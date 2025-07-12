 # 兼容打包路径和中文目录，防止 pystray、Pillow、tkinter、win32print、websockets、requests、flask 依赖遗漏
import sys
import os
import tempfile
import requests
from flask import Flask, request, jsonify
import subprocess
import shutil
import threading
import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime

# 直接导入 win32print 模块 (Directly import win32print module)
import win32print

# 直接导入 Windows 注册表模块 (Directly import Windows Registry module)
import winreg

# 重新添加 PIL (Pillow) 模块和 asyncio 模块的导入 (Re-add imports for PIL (Pillow) and asyncio modules)
import asyncio
import websockets  # 显式导入 websockets，确保 PyInstaller 能正确识别依赖 (Explicitly import websockets to ensure PyInstaller recognizes dependencies)
import pystray
from PIL import Image, ImageDraw, ImageTk

# 导入 uuid 模块 (Import uuid module)
import uuid


"""
全局变量声明，确保跨线程/函数访问
"""
root = None
status_icon_label = None
status_text_label = None
port_value_label = None
ws_port_value_label = None
logbox = None
tray_icon_instance = None
tk_icons = {}

is_printing = False  # 打印中状态

# 获取程序运行路径 (Get program execution path)
if getattr(sys, 'frozen', False):
    # 如果是打包后的exe文件 (If it's a bundled .exe file)
    BASE_DIR = os.path.dirname(sys.executable)
    APP_EXEC_PATH = sys.executable
else:
    # 如果是直接运行的.py文件 (If it's a directly run .py file)
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    APP_EXEC_PATH = f'"{sys.executable}" "{os.path.abspath(sys.argv[0])}"' # Python interpreter + script path (Python解释器 + 脚本路径)

PDFTOPRINTER_PATH = os.path.join(BASE_DIR, 'PDFtoPrinter.exe')

# 路径全部用 BASE_DIR 拼接，确保中文目录、任意目录都能用 (All paths are joined with BASE_DIR to ensure compatibility with Chinese directories and arbitrary directories)
LOG_FILE = os.path.join(BASE_DIR, 'print.log')
CACHE_DIR = os.path.join(BASE_DIR, 'pdf_cache')
MAX_CACHE = 5

app = Flask(__name__)
PRINT_ALLOWED = True
PRINT_PAUSED = False
PORT = 12345  # Flask HTTP 端口 (Flask HTTP Port)
WS_PORT = 12346  # WebSocket 独立端口，需与前端 ws://localhost:12346 保持一致 (WebSocket independent port, needs to match frontend ws://localhost:12346)
current_paper_size = '' # This variable is not used to control paper size, it's fixed below (此变量不用于控制纸张尺寸，它在下面是固定的)

# --- 开机自启动功能相关代码 (Auto-start functionality related code) ---
AUTOSTART_REG_KEY = r"Software\Microsoft\Windows\CurrentVersion\Run"
APP_NAME = "HttpPrinterService" # 应用程序在注册表中的名称 (Application name in registry)

def is_autostart_enabled():
    """检查程序是否已设置为开机自启动 (Checks if the program is set to auto-start on boot)"""
    # winreg 模块现在是直接导入的，如果导入失败会直接报错退出 (winreg module is now directly imported, if import fails, it will exit with an error)
    try:
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, AUTOSTART_REG_KEY, 0, winreg.KEY_READ)
        value, _ = winreg.QueryValueEx(key, APP_NAME)
        winreg.CloseKey(key)
        # 检查注册表中的值是否与当前程序路径匹配 (Check if the value in the registry matches the current program path)
        return value == APP_EXEC_PATH
    except FileNotFoundError:
        return False # 注册表项不存在 (Registry key not found)
    except Exception as e:
        log(f"检查开机自启动状态失败: {e}")
        return False

def enable_autostart():
    """设置程序开机自启动 (Sets the program to auto-start on boot)"""
    # winreg 模块现在是直接导入的 (winreg module is now directly imported)
    try:
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, AUTOSTART_REG_KEY, 0, winreg.KEY_SET_VALUE)
        winreg.SetValueEx(key, APP_NAME, 0, winreg.REG_SZ, APP_EXEC_PATH)
        winreg.CloseKey(key)
        log(f"已设置开机自启动: {APP_EXEC_PATH}")
        return True
    except Exception as e:
        log(f"设置开机自启动失败: {e}")
        return False

def disable_autostart():
    """取消程序开机自启动 (Disables program auto-start on boot)"""
    # winreg 模块现在是直接导入的 (winreg module is now directly imported)
    try:
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, AUTOSTART_REG_KEY, 0, winreg.KEY_SET_VALUE)
        winreg.DeleteValue(key, APP_NAME)
        winreg.CloseKey(key)
        log("已取消开机自启动。")
        return True
    except FileNotFoundError:
        log("开机自启动项不存在，无需取消。")
        return True # 注册表项不存在，视为成功取消 (Registry entry not found, considered successfully disabled)
    except Exception as e:
        log(f"取消开机自启动失败: {e}")
        return False

# --- 现有代码保持不变 (Existing code remains unchanged) ---

def log(msg):
    try:
        # 先读取现有日志，保留最后99条 (Read existing logs first, keep the last 99 lines)
        lines = []
        if os.path.exists(LOG_FILE):
            with open(LOG_FILE, 'r', encoding='utf-8') as f:
                lines = f.readlines()[-99:]
        # 新日志追加 (Append new log)
        lines.append(f"{datetime.now()} {msg}\n")
        with open(LOG_FILE, 'w', encoding='utf-8') as f:
            f.writelines(lines)
    except Exception as e:
        print(f"[日志写入异常] {e}。缓存/日志目录无写入权限，请检查文件夹权限或以管理员身份运行。", file=sys.stderr)
    # 实时刷新日志到GUI (Refresh log to GUI in real-time)
    global root, logbox
    if logbox and root:
        root.after(0, lambda: _update_logbox(msg))

def _update_logbox(msg):
    global logbox
    if logbox:
        logbox.insert(tk.END, f"{datetime.now()} {msg}\n")
        logbox.see(tk.END)

@app.route('/print', methods=['POST'])
def print_pdf():
    global PRINT_ALLOWED, PRINT_PAUSED, is_printing
    try:
        if not PRINT_ALLOWED or PRINT_PAUSED:
            log('打印被暂停或禁止')
            return jsonify({'status': 'error', 'message': '打印被暂停或禁止'}), 403
        
        data = request.json
        pdf_url = data.get('pdfUrl')
        printer_name = data.get('printerName')
        
        # 参数类型校验 (Parameter type validation)
        if not isinstance(pdf_url, str) or (printer_name is not None and not isinstance(printer_name, str)):
            log('参数类型错误，请检查接口调用方式。')
            return jsonify({'status': 'error', 'message': '参数类型错误，请检查接口调用方式。'}), 400
        
        # 强制所有打印任务都用 100*150mm 纸张 (Force all print tasks to use 100x150mm paper)
        paper_size = '100x150'
        paper_size_arg = f'/s"{paper_size}"' # Use /s flag and quote the size value (使用 /s 标志并引用尺寸值)

        if not pdf_url:
            log('打印失败：未提供pdfUrl。建议：检查接口调用参数。')
            return jsonify({'status': 'error', 'message': 'pdfUrl required。建议：检查接口参数。'}), 400
        
        # 每次都下载，不允许缓存打印 (Download every time, no caching allowed for printing)
        try:
            is_printing = True
            if hasattr(start_gui, 'set_tray_blink'):
                start_gui.set_tray_blink(True)

            if pdf_url.startswith('http://') or pdf_url.startswith('https://'):
                temp_pdf = download_pdf(pdf_url) # 修复：将 'url' 改为 'pdf_url' (Fix: changed 'url' to 'pdf_url')
            else:
                temp_pdf = os.path.join(CACHE_DIR, f"{uuid.uuid4()}.pdf")
                try:
                    shutil.copy(pdf_url, temp_pdf)
                except Exception as e:
                    log(f'本地PDF拷贝失败：{pdf_url}，错误：{str(e)}。请检查文件路径和权限。')
                    raise Exception(f'本地PDF拷贝失败：{str(e)}。建议：检查文件路径、权限。')
                clean_cache()
            
            if not os.path.exists(PDFTOPRINTER_PATH):
                download_url = 'https://mendelson.org/pdftoprinter.html'  # 示例下载地址，请替换为实际可用链接 (Example download URL, please replace with actual available link)
                log(f'缺少 PDFtoPrinter.exe，无法打印。请确认该文件与本程序在同一目录。下载地址：{download_url}')
                raise Exception(f'缺少 PDFtoPrinter.exe，无法打印。建议：将 PDFtoPrinter.exe 放到本程序同目录。下载地址：<a href="{download_url}" target="_blank">点击下载</a>')
            
            # Determine the actual printer name (确定实际的打印机名称)
            actual_printer_name = ""
            if printer_name is not None:
                pn_from_request = printer_name.strip()
                if pn_from_request: # If frontend provided a non-empty printer name (如果前端提供了非空的打印机名称)
                    # Validate if it's a real printer name, not a mis-sent paper size argument (验证它是否是真实的打印机名称，而不是错误发送的纸张尺寸参数)
                    if pn_from_request.lower().startswith('/papersize=') or pn_from_request.lower().startswith('/s='):
                        log(f'警告: 前端发送了无效的打印机名，包含纸张尺寸参数: {repr(pn_from_request)}。将忽略此值并尝试使用默认打印机。')
                        # Fallback to default printer if invalid printerName is sent (如果发送了无效的printerName，则回退到默认打印机)
                        try:
                            actual_printer_name = win32print.GetDefaultPrinter()
                            log(f'已回退到系统默认打印机: {repr(actual_printer_name)}')
                        except Exception as e:
                            log(f'获取系统默认打印机失败：{e}，自动补 ""')
                            actual_printer_name = ""
                    else:
                        actual_printer_name = pn_from_request
                        try:
                            if actual_printer_name not in get_printers():
                                log(f'打印机名称无效：{actual_printer_name}。请检查打印机是否连接、名称是否正确。')
                                raise Exception(f'打印机名称无效：{actual_printer_name}。建议：检查打印机连接和名称。')
                        except Exception as e:
                            log(str(e))
                            raise # Re-raise printer validation error (重新抛出打印机验证错误)
                else: # printerName was provided but empty (printerName已提供但为空)
                    try:
                        actual_printer_name = win32print.GetDefaultPrinter()
                        log(f'前端未传有效 printerName 字段，自动获取系统默认打印机: {repr(actual_printer_name)}')
                    except Exception as e:
                        log(f'获取系统默认打印机失败：{e}，自动补 ""')
                        actual_printer_name = ""
            else: # printerName was not provided at all (printerName根本没有提供)
                try:
                    actual_printer_name = win32print.GetDefaultPrinter()
                    log(f'前端未传 printerName 字段，自动获取系统默认打印机: {repr(actual_printer_name)}')
                except Exception as e:
                    log(f'获取系统默认打印机失败：{e}，自动补 ""')
                    actual_printer_name = ""

            # Construct the command list (构造命令列表)
            cmd = [PDFTOPRINTER_PATH, temp_pdf]
            
            # Add the actual printer name (it will be quoted by subprocess.run if it contains spaces and shell=False)
            # If actual_printer_name is empty, this will append an empty string, which PDFtoPrinter.exe might handle as default.
            cmd.append(actual_printer_name) 

            # Add the paper size argument as the last argument (将纸张尺寸参数作为最后一个参数添加)
            cmd.append(paper_size_arg)

            log(f'最终执行命令: {cmd}')
            
            try:
                result = subprocess.run(cmd, capture_output=True, text=True, timeout=60, check=False) # check=False to handle non-zero exit codes manually (check=False 以手动处理非零退出代码)
            except subprocess.TimeoutExpired:
                is_printing = False
                if hasattr(start_gui, 'set_tray_blink'):
                    start_gui.set_tray_blink(False)
                log(f'打印超时：{pdf_url}，建议检查打印机连接和状态。')
                return jsonify({'status': 'error', 'message': '打印超时，建议检查打印机连接和状态。'}), 504
            except OSError as e:
                is_printing = False
                if hasattr(start_gui, 'set_tray_blink'):
                    start_gui.set_tray_blink(False)
                log(f'PDFtoPrinter.exe 无法执行，可能被杀毒软件拦截，请恢复文件并添加信任。错误：{e}')
                return jsonify({'status': 'error', 'message': 'PDFtoPrinter.exe 无法执行，可能被杀毒软件拦截，请恢复文件并添加信任。'}), 500
            
            finally:
                is_printing = False
                if hasattr(start_gui, 'set_tray_blink'):
                    start_gui.set_tray_blink(False)

            if result.returncode == 0:
                log(f'打印成功：{pdf_url} -> {actual_printer_name or "默认打印机"} 纸张:{paper_size} 缓存:{temp_pdf}')
                return jsonify({'status': 'ok', 'message': result.stdout, 'cachePath': temp_pdf})
            else:
                # PDF损坏或格式不支持检测 (PDF corrupted or format not supported detection)
                err_msg = result.stderr.lower()
                if 'invalid' in err_msg or 'corrupt' in err_msg:
                    log(f'PDF文件损坏或格式不受支持：{pdf_url}。建议重新生成或检查源文件。')
                    return jsonify({'status': 'error', 'message': 'PDF文件损坏或格式不受支持，建议重新生成或检查源文件。', 'cachePath': temp_pdf})
                # 如果没有错误信息，可能是未设置默认打印机 (If no error message, it might be that the default printer is not set)
                if not result.stderr.strip():
                    msg = '打印失败：未检测到默认打印机，或打印机不可用。请在系统设置中设置默认打印机并确保其可用。'
                    log(f'{msg} {pdf_url} 缓存:{temp_pdf}')
                    return jsonify({'status': 'error', 'message': msg, 'cachePath': temp_pdf})
                log(f'打印失败：{pdf_url} -> {actual_printer_name or "默认打印机"}，错误：{result.stderr} 缓存:{temp_pdf}。建议：检查打印机状态、纸张、驱动。')
                return jsonify({'status': 'error', 'message': result.stderr + "。建议：检查打印机状态、纸张、驱动。", 'cachePath': temp_pdf})
        except Exception as e:
            is_printing = False
            if hasattr(start_gui, 'set_tray_blink'):
                start_gui.set_tray_blink(False)
            log(f'打印异常：{pdf_url}，错误：{str(e)}。如多次出现此类错误，请联系技术支持。')
            return jsonify({'status': 'error', 'message': str(e) + "。如多次出现此类错误，请联系技术支持。"})
    except Exception as e:
        log(f'未知错误：{str(e)}')
        return jsonify({'status': 'error', 'message': f'未知错误：{str(e)}'})

def run_flask():
    global PORT
    try:
        app.run(host='127.0.0.1', port=PORT)
    except OSError as e:
        log(f'HTTP服务启动失败：{e}。端口被占用，请更换端口或关闭占用程序。重启软件后请按F5刷新打印网站。')
        print(f'[ERROR] HTTP服务启动失败：{e}。端口被占用，请更换端口或关闭占用程序。重启软件后请按F5刷新打印网站。', file=sys.stderr)

 # 原生 WebSocket 服务端实现 (Native WebSocket server implementation)
import websockets.exceptions
async def ws_handler(websocket):
    global PRINT_ALLOWED, PRINT_PAUSED
    try:
        async for message in websocket:
            print(f"[DEBUG] 收到原始消息: {message}")
            import json
            import urllib.parse
            # 先尝试URL解码 (First try URL decoding)
            try:
                decoded = urllib.parse.unquote(message)
                print(f"[DEBUG] URL解码后: {decoded}")
                data = json.loads(decoded)
            except Exception:
                # 兼容前端直接发送字符串指令的情况 (Compatible with frontend sending string commands directly)
                msg = message.strip().lower()
                if msg == 'getprinterlist':
                    data = {'method': 'getprinterlist'}
                elif msg == 'get_printers':
                    data = {'method': 'get_printers'}
                else:
                    await websocket.send(json.dumps({'status': 'error', 'message': '数据格式错误'}))
                    continue

            # 兼容网站前端JS：获取打印机列表 (Compatible with website frontend JS: Get printer list)
            if data.get('method') == 'getprinterlist':
                printers = get_printers()
                # 构造兼容前端的返回格式，增加 printers 字段，最大兼容 (Construct a return format compatible with the frontend, add printers field, maximum compatibility)
                resp = {
                    'method': 'getprinterlist',
                    'status': 'ok',
                    'data': [{'name': p} for p in printers],
                    'printers': printers
                }
                resp_json = json.dumps(resp, ensure_ascii=False)
                print(f"[DEBUG] get_printers() 返回: {printers}")
                print(f"[DEBUG] WebSocket 发送: {resp_json}")
                await websocket.send(resp_json)
                continue
            # 兼容旧接口 (Compatible with old interface)
            if data.get('method') == 'get_printers':
                printers = get_printers()
                await websocket.send(json.dumps({'status': 'ok', 'printers': printers}))
                continue

            if not PRINT_ALLOWED or PRINT_PAUSED:
                log('打印被暂停或禁止')
                await websocket.send(json.dumps({'status': 'error', 'message': '打印被暂停或禁止'}))
                continue
            pdf_url = data.get('pdfUrl') or data.get('PdfUrl')
            printer_name = data.get('printerName')
            
            # 强制所有打印任务都用 100*150mm 纸张 (Force all print tasks to use 100x150mm paper)
            paper_size = '100x150'
            paper_size_arg = f'/s"{paper_size}"' # Use /s flag and quote the size value (使用 /s 标志并引用尺寸值)

            if not pdf_url:
                log('打印失败：未提供pdfUrl')
                await websocket.send(json.dumps({'status': 'error', 'message': 'pdfUrl required'}))
                continue
            try:
                global is_printing
                is_printing = True
                if hasattr(start_gui, 'set_tray_blink'):
                    start_gui.set_tray_blink(True)
                
                # 参数类型校验 (Parameter type validation)
                if not isinstance(pdf_url, str) or (printer_name is not None and not isinstance(printer_name, str)):
                    log('参数类型错误，请检查接口调用方式。(WS)')
                    await websocket.send(json.dumps({'status': 'error', 'message': '参数类型错误，请检查接口调用方式。'}))
                    is_printing = False
                    if hasattr(start_gui, 'set_tray_blink'):
                        start_gui.set_tray_blink(False)
                    continue

                if pdf_url.startswith('http://') or pdf_url.startswith('https://'):
                    try:
                        temp_pdf = download_pdf(pdf_url) # 修复：将 'url' 改为 'pdf_url' (Fix: changed 'url' to 'pdf_url')
                    except Exception as e:
                        await websocket.send(json.dumps({'status': 'error', 'message': str(e) + '（WebSocket端）'}))
                        is_printing = False
                        if hasattr(start_gui, 'set_tray_blink'):
                            start_gui.set_tray_blink(False)
                        continue
                else:
                    temp_pdf = os.path.join(CACHE_DIR, f"{uuid.uuid4()}.pdf")
                    try:
                        shutil.copy(pdf_url, temp_pdf)
                    except Exception as e:
                        log(f'本地PDF拷贝失败(WS)：{pdf_url}，错误：{str(e)}。请检查文件路径和权限。')
                        await websocket.send(json.dumps({'status': 'error', 'message': f'本地PDF拷贝失败：{str(e)}。建议：检查文件路径、权限。'}))
                        is_printing = False
                        if hasattr(start_gui, 'set_tray_blink'):
                            start_gui.set_tray_blink(False)
                        continue
                    clean_cache()
                
                if not os.path.exists(PDFTOPRINTER_PATH):
                    download_url = 'https://www.gridsoft.cn/download/PDFtoPrinter.exe'  # 示例下载地址，请替换为实际可用链接 (Example download URL, please replace with actual available link)
                    log(f'缺少 PDFtoPrinter.exe，无法打印(WS)。请确认该文件与本程序在同一目录。下载地址：{download_url}')
                    await websocket.send(json.dumps({
                        'status': 'error',
                        'message': f'缺少 PDFtoPrinter.exe，无法打印。建议：将 PDFtoPrinter.exe 放到本程序同目录。下载地址：<a href="{download_url}" target="_blank">点击下载</a>'
                    }))
                    is_printing = False
                    if hasattr(start_gui, 'set_tray_blink'):
                        start_gui.set_tray_blink(False)
                    continue
                
                # Determine the actual printer name (确定实际的打印机名称)
                actual_printer_name = ""
                if printer_name is not None:
                    pn_from_request = printer_name.strip()
                    if pn_from_request: # If frontend provided a non-empty printer name (如果前端提供了非空的打印机名称)
                        # Validate if it's a real printer name, not a mis-sent paper size argument (验证它是否是真实的打印机名称，而不是错误发送的纸张尺寸参数)
                        if pn_from_request.lower().startswith('/papersize=') or pn_from_request.lower().startswith('/s='):
                            log(f'警告(WS): 前端发送了无效的打印机名，包含纸张尺寸参数: {repr(pn_from_request)}。将忽略此值并尝试使用默认打印机。')
                            try:
                                actual_printer_name = win32print.GetDefaultPrinter()
                                log(f'已回退到系统默认打印机(WS): {repr(actual_printer_name)}')
                            except Exception as e:
                                log(f'获取系统默认打印机失败(WS)：{e}，自动补 ""')
                                actual_printer_name = ""
                        else:
                            actual_printer_name = pn_from_request
                            try:
                                if actual_printer_name not in get_printers():
                                    log(f'打印机名称无效(WS)：{actual_printer_name}。请检查打印机是否连接、名称是否正确。')
                                    await websocket.send(json.dumps({'status': 'error', 'message': f'打印机名称无效：{actual_printer_name}。建议：检查打印机连接和名称。'}))
                                    is_printing = False
                                    if hasattr(start_gui, 'set_tray_blink'):
                                        start_gui.set_tray_blink(False)
                                    continue
                            except Exception as e:
                                log(str(e))
                                await websocket.send(json.dumps({'status': 'error', 'message': str(e)}))
                                is_printing = False
                                if hasattr(start_gui, 'set_tray_blink'):
                                    is_printing = False
                                    if hasattr(start_gui, 'set_tray_blink'):
                                        start_gui.set_tray_blink(False)
                                    continue
                    else: # printerName was provided but empty (printerName已提供但为空)
                        try:
                            actual_printer_name = win32print.GetDefaultPrinter()
                            log(f'前端未传有效 printerName 字段，自动获取系统默认打印机(WS): {repr(actual_printer_name)}')
                        except Exception as e:
                            log(f'获取系统默认打印机失败(WS)：{e}，自动补 ""')
                            actual_printer_name = ""
                else: # printerName was not provided at all (printerName根本没有提供)
                    try:
                        actual_printer_name = win32print.GetDefaultPrinter()
                        log(f'前端未传 printerName 字段，自动获取系统默认打印机(WS): {repr(actual_printer_name)}')
                    except Exception as e:
                        log(f'获取系统默认打印机失败(WS)：{e}，自动补 ""')
                        actual_printer_name = ""

                # Construct the command list (构造命令列表)
                cmd = [PDFTOPRINTER_PATH, temp_pdf]
                cmd.append(actual_printer_name) 
                cmd.append(paper_size_arg) # Add the paper size argument as the last argument (将纸张尺寸参数作为最后一个参数添加)

                log(f'最终执行命令(WS): {cmd}')

                try:
                    result = subprocess.run(cmd, capture_output=True, text=True, timeout=60, check=False)
                except subprocess.TimeoutExpired:
                    is_printing = False
                    if hasattr(start_gui, 'set_tray_blink'):
                        start_gui.set_tray_blink(False)
                    log(f'打印超时(WS)：{pdf_url}，建议检查打印机连接和状态。')
                    await websocket.send(json.dumps({'status': 'error', 'message': '打印超时，建议检查打印机连接和状态。'}))
                    continue
                except OSError as e:
                    is_printing = False
                    if hasattr(start_gui, 'set_tray_blink'):
                        start_gui.set_tray_blink(False)
                    log(f'PDFtoPrinter.exe 无法执行(WS)，可能被杀毒软件拦截，请恢复文件并添加信任。错误：{e}')
                    await websocket.send(json.dumps({'status': 'error', 'message': 'PDFtoPrinter.exe 无法执行，可能被杀毒软件拦截，请恢复文件并添加信任。'}))
                    continue
                
                finally:
                    is_printing = False
                    if hasattr(start_gui, 'set_tray_blink'):
                        start_gui.set_tray_blink(False)

                if result.returncode == 0:
                    log(f'打印成功(WS)：{pdf_url} -> {printer_name or "默认打印机"} 纸张:{paper_size} 缓存:{temp_pdf}')
                    await websocket.send(json.dumps({'status': 'ok', 'message': result.stdout, 'cachePath': temp_pdf}))
                else:
                    err_msg = result.stderr.lower()
                    if 'invalid' in err_msg or 'corrupt' in err_msg:
                        log(f'PDF文件损坏或格式不受支持(WS)：{pdf_url}。建议重新生成或检查源文件。')
                        await websocket.send(json.dumps({'status': 'error', 'message': 'PDF文件损坏或格式不受支持，建议重新生成或检查源文件。', 'cachePath': temp_pdf}))
                        continue
                    # 如果没有错误信息，可能是未设置默认打印机 (If no error message, it might be that the default printer is not set)
                    if not result.stderr.strip():
                        msg = '打印失败：未检测到默认打印机，或打印机不可用。请在系统设置中设置默认打印机并确保其可用。'
                        log(f'{msg} {pdf_url} 缓存:{temp_pdf}')
                        await websocket.send(json.dumps({'status': 'error', 'message': msg, 'cachePath': temp_pdf}))
                        continue
                    log(f'打印失败(WS)：{pdf_url} -> {printer_name or "默认打印机"}，错误：{result.stderr} 缓存:{temp_pdf}。建议：检查打印机状态、纸张、驱动。')
                    await websocket.send(json.dumps({'status': 'error', 'message': result.stderr + "。建议：检查打印机状态、纸张、驱动。", 'cachePath': temp_pdf}))
            except websockets.exceptions.ConnectionClosed:
                log('WebSocket 客户端已断开')
                break
            except Exception as e:
                is_printing = False
                if hasattr(start_gui, 'set_tray_blink'):
                    start_gui.set_tray_blink(False)
                log(f'打印异常(WS)：{pdf_url}，错误：{str(e)}。如多次出现此类错误，请联系技术支持。')
                try:
                    await websocket.send(json.dumps({'status': 'error', 'message': str(e) + "。如多次出现此类错误，请联系技术支持。"}))
                except websockets.exceptions.ConnectionClosed:
                    log('WebSocket 客户端已断开')
                    break
    except Exception as e:
        log(f'WebSocket连接异常：{str(e)}')

def start_ws_server():
    print(f"[DEBUG] WebSocket服务即将启动，监听端口: {WS_PORT}")
    import sys
    try:
        async def ws_main():
            try:
                async with websockets.serve(ws_handler, '127.0.0.1', WS_PORT):
                    print(f"[DEBUG] WebSocket服务已启动，监听端口: {WS_PORT}")
                    await asyncio.Future()  # run forever
            except OSError as e:
                # 只写一条简明日志，不重复输出异常堆栈 (Only write a concise log, do not repeatedly output exception stack)
                log('WebSocket服务启动失败：端口被占用，请更换端口或关闭占用程序。重启软件后请按F5刷新打印网站。')
                print(f"[ERROR] WebSocket服务启动失败：端口被占用，请更换端口或关闭占用程序。重启软件后请按F5刷新打印网站。", file=sys.stderr)
        asyncio.run(ws_main())
    except Exception as e:
        print(f"[ERROR] WebSocket服务启动失败: {e}", file=sys.stderr)


def open_printer_settings():
    try:
        subprocess.Popen('control printers', shell=True)
    except Exception as e:
        messagebox.showerror("错误", f"无法打开打印机设置: {e}")

def start_gui():
    # 托盘闪烁控制 (Tray blinking control)
    tray_blinking = {'flag': False, 'thread': None}
    def tray_blink_worker():
        import time
        state = False
        while tray_blinking['flag']:
            if tray_icon_instance:
                tray_icon_instance.icon = tray_icons_pil['on'] if state else tray_icons_pil['off']
            state = not state
            time.sleep(0.5)
        # 恢复为当前状态 (Restore to current state)
        if tray_icon_instance:
            if PRINT_ALLOWED and not PRINT_PAUSED:
                tray_icon_instance.icon = tray_icons_pil['on']
            else:
                tray_icon_instance.icon = tray_icons_pil['off']

    def set_tray_blink(blink):
        if blink:
            if not tray_blinking['flag']:
                tray_blinking['flag'] = True
                tray_blinking['thread'] = threading.Thread(target=tray_blink_worker, daemon=True)
                tray_blinking['thread'].start()
        else:
            tray_blinking['flag'] = False
    start_gui.set_tray_blink = set_tray_blink
    global root, status_icon_label, status_text_label, port_value_label, ws_port_value_label, logbox, tray_icon_instance, tk_icons, PRINT_ALLOWED, PRINT_PAUSED, tray_icons_pil
    root = tk.Tk()
    # 初始窗口标题 (Initial window title)
    def update_title():
        if PRINT_ALLOWED and not PRINT_PAUSED:
            root.title("本地静默打印服务 - 打印已启动")
        else:
            root.title("本地静默打印服务 - 打印已暂停")
    update_title()
    root.geometry("600x500")
    root.protocol("WM_DELETE_WINDOW", lambda: root.withdraw())  # 关闭按钮只隐藏窗口 (Close button only hides window)

    # 生成绿色/灰色圆形图标（PIL Image） (Generate green/gray circular icon (PIL Image))
    def make_icon_pil(color, size=64):
        img = Image.new('RGBA', (size, size), (255, 255, 255, 0))
        draw = ImageDraw.Draw(img)
        draw.ellipse((8, 8, size-8, size-8), fill=color, outline=(128, 128, 128))
        return img

    tray_icons_pil = {
        'on': make_icon_pil((0, 255, 0)),
        'off': make_icon_pil((180, 180, 180))
    }

    # Tkinter用PhotoImage生成小圆点 (Tkinter uses PhotoImage to generate small dots)
    def make_icon_tk(color, size=18):
        img = Image.new('RGBA', (size, size), (255, 255, 255, 0))
        draw = ImageDraw.Draw(img)
        draw.ellipse((2, 2, size-2, size-2), fill=color, outline=(128, 128, 128))
        return ImageTk.PhotoImage(img, master=root)
    tk_icons['on'] = make_icon_tk((0, 255, 0))
    tk_icons['off'] = make_icon_tk((180, 180, 180))

    # 状态指示灯和文字（主窗口左上角） (Status indicator light and text (top-left of main window))
    status_icon_label = tk.Label(root, image=tk_icons['on'])
    status_icon_label.place(x=20, y=10)
    status_text_label = tk.Label(root, text="打印已启动", font=("微软雅黑", 11, "bold"), fg="#00b300")
    status_text_label.place(x=50, y=10)

    # 端口信息（紧跟状态指示灯下方） (Port information (directly below status indicator light))
    tk.Label(root, text="HTTP端口:").place(x=20, y=50)
    port_value_label = tk.Label(root, text=str(PORT), fg="#0055cc", font=("微软雅黑", 10, "bold"))
    port_value_label.place(x=100, y=50)
    tk.Label(root, text="WebSocket端口:").place(x=20, y=80)
    ws_port_value_label = tk.Label(root, text=str(WS_PORT), fg="#0055cc", font=("微软雅黑", 10, "bold"))
    ws_port_value_label.place(x=120, y=80)

    def set_port_label(port):
        port_value_label.config(text=str(port))
    def set_ws_port_label(port):
        ws_port_value_label.config(text=str(port))
    start_gui.set_port_label = set_port_label
    start_gui.set_ws_port_label = set_ws_port_label

    # 日志显示 (Log display)
    tk.Label(root, text="实时打印日志:").place(x=20, y=170)
    logbox = tk.Text(root, height=15, width=70)
    logbox.place(x=20, y=200)

    def clear_log():
        try:
            with open(LOG_FILE, 'w', encoding='utf-8') as f:
                f.write("")
            logbox.delete(1.0, tk.END)
        except Exception as e:
            messagebox.showerror("错误", f"无法清空日志: {e}")

    def start_print():
        global PRINT_ALLOWED, PRINT_PAUSED
        PRINT_ALLOWED = True
        PRINT_PAUSED = False
        logbox.insert(tk.END, "[系统] 打印已启动\n")
        try:
            update_gui_status()
        except Exception:
            pass
    def pause_print():
        global PRINT_PAUSED
        PRINT_PAUSED = True
        logbox.insert(tk.END, "[系统] 打印已暂停\n")
        try:
            update_gui_status()
        except Exception:
            pass
    def stop_print():
        global tray_icon_instance
        if tray_icon_instance:
            tray_icon_instance.stop()
        root.quit()
        os._exit(0)

    btn_width = 12
    btn_gap = 20
    base_x = 20
    y_pos = 130
    tk.Button(root, text="开始打印", width=btn_width, command=start_print).place(x=base_x, y=y_pos)
    tk.Button(root, text="暂停打印", width=btn_width, command=pause_print).place(x=base_x + (btn_width+2)*7 + btn_gap, y=y_pos)
    tk.Button(root, text="退出打印", width=btn_width, command=stop_print).place(x=base_x + 2*((btn_width+2)*7 + btn_gap), y=y_pos)
    tk.Button(root, text="清空日志", width=btn_width, command=clear_log).place(x=base_x + 3*((btn_width+2)*7 + btn_gap), y=y_pos)
    start_print()

    tk.Label(root, text="缓存PDF目录(仅供核查):").place(x=20, y=460)
    def open_cache_dir():
        abs_path = os.path.abspath(CACHE_DIR)
        if not os.path.exists(abs_path):
            os.makedirs(abs_path, exist_ok=True)
        try:
            os.startfile(abs_path)
        except Exception as e:
            messagebox.showerror("错误", f"无法打开缓存目录: {e}")
    tk.Button(root, text="打开缓存目录", command=open_cache_dir).place(x=160, y=455)

    def refresh_log():
        try:
            with open(LOG_FILE, 'r', encoding='utf-8') as f:
                lines = f.readlines()[-100:]
            logbox.delete(1.0, tk.END)
            for line in lines:
                logbox.insert(tk.END, line)
        except:
            pass
        root.after(2000, refresh_log)
    refresh_log()

    # --- GUI中的开机自启动开关 (Auto-start switch in GUI) ---
    autostart_var = tk.BooleanVar(value=is_autostart_enabled())
    def toggle_autostart():
        if autostart_var.get():
            if enable_autostart():
                log("开机自启动已启用。")
            else:
                autostart_var.set(False) # Revert if failed (失败则还原)
                messagebox.showerror("错误", "设置开机自启动失败，请检查权限。")
        else:
            if disable_autostart():
                log("开机自启动已禁用。")
            else:
                autostart_var.set(True) # Revert if failed (失败则还原)
                messagebox.showerror("错误", "取消开机自启动失败，请检查权限。")
        # 更新托盘菜单状态 (Update tray menu status)
        if tray_icon_instance:
            tray_icon_instance.update_menu()

    # 由于winreg现在是直接导入，这里无需再检查winreg是否为None (Since winreg is now directly imported, no need to check if winreg is None here)
    autostart_checkbox = ttk.Checkbutton(root, text="开机自动启动", variable=autostart_var, command=toggle_autostart)
    autostart_checkbox.place(x=20, y=110) # 放置在合适位置 (Place in a suitable position)


    def update_gui_status():
        global tray_icon_instance, status_icon_label, status_text_label, root, tk_icons, PRINT_ALLOWED, PRINT_PAUSED, tray_icons_pil
        tip = "（如遇异常重启软件后请按F5刷新打印网站）"
        if PRINT_ALLOWED and not PRINT_PAUSED:
            status_icon_label.config(image=tk_icons['on'])
            status_text_label.config(text="打印已启动 " + tip, fg="#00b300")
            root.title("本地静默打印服务 - 打印已启动 " + tip)
            if tray_icon_instance:
                tray_icon_instance.icon = tray_icons_pil['on']
                tray_icon_instance.title = "本地静默打印服务 - 打印已启动 " + tip
        else:
            status_icon_label.config(image=tk_icons['off'])
            status_text_label.config(text="打印已暂停 " + tip, fg="#888888")
            root.title("本地静默打印服务 - 打印已暂停 " + tip)
            if tray_icon_instance:
                tray_icon_instance.icon = tray_icons_pil['off']
                tray_icon_instance.title = "本地静默打印服务 - 打印已暂停 " + tip

    # 托盘相关 (Tray related)
    def on_show_window(icon, item):
        root.after(0, lambda: root.deiconify())
    def on_start_print(icon, item):
        root.after(0, lambda: start_print())
    def on_pause_print(icon, item):
        root.after(0, lambda: pause_print())
    def on_exit(icon, item):
        if icon:
            icon.stop()
        root.after(0, lambda: root.quit())
    def on_status(icon, item):
        if PRINT_ALLOWED and not PRINT_PAUSED:
            icon.notify("打印已启动", "绿色灯亮")
        else:
            icon.notify("打印已暂停", "灰色灯")
    
    # --- 托盘菜单中的开机自启动选项 (Auto-start option in tray menu) ---
    def on_toggle_autostart_tray(icon, item):
        # item.checked 会反映当前状态 (item.checked will reflect the current state)
        if item.checked: # 当前是启用状态，点击后应禁用 (Currently enabled, should disable after click)
            if disable_autostart():
                log("通过托盘菜单禁用开机自启动。")
            else:
                messagebox.showerror("错误", "取消开机自启动失败，请检查权限。")
        else: # 当前是禁用状态，点击后应启用 (Currently disabled, should enable after click)
            if enable_autostart():
                log("通过托盘菜单启用开机自启动。")
            else:
                messagebox.showerror("错误", "设置开机自启动失败，请检查权限。")
        # 更新GUI中的复选框状态 (Update checkbox status in GUI)
        root.after(0, lambda: autostart_var.set(is_autostart_enabled()))


    def tray_loop():
        global tray_icon_instance
        # 动态创建菜单项，以反映自启动状态 (Dynamically create menu items to reflect auto-start status)
        def create_menu():
            menu_items = [
                pystray.MenuItem('显示主窗口', on_show_window),
                pystray.MenuItem('开始打印', on_start_print),
                pystray.MenuItem('暂停打印', on_pause_print),
                pystray.MenuItem('显示当前状态', on_status),
            ]
            # 由于winreg现在是直接导入，这里无需再检查winreg是否为None (Since winreg is now directly imported, no need to check if winreg is None here)
            menu_items.append(
                pystray.MenuItem(
                    '开机自动启动', 
                    on_toggle_autostart_tray, 
                    checked=lambda item: is_autostart_enabled() # 动态检查状态 (Dynamically check status)
                )
            )
            menu_items.append(pystray.MenuItem('退出打印', on_exit))
            return pystray.Menu(*menu_items)

        tray_icon_instance = pystray.Icon("HttpPrinter", tray_icons_pil['on'], "本地静默打印服务", create_menu())
        tray_icon_instance.run()

    tray_thread = threading.Thread(target=tray_loop, daemon=True)
    tray_thread.start()

    # 启动时默认更新一次状态 (Update status once on startup)
    update_gui_status()

    root.mainloop()

if __name__ == '__main__':
    # 启动前清空日志文件，避免加载上一次的日志 (Clear log file before startup to avoid loading previous logs)
    try:
        with open(LOG_FILE, 'w', encoding='utf-8') as f:
            f.write("")
    except Exception:
        pass
    # 先启动GUI（主线程，保证Tkinter/托盘/进度条正常） (Start GUI first (main thread, ensure Tkinter/tray/progress bar work correctly))
    def start_servers():
        flask_thread = threading.Thread(target=run_flask, daemon=True)
        flask_thread.start()
        ws_thread = threading.Thread(target=start_ws_server, daemon=True)
        ws_thread.start()
    threading.Thread(target=start_servers, daemon=True).start()
    start_gui()
