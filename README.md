# 本地静默打印服务 (Local Silent Printing Service)

## 开源声明 (Open Source Declaration)
本项目采用GNU Lesser General Public License (LGPL)协议，使用本项目需要：
1. 保留所有原始组件的版权声明和许可文件
2. 在软件显著位置标明使用了以下开源组件：
   - PyStray (LGPL)
   - Flask (BSD)
   - Pillow (PIL)
   - PDFtoPrinter (MIT)
3. 修改后的版本必须开源并采用相同协议

This project is licensed under LGPL. Users must:
1. Preserve all original copyright notices and license files
2. Clearly attribute the use of:
   - PyStray (LGPL)
   - Flask (BSD) 
   - Pillow (PIL)
   - PDFtoPrinter (MIT)
3. Modified versions must be open sourced under the same license

## 依赖组件 (Dependencies)
本项目使用了以下开源组件：  
This project uses the following open source components:

- Python - PSF开源协议 (Python Software Foundation License)
- Flask - BSD开源协议 (BSD License) 
- PyStray - LGPL开源协议 (GNU Lesser General Public License)
- Pillow - PIL开源协议 (PIL License)
- PyWin32 - PSF开源协议 (Python Software Foundation License)
- PDFtoPrinter - MIT开源协议 (MIT License)

## 功能描述 (Function Description)
这是一个基于Python的本地打印服务程序，提供以下功能：  
This is a Python-based local printing service program with the following features:

- HTTP/WebSocket API接口接收打印任务  
- 自动处理PDF打印任务  
- 支持指定打印机或使用默认打印机  
- 强制100x150mm纸张尺寸  
- 系统托盘图标显示服务状态  
- 开机自启动配置  

## 接口说明 (API Documentation)

### WebSocket API
地址: `ws://localhost:12346`  
Address: `ws://localhost:12346`

#### 支持指令 (Supported Commands)

1. 获取打印机列表 (Get Printer List)
```
{"method": "getprinterlist"}
```
响应:

```
{
    "method": "getprinterlist",
    "status": "ok",
    "printers": ["打印机1", "打印机2"]
}
```

2. 打印任务 (Print Task)
```
{
    "pdfUrl": "PDF文件URL或本地路径",
    "printerName": "可选打印机名称"
}
```
响应同HTTP API

#### 打印接口 (Print API)
- 路径: `/print`
- 方法: POST
- 请求格式: JSON

## 使用说明 (Usage Guide)

1. 将程序与PDFtoPrinter.exe放在同一目录
2. 运行app.py或打包后的exe
3. 通过HTTP或WebSocket接口发送打印请求
4. 系统托盘图标可控制打印状态

1. Place the program in the same directory as PDFtoPrinter.exe  
2. Run app.py or the packaged exe  
3. Send print requests via HTTP or WebSocket API  
4. System tray icon can control printing status  

## 注意事项 (Important Notes)

- 需要安装Python依赖：flask, pystray, pillow, pywin32等
- 确保防火墙允许12345/12346端口
- 打印过程中托盘图标会闪烁
- 日志文件存储在print.log中

- Python dependencies required: flask, pystray, pillow, pywin32 etc.  
- Ensure firewall allows ports 12345/12346  
- Tray icon blinks during printing  
- Logs are stored in print.log  

## 系统要求 (System Requirements)
Python 3.8 或更高版本 (Python 3.8 or higher)
需要安装以下依赖包 (The following dependencies are required):
安装指南 (Installation Guide)
克隆或下载项目 (Clone or download the project)
安装依赖 (Install dependencies):
bash
pip install flask pystray pillow pywin32
或者使用requirements.txt (Or use requirements.txt):

bash
pip install -r requirements.txt
确保PDFtoPrinter.exe在项目目录中 (Make sure PDFtoPrinter.exe is in the project directory)