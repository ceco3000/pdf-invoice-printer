# 发票打印助手（PDF 合并打印）

Windows 桌面小工具：选择多个 PDF 发票 / 行程单，按「每页 4 张 / 2 张、横向 / 竖向」自动排版到 A4，并发送到系统默认打印机。

## 功能

- 拖拽或文件对话框添加 PDF，自动读取页数
- 三种版式：4 张（横）、2 张（横）、2 张（竖）
- 自动去除页面旋转，等比缩放居中，最高 300 DPI 渲染
- 后台线程排版/打印，界面实时显示进度
- 打印完成后自动清理临时文件

## 目录结构

```
invoice_printer.py        # 主程序（Tkinter 界面 + 排版 + 打印）
启动发票打印助手.bat       # Windows 启动器：查找 Python、安装依赖并启动
requirements.txt          # Python 依赖
app_icon.ico              # 程序图标
测试发票/                 # 合成测试发票（TEST ONLY，非真实数据）
测试行程单/               # 合成测试行程单（TEST ONLY，非真实数据）
```

## 环境要求

- Windows 10 / 11
- Python 3.10+（推荐 3.12）
- 依赖：PyMuPDF、Pillow、pywin32、tkinterdnd2

## 使用

双击 `启动发票打印助手.bat`。脚本会自动查找本机 Python、在需要时安装依赖并启动界面。

也可以手动运行：

```bat
pip install -r requirements.txt
python invoice_printer.py
```

## 打包产物

`*.exe`、`*.zip` 等 PyInstaller 打包结果不纳入版本库（见 `.gitignore`），需要时在本地构建。

## 说明

打印功能依赖 Windows 打印 API（`win32print` / `win32ui`），仅在 Windows 上可用；仓库中的测试 PDF 均为合成数据。
