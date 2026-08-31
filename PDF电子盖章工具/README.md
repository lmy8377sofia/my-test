# PDF 电子盖章工具

本地桌面软件：给 PDF 电子盖章（去白底红章 / 手动拖放 / 多页多章 / 按关键词自动盖章 / 导出另存）。

- 仅本地运行，不上传网络；界面中文
- Windows 10/11 可用（Linux/macOS 亦可运行）
- 技术栈：Python 3.10+ · Tkinter · PyMuPDF · Pillow · NumPy（可选加速）

## 快速开始（Windows）

1. 双击 `启动盖章工具.bat`（首次运行自动安装依赖）
2. 或打包好的 exe：运行 `build_exe.bat` 后使用 `dist\PDF电子盖章工具.exe`

```bat
:: 源码运行
pip install -r requirements.txt
python pdf_stamp.py
```

## 项目结构

```text
PDF电子盖章工具/
  pdf_stamp.py        主程序（GUI + 核心算法 + 自检）
  requirements.txt    PyMuPDF / Pillow / numpy
  启动盖章工具.bat     一键启动（自动装依赖，无黑框）
  build_exe.bat       PyInstaller 一键打包
  使用说明.txt         中文使用文档
  assets/             app.ico 图标；demo/ 示例印章与示例 PDF
```

## 核心功能

| 功能 | 说明 |
| --- | --- |
| 印章预处理 | 白底自动透明（阈值可调）+ 边缘羽化 + 整体透明度；NumPy 向量化，无 NumPy 逐像素兜底 |
| 尺寸换算 | 直径按 mm 控制，默认 55mm；嵌入位图约 320 DPI，`px = mm/25.4*320` |
| 坐标系统 | 统一 fitz 左上原点（Y 向下），`size_pt = mm/25.4*72`，中心点定位 |
| 手动盖章 | 单击放置、拖动移动、选中后滑条实时改大小/透明度/旋转、Delete 删除 |
| 自动盖章 | 规则 A：第 1 页找「华信检测 / 有限公司」行中心；规则 B：第 2 页找「项目成员」标题中心；定位跟文字走 |
| 导出 | `page.insert_image(rect, stream=PNG, overlay=True, keep_proportion=True)`，另存 `xxx_已盖章.pdf`，`deflate + garbage=3` |

## 无界面自检

```bash
python pdf_stamp.py --selftest    # 验证预处理/自动定位/导出/体积/原文件不变
python pdf_stamp.py --demo        # 生成示例印章与示例 PDF 到 assets/demo/
```

## 打包 exe

```bat
build_exe.bat
:: 等价命令：
python -m PyInstaller --noconfirm --clean --onefile --windowed --name "PDF电子盖章工具" --icon assets/app.ico pdf_stamp.py
```
