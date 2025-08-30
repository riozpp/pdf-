## PDF 工具（拆分、合并、转 Word、转图片）

一个基于 Python 的桌面工具（Tkinter GUI），支持：
- 拆分 PDF（按页码范围）
- 合并多个 PDF
- PDF 转 Word（.docx）
- PDF 转图片（支持 PNG/JPG，可设定 DPI）

### 运行环境
- Windows 10/11
- Python 3.9+（建议 3.10/3.11）

### 安装依赖
```bash
python -m venv .venv
. .venv/Scripts/Activate.ps1
pip install -r requirements.txt
```

### 启动 GUI
```bash
python pdf_tool/app.py
```

### 打包为 exe（PyInstaller）
PowerShell 执行：
```powershell
. .venv/Scripts/Activate.ps1
powershell -ExecutionPolicy Bypass -File build_exe.ps1
```
成功后在 `dist/PDFTool/` 下获得可执行文件。

### 使用说明（GUI）
- 拆分：选择 PDF，填写页码范围（如 1-3,5,7-9），选择输出文件夹。
- 合并：按顺序添加多个 PDF，选择输出位置。
- 转 Word：选择 PDF，指定输出 .docx。
- 转图片：选择 PDF，设置 DPI、图片格式，指定输出文件夹。

### 注意
- PDF 转 Word 效果取决于文档结构（扫描件建议先 OCR）。
- PDF 转图片使用 PyMuPDF，性能和质量较好。
- 如遇杀软拦截，请将 `dist/PDFTool` 加白。



