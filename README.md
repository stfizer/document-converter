# 📄 文档转换助手 (Document Converter)

本地离线运行的文档格式转换与编辑工具，一键打包为便携 exe，无需安装、无浏览器依赖。

> bilibili@何止有点帅 · 如果软件有帮到你，欢迎到 [B 站主页](https://space.bilibili.com/702444737) 支持一下！

## ✨ 功能

### 🔄 格式转换（支持批量、拖拽添加）
| 任务 | 说明 |
|------|------|
| Word / Excel / PPT → PDF | COM 调用本机 MS Office 导出，保真度最高；无 Office 时自动切换内置引擎简易排版兜底 |
| PDF → Word | pdf2docx 布局还原（保留表格与排版） |
| PDF → 图片 | 逐页导出 PNG |
| 图片 → PDF | 多图合成为一个 PDF |
| Word ↔ TXT / Excel ↔ CSV | 纯文本互转 |

### 🧰 PDF 工具
合并、按页码范围拆分、AES-256 加密、解密

### 📝 文字与图片编辑
- **Word**：段落列表选段改文字（保留原格式）、增删段落、插入图片
- **Excel**：编辑单元格、按锚点单元格插入图片
- **PDF**：页面可视化点击添加文字（可选字号/颜色）、点击插入图片、全文查找替换/删除文字、删除页

## 🚀 运行

### 方式一：便携 exe（推荐）
双击 `dist/文档转换助手.exe` 即可（约 130MB，离线运行，Windows 10/11）。

### 方式二：源码运行
```bash
pip install -r requirements.txt
python run.py
```

### 打包 exe
```bash
python make_icon.py     # 生成图标
python build_exe.py     # PyInstaller onefile 打包
```

## 🧪 测试
```bash
python test_core.py            # 转换核心 13 项测试
python test_editors_logic.py   # 编辑器逻辑 7 项测试
python test_drop.py            # 拖拽功能测试
python test_screenshot.py      # GUI 截图
```

## 🛠 技术栈
PySide6 · PyMuPDF · pdf2docx · python-docx · openpyxl · pywin32(COM) · PyInstaller

## 📌 说明
- Word/Excel/PPT → PDF 的高保真模式依赖本机安装的 MS Office 或 WPS，其余功能完全离线
- 旧版 .xls/.doc 请先另存为 .xlsx/.docx
