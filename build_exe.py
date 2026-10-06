# -*- coding: utf-8 -*-
"""PyInstaller 打包脚本"""
import subprocess
import sys
import os

BASE = os.path.dirname(os.path.abspath(__file__))
LOG = os.path.join(BASE, "build_log.txt")

cmd = [
    sys.executable, "-m", "PyInstaller",
    "--noconfirm", "--clean", "--onefile", "--noconsole",
    "--icon", os.path.join(BASE, "app.ico"),
    "--name", "文档转换助手",
    "--collect-all", "pymupdf",
    "--collect-all", "pdf2docx",
    # 排除不需要的 Qt 大模块, 减小体积
    "--exclude-module", "PySide6.QtWebEngineCore",
    "--exclude-module", "PySide6.QtWebEngineWidgets",
    "--exclude-module", "PySide6.QtQml",
    "--exclude-module", "PySide6.QtQuick",
    "--exclude-module", "PySide6.QtQuick3D",
    "--exclude-module", "PySide6.Qt3DCore",
    "--exclude-module", "PySide6.QtCharts",
    "--exclude-module", "PySide6.QtDataVisualization",
    "--exclude-module", "PySide6.QtPdf",   # 我们用 pymupdf
    "--exclude-module", "PySide6.QtMultimedia",
    "--exclude-module", "PySide6.QtSql",
    "--exclude-module", "PySide6.QtTest",
    "--exclude-module", "tkinter",
    "--exclude-module", "_tkinter",
    os.path.join(BASE, "run.py"),
]

print("Building...")
with open(LOG, "w", encoding="utf-8", errors="replace") as f:
    p = subprocess.run(cmd, stdout=f, stderr=subprocess.STDOUT, cwd=BASE, timeout=1800)
print("exit =", p.returncode)
exe = os.path.join(BASE, "dist", "文档转换助手.exe")
print("exe exists:", os.path.exists(exe),
      "size:", os.path.getsize(exe) // 1024 // 1024, "MB" if os.path.exists(exe) else "")
sys.exit(p.returncode)
