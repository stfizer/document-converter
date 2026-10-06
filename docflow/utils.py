# -*- coding: utf-8 -*-
"""通用工具函数"""
import os
import sys
import subprocess
import datetime


def human_size(path):
    try:
        n = os.path.getsize(path)
    except OSError:
        return "?"
    for unit in ("B", "KB", "MB", "GB"):
        if n < 1024 or unit == "GB":
            return f"{n:.1f}{unit}" if unit != "B" else f"{n}B"
        n /= 1024
    return "?"


def timestamp():
    return datetime.datetime.now().strftime("%H:%M:%S")


def open_in_explorer(path):
    """在资源管理器中打开文件所在文件夹并选中文件"""
    try:
        if os.path.isfile(path):
            subprocess.Popen(["explorer", "/select,", os.path.normpath(path)])
        else:
            subprocess.Popen(["explorer", os.path.normpath(path)])
    except Exception:
        pass


def resource_path(rel):
    """兼容 PyInstaller onefile 的资源路径"""
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
    return os.path.join(base, rel)


WORD_EXT = (".docx", ".doc")
EXCEL_EXT = (".xlsx", ".xls", ".xlsm")
PPT_EXT = (".pptx", ".ppt")
PDF_EXT = (".pdf",)
IMAGE_EXT = (".png", ".jpg", ".jpeg", ".bmp", ".gif", ".webp", ".tif", ".tiff")
