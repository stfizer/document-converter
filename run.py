# -*- coding: utf-8 -*-
"""文档转换助手 - 程序入口"""
import ctypes
import sys
import os


def main():
    if sys.platform == "win32":
        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            pass
    base = os.path.dirname(os.path.abspath(__file__))
    from docflow.gui import launch
    launch(os.path.join(base, "app.ico"))


if __name__ == "__main__":
    main()
