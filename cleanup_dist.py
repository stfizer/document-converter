# -*- coding: utf-8 -*-
"""清理构建产物: build/ 永久删除(可再生); 用户测试文档移入回收站(可恢复)"""
import ctypes
import os
import shutil
import sys
from ctypes import wintypes

BASE = os.path.dirname(os.path.abspath(__file__))
os.chdir(BASE)

# ---------- 1. build/ 中间产物: 直接删除 ----------
shutil.rmtree(os.path.join(BASE, "build"), ignore_errors=True)
print("build/ removed:", not os.path.exists(os.path.join(BASE, "build")))

# ---------- 2. 用户文档 -> 回收站 ----------
class SHFILEOPSTRUCTW(ctypes.Structure):
    _fields_ = [("hwnd", wintypes.HWND), ("wFunc", ctypes.c_uint),
                ("pFrom", wintypes.LPCWSTR), ("pTo", wintypes.LPCWSTR),
                ("fFlags", ctypes.c_ushort), ("fAnyOperationsAborted", wintypes.BOOL),
                ("hNameMappings", ctypes.c_void_p), ("lpszProgressTitle", wintypes.LPCWSTR)]


FO_DELETE, FOF_ALLOWUNDO, FOF_NOCONFIRMATION, FOF_SILENT = 3, 0x40, 0x10, 0x4


def to_recycle_bin(paths):
    op = SHFILEOPSTRUCTW()
    op.hwnd = None
    op.wFunc = FO_DELETE
    op.pFrom = "\0".join(os.path.abspath(p) for p in paths) + "\0"
    op.fFlags = FOF_ALLOWUNDO | FOF_NOCONFIRMATION | FOF_SILENT
    rc = ctypes.windll.shell32.SHFileOperationW(ctypes.byref(op))
    return rc == 0 and not op.fAnyOperationsAborted


files = [
    "dist/4060电压频率表.pdf",
    "dist/4060电压频率表.xlsx",
    "dist/9800X3D 异步外频.docx",
    "dist/9800X3D 异步外频.txt",
    "dist/万古神帝.docx",
    "dist/万古神帝.pdf",
    "build_log.txt",
]
files = [f for f in files if os.path.exists(f)]
ok = to_recycle_bin(files)
print("recycled:", ok, f"({len(files)} files)")

# ---------- 3. 验证 ----------
print("\n=== dist 剩余内容 ===")
for n in sorted(os.listdir("dist")):
    size = os.path.getsize(os.path.join("dist", n))
    print(f"  {n}  ({size // 1024 // 1024}MB)" if size > 1048576 else f"  {n}  ({size}B)")

print("\n=== 根目录检查 ===")
print("build exists:", os.path.exists("build"))
print("build_log exists:", os.path.exists("build_log.txt"))
exe = os.path.join("dist", "文档转换助手.exe")
print("exe kept:", os.path.exists(exe), f"({os.path.getsize(exe) // 1024 // 1024}MB)")
