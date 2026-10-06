# -*- coding: utf-8 -*-
"""exe 冒烟测试 v2: PrintWindow 抓目标窗口, 不受遮挡影响"""
import os
import sys
import time
import ctypes
import subprocess

BASE = os.path.dirname(os.path.abspath(__file__))
EXE = os.path.join(BASE, "dist", "文档转换助手.exe")
SHOT = os.path.join(BASE, "test_data", "shots", "exe_window.png")

import win32gui
import win32ui
import win32con
from PIL import Image


def find_window_by_title(substr, timeout=25):
    end = time.time() + timeout
    while time.time() < end:
        hits = []
        def cb(hwnd, _):
            t = win32gui.GetWindowText(hwnd)
            if substr in t and win32gui.IsWindowVisible(hwnd):
                hits.append((hwnd, t))
        win32gui.EnumWindows(cb, None)
        if hits:
            return hits[0]
        time.sleep(0.5)
    return None


def capture_window(hwnd, path):
    l, t, r, b = win32gui.GetWindowRect(hwnd)
    w, h = r - l, b - t
    hwndDC = win32gui.GetWindowDC(hwnd)
    mfcDC = win32ui.CreateDCFromHandle(hwndDC)
    saveDC = mfcDC.CreateCompatibleDC()
    bmp = win32ui.CreateBitmap()
    bmp.CreateCompatibleBitmap(mfcDC, w, h)
    saveDC.SelectObject(bmp)
    ok = ctypes.windll.user32.PrintWindow(hwnd, saveDC.GetSafeHdc(), 2)  # PW_RENDERFULLCONTENT
    info = bmp.GetInfo()
    data = bmp.GetBitmapBits(True)
    img = Image.frombuffer("RGB", (info["bmWidth"], info["bmHeight"]), data, "raw", "BGRX", 0, 1)
    img.save(path)
    win32gui.DeleteObject(bmp.GetHandle())
    saveDC.DeleteDC()
    mfcDC.DeleteDC()
    win32gui.ReleaseDC(hwnd, hwndDC)
    return ok, (w, h)


proc = subprocess.Popen([EXE])
print("started pid =", proc.pid)

hit = find_window_by_title("文档转换助手")
if not hit:
    print("SMOKE FAIL: window not found in 25s")
    subprocess.run(["taskkill", "/PID", str(proc.pid), "/F"], capture_output=True)
    sys.exit(1)

hwnd, title = hit
time.sleep(2)  # 等界面完全渲染
ok, size = capture_window(hwnd, SHOT)
print(f"window: {title!r} size={size} printwindow_ok={ok}")
print("shot:", SHOT)

# 再等几秒确认进程稳定存活 (无闪退)
time.sleep(5)
r = subprocess.run(["tasklist", "/FI", f"PID eq {proc.pid}"],
                   capture_output=True, text=True, encoding="gbk", errors="replace")
alive = str(proc.pid) in (r.stdout or "")
print("still alive after 7s:", alive)
subprocess.run(["taskkill", "/PID", str(proc.pid), "/F"], capture_output=True)
print("SMOKE", "PASS" if alive else "FAIL")
sys.exit(0 if alive else 1)
