# -*- coding: utf-8 -*-
"""调试: make 后立即转换 vs 延迟转换"""
import os
import sys
import time

BASE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE)
os.chdir(BASE)

import docflow.converters as cv
import pymupdf

src = os.path.join(BASE, "test_data", "orient", "wide_table.xlsx")


def try_convert(tag):
    dst = os.path.join(BASE, "test_data", "orient", f"dbg_{tag}.pdf")
    if os.path.exists(dst):
        os.remove(dst)
    ok, msg = cv.excel_to_pdf(src, dst, log=lambda m: print("  log:", m))
    size = os.path.getsize(dst) if os.path.exists(dst) else 0
    pages = -1
    info = ""
    if size:
        d = pymupdf.open(dst)
        pages = len(d)
        if pages:
            p = d[0]
            info = "%.0fx%.0f" % (p.rect.width, p.rect.height)
    print(f"[{tag}] ok={ok} size={size} pages={pages} {info}")
    print(f"   msg: {msg[:120]}")
    return pages > 0


print("== 1) 刚 make 完立即转换 (复现问题?) ==")
sys.argv = ["x"]
code = open("test_excel_orient.py", encoding="utf-8").read().split("def main")[0]
code = code.replace("BASE = os.path.dirname(os.path.abspath(__file__))",
                    "BASE = os.path.dirname(os.path.abspath('__file__'))")
g = {"__name__": "not_main", "__file__": os.path.abspath("test_excel_orient.py")}
exec(code, g)
g["make_wide_xlsx"]()
ok1 = try_convert("fresh")

print("== 2) 等待 3 秒后重试同一文件 ==")
time.sleep(3)
ok2 = try_convert("after_wait")

print("\n结论: fresh_ok =", ok1, "| after_wait_ok =", ok2)
