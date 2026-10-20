# -*- coding: utf-8 -*-
"""Excel→PDF 左右(按列)分割验证"""
import os
import sys

BASE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE)
os.chdir(BASE)

import docflow.converters as cv
import pymupdf

TD = os.path.join(BASE, "test_data", "orient")
SRC = os.path.join(TD, "long_table.xlsx")


def pages(path):
    d = pymupdf.open(path)
    texts = [d[i].get_text() for i in range(len(d))]
    w, h = d[0].rect.width, d[0].rect.height
    d.close()
    return len(texts), texts, w, h


def check(name, cond, extra=""):
    print(f"[{'PASS' if cond else 'FAIL'}] {name} {extra}")
    return cond


def main():
    assert os.path.exists(SRC), "先运行 test_total_pages.py 生成 long_table.xlsx"
    log = lambda m: print("   ", m)
    ok_all = True

    # a) 左右分割 2 页 (10 列 → 每页 ~5 列, 行方向自然分页)
    d1 = os.path.join(TD, "colsplit2.pdf")
    cv.excel_to_pdf(src=SRC, dst=d1, log=log, landscape=True,
                    fit_wide=False, total_pages=2, split="col")
    n, texts, w, h = pages(d1)
    all_text = "".join(texts)
    first, last = texts[0], texts[-1]
    print(f"   左右分割2页: {n} 页, {w:.0f}x{h:.0f}")
    ok_all &= check("横向页面", w > h, f"{w:.0f}x{h:.0f}")
    ok_all &= check("第一组含左侧列(序号)", "序号" in first)
    ok_all &= check("第一组不含最右列(备注)", "备注" not in first)
    ok_all &= check("最后组含最右列(备注)", "备注" in last)
    ok_all &= check("全部 60 行数据存在", "测试单位第60号" in all_text
                    and "测试单位第1号" in all_text)

    # b) 左右分割 3 页
    d2 = os.path.join(TD, "colsplit3.pdf")
    cv.excel_to_pdf(src=SRC, dst=d2, log=log, landscape=True,
                    fit_wide=False, total_pages=3, split="col")
    n, texts, w, h = pages(d2)
    print(f"   左右分割3页: {n} 页")
    ok_all &= check("3 页模式正常输出", n >= 3, f"实际 {n} 页")
    ok_all &= check("列均分覆盖所有列", all_text.count("工作地点") >= 1
                    and "序号" in texts[0])

    # c) 回归: 上下分割 3 页仍然恰好 3 页
    d3 = os.path.join(TD, "rowsplit3.pdf")
    cv.excel_to_pdf(src=SRC, dst=d3, log=log, landscape=True,
                    fit_wide=True, total_pages=3, split="row")
    n, texts, w, h = pages(d3)
    print(f"   上下分割3页: {n} 页")
    ok_all &= check("上下分割仍恰好 3 页", n == 3, f"实际 {n} 页")

    # d) 兜底引擎左右分割
    d4 = os.path.join(TD, "colsplit_fb.pdf")
    cv._xlsx_text_fallback_pdf(src=SRC, dst=d4, landscape=True,
                               total_pages=2, split="col")
    n, texts, w, h = pages(d4)
    all_text = "".join(texts)
    print(f"   兜底左右分割: {n} 页")
    ok_all &= check("兜底横向输出", w > h, f"{w:.0f}x{h:.0f}")
    ok_all &= check("兜底含首列与末列内容", "序号" in all_text and "备注内容60" in all_text)

    print("\n=====", "全部通过" if ok_all else "存在失败", "=====")
    return 0 if ok_all else 1


if __name__ == "__main__":
    sys.exit(main())
