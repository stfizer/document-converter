# -*- coding: utf-8 -*-
"""验证自动摆正: 构造"内容旋转90°"的问题PDF -> 修复 -> 断言页横+字横排; 正常PDF不误修"""
import os
import sys

BASE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE)
os.chdir(BASE)

import docflow.converters as cv
import docflow.pdf2docx_hifi  # noqa 确保包完整
import pymupdf

TD = os.path.join(BASE, "test_data", "orient")


def make_rotated(src_pdf, out_pdf, target_swap=True):
    """模拟 Excel/WPS 的旋转 bug: 正向横向页 -> 内容旋转90°放入交换尺寸页"""
    src = pymupdf.open(src_pdf)
    d = pymupdf.open()
    for page in src:
        if target_swap:
            W, H = page.rect.height, page.rect.width     # 交换 = 页面方向改变
        else:
            W, H = page.rect.width, page.rect.height
        p = d.new_page(width=W, height=H)
        p.show_pdf_page(p.rect, src, page.number, rotate=90)
    d.save(out_pdf)
    d.close()
    src.close()


def orient(path):
    d = pymupdf.open(path)
    h, v = cv._text_orientation(d[0])
    w, hh = d[0].rect.width, d[0].rect.height
    n = len(d)
    d.close()
    return w, hh, h, v, n


def main():
    good = os.path.join(TD, "out_landscape.pdf")     # 已验证的正向横向 PDF
    assert os.path.exists(good), "先运行 test_excel_orient.py"

    print("1) 构造旋转问题 PDF (模拟 Excel/WPS 导出 bug)…")
    bad = os.path.join(TD, "rotated_problem.pdf")
    make_rotated(good, bad, target_swap=True)
    w, hh, h, v, n = orient(bad)
    print(f"   问题PDF: {w:.0f}x{hh:.0f} 横排行={h} 竖排行={v} 页数={n}")
    assert v > h, "构造失败: 应为竖排内容"

    print("2) 执行自动摆正…")
    cv._auto_fix_rotation(bad, log=lambda m: print("   ", m))
    w, hh, h, v, n = orient(bad)
    print(f"   修复后: {w:.0f}x{hh:.0f} 横排行={h} 竖排行={v}")
    ok1 = h > v
    ok2 = w > hh           # 页面横向
    print(f"   [{'PASS' if ok1 else 'FAIL'}] 文字正向横排")
    print(f"   [{'PASS' if ok2 else 'FAIL'}] 页面横向 (宽>高)")

    print("3) 回归: 正常横向 PDF 不应被误改…")
    normal = os.path.join(TD, "out_landscape.pdf")
    before = orient(normal)
    cv._auto_fix_rotation(normal, log=lambda m: print("   ", m))
    after = orient(normal)
    same = (abs(before[0] - after[0]) < 1 and abs(before[1] - after[1]) < 1
            and before[2] == after[2])
    print(f"   [{'PASS' if same else 'FAIL'}] 正常文件未变化")
    ok1 &= ok2
    print("\n=====", "全部通过" if (ok1 and ok2 and same) else "存在失败", "=====")
    return 0 if (ok1 and ok2 and same) else 1


if __name__ == "__main__":
    sys.exit(main())
