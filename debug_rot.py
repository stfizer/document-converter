# -*- coding: utf-8 -*-
"""诊断 8 种旋转组合的文字方向"""
import pymupdf

bad = "test_data/orient/rotated_problem.pdf"
src = pymupdf.open(bad)
pw, ph = src[0].rect.width, src[0].rect.height
print("源页: %.0fx%.0f" % (pw, ph))


def hv(page):
    h = v = 0
    for b in page.get_text("dict")["blocks"]:
        if b.get("type") != 0:
            continue
        for l in b.get("lines", []):
            x0, y0, x1, y1 = l["bbox"]
            if (x1 - x0) >= (y1 - y0):
                h += 1
            else:
                v += 1
    return h, v


for (W, H) in [(ph, pw), (pw, ph)]:
    for rot in (0, 90, 180, 270):
        t = pymupdf.open()
        tp = t.new_page(width=W, height=H)
        tp.show_pdf_page(tp.rect, src, 0, rotate=rot)
        h, v = hv(tp)
        print("W=%.0f H=%.0f rot=%3d -> 横排=%d 竖排=%d %s"
              % (W, H, rot, h, v, "<<< 字正" if h > v else ""))
        t.close()
src.close()
