# -*- coding: utf-8 -*-
"""高保真 PDF->Word 验证: 样例PDF -> 两种模式转换 -> Word COM 转回PDF -> 渲染对比"""
import os
import sys

BASE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE)
TD = os.path.join(BASE, "test_data")
OUT = os.path.join(TD, "hifi")
os.makedirs(OUT, exist_ok=True)

import pymupdf

FONT_SONG = r"C:\Windows\Fonts\simsun.ttc"
FONT_HEI = r"C:\Windows\Fonts\simhei.ttf"
FONT_YAHEI = r"C:\Windows\Fonts\msyh.ttc"


def make_sample():
    """构造复杂中文样例 PDF: 标题/正文/不同字号颜色/图片/表格线/色块"""
    path = os.path.join(OUT, "sample_complex.pdf")
    doc = pymupdf.open()
    page = doc.new_page(width=595, height=842)  # A4
    fhei = pymupdf.Font(fontfile=FONT_HEI)
    fsong = pymupdf.Font(fontfile=FONT_SONG)
    fyahei = pymupdf.Font(fontfile=FONT_YAHEI)
    page.insert_font(fontname="hei", fontfile=FONT_HEI)
    page.insert_font(fontname="song", fontfile=FONT_SONG)
    page.insert_font(fontname="yahei", fontfile=FONT_YAHEI)

    # 标题 (黑体 22pt 红色)
    page.insert_text((150, 80), "季度财务报告", fontsize=22,
                     fontname="hei", color=(0.7, 0, 0))
    # 副标题
    page.insert_text((210, 110), "2026 年第三季度 · 内部资料", fontsize=11,
                     fontname="yahei", color=(0.3, 0.3, 0.3))
    # 分隔线
    page.draw_line(pymupdf.Point(60, 125), pymupdf.Point(535, 125),
                   color=(0.2, 0.4, 0.8), width=1.2)
    # 正文多行 (宋体 12pt)
    lines = [
        "本季度公司整体经营状况良好，营业收入同比增长 18.7%。",
        "核心业务板块毛利率稳定在 42% 左右，费用控制成效显著。",
        "研发投入持续加大，新产品的市场反馈超出预期。",
        "The quick brown fox jumps over the lazy dog. 0123456789",
        "下半年将围绕三大战略方向展开：产品升级、渠道下沉与海外拓展。",
    ]
    y = 160
    for ln in lines:
        page.insert_text((60, y), ln, fontsize=12, fontname="song")
        y += 24
    # 色块 + 白字
    page.draw_rect(pymupdf.Rect(60, 300, 280, 330), color=None,
                   fill=(0.15, 0.39, 0.92))
    page.insert_text((75, 320), "重点提示：本报告未经审计", fontsize=12,
                     fontname="hei", color=(1, 1, 1))
    # 表格 (画线 + 单元格文字)
    tx, ty, cw, ch = 60, 360, 150, 28
    for i in range(4):  # 竖线
        page.draw_line(pymupdf.Point(tx + i * cw, ty),
                       pymupdf.Point(tx + i * cw, ty + ch * 4),
                       color=(0, 0, 0), width=0.8)
    for j in range(5):  # 横线
        page.draw_line(pymupdf.Point(tx, ty + j * ch),
                       pymupdf.Point(tx + 3 * cw, ty + j * ch),
                       color=(0, 0, 0), width=0.8)
    cells = [["项目", "Q2", "Q3", "环比"],
             ["营收(万)", "1,250", "1,484", "+18.7%"],
             ["毛利(万)", "512", "623", "+21.7%"],
             ["净利(万)", "210", "268", "+27.6%"]]
    for r, row in enumerate(cells):
        for c, cell in enumerate(row):
            page.insert_text((tx + c * cw + 10, ty + r * ch + 19), cell,
                             fontsize=11, fontname="song",
                             color=(0, 0, 0.5) if r == 0 else (0, 0, 0))
    # 图片 (右下)
    from PIL import Image, ImageDraw
    img = Image.new("RGB", (220, 160), (255, 255, 255))
    d = ImageDraw.Draw(img)
    for i in range(160):
        d.line([(0, i), (220, i)], fill=(66 + i, 133 - i // 2, 244 - i))
    d.rectangle([20, 30, 200, 130], outline=(30, 30, 30), width=3)
    d.text((55, 70), "TREND CHART", fill=(20, 20, 20))
    ipath = os.path.join(OUT, "chart.png")
    img.save(ipath)
    page.insert_image(pymupdf.Rect(330, 360, 550, 520), filename=ipath)
    # 页脚
    page.insert_text((60, 800), "第 1 页 · 机密文件请勿外传", fontsize=9,
                     fontname="song", color=(0.5, 0.5, 0.5))
    doc.save(path)
    doc.close()
    return path


def docx_to_pdf_via_word(docx_path, log=print):
    """Word COM: docx -> pdf (用于回转对比)"""
    import pythoncom
    import win32com.client
    pythoncom.CoInitialize()
    app = None
    try:
        app = win32com.client.DispatchEx("Word.Application")
        app.Visible = False
        try:
            app.DisplayAlerts = 0
        except Exception:
            pass
        s = os.path.abspath(docx_path)
        d = s[:-5] + "_回转.pdf" if s.lower().endswith(".docx") else s + ".pdf"
        doc = app.Documents.Open(s, ReadOnly=True)
        doc.ExportAsFixedFormat(os.path.abspath(d), 17)
        doc.Close(False)
        return d
    finally:
        try:
            if app is not None:
                app.Quit()
        except Exception:
            pass
        pythoncom.CoUninitialize()


def similarity(pdf_a, pdf_b):
    """渲染两份 PDF 第 1 页, 计算像素相似度 0~1"""
    import pymupdf
    import numpy as np
    a = pymupdf.open(pdf_a)[0].get_pixmap(dpi=80)
    b = pymupdf.open(pdf_b)[0].get_pixmap(dpi=80)
    arr_a = np.frombuffer(a.samples, dtype=np.uint8).reshape(a.height, a.width, a.n)[:, :, :3]
    arr_b = np.frombuffer(b.samples, dtype=np.uint8).reshape(b.height, b.width, b.n)[:, :, :3]
    h = min(arr_a.shape[0], arr_b.shape[0])
    w = min(arr_a.shape[1], arr_b.shape[1])
    diff = np.abs(arr_a[:h, :w].astype(int) - arr_b[:h, :w].astype(int))
    return 1 - diff.mean() / 255.0


def main():
    print("1) 生成复杂样例 PDF…")
    sample = make_sample()
    hifi_docx = os.path.join(OUT, "sample_hifi.docx")
    std_docx = os.path.join(OUT, "sample_std.docx")

    print("2) 高保真模式转换…")
    import docflow.pdf2docx_hifi as hifi
    ok, msg = hifi.pdf_to_docx_hifi(sample, hifi_docx, log=lambda m: print("   ", m))
    print("   ", msg)

    print("3) 标准 pdf2docx 模式转换…")
    import docflow.converters as cv
    ok2, msg2 = cv.pdf_to_word(sample, std_docx, log=lambda m: None)
    print("   ", msg2)

    print("4) Word COM 回转 PDF (验证排版)…")
    pdf_hifi = docx_to_pdf_via_word(hifi_docx)
    pdf_std = docx_to_pdf_via_word(std_docx)
    print("   ", pdf_hifi)
    print("   ", pdf_std)

    print("5) 与原 PDF 像素对比:")
    s_hifi = similarity(sample, pdf_hifi)
    s_std = similarity(sample, pdf_std)
    print(f"   高保真模式相似度: {s_hifi * 100:.2f}%")
    print(f"   标准模式相似度:   {s_std * 100:.2f}%")
    print(f"   提升: {(s_hifi - s_std) * 100:+.2f} 个百分点")
    return 0


if __name__ == "__main__":
    sys.exit(main())
