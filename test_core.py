# -*- coding: utf-8 -*-
"""核心功能自动化测试: 生成样例 -> 全链路转换 -> 验证输出"""
import os
import sys
import traceback

BASE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE)

TD = os.path.join(BASE, "test_data")
OUT = os.path.join(TD, "out")
os.makedirs(OUT, exist_ok=True)

PASS, FAIL = [], []


def check(name, fn):
    try:
        ok, msg = fn()
        (PASS if ok else FAIL).append(name)
        print(f"[{'PASS' if ok else 'FAIL'}] {name}: {msg}")
    except Exception as e:
        FAIL.append(name)
        print(f"[FAIL] {name}: 异常 {e}")
        traceback.print_exc()


# ------------------------------------------------ 生成样例
def make_samples():
    from docx import Document
    from openpyxl import Workbook
    import pymupdf

    d = Document()
    d.add_heading("测试文档标题", 0)
    d.add_paragraph("这是第一段测试文字，用于验证 Word 转 PDF 功能。")
    d.add_paragraph("Second paragraph for testing.")
    d.add_table(rows=2, cols=2).cell(0, 0).text = "表格"
    d.save(os.path.join(TD, "sample.docx"))

    wb = Workbook()
    ws = wb.active
    ws.title = "数据"
    for r in range(1, 21):
        for c in range(1, 6):
            ws.cell(row=r, column=c, value=f"R{r}C{c}" if r > 1 else f"列{c}")
    wb.save(os.path.join(TD, "sample.xlsx"))

    doc = pymupdf.open()
    for i in range(3):
        page = doc.new_page()
        page.insert_text((72, 100), f"PDF Test Page {i+1}", fontsize=20)
        page.insert_text((72, 140), "用于测试的示例 PDF 文本内容。", fontsize=12, fontname="china-s")
    doc.save(os.path.join(TD, "sample.pdf"))
    doc.close()

    from PIL import Image
    Image.new("RGB", (400, 300), (66, 133, 244)).save(os.path.join(TD, "img1.png"))
    Image.new("RGB", (300, 400), (234, 67, 53)).save(os.path.join(TD, "img2.jpg"))


import docflow.converters as cv

log = lambda m: print("   ", m)


def o(name):
    return os.path.join(OUT, name)


def main():
    make_samples()
    check("docx->pdf", lambda: cv.word_to_pdf(os.path.join(TD, "sample.docx"), o("sample.pdf"), log=log))
    check("xlsx->pdf", lambda: cv.excel_to_pdf(os.path.join(TD, "sample.xlsx"), o("sample_x.pdf"), log=log))
    check("pdf->docx", lambda: cv.pdf_to_word(os.path.join(TD, "sample.pdf"), o("sample.pdf.docx"), log=log))
    check("pdf->images", lambda: cv.pdf_to_images(os.path.join(TD, "sample.pdf"), o("imgs"), log=log))
    check("images->pdf", lambda: cv.images_to_pdf(
        [os.path.join(TD, "img1.png"), os.path.join(TD, "img2.jpg")], o("imgs.pdf"), log=log))
    check("docx->txt", lambda: cv.word_to_txt(os.path.join(TD, "sample.docx"), o("sample.txt"), log=log))
    check("txt->docx", lambda: cv.txt_to_word(o("sample.txt"), o("sample_txt.docx"), log=log))
    check("xlsx->csv", lambda: cv.excel_to_csv(os.path.join(TD, "sample.xlsx"), o("csvs"), log=log))
    check("csv->xlsx", lambda: cv.csv_to_excel(
        os.path.join(OUT, "csvs", "sample_数据.csv"), o("back.xlsx"), log=log))
    check("pdf-merge", lambda: cv.pdf_merge(
        [os.path.join(TD, "sample.pdf"), o("imgs.pdf")], o("merged.pdf"), log=log))
    check("pdf-split", lambda: cv.pdf_split(os.path.join(TD, "sample.pdf"), o("split"), ranges=["1", "2-3"], log=log))
    check("pdf-encrypt", lambda: cv.pdf_encrypt(os.path.join(TD, "sample.pdf"), o("enc.pdf"), "123", log=log))
    check("pdf-decrypt", lambda: cv.pdf_decrypt(o("enc.pdf"), o("dec.pdf"), "123", log=log))

    # 输出有效性抽查
    import pymupdf
    for f, pages in (("sample.pdf", 1), ("merged.pdf", 4), ("enc.pdf", 3)):
        try:
            d = pymupdf.open(o(f))
            print(f"[CHECK] {f}: {len(d)} 页, {os.path.getsize(o(f))//1024}KB")
            d.close()
        except Exception as e:
            print(f"[CHECK] {f}: 打不开! {e}")
            FAIL.append(f)

    print(f"\n===== 结果: {len(PASS)} 通过, {len(FAIL)} 失败 =====")
    if FAIL:
        print("失败项:", FAIL)
    return 0 if not FAIL else 1


if __name__ == "__main__":
    sys.exit(main())
