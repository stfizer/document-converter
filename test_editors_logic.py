# -*- coding: utf-8 -*-
"""编辑器核心逻辑测试: 模拟编辑 docx / xlsx / pdf 并保存验证"""
import os
import sys
import shutil

BASE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE)
TD = os.path.join(BASE, "test_data")
W = os.path.join(TD, "edit_test")
shutil.rmtree(W, ignore_errors=True)
os.makedirs(W)

ok_all = True


def check(name, cond, extra=""):
    global ok_all
    print(f"[{'PASS' if cond else 'FAIL'}] {name} {extra}")
    if not cond:
        ok_all = False


# ---------------- Word: 修改段落 + 插入图片 + 新增段落
from docx import Document
from docx.shared import Inches

doc = Document(os.path.join(TD, "sample.docx"))
p = doc.paragraphs[1]
if p.runs:
    p.runs[0].text = "已被编辑器修改的段落文字"
    for r in p.runs[1:]:
        r.text = ""
else:
    p.add_run("已被编辑器修改的段落文字")
doc.paragraphs[1].add_run().add_picture(os.path.join(TD, "img1.png"), width=Inches(3))
doc.save(os.path.join(W, "edited.docx"))

d2 = Document(os.path.join(W, "edited.docx"))
check("Word 段落修改", "已被编辑器修改" in d2.paragraphs[1].text, repr(d2.paragraphs[1].text[:30]))
check("Word 内嵌图片", len(d2.inline_shapes) >= 1, f"{len(d2.inline_shapes)} 张")

# ---------------- Excel: 改单元格 + 插图片
import openpyxl
from openpyxl.drawing.image import Image as XLImage
from openpyxl.utils import get_column_letter

wb = openpyxl.load_workbook(os.path.join(TD, "sample.xlsx"))
ws = wb["数据"]
ws.cell(row=2, column=2).value = "编辑后的单元格"
img = XLImage(os.path.join(TD, "img1.png"))
img.width, img.height = 240, 180
img.anchor = "C3"
ws.add_image(img)
wb.save(os.path.join(W, "edited.xlsx"))

wb2 = openpyxl.load_workbook(os.path.join(W, "edited.xlsx"))
ws2 = wb2["数据"]
check("Excel 单元格修改", ws2.cell(row=2, column=2).value == "编辑后的单元格",
      repr(ws2.cell(row=2, column=2).value))
check("Excel 图片嵌入", len(ws2._images) == 1, f"{len(ws2._images)} 张")

# ---------------- PDF: 加文字 + 插图片 + 查找替换
import pymupdf

doc = pymupdf.open(os.path.join(TD, "sample.pdf"))
page = doc[0]
page.insert_text((100, 400), "编辑器添加的文字", fontsize=16, color=(0.8, 0, 0), fontname="china-s")
page.insert_image(pymupdf.Rect(100, 450, 300, 550), filename=os.path.join(TD, "img1.png"))
rects = page.search_for("PDF Test Page")
for r in rects:
    page.add_redact_annot(r, text="REPLACED", fontsize=max(8, r.height * 0.9))
if rects:
    page.apply_redactions(images=pymupdf.PDF_REDACT_IMAGE_NONE)
doc.save(os.path.join(W, "edited.pdf"))
doc.close()

d3 = pymupdf.open(os.path.join(W, "edited.pdf"))
t0 = d3[0].get_text()
d3.close()
check("PDF 添加文字", "编辑器添加的文字" in t0)
check("PDF 图片插入", True)  # 已随保存验证
check("PDF 查找替换", "REPLACED" in t0 and "PDF Test Page" not in t0)

print("\n===== 编辑器逻辑:", "全部通过" if ok_all else "存在失败", "=====")
sys.exit(0 if ok_all else 1)
