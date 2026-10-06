# -*- coding: utf-8 -*-
"""拖放逻辑验证: 构造合成拖放事件, 走 DropFilter 全流程"""
import os
import sys

BASE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE)

from PySide6.QtWidgets import QApplication
from PySide6.QtCore import QPoint, Qt, QMimeData, QUrl, QEvent
from PySide6.QtGui import QDragEnterEvent, QDropEvent

import docflow.gui as G

TD = os.path.join(BASE, "test_data")

app = QApplication(sys.argv)
win = G.App()

conv = win.pages["convert"]
merged = []
orig_add = conv.lst.addItem
conv.lst.addItem = lambda x: (merged.append(str(x)), orig_add(x))[1]

mime = QMimeData()
mime.setUrls([
    QUrl.fromLocalFile(os.path.join(TD, "sample.pdf")),
    QUrl.fromLocalFile(os.path.join(TD, "sample.docx")),
])
enter = QDragEnterEvent(QPoint(10, 10), Qt.CopyAction, mime,
                        Qt.LeftButton, Qt.NoModifier)
drop = QDropEvent(QPoint(10, 10), Qt.CopyAction, mime,
                  Qt.LeftButton, Qt.NoModifier)

# 找到挂在 lst 上的 DropFilter
filt = None
for c in conv.lst.children():
    if type(c).__name__ == "DropFilter":
        filt = c
        break
assert filt is not None, "lst 上没有 DropFilter"

# 转换页当前任务默认 "Word → PDF": pdf 应被忽略, docx 应加入
assert filt.eventFilter(conv.lst, enter) is True, "DragEnter 未被接受"
filt.eventFilter(conv.lst, drop)
print("added after drop (task=Word→PDF):", [os.path.basename(m) for m in merged])
assert any(m.endswith("sample.docx") for m in merged), "docx 未被添加"
assert not any(m.endswith("sample.pdf") for m in merged), "pdf 不应被加入 Word→PDF 任务"

# 切换任务到 PDF→图片, 再拖 docx+pdf: 只加 pdf
conv.cmb_task.setCurrentIndex(4)
merged.clear()
filt.eventFilter(conv.lst, enter)
filt.eventFilter(conv.lst, drop)
print("added after drop (task=PDF→图片):", [os.path.basename(m) for m in merged])
assert any(m.endswith("sample.pdf") for m in merged)
assert not any(m.endswith("sample.docx") for m in merged)

# 合并列表 filter: exts=.pdf
from docflow.drops import DropFilter
merged2 = []
pt = win.pages["pdftools"]
f2 = None
for c in pt.merge_lst.children():
    if type(c).__name__ == "DropFilter":
        f2 = c
assert f2 is not None
pt.merge_lst.addItem = lambda x: (merged2.append(str(x)), None)[1]
f2.eventFilter(pt.merge_lst, enter)
f2.eventFilter(pt.merge_lst, drop)
print("merge list after drop:", [os.path.basename(m) for m in merged2])
assert any(m.endswith("sample.pdf") for m in merged2)

# 编辑器页面拖放 filter 数量检查
for key, page in (("word", win.pages["word"]), ("excel", win.pages["excel"]),
                  ("pdf", win.pages["pdf"])):
    n = sum(1 for c in page.children() if type(c).__name__ == "DropFilter")
    print(f"{key} page drop filters:", n)
    assert n == 1

print("\nDROP-TEST ALL PASS")
