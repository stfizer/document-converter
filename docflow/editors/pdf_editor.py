# -*- coding: utf-8 -*-
"""PDF 可视化编辑页 - Qt 版: 点击添加文字 / 插入图片 / 查找替换 / 页面管理"""
import os

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, QFileDialog,
    QMessageBox, QDialog, QTextEdit, QLineEdit, QComboBox, QFormLayout,
    QGraphicsView, QGraphicsScene, QGraphicsPixmapItem, QInputDialog,
)
from PySide6.QtCore import Qt
from PySide6.QtGui import QPixmap, QImage, QColor

from docflow.drops import enable_page_drop


def pixmap_from_page(page, zoom):
    import pymupdf
    pix = page.get_pixmap(matrix=pymupdf.Matrix(zoom, zoom), alpha=False)
    img = QImage(pix.samples, pix.width, pix.height, pix.stride, QImage.Format_RGB888)
    return QPixmap.fromImage(img.copy())


class PdfEditorPage(QWidget):
    def __init__(self, status_cb, parent=None):
        super().__init__(parent)
        self.status_cb = status_cb
        self.doc = None
        self.path = None
        self.page_no = 0
        self.zoom = 1.5
        self.mode = None  # None / text / image
        self._build()
        enable_page_drop(self, self._drop_open, (".pdf",))

    def _drop_open(self, files):
        ok = skipped = 0
        for p in files:
            if p.lower().endswith(".pdf"):
                self.open_path(p)
                ok += 1
            else:
                skipped += 1
        if skipped:
            self.status_cb(f"忽略 {skipped} 个非 PDF 文件")

    def _build(self):
        lay = QVBoxLayout(self)
        lay.setContentsMargins(16, 12, 16, 12)
        lay.setSpacing(8)

        bar = QHBoxLayout()
        for text, cb in (("📂 打开 PDF", self.open_file),
                         ("💾 另存为", self.save_as)):
            b = QPushButton(text)
            b.clicked.connect(cb)
            bar.addWidget(b)
        sep = QLabel("|")
        sep.setStyleSheet("color:#cbd5e1;")
        bar.addWidget(sep)
        for text, m in (("➕ 点击添加文字", "text"), ("🖼 点击插入图片", "image")):
            b = QPushButton(text)
            b.clicked.connect(lambda _=False, mm=m: self._set_mode(mm))
            bar.addWidget(b)
        b_fr = QPushButton("🔍 查找替换文字")
        b_fr.clicked.connect(self._find_replace)
        bar.addWidget(b_fr)
        sep2 = QLabel("|")
        sep2.setStyleSheet("color:#cbd5e1;")
        bar.addWidget(sep2)
        b_del = QPushButton("🗑 删除当前页")
        b_del.clicked.connect(self._del_page)
        bar.addWidget(b_del)
        bar.addStretch(1)
        lay.addLayout(bar)

        nav = QHBoxLayout()
        b_prev = QPushButton("◀ 上一页")
        b_prev.clicked.connect(lambda: self._go(-1))
        b_next = QPushButton("下一页 ▶")
        b_next.clicked.connect(lambda: self._go(1))
        self.lbl_page = QLabel("0 / 0")
        self.lbl_page.setStyleSheet("font-weight:bold;")
        self.lbl_page.setAlignment(Qt.AlignCenter)
        self.lbl_page.setFixedWidth(80)
        nav.addWidget(b_prev)
        nav.addWidget(self.lbl_page)
        nav.addWidget(b_next)
        nav.addWidget(QLabel("缩放:"))
        for z, name in ((1.0, "100%"), (1.5, "150%"), (2.0, "200%")):
            b = QPushButton(name)
            b.setFixedWidth(72)
            b.clicked.connect(lambda _=False, zz=z: self._set_zoom(zz))
            nav.addWidget(b)
        self.lbl_mode = QLabel("")
        self.lbl_mode.setStyleSheet("color:#2563eb; font-weight:bold;")
        nav.addWidget(self.lbl_mode, stretch=1)
        lay.addLayout(nav)

        self.scene = QGraphicsScene(self)
        self.view = QGraphicsView(self.scene)
        self.view.setBackgroundBrush(QColor("#334155"))
        self.view.setRenderHints(self.view.renderHints())
        self.view.mousePressEvent = self._on_click
        lay.addWidget(self.view, stretch=1)

    # ------------------------------------------------ 文件与渲染
    def open_file(self):
        p, _ = QFileDialog.getOpenFileName(self, "打开 PDF", "", "PDF 文档 (*.pdf)")
        if p:
            self.open_path(p)

    def open_path(self, p):
        try:
            import pymupdf
            doc = pymupdf.open(p)
            if doc.needs_pass:
                pw, ok = QInputDialog.getText(self, "密码", "该 PDF 已加密, 请输入密码:",
                                              QLineEdit.Password)
                if not ok or not pw or not doc.authenticate(pw):
                    QMessageBox.critical(self, "错误", "密码错误或未输入, 无法打开")
                    doc.close()
                    return
            if self.doc:
                self.doc.close()
            self.doc, self.path = doc, p
            self.page_no = 0
            self._render()
            self.status_cb(f"已打开 {p}")
        except Exception as e:
            QMessageBox.critical(self, "打开失败", f"无法打开该文件:\n{e}")

    def _render(self):
        if self.doc is None:
            return
        self.page_no = max(0, min(self.page_no, len(self.doc) - 1))
        pm = pixmap_from_page(self.doc[self.page_no], self.zoom)
        self.scene.clear()
        self.scene.addPixmap(pm)
        self.scene.setSceneRect(0, 0, pm.width(), pm.height())
        self.lbl_page.setText(f"{self.page_no + 1} / {len(self.doc)}")

    def _go(self, d):
        if self.doc is None:
            return
        n = self.page_no + d
        if 0 <= n < len(self.doc):
            self.page_no = n
            self._render()

    def _set_zoom(self, z):
        self.zoom = z
        if self.doc is not None:
            self._render()

    def _set_mode(self, m):
        if self.doc is None:
            QMessageBox.information(self, "提示", "请先打开一个 PDF")
            return
        self.mode = None if self.mode == m else m
        tips = {"text": "模式: 点击页面位置添加文字（再点按钮取消）",
                "image": "模式: 点击页面位置插入图片（再点按钮取消）"}
        self.lbl_mode.setText(tips.get(self.mode, ""))
        self.view.setCursor(Qt.CrossCursor if self.mode else Qt.ArrowCursor)

    # ------------------------------------------------ 点击编辑
    def _on_click(self, ev):
        if self.doc is None or self.mode is None:
            QGraphicsView.mousePressEvent(self.view, ev)
            return
        sp = self.view.mapToScene(ev.position().toPoint())
        x, y = sp.x() / self.zoom, sp.y() / self.zoom
        if self.mode == "text":
            self._add_text_at(x, y)
        else:
            self._add_image_at(x, y)

    def _finish_mode(self):
        self.mode = None
        self.lbl_mode.setText("")
        self.view.setCursor(Qt.ArrowCursor)

    def _add_text_at(self, x, y):
        dlg = QDialog(self)
        dlg.setWindowTitle("添加文字")
        dlg.setFixedWidth(380)
        form = QFormLayout(dlg)
        txt = QTextEdit()
        txt.setMinimumHeight(80)
        form.addRow("文字内容:", txt)
        ent_size = QLineEdit("14")
        form.addRow("字号:", ent_size)
        color = QComboBox()
        color.addItems(["黑色", "红色", "蓝色", "白色"])
        form.addRow("颜色:", color)
        h = QHBoxLayout()
        ok = QPushButton("✔ 确定添加")
        ok.setObjectName("Accent")
        cancel = QPushButton("取消")
        ok.clicked.connect(dlg.accept)
        cancel.clicked.connect(dlg.reject)
        h.addWidget(ok)
        h.addWidget(cancel)
        form.addRow(h)
        if dlg.exec() != QDialog.Accepted:
            return
        text = txt.toPlainText().strip()
        if not text:
            return
        try:
            size = float(ent_size.text() or 14)
        except ValueError:
            size = 14
        cmap = {"黑色": (0, 0, 0), "红色": (0.8, 0, 0),
                "蓝色": (0, 0, 0.8), "白色": (1, 1, 1)}
        page = self.doc[self.page_no]
        page.insert_text((x, y + size), text, fontsize=size,
                         color=cmap[color.currentText()], fontname="china-s")
        self._finish_mode()
        self._render()
        self.status_cb("已添加文字（另存后生效）")

    def _add_image_at(self, x, y):
        p, _ = QFileDialog.getOpenFileName(self, "选择图片", "",
                                           "图片 (*.png *.jpg *.jpeg *.bmp *.gif *.webp);;所有文件 (*.*)")
        if not p:
            return
        try:
            import pymupdf
            from PIL import Image as PILImage
            page = self.doc[self.page_no]
            with PILImage.open(p) as im:
                iw, ih = im.size
            w = min(iw / 72.0 * 1.5, page.rect.width / 2)
            h = ih / iw * w
            page.insert_image(pymupdf.Rect(x, y, x + w, y + h), filename=p)
            self._finish_mode()
            self._render()
            self.status_cb("已插入图片（另存后生效）")
        except Exception as e:
            QMessageBox.critical(self, "错误", f"插入图片失败:\n{e}")

    def _find_replace(self):
        if self.doc is None:
            QMessageBox.information(self, "提示", "请先打开一个 PDF")
            return
        find, ok = QInputDialog.getText(self, "查找", "要查找的文字:")
        if not ok or not find.strip():
            return
        rep, ok = QInputDialog.getText(self, "替换 / 删除",
                                       f"将「{find}」替换为 (留空 = 删除该文字):")
        if not ok:
            return
        rep = rep.strip()
        n = 0
        import pymupdf
        for page in self.doc:
            rects = page.search_for(find)
            for r in rects:
                page.add_redact_annot(r, text=rep if rep else None,
                                      fontsize=max(8, r.height * 0.9) if rep else None)
            if rects:
                page.apply_redactions(images=pymupdf.PDF_REDACT_IMAGE_NONE)
                n += len(rects)
        self._render()
        action = "替换" if rep else "删除"
        self.status_cb(f"共{action} {n} 处「{find}」（另存后生效）")
        QMessageBox.information(self, "完成",
                                f"共{action} {n} 处。注意: 文字样式可能与原文略有差异。")

    def _del_page(self):
        if self.doc is None:
            return
        if len(self.doc) <= 1:
            QMessageBox.information(self, "提示", "至少需要保留一页")
            return
        if QMessageBox.question(self, "确认",
                                f"确定删除第 {self.page_no + 1} 页吗？") != QMessageBox.Yes:
            return
        self.doc.delete_page(self.page_no)
        self.page_no = min(self.page_no, len(self.doc) - 1)
        self._render()
        self.status_cb("已删除该页")

    # ------------------------------------------------ 保存
    def save_as(self):
        if self.doc is None:
            QMessageBox.information(self, "提示", "请先打开一个 PDF")
            return
        p, _ = QFileDialog.getSaveFileName(self, "另存 PDF",
                                           os.path.basename(self.path), "PDF 文档 (*.pdf)")
        if not p:
            return
        try:
            self.doc.save(p, garbage=3, deflate=True)
        except Exception as e:
            QMessageBox.critical(self, "保存失败", str(e))
            return
        self.status_cb(f"已保存: {p}")
        QMessageBox.information(self, "保存成功", f"已保存到:\n{p}")
