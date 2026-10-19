# -*- coding: utf-8 -*-
"""文档转换助手 - 主窗口 (PySide6)"""
import os

from PySide6.QtCore import Qt, QThread, Signal, QProcess, QEvent, QObject
from PySide6.QtGui import QIcon, QFont
from PySide6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QListWidget,
    QListWidgetItem, QStackedWidget, QPushButton, QComboBox, QCheckBox,
    QLineEdit, QProgressBar, QPlainTextEdit, QFileDialog, QMessageBox,
    QGroupBox, QFrame, QTabWidget, QTreeWidget, QTreeWidgetItem, QHeaderView,
    QTextEdit, QInputDialog, QApplication, QSizePolicy, QSplitter, QAbstractItemView,
)

from docflow import converters as cv
from docflow import pdf2docx_hifi
from docflow.drops import DropFilter, enable_page_drop
from docflow.editors.word_editor import WordEditorPage
from docflow.editors.excel_editor import ExcelEditorPage
from docflow.editors.pdf_editor import PdfEditorPage

APP_TITLE = "文档转换助手 v1.0"

QSS = """
* { font-family: "Microsoft YaHei UI"; font-size: 13px; color: #0f172a; }
QMainWindow, QWidget { background: #f1f5f9; }
#Sidebar { background: #0f172a; }
#Sidebar QLabel#Logo { color: white; font-size: 17px; font-weight: bold; background: transparent; }
#Sidebar QLabel#Sub { color: #64748b; font-size: 10px; background: transparent; }
#Sidebar QLabel#Foot { color: #475569; font-size: 10px; background: transparent; }
QListWidget#Nav { background: transparent; border: none; outline: none; }
QListWidget#Nav::item { color: #cbd5e1; padding: 12px 14px; border-radius: 8px; margin: 2px 10px; }
QListWidget#Nav::item:selected { background: #1d4ed8; color: white; }
QListWidget#Nav::item:hover:!selected { background: #1e293b; }
QGroupBox { background: white; border: 1px solid #e2e8f0; border-radius: 10px;
            margin-top: 12px; padding-top: 6px; font-weight: bold; color: #334155; }
QGroupBox::title { subcontrol-origin: margin; left: 12px; padding: 0 4px; }
QPushButton { background: white; border: 1px solid #cbd5e1; border-radius: 7px;
              padding: 7px 14px; }
QPushButton:hover { background: #eff6ff; border-color: #93c5fd; color: #1d4ed8; }
QPushButton:pressed { background: #dbeafe; }
QPushButton:disabled { color: #94a3b8; background: #f8fafc; }
QPushButton#Accent { background: #2563eb; border: none; color: white; font-weight: bold; padding: 9px 22px; }
QPushButton#Accent:hover { background: #1d4ed8; }
QPushButton#Accent:pressed { background: #1e40af; }
QLineEdit, QComboBox, QSpinBox { background: white; border: 1px solid #cbd5e1;
              border-radius: 7px; padding: 6px 10px; }
QPlainTextEdit#Log, QTextEdit { background: #0f172a; color: #e2e8f0; border: none;
              border-radius: 8px; font-family: Consolas; font-size: 12px; }
QListWidget, QTreeWidget { background: white; border: 1px solid #e2e8f0; border-radius: 8px; }
QProgressBar { background: #e2e8f0; border: none; border-radius: 7px; height: 14px; text-align: center; color: #334155; }
QProgressBar::chunk { background: #2563eb; border-radius: 7px; }
QLabel#Status { background: #e2e8f0; color: #334155; padding: 4px 12px; }
QTabWidget::pane { border: 1px solid #e2e8f0; border-radius: 8px; background: white; }
QTabBar::tab { padding: 8px 18px; }
QTableWidget { gridline-color: #f1f5f9; }
QHeaderView::section { background: #f8fafc; border: none; border-bottom: 1px solid #e2e8f0; padding: 5px; }
"""


# ================================================================ 转换页
class ConvertWorker(QThread):
    log_sig = Signal(str)
    done_sig = Signal(bool)
    progress_sig = Signal(int)

    def __init__(self, task_name, files, out_dir, hifi=False,
                 excel_landscape=None, excel_fitwide=False):
        super().__init__()
        self.task_name = task_name
        self.files = files
        self.out_dir = out_dir
        self.hifi = hifi
        self.excel_landscape = excel_landscape
        self.excel_fitwide = excel_fitwide

    def run(self):
        ok_all = True
        filt, func, ext = cv_tasks()[self.task_name]
        total = len(self.files)
        for i, src in enumerate(self.files, 1):
            self.log_sig.emit(f"[{i}/{total}] {os.path.basename(src)}")
            try:
                base = os.path.splitext(os.path.basename(src))[0]
                out_base = self.out_dir or os.path.dirname(src)
                if func == "pdf_images":
                    dst = os.path.join(out_base, base + "_图片")
                    ok, msg = cv.pdf_to_images(src, dst, log=self.log_sig.emit)
                elif func == "images_pdf":
                    dst = os.path.join(out_base, base + ".pdf")
                    ok, msg = cv.images_to_pdf(self.files, dst, log=self.log_sig.emit)
                    if ok:
                        self.log_sig.emit(f"  ✔ 输出: {dst}")
                        self.progress_sig.emit(100)
                        break
                elif func == "xlsx_csv":
                    dst = os.path.join(out_base, base + "_CSV")
                    ok, msg = cv.excel_to_csv(src, dst, log=self.log_sig.emit)
                elif func is cv.excel_to_pdf:
                    dst = os.path.join(out_base, base + ext)
                    ok, msg = cv.excel_to_pdf(src, dst, log=self.log_sig.emit,
                                              landscape=self.excel_landscape,
                                              fit_wide=self.excel_fitwide)
                    if ok:
                        self.log_sig.emit(f"  ✔ 输出: {dst}")
                elif func is cv.pdf_to_word and self.hifi:
                    dst = os.path.join(out_base, base + ext)
                    ok, msg = pdf2docx_hifi.pdf_to_docx_hifi(src, dst, log=self.log_sig.emit)
                    if ok:
                        self.log_sig.emit(f"  ✔ 输出: {dst}")
                else:
                    dst = os.path.join(out_base, base + ext)
                    ok, msg = func(src, dst, log=self.log_sig.emit)
                    if ok:
                        self.log_sig.emit(f"  ✔ 输出: {dst}")
                self.log_sig.emit(("  ✔ " if ok else "  ✘ ") + msg)
                ok_all = ok_all and ok
            except Exception as e:
                ok_all = False
                self.log_sig.emit(f"  ✘ 转换出错: {e}")
            self.progress_sig.emit(int(i * 100 / total))
        self.log_sig.emit("—— 全部任务结束 ——")
        self.done_sig.emit(ok_all)


def cv_tasks():
    return {
        "Word → PDF": (("Word 文档", "*.docx *.doc"), cv.word_to_pdf, ".pdf"),
        "Excel → PDF": (("Excel 工作簿", "*.xlsx *.xlsm *.xls"), cv.excel_to_pdf, ".pdf"),
        "PPT → PDF": (("PPT 演示文稿", "*.pptx *.ppt"), cv.ppt_to_pdf, ".pdf"),
        "PDF → Word": (("PDF 文档", "*.pdf"), cv.pdf_to_word, ".docx"),
        "PDF → 图片 (PNG)": (("PDF 文档", "*.pdf"), "pdf_images", None),
        "图片 → PDF": (("图片文件", "*.png *.jpg *.jpeg *.bmp *.gif *.webp"), "images_pdf", ".pdf"),
        "Word → 纯文本 TXT": (("Word 文档", "*.docx"), cv.word_to_txt, ".txt"),
        "TXT → Word": (("文本文件", "*.txt"), cv.txt_to_word, ".docx"),
        "Excel → CSV": (("Excel 工作簿", "*.xlsx *.xlsm"), "xlsx_csv", None),
        "CSV → Excel": (("CSV 文件", "*.csv"), cv.csv_to_excel, ".xlsx"),
    }


class ConvertPage(QWidget):
    def __init__(self, status_cb, parent=None):
        super().__init__(parent)
        self.status_cb = status_cb
        self.worker = None
        self._build()

    def _build(self):
        lay = QVBoxLayout(self)
        lay.setContentsMargins(16, 12, 16, 12)
        lay.setSpacing(10)

        g1 = QGroupBox("1️⃣  选择转换任务")
        h = QHBoxLayout(g1)
        self.cmb_task = QComboBox()
        self.cmb_task.addItems(list(cv_tasks().keys()))
        self.cmb_task.setFixedWidth(220)
        self.cmb_task.setFont(QFont("Microsoft YaHei UI", 11))
        h.addWidget(self.cmb_task)
        self.lbl_hint = QLabel("")
        self.lbl_hint.setStyleSheet("color:#64748b; font-weight:normal;")
        h.addWidget(self.lbl_hint)
        h.addStretch(1)
        self.chk_hifi = QCheckBox("高保真模式（文本框定位，排版与 PDF 一致）")
        self.chk_hifi.setToolTip(
            "仅对 PDF → Word 生效。\n"
            "开启后按原位置逐行还原文字、表格线、图片、颜色, 排版与 PDF 几乎一致;\n"
            "文字位于文本框中, 适合查看/打印/少量修改。\n"
            "关闭为常规模式 (流式排版, 便于大段编辑, 但复杂布局可能有偏差)。")
        h.addWidget(self.chk_hifi)

        # ---- Excel→PDF 页面设置 (仅该任务显示; 须在首次 _hint 前创建)
        self.row_ps = QWidget()
        hps = QHBoxLayout(self.row_ps)
        hps.setContentsMargins(4, 0, 4, 0)
        hps.addWidget(QLabel("页面方向:"))
        self.cmb_orient = QComboBox()
        self.cmb_orient.addItems(["跟随文档设置", "纵向", "横向（宽表格推荐）"])
        self.cmb_orient.setFixedWidth(190)
        hps.addWidget(self.cmb_orient)
        self.chk_fitwide = QCheckBox("列缩放至一页宽（列不再被拆分到其他页）")
        self.chk_fitwide.setChecked(True)
        self.chk_fitwide.setToolTip(
            "开启后忽略 Excel 原有打印缩放, 将所有列自动缩放到一页宽度内,\n"
            "行数较多时向下自动分页。宽表格 (如职位表/台账) 推荐配合横向使用。")
        hps.addWidget(self.chk_fitwide)
        hps.addStretch(1)
        lay.addWidget(self.row_ps)
        self.row_ps.setVisible(False)

        self.cmb_task.currentIndexChanged.connect(self._hint)
        self._hint(0)
        lay.addWidget(g1)

        g2 = QGroupBox("2️⃣  添加要转换的文件（可多选批量，也可直接拖入文件/文件夹）")
        v = QVBoxLayout(g2)
        bar = QHBoxLayout()
        for text, cb in (("➕ 添加文件", self.add_files),
                         ("➖ 移除选中", self.remove_sel),
                         ("🧹 清空列表", self.clear_files)):
            b = QPushButton(text)
            b.clicked.connect(cb)
            bar.addWidget(b)
        bar.addStretch(1)
        v.addLayout(bar)
        self.lst = QListWidget()
        self.lst.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.lst.setAlternatingRowColors(True)
        v.addWidget(self.lst)
        lay.addWidget(g2, stretch=3)

        g3 = QGroupBox("3️⃣  输出设置")
        v3 = QVBoxLayout(g3)
        self.chk_src = QCheckBox("输出到源文件所在文件夹（取消勾选可指定其他文件夹）")
        self.chk_src.setChecked(True)
        self.chk_src.toggled.connect(lambda on: self.ent_out.setEnabled(not on))
        v3.addWidget(self.chk_src)
        h3 = QHBoxLayout()
        self.ent_out = QLineEdit()
        self.ent_out.setEnabled(False)
        self.ent_out.setPlaceholderText("输出文件夹…")
        b3 = QPushButton("浏览…")
        b3.clicked.connect(self.pick_out)
        h3.addWidget(self.ent_out, stretch=1)
        h3.addWidget(b3)
        v3.addLayout(h3)
        h4 = QHBoxLayout()
        self.btn_go = QPushButton("🚀 开始转换")
        self.btn_go.setObjectName("Accent")
        self.btn_go.clicked.connect(self.run)
        self.pb = QProgressBar()
        self.pb.setFixedWidth(340)
        h4.addWidget(self.btn_go)
        h4.addWidget(self.pb)
        h4.addStretch(1)
        v3.addLayout(h4)
        lay.addWidget(g3)

        g4 = QGroupBox("转换日志")
        v4 = QVBoxLayout(g4)
        self.log = QPlainTextEdit()
        self.log.setObjectName("Log")
        self.log.setReadOnly(True)
        self.log.setMaximumBlockCount(500)
        v4.addWidget(self.log)
        lay.addWidget(g4, stretch=2)

        # 支持拖拽文件到列表添加
        DropFilter(self.lst, self._drop_files)

    def _drop_files(self, files):
        t = list(cv_tasks().values())[self.cmb_task.currentIndex()]
        exts = tuple(e.strip().lower().lstrip("*") for e in t[0][1].split() if e.startswith("*."))
        ok = [f for f in files if f.lower().endswith(exts)]
        bad = len(files) - len(ok)
        for fp in ok:
            self.lst.addItem(fp)
        if ok:
            self.status_cb(f"已通过拖放添加 {len(ok)} 个文件")
        if bad:
            self.append_log(f"⚠ 拖放: 忽略 {bad} 个与当前任务格式不符的文件（支持 {t[0][1]}）")
            self.status_cb(f"忽略 {bad} 个格式不符文件（支持 {t[0][1]}）")

    # ---- slots
    def _hint(self, idx):
        t = list(cv_tasks().values())[idx]
        self.lbl_hint.setText(f"支持格式: {t[0][1]}")
        task = self.cmb_task.currentText()
        self.chk_hifi.setEnabled(task == "PDF → Word")
        if task != "PDF → Word":
            self.chk_hifi.setChecked(False)
        self.row_ps.setVisible(task == "Excel → PDF")

    def add_files(self):
        t = list(cv_tasks().values())[self.cmb_task.currentIndex()]
        files, _ = QFileDialog.getOpenFileNames(self, "选择文件", "",
                                                f"{t[0][0]} ({t[0][1]});;所有文件 (*.*)")
        for f in files:
            self.lst.addItem(f)

    def remove_sel(self):
        for it in self.lst.selectedItems():
            self.lst.takeItem(self.lst.row(it))

    def clear_files(self):
        self.lst.clear()

    def pick_out(self):
        d = QFileDialog.getExistingDirectory(self, "选择输出文件夹")
        if d:
            self.ent_out.setText(d)

    def append_log(self, msg):
        self.log.appendPlainText(msg)

    def run(self):
        if self.worker is not None and self.worker.isRunning():
            return
        files = [self.lst.item(i).text() for i in range(self.lst.count())]
        if not files:
            QMessageBox.information(self, "提示", "请先添加要转换的文件")
            return
        out_dir = None if self.chk_src.isChecked() else self.ent_out.text().strip()
        if not self.chk_src.isChecked() and not out_dir:
            QMessageBox.information(self, "提示", "请选择输出文件夹")
            return
        task_name = self.cmb_task.currentText()
        self.log.appendPlainText(f"—— 开始任务: {task_name} ({len(files)} 个文件) ——")
        self.btn_go.setEnabled(False)
        self.pb.setValue(0)
        orient = self.cmb_orient.currentIndex()
        landscape = {0: None, 1: False, 2: True}[orient]
        self.worker = ConvertWorker(task_name, files, out_dir,
                                    hifi=self.chk_hifi.isChecked(),
                                    excel_landscape=landscape,
                                    excel_fitwide=self.chk_fitwide.isChecked())
        self.worker.log_sig.connect(self.append_log)
        self.worker.progress_sig.connect(self.pb.setValue)
        self.worker.done_sig.connect(self._finished)
        self.worker.start()
        self.status_cb("转换中…")

    def _finished(self, ok):
        self.btn_go.setEnabled(True)
        self.status_cb("转换完成 ✔" if ok else "转换完成（部分文件失败，见日志）")


# ================================================================ PDF 工具页
class PdfToolsPage(QWidget):
    def __init__(self, status_cb, parent=None):
        super().__init__(parent)
        self.status_cb = status_cb
        self._build()

    def _build(self):
        lay = QVBoxLayout(self)
        lay.setContentsMargins(16, 12, 16, 12)
        tabs = QTabWidget()
        lay.addWidget(tabs)

        # ---- 合并
        f1 = QWidget()
        v1 = QVBoxLayout(f1)
        v1.setContentsMargins(12, 12, 12, 12)
        self.merge_lst = QListWidget()
        self.merge_lst.setSelectionMode(QAbstractItemView.ExtendedSelection)
        v1.addWidget(self.merge_lst)
        h1 = QHBoxLayout()
        b_add = QPushButton("➕ 添加 PDF")
        b_add.clicked.connect(self._merge_add)
        b_rm = QPushButton("➖ 移除选中")
        b_rm.clicked.connect(lambda: [self.merge_lst.takeItem(self.merge_lst.row(i))
                                      for i in self.merge_lst.selectedItems()])
        b_go = QPushButton("🔗 合并为一个 PDF")
        b_go.setObjectName("Accent")
        b_go.clicked.connect(self._do_merge)
        for b in (b_add, b_rm):
            h1.addWidget(b)
        h1.addStretch(1)
        h1.addWidget(b_go)
        v1.addLayout(h1)
        tabs.addTab(f1, "合并 PDF")

        # ---- 拆分
        f2 = QWidget()
        v2 = QVBoxLayout(f2)
        v2.setContentsMargins(12, 12, 12, 12)
        h2 = QHBoxLayout()
        b2 = QPushButton("📂 选择 PDF")
        b2.clicked.connect(self._pick_split)
        self.lbl_split = QLabel("未选择")
        self.lbl_split.setStyleSheet("color:#64748b;")
        h2.addWidget(b2)
        h2.addWidget(self.lbl_split, stretch=1)
        v2.addLayout(h2)
        v2.addWidget(QLabel("页码范围（如 1-3,5,8-10；留空 = 每页一个文件）:"))
        self.ent_range = QLineEdit()
        v2.addWidget(self.ent_range)
        b2go = QPushButton("✂ 开始拆分")
        b2go.setObjectName("Accent")
        b2go.clicked.connect(self._do_split)
        v2.addWidget(b2go)
        v2.addStretch(1)
        tabs.addTab(f2, "拆分 PDF")

        # ---- 加密解密
        f3 = QWidget()
        v3 = QVBoxLayout(f3)
        v3.setContentsMargins(12, 12, 12, 12)
        h3 = QHBoxLayout()
        b3 = QPushButton("📂 选择 PDF")
        b3.clicked.connect(self._pick_enc)
        self.lbl_enc = QLabel("未选择")
        self.lbl_enc.setStyleSheet("color:#64748b;")
        h3.addWidget(b3)
        h3.addWidget(self.lbl_enc, stretch=1)
        v3.addLayout(h3)
        v3.addWidget(QLabel("密码:"))
        self.ent_pw = QLineEdit()
        self.ent_pw.setEchoMode(QLineEdit.Password)
        v3.addWidget(self.ent_pw)
        h3b = QHBoxLayout()
        b_enc = QPushButton("🔒 加密")
        b_dec = QPushButton("🔓 解密")
        b_enc.clicked.connect(lambda: self._do_crypt(True))
        b_dec.clicked.connect(lambda: self._do_crypt(False))
        h3b.addWidget(b_enc)
        h3b.addWidget(b_dec)
        h3b.addStretch(1)
        v3.addLayout(h3b)
        v3.addStretch(1)
        tabs.addTab(f3, "加密 / 解密")

        self.split_src = None
        self.enc_src = None

        # 拖放支持: 合并列表收多个 PDF; 拆分/加密页拖入单个 PDF 直接选中
        DropFilter(self.merge_lst, self._merge_drop, exts=(".pdf",))
        DropFilter(f2, self._split_drop, exts=(".pdf",), highlight=False)
        DropFilter(f3, self._split_drop, exts=(".pdf",), highlight=False)

    def _merge_drop(self, files):
        for fp in files:
            self.merge_lst.addItem(fp)
        self.status_cb(f"已通过拖放添加 {len(files)} 个 PDF 到合并列表")

    def _split_drop(self, files):
        p = files[0]
        self.split_src = p
        self.enc_src = p
        self.lbl_split.setText(p)
        self.lbl_split.setStyleSheet("color:#0f172a;")
        self.lbl_enc.setText(p)
        self.lbl_enc.setStyleSheet("color:#0f172a;")
        self.status_cb(f"已通过拖放选中: {os.path.basename(p)}（拆分/加密共用）")

    # 合并
    def _merge_add(self):
        files, _ = QFileDialog.getOpenFileNames(self, "选择 PDF", "", "PDF (*.pdf)")
        for f in files:
            self.merge_lst.addItem(f)

    def _do_merge(self):
        files = [self.merge_lst.item(i).text() for i in range(self.merge_lst.count())]
        if len(files) < 2:
            QMessageBox.information(self, "提示", "请至少添加 2 个 PDF")
            return
        dst, _ = QFileDialog.getSaveFileName(self, "保存合并结果", "合并结果.pdf", "PDF (*.pdf)")
        if not dst:
            return
        QApplication.setOverrideCursor(Qt.WaitCursor)
        QApplication.processEvents()
        ok, msg = cv.pdf_merge(files, dst, log=self.status_cb)
        QApplication.restoreOverrideCursor()
        QMessageBox.information(self, "完成" if ok else "失败", msg)

    # 拆分
    def _pick_split(self):
        p, _ = QFileDialog.getOpenFileName(self, "选择 PDF", "", "PDF (*.pdf)")
        if p:
            self.split_src = p
            self.lbl_split.setText(p)
            self.lbl_split.setStyleSheet("color:#0f172a;")

    def _do_split(self):
        if not self.split_src:
            QMessageBox.information(self, "提示", "请先选择 PDF")
            return
        rng = [r for r in self.ent_range.text().replace("，", ",").split(",") if r.strip()]
        out = os.path.join(os.path.dirname(self.split_src),
                           os.path.splitext(os.path.basename(self.split_src))[0] + "_拆分")
        QApplication.setOverrideCursor(Qt.WaitCursor)
        QApplication.processEvents()
        ok, msg = cv.pdf_split(self.split_src, out, ranges=rng or None, log=self.status_cb)
        QApplication.restoreOverrideCursor()
        QMessageBox.information(self, "完成" if ok else "失败", msg + f"\n输出目录: {out}")

    # 加密解密
    def _pick_enc(self):
        p, _ = QFileDialog.getOpenFileName(self, "选择 PDF", "", "PDF (*.pdf)")
        if p:
            self.enc_src = p
            self.lbl_enc.setText(p)
            self.lbl_enc.setStyleSheet("color:#0f172a;")

    def _do_crypt(self, encrypt):
        if not self.enc_src:
            QMessageBox.information(self, "提示", "请先选择 PDF")
            return
        pw = self.ent_pw.text()
        if not pw:
            QMessageBox.information(self, "提示", "请输入密码")
            return
        base = os.path.splitext(self.enc_src)[0]
        dst = base + ("_加密.pdf" if encrypt else "_解密.pdf")
        QApplication.setOverrideCursor(Qt.WaitCursor)
        QApplication.processEvents()
        if encrypt:
            ok, msg = cv.pdf_encrypt(self.enc_src, dst, pw)
        else:
            ok, msg = cv.pdf_decrypt(self.enc_src, dst, pw)
        QApplication.restoreOverrideCursor()
        if ok:
            self.status_cb(f"输出: {dst}")
        QMessageBox.information(self, "成功" if ok else "失败", msg + (f"\n输出: {dst}" if ok else ""))


# ================================================================ 感谢作者页
VIDEO_URL = ("https://www.bilibili.com/video/BV1p1daBMErQ/?"
             "spm_id_from=333.1387.homepage.video_card.click"
             "&vd_source=40169949e69893de442bdc5d692f88e5")
SPACE_URL = "https://space.bilibili.com/702444737?spm_id_from=333.1007.0.0"


class ThanksPage(QWidget):
    def __init__(self, status_cb, parent=None):
        super().__init__(parent)
        self.status_cb = status_cb
        self._build()

    def _build(self):
        lay = QVBoxLayout(self)
        lay.setContentsMargins(16, 12, 16, 12)
        lay.addStretch(1)

        card = QGroupBox("💖  感谢作者")
        card.setFixedWidth(640)
        v = QVBoxLayout(card)
        v.setContentsMargins(24, 16, 24, 20)
        v.setSpacing(14)

        title = QLabel(
            '<a href="' + SPACE_URL + '" '
            'style="color:#fb7299;font-weight:bold;font-size:16px;'
            'text-decoration:underline">bilibili@何止有点帅</a>'
            '<span style="color:#334155;font-size:13px">'
            ' · 软件作者（点击昵称访问 up 主空间）</span>')
        title.setTextInteractionFlags(Qt.TextBrowserInteraction)
        title.setOpenExternalLinks(True)
        v.addWidget(title)

        body = QLabel(
            '<html><body style="line-height:190%">'
            '<span style="color:#0f172a;font-size:14px">'
            "如果软件有帮到你，希望对 up 的软件表达感谢和支持，"
            "可以为 up 的另一个视频"
            '<a href="' + VIDEO_URL + '" '
            'style="color:#2563eb;font-weight:bold;font-size:15px;'
            'text-decoration:underline">增加播放量支持 →</a>，'
            "让 up 领取播放量奖励，感谢各位！"
            "</span></body></html>")
        body.setWordWrap(True)
        body.setMinimumHeight(84)
        body.setTextInteractionFlags(Qt.TextBrowserInteraction)
        body.setOpenExternalLinks(True)   # 点击蓝色链接用系统浏览器打开
        v.addWidget(body)

        tip = QLabel("↑ 点击蓝色「增加播放量支持 →」即可跳转视频（链接已内嵌，无需复制）")
        tip.setStyleSheet("color:#94a3b8; font-size:11px;")
        v.addWidget(tip)

        btn = QPushButton("▶  前往为 up 投放播放量支持")
        btn.setObjectName("Accent")
        btn.clicked.connect(self._open_video)
        h = QHBoxLayout()
        h.addStretch(1)
        h.addWidget(btn)
        h.addStretch(1)
        v.addLayout(h)

        lay.addWidget(card, alignment=Qt.AlignHCenter)
        lay.addStretch(2)

    def _open_video(self):
        from PySide6.QtGui import QDesktopServices
        from PySide6.QtCore import QUrl
        QDesktopServices.openUrl(QUrl(VIDEO_URL))
        self.status_cb("已在浏览器中打开 up 的视频，感谢支持！")


# ================================================================ 主窗口
class App(QMainWindow):
    PAGES = [("🔄", "格式转换", "convert"), ("🧰", "PDF 工具", "pdftools"),
             ("📝", "Word 编辑", "word"), ("📊", "Excel 编辑", "excel"),
             ("📕", "PDF 编辑", "pdf"), ("💖", "感谢作者", "thanks")]

    def __init__(self):
        super().__init__()
        self.setWindowTitle(APP_TITLE)
        self.resize(1180, 760)
        self.setMinimumSize(980, 640)
        self._build()

    def _build(self):
        central = QWidget()
        self.setCentralWidget(central)
        root = QHBoxLayout(central)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        # ---- 侧边栏
        sb = QWidget()
        sb.setObjectName("Sidebar")
        sb.setFixedWidth(190)
        v = QVBoxLayout(sb)
        v.setContentsMargins(0, 20, 0, 14)
        logo = QLabel("📄  文档转换助手")
        logo.setObjectName("Logo")
        logo.setAlignment(Qt.AlignCenter)
        sub = QLabel("Convert · Edit · Pack")
        sub.setObjectName("Sub")
        sub.setAlignment(Qt.AlignCenter)
        v.addWidget(logo)
        v.addWidget(sub)
        v.addSpacing(14)
        self.nav = QListWidget()
        self.nav.setObjectName("Nav")
        for icon, name, _k in self.PAGES:
            self.nav.addItem(QListWidgetItem(f"  {icon}  {name}"))
        self.nav.setCurrentRow(0)
        v.addWidget(self.nav, stretch=1)
        foot = QLabel("v1.0 · 本地离线运行\nWord / Excel / PPT\nPDF · 图片 · CSV")
        foot.setObjectName("Foot")
        foot.setAlignment(Qt.AlignLeft)
        foot.setContentsMargins(16, 0, 0, 0)
        v.addWidget(foot)
        root.addWidget(sb)

        # ---- 右侧
        right = QWidget()
        rv = QVBoxLayout(right)
        rv.setContentsMargins(0, 0, 0, 0)
        rv.setSpacing(0)
        self.stack = QStackedWidget()
        self.pages = {
            "convert": ConvertPage(self._status),
            "pdftools": PdfToolsPage(self._status),
            "word": WordEditorPage(self._status),
            "excel": ExcelEditorPage(self._status),
            "pdf": PdfEditorPage(self._status),
            "thanks": ThanksPage(self._status),
        }
        for p in self.pages.values():
            self.stack.addWidget(p)
        rv.addWidget(self.stack, stretch=1)
        self.status = QLabel("就绪 — 请选择左侧功能")
        self.status.setObjectName("Status")
        rv.addWidget(self.status)
        root.addWidget(right, stretch=1)

        self.nav.currentRowChanged.connect(self._switch)

    def _switch(self, row):
        self.stack.setCurrentIndex(row)

    def _status(self, msg):
        self.status.setText(str(msg))


def launch(icon_path=None):
    import sys
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    if icon_path and os.path.exists(icon_path):
        app.setWindowIcon(QIcon(icon_path))
    app.setStyleSheet(QSS)
    win = App()
    win.show()
    sys.exit(app.exec())
