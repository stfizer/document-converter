# -*- coding: utf-8 -*-
"""启动 GUI, 逐页离屏截图 (QWidget.grab, 无需前台窗口)"""
import os
import sys

BASE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE)

from PySide6.QtWidgets import QApplication
from PySide6.QtCore import QTimer

import docflow.gui as G

SHOTS = os.path.join(BASE, "test_data", "shots")


def main():
    os.makedirs(SHOTS, exist_ok=True)
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    app.setStyleSheet(G.QSS)
    win = G.App()
    win.show()

    keys = [("convert", "1_convert.png"), ("pdftools", "2_pdftools.png"),
            ("word", "3_word.png"), ("excel", "4_excel.png"), ("pdf", "5_pdf.png")]
    state = {"i": 0}

    def step():
        if state["i"] >= len(keys):
            app.quit()
            return
        key, name = keys[state["i"]]
        win.nav.setCurrentRow(list(G.App.PAGES).index([p for p in G.App.PAGES if p[2] == key][0]))
        state["i"] += 1

        def snap():
            pm = win.grab()
            pm.save(os.path.join(SHOTS, name))
            print("shot:", name)
            QTimer.singleShot(150, step)

        QTimer.singleShot(400, snap)

    QTimer.singleShot(1200, step)
    sys.exit(app.exec())
    print("all shots done")


if __name__ == "__main__":
    main()
