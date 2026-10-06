# -*- coding: utf-8 -*-
"""通用拖放支持: 事件过滤器方式, 让任意 widget / 整页接受文件拖入"""
import os

from PySide6.QtCore import QEvent, QObject


class DropFilter(QObject):
    """事件过滤器: 让现有 widget 支持拖入文件, 无需子类化"""

    def __init__(self, widget, on_files, exts=None, highlight=True):
        super().__init__(widget)
        self.widget = widget
        self.on_files = on_files
        self.exts = [e.lower() for e in exts] if exts else None
        self.highlight = highlight
        widget.setAcceptDrops(True)
        widget.installEventFilter(self)

    def _highlight(self, on):
        if not self.highlight:
            return
        self.widget.setStyleSheet(
            "border: 2px dashed #2563eb; border-radius: 8px; background: #eff6ff;"
            if on else "")

    def eventFilter(self, obj, ev):
        t = ev.type()
        if t == QEvent.DragEnter and ev.mimeData().hasUrls():
            ev.acceptProposedAction()
            self._highlight(True)
            return True
        if t == QEvent.DragMove:
            ev.acceptProposedAction()
            return True
        if t == QEvent.DragLeave:
            self._highlight(False)
            return True
        if t == QEvent.Drop:
            self._highlight(False)
            files = []
            for u in ev.mimeData().urls():
                p = u.toLocalFile() if u.isLocalFile() else None
                if not p:
                    continue
                if os.path.isdir(p):
                    for root, _dirs, names in os.walk(p):
                        for n in names:
                            files.append(os.path.join(root, n))
                else:
                    files.append(p)
            if self.exts is not None:
                files = [f for f in files if f.lower().endswith(tuple(self.exts))]
            if files:
                self.on_files(files)
            ev.acceptProposedAction()
            return True
        return False


def enable_page_drop(page, on_files, exts):
    """让整个页面(含所有子控件)接受拖放"""
    f = DropFilter(page, on_files, exts, highlight=False)
    page.setAcceptDrops(True)
    page.installEventFilter(f)
    for w in page.findChildren(QObject):
        if w.isWidgetType():
            w.setAcceptDrops(True)
            w.installEventFilter(f)
    return f
