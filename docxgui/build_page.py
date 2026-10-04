"""第三页：外部 run 进程的实时日志与结果。"""
import os

from PyQt6.QtCore import QUrl, pyqtSignal
from PyQt6.QtGui import QDesktopServices, QTextCursor
from PyQt6.QtWidgets import (QHBoxLayout, QLabel, QPlainTextEdit, QPushButton,
                             QVBoxLayout, QWidget)


class BuildPage(QWidget):
    """只展示与控制 RunProcess，自身不实现任何构建逻辑。"""

    cancel_requested = pyqtSignal()
    back_requested = pyqtSignal()

    def __init__(self, output_dir: str):
        super().__init__()
        self._output_dir = output_dir
        layout = QVBoxLayout(self)

        self.title = QLabel("正在生成 Word（run：架构图 → 构建 → 目录 → 审计）")
        layout.addWidget(self.title)

        self.log = QPlainTextEdit()
        self.log.setReadOnly(True)
        self.log.setLineWrapMode(QPlainTextEdit.LineWrapMode.NoWrap)
        layout.addWidget(self.log, 1)

        row = QHBoxLayout()
        self.btn_back = QPushButton("返回审阅")
        self.btn_back.clicked.connect(self.back_requested.emit)
        row.addWidget(self.btn_back)
        row.addStretch(1)
        self.btn_open = QPushButton("打开 output 目录")
        self.btn_open.clicked.connect(
            lambda: QDesktopServices.openUrl(
                QUrl.fromLocalFile(self._output_dir)))
        self.btn_cancel = QPushButton("终止构建")
        self.btn_cancel.clicked.connect(self.cancel_requested.emit)
        row.addWidget(self.btn_open)
        row.addWidget(self.btn_cancel)
        layout.addLayout(row)

    def reset(self):
        """每次进入构建页前清空旧日志与状态。"""
        self.log.clear()
        self.title.setText("正在生成 Word（run：架构图 → 构建 → 目录 → 审计）")
        self.btn_cancel.setEnabled(True)
        self.btn_back.setEnabled(False)

    def append_log(self, text: str):
        self.log.moveCursor(QTextCursor.MoveOperation.End)
        self.log.insertPlainText(text)
        self.log.moveCursor(QTextCursor.MoveOperation.End)

    def show_result(self, ok: bool):
        self.btn_cancel.setEnabled(False)
        self.btn_back.setEnabled(True)
        if ok:
            self.title.setText("构建成功完成，成品在 output 目录（可查看审计结果）")
        else:
            self.title.setText("构建未成功（进程非零退出），请查看上方日志定位问题")
