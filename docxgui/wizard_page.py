"""第一页：选择 docx、处理模式（默认/AI）、模板，发起解析。"""
import os

from PyQt6.QtCore import pyqtSignal
from PyQt6.QtWidgets import (QButtonGroup, QCheckBox, QFileDialog, QGroupBox,
                             QHBoxLayout, QLabel, QLineEdit, QPushButton,
                             QRadioButton, QVBoxLayout, QWidget)

from docxai import config as aiconf
from docxflow.cli import DEFAULT_TEMPLATE


class WizardPage(QWidget):
    """收集 run 所需输入。解析动作交给主窗口用 worker 执行。"""

    start_requested = pyqtSignal(dict)
    open_settings = pyqtSignal()
    open_template = pyqtSignal(str)       # 路径可能为空串（调用方回退默认模板）

    def __init__(self):
        super().__init__()
        self._build()
        self.refresh_ai_status()

    def _build(self):
        root = QVBoxLayout(self)

        # 1. 源文档
        doc_box = QGroupBox("1. 选择待处理的 Word 文档")
        v1 = QVBoxLayout(doc_box)
        row1 = QHBoxLayout()
        self.docx_edit = QLineEdit()
        self.docx_edit.setPlaceholderText("别人撰写的论文 .docx")
        btn_docx = QPushButton("浏览...")
        btn_docx.clicked.connect(self._pick_docx)
        row1.addWidget(self.docx_edit)
        row1.addWidget(btn_docx)
        v1.addLayout(row1)
        self.strip_chk = QCheckBox("丢弃第一个数字编号一级标题之前的前置内容"
                                   "（封面、诚信承诺书、目录等）")
        v1.addWidget(self.strip_chk)
        root.addWidget(doc_box)

        # 2. 处理模式
        mode_box = QGroupBox("2. 扁平化处理方式")
        v2 = QVBoxLayout(mode_box)
        self.rb_default = QRadioButton(
            "默认模式：直接按 Word 样式规则解析（离线、零成本，"
            "适合标题规范的文档）")
        self.rb_ai = QRadioButton(
            "AI 模式：规则解析后自动按章调 AI 修正结构分类（不改实质内容，"
            "适合样式混乱的文档）")
        self.rb_default.setChecked(True)
        grp = QButtonGroup(self)
        grp.addButton(self.rb_default)
        grp.addButton(self.rb_ai)
        v2.addWidget(self.rb_default)
        v2.addWidget(self.rb_ai)
        row2 = QHBoxLayout()
        self.ai_status = QLabel()
        btn_settings = QPushButton("AI 配置...")
        btn_settings.clicked.connect(self.open_settings.emit)
        row2.addWidget(self.ai_status, 1)
        row2.addWidget(btn_settings)
        v2.addLayout(row2)
        root.addWidget(mode_box)

        # 3. 模板
        tpl_box = QGroupBox("3. 模板（run 构建与审计共用）")
        v3 = QVBoxLayout(tpl_box)
        row3 = QHBoxLayout()
        self.tpl_edit = QLineEdit()
        self.tpl_edit.setPlaceholderText(DEFAULT_TEMPLATE)
        btn_tpl = QPushButton("浏览...")
        btn_tpl.clicked.connect(self._pick_template)
        row3.addWidget(self.tpl_edit, 1)
        row3.addWidget(btn_tpl)
        btn_tpl_check = QPushButton("模板预览与修复...")
        btn_tpl_check.clicked.connect(
            lambda: self.open_template.emit(self.tpl_edit.text().strip()))
        row3.addWidget(btn_tpl_check)
        v3.addLayout(row3)
        root.addWidget(tpl_box)

        # 启动
        self.start_btn = QPushButton("开始解析并进入审阅")
        self.start_btn.setMinimumHeight(36)
        self.start_btn.clicked.connect(self._emit_start)
        root.addWidget(self.start_btn)
        root.addStretch(1)

    def refresh_ai_status(self):
        """配置弹窗关闭后刷新状态文字；未配置时 AI 模式仍可选但会再次引导。"""
        self.ai_status.setText("当前：" + aiconf.describe())

    def _pick_docx(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "选择 Word 文档", "", "Word 文档 (*.docx)")
        if path:
            self.docx_edit.setText(path)

    def _pick_template(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "选择模板", "", "Word 文档 (*.docx)")
        if path:
            self.tpl_edit.setText(path)

    def _emit_start(self):
        docx = self.docx_edit.text().strip()
        if not docx or not os.path.exists(docx):
            self.ai_status.setText("请先选择存在的 .docx 文档")
            return
        if not docx.lower().endswith(".docx"):
            self.ai_status.setText("只支持 .docx 格式")
            return
        if self.rb_ai.isChecked() and not aiconf.configured():
            self.ai_status.setText("AI 模式需要先完成 AI 配置")
            self.open_settings.emit()
            return
        self.start_requested.emit({
            "docx": os.path.abspath(docx),
            "mode": "ai" if self.rb_ai.isChecked() else "default",
            "template": os.path.abspath(self.tpl_edit.text().strip())
                        if self.tpl_edit.text().strip() else "",
            "strip_front": self.strip_chk.isChecked(),
        })
