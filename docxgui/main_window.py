"""GUI 主窗口：向导 → 审阅 → 构建 三页编排与全局设置入口。

主窗口只做编排与页面间数据传递：
- 解析/AI/构建均由 workers 后台执行，本类不写业务逻辑；
- 审阅确认后把块 render 成 config/章节/<docx名>.md（run 的既定落点），
  再由外部进程 `python3 main.py run` 完成全部后续阶段。
"""
import os

from PyQt6.QtGui import QAction
from PyQt6.QtWidgets import (QMainWindow, QMessageBox, QStackedWidget,
                             QStatusBar)

from docxai import HERE
from docxconvert.markdown import render
from docxgui.build_page import BuildPage
from docxgui.review_page import ReviewPage
from docxgui.settings_dialog import SettingsDialog
from docxgui.template_page import TemplateDialog
from docxgui.wizard_page import WizardPage
from docxgui.workers import ParseWorker, RunProcess
from docxflow.cli import DEFAULT_TEMPLATE

CONFIG_DIR = os.path.join(HERE, "config")
CHAPTERS_DIR = os.path.join(CONFIG_DIR, "章节")
IMAGES_ROOT = os.path.join(CONFIG_DIR, "images")
OUTPUT_DIR = os.path.join(HERE, "output")


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("thesis-docx-tool 图形界面")
        self.resize(1100, 760)
        self.setStatusBar(QStatusBar(self))

        self.wizard = WizardPage()
        self.review = ReviewPage(CONFIG_DIR)
        self.build_page = BuildPage(OUTPUT_DIR)
        self.stack = QStackedWidget()
        for w in (self.wizard, self.review, self.build_page):
            self.stack.addWidget(w)
        self.setCentralWidget(self.stack)

        self._parse_worker = None
        self._runner = RunProcess(HERE)
        self._cur = {}          # 本次任务参数（docx/mode/template/md 路径）

        self._wire()
        self._build_menu()
        self._refresh_status()

    # ---------------- 信号接线 ----------------
    def _wire(self):
        self.wizard.start_requested.connect(self._start_parse)
        self.wizard.open_settings.connect(self._open_settings)
        self.wizard.open_template.connect(self._open_template)
        self.review.back_requested.connect(lambda: self.stack.setCurrentIndex(0))
        self.review.confirm_requested.connect(self._confirm)
        self.review.settings_requested.connect(self._open_settings)
        self.review.status.connect(self.statusBar().showMessage)
        self.build_page.cancel_requested.connect(self._runner.stop)
        self.build_page.back_requested.connect(lambda: self.stack.setCurrentIndex(1))
        self._runner.log.connect(self.build_page.append_log)
        self._runner.finished_ok.connect(self.build_page.show_result)

    def _build_menu(self):
        settings_act = QAction("AI 配置...", self)
        settings_act.triggered.connect(self._open_settings)
        template_act = QAction("模板预览与修复...", self)
        template_act.triggered.connect(lambda: self._open_template(""))
        menu = self.menuBar().addMenu("设置")
        menu.addAction(template_act)
        menu.addAction(settings_act)

    def _open_template(self, path: str = ""):
        path = path or self.wizard.tpl_edit.text().strip() or DEFAULT_TEMPLATE
        dlg = TemplateDialog(os.path.abspath(path), HERE, self)
        dlg.template_changed.connect(self.wizard.tpl_edit.setText)
        dlg.exec()

    def _refresh_status(self):
        from docxai import config as aiconf
        self.statusBar().showMessage(aiconf.describe())

    def _open_settings(self):
        dlg = SettingsDialog(self)
        dlg.exec()
        self.wizard.refresh_ai_status()
        self._refresh_status()

    # ---------------- 阶段一：解析 ----------------
    def _start_parse(self, params: dict):
        base = os.path.splitext(os.path.basename(params["docx"]))[0]
        images_dir = os.path.join(IMAGES_ROOT, f"{base}_images")
        md_target = os.path.join(CHAPTERS_DIR, f"{base}.md")
        self._cur = {**params, "base": base,
                     "images_dir": images_dir, "md_target": md_target}

        self.statusBar().showMessage("正在解析 Word 并抽取图片...")
        worker = ParseWorker(params["docx"], images_dir, CONFIG_DIR,
                             params["strip_front"])
        self._parse_worker = worker
        worker.finished_blocks.connect(self._on_parsed)
        worker.failed.connect(self._on_parse_failed)
        self.stack.setCurrentIndex(1)
        worker.start()

    def _on_parsed(self, blocks: list, images_dir: str):
        self.review.load_blocks(blocks)
        self.statusBar().showMessage(
            f"解析完成：{len(blocks)} 个块；图片在 "
            f"{os.path.relpath(images_dir, HERE)}")
        # AI 模式：进入审阅后自动跑首轮结构修正（默认模式留给用户手动触发）
        if self._cur.get("mode") == "ai":
            self.review.run_ai_structure()

    def _on_parse_failed(self, msg: str):
        QMessageBox.critical(self, "解析失败", msg)
        self.stack.setCurrentIndex(0)
        self.statusBar().showMessage("解析失败，请检查文档后重试")

    # ---------------- 阶段二：确认导出 ----------------
    def _confirm(self, do_build: bool):
        md_target = self._cur["md_target"]
        try:
            os.makedirs(os.path.dirname(md_target), exist_ok=True)
            with open(md_target, "w", encoding="utf-8") as fh:
                fh.write(render(self.review.blocks))
        except OSError as e:
            QMessageBox.critical(self, "导出失败", f"无法写入 {md_target}：{e}")
            return

        rel_md = os.path.relpath(md_target, HERE)
        if not do_build:
            QMessageBox.information(
                self, "已导出",
                f"Markdown 已写入：{rel_md}\n可自行运行 python3 main.py run 出稿")
            self.statusBar().showMessage(f"已导出 {rel_md}")
            return

        self.statusBar().showMessage(f"已导出 {rel_md}，开始 run 构建...")
        self.build_page.reset()
        self.stack.setCurrentIndex(2)
        template = self._cur.get("template") or None
        self._runner.start(template)
