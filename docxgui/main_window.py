"""GUI 主窗口：向导 → 审阅 → 构建 三页编排与全局设置入口。

主窗口只做编排与页面间数据传递：
- 解析/AI/构建均由 workers 后台执行，本类不写业务逻辑；
- 审阅确认后把块按 H1 拆成数字开头的分章 md 写入 config/章节/（清理旧分章），
  再由外部进程 `python3 main.py run` 完成全部后续阶段。
"""
import os

from PyQt6.QtGui import QAction
from PyQt6.QtWidgets import (QMainWindow, QMessageBox, QStackedWidget,
                             QStatusBar)

import re

from docxai import HERE
from docxai.blocks_json import split_chapters
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
        blocks = self.review.blocks
        # 按 H1 拆成多章，分别写成数字开头的分章 md，写入前清理旧分章文件，
        # 确保 build 只用本次审阅修正后的内容，不会与历史章节混在一起。
        chapters = split_chapters(blocks)
        try:
            os.makedirs(CHAPTERS_DIR, exist_ok=True)
            # 清理旧的数字开头章节文件（非数字开头的如模板示例不动，由 build 自行跳过）
            for fn in os.listdir(CHAPTERS_DIR):
                if fn.endswith(".md") and fn[:1].isdigit():
                    os.remove(os.path.join(CHAPTERS_DIR, fn))
            written = []
            for i, ch in enumerate(chapters, 1):
                ch_blocks = [blocks[idx] for idx in ch["indexes"]]
                title = self._safe_filename(ch["title"])
                fn = f"{i:02d}_{title}.md"
                with open(os.path.join(CHAPTERS_DIR, fn), "w", encoding="utf-8") as fh:
                    fh.write(render(ch_blocks))
                written.append(fn)
        except OSError as e:
            QMessageBox.critical(self, "导出失败", f"无法写入章节文件：{e}")
            return

        rel_dir = os.path.relpath(CHAPTERS_DIR, HERE)
        summary = f"{len(written)} 个章节 → {rel_dir}"
        if not do_build:
            QMessageBox.information(
                self, "已导出",
                f"已写入 {summary}\n可自行运行 python3 main.py run 出稿")
            self.statusBar().showMessage(f"已导出 {summary}")
            return

        self.statusBar().showMessage(f"已导出 {summary}，开始 run 构建...")
        self.build_page.reset()
        self.stack.setCurrentIndex(2)
        template = self._cur.get("template") or None
        self._runner.start(template)

    @staticmethod
    def _safe_filename(title: str) -> str:
        """把章节标题转成合法文件名：去掉路径分隔符等非法字符，空标题回退为“章节”。"""
        name = re.sub(r'[\\/:*?"<>|\r\n]+', "_", title).strip()
        return name or "章节"
