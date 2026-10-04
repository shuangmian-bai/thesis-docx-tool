"""第二页：审阅扁平化块序列。

左：按 H1 章分组的块树（纯文本，不渲染 Word 格式），可按类型筛选；
右：当前块编辑器（类型互转/标题级别/文本/代码/表格/图片路径与图题）。
结构修复手段：标题↔正文等类型误判可直接改类型，可插入正文/标题/有序列表项/
表格/图片、删除块，支持整篇 AI 结构修正与单块 AI 内容改写；插入/删除/改类型/
AI 整批可逐步撤销，文字改动可逐块还原。确认前只在内存中，不写任何 md。
"""
import os
from typing import Optional

from PyQt6.QtCore import Qt, pyqtSignal
from PyQt6.QtGui import QPixmap
from PyQt6.QtWidgets import (QComboBox, QFileDialog, QHBoxLayout, QInputDialog,
                             QLabel, QLineEdit, QListWidget, QListWidgetItem,
                             QMenu, QMessageBox, QPlainTextEdit, QProgressDialog,
                             QPushButton, QSpinBox, QSplitter, QTableWidget,
                             QTableWidgetItem, QToolButton, QVBoxLayout, QWidget)

from docxai import config as aiconf
from docxai.blocks_json import split_chapters
from docxai.client import AIClient
from docxgui import block_ops as ops
from docxgui import blocks_model as bm
from docxgui.workers import RewriteWorker, StructureWorker

#: 树项里保存全局块下标的 role
_IDX = Qt.ItemDataRole.UserRole


class ReviewPage(QWidget):
    """blocks 的审阅/编辑。所有 AI 任务在 worker 线程跑，界面不阻塞。"""

    back_requested = pyqtSignal()
    confirm_requested = pyqtSignal(bool)   # True=导出后继续 run；False=只导出 md
    settings_requested = pyqtSignal()
    status = pyqtSignal(str)

    def __init__(self, config_dir: str, images_dir: Optional[str] = None):
        super().__init__()
        self._config_dir = config_dir       # 图片路径基准（config/）
        self._images_dir = images_dir       # 本文档图片抽取目录（插入图片落点）
        self.blocks: list = []
        self._originals: dict = {}          # idx -> 首次改动前的块（还原依据）
        self._ai_changed: set = set()
        self._cur: Optional[int] = None
        self._loading = False
        self._struct_worker = None
        self._rewrite_worker = None
        self._history: list = []            # 结构操作栈（插入/删除/改类型/AI 整批）
        self._build()

    # ---------------- 界面搭建 ----------------
    def _build(self):
        layout = QVBoxLayout(self)

        bar = QHBoxLayout()
        bar.addWidget(QLabel("类型筛选："))
        self.filter = QComboBox()
        self.filter.addItem("全部类型", "")
        for k in bm.FILTER_KINDS:
            self.filter.addItem(bm.KIND_LABELS[k], k)
        self.filter.currentIndexChanged.connect(self._apply_filter)
        bar.addWidget(self.filter)
        self.btn_struct = QPushButton("AI 修正整篇结构")
        self.btn_struct.clicked.connect(self._ai_structure)
        self.btn_rewrite = QPushButton("AI 改写当前块")
        self.btn_rewrite.clicked.connect(self._ai_rewrite)
        self.btn_restore = QPushButton("还原当前块")
        self.btn_restore.clicked.connect(self._restore_current)
        bar.addWidget(self.btn_struct)
        bar.addWidget(self.btn_rewrite)
        bar.addWidget(self.btn_restore)

        self.btn_insert = QToolButton()
        self.btn_insert.setText("插入块 ▾")
        self.btn_insert.setPopupMode(
            QToolButton.ToolButtonPopupMode.InstantPopup)
        menu = QMenu(self)
        for kind, text in (("p", "正文段落"), ("h", "标题"),
                           ("ol", "有序列表项"), ("table", "表格"),
                           ("img", "图片（选择文件）")):
            menu.addAction(text, lambda k=kind: self._insert_block(k))
        self.btn_insert.setMenu(menu)
        bar.addWidget(self.btn_insert)
        self.btn_delete = QPushButton("删除当前块")
        self.btn_delete.clicked.connect(self._delete_current)
        bar.addWidget(self.btn_delete)
        self.btn_undo = QPushButton("撤销结构操作")
        self.btn_undo.clicked.connect(self._undo)
        bar.addWidget(self.btn_undo)
        bar.addStretch(1)
        layout.addLayout(bar)

        splitter = QSplitter(Qt.Orientation.Horizontal)
        self.tree = QListWidget()
        self.tree.currentRowChanged.connect(self._on_select)
        splitter.addWidget(self.tree)
        splitter.addWidget(self._build_editor())
        splitter.setSizes([460, 540])
        layout.addWidget(splitter, 1)

        bottom = QHBoxLayout()
        btn_back = QPushButton("上一步")
        btn_back.clicked.connect(self.back_requested.emit)
        bottom.addWidget(btn_back)
        bottom.addStretch(1)
        self.hint = QLabel("纯文本审阅；改类型/插入/删除均可撤销，确认结构无误后再导出")
        bottom.addWidget(self.hint)
        btn_md = QPushButton("仅导出 Markdown")
        btn_md.clicked.connect(lambda: self.confirm_requested.emit(False))
        btn_run = QPushButton("确认并生成 Word（run）")
        btn_run.setMinimumWidth(200)
        btn_run.clicked.connect(lambda: self.confirm_requested.emit(True))
        bottom.addWidget(btn_md)
        bottom.addWidget(btn_run)
        layout.addLayout(bottom)

    def _build_editor(self) -> QWidget:
        box = QVBoxLayout()
        self.kind_label = QLabel("未选择块")
        box.addWidget(self.kind_label)

        type_row = QHBoxLayout()
        type_row.addWidget(QLabel("修改类型："))
        self.kind_combo = QComboBox()
        for key in ops.CONVERT_KEYS:
            self.kind_combo.addItem(bm.KIND_LABELS[key], key)
        self.kind_combo.setEnabled(False)
        self.kind_combo.activated.connect(self._convert_current)
        type_row.addWidget(self.kind_combo)
        type_row.addWidget(QLabel("（标题↔正文等误判可在此修正）"))
        type_row.addStretch(1)
        box.addLayout(type_row)

        lv_row = QHBoxLayout()
        lv_row.addWidget(QLabel("标题级别："))
        self.level_spin = QSpinBox()
        self.level_spin.setRange(1, 4)
        self.level_spin.setEnabled(False)
        self.level_spin.valueChanged.connect(self._write_level)
        lv_row.addWidget(self.level_spin)
        lv_row.addStretch(1)
        box.addLayout(lv_row)

        self.text_edit = QPlainTextEdit()
        self.text_edit.setVisible(False)
        self.text_edit.textChanged.connect(self._write_text)
        box.addWidget(self.text_edit, 1)

        self.table = QTableWidget()
        self.table.setVisible(False)
        self.table.cellChanged.connect(self._write_table)
        box.addWidget(self.table, 1)

        self._table_bar = QWidget()
        tbar = QHBoxLayout(self._table_bar)
        tbar.setContentsMargins(0, 0, 0, 0)
        for text, slot in (("加行", lambda: self._resize_table(1, 0)),
                           ("加列", lambda: self._resize_table(0, 1)),
                           ("删行", lambda: self._resize_table(-1, 0)),
                           ("删列", lambda: self._resize_table(0, -1))):
            btn = QPushButton(text)
            btn.clicked.connect(slot)
            tbar.addWidget(btn)
        tbar.addStretch(1)
        self._table_bar.setVisible(False)
        box.addWidget(self._table_bar)

        self.img_preview = QLabel("（图片预览）")
        self.img_preview.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.img_path = QLineEdit()
        self.img_path.setReadOnly(True)
        self.img_caption = QLineEdit()
        self.img_caption.textChanged.connect(self._write_caption)
        cap_row = QHBoxLayout()
        cap_row.addWidget(QLabel("图题："))
        cap_row.addWidget(self.img_caption, 1)
        box.addWidget(self.img_preview, 1)
        box.addWidget(QLabel("图片路径（只读，替换文件请到素材目录）："))
        box.addWidget(self.img_path)
        box.addLayout(cap_row)
        self._set_image_visible(False)

        host = QWidget()
        host.setLayout(box)
        return host

    def _set_image_visible(self, visible: bool):
        for w in (self.img_preview, self.img_path, self.img_caption):
            w.setVisible(visible)

    # ---------------- 数据装载 ----------------
    def set_images_dir(self, images_dir: str):
        """主窗口在解析完成后告知本文档图片目录（插入图片的落点）。"""
        self._images_dir = images_dir

    def load_blocks(self, blocks: list):
        """初次装载：重置一切备份与 AI 标记。"""
        self.blocks = list(blocks)
        self._originals.clear()
        self._ai_changed.clear()
        self._history.clear()
        self._cur = None
        self._rebuild_tree()
        self.status.emit(f"共 {len(self.blocks)} 个块，请核对结构分类")

    def _rebuild_tree(self):
        """按 H1 章分组重灌列表；项文本带类型标签与 AI 标记。"""
        self.tree.blockSignals(True)
        self.tree.clear()
        for ch in split_chapters(self.blocks):
            chapter_item = QListWidgetItem(f"■ {ch['title']}")
            chapter_item.setFlags(Qt.ItemFlag.NoItemFlags)
            font = chapter_item.font()
            font.setBold(True)
            chapter_item.setFont(font)
            self.tree.addItem(chapter_item)
            for idx in ch["indexes"]:
                self.tree.addItem(self._make_item(idx))
        self.tree.blockSignals(False)
        self._apply_filter()

    def _make_item(self, idx: int) -> QListWidgetItem:
        blk = self.blocks[idx]
        mark = "  [AI 已改]" if idx in self._ai_changed else ""
        item = QListWidgetItem(f"    [{bm.label(blk)}] {bm.summary(blk)}{mark}")
        item.setData(_IDX, idx)
        return item

    def _apply_filter(self):
        """按类型筛选；章下无可见块时整章隐藏。"""
        want = self.filter.currentData()
        last_chapter = None
        for i in range(self.tree.count()):
            item = self.tree.item(i)
            idx = item.data(_IDX)
            if idx is None:                      # 章标题
                last_chapter = item
                item.setHidden(True)
                continue
            visible = (not want) or bm.kind_key(self.blocks[idx]) == want
            item.setHidden(not visible)
            if visible and last_chapter is not None:
                last_chapter.setHidden(False)

    def _refresh_item(self, idx: int):
        for i in range(self.tree.count()):
            item = self.tree.item(i)
            if item.data(_IDX) == idx:
                old = self.tree.currentRow()
                self.tree.blockSignals(True)
                self.tree.takeItem(i)
                self.tree.insertItem(i, self._make_item(idx))
                self.tree.blockSignals(False)
                self.tree.setCurrentRow(old)
                return

    # ---------------- 块编辑 ----------------
    def _on_select(self, _row: int):
        item = self.tree.currentItem()
        idx = item.data(_IDX) if item else None
        # 离开旧块前再刷新它在树里的摘要（编辑过程中不刷，避免输入丢焦点）
        if self._cur is not None and self._cur != idx:
            self._refresh_item(self._cur)
        self._cur = idx
        self._loading = True
        try:
            if idx is None:
                self.kind_label.setText("章标题" if item else "未选择块")
                self.kind_combo.setEnabled(False)
                self._show_editor("none")
                return
            blk = self.blocks[idx]
            self.kind_label.setText(f"块 #{idx} · 类型：{bm.label(blk)}")
            self._fill_kind_combo(blk)
            if blk[0] == "h":
                self.level_spin.setValue(blk[1])
                self.text_edit.setPlainText(blk[2])
                self._show_editor("h")
            elif blk[0] in ("p", "ol", "quote", "caption", "ref"):
                self.text_edit.setPlainText(blk[1])
                self._show_editor("text")
            elif blk[0] == "code":
                self.text_edit.setPlainText("\n".join(blk[1]))
                self._show_editor("text")
                self.kind_label.setText(f"块 #{idx} · 类型：代码（每行一条）")
            elif blk[0] == "table":
                self._fill_table(blk[1])
                self._show_editor("table")
            else:  # img
                self.img_path.setText(blk[1])
                self.img_caption.setText(blk[2])
                self._fill_image(blk[1])
                self._show_editor("image")
        finally:
            self._loading = False

    def _show_editor(self, mode: str):
        """切换右侧编辑器：none/h/text/table/image。"""
        self.level_spin.setEnabled(mode == "h")
        self.text_edit.setVisible(mode in ("h", "text"))
        self.table.setVisible(mode == "table")
        self._table_bar.setVisible(mode == "table")
        self._set_image_visible(mode == "image")

    def _fill_table(self, rows: list):
        self.table.blockSignals(True)
        ncol = max(len(r) for r in rows)
        self.table.setRowCount(len(rows))
        self.table.setColumnCount(ncol)
        for i, r in enumerate(rows):
            for j in range(ncol):
                self.table.setItem(i, j,
                                   QTableWidgetItem(r[j] if j < len(r) else ""))
        self.table.blockSignals(False)

    def _fill_image(self, rel_path: str):
        full = os.path.join(self._config_dir, rel_path)
        pix = QPixmap(full)
        if pix.isNull():
            self.img_preview.setText(f"图片加载失败：{rel_path}")
            self.img_preview.setPixmap(QPixmap())
        else:
            self.img_preview.setText("")
            self.img_preview.setPixmap(pix.scaled(
                360, 360, Qt.AspectRatioMode.KeepAspectRatio,
                Qt.TransformationMode.SmoothTransformation))

    def _mutate(self, idx: int, new_block: tuple):
        """统一改动入口：首次改动前备份原块，再写回活数据。

        树摘要不在每次输入时刷新（由切换选择时统一刷新），避免编辑丢焦点。
        """
        if idx not in self._originals:
            self._originals[idx] = self.blocks[idx]
        self.blocks[idx] = new_block

    def _write_text(self):
        if self._loading or self._cur is None:
            return
        blk = self.blocks[self._cur]
        val = self.text_edit.toPlainText()
        if blk[0] == "h":
            self._mutate(self._cur, ("h", blk[1], val))
        elif blk[0] == "code":
            self._mutate(self._cur, ("code", val.split("\n")))
        else:
            self._mutate(self._cur, (blk[0], val))

    def _write_level(self, lv: int):
        if self._loading or self._cur is None:
            return
        blk = self.blocks[self._cur]
        if blk[0] == "h":
            self._mutate(self._cur, ("h", lv, blk[2]))
            self._refresh_item(self._cur)

    def _write_table(self, _r, _c):
        if self._loading or self._cur is None:
            return
        rows = [[self.table.item(i, j).text() if self.table.item(i, j) else ""
                 for j in range(self.table.columnCount())]
                for i in range(self.table.rowCount())]
        self._mutate(self._cur, ("table", rows))
        self._refresh_item(self._cur)

    def _write_caption(self, text: str):
        if self._loading or self._cur is None:
            return
        blk = self.blocks[self._cur]
        if blk[0] == "img":
            self._mutate(self._cur, ("img", blk[1], text))
            self._refresh_item(self._cur)

    def _resize_table(self, dr: int, dc: int):
        self._loading = True
        rows = max(1, self.table.rowCount() + dr)
        cols = max(1, self.table.columnCount() + dc)
        self.table.setRowCount(rows)
        self.table.setColumnCount(cols)
        self._loading = False
        self._write_table(0, 0)

    def _restore_current(self):
        if self._cur is None or self._cur not in self._originals:
            self.status.emit("该块未被修改，无需还原")
            return
        idx = self._cur
        self.blocks[idx] = self._originals.pop(idx)
        self._ai_changed.discard(idx)
        row = self.tree.currentRow()
        self._refresh_item(idx)
        self.tree.setCurrentRow(row)
        self._on_select(row)
        self.status.emit(f"块 #{idx} 已还原为原始内容")

    # ---------------- 结构操作：改类型 / 插入 / 删除 / 撤销 ----------------
    def _fill_kind_combo(self, blk: tuple):
        """按当前块设置类型下拉：可互转类型启用并选中当前类型，表格/图片禁用。"""
        ok = ops.can_convert(blk)
        self.kind_combo.setEnabled(ok)
        if not ok:
            return
        row = self.kind_combo.findData(bm.kind_key(blk))
        self.kind_combo.blockSignals(True)
        if row >= 0:
            self.kind_combo.setCurrentIndex(row)
        self.kind_combo.blockSignals(False)

    def _convert_current(self, _row: int = -1):
        if self._cur is None:
            return
        idx = self._cur
        key = self.kind_combo.currentData()
        old = self.blocks[idx]
        if not key or key == bm.kind_key(old):
            return
        new = ops.convert(old, key)
        self._history.append(("conv", idx, old, new))
        self._mutate(idx, new)
        self._rebuild_tree()
        self._select_idx(idx)
        self.status.emit(f"块 #{idx} 已改为「{bm.KIND_LABELS[key]}」，"
                         f"可点「撤销结构操作」或「还原当前块」")

    def _select_idx(self, idx: int):
        """按全局块下标选中树项（章标题项无 _IDX）。"""
        for i in range(self.tree.count()):
            if self.tree.item(i).data(_IDX) == idx:
                self.tree.setCurrentRow(i)
                return

    def _shift_marks_up(self, pos: int):
        """在 pos 处插入一块后：备份表与 AI 标记中 >=pos 的下标整体后移。"""
        for k in sorted((k for k in self._originals if k >= pos), reverse=True):
            self._originals[k + 1] = self._originals.pop(k)
        self._ai_changed = {k + 1 if k >= pos else k for k in self._ai_changed}

    def _shift_marks_down(self, idx: int):
        """删除 idx 处块后：备份表与 AI 标记中 >idx 的下标整体前移。"""
        self._originals.pop(idx, None)
        for k in sorted(k for k in self._originals if k > idx):
            self._originals[k - 1] = self._originals.pop(k)
        self._ai_changed = {k - 1 if k > idx else k
                            for k in self._ai_changed if k != idx}

    def _insert_block(self, kind: str):
        if kind == "img":
            self._insert_image()
            return
        self._insert_at(ops.new_block(kind), bm.KIND_LABELS[
            kind if kind != "h" else "h2"])

    def _insert_at(self, blk: tuple, label: str):
        pos = self._cur + 1 if self._cur is not None else len(self.blocks)
        self.blocks.insert(pos, blk)
        self._shift_marks_up(pos)
        self._history.append(("ins", pos, blk))
        self._cur = None
        self._rebuild_tree()
        self._select_idx(pos)
        if blk[0] in ("p", "h", "ol"):
            self.text_edit.setFocus()
        self.status.emit(f"已在 #{pos} 插入「{label}」，请编辑内容（可撤销）")

    def _insert_image(self):
        if not self._images_dir:
            QMessageBox.warning(self, "无法插入图片",
                                "请先解析一个 Word 文档（图片需复制到本文档的素材目录）")
            return
        src, _ = QFileDialog.getOpenFileName(
            self, "选择要插入的图片", "",
            "图片 (*.png *.jpg *.jpeg *.gif *.bmp)")
        if not src:
            return
        rel = ops.copy_image(src, self._images_dir, self._config_dir)
        self._insert_at(ops.new_image_block(rel), "图片")

    def _delete_current(self):
        if self._cur is None:
            QMessageBox.information(self, "提示", "请先在左侧选择一个块（章标题不能删除）")
            return
        idx = self._cur
        if QMessageBox.question(
                self, "确认删除",
                f"确定删除块 #{idx}（{bm.label(self.blocks[idx])}）吗？"
                "删除后可用「撤销结构操作」恢复。") \
                != QMessageBox.StandardButton.Yes:
            return
        blk = self.blocks.pop(idx)
        self._shift_marks_down(idx)
        self._history.append(("del", idx, blk))
        self._cur = None
        self._rebuild_tree()
        if self.blocks:
            self._select_idx(min(idx, len(self.blocks) - 1))
        self.status.emit(f"已删除块 #{idx}（可撤销）")

    def _undo(self):
        if not self._history:
            self.status.emit("没有可撤销的结构操作")
            return
        op = self._history.pop()
        select = None
        if op[0] == "ins":
            pos = op[1]
            self.blocks.pop(pos)
            self._originals.pop(pos, None)
            for k in sorted(k for k in self._originals if k > pos):
                self._originals[k - 1] = self._originals.pop(k)
            self._ai_changed = {k - 1 if k > pos else k
                                for k in self._ai_changed if k != pos}
            select = min(pos, len(self.blocks) - 1)
        elif op[0] == "del":
            pos, blk = op[1], op[2]
            self.blocks.insert(pos, blk)
            self._shift_marks_up(pos)
            select = pos
        elif op[0] == "conv":
            idx, old = op[1], op[2]
            self.blocks[idx] = old
            select = idx
        elif op[0] == "ai":
            self.blocks = list(op[1])
            self._ai_changed.clear()
        self._cur = None
        self._rebuild_tree()
        if select is not None and 0 <= select < len(self.blocks):
            self._select_idx(select)
        left = len(self._history)
        self.status.emit(f"已撤销上一步结构操作（历史剩 {left} 步）")

    # ---------------- AI 操作 ----------------
    def run_ai_structure(self):
        """公开入口：供主窗口在「AI 模式」解析完成后自动触发首轮结构修正。"""
        self._ai_structure()

    def _client_or_open_settings(self) -> Optional[AIClient]:
        if not aiconf.configured():
            QMessageBox.information(self, "需要 AI 配置",
                                    "请先在「设置 → AI 配置」中填写服务商与 API Key")
            self.settings_requested.emit()
            return None
        try:
            return AIClient(aiconf.load())
        except Exception as e:
            QMessageBox.warning(self, "AI 配置无效", str(e))
            self.settings_requested.emit()
            return None

    def _ai_structure(self):
        client = self._client_or_open_settings()
        if client is None:
            return
        known = {b[1] for b in self.blocks if b[0] == "img"}
        prog = QProgressDialog("AI 正在按章修正结构...", "取消", 0, 100, self)
        prog.setWindowTitle("AI 结构修正")
        prog.setWindowModality(Qt.WindowModality.WindowModal)
        prog.setMinimumDuration(0)

        worker = StructureWorker(client, self.blocks, known)
        self._struct_worker = worker
        worker.progress.connect(
            lambda d, t, name: (
                prog.setValue(int(d * 100 / max(t, 1))),
                prog.setLabelText(f"正在处理（{d}/{t}）：{name}")))
        prog.canceled.connect(worker.cancel)

        def done(result):
            prog.close()
            # 修正前把将被替换的块备份（已有更早的手动改动备份则保留更早版本）
            for idx in result.changed:
                self._originals.setdefault(idx, self.blocks[idx])
            self._history.append(("ai", list(self.blocks)))
            self.blocks = result.blocks
            self._ai_changed = set(result.changed)
            self._cur = None
            self._rebuild_tree()
            msg = (f"AI 结构修正完成：成功 {result.ok_chapters}/"
                   f"{result.total_chapters} 章，修改 {len(result.changed)} 块")
            if result.skipped:
                msg += ("\n\n以下章节校验未通过，已整批保留原文：\n"
                        + "\n".join(f"· {s['chapter']}：{s['reason'][:60]}"
                                    for s in result.skipped))
            QMessageBox.information(self, "结构修正结果", msg)
            self.status.emit("AI 结构修正完成，标记 [AI 已改] 的块可逐块还原")

        def fail(msg):
            prog.close()
            QMessageBox.warning(self, "AI 结构修正失败", msg)

        worker.finished_result.connect(done)
        worker.failed.connect(fail)
        worker.start()
        prog.exec()

    def _ai_rewrite(self):
        if self._cur is None:
            QMessageBox.information(self, "提示", "请先在左侧选择一个文本块")
            return
        blk = self.blocks[self._cur]
        if blk[0] not in bm.TEXT_EDITABLE:
            QMessageBox.information(self, "不支持",
                                    "只支持标题、正文、引用、图题表题、文献块的内容改写")
            return
        client = self._client_or_open_settings()
        if client is None:
            return
        instruction, ok = QInputDialog.getText(
            self, "AI 改写", "修改要求（如：润色为更正式的学术语气、精简到 100 字）：")
        if not ok or not instruction.strip():
            return
        text = blk[2] if blk[0] == "h" else blk[1]
        prog = QProgressDialog("AI 正在改写...", "取消", 0, 0, self)
        prog.setWindowModality(Qt.WindowModality.WindowModal)
        worker = RewriteWorker(client, text, instruction.strip())
        self._rewrite_worker = worker
        prog.canceled.connect(worker.cancel)
        idx = self._cur

        def done(new_text):
            prog.close()
            if blk[0] == "h":
                self._mutate(idx, ("h", blk[1], new_text))
            else:
                self._mutate(idx, (blk[0], new_text))
            row = self.tree.currentRow()
            self._on_select(row)
            self.status.emit(f"块 #{idx} 已由 AI 改写，可点「还原当前块」撤销")

        def fail(msg):
            prog.close()
            QMessageBox.warning(self, "AI 改写失败", msg)

        worker.finished_text.connect(done)
        worker.failed.connect(fail)
        worker.start()
        prog.exec()
