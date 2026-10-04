"""模板预览与修复对话框：把模板骨架从黑盒变成可检视、可修复的中间结果。

- 预览：tplinspect 的只读检视结果（骨架纯文本序列、封面字段、承诺书题目占位符、
  TOC 域、样式、编号定义、页面设置），问题分级为错误/警告；
- 修复：封面字段写入 config/cover.json、承诺书签名图替换 config/images/signature.png；
  骨架/样式本体不在界面里改 OOXML——在 Word 中修好模板文件后点「重新检测」复验，
  全程人工兜底、过程完全可控（见 README「设计理念」）。
"""
import os
import shutil

from PyQt6.QtCore import pyqtSignal
from PyQt6.QtWidgets import (QDialog, QFileDialog, QFormLayout,
                             QGroupBox, QHBoxLayout, QLabel, QLineEdit,
                             QPlainTextEdit, QPushButton, QScrollArea,
                             QVBoxLayout, QWidget)

from docxbuild import tplinspect
from docxbuild.docinfo import COVER


class TemplateDialog(QDialog):
    """检视与修复模板。template_path 可为空串（由调用方给默认值后再传入）。"""

    template_changed = pyqtSignal(str)

    def __init__(self, template_path: str, tool_root: str, work_dir: str = None,
                 parent=None):
        super().__init__(parent)
        self.setWindowTitle("模板预览与修复")
        self.resize(860, 720)
        self._root = tool_root
        self._path = template_path
        self._work_dir = work_dir  # 当前论文的哈希工作目录（None=全局设置）
        self._cover_edits = {}
        self._build()
        self._load_cover_form()
        self.refresh()

    # ---------------- 界面 ----------------
    def _build(self):
        root = QVBoxLayout(self)

        row = QHBoxLayout()
        row.addWidget(QLabel("模板文件："))
        self.path_edit = QLineEdit(self._path)
        self.path_edit.setReadOnly(True)
        row.addWidget(self.path_edit, 1)
        btn_pick = QPushButton("更换模板...")
        btn_pick.clicked.connect(self._pick_template)
        row.addWidget(btn_pick)
        btn_check = QPushButton("重新检测")
        btn_check.clicked.connect(self.refresh)
        row.addWidget(btn_check)
        root.addLayout(row)

        # 检视报告（只读纯文本）
        self.report = QPlainTextEdit()
        self.report.setReadOnly(True)
        root.addWidget(self.report, 3)

        # 封面字段修复表单
        save_loc = "当前论文工作目录" if self._work_dir else "config/cover.json"
        box = QGroupBox(f"修复一：封面与承诺书字段（保存到 {save_loc}，"
                        "构建时自动填入；不修改模板本体）")
        form_host = QWidget()
        form = QFormLayout(form_host)
        self.title_edit = QLineEdit()
        self.version_edit = QLineEdit()
        self.tpl_title_edit = QLineEdit()
        form.addRow("论文题目 title：", self.title_edit)
        form.addRow("版本号 version：", self.version_edit)
        form.addRow("承诺书模板题目 template_title：", self.tpl_title_edit)
        for k in COVER:
            edit = QLineEdit()
            self._cover_edits[k] = edit
            form.addRow(f"封面「{k}」：", edit)
        scroll = QScrollArea()
        scroll.setWidget(form_host)
        scroll.setWidgetResizable(True)
        bl = QVBoxLayout(box)
        bl.addWidget(scroll)
        btn_save = QPushButton(f"保存封面字段到 {save_loc}")
        btn_save.clicked.connect(self._save_cover)
        bl.addWidget(btn_save)
        root.addWidget(box, 2)

        # 签名图修复
        sbox = QGroupBox("修复二：诚信承诺书签名图（替换 config/images/signature.png，"
                         "构建时自动换入并按比例缩放）")
        srow = QHBoxLayout(sbox)
        self.sign_status = QLabel()
        srow.addWidget(self.sign_status, 1)
        btn_sign = QPushButton("选择图片并替换...")
        btn_sign.clicked.connect(self._replace_signature)
        srow.addWidget(btn_sign)
        root.addWidget(sbox)

        tip = QLabel(
            "修复三：骨架部件或样式报「错误/警告」时，请在 Word 中修改该模板文件后"
            "点「重新检测」复验；本工具不直接改写模板内部 XML，避免损坏封面/目录域。")
        tip.setWordWrap(True)
        root.addWidget(tip)

        btn_close = QPushButton("关闭")
        btn_close.clicked.connect(self.accept)
        root.addWidget(btn_close)

    # ---------------- 数据 ----------------
    def refresh(self):
        """重新检视模板并渲染报告。"""
        rep = tplinspect.inspect_template(self._path)
        self.report.setPlainText(self._render(rep))
        sp = tplinspect.signature_path(self._root)
        if os.path.exists(sp):
            self.sign_status.setText(f"已就绪：{os.path.relpath(sp, self._root)}")
        else:
            self.sign_status.setText(
                f"缺失：{os.path.relpath(sp, self._root)}（构建会中断，请替换）")

    def _render(self, rep: dict) -> str:
        if not rep["exists"]:
            return f"[错误] {rep['path']} 不存在，请点「更换模板」选择学校模板 docx。"
        L = []
        L.append(f"模板文件：{rep['path']}")
        L.append("=" * 70)
        if rep["problems"]:
            L.append("检测结果：发现问题（按严重程度排序）")
            for lv, msg in rep["problems"]:
                L.append(f"  [{'错误' if lv == 'error' else '警告'}] {msg}")
        else:
            L.append("检测结果：通过。构建所需的骨架、样式、编号、签名部件齐备。")
        L.append("")

        L.append(f"一、保留骨架（{rep.get('skeleton_count', 0)} 个部件，"
                 "构建时原样保留，正文从其后开始替换）：")
        for kind, txt in rep.get("skeleton", []):
            L.append(f"  [{kind}] {txt}")
        L.append("")

        cov = rep.get("cover", {})
        L.append("二、封面表格字段（模板中存在的行）：")
        for lab in cov.get("labels", []):
            L.append(f"  · {lab}")
        if cov.get("missing"):
            L.append("  缺少：" + "、".join(cov["missing"]))
        L.append(f"  承诺书题目占位符：{'有' if rep.get('commitment_placeholder') else '无'}"
                 f"    TOC 目录域：{'有' if rep.get('has_toc') else '无'}")
        L.append("")

        st = rep.get("styles", {})
        L.append("三、样式（构建端依赖）：")
        present = set(st.get("present", []))
        for sid in st.get("required", []) + st.get("optional", []):
            L.append(f"  {'有' if sid in present else '缺'}  styleId={sid}")
        L.append("")

        num = rep.get("numbering", {})
        L.append(f"四、编号定义：numbering.xml "
                 f"{'有' if rep['parts'].get('numbering') else '无'}"
                 f"；decimal 自动编号：{'有' if num.get('decimal') else '无'}"
                 f"（有序列表需要）；现有 num 实例 {num.get('num_count', 0)} 个")
        L.append("")

        pg = rep.get("page", {})
        if pg:
            m = pg.get("margin_cm", {})
            L.append(f"五、页面设置：{pg.get('width_cm')} × {pg.get('height_cm')} cm；"
                     f"页边距 上{m.get('top')} 下{m.get('bottom')} "
                     f"左{m.get('left')} 右{m.get('right')} cm")
        L.append(f"     页眉 {len(rep['parts'].get('header', []))} 个 · "
                 f"页脚 {len(rep['parts'].get('footer', []))} 个 · "
                 f"媒体 {len(rep['parts'].get('media', []))} 个 · "
                 f"签名部件 {'有' if rep['parts'].get('has_signature_part') else '缺'}")
        return "\n".join(L)

    def _load_cover_form(self):
        form = tplinspect.load_cover_form(self._root, self._work_dir)
        self.title_edit.setText(form["title"])
        self.version_edit.setText(form["version"])
        self.tpl_title_edit.setText(form["template_title"])
        for k, edit in self._cover_edits.items():
            edit.setText(form["cover"][k])

    def _collect_form(self) -> dict:
        return {"title": self.title_edit.text(),
                "version": self.version_edit.text(),
                "template_title": self.tpl_title_edit.text(),
                "cover": {k: e.text() for k, e in self._cover_edits.items()}}

    # ---------------- 修复动作 ----------------
    def _pick_template(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "选择学校模板", "", "Word 文档 (*.docx)")
        if path:
            self._path = os.path.abspath(path)
            self.path_edit.setText(self._path)
            self.template_changed.emit(self._path)
            self.refresh()

    def _save_cover(self):
        p = tplinspect.save_cover_form(self._root, self._collect_form(),
                                       self._work_dir)
        self.sign_status.setText(
            f"封面字段已保存：{os.path.relpath(p, self._root)}（重新构建即生效）")

    def _replace_signature(self):
        src, _ = QFileDialog.getOpenFileName(
            self, "选择签名图片", "", "图片 (*.png *.jpg *.jpeg)")
        if not src:
            return
        dst = tplinspect.signature_path(self._root)
        os.makedirs(os.path.dirname(dst), exist_ok=True)
        shutil.copyfile(src, dst)
        self.refresh()
