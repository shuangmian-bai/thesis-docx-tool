"""模板预览对话框：把模板骨架从黑盒变成可检视的中间结果。

- 预览：tplinspect 的只读检视结果（骨架纯文本序列、封面字段、
  TOC 域、样式、编号定义、页面设置），问题分级为错误/警告；
- 封面与承诺书（含签名）保留模板原样，不提供修复入口；
  骨架/样式本体不在界面里改 OOXML——在 Word 中修好模板文件后点「重新检测」复验，
  全程人工兜底、过程完全可控（见 README「设计理念」）。
"""
import os

from PyQt6.QtWidgets import (QDialog, QHBoxLayout, QLabel, QLineEdit,
                             QPlainTextEdit, QPushButton, QVBoxLayout,
                             QFileDialog)

from docxbuild import tplinspect


class TemplateDialog(QDialog):
    """只读检视模板。template_path 可为空串（由调用方给默认值后再传入）。"""

    def __init__(self, template_path: str, tool_root: str, parent=None):
        super().__init__(parent)
        self.setWindowTitle("模板预览")
        self.resize(860, 720)
        self._root = tool_root
        self._path = template_path
        self._build()
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
        root.addWidget(self.report, 1)

        tip = QLabel(
            "骨架部件或样式报「错误/警告」时，请在 Word 中修改该模板文件后"
            "点「重新检测」复验；本工具不直接改写模板内部 XML，"
            "封面与承诺书（含签名）保留模板原样。")
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
            L.append("检测结果：通过。构建所需的骨架、样式、编号齐备。")
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
        L.append(f"  TOC 目录域：{'有' if rep.get('has_toc') else '无'}")
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
                 f"媒体 {len(rep['parts'].get('media', []))} 个")
        return "\n".join(L)

    # ---------------- 动作 ----------------
    def _pick_template(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "选择学校模板", "", "Word 文档 (*.docx)")
        if path:
            self._path = os.path.abspath(path)
            self.path_edit.setText(self._path)
            self.refresh()
