"""AI 配置弹窗（菜单：设置 → AI 配置）。

provider 选择自动带出 base_url/默认 model；api_key 掩码输入；
保存到 config/ai.json；提供一次同步连通性测试（最坏等待配置的超时秒数）。
配置属于全局设置，不放进向导主流程。
"""
from PyQt6.QtCore import Qt
from PyQt6.QtWidgets import (QComboBox, QDialog, QDialogButtonBox, QFormLayout,
                             QLineEdit, QMessageBox, QPushButton, QSpinBox,
                             QVBoxLayout, QApplication)

from docxai import config as aiconf
from docxai.client import AIClient, AIError


class SettingsDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("AI 配置")
        self.setMinimumWidth(520)

        self.provider = QComboBox()
        for key, p in aiconf.PROVIDERS.items():
            self.provider.addItem(p["label"], key)
        self.base_url = QLineEdit()
        self.api_key = QLineEdit()
        self.api_key.setEchoMode(QLineEdit.EchoMode.Password)
        self.api_key.setPlaceholderText("sk-...（仅保存在本机 config/ai.json）")
        self.model = QLineEdit()
        self.model.setPlaceholderText("模型名或端点 ID")
        self.timeout = QSpinBox()
        self.timeout.setRange(5, 600)
        self.timeout.setSuffix(" 秒")

        form = QFormLayout()
        form.addRow("服务商", self.provider)
        form.addRow("接口地址", self.base_url)
        form.addRow("API Key", self.api_key)
        form.addRow("模型", self.model)
        form.addRow("超时", self.timeout)

        self.test_btn = QPushButton("测试连接")
        self.test_btn.clicked.connect(self._test)
        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Save
            | QDialogButtonBox.StandardButton.Cancel)
        buttons.accepted.connect(self._save)
        buttons.rejected.connect(self.reject)

        layout = QVBoxLayout(self)
        layout.addLayout(form)
        layout.addWidget(self.test_btn)
        layout.addWidget(buttons)

        self.provider.currentIndexChanged.connect(self._on_provider)
        self._load()

    def _load(self):
        """回填已有配置；无配置则用默认 provider 的预设。"""
        if aiconf.exists():
            try:
                cfg = aiconf.load()
            except aiconf.ConfigError:
                cfg = aiconf.default_for("deepseek")
        else:
            cfg = aiconf.default_for("deepseek")
        key = cfg.get("provider", "deepseek")
        idx = max(0, self.provider.findData(key))
        self.provider.setCurrentIndex(idx)
        self.base_url.setText(str(cfg.get("base_url", "")))
        self.api_key.setText(str(cfg.get("api_key", "")))
        self.model.setText(str(cfg.get("model", "")))
        self.timeout.setValue(int(cfg.get("timeout", aiconf.DEFAULT_TIMEOUT)))

    def _on_provider(self, _idx):
        """切换 provider 时带出预设，但不覆盖用户已填的 key。"""
        preset = aiconf.PROVIDERS[self.provider.currentData()]
        self.base_url.setText(preset["base_url"])
        if preset["model"]:
            self.model.setText(preset["model"])

    def _collect(self) -> dict:
        return {
            "provider": self.provider.currentData(),
            "base_url": self.base_url.text().strip(),
            "api_key": self.api_key.text().strip(),
            "model": self.model.text().strip(),
            "timeout": self.timeout.value(),
        }

    def _save(self):
        cfg = self._collect()
        problems = aiconf.validate(cfg)
        if problems:
            QMessageBox.warning(self, "配置不完整", "；".join(problems))
            return
        try:
            aiconf.save(cfg)
        except aiconf.ConfigError as e:
            QMessageBox.critical(self, "保存失败", str(e))
            return
        self.accept()

    def _test(self):
        """同步发一句最短对话验证配置；key 不随错误信息展示。"""
        cfg = self._collect()
        if aiconf.validate(cfg):
            QMessageBox.warning(self, "配置不完整",
                                "请先填完接口地址、API Key、模型再测试")
            return
        self.test_btn.setEnabled(False)
        self.test_btn.setText("测试中，请稍候...")
        QApplication.setOverrideCursor(Qt.CursorShape.WaitCursor)
        try:
            client = AIClient(cfg)
            client.chat([{"role": "user", "content": "ping"}], temperature=0)
            QMessageBox.information(self, "连接成功", "API 调用正常，可以使用 AI 模式")
        except AIError as e:
            QMessageBox.warning(self, "连接失败", str(e))
        finally:
            QApplication.restoreOverrideCursor()
            self.test_btn.setEnabled(True)
            self.test_btn.setText("测试连接")
