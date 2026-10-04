"""gui 子命令入口：延迟导入 PyQt6，缺失时给出安装指引。"""
import sys


def main(argv=None) -> int:
    try:
        from PyQt6 import QtWidgets  # noqa: F401
    except ImportError:
        print("GUI 模式需要 PyQt6，当前环境未安装。")
        print("安装方法：pip install -r requirements.txt")
        print("（命令行的 run/build/convert 等功能不依赖 PyQt6，可照常使用）")
        return 1

    from docxgui.main_window import MainWindow
    from PyQt6 import QtWidgets
    app = QtWidgets.QApplication(sys.argv if argv is None else argv)
    app.setApplicationName("thesis-docx-tool")
    win = MainWindow()
    win.show()
    return app.exec()
