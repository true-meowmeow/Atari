# Module: application entrypoint.
# Main: main().
# Example: python main.py

from PySide6.QtCore import Qt
from PySide6.QtGui import QColor, QPalette
from PySide6.QtWidgets import QApplication

from atari.ui.main_window import MainWindow


def _apply_startup_dark_palette(app: QApplication) -> None:
    """Apply a dark fallback palette before the first window is shown."""
    app.setStyle("Fusion")
    palette = app.palette()
    palette.setColor(QPalette.Window, QColor("#0f1115"))
    palette.setColor(QPalette.WindowText, QColor("#e9eef7"))
    palette.setColor(QPalette.Base, QColor("#141821"))
    palette.setColor(QPalette.AlternateBase, QColor("#11161f"))
    palette.setColor(QPalette.Text, QColor("#e9eef7"))
    palette.setColor(QPalette.Button, QColor("#1b2332"))
    palette.setColor(QPalette.ButtonText, QColor("#e9eef7"))
    palette.setColor(QPalette.ToolTipBase, QColor("#141821"))
    palette.setColor(QPalette.ToolTipText, QColor("#e9eef7"))
    palette.setColor(QPalette.Highlight, QColor("#2c3f60"))
    palette.setColor(QPalette.HighlightedText, QColor("#f4f8ff"))
    app.setPalette(palette)


def main(app_version: str = "1.0"):
    app = QApplication([])
    _apply_startup_dark_palette(app)

    w = MainWindow()
    version = str(app_version or "").strip()
    if version:
        w.setWindowTitle(f"Atari {version}")
    else:
        w.setWindowTitle("Atari")

    # Force Qt to create a native window (HWND) early.
    w.setAttribute(Qt.WA_NativeWindow, True)
    w.winId()  # HWND is available now.

    w.show()
    app.exec()


if __name__ == "__main__":
    main()
