"""
Copyright 2026 pl1a

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the “Software”), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED “AS IS”, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
"""

import os
import sys
import io
from datetime import datetime
from contextlib import redirect_stdout

from PySide6.QtCore import Qt, QThread, Signal
from PySide6.QtWidgets import (
    QApplication,
    QMainWindow,
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QFormLayout,
    QLabel,
    QLineEdit,
    QPushButton,
    QFileDialog,
    QTextEdit,
    QMessageBox,
    QGroupBox,
    QSizePolicy,
)

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
if SCRIPT_DIR not in sys.path:
    sys.path.insert(0, SCRIPT_DIR)

from exporter import parse_chat_logs

DEFAULT_MAP_URL = "https://map.example.net/nodes/towns.json"

class Worker(QThread):
    log = Signal(str)
    finished_ok = Signal(str)
    failed = Signal(str)

    def __init__(self, log_path: str, out_path: str, towns_url: str):
        super().__init__()
        self.log_path = log_path
        self.out_path = out_path
        self.towns_url = towns_url

    def run(self):
        try:
            buf = io.StringIO()
            with redirect_stdout(buf):
                parse_chat_logs(
                    self.log_path,
                    self.out_path,
                    towns_url=self.towns_url,
                )
            output = buf.getvalue().strip()
            if output:
                for line in output.splitlines():
                    self.log.emit(line)
            self.finished_ok.emit(self.out_path)
        except Exception as e:
            self.failed.emit(str(e))


class StatsWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Nodes War Stats Exporter")
        self.setMinimumSize(640, 480)
        self.resize(760, 560)

        self.worker = None
        self._build_ui()

    def _build_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        root = QVBoxLayout(central)
        root.setContentsMargins(14, 14, 14, 14)
        root.setSpacing(12)

        form_box = QGroupBox("Inputs")
        form = QFormLayout(form_box)
        form.setSpacing(8)
        form.setLabelAlignment(Qt.AlignRight)

        log_row = QHBoxLayout()
        self.log_edit = QLineEdit()
        self.log_edit.setPlaceholderText("Path to Minecraft chat log (e.g. latest.log)")
        log_browse = QPushButton("Browse…")
        log_browse.clicked.connect(self._browse_log)
        log_row.addWidget(self.log_edit)
        log_row.addWidget(log_browse)
        form.addRow("Chat log file:", log_row)

        url_row = QHBoxLayout()
        self.url_edit = QLineEdit(DEFAULT_MAP_URL)
        url_reset = QPushButton("Reset")
        url_reset.clicked.connect(lambda: self.url_edit.setText(DEFAULT_MAP_URL))
        url_row.addWidget(self.url_edit)
        url_row.addWidget(url_reset)
        form.addRow("Towns JSON URL:", url_row)

        out_row = QHBoxLayout()
        default_out = f"war-{datetime.now().strftime('%m-%d-%Y')}.xlsx"
        self.out_edit = QLineEdit(default_out)
        out_browse = QPushButton("Browse…")
        out_browse.clicked.connect(self._browse_out)
        out_row.addWidget(self.out_edit)
        out_row.addWidget(out_browse)
        form.addRow("Output Excel:", out_row)

        root.addWidget(form_box)

        btn_row = QHBoxLayout()
        btn_row.addStretch()
        self.run_btn = QPushButton("Generate Stats")
        self.run_btn.setMinimumWidth(160)
        self.run_btn.setMinimumHeight(36)
        self.run_btn.clicked.connect(self._run)
        btn_row.addWidget(self.run_btn)
        btn_row.addStretch()
        root.addLayout(btn_row)

        status_box = QGroupBox("Status")
        status_layout = QVBoxLayout(status_box)
        self.status = QTextEdit()
        self.status.setReadOnly(True)
        self.status.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        status_layout.addWidget(self.status)
        root.addWidget(status_box, stretch=1)

    def _browse_log(self):
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Select Minecraft chat log",
            "",
            "Log files (*.log);;Text files (*.txt);;All files (*.*)",
        )
        if path:
            self.log_edit.setText(path)

    def _browse_out(self):
        path, _ = QFileDialog.getSaveFileName(
            self,
            "Save Excel as",
            self.out_edit.text() or "player_stats.xlsx",
            "Excel files (*.xlsx);;All files (*.*)",
        )
        if path:
            self.out_edit.setText(path)

    def _append_status(self, msg: str):
        self.status.append(msg)

    def _run(self):
        log_path = self.log_edit.text().strip()
        towns_url = self.url_edit.text().strip()
        out_path = self.out_edit.text().strip()

        if not log_path:
            QMessageBox.critical(self, "Missing input", "Please select a chat log file.")
            return
        if not os.path.isfile(log_path):
            QMessageBox.critical(
                self, "File not found", f"Log file does not exist:\n{log_path}"
            )
            return
        if not towns_url:
            QMessageBox.critical(self, "Missing input", "Please enter the towns JSON URL.")
            return
        if not out_path:
            QMessageBox.critical(self, "Missing input", "Please enter an output Excel path.")
            return

        if not os.path.isabs(out_path):
            out_path = os.path.join(os.path.dirname(log_path) or os.getcwd(), out_path)
            self.out_edit.setText(out_path)

        self.status.clear()
        self._append_status("Starting…")
        self._append_status(f"  Log:   {log_path}")
        self._append_status(f"  URL:   {towns_url}")
        self._append_status(f"  Out:   {out_path}")
        self._append_status("")

        self.run_btn.setEnabled(False)

        self.worker = Worker(log_path, out_path, towns_url)
        self.worker.log.connect(self._append_status)
        self.worker.finished_ok.connect(self._on_success)
        self.worker.failed.connect(self._on_error)
        self.worker.start()

    def _on_success(self, out_path: str):
        self._append_status("")
        self._append_status(f"Done! Saved to:\n  {out_path}")
        self.run_btn.setEnabled(True)
        QMessageBox.information(self, "Success", f"Stats exported to:\n{out_path}")

    def _on_error(self, err: str):
        self._append_status(f"ERROR: {err}")
        self.run_btn.setEnabled(True)
        QMessageBox.critical(self, "Error", err)


def gui():
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    window = StatsWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
