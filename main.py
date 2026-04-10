"""
Photostimulator Controller — PySide6
Sends a space-separated string of 6 frequencies over a serial port.
"""

import sys
import time
from pathlib import Path

import serial
import serial.tools.list_ports
from PySide6.QtCore import Qt, QThread, Signal
from PySide6.QtGui import QColor, QFont, QPalette
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QGridLayout, QLabel, QComboBox, QPushButton, QTextEdit,
    QGroupBox, QFrame, QSizePolicy, QMessageBox,
)

# Frequency list: "0", then 1000/i for i = 1..500.
# 6 decimal places, trailing zeros stripped — matches Arduino float precision
# and avoids scientific notation which Serial.parseFloat() cannot handle.
FREQUENCIES: list[str] = ["0"] + [
    f"{1000.0 / i:.6f}".rstrip("0").rstrip(".")
    for i in range(1, 501)
]

NUM_CHANNELS = 6
SETTINGS_FILE = Path("Settings.txt")


# ── Settings persistence ─────────────────────────────────────────────────────

def load_settings() -> list[str]:
    try:
        values = SETTINGS_FILE.read_text(encoding="utf-8").split()
        values = [v for v in values if v in FREQUENCIES]
        if len(values) == NUM_CHANNELS:
            return values
    except Exception:
        pass
    return ["0"] * NUM_CHANNELS


def save_settings(freqs: list[str]) -> None:
    try:
        SETTINGS_FILE.write_text(" ".join(freqs), encoding="utf-8")
    except Exception:
        pass


# ── Serial worker thread ─────────────────────────────────────────────────────

class SerialSendWorker(QThread):
    result = Signal(str)   # emitted with log message
    error  = Signal(str)

    def __init__(self, port: str, data: str):
        super().__init__()
        self.port = port
        self.data = data

    def run(self):
        try:
            with serial.Serial(self.port, baudrate=9600, timeout=1) as ser:
                # Append '\n' so Arduino's parseFloat() knows the last number
                # is complete without waiting for its 1-second serial timeout.
                ser.write((self.data + "\n").encode("utf-8"))
                # Wait for Arduino to process all 6 values and send replies.
                time.sleep(0.5)
                response = ser.read_all().decode("utf-8", errors="replace").strip()
            msg = f"=> {self.data}"
            if response:
                msg += f"\n<= {response}"
            self.result.emit(msg)
        except Exception as exc:
            self.error.emit(str(exc))


# ── Main Window ───────────────────────────────────────────────────────────────

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Photostimulator Controller")
        self.setMinimumWidth(480)
        self.setFixedWidth(480)

        self._worker: SerialSendWorker | None = None

        central = QWidget()
        self.setCentralWidget(central)
        root_layout = QVBoxLayout(central)
        root_layout.setContentsMargins(16, 16, 16, 16)
        root_layout.setSpacing(12)

        # ── Title ────────────────────────────────────────────────────────────
        title = QLabel("Photostimulator")
        title_font = QFont()
        title_font.setPointSize(15)
        title_font.setWeight(QFont.Weight.Bold)
        title.setFont(title_font)
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        root_layout.addWidget(title)

        sep = QFrame()
        sep.setFrameShape(QFrame.Shape.HLine)
        sep.setFrameShadow(QFrame.Shadow.Sunken)
        root_layout.addWidget(sep)

        # ── COM port row ──────────────────────────────────────────────────────
        port_group = QGroupBox("Serial Port")
        port_layout = QHBoxLayout(port_group)
        port_layout.setSpacing(8)

        self.cb_port = QComboBox()
        self.cb_port.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        port_layout.addWidget(self.cb_port)

        self.btn_refresh = QPushButton("↺  Refresh")
        self.btn_refresh.setFixedWidth(90)
        self.btn_refresh.clicked.connect(self.refresh_ports)
        port_layout.addWidget(self.btn_refresh)

        root_layout.addWidget(port_group)

        # ── Channel frequency grid ────────────────────────────────────────────
        chan_group = QGroupBox("Channel Frequencies")
        chan_grid = QGridLayout(chan_group)
        chan_grid.setHorizontalSpacing(12)
        chan_grid.setVerticalSpacing(6)

        # Header
        hdr_ch = QLabel("Channel")
        hdr_ch.setAlignment(Qt.AlignmentFlag.AlignCenter)
        hdr_freq = QLabel("Frequency (Hz)")
        hdr_freq.setAlignment(Qt.AlignmentFlag.AlignCenter)
        hdr_font = QFont()
        hdr_font.setWeight(QFont.Weight.Bold)
        hdr_ch.setFont(hdr_font)
        hdr_freq.setFont(hdr_font)
        chan_grid.addWidget(hdr_ch,   0, 0)
        chan_grid.addWidget(hdr_freq, 0, 1)

        self.freq_combos: list[QComboBox] = []
        saved = load_settings()

        for i in range(NUM_CHANNELS):
            lbl = QLabel(f"CH {i + 1}")
            lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)

            cb = QComboBox()
            cb.addItems(FREQUENCIES)
            cb.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
            # restore saved value
            idx = cb.findText(saved[i])
            if idx >= 0:
                cb.setCurrentIndex(idx)

            chan_grid.addWidget(lbl, i + 1, 0)
            chan_grid.addWidget(cb,  i + 1, 1)
            self.freq_combos.append(cb)

        chan_grid.setColumnStretch(1, 1)
        root_layout.addWidget(chan_group)

        # ── Send button ───────────────────────────────────────────────────────
        self.btn_send = QPushButton("▶  Send to Device")
        self.btn_send.setFixedHeight(38)
        send_font = QFont()
        send_font.setPointSize(10)
        send_font.setWeight(QFont.Weight.Medium)
        self.btn_send.setFont(send_font)
        self.btn_send.clicked.connect(self.send_frequencies)
        root_layout.addWidget(self.btn_send)

        # ── Log ───────────────────────────────────────────────────────────────
        log_group = QGroupBox("Log")
        log_layout = QVBoxLayout(log_group)

        self.log = QTextEdit()
        self.log.setReadOnly(True)
        self.log.setFixedHeight(110)
        self.log.setFont(QFont("Consolas", 9))
        log_layout.addWidget(self.log)

        btn_clear = QPushButton("Clear")
        btn_clear.setFixedWidth(70)
        btn_clear.clicked.connect(self.log.clear)
        log_layout.addWidget(btn_clear, alignment=Qt.AlignmentFlag.AlignRight)

        root_layout.addWidget(log_group)

        self.refresh_ports()
        self.adjustSize()
        self.setFixedSize(self.size())

    # ── Helpers ───────────────────────────────────────────────────────────────

    def refresh_ports(self):
        previous = self.cb_port.currentText()
        self.cb_port.clear()
        ports = [p.device for p in serial.tools.list_ports.comports()]
        ports.sort()
        self.cb_port.addItems(ports)
        # re-select previously chosen port if still present
        idx = self.cb_port.findText(previous)
        if idx >= 0:
            self.cb_port.setCurrentIndex(idx)
        elif self.cb_port.count() == 1:
            self.cb_port.setCurrentIndex(0)

    def current_frequencies(self) -> list[str]:
        return [cb.currentText() for cb in self.freq_combos]

    def build_serial_string(self) -> str:
        return " ".join(self.current_frequencies())

    def append_log(self, text: str):
        self.log.append(text)

    # ── Send ──────────────────────────────────────────────────────────────────

    def send_frequencies(self):
        port = self.cb_port.currentText()
        if not port:
            QMessageBox.warning(self, "No Port", "Please select a COM port first.")
            return

        data = self.build_serial_string()
        self.btn_send.setEnabled(False)
        self.btn_send.setText("Sending…")

        self._worker = SerialSendWorker(port, data)
        self._worker.result.connect(self._on_send_ok)
        self._worker.error.connect(self._on_send_error)
        self._worker.finished.connect(self._on_worker_done)
        self._worker.start()

    def _on_send_ok(self, msg: str):
        self.append_log(msg)

    def _on_send_error(self, err: str):
        self.append_log(f"[ERROR] {err}")
        QMessageBox.critical(self, "Serial Error", err)

    def _on_worker_done(self):
        self.btn_send.setEnabled(True)
        self.btn_send.setText("▶  Send to Device")

    # ── Close ─────────────────────────────────────────────────────────────────

    def closeEvent(self, event):
        save_settings(self.current_frequencies())
        super().closeEvent(event)


# ── Entry point ───────────────────────────────────────────────────────────────

def main():
    app = QApplication(sys.argv)
    app.setStyle("Fusion")

    # Clean dark-ish palette
    palette = QPalette()
    palette.setColor(QPalette.ColorRole.Window,          QColor(245, 245, 248))
    palette.setColor(QPalette.ColorRole.WindowText,      QColor(30,  30,  40))
    palette.setColor(QPalette.ColorRole.Base,            QColor(255, 255, 255))
    palette.setColor(QPalette.ColorRole.AlternateBase,   QColor(235, 235, 240))
    palette.setColor(QPalette.ColorRole.Button,          QColor(220, 220, 228))
    palette.setColor(QPalette.ColorRole.ButtonText,      QColor(30,  30,  40))
    palette.setColor(QPalette.ColorRole.Highlight,       QColor(60,  120, 210))
    palette.setColor(QPalette.ColorRole.HighlightedText, QColor(255, 255, 255))
    app.setPalette(palette)

    app.setStyleSheet("""
        QGroupBox {
            font-weight: bold;
            border: 1px solid #c0c0cc;
            border-radius: 6px;
            margin-top: 8px;
            padding-top: 4px;
        }
        QGroupBox::title {
            subcontrol-origin: margin;
            subcontrol-position: top left;
            padding: 0 6px;
            color: #3c3c50;
        }
        QPushButton {
            border: 1px solid #aaaabc;
            border-radius: 5px;
            padding: 5px 14px;
            background: qlineargradient(x1:0,y1:0,x2:0,y2:1,
                stop:0 #f0f0f8, stop:1 #d8d8e8);
        }
        QPushButton:hover {
            background: qlineargradient(x1:0,y1:0,x2:0,y2:1,
                stop:0 #e0e8ff, stop:1 #c8d4f8);
            border-color: #7090d0;
        }
        QPushButton:pressed {
            background: qlineargradient(x1:0,y1:0,x2:0,y2:1,
                stop:0 #c0ccee, stop:1 #a8b8e0);
        }
        QPushButton:disabled {
            color: #999;
            background: #e0e0e8;
        }
        QComboBox {
            border: 1px solid #b0b0c0;
            border-radius: 4px;
            padding: 3px 8px;
            background: white;
            color: #1a1a28;
        }
        QComboBox:focus { border-color: #5080d0; }
        QComboBox QAbstractItemView {
            background: white;
            color: #1a1a28;
            selection-background-color: #3c78d2;
            selection-color: white;
        }
        QTextEdit {
            border: 1px solid #c0c0cc;
            border-radius: 4px;
            background: #1e1e2e;
            color: #a0e0a0;
        }
    """)

    win = MainWindow()
    win.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
