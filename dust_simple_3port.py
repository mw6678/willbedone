import csv
from datetime import datetime
import sys
import os
import serial
import time
import re
import traceback
import serial.tools.list_ports

from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QFrame, QComboBox, QPushButton, QDialog, QMessageBox)
from PyQt5.QtCore import QThread, pyqtSignal, Qt, QPoint
from PyQt5.QtGui import QFont

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

BAUD_RATE = 9600

# 미세먼지/초미세먼지/극미세먼지 기준표
PM10_LEVELS = [
    {"name": "좋음", "min": 0, "max": 30, "color": "#28A745"},
    {"name": "보통", "min": 31, "max": 80, "color": "#FFD700"},
    {"name": "민감군", "min": 81, "max": 120, "color": "#FD7E14"},
    {"name": "나쁨", "min": 121, "max": 150, "color": "#DC3545"},
    {"name": "매우 나쁨", "min": 151, "max": 300, "color": "#800080"},
    {"name": "위험", "min": 301, "max": 600, "color": "#795548"},
]

PM25_LEVELS = [
    {"name": "좋음", "min": 0, "max": 15, "color": "#28A745"},
    {"name": "보통", "min": 16, "max": 35, "color": "#FFD700"},
    {"name": "민감군", "min": 36, "max": 50, "color": "#FD7E14"},
    {"name": "나쁨", "min": 51, "max": 75, "color": "#DC3545"},
    {"name": "매우 나쁨", "min": 76, "max": 100, "color": "#800080"},
    {"name": "위험", "min": 101, "max": 500, "color": "#795548"},
]

PM1_LEVELS = [
    {"name": "좋음", "min": 0, "max": 10, "color": "#28A745"},
    {"name": "보통", "min": 11, "max": 25, "color": "#FFD700"},
    {"name": "민감군", "min": 26, "max": 35, "color": "#FD7E14"},
    {"name": "나쁨", "min": 36, "max": 50, "color": "#DC3545"},
    {"name": "매우 나쁨", "min": 51, "max": 75, "color": "#800080"},
    {"name": "위험", "min": 76, "max": 300, "color": "#795548"},
]

# 1. 포트 선택 다이얼로그
class PortSelectionDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.selected_port = ""
        self.setWindowTitle("미세먼지 센서 포트 설정")
        self.initUI()
        self.refresh_ports()

    def initUI(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
        
        self.combo = QComboBox()
        layout.addWidget(QLabel("연결할 COM 포트를 선택하세요:"))
        layout.addWidget(self.combo)

        self.start_btn = QPushButton("모니터링 시작")
        self.start_btn.setFixedHeight(40)
        self.start_btn.clicked.connect(self.on_start_clicked)
        layout.addWidget(self.start_btn)

    def refresh_ports(self):
        ports = [p.device for p in serial.tools.list_ports.comports()]
        self.combo.clear()
        self.combo.addItems(ports if ports else ["연결된 포트 없음"])

    def on_start_clicked(self):
        port = self.combo.currentText()
        if port == "연결된 포트 없음":
            QMessageBox.warning(self, "경고", "포트를 연결해 주세요.")
            return
        self.selected_port = port
        self.accept()

# 2. 백그라운드 시리얼 스레드 (3개 수치 처리)
class SerialThread(QThread):
    # UI로 전달할 시그널: (PM10, PM2.5, PM1.0)
    data_signal = pyqtSignal(int, int, int) 
    error_signal = pyqtSignal(str)

    def __init__(self, port_name):
        super().__init__()
        self.port_name = port_name

    def parse_dust_data(self, raw_data):
        try:
            parts = raw_data.replace(" ", "").split(',')
            if len(parts) >= 3:
                pm10 = int(float(parts[0]))
                pm25 = int(float(parts[1]))
                pm1 = int(float(parts[2]))
                return pm10, pm25, pm1
        except Exception:
            pass
        return None, None, None

    def run(self):
        ser = None
        try:
            ser = serial.Serial(self.port_name, BAUD_RATE, timeout=0.1)
            ser.reset_input_buffer()

            while not self.isInterruptionRequested():
                if ser.in_waiting > 0:
                    raw_data = ser.readline().decode("utf-8", errors="ignore").strip()
                    pm10, pm25, pm1 = self.parse_dust_data(raw_data)

                    if pm10 is not None:
                        self.data_signal.emit(pm10, pm25, pm1)
                        self.save_csv(pm10, pm25, pm1)

                QThread.msleep(100)

        except serial.SerialException as se:
            self.error_signal.emit("연결 실패")
        finally:
            if ser and ser.is_open:
                ser.close()

    def save_csv(self, pm10, pm25, pm1):
        now = datetime.now()
        save_dir = os.path.join(BASE_DIR, "CSV_Logs")
        os.makedirs(save_dir, exist_ok=True)
        file_path = os.path.join(save_dir, f"Dust_log_{now.strftime('%Y-%m-%d')}.csv")
        file_exists = os.path.exists(file_path)

        try:
            with open(file_path, mode='a', newline='', encoding='utf-8-sig') as f:
                writer = csv.writer(f)
                if not file_exists:
                    writer.writerow(["측정일자", "측정시간", "포트", "PM10", "PM2.5", "PM1.0"])
                writer.writerow([now.strftime('%Y-%m-%d'), now.strftime('%H:%M:%S'), self.port_name, pm10, pm25, pm1])
        except Exception:
            pass

# 3. 데이터 표시용 개별 UI 카드
class DustWidget(QWidget):
    def __init__(self, title, levels):
        super().__init__()
        self.levels = levels
        self.initUI(title)

    def initUI(self, title):
        layout = QVBoxLayout(self)
        self.title_lbl = QLabel(title)
        self.title_lbl.setFont(QFont("Malgun Gothic", 14, QFont.Bold))
        self.title_lbl.setAlignment(Qt.AlignCenter)
        
        self.val_lbl = QLabel("----")
        self.val_lbl.setFont(QFont("Arial", 36, QFont.Bold))
        self.val_lbl.setAlignment(Qt.AlignCenter)

        self.status_lbl = QLabel("대기 중...")
        self.status_lbl.setFont(QFont("Malgun Gothic", 12, QFont.Bold))
        self.status_lbl.setAlignment(Qt.AlignCenter)
        self.status_lbl.setFixedHeight(30)

        layout.addWidget(self.title_lbl)
        layout.addWidget(self.val_lbl)
        layout.addWidget(self.status_lbl)

    def update_val(self, val):
        self.val_lbl.setText(str(val))
        for lvl in self.levels:
            if val <= lvl['max']:
                self.status_lbl.setText(lvl['name'])
                self.status_lbl.setStyleSheet(f"background-color: {lvl['color']}; color: white; border-radius: 5px;")
                break

# 4. 메인 어플리케이션
class DustMonitorApp(QMainWindow):
    def __init__(self, port):
        super().__init__()
        self.port = port
        self.initUI()
        self.start_thread()

    def initUI(self):
        self.setWindowTitle("미세먼지 통합 모니터링")
        self.resize(750, 250)
        
        central = QWidget()
        self.setCentralWidget(central)
        layout = QHBoxLayout(central)

        self.w_pm10 = DustWidget("미세먼지 (PM10)", PM10_LEVELS)
        self.w_pm25 = DustWidget("초미세먼지 (PM2.5)", PM25_LEVELS)
        self.w_pm1 = DustWidget("극미세먼지 (PM1.0)", PM1_LEVELS)

        layout.addWidget(self.w_pm10)
        layout.addWidget(self.w_pm25)
        layout.addWidget(self.w_pm1)

    def start_thread(self):
        self.thread = SerialThread(self.port)
        self.thread.data_signal.connect(self.update_data)
        self.thread.start()

    def update_data(self, pm10, pm25, pm1):
        self.w_pm10.update_val(pm10)
        self.w_pm25.update_val(pm25)
        self.w_pm1.update_val(pm1)

    def closeEvent(self, event):
        self.thread.requestInterruption()
        self.thread.wait(1000)
        event.accept()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    dlg = PortSelectionDialog()
    if dlg.exec_() == QDialog.Accepted:
        win = DustMonitorApp(dlg.selected_port)
        win.show()
        sys.exit(app.exec_())