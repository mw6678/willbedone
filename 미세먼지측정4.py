import sys
import os
import serial
import serial.tools.list_ports
import time
import re
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QLabel, QFrame)
from PyQt5.QtCore import QThread, pyqtSignal, Qt, QRectF
from PyQt5.QtGui import QFont, QIcon, QPixmap, QPainter

# ==========================================
# 사용자 설정 영역
# ==========================================
BAUD_RATE = 9600

# 파일 절대 경로 지정
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
LOGO_FILE = os.path.join(BASE_DIR, "logo2.png")


# ==========================================
# 자동 포트 검색 및 대기질 평가 함수
# ==========================================
def auto_detect_port():
    ports = serial.tools.list_ports.comports()
    if not ports:
        return None
    for port in ports:
        description = port.description.upper()
        if any(keyword in description for keyword in ["USB", "UART", "CH340", "CP210", "SERIAL"]):
            return port.device
    return ports[0].device


def get_air_quality_korea(pm25, pm10):
    if pm25 <= 15:
        pm25_img, pm25_color = "face_good.png", "#0078D7"
    elif pm25 <= 35:
        pm25_img, pm25_color = "face_normal.png", "#107C10"
    elif pm25 <= 75:
        pm25_img, pm25_color = "face_bad.png", "#D83B01"
    else:
        pm25_img, pm25_color = "face_very_bad.png", "#A80000"

    if pm10 <= 30:
        pm10_img, pm10_color = "face_good.png", "#0078D7"
    elif pm10 <= 80:
        pm10_img, pm10_color = "face_normal.png", "#107C10"
    elif pm10 <= 150:
        pm10_img, pm10_color = "face_bad.png", "#D83B01"
    else:
        pm10_img, pm10_color = "face_very_bad.png", "#A80000"
        
    return pm25_img, pm25_color, pm10_img, pm10_color


# ==========================================
# 백그라운드 시리얼 통신 스레드
# ==========================================
class SerialThread(QThread):
    status_signal = pyqtSignal(str)
    data_signal = pyqtSignal(int, int, int)

    def run(self):
        port_name = auto_detect_port()

        if port_name is None:
            self.status_signal.emit("❌ 센서를 찾을 수 없습니다.")
            return

        try:
            ser = serial.Serial(port_name, BAUD_RATE, timeout=1)
            self.status_signal.emit(f"✅ 연결됨: {port_name}")

            while not self.isInterruptionRequested():
                if ser.in_waiting > 0:
                    raw_data = ser.readline().decode("utf-8", errors="ignore").strip()
                    numbers = re.findall(r'\d+', raw_data)

                    if len(numbers) >= 3:
                        try:
                            pm1 = int(numbers[0])
                            pm25 = int(numbers[1])
                            pm10 = int(numbers[2])
                            
                            self.data_signal.emit(pm1, pm25, pm10)
                        except ValueError:
                            pass
                
                QThread.msleep(100) 

        except serial.SerialException:
            self.status_signal.emit(f"❌ 포트({port_name}) 접근 거부됨 (사용중)")
        finally:
            if 'ser' in locals() and ser.is_open:
                ser.close()


# ==========================================
# 메인 GUI (PyQt5)
# ==========================================
class DustMonitorApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.initUI()
        self.start_thread()

    def initUI(self):
        self.setWindowTitle("Will  be  done")
        self.resize(450, 550)
        
        if os.path.exists(LOGO_FILE):
            self.setWindowIcon(QIcon(LOGO_FILE))

        self.setStyleSheet("QMainWindow { background-color: #005A9E; }")

        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(8, 8, 8, 8)

        inner_frame = QFrame()
        inner_frame.setStyleSheet("background-color: white; border-radius: 5px;")
        inner_layout = QVBoxLayout(inner_frame)
        inner_layout.setContentsMargins(30, 30, 30, 30)

        # -----------------------------------
        # 로고 영역 (완벽 자동 조절 적용)
        # -----------------------------------
        self.logo_label = QLabel()
        self.logo_label.setAlignment(Qt.AlignCenter)
        self.logo_label.setFixedHeight(70) # 로고가 들어갈 공간 높이 고정
        
        if os.path.exists(LOGO_FILE):
            self.original_logo_pixmap = QPixmap(LOGO_FILE)
            self.update_logo_image() # 크기에 맞춰 자동 조절 함수 호출
        else:
            self.logo_label.setText("WILL BE DONE")
            self.logo_label.setFont(QFont("Arial", 24, QFont.Bold))
            self.logo_label.setStyleSheet("color: #005A9E;")

        inner_layout.addWidget(self.logo_label)
        inner_layout.addSpacing(10)

        # -----------------------------------
        # 상태 표시 및 시간 영역
        # -----------------------------------
        self.status_label = QLabel("🔍 센서 찾는 중...")
        self.status_label.setFont(QFont("Malgun Gothic", 10))
        self.status_label.setStyleSheet("color: gray;")
        self.status_label.setAlignment(Qt.AlignCenter)

        self.time_label = QLabel("시간: --:--:--")
        self.time_label.setFont(QFont("Malgun Gothic", 10))
        self.time_label.setStyleSheet("color: gray;")
        self.time_label.setAlignment(Qt.AlignCenter)

        inner_layout.addWidget(self.status_label)
        inner_layout.addWidget(self.time_label)
        inner_layout.addSpacing(10)

        line = QFrame()
        line.setFrameShape(QFrame.HLine)
        line.setStyleSheet("background-color: #E0E0E0;")
        inner_layout.addWidget(line)
        inner_layout.addSpacing(10)

        # -----------------------------------
        # 데이터 표시 영역
        # -----------------------------------
        self.val_labels = {}
        self.stat_labels = {}

        self.create_data_row(inner_layout, "PM 1.0 (극초미세)", "pm1")
        self.create_data_row(inner_layout, "PM 2.5 (초미세)", "pm25", has_status=True)
        self.create_data_row(inner_layout, "PM 10 (미세먼지)", "pm10", has_status=True)

        inner_layout.addStretch(1)
        main_layout.addWidget(inner_frame)

    # 창 크기가 바뀔 때 로고가 자동으로 재조절되도록 이벤트 추가
    def resizeEvent(self, event):
        super().resizeEvent(event)
        if os.path.exists(LOGO_FILE) and hasattr(self, 'original_logo_pixmap'):
            self.update_logo_image()

    def update_logo_image(self):
        # 💡 라벨의 현재 크기에 맞춰 로고를 찌그러짐 없이 비율 유지하며 자동 축소/확대
        scaled_pixmap = self.original_logo_pixmap.scaled(
            self.logo_label.width(), 
            self.logo_label.height(), 
            Qt.KeepAspectRatio, 
            Qt.SmoothTransformation
        )
        self.logo_label.setPixmap(scaled_pixmap)

    def create_data_row(self, parent_layout, title, key, has_status=False):
        row_layout = QHBoxLayout()
        row_layout.setContentsMargins(0, 10, 0, 10)

        title_label = QLabel(title)
        title_label.setFont(QFont("Malgun Gothic", 13))
        title_label.setFixedWidth(150)
        row_layout.addWidget(title_label)

        row_layout.addStretch(1)

        stat_label = QLabel()
        stat_label.setAlignment(Qt.AlignCenter | Qt.AlignVCenter)
        stat_label.setFixedSize(40, 40)
        
        if not has_status:
            stat_label.hide()
            
        self.stat_labels[key] = stat_label
        row_layout.addWidget(stat_label)

        val_label = QLabel("--")
        val_label.setFont(QFont("Arial", 20, QFont.Bold))
        val_label.setStyleSheet("color: #333333;")
        val_label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        val_label.setFixedWidth(60)
        self.val_labels[key] = val_label
        row_layout.addWidget(val_label)

        unit_label = QLabel("µg/m³")
        unit_label.setFont(QFont("Malgun Gothic", 10))
        unit_label.setStyleSheet("color: gray;")
        unit_label.setAlignment(Qt.AlignBottom)
        unit_label.setContentsMargins(5, 0, 0, 4)
        row_layout.addWidget(unit_label)

        parent_layout.addLayout(row_layout)

    def start_thread(self):
        self.serial_thread = SerialThread()
        self.serial_thread.status_signal.connect(self.update_status)
        self.serial_thread.data_signal.connect(self.update_data)
        self.serial_thread.start()

    def update_status(self, msg):
        self.status_label.setText(msg)

    def update_data(self, pm1, pm25, pm10):
        self.val_labels["pm1"].setText(str(pm1))
        self.val_labels["pm25"].setText(str(pm25))
        self.val_labels["pm10"].setText(str(pm10))

        pm25_img, pm25_color, pm10_img, pm10_color = get_air_quality_korea(pm25, pm10)

        full_path_25 = os.path.join(BASE_DIR, pm25_img)
        if os.path.exists(full_path_25):
            pixmap_25 = QPixmap(full_path_25).scaled(38, 38, Qt.KeepAspectRatio, Qt.SmoothTransformation)
            self.stat_labels["pm25"].setPixmap(pixmap_25)
        else:
            self.stat_labels["pm25"].setText("😊")

        full_path_10 = os.path.join(BASE_DIR, pm10_img)
        if os.path.exists(full_path_10):
            pixmap_10 = QPixmap(full_path_10).scaled(38, 38, Qt.KeepAspectRatio, Qt.SmoothTransformation)
            self.stat_labels["pm10"].setPixmap(pixmap_10)
        else:
            self.stat_labels["pm10"].setText("😊")

        self.time_label.setText(f"측정 시간: {time.strftime('%H:%M:%S')}")

    def closeEvent(self, event):
        if hasattr(self, 'serial_thread'):
            self.serial_thread.requestInterruption()
            self.serial_thread.wait()
        event.accept()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    ex = DustMonitorApp()
    ex.show()
    sys.exit(app.exec_())