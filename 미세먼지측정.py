import sys
import os
import serial
import serial.tools.list_ports
import time
import re
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QLabel, QFrame)
from PyQt5.QtCore import QThread, pyqtSignal, Qt
from PyQt5.QtGui import QFont, QIcon, QPixmap

# ==========================================
# 사용자 설정 영역
# ==========================================
BAUD_RATE = 9600

# 파일 절대 경로 지정
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
LOGO_FILE = os.path.join(BASE_DIR, "logo2.png")


# ==========================================
# 자동 포트 검색 및 대기질 평가 함수 (얼굴 아이콘 추가)
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
    # 💡 [수정] 수치에 따라 사진과 같은 얼굴 표정 이모티콘이 뜨도록 변경했습니다.
    if pm25 <= 15:
        pm25_stat, pm25_color = "😀 좋음", "#0078D7"
    elif pm25 <= 35:
        pm25_stat, pm25_color = "😐 보통", "#107C10"
    elif pm25 <= 75:
        pm25_stat, pm25_color = "🙁 나쁨", "#D83B01"
    else:
        pm25_stat, pm25_color = "😡 매우나쁨", "#A80000"

    if pm10 <= 30:
        pm10_stat, pm10_color = "😀 좋음", "#0078D7"
    elif pm10 <= 80:
        pm10_stat, pm10_color = "😐 보통", "#107C10"
    elif pm10 <= 150:
        pm10_stat, pm10_color = "🙁 나쁨", "#D83B01"
    else:
        pm10_stat, pm10_color = "😡 매우나쁨", "#A80000"
        
    return pm25_stat, pm25_color, pm10_stat, pm10_color


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
        self.setWindowTitle("Willbedone Data Logger")
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
        # 로고 영역
        # -----------------------------------
        self.logo_label = QLabel()
        self.logo_label.setAlignment(Qt.AlignCenter)
        
        if os.path.exists(LOGO_FILE):
            pixmap = QPixmap(LOGO_FILE)
            scaled_pixmap = pixmap.scaledToWidth(350, Qt.SmoothTransformation)
            self.logo_label.setPixmap(scaled_pixmap)
        else:
            self.logo_label.setText("WILL BE DONE")
            self.logo_label.setFont(QFont("Arial", 24, QFont.Bold))
            self.logo_label.setStyleSheet("color: #005A9E;")

        inner_layout.addWidget(self.logo_label)
        inner_layout.addSpacing(20)

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

    def create_data_row(self, parent_layout, title, key, has_status=False):
        row_layout = QHBoxLayout()
        row_layout.setContentsMargins(0, 10, 0, 10)

        # 1. 타이틀 (좌측 고정)
        title_label = QLabel(title)
        title_label.setFont(QFont("Malgun Gothic", 14))
        title_label.setFixedWidth(160)
        row_layout.addWidget(title_label)

        # 2. 빈 공간 (이 구문이 나머지 요소들을 우측으로 쫙 밀어줍니다)
        row_layout.addStretch(1)

        # 3. 상태(얼굴 표정) (우측 정렬)
        if has_status:
            stat_label = QLabel("")
            stat_label.setFont(QFont("Malgun Gothic", 15, QFont.Bold))
            stat_label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
            stat_label.setFixedWidth(120) # 글자 길이에 따라 수치가 흔들리지 않게 폭 고정
            self.stat_labels[key] = stat_label
            row_layout.addWidget(stat_label)
        else:
            # 상태 표시가 없는 PM 1.0도 줄을 맞추기 위해 투명한 공간 추가
            empty_space = QLabel("")
            empty_space.setFixedWidth(120)
            row_layout.addWidget(empty_space)

        # 4. 수치 (우측 정렬)
        val_label = QLabel("--")
        val_label.setFont(QFont("Arial", 22, QFont.Bold))
        val_label.setStyleSheet("color: #333333;")
        val_label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        val_label.setFixedWidth(60)
        self.val_labels[key] = val_label
        row_layout.addWidget(val_label)

        # 5. 단위 (가장 우측)
        unit_label = QLabel("µg/m³")
        unit_label.setFont(QFont("Malgun Gothic", 10))
        unit_label.setStyleSheet("color: gray;")
        unit_label.setAlignment(Qt.AlignBottom)
        unit_label.setContentsMargins(5, 0, 0, 4) # 수치와 약간의 여백
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

        pm25_stat, pm25_color, pm10_stat, pm10_color = get_air_quality_korea(pm25, pm10)

        self.stat_labels["pm25"].setText(pm25_stat)
        self.stat_labels["pm25"].setStyleSheet(f"color: {pm25_color};")
        
        self.stat_labels["pm10"].setText(pm10_stat)
        self.stat_labels["pm10"].setStyleSheet(f"color: {pm10_color};")

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