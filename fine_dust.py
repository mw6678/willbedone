import csv
from datetime import datetime
import sys
import os
import serial
import serial.tools.list_ports
import time
import re
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                            QHBoxLayout, QLabel, QFrame)
from PyQt5.QtCore import QThread, pyqtSignal, Qt
from PyQt5.QtGui import QFont

# ==========================================
# ⚙️ [사용자 설정 영역]
# ==========================================
BAUD_RATE = 9600  # 시리얼 통신 보드레이트 설정

if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))


# ==========================================
# 📥 [수신 데이터 파싱 커스텀 함수]
# ==========================================
def parse_custom_data(raw_data):
    """콤마(,)로 구분된 데이터를 파싱하고, 값을 보정합니다."""
    try:
        parts = raw_data.replace(" ", "").split(',')
        
        if len(parts) >= 3:
            raw_pm1 = int(parts[0])
            raw_pm25 = int(parts[1])
            raw_pm10 = int(parts[2])
            
            pm1 = int(raw_pm1 * 1.0)  # 보정 계수 적용 (필요 시 조정 가능)
            pm25 = int(raw_pm25 * 1.0)
            pm10 = int(raw_pm10 * 1.0)
            
            return pm1, pm25, pm10
            
    except (ValueError, IndexError):
        pass
        
    return None, None, None


# ==========================================
# 🔍 [자동 포트 검색 및 대기질 평가 함수]
# ==========================================
def auto_detect_port():
    """시스템에 연결된 시리얼 포트 중 USB/UART 관련 포트를 자동 탐색합니다."""
    ports = serial.tools.list_ports.comports()
    if not ports:
        return None
    for port in ports:
        description = port.description.upper()
        if any(keyword in description for keyword in ["USB", "UART", "CH340", "CP210", "SERIAL"]):
            return port.device
    return ports[0].device


def get_air_quality_korea(pm25, pm10):
    """한국 환경부 기준에 따라 대기질 등급별 이모지와 색상을 반환합니다."""
    # 초미세먼지(PM2.5) 등급 판정
    if pm25 <= 15:
        pm25_emoji, pm25_color = "😊 좋음", "#0078D7"
    elif pm25 <= 35:
        pm25_emoji, pm25_color = "🙂 보통", "#107C10"
    elif pm25 <= 75:
        pm25_emoji, pm25_color = "😷 나쁨", "#D83B01"
    else:
        pm25_emoji, pm25_color = "😱 매우나쁨", "#A80000"

    # 미세먼지(PM10) 등급 판정
    if pm10 <= 30:
        pm10_emoji, pm10_color = "😊 좋음", "#0078D7"
    elif pm10 <= 80:
        pm10_emoji, pm10_color = "🙂 보통", "#107C10"
    elif pm10 <= 150:
        pm10_emoji, pm10_color = "😷 나쁨", "#D83B01"
    else:
        pm10_emoji, pm10_color = "😱 매우나쁨", "#A80000"
        
    return pm25_emoji, pm25_color, pm10_emoji, pm10_color


# ==========================================
# 🧵 [백그라운드 시리얼 통신 및 CSV 자동 저장 스레드]
# ==========================================
class SerialThread(QThread):
    status_signal = pyqtSignal(str)          # 연결 상태 메시지 전달 시그널
    data_signal = pyqtSignal(int, int, int)  # 측정된 미세먼지 데이터 전달 시그널

    def run(self):
        port_name = auto_detect_port()

        if port_name is None:
            self.status_signal.emit("❌ 센서를 찾을 수 없습니다.")
            return

        current_date = ""
        file_path = ""

        try:
            ser = serial.Serial(port_name, BAUD_RATE, timeout=1)
            self.status_signal.emit(f"✅ 연결됨: {port_name}")

            while not self.isInterruptionRequested():
                if ser.in_waiting > 0:
                    raw_data = ser.readline().decode("utf-8", errors="ignore").strip()
                    pm1, pm25, pm10 = parse_custom_data(raw_data)

                    if pm1 is not None and pm25 is not None and pm10 is not None:
                        self.data_signal.emit(pm1, pm25, pm10)

                        now = datetime.now()
                        today_str = now.strftime("%Y-%m-%d")
                        time_str = now.strftime("%H:%M:%S")

                        if current_date != today_str:
                            current_date = today_str
                            file_path = os.path.join(BASE_DIR, f"Dust_log_{current_date}.csv")
                            
                            is_new_file = not os.path.exists(file_path)
                            if is_new_file:
                                with open(file_path, mode='a', newline='', encoding='utf-8-sig') as f:
                                    writer = csv.writer(f)
                                    writer.writerow(["측정일자", "측정시간", "PM 1.0", "PM 2.5", "PM 10"])

                        try:
                            with open(file_path, mode='a', newline='', encoding='utf-8-sig') as f:
                                writer = csv.writer(f)
                                writer.writerow([today_str, time_str, pm1, pm25, pm10])
                        except PermissionError:
                            pass
                
                QThread.msleep(100) 

        except serial.SerialException:
            self.status_signal.emit(f"❌ 포트({port_name}) 접근 거부됨 (사용중)")
        finally:
            if 'ser' in locals() and ser.is_open:
                ser.close()


# ==========================================
# 🖥️ [메인 GUI 클래스 (PyQt5)]
# ==========================================
class DustMonitorApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.initUI()
        self.start_thread()

    def initUI(self):
        """메인 창 레이아웃 및 디자인을 초기화합니다."""
        self.setWindowTitle("WILL BE DONE - 미세먼지 모니터링")
        self.resize(450, 520)
        
        # 메인 윈도우 배경색 설정
        self.setStyleSheet("QMainWindow { background-color: #005A9E; }")

        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(8, 8, 8, 8)

        # 내부 카드형 프레임 생성 (흰색 배경)
        inner_frame = QFrame()
        inner_frame.setStyleSheet("background-color: white; border-radius: 8px;")
        inner_layout = QVBoxLayout(inner_frame)
        inner_layout.setContentsMargins(30, 30, 30, 30)

        # -----------------------------------
        # 타이틀 텍스트 영역 (로고 대체)
        # -----------------------------------
        self.title_label = QLabel("WILL BE DONE")
        self.title_label.setAlignment(Qt.AlignCenter)
        self.title_label.setFont(QFont("Arial", 22, QFont.Bold))
        self.title_label.setStyleSheet("color: #005A9E;")
        inner_layout.addWidget(self.title_label)
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
        inner_layout.addSpacing(15)

        # 구분선 생성
        line = QFrame()
        line.setFrameShape(QFrame.HLine)
        line.setStyleSheet("background-color: #E0E0E0;")
        inner_layout.addWidget(line)
        inner_layout.addSpacing(15)

        # -----------------------------------
        # 데이터 표시 행 생성 영역
        # -----------------------------------
        self.val_labels = {}
        self.stat_labels = {}

        self.create_data_row(inner_layout, "PM 1.0 (극초미세)", "pm1")
        self.create_data_row(inner_layout, "PM 2.5 (초미세먼지)", "pm25", has_status=True)
        self.create_data_row(inner_layout, "PM 10 (미세먼지)", "pm10", has_status=True)

        inner_layout.addStretch(1)
        main_layout.addWidget(inner_frame)

    def create_data_row(self, parent_layout, title, key, has_status=False):
        """미세먼지 측정값(PM1.0, PM2.5, PM10)을 표시할 UI 행을 동적으로 생성합니다."""
        row_layout = QHBoxLayout()
        row_layout.setContentsMargins(0, 10, 0, 10)

        # 항목 이름 레이블
        title_label = QLabel(title)
        title_label.setFont(QFont("Malgun Gothic", 12, QFont.Bold))
        title_label.setFixedWidth(250)
        row_layout.addWidget(title_label)

        row_layout.addStretch(1)

        # 상태 텍스트(이모지 등) 레이블
        stat_label = QLabel()
        stat_label.setAlignment(Qt.AlignCenter | Qt.AlignVCenter)
        stat_label.setFont(QFont("Malgun Gothic", 11, QFont.Bold))
        stat_label.setFixedWidth(90)
        
        if not has_status:
            stat_label.hide()  # 상태가 필요없는 항목(PM1.0 등)은 숨김
            
        self.stat_labels[key] = stat_label
        row_layout.addWidget(stat_label)

        # 수치 값 레이블
        val_label = QLabel("--")
        val_label.setFont(QFont("Arial", 18, QFont.Bold))
        val_label.setStyleSheet("color: #333333;")
        val_label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        val_label.setFixedWidth(60)
        self.val_labels[key] = val_label
        row_layout.addWidget(val_label)

        # 단위 레이블
        unit_label = QLabel("µg/m³")
        unit_label.setFont(QFont("Malgun Gothic", 9))
        unit_label.setStyleSheet("color: gray;")
        unit_label.setAlignment(Qt.AlignBottom)
        unit_label.setContentsMargins(5, 0, 0, 4)
        row_layout.addWidget(unit_label)

        parent_layout.addLayout(row_layout)

    def start_thread(self):
        """시리얼 통신 스레드를 생성하고 시그널을 슬롯 함수와 연결합니다."""
        self.serial_thread = SerialThread()
        self.serial_thread.status_signal.connect(self.update_status)
        self.serial_thread.data_signal.connect(self.update_data)
        self.serial_thread.start()

    def update_status(self, msg):
        """연결 상태 메시지 레이블을 갱신합니다."""
        self.status_label.setText(msg)

    def update_data(self, pm1, pm25, pm10):
        """수신된 미세먼지 수치를 UI 화면에 반영하고 상태별 이모지 및 색상을 업데이트합니다."""
        self.val_labels["pm1"].setText(str(pm1))
        self.val_labels["pm25"].setText(str(pm25))
        self.val_labels["pm10"].setText(str(pm10))

        # 등급 정보 가져오기
        pm25_emoji, pm25_color, pm10_emoji, pm10_color = get_air_quality_korea(pm25, pm10)

        # PM2.5 상태 적용
        self.stat_labels["pm25"].setText(pm25_emoji)
        self.stat_labels["pm25"].setStyleSheet(f"color: {pm25_color}; background-color: #F8F9FA; border-radius: 4px; padding: 4px;")

        # PM10 상태 적용
        self.stat_labels["pm10"].setText(pm10_emoji)
        self.stat_labels["pm10"].setStyleSheet(f"color: {pm10_color}; background-color: #F8F9FA; border-radius: 4px; padding: 4px;")

        # 현재 측정 시간 갱신
        self.time_label.setText(f"측정 시간: {time.strftime('%H:%M:%S')}")

    def closeEvent(self, event):
        """프로그램 종료 시 백그라운드 스레드를 안전하게 종료합니다."""
        if hasattr(self, 'serial_thread'):
            self.serial_thread.requestInterruption()
            self.serial_thread.wait()
        event.accept()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = DustMonitorApp()
    window.show()
    sys.exit(app.exec_())