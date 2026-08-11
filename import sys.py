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
from PyQt5.QtCore import QThread, pyqtSignal, Qt, QRectF
from PyQt5.QtGui import QFont, QIcon, QPixmap, QPainter

import sys
import os

# ==========================================
# ⚙️ [사용자 설정 영역]
# ==========================================
BAUD_RATE = 9600  # 시리얼 통신 보드레이트 설정

# [수정된 부분] .exe 파일로 실행할 때와 파이썬으로 실행할 때의 경로를 똑같이 맞춰줍니다.
if getattr(sys, 'frozen', False):
    # .exe 로 실행된 경우 실행 파일이 있는 폴더 경로
    BASE_DIR = os.path.dirname(sys.executable)
else:
    # 파이썬 스크립트로 실행된 경우 현재 파일 경로
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

LOGO_FILE = os.path.join(BASE_DIR, "logo2.png")


# ==========================================
# 📥 [수신 데이터 파싱 커스텀 함수]
# 💡 하이퍼터미널이나 센서가 보내는 데이터 형식에 맞추어 이 함수를 수정하세요!
# ==========================================
def parse_custom_data(raw_data):
    """ 콤마(,)로 구분된 데이터를 파싱하고, 값을 입맛에 맞게 수정(보정)합니다. """
    try:
        # 1. 공백 제거 후 콤마를 기준으로 분리
        parts = raw_data.replace(" ", "").split(',')
        
        if len(parts) >= 3:
            # 2. 먼저 원본 데이터를 정수로 변환해서 가져옴
            raw_pm1 = int(parts[0])
            raw_pm25 = int(parts[1])
            raw_pm10 = int(parts[2])
            
            # 3. 🎯 [여기서 원하는 대로 값을 수정/보정합니다] 🎯
            # (아래는 예시입니다. 상황에 맞게 수식을 변경하세요)
            
            pm1 = int(raw_pm1 * 1.0)       # 그대로 유지
            pm25 = int(raw_pm25 * 1.0)       # 예: 실제보다 낮게 나와서 일괄적으로 5를 더함
            pm10 = int(raw_pm10 * 1.0)     # 예: 측정값에 1.2배 가중치를 곱함
            
            # 만약 값이 음수가 나오는 것을 방지하려면 아래처럼 처리할 수도 있습니다.
            # pm25 = max(0, pm25) 
            
            # 4. 수정된 최종 값을 반환 (이 값이 UI에 출력됨)
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
        # 일반적인 USB 시리얼 변환기 키워드 탐색
        if any(keyword in description for keyword in ["USB", "UART", "CH340", "CP210", "SERIAL"]):
            return port.device
    # 키워드 매칭이 없을 경우 첫 번째 포트 반환
    return ports[0].device


def get_air_quality_korea(pm25, pm10):
    """한국 환경부 기준에 따라 초미세먼지(PM2.5)와 미세먼지(PM10) 등급별 이미지 및 색상을 반환합니다."""
    # 초미세먼지(PM2.5) 등급 판정
    if pm25 <= 15:
        pm25_img, pm25_color = "face_good.png", "#0078D7"      # 좋음 (파란색)
    elif pm25 <= 35:
        pm25_img, pm25_color = "face_normal.png", "#107C10"   # 보통 (초록색)
    elif pm25 <= 75:
        pm25_img, pm25_color = "face_bad.png", "#D83B01"      # 나쁨 (주황색)
    else:
        pm25_img, pm25_color = "face_very_bad.png", "#A80000" # 매우나쁨 (빨간색)

    # 미세먼지(PM10) 등급 판정
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
# 🧵 [백그라운드 시리얼 통신 스레드]
# UI 멈춤 현상을 방지하기 위해 별도 스레드에서 시리얼 데이터를 수신합니다.
# ==========================================
# ==========================================
# 🧵 [백그라운드 시리얼 통신 및 엑셀(CSV) 자동 저장 스레드]
# ==========================================
class SerialThread(QThread):
    status_signal = pyqtSignal(str)          # 연결 상태 메시지 전달 시그널
    data_signal = pyqtSignal(int, int, int)  # 측정된 미세먼지 데이터 전달 시그널

    def run(self):
        # 1. 포트 자동 검색 수행
        port_name = auto_detect_port()

        if port_name is None:
            self.status_signal.emit("❌ 센서를 찾을 수 없습니다.")
            return

        # 날짜별 파일 저장을 위한 변수 초기화
        current_date = ""
        file_path = ""

        try:
            # 2. 시리얼 포트 연결 시도
            ser = serial.Serial(port_name, BAUD_RATE, timeout=1)
            self.status_signal.emit(f"✅ 연결됨: {port_name}")

            # 3. 스레드가 중단될 때까지 반복 수신
            while not self.isInterruptionRequested():
                if ser.in_waiting > 0:
                    raw_data = ser.readline().decode("utf-8", errors="ignore").strip()
                    pm1, pm25, pm10 = parse_custom_data(raw_data)

                    if pm1 is not None and pm25 is not None and pm10 is not None:
                        # UI 화면 업데이트를 위해 시그널 전송
                        self.data_signal.emit(pm1, pm25, pm10)

                        # ----------------------------------------------------
                        # 💾 [엑셀(CSV) 파일 저장 로직 시작]
                        # ----------------------------------------------------
                        now = datetime.now()
                        today_str = now.strftime("%Y-%m-%d") # 예: 2026-08-11
                        time_str = now.strftime("%H:%M:%S")  # 예: 11:17:25

                        # 날짜가 바뀌었거나 프로그램 실행 후 처음 저장하는 경우
                        if current_date != today_str:
                            current_date = today_str
                            # 파일명 지정 (예: dust_log_2026-08-11.csv)
                            file_path = os.path.join(BASE_DIR, f"Dust_log_{current_date}.csv")
                            
                            # 파일이 없다면 새로 만들고 최상단에 헤더(제목) 작성
                            is_new_file = not os.path.exists(file_path)
                            if is_new_file:
                                # utf-8-sig 인코딩을 사용하면 엑셀에서 한글이 깨지지 않습니다.
                                with open(file_path, mode='a', newline='', encoding='utf-8-sig') as f:
                                    writer = csv.writer(f)
                                    writer.writerow(["측정일자", "측정시간", "PM 1.0", "PM 2.5", "PM 10"])

                        # 매번 수신된 데이터를 파일의 맨 마지막 줄에 추가(Append)
                        try:
                            with open(file_path, mode='a', newline='', encoding='utf-8-sig') as f:
                                writer = csv.writer(f)
                                writer.writerow([today_str, time_str, pm1, pm25, pm10])
                        except PermissionError:
                            # 만약 사용자가 해당 엑셀 파일을 띄워놓고 있어서 저장이 막힌 경우 예외 처리
                            pass
                        # ----------------------------------------------------
                
                # CPU 과부하 방지를 위한 미세 대기
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
        self.setWindowTitle("Will be done")
        self.resize(450, 550)
        
        # 윈도우 아이콘 설정
        if os.path.exists(LOGO_FILE):
            self.setWindowIcon(QIcon(LOGO_FILE))

        # 메인 윈도우 배경색 설정
        self.setStyleSheet("QMainWindow { background-color: #005A9E; }")

        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(8, 8, 8, 8)

        # 내부 카드형 프레임 생성 (흰색 배경)
        inner_frame = QFrame()
        inner_frame.setStyleSheet("background-color: white; border-radius: 5px;")
        inner_layout = QVBoxLayout(inner_frame)
        inner_layout.setContentsMargins(30, 30, 30, 30)

        # -----------------------------------
        # 로고 영역 (비율 유지 자동 조절)
        # -----------------------------------
        self.logo_label = QLabel()
        self.logo_label.setAlignment(Qt.AlignCenter)
        
        if os.path.exists(LOGO_FILE):
            self.original_logo_pixmap = QPixmap(LOGO_FILE)
            self.update_logo_image() 
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

        # 구분선 생성
        line = QFrame()
        line.setFrameShape(QFrame.HLine)
        line.setStyleSheet("background-color: #E0E0E0;")
        inner_layout.addWidget(line)
        inner_layout.addSpacing(10)

        # -----------------------------------
        # 데이터 표시 행 생성 영역
        # -----------------------------------
        self.val_labels = {}
        self.stat_labels = {}

        self.create_data_row(inner_layout, "PM 1.0 (극초미세)", "pm1")
        self.create_data_row(inner_layout, "PM 2.5 (초미세)", "pm25", has_status=True)
        self.create_data_row(inner_layout, "PM 10 (미세먼지)", "pm10", has_status=True)

        inner_layout.addStretch(1)
        main_layout.addWidget(inner_frame)

    def resizeEvent(self, event):
        """창 크기 변경 시 로고 이미지를 비율에 맞게 재조정합니다."""
        super().resizeEvent(event)
        if os.path.exists(LOGO_FILE) and hasattr(self, 'original_logo_pixmap'):
            self.update_logo_image()

    def update_logo_image(self):
        """로고 이미지의 찌그러짐 없이 크기에 맞춰 스케일링합니다."""
        scaled_pixmap = self.original_logo_pixmap.scaled(
            self.logo_label.width(), 
            self.logo_label.height(), 
            Qt.KeepAspectRatio, 
            Qt.SmoothTransformation
        )
        self.logo_label.setPixmap(scaled_pixmap)

    def create_data_row(self, parent_layout, title, key, has_status=False):
        """미세먼지 측정값(PM1.0, PM2.5, PM10)을 표시할 UI 행을 동적으로 생성합니다."""
        row_layout = QHBoxLayout()
        row_layout.setContentsMargins(0, 10, 0, 10)

        # 항목 이름 레이블
        title_label = QLabel(title)
        title_label.setFont(QFont("Malgun Gothic", 13))
        title_label.setFixedWidth(150)
        row_layout.addWidget(title_label)

        row_layout.addStretch(1)

        # 상태 아이콘(얼굴 이미지 등) 레이블
        stat_label = QLabel()
        stat_label.setAlignment(Qt.AlignCenter | Qt.AlignVCenter)
        stat_label.setFixedSize(40, 40)
        
        if not has_status:
            stat_label.hide()  # 상태가 필요없는 항목(PM1.0 등)은 숨김
            
        self.stat_labels[key] = stat_label
        row_layout.addWidget(stat_label)

        # 수치 값 레이블
        val_label = QLabel("--")
        val_label.setFont(QFont("Arial", 20, QFont.Bold))
        val_label.setStyleSheet("color: #333333;")
        val_label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        val_label.setFixedWidth(60)
        self.val_labels[key] = val_label
        row_layout.addWidget(val_label)

        # 단위 레이블
        unit_label = QLabel("µg/m³")
        unit_label.setFont(QFont("Malgun Gothic", 10))
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
        """수신된 미세먼지 수치를 UI 화면에 반영하고 상태별 아이콘을 업데이트합니다."""
        # 수치 텍스트 업데이트
        self.val_labels["pm1"].setText(str(pm1))
        self.val_labels["pm25"].setText(str(pm25))
        self.val_labels["pm10"].setText(str(pm10))

        # 대기질 평가 등급 이미지 가져오기
        pm25_img, pm25_color, pm10_img, pm10_color = get_air_quality_korea(pm25, pm10)

        # PM2.5 상태 이미지 설정
        full_path_25 = os.path.join(BASE_DIR, pm25_img)
        if os.path.exists(full_path_25):
            pixmap_25 = QPixmap(full_path_25).scaled(38, 38, Qt.KeepAspectRatio, Qt.SmoothTransformation)
            self.stat_labels["pm25"].setPixmap(pixmap_25)
        else:
            self.stat_labels["pm25"].setText("😊")

        # PM10 상태 이미지 설정
        full_path_10 = os.path.join(BASE_DIR, pm10_img)
        if os.path.exists(full_path_10):
            pixmap_10 = QPixmap(full_path_10).scaled(38, 38, Qt.KeepAspectRatio, Qt.SmoothTransformation)
            self.stat_labels["pm10"].setPixmap(pixmap_10)
        else:
            self.stat_labels["pm10"].setText("😊")

        # 현재 측정 시간 갱신
        self.time_label.setText(f"측정 시간: {time.strftime('%H:%M:%S')}")

    def closeEvent(self, event):
        """프로그램 종료 시 백그라운드 스레드를 안전하게 종료합니다."""
        if hasattr(self, 'serial_thread'):
            self.serial_thread.requestInterruption()
            self.serial_thread.wait()
        event.accept()
