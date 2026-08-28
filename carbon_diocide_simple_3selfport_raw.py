import csv
from datetime import datetime
import sys
import os
import serial
import time
import re
import traceback
import serial.tools.list_ports
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, 
    QHBoxLayout, QLabel, QFrame, QComboBox, 
    QPushButton, QDialog, QMessageBox)
from PyQt5.QtCore import QThread, pyqtSignal, Qt, QPoint
from PyQt5.QtGui import QFont

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# ==========================================
# [사용자 설정 영역]
# ==========================================
BAUD_RATE = 9600

# 센서별 2점 보정 파라미터 {센서index: (기울기 Slope, 절편 Offset)} 기울기 1.1 = ex) 300 -> 330
# 센서 1 = 0번, 센서 2 = 1번, 센서 3 = 2번
SENSOR_CALIB_PARAMS = {
    0: (1.0000, +0.0),  
    1: (1.094081, +0.0),  
    2: (1.0394, +0.0)   
}

# 센서 허용 측정 범위
CO2_MIN = 0
CO2_MAX = 30000

# 화면 표시용 이동평균 개수
SMOOTHING_WINDOW = 1

# ==========================================
#  [포트 선택 초기 화면 위젯]
# ==========================================
class PortSelectionDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.selected_ports = []
        self.port_rows = []  # 각 포트 설정 행(QWidget)을 관리하는 리스트

        self.setWindowTitle("CO2 센서 포트 설정")
        self.resize(360, 250)
        self.initUI()
        self.refresh_all_ports()  # 초기 포트 로드

    def initUI(self):
        self.main_layout = QVBoxLayout(self)
        self.main_layout.setContentsMargins(20, 20, 20, 20)
        self.main_layout.setSpacing(12)

        # 상단 타이틀 및 새로고침 버튼
        top_layout = QHBoxLayout()
        title = QLabel("CO2 센서 COM 포트 선택")
        title.setFont(QFont("Malgun Gothic", 12, QFont.Bold))
        
        self.refresh_btn = QPushButton("🔄 새로고침")
        self.refresh_btn.setFont(QFont("Malgun Gothic", 9))
        self.refresh_btn.setStyleSheet("padding: 4px 8px; border: 1px solid #ccc; border-radius: 4px;")
        self.refresh_btn.clicked.connect(self.refresh_all_ports)

        top_layout.addWidget(title)
        top_layout.addStretch(1)
        top_layout.addWidget(self.refresh_btn)
        self.main_layout.addLayout(top_layout)

        # 포트 선택 행들이 들어갈 컨테이너 레이아웃
        self.ports_layout = QVBoxLayout()
        self.ports_layout.setSpacing(8)
        self.main_layout.addLayout(self.ports_layout)

        # 기본으로 3개 행 추가 (초기 상태)
        for _ in range(3):
            self.add_port_row()

        # 하단 버튼 영역 (+ 센서 추가 버튼, 모니터링 시작 버튼)
        self.add_btn = QPushButton("+ 센서 추가")
        self.add_btn.setFixedHeight(30)
        self.add_btn.setFont(QFont("Malgun Gothic", 9))
        self.add_btn.setStyleSheet("background-color: #6C757D; color: white; border-radius: 4px;")
        self.add_btn.clicked.connect(self.add_port_row)
        self.main_layout.addWidget(self.add_btn)

        self.start_btn = QPushButton("모니터링 시작")
        self.start_btn.setFixedHeight(40)
        self.start_btn.setFont(QFont("Malgun Gothic", 10, QFont.Bold))
        self.start_btn.setStyleSheet("background-color: #007BFF; color: white; border-radius: 5px;")
        self.start_btn.clicked.connect(self.on_start_clicked)
        self.main_layout.addWidget(self.start_btn)

    def add_port_row(self):
        """포트 선택 행을 동적으로 추가"""
        row_widget = QWidget()
        row_layout = QHBoxLayout(row_widget)
        row_layout.setContentsMargins(0, 0, 0, 0)

        sensor_idx = len(self.port_rows) + 1
        label = QLabel(f"센서 {sensor_idx} 포트:")
        label.setFont(QFont("Malgun Gothic", 10))
        label.setFixedWidth(80)

        combo = QComboBox()
        combo.setStyleSheet("padding: 5px; border: 1px solid gray; border-radius: 3px;")

        del_btn = QPushButton("✕")
        del_btn.setFixedSize(28, 28)
        del_btn.setStyleSheet("background-color: #DC3545; color: white; border-radius: 3px; font-weight: bold;")
        del_btn.clicked.connect(lambda: self.remove_port_row(row_widget))

        row_layout.addWidget(label)
        row_layout.addWidget(combo)
        row_layout.addWidget(del_btn)

        self.ports_layout.addWidget(row_widget)
        self.port_rows.append({"widget": row_widget, "label": label, "combo": combo})

        # 새로 추가된 콤보박스에 최신 포트 목록 채우기
        self.refresh_combo_ports(combo)
        self.update_labels()

    def remove_port_row(self, row_widget):
        """포트 선택 행 삭제"""
        if len(self.port_rows) <= 1:
            QMessageBox.warning(self, "경고", "최소 1개 이상의 센서는 유지해야 합니다.")
            return

        for item in self.port_rows:
            if item["widget"] == row_widget:
                self.port_rows.remove(item)
                row_widget.deleteLater()
                break

        self.update_labels()

    def update_labels(self):
        """행 삭제/추가 시 센서 번호 갱신 (센서 1, 센서 2...)"""
        for idx, item in enumerate(self.port_rows):
            item["label"].setText(f"센서 {idx + 1} 포트:")

    def refresh_combo_ports(self, combo):
        """단일 콤보박스 포트 목록 갱신"""
        available_ports = [p.device for p in serial.tools.list_ports.comports()]
        port_options = ["선택 안함"] + available_ports if available_ports else ["선택 안함", "연결된 포트 없음"]

        current_text = combo.currentText()
        combo.clear()
        combo.addItems(port_options)

        if current_text in port_options:
            combo.setCurrentText(current_text)
        else:
            combo.setCurrentIndex(0)

    def refresh_all_ports(self):
        """전체 콤보박스 새로고침 및 기본 포트 자동 할당"""
        available_ports = [p.device for p in serial.tools.list_ports.comports()]
        
        for i, item in enumerate(self.port_rows):
            combo = item["combo"]
            current_text = combo.currentText()
            self.refresh_combo_ports(combo)

            # 기존 선택이 유지 가능한 경우 유지, 아니면 사용 가능한 포트를 순서대로 자동 배치
            if current_text in [p.device for p in serial.tools.list_ports.comports()]:
                combo.setCurrentText(current_text)
            elif len(available_ports) > i:
                combo.setCurrentText(available_ports[i])

    def on_start_clicked(self):
        raw_ports = [item["combo"].currentText() for item in self.port_rows]
        valid_ports = [p for p in raw_ports if p not in ["선택 안함", "연결된 포트 없음"]]

        if not valid_ports:
            QMessageBox.warning(self, "경고", "최소 하나 이상의 센서 포트를 선택해주세요.")
            return

        if len(set(valid_ports)) != len(valid_ports):
            reply = QMessageBox.question(self, "확인", "중복된 포트가 선택되었습니다.\n그래도 진행하시겠습니까?", QMessageBox.Yes | QMessageBox.No)
            if reply == QMessageBox.No:
                return

        self.selected_ports = valid_ports
        self.accept()

    def get_selected_ports(self):
        return self.selected_ports

# ==========================================
#  [백그라운드 개별 통신 스레드]
# ==========================================
class SerialThread(QThread):
    data_signal = pyqtSignal(int, int) 
    error_signal = pyqtSignal(int, str) 
    base_co2_signal = pyqtSignal(float)

    def __init__(self, port_name, sensor_index): 
        super().__init__()
        self.current_port = port_name  
        self.sensor_index = sensor_index
        self.smoothing_window = 2 
        self.data_buffer = []
        self.minute_data_buffer = [] 
        self.last_recorded_status = "정상"  # 1분 동안 발생한 에러 상태 추적용

    def parse_data(self, raw_data):
        try:
            raw_data = raw_data.strip()
            if not raw_data:
                return None

            raw_co2 = None
            if "CO2" in raw_data.upper():
                match = re.search(r'CO2\s*[:=]?\s*([-+]?\d*\.?\d+)', raw_data, re.IGNORECASE)
                if match:
                    raw_co2 = float(match.group(1))
            else:
                numbers = re.findall(r"[-+]?\d*\.?\d+", raw_data)
                if len(numbers) == 1:
                    raw_co2 = float(numbers[0])

            if raw_co2 is None:
                return None

            if raw_co2 < CO2_MIN or raw_co2 > CO2_MAX:
                return "OUT_OF_RANGE"

            return raw_co2
        except Exception:
            return None

    def safe_sleep(self, ms): 
        for _ in range(ms // 100):
            if self.isInterruptionRequested():
                return  
            QThread.msleep(100)

    def write_log(self, message):
        log_dir = os.path.join(BASE_DIR, "Logs")
        os.makedirs(log_dir, exist_ok=True)
        log_file = os.path.join(log_dir, f"error_log_Sensor{self.sensor_index + 1}.txt")
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        try:
            with open(log_file, "a", encoding="utf-8") as f:
                # 에러 이유가 명확히 남도록 메시지 포맷 강화
                f.write(f"{timestamp} | 포트: {self.current_port} | 상태/사유: {message}\n")
        except Exception:
            pass

    def cleanup_old_logs(self):
        retention_days = 30
        threshold_time = time.time() - (retention_days * 24 * 60 * 60)
        directories = ["CSV_Logs", "Logs"]

        for dir_name in directories:
            target_dir = os.path.join(BASE_DIR, dir_name)
            if not os.path.exists(target_dir):
                continue
            for filename in os.listdir(target_dir):
                file_path = os.path.join(target_dir, filename)
                if os.path.isfile(file_path):
                    try:
                        if os.path.getmtime(file_path) < threshold_time:
                            os.remove(file_path)
                    except Exception:
                        pass

    def run(self):
        current_date = ""
        last_ui_update_time =  0  
        is_connected = False 
        is_no_data_error_sent = False 

        last_saved_minute_key = datetime.now().strftime("%Y-%m-%d %H:%M")

        while not self.isInterruptionRequested():
            ser = None
            try:
                ser = serial.Serial(self.current_port, BAUD_RATE, timeout=0.1)
                
                if not is_connected: 
                    self.write_log(f"통신 연결 성공")
                    is_connected = True
                    self.data_buffer.clear()
                    self.last_recorded_status = "정상"
                    ser.reset_input_buffer()

                last_data_time = time.time()

                while ser.is_open and not self.isInterruptionRequested():
                    calibrated_co2_value = None
                    current_time = time.time()

                    while ser.in_waiting > 0:
                        raw_data = ser.readline().decode("utf-8", errors="ignore").strip()
                        if not raw_data: continue

                        parsed = self.parse_data(raw_data)
                        
                        if parsed == "OUT_OF_RANGE":
                            err_msg = "범위 초과 (위험)"
                            self.error_signal.emit(self.sensor_index, err_msg)
                            self.write_log(err_msg)
                            self.last_recorded_status = err_msg
                            last_data_time = current_time
                            continue

                        if parsed is not None:
                            # 1. 해당 센서의 (기울기, 절편) 가져오기 (설정이 없으면 기본값 1.0, 0.0)
                            slope, offset = SENSOR_CALIB_PARAMS.get(self.sensor_index, (1.0, 0.0))
    
                            # 2. 2점 보정 공식 적용: (Raw * 기울기) + 절편
                            calibrated_co2_value = (parsed * slope) + offset
    
                            last_data_time = current_time
                            self.last_recorded_status = "정상"

                    # 1. 통신 끊김 (30초 이상 데이터 없음)
                    if current_time - last_data_time >= 30.0:
                        if not is_no_data_error_sent:
                            err_msg = "데이터 없음 (30초 초과)"
                            self.error_signal.emit(self.sensor_index, "데이터 없음")
                            self.write_log(err_msg)
                            self.last_recorded_status = "데이터 없음"
                            is_no_data_error_sent = True
                            self.data_buffer.clear()

                    # 2. 정상 데이터 버퍼 추가
                    elif calibrated_co2_value is not None and isinstance(calibrated_co2_value, (int, float)):
                        is_no_data_error_sent = False
                        self.data_buffer.append(calibrated_co2_value)
                        self.minute_data_buffer.append(calibrated_co2_value)
                        
                        if len(self.data_buffer) > self.smoothing_window:
                            self.data_buffer.pop(0)

                        if current_time - last_ui_update_time >= 1.0:
                            last_ui_update_time = current_time
                            smoothed_co2 = int(sum(self.data_buffer) / len(self.data_buffer))
                            self.data_signal.emit(self.sensor_index, smoothed_co2)

                    # 3. 1분 단위 CSV 저장 로직 (데이터가 없거나 에러여도 1분마다 기록)
                    now = datetime.now()
                    current_minute_key = now.strftime("%Y-%m-%d %H:%M")
                    
                    if current_minute_key != last_saved_minute_key:
                        prev_minute_dt = datetime.strptime(last_saved_minute_key, "%Y-%m-%d %H:%M")
                        today_str = prev_minute_dt.strftime("%Y-%m-%d")
                        time_str = prev_minute_dt.strftime("%H:%M:00")

                        save_dir = os.path.join(BASE_DIR, "CSV_Logs")
                        os.makedirs(save_dir, exist_ok=True)

                        if current_date != today_str:
                            current_date = today_str
                            self.cleanup_old_logs()
                            
                        file_path = os.path.join(save_dir, f"CO2_log_{current_date}_Sensor{self.sensor_index + 1}.csv")
                        file_exists = os.path.exists(file_path)

                        # 저장할 값 결정 (정상이면 1분 평균, 아니면 발생했던 에러 상태 문자열 기록)
                        if self.minute_data_buffer:
                            record_value = int(sum(self.minute_data_buffer) / len(self.minute_data_buffer))
                            self.minute_data_buffer.clear()
                        else:
                            record_value = f"Error: {self.last_recorded_status}"

                        try:
                            with open(file_path, mode='a', newline='', encoding='utf-8-sig') as f:
                                writer = csv.writer(f)
                                if not file_exists:
                                    writer.writerow(["측정일자", "측정시간", "센서번호", "포트", "Co2 1분 평균(ppm)"])
                                writer.writerow([today_str, time_str, self.sensor_index + 1, self.current_port, record_value])
                            
                            last_saved_minute_key = current_minute_key  

                        except Exception as e:
                            self.write_log(f"CSV 저장 중 시스템 오류 발생: {str(e)}")
                            self.error_signal.emit(self.sensor_index, "CSV 저장 오류")

                    self.safe_sleep(100)

            except serial.SerialException as se:
                err_msg = f"시리얼 연결 실패/끊김: {se}"
                if is_connected: 
                    self.write_log(err_msg)
                    is_connected = False
                self.data_buffer.clear()
                self.error_signal.emit(self.sensor_index, "연결 실패/끊김")
                self.last_recorded_status = "연결 끊김"
            except Exception as e:
                err_msg = f"시스템 예외 오류: {traceback.format_exc()}"
                if is_connected:
                    self.write_log(err_msg)
                    is_connected = False
                self.data_buffer.clear()
                self.error_signal.emit(self.sensor_index, "시스템 오류")
                self.last_recorded_status = "시스템 오류"
                
            finally:
                if ser is not None and ser.is_open: 
                    try:
                        ser.close() 
                    except Exception:
                        pass
                self.safe_sleep(3000)
                
# ==========================================
#  [CO2 레벨 표시 위젯] (기존 동일)
# ==========================================
class CO2LevelWidget(QWidget): #사랑넷 원하는 기준표
    LEVELS = [
        {"name": "좋음", "min": 1, "max": 500, "color": "#28A745"},        
        {"name": "보통", "min": 501, "max": 1000, "color": "#FFD700"},        
        {"name": "민감군", "min": 1001, "max": 3000, "color": "#FD7E14"},     
        {"name": "나쁨", "min": 3001, "max": 5000, "color": "#DC3545"},      
        {"name": "매우 나쁨", "min": 5001, "max": 10000, "color": "#800080"}, 
        {"name": "위험", "min": 10001, "max": 30000, "color": "#795548"},      
    ]

    def __init__(self, sensor_index, title="센서", parent=None):
        super().__init__(parent)
        self.sensor_index = sensor_index # 센서 번호 저장
        self.title = title
        self.current_value = 400 
        self.initUI()

    def initUI(self):
        self.setStyleSheet("background-color: transparent; border: none;")
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(20, 20, 20, 20)
        main_layout.setSpacing(5)

        self.title_label = QLabel(self.title)
        self.title_label.setFont(QFont("Malgun Gothic", 16, QFont.Bold))
        self.title_label.setAlignment(Qt.AlignCenter)
        self.title_label.setStyleSheet("color: black; border: none;")
        main_layout.addWidget(self.title_label)
        
        main_layout.addSpacing(5)

        value_layout = QHBoxLayout()
        value_layout.setAlignment(Qt.AlignCenter)
        value_layout.setSpacing(2)

        self.co2_value_label = QLabel("----")
        self.co2_value_label.setFont(QFont("Arial", 42, QFont.Bold))
        self.co2_value_label.setStyleSheet("color: black; border: none;")
        value_layout.addWidget(self.co2_value_label)

        ppm_label = QLabel("ppm")
        ppm_label.setFont(QFont("Arial", 16, QFont.Bold))
        ppm_label.setStyleSheet("color: black; margin-bottom: 8px; border: none;") 
        ppm_label.setAlignment(Qt.AlignBottom)
        value_layout.addWidget(ppm_label)

        main_layout.addLayout(value_layout)
        main_layout.addSpacing(10)

        self.arrow_container = QWidget()
        self.arrow_container.setStyleSheet("border: none;")
        self.arrow_container.setFixedHeight(20)
        
        self.arrow_label = QLabel("▼")
        self.arrow_label.setFont(QFont("Arial", 12, QFont.Bold))
        self.arrow_label.setStyleSheet("color: black; border: none;")
        self.arrow_label.setAlignment(Qt.AlignCenter)
        self.arrow_label.setFixedWidth(20)
        self.arrow_label.setParent(self.arrow_container)
        self.arrow_label.hide()
        main_layout.addWidget(self.arrow_container)

        level_bar_frame = QFrame()
        level_bar_frame.setStyleSheet("border: none;")
        self.level_bar_layout = QHBoxLayout(level_bar_frame)
        self.level_bar_layout.setContentsMargins(0, 0, 0, 0)
        self.level_bar_layout.setSpacing(2) 

        self.level_bars = []
        for i, level in enumerate(self.LEVELS):
            bar = QFrame()
            bar.setFixedHeight(14)
            
            border_radius = ""
            if i == 0:
                border_radius = "border-top-left-radius: 7px; border-bottom-left-radius: 7px;"
            elif i == len(self.LEVELS) - 1:
                border_radius = "border-top-right-radius: 7px; border-bottom-right-radius: 7px;"
                
            bar.setStyleSheet(f"background-color: {level['color']}; {border_radius}")
            self.level_bars.append(bar)
            self.level_bar_layout.addWidget(bar)

        main_layout.addWidget(level_bar_frame)
        main_layout.addSpacing(15)

        self.status_text_label = QLabel("대기 중...")
        self.status_text_label.setFont(QFont("Malgun Gothic", 16, QFont.Bold))
        self.status_text_label.setAlignment(Qt.AlignCenter)
        self.status_text_label.setFixedHeight(45)
        self.status_text_label.setStyleSheet("background-color: #E0E0E0; border-radius: 8px; color: gray;")
        
        main_layout.addWidget(self.status_text_label)
        # ----------------------------------------------------
        main_layout.addStretch(1)

    def update_co2(self, value):
        #  추가: 파싱된 데이터가 들어왔을 때 비로소 화살표와 보정창을 표시
        self.arrow_label.show()         
        
        self.current_value = value
        self.co2_value_label.setText(str(value))

        current_level = None
        for level in self.LEVELS:
            if value <= level['max']:
                current_level = level
                break
        
        if value < self.LEVELS[0]['min']: current_level = self.LEVELS[0]
        elif value > self.LEVELS[-1]['max']: current_level = self.LEVELS[-1]

        if current_level:
            self.status_text_label.setText(current_level['name'])
            color = current_level['color']
            text_color = "black" if current_level['name'] == "보통" else "white"
            self.status_text_label.setStyleSheet(f"background-color: {color}; border-radius: 8px; color: {text_color};")

        self.update_arrow_position()

    def set_error_state(self, msg):
        self.co2_value_label.setText("----")
        self.status_text_label.setText(msg)
        self.status_text_label.setStyleSheet("background-color: #FFCDD2; border-radius: 8px; color: #B71C1C;")
        
        self.current_value = self.LEVELS[0]['min']
        self.update_arrow_position()
        self.arrow_label.hide()

    def update_arrow_position(self):
        if not self.level_bars or self.level_bars[0].geometry().width() == 0:
            return 

        value = self.current_value
        
        level_index = 0
        for i, level in enumerate(self.LEVELS):
            if value <= level['max']:
                level_index = i
                break
        else:
            level_index = len(self.LEVELS) - 1

        current_level = self.LEVELS[level_index]
        level_min = current_level['min']
        level_max = current_level['max']

        if value <= level_min: ratio = 0.0
        elif value >= level_max: ratio = 1.0
        else: ratio = (value - level_min) / (level_max - level_min)

        target_bar = self.level_bars[level_index]
        bar_x = target_bar.geometry().x()
        bar_width = target_bar.geometry().width()

        target_x = bar_x + (bar_width * ratio)
        arrow_x = int(target_x - (self.arrow_label.width() / 2))
        
        # 화살표가 컨테이너 밖으로 나가지 않도록 최소 0, 최대(컨테이너 너비 - 화살표 너비)로 제한(Clamping)
        max_x = self.arrow_container.width() - self.arrow_label.width()
        arrow_x = max(0, min(arrow_x, max_x))
        
        arrow_y = self.arrow_container.height() - self.arrow_label.height()

        self.arrow_label.move(QPoint(arrow_x, int(arrow_y)))

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self.update_arrow_position()

# ==========================================
#  [메인 GUI 클래스]
# ==========================================
class CO2MonitorApp(QMainWindow):
    def __init__(self, target_ports):
        super().__init__()
        self.target_ports = target_ports
        self.threads = []
        self.initUI()
        self.start_threads()
        self.dragPos = QPoint()

    def initUI(self):
        self.setWindowTitle("이산화탄소 다중 모니터링")
        
        # [수정] 선택된 포트 개수에 맞춰 창 가로 너비를 동적으로 계산 (1개당 약 260px)
        num_ports = len(self.target_ports)
        window_width = max(280, num_ports * 260)
        self.resize(window_width, 360) 
        
        self.setStyleSheet("QMainWindow { background-color: #E9ECEF; border: none; }") 
        self.setWindowFlags(Qt.FramelessWindowHint)

        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QHBoxLayout(central_widget)
        main_layout.setContentsMargins(15, 15, 15, 15)
        main_layout.setSpacing(15) 

        self.widgets = []
        for i, port in enumerate(self.target_ports):
            widget = CO2LevelWidget(sensor_index=i, title=f"센서 {i+1} ({port})") 
            self.widgets.append(widget)
            main_layout.addWidget(widget)

    def start_threads(self):
        for i, port_name in enumerate(self.target_ports): 
            thread = SerialThread(port_name, i) 
            thread.data_signal.connect(self.update_data)
            thread.error_signal.connect(self.handle_error) 
            self.threads.append(thread)
            
        for thread in self.threads:
            thread.start()

    def update_data(self, sensor_index, co2_value):
        if 0 <= sensor_index < len(self.widgets):
            self.widgets[sensor_index].update_co2(co2_value)

    def handle_error(self, sensor_index, error_msg):
        if 0 <= sensor_index < len(self.widgets):
            self.widgets[sensor_index].set_error_state(error_msg)

    def closeEvent(self, event):
        for thread in self.threads:
            thread.requestInterruption()
            thread.wait(1000)
        event.accept()

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.dragPos = event.globalPos() - self.frameGeometry().topLeft()

    def mouseMoveEvent(self, event):
        if event.buttons() == Qt.LeftButton:
            self.move(event.globalPos() - self.dragPos)

    def keyPressEvent(self, event):
        if event.key() == Qt.Key_Escape:
            self.close()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    
    # 1. 설정 다이얼로그 먼저 실행
    setup_dialog = PortSelectionDialog()
    
    # 2. 다이얼로그에서 '모니터링 시작'을 눌러 정상적으로 포트가 선택되었을 경우에만 메인 앱 실행
    if setup_dialog.exec_() == QDialog.Accepted:
        selected_ports = setup_dialog.get_selected_ports()
        
        # 메인 윈도우 생성 시 선택된 포트 리스트를 전달
        window = CO2MonitorApp(selected_ports)
        window.show()
        sys.exit(app.exec_())