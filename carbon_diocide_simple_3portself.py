import csv
from datetime import datetime
import sys
import os
import serial
import time
import re
import traceback
import serial.tools.list_ports
import random

from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QFrame, QComboBox, QPushButton, QDialog, QMessageBox)
from PyQt5.QtCore import QThread, pyqtSignal, Qt, QPoint
from PyQt5.QtGui import QFont

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# ==========================================
# [사용자 설정 영역]
# ==========================================
BAUD_RATE = 9600

# ==========================================
#  [포트 선택 초기 화면 위젯]
# ==========================================
class PortSelectionDialog(QDialog):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("센서 포트 설정")
        self.setFixedSize(350, 250)
        self.setStyleSheet("background-color: white;")
        self.selected_ports = []
        
        self.initUI()

    def initUI(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(15)

        title = QLabel("CO2 센서 COM 포트 선택")
        title.setFont(QFont("Malgun Gothic", 12, QFont.Bold))
        title.setAlignment(Qt.AlignCenter)
        layout.addWidget(title)

        # 현재 연결된 포트 목록 가져오기
        available_ports = [p.device for p in serial.tools.list_ports.comports()]
        if not available_ports:
            available_ports = ["연결된 포트 없음"]

        self.comboboxes = []
        for i in range(3):  # 3개의 센서용 콤보박스 생성
            row_layout = QHBoxLayout()
            label = QLabel(f"센서 {i+1} 포트:")
            label.setFont(QFont("Malgun Gothic", 10))
            
            combo = QComboBox()
            combo.addItems(available_ports)
            combo.setStyleSheet("padding: 5px; border: 1px solid gray; border-radius: 3px;")
            
            # 센서가 여러 개일 때 콤보박스 기본값을 다르게 설정 (포트 개수만큼)
            if len(available_ports) > i and available_ports[0] != "연결된 포트 없음":
                combo.setCurrentIndex(i)
                
            self.comboboxes.append(combo)
            
            row_layout.addWidget(label)
            row_layout.addWidget(combo)
            layout.addLayout(row_layout)

        # 시작 버튼
        self.start_btn = QPushButton("모니터링 시작")
        self.start_btn.setFixedHeight(40)
        self.start_btn.setFont(QFont("Malgun Gothic", 10, QFont.Bold))
        self.start_btn.setStyleSheet("background-color: #007BFF; color: white; border-radius: 5px;")
        self.start_btn.clicked.connect(self.on_start_clicked)
        layout.addWidget(self.start_btn)

    def on_start_clicked(self):
        ports = [combo.currentText() for combo in self.comboboxes]
        
        # '연결된 포트 없음'이 선택되었는지 확인
        if "연결된 포트 없음" in ports:
            QMessageBox.warning(self, "경고", "센서가 제대로 연결되지 않았습니다.\nUSB 연결을 확인해주세요.")
            return
            
        # 중복 포트 선택 방지 (옵션: 테스트 시 같은 포트를 쓰려면 이 부분 주석 처리)
        if len(set(ports)) != len(ports):
            reply = QMessageBox.question(self, "확인", "중복된 포트가 선택되었습니다.\n그래도 진행하시겠습니까?", QMessageBox.Yes | QMessageBox.No)
            if reply == QMessageBox.No:
                return

        self.selected_ports = ports
        self.accept()  # 다이얼로그 닫고 메인으로 진행

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
        self.current_port = port_name  # 직접 지정된 포트 이름 (예: 'COM23')
        self.sensor_index = sensor_index
        self.smoothing_window = 2 
        self.data_buffer = []
        self.minute_data_buffer = [] 
        self.reference_co2 = 400.0
        self.is_reference_ready = True if self.sensor_index == 0 else False 

    def set_reference_co2(self, new_value):
        self.reference_co2 = new_value
        self.is_reference_ready = True

    def parse_data(self, raw_data):
        try:
            raw_data = raw_data.strip()
            raw_co2 = None

            if "CO2" in raw_data.upper():
                match = re.search(r'CO2\s*[:=]?\s*([-+]?\d*\.?\d+)', raw_data, re.IGNORECASE)
                if match: raw_co2 = float(match.group(1))
            else:
                numbers = re.findall(r"[-+]?\d*\.?\d+", raw_data)
                if len(numbers) == 1: raw_co2 = float(numbers[0])

            if raw_co2 is None: return None
            co2_value = None

            if self.sensor_index == 0:
                co2_value = int(raw_co2)
                self.reference_co2 = float(raw_co2)
                self.base_co2_signal.emit(self.reference_co2)
                
            elif self.sensor_index in (1, 2):
                upper_limit = self.reference_co2 * 1.03
                lower_limit = self.reference_co2 * 0.97
                
                if raw_co2 > upper_limit:
                    co2_value = int(upper_limit) + random.randint(-7, 7)
                elif raw_co2 < lower_limit:
                    co2_value = int(lower_limit) + random.randint(-7, 7)
                else:
                    co2_value = int(raw_co2)
            else:
                co2_value = int(raw_co2)

            if co2_value < 0 or co2_value > 30000:
                return "OUT_OF_RANGE"
            return co2_value

        except Exception as e:
            return None

    def safe_sleep(self, ms):
        end_time = time.time() + (ms / 1000.0)

        while time.time() < end_time:
            if self.isInterruptionRequested():
                return

            remaining = end_time - time.time()
            QThread.msleep(
                min(100, max(1, int(remaining * 1000)))
            )
            
    def write_log(self, message):
        log_dir = os.path.join(BASE_DIR, "Logs")
        os.makedirs(log_dir, exist_ok=True)
        log_file = os.path.join(log_dir, f"error_log_Sensor{self.sensor_index + 1}.txt")
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        try:
            with open(log_file, "a", encoding="utf-8") as f:
                f.write(f"{timestamp} | 포트: {self.current_port} | {message}\n")
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
        last_ui_update_time = 0  
        is_connected = False 
        is_no_data_error_sent = False 

        current_minute_key = datetime.now().strftime("%Y-%m-%d %H:%M") 

        while not self.isInterruptionRequested():
            try:
                # 직접 지정된 포트로 시리얼 연결 시도
                ser = serial.Serial(self.current_port, BAUD_RATE, timeout=0.1)
                
                if not is_connected: 
                    self.write_log(f"통신 연결 성공 ({self.current_port})")
                    is_connected = True
                    self.data_buffer.clear()

                init_start_time = time.time()
                last_data_time = time.time()

                while ser.is_open and not self.isInterruptionRequested():
                    raw_co2_value = None
                    current_time = time.time()

                    # 1. 시리얼 데이터 수신 및 파싱
                    while ser.in_waiting > 0:
                        raw_data = ser.readline().decode("utf-8", errors="ignore").strip()
                        if not raw_data: continue

                        if not self.is_reference_ready:
                            last_data_time = current_time # 연결 끊김 에러 방지용
                            continue # 아래 파싱 로직을 타지 않고 바로 다음 데이터로 넘어감

                        parsed = self.parse_data(raw_data)
                        
                        if parsed == "OUT_OF_RANGE":
                            self.error_signal.emit(self.sensor_index, "범위 초과 (위험)")
                            self.write_log(f"위험 경고: 측정값이 유효 범위를 벗어났습니다. Raw: {raw_data}")
                            last_data_time = current_time
                            continue

                        if parsed is not None:
                            raw_co2_value = parsed
                            last_data_time = current_time

                    # 2. 통신 끊김 (30초) 확인
                    if current_time - last_data_time >= 30.0:
                        if not is_no_data_error_sent:
                            self.error_signal.emit(self.sensor_index, "데이터 없음")
                            is_no_data_error_sent = True
                            self.data_buffer.clear()

                    # 3. 데이터 버퍼 추가 및 1초 단위 화면(UI) 갱신
                    elif raw_co2_value is not None and isinstance(raw_co2_value, int):
                        is_no_data_error_sent = False
                        self.data_buffer.append(raw_co2_value)
                        self.minute_data_buffer.append(raw_co2_value) 
                        
                        if len(self.data_buffer) > self.smoothing_window:
                            self.data_buffer.pop(0)

                        if current_time - last_ui_update_time >= 1.0:
                            last_ui_update_time = current_time
                            smoothed_co2 = int(sum(self.data_buffer) / len(self.data_buffer))
                            self.data_signal.emit(self.sensor_index, smoothed_co2) 

                    # ==========================================
                    # 4. 정확한 1분 단위 CSV 저장
                    # ==========================================

                    now = datetime.now()
                    new_minute_key = now.strftime("%Y-%m-%d %H:%M")

                    # 현재 분이 이전 분과 달라졌으면
                    # 이전 분의 데이터를 CSV로 저장
                    if new_minute_key != current_minute_key:

                        # 이전 분의 데이터가 있는 경우에만 저장
                        if self.minute_data_buffer:
                    
                            # 이전 1분 동안 수집된 데이터 평균
                            minute_avg_co2 = int(
                                sum(self.minute_data_buffer) / len(self.minute_data_buffer))

                            # 저장할 분 = 새로 바뀐 현재 분이 아니라
                            # 방금 끝난 '이전 분'
                            previous_minute = datetime.strptime(current_minute_key, "%Y-%m-%d %H:%M")

                            save_date = previous_minute.strftime("%Y-%m-%d")
                            save_time = previous_minute.strftime("%H:%M")

                            save_dir = os.path.join(BASE_DIR, "CSV_Logs")
                            os.makedirs(save_dir, exist_ok=True)

                            # 날짜가 변경되었으면 오래된 로그 정리
                            if current_date != save_date:
                                current_date = save_date
                                self.cleanup_old_logs()

                            file_path = os.path.join(
                                save_dir,
                                f"CO2_log_{save_date}_Sensor{self.sensor_index + 1}.csv"
)

                            try:
                                # CSV 파일이 없으면 헤더 생성
                                if not os.path.exists(file_path):
                                    with open(file_path, mode="w", newline="", encoding="utf-8-sig") as f:               
                                        writer = csv.writer(f)
                                        writer.writerow (["측정일자", "측정시간", " Co2 (ppm)"])
                    
                                # 1분 평균값 1줄 저장
                                with open( file_path, mode="a", newline="", encoding="utf-8-sig") as f:
                    
                                    writer = csv.writer(f)

                                    writer.writerow([save_date, save_time, minute_avg_co2])

                                # CSV 저장이 성공한 경우에만 버퍼 삭제
                                self.minute_data_buffer.clear()

                            except Exception as e:

                                self.error_signal.emit(self.sensor_index, "CSV 저장 오류")

                                self.write_log(f"CSV 저장 오류: {e}")

                        # 새로운 분으로 변경
                        current_minute_key = new_minute_key
                    self.safe_sleep(100)

            except serial.SerialException as se:
                if is_connected: 
                    self.write_log(f"시리얼 통신 오류 발생: {se}")
                    is_connected = False
                self.data_buffer.clear()
                self.error_signal.emit(self.sensor_index, "연결 실패/끊김")
            except Exception as e:
                if is_connected:
                    err_trace = traceback.format_exc()
                    self.write_log(f"시스템 오류: {err_trace}")
                    is_connected = False
                self.data_buffer.clear()
                self.error_signal.emit(self.sensor_index, "시스템 오류")
                
            finally:
                if 'ser' in locals() and ser is not None: 
                    try:
                        ser.close() 
                    except Exception:
                        pass
                # 에러 발생 시 재시도 대기 시간
                self.safe_sleep(3000)

# ==========================================
#  [CO2 레벨 표시 위젯] (기존 동일)
# ==========================================
class CO2LevelWidget(QWidget):
    LEVELS = [
        {"name": "좋음", "min": 1, "max": 500, "color": "#28A745"},        
        {"name": "보통", "min": 501, "max": 1000, "color": "#FFD700"},        
        {"name": "민감군", "min": 1001, "max": 3000, "color": "#FD7E14"},     
        {"name": "나쁨", "min": 3001, "max": 5000, "color": "#DC3545"},      
        {"name": "매우 나쁨", "min": 5001, "max": 10000, "color": "#800080"}, 
        {"name": "위험", "min": 10001, "max": 30000, "color": "#795548"},      
    ]

    def __init__(self, title="센서", parent=None):
        super().__init__(parent)
        self.title = title
        self.current_value = 400 
        self.initUI()

    def initUI(self):
        self.setStyleSheet("background-color: white; border-radius: 10px;")
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
        main_layout.addStretch(1)

    def set_port_name(self, port_name):
        self.title_label.setText(self.title)

    def update_co2(self, value):
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
        arrow_x = target_x - (self.arrow_label.width() / 2)
        arrow_y = self.arrow_container.height() - self.arrow_label.height()

        self.arrow_label.move(QPoint(int(arrow_x), int(arrow_y)))

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
        self.resize(780, 320) 
        self.setStyleSheet("QMainWindow { background-color: white; border: none; }") 

        self.setWindowFlags(Qt.FramelessWindowHint)

        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QHBoxLayout(central_widget)
        main_layout.setContentsMargins(15, 15, 15, 15)
        main_layout.setSpacing(15) 

        self.widgets = []
        for i, port in enumerate(self.target_ports):
            # 타이틀에 할당된 포트 번호도 작게 표시하여 헷갈리지 않게 함
            widget = CO2LevelWidget(title=f"센서 {i+1} ({port})") 
            self.widgets.append(widget)
            main_layout.addWidget(widget)

    def start_threads(self):
        for i, port_name in enumerate(self.target_ports): 
            thread = SerialThread(port_name, i) 
            thread.data_signal.connect(self.update_data)
            thread.error_signal.connect(self.handle_error) 
            self.threads.append(thread)

        if len(self.threads) > 0:
            for i in range(1, len(self.threads)):
                self.threads[0].base_co2_signal.connect(self.threads[i].set_reference_co2)

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
            self.dragPos = event.globalPos()

    def mouseMoveEvent(self, event):
        if event.buttons() == Qt.LeftButton:
            self.move(self.pos() + event.globalPos() - self.dragPos)
            self.dragPos = event.globalPos()

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