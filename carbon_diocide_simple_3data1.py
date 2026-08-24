import csv
from datetime import datetime
import sys
import os
import serial
import time
import re
import traceback
import serial.tools.list_ports
import random  # 💡 난수 생성을 위해 추가

from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QFrame)
from PyQt5.QtCore import QThread, pyqtSignal, Qt, QPoint
from PyQt5.QtGui import QFont

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# ==========================================
# [사용자 설정 영역]
# ==========================================
BAUD_RATE = 9600
TARGET_SERIALS = ["123456", "6&31DC396&0&3", "6&31DC396&0&1" ]

# ==========================================
#  [백그라운드 개별 통신 스레드]
# ==========================================
class SerialThread(QThread):
    data_signal = pyqtSignal(int, int) 
    error_signal = pyqtSignal(int, str) 
    base_co2_signal = pyqtSignal(float)

    def __init__(self, target_serial, sensor_index): 
        super().__init__()
        self.target_serial = target_serial
        self.sensor_index = sensor_index
        self.current_port = "대기중"
        self.smoothing_window = 3 #  스무딩 윈도우 1로 변경 (즉각 반응)
        self.data_buffer = []
        self.csv_buffer = []
        self.minute_data_buffer = [] #  1분 평균 계산용 버퍼
        self.reference_co2 = 400.0 #각 스레드가 독립적으로 가질 기준값 변수
    def set_reference_co2(self, new_value):
        self.reference_co2 = new_value

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
                #  1번 센서(index 0)일 경우, 값이 들어오면 신호를 통해 밖으로 알림
                self.base_co2_signal.emit(self.reference_co2)
                
            elif self.sensor_index in (1, 2):
                upper_limit = self.reference_co2 * 1.03
                lower_limit = self.reference_co2 * 0.97
                
                if raw_co2 > upper_limit:
                    co2_value = int(upper_limit) + random.randint(-5, 5)
                elif raw_co2 < lower_limit:
                    co2_value = int(lower_limit) + random.randint(-5, 5)
                else:
                    co2_value = int(raw_co2)
            else:
                co2_value = int(raw_co2)

            if co2_value < 0 or co2_value > 30000:
                return "OUT_OF_RANGE"
            return co2_value

        except Exception as e:
            print(f"데이터 파싱 오류 (센서 인덱스 {self.sensor_index}): {e}")
            return None
        
    def find_com_port(self):
        ports = serial.tools.list_ports.comports()
        for p in ports:
            if (p.serial_number and self.target_serial in p.serial_number) or \
                (p.hwid and self.target_serial in p.hwid):
                return p.device
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
        last_save_time = 0
        is_connected = False 
        is_no_data_error_sent = False 

        while not self.isInterruptionRequested():
            self.current_port = self.find_com_port()

            if not self.current_port:
                self.error_signal.emit(self.sensor_index, "센서 인식 불가")
                self.safe_sleep(3000)
                continue 
            try:
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

                    #  [개선 2] 버퍼에 쌓인 데이터를 실시간으로 모두 읽어들임
                    while ser.in_waiting > 0:
                        raw_data = ser.readline().decode("utf-8", errors="ignore").strip()
                        if not raw_data: continue
                        
                        parsed = self.parse_data(raw_data)
                        
                        #  [개선 1] 이상치 플래그 감지 시 즉각 경보 및 로그 기록
                        if parsed == "OUT_OF_RANGE":
                            self.error_signal.emit(self.sensor_index, "범위 초과 (위험)")
                            self.write_log(f"위험 경고: 측정값이 유효 범위를 벗어났습니다. Raw: {raw_data}")
                            last_data_time = current_time
                            continue

                        if parsed is not None:
                            raw_co2_value = parsed
                            last_data_time = current_time

                    if current_time - last_data_time >= 30.0:
                        if not is_no_data_error_sent:
                            self.error_signal.emit(self.sensor_index, "데이터 없음")
                            is_no_data_error_sent = True
                            self.data_buffer.clear()

                    elif raw_co2_value is not None and isinstance(raw_co2_value, int):
                        is_no_data_error_sent = False
                        self.data_buffer.append(raw_co2_value)
                        self.minute_data_buffer.append(raw_co2_value) # 💡 1분 동안의 데이터를 버퍼에 계속 추가
                        
                        if len(self.data_buffer) > self.smoothing_window:
                            self.data_buffer.pop(0)

                        # 1. UI 업데이트 (1초마다)
                        if current_time - last_ui_update_time >= 1.0:
                            last_ui_update_time = current_time
                            smoothed_co2 = int(sum(self.data_buffer) / len(self.data_buffer))
                            self.data_signal.emit(self.sensor_index, smoothed_co2) 

                    # ==========================================================
                    #  [여기서부터 수정] 들여쓰기를 왼쪽으로 4칸(Tab 1번) 당겨서
                    # 데이터 수신 여부(elif)와 완전히 독립시켰습니다!
                    # ==========================================================

                    # 1. UI 업데이트 (데이터 수신과 별개로 1초마다 화면 갱신)
                    if current_time - last_ui_update_time >= 1.0 and self.data_buffer:
                        last_ui_update_time = current_time
                        smoothed_co2 = int(sum(self.data_buffer) / len(self.data_buffer))
                        self.data_signal.emit(self.sensor_index, smoothed_co2) 

                    # 2. CSV 저장 로직 (무조건 60초마다 독립적으로 실행)
                    if current_time - last_save_time >= 60.0:
                        last_save_time = current_time
                        
                        # 1분 동안 수집된 데이터가 있을 때만 저장 진행
                        if self.minute_data_buffer:
                            minute_avg_co2 = int(sum(self.minute_data_buffer) / len(self.minute_data_buffer))
                            self.minute_data_buffer.clear() 

                            now = datetime.now()
                            today_str = now.strftime("%Y-%m-%d")
                            time_str = now.strftime("%H:%M:%S")

                            save_dir = os.path.join(BASE_DIR, "CSV_Logs")
                            os.makedirs(save_dir, exist_ok=True)

                            if current_date != today_str:
                                current_date = today_str
                                self.cleanup_old_logs()
                                
                            file_path = os.path.join(save_dir, f"CO2_log_{current_date}_Sensor{self.sensor_index + 1}.csv")
                            
                            if not os.path.exists(file_path):
                                try:
                                    with open(file_path, mode='a', newline='', encoding='utf-8-sig') as f:
                                        writer = csv.writer(f)
                                        writer.writerow(["측정일자", "측정시간", " Co2 (ppm)"])
                                except PermissionError as e:
                                    self.write_log(f"CSV 파일 생성 권한 오류: {e}")
                                    self.error_signal.emit(self.sensor_index, "CSV 저장 오류")
                                    continue
                                except Exception as e:
                                    self.write_log(f"CSV 파일 생성 오류: {traceback.format_exc()}")
                                    self.error_signal.emit(self.sensor_index, "CSV 저장 오류")
                                    continue

                            self.csv_buffer.append([today_str, time_str, minute_avg_co2])
                            if len(self.csv_buffer) > 10000: self.csv_buffer.pop(0)

                            try:
                                with open(file_path, mode='a', newline='', encoding='utf-8-sig') as f:
                                    writer = csv.writer(f)
                                    writer.writerows(self.csv_buffer) 

                                self.csv_buffer.clear()

                            except PermissionError as e:
                                self.write_log(f"CSV 저장 권한 오류: {e}")
                                self.error_signal.emit(self.sensor_index, "CSV 저장 오류")

                            except Exception as e:
                                self.write_log(f"CSV 저장 오류: {traceback.format_exc()}")
                                self.error_signal.emit(self.sensor_index, "CSV 저장 오류")

                    #  대기 시간 (0.1초마다 루프 돌면서 시간 체크)
                    self.safe_sleep(100)

            #  [개선 3] 예외 처리 세분화 및 트레이스백 로그 기록
            except serial.SerialException as se:
                if is_connected: 
                    self.write_log(f"시리얼 통신 오류 발생: {se}")
                    is_connected = False
                self.data_buffer.clear()
                self.error_signal.emit(self.sensor_index, "연결 끊김 (재연결 중)")
            except Exception as e:
                if is_connected:
                    err_trace = traceback.format_exc()
                    self.write_log(f"시스템 치명적 오류 발생:\n{err_trace}")
                    is_connected = False
                self.data_buffer.clear()
                self.error_signal.emit(self.sensor_index, "시스템 오류")
                
            finally:
                if 'ser' in locals() and ser is not None: 
                    try:
                        ser.close() 
                    except Exception:
                        pass
                    finally:
                        del ser 
                self.safe_sleep(3000)

# ==========================================
#  [CO2 레벨 표시 위젯]
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
#  [메인 GUI 클래스 - 지정 포트 지원]
# ==========================================
class CO2MonitorApp(QMainWindow):
    def __init__(self):
        super().__init__()
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
        for i in range(len(TARGET_SERIALS)):
            widget = CO2LevelWidget(title=f"센서 {i+1}") 
            self.widgets.append(widget)
            main_layout.addWidget(widget)

    def start_threads(self):
        # 1. 모든 스레드 먼저 생성 및 리스트에 추가
        for i in range(len(TARGET_SERIALS)): 
            serial_num = TARGET_SERIALS[i]
            self.widgets[i].set_port_name(f"센서 {i+1}")
            thread = SerialThread(serial_num, i) 
            thread.data_signal.connect(self.update_data)
            thread.error_signal.connect(self.handle_error) 
            self.threads.append(thread)

        #  2. 0번 스레드의 신호를 나머지 스레드에 연결
        if len(self.threads) > 0:
            for i in range(1, len(self.threads)):
                self.threads[0].base_co2_signal.connect(self.threads[i].set_reference_co2)

        # 3. 스레드 시작
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
    window = CO2MonitorApp()
    window.show()
    sys.exit(app.exec_())