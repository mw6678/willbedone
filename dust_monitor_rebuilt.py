import csv
from datetime import datetime
import sys
import os
import serial
import time
import traceback
import serial.tools.list_ports

from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QFrame, QPushButton, QDialog, QMessageBox,
    QListWidget, QListWidgetItem, QScrollArea, QGroupBox
)
from PyQt5.QtCore import QThread, pyqtSignal, Qt, QPoint
from PyQt5.QtGui import QFont

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# ==========================================
# 1. 설정 관리 (Config)
# ==========================================
class Config:
    BAUD_RATE = 9600
    # 센서 포트별 [PM10 보정값, PM2.5 보정값, PM1.0 보정값]
    SENSOR_OFFSETS = {
        0: (0.0, 0.0, 0.0),
        1: (0.0, 0.0, 0.0),
        2: (0.0, 0.0, 0.0),
        3: (0.0, 0.0, 0.0),
    }
    SMOOTHING_WINDOW = 1
    NO_DATA_TIMEOUT = 30.0
    RECONNECT_DELAY_MS = 3000
    LOG_RETENTION_DAYS = 30

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

# ==========================================
# 2. 데이터 파싱 로직 (Data Parser)
# ==========================================
class DustParser:
    @staticmethod
    def parse(raw_data, offsets):
        try:
            raw_data = raw_data.strip()
            if not raw_data:
                return None

            parts = [item.strip() for item in raw_data.replace(" ", "").split(",")]
            if len(parts) < 3:
                return None

            # 튜플로 전달받은 3개의 보정값을 각각 변수에 분리
            offset_pm10, offset_pm25, offset_pm1 = offsets

            # 각각의 보정값을 개별적으로 적용
            pm10 = max(0, int(round(float(parts[2]) + offset_pm10)))
            pm25 = max(0, int(round(float(parts[1]) + offset_pm25)))
            pm1 = max(0, int(round(float(parts[0]) + offset_pm1)))

            if pm1 > 1000 or pm25 > 1000 or pm10 > 2000:
                return "OUT_OF_RANGE"

            return pm10, pm25, pm1
        except Exception:
            return None

# ==========================================
# 3. 로깅 및 파일 처리 (Sensor Logger)
# ==========================================
class SensorLogger:
    def __init__(self, port_name, sensor_index):
        self.port_name = port_name
        self.sensor_index = sensor_index
        self.log_dir = os.path.join(BASE_DIR, "Logs")
        self.csv_dir = os.path.join(BASE_DIR, "CSV_Logs")
        os.makedirs(self.log_dir, exist_ok=True)
        os.makedirs(self.csv_dir, exist_ok=True)

    def write_error(self, message):
        log_file = os.path.join(self.log_dir, f"error_log_Sensor{self.sensor_index + 1}.txt")
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        try:
            with open(log_file, "a", encoding="utf-8") as f:
                f.write(f"{timestamp} | 포트: {self.port_name} | {message}\n")
        except Exception:
            pass

    def save_minute_data(self, minute_key, buffers):
        if not buffers["pm1"]:
            return

        try:
            avg_pm10 = round(sum(buffers["pm10"]) / len(buffers["pm10"]), 1)
            avg_pm25 = round(sum(buffers["pm25"]) / len(buffers["pm25"]), 1)
            avg_pm1 = round(sum(buffers["pm1"]) / len(buffers["pm1"]), 1)

            saved_dt = datetime.strptime(minute_key, "%Y-%m-%d %H:%M")
            date_str = saved_dt.strftime("%Y-%m-%d")
            time_str = saved_dt.strftime("%H:%M:00")

            file_path = os.path.join(self.csv_dir, f"Dust_log_{date_str}_Sensor{self.sensor_index + 1}.csv")
            file_exists = os.path.exists(file_path)

            with open(file_path, mode="a", newline="", encoding="utf-8-sig") as f:
                writer = csv.writer(f)
                if not file_exists:
                    writer.writerow(["측정일자", "측정시간", "포트", "PM10_평균", "PM2.5_평균", "PM1.0_평균"])
                writer.writerow([date_str, time_str, self.port_name, avg_pm10, avg_pm25, avg_pm1])

            self.cleanup_old_logs()
        except Exception as e:
            self.write_error(f"CSV 저장 오류: {e}")

    def cleanup_old_logs(self):
        threshold_time = time.time() - (Config.LOG_RETENTION_DAYS * 24 * 60 * 60)
        for target_dir in [self.log_dir, self.csv_dir]:
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

# ==========================================
# 4. 시리얼 통신 스레드 (Serial Thread)
# ==========================================
class SerialThread(QThread):
    data_signal = pyqtSignal(str, int, int, int)
    error_signal = pyqtSignal(str, str)

    def __init__(self, port_name, sensor_index):
        super().__init__()
        self.port_name = port_name
        
        # 단일 값이 아닌 튜플(3개 값)을 통째로 가져옴 (기본값도 튜플로)
        self.offsets = Config.SENSOR_OFFSETS.get(sensor_index, (0.0, 0.0, 0.0))
        
        self.logger = SensorLogger(port_name, sensor_index)
        self.data_buffer = {"pm10": [], "pm25": [], "pm1": []}
        self.minute_data_buffer = {"pm10": [], "pm25": [], "pm1": []}
        self.last_known_values = (0, 0, 0)

    def safe_sleep(self, ms):
        elapsed = 0
        while elapsed < ms:
            if self.isInterruptionRequested():
                return
            sleep_time = min(100, ms - elapsed)
            QThread.msleep(sleep_time)
            elapsed += sleep_time

    def clear_buffers(self):
        for key in self.data_buffer:
            self.data_buffer[key].clear()
            self.minute_data_buffer[key].clear()

    def run(self):
        is_connected = False
        no_data_error_sent = False
        last_saved_minute_key = datetime.now().strftime("%Y-%m-%d %H:%M")
        last_ui_update_time = 0.0
        last_data_time = time.time()

        while not self.isInterruptionRequested():
            ser = None
            try:
                ser = serial.Serial(self.port_name, Config.BAUD_RATE, timeout=0.1)
                ser.reset_input_buffer()

                if not is_connected:
                    self.logger.write_error(f"통신 연결 성공 ({self.port_name})")

                is_connected = True
                no_data_error_sent = False
                last_data_time = time.time()

                while ser.is_open and not self.isInterruptionRequested():
                    raw_values = None
                    
                    # 1. 데이터 수신 및 파싱
                    while ser.in_waiting > 0:
                        raw_data = ser.readline().decode("utf-8", errors="ignore").strip()
                        if not raw_data: continue

                        parsed = DustParser.parse(raw_data, self.offsets)
                        if parsed == "OUT_OF_RANGE":
                            self.error_signal.emit(self.port_name, "측정값 범위 초과")
                            self.logger.write_error(f"범위 초과 Raw: {raw_data}")
                            continue
                        elif parsed is not None:
                            raw_values = parsed
                            last_data_time = time.time()

                    # 2. 타임아웃 체크
                    if time.time() - last_data_time >= Config.NO_DATA_TIMEOUT:
                        if not no_data_error_sent:
                            self.error_signal.emit(self.port_name, "데이터 없음")
                            self.logger.write_error("30초 이상 데이터 없음")
                            no_data_error_sent = True
                            self.clear_buffers()

                    # 3. 정상 데이터 버퍼링 및 UI 업데이트
                    elif raw_values is not None:
                        no_data_error_sent = False
                        pm10, pm25, pm1 = raw_values

                        for key, val in zip(["pm10", "pm25", "pm1"], raw_values):
                            self.data_buffer[key].append(val)
                            self.minute_data_buffer[key].append(val)
                            if len(self.data_buffer[key]) > Config.SMOOTHING_WINDOW:
                                self.data_buffer[key].pop(0)

                        if time.time() - last_ui_update_time >= 1.0:
                            last_ui_update_time = time.time()
                            avgs = [int(sum(self.data_buffer[k]) / len(self.data_buffer[k])) for k in ["pm10", "pm25", "pm1"]]
                            self.last_known_values = tuple(avgs)
                            self.data_signal.emit(self.port_name, *avgs)

                    # 4. 분 단위 CSV 저장
                    current_minute_key = datetime.now().strftime("%Y-%m-%d %H:%M")
                    if current_minute_key != last_saved_minute_key:
                        self.logger.save_minute_data(last_saved_minute_key, self.minute_data_buffer)
                        for key in self.minute_data_buffer:
                            self.minute_data_buffer[key].clear()
                        last_saved_minute_key = current_minute_key

                    self.safe_sleep(100)

            except serial.SerialException as e:
                if is_connected: self.logger.write_error(f"시리얼 통신 오류: {e}")
                is_connected = False
                self.clear_buffers()
                self.error_signal.emit(self.port_name, "연결 실패/끊김")
            except Exception:
                if is_connected: self.logger.write_error("시스템 오류:\n" + traceback.format_exc())
                is_connected = False
                self.clear_buffers()
                self.error_signal.emit(self.port_name, "시스템 오류")
            finally:
                if ser is not None:
                    try:
                        if ser.is_open: ser.close()
                    except Exception: pass
                if not self.isInterruptionRequested():
                    self.safe_sleep(Config.RECONNECT_DELAY_MS)


# ==========================================
# 5. UI 위젯 및 앱 메인 로직 (기존과 동일)
# ==========================================
class PortSelectionDialog(QDialog):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("미세먼지 센서 포트 선택")
        self.resize(350, 400)
        self.setStyleSheet("background-color: white;")
        self.selected_ports = []
        self.initUI()
        self.refresh_ports()

    def initUI(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(12)

        title = QLabel("연결할 COM 포트를 모두 선택하세요")
        title.setFont(QFont("Malgun Gothic", 11, QFont.Bold))
        title.setAlignment(Qt.AlignCenter)
        layout.addWidget(title)

        self.list_widget = QListWidget()
        layout.addWidget(self.list_widget)

        refresh_btn = QPushButton("포트 새로고침")
        refresh_btn.clicked.connect(self.refresh_ports)
        layout.addWidget(refresh_btn)

        self.start_btn = QPushButton("모니터링 시작")
        self.start_btn.setFixedHeight(40)
        self.start_btn.setFont(QFont("Malgun Gothic", 10, QFont.Bold))
        self.start_btn.setStyleSheet("background-color: #007BFF; color: white; border-radius: 5px;")
        self.start_btn.clicked.connect(self.on_start_clicked)
        layout.addWidget(self.start_btn)

    def refresh_ports(self):
        self.list_widget.clear()
        ports = [p.device for p in serial.tools.list_ports.comports()]
        if not ports:
            item = QListWidgetItem("연결된 포트 없음")
            item.setFlags(Qt.NoItemFlags)
            self.list_widget.addItem(item)
            return

        for port in ports:
            item = QListWidgetItem(port)
            item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
            item.setCheckState(Qt.Unchecked)
            self.list_widget.addItem(item)

    def on_start_clicked(self):
        self.selected_ports = [self.list_widget.item(i).text() for i in range(self.list_widget.count()) if self.list_widget.item(i).checkState() == Qt.Checked]
        if not self.selected_ports:
            QMessageBox.warning(self, "경고", "최소 하나 이상의 포트를 선택해 주세요.")
            return
        if len(self.selected_ports) != len(set(self.selected_ports)):
            QMessageBox.warning(self, "경고", "같은 COM 포트를 중복 선택할 수 없습니다.")
            return
        self.accept()

class DustLevelWidget(QWidget):
    def __init__(self, title="미세먼지", levels=None, parent=None):
        super().__init__(parent)
        self.title = title
        self.levels = levels if levels else Config.PM10_LEVELS
        self.current_value = 0
        self.initUI()

    def initUI(self):
        self.setStyleSheet("background-color: white; border-radius: 10px; border: 1px solid #E0E0E0;")
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(10, 10, 10, 10)
        main_layout.setSpacing(3)

        self.title_label = QLabel(self.title)
        self.title_label.setFont(QFont("Malgun Gothic", 11, QFont.Bold))
        self.title_label.setAlignment(Qt.AlignCenter)
        self.title_label.setStyleSheet("color: black; border: none;")
        main_layout.addWidget(self.title_label)

        value_layout = QHBoxLayout()
        value_layout.setAlignment(Qt.AlignCenter)

        self.val_label = QLabel("----")
        self.val_label.setFont(QFont("Arial", 28, QFont.Bold))
        self.val_label.setStyleSheet("color: black; border: none;")
        value_layout.addWidget(self.val_label)

        unit_label = QLabel("µg/m³")
        unit_label.setFont(QFont("Arial", 10, QFont.Bold))
        unit_label.setStyleSheet("color: #666666; margin-bottom: 4px; border: none;")
        unit_label.setAlignment(Qt.AlignBottom)
        value_layout.addWidget(unit_label)

        main_layout.addLayout(value_layout)

        self.arrow_container = QWidget()
        self.arrow_container.setStyleSheet("border: none;")
        self.arrow_container.setFixedHeight(14)

        self.arrow_label = QLabel("▼")
        self.arrow_label.setFont(QFont("Arial", 9, QFont.Bold))
        self.arrow_label.setStyleSheet("color: black; border: none;")
        self.arrow_label.setAlignment(Qt.AlignCenter)
        self.arrow_label.setFixedWidth(14)
        self.arrow_label.setParent(self.arrow_container)
        main_layout.addWidget(self.arrow_container)

        level_bar_frame = QFrame()
        level_bar_frame.setStyleSheet("border: none;")
        self.level_bar_layout = QHBoxLayout(level_bar_frame)
        self.level_bar_layout.setContentsMargins(0, 0, 0, 0)
        self.level_bar_layout.setSpacing(2)

        self.level_bars = []
        for i, level in enumerate(self.levels):
            bar = QFrame()
            bar.setFixedHeight(10)
            border_radius = ""
            if i == 0: border_radius = "border-top-left-radius: 5px; border-bottom-left-radius: 5px;"
            elif i == len(self.levels) - 1: border_radius = "border-top-right-radius: 5px; border-bottom-right-radius: 5px;"
            bar.setStyleSheet(f"background-color: {level['color']}; {border_radius}")
            self.level_bars.append(bar)
            self.level_bar_layout.addWidget(bar)
        main_layout.addWidget(level_bar_frame)

        self.status_text_label = QLabel("대기 중...")
        self.status_text_label.setFont(QFont("Malgun Gothic", 11, QFont.Bold))
        self.status_text_label.setAlignment(Qt.AlignCenter)
        self.status_text_label.setFixedHeight(30)
        self.status_text_label.setStyleSheet("background-color: #E0E0E0; border-radius: 6px; color: gray;")
        main_layout.addWidget(self.status_text_label)

    def update_val(self, value):
        self.arrow_label.show()
        self.current_value = value
        self.val_label.setText(str(value))

        current_level = self.levels[0] if value < self.levels[0]["min"] else self.levels[-1]
        for level in self.levels:
            if value <= level["max"]:
                current_level = level
                break

        self.status_text_label.setText(current_level["name"])
        text_color = "black" if current_level["name"] == "보통" else "white"
        self.status_text_label.setStyleSheet(f"background-color: {current_level['color']}; border-radius: 6px; color: {text_color}; border: none;")
        self.update_arrow_position()

    def set_error_state(self, msg):
        self.val_label.setText("----")
        self.status_text_label.setText(msg)
        self.status_text_label.setStyleSheet("background-color: #FFCDD2; border-radius: 6px; color: #B71C1C; border: none;")
        self.current_value = self.levels[0]["min"]
        self.update_arrow_position()
        self.arrow_label.hide()

    def update_arrow_position(self):
        if not self.level_bars or self.level_bars[0].geometry().width() == 0:
            return

        value = self.current_value
        level_index = len(self.levels) - 1
        for i, level in enumerate(self.levels):
            if value <= level["max"]:
                level_index = i
                break

        current_level = self.levels[level_index]
        level_min, level_max = current_level["min"], current_level["max"]

        if value <= level_min: ratio = 0.0
        elif value >= level_max: ratio = 1.0
        else: ratio = (value - level_min) / (level_max - level_min)

        target_bar = self.level_bars[level_index]
        target_x = target_bar.geometry().x() + (target_bar.geometry().width() * ratio)
        
        self.arrow_label.move(QPoint(int(target_x - (self.arrow_label.width() / 2)), int(self.arrow_container.height() - self.arrow_label.height() + 2)))

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self.update_arrow_position()

class DustMonitorApp(QMainWindow):
    def __init__(self, ports):
        super().__init__()
        self.ports = ports
        self.threads = []
        self.port_widgets = {}
        self.initUI()
        self.start_threads()

    def initUI(self):
        self.setWindowTitle("다중 미세먼지 모니터링 시스템")
        self.resize(1000, 700)
        self.setStyleSheet("QMainWindow { background-color: white; }")

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        self.setCentralWidget(scroll)

        container = QWidget()
        scroll.setWidget(container)
        main_layout = QVBoxLayout(container)
        main_layout.setContentsMargins(15, 15, 15, 15)
        main_layout.setSpacing(15)

        for port in self.ports:
            group_box = QGroupBox(f"센서 포트: {port}")
            group_box.setFont(QFont("Malgun Gothic", 11, QFont.Bold))
            port_layout = QHBoxLayout(group_box)
            port_layout.setSpacing(10)

            w_pm10 = DustLevelWidget("PM10 (미세먼지)", Config.PM10_LEVELS)
            w_pm25 = DustLevelWidget("PM2.5 (초미세먼지)", Config.PM25_LEVELS)
            w_pm1 = DustLevelWidget("PM1.0 (극미세먼지)", Config.PM1_LEVELS)

            port_layout.addWidget(w_pm10)
            port_layout.addWidget(w_pm25)
            port_layout.addWidget(w_pm1)
            main_layout.addWidget(group_box)

            self.port_widgets[port] = {"pm10": w_pm10, "pm25": w_pm25, "pm1": w_pm1}

        main_layout.addStretch(1)

    def start_threads(self):
        for index, port in enumerate(self.ports):
            thread = SerialThread(port, sensor_index=index)
            thread.data_signal.connect(self.update_data)
            thread.error_signal.connect(self.handle_error)
            self.threads.append(thread)
            thread.start()

    def update_data(self, port, pm10, pm25, pm1):
        widgets = self.port_widgets.get(port)
        if widgets:
            widgets["pm10"].update_val(pm10)
            widgets["pm25"].update_val(pm25)
            widgets["pm1"].update_val(pm1)

    def handle_error(self, port, error_msg):
        widgets = self.port_widgets.get(port)
        if widgets:
            widgets["pm10"].set_error_state(error_msg)
            widgets["pm25"].set_error_state(error_msg)
            widgets["pm1"].set_error_state(error_msg)

    def closeEvent(self, event):
        for thread in self.threads: thread.requestInterruption()
        for thread in self.threads: thread.wait(1500)
        event.accept()

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton: self.dragPos = event.globalPos()

    def mouseMoveEvent(self, event):
        if event.buttons() == Qt.LeftButton and hasattr(self, "dragPos"):
            self.move(self.pos() + event.globalPos() - self.dragPos)
            self.dragPos = event.globalPos()

    def keyPressEvent(self, event):
        if event.key() == Qt.Key_Escape: self.close()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    setup_dialog = PortSelectionDialog()
    if setup_dialog.exec_() == QDialog.Accepted:
        window = DustMonitorApp(setup_dialog.selected_ports)
        window.show()
        sys.exit(app.exec_())