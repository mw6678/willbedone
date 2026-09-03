import csv
import sqlite3
from datetime import datetime
import sys
import os
import serial
import time
import traceback
import serial.tools.list_ports
from collections import deque
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QFrame, QPushButton, QDialog, QMessageBox,
    QListWidget, QListWidgetItem, QScrollArea, QGroupBox, QComboBox
)
from PyQt5.QtCore import QThread, pyqtSignal, Qt, QPoint
from PyQt5.QtGui import QFont
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# ==========================================
# 1. 설정 관리 (Config)
# ==========================================
class Config:
    BAUD_RATE = 9600
    SENSOR_CALIBRATION = {
        0: {"scale": (1.0, 2.0, 1.0), "offset": (0.0, 1.0, 0.0)},
        1: {"scale": (1.0, 3.0, 1.0), "offset": (0.0, 1.0, 0.0)},
        2: {"scale": (1.0, 1.0, 1.0), "offset": (0.0, 0.0, 0.0)},
        3: {"scale": (1.0, 1.0, 1.0), "offset": (0.0, 0.0, 0.0)},
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
    def parse(raw_data, calib_params):
        try:
            raw_data = raw_data.strip()
            if not raw_data:
                return None

            parts = [
                item.strip() for item in raw_data.replace(" ", "").split(",")
            ]
            if len(parts) < 3:
                return None

            scales = calib_params.get("scale", (1.0, 1.0, 1.0))
            offsets = calib_params.get("offset", (0.0, 0.0, 0.0))

            scale_pm10, scale_pm25, scale_pm1 = scales
            offset_pm10, offset_pm25, offset_pm1 = offsets

            raw_pm10 = float(parts[2])
            raw_pm25 = float(parts[1])
            raw_pm1 = float(parts[0])

            pm10 = max(0, int(round((raw_pm10 + offset_pm10) * scale_pm10)))
            pm25 = max(0, int(round((raw_pm25 + offset_pm25) * scale_pm25)))
            pm1 = max(0, int(round((raw_pm1 + offset_pm1) * scale_pm1)))

            if pm1 > 1000 or pm25 > 1000 or pm10 > 2000:
                return "OUT_OF_RANGE"

            return {
                "pm10": pm10, 
                "pm25": pm25, 
                "pm1": pm1
            }
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
        self.excel_dir = os.path.join(BASE_DIR, "Excel_Logs")
        self.db_dir = os.path.join(BASE_DIR, "Data")
        
        os.makedirs(self.log_dir, exist_ok=True)
        os.makedirs(self.excel_dir, exist_ok=True)
        os.makedirs(self.db_dir, exist_ok=True)

        self.db_path = os.path.join(self.db_dir, "dust_measurement.db")

    @staticmethod
    def init_global_database(db_path):
        try:
            with sqlite3.connect(db_path) as conn:
                conn.execute("PRAGMA journal_mode=WAL;")
                conn.execute("""
                    CREATE TABLE IF NOT EXISTS measurements (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        measured_at TEXT NOT NULL,
                        sensor_index INTEGER NOT NULL,
                        port TEXT NOT NULL,
                        pm10 REAL,
                        pm25 REAL,
                        pm1 REAL,
                        status TEXT DEFAULT 'NORMAL'
                    )
                """)
                conn.execute("CREATE INDEX IF NOT EXISTS idx_measurements_time ON measurements(measured_at)")
                conn.execute("CREATE INDEX IF NOT EXISTS idx_measurements_sensor ON measurements(sensor_index)")
                conn.commit()
        except Exception as e:
            print(f"Global SQLite 초기화 오류: {e}")

    @staticmethod
    def cleanup_old_logs_global():
        threshold_time = time.time() - (Config.LOG_RETENTION_DAYS * 24 * 60 * 60)
        for target_dir in [os.path.join(BASE_DIR, "Logs"), os.path.join(BASE_DIR, "Excel_Logs")]:
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

    def write_error(self, message):
        log_file = os.path.join(self.log_dir, f"error_log_Sensor{self.sensor_index + 1}.txt")
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        try:
            with open(log_file, "a", encoding="utf-8-sig") as f:
                f.write(f"{timestamp} | 포트: {self.port_name} | {message}\n")
        except Exception:
            pass

    def save_measurements_batch(self, measurements):
        if not measurements:
            return True
        try:
            with sqlite3.connect(self.db_path) as conn:
                conn.executemany("""
                    INSERT INTO measurements (
                        measured_at, sensor_index, port,
                        pm10, pm25, pm1, status
                    ) VALUES (?, ?, ?, ?, ?, ?, ?)
                """, measurements)
                conn.commit()
            return True
        except Exception as e:
            self.write_error(f"SQLite 일괄 저장 오류: {e}")
            return False

    def save_measurement(self, measured_at, pm10, pm25, pm1, status="NORMAL"):
        try:
            with sqlite3.connect(self.db_path) as conn:
                conn.execute("""
                    INSERT INTO measurements (
                        measured_at, sensor_index, port,
                        pm10, pm25, pm1, status
                    ) VALUES (?, ?, ?, ?, ?, ?, ?)
                """, (measured_at, self.sensor_index, self.port_name, pm10, pm25, pm1, status))
                conn.commit()
            return True
        except Exception as e:
            self.write_error(f"SQLite 측정 데이터 저장 오류: {e}")
            return False

    def export_excel(self, date_str=None):
        try:
            if not date_str:
                date_str = datetime.now().strftime("%Y-%m-%d")

            with sqlite3.connect(self.db_path) as conn:
                rows = conn.execute("""
                    SELECT measured_at, port, pm10, pm25, pm1, status
                    FROM measurements 
                    WHERE sensor_index = ? AND measured_at LIKE ? 
                    ORDER BY measured_at
                """, (self.sensor_index, f"{date_str}%")).fetchall()

            if not rows:
                return False

            filename = f"Dust_log_{date_str}_Sensor{self.sensor_index + 1}.xlsx"
            file_path = os.path.join(self.excel_dir, filename)

            wb = Workbook()
            ws = wb.active
            ws.title = f"{date_str} 측정 데이터"

            headers = ["측정일시", "포트", "PM10", "PM2.5", "PM1.0", "상태"]
            ws.append(headers)
            ws.row_dimensions[1].height = 25

            thin_border = Border(
                left=Side(style='thin', color='D3D3D3'), right=Side(style='thin', color='D3D3D3'),
                top=Side(style='thin', color='D3D3D3'), bottom=Side(style='thin', color='D3D3D3')
            )
            header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
            header_font = Font(name="Malgun Gothic", size=10, bold=True, color="FFFFFF")
            data_font = Font(name="Malgun Gothic", size=9)
            center_align = Alignment(horizontal="center", vertical="center")
            right_align = Alignment(horizontal="right", vertical="center")

            for col_num, header in enumerate(headers, 1):
                cell = ws.cell(row=1, column=col_num)
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = center_align
                cell.border = thin_border

            for row_data in rows:
                ws.append(list(row_data))
                current_row = ws.max_row
                ws.row_dimensions[current_row].height = 18

                for col_num, val in enumerate(row_data, 1):
                    cell = ws.cell(row=current_row, column=col_num)
                    cell.font = data_font
                    cell.border = thin_border

                    if col_num in [1, 2, 6]:
                        cell.alignment = center_align
                    else:
                        cell.alignment = right_align
                        if isinstance(val, (int, float)):
                            cell.number_format = '#,##0'

            for col in ws.columns:
                max_len = 0
                col_letter = get_column_letter(col[0].column)
                for cell in col:
                    val_str = str(cell.value or '')
                    max_len = max(max_len, len(val_str.encode('utf-8')) if cell.row == 1 else len(val_str))
                ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

            wb.save(file_path)
            return True

        except Exception as e:
            self.write_error(f"Excel Export 오류: {e}")
            return False

# ==========================================
# 4. 시리얼 통신 스레드 (Serial Thread)
# ==========================================
class SerialThread(QThread):
    data_signal = pyqtSignal(int, int, int, int)
    error_signal = pyqtSignal(int, str)

    def __init__(self, port_name, sensor_index):
        super().__init__()
        self.port_name = port_name
        self.sensor_index = sensor_index

        default_calib = {"scale": (1.0, 1.0, 1.0), "offset": (0.0, 0.0, 0.0)}
        self.calib_params = Config.SENSOR_CALIBRATION.get(
            sensor_index, default_calib
        )

        self.logger = SensorLogger(port_name, sensor_index)
        self.data_buffer = {
            "pm10": deque(maxlen=Config.SMOOTHING_WINDOW), 
            "pm25": deque(maxlen=Config.SMOOTHING_WINDOW), 
            "pm1": deque(maxlen=Config.SMOOTHING_WINDOW)
        }
        # 리스트 대신 합계와 개수로 1분 평균 데이터 관리
        self.minute_sum = {"pm10": 0, "pm25": 0, "pm1": 0}
        self.minute_count = 0
        
        self.last_minute = datetime.now().minute       
        # 1분 주기 평균 저장을 위한 시간 변수 추가
        self.last_minute_save_time = time.time()

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
        
        # 합계와 카운트 초기화로 변경
        self.minute_sum = {"pm10": 0, "pm25": 0, "pm1": 0}
        self.minute_count = 0

    def save_current_minute_average(self):
        # 데이터가 1건이라도 쌓였다면
        if self.minute_count > 0:
            avg_pm10 = int(self.minute_sum["pm10"] / self.minute_count)
            avg_pm25 = int(self.minute_sum["pm25"] / self.minute_count)
            avg_pm1 = int(self.minute_sum["pm1"] / self.minute_count)
            
            measured_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            
            batch_data = [(
                measured_at, self.sensor_index, self.port_name,
                avg_pm10, avg_pm25, avg_pm1, "NORMAL"
            )]
            self.logger.save_measurements_batch(batch_data)
            
            # 저장 후 합계와 카운트 초기화
            self.minute_sum = {"pm10": 0, "pm25": 0, "pm1": 0}
            self.minute_count = 0
            
        self.last_minute_save_time = time.time()

    def run(self):
        is_connected = False
        no_data_error_sent = False
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
                self.last_minute_save_time = time.time() # 연결 시점 초기화

                while ser.is_open and not self.isInterruptionRequested():
                    parsed_values = []
                    while ser.in_waiting > 0:
                        raw_data = ser.readline().decode("utf-8", errors="ignore").strip()
                        if not raw_data:
                            continue

                        parsed = DustParser.parse(raw_data, self.calib_params)
                        if parsed == "OUT_OF_RANGE":
                            self.error_signal.emit(self.sensor_index, "측정값 범위 초과")
                            self.logger.write_error(f"범위 초과 Raw: {raw_data}")
                            continue
                        elif parsed is not None:
                            parsed_values.append(parsed)
                            last_data_time = time.time()

                    if time.time() - last_data_time >= Config.NO_DATA_TIMEOUT:
                        if not no_data_error_sent:
                            self.error_signal.emit(self.sensor_index, "데이터 없음")
                            self.logger.write_error("30초 이상 데이터 없음")
                            no_data_error_sent = True
                            self.clear_buffers()

                    elif parsed_values:
                        no_data_error_sent = False

                        for parsed in parsed_values:
                            pm10 = parsed["pm10"]
                            pm25 = parsed["pm25"]
                            pm1 = parsed["pm1"]
                            
                            # 실시간 화면 표시용 버퍼 (기존 유지)
                            for parsed in parsed_values:
                                pm10 = parsed["pm10"]
                                pm25 = parsed["pm25"]
                                pm1 = parsed["pm1"]
                            
                            # 1. 화면 표시용 덱(deque) 버퍼
                            for key, val in zip(["pm10", "pm25", "pm1"], (pm10, pm25, pm1)):
                                self.data_buffer[key].append(val)

                                # 1분 평균용 합계 및 개수 누적
                                self.minute_sum["pm10"] += pm10
                                self.minute_sum["pm25"] += pm25
                                self.minute_sum["pm1"] += pm1
                                self.minute_count += 1

                        # 시스템 시각의 '분'이 바뀌었는지 체크
                        current_minute = datetime.now().minute
                        if current_minute != self.last_minute:
                            self.save_current_minute_average()
                            self.last_minute = current_minute

                        # 화면 UI는 1초마다 실시간 갱신
                        if time.time() - last_ui_update_time >= 1.0:
                            last_ui_update_time = time.time()
                            avgs = [
                                int(sum(self.data_buffer[k]) / len(self.data_buffer[k]))
                                for k in ["pm10", "pm25", "pm1"]
                            ]
                            self.data_signal.emit(self.sensor_index, *avgs)

                    self.safe_sleep(100)

            except serial.SerialException as e:
                if is_connected:
                    self.logger.write_error(f"시리얼 통신 오류: {e}")
                is_connected = False
                self.clear_buffers()
                self.error_signal.emit(self.sensor_index, "연결 실패/끊김")
            except Exception:
                if is_connected:
                    self.logger.write_error("시스템 오류:\n" + traceback.format_exc())
                is_connected = False
                self.clear_buffers()
                self.error_signal.emit(self.sensor_index, "시스템 오류")
            finally:
                if ser is not None:
                    try:
                        if ser.is_open:
                            ser.close()
                    except Exception:
                        pass
                if not self.isInterruptionRequested():
                    self.safe_sleep(Config.RECONNECT_DELAY_MS)

# ==========================================
# 5. UI 위젯 및 앱 메인 로직
# ==========================================
class DustLevelWidget(QWidget):
    def __init__(self, title="미세먼지", levels=None, parent=None):
        super().__init__(parent)
        self.title = title
        self.levels = levels if levels else Config.PM10_LEVELS
        self.current_value = 0
        self.initUI()

    def initUI(self):
        self.setStyleSheet("background-color: transparent; border: none;")        
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

        current_level = self.levels[self._find_level_index(value)]

        self.status_text_label.setText(current_level["name"])
        text_color = "black" if current_level["name"] == "보통" else "white"
        self.status_text_label.setStyleSheet(f"background-color: {current_level['color']}; border-radius: 6px; color: {text_color}; border: none;")
        self.update_arrow_position()

    def _find_level_index(self, value):
        for i, level in enumerate(self.levels):
            if value <= level["max"]:
                return i
        return len(self.levels) - 1

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
        level_index = self._find_level_index(value)

        current_level = self.levels[level_index]
        level_min, level_max = current_level["min"], current_level["max"]

        if value <= level_min: ratio = 0.0
        elif value >= level_max: ratio = 1.0
        else: ratio = (value - level_min) / (level_max - level_min)

        target_bar = self.level_bars[level_index]
        target_x = target_bar.geometry().x() + (target_bar.geometry().width() * ratio)
        arrow_x = int(target_x - (self.arrow_label.width() / 2))
        max_x = self.arrow_container.width() - self.arrow_label.width()
        arrow_x = max(0, min(arrow_x, max_x))
        self.arrow_label.move(QPoint(arrow_x, int(self.arrow_container.height() - self.arrow_label.height() + 2)))

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self.update_arrow_position()

# ==========================================
# 6. 포트 선택 팝업 다이얼로그
# ==========================================
class PortSelectionDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.selected_ports = {}
        self.available_ports = [port.device for port in serial.tools.list_ports.comports()]
        self.initUI()

    def initUI(self):
        self.setWindowTitle("센서 포트 설정")
        self.resize(350, 250)
        self.setStyleSheet("background-color: white;")

        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(15)

        title_label = QLabel("모니터링할 센서 포트를 선택하세요")
        title_label.setFont(QFont("Malgun Gothic", 10, QFont.Bold))
        layout.addWidget(title_label)

        self.combos = []
        for i in range(4):
            h_layout = QHBoxLayout()
            label = QLabel(f"센서 {i+1} 포트:")
            label.setFont(QFont("Malgun Gothic", 9))
            
            combo = QComboBox()
            combo.addItem("선택 안 함")
            combo.addItems(self.available_ports)
            combo.setFont(QFont("Malgun Gothic", 9))
            # 포트 선택이 바뀔 때마다 다른 콤보박스 목록을 갱신하도록 연결
            combo.currentIndexChanged.connect(self.update_combo_items)
            
            h_layout.addWidget(label)
            h_layout.addWidget(combo)
            layout.addLayout(h_layout)
            self.combos.append(combo)

        self.start_btn = QPushButton("모니터링 시작")
        self.start_btn.setFixedHeight(35)
        self.start_btn.setFont(QFont("Malgun Gothic", 10, QFont.Bold))
        self.start_btn.setStyleSheet("background-color: #007BFF; color: white; border-radius: 4px;")
        self.start_btn.clicked.connect(self.accept_selection)
        layout.addWidget(self.start_btn)

    def update_combo_items(self):
        # 현재 선택되어 있는 포트 목록 수집
        selected_values = [
            combo.currentText() for combo in self.combos 
            if combo.currentText() and combo.currentText() != "선택 안 함"
        ]

        # 각 콤보박스의 선택 상태를 유지하면서 이미 다른 곳에서 선택된 포트는 목록에서 숨기거나 막기
        for combo in self.combos:
            current_selection = combo.currentText()
            combo.blockSignals(True) # 시그널 루프 방지
            combo.clear()
            combo.addItem("선택 안 함")
            
            for port in self.available_ports:
                # 다른 콤보박스에서 이미 골랐고, 현재 콤보박스에서 고른 게 아니라면 추가 안 함
                if port in selected_values and port != current_selection:
                    continue
                combo.addItem(port)
                
            # 기존에 선택했던 값 복구 시도
            index = combo.findText(current_selection)
            if index != -1:
                combo.setCurrentIndex(index)
            else:
                combo.setCurrentIndex(0)
            combo.blockSignals(False)

    def accept_selection(self):
        self.selected_ports = {
            i: combo.currentText() for i, combo in enumerate(self.combos)
            if combo.currentText() and combo.currentText() != "선택 안 함"
        }
        if not self.selected_ports:
            QMessageBox.warning(self, "경고", "최소 하나 이상의 센서 포트를 선택해 주세요.")
            return
        self.accept()

# ==========================================
# 7. 메인 모니터링 창
# ==========================================
class DustMonitorApp(QMainWindow):
    def __init__(self, slot_mapping):
        super().__init__()
        self.threads = []
        self.sensor_widgets = {}
        self.slot_mapping = slot_mapping
        self.initUI()
        self.start_monitoring()

    def initUI(self):
        self.setWindowTitle("다중 미세먼지 모니터링 시스템")
        self.setStyleSheet("QMainWindow { background-color: white; }")
    
    # 강제로 크기를 지정하는 대신 최소 크기만 설정
        self.setMinimumSize(400, 150)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        self.setCentralWidget(scroll)

        scroll_content = QWidget()
        scroll.setWidget(scroll_content)
        self.scroll_layout = QVBoxLayout(scroll_content)
        self.scroll_layout.setContentsMargins(15, 15, 15, 15)
        self.scroll_layout.setSpacing(15)
        self.scroll_layout.addStretch(1)

    def start_monitoring(self):
        for sensor_index, port_name in self.slot_mapping.items():
            group_box = QGroupBox(f"센서 {sensor_index + 1} 포트: {port_name}")
            group_box.setFont(QFont("Malgun Gothic", 11, QFont.Bold))
            port_layout = QHBoxLayout(group_box)
            port_layout.setSpacing(10)

            w_pm10 = DustLevelWidget("PM10 (미세먼지)", Config.PM10_LEVELS)
            w_pm25 = DustLevelWidget("PM2.5 (초미세먼지)", Config.PM25_LEVELS)
            w_pm1 = DustLevelWidget("PM1.0 (극미세먼지)", Config.PM1_LEVELS)

            port_layout.addWidget(w_pm10)
            port_layout.addWidget(w_pm25)
            port_layout.addWidget(w_pm1)
            
            self.scroll_layout.insertWidget(self.scroll_layout.count() - 1, group_box)
            self.sensor_widgets[sensor_index] = {"pm10": w_pm10, "pm25": w_pm25, "pm1": w_pm1}

            thread = SerialThread(port_name, sensor_index)
            thread.data_signal.connect(self.update_data)
            thread.error_signal.connect(self.handle_error)
            thread.start()
            self.threads.append(thread)

        sensor_count = len(self.slot_mapping)
        calculated_height = max(220, sensor_count * 245 + 60)
        self.resize(980, calculated_height)

    def update_data(self, sensor_index, pm10, pm25, pm1):
        widgets = self.sensor_widgets.get(sensor_index)
        if widgets:
            widgets["pm10"].update_val(pm10)
            widgets["pm25"].update_val(pm25)
            widgets["pm1"].update_val(pm1)

    def handle_error(self, sensor_index, error_msg):
        widgets = self.sensor_widgets.get(sensor_index)
        if widgets:
            widgets["pm10"].set_error_state(error_msg)
            widgets["pm25"].set_error_state(error_msg)
            widgets["pm1"].set_error_state(error_msg)

    def closeEvent(self, event):
        for thread in self.threads:
            thread.requestInterruption()
        
        for thread in self.threads:
            thread.wait(2000)
    
        for thread in self.threads:
            thread.logger.export_excel()
            
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
    
    db_path = os.path.join(BASE_DIR, "Data", "dust_measurement.db")
    SensorLogger.init_global_database(db_path)
    SensorLogger.cleanup_old_logs_global()

    dialog = PortSelectionDialog()
    if dialog.exec_() == QDialog.Accepted:
        main_window = DustMonitorApp(dialog.selected_ports)
        main_window.show()
        sys.exit(app.exec_())
    else:
        sys.exit(0)