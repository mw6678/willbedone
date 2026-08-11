import sys
import pandas as pd
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QLabel, QPushButton,
    QVBoxLayout, QHBoxLayout, QTabWidget, QTableWidget, 
    QTableWidgetItem, QFileDialog, QMessageBox, QFrame, QHeaderView, QComboBox
)
from PyQt5.QtCore import QTimer, Qt
from PyQt5.QtGui import QFont

# 차트 출력을 위한 pyqtgraph
import pyqtgraph as pg

# 시리얼 통신 라이브러리 검사
try:
    import serial
    import serial.tools.list_ports
except ImportError as e:
    app = QApplication(sys.argv)
    QMessageBox.critical(None, "라이브러리 미설치", f"필요한 부품이 설치되지 않았습니다:\n{e}\n\ncmd에서 'python -m pip install pyserial pyqtgraph'를 실행하세요.")
    sys.exit(1)

class WillbedoneApp(QMainWindow):
    def __init__(self):
        super().__init__()
        
        # 차트 데이터 저장용 리스트
        self.time_indexes = []
        self.visc_data = []
        self.temp_data = []
        self.data_count = 0

        self.initUI()
        
        self.ser = None
        self.is_running = False
        
        self.timer = QTimer()
        self.timer.timeout.connect(self.read_serial_data)

        self.refresh_ports()

    def initUI(self):
        self.setWindowTitle('Willbedone Data Logger')
        self.resize(1050, 700)
        self.setStyleSheet("background-color: #ECECEC;")

        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        main_layout = QVBoxLayout()
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)

        # ----------------------------------------------------
        # 1. 상단 타이틀 바
        # ----------------------------------------------------
        header_frame = QFrame()
        header_frame.setStyleSheet("background-color: #0F5C9E; color: white;")
        header_frame.setFixedHeight(60)
        header_layout = QHBoxLayout(header_frame)

        title_label = QLabel("Willbedone")
        title_label.setFont(QFont("Arial", 24, QFont.Bold))
        header_layout.addWidget(title_label)
        header_layout.addStretch(1)
        main_layout.addWidget(header_frame)

        # ----------------------------------------------------
        # 2. 대시보드 및 포트 선택
        # ----------------------------------------------------
        content_layout = QVBoxLayout()
        content_layout.setContentsMargins(15, 15, 15, 15)

        port_layout = QHBoxLayout()
        port_label = QLabel("연결된 기기 선택:")
        port_label.setFont(QFont("Arial", 12, QFont.Bold))
        
        self.port_combo = QComboBox()
        self.port_combo.setFixedHeight(30)
        self.port_combo.setMinimumWidth(320)
        
        self.btn_refresh = QPushButton("🔄 새로고침")
        self.btn_refresh.setFixedHeight(30)
        self.btn_refresh.clicked.connect(self.refresh_ports)

        port_layout.addWidget(port_label)
        port_layout.addWidget(self.port_combo)
        port_layout.addWidget(self.btn_refresh)
        port_layout.addStretch(1)
        content_layout.addLayout(port_layout)

        dash_layout = QHBoxLayout()

        # Viscosity 표시
        visc_box = QFrame()
        visc_box.setStyleSheet("background-color: white; border: 2px solid #1B6EA2; border-radius: 4px;")
        visc_layout = QVBoxLayout(visc_box)
        visc_title = QLabel("Viscosity")
        visc_title.setStyleSheet("background-color: #1B6EA2; color: white; padding: 2px 10px; font-weight: bold;")
        self.visc_value = QLabel("- - - - - mPa.s")
        self.visc_value.setFont(QFont("Arial", 24, QFont.Bold))
        self.visc_value.setAlignment(Qt.AlignCenter)
        visc_layout.addWidget(visc_title)
        visc_layout.addWidget(self.visc_value)

        # Temperature 표시
        temp_box = QFrame()
        temp_box.setStyleSheet("background-color: white; border: 2px solid #1B6EA2; border-radius: 4px;")
        temp_layout = QVBoxLayout(temp_box)
        temp_title = QLabel("Temperature")
        temp_title.setStyleSheet("background-color: #1B6EA2; color: white; padding: 2px 10px; font-weight: bold;")
        self.temp_value = QLabel("- - - - - °C")
        self.temp_value.setFont(QFont("Arial", 24, QFont.Bold))
        self.temp_value.setAlignment(Qt.AlignCenter)
        temp_layout.addWidget(temp_title)
        temp_layout.addWidget(self.temp_value)

        # 제어 버튼
        btn_layout = QVBoxLayout()
        self.btn_start = QPushButton("START")
        self.btn_start.setFixedHeight(40)
        self.btn_start.setStyleSheet("background-color: #2D4059; color: white; font-weight: bold; font-size: 14px;")
        self.btn_start.clicked.connect(self.start_measurement)

        self.btn_stop = QPushButton("STOP")
        self.btn_stop.setFixedHeight(40)
        self.btn_stop.setStyleSheet("background-color: #888888; color: white; font-weight: bold; font-size: 14px;")
        self.btn_stop.clicked.connect(self.stop_measurement)

        btn_layout.addWidget(self.btn_start)
        btn_layout.addWidget(self.btn_stop)

        dash_layout.addWidget(visc_box, 4)
        dash_layout.addWidget(temp_box, 4)
        dash_layout.addLayout(btn_layout, 2)
        content_layout.addLayout(dash_layout)

        # ----------------------------------------------------
        # 3. 하단 탭 (Data List / Chart) 및 저장 버튼
        # ----------------------------------------------------
        bottom_layout = QHBoxLayout()
        self.tabs = QTabWidget()

        # [Tab 1] Data List
        self.tab_list = QWidget()
        list_layout = QVBoxLayout(self.tab_list)
        self.table = QTableWidget()
        self.table.setColumnCount(7)
        self.table.setHorizontalHeaderLabels([
            "Date", "Time", "Temperature", "Rotor", "Speed", "Viscosity", "Percent"
        ])
        header = self.table.horizontalHeader()
        for i in range(7):
            header.setSectionResizeMode(i, QHeaderView.Stretch)
        list_layout.addWidget(self.table)
        self.tabs.addTab(self.tab_list, "Data List")

        # [Tab 2] Chart (실시간 그래프)
        self.tab_chart = QWidget()
        chart_layout = QVBoxLayout(self.tab_chart)
        
        # pyqtgraph 차트 위젯 설정
        pg.setConfigOption('background', 'w')
        pg.setConfigOption('foreground', 'k')
        self.graphWidget = pg.PlotWidget()
        self.graphWidget.addLegend()
        self.graphWidget.showGrid(x=True, y=True)
        self.graphWidget.setLabel('left', 'Value')
        self.graphWidget.setLabel('bottom', 'Data Count')

        # Viscosity (파란색) & Temperature (빨간색) 곡선 생성
        self.visc_line = self.graphWidget.plot(pen=pg.mkPen(color='b', width=2), name="Viscosity (mPa.s)")
        self.temp_line = self.graphWidget.plot(pen=pg.mkPen(color='r', width=2), name="Temperature (°C)")

        chart_layout.addWidget(self.graphWidget)
        self.tabs.addTab(self.tab_chart, "Chart")

        # 우측 액션 버튼
        side_btn_layout = QVBoxLayout()
        
        self.btn_save = QPushButton("💾\nSAVE...")
        self.btn_save.setFixedSize(80, 80)
        self.btn_save.setStyleSheet("background-color: white; border: 1px solid #CCCCCC; font-weight: bold;")
        self.btn_save.clicked.connect(self.save_to_excel)

        self.btn_clear = QPushButton("Clear")
        self.btn_clear.setFixedWidth(80)
        self.btn_clear.clicked.connect(self.clear_table)

        side_btn_layout.addWidget(self.btn_save)
        side_btn_layout.addWidget(self.btn_clear)
        side_btn_layout.addStretch(1)

        bottom_layout.addWidget(self.tabs, 9)
        bottom_layout.addLayout(side_btn_layout, 1)

        content_layout.addLayout(bottom_layout)
        main_layout.addLayout(content_layout)
        
        main_widget.setLayout(main_layout)

    # ----------------------------------------------------
    # 포트 탐색 및 통신 처리
    # ----------------------------------------------------
    def refresh_ports(self):
        """PC에 연결된 실제 포트 목록을 탐색하여 콤보박스에 추가합니다."""
        self.port_combo.clear()
        try:
            ports = serial.tools.list_ports.comports()
            if not ports:
                self.port_combo.addItem("연결된 기기 없음", None)
            else:
                for port in ports:
                    # 'COM포트 - 장치 설명' 형태로 명확하게 표시
                    display_text = f"{port.device} - {port.description}"
                    self.port_combo.addItem(display_text, port.device)
                self.port_combo.setCurrentIndex(0)
        except Exception:
            self.port_combo.addItem("포트 조회 실패", None)

    def start_measurement(self):
        selected_port = self.port_combo.currentData()
        
        if not selected_port:
            QMessageBox.warning(self, "경고", "먼저 점도계를 연결하고 '새로고침'을 눌러주세요.")
            return

        if not self.is_running:
            try:
                self.ser = serial.Serial(selected_port, 9600, timeout=1)
                self.is_running = True
                self.port_combo.setEnabled(False)  
                self.btn_refresh.setEnabled(False)
                self.timer.start(500)  
                self.btn_start.setStyleSheet("background-color: #008CBA; color: white; font-weight: bold;")
            except Exception as e:
                QMessageBox.critical(self, "포트 연결 오류", f"기기를 열 수 없습니다. 다른 프로그램(PuTTY 등)이 포트를 사용 중인지 확인하세요.\n\n오류: {e}")

    def stop_measurement(self):
        self.is_running = False
        self.timer.stop()
        if self.ser and self.ser.is_open:
            self.ser.close()
        
        self.port_combo.setEnabled(True)
        self.btn_refresh.setEnabled(True)
        self.btn_start.setStyleSheet("background-color: #2D4059; color: white; font-weight: bold;")

    def read_serial_data(self):
        if self.ser and self.ser.in_waiting > 0:
            try:
                raw_data = self.ser.readline().decode('utf-8').strip()
                
                if raw_data.startswith('$') and raw_data.endswith('*'):
                    parsed = raw_data.strip('$*').split(',')
                    
                    if len(parsed) >= 11:
                        date_str = f"20{parsed[0].strip()}-{parsed[1].strip()}-{parsed[2].strip()}"
                        time_str = f"{parsed[3].strip()}:{parsed[4].strip()}:{parsed[5].strip()}"
                        temp_val = float(parsed[6].strip())
                        rotor = parsed[7].strip()
                        speed = parsed[8].strip()
                        visc_val = float(parsed[9].strip())
                        percent = parsed[10].strip()

                        # 1. 수치 표시 업데이트
                        self.visc_value.setText(f"{visc_val:.2f} mPa.s")
                        self.temp_value.setText(f"{temp_val:.1f} °C")

                        # 2. Data List 표 업데이트
                        row_idx = self.table.rowCount()
                        self.table.insertRow(row_idx)
                        self.table.setItem(row_idx, 0, QTableWidgetItem(date_str))
                        self.table.setItem(row_idx, 1, QTableWidgetItem(time_str))
                        self.table.setItem(row_idx, 2, QTableWidgetItem(str(temp_val)))
                        self.table.setItem(row_idx, 3, QTableWidgetItem(rotor))
                        self.table.setItem(row_idx, 4, QTableWidgetItem(speed))
                        self.table.setItem(row_idx, 5, QTableWidgetItem(str(visc_val)))
                        self.table.setItem(row_idx, 6, QTableWidgetItem(percent))
                        self.table.scrollToBottom()

                        # 3. Chart 데이터 기록 및 업데이트
                        self.data_count += 1
                        self.time_indexes.append(self.data_count)
                        self.visc_data.append(visc_val)
                        self.temp_data.append(temp_val)

                        self.visc_line.setData(self.time_indexes, self.visc_data)
                        self.temp_line.setData(self.time_indexes, self.temp_data)

            except Exception:
                pass 

    def clear_table(self):
        self.table.setRowCount(0)
        self.visc_value.setText("- - - - - mPa.s")
        self.temp_value.setText("- - - - - °C")
        
        # 차트 데이터 초기화
        self.time_indexes.clear()
        self.visc_data.clear()
        self.temp_data.clear()
        self.data_count = 0
        self.visc_line.setData([], [])
        self.temp_line.setData([], [])

    def save_to_excel(self):
        row_count = self.table.rowCount()
        column_count = self.table.columnCount()

        if row_count == 0:
            QMessageBox.warning(self, "경고", "저장할 데이터가 없습니다.")
            return

        headers = [self.table.horizontalHeaderItem(i).text() for i in range(column_count)]
        data = []

        for row in range(row_count):
            row_data = []
            for col in range(column_count):
                item = self.table.item(row, col)
                row_data.append(item.text() if item else "")
            data.append(row_data)

        df = pd.DataFrame(data, columns=headers)
        
        path, _ = QFileDialog.getSaveFileName(self, "Excel 저장", "Willbedone_Data.xlsx", "Excel Files (*.xlsx)")
        if path:
            df.to_excel(path, index=False)
            QMessageBox.information(self, "성공", "데이터가 성공적으로 엑셀로 저장되었습니다.")

if __name__ == '__main__':
    try:
        app = QApplication(sys.argv)
        ex = WillbedoneApp()
        ex.show()
        sys.exit(app.exec_())
    except Exception as e:
        print(f"Error: {e}")