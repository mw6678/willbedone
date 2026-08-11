import sys
import requests
import pandas as pd
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QLabel, QPushButton,
    QVBoxLayout, QHBoxLayout, QTabWidget, QTableWidget, 
    QTableWidgetItem, QFileDialog, QMessageBox, QFrame, QHeaderView, QLineEdit
)
from PyQt5.QtCore import QTimer, Qt
from PyQt5.QtGui import QFont

class WillbedoneApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.initUI()
        
        self.is_running = False
        
        # 타이머 설정 (지정된 주기마다 하이퍼링크에서 데이터 수신)
        self.timer = QTimer()
        self.timer.timeout.connect(self.fetch_web_data)

    def initUI(self):
        self.setWindowTitle('Willbedone Data Logger (Web Mode)')
        self.resize(1000, 650)
        self.setStyleSheet("background-color: #ECECEC;")

        # 메인 위젯 및 레이아웃
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
        # 2. 대시보드 및 컨트롤 영역
        # ----------------------------------------------------
        content_layout = QVBoxLayout()
        content_layout.setContentsMargins(15, 15, 15, 15)

        # [추가됨] 하이퍼링크(URL) 입력 영역
        url_layout = QHBoxLayout()
        url_label = QLabel("Data URL (하이퍼링크):")
        url_label.setFont(QFont("Arial", 12, QFont.Bold))
        self.url_input = QLineEdit("http://192.168.0.100/data") # 기본 예시 주소
        self.url_input.setFixedHeight(30)
        url_layout.addWidget(url_label)
        url_layout.addWidget(self.url_input)
        content_layout.addLayout(url_layout)

        dash_layout = QHBoxLayout()

        # (1) Viscosity 패널
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

        # (2) Temperature 패널
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

        # (3) 제어 버튼 (START / STOP)
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
        # 3. 하단 탭 (List) 및 우측 기능 버튼
        # ----------------------------------------------------
        bottom_layout = QHBoxLayout()

        self.tabs = QTabWidget()
        
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

        # ✨ [핵심 해결 부분] 조립된 레이아웃을 메인 위젯에 최종 부착합니다! ✨
        main_widget.setLayout(main_layout)

    # ----------------------------------------------------
    # 로직 및 웹(URL) 통신 처리
    # ----------------------------------------------------
    def start_measurement(self):
        url = self.url_input.text().strip()
        if not url.startswith("http"):
            QMessageBox.warning(self, "URL 오류", "올바른 하이퍼링크(http:// 또는 https://)를 입력해 주세요.")
            return

        if not self.is_running:
            self.is_running = True
            self.url_input.setReadOnly(True)  # 작동 중에는 주소 변경 불가
            self.timer.start(1000)  # 1초마다 데이터 수신 확인 (서버 부하 방지를 위해 1초 권장)
            self.btn_start.setStyleSheet("background-color: #008CBA; color: white; font-weight: bold;")

    def stop_measurement(self):
        self.is_running = False
        self.timer.stop()
        self.url_input.setReadOnly(False)
        self.btn_start.setStyleSheet("background-color: #2D4059; color: white; font-weight: bold;")

    def fetch_web_data(self):
        if not self.is_running:
            return
            
        url = self.url_input.text().strip()
        try:
            # 설정한 하이퍼링크(URL)에 접속하여 텍스트 데이터를 가져옴
            response = requests.get(url, timeout=2)
            
            if response.status_code == 200:
                raw_data = response.text.strip()
                
                # 기기 데이터 수신 예: $26,08,05,21,24,43,25.0,1,6, 0.00, 0.0*
                if raw_data.startswith('$') and raw_data.endswith('*'):
                    parsed = raw_data.strip('$*').split(',')
                    
                    if len(parsed) >= 11:
                        # 데이터 분리 및 가공
                        date_str = f"20{parsed[0].strip()}-{parsed[1].strip()}-{parsed[2].strip()}"
                        time_str = f"{parsed[3].strip()}:{parsed[4].strip()}:{parsed[5].strip()}"
                        temp = parsed[6].strip()
                        rotor = parsed[7].strip()
                        speed = parsed[8].strip()
                        visc = parsed[9].strip()
                        percent = parsed[10].strip()

                        # 1. 상단 대시보드 UI 업데이트
                        self.visc_value.setText(f"{visc} mPa.s")
                        self.temp_value.setText(f"{temp} °C")

                        # 2. 하단 표(Table)에 데이터 추가
                        row_idx = self.table.rowCount()
                        self.table.insertRow(row_idx)
                        
                        self.table.setItem(row_idx, 0, QTableWidgetItem(date_str))
                        self.table.setItem(row_idx, 1, QTableWidgetItem(time_str))
                        self.table.setItem(row_idx, 2, QTableWidgetItem(temp))
                        self.table.setItem(row_idx, 3, QTableWidgetItem(rotor))
                        self.table.setItem(row_idx, 4, QTableWidgetItem(speed))
                        self.table.setItem(row_idx, 5, QTableWidgetItem(visc))
                        self.table.setItem(row_idx, 6, QTableWidgetItem(percent))
                        
                        self.table.scrollToBottom()
        except requests.exceptions.RequestException as e:
            # 네트워크 오류 시 출력을 남기고 계속 대기
            print(f"Web Connection Error: {e}")

    def clear_table(self):
        self.table.setRowCount(0)
        self.visc_value.setText("- - - - - mPa.s")
        self.temp_value.setText("- - - - - °C")

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
            QMessageBox.information(self, "성공", "데이터가 성공적으로 엑셀 파일로 저장되었습니다.")

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = WillbedoneApp()
    ex.show()
    sys.exit(app.exec_())