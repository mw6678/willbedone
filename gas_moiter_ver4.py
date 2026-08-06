import tkinter as tk
from tkinter import ttk, messagebox
import serial
import threading
import time
from datetime import datetime
import csv
import os
from collections import deque
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

# --- UI 테마 색상 (오리지널 프로그램 스타일) ---
BG_COLOR = "#F0F2F5"
CARD_COLOR = "#FFFFFF"
PRIMARY_BLUE = "#1890FF"
ALARM_RED = "#FF4D4F"
TEXT_DARK = "#333333"
TEXT_LIGHT = "#888888"

class GasMonitorApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("New Gas Detector Software")
        self.geometry("1200x800")
        self.configure(bg=BG_COLOR)
        
        # 통신 및 데이터 변수
        self.serial_port = None
        self.is_running = False
        self.o2_data = deque([0.0]*60, maxlen=60) # 실시간 그래프용 60개 데이터
        self.alarm_count = 0
        
        # 엑셀(CSV) 데이터베이스 파일 경로 설정 (실행 파일과 같은 폴더에 생성)
        self.base_dir = os.path.dirname(os.path.abspath(__file__))
        self.csv_file = os.path.join(self.base_dir, "gas_history_db.csv")
        
        # DB 파일이 없으면 헤더 생성
        if not os.path.exists(self.csv_file):
            with open(self.csv_file, "w", newline="", encoding="utf-8-sig") as f:
                writer = csv.writer(f)
                writer.writerow(["RecordTime", "ProbeName", "GasType", "Value", "Unit", "State"])

        self.create_style()
        self.create_layout()
        self.show_page("Home")

    def create_style(self):
        style = ttk.Style(self)
        try: style.theme_use("clam")
        except tk.TclError: pass
        
        style.configure("TFrame", background=BG_COLOR)
        style.configure("Card.TFrame", background=CARD_COLOR)
        style.configure("Sidebar.TFrame", background=CARD_COLOR)
        
        style.configure("Menu.TButton", font=("맑은 고딕", 10, "bold"), padding=10, background=CARD_COLOR, borderwidth=0)
        style.configure("Action.TButton", font=("맑은 고딕", 10, "bold"), background=PRIMARY_BLUE, foreground="white")
        style.configure("Treeview", rowheight=28, font=("맑은 고딕", 9))
        style.configure("Treeview.Heading", font=("맑은 고딕", 9, "bold"), background="#FAFAFA")

    def create_layout(self):
        # 1. 좌측 사이드바 메뉴 (Home, RealData, History)
        sidebar = ttk.Frame(self, style="Sidebar.TFrame", width=160)
        sidebar.pack(side="left", fill="y")
        sidebar.pack_propagate(False)

        ttk.Label(sidebar, text="  Cloud Gas", font=("맑은 고딕", 13, "bold"), background=CARD_COLOR).pack(pady=20, anchor="w")
        
        menus = [("Home", "🏠 Home"), ("RealData", "📈 RealData"), ("History", "📜 History")]
        for key, text in menus:
            btn = ttk.Button(sidebar, text=text, style="Menu.TButton", command=lambda k=key: self.show_page(k))
            btn.pack(fill="x", pady=2, padx=10)

        # 2. 우측 메인 영역
        self.main_area = ttk.Frame(self, style="TFrame")
        self.main_area.pack(side="right", fill="both", expand=True, padx=20, pady=20)
        
        # 페이지 사전 생성
        self.pages = {
            "Home": self.create_home_page(),
            "RealData": self.create_realdata_page(),
            "History": self.create_history_page()
        }

    def show_page(self, page_name):
        for frame in self.pages.values():
            frame.pack_forget()
        self.pages[page_name].pack(fill="both", expand=True)

    # ==========================================
    # 탭 1. Home 화면 (사진 1번 형태 구현)
    # ==========================================
    def create_home_page(self):
        frame = ttk.Frame(self.main_area)
        
        # 상단 요약 카드 4개 (Fixes, Probes, Portables, Alarms)
        top_frame = ttk.Frame(frame)
        top_frame.pack(fill="x", pady=(0, 15))
        
        cards_info = [("Fixes", "1"), ("Probes", "3"), ("Portables", "0")]
        for title, val in cards_info:
            card = ttk.Frame(top_frame, style="Card.TFrame", padding=15)
            card.pack(side="left", fill="both", expand=True, padx=(0, 10))
            ttk.Label(card, text=title, font=("맑은 고딕", 9), background=CARD_COLOR, foreground=TEXT_LIGHT).pack(anchor="w")
            ttk.Label(card, text=val, font=("맑은 고딕", 20, "bold"), background=CARD_COLOR, foreground=TEXT_DARK).pack(anchor="w", pady=5)

        # 알람 카운트 카드 (오리지널 대시보드 스타일)
        alarm_card = ttk.Frame(top_frame, style="Card.TFrame", padding=15)
        alarm_card.pack(side="left", fill="both", expand=True)
        ttk.Label(alarm_card, text="Alarms (경고 횟수)", font=("맑은 고딕", 9), background=CARD_COLOR, foreground=TEXT_LIGHT).pack(anchor="w")
        self.lbl_alarm_count = ttk.Label(alarm_card, text="0", font=("맑은 고딕", 20, "bold"), background=CARD_COLOR, foreground=ALARM_RED)
        self.lbl_alarm_count.pack(anchor="w", pady=5)

        # 중앙 알람 그래프 영역
        mid_frame = ttk.Frame(frame, style="Card.TFrame", padding=15)
        mid_frame.pack(fill="both", expand=True, pady=(0, 15))
        ttk.Label(mid_frame, text="▌ Alarm Count Trend", font=("맑은 고딕", 11, "bold"), background=CARD_COLOR).pack(anchor="w")
        
        fig = Figure(figsize=(8, 2.5), dpi=100)
        ax = fig.add_subplot(111)
        ax.set_ylim(0, 1)
        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)
        canvas = FigureCanvasTkAgg(fig, master=mid_frame)
        canvas.get_tk_widget().pack(fill="both", expand=True, pady=5)

        return frame

    # ==========================================
    # 탭 2. RealData 화면 (사진 3번 형태 구현)
    # ==========================================
    def create_realdata_page(self):
        frame = ttk.Frame(self.main_area)
        
        # 상단 제어 바 (COM 포트 설정 및 Start 버튼)
        ctrl_frame = ttk.Frame(frame)
        ctrl_frame.pack(fill="x", pady=(0, 10))
        
        ttk.Label(ctrl_frame, text="COM Port:", background=BG_COLOR).pack(side="left", padx=(0, 5))
        self.port_var = tk.StringVar(value="COM3")
        ttk.Entry(ctrl_frame, textvariable=self.port_var, width=10).pack(side="left", padx=(0, 15))
        
        self.btn_start = ttk.Button(ctrl_frame, text="▶ Start & Connect", style="Action.TButton", command=self.toggle_realdata)
        self.btn_start.pack(side="left")
        
        ttk.Label(frame, text="Devices (Fixed)", font=("맑은 고딕", 14, "bold"), background=BG_COLOR).pack(anchor="w", pady=(10, 5))
        
        # 프로브 카드 영역 (Probe1: H2S, Probe2: O2, Probe3: CH4)
        cards_frame = ttk.Frame(frame)
        cards_frame.pack(fill="x", pady=(0, 15))
        
        self.probe_cards = {}
        probes = [("Probe1", "H2S"), ("Probe2", "O2"), ("Probe3", "CH4")]
        
        for p_id, gas in probes:
            card = ttk.Frame(cards_frame, style="Card.TFrame", padding=15)
            card.pack(side="left", fill="both", expand=True, padx=(0, 10))
            
            ttk.Label(card, text=f"⚙ {p_id}", font=("맑은 고딕", 9, "bold"), background=CARD_COLOR).pack(anchor="w")
            ttk.Label(card, text=gas, font=("맑은 고딕", 10), background=CARD_COLOR, foreground=TEXT_LIGHT).pack(pady=(5, 0))
            
            # 수치 표시 라벨
            lbl_val = ttk.Label(card, text="--", font=("맑은 고딕", 22, "bold"), background="#757575", foreground="white", anchor="center")
            lbl_val.pack(fill="x", pady=8)
            
            lbl_stat = ttk.Label(card, text="Lost", font=("맑은 고딕", 9), background=CARD_COLOR, foreground=TEXT_LIGHT)
            lbl_stat.pack()
            
            self.probe_cards[gas] = {"val": lbl_val, "stat": lbl_stat}

        # 하단 실시간 꺾은선 그래프 (O2 센서 기준)
        graph_card = ttk.Frame(frame, style="Card.TFrame", padding=10)
        graph_card.pack(fill="both", expand=True)
        ttk.Label(graph_card, text="▌ Real time data curve (O2 Sensor)", font=("맑은 고딕", 11, "bold"), background=CARD_COLOR).pack(anchor="w")
        
        self.fig = Figure(figsize=(8, 3), dpi=100)
        self.ax = self.fig.add_subplot(111)
        self.ax.set_ylim(0, 30)
        self.ax.spines['top'].set_visible(False)
        self.ax.spines['right'].set_visible(False)
        self.ax.grid(True, linestyle='--', alpha=0.3)
        self.line, = self.ax.plot(range(60), self.o2_data, color=PRIMARY_BLUE, linewidth=2)
        
        self.canvas = FigureCanvasTkAgg(self.fig, master=graph_card)
        self.canvas.get_tk_widget().pack(fill="both", expand=True)

        return frame

    # ==========================================
    # 탭 3. History 화면 (사진 4번 형태 구현)
    # ==========================================
    def create_history_page(self):
        frame = ttk.Frame(self.main_area)
        
        ttk.Label(frame, text="Devices History (Excel DB 조회)", font=("맑은 고딕", 14, "bold"), background=BG_COLOR).pack(anchor="w", pady=(0, 10))
        
        info_frame = ttk.Frame(frame, style="Card.TFrame", padding=15)
        info_frame.pack(fill="both", expand=True)
        
        # 검색 및 엑셀 내보내기 제어 바
        ctrl = ttk.Frame(info_frame, style="Card.TFrame")
        ctrl.pack(fill="x", pady=(0, 15))
        
        ttk.Button(ctrl, text="🔍 Query (데이터 불러오기)", style="Action.TButton", command=self.load_history_data).pack(side="left", padx=(0, 10))
        ttk.Label(ctrl, text=f"저장 위치: {self.csv_file}", font=("맑은 고딕", 9), background=CARD_COLOR, foreground=TEXT_LIGHT).pack(side="left")
        
        # 엑셀 표 형태의 Treeview
        cols = ("No", "RecordTime", "ProbeName", "GasType", "Value", "Unit", "State")
        self.history_tree = ttk.Treeview(info_frame, columns=cols, show="headings", height=18)
        
        col_widths = [50, 150, 100, 80, 80, 80, 100]
        for i, c in enumerate(cols):
            self.history_tree.heading(c, text=c)
            self.history_tree.column(c, width=col_widths[i], anchor="center")
            
        self.history_tree.pack(fill="both", expand=True)
        return frame

    def load_history_data(self):
        # 표 초기화
        for item in self.history_tree.get_children():
            self.history_tree.delete(item)
            
        # CSV 파일에서 데이터 읽어와서 표에 채우기
        if os.path.exists(self.csv_file):
            with open(self.csv_file, "r", encoding="utf-8-sig") as f:
                reader = csv.reader(f)
                next(reader) # 첫 줄(제목) 스킵
                rows = list(reader)
                for i, row in enumerate(reversed(rows)): # 최신 데이터가 위로 오게 역순 정렬
                    self.history_tree.insert("", "end", values=(i+1, row[0], row[1], row[2], row[3], row[4], row[5]))

    # ==========================================
    # 통신 및 데이터베이스 저장 로직 (핵심 기능)
    # ==========================================
    def modbus_crc16(self, data):
        crc = 0xFFFF
        for pos in data:
            crc ^= pos
            for _ in range(8):
                if (crc & 1) != 0: crc = (crc >> 1) ^ 0xA001
                else: crc >>= 1
        return crc

    def toggle_realdata(self):
        if not self.is_running:
            port = self.port_var.get().strip()
            try:
                self.serial_port = serial.Serial(port=port, baudrate=9600, timeout=0.5)
                self.is_running = True
                self.btn_start.config(text="■ Stop")
                threading.Thread(target=self.poll_gas_device, daemon=True).start()
            except Exception as e:
                messagebox.showerror("통신 에러", f"포트를 열 수 없습니다.\nCOM 포트 번호를 확인하세요.\n\n에러내용: {e}")
        else:
            self.is_running = False
            self.btn_start.config(text="▶ Start & Connect")
            if self.serial_port and self.serial_port.is_open:
                self.serial_port.close()

    def poll_gas_device(self):
        # 0x0000 번지 산소(O2) 데이터 요청 패킷 (ID 1, Func 3, Start 0, Len 1)
        tx_data = bytearray([0x01, 0x03, 0x00, 0x00, 0x00, 0x01])
        crc = self.modbus_crc16(tx_data)
        tx_data.append(crc & 0xFF)
        tx_data.append((crc >> 8) & 0xFF)
        
        while self.is_running:
            try:
                self.serial_port.reset_input_buffer()
                self.serial_port.write(tx_data)
                time.sleep(0.2)
                rx_data = self.serial_port.read(256)

                if len(rx_data) >= 7 and rx_data[1] == 0x03:
                    cal_crc = self.modbus_crc16(rx_data[:-2])
                    rx_crc = rx_data[-2] | (rx_data[-1] << 8)

                    if cal_crc == rx_crc:
                        raw_value = (rx_data[3] << 8) | rx_data[4]
                        actual_value = raw_value / 10.0 # 예: 209 -> 20.9
                        
                        # 산소 경고 범위 판별 (19.5% 미만 또는 23.5% 초과 시 알람)
                        state = "Normal"
                        is_alarm = False
                        if actual_value < 19.5 or actual_value > 23.5:
                            state = "ALARM"
                            is_alarm = True
                        
                        # 화면 갱신 및 엑셀(CSV) DB 자동 저장
                        self.after(0, self.update_dashboard_ui, actual_value, state, is_alarm)
                        self.save_to_excel_db(actual_value, state)
            except Exception:
                self.is_running = False
                break
            time.sleep(0.8)

    def save_to_excel_db(self, value, state):
        # 엑셀(CSV) 파일에 1초마다 실시간 데이터 행 추가 기록
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with open(self.csv_file, "a", newline="", encoding="utf-8-sig") as f:
            writer = csv.writer(f)
            writer.writerow([now, "Probe2", "O2", value, "VOL%", state])

    def update_dashboard_ui(self, value, state, is_alarm):
        if not self.is_running: return
        
        bg_col = ALARM_RED if is_alarm else PRIMARY_BLUE
        txt_col = ALARM_RED if is_alarm else "green"
        
        if is_alarm:
            self.alarm_count += 1
            self.lbl_alarm_count.config(text=str(self.alarm_count))
            
        # O2 카드 디자인 및 수치 갱신
        self.probe_cards["O2"]["val"].config(text=f"{value:.1f}", background=bg_col)
        self.probe_cards["O2"]["stat"].config(text=f"Status: {state}", foreground=txt_col)
        
        # 실시간 그래프 갱신
        self.o2_data.append(value)
        self.line.set_ydata(self.o2_data)
        self.canvas.draw()


if __name__ == "__main__":
    app = GasMonitorApp()
    app.mainloop()