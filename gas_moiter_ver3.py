import tkinter as tk
from tkinter import ttk, messagebox
import serial
import threading
import time
from collections import deque
from datetime import datetime
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

# --- 색상 테마 정의 ---
BG_COLOR = "#F0F2F5"      # 전체 배경색
CARD_COLOR = "#FFFFFF"    # 카드(흰색) 배경색
PRIMARY_BLUE = "#1890FF"  # 포인트 블루 색상
TEXT_DARK = "#333333"     # 짙은 글씨
TEXT_LIGHT = "#888888"    # 옅은 글씨

class GasDashboardApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("New Gas Detector Software")
        self.geometry("1200x800")
        self.configure(bg=BG_COLOR)
        
        # 통신 상태 제어용 변수
        self.serial_port = None
        self.is_running = False
        self.o2_data = deque([0.0]*60, maxlen=60)
        
        self.create_style()
        self.create_layout()
        
        # 프로그램 시작 시 기본 화면(Home) 띄우기
        self.show_page("Home")

    def create_style(self):
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
            
        style.configure("TFrame", background=BG_COLOR)
        style.configure("Card.TFrame", background=CARD_COLOR)
        style.configure("Sidebar.TFrame", background=CARD_COLOR)
        
        # 라벨 스타일
        style.configure("Title.TLabel", font=("맑은 고딕", 16, "bold"), background=BG_COLOR)
        style.configure("CardTitle.TLabel", font=("맑은 고딕", 10), background=CARD_COLOR, foreground=TEXT_LIGHT)
        style.configure("CardVal.TLabel", font=("맑은 고딕", 24, "bold"), background=CARD_COLOR, foreground=TEXT_DARK)
        
        # 버튼 스타일
        style.configure("Menu.TButton", font=("맑은 고딕", 11, "bold"), padding=10, background=CARD_COLOR, borderwidth=0)
        style.configure("Action.TButton", font=("맑은 고딕", 10, "bold"), background=PRIMARY_BLUE, foreground="white")
        style.configure("Danger.TButton", font=("맑은 고딕", 10), background="#FF4D4F", foreground="white")

        # Treeview (표) 스타일
        style.configure("Treeview", rowheight=30, borderwidth=0, font=("맑은 고딕", 9))
        style.configure("Treeview.Heading", font=("맑은 고딕", 9, "bold"), background="#FAFAFA", foreground=TEXT_DARK)

    def create_layout(self):
        # --- 1. 좌측 사이드바 (메뉴 영역) ---
        sidebar = ttk.Frame(self, style="Sidebar.TFrame", width=160)
        sidebar.pack(side="left", fill="y")
        sidebar.pack_propagate(False)

        ttk.Label(sidebar, text="  Cloud Gas", font=("맑은 고딕", 14, "bold"), background=CARD_COLOR).pack(pady=20, anchor="w")
        
        self.menu_btns = {}
        menus = [("Home", "🏠 Home"), ("Device", "🎛️ Device"), ("RealData", "📈 RealData"), ("History", "📜 History")]
        
        for key, text in menus:
            btn = ttk.Button(sidebar, text=text, style="Menu.TButton", 
                             command=lambda k=key: self.show_page(k))
            btn.pack(fill="x", pady=2, padx=10)
            self.menu_btns[key] = btn

        # --- 2. 우측 메인 영역 (페이지들이 교체될 공간) ---
        self.main_area = ttk.Frame(self, style="TFrame")
        self.main_area.pack(side="right", fill="both", expand=True, padx=20, pady=20)
        
        # 각 페이지 프레임 생성
        self.pages = {
            "Home": self.create_home_page(),
            "Device": self.create_device_page(),
            "RealData": self.create_realdata_page(),
            "History": self.create_history_page()
        }

    def show_page(self, page_name):
        # 모든 페이지 숨기기
        for frame in self.pages.values():
            frame.pack_forget()
            
        # 선택한 페이지만 보이기
        self.pages[page_name].pack(fill="both", expand=True)
        
        # TODO: 선택된 메뉴 버튼의 색상을 변경하는 로직(생략)

    # ==========================================
    # 탭 1. Home 화면[cite: 1]
    # ==========================================
    def create_home_page(self):
        frame = ttk.Frame(self.main_area)
        
        # 상단 요약 카드 4개
        top_frame = ttk.Frame(frame)
        top_frame.pack(fill="x", pady=(0, 15))
        
        cards_info = [("Fixeds", "1"), ("Probes", "3"), ("Portables", "0"), ("Alarms", "0")]
        for title, val in cards_info:
            card = ttk.Frame(top_frame, style="Card.TFrame", padding=15)
            card.pack(side="left", fill="both", expand=True, padx=(0, 10))
            ttk.Label(card, text=title, style="CardTitle.TLabel").pack(anchor="w")
            ttk.Label(card, text=val, style="CardVal.TLabel").pack(anchor="w", pady=5)
            
        # 중앙 그래프 영역 (Alarm / Alarm count)
        mid_frame = ttk.Frame(frame)
        mid_frame.pack(fill="both", expand=True, pady=(0, 15))
        
        alarm_card = ttk.Frame(mid_frame, style="Card.TFrame", padding=15)
        alarm_card.pack(side="left", fill="both", expand=True, padx=(0, 10))
        ttk.Label(alarm_card, text="▌ Alarm", font=("맑은 고딕", 11, "bold"), background=CARD_COLOR).pack(anchor="w")
        
        graph_card = ttk.Frame(mid_frame, style="Card.TFrame", padding=15)
        graph_card.pack(side="left", fill="both", expand=True)
        ttk.Label(graph_card, text="▌ Alarm count", font=("맑은 고딕", 11, "bold"), background=CARD_COLOR).pack(anchor="w")
        
        # Alarm count용 임시 빈 그래프 
        fig = Figure(figsize=(4, 2), dpi=100)
        ax = fig.add_subplot(111)
        ax.set_ylim(0, 1)
        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)
        canvas = FigureCanvasTkAgg(fig, master=graph_card)
        canvas.get_tk_widget().pack(fill="both", expand=True, pady=10)

        # 하단 알람 기록 테이블
        bot_frame = ttk.Frame(frame, style="Card.TFrame", padding=15)
        bot_frame.pack(fill="both", expand=True)
        ttk.Label(bot_frame, text="▌ Recent alarm records", font=("맑은 고딕", 11, "bold"), background=CARD_COLOR).pack(anchor="w", pady=(0,10))
        
        cols = ("ID", "Alarm time", "Device number", "Host name", "Alarm probe", "GasType", "Alarm value", "Unit", "Data type")
        tree = ttk.Treeview(bot_frame, columns=cols, show="headings")
        for c in cols:
            tree.heading(c, text=c)
            tree.column(c, width=100, anchor="center")
        tree.pack(fill="both", expand=True)
        
        return frame

    # ==========================================
    # 탭 2. Device 화면
    # ==========================================
    def create_device_page(self):
        frame = ttk.Frame(self.main_area)
        
        # 상단 컨트롤 바
        ctrl_frame = ttk.Frame(frame, style="Card.TFrame", padding=15)
        ctrl_frame.pack(fill="x", pady=(0, 15))
        
        ttk.Label(ctrl_frame, text="COM", background=CARD_COLOR).pack(side="left", padx=5)
        self.dev_port_var = tk.StringVar(value="COM1")
        ttk.Combobox(ctrl_frame, textvariable=self.dev_port_var, values=("COM1", "COM2", "COM3", "COM4"), width=10).pack(side="left", padx=5)
        
        ttk.Label(ctrl_frame, text="Device type", background=CARD_COLOR).pack(side="left", padx=(15, 5))
        ttk.Combobox(ctrl_frame, values=("Fixed Gas Detector",), width=20).pack(side="left", padx=5)
        
        ttk.Button(ctrl_frame, text="🔗 Connect", style="Action.TButton").pack(side="left", padx=15)
        ttk.Button(ctrl_frame, text="+ New").pack(side="right")
        
        # 디바이스 목록 테이블
        table_frame = ttk.Frame(frame, style="Card.TFrame")
        table_frame.pack(fill="both", expand=True)
        
        cols = ("ID", "Host name", "Device number", "Probes", "Ownership", "CreateTime", "Location")
        tree = ttk.Treeview(table_frame, columns=cols, show="headings")
        for c in cols:
            tree.heading(c, text=c)
            tree.column(c, width=120, anchor="center")
        
        # 기본 데이터 삽입
        tree.insert("", "end", values=("1", "고정식", "---", "3", "Native", "2024-12-31 15:22:58", ""))
        tree.pack(fill="both", expand=True, padx=2, pady=2)
        
        return frame

    # ==========================================
    # 탭 3. RealData 화면
    # ==========================================
    def create_realdata_page(self):
        frame = ttk.Frame(self.main_area)
        
        # 상단 컨트롤 바
        ctrl_frame = ttk.Frame(frame)
        ctrl_frame.pack(fill="x", pady=(0, 10))
        
        ttk.Label(ctrl_frame, text="Interval(s)", background=BG_COLOR).pack(side="left", padx=5)
        ttk.Combobox(ctrl_frame, values=("1", "2", "5", "10"), width=5).pack(side="left", padx=5)
        
        self.btn_start = ttk.Button(ctrl_frame, text="▶ Start", style="Action.TButton", command=self.toggle_realdata)
        self.btn_start.pack(side="left", padx=15)
        
        ttk.Label(frame, text="Devices", style="Title.TLabel").pack(anchor="w", pady=(10, 5))
        
        # 중앙 프로브 카드 영역
        cards_frame = ttk.Frame(frame)
        cards_frame.pack(fill="x", pady=(0, 15))
        
        self.real_cards = {}
        probes = [("Probe1", "H2S", "0"), ("Probe2", "O2", "0"), ("Probe3", "CH4", "0")]
        
        for probe_id, gas, val in probes:
            card = ttk.Frame(cards_frame, style="Card.TFrame", padding=15)
            card.pack(side="left", fill="both", expand=True, padx=(0, 10))
            
            ttk.Label(card, text=f" ⚙ {probe_id}", font=("맑은 고딕", 9, "bold"), background=CARD_COLOR).pack(anchor="w")
            ttk.Label(card, text=gas, font=("맑은 고딕", 10), background=CARD_COLOR).pack(pady=(10, 0))
            
            lbl_val = ttk.Label(card, text=val, font=("맑은 고딕", 20, "bold"), background="#757575", foreground="white", padding=5, width=10, anchor="center")
            lbl_val.pack(pady=5)
            
            lbl_stat = ttk.Label(card, text="Lost", font=("맑은 고딕", 9), background=CARD_COLOR, foreground=TEXT_LIGHT)
            lbl_stat.pack()
            
            self.real_cards[gas] = {"val": lbl_val, "stat": lbl_stat}

        # 하단 실시간 그래프
        graph_card = ttk.Frame(frame, style="Card.TFrame", padding=10)
        graph_card.pack(fill="both", expand=True)
        ttk.Label(graph_card, text="▌ Real time data curve", font=("맑은 고딕", 11, "bold"), background=CARD_COLOR).pack(anchor="w")
        
        self.fig = Figure(figsize=(8, 3), dpi=100)
        self.ax = self.fig.add_subplot(111)
        self.ax.set_ylim(0, 30)
        self.ax.spines['top'].set_visible(False)
        self.ax.spines['right'].set_visible(False)
        self.ax.grid(True, linestyle='--', alpha=0.3)
        self.line, = self.ax.plot(range(60), self.o2_data, color="#E53935", linewidth=2, label="Series")
        self.ax.legend(loc="lower right", frameon=False)
        
        self.canvas = FigureCanvasTkAgg(self.fig, master=graph_card)
        self.canvas.get_tk_widget().pack(fill="both", expand=True)

        return frame

    # ==========================================
    # 탭 4. History 화면
    # ==========================================
    def create_history_page(self):
        frame = ttk.Frame(self.main_area)
        
        ttk.Label(frame, text="Devices", style="Title.TLabel").pack(anchor="w", pady=(0, 10))
        
        info_frame = ttk.Frame(frame, style="Card.TFrame", padding=15)
        info_frame.pack(fill="both", expand=True)
        
        # 컨트롤 바
        ctrl = ttk.Frame(info_frame, style="Card.TFrame")
        ctrl.pack(fill="x", pady=(0, 15))
        
        ttk.Label(ctrl, text="Probe", background=CARD_COLOR).pack(side="left")
        ttk.Combobox(ctrl, values=("탐두1",), width=8).pack(side="left", padx=5)
        
        ttk.Label(ctrl, text="  DateTime", background=CARD_COLOR).pack(side="left")
        ttk.Entry(ctrl, width=20).pack(side="left", padx=5) # 시작시간
        ttk.Label(ctrl, text="To", background=CARD_COLOR).pack(side="left")
        ttk.Entry(ctrl, width=20).pack(side="left", padx=5) # 종료시간
        
        ttk.Button(ctrl, text="🔍 Query", style="Action.TButton").pack(side="left", padx=10)
        ttk.Button(ctrl, text="📥 Export").pack(side="left")
        
        # 테이블
        cols = ("ID", "Probe name", "GasType", "Value", "Unit", "State", "RecordTime")
        tree = ttk.Treeview(info_frame, columns=cols, show="headings", height=20)
        for c in cols:
            tree.heading(c, text=c)
            tree.column(c, width=120, anchor="center")
        tree.pack(fill="both", expand=True)
        
        return frame

    # ==========================================
    # 통신 로직 (RealData 탭 전용)
    # ==========================================
    def toggle_realdata(self):
        if not self.is_running:
            port = self.dev_port_var.get().strip() # Device 탭에서 입력한 COM 포트 사용
            try:
                self.serial_port = serial.Serial(port=port, baudrate=9600, timeout=0.5)
                self.is_running = True
                self.btn_start.config(text="■ Stop")
                threading.Thread(target=self.poll_data, daemon=True).start()
            except Exception as e:
                messagebox.showerror("Error", f"포트를 열 수 없습니다.\nDevice 탭의 COM 포트를 확인하세요.")
        else:
            self.is_running = False
            self.btn_start.config(text="▶ Start")
            if self.serial_port and self.serial_port.is_open:
                self.serial_port.close()

    def poll_data(self):
        tx_data = bytearray([0x01, 0x03, 0x00, 0x00, 0x00, 0x01]) # O2 센서 요청
        # CRC 계산은 생략(하드코딩 가능하나 이전 코드와 동일하게 적용하면 됨)
        # 예제를 위해 더미 통신 시뮬레이션 코드 추가
        
        while self.is_running:
            try:
                # *실제 통신 로직 대신 모니터링 시뮬레이션 적용*
                time.sleep(1)
                
                # 임의의 정상 가스값 (20.8 ~ 21.0)
                import random
                actual_value = round(random.uniform(20.8, 21.0), 1)
                
                self.after(0, self.update_realdata, actual_value)

            except Exception:
                self.is_running = False
                break

    def update_realdata(self, value):
        if not self.is_running: return
        
        # 프로브 상태 업데이트 (회색 배경 -> 파란 배경)
        self.real_cards["O2"]["val"].config(text=f"{value:.1f}", background=PRIMARY_BLUE)
        self.real_cards["O2"]["stat"].config(text="Normal", foreground="green")
        
        # 그래프 데이터 갱신
        self.o2_data.append(value)
        self.line.set_ydata(self.o2_data)
        self.canvas.draw()


if __name__ == "__main__":
    app = GasDashboardApp()
    app.mainloop()