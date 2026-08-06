import tkinter as tk
from tkinter import ttk, messagebox
import serial
import threading
import time
from collections import deque
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

class GasDashboardApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("New Gas Detector Software")
        self.geometry("1100x700")
        self.configure(bg="#F0F2F5")
        
        # 통신 및 데이터 변수
        self.serial_port = None
        self.is_running = False
        self.o2_data = deque([0.0]*60, maxlen=60) # 최근 60개 데이터 저장 (그래프용)
        
        self.create_style()
        self.create_layout()

    def create_style(self):
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
            
        # 색상 및 폰트 설정
        style.configure("Sidebar.TFrame", background="#FFFFFF")
        style.configure("Main.TFrame", background="#F0F2F5")
        style.configure("Card.TFrame", background="#FFFFFF", relief="flat")
        
        style.configure("Menu.TButton", font=("맑은 고딕", 11), padding=10, background="#FFFFFF", borderwidth=0)
        style.configure("Action.TButton", font=("맑은 고딕", 10, "bold"), background="#1890FF", foreground="white")
        
        style.configure("Header.TLabel", font=("맑은 고딕", 16, "bold"), background="#F0F2F5")
        style.configure("CardTitle.TLabel", font=("맑은 고딕", 11), background="#FFFFFF", foreground="#555555")
        style.configure("CardValue.TLabel", font=("맑은 고딕", 24, "bold"), background="#FFFFFF")
        style.configure("CardStatus.TLabel", font=("맑은 고딕", 10), background="#FFFFFF", foreground="#888888")

    def create_layout(self):
        # 1. 좌측 사이드바 (메뉴)
        sidebar = ttk.Frame(self, style="Sidebar.TFrame", width=150)
        sidebar.pack(side="left", fill="y")
        sidebar.pack_propagate(False)

        ttk.Label(sidebar, text="  Cloud Gas", font=("맑은 고딕", 14, "bold"), background="#FFFFFF").pack(pady=20, anchor="w")
        
        menus = ["Home", "Device", "RealData", "History"]
        for m in menus:
            btn = ttk.Button(sidebar, text=f"  {m}", style="Menu.TButton", cursor="hand2")
            btn.pack(fill="x", pady=2, padx=10)

        # 2. 우측 메인 영역
        self.main_area = ttk.Frame(self, style="Main.TFrame")
        self.main_area.pack(side="right", fill="both", expand=True, padx=20, pady=20)

        # 상단 타이틀
        header_frame = ttk.Frame(self.main_area, style="Main.TFrame")
        header_frame.pack(fill="x", pady=(0, 15))
        ttk.Label(header_frame, text="RealData", style="Header.TLabel").pack(side="left")

        # 컨트롤 바 (연결 설정)
        control_frame = ttk.Frame(self.main_area, style="Main.TFrame")
        control_frame.pack(fill="x", pady=(0, 15))
        
        ttk.Label(control_frame, text="COM Port:", background="#F0F2F5").pack(side="left", padx=(0, 5))
        self.port_var = tk.StringVar(value="COM3")
        ttk.Entry(control_frame, textvariable=self.port_var, width=10).pack(side="left", padx=(0, 15))
        
        self.btn_connect = ttk.Button(control_frame, text="▶ Connect & Start", style="Action.TButton", command=self.toggle_connection)
        self.btn_connect.pack(side="left")

        # 3. 중앙 카드 영역 (센서 프로브들)
        cards_frame = ttk.Frame(self.main_area, style="Main.TFrame")
        cards_frame.pack(fill="x", pady=(0, 20))

        self.cards = {}
        probes = [("Probe1 (H2S)", "0", "Lost"), ("Probe2 (O2)", "--", "Wait"), ("Probe3 (CH4)", "0", "Lost")]
        
        for name, val, stat in probes:
            card = ttk.Frame(cards_frame, style="Card.TFrame", padding=15)
            card.pack(side="left", fill="both", expand=True, padx=5)
            
            ttk.Label(card, text=name, style="CardTitle.TLabel").pack(anchor="w")
            
            # 값을 표시할 라벨 저장
            lbl_val = ttk.Label(card, text=val, style="CardValue.TLabel", foreground="#888888" if stat != "Wait" else "#1890FF")
            lbl_val.pack(pady=10)
            
            lbl_stat = ttk.Label(card, text=stat, style="CardStatus.TLabel")
            lbl_stat.pack()
            
            self.cards[name] = {"val": lbl_val, "stat": lbl_stat}

        # 4. 하단 실시간 그래프 영역
        graph_frame = ttk.Frame(self.main_area, style="Card.TFrame")
        graph_frame.pack(fill="both", expand=True, padx=5)
        
        ttk.Label(graph_frame, text="Real time data curve (O2)", style="CardTitle.TLabel").pack(anchor="w", padx=10, pady=10)
        
        # Matplotlib Figure 생성
        self.fig = Figure(figsize=(8, 3), dpi=100)
        self.ax = self.fig.add_subplot(111)
        self.ax.set_ylim(0, 30) # 산소 농도 범위 (0~30%)
        self.ax.margins(x=0)
        self.ax.grid(True, linestyle='--', alpha=0.6)
        self.line, = self.ax.plot(range(60), self.o2_data, color="#E53935", linewidth=2)
        
        self.canvas = FigureCanvasTkAgg(self.fig, master=graph_frame)
        self.canvas.get_tk_widget().pack(fill="both", expand=True, padx=10, pady=(0, 10))

    # --- 통신 및 로직 ---
    def modbus_crc16(self, data):
        crc = 0xFFFF
        for pos in data:
            crc ^= pos
            for _ in range(8):
                if (crc & 1) != 0:
                    crc >>= 1
                    crc ^= 0xA001
                else:
                    crc >>= 1
        return crc

    def toggle_connection(self):
        if not self.is_running:
            port = self.port_var.get().strip()
            try:
                self.serial_port = serial.Serial(port=port, baudrate=9600, timeout=0.5)
                self.is_running = True
                self.btn_connect.config(text="■ Stop")
                self.cards["Probe2 (O2)"]["stat"].config(text="Normal", foreground="green")
                
                # 통신 스레드 시작
                threading.Thread(target=self.poll_data, daemon=True).start()
            except Exception as e:
                messagebox.showerror("Connection Error", f"포트를 열 수 없습니다.\n{e}")
        else:
            self.is_running = False
            self.btn_connect.config(text="▶ Connect & Start")
            self.cards["Probe2 (O2)"]["stat"].config(text="Wait", foreground="#888888")
            if self.serial_port and self.serial_port.is_open:
                self.serial_port.close()

    def poll_data(self):
        # O2 센서 데이터 요청 패킷 (ID 1, Func 3, Addr 0, Len 1)
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
                        # 0x0000 번지 데이터를 읽어 10으로 나눔 (예: 209 -> 20.9)
                        raw_value = (rx_data[3] << 8) | rx_data[4]
                        actual_value = raw_value / 10.0
                        self.after(0, self.update_dashboard, actual_value)
                    else:
                        self.after(0, self.update_error, "CRC Error")
                else:
                    self.after(0, self.update_error, "Timeout")

            except Exception:
                self.is_running = False
                self.after(0, self.update_error, "Disconnected")
                self.after(0, lambda: self.btn_connect.config(text="▶ Connect & Start"))
                break

            time.sleep(0.8) # 1초 주기 유지

    def update_dashboard(self, value):
        if not self.is_running: return
        
        # 1. 카드 업데이트
        self.cards["Probe2 (O2)"]["val"].config(text=f"{value:.1f}")
        self.cards["Probe2 (O2)"]["stat"].config(text="Normal", foreground="green")
        
        # 2. 그래프 업데이트
        self.o2_data.append(value)
        self.line.set_ydata(self.o2_data)
        self.canvas.draw()

    def update_error(self, msg):
        if not self.is_running: return
        self.cards["Probe2 (O2)"]["stat"].config(text=msg, foreground="red")


if __name__ == "__main__":
    app = GasDashboardApp()
    app.mainloop()