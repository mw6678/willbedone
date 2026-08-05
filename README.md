import tkinter as tk
from tkinter import ttk, messagebox
import serial
import threading
import time

class GasMonitorApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("BH-DX103 실시간 가스 농도 모니터")
        self.geometry("500x380")
        self.resizable(False, False)

        # 통신 상태 제어용 변수
        self.serial_port = None
        self.is_running = False

        self.create_style()
        self.create_widgets()

    def create_style(self):
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure("Title.TLabel", font=("맑은 고딕", 12, "bold"))
        style.configure("Value.TLabel", font=("맑은 고딕", 65, "bold"))
        style.configure("Status.TLabel", font=("맑은 고딕", 14, "bold"))

    def create_widgets(self):
        # 상단 설정 영역 (COM 포트 입력 및 연결 버튼)
        top_frame = ttk.Frame(self, padding=10)
        top_frame.pack(fill="x")

        ttk.Label(top_frame, text="COM 포트:", font=("맑은 고딕", 10)).pack(side="left")
        
        self.port_var = tk.StringVar(value="COM3")
        ttk.Entry(top_frame, textvariable=self.port_var, width=10).pack(side="left", padx=10)

        self.btn_connect = ttk.Button(top_frame, text="연결 및 측정 시작", command=self.toggle_connection)
        self.btn_connect.pack(side="left", padx=5)

        # 중앙 표시 영역 (상태 및 농도 표시)
        mid_frame = ttk.Frame(self, padding=20)
        mid_frame.pack(expand=True, fill="both")

        self.lbl_status = ttk.Label(mid_frame, text="대기 중...", style="Status.TLabel", foreground="gray")
        self.lbl_status.pack(pady=5)

        self.lbl_value = ttk.Label(mid_frame, text="--", style="Value.TLabel", foreground="#1F4E78")
        self.lbl_value.pack(pady=10)

        # 하단 로그 영역 (디버그용 데이터 확인)
        bottom_frame = ttk.LabelFrame(self, text="수신 데이터 (Hex)", padding=8)
        bottom_frame.pack(fill="x", padx=15, pady=15)
        
        self.lbl_log = ttk.Label(bottom_frame, text="통신 대기 중...", font=("Consolas", 11), foreground="#555555")
        self.lbl_log.pack(anchor="w")

    def toggle_connection(self):
        # 연결이 안 되어 있을 때 -> 연결 시도
        if not self.is_running:
            port = self.port_var.get().strip()
            try:
                # 시리얼 포트 열기
                self.serial_port = serial.Serial(
                    port=port,
                    baudrate=9600,
                    bytesize=8,
                    parity='N',
                    stopbits=1,
                    timeout=0.2
                )
                self.is_running = True
                self.btn_connect.config(text="측정 중지")
                self.lbl_status.config(text="통신 상태: 연결됨", foreground="green")

                # 1초마다 데이터를 읽어오는 백그라운드 스레드 시작
                self.thread = threading.Thread(target=self.poll_data, daemon=True)
                self.thread.start()

            except Exception as e:
                messagebox.showerror("연결 오류", f"{port} 포트를 열 수 없거나 장치를 찾을 수 없습니다.\n\n{e}")
        
        # 연결되어 있을 때 -> 중지 시도
        else:
            self.is_running = False
            self.btn_connect.config(text="연결 및 측정 시작")
            self.lbl_status.config(text="대기 중...", foreground="gray")
            self.lbl_value.config(text="--")
            self.lbl_log.config(text="통신 대기 중...")
            
            if self.serial_port and self.serial_port.is_open:
                self.serial_port.close()

    def modbus_crc16(self, data):
        # CRC-16 계산 로직
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

    def poll_data(self):
        # 전송할 데이터 세팅 (ID:1, Func:3, Addr:0, Count:1)
        tx_data = bytearray([0x01, 0x03, 0x00, 0x00, 0x00, 0x01])
        crc = self.modbus_crc16(tx_data)
        tx_data.append(crc & 0xFF)
        tx_data.append((crc >> 8) & 0xFF)

        while self.is_running:
            try:
                if self.serial_port and self.serial_port.is_open:
                    # 버퍼 비우기 및 데이터 전송
                    self.serial_port.reset_input_buffer()
                    self.serial_port.write(tx_data)

                    # 장비 응답 대기 (200ms)
                    time.sleep(0.2)
                    rx_data = self.serial_port.read(256)

                    # 수신 결과 처리
                    if len(rx_data) > 0:
                        hex_str = " ".join([f"{b:02X}" for b in rx_data])
                        self.lbl_log.config(text=f"RX: {hex_str}")

                        # 정상 포맷인지 확인 (최소 7바이트: ID, Func, Size, Data(2), CRC(2))
                        if len(rx_data) >= 7 and rx_data[1] == 0x03:
                            cal_crc = self.modbus_crc16(rx_data[:-2])
                            rx_crc = rx_data[-2] | (rx_data[-1] << 8)

                            if cal_crc == rx_crc:
                                value = (rx_data[3] << 8) | rx_data[4]
                                # 안전하게 화면 업데이트
                                self.after(0, self.update_value, value, "통신 상태: 정상", "green")
                            else:
                                self.after(0, self.update_value, "--", "통신 상태: CRC 에러", "red")
                        else:
                            self.after(0, self.update_value, "--", "통신 상태: 알 수 없는 포맷", "orange")
                    else:
                        self.after(0, self.update_value, "--", "통신 상태: 응답 없음 (결선/전원 확인)", "red")

            except Exception:
                self.after(0, self.update_value, "--", "통신 상태: 포트 끊김", "red")
                self.is_running = False
                self.after(0, lambda: self.btn_connect.config(text="연결 및 측정 시작"))
                break

            # 1초 주기를 맞추기 위한 나머지 0.8초 대기
            time.sleep(0.8)

    def update_value(self, val, status_text, color):
        if not self.is_running: return
        self.lbl_value.config(text=str(val))
        self.lbl_status.config(text=status_text, foreground=color)


if __name__ == "__main__":
    app = GasMonitorApp()
    app.mainloop()
