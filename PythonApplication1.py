import traceback

try:
    # --- 여기서부터 원래 코드 ---
    import serial
    import serial.tools.list_ports
    import threading
    import time
    import re
    import sys
    import tkinter as tk
    from tkinter import font

    BAUD_RATE = 9600

    def auto_detect_port():
        ports = serial.tools.list_ports.comports()
        if not ports:
            return None
        for port in ports:
            description = port.description.upper()
            if any(keyword in description for keyword in ["USB", "UART", "CH340", "CP210", "SERIAL"]):
                return port.device
        return ports[0].device

    def get_air_quality_korea(pm25, pm10):
        if pm25 <= 15:
            pm25_stat, pm25_color = "좋음", "#0078D7"
        elif pm25 <= 35:
            pm25_stat, pm25_color = "보통", "#107C10"
        elif pm25 <= 75:
            pm25_stat, pm25_color = "나쁨", "#D83B01"
        else:
            pm25_stat, pm25_color = "매우 나쁨", "#A80000"

        if pm10 <= 30:
            pm10_stat, pm10_color = "좋음", "#0078D7"
        elif pm10 <= 80:
            pm10_stat, pm10_color = "보통", "#107C10"
        elif pm10 <= 150:
            pm10_stat, pm10_color = "나쁨", "#D83B01"
        else:
            pm10_stat, pm10_color = "매우 나쁨", "#A80000"
            
        return pm25_stat, pm25_color, pm10_stat, pm10_color

    class DustMonitorApp:
        def __init__(self, root):
            self.root = root
            self.root.title("willbedone - 미세먼지 모니터링")
            self.root.geometry("450x550")
            self.root.configure(bg="#005A9E")
            
            self.is_running = True

            self.title_font = font.Font(family="Helvetica", size=24, weight="bold")
            self.label_font = font.Font(family="Malgun Gothic", size=14)
            self.value_font = font.Font(family="Helvetica", size=20, weight="bold")
            self.status_font = font.Font(family="Malgun Gothic", size=16, weight="bold")
            self.info_font = font.Font(family="Malgun Gothic", size=10)

            self.main_frame = tk.Frame(self.root, bg="white")
            self.main_frame.pack(expand=True, fill="both", padx=8, pady=8)

            self.company_label = tk.Label(self.main_frame, text="willbedone", font=self.title_font, fg="#005A9E", bg="white")
            self.company_label.pack(pady=(30, 10))

            self.status_var = tk.StringVar(value="🔍 센서 찾는 중...")
            self.time_var = tk.StringVar(value="시간: --:--:--")
            tk.Label(self.main_frame, textvariable=self.status_var, font=self.info_font, fg="gray", bg="white").pack()
            tk.Label(self.main_frame, textvariable=self.time_var, font=self.info_font, fg="gray", bg="white").pack(pady=(0, 20))

            tk.Frame(self.main_frame, bg="#E0E0E0", height=2).pack(fill="x", padx=40, pady=10)

            self.create_data_row("PM 1.0 (극초미세)", "pm1")
            self.create_data_row("PM 2.5 (초미세)", "pm25", has_status=True)
            self.create_data_row("PM 10 (미세먼지)", "pm10", has_status=True)

            self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

            self.serial_thread = threading.Thread(target=self.read_serial_data)
            self.serial_thread.daemon = True
            self.serial_thread.start()

        def create_data_row(self, title, prefix, has_status=False):
            frame = tk.Frame(self.main_frame, bg="white")
            frame.pack(fill="x", padx=40, pady=10)

            title_label = tk.Label(frame, text=title, font=self.label_font, bg="white", width=14, anchor="w")
            title_label.pack(side="left")

            val_var = tk.StringVar(value="--")
            setattr(self, f"{prefix}_var", val_var)
            
            val_label = tk.Label(frame, textvariable=val_var, font=self.value_font, bg="white", fg="#333333", width=4, anchor="e")
            val_label.pack(side="left")

            unit_label = tk.Label(frame, text="µg/m³", font=self.info_font, bg="white", fg="gray")
            unit_label.pack(side="left", padx=(5, 10), side="left", anchor="s", pady=(0, 3))

            if has_status:
                stat_var = tk.StringVar(value="")
                setattr(self, f"{prefix}_stat_var", stat_var)
                stat_label = tk.Label(frame, textvariable=stat_var, font=self.status_font, bg="white")
                stat_label.pack(side="right")
                setattr(self, f"{prefix}_stat_label", stat_label)

        def update_ui(self, pm1, pm25, pm10, pm25_stat, pm25_color, pm10_stat, pm10_color):
            self.pm1_var.set(str(pm1))
            self.pm25_var.set(str(pm25))
            self.pm10_var.set(str(pm10))
            
            self.pm25_stat_var.set(pm25_stat)
            self.pm25_stat_label.config(fg=pm25_color)
            
            self.pm10_stat_var.set(pm10_stat)
            self.pm10_stat_label.config(fg=pm10_color)
            
            self.time_var.set(f"측정 시간: {time.strftime('%H:%M:%S')}")

        def read_serial_data(self):
            port_name = auto_detect_port()
            if port_name is None:
                self.status_var.set("❌ 센서를 찾을 수 없습니다.")
                return

            try:
                ser = serial.Serial(port_name, BAUD_RATE, timeout=1)
                self.status_var.set(f"✅ 연결됨: {port_name}")
                
                while self.is_running:
                    if ser.in_waiting > 0:
                        raw_data = ser.readline().decode('utf-8', errors='ignore').strip()
                        numbers = re.findall(r'\d+', raw_data)
                        
                        if len(numbers) >= 3:
                            try:
                                pm1 = int(numbers[0])
                                pm25 = int(numbers[1])
                                pm10 = int(numbers[2])
                                pm25_stat, pm25_color, pm10_stat, pm10_color = get_air_quality_korea(pm25, pm10)
                                self.root.after(0, self.update_ui, pm1, pm25, pm10, pm25_stat, pm25_color, pm10_stat, pm10_color)
                            except ValueError:
                                pass
                    time.sleep(0.1)
            except serial.SerialException:
                self.status_var.set(f"❌ 포트({port_name}) 접근 거부됨 (사용중)")
            finally:
                if 'ser' in locals() and ser.is_open:
                    ser.close()

        def on_closing(self):
            self.is_running = False
            self.root.destroy()

    if __name__ == "__main__":
        root = tk.Tk()
        app = DustMonitorApp(root)
        root.mainloop()

# --- 방어 코드 (에러가 나면 화면을 멈춤) ---
except Exception as e:
    print("\n" + "="*50)
    print("🚨 프로그램 실행 중 에러가 발생해서 창이 꺼지는 것을 막았습니다! 🚨")
    print("="*50)
    traceback.print_exc()
    print("="*50)
    input("위 영어 에러 메시지를 마우스로 드래그해서 복사한 뒤 알려주세요. (엔터 키를 누르면 창이 닫힙니다)")
