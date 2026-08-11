import serial
import serial.tools.list_ports
import time
import re
import sys

# ==========================================
# 사용자 설정 영역
# ==========================================
BAUD_RATE = 9600   # 센서 통신 속도 (보통 9600, 필요시 수정)
# ==========================================

def auto_detect_port():
    """연결된 시리얼 포트를 자동으로 찾아서 반환합니다."""
    ports = serial.tools.list_ports.comports()
    
    if not ports:
        print("❌ 연결된 시리얼 포트를 찾을 수 없습니다.")
        print("센서 케이블이 PC에 잘 연결되어 있는지 확인해 주세요.")
        return None

    # 1순위: 미세먼지 센서에서 흔히 쓰이는 USB-to-Serial 변환기 이름 스캔
    for port in ports:
        description = port.description.upper()
        if any(keyword in description for keyword in ["USB", "UART", "CH340", "CP210", "SERIAL"]):
            print(f"✅ 센서 포트 자동 감지됨: {port.device} ({port.description})")
            return port.device
            
    # 2순위: 특별한 키워드가 없다면 검색된 첫 번째 포트를 사용
    print(f"✅ 포트 자동 감지됨: {ports[0].device} ({ports[0].description})")
    return ports[0].device

def get_air_quality_korea(pm25, pm10):
    """한국 환경부 미세먼지 기준에 따라 상태를 반환합니다."""
    # 초미세먼지(PM2.5) 상태 평가
    if pm25 <= 15:
        pm25_status = "🔵 좋음"
    elif pm25 <= 35:
        pm25_status = "🟢 보통"
    elif pm25 <= 75:
        pm25_status = "🟡 나쁨"
    else:
        pm25_status = "🔴 매우 나쁨"

    # 미세먼지(PM10) 상태 평가
    if pm10 <= 30:
        pm10_status = "🔵 좋음"
    elif pm10 <= 80:
        pm10_status = "🟢 보통"
    elif pm10 <= 150:
        pm10_status = "🟡 나쁨"
    else:
        pm10_status = "🔴 매우 나쁨"
        
    return pm25_status, pm10_status

def main():
    # 포트 자동 탐색
    print("🔍 연결된 센서를 찾고 있습니다...")
    port_name = auto_detect_port()
    
    if port_name is None:
        sys.exit() # 포트를 찾지 못하면 프로그램 종료

    try:
        # 시리얼 포트 연결
        ser = serial.Serial(port_name, BAUD_RATE, timeout=1)
        print(f"\n[{port_name}] 통신이 성공적으로 연결되었습니다! 데이터를 기다리는 중...\n")
        print("-" * 50)
        
        while True:
            if ser.in_waiting > 0:
                # 센서로부터 한 줄을 읽어와서 문자열로 변환
                raw_data = ser.readline().decode('utf-8', errors='ignore').strip()
                
                # 정규표현식을 사용하여 문자열에서 연속된 숫자들만 추출
                numbers = re.findall(r'\d+', raw_data)
                
                # PM1.0, PM2.5, PM10 3개의 값이 정상적으로 들어왔는지 확인
                if len(numbers) >= 3:
                    try:
                        pm1 = int(numbers[0])
                        pm25 = int(numbers[1])
                        pm10 = int(numbers[2])
                        
                        # 대기질 상태 평가
                        pm25_status, pm10_status = get_air_quality_korea(pm25, pm10)
                        
                        # 사람들이 보기 쉽게 콘솔에 출력
                        print(f"🕒 측정 시간: {time.strftime('%H:%M:%S')}")
                        print(f" 🌫️ 극초미세먼지 (PM1.0) : {pm1} µg/m³")
                        print(f" 🌫️ 초미세먼지   (PM2.5) : {pm25} µg/m³\t[{pm25_status}]")
                        print(f" 🌫️ 미세먼지     (PM10)  : {pm10} µg/m³\t[{pm10_status}]")
                        print("-" * 50)
                        
                    except ValueError:
                        pass 
                else:
                    pass 
                    
            time.sleep(0.1) # CPU 과부하 방지
            
    except serial.SerialException:
        print(f"\n❌ 오류: {port_name} 포트를 열 수 없습니다.")
        print("💡 하이퍼터미널 등 다른 프로그램이 이 포트를 이미 사용 중인지 확인하고 종료해 주세요.")
    except KeyboardInterrupt:
        print("\n모니터링을 종료합니다.")
    finally:
        if 'ser' in locals() and ser.is_open:
            ser.close()

if __name__ == "__main__":
    main()