import os
import json
import sqlite3
import math
import sys
import re
from datetime import datetime

# pandas 사용
try:
    import pandas as pd
except ImportError:
    print("❌ pandas 및 필수 엑셀 라이브러리가 설치되어 있지 않습니다.")
    print("터미널에서 아래 명령어를 실행하여 설치해주세요:")
    print("pip install pandas xlrd openpyxl")
    sys.exit(1)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE_DIR, "Data", "dust_measurement.db")
CALIB_FILE = os.path.join(BASE_DIR, "Data", "sensor_calibration.json")

MIN_MINUTE_COUNT = 45 

# KOTITI 키워드 포함
REF_KEYWORDS = ["기준기", "기준값", "기준농도", "REFERENCE", "BAM", "GRIMM", "PM2.5", "PM25", "KOTITI"]

def find_reference_file(base_dir):
    ignore_prefixes = ("~$", "Dust_log_")
    valid_extensions = (".xlsx", ".csv", ".xls")
    try:
        candidates = [f for f in os.listdir(base_dir) if any(f.lower().endswith(ext) for ext in valid_extensions) and not any(f.startswith(p) for p in ignore_prefixes) and os.path.isfile(os.path.join(base_dir, f))]
    except OSError: 
        return None
        
    if not candidates: 
        return None
    if len(candidates) == 1: 
        return candidates[0]
        
    print("\n📁 여러 개의 비교군 파일 발견:")
    for idx, fname in enumerate(candidates, 1): 
        print(f"   [{idx}] {fname}")
    while True:
        try:
            choice = input("\n사용할 파일 번호를 입력하세요: ").strip()
            sel_idx = int(choice) - 1
            if 0 <= sel_idx < len(candidates): return candidates[sel_idx]
        except (ValueError, EOFError): 
            pass

def get_sensor_hourly_averages(sensor_index=0):
    if not os.path.exists(DB_PATH): 
        raise FileNotFoundError(f"DB 파일을 찾을 수 없습니다: {DB_PATH}")
    
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    
    query = """
        SELECT 
            strftime('%Y-%m-%d %H:00:00', measured_at) AS hour_bucket,
            AVG(raw_pm25) AS avg_pm25,
            COUNT(*) AS data_count
        FROM measurements
        WHERE sensor_index = ? AND status = 'NORMAL'
        GROUP BY hour_bucket
        HAVING data_count >= ?
        ORDER BY hour_bucket ASC
    """
    cursor.execute(query, (sensor_index, MIN_MINUTE_COUNT))
    rows = cursor.fetchall()
    conn.close()
    return {row[0]: round(row[1], 2) for row in rows}

def load_reference_file(file_path):
    ext = file_path.lower().split('.')[-1]
    filename = os.path.basename(file_path)
    
    # 🟢 1. 파일 이름(예: 기준측정기 데이터_260929-260930.xls)에서 연도와 월 추출
    # 추출 실패 시 현재 연/월 사용
    file_year = datetime.now().year
    file_month = datetime.now().month
    match = re.search(r'(\d{2})(\d{2})\d{2}', filename)
    if match:
        file_year = 2000 + int(match.group(1)) # 26 -> 2026
        file_month = int(match.group(2))       # 09 -> 9
        
    try:
        if ext in ['xls', 'xlsx']:
            df = pd.read_excel(file_path, header=None)
        elif ext == 'csv':
            df = pd.read_csv(file_path, header=None, encoding="utf-8-sig")
        else:
            raise ValueError("지원하지 않는 파일 형식입니다.")
    except Exception as e:
        raise ValueError(f"파일을 읽어오는 데 실패했습니다: {e}")

    date_col_idx = None
    time_col_idx = None
    ref_col_idx = None
    header_row_idx = None
    
    # 상위 15개 행 스캔하여 헤더 위치 탐색
    for r_idx, row in df.head(15).iterrows():
        for c_idx, val in enumerate(row):
            val_str = str(val).upper().strip() if pd.notna(val) else ""
            
            if date_col_idx is None and any(k in val_str for k in ["날짜", "DATE", "일자", "조회일자", "일"]): 
                date_col_idx = c_idx
            elif time_col_idx is None and any(k in val_str for k in ["시간", "TIME"]): 
                time_col_idx = c_idx
                
            if ref_col_idx is None and any(k in val_str for k in REF_KEYWORDS): 
                ref_col_idx = c_idx
                
        # 일, 시간, 기준값 열을 모두 찾으면 헤더로 인식
        if date_col_idx is not None and time_col_idx is not None and ref_col_idx is not None:
            header_row_idx = r_idx
            break

    if ref_col_idx is None: 
        raise ValueError("파일에서 'KOTITI' 등 기준측정기 열을 찾지 못했습니다.")
    if date_col_idx is None or time_col_idx is None:
        raise ValueError("파일에서 '일' 또는 '시간' 열을 찾지 못했습니다.")

    ref_data = {}
    
    # 🟢 2. 실제 데이터 행 처리
    for _, row in df.iloc[header_row_idx + 1:].iterrows():
        # 빈칸이거나 결측치면 건너뜀
        if pd.isna(row[date_col_idx]) or pd.isna(row[time_col_idx]) or pd.isna(row[ref_col_idx]):
            continue
            
        raw_date_str = str(row[date_col_idx]).strip()
        raw_time_str = str(row[time_col_idx]).strip()
        raw_val = row[ref_col_idx]
        
        # "29일"에서 숫자만 추출 -> 29
        day_match = re.search(r'\d+', raw_date_str)
        # "0", "1" 등 시간에서 숫자만 추출
        hour_match = re.search(r'\d+', raw_time_str)
        
        if day_match and hour_match:
            try:
                day = int(day_match.group())
                hour = int(hour_match.group())
                
                # 파일명에서 얻은 연/월과 엑셀의 일/시간을 결합하여 완벽한 DB 날짜 형식 생성
                norm_time = f"{file_year}-{file_month:02d}-{day:02d} {hour:02d}:00:00"
                
                # 기준값 추가
                ref_data[norm_time] = float(raw_val)
            except (ValueError, TypeError):
                pass
                
    return ref_data

def calculate_regression(x_vals, y_vals):
    n = len(x_vals)
    mean_x, mean_y = sum(x_vals) / n, sum(y_vals) / n
    ss_xx = sum((x - mean_x) ** 2 for x in x_vals)
    ss_yy = sum((y - mean_y) ** 2 for y in y_vals)
    ss_xy = sum((x - mean_x) * (y - mean_y) for x, y in zip(x_vals, y_vals))
    if ss_xx == 0: 
        raise ValueError("센서 값의 변동폭이 없어 회귀 계산이 불가능합니다.")
        
    scale = ss_xy / ss_xx
    offset = mean_y - (scale * mean_x)
    r2 = (ss_xy / math.sqrt(ss_xx * ss_yy)) ** 2 if ss_yy != 0 else 0.0
    return scale, offset, r2

def apply_calibration(sensor_index, scale_pm25, offset_pm25):
    calib_data = {}
    if os.path.exists(CALIB_FILE):
        try:
            with open(CALIB_FILE, "r", encoding="utf-8") as f: 
                calib_data = json.load(f)
        except Exception: 
            pass
            
    idx_str = str(sensor_index)
    if idx_str not in calib_data: 
        calib_data[idx_str] = {"scale": [1.0, 1.0, 1.0], "offset": [0.0, 0.0, 0.0]}
    
    calib_data[idx_str]["scale"][1] = round(scale_pm25, 4)
    calib_data[idx_str]["offset"][1] = round(offset_pm25, 4)
    
    temp_file = CALIB_FILE + ".tmp"
    with open(temp_file, "w", encoding="utf-8") as f: 
        json.dump(calib_data, f, ensure_ascii=False, indent=2)
    os.replace(temp_file, CALIB_FILE)

def run(sensor_index=0, ref_file_name=None):
    ref_path = os.path.join(BASE_DIR, ref_file_name)
    print(f"\n==================================================")
    print(f"       [센서 {sensor_index + 1}] PM2.5 자동 교정 시작")
    print(f"==================================================")
    
    try: 
        sensor_hourly = get_sensor_hourly_averages(sensor_index)
    except Exception as e: 
        print(f"❌ DB 읽기 오류: {e}")
        return
    
    try: 
        ref_hourly = load_reference_file(ref_path)
    except Exception as e: 
        print(f"❌ 엑셀/CSV 파싱 오류: {e}")
        return
    
    matched_x, matched_y = [], []
    for time_bucket, s_val in sensor_hourly.items():
        if time_bucket in ref_hourly:
            matched_x.append(s_val)
            matched_y.append(ref_hourly[time_bucket])
            
    if len(matched_x) < 5: 
        print("❌ DB 데이터와 기준측정기 엑셀의 시간대가 일치하는 매칭 데이터가 5개 미만입니다.")
        return
        
    scale, offset, r2 = calculate_regression(matched_x, matched_y)
    print(f"\n▶ 매칭 시간대: {len(matched_x)}시간 (최소 요구 조건: {MIN_MINUTE_COUNT}분/시간)")
    print(f"▶ 획득 Scale: {scale:.4f}")
    print(f"▶ 획득 Offset: {offset:.4f}")
    print(f"▶ 결정계수 (R²): {r2:.4f}")
    
    apply_calibration(sensor_index, scale, offset)
    print("\n✔ sensor_calibration.json 파일에 새 보정값이 반영되었습니다.")
    
    print("==================================================")
    print("⚠️ 안내: 메인 모니터링 프로그램이 현재 실행 중인 경우,")
    print("   업데이트된 보정값을 화면에 적용하려면 메인 프로그램의")
    print("   [🔄 외부 보정값 파일 새로고침] 버튼을 눌러주세요.")
    print("==================================================")

if __name__ == "__main__":
    TARGET_SENSOR = 0 
    SPECIFIED_FILE = ""
    
    target_file = SPECIFIED_FILE if SPECIFIED_FILE else find_reference_file(BASE_DIR)
    if not target_file: 
        print("❌ 기준 데이터를 담은 엑셀(.xls, .xlsx) 또는 CSV 파일을 찾을 수 없습니다.")
        sys.exit(1)
        
    run(sensor_index=TARGET_SENSOR, ref_file_name=target_file)