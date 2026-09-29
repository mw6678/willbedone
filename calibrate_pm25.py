import os
import json
import sqlite3
import csv
import math
import sys
from datetime import datetime
import openpyxl

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE_DIR, "Data", "dust_measurement.db")
CALIB_FILE = os.path.join(BASE_DIR, "Data", "sensor_calibration.json")

MIN_MINUTE_COUNT = 60
TIME_KEYWORDS = ["일시", "시간", "날짜", "측정일시", "DATE", "TIME"]
REF_KEYWORDS = ["기준기", "기준값", "기준농도", "REFERENCE", "BAM", "GRIMM", "PM2.5", "PM25"]

def find_reference_file(base_dir):
    ignore_prefixes = ("~$", "Dust_log_")
    valid_extensions = (".xlsx", ".csv", ".xls")
    try:
        candidates = [f for f in os.listdir(base_dir) if any(f.lower().endswith(ext) for ext in valid_extensions) and not any(f.startswith(p) for p in ignore_prefixes) and os.path.isfile(os.path.join(base_dir, f))]
    except OSError as e: return None
    if not candidates: return None
    if len(candidates) == 1: return candidates[0]
    print("\n📁 여러 개의 비교군 파일 발견:")
    for idx, fname in enumerate(candidates, 1): print(f"   [{idx}] {fname}")
    while True:
        try:
            choice = input("\n사용할 파일 번호를 입력하세요: ").strip()
            sel_idx = int(choice) - 1
            if 0 <= sel_idx < len(candidates): return candidates[sel_idx]
        except (ValueError, EOFError): pass

def get_sensor_hourly_averages(sensor_index=0):
    if not os.path.exists(DB_PATH): raise FileNotFoundError(f"DB 파일을 찾을 수 없습니다: {DB_PATH}")
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    
    # 여기서 기존 'AVG(pm25)'를 'AVG(raw_pm25)'로 수정하여 순수 원본값만 가져옵니다.
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

def parse_datetime_value(val):
    if isinstance(val, datetime): return val.strftime("%Y-%m-%d %H:00:00")
    val_str = str(val).strip()
    formats = ["%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%Y.%m.%d %H:%M:%S", "%Y.%m.%d %H:%M", "%Y/%m/%d %H:%M:%S", "%Y/%m/%d %H:%M"]
    for fmt in formats:
        try: return datetime.strptime(val_str[:16], fmt[:len(val_str[:16])]).strftime("%Y-%m-%d %H:00:00")
        except ValueError: continue
    return None

def load_reference_excel(file_path):
    wb = openpyxl.load_workbook(file_path, data_only=True)
    ws = wb.active
    time_col_idx, ref_col_idx, data_start_row = None, None, None
    for r in range(1, min(16, ws.max_row + 1)):
        row_vals = [str(ws.cell(r, c).value or "").strip() for c in range(1, ws.max_column + 1)]
        for c_idx, val in enumerate(row_vals, 1):
            if time_col_idx is None and any(k in val.upper() for k in TIME_KEYWORDS): time_col_idx = c_idx
            if ref_col_idx is None and any(k in val.upper() for k in REF_KEYWORDS): ref_col_idx = c_idx
        if time_col_idx and ref_col_idx:
            data_start_row = r + 1; break
    if not time_col_idx or not ref_col_idx: raise ValueError("엑셀에서 열을 찾지 못했습니다.")
    ref_data = {}
    for r in range(data_start_row, ws.max_row + 1):
        raw_time, raw_val = ws.cell(r, time_col_idx).value, ws.cell(r, ref_col_idx).value
        if raw_time is None or raw_val is None: continue
        norm_time = parse_datetime_value(raw_time)
        if norm_time:
            try: ref_data[norm_time] = float(raw_val)
            except ValueError: pass
    return ref_data

def load_reference_csv(file_path):
    ref_data = {}
    with open(file_path, "r", encoding="utf-8-sig") as f:
        reader = csv.reader(f)
        next(reader, None)
        for row in reader:
            if len(row) < 2: continue
            norm_time = parse_datetime_value(row[0])
            if norm_time:
                try: ref_data[norm_time] = float(row[1])
                except ValueError: pass
    return ref_data

def load_reference_file(file_path):
    if file_path.lower().endswith((".xlsx", ".xls")): return load_reference_excel(file_path)
    elif file_path.lower().endswith(".csv"): return load_reference_csv(file_path)
    else: raise ValueError("지원하지 않는 파일 형식입니다.")

def calculate_regression(x_vals, y_vals):
    n = len(x_vals)
    mean_x, mean_y = sum(x_vals) / n, sum(y_vals) / n
    ss_xx = sum((x - mean_x) ** 2 for x in x_vals)
    ss_yy = sum((y - mean_y) ** 2 for y in y_vals)
    ss_xy = sum((x - mean_x) * (y - mean_y) for x, y in zip(x_vals, y_vals))
    if ss_xx == 0: raise ValueError("센서 값의 변동폭이 없습니다.")
    scale = ss_xy / ss_xx
    offset = mean_y - (scale * mean_x)
    r2 = (ss_xy / math.sqrt(ss_xx * ss_yy)) ** 2 if ss_yy != 0 else 0.0
    return scale, offset, r2

def apply_calibration(sensor_index, scale_pm25, offset_pm25):
    calib_data = {}
    if os.path.exists(CALIB_FILE):
        try:
            with open(CALIB_FILE, "r", encoding="utf-8") as f: calib_data = json.load(f)
        except Exception: pass
    idx_str = str(sensor_index)
    if idx_str not in calib_data: calib_data[idx_str] = {"scale": [1.0, 1.0, 1.0], "offset": [0.0, 0.0, 0.0]}
    
    calib_data[idx_str]["scale"][1] = round(scale_pm25, 4)
    calib_data[idx_str]["offset"][1] = round(offset_pm25, 4)
    temp_file = CALIB_FILE + ".tmp"
    with open(temp_file, "w", encoding="utf-8") as f: json.dump(calib_data, f, ensure_ascii=False, indent=2)
    os.replace(temp_file, CALIB_FILE)

def run(sensor_index=0, ref_file_name=None):
    ref_path = os.path.join(BASE_DIR, ref_file_name)
    print(f"\n==================================================\n       [센서 {sensor_index + 1}] PM2.5 자동 교정 시작\n==================================================")
    try: sensor_hourly = get_sensor_hourly_averages(sensor_index)
    except Exception as e: print(f"❌ {e}"); return
    
    try: ref_hourly = load_reference_file(ref_path)
    except Exception as e: print(f"❌ 파일 읽기 오류: {e}"); return
    
    matched_x, matched_y = [], []
    for time_bucket, s_val in sensor_hourly.items():
        if time_bucket in ref_hourly:
            matched_x.append(s_val); matched_y.append(ref_hourly[time_bucket])
            
    if len(matched_x) < 5: print("❌ 일치하는 데이터가 너무 적습니다."); return
    scale, offset, r2 = calculate_regression(matched_x, matched_y)
    print(f"\n▶ 매칭 시간대: {len(matched_x)}시간\n▶ Scale: {scale:.4f}\n▶ Offset: {offset:.4f}\n▶ 결정계수 (R²): {r2:.4f}")
    apply_calibration(sensor_index, scale, offset)
    print("✔ 새 보정값이 반영되었습니다.")

if __name__ == "__main__":
    TARGET_SENSOR = 0
    SPECIFIED_FILE = ""
    target_file = SPECIFIED_FILE if SPECIFIED_FILE else find_reference_file(BASE_DIR)
    if not target_file: sys.exit(1)
    run(sensor_index=TARGET_SENSOR, ref_file_name=target_file)