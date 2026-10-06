import os
import json
import sqlite3
import math
import sys
import re
from datetime import datetime, timedelta

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
    
    file_year = datetime.now().year
    file_month = datetime.now().month
    match = re.search(r'(\d{2})(\d{2})\d{2}', filename)
    if match:
        file_year = 2000 + int(match.group(1)) 
        file_month = int(match.group(2))       
        
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
    
    for r_idx, row in df.head(15).iterrows():
        for c_idx, val in enumerate(row):
            val_str = str(val).upper().strip() if pd.notna(val) else ""
            
            if date_col_idx is None and any(k in val_str for k in ["날짜", "DATE", "일자", "조회일자", "일"]): 
                date_col_idx = c_idx
            elif time_col_idx is None and any(k in val_str for k in ["시간", "TIME"]): 
                time_col_idx = c_idx
                
            if ref_col_idx is None and any(k in val_str for k in REF_KEYWORDS): 
                ref_col_idx = c_idx
                
        if date_col_idx is not None and time_col_idx is not None and ref_col_idx is not None:
            header_row_idx = r_idx
            break

    if ref_col_idx is None: 
        raise ValueError("파일에서 'KOTITI' 등 기준측정기 열을 찾지 못했습니다.")
    if date_col_idx is None or time_col_idx is None:
        raise ValueError("파일에서 '일' 또는 '시간' 열을 찾지 못했습니다.")

    ref_data = {}
    
    for _, row in df.iloc[header_row_idx + 1:].iterrows():
        if pd.isna(row[date_col_idx]) or pd.isna(row[time_col_idx]) or pd.isna(row[ref_col_idx]):
            continue
            
        raw_date_str = str(row[date_col_idx]).strip()
        raw_time_str = str(row[time_col_idx]).strip()
        raw_val = row[ref_col_idx]
        
        day_match = re.search(r'\d+', raw_date_str)
        hour_match = re.search(r'\d+', raw_time_str)
        
        if day_match and hour_match:
            try:
                day = int(day_match.group())
                hour = int(hour_match.group())
                
                # 24시 처리 및 1시간 시차 교정 적용
                if hour == 24:
                    dt_obj = datetime(file_year, file_month, day, 0) + timedelta(days=1)
                else:
                    dt_obj = datetime(file_year, file_month, day, hour)
                    
                dt_obj -= timedelta(hours=1)
                norm_time = dt_obj.strftime("%Y-%m-%d %H:00:00")
                
                ref_data[norm_time] = float(raw_val)
            except (ValueError, TypeError):
                pass
                
    return ref_data

def calc_lin_reg(x_vals, y_vals):
    """단일 구간에 대한 선형 회귀 분석 (Scale, Offset 반환)"""
    n = len(x_vals)
    if n < 2: 
        return 1.0, 0.0 # 최소 2개 이상의 데이터가 없으면 계산 불가
    
    mean_x, mean_y = sum(x_vals) / n, sum(y_vals) / n
    ss_xx = sum((x - mean_x) ** 2 for x in x_vals)
    ss_xy = sum((x - mean_x) * (y - mean_y) for x, y in zip(x_vals, y_vals))
    
    if ss_xx == 0: 
        return 1.0, mean_y - mean_x
        
    scale = ss_xy / ss_xx
    offset = mean_y - (scale * mean_x)
    return scale, offset

def calculate_piecewise_regression(x_vals, y_vals):
    """
    🚀 [핵심 로직] 저농도/고농도 분리 보정값 최적화 알고리즘
    임계값(TH)을 2.0부터 10.0까지 변경해가며 가장 오차율(MAPE)이 적고 ±30% 합격률이 높은 5개의 파라미터를 찾습니다.
    """
    best_pass = -1
    best_mape = float('inf')
    best_params = (4.0, 1.0, 0.0, 1.0, 0.0) # (th, ls, lo, hs, ho) 기본값
    
    # 평가할 임계값 후보군 (2.0 ~ 10.0까지 0.5 단위로 탐색)
    thresholds = [th / 10.0 for th in range(20, 105, 5)]
    
    for th in thresholds:
        low_x, low_y, high_x, high_y = [], [], [], []
        
        for x, y in zip(x_vals, y_vals):
            if x < th:
                low_x.append(x)
                low_y.append(y)
            else:
                high_x.append(x)
                high_y.append(y)
        
        # 각 구간별 보정 계수 산출
        ls, lo = calc_lin_reg(low_x, low_y) if len(low_x) >= 2 else (1.0, 0.0)
        hs, ho = calc_lin_reg(high_x, high_y) if len(high_x) >= 2 else (1.0, 0.0)
        
        # 데이터 쏠림 현상 방어 (한쪽 구간에 데이터가 없으면 다른 쪽 값 차용)
        if len(low_x) < 2 and len(high_x) >= 2:
            ls, lo = hs, ho
        elif len(high_x) < 2 and len(low_x) >= 2:
            hs, ho = ls, lo
        elif len(low_x) < 2 and len(high_x) < 2:
            ls, lo = calc_lin_reg(x_vals, y_vals)
            hs, ho = ls, lo
            
        # 성능 평가 (30% 이내 통과 개수 및 평균 오차율)
        pass_count = 0
        errors = []
        for x, y in zip(x_vals, y_vals):
            pred = (x * ls + lo) if x < th else (x * hs + ho)
            err = abs((pred - y) / y) * 100
            errors.append(err)
            if err <= 30.0:
                pass_count += 1
                
        mape = sum(errors) / len(errors) if errors else float('inf')
        
        # 더 나은 성능을 찾았을 경우 갱신
        if pass_count > best_pass or (pass_count == best_pass and mape < best_mape):
            best_pass = pass_count
            best_mape = mape
            best_params = (th, ls, lo, hs, ho)
            
    return best_params, best_pass, best_mape, len(x_vals)

def apply_calibration(sensor_index, best_params):
    th, ls, lo, hs, ho = best_params
    
    # 1. JSON 파일 덮어쓰기 (메인 프로그램 실시간 연동용)
    calib_data = {}
    if os.path.exists(CALIB_FILE):
        try:
            with open(CALIB_FILE, "r", encoding="utf-8") as f: 
                calib_data = json.load(f)
        except Exception: 
            pass
            
    idx_str = str(sensor_index)
    
    # 이전 버전의 단일 scale/offset 포맷이거나 없으면 새 구간별 포맷으로 초기화
    if idx_str not in calib_data or "threshold" not in calib_data[idx_str]: 
        calib_data[idx_str] = {
            "threshold": [4.0, 4.0, 4.0],
            "low_scale": [1.0, 1.0, 1.0], "low_offset": [0.0, 0.0, 0.0],
            "high_scale": [1.0, 1.0, 1.0], "high_offset": [0.0, 0.0, 0.0]
        }
    
    # PM2.5 (인덱스 1) 위치의 값만 교체
    calib_data[idx_str]["threshold"][1] = round(th, 2)
    calib_data[idx_str]["low_scale"][1] = round(ls, 4)
    calib_data[idx_str]["low_offset"][1] = round(lo, 4)
    calib_data[idx_str]["high_scale"][1] = round(hs, 4)
    calib_data[idx_str]["high_offset"][1] = round(ho, 4)
    
    temp_file = CALIB_FILE + ".tmp"
    with open(temp_file, "w", encoding="utf-8") as f: 
        json.dump(calib_data, f, ensure_ascii=False, indent=2)
    os.replace(temp_file, CALIB_FILE)

    # 🚀 2. CSV에 도출된 보정값 누적 기록 남기기 (이력 관리용 - 컬럼 확장됨)
    history_file = os.path.join(BASE_DIR, "Data", "calibration_history.csv")
    file_exists = os.path.exists(history_file)
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    try:
        with open(history_file, "a", encoding="utf-8-sig") as f:
            if not file_exists:
                f.write("적용일시,센서번호,TH,Low_Scale,Low_Offset,High_Scale,High_Offset\n")
            f.write(f"{timestamp},센서 {sensor_index + 1},{th:.1f},{ls:.4f},{lo:.4f},{hs:.4f},{ho:.4f}\n")
    except Exception as e:
        print(f"⚠️ 히스토리 CSV 저장 중 오류 발생: {e}")

def run(sensor_index=0, ref_file_name=None):
    ref_path = os.path.join(BASE_DIR, ref_file_name)
    print(f"\n==================================================")
    print(f"       [센서 {sensor_index + 1}] PM2.5 구간별 보정 최적화")
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
            r_val = ref_hourly[time_bucket]
            
            # 기준기 값이 5 이상인 경우만 최적화에 사용
            if r_val >= 5:
                matched_x.append(s_val)
                matched_y.append(r_val)
                
    if len(matched_x) < 5: 
        print("❌ 유효한 매칭 데이터가 5개 미만입니다.")
        return
        
    best_params, pass_count, mape, total_count = calculate_piecewise_regression(matched_x, matched_y)
    th, ls, lo, hs, ho = best_params
    
    print(f"▶ 유효 매칭 시간대: {total_count}시간 (기준기 5 ㎍/㎥ 미만 제외)")
    print(f"▶ 탐색된 최적 임계값(TH) : {th:.1f}")
    print(f"▶ 저농도(<TH) 보정식   : (원본 × {ls:.4f}) + {lo:.4f}")
    print(f"▶ 고농도(>=TH) 보정식  : (원본 × {hs:.4f}) + {ho:.4f}")
    print(f"▶ 예상 검사 결과        : 합격 {pass_count}건 / {total_count}건 (평균오차율 {mape:.1f}%)")
    
    apply_calibration(sensor_index, best_params)
    print("\n✔ JSON 파일 및 CSV 히스토리에 새 보정값이 반영되었습니다.")
    
    print("==================================================")
    print("⚠️ 안내: 메인 모니터링 프로그램이 현재 실행 중인 경우,")
    print("   업데이트된 보정값을 화면에 적용하려면 메인 프로그램의")
    print("   [🔄 외부 보정값 파일 새로고침] 버튼을 눌러주세요.")
    print("==================================================")

if __name__ == "__main__":
    SPECIFIED_FILE = ""
    
    target_file = SPECIFIED_FILE if SPECIFIED_FILE else find_reference_file(BASE_DIR)
    if not target_file: 
        print("❌ 기준 데이터를 담은 엑셀(.xls, .xlsx) 또는 CSV 파일을 찾을 수 없습니다.")
        sys.exit(1)
        
    for i in range(4):
        run(sensor_index=i, ref_file_name=target_file)