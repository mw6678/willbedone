import os
import json
import sqlite3
import csv
import math
import sys
from datetime import datetime
import openpyxl

# ============================================================
# 설정 (경로 및 컬럼 키워드)
# ============================================================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE_DIR, "Data", "dust_measurement.db")
CALIB_FILE = os.path.join(BASE_DIR, "Data", "sensor_calibration.json")

# 1시간(60분) 중 최소 몇 분 이상 정상 수집되어야 유효한 데이터로 볼 것인가
MIN_MINUTE_COUNT = 45

# KTL/KOTITI 엑셀 헤더에서 찾을 열 이름 키워드
TIME_KEYWORDS = ["일시", "시간", "날짜", "측정일시", "DATE", "TIME"]
REF_KEYWORDS = ["기준기", "기준값", "기준농도", "REFERENCE", "BAM", "GRIMM", "PM2.5", "PM25"]


# ============================================================
# 1. 기준기 파일 자동 탐색 함수
# ============================================================
def find_reference_file(base_dir):
    """
    현재 폴더에서 기준기 엑셀 또는 CSV 파일을 자동 검색합니다.
    - 임시 파일(~$...) 및 프로그램 내보내기 파일(Dust_log_...) 자동 제외
    """
    ignore_prefixes = ("~$", "Dust_log_")
    valid_extensions = (".xlsx", ".csv", ".xls")

    try:
        candidates = [
            f for f in os.listdir(base_dir)
            if any(f.lower().endswith(ext) for ext in valid_extensions)
            and not any(f.startswith(p) for p in ignore_prefixes)
            and os.path.isfile(os.path.join(base_dir, f))
        ]
    except OSError as e:
        print(f"❌ 디렉터리 탐색 오류: {e}")
        return None

    if not candidates:
        return None

    # 파일이 1개만 있으면 자동 선택
    if len(candidates) == 1:
        print(f"📁 기준기 파일을 자동으로 감지했습니다: '{candidates[0]}'")
        return candidates[0]

    # 파일이 여러 개 있으면 사용자에게 선택 요청
    print("\n📁 여러 개의 비교군(기준기) 파일이 발견되었습니다:")
    for idx, fname in enumerate(candidates, 1):
        print(f"   [{idx}] {fname}")

    while True:
        try:
            choice = input("\n사용할 파일 번호를 입력하세요: ").strip()
            sel_idx = int(choice) - 1
            if 0 <= sel_idx < len(candidates):
                return candidates[sel_idx]
            print(f"1부터 {len(candidates)} 사이의 번호를 입력해주세요.")
        except (ValueError, EOFError):
            print("올바른 숫자를 입력해주세요.")


# ============================================================
# 2. SQLite에서 센서의 1시간 평균값 집계
# ============================================================
def get_sensor_hourly_averages(sensor_index=0):
    if not os.path.exists(DB_PATH):
        raise FileNotFoundError(f"DB 파일을 찾을 수 없습니다: {DB_PATH}")

    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()

    query = """
        SELECT 
            strftime('%Y-%m-%d %H:00:00', measured_at) AS hour_bucket,
            AVG(pm25) AS avg_pm25,
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


# ============================================================
# 3. 기준기 데이터 읽기 (.xlsx 및 .csv 자동 판별)
# ============================================================
def parse_datetime_value(val):
    """다양한 날짜/시간 포맷을 'YYYY-MM-DD HH:00:00' 형태로 통일"""
    if isinstance(val, datetime):
        return val.strftime("%Y-%m-%d %H:00:00")

    val_str = str(val).strip()
    formats = [
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%d %H:%M",
        "%Y.%m.%d %H:%M:%S",
        "%Y.%m.%d %H:%M",
        "%Y/%m/%d %H:%M:%S",
        "%Y/%m/%d %H:%M"
    ]
    for fmt in formats:
        try:
            dt = datetime.strptime(val_str[:16], fmt[:len(val_str[:16])])
            return dt.strftime("%Y-%m-%d %H:00:00")
        except ValueError:
            continue
    return None


def load_reference_excel(file_path):
    wb = openpyxl.load_workbook(file_path, data_only=True)
    ws = wb.active  # 첫 번째 시트 사용

    time_col_idx = None
    ref_col_idx = None
    data_start_row = None

    # 헤더 행 탐색 (상단 1~15행 사이 탐색)
    for r in range(1, min(16, ws.max_row + 1)):
        row_vals = [str(ws.cell(r, c).value or "").strip() for c in range(1, ws.max_column + 1)]

        for c_idx, val in enumerate(row_vals, 1):
            val_upper = val.upper()
            if time_col_idx is None and any(k in val_upper for k in TIME_KEYWORDS):
                time_col_idx = c_idx
            if ref_col_idx is None and any(k in val_upper for k in REF_KEYWORDS):
                ref_col_idx = c_idx

        if time_col_idx and ref_col_idx:
            data_start_row = r + 1
            break

    if not time_col_idx or not ref_col_idx:
        raise ValueError("엑셀에서 '시간' 열 또는 '기준기' 열을 찾지 못했습니다. 열 이름을 확인해주세요.")

    ref_data = {}
    for r in range(data_start_row, ws.max_row + 1):
        raw_time = ws.cell(r, time_col_idx).value
        raw_val = ws.cell(r, ref_col_idx).value

        if raw_time is None or raw_val is None:
            continue

        norm_time = parse_datetime_value(raw_time)
        if not norm_time:
            continue

        try:
            ref_data[norm_time] = float(raw_val)
        except (ValueError, TypeError):
            continue

    return ref_data


def load_reference_csv(file_path):
    ref_data = {}
    with open(file_path, "r", encoding="utf-8-sig") as f:
        reader = csv.reader(f)
        header = next(reader, None)
        for row in reader:
            if len(row) < 2:
                continue
            norm_time = parse_datetime_value(row[0])
            if norm_time:
                try:
                    ref_data[norm_time] = float(row[1])
                except ValueError:
                    continue
    return ref_data


def load_reference_file(file_path):
    if file_path.lower().endswith((".xlsx", ".xls")):
        return load_reference_excel(file_path)
    elif file_path.lower().endswith(".csv"):
        return load_reference_csv(file_path)
    else:
        raise ValueError("지원하지 않는 파일 형식입니다. (.xlsx, .xls, .csv)")


# ============================================================
# 4. 최소자승법 선형회귀 및 보정값 저장
# ============================================================
def calculate_regression(x_vals, y_vals):
    n = len(x_vals)
    mean_x = sum(x_vals) / n
    mean_y = sum(y_vals) / n

    ss_xx = sum((x - mean_x) ** 2 for x in x_vals)
    ss_yy = sum((y - mean_y) ** 2 for y in y_vals)
    ss_xy = sum((x - mean_x) * (y - mean_y) for x, y in zip(x_vals, y_vals))

    if ss_xx == 0:
        raise ValueError("센서 값의 변동폭이 없어 기울기를 계산할 수 없습니다.")

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
            calib_data = {}

    idx_str = str(sensor_index)
    if idx_str not in calib_data:
        calib_data[idx_str] = {"scale": [1.0, 1.0, 1.0], "offset": [0.0, 0.0, 0.0]}

    # PM2.5 계수 업데이트 (인덱스 1번: [PM10, PM2.5, PM1.0])
    calib_data[idx_str]["scale"][1] = round(scale_pm25, 4)
    calib_data[idx_str]["offset"][1] = round(offset_pm25, 4)

    temp_file = CALIB_FILE + ".tmp"
    with open(temp_file, "w", encoding="utf-8") as f:
        json.dump(calib_data, f, ensure_ascii=False, indent=2)
    os.replace(temp_file, CALIB_FILE)


# ============================================================
# 실행 메인
# ============================================================
def run(sensor_index=0, ref_file_name=None):
    ref_path = os.path.join(BASE_DIR, ref_file_name)
    if not os.path.exists(ref_path):
        print(f"❌ '{ref_file_name}' 파일을 찾을 수 없습니다.")
        return

    print(f"\n==================================================")
    print(f"       [센서 {sensor_index + 1}] PM2.5 자동 교정 시작")
    print(f"==================================================")

    print("1. DB에서 센서 1시간 단위 평균 집계 중...")
    try:
        sensor_hourly = get_sensor_hourly_averages(sensor_index)
    except FileNotFoundError as e:
        print(f"❌ {e}")
        return

    print(f"   -> 센서 1시간 평균 데이터: 총 {len(sensor_hourly)}개 시간대 추출됨")

    print(f"2. 기준기 데이터('{ref_file_name}') 읽는 중...")
    try:
        ref_hourly = load_reference_file(ref_path)
        print(f"   -> 기준기 1시간 데이터: 총 {len(ref_hourly)}개 시간대 추출됨")
    except Exception as e:
        print(f"❌ 파일 읽기 오류: {e}")
        return

    matched_x, matched_y = [], []
    for time_bucket, s_val in sensor_hourly.items():
        if time_bucket in ref_hourly:
            matched_x.append(s_val)
            matched_y.append(ref_hourly[time_bucket])

    print(f"3. 센서-기준기 동시간대 매칭: 총 {len(matched_x)}개 시간대 일치")
    if len(matched_x) < 5:
        print("❌ 일치하는 데이터가 너무 적어 교정을 진행할 수 없습니다. (최소 5시간 이상 필요)")
        return

    scale, offset, r2 = calculate_regression(matched_x, matched_y)

    print("\n" + "=" * 48)
    print(f"       [센서 {sensor_index + 1} PM2.5 교정 결과]")
    print("=" * 48)
    print(f" ▶ 매칭된 시간대 수 : {len(matched_x)}시간")
    print(f" ▶ Scale (기울기)   : {scale:.4f}")
    print(f" ▶ Offset (절편)    : {offset:.4f}")
    print(f" ▶ 결정계수 (R²)   : {r2:.4f} " + ("(유효함)" if r2 >= 0.8 else "(주의: 상관성 낮음)"))
    print(f" ▶ 적용 공식: (Raw값 × {scale:.4f}) + ({offset:.4f})")
    print("=" * 48)

    apply_calibration(sensor_index, scale, offset)
    print(f"\n✔ 'Data/sensor_calibration.json'에 계수가 자동 저장되었습니다.")
    print("✔ [안내] 모니터링 프로그램(dust_monitor_optimized_v2.py)을 재시작하면 새 보정값이 반영됩니다.\n")


if __name__ == "__main__":
    # 교정 대상 센서 번호 (0: 센서1, 1: 센서2, 2: 센서3, 3: 센서4)
    TARGET_SENSOR = 0

    # 특정 파일명을 고정 지정하려면 따옴표 안에 파일명 입력 (비워두면 자동 탐색)
    SPECIFIED_FILE = ""

    if SPECIFIED_FILE:
        target_file = SPECIFIED_FILE
    else:
        target_file = find_reference_file(BASE_DIR)

    if not target_file:
        print("❌ 같은 폴더에 비교군 엑셀(.xlsx) 또는 CSV(.csv) 파일이 존재하지 않습니다.")
        sys.exit(1)

    run(sensor_index=TARGET_SENSOR, ref_file_name=target_file)