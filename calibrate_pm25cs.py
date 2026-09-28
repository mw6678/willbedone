import os
import json
import sqlite3
import csv
import math
from datetime import datetime

# ============================================================
# 설정 (경로 및 기준)
# ============================================================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE_DIR, "Data", "dust_measurement.db")
CALIB_FILE = os.path.join(BASE_DIR, "Data", "sensor_calibration.json")

# 1시간(60분) 중 최소 몇 분 이상 데이터가 있어야 유효한 1시간 데이터로 인정할지 (예: 45분 이상)
MIN_MINUTE_COUNT = 45


# ============================================================
# 1. SQLite에서 1시간 단위 평균 데이터 추출
# ============================================================
def get_sensor_hourly_averages(sensor_index=0, start_date=None, end_date=None):
    """
    DB에서 1분 데이터를 1시간 단위로 그룹화하여 평균(AVG)을 계산합니다.
    반환값: { "YYYY-MM-DD HH:00:00": pm25_avg, ... }
    """
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
    """
    params = [sensor_index]

    if start_date:
        query += " AND measured_at >= ?"
        params.append(f"{start_date} 00:00:00")
    if end_date:
        query += " AND measured_at <= ?"
        params.append(f"{end_date} 23:59:59")

    query += """
        GROUP BY hour_bucket
        HAVING data_count >= ?
        ORDER BY hour_bucket ASC
    """
    params.append(MIN_MINUTE_COUNT)

    cursor.execute(query, params)
    rows = cursor.fetchall()
    conn.close()

    # { 시간: 평균값 } 형태로 변환
    hourly_data = {row[0]: round(row[1], 2) for row in rows}
    return hourly_data


# ============================================================
# 2. 기준기(Reference) 1시간 데이터 읽기
# ============================================================
def load_reference_data(ref_file_path):
    """
    기준기 CSV 파일을 읽어옵니다.
    CSV 형식 예시:
        시간,기준값
        2026-09-28 09:00:00,23.5
        2026-09-28 10:00:00,25.1
    """
    if not os.path.exists(ref_file_path):
        raise FileNotFoundError(f"기준기 파일을 찾을 수 없습니다: {ref_file_path}")

    ref_data = {}
    with open(ref_file_path, "r", encoding="utf-8-sig") as f:
        reader = csv.reader(f)
        header = next(reader, None)  # 헤더 스킵

        for row in reader:
            if len(row) < 2:
                continue
            time_str = row[0].strip()
            try:
                # 시간 포맷 정규화 (YYYY-MM-DD HH:00:00)
                dt = datetime.strptime(time_str[:13], "%Y-%m-%d %H")
                norm_time = dt.strftime("%Y-%m-%d %H:00:00")
                ref_val = float(row[1].strip())
                ref_data[norm_time] = ref_val
            except ValueError:
                continue

    return ref_data


# ============================================================
# 3. 최소자승법(OLS) 선형 회귀분석 및 교정 계수 산출
# ============================================================
def calculate_calibration_factors(x_vals, y_vals):
    """
    x: 센서 1시간 평균값 리스트
    y: 기준기 1시간 평균값 리스트
    회귀식: y = a * x + b
    - Scale (a): 기울기
    - Offset (b): 절편
    - R²: 결정계수
    """
    n = len(x_vals)
    if n < 5:
        raise ValueError(f"데이터 쌍이 너무 적습니다 ({n}개). 최소 5개 이상의 1시간 데이터가 필요합니다.")

    mean_x = sum(x_vals) / n
    mean_y = sum(y_vals) / n

    ss_xx = sum((x - mean_x) ** 2 for x in x_vals)
    ss_yy = sum((y - mean_y) ** 2 for y in y_vals)
    ss_xy = sum((x - mean_x) * (y - mean_y) for x, y in zip(x_vals, y_vals))

    if ss_xx == 0:
        raise ValueError("센서 데이터에 변화량이 없어 기울기를 계산할 수 없습니다.")

    # 1) Scale (기울기 a) & Offset (절편 b)
    scale = ss_xy / ss_xx
    offset = mean_y - (scale * mean_x)

    # 2) 결정계수 R² 산출
    if ss_yy == 0:
        r_squared = 1.0 if ss_xx == 0 else 0.0
    else:
        r = ss_xy / math.sqrt(ss_xx * ss_yy)
        r_squared = r ** 2

    return scale, offset, r_squared


# ============================================================
# 4. 프로그램 설정 파일(sensor_calibration.json)에 적용
# ============================================================
def apply_calibration_to_json(sensor_index, scale_pm25, offset_pm25):
    """
    기존 sensor_calibration.json을 열어 해당 센서의 PM2.5(인덱스 1번)만 갱신합니다.
    인덱스 구조: [0]: PM10, [1]: PM2.5, [2]: PM1.0
    """
    calib_data = {}
    if os.path.exists(CALIB_FILE):
        try:
            with open(CALIB_FILE, "r", encoding="utf-8") as f:
                calib_data = json.load(f)
        except Exception:
            calib_data = {}

    idx_str = str(sensor_index)
    if idx_str not in calib_data:
        calib_data[idx_str] = {
            "scale": [1.0, 1.0, 1.0],
            "offset": [0.0, 0.0, 0.0]
        }

    # PM2.5 계수 업데이트 (소수점 4자리 반올림)
    calib_data[idx_str]["scale"][1] = round(scale_pm25, 4)
    calib_data[idx_str]["offset"][1] = round(offset_pm25, 4)

    # 안전하게 임시 파일 생성 후 치환 저장
    temp_file = CALIB_FILE + ".tmp"
    with open(temp_file, "w", encoding="utf-8") as f:
        json.dump(calib_data, f, ensure_ascii=False, indent=2)
    os.replace(temp_file, CALIB_FILE)


# ============================================================
# 실행 메인 루틴
# ============================================================
def run_calibration(sensor_index=0, ref_csv_path="reference_pm25.csv"):
    print(f"\n==================================================")
    print(f"       [센서 {sensor_index + 1}] PM2.5 교정 계수 산출 작업")
    print(f"==================================================")

    # 1. DB에서 센서 1시간 평균 가져오기
    print("1. DB에서 1시간 단위 평균치 집계 중...")
    sensor_hourly = get_sensor_hourly_averages(sensor_index=sensor_index)
    print(f"   -> 센서 1시간 평균 데이터 추출 완료: 총 {len(sensor_hourly)}개 시간대")

    # 2. 기준기 1시간 평균 가져오기
    print(f"2. 기준기 데이터({ref_csv_path}) 로딩 중...")
    ref_hourly = load_reference_data(ref_csv_path)
    print(f"   -> 기준기 데이터 추출 완료: 총 {len(ref_hourly)}개 시간대")

    # 3. 시간대 매칭 (센서 시간 == 기준기 시간)
    matched_x = []  # 센서 PM2.5 (x)
    matched_y = []  # 기준기 PM2.5 (y)

    for time_bucket, s_val in sensor_hourly.items():
        if time_bucket in ref_hourly:
            matched_x.append(s_val)
            matched_y.append(ref_hourly[time_bucket])

    print(f"3. 시간 동기화 매칭 완료: 총 {len(matched_x)}개 시간대 매칭 성공")
    if len(matched_x) < 5:
        print("❌ 오류: 기준기와 센서 간 일치하는 시간대 데이터가 부족합니다.")
        return

    # 4. 회귀 계수 산출
    scale, offset, r2 = calculate_calibration_factors(matched_x, matched_y)

    print("\n---------------- [교정 분석 결과] ----------------")
    print(f" ▶ 매칭 샘플 수: {len(matched_x)}개 (시간)")
    print(f" ▶ 산출된 Scale (기울기) : {scale:.4f}")
    print(f" ▶ 산출된 Offset (절편)  : {offset:.4f}")
    print(f" ▶ 결정계수 (R²)        : {r2:.4f} " + ("(적합함)" if r2 >= 0.8 else "(주의: 상관관계 낮음)"))
    print(f" ▶ 보정 공식: 보정값 = (Raw값 × {scale:.4f}) + ({offset:.4f})")
    print("--------------------------------------------------")

    # 5. 설정 파일 자동 저장
    apply_calibration_to_json(sensor_index, scale, offset)
    print(f"✔ [성공] '{CALIB_FILE}'에 센서 {sensor_index + 1}의 PM2.5 보정 계수가 적용되었습니다.")
    print("✔ 모니터링 프로그램이 실행 중이라면 [센서 보정 설정] 창을 열어 즉시 확인 가능합니다.\n")


if __name__ == "__main__":
    # 교정할 센서 번호 (0: 센서1, 1: 센서2, 2: 센서3, 3: 센서4)
    TARGET_SENSOR_INDEX = 0

    # 기준기 데이터 파일명
    REFERENCE_CSV_FILE = os.path.join(BASE_DIR, "reference_pm25.csv")

    # 테스트용 기준기 샘플 파일이 없는 경우 자동 생성 안내
    if not os.path.exists(REFERENCE_CSV_FILE):
        print(f"'{REFERENCE_CSV_FILE}' 파일이 없어 기본 템플릿 파일을 생성합니다.")
        with open(REFERENCE_CSV_FILE, "w", encoding="utf-8-sig", newline="") as f:
            writer = csv.writer(f)
            writer.writerow(["시간", "기준값"])
            writer.writerow(["2026-09-28 09:00:00", "25.0"])
            writer.writerow(["2026-09-28 10:00:00", "32.4"])
            writer.writerow(["2026-09-28 11:00:00", "18.2"])
        print("생성된 CSV 파일에 실제 기준기 데이터를 입력 후 다시 실행해주세요.")
    else:
        run_calibration(sensor_index=TARGET_SENSOR_INDEX, ref_csv_path=REFERENCE_CSV_FILE)