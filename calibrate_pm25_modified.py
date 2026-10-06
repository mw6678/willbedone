import os
import json
import sqlite3
import math
import sys
import re
from datetime import datetime

try:
    import pandas as pd
except ImportError:
    print("❌ pandas가 설치되어 있지 않습니다.")
    print("터미널: pip install pandas openpyxl xlrd")
    sys.exit(1)


# ============================================================
# 설정
# ============================================================

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# DB 위치
DB_PATH = os.path.join(BASE_DIR, "Data", "dust_measurement.db")

# 보정계수 저장 위치
CALIB_FILE = os.path.join(BASE_DIR, "Data", "sensor_calibration.json")

# 기준측정기 파일.
# 빈 문자열이면 BASE_DIR에서 자동으로 xlsx/xls/csv를 찾습니다.
SPECIFIED_FILE = ""

# 1시간 평균을 만들 때 최소 데이터 개수
# 기존 코드의 45분 조건을 유지합니다.
MIN_MINUTE_COUNT = 45

# 회귀에 사용할 최소 매칭 시간
MIN_MATCHED_HOURS = 5

# 센서 개수
SENSOR_COUNT = 4

# 기준측정기 열을 찾기 위한 키워드
REF_KEYWORDS = [
    "기준기",
    "기준값",
    "기준농도",
    "REFERENCE",
    "BAM",
    "GRIMM",
    "PM2.5",
    "PM25",
    "KOTITI",
    "SPREAD",
]


# ============================================================
# 공통 함수
# ============================================================

def find_reference_file(base_dir):
    """기준측정기 Excel/CSV 파일을 자동으로 찾습니다."""
    ignore_prefixes = ("~$", "Dust_log_")
    valid_extensions = (".xlsx", ".csv", ".xls")

    try:
        candidates = [
            f for f in os.listdir(base_dir)
            if f.lower().endswith(valid_extensions)
            and not any(f.startswith(p) for p in ignore_prefixes)
            and os.path.isfile(os.path.join(base_dir, f))
        ]
    except OSError as e:
        print(f"❌ 폴더를 읽을 수 없습니다: {e}")
        return None

    if not candidates:
        return None

    if len(candidates) == 1:
        return candidates[0]

    print("\n📁 여러 개의 기준측정기 파일이 발견되었습니다.")
    for idx, fname in enumerate(candidates, 1):
        print(f"   [{idx}] {fname}")

    while True:
        try:
            choice = input("\n사용할 파일 번호: ").strip()
            sel_idx = int(choice) - 1

            if 0 <= sel_idx < len(candidates):
                return candidates[sel_idx]

        except (ValueError, EOFError):
            pass

        print("❌ 올바른 번호를 입력하세요.")


def detect_db_schema(conn):
    """
    DB 구조를 확인하여 현재 코드에서 사용하는
    measurements / measured_at / raw_pm25 / sensor_index / status
    컬럼이 실제로 존재하는지 검사합니다.
    """
    cursor = conn.cursor()

    tables = cursor.execute(
        "SELECT name FROM sqlite_master "
        "WHERE type='table' AND name NOT LIKE 'sqlite_%'"
    ).fetchall()

    table_names = [row[0] for row in tables]

    if "measurements" not in table_names:
        raise ValueError(
            "DB에 'measurements' 테이블이 없습니다.\n"
            f"발견된 테이블: {table_names}"
        )

    columns = cursor.execute("PRAGMA table_info(measurements)").fetchall()
    column_names = [row[1] for row in columns]

    required = [
        "measured_at",
        "raw_pm25",
        "sensor_index",
    ]

    missing = [c for c in required if c not in column_names]

    if missing:
        raise ValueError(
            "measurements 테이블에 필요한 컬럼이 없습니다.\n"
            f"누락: {missing}\n"
            f"현재 컬럼: {column_names}"
        )

    return column_names


# ============================================================
# 센서 데이터
# ============================================================

def get_sensor_hourly_averages(sensor_index=0):
    """
    DB의 raw_pm25를 1시간 평균으로 변환합니다.

    결과:
        {
            "2026-09-30 13:00:00": 25.31,
            ...
        }
    """
    if not os.path.exists(DB_PATH):
        raise FileNotFoundError(
            f"DB 파일을 찾을 수 없습니다:\n{DB_PATH}"
        )

    conn = sqlite3.connect(DB_PATH)

    try:
        columns = detect_db_schema(conn)

        # status 컬럼이 있으면 NORMAL만 사용
        status_condition = ""
        params = [sensor_index]

        if "status" in columns:
            status_condition = "AND status = 'NORMAL'"

        query = f"""
            SELECT
                strftime('%Y-%m-%d %H:00:00', measured_at) AS hour_bucket,
                AVG(CAST(raw_pm25 AS REAL)) AS avg_pm25,
                COUNT(*) AS data_count
            FROM measurements
            WHERE sensor_index = ?
              {status_condition}
              AND raw_pm25 IS NOT NULL
            GROUP BY hour_bucket
            HAVING data_count >= ?
            ORDER BY hour_bucket ASC
        """

        params.append(MIN_MINUTE_COUNT)

        rows = conn.execute(query, params).fetchall()

        result = {}

        for hour_bucket, avg_pm25, data_count in rows:
            if hour_bucket is None or avg_pm25 is None:
                continue

            result[hour_bucket] = {
                "value": float(avg_pm25),
                "count": int(data_count),
            }

        return result

    finally:
        conn.close()


# ============================================================
# 기준측정기 Excel
# ============================================================

def extract_file_year_month(filename):
    """
    파일명에서 YYMMDD 형태를 찾아 연/월을 추출합니다.

    예:
        기준측정기 데이터_260930-261001.xlsx
        -> 2026년 09월
    """
    match = re.search(r"(\d{2})(\d{2})\d{2}", filename)

    if match:
        return (
            2000 + int(match.group(1)),
            int(match.group(2)),
        )

    now = datetime.now()
    return now.year, now.month


def normalize_hour(value):
    """
    Excel의 시간 값을 HH:00:00 형태로 변환합니다.
    """
    if pd.isna(value):
        return None

    # datetime / Timestamp
    if isinstance(value, (pd.Timestamp, datetime)):
        return int(value.hour)

    # Excel에서 시간이 숫자로 들어오는 경우
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        number = float(value)

        # 0~1이면 Excel 시간 serial
        if 0 <= number < 1:
            return int(number * 24) % 24

        # 0~23이면 시간
        if 0 <= number <= 23:
            return int(number)

    text = str(value).strip()

    # "13:00", "13:00:00", "13시" 등
    match = re.search(r"(\d{1,2})", text)

    if match:
        hour = int(match.group(1))

        if 0 <= hour <= 23:
            return hour

    return None


def normalize_day(value):
    """Excel의 날짜/일 값을 일(day) 숫자로 변환합니다."""
    if pd.isna(value):
        return None

    if isinstance(value, (pd.Timestamp, datetime)):
        return int(value.day)

    text = str(value).strip()

    match = re.search(r"(\d{1,2})", text)

    if match:
        day = int(match.group(1))

        if 1 <= day <= 31:
            return day

    return None


def find_reference_column(df):
    """
    Excel 상단 영역에서 기준측정기/Spread 열을 찾습니다.
    """
    # 우선순위가 높은 키워드
    priority_keywords = [
        "SPREAD",
        "기준농도",
        "기준값",
        "기준기",
        "REFERENCE",
        "BAM",
        "GRIMM",
        "KOTITI",
        "PM2.5",
        "PM25",
    ]

    for keyword in priority_keywords:
        for c_idx, value in enumerate(df.iloc[:20].astype(str).values.flatten()):
            if keyword in value.upper():
                return c_idx

    return None


def find_header_columns(df):
    """
    Excel에서 날짜 / 시간 / 기준값 열을 자동 탐색합니다.
    """
    date_col_idx = None
    time_col_idx = None
    ref_col_idx = find_reference_column(df)
    header_row_idx = None

    # 상위 20행 검사
    for r_idx, row in df.head(20).iterrows():

        for c_idx, val in enumerate(row):
            if pd.isna(val):
                continue

            val_str = str(val).upper().strip()

            if date_col_idx is None:
                if any(
                    k in val_str
                    for k in ["날짜", "DATE", "일자", "조회일자"]
                ):
                    date_col_idx = c_idx

            if time_col_idx is None:
                if any(
                    k in val_str
                    for k in ["시간", "TIME"]
                ):
                    time_col_idx = c_idx

        if (
            date_col_idx is not None
            and time_col_idx is not None
            and ref_col_idx is not None
        ):
            header_row_idx = r_idx
            break

    return (
        date_col_idx,
        time_col_idx,
        ref_col_idx,
        header_row_idx,
    )


def load_reference_file(file_path):
    """
    기준측정기 Excel/CSV를 읽어서
    {
        "YYYY-MM-DD HH:00:00": 기준값
    }
    형태로 변환합니다.
    """
    if not os.path.exists(file_path):
        raise FileNotFoundError(
            f"기준측정기 파일을 찾을 수 없습니다:\n{file_path}"
        )

    ext = file_path.lower().split(".")[-1]
    filename = os.path.basename(file_path)

    file_year, file_month = extract_file_year_month(filename)

    try:
        if ext in ("xls", "xlsx"):
            df = pd.read_excel(
                file_path,
                header=None
            )

        elif ext == "csv":
            df = pd.read_csv(
                file_path,
                header=None,
                encoding="utf-8-sig"
            )

        else:
            raise ValueError(
                "지원하지 않는 파일 형식입니다."
            )

    except Exception as e:
        raise ValueError(
            f"기준측정기 파일을 읽지 못했습니다: {e}"
        )

    (
        date_col_idx,
        time_col_idx,
        ref_col_idx,
        header_row_idx,
    ) = find_header_columns(df)

    if ref_col_idx is None:
        raise ValueError(
            "기준측정기 열을 찾지 못했습니다.\n"
            "Excel 헤더에 'SPREAD', '기준값', '기준농도', "
            "'KOTITI', 'PM2.5' 등의 이름이 있는지 확인하세요."
        )

    if date_col_idx is None or time_col_idx is None:
        raise ValueError(
            "Excel에서 날짜/시간 열을 찾지 못했습니다."
        )

    if header_row_idx is None:
        # 헤더를 찾았지만 같은 행에서 모두 발견되지 않은 경우
        header_row_idx = 0

    ref_data = {}

    for _, row in df.iloc[header_row_idx + 1:].iterrows():

        if (
            pd.isna(row[date_col_idx])
            or pd.isna(row[time_col_idx])
            or pd.isna(row[ref_col_idx])
        ):
            continue

        day = normalize_day(row[date_col_idx])
        hour = normalize_hour(row[time_col_idx])

        if day is None or hour is None:
            continue

        try:
            value = float(
                str(row[ref_col_idx])
                .replace(",", "")
                .strip()
            )
        except (ValueError, TypeError):
            continue

        if not math.isfinite(value):
            continue

        norm_time = (
            f"{file_year}-{file_month:02d}-"
            f"{day:02d} {hour:02d}:00:00"
        )

        ref_data[norm_time] = value

    if not ref_data:
        raise ValueError(
            "기준측정기 데이터를 한 건도 읽지 못했습니다."
        )

    print(
        f"▶ 기준측정기 데이터: {len(ref_data)}시간"
    )

    return ref_data


# ============================================================
# 회귀 / 성능 계산
# ============================================================

def calculate_regression(x_vals, y_vals):
    """
    x = 센서 원본 dust
    y = 기준측정기 spread

    y = scale * x + offset
    """
    if len(x_vals) != len(y_vals):
        raise ValueError(
            "회귀 입력 데이터의 길이가 다릅니다."
        )

    n = len(x_vals)

    if n < 2:
        raise ValueError(
            "회귀 계산에는 최소 2개 이상의 데이터가 필요합니다."
        )

    mean_x = sum(x_vals) / n
    mean_y = sum(y_vals) / n

    ss_xx = sum(
        (x - mean_x) ** 2
        for x in x_vals
    )

    ss_yy = sum(
        (y - mean_y) ** 2
        for y in y_vals
    )

    ss_xy = sum(
        (x - mean_x) * (y - mean_y)
        for x, y in zip(x_vals, y_vals)
    )

    if ss_xx == 0:
        raise ValueError(
            "센서 dust 값의 변동폭이 없어 "
            "회귀 계산이 불가능합니다."
        )

    scale = ss_xy / ss_xx
    offset = mean_y - scale * mean_x

    if ss_yy != 0:
        r2 = (
            ss_xy /
            math.sqrt(ss_xx * ss_yy)
        ) ** 2
    else:
        r2 = 0.0

    return scale, offset, r2


def calculate_metrics(x_vals, y_vals, scale, offset):
    """
    보정 전/후 RMSE, MAE를 계산합니다.
    """
    if not x_vals:
        return {}

    raw_errors = [
        x - y
        for x, y in zip(x_vals, y_vals)
    ]

    corrected_errors = [
        (scale * x + offset) - y
        for x, y in zip(x_vals, y_vals)
    ]

    raw_mae = sum(
        abs(e) for e in raw_errors
    ) / len(raw_errors)

    corrected_mae = sum(
        abs(e) for e in corrected_errors
    ) / len(corrected_errors)

    raw_rmse = math.sqrt(
        sum(e ** 2 for e in raw_errors)
        / len(raw_errors)
    )

    corrected_rmse = math.sqrt(
        sum(e ** 2 for e in corrected_errors)
        / len(corrected_errors)
    )

    return {
        "raw_mae": raw_mae,
        "corrected_mae": corrected_mae,
        "raw_rmse": raw_rmse,
        "corrected_rmse": corrected_rmse,
    }


# ============================================================
# 보정값 저장
# ============================================================

def apply_calibration(
    sensor_index,
    scale_pm25,
    offset_pm25,
    r2=None,
    matched_hours=None,
):
    """
    sensor_calibration.json에 PM2.5 보정 계수를 저장합니다.

    기존 temperature/humidity 등의 설정은 유지하고
    PM2.5(scale[1], offset[1])만 변경합니다.
    """
    calib_data = {}

    if os.path.exists(CALIB_FILE):
        try:
            with open(
                CALIB_FILE,
                "r",
                encoding="utf-8"
            ) as f:
                calib_data = json.load(f)

        except Exception as e:
            print(
                f"⚠️ 기존 calibration 파일을 읽지 못했습니다: {e}"
            )

    idx_str = str(sensor_index)

    if idx_str not in calib_data:
        calib_data[idx_str] = {}

    sensor = calib_data[idx_str]

    if not isinstance(sensor.get("scale"), list):
        sensor["scale"] = [1.0, 1.0, 1.0]

    if not isinstance(sensor.get("offset"), list):
        sensor["offset"] = [0.0, 0.0, 0.0]

    while len(sensor["scale"]) < 3:
        sensor["scale"].append(1.0)

    while len(sensor["offset"]) < 3:
        sensor["offset"].append(0.0)

    # PM2.5 = index 1
    sensor["scale"][1] = round(
        float(scale_pm25),
        6
    )

    sensor["offset"][1] = round(
        float(offset_pm25),
        6
    )

    # 사람이 확인할 수 있도록 메타정보 저장
    sensor["pm25_calibration"] = {
        "formula": "corrected_pm25 = scale * raw_pm25 + offset",
        "scale": round(float(scale_pm25), 6),
        "offset": round(float(offset_pm25), 6),
        "r2": round(float(r2), 6) if r2 is not None else None,
        "matched_hours": matched_hours,
        "updated_at": datetime.now().isoformat(
            timespec="seconds"
        ),
    }

    os.makedirs(
        os.path.dirname(CALIB_FILE),
        exist_ok=True
    )

    temp_file = CALIB_FILE + ".tmp"

    with open(
        temp_file,
        "w",
        encoding="utf-8"
    ) as f:
        json.dump(
            calib_data,
            f,
            ensure_ascii=False,
            indent=2
        )

    os.replace(
        temp_file,
        CALIB_FILE
    )


# ============================================================
# 센서 1개 보정
# ============================================================

def run(sensor_index=0, ref_file_name=None):
    print("\n" + "=" * 60)
    print(
        f"       [센서 {sensor_index + 1}] "
        "PM2.5 자동 교정"
    )
    print("=" * 60)

    if not ref_file_name:
        print("❌ 기준측정기 파일이 지정되지 않았습니다.")
        return

    ref_path = (
        ref_file_name
        if os.path.isabs(ref_file_name)
        else os.path.join(BASE_DIR, ref_file_name)
    )

    # --------------------------------------------------------
    # 1. DB 센서 데이터
    # --------------------------------------------------------
    try:
        sensor_hourly = get_sensor_hourly_averages(
            sensor_index
        )

    except Exception as e:
        print(f"❌ DB 읽기 오류: {e}")
        return

    print(
        f"▶ 센서 {sensor_index + 1}: "
        f"{len(sensor_hourly)}시간 데이터"
    )

    # --------------------------------------------------------
    # 2. 기준측정기 데이터
    # --------------------------------------------------------
    try:
        ref_hourly = load_reference_file(
            ref_path
        )

    except Exception as e:
        print(
            f"❌ 기준측정기 파일 오류: {e}"
        )
        return

    # --------------------------------------------------------
    # 3. 같은 시간끼리 매칭
    # --------------------------------------------------------
    matched_x = []
    matched_y = []
    matched_times = []

    for time_bucket, sensor_info in sensor_hourly.items():

        if time_bucket not in ref_hourly:
            continue

        sensor_value = sensor_info["value"]
        ref_value = ref_hourly[time_bucket]

        if not (
            math.isfinite(sensor_value)
            and math.isfinite(ref_value)
        ):
            continue

        # 음수 PM2.5는 회귀에서 제외
        if sensor_value < 0 or ref_value < 0:
            continue

        matched_x.append(sensor_value)
        matched_y.append(ref_value)
        matched_times.append(time_bucket)

    matched_count = len(matched_x)

    print(
        f"▶ 시간 매칭 결과: "
        f"{matched_count}시간"
    )

    if matched_count < MIN_MATCHED_HOURS:
        print(
            f"❌ 매칭 데이터가 부족합니다. "
            f"(현재 {matched_count}, "
            f"최소 {MIN_MATCHED_HOURS})"
        )

        if matched_times:
            print(
                "▶ 매칭된 시간 예:"
            )
            for t in matched_times[:10]:
                print(f"   {t}")

        return

    # --------------------------------------------------------
    # 4. dust → 기준측정기(spread) 회귀
    #
    #     spread = scale * dust + offset
    # --------------------------------------------------------
    try:
        scale, offset, r2 = calculate_regression(
            matched_x,
            matched_y
        )

    except Exception as e:
        print(f"❌ 회귀 계산 오류: {e}")
        return

    # --------------------------------------------------------
    # 5. 보정 전/후 성능 비교
    # --------------------------------------------------------
    metrics = calculate_metrics(
        matched_x,
        matched_y,
        scale,
        offset
    )

    print("\n📊 PM2.5 보정 결과")
    print("-" * 60)
    print(
        f"▶ 보정식:"
        f"  corrected_pm25 = "
        f"{scale:.6f} × raw_pm25 "
        f"+ {offset:.6f}"
    )
    print(f"▶ Scale : {scale:.6f}")
    print(f"▶ Offset: {offset:.6f}")
    print(f"▶ R²    : {r2:.6f}")

    print("\n📉 오차 비교")
    print(
        f"▶ 보정 전 MAE : "
        f"{metrics['raw_mae']:.4f}"
    )
    print(
        f"▶ 보정 후 MAE : "
        f"{metrics['corrected_mae']:.4f}"
    )
    print(
        f"▶ 보정 전 RMSE: "
        f"{metrics['raw_rmse']:.4f}"
    )
    print(
        f"▶ 보정 후 RMSE: "
        f"{metrics['corrected_rmse']:.4f}"
    )

    # --------------------------------------------------------
    # 6. 실제 예시
    # --------------------------------------------------------
    print("\n🔎 보정 예시")
    print("-" * 60)

    for i in range(min(5, matched_count)):
        raw = matched_x[i]
        reference = matched_y[i]
        corrected = scale * raw + offset

        print(
            f"{matched_times[i]} | "
            f"dust={raw:.2f} → "
            f"보정={corrected:.2f} | "
            f"기준={reference:.2f}"
        )

    # --------------------------------------------------------
    # 7. calibration 파일 저장
    # --------------------------------------------------------
    try:
        apply_calibration(
            sensor_index=sensor_index,
            scale_pm25=scale,
            offset_pm25=offset,
            r2=r2,
            matched_hours=matched_count,
        )

    except Exception as e:
        print(
            f"❌ 보정값 저장 실패: {e}"
        )
        return

    print(
        "\n✔ sensor_calibration.json에 "
        "PM2.5 보정값이 저장되었습니다."
    )

    print(
        "\n⚠️ 중요:"
        "\n  DB의 raw_pm25 원본값 자체는 변경하지 않습니다."
        "\n  실제 표시/전송 단계에서"
        "\n  corrected = scale × raw_pm25 + offset"
        "\n  를 적용해야 합니다."
    )

    print("=" * 60)


# ============================================================
# 프로그램 시작
# ============================================================

if __name__ == "__main__":

    target_file = (
        SPECIFIED_FILE
        if SPECIFIED_FILE
        else find_reference_file(BASE_DIR)
    )

    if not target_file:
        print(
            "❌ 기준측정기 Excel/CSV 파일을 "
            "찾을 수 없습니다."
        )
        sys.exit(1)

    print(
        f"\n▶ 기준측정기 파일: {target_file}"
    )

    # 센서 0~3
    for sensor_index in range(SENSOR_COUNT):
        run(
            sensor_index=sensor_index,
            ref_file_name=target_file
        )
