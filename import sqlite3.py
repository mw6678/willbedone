import sqlite3
import pandas as pd

def get_calibration_average(db_path, target_port, start_time, end_time):
    """
    지정된 포트와 시간대의 3개 미세먼지(PM10, PM25, PM1) 평균값을 반환합니다.
    """
    conn = sqlite3.connect(db_path)

    # SQL 쿼리에 포트와 시간 조건을 추가하여 필요한 데이터만 빠르게 추출합니다.
    query = """
        SELECT pm10, pm25, pm1
        FROM measurements
        WHERE status = 'NORMAL'
          AND port = ?
          AND measured_at >= ?
          AND measured_at < ?
    """

    # 튜플 형태로 파라미터를 안전하게 전달하여 쿼리를 실행합니다.
    df = pd.read_sql_query(query, conn, params=(target_port, start_time, end_time))
    conn.close()

    # 해당 구간에 수집된 데이터가 없는 경우 예외 처리
    if df.empty:
        print(f"[{target_port}] 지정된 시간({start_time} ~ {end_time})의 데이터가 없습니다.")
        return None

    # 3개 항목에 대한 평균 계산 (소수점 2자리 반올림 적용)
    avg_results = df[['pm10', 'pm25', 'pm1']].mean().round(2)

    print(f"--- [교정용 데이터 추출 완료] ---")
    print(f"▶ 포트: {target_port}")
    print(f"▶ 구간: {start_time} ~ {end_time} (총 {len(df)}개)")
    print(f"▶ 평균: PM10={avg_results['pm10']}, PM2.5={avg_results['pm25']}, PM1.0={avg_results['pm1']}\n")

    return avg_results

# --- 교정 프로그램 적용 예시 ---
if __name__ == "__main__":
    db_file = 'dust_measurement.db'

    # 1. 원하는 통신 포트와 1시간 단위 구간 지정
    target_port = 'COM3'
    start_time = '2026-09-03 15:00:00'
    end_time   = '2026-09-03 16:00:00'

    # 2. 함수 호출하여 평균값 획득
    cal_avg = get_calibration_average(db_file, target_port, start_time, end_time)

    # 3. 추출된 데이터를 바탕으로 오차율 계산 등의 후속 로직 진행
    if cal_avg is not None:
        # 기준 장비의 데이터(Reference)와 연계하는 로직
        reference_pm10 = 50.0  # 가상의 기준 장비 농도
        error_rate = ((cal_avg['pm10'] - reference_pm10) / reference_pm10) * 100
        print(f"PM10 오차율: {error_rate:.2f}%")