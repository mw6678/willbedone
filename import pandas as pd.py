import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.dates as mdates

# 폰트 설정 (그래프 한글 깨짐 방지)
plt.rcParams['font.family'] = 'Malgun Gothic'
plt.rcParams['axes.unicode_minus'] = False

# ==========================================
# [사용자 정의 함수] 구간별(저농도/고농도) 보정 함수
# ==========================================
def apply_piecewise_calibration(raw_val, threshold, low_scale, low_offset, high_scale, high_offset):
    # 센서 원본값이 임계값(threshold)보다 작으면 저농도 수식 적용
    if raw_val < threshold:
        return (raw_val * low_scale) + low_offset
    # 센서 원본값이 임계값 이상이면 고농도 수식 적용
    else:
        return (raw_val * high_scale) + high_offset

# ==========================================
# 1. 파일 경로 설정
# ==========================================
ref_file = "기준측정기 데이터_261002-261006.xls"
sensor_file = "시간당_센서평균_리포트_20261006_1247.xlsx"

# ==========================================
# 2. 기준측정기 데이터 전처리
# ==========================================
ref_df = pd.read_excel(ref_file, header=1) 
ref_df.rename(columns={'일': 'day', '시간': 'hour', 'KOTITI\n기준측정기': 'Reference'}, inplace=True)
ref_df.dropna(subset=['day', 'hour'], inplace=True)

ref_df['day'] = ref_df['day'].astype(str).str.replace('일', '').astype(int)
ref_df['hour'] = ref_df['hour'].astype(int)
ref_df['datetime'] = pd.to_datetime({'year': 2026, 'month': 10, 'day': ref_df['day'], 'hour': ref_df['hour']})
ref_df = ref_df[['datetime', 'Reference']]

# ==========================================
# 3. 센서 데이터 전처리 및 구간별 보정(Calibration)
# ==========================================
sensor_df = pd.read_excel(sensor_file)
sensor_df['시간'] = pd.to_datetime(sensor_df['시간'])
sensor_columns = [col for col in sensor_df.columns if '센서' in col]

# [데이터 기반 산출] 구간별 최적 보정 계수
calibration_params = {
    '센서 1': {'threshold': 3.5, 'low_scale': 0.76, 'low_offset': 6.18, 'high_scale': 0.15, 'high_offset': 9.21},
    '센서 2': {'threshold': 3.5, 'low_scale': 0.79, 'low_offset': 6.12, 'high_scale': 0.15, 'high_offset': 9.20},
    '센서 3': {'threshold': 4.0, 'low_scale': 0.61, 'low_offset': 6.10, 'high_scale': 0.15, 'high_offset': 9.17},
    '센서 4': {'threshold': 4.5, 'low_scale': 0.68, 'low_offset': 6.03, 'high_scale': 0.15, 'high_offset': 9.18}
}

# 설정한 구간별 수식을 각 센서 데이터에 일괄 적용
for sensor in sensor_columns:
    if sensor in calibration_params:
        p = calibration_params[sensor]
        sensor_df[sensor] = sensor_df[sensor].apply(
            lambda x: apply_piecewise_calibration(
                x, p['threshold'], 
                p['low_scale'], p['low_offset'], 
                p['high_scale'], p['high_offset']
            )
        )

# ==========================================
# 4. 데이터 병합 & 5미만 필터링
# ==========================================
merged_df = pd.merge(ref_df, sensor_df, left_on='datetime', right_on='시간', how='inner')
merged_df.drop('시간', axis=1, inplace=True)
merged_df.set_index('datetime', inplace=True)

# 기준측정기 값이 5 미만인 데이터는 평가에서 제외 (현재 데이터는 모두 5 이상이긴 함)
merged_df = merged_df[merged_df['Reference'] >= 5.0]

# ==========================================
# 5. 상대오차(%) 및 상대정확도(%) 계산
# ==========================================
for sensor in sensor_columns:
    merged_df[f'{sensor}_상대오차(%)'] = ((merged_df[sensor] - merged_df['Reference']) / merged_df['Reference']) * 100
    merged_df[f'{sensor}_상대정확도(%)'] = 100 - abs(merged_df[f'{sensor}_상대오차(%)'])

print("✅ 데이터 보정(저/고농도 분리), 상대오차 계산 완료!")

# 엑셀 파일 저장
merged_df.to_excel("구간별보정_최종_기능검사_결과.xlsx")

# ==========================================
# 6. 결과 시각화
# ==========================================
fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(14, 10), sharex=True)

# ----- [위쪽: 측정값 비교] -----
ax1.plot(merged_df.index, merged_df['Reference'], label='기준측정기(Reference)', color='black', linewidth=2, marker='o')
for sensor in sensor_columns:
    ax1.plot(merged_df.index, merged_df[sensor], label=f'{sensor}(구간보정됨)', linestyle='--', marker='x')

ax1.set_title('시간별 기준측정기 및 센서 측정값 비교 (저/고농도 분리 보정)', fontsize=14)
ax1.set_ylabel('측정값')
ax1.legend()
ax1.grid(True, linestyle=':', alpha=0.6)

# ----- [아래쪽: 시간당 상대오차(%) 비교] -----
ax2.axhline(0, color='black', linewidth=1.5, linestyle='-')
ax2.axhspan(-30, 30, color='green', alpha=0.1, label='±30% 오차 범위') # 30% 기준으로 음영 표시

for sensor in sensor_columns:
    ax2.plot(merged_df.index, merged_df[f'{sensor}_상대오차(%)'], label=f'{sensor} 상대오차', linestyle='-', marker='s')

ax2.set_title('시간당 센서 상대오차 (%)', fontsize=14)
ax2.set_ylabel('상대오차 (%)')
ax2.set_xlabel('시간')
ax2.legend()
ax2.grid(True, linestyle=':', alpha=0.6)

ax2.xaxis.set_major_formatter(mdates.DateFormatter('%m-%d %H:%M'))
ax2.xaxis.set_major_locator(mdates.HourLocator(interval=6))
plt.xticks(rotation=45)

plt.tight_layout()
plt.show()