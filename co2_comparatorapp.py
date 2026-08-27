import sys
import os
import pandas as pd
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, 
    QHBoxLayout, QLabel, QPushButton, QFileDialog, 
    QTextEdit, QMessageBox
)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont

class CO2ComparatorApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.base_file_path = None
        self.target_file_path = None
        self.last_report_text = ""
        self.initUI()

    def initUI(self):
        self.setWindowTitle("CO2 센서 성능인증(1등급) 오차 비교 프로그램")
        self.resize(1000, 800)
        self.setStyleSheet("background-color: #F8F9FA;")

        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(20, 20, 20, 20)
        main_layout.setSpacing(15)

        # 타이틀
        title_label = QLabel("📊 CO2 센서 성능인증(1등급 기준) 오차 비교 분석기")
        title_label.setFont(QFont("Malgun Gothic", 14, QFont.Bold))
        title_label.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(title_label)

        # 파일 선택 레이아웃
        file_layout = QHBoxLayout()
        
        # 기준 파일 영역
        base_box = QVBoxLayout()
        self.base_label = QLabel("기준 파일 (기준기 데이터): 선택되지 않음")
        self.base_label.setStyleSheet("color: #555; font-size: 11pt;")
        btn_base = QPushButton("기준 CSV 선택")
        btn_base.setFixedHeight(35)
        btn_base.setStyleSheet("background-color: #6C757D; color: white; border-radius: 4px; font-weight: bold;")
        btn_base.clicked.connect(self.select_base_file)
        base_box.addWidget(self.base_label)
        base_box.addWidget(btn_base)

        # 비교 파일 영역
        target_box = QVBoxLayout()
        self.target_label = QLabel("비교 파일 (측정기 데이터): 선택되지 않음")
        self.target_label.setStyleSheet("color: #555; font-size: 11pt;")
        btn_target = QPushButton("비교 CSV 선택")
        btn_target.setFixedHeight(35)
        btn_target.setStyleSheet("background-color: #007BFF; color: white; border-radius: 4px; font-weight: bold;")
        btn_target.clicked.connect(self.select_target_file)
        target_box.addWidget(self.target_label)
        target_box.addWidget(btn_target)

        file_layout.addLayout(base_box)
        file_layout.addLayout(target_box)
        main_layout.addLayout(file_layout)

        # 분석 실행 버튼과 저장 버튼
        action_layout = QHBoxLayout()
        
        self.analyze_btn = QPushButton("🔍 1등급 기준 오차 비교 분석 실행")
        self.analyze_btn.setFixedHeight(45)
        self.analyze_btn.setFont(QFont("Malgun Gothic", 11, QFont.Bold))
        self.analyze_btn.setStyleSheet("background-color: #28A745; color: white; border-radius: 5px;")
        self.analyze_btn.clicked.connect(self.run_analysis)
        
        self.save_btn = QPushButton("💾 분석 결과 텍스트로 저장")
        self.save_btn.setFixedHeight(45)
        self.save_btn.setFont(QFont("Malgun Gothic", 11, QFont.Bold))
        self.save_btn.setStyleSheet("background-color: #17A2B8; color: white; border-radius: 5px;")
        self.save_btn.setEnabled(False)
        self.save_btn.clicked.connect(self.save_result_to_txt)

        action_layout.addWidget(self.analyze_btn, stretch=3)
        action_layout.addWidget(self.save_btn, stretch=2)
        main_layout.addLayout(action_layout)

        # 결과 출력 창
        result_label = QLabel("📋 1등급 인증 시뮬레이션 및 전체 상세 비교 결과")
        result_label.setFont(QFont("Malgun Gothic", 11, QFont.Bold))
        main_layout.addWidget(result_label)

        self.result_text = QTextEdit()
        self.result_text.setReadOnly(True)
        self.result_text.setLineWrapMode(QTextEdit.NoWrap)
        self.result_text.setFont(QFont("Consolas", 10))
        self.result_text.setStyleSheet("background-color: white; border: 1px solid #CCC; border-radius: 5px; padding: 10px;")
        main_layout.addWidget(self.result_text)

    def select_base_file(self):
        file_name, _ = QFileDialog.getOpenFileName(self, "기준 CSV 파일 선택", "", "All Files (*);;CSV Files (*.csv *.CSV)")
        if file_name:
            self.base_file_path = file_name
            self.base_label.setText(f"기준 파일: {os.path.basename(file_name)}")

    def select_target_file(self):
        file_name, _ = QFileDialog.getOpenFileName(self, "비교 CSV 파일 선택", "", "All Files (*);;CSV Files (*.csv *.CSV)")
        if file_name:
            self.target_file_path = file_name
            self.target_label.setText(f"비교 파일: {os.path.basename(file_name)}")

    def run_analysis(self):
        if not self.base_file_path or not self.target_file_path:
            QMessageBox.warning(self, "경고", "기준 파일과 비교 파일을 모두 선택해주세요.")
            return

        try:
            df_base = pd.read_csv(self.base_file_path, encoding='utf-8-sig')
            df_target = pd.read_csv(self.target_file_path, encoding='utf-8-sig')

            required_cols = ['측정시간', 'Co2 1분 평균(ppm)']
            for col in required_cols:
                if col not in df_base.columns or col not in df_target.columns:
                    raise ValueError(f"CSV 파일에 필수 컬럼('{col}')이 누락되었습니다.")

            df_base['Co2 1분 평균(ppm)'] = pd.to_numeric(df_base['Co2 1분 평균(ppm)'], errors='coerce')
            df_target['Co2 1분 평균(ppm)'] = pd.to_numeric(df_target['Co2 1분 평균(ppm)'], errors='coerce')

            merged = pd.merge(df_base, df_target, on='측정시간', suffixes=('_기준', '_비교'))

            if merged.empty:
                QMessageBox.warning(self, "결과 없음", "두 파일 간에 일치하는 '측정시간' 데이터가 존재하지 않습니다.")
                return

            merged = merged.dropna(subset=['Co2 1분 평균(ppm)_기준', 'Co2 1분 평균(ppm)_비교'])

            if merged.empty:
                QMessageBox.warning(self, "결과 없음", "비교할 수 있는 유효한 숫자 데이터가 겹치지 않습니다.")
                return

            base_vals = merged['Co2 1분 평균(ppm)_기준']
            target_vals = merged['Co2 1분 평균(ppm)_비교']

            diff = target_vals - base_vals               
            abs_diff = diff.abs()                        
            error_rate = (abs_diff / base_vals) * 100    

            merged['오차(Diff)'] = diff
            merged['절대오차'] = abs_diff
            merged['오차율(%)'] = error_rate.round(2)

            # 통계치 계산
            mean_base = base_vals.mean()
            mean_target = target_vals.mean()
            mean_abs_diff = abs_diff.mean()
            max_abs_diff = abs_diff.max()
            rmse = ((diff ** 2).mean()) ** 0.5

            # --- [KTL 등급 기준 시뮬레이션 판정 로직] ---
            # 예시 기준으로 1등급 판정 허용 오차율을 ±10% 이내로 설정 (필요시 조절 가능)
            ALLOWED_ERROR_RATE = 10.0 
            pass_count = (error_rate <= ALLOWED_ERROR_RATE).sum()
            total_count = len(merged)
            pass_ratio = (pass_count / total_count) * 100

            is_grade_1_candidate = pass_ratio >= 95.0  # 전체 데이터의 95% 이상이 오차 범위 내에 들면 적합 판정 예시

            report = []
            report.append("=" * 80)
            report.append(" 🏆 KTL CO2 간이측정기 성능인증(1등급) 오차 비교 분석 리포트")
            report.append("=" * 80)
            report.append(f"• 기준 파일: {os.path.basename(self.base_file_path)}")
            report.append(f"• 비교 파일: {os.path.basename(self.target_file_path)}")
            report.append(f"• 총 매칭된 데이터 개수: {total_count} 개 분(minute)")
            report.append("-" * 80)
            report.append(f"• 기준 파일 평균 PPM : {mean_base:.2f} ppm")
            report.append(f"• 비교 파일 평균 PPM : {mean_target:.2f} ppm")
            report.append(f"• 평균 절대 오차 (MAE): {mean_abs_diff:.2f} ppm")
            report.append(f"• 최대 절대 오차     : {max_abs_diff:.2f} ppm")
            report.append(f"• 평균 제곱근 오차(RMSE): {rmse:.2f} ppm")
            report.append("-" * 80)
            report.append(" [ 1등급 기준 시뮬레이션 판정 결과 ]")
            report.append(f"  - 기준 허용 오차율   : ±{ALLOWED_ERROR_RATE}% 이내")
            report.append(f"  - 기준 충족 데이터   : {pass_count} / {total_count} ({pass_ratio:.1f}%)")
            
            if is_grade_1_candidate:
                report.append("  📌 판정 결과       : [1등급 기준 충족 (Pass 검토 가능)] 🎉")
            else:
                report.append("  📌 판정 결과       : [기준 미달 (Fail / 보정 필요)] ⚠️")
            
            report.append("=" * 80)
            report.append("\n[ 전체 시간대별 상세 비교 및 1등급 기준 만족 여부 ]")
            report.append(f"{'측정시간':<15} | {'기준(ppm)':<10} | {'비교(ppm)':<10} | {'오차(Diff)':<10} | {'오차율(%)':<10} | {'판정'}")
            report.append("-" * 75)

            for idx, row in merged.iterrows():
                t = row['측정시간']
                b_val = row['Co2 1분 평균(ppm)_기준']
                t_val = row['Co2 1분 평균(ppm)_비교']
                d_val = row['오차(Diff)']
                r_val = row['오차율(%)']
                
                status = "PASS" if r_val <= ALLOWED_ERROR_RATE else "FAIL"
                report.append(f"{str(t):<15} | {b_val:<10.1f} | {t_val:<10.1f} | {d_val:<+10.1f} | {r_val}%        | {status}")

            report.append("=" * 80)
            report.append(" [분석 완료] 모든 데이터가 출력되었습니다.")

            self.last_report_text = "\n".join(report)
            self.result_text.setText(self.last_report_text)
            self.save_btn.setEnabled(True)

        except Exception as e:
            QMessageBox.critical(self, "오류 발생", f"파일을 분석하는 동안 오류가 발생했습니다.\n상세 내용: {str(e)}")

    def save_result_to_txt(self):
        if not self.last_report_text:
            QMessageBox.warning(self, "경고", "저장할 분석 결과가 없습니다. 먼저 분석을 실행해주세요.")
            return

        file_name, _ = QFileDialog.getSaveFileName(self, "분석 결과 저장", "CO2_Grade1_Report.txt", "Text Files (*.txt);;All Files (*)")
        if file_name:
            try:
                with open(file_name, 'w', encoding='utf-8-sig') as f:
                    f.write(self.last_report_text)
                QMessageBox.information(self, "성공", f"분석 결과가 성공적으로 저장되었습니다.\n경로: {file_name}")
            except Exception as e:
                QMessageBox.critical(self, "저장 오류", f"파일을 저장하는 중 오류가 발생했습니다.\n상세 내용: {str(e)}")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = CO2ComparatorApp()
    window.show()
    sys.exit(app.exec_())