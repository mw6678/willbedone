import sys
import os
import pandas as pd
import numpy as np
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, 
    QHBoxLayout, QLabel, QPushButton, QFileDialog, 
    QTextEdit, QMessageBox
)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont

class CO2PrecisionApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.file_path = None
        self.last_report_text = ""
        self.initUI()

    def initUI(self):
        self.setWindowTitle("CO2 센서 정밀도 및 정확성(변동계수/RSD) 검증기")
        self.resize(900, 750)
        self.setStyleSheet("background-color: #F8F9FA;")

        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(20, 20, 20, 20)
        main_layout.setSpacing(15)

        # 타이틀
        title_label = QLabel("📊 CO2 센서 데이터 정밀도/변동성(RSD) 분석기")
        title_label.setFont(QFont("Malgun Gothic", 14, QFont.Bold))
        title_label.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(title_label)

        # 파일 선택 레이아웃
        file_box = QVBoxLayout()
        self.file_label = QLabel("선택된 파일: 없음")
        self.file_label.setStyleSheet("color: #555; font-size: 11pt;")
        btn_file = QPushButton("로그 파일 선택 (CSV / 엑셀)")
        btn_file.setFixedHeight(40)
        btn_file.setStyleSheet("background-color: #007BFF; color: white; border-radius: 4px; font-weight: bold; font-size: 11pt;")
        btn_file.clicked.connect(self.select_file)
        file_box.addWidget(self.file_label)
        file_box.addWidget(btn_file)
        main_layout.addLayout(file_box)

        # 분석 실행 버튼과 저장 버튼
        action_layout = QHBoxLayout()
        
        self.analyze_btn = QPushButton("🔍 정밀도(표준편차/평균*100) 분석 실행")
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
        result_label = QLabel("📋 분석 리포트 및 시간대별 상세 측정 결과")
        result_label.setFont(QFont("Malgun Gothic", 11, QFont.Bold))
        main_layout.addWidget(result_label)

        self.result_text = QTextEdit()
        self.result_text.setReadOnly(True)
        self.result_text.setLineWrapMode(QTextEdit.NoWrap)
        self.result_text.setFont(QFont("Consolas", 10))
        self.result_text.setStyleSheet("background-color: white; border: 1px solid #CCC; border-radius: 5px; padding: 10px;")
        main_layout.addWidget(self.result_text)

    def select_file(self):
        file_name, _ = QFileDialog.getOpenFileName(self, "CO2 로그 파일 선택", "", "All Files (*);;CSV Files (*.csv);;Excel Files (*.xlsx *.xls)")
        if file_name:
            self.file_path = file_name
            self.file_label.setText(f"선택된 파일: {os.path.basename(file_name)}")

    def run_analysis(self):
        if not self.file_path:
            QMessageBox.warning(self, "경고", "먼저 분석할 로그 파일을 선택해주세요.")
            return

        try:
            # 파일 형식(CSV 또는 엑셀)에 따라 로드
            if self.file_path.endswith('.csv'):
                df = pd.read_csv(self.file_path, encoding='utf-8-sig')
            else:
                df = pd.read_excel(self.file_path, sheet_name=0)

            target_col = 'Co2 1분 평균(ppm)'
            if target_col not in df.columns:
                raise ValueError(f"파일에 필수 컬럼('{target_col}')이 누락되었습니다.")

            # 숫자가 아닌 에러 문자열 및 결측치 제거
            df['Clean_Co2'] = pd.to_numeric(df[target_col], errors='coerce')
            df_valid = df.dropna(subset=['Clean_Co2'])

            if df_valid.empty:
                QMessageBox.warning(self, "결과 없음", "유효한 숫자 형태의 CO2 데이터가 존재하지 않습니다.")
                return

            sensor_vals = df_valid['Clean_Co2']
            
            # 통계치 계산
            total_count = len(sensor_vals)
            mean_val = sensor_vals.mean()
            std_val = sensor_vals.std(ddof=1) if total_count > 1 else 0.0
            
            # 변동계수(RSD, %) = (표준편차 / 평균) * 100
            rsd_val = (std_val / mean_val) * 100 if mean_val != 0 else 0.0

            min_val = sensor_vals.min()
            max_val = sensor_vals.max()

            report = []
            report.append("=" * 80)
            report.append(" 📊 CO2 센서 정밀도 및 변동성(RSD) 분석 리포트")
            report.append("=" * 80)
            report.append(f"• 분석 파일명       : {os.path.basename(self.file_path)}")
            report.append(f"• 유효 데이터 개수  : {total_count} 개 (분 단위)")
            report.append("-" * 80)
            report.append(f"• 전체 데이터 평균  : {mean_val:.2f} ppm")
            report.append(f"• 전체 표준편차(σ)  : {std_val:.2f} ppm")
            report.append(f"• 최솟값 / 최댓값   : {min_val:.1f} ppm / {max_val:.1f} ppm")
            report.append("-" * 80)
            report.append(f"• [정밀도/변동계수 RSD] : (표준편차 / 평균) * 100 = {rsd_val:.2f} %")
            report.append("=" * 80)
            report.append("\n[ 시간대별 상세 측정값 목록 ]")
            report.append(f"{'측정시간':<20} | {'측정값(ppm)':<15}")
            report.append("-" * 40)

            for idx, row in df_valid.iterrows():
                t = row['측정시간'] if '측정시간' in df_valid.columns else f"Row {idx}"
                val = row['Clean_Co2']
                report.append(f"{str(t):<20} | {val:<15.1f}")

            report.append("=" * 80)
            report.append(" [분석 완료] 모든 데이터가 정상적으로 처리되었습니다.")

            self.last_report_text = "\n".join(report)
            self.result_text.setText(self.last_report_text)
            self.save_btn.setEnabled(True)

        except Exception as e:
            QMessageBox.critical(self, "오류 발생", f"파일을 분석하는 중 오류가 발생했습니다.\n상세 내용: {str(e)}")

    def save_result_to_txt(self):
        if not self.last_report_text:
            QMessageBox.warning(self, "경고", "저장할 분석 결과가 없습니다. 먼저 분석을 실행해주세요.")
            return

        file_name, _ = QFileDialog.getSaveFileName(self, "분석 결과 저장", "CO2_Precision_Report.txt", "Text Files (*.txt);;All Files (*)")
        if file_name:
            try:
                with open(file_name, 'w', encoding='utf-8-sig') as f:
                    f.write(self.last_report_text)
                QMessageBox.information(self, "성공", f"분석 결과가 성공적으로 저장되었습니다.\n경로: {file_name}")
            except Exception as e:
                QMessageBox.critical(self, "저장 오류", f"파일을 저장하는 중 오류가 발생했습니다.\n상세 내용: {str(e)}")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = CO2PrecisionApp()
    window.show()
    sys.exit(app.exec_())