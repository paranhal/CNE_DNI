# -*- coding: utf-8 -*-
"""
무선망(ISP) 측정 데이터 전처리 - 경로 및 설정

[입력]
- CNE_ISP_MEASURE.XLSX (통합된 단일 파일)
- 원본 시트: ISP측정

[출력]
- 동일 파일에 새 시트 추가: ISP측정_학교별평균
- 열: 학교명, 학교코드, 다운로드, 업로드, RTT, RSSI, CH, 다운로드 진단, 업로드 진단
"""
import os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CNE_DIR = os.path.join(BASE_DIR, "CNE")
DNI_DIR = os.path.join(BASE_DIR, "DNI")

# ISP 측정 원본 파일 (지역별) - CNE/DNI 폴더 또는 measure 폴더 직접
ISP_MEASURE_FILE = {
    "CNE": os.path.join(CNE_DIR, "CNE_ISP_MEASURE.XLSX"),
    "DNI": os.path.join(DNI_DIR, "DNI_ISP_MEASURE.XLSX"),
}

# 파일 없을 때 탐색할 후보 경로 (measure 폴더 직접)
ISP_MEASURE_CANDIDATES = {
    "CNE": [
        os.path.join(CNE_DIR, "CNE_ISP_MEASURE.XLSX"),
        os.path.join(BASE_DIR, "CNE_ISP_MEASURE.XLSX"),
    ],
    "DNI": [
        os.path.join(DNI_DIR, "DNI_ISP_MEASURE.XLSX"),
        os.path.join(BASE_DIR, "DNI_ISP_MEASURE.XLSX"),
    ],
}

# 원본 시트명
SHEET_ISP_SOURCE = "ISP측정"

# 새 시트명 (학교별 평균)
SHEET_ISP_AVG = "ISP측정_학교별평균"

# 원본 열 (1-based): C=장비관리번호, E~I=다운로드,업로드,RTT,RSSI,CH
COL_MGMT_NUM = 3   # C열: 장비관리번호
COL_DOWNLOAD = 5   # E열: 다운로드
COL_UPLOAD = 6     # F열: 업로드
COL_RTT = 7        # G열: RTT
COL_RSSI = 8       # H열: RSSI
COL_CH = 9         # I열: CH

# 진단 기준: 375 Mbps 이상 → 양호, 미만 → 미흡
ISP_THRESHOLD_MBPS = 375

# 학교 리스트 경로 (measure → split 순으로 탐색)
SCHOOL_LIST_SEARCH_DIRS = [
    BASE_DIR,
    os.path.join(os.path.dirname(BASE_DIR), "split"),
]
