# -*- coding: utf-8 -*-
"""
전부하 원시 데이터(FULLLOAD_RAWA_1) → 통계용 원본 → 학교별 평균

[입력]
- FULLLOAD_RAWA_1.xlsx: 1차 측정, 2차 측정 시트

[선택 로직]
- 1차 양호 → 1차 자료 사용
- 1차 미흡 + 2차 양호 → 2차 자료 사용
- 1차 미흡 + 2차 미흡 → 1차 자료 사용
- 1차만 있음 → 1차 자료 사용
- 2차만 있음 → 2차 자료 사용

[출력]
- CNE_FULLLOAD_MEASURE.xlsx (통계용 원본: 전부하측정 시트)
- 전부하측정_학교별평균 시트 (학교별만, 전체 행 없음)
"""
import os
from datetime import datetime

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CNE_DIR = os.path.join(BASE_DIR, "CNE")

# 원시 데이터 파일 (CNE 폴더 또는 measure 폴더)
FULLLOAD_RAW_CANDIDATES = [
    os.path.join(CNE_DIR, "FULLLOAD_RAWA_1.xlsx"),
    os.path.join(BASE_DIR, "FULLLOAD_RAWA_1.xlsx"),
]

# 1차/2차 시트명 (실제 시트명에 맞게 수정)
SHEET_1ST = "1차측정"
SHEET_2ND = "2차측정"

# 시트명 후보 (자동 검색용)
SHEET_1ST_CANDIDATES = ["1차측정", "1차", "1차 측정", "1차측정결과"]
SHEET_2ND_CANDIDATES = ["2차측정", "2차", "2차 측정", "2차측정결과"]

# 출력 파일
FULLLOAD_OUTPUT = os.path.join(CNE_DIR, "CNE_FULLLOAD_MEASURE.xlsx")
FULLLOAD_STATS_OUTPUT = os.path.join(CNE_DIR, "CNE_FULLLOAD_MEASURE_통계.xlsx")
SHEET_SOURCE = "전부하측정"
SHEET_AVG = "전부하측정_학교별평균"

# 진단 기준 (Mbps)
THRESHOLD_MBPS = 375

# 로그 파일
LOG_DIR = os.path.join(BASE_DIR, "logs")
LOG_PREFIX = "fullload_raw"
