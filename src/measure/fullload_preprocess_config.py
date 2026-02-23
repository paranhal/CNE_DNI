# -*- coding: utf-8 -*-
"""
전부하 측정 데이터 전처리 - 경로 및 설정

[입력]
- CNE_FULLLOAD_MEASURE.xlsx (통합된 단일 파일)
- 원본 시트: 전부하측정

[출력]
- 동일 파일에 새 시트 추가: 전부하측정_학교별평균
- 열: 학교코드(A), 학교명, 다운로드, 업로드, (D열명), (E열명), 다운로드 진단, 업로드 진단
- D~G열: 측정값 (다운로드, 업로드, 기타 2개)
"""
import os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CNE_DIR = os.path.join(BASE_DIR, "CNE")
DNI_DIR = os.path.join(BASE_DIR, "DNI")

# 전부하 측정 원본 파일 (지역별)
FULLLOAD_MEASURE_FILE = {
    "CNE": os.path.join(CNE_DIR, "CNE_FULLLOAD_MEASURE.xlsx"),
    "DNI": os.path.join(DNI_DIR, "DNI_FULLLOAD_MEASURE.xlsx"),
}

FULLLOAD_MEASURE_CANDIDATES = {
    "CNE": [
        os.path.join(CNE_DIR, "CNE_FULLLOAD_MEASURE.xlsx"),
        os.path.join(BASE_DIR, "CNE_FULLLOAD_MEASURE.xlsx"),
    ],
    "DNI": [
        os.path.join(DNI_DIR, "DNI_FULLLOAD_MEASURE.xlsx"),
        os.path.join(BASE_DIR, "DNI_FULLLOAD_MEASURE.xlsx"),
    ],
}

# 원본 시트명
SHEET_FULLLOAD_SOURCE = "전부하측정"

# 새 시트명 (학교별 평균)
SHEET_FULLLOAD_AVG = "전부하측정_학교별평균"

# 원본 열 (1-based): 장비관리번호는 헤더에서 검색, 없으면 C열(3) 사용
COL_MGMT_NUM_DEFAULT = 3   # C열 (장비관리번호 기본 위치)
COL_DOWNLOAD = 4   # D열
COL_UPLOAD = 5     # E열
COL_MEASURE3 = 6   # F열
COL_MEASURE4 = 7   # G열

# 진단 기준: 375 Mbps 이상 → 양호, 미만 → 미흡 (ISP와 동일)
FULLLOAD_THRESHOLD_MBPS = 375

# 학교 리스트 경로
SCHOOL_LIST_SEARCH_DIRS = [
    BASE_DIR,
    os.path.join(os.path.dirname(BASE_DIR), "split"),
]

# 출력 숫자 형식: 소숫점 1자리 + 단위
NUMBER_FORMAT_DOWNLOAD = '0.0 "Mbps"'
NUMBER_FORMAT_UPLOAD = '0.0 "Mbps"'
NUMBER_FORMAT_MS = '0.0 "ms"'      # RTT, 지연 등
NUMBER_FORMAT_DEFAULT = '0.0'     # 기타
