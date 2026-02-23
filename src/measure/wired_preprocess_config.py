# -*- coding: utf-8 -*-
"""
유선망 측정 데이터 전처리 - 경로 및 설정

[폴더 구조]
- CNE: 충남 유선망 1차 측정 결과
- DNI: 대전 유선망 1차 측정 결과

[출력]
- CNE: CNE_WIRED_MEANSURE_V1.XLSX
  - CNE_TOTAL: 통합 원본 데이터
  - CNE_WIRED_MEANSURE_AVG: 학교별 평균값
"""
import os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CNE_DIR = os.path.join(BASE_DIR, "CNE")
DNI_DIR = os.path.join(BASE_DIR, "DNI")

# 유선망 1차 측정 결과 소스 폴더
WIRED_1ST_SOURCE = {
    "CNE": os.path.join(CNE_DIR, "유선망품질측정결과1차_충남"),
    "DNI": os.path.join(DNI_DIR, "유선망품질측정결과1차_대전"),
}

# 출력 파일 (지역별)
WIRED_OUTPUT_FILE = {
    "CNE": os.path.join(CNE_DIR, "CNE_WIRED_MEANSURE_V1.XLSX"),
    "DNI": os.path.join(DNI_DIR, "DNI_WIRED_MEANSURE_V1.XLSX"),
}

# 시트 이름
SHEET_TOTAL = "CNE_TOTAL"      # CNE용 (DNI는 DNI_TOTAL)
SHEET_AVG = "CNE_WIRED_MEANSURE_AVG"  # CNE용 (DNI는 DNI_WIRED_MEANSURE_AVG)

SHEET_NAMES = {
    "CNE": {"total": "CNE_TOTAL", "avg": "CNE_WIRED_MEANSURE_AVG"},
    "DNI": {"total": "DNI_TOTAL", "avg": "DNI_WIRED_MEANSURE_AVG"},
}

# 학교코드 추출: 파일명에서 12자리 또는 G/E/N + 숫자 패턴
import re
SCHOOL_CODE_PATTERN = re.compile(r"([GNE]\d{9}[A-Z]{2}|\d{12})")

# 진단결과 기준: Avg Throughput (Mbps) >= THROUGHPUT_THRESHOLD → 양호, 이하면 미흡
THROUGHPUT_THRESHOLD_MBPS = 700

# 학교 리스트 경로 (measure → split 순으로 탐색)
SCHOOL_LIST_SEARCH_DIRS = [
    BASE_DIR,
    os.path.join(os.path.dirname(BASE_DIR), "split"),
]

# AVG 시트 열 선택 (현재 출력 기준 1-based: A=1,B=2,...,O=15)
# 결과 순서: 학교코드(D열), 학교명(E열), 장비개수, K열, L열, M열, N열, 진단결과(K열 700Mbps 기준)
AVG_COLUMN_SELECT = {
    "keep_cols": [11, 12, 13, 14],  # K, L, M, N (측정값만)
}

# 열별 표시 형식 (헤더 키워드 → 숫자 서식, 실제 값은 숫자 유지)
# Loss Ratio: 소숫점 2자리 %, 나머지: 소숫점 1자리
COLUMN_DISPLAY_FORMATS = [
    (["Throughput", "Mbps", "대역폭", "처리량"], '0.0 "Mbps"'),
    (["Loss", "Loss Rate", "Loss Ratio", "손실률", "패킷손실"], '0.00"%"'),
    (["Latency", "ms", "지연", "지연시간"], '0.0 "ms"'),
    (["Jitter", "지터"], '0.0 "ms"'),
]
