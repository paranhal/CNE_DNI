# -*- coding: utf-8 -*-
"""
학교별 측정 리포트 생성 - 경로 및 설정

[입력]
- 템플릿: 최종_측정값_템플릿.xlsx (우선 사용)
- 통계: TOTAL_MEASURE_LIST_V1.XLSX (통합 통계)

[출력]
- 학교별 xlsx 파일 (템플릿 복사 + 측정값/판정 입력)
"""
import os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CNE_DIR = os.path.join(BASE_DIR, "CNE")
OUTPUT_DIR = os.path.join(CNE_DIR, "학교별_리포트")

# 템플릿: 최종_측정값_템플릿.xlsx 우선 (measure 폴더)
TEMPLATE_DIR = os.path.join(BASE_DIR, "측정값_템플릿")
TEMPLATE_DIR_ALT = os.path.join(BASE_DIR, "측정값_템플릿")
TEMPLATE_CANDIDATES = [
    os.path.join(BASE_DIR, "최종_측정값_템플릿.xlsx"),
    os.path.join(TEMPLATE_DIR, "최종_측정값_템플릿.xlsx"),
    os.path.join(BASE_DIR, "측정값_템플릿.xlsx"),
    os.path.join(TEMPLATE_DIR, "측정값_템플릿.xlsx"),
    os.path.join(TEMPLATE_DIR, "템플릿.xlsx"),
    os.path.join(TEMPLATE_DIR_ALT, "최종_측정값_템플릿.xlsx"),
    os.path.join(TEMPLATE_DIR_ALT, "측정값_템플릿.xlsx"),
    os.path.join(TEMPLATE_DIR_ALT, "템플릿.xlsx"),
]

# 통합 통계 파일
TOTAL_MEASURE_LIST = os.path.join(CNE_DIR, "TOTAL_MEASURE_LIST_V1.XLSX")

# 최종_측정값_템플릿: 측정값=F열(6), 평가결과=H열(8)
# (행, 시트명, 소스열1, 소스열2_옵션) - 셀번호 F11→행11, F22→행22 ...
J_OUTPUT_MAP = [
    (11, "ISP측정_학교별평균", 6, None),    # F11: ISP RSSI
    (22, "전부하측정_학교별평균", 3, 4),    # F22: 전부하 C/D
    (23, "ISP측정_학교별평균", 3, None),    # F23: ISP C열
    (24, "ISP측정_학교별평균", 4, None),    # F24: ISP D열
    (25, "ISP측정_학교별평균", 6, None),    # F25: ISP F열(RSSI)
    (28, "집선ISP", 5, None),               # F28: 집선ISP E열
    (29, "집선ISP", 3, None),               # F29: 집선ISP C열
    (30, "집선ISP", 4, None),               # F30: 집선ISP D열
    (31, "CNE_WIRED_MEANSURE_AVG", 5, None),   # F31: CNE_WIRED E열
    (32, "CNE_WIRED_MEANSURE_AVG", 6, None),   # F32: CNE_WIRED F열
    (33, "CNE_WIRED_MEANSURE_AVG", 4, None),   # F33: CNE_WIRED D열
    (34, "CNE_WIRED_MEANSURE_AVG", 4, None),   # F34: CNE_WIRED D열
]

# H열 판정: (행, 연산, 기준값) - F열 데이터 기준
L_JUDGMENT_MAP = [
    (11, "ge", -60),   # H11: -60 이상 (RSSI)
    (22, "ge", 375),
    (23, "ge", 375),
    (24, "ge", 375),
    (25, "ge", -60),   # H25: -60 이상 (RSSI)
    (28, "le", 10),
    (29, "ge", 700),
    (30, "ge", 700),
    (31, "le", 1.0),
    (32, "le", 5),
    (33, "ge", 700),
    (34, "ge", 700),
]

J_COL = 6   # F열 측정값
L_COL = 8   # H열 평가결과

# 로그
LOG_DIR = os.path.join(BASE_DIR, "logs")
LOG_PREFIX = "school_report"
