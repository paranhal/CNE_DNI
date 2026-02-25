# -*- coding: utf-8 -*-
"""
대전(DNI) 학교별 측정 리포트 생성 - 경로 및 설정

[입력]
- 템플릿: 최종_측정값_템플릿.xlsx
- 통계: DNI_TOTAL_MEASURE_LIST_V1.xlsx

[출력]
- DNI/학교별_리포트/ 하위에 학교별 xlsx 파일
"""
import os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DNI_DIR = os.path.join(BASE_DIR, "DNI")
OUTPUT_DIR = os.path.join(DNI_DIR, "학교별_리포트")

TEMPLATE_DIR = os.path.join(BASE_DIR, "측정값_템플릿")
TEMPLATE_CANDIDATES = [
    os.path.join(BASE_DIR, "최종_측정값_템플릿.xlsx"),
    os.path.join(TEMPLATE_DIR, "최종_측정값_템플릿.xlsx"),
    os.path.join(BASE_DIR, "측정값_템플릿.xlsx"),
    os.path.join(TEMPLATE_DIR, "측정값_템플릿.xlsx"),
    os.path.join(TEMPLATE_DIR, "템플릿.xlsx"),
]

TOTAL_MEASURE_LIST = os.path.join(DNI_DIR, "DNI_TOTAL_MEASURE_LIST_V1.xlsx")

# F열 데이터 매핑: (행, 시트, 열1, 열2_또는_리스트, format_type)
J_OUTPUT_MAP = [
    (2, "학교별통신장비현황", 16, None, None),       # F2: P열
    (3, "h_only", None, None, "h_only"),
    (4, "POE", 3, None, None),                       # F4: C열
    (6, "케이블통계", [4, 5, 6, 7, 8], None, "cable"),
    (10, "AP_장비통계", 3, None, None),              # F10: C열 (없으면 빈칸)
    (11, "ISP측정_학교별평균", 6, None, None),       # F11: F열
    (12, "fixed", "80 Mhz", None, "fixed"),
    (14, "POE", 4, None, None),
    (15, "대전AP", 7, 8, "location5"),               # 대전AP
    (16, "h_only", None, None, "h_only"),
    (17, "대전AP", [4, 5, 6], None, "n_ac_ax"),      # 대전AP
    (18, "학교별통신장비현황", 47, None, None),
    (19, "케이블통계", 8, None, None),
    (20, "학교별통신장비현황", 9, None, None),
    (21, "전부하측정_학교별평균", 3, 4, "fullload"),
    (22, "전부하측정_학교별평균", 3, 4, None),
    (23, "ISP측정_학교별평균", 3, None, None),
    (24, "ISP측정_학교별평균", 4, None, None),
    (25, "ISP측정_학교별평균", 6, None, None),
    (26, "POE", 4, None, None),
    (28, "집선ISP", 5, None, None),
    (29, "집선ISP", 3, None, None),
    (30, "집선ISP", 4, None, None),
    (31, "DNI_WIRED_MEANSURE_AVG", 5, None, None),   # DNI_WIRED
    (32, "DNI_WIRED_MEANSURE_AVG", 6, None, None),
    (33, "DNI_WIRED_MEANSURE_AVG", 4, None, None),
    (34, "DNI_WIRED_MEANSURE_AVG", 4, None, None),
]

L_JUDGMENT_MAP = [
    (2, "le", 0),
    (3, "always", "정상"),
    (6, "always", "정상"),
    (4, "le", 12),
    (10, "zero_or_empty_ok", 0),
    (11, "ge", -60),
    (12, "always", "정상"),
    (14, "le", 3),
    (16, "always", "정상"),
    (15, "has_value", None),
    (17, "has_value", None),
    (18, "ge_before_keyword", 3),
    (19, "le", 0),
    (20, "split_exact", 0),
    (21, "both_ge", 375),
    (22, "ge", 375),
    (23, "ge", 375),
    (24, "ge", 375),
    (25, "ge", -60),
    (26, "le", 3),
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
G_COL = 7   # G열

JUDGMENT_ROW_START = 2
JUDGMENT_ROW_END = 34

FONT_BLACK_ROWS = (28, 29, 30)
JUDGE_BY_V2_ROWS = (15,)
JUDGE_BOTH_ROWS = (21,)

ROUND_1_ROWS = {14, 26, 27}

LOG_DIR = os.path.join(BASE_DIR, "logs")
LOG_PREFIX = "dni_school_report"

# 학교 리스트 파일명 (DNI용)
SCHOOL_LIST_FILES = [
    "school_reg_list_DNI.xlsx",
    "school_reg_list_DNI.csv",
    "SCHOOL_REG_LIST_DNI.xlsx",
]
