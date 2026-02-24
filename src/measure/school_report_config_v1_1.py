# -*- coding: utf-8 -*-
"""
학교별 측정 리포트 생성 - 경로 및 설정 (V1.1)

[입력]
- 템플릿: 최종_측정값_템플릿.xlsx (우선 사용)
- 통계: TOTAL_MEASURE_LIST_V1.XLSX (통합 통계)

[출력]
- 학교별 xlsx 파일 (템플릿 복사 + 측정값/판정 입력)

[데이터 매핑 형식]
- 단순: (행, 시트, 열1, 열2_옵션, None) -> "v1" 또는 "v1 / v2"
- 케이블: (행, 시트, [D,E,F,G,H열], "cable") -> SM(D)\nMM(E)\nCAT6(F)\nCAT5e(G)\nCAT5(H)
- 전부하: (행, 시트, C열, D열, "fullload") -> C(Mbps)\nD(Mbps)
"""
import os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CNE_DIR = os.path.join(BASE_DIR, "CNE")
OUTPUT_DIR = os.path.join(CNE_DIR, "학교별_리포트_V1.1")

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

# 통합 통계 파일 (TOTAL_MEASURE_LIST_V1.xlsx 사용)
TOTAL_MEASURE_LIST = os.path.join(CNE_DIR, "TOTAL_MEASURE_LIST_V1.xlsx")

# F열 데이터 매핑: (행, 시트, 열1, 열2_또는_리스트, format_type)
# format_type: None=단순, "cable"=케이블통계, "fullload"=전부하 Mbps
# 열: A=1, B=2, ..., P=16, Q=17, AU=47, AV=48
J_OUTPUT_MAP = [
    (2, "학교별통신장비현황", 16, None, None),       # F2: P열
    (3, "h_only", None, None, "h_only"),           # F3: (H3만 판정, F3 유지)
    (4, "POE", 3, None, None),                     # F4: C열
    (6, "케이블통계", [4, 5, 6, 7, 8], None, "cable"),  # F6: D,E,F,G,H (H6 무조건 정상)
    (10, "AP_장비통계", 3, None, None),            # F10: C열
    (11, "ISP측정_학교별평균", 6, None, None),     # F11: F열
    (12, "fixed", "80 Mhz", None, "fixed"),       # F12: 고정값
    (14, "POE", 4, None, None),                   # F14: D열
    (15, "충남AP", 7, 8, "location5"),            # F15: G,H열 -> 위치5/위치5외/보관및확인불가
    (16, "h_only", None, None, "h_only"),         # F16: (H16만 판정, F16 유지)
    (17, "충남AP", [4, 5, 6], None, "n_ac_ax"),   # F17: D,E,F열 -> N/AC/AX식
    (18, "학교별통신장비현황", 47, None, None),    # F18: AU열
    (19, "케이블통계", 8, None, None),             # F19: H열
    (20, "학교별통신장비현황", 9, None, None),     # F20: I열
    (21, "전부하측정_학교별평균", 3, 4, "fullload"),  # F21: C,D -> C(Mbps)\nD(Mbps)
    (22, "전부하측정_학교별평균", 3, 4, None),     # F22: C/D (판정 375 이상)
    (23, "ISP측정_학교별평균", 3, None, None),     # F23: C열
    (24, "ISP측정_학교별평균", 4, None, None),     # F24: D열
    (25, "ISP측정_학교별평균", 6, None, None),    # F25: F열
    (26, "POE", 4, None, None),                   # F26: D열
    (28, "집선ISP", 5, None, None),               # F28: E열
    (29, "집선ISP", 3, None, None),               # F29: C열
    (30, "집선ISP", 4, None, None),               # F30: D열
    (31, "CNE_WIRED_MEANSURE_AVG", 5, None, None),   # F31: E열
    (32, "CNE_WIRED_MEANSURE_AVG", 6, None, None),   # F32: F열
    (33, "CNE_WIRED_MEANSURE_AVG", 4, None, None),   # F33: D열
    (34, "CNE_WIRED_MEANSURE_AVG", 4, None, None),   # F34: D열
]

# H열 판정: (행, 연산, 기준값)  연산: ge=이상, le=이하
# 특수: H2 "0식", H12 "무조건 정상", H18 "계위 앞 숫자", H19 "0", H20 "분리"
L_JUDGMENT_MAP = [
    (2, "le", 0),      # H2: 0식
    (3, "always", "정상"),  # H3: 무조건 정상
    (6, "always", "정상"),  # H6: 무조건 정상
    (4, "le", 12),     # H4: 12개 이하
    (10, "zero_or_empty_ok", 0),  # H10: 0 또는 없으면 정상
    (11, "ge", -60),   # H11: -60 이상
    (12, "always", "정상"),  # H12: 무조건 정상
    (14, "le", 3),     # H14: 3 이하
    (16, "always", "정상"),  # H16: 무조건 정상
    (15, "has_value", None),  # H15: H열 값 있으면 개선필요
    (17, "has_value", None),  # H17: D열 값 있으면 개선필요
    (18, "ge_before_keyword", 3),  # H18: "계위" 앞 숫자 3 이상이면 개선필요
    (19, "le", 0),     # H19: 0
    (20, "split_exact", 0),  # H20: "분리"=정상, "미분리"=개선필요
    (21, "both_ge", 375),  # H21: C,D 중 1개라도 375 이하면 개선필요
    (22, "ge", 375),   # H22: 375 이상
    (23, "ge", 375),   # H23: 375 이상
    (24, "ge", 375),   # H24: 375 이상
    (25, "ge", -60),   # H25: -60 이상
    (26, "le", 3),     # H26: 3 이하
    (28, "le", 10),    # H28: 10 이하
    (29, "ge", 700),   # H29: 700 이상
    (30, "ge", 700),   # H30: 700 이상
    (31, "le", 1.0),   # H31: 1.0% 이하
    (32, "le", 5),     # H32: 5 이하
    (33, "ge", 700),   # H33: 700 이상
    (34, "ge", 700),   # H34: 700 이상
]

J_COL = 6   # F열 측정값
L_COL = 8   # H열 평가결과
G_COL = 7   # G열

# H2~H34 판정 행 (정상/개선필요 카운트 대상)
JUDGMENT_ROW_START = 2
JUDGMENT_ROW_END = 34

# 측정값 폰트 검정색 적용 행 (집선ISP 등)
FONT_BLACK_ROWS = (28, 29, 30)

# H열 판정 시 v2(두번째 열) 값으로 판정하는 행 (예: F15는 H열로 판정)
JUDGE_BY_V2_ROWS = (15,)

# H열 판정 시 v1, v2 둘 다 필요한 행 (예: H21은 C,D 둘 다 375 이상 체크)
JUDGE_BOTH_ROWS = (21,)

# 로그 (V1.1)
LOG_DIR = os.path.join(BASE_DIR, "logs")
LOG_PREFIX = "school_report_v1_1"
