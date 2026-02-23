# -*- coding: utf-8 -*-
"""
학교 관련 공통 유틸
- 관리번호 앞 12자리 = 학교코드 (원본에서 학교코드 열 삭제 시 사용)
- 로그 명명: split_log_{장비}_{지역}_{날짜}.csv (예: split_log_AP_DNI_20260222.csv)
"""
import os
from datetime import datetime

SCHOOL_CODE_LEN = 12

# 장비별 로그 prefix (split_log_{장비}_)
EQUIPMENT_LOG_NAMES = {"AP": "AP", "switch": "switch", "security": "security", "poe": "poe"}

# 지역별 학교 리스트 후보 (우선순위)
SCHOOL_LIST_BY_REGION = {
    "DNI": [
        "SCHOOL_REG_LIST_DNI.xlsx", "SCHOOL_REG_LIST_DNI.XLSX", "school_reg_list_dni.xlsx",
        "school_reg_list_DNI.csv", "SCHOOL_REG_LIST_DNI.csv",
        "SCHOOL_REG_LIST.XLSX", "school_reg_list.xlsx", "SCHOOL_REG_LIST.csv",
    ],
    "CNE": [
        "SCHOOL_REG_LIST_CNE.xlsx", "SCHOOL_REG_LIST_CNE.XLSX", "school_reg_list_cne.xlsx",
        "school_reg_list_CNE.csv", "SCHOOL_REG_LIST_CNE.csv",
        "SCHOOL_REG_LIST.XLSX", "school_reg_list.xlsx", "SCHOOL_REG_LIST.csv",
    ],
}


def get_school_list_path(region_key, base_dir=None):
    """지역(DNI/CNE)에 맞는 학교 리스트 파일 경로 반환 (존재하는 첫 번째)"""
    base_dir = base_dir or os.path.dirname(os.path.abspath(__file__))
    candidates = SCHOOL_LIST_BY_REGION.get(region_key, SCHOOL_LIST_BY_REGION["DNI"])
    for name in candidates:
        p = os.path.join(base_dir, name)
        if os.path.exists(p):
            return p
    return os.path.join(base_dir, candidates[0])  # 없으면 첫 후보 경로 (에러용)

# 처리 순서: 지역 순 (구/시군)
REGION_ORDER_BY_KEY = {
    "DNI": ['동구', '중구', '서구', '유성구', '대덕구'],  # 대전
    "CNE": ['계룡', '공주', '금산', '논산', '당진', '보령', '부여', '서산', '서천', '아산', '예산', '천안', '청양', '태안', '홍성'],  # 충남
}
REGION_ORDER = REGION_ORDER_BY_KEY["DNI"]  # 기본값 (하위호환)


def extract_school_code_from_mgmt_num(mgmt_num):
    """
    관리번호에서 학교코드 추출 (앞 12자리)
    관리번호가 None/빈값이면 '' 반환
    """
    if mgmt_num is None:
        return ''
    s = str(mgmt_num).strip()
    if not s:
        return ''
    return s[:SCHOOL_CODE_LEN]


def find_mgmt_col(ws, header_row=2):
    """
    헤더 행에서 '관리번호' 열 인덱스 반환 (1-based).
    없으면 None
    """
    for col in range(1, ws.max_column + 1):
        hdr = ws.cell(row=header_row, column=col).value
        if hdr and '관리번호' in str(hdr).strip():
            return col
    return None


def find_school_code_col(ws, header_row=2):
    """
    헤더 행에서 '학교코드' 열 인덱스 반환 (1-based).
    없으면 None (대전 DNI 등 A열에 학교코드 있는 경우 사용)
    """
    for col in range(1, ws.max_column + 1):
        hdr = ws.cell(row=header_row, column=col).value
        if hdr and '학교코드' in str(hdr).strip():
            return col
    return None


def sort_schools_by_region(schools_with_data, region_key="DNI"):
    """(school, rows) 리스트를 지역 순으로 정렬 (region_key: DNI/CNE)"""
    order = REGION_ORDER_BY_KEY.get(region_key, REGION_ORDER)
    def _key(item):
        school, _ = item
        r = school.get('region') or ''
        idx = order.index(r) if r in order else 999
        return (idx, r, school.get('code') or '')
    return sorted(schools_with_data, key=_key)


def get_split_log_path(equipment, region, base_dir=None, date=None, suffix=None):
    """
    split 로그 파일 경로 생성 (split_log_{장비}_{지역}_{날짜}.csv)
    - equipment: AP, switch, security, poe
    - region: DNI, CNE
    - date: YYYYMMDD (None이면 오늘)
    - suffix: --new-log 시 사용 (예: HHMMSS)
    """
    base_dir = base_dir or os.path.dirname(os.path.abspath(__file__))
    eq_name = EQUIPMENT_LOG_NAMES.get(equipment, equipment)
    dt = date or datetime.now().strftime("%Y%m%d")
    base = f"split_log_{eq_name}_{region}_{dt}"
    if suffix:
        return os.path.join(base_dir, f"{base}_{suffix}.csv")
    return os.path.join(base_dir, f"{base}.csv")


def get_split_log_prefix(equipment, region):
    """get_processed_school_codes 등에서 사용할 로그 파일 prefix"""
    eq_name = EQUIPMENT_LOG_NAMES.get(equipment, equipment)
    return f"split_log_{eq_name}_{region}_"


def get_output_cols(ws, header_row=2, exclude_school_code=True):
    """
    출력할 열 인덱스 목록 (1-based).
    exclude_school_code=True면 '학교코드' 열 제외
    """
    cols = []
    for col in range(1, ws.max_column + 1):
        hdr = ws.cell(row=header_row, column=col).value
        hdr_str = str(hdr).strip() if hdr else ''
        if exclude_school_code and hdr_str == '학교코드':
            continue
        cols.append(col)
    return cols if cols else list(range(1, ws.max_column + 1))
