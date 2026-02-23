# -*- coding: utf-8 -*-
"""
측정 데이터 전처리 공통 유틸

- 관리번호에서 학교코드 추출 (ISP, 전부하 등 공통 사용)
"""
SCHOOL_CODE_LEN = 12


def extract_school_code_from_mgmt_num(mgmt_val):
    """
    관리번호에서 학교코드 추출
    - 앞 12자리, '-' 이전까지
    - 공백 제거, 대소문자 통일(그룹핑 시 사용)
    - 예: 'G107441266MS-001' → 'G107441266MS'

    Args:
        mgmt_val: 장비관리번호 값 (str, int, float 등)

    Returns:
        str: 학교코드 (12자리 이내), 없으면 ''
    """
    if mgmt_val is None:
        return ""
    s = str(mgmt_val).strip()
    if not s:
        return ""
    part = s.split("-")[0].strip() if "-" in s else s.strip()
    code = part[:SCHOOL_CODE_LEN]
    # 대소문자 통일 (N108231014Ms vs N108231014MS 중복 방지)
    return code.upper()
