# -*- coding: utf-8 -*-
"""
학교별 장비 분리 스크립트 - 원본 파일·시트 규칙 정의

[폴더 규칙]
- BASE_DIR: 스크립트 기준 작업 폴더
- DNI_DIR: BASE_DIR/DNI (대전 AP 등 DNI 전용 원본)
- 그 외 장비: BASE_DIR에 직접 배치

[파일명 규칙]
- 대전(DNI): DNI 폴더 또는 작업폴더
  - AP: DNI/DNI_AP_LIST.XLSX
  - 스위치: DNI/DNI_SWITCH_LIST.xlsx
  - 보안: DNI/DNI_SEUTM_LIST.xlsx
  - POE: DNI_POE_LIST.xlsx (작업폴더)
- 충남(CNE): 작업폴더
  - AP: CNE_AP_LIST.xlsx
  - 스위치: CNE_SWITCH_LIST.xlsx
  - 보안: CNE_SEUTM_LIST.xlsx
  - POE: CNE_POE_LIST.xlsx

[시트 규칙]
- 장비별 시트명 우선순위 (앞쪽 우선 사용)
- 공통 패턴: {장비}_{지역}전체, {장비명}, Sheet1
"""
import os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DNI_DIR = os.path.join(BASE_DIR, "DNI")
CNE_DIR = os.path.join(BASE_DIR, "CNE")

# ========== 원본 파일 경로 규칙 ==========
# 지역별 기본 폴더 (원본 탐색 시 사용)
SOURCE_DIR_BY_REGION = {
    "DNI": DNI_DIR,   # 대전: DNI 폴더 우선
    "CNE": BASE_DIR,  # 충남: 작업 폴더
}

# 장비별 원본 파일명 규칙 (지역코드_장비 → (폴더, 파일명))
# 폴더: "DNI"=DNI_DIR, "CNE"=CNE_DIR, ""=BASE_DIR(작업폴더)
SOURCE_FILENAME_RULES = {
    ("DNI", "AP"): ("DNI", "DNI_AP_LIST.XLSX"),
    ("CNE", "AP"): ("CNE", "CNE_AP_LIST.xlsx"),  # 작업폴더
    ("DNI", "switch"): ("DNI", "DNI_SWITCH_LIST.xlsx"),
    ("CNE", "switch"): ("CNE", "CNE_SWITCH_LIST.xlsx"),
    ("DNI", "security"): ("DNI", "DNI_SEUTM_LIST.xlsx"),
    ("CNE", "security"): ("CNE", "CNE_SEUTM_LIST.xlsx"),
    ("DNI", "poe"): ("DNI", "DNI_POE_LIST.xlsx"),
    ("CNE", "poe"): ("CNE", "CNE_POE_LIST.xlsx"),
}


def get_source_path(region_key, equipment):
    """지역·장비에 따른 원본 파일 전체 경로 반환"""
    key = (region_key, equipment)
    folder_rel, filename = SOURCE_FILENAME_RULES.get(key, ("", ""))
    if not filename:
        return None
    if folder_rel == "DNI":
        base = DNI_DIR
    elif folder_rel == "CNE":
        base = CNE_DIR
    else:
        base = BASE_DIR
    return os.path.join(base, filename)


# ========== 시트 규칙 ==========
# 장비별 시트명 우선순위 (앞쪽부터 탐색, 존재하는 첫 시트 사용)
SHEET_PRIORITY_BY_EQUIPMENT = {
    "AP": [
        "AP자산", "AP",
    ],
    "switch": [
        "스위치", "Switch",
    ],
    "security": [
        "보안장비", "Security",
    ],
    "poe": [
        "POE자산", "POE",
    ],
}


def get_sheet_candidates(equipment):
    """장비에 따른 시트명 후보 목록"""
    return SHEET_PRIORITY_BY_EQUIPMENT.get(equipment, ["Sheet1"])


# ========== 출력 규칙 ==========
OUTPUT_BASE_BY_REGION = {
    "DNI": r"Y:\DJE_DNI\_지역별 산출물 취합\_00.별첨자료_학교별\DJE",
    "CNE": r"Y:\CNE_DNI\_지역별 산출물 취합\_00.별첨자료_학교별\CNE",
}
OUTPUT_BASE_TEST = os.path.join(BASE_DIR, "OUTPUT")
