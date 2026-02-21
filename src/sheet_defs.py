"""
통합 DB 및 학교별 DB에서 '실데이터' 시트만 구분하여 로딩하기 위한 정의.

- 가상자산: 장비 시트(PoE, AP, 스위치, 보안장비) + 학교정보
- 구성정보: 장비 시트(AP, PoE, 스위치, 보안장비)
- 관리번호·설명 등 비실데이터 시트는 로딩 대상에서 제외.
"""

from __future__ import annotations

# 가상자산 DB에서 실데이터로 사용할 시트 (장비 4종 + 학교정보)
VA_DATA_SHEETS_EQUIPMENT = ("PoE", "AP", "스위치", "보안장비")
VA_DATA_SHEETS_ALL = (*VA_DATA_SHEETS_EQUIPMENT, "학교정보")

# 구성정보 DB에서 실데이터로 사용할 시트
CFG_DATA_SHEETS = ("AP", "PoE", "스위치", "보안장비")

# 장비 시트 1행 헤더에 있어야 하는 컬럼 (실데이터 시트 판별용)
VA_HEADER_MUST_HAVE = ("관리번호", "장비명")
CFG_HEADER_MUST_HAVE = ("관리번호", "장비관리번호")
SCHOOL_INFO_HEADER_MUST_HAVE = ("학교코드", "학교명")


def get_va_header_row(sheet_name: str) -> int:
    """가상자산 파일에서 해당 시트의 헤더 행(1-based)."""
    if sheet_name == "학교정보":
        return 3
    return 1


def get_cfg_header_row(sheet_name: str) -> int:
    """구성정보 파일에서 해당 시트의 헤더 행(1-based)."""
    return 1


def get_data_start_row(sheet_name: str, file_kind: str) -> int:
    """헤더 다음 행 = 데이터 시작 행(1-based). file_kind: 'va' | 'cfg'."""
    if file_kind == "va":
        return get_va_header_row(sheet_name) + 1
    return get_cfg_header_row(sheet_name) + 1


def is_va_data_sheet(sheet_name: str) -> bool:
    """가상자산에서 실데이터로 쓸 시트인지 여부."""
    return sheet_name in VA_DATA_SHEETS_ALL


def is_cfg_data_sheet(sheet_name: str) -> bool:
    """구성정보에서 실데이터로 쓸 시트인지 여부."""
    return sheet_name in CFG_DATA_SHEETS


def va_sheets_to_load(sheet_names: list[str]) -> list[str]:
    """가상자산 파일의 시트 목록 중 실데이터 시트만 순서 유지하여 반환."""
    return [s for s in VA_DATA_SHEETS_ALL if s in sheet_names]


def cfg_sheets_to_load(sheet_names: list[str]) -> list[str]:
    """구성정보 파일의 시트 목록 중 실데이터 시트만 순서 유지하여 반환."""
    return [s for s in CFG_DATA_SHEETS if s in sheet_names]
