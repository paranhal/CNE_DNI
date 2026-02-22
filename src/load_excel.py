"""
통합/학교별 가상자산·구성정보 엑셀에서 실데이터 시트만 읽기.

- 가상자산: PoE, AP, 스위치, 보안장비, 학교정보 시트만 로딩 (관리번호·설명 시트 제외)
- 구성정보: AP, PoE, 스위치, 보안장비 시트만 로딩
- 717 = 학교 수; 장비 행은 그보다 훨씬 많음. 유니크 학교 수·총 행 수 검증은 load_validation.
"""

from __future__ import annotations

from pathlib import Path

import pandas as pd

from .sheet_defs import (
    get_data_start_row,
    get_va_header_row,
    get_cfg_header_row,
    va_sheets_to_load,
    cfg_sheets_to_load,
    normalize_mgmt_column,
)
from .load_validation import EXPECTED_MIN_ROWS, _count_data_rows


def _read_sheet(
    path: Path,
    sheet_name: str,
    header_row: int,
) -> pd.DataFrame:
    """엑셀 시트를 header_row(1-based)를 헤더로 읽어 DataFrame 반환."""
    return pd.read_excel(
        path,
        sheet_name=sheet_name,
        header=header_row - 1,
        engine="openpyxl",
    )


def load_va_sheet(path: Path, sheet_name: str, header_row_1based: int | None = None) -> pd.DataFrame:
    """가상자산 파일에서 지정 시트만 로딩. header_row_1based 없으면 시트별 기본값."""
    header_row = header_row_1based if header_row_1based is not None else get_va_header_row(sheet_name)
    return _read_sheet(path, sheet_name, header_row)


def load_cfg_sheet(path: Path, sheet_name: str, header_row_1based: int | None = None) -> pd.DataFrame:
    """구성정보 파일에서 지정 시트만 로딩. header_row_1based 없으면 1행."""
    header_row = header_row_1based if header_row_1based is not None else get_cfg_header_row(sheet_name)
    return _read_sheet(path, sheet_name, header_row)


def _load_va_sheet_robust(path: Path, sheet_name: str) -> pd.DataFrame:
    """장비 시트: 헤더 1..4행 시도. 관리번호/관리코드(2개면 학교코드-장비 형식 있는 쪽) 해석 후 '관리번호'로 정규화."""
    best_df: pd.DataFrame | None = None
    best_count = 0
    for header_1 in range(1, 5):
        try:
            df = _read_sheet(path, sheet_name, header_1)
            normalized = normalize_mgmt_column(df, "va")
            if "관리번호" not in normalized.columns:
                continue
            n = _count_data_rows(normalized)
            if n > best_count:
                best_count = n
                best_df = normalized
        except Exception:
            continue
    if best_df is not None:
        return best_df
    df = load_va_sheet(path, sheet_name)
    return normalize_mgmt_column(df, "va")


def _load_cfg_sheet_robust(path: Path, sheet_name: str) -> pd.DataFrame:
    """구성정보 시트: 헤더 1..4행 시도. 장비관리번호/장비관리코드/관리번호/관리코드 해석 후 '관리번호'로 정규화."""
    best_df: pd.DataFrame | None = None
    best_count = 0
    for header_1 in range(1, 5):
        try:
            df = _read_sheet(path, sheet_name, header_1)
            normalized = normalize_mgmt_column(df, "cfg")
            if "관리번호" not in normalized.columns:
                continue
            n = _count_data_rows(normalized)
            if n > best_count:
                best_count = n
                best_df = normalized
        except Exception:
            continue
    if best_df is not None:
        return best_df
    df = load_cfg_sheet(path, sheet_name)
    return normalize_mgmt_column(df, "cfg")


def load_va_data_sheets(path: Path, robust_headers: bool = True) -> dict[str, pd.DataFrame]:
    """가상자산 파일에서 실데이터 시트만 로딩. robust_headers=True면 장비 시트는 헤더 1~4행 중 데이터 많은 행 사용."""
    xl = pd.ExcelFile(path, engine="openpyxl")
    to_load = va_sheets_to_load(xl.sheet_names)
    xl.close()
    out = {}
    for name in to_load:
        if name == "학교정보":
            out[name] = load_va_sheet(path, name)
        else:
            if robust_headers:
                out[name] = _load_va_sheet_robust(path, name)
            else:
                out[name] = normalize_mgmt_column(load_va_sheet(path, name), "va")
    return out


def load_cfg_data_sheets(path: Path, robust_headers: bool = True) -> dict[str, pd.DataFrame]:
    """구성정보 파일에서 실데이터 시트만 로딩. robust_headers=True면 헤더 1~4행 중 데이터 많은 행 사용."""
    xl = pd.ExcelFile(path, engine="openpyxl")
    to_load = cfg_sheets_to_load(xl.sheet_names)
    xl.close()
    out = {}
    for name in to_load:
        if robust_headers:
            out[name] = _load_cfg_sheet_robust(path, name)
        else:
            out[name] = normalize_mgmt_column(load_cfg_sheet(path, name), "cfg")
    return out


def sheet_info_va(path: Path) -> list[dict]:
    """가상자산 파일의 실데이터 시트별 헤더 행·데이터 시작 행 정보."""
    xl = pd.ExcelFile(path, engine="openpyxl")
    to_load = va_sheets_to_load(xl.sheet_names)
    xl.close()
    return [
        {
            "sheet": name,
            "header_row": get_va_header_row(name),
            "data_start_row": get_data_start_row(name, "va"),
        }
        for name in to_load
    ]


def sheet_info_cfg(path: Path) -> list[dict]:
    """구성정보 파일의 실데이터 시트별 헤더 행·데이터 시작 행 정보."""
    xl = pd.ExcelFile(path, engine="openpyxl")
    to_load = cfg_sheets_to_load(xl.sheet_names)
    xl.close()
    return [
        {
            "sheet": name,
            "header_row": get_cfg_header_row(name),
            "data_start_row": get_data_start_row(name, "cfg"),
        }
        for name in to_load
    ]
