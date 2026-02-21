"""
통합/학교별 가상자산·구성정보 엑셀에서 실데이터 시트만 읽기.

- 가상자산: PoE, AP, 스위치, 보안장비, 학교정보 시트만 로딩 (관리번호·설명 시트 제외)
- 구성정보: AP, PoE, 스위치, 보안장비 시트만 로딩
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
)


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


def load_va_sheet(path: Path, sheet_name: str) -> pd.DataFrame:
    """가상자산 파일에서 지정 시트만 로딩 (헤더 행은 시트별 정의 적용)."""
    header_row = get_va_header_row(sheet_name)
    return _read_sheet(path, sheet_name, header_row)


def load_cfg_sheet(path: Path, sheet_name: str) -> pd.DataFrame:
    """구성정보 파일에서 지정 시트만 로딩."""
    header_row = get_cfg_header_row(sheet_name)
    return _read_sheet(path, sheet_name, header_row)


def load_va_data_sheets(path: Path) -> dict[str, pd.DataFrame]:
    """가상자산 파일에서 실데이터 시트만 로딩. 키=시트명, 값=DataFrame."""
    xl = pd.ExcelFile(path, engine="openpyxl")
    to_load = va_sheets_to_load(xl.sheet_names)
    xl.close()
    return {name: load_va_sheet(path, name) for name in to_load}


def load_cfg_data_sheets(path: Path) -> dict[str, pd.DataFrame]:
    """구성정보 파일에서 실데이터 시트만 로딩. 키=시트명, 값=DataFrame."""
    xl = pd.ExcelFile(path, engine="openpyxl")
    to_load = cfg_sheets_to_load(xl.sheet_names)
    xl.close()
    return {name: load_cfg_sheet(path, name) for name in to_load}


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
