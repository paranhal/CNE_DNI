"""
장비별·가상/구성별 파일 생성 기반 설정.

- 장비 종류: PoE, AP, 스위치, 보안장비 (VA/CFG 공통, sheet_defs와 동기화)
- 데이터 종류: 가상자산, 구성정보
- 출력: df_{장비}_가상자산.csv, df_{장비}_구성정보.csv (대상 학교 필터 적용 후)
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

from .sheet_defs import VA_DATA_SHEETS_EQUIPMENT

# 장비 시트 순서 = 내보내기 순서 (sheet_defs 단일 소스)
EQUIPMENT_TYPES: tuple[str, ...] = VA_DATA_SHEETS_EQUIPMENT

# 데이터 종류(출력 파일 접미사)
DATA_KINDS = ("가상자산", "구성정보")

# 출력 파일명 패턴: df_{장비}_{가상자산|구성정보}.csv
def get_export_filename(equipment: str, kind: str) -> str:
    """장비·데이터종류별 출력 CSV 파일명."""
    if equipment not in EQUIPMENT_TYPES or kind not in DATA_KINDS:
        raise ValueError(f"equipment={equipment!r}, kind={kind!r}")
    return f"df_{equipment}_{kind}.csv"


def get_export_path(output_dir: Path, equipment: str, kind: str) -> Path:
    """출력 디렉터리 + 파일명."""
    return output_dir / get_export_filename(equipment, kind)


def list_export_tasks(
    va_data: dict[str, Any],
    cfg_data: dict[str, Any],
    output_dir: Path,
) -> list[tuple[Any, Path, str, str]]:
    """
    현재 VA/CFG 데이터에 따라 (df, path, equipment, kind) 작업 목록 생성.
    반환: [(df, output_path, equipment, kind), ...]
    """
    tasks = []
    for eq in EQUIPMENT_TYPES:
        if eq in va_data and va_data[eq] is not None:
            tasks.append((
                va_data[eq],
                get_export_path(output_dir, eq, "가상자산"),
                eq,
                "가상자산",
            ))
        if eq in cfg_data and cfg_data[eq] is not None:
            tasks.append((
                cfg_data[eq],
                get_export_path(output_dir, eq, "구성정보"),
                eq,
                "구성정보",
            ))
    return tasks
