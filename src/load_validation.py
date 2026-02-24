"""
VA/CFG 로드 결과 검증.

- 717 = 학교 수. 장비 시트에는 학교당 여러 대가 있으므로 행 수는 717보다 훨씬 많아야 함.
- 검증: (1) 장비 시트에 등장하는 유니크 학교 수 >= EXPECTED_MIN_SCHOOLS
        (2) 장비 시트 총 행 수 >= EXPECTED_MIN_EQUIPMENT_ROWS (데이터를 제대로 찾았는지)
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

import pandas as pd

from .sheet_defs import VA_DATA_SHEETS_EQUIPMENT, CFG_DATA_SHEETS

# 정상 719개 학교, 최소 717개 학교. 장비 시트에 등장하는 유니크 학교 수가 717 미만이면 잘못 읽은 것.
EXPECTED_MIN_SCHOOLS = 717
# 장비 시트 총 행 수: 학교당 여러 대이므로 717보다 훨씬 많아야 함. 이보다 적으면 시트/헤더가 다른 형태.
EXPECTED_MIN_EQUIPMENT_ROWS = 2000

# 하위 호환: 예전 검증에서 "최소 행 수"로 쓰이던 값 → 장비 최소 행 수로 통일
EXPECTED_MIN_ROWS = EXPECTED_MIN_EQUIPMENT_ROWS


def _count_data_rows(df: pd.DataFrame, mgmt_col: str = "관리번호") -> int:
    """관리번호가 채워진 유효 데이터 행 수 (헤더 제외)."""
    if df.empty or mgmt_col not in df.columns:
        return 0
    col = df[mgmt_col]
    if isinstance(col, pd.DataFrame):
        col = col.iloc[:, 0]
    filled = col.notna() & (col.astype(str).str.strip() != "")
    return len(df.loc[filled])


def _count_unique_schools(df: pd.DataFrame, mgmt_col: str = "관리번호") -> int:
    """관리번호 컬럼에서 '학교코드-...' prefix(학교코드) 유니크 개수."""
    if df.empty or mgmt_col not in df.columns:
        return 0
    col = df[mgmt_col]
    if isinstance(col, pd.DataFrame):
        col = col.iloc[:, 0]
    prefixes = col.dropna().astype(str).str.strip()
    codes = set()
    for v in prefixes:
        if "-" in v:
            codes.add(v.split("-", 1)[0].strip())
    return len(codes)


def validate_va_loaded(
    va_data: dict[str, pd.DataFrame],
    source_path: Path | str | None = None,
    min_schools: int = EXPECTED_MIN_SCHOOLS,
    min_rows: int = EXPECTED_MIN_EQUIPMENT_ROWS,
) -> list[dict[str, Any]]:
    """
    가상자산 로드 결과 검증.
    장비 시트: 유니크 학교 수 >= min_schools, 총 행 수 >= min_rows (717학교 × 여러 대이므로 행이 훨씬 많아야 함).
    Returns: [{"sheet", "unique_schools", "total_rows", "expected_min_schools", "expected_min_rows", "message", "ok"}, ...]
    """
    issues = []
    for sheet in VA_DATA_SHEETS_EQUIPMENT:
        if sheet not in va_data:
            issues.append({
                "sheet": sheet,
                "unique_schools": 0,
                "total_rows": 0,
                "expected_min_schools": min_schools,
                "expected_min_rows": min_rows,
                "message": f"시트 없음 (파일에 '{sheet}' 시트가 없거나 이름이 다름)",
                "ok": False,
            })
            continue
        df = va_data[sheet]
        total = _count_data_rows(df)
        schools = _count_unique_schools(df)
        ok = schools >= min_schools and total >= min_rows
        if not ok:
            msg_parts = []
            if schools < min_schools:
                msg_parts.append(f"유니크 학교 {schools}개 (최소 {min_schools}개)")
            if total < min_rows:
                msg_parts.append(f"총 행 {total}건 (최소 {min_rows}건)")
            issues.append({
                "sheet": sheet,
                "unique_schools": schools,
                "total_rows": total,
                "expected_min_schools": min_schools,
                "expected_min_rows": min_rows,
                "message": "데이터 부족: " + ", ".join(msg_parts) + ". 시트명/헤더가 다를 수 있음.",
                "ok": False,
            })
        else:
            issues.append({
                "sheet": sheet,
                "unique_schools": schools,
                "total_rows": total,
                "expected_min_schools": min_schools,
                "expected_min_rows": min_rows,
                "message": f"OK: {schools}개 학교, {total}건",
                "ok": True,
            })
    return issues


def validate_cfg_loaded(
    cfg_data: dict[str, pd.DataFrame],
    source_path: Path | str | None = None,
    min_schools: int = EXPECTED_MIN_SCHOOLS,
    min_rows: int = EXPECTED_MIN_EQUIPMENT_ROWS,
) -> list[dict[str, Any]]:
    """구성정보 로드 결과 검증. 장비 시트: 유니크 학교 수·총 행 수 모두 기준 충족."""
    issues = []
    for sheet in CFG_DATA_SHEETS:
        if sheet not in cfg_data:
            issues.append({
                "sheet": sheet,
                "unique_schools": 0,
                "total_rows": 0,
                "expected_min_schools": min_schools,
                "expected_min_rows": min_rows,
                "message": f"시트 없음 (파일에 '{sheet}' 시트가 없거나 이름이 다름)",
                "ok": False,
            })
            continue
        df = cfg_data[sheet]
        mgmt = "관리번호" if "관리번호" in df.columns else "장비관리번호"
        total = _count_data_rows(df, mgmt)
        schools = _count_unique_schools(df, mgmt)
        ok = schools >= min_schools and total >= min_rows
        if not ok:
            msg_parts = []
            if schools < min_schools:
                msg_parts.append(f"유니크 학교 {schools}개 (최소 {min_schools}개)")
            if total < min_rows:
                msg_parts.append(f"총 행 {total}건 (최소 {min_rows}건)")
            issues.append({
                "sheet": sheet,
                "unique_schools": schools,
                "total_rows": total,
                "expected_min_schools": min_schools,
                "expected_min_rows": min_rows,
                "message": "데이터 부족: " + ", ".join(msg_parts) + ". 시트명/헤더가 다를 수 있음.",
                "ok": False,
            })
        else:
            issues.append({
                "sheet": sheet,
                "unique_schools": schools,
                "total_rows": total,
                "expected_min_schools": min_schools,
                "expected_min_rows": min_rows,
                "message": f"OK: {schools}개 학교, {total}건",
                "ok": True,
            })
    return issues


def detect_sheet_structure(
    path: Path,
    file_kind: str = "va",
    max_header_try: int = 5,
) -> list[dict[str, Any]]:
    """
    엑셀 파일의 모든 시트를 열어, 관리번호/관리코드/장비관리번호/장비관리코드 해석 후 데이터 행 수·유니크 학교 수 반환.
    Returns: [{"sheet_name", "header_row_0based", "data_rows", "unique_schools", "has_mgmt_col"}, ...]
    """
    from .sheet_defs import normalize_mgmt_column

    xl = pd.ExcelFile(path, engine="openpyxl")
    result = []
    for sheet_name in xl.sheet_names:
        for header_0 in range(max_header_try):
            try:
                df = pd.read_excel(path, sheet_name=sheet_name, header=header_0, engine="openpyxl")
                if df.empty:
                    continue
                normalized = normalize_mgmt_column(df, file_kind)
                if "관리번호" not in normalized.columns:
                    continue
                n = _count_data_rows(normalized)
                schools = _count_unique_schools(normalized)
                result.append({
                    "sheet_name": sheet_name,
                    "header_row_0based": header_0,
                    "data_rows": n,
                    "unique_schools": schools,
                    "has_mgmt_col": True,
                })
                break
            except Exception:
                continue
        else:
            result.append({
                "sheet_name": sheet_name,
                "header_row_0based": None,
                "data_rows": 0,
                "unique_schools": 0,
                "has_mgmt_col": False,
            })
    xl.close()
    return result


def run_validation_report(
    va_path: Path | None = None,
    cfg_path: Path | None = None,
    min_schools: int = EXPECTED_MIN_SCHOOLS,
    min_rows: int = EXPECTED_MIN_EQUIPMENT_ROWS,
) -> dict[str, object]:
    """VA/CFG 파일 로드 후 검증 + 구조 탐지 보고. CLI용."""
    from .load_excel import load_va_data_sheets, load_cfg_data_sheets

    report: dict[str, object] = {"va": None, "cfg": None}
    if va_path and va_path.exists():
        va_data = load_va_data_sheets(va_path)
        va_issues = validate_va_loaded(va_data, va_path, min_schools=min_schools, min_rows=min_rows)
        det = detect_sheet_structure(va_path, "va")
        report["va"] = {
            "path": str(va_path),
            "validation": va_issues,
            "all_sheets_detection": det,
        }
    if cfg_path and cfg_path.exists():
        cfg_data = load_cfg_data_sheets(cfg_path)
        cfg_issues = validate_cfg_loaded(cfg_data, cfg_path, min_schools=min_schools, min_rows=min_rows)
        det = detect_sheet_structure(cfg_path, "cfg")
        report["cfg"] = {
            "path": str(cfg_path),
            "validation": cfg_issues,
            "all_sheets_detection": det,
        }
    return report


if __name__ == "__main__":
    import sys
    from pathlib import Path

    min_schools, min_rows = EXPECTED_MIN_SCHOOLS, EXPECTED_MIN_EQUIPMENT_ROWS
    paths = [Path(p) for p in sys.argv[1:] if not p.startswith("--")]
    for a in sys.argv:
        if a.startswith("--min-schools="):
            min_schools = int(a.split("=", 1)[1])
        elif a.startswith("--min-rows="):
            min_rows = int(a.split("=", 1)[1])
    va_path = paths[0] if len(paths) > 0 else None
    cfg_path = paths[1] if len(paths) > 1 else None
    r = run_validation_report(va_path=va_path, cfg_path=cfg_path, min_schools=min_schools, min_rows=min_rows)
    print("기대: 유니크 학교 수 >=", min_schools, ", 장비 시트 총 행 수 >=", min_rows)
    for kind, data in [("가상자산", r.get("va")), ("구성정보", r.get("cfg"))]:
        if data is None:
            continue
        print(f"\n=== {kind} {data['path']} ===")
        for u in data["validation"]:
            mark = "OK" if u["ok"] else "FAIL"
            print(f"  [{mark}] {u['sheet']}: {u.get('unique_schools', u.get('actual', 0))}개 학교, {u.get('total_rows', 0)}건")
        print("  시트 구조 탐지:")
        for d in data["all_sheets_detection"]:
            if d.get("data_rows", 0) > 0:
                print(f"    시트={d['sheet_name']!r} header_0based={d['header_row_0based']} → {d['data_rows']}행, {d.get('unique_schools', 0)}개 학교")
