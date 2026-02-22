"""
raw_data/구성, 자산, 충남을 모두 스캔해 로드한 뒤 병합합니다.
- 충남(학교별) 데이터를 기준으로, 구성·자산에만 있는 행을 추가해 최종 VA/CFG를 만듦.
- 장비: AP, PoE, 스위치, 보안장비 (파일명 SW→스위치, FW→보안장비)
"""

from __future__ import annotations

import unicodedata
from pathlib import Path
from typing import Any

import pandas as pd
from tqdm import tqdm

from .sheet_defs import (
    CFG_DATA_SHEETS,
    VA_DATA_SHEETS_EQUIPMENT,
    normalize_mgmt_column,
)
from .load_validation import _count_data_rows


EQUIPMENT_ALIAS = {
    "AP": "AP",
    "PoE": "PoE",
    "스위치": "스위치",
    "SW": "스위치",
    "보안장비": "보안장비",
    "FW": "보안장비",
}


def _equipment_from_filename(name: str) -> str | None:
    name = Path(name).stem
    for key, eq in EQUIPMENT_ALIAS.items():
        if f"_{key}_" in name or name.startswith(f"{key}_"):
            return eq
    return None


def _norm(s: str) -> str:
    return unicodedata.normalize("NFC", s)


def _list_구성_xlsx(root: Path) -> list[tuple[Path, str]]:
    out = []
    dir_ = root / "구성"
    if not dir_.exists():
        return out
    for p in dir_.rglob("*.xlsx"):
        n = _norm(p.name)
        if "일괄업로드용" not in n or "구성정보" not in n:
            continue
        eq = _equipment_from_filename(n)
        if eq:
            out.append((p, eq))
    return out


def _list_자산_xlsx(root: Path) -> list[tuple[Path, str]]:
    out = []
    dir_ = root / "자산"
    if not dir_.exists():
        return out
    for p in dir_.rglob("*.xlsx"):
        n = _norm(p.name)
        if "일괄업로드용" not in n or "가상자산" not in n:
            continue
        eq = _equipment_from_filename(n)
        if eq:
            out.append((p, eq))
    return out


def _list_충남_xlsx(root: Path) -> list[tuple[Path, str, str]]:
    """(path, equipment, 'va'|'cfg')"""
    out = []
    dir_ = root / "충남"
    if not dir_.exists():
        return out
    for p in dir_.rglob("*.xlsx"):
        n = _norm(p.name)
        if "일괄업로드용" not in n:
            continue
        eq = _equipment_from_filename(n)
        if not eq:
            continue
        if "가상자산" in n:
            out.append((p, eq, "va"))
        elif "구성정보" in n or "구성정봐" in n:
            out.append((p, eq, "cfg"))
    return out


def _load_sheet_robust(path: Path, file_kind: str) -> pd.DataFrame | None:
    try:
        xl = pd.ExcelFile(path, engine="openpyxl")
    except Exception:
        return None
    best_df = None
    best_count = 0
    for sheet_name in xl.sheet_names:
        for header_0 in range(4):
            try:
                df = pd.read_excel(xl, sheet_name=sheet_name, header=header_0, engine="openpyxl")
                normalized = normalize_mgmt_column(df, file_kind)
                if "관리번호" not in normalized.columns:
                    continue
                n = _count_data_rows(normalized)
                if n > best_count:
                    best_count = n
                    best_df = normalized
            except Exception:
                continue
    xl.close()
    return best_df


def _load_one_cfg(path: Path) -> pd.DataFrame | None:
    return _load_sheet_robust(path, "cfg")


def _load_one_va(path: Path) -> pd.DataFrame | None:
    return _load_sheet_robust(path, "va")


def _extract_mgmt_set(df: pd.DataFrame) -> set:
    if df is None or df.empty or "관리번호" not in df.columns:
        return set()
    col = df["관리번호"]
    if isinstance(col, pd.DataFrame):
        col = col.iloc[:, 0]
    return set(col.dropna().astype(str).str.strip().tolist())


def _concat_dedupe(dfs: list[pd.DataFrame], keep: str = "first") -> pd.DataFrame:
    valid = [d for d in dfs if d is not None and not d.empty]
    if not valid:
        return pd.DataFrame()
    combined = pd.concat(valid, ignore_index=True)
    if "관리번호" not in combined.columns:
        return combined
    return combined.drop_duplicates(subset=["관리번호"], keep=keep)


def _merge_adding_missing(base_dfs: list[pd.DataFrame], add_dfs: list[pd.DataFrame]) -> pd.DataFrame:
    base = _concat_dedupe(base_dfs, keep="first")
    if base.empty:
        return _concat_dedupe(add_dfs, keep="first")
    existing = _extract_mgmt_set(base)
    to_add = []
    for d in add_dfs:
        if d is None or d.empty or "관리번호" not in d.columns:
            continue
        col = d["관리번호"]
        if isinstance(col, pd.DataFrame):
            col = col.iloc[:, 0]
        mask = ~col.astype(str).str.strip().isin(existing)
        if mask.any():
            to_add.append(d.loc[mask].copy())
    if not to_add:
        return base
    return _concat_dedupe([base] + to_add, keep="first")


def load_and_merge_raw(
    raw_data_root: Path | None = None,
    progress: bool = True,
) -> tuple[dict[str, pd.DataFrame], dict[str, pd.DataFrame], dict[str, Any]]:
    """
    구성/자산/충남을 로드해 장비별 VA/CFG 병합.
    충남을 기준으로, 구성(CFG)·자산(VA)에서 충남에 없는 행만 추가.
    Returns: (va_data, cfg_data, report).
    """
    if raw_data_root is None:
        from .verify_raw_data import get_raw_data_root
        raw_data_root = get_raw_data_root()
    root = Path(raw_data_root)
    if not root.exists():
        raise FileNotFoundError(f"raw_data_root가 없습니다: {root}")

    report: dict[str, Any] = {
        "구성_파일_수": 0,
        "자산_파일_수": 0,
        "충남_파일_수": 0,
        "장비별_충남_va행": {},
        "장비별_충남_cfg행": {},
        "장비별_구성_추가행": {},
        "장비별_자산_추가행": {},
        "장비별_최종_va행": {},
        "장비별_최종_cfg행": {},
    }

    list_구성 = _list_구성_xlsx(root)
    list_자산 = _list_자산_xlsx(root)
    list_충남 = _list_충남_xlsx(root)
    report["구성_파일_수"] = len(list_구성)
    report["자산_파일_수"] = len(list_자산)
    report["충남_파일_수"] = len(list_충남)

    # 장비별로 (충남 VA, 자산 VA) / (충남 CFG, 구성 CFG) 수집
    충남_va: dict[str, list[pd.DataFrame]] = {eq: [] for eq in VA_DATA_SHEETS_EQUIPMENT}
    충남_cfg: dict[str, list[pd.DataFrame]] = {eq: [] for eq in CFG_DATA_SHEETS}
    자산_va: dict[str, list[pd.DataFrame]] = {eq: [] for eq in VA_DATA_SHEETS_EQUIPMENT}
    구성_cfg: dict[str, list[pd.DataFrame]] = {eq: [] for eq in CFG_DATA_SHEETS}

    it_충남 = tqdm(list_충남, desc="충남 로드", unit="파일") if progress else list_충남
    for path, eq, kind in it_충남:
        if kind == "va":
            df = _load_one_va(path)
            if df is not None and not df.empty and eq in 충남_va:
                충남_va[eq].append(df)
        else:
            df = _load_one_cfg(path)
            if df is not None and not df.empty and eq in 충남_cfg:
                충남_cfg[eq].append(df)

    it_구성 = tqdm(list_구성, desc="구성 로드", unit="파일") if progress else list_구성
    for path, eq in it_구성:
        df = _load_one_cfg(path)
        if df is not None and not df.empty and eq in 구성_cfg:
            구성_cfg[eq].append(df)

    it_자산 = tqdm(list_자산, desc="자산 로드", unit="파일") if progress else list_자산
    for path, eq in it_자산:
        df = _load_one_va(path)
        if df is not None and not df.empty and eq in 자산_va:
            자산_va[eq].append(df)

    va_data: dict[str, pd.DataFrame] = {}
    cfg_data: dict[str, pd.DataFrame] = {}

    for eq in VA_DATA_SHEETS_EQUIPMENT:
        base_va = 충남_va.get(eq, [])
        add_va = 자산_va.get(eq, [])
        base_va_only = _concat_dedupe(base_va) if base_va else pd.DataFrame()
        merged_va = _merge_adding_missing(base_va, add_va)
        va_data[eq] = merged_va
        base_va_rows = len(base_va_only)
        report["장비별_충남_va행"][eq] = base_va_rows
        report["장비별_자산_추가행"][eq] = max(0, len(merged_va) - base_va_rows)
        report["장비별_최종_va행"][eq] = len(merged_va)

    for eq in CFG_DATA_SHEETS:
        base_cfg = 충남_cfg.get(eq, [])
        add_cfg = 구성_cfg.get(eq, [])
        base_cfg_only = _concat_dedupe(base_cfg) if base_cfg else pd.DataFrame()
        merged_cfg = _merge_adding_missing(base_cfg, add_cfg)
        cfg_data[eq] = merged_cfg
        base_cfg_rows = len(base_cfg_only)
        report["장비별_충남_cfg행"][eq] = base_cfg_rows
        report["장비별_구성_추가행"][eq] = max(0, len(merged_cfg) - base_cfg_rows)
        report["장비별_최종_cfg행"][eq] = len(merged_cfg)

    return va_data, cfg_data, report


def run_full(
    raw_data_root: Path | None = None,
    output_dir: Path | None = None,
    max_workers: int = 4,
) -> dict[str, Any]:
    """
    구성/자산/충남 스캔 → 병합 → 대상 필터 → 장비별 CSV 저장 → 데이터 없는 학교 리스트 생성 → 보고서 저장.
    """
    import json
    from datetime import datetime

    if output_dir is None:
        try:
            from .config_loader import get_path
            output_dir = Path(get_path("output_root"))
        except Exception:
            output_dir = Path(__file__).resolve().parent.parent / "output"
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    va_data, cfg_data, report = load_and_merge_raw(raw_data_root=raw_data_root, progress=True)
    from .integrate_export import run_integrate_export
    export_result = run_integrate_export(
        va_data=va_data,
        cfg_data=cfg_data,
        output_dir=output_dir,
        max_workers=max_workers,
    )
    from .make_missing_school_list import run as run_missing
    missing_result = run_missing(output_dir=output_dir)

    report["실행시각"] = datetime.now().isoformat()
    report["export_파일"] = export_result.get("files", [])
    report["전혀_없는_학교_수"] = missing_result.get("전혀_없는_학교_수", 0)
    report["검증리스트_수"] = missing_result.get("검증리스트_전체_수", 0)

    report_path = output_dir / "raw_data_병합_보고.json"
    with open(report_path, "w", encoding="utf-8") as f:
        json.dump(report, f, ensure_ascii=False, indent=2)

    lines = [
        "=" * 60,
        "raw_data 병합·통합 export 보고",
        "=" * 60,
        f"실행: {report['실행시각']}",
        "",
        "[소스 파일 수]",
        f"  구성(구성정보): {report['구성_파일_수']}개",
        f"  자산(가상자산):  {report['자산_파일_수']}개",
        f"  충남(학교별):    {report['충남_파일_수']}개",
        "",
        "[장비별 행 수 (충남 → 구성/자산 추가 → 최종)]",
    ]
    for eq in VA_DATA_SHEETS_EQUIPMENT:
        c_va = report["장비별_충남_va행"].get(eq, 0)
        a_va = report["장비별_자산_추가행"].get(eq, 0)
        f_va = report["장비별_최종_va행"].get(eq, 0)
        c_cfg = report["장비별_충남_cfg행"].get(eq, 0)
        g_cfg = report["장비별_구성_추가행"].get(eq, 0)
        f_cfg = report["장비별_최종_cfg행"].get(eq, 0)
        lines.append(f"  {eq}")
        lines.append(f"    가상자산: 충남 {c_va} + 자산 추가 {a_va} → 최종 {f_va}")
        lines.append(f"    구성정보: 충남 {c_cfg} + 구성 추가 {g_cfg} → 최종 {f_cfg}")
    lines.extend([
        "",
        "[출력]",
        f"  장비별 CSV: {export_result.get('files', [])}",
        f"  데이터 전혀 없는 학교: {report['전혀_없는_학교_수']}개",
        f"  검증리스트 전체: {report['검증리스트_수']}개",
        "",
        f"상세 JSON: {report_path.name}",
        "=" * 60,
    ])
    report_txt = output_dir / "raw_data_병합_보고.txt"
    report_txt.write_text("\n".join(lines), encoding="utf-8")

    return {
        "report": report,
        "report_path": str(report_path),
        "report_txt": str(report_txt),
        "export_result": export_result,
        "missing_result": missing_result,
    }


if __name__ == "__main__":
    import sys
    r = run_full(max_workers=4)
    print(r["report_txt"])
    with open(r["report_txt"], encoding="utf-8") as f:
        print(f.read())
    if "error" in r.get("missing_result", {}):
        sys.exit(1)
