"""
통합 데이터를 장비별·가상/구성별 CSV로 내보내기.

- 719 대상만 필터 후 df_{장비}_가상자산.csv, df_{장비}_구성정보.csv 생성
- tqdm 진행률, 병렬 저장, 학교별 건수 로그
"""

from __future__ import annotations

import json
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from pathlib import Path
from typing import Any

import pandas as pd
from tqdm import tqdm

from .load_excel import load_va_data_sheets, load_cfg_data_sheets
from .sheet_defs import VA_DATA_SHEETS_EQUIPMENT, CFG_DATA_SHEETS
from .verify_schools import get_target_school_codes, filter_va_data_by_target

# 구성정보 통합 파일 기본 경로
DEFAULT_CFG_PATH = Path(
    "/Users/paranhal/Library/CloudStorage/GoogleDrive-paranhanl66@gmail.com"
    "/내 드라이브/20260215_D드라이브/Lee_20260202/충남_통합_구성정보DB_20260203_103132.xlsx"
)


def _extract_school_code_prefix(series: pd.Series) -> pd.Series:
    """관리번호 컬럼 시리즈에서 학교코드(prefix) 추출."""
    def one(val):
        if pd.isna(val) or "-" not in str(val):
            return None
        return str(val).strip().split("-", 1)[0].strip()
    return series.apply(one)


def filter_cfg_data_by_target(
    cfg_data: dict[str, pd.DataFrame],
    target_codes: set[str] | None = None,
) -> dict[str, pd.DataFrame]:
    """구성정보를 719 대상 학교만 남김."""
    if target_codes is None:
        target_codes = get_target_school_codes()
    if not target_codes:
        return cfg_data
    out = {}
    for sheet in CFG_DATA_SHEETS:
        if sheet not in cfg_data or "관리번호" not in cfg_data[sheet].columns:
            continue
        df = cfg_data[sheet]
        prefixes = _extract_school_code_prefix(df["관리번호"])
        out[sheet] = df[prefixes.isin(target_codes)].copy()
    return out


def _school_counts(df: pd.DataFrame, mgmt_col: str = "관리번호") -> dict[str, int]:
    """DataFrame에서 관리번호 prefix별 행 수."""
    if df.empty or mgmt_col not in df.columns:
        return {}
    prefixes = _extract_school_code_prefix(df[mgmt_col])
    return prefixes.dropna().value_counts().astype(int).to_dict()


def _export_one(
    df: pd.DataFrame,
    out_path: Path,
    equipment: str,
    kind: str,
) -> dict[str, Any]:
    """한 개 CSV 저장 및 학교별 건수 반환."""
    out_path.parent.mkdir(parents=True, exist_ok=True)
    df.to_csv(out_path, index=False, encoding="utf-8-sig")
    counts = _school_counts(df)
    return {
        "file": str(out_path.name),
        "path": str(out_path),
        "equipment": equipment,
        "kind": kind,
        "total_rows": len(df),
        "school_count": len(counts),
        "per_school": counts,
    }


def run_integrate_export(
    va_path: Path | str | None = None,
    cfg_path: Path | str | None = None,
    output_dir: Path | str | None = None,
    max_workers: int = 4,
) -> dict[str, Any]:
    """
    가상자산·구성정보 로드 → 719 필터 → 장비별 CSV 8개 저장 + 로그.
    Returns: {"output_dir", "files": [...], "log_path", "summary"}.
    """
    va_path = Path(va_path) if va_path else _get_va_path()
    cfg_path = Path(cfg_path) if cfg_path else DEFAULT_CFG_PATH
    if output_dir is None:
        output_dir = Path(__file__).resolve().parent.parent / "output"
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    target_codes = get_target_school_codes()
    if not target_codes:
        raise RuntimeError("719 대상 학교 리스트를 로드할 수 없습니다.")

    # 1) 가상자산 로드 (한 번에 실데이터 시트만)
    with tqdm(total=1, desc="가상자산 로드", unit="파일") as pbar:
        va_data = load_va_data_sheets(va_path)
        pbar.update(1)
    va_data = filter_va_data_by_target(va_data, target_codes)

    # 2) 구성정보 로드
    with tqdm(total=1, desc="구성정보 로드", unit="파일") as pbar:
        cfg_data = load_cfg_data_sheets(cfg_path)
        pbar.update(1)
    cfg_data = filter_cfg_data_by_target(cfg_data, target_codes)

    # 3) 8개 파일 저장 (병렬)
    tasks = []
    for eq in VA_DATA_SHEETS_EQUIPMENT:
        if eq in va_data:
            tasks.append((va_data[eq], output_dir / f"df_{eq}_가상자산.csv", eq, "가상자산"))
        if eq in cfg_data:
            tasks.append((cfg_data[eq], output_dir / f"df_{eq}_구성정보.csv", eq, "구성정보"))

    results = []
    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        futures = {
            executor.submit(_export_one, df, path, eq, kind): (eq, kind)
            for df, path, eq, kind in tasks
        }
        for fut in tqdm(as_completed(futures), total=len(futures), desc="CSV 저장", unit="파일"):
            results.append(fut.result())

    # 4) 로그 기록
    log_path = output_dir / "통합_export_로그.txt"
    log_lines = [
        f"통합 내보내기 실행: {datetime.now().isoformat()}",
        f"가상자산: {va_path}",
        f"구성정보: {cfg_path}",
        f"대상 학교 수(719): {len(target_codes)}",
        f"출력 디렉터리: {output_dir}",
        "",
        "※ 통합 파일의 '학교 수'가 719보다 작은 이유:",
        "  719 = 분석 대상 학교 목록(충남_대상학교_리스트.xlsx)입니다.",
        "  원본(가상자산/구성정보)에 해당 장비 데이터가 1건이라도 있는 학교만 통합됩니다.",
        "  따라서 원본에 그 장비 데이터가 없는 학교는 '학교 수'에 포함되지 않습니다.",
        "  → 719 중 데이터가 없는 학교 목록: 719중_장비별_데이터없는_학교.csv",
        "",
    ]
    summary = {}
    for r in results:
        key = f"{r['equipment']}_{r['kind']}"
        summary[key] = {"total_rows": r["total_rows"], "school_count": r["school_count"], "file": r["file"]}
        log_lines.append(f"--- {r['file']} ---")
        log_lines.append(f"  총 행: {r['total_rows']}, 학교 수: {r['school_count']}")
        for sc, cnt in sorted(r["per_school"].items(), key=lambda x: -x[1])[:50]:
            log_lines.append(f"    {sc}: {cnt}")
        if len(r["per_school"]) > 50:
            log_lines.append(f"    ... 외 {len(r['per_school']) - 50}개 학교")
        log_lines.append("")
    with open(log_path, "w", encoding="utf-8") as f:
        f.write("\n".join(log_lines))

    # JSON 요약 (학교별 건수 전체)
    summary_path = output_dir / "통합_export_요약.json"
    per_file_school = {f"{r['equipment']}_{r['kind']}": r["per_school"] for r in results}
    with open(summary_path, "w", encoding="utf-8") as f:
        json.dump(
            {
                "timestamp": datetime.now().isoformat(),
                "target_school_count": len(target_codes),
                "files": [{"file": r["file"], "total_rows": r["total_rows"], "school_count": r["school_count"]} for r in results],
                "per_file_per_school": per_file_school,
            },
            f,
            ensure_ascii=False,
            indent=2,
        )

    # 719 중 장비별로 데이터가 없는 학교 목록 (원본에 해당 장비 데이터가 없음)
    from .verify_schools import load_target_school_list
    target_df = load_target_school_list()
    missing_path = None
    if not target_df.empty:
        missing_report = target_df.copy()
        for key, counts in per_file_school.items():
            missing_report[key] = missing_report["학교코드"].map(lambda c: 1 if c in counts else 0)
        missing_path = output_dir / "719중_장비별_데이터없는_학교.csv"
        missing_report.to_csv(missing_path, index=False, encoding="utf-8-sig")
        log_lines.insert(
            11,  # "출력 디렉터리" 다음, ※ 설명 다음에 한 줄 추가
            f"719 중 장비별 데이터 유무: {missing_path.name} (1=있음, 0=없음)",
        )
    with open(log_path, "w", encoding="utf-8") as f:
        f.write("\n".join(log_lines))

    return {
        "output_dir": str(output_dir),
        "files": [r["file"] for r in results],
        "log_path": str(log_path),
        "summary_path": str(summary_path),
        "summary": summary,
    }


def _get_va_path() -> Path:
    """가상자산 경로 (수정본 우선)."""
    proj = Path(__file__).resolve().parent.parent
    revised = proj / "output" / "충남_통합_가상자산DB_Lee_수정본.xlsx"
    if revised.exists():
        return revised
    from .data_quality import DEFAULT_VA_PATH
    return DEFAULT_VA_PATH


def main() -> None:
    import sys
    try:
        from .config_loader import get_path, ensure_runtime_dirs
        ensure_runtime_dirs()
        output_dir = get_path("output_root")
    except Exception:
        output_dir = Path(__file__).resolve().parent.parent / "output"

    n = int(sys.argv[1]) if len(sys.argv) > 1 else 4
    r = run_integrate_export(output_dir=output_dir, max_workers=n)
    print("저장된 파일:", r["files"])
    print("로그:", r["log_path"])
    print("요약 JSON:", r["summary_path"])


if __name__ == "__main__":
    main()
