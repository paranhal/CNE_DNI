"""
대상 학교 수 검증 및 누락 학교 확인.

- 대상 학교 리스트: output/CNE_LIST.xlsx (충남 작업 기준)
- 가상자산 학교정보 시트와 비교 → 대상인데 데이터에 없는 학교 / 데이터에 있지만 비대상 학교
- 품질 검사 재실행하여 미수정 오류 잔여 건수 확인
"""

from __future__ import annotations

from pathlib import Path

import pandas as pd

from .load_excel import load_va_data_sheets
from .data_quality import run_va_quality_checks, DEFAULT_VA_PATH

_PROJECT_ROOT = Path(__file__).resolve().parent.parent

# 대상 학교 리스트 파일 (학교코드·학교명 기준). 설정에서 target_school_list로 덮어쓸 수 있음.
DEFAULT_TARGET_SCHOOL_LIST_PATH = _PROJECT_ROOT / "output" / "CNE_LIST.xlsx"


def _resolve_target_list_path(path: Path | str | None) -> Path:
    """대상 학교 리스트 경로 결정: 인자 > config target_school_list > 기본(output/CNE_LIST.xlsx)."""
    if path is not None:
        return Path(path).resolve()
    try:
        from .config_loader import get_path
        return get_path("target_school_list")
    except (KeyError, FileNotFoundError):
        pass
    p = DEFAULT_TARGET_SCHOOL_LIST_PATH
    return p.resolve() if p.is_absolute() else (_PROJECT_ROOT / p).resolve()


def load_target_school_list(path: Path | str | None = None) -> pd.DataFrame:
    """대상 학교 리스트 로드(CNE_LIST.xlsx 등). 컬럼: 학교코드, 학교명(있으면)."""
    resolved = _resolve_target_list_path(path)
    if not resolved.exists():
        return pd.DataFrame(columns=["학교코드", "지역", "학교명"])
    xl = pd.ExcelFile(resolved, engine="openpyxl")
    # CNE_LIST.xlsx: 시트 'CNE', 헤더가 4행(0-based index 3): 학교코드, 지역, 학교명
    if "CNE" in xl.sheet_names:
        df = pd.read_excel(resolved, sheet_name="CNE", header=3, engine="openpyxl")
    else:
        try:
            df = pd.read_excel(resolved, sheet_name="시트 1", header=0, engine="openpyxl")
        except ValueError:
            df = pd.read_excel(resolved, sheet_name=0, header=0, engine="openpyxl")
    xl.close()
    # 학교코드 컬럼
    if "학교코드" not in df.columns and len(df.columns):
        df = df.rename(columns={df.columns[0]: "학교코드"})
    df = df.dropna(subset=["학교코드"])
    df["학교코드"] = df["학교코드"].astype(str).str.strip()
    if "학교명" not in df.columns:
        df["학교명"] = ""
    else:
        df["학교명"] = df["학교명"].fillna("").astype(str).str.strip()
    if "지역" not in df.columns:
        df["지역"] = ""
    else:
        df["지역"] = df["지역"].fillna("").astype(str).str.strip()
    return df[["학교코드", "지역", "학교명"]]


def get_va_path() -> Path:
    """가상자산 파일 경로 (수정본 우선, 없으면 원본)."""
    try:
        from .config_loader import get_path
        base = get_path("raw_data_root")
        p = base / "충남_통합_가상자산DB_Lee.xlsx"
        if p.exists():
            return p
    except Exception:
        pass
    # 수정본이 output에 있으면 사용
    proj = Path(__file__).resolve().parent.parent
    revised = proj / "output" / "충남_통합_가상자산DB_Lee_수정본.xlsx"
    if revised.exists():
        return revised
    return DEFAULT_VA_PATH


def get_target_school_codes(path: Path | str | None = None) -> set[str]:
    """대상 학교 리스트(CNE_LIST.xlsx)의 학교코드 집합 반환. 분석 시 이 집합으로만 필터링."""
    df = load_target_school_list(path)
    return set(df["학교코드"].tolist()) if not df.empty else set()


def filter_va_data_by_target(
    va_data: dict[str, pd.DataFrame],
    target_codes: set[str] | None = None,
) -> dict[str, pd.DataFrame]:
    """가상자산 데이터를 대상 학교 리스트(CNE_LIST.xlsx) 기준으로만 남김. target_codes 없으면 load_target_school_list()로 로드."""
    import re
    if target_codes is None:
        target_codes = get_target_school_codes()
    if not target_codes:
        return va_data
    out = {}
    # 학교정보: 학교코드 컬럼으로 필터
    if "학교정보" in va_data:
        df = va_data["학교정보"]
        if "학교코드" in df.columns:
            sc = df["학교코드"].astype(str).str.strip()
            out["학교정보"] = df[sc.isin(target_codes)].copy()
        else:
            out["학교정보"] = va_data["학교정보"].copy()
    # 장비 시트: 관리번호에서 학교코드 추출 후 필터 (학교정보에 있는 모든 형식 허용)
    for sheet in ("PoE", "AP", "스위치", "보안장비"):
        if sheet not in va_data or "관리번호" not in va_data[sheet].columns:
            continue
        df = va_data[sheet]
        def prefix(row):
            v = row.get("관리번호")
            if pd.isna(v) or "-" not in str(v):
                return None
            return str(v).strip().split("-", 1)[0].strip()
        prefixes = df.apply(prefix, axis=1)
        out[sheet] = df[prefixes.isin(target_codes)].copy()
    return out


def collect_school_codes_from_equipment(va_data: dict[str, pd.DataFrame]) -> set[str]:
    """장비 시트들에서 관리번호에 포함된 학교코드 추출 (학교정보 기준 형식)."""
    import re
    pattern = re.compile(r"^[A-Z0-9]{10,}$")  # 10자 이상 영숫자 (K, N, HT, BH 등 포함)
    codes = set()
    for sheet in ("PoE", "AP", "스위치", "보안장비"):
        if sheet not in va_data or "관리번호" not in va_data[sheet].columns:
            continue
        for val in va_data[sheet]["관리번호"].dropna():
            s = str(val).strip()
            if "-" in s:
                prefix = s.split("-", 1)[0].strip()
                if pattern.match(prefix):
                    codes.add(prefix)
    return codes


def run_verification(
    va_path: Path | str | None = None,
    output_dir: Path | str | None = None,
) -> dict:
    """
    검증 실행.
    Returns: {
        "va_path": path,
        "school_count": int,
        "target_count": int (대상 리스트 학교 수),
        "school_list": DataFrame (학교코드, 학교명),
        "in_equipment_not_in_info": list,  # 장비에만 있고 학교정보에 없는 학교코드
        "in_info_not_in_equipment": list,  # 학교정보에만 있고 장비에 없는 학교코드 (참고)
        "quality_issue_count": int,
        "quality_issue_by_type": dict,
    }
    """
    va_path = Path(va_path) if va_path else get_va_path()
    if not va_path.exists():
        return {"error": f"파일 없음: {va_path}"}

    va_data = load_va_data_sheets(va_path)
    if "학교정보" not in va_data:
        return {"error": "학교정보 시트 없음"}

    school_info = va_data["학교정보"]
    if "학교코드" not in school_info.columns or "학교명" not in school_info.columns:
        return {"error": "학교정보에 학교코드/학교명 컬럼 없음"}

    # 학교정보 기준 학교 수 (유니크 학교코드)
    school_list = school_info[["학교코드", "학교명"]].dropna(subset=["학교코드"])
    school_list = school_list.drop_duplicates(subset=["학교코드"])
    school_list["학교코드"] = school_list["학교코드"].astype(str).str.strip()
    school_count = len(school_list)
    info_codes = set(school_list["학교코드"].tolist())

    # 대상 학교 리스트(719) 기준 비교
    target_df = load_target_school_list()
    target_codes = set(target_df["학교코드"].tolist()) if not target_df.empty else set()
    target_count = len(target_codes)
    missing_from_data = sorted(target_codes - info_codes)  # 대상인데 가상자산 학교정보에 없음
    extra_in_data = sorted(info_codes - target_codes)  # 가상자산에 있지만 대상 리스트에 없음 (비대상)
    missing_df = target_df[target_df["학교코드"].isin(missing_from_data)] if not target_df.empty and missing_from_data else pd.DataFrame(columns=["학교코드", "지역", "학교명"])

    # 장비 시트에서 등장하는 학교코드 (9자리 형식만)
    equipment_codes = collect_school_codes_from_equipment(va_data)
    in_equipment_not_in_info = sorted(equipment_codes - info_codes)
    in_info_not_in_equipment = sorted(info_codes - equipment_codes)

    # 품질 검사
    issues = run_va_quality_checks(va_data)
    issue_df = pd.DataFrame(issues)
    quality_issue_count = len(issue_df)
    quality_issue_by_type = issue_df["issue_type"].value_counts().to_dict() if not issue_df.empty else {}

    result = {
        "va_path": str(va_path),
        "school_count": school_count,
        "target_count": target_count,
        "target_school_list_path": str(_resolve_target_list_path(None)),
        "school_list": school_list,
        "missing_from_data": missing_from_data,
        "missing_from_data_df": missing_df,
        "extra_in_data": extra_in_data,
        "in_equipment_not_in_info": in_equipment_not_in_info,
        "in_info_not_in_equipment": in_info_not_in_equipment,
        "quality_issue_count": quality_issue_count,
        "quality_issue_by_type": quality_issue_by_type,
    }

    if output_dir:
        out = Path(output_dir)
        out.mkdir(parents=True, exist_ok=True)
        school_list.to_csv(out / "학교목록_학교정보기준.csv", index=False, encoding="utf-8-sig")
        if missing_from_data:
            missing_df.to_csv(out / "누락_대상인데_데이터에_없는_학교.csv", index=False, encoding="utf-8-sig")
        if extra_in_data:
            sl_extra = school_list[school_list["학교코드"].isin(extra_in_data)]
            sl_extra.to_csv(out / "비대상_데이터에만_있는_학교.csv", index=False, encoding="utf-8-sig")
        if in_equipment_not_in_info:
            pd.DataFrame({"학교코드": in_equipment_not_in_info}).to_csv(
                out / "누락_장비에는_있는데_학교정보에_없음.csv", index=False, encoding="utf-8-sig"
            )
        # 검증 결과 요약 (검증용)
        with open(out / "검증결과_요약.txt", "w", encoding="utf-8") as f:
            f.write(f"대상 가상자산 파일: {va_path}\n")
            f.write(f"대상 학교 리스트: {_resolve_target_list_path(None)}\n")
            f.write(f"대상 학교 수(리스트): {target_count}개\n")
            f.write(f"가상자산 학교정보 학교 수: {school_count}개\n")
            f.write(f"대상인데 데이터에 없는 학교: {len(missing_from_data)}개 → 누락_대상인데_데이터에_없는_학교.csv\n")
            f.write(f"데이터에 있으나 비대상 학교: {len(extra_in_data)}개 → 비대상_데이터에만_있는_학교.csv\n")
            f.write(f"품질 검사 이상 건수: {quality_issue_count}\n")

    return result


def main() -> None:
    try:
        from .config_loader import get_path, ensure_runtime_dirs
        ensure_runtime_dirs()
        output_dir = get_path("output_root")
    except Exception:
        output_dir = Path(__file__).resolve().parent.parent / "output"

    va_path = get_va_path()
    print("대상 파일:", va_path)
    if not va_path.exists():
        print("파일이 없습니다.")
        return

    r = run_verification(va_path, output_dir=output_dir)
    if "error" in r:
        print("오류:", r["error"])
        return

    print()
    print("=== 대상 학교 리스트 기준 검증 ===")
    print("대상 학교 리스트:", r.get("target_school_list_path", "output/CNE_LIST.xlsx"))
    print(f"대상(리스트) 학교 수: {r['target_count']}개")
    print(f"가상자산 학교정보 학교 수: {r['school_count']}개")

    if r["missing_from_data"]:
        print()
        print("=== 대상인데 가상자산 데이터에 없는 학교 (학교코드, 학교명) ===")
        md = r["missing_from_data_df"]
        for _, row in md.head(30).iterrows():
            print(f"  {row['학교코드']}  {row.get('학교명', '')}")
        if len(r["missing_from_data"]) > 30:
            print(f"  ... 외 {len(r['missing_from_data']) - 30}건")
        if output_dir:
            print("  전체: output/누락_대상인데_데이터에_없는_학교.csv")
    else:
        print()
        print("→ 대상 리스트의 모든 학교가 가상자산 학교정보에 있습니다.")

    if r["extra_in_data"]:
        print()
        print("=== 가상자산에는 있으나 대상 리스트에 없는 학교 (비대상) ===")
        sl = r["school_list"].set_index("학교코드")["학교명"]
        for c in r["extra_in_data"][:20]:
            print(f"  {c}  {sl.get(c, '')}")
        if len(r["extra_in_data"]) > 20:
            print(f"  ... 외 {len(r['extra_in_data']) - 20}건")
        if output_dir:
            print("  전체: output/비대상_데이터에만_있는_학교.csv")

    if r["in_equipment_not_in_info"]:
        print()
        print("=== 장비 데이터에는 있는데 학교정보 시트에 없는 학교코드 ===")
        for c in r["in_equipment_not_in_info"][:15]:
            print(f"  {c}")
        if len(r["in_equipment_not_in_info"]) > 15:
            print(f"  ... 외 {len(r['in_equipment_not_in_info']) - 15}건")
        if output_dir:
            print("  전체: output/누락_장비에는_있는데_학교정보에_없음.csv")

    print()
    print("=== 품질 검사 (미수정 오류) ===")
    print(f"총 이상 건수: {r['quality_issue_count']}")
    if r["quality_issue_by_type"]:
        for k, v in r["quality_issue_by_type"].items():
            print(f"  - {k}: {v}건")
    if r["quality_issue_count"] > 0:
        print("→ 오류가 모두 수정된 것은 아닙니다. 자동 수정 가능한 8자리→9자리 등만 수정본에 반영됩니다.")
        print("  나머지는 수동 검토 후 수정이력에 따라 수정하세요.")


def list_missing_vs_reference(
    current_school_list_path: Path | str,
    reference_719_path: Path | str,
    output_path: Path | str | None = None,
) -> pd.DataFrame:
    """
    기준 목록(CSV/엑셀, 예: CNE_LIST.xlsx)과 현재 학교 목록을 비교해,
    기준에는 있는데 현재 파일에 없는 학교(학교코드, 학교명) 반환.
    reference_719_path: 학교코드 또는 학교명 컬럼이 있는 CSV/엑셀.
    """
    current = Path(current_school_list_path)
    ref = Path(reference_719_path)
    if not current.exists():
        raise FileNotFoundError(f"현재 목록 없음: {current}")
    if not ref.exists():
        raise FileNotFoundError(f"기준 719 목록 없음: {ref}")

    cur_df = pd.read_csv(current, encoding="utf-8-sig") if current.suffix.lower() == ".csv" else pd.read_excel(current)
    if ref.suffix.lower() == ".xlsx" or ref.suffix.lower() == ".xls":
        ref_df = pd.read_excel(ref)
    else:
        ref_df = pd.read_csv(ref, encoding="utf-8-sig")

    code_col = "학교코드" if "학교코드" in ref_df.columns else ref_df.columns[0]
    name_col = "학교명" if "학교명" in ref_df.columns else (ref_df.columns[1] if len(ref_df.columns) > 1 else None)
    ref_codes = set(ref_df[code_col].astype(str).str.strip().dropna().unique())
    cur_codes = set(cur_df["학교코드"].astype(str).str.strip().dropna().unique())
    missing_codes = ref_codes - cur_codes
    if not missing_codes:
        return pd.DataFrame(columns=["학교코드", "학교명"])

    ref_sub = ref_df[ref_df[code_col].astype(str).str.strip().isin(missing_codes)].copy()
    ref_sub = ref_sub.rename(columns={code_col: "학교코드", name_col: "학교명"}) if name_col else ref_sub.rename(columns={code_col: "학교코드"})
    if "학교명" not in ref_sub.columns and name_col:
        ref_sub["학교명"] = ref_sub.get(name_col, "")
    out = ref_sub[["학교코드", "학교명"]].drop_duplicates() if "학교명" in ref_sub.columns else ref_sub[["학교코드"]]
    if output_path:
        out.to_csv(output_path, index=False, encoding="utf-8-sig")
    return out


if __name__ == "__main__":
    main()
