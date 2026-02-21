"""
데이터 품질 검사: 관리코드·학교코드 검증 및 이상 케이스 리포트.

- 관리번호가 관리코드 형식(학교코드-장비명+일련번호)인지 검사
- 관리번호에 학교코드만 있는 경우 탐지
- 학교코드 형식 오류(ES/MS/HS/SS 누락·자릿수 등) 탐지
- 학교정보 기준 코드 목록과 대조
"""

from __future__ import annotations

import re
from pathlib import Path
from typing import Any

import pandas as pd

from .load_excel import load_va_data_sheets
from .sheet_defs import VA_DATA_SHEETS_EQUIPMENT

# 관리코드: 학교코드(12자)-장비명+일련번호  예: N108140063HS-SWL20002
# 학교코드: N + 9자리 숫자 + ES|MS|HS|SS  (12자). 8자리는 오류.
SCHOOL_CODE_PATTERN = re.compile(r"^N\d{9}(ES|MS|HS|SS)$")
MANAGEMENT_CODE_PATTERN = re.compile(
    r"^N\d{9}(ES|MS|HS|SS)-[A-Za-z0-9]+\d+$"
)


def is_valid_school_code(value: Any) -> bool:
    """학교코드 형식(12자, N+9자리+ES|MS|HS|SS) 여부."""
    if pd.isna(value) or not isinstance(value, str):
        return False
    s = str(value).strip()
    return bool(SCHOOL_CODE_PATTERN.match(s))


def is_valid_management_code(value: Any) -> bool:
    """관리코드 형식(학교코드-장비명+일련번호) 여부."""
    if pd.isna(value) or not isinstance(value, str):
        return False
    s = str(value).strip()
    return bool(MANAGEMENT_CODE_PATTERN.match(s))


def extract_school_code_from_management(code: str) -> str | None:
    """관리코드에서 학교코드 부분만 추출. 형식이 아니면 None."""
    if not code or "-" not in code:
        return None
    part = code.split("-", 1)[0].strip()
    return part if is_valid_school_code(part) else None


def classify_management_value(value: Any) -> str:
    """관리번호 컬럼 값 분류: 'valid' | 'school_code_only' | 'invalid_format' | 'empty'."""
    if pd.isna(value) or (isinstance(value, str) and not value.strip()):
        return "empty"
    s = str(value).strip()
    if is_valid_management_code(s):
        return "valid"
    if is_valid_school_code(s):
        return "school_code_only"
    return "invalid_format"


def check_va_sheet(
    df: pd.DataFrame,
    sheet_name: str,
    valid_school_codes: set[str] | None = None,
) -> list[dict]:
    """가상자산 장비 시트 한 개 검사. 학교정보에 있는 학교코드는 유효로 간주(K, HT, BH 등 포함)."""
    issues = []
    if "관리번호" not in df.columns:
        return [{"sheet": sheet_name, "issue": "no_column_관리번호", "row": None}]

    for idx, row in df.iterrows():
        val = row.get("관리번호")
        excel_row = idx + 2  # 1-based + 헤더 1행
        if pd.isna(val) or (isinstance(val, str) and not val.strip()):
            continue
        s = str(val).strip()
        # 학교정보 기준: prefix가 학교정보에 있으면 유효 (K, HT, BH 등 포함)
        if valid_school_codes and "-" in s:
            prefix = s.split("-", 1)[0].strip()
            if prefix in valid_school_codes:
                continue  # 유효, 이슈 없음 (719 대상 여부는 분석 시 필터링)
        kind = classify_management_value(val)
        if kind == "school_code_only":
            issues.append({
                "sheet": sheet_name,
                "excel_row": excel_row,
                "column": "관리번호",
                "value": val,
                "issue_type": "관리번호에_학교코드만_있음",
            })
            continue
        if kind == "invalid_format":
            # prefix가 학교정보에 없으면 형식 오류 (또는 학교정보에 없는 코드)
            issues.append({
                "sheet": sheet_name,
                "excel_row": excel_row,
                "column": "관리번호",
                "value": val,
                "issue_type": "관리코드_형식_오류",
            })
            continue
        # kind == "valid" (정규식 통과) → 학교정보에 없을 때만 이슈 (정규식만 맞고 학교정보에 없는 경우)
        if valid_school_codes:
            school = extract_school_code_from_management(s)
            if school and school not in valid_school_codes:
                issues.append({
                    "sheet": sheet_name,
                    "excel_row": excel_row,
                    "column": "관리번호",
                    "value": val,
                    "issue_type": "관리코드_형식_오류",
                })
    return issues


def check_school_info_sheet(df: pd.DataFrame, sheet_name: str = "학교정보") -> list[dict]:
    """학교정보 시트의 학교코드 검사. 기준은 학교정보 자체이므로 길이·형식만 완화 검사."""
    issues = []
    if "학교코드" not in df.columns:
        return [{"sheet": sheet_name, "issue": "no_column_학교코드", "row": None}]

    for idx, row in df.iterrows():
        val = row.get("학교코드")
        excel_row = idx + 4  # 헤더 3행 + 1-based
        if pd.isna(val) or (isinstance(val, str) and not val.strip()):
            continue
        s = str(val).strip()
        # 학교정보 기준이므로 N+9자리+ES/MS/HS만이 아닌, 10자 이상·영숫자 형태면 통과
        if len(s) < 10 or not any(c.isalnum() for c in s):
            issues.append({
                "sheet": sheet_name,
                "excel_row": excel_row,
                "column": "학교코드",
                "value": val,
                "issue_type": "학교코드_형식_오류",
            })
    return issues


def run_va_quality_checks(va_data: dict[str, pd.DataFrame]) -> list[dict]:
    """가상자산 로딩 결과 전체에 대해 품질 검사. 학교정보에 있는 코드는 유효(K, HT, BH 등)."""
    all_issues = []
    valid_school_codes = None
    if "학교정보" in va_data:
        sc = va_data["학교정보"]["학교코드"].dropna().astype(str).str.strip()
        valid_school_codes = set(sc[sc.str.len() >= 10])  # 학교정보 전체 코드 기준 (K, HT, BH 등 포함)

    for sheet_name in VA_DATA_SHEETS_EQUIPMENT:
        if sheet_name not in va_data:
            continue
        all_issues.extend(
            check_va_sheet(va_data[sheet_name], sheet_name, valid_school_codes)
        )
    if "학교정보" in va_data:
        all_issues.extend(check_school_info_sheet(va_data["학교정보"]))
    return all_issues


def run_quality_report(
    va_path: Path | str,
    output_dir: Path | str | None = None,
) -> pd.DataFrame:
    """가상자산 파일 로딩 후 품질 검사 실행, 이슈를 DataFrame으로 반환. 저장 optional."""
    va_path = Path(va_path)
    va_data = load_va_data_sheets(va_path)
    issues = run_va_quality_checks(va_data)
    df = pd.DataFrame(issues)
    if output_dir is not None and len(df) > 0:
        out = Path(output_dir)
        out.mkdir(parents=True, exist_ok=True)
        out_file = out / "품질검사_이상목록.csv"
        df.to_csv(out_file, index=False, encoding="utf-8-sig")
    return df


# 통합 가상자산 기본 경로 (config에 없을 때 사용)
DEFAULT_VA_PATH = Path(
    "/Users/paranhal/Library/CloudStorage/GoogleDrive-paranhanl66@gmail.com"
    "/내 드라이브/20260215_D드라이브/Lee_20260202/충남_통합_가상자산DB_Lee.xlsx"
)


def main() -> None:
    """통합 가상자산 경로를 config 또는 기본값으로 읽어 품질 검사 실행."""
    try:
        from .config_loader import get_path, ensure_runtime_dirs
        ensure_runtime_dirs()
        base = get_path("raw_data_root")
        va_path = base / "충남_통합_가상자산DB_Lee.xlsx"
        output_dir = get_path("output_root")
    except Exception:
        va_path = Path(__file__).resolve().parent.parent / "output" / "충남_통합_가상자산DB_Lee.xlsx"
        output_dir = Path(__file__).resolve().parent.parent / "output"
    if not va_path.exists():
        va_path = DEFAULT_VA_PATH
    if not va_path.exists():
        print("가상자산 파일을 찾을 수 없습니다:", va_path)
        return
    print("품질 검사 실행:", va_path)
    df = run_quality_report(va_path, output_dir=output_dir)
    print(f"총 이상 건수: {len(df)}")
    if len(df) > 0:
        print(df["issue_type"].value_counts())
        if output_dir:
            print("저장:", Path(output_dir) / "품질검사_이상목록.csv")
    else:
        print("이상 없음.")


if __name__ == "__main__":
    main()
