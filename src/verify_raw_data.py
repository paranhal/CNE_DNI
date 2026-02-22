"""
raw_data 아래 구성/자산(및 충남) 폴더를 스캔해 검증용 자료 목록을 출력합니다.

- 통합 파이프라인(integrate_export) 결과와 함께 검증할 때 사용.
- 지역별·장비별로 구분된 엑셀이 섞여 있을 수 있음.
사용: python -m src.verify_raw_data [--detail] [--compare]
"""

from __future__ import annotations

import json
from pathlib import Path


def get_raw_data_root() -> Path:
    proj = Path(__file__).resolve().parent.parent
    local_raw = proj / "raw_data"
    # 프로젝트 안 raw_data가 있으면 우선 사용(제안 위치로 옮긴 경우)
    if local_raw.exists():
        return local_raw
    try:
        from .config_loader import get_path
        return Path(get_path("raw_data_root"))
    except Exception:
        return local_raw


def scan_xlsx(dir_path: Path) -> list[Path]:
    if not dir_path.is_dir():
        return []
    return sorted(dir_path.rglob("*.xlsx"))


def get_sheet_info(xlsx_path: Path) -> list[tuple[str, int]]:
    """엑셀 파일의 시트별 이름과 데이터 행 수(대략) 반환."""
    try:
        import pandas as pd
        xl = pd.ExcelFile(xlsx_path)
        out = []
        for name in xl.sheet_names:
            try:
                df = pd.read_excel(xl, sheet_name=name, header=None)
                out.append((name, max(0, len(df) - 1)))  # 헤더 1행 가정
            except Exception:
                out.append((name, -1))
        return out
    except Exception:
        return []


def run(
    raw_data_root: Path | None = None,
    detail: bool = False,
    compare: bool = False,
    output_root: Path | None = None,
) -> dict:
    if raw_data_root is None:
        raw_data_root = get_raw_data_root()
    raw_data_root = Path(raw_data_root)
    if not raw_data_root.exists():
        return {
            "ok": False,
            "message": f"raw_data_root가 없습니다: {raw_data_root}. 구성/자산 폴더를 둔 뒤 다시 실행하세요.",
            "folders": {},
            "compare": None,
        }

    proj = Path(__file__).resolve().parent.parent
    if output_root is None:
        try:
            from .config_loader import get_path
            output_root = Path(get_path("output_root"))
        except Exception:
            output_root = proj / "output"

    folders = ("구성", "자산", "충남")
    result: dict = {"ok": True, "folders": {}, "compare": None}

    for folder in folders:
        d = raw_data_root / folder
        files = scan_xlsx(d)
        result["folders"][folder] = [str(p.relative_to(raw_data_root)) for p in files]
        if detail and files:
            result["folders"][f"{folder}_detail"] = {}
            for p in files:
                rel = str(p.relative_to(raw_data_root))
                result["folders"][f"{folder}_detail"][rel] = get_sheet_info(p)

    if compare and (output_root / "통합_export_요약.json").exists():
        with open(output_root / "통합_export_요약.json", encoding="utf-8") as f:
            summary = json.load(f)
        result["compare"] = {
            "통합_export_요약": summary,
            "raw_data_파일_수": {k: len(v) for k, v in result["folders"].items() if k in folders and isinstance(v, list)},
        }

    return result


def main() -> None:
    import sys
    detail = "--detail" in sys.argv
    compare = "--compare" in sys.argv
    r = run(detail=detail, compare=compare)
    if not r["ok"]:
        print(r["message"])
        return
    root = get_raw_data_root()
    print(f"raw_data_root: {root}\n")
    for folder, files in r["folders"].items():
        if folder.endswith("_detail"):
            continue
        if not isinstance(files, list):
            continue
        print(f"[{folder}] {len(files)}개 파일")
        for f in files[:30]:
            print(f"  {f}")
        if len(files) > 30:
            print(f"  ... 외 {len(files) - 30}개")
        if detail and f"{folder}_detail" in r["folders"]:
            for path, sheets in r["folders"][f"{folder}_detail"].items():
                if isinstance(sheets, list) and sheets:
                    print(f"  ※ {path}")
                    for name, rows in sheets:
                        print(f"      시트 {name!r}: 약 {rows}행")
        print()
    if r.get("compare"):
        c = r["compare"]
        print("--- 통합 export 요약 (비교 참고) ---")
        print(json.dumps(c.get("통합_export_요약", {}), ensure_ascii=False, indent=2)[:2000])
        if "raw_data_파일_수" in c:
            print("\nraw_data 폴더별 xlsx 파일 수:", c["raw_data_파일_수"])


if __name__ == "__main__":
    main()
