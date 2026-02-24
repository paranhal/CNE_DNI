"""
719중_장비별_데이터없는_학교.csv를 읽어, 검증용 '데이터 없는 학교' 리스트를 생성.

- 데이터_전혀_없는_학교.csv: 장비 데이터가 하나도 없는 학교 (학교코드, 지역, 학교명)
- 데이터_없는_학교_검증리스트.csv: 전혀 없음 + 일부 없음 통합 (학교코드, 지역, 학교명, 유형, 없는_장비)
"""

from __future__ import annotations

from pathlib import Path


def run(output_dir: Path | None = None) -> dict:
    import pandas as pd

    if output_dir is None:
        output_dir = Path(__file__).resolve().parent.parent / "output"
    output_dir = Path(output_dir)
    path = output_dir / "719중_장비별_데이터없는_학교.csv"
    if not path.exists():
        return {"error": f"파일 없음: {path}. 먼저 python -m src.integrate_export 를 실행하세요."}

    # 대상 학교 리스트(CNE_LIST.xlsx)에서 학교명·지역 로드하여 채움
    try:
        from .verify_schools import load_target_school_list
        target_df = load_target_school_list()
        target_df = target_df.drop_duplicates(subset=["학교코드"])
        target_df["학교코드"] = target_df["학교코드"].astype(str).str.strip()
        name_map = target_df.set_index("학교코드")["학교명"].to_dict()
        region_map = target_df.set_index("학교코드")["지역"].to_dict() if "지역" in target_df.columns else {}
    except Exception:
        name_map = {}
        region_map = {}

    df = pd.read_csv(path, encoding="utf-8-sig")
    df = df[df["학교코드"].astype(str).str.strip() != "학교코드"].copy()
    df = df[df["학교코드"].notna() & (df["학교코드"].astype(str).str.strip() != "")].copy()
    df["학교코드"] = df["학교코드"].astype(str).str.strip()
    # 학교명·지역: CNE_LIST.xlsx 기준으로 채움
    if name_map:
        df["학교명"] = df["학교코드"].map(lambda c: name_map.get(c, "") or "").astype(str).str.strip()
    if "학교명" not in df.columns:
        df["학교명"] = ""
    if region_map:
        df["지역"] = df["학교코드"].map(lambda c: region_map.get(c, "") or "").astype(str).str.strip()
    if "지역" not in df.columns:
        df["지역"] = ""

    eq_cols = [c for c in df.columns if c not in ("학교코드", "지역", "학교명")]
    df["_sum"] = df[eq_cols].astype(int).sum(axis=1)

    def missing_list(row):
        return ", ".join([c for c in eq_cols if row.get(c, 0) == 0 or pd.isna(row.get(c))])

    # 전혀 없는 학교 (학교코드, 지역, 학교명)
    none_df = df[df["_sum"] == 0][["학교코드", "지역", "학교명"]].copy()
    none_path = output_dir / "데이터_전혀_없는_학교.csv"
    none_df.to_csv(none_path, index=False, encoding="utf-8-sig")

    # 안내: 이 목록의 학교는 통합 파일에 없지만 학교별 폴더에 자료가 있을 수 있음
    note_path = output_dir / "데이터_전혀_없는_학교_안내.txt"
    note_path.write_text(
        "이 CSV의 학교는 '통합 VA/CFG' 파일에는 데이터가 없습니다.\n"
        "일부 학교는 클라우드/학교별 폴더에만 자료가 있을 수 있습니다.\n"
        "경로 패턴: .../충남/{지역번호}_{지역명}/{학교폴더명}/\n"
        "  → {학교코드}_{학교명}_{AP|PoE|스위치|보안장비}_일괄업로드용_가상자산DB.xlsx / 구성정보DB.xlsx\n"
        "자세한 설명: 프로젝트 docs/EXPORT_BASE.md '데이터 전혀 없는 학교 → 학교별 폴더' 섹션 참고.\n",
        encoding="utf-8",
    )

    # 검증 리스트 전체 (전혀 없음 + 일부 없음, 학교명 채움)
    all_missing = df[df["_sum"] < len(eq_cols)].copy()
    all_missing["없는_장비"] = all_missing.apply(missing_list, axis=1)
    all_missing["유형"] = all_missing["_sum"].apply(lambda s: "전혀 없음" if s == 0 else "일부 없음")
    verify = all_missing[["학교코드", "지역", "학교명", "유형", "없는_장비"]].sort_values(["유형", "지역", "학교코드"])
    verify_path = output_dir / "데이터_없는_학교_검증리스트.csv"
    verify.to_csv(verify_path, index=False, encoding="utf-8-sig")

    return {
        "전혀_없는_학교_수": len(none_df),
        "검증리스트_전체_수": len(verify),
        "전혀_없는_학교_파일": str(none_path),
        "전혀_없는_학교_안내": str(note_path),
        "검증리스트_파일": str(verify_path),
    }


if __name__ == "__main__":
    import sys
    try:
        from .config_loader import get_path
        output_dir = get_path("output_root")
    except Exception:
        output_dir = Path(__file__).resolve().parent.parent / "output"
    if len(sys.argv) > 1:
        output_dir = Path(sys.argv[1])
    r = run(output_dir=output_dir)
    if "error" in r:
        print(r["error"])
        raise SystemExit(1)
    print("데이터 전혀 없는 학교:", r["전혀_없는_학교_수"], "개 →", r["전혀_없는_학교_파일"])
    print("검증 리스트 전체:", r["검증리스트_전체_수"], "개 →", r["검증리스트_파일"])
