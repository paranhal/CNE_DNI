# -*- coding: utf-8 -*-
"""전부하 최종데이터: W열='평균값 데이터' 행만 필터, 학교/장비별 600개 집계 (pandas)."""
import pandas as pd
from pathlib import Path
from collections import defaultdict

DATA_DIR = Path(r"G:\내 드라이브\CNE\최종데이터")
OUT_FILE = Path(__file__).resolve().parent.parent / "docs" / "전부하_최종데이터_분석결과.txt"


def analyze_one(path: Path, lines: list) -> int:
    df = pd.read_excel(path, header=0)
    w_col_name = df.columns[-1]
    for c in df.columns:
        if "평균값" in str(c):
            w_col_name = c
            break
    df_avg = df[df[w_col_name].astype(str).str.strip() == "평균값 데이터"].copy()
    total_avg = len(df_avg)
    lines.append(f"[파일] {path.name}")
    lines.append(f"  전체 행: {len(df)}, W열 '평균값 데이터' 행: {total_avg}")
    if total_avg == 0:
        lines.append("  (해당 행 없음)")
        lines.append("")
        return 0
    sc_col = "학교코드" if "학교코드" in df_avg.columns else df_avg.columns[21]
    m1_col = "메모1" if "메모1" in df_avg.columns else df_avg.columns[5]
    g = df_avg.groupby([sc_col, m1_col]).size()
    by_school = defaultdict(int)
    for (sc, m1), cnt in g.items():
        by_school[sc] += cnt
    lines.append(f"  학교 수: {len(by_school)}, 학교·장비 조합 수: {len(g)}")
    not_600 = [(k, v) for k, v in g.items() if v != 600]
    if not_600:
        lines.append(f"  600개 아님: {len(not_600)}건 (처음 5건: {not_600[:5]})")
    else:
        lines.append("  모든 학교·장비별 600개 OK")
    for school in sorted(by_school.keys())[:8]:
        n = by_school[school]
        lines.append(f"    {school}: 총 {n}행")
    if len(by_school) > 8:
        lines.append(f"    ... 외 {len(by_school) - 8}개 학교")
    lines.append("")
    return total_avg


def main():
    if not DATA_DIR.exists():
        OUT_FILE.write_text(f"폴더 없음: {DATA_DIR}", encoding="utf-8")
        return
    files = sorted(DATA_DIR.glob("전부하_원데이터_*.xlsx"))
    lines = [
        "=" * 60,
        "전부하 최종데이터 분석 (W열 = '평균값 데이터' 행만 사용)",
        f"경로: {DATA_DIR}",
        f"파일 수: {len(files)}",
        "=" * 60,
        "",
    ]
    total_avg_all = 0
    for path in files:
        try:
            total_avg_all += analyze_one(path, lines)
        except Exception as e:
            lines.append(f"[파일] {path.name} -> 오류: {e}\n")
    lines.append("=" * 60)
    lines.append(f"전체 '평균값 데이터' 행 합계: {total_avg_all}")
    lines.append("=" * 60)
    text = "\n".join(lines)
    OUT_FILE.parent.mkdir(parents=True, exist_ok=True)
    OUT_FILE.write_text(text, encoding="utf-8")
    print(text)
    print(f"\n저장: {OUT_FILE}")


if __name__ == "__main__":
    main()
