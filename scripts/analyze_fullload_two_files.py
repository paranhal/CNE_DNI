# -*- coding: utf-8 -*-
"""전부하 최종데이터: 계룡 1차·2차 2개 파일만 분석 (W열 평균값 데이터)."""
import pandas as pd
from pathlib import Path
from collections import defaultdict

DATA_DIR = Path(r"G:\내 드라이브\CNE\최종데이터")
# 상대적으로 작은 파일 2개
FILES = [
    "전부하_원데이터_1차_20260315_133425_계룡.xlsx",
    "전부하_원데이터_2차_20260315_143129_계룡.xlsx",
]
OUT_FILE = Path(__file__).resolve().parent.parent / "docs" / "전부하_최종데이터_분석결과.txt"


def main():
    lines = [
        "전부하 최종데이터 분석 (W열 = '평균값 데이터' 행만 사용)",
        f"경로: {DATA_DIR}",
        "분석 대상: 계룡 1차·2차 2개 파일",
        "=" * 60,
        "",
    ]
    total_avg = 0
    for fname in FILES:
        path = DATA_DIR / fname
        if not path.exists():
            lines.append(f"[파일 없음] {fname}\n")
            continue
        df = pd.read_excel(path, header=0)
        w_col = df.columns[-1]
        df_avg = df[df[w_col].astype(str).str.strip() == "평균값 데이터"].copy()
        n = len(df_avg)
        total_avg += n
        lines.append(f"[파일] {fname}")
        lines.append(f"  전체 행: {len(df)}, '평균값 데이터' 행: {n}")
        if n == 0:
            lines.append("")
            continue
        g = df_avg.groupby(["학교코드", "메모1"]).size()
        not600 = g[g != 600]
        lines.append(f"  학교·장비 조합 수: {len(g)}, 600개 아님: {len(not600)}건")
        for (sc, m1), cnt in g.head(12).items():
            lines.append(f"    {sc} | {m1} -> {cnt}")
        lines.append("")
    lines.append("=" * 60)
    lines.append(f"합계 '평균값 데이터' 행: {total_avg}")
    text = "\n".join(lines)
    OUT_FILE.parent.mkdir(parents=True, exist_ok=True)
    OUT_FILE.write_text(text, encoding="utf-8")
    print(text)


if __name__ == "__main__":
    main()
