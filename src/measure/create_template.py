# -*- coding: utf-8 -*-
"""최소 템플릿 생성 (사용자가 자체 서식으로 교체 가능)"""
import os
from openpyxl import Workbook

BASE = os.path.dirname(os.path.abspath(__file__))
for folder in ["측정밗_템플릿", "측정값_템플릿"]:
    path = os.path.join(BASE, folder)
    os.makedirs(path, exist_ok=True)
    out = os.path.join(path, "템플릿.xlsx")
    wb = Workbook()
    ws = wb.active
    ws.title = "측정"
    for r in [20, 22, 23, 27, 28, 29, 30, 31, 32, 33]:
        ws.cell(r, 1, value=f"행{r}")
    wb.save(out)
    print(f"생성: {out}")
