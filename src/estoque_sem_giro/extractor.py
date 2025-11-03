from __future__ import annotations
from typing import Iterable
from .excel_reader import _clean_str, find_header_row

_COL_POS = {"A": 1, "C": 3, "E": 5, "I": 9, "J": 10}  # 1-based

def extract_sheet(ws, marca: str) -> list[dict]:
    header_row = find_header_row(ws, cols=("A","C","E","I","J"))
    start_row  = (header_row + 1) if header_row else 1
    max_col    = max(_COL_POS.values())

    out: list[dict] = []
    for row in ws.iter_rows(min_row=start_row, min_col=1, max_col=max_col, values_only=True):
        def at(i: int) -> str:
            return _clean_str(row[i-1] if len(row) >= i else "")
        sku, desc, curva, pdv, est = at(1), at(3), at(5), at(9), at(10)
        if not any((sku, desc, curva, pdv, est)):
            continue
        if not sku or not pdv:
            continue
        out.append({
            "PDV":           pdv,
            "SKU":           sku,
            "DESCRIÇÃO":     desc,
            "MARCA":         marca,
            "CURVA":         curva,
            "ESTOQUE_ATUAL": est,
        })
    return out

def extract_all(wb, expected_sheets: Iterable[str]) -> list[dict]:
    data: list[dict] = []
    for sheet in expected_sheets:
        if sheet in wb.sheetnames:
            ws = wb[sheet]
            recs = extract_sheet(ws, marca=sheet)
            data.extend(recs)
    return data
