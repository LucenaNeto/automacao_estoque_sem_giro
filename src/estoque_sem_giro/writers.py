from __future__ import annotations

from pathlib import Path
import csv
import re, os
from .config import Config, yesterday_str
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter

def write_consolidated_csv(records: list[dict], cfg: Config) -> Path:
    if not records:
        raise ValueError("Nenhum registro válido para salvar.")
    records = sorted(records, key=lambda r: (r.get("PDV",""), r.get("SKU","")))
    cfg.output_dir.mkdir(parents=True, exist_ok=True)
    out = cfg.output_dir / f"{cfg.output_basename}_{yesterday_str(cfg)}.csv"
    with out.open("w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=cfg.final_fields)
        w.writeheader()
        for rec in records:
            w.writerow({k: rec.get(k, "") for k in cfg.final_fields})
    return out

def write_csvs_by_pdv(records: list[dict], cfg: Config) -> dict[str, Path]:
    if not records:
        return {}
    date_str = yesterday_str(cfg)
    folder = cfg.output_dir / f"por_pdv_{date_str}"
    folder.mkdir(parents=True, exist_ok=True)

    groups: dict[str, list[dict]] = {}
    for rec in records:
        pdv = (rec.get("PDV") or "").strip() or "SEM_PDV"
        pdv = re.sub(r"[^\w\-]+", "_", pdv)
        groups.setdefault(pdv, []).append(rec)

    paths: dict[str, Path] = {}
    for pdv, rows in sorted(groups.items(), key=lambda kv: kv[0]):
        rows_sorted = sorted(rows, key=lambda r: (r.get("SKU","")))
        path = folder / f"{cfg.output_basename}_{date_str}_PDV_{pdv}.csv"
        with path.open("w", newline="", encoding="utf-8-sig") as f:
            w = csv.DictWriter(f, fieldnames=cfg.final_fields)
            w.writeheader()
            for rec in rows_sorted:
                w.writerow({k: rec.get(k, "") for k in cfg.final_fields})
        paths[pdv] = path
    return paths


def write_reports_xlsx_by_pdv(records: list[dict], cfg: Config) -> dict[str, Path]:
    """
    Cria um arquivo .xlsx por PDV em:
      data/output/relatorio_finalizado_DD_MM_AAAA/relatorio_finalizado_DD_MM_AAAA_PDV_<pdv>.xlsx

    Abas:
      - cfg.report_sheet_main  (preenchida com os dados)
      - cfg.report_sheet_disc  (vazia por enquanto)

    Layout:
      - Linha 1: título "Grupo Ana Sobral" (mesclado A1:ÚltimaColuna1)
      - Linha 2: espaço para LOGO (altura maior)
      - Linha 3: vazia (respiro)
      - Linha 4: cabeçalho da tabela
      - A tabela começa na linha 5
      - Coluna "CURVA": A/B=verde, C=amarelo, D/E=vermelho
    """
    if not records:
        return {}
    GROUP_NAME = "Grupo Ana Sobral"
    date_str = yesterday_str(cfg)
    folder = cfg.output_dir / f"{cfg.report_folder_prefix}_{date_str}"
    folder.mkdir(parents=True, exist_ok=True)

    # Agrupar por PDV
    groups: dict[str, list[dict]] = {}
    for rec in records:
        pdv_raw = (rec.get("PDV") or "").strip() or "SEM_PDV"
        pdv = re.sub(r"[^\w\-]+", "_", pdv_raw)
        groups.setdefault(pdv, []).append(rec)

    header = list(cfg.final_fields)
    ncols = len(header)
    last_col_letter = get_column_letter(ncols)

    # larguras sugeridas
    widths = {
        "PDV": 12, "SKU": 14, "DESCRIÇÃO": 50, "MARCA": 16, "CURVA": 10, "ESTOQUE_ATUAL": 18
    }

    # Fills para CURVA
    FILL_GREEN  = PatternFill(fill_type="solid", start_color="C6EFCE", end_color="C6EFCE")  # verde claro
    FILL_YELLOW = PatternFill(fill_type="solid", start_color="FFEB9C", end_color="FFEB9C")  # amarelo claro
    FILL_RED    = PatternFill(fill_type="solid", start_color="FFC7CE", end_color="FFC7CE")  # vermelho claro

    out_paths: dict[str, Path] = {}

    for pdv, rows in sorted(groups.items(), key=lambda kv: kv[0]):
        rows_sorted = sorted(rows, key=lambda r: (r.get("SKU", "")))

        # === Workbook e folha principal ===
        wb = Workbook()
        ws = wb.active
        ws.title = cfg.report_sheet_main

        # --- Cabeçalho visual (linhas 1-3) ---
        # Título
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncols)
        cell_title = ws.cell(row=1, column=1, value=GROUP_NAME)
        cell_title.font = Font(bold=True, size=16)
        cell_title.alignment = Alignment(horizontal="center", vertical="center")

        # Espaço para LOGO (só reservando altura por enquanto)
        ws.row_dimensions[2].height = 40  # facilita inserir logo futuramente
        # Linha 3 deixamos vazia (respiro)

        # --- Cabeçalho da tabela na linha 4 ---
        header_row = 4
        for col_idx, col_name in enumerate(header, start=1):
            c = ws.cell(row=header_row, column=col_idx, value=col_name)
            c.font = Font(bold=True)
            c.alignment = Alignment(horizontal="center", vertical="center")
            ws.column_dimensions[get_column_letter(col_idx)].width = widths.get(col_name, 16)

        ws.freeze_panes = f"A{header_row+1}"  # congela até o cabeçalho
        ws.auto_filter.ref = f"A{header_row}:{last_col_letter}{header_row}"

        # --- Dados: começam na linha 5 ---
        first_data_row = header_row + 1
        for r_idx, rec in enumerate(rows_sorted, start=first_data_row):
            for c_idx, col_name in enumerate(header, start=1):
                ws.cell(row=r_idx, column=c_idx, value=rec.get(col_name, ""))

        # --- Coloração condicional (CURVA) ---
        if "CURVA" in header:
            curva_col_idx = header.index("CURVA") + 1
            for r in range(first_data_row, ws.max_row + 1):
                val = ws.cell(row=r, column=curva_col_idx).value
                if val is None:
                    continue
                v = str(val).strip().upper()
                if v in {"A", "B"}:
                    ws.cell(row=r, column=curva_col_idx).fill = FILL_GREEN
                elif v == "C":
                    ws.cell(row=r, column=curva_col_idx).fill = FILL_YELLOW
                elif v in {"D", "E"}:
                    ws.cell(row=r, column=curva_col_idx).fill = FILL_RED

        # === Segunda aba: Descontinuados (com o mesmo topo visual e cabeçalho vazio por enquanto) ===
        ws2 = wb.create_sheet(cfg.report_sheet_disc)

        # topo visual
        ws2.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncols)
        cell_title2 = ws2.cell(row=1, column=1, value=GROUP_NAME)
        cell_title2.font = Font(bold=True, size=16)
        cell_title2.alignment = Alignment(horizontal="center", vertical="center")
        ws2.row_dimensions[2].height = 40

        # cabeçalho da tabela
        for col_idx, col_name in enumerate(header, start=1):
            c = ws2.cell(row=header_row, column=col_idx, value=col_name)
            c.font = Font(bold=True)
            c.alignment = Alignment(horizontal="center", vertical="center")
            ws2.column_dimensions[get_column_letter(col_idx)].width = widths.get(col_name, 16)

        ws2.freeze_panes = f"A{header_row+1}"
        ws2.auto_filter.ref = f"A{header_row}:{last_col_letter}{header_row}"

        # --- Salvar de forma atômica ---
        path = folder / f"{cfg.report_folder_prefix}_{date_str}_PDV_{pdv}.xlsx"
        tmp  = path.with_suffix(path.suffix + ".tmp")
        wb.save(tmp)
        os.replace(tmp, path)
        out_paths[pdv] = path

    return out_paths
