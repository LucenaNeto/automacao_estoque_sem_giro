from __future__ import annotations

from pathlib import Path
import logging
# Adicionado 'yesterday_str' para a nova linha de log
from .config import Config, ensure_dirs, yesterday_str
from .logging_config import setup_logging
from .excel_reader import open_workbook, preview_sheet
from .extractor import extract_all, extract_discontinued_all
from .writers import (
    write_consolidated_csv,
    write_csvs_by_pdv,
    write_reports_xlsx_by_pdv,
    write_discontinued_csvs_by_pdv,
)
from .archiver import archive_xlsx


def is_excel(p: Path) -> bool:
    return p.is_file() and p.suffix.lower() == ".xlsx" and not p.name.startswith("~$")


def list_excels(input_dir: Path) -> list[Path]:
    if not input_dir.exists():
        return []
    xs = [p for p in input_dir.iterdir() if is_excel(p)]
    xs.sort(key=lambda x: (x.stat().st_mtime, x.name))
    return xs


def latest_excel(input_dir: Path) -> Path | None:
    xs = list_excels(input_dir)
    return xs[-1] if xs else None


def process_latest(cfg: Config) -> bool:
    setup_logging()
    ensure_dirs(cfg)
    
    # 1. Lista inicializada
    discontinued: list[dict] = []

    xlsx = latest_excel(cfg.input_dir)
    if not xlsx:
        logging.error("Nenhum .xlsx encontrado em %s", cfg.input_dir)
        return False

    logging.info("Arquivo selecionado (mais recente): %s", xlsx.name)

    # 1) data_only=True
    try:
        wb = open_workbook(xlsx, data_only=True)
    except Exception as e:
        logging.exception("Falha ao abrir workbook: %s", e)
        return False

    # prévias (opcional)
    for sheet in cfg.expected_sheets:
        if sheet in wb.sheetnames:
            try:
                samples = preview_sheet(
                    wb[sheet], cols=("A", "C", "E", "I", "J"), max_rows=cfg.preview_max_rows
                )
                if samples:
                    header = " | ".join(("A", "C", "E", "I", "J"))
                    logging.info("Prévia %s: %s", sheet, header)
                    for row in samples:
                        logging.info(" | ".join(row))
            except Exception:
                pass

    records = extract_all(wb, cfg.expected_sheets)
    # 2. extrai descontinuados (já estava presente)
    discontinued = extract_discontinued_all(wb, cfg.expected_sheets)
    try:
        wb.close()
    except Exception:
        pass

    # 2) fallback sem cache de fórmulas
    if not records and cfg.enable_fallback:
        logging.warning("Sem registros; tentando fallback data_only=False.")
        try:
            wb2 = open_workbook(xlsx, data_only=False)
            records = extract_all(wb2, cfg.expected_sheets)
            # 3. extrai descontinuados no fallback (já estava presente)
            discontinued = extract_discontinued_all(wb2, cfg.expected_sheets)
            try:
                wb2.close()
            except Exception:
                pass
        except Exception as e:
            logging.exception("Falha no fallback: %s", e)

    if not records:
        logging.error(
            "Ainda sem registros. Reabra no Excel, Ctrl+Alt+F9 (recalcular) e salve."
        )
        return False

    out = write_consolidated_csv(records, cfg)
    logging.info("[OK] Consolidado: %s", out)

    if cfg.generate_by_pdv:
        paths = write_csvs_by_pdv(records, cfg)
        logging.info("[OK] %d CSVs por PDV gerados.", len(paths))

    # 4. Bloco de escrita dos relatórios ATUALIZADO
    # Relatórios Excel por PDV (preenche aba principal e a de descontinuados)
    reports = write_reports_xlsx_by_pdv(records, discontinued, cfg)
    logging.info("[OK] %d relatórios Excel por PDV em %s", len(reports), (cfg.output_dir / f"{cfg.report_folder_prefix}_{yesterday_str(cfg)}"))

    # CSVs de DESCONTINUADOS por PDV (bloco existente)
    if discontinued:
        disc_paths = write_discontinued_csvs_by_pdv(discontinued, cfg)
        logging.info(
            "[OK] %d CSVs de descontinuados por PDV gerados em %s",
            len(disc_paths),
            (cfg.output_dir / f"{cfg.discontinued_folder_prefix}_{yesterday_str(cfg)}"),
        )
    else:
        logging.info("Nenhum registro de descontinuados encontrado nas abas esperadas.")

    try:
        archived = archive_xlsx(xlsx, cfg.archive_dir)
        logging.info("[OK] Arquivado: %s", archived)
    except Exception as e:
        logging.warning("Não foi possível arquivar %s: %s", xlsx.name, e)

    return True