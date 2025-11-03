from __future__ import annotations

from pathlib import Path
import csv
import re
from .config import Config, yesterday_str

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
