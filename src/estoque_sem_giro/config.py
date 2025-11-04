from __future__ import annotations
from dataclasses import dataclass
from pathlib import Path
from zoneinfo import ZoneInfo
from datetime import datetime, timedelta
import sys

@dataclass(frozen=True)
class Config:
    project_root: Path
    data_dir: Path
    input_dir: Path
    output_dir: Path
    archive_dir: Path
    timezone: ZoneInfo
    expected_sheets: tuple[str, ...]
    final_fields: tuple[str, ...]
    output_basename: str
    generate_by_pdv: bool
    preview_max_rows: int
    enable_fallback: bool
    report_folder_prefix: str
    report_sheet_main: str
    report_sheet_disc: str
    discontinued_folder_prefix: str
    discontinued_fields: tuple[str, ...]
    # LOGO
    logo_path: Path | None
    logo_max_width_px: int
    logo_row_height: int

def load_config() -> Config:
    # Se estiver “frozen” (PyInstaller), usa a pasta do executável como raiz
    if getattr(sys, "frozen", False):
        base = Path(sys.executable).parent
    else:
        base = Path(__file__).resolve().parents[2]

    data_dir = (base / "data")
    logo_file = data_dir / "assets" / "logo.png"

    return Config(
        project_root=base,
        data_dir=data_dir,
        input_dir=data_dir / "input",
        output_dir=data_dir / "output",
        archive_dir=data_dir / "archived",
        timezone=ZoneInfo("America/Maceio"),  # Maceió (ou Recife, se preferir)
        expected_sheets=("EUD", "BOT", "QDB"),
        final_fields=("PDV", "SKU", "DESCRIÇÃO", "MARCA", "CURVA", "ESTOQUE_ATUAL"),
        output_basename="Estoque_sem_giro",
        generate_by_pdv=True,
        preview_max_rows=3,
        enable_fallback=True,
        report_folder_prefix="relatorio_finalizado",
        report_sheet_main="Estoque sem Giro",
        report_sheet_disc="Descontinuados",
        discontinued_folder_prefix="descontinuados",
        discontinued_fields=("PDV","SKU","SKU_PARA","DESCRIÇÃO","ESTOQUE ATUAL","FASES DO PRODUTO","MARCA"),
        logo_path=logo_file if logo_file.exists() else None,
        logo_max_width_px=380,
        logo_row_height=60,
    )

def ensure_dirs(cfg: Config) -> None:
    cfg.input_dir.mkdir(parents=True, exist_ok=True)
    cfg.output_dir.mkdir(parents=True, exist_ok=True)
    cfg.archive_dir.mkdir(parents=True, exist_ok=True)

def yesterday_str(cfg: Config) -> str:
    return (datetime.now(cfg.timezone) - timedelta(days=1)).strftime("%d_%m_%Y")
