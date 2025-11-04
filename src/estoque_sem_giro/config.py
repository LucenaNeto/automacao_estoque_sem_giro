from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from zoneinfo import ZoneInfo
from datetime import datetime, timedelta

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
    #adicionando novas configurações para relatórios futuros
    report_folder_prefix: str
    report_sheet_main: str
    report_sheet_disc: str
    # para descontinuados 
    discontinued_folder_prefix: str
    discontinued_fields: tuple[str, ...]   

def load_config() -> Config:
    # projeto com layout "src/"; o root é 2 níveis acima deste arquivo
    project_root = Path(__file__).resolve().parents[2]
    data_dir = project_root / "data"
    return Config(
        project_root=project_root,
        data_dir=data_dir,
        input_dir=data_dir / "input",
        output_dir=data_dir / "output",
        archive_dir=data_dir / "archived",
        timezone=ZoneInfo("America/Recife"),
        expected_sheets=("EUD", "BOT", "QDB"),
        final_fields=("PDV", "SKU", "DESCRIÇÃO", "MARCA", "CURVA", "ESTOQUE_ATUAL"),
        output_basename="Estoque_sem_giro",
        generate_by_pdv=True,
        preview_max_rows=3,
        enable_fallback=True,
        #adicionando novas configurações para relatórios futuros
        report_folder_prefix="relatorios_finalizado",
        report_sheet_main="Estoque sem Giro",
        report_sheet_disc="Descontinuados",
        # para descontinuados
        discontinued_folder_prefix="descontinuados",
        discontinued_fields=("PDV", "SKU","SKU_PARA" ,"DESCRIÇÃO", "ESTOQUE ATUAL","FASES DO PRODUTO","MARCA"),   
    )

def ensure_dirs(cfg: Config) -> None:
    cfg.input_dir.mkdir(parents=True, exist_ok=True)
    cfg.output_dir.mkdir(parents=True, exist_ok=True)
    cfg.archive_dir.mkdir(parents=True, exist_ok=True)

def yesterday_str(cfg: Config) -> str:
    return (datetime.now(cfg.timezone) - timedelta(days=1)).strftime("%d_%m_%Y")
