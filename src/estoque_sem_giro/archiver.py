from __future__ import annotations
from pathlib import Path
import shutil
from datetime import datetime

def archive_xlsx(xlsx: Path, archive_dir: Path) -> Path:
    archive_dir.mkdir(parents=True, exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    dest = archive_dir / f"{xlsx.stem}__processado_{ts}{xlsx.suffix}"
    shutil.move(str(xlsx), str(dest))
    return dest
