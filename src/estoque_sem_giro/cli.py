from __future__ import annotations
import argparse
import logging
from .config import load_config
from .pipeline import process_latest

def main():
    parser = argparse.ArgumentParser(prog="estoque-sem-giro", description="Consolidado + por PDV + arquivamento")
    # Mantemos simples: um único comando (processar o mais recente)
    args = parser.parse_args()
    ok = process_latest(load_config())
    raise SystemExit(0 if ok else 4)
