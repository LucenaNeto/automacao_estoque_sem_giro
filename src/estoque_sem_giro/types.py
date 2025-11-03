from __future__ import annotations
from typing import TypedDict

class Registro(TypedDict, total=False):
    PDV: str
    SKU: str
    DESCRIÇÃO: str
    MARCA: str
    CURVA: str
    ESTOQUE_ATUAL: str
