# estoque-sem-giro

Automação para:
- Ler abas **EUD/BOT/QDB**
- Extrair colunas **A (SKU), C (DESCRIÇÃO), E (CURVA/CLASSE), I (PDV), J (ESTOQUE_ATUAL)**
- Gerar **CSV consolidado** (agrupado por PDV) e **CSV por PDV**
- Arquivar o `.xlsx` processado

## Instalação (modo dev)
```bash
pip install -e .
