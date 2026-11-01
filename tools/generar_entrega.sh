#!/usr/bin/env bash
# Regenerates every deliverable of the workshop from source.
#   tools/generar_entrega.sh              -> reuses resultados/valores_excel.json
#   tools/generar_entrega.sh --recalcular -> new spreadsheet run in LibreOffice
set -euo pipefail
cd "$(dirname "$0")/.."
PREFIJO="YoungHoyos_JuanAndres_Taller2"
mkdir -p build resultados informe

# Part B: Python program (verification, statistics, histograms, policies)
uv run python "${PREFIJO}.py" | tee resultados/salida_python.txt

# Part A: workbook with formulas -> spreadsheet computation -> final workbook
uv run python tools/construir_excel.py build/base.xlsx
if [[ "${1:-}" == "--recalcular" || ! -f resultados/valores_excel.json ]]; then
  /usr/bin/python3 tools/calcular_libreoffice.py build/base.xlsx resultados/valores_excel.json
fi
uv run python tools/construir_excel.py "${PREFIJO}.xlsx" resultados/valores_excel.json
uv run python tools/inyectar_valores.py "${PREFIJO}.xlsx" resultados/valores_excel.json

# Numbers and tables for the LaTeX report (Overleaf)
uv run python tools/exportar_resultados.py resultados/valores_excel.json informe/resultados.tex
echo "Listo: ${PREFIJO}.xlsx, ${PREFIJO}.py, informe/resultados.tex"
