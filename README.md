# Taller 2 – Simulación Montecarlo de un sistema de inventarios

Métodos Cuantitativos, Universidad EAFIT · Juan Andrés Young Hoyos (código 1000337497) · entrega individual

Política recomendada: **s = 7, Q = 15** (cero ventas perdidas con ambas demandas y ganancia
estadísticamente equivalente a la máxima).

## Entregables (EAFIT Interactiva)

| Archivo | Contenido |
|---|---|
| `YoungHoyos_JuanAndres_Taller2.xlsx` | Hojas `Parametros`, `Uniforme`, `Empirica` y `Politicas`, con la política recomendada cargada |
| `YoungHoyos_JuanAndres_Taller2.py` | Programa completo: `python YoungHoyos_JuanAndres_Taller2.py` |
| `YoungHoyos_JuanAndres_Taller2_Python.pdf` | Análisis de las Partes A y B (se descarga del proyecto de Overleaf) |

## Reproducir todo

```bash
uv sync
tools/generar_entrega.sh               # reutiliza la corrida guardada de la hoja de cálculo
tools/generar_entrega.sh --recalcular  # nueva corrida de las 100 réplicas y de Politicas
```

El script ejecuta el programa de Python, construye el libro de Excel con las fórmulas del
enunciado (`tools/construir_excel.py`), lo calcula con el motor de LibreOffice Calc
(`tools/calcular_libreoffice.py`), guarda los valores calculados en el libro
(`tools/inyectar_valores.py`) y genera `informe/resultados.tex` con todas las cifras del
informe (`tools/exportar_resultados.py`).

- `informe/contenido.tex` → `Chapters/1. Content.tex` en Overleaf
- `informe/resultados.tex` → `General/Resultados.tex` en Overleaf
- `informe/referencias.bib` → `General/References.bib` en Overleaf
- `resultados/valores_excel.json`: corrida guardada del libro (réplicas, políticas, verificación)
- `resultados/salida_python.txt`: salida de consola del programa

Enunciado: `Taller 2 - Simulación Montecarlo.pdf`.
