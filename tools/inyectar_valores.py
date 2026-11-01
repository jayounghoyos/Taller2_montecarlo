"""Stores the computed value of every formula cell inside the .xlsx (cached <v>).

openpyxl writes formulas without results. This step adds the values computed
by the spreadsheet engine (tools/calcular_libreoffice.py) so the workbook opens
showing the saved replication in any viewer, while every formula and the data
table remain live.

Usage: python tools/inyectar_valores.py LIBRO.xlsx valores.json
"""

import json
import re
import shutil
import sys
import tempfile
import zipfile
from xml.sax.saxutils import escape

from openpyxl.utils import get_column_letter

SHEET_FILES = {"Parametros": "xl/worksheets/sheet1.xml",
               "Uniforme": "xl/worksheets/sheet2.xml",
               "Empirica": "xl/worksheets/sheet3.xml"}

CELL_RE = re.compile(r'<c r="([A-Z]+\d+)"([^>]*?)>(<f[^>]*?(?:/>|>.*?</f>))<v\s*/>(</c>)', re.S)


def inject(xml, values):
    missing = []

    def repl(m):
        ref, attrs, formula, close = m.groups()
        if ref not in values:
            missing.append(ref)
            return m.group(0)
        v = values[ref]
        attrs = re.sub(r'\s+t="[^"]*"', "", attrs)
        if isinstance(v, str):
            return f'<c r="{ref}"{attrs} t="str">{formula}<v>{escape(v)}</v>{close}'
        number = int(v) if v == int(v) else repr(float(v))
        return f'<c r="{ref}"{attrs}>{formula}<v>{number}</v>{close}'

    return CELL_RE.sub(repl, xml), missing


def main():
    path, json_path = sys.argv[1], sys.argv[2]
    data = json.load(open(json_path, encoding="utf-8"))
    by_sheet = {}
    for sheet, cells in data["cached"].items():
        by_sheet[sheet] = {}
        for key, value in cells.items():
            col, row = (int(x) for x in key.split(","))
            by_sheet[sheet][f"{get_column_letter(col + 1)}{row + 1}"] = value
    # X7 holds the data table formula; its value is replication 1, column X
    by_sheet["Uniforme"]["X7"] = data["replications"]["Uniforme"][0][0]
    by_sheet["Empirica"]["X7"] = data["replications"]["Empirica"][0][0]

    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx").name
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            content = zin.read(item.filename)
            for sheet, fname in SHEET_FILES.items():
                if item.filename == fname:
                    xml, missing = inject(content.decode("utf-8"), by_sheet[sheet])
                    if missing:
                        raise SystemExit(f"{sheet}: no value for {missing[:10]}")
                    content = xml.encode("utf-8")
            zout.writestr(item, content)
    shutil.move(tmp, path)


if __name__ == "__main__":
    main()
