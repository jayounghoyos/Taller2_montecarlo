"""Computes the formulas of the "Politicas" sheet with LibreOffice Calc.

Run with the system Python that ships the LibreOffice UNO bridge:
    /usr/bin/python3 tools/calcular_politicas.py LIBRO.xlsx salida.json

The sheet only depends on its own recorded values and on Parametros, so the
values computed here are exactly what Excel shows. They are stored as cached
values by tools/inyectar_valores.py.
"""

import json
import os
import subprocess
import sys
import time

import uno  # noqa: F401  (LibreOffice bridge)
from com.sun.star.beans import PropertyValue


def prop(name, value):
    p = PropertyValue()
    p.Name, p.Value = name, value
    return p


def main():
    src, out = os.path.abspath(sys.argv[1]), sys.argv[2]
    proc = subprocess.Popen(
        ["soffice", "--headless", "--invisible", "--norestore", "--nologo",
         "--accept=socket,host=localhost,port=2086;urp;"],
        stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    local = uno.getComponentContext()
    resolver = local.ServiceManager.createInstanceWithContext(
        "com.sun.star.bridge.UnoUrlResolver", local)
    for _ in range(60):
        try:
            ctx = resolver.resolve(
                "uno:socket,host=localhost,port=2086;urp;StarOffice.ComponentContext")
            break
        except Exception:
            time.sleep(0.5)
    else:
        raise RuntimeError("Could not connect to LibreOffice")
    desktop = ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
    try:
        doc = desktop.loadComponentFromURL(uno.systemPathToFileUrl(src), "_blank", 0,
                                           (prop("Hidden", True),))
        doc.calculateAll()
        sheet = doc.Sheets.getByName("Politicas")
        cursor = sheet.createCursor()
        cursor.gotoEndOfUsedArea(False)
        area = sheet.getCellRangeByPosition(0, 0, cursor.RangeAddress.EndColumn,
                                            cursor.RangeAddress.EndRow)
        forms, vals = area.getFormulaArray(), area.getDataArray()
        cached = {}
        errors = []
        for i, frow in enumerate(forms):
            for j, f in enumerate(frow):
                if isinstance(f, str) and f.startswith("="):
                    cell = sheet.getCellByPosition(j, i)
                    if cell.getError():
                        errors.append((j, i, f))
                    cached[f"{j},{i}"] = vals[i][j]
        doc.close(True)
    finally:
        try:
            desktop.terminate()
        except Exception:
            pass
        proc.wait(timeout=30)
    if errors:
        raise SystemExit(f"Formula errors in Politicas: {errors[:5]}")
    with open(out, "w", encoding="utf-8") as fh:
        json.dump({"cached": {"Politicas": cached}}, fh, ensure_ascii=False, indent=1)


if __name__ == "__main__":
    main()
