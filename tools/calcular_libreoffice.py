"""Computes the Part A workbook with a real spreadsheet engine (LibreOffice Calc).

Run with the system Python that ships the LibreOffice UNO bridge:
    /usr/bin/python3 tools/calcular_libreoffice.py build/base.xlsx build/valores.json

Steps (all on the workbook built by tools/construir_excel.py):
 1. Verification example of the statement: s=6, Q=10, demands 4, 7, 5 typed in
    G6:G8 of both sheets -> daily profits must be -4.320.640, 4.079.760, -2.640.600.
 2. Step A9: for each candidate policy, set s and Q in Parametros, recalculate
    (each recalculation re-evaluates RAND() and the data table) and record the
    four statistics of each sheet.
 3. Recommended policy: final recalculation; the 100 replications of the data
    table and the value of every formula cell are exported to JSON so that they
    can be stored in the delivered .xlsx as its saved state.
"""

import json
import os
import subprocess
import sys
import time

import uno  # noqa: F401  (LibreOffice bridge)
from com.sun.star.beans import PropertyValue

# The recommended policy goes last: its run is the state saved in the workbook,
# so the Politicas row and the Uniforme/Empirica sheets show the same numbers.
POLICIES = [(5, 10), (6, 10), (7, 10), (5, 15), (6, 15), (7, 20), (7, 30), (8, 15), (7, 15)]
RECOMMENDED = (7, 15)
SHEETS = ("Uniforme", "Empirica")
CASE_LABEL = {"Uniforme": "Uniforme", "Empirica": "Empírica"}


def prop(name, value):
    p = PropertyValue()
    p.Name, p.Value = name, value
    return p


def connect():
    proc = subprocess.Popen(
        ["soffice", "--headless", "--invisible", "--norestore", "--nologo",
         "--accept=socket,host=localhost,port=2083;urp;"],
        stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    local = uno.getComponentContext()
    resolver = local.ServiceManager.createInstanceWithContext(
        "com.sun.star.bridge.UnoUrlResolver", local)
    for _ in range(60):
        try:
            ctx = resolver.resolve(
                "uno:socket,host=localhost,port=2083;urp;StarOffice.ComponentContext")
            break
        except Exception:
            time.sleep(0.5)
    else:
        raise RuntimeError("Could not connect to LibreOffice")
    desktop = ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
    return proc, desktop


def cell(doc, sheet, addr):
    return doc.Sheets.getByName(sheet).getCellRangeByName(addr)


def stats(doc, sheet):
    v = [cell(doc, sheet, f"AC{r}").getValue() for r in (5, 6, 7, 8)]
    return {"mean": v[0], "std": v[1], "p_lost": v[2], "lost_day": v[3]}


def set_policy(doc, s, q):
    cell(doc, "Parametros", "B4").setValue(s)
    cell(doc, "Parametros", "B5").setValue(q)
    doc.calculateAll()


def run_data_table(doc, n_reps=100):
    """Replicates what Excel's one-variable data table does for every row.

    Excel writes the replication number in the (unused) input cell W2 and
    recalculates the whole workbook, so RAND() is drawn again for each row.
    LibreOffice's TABLE()/MULTIPLE.OPERATIONS does not re-draw volatile
    functions, so the same procedure is executed explicitly here and the
    results are written into X7:AA106 (columns X, Y, Z, AA = U6, U7, U8, U9).
    """
    rows = {sh: [] for sh in SHEETS}
    for i in range(1, n_reps + 1):
        cell(doc, "Uniforme", "W2").setValue(i)
        cell(doc, "Empirica", "W2").setValue(i)
        doc.calculateAll()
        for sh in SHEETS:
            rows[sh].append([cell(doc, sh, f"U{r}").getValue() for r in (6, 7, 8, 9)])
    for sh in SHEETS:
        cell(doc, sh, "W2").setString("")
        doc.Sheets.getByName(sh).getCellRangeByName("X7:AA106").setDataArray(
            tuple(tuple(r) for r in rows[sh]))
    doc.calculateAll()   # statistics and histogram now read the 100 replications
    return rows


def main():
    src, out = os.path.abspath(sys.argv[1]), sys.argv[2]
    proc, desktop = connect()
    try:
        doc = desktop.loadComponentFromURL(uno.systemPathToFileUrl(src), "_blank", 0,
                                           (prop("Hidden", True),))
        result = {"recommended": list(RECOMMENDED)}

        # Diagnostics: was the data table imported as a live formula?
        x8 = cell(doc, "Uniforme", "X8")
        result["data_table_formula_as_imported"] = x8.getFormula()

        # LibreOffice cannot evaluate Excel's data table with RAND(); remove it
        # from this in-memory copy and run the equivalent loop (run_data_table).
        for sh in SHEETS:
            doc.Sheets.getByName(sh).getCellRangeByName("X7:AA106").clearContents(1 | 2 | 4 | 16)

        # 1. Verification example
        verification = {}
        set_policy(doc, 6, 10)
        for sh in SHEETS:
            saved = [cell(doc, sh, f"G{r}").getFormula() for r in (6, 7, 8)]
            for r, d in zip((6, 7, 8), (4, 7, 5)):
                cell(doc, sh, f"G{r}").setValue(d)
            doc.calculateAll()
            rows = []
            for r in (6, 7, 8):
                rows.append({k: cell(doc, sh, f"{c}{r}").getValue() for k, c in
                             (("dia", "A"), ("R", "E"), ("II", "F"), ("D", "G"), ("V", "H"),
                              ("P", "I"), ("IF", "J"), ("pedido", "L"), ("utilidad", "R"))})
            verification[sh] = {
                "rows": rows,
                "A318": cell(doc, sh, "A318").getValue(),
                "B318": cell(doc, sh, "B318").getString(),
                "sumD": sum(cell(doc, sh, f"D{r}").getValue() for r in range(6, 319)),
                "check": cell(doc, sh, "U10").getString(),
            }
            for r, f in zip((6, 7, 8), saved):
                cell(doc, sh, f"G{r}").setFormula(f)
        result["verification"] = verification

        # 2. Step A9: policy comparison
        policies = {"Uniforme": [], "Empírica": []}
        assert POLICIES[-1] == RECOMMENDED
        for s, q in POLICIES:
            set_policy(doc, s, q)
            reps = run_data_table(doc)
            for sh in SHEETS:
                policies[CASE_LABEL[sh]].append({"s": s, "Q": q, **stats(doc, sh)})
        # Present the table sorted by demand case, s and Q
        for case in policies:
            policies[case].sort(key=lambda r: (r["s"], r["Q"]))
        result["policies"] = policies

        # 3. The workbook is now in the state of the recommended policy (last run)
        formulas = {}
        for sh in ("Parametros",) + SHEETS:
            sheet = doc.Sheets.getByName(sh)
            used = sheet.createCursor()
            used.gotoEndOfUsedArea(False)
            last_row, last_col = used.RangeAddress.EndRow, used.RangeAddress.EndColumn
            rng = sheet.getCellRangeByPosition(0, 0, last_col, last_row)
            forms = rng.getFormulaArray()
            vals = rng.getDataArray()
            sheet_vals = {}
            for i, frow in enumerate(forms):
                for j, f in enumerate(frow):
                    if isinstance(f, str) and f.startswith("="):
                        sheet_vals[f"{j},{i}"] = vals[i][j]
            formulas[sh] = sheet_vals
        result["replications"] = reps
        result["cached"] = formulas
        result["stats"] = {CASE_LABEL[sh]: stats(doc, sh) for sh in SHEETS}
        result["histogram"] = {
            CASE_LABEL[sh]: [[cell(doc, sh, f"{c}{r}").getValue() for c in ("AC", "AD", "AE", "AF")]
                             for r in range(17, 27)] for sh in SHEETS}
        result["extra"] = {
            CASE_LABEL[sh]: {k: cell(doc, sh, a).getValue() for k, a in
                             (("se", "AC9"), ("min", "AC12"), ("max", "AC13"), ("width", "AC14"),
                              ("total", "AF27"), ("inv_final_mean", "AC30"),
                              ("corr", "AC31"), ("slope", "AC32"), ("r2", "AC33"))}
            for sh in SHEETS}
        doc.close(True)
    finally:
        try:
            desktop.terminate()
        except Exception:
            pass
        proc.wait(timeout=30)

    with open(out, "w", encoding="utf-8") as fh:
        json.dump(result, fh, ensure_ascii=False, indent=1)


if __name__ == "__main__":
    main()
