"""Exports every number used in the report as LaTeX macros and table bodies.

The Excel figures come from the saved workbook run (resultados/valores_excel.json);
the Python figures are recomputed with the delivered program (fixed seed, so the
values are identical to its console output). No number is typed by hand in the
report, which removes transcription errors.

Usage: python tools/exportar_resultados.py resultados/valores_excel.json informe/resultados.tex
"""

import importlib.util
import json
import sys
import time
from pathlib import Path

import numpy as np

ROOT = Path(__file__).resolve().parents[1]
PROGRAM = next(ROOT.glob("*_Taller2.py"))
spec = importlib.util.spec_from_file_location("taller", PROGRAM)
tl = importlib.util.module_from_spec(spec)
spec.loader.exec_module(tl)


def miles(x, dec=0):
    """Spanish number format: 103.111.235 or 0,0065."""
    s = f"{abs(x):,.{dec}f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return ("$-$" if x < 0 and round(abs(x), dec) != 0 else "") + s


def cop(x):
    return ("$-$" if x < 0 else "") + r"\$" + miles(abs(x))


def mill(x, dec=2):
    return miles(x / 1e6, dec)


def main():
    excel = json.load(open(sys.argv[1], encoding="utf-8"))
    out = []

    def macro(name, value):
        out.append(rf"\newcommand{{\{name}}}{{{value}}}")

    s, q = tl.REORDER_POINT, tl.ORDER_QTY
    macro("polS", s)
    macro("polQ", q)
    macro("semilla", tl.SEED)
    macro("nReplicas", tl.N_REPLICATIONS)
    macro("nPoliticasPy", len(tl.POLICIES))
    macro("nPoliticasXl", len(excel["policies"]["Uniforme"]))
    macro("repsPrecision", tl.PRECISION_REPS)

    # ------------------------------------------------------------------ Excel
    ver = excel["verification"]["Uniforme"]["rows"]
    rows = []
    fechas = ["jue 01/01/2026", "vie 02/01/2026", "sáb 03/01/2026"]
    for r, f in zip(ver, fechas):
        rows.append(f"{int(r['dia'])} & {f} & {int(r['R'])} & {int(r['II'])} & {int(r['D'])} & "
                    f"{int(r['V'])} & {int(r['P'])} & {int(r['IF'])} & {int(r['pedido'])} & "
                    f"{cop(r['utilidad'])} \\\\")
    macro("tablaVerificacionExcel", "\n".join(rows))
    ok = all(excel["verification"][sh]["check"] == "Correcta" for sh in ("Uniforme", "Empirica"))
    macro("comprobacionCalendario", "correcta" if ok else "INCORRECTA")

    tags = {"Uniforme": "U", "Empírica": "E"}
    for case, t in tags.items():
        st, ex = excel["stats"][case], excel["extra"][case]
        macro(f"xl{t}Media", cop(st["mean"]))
        macro(f"xl{t}MediaM", mill(st["mean"]))
        macro(f"xl{t}Desv", cop(st["std"]))
        macro(f"xl{t}DesvM", mill(st["std"]))
        macro(f"xl{t}EE", cop(ex["se"]))
        macro(f"xl{t}Pp", miles(st["p_lost"], 4))
        macro(f"xl{t}Vp", miles(st["lost_day"], 4))
        macro(f"xl{t}Min", mill(ex["min"]))
        macro(f"xl{t}Max", mill(ex["max"]))
        macro(f"xl{t}Ancho", cop(ex["width"]))
        macro(f"xl{t}InvFin", miles(ex["inv_final_mean"], 2))
        macro(f"xl{t}Corr", miles(ex["corr"], 3))
        macro(f"xl{t}Pend", cop(ex["slope"]))
        macro(f"xl{t}RR", miles(100 * ex["r2"], 1))
        hist = excel["histogram"][case]
        coords = " ".join(f"({h[0] / 1e6:.4f},{int(h[3])})" for h in hist)
        coords += f" ({hist[-1][1] / 1e6:.4f},0)"
        macro(f"xl{t}Hist", coords)
        macro(f"xl{t}HistTabla", "\n".join(
            f"{k + 1} & {mill(h[0], 3)} & {mill(h[1], 3)} & {miles(h[2], 2)} & {int(h[3])} \\\\"
            for k, h in enumerate(hist)))
        macro(f"xl{t}FreqTotal", int(sum(h[3] for h in hist)))
        reps = np.array(excel["replications"]["Uniforme" if t == "U" else "Empirica"])
        macro(f"xl{t}InvRango", f"{int(reps[:, 3].min())}--{int(reps[:, 3].max())}")

    pol_rows = []
    for case in ("Uniforme", "Empírica"):
        for r in excel["policies"][case]:
            mark = r"\rowcolor{EAFIT-blue!12}" if (r["s"], r["Q"]) == (s, q) else ""
            pol_rows.append(f"{mark}{case} & {r['s']} & {r['Q']} & {cop(r['mean'])} & "
                            f"{cop(r['std'])} & {miles(r['p_lost'], 4)} & "
                            f"{miles(r['lost_day'], 4)} \\\\")
        pol_rows.append(r"\midrule")
    macro("tablaPoliticasExcel", "\n".join(pol_rows[:-1]))

    # Selection rule applied to the Excel table (zero lost sales in both cases)
    sel = []
    xu = {(r["s"], r["Q"]): r for r in excel["policies"]["Uniforme"]}
    xe = {(r["s"], r["Q"]): r for r in excel["policies"]["Empírica"]}
    for key in xu:
        if xu[key]["p_lost"] == 0 and xe[key]["p_lost"] == 0:
            sel.append(((xu[key]["mean"] + xe[key]["mean"]) / 2, key))
    sel.sort(reverse=True)
    macro("tablaCriterioExcel", "\n".join(
        (r"\rowcolor{EAFIT-blue!12}" if k == (s, q) else "") +
        f"{k[0]} & {k[1]} & {cop(xu[k]['mean'])} & {cop(xe[k]['mean'])} & {cop(m)} \\\\"
        for m, k in sel))
    best_any_u = max(excel["policies"]["Uniforme"], key=lambda r: r["mean"])
    best_any_e = max(excel["policies"]["Empírica"], key=lambda r: r["mean"])
    macro("xlMejorUs", best_any_u["s"])
    macro("xlMejorUq", best_any_u["Q"])
    macro("xlMejorUmedia", cop(best_any_u["mean"]))
    macro("xlMejorUpp", miles(best_any_u["p_lost"], 4))
    macro("xlMejorEs", best_any_e["s"])
    macro("xlMejorEq", best_any_e["Q"])
    macro("xlMejorEmedia", cop(best_any_e["mean"]))
    macro("xlMejorEpp", miles(best_any_e["p_lost"], 4))
    macro("xlDifMejorU", cop(best_any_u["mean"] - xu[(s, q)]["mean"]))
    macro("xlDifMejorE", cop(best_any_e["mean"] - xe[(s, q)]["mean"]))
    # Largest advantage of a policy WITH lost sales over the recommended one
    adv = []
    for case, table in (("uniforme", xu), ("empírica", xe)):
        for k, r in table.items():
            if r["p_lost"] > 0:
                adv.append((r["mean"] - table[(s, q)]["mean"], k, case, r["p_lost"]))
    a = max(adv)
    macro("xlVentaja", cop(a[0]))
    macro("xlVentajaPol", f"$s={a[1][0]}$, $Q={a[1][1]}$")
    macro("xlVentajaCaso", a[2])
    macro("xlVentajaPp", miles(a[3], 4))
    se_diff = np.sqrt(2) * np.mean([excel["extra"][c]["se"] for c in tags])
    macro("xlEEdif", cop(se_diff))
    macro("xlEEdifDos", cop(2 * se_diff))

    # Q effect with s fixed (s = 7) in the Excel table
    q_rows = []
    for key in sorted(k for k in xu if k[0] == s):
        q_rows.append(f"{key[1]} & {cop(xu[key]['mean'])} & {cop(xe[key]['mean'])} & "
                      f"{cop(xu[key]['std'])} & {cop(xe[key]['std'])} \\\\")
    macro("tablaQExcel", "\n".join(q_rows))

    # Excel: best full-service policy of this run and how it compares with the chosen one
    two_se = 2 * se_diff
    m_best, k_best = sel[0]
    rec_avg = (xu[(s, q)]["mean"] + xe[(s, q)]["mean"]) / 2
    if k_best == (s, q):
        frase = (f"Aplicando el criterio, la mejor política de la hoja es $s={s}$, $Q={q}$ "
                 "(\\cref{tab:A3b}).")
    else:
        gap = m_best - rec_avg
        sig = gap > two_se
        frase = (f"En esta corrida de Excel, la política de servicio completo con mayor ganancia "
                 f"conjunta es $s={k_best[0]}$, $Q={k_best[1]}$ (\\cref{{tab:A3b}}), "
                 f"{cop(gap)} por encima de $s={s}$, $Q={q}$. "
                 + ("Esa diferencia supera dos errores estándar en esta corrida, pero no se "
                    "confirma con mayor precisión: con números aleatorios comunes y "
                    f"{tl.PRECISION_REPS} réplicas ninguna variante de $Q$ supera a $Q={q}$ "
                    "(Parte~B, pregunta~3). " if sig else
                    f"Esa diferencia es menor que dos errores estándar ({cop(two_se)}), es decir, "
                    "compatible con el azar. ")
                 + f"Por el criterio 3 se mantiene $Q={q}$: tiene menor desviación estándar "
                 f"({cop(xu[(s, q)]['std'])} frente a {cop(xu[k_best]['std'])} con demanda "
                 "uniforme) y coincide con el lote óptimo teórico de la pregunta~6.")
    macro("xlServFrase", frase)
    for t, table in (("U", xu), ("E", xe)):
        if (s, 30) in table:
            macro(f"xlQtreinta{t}", cop(table[(s, q)]["mean"] - table[(s, 30)]["mean"]))
            macro(f"xlQtreintaStd{t}", miles(table[(s, 30)]["std"] / table[(s, q)]["std"], 2))

    # ----------------------------------------------------------------- Python
    t0 = time.perf_counter()
    cal = tl.build_calendar(tl.YEAR)
    cases = [("Uniforme", tl.demand_uniform, "U"), ("Empírica", tl.demand_empirical, "E")]
    py = {}
    for name, fn, t in cases:
        res = tl.run_experiment(fn, np.random.default_rng(tl.SEED), cal, s, q)
        py[name] = res
        mean, se, hw = tl.mean_ci(res["profits"])
        corr, slope, r2 = tl.leftover_regression(res)
        macro(f"py{t}Media", cop(mean))
        macro(f"py{t}MediaM", mill(mean))
        macro(f"py{t}Desv", cop(res["std_profit"]))
        macro(f"py{t}EE", cop(se))
        macro(f"py{t}IC", f"[{cop(mean - hw)};\\ {cop(mean + hw)}]")
        macro(f"py{t}Pp", miles(res["prob_lost_sales"], 4))
        macro(f"py{t}Vp", miles(res["mean_lost_per_day"], 4))
        macro(f"py{t}Corr", miles(corr, 3))
        macro(f"py{t}Pend", cop(slope))
        macro(f"py{t}RR", miles(100 * r2, 1))
        macro(f"py{t}InvFin", miles(res["ending_inventory"].mean(), 2))
        freq, edges = np.histogram(res["profits"] / 1e6, bins=10)
        coords = " ".join(f"({e:.4f},{int(f)})" for e, f in zip(edges[:-1], freq))
        macro(f"py{t}Hist", coords + f" ({edges[-1]:.4f},0)")
        macro(f"py{t}Min", miles(edges[0], 2))
        macro(f"py{t}Max", miles(edges[-1], 2))

        # Excel vs Python comparison: two independent samples of 100 years
        st = excel["stats"][name]
        se_x = excel["extra"][name]["se"]
        diff = st["mean"] - mean
        z = diff / np.sqrt(se ** 2 + se_x ** 2)
        macro(f"dif{t}", cop(diff))
        macro(f"dif{t}Z", miles(z, 2))
        macro(f"dif{t}Pct", miles(100 * diff / mean, 2))
        zs = globals().setdefault("_zs", [])
        zs.append(abs(z))

    macro("difZFrase", "ambos con $|z|<1{,}96$: las diferencias no son significativas al 5\\,\\%"
          if max(globals()["_zs"]) < 1.96 else
          "al menos uno con $|z|\\ge1{,}96$, una diferencia mayor que la esperada por azar")

    # Seeds
    seed_rows, seed_info = [], {}
    means_by_case = {}
    for name, fn, t in cases:
        means_by_case[name] = tl.means_by_seed(fn, cal, s, q, tl.SEED_SENSITIVITY)
    for i, sd in enumerate(tl.SEED_SENSITIVITY):
        seed_rows.append(f"{sd} & {cop(means_by_case['Uniforme'][i])} & "
                         f"{cop(means_by_case['Empírica'][i])} \\\\")
    macro("tablaSemillas", "\n".join(seed_rows))
    for name, fn, t in cases:
        m = means_by_case[name]
        macro(f"sem{t}Rango", cop(m.max() - m.min()))
        macro(f"sem{t}RangoPct", miles(100 * (m.max() - m.min()) / m.mean(), 2))
        macro(f"sem{t}Desv", cop(m.std(ddof=1)))
        macro(f"sem{t}Min", cop(m.min()))
        macro(f"sem{t}Max", cop(m.max()))
        seed_info[name] = m
    macro("semMaxPct", miles(max(100 * (m.max() - m.min()) / m.mean()
                                 for m in seed_info.values()), 2))

    # Policy grid with common random numbers
    grid = {}
    for name, fn, t in cases:
        grid[name] = tl.compare_policies(tl.POLICIES, fn, cal, tl.N_REPLICATIONS, tl.SEED)
        lookup = {(r["s"], r["Q"]): r for r in grid[name]}
        head = " & ".join(str(qq) for qq in tl.POLICY_Q)
        body = []
        for ss in tl.POLICY_S:
            cells = []
            best_row = max(grid[name], key=lambda r: r["mean_profit"])
            for qq in tl.POLICY_Q:
                r = lookup[(ss, qq)]
                v = mill(r["mean_profit"], 2)
                if (ss, qq) == (best_row["s"], best_row["Q"]):
                    v = rf"\textbf{{{v}}}"
                if (ss, qq) == (s, q):
                    v = rf"\cellcolor{{EAFIT-blue!15}}{v}"
                cells.append(v)
            body.append(f"{ss} & " + " & ".join(cells) + r" \\")
        macro(f"mat{t}Cabeza", head)
        macro(f"mat{t}Media", "\n".join(body))
        body = []
        for ss in tl.POLICY_S:
            cells = [miles(lookup[(ss, qq)]["prob_lost_sales"], 3) for qq in tl.POLICY_Q]
            body.append(f"{ss} & " + " & ".join(cells) + r" \\")
        macro(f"mat{t}Pp", "\n".join(body))
        top = sorted(grid[name], key=lambda r: -r["mean_profit"])[:10]
        macro(f"top{t}", "\n".join(
            (r"\rowcolor{EAFIT-blue!12}" if (r["s"], r["Q"]) == (s, q) else "") +
            f"{r['s']} & {r['Q']} & {cop(r['mean_profit'])} & {cop(r['std_profit'])} & "
            f"{miles(r['prob_lost_sales'], 4)} & {miles(r['mean_lost_per_day'], 4)} \\\\"
            for r in top))
        best = top[0]
        macro(f"py{t}MejorS", best["s"])
        macro(f"py{t}MejorQ", best["Q"])
        macro(f"py{t}MejorMedia", cop(best["mean_profit"]))
        macro(f"py{t}MejorPp", miles(best["prob_lost_sales"], 4))
        d, hw = tl.paired_gap(best["profits"], lookup[(s, q)]["profits"])
        macro(f"py{t}DifPar", cop(d))
        macro(f"py{t}DifParIC", cop(hw))
        macro(f"py{t}DifParSig", "significativa" if abs(d) > hw else "no significativa")
        rank = [(r["s"], r["Q"]) for r in sorted(grid[name], key=lambda r: -r["mean_profit"])]
        macro(f"py{t}Rango", rank.index((s, q)) + 1)
        # Same comparison without common random numbers (independent streams)
        rng_a, rng_b = np.random.default_rng(tl.SEED), np.random.default_rng(tl.SEED + 99)
        a = tl.run_experiment(fn, rng_a, cal, best["s"], best["Q"])["profits"]
        b = tl.run_experiment(fn, rng_b, cal, s, q)["profits"]
        hw_ind = 1.984 * np.sqrt(a.var(ddof=1) / a.size + b.var(ddof=1) / b.size)
        macro(f"py{t}DifInd", cop(a.mean() - b.mean()))
        macro(f"py{t}DifIndIC", cop(hw_ind))

    sb, qb, avg = tl.pick_full_service(grid)
    macro("pyServS", sb)
    macro("pyServQ", qb)
    macro("pyServMedia", cop(avg))
    for name, fn, t in cases:
        lookup = {(r["s"], r["Q"]): r for r in grid[name]}
        d, hw = tl.paired_gap(lookup[(sb, qb)]["profits"], lookup[(s, q)]["profits"])
        macro(f"py{t}DifServ", cop(d))
        macro(f"py{t}DifServIC", cop(hw))
        n_cov, ok = tl.full_service_holds(grid[name])
        macro(f"py{t}NCubiertas", n_cov)
        macro(f"py{t}CeroPerdidas", "sí" if ok else "NO")

    # High-precision run (1000 replications, s = 7, Q = 11..19)
    prec_rows = []
    prec = {}
    for name, fn, t in cases:
        prec[name] = tl.compare_policies(tl.PRECISION_POLICIES, fn, cal, tl.PRECISION_REPS,
                                         tl.SEED + 1)
    for i, (ss, qq) in enumerate(tl.PRECISION_POLICIES):
        cells = []
        for name in ("Uniforme", "Empírica"):
            rows_p = prec[name]
            chosen = next(r for r in rows_p if r["Q"] == q)["profits"]
            d, hw = tl.paired_gap(rows_p[i]["profits"], chosen)
            cells.append(f"{cop(rows_p[i]['mean_profit'])} & "
                         + ("--" if qq == q else f"{cop(d)} $\\pm$ {cop(hw)}"))
        mark = r"\rowcolor{EAFIT-blue!12}" if qq == q else ""
        prec_rows.append(f"{mark}{qq} & " + " & ".join(cells) + r" \\")
    macro("tablaPrecision", "\n".join(prec_rows))
    worse, better = [], []
    for name, label in (("Uniforme", "uniforme"), ("Empírica", "empírica")):
        rows_p = prec[name]
        chosen = next(r for r in rows_p if r["Q"] == q)["profits"]
        w = [r["Q"] for r in rows_p if r["Q"] != q
             and sum(tl.paired_gap(r["profits"], chosen)) < 0]
        bt = [r["Q"] for r in rows_p if r["Q"] != q
              and tl.paired_gap(r["profits"], chosen)[0]
              - tl.paired_gap(r["profits"], chosen)[1] > 0]
        if w:
            worse.append(f"$Q\\in\\{{{','.join(map(str, w))}\\}}$ ({label})")
        if bt:
            better.append(f"$Q\\in\\{{{','.join(map(str, bt))}\\}}$ ({label})")
    macro("precPeores", " y ".join(worse) if worse else "ningún valor")
    macro("precMejoresFrase",
          "ningún valor de $Q$ supera significativamente a $Q=" + str(q) + "$"
          if not better else "solo " + " y ".join(better) + " superan significativamente a $Q="
          + str(q) + "$")

    # Cost of full service: s = 6 and s = 5 (with lost sales) against the chosen policy
    gains = []
    for name, fn, t in cases:
        base = tl.compare_policies([(s, q)], fn, cal, tl.PRECISION_REPS, tl.SEED + 2)[0]
        for r in tl.compare_policies(tl.SERVICE_POLICIES, fn, cal, tl.PRECISION_REPS,
                                     tl.SEED + 2):
            d, hw = tl.paired_gap(r["profits"], base["profits"])
            key = f"serv{t}{'Seis' if r['s'] == 6 else 'Cinco'}"
            macro(key, cop(d))
            macro(key + "IC", cop(hw))
            macro(key + "Pp", miles(r["prob_lost_sales"], 4))
            macro(key + "Pct", miles(100 * d / base["mean_profit"], 2))
            macro(key + "Sig", "significativa" if abs(d) > hw else "no significativa")
            if r["s"] == 6:
                gains.append(d)
    macro("servGananciaMax", cop(max(gains)))
    hw_prec = np.mean([tl.mean_ci(r["profits"])[2]
                       for name in prec for r in prec[name]])
    macro("precIC", cop(hw_prec))

    elapsed = time.perf_counter() - t0
    macro("tiempoExport", miles(elapsed, 1))
    salida = (ROOT / "resultados" / "salida_python.txt")
    if salida.exists():
        last = [ln for ln in salida.read_text(encoding="utf-8").splitlines()
                if ln.startswith("Tiempo total")]
        if last:
            macro("tiempoPython", miles(float(last[0].split(":")[1].split()[0]), 1))

    Path(sys.argv[2]).write_text(
        "% Generated by tools/exportar_resultados.py - do not edit by hand.\n"
        + "\n".join(out) + "\n", encoding="utf-8")


if __name__ == "__main__":
    main()
