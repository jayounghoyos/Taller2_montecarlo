"""Monte Carlo simulation of a single-product (s, Q) inventory system - Part B.

Juan Andrés Young Hoyos · student ID 1000337497
Métodos Cuantitativos, Universidad EAFIT - Workshop 2

How to run:  python YoungHoyos_JuanAndres_Taller2.py
Needs:       Python 3.10+, NumPy, Matplotlib

"""

import datetime as dt
import time
from pathlib import Path

import matplotlib

matplotlib.use("Agg")  # draw straight to PNG files; no window is opened
import matplotlib.pyplot as plt  # noqa: E402
import numpy as np  # noqa: E402

# =============================================================================
# B1 · Inputs. Everything that can change lives here and nowhere else.
# =============================================================================
YEAR = 2026
N_REPLICATIONS = 100

REORDER_POINT = 7      # s: an order is placed when closing stock falls below s
ORDER_QTY = 15         # Q: units per order; also the stock received on day 1

PRICE = 680_000        # COP earned for every unit sold
UNIT_COST = 600_000    # COP paid for every unit received
ORDER_COST = 40_000    # COP paid each time an order arrives
HOLDING_COST = 80      # COP per unit per working day, on the day's average stock
FIXED_MONTHLY = 1_000_000  # COP charged on the first working day of every month

SEED = 1000337497      # my student ID: the base seed of every experiment

# Empirical demand: possible values and their probabilities (must add to 1)
EMP_VALUES = (3, 4, 5, 6, 7)
EMP_PROBS = (0.20, 0.05, 0.40, 0.25, 0.10)

# Policy grid for the comparison (6 x 12 = 72 combinations; Excel uses 9)
POLICY_S = (4, 5, 6, 7, 8, 9)
POLICY_Q = (8, 10, 12, 13, 14, 15, 16, 17, 18, 20, 25, 30)
POLICIES = [(s, q) for s in POLICY_S for q in POLICY_Q]

SEED_SENSITIVITY = (SEED, 2026, 42, 777, 31415)   # five seeds for question 2

# Longer runs used to settle close calls (questions 2 and 3)
PRECISION_REPS = 1000
PRECISION_POLICIES = [(7, q) for q in range(11, 20)]   # s fixed, Q around 15
SERVICE_POLICIES = [(6, 15), (5, 15)]                  # lower s, accepting stockouts

T_CRITICAL = 1.9842    # t(0.975; 99 d.f.) for 95 % intervals with 100 values
HERE = Path(__file__).resolve().parent


# =============================================================================
# B2 · Calendar: the working days of the year (Excel column B)
# =============================================================================
def build_calendar(year):
    """List every Monday-to-Saturday date of `year`; Sundays are skipped."""
    days = []
    current = dt.date(year, 1, 1)
    while current.year == year:
        if current.weekday() < 6:          # 0 = Monday ... 5 = Saturday
            days.append(current)
        current += dt.timedelta(days=1)
    return days


# =============================================================================
# B3 · Demand generators (Excel column G)
# =============================================================================
def demand_uniform(rng):
    """One day of demand, equally likely to be 3, 4, 5, 6 or 7 units."""
    return int(rng.integers(3, 8))         # the upper limit 8 is never drawn


def demand_empirical(rng):
    """One day of demand drawn from the empirical table by inverse transform.

    A uniform number u in [0, 1) is compared with the running total of the
    probabilities; the first value whose cumulative total exceeds u is the
    demand. This is the same rule as VLOOKUP(RAND(), table, 2, TRUE) in Excel.
    """
    u = rng.random()
    running_total = 0.0
    for value, probability in zip(EMP_VALUES, EMP_PROBS):
        running_total += probability
        if u < running_total:
            return value
    return EMP_VALUES[-1]                   # only reached through rounding


# =============================================================================
# B4 · One simulated year
# =============================================================================
def simulate_year(demands, calendar, s, q, keep_log=False):
    """Play one year day by day for a given list of daily demands.

    The demands are an input (not drawn inside) so the same sequence can be
    replayed under different policies, or typed by hand to check the model.
    Returns the annual profit, the share of days with a shortage, the lost
    units per day and the stock left on the last day.
    """
    if len(demands) != len(calendar):
        raise ValueError("Expected one demand per working day.")
    if s < 0 or q <= 0:
        raise ValueError("Use s >= 0 and Q > 0.")

    arriving = q            # day 1 starts by receiving the first Q units
    stock = 0
    year_profit = 0.0
    shortage_days = 0
    units_lost = 0
    log = []

    for i, today in enumerate(calendar):
        received = arriving                         # R_t
        opening = stock + received                  # II_t
        wanted = int(demands[i])                    # D_t
        sold = min(opening, wanted)                 # V_t
        lost = wanted - sold                        # P_t
        stock = opening - sold                      # IF_t
        arriving = q if stock < s else 0            # O_t: Q units arrive tomorrow

        new_month = i == 0 or calendar[i - 1].month != today.month   # M_t
        day_profit = (PRICE * sold
                      - UNIT_COST * received
                      - (ORDER_COST if received else 0)
                      - HOLDING_COST * (opening + stock) / 2
                      - (FIXED_MONTHLY if new_month else 0))         # U_t

        year_profit += day_profit
        units_lost += lost
        shortage_days += lost > 0

        if keep_log:
            log.append({"day": i + 1, "date": today, "received": received,
                        "opening": opening, "demand": wanted, "sold": sold,
                        "lost": lost, "closing": stock, "order": int(arriving > 0),
                        "profit": day_profit})

    n = len(calendar)
    return {"profit": year_profit,
            "shortage_fraction": shortage_days / n,
            "lost_per_day": units_lost / n,
            "ending_inventory": stock,
            "log": log}


# =============================================================================
# B5 · Many independent years for one demand case
# =============================================================================
def run_experiment(demand_fn, rng, calendar, s, q, n_reps=N_REPLICATIONS):
    """Simulate `n_reps` years with fresh demands and collect their results."""
    profit = np.empty(n_reps)
    shortage = np.empty(n_reps)
    lost = np.empty(n_reps)
    leftover = np.empty(n_reps)

    for k in range(n_reps):
        demands = [demand_fn(rng) for _ in calendar]   # 313 new demands
        year = simulate_year(demands, calendar, s, q)
        profit[k] = year["profit"]
        shortage[k] = year["shortage_fraction"]
        lost[k] = year["lost_per_day"]
        leftover[k] = year["ending_inventory"]

    # Every year has 313 days, so the mean of the yearly shares equals the
    # share over all simulated days (the definition in the assignment).
    return {"profits": profit,
            "ending_inventory": leftover,
            "mean_profit": float(profit.mean()),
            "std_profit": float(profit.std(ddof=1)),   # sample std, as STDEV.S
            "prob_lost_sales": float(shortage.mean()),
            "mean_lost_per_day": float(lost.mean())}


def mean_ci(values):
    """Mean, standard error and 95 % half-width of a sample of yearly results."""
    values = np.asarray(values, dtype=float)
    se = values.std(ddof=1) / np.sqrt(values.size)
    return float(values.mean()), float(se), float(T_CRITICAL * se)


def leftover_regression(result):
    """How much of the profit spread is explained by stock left on 31/12.

    Returns the correlation, the fitted COP lost per leftover unit and R^2.
    """
    x, y = result["ending_inventory"], result["profits"]
    r = float(np.corrcoef(x, y)[0, 1])
    return r, float(np.polyfit(x, y, 1)[0]), r * r


# =============================================================================
# B6 · Histogram of the annual profit
# =============================================================================
def plot_histogram(profits, title, filename):
    """Save a 10-bin histogram (equal widths from min to max) as a PNG file."""
    fig, ax = plt.subplots(figsize=(8, 4.5))
    ax.hist(np.asarray(profits) / 1e6, bins=10, color="#1d5fbf", edgecolor="white")
    ax.set_title(title)
    ax.set_xlabel("Ganancia anual (millones de COP)")
    ax.set_ylabel("Número de réplicas")
    ax.grid(axis="y", color="#e6e6e6")
    ax.set_axisbelow(True)
    fig.tight_layout()
    fig.savefig(HERE / filename, dpi=150)
    plt.close(fig)


# =============================================================================
# B7 · Comparing policies on the same random demands
# =============================================================================
def compare_policies(policies, demand_fn, calendar, n_reps, seed):
    """Run every (s, Q) pair against identical demand sequences.

    Restarting the generator from the same seed for each policy (common random
    numbers) means all policies live through exactly the same 100 years, so a
    difference between them is caused by the policy and not by luck.
    """
    table = []
    for s, q in policies:
        outcome = run_experiment(demand_fn, np.random.default_rng(seed), calendar,
                                 s, q, n_reps)
        table.append({"s": s, "Q": q,
                      "mean_profit": outcome["mean_profit"],
                      "std_profit": outcome["std_profit"],
                      "prob_lost_sales": outcome["prob_lost_sales"],
                      "mean_lost_per_day": outcome["mean_lost_per_day"],
                      "profits": outcome["profits"]})
    return table


def paired_gap(profits_a, profits_b):
    """Average of A - B year by year, with its 95 % half-width (needs CRN)."""
    mean, _, half = mean_ci(np.asarray(profits_a) - np.asarray(profits_b))
    return mean, half


def pick_full_service(results_by_case):
    """Selection rule of the report, applied to a grid of policies.

    Keep only the policies without a single lost sale in any demand case, then
    take the highest profit averaged over the cases; on a tie, the lower
    average standard deviation wins. Returns (s, Q, averaged mean profit).
    """
    cases = list(results_by_case.values())
    best = None
    for i, row in enumerate(cases[0]):
        same_policy = [case[i] for case in cases]
        if any(r["prob_lost_sales"] > 0 for r in same_policy):
            continue
        score = (np.mean([r["mean_profit"] for r in same_policy]),
                 -np.mean([r["std_profit"] for r in same_policy]))
        if best is None or score > best[0]:
            best = (score, row["s"], row["Q"])
    return best[1], best[2], float(best[0][0])


def full_service_holds(table):
    """Check question 5 on the grid: s >= 7 and Q >= 7 never lose a sale."""
    covered = [r for r in table if r["s"] >= 7 and r["Q"] >= 7]
    return len(covered), all(r["prob_lost_sales"] == 0 for r in covered)


def check_statement_example(calendar):
    """Replay the 3-day example of the assignment and stop if any profit differs."""
    run = simulate_year([4, 7, 5], calendar[:3], s=6, q=10, keep_log=True)
    expected = [-4_320_640, 4_079_760, -2_640_600]
    got = [round(day["profit"]) for day in run["log"]]
    if got != expected:
        raise AssertionError(f"The model does not match the statement: {got}")
    return run["log"]


def means_by_seed(demand_fn, calendar, s, q, seeds):
    """Mean annual profit of 100 years for each seed in `seeds`."""
    return np.array([run_experiment(demand_fn, np.random.default_rng(seed), calendar,
                                    s, q)["mean_profit"] for seed in seeds])


# =============================================================================
# Console output
# =============================================================================
def show_statistics(name, outcome):
    """Print the four statistics of the assignment plus their precision."""
    mean, se, half = mean_ci(outcome["profits"])
    r, slope, r2 = leftover_regression(outcome)
    left = outcome["ending_inventory"]
    print(f"--- Demanda {name} | s={REORDER_POINT}, Q={ORDER_QTY}, "
          f"{N_REPLICATIONS} réplicas ---")
    print(f"Ganancia anual promedio        {mean:>15,.0f} COP")
    print(f"Desviación estándar muestral   {outcome['std_profit']:>15,.0f} COP")
    print(f"Error estándar de la media     {se:>15,.0f} COP")
    print(f"IC 95 % de la media            [{mean - half:,.0f} ; {mean + half:,.0f}]")
    print(f"P(día con ventas perdidas)     {outcome['prob_lost_sales']:.4f}")
    print(f"Ventas perdidas por día        {outcome['mean_lost_per_day']:.4f} unidades")
    print(f"Inventario al 31/12            promedio {left.mean():.2f}, "
          f"de {left.min():.0f} a {left.max():.0f} unidades")
    print(f"Ganancia vs. inventario 31/12  r = {r:.3f}, R^2 = {r2:.3f}, "
          f"{slope:,.0f} COP por unidad sobrante\n")


def show_grid(name, table, top=15):
    """Print the best policies of the grid, highest mean profit first."""
    print(f"--- {len(table)} políticas, demanda {name} (números aleatorios comunes) ---")
    print(f"{'s':>3} {'Q':>4} {'Ganancia media':>16} {'Desv. est.':>13} "
          f"{'P(pérdida)':>11} {'Perd./día':>10}")
    for r in sorted(table, key=lambda row: row["mean_profit"], reverse=True)[:top]:
        print(f"{r['s']:>3} {r['Q']:>4} {r['mean_profit']:>16,.0f} "
              f"{r['std_profit']:>13,.0f} {r['prob_lost_sales']:>11.4f} "
              f"{r['mean_lost_per_day']:>10.4f}")
    print()


# =============================================================================
# B8 · Main program
# =============================================================================
def main():
    """Verification, both demand cases, seeds, policy grid and close calls."""
    clock = time.perf_counter()
    calendar = build_calendar(YEAR)
    if len(calendar) != 313:
        raise AssertionError("2026 should have 313 working days.")
    print(f"Días hábiles de {YEAR}: {len(calendar)} "
          f"({calendar[0]:%d/%m/%Y} a {calendar[-1]:%d/%m/%Y})")
    print(f"Política s = {REORDER_POINT}, Q = {ORDER_QTY} · {N_REPLICATIONS} réplicas · "
          f"semilla {SEED}\n")

    print("--- Ejemplo del enunciado: s=6, Q=10, demandas 4, 7, 5 ---")
    for d in check_statement_example(calendar):
        print(f"Día {d['day']} {d['date']:%d/%m/%Y}  R={d['received']:>2} "
              f"II={d['opening']:>2} D={d['demand']} V={d['sold']} P={d['lost']} "
              f"IF={d['closing']:>2} Pedido={d['order']} Utilidad={d['profit']:>13,.0f}")
    print("Las tres utilidades coinciden con el enunciado.\n")

    cases = [("Uniforme", demand_uniform), ("Empírica", demand_empirical)]

    # Question 1: statistics and histograms for the chosen policy
    for name, fn in cases:
        outcome = run_experiment(fn, np.random.default_rng(SEED), calendar,
                                 REORDER_POINT, ORDER_QTY)
        show_statistics(name, outcome)
        png = f"hist_{'uniforme' if name == 'Uniforme' else 'empirica'}.png"
        plot_histogram(outcome["profits"],
                       f"Ganancia anual, demanda {name.lower()} "
                       f"(s={REORDER_POINT}, Q={ORDER_QTY}, {N_REPLICATIONS} réplicas)", png)
        print(f"Histograma guardado en {png}\n")

    # Question 2: the same policy under five different seeds
    for name, fn in cases:
        means = means_by_seed(fn, calendar, REORDER_POINT, ORDER_QTY, SEED_SENSITIVITY)
        print(f"--- Cinco semillas, demanda {name} ---")
        for seed, m in zip(SEED_SENSITIVITY, means):
            print(f"Semilla {seed:>10}: ganancia media {m:,.0f} COP")
        print(f"Rango {means.max() - means.min():,.0f} COP · "
              f"desv. est. entre semillas {means.std(ddof=1):,.0f} COP\n")

    # Question 3: wide policy grid with common random numbers
    grid = {}
    for name, fn in cases:
        grid[name] = compare_policies(POLICIES, fn, calendar, N_REPLICATIONS, SEED)
        show_grid(name, grid[name])
        top = max(grid[name], key=lambda r: r["mean_profit"])
        print(f"Mayor ganancia sin restricción ({name}): s={top['s']}, Q={top['Q']}, "
              f"{top['mean_profit']:,.0f} COP\n")

    s_best, q_best, avg = pick_full_service(grid)
    print(f"Mejor con cero ventas perdidas en ambos casos: s={s_best}, Q={q_best} "
          f"({avg:,.0f} COP en promedio)")
    for name in grid:
        n, holds = full_service_holds(grid[name])
        print(f"s>=7 y Q>=7, demanda {name}: {n} políticas, "
              f"todas sin ventas perdidas = {holds}")
    print()

    for name in grid:
        by_policy = {(r["s"], r["Q"]): r for r in grid[name]}
        chosen = by_policy[(REORDER_POINT, ORDER_QTY)]["profits"]
        top = max(grid[name], key=lambda r: r["mean_profit"])
        for label, key in (("mejor sin restricción", (top["s"], top["Q"])),
                           ("mejor con cero pérdidas", (s_best, q_best))):
            gap, half = paired_gap(by_policy[key]["profits"], chosen)
            print(f"Demanda {name}, {label}: (s={key[0]}, Q={key[1]}) - "
                  f"(s={REORDER_POINT}, Q={ORDER_QTY}) = {gap:>10,.0f} ± {half:,.0f} COP "
                  f"(IC 95 % pareado)")
    print()

    # Close calls: s = 7 with Q around 15, 1000 years per policy
    for name, fn in cases:
        rows = compare_policies(PRECISION_POLICIES, fn, calendar, PRECISION_REPS, SEED + 1)
        chosen = next(r for r in rows if r["Q"] == ORDER_QTY)["profits"]
        print(f"--- s=7 con {PRECISION_REPS} réplicas por política, demanda {name} ---")
        for r in rows:
            gap, half = paired_gap(r["profits"], chosen)
            print(f"Q={r['Q']:>2}: media {r['mean_profit']:>14,.0f} COP | "
                  f"frente a Q={ORDER_QTY}: {gap:>10,.0f} ± {half:>8,.0f} COP")
        print()

    # Cost of full service: what a lower s gains and what it gives up
    for name, fn in cases:
        base = compare_policies([(REORDER_POINT, ORDER_QTY)], fn, calendar,
                                PRECISION_REPS, SEED + 2)[0]
        print(f"--- Costo del servicio completo, {PRECISION_REPS} réplicas, "
              f"demanda {name} ---")
        for r in compare_policies(SERVICE_POLICIES, fn, calendar, PRECISION_REPS, SEED + 2):
            gap, half = paired_gap(r["profits"], base["profits"])
            print(f"(s={r['s']}, Q={r['Q']}) - (s={REORDER_POINT}, Q={ORDER_QTY}) = "
                  f"{gap:>10,.0f} ± {half:,.0f} COP | "
                  f"P(pérdida) = {r['prob_lost_sales']:.4f}")
        print()

    print(f"Tiempo total de cómputo: {time.perf_counter() - clock:.1f} s")


if __name__ == "__main__":
    main()
