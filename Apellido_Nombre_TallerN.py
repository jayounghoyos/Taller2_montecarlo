"""Taller 2 - Simulación Monte Carlo de un sistema de inventarios.

Parte B: implementación en Python.

Requisitos: Python 3, NumPy y Matplotlib.
Ejecución: python Apellido_Nombre_TallerN.py

La semilla 12345 permite reproducir los resultados. Si el curso exige usar el
código de estudiante como semilla, cambie SEED por ese valor antes de entregar.
"""

import datetime as dt
from pathlib import Path

import matplotlib.pyplot as plt
import numpy as np

# -----------------------------------------------------------------------------
# Parameters
# -----------------------------------------------------------------------------
YEAR = 2026
N_REPLICATIONS = 100
REORDER_POINT = 7       # s: place an order when final inventory is strictly below s
ORDER_QTY = 15          # Q: fixed order quantity and initial inventory

PRICE = 680_000         # revenue per unit sold (COP)
UNIT_COST = 600_000     # purchase cost per unit received (COP)
ORDER_COST = 40_000     # fixed cost per received order (COP)
HOLDING_COST = 80       # cost per unit per day, applied to average daily inventory
FIXED_MONTHLY = 1_000_000  # charged on the first working day of each month (COP)

SEED = 12345
EMP_VALUES = (3, 4, 5, 6, 7)
EMP_PROBS = (0.20, 0.05, 0.40, 0.25, 0.10)

# More than the 20 policy combinations required by the assignment (30 total).
POLICIES = [(s, q) for s in (5, 6, 7, 8, 9) for q in (7, 10, 15, 20, 30, 50)]
SEED_SENSITIVITY = (12345, 2026, 42, 777, 2025)
OUTPUT_DIR = Path(__file__).resolve().parent


def build_calendar(year):
    """Return every date from Monday to Saturday in the selected year."""
    days = []
    day = dt.date(year, 1, 1)
    while day.year == year:
        if day.weekday() != 6:  # Python weekday(): Monday=0, Sunday=6
            days.append(day)
        day += dt.timedelta(days=1)
    return days


def demand_uniform(rng):
    """Generate discrete uniform demand on {3, 4, 5, 6, 7}."""
    return int(rng.integers(3, 8))  # upper limit is excluded


def demand_empirical(rng):
    """Generate empirical demand using the inverse-transform method."""
    u = rng.random()
    cumulative = 0.0
    for value, probability in zip(EMP_VALUES, EMP_PROBS):
        cumulative += probability
        if u < cumulative:
            return value
    return EMP_VALUES[-1]  # protects against floating-point rounding


def simulate_year(demands, calendar, s, q, keep_log=False):
    """Simulate one year for a given demand sequence and (s, Q) policy.

    An order placed at the end of a working day arrives at the beginning of
    the next working day. The initial inventory on day 1 is Q.
    """
    if len(demands) != len(calendar):
        raise ValueError("The demand sequence and calendar must have equal lengths.")
    if s < 0 or q <= 0:
        raise ValueError("s must be non-negative and Q must be positive.")

    annual_profit = 0.0
    days_with_shortage = 0
    total_lost_units = 0
    final_inventory = 0
    order_pending = False
    log = []

    for t, day in enumerate(calendar):
        # Receive the initial order on day 1, or an order placed the prior day.
        received = q if (t == 0 or order_pending) else 0
        initial_inventory = final_inventory + received

        # Meet demand from available stock; unmet demand is lost.
        demand = int(demands[t])
        sales = min(initial_inventory, demand)
        lost_units = demand - sales
        final_inventory = initial_inventory - sales

        # Review inventory at the end of the day.
        order_pending = final_inventory < s

        # Daily revenues and costs.
        month_start = (t == 0) or (calendar[t - 1].month != day.month)
        revenue = PRICE * sales
        purchase_cost = UNIT_COST * received
        ordering_cost = ORDER_COST if received > 0 else 0
        holding_cost = HOLDING_COST * (initial_inventory + final_inventory) / 2
        fixed_cost = FIXED_MONTHLY if month_start else 0
        daily_profit = revenue - purchase_cost - ordering_cost - holding_cost - fixed_cost
        annual_profit += daily_profit

        total_lost_units += lost_units
        days_with_shortage += int(lost_units > 0)

        if keep_log:
            log.append({
                "day": t + 1,
                "date": day,
                "received": received,
                "initial_inventory": initial_inventory,
                "demand": demand,
                "sales": sales,
                "lost_units": lost_units,
                "final_inventory": final_inventory,
                "order_at_close": int(order_pending),
                "revenue": revenue,
                "purchase_cost": purchase_cost,
                "ordering_cost": ordering_cost,
                "holding_cost": holding_cost,
                "fixed_cost": fixed_cost,
                "daily_profit": daily_profit,
            })

    n_days = len(calendar)
    return {
        "profit": annual_profit,
        "shortage_fraction": days_with_shortage / n_days,
        "lost_per_day": total_lost_units / n_days,
        "total_lost_units": total_lost_units,
        "days_with_shortage": days_with_shortage,
        "ending_inventory": final_inventory,
        "log": log,
    }


def run_experiment(demand_fn, rng, calendar, s, q, n_reps=N_REPLICATIONS):
    """Run independent annual replications and calculate four performance metrics."""
    profits = np.zeros(n_reps, dtype=float)
    shortage_fraction = np.zeros(n_reps, dtype=float)
    lost_per_day = np.zeros(n_reps, dtype=float)
    ending_inventory = np.zeros(n_reps, dtype=float)

    for r in range(n_reps):
        demands = [demand_fn(rng) for _ in calendar]
        result = simulate_year(demands, calendar, s, q)
        profits[r] = result["profit"]
        shortage_fraction[r] = result["shortage_fraction"]
        lost_per_day[r] = result["lost_per_day"]
        ending_inventory[r] = result["ending_inventory"]

    return {
        "profits": profits,
        "mean_profit": float(profits.mean()),
        "std_profit": float(profits.std(ddof=1)),  # sample standard deviation
        "prob_lost_sales": float(shortage_fraction.mean()),
        "mean_lost_per_day": float(lost_per_day.mean()),
        "mean_ending_inventory": float(ending_inventory.mean()),
    }


def plot_histogram(profits, title, filename):
    """Save a histogram with ten equal-width bins, expressed in millions of COP."""
    fig, ax = plt.subplots(figsize=(9, 5.2))
    ax.hist(profits / 1_000_000, bins=10, edgecolor="black")
    ax.set_title(title)
    ax.set_xlabel("Ganancia anual (millones de COP)")
    ax.set_ylabel("Frecuencia de réplicas")
    ax.grid(axis="y", alpha=0.25)
    fig.tight_layout()
    fig.savefig(OUTPUT_DIR / filename, dpi=180)
    plt.close(fig)


def compare_policies(policies, demand_fn, calendar, n_reps, seed):
    """Compare policies using common random numbers (the same seed per policy)."""
    rows = []
    for s, q in policies:
        # Resetting the generator makes each policy face the same demand sequences.
        rng = np.random.default_rng(seed)
        result = run_experiment(demand_fn, rng, calendar, s, q, n_reps)
        rows.append({
            "s": s,
            "Q": q,
            "mean_profit": result["mean_profit"],
            "std_profit": result["std_profit"],
            "prob_lost_sales": result["prob_lost_sales"],
            "mean_lost_per_day": result["mean_lost_per_day"],
        })
    return rows


def verify_required_example(calendar):
    """Check the three required daily profits from the workshop statement."""
    result = simulate_year([4, 7, 5], calendar[:3], s=6, q=10, keep_log=True)
    expected = [-4_320_640, 4_079_760, -2_640_600]
    observed = [round(row["daily_profit"]) for row in result["log"]]
    if observed != expected:
        raise AssertionError(f"Verification failed: {observed} != {expected}")
    return result["log"]


def seed_sensitivity(demand_fn, calendar, s, q, seeds):
    """Measure the variation in mean annual profit across five random seeds."""
    results = []
    for seed in seeds:
        rng = np.random.default_rng(seed)
        result = run_experiment(demand_fn, rng, calendar, s, q, N_REPLICATIONS)
        results.append({"seed": seed, "mean_profit": result["mean_profit"]})
    values = np.array([row["mean_profit"] for row in results])
    return results, float(values.min()), float(values.max()), float(values.max() - values.min())


def print_experiment_summary(name, result):
    """Print the required performance statistics in a readable format."""
    print(f"\n--- Demanda {name} | política s={REORDER_POINT}, Q={ORDER_QTY} ---")
    print(f"Ganancia anual promedio       : ${result['mean_profit']:,.0f} COP")
    print(f"Desviación estándar muestral  : ${result['std_profit']:,.0f} COP")
    print(f"Probabilidad de ventas perdidas: {result['prob_lost_sales']:.4f}")
    print(f"Unidades perdidas promedio/día: {result['mean_lost_per_day']:.4f}")
    print(f"Inventario final promedio     : {result['mean_ending_inventory']:.2f} unidades")


def print_policy_table(name, rows):
    """Print policy results ordered by mean annual profit, from highest to lowest."""
    print(f"\n--- Comparación de políticas: demanda {name} ---")
    print(f"{'s':>3} {'Q':>4} {'Ganancia media':>17} {'Desv. estándar':>17} "
          f"{'Pérdida/día':>13} {'Unid. perdidas/día':>20}")
    for row in sorted(rows, key=lambda item: item["mean_profit"], reverse=True):
        print(f"{row['s']:>3} {row['Q']:>4} {row['mean_profit']:>17,.0f} "
              f"{row['std_profit']:>17,.0f} {row['prob_lost_sales']:>13.4f} "
              f"{row['mean_lost_per_day']:>20.4f}")


def main():
    calendar = build_calendar(YEAR)
    print("TALLER 2 - SIMULACIÓN MONTE CARLO DE INVENTARIOS")
    print(f"Año simulado: {YEAR} | Días hábiles (lunes a sábado): {len(calendar)}")
    print(f"Réplicas: {N_REPLICATIONS} | Semilla base: {SEED}")
    print(f"Política principal: s={REORDER_POINT}, Q={ORDER_QTY}")
    if len(calendar) != 313:
        raise AssertionError("The calendar should contain exactly 313 working days.")

    # Mandatory verification case from the assignment.
    print("\n--- Prueba obligatoria (s=6, Q=10; demandas 4, 7, 5) ---")
    for row in verify_required_example(calendar):
        print(f"Día {row['day']} | {row['date']:%a %d/%m/%Y} | "
              f"Inventario inicial={row['initial_inventory']:>2} | "
              f"Demanda={row['demand']} | Ventas={row['sales']} | "
              f"Ventas perdidas={row['lost_units']} | "
              f"Inventario final={row['final_inventory']:>2} | "
              f"Utilidad=${row['daily_profit']:,.0f}")

    cases = [("Uniforme", demand_uniform), ("Empírica", demand_empirical)]
    summary = {}
    for name, demand_fn in cases:
        rng = np.random.default_rng(SEED)
        result = run_experiment(demand_fn, rng, calendar,
                                REORDER_POINT, ORDER_QTY, N_REPLICATIONS)
        summary[name] = result
        print_experiment_summary(name, result)
        plot_histogram(result["profits"],
                       f"Distribución de la ganancia anual - Demanda {name.lower()}",
                       f"histograma_{name.lower()}.png")
        print(f"Histograma guardado en: histograma_{name.lower()}.png")

    # Sensitivity to random seeds, as required in Part B analysis question 2.
    for name, demand_fn in cases:
        seed_rows, minimum, maximum, spread = seed_sensitivity(
            demand_fn, calendar, REORDER_POINT, ORDER_QTY, SEED_SENSITIVITY
        )
        print(f"\n--- Sensibilidad a cinco semillas: demanda {name} ---")
        for row in seed_rows:
            print(f"Semilla {row['seed']:>5}: ganancia media = ${row['mean_profit']:,.0f} COP")
        print(f"Rango de ganancias medias: ${minimum:,.0f} a ${maximum:,.0f} COP")
        print(f"Diferencia máximo - mínimo: ${spread:,.0f} COP")

    # Policy comparison with common random numbers: 30 (s, Q) combinations.
    print(f"\nSe comparan {len(POLICIES)} políticas (s, Q), usando números aleatorios comunes.")
    policy_results = {}
    for name, demand_fn in cases:
        rows = compare_policies(POLICIES, demand_fn, calendar, N_REPLICATIONS, SEED)
        policy_results[name] = rows
        print_policy_table(name, rows)
        best = max(rows, key=lambda item: item["mean_profit"])
        print(f"Mejor por ganancia media para demanda {name}: "
              f"s={best['s']}, Q={best['Q']}, media=${best['mean_profit']:,.0f} COP")

    print("\n--- Comprobación teórica de ventas perdidas ---")
    print("Si s >= 7 y Q >= 7, la simulación debe producir cero ventas perdidas,")
    print("porque la demanda máxima es 7 y el inventario recibido al inicio del día")
    print("siguiente es al menos Q cuando se activa el pedido. Verifique las filas")
    print("con s=7, 8 o 9 y Q>=7 en la tabla de políticas.")
    print("\nEjecución terminada correctamente.")


if __name__ == "__main__":
    main()
