"""Q1 reference solver and alignment checks.

This is an independent verification script. It uses attachment 1 only and
solves the deterministic daily storage model with a MILP charge/discharge
exclusivity formulation.

Note on the printed cost
------------------------
``Purchase cost`` printed below is the *lexicographic second-stage* cost
(~35126.9837). The primary acceptance anchor is the single-objective optimum

    C1* = 35126.948589 yuan

which the main LP in ``src/q1/run.py`` must reproduce. The gap (~0.035) is the
tolerance spent by lexicographic second-stage minimisation; it is expected and
does not indicate a modelling error. The data benchmarks and the six specified
periods printed here are the values to compare against.
"""

from __future__ import annotations

import argparse
from pathlib import Path

import numpy as np
import openpyxl
from scipy.optimize import Bounds, LinearConstraint, milp


DT = 1.0 / 6.0
POWER_LIMIT = 5000.0 / 6.0
ETA_C = 0.90
ETA_D = 0.90
SOC_MIN = 1200.0
SOC_MAX = 10800.0
SOC_INITIAL = 6000.0


def load_attachment(path: Path) -> tuple[np.ndarray, np.ndarray, np.ndarray]:
    """Read price, load and photovoltaic power and convert power to kWh."""
    workbook = openpyxl.load_workbook(path, read_only=True, data_only=True)
    rows = list(workbook["Sheet1"].iter_rows(min_row=2, values_only=True))
    workbook.close()

    price = np.array([float(row[1]) for row in rows])
    load = np.array([float(row[2]) for row in rows]) * DT
    pv = np.array([float(row[3]) for row in rows]) * DT
    if len(price) != 144:
        raise ValueError(f"Expected 144 ten-minute records, got {len(price)}")
    return price, load, pv


def solve_q1(price: np.ndarray, load: np.ndarray, pv: np.ndarray) -> dict[str, np.ndarray | float]:
    """Solve the daily Q1 storage model with terminal SOC equal to initial SOC."""
    n_periods = len(price)
    variables_per_period = 6  # grid, charge, discharge, curtailment, SOC, binary
    n_variables = variables_per_period * n_periods
    grid, charge, discharge, curtailment, soc, binary = range(6)

    def index(period: int, variable: int) -> int:
        return variables_per_period * period + variable

    cost = np.zeros(n_variables)
    for period in range(n_periods):
        cost[index(period, grid)] = price[period]

    rows: list[np.ndarray] = []
    lower: list[float] = []
    upper: list[float] = []

    # grid - charge + discharge - curtailment = load - photovoltaic
    for period in range(n_periods):
        row = np.zeros(n_variables)
        row[index(period, grid)] = 1.0
        row[index(period, charge)] = -1.0
        row[index(period, discharge)] = 1.0
        row[index(period, curtailment)] = -1.0
        rows.append(row)
        rhs = load[period] - pv[period]
        lower.append(rhs)
        upper.append(rhs)

    # SOC transition
    for period in range(n_periods):
        row = np.zeros(n_variables)
        row[index(period, soc)] = 1.0
        row[index(period, charge)] = -ETA_C
        row[index(period, discharge)] = 1.0 / ETA_D
        if period > 0:
            row[index(period - 1, soc)] = -1.0
        rows.append(row)
        rhs = SOC_INITIAL if period == 0 else 0.0
        lower.append(rhs)
        upper.append(rhs)

    # Charge/discharge mutual exclusion: c <= M z; d <= M (1-z).
    for period in range(n_periods):
        row = np.zeros(n_variables)
        row[index(period, charge)] = 1.0
        row[index(period, binary)] = -POWER_LIMIT
        rows.append(row)
        lower.append(-np.inf)
        upper.append(0.0)

        row = np.zeros(n_variables)
        row[index(period, discharge)] = 1.0
        row[index(period, binary)] = POWER_LIMIT
        rows.append(row)
        lower.append(-np.inf)
        upper.append(POWER_LIMIT)

    # Daily cycle requirement.
    row = np.zeros(n_variables)
    row[index(n_periods - 1, soc)] = 1.0
    rows.append(row)
    lower.append(SOC_INITIAL)
    upper.append(SOC_INITIAL)

    matrix = np.array(rows)
    lb = np.full(n_variables, -np.inf)
    ub = np.full(n_variables, np.inf)
    integrality = np.zeros(n_variables)
    for period in range(n_periods):
        lb[index(period, grid)] = 0.0
        lb[index(period, charge)] = 0.0
        ub[index(period, charge)] = POWER_LIMIT
        lb[index(period, discharge)] = 0.0
        ub[index(period, discharge)] = POWER_LIMIT
        lb[index(period, curtailment)] = 0.0
        ub[index(period, curtailment)] = pv[period]
        lb[index(period, soc)] = SOC_MIN
        ub[index(period, soc)] = SOC_MAX
        lb[index(period, binary)] = 0.0
        ub[index(period, binary)] = 1.0
        integrality[index(period, binary)] = 1.0

    first_stage = milp(
        cost,
        constraints=[LinearConstraint(matrix, np.array(lower), np.array(upper))],
        integrality=integrality,
        bounds=Bounds(lb, ub),
    )
    if not first_stage.success:
        raise RuntimeError(first_stage.message)

    # Lexicographic second stage: preserve cost, then minimize battery throughput.
    throughput = np.zeros(n_variables)
    for period in range(n_periods):
        throughput[index(period, charge)] = 1.0
        throughput[index(period, discharge)] = 1.0
    extended_matrix = np.vstack([matrix, cost[None, :]])
    extended_lower = np.append(lower, -np.inf)
    extended_upper = np.append(upper, (1.0 + 1e-6) * first_stage.fun)
    second_stage = milp(
        throughput,
        constraints=[LinearConstraint(extended_matrix, extended_lower, extended_upper)],
        integrality=integrality,
        bounds=Bounds(lb, ub),
    )
    if not second_stage.success:
        raise RuntimeError(second_stage.message)

    solution = second_stage.x
    return {
        "grid": np.array([solution[index(i, grid)] for i in range(n_periods)]),
        "charge": np.array([solution[index(i, charge)] for i in range(n_periods)]),
        "discharge": np.array([solution[index(i, discharge)] for i in range(n_periods)]),
        "curtailment": np.array([solution[index(i, curtailment)] for i in range(n_periods)]),
        "soc": np.array([solution[index(i, soc)] for i in range(n_periods)]),
        "first_stage_cost": float(first_stage.fun),
    }


def report(price: np.ndarray, load: np.ndarray, pv: np.ndarray, result: dict[str, np.ndarray | float]) -> None:
    """Print independent Q1 reconciliation metrics."""
    grid = result["grid"]
    charge = result["charge"]
    discharge = result["discharge"]
    curtailment = result["curtailment"]
    soc = result["soc"]
    assert isinstance(grid, np.ndarray)
    assert isinstance(charge, np.ndarray)
    assert isinstance(discharge, np.ndarray)
    assert isinstance(curtailment, np.ndarray)
    assert isinstance(soc, np.ndarray)

    gap = np.maximum(0.0, load - pv)
    base_cost = float(np.sum(price * gap))
    solved_cost = float(np.sum(price * grid))

    print("=== Attachment-1 benchmarks ===")
    print(f"Load total:                 {load.sum():12.2f} kWh")
    print(f"PV total:                   {pv.sum():12.2f} kWh")
    print(f"No-storage purchase cost:   {base_cost:12.4f} yuan")
    print()
    print("=== Q1 reference MILP ===")
    print(f"Purchase cost:              {solved_cost:12.4f} yuan")
    print(f"Grid purchase total:        {grid.sum():12.4f} kWh")
    print(f"Charge / discharge total:   {charge.sum():12.4f} / {discharge.sum():.4f} kWh")
    print(f"Curtailment total:          {curtailment.sum():12.4f} kWh")
    print(f"SOC minimum / maximum:      {soc.min():.2f} / {soc.max():.2f} kWh")
    print(f"Terminal SOC:               {soc[-1]:12.4f} kWh")
    print(f"Efficiency identity d=.81c: {discharge.sum():.4f} vs {(0.81 * charge.sum()):.4f}")
    print()
    print("Specified purchase periods (model t = 61, 73, 85, 97, 109, 121):")
    for period in [61, 73, 85, 97, 109, 121]:
        print(f"  t={period:3d}: {grid[period - 1]:8.2f} kWh")


def main() -> None:
    parser = argparse.ArgumentParser(description="Solve and verify the Q1 reference model.")
    default_input = Path(__file__).resolve().parents[1] / "data" / "附件1.xlsx"
    parser.add_argument("--input", type=Path, default=default_input, help="Path to 附件1.xlsx")
    args = parser.parse_args()
    price, load, pv = load_attachment(args.input)
    report(price, load, pv, solve_q1(price, load, pv))


if __name__ == "__main__":
    main()
