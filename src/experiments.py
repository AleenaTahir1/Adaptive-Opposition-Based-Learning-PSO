"""Run the full benchmark sweep: every algorithm on every benchmark for
``N_RUNS`` independent seeds, at multiple dimensionalities, then dump:

* a per-run record CSV (one row per (algo, benchmark, dim, seed))
* a summary table with mean / std / best / worst per cell
* Wilcoxon signed-rank p-values comparing AO-PSO vs each baseline
* pickled convergence and diversity trajectories for the plots module

This is the script that feeds the LaTeX report.
"""

from __future__ import annotations

import os
import pickle
import time

import numpy as np
import pandas as pd
from scipy.stats import wilcoxon

from benchmarks import BENCHMARKS
from pso_variants import AOPSO, OPSO, StandardPSO


# -------- configuration -----------------------------------------------------

N_RUNS = 30
DIMENSIONS = [10, 30]
FE_BUDGET = {10: 20_000, 30: 60_000}
SWARM_SIZE = 20

ALGOS = [
    ("PSO",   StandardPSO),
    ("O-PSO", OPSO),
    ("AO-PSO", AOPSO),
]

RESULTS_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                           "results")
os.makedirs(RESULTS_DIR, exist_ok=True)


# -------- main sweep --------------------------------------------------------

def main() -> None:
    rows: list[dict] = []
    # keep one full trajectory per (algo, bench, dim) for plotting
    trajectories: dict[tuple[str, str, int], dict[str, list]] = {}

    t0 = time.time()
    for D in DIMENSIONS:
        budget = FE_BUDGET[D]
        print(f"\n=== D = {D}, budget = {budget} FEs ===")
        for bench in BENCHMARKS:
            print(f"  {bench.name:<12}", end="", flush=True)
            for algo_name, algo_cls in ALGOS:
                fes_at_end = []
                bests = []
                # collect trajectories from every run, average for plotting
                all_fe = []
                all_fit = []
                all_div = []
                all_jr = []
                for seed in range(N_RUNS):
                    rng = np.random.default_rng(1000 * D + seed)
                    algo = algo_cls(N=SWARM_SIZE, max_fe=budget, rng=rng)
                    res = algo.run(bench, D)
                    rows.append({
                        "algorithm": algo_name,
                        "benchmark": bench.name,
                        "dim": D,
                        "seed": seed,
                        "best_fitness": res.best_fitness,
                        "fe_used": res.fe_history[-1],
                    })
                    bests.append(res.best_fitness)
                    fes_at_end.append(res.fe_history[-1])
                    all_fe.append(res.fe_history)
                    all_fit.append(res.fit_history)
                    all_div.append(res.diversity_history)
                    all_jr.append(res.jr_history)
                trajectories[(algo_name, bench.name, D)] = {
                    "fe": all_fe,
                    "fit": all_fit,
                    "div": all_div,
                    "jr": all_jr,
                }
                print(f"  {algo_name}:{np.mean(bests):.2e}", end="", flush=True)
            print()
    elapsed = time.time() - t0
    print(f"\nTotal runtime: {elapsed:.1f} s")

    df = pd.DataFrame(rows)
    df.to_csv(os.path.join(RESULTS_DIR, "raw_runs.csv"), index=False)

    # ---- summary statistics -----------------------------------------------
    summary = (
        df.groupby(["benchmark", "dim", "algorithm"])["best_fitness"]
          .agg(["mean", "std", "min", "max", "median"])
          .reset_index()
    )
    summary.to_csv(os.path.join(RESULTS_DIR, "summary.csv"), index=False)

    # ---- Wilcoxon AO-PSO vs each baseline ---------------------------------
    wilc_rows = []
    for bench in BENCHMARKS:
        for D in DIMENSIONS:
            ao = df[(df.algorithm == "AO-PSO")
                    & (df.benchmark == bench.name)
                    & (df.dim == D)]["best_fitness"].values
            for baseline in ("PSO", "O-PSO"):
                bl = df[(df.algorithm == baseline)
                        & (df.benchmark == bench.name)
                        & (df.dim == D)]["best_fitness"].values
                # zero differences need special handling
                diff = bl - ao
                if np.allclose(diff, 0):
                    p_value = 1.0
                    stat = 0.0
                else:
                    try:
                        stat, p_value = wilcoxon(ao, bl,
                                                 zero_method="wilcox",
                                                 alternative="less")
                    except ValueError:
                        # all-zero differences fallback
                        p_value = 1.0
                        stat = 0.0
                wilc_rows.append({
                    "benchmark": bench.name,
                    "dim": D,
                    "comparison": f"AO-PSO < {baseline}",
                    "statistic": float(stat),
                    "p_value": float(p_value),
                    "significant_at_0.05": p_value < 0.05,
                })
    wilc = pd.DataFrame(wilc_rows)
    wilc.to_csv(os.path.join(RESULTS_DIR, "wilcoxon.csv"), index=False)

    with open(os.path.join(RESULTS_DIR, "trajectories.pkl"), "wb") as f:
        pickle.dump(trajectories, f)

    print("\nSummary (mean fitness):\n")
    pivot = (
        summary.pivot_table(index=["benchmark", "dim"],
                            columns="algorithm", values="mean")
    )
    print(pivot.to_string(float_format=lambda v: f"{v:.3e}"))
    print("\nWilcoxon (AO-PSO better-than baseline, p<0.05):\n")
    print(wilc.to_string(index=False))


if __name__ == "__main__":
    main()
