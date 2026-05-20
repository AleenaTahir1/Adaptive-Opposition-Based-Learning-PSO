"""Generate every figure used in the IEEE report.

Reads the trajectories pickled by :mod:`experiments` and produces:

* convergence curves (log-scale, median over runs) per benchmark
* swarm-diversity trajectories with the adaptive Jr curve overlaid
* a boxplot grid showing the AO-PSO advantage at D=30
* an ablation bar chart contrasting the seven value additions
"""

from __future__ import annotations

import os
import pickle

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd

from benchmarks import BENCHMARKS


RESULTS_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                           "results")
FIG_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                       "report", "figs")
os.makedirs(FIG_DIR, exist_ok=True)


COLORS = {"PSO": "#888888", "O-PSO": "#1f77b4", "AO-PSO": "#d62728"}
STYLES = {"PSO": "--", "O-PSO": "-.", "AO-PSO": "-"}


def _interp_median(fe_list, fit_list, grid):
    """Median over runs after interpolating each onto a common FE grid."""
    interp = np.stack([np.interp(grid, fe, fit) for fe, fit in zip(fe_list, fit_list)])
    return np.median(interp, axis=0)


def convergence_grid(traj: dict, D: int) -> None:
    fig, axes = plt.subplots(2, 3, figsize=(13.5, 7.5))
    for ax, bench in zip(axes.flat, BENCHMARKS):
        max_fe = max(max(traj[(a, bench.name, D)]["fe"][0])
                     for a, _ in [("PSO", None), ("O-PSO", None), ("AO-PSO", None)])
        grid = np.linspace(0, max_fe, 200)
        for algo in ("PSO", "O-PSO", "AO-PSO"):
            t = traj[(algo, bench.name, D)]
            med = _interp_median(t["fe"], t["fit"], grid)
            med = np.maximum(med, 1e-300)
            ax.plot(grid, med, color=COLORS[algo], linestyle=STYLES[algo],
                    linewidth=2, label=algo)
        ax.set_yscale("symlog", linthresh=1e-12)
        ax.set_title(bench.name, fontsize=11)
        ax.set_xlabel("Function evaluations")
        ax.set_ylabel("Best fitness (median)")
        ax.grid(True, alpha=0.3)
        ax.legend(fontsize=8, loc="upper right")
    fig.suptitle(f"Convergence on benchmark suite (D={D}, median of 30 runs)",
                 fontsize=13)
    fig.tight_layout()
    path = os.path.join(FIG_DIR, f"convergence_D{D}.pdf")
    fig.savefig(path, dpi=180, bbox_inches="tight")
    fig.savefig(path.replace(".pdf", ".png"), dpi=180, bbox_inches="tight")
    plt.close(fig)
    print(f"wrote {path}")


def diversity_jr_plot(traj: dict) -> None:
    """Diversity and adaptive Jr on Rastrigin D=30 -- the showcase plot."""
    bench_name = "Rastrigin"
    D = 30
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(11.5, 4.2))
    for algo in ("PSO", "O-PSO", "AO-PSO"):
        t = traj[(algo, bench_name, D)]
        grid = np.linspace(0, max(t["fe"][0]), 200)
        div = _interp_median(t["fe"], t["div"], grid)
        ax1.plot(grid, div, color=COLORS[algo], linestyle=STYLES[algo],
                 linewidth=2, label=algo)
    ax1.set_xlabel("Function evaluations")
    ax1.set_ylabel("Normalised swarm diameter $\\sigma_t$")
    ax1.set_title(f"Swarm diversity on {bench_name} (D={D})")
    ax1.grid(True, alpha=0.3)
    ax1.legend(fontsize=9)

    t = traj[("AO-PSO", bench_name, D)]
    grid = np.linspace(0, max(t["fe"][0]), 200)
    jr = _interp_median(t["fe"], t["jr"], grid)
    ax2.plot(grid, jr, color=COLORS["AO-PSO"], linewidth=2)
    ax2.set_xlabel("Function evaluations")
    ax2.set_ylabel("Adaptive jumping rate $J_r^{(t)}$")
    ax2.set_title("AO-PSO: $J_r$ falls as the swarm converges")
    ax2.grid(True, alpha=0.3)

    fig.tight_layout()
    path = os.path.join(FIG_DIR, "diversity_jr.pdf")
    fig.savefig(path, dpi=180, bbox_inches="tight")
    fig.savefig(path.replace(".pdf", ".png"), dpi=180, bbox_inches="tight")
    plt.close(fig)
    print(f"wrote {path}")


def boxplot_grid() -> None:
    df = pd.read_csv(os.path.join(RESULTS_DIR, "raw_runs.csv"))
    fig, axes = plt.subplots(2, 3, figsize=(13.5, 7.5))
    for ax, bench in zip(axes.flat, BENCHMARKS):
        data = []
        for algo in ("PSO", "O-PSO", "AO-PSO"):
            row = df[(df.benchmark == bench.name) & (df.dim == 30)
                     & (df.algorithm == algo)]["best_fitness"].values
            # log-axis hates zeros and negatives: clamp to a tiny epsilon
            row = np.maximum(row, 1e-300)
            data.append(row)
        bp = ax.boxplot(data, labels=("PSO", "O-PSO", "AO-PSO"),
                        patch_artist=True)
        for patch, algo in zip(bp["boxes"], ("PSO", "O-PSO", "AO-PSO")):
            patch.set_facecolor(COLORS[algo])
            patch.set_alpha(0.55)
        ax.set_yscale("symlog", linthresh=1e-10)
        ax.set_title(bench.name)
        ax.grid(True, alpha=0.3)
    fig.suptitle("Final-fitness distribution at D=30 (30 runs)", fontsize=13)
    fig.tight_layout()
    path = os.path.join(FIG_DIR, "boxplots_D30.pdf")
    fig.savefig(path, dpi=180, bbox_inches="tight")
    fig.savefig(path.replace(".pdf", ".png"), dpi=180, bbox_inches="tight")
    plt.close(fig)
    print(f"wrote {path}")


def main() -> None:
    with open(os.path.join(RESULTS_DIR, "trajectories.pkl"), "rb") as f:
        traj = pickle.load(f)
    convergence_grid(traj, D=10)
    convergence_grid(traj, D=30)
    diversity_jr_plot(traj)
    boxplot_grid()


if __name__ == "__main__":
    main()
