"""Opposition-Based Learning primitives shared by the PSO variants.

Implements the three OBL building blocks used by AO-PSO:

* Tizhoosh's standard linear opposite (used by the focal O-PSO paper).
* Wang et al.'s Generalized OBL (GOBL) -- value-addition #2.
* The logistic chaotic map for seeding -- value-addition #1.

Plus a helper that returns the swarm's normalized diameter so the
adaptive jumping rate (value-addition #4) can react to live diversity.
"""

from __future__ import annotations

import numpy as np


def standard_opposite(X: np.ndarray, low: np.ndarray, high: np.ndarray) -> np.ndarray:
    """Tizhoosh (2005): xhat = a + b - x, applied component-wise."""
    return low + high - X


def generalized_opposite(
    X: np.ndarray,
    low: np.ndarray,
    high: np.ndarray,
    rng: np.random.Generator,
) -> np.ndarray:
    """Wang et al. GOBL:

        xhat_j = a_j + b_j - beta * x_j - (1 - beta) * x_bar_j

    where ``beta ~ U(0,1)`` (one beta per particle, broadcast across
    dimensions) and ``x_bar`` is the per-dimension population mean.

    Result is clipped to the [low, high] box because the formula can
    occasionally drift outside the search domain.
    """
    x_bar = X.mean(axis=0)  # (D,)
    beta = rng.random((X.shape[0], 1))  # one beta per particle
    Xhat = low + high - beta * X - (1.0 - beta) * x_bar
    return np.clip(Xhat, low, high)


def logistic_chaotic_init(
    N: int,
    D: int,
    low: np.ndarray,
    high: np.ndarray,
    rng: np.random.Generator,
    mu: float = 4.0,
    burn_in: int = 200,
) -> np.ndarray:
    """Generate an N x D initial population from a logistic chaotic map.

    Uses ``x_{n+1} = mu * x_n * (1 - x_n)`` with ``mu = 4`` (fully chaotic
    regime). A small burn-in is discarded to avoid early transients. The
    resulting [0,1] sequence is linearly mapped to the search box.
    """
    # one independent chaotic stream per dimension reduces axis-aligned
    # artefacts that you get from reshaping a single long sequence.
    seeds = rng.uniform(0.05, 0.95, size=D)
    # avoid fixed points (0, 0.25, 0.5, 0.75, 1) of the logistic map.
    bad = np.isclose(seeds, 0.0) | np.isclose(seeds, 0.25) | \
        np.isclose(seeds, 0.5) | np.isclose(seeds, 0.75) | np.isclose(seeds, 1.0)
    seeds[bad] += 0.013

    series = np.empty((N + burn_in, D))
    series[0] = seeds
    for n in range(N + burn_in - 1):
        series[n + 1] = mu * series[n] * (1.0 - series[n])

    unit = series[burn_in:]                       # (N, D) in [0,1]
    return low + unit * (high - low)              # scaled into the search box


def swarm_diameter(X: np.ndarray, low: np.ndarray, high: np.ndarray) -> float:
    """Diameter measure normalised to [0,1] for use in the adaptive Jr.

    We use the mean per-dimension standard deviation, divided by the
    per-dimension box width. This is cheaper than the all-pairs maximum
    distance and is well-behaved for the adaptive jumping schedule.
    """
    std_per_dim = X.std(axis=0)
    width = high - low
    return float(np.mean(std_per_dim / np.where(width > 0, width, 1.0)))
