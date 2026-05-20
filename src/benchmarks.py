"""Benchmark functions used to evaluate AO-PSO.

The first four (Sphere, Rosenbrock, Rastrigin, Griewank) are the exact
benchmarks reported by Jabeen, Jalil, and Baig (GECCO 2009) so we can
compete head-to-head with the focal paper's published numbers.

Two additional well-known multimodal functions (Ackley, Schwefel) are
included to demonstrate robustness beyond the focal paper's footprint.

Each Benchmark exposes:
    name      - human-readable label
    func(x)   - scalar objective; x is a (D,) numpy array
    bounds    - (low, high) per-dimension box; symmetric here
    optimum   - known global minimum value (typically 0)
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Callable

import numpy as np


@dataclass
class Benchmark:
    name: str
    func: Callable[[np.ndarray], float]
    low: float
    high: float
    optimum: float = 0.0

    def bounds(self, D: int) -> tuple[np.ndarray, np.ndarray]:
        return np.full(D, self.low), np.full(D, self.high)


def sphere(x: np.ndarray) -> float:
    return float(np.sum(x * x))


def rosenbrock(x: np.ndarray) -> float:
    return float(np.sum(100.0 * (x[1:] - x[:-1] ** 2) ** 2 + (x[:-1] - 1.0) ** 2))


def rastrigin(x: np.ndarray) -> float:
    A = 10.0
    return float(A * x.size + np.sum(x * x - A * np.cos(2.0 * np.pi * x)))


def griewank(x: np.ndarray) -> float:
    s = np.sum(x * x) / 4000.0
    p = np.prod(np.cos(x / np.sqrt(np.arange(1, x.size + 1))))
    return float(s - p + 1.0)


def ackley(x: np.ndarray) -> float:
    D = x.size
    s1 = np.sum(x * x)
    s2 = np.sum(np.cos(2.0 * np.pi * x))
    return float(-20.0 * np.exp(-0.2 * np.sqrt(s1 / D))
                 - np.exp(s2 / D) + 20.0 + np.e)


def schwefel(x: np.ndarray) -> float:
    # Shifted so the minimum lies at the origin = 0.
    D = x.size
    return float(418.9828872724337 * D - np.sum(x * np.sin(np.sqrt(np.abs(x)))))


BENCHMARKS: list[Benchmark] = [
    Benchmark("Sphere",     sphere,     -100.0, 100.0, 0.0),
    Benchmark("Rosenbrock", rosenbrock,  -30.0,  30.0, 0.0),
    Benchmark("Rastrigin",  rastrigin,    -5.12, 5.12, 0.0),
    Benchmark("Griewank",   griewank,   -600.0, 600.0, 0.0),
    Benchmark("Ackley",     ackley,      -32.0,  32.0, 0.0),
    Benchmark("Schwefel",   schwefel,   -500.0, 500.0, 0.0),
]


def get_benchmark(name: str) -> Benchmark:
    for b in BENCHMARKS:
        if b.name.lower() == name.lower():
            return b
    raise KeyError(name)
