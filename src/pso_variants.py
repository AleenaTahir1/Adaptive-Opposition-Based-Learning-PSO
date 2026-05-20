"""Three PSO variants compared in the AO-PSO study.

* :class:`StandardPSO` -- Clerc-constriction PSO baseline.
* :class:`OPSO`        -- Jabeen, Jalil, Baig (GECCO 2009), the focal paper.
* :class:`AOPSO`       -- our proposed Adaptive Opposition-Based PSO,
  carrying the seven value additions documented in the report.

Each optimiser returns a :class:`RunResult` containing the best fitness
trajectory indexed by function evaluations (FEs), so all algorithms are
compared on the same wall-clock-fair x-axis -- one of the open problems
the literature review (Phase 4) flagged in the OBL-PSO line.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Callable, Optional

import numpy as np

from obl import (
    generalized_opposite,
    logistic_chaotic_init,
    standard_opposite,
    swarm_diameter,
)


# ---------- helpers ---------------------------------------------------------

def clerc_constriction(c1: float, c2: float) -> float:
    """Clerc's constriction factor chi."""
    phi = c1 + c2
    assert phi > 4.0, "Clerc constriction requires c1 + c2 > 4"
    return 2.0 / abs(2.0 - phi - np.sqrt(phi * phi - 4.0 * phi))


@dataclass
class RunResult:
    algorithm: str
    benchmark: str
    dim: int
    best_fitness: float
    best_position: np.ndarray
    fe_history: list[int] = field(default_factory=list)
    fit_history: list[float] = field(default_factory=list)
    diversity_history: list[float] = field(default_factory=list)
    jr_history: list[float] = field(default_factory=list)


@dataclass
class HookOutcome:
    X: np.ndarray
    V: np.ndarray
    fits: np.ndarray
    pbest: np.ndarray
    pbest_fit: np.ndarray
    gbest: np.ndarray
    gbest_fit: float
    fe_added: int = 0
    jr: float = 0.0


# ---------- core PSO body shared by every variant ---------------------------

class _PSOCore:
    """Implements the Clerc-constriction PSO update loop.

    Subclasses override ``_init_swarm`` (initialisation step),
    ``_extra_velocity_term`` (extra velocity contribution), and
    ``_post_update_hook`` (per-generation extras) to inject the
    O-PSO / AO-PSO behaviours.
    """

    def __init__(
        self,
        N: int = 20,
        c1: float = 2.05,
        c2: float = 2.05,
        v_clamp_frac: float = 0.5,
        max_fe: int = 30_000,
        rng: Optional[np.random.Generator] = None,
    ):
        self.N = N
        self.c1 = c1
        self.c2 = c2
        self.chi = clerc_constriction(c1, c2)
        self.v_clamp_frac = v_clamp_frac
        self.max_fe = max_fe
        self.rng = rng if rng is not None else np.random.default_rng()

    # ---- to override -------------------------------------------------------

    def _init_swarm(
        self,
        D: int,
        low: np.ndarray,
        high: np.ndarray,
        fitness: Callable[[np.ndarray], float],
    ) -> tuple[np.ndarray, np.ndarray, int]:
        """Return ``(X, fits, fe_used)``. Pure random by default."""
        X = self.rng.uniform(low, high, size=(self.N, D))
        fits = np.array([fitness(x) for x in X])
        return X, fits, self.N

    def _extra_velocity_term(self, X, pbest, pbest_fit):
        """No extra velocity term by default. AO-PSO overrides for rebel."""
        return 0.0

    def _post_update_hook(
        self,
        gen: int,
        X: np.ndarray,
        V: np.ndarray,
        fits: np.ndarray,
        pbest: np.ndarray,
        pbest_fit: np.ndarray,
        gbest: np.ndarray,
        gbest_fit: float,
        improved_this_gen: np.ndarray,
        low: np.ndarray,
        high: np.ndarray,
        fitness: Callable[[np.ndarray], float],
    ) -> HookOutcome:
        """Default hook: identity (no extra OBL)."""
        return HookOutcome(X, V, fits, pbest, pbest_fit, gbest, gbest_fit,
                           fe_added=0, jr=0.0)

    # ---- main loop ---------------------------------------------------------

    def run(self, benchmark, D: int) -> RunResult:
        low, high = benchmark.bounds(D)
        width = high - low
        v_max = self.v_clamp_frac * width

        result = RunResult(
            algorithm=self.__class__.__name__,
            benchmark=benchmark.name,
            dim=D,
            best_fitness=np.inf,
            best_position=np.zeros(D),
        )

        X, fits, fe_used = self._init_swarm(D, low, high, benchmark.func)
        V = self.rng.uniform(-v_max, v_max, size=(self.N, D))

        pbest = X.copy()
        pbest_fit = fits.copy()
        g_idx = int(np.argmin(pbest_fit))
        gbest = pbest[g_idx].copy()
        gbest_fit = float(pbest_fit[g_idx])

        result.fe_history.append(fe_used)
        result.fit_history.append(gbest_fit)
        result.diversity_history.append(swarm_diameter(X, low, high))
        result.jr_history.append(0.0)

        gen = 0
        while fe_used < self.max_fe:
            gen += 1
            r1 = self.rng.random((self.N, D))
            r2 = self.rng.random((self.N, D))

            extra = self._extra_velocity_term(X, pbest, pbest_fit)
            V = self.chi * (V
                            + self.c1 * r1 * (pbest - X)
                            + self.c2 * r2 * (gbest - X)
                            + extra)
            V = np.clip(V, -v_max, v_max)
            X = np.clip(X + V, low, high)

            fits = np.array([benchmark.func(x) for x in X])
            fe_used += self.N

            improved = fits < pbest_fit
            pbest = np.where(improved[:, None], X, pbest)
            pbest_fit = np.where(improved, fits, pbest_fit)
            g_idx = int(np.argmin(pbest_fit))
            if pbest_fit[g_idx] < gbest_fit:
                gbest_fit = float(pbest_fit[g_idx])
                gbest = pbest[g_idx].copy()

            outcome = self._post_update_hook(
                gen, X, V, fits, pbest, pbest_fit, gbest, gbest_fit,
                improved, low, high, benchmark.func,
            )
            X = outcome.X
            V = outcome.V
            fits = outcome.fits
            pbest = outcome.pbest
            pbest_fit = outcome.pbest_fit
            gbest = outcome.gbest
            gbest_fit = outcome.gbest_fit
            fe_used += outcome.fe_added

            result.fe_history.append(fe_used)
            result.fit_history.append(gbest_fit)
            result.diversity_history.append(swarm_diameter(X, low, high))
            result.jr_history.append(outcome.jr)

        result.best_fitness = gbest_fit
        result.best_position = gbest
        return result


# ---------- variants --------------------------------------------------------

class StandardPSO(_PSOCore):
    """Plain Clerc-constriction PSO. No OBL anywhere."""


class OPSO(_PSOCore):
    """Jabeen, Jalil, Baig (GECCO 2009) -- the focal paper.

    Differences from StandardPSO:

    * Initialise N particles uniformly, plus N opposites using Tizhoosh's
      linear ``a + b - x`` formula. Evaluate all 2N, keep the best N.

    No OBL is applied after the initial generation -- this is precisely
    the limitation AO-PSO addresses.
    """

    def _init_swarm(self, D, low, high, fitness):
        X = self.rng.uniform(low, high, size=(self.N, D))
        Xhat = standard_opposite(X, low, high)
        combined = np.vstack([X, Xhat])
        combined_fits = np.array([fitness(c) for c in combined])
        order = np.argsort(combined_fits)[: self.N]
        X = combined[order]
        fits = combined_fits[order]
        return X, fits, 2 * self.N


class AOPSO(_PSOCore):
    """Adaptive Opposition-Based PSO -- our proposed algorithm.

    Carries every value addition spelled out in the report:

    1. Chaotic (logistic-map) initialisation.
    2. Generalized OBL formula for the initial opposite pool.
    3. Generation jumping (OBL applied throughout the search).
    4. Adaptive Jr driven by live swarm diameter.
    5. pbest opposition when a particle stagnates for ``t_stag`` generations.
    6. Rebel-repulsion velocity term away from the worst pbest.
    7. Function-evaluation budget tracking.
    """

    def __init__(
        self,
        N: int = 20,
        c1: float = 2.05,
        c2: float = 2.05,
        v_clamp_frac: float = 0.5,
        max_fe: int = 30_000,
        jr_max: float = 0.30,
        lam: float = 25.0,
        t_stag: int = 5,
        alpha_rebel: float = 0.05,
        rng: Optional[np.random.Generator] = None,
    ):
        super().__init__(N=N, c1=c1, c2=c2, v_clamp_frac=v_clamp_frac,
                         max_fe=max_fe, rng=rng)
        self.jr_max = jr_max
        self.lam = lam
        self.t_stag = t_stag
        self.alpha_rebel = alpha_rebel
        self._stag_counter: Optional[np.ndarray] = None

    # ---- value additions 1 + 2: chaotic init with generalized opposites ----
    def _init_swarm(self, D, low, high, fitness):
        X = logistic_chaotic_init(self.N, D, low, high, self.rng)
        Xhat = generalized_opposite(X, low, high, self.rng)
        combined = np.vstack([X, Xhat])
        combined_fits = np.array([fitness(c) for c in combined])
        order = np.argsort(combined_fits)[: self.N]
        X = combined[order]
        fits = combined_fits[order]
        self._stag_counter = np.zeros(self.N, dtype=int)
        return X, fits, 2 * self.N

    # ---- value addition 6: rebel repulsion ---------------------------------
    def _extra_velocity_term(self, X, pbest, pbest_fit):
        worst_idx = int(np.argmax(pbest_fit))
        worst = pbest[worst_idx]
        return -self.alpha_rebel * (worst - X)

    # ---- value additions 3 + 4 + 5: adaptive generation jumping etc --------
    def _post_update_hook(self, gen, X, V, fits, pbest, pbest_fit,
                          gbest, gbest_fit, improved_this_gen,
                          low, high, fitness):
        if self._stag_counter is None:
            self._stag_counter = np.zeros(self.N, dtype=int)
        self._stag_counter = np.where(improved_this_gen, 0,
                                      self._stag_counter + 1)

        sigma = swarm_diameter(X, low, high)
        jr = self.jr_max * (1.0 - np.exp(-self.lam * sigma))
        fe_added = 0

        # generation jumping: OBL on current positions with probability Jr
        if self.rng.random() < jr:
            Xhat = generalized_opposite(X, low, high, self.rng)
            new_fits = np.array([fitness(x) for x in Xhat])
            fe_added += self.N
            better = new_fits < fits
            X = np.where(better[:, None], Xhat, X)
            fits = np.where(better, new_fits, fits)
            new_pb = fits < pbest_fit
            pbest = np.where(new_pb[:, None], X, pbest)
            pbest_fit = np.where(new_pb, fits, pbest_fit)

        # pbest opposition on stagnated particles
        stagnated = np.where(self._stag_counter >= self.t_stag)[0]
        if stagnated.size > 0:
            sub_X = pbest[stagnated]
            sub_hat = generalized_opposite(sub_X, low, high, self.rng)
            sub_fits = np.array([fitness(x) for x in sub_hat])
            fe_added += stagnated.size
            for k, idx in enumerate(stagnated):
                if sub_fits[k] < pbest_fit[idx]:
                    pbest[idx] = sub_hat[k]
                    pbest_fit[idx] = sub_fits[k]
                self._stag_counter[idx] = 0

        g_idx = int(np.argmin(pbest_fit))
        if pbest_fit[g_idx] < gbest_fit:
            gbest_fit = float(pbest_fit[g_idx])
            gbest = pbest[g_idx].copy()

        return HookOutcome(X, V, fits, pbest, pbest_fit, gbest, gbest_fit,
                           fe_added=fe_added, jr=jr)
