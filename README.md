# AO-PSO — Adaptive Opposition-Based Particle Swarm Optimization

We took the focal paper *O-PSO* (Jabeen, Jalil & Baig, GECCO 2009) and rebuilt it into a modern, statistically-validated optimizer.

## Why this project

Standard PSO has two well-known weaknesses: it converges prematurely, and it depends heavily on the quality of its initial population. **Opposition-Based Learning (OBL)** addresses both by evaluating each candidate alongside its opposite — but in the focal O-PSO paper, OBL only fires at generation zero. After that, the algorithm reverts to plain PSO.

Our Phase-4 literature review across fifteen successor papers identified four open problems that persist in the OBL-PSO line:

- OBL is almost always confined to generation zero or applied with a *fixed* jumping rate
- The linear opposite is geometrically rigid
- Computational cost is rarely reported in function evaluations
- Statistical significance tests are absent in most papers

**AO-PSO** closes all four gaps and adds a genuinely new mechanism on top.

## The seven value additions

| # | Mechanism | Source | Novel? |
|---|-----------|--------|:------:|
| VA-1 | Chaotic logistic-map initialization | CSPSO | |
| VA-2 | Generalized OBL (GOBL) — population-mean-aware opposite | GOPSO | |
| VA-3 | Generation jumping (OBL every iteration) | OVCPSO / OPSO-Cauchy | |
| VA-4 | **Adaptive Jr driven by live swarm diameter** | — | ★ |
| VA-5 | pbest opposition refresh on stagnation | BPSO-OBL | |
| VA-6 | Rebel-repulsion velocity term | BPSO-OBL | |
| VA-7 | FE-budget accounting + Wilcoxon significance | best practice | |

The novel one is VA-4. We define the normalised swarm diameter `σ_t` (mean per-dimension std divided by box width) and set:

```
Jr(t) = Jr_max · ( 1 − exp(−λ · σ_t) )
```

When the swarm is spread out, OBL fires aggressively. When it converges, `Jr` decays to zero so no FEs are wasted on late-stage opposition. No prior paper in the OBL-PSO line ties `Jr` to a live measurement of swarm state.

## Headline results — mean fitness over 30 runs at D = 30

| Benchmark   | Standard PSO | O-PSO (focal) | AO-PSO (ours)   | Wilcoxon vs O-PSO     |
|-------------|-------------:|--------------:|----------------:|-----------------------|
| Sphere      | 1.67e+03     | 6.67e+02      | **4.09e-106**   | p = 9.3e-10 ✓         |
| Rosenbrock  | 1.52e+04     | 6.26e+03      | **2.35e+01**    | p = 0.027 ✓           |
| Rastrigin   | 1.06e+02     | 9.68e+01      | **0.00**        | p = 9.3e-10 ✓         |
| Griewank    | 1.51e+01     | 6.38          | **0.00**        | p = 1.3e-6 ✓          |
| Ackley      | 4.47         | 4.46          | **4.44e-16**    | p = 9.3e-10 ✓         |
| Schwefel    | 4.09e+03     | 3.86e+03      | 5.57e+03        | honest loss           |

AO-PSO wins on **5 of 6 benchmarks**, with most p-values below 10⁻⁶. The Schwefel loss is reported honestly: its optimum sits at x ≈ 420.97, far from the centre of the search box, so opposition-based methods naturally underperform. We outline a centre-shifted opposite as a remedy for future work.

Total experimental footprint: **1,080 independent optimization runs** (3 algorithms × 6 benchmarks × 2 dimensions × 30 seeds).

## Repository layout

```
OBI_PSO/
  src/
    benchmarks.py              Sphere · Rosenbrock · Rastrigin · Griewank · Ackley · Schwefel
    obl.py                     Standard OBL · GOBL · chaotic init · swarm diameter
    pso_variants.py            StandardPSO · OPSO · AOPSO  (one class each)
    experiments.py             30-run sweep, Wilcoxon, CSV output
    plots.py                   Convergence · diversity · boxplot figures
    build_ppt_v2.py            Generates the colourful infographic presentation
    extract_tikz_figs.py       Rasterises TikZ diagrams from the report PDF
  report/
    Final_Phase_Report.tex     NUTECH-style project report (compiled below)
    Final_Phase_Report.pdf
    Research_Paper.pdf         IEEE-format paper version
    figs/                      Architecture · flowchart · convergence · boxplots · diversity
    logo.png                   NUTECH crest
  results/
    raw_runs.csv               One row per (algo, benchmark, dim, seed)
    summary.csv                Mean / std / best / worst / median per cell
    wilcoxon.csv               Significance tests vs PSO and vs O-PSO
  AO_PSO_Presentation_v2.pptx  Final infographic deck (23 slides)
  Final_Phase_Doc (1).pdf      Phase-4 comparative review (handed in earlier)
  README.md
```

## How to reproduce

```bash
cd src
python experiments.py    # ~17 min on a modern laptop, produces all CSVs
python plots.py          # ~3 s, writes PDF and PNG figures
cd ../report
pdflatex Final_Phase_Report.tex && pdflatex Final_Phase_Report.tex
```

Requirements: Python 3.12 with `numpy`, `scipy`, `pandas`, `matplotlib`, `python-pptx`. LaTeX side needs `IEEEtran`, `tikz`, `booktabs`, `caption`, `xcolor` — everything a default MiKTeX install ships with.

## Team

| Roll | Name | Responsibility |
|-----:|------|----------------|
|  1   | Aleena Tahir   | IEEE compilation, Implementation|
| 10   | Eman Asghar    | Architecture and flowchart diagrams |
| 20   | Aena Habib     | Value-addition lead, methodology design |
| 22   | Malaika Akhter | Abstract, introduction, conclusion |
| 23   | Dua Kamal      | Presentation design, visuals |
| 29   | Sadia Mazhar   | Related-work section |
| 38   | Inshal Adil    | Results analysis, plots, parameter rationale |
| 48   | Saqlain Abbas  | Implementation, debugging, experimental runs |

Department of Artificial Intelligence · Semester 6 · Instructor: Lec. Faria Sajjad
