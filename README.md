creates actuariel trinagle excel file, still in production
known issues
problem when excel starts up, 
problem with deciding between 1 2 3 4 triangles. 
need to fix so the triangle doesnt have a higher paid than ibnf

# Actuarial Triangle Automation System — V4

Automated Excel workbook builder for actuarial loss reserving. Generates a fully formatted `.xlsx` file with 4 linked triangles, 8 reserving methods, user-controlled blending weights, and a completed square — all from a single Python script.

---

## What It Does

Takes raw loss triangle data and outputs a structured Excel workbook with:
- **8 reserving methods** per triangle (Chain Ladder, Mack, BF, Cape Cod, Loss Ratio, Freq-Sev, GLM, Bayesian)
- **Blended ultimate** per origin year, weighted by user-defined method weights
- **Completed square** projecting all future triangle cells
- **Diagnostics** (monotonicity, CV of LDFs, % reported, Mack SE, GLM dispersion)
- **Python-computed methods** (Mack variance, GLM ODP, Freq-Sev, Bayesian) written directly into cells

---

## Requirements

```
pip install openpyxl
```

Python 3.8+. No other dependencies.

---

## Output

```
C:\Users\yosef\Downloads\build_v4.12\ActuarialTriangle_V4.xlsx
```

Change `OUTPUT_PATH` at the top of the script to redirect.

---

## Workbook Structure

| Sheet | Contents |
|---|---|
| `_index` | Registry of all triangles (metadata) |
| `_control` | Global settings (tail factor, risk sensitivity, thresholds) |
| `T1_paid` | Paid loss triangle — primary triangle |
| `T2_incurred` | Case incurred loss triangle |
| `T3_counts` | Reported claim counts (placeholder — AXIS does not disclose) |
| `T4_premium` | Earned premium by origin year |

---

## Methods

| Code | Method | Type |
|---|---|---|
| CL | Chain Ladder | Excel formula |
| Mack | Mack 1993/1999 | Python (variance + SE) |
| BF | Bornhuetter-Ferguson | Excel formula |
| CC | Cape Cod | Excel formula |
| LR | Loss Ratio | Excel formula |
| FS | Frequency-Severity | Python (counts + severity) |
| GLM | Over-dispersed Poisson GLM | Python (Renshaw-Verrall) |
| Bayes | Bayesian Gamma-Poisson | Python (conjugate update) |

> Green cells = Excel formulas. Blue-green cells = Python-computed static values.

---

## Method Weights

- Toggle between **Uniform** (one weight row applies to all years) or **Per-AY** (per accident year weights)
- Weights must sum to **100%** — workbook validates and shows ERROR/INFO/OK
- Default weights are auto-calculated from data signals (maturity, LDF volatility, premium availability) following Friedland Ch. 7–15

---

## How to Run

1. Install `openpyxl`
2. Set `OUTPUT_PATH` to your desired output location
3. Replace `paid_matrix`, `incd_matrix`, `counts_matrix`, `premium_matrix` with your data (`None` = missing cell)
4. Run:

```bash
python actuarial_triangle_v4.py
```

5. Open the generated `.xlsx` in Excel
6. Adjust weights and ELR inputs as needed — all formula cells recalculate automatically

> If you get a `PermissionError`, close the file in Excel first then re-run.

---

## Data Format

```python
# origin_periods: list of accident years
origin_periods = [2014, 2015, ..., 2023]

# dev_periods: list of development ages (months)
dev_periods = [12, 24, 36, ..., 120]

# data_matrix: list of lists, None for missing cells
paid_matrix = [
    [359685, 937044, ..., 1906497],  # 2014 — fully developed
    [326645, 863573, ..., None],     # 2015 — last cell missing
    ...
]
```

---

## Key Inputs Per Triangle

| Input | Where | Notes |
|---|---|---|
| ELR | Method Inputs section | Auto-defaults to Cape Cod derived ELR; user can override |
| Premium | Method Inputs section | Can link to T4_premium sheet or enter manually |
| Tail Factor | LDF row | Default 1.000; enter >1.000 if development continues past last age |
| Bayesian Prior | Method Inputs section | Total prior ultimate across all origins |

---

## Sample Data Source

AXIS Capital 2023 Global Loss Triangles (consolidated total, accident year basis, USD).  
Counts triangle is a placeholder — AXIS does not disclose reported claim counts.

---

## Architecture

```
main()
├── build_index_sheet()       # _index tab
├── build_control_sheet()     # _control tab
├── add_triangle(...)         # one call per triangle → builds all 9 sections
│   ├── Section 1: Metadata
│   ├── Section 2: Data Matrix
│   ├── Section 3: Link Ratios
│   ├── Section 4: LDFs / CDFs / Tail
│   ├── Section 5: Method Ultimates (8 methods)
│   ├── Section 6: Method Inputs (ELR, Premium, Prior)
│   ├── Section 7: Method Weights (Uniform + Per-AY)
│   ├── Section 8: Blended Ultimate & IBNR
│   ├── Section 8b: Completed Square
│   └── Section 9: Diagnostics
└── run_python_methods(...)   # writes Mack/GLM/FS/Bayes values into cells
```
