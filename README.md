# Local Volatility Surface Modeling in Excel

**Microsoft Excel | Options Analytics | Implied Volatility | Finite-Difference Modeling**

An Excel-based quantitative finance project that transforms a grid of **110 implied-volatility observations** (11 strikes × 10 maturities) into a **9 × 9 grid of 81 local-volatility estimates** using Black–Scholes intermediate calculations and finite-difference approximations. The workbook demonstrates formula-driven financial modeling and multi-dimensional sensitivity analysis without requiring VBA or Python.

> **Scope note:** This is a spreadsheet implementation of a local-volatility approximation based on an implied-volatility grid, not a production-calibrated or independently validated trading model.

## Project highlights

| Metric | Value |
| --- | ---: |
| Implied-volatility input observations | **110** |
| Strike prices | **11**, from **430 to 530**, in increments of 10 |
| Expiries | **10**, spanning **0.14 to 1.59 years** |
| Local-volatility estimates populated | **81** (9 strikes × 9 maturities) |
| Underlying spot price assumed | **423.46** |
| Risk-free interest rate assumed | **1.00%** |
| Computed local-volatility range | **10.67%–111.85%** |
| Mean of populated local-volatility cells | **27.96%** |

The local-volatility range is a **description of model outputs**, not a measure of predictive accuracy. One populated cell is above 100%, illustrating the sensitivity of derivative-based estimates to the volatility grid and numerical approximation.

## Repository structure

```text
local-volatility-modeling-excel/
├── README.md
└── Local Volatility Calculation.xlsx
```

## Objective

Implied volatility is an option-market pricing input that varies across strike and maturity. Local volatility seeks to describe volatility as a function of the underlying asset level and time. This project works through the numerical ingredients of such a transformation in Excel by computing:

1. Black–Scholes **d₁** across the strike/maturity grid.
2. The first finite difference of implied volatility with respect to **maturity**.
3. The first finite difference of implied volatility with respect to **strike**.
4. The second finite difference of implied volatility with respect to **strike**.
5. The numerator, denominator, and square-root ratio used to obtain the workbook's **local-volatility estimates**.

## Workbook walkthrough

### `Sheet1` — Local-volatility calculation workflow

| Excel section | Main range | Purpose |
| --- | --- | --- |
| Model assumptions | `C2:D3`, `A25:B26` | Spot price and risk-free rate |
| Maturity and input IV surface | `A11:K23` | 11 strikes × 10 maturities (110 IV inputs) |
| Black–Scholes d₁ | `A28:K40` | d₁ values at strike–maturity coordinates |
| First difference across maturities | `A43:K56` | Approximation to ∂σ/∂T |
| First difference across strikes | `A59:K72` | Approximation to ∂σ/∂K |
| Second difference across strikes | `A76:K89` | Approximation to ∂²σ/∂K² |
| Local-volatility numerator | `A92:K105` | Formula component combining IV and derivative terms |
| Local-volatility denominator | `A108:K120` | Formula component combining d₁ and strike derivatives |
| Local-volatility output | `B126:K137` | Final **81 populated** volatility values |

`Sheet2` contains an additional numerical volatility grid. It is not a separate dashboard or a VBA-powered workflow; some entries differ from the `Sheet1` input grid, so the two ranges should not automatically be described as identical.

## Methodology and example Excel formulas

### 1. Black–Scholes d₁

The spreadsheet calculates:

\[
d_1 = \frac{\ln(S/K) + (r+\tfrac{1}{2}\sigma^2)T}{\sigma\sqrt{T}}
\]

Example from `Sheet1!B30`:

```excel
=(LN($B$25/$A30)+($B$26+0.5*B13*B13)*B$29)/(B13*SQRT(B$29))
```

This shows the use of `LN`, `SQRT`, relative references, and mixed/absolute cell references to make a two-dimensional calculation grid reusable.

### 2. Finite differences

The workbook estimates implied-volatility changes across neighboring maturity and strike observations. Representative **actual** formulas include:

**Change with maturity** — `Sheet1!C46`:

```excel
=(C13-B13)/(C$45-B$45)
```

**Change with strike** — `Sheet1!B63`:

```excel
=(B14-B13)/($A63-$A62)
```

**Second change with strike** — `Sheet1!C80`:

```excel
=(C15-2*C14+C13)/(10*10)
```

The second-difference example uses a hard-coded strike increment of 10; this is appropriate for the workbook's current uniform grid but would need revision for nonuniform strikes.

### 3. Local-volatility estimates

The workbook builds a numerator and denominator using the finite differences, then calculates the square root of their ratio. Example from `Sheet1!C128`:

```excel
=SQRT(C96/C111)
```

That formula corresponds to the workbook's local-volatility implementation. The resulting 9 × 9 interior grid excludes boundary locations where the numerical differencing method does not produce a full set of required terms.

## Results and interpretation

- **110 implied-volatility inputs** provide the initial strike-by-maturity surface.
- **81 local-volatility estimates** are populated after finite-difference processing and boundary exclusions.
- The populated results range from approximately **10.67% to 111.85%**, averaging **27.96%**.
- Values differ by strike and maturity, illustrating how the local-volatility estimate responds to the curvature and term structure of the input implied-volatility grid.
- **One local-volatility estimate exceeds 100%**. This is a reason to inspect numerical stability and the raw data, not an assertion of extraordinary model performance.

These are **descriptive spreadsheet outputs**, not backtest returns, forecast performance, or demonstrated improvements in pricing accuracy.

## Excel skills demonstrated

- **Formula-driven quantitative modeling** across multi-dimensional strike and maturity matrices.
- Financial mathematics with **`LN`**, **`SQRT`**, powers, and arithmetic expressions.
- **Absolute, relative, and mixed cell referencing** (`$B$25`, `$A30`, `B$29`) for consistent formula replication.
- **Finite-difference methods**, including first- and second-order numerical approximations.
- Building and maintaining connected **intermediate calculation blocks** and a final model-output grid.
- Investigating **numerical sensitivity and anomalous outputs**.


## How to explore the workbook

1. Download and open `Local Volatility Calculation.xlsx` in desktop Microsoft Excel.
2. Review the spot price (**423.46**) and annual risk-free rate (**1%**) near the top of `Sheet1`.
3. Inspect the input IV surface in `A11:K23` and the maturity values in row 12.
4. Follow the calculations through the `d₁`, first-difference, second-difference, numerator, and denominator blocks.
5. Inspect final local-volatility estimates in `C128:K136` and select cells to view the underlying formulas.

> **Recalculation caution:** This workbook contains calculated spreadsheet outputs. Changing assumptions or inputs may materially change results, and numerical instability can occur when denominators become small or non-positive.

## Assumptions and limitations

- The spot price and risk-free interest rate are fixed at **423.46** and **1%**, respectively.
- The grid uses **11 evenly spaced strikes** and **10 maturity points**, but the final local-volatility block is smaller due to differencing boundaries.
- The workbook does not document the market-data vendor, collection date, or independent validation of the implied-volatility inputs.
- The finite-difference method can be sensitive to noisy inputs, spacing choices, and numerical denominators; the output above 100% deserves review.
- The spreadsheet presents a **local-volatility calculation**, but does not include a documented external benchmark, production-grade calibration, pricing-error comparison, or out-of-sample backtest.
- `Sheet2` contains a related volatility grid with some values that differ from `Sheet1`; its provenance should be documented before treating it as a reconciled reference dataset.

## Potential extensions

- Add heatmaps and strike-by-maturity surface charts for both implied and local volatility.
- Add spreadsheet checks for non-positive denominators, missing values, and implausibly large outputs.
- Replace hard-coded strike spacing with grid-aware formulas for irregularly spaced strikes.
- Compare against a benchmark volatility surface and document the input-data source.
- Extend the model to price vanilla options and compare model prices to observed market prices.

---

**Tools:** Microsoft Excel · Black–Scholes calculations · Options volatility surfaces · Finite differences · Quantitative financial modeling

**Disclaimer:** Educational modeling project; not financial or investment advice.
