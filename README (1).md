# Basel Risk Command Center

> End-to-end credit risk analytics platform — IRB capital calculations, ICAAP stress testing, and multi-jurisdiction RWA comparison, built on an open-source Python library with an interactive web dashboard.

---

## What It Is

A Python-generated credit risk portfolio that computes Basel III/IV capital requirements from scratch — starting from raw trade inputs — and presents results through an interactive HTML dashboard. The project covers the full analytical chain a risk analyst performs: EAD calculation → PD/LGD inputs → IRB risk weights → RWA → stress testing → jurisdiction comparison.

Built using [`creditriskengine`](https://pypi.org/project/creditriskengine/) (v0.4), a production-grade open-source library that implements Basel regulatory formulas traced to their exact BCBS paragraph references.

---

## Why It Was Built

Most Basel capital analytics happen in Excel spreadsheets or expensive vendor platforms. This project demonstrates that a risk professional with domain expertise — but limited coding background — can use Python and open-source tooling to replicate institutional-grade analysis, understand the mechanics behind the numbers, and present results interactively.

The project was built incrementally: starting with a single IRB risk weight calculation, expanding to a full multi-asset portfolio, adding stress scenarios and jurisdiction comparison, then packaging everything into a formatted Excel report and a live web dashboard.

---

## Features

**Portfolio**
- 134 exposures across 8 asset classes: Sovereign, Bank, Corporate, SME, Residential Mortgage, QRRE, Other Retail, Derivatives
- 15 derivative types: IRS, Cross Currency Swap, FX Forward/Option, Equity TRS/Option/Forward, CDS, Commodity Swap/Forward, Bond Forward, Swaption
- EAD computed from raw inputs — drawn amount + CCF × undrawn for loans; max(MTM,0) + notional × add-on for derivatives
- Covers 17+ countries across India, USA, UK, Europe, Japan, Singapore, Brazil, South Africa

**IRB Stress Testing (ICAAP)**
- Three scenarios: Baseline (1×), Adverse (1.5× PD), Severely Adverse (2.5× PD)
- RWA and minimum capital (8%) computed per scenario
- Capital uplift visible across all exposures and in aggregate

**Multi-Jurisdiction SA Comparison**
- Same portfolio run through 4 regulatory frameworks: BCBS, EU CRR3, UK PRA, RBI (India)
- Highlights divergences: RBI flat LTV-based mortgage risk weights, EU SME supporting factor (0.7619×), UK PRA loan-splitting for residential real estate

**Excel Report (`ICAAP_Dashboard.xlsx`)**
- Sheet 1 — Portfolio Detail: every exposure with full calculation transparency (Drawn → CCF → EAD → PD → LGD → RW% → RWA → Capital)
- Sheet 2 — IRB Stress Test: per-exposure scenario comparison
- Sheet 3 — Jurisdiction Comparison: BCBS vs EU vs UK vs India side by side
- Sheet 4 — Executive Summary: sector breakdown, scenario summary, jurisdiction totals

**Web Dashboard (`index.html`)**
- 4 pages: Executive Dashboard, Risk Analytics, Portfolio, About
- Approach dropdown: 7 options (IRB ×3 scenarios + SA ×4 jurisdictions)
- Live KPI cards: Total EAD, Total RWA, Min Capital, Avg Risk Weight
- Charts: Sector RWA, IRB Scenario Comparison, SA Jurisdiction Comparison, EAD by Asset Class, Country distribution, Rating distribution, Top 15 by RWA, PD histogram, Derivative EAD breakdown
- Portfolio table: sortable, filterable by sector/country/rating/type, search by name
- Responsive: desktop sidebar nav, mobile top nav

---

## Tech Stack

| Layer | Tools |
|---|---|
| Risk calculations | `creditriskengine` v0.4 (Basel III/IV IRB, SA, EAD, ECL) |
| Excel output | `openpyxl` |
| JSON output | Python `json` |
| Frontend | Vanilla JS, Apache ECharts, HTML5/CSS3 |
| Fonts | IBM Plex Sans, DM Mono |
| Python version | 3.11+ |

---

## Setup

**1. Install dependencies**
```bash
pip install creditriskengine openpyxl
```

**2. Run the Python script**
```bash
python dashboard.py
```

This generates two files in the working directory:
- `ICAAP_Dashboard.xlsx` — formatted Excel report
- `portfolio.json` — full portfolio dataset consumed by the dashboard

**3. Open the dashboard**

Place `index.html` and `portfolio.json` in the same folder, then open `index.html` in a browser. The dashboard reads `portfolio.json` via a local `fetch()` call.

> **Note:** Some browsers block local file fetch requests. If the dashboard shows a loading error, either serve the files via a local server (`python -m http.server`) or use VS Code Live Server extension.

---

## File Structure

```
├── dashboard.py          # Main Python script — portfolio + calculations + outputs
├── portfolio.json        # Generated: full exposure data with all computed fields
├── ICAAP_Dashboard.xlsx  # Generated: 4-sheet Excel report
└── index.html            # Web dashboard
```

---

## Key Concepts

**EAD Calculation**
- Loans: `EAD = Drawn + CCF × Undrawn` — CCF values from Basel supervisory tables
- Derivatives: `EAD = max(MTM, 0) + Notional × Add-on%` — add-on rates vary by asset class (IR: 0.5–1.5%, FX: 1–1.5%, Equity: 6%, Commodity: 15%, Credit: 5%)

**IRB Risk Weight Formula (CRE31.4–31.10)**
- `K = LGD × [N(G(PD)/√(1−R) + √(R/(1−R)) × G(0.999)) − PD]`
- `RW = K × 12.5 × MA` (corporate/sovereign/bank); `RW = K × 12.5` (retail)
- Asset correlation R differs by class: corporate [12–24%], mortgage 15% fixed, QRRE 4% fixed

**Stress Testing Logic**
- PD multiplied by scenario factor: Adverse 1.5×, Severely Adverse 2.5×
- PD capped at 100%; all other parameters (LGD, EAD, maturity) held constant
- Capital impact = change in RWA × 8%

---

## Regulatory Coverage

| Framework | Reference |
|---|---|
| Basel III IRB | BCBS CRE31–CRE33 |
| Basel III SA | BCBS CRE20 |
| EU Standardised | CRR3 (incl. SME supporting factor Art. 501) |
| UK PRA | PS9/24 (loan-splitting for residential RE) |
| India RBI | RBI Master Circular on Basel III |
| EAD/CCF | BCBS CRE32.26–32.33 |

---

## Disclaimer

This project is for educational and analytical purposes. Calculations are based on the `creditriskengine` library which has not been reviewed or endorsed by any regulatory authority. Not for use in production capital calculations without independent validation.
