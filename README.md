# Swap Curve Bootstrapping - VBA + Excel

## 📋 Project Information

| | |
|---|---|
| **School** | École Supérieure d'Ingénieurs Léonard-de-Vinci (ESILV) |
| **Program** | Master 1 — Financial Engineering (Ingénierie Financière) |
| **Group** | A4-IF3 (ACTIF 4) |
| **Course** | VBA for Finance |
| **Professor** | Imad ROUGUI |
| **Academic Year** | 2025–2026 |

## 👤 Author

- **Matthieu HANNA GERGUIS**

## 📖 Description

Construction of a simple interest rate swap curve from market quotes using bootstrapping in VBA and Excel. The project builds discount factors and continuously compounded zero rates from three instrument types: Cash deposits (short end), Futures (middle of the curve), and Interest Rate Swaps (long end).

Each computation is implemented twice — once with Excel formulas and once with VBA macros — with an automated cross-validation to ensure consistency between both approaches.

## 📊 Instruments & Methodology

| Segment | Instrument | Bootstrapping Formula |
|---|---|---|
| Short end | Cash Deposits | DF = 1 / (1 + rτ) |
| Middle | Futures | DFᵢ = DFᵢ₋₁ / (1 + fᵢτᵢ) |
| Long end | Swaps | DFₙ = (1 − S·Σ αₖDFₖ) / (1 + S·αₙ) |

**Zero rate extraction:**

z(T) = −ln(DF(T)) / T

## 🏗️ File Structure

```
├── Project_Excel_VBA.xlsm          # Main workbook with macros
│   ├── Inputs Sheet                 # Market data, dates, accrual factors
│   │   ├── Excel formulas (top)     # Transparent step-by-step computation
│   │   └── VBA Results (bottom)     # Macro-generated values for validation
│   └── Curve Sheet                  # Bootstrapped discount factors & zero rates
│       ├── Excel formulas (top)     # Full curve with intermediate columns
│       ├── Validation box           # PASS/FAIL status + max absolute error
│       └── VBA Results (bottom)     # Independent VBA recomputation
│
├── Projet_VBA_ACTIF_4.xlsx         # Project subject / instructions
├── Report_VBA_Project_P4.pdf       # Full project report with explanations & VBA code
└── README.md
```

## ⚙️ Key Features

- **Dual computation**: every derived value is computed both in Excel formulas and in VBA macros for independent validation.
- **Day count conventions**: ACT/360, 30/360, ACT/365 — implemented manually and via `YEARFRAC` for cross-checking.
- **Futures chaining**: start date of each future is set to the previous instrument's end date, creating a clean quarterly sequence.
- **Swap bootstrapping**: par swap formula with annual fixed leg, using a helper function `DFAtYear` to retrieve discount factors at integer maturities.
- **Automated validation**: PASS/FAIL check comparing Excel and VBA outputs with maximum absolute difference displayed.

## 🔧 VBA Macros

| Macro | Description |
|---|---|
| `CalcInputs_VBA_Outputs` | Recomputes all derived input columns (dates, accruals, coupon factors) and writes them as values |
| `CalcCurve_VBA_Outputs` | Bootstraps the full curve (RateUsed, DF, zero rates) independently from Excel formulas |

**Helper functions:**
- `TenorToMonths` — converts tenor strings (e.g. "3M", "5Y") to months
- `BasisFromDayCount` — maps day count convention to Excel basis code
- `AccrualManual` — computes accrual factor using actual day count formulas
- `RateUsedFromQuote` — converts market quote to a consistent rate (handles futures price → forward rate)
- `DFAtYear` — retrieves the discount factor at a given integer year from the bootstrapped curve

## 📈 Curve Outputs

- **Discount factors** for all maturities (cash, futures, swap tenors)
- **Continuously compounded zero rates** derived from discount factors
- **Validation status**: PASS with max |ΔDF| displayed

## 🚀 How to Use

1. Open `Project_Excel_VBA.xlsm` in Excel (enable macros)
2. Review the **Inputs** sheet — market data and derived columns
3. Review the **Curve** sheet — bootstrapped discount factors and zero rates
4. Run macros via `Alt + F8`:
   - `CalcInputs_VBA_Outputs` — validates input computations
   - `CalcCurve_VBA_Outputs` — validates curve bootstrapping
5. Check the validation box on the Curve sheet for PASS/FAIL status

## 📚 References

- Hull, J.C. — *Options, Futures, and Other Derivatives*
- Course materials — VBA for Finance, Imad ROUGUI, ESILV 2025
