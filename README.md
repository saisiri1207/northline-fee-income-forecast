# Northline Fee Income Forecast (Excel)

Monthly fee income forecast by fee type, built from transaction volumes, fee rates and seasonality, with forecast vs actual accuracy.

> Northline Bank Ops is a fictional sample in a bank operations style. Figures are generated sample data, not from PNC or any real bank.

## Problem

Northline forecast fee income by copying the same month last year. That ignored volume growth and the 2026 fee schedule, so the forecast ran low and accuracy was weak.

## What I built

`Northline_Fee_Income_Forecast.xlsx`, built with `build_workbook.py` (openpyxl). Yellow cells are inputs and black cells are formulas.

| Sheet | Contents |
|---|---|
| Actuals | Jan 2024 to Jun 2026 volume and fee per transaction for 4 fee types. Fee income = volume x rate |
| Drivers | Year over year volume growth (SUMIFS) and the latest fee per transaction |
| Forecast | Jan to Jun 2026 naive forecast vs driver based forecast, absolute % error, accuracy (1 minus MAPE) |

Fee types: wire transfer fees, ACH origination fees, treasury service charges, card interchange.

Driver forecast = same month last year volume x (1 + volume growth) x latest fee rate. Seasonality comes through the prior year month.

## How to run

```
pip install openpyxl
python3 build_workbook.py
```

## Key results (sample figures)

| Method | Accuracy (1 minus MAPE) | Jan to Jun 2026 forecast |
|---|---|---|
| Naive (same month last year) | 85.3% | $13,010,108 |
| Driver based | **92.9%** | $14,794,485 |
| Actual | | $15,261,240 |

FY2025 fee income in the sample was $26.6M across 4 fee types.

## Conclusion

Adding volume growth and the current fee schedule raised accuracy from about **85% to 93%** and closed most of the $2.25M gap the naive method left in the first half. The driver layout also makes it clear whether a miss came from volume or from rate.

## Stack

Excel, Python (openpyxl). No macros.
