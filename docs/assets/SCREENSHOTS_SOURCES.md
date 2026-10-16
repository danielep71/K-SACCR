# Workbook screenshots

The `SACCR_*.png` screenshots show sheets of the workbook built from
`src/workbook/SACCR_Template.xlsx`, with the demonstration portfolio. The owner
took them in Excel for Windows on 2026-10-16.

| Image | Shows |
| --- | --- |
| `SACCR_NettingSets.png`, `SACCR_Trades.png` | The input sheets |
| `SACCR_Params.png`, `SACCR_Params_Supervisory.png` | Run parameters, and the supervisory table |
| `SACCR_README_sheet.png` | The in-workbook guide |
| `SACCR_Results_AddOns.png`, `SACCR_Results_EAD.png` | Results after a run: RC and add-ons, then PFE, EAD, cap and status |
| `SACCR_TradeCalc_Inputs.png`, `SACCR_TradeCalc_AddOns.png` | Every trade-level quantity of that run: dates, duration, adjusted notional and delta, then maturity factors, effective notionals and add-on |
| `SACCR_HedgingSets.png`, `SACCR_Buckets.png` | The hedging-set and bucket detail of that run |

The input sheets show the template's own content; the output sheets show a run
of that portfolio on that day.

- **Processing:** cropped to the sheet area (no window title, sheet tabs or
  status bar) and reduced to a 256-colour PNG palette without dithering. They
  are not otherwise edited.
- **Number format:** the values follow the regional settings of the machine
  that took them (decimal comma, thousands point); the workbook itself reads
  and writes locale-independent values.
- **Use:** the root README shows both Results images, the calculation detail
  and the input sheets in collapsible sections; the Wiki can reference any of
  them.
- **CRE99 Example 3:** the demonstration workbook shows an EAD of 5,405
  (thousands), the published figure 5,406. The demo enters the WTI trade with
  actual dates over 365 days a year (TradeCalc shows M = 0.7507); the
  regression fixture enters 270 days over 360, exactly 0.75 years as printed,
  and matches 5,406.

Retake a screenshot when the layout or content of its sheet changes, and keep
the same processing.
