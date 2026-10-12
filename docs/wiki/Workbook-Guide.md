# 📗 Workbook Guide

> **Guide, not policy:** [Installation](../../INSTALLATION.md) owns setup, import, and execution.

## Inputs and names

| Sheet | Current contract |
| --- | --- |
| Params | Reporting date/currency, regime, parameters, supervisory factors, and user-supplied FX rates |
| NettingSets | One declared netting set per row; margin terms, collateral, and overrides |
| Trades | One leg per row, unique Trade ID; grouped by Economic trade ID where needed |

NettingSets_Header and Trades_Header identify the first header cell in column C; data begins on the next row. The reviewed template has NettingSets headers on row 6 and Trades headers on row 4. Params uses Params_Code, Params_Value, SF_Table, FX_Table, and parameter names.

Loaders can follow supported named-range moves, but the stale-results fingerprint does not yet cover every relocated Params input ([#152](https://github.com/danielep71/VBA-SACCR-Toolkit/issues/152)). Preserve the standard layout and rerun after changes.

## Economic trades and legs

- Column Z: Economic trade ID.
- Column AA: MtM leg (Y/N).
- Column AB: Comment.
- Blank Economic trade ID: the row is a whole trade.
- Grouped trade: exactly one Y leg carries the economic MtM; all other legs carry zero.
- Every leg shares a netting set and regime. Notional and risk attributes remain per leg.
- Counts distinguish economic trades from legs.

The engine checks the declared structure; it does not discover missing risk drivers or duplicates entered under unrelated identities.

## FormulaTest

FormulaTest replaces the former Tests sheet. The sheet code name remains shTests.

| Position | Meaning |
| --- | --- |
| E4 | Tolerance; workbook name FormulaTest_Tolerance; default 0.001% |
| Row 6 | Headers |
| Rows 7–28 | 22 formula-test rows |
| E:K | Editable arguments |
| L | Expected result calculated with Excel functions |
| M | SACCR_* VBA result |
| N | Normalized difference |
| O | PASS/FAIL |
| P | Rule reference |

The column labeled **Rel. diff** uses `abs(Function - Expected) / max(1, abs(Expected))`. It is a normalized difference, not a conventional relative error when the expected magnitude is below 1. PASS requires it to be no greater than E4. The default 0.001% is 0.00001 as a decimal.

The expected formulas handle a negative duration start with a zero floor, a nonpositive multiplier add-on with a result of 1, and a positive V minus C before evaluating the exponential. These corrections prevent false failures at supported boundaries. Formula agreement remains a regression check, not an independent regulatory certification.

Randomized inputs are planned in [#165](https://github.com/danielep71/VBA-SACCR-Toolkit/issues/165); they are not delivered in the reviewed baseline.

## Outputs and upgrades

Results reports summary and status above its output table; detailed aggregation and trade results are available on their own sheets. Read Checks and validity as well as the numerical output.

The VBA synchronizer does not migrate worksheet cells, formulas, names, validation, or buttons. Rebuild from the template for layout changes and import compatible modules from the same commit. See [Getting Started](Getting-Started.md) for synchronization limitations.
