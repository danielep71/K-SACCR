# Workbook editability

This document lists, sheet by sheet, the cells a user may change in normal use
of the workbook (#420). Every other cell is locked: formulas, results,
headings, instructions, layout and the reference regulatory parameters.
Protection guards against accidental edits. It is not a security boundary and
does not replace input validation, the freshness status or the output checks.

The machine-readable matrix is [`protection.json`](protection.json). The
`check_workbook_protection` gate checks the lock flag of every cell of
`src/workbook/SACCR_Template.xlsx` against it.

## 🚦 Status

The template carries the lock flags of the matrix; no sheet is protected yet,
so every cell can still be edited. Turning protection on, without a
password, keeping it on through builds and macros, and the maintenance unlock
come with #421; the tests of protected workflows with #422.

## 🗂️ Matrix

| Sheet | Role | Unlocked |
| --- | --- | --- |
| `README` | Instructions | — |
| `Params` | Configuration and reference regulatory parameters | `E6` (AsOfDate), `E7` (ReportingCcy), `E23` (Regime), `C58:E257` (FX table and the rows below it for added currencies) |
| `NettingSets` | Input | Columns `C:S` from row 7 down |
| `Trades` | Input | Columns `C:AD` from row 5 down |
| `Import_NettingSets` | Staging | Columns `D:T` from row 10 down |
| `Import_Trades` | Staging | Columns `D:AE` from row 10 down |
| `Results`, `HedgingSets`, `Buckets`, `TradeCalc`, `Checks` | Output | — |
| `BankNettingSets` | Bank figures | `D4:D6` (amount scale and tolerances) and columns `C:P` from row 9 down |
| `Reconciliation` | Output | — |
| `BankTrades` | Bank figures | `D4:D6` and columns `C:O` from row 9 down |
| `TradeReconciliation` | Output | — |
| `TestCatalogue` | Test | `L7:L34` |
| `FormulaTest` | Test | `E4`, `E7:E28`, `F7:F28`, `G9:G28`, `H13:J21`, `K15` |
| `Basel_vs_CRR` | Reference | — |

## 📏 Rules

- **Input columns are open to the bottom of the sheet.** An input column is
  unlocked through its column format, not only in the rows the template
  stores, so a pasted or appended row is editable however many rows the
  portfolio has. The header rows above stay locked.
- **Params separates configuration from reference parameters.** The reporting
  date, the reporting currency, the default regime and the FX rates are
  configuration and keep the yellow input fill. Every other value on Params
  is a reference regulatory parameter, filled grey: it is changed only on an
  unprotected sheet, by a deliberate maintenance unlock, and a run on changed
  values stays `WHAT-IF`, as today.
- **Staging keeps its provenance.** On `Import_NettingSets` and
  `Import_Trades` the staged values are editable; the CSV line column, the
  status, the source file and the other metadata above the block stay locked.
- **Bank figures are inputs.** The bank's netting-set and trade figures, the
  amount scale and the two tolerances are editable; the reconciliation sheets
  built from them are outputs.
- **Yellow means editable.** No locked cell carries the yellow input fill.
- **Formulas stay visible.** No cell is hidden; a locked formula can still be
  selected, read and copied.

## 🔧 Changing the matrix

A change to the editable cells changes `protection.json`, the template and
this document in the same pull request; the gate fails when the template and
the matrix disagree. A new sheet needs a row in the matrix, even when nothing
on it is editable.
