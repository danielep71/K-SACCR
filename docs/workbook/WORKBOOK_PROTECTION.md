# Workbook protection

Every sheet of the workbook is protected, without a password, and so is the
workbook structure (#420, #421). This document lists, sheet by sheet, the
cells a user may change in normal use and what else each sheet allows. Every
other cell is locked: formulas, results, headings, instructions, layout and
the reference regulatory parameters. Protection guards against accidental
edits. It is not a security boundary and does not replace input validation,
the freshness status or the output checks.

The machine-readable matrix is [`protection.json`](protection.json). The
`check_workbook_protection` gate checks the template
`src/workbook/SACCR_Template.xlsx` against it: the lock flag of every cell,
the saved protection of every sheet and of the structure, and the `PROTECT_`
lists of `CORE_Config` that the macros apply.

## 🔒 How it works

- **Saved protected.** The template, and so every workbook built from it, is
  saved with every sheet and the structure protected. The protection holds
  with macros disabled too; the macros themselves need macros enabled.
- **Macros write, users edit inputs.** When the workbook opens with macros
  enabled, `Auto_Open` runs `M_Main.ProtectWorkbook`, which protects every
  sheet again for the user interface only: the macros write every sheet,
  the user changes only the unlocked cells. Excel does not save that
  setting, so every operation (Run SA-CCR, Validate inputs, Clear outputs,
  Import CSV, Validate staging, Commit import) runs `ProtectWorkbook` first,
  also when the workbook was opened without `Auto_Open`. A sheet already
  protected this way is left as it is.
- **Structure.** Sheets cannot be added, deleted, renamed, moved, hidden or
  shown. No macro of the workbook needs to.
- **No password.** Review > Unprotect Sheet, or Review > Protect Workbook
  for the structure, removes the protection at once.

## 🔧 Maintenance unlock

To change a locked cell, for example a reference regulatory parameter on
Params, unprotect that sheet (Review > Unprotect Sheet), make the change and
either protect it again (macro `ProtectWorkbook`) or simply start the next
operation: every operation protects every sheet and the structure again
first. The change itself is kept, and a run on changed regulatory
parameters is still classified `WHAT-IF`; the freshness status and the
evidence checks work as before.

## 🗂️ Matrix

| Sheet | Role | Unlocked | Also allowed |
| --- | --- | --- | --- |
| `README` | Instructions | — | Opening and closing row groups |
| `Params` | Configuration and reference regulatory parameters | `E6` (AsOfDate), `E7` (ReportingCcy), `E23` (Regime), `C58:E257` (FX table and the rows below it for added currencies) | — |
| `NettingSets` | Input | Columns `C:S` from row 7 down | Sorting, inserting rows |
| `Trades` | Input | Columns `C:AD` from row 5 down | Sorting, filtering, inserting rows |
| `Import_NettingSets` | Staging | Columns `D:T` from row 10 down | Inserting rows |
| `Import_Trades` | Staging | Columns `D:AE` from row 10 down | Inserting rows |
| `Results`, `HedgingSets`, `Buckets`, `TradeCalc`, `Checks` | Output | — | — |
| `BankNettingSets` | Bank figures | `D4:D6` (amount scale and tolerances) and columns `C:P` from row 9 down | Sorting, inserting rows |
| `Reconciliation` | Output | — | — |
| `BankTrades` | Bank figures | `D4:D6` and columns `C:O` from row 9 down | Sorting, inserting rows |
| `TradeReconciliation` | Output | — | — |
| `TestCatalogue` | Test | `L7:L34` | — |
| `FormulaTest` | Test | `E4`, `E7:E28`, `F7:F28`, `G9:G28`, `H13:J21`, `K15` | — |
| `Basel_vs_CRR` | Reference | — | Filtering |

Selecting and copying any cell is always allowed. Deleting rows or columns,
formatting cells and inserting columns are not: to remove an input row, clear
its cells. Filtering works through the AutoFilter the template gives a sheet
(`Trades`, `Basel_vs_CRR`); a protected sheet cannot get a new one.

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
  Loading and clearing a staging block keep it that way.
- **Bank figures are inputs.** The bank's netting-set and trade figures, the
  amount scale and the two tolerances are editable; the reconciliation sheets
  built from them are outputs.
- **Yellow means editable.** No locked cell carries the yellow input fill.
- **Formulas stay visible.** No cell is hidden; a locked formula can still be
  selected, read and copied.

## 🧱 Limits

- Rows and columns cannot be deleted, and columns cannot be inserted,
  resized or formatted, on a protected sheet. To remove an input row, clear
  its cells; to change the layout, unprotect the sheet for maintenance.
- Filtering works only on `Trades` and `Basel_vs_CRR`, through the
  AutoFilter the template gives them, over the rows it covers.
- A sheet protected by hand from the Review tab, with other options, gets the
  workbook's options back at the next operation.
- Protection prevents accidental edits only. Anyone can remove it; the
  freshness status, the output fingerprint and the evidence export guards
  still tell whether results match their inputs.

## ⬆️ Upgrading an existing workbook

A workbook built from an earlier template has every cell locked and no
protection. Build a new workbook from the current template, as in
[INSTALLATION.md](../../INSTALLATION.md), and bring the inputs across: paste
values into the unlocked columns of `NettingSets`, `Trades` and the bank
sheets, or import them as CSV files through the staging sheets, and copy the
reporting date, reporting currency, regime and FX rates on Params. Run
SA-CCR once to rebuild the outputs.

<a id="manual-checks"></a>

## ✅ Manual checks

The Excel suites run every operation with the protection active, and
`TEST_MainState` checks the protection itself, but they act through the
macros, which protection does not restrict. These checks act as a user does.
For release evidence (#212), run them on the candidate workbook, built from
the template, and note each outcome in `session.log`
([EXCEL_EVIDENCE.md](../EXCEL_EVIDENCE.md#manual-procedure)).

| # | Check | Expected |
| ---: | --- | --- |
| 1 | Open the workbook with macros enabled; look at Review on every sheet | Every sheet offers **Unprotect Sheet**; **Protect Workbook** is on |
| 2 | Type into a header (`Trades` C4), a grey Params value (E8), an output cell and a staging metadata cell | Refused |
| 3 | Type into `Trades` C5, a `Trades` cell a few thousand rows down, `NettingSets` C7, Params E6 and an FX rate, a bank figure and a bank tolerance | Accepted |
| 4 | Paste several rows of values into the input columns of `Trades`, then clear them | Accepted |
| 5 | Filter `Trades` from its header; insert a row on `Trades` and on `NettingSets`; open and close a README group; select and copy output cells and a formula | Accepted; formulas visible |
| 6 | Right-click a sheet tab | Delete, Rename, Move, Hide are unavailable |
| 7 | Run SA-CCR, Validate inputs, Clear outputs, then Run SA-CCR again | Each succeeds |
| 8 | Import CSV on both staging sheets, correct a staged value, Validate staging, Commit import (once Cancel, once OK), Export import specifications, Run SA-CCR | Each succeeds; the staged value is editable, the CSV line column is not |
| 9 | Enter bank figures, Run SA-CCR, Export evidence | Reconciliation and export succeed |
| 10 | Unprotect Params, change a grey parameter, Run SA-CCR | The run is `WHAT-IF`; Params is protected again; restore the value |
| 11 | Unprotect Results, change an output cell, Export evidence, then Run SA-CCR | The export is refused as `EDITED`; the run rebuilds Results and protects it again |
| 12 | Save, close and reopen, then Run SA-CCR | Every sheet still protected; the run succeeds |
| 13 | Open with macros disabled | Every sheet and the structure are protected; nothing else works until macros are enabled |

## 🧭 Changing the matrix

A change to the editable cells or to what a sheet allows changes
`protection.json`, the template, the `PROTECT_` lists of `CORE_Config` where
it concerns them, and this document in the same pull request; the gate fails
when they disagree. A new sheet needs a row in the matrix, even when nothing
on it is editable.
