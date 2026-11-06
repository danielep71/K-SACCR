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
parameters is still classified `WHAT-IF`, except for `DaysPerYear` and
`IRBucketOffset`: they are conventions, which `CORE_Engine.ClassifyRun` does
not compare. The freshness status and the evidence checks work as before.

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
  values stays `WHAT-IF`, as today. The exceptions are `DaysPerYear` and
  `IRBucketOffset`, conventions that do not change the run class.
  **Restore parameters** on the K-SACCR tab puts them all back to the
  regulatory set without an unlock (#357).
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
For release evidence (#212), `tools/Certify-SACCR.ps1` runs them through
`tools/Check-SACCR-Protection.ps1` on a production workbook built from the
candidate, and writes each outcome to `protection-checks.txt` in the bundle
([one-command certification](../EXCEL_EVIDENCE.md#one-command-certification)).
The script can also be run on its own.

How the script checks them:

- **Macros disabled (2, 3, 4, 6, 14).** It opens the workbook with its macros
  disabled. Nothing then protects the sheets for the user interface only, so
  Excel refuses a write to a locked cell from automation as it refuses a
  user's typing. The script writes the cells of checks 2 and 3, pastes and
  clears the block of check 4, and tries to add, rename and delete a sheet for
  check 6.
- **Check 5** reads what Excel enforces: filtering and an AutoFilter on Trades,
  inserting rows on Trades and NettingSets, formulas not hidden, output cells
  copied. The README row groups are checked under check 1, with the macros.
- **Macros enabled (1, 7 to 13).** It opens the workbook following the Trust
  Center setting, runs `Auto_Open`, and runs the operations through the silent
  entry points of `M_Main`. It saves, closes and reopens the workbook where a
  check says so. Check 8 imports the import specification's example files.
- **Not automated.** The confirmation dialog of **Commit import** (check 8),
  which the silent entry point does not show, and the sheet-tab menu
  (check 6), whose effect, a protected workbook structure, the script tests.
  A person may still run any check by hand.

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
| 10 | Unprotect Params, change Alpha (E8), Run SA-CCR | The run is `WHAT-IF`; Params is protected again; restore the value |
| 11 | Unprotect Results, change an output cell, Export evidence, then Run SA-CCR | The export is refused as `EDITED`; the run rebuilds Results and protects it again |
| 12 | Save, close and reopen, then Run SA-CCR | Every sheet still protected; the run succeeds |
| 13 | Import CSV on both staging sheets and Validate staging until both read `VALID`; save, close and reopen; Commit import, then Cancel. Change a staged value; save, close and reopen; Commit import | After the first reopen both sheets still read `VALID` and Commit import asks for confirmation. After the second, the changed sheet reads `EDITED` and Commit import is refused: *Nothing was committed. The staged rows have no VALID validation that still matches them, the input sheets and Params.* (#208) |
| 14 | Open with macros disabled | Every sheet and the structure are protected; nothing else works until macros are enabled |

## 🧭 Changing the matrix

A change to the editable cells or to what a sheet allows changes
`protection.json`, the template, the `PROTECT_` lists of `CORE_Config` where
it concerns them, and this document in the same pull request; the gate fails
when they disagree. A new sheet needs a row in the matrix, even when nothing
on it is editable.
