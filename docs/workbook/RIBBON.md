# Ribbon

The workbook carries its own **K-SACCR** Ribbon tab (#213). The tab holds
every command, the status of the results and of both staging sheets, and a
way to reach every sheet. This document is the contract. The RibbonX source
that implements it is [`src/workbook/customUI14.xml`](../../src/workbook/customUI14.xml),
and `tools/test_ribbon_contract.py` checks that the two agree: the same
controls, callbacks and targets.

The Ribbon is a second way to reach the existing macros, not a new
implementation. Every command runs an `M_Main` macro, and that macro keeps
its protection, busy guard, error handling and every gate: staging
validation, commit and export eligibility. The Ribbon's enabled states help
the user. They never replace those gates.

## 🖥️ Host and packaging

- **Host:** Microsoft 365 or Excel 2016 and later, Windows 64-bit; 32-bit is
  best effort. The source uses the `customUI14` namespace
  (`http://schemas.microsoft.com/office/2009/07/customui`), which those hosts
  read. Icons are built-in Office icons (`imageMso`), so the workbook ships no
  image file.
- **Packaging:** the macro-free template carries the source as its
  `customUI/customUI14.xml` part (#214). The build's `SaveAs .xlsm` keeps
  it, and so do ordinary saves; the owner's spike of 2026-11-05 checked
  both in Excel. `tools/ribbon_template.py --write` embeds the source, and
  the `ribbon-template` gate of `tools/check.py` fails when the template's
  part differs from it.
- **Which workbook:** a tab embedded in a workbook shows only while that
  workbook is active, and its callbacks run in that workbook's VBA project.
  Every callback acts on `ThisWorkbook`, so another open workbook is never
  touched.
- **Macros disabled:** the tab still shows, but no callback runs: commands,
  states and navigation are unavailable. The workbook stays readable, and
  the links on Home (#330) and the sheet tabs reach every sheet.

## 🧭 The tab

| Group | Controls |
| --- | --- |
| Calculation | Split button **Run**: the primary click runs SA-CCR; the menu holds Run SA-CCR, Validate inputs and Clear outputs |
| CSV import | Menu **Import CSV** (NettingSets file, Trades file) · **Validate staging** · **Commit import** · menu **Templates** (NettingSets template, Trades template, Import specification) |
| Evidence | **Export evidence** |
| Parameters | **Restore parameters**, which restores the regulatory parameters (#357) |
| Status | Three read-only labels: Results, NettingSets staging, Trades staging |
| Navigate | Split button **Go to**: the primary click opens Home; the menu lists every sheet in six titled sections (#329) |

Labels on the large buttons are short, so that they wrap cleanly; the full
name is in the screentip. The tab's keytip is `KS`, so **Alt, K, S** opens it.

## 🔘 Commands

Each command calls `RibbonAction`, which runs the macro named here.

| Control ID | Label | Keytip | Macro |
| --- | --- | --- | --- |
| `ksRun` | Run | R (split button) | `RunSACCR` |
| `ksRunMenuRun` | Run SA-CCR | R | `RunSACCR` |
| `ksValidateInputs` | Validate inputs | V | `ValidateInputs` |
| `ksClearOutputs` | Clear outputs | C | `ClearOutputs` |
| `ksImportNettingSets` | NettingSets file... | I, N | `ImportNettingSetsCsv` |
| `ksImportTrades` | Trades file... | I, T | `ImportTradesCsv` |
| `ksValidateStaging` | Validate staging | V | `ValidateStaging` |
| `ksCommitImport` | Commit import | M | `CommitImport` |
| `ksTemplateNettingSets` | NettingSets template... | P, N | `ExportNettingSetsCsvTemplate` |
| `ksTemplateTrades` | Trades template... | P, T | `ExportTradesCsvTemplate` |
| `ksImportSpecification` | Import specification... | P, S | `ExportImportSpecification` |
| `ksExportEvidence` | Export evidence | E | `ExportEvidence` |
| `ksRestoreParameters` | Restore parameters | A | `RestoreRegulatoryParameters` (#357) |

`ksRunSplit`, `ksImportMenu` and `ksTemplatesMenu` hold the controls above
and call no macro themselves. The Templates menu writes blank CSV templates
or the import specification; **Export evidence** writes the evidence CSV of a
run. They are different operations.

## 🗺️ Navigation

Each entry calls `RibbonGoTo`, which activates the sheet with that CodeName.
A renamed sheet tab therefore cannot break an entry. The primary click of
**Go to** (`ksGo`) opens Home.

| Section | Control ID | Label | Sheet CodeName |
| --- | --- | --- | --- |
| — | `ksGo` | Go to | `shHome` |
| Start | `ksGoHome` | Home | `shHome` |
| Start | `ksGoReadme` | README | `shReadme` |
| Inputs | `ksGoParams` | Params | `shParams` |
| Inputs | `ksGoNettingSets` | NettingSets | `shNettingSets` |
| Inputs | `ksGoTrades` | Trades | `shTrades` |
| Staging | `ksGoImportNettingSets` | Import_NettingSets | `shImportNettingSets` |
| Staging | `ksGoImportTrades` | Import_Trades | `shImportTrades` |
| Outputs | `ksGoResults` | Results | `shResults` |
| Outputs | `ksGoHedgingSets` | HedgingSets | `shHedgingSets` |
| Outputs | `ksGoBuckets` | Buckets | `shBuckets` |
| Outputs | `ksGoTradeCalc` | TradeCalc | `shTradeCalc` |
| Outputs | `ksGoChecks` | Checks | `shChecks` |
| Reconciliation | `ksGoBankNettingSets` | BankNettingSets | `shBankNettingSets` |
| Reconciliation | `ksGoReconciliation` | Reconciliation | `shReconciliation` |
| Reconciliation | `ksGoBankTrades` | BankTrades | `shBankTrades` |
| Reconciliation | `ksGoTradeReconciliation` | TradeReconciliation | `shTradeReconciliation` |
| Reference and tests | `ksGoTestCatalogue` | TestCatalogue | `shTestCatalogue` |
| Reference and tests | `ksGoFormulaTest` | FormulaTest | `shTests` |
| Reference and tests | `ksGoBaselVsCRR` | Basel_vs_CRR | `shBaselVsCRR` |

`shHome` is the Home sheet of #330. `RibbonGoTo` reports a target sheet the
workbook does not have and changes nothing, so an entry can never act on
another sheet or workbook.

## 🚦 Status

Each label calls `RibbonGetLabel`, which only reads a status cell.

| Control ID | Reads | Shows |
| --- | --- | --- |
| `ksStatusResults` | Results C5 | `Results: CURRENT` when the inputs are unchanged since the last run, `Results: OUT OF DATE` when they changed, `Results: NONE` before any run |
| `ksStatusNettingSets` | Import_NettingSets D7 | `NettingSets staging: ` and the status: `EMPTY`, `LOADED`, `EDITED`, `VALID`, `INVALID`, `OUT OF DATE` or `COMMITTED` |
| `ksStatusTrades` | Import_Trades D7 | `Trades staging: ` and the same statuses |

`EDITED` results are not shown: detecting them means recomputing the output
fingerprint, which a Ribbon callback must not do. The evidence export still
refuses edited results on its own.

## ✅ Enabled states

Every command calls `RibbonGetEnabled`. Like the labels, it reads only
status cells and the busy flag of `M_Main`. The macros enforce the same
rules again, whatever the Ribbon shows. #216 implements this table and the
refresh; until then every command is enabled.

| Control | Enabled when |
| --- | --- |
| Every command | No operation is running |
| `ksValidateStaging` | A staging sheet holds rows |
| `ksCommitImport` | A staging sheet reads `VALID` |
| `ksExportEvidence` | Results C5 says the inputs are unchanged since the last run |
| Every other command | Always, apart from a running operation |

The labels and enabled states are refreshed at the end of every operation,
on a staging edit, when a sheet or the workbook is activated, and after a
VBA state reset (#216). Until then a command run from the Ribbon refreshes
them, and only while the session holds the Ribbon object, which Excel
passes to `RibbonOnLoad` when it loads the tab. `?M_Ribbon.RibbonLoaded()`
in the Immediate window tells whether it does.

## 🔌 Callbacks

All callbacks live in one standard module, `src/workbook/M_Ribbon.bas`
(#215, #216, #329), with the signatures Office requires:

```vb
Public Sub RibbonOnLoad(ByVal Ribbon As IRibbonUI)
Public Sub RibbonAction(ByVal Control As IRibbonControl)
Public Sub RibbonGoTo(ByVal Control As IRibbonControl)
Public Sub RibbonGetEnabled(ByVal Control As IRibbonControl, ByRef ReturnedVal)
Public Sub RibbonGetLabel(ByVal Control As IRibbonControl, ByRef ReturnedVal)
```

`IRibbonUI` and `IRibbonControl` come from the Microsoft Office Object
Library, already one of the workbook's four references. The control-to-macro
and control-to-sheet maps are pure functions, so that tests can check them
against this contract; `IRibbonControl` cannot be created from VBA. No
callback raises to Excel: a control with no command, or a command whose
macro the workbook does not have yet, is reported in a message box and runs
nothing. Until #357 delivers `RestoreRegulatoryParameters`, **Restore
parameters** says the command is not available yet; until #330 adds Home,
**Go to** says the workbook has no sheet for that entry.
