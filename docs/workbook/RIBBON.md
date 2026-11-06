# Ribbon

The workbook carries its own **K-SACCR** Ribbon tab (#213). The tab holds
every command, the status of the results and of both staging sheets, and a
way to reach every sheet. This document is the contract. The RibbonX source
that implements it is [`src/workbook/customUI14.xml`](../../src/workbook/customUI14.xml),
and `tools/test_ribbon_contract.py` checks that the two agree: the same
controls, callbacks and targets.

The Ribbon reaches the existing macros; it is not a new implementation.
Since #217 it is the only place for the calculation and export commands: the
only worksheet buttons left are those of the two staging sheets. Every command runs an `M_Main` macro, and that macro keeps
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
| Parameters | **Restore parameters**, which puts the reference parameters and the factor table back to the regulatory set (#357) |
| Status | Three read-only labels: Results, NettingSets staging, Trades staging |
| Navigate | Split button **Go to**: the primary click opens Home; the menu lists every sheet in six titled sections, in the order of the sheet tabs (#329) |

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
| `ksRestoreParameters` | Restore parameters | A | `RestoreRegulatoryParameters` |

A completed run leaves the user on Results; a run that stops, and Validate
inputs, leave the user on Checks (#589). Clear outputs keeps the active sheet.

**Restore parameters** (#357) writes the regulatory parameter set
`CRE52-CRR-2026-10-09` back to Params, from the engine's own definition of the
set, the one every run is compared with:

- **Restored:** the reference parameters `Alpha`, `MultiplierFloor`,
  `BusinessDaysPerYear`, `MinMaturityBD`, `SDFloorBD`, the three MPOR floors,
  `BasisFactor`, `VolatilityFactor`, the three IR correlations and the two
  CRR lambda thresholds; the conventions `DaysPerYear` (365) and
  `IRBucketOffset` (TRUE); and the 26 rows of the supervisory-factor table,
  from the key to the regimes.
- **Not changed:** `AsOfDate`, `ReportingCcy`, `Regime`, the FX table, the
  Reference column of the factor table, and the alpha overrides on
  NettingSets, which are inputs. A run with an alpha override stays `WHAT-IF`.
- **Confirmation:** it lists what will be overwritten, the value now and in
  the set, with **Cancel** as the default. When Params already holds the set
  it says so and changes nothing.
- **Write:** through the macros' protection, with no maintenance unlock. A
  failed write puts Params back as it was and checks it, as **Commit import**
  does. A factor row below the 26 rows of the factor table, right below
  them or further down above the FX table, refuses the restore, which would
  leave it there.
- **Effect:** Results reads `OUT OF DATE` until the next run, and the
  evidence export refuses until then. Automation calls
  `RestoreRegulatoryParameters_Silent`, which returns a result line.

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
| Reference and tests | `ksGoReadme` | README | `shReadme` |

`shHome` is the Home sheet (#330), the first sheet of the workbook. `RibbonGoTo` reports a target sheet the
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

Every command calls `RibbonGetEnabled` (#216). Like the labels, it reads
only the busy flag of `M_Main` (`OperationRunning`), Results C5 and the two
staging status cells; it never calculates or writes. The pure function
`EnabledFor` holds the table, and `TEST_MainState` checks it. The macros
enforce the same rules again, whatever the Ribbon shows.

| Control | Enabled when |
| --- | --- |
| Every control | No operation is running |
| `ksValidateStaging` | A staging sheet holds rows: its status is not `EMPTY` |
| `ksCommitImport` | A staging sheet reads `VALID` |
| `ksExportEvidence` | Results C5 says the inputs are unchanged since the last run: not `OUT OF DATE` and not `NONE` |
| Every other control | Always, apart from a running operation |

A status that cannot be read enables its control, and so does an error or a
call made while another state callback runs: a control is never disabled on a
guess, and the macro decides. The states carry no colour; the Status labels
give the state in words.

**Refresh.** The tab reads its labels and states again:

- after every operation, once Excel is idle (`Application.OnTime`), so that
  the labels read Results C5 after its recalculation: a command run from the
  Ribbon, a staging button or the Macros dialog, and the change and activation
  handlers of the staging sheets (`M_Main.RibbonStateChanged`, not in silent
  mode or while Excel is hidden, as under automation);
- after any cell edit in the workbook, once Excel is idle, since an input
  edit recalculates Results C5 (`ThisWorkbook.Workbook_SheetChange`);
- when the workbook or one of its sheets is activated, so switching
  workbooks refreshes it (`ThisWorkbook`);
- when the workbook is reopened in a running Excel, as Excel loads the tab
  again.

**After a VBA state reset.** An unhandled error or a code edit clears the
Ribbon object `RibbonOnLoad` kept. `RibbonOnLoad` therefore also stores the
object's pointer, with the window handle of this Excel session, in the hidden
workbook name `ksRibbonPointer`. After a reset, `RefreshRibbon` rebuilds the
object from that pointer with `RtlMoveMemory` of `kernel32`, a function that
is part of Windows; nothing is installed. The pointer is used only when the
handle is this session's; a negative pointer, as 32-bit Excel can have, is
read back as written. The name never marks the workbook as changed: it is
written with the `Saved` flag kept. It never reaches the file either:
`ThisWorkbook` removes it before every save and puts it back after
(`HidePointer`, `ShowPointer`), so a workbook never opens with the pointer of
an earlier session. Without a usable pointer the tab is simply not refreshed
until the workbook is reopened; the commands still work.
`?M_Ribbon.RibbonLoaded()` in the Immediate window tells whether the session
holds or can recover the object.

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
and control-to-sheet maps, the state table (`EnabledFor`) and the pointer's
format (`PointerText`, `PointerFromText`) are pure functions, so that tests
can check them against this contract; `IRibbonControl` cannot be created from
VBA. No
callback raises to Excel: a control with no command, or a command whose
macro `RibbonAction` does not call, is reported in a message box and runs
nothing.
