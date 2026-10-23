<div align="center">

# 📦 Installation and Developer Setup

### Set up, build, validate and upgrade K-SACCR in Excel

[![Release](https://img.shields.io/badge/Release-1.0.2-2ea44f?style=flat-square)](#current-status)
[![Checks](https://img.shields.io/badge/Checks-python_tools%2Fcheck.py-0969da?style=flat-square)](#run-the-checks)
[![Security](https://img.shields.io/badge/Security-Private_policy-d73a49?style=flat-square)](SECURITY.md)

<br>

**One identifiable commit · Clean import · Compile · Validate · Preserve caller state**

</div>

---

This document is authoritative for **developer setup, import, upgrade and
removal**. Contribution workflow is owned by
[`CONTRIBUTING.md`](CONTRIBUTING.md), vulnerability handling by
[`SECURITY.md`](SECURITY.md), and publication by [`RELEASING.md`](RELEASING.md).

> [!IMPORTANT]
> VBA executes with the user's Office permissions. Review the exact source,
> follow organizational macro policy, and never enable macros in an untrusted
> workbook.

<a id="current-status"></a>

## 🧭 Current status

The current release is **1.0.2**, tagged `v1.0.2`; its
[GitHub Release](https://github.com/danielep71/VBA-SACCR-Toolkit/releases/tag/v1.0.2)
carries a built production workbook ([released workbook](#released-workbook)).
Development continues on the
[active release branch](docs/GOVERNANCE.md#branches). The workbook is built
from the macro-free template `src/workbook/SACCR_Template.xlsx` and the VBA in
`src/` (engine, worksheet functions, sheet macros) and `tests/modules/`
(regression harness and test suites).

| Topic | Status |
| --- | --- |
| VBA source layout | Defined in [`docs/REPOSITORY_STRUCTURE.md`](docs/REPOSITORY_STRUCTURE.md) |
| Import/export conventions and host support | Defined below and in [`docs/VBA_HOUSE_STYLE.md`](docs/VBA_HOUSE_STYLE.md) |
| Test suites | `tests/modules/`: the harness `TEST_Harness`, eight suites (`TEST_Formulas`, `TEST_MainState`, `TEST_Cases`, `TEST_InputValidation`, `TEST_Aggregation`, `TEST_Invariants`, `TEST_Reconcile`, `TEST_CsvReader`), the shared case runner `TEST_CaseRunner`, `TEST_Suite`, which runs the harness and every suite, and the benchmark `TEST_Benchmark`, which is not a suite; see [running the harness](#running-the-harness) |
| Excel evidence | Record and validator defined in [`docs/EXCEL_EVIDENCE.md`](docs/EXCEL_EVIDENCE.md) |
| Released versions | 1.0.2 (2026-10-23), tag `v1.0.2`; 1.0.1 (2026-10-18), tag `v1.0.1`; 1.0.0 (2026-10-17), tag `v1.0.0` |

## 🧰 Prerequisites

| Tool | Needed for |
| --- | --- |
| Git | Cloning and all repository work |
| Python 3.10 or later | `python tools/check.py`; standard library only, no packages |
| Node.js 20 or later | Optional: local label-catalogue checks |
| Microsoft Excel for Windows | Importing, compiling and running the VBA; see [supported hosts](#supported-hosts) |
| Windows PowerShell, as shipped with Windows | `tools/Sync-SACCR-VBA.cmd`, which runs `tools/Sync-SACCR-VBA.ps1` through `powershell.exe`; not needed to import the VBA by hand |

<a id="supported-hosts"></a>

## 🖥️ Supported hosts

Decided on 2026-10-06 in issue #4.

| Host | Support level |
| --- | --- |
| Microsoft 365 or Excel 2016 and later, **Windows, 64-bit** | Supported target |
| Same versions, Windows, 32-bit | Best effort: kept compiling, not certified |
| Excel for Mac, Excel for the web, Excel 2013 and earlier | Not supported |

**References:** only the four defaults of a new workbook: *Visual Basic For
Applications*, *Microsoft Excel 16.0 Object Library*, *OLE Automation* and
*Microsoft Office 16.0 Object Library*. Adding any other reference needs an
issue and an update to this section.

**Evidence:** this table is a support **commitment**, not a test result. Host
evidence is recorded per change under [Validation record](#validation-record),
or as an evidence bundle bound to a commit as described in
[`docs/EXCEL_EVIDENCE.md`](docs/EXCEL_EVIDENCE.md).

<a id="portfolio-size"></a>

### Portfolio size

Measured with the [benchmark](#benchmark) on 2026-10-17, on the
supported 64-bit Windows host (#291): every leg was used and every netting
set was `VALID` at each size.

| Legs | Netting sets | Writing the inputs | Engine run |
| ---: | ---: | ---: | ---: |
| 1,000 | 2 | 0.04 s | 0.31 s |
| 10,000 | 20 | 0.28 s | 1.62 s |
| 50,000 | 100 | 1.85 s | 8.50 s |
| 500,000 | 1,000 | 15.58 s | 113.93 s |
| 1,000, used range inflated to row 105,005 | 2 | 0.03 s | 1.79 s |

- The engine has no fixed trade limit: 500,000 legs, calculated in under two
  minutes, is the largest size measured, not a cap. The run time grows about
  linearly with the legs (0.16 to 0.23 ms per leg here), and memory with it.
  Times depend on the machine.
- The 500,000-leg run used the benchmark at `8d14e30`, which differs from
  the merged module only in restoring the sheets' used range between
  scenarios; its used-range scenario, distorted by that leftover range, is
  not reported. The used-range row comes from the run of the merged module,
  in which 1,000 legs took 0.33 s without the stray row.
- These timings were measured on the five-class mix of #291: interest
  rates, FX, credit, equity and commodities. The benchmark now adds an
  other-risks leg to every 21 legs (#309); on that six-class mix it
  measured, on the same host on 2026-10-17 with the benchmark at `beff699`,
  1,000 legs in 0.06 s and 0.49 s (writing, engine run), 10,000 legs in
  0.34 s and 1.84 s, and 1,000 legs with the used range inflated to row
  105,005 in 0.03 s and 2.10 s, every leg used and every netting set
  `VALID`. The 50,000- and 500,000-leg sizes have not been rerun on it.
- The timings measure the engine: reading the inputs, calculating and
  writing all its outputs (TradeCalc, Buckets, HedgingSets, Results, Checks
  and the run information). They do not include what **Run SA-CCR** does
  around it, such as the reconciliation against bank figures, nor the
  evidence export. Their time on large portfolios has not been measured.
- The template's Trades dropdowns, input validation and formatting cover
  rows 5 to 5,000. Rows below are read and calculated the same way, but are
  entered without dropdowns or validation messages.
- A used range that reaches far below the data slows every run: the engine
  scans to its last row. Delete stray rows far below the last trade (deleting
  the cells, not only their contents) and save the workbook.

<a id="released-workbook"></a>

## 📦 Released workbook

The [1.0.2 release](https://github.com/danielep71/VBA-SACCR-Toolkit/releases/tag/v1.0.2)
carries `K-SACCR_1.0.2.xlsm`, a **production** workbook: the template and the
engine and sheet macros of `src/` at `v1.0.2`, without the test modules. It
holds the demonstration portfolio of the template.

1. Download `K-SACCR_1.0.2.xlsm` and `SHA256SUMS.txt` from the release.
2. Check the workbook's SHA-256 in PowerShell, in the download folder:

   ```powershell
   Get-FileHash .\K-SACCR_1.0.2.xlsm -Algorithm SHA256
   ```

   The hash must equal the line for `K-SACCR_1.0.2.xlsm` in `SHA256SUMS.txt`
   (letter case aside). If it differs, delete the file.
3. Windows marks a downloaded file as coming from the internet, and Office
   then blocks its macros. If your organization's macro policy allows it,
   right-click the file → **Properties** → tick **Unblock** → **OK**, or run
   `Unblock-File .\K-SACCR_1.0.2.xlsm`. Do this only after the hash check.
4. Open the workbook and enable macros as your policy allows.

To run the test suites, build a **development** workbook from source instead,
as below; the tests write to the input and output sheets.

<a id="get-the-source"></a>

## 📥 Get the source

Use a **Git clone**:

```sh
git clone https://github.com/danielep71/VBA-SACCR-Toolkit.git
cd VBA-SACCR-Toolkit
git switch release/1.1.0
```

The branch to switch to is the
[active release branch](docs/GOVERNANCE.md#branches); `main` receives it only
on the owner's request. For the released source, run
`git switch --detach v1.0.2` instead.

Do not use **Code → Download ZIP**. `.gitattributes` excludes repository
plumbing such as `.gitattributes`, `.gitignore` and `.editorconfig` from GitHub
archives, so a ZIP is not a complete checkout and the checks will not behave as
documented.

A clone checks VBA files out with the CRLF line endings the VBE expects. Do not
download individual modules through GitHub's **Raw** view: it serves the LF
form stored in Git.

<a id="run-the-checks"></a>

## ✅ Run the checks

From the repository root:

```sh
python tools/check.py
```

All gates must pass. Reports and logs are written to the ignored
`test-results/` directory. Before pushing a committed candidate, the stricter
form also requires a clean tree:

```sh
python tools/check.py --ci
python tools/check.py --ci --base <base-sha>   # for a pull request's full range
```

Optional label-catalogue checks are in [`docs/GOVERNANCE.md`](docs/GOVERNANCE.md#labels).
The **Repository integrity** job in CI also runs Ruff, mypy and actionlint,
which `check.py` does not; [GitHub Actions](tools/README.md#github-actions)
shows how to run Ruff and mypy locally.

The gates are described in [`tools/README.md`](tools/README.md). They are static
checks: they never compile VBA or run Excel.

<a id="importing-vba-into-excel"></a>

## 📂 Importing VBA into Excel

The layout and file conventions are defined in
[`docs/REPOSITORY_STRUCTURE.md`](docs/REPOSITORY_STRUCTURE.md) and
[`docs/VBA_HOUSE_STYLE.md`](docs/VBA_HOUSE_STYLE.md#export-compatibility).

### Building a workbook from source

1. Start from one exact commit in a Git checkout, so `.bas`, `.cls` and `.frm`
   files have CRLF line endings. Never mix components from different commits.
2. Open `src/workbook/SACCR_Template.xlsx` from the checkout and save it with
   **File → Save As** as **Excel Macro-Enabled Workbook (`.xlsm`)** outside the
   checkout, or in an ignored location such as `test-results/`. The template
   holds the 18 sheets, named ranges, formulas and buttons, and no VBA. Built
   workbooks are never committed.

   Instead of the manual import in steps 4 and 5, drag the saved workbook onto
   `tools/Sync-SACCR-VBA.cmd`, or pass its path to the launcher; it has no
   default workbook. It first asks which workbook you are building:
   **P**roduction, with the engine, the worksheet functions and the
   sheet-button macros, or **D**evelopment, which adds every test module of
   `tests/modules/`. The two component sets are listed in
   `tools/vba-projects.json`. Build a production workbook to calculate and a
   development workbook only to run the tests, which write to the input and
   output sheets. It then asks whether to fill the worksheet and
   `ThisWorkbook` document modules as well (they hold only `Option Explicit`,
   so either answer works). It needs Excel's **Trust
   access to the VBA project object model** setting; see
   [local VBA synchronization](tools/README.md#local-vba-synchronization).
3. Open the VBE (**Alt+F11**) and check **Tools → References** shows only the
   four default references listed under [supported hosts](#supported-hosts).
4. Import with **File → Import File**, in this order:
   1. every file in `src/core/`;
   2. every file in `src/classes/` (none at present);
   3. every file in `src/modules/`;
   4. every standard module (`.bas`) in `src/workbook/`, currently `M_Main`;
      the `.cls` document modules there are pasted, not imported (step 5);
   5. every `.frm` in `src/forms/` (none at present; its `.frx` loads with it;
      never import a `.frx`);
   6. for a development workbook only: every module in `tests/modules/`.

   Current components, in import order. Within each folder the order is
   alphabetical by file name, as `tools/Sync-SACCR-VBA.ps1` imports them:

   | # | File | Component | Role |
   | ---: | --- | --- | --- |
   | 1 | `src/core/CORE_Config.bas` | `CORE_Config` | Internal; sheet names, the workbook names of the tables, column numbers, parameter codes and error numbers |
   | 2 | `src/core/CORE_CsvReader.bas` | `CORE_CsvReader` | Internal; reads a Trades or NettingSets CSV file under the [CSV input contract](docs/csv/CSV_CONTRACT.md) and reports its findings; writes nothing |
   | 3 | `src/core/CORE_CsvSchema.bas` | `CORE_CsvSchema` | Internal; the CSV input contract for `CORE_CsvReader`, generated from `docs/csv/contract.json` |
   | 4 | `src/core/CORE_Engine.bas` | `CORE_Engine` | Internal; calculates EAD per netting set and writes the output sheets |
   | 5 | `src/core/CORE_Export.bas` | `CORE_Export` | Internal; writes the evidence report as CSV text and UTF-8 bytes |
   | 6 | `src/core/CORE_Reconcile.bas` | `CORE_Reconcile` | Internal; compares a bank's figures with Results and TradeCalc |
   | 7 | `src/core/CORE_Util.bas` | `CORE_Util` | Internal; value conversions, keyed lookups and sheet helpers |
   | 8 | `src/modules/SACCR_Formulas.bas` | `SACCR_Formulas` | Public worksheet functions `SACCR_*` (`docs/PUBLIC_API.txt`) |
   | 9 | `src/workbook/M_Main.bas` | `M_Main` | Sheet-button macros `RunSACCR`, `ValidateInputs`, `ClearOutputs`, `ExportEvidence`, `ExportTradesCsvTemplate` and `ExportNettingSetsCsvTemplate`; the CSV check macros `CheckTradesCsv` and `CheckNettingSetsCsv`; the [automation entry points](#automation-entry-points) `RunSACCR_Silent`, `ExportEvidence_Silent`, `ExportCsvTemplate_Silent`, `ReadCsv_Silent` and `ResultsStatus`; the status-cell worksheet functions `ResultsStatusText`, `ReconciliationStatusText` and `TradeReconciliationStatusText`; reads and writes the reconciliation sheets |
   | 10 | `tests/modules/TEST_Aggregation.bas` | `TEST_Aggregation` | Test suite; row-order independence and offsets; development workbook only |
   | 11 | `tests/modules/TEST_Benchmark.bas` | `TEST_Benchmark` | Benchmark of the engine on large synthetic portfolios; not a suite; development workbook only |
   | 12 | `tests/modules/TEST_CaseRunner.bas` | `TEST_CaseRunner` | Shared runner for the case suites and the benchmark; development workbook only |
   | 13 | `tests/modules/TEST_Cases.bas` | `TEST_Cases` | Generated numerical test cases; development workbook only |
   | 14 | `tests/modules/TEST_CsvReader.bas` | `TEST_CsvReader` | Test suite, generated from `tests/csv`; the CSV reader against the reference implementation; development workbook only |
   | 15 | `tests/modules/TEST_Formulas.bas` | `TEST_Formulas` | Test suite; domain and error contract of the worksheet functions; development workbook only |
   | 16 | `tests/modules/TEST_Harness.bas` | `TEST_Harness` | Regression harness for the worksheet functions; development workbook only |
   | 17 | `tests/modules/TEST_InputValidation.bas` | `TEST_InputValidation` | Test suite; rejection of malformed inputs; development workbook only |
   | 18 | `tests/modules/TEST_Invariants.bas` | `TEST_Invariants` | Test suite; relations every result must satisfy; development workbook only |
   | 19 | `tests/modules/TEST_MainState.bas` | `TEST_MainState` | Test suite; Excel state and failures of the workbook macros; development workbook only |
   | 20 | `tests/modules/TEST_Reconcile.bas` | `TEST_Reconcile` | Test suite; reconciliation of bank figures with Results; development workbook only |
   | 21 | `tests/modules/TEST_Suite.bas` | `TEST_Suite` | Runs every suite and prints a summary (`Test_RunAll`); print routine of the suites; development workbook only |

   The sheet buttons call `RunSACCR`, `ValidateInputs`, `ClearOutputs`,
   `ExportEvidence`, `ExportTradesCsvTemplate` and
   `ExportNettingSetsCsvTemplate`, so they work once `M_Main` is imported.
5. **Document modules** in `src/workbook/` (`ThisWorkbook.cls` and sheet
   modules) cannot be imported: the VBE would create a new class such as
   `ThisWorkbook1`. Instead, open the `.cls` file in a text editor, copy the code
   below the `Attribute` lines, and paste it into the existing `ThisWorkbook` or
   sheet module. Sheet code names must match the file names (the FormulaTest
   sheet keeps the code name `shTests`). The document modules hold only
   `Option Explicit`.
6. Run **Debug → Compile VBAProject**; it must complete with no error.
7. Save, close and reopen when a clean session is needed, then run the harness
   and the specific scenario under test.

Do not paste source into arbitrarily named modules: the component name
(`VB_Name`) is part of the contract, and `tools/check_source.py` requires it to
match the file name.

### Exporting changes back

1. In the VBE, right-click each changed component → **Export File**, and save
   it over its file in the checkout, in the same folder. For a document module,
   copy its code back into the matching `src/workbook/` file below the existing
   `Attribute` lines instead.
2. Review `git diff`. The VBE silently changes the capitalization of an
   identifier everywhere it appears when one declaration changes case; revert
   case-only changes you did not intend.
3. Run `python tools/check.py`. It verifies placement, `VB_Name`,
   `Option Explicit`, `Option Private Module`, form resources, encoding and
   line endings.
4. Commit. Git stores the files with LF; you do not convert anything by hand.

<a id="running-the-harness"></a>

## ▶️ Running the harness

To run every suite at once, type `Test_RunAll` in the Immediate window
(**Ctrl+G**). It runs `TEST_Harness`, `TEST_Formulas`, `TEST_MainState`,
`TEST_Cases`, `TEST_InputValidation`, `TEST_Aggregation`, `TEST_Invariants`,
`TEST_Reconcile` and `TEST_CsvReader` in turn and prints only a summary: one line per suite
with its verdict, counts and duration, the failure lines of a suite that
failed, and a final `RESULT=` line with the number of suites passed and
failed. An error raised by a suite counts as its failure and the next suite
still runs. The summary is not Excel evidence: the evidence record needs the
full log of every suite, each run on its own, and checks the summary against
them when it is added (see [`docs/EXCEL_EVIDENCE.md`](docs/EXCEL_EVIDENCE.md)).

Each suite below can also be run alone, with its full log, by typing its entry
point in the Immediate window. A passing run ends with a `RESULT=PASS` line
with `failures=0`; the expected case and check counts of every suite are in
`.github/excel-evidence-policy.json` and in the
[suite table](README.md#test-suites) of the README. The header comment of each
module in `tests/modules/` describes its cases in detail. Every suite except
the harness runs the real engine or writes to the sheets: use a development
workbook.

### Worksheet function harness

```text
TEST_Harness.RunTests
```

Runs the worksheet-function regression cases. It prints a `MODE=NORMAL`
line, the environment, one `CASE=` line per case, the counts and a cleanup
verdict. Anything but `RESULT=PASS` is a failure, including a run that stops
early: the harness requires every case and assertion to run before it can
report `PASS`, and it raises an error after printing a failed result.

`TEST_Harness.RunTestsWithInjectedFailure` runs the same cases with one
deliberately wrong expectation, in `replacement-cost.exact`, and must print
`MODE=INJECTED_FAILURE` and a `RESULT=FAIL` line with `failures=1`, followed
by the suite failure error. If a run is interrupted, run
`TEST_Harness.ResetTests` before running again.

The harness never changes Excel settings. It checks that `Calculation`,
`DisplayAlerts`, `EnableEvents` and `ScreenUpdating` are the same after the run
as before, and reports `cleanup=FAIL` if not.

### Worksheet function contract tests

```text
TEST_Formulas.RunFormulaTests
```

Calls the `SACCR_*` worksheet functions directly and checks their domain and
error contract: an input outside its domain gives an Excel error, never a
runtime error or a misleading number, and boundary values, arrays and ranges
behave as documented. Its last case writes the formulas on a temporary
worksheet and deletes it again. On 32-bit Excel the two `LongLong` checks are
skipped.

### Workbook macro state tests

```text
TEST_MainState.RunMainStateTests
```

Runs the real workbook macros. It checks that **Run SA-CCR**, **Validate
inputs** and **Clear outputs** put calculation mode, events and screen
updating back as they found them; that an operation failure and a cleanup
failure are raised separately and leave the workbook ready for the next run,
with the results of a failed run withdrawn; that the results are reported out
of date after an input edit, also after the Params tables have moved, and not
after a formatting change; that a run in manual calculation recalculates its
inputs first; that every output table is located through the workbook name of
its first header cell (`layout.output-header-names`, #253); and that
`RunSACCR_Silent` restores its silent flag (`silent.flag-preserved`).

Some cases protect output sheets, insert a column before the Params tables or
a row above the HedgingSets table, switch the workbook to the 1904 date system
or break the `Checks_Header` name; each change is undone, also after an
unexpected error, and the caller's Excel settings are restored. The last case
leaves the outputs of a normal run.

<a id="numerical-test-cases"></a>

### Numerical test cases

```text
TEST_Cases.RunCaseTests
```

`tests/modules/TEST_Cases.bas` is generated from `tests/fixtures` and
`tests/expected`; never edit it by hand. After changing a JSON file, run
`python tools/generate_case_tests.py` and import the regenerated module with
`tests/modules/TEST_CaseRunner.bas`.

Each case writes its fixture into the NettingSets and Trades rows, the
AsOfDate and ReportingCcy parameters and any parameter the fixture sets, runs
the engine and checks the outputs; the other Params values must be the
template's. At the end the inputs and parameters are written back, the engine
runs once more and Excel settings are restored. The results are counted per
reference class; illustrative results validate nothing.

To see the failure path of any suite run through `TEST_CaseRunner`, arm the
injection on the same Immediate-window line as the suite:

```text
TEST_CaseRunner.InjectFailure: TEST_Cases.RunCaseTests
```

The suite moves the expected value of its first numeric check by ten times its
tolerance plus one, so the comparison must reject it. It prints
`MODE=INJECTED_FAILURE`, an `INJECTED=` line naming the check and one
`FAILURE=` line for it, and ends with `RESULT=FAIL`. The injection applies to
that one suite only. A suite with no numeric check, such as
`TEST_Aggregation`, fails with a message saying so.

### Invalid-input tests

```text
TEST_InputValidation.RunInputValidationTests
```

Needs `TEST_CaseRunner` and a workbook built from the template. It writes
inputs that a JSON fixture cannot express: missing and duplicate IDs,
unreadable dates and codes, Excel error values, broken layout names, values
outside the ranges of the data validation, invalid netting-set terms, economic
trades entered as legs, and the date, IR bucket, FX pair and number-format
conventions. It checks that each rejects the trade, makes the netting set
`INCOMPLETE` or `INVALID`, stops the run, or is accepted with its default or
a warning, and that the run is classified `REGULATORY` or `WHAT-IF` as it
should be. Every patched cell and name, the number format and the date system
are restored.

### Aggregation tests

```text
TEST_Aggregation.RunAggregationTests
```

Needs `TEST_CaseRunner`. It runs the workbook's own portfolio in its row
order, reversed and rotated, and checks that Results is the same each time;
and it checks offsets: two opposite swaps in one bucket offset exactly, and
nothing offsets across currencies, between interest rate and inflation,
across netting sets or credit entities, or between a standard and a basis
hedging set. The inputs are restored.

### Invariant tests

```text
TEST_Invariants.RunInvariantTests
```

Needs `TEST_CaseRunner`. On the workbook's own portfolio it recomputes, from
Results and NettingSets, the relations every result must satisfy: the
aggregate add-on, RC, multiplier, PFE, EAD, the cap of margined netting sets
and the TOTAL row of each regime. It checks that a second run writes identical
outputs, that switching `IRBucketOffset` off makes a CRR netting set with
interest-rate trades `INVALID`, and that the relations catch deliberately
corrupted copies of the outputs. The inputs are restored.

### Reconciliation tests

```text
TEST_Reconcile.RunReconcileTests
```

Calls `CORE_Reconcile` with synthetic figures and no sheet, per netting set
and per trade leg: tolerances, amount scale, figures missing on either side,
duplicates, other regimes, withheld netting sets, excluded legs and unusable
rows. It then runs the macros on the workbook's own portfolio and checks the
Reconciliation and TradeReconciliation sheets and their status. Three cases
cover the evidence export: the CSV encoding of each field; in the workbook, an
export that is the same byte for byte when written twice and is refused after
an input, bank-figure or output edit; and a write that fails, or meets a
read-only file, leaving the older file as it was. A last case writes each
blank CSV template, checks it byte for byte against its header line, and
checks that a failed write leaves the older file as it was. The
BankNettingSets and BankTrades tables are put back and the exported files
removed.

### CSV reader tests

```text
TEST_CsvReader.RunCsvReaderTests
```

Generated by `tools/generate_csv_vba.py` from `tests/csv`: reads every CSV
fixture as written, with a byte-order mark and with CRLF line ends, and every
inline case of `tests/csv/cases.json`, through `CORE_CsvReader.ReadCsvBytes`,
and expects exactly the findings the reference implementation
`tools/csv_contract.py` pinned, the same record count and, for an accepted
file, every value as written. The cases cover invalid UTF-8, characters
outside the basic plane, spaces and control characters, formula starts,
number forms, digits and magnitudes, dates, keys, quoting, a comma or tab
delimiter and the record and size limits. Six more cases read a fixture from a file in `%TEMP%`, refuse a
file too large from its size, expect `ERR_CSV_READ` for a missing file, a
malformed path and an unknown kind, and check what `M_Main.ReadCsv_Silent`
returns for an accepted and a rejected file. It touches no sheet.

### Benchmark

```text
TEST_Benchmark.RunBenchmark
TEST_Benchmark.RunBenchmark "1000,10000,50000"
```

Not a suite: `Test_RunAll` does not run it and the Excel evidence does not
include it, because its timings depend on the machine. Needs
`TEST_CaseRunner` and a workbook built from the template. It writes synthetic
CRR portfolios of the given numbers of legs (1,000 and 10,000 by default; up
to 1,000,000, which needs memory and time in proportion), in netting sets of
500 legs, alternately unmargined and margined, with interest-rate swaps, FX
forwards and options, credit default swaps, equity forwards and options,
commodity forwards and other-risks forwards, all valid; then, last, a
portfolio of 1,000 legs with a
stray comment 100,000 rows below the end of the Trades used range, which
inflates the sheet's used range. For each it times the writing of the inputs
and the engine run, prints a `BENCH` line such as this one, with the seconds
in the locale's number format,

```text
BENCH case=bench-10000; legs=10000; netting_sets=20; used=10000; withheld=0; errors=0; warnings=0; tradecalc_rows=10000; results_rows=21; write_s=<seconds>; run_s=<seconds>
```

and checks that the run completed with every leg used, every netting set
`VALID`, no error on Checks, one Results row per netting set and the `TOTAL
CRR` row, and one TradeCalc row per leg; the used-range scenario also shows
`used_last_row` and checks that the comment did inflate the used range and
that deleting its row gives the used range back its extent. After every
scenario the rows it wrote below the sheets' original extent are deleted and
checked, so a large portfolio does not leave the next scenario, or the
workbook, with an inflated used range. The inputs are
restored. Sizes that cannot be read print a `BENCH_REJECTED=` line and change
nothing. The measured results are under [Portfolio size](#portfolio-size).

```text
TEST_Benchmark.RunCsvBenchmark
TEST_Benchmark.RunCsvBenchmark "10000,100000,100001"
```

Times the [CSV reader](docs/csv/CSV_CONTRACT.md#-the-workbooks-reader) on
synthetic files written to `%TEMP%`, under names not already taken, and
deleted: one Trades file per size (10,000 and 100,000 records, the contract's
limit, by default) and a NettingSets file of 10,000 records. Each prints a
`BENCH_CSV` line with the file's bytes, the seconds to build and write it and
the seconds to read it, and checks the outcome: accepted with every record
within the limit, rejected with `CSV-TOO-MANY-RECORDS` above it. A last
`BENCH_CSV_RESULT=` line says whether every outcome was as expected. It touches no sheet.

<a id="automation-entry-points"></a>

## 🤖 Automation entry points

`M_Main` has entry points for a caller that runs the workbook without the
buttons, such as a script or a test. None of them shows a message box.

### RunSACCR_Silent

`RunSACCR_Silent` runs **Run SA-CCR**, writes its messages to the Immediate
window and raises any failure to the caller. It returns a result line; for the
demonstration portfolio with the bank figures of the reconciliation
screenshot:

```text
RESULT=OK; operation=run; errors=5; warnings=0; trades_used=62; trades_read=65; economic_trades_used=62; economic_trades_read=65; incomplete=2; total_ead_crr=48595558.4974488; total_ead_bcbs=40784020.3398879; inputs=3B2754CA; run_class=REGULATORY; reconciliation=DONE; trade_reconciliation=NONE; cleanup=PASS
```

`RESULT=STOPPED` means the inputs could not be loaded; see Checks. `errors`
and `warnings` count the errors and warnings on Checks, and `incomplete` the
netting sets whose EAD was withheld. `reconciliation` is `DONE` when the bank
figures on BankNettingSets were compared, `NONE` when there are none,
`FAILED` when their settings or headers are unusable (the Reconciliation
summary in C4 says why), and `UNAVAILABLE` in a workbook built from a template
without the two sheets; `trade_reconciliation` says the same for BankTrades
and TradeReconciliation. `total_ead_crr` and `total_ead_bcbs` are the totals
of each regime's `VALID` netting sets, never added together. `inputs` is the
fingerprint of the inputs the results were calculated from, also shown in the
run summary on Results C4.

### ResultsStatus and the status cells

`M_Main.ResultsStatus()` returns `CURRENT` when the results on the sheets
match the inputs as they are now and the outputs are still those the run
wrote, `EDITED` when the inputs match but an output cell on Results, Checks or
a reconciliation sheet was changed after the run, `STALE` when an input
changed since, and `NONE` when there is no valid completed run; `NONE` does not prove that every
output cell is empty, as a protected sheet may keep old values after a failed
clear. Params is read through its workbook names, so an edit is found also
after its tables have moved; descriptions, references and formatting do not
count as edits. The fingerprint also covers the Trades and NettingSets
headers, the type of each value (the number 47756 and the text `'47756`
differ, as do two error values) and the workbook's date system; a header
changed only in case or spacing, which the schema check accepts, does not
count.

Results C5 shows the input status through the formula
`=ResultsStatusText(...)`, which watches the input sheets and Params whole, so
Excel recalculates it when an input changes. It checks the inputs only and
says so: `Inputs unchanged since the last run (inputs …); outputs are checked
when evidence is exported`. Watching the outputs as well would recompute their
fingerprint after every edit, for every row of a large portfolio; an edited
output is reported by `ResultsStatus()` as `EDITED` and refused by the
evidence export. C5 turns red when it reads
`OUT OF DATE`, and reads `Status unavailable:` with the reason when a Params
workbook name is missing or broken. A formula, unlike a macro writing to the
sheet, leaves Excel's Undo history intact. With calculation set to manual,
press F9 to refresh it. Reconciliation C5 does the same for the comparison
with the bank's figures through `=ReconciliationStatusText(...)`, which also
watches BankNettingSets: it reads `OUT OF DATE` once a bank figure, a setting
or an input changes after the run, and `No reconciliation` when there is none
or no current results. TradeReconciliation C5 does the same through
`=TradeReconciliationStatusText(...)`, which watches BankTrades.

### ExportEvidence_Silent

`ExportEvidence_Silent(filePath)` writes the evidence report without a dialog
and returns `RESULT=OK; lines=...; file=...`, or `RESULT=REFUSED; reason=...`
while Excel has a recalculation pending (manual calculation: press F9), when
`ResultsStatus()` is not `CURRENT`, when a reconciliation does not match its
bank figures, settings and inputs as they are now (also after every bank
figure was removed), or when an output cell on Results, Checks or a
reconciliation sheet differs from what the run wrote: a completed run stores
the fingerprint of the exported outputs in the hidden name `SACCR_RunOutputs`,
which every operation empties first, and the report gives it as
`RUN;outputs`.

The file is `;`-separated, with a `.` decimal point, CR LF line ends and UTF-8
without a byte-order mark; the first field of every line names its section
(`RUN`, `PARAMS`, `SF_TABLE`, `FX_TABLE`, `RESULTS`, `RECONCILIATION`,
`TRADE_RECONCILIATION`, `CHECKS`, `LIMITATIONS`), and the first line is
`RUN;format;K-SACCR evidence 1`. `RUN;reference_parameter_set` names the set
the template ships; whether the run used it unchanged (`REGULATORY`) or not
(`WHAT-IF`) is in `RUN;run_summary`. A reconciliation reads `CURRENT` or
`NOT MADE`; its summary says why one was not made. The workbook's file name
is not in the report, as a Save As would change it for the same run. Text
that Excel would read as a formula when the file is opened (starting with
`=`, `+`, `-`, `@`, a tab or a carriage return) is written with a leading
apostrophe; numbers keep their sign. Error values are written by their Excel
names (`#N/A`, `#SPILL!`, ...). A reconciliation's scale and tolerances are in
`RUN` only when it was made, as they are then part of its fingerprint.
`SF_TABLE` and `FX_TABLE` hold the rows the run read: from the header of
`SF_Table` or `FX_Table` down to the first blank key, also below the name's
range.

### ReadCsv_Silent

`ReadCsv_Silent(kind, filePath)` checks a CSV file of `kind` (`Trades` or
`NettingSets`) against the [CSV input contract](docs/csv/CSV_CONTRACT.md)
without a dialog and writes nothing. It returns
`RESULT=ACCEPTED; records=N; file=...`, or
`RESULT=REJECTED; findings=N; first=CODE|line|record|column; file=...` with
the first finding in the reader's order. An unknown kind and a file that
does not exist or cannot be read raise `ERR_CSV_READ`. The macros
`CheckTradesCsv` and `CheckNettingSetsCsv`, run from the Macros dialog, do
the same for a file chosen in a dialog and show the outcome with the first
ten findings; cancelling the dialog does nothing.

### ExportCsvTemplate_Silent

`ExportCsvTemplate_Silent(kind, filePath)` writes the blank CSV template of
`kind` (`Trades` or `NettingSets`) without a dialog and returns
`RESULT=OK; file=...`, or `RESULT=REFUSED; reason=...` for another kind. The
file is the template of the [CSV input contract](docs/csv/CSV_CONTRACT.md) in
`docs/csv/templates/`, byte for byte: the header line and a line feed, ASCII
without a byte-order mark. It is replaced as the evidence report is: written
whole beside the target first, so that a failed write leaves an older file as
it was, and a read-only file is not replaced. A file that cannot be written
raises to the caller.

<a id="validation-record"></a>

## 🧪 Validation record

A successful import is not certification: record each Excel run with the
[record template](docs/EXCEL_EVIDENCE.md#record-template) of
`docs/EXCEL_EVIDENCE.md`, which `tools/check_excel_evidence.py` validates
against the exact commit.

## ⬆️ Upgrade

Read the notes in the [changelog](CHANGELOG.md) for the changes since the
commit the workbook was built from. Then:

- **Only the VBA changed:** drag the workbook onto `tools/Sync-SACCR-VBA.cmd`
  again, or import the changed modules by hand as above, run
  **Debug → Compile VBAProject** and run the tests that apply.
- **The template changed** (its sheets, layout or workbook names): build a new
  workbook from `src/workbook/SACCR_Template.xlsx` as described under
  [building a workbook from source](#building-a-workbook-from-source), and
  copy your inputs into it. `git diff --stat <old-sha> HEAD --
  src/workbook/SACCR_Template.xlsx` shows whether the template changed. A run
  that stops with `Workbook name '…_Header' is missing or no longer refers to
  the first header cell, in column C, of … - rebuild the workbook from the
  template.` comes from a workbook built from an older template.

Treat a locally modified copy as a fork: diff it against the old and new
source, reapply modifications deliberately and retest them.

## 🧯 Troubleshooting

| Symptom | Check first |
| --- | --- |
| Compile error or missing procedure | All components come from one commit, and only the four default references are set. |
| Ambiguous name | Remove duplicate or older copies of a component. |
| A run stops with `Workbook name '…' is missing or no longer refers to …` | The workbook was built from an older template, or a name was deleted; build it again from `src/workbook/SACCR_Template.xlsx`. |
| A production workbook holds `TEST_` modules | A production synchronization never removes modules; build it again from the template. |
| Garbled accented characters after import | The module was copied from a browser or re-saved in another encoding; import the file from the checkout. |
| Results C5 still reads current after **Use 1904 date system** was switched | Excel need not recalculate C5 when only the date system changes, as no input cell changed. Press Ctrl+Alt+F9, which recalculates every formula, or edit an input: C5 then reads `OUT OF DATE`. A run under the 1904 date system stops with a message, so no result is produced from it; clear the option, check the dates and run again. |
| Different result from a reference | Confirm the commit, the inputs, the Params values and the reference used. |
| Security warning | Verify the source origin and organizational macro settings. |
| `tools/check.py` reports CRLF in Git | Renormalize with `git add --renormalize .` and commit the result separately. |

Suspected security problems follow [`SECURITY.md`](SECURITY.md), not issues.

## 🗑️ Removal

The built workbook is the only thing installed: nothing is registered with
Excel or Windows. Delete the `.xlsm`, the `*.pre-vba-sync.*` backups the
synchronizer wrote beside it, and any evidence files you exported. If you
enabled **Trust access to the VBA project object model** only to build the
workbook, turn it off again.

## 📚 Related documents

- [`README.md`](README.md) — overview and status
- [`CONTRIBUTING.md`](CONTRIBUTING.md) — change and review workflow
- [`tools/README.md`](tools/README.md) — static checks and CI
- [`RELEASING.md`](RELEASING.md) — release sequence
- [`SECURITY.md`](SECURITY.md) — vulnerability and data-handling policy

---

**Installation principle:** use one identifiable commit, compile it, exercise
its real behavior in Excel, and record what was and was not validated.
