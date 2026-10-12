<div align="center">

# 📦 Installation and Developer Setup

### Set up, validate, and later import SACCR into Excel

[![Status](https://img.shields.io/badge/Status-Pre--release-6e7781?style=flat-square)](#current-status)
[![Checks](https://img.shields.io/badge/Checks-python_tools%2Fcheck.py-0969da?style=flat-square)](#run-the-checks)
[![Security](https://img.shields.io/badge/Security-Private_policy-d73a49?style=flat-square)](SECURITY.md)

<br>

**One identifiable commit · Clean import · Compile · Validate · Preserve caller state**

</div>

---

This document is authoritative for **developer setup, import, upgrade, recovery
and removal**. Contribution workflow is owned by
[`CONTRIBUTING.md`](CONTRIBUTING.md), vulnerability handling by
[`SECURITY.md`](SECURITY.md), and publication by [`RELEASING.md`](RELEASING.md).

> [!IMPORTANT]
> VBA executes with the user's Office permissions. Review the exact source,
> follow organizational macro policy, and never enable macros in an untrusted
> workbook.

<a id="current-status"></a>

## 🧭 Current status

There is **no SA-CCR release yet**; work is on `release/1.0.0`. The workbook
is built from the macro-free template `src/workbook/SACCR_Template.xlsx` and
the VBA in `src/` (engine, worksheet functions, sheet macros) and
`tests/modules/` (regression harness and test suites).

| Topic | Status |
| --- | --- |
| VBA source layout | Defined in [`docs/REPOSITORY_STRUCTURE.md`](docs/REPOSITORY_STRUCTURE.md) |
| Import/export conventions and host support | Defined below and in [`docs/VBA_HOUSE_STYLE.md`](docs/VBA_HOUSE_STYLE.md) |
| Test suites | `tests/modules/`: the harness, the case runner and five suites; see [running the harness](#running-the-harness) |
| Excel evidence | Record and validator defined in [`docs/EXCEL_EVIDENCE.md`](docs/EXCEL_EVIDENCE.md) |
| Released versions | None |

## 🧰 Prerequisites

| Tool | Needed for |
| --- | --- |
| Git | Cloning and all repository work |
| Python 3.10 or later | `python tools/check.py`; standard library only, no packages |
| Node.js 20 or later | Optional: local label-catalogue checks |
| Microsoft Excel for Windows | Importing, compiling and running the VBA; see [supported hosts](#supported-hosts) |

<a id="supported-hosts"></a>

## 🖥️ Supported hosts

Decided by the owner on 2026-10-06 in issue #4.

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

## 📥 Get the source

Use a **Git clone**:

```sh
git clone https://github.com/danielep71/VBA-SACCR-Toolkit.git
cd VBA-SACCR-Toolkit
git switch release/1.0.0
```

`release/1.0.0` is the active development branch; `main` receives it only on
the owner's request.

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

Optional label-catalogue checks:

```sh
node .github/scripts/labels-sync.mjs --policy .github/label-policy.json --self-test
node .github/scripts/labels-sync.mjs --policy .github/label-policy.json --mode validate
node .github/scripts/labels-drift.mjs --policy .github/label-policy.json --self-test
```

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
   holds the 12 sheets, named ranges, formulas and buttons, and no VBA. Built
   workbooks are never committed.

   Instead of the manual import in steps 4 and 5, drag the saved workbook onto
   `tools/Sync-SACCR-VBA.cmd`: it imports every module from `src/` and
   `tests/modules/` and asks whether to fill the worksheet and
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
   6. for a development workbook only: `tests/` and `examples/` modules.

   Current components, in import order. Within each folder the order is
   alphabetical by file name, as `tools/Sync-SACCR-VBA.ps1` imports them:

   | # | File | Component | Role |
   | ---: | --- | --- | --- |
   | 1 | `src/core/CORE_Config.bas` | `CORE_Config` | Internal; sheet layout constants, parameter codes and error numbers |
   | 2 | `src/core/CORE_Engine.bas` | `CORE_Engine` | Internal; calculates EAD per netting set and writes the output sheets |
   | 3 | `src/core/CORE_Util.bas` | `CORE_Util` | Internal; value conversions, keyed lookups and sheet helpers |
   | 4 | `src/modules/SACCR_Formulas.bas` | `SACCR_Formulas` | Public worksheet functions `SACCR_*` (`docs/PUBLIC_API.txt`) |
   | 5 | `src/workbook/M_Main.bas` | `M_Main` | Sheet-button macros `RunSACCR`, `ValidateInputs`, `ClearOutputs` |
   | 6 | `tests/modules/TEST_Aggregation.bas` | `TEST_Aggregation` | Test suite; row-order independence and offsets; development workbook only |
   | 7 | `tests/modules/TEST_CaseRunner.bas` | `TEST_CaseRunner` | Shared runner for the case suites; development workbook only |
   | 8 | `tests/modules/TEST_Cases.bas` | `TEST_Cases` | Generated numerical test cases; development workbook only |
   | 9 | `tests/modules/TEST_Formulas.bas` | `TEST_Formulas` | Test suite; domain and error contract of the worksheet functions; development workbook only |
   | 10 | `tests/modules/TEST_Harness.bas` | `TEST_Harness` | Regression harness for the worksheet functions; development workbook only |
   | 11 | `tests/modules/TEST_InputValidation.bas` | `TEST_InputValidation` | Test suite; rejection of malformed inputs; development workbook only |
   | 12 | `tests/modules/TEST_Invariants.bas` | `TEST_Invariants` | Test suite; relations every result must satisfy; development workbook only |
   | 13 | `tests/modules/TEST_MainState.bas` | `TEST_MainState` | Test suite; Excel state and failures of the workbook macros; development workbook only |

   The sheet buttons call `RunSACCR`, `ValidateInputs` and `ClearOutputs`, so
   they work once `M_Main` is imported.
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

After importing the components above and compiling, open the Immediate
window (**Ctrl+G**) and run:

```text
TEST_Harness.RunTests
```

A passing run prints a `MODE=NORMAL` line, the environment, one `CASE=` line per
case, the counts, the cleanup verdict and finally:

```text
RESULT=PASS; completeness=COMPLETE; cases=4; assertions=6; failures=0; cleanup=PASS
```

Anything else is a failure, including a run that stops early: the harness
requires exactly 4 cases and 6 assertions before it can report `PASS`, and it
raises an error after printing a failed result.

To see the failure path, run:

```text
TEST_Harness.RunTestsWithInjectedFailure
```

It runs the same suite with one deliberately wrong expectation, in
`replacement-cost.exact`, and must print
`MODE=INJECTED_FAILURE` and
`RESULT=FAIL; completeness=COMPLETE; cases=4; assertions=6; failures=1; cleanup=PASS`,
followed by the suite failure error. If a run is interrupted, run
`TEST_Harness.ResetTests` before running again.

The harness never changes Excel settings. It checks that `Calculation`,
`DisplayAlerts`, `EnableEvents` and `ScreenUpdating` are the same after the run
as before, and reports `cleanup=FAIL` if not.

### Workbook macro state tests

`tests/modules/TEST_MainState.bas` checks that **Run SA-CCR**, **Validate inputs**
and **Clear outputs** put calculation mode, events and screen updating back as
they found them, also when they were off, and that an operation failure and a
cleanup failure are each raised and leave the workbook ready for the next run.
It also protects the Checks sheet for one run, and makes one run fail after
the first output sheets are written, and checks that each fails with its
results withdrawn; and it edits one trade, and on Params a factor's Regimes
field and its header, a parameter and an FX rate, after a run and checks that the results
are reported out of date until the edit is undone, once with the template
layout and once with a column inserted before the Params tables, which the
case deletes again (#152). Unlike the
harness, it runs the real macros, which rewrite the output sheets, so use a
development workbook. The seven protected-output scenarios cover each output
alone, TradeCalc + HedgingSets, Buckets + Results, and TradeCalc + Results +
Checks. They assert every writable table is empty, every failed clear is named
with its error number/source/description, the original error remains primary,
and a subsequent unprotected run succeeds. Another case injects a known primary
write failure and two cleanup failures with a different number and source,
proving secondary errors do not replace the original. A protected Results sheet may retain
its old table and C4 summary: the raised error (and Checks, when writable) must
say cleanup is incomplete. `ResultsStatus() = "NONE"` means no valid completed
run fingerprint, not proof that all output cells are empty. The tests preserve
C5, retaining the formula-based status implemented in PR #89 without a cleanup
write to that cell.

Run:

```text
TEST_MainState.RunMainStateTests
```

A passing run prints twenty `CASE=` lines and ends with
`RESULT=PASS; cases=20; checks=...; failures=0; caller_state=RESTORED`.

### Worksheet function contract tests

`tests/modules/TEST_Formulas.bas` calls the `SACCR_*` worksheet functions
directly, without the engine or the worksheet's data validation, and checks
their domain and error contract (#42): an input outside its documented domain
gives `#NUM!`, `#VALUE!`, `#REF!` or the error passed in, never a runtime
error or a misleading number; valid inputs at the boundaries keep their
results; single values, one-element arrays, vectors in either orientation and
2-D ranges are handled as documented. Its last case writes the same formulas
on a temporary worksheet, which it deletes again, so use a development
workbook. Run:

```text
TEST_Formulas.RunFormulaTests
```

A passing run prints nine `CASE=` lines and ends with
`RESULT=PASS; cases=9; checks=82; failures=0` on 64-bit Excel (80 on
32-bit, which has no `LongLong`).

<a id="numerical-test-cases"></a>

### Numerical test cases

`tests/modules/TEST_Cases.bas` is generated from `tests/fixtures` and
`tests/expected`; never edit it by hand. After changing a JSON file, run
`python tools/generate_case_tests.py` and import the regenerated module.
Import it with `tests/modules/TEST_CaseRunner.bas`, then run:

```text
TEST_Cases.RunCaseTests
```

Each case writes its fixture into the NettingSets and Trades rows and the
AsOfDate and ReportingCcy parameters, runs the engine and checks the outputs;
the other Params values must be the template's, except those a fixture sets in
its `parameters`, which are restored after the case. At the end the inputs and
parameters are written back, the engine is run once more and Excel settings
are restored. A passing run prints one `CASE=` line per expected file, the
results per reference class, and
`RESULT=PASS; cases=20; checks=39; failures=0; restore=PASS`. Illustrative
results are counted separately and validate nothing.

To see the failure path of any suite run through `TEST_CaseRunner`, arm the
injection on the same Immediate-window line as the suite:

```text
TEST_CaseRunner.InjectFailure: TEST_Cases.RunCaseTests
```

The suite moves the expected value of its first numeric check by ten times its
tolerance plus one, so the comparison must reject it. It prints
`MODE=INJECTED_FAILURE` and an `INJECTED=` line naming the check, then one
`FAILURE=` line for it; for `TEST_Cases` that is T01, so the summary shows
`PUBLISHED: passed=13; failed=1` and
`RESULT=FAIL; cases=20; checks=39; failures=1; restore=PASS`. The injection
applies to that one suite only. A suite with no numeric check, such as
`TEST_Aggregation`, fails with a message saying so.

### Invalid-input tests

`tests/modules/TEST_InputValidation.bas` writes inputs that a JSON fixture
cannot express (a trade row without an ID, a duplicate trade ID, a missing
MtM, an unreadable date, a typo in a netting-set flag or amount) through
`TEST_CaseRunner` and checks that each makes the netting set `INCOMPLETE` or
`INVALID`, and that blank optional fields still take their defaults. It
repeats a netting-set ID and puts `#N/A` in a trade, a netting set and a
parameter. It also renames a header, breaks a parameter name, duplicates a parameter, a
supervisory-factor key and a currency, and puts a blank row or a blank key in
the factor and FX tables, and checks that the run stops, or continues when the
duplicate has the same values. For aggregation it gives one credit reference
two sub-classes, in both orders, and adds a basis trade without a label, a
reference with a reserved character and an interest-rate risk factor that is
not a currency. For the CRR scope it checks that the sum of absolute IR
bucket values makes a CRR netting set `INVALID` but is accepted under BCBS, and
that an inflation risk factor such as `EUR-INFL` is accepted while another
suffix is rejected. For the netting-set terms it rejects VM held on an
unmargined netting set, VM posted on an unmargined CRR netting set, a clearing
role or margin-agreement code that is not N, CM, CLIENT, MIXED or SHARED, and
the MIXED and SHARED structures; it accepts VM posted on an unmargined BCBS
netting set and checks its replacement cost, and checks the MPOR of each
clearing role and of disputes and a large netting set with weekly
remargining, and the warning for a margined netting set of more than 5,000
trades whose large or illiquid flag is N, and checks that a broken Params,
NettingSets or Trades layout name stops the run. It writes values outside the
ranges of the data validation, which validation would not stop because VBA
writes them: a multiplier floor of 1, correlations above 1 or that do not form
a correlation matrix, an MPOR floor of 0, a fractional business-day count and
a negative supervisory factor stop the run; a negative threshold or MTA, a
remargining frequency or MPOR override that is not a whole number of business
days in range, and an alpha override of zero or less make the netting set
`INVALID`; a remargining frequency of 0 is raised to daily with a
warning. It checks that a run with the regulatory parameters is classified
`REGULATORY` and names the set in the run summary, that a DaysPerYear of 360
keeps it regulatory, that a different alpha, a zero supervisory factor or an
alpha override makes it `WHAT-IF` with the difference listed on Checks, and
that a reporting-currency FX rate other than 1 stops the run, as does a
reporting currency that is not three letters A to Z (digits, punctuation,
two or four letters, blank; lower case and spaces are accepted), and the
remargining frequency at its bounds (0.5 and 10001 rejected, 10000 accepted
with MPOR 10009). For economic
trades entered as legs (#38) it checks a worked two-category example (an
equity forward with an FX leg: RC 30, add-ons 3,200 and 400, EAD 5,082, and
the economic-trade and leg counts), a call spread entered as two option legs,
and the rejection of two MtM legs, a repeated MtM on a leg that is not the
MtM leg, no MtM leg, legs in different netting sets, a swap split into two
linear legs, a whole trade flagged N and a flag that is not Y or N. Every
patched cell and name is restored. It needs a workbook
built from the template. Import it with `TEST_CaseRunner` and run
`TEST_InputValidation.RunInputValidationTests`; it ends with
`RESULT=PASS; cases=97; checks=129; failures=0; restore=PASS`.

### Aggregation tests

`tests/modules/TEST_Aggregation.bas` runs the workbook's own portfolio in its
row order, reversed and rotated, and checks that Results is the same each time;
and it checks offsets by comparing outputs with each other: two opposite swaps
in one bucket offset exactly, and nothing offsets across currencies, between
interest rate and inflation in one currency, across netting sets or credit
entities, or between a standard and a basis hedging set. Import it
with `TEST_CaseRunner` and run `TEST_Aggregation.RunAggregationTests`; it ends
with `RESULT=PASS; cases=11; checks=11; failures=0; restore=PASS`.

### Invariant tests

`tests/modules/TEST_Invariants.bas` checks relations that every result must
satisfy, recomputed from the Results columns and the NettingSets inputs. On
the workbook's own portfolio it checks, for each `VALID` netting set, that the
aggregate add-on is the sum of the asset-class add-ons; RC follows its
margined or unmargined formula and is never negative; the multiplier follows
its formula, lies between the floor and 1 and is 1 when V - C is not
negative; PFE is the multiplier times the aggregate add-on;
EAD = alpha * (RC + PFE); and a margined netting set's EAD is the lower of the
margined EAD and the unmargined cap, with the cap flag set exactly when the
cap is lower. Withheld netting sets must show no figures. The TOTAL row must
hold, for V, C, RC, each asset-class add-on, the aggregate add-on, PFE and
EAD, the sum over the `VALID` netting sets, and nothing in its other columns.
It first checks that the portfolio has unmargined and margined,
CRR and BCBS netting sets, a cap that applies and one that does not, and a
multiplier below 1. A second run must write identical Results, TradeCalc,
HedgingSets and Buckets. With `IRBucketOffset` set to FALSE a CRR netting set
with interest-rate trades must become `INVALID` (Art. 280a(3) allows only the
formula with offsets); every other netting set `VALID` before must stay
`VALID`, no IR add-on may fall, and at least one must rise. Two checks then corrupt copies of the outputs, each TOTAL column in turn
and one netting set made invalid in the second run, and pass only if the
relations fail and name what was corrupted. Two opposite swaps, which give a zero
add-on, must give EAD = 1.4 * RC. Import it with `TEST_CaseRunner` and run
`TEST_Invariants.RunInvariantTests`; it ends with
`RESULT=PASS; cases=4; checks=23; failures=0; restore=PASS`.

For automation, `RunSACCR_Silent` returns a result line such as
`RESULT=OK; operation=run; errors=0; warnings=0; trades_used=62; trades_read=65; economic_trades_used=62; economic_trades_read=65; incomplete=2; total_ead=...; inputs=...; run_class=REGULATORY; cleanup=PASS`
and raises any failure instead of showing a message box. `inputs` is the
fingerprint of the inputs the results were calculated from, also shown in the
run summary on Results C4. `M_Main.ResultsStatus()` returns `CURRENT` when the
results on the sheets match the inputs as they are now, `STALE` when an input
changed since, and `NONE` when there is no valid completed run. Params is
read through its workbook names, so an edit is found also after its tables
have moved; descriptions, references and formatting do not count as edits.
Results C5 shows the same status through the formula
`=ResultsStatusText(...)`, which watches the input sheets, Params whole, and
Excel recalculates when an input changes, and turns red when it reads
`OUT OF DATE`. It reads `Status unavailable:` with the reason when a Params
workbook name is missing or broken. A formula, unlike a macro writing to the sheet, leaves Excel's
Undo history intact. With calculation set to manual, press F9 to refresh it.

<a id="validation-record"></a>

## 🧪 Validation record

A successful import is not certification. Record at least:

```text
Commit (full SHA):
Files imported:
Excel / Office version and build:
Office bitness:
Operating system:
Compile:
Regression harness:
Specific scenario:
Cleanup:
Skipped or unverified:
```

A skipped, incomplete or cleanup-failed run is not a pass. Static checks cannot
substitute for Excel execution.

When the run must be bound to an exact commit, for example for a release or an
issue's acceptance evidence, record it as an evidence bundle and validate it
with `tools/check_excel_evidence.py`; see
[`docs/EXCEL_EVIDENCE.md`](docs/EXCEL_EVIDENCE.md).

## ⬆️ Upgrade

Before moving to a newer commit:

1. read the changelog between the two commits;
2. back up the host and any user configuration;
3. identify the complete component set at the target commit; and
4. review compatibility notes and known limitations.

Replace the whole component set, compile again and repeat the validation. Do not
infer backward compatibility from successful compilation alone.

Treat a locally modified copy as a fork: diff it against the old and new source,
reapply modifications deliberately and retest them.

## 🧯 Troubleshooting

| Symptom | Check first |
| --- | --- |
| Compile error or missing procedure | All components come from one commit, and required references are set. |
| Ambiguous name | Remove duplicate or older copies of a component. |
| Form controls missing or corrupt | Re-import the exact `.frm` with its adjacent `.frx`. |
| Garbled accented characters after import | The module was copied from a browser or re-saved in another encoding; import the file from the checkout. |
| Different result from a reference | Confirm the commit, inputs, configuration and the independent reference used. |
| Security warning | Verify the source origin and organizational macro settings. |
| `tools/check.py` reports CRLF in Git | Renormalize with `git add --renormalize .` and commit the result separately. |

Suspected security problems follow [`SECURITY.md`](SECURITY.md), not issues.

## 🗑️ Removal

1. Run any project-owned shutdown or cleanup procedure.
2. Remove the components and integrations the project owns.
3. Compile the remaining VBA project.
4. Close and reopen, and verify nothing still depends on removed components.

Removing modules does not remove formulas, names, links, connections or other
host integrations. Remove only state the project owns.

## 📚 Related documents

- [`README.md`](README.md) — overview and status
- [`CONTRIBUTING.md`](CONTRIBUTING.md) — change and review workflow
- [`tools/README.md`](tools/README.md) — static checks and CI
- [`RELEASING.md`](RELEASING.md) — release sequence
- [`SECURITY.md`](SECURITY.md) — vulnerability and data-handling policy

---

**Installation principle:** use one identifiable commit, compile it, exercise
its real behavior in Excel, and record what was and was not validated.
