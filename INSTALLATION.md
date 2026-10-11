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

Repository setup, milestone **v0.0.1**, is complete. There is **no SA-CCR
release yet**: the VBA is the prototype engine imported from
`SACCR_Calculator.xlsm`, plus the regression harness. The
workbook is built from the macro-free template `src/workbook/SACCR_Template.xlsx`
and the VBA in `src/` and `tests/`.

| Topic | Status |
| --- | --- |
| VBA source layout | Defined in [`docs/REPOSITORY_STRUCTURE.md`](docs/REPOSITORY_STRUCTURE.md) |
| Import/export conventions and host support | Defined below and in [`docs/VBA_HOUSE_STYLE.md`](docs/VBA_HOUSE_STYLE.md) |
| Regression harness | `tests/modules/TEST_Harness.bas`; see [running the harness](#running-the-harness) |
| Excel evidence | Record and validator defined in [`docs/EXCEL_EVIDENCE.md`](docs/EXCEL_EVIDENCE.md) |
| Released versions | None |

The component list under [importing](#importing-vba-into-excel) grows as the
SA-CCR engine is added.

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
The exact component list is added here with the first VBA source.

### Building a workbook from source

1. Start from one exact commit in a Git checkout, so `.bas`, `.cls` and `.frm`
   files have CRLF line endings. Never mix components from different commits.
2. Open `src/workbook/SACCR_Template.xlsx` from the checkout and save it with
   **File → Save As** as **Excel Macro-Enabled Workbook (`.xlsm`)** outside the
   checkout, or in an ignored location such as `test-results/`. The template
   holds the 12 sheets, named ranges, formulas and buttons, and no VBA. Built
   workbooks are never committed.
3. Open the VBE (**Alt+F11**) and check **Tools → References** shows only the
   four default references listed under [supported hosts](#supported-hosts).
4. Import with **File → Import File**, in this order:
   1. every file in `src/core/`;
   2. every file in `src/classes/`;
   3. every file in `src/modules/`;
   4. every standard module (`.bas`) in `src/workbook/`, currently `M_Main`;
      the `.cls` document modules there are pasted, not imported (step 5);
   5. every `.frm` in `src/forms/` (its `.frx` loads with it; never import a
      `.frx`);
   6. for a development workbook only: `tests/` and `examples/` modules.

   Current components, in import order:

   | # | File | Component | Role |
   | ---: | --- | --- | --- |
   | 1 | `src/core/CORE_Config.bas` | `CORE_Config` | Internal; sheet layout and parameter constants |
   | 2 | `src/core/CORE_Engine.bas` | `CORE_Engine` | Internal; SA-CCR calculation run |
   | 3 | `src/core/CORE_Util.bas` | `CORE_Util` | Internal; conversions and sheet helpers |
   | 4 | `src/modules/SACCR_Formulas.bas` | `SACCR_Formulas` | Public worksheet functions `SACCR_*` (`docs/PUBLIC_API.txt`) |
   | 5 | `src/workbook/M_Main.bas` | `M_Main` | Sheet-button macros `RunSACCR`, `ValidateInputs`, `ClearOutputs` |
   | 6 | `tests/modules/TEST_Harness.bas` | `TEST_Harness` | Regression harness; development workbook only |

   The sheet buttons call `RunSACCR`, `ValidateInputs` and `ClearOutputs`, so
   they work once `M_Main` is imported.
5. **Document modules** in `src/workbook/` (`ThisWorkbook.cls` and sheet
   modules) cannot be imported: the VBE would create a new class such as
   `ThisWorkbook1`. Instead, open the `.cls` file in a text editor, copy the code
   below the `Attribute` lines, and paste it into the existing `ThisWorkbook` or
   sheet module. Sheet code names must match the file names. The document
   modules hold only `Option Explicit`.
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
results withdrawn; and it edits one trade after a run and checks that the
results are reported out of date until the edit is undone. Unlike the
harness, it runs the real macros, which rewrite the output sheets, so use a
development workbook. The seven protected-output scenarios cover each output
alone, TradeCalc + HedgingSets, Buckets + Results, and TradeCalc + Results +
Checks. They assert every writable table is empty, every failed clear is named
with its error number/source/description, the original error remains primary,
and a subsequent unprotected run succeeds. Another case injects a known primary
write failure and two cleanup failures with a different number and source,
proving secondary errors do not replace the original. A protected Results sheet may retain
its old table and A2 summary: the raised error (and Checks, when writable) must
say cleanup is incomplete. `ResultsStatus() = "NONE"` means no valid completed
run fingerprint, not proof that all output cells are empty. The tests preserve
A3, retaining the formula-based status implemented in PR #89 without a cleanup
write to that cell.

Run:

```text
TEST_MainState.RunMainStateTests
```

A passing run prints eighteen `CASE=` lines and ends with
`RESULT=PASS; cases=18; checks=...; failures=0; caller_state=RESTORED`.

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
`RESULT=PASS; cases=19; checks=37; failures=0; restore=PASS`. Illustrative
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
`RESULT=FAIL; cases=19; checks=37; failures=1; restore=PASS`. The injection
applies to that one suite only. A suite with no numeric check, such as
`TEST_InputValidation`, fails with a message saying so.

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
remargining. Every patched cell and name is restored. It needs a workbook
built from the template. Import it with `TEST_CaseRunner` and run
`TEST_InputValidation.RunInputValidationTests`; it ends with
`RESULT=PASS; cases=53; checks=60; failures=0; restore=PASS`.

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
cap is lower. Withheld netting sets must show no figures and the TOTAL row
must add up. It first checks that the portfolio has unmargined and margined,
CRR and BCBS netting sets, a cap that applies and one that does not, and a
multiplier below 1. A second run must write identical Results, TradeCalc,
HedgingSets and Buckets. With `IRBucketOffset` set to FALSE no IR add-on may
fall, and at least one must rise. Two opposite swaps, which give a zero
add-on, must give EAD = 1.4 * RC. Import it with `TEST_CaseRunner` and run
`TEST_Invariants.RunInvariantTests`; it ends with
`RESULT=PASS; cases=4; checks=21; failures=0; restore=PASS`.

For automation, `RunSACCR_Silent` returns a result line such as
`RESULT=OK; operation=run; errors=0; warnings=0; trades_used=62; trades_read=65; incomplete=2; total_ead=...; inputs=...; cleanup=PASS`
and raises any failure instead of showing a message box. `inputs` is the
fingerprint of the inputs the results were calculated from, also shown in the
run summary on Results A2. `M_Main.ResultsStatus()` returns `CURRENT` when the
results on the sheets match the inputs as they are now, `STALE` when an input
changed since, and `NONE` when there are no results. Results A3 shows the same
status through the formula `=ResultsStatusText(...)`, which Excel recalculates
when an input changes, and turns red when it reads `OUT OF DATE`. A formula,
unlike a macro writing to the sheet, leaves Excel's Undo history intact. With
calculation set to manual, press F9 to refresh it.

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
