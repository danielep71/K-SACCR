# 🪟 Excel Evidence

[![Execution: manual](https://img.shields.io/badge/execution-manual-1D76DB)](#manual-procedure)
[![Binding: exact SHA](https://img.shields.io/badge/binding-exact%20SHA-217346)](#what-a-record-binds)
[![Validator: offline](https://img.shields.io/badge/validator-offline-6f42c1)](#validating-a-bundle)

This document defines how a run of K-SACCR in Excel is recorded so that it can be
tied to one exact commit and checked. It is authoritative for the **evidence
record, the manual Windows/Excel procedure and the validator**. How to build the
workbook is in [`INSTALLATION.md`](../INSTALLATION.md#importing-vba-into-excel);
when evidence is required is in [`CONTRIBUTING.md`](../CONTRIBUTING.md) and
[`RELEASING.md`](../RELEASING.md).

> [!IMPORTANT]
> The validator, `tools/check_excel_evidence.py`, never runs Excel. It checks
> that a record is complete and consistent and that its sources and logs match
> the candidate commit. It cannot prove what a person did. `NOT_RUN`,
> `UNAVAILABLE`, static checks and synthetic fixtures are never an Excel pass.

Execution is **manual**: an operator runs Excel interactively. There is no
self-hosted runner, no automated Excel job, and the procedure never changes
macro security or Trust Center settings.

<a id="what-a-record-binds"></a>

## 🔗 What a record binds

An evidence bundle is a folder outside the checkout containing `host.json` and
the logs it cites. The record binds:

| Item | Bound to |
| --- | --- |
| Candidate | Full 40-character commit SHA; the import log must name it |
| Sources | Every `.bas`, `.cls` and `.frm` under `src/` and `tests/` at that commit, plus each form's `.frx` and the workbook template `src/workbook/SACCR_Template.xlsx`, with the SHA-256 of the exact bytes stored in Git. `examples/` is not imported |
| Policy | `.github/excel-evidence-policy.json` **at that commit**: repository, the harness entry point, ordered cases, assertion count and expected-error cases, the required references, and every other suite with its header, case count, check count per bitness and restoration field |
| Environment | Excel version and build, Office bitness, Windows version and architecture, VBA runtime, references, macro policy and VBA-project access as found, and `trust_changes: false` |
| Logs | Each cited log's SHA-256; a log changed after the record was written fails |
| Harness | For a passing or failed regression: exactly one normal-mode `TEST_Harness.RunTests` report whose `CANDIDATE`, `ENVIRONMENT`, `CASE`, count, `CLEANUP` and `RESULT` lines agree with the record and the policy |
| Suites | Every other suite the policy lists, in its order, each run on its own: one normal-mode report whose `CANDIDATE`, header, `CASE` and `RESULT` lines agree with the record and, for a pass, with the policy's counts and restoration field |
| Summary | Optional: one `TEST_Suite.Test_RunAll` summary whose line for each suite agrees with that suite's own report |

Source digests are of the bytes **stored in Git (LF)**, not of the CRLF files
in a Windows checkout, so do not compute them with `Get-FileHash`. Print them
with the validator:

```sh
python tools/check_excel_evidence.py --candidate-sha FULL_SHA --inventory
```

The policy changes together with the suites: adding a case or a check to a
test module updates `.github/excel-evidence-policy.json` in the same pull
request. `python tools/check.py` (gate `excel-evidence-policy`) fails when the
policy's counts differ from those the modules declare (`EXPECTED_CASES`,
`EXPECTED_CHECKS`, `CASES`, `CHECKS`, `BeginSuite`), when a header or `MODE`
line differs, or when `Test_RunAll` runs other suites. A module that declares
no check count (`TEST_MainState`) is held only by its evidence log.

<a id="manual-procedure"></a>

## 🧑‍💻 Manual procedure

Use Excel for Windows on a [supported host](../INSTALLATION.md#supported-hosts).
Python 3.10 or later is needed on the same machine, or on any machine with a
clone, to print the inventory and validate the bundle.

1. **Pick the candidate.** Fetch and check out the exact commit, and confirm
   the tree is clean:

   ```sh
   git fetch origin
   git switch --detach FULL_SHA
   git status
   ```

2. **Create the bundle folder** outside the checkout, for example
   `..\saccr-evidence-SHORTSHA\`. Start `session.log` there in Notepad and save
   it as UTF-8. Write the full candidate SHA and the start time with its UTC
   offset.
3. **Record the environment without changing it.** In `session.log`, note:
   - **File → Account → About Excel**: version, build and bitness;
   - **winver** and **Settings → System → About**: Windows edition, version and
     architecture;
   - **File → Options → Trust Center → Trust Center Settings**: the macro
     setting and whether *Trust access to the VBA project object model* is on.
     Do not change either.
4. **Build the workbook** exactly as in
   [Importing VBA into Excel](../INSTALLATION.md#importing-vba-into-excel): the
   template from this checkout saved as `.xlsm` outside the checkout, the four
   default references, every VBA component the `--inventory` output lists (all
   of `src/` and `tests/modules/`: a **development** workbook) imported in
   order from this checkout, with the document modules pasted as
   `INSTALLATION.md` describes, or with `tools/Sync-SACCR-VBA.ps1 -Mode
   Development`, then
   **Debug → Compile VBAProject**. Write in `session.log` the files imported, the references
   listed in **Tools → References** and the compile result.
5. **Run each suite once, on its own.** For the harness and for every suite in
   the policy, in its order: clear the Immediate window (**Ctrl+G**, then
   **Ctrl+A**, **Delete**), run the entry point (`TEST_Harness.RunTests`,
   `TEST_Formulas.RunFormulaTests`, and so on), and copy the whole
   Immediate-window output into its own log, saved as UTF-8: `harness.log`,
   `TEST_Formulas.log`, and so on. Start every log with the line
   `CANDIDATE=FULL_SHA`, typed before the pasted output, so that it cannot be
   cited for another commit. Keep a failing report as it is: never replace it
   with a later passing one. Do not put an injected-failure run in these logs.
   Optionally, run `TEST_Suite.Test_RunAll` last and keep its output, after a
   `CANDIDATE=` line, as `summary.log`; it must then agree with every suite's
   own log.
6. **Clean up.** Note the harness `CLEANUP=` line, close the test workbook, and
   record in `session.log` that it was closed and the finish time.
7. **Hash the logs**, after their last edit:

   ```sh
   python -c "import hashlib,sys;[print(hashlib.sha256(open(f,'rb').read()).hexdigest(), f) for f in sys.argv[1:]]" *.log
   ```

   PowerShell's `Get-FileHash -Algorithm SHA256 harness.log` gives the same
   value; this is correct for logs but not for sources.
8. **Write `host.json`** in the bundle from the [record template](#record-template),
   with the inventory from the command above and the observed values.
9. **Validate** as in [Validating a bundle](#validating-a-bundle) and keep the
   bundle with the summary. For setup and pull-request evidence, attach the
   zipped bundle and the summary to the issue or pull request.

If Excel cannot be run, write an `unavailable` record (below) and say so. Do
not fill an environment, stage or harness with values that were not observed.

<a id="source-round-trip"></a>

## 🔁 Source round trip

A workbook built from the template and the source of one commit must hold
exactly that source: the components the mode lists in
`tools/vba-projects.json`, each with the same code and attributes. The round
trip shows it (#34). It needs Windows Excel with *Trust access to the VBA
project object model* enabled, and Python on the same machine.

1. **Clean checkout of the candidate**, as in step 1 of the manual procedure,
   with `git status` clean.
2. **Build** the workbook from the checkout's template, saved as `.xlsm`
   outside the checkout, and synchronize it:

   ```powershell
   .\tools\Sync-SACCR-VBA.ps1 -WorkbookPath "C:\path\K-SACCR.xlsm" -Mode Production -IncludeSheets
   ```

   For a development workbook, use `-Mode Development`.
3. **Compile and run.** Open the workbook, run **Debug → Compile
   VBAProject**, then **Run SA-CCR** on the demonstration portfolio (and, in a
   development workbook, `TEST_Suite.Test_RunAll`). Save and close.
4. **Export** to a new folder; the workbook is opened read-only and not
   changed:

   ```powershell
   .\tools\Export-SACCR-VBA.ps1 -WorkbookPath "C:\path\K-SACCR.xlsm" -OutDir "C:\path\export"
   ```

5. **Compare** with the source of the candidate:

   ```sh
   python tools/compare_vba_export.py --export C:\path\export --mode production --candidate-sha FULL_SHA --workbook C:\path\K-SACCR.xlsm
   ```

   It prints the candidate, the template and workbook SHA-256 digests, the
   number of re-cased lines, one `FAIL` line per difference and
   `RESULT=PASS` or `RESULT=FAIL`. Code and the `Attribute` lines of standard
   modules must match. Ignored:
   - line endings, trailing blank lines and the `VERSION`/`BEGIN`/`END` block
     of a class file;
   - letter case in code outside string literals and comments. VBA names are
     case-insensitive and the VBE keeps one spelling per name across the
     project, so it re-cases identifiers on import. The lines that differ only
     that way are counted, not failed;
   - the attributes of the document modules. The synchronizer replaces only
     their code, and Excel writes their attributes.
6. **Retain** the comparison output, the compile and run results and the
   source digests from `check_excel_evidence.py --inventory`, for example in
   the issue or pull request that needs the evidence.

<a id="record-template"></a>

## 🧾 Record template

Replace every value with what was observed. This shape is not evidence.

```json
{
  "schema_version": 2,
  "repository": "danielep71/VBA-SACCR-Toolkit",
  "candidate_sha": "FULL_40_CHARACTER_SHA",
  "execution": "manual",
  "availability_reason": null,
  "started_at": "2026-10-06T14:00:00+02:00",
  "finished_at": "2026-10-06T14:20:00+02:00",
  "operator": "Name, workstation",
  "environment": {
    "excel_version": "16.0",
    "excel_build": "Version 2608 (Build 16.0.20326.20072)",
    "office_bitness": "64-bit",
    "os": "Windows 11 Pro 24H2 (build 26100.0000)",
    "os_architecture": "x64",
    "runtime": "VBA7+",
    "references": [
      "Visual Basic For Applications",
      "Microsoft Excel 16.0 Object Library",
      "OLE Automation",
      "Microsoft Office 16.0 Object Library"
    ],
    "macro_policy": "As found, for example: Disable VBA macros with notification",
    "vba_project_access": "As found, for example: Off; modules imported by hand",
    "trust_changes": false
  },
  "sources": "PASTE THE --inventory OUTPUT HERE (a JSON array)",
  "stages": {
    "import": {"status": "PASS", "detail": "Imported the inventory into the template saved as .xlsm", "log": {"path": "session.log", "sha256": "SESSION_LOG_DIGEST"}},
    "compile": {"status": "PASS", "detail": "Debug > Compile VBAProject completed with no error", "log": {"path": "session.log", "sha256": "SESSION_LOG_DIGEST"}},
    "regression": {"status": "PASS", "detail": "TEST_Harness.RunTests, complete report", "log": {"path": "harness.log", "sha256": "HARNESS_LOG_DIGEST"}},
    "cleanup": {"status": "PASS", "detail": "Harness cleanup PASS; test workbook closed", "log": {"path": "session.log", "sha256": "SESSION_LOG_DIGEST"}}
  },
  "harness": {
    "entry_point": "TEST_Harness.RunTests",
    "cases": 4,
    "assertions": 6,
    "failures": 0,
    "completeness": "COMPLETE",
    "expected_errors": [
      {"case": "cdo-delta.invalid-tranche", "status": "PASS", "detail": "Implied by the complete passing suite"}
    ]
  },
  "suites": [
    {"entry_point": "TEST_Formulas.RunFormulaTests", "status": "PASS", "cases": 9, "checks": 82, "failures": 0,
     "log": {"path": "TEST_Formulas.log", "sha256": "FORMULAS_LOG_DIGEST"}},
    "... one object for each suite in .github/excel-evidence-policy.json, in its order ..."
  ],
  "summary": {"path": "summary.log", "sha256": "SUMMARY_LOG_DIGEST"}
}
```

`summary` is `null` when `Test_RunAll` was not run.

Rules the validator applies:

- `excel_version`, `office_bitness` and `runtime` must equal the `version`,
  `office` and `runtime` values on the harness `ENVIRONMENT=` line, and that
  line's `os` must identify Windows.
- `references` must be exactly the policy's required references, in any order.
- A stage that did not run has `"status": "NOT_RUN"` and `"log": null`. Compile
  cannot run after a failed import, nor the regression after a failed compile.
  When the regression did not run or timed out, `harness` is `null`; when it
  did not run, no suite ran either.
- Every log of a test run starts with `CANDIDATE=` and the full candidate SHA;
  a log naming another commit, or none, fails.
- `suites` lists every suite of the policy once, in its order. A suite that
  ran is `PASS` or `FAIL` with the counts its `RESULT` line shows, and as many
  `CASE` lines as cases. A `PASS` needs the policy's case count and its check
  count for the recorded bitness, no failure and, where the policy names one,
  the restoration field (`restore=PASS`, `caller_state=RESTORED`); counts of an
  older commit fail. A suite that did not run is `NOT_RUN` with `null` counts
  and log; one abandoned as hung is `TIMEOUT` with `null` counts and a log
  without a `RESULT=` line.
- The regression stage is `FAIL` when the harness or a suite failed, and
  `PASS` otherwise; a suite that did not finish makes the record `INCOMPLETE`
  or `EXECUTION_TIMEOUT`, not `TEST_FAILED`.
- A summary needs every suite's own report; its line for each suite must show
  the same verdict and `RESULT` fields, and its `RESULT` line the right
  counts.
- A failed regression keeps the observed counts, with `completeness` set to
  `INCOMPLETE` when the harness stopped early but still printed its report. Its
  log must be one normal-mode report whose `CASE` lines follow the policy order
  and whose counts, `CLEANUP` and `RESULT=FAIL` lines match the record; a `FAIL`
  record cannot cite a passing report.
- A regression abandoned as hung is `TIMEOUT`: keep whatever it printed as its
  log, set `harness` to `null`, and claim no counts. The log is hashed but not
  parsed, and it must have no `RESULT=` line: a run that printed one finished
  and is recorded as `PASS` or `FAIL`.
- The expected-error result is inferred from the complete passing suite, which
  includes the assertion that an invalid CDO tranche returns `#NUM!`; it is not
  a separate observation.
- If the harness reports `cleanup=FAIL`, the cleanup stage must be `FAIL`.
- `environment.os` must start with Windows, and `office_bitness` must be
  exactly `32-bit` or `64-bit`.
- Every other environment string must be nonempty.
- `started_at` and `finished_at` must carry a UTC offset, and the finish may
  not precede the start.
- Log paths must be relative, stay inside the bundle and not pass through a
  symbolic link.

When Excel was not run, use the same identity and time fields with
`"execution": "unavailable"`, a nonempty `availability_reason`, `operator`,
`environment`, `harness` and `summary` set to `null`, `"sources": []`,
`"stages": {}` and `"suites": []`.

<a id="validating-a-bundle"></a>

## 🚦 Validating a bundle

From the checkout:

```sh
python tools/check_excel_evidence.py --candidate-sha FULL_SHA \
  --evidence ../saccr-evidence-SHORTSHA/host.json \
  --output test-results/excel-evidence.json --summary test-results/excel-evidence.md
```

The candidate commit must be present in the clone; the checkout does not need
to be on it. Reports must be written outside the bundle. Exit code 0 means the
record is a consistent pass, 1 a non-passing or invalid record, and 2 a command
or output error.

| Outcome | Meaning |
| --- | --- |
| `PASS` | Complete, consistent record; every stage and the expected-error result passed |
| `IMPORT_FAILED` | Import failed |
| `COMPILE_FAILED` | Compilation failed |
| `TEST_FAILED` | The harness or a suite reported a failure |
| `CLEANUP_FAILED` | Harness or operator cleanup failed |
| `EXECUTION_TIMEOUT` | A stage was abandoned as hung |
| `INCOMPLETE` | At least one stage did not run |
| `UNAVAILABLE` | No Excel run was attempted |
| `EVIDENCE_INVALID` | Missing, malformed, contradictory or stale-source record or logs |

A record can carry several outcomes, for example `COMPILE_FAILED, INCOMPLETE`.
Keep non-passing bundles too until the finding is resolved.

## 🧱 Limits

- Source and log digests are identity controls. They do not prove what a
  workbook contained or authenticate the operator.
- A pass covers the suites on the recorded host only. It validates SA-CCR
  numbers only as far as the cases do (published values; illustrative ones
  validate nothing), and no other Excel version or bitness.
- The `CANDIDATE=` line is typed by the operator: it stops a log of one commit
  from being cited, by mistake, for another, not a deliberate forgery.
- The summary of `TEST_Suite.Test_RunAll` is an index, not a report: it is
  checked against the suites' own logs and never replaces them.
- Keep logs free of credentials, client data and unredacted screenshots; see
  [`SECURITY.md`](../SECURITY.md#data-and-secrets).

## 🧪 Validator tests

`tools/test_excel_evidence.py` runs inside `python tools/check.py`
(`tool-tests`). It uses synthetic records and a temporary Git repository to
cover a valid record and wrong source identity, a different candidate,
logs without or with another `CANDIDATE=` line, a missing or reordered suite,
stale suite counts, suite records that contradict their logs, injected-mode or
repeated suite runs, a suite that did not restore its state, failed, not-run and
hung suites, summaries that contradict the suites' reports, the policy check
against the test modules,
incomplete and interrupted execution, failed cleanup, import, compile and test
failures, failure records that contradict their logs, timeouts that claim
results or cite a finished log,
injected-failure or repeated logs, tampered, missing, escaping or
symlinked logs, environment and reference mismatches, a non-Windows host,
unavailable records and
the command-line exit codes. Those tests are not Excel evidence.

The validator and policy are adapted from EXCEL-VBA-PROJECT-TEMPLATE at commit
`b903fe44ef6a032c4689870b83745afa1c22490d`. K-SACCR keeps manual execution only,
derives the source inventory from the repository layout instead of a profile
file, requires the default references, binds the harness `ENVIRONMENT` line
(including a Windows host), and
does not include the template's release-evidence binding.
