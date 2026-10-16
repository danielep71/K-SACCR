# K-SACCR tooling

Requirements: Git and Python 3.10+ on PATH. `check.py` needs no Python packages
and no Excel. Run from the repository root before every push:

```sh
python tools/check.py
```

All gates must pass; findings return a nonzero exit code. Reports and logs go to
the ignored `test-results/` directory. `--ci` checks committed whitespace instead
of local edits and requires a clean tree; add `--base <base-sha>` to inspect a
pull request's whole range.

| Gate | Checks |
| --- | --- |
| `tool-tests` | Fixtures for the source gate and public-API roles (`test_tooling.py`), regression cases from reviews (`test_review_regressions.py`), synthetic Excel evidence records (`test_excel_evidence.py`), test-case files (`test_test_cases.py`), the case-module generator (`test_generate_case_tests.py`), the export comparison of the source round trip (`test_compare_vba_export.py`) and the synchronizer's folder selection and document-module decoding in PowerShell 7 (`test_sync_script.py`, skipped without `pwsh`) |
| `check_source` | Every tracked text file is stored with LF in Git (CRLF, mixed or lone-CR blobs declared as text fail); VBA components sit only in the locations defined in `docs/REPOSITORY_STRUCTURE.md`, and modules in `src/core/` declare `Option Private Module`; exported VBA (`.bas`, `.cls`, `.frm`) checks out as CRLF, decodes as cp1252, has `Option Explicit`, a `VB_Name` equal to its filename and unique in the project and no string literal left open on its line (a compile error the VBE reports only on import; quotes after `'`, outside `[...]`, or after `Rem`, also behind a line number, are comment); each `.frm` references a tracked `.frx` large enough for its offset; `Sync-SACCR-VBA.ps1` and the template exist, every component lies in an active directory of the synchronizer's `$sourceRoots`, and the document modules match the template's sheets by CodeName; `CHANGELOG.md` starts with `## [Unreleased]`, uses `## [X.Y.Z] - YYYY-MM-DD` release headings with real calendar dates and has a link reference for each; once a release exists, `VERSION` holds the newest one |
| `check_committed_whitespace` | `git diff --check` on staged and unstaged changes (local) or on the committed range (`--ci`) |
| `check_vba_jumps` | Every `GoTo`, `GoSub`, `Resume` and `On Error GoTo` target is a label in the same procedure |
| `check_vba_conditionals` | `#If`/`#ElseIf`/`#Else`/`#End If` are balanced and use only `VBA6`, `VBA7`, `Win32`, `Win64`; `Declare` in reachable 64-bit branches is `PtrSafe`; no `#Const` |
| `check_test_cases` | Every fixture and expected file under `tests/` follows `docs/methodology/TEST_CASES.md`: envelope, declared fields with their types and units, quantity names and forms, trade references, tolerances, reference classes with their required fields and registered sources, the `illustrative-` naming rule and unique catalogue IDs; a registered EBA Q&A is cited by its source ID (such as `EBA-QA-2023_6962`), not in prose (`EBA Q&A 2023_6962` or `EBA Single Rulebook Q&A 2023_6962`), in the expected files, the VBA source and the methodology; only the source register's row that defines the ID may use the full name. It does not run any case |
| `check_vba_public_api` | Every `Public` declaration in `src/modules/` is listed, with its exact signature, in `docs/PUBLIC_API.txt`, and nothing else is; no implicit public procedures; one identifier per public `Const` or variable; no name collisions |
| `excel-evidence-policy` | `.github/excel-evidence-policy.json` matches the test modules: the harness's case and assertion counts, each suite's case and check counts (`EXPECTED_CASES`, `EXPECTED_CHECKS`, `CASES`, `CHECKS` or `BeginSuite`; a check count under `#If Win64` is the 64-bit one and under its `#Else` the 32-bit one; a module without a check count is held by its evidence log only), header and `MODE` line, a case-runner suite's first case (the first `BeginCase` or `StartCase` in its module), and the suites `TEST_Suite.Test_RunAll` runs, with `SUITE_COUNT`; and the root `README.md` shows the same suites and 64-bit counts in its Excel suites and Regression badges, the Testing capability, the suite table with its total, and the `Test_RunAll` output |
| `generated-case-tests` | `tests/modules/TEST_Cases.bas` is exactly what `tools/generate_case_tests.py` generates from the current fixtures and expected files |
| `*-fixtures` | Each VBA checker and the whitespace gate run their own positive and negative self-tests first |

`check_excel_evidence.py` validates a manual Excel evidence bundle against a
candidate commit: the harness and every other suite the policy lists, each with
its own log, and the optional `Test_RunAll` summary. `--inventory` prints the
source digests a record needs, and `--check-policy` is the
`excel-evidence-policy` gate above. Its synthetic tests, `test_excel_evidence.py`, run in `tool-tests`.
See [`docs/EXCEL_EVIDENCE.md`](../docs/EXCEL_EVIDENCE.md).

`Export-SACCR-VBA.ps1` and `compare_vba_export.py` are not gates either: they
make up the [source round trip](../docs/EXCEL_EVIDENCE.md#source-round-trip).
The script opens a built workbook read-only in Excel and exports every VBA
component to a new folder; the comparison checks that folder against the
source of one commit for a production or development workbook, and prints the
template and workbook digests. It ignores the letter case of code outside
strings and comments, which the VBE re-cases on import, and the attributes of
document modules, which Excel writes. Its tests run in `tool-tests`, without
Excel.

These are static checks only. None of them compiles VBA, opens Excel, runs a
test harness or validates any SA-CCR number, and a pass is never evidence that
the workbook works in Excel.

## Local VBA synchronization

`Sync-SACCR-VBA.ps1` synchronizes the repository VBA source into an existing
local `.xlsm` workbook without intentionally changing worksheets, cells,
formulas, names, tables or formatting.

For normal use on Windows, double-click `tools/Sync-SACCR-VBA.cmd`. The launcher uses the workbook path set in `DEFAULT_WORKBOOK` at the top of the file, runs the PowerShell synchronizer, shows the result and pauses before closing. Drag another `.xlsm` file onto the `.cmd` file to use that one instead.

PowerShell remains available for direct use from the repository root:

```powershell
.\tools\Sync-SACCR-VBA.ps1 -WorkbookPath "C:\path\to\SACCR.xlsm"
```

A full synchronization first asks which workbook it is building, unless
`-Mode Production` or `-Mode Development` gives the answer:

| Mode | Components | Use |
| --- | --- | --- |
| Production | The `production` configuration of `tools/vba-projects.json` (the five standard modules of `src/`) and its document modules | A workbook that calculates |
| Development | The `development` configuration: the production modules and every test module of `tests/modules/` | A workbook that runs the test suites, which write to the input and output sheets |

Any other answer cancels with exit code 2 and the workbook is not changed.
The manifest is the single list of both sets, shared with
`check_vba_calls.py`: a component in `src/` or `tests/modules/` that it does
not list stops the synchronization before Excel starts. A production
synchronization never removes modules; if the workbook already holds `TEST_`
modules, it says so and leaves them, and a production workbook should then
be built again from the template. `-Mode` cannot be combined with
`-SourceDir`, which imports what the folder holds.

Before it opens the workbook, the script lists the worksheet and
`ThisWorkbook` modules in the source and asks whether to replace their code
as well. Answer N to keep the sheet code of the workbook, for example one
built from a newer template; the standard modules are imported either way.
`-IncludeSheets` or `-SkipSheets` gives the answer in advance. When N leaves
nothing to import, the script exits with code 2 and the workbook is not
changed. The repository's sheet modules hold only `Option Explicit` and
comments, so skipping them changes no behaviour.

To test modules received outside the repository, for example downloaded
during a review, without copying them into the working tree, double-click
`tools/Sync-SACCR-VBA-Downloads.cmd`. It asks for the folder (Enter for
`%USERPROFILE%\Downloads`) and runs the synchronizer with `-SourceDir`:

```powershell
.\tools\Sync-SACCR-VBA.ps1 -WorkbookPath "C:\path\to\SACCR.xlsm" -SourceDir "$env:USERPROFILE\Downloads"
```

In that mode only the modules found in the folder are imported; every other
module in the workbook is left as it is. A file is identified only by the
`Attribute VB_Name` it declares, never by its file name: a file with no such
line, a malformed one, several (even identical) ones or a name that is not a valid
VBA name is listed as ignored with the reason. A file is taken when its
`VB_Name` is a component of the repository source, with the same extension. A
standard module (`.bas`) that the checkout does not have yet, such as a test
suite added by the pull request under test, is also taken when its name
starts with `CORE_`, `SACCR_`, `M_` or `TEST_`; it is listed apart as a new
module before the confirmation. Other files in the folder are ignored. When several files carry the same
`VB_Name`, such as `CORE_Engine.bas` and `CORE_Engine (1).bas`, the newest
is taken. The script lists the chosen files with their dates and the ignored
ones with the reason (also when no file is usable, before it stops), and asks
for confirmation (`-Force` skips it; answering no exits with
code 2 and leaves the workbook unchanged). The chosen files are copied to a
temporary folder with CRLF line endings, imported from there, and the copy
is deleted afterwards. The repository files are never changed, so the
working tree stays clean for merging. A later run of `Sync-SACCR-VBA.cmd`
brings the workbook back to the repository source.

The code of a document module is read as cp1252, the code page the VBE
exports in, so characters such as `£` or `é` in comments and strings
survive. `tools/test_sync_script.py` checks the folder selection, this
decoding, the answer to the production/development question and the
component set of each mode in PowerShell 7 (`pwsh`), without Excel; it is
skipped where `pwsh` is not installed.

The script reads the components the manifest lists for the chosen mode. Standard
modules are replaced from source; `ThisWorkbook` and worksheet document modules
keep their workbook objects and only their code text is replaced. It disables
Excel events, macro execution, external-link updates and automatic calculation
while synchronizing, checks that the sheet names/CodeNames are unchanged, and
creates a timestamped safety backup unless `-NoBackup` is supplied. As requested
by the owner, it deletes this temporary backup after a successful save and
retains it on failure. It is not a persistent rollback archive.

`check_source.py` keeps the synchronizer complete: it fails when the
synchronizer or the template is not tracked or missing, when a VBA component
lies outside the directories in `$sourceRoots`, when one of those directories
is missing, and when a document module in `src/workbook/` and the template's
sheets (by CodeName) do not match one to one. It reads only the active
entries of `$sourceRoots`, one `(Join-Path $repoFullPath "...")` per line:
commented-out entries do not count, and any other entry fails the check.
Comments are found as PowerShell reads them: a `#` or `<#` inside a string,
here-string or backtick escape such as `` `" `` is text. A new source
directory must be added to `$sourceRoots`; a new worksheet needs both its sheet
in the template and its document module.

Because it changes code only, a change to a sheet layout (a column or header)
does not reach an existing workbook. The engine's header check then stops the
run with the expected header; the CHANGELOG entry of such a change says which
cells to update, or the workbook is rebuilt from the template.

The synchronizer rejects known retired repository modules before changing any
VBA component. Rebuild such a workbook from the current template and source;
it does not migrate old module names or delete unrelated user components.

It requires Windows desktop Excel and Excel's **Trust access to the VBA project
object model** setting. The workbook VBA project must not be password-locked.

## VBA checkers: source and adaptations

`check_vba_jumps.py`, `check_vba_conditionals.py` and `check_vba_public_api.py`
come from EXCEL-VBA-PROJECT-TEMPLATE at commit
`b903fe44ef6a032c4689870b83745afa1c22490d`. K-SACCR now evaluates jumps per reachable compilation environment and checks
`PtrSafe` in the declaration modifier position. The
public-API checker takes each component's role from its folder, as defined in
[`docs/REPOSITORY_STRUCTURE.md`](../docs/REPOSITORY_STRUCTURE.md): `src/modules/`
is public; the other `src/` folders are internal; `tests/` and `examples/` are
test and example. The template instead reads a component registry from
`.github/repository-profile.json`, which K-SACCR does not use. `test_tooling.py`
adds fixtures for that role mapping. `test_review_regressions.py` covers the
post-merge findings, including colon-separated public declarations.

### Project-wide call resolution: advisory

```sh
python tools/check_vba_calls.py --root . --output test-results/check_vba_calls.json \
  --summary test-results/check_vba_calls.md
```

`check_vba_calls.py` is vendored from VBA-DATETIMEPICKER at commit `a66790a`
and adopted **advisory only** (#60). `check.py` does not run it. The static
workflow runs it in a non-blocking step that adds its summary to the job page.
Its fixtures (`test_vba_calls.py`) do run in `tool-tests`. It is not a VBA
compiler: compiling in Excel stays the authoritative call check.

**What it checks.** For each configuration in `tools/vba-projects.json`, it
checks only calls it can resolve with certainty, and does so in every
conditional-compilation environment of `check_vba_conditionals.py`:
- unqualified and module-qualified calls;
- members of project classes;
- Private members used from another module;
- ambiguous public names;
- argument counts and named arguments;
- literal `Application.Run`, `OnTime`, `OnKey`, `OnAction`, `CallByName` and `AddressOf` targets.

Anything else is reported as *unknown*, never as a defect.

**Configurations.**

| Configuration | Modules |
| --- | --- |
| `production` | The five modules of a workbook built from the template |
| `development` | Those plus the nine test modules, as in INSTALLATION.md |

The document modules (`ThisWorkbook` and the sheet modules) are listed under
`document_modules` and are not analysed. This key is a K-SACCR addition.

**Evidence at adoption** (release head `275dff6`):

| Configuration | Resolved calls | Findings | Unknown calls (host or library, untyped receiver, `With` block) |
| --- | ---: | --- | --- |
| `production` | 701 | none | 442, 388 and 89 |
| `development` | 2839 | none | 803, 745 and 91 |

Three defects were planted in a scratch copy to test detection:

| Planted defect | Result |
| --- | --- |
| A missing argument | Reported (`VBA-CALL-010`) |
| A Private engine member called from `M_Main` | Reported (`VBA-CALL-003`) |
| A misspelled internal procedure (`ComputeHedgingSet`) | Not reported |

**Limits.**
- **Undeclared names.** An unqualified call to an undeclared name is reported only when the name matches a project prefix (`CORE_`, `SACCR_`, `M_`, `TEST_`). The engine's internal procedures have no prefix, so a misspelled call to one is treated as a host or library call; the VBA compiler catches it.
- **Document modules.** Listed as classes, they produce false positives: their Excel members (`ThisWorkbook.Names`, `.Worksheets`, `.Date1904`) are unknown to the checker. This produced 9 false findings in production and 21 in development, which is why they are excluded.
- **Not checked:** dynamic call targets, `Object` and `Variant` receivers, and argument types.

**K-SACCR changes to the vendored file.** Type annotations for `mypy --strict`
(no logic change), the Ribbon parser copied from that repository's
`check_source.py`, and `document_modules`. Three upstream parser functions
exceed the Ruff complexity ceiling of 15. `pyproject.toml` exempts this one
file from C901, so it can still be resynchronised with upstream.

It does not become a blocking gate until its false-positive behaviour has been
validated over further changes.

## GitHub Actions

`.github/workflows/static-checks.yml` runs the same command with `--ci` on
pull requests targeting `main` or `release/**`, pushes to those branches, and
manual dispatch. It is installed on both `main` and the active release branch.
The job is named **Repository integrity**.

Pull requests pass their base SHA through `--base` to check the complete
proposed change. Checkout fetches the full history and verifies the exact
event SHA; for pull requests this is GitHub's test merge commit.
Push and manual runs use the runner's first-parent whitespace scope.

All gates must pass. Reports and logs from `test-results/` are uploaded
even when checks fail and retained for 30 days. The artifact name includes
the candidate SHA, run ID and attempt; `static-checks.json` also records the
candidate SHA. Missing evidence fails the upload step.

The workflow uses a read-only token, does not persist checkout credentials,
and pins actions to full commit SHAs taken from the template.
It uses GitHub-hosted Ubuntu and Python 3.10.
A successful job is static-check evidence only, not Excel validation.

After `check.py`, the same job runs three quality steps. Each runs even if an
earlier step failed, so one run reports every problem, and each failure fails the
job:

| Step | Checks |
| --- | --- |
| Ruff | `ruff check tools` with the rules in `pyproject.toml`; a probe proves the complexity ceiling of 15 is active |
| mypy | `mypy --strict` over every module in `tools/`, tests included; a probe proves an unannotated function is rejected |
| actionlint | Every workflow in `.github/workflows/`, including shellcheck on `run:` scripts |

Ruff and mypy install from `tools/requirements-quality-ci.txt` with
`--require-hashes --only-binary=:all:`. The actionlint archive is downloaded at a
pinned version and verified against its SHA-256 before use. Their outputs are in
the uploaded report. To run Ruff and mypy locally, with Python 3.10:

```sh
python -m pip install --require-hashes --only-binary=:all: -r tools/requirements-quality-ci.txt
ruff check tools
mypy
```

## Traffic history

`.github/workflows/daily-traffic.yml`, adapted from VBA-DATETIMEPICKER, runs
daily at 06:00 UTC and on manual dispatch, from the default branch only. It
reads GitHub's traffic API, which keeps just 14 days, and appends the snapshot
to CSV files under `data/` on the orphan `traffic-history` branch: totals,
daily views and clones, referrers and popular paths. The first run creates the
branch. It also opens an alert issue, assigned and labelled **P3**, on a views
or clones spike, a new star or fork, or a new referrer with at least five
views.

It needs a fine-grained personal access token with `Administration: read` on
this repository, stored as `TRAFFIC_TOKEN` in the `analytics` environment;
without it the run fails with a clear error. The badge files it writes under
`data/badges/` are shields.io endpoint files that a README can use, since the
repository is public. This is analytics, not CI: it runs no repository code
and must never be extended to.

## Dependency updates

`.github/dependabot.yml` asks Dependabot for weekly version updates of the
GitHub Actions pinned in the workflows. They are opened against the active
release branch (`target-branch`) and labelled `ci` and `P3`. Dependabot reads
this file from the default branch; PR #24 integrated it into `main`. Update `target-branch` whenever a new release branch is opened.

Review each update pull request like any other:

1. Check the new version's release notes and the commit it pins. The pin must
   stay a full 40-character SHA with the version as a trailing comment.
2. Let **Repository integrity** run, including actionlint, and confirm it is
   green.
3. Squash-merge it yourself. Nothing is merged or approved automatically, and
   Dependabot alerts are never dismissed automatically.

The quality tools are not covered by Dependabot. To change Ruff, mypy or
actionlint, update the version and its hash or checksum in
`tools/requirements-quality-ci.txt` or the workflow `env` in one reviewed pull
request.
Whether GitHub blocks merging on this check depends on the repository's
branch rulesets.
