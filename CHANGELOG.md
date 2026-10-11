<div align="center">

# 📜 Changelog

### Release history for SACCR counterparty credit risk in Excel/VBA

[![Format](https://img.shields.io/badge/Format-Keep_a_Changelog-0969da?style=flat-square)](https://keepachangelog.com/en/1.1.0/)
[![Versioning](https://img.shields.io/badge/Versioning-SemVer-6f42c1?style=flat-square)](https://semver.org/spec/v2.0.0.html)
[![Dates](https://img.shields.io/badge/Dates-YYYY--MM--DD-217346?style=flat-square)](#date-and-version-rules)
[![Staging](https://img.shields.io/badge/Staging-Unreleased_first-d97706?style=flat-square)](#unreleased)
[![Contributing](https://img.shields.io/badge/Changes-Contribution_guide-2ea44f?style=flat-square)](CONTRIBUTING.md)

<br>

**User-visible history · Explicit compatibility · Reproducible evidence · Immutable releases**

</div>

---

All notable changes to **SACCR** are documented here.

This changelog follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/)
and [Semantic Versioning](https://semver.org/spec/v2.0.0.html). It records
released behavior and material unreleased changes; it is not a commit log, issue
tracker, or substitute for release evidence.

---

## 🧭 Maintenance policy

- Add material changes under **Unreleased** in the same pull request as the
  behavior or documentation they describe.
- Write from the user's perspective: describe the observable result, contract,
  compatibility impact, and migration need.
- Link the owning issue or pull request when it contains useful engineering
  detail.
- Record only validation actually performed. Static checks are never Excel
  evidence; state what was and was not run in Excel.
- Move Unreleased entries into a dated version section only when a release is
  explicitly prepared.
- Do not edit a published release entry except to correct a demonstrable factual
  or link error; annotate material corrections instead of rewriting history.

See [CONTRIBUTING.md](CONTRIBUTING.md) for change and evidence requirements and
[SECURITY.md](SECURITY.md) for private vulnerability reporting.

<a id="date-and-version-rules"></a>

### Date and version rules

| Rule | Standard |
|---|---|
| Version | `MAJOR.MINOR.PATCH`, without the leading `v` in headings |
| Release heading | `## [X.Y.Z] - YYYY-MM-DD` |
| Date | Gregorian calendar date in ISO `YYYY-MM-DD` format |
| Ordering | Unreleased first; released versions newest to oldest |
| Comparison | Unreleased → latest tag; each release → preceding tag |
| Patch | Backward-compatible correction or hardening |
| Minor | Backward-compatible capability |
| Major | Incompatible public-contract change |

The repository remains below `1.0.0` while its supported surface is still
forming. Pre-release status does not excuse undocumented breaking changes.

<details>
<summary><strong>Entry categories</strong></summary>

<br>

| Category | Use for |
|---|---|
| **Added** | New supported capabilities, APIs, files, or tests |
| **Changed** | Changes to existing behavior, contracts, tooling, or documentation |
| **Deprecated** | Supported behavior scheduled for removal |
| **Removed** | Removed capabilities or compatibility |
| **Fixed** | Corrected defects |
| **Security** | Safely disclosed security corrections |
| **Documentation** | Material documentation-only changes |
| **Validation** | Evidence actually produced |
| **Compatibility** | Upgrade or migration effects |
| **Known limitations** | Deliberate, unresolved boundaries |

Use only the categories needed by a release.

</details>

---

<a id="unreleased"></a>

## [Unreleased]

> Not yet released. Development takes place on the active release branch,
> `release/1.0.0`; changes to `main` require an explicit owner instruction.
> Repository setup (milestone v0.0.1) closed without a release.

### Added

- A margined netting set of more than 5,000 trades whose large or illiquid flag
  is N gives a warning on Checks (#39): the 20-day MPOR floor follows from the
  count in the previous quarter (CRE52.51(1); CRR Art. 285(3)(a)), which the
  workbook does not see. `TEST_CaseRunner.ExpectChecksLine` checks such a line;
  `TEST_InputValidation` has 54 cases and 62 checks.

- Value test for the Basel sum of absolute IR bucket values (#39):
  `illustrative-ir-sum-of-absolutes` runs the CRE99 Example 1 inputs under BCBS
  with `IRBucketOffset` = FALSE and expects an IR add-on of 625,153.16 USD and
  an EAD of 959,214.42 USD, from a hand calculation (illustrative). Fixtures
  accept `IRBucketOffset` in `parameters`, and `TEST_CaseRunner.SetParameter`
  writes Y and N as TRUE and FALSE. `TEST_Cases` has 20 cases and 39 checks.

- `docs/methodology/sources/` holds copies of the regulatory texts the
  methodology was checked against: the SA-CCR articles of the consolidated CRR
  (26.06.2026), Delegated Regulation 2021/931 as amended (25.05.2025), Basel
  CRE52 and CRE99 (effective 1 January 2023) and EBA Q&A 2023_6962, each with
  its address, retrieval date and SHA-256.
- A support matrix in `docs/methodology/README.md` lists, for v1.0.0, which
  methods, asset classes, instruments, netting-set and collateral structures,
  parameters and conventions are supported, which the user enters in a
  supported form, and which cannot be entered (#33).
- The traceability table in `docs/methodology/README.md` lists each
  implemented rule with its CRE52 and CRR references, its VBA procedures, its
  test cases and its status per regime (#44): Basel rules covered by a
  published CRE99 output are validated, others are exercised by a published
  EAD or not validated, and no CRR rule is validated in v1.0.0. Features that
  are not implemented are listed below the table.
- `TEST_CaseRunner.InjectFailure` proves that the case suites catch a wrong
  value (#44): the next suite moves the expected value of its first numeric
  check beyond its tolerance and must end with `RESULT=FAIL`, one failure and
  the workbook restored. Every runner suite now prints `MODE=NORMAL` or
  `MODE=INJECTED_FAILURE`, and a numeric failure line shows the expected value
  that was compared.
- `TEST_Invariants` checks relations every result must satisfy (#44), on
  the demo portfolio: the aggregate add-on is the sum of the asset-class
  add-ons, RC and the multiplier follow their formulas, PFE is the
  multiplier times the add-on, EAD = alpha * (RC + PFE), a margined netting
  set's EAD is the lower of the margined EAD and the unmargined cap, withheld
  netting sets show no figures, and the TOTAL row holds the sum over the
  `VALID` netting sets of V, C, RC, each asset-class add-on, the aggregate
  add-on, PFE and EAD, and nothing else (#104). It also checks that a second
  run gives identical outputs, that with the IR sum of absolute bucket values
  a CRR netting set with interest-rate trades becomes `INVALID` while every
  other netting set stays `VALID` with no lower IR add-on (#103), and
  that a zero add-on is calculated without failing. Two further checks
  corrupt copies of the outputs and pass only if the TOTAL and IR relations
  then fail and name the corrupted column or netting set. The checks compare
  outputs with each other and are reported as illustrative.
- An interest-rate add-on flowchart in `docs/assets/`
  (`SACCR_IR_AddOn_flow.svg` and a 2080 px PNG). It traces the steps from trade
  selection to the aggregate add-on, including the maturity factor, basis and
  volatility hedging sets and time buckets by end date, with Basel CRE52 and
  CRR references for each step.
- A source-controlled GitHub Wiki publication system keeps every maintained
  Wiki page under `docs/wiki/`, validates its catalogue and authority links,
  publishes the complete page set through `tools/Publish-SACCR-Wiki.cmd`, and
  records the exact source commit and page hashes in `Wiki-Source.json`.
  Hosted checks validate the source bundle and a separate observation detects
  published-Wiki drift.
- Repository-control files adopted from the Excel VBA project template:
  `.editorconfig`, `.gitattributes` and `.gitignore`. Exported VBA source is
  stored with LF in Git and checked out as CRLF; Office packages are binary and
  ignored unless an exact path is re-included.
- One local command, `python tools/check.py`, runs the static gates before every
  push: Git storage and line endings, exported VBA integrity (`Option Explicit`,
  `VB_Name`, form resources), changelog structure and whitespace. See
  [`tools/README.md`](tools/README.md).
- Automatic **Repository integrity** checks on pull requests and pushes to
  `main` and `release/**`, plus manual execution. Reports and logs identify the
  checked commit and are retained for 30 days, including failed checks.
- This changelog.
- Repository governance records the active release branch, the pull-request
  workflow, issue metadata requirements, the review checklist and the current
  GitHub controls.
- Twenty-label issue catalogue from the Excel VBA project template, with a
  workflow that reconciles live labels from `.github/labels.json` and a
  read-only daily drift check. See [`docs/LABELS.md`](docs/LABELS.md).
- Root project documents in the same form as the other Excel/VBA repositories:
  [`CONTRIBUTING.md`](CONTRIBUTING.md), [`INSTALLATION.md`](INSTALLATION.md),
  [`SECURITY.md`](SECURITY.md), [`CODE_OF_CONDUCT.md`](CODE_OF_CONDUCT.md) and
  [`RELEASING.md`](RELEASING.md), tailored to SACCR's pre-release state. The
  README now gives the project status, a getting-started path and a
  documentation map.
- Repository structure for the **application** profile:
  [`docs/REPOSITORY_STRUCTURE.md`](docs/REPOSITORY_STRUCTURE.md) records the
  profile decision, where core, facade, workbook glue, tests, fixtures, examples
  and methodology live, the dependency direction and the public-API boundary,
  with the [`docs/PUBLIC_API.txt`](docs/PUBLIC_API.txt) manifest.
  `python tools/check.py` now rejects VBA components outside those locations and
  core modules without `Option Private Module`.
- VBA conventions and host support:
  [`docs/VBA_HOUSE_STYLE.md`](docs/VBA_HOUSE_STYLE.md) defines naming, units and
  domains, error contracts, Excel-state cleanup that preserves the original
  error, and export rules (cp1252, LF in Git and CRLF on checkout, forms,
  workbook policy). [`INSTALLATION.md`](INSTALLATION.md) declares the supported
  hosts (Excel 2016+/Microsoft 365 on Windows, 64-bit supported, 32-bit best
  effort, default references only) and gives the step-by-step build-from-source
  and export procedure.
- Issue forms for bugs, features and documentation, and a pull-request
  template, adapted to SACCR and its label catalogue. Forms set the assignee,
  type label and a default priority; `docs/GOVERNANCE.md` documents the manual
  milestone step and a three-search triage check for issue metadata.
- Three VBA static checks from the template run in `python tools/check.py`,
  each with its own fixtures: jump targets stay within their procedure,
  conditional compilation is balanced and uses known symbols with `PtrSafe`
  declarations, and every public facade declaration matches
  `docs/PUBLIC_API.txt`. None of them compiles VBA or checks SA-CCR results.
- The **Repository integrity** job also lints the Python tooling with Ruff and
  strict mypy, from hash-locked versions, and the workflows with a
  checksum-verified actionlint. Dependabot proposes weekly GitHub Actions
  updates to the release branch for manual review; nothing merges
  automatically.
- A deterministic VBA regression harness, `tests/modules/TEST_Harness.bas`, run
  with `TEST_Harness.RunTests`. It supports exact, tolerance and expected-error
  assertions with stable case names, refuses to report `PASS` unless all
  expected cases and assertions ran, verifies Excel settings are unchanged, and
  prints a machine-readable `RESULT=` line. `RunTestsWithInjectedFailure`
  demonstrates the failure path. Its four cases exercise the engine's worksheet
  functions: replacement cost, maturity factor, the CDO-delta `#NUM!` result
  for an invalid tranche, and option-delta repeatability. The v0.0.1 setup
  scaffold it first tested (`CoreScaffold`, `SaccrScaffold`) is removed.
- Excel evidence bound to an exact commit:
  [`docs/EXCEL_EVIDENCE.md`](docs/EXCEL_EVIDENCE.md) gives the manual
  Windows/Excel procedure and record format, and
  `tools/check_excel_evidence.py` checks a record against the candidate's
  source digests, `.github/excel-evidence-policy.json`, the default references
  and the retained harness log, including failed reports. It
  reports import, compile, test, cleanup, incomplete and unavailable outcomes
  separately and never runs Excel.
- Methodology basis and numerical test-case format: the EU CRR is the baseline
  regime and Basel CRE52 is also supported, as recorded differences.
  [`docs/methodology/`](docs/methodology/README.md) holds the source register,
  assumptions, open decisions and traceability table;
  [`TEST_CASES.md`](docs/methodology/TEST_CASES.md) defines fixture and
  expected-result files, reference classes (published, independent,
  illustrative), tolerances, case categories and the completeness policy, with
  one clearly illustrative case. No SA-CCR formula is implemented.
- Versioning conventions in [`RELEASING.md`](RELEASING.md#versioning): a root
  `VERSION` file is created when the first release is prepared, and
  `python tools/check.py` then requires it to match the newest dated changelog
  release. Release branches are `release/X.Y.Z`; tags are annotated `vX.Y.Z`.
  Development continues on `release/1.0.0`.
- Daily traffic export, as in VBA-DATETIMEPICKER: GitHub's 14-day traffic data
  is kept on the orphan `traffic-history` branch, with alert issues on spikes
  and new referrers. See [`tools/README.md`](tools/README.md#traffic-history).
- Issue forms and the pull-request template rewritten in the style of the
  Excel VBA project template: bug and feature forms ask for the regime (CRR,
  Basel CRE52 or both), evidence and alternatives; the issue chooser links to
  installation help and the methodology; the PR template covers candidate
  identity, static checks, Excel evidence, regression coverage, risk,
  provenance and SA-CCR-specific review.

- Prototype SA-CCR engine imported from the owner's `SACCR_Calculator.xlsm`
  (engine v1.1.0): `CORE_Config`, `CORE_Engine` and `CORE_Util` in `src/core`, the ten
  `SACCR_*` worksheet functions in `src/modules/SACCR_Formulas` (added to
  `docs/PUBLIC_API.txt`), and the sheet-button macros in `src/workbook/M_Main`
  with the 13 document modules. It covers both regimes (CRR default, Basel
  CRE52 per netting set), the CRR other-risks class and the margined-EAD cap.
  The five standard modules are laid out in the house style (#60): module and
  procedure banners giving purpose, inputs, results and regulatory
  references, one commented declaration per line, and section comments. Their
  statements are unchanged from the prototype apart from `Option Private
  Module` in the three core modules, which the import had declared twice.
- Workbook template `src/workbook/SACCR_Template.xlsx` (#58): the prototype's
  12 sheets, formulas, named ranges and buttons with its VBA project and its
  document properties (`docProps/core.xml`, `docProps/app.xml`) removed;
  `tools/check_source.py` rejects a committed workbook that carries either.
  The workbook is built by saving it as `.xlsm` and importing the source; the
  template is part of the Excel evidence source inventory. Verified in Excel by
  the owner (#58): a workbook built from it with the repository source ran the
  harness and Run SA-CCR with the prototype's results.

- The prototype TestCatalogue's 28 checks as versioned test cases (#59): 12
  fixtures in `tests/fixtures` with expected files per regime in
  `tests/expected`, mapped in `tests/README.md`. The 14 figures printed in BCBS
  279 Annex 4 are `published`; the 14 that came from an unavailable Python
  script or unreviewed hand calculations are `illustrative`. Each value and
  tolerance matches the catalogue. `docs/methodology/TEST_CASES.md` defines
  the full fixture vocabulary, trade-level and text results, and currencies,
  and `tools/check_test_cases.py` validates the files in `python
  tools/check.py`.

- The numerical test cases run against the engine in Excel (#44, decision 7):
  `tools/generate_case_tests.py` generates `tests/modules/TEST_Cases.bas` from
  the JSON files, checked in `python tools/check.py`, and
  `tests/modules/TEST_CaseRunner.bas` writes each fixture into the input sheets,
  runs the engine, compares the outputs and restores the workbook. Results
  are reported per reference class.

### Changed

- Every parameter value on Params (E6 to E25) has data validation with an
  input message giving its meaning, regulatory value and reference, and a
  Stop alert with the accepted range: dates 2000 to 2100, a three-letter
  currency, alpha above zero, a multiplier floor below 100%, day counts,
  maturity and MPOR floors in business days, factors, IR correlations
  that form a valid correlation matrix, TRUE/FALSE, CRR/BCBS and the
  lambda thresholds. The supervisory-factor table validates unique keys,
  asset class, factor, correlation, option volatility, commodity hedging
  set and regimes; the FX table validates unique three-letter currencies
  and positive rates, with 1 for the reporting currency. Validation stops
  typed values only; the engine's own checks still apply to pasted values.
  NettingSets (rows 7 to 206) has the same kind of validation: unique IDs,
  Y/N flags, the clearing role, whole business days for the remargin
  frequency and the MPOR override, numbers for VM and NICA, TH and MTA not
  negative, alpha above zero, the regime and the margin-agreement
  structure. The Trades NettingSetID drop-down now explains itself and warns
  when an ID is not on NettingSets. Validation does not apply to pasted
  values, so a repeated netting-set ID, supervisory-factor key or FX
  currency is also shown in red by conditional formatting, however it was
  entered.
- NettingSets and Trades are located through workbook names, as Params is:
  `NettingSets_Header` and `Trades_Header` name the first header cell of each
  table (in column C), and the data start on the next row, so their rows can
  move (`CORE_Util.TableHeaderRow`, `TableFirstRow`). A missing or broken name
  stops the run. The test runner writes, clears and restores the input rows
  through the same names. Trades takes the owner's final layout: title band
  in B1:B2, description in C3, a hint row on row 4, headers on row 5 and data
  from row 6; the Run SA-CCR, Validate inputs and Clear outputs buttons sit on
  row 3 to the right of the description. Every Trades input column has data
  validation with an input message (rows 6 to 5001), following the engine's
  trade checks, and repeated Trade IDs are shown in red.
  `TEST_InputValidation` has 57 cases and 65 checks.
  Upgrading a workbook built from an earlier template: add the two names, or
  rebuild from the template.
- README, Params and NettingSets carry the owner's final layout; NettingSets
  has its title band in B1:B2 and the description in C3. Column headers are
  now compared ignoring case, spaces and line breaks, so the NettingSets
  headers `Client clearing role (N / CM / CLIENT)` and
  `Margin agreements (blank / MIXED / SHARED)` match.
- Params is located through workbook names instead of fixed rows and
  columns: `Params_Code` and `Params_Value` name the "Code" and "Value"
  header cells of the parameter rows, and the top-left cells of `SF_Table`
  and `FX_Table` the headers of the two lookup tables. The engine and the
  tests read Params only through these names, so its rows and columns can
  move; a missing or broken name stops the run (`CORE_Util.ParamsCodeCol`,
  `ParamsValueCol`, `ParamsHeaderRow`, `NamedCell`). Each of the 20
  parameters is read only through its workbook name, which is now required:
  a missing or broken one stops the run, where the engine used to fall back
  to the Code column. `TEST_InputValidation` has 55 cases and 63 checks.
- README names the engine v1.0.0 instead of the prototype's v1.1.0, and its
  note on the cap basis points to Results columns Z-AC. Every font of the
  template names Aptos on the style instead of following the workbook theme,
  so a sheet copied into another workbook keeps Aptos.

- Sheet layout: every sheet of the template has two margin columns (A:B,
  width 4) and two margin rows (row 1, 24 pt, 14 pt font; row 2, 26 pt, 16 pt
  font), and every font, including the theme fonts and the sheet buttons, is
  Aptos. Tables now start in column C, with headers on row 6 and data from
  row 7; the Results run summary is in C4 and the status formula in C5,
  which watches NettingSets `$C:$Q`, Trades `$C:$Y` and Params `$C:$J`.
  Formulas, workbook names, data validation, conditional formatting, merged
  cells, filters, frozen panes and the buttons moved with the cells. The VBA
  addresses tables through `CORE_Util.TableCell` and `TableRange` with
  `CORE_Config.COL_OFFSET` = 2; `HEADER_ROW` is 6 and `FIRST_DATA_ROW` 7.
  A workbook built from the earlier template must be rebuilt from this one:
  the synchronizer changes code only, and the new code stops at the header
  check of an old layout.

- `tools/check_source.py` checks that `Sync-SACCR-VBA.ps1` reaches every VBA
  component: each must lie in a directory of its `$sourceRoots`, each listed
  directory must exist, and the document modules in `src/workbook/` must match
  the template's sheets by CodeName.

- `tools/Sync-SACCR-VBA.ps1` also imports the test modules from
  `tests/modules/`, so one synchronization updates every engine and test
  module of the development workbook.

- Netting-set terms by regime (#39). The cap of a margined netting set now uses
  C = NICA in both regimes; the BCBS cap used VM + NICA, which CRE52.2 and
  CRE52.10 do not support (T16 becomes 40.77m, cap applied, as under CRR). VM
  held on an unmargined netting set is rejected; VM posted on one is a one-way
  agreement, kept in C under BCBS and rejected under CRR, where the set must be
  margined. NettingSets column D is now the client clearing role, N, CM or
  CLIENT: the 5-day MPOR floor applies to CM in both regimes and to CLIENT
  under CRR only. A new column O, Margin agreements, declares MIXED or SHARED
  structures, which are rejected. The MPOR after disputes, 2 x (F + N - 1),
  and of a large or illiquid set, 20 + N - 1, are kept and recorded as
  decisions. Fixtures use `clearing_role` instead of `cleared`; the Results
  status formula watches NettingSets A:O. `TEST_InputValidation` has 53 cases
  and 60 checks; `TEST_Cases` has 37 checks.
  Upgrading a workbook built from the earlier template: the synchronizer
  changes code only, and a run stops at the header check until NettingSets D4
  reads `Client clearing role (N/CM/CLIENT)`, O4 reads
  `Margin agreements (blank/MIXED/SHARED)` and Results A3 watches
  `NettingSets!$A:$O`, as in the new template; or rebuild from the template.

- CRR scope checked against the CRR text (#33). With `IRBucketOffset` set to
  FALSE, a CRR netting set with interest-rate trades is now `INVALID` and its
  EAD withheld, because Article 280a(3) has only the formula with offsets
  across maturity buckets; Basel netting sets keep both formulas. Inflation
  trades can be entered as interest rate with a risk factor such as
  `EUR-INFL`, each in its own hedging set (Articles 277(4)(a), 277a(1));
  before, any IR risk factor other than a currency was rejected. The template
  notes confirm the 18% climatic-conditions factor (Article 280e(5)), explain
  the unrated credit mapping (Article 280c(5)(a)) and state that the optional
  sold-option and sold-protection reductions are not applied.
  `TEST_InputValidation` has 39 cases and `TEST_Aggregation` 11.
- The CRE99 test cases (`cre99-example-*`), T01 to T14, are checked against
  the Basel Framework chapter CRE99 (effective 1 January 2023); T02 now
  expects the printed -0.2694 instead of -0.27. They are in USD, unscaled: CRE99 prints USD thousands, so amounts and expected values are
  the printed figures times 1,000 and the tolerance is half a printed unit,
  500 USD (#44). T10 used 0.6 thousand and T07 2 thousand; both now use the
  half-unit rule. Example 3 prints a 9-month maturity, which no number of days
  gives at 365 days a year, so a fixture can now set `DaysPerYear` in a new
  `parameters` field and Example 3 runs with 360 and 270 days.
  `TEST_CaseRunner.SetParameter` applies it for one case, and each case now
  sets the FX-table rate of its calculation currency to 1, because the
  template's rates are per EUR. The source register records CRE99 and the
  BCBS 279 version.
- VBA modules carry an upper-case role prefix: core modules `CORE_Config`,
  `CORE_Engine` and `CORE_Util`; the public worksheet functions in
  `SACCR_Formulas`; and the test modules `TEST_Harness`, `TEST_MainState`,
  `TEST_CaseRunner`, `TEST_Cases` and `TEST_InputValidation` (previously
  `M_Config`, `M_Engine`, `M_Util`, `M_Formulas`, `TestHarness`,
  `TestMainState`, `CaseRunner`, `TestCases` and `TestInputValidation`).
  `M_Main` keeps its name, so the sheet buttons are unchanged, and the
  `SACCR_*` worksheet functions keep theirs. `tools/check_source.py` enforces
  the prefixes. A workbook built before the rename must be rebuilt from the
  template: importing the renamed modules next to the old ones gives
  duplicate declarations.
- The project is licensed under the Mozilla Public License 2.0 instead of the
  MIT License. MPL 2.0 is file-level copyleft: modified versions of the
  project's files must stay under MPL 2.0 and their source must be available
  when distributed; the files can still be combined with code under other
  licenses.

### Fixed

- EBA Q&A 2023_6962 is cited by its source ID `EBA-QA-2023_6962` in the
  expected files, the VBA comments and the methodology (#129); the source of
  the posted-VM records stays `CRR`, whose Art. 274(3) is the rule, and the
  values and classes are unchanged. The engine header no longer says the Basel
  cap uses VM + NICA. The T17 note names EBA Q&A 2022_6354 as the source of the
  scenario: a question the EBA rejected on 26/07/2022 without an answer.
  `check_test_cases.py` now reads source IDs containing an underscore and fails
  when a registered EBA Q&A is cited in prose.

- The synchronizer coverage check of `tools/check_source.py` no longer passes
  when it cannot do its job: a commented-out `$sourceRoots` entry is not
  counted, an entry in any other form than `(Join-Path $repoFullPath "...")`
  fails (#136), and a missing `tools/Sync-SACCR-VBA.ps1` or
  `src/workbook/SACCR_Template.xlsx` fails instead of skipping the check
  (#137).

- Methodology traceability: the MPOR row names the 5-business-day floor for
  client clearing next to the 10-day bilateral floor; the dispute row says the
  engine's doubling of F + N - 1 follows CRE52.50–.51 and exceeds the CRR text
  (Art. 285(4)–(5)) only when remargining is not daily; Article 278(2) is listed
  as not implemented; EBA Q&A 2023_6962 has its own source ID and SHA-256.

- Installation instructions now show 35 input-validation cases and checks,
  matching the merged suite and recorded Excel result (#91), and describe
  the formula-based status from #89 as integrated rather than pending.

- Runtime review safeguards: reject zero/negative global alpha (#71), recognize
  parameter rows below the FX table (#80), and attempt each output cleanup even
  when another sheet is protected (#85). Cleanup reports every failed sheet
  with its error number, source and description, appended to the original run
  error. A secondary Checks-write failure is also reported. MainState now has
  18 cases, including multiple protected outputs, protected Results + Checks,
  preservation of the primary error and successful recovery. An empty run
  fingerprint means no valid results, even if protected cells remain visible.
  New input regressions are included; Excel verification is still required
  before merge.
- VBA synchronization refuses known retired repository components before
  changing the project (#71, #84, #87). Rebuild legacy workbooks from the current
  template; unrelated user components are left alone. Windows/Excel validation
  remains required.

- The IR add-on diagram includes both ten-business-day floors, separate
  inflation hedging sets, and the distinct CRR/Basel one-year bucket boundary
  (#77). The SD floor also applies under corrected CRR Article 279b.

- The template README now lists CORE_Engine, SACCR_Formulas, CORE_Config and
  CORE_Util, matching the renamed modules (#84). Only the C70 text changes;
  formulas, styles, names, sheet CodeNames and all other package parts are preserved.

- Fixture validation enforces quantity-specific units and rejects non-finite
  JSON numbers, including exponent overflow (#65; carries forward the unmerged
  fix from #76). The workbook gate now inspects every relationship part for
  dangling VBA references (#78). Sync documentation records the owner-selected
  temporary-backup lifecycle (#87).

- Malformed inputs are rejected instead of silently changing the portfolio
  (#35). A trade row with data but no Trade ID, including rows after the last
  ID, and a repeated Trade ID are rejected, making the netting set
  `INCOMPLETE`. A missing or non-numeric MtM is an error rather than an
  assumed 0, and an unreadable date or lambda is an error rather than blank. A
  netting-set flag that is not Y/N, an amount that is not a number or an
  unknown regime override makes the netting set `INVALID` with its EAD
  withheld; a NettingSets row with data but no ID is reported. On Params, a
  value that is present but not a number, an unreadable `IRBucketOffset` and a
  non-numeric supervisory factor, correlation or volatility stop the run.
  Blank optional fields still take their documented defaults.
  `tests/modules/TEST_InputValidation.bas` covers these cases in Excel.

- The run stops when the workbook layout or the parameter tables are
  ambiguous (#35). Before reading any input, the engine checks that the
  Params, NettingSets and Trades sheets exist and that every column it reads
  by number has its expected header (`CORE_Config` `_HEADERS` constants), so an
  inserted, deleted or moved column cannot shift values into the wrong
  fields. A parameter workbook name that no longer refers to a cell (`#REF!`)
  and a parameter listed twice on Params with different values stop the run.
  A supervisory-factor key or currency listed twice with different values
  also stops it, instead of the first row being used silently; a duplicate
  with the same values is a warning. `TEST_InputValidation` adds ten cases for
  these rules.

- A supervisory factor or FX rate below a blank row of its table on Params
  is reported and stops the run (#35). Each table ends at its first blank
  key, so such rows, and a value whose key is blank, were silently not read.
  `TEST_InputValidation` adds four cases.

- A netting-set ID listed twice makes that netting set `INVALID` with its
  EAD withheld (#35). The second row was ignored with an error, but the set
  was still reported `VALID` with the first row's collateral terms.
  `TEST_InputValidation` adds this case and three with an Excel error value
  (`#N/A`) in a trade's MtM, a netting-set amount and a parameter, each of
  which is rejected rather than read as blank.

- A netting set with a rejected trade no longer reports an EAD computed on
  its remaining trades (#36). Results has a Status column and a row for
  every input netting set: `VALID`, `INCOMPLETE: r of n trade(s) rejected`
  with the exposure columns blank, an error on Checks and no contribution to
  the TOTAL, or `NO TRADES`. Every run, and every validation, first clears
  the output sheets, so a failed or validation-only run cannot leave earlier
  results looking current. The run message, the run summary and
  `RunSACCR_Silent` (`incomplete=`) report the count of withheld netting
  sets. In the template, the deliberate error trades `OT-X` and `A11` move
  to their own netting set `DEMO-REJECTED`, so the CRE99 examples and
  `NS-ALPHA` still report. The template no longer carries the prototype's
  saved output rows, which no longer matched these inputs; Results A2 says
  that no results exist until the workbook is run.

- A Checks sheet that cannot be written, for example because it is
  protected, now fails the run (#36). The errors were swallowed: the new
  results were shown while Checks still held the previous run's messages.
  The run now clears the other output sheets, says why in Results A2 and
  raises `ERR_CHECKS_WRITE`; the button shows the reason and
  `RunSACCR_Silent` raises it. `TEST_MainState` adds the case.

- A run stopped by an unexpected error no longer leaves output tables from
  two runs (#36). An error after the engine had started writing, for
  example on Results or HedgingSets, left the sheets already written with
  this run's values and the others empty or from the previous run. The run
  now clears every output sheet, logs the error on Checks and in Results
  A2, and raises it unchanged. `TEST_MainState` adds the case, through a
  test seam that fails the run after TradeCalc, Results and Buckets.

- Results are recognised as out of date once an input changes (#36). A
  completed run stores a fingerprint of everything it read (NettingSets,
  Trades without the Comment column, and Params) in a hidden workbook name,
  shows it in Results A2 and in the `RunSACCR_Silent` result line
  (`inputs=`), and `M_Main.ResultsStatus()` compares it with the inputs as
  they are now: `CURRENT`, `STALE` or `NONE`. Results A3 holds the formula
  `=ResultsStatusText(SACCR_RunInputs, NettingSets!$A:$N, Trades!$A:$W,
  Params!$A:$H)`, which Excel recalculates when an input or the fingerprint
  changes, and is red when it reads `OUT OF DATE`. A formula is used rather
  than a macro because a macro writing to the workbook clears Excel's undo
  history; the first version marked A2 when the Results sheet was activated
  and so lost the user's Undo. A failed or validation-only run and Clear
  outputs empty the fingerprint. `TEST_MainState` adds the case.

- Aggregation no longer depends on the order of the trade rows (#37). A
  credit or equity entity, or a commodity, given two sub-classes in one
  netting set took the factor and correlation of whichever trade came
  first, with a warning; the netting set is now `INVALID` with its EAD
  withheld, whatever the order. A basis or volatility trade without a
  hedging-set label, which shared one blank hedging set, is rejected, as is
  a reference or label containing `|` or `#`, which the grouping keys use,
  and an interest-rate risk factor that is not a 3-letter currency code.
  References and labels are compared after trimming and in upper case, with
  no aliases. `TEST_InputValidation` adds five cases and the new
  `TEST_Aggregation` checks that the demo portfolio gives the same Results
  in reversed and rotated row order, and that offsets happen exactly within
  a bucket and never across currencies, netting sets, credit entities or
  hedging sets.

- The workbook macros restore calculation mode, events and screen updating to
  the values they found, including settings that were off, instead of
  switching events and screen updating on (#43). Each setting is restored
  independently; the original error is kept and a cleanup failure is reported
  separately (`ERR_CLEANUP_FAILED`). A second operation started while one is
  running is refused. `RunSACCR_Silent` restores the previous silent flag,
  returns a machine-readable result line and raises failures instead of
  printing them. `tests/modules/TEST_MainState.bas` covers these cases in Excel.

- Allow scheduled traffic exports without relying on a webhook payload, while
  retaining the default-branch restriction for manual dispatches (review #53).
  The analytics environment also restricts access to main independently.
- VBA jump checks resolve labels separately in each reachable compilation
  environment; mutually exclusive labels cannot hide missing targets or create
  false duplicates (PRs #21 and #24).
- Public API checks inspect colon-separated statements without splitting strings
  or named arguments; the conditional checker requires `PtrSafe` in the actual
  declaration modifier position (PRs #21 and #24).
- The cleanup example reports restoration failures separately from the primary
  error and attempts both restorations. Metadata triage includes closed issues
  and documents the maintainer blank-issue bypass (PRs #18 and #19).
- Validation hardening: a failed Excel run must cite a matching failure report, a
  timed-out run claims no results, the harness must report a Windows host, and
  changelog releases must be listed newest first (PRs #28, #30, #47 and #49).

### Known limitations

- The imported engine is not yet validated under the repository's test-case
  policy and keeps the deviations listed in `docs/REPOSITORY_STRUCTURE.md`. No
  automated check compiles VBA or runs Excel.
- Excel evidence is manual and covers one 64-bit host; 32-bit is untested.
- The methodology sources are registered but not yet verified against their
  official texts.

---

[Unreleased]: https://github.com/danielep71/VBA-SACCR-Toolkit/commits/main
