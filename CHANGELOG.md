<div align="center">

# 📜 Changelog

*Release history for K-SACCR counterparty credit risk in Excel/VBA*

[![Format](https://img.shields.io/badge/Format-Keep_a_Changelog-0969da?style=flat-square)](https://keepachangelog.com/en/1.1.0/)
[![Versioning](https://img.shields.io/badge/Versioning-SemVer-6f42c1?style=flat-square)](https://semver.org/spec/v2.0.0.html)
[![Dates](https://img.shields.io/badge/Dates-YYYY--MM--DD-217346?style=flat-square)](#date-and-version-rules)
[![Staging](https://img.shields.io/badge/Staging-Unreleased_first-d97706?style=flat-square)](#unreleased)
[![Contributing](https://img.shields.io/badge/Changes-Contribution_guide-2ea44f?style=flat-square)](CONTRIBUTING.md)

<br>

**User-visible history · Explicit compatibility · Reproducible evidence · Immutable releases**

</div>

---

All notable changes to **K-SACCR** are documented here.

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
- Until a version is released, edit the existing entry of a feature instead of
  appending a new entry that describes a change relative to an unreleased
  state; entries describe the state a user of the release gets.
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

The first release is `1.0.0`, tagged `v1.0.0`. Later releases record every
breaking change under **Compatibility**.

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

## [1.0.1] - 2026-10-18

### Added

- `tools/rsaccr_crosscheck.py` and `tools/rsaccr_run.R` compare the expected
  values of the test cases with the CRAN package `SACCR` 3.4, an independent
  CRR2 implementation (#352). With R 4.3.3, `SACCR` 3.4 and `Trading` 3.2, the
  24 numeric values it can compute, from 12 of the 15 fixtures, agree within
  each expected file's tolerance. Two fixtures are not run because they need
  a commodity option or the `IRBucketOffset` parameter, which the package
  lacks; the third has a matured trade, for which the engine withholds the
  netting set's EAD. It is not a gate:
  it needs R, and agreement between two implementations does not show that
  either is right.

### Changed

- The comments of every core, workbook and test module were checked against
  the code and corrected where stale, wrong or unclear; local variables and
  parameters follow the house-style naming (descriptive PascalCase, single
  lower-case letters for loop counters only). No behaviour changes: each
  module's change reproduces its previous code under the rename map, and the
  eight suites pass unchanged. The evidence-policy gate reads the case-ID
  parameter of `TEST_InputValidation.StartCase` and the `Test_RunAll` suite
  list regardless of the names' letter case.
- Results C5 says what it checks when the inputs match the last run:
  "Inputs unchanged since the last run (inputs …); outputs are checked when
  evidence is exported", instead of "Results current". `ResultsStatus()`
  returns a new state, `EDITED`, when the inputs match but an output cell on
  Results, Checks or a reconciliation sheet was changed after the run; the
  evidence export refuses it as before (#315). C5 does not watch the outputs:
  their fingerprint would be recomputed after every edit. `TEST_MainState`
  gains the case `results.edited-output`: 28 to 29 cases, 491 to 501 checks.

- The benchmark `TEST_Benchmark` covers all six CRR asset classes (#309):
  its mix of 21 legs adds an other-risks (OT) forward on three indices, long
  and short, to the interest-rate, FX, credit, equity and commodity legs.
  On the owner's Windows Excel host, 1,000 and 10,000 legs and the
  inflated used range passed with every leg used and every netting set
  `VALID` (23 checks, restoration PASS). `INSTALLATION.md` gives those
  timings, states that the published larger sizes were measured on the
  five-class mix, and that the timings measure the engine with all its
  outputs, not the reconciliation or the evidence export.

- The source round trip binds an export to its workbook (#313).
  `Export-SACCR-VBA.ps1` writes `export-manifest.json` with the workbook's
  SHA-256, taken before opening and checked after closing, and the SHA-256 of
  every exported file. `compare_vba_export.py --workbook` checks both against
  it and reports `binding=VERIFIED`, `FAIL` (which fails the result) or
  `UNVERIFIED` for an export without a manifest. Before, `--workbook` only
  printed a digest, so an export from one workbook could be presented with
  another. A release workbook now needs `binding=VERIFIED`.

### Fixed

- Exporting evidence over an existing report no longer deletes it before the
  new one is written (#312). The report is written whole to a new file beside
  it, then replaces the older file, which is put back with its attributes if
  the replacement fails; the first error is reported, not one from the
  cleanup. A read-only file is still left as it is. `TEST_Reconcile` injects a
  failure before the new file is opened, while it is written and while the
  older file is replaced, and checks that the older file is unchanged and
  still hidden: 19 to 20 cases, 144 to 162 checks.
- A failed reconciliation now keeps its cause in the run message: the
  message read the error after the cleanup of the reconciliation sheet had
  cleared it, so it said only "failed:". `TEST_Reconcile` checks it with
  an injected failure (#318); its checks go from 140 to 144.
- The PowerShell commands in `docs/EXCEL_EVIDENCE.md`, `tools/README.md` and
  the scripts' own help run through `powershell.exe -ExecutionPolicy Bypass
  -File`, as `Sync-SACCR-VBA.cmd` does: on a machine whose execution policy
  disables scripts, the 1.0.0 export command was refused.
- The Excel evidence policy names `TEST_InputValidation`'s first case as its
  log prints it, `input-trade-middle-row-without-id`, and the policy gate
  reads it from `StartCase` in the same way. With the 1.0.0 policy the
  validator rejected every real log of that suite; the 1.0.0 certification
  records this as its deviation.

### Documentation

- Three explainers in `docs/assets/`, drawn for this repository:
  `SACCR_aggregation.svg`, how a netting set is aggregated from trade to EAD,
  in the README's "How a run works" and on the website;
  `SACCR_EAD_tree.svg`, the exposure value broken down into its inputs, in
  the methodology; and `SACCR_RC_PFE.svg`, the replacement cost of a margined
  netting set and the PFE multiplier, also in the 1.0.1 user manual. Each has
  a PNG for the Wiki.
- The 1.0.1 user manual, in `docs/user-manual/v1.0.1/`, is written for
  someone who receives only `K-SACCR_1.0.1.xlsm` and `SHA256SUMS.txt`.
  Installation is now checking the workbook's SHA-256, unblocking and
  opening it, keeping the delivered file as a reference copy and moving
  inputs to a new version. Building from source, the development workbook,
  the regression-suite inventory, the release certification steps and
  every reference to the source repository and its hosting service are
  gone; troubleshooting sends a damaged workbook back to the reference copy
  instead of a rebuild. It also covers this release's changes (the Results
  C5 text, the `EDITED` state, the export that keeps the older file, the
  1904 date-system recalculation) and the replacement-cost and PFE
  explainer. The 1.0.0 manual is unchanged.
- The roadmap, the support matrix and the diagram of where the CRR uses
  SA-CCR plan CCP exposures and default funds (#112) in v1.0.2, Clearing coverage and CCP completeness,
  instead of v1.7.0, as the milestones now do.
- `docs/VBA_HOUSE_STYLE.md` names local variables and parameters in
  PascalCase, spelled out (`NotionalAmount`), and keeps a single lower-case
  letter for loop counters only. A renaming to these rules is the one name
  change a formatting-only change may make, and comments are kept true: a
  change updates every header, `INPUTS` name and line comment it makes stale.
- `RELEASING.md` makes two checks deferred from the 1.0.0 release standing
  steps of every release (#308): the source round trip of the certified
  development workbook, beside the production one, and the published
  workbook used as a user would, from download and hash check to a refused
  export after an edit. They are first carried out on 1.0.1.
- A diagram of where the CRR uses the SA-CCR exposure value (counterparty
  credit risk, CCPs, large exposures, leverage ratio, CVA risk, output floor),
  with the K-SACCR milestone or scope of each, in the README roadmap and
  `docs/assets/SACCR_uses_in_CRR.svg`.
- The website shows the same diagram in a section of its own, with the
  leverage-ratio and output-floor adjustments; on narrow screens it keeps a
  readable size and scrolls sideways. The website also points those who
  download the workbook to its SHA-256 check and unblocking in
  `INSTALLATION.md`, and its "Build the workbook" button reads "Build from
  source" and opens the build steps.
- The website opens with a banner under the navigation: a photograph supplied
  by the owner, `docs/assets/SACCR_banner.webp`, behind the text "SA-CCR for
  derivatives, Basel CRE52 · EU CRR", set as HTML.
- Troubleshooting in `INSTALLATION.md` explains that Results C5 can keep its
  earlier status after only the workbook's date system is switched, until a
  full recalculation (Ctrl+Alt+F9), and that a run under the 1904 date system
  stops (#317). The behaviour is unchanged: Excel raises no event for the
  change, and a volatile C5 would recompute the input fingerprint on every
  edit.

## [1.0.0] - 2026-10-17

First release. Everything below is new in it; engineering detail is in the
linked issues and pull requests.

### Changed

- Refresh the project website with a deep-blue and azure palette, pale blue
  surfaces, and coordinated navigation, links, and example-card accents (#284).

### Added

- Add the owner-supplied Google Search Console verification tag to the project
  website, enabling ownership verification after deployment (#287).

- Public GitHub Pages project website with a workflow overview, workbook
  screenshot, published Basel example, explicit validation limits, and links
  to source and documentation (#279).

#### Calculation

- SA-CCR exposure at default under the EU CRR or Basel CRE52, chosen per
  netting set with the Params `Regime` as default: replacement cost, the PFE
  multiplier, the add-ons of interest rate, FX, credit, equity, commodity and,
  under CRR only, other risks, and EAD = alpha x (RC + PFE). A margined
  netting set's EAD is capped at its unmargined EAD, with C = NICA in the cap
  in both regimes (CRE52.2, CRE52.10; #39).
- Interest rate (#33, #191): maturity buckets by end date, bucket 1 ending at
  exactly 1 year under CRR (Art. 280a(3), Table 2) and below 1 year under Basel
  (CRE52.57(3)); `IRBucketOffset` TRUE takes the formula with offsets across
  buckets, FALSE the sum of absolute bucket values (CRE52.57(5)), which is
  Basel only: a CRR netting set with interest-rate trades is then `INVALID`.
  Inflation is entered as interest rate with a risk factor such as `EUR-INFL`,
  in its own hedging set (CRR Art. 277(4)(a), 277a(1)); any other interest-rate
  risk factor is a three-letter currency code.
- Credit and commodity factors: CRR single names by credit quality step
  `CQS1` to `CQS6` (Art. 280c(5)), Basel by rating `AAA` to `CCC` (CRE52.72);
  an unrated name is entered as `CQS3`, or `CQS5` where Art. 128 applies
  (Art. 280c(5)(a)). Climatic conditions are a CRR commodity hedging set with
  an 18% factor (Art. 277a(1)(e)(v), 280e(5)).
- Options: CRR lambda shift for interest-rate and commodity options (`CRR-RTS`
  Art. 5). An interest-rate or credit option needs the start date of its
  underlying, for a Bermudan the earliest exercise (CRE52.31(2), CRE52.42(2));
  an expiry after maturity is rejected (#40). An FX option is entered in
  market quote order: EUR, GBP, AUD, NZD, USD, CAD, CHF, JPY, then other
  currencies alphabetically (CRE52.59 FAQ1).
- Economic trades entered as legs (#38): rows sharing an Economic trade ID
  form one transaction, with the MtM on exactly one leg flagged Y. Legs in
  different asset classes map one transaction to several risk categories (CRR
  Art. 277(3); CRE52.28); legs in one class must all be options and form an
  option combination (CRR Art. 274(6); CRE52.42(4), CRE52.43). A malformed
  economic trade rejects every leg.
- Every run is `REGULATORY` or `WHAT-IF` (#41): the Params values, the
  supervisory-factor table and alpha overrides are compared with the built-in
  parameter set `CRE52-CRR-2026-10-09`, and each difference is listed on Checks
  as a "What-if:" warning. `DaysPerYear` and `IRBucketOffset` are conventions
  and do not change the class.
- Results do not depend on the order of the trade rows (#37); references and
  labels are compared trimmed and in upper case, with no aliases.

#### Netting sets and margining

- MPOR floors (#39): 10 business days bilateral; 5 for client clearing, for the
  clearing member (CM) in both regimes and for the client (CLIENT) under CRR
  only (CRE54.8, CRE54.12, CRE54.14; CRR Art. 279c(1)). After margin disputes
  the MPOR is 2 x (F + N - 1) in both regimes, as CRE52.51(3) gives; under CRR
  this is higher than Art. 285(4)-(5), 2F + N - 1, when remargining is not
  daily (a recorded decision). A large or illiquid set, flagged by the user,
  takes 20 + N - 1 (CRE52.51; CRR Art. 285). A margined netting set of more than
  5,000 trades with the flag N gives a warning, since the floor follows from
  the previous quarter's count.
- Variation margin (#39): VM received on an unmargined netting set is
  rejected; VM posted on one is a one-way agreement, kept in C under Basel and
  rejected under CRR, where the set must be margined (CRR Art. 272(7a)).
- Several margin agreements in one netting set, or one agreement over several
  netting sets, are declared as `MIXED` or `SHARED` and rejected (CRR Art.
  274(4), 275(3); CRE52.74–.75).

#### Inputs and validation

- Malformed inputs are rejected instead of silently changing the portfolio
  (#35, #147, #194): a trade row without a Trade ID, a repeated Trade ID or
  netting-set ID, a missing MtM, an error value such as `#N/A`, and values
  outside the stated ranges, also when pasted. Blank optional fields take
  their documented defaults.
- The run stops when the layout or parameters are ambiguous (#35): a missing
  sheet, workbook name or expected column header (compared ignoring case,
  spaces and line breaks), a parameter or supervisory-factor key listed twice
  with different values, rows below a blank row of a Params table, an FX rate
  of the reporting currency other than 1 (#41), or a reporting currency that
  is not three letters A to Z.
- Dates (#40): a date, a 1900-system serial or text in ISO `yyyy-mm-dd` form;
  other text rejects the trade, or as the reporting date stops the run, and a
  workbook in the 1904 date system stops the run. An FX trade needs FX rates
  for both currencies of its pair, and its notional currency is one of them.
- Inputs are read with `Range.Value2`, so results do not depend on the number
  format (#196); a run recalculates the workbook first, and a recalculation
  that does not finish stops it with `ERR_RECALC_PENDING` (#195, #204).
- Data validation with input messages on Params, the factor and FX tables,
  NettingSets and Trades, and repeated keys shown in red.

#### Results and run status

- One Results row per input netting set with a status: `VALID`, `INVALID`,
  `INCOMPLETE` (a rejected trade; the EAD is withheld rather than computed on
  the remaining trades) or `NO TRADES` (#36). Results ends with one total row
  per regime present in the inputs, `TOTAL CRR` and/or `TOTAL BCBS`, each
  summing that regime's `VALID` netting sets and saying how many are covered; there is no total across regimes, and `TOTAL`,
  `TOTAL CRR` and `TOTAL BCBS` are reserved IDs (#198).
- Results are recognised as out of date (#36, #152, #197): a completed run
  stores a fingerprint of its inputs in the hidden name `SACCR_RunInputs`, and
  a status formula on Results reads `OUT OF DATE` once an input, a header or
  the date system changes; formatting does not count.
- Every run and validation first clears the output sheets; the macros restore
  Excel's settings as they found them and refuse a second operation (#36,
  #43). Errors: `ERR_RUN_ACTIVE`, `ERR_CLEANUP_FAILED`, `ERR_CHECKS_WRITE`,
  `ERR_PARAMS_LAYOUT`, `ERR_FORMULA_DOMAIN`, `ERR_RECALC_PENDING`,
  `ERR_EXPORT_WRITE`, `ERR_OUTPUT_LAYOUT`.
- `RunSACCR_Silent` returns a result line: `RESULT`, `operation`, `errors`,
  `warnings`, `trades_used`, `trades_read`, `economic_trades_used`,
  `economic_trades_read`, `incomplete`, `total_ead_crr`, `total_ead_bcbs`,
  `inputs`, `run_class`, `reconciliation`, `trade_reconciliation`, `cleanup`.

#### Reconciliation and evidence export

- Reconciliation with a bank's own figures (#45): per netting set,
  `BankNettingSets` (EAD, RC, PFE, add-ons, multiplier) against Results on
  `Reconciliation`; per trade leg, `BankTrades` against TradeCalc on
  `TradeReconciliation`, with the leg's Economic trade ID. Results: `MATCH`,
  `DIFFERENCE`, `MISSING IN CALCULATOR`, `NOT IN BANK FILE`, `DUPLICATE`,
  `REGIME MISMATCH`, `WITHHELD`, `INVALID INPUT`, and per leg `EXCLUDED` and
  `NOT APPLICABLE`. Amounts are scaled (`BankNS_Scale`, `BankTR_Scale`) and
  compared within an absolute or relative tolerance (`BankNS_AbsTol`,
  `BankNS_RelTol`, `BankTR_AbsTol`, `BankTR_RelTol`); figures of different
  regimes are never compared. Status formulas use the fingerprints
  `SACCR_ReconInputs` and `SACCR_TradeReconInputs`.
- **Export evidence** on `Reconciliation` (#45) writes one CSV file, `;`
  separated, `.` decimal point, UTF-8 without byte-order mark, CR LF: run
  summary and fingerprints, parameters, the factor and FX tables, results,
  both reconciliations, Checks and the limitations. It refuses while a
  recalculation is pending, the results or a reconciliation are not current,
  or an output cell was edited after the run (`SACCR_RunOutputs`).
  `ExportEvidence_Silent` returns `RESULT=OK` or `RESULT=REFUSED; reason=...`.

#### Worksheet functions

- Ten `SACCR_*` worksheet functions in `SACCR_Formulas`, listed in
  `docs/PUBLIC_API.txt`. All but `SACCR_NormCDF` return `Variant` and give
  `#NUM!` for a value outside its domain, `#VALUE!` for text or a Boolean where
  a number is needed, and `#REF!` for ranges of different sizes or shapes; an
  error value in an input is passed on (#42).

#### Workbook and sheets

- Template `src/workbook/SACCR_Template.xlsx` with 16 sheets: README, Params,
  NettingSets, Trades, Results, HedgingSets, Buckets, TradeCalc, Checks,
  BankNettingSets, Reconciliation, BankTrades, TradeReconciliation,
  TestCatalogue, FormulaTest and Basel_vs_CRR, and no VBA project: the
  workbook is built by saving it as `.xlsm` and importing the source (#58).
- Every table is located through workbook names (`Params_Code`,
  `NettingSets_Header`, `Results_Header` and the other `*_Header` names), so
  rows and columns can move (#141, #253). The demonstration portfolio holds
  the CRE99 examples; FormulaTest checks each `SACCR_*` function against
  Excel's own formulas.
- The project is named K-SACCR; the repository keeps the name
  VBA-SACCR-Toolkit and identifiers keep the `SACCR` prefix. Licensed under
  the Mozilla Public License 2.0.

#### Testing and Excel evidence

- Eight suites run in Excel, together through `TEST_Suite.Test_RunAll`:
  `TEST_Harness` (4 cases, 6 assertions), `TEST_Formulas` (9, 82),
  `TEST_MainState` (28, 491), `TEST_Cases` (20, 39), `TEST_InputValidation`
  (133, 209), `TEST_Aggregation` (11, 11), `TEST_Invariants` (4, 23) and
  `TEST_Reconcile` (19, 140): 228 cases and 1,001 checks on 64-bit Excel
  (`TEST_Formulas` has 80 checks on 32-bit).
- A benchmark for the supported portfolio size, `TEST_Benchmark.RunBenchmark`,
  run on its own in a development workbook (#291): it times the writing of the
  inputs and the engine run on synthetic CRR portfolios of 1,000 and 10,000
  legs, or other sizes up to 1,000,000, and on 1,000 legs with an
  inflated Trades used range, and checks that every leg is used and every
  netting set is `VALID`. It is not a suite: `Test_RunAll` and the Excel
  evidence do not include it. Measured on 2026-10-17: 50,000 legs in 8.5 s
  and 500,000 legs in 114 s; the results are in INSTALLATION under Portfolio
  size.
- Numerical test cases as JSON fixtures, from which `TEST_Cases` is generated
  (#59, #44): the CRE99 examples are published references, the others
  illustrative. A failure-injection mode proves the suites catch a wrong value.
- Excel evidence bound to an exact commit (#46): `docs/EXCEL_EVIDENCE.md` gives
  the manual procedure, `.github/excel-evidence-policy.json` the expected
  suites and counts, and `tools/check_excel_evidence.py` validates a record.
  `TEST_MainState` and `TEST_Reconcile`, run on their own, keep the macros'
  run messages out of their report, which otherwise overran the VBE
  Immediate window's 200 lines and lost its header and first cases.
  A source round trip (`Export-SACCR-VBA.ps1`, `compare_vba_export.py`) checks
  a built workbook's code against a commit (#34).

#### Tooling and repository

- `python tools/check.py` runs the static gates before every push and in CI:
  Git storage and line endings, exported VBA integrity, component locations,
  VBA jump, conditional-compilation and public-API checks, test-case files,
  the generated case module, the evidence policy, the changelog and
  whitespace. None compiles VBA or runs Excel.
- `tools/Sync-SACCR-VBA.ps1` builds a production or development workbook from
  the source (#34), with drag-and-drop `.cmd` launchers.
- CI with Ruff, mypy and actionlint; Dependabot; issue forms and a PR
  template; label sync and drift workflows; a daily traffic export.

#### Documentation

- README with screenshots of the workbook, quick start, regime differences and
  limitations; INSTALLATION with supported hosts and the build procedure;
  CONTRIBUTING, SECURITY, RELEASING and a Contributor Covenant code of conduct.
- Methodology in `docs/methodology/`: source register (CRR, `CRR-RTS`, CRE52,
  CRE54, CRE99, EBA Q&A 2023_6962) with copies of the texts, assumptions and
  decisions, support matrix, traceability table and test-case format.
- `docs/GOVERNANCE.md` (branch policy, labels, review), repository structure,
  VBA house style and `docs/assets/README.md` for the images: a one-page EAD
  reference card (also shown on the workbook's README sheet), workbook
  screenshots and the social preview.
- A user manual, as pandoc Markdown with its build instructions in
  `docs/user-manual/v1.0.0/`; its PDF will be attached to the 1.0.0 release.

### Differences from the prototype workbook

The prototype `SACCR_Calculator.xlsm` (engine v1.1.0) gives different results
in these cases:

- The Basel cap of a margined netting set uses C = NICA, not VM + NICA (#39).
- Under CRR an interest-rate trade ending exactly 1 year after the reporting
  date is in bucket 1, not bucket 2 (#191).
- A credit or equity entity or a commodity given two sub-classes makes the
  netting set `INVALID`; the prototype took the first trade's, so results
  depended on row order (#37).
- A netting set with a rejected trade is `INCOMPLETE` with its EAD withheld;
  the prototype reported an EAD on the remaining trades (#36).
- Results has a total per regime; the prototype's single `TOTAL` added CRR and
  Basel EADs together (#198).
- A reporting-currency FX rate other than 1 stops the run instead of giving a
  warning (#41).
- Inputs the prototype read as defaults are rejected: a missing MtM (was 0),
  an error value (was blank), non-ISO date text (read with the regional
  settings), an option without start date (S = 0). A Currency- or
  Accounting-formatted input is no longer rounded to four decimals (#35, #40,
  #194, #196).

### Notes for pre-release builds

For a workbook built from `release/1.0.0` or `main` before the release:

- Rebuild it from `src/workbook/SACCR_Template.xlsx`: the sheet layout and the
  named output headers (`*_Header`) changed (#141, #253).
- Module names: `M_Config`, `M_Engine`, `M_Util` and `M_Formulas` are now
  `CORE_Config`, `CORE_Engine`, `CORE_Util` and `SACCR_Formulas`.
- The result line's `total_ead` is replaced by `total_ead_crr` and
  `total_ead_bcbs`; `SACCR_*` functions return `Variant` with worksheet errors.
- The results fingerprint format changed: Results show `OUT OF DATE` until run
  again. Evidence records of schema 1 are no longer accepted.

### Known limitations

- Only Basel rules covered by a published CRE99 output are validated; no CRR
  rule is validated in this release. No automated check compiles VBA or runs
  Excel.
- Excel evidence is manual and covers one 64-bit Windows host; 32-bit is
  best effort, untested; Excel for Mac and for the web are not supported.
- Digital options are not supported in this release (#120). Option legs of one
  economic trade are calculated as an ordinary option combination: neither the
  CRR payoff-sized collar nor the BCBS cap of the effective notional is
  applied.
- The optional reductions for sold options and sold credit protection are not
  applied, which gives a higher exposure value.

---

[Unreleased]: https://github.com/danielep71/VBA-SACCR-Toolkit/compare/v1.0.1...HEAD
[1.0.1]: https://github.com/danielep71/VBA-SACCR-Toolkit/releases/tag/v1.0.1
[1.0.0]: https://github.com/danielep71/VBA-SACCR-Toolkit/releases/tag/v1.0.0
