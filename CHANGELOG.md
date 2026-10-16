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

Nothing has been released yet; the first release is planned as `1.0.0` from
`release/1.0.0`. Pre-release status does not excuse undocumented breaking
changes.

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

First release, planned as `1.0.0` from the release branch `release/1.0.0`.
Everything below is new in it; engineering detail is in the linked issues and
pull requests.

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
- Numerical test cases as JSON fixtures, from which `TEST_Cases` is generated
  (#59, #44): the CRE99 examples are published references, the others
  illustrative. A failure-injection mode proves the suites catch a wrong value.
- Excel evidence bound to an exact commit (#46): `docs/EXCEL_EVIDENCE.md` gives
  the manual procedure, `.github/excel-evidence-policy.json` the expected
  suites and counts, and `tools/check_excel_evidence.py` validates a record.
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

[Unreleased]: https://github.com/danielep71/VBA-SACCR-Toolkit/commits/release/1.0.0
