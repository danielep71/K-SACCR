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

### Added

- Define version 1 of the Trades and NettingSets CSV input contract (#206):
  `docs/csv/contract.json` and `docs/csv/CSV_CONTRACT.md` specify the columns,
  types, allowed values, blank rules, encoding, `;` as the field separator,
  quoting, header and version policy, limits and finding codes, with blank templates generated from the
  contract. `tools/csv_contract.py` is the reference implementation, tested
  on synthetic valid and invalid fixtures in `tests/csv/`, and the new
  `check_csv_contract` gate keeps the contract consistent with the engine's
  headers, the template's validation lists, the templates and the fixtures.
  No workbook import yet: that is #207 to #211.
- Add an **Export CSV template** button to `Trades` and `NettingSets` (#206).
  It writes the blank template of the CSV input contract, the header line and
  a line feed, to a file the user chooses; the file is the committed template
  in `docs/csv/templates/` byte for byte. `ExportCsvTemplate_Silent` does the
  same without a dialog for automation and tests. The evidence export and the
  template export share one writer, which writes the new file whole before it
  replaces an older one. `check_csv_contract` also compares the header lines
  in `CORE_Config` with the contract. `TEST_Reconcile` gains the case
  `export.csv-templates`: 21 cases, 180 checks.
- Add the workbook's CSV reader (#207): `CORE_CsvReader` reads a Trades or
  NettingSets file under the CSV input contract and reports every finding
  with its code, line, record and column, or returns the records of an
  accepted file as their raw text. It decodes UTF-8 itself and checks numbers
  and dates as text, so the outcome does not depend on the Windows locale,
  and it writes nothing; staging the records is #208. The contract gains a
  file size limit of 128 MiB, checked from the size before reading
  (`CSV-TOO-LARGE`), and states that keys compare in upper case one character
  at a time, as the engine compares IDs. `tools/generate_csv_vba.py` writes
  the contract into `CORE_CsvSchema` and the fixtures and the new inline
  cases of `tests/csv/cases.json` into the generated suite `TEST_CsvReader`,
  which expects exactly the findings of the reference implementation;
  `Test_RunAll` runs nine suites.
- Check a CSV file from the workbook (#207): the macros `CheckTradesCsv` and
  `CheckNettingSetsCsv`, run from the Macros dialog, read a chosen file with
  the CSV reader and show whether it follows the contract, with its record
  count or its first findings, as many as the message box shows whole;
  `ReadCsv_Silent` does the same for automation. Nothing is written to the
  workbook. A file whose header is one field holding a comma or a tab, such
  as Excel's comma-separated save, gets one `CSV-DELIMITER` finding instead
  of a finding per column. `TEST_Benchmark` gains
  `RunCsvBenchmark`, which times the reader on synthetic files of up to the
  record limit. The reader now trims the zeros of a number and of its
  exponent in one pass, and raises `ERR_CSV_READ` also when the path itself
  is malformed (#446). `TEST_CsvReader`: 93 cases, 199 checks.
- Add the staging sheets `Import_NettingSets` and `Import_Trades` after
  Trades (#208), where a CSV file will be loaded for checking and correction
  before it replaces NettingSets or Trades. Each has the source file, load
  time, record count and status at the top, then a header row with a
  `CSV line` column and the contract's columns in contract order, their CSV
  names above, and data columns formatted as text so that every value stays
  as written. Nothing reads them yet: loading comes with the import macros.
  `check_csv_contract` checks their headers and anchors against the
  contract. Rebuild the workbook from the template: a workbook built from an
  older template has no staging sheets, and a sheet-module synchronization
  stops on the missing modules.

### Changed

- Require pull requests into `release/**` to be up to date with their target
  branch before merging, with **Repository integrity** passing on the updated
  candidate. An Excel verification of a VBA or template change survives such
  an update only when the incoming changes leave the VBA and the template
  untouched, as the pull-request checklist now says; release certification
  still runs on the exact candidate commit.
- Set `release/1.1.0` as the active development branch and the target for Dependabot updates; preserve released history through immutable tags and published artifacts.

### Documentation

- Explain in the README roadmap why there is no v1.7.0 milestone: its CCP
  scope was delivered in v1.0.2 and default-fund contributions moved to
  v1.10.0.
- Count the release website as updated only once its change is integrated
  into `main` and the Pages deployment shows the new release, and add the
  README roadmap section of the current milestone to the places updated when
  a new release branch is created (#434).
- Correct the references of illustrative expected values: the MPOR outputs of
  the CCP cases cite `CRR-CCP` and `CRE54`, the CCP articles' own sources; the
  Basel commodity add-on cites CRE52.69, CRE52.70 and CRE52.72 instead of the
  basis and volatility multiplier of CRE52.73; T21 cites `CRR-RTS` with rule
  `CRR-RTS.5`, as T20 does; T19 gets a CRR locator; the unmargined swap cites
  CRE52.10; and the notes of T15, T19 and T21 no longer call their Python
  implementation independent. The CCP derivation script now prints the
  non-qualifying case its expected file cites. No expected value changes.
- Correct inconsistencies found in a repository audit: the expected files
  without a catalogue ID and their origins in `tests/README.md`, the boundary
  fixtures and the clearing fixtures among the regime differences, the CCP
  articles K-SACCR applies, the 10-day CCP MPOR floor in the parameter table,
  the CCP label of the CRR diagram, the core's host independence as a goal
  with known deviations, the number of CSV files `TEST_Reconcile` writes, the
  harness and seven policy suites of the release certification, the website
  update after publication, every place that names the active branch, the CI
  steps beyond `check.py`, the `check_test_cases` rules, the roadmap rows of
  v1.2.0 and v1.3.0, and a broken link in the VBA house style.
- Describe measurements, screenshots, sources and test derivations without
  referring to the author in the third person: descriptive text no longer
  names "the owner" or "the maintainer", and the expected-value notes of the
  illustrative cases say "derived within the project". Process rules that
  name the owner's role are unchanged.
- Record the owner-authorized retirement of historical branches `release/1.0.1`
  and `release/1.0.2`, preserving all commits in `main`, published tags and
  release artifacts, and retaining `traffic-history` for automation.
- Correct the branch governance: retain and safely fast-forward the active
  release branch after integration into `main`, describe the separate deletion
  ruleset with no branch exclusions or bypass actors, and record
  `release/1.0.0` as deleted with its tag preserved.
- Align the v1.1.0 roadmap with the approved CSV/staging and workbook-protection
  scope, the released v1.0.2 baseline and combined closeout under #212;
  distinguish planned work from released functionality.
- The website's download buttons, checksum link, status line and release
  links point to release 1.0.2, now published.

## [1.0.2] - 2026-10-23

### Added

- `TEST_Aggregation` adds a mixed clearing population (#361): one CRR
  portfolio with an ordinary, a clearing-member, a client and a CCP
  netting set, margined with daily remargining and holding the same swap,
  and a CCP netting set without its QCCP status. It checks that each
  `VALID` set takes its own role's MPOR floor (10, 5, 5 and 10 business
  days), so no role is calculated as another; that `TOTAL CRR` is the sum
  of the four `VALID` sets' EADs and reads
  `4 of 5 netting set(s) VALID; 1 withheld`; that the invalid set is
  `INVALID: 1 input error(s)`; and that the CCP line on Checks names one of
  two CCP sets `VALID` with its EAD. Totals: `TEST_Aggregation` 12 cases and
  15 checks; all suites 288 cases and 1,283 checks.
- NettingSets: trade exposures to a central counterparty (#112, part (a)).
  Client clearing role `CCP` is the bank's trade exposure, as clearing
  member, to a CCP, for its own trades and for client trades it must
  reimburse on the CCP's default; client trades it need not reimburse are
  not entered. Two columns after Margin agreements are required for `CCP`
  and must be blank for every other role: *QCCP (Y/N)* in R (under `CRR` a
  CCP authorised or recognised under EMIR, Art. 4(1)(88); under `BCBS` a
  QCCP of CRE50.3) and *CCP retains VM, collateral not protected (Y/N)* in
  S, which needs Margined = Y. A missing, invalid or misplaced value makes
  the netting set `INVALID`, with an error on Checks. A margined `CCP` set
  has an MPOR floor of 10 business days in both regimes [CRE54.8(2); CRR
  Art. 304(3)(b)], plus N - 1, doubled with disputes and raised by an
  override that is not below it; the large or illiquid flag still sets 20
  days, a conservative limitation, although the 5,000-trade floor does not
  apply to trade exposures. With retained VM each trade takes the larger of
  the netting set's MPOR and the lower of one year and its remaining
  maturity, at least 10 business days, in its own margined maturity factor
  [CRE54.8(3); CRR Art. 304(3)(d)]. The exposure value stays in the regime's
  EAD total. Checks gives each `CCP` set an INFO line with its role, floor,
  rule and risk-weight rule, stated and not applied (2% for a QCCP; for a
  non-qualifying CCP the standardised weight, to be determined outside
  K-SACCR), and each regime with `CCP` sets an INFO line with their number,
  how many are `VALID`, and their EAD. A blank or invalid QCCP value leaves
  the status unknown: the INFO line says so and names no risk-weight rule.
  Before this change, from the start of
  this release, role `CCP` rejected the netting set as not supported; that
  interim status `INVALID: facing a CCP, not supported` no longer exists.
  Default-fund contributions, their own funds requirements and the cap of
  Art. 303(2) / CRE54.40 are not calculated (#400). A CCP entered as the
  counterparty of a role-`N` netting set is still calculated as bilateral.
  The template has the columns, their dropdowns and input messages, and
  guidance; the status formulas of Results, Reconciliation and
  TradeReconciliation C5 watch `NettingSets!$C:$S`. Upgrade: a workbook
  built from an earlier template stops at the NettingSets header check until
  it is rebuilt from the new template, or the two columns are added after
  Margin agreements with the new headers, `NettingSets!$C:$Q` is changed to
  `NettingSets!$C:$S` in those three formulas, and `CCP` and `MULTI` are
  added to the validation list of column F; with R and S blank every other
  netting set keeps its result. Tests: the illustrative fixtures
  `illustrative-ccp-trade-exposure` (CRR and BCBS, MPOR 10, the same EAD),
  `illustrative-ccp-retained-vm` (CRR and BCBS, trade maturity factors 0.3,
  1.0592 and 1.5) and `illustrative-ccp-non-qualifying`, with values from
  `tests/derivations/ccp_trade_exposure.py`; 16 `TEST_InputValidation`
  cases replace the two of the interim rejection. The fixture format gains
  `qccp`, `ccp_retains_variation_margin` and the `ccp` role, and the
  expected format the trade quantity `maturity_factor_margined` (margined
  netting sets only). The source register adds `CRE50`, with a registered
  copy. Totals: `TEST_Cases` 33 cases and 186 checks,
  `TEST_InputValidation` 177 cases and 308 checks.
- Tests for client clearing (#359): two illustrative numerical cases, a
  listed equity index future and call option in a margined netting set where
  the bank is the clearing member (`illustrative-clearing-member-listed`,
  daily remargining, MPOR 5 in both regimes) or the client
  (`illustrative-clearing-client-listed`, weekly remargining, MPOR 9 under
  CRR and 14 under BCBS), each under CRR and BCBS, with MPOR, RC, the
  option's delta, the equity add-on, multiplier, EAD, cap flag and status;
  their values come from `tests/derivations/clearing_listed.py`, a separate
  Python implementation, and stay illustrative until reviewed by someone
  other than the code's author. `TEST_InputValidation` adds the clearing
  member's large-or-illiquid floor (20 days in both regimes, the
  conservative limitation), the client's floor doubled by disputes (10 under
  CRR, 20 under BCBS) and an MPOR override below and above the clearing
  member's floor. `tools/check_test_cases.py` accepts `CRE54.<paragraph>`
  rule IDs. The owner decided that there will be no independent reviewer,
  so these values stay illustrative and #359's criterion on independently
  reviewed values is recorded as not met (decision 8 of the clearing
  coverage). Totals: `TEST_Cases` 24 cases and 71 checks,
  `TEST_InputValidation` 143 cases and 235 checks.
- NettingSets: Client clearing role `MULTI` declares a multi-level client
  structure (a client clearing for its own clients), whose conditions apply
  at every level (CRE54.14(3); CRR Art. 305(4)) and are not modelled: the
  netting set is `INVALID: multi-level client structure, not supported`,
  with an error on Checks, instead of taking an ordinary role's floor. The
  template's dropdown, input message and guidance list it; as for `CCP`, an
  existing workbook needs the new template, or `MULTI` added to the
  validation list of NettingSets column F, to enter it (#359).
- Checks: every `CM` and `CLIENT` netting set gets an INFO line naming its
  role, the MPOR floor and the rule applied, the BCBS client's 10 days marked
  as the conservative interpretation, and the 20 days of the large or
  illiquid flag when it raises the floor; an unmargined one says that no
  floor applies. The role and method are thus in the run's evidence without a
  new Results column (#359).
- Basel_vs_CRR (#389): eight rows completed or qualified after the owner's
  review: Basel's lambda is one value per currency for IR options, and the
  engine accepts an entered lambda for other options only as a technical
  shift; listed credit indices and other multi-name positions under CRR
  Art. 280c(5)(b), with the engine's limits; the BCBS client's 10 days as an
  adopted conservative interpretation, the dispute difference, the clearing
  perimeter and the specific client-clearing tests; the CRR conditions on
  basis hedging sets; the conditions of the optional reductions for sold
  options and sold protection, whose omission gives an EAD not lower than
  with them, not necessarily higher; the CRE52.74 FAQ on one RC and add-ons
  by MPOR; and the cap columns on Results. Three rows are added, as grouped
  on the wiki page: equity and commodity adjusted notional, variable notional
  or quantity, and multiple exchanges and leverage. The Source check header says it records a text
  check, not numerical validation; column J is wider and the edited rows
  taller. The methodology's wording on the omitted reductions is aligned.
- TestCatalogue T07 (CRE99 Example 3 on the demo) says why the demo is
  1.38 below the published 5,406: its 9-month leg is 274 days, 0.7507 years,
  because no date gives 0.75 years at 365 days; tolerance 2 covers this
  input difference, not engine rounding, and the case `cre99-example-3`
  reproduces the example exactly with `DaysPerYear` 360. The owner kept the
  274-day input.
- Results: a conditional format highlights in light red every netting-set
  row whose status is `INVALID` or `INCOMPLETE`, so withheld EADs stand out;
  `VALID`, `NO TRADES` and the `TOTAL` rows are not highlighted. It is part of
  the template: a workbook that only receives the synchronized VBA keeps its
  Results sheet without it.
- Trades: settled-to-market trades (#360). Two columns before Comment,
  *Settlement type (N/STM)* in AB and *Next reset date* in AC; Comment moves
  from AB to AD. `STM` declares that the contract settles its outstanding
  exposure and resets its value to zero on set dates (daily variation margin
  alone is not `STM`; the qualification is the user's). Its remaining
  maturity M is then the time to the next reset date, floored at 10 business
  days in the maturity factor [CRE52.37(5), CRE52.48; CRR Art. 279c(1)(a),
  (2)], while S, E and the option expiry T keep their meaning, so a settled
  option can expire after its reset date. The reset date must be after the
  reporting date (the MtM is the value after that day's settlement) and not
  after the maturity as entered; a reset on the maturity date is accepted.
  Blank or `N` with a reset date, `STM` without one, another code, a reset
  that is not a date, and legs of one economic trade that differ in
  settlement type (blank and `N` alike) or reset date are rejected, the
  trade `EXCLUDED` and its netting set `INCOMPLETE`. An `STM` trade in a
  margined netting set makes it `INVALID: STM requires unmargined treatment;
  mixed margined/unmargined netting sets are not supported`, a limitation of
  the engine (an all-`STM` set is misclassified; a mixed set needs the
  CRE52.74 treatment, not implemented). With both columns blank every trade
  keeps its result. The template has the columns, their dropdown and date
  validation, and a sentence on README D19; the status formulas of Results,
  Reconciliation and TradeReconciliation C5 watch `Trades!$C:$AC`.
  Upgrade: a workbook built from an earlier template stops at the Trades
  header check until it is rebuilt from the new template, or two columns
  are inserted before Comment with the new headers and `Trades!$C:$AA` is
  changed to `Trades!$C:$AC` in those three formulas. Tests: the
  illustrative fixtures `illustrative-settled-to-market-daily` (equity index
  future and call reset the next business day: MF 0.2) and
  `illustrative-settled-to-market-periodic` (a five-year swap, an equity call
  and a Brent forward reset in 30 days, the swap keeping its bucket by E, the
  forward's reset on its maturity), each under CRR and BCBS, with values from
  `tests/derivations/settled_to_market.py`; 20 `TEST_InputValidation` cases
  for the rules. The fixture format gains `settlement_type` and
  `next_reset_date`, and the expected format the trade quantities
  `remaining_maturity` and `maturity_factor` (unmargined netting sets only).
  The R cross-check translates a settled-to-market trade with M as the time
  to its reset, and skips the cases it cannot represent. Totals: `TEST_Cases`
  28 cases and 131 checks, `TEST_InputValidation` 163 cases and 267 checks.

### Fixed

- `tools/Sync-SACCR-VBA.ps1` drops exported member attributes, such as
  `Attribute Worksheet_Change.VB_Description = "..."`, from the code it
  writes into a document module (#394). They are not valid as code, so a
  synchronized worksheet or workbook module with one would not compile. The
  lines around them, and comments and strings that name `Attribute`, are
  kept, with or without a module header, CRLF or LF, and for member names
  with code-page letters such as `Évaluer`. The header scan now matches
  only `Attribute VB_<name> =`, so a member whose name starts with `VB_` can
  no longer be taken for the header and cut the code above it. No module in
  the repository has a member attribute today. Synchronization only: no
  SA-CCR calculation changes.
- `tools/check_test_cases.py` validates the fields of a numerical fixture
  that depend on each other with the engine's rules (#401). Each field was
  checked alone, so a fixture the engine would reject could pass the gate
  while its expected figures described a calculation that never happens: an
  ordinary trade with a next reset date, a settled-to-market trade without
  one, or with one on or before the valuation date or after M (the maturity
  date, else the end date, else the option expiry); a settled-to-market trade
  in a margined netting set; and, for the CCP inputs of #112, `qccp` or
  `ccp_retains_variation_margin` missing for role `ccp` or given for another
  role, or retained VM on an unmargined netting set. Such an input is now
  accepted only where every expected file of the fixture declares the
  engine's outcome, `trade_status` `EXCLUDED` or a `netting_set_status`
  starting with `INVALID`; a reset on M stays valid. `TEST_CASES.md` states
  the rule. Tooling only: no fixture, expected value or calculation changes.

### Documentation

- Independent review is no longer required anywhere: the project has one
  maintainer and no reviewer (decision 8 of the clearing coverage). An
  `independent` value is now one computed by an external implementation the
  maintainer did not write, registered as a source (for example the R
  package SACCR of #352); it names that implementation in `derived_by`, and
  `reviewed_by` may be null. `tools/check_test_cases.py` no longer requires
  `reviewed_by` for it, and now requires its `source` to be an external
  implementation, registered with an ID starting `IMPL-`; a `published`
  value cannot cite one. Decision 8, the methodology, the tests README,
  the `derivation` texts of the expected files and the reference-class
  comment in `TEST_CaseRunner` say the same (#411). `TEST_CASES.md` says
  that acceptance criteria ask for `published` values where a registered
  source prints them and otherwise for `illustrative` values from a
  derivation separate from the code under test, never for independent
  review; CONTRIBUTING and the bug and feature issue templates say the
  same.
- The user manual for 1.0.2 is prepared in `docs/user-manual/v1.0.2/`,
  from the 1.0.1 manual: trade exposures to a CCP and the NettingSets
  columns `QCCP (Y/N)` and `CCP retains VM, collateral not protected (Y/N)`;
  settled-to-market trades and the Trades columns `Settlement type (N/STM)`
  and `Next reset date`, with Comment moving two columns right; the role
  `MULTI`; the INFO lines on Checks for client clearing and CCPs and the CCP
  subtotal per regime; the highlighted Results rows; moving inputs from
  1.0.1; the known limitations; the test counts; and CRE50 and CRR Articles
  300–311 among the sources. The worksheet-function table of section 10.2,
  whose signatures overflowed into the next column in the 1.0.0 and 1.0.1
  PDFs, is a list with the full signatures, `SACCR_MaturityFactor`'s
  included, and a note that arguments in square brackets are optional
  (#379); the built PDF has no line wider than the page. The 1.0.0 and
  1.0.1 manuals are not changed.
- `docs/methodology/CLEARING_COVERAGE.md` specifies the input contract for
  CCP trade exposures (#112), which the NettingSets entry under Added
  implements. Role `CCP` becomes the
  clearing member's trade exposure to a CCP, calculated with SA-CCR, with two
  NettingSets columns after Margin agreements: *QCCP (Y/N)* in R and *CCP
  retains VM, collateral not protected (Y/N)* in S, required for `CCP` and
  blank otherwise. A margined `CCP` set has an MPOR floor of 10 business days
  in both regimes; with retained VM each trade takes the larger of that MPOR
  and the lower of one year and its remaining maturity, floored at 10
  business days. The combined large-or-illiquid flag stays, as a
  conservative limitation; client trades the member need not reimburse are
  not entered. QCCP status follows the netting set's regime (EMIR under
  CRR, CRE50 under BCBS). Checks names each `CCP` set's role, MPOR rule and
  risk-weight rule (2% for a QCCP, stated and not applied; for a
  non-qualifying CCP, a standardised weight to be determined outside
  K-SACCR), and each regime's CCP EAD subtotal. Default-fund contributions, own funds requirements and the
  Art. 303(2) cap are left for a later issue. Decisions 9 to 14 record the
  choices.
- `docs/methodology/CLEARING_COVERAGE.md` records the input contract for
  settled-to-market trades (#360), which the Trades entry under Added
  implements: the columns, the seven rules and their rationale, the upgrade
  and the tests. The methodology's scope, status and traceability tables,
  `TEST_CASES.md` and the README's capability table describe it.
- The README's Regimes section says why the toolkit has two regimes: `CRR`
  is the default and binding for EU institutions; `BCBS` does not replace it
  but shows the effect of the EU deviations, traces results to the Basel
  standard and supports Basel monitoring and non-EU group entities, with a
  link to the wiki's "Why two regimes".
- The source descriptions are more exact. The consolidated CRR text has no
  legal effect; the methodology implements Articles 274–280f, and Articles
  273a, 281 and 282 are context only. The client's 10-day MPOR floor under
  `BCBS` cites the conditions of CRE54.15, the 4% risk weight of
  CRE54.16, the bilateral treatment of CRE54.17 and the OTC scope of
  CRE54.8(2). BCBS 279 (March 2014, revised April 2014) is now copied in
  `docs/methodology/sources/` with its hash, as the BIS terms allow
  non-commercial redistribution; the sources README says the copies stay
  under their publishers' terms, not the repository's licence. Only `published` and `independent` values qualify as
  numerical-validation references; CRR provisions were checked against the
  registered sources, and CRR-specific numerical validation remains
  incomplete.
- The `RunSACCR_Silent` result line in `INSTALLATION.md` gives the input
  fingerprint the demonstration portfolio has today, `3B2754CA`, as seen in a
  workbook built from the template with the CCP columns of NettingSets
  (#112), whose headers are part of the fingerprint. The 1.0.1 release
  workbook gives `7B68D18F`; `74A6C31F` and `409B031E` (the template with
  the settled-to-market columns, #360) were recorded before later changes to
  the template's inputs. The 1.0.0 and 1.0.1 user manuals are released and
  keep the older value in an example they label illustrative.
- The README shows the total downloads of the release assets as a badge
  linked to the latest release.
- The website's download button starts the download of the 1.0.1 workbook
  instead of opening the release page, a second button downloads the user
  manual (PDF), and the SHA-256 note links `SHA256SUMS.txt`. A short script
  starts these downloads in a hidden frame, so the page no longer visibly
  passes through GitHub; without JavaScript the links work as before.
- `docs/methodology/CLEARING_COVERAGE.md` records what K-SACCR does for
  client clearing, CCPs and listed derivatives (#358): the arrangements and
  their MPOR floors under both regimes, QCCP and non-QCCP treatment,
  products, collateral and segregation, settlement and reset, the mapping of
  each supported rule to its input, code and test, the cases still to be
  rejected, and the owner's decisions: a `BCBS` client keeps 10 days as a
  conservative interpretation of CRE54.2; CCP-facing netting sets get their
  own role, QCCP, retained-VM and protection inputs, with CCP trade EAD in the
  EAD totals and risk weights and capital outside them; the single *large or
  illiquid* flag stays as a conservative limitation; settled-to-market trades
  get a settlement type and a next reset date; and a declared CCP-facing
  netting set is rejected until CCP support exists, an interim rejection the
  CCP trade exposures under Added replace. A netting set facing a CCP but entered with role N
  is still calculated as bilateral, as the support matrix says. The source
  register adds `CRR-CCP`, Articles 4(1)(88)–(91) and 300–311 of the
  consolidated CRR, with a registered extract. The documentation itself
  changes no calculation.

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

- The workbook's own texts no longer refer to the source repository or its
  issues, for a user who receives only the workbook. TestCatalogue calls its
  computed illustrative values "separate unreviewed calculation" instead of
  "Python value, not in repository", says a source was "not checked" rather
  than "not registered", and points to the input-validation tests run before
  release instead of a test module and an issue number; Basel_vs_CRR drops
  the issue numbers (#38, #39) and the test-module name. README C3 drops
  "VBA engine v1.0.0", a version stamp that was no longer true. Only the
  text of 13 cells changes.

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

[Unreleased]: https://github.com/danielep71/VBA-SACCR-Toolkit/compare/v1.0.2...HEAD
[1.0.2]: https://github.com/danielep71/VBA-SACCR-Toolkit/releases/tag/v1.0.2
[1.0.1]: https://github.com/danielep71/VBA-SACCR-Toolkit/releases/tag/v1.0.1
[1.0.0]: https://github.com/danielep71/VBA-SACCR-Toolkit/releases/tag/v1.0.0
