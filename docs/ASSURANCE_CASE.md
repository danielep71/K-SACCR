# Security assurance case

This document argues that K-SACCR meets its security requirements. It states
those requirements, the threats and the trust boundaries they assume, how
secure design principles are applied, and how the common weaknesses of this
kind of software are countered, with the evidence for each claim.
[`SECURITY.md`](../SECURITY.md) is the policy: how to report a vulnerability,
which versions are supported and the runtime risk surface. This document is
the argument behind it.

It covers the security of the software and of its delivery. It does not argue
that an exposure figure is correct: the numerical evidence is in the
[methodology](methodology/README.md) and the
[test-case policy](methodology/TEST_CASES.md).

## 🎯 Security requirements

What a user can expect from the workbook and its releases:

| # | Requirement |
| --- | --- |
| R1 | **Confidentiality.** The workbook sends no data anywhere: it has no network access, and it writes files only where the user chooses in a save dialog. |
| R2 | **No code from data.** Nothing the workbook reads (a CSV file, a cell, a parameter) is executed as code or as an Excel formula. |
| R3 | **Input integrity.** A CSV file is accepted or rejected as a whole against a published contract, and staged rows reach the input sheets only after a validation that still holds and an explicit confirmation. |
| R4 | **Result integrity.** Results that no longer match their inputs are shown as out of date, and the evidence export refuses them, so stale or hand-edited results are not exported as a run. |
| R5 | **Bounded resources.** A malformed or oversized file is rejected before it can exhaust memory or time. |
| R6 | **Host settings.** The workbook changes no Excel Trust Center setting, installs nothing and loads no library beyond the one Windows function it declares. |
| R7 | **Delivery integrity.** Releases are delivered over HTTPS with published SHA-256 sums; the source is public and every change reaches it through a reviewed pull request with required checks. |
| R8 | **No secrets in the repository.** The repository and the releases hold no credentials and no real portfolio data. |

## 🧭 Assets, threats and trust boundaries

**Assets:** the user's portfolio data and results inside the workbook; the
integrity of the figures and of the evidence exported from them; the user's
Excel and Windows session; the repository, its releases and the credentials
of its automation.

**Threats considered:**

| Actor | Can | Aim |
| --- | --- | --- |
| Author of a CSV file the user imports | Control every byte of the file | Run code or formulas, corrupt the inputs, exhaust the session |
| Contributor or attacker with a pull request | Control the PR's files | Run code with repository credentials, slip a change past review |
| Compromised dependency or GitHub Action | Change code the CI runs | Run code in CI, exfiltrate credentials |
| Anyone reading the public repository | Read history and files | Find credentials or real data |

**Out of scope**, as in [`SECURITY.md`](../SECURITY.md#scope): a modified
copy of the workbook, malicious VBA already trusted in the same Excel
process, and vulnerabilities of Excel, Office, Windows, GitHub or Python
themselves. The user decides to enable the workbook's macros: Excel's trust
model is the boundary there, and K-SACCR does not try to replace it.

**Trust boundaries:**

1. **File to workbook.** Every CSV file is untrusted. It crosses into the
   workbook only through `CORE_CsvReader`, which applies the
   [CSV contract](csv/CSV_CONTRACT.md) before anything is written to a sheet.
2. **Staging to inputs.** Staged rows are untrusted until a validation of
   both staging sheets, against each other and against Params, is `VALID` and
   still current; only then can **Commit import** write the input sheets,
   after a confirmation.
3. **Sheets to engine.** Values typed on the input sheets are validated by
   the engine at every run; an invalid value makes its trade or netting set
   invalid, never a silently different figure.
4. **Pull request to CI.** Code from a pull request runs only in jobs with a
   read-only token and no secrets. The one workflow that runs with write
   access to pull requests (`pr-milestones.yml`, `pull_request_target`) never
   checks out or runs the pull request's code.
5. **Release to user.** Releases leave the repository through GitHub over
   HTTPS, with `SHA256SUMS.txt` for each asset.

## 🧱 Secure design principles

The principles of Saltzer and Schroeder, as applied here:

| Principle | How K-SACCR applies it | Evidence |
| --- | --- | --- |
| Economy of mechanism | One CSV reader for every path (check, import, staging validation), generated from one machine-readable contract; one engine for every run | `docs/csv/contract.json`, `tools/generate_csv_vba.py`, the `generated-csv-modules` gate |
| Fail-safe defaults | A file is rejected as a whole on any finding; an unreadable value makes a trade invalid rather than taking a default; a status that cannot be read leaves a Ribbon control enabled while the macro still refuses the operation; the commit confirmation defaults to Cancel | [CSV contract](csv/CSV_CONTRACT.md), `TEST_InputValidation`, `TEST_Import` |
| Complete mediation | Every run validates every input row; every commit re-checks that the validation is still current; every evidence export re-checks that the results match their inputs | `M_Main.StagingValidationCurrent`, `M_Main.ResultsStatus`, `TEST_MainState`, `TEST_Import` |
| Open design | All source, contracts, tests and evidence rules are public; no secret is part of the design | This repository |
| Separation of privilege | A change to the release branch needs a pull request and the required check **Repository integrity**; VBA changes also need an Excel run by the owner before merge | [`docs/GOVERNANCE.md`](GOVERNANCE.md), [`CONTRIBUTING.md`](../CONTRIBUTING.md#validation-and-evidence) |
| Least privilege | Workflows default to no or read-only permissions and grant write scopes per job; the traffic export uses a GitHub App with only Administration: read, a short-lived token and an environment secret seen by one step | `.github/workflows/`, [`tools/README.md`](../tools/README.md#github-actions) |
| Least common mechanism | The workbook shares no state with other workbooks: every callback acts on `ThisWorkbook`; the Ribbon pointer is kept per Excel session and removed before every save | [`docs/workbook/RIBBON.md`](workbook/RIBBON.md), `M_Ribbon` |
| Psychological acceptability | Findings name the CSV line, the record and the field and link to the staged cell; protection guards against mistakes without a password the user must manage | [`docs/workbook/WORKBOOK_PROTECTION.md`](workbook/WORKBOOK_PROTECTION.md) |

## 🛡️ Common weaknesses countered

| Weakness | Countermeasure | Evidence |
| --- | --- | --- |
| CSV formula injection (CWE-1236) | A field of any type other than a number that starts with `=`, `+`, `-`, `@`, a tab or a carriage return is rejected (`CSV-FORMULA`); staged values are written as text into cells formatted as text, so nothing is evaluated | CSV case `formula-starts`; `TEST_CsvReader`; `TEST_Import` |
| Improper input validation (CWE-20) | Every CSV field is checked against its type, an allowlist of codes, its range and its length; rules between fields, between records and against Params follow; the engine validates the sheets again at every run | [CSV contract](csv/CSV_CONTRACT.md); `TEST_CsvReader` (every fixture and case the Python reference pins); `TEST_InputValidation` |
| Encoding errors (CWE-172, CWE-838) | Strict UTF-8 decoding: invalid, overlong and surrogate sequences and NUL characters reject the file | CSV cases `utf8-*`, `nul-on-line-3` |
| Uncontrolled resource consumption (CWE-400, CWE-770) | The file size is checked against 128 MiB before the file is read; record counts are capped per file; at most 10,000 findings are listed and marked | CSV cases `too-large`, `too-many-records`, `too-many-empty-lines`; `CSV-TOO-LARGE` |
| Code or command injection (CWE-78, CWE-94) | The sources contain no `Shell`, no `CreateObject`, no network access and no evaluation of text as code; the only file opened in another application is the import specification the workbook itself wrote, and only after the user agrees | [`SECURITY.md`](../SECURITY.md#runtime-risk-surface) |
| Path manipulation (CWE-22, CWE-73) | No file path comes from data: every file is chosen in an Excel open or save dialog; the test suites write their own files to `%TEMP%` under names they choose | [`SECURITY.md`](../SECURITY.md#runtime-risk-surface) |
| Memory safety (CWE-119) | VBA and Python are memory-safe. The one Windows call, `RtlMoveMemory`, copies only a pointer the same Excel session stored, checked against that session's window handle | `M_Ribbon.CurrentRibbon`; `TEST_MainState` case `ribbon.pointer` |
| Integrity of results (CWE-345) | An input fingerprint binds results to their inputs; out-of-date or edited results are flagged and refused by the evidence export; the evidence record binds logs to a commit with SHA-256 digests | `CORE_Engine.InputFingerprint`, `M_Main.ResultsStatus`; [`docs/EXCEL_EVIDENCE.md`](EXCEL_EVIDENCE.md) |
| Untrusted code in CI and the supply chain (CWE-494, CWE-829) | Actions pinned by full commit SHA with version comments; Python tools installed from hash-locked requirements; Dependabot alerts and updates; actionlint and CodeQL on the workflows; `pull_request_target` never runs PR code | `tools/check_workflows.py`; `tools/requirements-quality-ci.txt`; [`docs/GOVERNANCE.md`](GOVERNANCE.md) |
| Credentials in the repository (CWE-798, CWE-522) | No credential is stored in the repository; secret scanning and push protection are enabled; the only automation key is an environment secret, rotated as `SECURITY.md` prescribes | [`SECURITY.md`](../SECURITY.md#runtime-risk-surface); [`docs/GOVERNANCE.md`](GOVERNANCE.md) |
| Static analysis gaps | CodeQL covers the Python tooling and the workflows; VBA, which CodeQL does not support, is covered by the repository's own gates: undeclared identifiers, call resolution, dependency direction, jump targets, conditional compilation and public API | [`tools/README.md`](../tools/README.md#gates) |

## ⚠️ Residual risks

- **Macros must be enabled.** The workbook is VBA: enabling its macros
  trusts the file. A modified copy is outside this case; users should obtain
  the workbook from the GitHub release and check it against
  `SHA256SUMS.txt`.
- **Protection is not a security boundary.** The sheets are protected
  without a password, against accidental edits only.
- **Releases are not yet signed.** They are delivered over HTTPS with
  SHA-256 sums; a cryptographic signature of the release assets is planned.
- **One maintainer.** Every change is reviewed by automated checks and by
  the owner, not by a second person.
- **No dynamic analysis tool.** There is no fuzzer; the CSV cases and the
  input-validation suite cover malformed input, and randomized inputs are
  planned in #165.

## 🔁 Maintenance

This case is reviewed with every change to [`SECURITY.md`](../SECURITY.md), to
a trust boundary above (a new way data enters the workbook, a new file
written, a new workflow with write access or a secret) and before each
release. Last reviewed: 2026-11-12, on `release/1.3.0`.
