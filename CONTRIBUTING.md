<div align="center">

# 🤝 Contributing to K-SACCR

### Counterparty credit risk under SA-CCR, built in Excel/VBA

[![Conduct](https://img.shields.io/badge/Conduct-Required-6f42c1?style=flat-square)](CODE_OF_CONDUCT.md)
[![Security](https://img.shields.io/badge/Security-Private_reporting-d73a49?style=flat-square)](SECURITY.md)
[![Evidence](https://img.shields.io/badge/Evidence-Exact_source-0969da?style=flat-square)](#validation-and-evidence)
[![Workflow](https://img.shields.io/badge/Workflow-PR_into_release-217346?style=flat-square)](#development-workflow)

<br>

**Focused scope · Reviewable source · Reproducible evidence · Honest limitations**

</div>

---

This document is authoritative for the **contribution and review workflow**.
Branch, issue and review policy is owned by
[`docs/GOVERNANCE.md`](docs/GOVERNANCE.md); developer setup by
[`INSTALLATION.md`](INSTALLATION.md); vulnerability handling by
[`SECURITY.md`](SECURITY.md); publication by [`RELEASING.md`](RELEASING.md).

Participation is governed by the [Code of Conduct](CODE_OF_CONDUCT.md).
Suspected vulnerabilities must never be disclosed in an issue or pull request.

> [!NOTE]
> The SA-CCR engine is in `src/`: CRR baseline, Basel CRE52 also supported.
> Its regulatory basis and the numerical test-case format are in
> [`docs/methodology/`](docs/methodology/README.md). The regression harness is
> `TEST_Harness.RunTests`; it and the other test suites are listed under
> [running the harness](INSTALLATION.md#running-the-harness). The source
> layout and its known deviations are in
> [`docs/REPOSITORY_STRUCTURE.md`](docs/REPOSITORY_STRUCTURE.md).

## 🌱 Ways to contribute

| Contribution | First action |
| --- | --- |
| 🐛 Reproducible defect | Open an issue with exact commit, environment, expected and observed behavior, and evidence. |
| ✨ Feature or calculation change | Open an issue defining the regulatory basis, observable contract, compatibility, non-goals and validation. |
| 🧪 Test or evidence improvement | Explain provenance, independence, coverage and what failure the test detects. |
| 📖 Documentation | Identify the authoritative document being corrected. |
| ⚙️ Tooling or governance | Explain failure behavior, portability, trust boundary and maintenance cost. |
| 🔐 Security concern | Follow [`SECURITY.md`](SECURITY.md) privately. |

Open issues with the forms (bug, feature or calculation change, documentation).
Every issue is assigned to `danielep71`, carries one priority label (`P1`, `P2`
or `P3`) and has the relevant milestone; the forms set the first two, and triage
sets the milestone. See
[issue metadata and completion](docs/GOVERNANCE.md#issue-metadata-and-completion).
Pull requests follow [the PR template](.github/PULL_REQUEST_TEMPLATE.md).

<a id="development-workflow"></a>

## 🌿 Development workflow

1. Start from the active release branch, currently `release/1.0.0`. Never assume
   GitHub's default branch is the right base.
2. Create one focused task branch with a descriptive name in the repository
   convention: `fix/<issue>-<slug>`, `docs/<slug>` or `chore/<slug>`.
3. Reproduce the current behavior before changing it.
4. Define the observable contract, affected callers, compatibility impact and
   validation plan.
5. Make the smallest coherent change; avoid unrelated formatting, generated
   output or opportunistic refactoring.
6. Run `python tools/check.py` before every push. All gates must pass. Check
   the committed candidate with `python tools/check.py --ci --base <base-sha>`.
7. Update the affected documentation and add a user-facing entry under
   `## [Unreleased]` in [`CHANGELOG.md`](CHANGELOG.md).
8. Review the complete diff, then open a pull request **against the release
   branch** with evidence and explicit limitations.
9. Pull requests into the release branch are squash-merged once the
   **Repository integrity** check is green, review findings are resolved, and
   every review conversation, including the automatic Codex review, is
   resolved. Merged branches are deleted automatically.

Nothing is committed directly to the release branch or to `main`. Integration
of the release branch into `main` happens only on the owner's request; see
[`RELEASING.md`](RELEASING.md).

### Commit messages

Write an imperative subject line, then a body that states:

- **what** changed;
- **why** it changed; and
- **what was validated** at that commit, separating static checks from Excel
  runs.

Reference the issue when one exists. Do not place credentials, private links,
attribution boilerplate or unverifiable test claims in commit messages.

## 📦 Source-change discipline

The repository-control files define how source is stored:

- exported VBA (`.bas`, `.cls`, `.frm`) is the reviewable source of truth. It
  is stored with LF in Git and checked out with CRLF;
- component filenames match their `VB_Name`, and every module declares
  `Option Explicit`;
- `.frm`/`.frx` pairs stay together, and `.frx` stays binary;
- workbooks and add-ins are generated artifacts, ignored by Git unless an exact
  path is re-included; a workbook is never the only record of a source change;
  and
- test data and examples are synthetic.

`python tools/check.py` enforces the storage, naming, placement and
`Option Explicit` rules, plus the committed-whitespace check, the VBA jump,
conditional-compilation and public-API checks, the validation of the test-case
files and the check that the generated `TEST_Cases.bas` is current.
Where each component belongs, the dependency direction and the public-API
boundary are defined in
[`docs/REPOSITORY_STRUCTURE.md`](docs/REPOSITORY_STRUCTURE.md); supported
declarations are listed in [`docs/PUBLIC_API.txt`](docs/PUBLIC_API.txt). Naming,
contracts, errors and Excel-state handling follow
[`docs/VBA_HOUSE_STYLE.md`](docs/VBA_HOUSE_STYLE.md); import and export follow
[`INSTALLATION.md`](INSTALLATION.md#importing-vba-into-excel).

Do not weaken a calculation, numerical or packaging gate merely because the
generic repository gate passes.

## 🔄 Compatibility and calculation contracts

A change to documented procedures, functions, classes, enums, parameters,
defaults, return values, errors, side effects, input or output layouts, or
supported platforms is a contract change. So is any change that alters a
computed exposure, add-on, replacement cost or multiplier for the same inputs.

Such a contribution must identify callers and migration impact, update
regression coverage, update the user documentation and state the release
impact. A calculation change must also cite the regulatory provision or
published example it implements.

Excel and Windows state belongs to the caller or host unless a component
explicitly owns it. Capture before mutation, restore only state successfully
changed and still owned, and never let cleanup conceal the original failure.

<a id="validation-and-evidence"></a>

## 🧪 Validation and evidence

Validation must be reproducible from the exact source under review. Record:

```text
Source
------
Commit:
Files or components changed:

Environment
-----------
Excel / Office build:
Office bitness:
Operating system:
Locale / date system:

Checks
------
python tools/check.py:
Compile (Debug → Compile VBAProject):
Regression harness:
Specific scenario verified:
Cleanup:

Limitations
-----------
Not run or unverified:
Follow-up:
```

Use only the applicable fields, but never omit a material limitation.

- **Static checks are not Excel evidence.** `tools/check.py` and the hosted
  workflow inspect files. They do not compile VBA or run Excel. Evidence that
  must be bound to an exact commit follows
  [`docs/EXCEL_EVIDENCE.md`](docs/EXCEL_EVIDENCE.md).
- **VBA changes are not merged until verified in Excel:** compile, run the
  harness, and exercise the specific scenario the change addresses. State
  plainly what was and was not run in Excel.
- **A skipped check is not a pass.**
- **Reference values must be independent** of the implementation under test,
  for example published regulatory examples or an independent calculation.

When sharing a VBA module for review or import, send the CRLF file from a Git
checkout. Do not use GitHub's **Raw** view, which serves the LF form stored in
Git.

<a id="pull-requests"></a>

## 🚀 Pull requests

A pull request answers five questions:

1. What problem does this solve?
2. What observable contract changes?
3. What remains compatible?
4. How was the exact source validated?
5. What remains unverified?

### Review checklist

- [ ] The PR targets the active release branch, and the related issue is linked.
- [ ] Scope is focused; calculation, API and compatibility impact are assessed.
- [ ] Exported VBA and required binary companions are synchronized.
- [ ] `python tools/check.py` passes, and the **Repository integrity** check is green.
- [ ] Excel compile, harness and scenario results are recorded for VBA changes.
- [ ] Error, boundary, recovery and cleanup paths are covered where affected.
- [ ] Documentation and the `[Unreleased]` changelog entry are updated.
- [ ] No confidential, real-portfolio, accidental binary or generated material is added.
- [ ] Unverified environments and skipped checks are stated plainly.
- [ ] The final diff contains no unrelated formatting or local artifacts.

The minimum review checklist is in
[`docs/GOVERNANCE.md`](docs/GOVERNANCE.md#ownership-authorization-and-review).
Discussion stays technical and respectful under the
[Code of Conduct](CODE_OF_CONDUCT.md).

## 📚 Where detailed rules live

| Need | Document |
| --- | --- |
| Source layout and public API | [`docs/REPOSITORY_STRUCTURE.md`](docs/REPOSITORY_STRUCTURE.md) |
| VBA naming, contracts and errors | [`docs/VBA_HOUSE_STYLE.md`](docs/VBA_HOUSE_STYLE.md) |
| Regulatory basis and test cases | [`docs/methodology/`](docs/methodology/README.md) |
| Commit-bound Excel evidence | [`docs/EXCEL_EVIDENCE.md`](docs/EXCEL_EVIDENCE.md) |
| Branches, issues, review, GitHub controls | [`docs/GOVERNANCE.md`](docs/GOVERNANCE.md) |
| Developer setup and import | [`INSTALLATION.md`](INSTALLATION.md) |
| Static checks and CI | [`tools/README.md`](tools/README.md) |
| Issue labels | [`docs/LABELS.md`](docs/LABELS.md) |
| Vulnerability reporting | [`SECURITY.md`](SECURITY.md) |
| Release procedure | [`RELEASING.md`](RELEASING.md) |
| Change history | [`CHANGELOG.md`](CHANGELOG.md) |

## 📄 Licensing and maintainer

This project is distributed under the [Mozilla Public License 2.0](LICENSE).
Contributions are accepted under the same license. Contributors must have the
right to submit every part of a contribution, including code, tests, data,
images and generated material.

Maintained by **Daniele Penza**.

---

**Contribution principle:** make the contract explicit, keep the diff focused,
and leave evidence another person can reproduce.
