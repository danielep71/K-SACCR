# 🤝 Contributing to K-SACCR

Participation is governed by the [Code of Conduct](CODE_OF_CONDUCT.md).
Suspected vulnerabilities must never be disclosed in an issue or pull request;
follow [`SECURITY.md`](SECURITY.md).

The SA-CCR engine is in `src/` (CRR baseline, Basel CRE52 also supported); the
test suites are listed under [running the harness](INSTALLATION.md#running-the-harness).

## 🌱 Ways to contribute

| Contribution | First action |
| --- | --- |
| 🐛 Reproducible defect | Open an issue with exact commit, environment, expected and observed behavior, and evidence. |
| ✨ Feature or calculation change | Open an issue defining the regulatory basis, observable contract, compatibility, non-goals and validation. |
| 🧪 Test or evidence improvement | Explain provenance, independence, coverage and what failure the test detects. |
| 📖 Documentation | Identify the authoritative document being corrected. |
| ⚙️ Tooling or governance | Explain failure behavior, portability, trust boundary and maintenance cost. |

Open issues with the forms. Issue metadata (assignee, exactly one priority
label, milestone) is described in
[`docs/GOVERNANCE.md`](docs/GOVERNANCE.md#issue-metadata-and-completion).

<a id="development-workflow"></a>

## 🌿 Development workflow

1. Start from the active release branch named in
   [`docs/GOVERNANCE.md`](docs/GOVERNANCE.md#branches). Never assume GitHub's
   default branch is the right base.
2. Create one focused task branch: `fix/<issue>-<slug>`, `docs/<slug>` or
   `chore/<slug>`.
3. Reproduce the current behavior, then make the smallest coherent change;
   avoid unrelated formatting, generated output or opportunistic refactoring.
4. Run `python tools/check.py` before every push; all gates must pass. The
   gates are described in [`tools/README.md`](tools/README.md).
5. Update the affected documentation and add a user-facing entry under
   `## [Unreleased]` in [`CHANGELOG.md`](CHANGELOG.md).
6. Open a pull request **against the active release branch** using
   [the PR template](.github/PULL_REQUEST_TEMPLATE.md), which holds the review
   checklist.
7. Keep the task branch up to date with the target release branch before
   merging. If the target advances, incorporate its latest changes and rerun
   validation. The pull request is squash-merged once **Repository integrity**
   is green on the updated candidate and every review conversation, including
   the automated code review, is resolved. For a VBA or template change, the
   owner's Excel verification still counts after such an update only if the
   incoming changes touch nothing in `src/`, `tests/modules/` or the
   template; otherwise compile and the affected suites are run again.

### Commit messages

Write an imperative subject line, then a body that states:

- **what** changed;
- **why** it changed; and
- **what was validated** at that commit, separating static checks from Excel
  runs.

Reference the issue when one exists. Do not place credentials, private links,
attribution boilerplate or unverifiable test claims in commit messages.

## 📦 Source and compatibility

Exported VBA (`.bas`, `.cls`, `.frm`) is the reviewable source of truth; a
workbook is never the only record of a source change. Encoding, line endings,
forms and committed workbooks follow
[export compatibility](docs/VBA_HOUSE_STYLE.md#export-compatibility); Excel
state follows [operation and cleanup](docs/VBA_HOUSE_STYLE.md#operation-and-cleanup).
Test data and examples are synthetic.

A change to documented procedures, parameters, defaults, return values, errors,
side effects, input or output layouts, or supported platforms is a contract
change, and so is any change that alters a computed exposure for the same
inputs. It must identify callers and migration impact, update regression
coverage and user documentation, and state the release impact. A calculation
change must cite the regulatory provision or published example it implements.

<a id="validation-and-evidence"></a>

## 🧪 Validation and evidence

Validation must be reproducible from the exact source under review. Record it
with the [record template](docs/EXCEL_EVIDENCE.md#record-template) of
`docs/EXCEL_EVIDENCE.md`.

- **Static checks are not Excel evidence.** `tools/check.py` and the hosted
  workflow inspect files. They do not compile VBA or run Excel.
- **VBA changes are not merged until verified in Excel:** compile, run the
  suites, and exercise the specific scenario the change addresses. State
  plainly what was and was not run.
- **A skipped test is not a pass.**
- **Expected values never come from running the implementation under test.**
  They come from published regulatory examples, from an external
  implementation registered as a source, or, as `illustrative` values, from a
  derivation separate from the code under test. No independent reviewer is
  required: the project has one maintainer
  ([test-case policy](docs/methodology/TEST_CASES.md#reference-classes)).

## 📚 Where detailed rules live

| Need | Document |
| --- | --- |
| Source layout and public API | [`docs/REPOSITORY_STRUCTURE.md`](docs/REPOSITORY_STRUCTURE.md) |
| VBA naming, contracts and errors | [`docs/VBA_HOUSE_STYLE.md`](docs/VBA_HOUSE_STYLE.md) |
| Regulatory basis and test cases | [`docs/methodology/`](docs/methodology/README.md) |
| Commit-bound Excel evidence | [`docs/EXCEL_EVIDENCE.md`](docs/EXCEL_EVIDENCE.md) |
| Branches, issues, labels, GitHub settings | [`docs/GOVERNANCE.md`](docs/GOVERNANCE.md) |
| Pull requests and review checklist | [`.github/PULL_REQUEST_TEMPLATE.md`](.github/PULL_REQUEST_TEMPLATE.md) |
| Developer setup and import | [`INSTALLATION.md`](INSTALLATION.md) |
| Static checks and CI | [`tools/README.md`](tools/README.md) |
| Vulnerability reporting | [`SECURITY.md`](SECURITY.md) |
| Release procedure | [`RELEASING.md`](RELEASING.md) |
| Change history | [`CHANGELOG.md`](CHANGELOG.md) |

## 📄 Licensing

This project is distributed under the [Mozilla Public License 2.0](LICENSE).
Contributions are accepted under the same license. Contributors must have the
right to submit every part of a contribution, including code, tests, data,
images and generated material. Maintained by @danielep71.
