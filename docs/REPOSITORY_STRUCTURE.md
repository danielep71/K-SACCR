# 🗂️ Repository Structure

[![Profile: application](https://img.shields.io/badge/profile-application-217346)](#profile-decision)
[![Model: source first](https://img.shields.io/badge/model-source--first-0969da)](#repository-layout)
[![Boundary: facade over core](https://img.shields.io/badge/boundary-facade%20over%20core-6f42c1)](#public-api-boundary)

This document is authoritative for **the project profile, where every durable
artifact belongs and the boundary between the calculation core, the public
interface and the Excel host**. It adapts
`docs/REPOSITORY_STRUCTURE.md` from EXCEL-VBA-PROJECT-TEMPLATE to SACCR.
Import order and host support are owned by [`INSTALLATION.md`](../INSTALLATION.md);
how VBA is written by [`VBA_HOUSE_STYLE.md`](VBA_HOUSE_STYLE.md).

<a id="profile-decision"></a>

## 🧭 Profile decision

**K-SACCR uses the `application` profile.** Decided by the owner on 2026-10-06
in issue #3.

| Question | Answer for K-SACCR |
| --- | --- |
| Who calls it? | A user working in the K-SACCR workbook: entering trades and netting sets, running the calculation, reading results. |
| What does it own? | The workbook: input and result sheets, the calculation engine and their startup and packaging. |
| Lifecycle | The workbook is the deliverable; it is built from the exported source in this repository. |
| Supported environments | Excel for Windows; see [supported hosts](../INSTALLATION.md#supported-hosts). |

**Why not `library`:** the milestone plan defers business UI and production
workbook packaging to later milestones, so the end product is a workbook, not
only callable functions. Choosing `library` now would force a profile change
later.

**What stays library-like:** the calculation core is host-independent. It never
touches workbook objects, so the engine remains testable without a workbook and
reusable if a library deliverable is ever wanted.

`.github/label-policy.json` already selects the `application` profile for the
label catalogue; this document makes it the architecture decision as well.

<a id="repository-layout"></a>

## 📁 Repository layout

```text
.github/            workflows, issue and pull-request templates, label catalogue
                    and its scripts, Excel evidence policy
docs/               contracts, architecture, governance, evidence record
  index.html        public project landing page (GitHub Pages); WEBSITE.md is its runbook
  assets/           images and diagrams of the documentation, with their sources
  methodology/      regulatory basis, support matrix, traceability, test-case format
    sources/        regulatory source documents
  user-manual/      user manual source, one folder per version (v1.0.0/)
src/                production VBA source: the only input to the workbook
  core/             calculation engine (internal)
  modules/          public worksheet functions
  workbook/         template, M_Main and the document modules
tests/              verification source and test data
  modules/          harness, case runner, suites and Test_RunAll
  fixtures/         synthetic case inputs
  expected/         expected results with their reference class
tools/              static checks, case generator, VBA synchronizer and
                    exporter, Excel evidence validator
```

The root holds the project documents (`README.md`, `INSTALLATION.md`,
`CONTRIBUTING.md`, `RELEASING.md`, `SECURITY.md`, `CODE_OF_CONDUCT.md`,
`CHANGELOG.md`, `LICENSE`) and the repository configuration (`.editorconfig`,
`.gitattributes`, `.gitignore`, `pyproject.toml`). `src/classes/`,
`src/forms/` and `examples/` are reserved locations: none exists yet, and each
is created with its first file.

| Directory | Owns | Must not own |
| --- | --- | --- |
| `src/` | Production components that go into the workbook, and the workbook template | Tests, examples, built workbooks |
| `tests/` | Test modules, synthetic fixtures, reviewed expected values | Production entry points, run output |
| `examples/` | Examples that use only the public API; created with the first example | Tests, real data |
| `docs/` | Contracts, architecture, [methodology](methodology/README.md) | Copies of root documents, Wiki pages |
| `tools/` | Deterministic checks, build and evidence scripts | Calculation logic |

`src/`, `tests/` and `tools/` each have a `README.md`, as do `docs/assets/`,
`docs/methodology/`, `docs/methodology/sources/` and each version folder of
`docs/user-manual/`.
Subdirectories are created with their first real file, never empty.

### Production source

| Location | Holds | Visibility |
| --- | --- | --- |
| `src/core/` | Calculation engine: configuration, validation, parsing, aggregation, reconciliation, evidence encoding, shared helpers | In-project only. Every module declares `Option Private Module`. |
| `src/modules/` | Public facade: the supported entry points, constants and enums | Supported API, listed in [`PUBLIC_API.txt`](PUBLIC_API.txt) |
| `src/classes/` | Class modules, e.g. trade or netting-set objects, state managers; reserved, none yet | Status stated in each class header |
| `src/workbook/` | Exported document modules: `ThisWorkbook` and sheet modules, and the macro-free template `SACCR_Template.xlsx`; also `M_Main` (a [known deviation](#known-deviations)) | Host glue only; no calculation logic |
| `src/forms/` | UserForms, `.frm` beside its `.frx`, only if a form is ever needed; reserved, none yet | — |

### Wiki

The GitHub Wiki is maintained directly in its own repository, in two
sections, Regulation and Workbook. It links to files and images of this
repository only on `main`, and it is not part of this repository's source or
checks. Where it differs from the versioned documents here, the versioned
documents govern.

### Verification and examples

| Location | Holds |
| --- | --- |
| `tests/modules/` | The harness, the case runner, the suites and `TEST_Suite` |
| `tests/fixtures/` | Synthetic inputs: trades, netting sets, collateral terms |
| `tests/expected/` | Expected values per regime with their provenance ([test cases](methodology/TEST_CASES.md)) |
| `examples/modules/` | Example modules calling the public facade; created with the first example |

`tools/check_source.py` enforces the placement rules below on every run of
`python tools/check.py`.

## 🔀 Dependency direction

```text
src/workbook  ──▶  src/modules (facade)  ──▶  src/core
                         │                       ▲
src/classes  ◀───────────┘                       │
tests/       ────────────────────────────────────┘ (may test core directly)
examples/    ──▶  src/modules only
```

- `src/core` depends on nothing outside `src/core`. No Excel object model, no
  workbook, sheet or range access, no UI.
- `src/modules` validates inputs at the public boundary and delegates to the
  core. It holds no formulas of its own.
- `src/workbook` connects sheets and events to the facade. It does not compute.
- `tests` may call core procedures directly; `examples` may not.

<a id="public-api-boundary"></a>

## 🧩 Public API boundary

A VBA `Public` declaration is not automatically supported API.

- **Supported:** declarations in `src/modules/` listed in
  [`PUBLIC_API.txt`](PUBLIC_API.txt). Changing one is a contract change under
  [`CONTRIBUTING.md`](../CONTRIBUTING.md).
- **In-project only:** everything in `src/core/`, kept off the external surface
  by `Option Private Module`, and anything in `src/modules/` not listed in the
  manifest.
- **Never supported:** test and example modules. They are not part of the
  workbook.

The manifest lists the ten `SACCR_*` worksheet functions in `SACCR_Formulas`.

<a id="known-deviations"></a>

### Known deviations: imported prototype engine

The engine imported from the prototype workbook `SACCR_Calculator.xlsm` keeps
its original structure until it is refactored:

| Rule | Deviation |
| --- | --- |
| `src/core` never touches Excel | `CORE_Engine` and `CORE_Util` read the input sheets and write the output sheets |
| `src/workbook` holds document modules | `M_Main`, a standard module, holds the sheet-button macros `RunSACCR`, `ValidateInputs`, `ClearOutputs` and `ExportEvidence`, and reads and writes the reconciliation sheets |
| Naming rules in `VBA_HOUSE_STYLE.md` | `M_Main` keeps the prototype's `M_` prefix; it has no role prefix of its own until it moves out of the standard modules |
| Dependency direction | `CORE_Engine` calls `SACCR_*` worksheet functions in `src/modules` |
| Facade holds no formulas | `SACCR_Formulas` implements the SA-CCR formulas itself |
| `src/workbook` calls only the facade | `M_Main` calls `CORE_Engine` and `CORE_Reconcile` directly |
| No module state in `src/core` | `CORE_Engine` keeps run state in module variables |

Each deviation is removed by a reviewed change, not by reformatting. Their
removal is planned in #220 (v1.3.0).

## 🚦 Placement rules

1. VBA components (`.bas`, `.cls`, `.frm`, `.frx`) live only in
   `src/core/`, `src/modules/`, `src/classes/`, `src/workbook/`, `src/forms/`,
   `tests/` or `examples/`.
2. Every standard module in `src/core/` declares `Option Private Module`.
3. Test and example modules never go into the production import set.
4. Fixtures are synthetic; never commit real trades, counterparties or
   portfolios.
5. Workbooks are built from source and never committed, except at an exact
   path explicitly re-included in `.gitignore`. The one such path is
   `src/workbook/SACCR_Template.xlsx`, the macro-free template that holds the
   sheets, formulas, named ranges and buttons; it contains no VBA project and
   no document properties (`docProps/`).
6. A new location for VBA components, fixtures or the workbook becomes
   contractual only when this document, the directory README,
   `INSTALLATION.md` and the checks are updated together. Documentation
   folders such as `docs/assets/` and `docs/user-manual/` need only this
   document and their own README.

Rules 1 and 2, and the content of a committed workbook under rule 5, are
enforced by `tools/check_source.py`; the others by review.

## 📚 Directory guides

- [`../src/README.md`](../src/README.md)
- [`../tests/README.md`](../tests/README.md)
- [`assets/README.md`](assets/README.md)
- [`methodology/README.md`](methodology/README.md)
- [`user-manual/v1.0.0/README.md`](user-manual/v1.0.0/README.md)
- [`../tools/README.md`](../tools/README.md)

---

**Structure principle:** one home per responsibility, the engine never touches
Excel, and only the facade is supported API.
