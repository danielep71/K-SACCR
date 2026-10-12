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

**SACCR uses the `application` profile.** Decided by the owner on 2026-10-06
in issue #3.

| Question | Answer for SACCR |
| --- | --- |
| Who calls it? | A user working in the SACCR workbook: entering trades and netting sets, running the calculation, reading results. |
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
src/        production VBA source: the only input to the workbook
tests/      regression modules, synthetic fixtures and expected results
examples/   runnable examples of the supported API
docs/       contracts, architecture and methodology
tools/      static checks and evidence tooling; later, build tooling
.github/    workflows, label catalogue and scripts
```

| Directory | Owns | Must not own |
| --- | --- | --- |
| `src/` | Production components that go into the workbook, and the workbook template | Tests, examples, built workbooks |
| `tests/` | Test modules, synthetic fixtures, reviewed expected values | Production entry points, run output |
| `examples/` | Examples that use only the public API | Tests, real data |
| `docs/` | Contracts, architecture, [methodology](methodology/README.md) | Copies of root documents, Wiki pages |
| `tools/` | Deterministic checks, build and evidence scripts | Calculation logic |

`src/`, `tests/`, `examples/` and `tools/` each have a `README.md`, as do
`docs/methodology/` and `docs/methodology/sources/`.
Subdirectories are created with their first real file, never empty.

### Production source

| Location | Holds | Visibility |
| --- | --- | --- |
| `src/core/` | Calculation engine: SA-CCR formulas, validation, parsing, shared helpers | In-project only. Every module declares `Option Private Module`. |
| `src/modules/` | Public facade: the supported entry points, constants and enums | Supported API, listed in [`PUBLIC_API.txt`](PUBLIC_API.txt) |
| `src/classes/` | Class modules, e.g. trade or netting-set objects, state managers | Status stated in each class header |
| `src/workbook/` | Exported document modules: `ThisWorkbook` and sheet modules, and the macro-free template `SACCR_Template.xlsx` | Host glue only; no calculation logic |
| `src/forms/` | UserForms, `.frm` beside its `.frx`, only if a form is ever needed | — |

### Wiki

The GitHub Wiki is maintained directly in its own repository and is
self-contained: it does not link to files of this repository, and it is not
part of this repository's source or checks. Where it differs from the
versioned documents here, the versioned documents govern.

### Verification and examples

| Location | Holds |
| --- | --- |
| `tests/modules/` | Regression modules and the harness entry point (#7) |
| `tests/fixtures/` | Synthetic inputs: trades, netting sets, collateral terms |
| `tests/expected/` | Expected values per regime with their provenance ([test cases](methodology/TEST_CASES.md)) |
| `examples/modules/` | Example modules calling the public facade |

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
| `src/workbook` holds document modules | `M_Main`, a standard module, holds the sheet-button macros `RunSACCR`, `ValidateInputs` and `ClearOutputs` |
| Naming rules in `VBA_HOUSE_STYLE.md` | `M_Main` keeps the prototype's `M_` prefix; it has no role prefix of its own until it moves out of the standard modules |
| Dependency direction | `CORE_Engine` calls `SACCR_*` worksheet functions in `src/modules` |
| Facade holds no formulas | `SACCR_Formulas` implements the SA-CCR formulas itself |
| `src/workbook` calls only the facade | `M_Main` calls `CORE_Engine` directly |
| No module state in `src/core` | `CORE_Engine` keeps run state in module variables |
| `On Error Resume Next` around single statements | `CORE_Engine.WriteRunInfo` runs under it as a whole |

Each deviation is removed by a reviewed change, not by reformatting.

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
6. A new location becomes contractual only when this document, the directory
   README, `INSTALLATION.md` and the checks are updated together.

Rules 1 and 2, and the content of a committed workbook under rule 5, are
enforced by `tools/check_source.py`; the others by review.

## 📚 Directory guides

- [`../src/README.md`](../src/README.md)
- [`../tests/README.md`](../tests/README.md)
- [`../examples/README.md`](../examples/README.md)
- [`methodology/README.md`](methodology/README.md)
- [`../tools/README.md`](../tools/README.md)

---

**Structure principle:** one home per responsibility, the engine never touches
Excel, and only the facade is supported API.
