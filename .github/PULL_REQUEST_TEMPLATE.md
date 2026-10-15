<!--
  Target: the active release branch (release/1.0.0). Never main, unless the
  owner asked for an integration.
  One coherent outcome per pull request. Complete every common section; delete
  the optional review blocks that do not apply.
  Write NOT RUN or NOT APPLICABLE with a reason; never manufacture PASS
  evidence. Static checks are never Excel evidence.
  Synthetic data only. Report vulnerabilities privately through SECURITY.md.
-->

<div align="center">

# 🔀 K-SACCR Pull Request

### Exact evidence · Reviewable change · Honest boundaries

[![Contributing](https://img.shields.io/badge/guide-CONTRIBUTING-217346?style=flat-square)](https://github.com/danielep71/VBA-SACCR-Toolkit/blob/main/CONTRIBUTING.md)
[![Excel evidence](https://img.shields.io/badge/evidence-EXCEL__EVIDENCE-1D76DB?style=flat-square)](https://github.com/danielep71/VBA-SACCR-Toolkit/blob/main/docs/EXCEL_EVIDENCE.md)
[![Security](https://img.shields.io/badge/security-private%20reporting-d73a49?style=flat-square)](https://github.com/danielep71/VBA-SACCR-Toolkit/blob/main/SECURITY.md)
[![Changelog](https://img.shields.io/badge/changes-Unreleased-d97706?style=flat-square)](https://github.com/danielep71/VBA-SACCR-Toolkit/blob/main/CHANGELOG.md)

</div>

---

## 📌 Summary

<!-- The observable outcome and why it is needed, in two or three sentences. -->

## 🔗 Related issues

- Closes #
- Related to #

Use a closing keyword only when this pull request meets the issue's complete acceptance criteria.

## 🧭 Change classification

- [ ] 🐛 Defect correction
- [ ] ✨ Backward-compatible capability
- [ ] 💥 Breaking API, result or layout change
- [ ] 🧮 Calculation or methodology change
- [ ] ♻️ Internal refactor, no intended behavior change
- [ ] 🧪 Test, fixture or expected-value change
- [ ] 🔐 Security or trust-boundary hardening
- [ ] 📖 Documentation only
- [ ] 🛠️ Tooling, workflow or governance
- [ ] 📦 Release preparation

## 🎚️ Affected surface

- [ ] Calculation engine (`src/core`)
- [ ] Public API (`src/modules`, `docs/PUBLIC_API.txt`)
- [ ] Workbook or user interface (`src/workbook`, the template `src/workbook/SACCR_Template.xlsx`, `src/forms`)
- [ ] Tests, fixtures or expected values (`tests/`)
- [ ] Tooling or CI (`tools/`, `.github/`)
- [ ] No runtime surface: documentation or repository only

---

## 📐 Scope and contract impact

**In scope:**

- <!-- Deliberate outcome -->

**Out of scope:**

- <!-- Adjacent work deliberately deferred -->

```text
Supported behavior changed:        Yes / No
Backward compatible:               Yes / No / Uncertain
Suggested release impact:          none / patch / minor / major
Public declarations added/removed:
Changed results, errors or state:
Regime affected:                   CRR / BCBS / both / none
Migration required:
Known limitation introduced or kept:
```

- [ ] Components and import order are unchanged.
- [ ] Components or import order changed, and `INSTALLATION.md` is updated.
- [ ] No production source impact.

## 🔧 Implementation notes

```text
Approach and key invariant:
Alternatives considered:
New reference, dependency or generated input:
Excel state ownership and cleanup:
Failure behavior:
```

<!-- Explain what a reviewer cannot safely infer from the diff. For a GitHub
Actions or quality-tool update, follow tools/README.md → Dependency updates and
record old and new pins, release notes reviewed and rollback. -->

---

## ✅ Verification

### Candidate identity

| Evidence | Result |
| --- | --- |
| PR head SHA | <!-- full 40-character SHA --> |
| Base branch and SHA | <!-- release/1.0.0 + full SHA --> |
| Local working tree | <!-- clean / dirty: explain --> |

Evidence from another commit does not certify this one.

### Static checks

| Check | Result |
| --- | --- |
| `python tools/check.py --ci --base <base-sha>` | <!-- PASS / FAIL --> |
| **Repository integrity** (CI) | <!-- PASS / FAIL + run link --> |
| Ruff · mypy · actionlint (in CI) | <!-- PASS / FAIL / N/A --> |

### Excel and VBA execution

- [ ] Required and completed against the PR head.
- [ ] Required but not complete: reason and merge consequence stated.
- [ ] Not required: no VBA, workbook or packaging change.

| Evidence | Result |
| --- | --- |
| Tested commit | <!-- full SHA or N/A --> |
| Debug → Compile VBAProject | <!-- PASS / FAIL / NOT RUN / N/A --> |
| `TEST_Harness.RunTests` | <!-- RESULT= line, verbatim --> |
| Cases / assertions / failures / cleanup | <!-- e.g. 4 / 6 / 0 / PASS --> |
| Other suites (`RunMainStateTests`, `RunCaseTests`, `RunInputValidationTests`, `RunInvariantTests`, `RunAggregationTests`) | <!-- RESULT= lines, or N/A --> |
| Specific scenario | <!-- what was exercised and the outcome --> |
| Evidence bundle | <!-- validator result per docs/EXCEL_EVIDENCE.md, or N/A --> |

```text
Excel version and build:
Office bitness:                    32-bit / 64-bit
Windows version:
Locale / decimal separator:
```

Record only environments actually used. One bitness never proves the other.

### Regression coverage

- [ ] Existing cases cover the changed path.
- [ ] New or amended cases cover each corrected defect.
- [ ] Boundary, invalid-input and cleanup paths are covered as applicable.
- [ ] Harness counts and `.github/excel-evidence-policy.json` are updated together.
- [ ] Every expected value carries a reference class per `docs/methodology/TEST_CASES.md`; only `published` and `independent` values count as validation, and none is produced by running this code.
- [ ] No regression change needed: rationale below.

```text
New or changed cases:
Coverage deliberately deferred:
```

---

## ⚠️ Risk and rollback

- [ ] 🟢 Low: documentation, metadata or mechanically verified change.
- [ ] 🟡 Medium: bounded runtime, tooling or compatibility impact.
- [ ] 🔴 High: numerical results, shared Excel state, security, release or breaking impact.

```text
Principal failure modes:
Residual risk:
Rollback procedure:
```

## 🔐 Security, data and provenance

- [ ] No credential, token, internal URL or personal path.
- [ ] No real trades, netting sets, counterparties, collateral or client data; fixtures are synthetic.
- [ ] Regulatory texts, examples and external code are cited with source and version.
- [ ] No vulnerability detail that belongs in private disclosure.

## 📚 Documentation and release hygiene

- [ ] `CHANGELOG.md` records the change under `[Unreleased]`, or it is not material.
- [ ] `README.md`, `INSTALLATION.md` and `CONTRIBUTING.md` match the new behavior.
- [ ] `docs/PUBLIC_API.txt` and `docs/methodology/` are updated where affected.
- [ ] No version, tag or `VERSION` change, unless this is the approved release PR.
- [ ] No documentation change needed: reason below.

---

## 🧩 Specialist review

<details>
<summary><strong>🧮 SA-CCR calculation or methodology</strong></summary>

Keep only for changes to formulas, parameters, mappings or expected values.

- [ ] Each rule cites its source ID and locator, such as `CRR Article 275(1)`, and appears in the traceability table.
- [ ] The regime (CRR, BCBS or both) is explicit, and any CRR/BCBS difference is recorded first.
- [ ] Units, signs, dates, maturity conventions and valid input domains are explicit.
- [ ] Expected values carry a reference class, and tolerances follow `docs/methodology/TEST_CASES.md`.
- [ ] Boundary cases are covered: zero, floors, caps, offsetting trades, single trade.

</details>
<details>
<summary><strong>🪟 Workbook, UI or Excel state</strong></summary>

Keep only for workbook modules, forms, events or any change to Excel settings.

- [ ] Caller-owned Excel state is captured before change and restored on success and failure.
- [ ] Every event, form or registration has an owner and a teardown rule.
- [ ] 32-bit and 64-bit declarations are assessed.

</details>
<details>
<summary><strong>📦 Release preparation</strong></summary>

Keep only for the release PR. Follow `RELEASING.md`.

- [ ] `CHANGELOG.md` is dated and `VERSION` matches it.
- [ ] Excel certification passed on the exact candidate, with a validated evidence bundle.
- [ ] Any artifact is built from the candidate, reopened, tested and hashed.

</details>

---

## 👀 Reviewer focus

```text
Highest-risk decision:
Files and procedures to inspect first:
Evidence to challenge:
Known boundary this PR does not prove:
```

## ☑️ Final author check

- [ ] The title states the observable outcome.
- [ ] One coherent purpose, no unrelated churn.
- [ ] Linked acceptance criteria are met, or the remaining work is explicit.
- [ ] Evidence belongs to the exact head SHA.
- [ ] VBA changes were compiled and tested in Excel, or are explicitly not yet verified.
- [ ] The whole diff, including comments and documentation, was reviewed.
- [ ] No placeholder, unexplained N/A or private material remains.

---

**Review principle:** approve the smallest coherent change whose contract, evidence, risk and rollback can all be explained from this pull request.
