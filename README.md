<div align="center">

# 🏦 SA-CCR Benchmark — Independent Calculation and Validation Toolkit

### Counterparty credit risk under SA-CCR, built in Excel/VBA

[![Excel VBA](https://img.shields.io/badge/Excel_VBA-Windows-217346?style=for-the-badge&logo=microsoft-excel&logoColor=white)](#status)
[![Status](https://img.shields.io/badge/Status-Pre--release-6e7781?style=for-the-badge)](#status)
[![Branch](https://img.shields.io/badge/Branch-release%2F1.0.0-6f42c1?style=for-the-badge)](https://github.com/danielep71/VBA-SACCR-Toolkit/tree/release/1.0.0)
[![License](https://img.shields.io/badge/License-MPL_2.0-2ea44f?style=for-the-badge)](LICENSE)

<br>

**Source-first · Synthetic data only · Static checks on every change · Excel-verified calculations**

</div>

---

## ✨ What this project is

**SA-CCR (Standardized Approach for Counterparty Credit Risk)** is the Basel
regulatory framework for determining the exposure at default (EAD) of derivative
transactions and long-settlement transactions. It is a standardized,
non-modelled approach that reflects current exposure, potential future exposure,
netting, collateral and margining arrangements.

**SA-CCR Benchmark** is an independent Excel/VBA toolkit designed to reproduce,
test and reconcile SA-CCR calculations under the Basel Framework and EU CRR.

See the [project Wiki](https://github.com/danielep71/VBA-SACCR-Toolkit/wiki) for the regulatory concepts, calculation
structure and validation approach.

<a id="status"></a>

## 🧭 Status

Repository setup, milestone **v0.0.1**, is complete; its accepted baseline is
recorded in [issue #13](https://github.com/danielep71/VBA-SACCR-Toolkit/issues/13). The
prototype engine (CRR and Basel CRE52) has been imported as source; it is not
yet validated under the repository's test-case policy. SACCR uses the **application** profile: the
deliverable is a workbook built from the exported source, around a
host-independent calculation core
([repository structure](docs/REPOSITORY_STRUCTURE.md)). Setup established:

- repository conventions, static checks and CI;
- governance, issue metadata and labels;
- the methodology basis, CRR with Basel CRE52 also supported, and the numerical
  test-case format ([`docs/methodology/`](docs/methodology/README.md)); and
- Excel evidence bound to a commit ([`docs/EXCEL_EVIDENCE.md`](docs/EXCEL_EVIDENCE.md))
  and the first real Excel smoke test (#9).

Nothing has been released. See the [changelog](CHANGELOG.md) for what has been
added so far.

## 🌿 Development

Development takes place on the active release branch, `release/1.0.0`. Every
change goes through a task branch and a pull request into that branch. Changes
to `main` require an explicit owner instruction.

## 🚀 Getting started

```sh
git clone https://github.com/danielep71/VBA-SACCR-Toolkit.git
cd VBA-SACCR-Toolkit
git switch release/1.0.0
python tools/check.py
```

Requirements, the import procedure and troubleshooting are in
[`INSTALLATION.md`](INSTALLATION.md).

## 📚 Documentation

| Document | Covers |
| --- | --- |
| [Wiki](https://github.com/danielep71/VBA-SACCR-Toolkit/wiki) | SA-CCR overview, regulatory sources, worked examples and workbook guide, maintained in the Wiki repository |
| [`INSTALLATION.md`](INSTALLATION.md) | Developer setup, checks, Excel import and validation record |
| [`CONTRIBUTING.md`](CONTRIBUTING.md) | Workflow, commit messages, evidence and pull requests |
| [`docs/REPOSITORY_STRUCTURE.md`](docs/REPOSITORY_STRUCTURE.md) | Profile decision, source layout, public API boundary |
| [`docs/VBA_HOUSE_STYLE.md`](docs/VBA_HOUSE_STYLE.md) | VBA naming, contracts, errors, Excel-state cleanup, export rules |
| [`docs/EXCEL_EVIDENCE.md`](docs/EXCEL_EVIDENCE.md) | Manual Excel run record, procedure and validator |
| [`docs/methodology/`](docs/methodology/README.md) | Regulatory basis, source register, numerical test-case format |
| [`docs/GOVERNANCE.md`](docs/GOVERNANCE.md) | Branches, issue metadata, review checklist, GitHub controls |
| [`tools/README.md`](tools/README.md) | Static checks, CI, evidence validator and traffic export |
| [`docs/LABELS.md`](docs/LABELS.md) | Issue-label catalogue and its workflows |
| [`RELEASING.md`](RELEASING.md) | Integration into `main` and the release sequence |
| [`SECURITY.md`](SECURITY.md) | Private vulnerability reporting and data handling |
| [`CODE_OF_CONDUCT.md`](CODE_OF_CONDUCT.md) | Participant behavior and conduct reporting |
| [`CHANGELOG.md`](CHANGELOG.md) | Change history |

## 🔐 Data and security

Use synthetic data only. Never commit or attach real trades, netting sets,
counterparty or collateral data, or any client or employer material. Report
suspected vulnerabilities privately as described in [`SECURITY.md`](SECURITY.md).

## 👤 Author

**Daniele Penza** — [@danielep71](https://github.com/danielep71)

## 📄 License

[Mozilla Public License 2.0](LICENSE)
