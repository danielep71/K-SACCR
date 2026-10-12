# 🛡️ SA-CCR Benchmark Wiki

> **Guide, not policy:** [README](../../README.md) defines the project purpose; [Methodology](../methodology/README.md) defines the supported scope and validation boundary.

SA-CCR Benchmark is an Excel/VBA toolkit for calculating derivative counterparty exposure and examining how that exposure was built.

**Status as of October 12, 2026:** v1.0.0 is under development. The prototype's internal v1.1.0 label was not a published release. The reviewed baseline is integration commit `21978de6e783df697611be09290dc135d8edd065` (#166).

## 🧭 Start here

| Page | Purpose |
| --- | --- |
| [Getting Started](Getting-Started.md) | Build a development workbook, compile, run, and inspect results |
| [SA-CCR Overview](SA-CCR-Overview.md) | Understand RC, PFE, and EAD |
| [Regulatory Sources](Regulatory-Sources.md) | Find registered texts and understand their evidence status |
| [Regulatory Basis](Regulatory-Basis.md) | Read the project's regime and scope policy |
| [Basel vs CRR](Basel-vs-CRR.md) | Understand implemented regime choices and corrections |
| [How It Works](How-It-Works.md) | Follow the calculation and inspect the modules |
| [Workbook Guide](Workbook-Guide.md) | Use input tables, economic legs, and FormulaTest |
| [Worked Examples](Worked-Examples.md) | Distinguish published values from illustrative calculations |
| [Test Catalogue](Test-Catalogue.md) | Find all 28 catalogue mappings and the test suites |
| [Validation Approach](Validation-Approach.md) | Understand what a PASS does and does not establish |
| [Open Items and Limitations](Open-Items-and-Limitations.md) | See unfinished release work and later milestones |

## ✅ What is available

The engine supports the project's declared CRR and BCBS portfolios, exposes trade and aggregation intermediates, validates inputs, withholds results for invalid/incomplete netting sets, groups economic trades and legs, and identifies reference versus what-if parameter calibration. The template includes the editable **FormulaTest** worksheet.

## 🧪 What has been validated

Under the accepted v1.0.0 policy, only the published Basel CRE99 checks T01–T14 count as regulatory numerical validation. Other cases are illustrative regression evidence. **CRR is implemented but not numerically validated under that policy.** A REGULATORY parameter label does not mean regulatory approval or independent certification.

Bank-result reconciliation and the final release package remain unfinished. Use synthetic data only.

Every maintained page belongs in `docs/wiki/`; publish from a clean, committed source tree using the [publication procedure](../WIKI_PUBLICATION.md).
