# ⚖️ Regulatory Basis

> **Guide, not policy:** [Methodology](../methodology/README.md) owns source versions, scope decisions, and regime differences.

The project implements two explicitly identified regimes:

| Regime | Project role |
| --- | --- |
| CRR | EU implementation and default project baseline |
| BCBS | Basel CRE52 implementation, with differences handled explicitly |

Results and test expectations must identify their regime. Do not relabel a Basel worked example as a published CRR result because the same numerical value happens to occur.

## First-release boundary

The declared scope is full SA-CCR for supported portfolios. It does not include the complete prudential capital calculation. Simplified SA-CCR, OEM, CCP exposure/default-fund treatment, SFTs, CVA, and RWA have separate future milestones; see [Open Items and Limitations](Open-Items-and-Limitations.md).

The support matrix distinguishes supported inputs, user-supplied classifications, and rejected structures. Margining and selected client-clearing MPOR inputs do not imply that the full CCP exposure/default-fund method is implemented.

## Parameters and validity are different controls

Reference set `CRE52-CRR-2026-10-09` is compared with parameter values, factor rows, and alpha overrides. Differences produce **WHAT-IF** and individual Checks entries. Matching calibration produces **REGULATORY**. The owner's policy treats `DaysPerYear` and `IRBucketOffset` as conventions for classification; regime-specific validity checks still apply.

REGULATORY describes calibration of accepted results. It does not certify an entire submitted portfolio, validate withheld results, establish legal eligibility, or provide regulatory approval.

## Validation boundary

Published CRE99 checks support the limited Basel validation claim described in the methodology. CRR behavior and additional engineering cases remain illustrative under the accepted v1.0.0 policy. Independent second-source comparison is tracked in [#100](https://github.com/danielep71/VBA-SACCR-Toolkit/issues/100).
