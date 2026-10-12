# 🧪 Validation Approach

> **Guide, not policy:** [Tests](../../tests/README.md) and [test-case methodology](../methodology/TEST_CASES.md) define evidence requirements.

## Evidence is not a single PASS badge

| Evidence | What it supports | What it does not establish |
| --- | --- | --- |
| Static checks and CI | Source, fixture, tooling, and publication contracts | Excel compilation or numerical execution |
| Compile and smoke tests | Operation on a recorded Excel host | Complete methodology correctness |
| Published CRE99 checks | The specific Basel quantities and tolerances represented | All Basel rules or CRR numerical validation |
| Illustrative cases and invariants | Regression behavior and useful consistency checks | Independent numerical certification |
| FormulaTest | Excel-formula versus VBA agreement for selected arguments | Independent review of the underlying interpretation |
| Independent comparator | Future cross-check evidence when reviewed | Regulatory authority by itself |

## Accepted v1.0.0 policy

Decision 8 limits numerical validation to published CRE99 values. T01–T14 are published Basel checks; T15–T28 are illustrative. The project does not claim independently reviewed expected values for v1.0.0. **CRR is implemented but not numerically validated under this policy.**

Run classification is separate: REGULATORY means matching parameter calibration, not complete input validity or certification. Check rejected/withheld netting sets independently.

## Reproduce a run

Record the full source SHA, template used, imports, Excel build and bitness, suite names and counts, failures, cleanup/restoration outcome, and skipped scenarios. A partial run or failed cleanup is not a pass. Use [Excel Evidence](../EXCEL_EVIDENCE.md) for the formal contract.

Historical evidence includes #156's InputValidation 88 cases / 119 checks, #166's owner confirmation of the final FormulaTest template, and #168's InputValidation 97 cases / 129 checks. These are recorded development results, not fresh tests performed for this wiki review or final release certification.

## Reconcile progressively

Compare population and units first, then EAD, RC/PFE, multiplier, aggregate add-on, asset class, hedging set/bucket, and trade inputs/intermediates. Compare economic-trade counts with economic-trade counts, not leg counts.

The reproducible bank-comparison/export workflow is still [#45](https://github.com/danielep71/VBA-SACCR-Toolkit/issues/45). Independent second-source validation is [#100](https://github.com/danielep71/VBA-SACCR-Toolkit/issues/100); final package and candidate evidence are [#46](https://github.com/danielep71/VBA-SACCR-Toolkit/issues/46).
