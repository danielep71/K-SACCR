# 🧪 Test Catalogue

> **Guide, not policy:** [Tests](../../tests/README.md) owns the test inventory; machine-readable [expected files](../../tests/expected/) retain exact values and provenance.

## The 28 catalogue mappings

The table maps the historical catalogue IDs to versioned cases. It contains expectations, not a claim that the current workbook has just been executed. T16 is corrected to the current expected JSON after #39; the older tests README still carries its prototype number.

| Test | Case and regime | Quantity (trade) | Expected | Class |
| --- | --- | --- | --- | --- |
| T01 | `cre99-example-1.bcbs` | `adjusted_notional` (EX1-T1) | 78,694,000 USD | published |
| T02 | `cre99-example-1.bcbs` | `supervisory_delta` (EX1-T3) | -0.2694 | published |
| T03 | `cre99-example-1.bcbs` | `add_on.interest_rate` | 347,000 USD | published |
| T04 | `cre99-example-1.bcbs` | `exposure_value` | 569,000 USD | published |
| T05 | `cre99-example-2.bcbs` | `multiplier` | 0.965 | published |
| T06 | `cre99-example-2.bcbs` | `exposure_value` | 381,000 USD | published |
| T07 | `cre99-example-3.bcbs` | `exposure_value` | 5,406,000 USD | published |
| T08 | `cre99-example-4.bcbs` | `replacement_cost` | 40,000 USD | published |
| T09 | `cre99-example-4.bcbs` | `aggregate_add_on` | 629,000 USD | published |
| T10 | `cre99-example-4.bcbs` | `exposure_value` | 936,000 USD | published |
| T11 | `cre99-example-5.bcbs` | `margin_period_of_risk` | 14 business days | published |
| T12 | `cre99-example-5.bcbs` | `aggregate_add_on` | 1,401,000 USD | published |
| T13 | `cre99-example-5.bcbs` | `multiplier` | 0.958 | published |
| T14 | `cre99-example-5.bcbs` | `exposure_value` | 1,879,000 USD | published |
| T15 | `cre99-example-5.crr` | `exposure_value` | 1,879,000 USD | illustrative |
| T16 | `illustrative-posted-vm-cap.bcbs` | `exposure_value` | 40774849.34860364 EUR | illustrative |
| T17 | `illustrative-posted-vm-cap.crr` | `exposure_value` | 40774849.34860364 EUR | illustrative |
| T18 | `illustrative-posted-vm-cap.crr` | `cap_applied` | `Y` | illustrative |
| T19 | `illustrative-swaption-forward-1pct.crr` | `exposure_value` | 760.4939328393385 EUR | illustrative |
| T20 | `illustrative-negative-ir-forward.crr` | `lambda_shift` (NIR-C) | 0.003 | illustrative |
| T21 | `illustrative-negative-ir-forward.crr` | `supervisory_delta` (NIR-C) | -0.9941752720429311 | illustrative |
| T22 | `illustrative-negative-ir-forward.bcbs` | `trade_status` (NIR-C) | `EXCLUDED` | illustrative |
| T23 | `illustrative-negative-commodity-price.crr` | `lambda_shift` (NCO-1) | 5.5 price | illustrative |
| T24 | `illustrative-other-risks.crr` | `add_on.other` | 880 EUR | illustrative |
| T25 | `illustrative-other-risks.crr` | `exposure_value` | 1,232 EUR | illustrative |
| T26 | `illustrative-other-risks.bcbs` | `trade_status` (OT-1) | `EXCLUDED` | illustrative |
| T27 | `illustrative-climatic-commodity.crr` | `hedging_set` (CL-1) | `CO\|CLIMATIC\|STD` | illustrative |
| T28 | `illustrative-credit-quality-step.crr` | `supervisory_factor` (B03) | 0.0042 | illustrative |

## Units and evidence

Published source amounts use USD thousands; fixture amounts are unscaled USD. Quantity-specific tolerances remain in each expected file. Published delta and multiplier tolerances must not be inferred from the monetary rounding tolerance.

T01–T14 are published Basel checks. T15–T28 are illustrative. The extra regression suites cover more cases than these 28 catalogue IDs; those counts are not interchangeable.

## Running and extending tests

Start with TEST_Harness.RunTests as described in [Installation](../../INSTALLATION.md#running-the-harness), then run the applicable complete suites, including TEST_Cases.RunCaseTests, input validation, invariants, aggregation, and state/error tests.

FormulaTest has 22 editable worksheet checks with tolerance in E4. Passing those checks does not replace the engine suites.

Add stable synthetic inputs and expected records under tests/fixtures and tests/expected, supply regime, units, locator, evidence class, and tolerance, and regenerate the checked VBA cases. Never turn current engine output into a “published” expectation.

Known fixture/engine whole-day validation parity is tracked in [#161](https://github.com/danielep71/VBA-SACCR-Toolkit/issues/161). Randomized FormulaTest inputs remain [#165](https://github.com/danielep71/VBA-SACCR-Toolkit/issues/165).
