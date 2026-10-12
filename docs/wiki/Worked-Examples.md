# 🧮 Worked Examples

> **Guide, not policy:** [Methodology](../methodology/README.md) defines scope; [expected files](../../tests/expected/) retain values, provenance, units, and tolerances.

## Published Basel examples

These are the registered published EAD targets, not newly executed engine outputs. The source prints amounts in USD thousands; repository fixtures use unscaled USD.

| CRE99 example | Published EAD, USD thousands | Fixture target, USD |
| --- | ---: | ---: |
| 1 — interest rate | 569 | 569,000 |
| 2 — credit | 381 | 381,000 |
| 3 — commodities | 5,406 | 5,406,000 |
| 4 — cross-asset | 936 | 936,000 |
| 5 — margined | 1,879 | 1,879,000 |

Use the fixture's exact tolerance and inputs. Example 3 uses 270 days with DaysPerYear = 360 to represent exactly 0.75 years. The old wiki's 274/365 scenario and its 5,404.62 result are not the current published-example fixture. T02 uses the printed delta -0.2694, not the prototype's -0.27.

To reproduce the complete examples, use the [versioned fixtures](../../tests/fixtures/) and [expected outputs](../../tests/expected/) with TEST_Cases. Do not reconstruct full-precision inputs from rounded display tables.

## Illustrative posted-VM cap scenario

The old wiki's “Basel 376.6m versus CRR 40.8m” conclusion is superseded by #39. Current expected files use **40,774,849.34860364 EUR in both regimes**, with the cap applied. This is a fixture expectation, not a fresh run result.

The scenario uses V = -1.1 billion, posted VM = -1.3 billion, and NICA = 0. The current cap calculation uses NICA in both regimes. See [BCBS expected file](../../tests/expected/illustrative-posted-vm-cap.bcbs.json) and [CRR expected file](../../tests/expected/illustrative-posted-vm-cap.crr.json) for the derivation and tolerances.

These values are **illustrative**. Neither a rejected Q&A nor an unreviewed external calculation makes them published or independently certified.

## Illustrative economic-trade example

The regression added by #156 has one economic trade with two risk-category legs:

| Leg | Category | Notional | MtM | MtM leg |
| --- | --- | ---: | ---: | --- |
| L1 | Equity single name | 10,000 | 30 | Y |
| L2 | FX | 10,000 | 0 | N |

Both use Economic trade ID E1. The test expects RC 30, equity add-on 3,200, FX add-on 400, multiplier 1, and EAD 5,082. It expects one economic trade and two legs. Repeating MtM 30 on L2 is a deliberate rejection case.

This illustrates the grouping contract; it does not independently validate product mapping, digital-option sizing, or regulatory eligibility.

## Other engineering examples

Negative-rate and negative-commodity option shifts, other risks, climatic mapping, and CQS mapping remain illustrative. Their current values and identifiers are in the versioned expected files and [Test Catalogue](Test-Catalogue.md). Read their classification before using them as a reference.
