# 🧪 Tests

`tests/` holds all verification source and stable test data. Layout rules are in
[`docs/REPOSITORY_STRUCTURE.md`](../docs/REPOSITORY_STRUCTURE.md).

| Location | Contents |
| --- | --- |
| `modules/` | The harness `TEST_Harness.bas`, the shared case runner `TEST_CaseRunner.bas`, the generated `TEST_Cases.bas`, the `TEST_Formulas`, `TEST_MainState`, `TEST_InputValidation`, `TEST_Aggregation`, `TEST_Invariants` and `TEST_Reconcile` suites, `TEST_Suite.bas`, whose `Test_RunAll` runs them all, and the benchmark `TEST_Benchmark.bas` |
| `fixtures/` | Synthetic inputs: trades, netting sets, collateral terms ([`TEST_CASES.md`](../docs/methodology/TEST_CASES.md)) |
| `expected/` | Expected values per regime, each with its reference class and source ([`TEST_CASES.md`](../docs/methodology/TEST_CASES.md)) |
| `csv/` | Synthetic CSV input files of the [CSV contract](../docs/csv/CSV_CONTRACT.md): `valid/` files that must be accepted, `invalid/` files that must be rejected, and `expected.json` with the findings of each |
| `derivations/` | Separate Python implementations, sharing no code with the engine, that compute the expected values an `expected/` file cites in its `derivation`; run from the repository root |

`TEST_Suite.Test_RunAll` runs the harness and every suite and prints a
summary. In the order it runs them, the entry points of the single suites are
`TEST_Harness.RunTests`, `TEST_Formulas.RunFormulaTests`,
`TEST_MainState.RunMainStateTests`, `TEST_Cases.RunCaseTests`,
`TEST_InputValidation.RunInputValidationTests`,
`TEST_Aggregation.RunAggregationTests`, `TEST_Invariants.RunInvariantTests`
and `TEST_Reconcile.RunReconcileTests`; `TEST_Cases`, `TEST_InputValidation`,
`TEST_Aggregation` and `TEST_Invariants` need `TEST_CaseRunner`. How to run
them and what a passing log looks like is in
[`INSTALLATION.md`](../INSTALLATION.md#running-the-harness).

`TEST_Benchmark.RunBenchmark` is not a suite: it times the engine on large
synthetic portfolios (#291), its timings depend on the machine, and neither
`Test_RunAll` nor the Excel evidence includes it. It also needs
`TEST_CaseRunner`; see [the benchmark](../INSTALLATION.md#benchmark).

Rules:

- Test modules are never part of the production import set.
- Fixtures are synthetic. Never commit real trades, counterparties, collateral
  agreements or portfolio extracts.
- Expected values never come from running the implementation: they come from
  published regulatory examples, from an external implementation registered
  as a source, or, for `illustrative` cases, from a separate derivation.
  Cases named `illustrative-*` hold only illustrative values: they show a
  behaviour and validate nothing.
- Run output, logs and generated workbooks are not committed.
- `python tools/check.py` validates every fixture and expected file against
  the format (`check_test_cases`) and checks that the generated
  `modules/TEST_Cases.bas` is current. `TEST_Cases.RunCaseTests` runs them in
  Excel through `modules/TEST_CaseRunner.bas`.

<a id="ported-prototype-catalogue"></a>

## Ported prototype catalogue

The 28 checks of the prototype workbook's TestCatalogue sheet (#59), with the
catalogue's expected value. Inputs come from the template's NettingSets and
Trades sheets, valued on 2026-09-30. The `cre99-example-*` cases are in USD,
unscaled: CRE99 prints USD thousands, so their amounts and expected values are
the printed figures times 1,000, with a tolerance of half a printed unit
(500 USD). The other cases keep the template's figures, labelled EUR. Netting sets
that the workbook duplicates per regime with identical inputs (`BCBS-EX5` and
`CRR-EX5`, `BCBS-VMCAP` and `CRR-VMCAP`, `BCBS-NEGIR` and `CRR-NEGIR`) are one
fixture with an expected file per regime. T22 and T26 read trades `NIR-B` and
`OT-X` in the workbook; the fixtures hold the same inputs as `NIR-C` and
`OT-1`. T28 needs only trade `B03` of `NS-GAMMA`.

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

The 17 expected files with no catalogue ID are additional illustrative cases
that do not come from the catalogue; each records its own derivation:

- `illustrative-ir-sum-of-absolutes.bcbs`, `illustrative-ir-swap-unmargined.bcbs`,
  `illustrative-ir-swap-unmargined.crr` and `illustrative-partly-rejected.crr`:
  a hand calculation, a format example or the engine behaviour it checks;
- the five CCP files `illustrative-ccp-*`, the four client-clearing files
  `illustrative-clearing-*` and the four settled-to-market files
  `illustrative-settled-to-market-*`: the scripts in
  [`derivations/`](derivations/), which share no code with the engine.

`published` values are printed in the Basel Framework chapter CRE99 (and BCBS
279 Annex 4a), rounded to the printed digit, and were checked against the text
on 2026-10-09. T02 is printed as -0.2694; the prototype catalogue had -0.27,
and the workbook's TestCatalogue sheet now uses -0.2694 (#267). Example 3
prints a residual maturity of 9 months for trade 1, so its case runs with
`DaysPerYear` 360 and 270 days to give exactly 0.75 years. `illustrative`
values came from a Python implementation that is not in the repository, from
hand calculations or from the scripts in `derivations/`, and stay illustrative unless an external implementation
registered as a source reproduces them (#100). The values and tolerances of
T15 to T28 are the catalogue's, except T16, revised in #39 (#130) to the cap
with C = NICA only.
