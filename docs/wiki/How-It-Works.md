# ⚙️ How It Works

> **Guide, not policy:** [Methodology](../methodology/README.md) owns calculation rules and [Repository Structure](../REPOSITORY_STRUCTURE.md) owns architecture.

## Calculation flow

1. Read named inputs and resolve netting-set regimes.
2. Validate inputs and economic-trade grouping.
3. Compute trade-level quantities, including adjusted notional, delta, and maturity factor.
4. Aggregate through risk drivers, hedging sets, and asset classes.
5. Calculate replacement cost, PFE, applicable cap, and final EAD.
6. Write intermediates, Checks, validity/freshness status, and parameter classification.

The broad structure is:

$$
EAD = \alpha (RC + PFE), \qquad PFE = multiplier \times AddOn_{aggregate}.
$$

The ordinary reference calibration uses alpha 1.4 and multiplier floor 5%. The precise collateral basis and permitted margin treatment come from the selected regime and support contract.

Economic MtM enters the netting-set value once. Risk attributes and notional remain on each accepted leg. Rows with unrelated IDs cannot be recognized automatically as the same economic position.

## Modules

| Component | Responsibility |
| --- | --- |
| CORE_Config | Names, layout constants, parameter codes, and reference calibration |
| CORE_Engine | Input loading, validation, grouping, calculation, aggregation, and output |
| CORE_Util | Conversions, keyed lookups, and worksheet helpers |
| SACCR_Formulas | Public SACCR_* worksheet functions |
| M_Main | RunSACCR, ValidateInputs, ClearOutputs, and workbook entry-point behavior |
| TEST_* modules | Development regression suites |

The imported engine still accesses worksheets. A fully separated calculation core is a goal tracked in [#60](https://github.com/danielep71/VBA-SACCR-Toolkit/issues/60), not a completed architecture claim.

## Inspect a result

Use Results for netting-set outcomes, TradeCalc for trade intermediates, HedgingSets and Buckets for aggregation, and Checks for errors and parameter deviations. Invalid or incomplete sets must not be treated as valid partial exposure estimates.

The freshness mechanism has a known relocated-Params gap ([#152](https://github.com/danielep71/VBA-SACCR-Toolkit/issues/152)); rerun after edits and do not rely on a “current” indicator after moving those input ranges until fixed.

For public-function argument and error-contract work, see [#42](https://github.com/danielep71/VBA-SACCR-Toolkit/issues/42). Worksheet data validation alone cannot protect every direct UDF call.
