# 🧪 Numerical Test Cases

This document defines how K-SACCR's numerical test cases are written: inputs,
units, expected results, their provenance, tolerances, case categories and when
the suite counts as complete. Sources and rule IDs come from the
[methodology](README.md).

> [!IMPORTANT]
> The prototype's 28 TestCatalogue checks are ported as cases (#59): 14 values
> printed in CRE99 (the BCBS 279 Annex 4a examples) are `published`; the other
> 14 are `illustrative`. They run in Excel through `TEST_Cases.RunCaseTests`
> ([consumption](#consumption)). Which rules they validate, and for which
> regime, is in the [traceability](README.md#traceability) table. The mapping
> is in [`tests/README.md`](../../tests/README.md#ported-prototype-catalogue).

## 📁 Files

A case is one input file and one expected-result file per regime:

| File | Holds |
| --- | --- |
| `tests/fixtures/<case-id>.json` | Synthetic inputs: netting set, trades, collateral |
| `tests/expected/<case-id>.<regime>.json` | Expected results for one regime (`crr` or `bcbs`), each with its provenance |

Separating them lets the same inputs carry a `CRR` and a `BCBS` result. Case IDs
are lowercase kebab-case and never reused. A case whose expected values are all
illustrative starts with `illustrative-`.

Files are UTF-8 JSON, stored with LF. All data is synthetic: never real trades,
counterparties, collateral agreements or portfolio extracts.

<a id="fixture-format"></a>

## 📥 Fixture format

```json
{
  "schema_version": 1,
  "kind": "saccr-fixture",
  "id": "illustrative-ir-swap-unmargined",
  "synthetic": true,
  "description": "What the case exercises, in one sentence",
  "category": "nominal",
  "calculation_currency": "EUR",
  "valuation_date": "2026-01-02",
  "parameters": {},
  "netting_set": {
    "id": "NS1",
    "margined": false,
    "clearing_role": "none",
    "remargin_period_business_days": null,
    "large_or_illiquid": false,
    "margin_disputes": false,
    "mpor_override_business_days": null,
    "variation_margin_net_amount": 0,
    "independent_collateral_net_amount": 0,
    "threshold_amount": 0,
    "minimum_transfer_amount": 0,
    "alpha_factor": null,
    "qccp": null,
    "ccp_retains_variation_margin": null
  },
  "trades": [
    {
      "id": "T1",
      "asset_class": "interest_rate",
      "sub_class": null,
      "risk_factor": "EUR",
      "instrument": "linear",
      "direction": "long",
      "option_type": null,
      "nature": "standard",
      "hedging_set_label": null,
      "notional_amount": 10000000,
      "market_value_amount": 100000,
      "start_date": "2026-01-02",
      "end_date": "2031-01-02",
      "maturity_date": "2031-01-02",
      "option_expiry_date": null,
      "underlying_price": null,
      "strike_price": null,
      "lambda_price": null,
      "attachment_rate": null,
      "detachment_rate": null,
      "settlement_type": "ordinary",
      "next_reset_date": null,
      "description": "Five-year swap, format example"
    }
  ]
}
```

| Rule | Detail |
| --- | --- |
| Envelope | `schema_version`, `kind`, `id` (equal to the file name), `synthetic: true`, `description` and `category` are required |
| Units | In `netting_set` and `trades`, a field's suffix states its unit: `_amount` in `calculation_currency`, unscaled; `_date` as ISO `YYYY-MM-DD`; `_years` in years; `_business_days` in business days; `_rate` and `_factor` as decimals; `underlying_price`, `strike_price` and `lambda_price` share the underlying's unit, a decimal rate for an interest-rate option and a price otherwise. A quantity there never lacks a unit suffix. Envelope fields such as `schema_version` are metadata, not quantities |
| Currencies | Every amount is in `calculation_currency`; amounts in other currencies are converted before they enter a fixture ([currency](README.md#assumptions-and-scope)). `risk_factor` names what the trade is exposed to: the interest-rate currency, the FX pair, the credit or equity reference, or the commodity |
| Signs | Market values are from the bank's side: positive is an asset. `variation_margin_net_amount` and `independent_collateral_net_amount` are held minus posted. `notional_amount` is unsigned; `direction` gives the sign |
| Parameters | A fixture runs with the default parameters of the workbook template's Params sheet (for example 365 days and 250 business days a year). `parameters` lists the Params values the case runs with instead, by code; `{}` for none. Accepted: `DaysPerYear` (a number from 360 to 366, the engine's range), so that a published maturity such as 0.75 years can be entered as dates, and `IRBucketOffset` (Boolean), so that a BCBS case can use the sum of absolute bucket values; a case that needs another parameter adds it to the format first |
| Fields | Every field below is present in every fixture; `null` means not applicable or the default. `tools/check_test_cases.py` rejects an unknown, missing or mistyped field |

| Netting-set field | Type | Meaning |
| --- | --- | --- |
| `id` | text | Netting-set ID |
| `margined`, `large_or_illiquid`, `margin_disputes` | Boolean | As on the NettingSets sheet |
| `clearing_role` | `none`, `clearing_member`, `client` or `ccp` | Client clearing role on the NettingSets sheet: N, CM, CLIENT or CCP (#112) |
| `remargin_period_business_days` | whole number from 0 to 10000, or `null` | Remargining period of a margined set; `null` leaves the cell blank, so the engine uses 1, and the engine raises 0 to 1 with a warning |
| `mpor_override_business_days` | whole number above 0, or `null` | MPOR override, with no upper bound; `null` for none, never 0 |
| `variation_margin_net_amount`, `independent_collateral_net_amount` | amount | Net VM and NICA, held minus posted |
| `threshold_amount`, `minimum_transfer_amount` | amount, at least 0 | Margin agreement terms |
| `alpha_factor` | number above 0, or `null` | Alpha override; `null` uses the regime's alpha |
| `qccp`, `ccp_retains_variation_margin` | Boolean, or `null` | NettingSets *QCCP* and *CCP retains VM, collateral not protected* (columns R and S): `true` writes Y, `false` N, `null` leaves the cell blank. The engine requires both for `ccp` and neither for another role (#112) |

| Trade field | Type | Meaning |
| --- | --- | --- |
| `id` | text | Trade ID, unique in the fixture |
| `asset_class` | `interest_rate`, `foreign_exchange`, `credit`, `equity`, `commodity`, `other` | Asset class; `other` exists only under CRR |
| `sub_class` | text or `null` | Rating, credit quality step, index or commodity type, as in the supervisory-factor table |
| `risk_factor` | text | See Currencies above |
| `instrument` | `linear`, `option`, `cdo` | Instrument type |
| `direction` | `long`, `short` | Position |
| `option_type` | `call`, `put` or `null` | Options only |
| `nature` | `standard`, `basis`, `volatility` | Transaction nature |
| `hedging_set_label` | text or `null` | Basis and volatility transactions only |
| `notional_amount` | amount, at least 0 | Notional |
| `market_value_amount` | amount | Market value |
| `start_date`, `end_date`, `maturity_date`, `option_expiry_date` | date or `null` | As on the Trades sheet; `null` when not given |
| `underlying_price`, `strike_price`, `lambda_price` | number or `null` | Options; `lambda_price` is an entered shift |
| `attachment_rate`, `detachment_rate` | decimal or `null` | CDO tranches |
| `settlement_type` | `ordinary` or `settled_to_market` | Trades *Settlement type*: `ordinary` leaves it blank, `settled_to_market` writes `STM` (#360) |
| `next_reset_date` | date or `null` | Trades *Next reset date*; `null` when not given, as for an ordinary trade |
| `description` | text or `null` | Free text; never read by the engine |

<a id="expected-format"></a>

## 📤 Expected-result format

```json
{
  "schema_version": 1,
  "kind": "saccr-expected",
  "fixture": "illustrative-ir-swap-unmargined",
  "regime": "CRR",
  "outputs": [
    {
      "quantity": "replacement_cost",
      "rule": "CRR.275.1",
      "value": 100000,
      "unit": "EUR",
      "tolerance": {"absolute": 0.000001, "relative": 0.000000001},
      "reference": {
        "class": "illustrative",
        "source": "CRR",
        "locator": "Article 275(1)",
        "derivation": "How the value was obtained",
        "derived_by": "Name",
        "reviewed_by": null,
        "date": "2026-10-06"
      }
    }
  ]
}
```

An output is a number (`value`, `unit`, `tolerance`) or a text result
(`text`, compared exactly, no unit or tolerance). It may also carry
`catalogue`, the prototype TestCatalogue test it was ported from, such as
`T01`. An invalid input is expected as a rejection status, `trade_status` or
`netting_set_status`.

A fixture holds inputs the engine accepts, unless it tests a rejection. For
the inputs whose fields depend on each other, `tools/check_test_cases.py`
applies the engine's rules (#401):

- **Settlement (#360):** a `next_reset_date` only with `settled_to_market`,
  which needs one, after `valuation_date` and not after M (`maturity_date`,
  else `end_date`, else `option_expiry_date`); a reset on M is accepted. A
  trade that breaks a rule is rejected by the engine.
- **Netting set:** a `settled_to_market` trade in a margined netting set;
  `qccp` and `ccp_retains_variation_margin` required for `clearing_role`
  `ccp` and `null` for any other role; `ccp_retains_variation_margin` true
  only on a margined netting set (#112). A netting set that breaks a rule is
  `INVALID`.

Such a fixture is accepted only when every one of its expected files
declares the outcome: `trade_status` `EXCLUDED` for the rejected trade, or a
`netting_set_status` starting with `INVALID`. Otherwise it is a malformed
fixture, whose figures would not describe a calculation, and the check fails.

A third form, `error`, is reserved and not supported. It would replace
`value`, `unit` and `tolerance` with the name of a public `SACCR_ERROR_*`
constant. `tools/check_test_cases.py` accepts that form, but no such constant
exists, no expected file uses it, and `tools/generate_case_tests.py` refuses it
because the engine runner cannot check it.

Quantities use stable names, each with one meaning only:

| Level | Quantity | Form and unit |
| --- | --- | --- |
| Netting set | `exposure_value`, `replacement_cost`, `potential_future_exposure`, `aggregate_add_on`, `add_on.<asset_class>` | Number in the calculation currency |
| Netting set | `multiplier` | Number, unit `1` |
| Netting set | `margin_period_of_risk` | Number, unit `business_days` |
| Netting set | `cap_applied` | Text: `Y` when the unmargined cap sets the exposure value |
| Netting set | `netting_set_status` | Text: `VALID`; `INCOMPLETE: <r> of <n> trade(s) rejected`, `INVALID: <e> input error(s)` or, for a netting set declared as a multi-level client structure, `INVALID: multi-level client structure, not supported`, or, for a settled-to-market trade in a margined netting set, `INVALID: STM requires unmargined treatment; mixed margined/unmargined netting sets are not supported` (followed by `; <e> other input error(s)` when there are), when the exposure value is withheld; or `NO TRADES` |
| Trade | `adjusted_notional` | Number in the calculation currency |
| Trade | `supervisory_delta`, `supervisory_factor` | Number, unit `1` |
| Trade | `lambda_shift` | Number in the underlying's unit: `1` for a rate, `price` otherwise |
| Trade | `remaining_maturity` | Number, unit `years`: M as the maturity factor uses it, the time to the next reset date for a settled-to-market trade (TradeCalc column M) |
| Trade | `maturity_factor` | Number, unit `1`: the unmargined maturity factor (TradeCalc column MF unmargined); only for a netting set that is not margined |
| Trade | `maturity_factor_margined` | Number, unit `1`: the margined maturity factor (TradeCalc column MF margined), the netting set's or, where a CCP retains VM, the trade's own (#112); only for a margined netting set |
| Trade | `hedging_set` | Text: the engine's hedging-set key, for example `CO\|CLIMATIC\|STD` |
| Trade | `bucket` | Text: the bucket within the hedging set: the maturity bucket `<1Y`, `1-5Y` or `>5Y` of an interest-rate trade, the reference or commodity type otherwise |
| Trade | `trade_status` | Text: `OK` or `EXCLUDED` |

A trade-level output names its trade in `trade`; a netting-set output has no
`trade`. New quantities are added to this table and to the validator together.

Rule IDs follow the [traceability](README.md#traceability) convention:
`CRR.<article>[.<paragraph>]`, `CRR-RTS.<article>` for Delegated Regulation
(EU) 2021/931, `CRE52.<paragraph>`, with `CRE52` alone for a paragraph not
yet located, and `CRE54.<paragraph>` for the client-clearing rules of CRE54. The rule IDs of the ported cases are those the prototype cites and
are verified with the sources.

<a id="reference-classes"></a>

## 🏷️ Reference classes

Every expected value carries exactly one class. Only the first two qualify as
numerical-validation references under this policy. Regression, boundary and
rejection tests are evidence too, that the software behaves as specified, but
not that a figure is correct.

| Class | Meaning | Required fields |
| --- | --- | --- |
| `published` | Printed in a source from the register, for example a worked example | `source`, exact `locator` (section, example, table) |
| `independent` | Computed by an external implementation by another author, registered as a source, from the same inputs, for example the R package SACCR (#352) | `source` (a register ID starting `IMPL-`), `locator`, `derivation`, `derived_by` (the implementation and its version); `reviewed_by` may be `null` |
| `illustrative` | Shows the format or a behaviour; not a validated SA-CCR figure | `derivation` saying why it exists |

An `independent` value is never produced by running K-SACCR itself, and its
`source` is an external implementation, registered with an ID starting
`IMPL-`; a `published` value never cites one. The checker rejects either
mismatch, so a value derived within the project from a regulation stays
`illustrative` whatever its `derived_by` says. No
independent reviewer is required, and none is available: the code under test
and its expected values come from the same hand ([decision 8](CLEARING_COVERAGE.md#decisions)
of the clearing coverage). Acceptance criteria therefore ask for `published`
values where a registered source prints them, and otherwise for
`illustrative` values from a derivation separate from the code under test
(for example a script under `tests/derivations/`), with their provenance and
tolerances; they do not ask for independent review. Test output reports
illustrative results separately and never as validation.

<a id="tolerance"></a>

## 📏 Tolerance

An actual result `a` matches the expected value `e` when

```text
|a − e| ≤ max(absolute, relative × |e|)
```

Both bounds are stated explicitly on every output; there is no hidden default.
Use these conventions when choosing them:

| Reference | `absolute` | `relative` |
| --- | --- | --- |
| `independent` or `illustrative`, full precision | `1e-6` currency units (or `1e-9` for factors) | `1e-9` |
| `published`, rounded | Half a unit of the last printed digit: printed in thousands → `500` | `0` |

Ported illustrative catalogue cases keep the catalogue's tolerances; see
[`tests/README.md`](../../tests/README.md#ported-prototype-catalogue).

Near zero, the absolute bound governs: when `|e|` is small the relative term
vanishes, so an expected `0` is matched within `absolute` only. Domain rules
apply before tolerance: a quantity that must be non-negative, such as
`exposure_value`, fails if it is negative, however small.

<a id="categories"></a>

## 🗂️ Case categories

Each fixture declares one `category`:

| Category | Covers |
| --- | --- |
| `nominal` | Typical inputs within the domain |
| `boundary` | Values on a domain edge: zero notional or market value, results that should be exactly zero (fully offsetting trades, negative net value floored at zero), maturity at a regulatory floor or cap, a single trade, a single hedging set |
| `invalid` | Inputs the engine must reject: missing or unknown fields, negative amounts where not allowed, end date before start date, unknown asset class or currency. Expected results are rejection statuses (`trade_status`, `netting_set_status`), not values. JSON cannot hold NaN or infinity, so non-finite inputs are tested directly in a VBA test module instead of a fixture |
| `regime` | Inputs chosen to show a recorded [regime difference](README.md#regime-differences), normally with a `crr` and a `bcbs` expected file |

<a id="completeness-policy"></a>

## ✅ Completeness policy

- A rule in the [traceability table](README.md#traceability) is **validated**
  for a regime only when at least one `published` or `independent` case passes
  for it in that regime.
- Each implemented rule also needs at least one boundary case, and each public
  input domain at least one invalid case. Such a case is a fixture or a case
  of a VBA test suite, mostly `TEST_InputValidation`, which also covers inputs
  JSON cannot express. Three fixtures have category `boundary`:
  `illustrative-ccp-non-qualifying`, `illustrative-ccp-retained-vm` and
  `illustrative-settled-to-market-daily`; the other boundary cases are suite
  cases, such as `ir-bucket-boundaries-crr` and `ir-bucket-boundaries-bcbs`.
  No tool checks that every rule has one; the
  [traceability](README.md#traceability) table lists the cases of each rule.
- Every recorded regime difference has a case that shows it, as a fixture or
  a suite case. The table below lists them; where only one regime has a case,
  it says so.
- A release states which rules are validated, for which regime, and lists
  every rule that is implemented but not validated. Illustrative cases never
  count toward any of this.
- The harness keeps its own completeness check: a run that does not execute
  every expected case and assertion is incomplete, not a pass.
- v1.0.2 validates rules with `published` values only
  ([evidence for v1.0.2](README.md#assumptions-and-scope)); `independent` values
  can come only from an external implementation registered as a source, such
  as the R package SACCR of #352, not from a reviewer.

Where each [regime difference](README.md#regime-differences) is covered:

| Regime difference | Covered by |
| --- | --- |
| Interest-rate aggregation across maturity buckets | Fixture `illustrative-ir-sum-of-absolutes` (`BCBS`); `TEST_InputValidation` cases `params-ir-sum-of-absolutes-crr` and `params-ir-sum-of-absolutes-bcbs` |
| Asset classes (other risks) | Fixture `illustrative-other-risks`, both regimes (T24–T26) |
| Commodity hedging sets (climatic conditions) | Fixture `illustrative-climatic-commodity`, `CRR` only (T27) |
| Single-name credit factor | Fixture `illustrative-credit-quality-step`, `CRR` only (T28) |
| Supervisory delta of interest-rate and commodity options | Fixture `illustrative-negative-ir-forward`, both regimes (T20–T22); fixture `illustrative-negative-commodity-price`, `CRR` only (T23) |
| Interest-rate maturity bucket 1 | `TEST_InputValidation` cases `ir-bucket-boundaries-crr`, `ir-bucket-boundaries-bcbs`, `ir-bucket-offset-crr` and `ir-bucket-offset-bcbs` |
| MPOR floor of a client facing its clearing member | Fixture `illustrative-clearing-client-listed`, both regimes (MPOR 9 under `CRR`, 14 under `BCBS`); `TEST_InputValidation` cases `netting-set-clearing-client-crr` and `netting-set-clearing-client-bcbs` |
| Variation margin posted under a one-way margin agreement | `TEST_InputValidation` cases `netting-set-unmargined-vm-posted-crr` and `netting-set-unmargined-vm-posted-bcbs` |

The fixture `illustrative-posted-vm-cap` has category `regime`, but the cap of
a margined netting set that it tests is not a regime difference: both regimes
give the same capped EAD (T16, T17), as the note under
[Regime differences](README.md#regime-differences) explains. Likewise, the
fixture `illustrative-clearing-member-listed` has category `regime` but gives
the same outputs in both regimes: a clearing member's MPOR of 5 business days
is the same under both.

<a id="consumption"></a>

## ⚙️ Consumption

`tools/generate_case_tests.py` generates `tests/modules/TEST_Cases.bas` from the
JSON files, and `python tools/check.py` fails when the committed module is out
of date. The generated module holds data only; the hand-written
`tests/modules/TEST_CaseRunner.bas` writes each fixture into the NettingSets and
Trades sheets and the AsOfDate and ReportingCcy parameters, also sets the
case's `parameters` on Params and the FX rate of the calculation currency to
1, runs the engine with the case's regime as the netting set's override,
compares the output cells with the tolerance rule above, and restores the
workbook. Running the cases is described in
[`INSTALLATION.md`](../../INSTALLATION.md#numerical-test-cases).

The file format is enforced by `tools/check_test_cases.py`, part of
`python tools/check.py`: envelopes, field names, types and units, quantity
names and forms, trade references, tolerances, reference classes and their
required fields, sources from the register, the `illustrative-` naming rule,
unique catalogue IDs, and the settlement and CCP inputs above together with
the rejection their expected files declare. It does not check any value
against the engine.
