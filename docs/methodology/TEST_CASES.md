# 🧪 Numerical Test Cases

This document defines how SACCR's numerical test cases are written: inputs,
units, expected results, their provenance, tolerances, case categories and when
the suite counts as complete. Sources and rule IDs come from the
[methodology](README.md).

> [!IMPORTANT]
> The prototype's 28 TestCatalogue checks are ported as cases (#59): 14 values
> printed in BCBS 279 Annex 4 are `published`; the other 14 are `illustrative`.
> They run in Excel through `TEST_Cases.RunCaseTests` ([consumption](#consumption));
> no rule is validated until a run passes and its published values are checked
> against the source. The mapping is in
> [`tests/README.md`](../../tests/README.md#ported-prototype-catalogue).

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
    "alpha_factor": null
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
      "description": "Five-year swap, format example"
    }
  ]
}
```

| Rule | Detail |
| --- | --- |
| Envelope | `schema_version`, `kind`, `id` (equal to the file name), `synthetic: true`, `description` and `category` are required |
| Units | In `netting_set` and `trades`, a field's suffix states its unit: `_amount` in `calculation_currency`, unscaled; `_date` as ISO `YYYY-MM-DD`; `_years` in years; `_business_days` in business days; `_rate` and `_factor` as decimals; `underlying_price`, `strike_price` and `lambda_price` share the underlying's unit, a decimal rate for an interest-rate option and a price otherwise. A quantity there never lacks a unit suffix. Envelope fields such as `schema_version` are metadata, not quantities |
| Currencies | Every amount is in `calculation_currency`; amounts in other currencies are converted before they enter a fixture ([decision 6](README.md#assumptions-and-scope)). `risk_factor` names what the trade is exposed to: the interest-rate currency, the FX pair, the credit or equity reference, or the commodity |
| Signs | Market values are from the bank's side: positive is an asset. `variation_margin_net_amount` and `independent_collateral_net_amount` are held minus posted. `notional_amount` is unsigned; `direction` gives the sign |
| Parameters | A fixture runs with the default parameters of the workbook template's Params sheet (for example 365 days and 250 business days a year). `parameters` lists the Params values the case runs with instead, by code; `{}` for none. Only `DaysPerYear` (positive) is accepted so far, so that a published maturity such as 0.75 years can be entered as dates; a case that needs another parameter adds it to the format first |
| Fields | Every field below is present in every fixture; `null` means not applicable or the default. `tools/check_test_cases.py` rejects an unknown, missing or mistyped field |

| Netting-set field | Type | Meaning |
| --- | --- | --- |
| `id` | text | Netting-set ID |
| `margined`, `large_or_illiquid`, `margin_disputes` | Boolean | As on the NettingSets sheet |
| `clearing_role` | `none`, `clearing_member` or `client` | Client clearing role on the NettingSets sheet: N, CM or CLIENT |
| `remargin_period_business_days` | number or `null` | Remargining period of a margined set |
| `mpor_override_business_days` | number or `null` | MPOR override; `null` for none |
| `variation_margin_net_amount`, `independent_collateral_net_amount` | amount | Net VM and NICA, held minus posted |
| `threshold_amount`, `minimum_transfer_amount` | amount, at least 0 | Margin agreement terms |
| `alpha_factor` | number or `null` | Alpha override; `null` uses the regime's alpha |

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

An invalid-input case replaces `value`, `unit` and `tolerance` with the error
the facade must raise, by its public constant name. The name below is an
example; no such constant exists yet:

```json
{"quantity": "exposure_value", "rule": "CRR.274.2", "error": "SACCR_ERROR_INVALID_NOTIONAL", "reference": {"class": "illustrative", "...": "..."}}
```

An output is one of three forms: a number (`value`, `unit`, `tolerance`), a
text result (`text`, compared exactly, no unit or tolerance) or an error
(`error`). It may also carry `catalogue`, the prototype TestCatalogue test it
was ported from, such as `T01`.

Quantities use stable names, each with one meaning only:

| Level | Quantity | Form and unit |
| --- | --- | --- |
| Netting set | `exposure_value`, `replacement_cost`, `potential_future_exposure`, `aggregate_add_on`, `add_on.<asset_class>` | Number in the calculation currency |
| Netting set | `multiplier` | Number, unit `1` |
| Netting set | `margin_period_of_risk` | Number, unit `business_days` |
| Netting set | `cap_applied` | Text: `Y` when the unmargined cap sets the exposure value |
| Netting set | `netting_set_status` | Text: `VALID`; `INCOMPLETE: <r> of <n> trade(s) rejected` or `INVALID: <e> input error(s)`, when the exposure value is withheld; or `NO TRADES` |
| Trade | `adjusted_notional` | Number in the calculation currency |
| Trade | `supervisory_delta`, `supervisory_factor` | Number, unit `1` |
| Trade | `lambda_shift` | Number in the underlying's unit: `1` for a rate, `price` otherwise |
| Trade | `hedging_set` | Text: the engine's hedging-set key, for example `CO\|CLIMATIC\|STD` |
| Trade | `trade_status` | Text: `OK` or `EXCLUDED` |

A trade-level output names its trade in `trade`; a netting-set output has no
`trade`. New quantities are added to this table and to the validator together.

Rule IDs follow the [traceability](README.md#traceability) convention:
`CRR.<article>[.<paragraph>]`, `CRR-RTS.<article>` for Delegated Regulation
(EU) 2021/931, and `CRE52.<paragraph>`, with `CRE52` alone for a paragraph not
yet located. The rule IDs of the ported cases are those the prototype cites and
are verified with the sources.

<a id="reference-classes"></a>

## 🏷️ Reference classes

Every expected value carries exactly one class. Only the first two are
evidence that the implementation is correct.

| Class | Meaning | Required fields |
| --- | --- | --- |
| `published` | Printed in a source from the register, for example a worked example | `source`, exact `locator` (section, example, table) |
| `independent` | Computed outside SACCR's code, by a separate method such as a reviewer's spreadsheet or script, from the cited rules | `source`, `locator`, `derivation`, `derived_by`, `reviewed_by` |
| `illustrative` | Shows the format or a behaviour; not a validated SA-CCR figure | `derivation` saying why it exists |

An `independent` value is never produced by running SACCR itself, and never
reviewed only by the person who wrote the code under test. Test output reports
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
| `invalid` | Inputs the facade must reject: missing or unknown fields, negative amounts where not allowed, end date before start date, unknown asset class or currency. Expected results are errors, not values. JSON cannot hold NaN or infinity, so non-finite inputs are tested directly in a VBA test module instead of a fixture |
| `regime` | Inputs chosen to show a recorded [regime difference](README.md#regime-differences), with a `crr` and a `bcbs` expected file |

<a id="completeness-policy"></a>

## ✅ Completeness policy

- A rule in the [traceability table](README.md#traceability) is **validated**
  for a regime only when at least one `published` or `independent` case passes
  for it in that regime.
- Each implemented rule also has at least one `boundary` case, and each public
  input domain has at least one `invalid` case, as a fixture or, for inputs
  JSON cannot express, as a direct test in a VBA test module.
- Every recorded regime difference has a `regime` case.
- A release states which rules are validated, for which regime, and lists
  every rule that is implemented but not validated. Illustrative cases never
  count toward any of this.
- The harness keeps its own completeness check: a run that does not execute
  every expected case and assertion is incomplete, not a pass.
- v1.0.0 validates rules with `published` values only
  ([decision 8](README.md#assumptions-and-scope)); `independent` values
  start in a later release, with a reviewer other than the author of the
  code under test.

<a id="consumption"></a>

## ⚙️ Consumption

Decided by the owner on 2026-10-06 (decision 7, #44):
`tools/generate_case_tests.py` generates `tests/modules/TEST_Cases.bas` from the
JSON files, and `python tools/check.py` fails when the committed module is out
of date. The generated module holds data only; the hand-written
`tests/modules/TEST_CaseRunner.bas` writes each fixture into the NettingSets and
Trades sheets and the AsOfDate and ReportingCcy parameters, runs the engine
with the case's regime as the netting set's override, compares the output
cells with the tolerance rule above, and restores the workbook. Running the
cases is described in
[`INSTALLATION.md`](../../INSTALLATION.md#numerical-test-cases).

The file format is enforced by `tools/check_test_cases.py`, part of
`python tools/check.py`: envelopes, field names, types and units, quantity
names and forms, trade references, tolerances, reference classes and their
required fields, sources from the register, the `illustrative-` naming rule,
and unique catalogue IDs. It does not check any value against the engine.
