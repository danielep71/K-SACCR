# CSV input contract

This document specifies version 1 of the two CSV files that carry the inputs of
the `NettingSets` and `Trades` sheets: what a file must contain, how each field
is written and read, and how a file is rejected. It is the contract for the
CSV import of v1.1.0 (#206). Parsing, staging, validation and commit are built
on it in #207 to #211.

The machine-readable contract is [`contract.json`](contract.json). The field
tables below are generated from it, and
[`tools/csv_contract.py`](../../tools/csv_contract.py) implements it as the
reference against which the importer is tested. The
`check_csv_contract` gate keeps the contract, the engine's expected headers,
the template's validation lists, the blank templates, the header lines the
workbook exports, this document's tables and the
[test fixtures](../../tests/csv/) consistent.

## 📄 Files and templates

| File | Sheet | Key | Blank template |
| --- | --- | --- | --- |
| NettingSets | `NettingSets` | `netting_set_id` | [`K-SACCR_NettingSets_v1.csv`](templates/K-SACCR_NettingSets_v1.csv) |
| Trades | `Trades` | `trade_id` | [`K-SACCR_Trades_v1.csv`](templates/K-SACCR_Trades_v1.csv) |

A template is the header record alone. Each CSV column maps to exactly one
column of its sheet, named in the tables below; the CSV uses short
`snake_case` names because several sheet headers contain commas and
brackets.

In the workbook, the **Export CSV template** button on `Trades` and on
`NettingSets` writes the same file, byte for byte, to a file the user
chooses. The macros take the header line from constants in `CORE_Config`,
which the `check_csv_contract` gate compares with the contract.

One record of `Trades` is one **leg**, not necessarily one economic trade. An
economic trade with several risk drivers has one record per leg, all with the
same `economic_trade_id` and `netting_set_id`; exactly one leg has
`mtm_leg` `Y` and carries the whole MtM, and the others have `mtm_leg` `N` and
an MtM of `0`. The engine rejects every leg of an economic trade that breaks
this rule (see the
[user manual](../user-manual/v1.0.2/K-SACCR_v1.0.2_User_Manual.md#multi-leg-economic-trades-prevent-mtm-double-counting)).

## 🧾 File format

- **Encoding.** UTF-8, with or without a byte-order mark. A file that is not
  valid UTF-8, or contains a NUL character, is rejected.
- **Records.** Records are separated by CRLF or LF; the last record may end
  without one. A carriage return anywhere else, outside a quoted field, is an
  error.
- **Fields.** Fields are separated by semicolons (`;`), as in the evidence
  export. Where the comma is the decimal or thousands separator, as in Italy,
  Excel also uses `;` as its list separator, so the file opens in columns;
  and a value such as `1,5` or `1,000` stays one field and is rejected as a
  number, instead of shifting the fields after it. A field that contains a
  semicolon, a quote or a line break is enclosed in double quotes, and a
  quote inside it is written twice (`""`), as in RFC 4180 with `;` in place
  of the comma. A quote inside an unquoted field, text after a closing quote,
  or a quoted field left open at the end of the file is a syntax error. A
  file saved by Excel in such a locale still writes decimal commas and local
  dates, which the [value types](#-value-types) reject.
- **Lines and records.** A quoted field can span several physical lines.
  Every finding names the physical line on which its record starts and the
  record number, counted from 1 after the header; the header is record 0.
- **Spaces.** Spaces are part of a field. A value of type `id`, `code`,
  `flag`, `currency`, `integer`, `number` or `date` with a leading or trailing
  space is rejected, never trimmed.

## 🏷️ Header and schema version

The first record is the header. It names every column of the file's
contract exactly once, in any order; names are case-sensitive. A repeated,
missing or unknown column is an error, and the file is not read further.
There are no optional columns: an optional field is a column whose value may
be blank.

The column set is the schema version. These are the version 1 files, and the
templates carry `_v1` in their name. A later version that adds, removes or
renames a column is a new contract with its own templates; a file is never
mapped to a contract by guessing what a column means.

## 🔤 Value types

| Type | Written as | Rules |
| --- | --- | --- |
| `id` | Text | At most 64 characters; no `\|` or `#`; no control character (Unicode category Cc, which includes DEL and the C1 controls). Read as text, so leading zeros are kept (`00123`). Keys are compared case-insensitively, as the engine compares them: in upper case, one character at a time |
| `text` | Text | At most 500 characters; line breaks allowed, other control characters (category Cc) not |
| `code` | One of the listed values | Exact spelling and case, as in the sheet's validation list (`Linear`, `IR`, `CLIENT`); `linear` is rejected |
| `flag` | `Y` or `N` | Exact; `y`, `Yes`, `TRUE` or `1` are rejected |
| `currency` | Three capital letters | `EUR`, not `eur` or `€` |
| `integer` | `-?digits` | No sign `+`, no leading zero, no decimal point |
| `number` | `-?digits[.digits][e±digits]` | Decimal point `.`, no thousands separator, no sign `+`; at most 15 significant digits, so a value is never rounded when it is read; a non-zero magnitude from 2.2250738585072014e-308 to 9.99999999999999e307, as a worksheet cell holds; `NaN` and `inf` are not numbers |
| `date` | `yyyy-mm-dd` | A calendar date from 1950-01-01 to 2200-12-31, as the sheet's validation allows; `31/03/2027` and serial numbers are rejected |

Text that starts with `=`, `+`, `-`, `@`, a tab or a carriage return is
rejected (`CSV-FORMULA`), so a field can never become a formula when it is
written to a sheet. The rule covers `id` and `text` fields; a negative
`number` or `integer` is a value, not a formula.

## ⬜ Blank fields, blank records and empty files

- A **blank field** is an empty one. A required field must not be blank; an
  optional blank field takes the meaning the tables give, which is the
  sheet's.
- A **blank record**, whose fields are all empty, and an **empty line** are
  errors; a file is not padded with them.
- A **header-only file** is valid and holds no record. Committing it never
  clears the live inputs on its own: replacing a sheet's inputs with none is
  an explicit choice, made when the file is committed (#211).
- A file is accepted or rejected **as a whole**: one finding rejects every
  record, and no record is ever dropped or truncated.

## 📏 Limits

| File | Records per file |
| --- | --- |
| NettingSets | 10,000 |
| Trades | 100,000 |

A file of more than 128 MiB (134,217,728 bytes, a byte-order mark included)
is rejected from its size alone, before any of it is read
(`CSV-TOO-LARGE`). 100,000 Trades records of the size of the fixtures' are
about 10 MB.

A file with more records is rejected whole (`CSV-TOO-MANY-RECORDS`). These
limits are the import's, and differ from two other numbers. The sheets'
validation lists cover `NettingSets` rows 7 to 206 and `Trades` rows 5 to
5,000, which bounds only typing on the sheet. The engine reads every row
down to the last one used, and was measured with 500,000 legs
([portfolio size](../../INSTALLATION.md#portfolio-size)). How the staging
sheets and the protected input sheets hold files up to these limits is
agreed in #207, #208 and #420.

## 📋 NettingSets columns

<!-- BEGIN GENERATED: NettingSets -->
| CSV column | Sheet header | Type | Required | Blank means | Description | Example |
| --- | --- | --- | --- | --- | --- | --- |
| `netting_set_id` | NettingSetID | id (<= 64 characters) | yes |  | Unique netting-set ID, compared case-insensitively; referenced by every trade. | `NS-001` |
| `counterparty` | Counterparty | text (<= 500 characters) | no |  | Descriptive counterparty name or identifier. | `CPTY-A` |
| `margined` | Margined (Y/N) | flag: `Y`, `N` | no | N | Y for a supported margin agreement. | `Y` |
| `clearing_role` | Client clearing role (N / CM / CLIENT) | code: `N`, `CM`, `CLIENT`, `CCP`, `MULTI` | no | N | Bank's clearing role; MULTI is accepted by the file and makes the netting set INVALID in the engine. | `CM` |
| `remargin_period_business_days` | Remargin frequency N (BD) | integer (>= 0, <= 10000) | no | 1 | Remargining every N business days; 0 is read as daily with a warning. | `1` |
| `large_or_illiquid` | >5,000 trades or illiquid collateral (Y/N) | flag: `Y`, `N` | no | N | Y applies the 20-business-day MPOR floor. | `N` |
| `margin_disputes` | Margin disputes (Y/N) | flag: `Y`, `N` | no | N | Y doubles the MPOR. | `N` |
| `mpor_override_business_days` | MPOR override (BD) | integer (>= 1) | no | no override | MPOR in business days, applied only when not below the computed MPOR. | `20` |
| `variation_margin_net_amount` | Net VM held (+) / posted (-) | number | no | 0 | Net variation margin: held positive, posted negative. | `-150000` |
| `independent_collateral_net_amount` | NICA | number | no | 0 | Net independent collateral amount. | `250000` |
| `threshold_amount` | Threshold TH | number (>= 0) | no | 0 | Margin threshold, unscaled. | `0` |
| `minimum_transfer_amount` | MTA | number (>= 0) | no | 0 | Minimum transfer amount, unscaled. | `100000` |
| `alpha_factor` | Alpha override | number (> 0) | no | the Params alpha | Alpha for this netting set; a value other than 1.4 classifies the run WHAT-IF. | `1.4` |
| `regime_override` | Regime override (blank = Params) | code: `CRR`, `BCBS` | no | the Params regime | Regime of this netting set and all its trades. | `CRR` |
| `margin_agreements` | Margin agreements (blank / MIXED / SHARED) | code: `MIXED`, `SHARED` | no | a single agreement or none | MIXED or SHARED declares an unsupported structure, which the engine rejects. |  |
| `qccp` | QCCP (Y/N) | flag: `Y`, `N` | no |  | Role CCP only, and then required: Y if the CCP is qualifying. |  |
| `ccp_retains_variation_margin` | CCP retains VM, collateral not protected (Y/N) | flag: `Y`, `N` | no |  | Role CCP only, and then required: Y if the CCP retains VM and the collateral is not protected. |  |
<!-- END GENERATED: NettingSets -->

Reserved keys: `TOTAL`, `TOTAL CRR` and `TOTAL BCBS`, in any case, label the
total rows on `Results` and are rejected as netting-set IDs.

## 📋 Trades columns

<!-- BEGIN GENERATED: Trades -->
| CSV column | Sheet header | Type | Required | Blank means | Description | Example |
| --- | --- | --- | --- | --- | --- | --- |
| `trade_id` | TradeID | id (<= 64 characters) | yes |  | Unique leg ID, compared case-insensitively. | `TRD-101-IR` |
| `netting_set_id` | NettingSetID | id (<= 64 characters) | yes |  | ID of the leg's netting set in NettingSets. | `NS-001` |
| `asset_class` | Asset class | code: `IR`, `FX`, `CR`, `EQ`, `CO`, `OT` | yes |  | Asset class; OT is CRR only. | `IR` |
| `sub_class` | Sub-class | id (<= 64 characters) | no |  | Regulatory sub-class, such as AAA or CQS3 for credit, SINGLE or INDEX for equity; required by the engine where the asset class needs one. |  |
| `risk_factor` | Risk factor / reference | id (<= 64 characters) | yes |  | Currency, currency pair, reference entity or other hedging key. | `EUR` |
| `instrument` | Instrument | code: `Linear`, `Option`, `CDO` | no | Linear | Instrument type. | `Linear` |
| `direction` | Direction | code: `Long`, `Short` | yes |  | Long or Short. | `Long` |
| `option_type` | Option type | code: `Call`, `Put` | no |  | Call or Put; required by the engine for an option. |  |
| `nature` | Nature | code: `Standard`, `Basis`, `Volatility` | no | Standard | Hedging-set nature. | `Standard` |
| `hedging_set_label` | Basis / vol hedging-set label | id (<= 64 characters) | no |  | Label of a basis or volatility hedging set; required by the engine for those. |  |
| `notional_amount` | Notional | number (> 0) | yes |  | Unscaled notional of this leg. | `10000000` |
| `notional_currency` | Notional ccy | currency | no | the reporting currency | Currency of the notional. | `EUR` |
| `market_value_amount` | MtM | number | yes |  | Mark-to-market of the leg; in an economic trade, only its MtM leg carries a value. | `250000` |
| `market_value_currency` | MtM ccy | currency | no | the notional currency | Currency of the MtM. |  |
| `start_date` | Start date (S) | date | no |  | Start date S. | `2026-10-01` |
| `end_date` | End date (E) | date | no |  | End date E; after the start date when both are given. | `2031-10-01` |
| `maturity_date` | Maturity date (M) | date | no |  | Contractual maturity M. | `2031-10-01` |
| `option_expiry_date` | Option expiry (T) | date | no |  | Option expiry T; required by the engine for an option. |  |
| `underlying_price` | Underlying price P | number | no |  | Option underlying price or rate P. |  |
| `strike_price` | Strike K | number | no |  | Option strike K. |  |
| `lambda_price` | Lambda shift | number (>= 0) | no |  | Lambda shift added to P and K. |  |
| `attachment_rate` | Attachment A | number (>= 0, <= 0.9999) | no |  | CDO tranche attachment point, a fraction. |  |
| `detachment_rate` | Detachment D | number (> 0, <= 1) | no |  | CDO tranche detachment point, a fraction above the attachment point. |  |
| `economic_trade_id` | Economic trade ID | id (<= 64 characters) | no | the row is a whole trade | Groups the legs of one economic trade. | `TRD-101` |
| `mtm_leg` | MtM leg (Y/N) | flag: `Y`, `N` | no |  | Y on the one leg of an economic trade that carries its whole MtM. | `Y` |
| `settlement_type` | Settlement type (N/STM) | code: `N`, `STM` | no | N | STM for a settled-to-market trade. |  |
| `next_reset_date` | Next reset date | date | no |  | STM only, and then required: the first settlement date after the reporting date. |  |
| `comment` | Comment | text (<= 500 characters) | no |  | Free-text note; not read by the engine. |  |
<!-- END GENERATED: Trades -->

## 🔗 Rules between fields of a record

| File | Rule | Requirement |
| --- | --- | --- |
| NettingSets | `ccp-flags` | With `clearing_role` `CCP`, `qccp` and `ccp_retains_variation_margin` are required; with any other role both are blank |
| Trades | `stm-reset` | With `settlement_type` `STM`, `next_reset_date` is required; otherwise it is blank |
| Trades | `end-after-start` | When both are given, `end_date` is after `start_date` |
| Trades | `detachment-above-attachment` | When both are given, `detachment_rate` is above `attachment_rate` |

A rule is not judged when one of its fields, the field its condition reads included, already has a finding of its own.

## 🚫 What the contract does not check

The contract checks each file on its own. The engine's validation, run on the
staged inputs (#209), still applies, unchanged, to everything that depends on
other records, on the other file or on Params:

- that every `netting_set_id` of `Trades` exists in `NettingSets`;
- the sub-class required by an asset class and regime, the supervisory-factor
  key, `OT` being CRR-only;
- an option's type, expiry and price and strike, the FX pair and its quote
  order, the hedging-set label of a basis or volatility trade;
- the maturity relative to the reporting date, the reset date relative to the
  maturity, an `STM` trade in a margined netting set;
- the rules of an economic trade across its legs;
- the currencies having a rate in the FX table.

A file that passes the contract can therefore still have netting sets that
are `INVALID` and trades that are excluded, exactly as on the sheets.

## ⚠️ Findings

Each finding has a stable code, the physical line, the record and the column.

| Code | Meaning |
| --- | --- |
| `CSV-ENCODING` | Not valid UTF-8, or a NUL character |
| `CSV-SYNTAX` | Not valid CSV; reading stops at the first syntax error |
| `CSV-EMPTY` | No header record |
| `CSV-HEADER-DUPLICATE` | A column named twice in the header |
| `CSV-HEADER-MISSING` | A contract column missing from the header |
| `CSV-HEADER-UNKNOWN` | A header column not in the contract |
| `CSV-FIELD-COUNT` | A record with more or fewer fields than the header |
| `CSV-BLANK-RECORD` | A record with every field blank, or an empty line |
| `CSV-TOO-MANY-RECORDS` | More records than the file's limit |
| `CSV-TOO-LARGE` | More bytes than the file size limit |
| `CSV-REQUIRED` | A required field is blank |
| `CSV-TYPE` | A value not of its type's form |
| `CSV-RANGE` | A value outside its allowed range |
| `CSV-ALLOWED` | A value not among the allowed values |
| `CSV-LENGTH` | Text longer than its type allows |
| `CSV-FORMULA` | Text that starts like a formula |
| `CSV-DUPLICATE-KEY` | A key that repeats an earlier record, ignoring case |
| `CSV-RESERVED-KEY` | A reserved key |
| `CSV-ROW-RULE` | A rule between fields of the record broken |

## 🧪 Reference implementation and fixtures

```sh
python tools/csv_contract.py --kind Trades path/to/trades.csv
python tools/csv_contract.py --write-templates
```

The first command validates a file and lists its findings; the second
writes the blank templates and the field tables of this document from the
contract. The fixtures in [`tests/csv/`](../../tests/csv/) are synthetic:
`valid/` files must be accepted, `invalid/` files must be rejected with
exactly the findings recorded in
[`expected.json`](../../tests/csv/expected.json). The inline cases of
[`cases.json`](../../tests/csv/cases.json) pin the edge cases the fixtures
leave out: invalid UTF-8, characters outside the basic plane, spaces and
control characters, number forms and magnitudes, dates, keys, quoting and the
limits. The tests in `tools/test_csv_contract.py` also read every fixture
with a byte-order mark and with CRLF line ends, and expect the same outcome.

## 📥 The workbook's reader

The workbook reads a file with `CORE_CsvReader` (#207), which applies this
contract without Excel converting anything: it decodes the bytes itself,
checks numbers as text and converts them with `Val`, which always reads a
`.` decimal point, and checks dates digit by digit, so the outcome does not
depend on the Windows locale. It writes nothing: a file is accepted or
rejected whole, and an accepted file's records are returned as their raw
text, in contract column order, for the staging sheets (#208). Until they
exist, the macros `CheckTradesCsv` and `CheckNettingSetsCsv` (Macros dialog,
Alt+F8) check a file and list its first findings with their line, record and
column, and `TEST_Benchmark.RunCsvBenchmark` measures the reader on files of
up to the record limit.

`tools/generate_csv_vba.py` writes the contract into `CORE_CsvSchema` and
every fixture and case into the test suite `TEST_CsvReader`, which expects
the reader to report exactly the findings `tools/csv_contract.py` reports;
`python tools/check.py` fails when either module is out of date.
