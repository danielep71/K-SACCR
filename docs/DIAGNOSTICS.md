# Stable Checks diagnostics

From v1.3.0, Checks and the exported `CHECKS` section have eight columns:

| Position | Header | Contract |
| --- | --- | --- |
| 1–6 | Severity, Sheet, Row, Reference, Message, Origin | Existing positions preserved. |
| 7 | Code | Stable, case-sensitive rule identifier. |
| 8 | Field | Field concerned; empty for a rule spanning fields or the whole run. |

`CORE_Config.bas` declares the codes. [diagnostics.json](diagnostics.json) is the
reviewable registry of meanings and active/retired status. No retired values may
be reused or deleted. Change wording freely; a changed rule meaning needs a new
code. Shared validators use one code with explicit field context (for example
`NS-FLAG` + `QCCP`), not a separate identifier per field. INFO rows also have codes.
The existing `CSV-*` identifiers and staging message prefixes retain their meaning.
CSV field names remain the contract's snake_case names; engine fields use workbook
labels. CSV diagnostics are documented in [the CSV contract](csv/CSV_CONTRACT.md).

The first run, input validation, staging validation or clear operation upgrades
older templates in place: Code and Field are appended after Origin, with no
column insertion. Rebuild from current sources before using the new interface.
The headers and data are covered by the existing output-layout and protection
lifecycle. There is no calculation-formula change.

Trade violations are now separate Checks rows instead of one concatenated ERROR
per rejected trade. The TradeCalc reason text stays aggregated; error counts now
count individual violations. Rejected-trade counts and EAD withholding do not
change. Economic-trade violations appear for each affected leg. Consumers must
match Code (and Field when appropriate), rather than message wording or a fixed
number of error lines. Existing INFO prose assertions can still verify detailed
regulatory explanations.

`RunSACCR_Silent` and `ValidateInputs_Silent` add `; codes=...` to the result
line, just before the closing `; cleanup=...`: distinct ERROR/WARNING codes in
encounter order, comma-separated, or empty if there are none. Codes come from the current engine log. Clear does not report
codes. Existing result fields retain their meaning except the more granular
error count. Consumers should parse named fields, not require an exact whole-line
match. Unexpected raised exceptions still use the established exception contract.

Evidence CSV export and output fingerprints use `CK_NCOLS = 8`: Code/Field are
included and changing either makes exported output stale. Earlier evidence stays
historical evidence with its original six-column schema; it must not be rewritten.

## Validation

`python tools/check_diagnostics.py --root .` runs in `tools/check.py`. It rejects
a Checks width other than eight or missing/mispositioned Code/Field constants,
missing/unknown engine emission codes, registry/constant mismatches, retired code
use and unused active codes. Only the private, reviewed forwarding wrappers may
pass a dynamic code. Mutation tests cover comments, strings, continuations,
missing/unknown codes, retirement, registry drift and the shared output width. Static analysis does not
replace Excel compilation or runtime verification.

The input-validation suite checks ERROR/WARNING identifiers, with explicit field
checks for shared text/date/flag validators; `RunCaseExpectingStop` now takes a
code. INFO-detail tests retain their wording checks. The staging suite verifies
Code/Field alongside its existing origin, cell-link and message contract.
Run `tools/Verify-SACCR-Branch.ps1 -Branch feat/81-stable-diagnostics -Quick` in
Excel before merge, including compilation, all suites and protection checks.

## Registry

| Code | Status | Meaning / existing message context |
| --- | --- | --- |
| `CSV-ALLOWED` | active | Existing CSV contract diagnostic; meaning unchanged. |
| `CSV-BLANK-RECORD` | active | Existing CSV contract diagnostic; meaning unchanged. |
| `CSV-CONTEXT-CURRENCY` | active | Existing CSV contract diagnostic; meaning unchanged. |
| `CSV-CONTEXT-EXPIRED` | active | Existing CSV contract diagnostic; meaning unchanged. |
| `CSV-CONTEXT-FACTOR` | active | Existing CSV contract diagnostic; meaning unchanged. |
| `CSV-CONTEXT-MATURED` | active | Existing CSV contract diagnostic; meaning unchanged. |
| `CSV-CONTEXT-PAIR` | active | Existing CSV contract diagnostic; meaning unchanged. |
| `CSV-CONTEXT-PARAMS` | active | Existing CSV contract diagnostic; meaning unchanged. |
| `CSV-CONTEXT-REGIME` | active | Existing CSV contract diagnostic; meaning unchanged. |
| `CSV-CONTEXT-RESET` | active | Existing CSV contract diagnostic; meaning unchanged. |
| `CSV-DELIMITER` | active | Existing CSV contract diagnostic; meaning unchanged. |
| `CSV-DUPLICATE-KEY` | active | Existing CSV contract diagnostic; meaning unchanged. |
| `CSV-ECON-ASSET-CLASS` | active | Existing CSV contract diagnostic; meaning unchanged. |
| `CSV-ECON-MTM-LEG` | active | Existing CSV contract diagnostic; meaning unchanged. |
| `CSV-ECON-MTM-VALUE` | active | Existing CSV contract diagnostic; meaning unchanged. |
| `CSV-ECON-NETTING-SET` | active | Existing CSV contract diagnostic; meaning unchanged. |
| `CSV-ECON-SETTLEMENT` | active | Existing CSV contract diagnostic; meaning unchanged. |
| `CSV-EMPTY` | active | Existing CSV contract diagnostic; meaning unchanged. |
| `CSV-ENCODING` | active | Existing CSV contract diagnostic; meaning unchanged. |
| `CSV-FIELD-COUNT` | active | Existing CSV contract diagnostic; meaning unchanged. |
| `CSV-FORMULA` | active | Existing CSV contract diagnostic; meaning unchanged. |
| `CSV-HEADER-DUPLICATE` | active | Existing CSV contract diagnostic; meaning unchanged. |
| `CSV-HEADER-MISSING` | active | Existing CSV contract diagnostic; meaning unchanged. |
| `CSV-HEADER-UNKNOWN` | active | Existing CSV contract diagnostic; meaning unchanged. |
| `CSV-LENGTH` | active | Existing CSV contract diagnostic; meaning unchanged. |
| `CSV-NETTING-SET-UNKNOWN` | active | Existing CSV contract diagnostic; meaning unchanged. |
| `CSV-RANGE` | active | Existing CSV contract diagnostic; meaning unchanged. |
| `CSV-REQUIRED` | active | Existing CSV contract diagnostic; meaning unchanged. |
| `CSV-RESERVED-KEY` | active | Existing CSV contract diagnostic; meaning unchanged. |
| `CSV-ROW-RULE` | active | Existing CSV contract diagnostic; meaning unchanged. |
| `CSV-SYNTAX` | active | Existing CSV contract diagnostic; meaning unchanged. |
| `CSV-TOO-LARGE` | active | Existing CSV contract diagnostic; meaning unchanged. |
| `CSV-TOO-MANY-RECORDS` | active | Existing CSV contract diagnostic; meaning unchanged. |
| `CSV-TYPE` | active | Existing CSV contract diagnostic; meaning unchanged. |
| `ECON-ASSET-CLASS` | active | it has   legs in asset class |
| `ECON-ID` | active | Economic trade ID must be text or blank (found an error value on row |
| `ECON-MTM-COUNT` | active | exactly one leg must be the MtM leg (MtM leg = Y), found |
| `ECON-MTM-FLAG` | active | MtM leg must be Y or N (found an error value on row |
| `ECON-MTM-VALUE` | active | leg   is not the MtM leg, so its MtM |
| `ECON-NETTING-SET` | active | its legs are in different netting sets ( , |
| `ECON-RESET` | active | its legs have different next reset dates |
| `ECON-SETTLEMENT` | active | its legs have different settlement types ( |
| `ECON-WHOLE-MTM` | active | a trade without an Economic trade ID carries its own MtM, so |
| `FX-DUP` | active | Currency listed twice with different rates. |
| `FX-DUP-SAME` | active | Currency listed twice with the same rate - first row used. |
| `FX-RATE` | active | FX rate missing or not positive - currency ignored. |
| `FX-REPORTING-RATE` | active | Rate of the reporting currency must be 1 (found |
| `FX-TABLE` | active | FX table not found (workbook name ' '). |
| `IMPORT-SUMMARY` | active | Staging validation summary. |
| `LAYOUT-DATE-SYSTEM` | active | The workbook uses the 1904 date system, which is not supported. |
| `LAYOUT-HEADER` | active | A table header differs from the required column name. |
| `LAYOUT-HEADER-NAME` | active | Workbook name ' |
| `LAYOUT-PARAM-NAME` | active | Workbook name ' |
| `LAYOUT-PARAM-POSITION` | active | Workbook name ' |
| `LAYOUT-PARAM-ROW` | active | Workbook names ' ' and ' |
| `LAYOUT-SHEET` | active | Sheet ' ' not found. |
| `NS-ALPHA` | active | Alpha override must be greater than zero (found |
| `NS-CCP-FLAG-REQUIRED` | active | (Y/N) is required for Client clearing role CCP. |
| `NS-CCP-UNMARGINED` | active | on an unmargined netting set: no MPOR |
| `NS-CLEARING-UNMARGINED` | active | Client clearing role   on an unmargined |
| `NS-CODE` | active | must be  / , |
| `NS-DUP-ID` | active | Netting set ID listed twice - row ignored and EAD withheld. |
| `NS-EMPTY` | active | No netting sets defined. |
| `NS-FLAG` | active | must be Y or N (found  ). |
| `NS-ID-MISSING` | active | Row has netting-set data but no ID - row ignored. |
| `NS-ID-RESERVED` | active | Netting set ID   is reserved for a TOTAL row on Results - row ignored. |
| `NS-INPUT-INCOMPLETE` | active | EAD withheld: |
| `NS-IR-OFFSET` | active | IRBucketOffset is FALSE, but the |
| `NS-LARGE` | active | Netting set has |
| `NS-MARGIN-MULTIPLE` | active | Several margin agreements, or margined and unmargined |
| `NS-MARGIN-SHARED` | active | One margin agreement over several netting sets is not |
| `NS-MPOR` | active | MPOR override must be a whole number of business days |
| `NS-MPOR-FLOOR` | active | MPOR override |
| `NS-MPOR-INFO` | active | . |
| `NS-MTA` | active | MTA must not be negative (found  ). |
| `NS-MULTILEVEL` | active | A multi-level client structure is not supported |
| `NS-NO-TRADES` | active | Netting set has no trades - no EAD. |
| `NS-NONE-VALID` | active | No valid netting sets. |
| `NS-NUMBER` | active | must be a number (found  ). |
| `NS-ONEWAY` | active | VM posted under a one-way margin agreement: unmargined |
| `NS-QCCP-ROLE` | active | QCCP applies only to Client clearing role CCP - |
| `NS-REGIME` | active | Regime override must be BCBS or CRR (found |
| `NS-REMARGIN` | active | Remargin frequency must be a whole number of business days |
| `NS-REMARGIN-FLOOR` | active | Remargin frequency |
| `NS-RETAINS-ROLE` | active | CCP retains VM applies only to Client clearing role |
| `NS-RETAINS-UNMARGINED` | active | CCP retains VM is Y on an unmargined netting set: |
| `NS-STM-MARGINED` | active | Trade   is settled-to-market (STM) in a |
| `NS-TH` | active | Threshold TH must not be negative (found  ). |
| `NS-TRADE-INCOMPLETE` | active | EAD withheld: |
| `NS-VM-POSTED` | active | VM posted on an unmargined CRR netting set: under CRR |
| `NS-VM-RECEIVED` | active | VM received on an unmargined netting set: a counterparty |
| `PRM-ALPHA` | active | Alpha must be greater than zero. |
| `PRM-ASOF` | active | Reporting date missing or not a date (enter a date, |
| `PRM-BOOLEAN` | active | Parameter must be TRUE or FALSE (found |
| `PRM-CURRENCY` | active | Reporting currency must be three letters A to Z, |
| `PRM-CURRENCY-MISSING` | active | Reporting currency missing. |
| `PRM-DEFAULT` | active | Parameter missing - default   used. |
| `PRM-DUP` | active | Parameter listed twice with different values (rows |
| `PRM-DUP-SAME` | active | Parameter listed twice with the same value (rows |
| `PRM-IR-CORRELATION` | active | The IR bucket correlations  , |
| `PRM-NAME` | active | Workbook name ' |
| `PRM-NUMBER` | active | Parameter must be a number (found  ). |
| `PRM-RANGE` | active | must be   (found  ). |
| `PRM-READ` | active | Cannot read parameters: |
| `PRM-REGIME` | active | Regime must be BCBS or CRR (found an error value). |
| `PRM-REGIME-DEFAULT` | active | Regime missing - CRR assumed. |
| `RUN-CLEAN` | active | No issues found. |
| `RUN-ERROR` | active | Run stopped by error  : |
| `RUN-TOTAL` | active | TOTAL |
| `RUN-WHATIF` | active | What-if: |
| `SF-DUP` | active | Supervisory factor key listed twice with different values. |
| `SF-DUP-SAME` | active | Supervisory factor key listed twice with the same values - |
| `SF-EMPTY` | active | Supervisory factor table is empty. |
| `SF-NUMBER` | active | Supervisory factor, correlation and volatility must be numbers. |
| `SF-RANGE` | active | Supervisory factor and correlation must be from 0 to 1 and |
| `SF-TABLE` | active | Supervisory factor table not found (workbook name ' |
| `TABLE-GAP` | active | row below a blank row - not read; |
| `TABLE-KEY` | active | row has a value but no key - not read. |
| `TR-ASSET-CLASS` | active | asset class must be IR, FX, CR, EQ, CO or OT |
| `TR-BASIS-LABEL` | active | basis trade needs a hedging-set label |
| `TR-CDO-CLASS` | active | CDO tranches are credit (CR) trades |
| `TR-CDO-POINTS` | active | CDO tranche needs attachment and detachment points |
| `TR-CDO-RANGE` | active | attachment/detachment must satisfy 0 <= A < D <= 1 |
| `TR-COMMODITY-GROUP` | active | commodity hedging group missing for ' ' |
| `TR-DATE` | active | is not a date (found  ) |
| `TR-DATE-ORDER` | active | end date must be after start date |
| `TR-DIRECTION` | active | direction must be Long or Short |
| `TR-DUP-ID` | active | duplicate Trade ID (first used on row  ) |
| `TR-EMPTY` | active | No trades found. |
| `TR-EXPIRED` | active | The option has expired on or before the reporting date. |
| `TR-EXPIRY-MATURITY` | active | option expiry is after the maturity date (M is the latest |
| `TR-EXPIRY-MISSING` | active | option expiry missing |
| `TR-FACTOR` | active | No supervisory-factor row exists for the asset class/sub-class. |
| `TR-FACTOR-REGIME` | active | The supervisory-factor row excludes the effective regime. |
| `TR-FX-NOTIONAL` | active | notional currency ' ' is not a currency of the pair |
| `TR-FX-PAIR` | active | FX risk factor must be a currency pair such as EUR/USD |
| `TR-FX-PAIR-RATE` | active | A currency in the FX pair has no usable FX rate. |
| `TR-FX-QUOTE` | active | FX option must be entered in market quote order as |
| `TR-ID-MISSING` | active | Row has trade data but no Trade ID - trade rejected. |
| `TR-INSTRUMENT` | active | instrument type must be Linear, Option or CDO |
| `TR-IR-CURRENCY` | active | interest-rate risk factor must be a 3-letter currency code such as EUR, |
| `TR-LABEL-CHAR` | active | hedging-set label must not contain '/' or '#' |
| `TR-LAMBDA-IGNORED` | active | lambda input ignored under CRR - RTS lambda applied |
| `TR-LAMBDA-NUMBER` | active | lambda must be a number (found  ) |
| `TR-MATURED` | active | The trade maturity is on or before the reporting date. |
| `TR-MATURITY-MISSING` | active | maturity / end date missing |
| `TR-MTM-CURRENCY` | active | MtM |
| `TR-MTM-MISSING` | active | MtM missing (enter 0 if the trade has no value) |
| `TR-MTM-NUMBER` | active | MtM must be a number (found  ) |
| `TR-NATURE` | active | nature must be Standard, Basis or Volatility |
| `TR-NETTING-SET` | active | unknown netting set ' ' |
| `TR-NOTIONAL-CURRENCY` | active | notional |
| `TR-NOTIONAL-MISSING` | active | notional missing |
| `TR-NOTIONAL-NEGATIVE` | active | notional must not be negative (use Direction for the sign) |
| `TR-NOTIONAL-NUMBER` | active | notional must be a number (found  ) |
| `TR-OPTION-INPUTS` | active | option needs underlying price and strike |
| `TR-OPTION-TYPE` | active | option type must be Call or Put |
| `TR-OTHER-REGIME` | active | Other risks (OT) is unavailable under the effective regime. |
| `TR-REFERENCE-CHAR` | active | risk factor / reference must not contain '/' or '#' |
| `TR-REFERENCE-CONFLICT` | active | Reference ' ' is   here but |
| `TR-REFERENCE-MISSING` | active | risk factor / reference missing |
| `TR-RESET-MATURITY` | active | next reset date is after the maturity date |
| `TR-RESET-MISSING` | active | a settled-to-market trade (STM) needs its next reset date |
| `TR-RESET-NONSTM` | active | a next reset date is given but the settlement type is not STM |
| `TR-RESET-PAST` | active | The STM next reset date is on or before the reporting date. |
| `TR-SETTLEMENT` | active | settlement type must be N, STM or blank (found |
| `TR-SHIFTED-PRICE` | active | price and strike must be > 0 after lambda shift (BCBS: enter lambda per CRE52.40 FAQ) |
| `TR-START-REQUIRED` | active | an interest-rate or credit option needs the start date S of its underlying |
| `TR-TEXT` | active | must be text or blank (found an error value) |
| `TR-VOL-LABEL` | active | volatility trade needs a hedging-set label |
| `TR-VOLATILITY` | active | supervisory volatility missing for ' ' |
