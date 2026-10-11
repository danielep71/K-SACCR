# 📐 Methodology

This directory records **what SACCR implements and on what authority**: the
regulatory texts and their versions, assumptions, parameter sources, open
decisions, and the trace from each rule to code and tests. The numerical
test-case format is in [`TEST_CASES.md`](TEST_CASES.md).

> [!IMPORTANT]
> The engine implements CRR and Basel CRE52. Under
> [`TEST_CASES.md`](TEST_CASES.md), only Basel rules are validated, by the
> published CRE99 examples (T01–T14); no CRR rule is validated in v1.0.0
> ([decision 8](#assumptions-and-scope)). The [traceability](#traceability)
> table gives the status of each rule. The CRR and CRE52 versions in the
> source register and the supervisory parameters were checked against the
> official texts on 2026-10-09.

<a id="regulatory-basis"></a>

## ⚖️ Regulatory basis

Decided by the owner on 2026-10-06 in issue #8.

| Regime | Identifier | Role |
| --- | --- | --- |
| EU Capital Requirements Regulation, SA-CCR | `CRR` | **Baseline.** Default for every calculation and test case |
| Basel Framework, CRE52 | `BCBS` | Also supported; implemented as differences from the baseline |

A caller may choose the regime; when it does not, the facade uses `CRR`. Every
result reports the regime it was calculated under, and every expected result
names its regime. Where the two texts agree, the engine has one implementation
tagged with both rule identifiers. Where they differ, the difference is recorded
in [Regime differences](#regime-differences) before any code depends on it.

<a id="source-register"></a>

## 📚 Source register

Cite sources by ID in code `REFERENCE` sections, in the traceability table and
in test cases. Record the exact version date when a source is first relied on;
a later version is a new row, not an edit.

| ID | Source | Scope relied on | Version / date | Status |
| --- | --- | --- | --- | --- |
| `CRR` | Regulation (EU) No 575/2013, Part Three, Title II, Chapter 6, Section 3 (Articles 274–280f), as amended by Regulation (EU) 2019/876 and Regulation (EU) 2024/1623 | SA-CCR exposure value, replacement cost, PFE, add-ons | Consolidated text of 26.06.2026 (version 021.001), EUR-Lex | Articles 274(2)–(7), 275, 277, 277a, 278(3), 279a–279c, 280–280f and 285(2)–(5) checked 2026-10-09; Article 278(1)–(2) checked 2026-10-10 in the EUR-Lex PDF |
| `CRR-RTS` | Commission Delegated Regulation (EU) 2021/931, as amended by Delegated Regulation (EU) 2025/855 | Primary risk driver and risk-category mapping; supervisory delta and lambda shift of interest-rate and commodity options | Consolidated text of 25.05.2025 (version 001.001), EUR-Lex | Article 5 checked 2026-10-09; the mapping articles not yet |
| `EBA-QA-2023_6962` | EBA Single Rulebook Q&A 2023_6962 | Collateral C in the cap on the exposure value of a margined netting set (CRR Art. 274(3)): NICA | Final, published 15/03/2024 | Answer checked 2026-10-10 |
| `BCBS` | Basel Framework, chapter CRE52 | Basel SA-CCR | Effective 1 January 2023, last updated 5 June 2020 | Paragraph locators and the CRE52.72 table checked 2026-10-09 |
| `BCBS-279` | BCBS, *The standardised approach for measuring counterparty credit risk exposures* (2014) | Background; Annex 4a worked examples (printed pages 22–30) | March 2014, revised April 2014 | Example 4 checked 2026-10-09 |
| `CRE99` | Basel Framework, chapter CRE99, *Application guidance*: the SA-CCR worked examples | Published values of the `cre99-example-*` cases. Amounts in USD thousands; intermediate results unrounded, displayed results and final EAD rounded (CRE99.20); 250 business days a year (CRE99.28) | Effective 1 January 2023, last updated 27 March 2020 | Examples 1–5: the values used by the cases checked 2026-10-09 |

Reference values may only be taken from a source in this register. Add a row
before citing a new one.

Copies of the checked versions, with their addresses and SHA-256 hashes, are in
[`sources/`](sources/README.md).

<a id="assumptions-and-scope"></a>

## 🧭 Assumptions and scope

| Topic | Current position |
| --- | --- |
| Units | Amounts in the netting set's calculation currency, unscaled; rates and factors as decimals; periods in years. See [`VBA_HOUSE_STYLE.md`](../VBA_HOUSE_STYLE.md#units-and-domains) |
| Inputs | Trade-level data supplied by the caller; SACCR does not price trades or source market data |
| Supervisory parameters | Taken from the cited regime text and recorded in the parameter table below, never typed in code without a reference |
| Running test cases | Decided 2026-10-06 (decision 7, #44): VBA generated from the JSON by a checked tool; cases feed the engine through its input sheets. See [`TEST_CASES.md`](TEST_CASES.md#consumption) |
| Evidence for v1.0.0 | Decided 2026-10-09 (decision 8, #44): v1.0.0 validates rules only with the `published` values of BCBS 279 Annex 4 (CRE99), checked against the registered text. It has no `independent` values: the project has one maintainer, who also wrote the code under test. Every other case stays `illustrative`. Because the published examples are Basel examples, rules under `CRR` are implemented but not validated in v1.0.0, and the release says so. A later release cross-checks the results against a second source |
| Sources | Decided 2026-10-09 (decisions 1 and 2, #33): `CRR` is the consolidated CRR of 26.06.2026 (version 021.001), checked on 2026-10-09, with Delegated Regulation (EU) 2021/931 as amended by (EU) 2025/855 and EBA Q&A 2023_6962; `BCBS` is the Basel Framework CRE52 effective 1 January 2023, last updated 5 June 2020, checked on 2026-10-09 |
| Scope | Decided 2026-10-09 (decision 3, #33): the full SA-CCR only. Simplified SA-CCR (CRR Article 281), the original exposure method (Article 282), exposures to central counterparties, securities financing transactions, CVA and RWA are outside v1.0.0. They are planned as simplified SA-CCR in #107 (v1.2.0), the original exposure method in #111 (v1.3.0), CCP exposures in #112 (v1.4.0), SFTs in #113 (v1.5.0), CVA in #114 (v1.6.0) and RWA in #115 (v1.7.0). Inputs that ask for them are rejected; a limitation stated only in prose is not enough |
| Margined netting sets | Decided 2026-10-09 (decision 4, #33): in scope for v1.0.0, with VM, NICA, threshold, MTA and MPOR as inputs |
| Day count | Decided 2026-10-09 (decision 5, #33), confirmed after the CRR check: S, E, M and T in years of 365 calendar days (`DaysPerYear`); floors and MPOR in business days of 250 a year (`BusinessDaysPerYear`). Articles 279b and 279c say "the relevant business day convention"; this is the convention SACCR uses, without holiday calendars. A test case may set `DaysPerYear` to reproduce a published maturity |
| Currency | Decided 2026-10-09 (decision 6, #33): trade amounts are converted into the reporting currency with the FX table on Params, whose rates the user supplies; SACCR sources no market data. A test fixture holds amounts already in its calculation currency, whose rate is set to 1 |
| Interest-rate aggregation | Decided 2026-10-09 (decision 9, #33): under `BCBS` both CRE52.57 formulas are supported, chosen with `IRBucketOffset`: the bucket formula that recognises offsets across maturity buckets, and the sum of absolute bucket values. Under `CRR` only what Article 280a allows is supported, once its text is checked. Each supported formula needs a value test |
| Open treatments | Decided 2026-10-09 (decision 10, #33), revised after the CRR check: climatic conditions take 18% (Articles 277a(1)(e)(v), 280e(5)); there is no unrated sub-class, and the user enters CQS3, CQS5 where Article 128 applies, or the step mapped from an internal rating (Article 280c(5)(a)); other risks keep one hedging set per identical primary risk driver (Article 277a(1)(f)); inflation is interest rate, entered as a currency followed by `-INFL` (such as `EUR-INFL`) in its own hedging set (Articles 277(4)(a), 277a(1)); the optional reductions for sold options and sold credit protection (Article 274(5) and (7)) are not applied, which gives a higher exposure value and is permitted; several margin agreements in one netting set and one agreement over several netting sets (Articles 274(4), 275(3)) move to #39 |
| Collateral of the cap and one-way agreements | Decided 2026-10-10 (#39): the cap of a margined netting set uses C = NICA in both regimes, the collateral of the same netting set without a margin agreement (CRE52.2, CRE52.10; CRR Art. 274(3), 275(1), EBA Q&A 2023_6962). VM held on an unmargined netting set is rejected, because a counterparty that posts VM makes it margined (CRE52.2; CRR Art. 272(7)). VM posted on an unmargined netting set is the one-way agreement in favour of the counterparty: under `BCBS` it stays unmargined and C includes that VM with a negative sign (CRE52.10, footnote 2); under `CRR` it is rejected, because a one-way agreement is a margin agreement (Art. 272(7a)) and the netting set must be entered as margined |
| Client clearing | Decided 2026-10-10 (#39): the NettingSets column is the bank's role, N, CM or CLIENT. The 5-day MPOR floor applies to CM in both regimes (CRE54.12; CRR Art. 279c(1)) and to CLIENT under `CRR` only; a `BCBS` client keeps the 10-day floor (CRE54.8, CRE54.14). Eligibility of the arrangement is the user's to check |
| MPOR combinations | Decided 2026-10-10 (#39): after margin disputes the engine doubles F + N - 1, which is the `BCBS` rule (CRE52.50(2), CRE52.51(3)) and meets every reading of CRR Art. 285(4)–(5), being higher than the CRR minimum when remargining is not daily. A large or illiquid netting set takes 20 + N - 1 in both regimes: CRR Art. 285(3) and (5) give this; CRE52.51 sets 20 days without saying whether N - 1 is added, and adding it is not lower than either reading |
| Margin-agreement structures | Decided 2026-10-09 (#33) and implemented 2026-10-10 (#39): several margin agreements, or margined and unmargined trades, in one netting set (CRE52.74; CRR Art. 274(4)) and one agreement over several netting sets (CRE52.75; CRR Art. 275(3), 278(2)) are declared in the Margin agreements column as MIXED or SHARED, and the netting set is rejected |

<a id="support-matrix"></a>

## 🧾 Support matrix

What v1.0.0 calculates, what the user must enter in a supported form, and what
is outside it (#33). "Not supported" means the input sheets cannot express the
case; nothing is calculated for it. Whether a supported rule is also validated
is in the [traceability](#traceability) table.

| Area | Item | Status | How it is entered or treated | Reference |
| --- | --- | --- | --- | --- |
| Methods | Full SA-CCR, CRR (default) and Basel CRE52, chosen per netting set | Supported | Regime on Params, override per netting set | Decisions 1–3 |
| Methods | Simplified SA-CCR, original exposure method, CCP default-fund exposures, SFTs, CVA, RWA | Not supported | Cannot be entered; planned in #107 (simplified, v1.2.0), #111 (OEM, v1.3.0), #112 (CCP, v1.4.0), #113 (SFT, v1.5.0), #114 (CVA, v1.6.0), #115 (RWA, v1.7.0) | Decision 3 |
| Asset classes | Interest rate, foreign exchange, credit, equity, commodity | Supported | Asset class IR, FX, CR, EQ, CO | CRR Art. 277; CRE52.72 |
| Asset classes | Other risks | Supported under CRR; rejected under BCBS | Asset class OT; one hedging set per identical primary risk driver | CRR Art. 277(1)(f), 277a(1)(f), 280f |
| Asset classes | Inflation | Supported | IR with risk factor such as `EUR-INFL`, own hedging set per currency | CRR Art. 277(4)(a), 277a(1) |
| Asset classes | Climatic conditions | Supported under CRR | CO sub-class `CLIMATIC`, own hedging set, factor 18% | CRR Art. 277a(1)(e)(v), 280e(5) |
| Instruments | Linear derivatives (swaps, forwards, futures, single-name and index credit derivatives) | Supported | Instrument `Linear`; delta +1 long, -1 short | CRE52.39; CRR Art. 279a(1)(c) |
| Instruments | Options: European, American and Bermudan (the latest exercise date), Asian (the current value of the average) | Supported | Instrument `Option` with price, strike and expiry; CRR lambda for IR and commodity options | CRE52.40; CRR Art. 279a(1)(a); `CRR-RTS` Art. 5 |
| Instruments | Combinations of European option payoffs (collar, spread, straddle, strangle) and caps or floors | Entered by the user | One row per component option; a cap or floor as its caplets or floorlets | CRR Art. 274(6); CRE52.42, CRE52.43, CRE52.1 FAQ3 |
| Instruments | Digital options | Entered by the user under CRR; not supported under BCBS | CRR: the collar of a bought and a sold option at 0.95 K and 1.05 K, sized so that it reproduces the digital payoff outside the two strikes. BCBS also caps the digital's effective notional at its payoff divided by the supervisory factor, which the engine does not apply | CRR Art. 274(6); CRE52.42 |
| Instruments | CDO tranches and nth-to-default credit derivatives | Supported | Instrument `CDO` with attachment and detachment; nth-to-default as A = (n - 1)/k, D = n/k | CRE52.41; CRR Art. 279a(1)(b) |
| Instruments | Basis and volatility transactions | Supported | Nature `Basis` or `Volatility` with a hedging-set label | CRE52.46–.47; CRR Art. 277a(2), 280 |
| Instruments | Transactions with several material risk drivers | Entered by the user | One row per risk-category leg, mapped by the user. The trade's market value must be counted once: entered on one row with 0 on the others, or split so that the rows add up to it. The engine adds every row's market value to V and does not check this (#38) | CRR Art. 277(3); `CRR-RTS` |
| Instruments | Sold options and sold credit protection | Supported, without the optional reductions | Calculated as ordinary trades; the zero EAD and premium cap are not applied, which gives a higher EAD | CRR Art. 274(5), (7); CRE52.1 FAQ1–2 |
| Credit factors | Rated single names; quoted IG and SG indices | Supported | Sub-class AAA–CCC (BCBS) or CQS1–CQS6 (CRR); `IG_INDEX`, `SG_INDEX` | CRE52.72; CRR Art. 280c(5) |
| Credit factors | Unrated single names; unlisted multi-name positions | Entered by the user | CQS3, CQS5 where Art. 128 applies, or the mapped step; for an unlisted multi-name position, the step of its weighted factor | CRR Art. 280c(5) |
| Netting sets | Unmargined netting set | Supported | Margined = N; collateral in NICA | CRE52.10; CRR Art. 275(1) |
| Netting sets | Margined netting set with VM, NICA, threshold and MTA | Supported | Margined = Y; MPOR from floor, remargining frequency, large or illiquid flag, disputes and override | CRE52.18, .50–.53; CRR Art. 275(2), 279c, 285 |
| Netting sets | Cap of a margined netting set | Supported | Cap with C = NICA in both regimes | CRE52.2, CRE52.10; CRR Art. 274(3), 275(1) |
| Netting sets | One-way margin agreement where only the bank posts VM | Supported under BCBS; entered as margined under CRR | BCBS: Margined = N with the posted VM, included in C. CRR: Margined = Y, with the threshold of the agreement; VM on an unmargined CRR netting set is rejected | CRE52.2, CRE52.10 footnote 2; CRR Art. 272(7a), 275(2) |
| Netting sets | Several margin agreements, or margined and unmargined trades, in one netting set; one agreement over several netting sets | Not supported | Declared as MIXED or SHARED in the Margin agreements column; the netting set is rejected | CRR Art. 274(4), 275(3), 278(2); CRE52.74–.76 |
| Netting sets | Client clearing | Supported with a user declaration | Client clearing role CM applies the 5-business-day MPOR floor in both regimes, CLIENT under CRR only; a BCBS client keeps 10 days. Eligibility is the user's to check | CRE54.12, CRE54.14; CRR Art. 279c(1) |
| Collateral | Haircuts, eligibility, segregated initial margin | Entered by the user | Amounts after haircuts; segregated collateral posted by the bank excluded from NICA | CRE52.11, .17; CRR Art. 276 |
| Parameters | IR aggregation without offsets across buckets | Supported under BCBS only | `IRBucketOffset` = FALSE; a CRR netting set with IR trades is then INVALID | CRE52.57(5); CRR Art. 280a(3) |
| Parameters | Alpha override per netting set | Supported for what-if runs | Alpha column on NettingSets; regulatory runs use 1.4 (#41) | CRE52.1; CRR Art. 274(2) |
| Currency and dates | Reporting currency and conversion | Supported | FX table on Params, rates supplied by the user | Decision 6 |
| Currency and dates | Day count | Supported | 365 calendar days per year, 250 business days for floors and MPOR | Decision 5 |

<a id="open-decisions"></a>

## ❓ Open decisions

These are deliberately undecided. Each is settled by the owner in an issue and
moved into the tables above; none may be assumed by code or tests meanwhile.

None at present.

Decisions 1 to 6, 9 and 10 were settled on 2026-10-09 in #33 and are recorded
in [Assumptions and scope](#assumptions-and-scope).

<a id="regime-differences"></a>

## 🔀 Regime differences

Each confirmed difference between `CRR` and `BCBS` that affects a result gets a
row, with locators in both texts. No difference is asserted until both texts
have been compared for that rule.

| Topic | `CRR` (locator) | `BCBS` (locator) | Effect | Test cases |
| --- | --- | --- | --- | --- |
| Interest-rate aggregation across maturity buckets | Only the formula with offsets (Art. 280a(3)) | The formula with offsets, or the sum of absolute bucket values at the bank's choice (CRE52.57(5)) | `IRBucketOffset` = FALSE is valid only under `BCBS` | `TEST_Invariants` (relation) |
| Asset classes | Six, with other risks (Art. 277(1)(f), 280f: factor 8%) | Five (CRE52.72) | Other-risk trades are rejected under `BCBS` | T24, T25, T26 (illustrative) |
| Commodity hedging sets | Energy, metals, agricultural goods, other commodities and climatic conditions (Art. 277a(1)(e)); climatic conditions take 18% (Art. 280e(5)) | Energy, metals, agricultural and other commodities (CRE52.45, CRE52.70 step 2); CRE52.72 sets the factors by commodity type (electricity, oil/gas, metals, agricultural, other), not the hedging sets | Climatic trades form their own hedging set under `CRR` only | T27 (illustrative) |
| Single-name credit factor | By credit quality step 1–6 (Art. 280c(5), Table 3) | By rating AAA to CCC (CRE52.72) | Same values: 0.38%, 0.38%/0.42%, 0.54%, 1.06%, 1.6%, 6.0% | T28 (illustrative) |
| Supervisory delta of interest-rate and commodity options | Shifted by the regulatory lambda (`CRR-RTS` Art. 5) | CRE52.40 | Delta differs when a price or strike is near or below zero | T20–T23 (illustrative) |

The cap of a margined netting set is not recorded as a difference. Both texts
cap the EAD at that of the same netting set without margin (CRR Art. 274(3);
CRE52.2), and both define the unmargined collateral without variation margin:
CRR Art. 275(1) uses NICA (EBA Q&A 2023_6962), and CRE52.10 uses the net
collateral of the NICA methodology in CRE52.17, adding posted variation margin
with a negative sign only for a one-way margin agreement where the bank alone
posts (CRE52.10, footnote 2), which CRE52.2 treats as unmargined. Since #39
the engine's cap uses NICA in both regimes; before, its `BCBS` cap included
variation margin (C = VM + NICA).

<a id="parameters"></a>

## 🔢 Supervisory parameters

Values checked on 2026-10-09 against the consolidated CRR of 26.06.2026,
`CRR-RTS` Article 5 and the CRE52.72 table; each matches the template's Params
sheet. Rows marked `CRR` are as printed in the CRR; rows marked `BCBS` are the
CRE52.72 values, which equal the CRR values wherever both texts give one.

| Parameter | Value | Regime | Source locator | Used by |
| --- | --- | --- | --- | --- |
| Alpha | 1.4 | CRR | Art. 274(2) | `Alpha` |
| Multiplier floor | 5% | CRR | Art. 278(3) | `MultiplierFloor` |
| Supervisory discount rate R | 5% | CRR | Art. 279b(1)(a) | `SACCR_SupervisoryDuration` |
| Supervisory duration floor | 10 / OneBusinessYear | CRR | Art. 279b(1)(a) | `SDFloorBD` |
| Maturity floor, unmargined | 10 / OneBusinessYear | CRR | Art. 279c(1)(a) | `MinMaturityBD` |
| Maturity factor, margined | 3/2 x sqrt(MPOR / OneBusinessYear) | CRR | Art. 279c(1)(b) | `SACCR_MaturityFactor` |
| MPOR floor | 10 business days; 5 for transactions between a client and a clearing member | CRR | Art. 285(2)(b); Art. 279c(1)(b) | `MPORFloorBilateral`, `MPORFloorCleared` |
| MPOR floor, over 5 000 trades or illiquid collateral or a hard-to-replace OTC derivative | 20 business days | CRR | Art. 285(3) | `MPORFloorLarge` |
| MPOR with margin disputes | at least double | CRR | Art. 285(4) | `CORE_Engine.LoadNettingSets` |
| MPOR with remargining every N days | F + N - 1 | CRR | Art. 285(5) | `CORE_Engine.LoadNettingSets` |
| Hedging-set coefficient | 1; 5 for volatility; 0.5 for basis | CRR | Art. 280 | `VolatilityFactor`, `BasisFactor` |
| Interest rate: factor; option volatility | 0.5%; 50% | CRR | Art. 280a(2); `CRR-RTS` Art. 5(3) | Params row `IR` |
| Interest rate: buckets; cross terms | <= 1, 1 to 5, > 5 years; 1.4, 1.4, 0.6 | CRR | Art. 280a(3), Table 2 | `IRCorrBucket12`, `IRCorrBucket23`, `IRCorrBucket13` |
| Foreign exchange: factor; option volatility | 4%; 15% | CRR | Art. 280b(2); Art. 279a Table 1 | Params row `FX` |
| Credit, single name: factor by credit quality step 1 to 6 | 0.38%, 0.42%, 0.54%, 1.06%, 1.6%, 6.0% | CRR | Art. 280c(5), Table 3 | Params rows `CR_CQS1` to `CR_CQS6` |
| Credit, unrated single name | 0.54% under the standardised approach, 1.6% where Art. 128 applies; IRB banks map the internal rating | CRR | Art. 280c(5)(a) | Entered as `CQS3`, `CQS5` or the mapped step |
| Credit, quoted index: factor | 0.38% investment grade; 1.06% non-investment grade | CRR | Art. 280c(5), Table 4 | Params rows `CR_IG_INDEX`, `CR_SG_INDEX` |
| Credit: correlation; option volatility | 50% single name, 80% index; 100% single name, 80% index | CRR | Art. 280c(3); Art. 279a Table 1 | Params credit rows |
| Equity: factor; correlation; option volatility | 32%, 50%, 120% single name; 20%, 80%, 75% index | CRR | Art. 280d(3)–(4); Art. 279a Table 1 | Params rows `EQ_SINGLE`, `EQ_INDEX` |
| Commodity: factor | 18%; 40% for electricity | CRR | Art. 280e(5) | Params `CO_` rows |
| Commodity, climatic conditions: factor | 18% (a commodity hedging set other than electricity) | CRR | Art. 277a(1)(e)(v); Art. 280e(5) | Params row `CO_CLIMATIC` |
| Commodity: correlation; option volatility | 40%; 70%, 150% for electricity | CRR | Art. 280e(4); `CRR-RTS` Art. 5(3) | Params `CO_` rows |
| Other risks: factor; option volatility | 8%; 150% | CRR | Art. 280f(2); Art. 279a Table 1 | Params row `OT` |
| Lambda thresholds | 0.10% for interest rate; 0.1 for commodity | CRR | `CRR-RTS` Art. 5(2) | `LambdaThresholdIR`, `LambdaThresholdCO` |
| Interest rate: factor; option volatility | 0.50%; 50% | BCBS | CRE52.72 | Params row `IR` |
| Foreign exchange: factor; option volatility | 4.0%; 15% | BCBS | CRE52.72 | Params row `FX` |
| Credit, single name by rating AAA, AA, A, BBB, BB, B, CCC | 0.38%, 0.38%, 0.42%, 0.54%, 1.06%, 1.6%, 6.0%; correlation 50%; volatility 100% | BCBS | CRE52.72 | Params rows `CR_AAA` to `CR_CCC` |
| Credit, index IG, SG | 0.38%, 1.06%; correlation 80%; volatility 80% | BCBS | CRE52.72 | Params rows `CR_IG_INDEX`, `CR_SG_INDEX` |
| Equity, single name; index | 32%, correlation 50%, volatility 120%; 20%, 80%, 75% | BCBS | CRE52.72 | Params rows `EQ_SINGLE`, `EQ_INDEX` |
| Commodity: electricity; oil/gas, metals, agricultural, other | 40%; 18%; correlation 40%; volatility 150%; 70% | BCBS | CRE52.72 | Params `CO_` rows |

How the points found in the CRR text are handled (#33):

- Article 280a prints only the formula with offsets across maturity buckets.
  With `IRBucketOffset` = FALSE, a `CRR` netting set with interest-rate trades
  is `INVALID` and its EAD withheld; `BCBS` netting sets may use either
  formula.
- Inflation trades are entered as `EUR-INFL` and similar, each in its own
  hedging set.
- The optional reductions of Article 274(5) and (7) are not applied
  ([decision 10](#assumptions-and-scope)). As printed, Article 274(5) allows a
  zero exposure value only for a netting set that consists solely of sold
  options, whose current market value is negative at all times, whose
  premiums were all received upfront, and that is not subject to a margin
  agreement. Article 274(7) allows a credit derivative long the underlying
  (protection sold) to be capped at the outstanding unpaid premium only if it
  is its own netting set, not subject to a margin agreement. CRE52.1 FAQ1 and
  FAQ2 give the Basel counterparts for options and protection outside netting
  and margin agreements. Implementing either would need these conditions as
  inputs.
- Article 274(4) and 275(3) cannot be expressed in the input sheets and are
  handled in #39.
- Article 280c(5)(b)(ii) (an unlisted multi-name credit position takes the
  notional-weighted factor of its constituents) is not automated: the user
  enters the resulting factor's credit quality step or a quoted index.

<a id="traceability"></a>

## 🔗 Traceability

One row per implemented rule. Rule IDs are the regime identifier and the
article or paragraph, for example `CRR.275.1` or `CRE52.<paragraph>`; `CRE52` alone
marks a paragraph not yet located. A rule shared by
both regimes lists both IDs.

Article numbers were checked against the CRR on 2026-10-09. CRE52 paragraph
numbers were checked against the CRE52 contents for EAD and alpha (.1), the cap
(.2), RC (.10, .18), PFE (.20), the multiplier (.22–.23), the aggregate add-on
(.25), supervisory duration (.34), adjusted notional (.35, .36), delta (.39,
.40, .41), basis and volatility (.46, .47, .73), maturity factor (.48–.49,
.52–.53), MPOR (.50, .51), interest rate (.56, .57), FX (.58, .59), credit
(.60, .61, .64), equity (.65, .66, .68), commodity (.69, .70) and the
parameters (.72).
CRE52.64 is about credit and CRE52.68 about equity factors; the code cited them
for equity and commodity until #33 corrected it. CRE52.65 and CRE52.69
introduce the equity and commodity offsets.
Test IDs `T01`–`T28` are listed in
[`tests/README.md`](../../tests/README.md#ported-prototype-catalogue). Status
uses four terms:

- **Validated**: a `published` output whose rule is this one passes in Excel
  for that regime. Only these count for the
  [completeness policy](TEST_CASES.md#completeness-policy).
- **Exercised**: no published output names the rule, but a published value
  that passes depends on it, so an error in it would very probably fail that
  value. Not counted as validated.
- **Not validated**: implemented; any cases are illustrative.
- **Not implemented**: listed below the table.

| Rule ID | Requirement | VBA procedure | Test cases | Status |
| --- | --- | --- | --- | --- |
| `CRE52.1`, `CRR.274.2` | EAD = alpha x (RC + PFE), alpha 1.4 | `CORE_Engine.ComputeNettingSets` | Published: T04, T06, T07, T10, T14. Illustrative: T15, T19, T25. Relations: `TEST_Invariants` | BCBS validated; CRR not validated |
| `CRE52.2`, `CRR.274.3` | EAD of a margined netting set capped at its unmargined EAD; the cap uses C = NICA in both regimes (CRE52.10; `EBA-QA-2023_6962`) | `CORE_Engine.ComputeNettingSets` | Illustrative: T15, T16, T17, T18 and the T16 cap flag. Relations: `TEST_Invariants` | Not validated: no published example makes the cap bind |
| `CRE52.10`, `CRR.275.1` | RC of an unmargined netting set = max(V - C, 0); under BCBS, C includes VM posted under a one-way agreement (CRE52.10 footnote 2); VM held, or VM posted under CRR, is rejected | `SACCR_ReplacementCost`, `CORE_Engine.LoadNettingSets`, `CORE_Engine.ComputeNettingSets` | Published: T08. Illustrative: `illustrative-ir-swap-unmargined`, `TEST_InputValidation` (one-way agreement) | BCBS validated; CRR not validated |
| `CRE52.18`, `CRR.275.2` | RC of a margined netting set = max(V - C, TH + MTA - NICA, 0) | `SACCR_ReplacementCost`, `CORE_Engine.ComputeNettingSets` | Published EAD T14 (Example 5) | BCBS exercised; CRR not validated |
| `CRE52.20`, `CRR.278.1` | PFE = multiplier x aggregate add-on | `CORE_Engine.ComputeNettingSets` | Published EADs T04, T06, T07, T10, T14. Relations: `TEST_Invariants` | BCBS exercised; CRR not validated |
| `CRE52.23`, `CRR.278.3` | Multiplier, floor 5% | `SACCR_Multiplier`, `CORE_Engine.ComputeNettingSets` | Published: T05, T13. Relations: `TEST_Invariants` | BCBS validated; CRR not validated |
| `CRE52.25`, `CRR.278.1` | Aggregate add-on = sum of the asset-class add-ons; under CRR only the risk categories to which at least one trade of the netting set is mapped | `CORE_Engine.ComputeNettingSets` | Published: T09, T12. Relations: `TEST_Invariants` | BCBS validated; CRR not validated |
| `CRE52.34`, `CRR.279b` | Supervisory duration, floored at 10 business days, and adjusted notional of interest-rate and credit trades | `SACCR_SupervisoryDuration`, `CORE_Engine.ProcessTrades` | Published: T01 | BCBS validated; CRR not validated |
| `CRE52.40`, `CRR.279a` | Supervisory delta of options | `SACCR_OptionDelta`, `CORE_Engine.ProcessTrades` | Published: T02. Illustrative: T21 | BCBS validated; CRR not validated |
| `CRR-RTS.5` | CRR lambda shift for interest-rate and commodity options with negative or low prices | `SACCR_LambdaCRR`, `CORE_Engine.ProcessTrades` | Illustrative: T20, T23; T22 (rejected under BCBS) | Not validated |
| `CRE52.41` | Supervisory delta of CDO tranches | `SACCR_CDODelta`, `CORE_Engine.ProcessTrades` | `TEST_Harness` (input domain only) | Not validated |
| `CRE52.46`, `CRE52.73`, `CRR.277a`, `CRR.280` | Basis transactions: own hedging set per basis, factor x 0.5 | `CORE_Engine.ProcessTrades` | `TEST_Aggregation` (relation: a basis swap adds half a standard swap's add-on) | Not validated |
| `CRE52.47`, `CRE52.73`, `CRR.277a`, `CRR.280` | Volatility transactions: own hedging sets built as in CRE52.45, factor x 5 | `CORE_Engine.ProcessTrades` | None: no case exercises a volatility transaction | Not validated |
| `CRE52.48`, `CRE52.49`, `CRR.279c` | Maturity factor of an unmargined trade, sqrt(min(M, 1)), M floored at 10 business days | `SACCR_MaturityFactor`, `CORE_Engine.ProcessTrades` | Published EAD T07 (9-month trade) | BCBS exercised; CRR not validated |
| `CRE52.50`, `CRE52.52`, `CRE52.53`, `CRR.279c`, `CRR.285` | MPOR floor F plus the remargining period N less one: F is 10 business days for a bilateral netting set and 5 for client clearing: the clearing member facing its client in both regimes (CRE54.12, `CRR.279c.1`), the client facing its clearing member under CRR only, a BCBS client keeping 10 (CRE54.8, CRE54.14); maturity factor of a margined netting set, 1.5 x sqrt(MPOR / 250) | `CORE_Engine.LoadNettingSets`, `SACCR_MaturityFactor` | Published: T11 (weekly remargining, 14 days); EAD T14. Illustrative: `TEST_InputValidation` (MPOR of each clearing role) | BCBS validated (MPOR with remargining), exercised (maturity factor); CRR not validated |
| `CRE52.51`, `CRR.285` | MPOR exceptions: 20 business days for over 5,000 trades or illiquid collateral or a hard-to-replace derivative; the floor doubled after margin disputes | `CORE_Engine.LoadNettingSets` | Illustrative: `TEST_InputValidation` (disputes, 28 days, and a large netting set, 24 days, with weekly remargining) | Not validated. The engine doubles F + N - 1 after disputes. Under BCBS this is the text: the floor for non-daily remargining is nine business days plus N (CRE52.50), and CRE52.51(3) doubles that floor. Under CRR, Art. 285(4) doubles F and Art. 285(5) adds N - 1 to F, so the engine's MPOR is higher than the CRR text requires when remargining is not daily. A large or illiquid set takes 20 + N - 1 in both regimes. Both decided 2026-10-10 ([MPOR combinations](#assumptions-and-scope)) |
| `CRE52.57`, `CRR.280a` | Interest rate: hedging set per currency (CRE52.57(2)), maturity buckets (.57(3)), bucket formula or, under `BCBS` only, sum of absolutes (.57(5)) | `CORE_Engine.ProcessTrades`, `SACCR_IREffectiveNotional`, `CORE_Engine.ComputeHedgingSets` | Published: T03. Relations: `TEST_Invariants`, `TEST_Aggregation` | BCBS validated; CRR not validated |
| `CRE52.58`, `CRE52.59`, `CRR.280b` | Foreign exchange: hedging set per currency pair, add-on = factor x abs(effective notional) | `CORE_Engine.ProcessTrades`, `CORE_Engine.ComputeHedgingSets` | None | Not validated |
| `CRE52.60`, `CRE52.61`, `CRE52.64`, `CRR.280c` | Credit: entity buckets, single-factor aggregation, factor by rating (BCBS) or credit quality step (CRR) | `CORE_Engine.ProcessTrades`, `SACCR_FactorAggregation`, `CORE_Engine.ComputeHedgingSets` | Published EAD T06. Illustrative: T28. Relations: `TEST_Aggregation` | BCBS exercised; CRR not validated |
| `CRE52.65`, `CRE52.66`, `CRE52.68`, `CRR.280d` | Equity: entity buckets, single-factor aggregation | `CORE_Engine.ProcessTrades`, `SACCR_FactorAggregation`, `CORE_Engine.ComputeHedgingSets` | None | Not validated |
| `CRE52.69`, `CRE52.70`, `CRR.280e` | Commodity: hedging sets energy, metals, agricultural and other, commodity-type buckets, single-factor aggregation | `CORE_Engine.ProcessTrades`, `SACCR_FactorAggregation`, `CORE_Engine.ComputeHedgingSets` | Published EADs T07, T12 | BCBS exercised; CRR not validated |
| `CRR.277a`, `CRR.280e` | CRR climatic-conditions commodity hedging set, factor 18% | `CORE_Engine.ProcessTrades` | Illustrative: T27 | Not validated. The 18% factor was checked against Art. 280e(5) on 2026-10-09; that is a source check, not a numerical validation |
| `CRR.277.1`, `CRR.280f` | CRR other-risks asset class, factor 8%; rejected under BCBS | `CORE_Engine.ProcessTrades`, `CORE_Engine.ComputeHedgingSets` | Illustrative: T24, T25, T26 | Not validated |
| `CRE52.72` | Supervisory factors, correlations and option volatilities on Params | `CORE_Engine.LoadSFTable` | Published values T01–T14, for the factors their trades use | BCBS exercised for those factors; CRR not validated |

Not implemented, as recorded on the template's Basel_vs_CRR sheet; where the
sheet says so, the user enters such trades in a form the engine supports: several margin agreements in one netting set
(CRE52.74, `CRR.274.4`) and one margin agreement covering several netting sets
(`CRR.275.3`, `CRR.278.2`), which the user declares so that the netting set is
rejected; the special treatment of sold options and sold credit
protection (`CRR.274.5`, `CRR.274.7`); decomposition of option combinations
(`CRR.274.6`); automated mapping of risk drivers (`CRR.277.2`, `CRR.277.3`);
simplified SA-CCR (`CRR.281`) and the original exposure method (`CRR.282`).
Collateral is an input after haircuts, and trade amounts are converted with
the FX table on Params ([decision 6](#assumptions-and-scope)); neither is
validated.

A rule is **validated** only when at least one case with a `published` or
`independent` reference passes for it in each regime it claims; see the
[completeness policy](TEST_CASES.md#completeness-policy). Illustrative cases
never count.
