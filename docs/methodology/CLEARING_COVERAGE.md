# 🧭 Clearing coverage

What K-SACCR does today for client clearing, central counterparties (CCPs) and
listed derivatives, which rules govern each arrangement, and what is still to
be defined (#358). It is the specification for the clearing work of v1.0.2:
client-clearing verification (#359), settlement and reset inputs (#360), CCP
exposures and default funds (#112), and the closing evidence (#361). It changes
no calculation.

The inventory was taken at commit `d16eeba` of `release/1.0.2`, whose engine,
workbook and tests are those of release 1.0.1. Rules are cited by the IDs of the
[source register](README.md#source-register); every citation below was checked
against the registered texts.

## 🏷️ Coverage classes

| Class | Meaning |
| --- | --- |
| Implemented | The engine applies the rule, from an input it reads and checks |
| Exercised | A test runs the rule and checks its result; its expected value is `illustrative` |
| Validated | A `published` or `independent` value checks the result ([`TEST_CASES.md`](TEST_CASES.md)) |
| User's | The engine calculates what the user enters; whether the input meets the rule is the user's to check |
| Not supported | Nothing is calculated for it. Where it cannot be told apart from a supported case, this page says so |

No clearing rule is validated. Every published value of the register (CRE99,
Examples 1–5) is a bilateral netting set, and no `independent` value exists
yet. The clearing tests are all `illustrative`, including the expected values
of the `illustrative-clearing-*-listed` cases, computed by
[`tests/derivations/clearing_listed.py`](../../tests/derivations/clearing_listed.py),
a separate implementation that no reviewer other than the code's author has
checked.

## 🔀 Arrangements

The bank's role is the NettingSets column *Client clearing role*: `N` (blank),
`CM`, `CLIENT`, `CCP` or `MULTI`. `CCP` declares a netting set facing a CCP and
`MULTI` a multi-level client structure; both are rejected ([decisions 5 and
6](#decisions)). `CM` and `CLIENT` act only on a margined netting set, through
the MPOR floor, and each such netting set gets an INFO line on Checks naming
its role, floor and rule ([decision 7](#decisions)); an unmargined netting set takes its maturity factor from each trade's
remaining maturity whatever the role.

| Arrangement | Entered as | MPOR floor, CRR | MPOR floor, BCBS | Coverage |
| --- | --- | --- | --- | --- |
| Bilateral, no clearing | Role `N` | 10 business days (Art. 285(2)(b)) | 10 business days with daily margining, 9 + N otherwise (CRE52.50(1)–(2)) | Implemented; exercised; validated under BCBS through CRE99 (MPOR with remargining) |
| Clearing member facing its client | Role `CM` | 5 business days (Art. 279c(1), last subparagraph; Art. 304(3)(a)) | 5 business days (CRE54.12) | Implemented; exercised, also with a listed future and option (`illustrative-clearing-member-listed`, both regimes) |
| Bank as client facing its clearing member | Role `CLIENT` | 5 business days (Art. 279c(1), last subparagraph) | 10 business days: the CCP treatment where CRE54.15 is met (CRE54.14; 10 days for OTC derivatives, CRE54.8(2)), with a 4% risk weight instead of 2% where only the joint-default protection is missing (CRE54.16); a bilateral trade otherwise (CRE54.17; CRE52.50) | Implemented; exercised, also with a listed future and option and weekly remargining: MPOR 9 under CRR, 14 under BCBS (`illustrative-clearing-client-listed`). The conditions of CRE54.15 and CRR Art. 305(2)–(3) are the user's. They set the risk weight and, under BCBS, whether the 5,000-trade floor applies (CRE54.8(1), not implemented); the MPOR floor stays as shown |
| Bank facing a CCP: a clearing member's trade exposure, for itself or for clients, or a client's exposure to the CCP guaranteed by its clearing member (CRE54.14(2)) | Role `CCP`, which rejects the netting set | At least 10 business days (Art. 304(3)(b)); the 20-day floor for more than 5,000 trades does not apply to trade exposures (Art. 285(3)(a)); VM retained by the CCP without protection of the member's collateral: the lesser of one year and the remaining maturity, floored at 10 business days (Art. 304(3)(d)) | 10 days for OTC derivatives (CRE54.8(2)); the 5,000-trade floor does not apply (CRE54.8(1)); VM retained by the CCP: as CRR (CRE54.8(3)) | Not supported, as in v1.0.0 and v1.0.1. Declared with role `CCP`, the netting set is `INVALID` (`netting-set-clearing-role-ccp`, `netting-set-clearing-role-ccp-margined`). A CCP entered as the counterparty of a role-`N` netting set still cannot be told apart from a bilateral one and gets the bilateral floors (#112) |
| Multi-level client structure: a higher-level client facing a lower-level client | Role `MULTI`, which rejects the netting set | The CRR five-day floor is for transactions "between a client and a clearing member" (Art. 279c(1)); the multi-level treatment is in Art. 300(6)–(9), 304(7), 305(4) | Treated as the clearing member–client leg (CRE54.2, CRE54.13, CRE54.14(3)) | Not supported. Declared with role `MULTI`, the netting set is `INVALID: multi-level client structure, not supported` (`netting-set-clearing-role-multi-level`); entered as `CM` or `CLIENT` instead, it cannot be detected |

The clearing-member–client leg of an exchange-traded derivative conducted under
a bilateral agreement is capitalised as an OTC derivative by both parties
(CRE54.2). In K-SACCR that is the `CM` or `CLIENT` netting set above.

## 🏦 Qualifying and non-qualifying CCPs

A QCCP is a CCP authorised under Article 14 or recognised under Article 25 of
Regulation (EU) No 648/2012 (`CRR-CCP`, Art. 4(1)(88)). The Basel definition is
in CRE50, which is not in the register yet; #112 adds it before relying on it.
The status changes risk weights and capital, not the SA-CCR exposure value.
K-SACCR calculates exposure values only, so none of the following is
calculated.

| Item | QCCP | Non-qualifying CCP | K-SACCR |
| --- | --- | --- | --- |
| Clearing member's trade exposure | 2% risk weight (Art. 306(1)(a); CRE54.7) | Standardised risk weight of the counterparty (Art. 306(1)(b); CRE54.41) | Not supported (#112) |
| Client's trade exposure meeting the segregation and portability conditions | 2%; 4% when only the joint-default protection is missing (Art. 305(2)–(3); CRE54.14–.16) | Not available: Art. 305(2)(d) requires a QCCP | Not supported |
| Clearing member's trade exposure to the CCP for a client trade | Exposure value zero when the member does not reimburse the client for losses on the CCP's default (Art. 306(1)(c)); otherwise as its own trade exposure (Art. 306(1)(d); CRE54.7) | Under CRR the same: zero without reimbursement (Art. 306(1)(c), not limited to QCCPs), otherwise the non-qualifying treatment of its own trade exposure (Art. 306(1)(b), (d)); CRE54.41 does not distinguish the two cases | Not supported |
| Pre-funded default-fund contribution | K = max(K_CCP · DF_i / (DF_CCP + DF_CM); 8% · 2% · DF_i) (Art. 308; CRE54.25–.39) | K = DF + UC (Art. 309); 1250% (CRE54.42) | Not supported (#112) |
| Unfunded default-fund contribution | 0% (Art. 310) | In K = DF + UC (Art. 309) | Not supported |
| Cap | Trade exposure and default-fund requirements together capped at the non-qualifying requirement (Art. 303(2); CRE54.40) | — | Not supported |
| CCP ceasing to qualify | Non-qualifying treatment within three months (Art. 311; CRE54.6) | — | Not supported |

## 📦 Products

K-SACCR has no exchange or listing input. The treatment follows the netting
set's role and margining, never the fact that a contract is listed: listing
does not make a trade cleared, and a cleared trade is not necessarily listed.

| Product | Entered as | Coverage |
| --- | --- | --- |
| Listed future | Instrument `Linear`, with M, and S and E for interest-rate and credit futures | Implemented; exercised as a bilateral linear trade, and in the `CM` and `CLIENT` netting sets of `illustrative-clearing-member-listed` and `illustrative-clearing-client-listed` (equity index future) |
| Ordinary listed option, on a security, an index or a future | Instrument `Option`, with price, strike and expiry T; for an option on an interest-rate future, S and E of the underlying | Implemented; exercised as a bilateral option, and in the same clearing cases (equity index call; its supervisory delta is checked) |
| Daily settled (settled-to-market) contract | Settlement type `STM` and its next reset date, in an unmargined netting set (#360) | Implemented; the qualification is the user's: see [Settlement and reset](#settlement-and-reset). Exercised by `illustrative-settled-to-market-daily` and `illustrative-settled-to-market-periodic` |
| Exotic options | No separate input | Relevant to the clearing member's relief from the 20-day floor (Art. 304(3)(c); CRE54.8(1)), which is not implemented |
| SFTs, long settlement transactions, cash transactions cleared through a CCP | Cannot be entered | Not supported. SFTs are #113; cash transactions are outside the CCP rules (Art. 301(1); CRE54.1) |

## 🛡️ Margin, collateral and segregation

| Item | Rule | K-SACCR |
| --- | --- | --- |
| Variation margin under a margin agreement | Margined netting set (CRE52.18; Art. 275(2)) | Implemented: Margined = Y, VM, threshold, MTA |
| Initial margin received | Part of NICA (CRE52.15–.17; Art. 272(12a)) | User's: entered in NICA after haircuts |
| Collateral posted and held bankruptcy remote | Exposure value zero for that collateral (Art. 306(2); CRE54.21); not part of NICA (CRE52.17) | User's: left out of NICA |
| Collateral posted and not held bankruptcy remote | Counted in NICA, with a negative sign (CRE52.17; CRE54.23) | User's: entered in NICA with a negative sign |
| Client collateral passed on to the CCP by the clearing member | Recognised against both the CCP and the client leg (Art. 304(7); CRE54.13) | User's: entered in the NICA of the `CM` netting set |
| Initial margin used by the CCP to mutualise losses | A default-fund contribution, not initial margin (Art. 301(2)(c)) | Must not be entered in NICA; default funds are not supported |
| Excess collateral the CCP or the member can stop being withdrawn | Initial margin (Art. 301(2)(b)) | User's: in NICA |
| Segregation and portability of client positions | Condition for the client's 2% or 4% risk weight (Art. 305(2)–(3); CRE54.15–.16) | Not checked; affects risk weights only |

## ⏱️ MPOR

| Condition | CRR | BCBS | K-SACCR |
| --- | --- | --- | --- |
| Floor by role | 10; 5 for `CM` and `CLIENT` (Art. 279c(1), 285(2)(b)) | 10; 5 for `CM` (CRE52.50, CRE54.12) | Implemented; exercised by `netting-set-clearing-member-crr`, `-bcbs`, `netting-set-clearing-client-crr`, `-bcbs` and the two `illustrative-clearing-*-listed` cases |
| Remargining every N business days | F + N − 1 (Art. 285(5)) | 9 + N for bilateral non-daily (CRE52.50(2)) | Implemented: F + N − 1 for every role; exercised for the client with N = 5 (`illustrative-clearing-client-listed`) |
| More than 5,000 trades | 20 days, except for trade exposures (Art. 285(3)(a)); a clearing member may disregard it towards its clients, without illiquid collateral, disputes or exotic options (Art. 304(3)(c)) | 20 days for netting sets not with a CCP (CRE52.51(1)); not for trade exposures to a CCP (CRE54.8(1)) | Implemented conservatively, a recorded limitation ([decision 3](#decisions)): one *large or illiquid* flag sets 20 days for every role, so the relief cannot be applied; exercised for the clearing member (`netting-set-clearing-member-large-crr`, `-bcbs`). A set over 5,000 trades with the flag at N is VALID with a warning (`netting-set-over-5000-trades`) |
| Illiquid collateral or a hard-to-replace OTC derivative | 20 days (Art. 285(3)(b)) | 20 days (CRE52.51(2)) | Implemented; exercised (`netting-set-mpor-large-weekly-bcbs`) |
| Margin disputes | Doubled (Art. 285(4)) | Floor doubled (CRE52.51(3)) | Implemented: F + N − 1 doubled, as recorded in [Assumptions](README.md#assumptions-and-scope); exercised for the client, 10 under CRR and 20 under BCBS (`netting-set-clearing-client-disputes-crr`, `-bcbs`) |
| VM retained by the CCP, member collateral unprotected | Lesser of one year and remaining maturity, floor 10 (Art. 304(3)(d)) | Same (CRE54.8(3)) | Not supported, and not detectable; inputs required by [decision 2](#decisions) (#112) |

<a id="settlement-and-reset"></a>

## 🔁 Settlement and reset

Settlement and margining differ. A contract whose outstanding exposure is
settled on set dates, with terms reset so that its value is zero, has as
remaining maturity the time to the next reset date (CRE52.37(5); CRR Art.
279c(2)); the settlement is a contractual payment of the same transaction
(Art. 279b(1)(a)). For a daily settled future the Basel FAQ to CRE52.31–.32
treats the trade as unmargined, with M at its floor of 10 business days, and
keeps E, so the interest-rate maturity bucket does not change. A contract
that is daily margined but not settled (collateralised-to-market) is a
margined netting set with an MPOR.

| Item | K-SACCR | Note |
| --- | --- | --- |
| Settled-to-market linear contract | Settlement type `STM` and the next reset date, in an unmargined netting set: M is the time to the reset, floored at 10 business days in the maturity factor; E stays the contractual end | Implemented (#360); the [input contract](#settled-to-market-input-contract) below |
| Settled-to-market option | The same, with the expiry T kept for the supervisory delta, so the expiry may be after the reset | Implemented (#360). Before, the next reset date had to be entered as M, and an option expiring after it was rejected |
| Collateralised-to-market contract | Margined = Y with the MPOR rules above | — |
| Exchange calendar | None: business days are 250 a year without holidays | Listing does not define a calendar; the user enters the reset date |
| Past or inconsistent dates | M on or before the reporting date is rejected; M after E is not checked. A reset date on or before the reporting date or after M is rejected | [Input contract](#settled-to-market-input-contract), rules 3 and 4 |

<a id="settled-to-market-input-contract"></a>

### Settled-to-market input contract (#360)

Specified from [decision 4](#decisions) and implemented in #360.

**Columns.** Trades has two columns immediately before *Comment*, after
#38's *Economic trade ID* and *MtM leg*. Comment moved from AB to AD:

| Column | Header | Content |
| --- | --- | --- |
| Z | Economic trade ID | Unchanged (#38) |
| AA | MtM leg (Y/N) | Unchanged (#38) |
| AB | Settlement type (N/STM) | Blank or `N`: ordinary treatment. `STM`: settled-to-market; lower case is read as upper |
| AC | Next reset date | A date. Required for `STM`; must be blank for blank or `N` |
| AD | Comment | Moved from AB; still not read |

**Rules.** Each rule holds under both regimes; the texts are the same in
substance (CRE52.37(5), CRE52.48; CRR Art. 279b(1)(a), 279c(1)(a), 279c(2)).

| # | Rule | Inputs that break it |
| --- | --- | --- |
| 1 | `STM` declares that, on specified dates, the contract settles its outstanding exposure and its terms reset so that its value is zero. The settlement is a contractual payment of the same transaction (Art. 279b(1)(a)). Daily variation margin alone does not qualify: a contract margined daily but not settled is collateralised-to-market and belongs in a margined netting set with an MPOR (Basel FAQ to CRE52.31–.32). `STM` is the user's attestation of these contractual facts: no input can prove them, so the engine validates the declaration and the dates (rules 2–7), not the contract | — |
| 2 | For an `STM` trade, the remaining maturity used in the unmargined maturity factor is the time from the reporting date to the next reset date, in years on the `DaysPerYear` basis. It keeps the floor of 10 business days (`MinMaturityBD` / `BusinessDaysPerYear`; CRE52.48; Art. 279c(1)(a)) | — |
| 3 | The other dates keep their meaning. M stays the contractual maturity and keeps its checks, including that an option's expiry T is not after M. E and S keep the supervisory duration and the interest-rate maturity bucket (Basel FAQ to CRE52.31–.32). T keeps the supervisory delta. The reset date is compared with the contractual maturity as entered, before the reset replaces it in the maturity factor; a reset on the maturity date is accepted, since the final reset may fall on it | Next reset date after M (or, where M is blank, after the E or T used in its place) |
| 4 | The MtM is the value at the reporting date after any settlement due that day, so the next reset date is the first reset after the reporting date. A reset on or before the reporting date is rejected: accepting the reporting date would give the 10-day floor even to a contract whose next reset is months away. The shortest horizon is the floor: a daily settled contract, reset the next business day, gets M = 10 business days and MF = √(10/250) = 0.2 with the regulatory parameters | Next reset date on or before the reporting date |
| 5 | Blank or `N` with a next reset date is contradictory, and `STM` needs one | Blank or `N` with a date; `STM` without a date; any other settlement type; a reset date that is not a date |
| 6 | The Basel FAQ to CRE52.31–.32 treats daily-settled trades as unmargined, so an `STM` trade needs the unmargined treatment. In a netting set with Margined = Y there are two cases. If every trade is `STM`, the netting set's classification is inconsistent; the set is not necessarily mixed. If it also holds genuinely margined trades, it is a mixed netting set: CRE52.74 and Art. 274(4) give a calculation treatment for it, which K-SACCR does not implement, as for `MIXED`. Rejecting both cases is a limitation of the current implementation, not a regulatory prohibition. The netting set is `INVALID: STM requires unmargined treatment; mixed margined/unmargined netting sets are not supported`, with an error on Checks | `STM` in a netting set with Margined = Y |
| 7 | The legs of one economic trade (#38), the rows that decompose one contractual trade under one Economic trade ID, share the settlement type and the next reset date. Blank and `N` are the same settlement type for this comparison. Separately contracted trades, such as a hedge or a back-to-back transaction, have their own Economic trade ID or none, and may have different settlement arrangements | Legs of one economic trade that differ in either, after reading blank as `N` |

A trade that breaks rules 3 to 5 is rejected like any other invalid trade:
an error on Checks names the rule, the trade is `EXCLUDED`, and its netting set
is `INCOMPLETE`. Under rule 7 every leg of the economic trade is rejected, as
for #38's other leg rules. Rule 6 makes the whole netting set `INVALID`. No calendar is derived: K-SACCR has no exchange calendar, and the
user enters the reset date.

**Existing workbooks.** With both columns blank, every trade keeps its
earlier treatment and result. A settled contract entered with Margined = N and
the next reset date as M keeps its result; to use the contract, enter `STM`,
the next reset date, and the contractual maturity as M. A settled option whose
expiry is after the reset date, rejected before, can now be entered.

**Upgrade.** A workbook built from an earlier template has Comment in AB. The
Trades header check rejects that layout, so the run stops rather than reading
the wrong columns. To upgrade it, rebuild it from the new template, or insert
two columns before Comment (AB:AC) with the headers above. In the second case
also change `Trades!$C:$AA` to `Trades!$C:$AC` in the formulas of Results C5,
Reconciliation C5 and TradeReconciliation C5, so that an edit in AB or AC
updates those statuses; data validation for the new columns is in the new
template only.

**Run validity.** A change in AB or AC marks the previous results as stale,
as a change in any other Trades input does: the input fingerprint covers
every Trades column before Comment.

**Tests.** The numerical cases have expected values from
[`tests/derivations/settled_to_market.py`](../../tests/derivations/settled_to_market.py),
a separate implementation, and are `illustrative`, as for #359:

- `illustrative-settled-to-market-daily`, both regimes: a listed equity
  index future and call reset the next business day, so M is the 10-day
  floor (MF 0.2) while the call keeps its expiry for the delta;
- `illustrative-settled-to-market-periodic`, both regimes: a five-year EUR
  swap, an equity index call and a Brent forward reset in 30 days (MF
  √(30/365)); the swap keeps its duration and its bucket over five years by
  E, and the forward's reset falls on its maturity date (accepted);
- `TEST_InputValidation`: a reset on and before the reporting date, after M,
  after the end date that fills a blank M (rejected) and on M (accepted); `STM`
  without a reset date, blank or `N` with one, another code, an error value
  and a reset that is not a date (rejected); lower-case `stm` with a daily
  reset (M = 1/365, MF 0.2); a margined netting set whose trades are all
  `STM` (both regimes), one mixing `STM` with ordinary trades and one with
  another invalid field (`INVALID` with the rule 6 status); legs with blank
  and `N` (accepted), with `STM` and `N` or different reset dates (every leg
  rejected), and two separately contracted trades with different settlement
  types (both accepted);
- every other input in the rules table that breaks a rule;
- regression showing that existing cases, with both columns blank, keep
  their results.

## 🔗 Rule, input, code and test

The supported clearing rules and where each is carried.

| Rule | Input | Code | Test | Coverage |
| --- | --- | --- | --- | --- |
| Clearing-member floor, 5 days (CRE54.12; Art. 279c(1), 304(3)(a)) | NettingSets *Client clearing role* `CM`, Margined = Y | `CORE_Engine.LoadNettingSets` | `netting-set-clearing-member-crr`, `netting-set-clearing-member-bcbs`, `-large-crr`, `-large-bcbs`, `-override-below-floor`, `-override-above-floor`; `illustrative-clearing-member-listed` | Implemented; exercised |
| Client floor, 5 days under CRR (Art. 279c(1)) | Role `CLIENT`, Margined = Y, regime CRR | `CORE_Engine.LoadNettingSets` | `netting-set-clearing-client-crr`, `netting-set-clearing-client-disputes-crr`; `illustrative-clearing-client-listed.crr` | Implemented; exercised |
| Client floor, 10 days under BCBS (CRE54.8(2), .14–.17; CRE52.50) | Role `CLIENT`, Margined = Y, regime BCBS | `CORE_Engine.LoadNettingSets`, with an information message on Checks | `netting-set-clearing-client-bcbs`, `netting-set-clearing-client-disputes-bcbs`; `illustrative-clearing-client-listed.bcbs` | Implemented; exercised |
| Role codes | Role other than blank, `N`, `CM`, `CLIENT`, `CCP`, `MULTI` | `CORE_Engine.LoadNettingSets` (`NsCode`) | `netting-set-clearing-role-not-a-code` | Implemented; exercised: the netting set is INVALID |
| CCP-facing netting set rejected ([decision 5](#decisions)) | Role `CCP`, margined or not | `CORE_Engine.LoadNettingSets` | `netting-set-clearing-role-ccp`, `netting-set-clearing-role-ccp-margined` | Implemented; exercised: the status is `INVALID: facing a CCP, not supported`, with any other invalid field counted after it, and Checks has the error |
| Multi-level client structure rejected ([decision 6](#decisions)) | Role `MULTI` | `CORE_Engine.LoadNettingSets` | `netting-set-clearing-role-multi-level` | Implemented; exercised: the status is `INVALID: multi-level client structure, not supported`, and Checks has the error |
| Role, floor and rule on Checks ([decision 7](#decisions)) | Role `CM` or `CLIENT`, margined or not | `CORE_Engine.LoadNettingSets` | The four role cases and `netting-set-clearing-member-unmargined` check the INFO line | Implemented; exercised |
| Next reset date as remaining maturity (CRE52.37(5); Art. 279c(2)) | Trades *Settlement type* `STM` and *Next reset date* ([input contract](#settled-to-market-input-contract)) | `CORE_Engine.ReadTradeSettlement`, `CORE_Engine.CheckEconomicTrades`, `SACCR_MaturityFactor` | `illustrative-settled-to-market-daily`, `illustrative-settled-to-market-periodic`; the `TEST_InputValidation` cases listed under the input contract | Implemented; exercised. The qualification of the contract is the user's |
| Posted collateral not bankruptcy remote in NICA (CRE54.23) | NettingSets NICA | `CORE_Engine` replacement cost and PFE multiplier | Existing NICA cases, none clearing-specific | User's |

## 🚫 Rejections and required attributes

Required behaviour, from the [decisions](#decisions). Until the implementing
issue is done, each case stays as described under *Today*.

| Case | Today | Required |
| --- | --- | --- |
| Netting set facing a CCP | Declared with role `CCP`: rejected. Entered with role `N`: calculated as bilateral, undetectable | Rejected when declared until CCP support exists ([decision 5](#decisions), implemented); then a CCP role with the attributes below, without which it is rejected ([decision 2](#decisions)) (#112) |
| VM retained by the CCP, member collateral unprotected | Cannot be declared; a margined set gets 10 days or less | Declared through the retained-VM and protection inputs and given the Art. 304(3)(d) / CRE54.8(3) horizon ([decision 2](#decisions)) (#112) |
| Multi-level client structure | Declared with role `MULTI`: rejected. Entered as `CM` or `CLIENT`: calculated as that role, undetectable | Rejected when declared ([decision 6](#decisions), implemented) |
| QCCP status missing or unknown | No input | Rejected once CCP-facing sets exist (#112) |
| Settled-to-market option with a reset date before expiry | Entered with `STM` and its reset date, M and T kept (#360, implemented) | [Decision 4](#decisions); [input contract](#settled-to-market-input-contract) |
| Settled-to-market trade in a margined netting set | `INVALID: STM requires unmargined treatment; mixed margined/unmargined netting sets are not supported` (#360, implemented) | A limitation: an all-`STM` set is misclassified, and a mixed set needs the CRE52.74 treatment, not implemented |
| More than 5,000 trades where the floor does not apply | 20 days if the flag is set: a conservative limitation | Separated with CCP support, waiving only the trade-count floor where eligible ([decision 3](#decisions)) (#112, #359) |

Attributes a CCP-facing netting set needs for #112: the CCP's identity; QCCP
status (authorised or recognised); whether the bank is a clearing member or a
client guaranteed by its member; for a clearing member's client trades,
whether it reimburses the client on the CCP's default (Art. 306(1)(c)–(d));
whether the CCP retains VM without protecting the member's collateral;
bankruptcy-remote treatment of posted collateral; and, for default funds, DF_i,
unfunded commitments, K_CCP, DF_CCP and DF_CM as communicated by the CCP
(Art. 308(2)).

## 🧮 Exposure, risk weights and capital

K-SACCR's outputs are exposure values: RC, PFE and EAD per netting set, and
their totals per regime. Risk weights, risk-weighted exposure amounts and own
funds requirements are different quantities and must never enter an EAD
column or an EAD total:

- A trade exposure to a CCP is an exposure value, calculated with SA-CCR (Art.
  306(3); CRE54.8), to which a risk weight is applied separately (Art.
  306(4)).
- A default-fund contribution is not a trade: it gives an own funds
  requirement (Art. 308–310; CRE54.25–.42), never an EAD. It is not entered on
  Trades or in NICA.
- Each trade row belongs to one netting set. A clearing member that
  clears for a client has two exposures to two counterparties, to the client
  and to the CCP, entered as two netting sets; the CCP leg may have an
  exposure value of zero (Art. 306(1)(c)). This is two exposures, not double
  counting.
- A client's exposure is either to its clearing member or, where the member
  guarantees it, to the CCP (CRE54.14(1)–(2)), never both for the same
  transactions.
- Client collateral passed on to the CCP may reduce both legs (Art. 304(7);
  CRE54.13); that is permitted reuse.
- Results totals add EAD across all `VALID` netting sets of a regime. CCP
  trade EAD stays in that total and is separately identifiable there
  ([decision 2](#decisions)).

<a id="decisions"></a>

## ✅ Decisions

Settled by the owner for the clearing work of v1.0.2. Each is an acceptance
criterion of the issue named; until that issue is done, the current behaviour
described above stays.

| # | Decision | Rationale | Acceptance criteria | Issue |
| --- | --- | --- | --- | --- |
| 1 | A `BCBS` client facing its clearing member keeps the 10-day MPOR floor; the clearing member facing its client keeps 5 days | A conservative interpretation. CRE54.2 has both parties capitalise the clearing member–client leg of an exchange-traded derivative cleared under a bilateral agreement as an OTC derivative, and its footnote says the treatment of CRE54.12 "would also apply". CRE54.12 grants at least five days to the clearing member for its exposure to its clients and does not name the client; the client's own treatment is CRE54.14–.17, with 10 days for OTC derivatives (CRE54.8(2)) or the bilateral floors (CRE52.50). Ten days is not lower than either reading of the footnote. `CRR` is unchanged: 5 days for both (Art. 279c(1)) | The rationale stays in the methodology; the four role cases keep their floors; cases cover the exchange-traded leg under a bilateral agreement in both directions | #359 |
| 2 | A netting set facing a CCP is declared with its own role and a QCCP column, with inputs for VM retained by the CCP and for protection of the bank's collateral against the CCP's insolvency | Retained VM without that protection sets the horizon to the lesser of one year and the remaining maturity, floored at 10 business days (CRE54.8(3); Art. 304(3)(d)); neither can be inferred from other inputs | Missing or inconsistent CCP inputs reject the netting set. Risk weights, risk-weighted amounts and default-fund capital are separate outputs and never enter an EAD column or total. CCP trade EAD stays in its regime's EAD total and is separately identifiable there | #112 |
| 3 | The single *large or illiquid* flag stays until CCP support arrives | The current treatment is a conservative limitation: 20 days whenever the flag is set, also where the trade-count floor does not apply | When split, only the trade-count floor is waived, and only where eligible: trade exposures to a CCP (Art. 285(3)(a); CRE54.8(1)) and a clearing member's netting sets with its clients without illiquid collateral, disputed trades or exotic options (Art. 304(3)(c)). The 20-day floor for illiquid collateral or hard-to-replace derivatives (Art. 285(3)(b); CRE52.51(2)) and the doubling for disputes (Art. 285(4); CRE52.51(3)) stay | #112, #359 |
| 4 | A settled-to-market trade is declared with a settlement type and a next reset date | The reset treatment applies only where the settlement extinguishes the accumulated exposure and resets the value to zero (CRE52.37(5); Art. 279c(2)). A daily VM payment alone does not: such a contract is collateralised-to-market and margined | The contractual maturity and the option expiry are kept as separate inputs from the next reset date. The engine validates the declaration and the consistency of the dates, and rejects a missing, past or inconsistent reset date where the settlement type requires one. Whether the settlement extinguishes the exposure is a contractual fact the user attests by declaring `STM`, and remains the user's responsibility; no input can prove it ([input contract](#settled-to-market-input-contract)) | #360 |
| 5 | Until CCP support is implemented, a netting set declared as facing a CCP is rejected | Documentation alone cannot prevent the understated exposure of a CCP-facing set calculated as bilateral, for example where the CCP retains VM (decision 2) | A recognised declaration makes the netting set `INVALID`, with a message naming the missing support; a netting set not declared stays undetectable, as stated above | #112 |
| 6 | A multi-level client structure is declared with role `MULTI` and rejected | Its treatment depends on conditions at every level (CRE54.14(3); CRR Art. 305(4)), which the engine does not model; entered as `CM` or `CLIENT` it would silently take an ordinary role's floor | The netting set is `INVALID: multi-level client structure, not supported`, with an error on Checks | #359 |
| 7 | The role and method applied are shown as an INFO line on Checks for every `CM` and `CLIENT` netting set, not as a Results column | The evidence must show the selected role; a Checks line does so without changing the Results layout, the reconciliation or the CSV export | Margined: the role, the MPOR floor and its rule (for a BCBS client, marked as the conservative interpretation), and the floor applied when the large or illiquid flag raises it; unmargined: that no floor applies | #359 |
| 8 | There will be no independent reviewer | The project has one maintainer, who wrote the code under test, so an `independent` value cannot exist under the evidence policy | The client-clearing values stay `illustrative`, computed by a separate implementation (`tests/derivations/clearing_listed.py`); #359's criterion on independently reviewed values is not met and is recorded as such | #359, #361 |

## 📜 Earlier releases

Releases 1.0.0 and 1.0.1 excluded exposures to central counterparties and
default-fund contributions; that remains a fact of those releases, and their
manuals and notes are not changed by this page.
