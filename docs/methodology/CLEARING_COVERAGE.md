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
yet. The clearing tests are all `illustrative`.

## 🔀 Arrangements

The bank's role is the NettingSets column *Client clearing role*: `N` (blank),
`CM` or `CLIENT`. It acts only on a margined netting set, through the MPOR
floor; an unmargined netting set takes its maturity factor from each trade's
remaining maturity whatever the role.

| Arrangement | Entered as | MPOR floor, CRR | MPOR floor, BCBS | Coverage |
| --- | --- | --- | --- | --- |
| Bilateral, no clearing | Role `N` | 10 business days (Art. 285(2)(b)) | 10 business days with daily margining, 9 + N otherwise (CRE52.50(1)–(2)) | Implemented; exercised; validated under BCBS through CRE99 (MPOR with remargining) |
| Clearing member facing its client | Role `CM` | 5 business days (Art. 279c(1), last subparagraph; Art. 304(3)(a)) | 5 business days (CRE54.12) | Implemented; exercised |
| Bank as client facing its clearing member | Role `CLIENT` | 5 business days (Art. 279c(1), last subparagraph) | 10 business days: the CCP treatment where CRE54.15 is met (CRE54.14; 10 days for OTC derivatives, CRE54.8(2)), with a 4% risk weight instead of 2% where only the joint-default protection is missing (CRE54.16); a bilateral trade otherwise (CRE54.17; CRE52.50) | Implemented; exercised. The conditions of CRE54.15 and CRR Art. 305(2)–(3) are the user's. They set the risk weight and, under BCBS, whether the 5,000-trade floor applies (CRE54.8(1), not implemented); the MPOR floor stays as shown |
| Bank facing a CCP: a clearing member's trade exposure, for itself or for clients, or a client's exposure to the CCP guaranteed by its clearing member (CRE54.14(2)) | No role exists | At least 10 business days (Art. 304(3)(b)); the 20-day floor for more than 5,000 trades does not apply to trade exposures (Art. 285(3)(a)); VM retained by the CCP without protection of the member's collateral: the lesser of one year and the remaining maturity, floored at 10 business days (Art. 304(3)(d)) | 10 days for OTC derivatives (CRE54.8(2)); the 5,000-trade floor does not apply (CRE54.8(1)); VM retained by the CCP: as CRR (CRE54.8(3)) | Not supported, as in v1.0.0 and v1.0.1. A CCP entered as the counterparty of a role-`N` netting set cannot be told apart from a bilateral one and gets the bilateral floors (#112) |
| Multi-level client structure: a higher-level client facing a lower-level client | No role exists | The CRR five-day floor is for transactions "between a client and a clearing member" (Art. 279c(1)); the multi-level treatment is in Art. 300(6)–(9), 304(7), 305(4) | Treated as the clearing member–client leg (CRE54.2, CRE54.13, CRE54.14(3)) | Not supported; nothing stops it being entered as `CM` or `CLIENT` |

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
| Listed future | Instrument `Linear`, with M, and S and E for interest-rate and credit futures | Implemented; exercised as a bilateral linear trade. No clearing-specific case (#359) |
| Ordinary listed option, on a security, an index or a future | Instrument `Option`, with price, strike and expiry T; for an option on an interest-rate future, S and E of the underlying | Implemented; exercised as a bilateral option. No clearing-specific case (#359) |
| Daily settled (settled-to-market) contract | Margined = N, and as M the next reset date | User's: see [Settlement and reset](#settlement-and-reset) |
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
| Floor by role | 10; 5 for `CM` and `CLIENT` (Art. 279c(1), 285(2)(b)) | 10; 5 for `CM` (CRE52.50, CRE54.12) | Implemented; exercised by `netting-set-clearing-member-crr`, `-bcbs`, `netting-set-clearing-client-crr`, `-bcbs` |
| Remargining every N business days | F + N − 1 (Art. 285(5)) | 9 + N for bilateral non-daily (CRE52.50(2)) | Implemented: F + N − 1 for every role |
| More than 5,000 trades | 20 days, except for trade exposures (Art. 285(3)(a)); a clearing member may disregard it towards its clients, without illiquid collateral, disputes or exotic options (Art. 304(3)(c)) | 20 days for netting sets not with a CCP (CRE52.51(1)); not for trade exposures to a CCP (CRE54.8(1)) | Implemented conservatively: one *large or illiquid* flag sets 20 days for every role, so the relief cannot be applied. A set over 5,000 trades with the flag at N is VALID with a warning (`netting-set-over-5000-trades`) |
| Illiquid collateral or a hard-to-replace OTC derivative | 20 days (Art. 285(3)(b)) | 20 days (CRE52.51(2)) | Implemented; exercised (`netting-set-mpor-large-weekly-bcbs`) |
| Margin disputes | Doubled (Art. 285(4)) | Floor doubled (CRE52.51(3)) | Implemented: F + N − 1 doubled, as recorded in [Assumptions](README.md#assumptions-and-scope) |
| VM retained by the CCP, member collateral unprotected | Lesser of one year and remaining maturity, floor 10 (Art. 304(3)(d)) | Same (CRE54.8(3)) | Not supported, and not detectable (#112) |

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

| Item | K-SACCR today | Gap (#360) |
| --- | --- | --- |
| Settled-to-market linear contract | Margined = N and M = next reset date; the unmargined maturity factor floors M at 10 business days; E stays the contractual end | No input says the trade is settled-to-market, so nothing checks that M is a reset date rather than a maturity, or that the contract qualifies |
| Settled-to-market option | Rejected when the expiry T is after M, so M cannot be the next reset date | The input contract for options must be defined |
| Collateralised-to-market contract | Margined = Y with the MPOR rules above | — |
| Exchange calendar | None: business days are 250 a year without holidays | Listing does not define a calendar; a date derivation would need an explicit one |
| Past or inconsistent dates | M on or before the reporting date is rejected; M after E is not checked | Rules for reset dates to be defined |

## 🔗 Rule, input, code and test

The supported clearing rules and where each is carried.

| Rule | Input | Code | Test | Coverage |
| --- | --- | --- | --- | --- |
| Clearing-member floor, 5 days (CRE54.12; Art. 279c(1), 304(3)(a)) | NettingSets *Client clearing role* `CM`, Margined = Y | `CORE_Engine.LoadNettingSets` | `netting-set-clearing-member-crr`, `netting-set-clearing-member-bcbs` | Implemented; exercised |
| Client floor, 5 days under CRR (Art. 279c(1)) | Role `CLIENT`, Margined = Y, regime CRR | `CORE_Engine.LoadNettingSets` | `netting-set-clearing-client-crr` | Implemented; exercised |
| Client floor, 10 days under BCBS (CRE54.8(2), .14–.17; CRE52.50) | Role `CLIENT`, Margined = Y, regime BCBS | `CORE_Engine.LoadNettingSets`, with an information message on Checks | `netting-set-clearing-client-bcbs` | Implemented; exercised |
| Role codes | Role other than blank, `N`, `CM`, `CLIENT` | `CORE_Engine.LoadNettingSets` (`NsCode`) | `netting-set-clearing-role-not-a-code` | Implemented; exercised: the netting set is INVALID |
| Next reset date as remaining maturity (CRE52.37(5); Art. 279c(2)) | Trades *Maturity date (M)* | `CORE_Engine` trade dates, `SACCR_MaturityFactor` | None specific | User's |
| Posted collateral not bankruptcy remote in NICA (CRE54.23) | NettingSets NICA | `CORE_Engine` replacement cost and PFE multiplier | Existing NICA cases, none clearing-specific | User's |

## 🚫 Rejections and required attributes

Behaviour to be defined by the implementing issues. Until then each stays as
described under *Today*.

| Case | Today | Required |
| --- | --- | --- |
| Netting set facing a CCP | Entered as role `N` and calculated as bilateral | A declared role for the CCP-facing set; without the CCP attributes below it is rejected (#112) |
| VM retained by the CCP, member collateral unprotected | Cannot be declared; a margined set gets 10 days or less | Declared and given the Art. 304(3)(d) / CRE54.8(3) horizon, or rejected (#112) |
| Multi-level client structure | Can be entered as `CM` or `CLIENT` | Declared, and either supported per level or rejected (#359) |
| QCCP status missing or unknown | No input | Rejected once CCP-facing sets exist (#112) |
| Settled-to-market option with a reset date before expiry | Rejected as expiry after maturity | Defined input contract (#360) |
| More than 5,000 trades where the floor does not apply | 20 days if the flag is set | Separate the trade count from illiquidity, or keep the conservative floor (open decision below) |

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
- Results totals add EAD across all `VALID` netting sets of a regime. Whether
  CCP trade exposures need a separate subtotal is part of #112.

<a id="open-decisions"></a>

## ❓ Open decisions

Settled by the owner in the named issue; until then no code or test assumes an
answer.

| Decision | Options | Recommendation | Issue |
| --- | --- | --- | --- |
| MPOR of a BCBS client whose exchange-traded trades are cleared under a bilateral clearing member–client agreement | Keep 10 days; or 5 days, reading CRE54.2's footnote ("the treatment in CRE54.12 would also apply") as extending the clearing-member floor to the client | Keep 10 days, the higher floor, unless a supervisory reading supports 5 | #359 |
| How a CCP-facing netting set is declared, and what #112 outputs | A new role code with a QCCP column, and EAD only; or also the risk-weighted amount and default-fund requirement in separate outputs | A new role code and QCCP column; risk weights and default-fund requirements in their own outputs, outside EAD totals | #112 |
| The 5,000-trade floor where it does not apply (CCP trade exposures; a clearing member towards clients) | Keep one *large or illiquid* flag (conservative); or separate the trade count from illiquidity | Separate them only together with the CCP role of #112 | #112, #359 |
| How a settled-to-market trade is entered | Keep M as the next reset date, documented; or add an explicit settlement-type and next-reset input, with an option rule | An explicit input, so that the treatment is declared and checked | #360 |

## 📜 Earlier releases

Releases 1.0.0 and 1.0.1 excluded exposures to central counterparties and
default-fund contributions; that remains a fact of those releases, and their
manuals and notes are not changed by this page.
