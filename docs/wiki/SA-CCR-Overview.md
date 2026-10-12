# SA-CCR Overview

> **Guide, not policy:** [Methodology](../methodology/README.md) owns the implemented regulatory basis, source versions and rule-to-test traceability.

## What SA-CCR is

The **Standardized Approach for Counterparty Credit Risk (SA-CCR)** is the Basel
regulatory framework for determining the **exposure at default (EAD)** of
derivative transactions and long-settlement transactions.

It is a **non-modelled, standardized exposure methodology**. Rather than relying
on an institution's internal exposure model, SA-CCR combines current exposure
with a regulatory estimate of potential future exposure and recognises, subject
to the applicable rules, netting, collateral and margining arrangements.

Counterparty credit risk is the risk that a counterparty to a transaction
defaults before final settlement while the transaction has positive economic
value to the institution. For derivatives this exposure can change materially
over time, even when the current mark-to-market is small or negative.

## Regulatory scope

Under the Basel Framework, SA-CCR applies to:

- over-the-counter (OTC) derivatives;
- exchange-traded derivatives; and
- long-settlement transactions.

The Basel rules are set out primarily in **CRE52**. In the European Union, the
SA-CCR framework is implemented through the **Capital Requirements Regulation
(CRR)**, principally Part Three, Title II, Chapter 6, Section 3
(Articles 274–280f), subject to the applicable consolidated version and related
delegated acts.

SA-CCR determines the counterparty-credit-risk exposure value. That EAD is then
used within the wider prudential framework to determine risk-weighted assets
and, ultimately, regulatory capital requirements. **SA-CCR itself is therefore
an exposure-measurement framework, not a complete capital-requirement
calculation.**

## Core exposure measure

For a netting set, the standard SA-CCR structure is:

$
EAD = \alpha \times (RC + PFE)
$

where:

- **EAD** is exposure at default;
- **α** is the regulatory multiplier (1.4 in the Basel SA-CCR framework);
- **RC** is replacement cost; and
- **PFE** is potential future exposure.

The calculation is performed at the relevant **netting-set** level. Transactions
that are not covered by a recognised legally enforceable netting agreement are
treated separately for exposure purposes.

## Replacement Cost (RC)

**Replacement Cost** captures current exposure: the amount that could be lost if
the counterparty defaulted today and the portfolio had to be replaced.

The exact formulation depends on whether the netting set is **unmargined** or
**margined** and on the treatment of collateral, variation margin, independent
collateral amounts, thresholds and minimum transfer amounts.

For margined relationships, the framework is designed to reflect the mechanics
of the collateral agreement rather than simply subtracting collateral from
mark-to-market.

## Potential Future Exposure (PFE)

**Potential Future Exposure** captures the possibility that the exposure grows
before the transactions mature or can be closed out.

At a high level:

$
PFE = multiplier \times AddOn_{aggregate}
$

The aggregate add-on is built from trade-level regulatory measures and
supervisory parameters. The calculation reflects factors such as:

- asset class;
- adjusted notional;
- supervisory delta;
- maturity factor;
- supervisory factor;
- hedging-set membership;
- regulatory correlations; and
- permitted offsetting within and across defined groups.

The multiplier can recognise over-collateralisation or a negative current
mark-to-market, subject to the regulatory floor.

## Asset classes and hedging sets

The Basel SA-CCR framework organises trades into five principal asset classes:

1. interest rate;
2. foreign exchange;
3. credit;
4. equity; and
5. commodity.

Each asset class has its own aggregation logic, supervisory parameters and
hedging-set rules. Interest-rate trades, for example, are aggregated by
currency and maturity bucket, while FX trades are grouped by currency pair.

A trade may require allocation to more than one risk category (CRR Art.
277(3); CRE52.28), and an option combination is split into its options (CRR
Art. 274(6); CRE52.42-43). The workbook takes such a trade as legs: rows
sharing an Economic trade ID, each with its own risk attributes and add-on,
of which exactly one, the MtM leg, carries the trade's market value. The
engine rejects the trade when the value would be counted more than once, when
its legs are in different netting sets, or when a linear product is split
within one asset class.

## Margined and unmargined netting sets

SA-CCR distinguishes between **unmargined** and **margined** relationships.

For unmargined sets, maturity enters directly into the maturity factor, subject
to regulatory floors. For margined sets, potential future exposure is driven by
the margin period of risk (MPOR) and by the contractual margining mechanics.

The treatment of variation margin, independent collateral amounts, thresholds,
minimum transfer amounts, one-way margining and special MPOR conditions is a
material source of implementation differences and must be validated separately
for each regulatory regime.

## Basel and EU CRR

The Basel Framework is the international prudential standard. The EU CRR is the
legally binding implementation applicable to EU institutions.

The two frameworks are closely related but must not be assumed to be identical.
Differences in legal text, parameters, scope, options, collateral treatment or
interpretation must be recorded explicitly and tested by regime.

**SA-CCR Benchmark treats CRR as the baseline regime and Basel CRE52 as a
separately supported regime.** A result must always identify the regime under
which it was calculated.

## What SA-CCR Benchmark does

**SA-CCR Benchmark** is an independent Excel/VBA calculation and validation
toolkit. Its purpose is not to price derivatives or source market data. Instead,
it takes trade- and netting-set-level inputs and reproduces the regulatory
exposure calculation so that a bank's own SA-CCR results can be independently
tested and reconciled.

The project is designed to support validation at multiple levels, including:

- trade-level regulatory quantities;
- hedging-set and maturity-bucket aggregation;
- asset-class add-ons;
- replacement cost;
- multiplier and PFE;
- EAD at netting-set level; and
- bank-versus-independent-calculator reconciliation.

Intermediate calculations are deliberately exposed so that a difference in EAD
can be traced to the calculation layer that caused it.

## Validation approach

The project separates implementation from validation evidence.

A calculation is not considered validated merely because it reproduces its own
previous output. Validation evidence is intended to come from a hierarchy of
sources, including:

1. published regulatory worked examples;
2. independently derived and reviewed expected results;
3. invariant and boundary tests;
4. independent external implementations used as comparators, not as regulatory
   authority; and
5. reproducible execution in Excel tied to the exact source commit.

Illustrative examples may demonstrate behaviour, but they do not count as
independent numerical validation.

## Current project boundary

The first release is focused on the **full SA-CCR** calculation needed for
independent bank-calculation validation.

Simplified SA-CCR, the Original Exposure Method (OEM), securities financing
transactions, CVA capital and full RWA calculation remain outside the first
release unless explicitly brought into scope and validated.

The detailed and authoritative project scope is maintained in
[Methodology](../methodology/README.md).

## Regulatory references

- **Basel Committee on Banking Supervision — Basel Framework, CRE52:
  Standardised approach to counterparty credit risk**\
  https://www.bis.org/committees/bcbs/basel-framework/standard/cre?allChapters=true&chapter=50
- **Basel Committee on Banking Supervision — CRE51:
  Counterparty credit risk overview**\
  https://www.bis.org/committees/bcbs/basel-framework/standard/cre/51/
- **Regulation (EU) No 575/2013 (CRR), as amended** — Part Three, Title II,
  Chapter 6, Section 3, Articles 274–280f. The exact consolidated version used by
  the calculator is maintained in the project's
  [regulatory source register](../methodology/README.md#source-register).
- **European Banking Authority — Single Rulebook Q&A**\
  Individual Q&As are used as interpretative evidence where relevant, but are
  not treated as substitutes for the Level 1 legal text or as general numerical
  benchmark suites.

---

> **Important:** this page is explanatory. For implementation decisions,
> supported products, parameter versions and rule-to-test traceability, use the
> project's [Methodology](../methodology/README.md) and its source register.
