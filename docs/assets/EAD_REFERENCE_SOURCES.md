# EAD calculation reference card

`SACCR_EAD_reference.svg` summarises the SA-CCR formulas and supervisory
parameters as the toolkit implements them. It is a documentation aid: it does
not certify the engine or replace the
[methodology](../methodology/README.md), which holds the paragraph-level
traceability and the source-register decisions.

- **Supervisory duration:** CRE52.34 floors SD at ten business days; CRR
  Article 279b(1)(a) has the same floor (see
  [IR_DIAGRAM_SOURCES.md](IR_DIAGRAM_SOURCES.md)).
- **Dates:** S, E, M and T are in years of 365 days, as in the engine.
- **Lambda:** shifted lognormal for negative rates or prices; CRR applies it to
  interest-rate and commodity options under Delegated Regulation (EU)
  2021/931, Article 5; under Basel the value is taken as entered.
- **Collateral:** C is positive when held. A margined netting set uses
  C = VM + NICA. An unmargined set uses C = NICA; under BCBS a one-way
  agreement in favour of the counterparty adds the posted VM with its negative
  sign (CRE52.10, footnote 2), while CRR treats that set as margined
  (Art. 272(7a)) and rejects it as unmargined. The margined cap is a separate
  rule: it always uses C = NICA (CRE52.2; CRR Art. 274(3)).
- **MPOR floor:** 10 business days bilaterally. The 5-day client-clearing
  floor applies to a clearing member (`CM`) under both regimes and to a client
  (`CLIENT`) under CRR only; a BCBS client keeps 10 days (CRE54.8, CRE54.12,
  CRE54.14; CRR Art. 279c(1)). The 20-day floor applies when the user sets the
  large-or-illiquid flag: more than 5,000 trades during the previous quarter,
  or illiquid collateral or hard-to-replace derivatives (CRE52.51; CRR
  Art. 285(3)). The workbook sees only the current count, so more than 5,000
  trades with the flag at N gives a warning and keeps the ordinary floor.
- **Other risks:** OT exists under CRR only (Art. 277(1)(f), 280f). Under
  BCBS an OT trade is rejected and its netting set's EAD withheld; the toolkit
  does not remap it, and the user enters it under the class of its primary
  risk driver.

The SVG is the source. Values were checked against the methodology parameter
table on the release branch at the time of the change; update the card when a
supervisory parameter or formula changes.
