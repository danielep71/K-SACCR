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
- **Other risks:** the OT column applies under CRR only; Basel maps such
  trades to the class of their primary risk driver.

The SVG is the source. Values were checked against the methodology parameter
table on the release branch at the time of the change; update the card when a
supervisory parameter or formula changes.
