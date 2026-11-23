# Dashboard

The **Dashboard** sheet (#675) summarises the last run on one page: what the
exposure is, with whom, in which asset classes, how concentrated, how it
grows from notional to EAD, and how clean the data are. It is an output sheet
like Results: every run writes it, Clear outputs empties it, and it shows the
results as out of date when the inputs change.

This document is the contract of its figures. The machine-readable form is
[`dashboard.json`](dashboard.json): one **block** per view, each a small
table with fixed rows and columns, under a workbook name such as
`Dashboard_KeyFigures`. A run writes only the values of these blocks, into
the data table on the sheet. The layout, the colours, the cards and the
charts are designed once in the template and read the data table; no formula
on the sheet aggregates an input or output sheet.
[`tools/dashboard.py`](../../tools/dashboard.py) computes every block from a
run's records, and its tests (`tools/test_dashboard.py`) pin the definitions
below on a portfolio with two regimes, an invalid and an incomplete netting
set, one out of the money and one holding collateral, and on an empty
portfolio.

## 🧭 Principles

- **No new calculation.** Every figure is a count, a sum, a share or a
  ranking of figures the run already wrote on Results, TradeCalc and Checks,
  of the inputs of NettingSets, or of the results of the two reconciliation
  sheets.
- **Regimes apart.** A block with amounts has one column per regime, CRR
  then BCBS, and amounts of different regimes are never added. A chart shows
  one series per regime; the add-on by asset class has one doughnut per
  regime. A regime with no netting set shows zeros, and its column is still
  written.
- **Valid netting sets.** Amounts, rankings and shares use the netting sets
  of the regime whose Status on Results is `VALID`: those whose EAD Results
  reports. Their sum is the TOTAL row of the regime. Counts of netting sets
  use every netting set of the regime, whatever its status.
- **Used trade legs.** Notional, maturity and market-value figures use the
  rows of TradeCalc with Status `OK` that belong to a valid netting set of
  the regime. A row of TradeCalc is a trade leg: an economic trade entered as
  several legs counts once per leg.
- **Units.** Amounts are in the reporting currency of Params. Each of the two
  families of amounts has its own scale, so that a large notional does not
  flatten the exposure figures:
  - *exposure*: EAD, RC, PFE, add-ons, gross positive value and value after
    netting;
  - *notional*: notional amounts.

  The largest absolute amount of the family, over both regimes, gives
  k = floor(log10(amount) / 3), held between 0 and 3 (0 when every amount is
  0). The family's amounts are written divided by 1, 1,000, 1,000,000 or
  1,000,000,000, and `Dashboard_Scale` gives the factor and the unit, for
  example `EUR thousand`. Every card and chart axis names its unit from that
  block. Counts, shares and the Herfindahl index have no scale.
- **Blank, not zero, when undefined.** A share of a total of 0, a
  concentration without EAD and the unused rows of a ranking are blank.

## 📋 Blocks

| Block | View | Rows × columns | Definition |
| --- | --- | --- | --- |
| `Dashboard_Run` | Run header | 6 × 1 | Reporting date and reporting currency of Params, the run's date and time, the default regime of Params, the parameter set (`REGULATORY` and the set's ID, or `WHAT-IF`) and the input fingerprint of the run summary on Results C4. The status of the results is not written: the sheet shows it with the formula of Results C5, so that it turns out of date when the inputs change |
| `Dashboard_Scale` | Units | 2 × 2 | Factor and unit of the exposure and notional families |
| `Dashboard_KeyFigures` | Key figures | 12 × 2 | Per regime: EAD, RC and PFE of the valid netting sets; netting sets in total, valid, invalid (status starting `INVALID`), incomplete (starting `INCOMPLETE`) and with no trades (`NO TRADES`); margined and unmargined (Margined `Y` or `N`); valid netting sets with the cap applied (`Y`); distinct counterparties (non-blank Counterparty names, compared after trimming) |
| `Dashboard_Counts` | Key figures | 8 × 1 | Trade legs on TradeCalc, used (`OK`) and excluded (`EXCLUDED`); netting sets whose regime is neither CRR nor BCBS; distinct counterparties over both regimes; rows of Checks by severity, ERROR, WARNING and INFO, without the single `RUN-CLEAN` row of a clean run |
| `Dashboard_Sector` | 1. EAD by counterparty sector | 7 × 2 | EAD by the counterparty sector of NettingSets (#674): `CB`, `GG`, `CI`, `OFC`, `NFC`, `HH`, and *Not classified* for a blank or unknown sector |
| `Dashboard_ClearingRole` | 2. EAD by clearing role | 5 × 2 | EAD by the client clearing role of NettingSets: bilateral (`N` or blank), clearing member (`CM`), client (`CLIENT`), CCP qualifying (`CCP` with QCCP `Y`) and not qualifying (`CCP` otherwise). `MULTI` makes a netting set invalid, so it has no EAD |
| `Dashboard_AddOn` | 3. Add-on by asset class | 6 × 2 | Sum of the add-on columns IR, FX, CR, EQ, CO and OT of Results |
| `Dashboard_Composition` | 4. EAD composition | 4 × 2 | RC and PFE (before alpha; with the cap, alpha × (RC + PFE) need not equal EAD), and EAD of the margined and of the unmargined netting sets, whose sum is the regime's EAD |
| `Dashboard_TopCRR`, `Dashboard_TopBCBS` | 5. Ten largest netting sets | 10 × 4 each | NettingSetID, counterparty, EAD and share of the regime's EAD, by EAD from the largest; equal EADs by NettingSetID in ordinal order |
| `Dashboard_Concentration` | 6. Concentration | 2 × 2 | Share of the regime's EAD held by its ten largest netting sets, and the Herfindahl index, the sum of the squared shares of every valid netting set: 1 for one netting set, 1/n for n of equal EAD |
| `Dashboard_Maturity` | 7. Maturity profile | 3 × 2 | Notional of the used legs by M on TradeCalc: under 1 year (M < 1), 1 to 5 years (1 ≤ M ≤ 5), over 5 years (M > 5) |
| `Dashboard_AssetClass` | 8. Trades by asset class | 6 × 4 | Per regime, used legs and their notional by the asset class on TradeCalc |
| `Dashboard_DataQuality` | 9. Data quality | 2 × 1 | Excluded trade legs, and netting sets whose EAD is withheld (status other than `VALID`) |
| `Dashboard_Codes` | 9. Diagnostic codes | 5 × 3 | The five codes with the most ERROR and WARNING rows on Checks, with ERROR if any of its rows is an error, and the number of rows; equal counts put ERROR first, then the code in ordinal order. A row without a code counts as `(none)` |
| `Dashboard_Reconciliation` | 10. Reconciliation | 5 × 2 | For Reconciliation and TradeReconciliation: `Y` when the sheet holds a comparison, else `N` and zeros; then the comparison rows that are `MATCH`, `DIFFERENCE` and `MISSING IN CALCULATOR`, and the rows of any other result |
| `Dashboard_NotionalToEAD` | 11. From notional to EAD | 6 × 4 | Per regime, each with its share of the notional: the notional of the used legs; their gross positive value, Σ max(MtM, 0) over the legs; the value after close-out netting, Σ max(V, 0) over the valid netting sets; RC; PFE; EAD. The notional is on the notional scale and is a card only; the five other rows are bars on one exposure scale, none magnified |
| `Dashboard_Reading` | 11. Reading | 2 × 1 | Per regime: "Netting removes X% of gross positive value; collateral a further Y%.", with X = (gross − netted) / gross and Y = (netted − RC) / gross, percentages to one decimal with halves away from zero. When Y is negative (a threshold, an MTA or collateral posted raises RC): "…; collateral and margin terms add Z%." with Z = −Y. "No trade has a positive market value." when gross is 0; blank for a regime with no valid netting set |

## 🔄 Life cycle

These rules are delivered with the sheet and the writer (#675):

- **Run:** a run that completes writes every block. A run that fails clears
  them, like the other output sheets, so the Dashboard never shows a partial
  run.
- **Clear outputs** empties every block.
- **Out of date:** the sheet's status cell shows the status of Results, so
  the Dashboard reads as out of date as soon as an input changes.
- **Reconciliation:** view 10 is written after the reconciliations that a
  run refreshes.
