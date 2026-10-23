<div align="center">

# 🏦 K-SACCR

### Independent SA-CCR exposure calculation and validation in Excel VBA

**EU CRR and Basel CRE52 · Trade-to-EAD audit trail · Regulatory or what-if runs · Published-example regression · No real portfolio data**

<br>

[![Excel VBA](https://img.shields.io/badge/Excel_VBA-64--bit-217346?style=for-the-badge&logo=microsoft-excel&logoColor=white)](#requirements)
[![Windows](https://img.shields.io/badge/Platform-Windows-0078D6?style=for-the-badge&logo=windows&logoColor=white)](#requirements)
[![Regimes](https://img.shields.io/badge/Regimes-EU_CRR_%C2%B7_Basel_CRE52-1f3a5f?style=for-the-badge)](#regimes)
[![Release](https://img.shields.io/badge/Release-1.0.2-2ea44f?style=for-the-badge)](#status)
[![Excel suites](https://img.shields.io/badge/Excel_suites-9%2F9_PASS-2ea44f?style=for-the-badge)](#validation)
[![Regression](https://img.shields.io/badge/Regression-382_cases_%C2%B7_1500_checks-2ea44f?style=for-the-badge)](#test-suites)
[![License](https://img.shields.io/badge/License-MPL_2.0-2ea44f?style=for-the-badge)](LICENSE)

<br>

[![Last commit](https://img.shields.io/github/last-commit/danielep71/VBA-SACCR-Toolkit/main?style=flat-square&label=last%20commit&color=217346)](https://github.com/danielep71/VBA-SACCR-Toolkit/commits/main)
[![Clones](https://img.shields.io/endpoint?url=https%3A%2F%2Fraw.githubusercontent.com%2Fdanielep71%2FVBA-SACCR-Toolkit%2Ftraffic-history%2Fdata%2Fbadges%2Fclones.json&style=flat-square&label=clones)](https://github.com/danielep71/VBA-SACCR-Toolkit)
[![Visitors](https://img.shields.io/endpoint?url=https%3A%2F%2Fraw.githubusercontent.com%2Fdanielep71%2FVBA-SACCR-Toolkit%2Ftraffic-history%2Fdata%2Fbadges%2Fvisitors.json&style=flat-square&label=visitors%20(14d))](https://github.com/danielep71/VBA-SACCR-Toolkit)
[![Visits](https://img.shields.io/endpoint?url=https%3A%2F%2Fraw.githubusercontent.com%2Fdanielep71%2FVBA-SACCR-Toolkit%2Ftraffic-history%2Fdata%2Fbadges%2Fvisits.json&style=flat-square&label=visits)](https://github.com/danielep71/VBA-SACCR-Toolkit)
[![Downloads](https://img.shields.io/github/downloads/danielep71/VBA-SACCR-Toolkit/total?style=flat-square&label=downloads&color=217346)](https://github.com/danielep71/VBA-SACCR-Toolkit/releases/latest)
[![P1 issues](https://img.shields.io/github/issues/danielep71/VBA-SACCR-Toolkit/P1?style=flat-square&label=P1&color=B60205)](https://github.com/danielep71/VBA-SACCR-Toolkit/issues?q=is%3Aissue%20is%3Aopen%20label%3AP1)
[![P2 issues](https://img.shields.io/github/issues/danielep71/VBA-SACCR-Toolkit/P2?style=flat-square&label=P2&color=D93F0B)](https://github.com/danielep71/VBA-SACCR-Toolkit/issues?q=is%3Aissue%20is%3Aopen%20label%3AP2)
[![P3 issues](https://img.shields.io/github/issues/danielep71/VBA-SACCR-Toolkit/P3?style=flat-square&label=P3&color=FBCA04)](https://github.com/danielep71/VBA-SACCR-Toolkit/issues?q=is%3Aissue%20is%3Aopen%20label%3AP3)
[![Stars](https://img.shields.io/github/stars/danielep71/VBA-SACCR-Toolkit?style=flat-square&color=bf8700)](https://github.com/danielep71/VBA-SACCR-Toolkit/stargazers)

<br>

**No add-in · No external reference · No DLL · No installer · Reviewable source for every calculation**

[Project website](https://danielep71.github.io/VBA-SACCR-Toolkit/)
&nbsp;·&nbsp;
[Quick start](#quick-start)
&nbsp;·&nbsp;
[How a run works](#how-a-run-works)
&nbsp;·&nbsp;
[Regimes](#regimes)
&nbsp;·&nbsp;
[Validation](#validation)
&nbsp;·&nbsp;
[Worksheet Functions](#worksheet-functions)
&nbsp;·&nbsp;
[Limitations](#limitations)
&nbsp;·&nbsp;
[Methodology](docs/methodology/README.md)
&nbsp;·&nbsp;
[Wiki](https://github.com/danielep71/VBA-SACCR-Toolkit/wiki)

---

<p align="center">
  <img width="100%"
       alt="K-SACCR: CRE99 Example 5 netting set, PFE multiplier, regime choice and the published regression cases"
       src="docs/assets/SACCR_social_preview.png">
</p>

---

</div>

> [!IMPORTANT]
> **Version 1.0.2** is the current release: the
> [1.0.2 GitHub Release](https://github.com/danielep71/VBA-SACCR-Toolkit/releases/tag/v1.0.2)
> carries the workbook and the user manual. The engine
> implements EU CRR and Basel CRE52, but only Basel rules are **validated** in
> v1.0.2, by the published CRE99 examples. CRR rules are implemented and tested
> against illustrative values, not validated ([methodology](docs/methodology/README.md#assumptions-and-scope)).

> [!NOTE]
> The workbook is **not stored in the repository as a macro-enabled file**. It
> is built from the macro-free template `src/workbook/SACCR_Template.xlsx` and
> the exported VBA source, so every line of calculation code is reviewable in a
> pull request.

## ✨ What this project is

**SA-CCR**, the Standardised Approach for Counterparty Credit Risk, is the
regulatory method for the exposure at default (EAD) of derivatives:

```text
EAD = alpha × (RC + PFE)         alpha = 1.4
PFE = multiplier × AddOn_aggregate
```

**K-SACCR** is an independent Excel/VBA implementation for reproducing,
testing and reconciling that calculation. You enter netting sets and trades;
the engine maps each trade to its asset class, hedging set and bucket, computes
every intermediate quantity, and writes the result per netting set with the
full trail behind it.

It is built for people who need to **check** an SA-CCR number rather than
only produce one:

- each trade's adjusted notional, supervisory delta, maturity factor and
  effective notional is visible on `TradeCalc`;
- each hedging set and bucket is visible on `HedgingSets` and `Buckets`;
- every rejected input is listed on `Checks` with the reason, and the netting
  set's EAD is withheld instead of being computed from a partial portfolio;
- every rule is traced to its CRR article and CRE52 paragraph in the
  [methodology](docs/methodology/README.md#traceability).

<a id="methodology-at-a-glance"></a>

### Methodology at a glance

The reference card below collects the formulas and supervisory parameters on
one page, with the points where CRR and Basel differ. Select it to open the
full-size SVG; the [methodology](docs/methodology/README.md) remains the
authoritative description.

<p align="center">
  <a href="docs/assets/SACCR_EAD_reference.svg">
    <img width="100%"
         alt="SA-CCR EAD calculation reference: netting-set exposure, trade-level inputs, asset-class parameters, aggregation, and CRR versus Basel differences"
         src="docs/assets/SACCR_EAD_reference.svg">
  </a>
</p>

---

## 🎯 Core capabilities

Most SA-CCR numbers come out of systems whose intermediate steps cannot be
inspected. This toolkit is the opposite:

| Area | Capability |
| --- | --- |
| Regimes | EU CRR (default) and Basel CRE52 in one engine, chosen per netting set |
| Asset classes | Interest rate, FX, credit, equity, commodity; other risks under CRR |
| Instruments | Linear derivatives; European, American, Bermudan and Asian options; CDO tranches and nth-to-default; basis and volatility transactions |
| Multi-risk trades | Economic trades entered as legs, one per risk category, with one MtM leg |
| Settled-to-market trades | Declared with settlement type `STM` and the next reset date, which sets the remaining maturity (floored at 10 business days) while the contractual dates keep their roles ([input contract](docs/methodology/CLEARING_COVERAGE.md#settled-to-market-input-contract)) |
| Netting sets | Unmargined and margined; VM, NICA, threshold, MTA; MPOR from floor, remargining frequency, large or illiquid flag, disputes and override; cap at the unmargined EAD |
| Client clearing | Clearing-member and client roles with the 5-business-day MPOR floor where the regime allows it, and a clearing member's trade exposures to a CCP with the 10-day floor, its QCCP status and retained variation margin; the CCP risk weight is stated, not applied (see the [clearing coverage](docs/methodology/CLEARING_COVERAGE.md)) |
| CRR specifics | Other-risks class (8%), climatic commodity hedging set (18%), credit quality steps, lambda-shifted option delta (`CRR-RTS` Art. 5) |
| Run classification | Every run is `REGULATORY` against parameter set `CRE52-CRR-2026-10-09`, otherwise `WHAT-IF` with each difference on Checks |
| Input validation | Unreadable or locale-dependent dates, the 1904 date system, duplicate IDs, codes, FX pairs, Excel error values, moved columns and broken names reject the trade or netting set or stop the run; nothing is silently read as blank |
| Reconciliation | A bank's own figures per netting set and per trade leg, compared with the engine's ([how](#3-run)) |
| Evidence export | One CSV file per run, refused while anything is out of date ([how](#3-run)) |
| Worksheet functions | Ten `SACCR_*` functions for the building blocks, returning Excel errors for invalid inputs |
| Deployment | Plain exported VBA and a macro-free template; no add-in, COM component, DLL or extra VBA reference |
| Testing | Nine Excel suites (382 cases, 1500 checks), `Test_RunAll` summary, static checks in CI, and a validator for the Excel evidence record |
| Scale | 500,000 synthetic trade legs in 1,000 netting sets calculated in under two minutes on the supported 64-bit Excel host, every leg used; no fixed trade limit ([portfolio size](INSTALLATION.md#portfolio-size)) |

The [support matrix](docs/methodology/README.md#support-matrix) lists what is
supported, what the user must enter in a given form, and what is out of scope.

---

<a id="quick-start"></a>

# ⚡ Quick start

To use the released workbook without building it, download
`K-SACCR_1.0.2.xlsm` from the
[1.0.2 release](https://github.com/danielep71/VBA-SACCR-Toolkit/releases/tag/v1.0.2)
and check it as in [INSTALLATION.md › Released workbook](INSTALLATION.md#released-workbook),
then go to [section 3](#3-run). The steps below build the workbook from source.

## 1. Clone and check

```sh
git clone https://github.com/danielep71/VBA-SACCR-Toolkit.git
cd VBA-SACCR-Toolkit
git switch --detach v1.0.2
python tools/check.py
```

`v1.0.2` is the released source; for current development, switch to the
[active release branch](docs/GOVERNANCE.md#branches) instead.

Use a clone, not **Download ZIP** or GitHub's **Raw** view: a clone checks the
VBA out with the CRLF line endings the VBE expects. The checks need Python 3.10
or later and no packages.

## 2. Build the workbook

1. Open `src/workbook/SACCR_Template.xlsx` and save it as an **Excel
   Macro-Enabled Workbook (`.xlsm`)** outside the checkout.
2. In Excel, enable **Trust access to the VBA project object model**.
3. Close the workbook and drag it onto `tools/Sync-SACCR-VBA.cmd`, or pass
   its path to the launcher; there is no default workbook. It asks which
   workbook to build: **P**roduction, the engine and the sheet macros from
   `src/`, or **D**evelopment, which adds the test modules of `tests/modules/`
   that [section 4](#4-test-development-workbook) needs. It then asks whether
   to replace the code of the worksheet and `ThisWorkbook` modules as well;
   they hold only `Option Explicit`, so either answer works. The tests write
   to the input and output sheets: never run them on real data.
4. Open the VBE (**Alt+F11**) and run **Debug → Compile VBAProject**.

The manual import order, references and troubleshooting are in
[`INSTALLATION.md`](INSTALLATION.md).

## 3. Run

The template ships with a demonstration portfolio: the five CRE99 examples and
further illustrative netting sets.

| Button | Macro | Does |
| --- | --- | --- |
| **Run SA-CCR** | `RunSACCR` | Validates the inputs, calculates and writes every output sheet |
| **Validate inputs** | `ValidateInputs` | Checks the inputs and lists the findings on Checks without calculating |
| **Clear outputs** | `ClearOutputs` | Empties the output sheets |
| **Export evidence** (on Reconciliation) | `ExportEvidence` | Writes the evidence report of the current run to a CSV file |
| **Export CSV template** (on Trades and NettingSets) | `ExportTradesCsvTemplate`, `ExportNettingSetsCsvTemplate` | Writes the blank CSV template of the [input contract](docs/csv/CSV_CONTRACT.md): its header line only |
| From the Macros dialog (Alt+F8) | `CheckTradesCsv`, `CheckNettingSetsCsv` | Checks a CSV file against the input contract and lists its findings; writes nothing |

To compare the results with your bank's own figures, enter them on
`BankNettingSets`, one row per netting set with its regime, in the amount scale
set at the top of the sheet. **Run SA-CCR** then writes the comparison to
`Reconciliation`: one row per netting set and figure, matched by ID and never
across regimes, `MATCH` within the absolute or the relative tolerance or
`DIFFERENCE`, and a row for every netting set that cannot be compared, with
the reason: missing on either side, listed twice, of another regime, withheld
or unusable. `BankTrades` and `TradeReconciliation` do the same per trade leg,
by Trade ID, against `TradeCalc`: notional, MtM, M, supervisory duration,
adjusted notional, delta, maturity factors, effective notionals and the
stand-alone add-on, with `EXCLUDED` legs and `NOT APPLICABLE` figures. Each
leg's row also shows its Economic trade ID from Trades, so legs can be grouped
as the bank groups its trades.

**Export evidence** then writes the run to one CSV file, separated by `;`
with a `.` decimal point and UTF-8, whatever Excel's regional settings: the
run's identity and input fingerprint, the parameters, supervisory-factor and
FX tables, Results, both reconciliations, Checks and the report's
limitations, each line starting with its section. It refuses while the
results or a reconciliation are out of date, or after an output cell was
edited, so the file always describes the inputs as they are and what the
run wrote.

<details>
<summary><b>Evidence file</b>: the start of the CSV written by Export evidence</summary>

<br>

<p align="center">
  <img width="100%"
       alt="Evidence CSV opened in a text editor: RUN lines with the format, run summary, result and reconciliation status, input and output fingerprints, reporting date, currency, default regime, totals and status per regime, Checks counts, reconciliation summary and tolerances, then the PARAMS lines and the start of the SF_TABLE section"
       src="docs/assets/SACCR_Evidence_CSV.png">
</p>

</details>

## 4. Test (development workbook)

In the Immediate window:

```vb
Test_RunAll
```

```text
SACCR TEST SUMMARY 2026-10-23 20:47:43
  PASS  TEST_Harness          completeness=COMPLETE; cases=4; assertions=6; failures=0; cleanup=PASS  (0,0 s)
  PASS  TEST_Formulas         cases=9; checks=82; failures=0  (0,2 s)
  PASS  TEST_MainState        cases=29; checks=501; failures=0; caller_state=RESTORED  (36,0 s)
  PASS  TEST_Cases            cases=33; checks=186; failures=0; restore=PASS  (10,0 s)
  PASS  TEST_InputValidation  cases=177; checks=308; failures=0; restore=PASS  (26,1 s)
  PASS  TEST_Aggregation      cases=12; checks=15; failures=0; restore=PASS  (2,5 s)
  PASS  TEST_Invariants       cases=4; checks=23; failures=0; restore=PASS  (1,3 s)
  PASS  TEST_Reconcile        cases=21; checks=180; failures=0  (8,5 s)
  PASS  TEST_CsvReader        cases=93; checks=199; failures=0  (0,2 s)
RESULT=PASS; suites=9; passed=9; failed=0; seconds=84,9
```

The suites rewrite inputs, outputs and Params and restore them afterwards: run
them in a development workbook.

---

<a id="how-a-run-works"></a>

# 🔄 How a run works

```text
Params ─┐
        ├─► validate ─► trades ─► hedging sets ─► asset-class ─► netting set
NettingSets              │          and buckets     add-ons         EAD
Trades ─┘                │                                           │
                         ▼                                           ▼
                      Checks                         Results · HedgingSets
                 (errors, warnings,                  Buckets · TradeCalc
                  what-if differences)                       │
                                                             ▼
BankNettingSets · BankTrades ─────────────────► Reconciliation · TradeReconciliation
                                                             │
                                                             ▼
                                                  Export evidence (CSV)
```

1. **Parameters.** Reporting date and currency, regime, supervisory factors,
   correlations, option volatilities and FX rates are read from Params through
   named ranges. A value outside its range, a missing reporting date or
   currency, or a broken name stops the run; a missing numeric parameter takes
   its regulatory default with a warning.
2. **Netting sets.** Margining, collateral, threshold, MTA, clearing role and
   MPOR inputs are read and checked.
3. **Trades.** Each trade gets its dates in years (S, E, M, T), adjusted
   notional, supervisory delta, maturity factor and effective notional. An
   invalid trade is excluded with all its reasons.
4. **Aggregation.** Trades are grouped into hedging sets and buckets; add-ons
   are aggregated per asset class with the regime's correlations.
5. **Exposure.** RC, multiplier, PFE and EAD per netting set, with the cap for
   margined netting sets. A netting set with a rejected trade or an invalid
   input reports `INCOMPLETE` or `INVALID` and no EAD.
6. **Reconciliation and evidence.** Bank figures, when entered, are compared
   with Results and TradeCalc, and **Export evidence** writes the run to a CSV
   file ([details](#3-run)).

Steps 3 to 5 in one picture: each trade is assigned to its netting set, asset
class and hedging set, then the figures are aggregated back from the trade to
the exposure value (select it to open it full size):

<p align="center">
  <a href="docs/assets/SACCR_aggregation.svg">
    <img width="100%"
         alt="How SA-CCR aggregates a netting set: trades are assigned to a netting set, an asset class (IR, FX, CR, EQ, CO and, under CRR only, OT) and a hedging set; each trade's adjusted notional, supervisory delta and maturity factor give its effective notional, which is netted within the hedging set (by maturity bucket for interest rates, by entity for credit and equity, by commodity type) and multiplied by the supervisory factor; within each hedging set, credit and equity entities and commodity types are combined with their correlations; every asset class sums its hedging-set add-ons, and the asset classes add up to the aggregate add-on, the PFE and EAD = 1.4 x (RC + PFE)"
         src="docs/assets/SACCR_aggregation.svg">
  </a>
</p>

| Sheet | Role |
| --- | --- |
| `README` | In-workbook guide |
| `Params` | Run parameters, supervisory-factor table, FX rates |
| `NettingSets` · `Trades` | Inputs |
| `Import_NettingSets` · `Import_Trades` | CSV staging: a file's records as text, to check and correct before they replace the inputs |
| `Results` | One row per netting set: RC, add-ons by asset class, multiplier, PFE, EAD, MPOR, cap flag, status; a TOTAL row per regime over its `VALID` netting sets; run summary and stale-result indicator |
| `HedgingSets` · `Buckets` | Effective notionals and add-ons per hedging set and bucket |
| `TradeCalc` | Every trade-level quantity, or the reasons a trade was excluded |
| `Checks` | Errors, warnings and what-if differences |
| `BankNettingSets` · `Reconciliation` | A bank's own figures per netting set, and their comparison with Results |
| `BankTrades` · `TradeReconciliation` | The same per trade leg, against TradeCalc |
| `TestCatalogue` · `FormulaTest` | Reference-test catalogue; worksheet-function examples |
| `Basel_vs_CRR` | The two regimes side by side |

Results after a run of the demonstration portfolio: the run summary and
stale-result indicator, RC and the add-ons per asset class, then the
multiplier, PFE, EAD, the cap for margined netting sets and the status, with a
TOTAL row per regime:

<p align="center">
  <img width="100%"
       alt="Results sheet after a run, left part: one row per netting set with regime, margining, trade count, MPOR, V, C, RC, add-ons by asset class, aggregate add-on and multiplier, and TOTAL CRR and TOTAL BCBS rows"
       src="docs/assets/SACCR_Results_AddOns.png">
</p>

<p align="center">
  <img width="100%"
       alt="Results sheet after a run, right part: aggregate add-on, multiplier, PFE, alpha, margined EAD, unmargined cap, EAD, cap applied and status per netting set; two netting sets are INCOMPLETE because trades were rejected"
       src="docs/assets/SACCR_Results_EAD.png">
</p>

The demonstration portfolio enters dates as calendar dates over 365 days a
year, so `BCBS-EX3` shows 5,405 against the 5,406 printed in CRE99: its WTI
trade has M = 0.7507 rather than the printed 0.75. The regression case enters
exactly 0.75 years and matches 5,406
([details](docs/assets/README.md)).

Checks lists every trade the run rejected, with its row and reason, and every
netting set whose EAD is withheld as a result:

<p align="center">
  <img width="100%"
       alt="Checks sheet after the demonstration run: five ERROR lines, three trades excluded with their Trades row and reason, and two netting sets whose EAD is withheld"
       src="docs/assets/SACCR_Checks.png">
</p>

<details>
<summary><b>Calculation detail</b>: TradeCalc, HedgingSets and Buckets</summary>

<br>

<p align="center">
  <img width="100%"
       alt="TradeCalc sheet, left part: per trade, netting set, asset class, status, hedging set, bucket, effective supervisory factor, notional and MtM in reporting currency, S, E, M, T, supervisory duration, adjusted notional and supervisory delta"
       src="docs/assets/SACCR_TradeCalc_Inputs.png">
</p>

<p align="center">
  <img width="75%"
       alt="TradeCalc sheet, right part: per trade, adjusted notional, supervisory delta, unmargined and margined maturity factors, effective notionals, IR bucket, stand-alone add-on, lambda and regime"
       src="docs/assets/SACCR_TradeCalc_AddOns.png">
</p>

<p align="center">
  <img width="100%"
       alt="HedgingSets sheet: per netting set and hedging set, the IR maturity-bucket effective notionals D1 to D3, systematic and idiosyncratic components, hedging-set add-on, unmargined add-on, basis and supervisory factor"
       src="docs/assets/SACCR_HedgingSets.png">
</p>

<p align="center">
  <img width="85%"
       alt="Buckets sheet: per netting set, hedging set and maturity bucket, entity or commodity type, the trade count, supervisory factor, correlation, effective notionals and entity add-ons, unmargined and margined"
       src="docs/assets/SACCR_Buckets.png">
</p>

</details>

<details>
<summary><b>Reconciliation</b>: bank figures on BankNettingSets, compared on Reconciliation</summary>

<br>

<p align="center">
  <img width="100%"
       alt="BankNettingSets sheet: amount scale, absolute and relative tolerance, and the bank's EAD for BCBS-EX1 and BCBS-EX3 and a row for an unknown netting set"
       src="docs/assets/SACCR_BankNettingSets.png">
</p>

<p align="center">
  <img width="100%"
       alt="Reconciliation sheet: the bank's EAD for BCBS-EX1 is a MATCH, for BCBS-EX3 a DIFFERENCE of 4.62 (0.0855%), an unknown bank netting set is MISSING IN CALCULATOR, and every other netting set is NOT IN BANK FILE, each with its note"
       src="docs/assets/SACCR_Reconciliation.png">
</p>

</details>

<details>
<summary><b>Input sheets</b>: the README guide, NettingSets, Trades and Params with its supervisory-factor table (yellow cells are inputs)</summary>

<br>

<p align="center">
  <img width="100%"
       alt="README sheet in the workbook: the seven How to use steps (enable macros, parameters, netting sets, trades, run, review, regime) and the sign and input conventions for direction, sub-class, risk factor, notional, dates, nature and options"
       src="docs/assets/SACCR_README_sheet.png">
</p>

<p align="center">
  <img width="100%"
       alt="NettingSets sheet: the five CRE99 examples and further illustrative netting sets, with margining, clearing role, MPOR, collateral and regime override columns"
       src="docs/assets/SACCR_NettingSets.png">
</p>

<p align="center">
  <img width="100%"
       alt="Trades sheet: one row per trade leg with netting set, asset class, sub-class, risk factor, instrument, direction, option type, nature, notional, MtM and dates"
       src="docs/assets/SACCR_Trades.png">
</p>

<p align="center">
  <img width="62%"
       alt="Params sheet: reporting date and currency, alpha, multiplier floor, MPOR floors, IR correlations, default regime and CRR lambda thresholds, each with its CRE52 or CRR reference"
       src="docs/assets/SACCR_Params.png">
</p>

<p align="center">
  <img width="100%"
       alt="Params sheet, supervisory-factor table: per SF_Key the asset class, category, supervisory factor, correlation, supervisory option volatility, commodity hedging set, the regimes it applies to and its CRE52 or CRR reference; the climatic commodity and other-risk rows apply under CRR only"
       src="docs/assets/SACCR_Params_Supervisory.png">
</p>

</details>

---

<a id="regimes"></a>

# ⚖️ Regimes

EU CRR is the baseline and the default; Basel CRE52 is implemented as the
differences from it. The regime is set on Params and can be overridden per
netting set, and every result reports the regime it used.

The toolkit supports two regimes: `CRR` (default, binding for EU institutions)
and `BCBS` (Basel CRE52). The Basel regime does not replace the CRR. It shows
the effect of the EU-specific deviations, traces results back to the
international standard, and supports Basel monitoring exercises and non-EU
group entities. See
[Why two regimes](https://github.com/danielep71/VBA-SACCR-Toolkit/wiki/Basel-vs-CRR#why-two-regimes).

| Topic | EU CRR | Basel CRE52 |
| --- | --- | --- |
| Asset classes | Six, with other risks (Art. 277(1)(f), 280f; 8%) | Five (CRE52.72); other-risk trades are rejected |
| Commodity hedging sets | Adds climatic conditions (Art. 277a(1)(e); 18%) | Energy, metals, agricultural, other |
| Single-name credit factor | By credit quality step (Art. 280c(5)) | By rating AAA to CCC (CRE52.72) |
| IR and commodity option delta | Lambda shift of `CRR-RTS` Art. 5 | CRE52.40 |
| IR maturity bucket 1 | End date up to and including 1 year (Art. 280a(3)) | Under 1 year (CRE52.57(3)) |
| IR aggregation across buckets | Formula with offsets only (Art. 280a(3)) | With offsets, or sum of absolutes (CRE52.57(5)) |

The [regime differences](docs/methodology/README.md#regime-differences) table
gives the locators and test cases of each row; the `Basel_vs_CRR` sheet shows
them in the workbook.

<details>
<summary><b>In the workbook</b>: the Basel_vs_CRR sheet</summary>

<br>

<p align="center">
  <img width="100%"
       alt="Basel_vs_CRR sheet, first 11 comparison rows: for each area the Basel CRE52 rule, the EU CRR rule, whether they differ, the effect on the numbers, the workbook treatment, the tests that cover it and the source check"
       src="docs/assets/SACCR_Basel_vs_CRR.png">
</p>

</details>

---

<a id="validation"></a>

# ✅ Validation

## Evidence classes

| Class | Meaning | Used for |
| --- | --- | --- |
| `published` | Printed in Basel CRE99 (BCBS 279 Annex 4a), checked against the text | T01–T14: validates Basel rules |
| `illustrative` | A derivation separate from the code under test, a hand calculation or a format example | Every other expected value; never counted as validation |

There are no `independent` values yet: they can come only from an external
implementation registered as a source (#100). No independent review is
required; the project has no separate reviewer.

## Published CRE99 examples

| Example | Published values checked | EAD (USD) |
| --- | --- | ---: |
| 1 · interest-rate swaps and a swaption | Adjusted notional, option delta, IR add-on | 569,000 |
| 2 · credit derivatives | Multiplier | 381,000 |
| 3 · commodity forwards | — | 5,406,000 |
| 4 · IR and credit, one netting set | Replacement cost, aggregate add-on | 936,000 |
| 5 · margined, weekly remargining | MPOR 14 business days, aggregate add-on, multiplier | 1,879,000 |

## Test suites

| Suite | Covers | Cases | Checks |
| --- | --- | ---: | ---: |
| `TEST_Harness` | Worksheet-function regression with a completeness and cleanup verdict | 4 | 6 |
| `TEST_Formulas` | Domain and Excel-error contract of every worksheet function | 9 | 82 |
| `TEST_MainState` | Excel state, stale-result detection, failure paths of the workbook macros and the output-table names | 29 | 501 |
| `TEST_Cases` | The numerical cases, generated from `tests/fixtures/` and `tests/expected/` | 33 | 186 |
| `TEST_InputValidation` | Rejection of malformed inputs and the input conventions | 177 | 308 |
| `TEST_Aggregation` | Row-order independence, offsets and a mixed clearing population | 12 | 15 |
| `TEST_Invariants` | Relations every result must satisfy | 4 | 23 |
| `TEST_Reconcile` | Netting-set and trade-leg reconciliation, the evidence export and the CSV templates | 21 | 180 |
| `TEST_CsvReader` | The VBA CSV reader against every fixture and case the reference implementation pins | 93 | 199 |
| **Total** | | **382** | **1500** |

The harness counts assertions as its checks. The counts are those of 64-bit
Excel; on 32-bit, `TEST_Formulas` skips two `LongLong` checks (80).

<details>
<summary><b>In the workbook</b>: TestCatalogue and FormulaTest</summary>

<br>

<p align="center">
  <img width="100%"
       alt="TestCatalogue sheet: 28 reference tests with area, source, source status, regime, the output cell read, expected value, tolerance, engine value, difference and PASS; summary 28 passed, 0 failed"
       src="docs/assets/SACCR_TestCatalogue.png">
</p>

<p align="center">
  <img width="100%"
       alt="FormulaTest sheet: unit tests of the SACCR_ worksheet functions, each with its inputs, the value recomputed with Excel's own functions, the function's value, the relative difference, PASS and the CRE52 or CRR rule"
       src="docs/assets/SACCR_FormulaTest.png">
</p>

</details>

`python tools/check.py` runs the [static gates](tools/README.md) locally and
in CI, where Ruff, mypy and actionlint run as well. Static checks never stand in for Excel evidence; the
[Excel evidence record](docs/EXCEL_EVIDENCE.md) defines what a run must show,
suite by suite, and `tools/check_excel_evidence.py` validates a record against
the exact commit it claims.

---

<a id="worksheet-functions"></a>

# 🧩 Worksheet functions

The supported API is the `SACCR_Formulas` module, listed with its signatures
in [`docs/PUBLIC_API.txt`](docs/PUBLIC_API.txt). An invalid input returns an
Excel error (`#NUM!`, `#VALUE!` or `#REF!`) instead of a misleading number.

| Function | Returns | Reference |
| --- | --- | --- |
| `SACCR_SupervisoryDuration(S, E, [FloorYears])` | Supervisory duration | CRE52.34; CRR Art. 279b |
| `SACCR_MaturityFactor(M, [Margined], [MPOR_BD], …)` | Maturity factor | CRE52.48–53; CRR Art. 279c |
| `SACCR_OptionDelta(P, K, T, Vol, IsCall, IsBought, [Lambda])` | Supervisory delta of an option | CRE52.40; CRR Art. 279a |
| `SACCR_LambdaCRR(P, K, IsInterestRate, [Threshold])` | CRR lambda shift | `CRR-RTS` Art. 5 |
| `SACCR_CDODelta(A, D, IsLong)` | Supervisory delta of a CDO tranche | CRE52.41 |
| `SACCR_IREffectiveNotional(D1, D2, D3, [Rho12], [Rho23], [Rho13])` | IR effective notional across buckets | CRE52.57 |
| `SACCR_FactorAggregation(AddOns, Rhos)` | Single-factor aggregation | CRE52.61, .68, .70 |
| `SACCR_Multiplier(VminusC, AddOn, [FloorPct])` | PFE multiplier | CRE52.23; CRR Art. 278(3) |
| `SACCR_ReplacementCost(V, C, [Margined], [TH], [MTA], [NICA])` | Replacement cost | CRE52.10, .18; CRR Art. 275 |
| `SACCR_NormCDF(x)` | Standard normal distribution | — |

Everything in `src/core/` is internal and may change without notice.

---

## 📦 Repository structure

```text
src/     engine (core/), worksheet functions (modules/), template and sheet macros (workbook/)
tests/   test modules, fixtures and expected results
docs/    methodology, evidence record, governance, house style, images
tools/   static checks, case generator, VBA synchronizer and exporter
```

The full layout and its rules are in
[`docs/REPOSITORY_STRUCTURE.md`](docs/REPOSITORY_STRUCTURE.md).

---

<a id="requirements"></a>

# 💻 Requirements

Desktop Excel for Windows: the [supported hosts](INSTALLATION.md#supported-hosts)
and the VBA references are listed in `INSTALLATION.md`.

---

<a id="limitations"></a>

## ⚠️ Known limitations

- **Released versions only are stable.** Behaviour on the
  [active release branch](docs/GOVERNANCE.md#branches) can change until the
  next release; use the `v1.0.2` tag or its workbook for stable results.
- **CRR is implemented, not validated.** The only published reference values
  are Basel examples; CRR results are tested against illustrative values.
- **One leg per FX trade.** The notional is the leg the user enters; choosing
  the leg when both are foreign is the user's.
- **FX options use a fixed quote order.** An FX option is rejected unless its
  pair follows the ranking EUR, GBP, AUD, NZD, USD, CAD, CHF, JPY (USD/JPY, not
  JPY/USD), because the two orders give different deltas (CRE52.59 FAQ1). Two
  currencies outside the ranking are ordered alphabetically, which may differ
  from local market practice.
- **Not supported:** digital options (#120), the optional reductions for sold
  options and sold credit protection, several margin agreements in one netting
  set, the 1904 date system, default-fund contributions to a CCP (#400), and
  the simplified SA-CCR, OEM, SFT, CVA and RWA calculations.
- **Large portfolios.** 500,000 legs calculated in under two minutes on the
  supported host, but the Trades dropdowns and validation cover rows 5 to 5,000
  only, and a used range inflated far below the data slows every run
  ([portfolio size](INSTALLATION.md#portfolio-size), #291).
- **Windows desktop Excel only.**

The full list is in the
[support matrix](docs/methodology/README.md#support-matrix).

---

<a id="roadmap"></a>

## 🗺️ Roadmap

K-SACCR calculates the risk-based SA-CCR exposure value. The CRR uses that
value, or the same method with its own adjustments, well beyond counterparty
credit risk, and the later milestones follow it there: CCP exposures, CVA risk
and risk-weighted assets.

<p align="center">
  <a href="docs/assets/SACCR_uses_in_CRR.svg">
    <img width="100%"
         alt="Where SA-CCR is used in the EU CRR: counterparty credit risk, exposures to CCPs and large exposures take the risk-based exposure value, and so does CVA risk; the leverage ratio and the output floor recalculate SA-CCR with their own adjustments; with the K-SACCR milestone or scope of each"
         src="docs/assets/SACCR_uses_in_CRR.svg">
  </a>
</p>

| Milestone | Scope |
| --- | --- |
| v1.0.0 | First supported Excel release: the current engine with its evidence |
| v1.0.1 | Post-release hardening: evidence export that keeps the previous file until the new one is written (#312), export bound to its workbook (#313), Results status `EDITED` (#315), troubleshooting (#317), reconciliation error kept (#318), Other Risks in the benchmark (#309), package checks as standing release steps (#308) |
| v1.0.2 | Clearing coverage and CCP completeness: coverage matrix and regulatory contract (#358), listed futures and options across client-clearing roles (#359), settlement and reset inputs (#360), CCP trade exposures (#112), reconciliation evidence and published perimeter (#361) |
| v1.1.0 | CSV import and staging for Trades and NettingSets (#206–#211), workbook protection against accidental edits (#420–#422), and integration/release evidence (#212) |
| v1.2.0 | Excel Ribbon interface (#213–#218), Home landing page (#330), worksheet navigation sidebar (#329) and a Restore Parameter Defaults command (#357) |
| v1.3.0 | Diagnostics and independent validation: engine separated from Excel I/O (#220), stable Checks codes (#81), independent cross-check (#100, #352), randomised FormulaTest inputs (#165), formatted run report next to the CSV evidence (#259) |
| v1.4.0 | Advanced products and FX provenance: digital options (#120), FX quote import (#157), TARGET holiday calendar (#328) |
| v1.5.0 | Simplified SA-CCR (#107) |
| v1.6.0 | Original exposure method (#111) |
| v1.8.0 | Securities financing transactions (#113) |
| v1.9.0 | Regulatory CVA risk (#114) |
| v1.10.0 | Risk-weighted assets and capital (#115), default-fund contributions to CCPs (#400) |

The [milestones](https://github.com/danielep71/VBA-SACCR-Toolkit/milestones)
hold the current scope of each release. v1.7.0 is not used: its scope,
exposures to CCPs, was delivered in v1.0.2 (#112), and default-fund
contributions moved to v1.10.0 (#400).

### v1.1.0 — CSV import, staging and workbook protection

Development starts from the released **v1.0.2** baseline on
`release/1.1.0`, retaining its clearing and settlement input contracts.
The following work is planned; it is not part of the released workbook:

| Workstream | Deliverables | Issues |
| --- | --- | --- |
| CSV input | Versioned Trades/NettingSets contracts and blank templates; deterministic local parsing that preserves identifiers and prevents formula execution | #206, #207 |
| Staging and commit | Editable staging, actionable validation findings, correction/revalidation and explicit replacement of selected live datasets with rollback and stale-result handling | #208–#211 |
| Workbook protection | Editable input cells, protected generated content and workbook structure, macro-compatible protection, recovery tests and documented maintenance unlock | #420–#422 |
| Integration and closeout | Combined regressions, candidate-bound Excel evidence, operating guidance and release checks for both workstreams | #212 |

Acceptance requires equivalent manual and CSV portfolios to produce the
same results, invalid or stale candidates to be refused at commit, and
normal supported workflows to work with protection active. Error and
cancellation paths must preserve or explicitly report recovery of data,
protection and Excel state. Protection prevents accidental edits; it does
not replace validation, freshness checks or output-integrity checks.

The scope does not change SA-CCR calculation methodology or add automatic
post-import calculation, incremental upserts, online feeds or XLSX/JSON
imports. The keyboard-interruption proposal (#417) was closed as not
planned and is excluded. The existing recalculation guard remains in place.

The milestone's due date is a planning target, not a release commitment.
Release publication remains subject to the owner's explicit authorization
and the gates in [RELEASING.md](RELEASING.md).

---

## 📚 Documentation

| Document | Covers |
| --- | --- |
| [Wiki](https://github.com/danielep71/VBA-SACCR-Toolkit/wiki) | SA-CCR overview, regulatory sources, worked examples and workbook guide |
| [`docs/methodology/`](docs/methodology/README.md) | Regulatory basis, source register, support matrix, regime differences, parameters, traceability |
| [`docs/methodology/TEST_CASES.md`](docs/methodology/TEST_CASES.md) | Numerical test-case format and evidence classes |
| [`docs/csv/CSV_CONTRACT.md`](docs/csv/CSV_CONTRACT.md) | Trades and NettingSets CSV input contract, version 1, and its blank templates (import planned for v1.1.0) |
| [`tests/README.md`](tests/README.md) | Test catalogue and suites |
| [`INSTALLATION.md`](INSTALLATION.md) | Developer setup, checks, Excel import, test suites, automation entry points, upgrade and removal |
| [`docs/user-manual/v1.0.2/`](docs/user-manual/v1.0.2/README.md) | Source of the user manual for 1.0.2; its PDF is attached to the [1.0.2 release](https://github.com/danielep71/VBA-SACCR-Toolkit/releases/tag/v1.0.2). The 1.0.0 and 1.0.1 manuals stay in [`docs/user-manual/v1.0.0/`](docs/user-manual/v1.0.0/README.md) and [`docs/user-manual/v1.0.1/`](docs/user-manual/v1.0.1/README.md) |
| [`CONTRIBUTING.md`](CONTRIBUTING.md) | Workflow, commit messages, evidence and pull requests; the review checklist is in the [pull-request template](.github/PULL_REQUEST_TEMPLATE.md) |
| [`docs/REPOSITORY_STRUCTURE.md`](docs/REPOSITORY_STRUCTURE.md) | Profile decision, source layout, public API boundary |
| [`docs/VBA_HOUSE_STYLE.md`](docs/VBA_HOUSE_STYLE.md) | VBA naming, contracts, errors, Excel-state cleanup, export rules |
| [`docs/EXCEL_EVIDENCE.md`](docs/EXCEL_EVIDENCE.md) | Manual Excel run record for every suite, procedure and validator |
| [`docs/GOVERNANCE.md`](docs/GOVERNANCE.md) | Branch model and integration into `main`, ownership and review, issue metadata, GitHub settings and labels |
| [`tools/README.md`](tools/README.md) | Static checks, CI, evidence validator, VBA synchronizer and traffic export |
| [`RELEASING.md`](RELEASING.md) | Release steps: versioning, release invariants and current status |
| [`SECURITY.md`](SECURITY.md) | Private vulnerability reporting and data handling |
| [`CODE_OF_CONDUCT.md`](CODE_OF_CONDUCT.md) | Participant behaviour and conduct reporting |
| [`CHANGELOG.md`](CHANGELOG.md) | Change history |

---

<a id="status"></a>

## 📌 Status

**Version 1.0.2**, released on 2026-10-23 and tagged `v1.0.2`, after 1.0.1
on 2026-10-18 and 1.0.0 on 2026-10-17. Its
[GitHub Release](https://github.com/danielep71/VBA-SACCR-Toolkit/releases/tag/v1.0.2)
carries the production workbook `K-SACCR_1.0.2.xlsm`, the user manual PDF and
their SHA-256 sums. The release notes and the tag message record the Excel
certification of the tagged source: compile, the eight suites each run on its
own, and the evidence validator.

Development continues on the
[active release branch](docs/GOVERNANCE.md#branches) through task branches and
pull requests; `main` receives it on the owner's instruction, and only a
release is versioned and tagged.

---

## 🔐 Data and security

No real portfolio data: never commit or attach real trades, netting sets,
counterparty or collateral data, or any client or employer material. VBA runs
with the user's Office permissions: review the source and follow your
organisation's macro policy. Report suspected vulnerabilities privately as
described in [`SECURITY.md`](SECURITY.md).

---

## 👤 Author

**Daniele Penza** — [@danielep71](https://github.com/danielep71)

## 📄 License

[Mozilla Public License 2.0](LICENSE)
