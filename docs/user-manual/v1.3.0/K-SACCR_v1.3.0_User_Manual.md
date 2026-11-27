---
title: "K-SACCR"
subtitle: "Comprehensive User Manual | Version 1.3.0"
author: "Independent SA-CCR Calculation and Validation Toolkit"
date: "Release 1.3.0"
lang: en-US
colorlinks: true
linkcolor: KBlue
urlcolor: KBlue
toc-title: "Contents"
---

> **RELEASE STATUS**\
> This manual documents **K-SACCR 1.3.0**. It is delivered with the workbook `K-SACCR_1.3.0.xlsm` and the file `SHA256SUMS.txt`, which holds the SHA-256 of the workbook and of this manual. The workbook was tested in Excel before release. This manual explains the software; it does not certify it, and it is not regulatory approval.

**Document purpose.** A practical and technical operating reference for analysts, validators, model owners, risk managers and reviewers using the K-SACCR workbook, a Windows Excel/VBA implementation of the **Standardised Approach for Counterparty Credit Risk (SA-CCR)**. It explains how to check and open the workbook, move around it with its **K-SACCR** Ribbon tab and the navigation bar of every sheet, what its protection lets you edit, enter portfolios by hand or import them from CSV files, validate them, interpret trade-to-EAD calculations and the Dashboard, reconcile an external bank calculation, export reproducible evidence and a formatted report, and understand the declared scope and limitations.

**Reader convention.** `CRR` means the European Union Capital Requirements Regulation implementation in this toolkit; `BCBS` means the Basel Framework CRE52 implementation. “Bank” is shorthand for the independent upstream calculation being compared. “Trade” may mean a row-level **leg**, not necessarily a whole economic trade. Amounts are **unscaled units** of the reporting currency unless an explicit bank-reconciliation scale is used. A result generated under one regime must not be added to the same portfolio's result under another regime.

**Precedence.** This manual describes version 1.3.0 only. The workbook's own `Home` and `README` sheets, the notes of its column headers and its messages take precedence over an outdated copy of the manual, and the applicable legal text takes precedence over both.

\newpage


\Needspace{12\baselineskip}

# Orientation and safe use


\Needspace{9\baselineskip}

## What the toolkit does

K-SACCR is an independent, locally executed Microsoft Excel/VBA calculator that implements the full SA-CCR methodology for supported derivatives. Its main purpose is **transparent validation**: not only producing an exposure-at-default figure, but also showing the trace from a trade leg through effective notionals, hedging sets, maturity buckets, asset-class add-ons, replacement cost and potential future exposure to the reported EAD. It is designed for regulatory-methodology analysis, second-line challenge, independent reconciliation, education and controlled what-if analysis.

The calculation is performed for each **netting set**. A netting set represents the group of transactions under a legally and operationally relevant netting arrangement whose eligible exposure is calculated together. Inputs consist of (a) model settings, supervisory factors, the day count and FX conversion rates; (b) netting-set and collateral conditions; and (c) derivatives and, where relevant, trade legs. Optional bank figures provide an external comparator. The workbook does not price derivatives. Values and descriptors must be supplied by the user, typed on the input sheets or imported from two CSV files through staging sheets where they are checked before they replace the inputs. The one market input the workbook can fetch, on request, is the FX table: **Download prices** writes closing prices or the ECB euro reference rates as its rates (see *FX_Table: conversion into reporting currency*).


\Needspace{6\baselineskip}

### Calculation overview

The high-level SA-CCR structure is:

\[
\mathrm{EAD}=\alpha\times(\mathrm{RC}+\mathrm{PFE}),\qquad
\mathrm{PFE}=\mathrm{Multiplier}\times\mathrm{AddOn}_{\mathrm{aggregate}}.
\]

For a standard regulatory run, **alpha is 1.4**. `RC` is replacement cost; `PFE` is potential future exposure; `AddOn_aggregate` combines the supervisory add-ons of the relevant asset classes. A separate unmargined cap is applied where the netting set is margined. The workbook also records intermediate quantities needed to explain why two EAD calculations differ.


\Needspace{6\baselineskip}

### Intended and unintended uses

**Appropriate uses:** reproduce published Basel examples; challenge a bank's SA-CCR calculation; inspect treatment of dates, FX conversion, option delta and margin terms; document reconciliation exceptions; produce a controlled evidence CSV and a readable report of a run; and test the effect of changed supervisory parameters, explicitly labeled `WHAT-IF`.

**Not established by this version:** universal regulatory certification, counterparty legal enforceability, collateral-eligibility assessments, derivative valuation, regulatory reporting submissions, default-fund contributions to a CCP, risk weights (including those of CCP trade exposures, which are stated but not applied), securities-financing transactions, CVA or RWA. These are distinct controls and processes. An output labeled `REGULATORY` indicates that the workbook parameters match the identified reference set; it is **not** a certification of the input data, legal netting arrangements or final release package.


\Needspace{9\baselineskip}

## Version, evidence and validation boundary

K-SACCR 1.3.0 implements both Basel CRE52 and EU CRR rules. What each rule has been checked against depends on the rule and the regime:

| Evidence | What it is | What it validates |
|:--|:--|:--|
| **Published examples** | The numerical examples 1–5 of Basel CRE99. | The Basel rules those examples exercise. |
| **Independent implementation** | An external SA-CCR implementation by another author, the R package `SACCR` 3.4. On a synthetic portfolio of 40 netting sets in every asset class, a run of the workbook agreed with it on every figure of 37 netting sets; the differences of the other three come from the package. Six of the agreeing netting sets, and further single values, are kept as test cases. | The CRR rules the package computes as the workbook does: among them EAD and alpha, RC (with threshold, MTA and NICA), the multiplier, PFE, the aggregate add-on, the interest-rate, FX, credit (credit quality steps 1 to 6 and indices), equity, commodity (positions of one sign) and other-risk add-ons. |
| **Production reconciliation** | A comparison with a bank's production SA-CCR engine on the bank's own portfolio, under the CRR at 30 June 2026: about 900 netting sets and about 4,300 trades. Every difference was traced; no defect was found in K-SACCR. The bank, its engine and its data are confidential. | The CRR rules that portfolio exercises: EAD and alpha, RC, the multiplier, PFE, supervisory duration, option delta and the lambda shift of interest-rate options, maturity factors and MPOR, interest-rate, FX and equity aggregation and their supervisory factors. |
| **Illustrative values** | Values derived separately from the code, by hand or by a separate script, with no independent review. | Nothing on their own. |

Rules no external source covers are **not validated**, in particular: the cap of a margined netting set where it binds, client clearing roles and trade exposures to a CCP, settled-to-market trades, basis and volatility transactions, the CRR climatic hedging set, the lambda shift of commodity options, the offset of commodity positions of opposite sign, and every Basel rule the published examples do not exercise. Describe the result accordingly: **“Basel rules validated against the published CRE99 examples; the CRR rules named above validated by an independent implementation and a production reconciliation; other rules implemented and illustratively tested”**. A comparison at one date with one bank's engine is not a certification of the CRR implementation.

Before release, the workbook's VBA code was compiled and tested in 64-bit Windows Excel with **thirteen automated test suites, a test harness among them**, all passing. The tests are not part of the delivered workbook. That testing covers the code as delivered: check the workbook's SHA-256 when you receive it (see *Getting started*) and do not change its VBA code.

**Changes from 1.2.0.** The SA-CCR rules are unchanged, with one exception: each regime's credit scale is now enforced (see below). New in the way you work with the workbook are:

* **Navigation.** Every sheet has a **navigation bar** in column A, with a link to every sheet and `Home` at its top; it works with macros disabled. **Navigation bar** on the K-SACCR tab, or the small triangle at the top left of every sheet, hides and shows it. The workbook opens on the K-SACCR tab, with every sheet at a zoom of 100%. See *The navigation bar*.
* **The K-SACCR tab.** Its commands have icons of their own. **Clear outputs** is a button of its own, and every output sheet has a **Clear outputs** button. **Export report** is new in the Evidence group. The long commands show their progress in Excel's status bar. See *The K-SACCR tab and the navigation bar*.
* **Two new sheets, 21 in all.** The **Dashboard**, the first output sheet, shows the last run in one page (see *The Dashboard*). **ReconSummary** sums the two reconciliations per measure.
* **The formatted run report.** **Export report** writes a report of the run as a PDF or an XLSX workbook, beside the evidence CSV (see *The formatted run report*).
* **Checks.** Two columns, **Code** and **Field**, give every line a stable diagnostic code and the field it concerns. A rejected trade now gets one line per violation.
* **Inputs.** `NettingSets` and `Trades` have short column headers, each with a note, the header on row 4, frozen with the rows above it, and the data from row 5. `NettingSets` has a new, optional last column, **Counterparty sector**. A committed import formats every row it writes.
* **Credit scale per regime.** The ratings `AAA` to `CCC` are accepted under BCBS only, and the credit quality steps `CQS1` to `CQS6` under CRR only.
* **Dates.** The days of the year (`YearConvention`) and the **calendar** (calendar days, the TARGET business days, and a table of closing days) are now settings on `Params`. The default, 365 calendar days, gives the same years as before. Each run records its day count.
* **FX rates.** The FX table is now quoted **as the market quotes it**: units of the currency per 1 unit of the reporting currency. **Download prices** writes closing prices or the ECB reference rates, a change of the reporting currency offers to rebase the table, and each run records where its rates came from.
* **CSV files.** NettingSets files are now of version 2, with the sector; version 1 files are still read. The import also checks that the notional currency of an FX trade is a currency of its pair. The bank figures can be loaded from two new CSV files.
* **Reconciliation.** Both bank sheets are delivered empty and have **Import CSV**, **Export CSV template** and **Clear bank figures** buttons. The two reconciliation sheets have summary tiles, filters and number formats per measure, and are no longer protected. The default tolerances are 1 unit and 1%.

See *Moving inputs from 1.2.0*.


\Needspace{9\baselineskip}

## Recommended user journey

![Inputs-to-results process map. The diagram is schematic; the workbook records full details.](assets/data_flow.png){width=98%}


1. Check that the workbook you received is the released file: its SHA-256 must equal the expected value, obtained from a source you trust.
2. Keep that file unchanged as a reference copy, and work in a copy of it.
3. Open the copy and enable macros as your policy allows. The workbook opens on `Home`, with the **K-SACCR** tab of the Ribbon selected. Inspect the worksheets, using the navigation bar on the left of every sheet, and work through the synthetic demonstration. The yellow cells are the ones you can change; every other input cell is protected.
4. Set Params, including the day count and the FX rates, then populate NettingSets and Trades with synthetic or properly authorized test inputs: type or paste them, or import two CSV files onto `Import_NettingSets` and `Import_Trades`, correct them there, press **Validate staging** until no finding remains and press **Commit import**.
5. Use **Validate inputs** before **Run SA-CCR**; resolve all `ERROR` findings. A commit calculates nothing: press **Run SA-CCR** after it.
6. Inspect the Dashboard, Results, TradeCalc, HedgingSets, Buckets and Checks; do not rely on EAD alone.
7. Optionally enter external bank figures on BankNettingSets and BankTrades, typed or loaded from CSV files; run again to refresh the two reconciliations, and read ReconSummary.
8. Export evidence only when the result and any reconciliation are `CURRENT` and the exported output cells are unmodified. The tab's Status group and the enabled state of **Export evidence** help you see this; the export checks it again itself. **Export report** writes a readable report of the same run, under the same conditions.
9. Archive the input description, expected sources, any imported CSV files, the evidence CSV and report, the workbook's file name and SHA-256, and review conclusions under your organization's controls.


\Needspace{9\baselineskip}

## A map of the 21 worksheets

The sheets are listed in the order of their tabs, which is also the order of the **Go to** menu and of the navigation bar.

| Sheet | Type | What you do / find |
|:--|:--|:--|
| `Home` | Start page | The first sheet, on which the workbook opens: what K-SACCR calculates and does not cover, the six steps of a run with the commands they use, and a link to every other sheet. Holds no input. |
| `Params` | Configuration and reference data | Reporting date, reporting currency, default regime, days of the year, calendar, FX rates with their source and the table of closing days, which you set; supervisory parameters and supervisory-factor lookup, which are locked reference values. |
| `NettingSets` | Input | One row per netting set; margin terms, VM, NICA, collateral conditions, regime and alpha override, counterparty sector. |
| `Trades` | Input | One row per trade leg; identifiers, instrument/risk drivers, currency, dates, pricing inputs and economic-trade links. |
| `Import_NettingSets` | Staging | Netting sets loaded from a CSV file, as text, to check and correct before they replace `NettingSets`. |
| `Import_Trades` | Staging | Trade legs loaded from a CSV file, as text, to check and correct before they replace `Trades`. |
| `Dashboard` | Output | The last run in one page: cards, key figures, breakdowns, concentration, data quality, reconciliation and the path from notional to EAD, with ten charts; each regime apart. |
| `Results` | Output | Netting-set RC, add-ons, PFE, EAD, capped EAD, status; separate totals for CRR and BCBS, run summary and freshness. |
| `HedgingSets` | Output | Per-hedging-set effective notionals and associated add-on components. |
| `Buckets` | Output | Per-bucket and underlying-reference aggregation details. |
| `TradeCalc` | Output | Accepted/excluded trade legs and every major leg-level numerical intermediate. |
| `Checks` | Output | Blocking errors, nonblocking warnings, explanatory notices and what-if deviations of a run, or the findings of a staging validation, each line with its origin, code and field. |
| `BankNettingSets` | Optional bank input | External netting-set metrics and reconciliation scale/tolerances; delivered empty. |
| `Reconciliation` | Output, not protected | Bank-vs-calculator comparisons, one netting set and metric per row, with summary tiles. |
| `BankTrades` | Optional bank input | External trade-leg metrics and trade comparison settings; delivered empty. |
| `TradeReconciliation` | Output, not protected | Bank-vs-calculator comparisons per trade leg and metric, with summary tiles; includes economic trade ID. |
| `ReconSummary` | Output | One row per measure of the two reconciliations: rows shown, results, and over the rows compared the sums and the largest difference. |
| `TestCatalogue` | Reference | Catalogue of numerical demonstrations and test expectations. |
| `FormulaTest` | Reference | Worksheet-function examples and expected formulas, with reproducible random inputs. |
| `Basel_vs_CRR` | Reference | Comparison of regulatory treatments and their impact. |
| `README` | In-workbook guide | The last tab: conventions, methodology summary, worksheet functions, scope and limitations, and a reference card. |

**Layout convention.** Column A of every sheet holds the navigation bar and column B is a narrow margin; the title is in `C1` and `C2`, and every worksheet table begins in **column C**. The input tables are found through workbook names pointing to their first header cell. Do not move or delete named header cells, break named ranges, or rearrange input columns. The workbook's protection refuses these changes in normal use; it is described in *What you can edit: workbook protection*.

![The Home sheet: what the workbook does, the steps of a run and a link to every sheet, with the navigation bar on the left.](assets/home.png){width=98%}


\Needspace{12\baselineskip}

# Getting started


\Needspace{9\baselineskip}

## Supported environment

| Component | Requirement / supported level |
|:--|:--|
| Host | **Microsoft Excel for Windows, 64-bit**, Microsoft 365 or Excel 2016 and later. |
| 32-bit Windows Excel | Best effort, **not certified**. |
| Excel for Mac, Excel for web, Excel 2013 or earlier | **Not supported**. |
| Macros | The workbook needs macros enabled. With macros disabled, the K-SACCR tab is shown but none of its commands runs; the navigation bar and the links on `Home` still work. |
| External components | None installed: no add-in, DLL, COM registration or extra VBA reference. The workbook calls one function that is part of Windows, `RtlMoveMemory`, described below. |
| Network | None, except **Download prices** on `Params`, on request and after a confirmation: the closing-price source needs Microsoft 365 and a connection (`STOCKHISTORY`); the ECB source sends one HTTPS request through Windows' own web component and proxy settings. See *Download prices*. |

**Macro security.** The workbook runs VBA with your Office permissions. Accept it only from a source you trust, check its SHA-256, and follow your organization's macro policy. Its VBA code can be read in the Visual Basic editor (**Alt+F11**). The workbook does not need *Trust access to the VBA project object model*; leave that setting off.

**The one Windows function.** The workbook's code declares one function of Windows, `RtlMoveMemory` of `kernel32`, which every Windows program already has. It is used for one purpose: after a VBA error, the K-SACCR tab would otherwise stop updating its Status labels and enabled states until the workbook is reopened, and this function lets the workbook find the tab again. The information it needs is kept in a hidden workbook name during the Excel session only, is removed before every save and never reaches the saved file. Nothing is installed, registered or loaded.


\Needspace{9\baselineskip}

## Check the delivered workbook

You receive `K-SACCR_1.3.0.xlsm` together with `SHA256SUMS.txt`, which holds one line per delivered file: its SHA-256 and its name.

1. Save both files in the same local folder.
2. Compute the workbook's SHA-256 in PowerShell, in that folder:

   ```powershell
   Get-FileHash .\K-SACCR_1.3.0.xlsm -Algorithm SHA256
   ```

   or in a Command Prompt: `certutil -hashfile K-SACCR_1.3.0.xlsm SHA256`.
3. Compare it with the line for `K-SACCR_1.3.0.xlsm` in `SHA256SUMS.txt` (letter case does not matter). If they differ, do not open the file: delete it and ask for a new copy.

   A match shows that the workbook arrived unchanged, not who made it: anyone able to replace the workbook on its way to you can replace `SHA256SUMS.txt` too. If both reached you through the same channel, confirm the expected SHA-256 with the supplier through a separate channel you trust, for example by telephone or from the supplier's own published release information, before you unblock the file and enable macros.
4. Windows marks a downloaded or e-mailed file as coming from the internet, and Office then blocks its macros. If your macro policy allows it, right-click the file → **Properties** → tick **Unblock** → **OK**, or run `Unblock-File .\K-SACCR_1.3.0.xlsm` in PowerShell. Do this only after the hash check.
5. Open the workbook in desktop Excel and enable macros as your policy allows. The workbook opens on `Home`, at a zoom of 100%, with the **K-SACCR** tab selected. Check that the tab has the groups Calculation, CSV import, Evidence, Parameters, Status and Navigate, and that its Status group shows the three status labels (`Results: NONE` before any run). Check that column A of `Home` shows the navigation bar. Open `Import_NettingSets` and `Import_Trades`, through the bar, **Go to** or the links on `Home`, and check that each has its four buttons.

A renamed copy (for example `K-SACCR_1.3.0 (1).xlsm`) has the same SHA-256 as long as its content is unchanged. Once you save the workbook, its SHA-256 changes: check it before the first save.


\Needspace{9\baselineskip}

## Keep a reference copy and work in a copy

Keep the delivered file unchanged as your reference copy. Use **File > Save As** to create a working copy, and keep business data out of the demonstration copy. If a working copy is damaged, for example after a header or named range was deleted, start again from the reference copy and enter your inputs again.

Do not change the workbook's VBA code, its worksheet structure, its named ranges, its Ribbon tab, its navigation bar or its buttons. A workbook whose code was changed is no longer the tested workbook. The protection of the sheets and of the workbook structure guards against such changes by accident, not against deliberate ones.


\Needspace{9\baselineskip}

## A new version, and removal

A new version of K-SACCR is delivered as a new workbook with its own file name, SHA-256 and manual. Check it as above and read any migration notes delivered with it. Then, in a copy of the new workbook:

1. Clear the contents of the demonstration rows on `NettingSets` and `Trades`, so that no demonstration value stays attached to your portfolio. Rows cannot be deleted on a protected sheet: clear their cells. The bank sheets are delivered empty.
2. Move your inputs by what identifies each value, never as one block, pasting values only (not formats or formulas): on `NettingSets`, `Trades`, `BankNettingSets` and `BankTrades` the data rows, column by column under the column of the same meaning, never the header row itself; on `Params` the run settings under the same `Code`, the FX rates in the FX table by the same currency code and the closing days in their table; and the reconciliation settings above the bank tables (`BankNS_Scale`, `BankNS_AbsTol`, `BankNS_RelTol`, `BankTR_Scale`, `BankTR_AbsTol`, `BankTR_RelTol`) by their names. A new version can add, rename or reorder columns and rows, and a value moved to the wrong place can still be accepted and give a wrong EAD. `NettingSets`, `Trades` and the bank figures can also be moved as CSV files.
3. Fill every column, code or key that is new in this version as its migration notes and manual say.
4. Run **Validate inputs**, then compare the results of the two versions on the same inputs before relying on the new one.

The other values on `Params`, the supervisory parameters and `SF_Table`, are the new workbook's reference values and are locked. Change one only if you changed it deliberately in the earlier version, for a what-if analysis, by the maintenance unlock described in *What you can edit: workbook protection*; the run is then classified `WHAT-IF`, unless the value changed is `DaycountConvention` or `IRBucketOffset`, conventions that do not change the classification. **Restore parameters** puts them back to the regulatory set at any time.

Do not copy VBA code or sheets between versions, and do not bring an earlier workbook up to date by changing it: a new version can change worksheets, headers and named ranges as well as code. Move the inputs into the new workbook instead.

Nothing is installed on the computer. Deleting the workbook and the CSV, PDF and XLSX files you exported or imported removes K-SACCR; keep any evidence your organization requires you to retain.


\Needspace{6\baselineskip}

### Moving inputs from 1.2.0

Version 1.3.0 changes the layout of every input sheet and the meaning of the FX rates. Moved as described here, a 1.2.0 portfolio gives the same EAD in 1.3.0, except a credit trade on the other regime's credit scale, which is now rejected. The differences that matter when moving are:

* **NettingSets.** The header is now on row 4 and the data start on row 5 (in 1.2.0, row 6 and row 7). The 17 columns `C` to `S` keep their order and meaning, under shorter headers: for example *Margined* for *Margined (Y/N)*, *Client clearing role* for *Client clearing role (N / CM / CLIENT / CCP; MULTI rejected)*, *Remargin freq. (BD)*, *Large or illiquid*, *Net VM* and *Threshold (TH)*. Column `T`, **Counterparty sector**, is new and optional: leave it blank, or enter a sector (see *NettingSets data contract*).
* **Trades.** The header stays on row 4 and the data on row 5; the 28 columns `C` to `AD` keep their order and meaning, under shorter headers such as *Hedging-set label*, *Notional CCY*, *Strike (K)*, *MtM leg* and *Settlement type*.
* **Credit scale.** Under CRR a single-name credit trade must use `CQS1` to `CQS6`, under BCBS `AAA` to `CCC`; a trade on the other scale is now rejected (`TR-FACTOR-REGIME`). Both scales give the same factors for the same steps, so re-enter such a trade with the step of its regime: for example `BBB` (0.54%) as `CQS3` under CRR.
* **FX rates: invert them.** In 1.2.0 a rate was the number of units of the reporting currency per 1 unit of the currency; in 1.3.0 it is the number of units of the currency per 1 unit of the reporting currency, as the market quotes it. Enter **1 divided by** each 1.2.0 rate, for example `1.0870` for a 1.2.0 USD rate of `0.92` with EUR as reporting currency, or press **Download prices**. The rate of the reporting currency stays 1. The workbook never converts a rate for you.
* **Day count.** `DaysPerYear` is gone. Its place is taken by `YearConvention`, the days of the year (365 by default), and `Calendar` (`TABLE` by default, with an empty table of closing days), which together give the calendar days over 365 of 1.2.0. Both are now run settings, in yellow, changed without a maintenance unlock. If you had changed `DaysPerYear` in 1.2.0, enter the same value as `YearConvention`.
* **The input fingerprint.** It covers the headers, the new column, the day count and the record of the FX download, so the same inputs give a different fingerprint (`inputs=` in the run summary and the evidence CSV) in 1.3.0 than in 1.2.0. Compare the two versions by their `Results`, not by their fingerprints.
* **Bank sheets.** The columns of `BankNettingSets` and `BankTrades` are unchanged. The default relative tolerance is now 1% (`0.01`) instead of 0.01%, and the absolute tolerance is 1 unit; move your own values if you set them.
* **CSV files.** NettingSets files of version 1, written for 1.2.0, import into 1.3.0 as they are, with the sector blank. Trades files are unchanged (version 1). The bank figures can now be moved as two CSV files too. `Params` is not imported: enter it as above.
* **What-if parameters.** If you changed reference parameters in the 1.2.0 workbook deliberately, make the same change again in the 1.3.0 workbook by a maintenance unlock; otherwise leave the delivered values, which **Restore parameters** can put back.

Either way, run both workbooks on the same inputs and compare their `Results`, as step 4 says.


\Needspace{12\baselineskip}

# First calculation: guided demonstration


\Needspace{9\baselineskip}

## What is included in the workbook

The delivered workbook contains a **synthetic demonstration portfolio** on `NettingSets` and `Trades`: the five published Basel CRE99 examples and additional illustrative netting sets, each with a counterparty sector, with the outputs cleared. Some illustrative rows are deliberately invalid to demonstrate the rejection mechanism. The bank sheets are empty. The demonstration rows do not always use the exact conventions of the published examples, as the next paragraph shows.

For example, CRE99 Example 3's published calculation uses a nine-month remaining maturity (`M = 0.75`); the demonstration row enters actual dates on a 365-day basis, giving `M = 0.7507` and an EAD around **5,405** in the example's printed thousand-unit scale rather than the published **5,406**. Using the stated nine-month period reproduces the published figure. Treat this as a **day-count illustration**, not a discrepancy requiring an arbitrary manual adjustment.


\Needspace{9\baselineskip}

## Walkthrough in Excel

1. **Preserve an untouched copy.** Save a separate working `.xlsm` from the delivered workbook. Use only synthetic information. The copy opens on `Home`; read it, and use the navigation bar on the left, its links or **Go to** on the K-SACCR tab to reach each sheet below.
2. **Inspect `Params`.** Find `AsOfDate`, `ReportingCcy`, `Regime`, `Alpha`, `BusinessDaysPerYear`, `YearConvention`, `Calendar`, `SF_Table` and the FX table; verify the reporting-currency FX row has rate 1 and that the other rates are quoted in units of the currency per 1 unit of the reporting currency.
3. **Inspect `NettingSets`.** Rest the pointer on a header to read its note. Note each ID, `Margined`, VM, NICA, threshold, MTA, any MPOR condition, `Regime override` and `Counterparty sector`.
4. **Inspect `Trades`.** Each row has a unique **TradeID** (the leg ID). Follow a sample row from `NettingSetID` to its asset class, risk factor, notional, currency and dates. If an economic trade has several legs, use the economic-trade columns.
5. On the K-SACCR tab, open the menu under **Run** and choose **Validate inputs** (`ValidateInputs`). The workbook opens `Checks`; read it. Errors describe excluded trades, unusable netting sets or a run-stopping defect; warnings describe accepted inputs needing attention. The Code and Field columns name each finding.
6. Click **Run** on the K-SACCR tab, which runs **Run SA-CCR** (`RunSACCR`). The status bar shows its progress. The macro validates, calculates and updates output sheets. It does **not** quietly compute a netting set's EAD using only valid trades when another trade in that set was rejected.
7. A completed run opens `Results` (a run that stops opens `Checks`). Locate the status for each netting set. Inspect `VALID`, `INCOMPLETE`, `INVALID` and `NO TRADES` behavior before focusing on numbers. The **CRR** and **BCBS** totals are separate, and they include only `VALID` netting sets.
8. Open the `Dashboard` for the run in one page: the totals per regime, the breakdowns and their charts.
9. For a valid example, trace its `TradeCalc` rows to `HedgingSets` / `Buckets` and the add-on columns on `Results`; verify the reported `RC`, multiplier, `PFE` and EAD.
10. Inspect `Checks` to understand all demonstration exclusions. Resolve exclusions on **copies**, never in your reference copy.
11. Optionally enter illustrative comparison figures on `BankNettingSets` and `BankTrades`, for example a netting set's EAD from `Results`; run again and study the tiles and result codes on `Reconciliation` and `TradeReconciliation`, and the sums on `ReconSummary`.
12. Click **Export evidence** on the K-SACCR tab, select an output CSV path and review the exported text. Click **Export report** and save a PDF. Repeat only after confirming freshness and output integrity.
13. Optionally, to see the effect of a reference parameter, change one on `Params` after a maintenance unlock (see *Maintenance unlock and relock*), run again and note the `WHAT-IF` classification and its warning on `Checks`; then press **Restore parameters** and run again.


\Needspace{9\baselineskip}

## Published Basel example reference points

The following are **published CRE99 target results**, not a promise that every demonstration row uses identical date assumptions. They are listed in original **USD units** here, whereas the source tables print USD thousands.

| Published case | Selected printed target in unscaled USD | What it helps check |
|:--|--:|:--|
| CRE99 Example 1 | EAD **569,000** | Interest-rate components, including option delta and add-on. |
| CRE99 Example 2 | EAD **381,000** | Credit add-on and multiplier (negative MtM, no collateral). |
| CRE99 Example 3 | EAD **5,406,000** | Commodity add-on and maturity precision. |
| CRE99 Example 4 | EAD **936,000** | Replacement cost and aggregation. |
| CRE99 Example 5 | EAD **1,879,000** | Margined MPOR (weekly remargining, 14 business days), aggregate add-on and multiplier. |

A reference comparison is meaningful only when **regime, reporting currency, unit scale, dates, day count, market inputs and netting terms** match the reference. Rounded values should be checked against their printed precision and published tolerance, not mistaken for machine-precision oracle values.


\Needspace{12\baselineskip}

# The K-SACCR tab and the navigation bar

Every command of the workbook is on its own Ribbon tab, **K-SACCR**, which also shows the status of the results and of both staging sheets and opens any sheet. The workbook opens with this tab selected. Press **Alt, K, S** to open it from the keyboard; Excel then shows the keytip of each control, and the keytip of a menu entry applies once its menu is open.

![The K-SACCR tab before any run: the Calculation, CSV import, Evidence, Parameters, Status and Navigate groups. Validate staging, Commit import and the two exports are grayed out, as there is nothing for them to act on.](assets/ribbon.png){width=98%}

* **This workbook only.** The tab is shown while the K-SACCR workbook is the active workbook, and every command acts on that workbook; another open workbook is never touched.
* **Selected at opening.** When the workbook opens, the tab is selected. An add-in that loads its own tab afterwards can take the selection away, so the workbook selects its tab once more about three seconds later, only while it is the active workbook; a tab you choose after that moment is left alone.
* **The same macros.** Each command runs the workbook macro named below, with all its checks: protection, the refusal while another operation runs, error handling, and every condition of a staging validation, a commit or an export. The enabled states of the tab help you; they never replace those checks.
* **Icons.** Each command has an icon of its own, a white sign on a navy tile, the color of the table headers; the entries of the Import CSV, Templates and Go to menus have one too, one sign for each kind of file or sheet.
* **Macros disabled.** The tab is still shown, but nothing on it works: the commands, the Status labels, the enabled states, **Go to** and **Navigation bar** all need macros. The workbook stays readable, and the navigation bar, the links on `Home` and the sheet tabs reach every sheet.

The labels of the large buttons are short; the full name is in the screentip that Excel shows when you rest the pointer on a control.


\Needspace{9\baselineskip}

## Ribbon map

| Group | Control (keytip) | What it does |
|:---|:-------|:----------|
| Calculation | **Run** (R), a split button | Its button runs **Run SA-CCR**; the arrow opens a menu with **Run SA-CCR** (R) and **Validate inputs** (V). |
|  | **Run SA-CCR** (R) | Validate the inputs, calculate the accepted netting sets and populate every output sheet, the Dashboard included; update the reconciliations where possible. A completed run opens `Results` (its message points to `Checks` when there are errors or warnings); a run that stops opens `Checks`. |
|  | **Validate inputs** (V) | Examine the inputs and write the findings to `Checks`, which it opens, **without calculating exposure**. It clears the previous run's outputs and reconciliations, so press **Run SA-CCR** again before reviewing or exporting. |
|  | **Clear outputs** (C), a button of its own | Clear the calculated output areas, the Dashboard and both reconciliations and invalidate the previous completed run; the inputs, `Params` and the bank figures stay, and the sheet you are on stays active. Do not assume that an error while clearing leaves no residual values. |
| CSV import | **Import CSV** (I), a menu: **NettingSets file...** (N), **Trades file...** (T) | Load a CSV file you choose onto `Import_NettingSets` or `Import_Trades` as text, for checking and correction; the input sheets are not changed. |
|  | **Import CSV** menu, *Bank figures*: **BankNettingSets file...** (B), **BankTrades file...** (L) | Replace the figures of `BankNettingSets` or `BankTrades` with the records of a CSV file, or change nothing and list what is wrong with the file. There is no staging sheet. See *Load the bank figures from a CSV file*. |
|  | **Import CSV** menu: **Clear BankNettingSets...** (C), **Clear BankTrades...** (R) | After a confirmation, clear every figure and comment of that bank sheet; the amount scale and the tolerances stay. |
|  | **Validate staging** (V) | Check the staged rows of both staging sheets, against each other, the input sheets and Params; list the findings on Checks with the origin `IMPORT`. The input sheets are not changed. |
|  | **Commit import** (M) | After a confirmation, replace `NettingSets`, `Trades` or both with the staged rows of a validation without findings that still holds. Calculates nothing. |
|  | **Templates** (P), a menu: **NettingSets template...** (N), **Trades template...** (T), **BankNettingSets template...** (B), **BankTrades template...** (L) | Write the blank CSV file of that sheet, its header line alone, to a file you choose. |
|  | **Templates** menu: **Import specification...** (S) | Save the description of the four CSV files, with an example of each, as one HTML page. Changes nothing in the workbook. |
| Evidence | **Export evidence** (E) | Interactive CSV export of a **current and unmodified** run; refused if a freshness or integrity requirement fails. |
|  | **Export report** (O) | A formatted report of the same run as a PDF or an XLSX workbook, refused in the same cases. See *The formatted run report*. |
| Parameters | **Restore parameters** (A) | After a confirmation, put the reference parameters and the supervisory-factor table on `Params` back to the regulatory set; see *Restore the regulatory parameters*. |
| Status | Three labels | Show the status of `Results` and of the two staging sheets; they only read cells. |
| Navigate | **Go to** (G), a split button | Its button opens `Home`; the arrow lists every sheet. |
|  | **Navigation bar** (B), a toggle | Hide or show the navigation bar on every sheet; pressed while the bar is shown. See *The navigation bar*. |

The macros behind the controls, for the Macros dialog (Alt+F8) and for automation:

| Control | Macro |
|:----------|:----------|
| **Run** (R), a split button | `RunSACCR` |
| **Run SA-CCR** (R) | `RunSACCR` |
| **Validate inputs** (V) | `ValidateInputs` |
| **Clear outputs** (C), a button of its own | `ClearOutputs` |
| **Import CSV** (I), a menu: **NettingSets file...** (N), **Trades file...** (T) | `ImportNettingSetsCsv`, `ImportTradesCsv` |
| **Import CSV** menu, *Bank figures*: **BankNettingSets file...** (B), **BankTrades file...** (L) | `ImportBankNettingSetsCsv`, `ImportBankTradesCsv` |
| **Import CSV** menu: **Clear BankNettingSets...** (C), **Clear BankTrades...** (R) | `ClearBankNettingSets`, `ClearBankTrades` |
| **Validate staging** (V) | `ValidateStaging` |
| **Commit import** (M) | `CommitImport` |
| **Templates** (P), a menu: **NettingSets template...** (N), **Trades template...** (T), **BankNettingSets template...** (B), **BankTrades template...** (L) | `ExportNettingSetsCsvTemplate`, `ExportTradesCsvTemplate`, `ExportBankNettingSetsCsvTemplate`, `ExportBankTradesCsvTemplate` |
| **Templates** menu: **Import specification...** (S) | `ExportImportSpecification` |
| **Export evidence** (E) | `ExportEvidence` |
| **Export report** (O) | `ExportReport` |
| **Restore parameters** (A) | `RestoreRegulatoryParameters` |
| **Navigation bar** (B), a toggle | `ToggleNavigationBar` |

The **Templates** menu writes blank CSV files or the import specification, which describe input files; **Export evidence** writes the evidence CSV of a run, and **Export report** a readable report of it. They are different operations.

After running, `Results!C4` contains the run summary, and `Results!C5` the freshness status. Reconciliation sheets likewise have summary (`C4`) and status (`C5`) positions. Each staging sheet shows its source file, load time, record count and status in `D4:D7`.


\Needspace{9\baselineskip}

## Enabled states

A control the tab disables is grayed out. The tab decides from the same cells the Status labels read, and never calculates or writes to do so:

| Control | Enabled when |
|:----|:--------|
| Every command and the **Navigation bar** toggle | No operation is running. While one runs, every command is disabled. |
| **Validate staging** | A staging sheet holds rows: its status is not `EMPTY`. |
| **Commit import** | A staging sheet reads `VALID`. |
| **Export evidence**, **Export report** | `Results!C5` says that the inputs are unchanged since the last run: it reads neither `OUT OF DATE` nor that there is no run. The two exports refuse in the same cases. |
| Every other control | Always, apart from a running operation. |

**Go to**, its entries and the Status labels have no enabled state and stay available, also while an operation runs. A status that cannot be read leaves its control enabled: a command is never disabled on a guess, and the macro then decides. An enabled command can therefore still refuse, with its own message: **Commit import** refuses unless the last **Validate staging** found no finding and nothing it judged has changed since, and the two exports refuse edited outputs, a staging validation on `Checks` or a stale reconciliation, which the tab does not detect (see *Why an export can be refused*).


\Needspace{9\baselineskip}

## Status labels

The Status group has three read-only labels. They describe the state in words; the controls carry no color.

| Label | Reads | Shows |
|:----|:----|:--------|
| **Results** | `Results!C5` | `Results: CURRENT` when the inputs are unchanged since the last run, `Results: OUT OF DATE` when they changed, `Results: NONE` before any run, `Results: UNKNOWN` for any other text in the cell. |
| **NettingSets staging** | `Import_NettingSets!D7` | `NettingSets staging:` followed by the status: `EMPTY`, `LOADED`, `EDITED`, `VALID`, `INVALID`, `OUT OF DATE` or `COMMITTED`. |
| **Trades staging** | `Import_Trades!D7` | `Trades staging:` followed by the same statuses. |

`Results: CURRENT` follows `Results!C5`, which checks the inputs only. An output edited after the run, which `ResultsStatus()` reports as `EDITED`, is not shown on the tab; the two exports still refuse it on their own (see *Freshness model: when is a result current?*).


\Needspace{9\baselineskip}

## Go to: reaching every sheet

The button of **Go to** opens `Home`. Its arrow lists every sheet, in the order of the sheet tabs, in six titled sections:

| Section | Sheets |
|:----|:--------|
| Start | `Home` |
| Inputs | `Params`, `NettingSets`, `Trades` |
| Staging | `Import_NettingSets`, `Import_Trades` |
| Outputs | `Dashboard`, `Results`, `HedgingSets`, `Buckets`, `TradeCalc`, `Checks` |
| Reconciliation | `BankNettingSets`, `Reconciliation`, `BankTrades`, `TradeReconciliation`, `ReconSummary` |
| Reference and tests | `TestCatalogue`, `FormulaTest`, `Basel_vs_CRR`, `README` |

`Home` links to every other sheet, grouped the same way. Its links are ordinary worksheet hyperlinks, so they also work with macros disabled.


\Needspace{9\baselineskip}

## The navigation bar

Column A of every sheet holds the **navigation bar**: a link to every sheet, in the six sections of the **Go to** menu, in its order, with the same labels and tooltips. Its links are shapes linked to the sheets, so they work on the protected sheets and with macros disabled.

* **Home at the top.** The *Start* section (`Home`) and the *Inputs* section (`Params`, `NettingSets`, `Trades`) form the top block of the bar. It stands in the frozen rows at the top of the sheet, so it stays in view whichever way the sheet scrolls.
* **The rest scrolls with the rows.** The other sections, *Staging*, *Outputs*, *Reconciliation* and *Reference and tests*, follow from the first row below the frozen rows. That part of the bar **scrolls with the rows**: on a long sheet, scroll back to the top to reach it.
* **The current sheet** is highlighted and is not a link.
* **Not printed.** The bar is not printed, and its cells are locked.
* **Same place on every sheet.** Columns A and B are frozen on every sheet, so the bar stays in view when a wide sheet scrolls to the right. Every sheet has the same title band, and sheets of the same kind freeze rows of the same height, so the view does not jump when you change sheet.

**Hiding the bar.** Wide sheets such as `Trades` and `TradeCalc` get the space back when the bar is hidden. Two controls do this, on every sheet at once:

* **Navigation bar** in the Navigate group of the K-SACCR tab, a toggle that is pressed while the bar is shown;
* the small triangle at the far left of `B1`, in the title band of every sheet: **◀** while the bar is shown hides it, **▶** while it is hidden shows it again, in the same place.

Only column A is hidden; column B, with the triangle, stays in view. The sheets stay protected, and hiding or showing the bar does not mark the workbook as changed, so it never asks you to save. Hiding lasts for the Excel session only: the workbook is always saved with the bar shown, so a copy opened later, also with macros disabled, has its bar. Both controls need macros; with macros disabled the triangles do nothing.


\Needspace{9\baselineskip}

## When the tab updates

The Status labels and the enabled states are read again:

* after every operation, once Excel is idle, whether it was started from the tab, from a button on a sheet or from the Macros dialog, and after a change on a staging sheet;
* after the navigation bar is shown or hidden;
* after any cell edit in the workbook, since an input edit changes `Results!C5`;
* when the workbook or one of its sheets is activated, so switching back from another workbook refreshes the tab;
* when the workbook is opened again.

The tab is refreshed only while the K-SACCR workbook is the active workbook; it is refreshed when you come back to it. With calculation set to manual, `Results!C5` changes only when Excel recalculates: press **F9**, then select another sheet, and the label follows.

**After a VBA error.** An unhandled VBA error, or an edit of the code in the Visual Basic editor, resets the workbook's code and with it the link the workbook keeps to its tab. The workbook then restores that link through the one Windows function described in *Supported environment*, so that the labels and states keep updating. If that is not possible, the tab simply stops updating its labels and states until the workbook is closed and opened again; its commands still work, and each still checks its own conditions. See *Diagnostic table*.


\Needspace{9\baselineskip}

## Progress in the status bar

The long commands say what they are doing in Excel's status bar, at the bottom of the window, while they run: the command, a bar of 20 blocks, one per 5% done, the percentage and the step, for example:

```text
K-SACCR - Import Trades  ███████░░░░░░░░░░░░░  36%  Trades.csv: checking record 14000 of 40000
```

| Command | Steps shown |
|:----|:--------|
| **Run** and **Validate inputs** | Withdrawing the previous outputs; the calculation, 80% of the bar: loading the inputs, each trade, the hedging sets, the netting sets, Buckets and HedgingSets, Checks; then the reconciliations and the Dashboard. |
| **Import CSV**, for the four files | Reading and checking the file, record by record, 90% of the bar; then writing its records to the sheet. |
| **Validate staging** | The rows of `Import_NettingSets`, then of `Import_Trades`, row by row; the legs against their netting sets; Params; writing Checks and marking the cells. |
| **Commit import** | Reading the staged rows; writing each input sheet; formatting its rows, column by column; reading them back. |
| **Export evidence** | Checking the run, then writing the evidence. |
| **Export report** | One step per section of the report and the save; the Charts step names the chart being pictured. |
| `CheckNettingSetsCsv`, `CheckTradesCsv` | Reading and checking the file, record by record. |

* The bar reaches 100% when the command is done. It stays while the final message is shown, and the status bar is given back to Excel when you close the message; a failure gives it back too.
* While the command runs, Excel takes no keyboard or mouse input, so that no cell can be edited while a sheet is being written. Input is given back before every message and question. A command started meanwhile is refused.
* The automation entry points (see *Automation and worksheet functions*) show no progress.


\Needspace{9\baselineskip}

## Buttons on the sheets

Some sheets carry buttons of their own, rounded shapes in the title band, right after the title: navy with white text for a command, light blue with navy text for the secondary one. They run the same macros as the tab and work on the protected sheets.

| Sheet | Buttons (macro) | What they do |
|:----|:--------|:--------|
| `Params` | **Download prices** (`DownloadClosingPrices`) | Write the rates of the FX table from the source chosen beside the button; see *Download prices*. Not on the tab. |
| `Import_NettingSets`, `Import_Trades` | **Import CSV** (`ImportNettingSetsCsv` or `ImportTradesCsv`), **Validate staging** (`ValidateStaging`), **Commit import** (`CommitImport`), **Export import specifications** (`ExportImportSpecification`) | As on the tab: load a CSV file onto that staging sheet, check the staged rows of both sheets, commit them, and save the import specification page. |
| `Results`, `HedgingSets`, `Buckets`, `TradeCalc`, `Checks`, `Reconciliation`, `TradeReconciliation` | **Clear outputs** (`ClearOutputs`) | As on the tab: clear every output sheet, not only the one you are on. |
| `BankNettingSets`, `BankTrades` | **Import CSV** (`ImportBankNettingSetsCsv` or `ImportBankTradesCsv`), **Export CSV template** (`ExportBankNettingSetsCsvTemplate` or `ExportBankTradesCsvTemplate`), **Clear bank figures** (`ClearBankNettingSets` or `ClearBankTrades`) | As in the Import CSV and Templates menus of the tab, for that sheet. |
| `FormulaTest` | **Randomize inputs** (`RandomizeFormulaTestInputs`), **Restore defaults** (`RestoreFormulaTestDefaults`) | Draw reproducible random inputs for the worksheet-function tests, or write the delivered inputs back; see *Worksheet function API*. Not on the tab. |

Three macros are run from Excel's Macros dialog (**Alt+F8**, select the macro, **Run**): `CheckNettingSetsCsv` and `CheckTradesCsv`, which check a CSV file without loading it, and `ProtectWorkbook`, which protects the sheets again after a maintenance unlock.


\Needspace{12\baselineskip}

# What you can edit: workbook protection


\Needspace{9\baselineskip}

## Protection in brief

Every sheet of the workbook is protected, without a password, except the two reconciliation sheets, and so is the workbook structure. Only the cells listed below can be changed in normal use: the inputs, the bank figures, the run configuration on `Params` and the test inputs of the two test sheets. Every other cell is locked: formulas, results, headings, instructions, layout and the reference regulatory parameters.

* **Yellow means editable.** No locked cell has the yellow input fill. On `Params`, the reference regulatory parameters are filled grey.
* **Formulas stay visible.** No cell is hidden: a locked cell, a formula included, can still be selected, read and copied.
* **Saved protected.** The workbook is delivered with every sheet but `Reconciliation` and `TradeReconciliation`, and the structure, protected, so the protection also holds with macros disabled. Nothing else works until macros are enabled.
* **Macros write, you edit inputs.** When the workbook opens with macros enabled, and when an operation starts, such as **Run SA-CCR**, **Validate inputs**, **Clear outputs**, **Import CSV**, **Validate staging**, **Commit import**, **Restore parameters** or the **Navigation bar** toggle, the workbook protects every sheet and the structure again, in a way that lets the macros write every sheet while you change only the unlocked cells. A sheet you unprotected, or protected by hand with other options, is protected again with the workbook's options at that moment.
* **Structure.** Sheets cannot be added, deleted, renamed, moved, hidden or shown.
* **Restoring parameters needs no unlock.** **Restore parameters** on the K-SACCR tab writes the regulatory set back to the locked reference parameters through the macros' protection, after a confirmation (see *Restore the regulatory parameters*). Only a change away from the regulatory set needs a maintenance unlock.

**Protection is not security.** It prevents accidental edits only. Anyone can remove it with **Review > Unprotect Sheet**, as no password is set. It does not replace the input validation, the freshness status of `Results`, the output fingerprint or the guards of the two exports, which still tell whether results match their inputs.


\Needspace{9\baselineskip}

## What each sheet lets you change

| Sheet | Cells you can change | Also allowed |
|:----|:--------|:----|
| `Home` | None: it holds no input | None |
| `Params` | `E6` (`AsOfDate`), `E7` (`ReportingCcy`), `E8` (`Regime`), `E26` (`YearConvention`), `E27` (`Calendar`), `D59` (the source of **Download prices**), `C61:E260` (the FX table, with rows below it for added currencies) and `I61:J160` (the table of closing days) | None |
| `NettingSets` | Columns `C:T` from row 5 down | Sorting, inserting rows |
| `Trades` | Columns `C:AD` from row 5 down | Sorting, filtering, inserting rows |
| `Import_NettingSets` | Columns `D:U` from row 10 down: the staged values | Inserting rows |
| `Import_Trades` | Columns `D:AE` from row 10 down: the staged values | Inserting rows |
| `Dashboard` | None | None |
| `Results`, `HedgingSets`, `Buckets`, `TradeCalc`, `Checks` | None | Filtering |
| `BankNettingSets` | `D4:D6` (amount scale and tolerances) and columns `C:P` from row 9 down | Sorting, inserting rows |
| `Reconciliation`, `TradeReconciliation` | Not protected | Anything; see below |
| `BankTrades` | `D4:D6` and columns `C:O` from row 9 down | Sorting, inserting rows |
| `ReconSummary` | None: it holds formulas only | None |
| `TestCatalogue` | `L7:L34`, the expected values | None |
| `FormulaTest` | The tolerance `E4`, the seed `H4`, the test inputs `E7:F28`, `G9:G28`, `H13:J21`, `K15`, and the range table `G33:H81`, `K33:L81` | None |
| `Basel_vs_CRR` | None | Filtering |
| `README` | None | Opening and closing row groups |

The sheets are listed in the order of their tabs. `Home` is protected like `README`; its links to the other sheets work all the same, also with macros disabled. On every sheet the navigation bar's cells are locked and its shapes are protected with the sheet; its links still work.

The input columns are unlocked down to the last row of the sheet, not only in the rows that hold data, so a pasted or appended row is editable however many rows the portfolio has; the header rows above them stay locked. On the staging sheets, the `CSV line` column (column C), the status, the source file and the other information above the staged rows stay locked.

**The two reconciliation sheets are working sheets.** `Reconciliation` and `TradeReconciliation` are never protected: you can sort and filter their tables, write notes and add columns to the right of the table. A run, **Validate inputs** and **Clear outputs** rewrite only the table. The table is part of what the run recorded, so sorting it or changing a cell of it makes `ResultsStatus()` read `EDITED`, and the two exports refuse, until the next run; filtering changes no cell.

Selecting and copying any cell is always allowed. These are **refused** on a protected sheet:

* typing into, pasting into or clearing a locked cell, such as a header, a grey `Params` value, an output cell or the status of a staging sheet;
* deleting rows or columns, inserting columns, and resizing or formatting columns or cells. To remove an input row, clear its cells;
* filtering on a protected sheet other than `Trades`, `Basel_vs_CRR` and the five output sheets `Results`, `HedgingSets`, `Buckets`, `TradeCalc` and `Checks`, which keep the AutoFilter the workbook is delivered with; a protected sheet cannot get a new AutoFilter. On the output sheets the filter covers the table down to the last row of the sheet, and a run, **Validate inputs** and **Clear outputs** keep the filter buttons and apply a filter you set again to the new rows;
* adding, deleting, renaming, moving, hiding or showing a sheet.

Excel itself refuses these edits, with its own message.


\Needspace{9\baselineskip}

## Maintenance unlock and relock

A locked cell is changed only on purpose. To change one, for example a supervisory parameter on `Params` for a what-if analysis:

1. Work in a copy, and write down the cell, its old value and the reason.
2. Select the sheet and choose **Review > Unprotect Sheet**. No password is asked.
3. Make the change.
4. Protect the sheet again: run the macro `ProtectWorkbook` (**Alt+F8**, select it, **Run**), or simply start the next operation, such as **Run SA-CCR**, which protects every sheet and the structure again first.

The change itself is kept. A run on changed reference parameters is classified **`WHAT-IF`**, except for the conventions `DaycountConvention` and `IRBucketOffset`, with a *What-if:* warning on `Checks` for each deviation (see *Regulatory versus what-if classification*); the freshness status and the evidence checks work as before. An output cell changed during an unlock makes `ResultsStatus()` read `EDITED` and the two exports refuse until the next run, which rewrites the outputs. To return to the reference values, press **Restore parameters** on the K-SACCR tab, which needs no unlock; a factor row you added below the supervisory-factor table must first be removed by hand, after a maintenance unlock, since the restore refuses while it is there.

Do not unlock a sheet to delete rows, move columns or rename headers: the input tables are found by their workbook names and headers, and a changed layout can stop a run (see *Dates, named ranges and error handling*).


\Needspace{12\baselineskip}

# Global settings and supervisory data


\Needspace{9\baselineskip}

## Params: identity and values

The `Params` sheet is a **named-range-driven table**. Each parameter has a code in the `Code` column, a value in the `Value` column and a workbook name of that code. The workbook also uses the names `Params_Code`, `Params_Value`, `SF_Table`, `FX_Table`, `FX_Source` and `Holiday_Table`. Names—not fixed row positions—locate the tables; their first header cells must be structurally correct. The header is on row 5; rows 1 to 4 and the `Code` column stay in view when the sheet scrolls.

`AsOfDate`, `ReportingCcy`, `Regime`, `YearConvention`, `Calendar`, the FX table with the source of **Download prices**, and the table of closing days are the **run configuration**: yellow, and changed in normal use. Every other value on `Params`, the settings from `Alpha` to `DaycountConvention` and `SF_Table`, is a **reference regulatory parameter**: grey and locked, changed only by a deliberate maintenance unlock (see *Maintenance unlock and relock*), after which a run on a changed value is `WHAT-IF`; `DaycountConvention` and `IRBucketOffset` are the exceptions, conventions that do not change the classification (see *Regulatory versus what-if classification*). **Restore parameters** puts every one of them, the two conventions included, back to the regulatory value shown below (see *Restore the regulatory parameters*).

| Setting | Typical regulatory value / convention | Meaning and user action |
|:--|:--|:--|
| `AsOfDate` | User-supplied valid reporting date | All dates and exposures are referenced to this date. |
| `ReportingCcy` | Three-letter code, e.g. `EUR` | All output amounts are in this currency. Its FX table rate must be **1**; a change offers to rebase the table (see *FX_Table: conversion into reporting currency*). |
| `Regime` | `CRR` by default; `BCBS` possible | Default rule set, with optional netting-set override. |
| `Alpha` | **1.4** | Supervisory EAD multiplier unless explicitly overridden for a netting set. |
| `MultiplierFloor` | **0.05** | Lower bound on the PFE multiplier. |
| `BusinessDaysPerYear` | **250** | Business-day convention for maturity floors and MPOR. |
| `MinMaturityBD` | **10** | Floor on the unmargined maturity factor's M. |
| `SDFloorBD` | **10** | Floor used in supervisory-duration treatment. |
| `MPORFloorBilateral` | **10** | Base bilateral MPOR floor. |
| `MPORFloorCleared` | **5** | Eligible client-clearing MPOR floor, subject to regime/role. |
| `MPORFloorLarge` | **20** | Large / illiquid netting-set floor. |
| `BasisFactor` | **0.5** | Factor for eligible basis hedging sets. |
| `VolatilityFactor` | **5** | Factor for volatility transactions. |
| `IRCorrBucket12`, `IRCorrBucket23`, `IRCorrBucket13` | **0.7**, **0.7**, **0.3** | Interest-rate maturity-bucket correlations. |
| `IRBucketOffset` | `TRUE` for offset formula | `FALSE` selects sum-of-absolutes **only for BCBS**; invalid for a CRR netting set with IR trades. |
| `LambdaThresholdIR` | **0.001** | CRR lambda threshold for interest-rate options. |
| `LambdaThresholdCO` | **0.1** | CRR lambda threshold for commodity options. |
| `DaycountConvention` | `ACTUAL` | Numerator of the day count, a convention: the days between the reporting date and a date, counted on the calendar. `ACTUAL` is the only value. |
| `YearConvention` | **365** (run setting) | Denominator of the day count: the days of the conventional year, 366, 365, 364, 360, 252 or 250. See *Dates and the day count*. |
| `Calendar` | `TABLE` (run setting) | Which days are counted: `TABLE`, `TARGET` or `TARGET+TABLE`. See *Dates and the day count*. |

**Unit discipline.** Rates, correlations and percentages are decimal values (`0.05` means 5%); notional, valuation and collateral are unscaled currency units; business-day settings are days; S/E/M/T are year fractions. Do not enter `5` to represent 5% unless a particular field expressly requires it. Missing optional numeric settings may be replaced by their regulatory defaults **with a warning**; invalid mandatory settings stop the calculation.

![The top of Params: the run settings in yellow and the reference regulatory parameters in grey, each with its description and its CRE52 or CRR reference.](assets/params.png){width=70%}


\Needspace{9\baselineskip}

## Regulatory versus what-if classification

The workbook records the regulatory reference set **`CRE52-CRR-2026-10-09`**. The engine compares the relevant Params numeric fields, supervisory-factor rows (their regimes included) and per-netting-set alpha overrides to the set. If they all match, the run is `REGULATORY`; otherwise `WHAT-IF`, with a **separate “What-if:” warning** in Checks for each deviation; the warning of a factor row names its regimes and the regulatory ones. `DaycountConvention` and `IRBucketOffset` are treated as conventions, and `YearConvention` and `Calendar` as run settings, like the reporting date: they do not by themselves change the classification; nevertheless they affect calculations and must be documented. A `REGULATORY` class is not a judgment that the user entered correct data. Since the reference values are locked, a `WHAT-IF` run on reference parameters follows a deliberate maintenance unlock; an alpha override on `NettingSets` is an input and needs none. **Restore parameters** makes the parameters and the factor table match the set again, so that the next run on them is `REGULATORY`; it does not change an alpha override, so a run with one stays `WHAT-IF`.


\Needspace{9\baselineskip}

## Restore the regulatory parameters

**Restore parameters**, in the Parameters group of the K-SACCR tab (macro `RestoreRegulatoryParameters`), writes the regulatory set **`CRE52-CRR-2026-10-09`** back to `Params`. Its values come from the same definition of the set that every run is compared with.

| | What it covers |
|:----|:--------|
| **Restored** | The reference parameters `Alpha`, `MultiplierFloor`, `BusinessDaysPerYear`, `MinMaturityBD`, `SDFloorBD`, `MPORFloorBilateral`, `MPORFloorCleared`, `MPORFloorLarge`, `BasisFactor`, `VolatilityFactor`, `IRCorrBucket12`, `IRCorrBucket23`, `IRCorrBucket13`, `LambdaThresholdIR` and `LambdaThresholdCO`; the conventions `DaycountConvention` (`ACTUAL`) and `IRBucketOffset` (`TRUE`); and the 26 rows of the supervisory-factor table, in its columns from `SF_Key` to `Regimes`. |
| **Not changed** | `AsOfDate`, `ReportingCcy`, `Regime`, `YearConvention`, `Calendar`, the FX table and the source of its rates, the table of closing days, the descriptive `Reference` column of the factor table, and the alpha overrides on `NettingSets`, which are inputs. |

1. **Confirmation.** The command lists every value it will overwrite, with its value now and in the set, and says what it leaves unchanged. **Cancel** is the default, and cancelling changes nothing.
2. **Nothing to restore.** When `Params` already holds the set, the command says so and changes nothing.
3. **Write.** On **OK**, it writes only the cells that differ from the set, through the macros' protection: no maintenance unlock is needed. It reads each cell back. If a write fails, it puts `Params` back as it was and checks it, as **Commit import** does with the input sheets, and its message says whether that succeeded; if `Params` could not be put back, check it before any run.
4. **Refusal.** A factor row that stands below the 26 rows of the factor table, right below them or further down above the FX table, refuses the restore, which would leave that row in the table: remove the row by hand after a maintenance unlock, then restore again. A `Params` workbook name that is missing or broken refuses it too. A refusal writes nothing and gives the reason.
5. **Effect.** The message gives the number of parameters and factor rows changed. Nothing is recalculated: `Results` reads **OUT OF DATE**, and the two exports refuse, until the next **Run SA-CCR**.


\Needspace{9\baselineskip}

## SF_Table: supervisory-factor register

The supervisory lookup contains the following columns, in this order: `SF_Key`; `Asset class`; `Category`; `Supervisory factor`; `Correlation`; `Supervisory option vol`; `Commodity hedging set`; `Regimes`, followed by a descriptive `Reference` column that the engine does not read. A key's classification and its regimes (`BOTH`, `CRR` or `BCBS`) determine what the engine may use. The table extends from its named header to the **first blank key**, including rows below the originally named range. Do not leave a blank key in the middle of the intended data.

Selected **regulatory** reference values are shown below. This is a navigation aid, not permission to overwrite the workbook's complete table without an independent review.

| Class/category | Supervisory factor | Correlation / supervisory volatility where relevant |
|:--|--:|:--|
| IR | 0.50% | Supervisory option volatility 50%. |
| FX | 4.00% | Supervisory option volatility 15%. |
| CR single name | 0.38% to 6.00% | By Basel rating (BCBS only) or CRR credit quality step (CRR only); single-name correlation 50%, option vol 100%. |
| CR index | 0.38% (IG), 1.06% (SG) | Index correlation 80%, option vol 80%; both regimes. |
| EQ single name | 32% | Correlation 50%, option vol 120%. |
| EQ index | 20% | Correlation 80%, option vol 75%. |
| CO electricity | 40% | Commodity correlation 40%, option vol 150%. |
| CO other categories | 18% | Commodity correlation 40%, option vol 70% for non-electricity commodities. |
| CO climatic conditions | 18% | **CRR-only** hedging set. |
| OT other risks | 8% | **CRR-only**; supervisory option volatility 150%. |

**Each regime's credit scale.** The rows `CR_AAA` to `CR_CCC` are the Basel ratings of CRE52.72 and apply under **BCBS only**; the rows `CR_CQS1` to `CR_CQS6` are the credit quality steps of CRR Article 280c(5), Table 3, and apply under **CRR only**. The factors are: CRR CQS1 through CQS6, **0.38%, 0.42%, 0.54%, 1.06%, 1.60%, 6.00%**; Basel AAA/AA/A/BBB/BB/B/CCC, **0.38%, 0.38%, 0.42%, 0.54%, 1.06%, 1.60%, 6.00%**. A single-name credit trade on the other regime's scale is rejected (`TR-FACTOR-REGIME` on its sub-class). The index rows `IG_INDEX` and `SG_INDEX` apply under both regimes. A bank must apply the appropriate rating or credit-quality-step mapping and report its methodology; K-SACCR does not fetch ratings.


\Needspace{9\baselineskip}

## FX_Table: conversion into reporting currency

The FX table has the header fields `Currency Code`, a descriptive `Currency Description` column (not read by the engine), and `Units of ccy per 1 unit of reporting ccy`. The rates are quoted **as the market quotes them**: enter the number of units of the currency per **one unit of the reporting currency**. Example: when reporting in EUR, a EUR/USD rate of 1.16 is entered as `USD -> 1.16`. An amount in a currency is converted into the reporting currency by **dividing** it by the currency's rate. The rate of the reporting currency itself **must be 1**. The code and direction matter: entering reciprocal quotes changes amounts. 1.2.0 used the reciprocal quote; see *Moving inputs from 1.2.0*.

**Example (illustrative only):** with EUR as reporting currency and a USD rate of 1.16, a USD 1,000,000 trade amount converts into EUR 862,068.97 (1,000,000 / 1.16). This is a **currency conversion example**, not an SA-CCR EAD prediction. The engine converts only notionals and MtM values: option prices and strikes are ratios, and collateral is entered in the reporting currency.

The delivered table holds the euro and the 20 currencies most traded against it, with illustrative rates. Add a currency directly below the last row, with no blank row in the `Currency Code` column; the table's validation covers its first 21 rows, and up to 200 rows below the header can be edited. For **FX derivatives**, two currencies must both be present in the FX table and the entered risk-factor currency pair must be coherent with the notional currency. The special pair-order rules for options and linear FX are detailed later in this manual.

**Changing the reporting currency.** The rates are relative to the reporting currency, so after you change `ReportingCcy` they are still those of the old one, and a run would stop: the new reporting currency's rate is not 1. The workbook therefore asks, with **No** as the default, whether to **rebase** the table on the new currency. **Yes** divides every rate by the rate of the new reporting currency, which becomes 1; every conversion stays the same, and the date and source of the rates stay too. A rate that is a formula, or not a number above zero, is left as it is and listed; a formula as the rate of the new reporting currency refuses the rebase. **Undo** cannot bring the rates back. Without the rebase, a run stops with `FX-REPORTING-RATE`, whose message says what to do: accept the rebase, press **Download prices**, or divide every rate by the rate of the reporting currency yourself. The question is not asked while an operation runs.


\Needspace{6\baselineskip}

### Download prices

**Download prices**, a button beside the FX table, writes the rates of the table from the source chosen in the list beside it (`D59`):

| Source | What it writes |
|:--|:--|
| `Closing price` (the default) | The closing price of each currency pair against the reporting currency, looked up with Excel's `STOCKHISTORY`: the last closing price up to the reporting date, seven days back at most. Needs Microsoft 365 and a connection. |
| `ECB reference rate` | The euro foreign exchange reference rate of the European Central Bank: for each currency, the last fixing on or before the reporting date, as published. With another reporting currency, the rate is crossed through the euro. The workbook sends one HTTPS request to the ECB data portal, through Windows' own connection and proxy settings, carrying only the three-letter codes of the table and the seven days up to the reporting date. |

* **Confirmation.** The command first asks, naming the source, with **Cancel** as the default.
* **What changes.** The rates are written as values; no formula, web query or connection stays in the workbook. A currency that gets no rate keeps its own and is listed, for example a currency the ECB does not publish. The rate of the reporting currency is set to 1. A currency added below the table is looked up too. The source and the date of the rates are shown beside the button, for example *ECB reference rates of 2026-09-30* or *Closing prices of 2026-09-30*.
* **Waiting.** Excel hands the closing prices back only when no macro runs, so the button starts the lookup and returns; the rates are written when the prices have come, within 30 seconds, and the status bar shows the wait. Saving or closing the workbook meanwhile ends the lookup without changing a rate. The ECB answer is awaited for up to 30 seconds.
* **Refusals.** Nothing is changed, and the message gives the reason, when the source cannot be used: an Excel without `STOCKHISTORY`; no connection, a time-out, an HTTP error, or an answer that is not the one expected, such as a proxy's sign-in page; no rate for the reporting currency, or for any currency; or a reporting date, reporting currency or currency code changed while the prices were awaited. The ECB rates are written all or none.
* **Afterwards.** Nothing is recalculated: `Results` reads **OUT OF DATE** until the next run. **Restore parameters** does not change the source or the rates.

**The source on the run.** Each run records where its rates came from, in the run summary on `Results!C4` and in the evidence CSV (`fx_source`): `Closing price of` or `ECB reference rate of` and the date while the FX table is as **Download prices** wrote it; `Manual (edited after … download)` once a rate, a currency or a row has changed since; `Manual` when no download is recorded. A rebase on a new reporting currency keeps a downloaded table downloaded. The record of the download is part of the input fingerprint, so results read **OUT OF DATE** after a new download even when the rates are the same.

Whatever the source, the user remains responsible for the rates used: establish the correct as-of prices and reconcile them with the bank's data.


\Needspace{9\baselineskip}

## Dates and the day count

S, E, M, T and the time to a reset date are year fractions. The workbook computes each one as the **days counted** after the reporting date, up to and including the date, divided by the **days of the conventional year**:

\[
\text{years}=\frac{\text{days counted on the Calendar after AsOfDate, up to the date}}{\text{YearConvention}}.
\]

| Setting | Values | Meaning |
|:--|:--|:--|
| `DaycountConvention` | `ACTUAL` (locked) | The numerator: the actual days, as the calendar counts them. |
| `YearConvention` | 366, **365** (default), 364, 360, 252, 250 | The denominator: the days of the conventional year. |
| `Calendar` | **`TABLE`** (default), `TARGET`, `TARGET+TABLE` | `TABLE` counts every day except those of the table of closing days, so an empty table gives calendar days. `TARGET` counts the business days of the EUR TARGET calendar. `TARGET+TABLE` counts the TARGET business days, with the days of the table closed too. |

* **The default** is calendar days over 365, with an empty table: the same years as 1.2.0.
* **Business days.** Under `TARGET`, Saturdays, Sundays and the TARGET closing days are not counted. The TARGET closing days come from a rule, not from a yearly list: 1 January, Good Friday, Easter Monday, 1 May, 25 and 26 December. The rule applies from 1 January 1999; a reporting date before then stops a run under `TARGET` (`PRM-CALENDAR`). Business days are normally used only with a `YearConvention` of 252 or 250, as the notes of the cells say.
* **The table of closing days** (`I61:J160` of `Params`, headers *Bank Holiday* and *Description*) holds up to 100 dates with a description. Every row is read, also below a blank row. A weekend date or a date listed twice counts once. An entry that is not a date from 1 March 1900 stops a run (`PRM-HOLIDAY`).
* **What the setting does not change.** The floors and periods counted in business days, the MPOR, `MinMaturityBD` and `SDFloorBD`, still use `BusinessDaysPerYear` (250). A date on or before the reporting date keeps the calendar-day measure; only the checks of a matured trade or an expired option, and the floor of S at zero, read it. The order of a start and an end date is checked on the dates themselves. An option with no counted day left before its expiry is rejected as expired (`TR-EXPIRED`).
* **Recorded on every run.** The run summary on `Results!C4` adds the day count, for example `| year 365 days, calendar TABLE | FX Manual`, with the number of entries and a fingerprint of the table of closing days when the calendar uses it; the evidence CSV and the result line of `RunSACCR_Silent` carry the same as `year_convention`, `calendar` and `holiday_table`. The day count and the table are part of the input fingerprint. Results on two day counts are different results: do not compare them as if they were the same.
* **Errors.** A `YearConvention` or `Calendar` outside its list stops a run (`PRM-DAYCOUNT`, `PRM-CALENDAR`).

**Which convention to choose.** The CRR expresses T, S, E and M “using the relevant business day convention”, without naming a calendar or a divisor. The choice is the bank's and should match its own engine when you reconcile with it: for example business days of the TARGET calendar over 252 or 250. The default, actual calendar days over 365, is the workbook's historical default, not a convention the regulation prescribes. The worksheet function `SACCR_BusinessDays` counts business days the same way (see *Worksheet function API*).


\Needspace{9\baselineskip}

## Dates, named ranges and error handling

The workbook supports an actual Excel date, a serial number in Excel's **1900 date system**, or literal `YYYY-MM-DD` text. Text such as `30/09/2031` is deliberately **rejected** because its interpretation depends on locale. A workbook using the 1904 date system stops the run. An invalid reporting date stops the run; an invalid trade date rejects that trade. Be deliberate with future maturities, option expiries and products with reset schedules.

**Never rearrange `NettingSets` or `Trades` columns or rewrite their headings.** Header comparison tolerates case, spaces and line breaks, but not a changed schema. A missing required workbook name, an invalid reference or a moved header can stop the run. Visible formatting alone is not a controlled way to change the data contract.

\Needspace{12\baselineskip}

# Enter netting sets and collateral terms


\Needspace{9\baselineskip}

## NettingSets data contract

The `NettingSets` input table has **18 fields** in an exact sequence, in columns `C` to `T`. The header is on row 4, the last of the frozen rows, so it stays in view when the sheet scrolls; the data start on row 5. Every header carries a note: rest the pointer on it to read what the field is, the values allowed, what a blank means, when it is required, its CSV field name and its regulatory reference. Each row represents one netting set; the same netting-set ID must not be entered twice. A row can override the default regime, but all trades assigned to it follow that regime. The `Counterparty` field is descriptive; legal enforceability, eligible margin arrangements and transaction classification remain external responsibilities. The two margin columns (`VM`, `NICA`) and the three credit support settings (`TH`, `MTA`, MPOR conditions) have a **material numerical effect** and must not be approximated casually. A NettingSets CSV file carries the same 18 fields under short names, such as `netting_set_id` and `margined` (see *Import a portfolio from CSV files*).

| # | Exact header / field | Accepted concept, sign and key checks |
|--:|:--|:--|
| 1 | `NettingSetID` | Required unique key; referenced by every trade. Do not use `TOTAL`, `TOTAL CRR` or `TOTAL BCBS` as IDs, because these are output total labels. |
| 2 | `Counterparty` | Description / controlled identifier. Use synthetic identifiers in public test workbooks. |
| 3 | `Margined` (Y/N) | `Y` for a supported margin agreement; `N` for an unmargined netting set. Drives RC and maturity-factor treatment. |
| 4 | `Client clearing role` | Bank's role: `N` ordinary, `CM` clearing member facing its client, `CLIENT` client facing its clearing member, `CCP` the bank's trade exposure, as clearing member, to a CCP, or `MULTI` a multi-level client structure, which is recognized only to reject the netting set. Blank is treated as `N`. |
| 5 | `Remargin freq. (BD)` | Whole number of business days; `0` is interpreted as daily **with a warning**; accepted maximum 10,000. Used in MPOR `F + N - 1`. |
| 6 | `Large or illiquid` (Y/N) | `Y` for more than 5,000 trades or illiquid collateral triggers relevant 20-business-day MPOR floor; the business definition also covers hard-to-replace OTC transactions. Confirm actual eligibility and threshold. |
| 7 | `Margin disputes` (Y/N) | `Y` triggers the toolkit's doubled MPOR treatment; verify whether dispute conditions actually apply. |
| 8 | `MPOR override (BD)` | Optional strictly positive whole number of business days. It is applied only if not below the computed MPOR; a lower value is ignored with a warning and the computed MPOR is used. A non-whole or non-positive value makes the netting set `INVALID`. |
| 9 | `Net VM` | Net variation margin. Held is **positive**; posted is **negative**. VM treatment differs for unmargined one-way agreements by regime. |
| 10 | `NICA` | Net independent collateral amount; its regulatory definition includes eligible independent collateral held less relevant posted independent collateral, subject to eligibility, haircut and segregation rules. |
| 11 | `Threshold (TH)` | Margin threshold, expressed as an unscaled amount in the calculation currency. |
| 12 | `MTA` | Minimum transfer amount, expressed as an unscaled amount. |
| 13 | `Alpha override` | Optional per-netting-set alpha; values different from 1.4 classify the run `WHAT-IF`. |
| 14 | `Regime override` | Blank inherits the Params regime; otherwise use `CRR` or `BCBS`. The comparison and output retain the effective regime. |
| 15 | `Margin agreements` | Blank represents the supported single agreement/no agreement structure; `MIXED` or `SHARED` explicitly declares **unsupported structures** and causes rejection. |
| 16 | `QCCP` (Y/N) | Role `CCP` only, and then required: `Y` if the CCP qualifies under the netting set's regime (CRR: authorised or recognised under Regulation (EU) No 648/2012; BCBS: a QCCP as defined in CRE50), `N` otherwise. It sets the risk weight stated on `Checks`, not the EAD. Blank for every other role. |
| 17 | `CCP retains VM` (Y/N) | Role `CCP` only, and then required: `Y` if the CCP retains variation margin against the trades and the bank's collateral is not protected against the CCP's insolvency; it needs `Margined = Y`. Blank for every other role. |
| 18 | `Counterparty sector` | Optional FINREP counterparty sector: `CB` central banks, `GG` general governments, `CI` credit institutions, `OFC` other financial corporations, `NFC` non-financial corporations, `HH` households. Blank means not classified; a CCP, recognized from its clearing role, needs none. No SA-CCR rule reads it, so it changes no figure: it feeds the breakdown of EAD by sector on the Dashboard and in the run report. Any other value is an error (`NS-CODE` on the field *Counterparty sector*) and makes the netting set `INVALID`. |

**Rows and formats.** The template formats the first 200 netting-set rows; the drop-down lists and the highlight of a repeated ID cover rows 5 to 10,004, the limit of a NettingSets CSV file. A **Commit import** gives every row it writes the format of the first data row, so a longer portfolio looks the same to its last row.

![NettingSets with the demonstration portfolio: the header on row 4, in the frozen rows, and one row per netting set from row 5, with the navigation bar on the left.](assets/nettingsets.png){width=98%}

**Data ownership rule.** One input row is not necessarily a legal agreement record. The analyst must ensure that exposures, collateral and margin settings are correctly mapped to the intended netting set under the applicable rule. A workbook that calculates without warnings can still contain a poor legal or economic mapping.


\Needspace{9\baselineskip}

## Margined and unmargined replacement cost

For an **unmargined** netting set, the relevant replacement cost is essentially:

\[
\mathrm{RC}_{u}=\max(V-C,0),
\]

where `V` is current portfolio value and `C` is the collateral recognized under the appropriate unmargined convention. For the **margined** calculation, the standard structure is:

\[
\mathrm{RC}_{m}=\max(V-C,\ \mathrm{TH}+\mathrm{MTA}-\mathrm{NICA},\ 0).
\]

The toolkit's **cap calculation** compares the margined EAD with the EAD of the *same netting set without margin*; for the cap, `C = NICA` under both supported regimes. **Do not include margined VM in the unmargined cap's C.** This treatment follows CRE52.10, CRR Article 274(3) and EBA Q&A 2023_6962, and is important when interpreting a cap flag.

For an unmargined BCBS one-way agreement in which **only the bank posts VM**, posted VM is negative and enters collateral appropriately. Conversely, **VM held** is not supported as an unmargined input. Under CRR, a one-way agreement is treated as a margin agreement: enter the set as margined; posted VM on an `N` CRR netting set is rejected. The bank must determine which regulatory description accurately represents its arrangement.


\Needspace{9\baselineskip}

## MPOR calculation and clearing role

The Margin Period of Risk (MPOR) is measured in **business days**. The engine calculates the relevant floor from the margin agreement, frequency, large/illiquid flag, dispute flag, client-clearing role, and any permissible override. The basic approach combines a floor `F` with remargining frequency `N` as `F + N - 1`. Under the documented implementation, dispute treatment doubles that result, while large/illiquid treatment applies the 20-day floor with the documented frequency adjustment. An entered MPOR override cannot reduce the calculated requirement.

| Condition | Toolkit treatment |
|:--|:--|
| Ordinary margined bilateral | Base floor of 10 business days. |
| Clearing member (`CM`) facing client | 5-day floor under both CRR and BCBS, subject to eligibility. |
| Client (`CLIENT`) facing clearing member | 5-day floor under **CRR**; under **BCBS** the 10-day floor, a conservative interpretation of the Basel text. |
| Trade exposure to a CCP (`CCP`), margined | 10-day floor under both regimes. If the CCP retains VM and the collateral is not protected, each trade's horizon is the lower of one year and its remaining maturity, at least 10 business days, when that exceeds the netting set's MPOR. |
| Trade exposure to a CCP (`CCP`), unmargined | No MPOR floor: each trade's maturity factor uses its remaining maturity, as in any unmargined netting set. |
| Multi-level client structure (`MULTI`) | Not calculated: the netting set is `INVALID`. |
| Large/illiquid conditions | Relevant 20-business-day floor, with frequency treatment. One flag covers both more than 5,000 trades and illiquid collateral, so it also sets 20 days where the trade-count floor does not apply (for example a `CCP` netting set): a conservative limitation. |
| Remargin every N business days | MPOR floor plus `N-1`. |
| Margin dispute | Doubled calculated period under toolkit's deliberately conservative rule. |
| User MPOR override | Applied only when not below the calculated MPOR; a lower override is ignored with a warning. |

**Important:** `CM`, `CLIENT` and `CCP` are bank-role declarations, **not** evidence that a transaction qualifies for the corresponding treatment. The user must separately establish regulatory eligibility and documentation, including the QCCP status. A CCP entered with role `N` is calculated as a bilateral counterparty and cannot be detected.

Every `CM`, `CLIENT` and `CCP` netting set gets an INFO line on `Checks` naming its role, the MPOR floor and the rule applied, or, for an unmargined set, that no floor applies. For `CCP` it also states the risk-weight rule, which K-SACCR does not apply: 2% for a QCCP, or the CCP's standardised risk weight, to be determined outside the workbook, for a non-qualifying CCP; with `QCCP` blank or invalid, it names a CCP of unknown QCCP status and states no rule. Each regime with `CCP` netting sets gets an INFO line with their number, how many are `VALID`, and their EAD, which stays in that regime's total. Client trades for which the clearing member need not reimburse the client on the CCP's default have an exposure value of zero and are not entered.


\Needspace{9\baselineskip}

## Collateral, haircuts and amount scales

Collateral amounts on `NettingSets` are entered **after the user has applied** legal/eligibility and applicable valuation/haircut treatment. The workbook does not price collateral, establish eligibility, determine legally enforceable close-out netting, or automatically remove segregated bank-posted initial margin from NICA. The user must apply these definitions before entering inputs. Use one internally consistent reporting-currency unit scale for every netting-set amount. A missing or incorrect conversion is a data-control failure, not merely a reconciliation tolerance issue.


\Needspace{9\baselineskip}

## Status and exclusion rules

The engine assigns the following status, in this precedence order:

| Status | Trigger | What is shown |
|:--|:--|:--|
| `INVALID` | Netting-set-level error (including a missing or misplaced `QCCP` or retained-VM value, or role `MULTI`), duplicate netting-set ID, inconsistent trade reference/sub-class combination, a settled-to-market trade in a margined netting set, or a CRR IR netting set with unsupported `IRBucketOffset=FALSE`. | The netting set is represented on Results, but no EAD is calculated. |
| `INCOMPLETE` | The netting set is otherwise structurally valid, but at least one associated trade leg is rejected. | EAD is **withheld** rather than calculated on a partial portfolio. |
| `NO TRADES` | No trade rows refer to the netting set (a set whose trades were all rejected is `INCOMPLETE`). | No EAD; it is not counted as a withheld incomplete set. |
| `VALID` | Netting set and all of its trades satisfy the supported input contract. | Calculation figures and EAD can be used for further review. |

The status cell starts with the code; `INVALID` and `INCOMPLETE` are followed by a count, for example `INCOMPLETE: 1 of 3 trade(s) rejected`, or by the reason, for example `INVALID: multi-level client structure, not supported`. `INVALID` and `INCOMPLETE` rows are highlighted in light red. The order matters: a netting set with both an invalid netting field and an excluded trade is `INVALID`. An excluded trade is listed on `TradeCalc` and `Checks` with its reason (a row without a TradeID appears only on `Checks`, by its sheet row). In a portfolio total, only `VALID` sets contribute, and withheld sets must be identified separately. **A reported total of valid netting sets is not necessarily the exposure of the full input portfolio.**


\Needspace{12\baselineskip}

# Enter derivative trades and trade legs


\Needspace{9\baselineskip}

## Trades data contract

`Trades` contains **28 fields**, in columns `C` to `AD`, under a header on row 4 that stays in view, with the data from row 5. As on `NettingSets`, every header carries a note. The **TradeID is a unique leg identifier**. Every leg is linked to a netting set by `NettingSetID`. An economic trade may require one row or multiple rows depending on its material risk drivers or decomposition. The workbook reads fixed column positions verified against expected headings, not arbitrary user-added columns. A Trades CSV file carries the same 28 fields under short names, such as `trade_id` and `notional_amount` (see *Import a portfolio from CSV files*).

| # | Exact header / field | Meaning, format and common mistakes |
|--:|:--|:--|
| 1 | `TradeID` | Required unique **leg** ID. Duplicate IDs are errors and affect reconciliation population diagnostics. |
| 2 | `NettingSetID` | Must reference a valid netting-set ID from `NettingSets`. |
| 3 | `Asset class` | `IR`, `FX`, `CR`, `EQ`, `CO`, or `OT` (`OT` only under CRR). |
| 4 | `Sub-class` | Regulatory category, e.g. `AAA` (BCBS) or `CQS3` (CRR) for credit, `SINGLE`/`INDEX` for equity; must be coherent with asset class and regime. |
| 5 | `Risk factor / reference` | Currency, pair, reference entity, commodity risk driver or other hedging key. Exact identifiers influence aggregation. |
| 6 | `Instrument` | `Linear`, `Option` or `CDO`. Digital options are not a supported instrument type. |
| 7 | `Direction` | `Long` or `Short`, affecting supervisory delta and aggregation. |
| 8 | `Option type` | `Call` or `Put` for options; only meaningful where relevant. |
| 9 | `Nature` | `Standard`, `Basis` or `Volatility`; changes hedging-set treatment. |
| 10 | `Hedging-set label` | User-supplied distinguishing label for basis and volatility transactions. |
| 11 | `Notional` | Unscaled monetary or relevant underlying notional, in the specified notional currency; product-specific adjusted-notional rules apply. Zero or more: the sign of a position is given by `Direction`, and a negative notional rejects the leg. A leg with a notional of zero adds nothing to the add-on, while its MtM still counts in V. |
| 12 | `Notional CCY` | Currency of notional; for FX trades, must be a currency of the specified pair (`TR-FX-NOTIONAL`; in a CSV file, `CSV-CONTEXT-FX-NOTIONAL`). |
| 13 | `MtM` | Current mark-to-market for this leg/economic trade, in the specified MtM currency. Exactly one leg carries it in a multi-leg economic trade. |
| 14 | `MtM CCY` | MtM currency; blank defaults to the notional currency. Must be convertible via Params. |
| 15 | `Start date (S)` | Relevant effective/start date; mandatory for specified underlying IR/credit options. |
| 16 | `End date (E)` | Date used in duration and IR bucket logic; often underlying end date. |
| 17 | `Maturity date (M)` | Contractual maturity date. For a settled-to-market trade, enter the reset date in field 27, not here. |
| 18 | `Option expiry (T)` | Required option expiry; must be after the reporting date and not after M. |
| 19 | `Underlying price (P)` | Option underlying price/rate (including FX quote conventions). |
| 20 | `Strike (K)` | Option strike in the same market convention as price. |
| 21 | `Lambda shift` | Optional shift added to P and K when they are negative or near zero (BCBS, CRE52.40 FAQ). For CRR interest-rate and commodity options the engine computes the RTS lambda from `LambdaThresholdIR`/`LambdaThresholdCO` and ignores this value, with a warning if it differs. Do not use it as an arbitrary delta adjustment. |
| 22 | `Attachment (A)` | CDO tranche attachment point as a decimal fraction. |
| 23 | `Detachment (D)` | CDO tranche detachment point as a decimal fraction; greater than A. |
| 24 | `Economic trade ID` | Optional group key identifying legs belonging to the same economic instrument. Blank means this row is a whole trade. |
| 25 | `MtM leg` (Y/N) | Exactly one leg of a grouped economic trade carries the full MtM (`Y`); other legs hold zero MtM. |
| 26 | `Settlement type` (N/STM) | Blank or `N`: ordinary treatment. `STM`: settled-to-market, see below. |
| 27 | `Next reset date` | `STM` only, and then required: the first contractual settlement date after the reporting date, not after M. Blank for an ordinary trade. |
| 28 | `Comment` | Free-text operator note; not read by the engine (its header is not checked). Do not rely on this field for mandatory machine-readable inputs. |

**Rows and formats.** The template formats the first 5,000 trade rows; the drop-down lists and the highlight of a repeated ID cover rows 5 to 100,004, the limit of a Trades CSV file. A **Commit import** formats every row it writes like the first data row.

**Safety distinction:** a trade being accepted means its **declared** instrument, dates, rates and structure comply with the engine's supported field contract; it does not prove that its declared underlying risk drivers, decomposition or model inputs are economically correct.


\Needspace{9\baselineskip}

## Assets and supported instruments

| Asset class | What is supported | What requires care |
|:--|:--|:--|
| `IR` | Linear interest-rate derivatives, IR options, inflation exposures declared as IR | Currency/`-INFL` identifiers, supervisory duration, maturity buckets and underlyings of swaptions/caps/floors. |
| `FX` | FX linear trades and options | Currency pair, correct notional leg, quote order and delta sign. |
| `CR` | Credit single names and indices, options and credit CDO/tranche rules | Rating versus CRR CQS, reference consistency, sub-class factors and tranche attachment/detachment. |
| `EQ` | Equity single-name and index instruments and options | Index versus single-name classifications, reference-based aggregation. |
| `CO` | Commodity instruments and options, CRR climatic hedging set | Commodity category, electricity factor, negative price/strike treatment and special CRR classification. |
| `OT` | Other-risk drivers **under CRR only** | Each identical primary risk driver maps to its own hedging set; rejected under Basel. |

A **linear trade** uses supervisory delta +1 when long and -1 when short, with the relevant product-level signed interpretation. **Options** use the supported supervisory-option-delta treatment. European, American, Bermudan and Asian styles are covered under the documented input conventions: for Bermudan use the **latest exercise** date as expiry; for Asian options the input price represents the current value of the relevant average. **Credit CDO tranches and nth-to-default** use attachment/detachment treatment. Basis or volatility nature changes hedging-set coefficients rather than creating a new asset class.

**Not included:** digital options under either regime, automatic decomposition/pricing of arbitrary exotic structures, optional zero-EAD treatment for qualifying sold options, the optional unpaid-premium cap for certain sold credit protection, unsupported multiple/shared margin-agreement structures, and other methods outside full SA-CCR.


\Needspace{9\baselineskip}

## Multi-leg economic trades: prevent MtM double-counting

A derivative can have several material risk drivers, or an option combination can require decomposition. The supported representation is **one row per eligible leg**, all sharing the same `Economic trade ID`. Exactly one leg is marked `MtM leg = Y` and carries the **entire economic trade's MtM**; the other legs are marked `N` and have MtM **zero**. All legs must belong to the **same netting set**. If this fails, the engine rejects all legs of that economic trade. The user remains responsible for mapping the risk drivers and choosing the correct economic decomposition.


\Needspace{6\baselineskip}

### Illustrative three-leg example

| TradeID | Economic trade ID | NettingSetID | MtM leg | MtM |
|:--|:--|:--|:--:|--:|
| `TRD-101-IR` | `TRD-101` | `NS-01` | `Y` | +250,000 |
| `TRD-101-FX` | `TRD-101` | `NS-01` | `N` | 0 |
| `TRD-101-EQ` | `TRD-101` | `NS-01` | `N` | 0 |

This example represents a **possible input structure only**: no claim is made that the instrument economically has these three risk categories. If the user repeated +250,000 in each row, a naive approach would triple the portfolio MtM. K-SACCR's grouping check is designed to reject such errors rather than silently aggregate an inconsistent economic trade.

**Option combinations** such as spreads, straddles, strangles, collars and caps/floors can be represented using suitable option component legs. Within the same asset class, grouped option combinations must be option legs; a normal linear swap should not be split as if it were an option combination. A digital payoff is **not made supported** merely by approximating it with a collection of standard options.


\Needspace{9\baselineskip}

## Date conventions and maturity inputs

The workbook derives year quantities on its day count, `YearConvention` and `Calendar` (see *Dates and the day count*). **S** is years from reporting date to the start date, floored to zero if blank or already begun except where a start date is mandatory. **E** is years to the end date (or the derived fallback); **M** is years to maturity (or eligible fallback / reset); **T** is years to option expiry. The **regulatory business-day** floors are calculated using `BusinessDaysPerYear` and MPOR settings, whatever the calendar.

* For an interest-rate or credit trade, **E must be later than S** where duration is required; the two dates are compared, not their year fractions.
* A maturity at or before the reporting date is invalid; options that are expired, or have expiry after M, are rejected.
* For an **option on an interest-rate or credit underlying**, S and E describe the underlying rather than only the option's own life. A swaption's S is ordinarily the exercise/start of the underlying swap and E its final end. A Bermudan swaption can need S as the earliest permitted exercise and T as the latest. Missing mandatory S is rejected.
* For an amortizing or accreting IR/credit trade, enter the correct **time-weighted average remaining notional**. For leveraged or notional-exchange products apply the required leverage/payment conventions before input.
* For a settled-to-market trade (`STM`), M becomes the time to the **next reset date**, floored at 10 business days in the maturity factor; S, E and the option expiry T keep their meaning, so a settled option can expire after its reset date. There is no automatic notional or reset schedule reader.

\Needspace{6\baselineskip}

### Settled-to-market trades

Declare `STM` only when the contract settles its outstanding exposure and resets its value to zero on set dates. **Daily variation margin alone is not settlement to market**: a margined contract stays an ordinary trade in a margined netting set. The qualification is the user's.

* The next reset date must be after the reporting date (the MtM is the value after any settlement due that day) and not after the maturity as entered; a reset on the maturity date is accepted.
* `STM` without a reset date, blank or `N` with one, another code, or a reset that is not a date rejects the trade.
* The legs of one economic trade must share the settlement type (blank and `N` alike) and the reset date.
* A settled trade needs the unmargined treatment: an `STM` trade in a margined netting set makes the set `INVALID` (*STM requires unmargined treatment; mixed margined/unmargined netting sets are not supported*), a limitation of the workbook rather than a regulatory prohibition.


\Needspace{6\baselineskip}

### Interest-rate maturity buckets

| End date E (years) | CRR bucket | BCBS bucket |
|:--|:--|:--|
| `0 < E < 1` | Bucket 1 | Bucket 1 |
| `E = 1` | **Bucket 1** | **Bucket 2** |
| `1 < E <= 5` | Bucket 2 | Bucket 2 |
| `E > 5` | Bucket 3 | Bucket 3 |

This **exact one-year boundary difference** is deliberate; it can change aggregation and EAD even with identical economic trade data. The bucket boundary at exactly five years is bucket 2 under both regimes.


\Needspace{9\baselineskip}

## FX forwards: pair orientation and notional

FX risk factor is two distinct three-letter currency codes, e.g. `EUR/USD`, `EURUSD` or the reverse. Both currencies need a rate in the FX table. A forward may be entered in either orientation; the hedging key is normalized in alphabetical order, and the supervisory-delta sign changes appropriately when input orientation is reversed.

The entered notional represents **one leg**, not the sum of both. It must be in one of the pair's currencies. Use the **foreign-currency leg** against reporting currency or, if both legs are foreign to reporting currency, the leg with the **larger reporting-currency value** under the relevant rule. The workbook checks that the notional currency is a pair currency, in either order of the pair, both in a run and in the CSV checks; a blank notional currency is read as the reporting currency. It **does not independently decide whether the selected leg was the correct one**.


\Needspace{9\baselineskip}

## FX options: mandatory market-quote order

Unlike linear FX, an FX option must be entered in the market-quote order used for the regulatory delta. The toolkit's priority ranking is:

**EUR, GBP, AUD, NZD, USD, CAD, CHF, JPY**.

Examples: **EUR/USD, GBP/USD, USD/JPY, EUR/CHF**. A ranked currency comes before an unranked one, so `USD/SEK` is accepted. If neither is in the ranking, the alphabetically earlier currency is the base, e.g. `NOK/SEK`. Options entered in reverse order are **rejected**, because a numerically analogous call/put on inverse quotes does not generally yield the same supervisory delta. Price P and strike K must be quoted in the entered orientation.


\Needspace{9\baselineskip}

## Credit, commodity, inflation and other-risk coding

For **credit single names**, classify by `AAA` through `CCC` under BCBS or `CQS1` through `CQS6` under CRR, by the regime of the netting set; a trade on the other regime's scale is rejected (`TR-FACTOR-REGIME`). Index categories `IG_INDEX` and `SG_INDEX` are supported under both regimes. For CRR unrated exposures, the appropriate CQS is an **input determination**, not automatically a separate “unrated” factor. Use the prescribed relevant CQS rule, including CQS5 where Article 128 applies. An **unlisted multi-name credit position** (a basket that is not an index listed on a recognized exchange) is not supported and cannot be detected: no factor row computes the notional-weighted factor of its constituents, and the `CQS1`–`CQS6` rows carry the single-name correlation of 50% where a multi-name reference entity takes 80%. Do not relabel such a basket as `IG_INDEX` or `SG_INDEX` or round its factor to a credit quality step: the add-on would be wrong in a run still classified `REGULATORY`. The only workaround is a factor-table row with the weighted factor and a correlation of 80%, added after a maintenance unlock; the run is then `WHAT-IF`, and **Restore parameters** refuses while that row is below the table.

For **inflation**, use class `IR` and a separate factor identifier such as `EUR-INFL` to keep inflation from ordinary interest-rate risk within the same currency. For **commodity climatic** exposures, the `CLIMATIC` sub-class and associated hedging treatment are CRR-only; a Basel climatic trade is not converted into an invented class. For **other risks**, `OT` is CRR-only and the same primary risk driver forms a hedging set. For **CDO/nth-to-default**, ensure attachment and detachment are proper fractions and that `A < D`; nth-to-default may be represented by `A = (n-1)/k`, `D = n/k` when that mapping is applicable.


\Needspace{9\baselineskip}

## Examples of inputs that must be corrected, not ignored

| Problem | Typical observation | Resolution principle |
|:--|:--|:--|
| Two rows with the same `TradeID` | Duplicate-ID error; trade-level bank matching may become non-comparable. | Allocate unique leg IDs; do not solve it by hiding a row. |
| A negative `Notional` | Excluded trade: *notional must not be negative (use Direction for the sign)*; in a Trades CSV file, a `CSV-RANGE` finding on `notional_amount`. | Enter the notional as a positive amount and the side in `Direction`. A notional of zero is accepted. |
| An unrecognized `NettingSetID` | Trade excluded and netting population exception. | Add/correct the corresponding netting set and its metadata. |
| Option expiry `T` earlier than or on AsOfDate | Excluded trade. | Verify the valuation date / expiry or remove already-matured trade appropriately. |
| Currency pair in reverse order for FX **option** | Excluded option despite economically equivalent wording. | Normalize to the mandated market quote orientation and adjust price/strike. |
| Invalid locale-formatted date string | Excluded row or stopped run. | Use a true Excel date or ISO `YYYY-MM-DD`. |
| Unsupported margin structure `MIXED` / `SHARED` | `INVALID` netting set; EAD withheld. | Split/restructure under a supported legal arrangement; do not misdeclare its status. |
| Inconsistent grouped MtM legs | All legs of economic trade rejected. | One shared economic ID, exactly one `Y` leg with full MtM, others zero. |
| CRR IR set and `IRBucketOffset=FALSE` | `INVALID`; CRR does not offer this option. | Set TRUE for CRR or use BCBS only when that is the intended comparison. |
| `OT` under Basel | Excluded trade. | Do not invent Basel support for sixth asset class. |
| Role `CCP` with `QCCP` blank, or `QCCP` filled for another role | `INVALID` netting set. | Enter `Y` or `N` for a CCP, leave it blank otherwise. |
| `STM` trade in a margined netting set | `INVALID` netting set. | Enter settled trades in an unmargined netting set; do not declare daily VM as `STM`. |
| A rating such as `BBB` in a CRR netting set, or a step such as `CQS3` in a BCBS one | Excluded trade (`TR-FACTOR-REGIME`); in a CSV file, `CSV-CONTEXT-REGIME`. | Use the scale of the netting set's regime: `CQS3` under CRR has the factor of `BBB` under BCBS. |
| An FX trade whose notional currency is not a currency of its pair | Excluded trade (`TR-FX-NOTIONAL`); in a CSV file, `CSV-CONTEXT-FX-NOTIONAL`. | Enter the notional of one of the pair's legs, in that leg's currency. |
| A counterparty sector outside the list | `INVALID` netting set (`NS-CODE`). | Enter `CB`, `GG`, `CI`, `OFC`, `NFC` or `HH`, or leave it blank. |

\Needspace{12\baselineskip}

# Import a portfolio from CSV files


\Needspace{9\baselineskip}

## What an import does

Instead of typing or pasting rows on `NettingSets` and `Trades`, you can load them from two CSV files: a **NettingSets file**, one record per netting set, and a **Trades file**, one record per trade leg. Each CSV column fills one column of its sheet, with the same meaning, allowed values and blank meaning as the sheet column. An import never writes to an input sheet directly. It passes through two staging sheets, and each step but the correction has its own command. The commands are in the CSV import group of the K-SACCR tab, where **Import CSV** is a menu with **NettingSets file...** and **Trades file...**, and **Templates** writes the blank files and the import specification. The same menus load and write the two files of the bank's figures, which have no staging sheet (see *Load the bank figures from a CSV file*). The four buttons of each staging sheet run the same commands; the steps below name them as the buttons do.

```
 NettingSets file              Trades file
        |                           |
   Import CSV                  Import CSV
        v                           v
 Import_NettingSets          Import_Trades     values kept as text;
        |                           |          correct them here
        +-------------+-------------+
                      |
              Validate staging    -> findings on Checks, origin IMPORT
                      |
               Commit import      -> replaces NettingSets and/or Trades
                      |
                Run SA-CCR        -> Results and the other outputs
```

1. **Import CSV** reads a file and writes its records onto the staging sheet as text, exactly as in the file.
2. If needed, you **correct** staged values on the staging sheet, or correct the file and import it again.
3. **Validate staging** checks the staged rows of both staging sheets, against each other, against the input sheets they will join and against `Params`.
4. **Commit import** replaces each input sheet whose staging sheet holds rows with those rows, once a validation without findings still holds.
5. **Run SA-CCR** calculates. **Nothing is calculated automatically**: after a commit, `Results!C5` reads **OUT OF DATE** until you run.

The Status group of the tab shows the status of both staging sheets, and **Validate staging** and **Commit import** are enabled only when there is something they can act on (see *Enabled states*).

**Run SA-CCR** and **Validate inputs** read `NettingSets` and `Trades` only, never the staging sheets. Until a commit, an import changes neither the input sheets nor `Results` nor the run fingerprints. `Params`, the bank sheets and their settings are not imported; enter them on their sheets.


\Needspace{9\baselineskip}

## The import specification page

**Import specification...**, in the **Templates** menu of the K-SACCR tab, and the button **Export import specifications**, on `Import_NettingSets` and on `Import_Trades`, save a description of the four CSV files as one HTML page, so that whoever prepares the files has every rule at hand, also without the workbook.

* A save dialog offers the file name `K-SACCR_Import_Specification_v3.html` in the workbook's folder. If the chosen file exists, the workbook asks whether to replace it, with **No** as the default; **No** writes nothing.
* The page is written whole before it replaces an older file; a failed write, or a read-only file, leaves the older file as it was.
* The message *Import specifications exported* then asks *Open the page now?*; **Yes** opens it with the program Windows uses for HTML files.
* The workbook is not changed, so the export works at any time, with the protection on and whatever the inputs hold.

The page is self-contained: it has no script and loads no other file, so it can be read offline, archived with the files it describes and printed (it is laid out for A4 landscape). It describes **contract version 3**: NettingSets files of version 2, whose blank template ends in `_v2`, and Trades, BankNettingSets and BankTrades files of version 1, whose templates end in `_v1`. A file's version changes only when a column is added, removed or renamed, not with each K-SACCR release.

| Part of the page | What it gives |
|:----|:--------|
| Files | The four files, their version, sheet, key column, record limit and blank template name, and how each is loaded. |
| File format, header, value types | The rules summarized in *CSV file format* below, in full. |
| Blank fields, empty files and limits | What a blank field means, header-only files, the record and size limits. |
| NettingSets file, Trades file, BankNettingSets file, BankTrades file | For each column: its position, exact CSV name, the sheet column it fills, meaning, type and limits, whether it is required, its allowed values, what a blank means and an example value. Then the rules between the fields of a record and, for Trades, the rules between records and against `Params`. |
| Example of each file | Synthetic records as CSV text to copy: three netting sets, five trade legs, and three records of each bank file. |
| Using the examples | How to save and import the examples. |
| Findings | Every finding code with its meaning. |

**The examples.** The two examples belong together: every leg's netting set is in the NettingSets example, and the two files import and validate together, without a finding, into a workbook whose `Params` are as delivered. They are synthetic and hold no real portfolio data. Between them they show an identifier with leading zeros (`00731`), a quoted text holding a semicolon and quotes, letters outside ASCII, blank optional fields, counterparty sectors, an economic trade with two legs, an option, a trade settled to market and a CCP netting set. The trades suppose a reporting date before 2026-12-31, FX rates for `EUR`, `GBP` and `USD`, and supervisory-factor rows for `EQ_SINGLE`, `FX` and `IR`; with other `Params`, a validation may report a matured trade or a missing rate.

To try them, in a copy of your reference copy, since the commit replaces the demonstration rows:

1. Open the page. Select the whole text of the NettingSets example, copy it into a plain-text editor and save it as **UTF-8** with the extension `.csv`. Do the same with the Trades example. A file saved in another encoding is rejected (`CSV-ENCODING`), because the examples hold letters outside ASCII.
2. On `Import_NettingSets`, press **Import CSV** and choose the NettingSets file; on `Import_Trades`, the same with the Trades file. On the K-SACCR tab, **Import CSV > NettingSets file...** and **Import CSV > Trades file...** do the same from any sheet.
3. Press **Validate staging**. Both staging sheets read `VALID`.
4. Press **Commit import** and confirm, then **Run SA-CCR**.

The examples are not a template for a real portfolio: start a real file from the blank template and the column tables.


\Needspace{9\baselineskip}

## CSV file format

| Item | Rule |
|:----|:--------|
| Encoding | UTF-8, with or without a byte-order mark. A file that is not valid UTF-8, or contains a NUL character, is rejected. |
| Delimiter | A semicolon, `;`, between fields, whatever Excel's regional settings. |
| Quoting | A field holding `;`, a double quote or a line break is enclosed in double quotes, and a quote inside it is written twice: `"A ""B"""` reads `A "B"`. A quote inside an unquoted field, text after a closing quote or a quoted field left open is a syntax error. |
| Records | Separated by CRLF or LF; the last may end without one. A quoted text field may span several lines. |
| Header | The first record names every column of the file exactly once, in any order, with the exact lower-case names of the specification. A missing, unknown or repeated column rejects the file. There are no optional columns: an optional field is a column whose value may be blank. |
| Numbers | Decimal point `.`, no thousands separator, no `+` sign, no leading zero (`0.5`, not `007.5`), an optional exponent (`1.25e6`), at most 15 significant digits. `1,5` and `1.000,00` are rejected. |
| Dates | `yyyy-mm-dd`, from 1950-01-01 to 2200-12-31. `31/03/2027` and serial numbers are rejected. |
| Codes and flags | Exact spelling and case: `Linear`, not `linear`; `Y` or `N`, not `y`, `Yes`, `TRUE` or `1`. Currencies are three capital letters. |
| Identifiers | At most 64 characters, no vertical bar or `#`; read as text, so leading zeros are kept. Keys are compared ignoring case, as the calculation compares IDs. |
| Text | At most 500 characters; line breaks allowed. |
| Spaces | Part of the field: a leading or trailing space in a value other than text is rejected, never trimmed. |
| Formulas | Text that starts with `=`, `+`, `-`, `@`, a tab or a carriage return is rejected, so that no field becomes a formula on a sheet. A negative number is a value. |

**Blank fields and empty files.** A blank field is an empty one. A required field must not be blank; an optional blank field takes the meaning the specification gives under *Blank means*, which is the sheet's. A record whose fields are all blank, and an empty line, are findings. A **header-only file** is valid and holds no record: importing it empties the staging sheet, and a commit never clears an input sheet for lack of staged rows.

**Limits.** A NettingSets file holds at most **10,000** records and a Trades file at most **100,000**; a BankNettingSets file at most **10,000** and a BankTrades file at most **50,000**. A file with more is rejected whole. A file larger than **128 MiB** (134,217,728 bytes, a byte-order mark included) is rejected from its size alone, before it is read. The drop-down lists of the input sheets cover these limits, `NettingSets` rows 5 to 10,004 and `Trades` rows 5 to 100,004, and a commit formats every row it writes. Loading, validating and committing a file at these limits has not been timed.

**Header and version.** The first record names every column of the file exactly once, in any order. A NettingSets file of **version 1**, whose header is exactly the 17 columns of 1.2.0, without `counterparty_sector`, is still read, with the sector blank in every record; any other header is judged against version 2, so a version 1 header with a column missing or added reports `counterparty_sector` as missing too. Trades files are of version 1, as in 1.2.0.

**Files saved by Excel.** Excel's own CSV save writes numbers and dates in the format of the computer's locale, such as a decimal comma or a day-first date, and in some locales separates fields with commas. Such a file is rejected: a comma-separated header gets one `CSV-DELIMITER` finding, and decimal commas and local dates get `CSV-TYPE`. Write the files with a program that follows these rules.


\Needspace{9\baselineskip}

## Prepare the files

* **Blank templates.** **NettingSets template...** and **Trades template...**, in the **Templates** menu of the K-SACCR tab, write the blank file of their sheet, the header line alone, to a file you choose. The names they offer, in the workbook's folder, are `K-SACCR_NettingSets_v2.csv` and `K-SACCR_Trades_v1.csv`. Add one record per netting set or trade leg below the header.
* **Columns.** The column tables of the import specification give, for each CSV column, the sheet column it fills and its rules. The CSV names are `snake_case` names, fixed by the contract; the note of each header on `NettingSets` and `Trades` names the CSV field of its column, and row 8 of each staging sheet lists them.
* **Legs and economic trades.** A record of the Trades file is one **leg**. An economic trade with several legs follows the rule of *Multi-leg economic trades: prevent MtM double-counting*: the same `economic_trade_id` and `netting_set_id` on every leg, exactly one leg with `mtm_leg` `Y` carrying the whole MtM, the others `N` with an MtM of `0`.
* **Both files or one.** You can import both files, or only one. Legs staged alone are checked against the netting sets on `NettingSets`; netting sets staged alone are checked with the legs on `Trades`.
* **Check a file without loading it.** The macros `CheckNettingSetsCsv` and `CheckTradesCsv`, which are not on the K-SACCR tab (**Alt+F8**, select, **Run**), read a file you choose and say whether it follows the contract, with its record count or its first findings. They write nothing to the workbook.

Keep the files you commit: they document where the inputs came from.


\Needspace{9\baselineskip}

## Import CSV onto a staging sheet

Each staging sheet shows, at the top, the **source file** (`D4`), the **load time** (`D5`), the **record count** (`D6`) and the **status** (`D7`). Below, row 8 holds the CSV column names, explained in a note on `D8` (the workbook does not read that row), and row 9 the sheet headers; the staged rows start at row 10. Column C, `CSV line`, holds the physical line of the file on which each record starts, and the values follow from column D, in the order of the input sheet's columns: on `Import_NettingSets` from `D` to `U`, the sector in `U`; on `Import_Trades` from `D` to `AE`. The value cells are formatted as text, so leading zeros, long identifiers, ISO dates and other values stay exactly as written, and nothing on a staging sheet is calculated.

Press **Import CSV** on `Import_NettingSets` for a NettingSets file, or on `Import_Trades` for a Trades file, or choose **Import CSV > NettingSets file...** or **Import CSV > Trades file...** on the K-SACCR tab, and choose the file.

* **A structural finding loads nothing:** a file that is not UTF-8, is not valid CSV, has no header, uses another delimiter, has a repeated, missing or unknown column, a record with the wrong number of fields, too many records or too many bytes. The message lists the findings and says *Nothing was loaded*: the staging sheet is unchanged.
* **Otherwise the records are loaded**, one row each, after the rows the sheet held before were cleared. A record whose fields are all blank is a finding and is not loaded as a row. When the file follows the contract, the message says *Loaded into … as text* and *Not validated yet*. When it has findings about values, such as a wrong date format or a missing required field, the records are loaded all the same, so that you can correct them on the sheet, and the message lists the first findings, each with its line, record, column, code and text. Every finding is judged again, and listed on `Checks`, by **Validate staging**.
* **Changes are protected.** When the staging sheet has changes since its last load, or rows that were not loaded from a file, the workbook asks before replacing them, with **No** as the default.
* A value longer than an Excel cell holds (32,767 characters) loads nothing. A load that fails while writing puts the staging sheet back as it was.

The sheet then reads `LOADED`. `NettingSets`, `Trades`, `Results` and the run fingerprints are not changed.

![Import_NettingSets after Import CSV of the example file of the import specification: the buttons, the source file, the load time, three records, the status LOADED, the row of CSV field names and the three staged netting sets as text, each with its CSV line.](assets/import_nettingsets.png){width=98%}


\Needspace{9\baselineskip}

## Staging statuses

| Status | Meaning |
|:----|:--------|
| `EMPTY` | Nothing has been loaded onto the sheet. |
| `LOADED` | A file was loaded and nothing on the sheet has changed since. A header-only file reads `LOADED` with 0 records. |
| `EDITED` | A staged cell was changed after the last load, validation or commit. |
| `VALID` | The last **Validate staging** found no finding on this sheet. |
| `INVALID` | The last **Validate staging** found at least one finding on this sheet. |
| `OUT OF DATE` | The sheet read `VALID`, but something that validation judged has changed since: a staged row, `NettingSets`, `Trades` or `Params`. Shown when the sheet is activated; it reads `VALID` again if the change is undone. |
| `COMMITTED` | The sheet's rows were written to its input sheet by **Commit import**. The rows stay as a record of what was loaded. |

The status is in `D7` of each staging sheet, and the Status group of the K-SACCR tab repeats it as *NettingSets staging:* and *Trades staging:*. It is for you to read; the tab uses it only to enable **Validate staging** and **Commit import**. **Commit import** never relies on it: it checks itself whether a validation without findings still matches the staged rows, the input sheets and `Params`, so a `VALID` saved with the workbook is never enough to commit.


\Needspace{9\baselineskip}

## Correct staged values

* **Edit a value** by typing over the staged cell. The cell is formatted as text, so type the value as the file would hold it: a date as `yyyy-mm-dd`, a number with a decimal point, a code with its exact spelling. The sheet reads `EDITED`.
* **Add a record** by typing it in the first empty row below the staged rows, or in an inserted row. Its `CSV line` stays blank; its findings name line 0.
* **Do not enter formulas.** A formula in a staged cell refuses the validation: *Nothing was validated: … holds a formula. Enter the value itself, as text.*
* **Remove records in the file.** Rows cannot be deleted on the protected sheet, and a row whose values you clear stays as a blank record, reported by `CSV-BLANK-RECORD`, because its `CSV line` remains. To drop records, correct the file and import it again.
* **Empty a staging sheet** by importing a header-only file, such as the blank template. This matters because a commit replaces every input sheet whose staging sheet holds rows, rows left from an earlier import included.

A file can always be corrected outside Excel and imported again; the import asks before it replaces changes made on the sheet.


\Needspace{9\baselineskip}

## Validate staging

**Validate staging**, on the K-SACCR tab or on either staging sheet, checks the rows of both. It first recalculates the workbook, as a run does, so that a formula on `Params`, `NettingSets` or `Trades` is judged by its current result also in manual calculation. It then judges:

1. **Each staged row**, as a record of a file: each field (required, type, allowed values, range, length, formula-like text), the rules between the fields of the row, duplicate and reserved keys, blank rows and the record limit.
2. **The rules between records**, for the staged legs: each leg's netting set exists, among the staged netting sets when `Import_NettingSets` holds rows, or else on `NettingSets`; and the rules of economic trades (one MtM leg, MtM `0` on the others, one netting set, one settlement type and reset date, and only options when two legs share an asset class). With only netting sets staged, the legs on `Trades` are checked against them, which finds the legs a commit of those netting sets would leave without a netting set.
3. **The rules against `Params`**, read as a run reads them: no trade matured, no option expired and no reset date of an `STM` trade on or before the reporting date; an FX rate for every notional, MtM and FX-pair currency; the notional currency of an FX trade, or the reporting currency when it is blank, a currency of its pair; a supervisory-factor row for every trade; and, under each netting set's regime (its staged or live override, or else the `Params` regime), `OT` and regime-limited factor rows only under their regime. With only netting sets staged, the factor and regime rules judge the legs on `Trades`, because a staged regime override can change their outcome. When a run would stop at `Params`, one `CSV-CONTEXT-PARAMS` finding gives the reason, on the cell concerned, and no staging sheet can be `VALID`.

**Findings on Checks.** The findings replace the lines of `Checks`, with **`IMPORT`** in the Origin column. The first line is an `INFO` summary with the time and the rows and errors of each sheet. Each finding follows as an `ERROR` line with its sheet and row, a link in the Reference column that leads to its cell, its code, column, text and CSV line in the message, and its code and CSV field name in the Code and Field columns. A finding that names no column is on the row's `CSV line` cell. At most 10,000 findings are listed; the summary gives the full count. Every finding is an error: there is no warning to accept.

**Marks.** The staged cells with a finding get a red fill, after the red fills of the previous validation are removed; other fills are left alone. A cell of `NettingSets`, `Trades` or `Params` is never marked or changed: a finding on such a cell is listed on `Checks` only.

**Status.** Each staging sheet that holds rows reads `VALID` or `INVALID`. Correct the findings and validate again until no finding remains; a sheet can read `VALID` while the other reads `INVALID`, and the commit then refuses. The message ends by saying that nothing was copied to `NettingSets` or `Trades`.

**Refusals.** Nothing is validated, and nothing changes, when a staged cell holds a formula, or when both staging sheets are empty (*Nothing was validated: Import_NettingSets and Import_Trades hold no rows. Import a CSV file first.*).

**Effect on the last run.** The input sheets, `Results` and the run fingerprints are not changed, but `Checks` no longer shows the checks of the last run. Until the next **Run SA-CCR**, whose lines carry the origin `RUN`, `ResultsStatus()` reads `EDITED` and **Export evidence** and **Export report** refuse: *Checks shows a staging validation, not the checks of the last run; press Run SA-CCR first*. The tab does not detect this: its Results label and the enabled states of the two exports follow `Results!C5`, which checks the inputs only, so the refusal comes from the export itself.

**What a validation does not judge.** The validation applies the rules a file can be judged by on its own and against `Params`. The calculation's own validation still applies at the next run to the rest: the sub-class an asset class and regime require; an option's type, price and strike, its expiry relative to the maturity; the form and quote order of an FX pair; the hedging-set label of a basis or volatility trade; the reset date relative to the maturity; an `STM` trade in a margined netting set; and the netting-set rules that depend on `Params`. A committed portfolio can therefore still have `INVALID` netting sets and excluded trades, exactly as one typed by hand.


\Needspace{9\baselineskip}

## Commit import

**Commit import**, on the K-SACCR tab or on either staging sheet, writes the staged rows to the input sheets.

* **It needs a current validation.** The commit recalculates the workbook and refuses, writing nothing, unless the last **Validate staging** found no finding and nothing it judged has changed since: the staged rows, `NettingSets`, `Trades` or `Params`, even when the change was made with Excel's events off. The message is *Nothing was committed. The staged rows have no VALID validation that still matches them, the input sheets and Params. Run Validate staging, correct any finding, then commit.* Rows with findings are therefore never committed.
* **It asks first.** A confirmation gives, for each input sheet, its rows now and the staged rows that will replace them, or that it is left as it is because its staging sheet holds no rows. It says that **Undo cannot take the commit back** and that the staged rows stay on the staging sheets. **Cancel** is the default.
* **It replaces whole sheets.** `NettingSets` is replaced by the rows of `Import_NettingSets` when that sheet holds any, and `Trades` by those of `Import_Trades` when it holds any. The staged rows are written from the first data row, and the input cells of every row below them that held a value or a formula are cleared. There is **no merge and no update by key**: a netting set or leg that is not staged is no longer on the input sheet after the commit. An input sheet whose staging sheet holds no rows is **left as it is**, so nothing is ever erased for lack of staged rows. Every row written gets the format of the first data row: number format, alignment, fill, font and borders, so that a portfolio longer than the rows the template formats looks the same to its last row. Drop-down lists, workbook names and the other sheets are not changed; a cell formatted *General* that receives a date takes a date format, as when a date is typed into it.
* **It writes values as typed by hand.** A number is written as a number, a date as an Excel date and every other value as text, so that leading zeros, as in `00731`, and text that looks like a number or a date stay as staged. A blank field stays blank. A run on committed rows gives the same results as a run on the same rows typed by hand.
* **It checks what it wrote, and puts the sheets back on failure.** Before writing, the commit keeps every input cell it may change; after writing, it reads each sheet back. On any failure it writes the kept cells back, reads them back and says so: *The input sheets were put back as they were and checked*, or, if that failed too, that they could not be put back, naming what to check in `NettingSets` and `Trades` by hand before any run.

After a commit the message gives the rows written to each input sheet, or *unchanged*. Each committed staging sheet reads `COMMITTED`, and the validation is used up: committing the same rows again needs a new validation. `Results`, `Checks` and the reconciliations are not changed.


\Needspace{6\baselineskip}

### Run after the commit

A commit **calculates nothing**. As after any edit of the inputs, `Results!C5` reads **OUT OF DATE**, `ResultsStatus()` reads `STALE` and the two exports refuse until you press **Run SA-CCR**. Then review `Checks` and `Results` as for any run: the netting-set statuses and the excluded trades come from the calculation's own validation, as described above.

Committed rows are ordinary inputs: you can edit them on `NettingSets` and `Trades` afterwards, which does not change the staging sheets. A later commit replaces the input sheets again, edits included.


\Needspace{9\baselineskip}

## Finding codes

Every finding of an import or a staging validation has a stable code. The import specification page lists every code with its meaning; they fall into these groups:

| Group | Codes | What they report |
|:----|:-------|:-------|
| Structural: nothing loaded | `CSV-ENCODING`, `CSV-SYNTAX`, `CSV-EMPTY`, `CSV-DELIMITER`, `CSV-HEADER-DUPLICATE`, `CSV-HEADER-MISSING`, `CSV-HEADER-UNKNOWN`, `CSV-FIELD-COUNT`, `CSV-TOO-MANY-RECORDS`, `CSV-TOO-LARGE` | The file cannot be read as a file of its kind: encoding, CSV syntax, header, delimiter, field count or a limit. |
| Values of a record | `CSV-BLANK-RECORD`, `CSV-REQUIRED`, `CSV-TYPE`, `CSV-RANGE`, `CSV-ALLOWED`, `CSV-LENGTH`, `CSV-FORMULA`, `CSV-DUPLICATE-KEY`, `CSV-RESERVED-KEY`, `CSV-ROW-RULE` | A blank record, a missing, malformed, out-of-range, unlisted, too long or formula-like value, a repeated or reserved key, or a broken rule between the fields of a record. |
| Between records | `CSV-NETTING-SET-UNKNOWN`, `CSV-ECON-MTM-LEG`, `CSV-ECON-MTM-VALUE`, `CSV-ECON-NETTING-SET`, `CSV-ECON-SETTLEMENT`, `CSV-ECON-ASSET-CLASS` | A leg whose netting set does not exist, or an economic trade whose legs break its rules. |
| Against Params | `CSV-CONTEXT-PARAMS`, `CSV-CONTEXT-MATURED`, `CSV-CONTEXT-EXPIRED`, `CSV-CONTEXT-RESET`, `CSV-CONTEXT-CURRENCY`, `CSV-CONTEXT-PAIR`, `CSV-CONTEXT-FX-NOTIONAL`, `CSV-CONTEXT-FACTOR`, `CSV-CONTEXT-REGIME` | `Params` cannot be read as a run reads it, or a trade is matured, expired or past its reset, lacks an FX rate, has an FX notional currency outside its pair, lacks a factor row, or does not fit its netting set's regime, its credit scale included. Staging validation only. |

\Needspace{12\baselineskip}

# Calculation anatomy: from trade leg to EAD


\Needspace{9\baselineskip}

## Traceable calculation path

Every accepted trade is enriched with normalized values and a supervisory interpretation. The workbook first calculates or assigns a **supervisory factor, supervisory delta, adjusted notional and maturity factor**. It then forms **effective notionals**, aggregates those by **hedging set** and, where relevant, **maturity bucket or reference**, produces **asset-class add-ons**, combines the add-ons into the **aggregate add-on**, and finally calculates **replacement cost, PFE and EAD** per netting set. The detailed views exist precisely to make this hierarchy challengeable.

```
Params + NettingSets + Trades
             |
             v
  Validate columns / dates / IDs / regimes / amounts
             |
             v
 TradeCalc: leg normalization and effective notionals
             |
             v
 HedgingSets + Buckets: grouping and regulatory offsets
             |
             v
 Results: add-ons -> multiplier -> PFE -> RC -> EAD
             |
             v
 Bank comparisons -> Reconciliation + TradeReconciliation
             |
             v
 Dashboard: the run in one page
             |
             v
 Current, unmodified run -> evidence CSV and report
```


\Needspace{9\baselineskip}

## Trade-level intermediary quantities

A **linear** trade's signed supervisory delta is ordinarily +1 for long or -1 for short. For options, the delta is obtained from the regulatory supervisory-option approach (the workbook exposes `SACCR_OptionDelta` and `SACCR_LambdaCRR` to inspect the building blocks). In IR and credit, adjusted notional depends on supervisory duration and the input notional; FX and other asset classes apply their respective definitions. The resulting **effective notional** combines adjusted notional, supervisory delta and the corresponding maturity factor. `TradeCalc` records both unmargined and, when applicable, margined versions to support comparison and cap analysis.

The basic supervisory duration function, when used, follows the continuous-rate expression:

\[
\mathrm{SD}(S,E)=\frac{e^{-0.05S}-e^{-0.05E}}{0.05},
\]

with the appropriate supervisory floor and product-specific date conventions. This is **not** a valuation-based DV01 or market-risk duration: it is a prescribed SA-CCR supervisory quantity.

For illustrative orientation, the maturity-factor structures are:

\[
\mathrm{MF}_{u}=\sqrt{\min(\max(M,\,10/250),\,1)},
\qquad
\mathrm{MF}_{m}=\frac{3}{2}\sqrt{\frac{\mathrm{MPOR}_{BD}}{250}}.
\]

The exact runtime uses configured parameters such as business days per year and minimum maturity business days; **do not use these display equations instead of the workbook's named parameter values or regime-specific conditions**. The margined factor applies only in the relevant margined treatment. In a `CCP` netting set where the CCP retains VM, a trade whose own horizon exceeds the MPOR uses that horizon in place of the MPOR in its margined factor, so `MF margined` on `TradeCalc` can differ between trades of the set.


\Needspace{9\baselineskip}

## Hedging sets, maturity buckets and correlation

Netting within a **hedging set** is not a simple sum across all trades. Regulatory rules differ by asset class and sub-class. The workbook exposes hedging sets, buckets and entity/reference dimensions so that offsetting, correlations and category distinctions can be reviewed. Pay special attention to an unexpected zero or low add-on: it could be valid netting or an incorrectly shared hedging identifier.

For interest-rate exposures, effective notionals are allocated to **three maturity buckets**, commonly represented as `D1`, `D2` and `D3`. The bucket formula with offsets has schematic form:

\[
\mathrm{EN}_{IR}=\sqrt{D_1^2+D_2^2+D_3^2 +
1.4D_1D_2+1.4D_2D_3+0.6D_1D_3}.
\]

Under **BCBS** the user may choose an alternative **sum of absolute bucket values** (`|D1|+|D2|+|D3|`) using `IRBucketOffset=FALSE`. Under **CRR**, that switch is not supported for a netting set with IR trades. Other asset classes apply their own aggregation rules using prescribed categories and correlations. **Never use an IR bucket formula as a general-purpose formula for equity, credit or commodities.**


\Needspace{9\baselineskip}

## Multiplier, PFE and replacement cost

The multiplier reflects portfolio MtM net of applicable collateral relative to aggregate add-on. The familiar supervisory form, when aggregate add-on is positive, is:

\[
\mathrm{Multiplier}=\min\left[1,\ f+(1-f)
\exp\left(\frac{V-C}{2(1-f)\mathrm{AddOn}_{aggregate}}\right)\right],
\quad f=0.05.
\]

The implementation handles input domains and exceptional cases; the user should inspect these through **Results**, **Checks** and the `FormulaTest` sheet rather than implementing the expression independently in an uncontrolled spreadsheet. PFE is the product of the multiplier and aggregate add-on. The exposure calculation applies alpha and, for a margined set, the appropriate unmargined cap.

![Replacement cost of a margined netting set on an illustrative value path, the add-on levels from trade to netting set, and the multiplier falling from 1 towards its 5% floor as collateral exceeds the value of the trades.](assets/rc_pfe.png){width=98%}

**Key interpretation:** a low multiplier does not imply the exposure can fall without bound, because the regulatory floor applies. A margined cap flag does not mean the netting set is invalid; it means the permitted unmargined ceiling constrains the margined calculation. A zero or blank EAD **with non-VALID status** is not the same as an eligible zero regulatory exposure.


\Needspace{9\baselineskip}

## Output field guide: Results

`Results` has **28 columns** and a row per input netting set, followed by one total per regime with netting sets. The principal measures include the effective regime, margining, trade count, MPOR, portfolio value `V`, recognized collateral `C`, replacement cost `RC`, IR/FX/CR/EQ/CO/OT add-ons, aggregate add-on, multiplier, PFE, alpha, margined EAD, unmargined cap, final EAD, cap-applied indicator and status. Read the **status column first**; any blank numerical values due to `INVALID`, `INCOMPLETE` or `NO TRADES` have a different interpretation from zero. `INVALID` and `INCOMPLETE` rows are highlighted in light red. The header row carries filter buttons, which work on the protected sheet; a run keeps them and applies a filter you set to the new rows.

**Run summary.** `Results!C4` holds the run summary: the time, the counts, the parameter set (`REGULATORY` or `WHAT-IF`), the input fingerprint, then the day count and the source of the FX rates, for example `| year 365 days, calendar TABLE | FX Manual`. `Results!C5` holds the freshness status.

![The right part of Results after a run of the demonstration portfolio: aggregate add-on, multiplier, PFE, alpha, margined EAD, unmargined cap, EAD, cap applied and status; two netting sets are INCOMPLETE.](assets/results_ead.png){width=98%}

**Totals:** the workbook creates `TOTAL CRR` and `TOTAL BCBS` as applicable. Each sums the figures of that regime's **VALID sets only** and reports a completeness statement such as “6 of 8 netting set(s) VALID; 2 withheld.” The two regime totals **must not be added together** because they may describe alternate regulatory views of the same trades. The message box and silent result line follow that policy.


\Needspace{9\baselineskip}

## Output field guide: TradeCalc, HedgingSets and Buckets

`TradeCalc` shows one row per input trade leg that has a TradeID (a row without one is reported only on `Checks`). Review at least: TradeID, netting set, asset class, `OK` or `EXCLUDED`, hedging set, bucket, supervisory factor, converted notional, converted MtM, S/E/M/T, supervisory duration, adjusted notional, supervisory delta, unmargined/margined maturity factors, effective notionals, stand-alone add-on, lambda and messages. Some fields are legitimately **not applicable** (for example option expiry for a linear trade; supervisory duration outside IR/credit; margined MF in an unmargined set).

The three sheets carry filter buttons on their header row, as `Results` does.

`HedgingSets` summarizes effective notionals, component add-ons and regulatory aggregation, including the three IR buckets where relevant. `Buckets` exposes the finer underlying, maturity or commodity-category contributions. Reconcile these data against the identifiers supplied on Trades before treating any numerical difference as a calculation bug.


\Needspace{9\baselineskip}

## Checks: interpret warnings and error messages

`Checks` has a severity field with **`ERROR`**, **`WARNING`** and **`INFO`** entries. An error can exclude a trade, invalidate a netting set or stop an entire run depending on its scope. A warning may record a defaulted optional parameter, interpreted remargining value or `WHAT-IF` deviation without invalidating the result. An INFO line provides trace context, for example the role, MPOR floor and rule of a client-clearing or CCP netting set, or the CCP share of a regime's total. The `Checks` table should be reviewed **on every run**, even when Results includes some valid EAD values.

Each line has eight columns: `Severity`, `Sheet`, `Row`, `Reference`, `Message`, **`Origin`**, **`Code`** and **`Field`**.

* **Origin** tells which operation wrote the line: **`RUN`** for **Run SA-CCR** and **Validate inputs**, **`IMPORT`** for **Validate staging**. The two never mix: each operation replaces all the lines of the other.
* **Code** is the stable diagnostic code of the rule, such as `TR-MATURED`, `NS-CODE` or `CSV-TYPE`. Codes are case-sensitive and keep their meaning from version to version; a rule whose meaning changes gets a new code, while the wording of a message can change. INFO lines have codes too, such as `NS-MPOR-INFO`, and a clean run writes one line, `RUN-CLEAN`. The codes group by prefix: `PRM-` and `LAYOUT-` for `Params` and the workbook's layout, `SF-` and `FX-` for the two tables, `TABLE-` for their rows, `NS-` for netting sets, `TR-` and `ECON-` for trades and economic trades, `RUN-` for the run, `CSV-` and `IMPORT-SUMMARY` for a staging validation.
* **Field** names the field the line concerns, such as `Sub-class` or `QCCP`, when it concerns one; it is blank for a rule that spans fields or the whole run. A rule shared by several fields has one code and names the field here.

**One line per violation.** A rejected trade gets one `ERROR` line for each rule it breaks, each with its code, rather than one line that joins every reason; its reason on `TradeCalc` still joins them. The error count of a run counts these lines, so it can be higher than the number of rejected trades; the number of rejected trades and the withheld EADs do not change. Filter `Checks` by `Code`, not by the wording of a message. The result line of `RunSACCR_Silent` lists the distinct ERROR and WARNING codes (see *Noninteractive calculation: RunSACCR_Silent*).

![Checks after a run of the demonstration portfolio: six ERROR lines, four for the three trades excluded and two for the netting sets whose EAD is withheld, each with its origin, code and field.](assets/checks.png){width=98%} After a staging validation, `Checks` shows its findings, each with a link to its staged cell (see *Validate staging*), and no longer the checks of the last run, until **Run SA-CCR** writes them again.

Recommended discipline: export or retain all error/warning rows; connect them to input table row numbers and trade/netting IDs; correct the **source input**; rerun the full calculation; compare fresh fingerprints; and document any consciously accepted warnings. `Checks` is protected and carries filter buttons on its header row; changing it after a maintenance unlock is **not** equivalent to resolving the input cause and also invalidates the integrity guard of the two exports.


\Needspace{12\baselineskip}

# The Dashboard

The `Dashboard` sheet, the first output tab, gives the last run in one page: what the exposure is, with whom, in which asset classes, how concentrated, how it grows from notional to EAD, and how clean the data are.

![The Dashboard after a run of the demonstration portfolio: the status and the parameter set at the top, the four cards, the run, the key figures per regime, the counts, the units, and the EAD by counterparty sector with its chart.](assets/dashboard.png){width=90%}


\Needspace{9\baselineskip}

## How the Dashboard is made

* **Written by every run.** A run that completes writes the Dashboard after its two reconciliations. Every operation, **Run SA-CCR**, **Validate inputs** and **Clear outputs**, first empties it, so a run that fails or stops leaves no Dashboard, and the sheet never shows part of a run.
* **No new calculation.** Every figure is a count, a sum, a share or a ranking of figures the run already wrote on `Results`, `TradeCalc` and `Checks`, of the inputs of `NettingSets`, or of the results of the two reconciliation sheets.
* **Each regime apart.** A table with amounts has one column per regime, CRR then BCBS, and amounts of the two regimes are never added; a chart shows one series per regime. A regime with no netting set shows zeros.
* **Valid netting sets.** Amounts, rankings and shares use the netting sets whose status on `Results` is `VALID`, whose sum is the regime's TOTAL row. Counts of netting sets use every netting set of the regime. Notional, maturity and market-value figures use the trade legs used (`OK` on `TradeCalc`) of the valid netting sets; a leg counts once.
* **Amounts in units.** Amounts are in **units of the reporting currency**, as on the input and output sheets, with a thousands separator and no decimals; every card and table title names the currency. Counts, shares and the Herfindahl index have no unit. A share of a total of zero, and an unused row of a ranking, are blank, not zero.
* **Out of date.** The status cell at the top shows the status of `Results` (`Results!C5`) and turns red when the inputs change after the run.
* **Protected.** The sheet is protected like the other output sheets; nothing on it is an input.


\Needspace{9\baselineskip}

## What the Dashboard shows

At the top, four **cards**: the EAD, the RC and the PFE of each regime, and the share of the regime's EAD held by its ten largest netting sets. Below, small tables, each with its chart beside or below it:

| View | Contents |
|:----|:--------|
| Run | Reporting date and currency, the time of the run, the default regime, the parameter set (`REGULATORY` with its ID, or `WHAT-IF`) and the input fingerprint. |
| Key figures | Per regime: EAD, RC and PFE; netting sets in total, valid, invalid, incomplete and with no trades; margined and unmargined; valid netting sets with the cap applied; distinct counterparties. |
| Counts and units | Trade legs read, used and excluded; netting sets without a regime; distinct counterparties; the rows of Checks by severity; the unit of the amounts. |
| 1. EAD by counterparty sector | EAD by the sector of `NettingSets`, with *Not classified* for a blank sector. |
| 2. EAD by clearing role | Bilateral, clearing member, client, CCP qualifying and CCP not qualifying. |
| 3. Add-on by asset class | The IR, FX, CR, EQ, CO and OT add-ons; one doughnut chart per regime. |
| 4. EAD composition | RC and PFE (before alpha), and the EAD of the margined and of the unmargined netting sets. |
| 5. Ten largest netting sets | Per regime, by EAD, with each one's share of the regime's EAD. |
| 6. Concentration | The share of the ten largest netting sets and the Herfindahl index: 1 for one netting set, 1/n for n of equal EAD. |
| 7. Maturity profile | The notional of the used legs under 1 year, from 1 to 5 years, and over 5 years, by M. |
| 8. Trades by asset class | Per regime, the used legs and their notional. |
| 9. Data quality and diagnostic codes | Excluded legs, netting sets whose EAD is withheld, and the five codes with the most ERROR and WARNING rows on `Checks`. |
| 10. Reconciliation | For each reconciliation, whether it was made, and its `MATCH`, `DIFFERENCE`, `MISSING IN CALCULATOR` and other rows. |
| 11. From notional to EAD | Per regime, each with its share of the notional: the notional of the used legs, their gross positive value, the value after netting, RC, PFE and EAD, and a sentence that reads it, such as *Netting removes X% of gross positive value; collateral a further Y%.* |

![The lower part of the Dashboard after the same run: the further tables per regime with their charts.](assets/dashboard_charts.png){width=90%}

The Dashboard is a summary for reading, not a source: the figures behind it are on the output sheets, and the formatted run report copies it (see *The formatted run report*). Do not use its rounded figures in place of those of `Results`.


\Needspace{12\baselineskip}

# Reconcile against a bank calculation


\Needspace{9\baselineskip}

## What the two reconciliations compare

K-SACCR provides **two independent comparison levels**:

1. **Netting-set reconciliation:** enter the bank's EAD and other high-level measures on `BankNettingSets`; `Reconciliation` compares those figures with `Results`, **by NettingSetID and regime**.
2. **Trade-leg reconciliation:** enter the bank's trade-leg quantities on `BankTrades`; `TradeReconciliation` compares them with `TradeCalc`, **by unique TradeID** and records the economic trade grouping ID for interpretation.

These are not substitute controls for one another. A close EAD could mask offsetting differences in inputs and components; a close set of leg-level inputs may still produce a different aggregated EAD if grouping, margin, correlation or cap treatment differs. Use both where the external system provides sufficiently granular data.

The bank's figures are typed on the two bank sheets or loaded from CSV files (see *Load the bank figures from a CSV file*). Both comparison mechanisms are **optional**, and both bank sheets are delivered empty. If no bank figures are provided, the calculator can run, but reconciliation will report that it was **not made**, rather than inventing comparisons. You can use the calculator alone for published examples and exploratory analysis.


\Needspace{9\baselineskip}

## BankNettingSets: scale, tolerances and fields

Before populating bank figures, enter the **amount scale** and two tolerances in the named settings at the top of `BankNettingSets`, `D4:D6`: `BankNS_Scale`, `BankNS_AbsTol` and `BankNS_RelTol`. The delivered values are a scale of 1, an absolute tolerance of **1 unit** and a relative tolerance of **1%** (`0.01`). The table's header is on row 8 and the figures start on row 9. For instance, if bank values are expressed in thousands while the calculator is in units, the amount scale is 1,000. The **multiplier is not an amount** and is **not scaled**. The absolute tolerance is stated in reporting-currency **units**, not necessarily in the units displayed in the bank extract; the relative tolerance is a unitless decimal fraction.

| Column | External metric | Reconciliation rule |
|:--|:--|:--|
| `NettingSetID` | Bank netting-set key | Match the calculator's exact set key; duplicate bank keys are diagnostic. |
| `Regime (BCBS / CRR)` | Bank method | Must match the effective calculator regime; cross-regime comparisons are **not** automatically accepted. |
| `EAD` | Bank final EAD | Amount, scaled by `BankNS_Scale`. |
| `RC` | Replacement cost | Amount. |
| `PFE` | Potential future exposure | Amount. |
| `AddOn aggregate` | Aggregate supervisory add-on | Amount. |
| `AddOn IR` | Interest-rate component | Amount. |
| `AddOn FX` | FX component | Amount. |
| `AddOn CR` | Credit component | Amount. |
| `AddOn EQ` | Equity component | Amount. |
| `AddOn CO` | Commodity component | Amount. |
| `AddOn OT` | Other-risk component | Amount; meaningful under CRR where supported. |
| `Multiplier` | PFE multiplier | **Dimensionless; not scaled.** |
| `Comment` | Analyst annotation; not read by the engine (its header is not checked) | Not read as a numerical comparison. |

![BankNettingSets with the bank's EAD for three netting sets: the buttons Import CSV, Export CSV template and Clear bank figures, the amount scale and the two tolerances, and one row with an ID the calculator does not have.](assets/banknettingsets.png){width=98%}

Blank bank figures are **not compared**, and are different from an explicit zero. The calculator creates comparison rows and population findings rather than silently assuming missing measurements equal zero.


\Needspace{9\baselineskip}

## BankTrades: scale, tolerances and fields

The analogous settings on `BankTrades` are `BankTR_Scale`, `BankTR_AbsTol` and `BankTR_RelTol`, with the same delivered values: 1, 1 unit and 1%. Amounts are converted to unscaled units; **M (years), supervisory duration, delta and maturity factors are not scaled**. Enter each bank metric against the **calculator's unique leg-level TradeID**—not just the aggregated economic-trade number—if a meaningful leg-level reconciliation is required.

| Column | External metric | Amount scale applied? |
|:--|:--|:--:|
| `TradeID` | Leg key | No |
| `Notional (rep ccy)` | Converted notional | **Yes** |
| `MtM (rep ccy)` | Converted MtM | **Yes** |
| `M (years)` | Remaining maturity | No |
| `Supervisory duration` | Supervisory duration | No |
| `Adjusted notional d` | Adjusted notional | **Yes** |
| `Supervisory delta` | Signed delta | No |
| `MF unmargined` | Unmargined maturity factor | No |
| `MF margined` | Margined maturity factor, when applicable | No |
| `Effective notional unmargined` | Unmargined effective notional | **Yes** |
| `Effective notional margined` | Margined effective notional | **Yes** |
| `Stand-alone add-on` | Trade-level amount | **Yes** |
| `Comment` | Analyst explanation; not read by the engine (its header is not checked) | No |

The `TradeReconciliation` output carries the **Economic trade ID** so that several legs can be examined together without incorrectly comparing one aggregate bank trade to only one calculator leg. If an input leg is `EXCLUDED`, do not force a numerical `MATCH`: first investigate the exclusion.


\Needspace{9\baselineskip}

## How a comparison is classified

For an applicable bank-versus-calculator metric, the tool calculates the absolute difference, **|calculator − bank|**, a relative difference, (calculator − bank) divided by the absolute bank figure where it is not zero, and a result. The rule depends on the kind of metric:

- **Amounts** (EAD, RC, PFE, add-ons, notionals, MtM, effective notionals): the bank figure is multiplied by the sheet's scale, and the figure is a `MATCH` if it meets the **absolute OR relative** tolerance. When the bank figure is zero there is no relative difference, so only the absolute tolerance applies.
- **Ratios** (`Multiplier` on BankNettingSets; `M`, supervisory duration, supervisory delta and both maturity factors on BankTrades): the scale is not applied and the absolute tolerance is **not** used. The figure is a `MATCH` only within the relative tolerance, or, when the bank figure is zero, only if the calculator figure is exactly zero.

A figure outside the applicable tolerance is a `DIFFERENCE`.

| Result on reconciliation | Meaning | Primary next action |
|:--|:--|:--|
| `MATCH` | Applicable comparison within the applicable tolerance (see above). | Check that source, regime, scale and coverage are correct; do not infer independent model certification from a match. |
| `DIFFERENCE` | Comparison outside the applicable tolerance. | Trace inputs, valuation, maturity, netting, categories and regulatory rules. |
| `MISSING IN CALCULATOR` | Bank key not represented by the corresponding calculator output. | Reconcile source populations and identifiers. |
| `NOT IN BANK FILE` | Calculator key has no bank-side population match. | Verify the bank extract's scope/filtering. |
| `DUPLICATE` | Duplicate identifiers make an unambiguous bank population comparison impossible. | Correct the duplicate-ID issue; do not select an arbitrary first row. |
| `REGIME MISMATCH` | Bank and calculator use different regulatory regimes. | Align CRR/BCBS before comparing amounts. |
| `WITHHELD` | Netting-set EAD is not usable because calculator status is not VALID. | Resolve input errors and rerun. |
| `INVALID INPUT` | A bank row or figure is unusable: missing ID, regime not `BCBS`/`CRR`, a figure that is not a number, or a row without figures. (An unusable scale, tolerance or header stops that reconciliation altogether: `FAILED`, with the reason in the sheet's C4.) | Correct the bank row or figure. |
| `EXCLUDED` | Calculator rejected the trade leg. | Resolve trade validation; do not treat blank figures as zero. |
| `NOT APPLICABLE` | The bank entered a figure the calculator does not have for that row (for example a supervisory duration for an FX or equity leg, or a margined maturity factor in an unmargined netting set). | Do not turn legitimate blanks into zero-valued comparisons. |

**Illustrative tolerance example:** bank EAD = 1,000,000; calculator EAD = 1,000,050; absolute tolerance = 100 reporting-currency units; relative tolerance = 0.001 (0.1%). The difference of 50 is within the **absolute** tolerance and is also within the relative tolerance; the result is `MATCH`. If the absolute tolerance were 10 but relative tolerance 0.001, it would still be `MATCH` because the relative criterion is satisfied. This example explains the classification rule, not a prescribed tolerance policy. With the delivered tolerances, 1 unit and 1%, an EAD of 5,300 against a calculator figure of 5,404.62 differs by 1.97% and is a `DIFFERENCE`.


\Needspace{9\baselineskip}

## Load the bank figures from a CSV file

The bank's figures can be typed on the two bank sheets or loaded from a CSV file. **BankNettingSets file...** and **BankTrades file...**, in the Import CSV menu of the K-SACCR tab, and the **Import CSV** button of each bank sheet, read a file you choose and write its records onto the sheet at once. There is no staging sheet: the sheet holds figures only, and the reconciliation reports what does not match.

* **The file.** A BankNettingSets file has one record per netting set: `netting_set_id`, `regime` (`BCBS` or `CRR`, required), `ead`, `rc`, `pfe`, `addon_aggregate`, the six add-ons from `addon_ir` to `addon_ot`, `multiplier` and `comment`. A BankTrades file has one record per trade leg: `trade_id`, `notional`, `mtm`, `maturity`, `supervisory_duration`, `adjusted_notional`, `supervisory_delta`, `mf_unmargined`, `mf_margined`, `effective_notional_unmargined`, `effective_notional_margined`, `standalone_addon` and `comment`. The columns are those of the sheet, in its order. The format is that of every CSV file of the workbook (see *CSV file format*); **BankNettingSets template...** and **BankTrades template...** in the Templates menu, or **Export CSV template** on the sheet, write the blank files `K-SACCR_BankNettingSets_v1.csv` and `K-SACCR_BankTrades_v1.csv`.
* **Amounts in the sheet's scale.** A file carries no scale and no tolerance: an amount is written in the scale set on the sheet, and the amount scale and the two tolerances above the table are neither read from the file nor changed.
* **All or nothing.** A file with a finding, a file with more records than the limit (10,000 netting sets, 50,000 legs) and a header-only file change nothing; the dialog lists the findings, and nothing is written to `Checks`.
* **A question first.** When the sheet holds figures, the workbook asks before replacing them, with **No** as the default.
* **How it writes.** The figures and comments of the sheet are cleared and the records written from the first row below the header. The identifier, the regime and the comment are written as text, so `00731` keeps its leading zeros; a figure is written as a number, read with the decimal point of the file in every locale; a blank figure is an empty cell, which the reconciliation does not compare. The rows below the first 200 netting sets or 500 legs, which the template formats with drop-down lists, get the number formats of the first row and no drop-down list. An error while the sheet is written puts the old figures back.
* **Unknown IDs.** An ID that is not on `NettingSets` or `Trades` is not a finding of the file: the reconciliation reports it as `MISSING IN CALCULATOR`.

The bank files are figures to compare, not inputs of a run: loading one changes neither the inputs nor `Results`, but it makes the reconciliations out of date until the next **Run SA-CCR**.


\Needspace{9\baselineskip}

## Clear bank figures

Both bank sheets are delivered **empty**, so that no demonstration figure is left beside a real portfolio. To empty a sheet again, press **Clear bank figures** on it, or choose **Clear BankNettingSets...** or **Clear BankTrades...** in the Import CSV menu of the K-SACCR tab. After a confirmation, with **No** as the default, every figure and comment of that sheet is cleared; the amount scale and the two tolerances stay. Run again to refresh the reconciliation: without bank figures, it reports that it was not made.


\Needspace{9\baselineskip}

## The reconciliation sheets

`Reconciliation` and `TradeReconciliation` have the run's summary in `C4` and their freshness status in `C5`, a block of summary tiles, and the comparison table from its header on row 11.

![Reconciliation after a run with bank figures for three netting sets: the summary and status, the tiles, with the measure shown (EAD) and the sums of the two rows compared, and the table, with a MATCH, a DIFFERENCE beyond the relative tolerance of 1%, a netting set missing in the calculator and the netting sets not in the bank file.](assets/reconciliation.png){width=98%}

**The tiles** sit above the table, each a label and a value: *Last reconciliation*, *Rows shown*, *Figures compared*, *Match* (green), *Difference* (red), *Not in bank file*, *Missing in calculator*, *Duplicate rows*, and *Other results*, which lists the other results with their counts, or *None*. A last row gives, for the measure shown, *Measure shown*, *Sum of bank figures*, *Sum of calc. figures*, *Sum of abs. diff.* and *Largest abs. diff.*:

* The tiles are linked to formulas on the sheet, so they follow the filter of the table and need no macro: *Rows shown* counts the rows the filter leaves.
* The sums and the largest difference concern one measure: filter the `Measure` column to one measure to read them; with several measures shown, the tiles say so. A factor measure, such as the multiplier, has no sum.
* The sums and the differences count **only the rows compared**, those whose result is `MATCH` or `DIFFERENCE`, so that the bank's sum and the calculator's cover the same rows. A row with a figure on one side only, `NOT IN BANK FILE` or `MISSING IN CALCULATOR` for example, is counted in its own tile and left out of the sums.

**The table.** The columns are the identifier, the bank row, the bank and calculator regimes, the calculator status, the measure, the bank figure (amounts in units of the reporting currency, the bank figure multiplied by the scale), the calculator figure, the **absolute difference**, `|calculator − bank|`, the **relative difference**, which keeps its sign, the result and a note; `TradeReconciliation` also gives each leg's Economic trade ID. The figures and the difference of an amount are shown with two decimals, those of a factor (the multiplier; M, the supervisory duration, the delta and the maturity factors) with six, whatever the order of the rows. A bank figure of 0 has no relative difference.

**Working on the sheets.** The two sheets are not protected. The header row carries an AutoFilter down to row 100,000, and a run applies a filter you set again to the new rows. You can also sort, write notes and add columns to the right of the table; a run, **Validate inputs** and **Clear outputs** rewrite only the table. Sorting the table or changing a cell of it makes the two exports refuse until the next run (see *What each sheet lets you change*). A reconciliation writes one row per figure, eleven per trade leg; when the rows it needs do not fit on the sheet, it is not made and the run message says how many it needs.


\Needspace{9\baselineskip}

## ReconSummary: one row per measure

`ReconSummary` sums the two reconciliations per measure: one table for the netting sets of `Reconciliation`, from EAD to the multiplier, and one for the trade legs of `TradeReconciliation`, from the notional to the stand-alone add-on.

![ReconSummary: one row per measure of the two reconciliations, with the rows shown, the count of each result and, over the rows compared, the sums and the largest absolute difference.](assets/reconsummary.png){width=98%}

| Column | Contents |
|:----|:--------|
| Rows shown | The rows of the measure that the filters of the reconciliation sheet leave. |
| Match, Difference, Other results | The rows of each result. |
| Sum of bank figures compared, Sum of calculator figures compared | Over the rows compared, `MATCH` and `DIFFERENCE` only, so that both sums cover the same rows. Amounts are in units of the reporting currency. |
| Sum of absolute differences, Largest absolute difference | Over the same rows; a factor has no sum, only its largest difference, with six decimals. |
| Row of the largest | The sheet row where the largest difference is. |

The sheet holds formulas only: it follows the filters of the two reconciliation sheets, so a filter on a result or a measure there changes its figures, and it needs no macro. It is protected as an output sheet.


\Needspace{9\baselineskip}

## Practical root-cause investigation sequence

For every material `DIFFERENCE`, follow a **top-down, then bottom-up** approach:

1. **Population and regime:** are the same trade legs in both systems, assigned to identical netting sets and regimes? Check `MISSING`, `NOT IN BANK FILE`, `DUPLICATE`, `WITHHELD` first.
2. **Data normalization:** do the same as-of date, currencies, FX rates, notionals, marks, signed directions and date conventions apply? Review converted notional/MtM on TradeCalc.
3. **Instrument treatment:** check asset class, category, option quote, option delta, supervisory duration, start/end/expiry and possible multi-leg decomposition.
4. **Margin and collateral:** check VM sign, NICA, TH, MTA, MPOR, clearing role, large/illiquid and disputes.
5. **Add-on aggregation:** check hedging-set and bucket assignment, offset option (where permitted) and regulatory correlations.
6. **Exposure build:** explain RC, aggregate add-on, multiplier, PFE, alpha and the cap flag.
7. **Documentation:** distinguish input issue, methodology difference, deliberate policy or conservatism, upstream data issue and suspected implementation defect. Retain the exact source and comparison evidence.

`ReconSummary` shows at once which measures differ, and by how much in total; filter `Reconciliation` or `TradeReconciliation` on one measure and on `DIFFERENCE` to list the rows behind it. For a trade-level issue, start with the metric furthest **upstream** that differs. If Notional or M is different, do not begin debugging the final stand-alone add-on. For a netting-set issue, work backwards from EAD to RC/PFE, and then PFE to multiplier/aggregate add-on.


\clearpage


# Evidence export, freshness and integrity


\Needspace{9\baselineskip}

## What the export is for

The evidence function produces **one semicolon-delimited CSV per completed run** that captures the run identity and the principal values needed to reproduce and review it. This is **not** a regulatory return, a full trade portfolio export, or a substitute for the release-certification evidence bundle. It is a compact, self-describing **run report** for audit and reconciliation.

The file uses **UTF-8 without BOM**, **CRLF** line endings, semicolon separators and a **dot** decimal separator, independent of Excel's regional settings. The first field in each line identifies the section. The first line identifies the format: `RUN;format;K-SACCR evidence 1`. Text that looks like an Excel formula (leading `=`, `+`, `-`, `@`, tab or carriage return) is escaped with a leading apostrophe to reduce spreadsheet-formula injection risk; numerical signs remain meaningful. Excel error values are serialized by their conventional names.


\Needspace{9\baselineskip}

## Sections in the evidence CSV

| CSV section | Typical contents |
|:--|:--|
| `RUN` | Format ID, run summary/classification, input and output fingerprints, reporting date/currency, default regime, the day count (`year_convention`, `calendar`, `holiday_table`), the source of the FX rates (`fx_source`), counts, totals per regime, reconciliation status and tolerances when comparisons were made. |
| `PARAMS` | The parameter codes/values used. |
| `SF_TABLE` | Supervisory-factor lookup table rows actually read. |
| `FX_TABLE` | FX conversion rows actually read, in units of the currency per 1 unit of the reporting currency. |
| `HOLIDAY_TABLE` | The header of the table of closing days and every closing day in its 100 rows. |
| `RESULTS` | Netting-set figures and status / applicable totals. |
| `RECONCILIATION` | Netting-set bank comparisons, where made. |
| `TRADE_RECONCILIATION` | Trade-level bank comparisons, where made. |
| `CHECKS` | Errors, warnings, information and what-if deviations, in the eight columns of `Checks`: with their origin, code and field. |
| `LIMITATIONS` | Limitations and caveats relevant to the run format. |

The bank comparison's amount scales and tolerances are included in the `RUN` section **when that reconciliation was made**, because they affect how differences are interpreted and enter the fingerprint. The workbook's current **file name is not used as run identity**: Save As could change a name without changing the calculation. Results of two day counts, or of two sources of FX rates, are different results: the `RUN` lines say which the run used. An evidence CSV written by 1.2.0 keeps its six-column `CHECKS` section; it stays the evidence of that run and must not be rewritten.


\Needspace{9\baselineskip}

## Freshness model: when is a result current?

K-SACCR stores **hidden workbook names** to distinguish the completed calculation from subsequent edits. The core indicators to understand are:

| Hidden name | Meaning |
|:--|:--|
| `SACCR_RunInputs` | Fingerprint of the calculator inputs used for the last completed run. Empty when no completed result exists. |
| `SACCR_ReconInputs` | Fingerprint of the netting-set bank comparison settings, bank figures and calculator context. |
| `SACCR_TradeReconInputs` | Fingerprint of the trade bank comparison settings, bank figures and calculator context. |
| `SACCR_RunOutputs` | Fingerprint of output cells from the run that are included in the evidence CSV; detects post-run output edits. |

`ResultsStatus()` returns **`CURRENT`** when the completed results match the input state and the outputs are still those the run wrote, **`EDITED`** when the inputs match but an output cell on Results, Checks or a reconciliation sheet was changed after the run, including by a staging validation, which rewrites Checks, **`STALE`** when the input state has changed, and **`NONE`** when there is no valid completed run. The visible `Results!C5` status formula checks the inputs only and says so: *Inputs unchanged since the last run (inputs …); outputs are checked when evidence is exported*. It becomes **OUT OF DATE** when an input changes. It does not watch the outputs, because their fingerprint would then be recomputed after every edit; an edited output is caught by `ResultsStatus()` and refused by the export. The Results label of the K-SACCR tab reads `Results!C5` and shows `CURRENT`, `OUT OF DATE` or `NONE`, never `EDITED`. The Dashboard shows the same status as `Results!C5`. The equivalent status cells on `Reconciliation!C5` and `TradeReconciliation!C5` watch bank data and comparison settings as well as calculator inputs. An empty or absent comparison is presented as no reconciliation, not a fictitious current comparison.

The input fingerprint covers dates, currency, parameter and lookup values, the day count and the table of closing days, the record of an FX download, the two input tables (including headers), Excel-error/number/text distinctions and the workbook date system. Because it covers the headers, the same inputs can have a different fingerprint in another version whose headers differ, as 1.2.0 and 1.3.0 do (see *Moving inputs from 1.2.0*). The staging sheets are not inputs of a run: loading, editing or validating them leaves `Results!C5` as it is, while a commit changes the input tables and makes it **OUT OF DATE**, and so do **Restore parameters** when it changes a value on `Params`, **Download prices** and a rebase of the FX table. The bank figures are not inputs of a run either: loading or clearing them makes the reconciliations, not `Results`, out of date. Moving Params rows without breaking the workbook names should still allow changes to be detected. Cosmetic formatting or a purely descriptive label may not count as a parameter change: do not confuse the fingerprint with a universal change-tracking or audit-log service.

**Manual calculation mode:** when Excel has recalculation pending, press **F9** and confirm status updates before export. Switching only the workbook's date system (**Use 1904 date system**) changes no input cell, so `Results!C5` can keep its earlier text until **Ctrl+Alt+F9** recalculates every formula; a run under the 1904 date system stops. The status formulas are designed to detect edits without a macro that erases Excel's Undo history, but they depend on calculation having taken place.


\Needspace{9\baselineskip}

## How to export evidence interactively

1. Finish a **Run SA-CCR** calculation and inspect `Checks` and the `Results` status.
2. If populated, inspect both reconciliations and their **CURRENT**/not-made status. Correct or intentionally clear bank-side data and **rerun** as appropriate.
3. Ensure Excel is not awaiting recalculation. Confirm no output result/check/comparison values have been manually edited after the run.
4. Click **Export evidence** (`ExportEvidence`) in the Evidence group of the K-SACCR tab, from any sheet. Choose a filename and directory that satisfy your local data-classification requirements. The tab disables **Export evidence** while `Results!C5` reads **OUT OF DATE** or there is no run, and while another operation runs; an enabled button does not mean the export will succeed, since the export makes every check below itself. The status bar shows its two steps, checking the run and writing the file.
5. Verify the resulting CSV exists, starts with the expected format record and contains the expected section labels. In a text editor, confirm its run identity and that numeric separators are invariant to locale.
6. Optionally, click **Export report** for a PDF or XLSX report of the same run (see *The formatted run report*).
7. Archive the CSV, and the report if any, together with any external bank source extracts, the workbook's file name and SHA-256, the review decision and the approved data-retention location.

If an export is **refused**, do **not** disable the guard or hand-edit the CSV to force success. Resolve the underlying stale/recalculation/output-integrity condition, rerun and export again.


\Needspace{9\baselineskip}

## Why an export can be refused

![The consistency checks governing a new run-evidence CSV.](assets/evidence_gate.png){width=96%}


The implementation deliberately refuses **Export evidence** and **Export report** when:

* Excel still has recalculation pending and the current input state cannot be trusted;
* `Checks` shows the findings of a staging validation instead of the checks of the last run;
* `ResultsStatus()` is not `CURRENT`, including after an input changed since the run (`STALE`), for example by a commit of staged rows, by **Restore parameters** or by **Download prices**, or an output cell was edited (`EDITED`), for example by sorting the table of a reconciliation sheet;
* either reconciliation is stale against its relevant bank input/settings, including after its bank figures were loaded from a file or cleared;
* an exported output cell on Results, Checks or a reconciliation sheet differs from what the run produced;
* another operational or file-write failure prevents generation of the complete report;
* for **Export report** only, a figure of the Dashboard differs from what the run wrote.

When the chosen file already exists, the new report is written whole to a new file beside it and then replaces the older one; if the replacement fails, the older file is put back unchanged and the first error is reported. A read-only file is left as it is.

A CSV that was once successfully exported does not remain evidence for a **new** workbook state after subsequent edits.


\Needspace{9\baselineskip}

## The formatted run report

**Export report**, in the Evidence group of the K-SACCR tab (macro `ExportReport`), writes a formatted report of the last run for a reviewer who reads results rather than parses them. The evidence CSV stays the machine-readable record of the run; the report is a readable view of the same run.

* **PDF or XLSX.** The Save As dialog offers the two file types; the extension you choose decides the format.
* **The same run as the CSV, never another.** The report is refused in the same cases as **Export evidence**, with the same reasons (see *Why an export can be refused*), and also when a figure of the Dashboard was changed after the run. **Export report** is enabled when **Export evidence** is.
* **The workbook is left as it is.** The report is built in a new, temporary workbook. The sheets, names, cells, selection, calculation mode and clipboard of the K-SACCR workbook are not changed, and the results stay current. While the report is written, no other operation can start, and the status bar shows its progress.
* **Safe replacement.** An older file of the same name is replaced only once the new one is saved; a failure leaves it as it was and leaves no partial file.

| Section | Contents | XLSX | PDF |
|:----|:--------|:----|:----|
| Run | The run summary, the status of the results, the reporting date and currency, the default regime, the regulatory parameter set, the day count, the calendar and the table of closing days, the source of the FX rates, the fingerprints of the inputs and outputs, the state of each reconciliation and the units. | | |
| Dashboard | Every table of the Dashboard, with its title, unit, headers and values. | Values, in the format of each cell | Text, in the format of each cell |
| Charts | The ten charts of the Dashboard, as pictures. | None | On a sheet of their own |
| Totals | The two TOTAL rows of `Results`. | | |
| Results | Every netting set, with its 28 columns, and the TOTAL rows. | | |
| Checks | A table with a row per code, its highest severity and its number of rows; then the rows of `Checks`, grouped by code. | Every row | At most 10 rows of each code |
| Reconciliation | The summary and the rows, when a reconciliation was made. | Every row | Every row that is not `MATCH` |
| Trade reconciliation | The summary and the rows, when a reconciliation was made. | Every row | The 50 rows with the result `DIFFERENCE` and the largest absolute difference |

* **What the PDF leaves out is said.** Under the summary of `Checks` and of each reconciliation, a line says how many rows the PDF leaves out; the summaries still count every row, and the XLSX has them all.
* **Formats.** Dates are `yyyy-mm-dd`. The XLSX holds values and number formats only: no formula, no macro and no link to the workbook, with amounts in `#,##0.00` and ratios in six decimals. The PDF writes its numbers as text that does not depend on the regional settings: a comma between thousands and a point before the decimals, for example `1,234,567.89`.
* **The PDF.** A4 landscape, one page wide; the header row of each table repeats on every page, and the footer gives the time of the run, the input fingerprint and the page number. To picture the charts, the export shows the Dashboard and scrolls each chart into view; the Dashboard's zoom, scroll position and selection are put back afterwards.
* **Units.** Amounts are in units of the reporting currency, unscaled, and each table says so; ratios and times are not scaled.

The report has no limitations section: the limitations stay in the evidence CSV. Archive the report with the evidence CSV of the same run; it does not replace it.


\Needspace{9\baselineskip}

## Silent export for controlled automation

`ExportEvidence_Silent(filePath)` creates the same evidence content without a file-dialog. On success, it returns a line like `RESULT=OK; lines=...; file=...`. On a guard failure, it returns `RESULT=REFUSED; reason=...`. It is used for controlled testing and external orchestration, not as a way to bypass the interactive guard.

`ExportReport_Silent(filePath)` writes the formatted report in the same way, a PDF or an XLSX by the extension of `filePath` (`.pdf` or `.xlsx`); it refuses another extension and the cases of the interactive command. It returns `RESULT=OK; format=PDF|XLSX; sheets=N; charts=N; page_setup=OK; file=...`, with `page_setup=FAILED (...)` when a printer driver refused a page setting, which does not stop a PDF, or `RESULT=REFUSED; reason=...`.

A reproducibility test can run the same inputs, export twice to separate files, and compare their contents under the established output identity. Preserve the generated CSV in its original byte encoding whenever deterministic comparison matters.


\Needspace{12\baselineskip}

# Automation and worksheet functions


\Needspace{9\baselineskip}

## Noninteractive calculation: RunSACCR_Silent

For a caller that automates Excel on Windows, `RunSACCR_Silent` calculates without message boxes, logs to the Immediate window and raises failures to the caller. Unlike **Run SA-CCR**, it activates no sheet. Its structured result line includes operation status; error/warning counts; trade-leg and economic-trade usage; withheld netting sets; separate CRR/BCBS total EAD; input fingerprint; run classification; reconciliation states; and cleanup status.

The result line has this form, the values shown as `…`:

```text
RESULT=OK; operation=run; errors=…; warnings=…;
trades_used=…; trades_read=…;
economic_trades_used=…; economic_trades_read=…;
incomplete=…; total_ead_crr=…; total_ead_bcbs=…;
inputs=…; run_class=REGULATORY; reconciliation=NONE;
trade_reconciliation=NONE; year_convention=365;
calendar=TABLE; holiday_table=empty; fx_source=Manual;
codes=TR-MATURED,NS-TRADE-INCOMPLETE; cleanup=PASS
```

It shows the format; its figures depend on the inputs. Parse the fields by name, not the whole line. `errors` and `warnings` count the ERROR and WARNING lines on `Checks`, one per violation; `codes` lists their distinct codes, in the order met, and is empty when there are none; `cleanup` is always last. `year_convention`, `calendar` and `holiday_table` give the day count (`holiday_table` is `unused` under `TARGET`, `empty`, or the number of entries and a fingerprint, such as `3 days, 1A2B3C4D`), and `fx_source` the source of the FX rates, as in the run summary. `RESULT=STOPPED` means input loading could not complete; inspect Checks and the raised error as applicable. Reconciliation state may be `DONE`, `NONE`, `FAILED` or `UNAVAILABLE` according to input presence and sheet/configuration conditions. A stopped or failed run must never be reported as successful merely because an earlier workbook cell still contains a number.

`ValidateInputs_Silent` and `ClearOutputs_Silent` run **Validate inputs** and **Clear outputs** in the same way: `operation=validate` with the same fields as a run, `codes` included, or `operation=clear` with only `cleanup`. None of the silent entry points shows a message or the progress in the status bar.

The silent entry points are useful for testing and orchestration, but a caller must manage Excel lifecycle, capture exceptions, verify cleanup (including changes to Excel's global state), and record the workbook's file name and SHA-256. They do not transform the workbook into a headless stand-alone engine independent of Excel.


\Needspace{9\baselineskip}

## Noninteractive CSV import

The CSV import has silent entry points that ask nothing and return a result line, for the same kind of controlled automation. `Kind` is `"NettingSets"` or `"Trades"`, except for the two bank entry points.

- `ReadCsv_Silent(Kind, FilePath)`\
  Checks a file and writes nothing: `RESULT=ACCEPTED; records=…` or `RESULT=REJECTED; findings=…; first=…`.
- `ImportCsv_Silent(Kind, FilePath, [ReplaceEdits])`\
  Loads a file onto its staging sheet: `RESULT=LOADED; records=…; findings=…`, `RESULT=REJECTED` for a structural finding, or `RESULT=REFUSED` when the staging sheet has changes and `ReplaceEdits` is not `True`.
- `ValidateStaging_Silent()`\
  Validates both staging sheets: `RESULT=VALID; errors=0; …`, `RESULT=INVALID; errors=…; …; first=…` or `RESULT=REFUSED; reason=…`.
- `CommitStaging_Silent()`\
  Commits **without a confirmation**, under the same conditions as **Commit import**: `RESULT=COMMITTED; nettingsets=…; trades=…` (0 for a sheet left as it was), `RESULT=REFUSED; reason=…`, or `RESULT=FAILED; error=…; rollback=…`, the rollback being `NONE`, `VERIFIED` or `FAILED` with what could not be put back.
- `StagingValidationCurrent()`\
  `True` while a validation without findings still matches the staged rows, the input sheets and `Params`.
- `ExportCsvTemplate_Silent(Kind, FilePath)`\
  Writes a blank template: `RESULT=OK; file=…`.
- `ExportImportSpecification_Silent(FilePath, [ReplaceExisting])`\
  Writes the import specification page: `RESULT=OK; …`, or `RESULT=REFUSED` when the file exists and `ReplaceExisting` is not `True`.
- `ImportBankCsv_Silent(Kind, FilePath, [ReplaceExisting])`\
  With `Kind` `"BankNettingSets"` or `"BankTrades"`, loads the bank figures onto their sheet: `RESULT=LOADED; records=…`, `RESULT=REJECTED; findings=…; first=…` when the file does not follow the contract, or `RESULT=REFUSED; reason=…` when the file has no record, or the sheet holds figures and `ReplaceExisting` is not `True`.
- `ClearBank_Silent(Kind)`\
  Clears every figure and comment of that bank sheet, without a confirmation: `RESULT=CLEARED; sheet=…; rows=…`. The scale and tolerances stay.

A caller that uses `CommitStaging_Silent` takes on the confirmation the command asks for: it should check the staged row counts first, and run `RunSACCR_Silent` afterwards, since a commit calculates nothing.


\Needspace{9\baselineskip}

## Noninteractive parameter restore

`RestoreRegulatoryParameters_Silent()` does what **Restore parameters** does after its confirmation, **without a confirmation** or a message: it writes the regulatory parameter set back to `Params`, with the same scope (see *Restore the regulatory parameters*). Only the cells that differ from the set are written, and each is read back; on any failure the cells kept before the write are written back and read back. It returns:

- `RESULT=RESTORED; parameters=…; factor_rows=…`\
  with the number of parameters and factor-table rows that were changed;
- `RESULT=UNCHANGED`\
  when `Params` already holds the set;
- `RESULT=REFUSED; reason=…`\
  when nothing was written: a `Params` workbook name is missing, or a factor row stands below the regulatory rows of the factor table;
- `RESULT=FAILED; error=…; rollback=…`\
  with the rollback `NONE`, `VERIFIED` or `FAILED` with what could not be put back, as for `CommitStaging_Silent`.

Nothing is recalculated: after `RESULT=RESTORED`, `ResultsStatus()` returns `STALE` until the next run. A caller takes on the confirmation the command asks for: it should know which values it is about to overwrite, and run `RunSACCR_Silent` afterwards.

Three more entry points change `Params` or the view without a question:

- `DownloadClosingPrices_Silent([Source])`\
  Does what **Download prices** does after its confirmation, from `Source`, `"Closing price"` or `"ECB reference rate"`, or from the source chosen on the sheet: `RESULT=PRICED; rates=…; without_price=…`, or `RESULT=REFUSED; reason=…` with no rate changed. It waits inside the call; with real closing prices, which Excel delivers only when no macro runs, it refuses with *no price came*, so use the button for that source.
- `RebaseFxTable_Silent([Currency])`\
  Rebases the FX table on `Currency`, or on the reporting currency: `RESULT=REBASED; currency=…; divisor=…; rates=…`, `RESULT=UNCHANGED` when its rate is already 1, or `RESULT=REFUSED; reason=…`.
- `SetNavigationBar_Silent(Shown)`\
  Shows (`True`) or hides (`False`) the navigation bar on every sheet: `RESULT=SHOWN; sheets=…` or `RESULT=HIDDEN; sheets=…`, `RESULT=REFUSED` while an operation runs. Every save still writes the bar shown.


\Needspace{9\baselineskip}

## Worksheet function API

The published user-accessible worksheet functions are the `SACCR_*` functions listed below. Inputs are generally dimensionless fractions or year quantities where stated; invalid inputs return appropriate Excel errors (`#NUM!`, `#VALUE!` or `#REF!`) instead of silently returning misleading numbers. These functions help reproduce **building blocks**; they are not a substitute for the entire netting-set engine or for regulatory classification.

- `SACCR_NormCDF(x)`\
  Standard normal cumulative distribution function. *Verify:* numeric domain and spreadsheet error behavior.
- `SACCR_SupervisoryDuration(S, E, [FloorYears])`\
  Supervisory duration for IR/credit adjusted notional. *Verify:* S/E in years, E > S, floor conventions.
- `SACCR_MaturityFactor(M, [Margined], [MPOR_BD], [MinBD], [BDPerYear])`\
  Unmargined/margined maturity factor. *Verify:* M in years; MPOR in business days.
- `SACCR_OptionDelta(P, K, T, Vol, IsCall, IsBought, [Lambda])`\
  Supervisory option delta. *Verify:* pricing-unit conventions, T, option sign and CRR lambda if applicable.
- `SACCR_LambdaCRR(P, K, IsInterestRate, [Threshold])`\
  CRR rate/commodity lambda shift. *Verify:* regulatory threshold and applicability.
- `SACCR_CDODelta(A, D, IsLong)`\
  Credit-tranche supervisory delta. *Verify:* attachment/detachment fractions and direction.
- `SACCR_IREffectiveNotional(D1, D2, D3, [Rho12], [Rho23], [Rho13])`\
  IR effective notional across maturity buckets. *Verify:* correct bucket populations, correlations and regime option.
- `SACCR_FactorAggregation(AddOns, Rhos)`\
  Factor/correlation-based aggregation. *Verify:* matching vector dimensions and valid correlations.
- `SACCR_Multiplier(VminusC, AddOn, [FloorPct])`\
  PFE multiplier. *Verify:* V-C, aggregate add-on and floor, not EAD directly.
- `SACCR_ReplacementCost(V, C, [Margined], [TH], [MTA], [NICA])`\
  Replacement cost. *Verify:* C, VM/NICA conventions and margin flag.
- `SACCR_BusinessDays(StartDate, EndDate, [Holidays], [UseTarget])`\
  Business days after `StartDate` up to and including `EndDate`, negative for a reversed interval. Saturdays and Sundays are never counted; the TARGET closing days are counted out unless `UseTarget` is `FALSE`, and `Holidays` is a date, a range or an array of further closing days. A date before 1 March 1900, or, under the TARGET rule, an interval that touches a year before 1999, gives `#NUM!`. *Verify:* the calendar matches the one on `Params`.

Arguments in square brackets are optional.

**FormulaTest.** The `FormulaTest` sheet tests these functions against Excel's own formulas, within the tolerance in `E4`. **Randomize inputs** draws every input of every test within its range, as values, from a seed written beside the tolerance (`H4`): type a seed and press the button to replay a draw. The ranges are in an editable table on the sheet; a range that leaves a function's domain is refused with a message. **Restore defaults** writes the delivered inputs back.

These worksheet functions are the workbook's stable public interface. The macros this manual describes (the commands of the K-SACCR tab and of the buttons on the sheets, `CheckNettingSetsCsv`, `CheckTradesCsv`, `ProtectWorkbook`, `ToggleNavigationBar`, `RunSACCR_Silent`, `ValidateInputs_Silent`, `ClearOutputs_Silent`, `ExportEvidence_Silent`, `ExportReport_Silent`, the silent CSV entry points, `ImportBankCsv_Silent`, `ClearBank_Silent`, `RestoreRegulatoryParameters_Silent`, `DownloadClosingPrices_Silent`, `RebaseFxTable_Silent`, `SetNavigationBar_Silent`, `ResultsStatus()` and `StagingValidationCurrent()`) are the documented entry points. Other VBA procedures and module members are internal and can change between versions; do not call them.


\Needspace{12\baselineskip}

# Troubleshooting and controlled recovery


\Needspace{9\baselineskip}

## Diagnostic table

| Symptom | Likely cause / first check | Recommended action |
|:--|:--|:--|
| The K-SACCR tab is missing | Another workbook is active: the tab is shown only while the K-SACCR workbook is the active workbook. | Activate the K-SACCR workbook, in a supported desktop Excel for Windows (see *Supported environment*). |
| The workbook opens on another add-in's tab | An add-in loaded its tab after K-SACCR's and kept the selection beyond the three seconds the workbook waits. | Click the **K-SACCR** tab, or press **Alt, K, S**. |
| Column A and the navigation bar are gone | The bar was hidden with **Navigation bar** on the tab or the triangle in `B1`. | Press **Navigation bar** again, or click **▶** at the top left of the sheet. A copy saved and opened again shows it. |
| The lower part of the navigation bar is not in view | Only the top block, `Home` and the inputs, stays in the frozen rows; the rest scrolls with the rows. | Scroll the sheet back to the top, or use **Go to**. |
| Excel takes no input while a command runs | The command shows its progress in the status bar and holds the keyboard and mouse until it ends. | Wait for its message. |
| The K-SACCR tab is shown, but nothing on it works, or a button or a triangle on a sheet does nothing | Macros are disabled or blocked, or the file is still marked as downloaded. | Unblock the file after the hash check, enable macros as your policy allows, and use the delivered workbook or an unchanged copy. Until then, use the navigation bar, the links on `Home` and the sheet tabs to read the sheets. |
| A command is grayed out | Another operation is still running, or there is nothing it can act on: no staged rows (**Validate staging**), no `VALID` staging sheet (**Commit import**), or `Results` out of date or never run (**Export evidence**, **Export report**). | Wait for the operation to finish; otherwise do what the Status labels show is missing, for example run again before exporting. See *Enabled states*. |
| A Status label does not match its sheet | Calculation is set to manual, so `Results!C5` has not been recalculated; or the tab has not been read again since the change. | Press **F9**, then select another sheet and come back, which reads the tab again. |
| After a VBA error message, the Status labels and enabled states stop changing | The error reset the workbook's code, and the tab could not be found again in this session. | The commands still work, and each checks its own conditions. Save your copy, close it and open it again to restore the tab's updates. If the error recurs, note its text and report it to whoever supplied the workbook. |
| A VBA error such as “Compile error”, “Ambiguous name” or a missing reference | The workbook's VBA code or references were changed, or the file is damaged. | Do not repair the code. Go back to your reference copy, check its SHA-256, and report the problem to whoever supplied the workbook. |
| Excel refuses to change a cell, delete a row or insert a column | The cell is locked, or the sheet does not allow the action (see *What you can edit: workbook protection*). | Change only the yellow input cells; clear cells instead of deleting rows. Unprotect a sheet only for deliberate maintenance. |
| Broken workbook name or missing header | A named range, header or column was edited, moved or deleted, after a sheet was unprotected. | Undo the change, or enter the inputs again in a copy of your reference copy. |
| `Results` is `OUT OF DATE` | Input, Params, header or date-system change after last run. | Press F9 as needed; rerun; do not rely on old figures. |
| `Results!C5` still reads unchanged after **Use 1904 date system** was switched | Only the date system changed, so Excel need not recalculate C5. | Press **Ctrl+Alt+F9**; a run under the 1904 date system stops, so clear the option, check the dates and run again. |
| `ResultsStatus()` is `EDITED` | An output cell on Results, Checks or a reconciliation sheet was changed after the run, the table of a reconciliation sheet was sorted, or Checks shows a staging validation. | Run SA-CCR again; never edit calculated cells. Filter the reconciliation tables rather than sort them before an export. |
| `ResultsStatus()` is `NONE` | No valid completed run; failed/cleared run. | Review Checks and operation status; complete a fresh successful run. |
| Trade `EXCLUDED` | Failed domain, date, code, pair, group or currency check. | Read its TradeCalc/Checks messages and the Code column, correct source and rerun. |
| `TR-FACTOR-REGIME` on a credit trade | A rating `AAA` to `CCC` in a CRR netting set, or a step `CQS1` to `CQS6` in a BCBS one. | Use the credit scale of the netting set's regime (see *SF_Table: supervisory-factor register*). |
| More ERROR lines on `Checks` than rejected trades | A trade that breaks several rules has one line per rule. | Expected; group the lines by `Code` or by the trade's row. |
| A run stops with `FX-REPORTING-RATE` | The rate of the reporting currency is not 1, typically after `ReportingCcy` was changed without a rebase. | Accept the rebase offered when you change the reporting currency, press **Download prices**, or divide every rate by that of the reporting currency. |
| Amounts in a currency are far off, by about the square of a rate | A rate entered the 1.2.0 way, reporting currency per unit of the currency. | Enter units of the currency per 1 unit of the reporting currency, as the market quotes it (see *FX_Table: conversion into reporting currency*). |
| A run stops with `PRM-DAYCOUNT`, `PRM-CALENDAR` or `PRM-HOLIDAY` | A `YearConvention` or `Calendar` outside its list, `TARGET` with a reporting date before 1999, or an entry of the table of closing days that is not a date. | Correct the cell the message names on `Params`. |
| **Download prices** changes nothing | No connection, a proxy that answers with its own page, an Excel without `STOCKHISTORY`, or the reporting date or currency changed while the prices were awaited. | Read the message; try the other source, or type the rates. |
| **Download prices** leaves some rates | The source has no rate for those currencies. | The message lists them: enter them yourself and check them. |
| Netting set `INCOMPLETE` | One or more associated trades rejected. | Trace every excluded leg; do not use the withheld EAD. |
| Netting set `INVALID` | Invalid netting data (including the CCP columns and role `MULTI`), duplicate set IDs, a settled-to-market trade in a margined set, or unsupported regime treatment. | Read the netting set's lines on Checks; correct netting terms, IDs, roles and regime/options. |
| `NO TRADES` | No trade rows refer to that netting set (rejected trades make it `INCOMPLETE` instead). | Confirm the trade population, the `NettingSetID` values on `Trades` and the as-of scope. |
| Reconciliation `REGIME MISMATCH` | Bank comparison is across BCBS/CRR. | Align regimes; never compare by EAD amount alone. |
| Reconciliation `DUPLICATE` | Duplicate population key on one side. | Audit all duplicate rows and identifiers. |
| Unexpected `DIFFERENCE` | Amount scale, FX quote, date convention, margin setting, class or aggregation. | Apply top-down/bottom-up reconciliation sequence. |
| Export `REFUSED` | Stale or edited data/results, a staging validation on Checks, a commit, a parameter restore or a price download since the run, a sorted reconciliation table, pending recalc or failed comparison; for the report, a Dashboard changed after the run. | Restore consistency; calculate; rerun and retry; keep guard enabled. |
| A reconciliation was not made, and the run message gives a number of rows | The comparison needs more rows than the sheet holds. | Reconcile a smaller population, for example a part of the bank file at a time. |
| Bank sheet Import CSV changes nothing | A finding in the file, more records than the limit, or a header-only file; the whole file is refused. | Read the findings in the dialog, correct the file and load it again. |
| A NettingSets file is refused with `counterparty_sector` missing | Its header is neither exactly the version 1 columns nor the version 2 columns. | Write the header of one version exactly, for example from the blank template. |
| `CSV-CONTEXT-FX-NOTIONAL` | The notional currency of an FX trade, or the reporting currency for a blank one, is not a currency of its pair. | Enter the notional of one of the pair's legs, in that leg's currency. |
| Import CSV: *Nothing was loaded* | A structural finding: not UTF-8, invalid CSV, another delimiter (`CSV-DELIMITER`), a missing, unknown or repeated column, a wrong field count, or a file over a limit. | Correct the file, for example save it as UTF-8 with `;` between fields and the exact header, and import it again. |
| Many `CSV-TYPE` findings on numbers or dates | The file has decimal commas, thousands separators or local dates, as Excel's own CSV save can write. | Write numbers with a decimal point and dates as `yyyy-mm-dd`, with a program that follows the file format. |
| Import CSV asks whether to replace changes | The staging sheet was edited since its last load, or holds rows not loaded from a file. | Answer **No** to keep them; **Yes** replaces them with the file's records. |
| Validate staging: *Nothing was validated* | A staged cell holds a formula, or both staging sheets are empty. | Enter the value itself as text; import a file first. |
| `CSV-NETTING-SET-UNKNOWN` on many legs | The netting sets were not staged and are not on `NettingSets`, or their IDs differ. | Import the NettingSets file too, or correct the IDs; IDs are compared ignoring case. |
| A staging sheet reads `OUT OF DATE` | A staged row, `NettingSets`, `Trades` or `Params` changed after the validation. | Undo the change, or press **Validate staging** again. |
| Restore parameters: *Nothing was restored* | A factor row stands below the factor table, a `Params` workbook name is missing or broken, or a write failed. | Read the message. Remove an added factor row by hand after a maintenance unlock, then restore again; for a broken name, start again from your reference copy. After a failed write, the message says whether `Params` was put back and checked; if not, check `Params` before any run. |
| Restore parameters says nothing was changed | `Params` already holds the regulatory set. | Nothing to do. A run that is still `WHAT-IF` has an alpha override on `NettingSets`; see the *What-if:* warnings on `Checks`. |
| Commit import: *Nothing was committed* | No validation without findings that still holds, or Excel did not finish recalculating. | Press **Validate staging**, correct every finding and commit again; press **F9** first if asked. |
| Commit import reports a failure | A write or its read-back failed. | Read the message: if the input sheets were put back and checked, nothing changed; if not, check `NettingSets` and `Trades` by hand, or start again from your reference copy, before any run. |
| A commit replaced an input sheet you meant to keep | Its staging sheet still held rows, for example from an earlier import. | Undo cannot take a commit back: re-enter the sheet, or import its file again. Empty a staging sheet with a header-only file before committing. |
| No externally visible output but old cells remain | Failure while clearing a protected/invalid output sheet. | Treat the result as invalid; use **Clear outputs** (K-SACCR tab or the button on an output sheet), or a copy of your reference copy, and run again. |
| A larger portfolio is slow | Large population and/or inflated Excel UsedRange. | See the size note under *Known limitations*; delete stray rows far below the last trade (the cells, not only their contents). |


\Needspace{9\baselineskip}

## When a run stops

A stopped operation is materially different from an excluded trade. A run can stop at the **workbook-level** when mandatory Params names, reporting date/currency, schema, error-valued settings, the 1904 date system or other structural prerequisites are unusable. In these cases, calculation cannot safely proceed. Capture the on-screen/Immediate-window error, preserve the source, fix the **root cause** and perform a clean rerun. Do not copy output cells from an earlier valid run into the newly failed workbook.


\Needspace{9\baselineskip}

## Safe manual recovery procedure

1. Save a **diagnostic copy** of the problematic workbook if organizational rules allow; do not overwrite an audit snapshot.
2. Write down the workbook's file name, the SHA-256 of the delivered file, Excel version/bitness, operation, status and error text.
3. Inspect `Checks`, `Results!C4:C5` and reconciliation summaries; determine whether the issue is **schema**, **input**, **regime**, **calculation**, **comparison** or **export**.
4. If the structure is broken, enter the inputs again in a copy of your **reference copy** rather than attempting to restore every named range by hand.
5. Correct source input or configuration. Never edit calculated cells to make a comparison match.
6. Repeat validation and a complete calculation, and verify the result is CURRENT.
7. Rerun both comparisons where input exists; check output integrity and export only after successful recomputation.
8. Record the corrective change and its effect, especially when previously reported EAD was incomplete or stale.


\Needspace{9\baselineskip}

## Known limitations at v1.3.0

* Basel rules are validated only where the published CRE99 examples exercise them. CRR rules are validated only where the independent implementation or the production reconciliation covers them (see *Version, evidence and validation boundary*); the others, the cap of a margined netting set, client clearing, CCP trade exposures and settled-to-market trades among them, are implemented and illustratively tested.
* Windows desktop Excel 64-bit is the supported target; 32-bit is best effort and not certified; Mac and web are not supported.
* The tool uses user-supplied marks, market values, dates, classifications, factor data and collateral amounts; it does not source or price these. The FX rates are the one exception: **Download prices** can write closing prices (Microsoft 365) or the ECB reference rates, which the user remains responsible for.
* FX derivatives use one selected notional leg; selecting the correct leg when both are foreign is the **user's** responsibility.
* FX options enforce the stated quote ordering, including alphabetical order for two unranked currencies, which may not match every local market convention.
* Digital options are not supported under either regime.
* Optional reductions/caps for certain qualifying sold options and sold credit protection are not implemented; using the ordinary method may produce a higher exposure.
* Multiple margin agreements in one netting set and one shared agreement across multiple netting sets are not modeled; declarative input is rejected. So is a netting set mixing margined and settled-to-market trades.
* Trade exposures to a CCP are calculated for a clearing member; QCCP status, retained VM and settlement to market are the user's declarations. Risk weights are stated, not applied; default-fund contributions, their own funds requirements and the cap that combines them with trade exposures are not supported. A CCP entered with role `N` is treated as bilateral.
* Multi-level client structures are rejected when declared with role `MULTI`; entered as `CM` or `CLIENT`, they cannot be detected.
* One *large or illiquid* flag covers more than 5,000 trades and illiquid collateral, so it can overstate the MPOR where the trade-count floor does not apply, for example for a CCP. Under BCBS a client facing its clearing member keeps the 10-day floor, a conservative interpretation.
* No 1904 date system. The only built-in calendar is the EUR TARGET calendar, by rule, from 1999; any other closing days are entered in the table of 100 rows on `Params`. The numerator of the day count is `ACTUAL` only. MPOR and the other floors in business days use `BusinessDaysPerYear` whatever the calendar.
* Variable-notional schedules, collateral haircuts, rating mapping and risk-driver decomposition must be prepared by the user.
* Unlisted multi-name credit positions (baskets that are not an index listed on a recognized exchange) are not supported and cannot be detected; relabeling one as an index or a credit quality step gives a wrong add-on in a `REGULATORY` run (see *Credit, commodity, inflation and other-risk coding*).
* Simplified SA-CCR, original exposure method, CCP default-fund exposures, SFT, CVA and RWA are beyond v1.3.0.
* There is no fixed trade limit, but no certified capacity either. On a 64-bit test machine, a synthetic portfolio of 500,000 trade legs in 1,000 netting sets calculated in about 114 seconds; times depend on the machine. The drop-down lists and input validation of the sheets cover `NettingSets` rows 5 to 10,004 and `Trades` rows 5 to 100,004, and a used range far below the data slows every run. The time of the reconciliations, the Dashboard and the report on large portfolios has not been measured.
* **CSV import replaces whole sheets.** A commit replaces `NettingSets`, `Trades` or both with the staged rows: there is no merge and no update by key, so a partial file removes every row it does not hold. An input sheet whose staging sheet holds no rows is left as it is. Rows with findings are never committed, a commit calculates nothing, and **Undo** cannot take it back.
* `NettingSets` and `Trades` are imported through the staging sheets, the bank figures directly onto their sheets; `Params` and the reconciliation settings are entered on their sheets. The workbook writes blank templates, not CSV files of the inputs it holds. Other CSV dialects, such as a comma-separated file or Excel's own CSV save with decimal commas or local dates, are rejected.
* An import holds at most 10,000 netting sets and 100,000 legs, and a bank file 10,000 netting sets or 50,000 legs, in files of at most 128 MiB. Loading, validating, committing and reconciling files at these limits has not been timed.
* Staged rows cannot be deleted on the protected staging sheets; a staging sheet is emptied by importing a header-only file. A staging validation leaves some rules to the calculation's own validation at the next run (see *Validate staging*).
* The K-SACCR tab is shown only while the K-SACCR workbook is active, and needs macros enabled for any command, label, enabled state, **Go to** entry or the **Navigation bar** toggle. Its Results label follows `Results!C5`, which checks the inputs only: it never shows `EDITED`, and the two exports can be enabled for an export that will refuse. The enabled states are an aid: every command checks its own conditions. After a VBA error the tab may stop updating its labels and states until the workbook is reopened; its commands still work.
* **Restore parameters** puts back the whole regulatory set at once, not a single value, and refuses while a factor row stands below the factor table. Write down a what-if value before you restore if you will need it again.
* Workbook protection has no password and prevents accidental edits only. On a protected sheet, rows and columns cannot be deleted and columns cannot be inserted, resized or formatted; filtering works only on `Trades`, `Basel_vs_CRR` and the five output sheets, through their existing AutoFilter. The two reconciliation sheets are not protected at all. Changing a locked value needs a deliberate maintenance unlock.
* The navigation bar's lower sections scroll with the rows; only `Home` and the inputs stay in view. Hiding the bar lasts for the session only.
* The Dashboard and the report summarize a run; they add no calculation, and the PDF report leaves rows out of `Checks` and the reconciliations, as it says. The report's PDF depends on the printer driver for its page setup.
* **Download prices** needs Microsoft 365 for closing prices, or access to the ECB data portal for the reference rates; the ECB does not publish every currency.
* The counterparty sector is a declaration for the breakdowns only; no SA-CCR rule reads it.


\Needspace{12\baselineskip}

# Operational control checklists


\Needspace{9\baselineskip}

## Checklist: before a first business run

- [ ] Check that the workbook's SHA-256 equals the expected value, confirmed through a channel you trust, before the first save.
- [ ] Use Windows desktop Excel in the supported environment; comply with macro policy.
- [ ] The workbook opens on `Home` and the **K-SACCR** tab is on the Ribbon, with its Status labels filled; the navigation bar is in column A.
- [ ] Keep the delivered file unchanged as a reference copy; never change the workbook's VBA code.
- [ ] Save an untouched synthetic demonstration copy; keep operational work separately.
- [ ] Set and independently check reporting date, currency, default regime, day count (`YearConvention`, `Calendar` and the closing days), reference parameters and FX rates, quoted in units of the currency per 1 unit of the reporting currency, with their source; unless a what-if analysis is intended, the run is `REGULATORY` (use **Restore parameters** if it is not).
- [ ] Validate every netting-set ID, regime, margin role, VM/NICA sign and supported legal structure; for a CCP, its QCCP status and whether it retains VM; the counterparty sector where used.
- [ ] Validate trade-leg IDs, economic IDs, one MtM leg per group, notionals (zero or more, with the side in `Direction`), dates, settlement type and reset dates, FX quote orientation and notional currency, and the credit scale of each regime.
- [ ] Confirm that source extracts and output quantities use the documented unit conventions.
- [ ] If the inputs come from CSV files, follow the import checklist below.
- [ ] Run **Validate inputs**, resolve errors and review all warnings.


\Needspace{9\baselineskip}

## Checklist: before committing an import

- [ ] Work in a copy; the reference copy stays unchanged, and a commit cannot be undone.
- [ ] The files follow the import specification: UTF-8, `;` between fields, the exact header of their version, numbers with a decimal point, dates as `yyyy-mm-dd`.
- [ ] Each staging sheet holds exactly the rows you mean to commit; a staging sheet left with rows from an earlier import has been emptied with a header-only file.
- [ ] Each file is the **whole** sheet: a commit replaces the input sheet and keeps no row the file does not hold.
- [ ] **Validate staging** reports no finding, and both sheets that hold rows read `VALID`, not `OUT OF DATE`, in `D7` and in the Status group of the K-SACCR tab.
- [ ] The confirmation's row counts for `NettingSets` and `Trades` are those you expect.
- [ ] After the commit: **Run SA-CCR**, then review `Checks` and the netting-set statuses, since the calculation's own validation still applies.
- [ ] The committed CSV files are kept with the run's evidence.


\Needspace{9\baselineskip}

## Checklist: before accepting an EAD

- [ ] `Results` is **CURRENT** after a successful completed run.
- [ ] The relevant netting set is **VALID**; `INCOMPLETE`, `INVALID` and `NO TRADES` have been addressed appropriately.
- [ ] The reported netting-set portfolio is complete; separate CRR/BCBS totals are not combined.
- [ ] Material differences in V, C, RC, add-ons, multiplier, PFE, alpha and cap are understood.
- [ ] FX rates, supervisory factors, days, thresholds and classifications have an attributable source.
- [ ] TradeCalc/HedgingSets/Buckets substantiate the key numerical steps.
- [ ] `Checks` has been reviewed, by code, and material warnings or what-if deviations are recorded.
- [ ] Bank comparison metrics were normalized to the same units, regime and day count, with appropriate tolerance policy; the bank sheets hold only the bank figures intended, and `ReconSummary` has been read.
- [ ] An independent reviewer has evaluated material exceptions and the assurance boundary.


\Needspace{9\baselineskip}

## Checklist: before exporting and archiving

- [ ] No pending Excel recalculation; status cells and the Status labels of the K-SACCR tab have refreshed (`Results: CURRENT`, which does not by itself rule out an edited output).
- [ ] Results and relevant reconciliations are **CURRENT**, or bank comparisons were intentionally not made.
- [ ] Calculated results, Checks, the Dashboard and reconciliation outputs have not been manually modified, nor the reconciliation tables sorted, after the run.
- [ ] Interactive or silent export succeeds **without disabling a guard**.
- [ ] CSV contains `RUN;format;K-SACCR evidence 1`, expected sections, correct currency, day count, FX source and regime-specific totals.
- [ ] Where a report is archived too, it was exported from the same run (its footer gives the input fingerprint).
- [ ] The archive identifies input data provenance, run scope, the workbook's file name and SHA-256, and the reviewer decision.
- [ ] Data classifications, confidentiality, retention and access controls are satisfied.


\clearpage

# Delivered files and their roles


\Needspace{9\baselineskip}

| File | Main purpose | What it is **not** |
|:--|:--|:--|
| **`K-SACCR_1.3.0.xlsm`** | The calculator, with its K-SACCR Ribbon tab, its navigation bar, its `Home` sheet and a synthetic demonstration portfolio; the bank sheets are empty. Nothing is installed with it. | Not a store for your evidence: export it. |
| **This user manual (PDF)** | Teach a user how to check, operate, interpret and troubleshoot the version. | Not proof that the workbook was tested. |
| **`SHA256SUMS.txt`** | The SHA-256 of the workbook and of this manual, to check that the files arrived unchanged. | Not proof of who made the files (confirm the values through a separate trusted channel), nor a check of a copy you have saved since. |
| **Evidence CSV** (you create it) | Document a particular **calculator run** and its inputs/outputs/comparisons. | Not a regulatory return or a test record. |
| **Run report, PDF or XLSX** (you create it) | A readable view of the same run: Dashboard, totals, Results, Checks by code, reconciliations. | Not the machine-readable record: the evidence CSV is; the PDF leaves rows out. |
| **Import specification** (you save it from the workbook) | Describe the four CSV files, with an example of each, to whoever prepares them. | Not a delivered file, and not a check of a file: the import and **Validate staging** check files. |
| **CSV templates and input files** (you create them) | Carry `NettingSets` and `Trades` into the workbook through the staging sheets, and the bank figures onto their sheets. | Not a backup of the workbook: `Params` and the reconciliation settings are not in them. |

A released version is never changed: a correction comes as a new version, with its own file name, SHA-256 and manual.


\Needspace{12\baselineskip}

# Regulatory references and terminology


\Needspace{9\baselineskip}

## Principal regulatory/documentary sources

| Source | Role in K-SACCR v1.3.0 |
|:--|:--|
| [Basel Framework CRE52](https://www.bis.org/basel_framework/chapter/CRE/52.htm) | Basel SA-CCR calculation, the version effective 1 January 2023. |
| [Basel Framework CRE50](https://www.bis.org/basel_framework/chapter/CRE/50.htm) | Definition of a qualifying CCP. |
| [Basel Framework CRE54](https://www.bis.org/basel_framework/chapter/CRE/54.htm) | Client clearing and CCP rules: MPOR floors, the horizon when a CCP retains VM, and the risk-weight rules stated on Checks. |
| [Basel Framework CRE99](https://www.bis.org/basel_framework/chapter/CRE/99.htm) | **Published numerical examples 1–5**, the external reference values for the Basel rules. |
| Regulation (EU) No 575/2013 (CRR), consolidated text as registered | EU CRR Articles 272–280f and related rules, and Articles 300–311 on exposures to CCPs, in the consolidated text of 26 June 2026. |
| Commission Delegated Regulation (EU) 2021/931, as amended by (EU) 2025/855 | Technical standards including the CRR option lambda-shift methodology. |
| EBA Q&A 2023_6962 | Regulatory clarification informing NICA/cap treatment. |
| R package `SACCR` 3.4 | The independent implementation whose values validate CRR rules; an evidence source, not a regulatory text. |
| FINREP counterparty sectors | The codes of the optional counterparty sector. |

When a conflict arises between general explanatory prose and the **actual applicable legal text**, investigate the exact legal rule and date rather than treating the manual as a replacement for the regulation. This manual and the `Basel_vs_CRR` sheet record K-SACCR's chosen interpretations and conservative choices; legal applicability remains the bank's responsibility.


\Needspace{9\baselineskip}

## Compact glossary

| Term | Meaning in this manual |
|:--|:--|
| **SA-CCR** | Standardised Approach for Counterparty Credit Risk. |
| **EAD** | Exposure at Default under the supported supervisory calculation. |
| **RC** | Replacement Cost, calculated for a netting set. |
| **PFE** | Potential Future Exposure = multiplier × aggregate add-on. |
| **MtM / V** | Mark-to-market value of a leg or netting-set portfolio, with specified sign/currency. |
| **C** | Recognized collateral value under the relevant replacement-cost/cap context. |
| **VM** | Variation Margin; positive when net held, negative when net posted. |
| **NICA** | Net Independent Collateral Amount, after the relevant eligibility/segregation treatment. |
| **TH / MTA** | Margin threshold / minimum transfer amount. |
| **MPOR** | Margin Period of Risk measured in business days. |
| **Supervisory delta** | Prescribed signed risk sensitivity used for SA-CCR notional scaling, not a full valuation-model Greek. |
| **Adjusted notional** | Supervisory scaling of input notional, with asset-class methodology. |
| **MF** | Maturity Factor, with margined/unmargined conventions. |
| **Hedging set / bucket** | Regulatory aggregation grouping / maturity or asset category. |
| **CQS** | Credit Quality Step, used for CRR credit supervisory factors. |
| **CCP / QCCP** | Central counterparty / qualifying CCP; the status sets the risk weight, not the exposure value. |
| **STM** | Settled-to-market: a contract that settles its exposure and resets its value to zero on set dates. |
| **CRR / BCBS** | EU CRR / Basel CRE52 calculation regimes. |
| **REGULATORY / WHAT-IF** | Parameter-reference classification; not a certification grade. |
| **CURRENT / EDITED / STALE / NONE** | Completed-run status: current, an output edited after the run, inputs changed since, or no completed run. |
| **INVALID / INCOMPLETE / NO TRADES / VALID** | Netting-set calculation status (INCOMPLETE is followed by a count, INVALID by a count or a reason); only VALID produces EAD. |
| **Economic trade ID** | Shared identifier linking modeled legs of one economic trade. |
| **CSV contract version 3** | The columns and rules of the four CSV files, described by the import specification page: NettingSets files of version 2 (version 1 still read), Trades, BankNettingSets and BankTrades files of version 1. |
| **Counterparty sector** | The optional FINREP sector of a netting set's counterparty (`CB`, `GG`, `CI`, `OFC`, `NFC`, `HH`), used by the Dashboard and the report only. |
| **Staging sheet** | `Import_NettingSets` or `Import_Trades`: where a CSV file is loaded as text, checked and corrected before a commit. Never read by a run. |
| **Staging status** | `EMPTY`, `LOADED`, `EDITED`, `VALID`, `INVALID`, `OUT OF DATE` or `COMMITTED`, shown in `D7` of each staging sheet and in the Status group of the K-SACCR tab. |
| **Commit** | **Commit import**: replacing whole input sheets with the staged rows of a current validation without findings; no merge, no calculation. |
| **Finding origin** | The Origin column of `Checks`: `RUN` for a run or **Validate inputs**, `IMPORT` for a staging validation. |
| **Diagnostic code** | The Code column of `Checks`: a stable, case-sensitive identifier of the rule a line reports, such as `TR-MATURED`; the Field column names the field. |
| **Maintenance unlock** | Unprotecting a sheet on purpose to change a locked cell; the next operation, or the macro `ProtectWorkbook`, protects it again. Not needed by **Restore parameters**. |
| **K-SACCR tab** | The workbook's own Ribbon tab (**Alt, K, S**): every command, the Status labels, **Go to** and the **Navigation bar** toggle. Selected at opening, shown while the workbook is active; works only with macros enabled. |
| **Status labels** | The three read-only labels of the tab: `Results: CURRENT`, `OUT OF DATE` or `NONE`, and the status of each staging sheet. |
| **Enabled state** | Whether the tab offers a command now; an aid only, since every command checks its own conditions. |
| **Restore parameters** | The command that writes the regulatory parameter set back to the reference parameters, the two conventions and the factor table on `Params`, after a confirmation. |
| **Home** | The first sheet, on which the workbook opens: scope, the steps of a run and a link to every sheet. |
| **Navigation bar** | Column A of every sheet: a link to every sheet, `Home` and the inputs at the top in the frozen rows, the other sections scrolling with the rows. Hidden and shown by the **Navigation bar** toggle or the triangle in `B1`. |
| **Dashboard** | The output sheet that gives the last run in one page, each regime apart. |
| **ReconSummary** | The sheet that sums the two reconciliations per measure, over the rows compared. |
| **Run report** | The PDF or XLSX that **Export report** writes of a current run. |
| **YearConvention / Calendar** | The days of the conventional year and the calendar on which the days of S, E, M and T are counted: `TABLE`, `TARGET` or `TARGET+TABLE`. |
| **Download prices** | The `Params` button that writes the FX rates from closing prices or the ECB reference rates. |
| **TradeID** | Unique input **leg** identifier and trade comparison key. |
| **Evidence CSV** | Semicolon-delimited versioned run report; not a regulatory return. |


\Needspace{9\baselineskip}

## End-user quick-reference card

**Start:** check the SHA-256 → unblock → open and enable macros → keep the delivered file as reference copy → work in a copy; change only yellow cells. Every command is on the **K-SACCR** tab (**Alt, K, S**); the navigation bar, **Go to** or the links on `Home` open any sheet.\
**Configure:** AsOfDate, ReportingCcy, Regime, YearConvention, Calendar and FX table (units per 1 reporting currency, typed or **Download prices**) → check the locked supervisory parameters against the reference values, or press **Restore parameters** → verify netting terms → load trade legs.\
**Import (optional):** Templates > Import specification... or a blank template → write UTF-8 files with `;`, a decimal point and `yyyy-mm-dd` dates → Import CSV for both files → correct → Validate staging until no finding → Commit import.\
**Calculate:** Validate inputs (opens Checks; read Code and Field) → resolve `ERROR` → Run (opens Results) → read **status** before numerical EAD → Dashboard → trace intermediates.\
**Reconcile:** normalize bank currency/units/regime/day count → BankNettingSets/BankTrades, typed or Import CSV → rerun → tiles, `MATCH`, `DIFFERENCE`, population exceptions, ReconSummary → Clear bank figures when done.\
**Archive:** ensure CURRENT and unedited outputs → Export evidence (and Export report) → record the workbook's file name and SHA-256, the data and the review conclusion.\
**Escalate:** unsupported instrument, missing legal/market input, or withheld EAD is an **explicit limitation**, not a blank cell to work around.

---

**Document control:** K-SACCR 1.3.0 user manual, delivered with `K-SACCR_1.3.0.xlsm`; the SHA-256 of both is in the `SHA256SUMS.txt` delivered with them. Each version of K-SACCR has its own manual.
