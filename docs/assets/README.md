# 🖼️ Documentation images

The images in this directory illustrate K-SACCR. They do not certify the
engine or replace the [methodology](../methodology/README.md), which holds the
regulatory sources, the support matrix and the traceability of each rule. Each
section below says what its images show, how they were made and when to update
them.

<a id="screenshots"></a>

## 📸 Screenshots

The images below show sheets of the workbook built from
`src/workbook/SACCR_Template.xlsx`, with the demonstration portfolio. They were
taken in Excel for Windows on 2026-10-16, except `SACCR_Basel_vs_CRR.png`,
retaken on 2026-10-17 to show the colours of #296 and on 2026-10-20 after the
revision of its rows (#389).

| Image | Shows |
| --- | --- |
| `SACCR_NettingSets.png`, `SACCR_Trades.png` | The input sheets |
| `SACCR_Params.png`, `SACCR_Params_Supervisory.png` | Run parameters, and the supervisory table |
| `SACCR_README_sheet.png` | The in-workbook guide |
| `SACCR_Results_AddOns.png`, `SACCR_Results_EAD.png` | Results after a run: RC and add-ons, then PFE, EAD, cap and status |
| `SACCR_TradeCalc_Inputs.png`, `SACCR_TradeCalc_AddOns.png` | Every trade-level quantity of that run: dates, duration, adjusted notional and delta, then maturity factors, effective notionals and add-on |
| `SACCR_HedgingSets.png`, `SACCR_Buckets.png` | The hedging-set and bucket detail of that run |
| `SACCR_Checks.png` | The errors of that run: excluded trades and withheld netting sets |
| `SACCR_BankNettingSets.png`, `SACCR_Reconciliation.png` | A reconciliation: the bank's EAD for BCBS-EX1, copied from Results (MATCH), 5,400 for BCBS-EX3 (DIFFERENCE), a netting set the calculator does not have (MISSING IN CALCULATOR), and the netting sets not in the bank file |
| `SACCR_TestCatalogue.png`, `SACCR_FormulaTest.png` | The reference-test catalogue and the worksheet-function unit tests, all passing |
| `SACCR_Basel_vs_CRR.png` | The first 11 rows of the regime comparison |
| `SACCR_Evidence_CSV.png` | The start of the evidence CSV of that run, opened in Notepad: the RUN and PARAMS sections and the start of SF_TABLE |

The input sheets show the template's own content, except BankNettingSets,
which is empty in the template: its rows were entered for the reconciliation
screenshot. The output sheets show a run of that portfolio on that day.

- **Processing:** cropped to the sheet area (no window title, sheet tabs or
  status bar; for the CSV, no menu bar or scroll bar) and reduced to a
  256-colour PNG palette without dithering. They are not otherwise edited.
- **Number format:** the sheet values follow the regional settings of the
  machine that took them (decimal comma, thousands point); the workbook itself
  reads and writes locale-independent values, as the CSV shows.
- **Use:** the root README shows Results and Checks, and in collapsible
  sections the calculation detail, the reconciliation, the evidence file, the
  input sheets (NettingSets, Trades, Params, the supervisory table and the
  README sheet), the regime comparison and the two test sheets.
- **CRE99 Example 3:** the demonstration workbook shows an EAD of 5,405
  (thousands), the published figure 5,406. The demo enters the WTI trade with
  actual dates over 365 days a year (TradeCalc shows M = 0.7507); the
  regression fixture enters 270 days over 360, exactly 0.75 years as printed,
  and matches 5,406.

Retake a screenshot when the layout or content of its sheet changes, and keep
the same processing.

<a id="ead-reference-card"></a>

## 🧮 EAD reference card

`SACCR_EAD_reference.svg` summarises the SA-CCR formulas and supervisory
parameters as K-SACCR implements them. What is supported, and on what
authority, is in the methodology: the
[support matrix](../methodology/README.md#support-matrix), the
[supervisory parameters](../methodology/README.md#parameters) and the
[regime differences](../methodology/README.md#regime-differences); the card
summarises them and is not a source.

- **Dates:** the card gives S, E, M and T in years of 365 days. That is the
  default of the `DaysPerYear` parameter on Params, which accepts any value
  from 360 to 366; the engine uses the value set.
- **Sources:** the card's source line names CRE52, the CRR and Delegated
  Regulation (EU) 2021/931. Its MPOR line also relies on CRE54.8, CRE54.12,
  CRE54.14 to CRE54.17, registered as `CRE54` in the
  [source register](../methodology/README.md#source-register).

The SVG is the source. Its values were checked against the methodology
parameter table on the release branch on 2026-10-15, when the SVG last
changed; update the card when a supervisory parameter or formula changes.

The workbook carries the card as a picture on its README sheet, under the
"Reference card" heading, so a user never has to leave Excel (#281).
README!D23 links to C114, a cell behind the lower part of the picture:
Excel scrolls only until the target cell is in view, so a link to the
heading would leave the picture below the window. The picture is a PNG of
this SVG: when the SVG changes, render it again and replace the picture in
the template (`xl/media/image1.png`, shown 1640 px wide, the width of
columns C:D):

1. Open the SVG in an HTML page that shows it at 1800 x 1444 px and take a
   screenshot with headless Chromium at a device scale factor of 1, in a
   window taller than 1444 px; crop it to 1800 x 1444.
2. Quantize it to 256 colours (Pillow `quantize(colors=256,
   method=MEDIANCUT, dither=NONE)`) and save it with `optimize=True`, as for
   the screenshots above.

<a id="saccr-uses-in-crr"></a>

## 🧭 Where SA-CCR is used in the CRR

`SACCR_uses_in_CRR.svg` shows the parts of the CRR that use SA-CCR. Solid
lines lead to those that take the risk-based exposure value of a netting set
as calculated; dashed lines to the leverage ratio and the output floor, which
recalculate SA-CCR with their own adjustments (see the calculation contexts
below):
- counterparty credit risk under the standardised approach (Art. 107);
- exposures to CCPs (Art. 301–311);
- large exposures (Art. 390);
- the leverage ratio, with RC net of eligible cash variation margin only and a
  PFE multiplier of 1, except for client trades cleared through a QCCP
  (Art. 429c);
- CVA risk, through the standardised method under CRR2 and BA-CVA under CRR3
  (Art. 382–386);
- the output floor of CRR3, 50% of standardised RWA in 2025 rising to 72.5%;
  until 31 December 2029, contracts whose un-floored exposure value comes from
  the internal model method may use α = 1 in SA-CCR for the floor
  (Art. 92(3), 465(4)).

Each part is coloured by the regulation that brought it in, CRR2 (Regulation
(EU) 2019/876, SA-CCR applied from 28 June 2021) or CRR3 (Regulation (EU)
2024/1623, applied from 1 January 2025), and names what K-SACCR does with it:
the roadmap milestone, or out of scope. It is an orientation picture, not a
source: K-SACCR calculates only the exposure value at its centre.

- **Origin:** drawn for this repository from the CRR. The idea of a single
  exposure value feeding several frameworks follows the European Banking
  Federation's annex on the use of SA-CCR across the prudential framework; no
  part of that figure is reproduced.
- **Sources:** Art. 429c(3)–(5) and 465(4) were checked on
  EUR-Lex on 2026-10-17, in the consolidated CRR of 26 June 2026 (the version
  registered as `CRR` in the
  [source register](../methodology/README.md#source-register)) and in
  Regulation (EU) 2024/1623 (CRR III). The other references name the parts
  of the CRR that take the exposure value; the picture is not a source for
  their rules.
- **Separate calculation contexts:** the leverage ratio keeps α = 1.4 and sets
  the PFE multiplier to 1 (Art. 429c(5)); RC still follows Art. 275, with
  only eligible cash variation margin received (Art. 429c(3)) and, for
  qualifying client trades cleared through a QCCP, the wider collateral of
  Art. 429c(4a). The α = 1 of Art. 465(4) applies only to the output floor
  and only where the same exposures are calculated with the IMM in the
  un-floored total risk exposure amount. Neither is the risk-based SA-CCR
  that K-SACCR calculates.

The SVG is the source, shown by the README and the website (`docs/index.html`);
`SACCR_uses_in_CRR.png`, used by the Wiki, is rendered
from it at 1600 x 1040 px with headless Chromium at a device scale factor of
1, then quantized to 256 colours as the screenshots above. Render it again
whenever the SVG changes.

<a id="ccr-framework"></a>

## 🏛️ Counterparty credit risk framework

`CCR_framework.svg` shows where K-SACCR belongs in counterparty credit risk,
from exposures to capital requirements, in four steps:

1. two exposure engines, K-SA-CCR for derivatives and K-SFT for SFTs;
2. a classification by counterparty, feeding the ordinary and CCP branches;
3. the ordinary capital engine (SA or IRB risk weights) and CCP capital
   (trade exposures and default-fund contributions, qualifying or not);
4. the aggregation into the capital requirement for counterparty default
   risk and its RWA.

CVA risk appears below them as a memo item only: it belongs to market risk
and is excluded from that total.

- **Origin:** drawn on 2026-10-26 at the maintainer's request (#462), as one
  picture shared by K-SACCR, K-SFTs and K-Risk Engines. It is an orientation
  picture, not a source: it does not certify implemented functionality, and
  K-SACCR calculates only the K-SA-CCR box.
- **Sources:** its reference line names Basel CRE20, CRE22, CRE30-36,
  CRE51-54 and MAR50, and CRR Art. 111-191, 192-241, 271-311 and 381-386.
  The risk weights of its CCP table and the Basel citation of the QCCP cap
  are a summary, not checked against the CRR text for this picture.
- **Rendering:** the SVG is the source, shown by the README.
  `CCR_framework.png` is rendered from it with headless Chromium at a device
  scale factor of 2 (3200 x 3056 px), then quantized to 256 colours as the
  screenshots above. The same pair is used in the other two repositories:
  change it there too when it changes here.

<a id="sa-ccr-explainers"></a>

## 🧩 SA-CCR explainers

Three pictures explain how the exposure value is built. Like the EAD reference
card they summarise the [methodology](../methodology/README.md) and are not a
source; the card holds the parameters and the full formulas.

| Image | Shows | Used by |
| --- | --- | --- |
| `SACCR_aggregation.svg` | How a netting set is aggregated: each trade assigned to its netting set, asset class and hedging set, then the effective notional, hedging-set, asset-class and aggregate add-ons, PFE and EAD | README ("How a run works"), the website, the Wiki (SA-CCR Overview) |
| `SACCR_EAD_tree.svg` | The exposure value broken down into RC, PFE and α, down to the trade-level inputs | The methodology, the Wiki (SA-CCR Overview) |
| `SACCR_RC_PFE.svg` | The replacement cost of a margined netting set on an illustrative value path, the add-on levels and the PFE multiplier, plotted from its formula | The Wiki (Replacement Cost and Collateral, Potential Future Exposure); the user manual from 1.0.1 |

- **Origin:** drawn for this repository from CRE52 and CRR Art. 274–280f and
  285, after commonly used teaching diagrams of SA-CCR (an aggregation flow, a
  breakdown of the EAD, an RC chart); no part of any of them is reproduced.
  The value path in `SACCR_RC_PFE.svg` is illustrative, not data.
- **Rendering:** the SVGs are the source. `SACCR_aggregation.png`,
  `SACCR_EAD_tree.png` and `SACCR_RC_PFE.png`, used by the Wiki, are rendered
  from them at their own size (1800 x 1550, 1820 x 1568 and 1800 x 1060 px)
  with headless Chromium at a device scale factor of 1, then quantized to 256
  colours as the screenshots above. Render them again whenever an SVG
  changes, and update the pictures when a formula or a supervisory parameter
  they show changes.

<a id="website-banner"></a>

## 🏙️ Website banner

`SACCR_banner.webp` is the background of the banner under the navigation of
the website (`docs/index.html`). It is a decorative photograph provided for the
project, cropped to 1536 x 512 px from the middle of the original so that none of
its own lettering remains, and saved as WebP at quality 80. The banner's text
(SA-CCR for derivatives, Basel CRE52, EU CRR) is HTML, not part of the image.
It is a promotional image: it shows no workbook, data or result.

<a id="social-preview"></a>

## 🪧 Social preview

`SACCR_social_preview.html` is the source for `SACCR_social_preview.png`, the
repository's social preview image (1280 x 640 px, under GitHub's 1 MB limit).
It is a promotional image.

- **Figures:** the netting-set card shows CRE99 Example 5 in USD thousands:
  MPOR 14 business days (CRE99.83), aggregate add-on 1,401 (CRE99.95),
  multiplier 0.958 (CRE99.96) and EAD 1,879 (CRE99.97). PFE is
  0.958 x 1,401 = 1,342 and RC is 0, since 1.4 x (0 + 1,342) = 1,879. The strip
  shows the published EAD of Examples 1 to 5 (CRE99.40, .58, .76, .79, .97),
  the `published` values in `tests/expected/cre99-example-*.bcbs.json`.
- **Curve:** the PFE multiplier m = min(1, F + (1 - F) exp((V - C) /
  (2 (1 - F) AddOn))) with F = 5% (CRE52.23), plotted for (V - C) / AddOn from
  -4 to +1.
- **Name:** the title is the project name, K-SACCR; the method, SA-CCR, is
  named in the label beside the status.
- **Status:** the image names the current release, "Release 1.0.2"; update
  it with each release.

Regenerate by opening the HTML in a Chromium browser with a 1280 x 640 viewport
at device scale factor 1, after the Google Fonts (Barlow Condensed, Barlow,
JetBrains Mono) have loaded, and capturing the full page. Then upload the PNG
in **Settings > General > Social preview**; GitHub has no API for this setting.
