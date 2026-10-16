# 🖼️ Documentation images

The images in this directory illustrate K-SACCR. They do not certify the
engine or replace the [methodology](../methodology/README.md), which holds the
regulatory sources, the support matrix and the traceability of each rule. Each
section below says what its images show, how they were made and when to update
them.

<a id="screenshots"></a>

## 📸 Screenshots

The images below show sheets of the workbook built from
`src/workbook/SACCR_Template.xlsx`, with the demonstration portfolio. The owner
took them in Excel for Windows on 2026-10-16.

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
  Regulation (EU) 2021/931. Its MPOR line also relies on CRE54.8, CRE54.12 and
  CRE54.14, registered as `CRE54` in the
  [source register](../methodology/README.md#source-register).

The SVG is the source. Its values were checked against the methodology
parameter table on the release branch on 2026-10-15, when the SVG last
changed; update the card when a supervisory parameter or formula changes.

The workbook carries the card as a picture on its README sheet, under the
"Reference card" heading, so a user never has to leave Excel (#281).
README!D23 links to C114, a cell behind the lower part of the picture:
Excel scrolls only until the target cell is in view, so a link to the
heading would leave the picture below the window. The picture is a PNG of this SVG: when the SVG changes,
render it again and replace the picture in the template
(`xl/media/image1.png`, shown 1420 px wide, the width of columns C:D):

1. Open the SVG in an HTML page that shows it at 1800 x 1444 px and take a
   screenshot with headless Chromium at a device scale factor of 1, in a
   window taller than 1444 px; crop it to 1800 x 1444.
2. Quantize it to 256 colours (Pillow `quantize(colors=256,
   method=MEDIANCUT, dither=NONE)`) and save it with `optimize=True`, as for
   the screenshots above.

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
- **Status:** the image carries "Under development" while no release exists;
  check it at the first release.

Regenerate by opening the HTML in a Chromium browser with a 1280 x 640 viewport
at device scale factor 1, after the Google Fonts (Barlow Condensed, Barlow,
JetBrains Mono) have loaded, and capturing the full page. Then upload the PNG
in **Settings > General > Social preview**; GitHub has no API for this setting.
