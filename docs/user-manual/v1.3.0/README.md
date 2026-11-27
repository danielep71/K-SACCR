# K-SACCR user manual, version 1.3.0

The source of the user manual for K-SACCR 1.3.0. The PDF built from it will be
attached to the
[1.3.0 GitHub Release](https://github.com/danielep71/K-SACCR/releases/tag/v1.3.0);
it is not stored in this repository.

| File | Contents |
| --- | --- |
| [`K-SACCR_v1.3.0_User_Manual.md`](K-SACCR_v1.3.0_User_Manual.md) | The manual, in pandoc Markdown |
| `assets/data_flow.png`, `assets/evidence_gate.png`, `assets/rc_pfe.png` | The three diagrams of the earlier manuals, unchanged from 1.2.0; `rc_pfe.png` is a copy of `docs/assets/SACCR_RC_PFE.png` |
| `assets/ribbon.png`, `assets/home.png`, `assets/params.png`, `assets/nettingsets.png`, `assets/import_nettingsets.png`, `assets/dashboard.png`, `assets/dashboard_charts.png`, `assets/results_ead.png`, `assets/checks.png`, `assets/banknettingsets.png`, `assets/reconciliation.png`, `assets/reconsummary.png` | Screenshots of the workbook, copies of `docs/assets/SACCR_Ribbon.png`, `SACCR_Home.png`, `SACCR_Params.png`, `SACCR_NettingSets.png`, `SACCR_Import_NettingSets.png`, `SACCR_Dashboard.png`, `SACCR_Dashboard_Charts.png`, `SACCR_Results_EAD.png`, `SACCR_Checks.png`, `SACCR_BankNettingSets.png`, `SACCR_Reconciliation.png` and `SACCR_ReconSummary.png`, taken on 2026-11-27 ([provenance](../../assets/README.md#screenshots)) |
| `header.tex` | The LaTeX header the PDF build needs (`\Needspace`, the link colour `KBlue`), unchanged from 1.2.0 |

The images are copies, so that the folder builds on its own and keeps the
pictures it was released with when `docs/assets/` is retaken later.

The Markdown is written for pandoc and LaTeX: page-break commands
(`\Needspace`, `\newpage`, `\clearpage`) and display formulas in `\[ … \]`
appear as plain text when the file is viewed on GitHub. Read the PDF.

## Status

Prepared for 1.3.0, not yet released. Started from the 1.2.0 manual and
updated for the changes of 1.3.0:

- the **navigation bar** of every sheet, in a new section: its sections in
  the order of **Go to**, `Home` and the inputs at the top in the frozen rows,
  the other sections scrolling with the rows, the highlighted current sheet,
  the links that work with macros disabled, and the two ways to hide and show
  it, the **Navigation bar** toggle of the K-SACCR tab and the triangle in
  `B1`, for the session only;
- the K-SACCR tab: its icons, **Clear outputs** as a button of its own, the
  bank entries of the Import CSV and Templates menus, **Export report**, the
  **Navigation bar** toggle, the tab selected at opening and the workbook at
  100%, the progress of the long commands in the status bar, and the buttons
  of every sheet (Params, the staging, output and bank sheets, FormulaTest);
- the map of the 21 worksheets, with **Dashboard** and **ReconSummary**, and
  the protection matrix: the reconciliation sheets no longer protected, the
  filters of the output sheets, the new editable cells of Params and
  FormulaTest;
- `Params`: the day count (`YearConvention`, `Calendar`, the table of closing
  days and the TARGET calendar) in a new section, in place of `DaysPerYear`;
  the FX rates quoted the market way, the rebase on a new reporting currency
  and **Download prices** from closing prices or the ECB reference rates;
  each regime's credit scale in the factor table;
- `NettingSets` and `Trades`: the short headers with their notes, the header
  on row 4 and the data from row 5, the counterparty sector, the formats of
  committed rows and the extended drop-down lists;
- the CSV import: contract version 3, NettingSets files of version 2 and the
  reading of version 1 files, the FX notional rule, the limits;
- `Checks`: the Code and Field columns, the stable diagnostic codes and one
  line per violation;
- the **Dashboard**, in a new chapter, and the **formatted run report** of
  **Export report**, in the evidence chapter;
- the reconciliation: the bank files and **Clear bank figures**, the empty
  bank sheets, the default tolerances of 1 unit and 1%, the tiles, the sheets'
  filters and formats, and **ReconSummary**, whose sums, like the tiles',
  count only the rows compared, `MATCH` and `DIFFERENCE`;
- the validation boundary: the independent implementation and the production
  reconciliation, and the rules that no external source covers;
- *Moving inputs from 1.2.0* in place of *Moving inputs from 1.1.0*, the
  automation entry points, the worksheet function `SACCR_BusinessDays`, the
  troubleshooting table, the known limitations, the checklists, the delivered
  files, the glossary and the quick-reference card.

As before, the manual is written for someone who receives only the workbook
`K-SACCR_1.3.0.xlsm` and `SHA256SUMS.txt`: it does not mention the source
repository, building the workbook from source or the release process, and of
the testing it gives only the summary result. It quotes no input fingerprint
and no total of the demonstration portfolio, which change with the template.
The release's PDF is built from this folder at the release tag, and its
SHA-256 is listed in the release's `SHA256SUMS.txt`. The 1.0.0, 1.0.1, 1.0.2,
1.1.0 and 1.2.0 manuals stay as released in
[`../v1.0.0/`](../v1.0.0/README.md), [`../v1.0.1/`](../v1.0.1/README.md),
[`../v1.0.2/`](../v1.0.2/README.md), [`../v1.1.0/`](../v1.1.0/README.md) and
[`../v1.2.0/`](../v1.2.0/README.md).

## Building the PDF

Requires pandoc 3 and a XeLaTeX distribution with the `needspace` package
(for example TeX Live with `texlive-xetex` and `texlive-latex-extra`) and the
DejaVu fonts. From this folder:

```sh
pandoc K-SACCR_v1.3.0_User_Manual.md \
  --from markdown+tex_math_single_backslash \
  --pdf-engine=xelatex --toc --number-sections \
  -H header.tex -V geometry:margin=2cm -V fontsize=10pt \
  -V mainfont="DejaVu Sans" -V monofont="DejaVu Sans Mono" \
  --resource-path=. -o K-SACCR_User_Manual_1.3.0.pdf
```

Edit the Markdown, not the PDF. A later version of the manual goes in its own
folder (`docs/user-manual/v1.3.1/` and so on); once 1.3.0 is released, this
one stays as released.
