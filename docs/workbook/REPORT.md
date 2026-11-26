# Run report

The **Export report** command (#259) writes a formatted report of the last run,
as a PDF or as an XLSX workbook, for a reviewer who reads results rather than
parses them. It opens with the figures of the [Dashboard](DASHBOARD.md) and then
gives the evidence behind them. The CSV that **Export evidence** writes stays
the machine-readable record of the run, with its limitations; the report is a
readable view of the same run. It has no limitations sheet: the owner found it
of no use in the PDF (2026-11-26).

This document is the contract of the report. The machine-readable form is
[`report.json`](report.json): the sections, their columns and kinds, the formats
and the rules for what the PDF shows. [`tools/report.py`](../../tools/report.py)
is the reference of the rules that are computations (the text of a number and a
date, Checks by code, the rows of a reconciliation that the PDF keeps), and its
tests (`tools/test_report.py`) pin them with the examples of `report.json`.
`tools/test_report_spec.py` checks the specification against the template and
the code: every table has the headers of its sheet, the evidence export reads
the same number of columns, and the run lines are lines of the evidence CSV.

## 🧭 Principles

- **The run of the CSV, never another.** The command refuses in the same cases
  as Export evidence, with the same reasons: recalculation pending, a staging
  validation on Checks, results that are not those of the current inputs, an
  output cell edited after the run, a reconciliation of bank figures that is not
  current, and outputs that the run did not record. A report never describes
  results that do not match the inputs.
- **No new calculation.** Every figure is read from the sheets as they are, as
  the evidence CSV reads them. The only figures the report computes are the
  number of rows of each code of Checks and the number of rows the PDF leaves
  out of Checks and of the reconciliations; it also orders the trade legs that
  differ by the size of their difference, which it reads.
- **The workbook is left as it is.** The report is built in a new, temporary
  workbook. The sheets, the names, the cells, the selection, the calculation
  mode and the clipboard of the K-SACCR workbook are not touched, and the
  fingerprints of the inputs and of the outputs are the same after the export,
  so the results stay current.
- **The look of the workbook.** The report has the font, the vertical alignment,
  the header fill (white bold text on `1F3864`) and no gridlines, like the K-SACCR
  sheets of [`sheet-style.json`](sheet-style.json); Run has the tab color of the
  reference sheets, the other sections that of the output sheets.
  `tools/test_report_spec.py` holds the writer's colors to that file.
- **Values only.** The XLSX holds values and formats, with no formula, no macro,
  no name that refers to another workbook and no link to the source workbook.
- **Units.** Amounts of Results, the totals and the Checks are in the reporting
  currency, unscaled, and each table says so. The Dashboard's amounts are on
  the scale of `Dashboard_Scale`, named in the title of each table. The bank
  figures of a reconciliation are in units of the reporting currency, as the
  header says; ratios and times are not scaled.

## 📋 Sections

One sheet of the report per section, in this order. The same workbook is saved
as the XLSX and exported as the PDF, with the differences of the last two
columns.

| Section | Sheet | Content | XLSX | PDF |
| --- | --- | --- | --- | --- |
| Run | `Run` | The identity of the run: the run summary of Results C4, the status of the results, the reporting date and currency, the default regime, the regulatory parameter set, the days of the year, the calendar and the table of closing days, the source of the FX rates, the fingerprints of the inputs and of the outputs, the state of each reconciliation (`CURRENT`, `NOT MADE` or `UNAVAILABLE`, as the evidence CSV says) and the units | | |
| Dashboard | `Dashboard` | Every block of the Dashboard's data table, as the sheet shows it: its title with the unit, its column headers, its row labels and its values (`dashboard.json`) | values, each in the number format of its cell | text in the format of its cell |
| Charts | `Charts` | The ten charts of the Dashboard sheet, as pictures of the charts on the sheet | none | on a sheet of their own |
| Totals | `Totals` | The two TOTAL rows of Results: for each regime the value, the collateral, RC, the add-ons, PFE, EAD and the status line | | |
| Results | `Results` | Every netting set of Results, with its 28 columns, and the TOTAL rows | | |
| Checks | `Checks` | A table with a row per code, its highest severity and its number of rows; then the rows of Checks, grouped by code | every row | at most 10 rows of each code |
| Reconciliation | `Reconciliation` | The summary of the reconciliation of netting sets and its rows, when the sheet holds a comparison | every row | every row that is not `MATCH` |
| Trade reconciliation | `TradeReconciliation` | The summary of the reconciliation of trade legs and its rows, when the sheet holds a comparison | every row | the 50 rows with the result `DIFFERENCE` and the largest absolute difference |

A reconciliation that was not made, or whose sheets the workbook does not have,
has no sheet in the report; the Run sheet says `NOT MADE` or `UNAVAILABLE` as the
evidence CSV does. In the PDF, a line under the summary of each reconciliation
says what is left out, so the omission is never silent:

- **Netting sets:** the rows that are not `MATCH` are shown, and the line counts the
  `MATCH` rows left out ("3 rows with the result MATCH are left out; the workbook
  version of the report has every row.").
- **Trade legs:** a real portfolio gave a PDF of 163 pages, most of it trade legs
  not in the bank's file, so the owner decided on 2026-11-26 that the PDF shows
  only the rows with the result `DIFFERENCE`, the largest absolute difference
  first (equal ones in their order on the sheet), at most 50
  (`pdf_trade_limit`). The summary above the table counts every row by result,
  and the line says what is shown and how many rows are left out ("The 50 rows
  with the result DIFFERENCE and the largest absolute difference are shown, of
  80; the other 5030 rows are left out, counted by result above; the workbook
  version of the report has every row.").

### Checks by code

The first table of the Checks sheet lists every code that Checks shows, with the
highest severity among its rows and its number of rows. A row without a code
belongs to the code `(none)`. The codes are ordered by that severity, ERROR then
WARNING then INFO, and then by code in ordinal order. The table below lists
the rows of Checks in the order of the summary, and the rows of a code in their
order on the sheet. A clean run shows its single `RUN-CLEAN` row. With a real
portfolio Checks can hold a row per trade, so the PDF shows at most 10 rows of each
code (`pdf_rows_per_code`), with a line under the title that says how many rows
are left out ("At most 10 rows of each code are shown; 412 rows are left out; the
workbook version of the report has every row."); the summary still counts every
row, and the XLSX has them all.

## 🔢 Formats

Dates are `yyyy-mm-dd`, with the time `yyyy-mm-dd hh:mm:ss`. The kinds of the
columns of `report.json` decide the number format of the XLSX, which writes the
numbers as numbers and the dates as dates:

| Kind | XLSX format | Used for |
| --- | --- | --- |
| `amount` | `#,##0.00` | Every amount |
| `ratio` | `0.000000` | Multipliers, alpha and relative differences |
| `count` | `0` | Counts, rows and business days |
| `date`, `datetime` | `yyyy-mm-dd`, `yyyy-mm-dd hh:mm:ss` | Dates |
| `text` | text | Everything else |

In the two reconciliations the format of the columns Bank figure, Calculator figure and
Absolute difference follows the measure of the row: `ratio` for the measures that
`CORE_Reconcile` does not scale and compares by relative tolerance only (the multiplier of a netting
set; the maturity, the supervisory duration, the delta and the two maturity factors of a trade
leg), `amount` for every other measure. A multiplier of 0.958585 against 0.958586 so reads as it is,
and the difference of a `DIFFERENCE` row is never shown as `0.00`. `report.json` lists these measures
as `measure_ratios`, and `tools/test_report_spec.py` checks them against `CORE_Reconcile`.

Writing the file does not depend on the regional settings of the machine: the
formats are the invariant codes of Excel, and the values are written as
`Double` and `Date`, never converted to text with the machine's separators. A
workbook opened later shows the numbers with the separators of its reader.

A PDF is read as it was written, so its numbers are text that does not depend
on the machine at all (`CORE_Export.ReportCellText`). An amount is the value to 15 significant digits rounded
half away from zero to two decimals, with a comma between thousands and a point
before the decimals, and a value that rounds to zero has no sign; the other
kinds follow `formats.text_rules` of `report.json`. The examples of that file
(`1234567.891` is `1,234,567.89`, `-2.675` is `-2.68`, `-0.004` is `0.00`,
serial `46295` is `2026-09-30`) are checked by `tools/test_report.py`, and
`TEST_Report` checks the same ones; `tools/test_report_spec.py` holds its list to
`report.json`. The Dashboard's figures keep the number format of their cell on the
sheet: the decimals and the grouping of the format, a percentage times 100 with its
sign, a date as `yyyy-mm-dd`, the same invariant separators
(`CORE_Export.ReportFormatText`, `tools/report.py` `format_text`).

## 🖨️ The PDF

A4 landscape, one page wide: a sheet wider than the page is scaled down, never
cut or continued on another page to the right. The header row of each table
repeats on every page, and the footer reads "K-SACCR report - run *time* - inputs
*fingerprint* - page *n* of *N*". The Dashboard tables come first and then the
pictures of the ten charts, exported from the Dashboard sheet one by one, as the
sheet shows them. Excel draws a chart only where it shows it: exported from a
sheet it did not show, with screen updating off, a chart came out empty, its
frame and title alone (the owner's verification of 2026-11-26). The export
therefore shows the Dashboard, with screen updating on, and scrolls each chart
into view before it is pictured, with the window zoomed to 300% so that each
picture has three times the screen's resolution; the Dashboard's zoom, scroll
position, selection and screen updating are then put back, and the report comes
to the front again. The pictures are laid out in rows, two wide charts or three
narrow ones, and a row that would cross a page starts a new one: the Charts sheet
is printed at the zoom at which its widest row fits the page, since Excel ignores
manual page breaks on a sheet printed to fit a number of pages. Every sheet of
the PDF has the same margins, 0.7 inch at the sides and 0.75 inch at the top and
bottom.
The XLSX holds the same figures as values and has no chart.

## 🔄 Life cycle

`M_Main` checks the run, asks for the file and calls the writer. The Save As
dialog has two file types, PDF and XLSX, and the extension the user types or
chooses decides the format; `ExportReport_Silent` takes the path and decides by
the extension; it returns `RESULT=OK; format=PDF|XLSX; sheets=N; charts=N;
page_setup=OK; file=...` or `RESULT=REFUSED; reason=...`. A printer driver that
refuses a page setting does not stop the PDF: the result says
`page_setup=FAILED (...)`. The command is **Export report** in the Evidence group
of the K-SACCR Ribbon tab, enabled like Export evidence while Results is current
([`RIBBON.md`](RIBBON.md)); it also runs from the Macros dialog as `ExportReport`. A refusal and a write failure are reported in a message, or in
the Immediate window in silent mode, and never raise. A failure leaves no
partial file: the report is written beside the target and then renamed, as the
evidence CSV is, and the temporary workbook is closed without saving whatever
happens. On success and on failure Excel's state is put back: the active
workbook and sheet, screen updating, alerts and calculation mode.

## ✅ Verification

- `tools/test_report.py`: the text rules, Checks by code and the rows the PDF
  keeps.
- `tools/test_report_spec.py`: the specification against the template and the
  evidence export.
- `TEST_Report` (5 cases, 54 checks): the text of the examples, Checks by code,
  the refusals, an XLSX read back from the file (sheets, values only, no link and
  no macro, the run's figures), a PDF with the charts, the results still current,
  and the failures `report-save`, `report-replace` and a read-only file, each
  leaving an older report as it was.
- The owner checks the PDF in Windows Excel 64-bit, which no automatic test
  can see.
