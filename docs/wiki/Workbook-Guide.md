# Workbook Guide

> **Guide, not policy:** [Installation](../../INSTALLATION.md) owns the supported local setup, import and execution procedure.

SA-CCR Benchmark follows a **source-first** model:

- exported VBA under `src/` is the reviewable source of truth;
- the local macro-enabled workbook is a development/execution artifact;
- fixtures and expected values live under `tests/`; and
- generated or local execution artifacts are not authoritative source.

For normal Windows development, `tools/Sync-SACCR-VBA.cmd` updates the VBA
project in the local workbook from repository source while preserving workbook
sheet objects and their contents. It retains a safety backup only when
synchronization fails.

Every sheet has the same layout: two narrow margin columns (A:B, width 4) and
a title band on rows 1 and 2 (B1 and B2). On the output tables, a description
or the run summary is in C4, the Results status in C5, the column headers on
row 6 and the data from row 7.
The input sheets are located through workbook names, so their rows can move:
NettingSets and Trades through `NettingSets_Header` and `Trades_Header`, the
first header cell of each table, which must stay in column C, with data from
the next row (today NettingSets headers on row 6, Trades headers on row 4).
On Trades each row is a leg with a unique Trade ID; rows sharing an Economic
trade ID (column Z) form one economic trade, and exactly one of them has MtM
leg = Y (column AA) and carries its MtM (#38). Comment is column AB.
Headers are compared ignoring case, spaces and line breaks, so rewording only
their spacing does not stop a run. Params is read
through the workbook names `Params_Code` and `Params_Value` (the "Code" and
"Value" header cells) and `SF_Table` and `FX_Table` (from their header
cells), and each parameter through the workbook name equal to its code, so
its rows and columns can move; the names move with the cells, and a missing
name stops the run.
Every font is Aptos, named on each style rather than taken from the workbook
theme, so a sheet copied into another workbook keeps it. The VBA reads and writes the tables through `CORE_Util.TableCell`
and `TableRange`, which add the margin columns, so a column number in
`CORE_Config` counts from column C.

The Tests sheet checks each `SACCR_*` worksheet function against the rule
it implements. A row holds the function's inputs in yellow cells (E:K),
Expected computes the rule with Excel's own functions from the same inputs,
and Function calls the VBA; the row passes when they agree within 1e-9
relative. Changing an input tests the function on new values; the test name
describes the inputs shipped with the template, and the Rule column cites
the source.

A source change should pass the repository static checks and, when it affects
executable VBA or numerical behaviour, be exercised in desktop Excel. See
[Excel Evidence](../EXCEL_EVIDENCE.md) and
[Repository Structure](../REPOSITORY_STRUCTURE.md).
