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
two margin rows (row 1, 24 pt high, 14 pt font; row 2, 26 pt, 16 pt font).
On the input and output tables, the title is in C3, a description or the
run summary in C4, the Results status in C5, the column headers on row 6 and
the data from row 7. README and Params have their own layout. Params is read
through the workbook names `Params_Code` and `Params_Value` (the "Code" and
"Value" header cells) and `SF_Table` and `FX_Table` (from their header
cells), and each parameter through the workbook name equal to its code, so
its rows and columns can move; the names move with the cells, and a missing
name stops the run.
Every font is Aptos, named on each style rather than taken from the workbook
theme, so a sheet copied into another workbook keeps it. The VBA reads and writes the tables through `CORE_Util.TableCell`
and `TableRange`, which add the margin columns, so a column number in
`CORE_Config` counts from column C.

A source change should pass the repository static checks and, when it affects
executable VBA or numerical behaviour, be exercised in desktop Excel. See
[Excel Evidence](../EXCEL_EVIDENCE.md) and
[Repository Structure](../REPOSITORY_STRUCTURE.md).
