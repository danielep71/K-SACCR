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
The title is in C3, a description or the run summary in C4, the Results
status in C5, the column headers on row 6 and the data from row 7. Every font
is Aptos. The VBA reads and writes the tables through `CORE_Util.TableCell`
and `TableRange`, which add the margin columns, so a column number in
`CORE_Config` counts from column C.

A source change should pass the repository static checks and, when it affects
executable VBA or numerical behaviour, be exercised in desktop Excel. See
[Excel Evidence](../EXCEL_EVIDENCE.md) and
[Repository Structure](../REPOSITORY_STRUCTURE.md).
