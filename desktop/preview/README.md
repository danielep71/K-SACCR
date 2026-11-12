# Desktop preview

A mock-up of the desktop application planned for v1.7.0 (#694 to #700), kept
as design input for the architecture decision of #694. It is not part of a
release, has no tests and is not maintained with the workbook.

It shows the planned sections (Dashboard, Counterparties, Trades, Netting
Sets, Calculation, Results, Reports and Settings) on demo data: the five
CRE99 worked examples of `tests/fixtures`, with the figures CRE99 prints from
`tests/expected`. The counterparties are invented. Nothing is calculated, no
workbook is opened and no file is written; the engine choice, the exports
and the editing buttons are placeholders.

From the repository root, with Python 3.10 or later:

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install PySide6
.\.venv\Scripts\python.exe desktop\preview\ksaccr_desktop_preview.py
```

`--screenshots FOLDER` saves one PNG per section and exits.
