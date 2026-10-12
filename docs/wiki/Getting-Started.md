# 🚀 Getting Started

> **Guide, not policy:** [Installation](../../INSTALLATION.md) owns setup, supported hosts, import, recovery, and execution.

## 1. Use the supported host

The supported target is Microsoft 365 or Excel 2016 and later on **Windows 64-bit**. Windows 32-bit is best effort and not certified. Mac, Excel for the web, and Excel 2013 or earlier are not supported. Use the default VBA/Excel/Office references specified in Installation.

## 2. Build from one source commit

There is no published v1.0.0 workbook to download yet.

1. Clone the repository and select `release/1.0.0`.
2. Run `python tools/check.py`.
3. Open `src/workbook/SACCR_Template.xlsx` and save a separate development copy as `.xlsm`.
4. Import the standard modules following Installation. Populate the existing sheet and ThisWorkbook document modules; do not import them as additional classes.
5. Include the current test modules for a development workbook.
6. Run **Debug → Compile VBAProject**, then run the applicable suites and record the exact commit and Excel environment.

Follow your organization's macro policy. A successful import or compile is not numerical validation.

## 3. Run a synthetic portfolio

| Step | Action |
| --- | --- |
| Params | Set valuation date, reporting currency, regime, parameter values, and supplied numeric FX rates |
| NettingSets | Enter netting-set identity, margin terms, collateral, and applicable regime |
| Trades | Enter unique leg IDs; group decompositions with Economic trade ID and designate one MtM leg |
| Validate | Run `ValidateInputs` and inspect Checks |
| Calculate | Run `RunSACCR` |
| Review | Inspect Results status, invalid/withheld populations, parameter class, Checks, and intermediates |

Do not accept a portfolio total while ignoring rejected rows or invalid netting sets. After changing inputs, rerun before relying on the outputs.

## 4. Update deliberately

The synchronizer updates VBA, **not worksheet layouts, names, formulas, or validation**. Rebuild from the accepted template when those change, including the economic-leg columns and FormulaTest migration.

The optional `-SourceDir` folder workflow has open identity and encoding defects ([#159](https://github.com/danielep71/VBA-SACCR-Toolkit/issues/159), [#160](https://github.com/danielep71/VBA-SACCR-Toolkit/issues/160)). Preserve a backup and verify selected modules and imported text. Prefer a complete checkout of one commit for reproducible testing.

See [Workbook Guide](Workbook-Guide.md) for the current layout and [Test Catalogue](Test-Catalogue.md) for the tests.
