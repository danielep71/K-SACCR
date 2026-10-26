# 🧩 Production Source

`src/` holds the production VBA source and the macro-free workbook template:
the only inputs from which the K-SACCR workbook is built. Layout, visibility and
dependency rules, including what each subdirectory holds, are defined in
[`docs/REPOSITORY_STRUCTURE.md`](../docs/REPOSITORY_STRUCTURE.md).

Current source:

- `core/CORE_Config.bas`: sheet names, the workbook names of the tables,
  column numbers, parameter codes and error numbers;
- `core/CORE_CsvReader.bas`: the reading of a Trades or NettingSets CSV file
  under the [CSV input contract](../docs/csv/CSV_CONTRACT.md); it writes no
  sheet;
- `core/CORE_CsvSchema.bas`: the CSV input contract for the reader, generated
  from `docs/csv/contract.json` by `tools/generate_csv_vba.py`;
- `core/CORE_Engine.bas`: the calculation of EAD per netting set, which reads
  the input sheets and writes the output sheets;
- `core/CORE_Export.bas`: the CSV encoding of the evidence report; it reads no
  sheet;
- `core/CORE_Reconcile.bas`: the comparison of a bank's figures with Results
  and TradeCalc; it reads no sheet;
- `core/CORE_Util.bas`: value conversions, keyed lookups and sheet helpers;
- `modules/SACCR_Formulas.bas`: the public `SACCR_*` worksheet functions,
  listed in [`docs/PUBLIC_API.txt`](../docs/PUBLIC_API.txt);
- `workbook/M_Main.bas`: the sheet-button macros, the CSV check macros, the
  automation entry points
  and status functions, and the reading and writing of the reconciliation
  sheets and the evidence file;
- `workbook/ThisWorkbook.cls` and the `sh*.cls` sheet modules: the document
  modules, which hold only `Option Explicit`;
- `workbook/SACCR_Template.xlsx`: the 18 sheets (README, Params, NettingSets,
  Trades, Import_NettingSets, Import_Trades, Results, HedgingSets, Buckets,
  TradeCalc, Checks, BankNettingSets, Reconciliation, BankTrades,
  TradeReconciliation, TestCatalogue, FormulaTest, Basel_vs_CRR) with their
  formulas, named ranges, data validation and buttons, and no VBA project or
  document properties.

The engine does not yet follow every rule of the layout: `CORE_Engine` and
`CORE_Util` read and write worksheets from `src/core`, and `M_Main` is a
standard module kept with the workbook glue because the sheet buttons call it.
See [known deviations](../docs/REPOSITORY_STRUCTURE.md#known-deviations).

Do not place tests, examples, built workbooks or generated output here.
