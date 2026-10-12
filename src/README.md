# 🧩 Production Source

`src/` holds the production VBA source and the macro-free workbook template:
the only inputs from which the SACCR workbook is built. Layout, visibility and
dependency rules are defined in
[`docs/REPOSITORY_STRUCTURE.md`](../docs/REPOSITORY_STRUCTURE.md).

| Location | Contents |
| --- | --- |
| `core/` | Calculation engine. Every module declares `Option Private Module`; no Excel object-model access. |
| `modules/` | Public facade: supported entry points listed in [`docs/PUBLIC_API.txt`](../docs/PUBLIC_API.txt). |
| `classes/` | Class modules; each header states whether it is public surface or internal (none yet). |
| `workbook/` | Exported `ThisWorkbook` and sheet modules; host glue only. Also `SACCR_Template.xlsx`, the macro-free template holding the 12 sheets. |
| `forms/` | UserForms with `.frm` beside `.frx`, only if one is ever needed (none yet). |

Current source:

- the engine, first imported from the prototype `SACCR_Calculator.xlsm` (#57):
  `core/CORE_Config`, `core/CORE_Util`, `core/CORE_Engine`, the worksheet
  functions in `modules/SACCR_Formulas`, and `workbook/M_Main` with the 13
  document modules;
- `workbook/SACCR_Template.xlsx`: the 12 sheets (README, Params, NettingSets,
  Trades, Results, HedgingSets, Buckets, TradeCalc, Checks, TestCatalogue,
  FormulaTest, Basel_vs_CRR) with their formulas, named ranges, data
  validation and buttons, and no VBA project or document properties.

The imported engine does not yet follow every rule here: `CORE_Engine` and
`CORE_Util` read and write worksheets from `src/core`, and `M_Main` is a standard
module kept with the workbook glue because the sheet buttons call it. See
[known deviations](../docs/REPOSITORY_STRUCTURE.md#known-deviations).

Do not place tests, examples, built workbooks or generated output here.
