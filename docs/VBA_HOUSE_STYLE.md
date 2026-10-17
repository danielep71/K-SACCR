# ✍️ VBA House Style

[![Scope: VBA exports](https://img.shields.io/badge/scope-VBA%20exports-217346)](REPOSITORY_STRUCTURE.md)
[![Contracts: explicit](https://img.shields.io/badge/contracts-explicit-6f42c1)](#procedure-contracts)
[![State: restored](https://img.shields.io/badge/Excel%20state-restored-0969da)](#operation-and-cleanup)

This document is authoritative for **how K-SACCR's VBA is written**: naming,
contracts, errors, Excel-state handling and source presentation. Where a
component lives is owned by
[`REPOSITORY_STRUCTURE.md`](REPOSITORY_STRUCTURE.md); importing and exporting
by [`INSTALLATION.md`](../INSTALLATION.md#importing-vba-into-excel). It adapts
`docs/VBA_HOUSE_STYLE.md` from EXCEL-VBA-PROJECT-TEMPLATE.

## 🏷️ Naming

The VBE lists all components in one flat project, so component names carry
their role in an upper-case prefix. `tools/check_source.py` enforces the
prefixes for core, facade and test modules.

| Element | Convention | Example |
| --- | --- | --- |
| Core module (`src/core/`) | `CORE_` + subject | `CORE_Engine` |
| Public API module (`src/modules/`) | `SACCR_` + subject | `SACCR_Formulas` |
| Class (`src/classes/`) | `C` + noun | `CNettingSet` |
| Workbook module (`src/workbook/`) | `ThisWorkbook`; sheets by code name `sh` + noun | `shParams` |
| Test module (`tests/`) | `TEST_` + subject | `TEST_Harness` |
| Example module (`examples/`) | `Example` + subject | `ExampleSingleNettingSet` |
| Worksheet function (`src/modules/`) | `SACCR_` + quantity; arguments use the regulatory symbols | `SACCR_ReplacementCost(V, C, ...)` |
| Procedure | PascalCase verb phrase | `CalculateAddOn` |
| Parameter, local variable | PascalCase noun, spelled out | `NotionalAmount` |
| Loop counter | one lower-case letter | `i`, `j`, `r`, `c` |
| Module-level variable | `m` + PascalCase | `mRunClass` |
| Public variable | `g` + PascalCase | `gEngineFault` |
| Constant | UPPER_SNAKE_CASE | `ALPHA_FACTOR` |
| Enum type / member | `Saccr` + noun / short prefix + PascalCase | `SaccrAssetClass` / `acInterestRate` |

The project has no class or example modules yet: `src/classes/` and
`examples/` do not exist. Their rows apply when such modules are introduced.

Names are English and spelled out; avoid abbreviations other than established
SA-CCR terms (`EAD`, `RC`, `PFE`, `MPOR`), written in full in the procedure's
comment the first time.

A single letter is reserved for a loop counter, the variable a `For` loop, or
a `Do` loop that steps through positions, counts with; it stays lower case.
Every other variable and parameter names what it holds, also in a short
helper (`Formatted`, not `s`; `CodePoint`, not `cp`). Worksheet functions
are the exception the table gives: their arguments are the regulatory
symbols. VBA names are case-insensitive and the VBE shows one spelling per
name across the project, so the source files, not the editor, carry the
convention.

## ⚙️ Module rules

- Every module declares `Option Explicit` (enforced by `tools/check_source.py`).
- Every module in `src/core/` and `tests/modules/` declares
  `Option Private Module` (enforced for `src/core/` by `tools/check_source.py`).
- No `Option Base 1`: arrays are declared with explicit bounds
  (`Dim Values(1 To ValueCount)`).
- No `Variant` where a specific type works. Arguments are `ByVal` unless the
  procedure's contract is to write an output parameter.
- No global mutable state in `src/core/`. Core procedures depend only on their
  arguments, except the engine imported from the prototype; see the
  [known deviations](REPOSITORY_STRUCTURE.md#known-deviations).
- A `Declare` statement, if ever needed, uses `PtrSafe` and `LongPtr`, because
  64-bit Excel is the primary target.

<a id="procedure-contracts"></a>

## 📜 Procedure contracts

Every procedure documents its contract in a banner immediately after its
signature. Omit sections that are empty.

| Section | Explain |
| --- | --- |
| `PURPOSE` | The responsibility the procedure owns |
| `INPUTS` | Meaning, **unit**, **valid domain** and ownership of each argument |
| `RETURNS` | Meaning and unit of the result and any output parameter |
| `ERROR POLICY` | Errors raised, propagated or contained, and cleanup |
| `DEPENDENCIES` | Boundaries that matter to callers or maintainers |
| `STATE OWNERSHIP` | Excel or module state read, changed or restored |
| `REFERENCE` | Regulatory paragraph or published example implemented |
| `UPDATED` | Date the procedure was last revised |

### Units and domains

- Amounts are `Double` in the netting set's calculation currency, never scaled
  (no thousands or millions).
- Rates, factors and correlations are decimals: `0.05`, never `5` meaning 5%.
- Time periods are `Double` in years unless the name says otherwise
  (`MporBusinessDays`). Dates are VBA `Date`.
- Every numeric input states its valid domain, for example `NotionalAmount >= 0`
  or `0 < correlation <= 1`. For an input with no meaningful result, the
  facade's worksheet functions return a worksheet error value (`CVErr`) or
  apply a documented floor.
- The exact day-count, maturity and supervisory parameter conventions are part
  of the [methodology](methodology/README.md) and are cited in `REFERENCE` by
  source ID and locator, for example `CRR Article 275(1)`.

### Errors

- Error numbers are named constants in the range `vbObjectError + 2048` to
  `vbObjectError + 4095`, defined once in a core module and re-exposed by the
  facade when callers need them.
- Raise with `Err.Raise number, source, description`. The source names the
  component and procedure (`"CORE_Engine.Calculate"`); the description states
  what was invalid and the offending value.
- Macros and the engine raise; the facade's worksheet functions do not (see
  [Units and domains](#units-and-domains)).
- `On Error Resume Next` is allowed only around a single statement whose error
  is checked on the next line, followed by `On Error GoTo 0` or the procedure's
  handler.
- Errors are never swallowed silently. A procedure either handles an error
  completely or propagates it unchanged.

### Public API changes

The supported API is the list in [`PUBLIC_API.txt`](PUBLIC_API.txt). Adding,
removing or changing a listed declaration's name, arguments, defaults, units,
return value or errors is a contract change. It updates the manifest, the
regression tests and the `[Unreleased]` changelog entry in the same pull
request, and states the compatibility impact.

<a id="operation-and-cleanup"></a>

## 🧹 Operation and cleanup

Any procedure that changes Excel application state owns restoring it.

1. **Capture** the values it will change before changing them
   (`Calculation`, `ScreenUpdating`, `EnableEvents`, `DisplayAlerts`, `Cursor`,
   `StatusBar`).
2. **Change** only what the operation needs.
3. **Restore** in one cleanup block reached on both the success and the error
   path, restoring only what this procedure changed.
4. **Preserve the primary error.** Copy `Err.Number`, `Err.Source` and
   `Err.Description` into locals before cleanup runs, and re-raise those after
   cleanup. Restoring is best effort: a failure inside cleanup must never
   replace the original error.

The reference implementation is `M_Main.ExecuteOperation` in
[`src/workbook/M_Main.bas`](../src/workbook/M_Main.bas), with
`RestoreSettings`, `RestoreCalculation`, `RestoreEvents` and
`RestoreScreenUpdating`. Each `Restore...` helper contains its own error and
returns what failed as text, empty on success, so every restoration is
attempted. The sketch below follows the same pattern; a cleanup-only failure
raises `ERR_CLEANUP_FAILED` (`src/core/CORE_Config.bas`), never success, and
`CleanupDetails` returns the failures to the caller also when the primary
error is re-raised.

```vb
'In a standard module named M_Example
Public Sub RunOperation(ByRef CleanupDetails As String)
    Dim SavedCalculation   As XlCalculation    'Calculation mode before the operation
    Dim SavedScreen        As Boolean          'ScreenUpdating before the operation
    Dim StateCaptured      As Boolean          'True once both were read
    Dim ErrNumber          As Long             'Primary error number; 0 for none
    Dim ErrSource          As String           'Primary error source
    Dim ErrDescription     As String           'Primary error description

        CleanupDetails = vbNullString
        On Error GoTo Failed
        SavedCalculation = Application.Calculation
        SavedScreen = Application.ScreenUpdating
        StateCaptured = True
        Application.ScreenUpdating = False
        Application.Calculation = xlCalculationManual

        '... work ...

CleanUp:
        On Error GoTo 0
        If StateCaptured Then
            CleanupDetails = AppendDetail(RestoreCalculation(SavedCalculation), _
                                          RestoreScreenUpdating(SavedScreen))
        End If
        If ErrNumber <> 0 Then Err.Raise ErrNumber, ErrSource, ErrDescription
        If Len(CleanupDetails) > 0 Then
            Err.Raise ERR_CLEANUP_FAILED, "M_Example.RunOperation", CleanupDetails
        End If
        Exit Sub

Failed:
        ErrNumber = Err.Number
        ErrSource = Err.Source
        ErrDescription = Err.Description
        Resume CleanUp
End Sub
```

Core procedures do not change Excel application settings, so they need none
of this. They are not free of Excel side effects: `CORE_Engine` and
`CORE_Util` read the input sheets and write the output sheets, one of the
[known deviations](REPOSITORY_STRUCTURE.md#known-deviations) of the engine
imported from the prototype.

## 🧱 Module layout

1. `Attribute VB_Name` on the first line of a `.bas` export.
2. A full-width module banner: name, purpose, public surface, dependencies,
   state ownership, error policy, compatibility, updated date and author.
3. `MODULE SETTINGS`: `Option Explicit`, plus `Option Private Module` in core.
4. `MODULE CONSTANTS` and, outside core, `MODULE STATE` when needed; the
   engine imported from the prototype is the exception (see the
   [known deviations](REPOSITORY_STRUCTURE.md#known-deviations)).
5. Centered, uppercase procedure-group headings, then the procedures.

Banners use an apostrophe and 78 `=`; section rules an apostrophe and 78 `-`.
Source decorations are ASCII only.

Generated modules (`TEST_Cases`, written by `tools/generate_case_tests.py`) are
exempt from the procedure and module banner rules; their header says they are
generated and must not be edited by hand.

## 📐 Indentation and wrapping

- Four-space levels; no tabs, no trailing whitespace.
- Local declarations at four spaces with `As` aligned in each group; body
  statements at eight spaces, nested blocks adding four.
- Procedure declarations, `End Sub` / `End Function`, labels and `#If`
  directives at column one.
- A signature with parameters puts one parameter per continuation line, and a
  function's return type on its own continuation line:

  ```vb
  Public Function CalculateAddOn( _
      ByVal NotionalAmount As Double, _
      ByVal SupervisoryFactor As Double) _
      As Double
  ```

- Two empty lines between procedures. Aim for 79-column comments and about 100
  columns of code.
- Comments explain intent and invariants, not assignments. Each declaration
  carries a short trailing comment stating its role or unit.

<a id="export-compatibility"></a>

## 💾 Export compatibility

- **Encoding:** the VBE exports in the Windows ANSI code page, which is cp1252
  on the supported Western-European hosts. Keep source ASCII; build any other
  character at run time with `ChrW`. `tools/check_source.py` rejects files that
  do not decode as cp1252. No byte-order mark.
- **Line endings:** stored with LF in Git, checked out with CRLF by
  `.gitattributes`, which is what the VBE expects. Never change these rules or
  the export header when reformatting.
- **Forms:** a `.frm` and its `.frx` are one component, kept side by side. The
  `.frx` is binary and is never edited as text.
- **Workbooks and other Office files** are build outputs, ignored by Git. One
  may be committed only at an exact path re-included in `.gitignore` and
  approved in an issue. The one such path is
  `src/workbook/SACCR_Template.xlsx`, which carries no VBA project or document
  properties.

## 🔎 Review and validation

- A formatting-only change must leave names, signatures, visibility, constants,
  calculations, tests and error behavior unchanged; a discovered defect goes in
  a separate change. Renaming local variables and parameters to the
  [naming](#naming) rules is the one exception for names: it changes no
  logic or string, and in comments only the renamed names (in `INPUTS` and
  wherever a comment names them). Its pull request shows that applying the
  map of old to new names to the previous file reproduces every code line.
  The exception does not cover the parameters of a declaration listed in
  [`PUBLIC_API.txt`](PUBLIC_API.txt): a caller may pass them by name, so
  renaming one is a [contract change](#public-api-changes).
- Comments are kept true. A change updates every module header, procedure
  header, `INPUTS` name and line comment it makes stale, and the `UPDATED`
  date of each procedure and module it touches.
- `python tools/check.py` must pass. Its gates are listed in
  [`tools/README.md`](../tools/README.md); it does not compile VBA.
- Changed VBA is merged only after it has been compiled and tested in Excel, as
  described in [`CONTRIBUTING.md`](../CONTRIBUTING.md#validation-and-evidence).

---

**Style principle:** state the contract, name the unit, restore what you
change, and never lose the original error.
