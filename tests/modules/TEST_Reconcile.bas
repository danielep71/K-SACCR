Attribute VB_Name = "TEST_Reconcile"
'==============================================================================
' MODULE: TEST_Reconcile
'------------------------------------------------------------------------------
' PURPOSE
'   Check the reconciliation of a bank's figures with the calculator's
'   (#45), per netting set against Results and per trade leg against
'   TradeCalc: figures within either tolerance MATCH and others are a
'   DIFFERENCE; bank amounts are scaled and the multiplier is not; IDs
'   missing on either side, listed twice, of another regime, withheld on
'   Results or unusable are each reported and never compared; TOTAL rows
'   are not netting sets; trade-leg ratios are not scaled, a figure the
'   calculator does not have is NOT APPLICABLE and an excluded leg is
'   EXCLUDED; and, in the workbook, a run writes the Reconciliation and
'   TradeReconciliation sheets, an edited bank figure makes them OUT OF
'   DATE, and validation withdraws them. The evidence report (CORE_Export,
'   M_Main.ExportEvidence_Silent) writes invariant CSV fields and UTF-8,
'   is the same for two exports of one run, and is refused when the
'   results or the bank figures changed after the run.
'
' PUBLIC SURFACE
'   RunReconcileTests is the entry point. Option Private Module keeps it out
'   of the external workbook automation API.
'
' DEPENDENCIES
'   CORE_Reconcile for the comparison cases, which use synthetic arrays and
'   no sheet; M_Main for the workbook case, which needs a workbook built
'   from the template with the full source and its demonstration inputs;
'   TEST_Suite.Say for the log lines. Every figure is synthetic.
'
' WORKSHEET SAFETY
'   The workbook cases write bank figures on BankNettingSets and
'   BankTrades, run the macros, which rewrite every output sheet, and put
'   the bank tables and settings back, also after an unexpected error. The
'   export case writes two files to the TEMP folder and deletes them, and
'   edits one input and one bank cell and puts them back. It leaves the
'   outputs of a normal run. Use a development workbook.
'
' STATE OWNERSHIP
'   Restores the caller's calculation mode, M_Main.gSilent and
'   M_Main.gQuietLog. gQuietLog is set for the run: the macros' run messages
'   would push the report out of the Immediate window, which keeps about
'   200 lines.
'
' ERROR POLICY
'   An unexpected error is recorded against its case and the remaining
'   cases still run.
'
' USAGE
'   Import with the production modules, compile, then run
'   TEST_Reconcile.RunReconcileTests from the Immediate window. It ends
'   with RESULT=PASS; cases=...; checks=...; failures=0.
'
' UPDATED
'   2026-10-17
'
' AUTHOR
'   Daniele Penza
'==============================================================================

'------------------------------------------------------------------------------
' MODULE SETTINGS
'------------------------------------------------------------------------------
    'Require explicit declarations; keep the tests out of the external API.
    Option Explicit
    Option Private Module

'------------------------------------------------------------------------------
' MODULE CONSTANTS
'------------------------------------------------------------------------------
        Private Const EXPECTED_CASES    As Long = 19              'Cases in a complete run
        Private Const EXPECTED_CHECKS   As Long = 144              'Checks in a complete run
        Private Const TOL               As Double = 0.000001      'Tolerance of numeric checks

'------------------------------------------------------------------------------
' MODULE STATE
'------------------------------------------------------------------------------
        Private mCaseCount      As Long      'Cases started
        Private mCheckCount     As Long      'Checks evaluated
        Private mFailureCount   As Long      'Failed checks and unexpected errors
        Private mCurrentCase    As String    'Case being run, for messages
        Private mCounts()       As Long      'Counts of the last comparison


'
'------------------------------------------------------------------------------
'
'                                 ENTRY POINT
'
'------------------------------------------------------------------------------
'

Public Sub RunReconcileTests()
'
'==============================================================================
'                              RunReconcileTests
'------------------------------------------------------------------------------
' PURPOSE
'   Run every case and print one CASE line per case, each failed check,
'   and a final RESULT line.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim callerCalculation   As XlCalculation    'Caller's calculation mode
    Dim callerSilent        As Boolean          'Caller's M_Main.gSilent
    Dim callerQuiet         As Boolean          'Caller's M_Main.gQuietLog

'------------------------------------------------------------------------------
' RUN
'------------------------------------------------------------------------------
        mCaseCount = 0
        mCheckCount = 0
        mFailureCount = 0
        callerCalculation = Application.Calculation
        callerSilent = M_Main.gSilent
        callerQuiet = M_Main.gQuietLog
        M_Main.gQuietLog = True
        TEST_Suite.Say "SACCR RECONCILIATION TESTS"

        CaseTolerances
        CaseScale
        CaseMultiplier
        CaseZeroBankFigure
        CaseMissingAndNotInBank
        CaseDuplicate
        CaseRegimeAndWithheld
        CaseInvalidInput
        CaseIdsAndTotals
        CaseSettingsAndHeaders
        CaseSummary
        CaseWorkbook
        CaseTradeFigures
        CaseTradePopulations
        CaseTradeCalcDuplicates
        CaseTradeSummary
        CaseTradeWorkbook
        CaseExportFields
        CaseExportWorkbook

        Application.Calculation = callerCalculation
        M_Main.gSilent = callerSilent
        M_Main.gQuietLog = callerQuiet
        If mCaseCount <> EXPECTED_CASES Then
            Fail "suite", "expected " & EXPECTED_CASES & " cases, ran " & mCaseCount
        End If
        If mCheckCount <> EXPECTED_CHECKS Then
            Fail "suite", "expected " & EXPECTED_CHECKS & " checks, ran " & mCheckCount
        End If
        TEST_Suite.Say "RESULT=" & IIf(mFailureCount = 0, "PASS", "FAIL") & _
                    "; cases=" & mCaseCount & "; checks=" & mCheckCount & _
                    "; failures=" & mFailureCount

End Sub


'
'------------------------------------------------------------------------------
'
'                              COMPARISON CASES
'
'------------------------------------------------------------------------------
'

Private Sub CaseTolerances()
'
'==============================================================================
' PURPOSE
'   A figure matches within the absolute or within the relative tolerance,
'   and is a DIFFERENCE outside both; difference = calculator - bank and
'   the relative difference is taken on the bank figure.
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim bank   As Variant    'Bank rows
    Dim calc   As Variant    'Results rows
    Dim found  As Variant    'Reconciliation rows

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
        On Error GoTo Unexpected
        BeginCase "compare.tolerances"
        bank = NewBlock(1, BN_NCOLS)
        SetBankRow bank, 1, "NS1", "CRR", 1000, 1000000, 500
        calc = NewBlock(1, RS_NCOLS)
        SetCalcRow calc, 1, "NS1", "CRR", "VALID", 1000.5, 1000050, 600, 0.5
        found = Compare(bank, calc, 1#, 1#, 0.0001)
        CheckText "EAD within the absolute tolerance", ResultOf(found, "NS1", "EAD"), RES_MATCH
        CheckText "RC within the relative tolerance", ResultOf(found, "NS1", "RC"), RES_MATCH
        CheckText "PFE outside both", ResultOf(found, "NS1", "PFE"), RES_DIFF
        CheckNear "PFE difference", CellOf(found, "NS1", "PFE", RC_DIFF), 100#
        CheckNear "PFE relative difference", CellOf(found, "NS1", "PFE", RC_REL), 0.2
        Check mCounts(RCN_COMPARED) = 3 And mCounts(RCN_MATCH) = 2 And mCounts(RCN_DIFF) = 1, _
              "counts compared/match/difference: " & mCounts(RCN_COMPARED) & "/" & mCounts(RCN_MATCH) & "/" & _
              mCounts(RCN_DIFF)
        Exit Sub

Unexpected:
        Fail mCurrentCase, "unexpected error " & Err.Number & ": " & Err.Description

End Sub


Private Sub CaseScale()
'
'==============================================================================
' PURPOSE
'   Bank amounts in thousands are compared in units, and the absolute
'   tolerance is in units too.
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim bank   As Variant    'Bank rows
    Dim calc   As Variant    'Results rows
    Dim found  As Variant    'Reconciliation rows

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
        On Error GoTo Unexpected
        BeginCase "compare.scale"
        bank = NewBlock(1, BN_NCOLS)
        SetBankRow bank, 1, "NS1", "BCBS", 48.6, 12.002
        calc = NewBlock(1, RS_NCOLS)
        SetCalcRow calc, 1, "NS1", "BCBS", "VALID", 48600, 12000, 0, 0
        found = Compare(bank, calc, 1000#, 1#, 0#)
        CheckNear "EAD bank figure in units", CellOf(found, "NS1", "EAD", RC_BANK), 48600#
        CheckText "EAD in thousands", ResultOf(found, "NS1", "EAD"), RES_MATCH
        CheckNear "RC difference in units", CellOf(found, "NS1", "RC", RC_DIFF), -2#
        CheckText "RC two units apart", ResultOf(found, "NS1", "RC"), RES_DIFF
        Exit Sub

Unexpected:
        Fail mCurrentCase, "unexpected error " & Err.Number & ": " & Err.Description

End Sub


Private Sub CaseMultiplier()
'
'==============================================================================
' PURPOSE
'   The multiplier is not scaled and the absolute tolerance does not apply
'   to it: only an equal value or the relative tolerance matches.
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim bank   As Variant    'Bank rows
    Dim calc   As Variant    'Results rows
    Dim found  As Variant    'Reconciliation rows

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
        On Error GoTo Unexpected
        BeginCase "compare.multiplier"
        bank = NewBlock(2, BN_NCOLS)
        SetBankRow bank, 1, "NS1", "CRR"
        bank(1, BN_MULT) = 0.5
        SetBankRow bank, 2, "NS2", "CRR"
        bank(2, BN_MULT) = 0.9
        calc = NewBlock(2, RS_NCOLS)
        SetCalcRow calc, 1, "NS1", "CRR", "VALID", 1, 1, 1, 0.5
        SetCalcRow calc, 2, "NS2", "CRR", "VALID", 1, 1, 1, 0.8
        found = Compare(bank, calc, 1000#, 1#, 0.0001)
        CheckNear "multiplier not scaled", CellOf(found, "NS1", "Multiplier", RC_BANK), 0.5
        CheckText "equal multiplier", ResultOf(found, "NS1", "Multiplier"), RES_MATCH
        CheckText "absolute tolerance not applied", ResultOf(found, "NS2", "Multiplier"), RES_DIFF
        Exit Sub

Unexpected:
        Fail mCurrentCase, "unexpected error " & Err.Number & ": " & Err.Description

End Sub


Private Sub CaseZeroBankFigure()
'
'==============================================================================
' PURPOSE
'   A bank figure of zero has no relative difference: only the absolute
'   tolerance decides.
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim bank   As Variant    'Bank rows
    Dim calc   As Variant    'Results rows
    Dim found  As Variant    'Reconciliation rows

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
        On Error GoTo Unexpected
        BeginCase "compare.zero-bank-figure"
        bank = NewBlock(1, BN_NCOLS)
        SetBankRow bank, 1, "NS1", "CRR", Empty, 0, 0
        calc = NewBlock(1, RS_NCOLS)
        SetCalcRow calc, 1, "NS1", "CRR", "VALID", 10, 0.5, 5, 0.5
        found = Compare(bank, calc, 1#, 1#, 0.5)
        CheckText "zero within the absolute tolerance", ResultOf(found, "NS1", "RC"), RES_MATCH
        CheckText "zero outside it", ResultOf(found, "NS1", "PFE"), RES_DIFF
        Check IsEmpty(CellOf(found, "NS1", "PFE", RC_REL)), "relative difference of a zero bank figure is not empty"
        Check RowOf(found, "NS1", "EAD") = 0, "a blank bank figure is compared"
        Exit Sub

Unexpected:
        Fail mCurrentCase, "unexpected error " & Err.Number & ": " & Err.Description

End Sub


Private Sub CaseMissingAndNotInBank()
'
'==============================================================================
' PURPOSE
'   A bank ID absent from Results is MISSING IN CALCULATOR, with its bank
'   EAD; a Results netting set absent from the bank file is NOT IN BANK
'   FILE, with its EAD and status.
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim bank   As Variant    'Bank rows
    Dim calc   As Variant    'Results rows
    Dim found  As Variant    'Reconciliation rows

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
        On Error GoTo Unexpected
        BeginCase "compare.missing-and-not-in-bank"
        bank = NewBlock(2, BN_NCOLS)
        SetBankRow bank, 1, "NS1", "CRR", 100
        SetBankRow bank, 2, "NS9", "CRR", 900
        calc = NewBlock(2, RS_NCOLS)
        SetCalcRow calc, 1, "NS1", "CRR", "VALID", 100, 0, 0, 0
        SetCalcRow calc, 2, "NS2", "BCBS", "VALID", 250, 0, 0, 0
        found = Compare(bank, calc, 1#, 0#, 0#)
        CheckText "bank ID not on Results", ResultOf(found, "NS9", "EAD"), RES_MISSING
        CheckNear "its bank EAD", CellOf(found, "NS9", "EAD", RC_BANK), 900#
        CheckText "Results netting set not in the bank file", ResultOf(found, "NS2", "EAD"), RES_NOTINBANK
        CheckNear "its Results EAD", CellOf(found, "NS2", "EAD", RC_CALC), 250#
        Check IsEmpty(CellOf(found, "NS2", "EAD", RC_BANKROW)), "a netting set not in the bank file has a bank row"
        Check mCounts(RCN_MISSING) = 1 And mCounts(RCN_NOTINBANK) = 1, _
              "counts missing/not in bank: " & mCounts(RCN_MISSING) & "/" & mCounts(RCN_NOTINBANK)
        Exit Sub

Unexpected:
        Fail mCurrentCase, "unexpected error " & Err.Number & ": " & Err.Description

End Sub


Private Sub CaseDuplicate()
'
'==============================================================================
' PURPOSE
'   An ID listed twice in the bank file is reported on each of its rows
'   and none of its figures is compared; its Results netting set is not
'   reported as NOT IN BANK FILE.
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim bank   As Variant    'Bank rows
    Dim calc   As Variant    'Results rows
    Dim found  As Variant    'Reconciliation rows

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
        On Error GoTo Unexpected
        BeginCase "compare.duplicate"
        bank = NewBlock(2, BN_NCOLS)
        SetBankRow bank, 1, "NS1", "CRR", 100
        SetBankRow bank, 2, "ns1", "CRR", 100
        calc = NewBlock(1, RS_NCOLS)
        SetCalcRow calc, 1, "NS1", "CRR", "VALID", 100, 0, 0, 0
        found = Compare(bank, calc, 1#, 0#, 0#)
        Check RowCount(found) = 2, "rows: " & RowCount(found)
        CheckText "first row", CStr(found(1, RC_RESULT)), RES_DUPLICATE
        CheckText "second row", CStr(found(2, RC_RESULT)), RES_DUPLICATE
        Check mCounts(RCN_COMPARED) = 0 And mCounts(RCN_NOTINBANK) = 0 And mCounts(RCN_DUPLICATE) = 2, _
              "counts compared/not in bank/duplicate: " & mCounts(RCN_COMPARED) & "/" & mCounts(RCN_NOTINBANK) & _
              "/" & mCounts(RCN_DUPLICATE)
        Exit Sub

Unexpected:
        Fail mCurrentCase, "unexpected error " & Err.Number & ": " & Err.Description

End Sub


Private Sub CaseRegimeAndWithheld()
'
'==============================================================================
' PURPOSE
'   Figures of another regime than Results, or of a netting set that has
'   no EAD on Results, are not compared: one row says why.
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim bank   As Variant    'Bank rows
    Dim calc   As Variant    'Results rows
    Dim found  As Variant    'Reconciliation rows

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
        On Error GoTo Unexpected
        BeginCase "compare.regime-and-withheld"
        bank = NewBlock(2, BN_NCOLS)
        SetBankRow bank, 1, "NS1", "BCBS", 100, 10
        SetBankRow bank, 2, "NS2", "CRR", 200, 20
        calc = NewBlock(2, RS_NCOLS)
        SetCalcRow calc, 1, "NS1", "CRR", "VALID", 100, 10, 0, 0
        SetCalcRow calc, 2, "NS2", "CRR", "INCOMPLETE: 1 of 3 trade(s) rejected", Empty, Empty, Empty, Empty
        found = Compare(bank, calc, 1#, 0#, 0#)
        Check RowCount(found) = 2, "rows: " & RowCount(found)
        CheckText "other regime", ResultOf(found, "NS1", ""), RES_REGIME
        CheckText "withheld netting set", ResultOf(found, "NS2", ""), RES_WITHHELD
        CheckText "its Results status", CStr(CellOf(found, "NS2", "", RC_CALCSTATUS)), _
                  "INCOMPLETE: 1 of 3 trade(s) rejected"
        Exit Sub

Unexpected:
        Fail mCurrentCase, "unexpected error " & Err.Number & ": " & Err.Description

End Sub


Private Sub CaseInvalidInput()
'
'==============================================================================
' PURPOSE
'   An unusable bank row or figure is INVALID INPUT: no ID, no valid
'   regime, no figure, a number entered as text, an error value or a
'   logical value. The usable figures of the same row are still compared.
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim bank   As Variant    'Bank rows
    Dim calc   As Variant    'Results rows
    Dim found  As Variant    'Reconciliation rows

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
        On Error GoTo Unexpected
        BeginCase "compare.invalid-input"
        bank = NewBlock(5, BN_NCOLS)
        SetBankRow bank, 1, "", "CRR", 100
        SetBankRow bank, 2, "NS2", "", 100
        SetBankRow bank, 3, "NS3", "CRR"
        bank(3, BN_COMMENT) = "comment only, no figure"
        SetBankRow bank, 4, "NS4", "CRR", "1,000", CVErr(xlErrNA), True, 50
        bank(5, BN_COMMENT) = "a comment alone is not a row"
        calc = NewBlock(3, RS_NCOLS)
        SetCalcRow calc, 1, "NS2", "CRR", "VALID", 100, 0, 0, 0
        SetCalcRow calc, 2, "NS3", "CRR", "VALID", 100, 0, 0, 0
        SetCalcRow calc, 3, "NS4", "CRR", "VALID", 1000, 0, 50, 50
        found = Compare(bank, calc, 1#, 0#, 0#)
        CheckText "row without ID", ResultOf(found, "", ""), RES_INVALID
        CheckText "row without regime", ResultOf(found, "NS2", ""), RES_INVALID
        CheckText "row without figures", ResultOf(found, "NS3", ""), RES_INVALID
        CheckText "number as text", ResultOf(found, "NS4", "EAD"), RES_INVALID
        CheckText "error value", ResultOf(found, "NS4", "RC"), RES_INVALID
        CheckText "logical value", ResultOf(found, "NS4", "PFE"), RES_INVALID
        CheckText "usable figure on the same row", ResultOf(found, "NS4", "AddOn aggregate"), RES_MATCH
        Check mCounts(RCN_INVALID) = 6 And RowCount(found) = 7, _
              "invalid count " & mCounts(RCN_INVALID) & ", rows " & RowCount(found)
        Check CLng(CellOf(found, "NS4", "EAD", RC_BANKROW)) = 4, "bank row of NS4: " & CellOf(found, "NS4", "EAD", RC_BANKROW)
        Exit Sub

Unexpected:
        Fail mCurrentCase, "unexpected error " & Err.Number & ": " & Err.Description

End Sub


Private Sub CaseIdsAndTotals()
'
'==============================================================================
' PURPOSE
'   IDs match regardless of case and surrounding spaces; the TOTAL rows of
'   Results are not netting sets, but a netting set whose ID only starts
'   with TOTAL is; an empty bank file has no rows.
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim bank   As Variant    'Bank rows
    Dim calc   As Variant    'Results rows
    Dim found  As Variant    'Reconciliation rows

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
        On Error GoTo Unexpected
        BeginCase "compare.ids-and-totals"
        bank = NewBlock(2, BN_NCOLS)
        SetBankRow bank, 1, " ns1 ", "crr", 100
        SetBankRow bank, 2, "TOTAL DESK", "CRR", 50
        calc = NewBlock(4, RS_NCOLS)
        SetCalcRow calc, 1, "NS1", "CRR", "VALID", 100, 0, 0, 0
        SetCalcRow calc, 2, "TOTAL DESK", "CRR", "VALID", 50, 0, 0, 0
        SetCalcRow calc, 3, "TOTAL CRR", "CRR", "2 of 2 netting set(s) VALID", 150, 0, 0, Empty
        SetCalcRow calc, 4, "TOTAL BCBS", "BCBS", "0 of 0", 0, 0, 0, Empty
        found = Compare(bank, calc, 1#, 0#, 0#)
        CheckText "ID and regime in another case", ResultOf(found, "NS1", "EAD"), RES_MATCH
        CheckText "netting set named TOTAL DESK", ResultOf(found, "TOTAL DESK", "EAD"), RES_MATCH
        Check RowCount(found) = 2, "TOTAL rows reported: " & RowCount(found)
        bank = NewBlock(2, BN_NCOLS)
        bank(2, BN_COMMENT) = "comment"
        Check Not CORE_Reconcile.HasBankRows(bank), "a table with a comment only has bank rows"
        Check Not CORE_Reconcile.HasBankRows(Empty), "no table has bank rows"
        Exit Sub

Unexpected:
        Fail mCurrentCase, "unexpected error " & Err.Number & ": " & Err.Description

End Sub


Private Sub CaseSettingsAndHeaders()
'
'==============================================================================
' PURPOSE
'   The scale must be a number above zero and the tolerances numbers of
'   zero or more; a moved bank column is found by its header.
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim headers   As Variant    'Header row, (1, column)
    Dim parts     As Variant    'Expected headers
    Dim c         As Long       'Column

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
        On Error GoTo Unexpected
        BeginCase "settings-and-headers"
        Check CORE_Reconcile.SettingsProblem(1000#, 1#, 0.0001) = "", "valid settings rejected"
        Check InStr(CORE_Reconcile.SettingsProblem(0#, 1#, 0#), NM_BN_SCALE) > 0, "zero scale accepted"
        Check InStr(CORE_Reconcile.SettingsProblem("1000", 1#, 0#), NM_BN_SCALE) > 0, "scale as text accepted"
        Check InStr(CORE_Reconcile.SettingsProblem(1#, -1#, 0#), NM_BN_ABSTOL) > 0, "negative tolerance accepted"
        Check InStr(CORE_Reconcile.SettingsProblem(1#, 0#, Empty), NM_BN_RELTOL) > 0, "blank tolerance accepted"
        parts = Split(BN_HEADERS, "|")
        headers = NewBlock(1, BN_NCOLS)
        For c = 0 To UBound(parts)
            headers(1, c + 1) = UCase$(Replace(parts(c), " ", "  "))
        Next c
        Check CORE_Reconcile.HeadersProblem(headers, BN_HEADERS) = "", "headers in another case rejected"
        headers(1, 4) = parts(4)
        headers(1, 5) = parts(3)
        Check InStr(CORE_Reconcile.HeadersProblem(headers, BN_HEADERS), "column 4") > 0, "swapped columns accepted"
        Exit Sub

Unexpected:
        Fail mCurrentCase, "unexpected error " & Err.Number & ": " & Err.Description

End Sub


Private Sub CaseSummary()
'
'==============================================================================
' PURPOSE
'   The summary line names every count, the scale and the tolerances.
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim counts()   As Long      'Counts
    Dim text       As String    'Summary line
    Dim i          As Long      'Count

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
        On Error GoTo Unexpected
        BeginCase "summary"
        ReDim counts(1 To RCN_COUNT)
        For i = 1 To RCN_COUNT
            counts(i) = 10 + i
        Next i
        text = CORE_Reconcile.SummaryText(counts, 1000#, 1#, 0.0001)
        Check InStr(text, "bank netting sets 11") > 0 And InStr(text, "figures compared 12") > 0 And _
              InStr(text, "MATCH 13") > 0 And InStr(text, "DIFFERENCE 14") > 0, "compared counts: " & text
        Check InStr(text, "missing in calculator 15") > 0 And InStr(text, "not in bank file 16") > 0 And _
              InStr(text, "duplicate rows 17") > 0 And InStr(text, "regime mismatch 18") > 0 And _
              InStr(text, "withheld 19") > 0 And InStr(text, "invalid input 20") > 0, "other counts: " & text
        Check InStr(text, "scale 1000") > 0 And InStr(text, "tolerance 1 or ") > 0, "settings: " & text
        Exit Sub

Unexpected:
        Fail mCurrentCase, "unexpected error " & Err.Number & ": " & Err.Description

End Sub


'
'------------------------------------------------------------------------------
'
'                                WORKBOOK CASE
'
'------------------------------------------------------------------------------
'

Private Sub CaseWorkbook()
'
'==============================================================================
' PURPOSE
'   In the workbook: a run with no bank figures makes no reconciliation; a
'   run with figures taken from Results writes MATCH, DIFFERENCE, MISSING
'   IN CALCULATOR and NOT IN BANK FILE rows, its summary and a current
'   status; a failed reconciliation keeps its cause in the run message
'   and the summary (#318); an edited bank figure makes the status OUT OF
'   DATE; a run whose recalculation does not finish and validation each
'   withdraw the reconciliation. The bank table and settings are put back.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim wsBank      As Worksheet    'BankNettingSets
    Dim wsRecon     As Worksheet    'Reconciliation
    Dim wsResults   As Worksheet    'Results
    Dim firstRow    As Long         'First bank row
    Dim saved       As Variant      'Bank table formulas, first 50 rows
    Dim savedSet    As Variant      'Settings formulas, (scale, abs, rel)
    Dim result      As String       'Result line of a run
    Dim r           As Long         'Results row of the first VALID netting set
    Dim nsId        As String       'Its ID
    Dim lastR       As Long         'Last Results row
    Dim outcomes    As Variant      'Reconciliation rows
    Dim i           As Long         'Row

'------------------------------------------------------------------------------
' SAVE THE BANK TABLE
'------------------------------------------------------------------------------
        On Error GoTo Unexpected
        BeginCase "workbook.run-edit-validate"
        Application.Calculation = xlCalculationAutomatic
        M_Main.gSilent = True
        Set wsBank = GetSheet(SH_BANK_NS)
        Set wsRecon = GetSheet(SH_RECON)
        Set wsResults = GetSheet(SH_RESULTS)
        firstRow = NamedRow(NM_BN_HEADER) + 1
        saved = TableRange(wsBank, firstRow, 1, firstRow + 49, BN_NCOLS).Formula
        savedSet = Array(NamedCell(NM_BN_SCALE).Formula, NamedCell(NM_BN_ABSTOL).Formula, _
                         NamedCell(NM_BN_RELTOL).Formula)

'------------------------------------------------------------------------------
' NO BANK FIGURES
'------------------------------------------------------------------------------
        TableRange(wsBank, firstRow, 1, firstRow + 49, BN_NCOLS).ClearContents
        result = M_Main.RunSACCR_Silent()
        Check InStr(result, "; reconciliation=NONE;") > 0, "run without bank figures: " & result
        Check InStr(SafeStr(wsRecon.Range(RECON_INFO_CELL).Value), "No bank figures") = 1, _
              "summary without bank figures: " & SafeStr(wsRecon.Range(RECON_INFO_CELL).Value)

'------------------------------------------------------------------------------
' FIGURES TAKEN FROM RESULTS
'------------------------------------------------------------------------------
    'The first VALID netting set: its EAD as on Results (MATCH) and its RC
    'ten percent higher (DIFFERENCE); and an ID that Results does not have.
        lastR = LastDataRow(wsResults, OutputFirstRow(SH_RESULTS), RS_ID)
        For i = lastR To OutputFirstRow(SH_RESULTS) Step -1
            If SafeStr(TableCell(wsResults, i, RS_STATUS_COL).Value) = "VALID" Then
                r = i
            End If
        Next i
        Check r > 0, "no VALID netting set on Results: build the workbook from the template"
        nsId = SafeStr(TableCell(wsResults, r, RS_ID).Value)
        NamedCell(NM_BN_SCALE).Value = 1
        NamedCell(NM_BN_ABSTOL).Value = 1
        NamedCell(NM_BN_RELTOL).Value = 0.0001
        TableCell(wsBank, firstRow, BN_ID).Value = nsId
        TableCell(wsBank, firstRow, BN_REGIME).Value = TableCell(wsResults, r, RS_REGIME).Value
        TableCell(wsBank, firstRow, BN_EAD).Value = TableCell(wsResults, r, RS_EAD).Value2
        TableCell(wsBank, firstRow, BN_EAD + 1).Value = TableCell(wsResults, r, RS_RC).Value2 * 1.1 + 100
        TableCell(wsBank, firstRow + 1, BN_ID).Value = "NOT-ON-RESULTS"
        TableCell(wsBank, firstRow + 1, BN_REGIME).Value = TableCell(wsResults, r, RS_REGIME).Value
        TableCell(wsBank, firstRow + 1, BN_EAD).Value = 1000
        result = M_Main.RunSACCR_Silent()
        Check InStr(result, "; reconciliation=DONE;") > 0, "run with bank figures: " & result
        outcomes = ReconciliationRows(wsRecon)
        CheckText "EAD as on Results", ResultOf(outcomes, nsId, "EAD"), RES_MATCH
        CheckText "RC ten percent higher", ResultOf(outcomes, nsId, "RC"), RES_DIFF
        CheckText "ID not on Results", ResultOf(outcomes, "NOT-ON-RESULTS", "EAD"), RES_MISSING
    'Every other netting set on Results is not in the bank file: all the
    'rows from the first data row but the TOTAL rows and the one compared.
        Check CountResult(outcomes, RES_NOTINBANK) = LastDataRow(wsResults, OutputFirstRow(SH_RESULTS), RS_ID) - OutputFirstRow(SH_RESULTS) - _
              CountTotals(wsResults), "NOT IN BANK FILE rows: " & CountResult(outcomes, RES_NOTINBANK)
        Check InStr(SafeStr(wsRecon.Range(RECON_INFO_CELL).Value), "Last reconciliation") = 1, _
              "summary: " & SafeStr(wsRecon.Range(RECON_INFO_CELL).Value)
        CheckStatus wsRecon, "Reconciliation current", "after the run"

'------------------------------------------------------------------------------
' FAILED RECONCILIATION
'------------------------------------------------------------------------------
    'Cleaning up a failed reconciliation clears Err: its cause must reach
    'both the summary and the run message (#318). A normal run follows.
        M_Main.gTestFault = "reconcile"
        M_Main.gLastMessage = ""
        result = M_Main.RunSACCR_Silent()
        M_Main.gTestFault = ""
        Check InStr(result, "; reconciliation=FAILED;") > 0, "run with a failed reconciliation: " & result
        Check InStr(M_Main.gLastMessage, "Reconciliation failed: Injected reconciliation failure.") > 0, _
              "run message after a failed reconciliation: " & M_Main.gLastMessage
        Check InStr(SafeStr(wsRecon.Range(RECON_INFO_CELL).Value), "Injected reconciliation failure.") > 0, _
              "summary after a failed reconciliation: " & SafeStr(wsRecon.Range(RECON_INFO_CELL).Value)
        result = M_Main.RunSACCR_Silent()
        Check InStr(result, "; reconciliation=DONE;") > 0, "run after the failed reconciliation: " & result

'------------------------------------------------------------------------------
' EDITED BANK FIGURE, STOPPED RUN, THEN VALIDATION
'------------------------------------------------------------------------------
    'A run that stops in its recalculation raises in silent mode; the
    'reconciliation must be gone all the same.
        TableCell(wsBank, firstRow + 1, BN_EAD).Value = 2000
        CheckStatus wsRecon, "OUT OF DATE", "after a bank figure was edited"
        M_Main.gTestFault = "recalc"
        On Error Resume Next
        result = M_Main.RunSACCR_Silent()
        Err.Clear
        On Error GoTo Unexpected
        M_Main.gTestFault = ""
        Check RowCount(ReconciliationRows(wsRecon)) = 0, "rows left after a run whose recalculation did not finish"
        result = M_Main.RunSACCR_Silent()
        Check RowCount(ReconciliationRows(wsRecon)) > 0, "no rows after a new run: " & result
        M_Main.ValidateInputs
        Check RowCount(ReconciliationRows(wsRecon)) = 0, "rows left after validation"
        CheckStatus wsRecon, "No reconciliation", "after validation"

'------------------------------------------------------------------------------
' RESTORE
'------------------------------------------------------------------------------
        RestoreBankTable wsBank, firstRow, BN_NCOLS, saved, savedSet, NM_BN_SCALE, NM_BN_ABSTOL, NM_BN_RELTOL
        result = M_Main.RunSACCR_Silent()
        Check Left$(result, 9) = "RESULT=OK", "closing run: " & result
        Exit Sub

Unexpected:
        Fail mCurrentCase, "unexpected error " & Err.Number & ": " & Err.Description
        M_Main.gTestFault = ""
        RestoreBankTable wsBank, firstRow, BN_NCOLS, saved, savedSet, NM_BN_SCALE, NM_BN_ABSTOL, NM_BN_RELTOL

End Sub


Private Sub RestoreBankTable( _
    ByVal wsBank As Worksheet, _
    ByVal firstRow As Long, _
    ByVal nCols As Long, _
    ByVal saved As Variant, _
    ByVal savedSet As Variant, _
    ByVal scaleName As String, _
    ByVal absName As String, _
    ByVal relName As String)
'
'==============================================================================
' PURPOSE
'   Put a bank table and its settings back as a workbook case found them.
'
' ERROR POLICY
'   Contained: a failure is reported as a failure of the case.
'==============================================================================
'
        On Error GoTo Failed
        If wsBank Is Nothing Or Not IsArray(saved) Or Not IsArray(savedSet) Then
            Exit Sub
        End If
        TableRange(wsBank, firstRow, 1, firstRow + 49, nCols).Formula = saved
        NamedCell(scaleName).Formula = savedSet(0)
        NamedCell(absName).Formula = savedSet(1)
        NamedCell(relName).Formula = savedSet(2)
        Exit Sub

Failed:
        Fail mCurrentCase, "could not restore " & wsBank.Name & ": " & Err.Description

End Sub


'
'------------------------------------------------------------------------------
'
'                                 TRADE CASES
'
'------------------------------------------------------------------------------
'

Private Sub CaseTradeFigures()
'
'==============================================================================
' PURPOSE
'   Trade-leg amounts are scaled and the delta and maturity factors are
'   not, with only the relative tolerance; a figure the calculator does
'   not have for the trade is NOT APPLICABLE; the trade's netting set and
'   asset class are shown.
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim bank   As Variant    'Bank rows
    Dim calc   As Variant    'TradeCalc rows
    Dim found  As Variant    'Reconciliation rows

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
        On Error GoTo Unexpected
        BeginCase "trade.figures"
        bank = NewBlock(1, BT_NCOLS)
        bank(1, BT_ID) = "T1"
        bank(1, BT_NOTIONAL) = 10
        bank(1, BT_NOTIONAL + 5) = 0.5
        bank(1, BT_NOTIONAL + 6) = 1
        bank(1, BT_NOTIONAL + 7) = 0.3
        calc = NewBlock(1, TC_NCOLS)
        SetTradeCalcRow calc, 1, "T1", "NS1", "IR", "OK", 10000.4, 0.6, 1, Empty, ""
        found = CompareT(bank, calc, 1000#, 1#, 0.0001)
        CheckText "notional in thousands", ResultOf(found, "T1", "Notional"), RES_MATCH
        CheckText "delta outside the relative tolerance", ResultOf(found, "T1", "Supervisory delta"), RES_DIFF
        CheckNear "delta not scaled", CellOf(found, "T1", "Supervisory delta", RC_BANK), 0.5
        CheckText "equal maturity factor", ResultOf(found, "T1", "MF unmargined"), RES_MATCH
        CheckText "no margined maturity factor", ResultOf(found, "T1", "MF margined"), RES_NA
        CheckText "netting set shown", CStr(CellOf(found, "T1", "Notional", TRC_NS)), "NS1"
        Check mCounts(RCN_NA) = 1 And mCounts(RCN_COMPARED) = 3, _
              "counts not applicable/compared: " & mCounts(RCN_NA) & "/" & mCounts(RCN_COMPARED)
        Exit Sub

Unexpected:
        Fail mCurrentCase, "unexpected error " & Err.Number & ": " & Err.Description

End Sub


Private Sub CaseTradePopulations()
'
'==============================================================================
' PURPOSE
'   A bank leg absent from TradeCalc is MISSING IN CALCULATOR; a TradeCalc
'   leg absent from the bank file is NOT IN BANK FILE with its notional; an
'   excluded leg is EXCLUDED with the calculator's reasons; a leg listed
'   twice is DUPLICATE on each row; a row without TradeID is INVALID INPUT.
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim bank   As Variant    'Bank rows
    Dim calc   As Variant    'TradeCalc rows
    Dim found  As Variant    'Reconciliation rows

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
        On Error GoTo Unexpected
        BeginCase "trade.populations"
        bank = NewBlock(5, BT_NCOLS)
        bank(1, BT_ID) = "T2"
        bank(1, BT_NOTIONAL) = 100
        bank(2, BT_ID) = "T4"
        bank(2, BT_NOTIONAL) = 100
        bank(3, BT_ID) = "T5"
        bank(3, BT_NOTIONAL) = 100
        bank(4, BT_ID) = "T5"
        bank(4, BT_NOTIONAL) = 100
        bank(5, BT_NOTIONAL) = 100
        calc = NewBlock(3, TC_NCOLS)
        SetTradeCalcRow calc, 1, "T3", "NS1", "FX", "OK", 300, 1, 1, Empty, ""
        SetTradeCalcRow calc, 2, "T4", "NS1", "IR", "EXCLUDED", Empty, Empty, Empty, Empty, "Maturity date missing"
        SetTradeCalcRow calc, 3, "T5", "NS1", "IR", "OK", 100, 1, 1, Empty, ""
        found = CompareT(bank, calc, 1#, 0#, 0#)
        CheckText "bank leg not on TradeCalc", ResultOf(found, "T2", "Notional"), RES_MISSING
        CheckText "TradeCalc leg not in the bank file", ResultOf(found, "T3", "Notional"), RES_NOTINBANK
        CheckNear "its notional", CellOf(found, "T3", "Notional", RC_CALC), 300#
        CheckText "excluded leg", ResultOf(found, "T4", ""), RES_EXCLUDED
        Check InStr(SafeStr(CellOf(found, "T4", "", RC_NOTE)), "Maturity date missing") > 0, _
              "reasons of the excluded leg: " & SafeStr(CellOf(found, "T4", "", RC_NOTE))
        CheckText "leg listed twice", ResultOf(found, "T5", ""), RES_DUPLICATE
        CheckText "row without TradeID", ResultOf(found, "", ""), RES_INVALID
        Check mCounts(RCN_DUPLICATE) = 2 And mCounts(RCN_EXCLUDED) = 1 And RowCount(found) = 6, _
              "duplicate/excluded/rows: " & mCounts(RCN_DUPLICATE) & "/" & mCounts(RCN_EXCLUDED) & "/" & RowCount(found)
        Exit Sub

Unexpected:
        Fail mCurrentCase, "unexpected error " & Err.Number & ": " & Err.Description

End Sub


Private Sub CaseTradeCalcDuplicates()
'
'==============================================================================
' PURPOSE
'   A Trade ID listed twice on TradeCalc (repeated on Trades, the later row
'   EXCLUDED by the engine) is DUPLICATE and not compared, with or without
'   a bank row; it is not NOT IN BANK FILE.
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim bank   As Variant    'Bank rows
    Dim calc   As Variant    'TradeCalc rows
    Dim found  As Variant    'Reconciliation rows

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
        On Error GoTo Unexpected
        BeginCase "trade.calc-duplicates"
        bank = NewBlock(1, BT_NCOLS)
        bank(1, BT_ID) = "T7"
        bank(1, BT_NOTIONAL) = 100
        calc = NewBlock(4, TC_NCOLS)
        SetTradeCalcRow calc, 1, "T7", "NS1", "IR", "OK", 100, 1, 1, Empty, ""
        SetTradeCalcRow calc, 2, "T7", "NS1", "IR", "EXCLUDED", Empty, Empty, Empty, Empty, "duplicate Trade ID"
        SetTradeCalcRow calc, 3, "T8", "NS1", "FX", "OK", 50, 1, 1, Empty, ""
        SetTradeCalcRow calc, 4, "T8", "NS1", "FX", "EXCLUDED", Empty, Empty, Empty, Empty, "duplicate Trade ID"
        found = CompareT(bank, calc, 1#, 0#, 0#)
        CheckText "bank leg listed twice on TradeCalc", ResultOf(found, "T7", ""), RES_DUPLICATE
        CheckText "leg listed twice on TradeCalc, not in the bank file", ResultOf(found, "T8", ""), RES_DUPLICATE
        Check mCounts(RCN_DUPLICATE) = 2 And mCounts(RCN_NOTINBANK) = 0 And mCounts(RCN_COMPARED) = 0 And _
              RowCount(found) = 2, "duplicate/not in bank/compared/rows: " & mCounts(RCN_DUPLICATE) & "/" & _
              mCounts(RCN_NOTINBANK) & "/" & mCounts(RCN_COMPARED) & "/" & RowCount(found)
        Exit Sub

Unexpected:
        Fail mCurrentCase, "unexpected error " & Err.Number & ": " & Err.Description

End Sub


Private Sub CaseTradeSummary()
'
'==============================================================================
' PURPOSE
'   The trade summary names its counts, and a setting problem names the
'   BankTrades cell.
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim counts()   As Long      'Counts
    Dim text       As String    'Summary line
    Dim i          As Long      'Count

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
        On Error GoTo Unexpected
        BeginCase "trade.summary"
        ReDim counts(1 To RCN_COUNT)
        For i = 1 To RCN_COUNT
            counts(i) = 10 + i
        Next i
        text = CORE_Reconcile.TradeSummaryText(counts, 1#, 1#, 0#)
        Check InStr(text, "bank trade legs 11") > 0 And InStr(text, "excluded 21") > 0 And _
              InStr(text, "not applicable 22") > 0 And InStr(text, "not in bank file 16") > 0, "summary: " & text
        Check InStr(CORE_Reconcile.SettingsProblem(0#, 1#, 0#, NM_BT_SCALE, NM_BT_ABSTOL, NM_BT_RELTOL), _
                    NM_BT_SCALE) > 0, "zero BankTrades scale accepted"
        Exit Sub

Unexpected:
        Fail mCurrentCase, "unexpected error " & Err.Number & ": " & Err.Description

End Sub


Private Sub CaseTradeWorkbook()
'
'==============================================================================
' PURPOSE
'   In the workbook: figures taken from the first OK leg on TradeCalc give
'   a MATCH and a DIFFERENCE, an unknown leg is MISSING IN CALCULATOR, the
'   run reports trade_reconciliation=DONE, the TradeReconciliation status
'   is current and an edited bank figure makes it OUT OF DATE; an Economic
'   trade ID set on Trades for the leg is shown on its rows. The BankTrades
'   table and settings, and the Trades cells, are put back.
'
' UPDATED
'   2026-10-15
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim wsBank     As Worksheet    'BankTrades
    Dim wsRecon    As Worksheet    'TradeReconciliation
    Dim wsCalc     As Worksheet    'TradeCalc
    Dim firstRow   As Long         'First bank row
    Dim saved      As Variant      'Bank table formulas, first 50 rows
    Dim savedSet   As Variant      'Settings formulas, (scale, abs, rel)
    Dim result     As String       'Result line of a run
    Dim r          As Long         'TradeCalc row of the first OK leg
    Dim i          As Long         'TradeCalc row
    Dim legId      As String       'Its TradeID
    Dim outcomes   As Variant      'TradeReconciliation rows
    Dim wsTrades   As Worksheet    'Trades
    Dim tradeRow   As Long         'Trades row of the leg
    Dim econCell   As Range        'Its Economic trade ID, set and restored
    Dim mtmLegCell As Range        'Its MtM leg flag, set and restored
    Dim savedEcon  As Variant      'Formula of econCell
    Dim savedMtmLeg As Variant     'Formula of mtmLegCell

'------------------------------------------------------------------------------
' SAVE THE BANK TABLE
'------------------------------------------------------------------------------
        On Error GoTo Unexpected
        BeginCase "trade.workbook-run-edit"
        Application.Calculation = xlCalculationAutomatic
        M_Main.gSilent = True
        Set wsBank = GetSheet(SH_BANK_TR)
        Set wsRecon = GetSheet(SH_TRADE_RECON)
        Set wsCalc = GetSheet(SH_TRADECALC)
        firstRow = NamedRow(NM_BT_HEADER) + 1
        saved = TableRange(wsBank, firstRow, 1, firstRow + 49, BT_NCOLS).Formula
        savedSet = Array(NamedCell(NM_BT_SCALE).Formula, NamedCell(NM_BT_ABSTOL).Formula, _
                         NamedCell(NM_BT_RELTOL).Formula)

'------------------------------------------------------------------------------
' FIGURES TAKEN FROM TRADECALC
'------------------------------------------------------------------------------
        TableRange(wsBank, firstRow, 1, firstRow + 49, BT_NCOLS).ClearContents
        result = M_Main.RunSACCR_Silent()
        For i = LastDataRow(wsCalc, OutputFirstRow(SH_TRADECALC), TC_ID) To OutputFirstRow(SH_TRADECALC) Step -1
            If SafeStr(TableCell(wsCalc, i, TC_STATUS).Value) = "OK" Then
                r = i
            End If
        Next i
        Check r > 0, "no OK leg on TradeCalc: build the workbook from the template"
        legId = SafeStr(TableCell(wsCalc, r, TC_ID).Value)
        NamedCell(NM_BT_SCALE).Value = 1
        NamedCell(NM_BT_ABSTOL).Value = 1
        NamedCell(NM_BT_RELTOL).Value = 0.0001
        TableCell(wsBank, firstRow, BT_ID).Value = legId
        TableCell(wsBank, firstRow, BT_NOTIONAL).Value = TableCell(wsCalc, r, TC_NOTIONAL).Value2
        TableCell(wsBank, firstRow, BT_NOTIONAL + 5).Value = TableCell(wsCalc, r, TC_DELTA).Value2 + 0.1
        TableCell(wsBank, firstRow + 1, BT_ID).Value = "NO-SUCH-LEG"
        TableCell(wsBank, firstRow + 1, BT_NOTIONAL).Value = 1000
        result = M_Main.RunSACCR_Silent()
        Check InStr(result, "; trade_reconciliation=DONE;") > 0, "run with bank trade figures: " & result
        outcomes = ReconciliationRows(wsRecon)
        CheckText "notional as on TradeCalc", ResultOf(outcomes, legId, "Notional"), RES_MATCH
        CheckText "delta 0.1 higher", ResultOf(outcomes, legId, "Supervisory delta"), RES_DIFF
        CheckText "leg not on TradeCalc", ResultOf(outcomes, "NO-SUCH-LEG", "Notional"), RES_MISSING
        CheckStatus wsRecon, "Reconciliation current", "after the run"
        TableCell(wsBank, firstRow + 1, BT_NOTIONAL).Value = 2000
        CheckStatus wsRecon, "OUT OF DATE", "after a bank figure was edited"

'------------------------------------------------------------------------------
' ECONOMIC TRADE ID
'------------------------------------------------------------------------------
    'The leg made an economic trade of one leg, which must then be its MtM
    'leg; its rows show the ID, a leg not on Trades shows none.
        Set wsTrades = GetSheet(SH_TRADES)
        For i = TableFirstRow(SH_TRADES) To LastDataRow(wsTrades, TableFirstRow(SH_TRADES), TR_ID)
            If UTxt(TableCell(wsTrades, i, TR_ID).Value) = UTxt(legId) Then
                tradeRow = i
            End If
        Next i
        Check tradeRow > 0, "leg " & legId & " not found on Trades"
        If tradeRow > 0 Then
            Set econCell = TableCell(wsTrades, tradeRow, TR_ECON)
            Set mtmLegCell = TableCell(wsTrades, tradeRow, TR_MTMLEG)
            savedEcon = econCell.Formula
            savedMtmLeg = mtmLegCell.Formula
            econCell.Value = "ECON-T1"
            mtmLegCell.Value = "Y"
            result = M_Main.RunSACCR_Silent()
            outcomes = ReconciliationRows(wsRecon)
            CheckText "Economic trade ID of the leg", SafeStr(CellOf(outcomes, legId, "Notional", TRC_ECON)), "ECON-T1"
            CheckText "Economic trade ID of a leg not on Trades", _
                      SafeStr(CellOf(outcomes, "NO-SUCH-LEG", "Notional", TRC_ECON)), ""
            econCell.Formula = savedEcon
            mtmLegCell.Formula = savedMtmLeg
            Check UBound(ColumnValues(wsTrades, tradeRow, tradeRow, TR_ID), 1) = 1, _
                  "a one-row Trades column is not read as an array"
        End If

'------------------------------------------------------------------------------
' RESTORE
'------------------------------------------------------------------------------
        RestoreBankTable wsBank, firstRow, BT_NCOLS, saved, savedSet, NM_BT_SCALE, NM_BT_ABSTOL, NM_BT_RELTOL
        result = M_Main.RunSACCR_Silent()
        Check Left$(result, 9) = "RESULT=OK", "closing run: " & result
        Exit Sub

Unexpected:
        Fail mCurrentCase, "unexpected error " & Err.Number & ": " & Err.Description
        RestoreCell econCell, savedEcon
        RestoreCell mtmLegCell, savedMtmLeg
        RestoreBankTable wsBank, firstRow, BT_NCOLS, saved, savedSet, NM_BT_SCALE, NM_BT_ABSTOL, NM_BT_RELTOL

End Sub


'
'------------------------------------------------------------------------------
'
'                                EXPORT CASES
'
'------------------------------------------------------------------------------
'

Private Sub CaseExportFields()
'
'==============================================================================
' PURPOSE
'   The evidence CSV writes numbers with a decimal point and no thousands
'   separator in every locale, quotes only text that needs it, writes text
'   Excel would read as a formula with a leading apostrophe, names error
'   values as Excel does, ends lines with CR LF and encodes UTF-8 without a
'   byte-order mark, including characters outside the BMP.
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim b()   As Byte    'UTF-8 bytes

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
        On Error GoTo Unexpected
        BeginCase "export.fields"
        CheckText "number", CORE_Export.NumberText(48595558.5), "48595558.5"
        CheckText "number below one", CORE_Export.NumberText(0.0001), "0.0001"
        CheckText "negative number below one", CORE_Export.NumberText(-0.5), "-0.5"
        CheckText "text with a separator", CORE_Export.CsvField("a;b"), """a;b"""
        CheckText "text with a quote", CORE_Export.CsvField("say ""x"""), """say """"x"""""""
        CheckText "plain text", CORE_Export.CsvField("NS-ALPHA"), "NS-ALPHA"
        CheckText "text read as a formula", CORE_Export.CsvField("=1+1"), "'=1+1"
        CheckText "text starting with a minus", CORE_Export.CsvField("-x;y"), """'-x;y"""
        CheckText "logical value", CORE_Export.CsvField(True), "TRUE"
        CheckText "error value", CORE_Export.CsvField(CVErr(xlErrNA)), "#N/A"
        CheckText "newer error value", CORE_Export.CsvField(CVErr(2045)), "#SPILL!"
        CheckText "empty cell", CORE_Export.CsvField(Empty), ""
        CheckText "line", CORE_Export.CsvLine("RUN", Array("alpha", 1.4)), "RUN;alpha;1.4" & vbCrLf
        b = CORE_Export.Utf8Bytes("A" & ChrW(&HE9) & ChrW(&H20AC))
        Check UBound(b) = 5 And b(0) = &H41 And b(1) = &HC3 And b(2) = &HA9 And b(3) = &HE2 And b(4) = &H82 And _
              b(5) = &HAC, "UTF-8 of A, e acute and the euro sign"
        b = CORE_Export.Utf8Bytes(ChrW(&HD83D) & ChrW(&HDE00))
        Check UBound(b) = 3 And b(0) = &HF0 And b(1) = &H9F And b(2) = &H98 And b(3) = &H80, _
              "UTF-8 of a character outside the BMP"
        Exit Sub

Unexpected:
        Fail mCurrentCase, "unexpected error " & Err.Number & ": " & Err.Description

End Sub


Private Sub CaseExportWorkbook()
'
'==============================================================================
' PURPOSE
'   In the workbook: after a run the evidence report is written with its
'   sections, and a second export of the same run is identical byte for
'   byte, also over a longer hidden file of that name; an edited input, bank figures entered after the run, bank
'   figures removed after a reconciliation and an edited output cell each
'   make the export refuse, as does a recalculation pending in manual
'   calculation; undoing them, or a new run, allows it again. A
'   currency added below the FX table's named range is in the report.
'   The edited cells and the files are put back or removed.
'
' UPDATED
'   2026-10-16
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim folder      As String       'Folder for the test files
    Dim fileA       As String       'First export
    Dim fileB       As String       'Second export
    Dim result      As String       'Result line
    Dim textA       As String       'First export as text
    Dim mtmCell     As Range        'An input edited and restored
    Dim savedMtm    As Variant      'Its formula
    Dim bankCell    As Range        'A bank figure entered and removed
    Dim savedBank   As Variant      'Its formula
    Dim outCell     As Range        'An output edited and restored
    Dim savedOut    As Variant      'Its value
    Dim wsParams    As Worksheet    'Params
    Dim fxCol       As Long         'Key column of the FX table
    Dim fxRow       As Long         'First blank row below it
    Dim fxCells     As Range        'A currency added and removed
    Dim savedFx     As Variant      'Their formulas
    Dim fileNo      As Integer      'File number of a file written by the case

'------------------------------------------------------------------------------
' EXPORT TWICE
'------------------------------------------------------------------------------
        On Error GoTo Unexpected
        BeginCase "export.workbook"
        Application.Calculation = xlCalculationAutomatic
        M_Main.gSilent = True
        folder = Environ$("TEMP")
        If Len(folder) = 0 Then
            folder = ThisWorkbook.Path
        End If
        fileA = folder & Application.PathSeparator & "K-SACCR_evidence_test_a.csv"
        fileB = folder & Application.PathSeparator & "K-SACCR_evidence_test_b.csv"
        result = M_Main.RunSACCR_Silent()
        result = M_Main.ExportEvidence_Silent(fileA)
        Check Left$(result, 9) = "RESULT=OK", "first export: " & result
        textA = FileText(fileA)
        Check Left$(textA, 31) = "RUN;format;K-SACCR evidence 1" & vbCrLf, "first line: " & Left$(textA, 40)
        Check InStr(textA, vbCrLf & "RUN;results_status;CURRENT" & vbCrLf) > 0, "results status line missing"
        Check InStr(textA, vbCrLf & "RUN;reference_parameter_set;") > 0 And InStr(textA, "RUN;workbook;") = 0, _
              "run identity lines"
        Check InStr(textA, vbCrLf & "RUN;reconciliation;NOT MADE" & vbCrLf) > 0 And _
              InStr(textA, vbCrLf & "RUN;reconciliation_scale;") = 0, "settings of a reconciliation that was not made"
        Check Not (SafeStr(GetSheet(SH_RESULTS).Range(RUNINFO_CELL).Value2) Like "*#,# s*"), _
              "run time with a decimal comma: " & SafeStr(GetSheet(SH_RESULTS).Range(RUNINFO_CELL).Value2)
        Check InStr(textA, vbCrLf & "PARAMS;Alpha;1.4" & vbCrLf) > 0, "Alpha line missing or not invariant"
        Check InStr(textA, vbCrLf & "TRADE_RECONCILIATION;TradeID;") > 0 And _
              InStr(textA, ";Note;Economic trade ID" & vbCrLf) > 0, "trade reconciliation header line"
        Check InStr(textA, vbCrLf & "RESULTS;NettingSetID;") > 0 And InStr(textA, vbCrLf & "CHECKS;") > 0 And _
              InStr(textA, vbCrLf & "LIMITATIONS;1;") > 0, "a section is missing"
        result = M_Main.ExportEvidence_Silent(fileB)
        Check FileText(fileB) = textA, "a second export of the same run differs"
        fileNo = FreeFile
        Open fileB For Output As #fileNo
        Print #fileNo, String$(Len(textA) + 100, "x")
        Close #fileNo
        SetAttr fileB, vbHidden
        result = M_Main.ExportEvidence_Silent(fileB)
        Check Left$(result, 9) = "RESULT=OK" And FileText(fileB) = textA, _
              "a longer hidden file is not replaced whole: " & result

'------------------------------------------------------------------------------
' REFUSE WHEN NOT CURRENT
'------------------------------------------------------------------------------
        Set mtmCell = TableCell(GetSheet(SH_TRADES), TableFirstRow(SH_TRADES), TR_MTM)
        savedMtm = mtmCell.Formula
        mtmCell.Value = ToDbl(mtmCell.Value, 0#) + 1#
        result = M_Main.ExportEvidence_Silent(fileB)
        Check Left$(result, 14) = "RESULT=REFUSED", "export after an input edit: " & result
        mtmCell.Formula = savedMtm
        result = M_Main.ExportEvidence_Silent(fileB)
        Check Left$(result, 9) = "RESULT=OK", "export after undoing the edit: " & result
    'In manual calculation the same value written again leaves the status
    'formulas to recalculate: the export waits for F9.
        Application.Calculation = xlCalculationManual
        mtmCell.Formula = savedMtm
        result = M_Main.ExportEvidence_Silent(fileB)
        Check Left$(result, 14) = "RESULT=REFUSED" And InStr(result, "recalculating") > 0, _
              "export with a recalculation pending: " & result
        Application.Calculate
        Application.Calculation = xlCalculationAutomatic
        result = M_Main.ExportEvidence_Silent(fileB)
        Check Left$(result, 9) = "RESULT=OK", "export after the recalculation: " & result
        Set bankCell = TableCell(GetSheet(SH_BANK_NS), NamedRow(NM_BN_HEADER) + 1, BN_ID)
        savedBank = bankCell.Formula
        bankCell.Value = "NEW-AFTER-RUN"
        result = M_Main.ExportEvidence_Silent(fileB)
        Check Left$(result, 14) = "RESULT=REFUSED" And InStr(result, SH_BANK_NS) > 0, _
              "export with bank figures entered after the run: " & result
        result = M_Main.RunSACCR_Silent()
        result = M_Main.ExportEvidence_Silent(fileB)
        Check Left$(result, 9) = "RESULT=OK", "export with a reconciliation: " & result
        Check InStr(FileText(fileB), vbCrLf & "RUN;reconciliation_scale;") > 0, "settings of the reconciliation missing"
        bankCell.Formula = savedBank
        result = M_Main.ExportEvidence_Silent(fileB)
        Check Left$(result, 14) = "RESULT=REFUSED" And InStr(result, SH_BANK_NS) > 0, _
              "export after every bank figure was removed: " & result

'------------------------------------------------------------------------------
' REFUSE AN EDITED OUTPUT
'------------------------------------------------------------------------------
        result = M_Main.RunSACCR_Silent()
        Set outCell = TableCell(GetSheet(SH_RESULTS), OutputFirstRow(SH_RESULTS), RS_EAD)
        savedOut = outCell.Value2
        outCell.Value2 = ToDbl(savedOut, 0#) + 1#
        result = M_Main.ExportEvidence_Silent(fileB)
        Check Left$(result, 14) = "RESULT=REFUSED" And InStr(result, "output cell") > 0, _
              "export after an output cell was edited: " & result
        outCell.Value2 = savedOut
        result = M_Main.ExportEvidence_Silent(fileB)
        Check Left$(result, 9) = "RESULT=OK", "export after restoring the output cell: " & result

'------------------------------------------------------------------------------
' EXPORT A CURRENCY ADDED BELOW THE FX TABLE
'------------------------------------------------------------------------------
    'The template asks for new currencies directly below the last row,
    'outside the range of the name FX_Table; the run reads them, so the
    'report must list them.
        Set wsParams = GetSheet(SH_PARAMS)
        fxCol = NamedTableCol(NM_FX_TABLE)
        fxRow = NamedRow(NM_FX_TABLE) + 1
        Do While Not IsBlankCell(TableCell(wsParams, fxRow, fxCol).Value)
            fxRow = fxRow + 1
        Loop
        Set fxCells = TableRange(wsParams, fxRow, fxCol, fxRow, fxCol + 2)
        savedFx = fxCells.Formula
        fxCells.Value = Array("ZZX", "Test currency", 2)
        result = M_Main.RunSACCR_Silent()
        result = M_Main.ExportEvidence_Silent(fileB)
        Check InStr(FileText(fileB), vbCrLf & "FX_TABLE;ZZX;Test currency;2" & vbCrLf) > 0, _
              "a currency added below the FX table is missing from the report: " & result
        fxCells.Formula = savedFx
        result = M_Main.RunSACCR_Silent()

'------------------------------------------------------------------------------
' CLEAN UP
'------------------------------------------------------------------------------
        RemoveFile fileA
        RemoveFile fileB
        Exit Sub

Unexpected:
        Fail mCurrentCase, "unexpected error " & Err.Number & ": " & Err.Description
        Application.Calculation = xlCalculationAutomatic
        RestoreCell mtmCell, savedMtm
        RestoreCell bankCell, savedBank
        If Not outCell Is Nothing Then
            outCell.Value2 = savedOut
        End If
        If Not fxCells Is Nothing Then
            fxCells.Formula = savedFx
        End If
        RemoveFile fileA
        RemoveFile fileB

End Sub


Private Function FileText( _
    ByVal filePath As String) _
    As String
'
'==============================================================================
' PURPOSE
'   Read a file's bytes as text, one character per byte; enough to check
'   the ASCII content of an export.
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim fileNo   As Integer    'File number
    Dim b()      As Byte       'The bytes

'------------------------------------------------------------------------------
' READ
'------------------------------------------------------------------------------
        fileNo = FreeFile
        Open filePath For Binary Access Read As #fileNo
        If LOF(fileNo) > 0 Then
            ReDim b(0 To LOF(fileNo) - 1)
            Get #fileNo, , b
            FileText = StrConv(b, vbUnicode)
        End If
        Close #fileNo

End Function


Private Sub RemoveFile( _
    ByVal filePath As String)
'
'==============================================================================
' PURPOSE
'   Delete a test file when it exists.
'
' ERROR POLICY
'   Contained.
'==============================================================================
'
        On Error Resume Next
        If Len(filePath) > 0 Then
            If Len(Dir$(filePath, vbHidden)) > 0 Then
                SetAttr filePath, vbNormal
                Kill filePath
            End If
        End If
        Err.Clear

End Sub


Private Sub RestoreCell( _
    ByVal target As Range, _
    ByVal saved As Variant)
'
'==============================================================================
' PURPOSE
'   Put back a cell edited by a case, after an unexpected error.
'
' ERROR POLICY
'   Contained: a failure is reported as a failure of the case.
'==============================================================================
'
        On Error GoTo Failed
        If Not target Is Nothing Then
            target.Formula = saved
        End If
        Exit Sub

Failed:
        Fail mCurrentCase, "could not restore an edited cell: " & Err.Description

End Sub


'
'------------------------------------------------------------------------------
'
'                                   HELPERS
'
'------------------------------------------------------------------------------
'

Private Function NewBlock( _
    ByVal nRows As Long, _
    ByVal nCols As Long) _
    As Variant
'
'==============================================================================
' PURPOSE
'   Return an empty 1-based block, as Range.Value2 gives one.
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim block()   As Variant    'The block

'------------------------------------------------------------------------------
' CREATE
'------------------------------------------------------------------------------
        ReDim block(1 To nRows, 1 To nCols)
        NewBlock = block

End Function


Private Sub SetBankRow( _
    ByRef bank As Variant, _
    ByVal r As Long, _
    ByVal id As String, _
    ByVal regimeCode As String, _
    Optional ByVal ead As Variant, _
    Optional ByVal rcValue As Variant, _
    Optional ByVal pfe As Variant, _
    Optional ByVal bankAddOn As Variant)
'
'==============================================================================
' PURPOSE
'   Fill a bank row: ID, regime and, when given, EAD, RC, PFE and aggregate
'   add-on. A string "" leaves the cell empty.
'==============================================================================
'
        If Len(id) > 0 Then bank(r, BN_ID) = id
        If Len(regimeCode) > 0 Then bank(r, BN_REGIME) = regimeCode
        If Not IsMissing(ead) Then bank(r, BN_EAD) = ead
        If Not IsMissing(rcValue) Then bank(r, BN_EAD + 1) = rcValue
        If Not IsMissing(pfe) Then bank(r, BN_EAD + 2) = pfe
        If Not IsMissing(bankAddOn) Then bank(r, BN_EAD + 3) = bankAddOn

End Sub


Private Sub SetCalcRow( _
    ByRef calc As Variant, _
    ByVal r As Long, _
    ByVal id As String, _
    ByVal regimeCode As String, _
    ByVal status As String, _
    ByVal ead As Variant, _
    ByVal rcValue As Variant, _
    ByVal pfe As Variant, _
    ByVal multiplier As Variant)
'
'==============================================================================
' PURPOSE
'   Fill a Results row: ID, regime, status, EAD, RC, PFE and multiplier;
'   the aggregate add-on equals the PFE.
'==============================================================================
'
        calc(r, RS_ID) = id
        calc(r, RS_REGIME) = regimeCode
        calc(r, RS_STATUS_COL) = status
        calc(r, RS_EAD) = ead
        calc(r, RS_RC) = rcValue
        calc(r, RS_PFE) = pfe
        calc(r, RS_ADDON) = pfe
        calc(r, RS_MULT) = multiplier

End Sub


Private Function Compare( _
    ByVal bank As Variant, _
    ByVal calc As Variant, _
    ByVal scaleFactor As Double, _
    ByVal absTol As Double, _
    ByVal relTol As Double) _
    As Variant
'
'==============================================================================
' PURPOSE
'   Run CORE_Reconcile.CompareNettingSets with the bank rows starting on sheet
'   row 1, and keep its counts in mCounts.
'==============================================================================
'
        Compare = CORE_Reconcile.CompareNettingSets(bank, 1, calc, scaleFactor, absTol, relTol, mCounts)

End Function


Private Function CompareT( _
    ByVal bank As Variant, _
    ByVal calc As Variant, _
    ByVal scaleFactor As Double, _
    ByVal absTol As Double, _
    ByVal relTol As Double) _
    As Variant
'
'==============================================================================
' PURPOSE
'   Run CORE_Reconcile.CompareTrades with the bank rows starting on sheet
'   row 1, and keep its counts in mCounts.
'==============================================================================
'
        CompareT = CORE_Reconcile.CompareTrades(bank, 1, calc, scaleFactor, absTol, relTol, mCounts)

End Function


Private Sub SetTradeCalcRow( _
    ByRef calc As Variant, _
    ByVal r As Long, _
    ByVal id As String, _
    ByVal setCode As String, _
    ByVal classCode As String, _
    ByVal status As String, _
    ByVal notional As Variant, _
    ByVal delta As Variant, _
    ByVal unmarginedMf As Variant, _
    ByVal marginedMf As Variant, _
    ByVal messages As String)
'
'==============================================================================
' PURPOSE
'   Fill a TradeCalc row: ID, netting set, asset class, status, notional,
'   delta, the two maturity factors and the messages.
'==============================================================================
'
        calc(r, TC_ID) = id
        calc(r, TC_NS) = setCode
        calc(r, TC_AC) = classCode
        calc(r, TC_STATUS) = status
        calc(r, TC_NOTIONAL) = notional
        calc(r, TC_DELTA) = delta
        calc(r, TC_MFU) = unmarginedMf
        calc(r, TC_MFM) = marginedMf
        calc(r, TC_MESSAGES) = messages

End Sub


Private Function RowOf( _
    ByVal outRows As Variant, _
    ByVal id As String, _
    ByVal measure As String) _
    As Long
'
'==============================================================================
' PURPOSE
'   Return the first Reconciliation row of an ID and measure; 0 when there
'   is none. A measure "" finds a row without a measure.
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim r   As Long    'Row

'------------------------------------------------------------------------------
' FIND
'------------------------------------------------------------------------------
        For r = 1 To RowCount(outRows)
            If SafeStr(outRows(r, RC_ID)) = id And SafeStr(outRows(r, RC_MEASURE)) = measure Then
                RowOf = r
                Exit Function
            End If
        Next r

End Function


Private Function ResultOf( _
    ByVal outRows As Variant, _
    ByVal id As String, _
    ByVal measure As String) _
    As String
'
'==============================================================================
' PURPOSE
'   Return the result of an ID and measure; "(no row)" when there is none.
'==============================================================================
'
        ResultOf = SafeStr(CellOf(outRows, id, measure, RC_RESULT))
        If RowOf(outRows, id, measure) = 0 Then
            ResultOf = "(no row)"
        End If

End Function


Private Function CellOf( _
    ByVal outRows As Variant, _
    ByVal id As String, _
    ByVal measure As String, _
    ByVal col As Long) _
    As Variant
'
'==============================================================================
' PURPOSE
'   Return one column of the row of an ID and measure; Empty when there is
'   no such row.
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim r   As Long    'Row

'------------------------------------------------------------------------------
' READ
'------------------------------------------------------------------------------
        r = RowOf(outRows, id, measure)
        If r > 0 Then
            CellOf = outRows(r, col)
        End If

End Function


Private Function CountResult( _
    ByVal outRows As Variant, _
    ByVal resultLabel As String) _
    As Long
'
'==============================================================================
' PURPOSE
'   Count the Reconciliation rows with a result.
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim r   As Long    'Row

'------------------------------------------------------------------------------
' COUNT
'------------------------------------------------------------------------------
        For r = 1 To RowCount(outRows)
            If SafeStr(outRows(r, RC_RESULT)) = resultLabel Then
                CountResult = CountResult + 1
            End If
        Next r

End Function


Private Function CountTotals( _
    ByVal wsResults As Worksheet) _
    As Long
'
'==============================================================================
' PURPOSE
'   Count the TOTAL rows on Results.
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim r   As Long    'Results row

'------------------------------------------------------------------------------
' COUNT
'------------------------------------------------------------------------------
        For r = OutputFirstRow(SH_RESULTS) To LastDataRow(wsResults, OutputFirstRow(SH_RESULTS), RS_ID)
            If UTxt(TableCell(wsResults, r, RS_ID).Value) = "TOTAL " & RG_CRR Or _
               UTxt(TableCell(wsResults, r, RS_ID).Value) = "TOTAL " & RG_BCBS Then
                CountTotals = CountTotals + 1
            End If
        Next r

End Function


Private Function ReconciliationRows( _
    ByVal wsRecon As Worksheet) _
    As Variant
'
'==============================================================================
' PURPOSE
'   Read the rows on the Reconciliation sheet; Empty when there are none.
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim topRow  As Long    'First data row
    Dim lastR   As Long    'Last row with a result

'------------------------------------------------------------------------------
' READ
'------------------------------------------------------------------------------
        topRow = OutputFirstRow(wsRecon.Name)
        lastR = LastDataRow(wsRecon, topRow, RC_RESULT)
        If lastR >= topRow Then
            ReconciliationRows = TableRange(wsRecon, topRow, 1, lastR, _
                                            IIf(wsRecon.Name = SH_TRADE_RECON, TRC_NCOLS, RC_NCOLS)).Value2
        End If

End Function


Private Function RowCount( _
    ByVal block As Variant) _
    As Long
'
'==============================================================================
' PURPOSE
'   Return the number of rows of a block; 0 when it is not an array.
'==============================================================================
'
        If IsArray(block) Then
            RowCount = UBound(block, 1) - LBound(block, 1) + 1
        End If

End Function


Private Sub CheckStatus( _
    ByVal wsRecon As Worksheet, _
    ByVal expectedStart As String, _
    ByVal moment As String)
'
'==============================================================================
' PURPOSE
'   Check that the status formula in Reconciliation C5 starts with
'   expectedStart, after Excel has recalculated what changed.
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim shown   As String    'Text in Reconciliation C5

'------------------------------------------------------------------------------
' CHECK
'------------------------------------------------------------------------------
        DoEvents
        Application.Calculate
        shown = SafeStr(wsRecon.Range(RECON_STATUS_CELL).Value)
        Check Left$(shown, Len(expectedStart)) = expectedStart, _
              "Reconciliation C5 " & moment & ": expected """ & expectedStart & "..."", found """ & shown & """"

End Sub


Private Sub CheckText( _
    ByVal what As String, _
    ByVal actual As String, _
    ByVal expected As String)
'
'==============================================================================
' PURPOSE
'   Check that a text equals the expected one.
'==============================================================================
'
        Check actual = expected, what & ": expected " & expected & ", found " & actual

End Sub


Private Sub CheckNear( _
    ByVal what As String, _
    ByVal actual As Variant, _
    ByVal expected As Double)
'
'==============================================================================
' PURPOSE
'   Check that a value is a number within TOL of the expected one.
'==============================================================================
'
        If IsError(actual) Or IsEmpty(actual) Or Not IsNumeric(actual) Then
            Check False, what & ": expected " & CStr(expected) & ", found no number"
        Else
            Check Abs(CDbl(actual) - expected) <= TOL, what & ": expected " & CStr(expected) & ", found " & _
                  CStr(actual)
        End If

End Sub


Private Sub BeginCase( _
    ByVal caseName As String)
'
'==============================================================================
' PURPOSE
'   Count a case and print its CASE line.
'==============================================================================
'
        mCaseCount = mCaseCount + 1
        mCurrentCase = caseName
        TEST_Suite.Say "CASE=" & caseName

End Sub


Private Sub Check( _
    ByVal passed As Boolean, _
    ByVal detail As String)
'
'==============================================================================
' PURPOSE
'   Count one check and record it when it fails.
'==============================================================================
'
        mCheckCount = mCheckCount + 1
        If Not passed Then
            Fail mCurrentCase, detail
        End If

End Sub


Private Sub Fail( _
    ByVal caseName As String, _
    ByVal detail As String)
'
'==============================================================================
' PURPOSE
'   Count a failure and print it.
'==============================================================================
'
        mFailureCount = mFailureCount + 1
        TEST_Suite.Say "FAILURE=" & caseName & ": " & detail

End Sub
