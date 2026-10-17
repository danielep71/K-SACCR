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
'   is the same for two exports of one run, is refused when the inputs,
'   the bank figures or an output cell changed after the run or a
'   recalculation is pending, and leaves an older report as it was when
'   writing or replacing it fails (#312). A failed reconciliation keeps
'   its cause in the run message and the summary (#318).
'
' PUBLIC SURFACE
'   RunReconcileTests is the entry point. Option Private Module keeps it out
'   of the external workbook automation API.
'
' DEPENDENCIES
'   CORE_Reconcile for the comparison cases and CORE_Export for
'   export.fields, which use synthetic arrays and values and no sheet;
'   M_Main, with its test seams gTestFault and gLastMessage, for the
'   workbook and export cases, which need a workbook built from the
'   template with the full source and its demonstration inputs; CORE_Util
'   and CORE_Config for sheet access and the column layout; TEST_Suite.Say
'   for the log lines. Every figure is synthetic.
'
' WORKSHEET SAFETY
'   The workbook cases write bank figures on BankNettingSets and
'   BankTrades, run the macros, which rewrite every output sheet, and put
'   the bank tables and settings back, also after an unexpected error. The
'   export cases write three files to the TEMP folder
'   (K-SACCR_evidence_test_a.csv, _b.csv and _c.csv) and delete them;
'   export.workbook also edits an input, a bank cell and an output cell and
'   adds a currency below the FX table, and puts them back. They leave the
'   outputs of a normal run. Use a development workbook.
'
' STATE OWNERSHIP
'   Restores the caller's calculation mode, M_Main.gSilent and
'   M_Main.gQuietLog. gQuietLog is set for the run: the macros' run messages
'   would push the report out of the Immediate window, which keeps about
'   200 lines. The cases that set M_Main.gTestFault set it back to "",
'   also after an unexpected error.
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
        Private Const EXPECTED_CASES    As Long = 20              'Cases in a complete run
        Private Const EXPECTED_CHECKS   As Long = 162              'Checks in a complete run
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
    Dim CallerCalculation   As XlCalculation    'Caller's calculation mode
    Dim CallerSilent        As Boolean          'Caller's M_Main.gSilent
    Dim CallerQuiet         As Boolean          'Caller's M_Main.gQuietLog

'------------------------------------------------------------------------------
' RUN
'------------------------------------------------------------------------------
        mCaseCount = 0
        mCheckCount = 0
        mFailureCount = 0
        CallerCalculation = Application.Calculation
        CallerSilent = M_Main.gSilent
        CallerQuiet = M_Main.gQuietLog
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
        CaseExportReplaceFailure

        Application.Calculation = CallerCalculation
        M_Main.gSilent = CallerSilent
        M_Main.gQuietLog = CallerQuiet
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
    Dim BankRows       As Variant    'Bank rows
    Dim CalcRows       As Variant    'Results rows
    Dim ComparedRows   As Variant    'Reconciliation rows

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
        On Error GoTo Unexpected
        BeginCase "compare.tolerances"
        BankRows = NewBlock(1, BN_NCOLS)
        SetBankRow BankRows, 1, "NS1", "CRR", 1000, 1000000, 500
        CalcRows = NewBlock(1, RS_NCOLS)
        SetCalcRow CalcRows, 1, "NS1", "CRR", "VALID", 1000.5, 1000050, 600, 0.5
        ComparedRows = Compare(BankRows, CalcRows, 1#, 1#, 0.0001)
        CheckText "EAD within the absolute tolerance", ResultOf(ComparedRows, "NS1", "EAD"), RES_MATCH
        CheckText "RC within the relative tolerance", ResultOf(ComparedRows, "NS1", "RC"), RES_MATCH
        CheckText "PFE outside both", ResultOf(ComparedRows, "NS1", "PFE"), RES_DIFF
        CheckNear "PFE difference", CellOf(ComparedRows, "NS1", "PFE", RC_DIFF), 100#
        CheckNear "PFE relative difference", CellOf(ComparedRows, "NS1", "PFE", RC_REL), 0.2
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
    Dim BankRows       As Variant    'Bank rows
    Dim CalcRows       As Variant    'Results rows
    Dim ComparedRows   As Variant    'Reconciliation rows

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
        On Error GoTo Unexpected
        BeginCase "compare.scale"
        BankRows = NewBlock(1, BN_NCOLS)
        SetBankRow BankRows, 1, "NS1", "BCBS", 48.6, 12.002
        CalcRows = NewBlock(1, RS_NCOLS)
        SetCalcRow CalcRows, 1, "NS1", "BCBS", "VALID", 48600, 12000, 0, 0
        ComparedRows = Compare(BankRows, CalcRows, 1000#, 1#, 0#)
        CheckNear "EAD bank figure in units", CellOf(ComparedRows, "NS1", "EAD", RC_BANK), 48600#
        CheckText "EAD in thousands", ResultOf(ComparedRows, "NS1", "EAD"), RES_MATCH
        CheckNear "RC difference in units", CellOf(ComparedRows, "NS1", "RC", RC_DIFF), -2#
        CheckText "RC two units apart", ResultOf(ComparedRows, "NS1", "RC"), RES_DIFF
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
    Dim BankRows       As Variant    'Bank rows
    Dim CalcRows       As Variant    'Results rows
    Dim ComparedRows   As Variant    'Reconciliation rows

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
        On Error GoTo Unexpected
        BeginCase "compare.multiplier"
        BankRows = NewBlock(2, BN_NCOLS)
        SetBankRow BankRows, 1, "NS1", "CRR"
        BankRows(1, BN_MULT) = 0.5
        SetBankRow BankRows, 2, "NS2", "CRR"
        BankRows(2, BN_MULT) = 0.9
        CalcRows = NewBlock(2, RS_NCOLS)
        SetCalcRow CalcRows, 1, "NS1", "CRR", "VALID", 1, 1, 1, 0.5
        SetCalcRow CalcRows, 2, "NS2", "CRR", "VALID", 1, 1, 1, 0.8
        ComparedRows = Compare(BankRows, CalcRows, 1000#, 1#, 0.0001)
        CheckNear "multiplier not scaled", CellOf(ComparedRows, "NS1", "Multiplier", RC_BANK), 0.5
        CheckText "equal multiplier", ResultOf(ComparedRows, "NS1", "Multiplier"), RES_MATCH
        CheckText "absolute tolerance not applied", ResultOf(ComparedRows, "NS2", "Multiplier"), RES_DIFF
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
    Dim BankRows       As Variant    'Bank rows
    Dim CalcRows       As Variant    'Results rows
    Dim ComparedRows   As Variant    'Reconciliation rows

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
        On Error GoTo Unexpected
        BeginCase "compare.zero-bank-figure"
        BankRows = NewBlock(1, BN_NCOLS)
        SetBankRow BankRows, 1, "NS1", "CRR", Empty, 0, 0
        CalcRows = NewBlock(1, RS_NCOLS)
        SetCalcRow CalcRows, 1, "NS1", "CRR", "VALID", 10, 0.5, 5, 0.5
        ComparedRows = Compare(BankRows, CalcRows, 1#, 1#, 0.5)
        CheckText "zero within the absolute tolerance", ResultOf(ComparedRows, "NS1", "RC"), RES_MATCH
        CheckText "zero outside it", ResultOf(ComparedRows, "NS1", "PFE"), RES_DIFF
        Check IsEmpty(CellOf(ComparedRows, "NS1", "PFE", RC_REL)), "relative difference of a zero bank figure is not empty"
        Check RowOf(ComparedRows, "NS1", "EAD") = 0, "a blank bank figure is compared"
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
    Dim BankRows       As Variant    'Bank rows
    Dim CalcRows       As Variant    'Results rows
    Dim ComparedRows   As Variant    'Reconciliation rows

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
        On Error GoTo Unexpected
        BeginCase "compare.missing-and-not-in-bank"
        BankRows = NewBlock(2, BN_NCOLS)
        SetBankRow BankRows, 1, "NS1", "CRR", 100
        SetBankRow BankRows, 2, "NS9", "CRR", 900
        CalcRows = NewBlock(2, RS_NCOLS)
        SetCalcRow CalcRows, 1, "NS1", "CRR", "VALID", 100, 0, 0, 0
        SetCalcRow CalcRows, 2, "NS2", "BCBS", "VALID", 250, 0, 0, 0
        ComparedRows = Compare(BankRows, CalcRows, 1#, 0#, 0#)
        CheckText "bank ID not on Results", ResultOf(ComparedRows, "NS9", "EAD"), RES_MISSING
        CheckNear "its bank EAD", CellOf(ComparedRows, "NS9", "EAD", RC_BANK), 900#
        CheckText "Results netting set not in the bank file", ResultOf(ComparedRows, "NS2", "EAD"), RES_NOTINBANK
        CheckNear "its Results EAD", CellOf(ComparedRows, "NS2", "EAD", RC_CALC), 250#
        Check IsEmpty(CellOf(ComparedRows, "NS2", "EAD", RC_BANKROW)), "a netting set not in the bank file has a bank row"
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
    Dim BankRows       As Variant    'Bank rows
    Dim CalcRows       As Variant    'Results rows
    Dim ComparedRows   As Variant    'Reconciliation rows

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
        On Error GoTo Unexpected
        BeginCase "compare.duplicate"
        BankRows = NewBlock(2, BN_NCOLS)
        SetBankRow BankRows, 1, "NS1", "CRR", 100
        SetBankRow BankRows, 2, "ns1", "CRR", 100
        CalcRows = NewBlock(1, RS_NCOLS)
        SetCalcRow CalcRows, 1, "NS1", "CRR", "VALID", 100, 0, 0, 0
        ComparedRows = Compare(BankRows, CalcRows, 1#, 0#, 0#)
        Check RowCount(ComparedRows) = 2, "rows: " & RowCount(ComparedRows)
        CheckText "first row", CStr(ComparedRows(1, RC_RESULT)), RES_DUPLICATE
        CheckText "second row", CStr(ComparedRows(2, RC_RESULT)), RES_DUPLICATE
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
    Dim BankRows       As Variant    'Bank rows
    Dim CalcRows       As Variant    'Results rows
    Dim ComparedRows   As Variant    'Reconciliation rows

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
        On Error GoTo Unexpected
        BeginCase "compare.regime-and-withheld"
        BankRows = NewBlock(2, BN_NCOLS)
        SetBankRow BankRows, 1, "NS1", "BCBS", 100, 10
        SetBankRow BankRows, 2, "NS2", "CRR", 200, 20
        CalcRows = NewBlock(2, RS_NCOLS)
        SetCalcRow CalcRows, 1, "NS1", "CRR", "VALID", 100, 10, 0, 0
        SetCalcRow CalcRows, 2, "NS2", "CRR", "INCOMPLETE: 1 of 3 trade(s) rejected", Empty, Empty, Empty, Empty
        ComparedRows = Compare(BankRows, CalcRows, 1#, 0#, 0#)
        Check RowCount(ComparedRows) = 2, "rows: " & RowCount(ComparedRows)
        CheckText "other regime", ResultOf(ComparedRows, "NS1", ""), RES_REGIME
        CheckText "withheld netting set", ResultOf(ComparedRows, "NS2", ""), RES_WITHHELD
        CheckText "its Results status", CStr(CellOf(ComparedRows, "NS2", "", RC_CALCSTATUS)), _
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
    Dim BankRows       As Variant    'Bank rows
    Dim CalcRows       As Variant    'Results rows
    Dim ComparedRows   As Variant    'Reconciliation rows

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
        On Error GoTo Unexpected
        BeginCase "compare.invalid-input"
        BankRows = NewBlock(5, BN_NCOLS)
        SetBankRow BankRows, 1, "", "CRR", 100
        SetBankRow BankRows, 2, "NS2", "", 100
        SetBankRow BankRows, 3, "NS3", "CRR"
        BankRows(3, BN_COMMENT) = "comment only, no figure"
        SetBankRow BankRows, 4, "NS4", "CRR", "1,000", CVErr(xlErrNA), True, 50
        BankRows(5, BN_COMMENT) = "a comment alone is not a row"
        CalcRows = NewBlock(3, RS_NCOLS)
        SetCalcRow CalcRows, 1, "NS2", "CRR", "VALID", 100, 0, 0, 0
        SetCalcRow CalcRows, 2, "NS3", "CRR", "VALID", 100, 0, 0, 0
        SetCalcRow CalcRows, 3, "NS4", "CRR", "VALID", 1000, 0, 50, 50
        ComparedRows = Compare(BankRows, CalcRows, 1#, 0#, 0#)
        CheckText "row without ID", ResultOf(ComparedRows, "", ""), RES_INVALID
        CheckText "row without regime", ResultOf(ComparedRows, "NS2", ""), RES_INVALID
        CheckText "row without figures", ResultOf(ComparedRows, "NS3", ""), RES_INVALID
        CheckText "number as text", ResultOf(ComparedRows, "NS4", "EAD"), RES_INVALID
        CheckText "error value", ResultOf(ComparedRows, "NS4", "RC"), RES_INVALID
        CheckText "logical value", ResultOf(ComparedRows, "NS4", "PFE"), RES_INVALID
        CheckText "usable figure on the same row", ResultOf(ComparedRows, "NS4", "AddOn aggregate"), RES_MATCH
        Check mCounts(RCN_INVALID) = 6 And RowCount(ComparedRows) = 7, _
              "invalid count " & mCounts(RCN_INVALID) & ", rows " & RowCount(ComparedRows)
        Check CLng(CellOf(ComparedRows, "NS4", "EAD", RC_BANKROW)) = 4, "bank row of NS4: " & CellOf(ComparedRows, "NS4", "EAD", RC_BANKROW)
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
    Dim BankRows       As Variant    'Bank rows
    Dim CalcRows       As Variant    'Results rows
    Dim ComparedRows   As Variant    'Reconciliation rows

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
        On Error GoTo Unexpected
        BeginCase "compare.ids-and-totals"
        BankRows = NewBlock(2, BN_NCOLS)
        SetBankRow BankRows, 1, " ns1 ", "crr", 100
        SetBankRow BankRows, 2, "TOTAL DESK", "CRR", 50
        CalcRows = NewBlock(4, RS_NCOLS)
        SetCalcRow CalcRows, 1, "NS1", "CRR", "VALID", 100, 0, 0, 0
        SetCalcRow CalcRows, 2, "TOTAL DESK", "CRR", "VALID", 50, 0, 0, 0
        SetCalcRow CalcRows, 3, "TOTAL CRR", "CRR", "2 of 2 netting set(s) VALID", 150, 0, 0, Empty
        SetCalcRow CalcRows, 4, "TOTAL BCBS", "BCBS", "0 of 0", 0, 0, 0, Empty
        ComparedRows = Compare(BankRows, CalcRows, 1#, 0#, 0#)
        CheckText "ID and regime in another case", ResultOf(ComparedRows, "NS1", "EAD"), RES_MATCH
        CheckText "netting set named TOTAL DESK", ResultOf(ComparedRows, "TOTAL DESK", "EAD"), RES_MATCH
        Check RowCount(ComparedRows) = 2, "TOTAL rows reported: " & RowCount(ComparedRows)
        BankRows = NewBlock(2, BN_NCOLS)
        BankRows(2, BN_COMMENT) = "comment"
        Check Not CORE_Reconcile.HasBankRows(BankRows), "a table with a comment only has bank rows"
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
    Dim Headers           As Variant    'Header row, (1, column)
    Dim ExpectedHeaders   As Variant    'Expected headers
    Dim c                 As Long       'Column

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
        ExpectedHeaders = Split(BN_HEADERS, "|")
        Headers = NewBlock(1, BN_NCOLS)
        For c = 0 To UBound(ExpectedHeaders)
            Headers(1, c + 1) = UCase$(Replace(ExpectedHeaders(c), " ", "  "))
        Next c
        Check CORE_Reconcile.HeadersProblem(Headers, BN_HEADERS) = "", "headers in another case rejected"
        Headers(1, 4) = ExpectedHeaders(4)
        Headers(1, 5) = ExpectedHeaders(3)
        Check InStr(CORE_Reconcile.HeadersProblem(Headers, BN_HEADERS), "column 4") > 0, "swapped columns accepted"
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
    Dim Counts()   As Long      'Counts
    Dim Summary    As String    'Summary line
    Dim i          As Long      'Count

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
        On Error GoTo Unexpected
        BeginCase "summary"
        ReDim Counts(1 To RCN_COUNT)
        For i = 1 To RCN_COUNT
            Counts(i) = 10 + i
        Next i
        Summary = CORE_Reconcile.SummaryText(Counts, 1000#, 1#, 0.0001)
        Check InStr(Summary, "bank netting sets 11") > 0 And InStr(Summary, "figures compared 12") > 0 And _
              InStr(Summary, "MATCH 13") > 0 And InStr(Summary, "DIFFERENCE 14") > 0, "compared counts: " & Summary
        Check InStr(Summary, "missing in calculator 15") > 0 And InStr(Summary, "not in bank file 16") > 0 And _
              InStr(Summary, "duplicate rows 17") > 0 And InStr(Summary, "regime mismatch 18") > 0 And _
              InStr(Summary, "withheld 19") > 0 And InStr(Summary, "invalid input 20") > 0, "other counts: " & Summary
        Check InStr(Summary, "scale 1000") > 0 And InStr(Summary, "tolerance 1 or ") > 0, "settings: " & Summary
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
    Dim BankSheet       As Worksheet    'BankNettingSets
    Dim ReconSheet      As Worksheet    'Reconciliation
    Dim ResultsSheet    As Worksheet    'Results
    Dim FirstRow        As Long         'First bank row
    Dim SavedTable      As Variant      'Bank table formulas, first 50 rows
    Dim SavedSettings   As Variant      'Settings formulas, (scale, abs, rel)
    Dim ResultLine      As String       'Result line of a run
    Dim ValidRow        As Long         'Results row of the first VALID netting set
    Dim ValidSetId      As String       'Its ID
    Dim LastRow         As Long         'Last Results row
    Dim ReconRows       As Variant      'Reconciliation rows
    Dim i               As Long         'Row

'------------------------------------------------------------------------------
' SAVE THE BANK TABLE
'------------------------------------------------------------------------------
        On Error GoTo Unexpected
        BeginCase "workbook.run-edit-validate"
        Application.Calculation = xlCalculationAutomatic
        M_Main.gSilent = True
        Set BankSheet = GetSheet(SH_BANK_NS)
        Set ReconSheet = GetSheet(SH_RECON)
        Set ResultsSheet = GetSheet(SH_RESULTS)
        FirstRow = NamedRow(NM_BN_HEADER) + 1
        SavedTable = TableRange(BankSheet, FirstRow, 1, FirstRow + 49, BN_NCOLS).Formula
        SavedSettings = Array(NamedCell(NM_BN_SCALE).Formula, NamedCell(NM_BN_ABSTOL).Formula, _
                         NamedCell(NM_BN_RELTOL).Formula)

'------------------------------------------------------------------------------
' NO BANK FIGURES
'------------------------------------------------------------------------------
        TableRange(BankSheet, FirstRow, 1, FirstRow + 49, BN_NCOLS).ClearContents
        ResultLine = M_Main.RunSACCR_Silent()
        Check InStr(ResultLine, "; reconciliation=NONE;") > 0, "run without bank figures: " & ResultLine
        Check InStr(SafeStr(ReconSheet.Range(RECON_INFO_CELL).Value), "No bank figures") = 1, _
              "summary without bank figures: " & SafeStr(ReconSheet.Range(RECON_INFO_CELL).Value)

'------------------------------------------------------------------------------
' FIGURES TAKEN FROM RESULTS
'------------------------------------------------------------------------------
    'The first VALID netting set: its EAD as on Results (MATCH) and its RC
    'ten percent higher (DIFFERENCE); and an ID that Results does not have.
        LastRow = LastDataRow(ResultsSheet, OutputFirstRow(SH_RESULTS), RS_ID)
        For i = LastRow To OutputFirstRow(SH_RESULTS) Step -1
            If SafeStr(TableCell(ResultsSheet, i, RS_STATUS_COL).Value) = "VALID" Then
                ValidRow = i
            End If
        Next i
        Check ValidRow > 0, "no VALID netting set on Results: build the workbook from the template"
        ValidSetId = SafeStr(TableCell(ResultsSheet, ValidRow, RS_ID).Value)
        NamedCell(NM_BN_SCALE).Value = 1
        NamedCell(NM_BN_ABSTOL).Value = 1
        NamedCell(NM_BN_RELTOL).Value = 0.0001
        TableCell(BankSheet, FirstRow, BN_ID).Value = ValidSetId
        TableCell(BankSheet, FirstRow, BN_REGIME).Value = TableCell(ResultsSheet, ValidRow, RS_REGIME).Value
        TableCell(BankSheet, FirstRow, BN_EAD).Value = TableCell(ResultsSheet, ValidRow, RS_EAD).Value2
        TableCell(BankSheet, FirstRow, BN_EAD + 1).Value = TableCell(ResultsSheet, ValidRow, RS_RC).Value2 * 1.1 + 100
        TableCell(BankSheet, FirstRow + 1, BN_ID).Value = "NOT-ON-RESULTS"
        TableCell(BankSheet, FirstRow + 1, BN_REGIME).Value = TableCell(ResultsSheet, ValidRow, RS_REGIME).Value
        TableCell(BankSheet, FirstRow + 1, BN_EAD).Value = 1000
        ResultLine = M_Main.RunSACCR_Silent()
        Check InStr(ResultLine, "; reconciliation=DONE;") > 0, "run with bank figures: " & ResultLine
        ReconRows = ReconciliationRows(ReconSheet)
        CheckText "EAD as on Results", ResultOf(ReconRows, ValidSetId, "EAD"), RES_MATCH
        CheckText "RC ten percent higher", ResultOf(ReconRows, ValidSetId, "RC"), RES_DIFF
        CheckText "ID not on Results", ResultOf(ReconRows, "NOT-ON-RESULTS", "EAD"), RES_MISSING
    'Every other netting set on Results is not in the bank file: all the
    'rows from the first data row but the TOTAL rows and the one compared.
        Check CountResult(ReconRows, RES_NOTINBANK) = LastDataRow(ResultsSheet, OutputFirstRow(SH_RESULTS), RS_ID) - OutputFirstRow(SH_RESULTS) - _
              CountTotals(ResultsSheet), "NOT IN BANK FILE rows: " & CountResult(ReconRows, RES_NOTINBANK)
        Check InStr(SafeStr(ReconSheet.Range(RECON_INFO_CELL).Value), "Last reconciliation") = 1, _
              "summary: " & SafeStr(ReconSheet.Range(RECON_INFO_CELL).Value)
        CheckStatus ReconSheet, "Reconciliation current", "after the run"

'------------------------------------------------------------------------------
' FAILED RECONCILIATION
'------------------------------------------------------------------------------
    'Cleaning up a failed reconciliation clears Err: its cause must reach
    'both the summary and the run message (#318). A normal run follows.
        M_Main.gTestFault = "reconcile"
        M_Main.gLastMessage = ""
        ResultLine = M_Main.RunSACCR_Silent()
        M_Main.gTestFault = ""
        Check InStr(ResultLine, "; reconciliation=FAILED;") > 0, "run with a failed reconciliation: " & ResultLine
        Check InStr(M_Main.gLastMessage, "Reconciliation failed: Injected reconciliation failure.") > 0, _
              "run message after a failed reconciliation: " & M_Main.gLastMessage
        Check InStr(SafeStr(ReconSheet.Range(RECON_INFO_CELL).Value), "Injected reconciliation failure.") > 0, _
              "summary after a failed reconciliation: " & SafeStr(ReconSheet.Range(RECON_INFO_CELL).Value)
        ResultLine = M_Main.RunSACCR_Silent()
        Check InStr(ResultLine, "; reconciliation=DONE;") > 0, "run after the failed reconciliation: " & ResultLine

'------------------------------------------------------------------------------
' EDITED BANK FIGURE, STOPPED RUN, THEN VALIDATION
'------------------------------------------------------------------------------
    'A run that stops in its recalculation raises in silent mode; the
    'reconciliation must be gone all the same.
        TableCell(BankSheet, FirstRow + 1, BN_EAD).Value = 2000
        CheckStatus ReconSheet, "OUT OF DATE", "after a bank figure was edited"
        M_Main.gTestFault = "recalc"
        On Error Resume Next
        ResultLine = M_Main.RunSACCR_Silent()
        Err.Clear
        On Error GoTo Unexpected
        M_Main.gTestFault = ""
        Check RowCount(ReconciliationRows(ReconSheet)) = 0, "rows left after a run whose recalculation did not finish"
        ResultLine = M_Main.RunSACCR_Silent()
        Check RowCount(ReconciliationRows(ReconSheet)) > 0, "no rows after a new run: " & ResultLine
        M_Main.ValidateInputs
        Check RowCount(ReconciliationRows(ReconSheet)) = 0, "rows left after validation"
        CheckStatus ReconSheet, "No reconciliation", "after validation"

'------------------------------------------------------------------------------
' RESTORE
'------------------------------------------------------------------------------
        RestoreBankTable BankSheet, FirstRow, BN_NCOLS, SavedTable, SavedSettings, NM_BN_SCALE, NM_BN_ABSTOL, NM_BN_RELTOL
        ResultLine = M_Main.RunSACCR_Silent()
        Check Left$(ResultLine, 9) = "RESULT=OK", "closing run: " & ResultLine
        Exit Sub

Unexpected:
        Fail mCurrentCase, "unexpected error " & Err.Number & ": " & Err.Description
        M_Main.gTestFault = ""
        RestoreBankTable BankSheet, FirstRow, BN_NCOLS, SavedTable, SavedSettings, NM_BN_SCALE, NM_BN_ABSTOL, NM_BN_RELTOL

End Sub


Private Sub RestoreBankTable( _
    ByVal BankSheet As Worksheet, _
    ByVal FirstRow As Long, _
    ByVal ColumnCount As Long, _
    ByVal SavedTable As Variant, _
    ByVal SavedSettings As Variant, _
    ByVal ScaleName As String, _
    ByVal AbsToleranceName As String, _
    ByVal RelToleranceName As String)
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
        If BankSheet Is Nothing Or Not IsArray(SavedTable) Or Not IsArray(SavedSettings) Then
            Exit Sub
        End If
        TableRange(BankSheet, FirstRow, 1, FirstRow + 49, ColumnCount).Formula = SavedTable
        NamedCell(ScaleName).Formula = SavedSettings(0)
        NamedCell(AbsToleranceName).Formula = SavedSettings(1)
        NamedCell(RelToleranceName).Formula = SavedSettings(2)
        Exit Sub

Failed:
        Fail mCurrentCase, "could not restore " & BankSheet.Name & ": " & Err.Description

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
    Dim BankRows       As Variant    'Bank rows
    Dim CalcRows       As Variant    'TradeCalc rows
    Dim ComparedRows   As Variant    'Reconciliation rows

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
        On Error GoTo Unexpected
        BeginCase "trade.figures"
        BankRows = NewBlock(1, BT_NCOLS)
        BankRows(1, BT_ID) = "T1"
        BankRows(1, BT_NOTIONAL) = 10
        BankRows(1, BT_NOTIONAL + 5) = 0.5
        BankRows(1, BT_NOTIONAL + 6) = 1
        BankRows(1, BT_NOTIONAL + 7) = 0.3
        CalcRows = NewBlock(1, TC_NCOLS)
        SetTradeCalcRow CalcRows, 1, "T1", "NS1", "IR", "OK", 10000.4, 0.6, 1, Empty, ""
        ComparedRows = CompareT(BankRows, CalcRows, 1000#, 1#, 0.0001)
        CheckText "notional in thousands", ResultOf(ComparedRows, "T1", "Notional"), RES_MATCH
        CheckText "delta outside the relative tolerance", ResultOf(ComparedRows, "T1", "Supervisory delta"), RES_DIFF
        CheckNear "delta not scaled", CellOf(ComparedRows, "T1", "Supervisory delta", RC_BANK), 0.5
        CheckText "equal maturity factor", ResultOf(ComparedRows, "T1", "MF unmargined"), RES_MATCH
        CheckText "no margined maturity factor", ResultOf(ComparedRows, "T1", "MF margined"), RES_NA
        CheckText "netting set shown", CStr(CellOf(ComparedRows, "T1", "Notional", TRC_NS)), "NS1"
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
    Dim BankRows       As Variant    'Bank rows
    Dim CalcRows       As Variant    'TradeCalc rows
    Dim ComparedRows   As Variant    'Reconciliation rows

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
        On Error GoTo Unexpected
        BeginCase "trade.populations"
        BankRows = NewBlock(5, BT_NCOLS)
        BankRows(1, BT_ID) = "T2"
        BankRows(1, BT_NOTIONAL) = 100
        BankRows(2, BT_ID) = "T4"
        BankRows(2, BT_NOTIONAL) = 100
        BankRows(3, BT_ID) = "T5"
        BankRows(3, BT_NOTIONAL) = 100
        BankRows(4, BT_ID) = "T5"
        BankRows(4, BT_NOTIONAL) = 100
        BankRows(5, BT_NOTIONAL) = 100
        CalcRows = NewBlock(3, TC_NCOLS)
        SetTradeCalcRow CalcRows, 1, "T3", "NS1", "FX", "OK", 300, 1, 1, Empty, ""
        SetTradeCalcRow CalcRows, 2, "T4", "NS1", "IR", "EXCLUDED", Empty, Empty, Empty, Empty, "Maturity date missing"
        SetTradeCalcRow CalcRows, 3, "T5", "NS1", "IR", "OK", 100, 1, 1, Empty, ""
        ComparedRows = CompareT(BankRows, CalcRows, 1#, 0#, 0#)
        CheckText "bank leg not on TradeCalc", ResultOf(ComparedRows, "T2", "Notional"), RES_MISSING
        CheckText "TradeCalc leg not in the bank file", ResultOf(ComparedRows, "T3", "Notional"), RES_NOTINBANK
        CheckNear "its notional", CellOf(ComparedRows, "T3", "Notional", RC_CALC), 300#
        CheckText "excluded leg", ResultOf(ComparedRows, "T4", ""), RES_EXCLUDED
        Check InStr(SafeStr(CellOf(ComparedRows, "T4", "", RC_NOTE)), "Maturity date missing") > 0, _
              "reasons of the excluded leg: " & SafeStr(CellOf(ComparedRows, "T4", "", RC_NOTE))
        CheckText "leg listed twice", ResultOf(ComparedRows, "T5", ""), RES_DUPLICATE
        CheckText "row without TradeID", ResultOf(ComparedRows, "", ""), RES_INVALID
        Check mCounts(RCN_DUPLICATE) = 2 And mCounts(RCN_EXCLUDED) = 1 And RowCount(ComparedRows) = 6, _
              "duplicate/excluded/rows: " & mCounts(RCN_DUPLICATE) & "/" & mCounts(RCN_EXCLUDED) & "/" & RowCount(ComparedRows)
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
    Dim BankRows       As Variant    'Bank rows
    Dim CalcRows       As Variant    'TradeCalc rows
    Dim ComparedRows   As Variant    'Reconciliation rows

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
        On Error GoTo Unexpected
        BeginCase "trade.calc-duplicates"
        BankRows = NewBlock(1, BT_NCOLS)
        BankRows(1, BT_ID) = "T7"
        BankRows(1, BT_NOTIONAL) = 100
        CalcRows = NewBlock(4, TC_NCOLS)
        SetTradeCalcRow CalcRows, 1, "T7", "NS1", "IR", "OK", 100, 1, 1, Empty, ""
        SetTradeCalcRow CalcRows, 2, "T7", "NS1", "IR", "EXCLUDED", Empty, Empty, Empty, Empty, "duplicate Trade ID"
        SetTradeCalcRow CalcRows, 3, "T8", "NS1", "FX", "OK", 50, 1, 1, Empty, ""
        SetTradeCalcRow CalcRows, 4, "T8", "NS1", "FX", "EXCLUDED", Empty, Empty, Empty, Empty, "duplicate Trade ID"
        ComparedRows = CompareT(BankRows, CalcRows, 1#, 0#, 0#)
        CheckText "bank leg listed twice on TradeCalc", ResultOf(ComparedRows, "T7", ""), RES_DUPLICATE
        CheckText "leg listed twice on TradeCalc, not in the bank file", ResultOf(ComparedRows, "T8", ""), RES_DUPLICATE
        Check mCounts(RCN_DUPLICATE) = 2 And mCounts(RCN_NOTINBANK) = 0 And mCounts(RCN_COMPARED) = 0 And _
              RowCount(ComparedRows) = 2, "duplicate/not in bank/compared/rows: " & mCounts(RCN_DUPLICATE) & "/" & _
              mCounts(RCN_NOTINBANK) & "/" & mCounts(RCN_COMPARED) & "/" & RowCount(ComparedRows)
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
    Dim Counts()   As Long      'Counts
    Dim Summary    As String    'Summary line
    Dim i          As Long      'Count

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
        On Error GoTo Unexpected
        BeginCase "trade.summary"
        ReDim Counts(1 To RCN_COUNT)
        For i = 1 To RCN_COUNT
            Counts(i) = 10 + i
        Next i
        Summary = CORE_Reconcile.TradeSummaryText(Counts, 1#, 1#, 0#)
        Check InStr(Summary, "bank trade legs 11") > 0 And InStr(Summary, "excluded 21") > 0 And _
              InStr(Summary, "not applicable 22") > 0 And InStr(Summary, "not in bank file 16") > 0, "summary: " & Summary
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
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim BankSheet         As Worksheet    'BankTrades
    Dim ReconSheet        As Worksheet    'TradeReconciliation
    Dim CalcSheet         As Worksheet    'TradeCalc
    Dim FirstRow          As Long         'First bank row
    Dim SavedTable        As Variant      'Bank table formulas, first 50 rows
    Dim SavedSettings     As Variant      'Settings formulas, (scale, abs, rel)
    Dim ResultLine        As String       'Result line of a run
    Dim OkRow             As Long         'TradeCalc row of the first OK leg
    Dim i                 As Long         'TradeCalc row, then Trades row
    Dim LegId             As String       'Its TradeID
    Dim ReconRows         As Variant      'TradeReconciliation rows
    Dim TradesSheet       As Worksheet    'Trades
    Dim TradeRow          As Long         'Trades row of the leg
    Dim EconomicIdCell    As Range        'Its Economic trade ID, set and restored
    Dim MtmLegCell        As Range        'Its MtM leg flag, set and restored
    Dim SavedEconomicId   As Variant      'Formula of EconomicIdCell
    Dim SavedMtmLeg       As Variant      'Formula of MtmLegCell

'------------------------------------------------------------------------------
' SAVE THE BANK TABLE
'------------------------------------------------------------------------------
        On Error GoTo Unexpected
        BeginCase "trade.workbook-run-edit"
        Application.Calculation = xlCalculationAutomatic
        M_Main.gSilent = True
        Set BankSheet = GetSheet(SH_BANK_TR)
        Set ReconSheet = GetSheet(SH_TRADE_RECON)
        Set CalcSheet = GetSheet(SH_TRADECALC)
        FirstRow = NamedRow(NM_BT_HEADER) + 1
        SavedTable = TableRange(BankSheet, FirstRow, 1, FirstRow + 49, BT_NCOLS).Formula
        SavedSettings = Array(NamedCell(NM_BT_SCALE).Formula, NamedCell(NM_BT_ABSTOL).Formula, _
                         NamedCell(NM_BT_RELTOL).Formula)

'------------------------------------------------------------------------------
' FIGURES TAKEN FROM TRADECALC
'------------------------------------------------------------------------------
        TableRange(BankSheet, FirstRow, 1, FirstRow + 49, BT_NCOLS).ClearContents
        ResultLine = M_Main.RunSACCR_Silent()
        For i = LastDataRow(CalcSheet, OutputFirstRow(SH_TRADECALC), TC_ID) To OutputFirstRow(SH_TRADECALC) Step -1
            If SafeStr(TableCell(CalcSheet, i, TC_STATUS).Value) = "OK" Then
                OkRow = i
            End If
        Next i
        Check OkRow > 0, "no OK leg on TradeCalc: build the workbook from the template"
        LegId = SafeStr(TableCell(CalcSheet, OkRow, TC_ID).Value)
        NamedCell(NM_BT_SCALE).Value = 1
        NamedCell(NM_BT_ABSTOL).Value = 1
        NamedCell(NM_BT_RELTOL).Value = 0.0001
        TableCell(BankSheet, FirstRow, BT_ID).Value = LegId
        TableCell(BankSheet, FirstRow, BT_NOTIONAL).Value = TableCell(CalcSheet, OkRow, TC_NOTIONAL).Value2
        TableCell(BankSheet, FirstRow, BT_NOTIONAL + 5).Value = TableCell(CalcSheet, OkRow, TC_DELTA).Value2 + 0.1
        TableCell(BankSheet, FirstRow + 1, BT_ID).Value = "NO-SUCH-LEG"
        TableCell(BankSheet, FirstRow + 1, BT_NOTIONAL).Value = 1000
        ResultLine = M_Main.RunSACCR_Silent()
        Check InStr(ResultLine, "; trade_reconciliation=DONE;") > 0, "run with bank trade figures: " & ResultLine
        ReconRows = ReconciliationRows(ReconSheet)
        CheckText "notional as on TradeCalc", ResultOf(ReconRows, LegId, "Notional"), RES_MATCH
        CheckText "delta 0.1 higher", ResultOf(ReconRows, LegId, "Supervisory delta"), RES_DIFF
        CheckText "leg not on TradeCalc", ResultOf(ReconRows, "NO-SUCH-LEG", "Notional"), RES_MISSING
        CheckStatus ReconSheet, "Reconciliation current", "after the run"
        TableCell(BankSheet, FirstRow + 1, BT_NOTIONAL).Value = 2000
        CheckStatus ReconSheet, "OUT OF DATE", "after a bank figure was edited"

'------------------------------------------------------------------------------
' ECONOMIC TRADE ID
'------------------------------------------------------------------------------
    'The leg made an economic trade of one leg, which must then be its MtM
    'leg; its rows show the ID, a leg not on Trades shows none.
        Set TradesSheet = GetSheet(SH_TRADES)
        For i = TableFirstRow(SH_TRADES) To LastDataRow(TradesSheet, TableFirstRow(SH_TRADES), TR_ID)
            If UTxt(TableCell(TradesSheet, i, TR_ID).Value) = UTxt(LegId) Then
                TradeRow = i
            End If
        Next i
        Check TradeRow > 0, "leg " & LegId & " not found on Trades"
        If TradeRow > 0 Then
            Set EconomicIdCell = TableCell(TradesSheet, TradeRow, TR_ECON)
            Set MtmLegCell = TableCell(TradesSheet, TradeRow, TR_MTMLEG)
            SavedEconomicId = EconomicIdCell.Formula
            SavedMtmLeg = MtmLegCell.Formula
            EconomicIdCell.Value = "ECON-T1"
            MtmLegCell.Value = "Y"
            ResultLine = M_Main.RunSACCR_Silent()
            ReconRows = ReconciliationRows(ReconSheet)
            CheckText "Economic trade ID of the leg", SafeStr(CellOf(ReconRows, LegId, "Notional", TRC_ECON)), "ECON-T1"
            CheckText "Economic trade ID of a leg not on Trades", _
                      SafeStr(CellOf(ReconRows, "NO-SUCH-LEG", "Notional", TRC_ECON)), ""
            EconomicIdCell.Formula = SavedEconomicId
            MtmLegCell.Formula = SavedMtmLeg
            Check UBound(ColumnValues(TradesSheet, TradeRow, TradeRow, TR_ID), 1) = 1, _
                  "a one-row Trades column is not read as an array"
        End If

'------------------------------------------------------------------------------
' RESTORE
'------------------------------------------------------------------------------
        RestoreBankTable BankSheet, FirstRow, BT_NCOLS, SavedTable, SavedSettings, NM_BT_SCALE, NM_BT_ABSTOL, NM_BT_RELTOL
        ResultLine = M_Main.RunSACCR_Silent()
        Check Left$(ResultLine, 9) = "RESULT=OK", "closing run: " & ResultLine
        Exit Sub

Unexpected:
        Fail mCurrentCase, "unexpected error " & Err.Number & ": " & Err.Description
        RestoreCell EconomicIdCell, SavedEconomicId
        RestoreCell MtmLegCell, SavedMtmLeg
        RestoreBankTable BankSheet, FirstRow, BT_NCOLS, SavedTable, SavedSettings, NM_BT_SCALE, NM_BT_ABSTOL, NM_BT_RELTOL

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
    Dim Bytes()   As Byte    'UTF-8 bytes

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
        Bytes = CORE_Export.Utf8Bytes("A" & ChrW(&HE9) & ChrW(&H20AC))
        Check UBound(Bytes) = 5 And Bytes(0) = &H41 And Bytes(1) = &HC3 And Bytes(2) = &HA9 And Bytes(3) = &HE2 And Bytes(4) = &H82 And _
              Bytes(5) = &HAC, "UTF-8 of A, e acute and the euro sign"
        Bytes = CORE_Export.Utf8Bytes(ChrW(&HD83D) & ChrW(&HDE00))
        Check UBound(Bytes) = 3 And Bytes(0) = &HF0 And Bytes(1) = &H9F And Bytes(2) = &H98 And Bytes(3) = &H80, _
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
'   byte, also over a longer hidden file of that name; an edited input,
'   bank figures entered after the run, bank figures removed after a
'   reconciliation and an edited output cell each make the export refuse,
'   as does a recalculation pending in manual calculation; undoing them,
'   or a new run, allows it again. A currency added below the FX table's
'   named range is in the report. The edited cells and the files are put
'   back or removed.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim TempFolder    As String       'Folder for the test files
    Dim FirstFile     As String       'First export
    Dim SecondFile    As String       'Second export
    Dim ResultLine    As String       'Result line
    Dim FirstText     As String       'First export as text
    Dim MtmCell       As Range        'An input edited and restored
    Dim SavedMtm      As Variant      'Its formula
    Dim BankCell      As Range        'A bank figure entered and removed
    Dim SavedBank     As Variant      'Its formula
    Dim OutputCell    As Range        'An output edited and restored
    Dim SavedOutput   As Variant      'Its value
    Dim ParamsSheet   As Worksheet    'Params
    Dim FxColumn      As Long         'Key column of the FX table
    Dim FxRow         As Long         'First blank row below it
    Dim FxCells       As Range        'A currency added and removed
    Dim SavedFx       As Variant      'Their formulas
    Dim FileNumber    As Integer      'File number of a file written by the case

'------------------------------------------------------------------------------
' EXPORT TWICE
'------------------------------------------------------------------------------
        On Error GoTo Unexpected
        BeginCase "export.workbook"
        Application.Calculation = xlCalculationAutomatic
        M_Main.gSilent = True
        TempFolder = Environ$("TEMP")
        If Len(TempFolder) = 0 Then
            TempFolder = ThisWorkbook.Path
        End If
        FirstFile = TempFolder & Application.PathSeparator & "K-SACCR_evidence_test_a.csv"
        SecondFile = TempFolder & Application.PathSeparator & "K-SACCR_evidence_test_b.csv"
        ResultLine = M_Main.RunSACCR_Silent()
        ResultLine = M_Main.ExportEvidence_Silent(FirstFile)
        Check Left$(ResultLine, 9) = "RESULT=OK", "first export: " & ResultLine
        FirstText = FileText(FirstFile)
        Check Left$(FirstText, 31) = "RUN;format;K-SACCR evidence 1" & vbCrLf, "first line: " & Left$(FirstText, 40)
        Check InStr(FirstText, vbCrLf & "RUN;results_status;CURRENT" & vbCrLf) > 0, "results status line missing"
        Check InStr(FirstText, vbCrLf & "RUN;reference_parameter_set;") > 0 And InStr(FirstText, "RUN;workbook;") = 0, _
              "run identity lines"
        Check InStr(FirstText, vbCrLf & "RUN;reconciliation;NOT MADE" & vbCrLf) > 0 And _
              InStr(FirstText, vbCrLf & "RUN;reconciliation_scale;") = 0, "settings of a reconciliation that was not made"
        Check Not (SafeStr(GetSheet(SH_RESULTS).Range(RUNINFO_CELL).Value2) Like "*#,# s*"), _
              "run time with a decimal comma: " & SafeStr(GetSheet(SH_RESULTS).Range(RUNINFO_CELL).Value2)
        Check InStr(FirstText, vbCrLf & "PARAMS;Alpha;1.4" & vbCrLf) > 0, "Alpha line missing or not invariant"
        Check InStr(FirstText, vbCrLf & "TRADE_RECONCILIATION;TradeID;") > 0 And _
              InStr(FirstText, ";Note;Economic trade ID" & vbCrLf) > 0, "trade reconciliation header line"
        Check InStr(FirstText, vbCrLf & "RESULTS;NettingSetID;") > 0 And InStr(FirstText, vbCrLf & "CHECKS;") > 0 And _
              InStr(FirstText, vbCrLf & "LIMITATIONS;1;") > 0, "a section is missing"
        ResultLine = M_Main.ExportEvidence_Silent(SecondFile)
        Check FileText(SecondFile) = FirstText, "a second export of the same run differs"
        FileNumber = FreeFile
        Open SecondFile For Output As #FileNumber
        Print #FileNumber, String$(Len(FirstText) + 100, "x")
        Close #FileNumber
        SetAttr SecondFile, vbHidden
        ResultLine = M_Main.ExportEvidence_Silent(SecondFile)
        Check Left$(ResultLine, 9) = "RESULT=OK" And FileText(SecondFile) = FirstText, _
              "a longer hidden file is not replaced whole: " & ResultLine

'------------------------------------------------------------------------------
' REFUSE WHEN NOT CURRENT
'------------------------------------------------------------------------------
        Set MtmCell = TableCell(GetSheet(SH_TRADES), TableFirstRow(SH_TRADES), TR_MTM)
        SavedMtm = MtmCell.Formula
        MtmCell.Value = ToDbl(MtmCell.Value, 0#) + 1#
        ResultLine = M_Main.ExportEvidence_Silent(SecondFile)
        Check Left$(ResultLine, 14) = "RESULT=REFUSED", "export after an input edit: " & ResultLine
        MtmCell.Formula = SavedMtm
        ResultLine = M_Main.ExportEvidence_Silent(SecondFile)
        Check Left$(ResultLine, 9) = "RESULT=OK", "export after undoing the edit: " & ResultLine
    'In manual calculation the same value written again leaves the status
    'formulas to recalculate: the export waits for F9.
        Application.Calculation = xlCalculationManual
        MtmCell.Formula = SavedMtm
        ResultLine = M_Main.ExportEvidence_Silent(SecondFile)
        Check Left$(ResultLine, 14) = "RESULT=REFUSED" And InStr(ResultLine, "recalculating") > 0, _
              "export with a recalculation pending: " & ResultLine
        Application.Calculate
        Application.Calculation = xlCalculationAutomatic
        ResultLine = M_Main.ExportEvidence_Silent(SecondFile)
        Check Left$(ResultLine, 9) = "RESULT=OK", "export after the recalculation: " & ResultLine
        Set BankCell = TableCell(GetSheet(SH_BANK_NS), NamedRow(NM_BN_HEADER) + 1, BN_ID)
        SavedBank = BankCell.Formula
        BankCell.Value = "NEW-AFTER-RUN"
        ResultLine = M_Main.ExportEvidence_Silent(SecondFile)
        Check Left$(ResultLine, 14) = "RESULT=REFUSED" And InStr(ResultLine, SH_BANK_NS) > 0, _
              "export with bank figures entered after the run: " & ResultLine
        ResultLine = M_Main.RunSACCR_Silent()
        ResultLine = M_Main.ExportEvidence_Silent(SecondFile)
        Check Left$(ResultLine, 9) = "RESULT=OK", "export with a reconciliation: " & ResultLine
        Check InStr(FileText(SecondFile), vbCrLf & "RUN;reconciliation_scale;") > 0, "settings of the reconciliation missing"
        BankCell.Formula = SavedBank
        ResultLine = M_Main.ExportEvidence_Silent(SecondFile)
        Check Left$(ResultLine, 14) = "RESULT=REFUSED" And InStr(ResultLine, SH_BANK_NS) > 0, _
              "export after every bank figure was removed: " & ResultLine

'------------------------------------------------------------------------------
' REFUSE AN EDITED OUTPUT
'------------------------------------------------------------------------------
        ResultLine = M_Main.RunSACCR_Silent()
        Set OutputCell = TableCell(GetSheet(SH_RESULTS), OutputFirstRow(SH_RESULTS), RS_EAD)
        SavedOutput = OutputCell.Value2
        OutputCell.Value2 = ToDbl(SavedOutput, 0#) + 1#
        ResultLine = M_Main.ExportEvidence_Silent(SecondFile)
        Check Left$(ResultLine, 14) = "RESULT=REFUSED" And InStr(ResultLine, "output cell") > 0, _
              "export after an output cell was edited: " & ResultLine
        OutputCell.Value2 = SavedOutput
        ResultLine = M_Main.ExportEvidence_Silent(SecondFile)
        Check Left$(ResultLine, 9) = "RESULT=OK", "export after restoring the output cell: " & ResultLine

'------------------------------------------------------------------------------
' EXPORT A CURRENCY ADDED BELOW THE FX TABLE
'------------------------------------------------------------------------------
    'The template asks for new currencies directly below the last row,
    'outside the range of the name FX_Table; the run reads them, so the
    'report must list them.
        Set ParamsSheet = GetSheet(SH_PARAMS)
        FxColumn = NamedTableCol(NM_FX_TABLE)
        FxRow = NamedRow(NM_FX_TABLE) + 1
        Do While Not IsBlankCell(TableCell(ParamsSheet, FxRow, FxColumn).Value)
            FxRow = FxRow + 1
        Loop
        Set FxCells = TableRange(ParamsSheet, FxRow, FxColumn, FxRow, FxColumn + 2)
        SavedFx = FxCells.Formula
        FxCells.Value = Array("ZZX", "Test currency", 2)
        ResultLine = M_Main.RunSACCR_Silent()
        ResultLine = M_Main.ExportEvidence_Silent(SecondFile)
        Check InStr(FileText(SecondFile), vbCrLf & "FX_TABLE;ZZX;Test currency;2" & vbCrLf) > 0, _
              "a currency added below the FX table is missing from the report: " & ResultLine
        FxCells.Formula = SavedFx
        ResultLine = M_Main.RunSACCR_Silent()

'------------------------------------------------------------------------------
' CLEAN UP
'------------------------------------------------------------------------------
        RemoveFile FirstFile
        RemoveFile SecondFile
        Exit Sub

Unexpected:
        Fail mCurrentCase, "unexpected error " & Err.Number & ": " & Err.Description
        Application.Calculation = xlCalculationAutomatic
        RestoreCell MtmCell, SavedMtm
        RestoreCell BankCell, SavedBank
        If Not OutputCell Is Nothing Then
            OutputCell.Value2 = SavedOutput
        End If
        If Not FxCells Is Nothing Then
            FxCells.Formula = SavedFx
        End If
        RemoveFile FirstFile
        RemoveFile SecondFile

End Sub


Private Sub CaseExportReplaceFailure()
'
'==============================================================================
' PURPOSE
'   A report that cannot be written or cannot replace the older one leaves
'   the older file as it was, byte for byte and still hidden, and no file
'   beside it: with a failure injected before the new file is opened,
'   while it is written and while the older file is replaced, and with a
'   read-only older file (#312). Without a failure the report replaces
'   the older file.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim TempFolder    As String     'Folder for the test file
    Dim TargetFile    As String     'The report's file
    Dim OlderText     As String     'Content of the older file
    Dim ResultLine    As String     'Result line
    Dim Faults        As Variant    'Injected failures
    Dim Fault         As Variant    'One of them
    Dim ErrorNumber   As Long       'Error raised by the export
    Dim FileNumber    As Integer    'File number of the older file

'------------------------------------------------------------------------------
' WRITE AN OLDER FILE
'------------------------------------------------------------------------------
        On Error GoTo Unexpected
        BeginCase "export.replace-failure"
        M_Main.gSilent = True
        TempFolder = Environ$("TEMP")
        If Len(TempFolder) = 0 Then
            TempFolder = ThisWorkbook.Path
        End If
        TargetFile = TempFolder & Application.PathSeparator & "K-SACCR_evidence_test_c.csv"
        RemoveFile TargetFile
        OlderText = "An older report" & vbCrLf
        FileNumber = FreeFile
        Open TargetFile For Output As #FileNumber
        Print #FileNumber, "An older report"
        Close #FileNumber
        SetAttr TargetFile, vbHidden
        ResultLine = M_Main.RunSACCR_Silent()

'------------------------------------------------------------------------------
' FAIL AND KEEP THE OLDER FILE
'------------------------------------------------------------------------------
        Faults = Array("evidence-open", "evidence-write", "evidence-replace")
        For Each Fault In Faults
            M_Main.gTestFault = CStr(Fault)
            ErrorNumber = 0
            On Error Resume Next
            ResultLine = M_Main.ExportEvidence_Silent(TargetFile)
            ErrorNumber = Err.Number
            On Error GoTo Unexpected
            M_Main.gTestFault = ""
            Check ErrorNumber = ERR_INJECTED_FAULT, Fault & ": error " & ErrorNumber
            Check FileText(TargetFile) = OlderText, Fault & ": the older file changed"
            Check (GetAttr(TargetFile) And vbHidden) <> 0, Fault & ": the older file is no longer hidden"
            Check Len(Dir$(TargetFile & ".saccr-*", vbHidden Or vbSystem)) = 0, Fault & ": a file was left beside it"
        Next Fault
        SetAttr TargetFile, vbReadOnly
        ErrorNumber = 0
        On Error Resume Next
        ResultLine = M_Main.ExportEvidence_Silent(TargetFile)
        ErrorNumber = Err.Number
        On Error GoTo Unexpected
        SetAttr TargetFile, vbNormal
        Check ErrorNumber = ERR_EXPORT_WRITE, "read-only file: error " & ErrorNumber
        Check FileText(TargetFile) = OlderText, "read-only file: it changed"
        Check Len(Dir$(TargetFile & ".saccr-*", vbHidden Or vbSystem)) = 0, "read-only file: a file was left beside it"

'------------------------------------------------------------------------------
' REPLACE IT WITHOUT A FAILURE
'------------------------------------------------------------------------------
        ResultLine = M_Main.ExportEvidence_Silent(TargetFile)
        Check Left$(ResultLine, 9) = "RESULT=OK", "export without a failure: " & ResultLine
        Check Left$(FileText(TargetFile), 29) = "RUN;format;K-SACCR evidence 1", "the report did not replace the older file"
        Check Len(Dir$(TargetFile & ".saccr-*", vbHidden Or vbSystem)) = 0, "a file was left beside the report"

'------------------------------------------------------------------------------
' CLEAN UP
'------------------------------------------------------------------------------
        RemoveFile TargetFile
        Exit Sub

Unexpected:
        Fail mCurrentCase, "unexpected error " & Err.Number & ": " & Err.Description
        M_Main.gTestFault = ""
        On Error Resume Next
        SetAttr TargetFile, vbNormal
        RemoveFile TargetFile

End Sub


Private Function FileText( _
    ByVal FilePath As String) _
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
    Dim FileNumber   As Integer    'File number
    Dim Bytes()      As Byte       'The bytes

'------------------------------------------------------------------------------
' READ
'------------------------------------------------------------------------------
        FileNumber = FreeFile
        Open FilePath For Binary Access Read As #FileNumber
        If LOF(FileNumber) > 0 Then
            ReDim Bytes(0 To LOF(FileNumber) - 1)
            Get #FileNumber, , Bytes
            FileText = StrConv(Bytes, vbUnicode)
        End If
        Close #FileNumber

End Function


Private Sub RemoveFile( _
    ByVal FilePath As String)
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
        If Len(FilePath) > 0 Then
            If Len(Dir$(FilePath, vbHidden)) > 0 Then
                SetAttr FilePath, vbNormal
                Kill FilePath
            End If
        End If
        Err.Clear

End Sub


Private Sub RestoreCell( _
    ByVal EditedCell As Range, _
    ByVal SavedFormula As Variant)
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
        If Not EditedCell Is Nothing Then
            EditedCell.Formula = SavedFormula
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
    ByVal NumberOfRows As Long, _
    ByVal NumberOfColumns As Long) _
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
    Dim Block()   As Variant    'The block

'------------------------------------------------------------------------------
' CREATE
'------------------------------------------------------------------------------
        ReDim Block(1 To NumberOfRows, 1 To NumberOfColumns)
        NewBlock = Block

End Function


Private Sub SetBankRow( _
    ByRef BankRows As Variant, _
    ByVal RowIndex As Long, _
    ByVal RowId As String, _
    ByVal RegimeCode As String, _
    Optional ByVal EadValue As Variant, _
    Optional ByVal RcValue As Variant, _
    Optional ByVal PfeValue As Variant, _
    Optional ByVal AddOnValue As Variant)
'
'==============================================================================
' PURPOSE
'   Fill a bank row: ID, regime and, when given, EAD, RC, PFE and aggregate
'   add-on. A string "" leaves the cell empty.
'==============================================================================
'
        If Len(RowId) > 0 Then BankRows(RowIndex, BN_ID) = RowId
        If Len(RegimeCode) > 0 Then BankRows(RowIndex, BN_REGIME) = RegimeCode
        If Not IsMissing(EadValue) Then BankRows(RowIndex, BN_EAD) = EadValue
        If Not IsMissing(RcValue) Then BankRows(RowIndex, BN_EAD + 1) = RcValue
        If Not IsMissing(PfeValue) Then BankRows(RowIndex, BN_EAD + 2) = PfeValue
        If Not IsMissing(AddOnValue) Then BankRows(RowIndex, BN_EAD + 3) = AddOnValue

End Sub


Private Sub SetCalcRow( _
    ByRef CalcRows As Variant, _
    ByVal RowIndex As Long, _
    ByVal RowId As String, _
    ByVal RegimeCode As String, _
    ByVal StatusText As String, _
    ByVal EadValue As Variant, _
    ByVal RcValue As Variant, _
    ByVal PfeValue As Variant, _
    ByVal Multiplier As Variant)
'
'==============================================================================
' PURPOSE
'   Fill a Results row: ID, regime, status, EAD, RC, PFE and multiplier;
'   the aggregate add-on equals the PFE.
'==============================================================================
'
        CalcRows(RowIndex, RS_ID) = RowId
        CalcRows(RowIndex, RS_REGIME) = RegimeCode
        CalcRows(RowIndex, RS_STATUS_COL) = StatusText
        CalcRows(RowIndex, RS_EAD) = EadValue
        CalcRows(RowIndex, RS_RC) = RcValue
        CalcRows(RowIndex, RS_PFE) = PfeValue
        CalcRows(RowIndex, RS_ADDON) = PfeValue
        CalcRows(RowIndex, RS_MULT) = Multiplier

End Sub


Private Function Compare( _
    ByVal BankRows As Variant, _
    ByVal CalcRows As Variant, _
    ByVal ScaleFactor As Double, _
    ByVal AbsTolerance As Double, _
    ByVal RelTolerance As Double) _
    As Variant
'
'==============================================================================
' PURPOSE
'   Run CORE_Reconcile.CompareNettingSets with the bank rows starting on sheet
'   row 1, and keep its counts in mCounts.
'==============================================================================
'
        Compare = CORE_Reconcile.CompareNettingSets(BankRows, 1, CalcRows, ScaleFactor, AbsTolerance, RelTolerance, mCounts)

End Function


Private Function CompareT( _
    ByVal BankRows As Variant, _
    ByVal CalcRows As Variant, _
    ByVal ScaleFactor As Double, _
    ByVal AbsTolerance As Double, _
    ByVal RelTolerance As Double) _
    As Variant
'
'==============================================================================
' PURPOSE
'   Run CORE_Reconcile.CompareTrades with the bank rows starting on sheet
'   row 1, and keep its counts in mCounts.
'==============================================================================
'
        CompareT = CORE_Reconcile.CompareTrades(BankRows, 1, CalcRows, ScaleFactor, AbsTolerance, RelTolerance, mCounts)

End Function


Private Sub SetTradeCalcRow( _
    ByRef CalcRows As Variant, _
    ByVal RowIndex As Long, _
    ByVal RowId As String, _
    ByVal NettingSetCode As String, _
    ByVal ClassCode As String, _
    ByVal StatusText As String, _
    ByVal Notional As Variant, _
    ByVal DeltaValue As Variant, _
    ByVal UnmarginedMf As Variant, _
    ByVal MarginedMf As Variant, _
    ByVal Messages As String)
'
'==============================================================================
' PURPOSE
'   Fill a TradeCalc row: ID, netting set, asset class, status, notional,
'   delta, the two maturity factors and the messages.
'==============================================================================
'
        CalcRows(RowIndex, TC_ID) = RowId
        CalcRows(RowIndex, TC_NS) = NettingSetCode
        CalcRows(RowIndex, TC_AC) = ClassCode
        CalcRows(RowIndex, TC_STATUS) = StatusText
        CalcRows(RowIndex, TC_NOTIONAL) = Notional
        CalcRows(RowIndex, TC_DELTA) = DeltaValue
        CalcRows(RowIndex, TC_MFU) = UnmarginedMf
        CalcRows(RowIndex, TC_MFM) = MarginedMf
        CalcRows(RowIndex, TC_MESSAGES) = Messages

End Sub


Private Function RowOf( _
    ByVal ReconRows As Variant, _
    ByVal RowId As String, _
    ByVal Measure As String) _
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
        For r = 1 To RowCount(ReconRows)
            If SafeStr(ReconRows(r, RC_ID)) = RowId And SafeStr(ReconRows(r, RC_MEASURE)) = Measure Then
                RowOf = r
                Exit Function
            End If
        Next r

End Function


Private Function ResultOf( _
    ByVal ReconRows As Variant, _
    ByVal RowId As String, _
    ByVal Measure As String) _
    As String
'
'==============================================================================
' PURPOSE
'   Return the result of an ID and measure; "(no row)" when there is none.
'==============================================================================
'
        ResultOf = SafeStr(CellOf(ReconRows, RowId, Measure, RC_RESULT))
        If RowOf(ReconRows, RowId, Measure) = 0 Then
            ResultOf = "(no row)"
        End If

End Function


Private Function CellOf( _
    ByVal ReconRows As Variant, _
    ByVal RowId As String, _
    ByVal Measure As String, _
    ByVal ColumnIndex As Long) _
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
    Dim FoundRow   As Long    'Row

'------------------------------------------------------------------------------
' READ
'------------------------------------------------------------------------------
        FoundRow = RowOf(ReconRows, RowId, Measure)
        If FoundRow > 0 Then
            CellOf = ReconRows(FoundRow, ColumnIndex)
        End If

End Function


Private Function CountResult( _
    ByVal ReconRows As Variant, _
    ByVal ResultLabel As String) _
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
        For r = 1 To RowCount(ReconRows)
            If SafeStr(ReconRows(r, RC_RESULT)) = ResultLabel Then
                CountResult = CountResult + 1
            End If
        Next r

End Function


Private Function CountTotals( _
    ByVal ResultsSheet As Worksheet) _
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
        For r = OutputFirstRow(SH_RESULTS) To LastDataRow(ResultsSheet, OutputFirstRow(SH_RESULTS), RS_ID)
            If UTxt(TableCell(ResultsSheet, r, RS_ID).Value) = "TOTAL " & RG_CRR Or _
               UTxt(TableCell(ResultsSheet, r, RS_ID).Value) = "TOTAL " & RG_BCBS Then
                CountTotals = CountTotals + 1
            End If
        Next r

End Function


Private Function ReconciliationRows( _
    ByVal ReconSheet As Worksheet) _
    As Variant
'
'==============================================================================
' PURPOSE
'   Read the rows on a reconciliation sheet, Reconciliation or
'   TradeReconciliation; Empty when there are none.
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim FirstRow   As Long    'First data row
    Dim LastRow    As Long    'Last row with a result

'------------------------------------------------------------------------------
' READ
'------------------------------------------------------------------------------
        FirstRow = OutputFirstRow(ReconSheet.Name)
        LastRow = LastDataRow(ReconSheet, FirstRow, RC_RESULT)
        If LastRow >= FirstRow Then
            ReconciliationRows = TableRange(ReconSheet, FirstRow, 1, LastRow, _
                                            IIf(ReconSheet.Name = SH_TRADE_RECON, TRC_NCOLS, RC_NCOLS)).Value2
        End If

End Function


Private Function RowCount( _
    ByVal Block As Variant) _
    As Long
'
'==============================================================================
' PURPOSE
'   Return the number of rows of a block; 0 when it is not an array.
'==============================================================================
'
        If IsArray(Block) Then
            RowCount = UBound(Block, 1) - LBound(Block, 1) + 1
        End If

End Function


Private Sub CheckStatus( _
    ByVal ReconSheet As Worksheet, _
    ByVal ExpectedStart As String, _
    ByVal Moment As String)
'
'==============================================================================
' PURPOSE
'   Check that the status formula in C5 of a reconciliation sheet,
'   Reconciliation or TradeReconciliation, starts with ExpectedStart, after
'   Excel has recalculated what changed.
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Shown   As String    'Text in Reconciliation C5

'------------------------------------------------------------------------------
' CHECK
'------------------------------------------------------------------------------
        DoEvents
        Application.Calculate
        Shown = SafeStr(ReconSheet.Range(RECON_STATUS_CELL).Value)
        Check Left$(Shown, Len(ExpectedStart)) = ExpectedStart, _
              "Reconciliation C5 " & Moment & ": expected """ & ExpectedStart & "..."", found """ & Shown & """"

End Sub


Private Sub CheckText( _
    ByVal CheckLabel As String, _
    ByVal Actual As String, _
    ByVal Expected As String)
'
'==============================================================================
' PURPOSE
'   Check that a text equals the expected one.
'==============================================================================
'
        Check Actual = Expected, CheckLabel & ": expected " & Expected & ", found " & Actual

End Sub


Private Sub CheckNear( _
    ByVal CheckLabel As String, _
    ByVal Actual As Variant, _
    ByVal Expected As Double)
'
'==============================================================================
' PURPOSE
'   Check that a value is a number within TOL of the expected one.
'==============================================================================
'
        If IsError(Actual) Or IsEmpty(Actual) Or Not IsNumeric(Actual) Then
            Check False, CheckLabel & ": expected " & CStr(Expected) & ", found no number"
        Else
            Check Abs(CDbl(Actual) - Expected) <= TOL, CheckLabel & ": expected " & CStr(Expected) & ", found " & _
                  CStr(Actual)
        End If

End Sub


Private Sub BeginCase( _
    ByVal CaseName As String)
'
'==============================================================================
' PURPOSE
'   Count a case and print its CASE line.
'==============================================================================
'
        mCaseCount = mCaseCount + 1
        mCurrentCase = CaseName
        TEST_Suite.Say "CASE=" & CaseName

End Sub


Private Sub Check( _
    ByVal Passed As Boolean, _
    ByVal Detail As String)
'
'==============================================================================
' PURPOSE
'   Count one check and record it when it fails.
'==============================================================================
'
        mCheckCount = mCheckCount + 1
        If Not Passed Then
            Fail mCurrentCase, Detail
        End If

End Sub


Private Sub Fail( _
    ByVal CaseName As String, _
    ByVal Detail As String)
'
'==============================================================================
' PURPOSE
'   Count a failure and print it.
'==============================================================================
'
        mFailureCount = mFailureCount + 1
        TEST_Suite.Say "FAILURE=" & CaseName & ": " & Detail

End Sub
