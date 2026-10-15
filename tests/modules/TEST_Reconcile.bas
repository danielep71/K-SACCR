Attribute VB_Name = "TEST_Reconcile"
'==============================================================================
' MODULE: TEST_Reconcile
'------------------------------------------------------------------------------
' PURPOSE
'   Check the reconciliation of a bank's netting-set figures with Results
'   (#45): figures within either tolerance MATCH and others are a
'   DIFFERENCE; bank amounts are scaled and the multiplier is not; IDs
'   missing on either side, listed twice, of another regime, withheld on
'   Results or unusable are each reported and never compared; TOTAL rows
'   are not netting sets; and, in the workbook, a run writes the
'   Reconciliation sheet, an edited bank figure makes it OUT OF DATE, and
'   validation withdraws it.
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
'   The workbook case writes bank figures on BankNettingSets, runs the
'   macros, which rewrite every output sheet, and puts the BankNettingSets
'   table and settings back, also after an unexpected error. It leaves the
'   outputs of a normal run. Use a development workbook.
'
' STATE OWNERSHIP
'   Restores the caller's calculation mode and M_Main.gSilent.
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
'   2026-10-15
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
        Private Const EXPECTED_CASES    As Long = 12              'Cases in a complete run
        Private Const EXPECTED_CHECKS   As Long = 68              'Checks in a complete run
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
'   2026-10-15
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim callerCalculation   As XlCalculation    'Caller's calculation mode
    Dim callerSilent        As Boolean          'Caller's M_Main.gSilent

'------------------------------------------------------------------------------
' RUN
'------------------------------------------------------------------------------
        mCaseCount = 0
        mCheckCount = 0
        mFailureCount = 0
        callerCalculation = Application.Calculation
        callerSilent = M_Main.gSilent
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

        Application.Calculation = callerCalculation
        M_Main.gSilent = callerSilent
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
'   Results are not netting sets; an empty bank file has no rows.
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
        bank = NewBlock(1, BN_NCOLS)
        SetBankRow bank, 1, " ns1 ", "crr", 100
        calc = NewBlock(3, RS_NCOLS)
        SetCalcRow calc, 1, "NS1", "CRR", "VALID", 100, 0, 0, 0
        SetCalcRow calc, 2, "TOTAL CRR", "CRR", "1 of 1 netting set(s) VALID", 100, 0, 0, Empty
        SetCalcRow calc, 3, "TOTAL BCBS", "BCBS", "0 of 0", 0, 0, 0, Empty
        found = Compare(bank, calc, 1#, 0#, 0#)
        CheckText "ID and regime in another case", ResultOf(found, "NS1", "EAD"), RES_MATCH
        Check RowCount(found) = 1, "TOTAL rows reported: " & RowCount(found)
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
'   status; an edited bank figure makes the status OUT OF DATE; validation
'   withdraws the reconciliation. The bank table and settings are put back.
'
' UPDATED
'   2026-10-15
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
        lastR = LastDataRow(wsResults, FIRST_DATA_ROW, RS_ID)
        For i = lastR To FIRST_DATA_ROW Step -1
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
    'rows from FIRST_DATA_ROW but the TOTAL rows and the one compared.
        Check CountResult(outcomes, RES_NOTINBANK) = LastDataRow(wsResults, FIRST_DATA_ROW, RS_ID) - FIRST_DATA_ROW - _
              CountTotals(wsResults), "NOT IN BANK FILE rows: " & CountResult(outcomes, RES_NOTINBANK)
        Check InStr(SafeStr(wsRecon.Range(RECON_INFO_CELL).Value), "Last reconciliation") = 1, _
              "summary: " & SafeStr(wsRecon.Range(RECON_INFO_CELL).Value)
        CheckStatus wsRecon, "Reconciliation current", "after the run"

'------------------------------------------------------------------------------
' EDITED BANK FIGURE, THEN VALIDATION
'------------------------------------------------------------------------------
        TableCell(wsBank, firstRow + 1, BN_EAD).Value = 2000
        CheckStatus wsRecon, "OUT OF DATE", "after a bank figure was edited"
        M_Main.ValidateInputs
        Check RowCount(ReconciliationRows(wsRecon)) = 0, "rows left after validation"
        CheckStatus wsRecon, "No reconciliation", "after validation"

'------------------------------------------------------------------------------
' RESTORE
'------------------------------------------------------------------------------
        RestoreBank wsBank, firstRow, saved, savedSet
        result = M_Main.RunSACCR_Silent()
        Check Left$(result, 9) = "RESULT=OK", "closing run: " & result
        Exit Sub

Unexpected:
        Fail mCurrentCase, "unexpected error " & Err.Number & ": " & Err.Description
        RestoreBank wsBank, firstRow, saved, savedSet

End Sub


Private Sub RestoreBank( _
    ByVal wsBank As Worksheet, _
    ByVal firstRow As Long, _
    ByVal saved As Variant, _
    ByVal savedSet As Variant)
'
'==============================================================================
' PURPOSE
'   Put the bank table and settings back as CaseWorkbook found them.
'
' ERROR POLICY
'   Contained: a failure is reported as a failure of the case.
'==============================================================================
'
        On Error GoTo Failed
        If wsBank Is Nothing Or Not IsArray(saved) Or Not IsArray(savedSet) Then
            Exit Sub
        End If
        TableRange(wsBank, firstRow, 1, firstRow + 49, BN_NCOLS).Formula = saved
        NamedCell(NM_BN_SCALE).Formula = savedSet(0)
        NamedCell(NM_BN_ABSTOL).Formula = savedSet(1)
        NamedCell(NM_BN_RELTOL).Formula = savedSet(2)
        Exit Sub

Failed:
        Fail mCurrentCase, "could not restore BankNettingSets: " & Err.Description

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
        For r = FIRST_DATA_ROW To LastDataRow(wsResults, FIRST_DATA_ROW, RS_ID)
            If Left$(UTxt(TableCell(wsResults, r, RS_ID).Value), 5) = "TOTAL" Then
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
    Dim lastR   As Long    'Last row with a result

'------------------------------------------------------------------------------
' READ
'------------------------------------------------------------------------------
        lastR = LastDataRow(wsRecon, FIRST_DATA_ROW, RC_RESULT)
        If lastR >= FIRST_DATA_ROW Then
            ReconciliationRows = TableRange(wsRecon, FIRST_DATA_ROW, 1, lastR, RC_NCOLS).Value2
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
