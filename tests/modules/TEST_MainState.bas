Attribute VB_Name = "TEST_MainState"
'==============================================================================
' MODULE: TEST_MainState
'------------------------------------------------------------------------------
' PURPOSE
'   Check that the workbook macros in M_Main put Excel back exactly as they
'   found it, on success and on failure, report operation and cleanup
'   failures separately, and leave the workbook ready for the next run
'   (issue #43); that a Checks sheet that cannot be written, or any error
'   after the outputs are being written, fails the run and withdraws its
'   results; and that results are reported as out of date once an input
'   changes (issue #36), also on Params after its tables have moved
'   (issue #152).
'
' PUBLIC SURFACE
'   RunMainStateTests is the entry point. Option Private Module keeps it out
'   of the external workbook automation API.
'
' DEPENDENCIES
'   M_Main and its test seam gTestFault, the CORE_Engine test seam
'   gEngineFault, and the error numbers in CORE_Config. The workbook must be
'   built from the template with the full source, because the macros
'   calculate and write the output sheets.
'
' WORKSHEET SAFETY
'   Unlike TEST_Harness, this module runs the real macros: they clear and
'   rewrite TradeCalc, Buckets, HedgingSets, Results and Checks, and
'   activate the Checks sheet. One case protects the Checks sheet for one
'   run and unprotects it again, also after an unexpected error. One case
'   inserts a column before the Params tables for its run and deletes it
'   again, also after an unexpected error. The last case leaves the
'   outputs of a normal run. Use a development workbook only.
'
' STATE OWNERSHIP
'   Sets calculation mode, events and screen updating to known values for
'   each case and restores the caller's values, gSilent, gTestFault and
'   gEngineFault at the end, also after an unexpected error.
'
' ERROR POLICY
'   Expected errors are captured and asserted. An unexpected error is
'   recorded against its case and the remaining cases still run.
'
' USAGE
'   Import with the production modules, compile, then run
'   TEST_MainState.RunMainStateTests from the Immediate window.
'
' UPDATED
'   2026-10-12
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
        Private Const EXPECTED_CASES   As Long = 20   'Cases in a complete run

'------------------------------------------------------------------------------
' MODULE STATE
'------------------------------------------------------------------------------
        Private mCaseCount      As Long      'Cases started
        Private mCheckCount     As Long      'Checks evaluated
        Private mFailureCount   As Long      'Failed checks and unexpected errors
        Private mCurrentCase    As String    'Case being run, for messages
        Private mParamsRunLine  As String    'Result line of the unshifted Params run


'
'------------------------------------------------------------------------------
'
'                                 ENTRY POINT
'
'------------------------------------------------------------------------------
'

Public Sub RunMainStateTests()
'
'==============================================================================
'                              RunMainStateTests
'------------------------------------------------------------------------------
' PURPOSE
'   Run every case and print one CASE line per case, each failed check,
'   and a final RESULT line.
'
' STATE OWNERSHIP
'   Restores the caller's calculation mode, events, screen updating,
'   gSilent and gTestFault before printing the result.
'
' UPDATED
'   2026-10-06
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim callerCalculation   As XlCalculation    'Caller's calculation mode
    Dim callerEvents        As Boolean          'Caller's EnableEvents
    Dim callerScreen        As Boolean          'Caller's ScreenUpdating
    Dim callerSilent        As Boolean          'Caller's M_Main.gSilent
    Dim restored            As Boolean          'Caller's settings read back equal

'------------------------------------------------------------------------------
' START
'------------------------------------------------------------------------------
        mCaseCount = 0
        mCheckCount = 0
        mFailureCount = 0
        callerCalculation = Application.Calculation
        callerEvents = Application.EnableEvents
        callerScreen = Application.ScreenUpdating
        callerSilent = M_Main.gSilent
        Debug.Print "SACCR MAIN STATE TESTS"

'------------------------------------------------------------------------------
' RUN CASES
'------------------------------------------------------------------------------
    'Silent mode keeps message boxes from blocking the run and makes the
    'macros raise their errors.
        M_Main.gSilent = True
        M_Main.gTestFault = ""
        CaseRestoresState "clear.automatic-on", "clear", xlCalculationAutomatic, True, True
        CaseRestoresState "clear.manual-off", "clear", xlCalculationManual, False, False
        CaseRestoresState "validate.manual-off", "validate", xlCalculationManual, False, False
        CaseRestoresState "run.manual-off", "run", xlCalculationManual, False, False
        CaseOperationFailure
        CaseCleanupFailure
        CaseChecksUnwritable
        CaseOutputWriteFailure
        CaseOutputCleanupFailure
        CaseProtectedOutput SH_TRADECALC
        CaseProtectedOutput SH_BUCKETS
        CaseProtectedOutput SH_HEDGING
        CaseProtectedOutput SH_RESULTS
        CaseProtectedOutput SH_TRADECALC, SH_HEDGING
        CaseProtectedOutput SH_BUCKETS, SH_RESULTS
        CaseProtectedOutput SH_TRADECALC, SH_RESULTS, True
        CaseStaleResults
        CaseStaleParams False
        CaseStaleParams True
        CaseSilentFlagPreserved

'------------------------------------------------------------------------------
' RESTORE THE CALLER AND REPORT
'------------------------------------------------------------------------------
        M_Main.gTestFault = ""
        CORE_Engine.gEngineFault = ""
        M_Main.gSilent = callerSilent
        Application.Calculation = callerCalculation
        Application.EnableEvents = callerEvents
        Application.ScreenUpdating = callerScreen
        restored = (Application.Calculation = callerCalculation) And _
                   (Application.EnableEvents = callerEvents) And _
                   (Application.ScreenUpdating = callerScreen)
        If mCaseCount <> EXPECTED_CASES Then
            Fail "suite", "expected " & EXPECTED_CASES & " cases, ran " & mCaseCount
        End If
        Debug.Print "CASES=" & mCaseCount & "; CHECKS=" & mCheckCount & "; FAILURES=" & mFailureCount
        Debug.Print "RESULT=" & IIf(mFailureCount = 0 And restored, "PASS", "FAIL") & _
                    "; cases=" & mCaseCount & "; checks=" & mCheckCount & _
                    "; failures=" & mFailureCount & "; caller_state=" & IIf(restored, "RESTORED", "CHANGED")

End Sub


'
'------------------------------------------------------------------------------
'
'                                    CASES
'
'------------------------------------------------------------------------------
'

Private Sub CaseRestoresState( _
    ByVal caseName As String, _
    ByVal operation As String, _
    ByVal calcMode As XlCalculation, _
    ByVal eventsOn As Boolean, _
    ByVal screenOn As Boolean)
'
'==============================================================================
'                              CaseRestoresState
'------------------------------------------------------------------------------
' PURPOSE
'   Set known settings, run one operation, and check that every setting is
'   back to its value, including settings that were off before the run.
'
' INPUTS
'   caseName: case identifier.
'   operation: "clear", "validate" or "run".
'   calcMode, eventsOn, screenOn: the settings before the operation.
'
' UPDATED
'   2026-10-06
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim result   As String    'Result line of a run

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
        On Error GoTo Unexpected
        BeginCase caseName
        SetState calcMode, eventsOn, screenOn
        Select Case operation
            Case "clear"
                M_Main.ClearOutputs
            Case "validate"
                M_Main.ValidateInputs
            Case "run"
                result = M_Main.RunSACCR_Silent()
                Check Left$(result, 9) = "RESULT=OK", "run result: " & result
        End Select
        CheckState calcMode, eventsOn, screenOn
        Exit Sub

'------------------------------------------------------------------------------
' HANDLE UNEXPECTED ERROR
'------------------------------------------------------------------------------
Unexpected:
        Fail caseName, "unexpected error " & Err.Number & ": " & Err.Description

End Sub


Private Sub CaseOperationFailure()
'
'==============================================================================
'                             CaseOperationFailure
'------------------------------------------------------------------------------
' PURPOSE
'   An error during the operation is raised unchanged after Excel has been
'   restored, and the next run works.
'
' UPDATED
'   2026-10-06
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim result      As String    'Result line of the follow-up run
    Dim errNumber   As Long      'Error raised by the failed run
    Dim errText     As String    'Its description

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
        On Error GoTo Unexpected
        BeginCase "run.operation-failure"
        SetState xlCalculationAutomatic, False, False
        M_Main.gTestFault = "operation"
        On Error Resume Next
        result = M_Main.RunSACCR_Silent()
        errNumber = Err.Number
        errText = Err.Description
        On Error GoTo Unexpected
        M_Main.gTestFault = ""
        Check errNumber = ERR_INJECTED_FAULT, "expected the injected error, got " & errNumber & ": " & errText
        CheckState xlCalculationAutomatic, False, False

    'The re-entry flag must have been cleared: the next run completes.
        result = M_Main.RunSACCR_Silent()
        Check Left$(result, 9) = "RESULT=OK", "follow-up run: " & result
        CheckState xlCalculationAutomatic, False, False
        Exit Sub

'------------------------------------------------------------------------------
' HANDLE UNEXPECTED ERROR
'------------------------------------------------------------------------------
Unexpected:
        Fail "run.operation-failure", "unexpected error " & Err.Number & ": " & Err.Description

End Sub


Private Sub CaseCleanupFailure()
'
'==============================================================================
'                              CaseCleanupFailure
'------------------------------------------------------------------------------
' PURPOSE
'   A failed restoration is reported as ERR_CLEANUP_FAILED, the other
'   settings are still restored, and the next run works.
'
' UPDATED
'   2026-10-06
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim result      As String    'Result line of the follow-up run
    Dim errNumber   As Long      'Error raised by the failed run
    Dim errText     As String    'Its description

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
    'The injected fault stops the calculation-mode restoration, so the mode
    'stays manual; events and screen updating must still come back False.
        On Error GoTo Unexpected
        BeginCase "run.cleanup-failure"
        SetState xlCalculationAutomatic, False, False
        M_Main.gTestFault = "cleanup"
        On Error Resume Next
        result = M_Main.RunSACCR_Silent()
        errNumber = Err.Number
        errText = Err.Description
        On Error GoTo Unexpected
        M_Main.gTestFault = ""
        Check errNumber = ERR_CLEANUP_FAILED, "expected the cleanup error, got " & errNumber & ": " & errText
        Check InStr(errText, "Calculation:") > 0, "cleanup detail names calculation: " & errText
        Check InStr(errText, "RESULT=OK") > 0, "cleanup error keeps the run result: " & errText
        CheckState xlCalculationManual, False, False

    'After the user restores the mode by hand, the next run works.
        Application.Calculation = xlCalculationAutomatic
        result = M_Main.RunSACCR_Silent()
        Check Left$(result, 9) = "RESULT=OK", "follow-up run: " & result
        CheckState xlCalculationAutomatic, False, False
        Exit Sub

'------------------------------------------------------------------------------
' HANDLE UNEXPECTED ERROR
'------------------------------------------------------------------------------
Unexpected:
        Fail "run.cleanup-failure", "unexpected error " & Err.Number & ": " & Err.Description

End Sub


Private Sub CaseChecksUnwritable()
'
'==============================================================================
'                             CaseChecksUnwritable
'------------------------------------------------------------------------------
' PURPOSE
'   When the Checks sheet cannot be written, here because it is protected,
'   the run raises ERR_CHECKS_WRITE, clears its results, says why on
'   Results, restores Excel, and the next run works once the sheet is
'   writable again (#36).
'
' UPDATED
'   2026-10-11
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim results     As Worksheet    'Results sheet
    Dim result      As String       'Result line of the follow-up run
    Dim errNumber   As Long         'Error raised by the failed run
    Dim errText     As String       'Its description

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
        On Error GoTo Unexpected
        BeginCase "run.checks-unwritable"
        SetState xlCalculationAutomatic, False, False
        Set results = GetSheet(SH_RESULTS)
        GetSheet(SH_CHECKS).Protect
        On Error Resume Next
        result = M_Main.RunSACCR_Silent()
        errNumber = Err.Number
        errText = Err.Description
        On Error GoTo Unexpected
        GetSheet(SH_CHECKS).Unprotect
        Check errNumber = ERR_CHECKS_WRITE, "expected the Checks write error, got " & errNumber & ": " & errText
        Check IsBlankCell(TableCell(results, FIRST_DATA_ROW, 1).Value), "Results rows were not withdrawn"
        Check InStr(1, SafeStr(results.Range(RUNINFO_CELL).Value), "Checks sheet could not be written", _
                    vbTextCompare) > 0, "run summary does not say why: " & SafeStr(results.Range(RUNINFO_CELL).Value)
        CheckState xlCalculationAutomatic, False, False

    'With the sheet writable again, the next run works.
        result = M_Main.RunSACCR_Silent()
        Check Left$(result, 9) = "RESULT=OK", "follow-up run: " & result
        CheckState xlCalculationAutomatic, False, False
        Exit Sub

'------------------------------------------------------------------------------
' HANDLE UNEXPECTED ERROR
'------------------------------------------------------------------------------
Unexpected:
        Fail "run.checks-unwritable", "unexpected error " & Err.Number & ": " & Err.Description
        UnprotectChecks

End Sub


Private Sub UnprotectChecks()
'
'==============================================================================
'                               UnprotectChecks
'------------------------------------------------------------------------------
' PURPOSE
'   Unprotect the Checks sheet after an unexpected error in
'   CaseChecksUnwritable, so that the protection never outlives the case.
'
' ERROR POLICY
'   Contained: a failure is reported as a test failure.
'
' UPDATED
'   2026-10-07
'==============================================================================
'
        On Error GoTo Failed
        GetSheet(SH_CHECKS).Unprotect
        Exit Sub

Failed:
        Fail "run.checks-unwritable", "could not unprotect the Checks sheet: " & Err.Description

End Sub


Private Sub CaseOutputWriteFailure()
'
'==============================================================================
'                            CaseOutputWriteFailure
'------------------------------------------------------------------------------
' PURPOSE
'   An error after TradeCalc, Results and Buckets have been written, and
'   before HedgingSets, is raised unchanged; the output sheets are cleared
'   so that no table of the stopped run is left, Checks and Results C4 say
'   why, Excel is restored, and the next run works (#36).
'
' UPDATED
'   2026-10-11
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim result      As String    'Result line of the follow-up run
    Dim errNumber   As Long      'Error raised by the failed run
    Dim errSource   As String    'Original error source
    Dim errText     As String    'Its description
    Dim summary     As String    'Results C4 after the failed run

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
        On Error GoTo Unexpected
        BeginCase "run.output-write-failure"
        SetState xlCalculationAutomatic, False, False
        CORE_Engine.gEngineFault = "outputs"
        On Error Resume Next
        result = M_Main.RunSACCR_Silent()
        errNumber = Err.Number
        errText = Err.Description
        errSource = Err.Source
        On Error GoTo Unexpected
        CORE_Engine.gEngineFault = ""
        Check errNumber = ERR_INJECTED_FAULT, "expected the injected error, got " & errNumber & ": " & errText
        Check errSource = "CORE_Engine.Calculate", "injected error source was replaced: " & errSource
        Check errText = "Injected output failure.", "injected error description was replaced: " & errText
        Check IsBlankCell(TableCell(GetSheet(SH_TRADECALC), FIRST_DATA_ROW, 1).Value), "TradeCalc rows were not withdrawn"
        Check IsBlankCell(TableCell(GetSheet(SH_RESULTS), FIRST_DATA_ROW, 1).Value), "Results rows were not withdrawn"
        Check IsBlankCell(TableCell(GetSheet(SH_BUCKETS), FIRST_DATA_ROW, 1).Value), "Buckets rows were not withdrawn"
        summary = SafeStr(GetSheet(SH_RESULTS).Range(RUNINFO_CELL).Value)
        Check InStr(1, summary, "results withdrawn", vbTextCompare) > 0, "run summary does not say why: " & summary
        Check ChecksMention("Injected output failure"), "Checks does not show the error"
        Check M_Main.ResultsStatus() = "NONE", "results status after a failed run: " & M_Main.ResultsStatus()
        CheckState xlCalculationAutomatic, False, False

    'The next run works and writes every sheet again.
        result = M_Main.RunSACCR_Silent()
        Check Left$(result, 9) = "RESULT=OK", "follow-up run: " & result
        CheckState xlCalculationAutomatic, False, False
        Exit Sub

'------------------------------------------------------------------------------
' HANDLE UNEXPECTED ERROR
'------------------------------------------------------------------------------
Unexpected:
        CORE_Engine.gEngineFault = ""
        Fail "run.output-write-failure", "unexpected error " & Err.Number & ": " & Err.Description

End Sub


Private Sub CaseOutputCleanupFailure()
'
'==============================================================================
'                           CaseOutputCleanupFailure
'------------------------------------------------------------------------------
' PURPOSE
'   A known primary write error survives two different cleanup errors. This
'   complements the real protection cases, whose Excel errors can be alike.
' ERROR POLICY
'   Always reset the fault seam; the normal recovery run clears retained data.
' UPDATED
'   2026-10-11
'==============================================================================
    Dim result As String       'Silent result
    Dim errNumber As Long      'Raised primary error
    Dim errSource As String    'Raised source
    Dim errText As String      'Primary text with appended cleanup diagnostics
        On Error GoTo Unexpected
        BeginCase "run.output-and-cleanup-failure"
        SetState xlCalculationAutomatic, False, False
        CORE_Engine.gEngineFault = "outputs+cleanup"
        On Error Resume Next
        result = M_Main.RunSACCR_Silent()
        errNumber = Err.Number
        errSource = Err.Source
        errText = Err.Description
        On Error GoTo Unexpected
        CORE_Engine.gEngineFault = ""
        Check errNumber = ERR_INJECTED_FAULT, "cleanup replaced the primary error number"
        Check errSource = "CORE_Engine.Calculate", "cleanup replaced the primary source"
        Check Left$(errText, Len("Injected output failure.")) = "Injected output failure.", "primary text lost"
        Check InStr(errText, SH_TRADECALC & " [error " & CStr(ERR_CLEANUP_FAILED)) > 0, "first cleanup error lost"
        Check InStr(errText, SH_RESULTS & " [error " & CStr(ERR_CLEANUP_FAILED)) > 0, "second cleanup error lost"
        Check InStr(errText, "source=CORE_Engine.TryClearOutput]: Injected output cleanup failure.") > 0, _
              "cleanup source/description lost"
        Check InStr(1, errText, "results withdrawn", vbTextCompare) = 0, "incomplete cleanup reported as withdrawn"
        Check Not IsBlankCell(TableCell(GetSheet(SH_RESULTS), FIRST_DATA_ROW, 1).Value), "expected retained partial Results"
        Check IsBlankCell(TableCell(GetSheet(SH_BUCKETS), FIRST_DATA_ROW, 1).Value), "Buckets skipped after first cleanup failure"
        Check IsBlankCell(TableCell(GetSheet(SH_HEDGING), FIRST_DATA_ROW, 1).Value), "HedgingSets not cleared"
        Check ChecksMention("Injected output failure."), "Checks lost primary failure"
        Check ChecksMention("Injected output cleanup failure."), "Checks lost cleanup failure"
        Check M_Main.ResultsStatus() = "NONE", "failed run left valid results"
        CheckState xlCalculationAutomatic, False, False
        result = M_Main.RunSACCR_Silent()
        Check Left$(result, 9) = "RESULT=OK", "recovery run: " & result
        Check M_Main.ResultsStatus() = "CURRENT", "recovery fingerprint missing"
        Check Not ChecksMention("Injected output cleanup failure."), "recovery retained cleanup diagnostic"
        CheckState xlCalculationAutomatic, False, False
        Exit Sub
Unexpected:
        CORE_Engine.gEngineFault = ""
        Fail "run.output-and-cleanup-failure", "unexpected error " & Err.Number & ": " & Err.Description
End Sub


Private Sub CaseProtectedOutput( _
    ByVal protectedSheet As String, _
    Optional ByVal secondSheet As String = "", _
    Optional ByVal protectChecks As Boolean = False)
'
'==============================================================================
'                             CaseProtectedOutput
'------------------------------------------------------------------------------
' PURPOSE
'   Protect early/late outputs, singly and together, with Checks writable or
'   protected. Every writable output must clear, all cleanup failures must
'   be named, the primary error must survive, and an unprotected rerun works.
' INPUTS
'   protectedSheet, secondSheet: outputs to protect (second may be empty).
'   protectChecks: also make the diagnostic sheet unwritable.
' ERROR POLICY
'   Undo only protection owned by this case, including on unexpected errors.
' UPDATED
'   2026-10-11
'==============================================================================
    Dim result As String              'Run result
    Dim errNumber As Long             'Reported run error
    Dim errSource As String           'Reported error source
    Dim errText As String             'Reported error description
    Dim primaryNumber As Long         'First failing clear, before running engine
    Dim primarySource As String       'Expected primary source
    Dim primaryText As String         'Expected primary description
    Dim probeNumber As Long           'One protected clear's error
    Dim probeSource As String         'One protected clear's source
    Dim probeText As String           'One protected clear's description
    Dim detail(0 To 3) As String       'Expected exact cleanup detail per sheet
    Dim lastRows(0 To 3) As Long       'Bounds before clearing (UsedRange may shrink)
    Dim ownsProtection(0 To 3) As Boolean
    Dim ownsChecks As Boolean         'Only undo protection set by this case
    Dim sheetNames As Variant         'Output names in engine clear order
    Dim widths As Variant             'Corresponding output table widths
    Dim i As Long                     'Output index
    Dim oldSummary As String          'Successful summary before protection
    Dim statusFormula As Variant      'C5 formula/text must not be overwritten
    Dim caseName As String            'Distinct regression identifier
        On Error GoTo Unexpected
        caseName = "run.protected-output." & protectedSheet
        If Len(secondSheet) > 0 Then caseName = caseName & "+" & secondSheet
        If protectChecks Then caseName = caseName & "+Checks"
        BeginCase caseName
        SetState xlCalculationAutomatic, False, False
        sheetNames = Array(SH_TRADECALC, SH_BUCKETS, SH_HEDGING, SH_RESULTS)
        widths = Array(TC_NCOLS, BK_NCOLS, HS_NCOLS, RS_NCOLS)
        For i = 0 To 3
            If GetSheet(CStr(sheetNames(i))).ProtectContents Then
                Fail caseName, "test requires initially unprotected " & CStr(sheetNames(i))
                Exit Sub
            End If
        Next i
        If GetSheet(SH_CHECKS).ProtectContents Then
            Fail caseName, "test requires initially unprotected Checks"
            Exit Sub
        End If
        result = M_Main.RunSACCR_Silent()
        Check Left$(result, 9) = "RESULT=OK", "setup run: " & result
        oldSummary = SafeStr(GetSheet(SH_RESULTS).Range(RUNINFO_CELL).Value)
        statusFormula = GetSheet(SH_RESULTS).Range(RESULTS_STATUS_CELL).Formula
        For i = 0 To 3
            TableCell(GetSheet(CStr(sheetNames(i))), FIRST_DATA_ROW, 1).Value = "OLD-OUTPUT"
            lastRows(i) = UsedLastRow(GetSheet(CStr(sheetNames(i))))
            If CStr(sheetNames(i)) = protectedSheet Or CStr(sheetNames(i)) = secondSheet Then
                GetSheet(CStr(sheetNames(i))).Protect
                ownsProtection(i) = True
                'Capture Excel's localized error independently, before the run.
                On Error Resume Next
                ClearOutputBlock GetSheet(CStr(sheetNames(i))), FIRST_DATA_ROW, CLng(widths(i))
                probeNumber = Err.Number
                probeSource = Err.Source
                probeText = Err.Description
                On Error GoTo Unexpected
                Check probeNumber <> 0, "protected clear probe did not fail"
                If primaryNumber = 0 Then
                    primaryNumber = probeNumber
                    primarySource = probeSource
                    primaryText = probeText
                End If
                detail(i) = CStr(sheetNames(i)) & " [error " & CStr(probeNumber) & _
                            "; source=" & probeSource & "]: " & probeText
            End If
        Next i
        If protectChecks Then
            GetSheet(SH_CHECKS).Protect
            ownsChecks = True
        End If
        On Error Resume Next
        result = M_Main.RunSACCR_Silent()
        errNumber = Err.Number
        errSource = Err.Source
        errText = Err.Description
        On Error GoTo Unexpected
        Check errNumber = primaryNumber And errNumber <> 0, "primary error number was replaced"
        Check errSource = primarySource, "primary error source was replaced: " & errSource
        Check Left$(errText, Len(primaryText)) = primaryText, "original error description was lost: " & errText
        Check InStr(1, errText, "output cleanup incomplete", vbTextCompare) > 0, "missing incomplete cleanup warning"
        Check InStr(1, errText, "results withdrawn", vbTextCompare) = 0, "reported withdrawal despite retained outputs"
        For i = 0 To 3
            If ownsProtection(i) Then
                Check SafeStr(TableCell(GetSheet(CStr(sheetNames(i))), FIRST_DATA_ROW, 1).Value) = "OLD-OUTPUT", _
                      CStr(sheetNames(i)) & " protected sentinel unexpectedly changed"
                Check InStr(1, errText, detail(i), vbBinaryCompare) > 0, "missing cleanup detail: " & detail(i)
                If Not protectChecks Then Check ChecksMention(detail(i)), "Checks lacks detail: " & detail(i)
            Else
                Check Application.CountA(GetSheet(CStr(sheetNames(i))).Range( _
                      TableCell(GetSheet(CStr(sheetNames(i))), FIRST_DATA_ROW, 1), _
                      TableCell(GetSheet(CStr(sheetNames(i))), lastRows(i), CLng(widths(i))))) = 0, _
                      CStr(sheetNames(i)) & " retained output cells"
            End If
        Next i
        If ownsProtection(3) Then
            Check SafeStr(GetSheet(SH_RESULTS).Range(RUNINFO_CELL).Value) = oldSummary, _
                  "protected Results summary unexpectedly changed"
            Check InStr(1, errText, "old results may remain", vbTextCompare) > 0, "retained Results not disclosed"
        Else
            Check InStr(1, SafeStr(GetSheet(SH_RESULTS).Range(RUNINFO_CELL).Value), _
                  "output cleanup incomplete", vbTextCompare) > 0, "Results summary hides cleanup failure"
        End If
        If protectChecks Then
            Check InStr(1, errText, "Checks could not be updated", vbTextCompare) > 0, "Checks failure was not reported"
        Else
            Check ChecksMention("output cleanup incomplete"), "Checks did not report incomplete cleanup"
        End If
        Check GetSheet(SH_RESULTS).Range(RESULTS_STATUS_CELL).Formula = statusFormula, "cleanup overwrote the status formula"
        Check M_Main.ResultsStatus() = "NONE", "failed run retained a current fingerprint"
        Check Len(CORE_Engine.RunInputs) = 0, "failed run retained in-memory run inputs"
        CheckState xlCalculationAutomatic, False, False
        For i = 0 To 3
            If ownsProtection(i) Then
                GetSheet(CStr(sheetNames(i))).Unprotect
                ownsProtection(i) = False
            End If
        Next i
        If ownsChecks Then
            GetSheet(SH_CHECKS).Unprotect
            ownsChecks = False
        End If
        result = M_Main.RunSACCR_Silent()
        Check Left$(result, 9) = "RESULT=OK", "recovery run: " & result
        Check M_Main.ResultsStatus() = "CURRENT", "recovery did not publish a valid fingerprint"
        Check Not ChecksMention("output cleanup incomplete"), "recovery retained obsolete failure diagnostics"
        For i = 0 To 3
            Check SafeStr(TableCell(GetSheet(CStr(sheetNames(i))), FIRST_DATA_ROW, 1).Value) <> "OLD-OUTPUT", _
                  CStr(sheetNames(i)) & " retained the old sentinel after recovery"
        Next i
        CheckState xlCalculationAutomatic, False, False
        Exit Sub
Unexpected:
        Fail caseName, "unexpected error " & Err.Number & ": " & Err.Description
        For i = 0 To 3
            If ownsProtection(i) Then UnprotectOutput CStr(sheetNames(i))
        Next i
        If ownsChecks Then UnprotectOutput SH_CHECKS
End Sub


Private Sub UnprotectOutput(ByVal sheetName As String)
'
'==============================================================================
' PURPOSE
'   Release protection owned by CaseProtectedOutput after an unexpected error.
'==============================================================================
        On Error GoTo Failed
        GetSheet(sheetName).Unprotect
        Exit Sub
Failed:
        Fail "protected-output", "could not unprotect " & sheetName & ": " & Err.Description
End Sub


Private Sub CaseStaleResults()
'
'==============================================================================
'                               CaseStaleResults
'------------------------------------------------------------------------------
' PURPOSE
'   After a run the results are CURRENT; an edited input makes them STALE
'   and the status formula in Results C5 says OUT OF DATE; undoing the edit
'   makes them CURRENT again; Clear outputs makes them NONE (#36). Each
'   status is checked both through ResultsStatus and in the cell.
'
' UPDATED
'   2026-10-11
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim result    As String     'Result line of a run
    Dim mtmCell   As Range      'MtM of the first trade, edited and restored
    Dim saved     As Variant    'Its formula

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
        On Error GoTo Unexpected
        BeginCase "results.stale-after-input-change"
        SetState xlCalculationAutomatic, False, False
        result = M_Main.RunSACCR_Silent()
        Check InStr(result, "; inputs=") > 0 And InStr(result, "; inputs=;") = 0, "result line names no inputs: " & result
        Check M_Main.ResultsStatus() = "CURRENT", "after a run: " & M_Main.ResultsStatus()
        CheckStatusCell "Results current", "after a run"

    'Edit one input, as a user would: the formula recalculates by itself.
        Set mtmCell = TableCell(GetSheet(SH_TRADES), TableFirstRow(SH_TRADES), TR_MTM)
        saved = mtmCell.Formula
        mtmCell.Value = ToDbl(mtmCell.Value, 0#) + 1#
        Check M_Main.ResultsStatus() = "STALE", "after an edit: " & M_Main.ResultsStatus()
        CheckStatusCell "OUT OF DATE", "after an edit"

    'Undoing the edit brings the results back in line.
        mtmCell.Formula = saved
        Check M_Main.ResultsStatus() = "CURRENT", "after undoing the edit: " & M_Main.ResultsStatus()
        CheckStatusCell "Results current", "after undoing the edit"

    'Clear outputs leaves no results; a new run makes them current again.
        M_Main.ClearOutputs
        Check M_Main.ResultsStatus() = "NONE", "after Clear outputs: " & M_Main.ResultsStatus()
        CheckStatusCell "No results", "after Clear outputs"
        result = M_Main.RunSACCR_Silent()
        Check M_Main.ResultsStatus() = "CURRENT", "after a new run: " & M_Main.ResultsStatus()
        CheckStatusCell "Results current", "after a new run"
        CheckState xlCalculationAutomatic, False, False
        Exit Sub

'------------------------------------------------------------------------------
' HANDLE UNEXPECTED ERROR
'------------------------------------------------------------------------------
Unexpected:
        Fail "results.stale-after-input-change", "unexpected error " & Err.Number & ": " & Err.Description
        RestoreCell mtmCell, saved

End Sub


Private Sub CaseStaleParams( _
    ByVal shifted As Boolean)
'
'==============================================================================
'                               CaseStaleParams
'------------------------------------------------------------------------------
' PURPOSE
'   After a run, an edit to the Regimes field of a factor, to a parameter
'   value or to an FX rate makes the results STALE and Results C5 OUT OF
'   DATE, and undoing it makes them CURRENT again; a changed parameter
'   description or cell colour leaves them CURRENT (#152). With shifted
'   True a column is first inserted before the Params tables, which moves
'   every Params input and workbook name one column to the right, as the
'   issue describes; the run must then give the same result line,
'   fingerprint included, as the unshifted run, and the column is deleted
'   again at the end.
'
' UPDATED
'   2026-10-12
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ws           As Worksheet    'Params sheet
    Dim inserted     As Boolean      'The extra column is on Params
    Dim result       As String       'Result line of the run
    Dim sfCol        As Long         'Table column of the factor keys
    Dim fxCol        As Long         'Table column of the FX currencies
    Dim descCell     As Range        'Description edited, Nothing once restored
    Dim descSaved    As Variant      'Its formula
    Dim colorCell    As Range        'Cell recoloured, Nothing once restored
    Dim colorSaved   As Variant      'Its fill colour

'------------------------------------------------------------------------------
' LAYOUT AND RUN
'------------------------------------------------------------------------------
        On Error GoTo Unexpected
        BeginCase IIf(shifted, "results.stale-params-shifted", "results.stale-params")
        SetState xlCalculationAutomatic, False, False
        Set ws = GetSheet(SH_PARAMS)
        If shifted Then
            ws.Columns(COL_OFFSET + 1).Insert
            inserted = True
        End If
        sfCol = NamedTableCol(NM_SF_TABLE)
        fxCol = NamedTableCol(NM_FX_TABLE)
        Check sfCol = IIf(shifted, 2, 1), "factor table in table column " & sfCol
        result = M_Main.RunSACCR_Silent()
        Check M_Main.ResultsStatus() = "CURRENT", "after a run: " & M_Main.ResultsStatus() & " (" & result & ")"
        CheckStatusCell "Results current", "after a run"
        If shifted Then
            Check result = mParamsRunLine, "shifted run differs: " & result & " | unshifted: " & mParamsRunLine
        Else
            mParamsRunLine = result
        End If

'------------------------------------------------------------------------------
' INPUTS: STALE, THEN CURRENT AGAIN
'------------------------------------------------------------------------------
    'Regimes is the eighth factor column: K once shifted, outside the
    'C:J block fingerprinted before #152.
        EditAndUndo TableCell(ws, NamedRow(NM_SF_TABLE) + 1, sfCol + 7), "factor Regimes"
        EditAndUndo NamedCell(PRM_ALPHA), "parameter Alpha"
        EditAndUndo TableCell(ws, NamedRow(NM_FX_TABLE) + 1, fxCol + 2), "FX rate"

'------------------------------------------------------------------------------
' NOT INPUTS: STILL CURRENT
'------------------------------------------------------------------------------
    'The description beside a parameter code, then a fill colour.
        Set descCell = TableCell(ws, NamedRow(PRM_ALPHA), ParamsCodeCol() + 1)
        descSaved = descCell.Formula
        descCell.Value = SafeStr(descCell.Value) & " (edited)"
        Check M_Main.ResultsStatus() = "CURRENT", "after a description edit: " & M_Main.ResultsStatus()
        CheckStatusCell "Results current", "after a description edit"
        descCell.Formula = descSaved
        Set descCell = Nothing
        Set colorCell = NamedCell(PRM_ALPHA)
        colorSaved = colorCell.Interior.Color
        colorCell.Interior.Color = RGB(255, 0, 0)
        Check M_Main.ResultsStatus() = "CURRENT", "after a colour change: " & M_Main.ResultsStatus()
        CheckStatusCell "Results current", "after a colour change"
        colorCell.Interior.Color = colorSaved
        Set colorCell = Nothing

'------------------------------------------------------------------------------
' RESTORE THE LAYOUT
'------------------------------------------------------------------------------
        If inserted Then
            ws.Columns(COL_OFFSET + 1).Delete
            inserted = False
            Check NamedTableCol(NM_SF_TABLE) = 1, "factor table not back in table column 1"
            Check M_Main.ResultsStatus() = "CURRENT", "after deleting the column: " & M_Main.ResultsStatus()
        End If
        CheckState xlCalculationAutomatic, False, False
        Exit Sub

'------------------------------------------------------------------------------
' HANDLE UNEXPECTED ERROR
'------------------------------------------------------------------------------
    'Undo whatever is still changed: the description, the colour, the
    'inserted column.
Unexpected:
        Fail mCurrentCase, "unexpected error " & Err.Number & ": " & Err.Description
        On Error Resume Next
        If Not descCell Is Nothing Then descCell.Formula = descSaved
        If Not colorCell Is Nothing Then colorCell.Interior.Color = colorSaved
        If inserted Then ws.Columns(COL_OFFSET + 1).Delete
        If Err.Number <> 0 Then
            Fail mCurrentCase, "could not restore Params: " & Err.Description
        End If

End Sub


Private Sub EditAndUndo( _
    ByVal target As Range, _
    ByVal what As String)
'
'==============================================================================
'                                 EditAndUndo
'------------------------------------------------------------------------------
' PURPOSE
'   Change an input cell, check that the results are STALE and Results C5
'   says OUT OF DATE, put the cell back, and check that they are CURRENT
'   again. A number is raised by 1; any other value gets an X appended.
'
' UPDATED
'   2026-10-12
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim saved   As Variant    'The cell's formula

'------------------------------------------------------------------------------
' EDIT AND UNDO
'------------------------------------------------------------------------------
        On Error GoTo Unexpected
        saved = target.Formula
        If VarType(target.Value2) = vbDouble Then
            target.Value = target.Value2 + 1#
        Else
            target.Value = SafeStr(target.Value2) & "X"
        End If
        Check M_Main.ResultsStatus() = "STALE", what & " edited: " & M_Main.ResultsStatus()
        CheckStatusCell "OUT OF DATE", what & " edited"
        target.Formula = saved
        Check M_Main.ResultsStatus() = "CURRENT", what & " restored: " & M_Main.ResultsStatus()
        CheckStatusCell "Results current", what & " restored"
        Exit Sub

'------------------------------------------------------------------------------
' HANDLE UNEXPECTED ERROR
'------------------------------------------------------------------------------
Unexpected:
        Fail mCurrentCase, what & ": unexpected error " & Err.Number & ": " & Err.Description
        RestoreCell target, saved

End Sub


Private Sub CheckStatusCell( _
    ByVal expectedStart As String, _
    ByVal moment As String)
'
'==============================================================================
'                               CheckStatusCell
'------------------------------------------------------------------------------
' PURPOSE
'   Check that the status formula in Results C5 starts with expectedStart,
'   after Excel has recalculated what changed.
'
' UPDATED
'   2026-10-11
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim shown   As String    'Text in Results C5

'------------------------------------------------------------------------------
' CHECK
'------------------------------------------------------------------------------
        Application.Calculate
        shown = SafeStr(GetSheet(SH_RESULTS).Range(RESULTS_STATUS_CELL).Value)
        Check Left$(shown, Len(expectedStart)) = expectedStart, _
              "Results C5 " & moment & ": expected """ & expectedStart & "..."", found """ & shown & """"

End Sub


Private Sub RestoreCell( _
    ByVal target As Range, _
    ByVal saved As Variant)
'
'==============================================================================
'                                 RestoreCell
'------------------------------------------------------------------------------
' PURPOSE
'   Put back an input cell edited by CaseStaleResults or EditAndUndo after
'   an unexpected error.
'
' ERROR POLICY
'   Contained: a failure is reported as a test failure of the current
'   case.
'
' UPDATED
'   2026-10-12
'==============================================================================
'
        On Error GoTo Failed
        If Not target Is Nothing Then
            target.Formula = saved
        End If
        Exit Sub

Failed:
        Fail mCurrentCase, "could not restore the edited input: " & Err.Description

End Sub


Private Function ChecksMention( _
    ByVal fragment As String) _
    As Boolean
'
'==============================================================================
'                                ChecksMention
'------------------------------------------------------------------------------
' PURPOSE
'   Tell whether a message on the Checks sheet contains fragment, ignoring
'   case.
'
' UPDATED
'   2026-10-11
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ws   As Worksheet    'Checks sheet
    Dim r    As Long         'Row being read

'------------------------------------------------------------------------------
' SEARCH
'------------------------------------------------------------------------------
    'Checks column 5 holds the message.
        Set ws = GetSheet(SH_CHECKS)
        For r = FIRST_DATA_ROW To UsedLastRow(ws)
            If InStr(1, SafeStr(TableCell(ws, r, 5).Value), fragment, vbTextCompare) > 0 Then
                ChecksMention = True
                Exit Function
            End If
        Next r

End Function


Private Sub CaseSilentFlagPreserved()
'
'==============================================================================
'                           CaseSilentFlagPreserved
'------------------------------------------------------------------------------
' PURPOSE
'   RunSACCR_Silent restores gSilent to its previous value, False or True,
'   on success and when the run raises.
'
' UPDATED
'   2026-10-06
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim result   As String    'Result line

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
        On Error GoTo Unexpected
        BeginCase "silent.flag-preserved"
        SetState xlCalculationAutomatic, True, True
        M_Main.gSilent = False
        result = M_Main.RunSACCR_Silent()
        Check M_Main.gSilent = False, "gSilent was False before a successful run"
        M_Main.gSilent = True
        M_Main.gTestFault = "operation"
        On Error Resume Next
        result = M_Main.RunSACCR_Silent()
        On Error GoTo Unexpected
        M_Main.gTestFault = ""
        Check M_Main.gSilent = True, "gSilent was True before a failed run"
        M_Main.gSilent = False
        M_Main.gTestFault = "operation"
        On Error Resume Next
        result = M_Main.RunSACCR_Silent()
        On Error GoTo Unexpected
        M_Main.gTestFault = ""
        Check M_Main.gSilent = False, "gSilent was False before a failed run"
        M_Main.gSilent = True
        CheckState xlCalculationAutomatic, True, True
        Exit Sub

'------------------------------------------------------------------------------
' HANDLE UNEXPECTED ERROR
'------------------------------------------------------------------------------
Unexpected:
        M_Main.gTestFault = ""
        M_Main.gSilent = True
        Fail "silent.flag-preserved", "unexpected error " & Err.Number & ": " & Err.Description

End Sub


'
'------------------------------------------------------------------------------
'
'                                   HELPERS
'
'------------------------------------------------------------------------------
'

Private Sub BeginCase( _
    ByVal caseName As String)
'
'==============================================================================
'                                  BeginCase
'------------------------------------------------------------------------------
' PURPOSE
'   Count a case and print its CASE line.
'
' UPDATED
'   2026-10-06
'==============================================================================
'
        mCaseCount = mCaseCount + 1
        mCurrentCase = caseName
        Debug.Print "CASE=" & caseName

End Sub


Private Sub SetState( _
    ByVal calcMode As XlCalculation, _
    ByVal eventsOn As Boolean, _
    ByVal screenOn As Boolean)
'
'==============================================================================
'                                   SetState
'------------------------------------------------------------------------------
' PURPOSE
'   Put the three settings into a known state before an operation.
'
' UPDATED
'   2026-10-06
'==============================================================================
'
        Application.Calculation = calcMode
        Application.EnableEvents = eventsOn
        Application.ScreenUpdating = screenOn

End Sub


Private Sub CheckState( _
    ByVal calcMode As XlCalculation, _
    ByVal eventsOn As Boolean, _
    ByVal screenOn As Boolean)
'
'==============================================================================
'                                  CheckState
'------------------------------------------------------------------------------
' PURPOSE
'   Check each of the three settings against its expected value.
'
' UPDATED
'   2026-10-06
'==============================================================================
'
        Check Application.Calculation = calcMode, _
              "Calculation expected " & calcMode & ", found " & Application.Calculation
        Check Application.EnableEvents = eventsOn, _
              "EnableEvents expected " & eventsOn & ", found " & Application.EnableEvents
        Check Application.ScreenUpdating = screenOn, _
              "ScreenUpdating expected " & screenOn & ", found " & Application.ScreenUpdating

End Sub


Private Sub Check( _
    ByVal passed As Boolean, _
    ByVal detail As String)
'
'==============================================================================
'                                    Check
'------------------------------------------------------------------------------
' PURPOSE
'   Count one check and record it when it fails.
'
' UPDATED
'   2026-10-06
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
'                                     Fail
'------------------------------------------------------------------------------
' PURPOSE
'   Count a failure and print it.
'
' UPDATED
'   2026-10-06
'==============================================================================
'
        mFailureCount = mFailureCount + 1
        Debug.Print "FAILURE=" & caseName & ": " & detail

End Sub
