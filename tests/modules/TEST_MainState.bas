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
'   (issue #152), when a number becomes the same digits as text, a header
'   is renamed or the date system changes, but not after a formatting
'   change (issue #197); that an output cell edited after a run is
'   reported as EDITED, refused by the evidence export and cleared by a
'   new run (issue #315); that a run summary that cannot be composed does
'   not fail the run (issue #60); that a run in manual calculation
'   recalculates its formula inputs before reading them, and stops without
'   reporting old results as current when that recalculation does not
'   finish (issue #195) or raises an error (issue #204); and that every
'   output table is located through the workbook name of its first header
'   cell, so that it can move.
'
' PUBLIC SURFACE
'   RunMainStateTests is the entry point. Option Private Module keeps it out
'   of the external workbook automation API.
'
' DEPENDENCIES
'   M_Main, its flags gSilent and gQuietLog and its test seam gTestFault,
'   the CORE_Engine test seam gEngineFault, the sheet and table helpers of
'   CORE_Util, the names, columns and error numbers in CORE_Config, and
'   TEST_Suite.Say for the log lines. The workbook must be built from the
'   template with the full source, because the macros calculate and write
'   the output sheets.
'
' WORKSHEET SAFETY
'   Unlike TEST_Harness, this module runs the real macros: they clear and
'   rewrite TradeCalc, Buckets, HedgingSets, Results, Checks and the
'   reconciliation sheets, and activate the Checks sheet. Several cases
'   edit an input or output cell and put it back. Some cases protect the
'   Checks sheet or output sheets for one run and unprotect them again. One
'   case inserts a column before the Params tables for its run and deletes
'   it again. One case switches the workbook to the 1904 date system and
'   back. One case inserts a row above the HedgingSets table for its run
'   and deletes it again, and breaks the Checks_Header name and restores
'   it. Each of these is undone also after an unexpected error. The last
'   case leaves the outputs of a normal run. Use a development workbook
'   only.
'
' STATE OWNERSHIP
'   Sets calculation mode, events and screen updating to known values for
'   each case. At the end it restores the caller's values, gSilent and
'   gQuietLog, and leaves gTestFault and gEngineFault empty; a case that
'   fails unexpectedly does not stop the others from running.
'   gQuietLog is set for the run: the macros' run messages would push the
'   report out of the Immediate window, which keeps about 200 lines.
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
        Private Const EXPECTED_CASES   As Long = 29   'Cases in a complete run

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
'   gSilent and gQuietLog, and empties gTestFault and gEngineFault, before
'   printing the result.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim CallerCalculation   As XlCalculation    'Caller's calculation mode
    Dim CallerEvents        As Boolean          'Caller's EnableEvents
    Dim CallerScreen        As Boolean          'Caller's ScreenUpdating
    Dim CallerSilent        As Boolean          'Caller's M_Main.gSilent
    Dim CallerQuiet         As Boolean          'Caller's M_Main.gQuietLog
    Dim Restored            As Boolean          'Caller's settings read back equal

'------------------------------------------------------------------------------
' START
'------------------------------------------------------------------------------
        mCaseCount = 0
        mCheckCount = 0
        mFailureCount = 0
        CallerCalculation = Application.Calculation
        CallerEvents = Application.EnableEvents
        CallerScreen = Application.ScreenUpdating
        CallerSilent = M_Main.gSilent
        CallerQuiet = M_Main.gQuietLog
        TEST_Suite.Say "SACCR MAIN STATE TESTS"

'------------------------------------------------------------------------------
' RUN CASES
'------------------------------------------------------------------------------
    'Silent mode keeps message boxes from blocking the run and makes the
    'macros raise their errors; gQuietLog keeps their run messages out of
    'the report.
        M_Main.gSilent = True
        M_Main.gQuietLog = True
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
        CaseRunInfoFailure
        CaseProtectedOutput SH_TRADECALC
        CaseProtectedOutput SH_BUCKETS
        CaseProtectedOutput SH_HEDGING
        CaseProtectedOutput SH_RESULTS
        CaseProtectedOutput SH_TRADECALC, SH_HEDGING
        CaseProtectedOutput SH_BUCKETS, SH_RESULTS
        CaseProtectedOutput SH_TRADECALC, SH_RESULTS, True
        CaseStaleResults
        CaseEditedOutput
        CaseStaleParams False
        CaseStaleParams True
        CaseStaleValueType
        CaseStaleHeaders
        CaseStaleDateSystem
        CaseManualRecalculation
        CaseRecalculationPending
        CaseRecalculationError
        CaseOutputHeaderNames
        CaseSilentFlagPreserved

'------------------------------------------------------------------------------
' RESTORE THE CALLER AND REPORT
'------------------------------------------------------------------------------
        M_Main.gTestFault = ""
        CORE_Engine.gEngineFault = ""
        M_Main.gSilent = CallerSilent
        M_Main.gQuietLog = CallerQuiet
        Application.Calculation = CallerCalculation
        Application.EnableEvents = CallerEvents
        Application.ScreenUpdating = CallerScreen
        Restored = (Application.Calculation = CallerCalculation) And _
                   (Application.EnableEvents = CallerEvents) And _
                   (Application.ScreenUpdating = CallerScreen)
        If mCaseCount <> EXPECTED_CASES Then
            Fail "suite", "expected " & EXPECTED_CASES & " cases, ran " & mCaseCount
        End If
        TEST_Suite.Say "CASES=" & mCaseCount & "; CHECKS=" & mCheckCount & "; FAILURES=" & mFailureCount
        TEST_Suite.Say "RESULT=" & IIf(mFailureCount = 0 And Restored, "PASS", "FAIL") & _
                    "; cases=" & mCaseCount & "; checks=" & mCheckCount & _
                    "; failures=" & mFailureCount & "; caller_state=" & IIf(Restored, "RESTORED", "CHANGED")

End Sub


'
'------------------------------------------------------------------------------
'
'                                    CASES
'
'------------------------------------------------------------------------------
'

Private Sub CaseRestoresState( _
    ByVal CaseName As String, _
    ByVal Operation As String, _
    ByVal CalcMode As XlCalculation, _
    ByVal EventsOn As Boolean, _
    ByVal ScreenOn As Boolean)
'
'==============================================================================
'                              CaseRestoresState
'------------------------------------------------------------------------------
' PURPOSE
'   Set known settings, run one operation, and check that every setting is
'   back to its value, including settings that were off before the run.
'
' INPUTS
'   CaseName: case identifier.
'   Operation: "clear", "validate" or "run".
'   CalcMode, EventsOn, ScreenOn: the settings before the operation.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Result   As String    'Result line of a run

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
        On Error GoTo Unexpected
        BeginCase CaseName
        SetState CalcMode, EventsOn, ScreenOn
        Select Case Operation
            Case "clear"
                M_Main.ClearOutputs
            Case "validate"
                M_Main.ValidateInputs
            Case "run"
                Result = M_Main.RunSACCR_Silent()
                Check Left$(Result, 9) = "RESULT=OK", "run result: " & Result
        End Select
        CheckState CalcMode, EventsOn, ScreenOn
        Exit Sub

'------------------------------------------------------------------------------
' HANDLE UNEXPECTED ERROR
'------------------------------------------------------------------------------
Unexpected:
        Fail CaseName, "unexpected error " & Err.Number & ": " & Err.Description

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
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Result        As String    'Result line of the follow-up run
    Dim ErrorNumber   As Long      'Error raised by the failed run
    Dim ErrorText     As String    'Its description

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
        On Error GoTo Unexpected
        BeginCase "run.operation-failure"
        SetState xlCalculationAutomatic, False, False
        M_Main.gTestFault = "operation"
        On Error Resume Next
        Result = M_Main.RunSACCR_Silent()
        ErrorNumber = Err.Number
        ErrorText = Err.Description
        On Error GoTo Unexpected
        M_Main.gTestFault = ""
        Check ErrorNumber = ERR_INJECTED_FAULT, "expected the injected error, got " & ErrorNumber & ": " & ErrorText
        CheckState xlCalculationAutomatic, False, False

    'The re-entry flag must have been cleared: the next run completes.
        Result = M_Main.RunSACCR_Silent()
        Check Left$(Result, 9) = "RESULT=OK", "follow-up run: " & Result
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
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Result        As String    'Result line of the follow-up run
    Dim ErrorNumber   As Long      'Error raised by the failed run
    Dim ErrorText     As String    'Its description

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
        Result = M_Main.RunSACCR_Silent()
        ErrorNumber = Err.Number
        ErrorText = Err.Description
        On Error GoTo Unexpected
        M_Main.gTestFault = ""
        Check ErrorNumber = ERR_CLEANUP_FAILED, "expected the cleanup error, got " & ErrorNumber & ": " & ErrorText
        Check InStr(ErrorText, "Calculation:") > 0, "cleanup detail names calculation: " & ErrorText
        Check InStr(ErrorText, "RESULT=OK") > 0, "cleanup error keeps the run result: " & ErrorText
        CheckState xlCalculationManual, False, False

    'After the user restores the mode by hand, the next run works.
        Application.Calculation = xlCalculationAutomatic
        Result = M_Main.RunSACCR_Silent()
        Check Left$(Result, 9) = "RESULT=OK", "follow-up run: " & Result
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
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ResultsSheet   As Worksheet    'Results sheet
    Dim Result         As String       'Result line of the follow-up run
    Dim ErrorNumber    As Long         'Error raised by the failed run
    Dim ErrorText      As String       'Its description

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
        On Error GoTo Unexpected
        BeginCase "run.checks-unwritable"
        SetState xlCalculationAutomatic, False, False
        Set ResultsSheet = GetSheet(SH_RESULTS)
        GetSheet(SH_CHECKS).Protect
        On Error Resume Next
        Result = M_Main.RunSACCR_Silent()
        ErrorNumber = Err.Number
        ErrorText = Err.Description
        On Error GoTo Unexpected
        GetSheet(SH_CHECKS).Unprotect
        Check ErrorNumber = ERR_CHECKS_WRITE, "expected the Checks write error, got " & ErrorNumber & ": " & ErrorText
        Check IsBlankCell(TableCell(ResultsSheet, OutputFirstRow(SH_RESULTS), 1).Value), "Results rows were not withdrawn"
        Check InStr(1, SafeStr(ResultsSheet.Range(RUNINFO_CELL).Value), "Checks sheet could not be written", _
                    vbTextCompare) > 0, "run summary does not say why: " & SafeStr(ResultsSheet.Range(RUNINFO_CELL).Value)
        CheckState xlCalculationAutomatic, False, False

    'With the sheet writable again, the next run works.
        Result = M_Main.RunSACCR_Silent()
        Check Left$(Result, 9) = "RESULT=OK", "follow-up run: " & Result
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
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Result        As String    'Result line of the follow-up run
    Dim ErrorNumber   As Long      'Error raised by the failed run
    Dim ErrorSource   As String    'Original error source
    Dim ErrorText     As String    'Its description
    Dim Summary       As String    'Results C4 after the failed run

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
        On Error GoTo Unexpected
        BeginCase "run.output-write-failure"
        SetState xlCalculationAutomatic, False, False
        CORE_Engine.gEngineFault = "outputs"
        On Error Resume Next
        Result = M_Main.RunSACCR_Silent()
        ErrorNumber = Err.Number
        ErrorText = Err.Description
        ErrorSource = Err.Source
        On Error GoTo Unexpected
        CORE_Engine.gEngineFault = ""
        Check ErrorNumber = ERR_INJECTED_FAULT, "expected the injected error, got " & ErrorNumber & ": " & ErrorText
        Check ErrorSource = "CORE_Engine.Calculate", "injected error source was replaced: " & ErrorSource
        Check ErrorText = "Injected output failure.", "injected error description was replaced: " & ErrorText
        Check IsBlankCell(TableCell(GetSheet(SH_TRADECALC), OutputFirstRow(SH_TRADECALC), 1).Value), "TradeCalc rows were not withdrawn"
        Check IsBlankCell(TableCell(GetSheet(SH_RESULTS), OutputFirstRow(SH_RESULTS), 1).Value), "Results rows were not withdrawn"
        Check IsBlankCell(TableCell(GetSheet(SH_BUCKETS), OutputFirstRow(SH_BUCKETS), 1).Value), "Buckets rows were not withdrawn"
        Summary = SafeStr(GetSheet(SH_RESULTS).Range(RUNINFO_CELL).Value)
        Check InStr(1, Summary, "results withdrawn", vbTextCompare) > 0, "run summary does not say why: " & Summary
        Check ChecksMention("Injected output failure"), "Checks does not show the error"
        Check M_Main.ResultsStatus() = "NONE", "results status after a failed run: " & M_Main.ResultsStatus()
        CheckState xlCalculationAutomatic, False, False

    'The next run works and writes every sheet again.
        Result = M_Main.RunSACCR_Silent()
        Check Left$(Result, 9) = "RESULT=OK", "follow-up run: " & Result
        CheckState xlCalculationAutomatic, False, False
        Exit Sub

'------------------------------------------------------------------------------
' HANDLE UNEXPECTED ERROR
'------------------------------------------------------------------------------
Unexpected:
        CORE_Engine.gEngineFault = ""
        Fail "run.output-write-failure", "unexpected error " & Err.Number & ": " & Err.Description

End Sub


Private Sub CaseRunInfoFailure()
'
'==============================================================================
'                              CaseRunInfoFailure
'------------------------------------------------------------------------------
' PURPOSE
'   A run summary that cannot be composed does not fail the run: the run
'   completes, its results are CURRENT, and Results C4 says the summary is
'   unavailable instead of being left empty or partial. The next run
'   writes the full summary again (#60).
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Result    As String    'Result line of a run
    Dim Summary   As String    'Results C4

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
        On Error GoTo Unexpected
        BeginCase "run.summary-unavailable"
        SetState xlCalculationAutomatic, False, False
        CORE_Engine.gEngineFault = "runinfo"
        Result = M_Main.RunSACCR_Silent()
        CORE_Engine.gEngineFault = ""
        Check Left$(Result, 9) = "RESULT=OK", "run with a failed summary: " & Result
        Check M_Main.ResultsStatus() = "CURRENT", "results status: " & M_Main.ResultsStatus()
        Summary = SafeStr(GetSheet(SH_RESULTS).Range(RUNINFO_CELL).Value)
        Check InStr(1, Summary, "run summary unavailable", vbTextCompare) > 0 And _
              InStr(1, Summary, "Injected run-summary failure", vbTextCompare) > 0, _
              "Results C4 does not say the summary is unavailable: " & Summary
        CheckState xlCalculationAutomatic, False, False

    'The next run writes the full summary again.
        Result = M_Main.RunSACCR_Silent()
        Summary = SafeStr(GetSheet(SH_RESULTS).Range(RUNINFO_CELL).Value)
        Check Left$(Result, 9) = "RESULT=OK", "follow-up run: " & Result
        Check InStr(1, Summary, " | inputs ", vbBinaryCompare) > 0, "follow-up summary: " & Summary
        CheckState xlCalculationAutomatic, False, False
        Exit Sub

'------------------------------------------------------------------------------
' HANDLE UNEXPECTED ERROR
'------------------------------------------------------------------------------
Unexpected:
        CORE_Engine.gEngineFault = ""
        Fail "run.summary-unavailable", "unexpected error " & Err.Number & ": " & Err.Description

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
'   2026-10-17
'==============================================================================
    Dim Result        As String    'Silent result
    Dim ErrorNumber   As Long      'Raised primary error
    Dim ErrorSource   As String    'Raised source
    Dim ErrorText     As String    'Primary text with appended cleanup diagnostics
        On Error GoTo Unexpected
        BeginCase "run.output-and-cleanup-failure"
        SetState xlCalculationAutomatic, False, False
        CORE_Engine.gEngineFault = "outputs+cleanup"
        On Error Resume Next
        Result = M_Main.RunSACCR_Silent()
        ErrorNumber = Err.Number
        ErrorSource = Err.Source
        ErrorText = Err.Description
        On Error GoTo Unexpected
        CORE_Engine.gEngineFault = ""
        Check ErrorNumber = ERR_INJECTED_FAULT, "cleanup replaced the primary error number"
        Check ErrorSource = "CORE_Engine.Calculate", "cleanup replaced the primary source"
        Check Left$(ErrorText, Len("Injected output failure.")) = "Injected output failure.", "primary text lost"
        Check InStr(ErrorText, SH_TRADECALC & " [error " & CStr(ERR_CLEANUP_FAILED)) > 0, "first cleanup error lost"
        Check InStr(ErrorText, SH_RESULTS & " [error " & CStr(ERR_CLEANUP_FAILED)) > 0, "second cleanup error lost"
        Check InStr(ErrorText, "source=CORE_Engine.TryClearOutput]: Injected output cleanup failure.") > 0, _
              "cleanup source/description lost"
        Check InStr(1, ErrorText, "results withdrawn", vbTextCompare) = 0, "incomplete cleanup reported as withdrawn"
        Check Not IsBlankCell(TableCell(GetSheet(SH_RESULTS), OutputFirstRow(SH_RESULTS), 1).Value), "expected retained partial Results"
        Check IsBlankCell(TableCell(GetSheet(SH_BUCKETS), OutputFirstRow(SH_BUCKETS), 1).Value), "Buckets skipped after first cleanup failure"
        Check IsBlankCell(TableCell(GetSheet(SH_HEDGING), OutputFirstRow(SH_HEDGING), 1).Value), "HedgingSets not cleared"
        Check ChecksMention("Injected output failure."), "Checks lost primary failure"
        Check ChecksMention("Injected output cleanup failure."), "Checks lost cleanup failure"
        Check M_Main.ResultsStatus() = "NONE", "failed run left valid results"
        CheckState xlCalculationAutomatic, False, False
        Result = M_Main.RunSACCR_Silent()
        Check Left$(Result, 9) = "RESULT=OK", "recovery run: " & Result
        Check M_Main.ResultsStatus() = "CURRENT", "recovery fingerprint missing"
        Check Not ChecksMention("Injected output cleanup failure."), "recovery retained cleanup diagnostic"
        CheckState xlCalculationAutomatic, False, False
        Exit Sub
Unexpected:
        CORE_Engine.gEngineFault = ""
        Fail "run.output-and-cleanup-failure", "unexpected error " & Err.Number & ": " & Err.Description
End Sub


Private Sub CaseProtectedOutput( _
    ByVal ProtectedSheet As String, _
    Optional ByVal SecondSheet As String = "", _
    Optional ByVal ProtectChecks As Boolean = False)
'
'==============================================================================
'                             CaseProtectedOutput
'------------------------------------------------------------------------------
' PURPOSE
'   Protect early/late outputs, singly and together, with Checks writable or
'   protected. Every writable output must clear, all cleanup failures must
'   be named, the primary error must survive, and an unprotected rerun works.
' INPUTS
'   ProtectedSheet, SecondSheet: outputs to protect (SecondSheet may be
'   empty).
'   ProtectChecks: also make the Checks sheet unwritable.
' ERROR POLICY
'   Undo only protection owned by this case, including on unexpected errors.
' UPDATED
'   2026-10-17
'==============================================================================
    Dim Result                    As String     'Run result
    Dim ErrorNumber               As Long       'Reported run error
    Dim ErrorSource               As String     'Reported error source
    Dim ErrorText                 As String     'Reported error description
    Dim PrimaryNumber             As Long       'Error of the first protected clear, probed before the run
    Dim PrimarySource             As String     'Expected primary source
    Dim PrimaryText               As String     'Expected primary description
    Dim ProbeNumber               As Long       'One protected clear's error
    Dim ProbeSource               As String     'One protected clear's source
    Dim ProbeText                 As String     'One protected clear's description
    Dim ExpectedDetails(0 To 3)   As String     'Expected exact cleanup detail per sheet
    Dim LastRows(0 To 3)          As Long       'Last used row before clearing (UsedRange may shrink)
    Dim OwnsProtection(0 To 3)    As Boolean    'Protection set by this case, per output
    Dim OwnsChecks                As Boolean    'Only undo protection set by this case
    Dim SheetNames                As Variant    'Output names in engine clear order
    Dim Widths                    As Variant    'Corresponding output table widths
    Dim i                         As Long       'Output index
    Dim OldSummary                As String     'Successful summary before protection
    Dim StatusFormula             As Variant    'C5 formula/text must not be overwritten
    Dim CaseName                  As String     'Distinct regression identifier
        On Error GoTo Unexpected
        CaseName = "run.protected-output." & ProtectedSheet
        If Len(SecondSheet) > 0 Then CaseName = CaseName & "+" & SecondSheet
        If ProtectChecks Then CaseName = CaseName & "+Checks"
        BeginCase CaseName
        SetState xlCalculationAutomatic, False, False
        SheetNames = Array(SH_TRADECALC, SH_BUCKETS, SH_HEDGING, SH_RESULTS)
        Widths = Array(TC_NCOLS, BK_NCOLS, HS_NCOLS, RS_NCOLS)
        For i = 0 To 3
            If GetSheet(CStr(SheetNames(i))).ProtectContents Then
                Fail CaseName, "test requires initially unprotected " & CStr(SheetNames(i))
                Exit Sub
            End If
        Next i
        If GetSheet(SH_CHECKS).ProtectContents Then
            Fail CaseName, "test requires initially unprotected Checks"
            Exit Sub
        End If
        Result = M_Main.RunSACCR_Silent()
        Check Left$(Result, 9) = "RESULT=OK", "setup run: " & Result
        OldSummary = SafeStr(GetSheet(SH_RESULTS).Range(RUNINFO_CELL).Value)
        StatusFormula = GetSheet(SH_RESULTS).Range(RESULTS_STATUS_CELL).Formula
        For i = 0 To 3
            TableCell(GetSheet(CStr(SheetNames(i))), OutputFirstRow(CStr(SheetNames(i))), 1).Value = "OLD-OUTPUT"
            LastRows(i) = UsedLastRow(GetSheet(CStr(SheetNames(i))))
            If CStr(SheetNames(i)) = ProtectedSheet Or CStr(SheetNames(i)) = SecondSheet Then
                GetSheet(CStr(SheetNames(i))).Protect
                OwnsProtection(i) = True
                'Capture Excel's localized error independently, before the run.
                On Error Resume Next
                ClearOutputBlock GetSheet(CStr(SheetNames(i))), OutputFirstRow(CStr(SheetNames(i))), CLng(Widths(i))
                ProbeNumber = Err.Number
                ProbeSource = Err.Source
                ProbeText = Err.Description
                On Error GoTo Unexpected
                Check ProbeNumber <> 0, "protected clear probe did not fail"
                If PrimaryNumber = 0 Then
                    PrimaryNumber = ProbeNumber
                    PrimarySource = ProbeSource
                    PrimaryText = ProbeText
                End If
                ExpectedDetails(i) = CStr(SheetNames(i)) & " [error " & CStr(ProbeNumber) & _
                            "; source=" & ProbeSource & "]: " & ProbeText
            End If
        Next i
        If ProtectChecks Then
            GetSheet(SH_CHECKS).Protect
            OwnsChecks = True
        End If
        On Error Resume Next
        Result = M_Main.RunSACCR_Silent()
        ErrorNumber = Err.Number
        ErrorSource = Err.Source
        ErrorText = Err.Description
        On Error GoTo Unexpected
        Check ErrorNumber = PrimaryNumber And ErrorNumber <> 0, "primary error number was replaced"
        Check ErrorSource = PrimarySource, "primary error source was replaced: " & ErrorSource
        Check Left$(ErrorText, Len(PrimaryText)) = PrimaryText, "original error description was lost: " & ErrorText
        Check InStr(1, ErrorText, "output cleanup incomplete", vbTextCompare) > 0, "missing incomplete cleanup warning"
        Check InStr(1, ErrorText, "results withdrawn", vbTextCompare) = 0, "reported withdrawal despite retained outputs"
        For i = 0 To 3
            If OwnsProtection(i) Then
                Check SafeStr(TableCell(GetSheet(CStr(SheetNames(i))), OutputFirstRow(CStr(SheetNames(i))), 1).Value) = "OLD-OUTPUT", _
                      CStr(SheetNames(i)) & " protected sentinel unexpectedly changed"
                Check InStr(1, ErrorText, ExpectedDetails(i), vbBinaryCompare) > 0, "missing cleanup detail: " & ExpectedDetails(i)
                If Not ProtectChecks Then Check ChecksMention(ExpectedDetails(i)), "Checks lacks detail: " & ExpectedDetails(i)
            Else
                Check Application.CountA(GetSheet(CStr(SheetNames(i))).Range( _
                      TableCell(GetSheet(CStr(SheetNames(i))), OutputFirstRow(CStr(SheetNames(i))), 1), _
                      TableCell(GetSheet(CStr(SheetNames(i))), LastRows(i), CLng(Widths(i))))) = 0, _
                      CStr(SheetNames(i)) & " retained output cells"
            End If
        Next i
        If OwnsProtection(3) Then
            Check SafeStr(GetSheet(SH_RESULTS).Range(RUNINFO_CELL).Value) = OldSummary, _
                  "protected Results summary unexpectedly changed"
            Check InStr(1, ErrorText, "old results may remain", vbTextCompare) > 0, "retained Results not disclosed"
        Else
            Check InStr(1, SafeStr(GetSheet(SH_RESULTS).Range(RUNINFO_CELL).Value), _
                  "output cleanup incomplete", vbTextCompare) > 0, "Results summary hides cleanup failure"
        End If
        If ProtectChecks Then
            Check InStr(1, ErrorText, "Checks could not be updated", vbTextCompare) > 0, "Checks failure was not reported"
        Else
            Check ChecksMention("output cleanup incomplete"), "Checks did not report incomplete cleanup"
        End If
        Check GetSheet(SH_RESULTS).Range(RESULTS_STATUS_CELL).Formula = StatusFormula, "cleanup overwrote the status formula"
        Check M_Main.ResultsStatus() = "NONE", "failed run retained a current fingerprint"
        Check Len(CORE_Engine.RunInputs) = 0, "failed run retained in-memory run inputs"
        CheckState xlCalculationAutomatic, False, False
        For i = 0 To 3
            If OwnsProtection(i) Then
                GetSheet(CStr(SheetNames(i))).Unprotect
                OwnsProtection(i) = False
            End If
        Next i
        If OwnsChecks Then
            GetSheet(SH_CHECKS).Unprotect
            OwnsChecks = False
        End If
        Result = M_Main.RunSACCR_Silent()
        Check Left$(Result, 9) = "RESULT=OK", "recovery run: " & Result
        Check M_Main.ResultsStatus() = "CURRENT", "recovery did not publish a valid fingerprint"
        Check Not ChecksMention("output cleanup incomplete"), "recovery retained obsolete failure diagnostics"
        For i = 0 To 3
            Check SafeStr(TableCell(GetSheet(CStr(SheetNames(i))), OutputFirstRow(CStr(SheetNames(i))), 1).Value) <> "OLD-OUTPUT", _
                  CStr(SheetNames(i)) & " retained the old sentinel after recovery"
        Next i
        CheckState xlCalculationAutomatic, False, False
        Exit Sub
Unexpected:
        Fail CaseName, "unexpected error " & Err.Number & ": " & Err.Description
        For i = 0 To 3
            If OwnsProtection(i) Then UnprotectOutput CStr(SheetNames(i))
        Next i
        If OwnsChecks Then UnprotectOutput SH_CHECKS
End Sub


Private Sub UnprotectOutput(ByVal SheetName As String)
'
'==============================================================================
' PURPOSE
'   Release protection owned by CaseProtectedOutput after an unexpected error.
'==============================================================================
        On Error GoTo Failed
        GetSheet(SheetName).Unprotect
        Exit Sub
Failed:
        Fail "protected-output", "could not unprotect " & SheetName & ": " & Err.Description
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
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Result         As String     'Result line of a run
    Dim MtmCell        As Range      'MtM of the first trade, edited and restored
    Dim SavedFormula   As Variant    'Its formula

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
        On Error GoTo Unexpected
        BeginCase "results.stale-after-input-change"
        SetState xlCalculationAutomatic, False, False
        Result = M_Main.RunSACCR_Silent()
        Check InStr(Result, "; inputs=") > 0 And InStr(Result, "; inputs=;") = 0, "result line names no inputs: " & Result
        Check M_Main.ResultsStatus() = "CURRENT", "after a run: " & M_Main.ResultsStatus()
        CheckStatusCell "Inputs unchanged", "after a run"

    'Edit one input, as a user would: the formula recalculates by itself.
        Set MtmCell = TableCell(GetSheet(SH_TRADES), TableFirstRow(SH_TRADES), TR_MTM)
        SavedFormula = MtmCell.Formula
        MtmCell.Value = ToDbl(MtmCell.Value, 0#) + 1#
        Check M_Main.ResultsStatus() = "STALE", "after an edit: " & M_Main.ResultsStatus()
        CheckStatusCell "OUT OF DATE", "after an edit"

    'Undoing the edit brings the results back in line.
        MtmCell.Formula = SavedFormula
        Check M_Main.ResultsStatus() = "CURRENT", "after undoing the edit: " & M_Main.ResultsStatus()
        CheckStatusCell "Inputs unchanged", "after undoing the edit"

    'Clear outputs leaves no results; a new run makes them current again.
        M_Main.ClearOutputs
        Check M_Main.ResultsStatus() = "NONE", "after Clear outputs: " & M_Main.ResultsStatus()
        CheckStatusCell "No results", "after Clear outputs"
        Result = M_Main.RunSACCR_Silent()
        Check M_Main.ResultsStatus() = "CURRENT", "after a new run: " & M_Main.ResultsStatus()
        CheckStatusCell "Inputs unchanged", "after a new run"
        CheckState xlCalculationAutomatic, False, False
        Exit Sub

'------------------------------------------------------------------------------
' HANDLE UNEXPECTED ERROR
'------------------------------------------------------------------------------
Unexpected:
        Fail "results.stale-after-input-change", "unexpected error " & Err.Number & ": " & Err.Description
        RestoreCell MtmCell, SavedFormula

End Sub


Private Sub CaseEditedOutput()
'
'==============================================================================
'                               CaseEditedOutput
'------------------------------------------------------------------------------
' PURPOSE
'   An EAD edited on Results after a run, with the inputs unchanged, makes
'   ResultsStatus() EDITED while Results C5, which checks the inputs only,
'   still says the inputs are unchanged; the evidence export refuses it.
'   An input edited as well makes the results STALE. Restoring the value
'   makes them CURRENT, and so does a new run after another edit (#315).
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Result        As String     'Result line of a run or an export
    Dim OutputCell    As Range      'EAD of the first netting set, edited
    Dim SavedOutput   As Variant    'Its value
    Dim MtmCell       As Range      'MtM of the first trade, edited and restored
    Dim SavedMtm      As Variant    'Its formula
    Dim ExportPath    As String     'Export file, never written

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
        On Error GoTo Unexpected
        BeginCase "results.edited-output"
        SetState xlCalculationAutomatic, False, False
        Result = M_Main.RunSACCR_Silent()
        Check M_Main.ResultsStatus() = "CURRENT", "after a run: " & M_Main.ResultsStatus()

    'Edit an output, as a user could: C5 checks the inputs only.
        Set OutputCell = TableCell(GetSheet(SH_RESULTS), OutputFirstRow(SH_RESULTS), RS_EAD)
        SavedOutput = OutputCell.Value2
        OutputCell.Value2 = ToDbl(SavedOutput, 0#) + 1#
        Check M_Main.ResultsStatus() = "EDITED", "after an output edit: " & M_Main.ResultsStatus()
        CheckStatusCell "Inputs unchanged", "after an output edit"
        ExportPath = Environ$("TEMP")
        If Len(ExportPath) = 0 Then
            ExportPath = ThisWorkbook.Path
        End If
        ExportPath = ExportPath & Application.PathSeparator & "K-SACCR_evidence_test_edited.csv"
        Result = M_Main.ExportEvidence_Silent(ExportPath)
        Check Left$(Result, 14) = "RESULT=REFUSED" And InStr(Result, "output cell") > 0, _
              "export after an output edit: " & Result

    'An input edited as well: the inputs decide first.
        Set MtmCell = TableCell(GetSheet(SH_TRADES), TableFirstRow(SH_TRADES), TR_MTM)
        SavedMtm = MtmCell.Formula
        MtmCell.Value = ToDbl(MtmCell.Value, 0#) + 1#
        Check M_Main.ResultsStatus() = "STALE", "after an output and an input edit: " & M_Main.ResultsStatus()
        MtmCell.Formula = SavedMtm

    'Restoring the output, or a new run, makes the results current again.
        OutputCell.Value2 = SavedOutput
        Check M_Main.ResultsStatus() = "CURRENT", "after restoring the output: " & M_Main.ResultsStatus()
        OutputCell.Value2 = ToDbl(SavedOutput, 0#) + 1#
        Result = M_Main.RunSACCR_Silent()
        Check M_Main.ResultsStatus() = "CURRENT", "after a new run: " & M_Main.ResultsStatus()
        CheckState xlCalculationAutomatic, False, False
        Exit Sub

'------------------------------------------------------------------------------
' HANDLE UNEXPECTED ERROR
'------------------------------------------------------------------------------
Unexpected:
        Fail "results.edited-output", "unexpected error " & Err.Number & ": " & Err.Description
        RestoreCell MtmCell, SavedMtm
        If Not OutputCell Is Nothing Then
            OutputCell.Value2 = SavedOutput
        End If

End Sub


Private Sub CaseStaleValueType()
'
'==============================================================================
'                              CaseStaleValueType
'------------------------------------------------------------------------------
' PURPOSE
'   After a run, the MtM of the first trade replaced by the same digits as
'   text makes the results STALE, since the engine reads the two
'   differently (#185); putting the number back makes them CURRENT. Two
'   error values give different fingerprints. A number format on the cell
'   leaves the results CURRENT (#197).
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim MtmCell            As Range      'MtM of the first trade
    Dim SavedFormula       As Variant    'Its formula
    Dim SavedFormat        As Variant    'Its number format
    Dim NaFingerprint      As String     'Fingerprint with #N/A in the cell
    Dim ValueFingerprint   As String     'Fingerprint with #VALUE! in the cell

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
        On Error GoTo Unexpected
        BeginCase "results.stale-value-type"
        SetState xlCalculationAutomatic, False, False
        Set MtmCell = TableCell(GetSheet(SH_TRADES), TableFirstRow(SH_TRADES), TR_MTM)
        SavedFormula = MtmCell.Formula
        SavedFormat = MtmCell.NumberFormat
        Check VarType(MtmCell.Value2) = vbDouble, "the first trade's MtM is not a number"
        M_Main.RunSACCR_Silent
        Check M_Main.ResultsStatus() = "CURRENT", "after a run: " & M_Main.ResultsStatus()

    'The same digits, entered as text.
        MtmCell.Value = "'" & CStr(MtmCell.Value2)
        Check VarType(MtmCell.Value2) = vbString, "the MtM was not stored as text"
        Check M_Main.ResultsStatus() = "STALE", "number as text: " & M_Main.ResultsStatus()
        CheckStatusCell "OUT OF DATE", "number as text"
        MtmCell.Formula = SavedFormula
        Check M_Main.ResultsStatus() = "CURRENT", "number restored: " & M_Main.ResultsStatus()
        CheckStatusCell "Inputs unchanged", "number restored"

    'Two error values.
        MtmCell.Value = CVErr(xlErrNA)
        NaFingerprint = CORE_Engine.InputFingerprint()
        MtmCell.Value = CVErr(xlErrValue)
        ValueFingerprint = CORE_Engine.InputFingerprint()
        Check NaFingerprint <> ValueFingerprint, "#N/A and #VALUE! give the same fingerprint " & NaFingerprint
        MtmCell.Formula = SavedFormula

    'Formatting only.
        MtmCell.NumberFormat = "0.000%"
        Check M_Main.ResultsStatus() = "CURRENT", "after a number format: " & M_Main.ResultsStatus()
        CheckStatusCell "Inputs unchanged", "after a number format"
        MtmCell.NumberFormat = SavedFormat
        CheckState xlCalculationAutomatic, False, False
        Exit Sub

'------------------------------------------------------------------------------
' HANDLE UNEXPECTED ERROR
'------------------------------------------------------------------------------
Unexpected:
        Fail mCurrentCase, "unexpected error " & Err.Number & ": " & Err.Description
        RestoreCell MtmCell, SavedFormula
        On Error Resume Next
        If Not MtmCell Is Nothing Then MtmCell.NumberFormat = SavedFormat

End Sub


Private Sub CaseStaleHeaders()
'
'==============================================================================
'                               CaseStaleHeaders
'------------------------------------------------------------------------------
' PURPOSE
'   After a run, a renamed header on Trades or NettingSets makes the
'   results STALE, since the run would stop (#35), and the original header
'   makes them CURRENT again. A bold header, or one changed only in case
'   and spacing, which the schema check accepts, leaves them CURRENT
'   (#197).
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim HeaderCell   As Range      'Header edited, Nothing once restored
    Dim SavedBold    As Variant    'Its bold setting
    Dim SavedText    As Variant    'Its formula

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
        On Error GoTo Unexpected
        BeginCase "results.stale-headers"
        SetState xlCalculationAutomatic, False, False
        M_Main.RunSACCR_Silent
        Check M_Main.ResultsStatus() = "CURRENT", "after a run: " & M_Main.ResultsStatus()
        EditAndUndo TableCell(GetSheet(SH_TRADES), TableHeaderRow(SH_TRADES), TR_MTM), "Trades MtM header"
        EditAndUndo TableCell(GetSheet(SH_NS), TableHeaderRow(SH_NS), NS_ID), "NettingSets ID header"

    'Formatting only.
        Set HeaderCell = TableCell(GetSheet(SH_TRADES), TableHeaderRow(SH_TRADES), TR_MTM)
        SavedBold = HeaderCell.Font.Bold
        HeaderCell.Font.Bold = Not CBool(SavedBold)
        Check M_Main.ResultsStatus() = "CURRENT", "after a bold header: " & M_Main.ResultsStatus()
        CheckStatusCell "Inputs unchanged", "after a bold header"
        HeaderCell.Font.Bold = SavedBold

    'Case and spacing only: "MtM" as " m tM ".
        SavedText = HeaderCell.Formula
        HeaderCell.Value = " " & LCase$(Left$(CStr(SavedText), 1)) & " " & Mid$(CStr(SavedText), 2) & " "
        Check M_Main.ResultsStatus() = "CURRENT", "after a case and spacing edit: " & M_Main.ResultsStatus()
        CheckStatusCell "Inputs unchanged", "after a case and spacing edit"
        HeaderCell.Formula = SavedText
        Set HeaderCell = Nothing
        CheckState xlCalculationAutomatic, False, False
        Exit Sub

'------------------------------------------------------------------------------
' HANDLE UNEXPECTED ERROR
'------------------------------------------------------------------------------
Unexpected:
        Fail mCurrentCase, "unexpected error " & Err.Number & ": " & Err.Description
        On Error Resume Next
        If Not HeaderCell Is Nothing Then
            HeaderCell.Font.Bold = SavedBold
            If Not IsEmpty(SavedText) Then HeaderCell.Formula = SavedText
        End If

End Sub


Private Sub CaseStaleDateSystem()
'
'==============================================================================
'                              CaseStaleDateSystem
'------------------------------------------------------------------------------
' PURPOSE
'   After a run, switching the workbook to the other date system makes the
'   results STALE, since the run stops in a 1904 workbook (#185), and
'   switching back makes them CURRENT (#197).
'
' METHOD
'   Excel need not recalculate Results C5 when only the date system
'   changes, because no input cell changed; the case recalculates the cell
'   itself, as the next input edit would.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim SavedDateSystem   As Boolean    'Date system of the workbook
    Dim Switched          As Boolean    'The date system is switched

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
        On Error GoTo Unexpected
        BeginCase "results.stale-date-system"
        SetState xlCalculationAutomatic, False, False
        SavedDateSystem = ThisWorkbook.Date1904
        M_Main.RunSACCR_Silent
        Check M_Main.ResultsStatus() = "CURRENT", "after a run: " & M_Main.ResultsStatus()
        ThisWorkbook.Date1904 = Not SavedDateSystem
        Switched = True
        Check M_Main.ResultsStatus() = "STALE", "date system switched: " & M_Main.ResultsStatus()
        GetSheet(SH_RESULTS).Range(RESULTS_STATUS_CELL).Calculate
        CheckStatusCell "OUT OF DATE", "date system switched"
        ThisWorkbook.Date1904 = SavedDateSystem
        Switched = False
        Check M_Main.ResultsStatus() = "CURRENT", "date system restored: " & M_Main.ResultsStatus()
        GetSheet(SH_RESULTS).Range(RESULTS_STATUS_CELL).Calculate
        CheckStatusCell "Inputs unchanged", "date system restored"
        CheckState xlCalculationAutomatic, False, False
        Exit Sub

'------------------------------------------------------------------------------
' HANDLE UNEXPECTED ERROR
'------------------------------------------------------------------------------
Unexpected:
        Fail mCurrentCase, "unexpected error " & Err.Number & ": " & Err.Description
        On Error Resume Next
        If Switched Then ThisWorkbook.Date1904 = SavedDateSystem
        If Err.Number <> 0 Then
            Fail mCurrentCase, "could not restore the date system: " & Err.Description
        End If

End Sub


Private Sub CaseManualRecalculation()
'
'==============================================================================
'                           CaseManualRecalculation
'------------------------------------------------------------------------------
' PURPOSE
'   In manual calculation, a formula input whose precedent changed since
'   the last recalculation is brought up to date before the run reads it
'   (#195). The first trade's MtM becomes a formula on its own Comment
'   cell, which the engine does not read; the Comment changes from 1,000 to
'   2,000 without a recalculation, and the run must use an MtM of 2,000
'   (the trade's MtM currency is the reporting currency). The results are
'   then current, and the caller's manual mode is kept.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Result      As String     'Result line of the run
    Dim MtmCell     As Range      'MtM of the first trade, made a formula
    Dim NoteCell    As Range      'Comment of the first trade, its precedent
    Dim SavedMtm    As Variant    'Formula of MtmCell before the case
    Dim SavedNote   As Variant    'Formula of NoteCell before the case
    Dim CalcRow     As Long       'TradeCalc row of the first trade

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
        On Error GoTo Unexpected
        BeginCase "run.manual-recalculates-inputs"
        Set MtmCell = TableCell(GetSheet(SH_TRADES), TableFirstRow(SH_TRADES), TR_MTM)
        Set NoteCell = TableCell(GetSheet(SH_TRADES), TableFirstRow(SH_TRADES), TR_COMMENT)
        SavedMtm = MtmCell.Formula
        SavedNote = NoteCell.Formula
        SetState xlCalculationManual, False, False
        NoteCell.Value = 1000
        MtmCell.Formula = "=" & NoteCell.Address(False, False)
        Application.Calculate
        NoteCell.Value = 2000
        Check ToDbl(MtmCell.Value, 0#) = 1000#, "set-up: the MtM formula was recalculated before the run"
        Result = M_Main.RunSACCR_Silent()

    'The first trade is the first TradeCalc row; column 9 is its MtM in
    'the reporting currency.
        CalcRow = OutputFirstRow(SH_TRADECALC)
        Check UTxt(TableCell(GetSheet(SH_TRADECALC), CalcRow, 1).Value) = _
              UTxt(TableCell(GetSheet(SH_TRADES), TableFirstRow(SH_TRADES), TR_ID).Value), _
              "the first TradeCalc row is not the first trade"
        Check ToDbl(TableCell(GetSheet(SH_TRADECALC), CalcRow, 9).Value, 0#) = 2000#, _
              "the run read an MtM of " & TableCell(GetSheet(SH_TRADECALC), CalcRow, 9).Value & ", expected 2000"
        Check M_Main.ResultsStatus() = "CURRENT", "after the run: " & M_Main.ResultsStatus()
        CheckState xlCalculationManual, False, False
        RestoreCell MtmCell, SavedMtm
        RestoreCell NoteCell, SavedNote
        Exit Sub

'------------------------------------------------------------------------------
' HANDLE UNEXPECTED ERROR
'------------------------------------------------------------------------------
Unexpected:
        Fail "run.manual-recalculates-inputs", "unexpected error " & Err.Number & ": " & Err.Description
        RestoreCell MtmCell, SavedMtm
        RestoreCell NoteCell, SavedNote

End Sub


Private Sub CaseRecalculationPending()
'
'==============================================================================
'                           CaseRecalculationPending
'------------------------------------------------------------------------------
' PURPOSE
'   When the recalculation before a run does not finish (#195), simulated
'   through gTestFault = "recalc", the run raises ERR_RECALC_PENDING, the
'   previous results are no longer reported as current, Excel is restored,
'   and the next run works.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Result        As String    'Result line of a run
    Dim ErrorNumber   As Long      'Error raised by the stopped run
    Dim ErrorText     As String    'Its description

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
        On Error GoTo Unexpected
        BeginCase "run.recalculation-pending"
        SetState xlCalculationManual, False, False
        Result = M_Main.RunSACCR_Silent()
        Check M_Main.ResultsStatus() = "CURRENT", "after a run: " & M_Main.ResultsStatus()
        M_Main.gTestFault = "recalc"
        On Error Resume Next
        Result = M_Main.RunSACCR_Silent()
        ErrorNumber = Err.Number
        ErrorText = Err.Description
        On Error GoTo Unexpected
        M_Main.gTestFault = ""
        Check ErrorNumber = ERR_RECALC_PENDING, "expected ERR_RECALC_PENDING, got " & ErrorNumber & ": " & ErrorText
        Check M_Main.ResultsStatus() <> "CURRENT", "results still reported as current after the stopped run"
        CheckState xlCalculationManual, False, False

    'The next run completes and its results are current again.
        Result = M_Main.RunSACCR_Silent()
        Check Left$(Result, 9) = "RESULT=OK", "follow-up run: " & Result
        Check M_Main.ResultsStatus() = "CURRENT", "after the follow-up run: " & M_Main.ResultsStatus()
        CheckState xlCalculationManual, False, False
        Exit Sub

'------------------------------------------------------------------------------
' HANDLE UNEXPECTED ERROR
'------------------------------------------------------------------------------
Unexpected:
        M_Main.gTestFault = ""
        Fail "run.recalculation-pending", "unexpected error " & Err.Number & ": " & Err.Description

End Sub


Private Sub CaseRecalculationError()
'
'==============================================================================
'                            CaseRecalculationError
'------------------------------------------------------------------------------
' PURPOSE
'   When the recalculation before a run raises an error (#204), simulated
'   through gTestFault = "recalc-error", the previous results are no longer
'   reported as current, even though a formula input still holds the value
'   the last run used. The first trade's MtM becomes a formula on its own
'   Comment cell, as in CaseManualRecalculation; after a run in manual
'   calculation the Comment changes without a recalculation, so the cached
'   MtM still equals the one the run fingerprinted. The stopped run must
'   raise the injected error itself, restore Excel, and leave the next run
'   working. The status cell on Results must follow each step although
'   calculation stays manual.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Result        As String     'Result line of a run
    Dim ErrorNumber   As Long       'Error raised by the stopped run
    Dim ErrorText     As String     'Its description
    Dim MtmCell       As Range      'MtM of the first trade, made a formula
    Dim NoteCell      As Range      'Comment of the first trade, its precedent
    Dim SavedMtm      As Variant    'Formula of MtmCell before the case
    Dim SavedNote     As Variant    'Formula of NoteCell before the case

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
        On Error GoTo Unexpected
        BeginCase "run.recalculation-error"
        Set MtmCell = TableCell(GetSheet(SH_TRADES), TableFirstRow(SH_TRADES), TR_MTM)
        Set NoteCell = TableCell(GetSheet(SH_TRADES), TableFirstRow(SH_TRADES), TR_COMMENT)
        SavedMtm = MtmCell.Formula
        SavedNote = NoteCell.Formula
        SetState xlCalculationManual, False, False
        NoteCell.Value = 1000
        MtmCell.Formula = "=" & NoteCell.Address(False, False)
        Result = M_Main.RunSACCR_Silent()
        Check Left$(Result, 9) = "RESULT=OK", "set-up run: " & Result
        Check M_Main.ResultsStatus() = "CURRENT", "after a run: " & M_Main.ResultsStatus()
        CheckStatusCell "Inputs unchanged", "after a run in manual calculation"

    'The precedent changes but nothing recalculates: the cached MtM is the
    'value the run used, so only the withdrawn fingerprint can show that
    'the results may be out of date.
        NoteCell.Value = 2000
        Check ToDbl(MtmCell.Value, 0#) = 1000#, "set-up: the MtM formula was recalculated too early"
        M_Main.gTestFault = "recalc-error"
        On Error Resume Next
        Result = M_Main.RunSACCR_Silent()
        ErrorNumber = Err.Number
        ErrorText = Err.Description
        On Error GoTo Unexpected
        M_Main.gTestFault = ""
        Check ErrorNumber = ERR_INJECTED_FAULT, "expected ERR_INJECTED_FAULT, got " & ErrorNumber & ": " & ErrorText
        Check ToDbl(MtmCell.Value, 0#) = 1000#, "the cached MtM changed although the recalculation failed"
        Check M_Main.ResultsStatus() <> "CURRENT", "results still reported as current after the failed recalculation"
        CheckStatusCell "No results", "after the failed recalculation"
        CheckState xlCalculationManual, False, False

    'The next run recalculates, reads the new MtM and is current again.
        Result = M_Main.RunSACCR_Silent()
        Check Left$(Result, 9) = "RESULT=OK", "follow-up run: " & Result
        Check ToDbl(TableCell(GetSheet(SH_TRADECALC), OutputFirstRow(SH_TRADECALC), 9).Value, 0#) = 2000#, _
              "the follow-up run read an MtM of " & TableCell(GetSheet(SH_TRADECALC), OutputFirstRow(SH_TRADECALC), 9).Value & ", expected 2000"
        Check M_Main.ResultsStatus() = "CURRENT", "after the follow-up run: " & M_Main.ResultsStatus()
        CheckStatusCell "Inputs unchanged", "after the follow-up run"
        CheckState xlCalculationManual, False, False
        RestoreCell MtmCell, SavedMtm
        RestoreCell NoteCell, SavedNote
        Exit Sub

'------------------------------------------------------------------------------
' HANDLE UNEXPECTED ERROR
'------------------------------------------------------------------------------
Unexpected:
        M_Main.gTestFault = ""
        Fail "run.recalculation-error", "unexpected error " & Err.Number & ": " & Err.Description
        RestoreCell MtmCell, SavedMtm
        RestoreCell NoteCell, SavedNote

End Sub


Private Sub CaseStaleParams( _
    ByVal Shifted As Boolean)
'
'==============================================================================
'                               CaseStaleParams
'------------------------------------------------------------------------------
' PURPOSE
'   After a run, an edit to the Regimes field of a factor or its header,
'   to a parameter value or to an FX rate makes the results STALE and
'   Results C5 OUT OF DATE, and undoing it makes them CURRENT again; a
'   changed parameter description or cell colour leaves them CURRENT
'   (#152). With Shifted True a column is first inserted before the Params
'   tables, which moves every Params input and workbook name one column to
'   the right, as the issue describes; the run must then give the same
'   result line, fingerprint included, as the unshifted run, and the
'   column is deleted again at the end.
'
' INPUTS
'   Shifted: run the case with a column inserted before the Params tables.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ParamsSheet        As Worksheet    'Params sheet
    Dim Inserted           As Boolean      'The extra column is on Params
    Dim Result             As String       'Result line of the run
    Dim FactorColumn       As Long         'Table column of the factor keys
    Dim FxColumn           As Long         'Table column of the FX currencies
    Dim DescriptionCell    As Range        'Description edited, Nothing once restored
    Dim SavedDescription   As Variant      'Its formula
    Dim ColorCell          As Range        'Cell recoloured, Nothing once restored
    Dim SavedColor         As Variant      'Its fill colour

'------------------------------------------------------------------------------
' LAYOUT AND RUN
'------------------------------------------------------------------------------
        On Error GoTo Unexpected
        BeginCase IIf(Shifted, "results.stale-params-shifted", "results.stale-params")
        SetState xlCalculationAutomatic, False, False
        Set ParamsSheet = GetSheet(SH_PARAMS)
        If Shifted Then
            ParamsSheet.Columns(COL_OFFSET + 1).Insert
            Inserted = True
        End If
        FactorColumn = NamedTableCol(NM_SF_TABLE)
        FxColumn = NamedTableCol(NM_FX_TABLE)
        Check FactorColumn = IIf(Shifted, 2, 1), "factor table in table column " & FactorColumn
        Result = M_Main.RunSACCR_Silent()
        Check M_Main.ResultsStatus() = "CURRENT", "after a run: " & M_Main.ResultsStatus() & " (" & Result & ")"
        CheckStatusCell "Inputs unchanged", "after a run"
        If Shifted Then
            Check Result = mParamsRunLine, "shifted run differs: " & Result & " | unshifted: " & mParamsRunLine
        Else
            mParamsRunLine = Result
        End If

'------------------------------------------------------------------------------
' INPUTS: STALE, THEN CURRENT AGAIN
'------------------------------------------------------------------------------
    'Regimes is the eighth factor column: K once shifted, outside the
    'C:J block fingerprinted before #152.
        EditAndUndo TableCell(ParamsSheet, NamedRow(NM_SF_TABLE) + 1, FactorColumn + 7), "factor Regimes"
        EditAndUndo TableCell(ParamsSheet, NamedRow(NM_SF_TABLE), FactorColumn + 7), "factor Regimes header"
        EditAndUndo NamedCell(PRM_ALPHA), "parameter Alpha"
        EditAndUndo TableCell(ParamsSheet, NamedRow(NM_FX_TABLE) + 1, FxColumn + 2), "FX rate"

'------------------------------------------------------------------------------
' NOT INPUTS: STILL CURRENT
'------------------------------------------------------------------------------
    'The description beside a parameter code, then a fill colour.
        Set DescriptionCell = TableCell(ParamsSheet, NamedRow(PRM_ALPHA), ParamsCodeCol() + 1)
        SavedDescription = DescriptionCell.Formula
        DescriptionCell.Value = SafeStr(DescriptionCell.Value) & " (edited)"
        Check M_Main.ResultsStatus() = "CURRENT", "after a description edit: " & M_Main.ResultsStatus()
        CheckStatusCell "Inputs unchanged", "after a description edit"
        DescriptionCell.Formula = SavedDescription
        Set DescriptionCell = Nothing
        Set ColorCell = NamedCell(PRM_ALPHA)
        SavedColor = ColorCell.Interior.Color
        ColorCell.Interior.Color = RGB(255, 0, 0)
        Check M_Main.ResultsStatus() = "CURRENT", "after a colour change: " & M_Main.ResultsStatus()
        CheckStatusCell "Inputs unchanged", "after a colour change"
        ColorCell.Interior.Color = SavedColor
        Set ColorCell = Nothing

'------------------------------------------------------------------------------
' RESTORE THE LAYOUT
'------------------------------------------------------------------------------
        If Inserted Then
            ParamsSheet.Columns(COL_OFFSET + 1).Delete
            Inserted = False
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
        If Not DescriptionCell Is Nothing Then DescriptionCell.Formula = SavedDescription
        If Not ColorCell Is Nothing Then ColorCell.Interior.Color = SavedColor
        If Inserted Then ParamsSheet.Columns(COL_OFFSET + 1).Delete
        If Err.Number <> 0 Then
            Fail mCurrentCase, "could not restore Params: " & Err.Description
        End If

End Sub


Private Sub EditAndUndo( _
    ByVal EditedCell As Range, _
    ByVal InputLabel As String)
'
'==============================================================================
'                                 EditAndUndo
'------------------------------------------------------------------------------
' PURPOSE
'   Change an input cell, check that the results are STALE and Results C5
'   says OUT OF DATE, put the cell back, and check that they are CURRENT
'   again. A number is raised by 1; any other value gets an X appended.
'
' INPUTS
'   EditedCell: the input cell to change.
'   InputLabel: what the cell is, for the failure messages.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim SavedFormula   As Variant    'The cell's formula

'------------------------------------------------------------------------------
' EDIT AND UNDO
'------------------------------------------------------------------------------
        On Error GoTo Unexpected
        SavedFormula = EditedCell.Formula
        If VarType(EditedCell.Value2) = vbDouble Then
            EditedCell.Value = EditedCell.Value2 + 1#
        Else
            EditedCell.Value = SafeStr(EditedCell.Value2) & "X"
        End If
        Check M_Main.ResultsStatus() = "STALE", InputLabel & " edited: " & M_Main.ResultsStatus()
        CheckStatusCell "OUT OF DATE", InputLabel & " edited"
        EditedCell.Formula = SavedFormula
        Check M_Main.ResultsStatus() = "CURRENT", InputLabel & " restored: " & M_Main.ResultsStatus()
        CheckStatusCell "Inputs unchanged", InputLabel & " restored"
        Exit Sub

'------------------------------------------------------------------------------
' HANDLE UNEXPECTED ERROR
'------------------------------------------------------------------------------
Unexpected:
        Fail mCurrentCase, InputLabel & ": unexpected error " & Err.Number & ": " & Err.Description
        RestoreCell EditedCell, SavedFormula

End Sub


Private Sub CheckStatusCell( _
    ByVal ExpectedStart As String, _
    ByVal Moment As String)
'
'==============================================================================
'                               CheckStatusCell
'------------------------------------------------------------------------------
' PURPOSE
'   Check that the status formula in Results C5 starts with ExpectedStart,
'   after Excel has recalculated what changed.
'
' INPUTS
'   ExpectedStart: the start of the expected text.
'   Moment: when the check is made, for the failure message.
'
' METHOD
'   DoEvents and Application.Calculate let Excel recalculate the cells an
'   edit or a changed workbook name has marked as changed, as Excel does
'   when a macro returns to a user, or on F9 in manual calculation; without
'   them the cell could still show the previous state. The cell itself is
'   not forced to recalculate (Range.Calculate), so the check still shows
'   that Excel marks the status as changed by itself.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Shown   As String    'Text in Results C5

'------------------------------------------------------------------------------
' CHECK
'------------------------------------------------------------------------------
        DoEvents
        Application.Calculate
        Shown = SafeStr(GetSheet(SH_RESULTS).Range(RESULTS_STATUS_CELL).Value)
        Check Left$(Shown, Len(ExpectedStart)) = ExpectedStart, _
              "Results C5 " & Moment & ": expected """ & ExpectedStart & "..."", found """ & Shown & """"

End Sub


Private Sub RestoreCell( _
    ByVal EditedCell As Range, _
    ByVal SavedFormula As Variant)
'
'==============================================================================
'                                 RestoreCell
'------------------------------------------------------------------------------
' PURPOSE
'   Put back a cell edited by a case (CaseStaleResults, CaseEditedOutput,
'   CaseStaleValueType, CaseManualRecalculation, CaseRecalculationError or
'   EditAndUndo), at its end or after an unexpected error.
'
' INPUTS
'   EditedCell: the cell, or Nothing when the case had not reached it.
'   SavedFormula: its formula before the edit.
'
' ERROR POLICY
'   Contained: a failure is reported as a test failure of the current
'   case.
'
' UPDATED
'   2026-10-17
'==============================================================================
'
        On Error GoTo Failed
        If Not EditedCell Is Nothing Then
            EditedCell.Formula = SavedFormula
        End If
        Exit Sub

Failed:
        Fail mCurrentCase, "could not restore the edited input: " & Err.Description

End Sub


Private Function ChecksMention( _
    ByVal Fragment As String) _
    As Boolean
'
'==============================================================================
'                                ChecksMention
'------------------------------------------------------------------------------
' PURPOSE
'   Tell whether a message on the Checks sheet contains Fragment, ignoring
'   case.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ChecksSheet   As Worksheet    'Checks sheet
    Dim r             As Long         'Row being read

'------------------------------------------------------------------------------
' SEARCH
'------------------------------------------------------------------------------
    'Checks column 5 holds the message.
        Set ChecksSheet = GetSheet(SH_CHECKS)
        For r = OutputFirstRow(SH_CHECKS) To UsedLastRow(ChecksSheet)
            If InStr(1, SafeStr(TableCell(ChecksSheet, r, 5).Value), Fragment, vbTextCompare) > 0 Then
                ChecksMention = True
                Exit Function
            End If
        Next r

End Function


Private Sub CaseOutputHeaderNames()
'
'==============================================================================
'                            CaseOutputHeaderNames
'------------------------------------------------------------------------------
' PURPOSE
'   Every output table is located through the workbook name of its first
'   header cell. A row inserted above the HedgingSets table moves the name,
'   and the next run writes directly below the moved headers, leaving the
'   inserted row empty; deleting the row moves the table back. A broken
'   Checks_Header stops OutputFirstRow with ERR_OUTPUT_LAYOUT, naming it.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim SheetNames         As Variant      'Output sheets
    Dim FirstHeaders       As Variant      'Their first headers
    Dim i                  As Long         'Output sheet
    Dim LocatedRow         As Long         'A header row, or OutputFirstRow's result
    Dim HedgingSheet       As Worksheet    'HedgingSets
    Dim InsertedRow        As Long         'Row inserted above HedgingSets; 0 when none
    Dim ChecksHeaderName   As Name         'Checks_Header, Nothing once restored
    Dim SavedReference     As String       'Its reference
    Dim ErrorNumber        As Long         'Error raised by OutputFirstRow
    Dim ErrorText          As String       'Its description

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
        On Error GoTo Unexpected
        BeginCase "layout.output-header-names"
        SetState xlCalculationAutomatic, False, False

    'Each name is on the first header of its table, in column C.
        SheetNames = Array(SH_RESULTS, SH_HEDGING, SH_BUCKETS, SH_TRADECALC, SH_CHECKS, SH_RECON, SH_TRADE_RECON)
        FirstHeaders = Array("NettingSetID", "NettingSetID", "NettingSetID", "TradeID", "Severity", "NettingSetID", "TradeID")
        For i = 0 To UBound(SheetNames)
            LocatedRow = TableHeaderRow(CStr(SheetNames(i)))
            Check LocatedRow > 0, CStr(SheetNames(i)) & ": " & TableHeaderName(CStr(SheetNames(i))) & " does not locate the table"
            If LocatedRow > 0 Then
                Check SafeStr(TableCell(GetSheet(CStr(SheetNames(i))), LocatedRow, 1).Value) = CStr(FirstHeaders(i)), _
                      CStr(SheetNames(i)) & ": the named cell is not the first header"
            End If
        Next i

    'A row inserted above HedgingSets moves the table with its name.
        Set HedgingSheet = GetSheet(SH_HEDGING)
        LocatedRow = TableHeaderRow(SH_HEDGING)
        HedgingSheet.Rows(LocatedRow).Insert
        InsertedRow = LocatedRow
        Check TableHeaderRow(SH_HEDGING) = LocatedRow + 1, "the HedgingSets header name did not move with the inserted row"
        M_Main.RunSACCR_Silent
        Check SafeStr(TableCell(HedgingSheet, LocatedRow + 1, 1).Value) = "NettingSetID", "the HedgingSets headers were overwritten"
        Check Not IsBlankCell(TableCell(HedgingSheet, LocatedRow + 2, 1).Value), "HedgingSets was not written below the moved headers"
        Check Application.CountA(HedgingSheet.Rows(LocatedRow)) = 0, "the run wrote into the row inserted above the headers"
        HedgingSheet.Rows(LocatedRow).Delete
        InsertedRow = 0
        Check TableHeaderRow(SH_HEDGING) = LocatedRow, "the HedgingSets header name did not move back"
        M_Main.RunSACCR_Silent
        Check Not IsBlankCell(TableCell(HedgingSheet, LocatedRow + 1, 1).Value), "HedgingSets was not written below its headers"

    'A broken name stops OutputFirstRow, naming the name.
        Set ChecksHeaderName = ThisWorkbook.Names(NM_CK_HEADER)
        SavedReference = ChecksHeaderName.RefersTo
        ChecksHeaderName.RefersTo = "=#REF!"
        On Error Resume Next
        LocatedRow = OutputFirstRow(SH_CHECKS)
        ErrorNumber = Err.Number
        ErrorText = Err.Description
        On Error GoTo Unexpected
        ChecksHeaderName.RefersTo = SavedReference
        Set ChecksHeaderName = Nothing
        Check ErrorNumber = ERR_OUTPUT_LAYOUT, "a broken " & NM_CK_HEADER & " raised error " & ErrorNumber
        Check InStr(1, ErrorText, NM_CK_HEADER, vbBinaryCompare) > 0, "the error does not name " & NM_CK_HEADER & ": " & ErrorText
        Check OutputFirstRow(SH_CHECKS) = TableHeaderRow(SH_CHECKS) + 1, "Checks_Header was not restored"
        CheckState xlCalculationAutomatic, False, False
        Exit Sub

'------------------------------------------------------------------------------
' HANDLE UNEXPECTED ERROR
'------------------------------------------------------------------------------
Unexpected:
        Fail mCurrentCase, "unexpected error " & Err.Number & ": " & Err.Description
        On Error Resume Next
        If Not ChecksHeaderName Is Nothing Then ChecksHeaderName.RefersTo = SavedReference
        If InsertedRow > 0 Then HedgingSheet.Rows(InsertedRow).Delete

End Sub


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
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Result   As String    'Result line

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
        On Error GoTo Unexpected
        BeginCase "silent.flag-preserved"
        SetState xlCalculationAutomatic, True, True
        M_Main.gSilent = False
        Result = M_Main.RunSACCR_Silent()
        Check M_Main.gSilent = False, "gSilent was False before a successful run"
        M_Main.gSilent = True
        M_Main.gTestFault = "operation"
        On Error Resume Next
        Result = M_Main.RunSACCR_Silent()
        On Error GoTo Unexpected
        M_Main.gTestFault = ""
        Check M_Main.gSilent = True, "gSilent was True before a failed run"
        M_Main.gSilent = False
        M_Main.gTestFault = "operation"
        On Error Resume Next
        Result = M_Main.RunSACCR_Silent()
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
    ByVal CaseName As String)
'
'==============================================================================
'                                  BeginCase
'------------------------------------------------------------------------------
' PURPOSE
'   Count a case and print its CASE line.
'
' UPDATED
'   2026-10-17
'==============================================================================
'
        mCaseCount = mCaseCount + 1
        mCurrentCase = CaseName
        TEST_Suite.Say "CASE=" & CaseName

End Sub


Private Sub SetState( _
    ByVal CalcMode As XlCalculation, _
    ByVal EventsOn As Boolean, _
    ByVal ScreenOn As Boolean)
'
'==============================================================================
'                                   SetState
'------------------------------------------------------------------------------
' PURPOSE
'   Put the three settings into a known state before an operation.
'
' UPDATED
'   2026-10-17
'==============================================================================
'
        Application.Calculation = CalcMode
        Application.EnableEvents = EventsOn
        Application.ScreenUpdating = ScreenOn

End Sub


Private Sub CheckState( _
    ByVal CalcMode As XlCalculation, _
    ByVal EventsOn As Boolean, _
    ByVal ScreenOn As Boolean)
'
'==============================================================================
'                                  CheckState
'------------------------------------------------------------------------------
' PURPOSE
'   Check each of the three settings against its expected value.
'
' UPDATED
'   2026-10-17
'==============================================================================
'
        Check Application.Calculation = CalcMode, _
              "Calculation expected " & CalcMode & ", found " & Application.Calculation
        Check Application.EnableEvents = EventsOn, _
              "EnableEvents expected " & EventsOn & ", found " & Application.EnableEvents
        Check Application.ScreenUpdating = ScreenOn, _
              "ScreenUpdating expected " & ScreenOn & ", found " & Application.ScreenUpdating

End Sub


Private Sub Check( _
    ByVal Passed As Boolean, _
    ByVal Detail As String)
'
'==============================================================================
'                                    Check
'------------------------------------------------------------------------------
' PURPOSE
'   Count one check and record it when it fails.
'
' UPDATED
'   2026-10-17
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
'                                     Fail
'------------------------------------------------------------------------------
' PURPOSE
'   Count a failure and print it.
'
' UPDATED
'   2026-10-17
'==============================================================================
'
        mFailureCount = mFailureCount + 1
        TEST_Suite.Say "FAILURE=" & CaseName & ": " & Detail

End Sub
