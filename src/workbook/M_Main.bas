Attribute VB_Name = "M_Main"
'==============================================================================
' MODULE: M_Main
'------------------------------------------------------------------------------
' PURPOSE
'   Provide the macros behind the buttons on the sheets, run each operation
'   with Excel prepared for speed, put Excel back exactly as it was, and
'   report the outcome to the user.
'
' PUBLIC SURFACE
'   RunSACCR: validate the inputs, calculate and write every output sheet.
'   ValidateInputs: validate only; findings go to the Checks sheet.
'   ClearOutputs: empty every output sheet.
'   RunSACCR_Silent: RunSACCR for automation. Returns a machine-readable
'   result line and raises errors instead of showing them.
'   ResultsStatus: whether the results on the sheets match the current
'   inputs (CURRENT, STALE or NONE).
'   ResultsStatusText: the same as text, for the status formula in Results
'   A3; Excel recalculates it when an input changes, without a macro
'   writing to the workbook, so Undo keeps working.
'   The sheet buttons call the first three by name. These macros are not
'   listed in docs/PUBLIC_API.txt.
'
' DEPENDENCIES
'   CORE_Engine for the calculation and its run counters; CORE_Util and CORE_Config
'   for the output sheets and the error numbers.
'
' STATE OWNERSHIP
'   Each operation captures calculation mode, events and screen updating,
'   changes them for the run, and restores the captured values on success
'   and on failure. Owns gSilent (messages to the Immediate window instead
'   of message boxes, and errors raised to the caller), the re-entry flag
'   mRunning, and the test seam gTestFault.
'
' ERROR POLICY
'   The primary error is copied before cleanup. Each setting is restored by
'   its own helper, so one failed restoration does not stop the others, and
'   a cleanup failure is reported separately from the primary error. With
'   gSilent False the outcome is shown in a message box; with gSilent True
'   it is raised to the caller. Errors in the inputs are not VBA errors:
'   the engine lists them on the Checks sheet.
'
' TEST SEAM
'   gTestFault = "operation" raises ERR_INJECTED_FAULT after the settings
'   have been changed; gTestFault = "cleanup" makes the calculation-mode
'   restoration fail. tests/modules/TEST_MainState.bas uses both. Production
'   code never sets it.
'
' KNOWN DEVIATION
'   This standard module lives in src/workbook. See the known deviations in
'   docs/REPOSITORY_STRUCTURE.md.
'
' COMPATIBILITY
'   Excel VBA; no references beyond the defaults.
'
' UPDATED
'   2026-10-10
'
' AUTHOR
'   Daniele Penza
'==============================================================================

'------------------------------------------------------------------------------
' MODULE SETTINGS
'------------------------------------------------------------------------------
    'Require explicit declarations. The macros must stay visible to the
    'sheet buttons, so this module is not private.
    Option Explicit

'------------------------------------------------------------------------------
' MODULE CONSTANTS
'------------------------------------------------------------------------------
    'Operations run by ExecuteOperation.
        Private Const OP_RUN        As Long = 1    'Calculate and write every output sheet
        Private Const OP_VALIDATE   As Long = 2    'Validate only
        Private Const OP_CLEAR      As Long = 3    'Clear every output sheet

'------------------------------------------------------------------------------
' MODULE STATE
'------------------------------------------------------------------------------
    'gSilent switches messages to the Immediate window and errors to the
    'caller. gTestFault is the test seam described above. mRunning refuses a
    'second operation while one is in progress.
        Public gSilent      As Boolean    'Silent mode for automation and tests
        Public gTestFault   As String     'Test seam: "", "operation" or "cleanup"
        Private mRunning    As Boolean    'True while an operation is in progress


'
'------------------------------------------------------------------------------
'
'                                BUTTON MACROS
'
'------------------------------------------------------------------------------
'

Public Sub RunSACCR()
'
'==============================================================================
'                                   RunSACCR
'------------------------------------------------------------------------------
' PURPOSE
'   Run the full SA-CCR calculation and write every output sheet, then
'   report the trades used, the total EAD and the number of errors and
'   warnings.
'
' USAGE
'   Assigned to the Run SA-CCR button.
'
' UPDATED
'   2026-10-06
'==============================================================================
'

'------------------------------------------------------------------------------
' RUN
'------------------------------------------------------------------------------
        ExecuteOperation OP_RUN

End Sub


Public Function RunSACCR_Silent() As String
'
'==============================================================================
'                               RunSACCR_Silent
'------------------------------------------------------------------------------
' PURPOSE
'   Run RunSACCR unattended: messages go to the Immediate window and any
'   failure is raised to the caller.
'
' RETURNS
'   A result line such as "RESULT=OK; operation=run; errors=0; warnings=0;
'   trades_used=62; trades_read=65; incomplete=0; total_ead=425197517.8;
'   inputs=5A428560; cleanup=PASS". incomplete counts netting sets whose
'   EAD was withheld; inputs is the fingerprint of the inputs the results
'   were calculated from, empty when there are none.
'   RESULT=STOPPED means the inputs could not be loaded; see Checks.
'
' STATE OWNERSHIP
'   Sets gSilent for the run and restores its previous value afterwards,
'   also when the run raises.
'
' ERROR POLICY
'   Re-raises the operation's error, or ERR_CLEANUP_FAILED when Excel
'   settings could not be restored, after gSilent has been restored.
'
' UPDATED
'   2026-10-06
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim previousSilent   As Boolean    'gSilent before this call
    Dim errNumber        As Long       'Error raised by the operation
    Dim errSource        As String     'Its source
    Dim errDescription   As String     'Its description

'------------------------------------------------------------------------------
' RUN SILENTLY
'------------------------------------------------------------------------------
        previousSilent = gSilent
        gSilent = True
        On Error GoTo Failed
        RunSACCR_Silent = ExecuteOperation(OP_RUN)
        gSilent = previousSilent
        Exit Function

'------------------------------------------------------------------------------
' RESTORE THE FLAG AND RE-RAISE
'------------------------------------------------------------------------------
Failed:
        errNumber = Err.Number
        errSource = Err.Source
        errDescription = Err.Description
        On Error GoTo 0
        gSilent = previousSilent
        Err.Raise errNumber, errSource, errDescription

End Function


Public Sub ValidateInputs()
'
'==============================================================================
'                                ValidateInputs
'------------------------------------------------------------------------------
' PURPOSE
'   Validate every input without writing results, show the Checks sheet and
'   report the number of errors and warnings.
'
' USAGE
'   Assigned to the Validate inputs button.
'
' UPDATED
'   2026-10-06
'==============================================================================
'

'------------------------------------------------------------------------------
' VALIDATE
'------------------------------------------------------------------------------
        ExecuteOperation OP_VALIDATE

End Sub


Public Sub ClearOutputs()
'
'==============================================================================
'                                 ClearOutputs
'------------------------------------------------------------------------------
' PURPOSE
'   Empty the five output sheets and stamp the time on the Results sheet.
'
' USAGE
'   Assigned to the Clear outputs button.
'
' UPDATED
'   2026-10-06
'==============================================================================
'

'------------------------------------------------------------------------------
' CLEAR
'------------------------------------------------------------------------------
        ExecuteOperation OP_CLEAR

End Sub


Public Function ResultsStatus() As String
'
'==============================================================================
'                                ResultsStatus
'------------------------------------------------------------------------------
' PURPOSE
'   Tell whether the results on the output sheets were calculated from the
'   inputs as they are now (#36).
'
' RETURNS
'   CURRENT: the inputs match the last completed run.
'   STALE: the inputs changed after it; the results are out of date.
'   NONE: no valid completed-run fingerprint. Protected output cells may
'   remain after failed cleanup; the run error reports that separately.
'
' UPDATED
'   2026-10-07
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim stored   As String    'Fingerprint stored by the last completed run

'------------------------------------------------------------------------------
' COMPARE
'------------------------------------------------------------------------------
        stored = CORE_Engine.LastRunInputs()
        If Len(stored) = 0 Then
            ResultsStatus = "NONE"
        ElseIf stored = CORE_Engine.InputFingerprint() Then
            ResultsStatus = "CURRENT"
        Else
            ResultsStatus = "STALE"
        End If

End Function


Public Function ResultsStatusText( _
    ByVal storedInputs As Variant, _
    ByVal nettingSetRows As Range, _
    ByVal tradeRows As Range, _
    ByVal paramRows As Range) _
    As String
'
'==============================================================================
'                              ResultsStatusText
'------------------------------------------------------------------------------
' PURPOSE
'   Worksheet function for the status cell in Results A3 (#36):
'     =ResultsStatusText(SACCR_RunInputs, NettingSets!$A:$O, Trades!$A:$W,
'                        Params!$A:$H)
'   Its arguments are what Excel watches: a run changes the hidden name
'   SACCR_RunInputs and an edit changes an input range, and either makes
'   Excel recalculate the cell. A formula never clears Excel's undo
'   history, as a macro writing to the sheet would.
'
' INPUTS
'   storedInputs: the value of SACCR_RunInputs.
'   nettingSetRows, tradeRows, paramRows: the input ranges, only so that
'   Excel recalculates; the fingerprint is computed from the sheets as the
'   engine reads them.
'
' RETURNS
'   "Results current (inputs XXXXXXXX)", "OUT OF DATE - inputs changed
'   since the last run; press Run SA-CCR." or "No results: press Run
'   SA-CCR."
'
' ERROR POLICY
'   Returns "Status unavailable" instead of raising: a worksheet function
'   must not interrupt Excel.
'
' UPDATED
'   2026-10-10
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim stored   As String    'Fingerprint of the last completed run

'------------------------------------------------------------------------------
' COMPARE
'------------------------------------------------------------------------------
        On Error GoTo Failed
        stored = SafeStr(storedInputs)
        If Len(stored) = 0 Then
            ResultsStatusText = "No results: press Run SA-CCR."
        ElseIf stored = CORE_Engine.InputFingerprint() Then
            ResultsStatusText = "Results current (inputs " & stored & ")"
        Else
            ResultsStatusText = "OUT OF DATE - inputs changed since the last run; press Run SA-CCR."
        End If
        Exit Function

Failed:
        ResultsStatusText = "Status unavailable"

End Function


'
'------------------------------------------------------------------------------
'
'                              OPERATION RUNNER
'
'------------------------------------------------------------------------------
'

Private Function ExecuteOperation( _
    ByVal operation As Long) _
    As String
'
'==============================================================================
'                               ExecuteOperation
'------------------------------------------------------------------------------
' PURPOSE
'   Run one operation with Excel prepared for speed, restore Excel, then
'   report the outcome.
'
' INPUTS
'   operation: OP_RUN, OP_VALIDATE or OP_CLEAR.
'
' RETURNS
'   The result line built by ReportOutcome; empty when the operation did
'   not complete.
'
' STATE OWNERSHIP
'   Captures calculation mode, events and screen updating before changing
'   them, and restores the captured values on every path. Nothing is
'   changed if the capture itself fails. Clears mRunning before reporting,
'   so a later operation can always run.
'
' ERROR POLICY
'   The primary error is copied into locals before cleanup and reported, or
'   raised in silent mode, after cleanup. A cleanup failure is reported
'   with it, or on its own when the operation succeeded.
'
' UPDATED
'   2026-10-06
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim savedCalculation   As XlCalculation    'Calculation mode before the operation
    Dim savedEvents        As Boolean          'EnableEvents before the operation
    Dim savedScreen        As Boolean          'ScreenUpdating before the operation
    Dim stateCaptured      As Boolean          'True once all three were read
    Dim ok                 As Boolean          'True when the engine completed the run
    Dim errNumber          As Long             'Primary error number; 0 for none
    Dim errSource          As String           'Primary error source
    Dim errDescription     As String           'Primary error description
    Dim cleanupDetails     As String           'Restoration failures; empty when all restored

'------------------------------------------------------------------------------
' REFUSE RE-ENTRY
'------------------------------------------------------------------------------
    'A button pressed while an operation is running must not start a
    'second one on half-written sheets.
        If mRunning Then
            ReportFailure ERR_RUN_ACTIVE, "M_Main.ExecuteOperation", _
                          "An SA-CCR operation is already running.", ""
            Exit Function
        End If
        mRunning = True
        On Error GoTo Failed

'------------------------------------------------------------------------------
' CAPTURE AND PREPARE EXCEL
'------------------------------------------------------------------------------
    'Read all three settings before changing any of them.
        savedCalculation = Application.Calculation
        savedEvents = Application.EnableEvents
        savedScreen = Application.ScreenUpdating
        stateCaptured = True
        Application.ScreenUpdating = False
        Application.EnableEvents = False
        Application.Calculation = xlCalculationManual

'------------------------------------------------------------------------------
' RUN THE OPERATION
'------------------------------------------------------------------------------
        If gTestFault = "operation" Then
            Err.Raise ERR_INJECTED_FAULT, "M_Main.ExecuteOperation", "Injected operation failure."
        End If
        Select Case operation
            Case OP_RUN
                ok = CORE_Engine.Calculate(True)
            Case OP_VALIDATE
                ok = CORE_Engine.Calculate(False)
            Case OP_CLEAR
                ClearAllOutputs
                ok = True
        End Select

'------------------------------------------------------------------------------
' RESTORE EXCEL AND REPORT
'------------------------------------------------------------------------------
    'Reached on success and, through Failed, after an error. The operation
    'handler is switched off first so that a reporting error cannot loop
    'back into it.
CleanUp:
        On Error GoTo 0
        If stateCaptured Then
            cleanupDetails = RestoreSettings(savedCalculation, savedEvents, savedScreen)
        End If
        mRunning = False
        If errNumber <> 0 Then
            ReportFailure errNumber, errSource, errDescription, cleanupDetails
        Else
            ExecuteOperation = ReportOutcome(operation, ok, cleanupDetails)
        End If
        Exit Function

'------------------------------------------------------------------------------
' HANDLE ERROR
'------------------------------------------------------------------------------
Failed:
        errNumber = Err.Number
        errSource = Err.Source
        errDescription = Err.Description
        Resume CleanUp

End Function


Private Sub ClearAllOutputs()
'
'==============================================================================
'                               ClearAllOutputs
'------------------------------------------------------------------------------
' PURPOSE
'   Clear the five output tables and write the time to the run-summary cell.
'
' UPDATED
'   2026-10-06
'==============================================================================
'

'------------------------------------------------------------------------------
' CLEAR
'------------------------------------------------------------------------------
        ClearOutputBlock GetSheet(SH_TRADECALC), FIRST_DATA_ROW, TC_NCOLS
        ClearOutputBlock GetSheet(SH_BUCKETS), FIRST_DATA_ROW, BK_NCOLS
        ClearOutputBlock GetSheet(SH_HEDGING), FIRST_DATA_ROW, HS_NCOLS
        ClearOutputBlock GetSheet(SH_RESULTS), FIRST_DATA_ROW, RS_NCOLS
        ClearOutputBlock GetSheet(SH_CHECKS), FIRST_DATA_ROW, CK_NCOLS
        CORE_Engine.ForgetRunInputs
        GetSheet(SH_RESULTS).Range(RUNINFO_CELL).Value = "Outputs cleared " & Format$(Now, "yyyy-mm-dd hh:mm:ss")

End Sub


'
'------------------------------------------------------------------------------
'
'                               EXCEL SETTINGS
'
'------------------------------------------------------------------------------
'

Private Function RestoreSettings( _
    ByVal savedCalculation As XlCalculation, _
    ByVal savedEvents As Boolean, _
    ByVal savedScreen As Boolean) _
    As String
'
'==============================================================================
'                               RestoreSettings
'------------------------------------------------------------------------------
' PURPOSE
'   Restore the three captured settings, each independently.
'
' RETURNS
'   The failures, separated by "; "; empty when all three were restored.
'
' UPDATED
'   2026-10-06
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim details   As String    'Collected failures
    Dim one       As String    'Failure of one restoration

'------------------------------------------------------------------------------
' RESTORE
'------------------------------------------------------------------------------
    'Every helper runs even when an earlier one failed.
        one = RestoreCalculation(savedCalculation)
        details = AppendDetail(details, one)
        one = RestoreEvents(savedEvents)
        details = AppendDetail(details, one)
        one = RestoreScreenUpdating(savedScreen)
        details = AppendDetail(details, one)
        RestoreSettings = details

End Function


Private Function RestoreCalculation( _
    ByVal savedValue As XlCalculation) _
    As String
'
'==============================================================================
'                              RestoreCalculation
'------------------------------------------------------------------------------
' PURPOSE
'   Set the calculation mode back and check that it took effect.
'
' RETURNS
'   Empty on success; otherwise what failed.
'
' ERROR POLICY
'   Contains its own error and returns it as text.
'
' UPDATED
'   2026-10-06
'==============================================================================
'

'------------------------------------------------------------------------------
' RESTORE
'------------------------------------------------------------------------------
        On Error GoTo Failed
        If gTestFault = "cleanup" Then
            Err.Raise ERR_INJECTED_FAULT, "M_Main.RestoreCalculation", "Injected cleanup failure."
        End If
        Application.Calculation = savedValue
        If Application.Calculation <> savedValue Then
            RestoreCalculation = "Calculation: restored value does not match"
        End If
        Exit Function

'------------------------------------------------------------------------------
' REPORT FAILURE
'------------------------------------------------------------------------------
Failed:
        RestoreCalculation = "Calculation: " & CStr(Err.Number) & " / " & Err.Description

End Function


Private Function RestoreEvents( _
    ByVal savedValue As Boolean) _
    As String
'
'==============================================================================
'                                RestoreEvents
'------------------------------------------------------------------------------
' PURPOSE
'   Set EnableEvents back and check that it took effect.
'
' RETURNS
'   Empty on success; otherwise what failed.
'
' ERROR POLICY
'   Contains its own error and returns it as text.
'
' UPDATED
'   2026-10-06
'==============================================================================
'

'------------------------------------------------------------------------------
' RESTORE
'------------------------------------------------------------------------------
        On Error GoTo Failed
        Application.EnableEvents = savedValue
        If Application.EnableEvents <> savedValue Then
            RestoreEvents = "EnableEvents: restored value does not match"
        End If
        Exit Function

'------------------------------------------------------------------------------
' REPORT FAILURE
'------------------------------------------------------------------------------
Failed:
        RestoreEvents = "EnableEvents: " & CStr(Err.Number) & " / " & Err.Description

End Function


Private Function RestoreScreenUpdating( _
    ByVal savedValue As Boolean) _
    As String
'
'==============================================================================
'                            RestoreScreenUpdating
'------------------------------------------------------------------------------
' PURPOSE
'   Set ScreenUpdating back and check that it took effect.
'
' RETURNS
'   Empty on success; otherwise what failed.
'
' ERROR POLICY
'   Contains its own error and returns it as text.
'
' UPDATED
'   2026-10-06
'==============================================================================
'

'------------------------------------------------------------------------------
' RESTORE
'------------------------------------------------------------------------------
        On Error GoTo Failed
        Application.ScreenUpdating = savedValue
        If Application.ScreenUpdating <> savedValue Then
            RestoreScreenUpdating = "ScreenUpdating: restored value does not match"
        End If
        Exit Function

'------------------------------------------------------------------------------
' REPORT FAILURE
'------------------------------------------------------------------------------
Failed:
        RestoreScreenUpdating = "ScreenUpdating: " & CStr(Err.Number) & " / " & Err.Description

End Function


Private Function AppendDetail( _
    ByVal details As String, _
    ByVal one As String) _
    As String
'
'==============================================================================
'                                 AppendDetail
'------------------------------------------------------------------------------
' PURPOSE
'   Append one failure to a "; "-separated list; an empty one is skipped.
'
' UPDATED
'   2026-10-06
'==============================================================================
'

'------------------------------------------------------------------------------
' APPEND
'------------------------------------------------------------------------------
        If Len(one) = 0 Then
            AppendDetail = details
        ElseIf Len(details) = 0 Then
            AppendDetail = one
        Else
            AppendDetail = details & "; " & one
        End If

End Function


'
'------------------------------------------------------------------------------
'
'                                  REPORTING
'
'------------------------------------------------------------------------------
'

Private Function ReportOutcome( _
    ByVal operation As Long, _
    ByVal ok As Boolean, _
    ByVal cleanupDetails As String) _
    As String
'
'==============================================================================
'                                ReportOutcome
'------------------------------------------------------------------------------
' PURPOSE
'   Tell the user how a completed operation went, and build its result line.
'
' INPUTS
'   operation: the operation that ran.
'   ok: the engine's result; False means the inputs could not be loaded.
'   cleanupDetails: restoration failures; empty when Excel was restored.
'
' RETURNS
'   The result line, for example "RESULT=OK; operation=run; ...;
'   cleanup=PASS".
'
' ERROR POLICY
'   In silent mode a cleanup failure raises ERR_CLEANUP_FAILED, with the
'   result line in its description; otherwise it is added to the message.
'
' UPDATED
'   2026-10-06
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim msg      As String           'Message for the user; empty for none
    Dim style    As VbMsgBoxStyle    'Message-box icon
    Dim result   As String           'Machine-readable result line

'------------------------------------------------------------------------------
' BUILD THE MESSAGE
'------------------------------------------------------------------------------
    'Messages and sheet activation are those of the original macros. A
    'completed run can still have excluded trades, so it points to Checks
    'whenever there are errors or warnings. ShowChecksSheet never raises.
        Select Case operation
            Case OP_RUN
                If ok Then
                    msg = "SA-CCR run completed." & vbCrLf & vbCrLf & _
                          "Trades used: " & CORE_Engine.TradesUsed & " of " & CORE_Engine.TradesRead & vbCrLf & _
                          "Total EAD: " & Format$(CORE_Engine.TotalEAD, "#,##0") & vbCrLf & _
                          "Errors: " & CORE_Engine.ErrorCount & "   Warnings: " & CORE_Engine.WarningCount
                    If CORE_Engine.IncompleteCount > 0 Then
                        msg = msg & vbCrLf & "EAD withheld for " & CORE_Engine.IncompleteCount & _
                              " netting set(s) with rejected trades or invalid inputs."
                    End If
                    If CORE_Engine.ErrorCount > 0 Or CORE_Engine.WarningCount > 0 Then
                        msg = msg & vbCrLf & vbCrLf & "See the Checks sheet for details."
                    End If
                    style = IIf(CORE_Engine.ErrorCount > 0, vbExclamation, vbInformation)
                Else
                    msg = "SA-CCR run stopped - see the Checks sheet."
                    style = vbCritical
                    ShowChecksSheet
                End If
            Case OP_VALIDATE
                ShowChecksSheet
                msg = "Validation finished: " & CORE_Engine.ErrorCount & " error(s), " & _
                      CORE_Engine.WarningCount & " warning(s)."
                style = IIf(CORE_Engine.ErrorCount > 0, vbExclamation, vbInformation)
            Case OP_CLEAR
                msg = ""
                style = vbInformation
        End Select

'------------------------------------------------------------------------------
' BUILD THE RESULT LINE
'------------------------------------------------------------------------------
    'Numbers use Str$, which always writes a decimal point, so the line
    'reads the same in every locale.
        result = "RESULT=" & IIf(ok, "OK", "STOPPED") & "; operation=" & OperationName(operation)
        If operation <> OP_CLEAR Then
            result = result & _
                     "; errors=" & CStr(CORE_Engine.ErrorCount) & _
                     "; warnings=" & CStr(CORE_Engine.WarningCount) & _
                     "; trades_used=" & CStr(CORE_Engine.TradesUsed) & _
                     "; trades_read=" & CStr(CORE_Engine.TradesRead) & _
                     "; incomplete=" & CStr(CORE_Engine.IncompleteCount) & _
                     "; total_ead=" & Trim$(Str$(CORE_Engine.TotalEAD)) & _
                     "; inputs=" & CORE_Engine.RunInputs
        End If
        result = result & "; cleanup=" & IIf(Len(cleanupDetails) = 0, "PASS", "FAIL")

'------------------------------------------------------------------------------
' REPORT A CLEANUP FAILURE
'------------------------------------------------------------------------------
    'Settings left changed matter more than the run summary: raise in
    'silent mode, otherwise add a warning and use the critical icon.
        If Len(cleanupDetails) > 0 Then
            If gSilent Then
                Err.Raise ERR_CLEANUP_FAILED, "M_Main.ExecuteOperation", _
                          "Excel settings could not be restored: " & cleanupDetails & " (" & result & ")"
            End If
            If Len(msg) > 0 Then
                msg = msg & vbCrLf & vbCrLf
            End If
            msg = msg & "Excel settings could not be restored: " & cleanupDetails & _
                  ". Check calculation mode, events and screen updating."
            style = vbCritical
        End If

'------------------------------------------------------------------------------
' SHOW
'------------------------------------------------------------------------------
        If Len(msg) > 0 Then
            Notify msg, style
        End If
        ReportOutcome = result

End Function


Private Sub ShowChecksSheet()
'
'==============================================================================
'                               ShowChecksSheet
'------------------------------------------------------------------------------
' PURPOSE
'   Bring the Checks sheet to the front so the user sees the findings.
'
' STATE OWNERSHIP
'   Activates the Checks sheet; does nothing in silent mode, where no one
'   is looking and the window may be hidden.
'
' ERROR POLICY
'   Best effort: activation fails when the window or the sheet is hidden.
'   That failure is contained, because showing the sheet is a convenience
'   and must not replace the operation's outcome.
'
' UPDATED
'   2026-10-06
'==============================================================================
'

'------------------------------------------------------------------------------
' SHOW
'------------------------------------------------------------------------------
        If gSilent Then
            Exit Sub
        End If
        On Error GoTo NotShown
        ThisWorkbook.Worksheets(SH_CHECKS).Activate
        Exit Sub

'------------------------------------------------------------------------------
' IGNORE A HIDDEN SHEET OR WINDOW
'------------------------------------------------------------------------------
NotShown:
        Err.Clear

End Sub


Private Sub ReportFailure( _
    ByVal errNumber As Long, _
    ByVal errSource As String, _
    ByVal errDescription As String, _
    ByVal cleanupDetails As String)
'
'==============================================================================
'                                ReportFailure
'------------------------------------------------------------------------------
' PURPOSE
'   Report an operation that did not complete: raise it in silent mode,
'   otherwise show it.
'
' INPUTS
'   errNumber, errSource, errDescription: the primary error, as copied
'      before cleanup.
'   cleanupDetails: restoration failures; empty when Excel was restored.
'
' ERROR POLICY
'   In silent mode raises the primary error unchanged, with any cleanup
'   failure appended to its description.
'
' UPDATED
'   2026-10-06
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim msg   As String    'Message for the user

'------------------------------------------------------------------------------
' RAISE IN SILENT MODE
'------------------------------------------------------------------------------
        If Len(cleanupDetails) > 0 Then
            errDescription = errDescription & " Excel settings could not be restored: " & cleanupDetails & "."
        End If
        If gSilent Then
            Err.Raise errNumber, errSource, errDescription
        End If

'------------------------------------------------------------------------------
' SHOW OTHERWISE
'------------------------------------------------------------------------------
    'The project's own errors carry a readable description; anything else
    'is shown with its number.
        If errNumber = ERR_RUN_ACTIVE Or errNumber = ERR_CHECKS_WRITE Then
            msg = errDescription
        Else
            msg = "Unexpected error " & errNumber & ": " & errDescription
        End If
        MsgBox msg, vbCritical, "SA-CCR"

End Sub


Private Function OperationName( _
    ByVal operation As Long) _
    As String
'
'==============================================================================
'                                OperationName
'------------------------------------------------------------------------------
' PURPOSE
'   Name an operation for the result line: run, validate or clear.
'
' UPDATED
'   2026-10-06
'==============================================================================
'

'------------------------------------------------------------------------------
' NAME
'------------------------------------------------------------------------------
        Select Case operation
            Case OP_RUN
                OperationName = "run"
            Case OP_VALIDATE
                OperationName = "validate"
            Case Else
                OperationName = "clear"
        End Select

End Function


Private Sub Notify( _
    ByVal msg As String, _
    ByVal style As VbMsgBoxStyle)
'
'==============================================================================
'                                    Notify
'------------------------------------------------------------------------------
' PURPOSE
'   Show a message to the user, or print it in silent mode.
'
' INPUTS
'   msg: the text to show.
'   style: message-box icon, for example vbInformation or vbCritical.
'
' UPDATED
'   2026-10-06
'==============================================================================
'

'------------------------------------------------------------------------------
' SHOW OR PRINT
'------------------------------------------------------------------------------
        If gSilent Then
            Debug.Print msg
        Else
            MsgBox msg, style, "SA-CCR"
        End If

End Sub
