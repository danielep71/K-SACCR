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
'   C5; Excel recalculates it when an input changes, without a macro
'   writing to the workbook, so Undo keeps working.
'   ReconciliationStatusText: the same for the reconciliation of the bank
'   figures, for the status formula in Reconciliation C5 (#45).
'   The sheet buttons call the first three by name. These macros are not
'   listed in docs/PUBLIC_API.txt.
'
' DEPENDENCIES
'   CORE_Engine for the calculation and its run counters; CORE_Reconcile
'   for the comparison of the bank figures with Results (#45), whose
'   sheets this module reads and writes; CORE_Util and CORE_Config for the
'   output sheets and the error numbers.
'
' STATE OWNERSHIP
'   Each operation captures calculation mode, events and screen updating,
'   changes them for the run, and restores the captured values on success
'   and on failure. Owns gSilent (messages to the Immediate window instead
'   of message boxes, and errors raised to the caller), the re-entry flag
'   mRunning, and the test seams gTestFault and gQuietLog (silent-mode
'   messages dropped instead of printed).
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
'   restoration fail; gTestFault = "recalc" takes the path of a
'   recalculation that did not finish (#195); gTestFault = "recalc-error"
'   raises ERR_INJECTED_FAULT where the recalculation starts, as an error
'   thrown by it would (#204). tests/modules/TEST_MainState.bas uses all
'   four. Production code never sets it.
'
' KNOWN DEVIATION
'   This standard module lives in src/workbook. See the known deviations in
'   docs/REPOSITORY_STRUCTURE.md.
'
' COMPATIBILITY
'   Excel VBA; no references beyond the defaults.
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

    'Title of every message box: the project name.
        Private Const MSG_TITLE     As String = "K-SACCR"

'------------------------------------------------------------------------------
' MODULE STATE
'------------------------------------------------------------------------------
    'gSilent switches messages to the Immediate window and errors to the
    'caller. gTestFault is the test seam described above. gQuietLog drops
    'the messages of silent mode instead of printing them, for
    'TEST_Suite.Test_RunAll. mRunning refuses a second operation while one
    'is in progress.
        Public gSilent      As Boolean    'Silent mode for automation and tests
        Public gTestFault   As String     'Test seam: "", "operation", "cleanup", "recalc" or "recalc-error"
        Public gQuietLog    As Boolean    'Test seam: silent-mode messages not printed
        Private mRunning    As Boolean    'True while an operation is in progress
    'The reconciliation of the last operation (#45), for its message and
    'result line: NONE, DONE, FAILED or UNAVAILABLE, and the message line.
        Private mReconState As String     'Reconciliation outcome
        Private mReconLine  As String     'Line for the run message; "" for none


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
'   trades_used=62; trades_read=65; incomplete=0; total_ead_crr=4251975.8;
'   total_ead_bcbs=8193300.1; inputs=5A428560; cleanup=PASS". incomplete
'   counts netting sets whose EAD was withheld; total_ead_crr and
'   total_ead_bcbs are the totals of each regime's VALID netting sets, 0
'   when it has none (#198); inputs is the fingerprint of the inputs the
'   results were calculated from, empty when there are none.
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
'   2026-10-13
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
' ERROR POLICY
'   Raises ERR_PARAMS_LAYOUT from CORE_Engine.InputFingerprint when a
'   Params workbook name is missing or broken (#152).
'
' UPDATED
'   2026-10-12
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
'   Worksheet function for the status cell in Results C5 (#36):
'     =ResultsStatusText(SACCR_RunInputs, NettingSets!$C:$Q, Trades!$C:$AA,
'                        Params!$1:$1048576)
'   Its arguments are what Excel watches: a run changes the hidden name
'   SACCR_RunInputs and an edit changes an input range, and either makes
'   Excel recalculate the cell. Params is watched whole, because its tables
'   are located through workbook names and can move (#152). A formula
'   never clears Excel's undo history, as a macro writing to the sheet
'   would.
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
'   Returns "Status unavailable: " and the reason instead of raising, for
'   example when a Params workbook name is missing (#152): a worksheet
'   function must not interrupt Excel.
'
' UPDATED
'   2026-10-12
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
        ResultsStatusText = "Status unavailable: " & Err.Description

End Function


'
'------------------------------------------------------------------------------
'
'                                RECONCILIATION
'
'------------------------------------------------------------------------------
'

Public Function ReconciliationStatusText( _
    ByVal storedRecon As Variant, _
    ByVal storedRun As Variant, _
    ByVal bankRows As Range, _
    ByVal nettingSetRows As Range, _
    ByVal tradeRows As Range, _
    ByVal paramRows As Range) _
    As String
'
'==============================================================================
'                           ReconciliationStatusText
'------------------------------------------------------------------------------
' PURPOSE
'   Worksheet function for the status cell in Reconciliation C5 (#45):
'     =ReconciliationStatusText(SACCR_ReconInputs, SACCR_RunInputs,
'        BankNettingSets!$1:$1048576, NettingSets!$C:$Q, Trades!$C:$AA,
'        Params!$1:$1048576)
'   As ResultsStatusText: a run changes the hidden names and an edit
'   changes a watched range, and either makes Excel recalculate the cell.
'   BankNettingSets is watched whole, because its settings and table are
'   located through workbook names.
'
' INPUTS
'   storedRecon: the value of SACCR_ReconInputs; storedRun: the value of
'   SACCR_RunInputs.
'   bankRows, nettingSetRows, tradeRows, paramRows: the watched ranges,
'   only so that Excel recalculates.
'
' RETURNS
'   "Reconciliation current (...)", "OUT OF DATE - ..." or "No
'   reconciliation: ...".
'
' ERROR POLICY
'   Returns "Status unavailable: " and the reason instead of raising.
'
' UPDATED
'   2026-10-15
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim stored   As String    'Fingerprint of the last reconciliation

'------------------------------------------------------------------------------
' COMPARE
'------------------------------------------------------------------------------
    'Without current results there is nothing the reconciliation can still
    'describe, whatever its own fingerprint says.
        On Error GoTo Failed
        stored = SafeStr(storedRecon)
        If Len(stored) = 0 Or Len(SafeStr(storedRun)) = 0 Then
            ReconciliationStatusText = "No reconciliation: enter bank figures on BankNettingSets and press Run SA-CCR."
        ElseIf stored = ReconFingerprint() Then
            ReconciliationStatusText = "Reconciliation current (bank figures and inputs " & stored & ")"
        Else
            ReconciliationStatusText = "OUT OF DATE - bank figures or inputs changed since the last run; " & _
                                       "press Run SA-CCR."
        End If
        Exit Function

Failed:
        ReconciliationStatusText = "Status unavailable: " & Err.Description

End Function


Private Function ReconcileAfterRun() As String
'
'==============================================================================
'                              ReconcileAfterRun
'------------------------------------------------------------------------------
' PURPOSE
'   After a completed run, compare the figures on BankNettingSets with
'   Results and write the Reconciliation sheet, its summary and its
'   fingerprint (#45). Sets mReconLine for the run message.
'
' RETURNS
'   DONE: rows written. NONE: no bank figures entered. UNAVAILABLE: the
'   workbook has no BankNettingSets or Reconciliation sheet, for example
'   one built from an older template. FAILED: the settings, the headers or
'   the reconciliation itself failed; the summary cell says why.
'
' ERROR POLICY
'   Never raises: the SA-CCR results stand whatever happens here. An error
'   leaves the Reconciliation sheet cleared, its fingerprint empty and the
'   reason in its summary cell.
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
    Dim headerRow   As Long         'Header row of the bank table
    Dim problem     As String       'Why the reconciliation cannot run; "" when it can
    Dim scaleValue  As Variant      'BankNS_Scale
    Dim absValue    As Variant      'BankNS_AbsTol
    Dim relValue    As Variant      'BankNS_RelTol
    Dim lastR       As Long         'Last row read
    Dim bank        As Variant      'Bank rows; Empty when none
    Dim calc        As Variant      'Results rows; Empty when none
    Dim counts()    As Long         'Counts from CompareNettingSets
    Dim found       As Variant      'Rows from CompareNettingSets; Empty when none
    Dim outRows()   As Variant      'The same as an array for WriteBlock
    Dim nRows       As Long         'Rows to write
    Dim summary     As String       'Summary line without its time stamp

'------------------------------------------------------------------------------
' FIND THE SHEETS AND SETTINGS
'------------------------------------------------------------------------------
        On Error GoTo Failed
        If Not ReconciliationSheetsPresent() Then
            ReconcileAfterRun = "UNAVAILABLE"
            Exit Function
        End If
        Set wsBank = GetSheet(SH_BANK_NS)
        Set wsRecon = GetSheet(SH_RECON)
        headerRow = BankHeaderRow()
        If headerRow = 0 Then
            problem = "the name " & NM_BN_HEADER & " is missing or not in column C of " & SH_BANK_NS
        Else
            problem = CORE_Reconcile.HeadersProblem( _
                      TableRange(wsBank, headerRow, 1, headerRow, BN_NCOLS).Value2, BN_HEADERS)
        End If
        If Len(problem) = 0 Then
            problem = NamedValue(NM_BN_SCALE, scaleValue)
        End If
        If Len(problem) = 0 Then
            problem = NamedValue(NM_BN_ABSTOL, absValue)
        End If
        If Len(problem) = 0 Then
            problem = NamedValue(NM_BN_RELTOL, relValue)
        End If
        If Len(problem) = 0 Then
            problem = CORE_Reconcile.SettingsProblem(scaleValue, absValue, relValue)
        End If
        If Len(problem) > 0 Then
            ClearReconciliation "Reconciliation not run " & Format$(Now, "yyyy-mm-dd hh:mm:ss") & ": " & problem & "."
            mReconLine = "Reconciliation not run: " & problem & "."
            ReconcileAfterRun = "FAILED"
            Exit Function
        End If

'------------------------------------------------------------------------------
' READ THE BANK FIGURES AND THE RESULTS
'------------------------------------------------------------------------------
        lastR = UsedLastRow(wsBank)
        If lastR > headerRow Then
            bank = TableRange(wsBank, headerRow + 1, 1, lastR, BN_NCOLS).Value2
        End If
        If Not CORE_Reconcile.HasBankRows(bank) Then
            ClearReconciliation "No bank figures on " & SH_BANK_NS & ": nothing to reconcile."
            ReconcileAfterRun = "NONE"
            Exit Function
        End If
        Set wsResults = GetSheet(SH_RESULTS)
        lastR = LastDataRow(wsResults, FIRST_DATA_ROW, RS_ID)
        If lastR >= FIRST_DATA_ROW Then
            calc = TableRange(wsResults, FIRST_DATA_ROW, 1, lastR, RS_NCOLS).Value2
        End If

'------------------------------------------------------------------------------
' COMPARE AND WRITE
'------------------------------------------------------------------------------
        found = CORE_Reconcile.CompareNettingSets(bank, headerRow + 1, calc, CDbl(scaleValue), CDbl(absValue), _
                                              CDbl(relValue), counts)
        ClearOutputBlock wsRecon, FIRST_DATA_ROW, RC_NCOLS
        If IsArray(found) Then
            outRows = found
            nRows = UBound(outRows, 1)
            WriteBlock wsRecon, FIRST_DATA_ROW, outRows, nRows, RC_NCOLS
            FormatColumns wsRecon, FIRST_DATA_ROW, nRows, "@|0|@|@|@|@|#,##0.00|#,##0.00|#,##0.00|0.0000%|@|@"
        End If
        summary = CORE_Reconcile.SummaryText(counts, CDbl(scaleValue), CDbl(absValue), CDbl(relValue))
        wsRecon.Range(RECON_INFO_CELL).Value = "Last reconciliation " & Format$(Now, "yyyy-mm-dd hh:mm:ss") & _
                                               " | " & summary
        RememberReconInputs ReconFingerprint()
        mReconLine = "Reconciliation: " & counts(RCN_COMPARED) & " figure(s) compared, " & _
                     counts(RCN_MATCH) & " MATCH, " & counts(RCN_DIFF) & " DIFFERENCE"
        If counts(RCN_MISSING) + counts(RCN_NOTINBANK) + counts(RCN_DUPLICATE) + counts(RCN_REGIME) + _
           counts(RCN_WITHHELD) + counts(RCN_INVALID) > 0 Then
            mReconLine = mReconLine & "; netting sets not compared: see the Reconciliation sheet"
        End If
        mReconLine = mReconLine & "."
        ReconcileAfterRun = "DONE"
        Exit Function

'------------------------------------------------------------------------------
' HANDLE ERROR
'------------------------------------------------------------------------------
Failed:
        ClearReconciliation "Reconciliation failed " & Format$(Now, "yyyy-mm-dd hh:mm:ss") & " (error " & _
                            Err.Number & ": " & Err.Description & ")."
        mReconLine = "Reconciliation failed: " & Err.Description & "."
        ReconcileAfterRun = "FAILED"

End Function


Private Sub ClearReconciliation( _
    ByVal infoText As String)
'
'==============================================================================
'                             ClearReconciliation
'------------------------------------------------------------------------------
' PURPOSE
'   Withdraw the reconciliation: empty its fingerprint and its rows and
'   write infoText to its summary cell. Every operation starts with this,
'   so that no reconciliation outlives the results it compared.
'
' ERROR POLICY
'   Best effort, and never raises: the fingerprint is emptied first, so a
'   sheet that cannot be cleared is still reported as not current.
'
' UPDATED
'   2026-10-15
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ws   As Worksheet    'Reconciliation

'------------------------------------------------------------------------------
' CLEAR
'------------------------------------------------------------------------------
        On Error GoTo Failed
        ForgetReconInputs
        If Not ReconciliationSheetsPresent() Then
            Exit Sub
        End If
        Set ws = GetSheet(SH_RECON)
        ClearOutputBlock ws, FIRST_DATA_ROW, RC_NCOLS
        ws.Range(RECON_INFO_CELL).Value = infoText

Failed:

End Sub


Private Function ReconFingerprint() As String
'
'==============================================================================
'                               ReconFingerprint
'------------------------------------------------------------------------------
' PURPOSE
'   Fingerprint what a reconciliation depends on: the run inputs
'   (CORE_Engine.InputFingerprint), the bank table from its header row,
'   without the comment column, and the scale and tolerance cells.
'
' ERROR POLICY
'   Raises when the bank table or a setting cannot be located, or as
'   CORE_Engine.InputFingerprint.
'
' UPDATED
'   2026-10-15
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim parts(1 To 5)   As String    'Text of each part
    Dim headerRow       As Long      'Header row of the bank table
    Dim i               As Long      'Setting
    Dim c               As Range     'A setting cell

'------------------------------------------------------------------------------
' FINGERPRINT
'------------------------------------------------------------------------------
        headerRow = BankHeaderRow()
        If headerRow = 0 Then
            Err.Raise ERR_PARAMS_LAYOUT, "M_Main.ReconFingerprint", _
                      "The name " & NM_BN_HEADER & " is missing or not in column C of " & SH_BANK_NS & "."
        End If
        parts(1) = CORE_Engine.InputFingerprint()
        parts(2) = CORE_Engine.BlockText(GetSheet(SH_BANK_NS), headerRow, BN_COMMENT - 1)
        For i = 0 To 2
            Set c = NamedCell(Choose(i + 1, NM_BN_SCALE, NM_BN_ABSTOL, NM_BN_RELTOL))
            If c Is Nothing Then
                Err.Raise ERR_PARAMS_LAYOUT, "M_Main.ReconFingerprint", _
                          "The name " & Choose(i + 1, NM_BN_SCALE, NM_BN_ABSTOL, NM_BN_RELTOL) & " is missing."
            End If
            parts(3 + i) = CORE_Engine.BlockText(c.Worksheet, c.Row, 1, c.Column - COL_OFFSET, c.Row)
        Next i
        ReconFingerprint = TextHash(Join(parts, Chr$(29)))

End Function


Private Sub RememberReconInputs( _
    ByVal fingerprint As String)
'
'==============================================================================
'                             RememberReconInputs
'------------------------------------------------------------------------------
' PURPOSE
'   Store the fingerprint of a completed reconciliation in a hidden
'   workbook name.
'
' ERROR POLICY
'   Best effort, as CORE_Engine.RememberRunInputs: a name that cannot be
'   changed leaves the reconciliation reported as not current.
'
' UPDATED
'   2026-10-15
'==============================================================================
'
        On Error GoTo Failed
        ThisWorkbook.Names.Add Name:=RECON_INPUTS_NAME, RefersTo:="=""" & fingerprint & """", Visible:=False
        Exit Sub

Failed:
        ForgetReconInputs

End Sub


Private Sub ForgetReconInputs()
'
'==============================================================================
'                              ForgetReconInputs
'------------------------------------------------------------------------------
' PURPOSE
'   Empty the stored reconciliation fingerprint. The name is kept, empty,
'   because the status formula on Reconciliation refers to it.
'
' ERROR POLICY
'   Contains any error.
'
' UPDATED
'   2026-10-15
'==============================================================================
'
        On Error GoTo Failed
        ThisWorkbook.Names.Add Name:=RECON_INPUTS_NAME, RefersTo:="=""""", Visible:=False

Failed:

End Sub


Private Function ReconciliationSheetsPresent() As Boolean
'
'==============================================================================
'                         ReconciliationSheetsPresent
'------------------------------------------------------------------------------
' PURPOSE
'   Tell whether the workbook has both the BankNettingSets and the
'   Reconciliation sheet. A workbook built from a template before #45 has
'   neither, and still runs.
'
' ERROR POLICY
'   Contains the expected error of a missing sheet.
'
' UPDATED
'   2026-10-15
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim wsBank    As Worksheet    'BankNettingSets; Nothing when absent
    Dim wsRecon   As Worksheet    'Reconciliation; Nothing when absent

'------------------------------------------------------------------------------
' LOOK UP
'------------------------------------------------------------------------------
        On Error Resume Next
        Set wsBank = ThisWorkbook.Worksheets(SH_BANK_NS)
        Set wsRecon = ThisWorkbook.Worksheets(SH_RECON)
        Err.Clear
        On Error GoTo 0
        ReconciliationSheetsPresent = Not wsBank Is Nothing And Not wsRecon Is Nothing

End Function


Private Function BankHeaderRow() As Long
'
'==============================================================================
'                                BankHeaderRow
'------------------------------------------------------------------------------
' PURPOSE
'   Return the header row of the bank table, from the workbook name
'   NM_BN_HEADER, which must be in column C of BankNettingSets.
'
' RETURNS
'   The row; 0 when the name is missing, broken, on another sheet or not
'   in column C.
'
' UPDATED
'   2026-10-15
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim c   As Range    'The named cell

'------------------------------------------------------------------------------
' RESOLVE
'------------------------------------------------------------------------------
        Set c = NamedCell(NM_BN_HEADER)
        If Not c Is Nothing Then
            If StrComp(c.Worksheet.Name, SH_BANK_NS, vbTextCompare) = 0 And c.Column = COL_OFFSET + 1 Then
                BankHeaderRow = c.Row
            End If
        End If

End Function


Private Function NamedValue( _
    ByVal nameText As String, _
    ByRef valueRead As Variant) _
    As String
'
'==============================================================================
'                                  NamedValue
'------------------------------------------------------------------------------
' PURPOSE
'   Read the value of a workbook name's cell on BankNettingSets.
'
' INPUTS
'   nameText: the workbook name.
'   valueRead: receives the cell's Value2.
'
' RETURNS
'   "" when it was read; otherwise why not.
'
' UPDATED
'   2026-10-15
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim c   As Range    'The named cell

'------------------------------------------------------------------------------
' READ
'------------------------------------------------------------------------------
        Set c = NamedCell(nameText)
        If c Is Nothing Then
            NamedValue = "the name " & nameText & " is missing"
        ElseIf StrComp(c.Worksheet.Name, SH_BANK_NS, vbTextCompare) <> 0 Then
            NamedValue = "the name " & nameText & " is not on " & SH_BANK_NS
        Else
            valueRead = c.Value2
        End If

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
'   2026-10-15
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
' BRING FORMULA INPUTS UP TO DATE
'------------------------------------------------------------------------------
    'In manual calculation a formula whose precedents changed since the last
    'recalculation still holds its old result, and the engine would
    'calculate from it and fingerprint it as current. Calculate brings every
    'dirty cell up to date first; it refreshes no external data connection
    '(#195). Clear outputs reads no input. The stored fingerprint is
    'withdrawn before the recalculation starts: if it does not finish, for
    'example after Esc, or raises an error, a dirty input could still show
    'the value the last run used, so the earlier results must not be
    'reported as current (#204). A run that completes records its own
    'fingerprint again; validation never keeps one.
        If operation <> OP_CLEAR Then
            CORE_Engine.ForgetRunInputs
            If gTestFault = "recalc-error" Then
                Err.Raise ERR_INJECTED_FAULT, "M_Main.ExecuteOperation", "Injected recalculation failure."
            End If
            Application.Calculate
            If Application.CalculationState <> xlDone Or gTestFault = "recalc" Then
                Err.Raise ERR_RECALC_PENDING, "M_Main.ExecuteOperation", _
                          "Excel did not finish recalculating the inputs, so the operation was not started " & _
                          "and the previous results are no longer reported as current."
            End If
        End If

'------------------------------------------------------------------------------
' RUN THE OPERATION
'------------------------------------------------------------------------------
        If gTestFault = "operation" Then
            Err.Raise ERR_INJECTED_FAULT, "M_Main.ExecuteOperation", "Injected operation failure."
        End If
    'Every operation first withdraws the reconciliation (#45); only a run
    'that completes makes a new one.
        mReconState = "NONE"
        mReconLine = ""
        Select Case operation
            Case OP_RUN
                ClearReconciliation "Reconciliation pending: the SA-CCR run has not completed."
            Case OP_VALIDATE
                ClearReconciliation "Outputs cleared by validation: press Run SA-CCR to reconcile."
            Case OP_CLEAR
                ClearReconciliation "Outputs cleared " & Format$(Now, "yyyy-mm-dd hh:mm:ss")
        End Select
        Select Case operation
            Case OP_RUN
                ok = CORE_Engine.Calculate(True)
                If ok Then
                    mReconState = ReconcileAfterRun()
                Else
                    ClearReconciliation "Reconciliation not run: the SA-CCR run stopped; see the Checks sheet."
                End If
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
    'back into it. The status cell is recalculated first: the operation
    'stored or withdrew the run fingerprint, and in manual calculation the
    'cell would keep the value computed while the fingerprint was empty.
CleanUp:
        On Error GoTo 0
        If stateCaptured Then
            RefreshResultsStatus
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


Private Sub RefreshResultsStatus()
'
'==============================================================================
'                             RefreshResultsStatus
'------------------------------------------------------------------------------
' PURPOSE
'   Recalculate the status formulas on Results and Reconciliation after an
'   operation has stored or withdrawn the run fingerprint. In manual calculation Excel
'   would otherwise keep the value computed during the recalculation
'   before the run, when the fingerprint had already been withdrawn
'   (#204), and show "No results" after a completed run.
'
' ERROR POLICY
'   Best effort: a status that cannot be recalculated here is refreshed by
'   the next recalculation, and the outcome of the operation is unchanged.
'
' UPDATED
'   2026-10-15
'==============================================================================
'
        On Error GoTo Failed
        GetSheet(SH_RESULTS).Range(RESULTS_STATUS_CELL).Calculate
        If ReconciliationSheetsPresent() Then
            GetSheet(SH_RECON).Range(RECON_STATUS_CELL).Calculate
        End If

Failed:

End Sub


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
'   reconciliation=DONE; cleanup=PASS".
'
' ERROR POLICY
'   In silent mode a cleanup failure raises ERR_CLEANUP_FAILED, with the
'   result line in its description; otherwise it is added to the message.
'
' UPDATED
'   2026-10-15
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
                          "Economic trades used: " & CORE_Engine.EconomicTradesUsed & " of " & _
                          CORE_Engine.EconomicTradesRead & vbCrLf & _
                          "Legs (rows) used: " & CORE_Engine.TradesUsed & " of " & CORE_Engine.TradesRead & vbCrLf & _
                          TotalLines() & _
                          "Errors: " & CORE_Engine.ErrorCount & "   Warnings: " & CORE_Engine.WarningCount & _
                          vbCrLf & RunClassLine()
                    If CORE_Engine.IncompleteCount > 0 Then
                        msg = msg & vbCrLf & "EAD withheld for " & CORE_Engine.IncompleteCount & _
                              " netting set(s) with rejected trades or invalid inputs."
                    End If
                    If Len(mReconLine) > 0 Then
                        msg = msg & vbCrLf & vbCrLf & mReconLine
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
                      CORE_Engine.WarningCount & " warning(s)." & vbCrLf & RunClassLine()
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
                     "; economic_trades_used=" & CStr(CORE_Engine.EconomicTradesUsed) & _
                     "; economic_trades_read=" & CStr(CORE_Engine.EconomicTradesRead) & _
                     "; incomplete=" & CStr(CORE_Engine.IncompleteCount) & _
                     "; total_ead_crr=" & Trim$(Str$(CORE_Engine.TotalEAD(RG_CRR))) & _
                     "; total_ead_bcbs=" & Trim$(Str$(CORE_Engine.TotalEAD(RG_BCBS))) & _
                     "; inputs=" & CORE_Engine.RunInputs & _
                     "; run_class=" & CORE_Engine.RunClass & _
                     "; reconciliation=" & mReconState
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
        MsgBox msg, vbCritical, MSG_TITLE

End Sub


Private Function TotalLines() As String
'
'==============================================================================
'                                  TotalLines
'------------------------------------------------------------------------------
' PURPOSE
'   The message-box lines with the total EAD of each regime that has
'   netting sets, for example "Total EAD CRR: 12,345 (3 of 4 netting
'   set(s) VALID; 1 withheld)", each ending with a line break. EADs under
'   CRR and BCBS are never added together (#198).
'
' UPDATED
'   2026-10-13
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim regimes   As Variant    'The regimes, CRR first
    Dim i         As Long       'Index into regimes

'------------------------------------------------------------------------------
' BUILD
'------------------------------------------------------------------------------
        regimes = Array(RG_CRR, RG_BCBS)
        For i = LBound(regimes) To UBound(regimes)
            If CORE_Engine.RegimeSetCount(CStr(regimes(i))) > 0 Then
                TotalLines = TotalLines & "Total EAD " & regimes(i) & ": " & _
                             Format$(CORE_Engine.TotalEAD(CStr(regimes(i))), "#,##0") & _
                             " (" & CORE_Engine.TotalStatusText(CStr(regimes(i))) & ")" & vbCrLf
            End If
        Next i

End Function


Private Function RunClassLine() As String
'
'==============================================================================
'                                 RunClassLine
'------------------------------------------------------------------------------
' PURPOSE
'   One line for the message box saying whether the run used the
'   regulatory parameter set or is a what-if run (#41); "" when the run was
'   not classified.
'
' UPDATED
'   2026-10-11
'==============================================================================
'
        Select Case CORE_Engine.RunClass
            Case RUN_REGULATORY
                RunClassLine = "Parameters: regulatory set " & REF_SET_ID
            Case RUN_WHATIF
                RunClassLine = "WHAT-IF RUN: " & CORE_Engine.ParameterDeviations & _
                               " parameter(s) differ from the regulatory set (listed on Checks)."
        End Select

End Function


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
'   Show a message to the user, or print it in silent mode unless
'   gQuietLog is set.
'
' INPUTS
'   msg: the text to show.
'   style: message-box icon, for example vbInformation or vbCritical.
'
' UPDATED
'   2026-10-12
'==============================================================================
'

'------------------------------------------------------------------------------
' SHOW OR PRINT
'------------------------------------------------------------------------------
        If gSilent Then
            If Not gQuietLog Then
                Debug.Print msg
            End If
        Else
            MsgBox msg, style, MSG_TITLE
        End If

End Sub
