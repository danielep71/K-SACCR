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
'   ReconciliationStatusText and TradeReconciliationStatusText: the same
'   for the reconciliations of the bank figures, for the status formulas
'   in Reconciliation C5 and TradeReconciliation C5 (#45).
'   ExportEvidence: write the evidence report of the last run to a CSV
'   file (#45); ExportEvidence_Silent: the same to a given file, for
'   automation and tests.
'   The sheet buttons call the first three by name. These macros are not
'   listed in docs/PUBLIC_API.txt.
'
' DEPENDENCIES
'   CORE_Engine for the calculation and its run counters; CORE_Reconcile
'   for the comparison of the bank figures with Results and TradeCalc
'   (#45), whose sheets this module reads and writes; CORE_Export for the
'   evidence report; CORE_Util and CORE_Config for the
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

    'Reconciliations of bank figures (#45).
        Private Const RECON_NS      As Long = 1    'BankNettingSets against Results
        Private Const RECON_TR      As Long = 2    'BankTrades against TradeCalc

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
    'The reconciliations of the last operation (#45), for its message and
    'result line: NONE, DONE, FAILED or UNAVAILABLE, and the message lines.
        Private mReconState      As String    'Netting-set reconciliation outcome
        Private mTradeReconState As String    'Trade reconciliation outcome
        Private mReconLine       As String    'Lines for the run message; "" for none


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
'   As ReconStatus.
'
' UPDATED
'   2026-10-15
'==============================================================================
'
        ReconciliationStatusText = ReconStatus(RECON_NS, storedRecon, storedRun)

End Function


Public Function TradeReconciliationStatusText( _
    ByVal storedRecon As Variant, _
    ByVal storedRun As Variant, _
    ByVal bankRows As Range, _
    ByVal nettingSetRows As Range, _
    ByVal tradeRows As Range, _
    ByVal paramRows As Range) _
    As String
'
'==============================================================================
'                        TradeReconciliationStatusText
'------------------------------------------------------------------------------
' PURPOSE
'   Worksheet function for the status cell in TradeReconciliation C5
'   (#45), as ReconciliationStatusText:
'     =TradeReconciliationStatusText(SACCR_TradeReconInputs,
'        SACCR_RunInputs, BankTrades!$1:$1048576, NettingSets!$C:$Q,
'        Trades!$C:$AA, Params!$1:$1048576)
'
' INPUTS
'   storedRecon: the value of SACCR_TradeReconInputs; the others as
'   ReconciliationStatusText, with BankTrades as bankRows.
'
' RETURNS
'   As ReconStatus.
'
' UPDATED
'   2026-10-15
'==============================================================================
'
        TradeReconciliationStatusText = ReconStatus(RECON_TR, storedRecon, storedRun)

End Function


Private Function ReconStatus( _
    ByVal kind As Long, _
    ByVal storedRecon As Variant, _
    ByVal storedRun As Variant) _
    As String
'
'==============================================================================
'                                 ReconStatus
'------------------------------------------------------------------------------
' PURPOSE
'   Compare the stored fingerprint of a reconciliation with the bank
'   figures and inputs as they are now.
'
' INPUTS
'   kind: RECON_NS or RECON_TR; storedRecon, storedRun: as the worksheet
'   functions.
'
' RETURNS
'   "Reconciliation current (...)", "OUT OF DATE - ..." or "No
'   reconciliation: ...".
'
' ERROR POLICY
'   Returns "Status unavailable: " and the reason instead of raising: a
'   worksheet function must not interrupt Excel.
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
            ReconStatus = "No reconciliation: enter bank figures on " & BankSheetOf(kind) & _
                          " and press Run SA-CCR."
        ElseIf stored = ReconFingerprint(kind) Then
            ReconStatus = "Reconciliation current (bank figures and inputs " & stored & ")"
        Else
            ReconStatus = "OUT OF DATE - bank figures or inputs changed since the last run; press Run SA-CCR."
        End If
        Exit Function

Failed:
        ReconStatus = "Status unavailable: " & Err.Description

End Function


Private Function ReconcileAfterRun( _
    ByVal kind As Long) _
    As String
'
'==============================================================================
'                              ReconcileAfterRun
'------------------------------------------------------------------------------
' PURPOSE
'   After a completed run, compare the figures on a bank sheet with the
'   calculator's sheet and write the reconciliation sheet, its summary and
'   its fingerprint (#45): BankNettingSets against Results into
'   Reconciliation, or BankTrades against TradeCalc into
'   TradeReconciliation. Adds a line to mReconLine for the run message.
'
' INPUTS
'   kind: RECON_NS or RECON_TR.
'
' RETURNS
'   DONE: rows written. NONE: no bank figures entered. UNAVAILABLE: the
'   workbook does not have the two sheets, for example one built from an
'   older template. FAILED: the settings, the headers or the
'   reconciliation itself failed; the summary cell says why.
'
' ERROR POLICY
'   Never raises: the SA-CCR results stand whatever happens here. An error
'   leaves the reconciliation sheet cleared, its fingerprint empty and the
'   reason in its summary cell.
'
' UPDATED
'   2026-10-15
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim wsBank      As Worksheet    'The bank sheet
    Dim wsRecon     As Worksheet    'The reconciliation sheet
    Dim wsCalc      As Worksheet    'Results or TradeCalc
    Dim headerRow   As Long         'Header row of the bank table
    Dim problem     As String       'Why the reconciliation cannot run; "" when it can
    Dim settings(1 To 3) As Variant 'Scale, absolute and relative tolerance
    Dim i           As Long         'Setting
    Dim lastR       As Long         'Last row read
    Dim bank        As Variant      'Bank rows; Empty when none
    Dim calc        As Variant      'Calculator rows; Empty when none
    Dim counts()    As Long         'Counts from the comparison
    Dim found       As Variant      'Rows from the comparison; Empty when none
    Dim outRows()   As Variant      'The same as an array for WriteBlock
    Dim nRows       As Long         'Rows to write
    Dim summary     As String       'Summary line without its time stamp
    Dim reconName   As String       'Name of the reconciliation in the run message
    Dim others      As Long         'Rows other than MATCH and DIFFERENCE

'------------------------------------------------------------------------------
' FIND THE SHEETS AND SETTINGS
'------------------------------------------------------------------------------
        On Error GoTo Failed
        reconName = IIf(kind = RECON_NS, "Reconciliation", "Trade reconciliation")
        If Not ReconciliationSheetsPresent(kind) Then
            ReconcileAfterRun = "UNAVAILABLE"
            Exit Function
        End If
        Set wsBank = GetSheet(BankSheetOf(kind))
        Set wsRecon = GetSheet(ReconSheetOf(kind))
        headerRow = BankHeaderRow(kind)
        If headerRow = 0 Then
            problem = "the name " & HeaderNameOf(kind) & " is missing or not in column C of " & BankSheetOf(kind)
        Else
            problem = CORE_Reconcile.HeadersProblem( _
                      TableRange(wsBank, headerRow, 1, headerRow, BankColsOf(kind)).Value2, BankHeadersOf(kind))
        End If
        For i = 1 To 3
            If Len(problem) = 0 Then
                problem = NamedValue(SettingNameOf(kind, i), BankSheetOf(kind), settings(i))
            End If
        Next i
        If Len(problem) = 0 Then
            problem = CORE_Reconcile.SettingsProblem(settings(1), settings(2), settings(3), _
                      SettingNameOf(kind, 1), SettingNameOf(kind, 2), SettingNameOf(kind, 3))
        End If
        If Len(problem) > 0 Then
            ClearReconciliation "Reconciliation not run " & Format$(Now, "yyyy-mm-dd hh:mm:ss") & ": " & problem & ".", _
                                kind
            AddReconLine reconName & " not run: " & problem & "."
            ReconcileAfterRun = "FAILED"
            Exit Function
        End If

'------------------------------------------------------------------------------
' READ THE BANK FIGURES AND THE CALCULATOR ROWS
'------------------------------------------------------------------------------
        lastR = UsedLastRow(wsBank)
        If lastR > headerRow Then
            bank = TableRange(wsBank, headerRow + 1, 1, lastR, BankColsOf(kind)).Value2
        End If
        If Not CORE_Reconcile.HasBankRows(bank, BankColsOf(kind)) Then
            ClearReconciliation "No bank figures on " & BankSheetOf(kind) & ": nothing to reconcile.", kind
            ReconcileAfterRun = "NONE"
            Exit Function
        End If
        If kind = RECON_NS Then
            Set wsCalc = GetSheet(SH_RESULTS)
            lastR = LastDataRow(wsCalc, FIRST_DATA_ROW, RS_ID)
            If lastR >= FIRST_DATA_ROW Then
                calc = TableRange(wsCalc, FIRST_DATA_ROW, 1, lastR, RS_NCOLS).Value2
            End If
        Else
            Set wsCalc = GetSheet(SH_TRADECALC)
            lastR = LastDataRow(wsCalc, FIRST_DATA_ROW, TC_ID)
            If lastR >= FIRST_DATA_ROW Then
                calc = TableRange(wsCalc, FIRST_DATA_ROW, 1, lastR, TC_NCOLS).Value2
            End If
        End If

'------------------------------------------------------------------------------
' COMPARE AND WRITE
'------------------------------------------------------------------------------
        If kind = RECON_NS Then
            found = CORE_Reconcile.CompareNettingSets(bank, headerRow + 1, calc, CDbl(settings(1)), _
                                                      CDbl(settings(2)), CDbl(settings(3)), counts)
            summary = CORE_Reconcile.SummaryText(counts, CDbl(settings(1)), CDbl(settings(2)), CDbl(settings(3)))
        Else
            found = CORE_Reconcile.CompareTrades(bank, headerRow + 1, calc, CDbl(settings(1)), _
                                                 CDbl(settings(2)), CDbl(settings(3)), counts)
            summary = CORE_Reconcile.TradeSummaryText(counts, CDbl(settings(1)), CDbl(settings(2)), _
                                                      CDbl(settings(3)))
        End If
        ClearOutputBlock wsRecon, FIRST_DATA_ROW, RC_NCOLS
        If IsArray(found) Then
            outRows = found
            nRows = UBound(outRows, 1)
            WriteBlock wsRecon, FIRST_DATA_ROW, outRows, nRows, RC_NCOLS
            FormatColumns wsRecon, FIRST_DATA_ROW, nRows, "@|0|@|@|@|@|#,##0.00########|#,##0.00########|" & _
                          "#,##0.00########|0.0000%|@|@"
        End If
        wsRecon.Range(RECON_INFO_CELL).Value = "Last reconciliation " & Format$(Now, "yyyy-mm-dd hh:mm:ss") & _
                                               " | " & summary
        RememberReconInputs kind, ReconFingerprint(kind)
        others = counts(RCN_MISSING) + counts(RCN_NOTINBANK) + counts(RCN_DUPLICATE) + counts(RCN_REGIME) + _
                 counts(RCN_WITHHELD) + counts(RCN_INVALID) + counts(RCN_EXCLUDED) + counts(RCN_NA)
        AddReconLine reconName & ": " & counts(RCN_COMPARED) & " figure(s) compared, " & counts(RCN_MATCH) & _
                     " MATCH, " & counts(RCN_DIFF) & " DIFFERENCE" & _
                     IIf(others > 0, "; " & others & " other row(s): see the " & ReconSheetOf(kind) & " sheet", "") & "."
        ReconcileAfterRun = "DONE"
        Exit Function

'------------------------------------------------------------------------------
' HANDLE ERROR
'------------------------------------------------------------------------------
Failed:
        ClearReconciliation "Reconciliation failed " & Format$(Now, "yyyy-mm-dd hh:mm:ss") & " (error " & _
                            Err.Number & ": " & Err.Description & ").", kind
        AddReconLine reconName & " failed: " & Err.Description & "."
        ReconcileAfterRun = "FAILED"

End Function


Private Sub AddReconLine( _
    ByVal lineText As String)
'
'==============================================================================
'                                 AddReconLine
'------------------------------------------------------------------------------
' PURPOSE
'   Add a line about a reconciliation to mReconLine, the text of the run
'   message.
'
' UPDATED
'   2026-10-15
'==============================================================================
'
        If Len(mReconLine) > 0 Then
            mReconLine = mReconLine & vbCrLf
        End If
        mReconLine = mReconLine & lineText

End Sub


Private Sub ClearReconciliation( _
    ByVal infoText As String, _
    Optional ByVal kind As Long = 0)
'
'==============================================================================
'                             ClearReconciliation
'------------------------------------------------------------------------------
' PURPOSE
'   Withdraw a reconciliation: empty its fingerprint and its rows and write
'   infoText to its summary cell. Every operation starts with this for
'   both, so that no reconciliation outlives the results it compared.
'
' INPUTS
'   infoText: the text for the summary cell.
'   kind: RECON_NS or RECON_TR; 0 for both.
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
    Dim k    As Long         'Reconciliation cleared
    Dim ws   As Worksheet    'Its sheet

'------------------------------------------------------------------------------
' CLEAR
'------------------------------------------------------------------------------
        On Error Resume Next
        For k = RECON_NS To RECON_TR
            If kind = 0 Or kind = k Then
                ForgetReconInputs k
                If ReconciliationSheetsPresent(k) Then
                    Set ws = GetSheet(ReconSheetOf(k))
                    ClearOutputBlock ws, FIRST_DATA_ROW, RC_NCOLS
                    ws.Range(RECON_INFO_CELL).Value = infoText
                End If
            End If
        Next k
        Err.Clear

End Sub


Private Function ReconFingerprint( _
    ByVal kind As Long) _
    As String
'
'==============================================================================
'                               ReconFingerprint
'------------------------------------------------------------------------------
' PURPOSE
'   Fingerprint what a reconciliation depends on: the run inputs
'   (CORE_Engine.InputFingerprint), the bank table from its header row,
'   without the comment column, and the scale and tolerance cells.
'
' INPUTS
'   kind: RECON_NS or RECON_TR.
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
        headerRow = BankHeaderRow(kind)
        If headerRow = 0 Then
            Err.Raise ERR_PARAMS_LAYOUT, "M_Main.ReconFingerprint", _
                      "The name " & HeaderNameOf(kind) & " is missing or not in column C of " & BankSheetOf(kind) & "."
        End If
        parts(1) = CORE_Engine.InputFingerprint()
        parts(2) = CORE_Engine.BlockText(GetSheet(BankSheetOf(kind)), headerRow, BankColsOf(kind) - 1)
        For i = 1 To 3
            Set c = NamedCell(SettingNameOf(kind, i))
            If c Is Nothing Then
                Err.Raise ERR_PARAMS_LAYOUT, "M_Main.ReconFingerprint", _
                          "The name " & SettingNameOf(kind, i) & " is missing."
            End If
            parts(2 + i) = CORE_Engine.BlockText(c.Worksheet, c.Row, 1, c.Column - COL_OFFSET, c.Row)
        Next i
        ReconFingerprint = TextHash(Join(parts, Chr$(29)))

End Function


Private Sub RememberReconInputs( _
    ByVal kind As Long, _
    ByVal fingerprint As String)
'
'==============================================================================
'                             RememberReconInputs
'------------------------------------------------------------------------------
' PURPOSE
'   Store the fingerprint of a completed reconciliation in its hidden
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
        ThisWorkbook.Names.Add Name:=InputsNameOf(kind), RefersTo:="=""" & fingerprint & """", Visible:=False
        Exit Sub

Failed:
        ForgetReconInputs kind

End Sub


Private Sub ForgetReconInputs( _
    ByVal kind As Long)
'
'==============================================================================
'                              ForgetReconInputs
'------------------------------------------------------------------------------
' PURPOSE
'   Empty the stored fingerprint of a reconciliation. The name is kept,
'   empty, because the status formula of its sheet refers to it.
'
' ERROR POLICY
'   Contains any error.
'
' UPDATED
'   2026-10-15
'==============================================================================
'
        On Error GoTo Failed
        ThisWorkbook.Names.Add Name:=InputsNameOf(kind), RefersTo:="=""""", Visible:=False

Failed:

End Sub


Private Function ReconciliationSheetsPresent( _
    ByVal kind As Long) _
    As Boolean
'
'==============================================================================
'                         ReconciliationSheetsPresent
'------------------------------------------------------------------------------
' PURPOSE
'   Tell whether the workbook has both sheets of a reconciliation. A
'   workbook built from an older template lacks them, and still runs.
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
    Dim wsBank    As Worksheet    'The bank sheet; Nothing when absent
    Dim wsRecon   As Worksheet    'The reconciliation sheet; Nothing when absent

'------------------------------------------------------------------------------
' LOOK UP
'------------------------------------------------------------------------------
        On Error Resume Next
        Set wsBank = ThisWorkbook.Worksheets(BankSheetOf(kind))
        Set wsRecon = ThisWorkbook.Worksheets(ReconSheetOf(kind))
        Err.Clear
        On Error GoTo 0
        ReconciliationSheetsPresent = Not wsBank Is Nothing And Not wsRecon Is Nothing

End Function


Private Function BankHeaderRow( _
    ByVal kind As Long) _
    As Long
'
'==============================================================================
'                                BankHeaderRow
'------------------------------------------------------------------------------
' PURPOSE
'   Return the header row of a bank table, from the workbook name of its
'   first header cell, which must be in column C of the bank sheet.
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
        Set c = NamedCell(HeaderNameOf(kind))
        If Not c Is Nothing Then
            If StrComp(c.Worksheet.Name, BankSheetOf(kind), vbTextCompare) = 0 And c.Column = COL_OFFSET + 1 Then
                BankHeaderRow = c.Row
            End If
        End If

End Function


Private Function NamedValue( _
    ByVal nameText As String, _
    ByVal sheetName As String, _
    ByRef valueRead As Variant) _
    As String
'
'==============================================================================
'                                  NamedValue
'------------------------------------------------------------------------------
' PURPOSE
'   Read the value of a workbook name's cell on a given sheet.
'
' INPUTS
'   nameText: the workbook name; sheetName: the sheet it must be on.
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
        ElseIf StrComp(c.Worksheet.Name, sheetName, vbTextCompare) <> 0 Then
            NamedValue = "the name " & nameText & " is not on " & sheetName
        Else
            valueRead = c.Value2
        End If

End Function


Private Function BankSheetOf( _
    ByVal kind As Long) _
    As String
'
'==============================================================================
' PURPOSE
'   The bank sheet of a reconciliation.
'==============================================================================
'
        BankSheetOf = IIf(kind = RECON_NS, SH_BANK_NS, SH_BANK_TR)

End Function


Private Function ReconSheetOf( _
    ByVal kind As Long) _
    As String
'
'==============================================================================
' PURPOSE
'   The output sheet of a reconciliation.
'==============================================================================
'
        ReconSheetOf = IIf(kind = RECON_NS, SH_RECON, SH_TRADE_RECON)

End Function


Private Function HeaderNameOf( _
    ByVal kind As Long) _
    As String
'
'==============================================================================
' PURPOSE
'   The workbook name of the first header cell of a bank table.
'==============================================================================
'
        HeaderNameOf = IIf(kind = RECON_NS, NM_BN_HEADER, NM_BT_HEADER)

End Function


Private Function SettingNameOf( _
    ByVal kind As Long, _
    ByVal i As Long) _
    As String
'
'==============================================================================
' PURPOSE
'   The workbook name of setting i of a bank sheet: 1 the scale, 2 the
'   absolute tolerance, 3 the relative tolerance.
'==============================================================================
'
        If kind = RECON_NS Then
            SettingNameOf = Choose(i, NM_BN_SCALE, NM_BN_ABSTOL, NM_BN_RELTOL)
        Else
            SettingNameOf = Choose(i, NM_BT_SCALE, NM_BT_ABSTOL, NM_BT_RELTOL)
        End If

End Function


Private Function BankColsOf( _
    ByVal kind As Long) _
    As Long
'
'==============================================================================
' PURPOSE
'   The number of columns of a bank table; the last is its comment.
'==============================================================================
'
        BankColsOf = IIf(kind = RECON_NS, BN_NCOLS, BT_NCOLS)

End Function


Private Function BankHeadersOf( _
    ByVal kind As Long) _
    As String
'
'==============================================================================
' PURPOSE
'   The expected headers of a bank table.
'==============================================================================
'
        BankHeadersOf = IIf(kind = RECON_NS, BN_HEADERS, BT_HEADERS)

End Function


Private Function InputsNameOf( _
    ByVal kind As Long) _
    As String
'
'==============================================================================
' PURPOSE
'   The hidden workbook name holding a reconciliation's fingerprint.
'==============================================================================
'
        InputsNameOf = IIf(kind = RECON_NS, RECON_INPUTS_NAME, TRADE_RECON_INPUTS_NAME)

End Function


'
'------------------------------------------------------------------------------
'
'                               EVIDENCE EXPORT
'
'------------------------------------------------------------------------------
'

Public Sub ExportEvidence()
'
'==============================================================================
'                                ExportEvidence
'------------------------------------------------------------------------------
' PURPOSE
'   Write the evidence report of the last run to a CSV file chosen by the
'   user (#45): run identity, parameters, Results, both reconciliations,
'   Checks and the report's limitations (CORE_Export). Refused when the
'   results or a reconciliation with bank figures are not current, or when
'   an output cell was edited after the run.
'
' USAGE
'   Assigned to the Export evidence button on Reconciliation.
'
' ERROR POLICY
'   Reports a refusal or a write failure in a message, or in the Immediate
'   window in silent mode; never raises.
'
' UPDATED
'   2026-10-16
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim problem   As String     'Why the report cannot be written; "" when it can
    Dim target    As Variant    'Chosen file, or False when cancelled
    Dim nLines    As Long       'Lines written

'------------------------------------------------------------------------------
' CHECK, ASK AND WRITE
'------------------------------------------------------------------------------
        On Error GoTo Failed
        problem = EvidenceProblem()
        If Len(problem) > 0 Then
            Notify "Evidence not exported: " & problem & ".", vbExclamation
            Exit Sub
        End If
        target = Application.GetSaveAsFilename(InitialFileName:=DefaultEvidenceName(), _
                                               FileFilter:="CSV files (*.csv), *.csv", _
                                               Title:="Export evidence")
        If VarType(target) = vbBoolean Then
            Exit Sub
        End If
        nLines = WriteEvidence(CStr(target))
        Notify "Evidence exported: " & CStr(target) & " (" & nLines & " lines).", vbInformation
        Exit Sub

Failed:
        Notify "Evidence not exported: " & Err.Description, vbCritical

End Sub


Public Function ExportEvidence_Silent( _
    ByVal filePath As String) _
    As String
'
'==============================================================================
'                            ExportEvidence_Silent
'------------------------------------------------------------------------------
' PURPOSE
'   ExportEvidence for automation and tests: write the report to filePath
'   without asking.
'
' RETURNS
'   "RESULT=OK; lines=N; file=..." or "RESULT=REFUSED; reason=...".
'
' ERROR POLICY
'   A file that cannot be written raises to the caller.
'
' UPDATED
'   2026-10-16
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim problem   As String    'Why the report cannot be written; "" when it can

'------------------------------------------------------------------------------
' CHECK AND WRITE
'------------------------------------------------------------------------------
        problem = EvidenceProblem()
        If Len(problem) > 0 Then
            ExportEvidence_Silent = "RESULT=REFUSED; reason=" & problem
        Else
            ExportEvidence_Silent = "RESULT=OK; lines=" & WriteEvidence(filePath) & "; file=" & filePath
        End If

End Function


Private Function EvidenceProblem() As String
'
'==============================================================================
'                               EvidenceProblem
'------------------------------------------------------------------------------
' PURPOSE
'   Tell why no evidence report may be written now: the results are not
'   those of the current inputs, or a sheet of bank figures holds figures
'   whose reconciliation is missing or out of date. An evidence report
'   must never describe results that do not match the inputs.
'
' RETURNS
'   "" when the report may be written; otherwise the reason.
'
' ERROR POLICY
'   Never raises: an error is itself the reason.
'
' UPDATED
'   2026-10-16
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim kind      As Long       'Reconciliation checked
    Dim stored    As String     'A stored fingerprint
    Dim current   As Boolean    'True when the reconciliation is current

'------------------------------------------------------------------------------
' CHECK
'------------------------------------------------------------------------------
        On Error GoTo Failed
    'In manual calculation a formula input may still show the value from
    'before its precedents changed, and the fingerprints would read it as
    'current.
        If Application.CalculationState <> xlDone Then
            EvidenceProblem = "Excel has not finished recalculating; press F9, then Run SA-CCR if the results " & _
                              "are out of date"
            Exit Function
        End If
        If ResultsStatus() <> "CURRENT" Then
            EvidenceProblem = "the results are not those of the current inputs; press Run SA-CCR first"
            Exit Function
        End If
    'A stored reconciliation must match its bank figures as they are now,
    'also when every bank row was removed after the run; bank figures
    'without a stored reconciliation were never compared.
        For kind = RECON_NS To RECON_TR
            If ReconciliationSheetsPresent(kind) Then
                stored = StoredText(InputsNameOf(kind))
                If Len(stored) > 0 Then
                    current = (stored = ReconFingerprint(kind))
                Else
                    current = Not BankRowsEntered(kind)
                End If
                If Not current Then
                    EvidenceProblem = "the reconciliation of the bank figures on " & BankSheetOf(kind) & _
                                      " is not current; press Run SA-CCR first"
                    Exit Function
                End If
            End If
        Next kind
    'The output sheets are not protected: what the report copies from them
    'must still be what the run wrote.
        stored = StoredText(RUN_OUTPUTS_NAME)
        If Len(stored) = 0 Then
            EvidenceProblem = "the outputs of the last run were not recorded; press Run SA-CCR first"
        ElseIf stored <> OutputFingerprint() Then
            EvidenceProblem = "an output cell on Results, Checks or a reconciliation sheet was changed after " & _
                              "the run; press Run SA-CCR first"
        End If
        Exit Function

Failed:
        EvidenceProblem = "the workbook could not be checked (" & Err.Description & ")"

End Function


Private Function WriteEvidence( _
    ByVal filePath As String) _
    As Long
'
'==============================================================================
'                                WriteEvidence
'------------------------------------------------------------------------------
' PURPOSE
'   Write the evidence report to filePath as UTF-8 without a byte-order
'   mark, replacing any file of that name, hidden ones included; a
'   read-only file is left as it is and the write fails.
'
' RETURNS
'   The number of lines written.
'
' ERROR POLICY
'   Raises when the file cannot be written, or when it holds more than
'   the report afterwards (an old file that could not be removed); the
'   file number is closed first.
'
' UPDATED
'   2026-10-16
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim text        As String    'The report
    Dim bytes()     As Byte      'Its UTF-8 bytes
    Dim fileNo      As Integer   'File number; 0 while none is open
    Dim errNumber   As Long      'Error to raise after closing the file
    Dim errText     As String    'Its description

'------------------------------------------------------------------------------
' WRITE
'------------------------------------------------------------------------------
        text = EvidenceText()
        bytes = CORE_Export.Utf8Bytes(text)
    'Open For Binary does not truncate: an old file must go first, also
    'when it is hidden, or a shorter report would keep its tail.
        On Error GoTo Failed
        If Len(Dir$(filePath, vbHidden Or vbSystem)) > 0 Then
            If (GetAttr(filePath) And vbReadOnly) = 0 Then
                SetAttr filePath, vbNormal
            End If
            Kill filePath
        End If
        fileNo = FreeFile
        Open filePath For Binary Access Write As #fileNo
        Put #fileNo, , bytes
        If LOF(fileNo) <> UBound(bytes) - LBound(bytes) + 1 Then
            Err.Raise ERR_EXPORT_WRITE, "M_Main.WriteEvidence", "an older file of that name could not be replaced"
        End If
        Close #fileNo
        WriteEvidence = (Len(text) - Len(Replace(text, vbCrLf, ""))) \ 2
        Exit Function

Failed:
        errNumber = Err.Number
        errText = Err.Description
        If fileNo <> 0 Then
            Close #fileNo
        End If
        Err.Raise errNumber, "M_Main.WriteEvidence", "Could not write " & filePath & ": " & errText

End Function


Private Function EvidenceText() As String
'
'==============================================================================
'                                 EvidenceText
'------------------------------------------------------------------------------
' PURPOSE
'   Build the evidence report from the sheets as they are: RUN (the run
'   identity), PARAMS, SF_TABLE and FX_TABLE (the parameter snapshot),
'   RESULTS, RECONCILIATION, TRADE_RECONCILIATION, CHECKS and LIMITATIONS.
'   Everything is read from the workbook, so the report is the same after
'   the workbook is closed and reopened.
'
' UPDATED
'   2026-10-16
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim s          As String       'The report
    Dim wsRes      As Worksheet    'Results
    Dim lastR      As Long         'Last row of a table
    Dim r          As Long         'Row
    Dim id         As String       'Results ID
    Dim codes      As Variant      'Parameter codes
    Dim i          As Long         'Code or reconciliation
    Dim nSev(1 To 3) As Long       'Checks lines: ERROR, WARNING, INFO
    Dim sev        As String       'Severity of a Checks line
    Dim prefix     As String       'Key prefix of a reconciliation
    Dim c          As Range        'A named cell

'------------------------------------------------------------------------------
' RUN
'------------------------------------------------------------------------------
        Set wsRes = GetSheet(SH_RESULTS)
        s = RunLine("format", "K-SACCR evidence 1")
        s = s & RunLine("workbook", ThisWorkbook.Name)
        s = s & RunLine("run_summary", wsRes.Range(RUNINFO_CELL).Value2)
        s = s & RunLine("results_status", ResultsStatus())
        s = s & RunLine("inputs", CORE_Engine.LastRunInputs())
        s = s & RunLine("outputs", StoredText(RUN_OUTPUTS_NAME))
        Set c = NamedCell(PRM_ASOF)
        If IsNumeric(c.Value2) And Not IsEmpty(c.Value2) Then
            s = s & RunLine("reporting_date", Format$(CDate(c.Value2), "yyyy-mm-dd"))
        Else
            s = s & RunLine("reporting_date", c.Value2)
        End If
        s = s & RunLine("reporting_ccy", NamedCell(PRM_REPCCY).Value2)
        s = s & RunLine("regime_default", NamedCell(PRM_REGIME).Value2)
        s = s & RunLine("parameter_set", REF_SET_ID)
        lastR = LastDataRow(wsRes, FIRST_DATA_ROW, RS_ID)
        For r = FIRST_DATA_ROW To lastR
            id = UTxt(TableCell(wsRes, r, RS_ID).Value2)
            If id = "TOTAL " & RG_CRR Or id = "TOTAL " & RG_BCBS Then
                s = s & RunLine("total_ead_" & LCase$(Mid$(id, 7)), TableCell(wsRes, r, RS_EAD).Value2)
                s = s & RunLine("total_status_" & LCase$(Mid$(id, 7)), TableCell(wsRes, r, RS_STATUS_COL).Value2)
            End If
        Next r
        lastR = LastDataRow(GetSheet(SH_CHECKS), FIRST_DATA_ROW, 1)
        For r = FIRST_DATA_ROW To lastR
            sev = UTxt(TableCell(GetSheet(SH_CHECKS), r, 1).Value2)
            If sev = SEV_ERROR Then nSev(1) = nSev(1) + 1
            If sev = SEV_WARN Then nSev(2) = nSev(2) + 1
            If sev = SEV_INFO Then nSev(3) = nSev(3) + 1
        Next r
        s = s & RunLine("checks_errors", nSev(1))
        s = s & RunLine("checks_warnings", nSev(2))
        s = s & RunLine("checks_info", nSev(3))
        For i = RECON_NS To RECON_TR
            prefix = IIf(i = RECON_NS, "reconciliation", "trade_reconciliation")
            If Not ReconciliationSheetsPresent(i) Then
                s = s & RunLine(prefix, "UNAVAILABLE")
            Else
                s = s & RunLine(prefix, IIf(Len(StoredText(InputsNameOf(i))) > 0, "CURRENT", "NONE"))
                s = s & RunLine(prefix & "_inputs", StoredText(InputsNameOf(i)))
                s = s & RunLine(prefix & "_summary", GetSheet(ReconSheetOf(i)).Range(RECON_INFO_CELL).Value2)
            'The settings are part of a reconciliation's fingerprint; without
            'one they were not used, and could change after the run.
                If Len(StoredText(InputsNameOf(i))) > 0 Then
                    s = s & RunLine(prefix & "_scale", NamedCell(SettingNameOf(i, 1)).Value2)
                    s = s & RunLine(prefix & "_abs_tol", NamedCell(SettingNameOf(i, 2)).Value2)
                    s = s & RunLine(prefix & "_rel_tol", NamedCell(SettingNameOf(i, 3)).Value2)
                End If
            End If
        Next i

'------------------------------------------------------------------------------
' PARAMETER SNAPSHOT
'------------------------------------------------------------------------------
        codes = CORE_Engine.ParameterCodes()
        For i = LBound(codes) To UBound(codes)
            s = s & CORE_Export.CsvLine("PARAMS", Array(codes(i), NamedCell(CStr(codes(i))).Value2))
        Next i
        s = s & ParamsTableLines("SF_TABLE", NM_SF_TABLE, SF_HEADERS)
        s = s & ParamsTableLines("FX_TABLE", NM_FX_TABLE, FX_HEADERS)

'------------------------------------------------------------------------------
' TABLES
'------------------------------------------------------------------------------
        s = s & OutputTableLines()
        s = s & CORE_Export.LimitationLines()
        EvidenceText = s

End Function


Private Function ParamsTableLines( _
    ByVal section As String, _
    ByVal nameText As String, _
    ByVal headers As String) _
    As String
'
'==============================================================================
'                               ParamsTableLines
'------------------------------------------------------------------------------
' PURPOSE
'   A Params table as the engine reads it: from the header row of its
'   workbook name down to the first blank key, across its header columns.
'   Rows added below the named range are read by the run, so they are
'   exported too.
'
' ERROR POLICY
'   Raises when the name does not locate a header row.
'
' UPDATED
'   2026-10-16
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ws        As Worksheet    'Params
    Dim headRow   As Long         'Header row
    Dim keyCol    As Long         'Key column
    Dim r         As Long         'Row read

'------------------------------------------------------------------------------
' READ
'------------------------------------------------------------------------------
        Set ws = GetSheet(SH_PARAMS)
        headRow = NamedRow(nameText)
        keyCol = NamedTableCol(nameText)
        If headRow = 0 Then
            Err.Raise ERR_PARAMS_LAYOUT, "M_Main.ParamsTableLines", "The name " & nameText & " is missing."
        End If
        r = headRow + 1
        Do While Not IsBlankCell(TableCell(ws, r, keyCol).Value)
            r = r + 1
        Loop
        ParamsTableLines = CORE_Export.TableLines(section, TableRange(ws, headRow, keyCol, r - 1, _
                                                  keyCol + UBound(Split(headers, "|"))).Value2)

End Function


Private Function SheetTableLines( _
    ByVal section As String, _
    ByVal sheetName As String, _
    ByVal nCols As Long) _
    As String
'
'==============================================================================
'                               SheetTableLines
'------------------------------------------------------------------------------
' PURPOSE
'   Write an output table as lines of a section: its header row, then its
'   rows down to the last with a value in the first column.
'
' UPDATED
'   2026-10-16
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ws      As Worksheet    'The sheet
    Dim lastR   As Long         'Last row of the table

'------------------------------------------------------------------------------
' WRITE
'------------------------------------------------------------------------------
        Set ws = GetSheet(sheetName)
        lastR = LastDataRow(ws, FIRST_DATA_ROW, 1)
        If lastR < HEADER_ROW Then
            lastR = HEADER_ROW
        End If
        SheetTableLines = CORE_Export.TableLines(section, TableRange(ws, HEADER_ROW, 1, lastR, nCols).Value2)

End Function


Private Function RunLine( _
    ByVal key As String, _
    ByVal entry As Variant) _
    As String
'
'==============================================================================
' PURPOSE
'   One line of the RUN section: key and value.
'==============================================================================
'
        RunLine = CORE_Export.CsvLine("RUN", Array(key, entry))

End Function


Private Function DefaultEvidenceName() As String
'
'==============================================================================
'                             DefaultEvidenceName
'------------------------------------------------------------------------------
' PURPOSE
'   The file name offered for a report: K-SACCR_evidence_ with the date and
'   time, in the workbook's folder when it has one.
'
' UPDATED
'   2026-10-16
'==============================================================================
'
        DefaultEvidenceName = "K-SACCR_evidence_" & Format$(Now, "yyyymmdd-hhnnss") & ".csv"
    'A workbook opened from OneDrive or SharePoint has a URL as its path,
    'which the save dialog cannot start from.
        If Len(ThisWorkbook.Path) > 0 And InStr(ThisWorkbook.Path, "://") = 0 Then
            DefaultEvidenceName = ThisWorkbook.Path & Application.PathSeparator & DefaultEvidenceName
        End If

End Function


Private Function BankRowsEntered( _
    ByVal kind As Long) _
    As Boolean
'
'==============================================================================
'                               BankRowsEntered
'------------------------------------------------------------------------------
' PURPOSE
'   Tell whether a bank sheet holds any figure: an unreadable table counts
'   as holding figures, so that it is never skipped silently.
'
' UPDATED
'   2026-10-16
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ws          As Worksheet    'The bank sheet
    Dim headerRow   As Long         'Its header row
    Dim lastR       As Long         'Its last used row

'------------------------------------------------------------------------------
' READ
'------------------------------------------------------------------------------
        headerRow = BankHeaderRow(kind)
        If headerRow = 0 Then
            BankRowsEntered = True
            Exit Function
        End If
        Set ws = GetSheet(BankSheetOf(kind))
        lastR = UsedLastRow(ws)
        If lastR > headerRow Then
            BankRowsEntered = CORE_Reconcile.HasBankRows( _
                TableRange(ws, headerRow + 1, 1, lastR, BankColsOf(kind)).Value2, BankColsOf(kind))
        End If

End Function


Private Function StoredText( _
    ByVal nameText As String) _
    As String
'
'==============================================================================
'                                  StoredText
'------------------------------------------------------------------------------
' PURPOSE
'   Return the fingerprint stored in a hidden workbook name (a
'   reconciliation's or the run outputs'), or "" when there is none, as
'   CORE_Engine.LastRunInputs.
'
' ERROR POLICY
'   Contains the expected error of a missing name.
'
' UPDATED
'   2026-10-16
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim wbName   As Name    'The stored name, Nothing when absent

'------------------------------------------------------------------------------
' READ
'------------------------------------------------------------------------------
        On Error Resume Next
        Set wbName = ThisWorkbook.Names(nameText)
        Err.Clear
        On Error GoTo 0
        If Not wbName Is Nothing Then
            StoredText = Replace(Replace(wbName.RefersTo, "=", ""), """", "")
        End If

End Function


Private Function OutputTableLines() As String
'
'==============================================================================
'                               OutputTableLines
'------------------------------------------------------------------------------
' PURPOSE
'   The output tables of the evidence report: RESULTS, RECONCILIATION and
'   TRADE_RECONCILIATION when their sheets exist, and CHECKS.
'
' UPDATED
'   2026-10-16
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim s   As String    'The lines

'------------------------------------------------------------------------------
' READ
'------------------------------------------------------------------------------
        s = SheetTableLines("RESULTS", SH_RESULTS, RS_NCOLS)
        If ReconciliationSheetsPresent(RECON_NS) Then
            s = s & SheetTableLines("RECONCILIATION", SH_RECON, RC_NCOLS)
        End If
        If ReconciliationSheetsPresent(RECON_TR) Then
            s = s & SheetTableLines("TRADE_RECONCILIATION", SH_TRADE_RECON, RC_NCOLS)
        End If
        OutputTableLines = s & SheetTableLines("CHECKS", SH_CHECKS, CK_NCOLS)

End Function


Private Function OutputFingerprint() As String
'
'==============================================================================
'                              OutputFingerprint
'------------------------------------------------------------------------------
' PURPOSE
'   Fingerprint what the evidence report copies from the output sheets:
'   the run summary on Results, the summary of each reconciliation sheet
'   and the output tables.
'
' ERROR POLICY
'   Raises when a sheet cannot be read.
'
' UPDATED
'   2026-10-16
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim s      As String    'The text fingerprinted
    Dim kind   As Long      'Reconciliation

'------------------------------------------------------------------------------
' FINGERPRINT
'------------------------------------------------------------------------------
        s = CORE_Export.CsvField(GetSheet(SH_RESULTS).Range(RUNINFO_CELL).Value2)
        For kind = RECON_NS To RECON_TR
            If ReconciliationSheetsPresent(kind) Then
                s = s & Chr$(29) & CORE_Export.CsvField(GetSheet(ReconSheetOf(kind)).Range(RECON_INFO_CELL).Value2)
            End If
        Next kind
        OutputFingerprint = TextHash(s & Chr$(29) & OutputTableLines())

End Function


Private Sub RememberRunOutputs()
'
'==============================================================================
'                              RememberRunOutputs
'------------------------------------------------------------------------------
' PURPOSE
'   Store the fingerprint of the output sheets when a run has completed,
'   its reconciliations included.
'
' ERROR POLICY
'   Best effort: when the fingerprint cannot be taken or stored it is left
'   empty, and the export refuses until the next run.
'
' UPDATED
'   2026-10-16
'==============================================================================
'
        On Error GoTo Failed
        ThisWorkbook.Names.Add Name:=RUN_OUTPUTS_NAME, RefersTo:="=""" & OutputFingerprint() & """", _
                               Visible:=False
        Exit Sub

Failed:
        ForgetRunOutputs

End Sub


Private Sub ForgetRunOutputs()
'
'==============================================================================
'                               ForgetRunOutputs
'------------------------------------------------------------------------------
' PURPOSE
'   Empty the stored fingerprint of the output sheets.
'
' ERROR POLICY
'   Contains any error.
'
' UPDATED
'   2026-10-16
'==============================================================================
'
        On Error GoTo Failed
        ThisWorkbook.Names.Add Name:=RUN_OUTPUTS_NAME, RefersTo:="=""""", Visible:=False

Failed:

End Sub


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
' WITHDRAW THE RECONCILIATION
'------------------------------------------------------------------------------
    'Every operation first withdraws both reconciliations (#45) and the
    'fingerprint of the output sheets, before anything that can stop it,
    'the recalculation included; only a run that completes makes new ones.
        mReconState = "NONE"
        mTradeReconState = "NONE"
        mReconLine = ""
        ForgetRunOutputs
        Select Case operation
            Case OP_RUN
                ClearReconciliation "Reconciliation pending: the SA-CCR run has not completed."
            Case OP_VALIDATE
                ClearReconciliation "Outputs cleared by validation: press Run SA-CCR to reconcile."
            Case OP_CLEAR
                ClearReconciliation "Outputs cleared " & Format$(Now, "yyyy-mm-dd hh:mm:ss")
        End Select

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
        Select Case operation
            Case OP_RUN
                ok = CORE_Engine.Calculate(True)
                If ok Then
                    mReconState = ReconcileAfterRun(RECON_NS)
                    mTradeReconState = ReconcileAfterRun(RECON_TR)
                    RememberRunOutputs
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
'   Recalculate the status formulas on Results, Reconciliation and
'   TradeReconciliation after an operation has stored or withdrawn the run
'   fingerprint. In manual calculation Excel
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
        If ReconciliationSheetsPresent(RECON_NS) Then
            GetSheet(SH_RECON).Range(RECON_STATUS_CELL).Calculate
        End If
        If ReconciliationSheetsPresent(RECON_TR) Then
            GetSheet(SH_TRADE_RECON).Range(RECON_STATUS_CELL).Calculate
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
'   reconciliation=DONE; trade_reconciliation=NONE; cleanup=PASS".
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
                     "; reconciliation=" & mReconState & _
                     "; trade_reconciliation=" & mTradeReconState
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
