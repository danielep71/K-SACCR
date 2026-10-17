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
'   inputs and are still as the run wrote them (CURRENT, EDITED, STALE or
'   NONE).
'   ResultsStatusText: the input part of that status as text, for the
'   status formula in Results C5; Excel recalculates it when an input
'   changes, without a macro writing to the workbook, so Undo keeps
'   working. An edited output is left to ResultsStatus and the evidence
'   export (#315).
'   ReconciliationStatusText and TradeReconciliationStatusText: the same
'   for the reconciliations of the bank figures, for the status formulas
'   in Reconciliation C5 and TradeReconciliation C5 (#45).
'   ExportEvidence: write the evidence report of the last run to a CSV
'   file (#45); ExportEvidence_Silent: the same to a given file, for
'   automation and tests.
'   The sheet buttons call the first three and ExportEvidence by name.
'   These macros are not listed in docs/PUBLIC_API.txt.
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
'   mRunning, the reconciliation outcomes of the last operation
'   (mReconState, mTradeReconState, mReconLine), and the test seams
'   gTestFault, gQuietLog (silent-mode messages dropped instead of
'   printed) and gLastMessage.
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
'   thrown by it would (#204); gTestFault = "reconcile" raises
'   ERR_INJECTED_FAULT in a reconciliation that has bank figures (#318);
'   gTestFault = "evidence-open", "evidence-write" and "evidence-replace"
'   raise ERR_INJECTED_FAULT in WriteEvidence before the new report is
'   opened, while it is written and while the older file is replaced
'   (#312). tests/modules/TEST_MainState.bas uses the first four and
'   tests/modules/TEST_Reconcile.bas the others. gLastMessage keeps the last
'   message shown or printed, so that a test can read it in silent mode.
'   Production code never sets gTestFault.
'
' KNOWN DEVIATION
'   This standard module lives in src/workbook. See the known deviations in
'   docs/REPOSITORY_STRUCTURE.md.
'
' COMPATIBILITY
'   Excel VBA; no references beyond the defaults.
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
    'caller. gTestFault and gLastMessage are the test seams described above.
    'gQuietLog drops the messages of silent mode instead of printing them,
    'for the test suites (TEST_Suite.Test_RunAll, TEST_MainState and
    'TEST_Reconcile), whose output would overflow the Immediate window.
    'mRunning refuses a second operation while one is in progress.
        Public gSilent      As Boolean    'Silent mode for automation and tests
        Public gTestFault   As String     'Test seam: "", "operation", "cleanup", "recalc", "recalc-error", "reconcile" or "evidence-*"
        Public gQuietLog    As Boolean    'Test seam: silent-mode messages not printed
        Public gLastMessage As String     'Test seam: the last message shown or printed
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
'   The result line of ReportOutcome: RESULT=OK or STOPPED,
'   operation=run, errors, warnings, trades_used, trades_read,
'   economic_trades_used, economic_trades_read, incomplete,
'   total_ead_crr, total_ead_bcbs, inputs, run_class, reconciliation,
'   trade_reconciliation and cleanup=PASS or FAIL, separated by "; ".
'   incomplete counts netting sets whose EAD was withheld; total_ead_crr
'   and total_ead_bcbs are the totals of each regime's VALID netting sets,
'   0 when it has none (#198); inputs is the fingerprint of the inputs
'   the results were calculated from, empty when there are none.
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
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim PreviousSilent     As Boolean    'gSilent before this call
    Dim ErrorNumber        As Long       'Error raised by the operation
    Dim ErrorSource        As String     'Its source
    Dim ErrorDescription   As String     'Its description

'------------------------------------------------------------------------------
' RUN SILENTLY
'------------------------------------------------------------------------------
        PreviousSilent = gSilent
        gSilent = True
        On Error GoTo Failed
        RunSACCR_Silent = ExecuteOperation(OP_RUN)
        gSilent = PreviousSilent
        Exit Function

'------------------------------------------------------------------------------
' RESTORE THE FLAG AND RE-RAISE
'------------------------------------------------------------------------------
Failed:
        ErrorNumber = Err.Number
        ErrorSource = Err.Source
        ErrorDescription = Err.Description
        On Error GoTo 0
        gSilent = PreviousSilent
        Err.Raise ErrorNumber, ErrorSource, ErrorDescription

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
'   Empty the five output tables and both reconciliation sheets, withdraw
'   the stored fingerprints, and stamp the time on the Results sheet.
'
' USAGE
'   Assigned to the Clear outputs button.
'
' UPDATED
'   2026-10-17
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
'   CURRENT: the inputs match the last completed run, and so do the
'   outputs it recorded.
'   EDITED: the inputs match, but an output cell on Results, Checks or a
'   reconciliation sheet was changed after the run (#315); the evidence
'   export refuses such outputs.
'   STALE: the inputs changed after it; the results are out of date.
'   NONE: no valid completed-run fingerprint. Protected output cells may
'   remain after failed cleanup; the run error reports that separately.
'
' METHOD
'   The outputs are compared only when the inputs match, and only here:
'   the status formula in Results C5 does not watch the output sheets,
'   whose fingerprint would be recomputed after every edit (#315). A run
'   whose outputs were not recorded is CURRENT when its inputs match.
'
' ERROR POLICY
'   Raises ERR_PARAMS_LAYOUT from CORE_Engine.InputFingerprint when a
'   Params workbook name is missing or broken (#152), and the error of
'   OutputFingerprint when an output sheet cannot be read.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim StoredInputs    As String    'Fingerprint stored by the last completed run
    Dim StoredOutputs   As String    'Fingerprint of the outputs it wrote

'------------------------------------------------------------------------------
' COMPARE
'------------------------------------------------------------------------------
        StoredInputs = CORE_Engine.LastRunInputs()
        If Len(StoredInputs) = 0 Then
            ResultsStatus = "NONE"
        ElseIf StoredInputs <> CORE_Engine.InputFingerprint() Then
            ResultsStatus = "STALE"
        Else
            StoredOutputs = StoredText(RUN_OUTPUTS_NAME)
            If Len(StoredOutputs) > 0 And StoredOutputs <> OutputFingerprint() Then
                ResultsStatus = "EDITED"
            Else
                ResultsStatus = "CURRENT"
            End If
        End If

End Function


Public Function ResultsStatusText( _
    ByVal StoredInputs As Variant, _
    ByVal NettingSetRows As Range, _
    ByVal TradeRows As Range, _
    ByVal ParamRows As Range) _
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
'   StoredInputs: the value of SACCR_RunInputs.
'   NettingSetRows, TradeRows, ParamRows: the input ranges, only so that
'   Excel recalculates; the fingerprint is computed from the sheets as the
'   engine reads them.
'
' RETURNS
'   "Inputs unchanged since the last run (inputs XXXXXXXX); outputs are
'   checked when evidence is exported", "OUT OF DATE - inputs changed
'   since the last run; press Run SA-CCR." or "No results: press Run
'   SA-CCR." The cell checks the inputs only: an edited output cell is
'   reported by ResultsStatus() and refused by the evidence export (#315).
'
' ERROR POLICY
'   Returns "Status unavailable: " and the reason instead of raising, for
'   example when a Params workbook name is missing (#152): a worksheet
'   function must not interrupt Excel.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim RunFingerprint   As String    'Fingerprint of the last completed run

'------------------------------------------------------------------------------
' COMPARE
'------------------------------------------------------------------------------
        On Error GoTo Failed
        RunFingerprint = SafeStr(StoredInputs)
        If Len(RunFingerprint) = 0 Then
            ResultsStatusText = "No results: press Run SA-CCR."
        ElseIf RunFingerprint = CORE_Engine.InputFingerprint() Then
            ResultsStatusText = "Inputs unchanged since the last run (inputs " & RunFingerprint & _
                                "); outputs are checked when evidence is exported"
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
    ByVal StoredRecon As Variant, _
    ByVal StoredRun As Variant, _
    ByVal BankRows As Range, _
    ByVal NettingSetRows As Range, _
    ByVal TradeRows As Range, _
    ByVal ParamRows As Range) _
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
'   StoredRecon: the value of SACCR_ReconInputs; StoredRun: the value of
'   SACCR_RunInputs.
'   BankRows, NettingSetRows, TradeRows, ParamRows: the watched ranges,
'   only so that Excel recalculates.
'
' RETURNS
'   As ReconStatus.
'
' UPDATED
'   2026-10-17
'==============================================================================
'
        ReconciliationStatusText = ReconStatus(RECON_NS, StoredRecon, StoredRun)

End Function


Public Function TradeReconciliationStatusText( _
    ByVal StoredRecon As Variant, _
    ByVal StoredRun As Variant, _
    ByVal BankRows As Range, _
    ByVal NettingSetRows As Range, _
    ByVal TradeRows As Range, _
    ByVal ParamRows As Range) _
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
'   StoredRecon: the value of SACCR_TradeReconInputs; the others as
'   ReconciliationStatusText, with BankTrades as BankRows.
'
' RETURNS
'   As ReconStatus.
'
' UPDATED
'   2026-10-17
'==============================================================================
'
        TradeReconciliationStatusText = ReconStatus(RECON_TR, StoredRecon, StoredRun)

End Function


Private Function ReconStatus( _
    ByVal ReconKind As Long, _
    ByVal StoredRecon As Variant, _
    ByVal StoredRun As Variant) _
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
'   ReconKind: RECON_NS or RECON_TR; StoredRecon, StoredRun: as the
'   worksheet functions.
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
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim StoredFingerprint   As String    'Fingerprint of the last reconciliation

'------------------------------------------------------------------------------
' COMPARE
'------------------------------------------------------------------------------
    'Without current results there is nothing the reconciliation can still
    'describe, whatever its own fingerprint says.
        On Error GoTo Failed
        StoredFingerprint = SafeStr(StoredRecon)
        If Len(StoredFingerprint) = 0 Or Len(SafeStr(StoredRun)) = 0 Then
            ReconStatus = "No reconciliation: enter bank figures on " & BankSheetOf(ReconKind) & _
                          " and press Run SA-CCR (on Results or Trades)."
        ElseIf StoredFingerprint = ReconFingerprint(ReconKind) Then
            ReconStatus = "Reconciliation current (bank figures and inputs " & StoredFingerprint & ")"
        Else
            ReconStatus = "OUT OF DATE - bank figures or inputs changed since the last run; press Run SA-CCR (on Results or Trades)."
        End If
        Exit Function

Failed:
        ReconStatus = "Status unavailable: " & Err.Description

End Function


Private Function ReconcileAfterRun( _
    ByVal ReconKind As Long) _
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
'   ReconKind: RECON_NS or RECON_TR.
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
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim BankSheet          As Worksheet    'The bank sheet
    Dim ReconSheet         As Worksheet    'The reconciliation sheet
    Dim CalcSheet          As Worksheet    'Results or TradeCalc
    Dim HeaderRow          As Long         'Header row of the bank table
    Dim Problem            As String       'Why the reconciliation cannot run; "" when it can
    Dim Settings(1 To 3)   As Variant      'Scale, absolute and relative tolerance
    Dim i                  As Long         'Setting
    Dim LastRow            As Long         'Last row read
    Dim CalcFirstRow       As Long         'First data row of Results or TradeCalc
    Dim ReconFirstRow      As Long         'First data row of the reconciliation sheet
    Dim BankRows           As Variant      'Bank rows; Empty when none
    Dim CalcRows           As Variant      'Calculator rows; Empty when none
    Dim Counts()           As Long         'Counts from the comparison
    Dim ComparedRows       As Variant      'Rows from the comparison; Empty when none
    Dim ReconRows()        As Variant      'The same as an array for WriteBlock
    Dim RowsToWrite        As Long         'Rows to write
    Dim Summary            As String       'Summary line without its time stamp
    Dim ReconName          As String       'Name of the reconciliation in the run message
    Dim OtherCount         As Long         'Rows other than MATCH and DIFFERENCE
    Dim ErrorNumber        As Long         'Error that stopped the reconciliation
    Dim ErrorText          As String       'Its description

'------------------------------------------------------------------------------
' FIND THE SHEETS AND SETTINGS
'------------------------------------------------------------------------------
        On Error GoTo Failed
        ReconName = IIf(ReconKind = RECON_NS, "Reconciliation", "Trade reconciliation")
        If Not ReconciliationSheetsPresent(ReconKind) Then
            ReconcileAfterRun = "UNAVAILABLE"
            Exit Function
        End If
        Set BankSheet = GetSheet(BankSheetOf(ReconKind))
        Set ReconSheet = GetSheet(ReconSheetOf(ReconKind))
        HeaderRow = BankHeaderRow(ReconKind)
        If HeaderRow = 0 Then
            Problem = "the name " & HeaderNameOf(ReconKind) & " is missing or not in column C of " & BankSheetOf(ReconKind)
        Else
            Problem = CORE_Reconcile.HeadersProblem( _
                      TableRange(BankSheet, HeaderRow, 1, HeaderRow, BankColsOf(ReconKind)).Value2, BankHeadersOf(ReconKind))
        End If
        For i = 1 To 3
            If Len(Problem) = 0 Then
                Problem = NamedValue(SettingNameOf(ReconKind, i), BankSheetOf(ReconKind), Settings(i))
            End If
        Next i
        If Len(Problem) = 0 Then
            Problem = CORE_Reconcile.SettingsProblem(Settings(1), Settings(2), Settings(3), _
                      SettingNameOf(ReconKind, 1), SettingNameOf(ReconKind, 2), SettingNameOf(ReconKind, 3))
        End If
        If Len(Problem) > 0 Then
            ClearReconciliation "Reconciliation not run " & Format$(Now, "yyyy-mm-dd hh:mm:ss") & ": " & Problem & ".", _
                                ReconKind
            AddReconLine ReconName & " not run: " & Problem & "."
            ReconcileAfterRun = "FAILED"
            Exit Function
        End If

'------------------------------------------------------------------------------
' READ THE BANK FIGURES AND THE CALCULATOR ROWS
'------------------------------------------------------------------------------
        LastRow = UsedLastRow(BankSheet)
        If LastRow > HeaderRow Then
            BankRows = TableRange(BankSheet, HeaderRow + 1, 1, LastRow, BankColsOf(ReconKind)).Value2
        End If
        If Not CORE_Reconcile.HasBankRows(BankRows, BankColsOf(ReconKind)) Then
            ClearReconciliation "No bank figures on " & BankSheetOf(ReconKind) & ": nothing to reconcile.", ReconKind
            ReconcileAfterRun = "NONE"
            Exit Function
        End If
        If gTestFault = "reconcile" Then
            Err.Raise ERR_INJECTED_FAULT, "M_Main.ReconcileAfterRun", "Injected reconciliation failure."
        End If
        If ReconKind = RECON_NS Then
            Set CalcSheet = GetSheet(SH_RESULTS)
            CalcFirstRow = OutputFirstRow(SH_RESULTS)
            LastRow = LastDataRow(CalcSheet, CalcFirstRow, RS_ID)
            If LastRow >= CalcFirstRow Then
                CalcRows = TableRange(CalcSheet, CalcFirstRow, 1, LastRow, RS_NCOLS).Value2
            End If
        Else
            Set CalcSheet = GetSheet(SH_TRADECALC)
            CalcFirstRow = OutputFirstRow(SH_TRADECALC)
            LastRow = LastDataRow(CalcSheet, CalcFirstRow, TC_ID)
            If LastRow >= CalcFirstRow Then
                CalcRows = TableRange(CalcSheet, CalcFirstRow, 1, LastRow, TC_NCOLS).Value2
            End If
        End If

'------------------------------------------------------------------------------
' COMPARE AND WRITE
'------------------------------------------------------------------------------
        If ReconKind = RECON_NS Then
            ComparedRows = CORE_Reconcile.CompareNettingSets(BankRows, HeaderRow + 1, CalcRows, CDbl(Settings(1)), _
                                                      CDbl(Settings(2)), CDbl(Settings(3)), Counts)
            Summary = CORE_Reconcile.SummaryText(Counts, CDbl(Settings(1)), CDbl(Settings(2)), CDbl(Settings(3)))
        Else
            ComparedRows = CORE_Reconcile.CompareTrades(BankRows, HeaderRow + 1, CalcRows, CDbl(Settings(1)), _
                                                 CDbl(Settings(2)), CDbl(Settings(3)), Counts)
            Summary = CORE_Reconcile.TradeSummaryText(Counts, CDbl(Settings(1)), CDbl(Settings(2)), _
                                                      CDbl(Settings(3)))
        End If
        ReconFirstRow = OutputFirstRow(ReconSheetOf(ReconKind))
        ClearOutputBlock ReconSheet, ReconFirstRow, ReconColsOf(ReconKind)
        If IsArray(ComparedRows) Then
            If ReconKind = RECON_NS Then
                ReconRows = ComparedRows
            Else
                ReconRows = WithEconomicTradeIds(ComparedRows)
            End If
            RowsToWrite = UBound(ReconRows, 1)
            WriteBlock ReconSheet, ReconFirstRow, ReconRows, RowsToWrite, ReconColsOf(ReconKind)
            FormatColumns ReconSheet, ReconFirstRow, RowsToWrite, "@|0|@|@|@|@|#,##0.00########|#,##0.00########|" & _
                          "#,##0.00########|0.0000%|@|@" & IIf(ReconKind = RECON_NS, "", "|@")
        End If
        ReconSheet.Range(RECON_INFO_CELL).Value = "Last reconciliation " & Format$(Now, "yyyy-mm-dd hh:mm:ss") & _
                                               " | " & Summary
        RememberReconInputs ReconKind, ReconFingerprint(ReconKind)
        OtherCount = Counts(RCN_MISSING) + Counts(RCN_NOTINBANK) + Counts(RCN_DUPLICATE) + Counts(RCN_REGIME) + _
                 Counts(RCN_WITHHELD) + Counts(RCN_INVALID) + Counts(RCN_EXCLUDED) + Counts(RCN_NA)
        AddReconLine ReconName & ": " & Counts(RCN_COMPARED) & " figure(s) compared, " & Counts(RCN_MATCH) & _
                     " MATCH, " & Counts(RCN_DIFF) & " DIFFERENCE" & _
                     IIf(OtherCount > 0, "; " & OtherCount & " other row(s): see the " & ReconSheetOf(ReconKind) & " sheet", "") & "."
        ReconcileAfterRun = "DONE"
        Exit Function

'------------------------------------------------------------------------------
' HANDLE ERROR
'------------------------------------------------------------------------------
Failed:
    'ClearReconciliation clears Err: keep the cause for the run message
    '(#318).
        ErrorNumber = Err.Number
        ErrorText = Err.Description
        ClearReconciliation "Reconciliation failed " & Format$(Now, "yyyy-mm-dd hh:mm:ss") & " (error " & _
                            ErrorNumber & ": " & ErrorText & ").", ReconKind
        AddReconLine ReconName & " failed: " & ErrorText & "."
        ReconcileAfterRun = "FAILED"

End Function


Private Sub AddReconLine( _
    ByVal LineText As String)
'
'==============================================================================
'                                 AddReconLine
'------------------------------------------------------------------------------
' PURPOSE
'   Add a line about a reconciliation to mReconLine, the text of the run
'   message.
'
' UPDATED
'   2026-10-17
'==============================================================================
'
        If Len(mReconLine) > 0 Then
            mReconLine = mReconLine & vbCrLf
        End If
        mReconLine = mReconLine & LineText

End Sub


Private Sub ClearReconciliation( _
    ByVal InfoText As String, _
    Optional ByVal ReconKind As Long = 0)
'
'==============================================================================
'                             ClearReconciliation
'------------------------------------------------------------------------------
' PURPOSE
'   Withdraw a reconciliation: empty its fingerprint and its rows and write
'   InfoText to its summary cell. Every operation starts with this for
'   both, so that no reconciliation outlives the results it compared.
'
' INPUTS
'   InfoText: the text for the summary cell.
'   ReconKind: RECON_NS or RECON_TR; 0 for both.
'
' ERROR POLICY
'   Best effort, and never raises: the fingerprint is emptied first, so a
'   sheet that cannot be cleared is still reported as not current.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim k            As Long         'Reconciliation cleared
    Dim ReconSheet   As Worksheet    'Its sheet

'------------------------------------------------------------------------------
' CLEAR
'------------------------------------------------------------------------------
        On Error Resume Next
        For k = RECON_NS To RECON_TR
            If ReconKind = 0 Or ReconKind = k Then
                ForgetReconInputs k
                If ReconciliationSheetsPresent(k) Then
                    Set ReconSheet = GetSheet(ReconSheetOf(k))
                    ClearOutputBlock ReconSheet, OutputFirstRow(ReconSheetOf(k)), ReconColsOf(k)
                    ReconSheet.Range(RECON_INFO_CELL).Value = InfoText
                End If
            End If
        Next k
        Err.Clear

End Sub


Private Function ReconFingerprint( _
    ByVal ReconKind As Long) _
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
'   ReconKind: RECON_NS or RECON_TR.
'
' ERROR POLICY
'   Raises when the bank table or a setting cannot be located, or as
'   CORE_Engine.InputFingerprint.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Parts(1 To 5)   As String    'Text of each part
    Dim HeaderRow       As Long      'Header row of the bank table
    Dim i               As Long      'Setting
    Dim SettingCell     As Range     'A setting cell

'------------------------------------------------------------------------------
' FINGERPRINT
'------------------------------------------------------------------------------
        HeaderRow = BankHeaderRow(ReconKind)
        If HeaderRow = 0 Then
            Err.Raise ERR_PARAMS_LAYOUT, "M_Main.ReconFingerprint", _
                      "The name " & HeaderNameOf(ReconKind) & " is missing or not in column C of " & BankSheetOf(ReconKind) & "."
        End If
        Parts(1) = CORE_Engine.InputFingerprint()
        Parts(2) = CORE_Engine.BlockText(GetSheet(BankSheetOf(ReconKind)), HeaderRow, BankColsOf(ReconKind) - 1)
        For i = 1 To 3
            Set SettingCell = NamedCell(SettingNameOf(ReconKind, i))
            If SettingCell Is Nothing Then
                Err.Raise ERR_PARAMS_LAYOUT, "M_Main.ReconFingerprint", _
                          "The name " & SettingNameOf(ReconKind, i) & " is missing."
            End If
            Parts(2 + i) = CORE_Engine.BlockText(SettingCell.Worksheet, SettingCell.Row, 1, SettingCell.Column - COL_OFFSET, SettingCell.Row)
        Next i
        ReconFingerprint = TextHash(Join(Parts, Chr$(29)))

End Function


Private Sub RememberReconInputs( _
    ByVal ReconKind As Long, _
    ByVal Fingerprint As String)
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
'   2026-10-17
'==============================================================================
'
        On Error GoTo Failed
        ThisWorkbook.Names.Add Name:=InputsNameOf(ReconKind), RefersTo:="=""" & Fingerprint & """", Visible:=False
        Exit Sub

Failed:
        ForgetReconInputs ReconKind

End Sub


Private Sub ForgetReconInputs( _
    ByVal ReconKind As Long)
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
'   2026-10-17
'==============================================================================
'
        On Error GoTo Failed
        ThisWorkbook.Names.Add Name:=InputsNameOf(ReconKind), RefersTo:="=""""", Visible:=False

Failed:

End Sub


Private Function ReconciliationSheetsPresent( _
    ByVal ReconKind As Long) _
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
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim BankSheet    As Worksheet    'The bank sheet; Nothing when absent
    Dim ReconSheet   As Worksheet    'The reconciliation sheet; Nothing when absent

'------------------------------------------------------------------------------
' LOOK UP
'------------------------------------------------------------------------------
        On Error Resume Next
        Set BankSheet = ThisWorkbook.Worksheets(BankSheetOf(ReconKind))
        Set ReconSheet = ThisWorkbook.Worksheets(ReconSheetOf(ReconKind))
        Err.Clear
        On Error GoTo 0
        ReconciliationSheetsPresent = Not BankSheet Is Nothing And Not ReconSheet Is Nothing

End Function


Private Function BankHeaderRow( _
    ByVal ReconKind As Long) _
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
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim HeaderCell   As Range    'The named cell

'------------------------------------------------------------------------------
' RESOLVE
'------------------------------------------------------------------------------
        Set HeaderCell = NamedCell(HeaderNameOf(ReconKind))
        If Not HeaderCell Is Nothing Then
            If StrComp(HeaderCell.Worksheet.Name, BankSheetOf(ReconKind), vbTextCompare) = 0 And HeaderCell.Column = COL_OFFSET + 1 Then
                BankHeaderRow = HeaderCell.Row
            End If
        End If

End Function


Private Function NamedValue( _
    ByVal NameText As String, _
    ByVal SheetName As String, _
    ByRef ValueRead As Variant) _
    As String
'
'==============================================================================
'                                  NamedValue
'------------------------------------------------------------------------------
' PURPOSE
'   Read the value of a workbook name's cell on a given sheet.
'
' INPUTS
'   NameText: the workbook name; SheetName: the sheet it must be on.
'   ValueRead: receives the cell's Value2.
'
' RETURNS
'   "" when it was read; otherwise why not.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ValueCell   As Range    'The named cell

'------------------------------------------------------------------------------
' READ
'------------------------------------------------------------------------------
        Set ValueCell = NamedCell(NameText)
        If ValueCell Is Nothing Then
            NamedValue = "the name " & NameText & " is missing"
        ElseIf StrComp(ValueCell.Worksheet.Name, SheetName, vbTextCompare) <> 0 Then
            NamedValue = "the name " & NameText & " is not on " & SheetName
        Else
            ValueRead = ValueCell.Value2
        End If

End Function


Private Function BankSheetOf( _
    ByVal ReconKind As Long) _
    As String
'
'==============================================================================
' PURPOSE
'   The bank sheet of a reconciliation.
'==============================================================================
'
        BankSheetOf = IIf(ReconKind = RECON_NS, SH_BANK_NS, SH_BANK_TR)

End Function


Private Function ReconSheetOf( _
    ByVal ReconKind As Long) _
    As String
'
'==============================================================================
' PURPOSE
'   The output sheet of a reconciliation.
'==============================================================================
'
        ReconSheetOf = IIf(ReconKind = RECON_NS, SH_RECON, SH_TRADE_RECON)

End Function


Private Function HeaderNameOf( _
    ByVal ReconKind As Long) _
    As String
'
'==============================================================================
' PURPOSE
'   The workbook name of the first header cell of a bank table.
'==============================================================================
'
        HeaderNameOf = IIf(ReconKind = RECON_NS, NM_BN_HEADER, NM_BT_HEADER)

End Function


Private Function SettingNameOf( _
    ByVal ReconKind As Long, _
    ByVal SettingIndex As Long) _
    As String
'
'==============================================================================
' PURPOSE
'   The workbook name of setting SettingIndex of a bank sheet: 1 the scale,
'   2 the absolute tolerance, 3 the relative tolerance.
'==============================================================================
'
        If ReconKind = RECON_NS Then
            SettingNameOf = Choose(SettingIndex, NM_BN_SCALE, NM_BN_ABSTOL, NM_BN_RELTOL)
        Else
            SettingNameOf = Choose(SettingIndex, NM_BT_SCALE, NM_BT_ABSTOL, NM_BT_RELTOL)
        End If

End Function


Private Function BankColsOf( _
    ByVal ReconKind As Long) _
    As Long
'
'==============================================================================
' PURPOSE
'   The number of columns of a bank table; the last is its comment.
'==============================================================================
'
        BankColsOf = IIf(ReconKind = RECON_NS, BN_NCOLS, BT_NCOLS)

End Function


Private Function ReconColsOf( _
    ByVal ReconKind As Long) _
    As Long
'
'==============================================================================
' PURPOSE
'   The number of columns of a reconciliation sheet.
'==============================================================================
'
        ReconColsOf = IIf(ReconKind = RECON_NS, RC_NCOLS, TRC_NCOLS)

End Function


Private Function WithEconomicTradeIds( _
    ByVal ComparedRows As Variant) _
    As Variant
'
'==============================================================================
'                             WithEconomicTradeIds
'------------------------------------------------------------------------------
' PURPOSE
'   Add to the rows of a trade reconciliation the Economic trade ID of each
'   leg as on Trades (#45), matched by Trade ID as the comparison matches
'   it; blank for a whole trade and for a leg that is not on Trades.
'
' INPUTS
'   ComparedRows: the rows from CORE_Reconcile.CompareTrades, RC_NCOLS
'   wide.
'
' RETURNS
'   The same rows, TRC_NCOLS wide.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim TradesSheet     As Worksheet     'Trades
    Dim FirstRow        As Long          'First data row of Trades
    Dim LastRow         As Long          'Last data row of Trades
    Dim LegIds          As Variant       'Trade IDs on Trades
    Dim EconomicIds     As Variant       'Economic trade IDs
    Dim LegIndex        As Collection    'Trade ID to row on Trades
    Dim WidenedRows()   As Variant       'The widened rows
    Dim r               As Long          'Row
    Dim c               As Long          'Column
    Dim TradesRow       As Long          'Row on Trades

'------------------------------------------------------------------------------
' INDEX TRADES
'------------------------------------------------------------------------------
        Set TradesSheet = GetSheet(SH_TRADES)
        Set LegIndex = New Collection
        FirstRow = TableFirstRow(SH_TRADES)
        LastRow = LastDataRow(TradesSheet, FirstRow, TR_ID)
        If LastRow >= FirstRow Then
            LegIds = ColumnValues(TradesSheet, FirstRow, LastRow, TR_ID)
            EconomicIds = ColumnValues(TradesSheet, FirstRow, LastRow, TR_ECON)
            For r = 1 To UBound(LegIds, 1)
                If Len(UTxt(LegIds(r, 1))) > 0 Then
                    If KeyIndex(LegIndex, UTxt(LegIds(r, 1))) = 0 Then
                        KeyAdd LegIndex, UTxt(LegIds(r, 1)), r
                    End If
                End If
            Next r
        End If

'------------------------------------------------------------------------------
' WIDEN
'------------------------------------------------------------------------------
        ReDim WidenedRows(1 To UBound(ComparedRows, 1), 1 To TRC_NCOLS)
        For r = 1 To UBound(ComparedRows, 1)
            For c = 1 To RC_NCOLS
                WidenedRows(r, c) = ComparedRows(r, c)
            Next c
            TradesRow = 0
            If Len(UTxt(ComparedRows(r, RC_ID))) > 0 Then
                TradesRow = KeyIndex(LegIndex, UTxt(ComparedRows(r, RC_ID)))
            End If
            If TradesRow > 0 Then
                If Not IsError(EconomicIds(TradesRow, 1)) Then
                    If Len(Trim$(CStr(EconomicIds(TradesRow, 1)))) > 0 Then
                        WidenedRows(r, TRC_ECON) = Trim$(CStr(EconomicIds(TradesRow, 1)))
                    End If
                End If
            End If
        Next r
        WithEconomicTradeIds = WidenedRows

End Function


Private Function BankHeadersOf( _
    ByVal ReconKind As Long) _
    As String
'
'==============================================================================
' PURPOSE
'   The expected headers of a bank table.
'==============================================================================
'
        BankHeadersOf = IIf(ReconKind = RECON_NS, BN_HEADERS, BT_HEADERS)

End Function


Private Function InputsNameOf( _
    ByVal ReconKind As Long) _
    As String
'
'==============================================================================
' PURPOSE
'   The hidden workbook name holding a reconciliation's fingerprint.
'==============================================================================
'
        InputsNameOf = IIf(ReconKind = RECON_NS, RECON_INPUTS_NAME, TRADE_RECON_INPUTS_NAME)

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
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Problem      As String     'Why the report cannot be written; "" when it can
    Dim ChosenFile   As Variant    'Chosen file, or False when cancelled
    Dim LineCount    As Long       'Lines written

'------------------------------------------------------------------------------
' CHECK, ASK AND WRITE
'------------------------------------------------------------------------------
        On Error GoTo Failed
        Problem = EvidenceProblem()
        If Len(Problem) > 0 Then
            Notify "Evidence not exported: " & Problem & ".", vbExclamation
            Exit Sub
        End If
        ChosenFile = Application.GetSaveAsFilename(InitialFileName:=DefaultEvidenceName(), _
                                               FileFilter:="CSV files (*.csv), *.csv", _
                                               Title:="Export evidence")
        If VarType(ChosenFile) = vbBoolean Then
            Exit Sub
        End If
        LineCount = WriteEvidence(CStr(ChosenFile))
        Notify "Evidence exported: " & CStr(ChosenFile) & " (" & LineCount & " lines).", vbInformation
        Exit Sub

Failed:
        Notify "Evidence not exported: " & Err.Description, vbCritical

End Sub


Public Function ExportEvidence_Silent( _
    ByVal FilePath As String) _
    As String
'
'==============================================================================
'                            ExportEvidence_Silent
'------------------------------------------------------------------------------
' PURPOSE
'   ExportEvidence for automation and tests: write the report to FilePath
'   without asking.
'
' RETURNS
'   "RESULT=OK; lines=N; file=..." or "RESULT=REFUSED; reason=...".
'
' ERROR POLICY
'   A file that cannot be written raises to the caller.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Problem   As String    'Why the report cannot be written; "" when it can

'------------------------------------------------------------------------------
' CHECK AND WRITE
'------------------------------------------------------------------------------
        Problem = EvidenceProblem()
        If Len(Problem) > 0 Then
            ExportEvidence_Silent = "RESULT=REFUSED; reason=" & Problem
        Else
            ExportEvidence_Silent = "RESULT=OK; lines=" & WriteEvidence(FilePath) & "; file=" & FilePath
        End If

End Function


Private Function EvidenceProblem() As String
'
'==============================================================================
'                               EvidenceProblem
'------------------------------------------------------------------------------
' PURPOSE
'   Tell why no evidence report may be written now: the results are not
'   those of the current inputs, an output cell was changed after the run,
'   or a sheet of bank figures holds figures whose reconciliation is
'   missing or out of date. An evidence report must never describe
'   results that do not match the inputs.
'
' RETURNS
'   "" when the report may be written; otherwise the reason.
'
' ERROR POLICY
'   Never raises: an error is itself the reason.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim kind                As Long       'Reconciliation checked
    Dim StoredFingerprint   As String     'A stored fingerprint
    Dim IsCurrent           As Boolean    'True when the reconciliation is current
    Dim ResultState         As String     'ResultsStatus()

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
        ResultState = ResultsStatus()
        If ResultState = "EDITED" Then
            EvidenceProblem = "an output cell on Results, Checks or a reconciliation sheet was changed after " & _
                              "the run; press Run SA-CCR first"
            Exit Function
        ElseIf ResultState <> "CURRENT" Then
            EvidenceProblem = "the results are not those of the current inputs; press Run SA-CCR first"
            Exit Function
        End If
    'A stored reconciliation must match its bank figures as they are now,
    'also when every bank row was removed after the run; bank figures
    'without a stored reconciliation were never compared.
        For kind = RECON_NS To RECON_TR
            If ReconciliationSheetsPresent(kind) Then
                StoredFingerprint = StoredText(InputsNameOf(kind))
                If Len(StoredFingerprint) > 0 Then
                    IsCurrent = (StoredFingerprint = ReconFingerprint(kind))
                Else
                    IsCurrent = Not BankRowsEntered(kind)
                End If
                If Not IsCurrent Then
                    EvidenceProblem = "the reconciliation of the bank figures on " & BankSheetOf(kind) & _
                                      " is not current; press Run SA-CCR first"
                    Exit Function
                End If
            End If
        Next kind
    'The output sheets are not protected: ResultsStatus has compared them
    'with what the run wrote, which a run that recorded nothing cannot be.
        StoredFingerprint = StoredText(RUN_OUTPUTS_NAME)
        If Len(StoredFingerprint) = 0 Then
            EvidenceProblem = "the outputs of the last run were not recorded; press Run SA-CCR first"
        End If
        Exit Function

Failed:
        EvidenceProblem = "the workbook could not be checked (" & Err.Description & ")"

End Function


Private Function WriteEvidence( _
    ByVal FilePath As String) _
    As Long
'
'==============================================================================
'                                WriteEvidence
'------------------------------------------------------------------------------
' PURPOSE
'   Write the evidence report to FilePath as UTF-8 without a byte-order
'   mark, replacing any file of that name, hidden ones included; a
'   read-only file is left as it is and the write fails.
'
' METHOD
'   The report is written whole to a new file beside FilePath first. Only
'   then is an older file moved aside, the new one renamed to FilePath and
'   the older one deleted, so a failed write leaves the older file as it
'   was (#312). A failed replacement puts the older file back with its
'   attributes. The replacement is as atomic as the file system's rename,
'   no more.
'
' RETURNS
'   The number of lines written.
'
' ERROR POLICY
'   Raises the first error when the report cannot be written or replaced,
'   after closing the file number, deleting the new file and putting an
'   older file back; an error in that cleanup does not replace the first
'   one. An older file that cannot be deleted once replaced is left beside
'   the report.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ReportText          As String     'The report
    Dim ReportBytes()       As Byte       'Its UTF-8 bytes
    Dim FileNumber          As Integer    'File number; 0 while none is open
    Dim NewPath             As String     'The report before it replaces FilePath; "" once renamed
    Dim OldPath             As String     'An older FilePath while it is moved aside; "" otherwise
    Dim OldAttributes       As Long       'Its attributes, put back if the replacement fails
    Dim AttributesCleared   As Boolean    'Its attributes were cleared
    Dim ErrorNumber         As Long       'First error; 0 on success
    Dim ErrorText           As String     'Its description

'------------------------------------------------------------------------------
' WRITE THE NEW FILE
'------------------------------------------------------------------------------
        ReportText = EvidenceText()
        ReportBytes = CORE_Export.Utf8Bytes(ReportText)
        On Error GoTo Failed
        If FileExists(FilePath) Then
            If (GetAttr(FilePath) And vbReadOnly) <> 0 Then
                Err.Raise ERR_EXPORT_WRITE, "M_Main.WriteEvidence", "the file is read-only"
            End If
        End If
        NewPath = SiblingPath(FilePath, "new")
        If gTestFault = "evidence-open" Then
            Err.Raise ERR_INJECTED_FAULT, "M_Main.WriteEvidence", "Injected failure before the report was opened."
        End If
        FileNumber = FreeFile
        Open NewPath For Binary Access Write As #FileNumber
        Put #FileNumber, , ReportBytes
        If gTestFault = "evidence-write" Then
            Err.Raise ERR_INJECTED_FAULT, "M_Main.WriteEvidence", "Injected failure while the report was written."
        End If
        If LOF(FileNumber) <> UBound(ReportBytes) - LBound(ReportBytes) + 1 Then
            Err.Raise ERR_EXPORT_WRITE, "M_Main.WriteEvidence", "the report was not written whole"
        End If
        Close #FileNumber
        FileNumber = 0

'------------------------------------------------------------------------------
' REPLACE THE OLDER FILE
'------------------------------------------------------------------------------
        If FileExists(FilePath) Then
            OldPath = SiblingPath(FilePath, "old")
            OldAttributes = GetAttr(FilePath) And (vbReadOnly Or vbHidden Or vbSystem Or vbArchive)
            SetAttr FilePath, vbNormal
            AttributesCleared = True
            Name FilePath As OldPath
        End If
        If gTestFault = "evidence-replace" Then
            Err.Raise ERR_INJECTED_FAULT, "M_Main.WriteEvidence", "Injected failure while the older file was replaced."
        End If
        Name NewPath As FilePath
        NewPath = ""
        WriteEvidence = (Len(ReportText) - Len(Replace(ReportText, vbCrLf, ""))) \ 2

'------------------------------------------------------------------------------
' CLEAN UP AND REPORT
'------------------------------------------------------------------------------
    'Reached on success and, through Failed, after an error. Errors are
    'ignored here (On Error Resume Next), so that a cleanup error cannot
    'replace the first.
CleanUp:
        On Error Resume Next
        If FileNumber <> 0 Then
            Close #FileNumber
        End If
        If Len(NewPath) > 0 Then
            If FileExists(NewPath) Then
                Kill NewPath
            End If
        End If
        If Len(OldPath) > 0 Then
            If ErrorNumber = 0 Then
                Kill OldPath
            ElseIf Not FileExists(FilePath) Then
                Name OldPath As FilePath
            End If
        End If
        If AttributesCleared And ErrorNumber <> 0 Then
            SetAttr FilePath, OldAttributes
        End If
        On Error GoTo 0
        If ErrorNumber <> 0 Then
            Err.Raise ErrorNumber, "M_Main.WriteEvidence", "Could not write " & FilePath & ": " & ErrorText
        End If
        Exit Function

'------------------------------------------------------------------------------
' HANDLE ERROR
'------------------------------------------------------------------------------
Failed:
        ErrorNumber = Err.Number
        ErrorText = Err.Description
        Resume CleanUp

End Function


Private Function FileExists( _
    ByVal FilePath As String) _
    As Boolean
'
'==============================================================================
'                                  FileExists
'------------------------------------------------------------------------------
' PURPOSE
'   Tell whether a file of that name exists, hidden and system files
'   included.
'
' UPDATED
'   2026-10-17
'==============================================================================
'
        FileExists = Len(Dir$(FilePath, vbHidden Or vbSystem)) > 0

End Function


Private Function SiblingPath( _
    ByVal FilePath As String, _
    ByVal Tag As String) _
    As String
'
'==============================================================================
'                                 SiblingPath
'------------------------------------------------------------------------------
' PURPOSE
'   A name for a file beside FilePath that does not exist yet, such as
'   "report.csv.saccr-new1": in the same folder, so that a rename does not
'   move the file to another drive.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim n           As Long      'Counter that makes the name unique
    Dim Candidate   As String    'Name tried

'------------------------------------------------------------------------------
' FIND A FREE NAME
'------------------------------------------------------------------------------
        Do
            n = n + 1
            Candidate = FilePath & ".saccr-" & Tag & n
        Loop While FileExists(Candidate)
        SiblingPath = Candidate

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
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Report                   As String       'The report
    Dim ResultsSheet             As Worksheet    'Results
    Dim LastRow                  As Long         'Last row of a table
    Dim FirstDataRow             As Long         'First data row of a table
    Dim r                        As Long         'Row
    Dim ResultId                 As String       'Results ID
    Dim ParamCodes               As Variant      'Parameter codes
    Dim i                        As Long         'Code or reconciliation
    Dim SeverityCounts(1 To 3)   As Long         'Checks lines: ERROR, WARNING, INFO
    Dim Severity                 As String       'Severity of a Checks line
    Dim KeyPrefix                As String       'Key prefix of a reconciliation
    Dim ReportingDateCell        As Range        'A named cell

'------------------------------------------------------------------------------
' RUN
'------------------------------------------------------------------------------
        Set ResultsSheet = GetSheet(SH_RESULTS)
        Report = RunLine("format", "K-SACCR evidence 1")
        Report = Report & RunLine("run_summary", ResultsSheet.Range(RUNINFO_CELL).Value2)
        Report = Report & RunLine("results_status", ResultsStatus())
        Report = Report & RunLine("inputs", CORE_Engine.LastRunInputs())
        Report = Report & RunLine("outputs", StoredText(RUN_OUTPUTS_NAME))
        Set ReportingDateCell = NamedCell(PRM_ASOF)
        If IsNumeric(ReportingDateCell.Value2) And Not IsEmpty(ReportingDateCell.Value2) Then
            Report = Report & RunLine("reporting_date", Format$(CDate(ReportingDateCell.Value2), "yyyy-mm-dd"))
        Else
            Report = Report & RunLine("reporting_date", ReportingDateCell.Value2)
        End If
        Report = Report & RunLine("reporting_ccy", NamedCell(PRM_REPCCY).Value2)
        Report = Report & RunLine("regime_default", NamedCell(PRM_REGIME).Value2)
    'The set the template ships, not a claim about this run: whether the run
    'used it unchanged is in run_summary (regulatory or what-if).
        Report = Report & RunLine("reference_parameter_set", REF_SET_ID)
        FirstDataRow = OutputFirstRow(SH_RESULTS)
        LastRow = LastDataRow(ResultsSheet, FirstDataRow, RS_ID)
        For r = FirstDataRow To LastRow
            ResultId = UTxt(TableCell(ResultsSheet, r, RS_ID).Value2)
            If ResultId = "TOTAL " & RG_CRR Or ResultId = "TOTAL " & RG_BCBS Then
                Report = Report & RunLine("total_ead_" & LCase$(Mid$(ResultId, 7)), TableCell(ResultsSheet, r, RS_EAD).Value2)
                Report = Report & RunLine("total_status_" & LCase$(Mid$(ResultId, 7)), TableCell(ResultsSheet, r, RS_STATUS_COL).Value2)
            End If
        Next r
        FirstDataRow = OutputFirstRow(SH_CHECKS)
        LastRow = LastDataRow(GetSheet(SH_CHECKS), FirstDataRow, 1)
        For r = FirstDataRow To LastRow
            Severity = UTxt(TableCell(GetSheet(SH_CHECKS), r, 1).Value2)
            If Severity = SEV_ERROR Then SeverityCounts(1) = SeverityCounts(1) + 1
            If Severity = SEV_WARN Then SeverityCounts(2) = SeverityCounts(2) + 1
            If Severity = SEV_INFO Then SeverityCounts(3) = SeverityCounts(3) + 1
        Next r
        Report = Report & RunLine("checks_errors", SeverityCounts(1))
        Report = Report & RunLine("checks_warnings", SeverityCounts(2))
        Report = Report & RunLine("checks_info", SeverityCounts(3))
        For i = RECON_NS To RECON_TR
            KeyPrefix = IIf(i = RECON_NS, "reconciliation", "trade_reconciliation")
            If Not ReconciliationSheetsPresent(i) Then
                Report = Report & RunLine(KeyPrefix, "UNAVAILABLE")
            Else
            'Without a stored fingerprint the reconciliation was not made:
            'no bank figures, or settings or headers it could not use; the
            'summary says which.
                Report = Report & RunLine(KeyPrefix, IIf(Len(StoredText(InputsNameOf(i))) > 0, "CURRENT", "NOT MADE"))
                Report = Report & RunLine(KeyPrefix & "_inputs", StoredText(InputsNameOf(i)))
                Report = Report & RunLine(KeyPrefix & "_summary", GetSheet(ReconSheetOf(i)).Range(RECON_INFO_CELL).Value2)
            'The settings are part of a reconciliation's fingerprint; without
            'one they were not used, and could change after the run.
                If Len(StoredText(InputsNameOf(i))) > 0 Then
                    Report = Report & RunLine(KeyPrefix & "_scale", NamedCell(SettingNameOf(i, 1)).Value2)
                    Report = Report & RunLine(KeyPrefix & "_abs_tol", NamedCell(SettingNameOf(i, 2)).Value2)
                    Report = Report & RunLine(KeyPrefix & "_rel_tol", NamedCell(SettingNameOf(i, 3)).Value2)
                End If
            End If
        Next i

'------------------------------------------------------------------------------
' PARAMETER SNAPSHOT
'------------------------------------------------------------------------------
        ParamCodes = CORE_Engine.ParameterCodes()
        For i = LBound(ParamCodes) To UBound(ParamCodes)
            Report = Report & CORE_Export.CsvLine("PARAMS", Array(ParamCodes(i), NamedCell(CStr(ParamCodes(i))).Value2))
        Next i
        Report = Report & ParamsTableLines("SF_TABLE", NM_SF_TABLE, SF_HEADERS)
        Report = Report & ParamsTableLines("FX_TABLE", NM_FX_TABLE, FX_HEADERS)

'------------------------------------------------------------------------------
' TABLES
'------------------------------------------------------------------------------
        Report = Report & OutputTableLines()
        Report = Report & CORE_Export.LimitationLines()
        EvidenceText = Report

End Function


Private Function ParamsTableLines( _
    ByVal Section As String, _
    ByVal NameText As String, _
    ByVal Headers As String) _
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
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ParamsSheet   As Worksheet    'Params
    Dim HeaderRow     As Long         'Header row
    Dim KeyColumn     As Long         'Key column
    Dim r             As Long         'Row read

'------------------------------------------------------------------------------
' READ
'------------------------------------------------------------------------------
        Set ParamsSheet = GetSheet(SH_PARAMS)
        HeaderRow = NamedRow(NameText)
        KeyColumn = NamedTableCol(NameText)
        If HeaderRow = 0 Then
            Err.Raise ERR_PARAMS_LAYOUT, "M_Main.ParamsTableLines", "The name " & NameText & " is missing."
        End If
        r = HeaderRow + 1
        Do While Not IsBlankCell(TableCell(ParamsSheet, r, KeyColumn).Value)
            r = r + 1
        Loop
        ParamsTableLines = CORE_Export.TableLines(Section, TableRange(ParamsSheet, HeaderRow, KeyColumn, r - 1, _
                                                  KeyColumn + UBound(Split(Headers, "|"))).Value2)

End Function


Private Function SheetTableLines( _
    ByVal Section As String, _
    ByVal SheetName As String, _
    ByVal ColumnCount As Long) _
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
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim TableSheet   As Worksheet    'The sheet
    Dim HeaderRow    As Long         'Header row of the table
    Dim LastRow      As Long         'Last row of the table

'------------------------------------------------------------------------------
' WRITE
'------------------------------------------------------------------------------
        Set TableSheet = GetSheet(SheetName)
        HeaderRow = OutputFirstRow(SheetName) - 1
        LastRow = LastDataRow(TableSheet, HeaderRow + 1, 1)
        If LastRow < HeaderRow Then
            LastRow = HeaderRow
        End If
        SheetTableLines = CORE_Export.TableLines(Section, TableRange(TableSheet, HeaderRow, 1, LastRow, ColumnCount).Value2)

End Function


Private Function RunLine( _
    ByVal EntryKey As String, _
    ByVal EntryValue As Variant) _
    As String
'
'==============================================================================
' PURPOSE
'   One line of the RUN section: key and value.
'==============================================================================
'
        RunLine = CORE_Export.CsvLine("RUN", Array(EntryKey, EntryValue))

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
    ByVal ReconKind As Long) _
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
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim BankSheet   As Worksheet    'The bank sheet
    Dim HeaderRow   As Long         'Its header row
    Dim LastRow     As Long         'Its last used row

'------------------------------------------------------------------------------
' READ
'------------------------------------------------------------------------------
        HeaderRow = BankHeaderRow(ReconKind)
        If HeaderRow = 0 Then
            BankRowsEntered = True
            Exit Function
        End If
        Set BankSheet = GetSheet(BankSheetOf(ReconKind))
        LastRow = UsedLastRow(BankSheet)
        If LastRow > HeaderRow Then
            BankRowsEntered = CORE_Reconcile.HasBankRows( _
                TableRange(BankSheet, HeaderRow + 1, 1, LastRow, BankColsOf(ReconKind)).Value2, BankColsOf(ReconKind))
        End If

End Function


Private Function StoredText( _
    ByVal NameText As String) _
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
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim StoredName   As Name    'The stored name, Nothing when absent

'------------------------------------------------------------------------------
' READ
'------------------------------------------------------------------------------
        On Error Resume Next
        Set StoredName = ThisWorkbook.Names(NameText)
        Err.Clear
        On Error GoTo 0
        If Not StoredName Is Nothing Then
            StoredText = Replace(Replace(StoredName.RefersTo, "=", ""), """", "")
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
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Lines   As String    'The lines

'------------------------------------------------------------------------------
' READ
'------------------------------------------------------------------------------
        Lines = SheetTableLines("RESULTS", SH_RESULTS, RS_NCOLS)
        If ReconciliationSheetsPresent(RECON_NS) Then
            Lines = Lines & SheetTableLines("RECONCILIATION", SH_RECON, RC_NCOLS)
        End If
        If ReconciliationSheetsPresent(RECON_TR) Then
            Lines = Lines & SheetTableLines("TRADE_RECONCILIATION", SH_TRADE_RECON, TRC_NCOLS)
        End If
        OutputTableLines = Lines & SheetTableLines("CHECKS", SH_CHECKS, CK_NCOLS)

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
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim FingerprintText   As String    'The text fingerprinted
    Dim kind              As Long      'Reconciliation

'------------------------------------------------------------------------------
' FINGERPRINT
'------------------------------------------------------------------------------
        FingerprintText = CORE_Export.CsvField(GetSheet(SH_RESULTS).Range(RUNINFO_CELL).Value2)
        For kind = RECON_NS To RECON_TR
            If ReconciliationSheetsPresent(kind) Then
                FingerprintText = FingerprintText & Chr$(29) & CORE_Export.CsvField(GetSheet(ReconSheetOf(kind)).Range(RECON_INFO_CELL).Value2)
            End If
        Next kind
        OutputFingerprint = TextHash(FingerprintText & Chr$(29) & OutputTableLines())

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
    ByVal Operation As Long) _
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
'   Operation: OP_RUN, OP_VALIDATE or OP_CLEAR.
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
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim SavedCalculation   As XlCalculation    'Calculation mode before the operation
    Dim SavedEvents        As Boolean          'EnableEvents before the operation
    Dim SavedScreen        As Boolean          'ScreenUpdating before the operation
    Dim StateCaptured      As Boolean          'True once all three were read
    Dim Completed          As Boolean          'True when the engine completed the run
    Dim ErrorNumber        As Long             'Primary error number; 0 for none
    Dim ErrorSource        As String           'Primary error source
    Dim ErrorDescription   As String           'Primary error description
    Dim CleanupDetails     As String           'Restoration failures; empty when all restored

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
        SavedCalculation = Application.Calculation
        SavedEvents = Application.EnableEvents
        SavedScreen = Application.ScreenUpdating
        StateCaptured = True
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
        Select Case Operation
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
        If Operation <> OP_CLEAR Then
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
        Select Case Operation
            Case OP_RUN
                Completed = CORE_Engine.Calculate(True)
                If Completed Then
                    mReconState = ReconcileAfterRun(RECON_NS)
                    mTradeReconState = ReconcileAfterRun(RECON_TR)
                    RememberRunOutputs
                Else
                    ClearReconciliation "Reconciliation not run: the SA-CCR run stopped; see the Checks sheet."
                End If
            Case OP_VALIDATE
                Completed = CORE_Engine.Calculate(False)
            Case OP_CLEAR
                ClearAllOutputs
                Completed = True
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
        If StateCaptured Then
            RefreshResultsStatus
            CleanupDetails = RestoreSettings(SavedCalculation, SavedEvents, SavedScreen)
        End If
        mRunning = False
        If ErrorNumber <> 0 Then
            ReportFailure ErrorNumber, ErrorSource, ErrorDescription, CleanupDetails
        Else
            ExecuteOperation = ReportOutcome(Operation, Completed, CleanupDetails)
        End If
        Exit Function

'------------------------------------------------------------------------------
' HANDLE ERROR
'------------------------------------------------------------------------------
Failed:
        ErrorNumber = Err.Number
        ErrorSource = Err.Source
        ErrorDescription = Err.Description
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
'   fingerprint. In manual calculation Excel would otherwise keep the value
'   computed during the recalculation before the run, when the fingerprint
'   had already been withdrawn (#204), and show "No results" after a
'   completed run.
'
' ERROR POLICY
'   Best effort: a status that cannot be recalculated here is refreshed by
'   the next recalculation, and the outcome of the operation is unchanged.
'
' UPDATED
'   2026-10-17
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
'   Clear the five output tables, withdraw the run's input fingerprint and
'   write the time to the run-summary cell. ExecuteOperation withdraws the
'   reconciliations and the output fingerprint.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' CLEAR
'------------------------------------------------------------------------------
        ClearOutputBlock GetSheet(SH_TRADECALC), OutputFirstRow(SH_TRADECALC), TC_NCOLS
        ClearOutputBlock GetSheet(SH_BUCKETS), OutputFirstRow(SH_BUCKETS), BK_NCOLS
        ClearOutputBlock GetSheet(SH_HEDGING), OutputFirstRow(SH_HEDGING), HS_NCOLS
        ClearOutputBlock GetSheet(SH_RESULTS), OutputFirstRow(SH_RESULTS), RS_NCOLS
        ClearOutputBlock GetSheet(SH_CHECKS), OutputFirstRow(SH_CHECKS), CK_NCOLS
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
    ByVal SavedCalculation As XlCalculation, _
    ByVal SavedEvents As Boolean, _
    ByVal SavedScreen As Boolean) _
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
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Details      As String    'Collected failures
    Dim OneFailure   As String    'Failure of one restoration

'------------------------------------------------------------------------------
' RESTORE
'------------------------------------------------------------------------------
    'Every helper runs even when an earlier one failed.
        OneFailure = RestoreCalculation(SavedCalculation)
        Details = AppendDetail(Details, OneFailure)
        OneFailure = RestoreEvents(SavedEvents)
        Details = AppendDetail(Details, OneFailure)
        OneFailure = RestoreScreenUpdating(SavedScreen)
        Details = AppendDetail(Details, OneFailure)
        RestoreSettings = Details

End Function


Private Function RestoreCalculation( _
    ByVal SavedValue As XlCalculation) _
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
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' RESTORE
'------------------------------------------------------------------------------
        On Error GoTo Failed
        If gTestFault = "cleanup" Then
            Err.Raise ERR_INJECTED_FAULT, "M_Main.RestoreCalculation", "Injected cleanup failure."
        End If
        Application.Calculation = SavedValue
        If Application.Calculation <> SavedValue Then
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
    ByVal SavedValue As Boolean) _
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
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' RESTORE
'------------------------------------------------------------------------------
        On Error GoTo Failed
        Application.EnableEvents = SavedValue
        If Application.EnableEvents <> SavedValue Then
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
    ByVal SavedValue As Boolean) _
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
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' RESTORE
'------------------------------------------------------------------------------
        On Error GoTo Failed
        Application.ScreenUpdating = SavedValue
        If Application.ScreenUpdating <> SavedValue Then
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
    ByVal Details As String, _
    ByVal NewDetail As String) _
    As String
'
'==============================================================================
'                                 AppendDetail
'------------------------------------------------------------------------------
' PURPOSE
'   Append one failure to a "; "-separated list; an empty one is skipped.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' APPEND
'------------------------------------------------------------------------------
        If Len(NewDetail) = 0 Then
            AppendDetail = Details
        ElseIf Len(Details) = 0 Then
            AppendDetail = NewDetail
        Else
            AppendDetail = Details & "; " & NewDetail
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
    ByVal Operation As Long, _
    ByVal Completed As Boolean, _
    ByVal CleanupDetails As String) _
    As String
'
'==============================================================================
'                                ReportOutcome
'------------------------------------------------------------------------------
' PURPOSE
'   Tell the user how a completed operation went, and build its result line.
'
' INPUTS
'   Operation: the operation that ran.
'   Completed: the engine's result; False means the inputs could not be
'   loaded.
'   CleanupDetails: restoration failures; empty when Excel was restored.
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
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Message      As String           'Message for the user; empty for none
    Dim IconStyle    As VbMsgBoxStyle    'Message-box icon
    Dim ResultLine   As String           'Machine-readable result line

'------------------------------------------------------------------------------
' BUILD THE MESSAGE
'------------------------------------------------------------------------------
    'Messages and sheet activation are those of the original macros. A
    'completed run can still have excluded trades, so it points to Checks
    'whenever there are errors or warnings. ShowChecksSheet never raises.
        Select Case Operation
            Case OP_RUN
                If Completed Then
                    Message = "SA-CCR run completed." & vbCrLf & vbCrLf & _
                          "Economic trades used: " & CORE_Engine.EconomicTradesUsed & " of " & _
                          CORE_Engine.EconomicTradesRead & vbCrLf & _
                          "Legs (rows) used: " & CORE_Engine.TradesUsed & " of " & CORE_Engine.TradesRead & vbCrLf & _
                          TotalLines() & _
                          "Errors: " & CORE_Engine.ErrorCount & "   Warnings: " & CORE_Engine.WarningCount & _
                          vbCrLf & RunClassLine()
                    If CORE_Engine.IncompleteCount > 0 Then
                        Message = Message & vbCrLf & "EAD withheld for " & CORE_Engine.IncompleteCount & _
                              " netting set(s) with rejected trades or invalid inputs: " & _
                              CORE_Engine.WithheldByRegimeText() & "."
                    End If
                    If Len(mReconLine) > 0 Then
                        Message = Message & vbCrLf & vbCrLf & mReconLine
                    End If
                    If CORE_Engine.ErrorCount > 0 Or CORE_Engine.WarningCount > 0 Then
                        Message = Message & vbCrLf & vbCrLf & "See the Checks sheet for details."
                    End If
                    IconStyle = IIf(CORE_Engine.ErrorCount > 0, vbExclamation, vbInformation)
                Else
                    Message = "SA-CCR run stopped - see the Checks sheet."
                    IconStyle = vbCritical
                    ShowChecksSheet
                End If
            Case OP_VALIDATE
                ShowChecksSheet
                Message = "Validation finished: " & CORE_Engine.ErrorCount & " error(s), " & _
                      CORE_Engine.WarningCount & " warning(s)." & vbCrLf & RunClassLine()
                IconStyle = IIf(CORE_Engine.ErrorCount > 0, vbExclamation, vbInformation)
            Case OP_CLEAR
                Message = ""
                IconStyle = vbInformation
        End Select

'------------------------------------------------------------------------------
' BUILD THE RESULT LINE
'------------------------------------------------------------------------------
    'Numbers use Str$, which always writes a decimal point, so the line
    'reads the same in every locale.
        ResultLine = "RESULT=" & IIf(Completed, "OK", "STOPPED") & "; operation=" & OperationName(Operation)
        If Operation <> OP_CLEAR Then
            ResultLine = ResultLine & _
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
        ResultLine = ResultLine & "; cleanup=" & IIf(Len(CleanupDetails) = 0, "PASS", "FAIL")

'------------------------------------------------------------------------------
' REPORT A CLEANUP FAILURE
'------------------------------------------------------------------------------
    'Settings left changed matter more than the run summary: raise in
    'silent mode, otherwise add a warning and use the critical icon.
        If Len(CleanupDetails) > 0 Then
            If gSilent Then
                Err.Raise ERR_CLEANUP_FAILED, "M_Main.ExecuteOperation", _
                          "Excel settings could not be restored: " & CleanupDetails & " (" & ResultLine & ")"
            End If
            If Len(Message) > 0 Then
                Message = Message & vbCrLf & vbCrLf
            End If
            Message = Message & "Excel settings could not be restored: " & CleanupDetails & _
                  ". Check calculation mode, events and screen updating."
            IconStyle = vbCritical
        End If

'------------------------------------------------------------------------------
' SHOW
'------------------------------------------------------------------------------
        If Len(Message) > 0 Then
            Notify Message, IconStyle
        End If
        ReportOutcome = ResultLine

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
    ByVal ErrorNumber As Long, _
    ByVal ErrorSource As String, _
    ByVal ErrorDescription As String, _
    ByVal CleanupDetails As String)
'
'==============================================================================
'                                ReportFailure
'------------------------------------------------------------------------------
' PURPOSE
'   Report an operation that did not complete: raise it in silent mode,
'   otherwise show it.
'
' INPUTS
'   ErrorNumber, ErrorSource, ErrorDescription: the primary error, as
'      copied before cleanup.
'   CleanupDetails: restoration failures; empty when Excel was restored.
'
' ERROR POLICY
'   In silent mode raises the primary error unchanged, with any cleanup
'   failure appended to its description.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Message   As String    'Message for the user

'------------------------------------------------------------------------------
' RAISE IN SILENT MODE
'------------------------------------------------------------------------------
        If Len(CleanupDetails) > 0 Then
            ErrorDescription = ErrorDescription & " Excel settings could not be restored: " & CleanupDetails & "."
        End If
        If gSilent Then
            Err.Raise ErrorNumber, ErrorSource, ErrorDescription
        End If

'------------------------------------------------------------------------------
' SHOW OTHERWISE
'------------------------------------------------------------------------------
    'The project's own errors carry a readable description; anything else
    'is shown with its number.
        If ErrorNumber = ERR_RUN_ACTIVE Or ErrorNumber = ERR_CHECKS_WRITE Then
            Message = ErrorDescription
        Else
            Message = "Unexpected error " & ErrorNumber & ": " & ErrorDescription
        End If
        MsgBox Message, vbCritical, MSG_TITLE

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
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Regimes   As Variant    'The regimes, CRR first
    Dim i         As Long       'Index into regimes

'------------------------------------------------------------------------------
' BUILD
'------------------------------------------------------------------------------
        Regimes = Array(RG_CRR, RG_BCBS)
        For i = LBound(Regimes) To UBound(Regimes)
            If CORE_Engine.RegimeSetCount(CStr(Regimes(i))) > 0 Then
                TotalLines = TotalLines & "Total EAD " & Regimes(i) & ": " & _
                             Format$(CORE_Engine.TotalEAD(CStr(Regimes(i))), "#,##0") & _
                             " (" & CORE_Engine.TotalStatusText(CStr(Regimes(i))) & ")" & vbCrLf
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
    ByVal Operation As Long) _
    As String
'
'==============================================================================
'                                OperationName
'------------------------------------------------------------------------------
' PURPOSE
'   Name an operation for the result line: run, validate or clear.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' NAME
'------------------------------------------------------------------------------
        Select Case Operation
            Case OP_RUN
                OperationName = "run"
            Case OP_VALIDATE
                OperationName = "validate"
            Case Else
                OperationName = "clear"
        End Select

End Function


Private Sub Notify( _
    ByVal Message As String, _
    ByVal IconStyle As VbMsgBoxStyle)
'
'==============================================================================
'                                    Notify
'------------------------------------------------------------------------------
' PURPOSE
'   Show a message to the user, or print it in silent mode unless
'   gQuietLog is set.
'
' INPUTS
'   Message: the text to show.
'   IconStyle: message-box icon, for example vbInformation or vbCritical.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' SHOW OR PRINT
'------------------------------------------------------------------------------
        gLastMessage = Message
        If gSilent Then
            If Not gQuietLog Then
                Debug.Print Message
            End If
        Else
            MsgBox Message, IconStyle, MSG_TITLE
        End If

End Sub
