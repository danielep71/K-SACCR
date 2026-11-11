Attribute VB_Name = "M_Main"
'==============================================================================
' MODULE: M_Main
'------------------------------------------------------------------------------
' PURPOSE
'   Provide the macros behind the K-SACCR Ribbon tab and the staging-sheet
'   buttons, run each operation
'   with Excel prepared for speed, put Excel back exactly as it was, and
'   report the outcome to the user.
'
' PUBLIC SURFACE
'   RunSACCR: validate the inputs, calculate and write every output sheet.
'   ValidateInputs: validate only; findings go to the Checks sheet.
'   ClearOutputs: empty every output sheet.
'   RunSACCR_Silent: RunSACCR for automation. Returns a machine-readable
'   result line and raises errors instead of showing them.
'   ValidateInputs_Silent and ClearOutputs_Silent: the same for
'   ValidateInputs and ClearOutputs, for the automated protection checks
'   of tools\Check-SACCR-Protection.ps1 (#535).
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
'   ExportNettingSetsCsvTemplate and ExportTradesCsvTemplate: write the
'   blank CSV template of the input contract to a file (#206);
'   ExportCsvTemplate_Silent: the same to a given file, for automation and
'   tests.
'   ExportImportSpecification: save the CSV import specification, an HTML
'   page with an example of each file, to a file and offer to open it
'   (#473); ExportImportSpecification_Silent: the same to a given file,
'   for automation and tests.
'   CheckTradesCsv and CheckNettingSetsCsv: check a CSV file against the
'   input contract and report the outcome, writing nothing (#207), from
'   the Macros dialog; ReadCsv_Silent: the same for a given file, for
'   automation and tests.
'   ImportNettingSetsCsv and ImportTradesCsv: load a CSV file into the
'   staging sheet Import_NettingSets or Import_Trades, as text, for
'   checking and correction before it replaces the inputs (#208);
'   ImportCsv_Silent: the same for a given file. StagingEdited: the
'   change handler of the staging sheets. StagingFormulaCell: the first
'   formula in a staging block, for staging validation and commit.
'   ValidateStaging: check the rows of both staging sheets against the
'   contract, each other and Params, and list the findings on Checks with
'   the origin IMPORT and a link to each cell, which is marked (#209);
'   ValidateStaging_Silent: the same, for automation and tests.
'   StagingValidationCurrent: whether the last staging validation was
'   VALID and nothing it judged has changed, for the commit (#210, #211).
'   RefreshStagingStatus: show OUT OF DATE on a staging sheet that reads
'   VALID once what the last validation judged has changed, and VALID
'   again when it matches (#210); called when a staging sheet is
'   activated.
'   CommitStaging_Silent: replace NettingSets and Trades with the staged
'   rows once a VALID staging validation still holds, putting them back
'   on any failure (#211); CommitImport: the same after a confirmation
'   that names each input sheet with its row counts.
'   RestoreRegulatoryParameters: put the reference parameters and the
'   factor table on Params back to the regulatory set, after a
'   confirmation that lists what is overwritten, putting Params back on
'   any failure (#357); RestoreRegulatoryParameters_Silent: the same
'   without a confirmation, for automation and tests.
'   Auto_Open: run by Excel when the workbook opens; protects every sheet
'   again for the user interface only. ProtectWorkbook: protect every
'   sheet and the workbook structure, without a password, so that the
'   macros write while the user changes only the unlocked cells (#421);
'   also run when each operation starts.
'   OperationRunning: whether an operation is in progress, for the
'   Ribbon's enabled states (#216); RibbonStateChanged: refresh the tab
'   once Excel is idle, also for ThisWorkbook's change handler.
'   NavigationBarShown, ShowNavigationBar and ToggleNavigationBar: whether
'   the navigation bar of every sheet is shown, and show or hide it on
'   every sheet at once (#620); SetNavigationBar_Silent: the same for
'   automation and tests. NavigationBarBeforeSave and
'   NavigationBarAfterSave: show the bar for a save and hide it again
'   after, for ThisWorkbook, so that a saved file always shows it.
'   The Ribbon (M_Ribbon) calls the first three, ExportEvidence, the two
'   template exports, the two imports, ValidateStaging, CommitImport,
'   ExportImportSpecification, RestoreRegulatoryParameters,
'   ShowNavigationBar and NavigationBarShown; the
'   buttons of the staging sheets call the
'   imports, ValidateStaging, CommitImport and ExportImportSpecification
'   by name (#217).
'   These macros are not listed in docs/PUBLIC_API.txt.
'
' DEPENDENCIES
'   M_Ribbon to refresh the K-SACCR tab after a macro that changes a state
'   it shows (#216); CORE_Engine for the calculation and its run counters; CORE_Reconcile
'   for the comparison of the bank figures with Results and TradeCalc
'   (#45), whose sheets this module reads and writes; CORE_Export for the
'   evidence report and its UTF-8 encoding; CORE_CsvReader and
'   CORE_CsvSchema for the CSV check, the staging load and the staging
'   validation; CORE_ImportSpec for the import specification page;
'   CORE_Util and CORE_Config for the output and staging
'   sheets, the CSV template header lines and the error numbers.
'
' STATE OWNERSHIP
'   Each operation captures calculation mode, events and screen updating,
'   changes them for the run, and restores the captured values on success
'   and on failure. Owns gSilent (messages to the Immediate window instead
'   of message boxes, and errors raised to the caller), the re-entry flag
'   mRunning, the reconciliation outcomes of the last operation
'   (mReconState, mTradeReconState, mReconLine), and the test seams
'   gTestFault, gQuietLog (silent-mode messages dropped instead of
'   printed), gLastMessage and gKeepProtection.
'
' ERROR POLICY
'   The primary error is copied before cleanup. Each setting is restored by
'   its own helper, so one failed restoration does not stop the others, and
'   a cleanup failure is reported separately from the primary error. With
'   gSilent False the outcome is shown in a message box; with gSilent True
'   it is raised to the caller. Errors in the inputs are not VBA errors:
'   the engine lists them on the Checks sheet, and so does the staging
'   validation for the staged rows.
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
'   raise ERR_INJECTED_FAULT in WriteTextFile, which writes the evidence
'   report, the CSV templates and the import specification, before the new file is opened, while it
'   is written and while the older file is replaced (#312).
'   gTestFault = "staging-write" raises ERR_INJECTED_FAULT in LoadStaging
'   once the new staging block has been written (#208).
'   gTestFault = "commit-write" raises ERR_INJECTED_FAULT in
'   CommitStaging once the first input sheet has been written;
'   "commit-verify" makes its read-back fail, and "commit-rollback" makes
'   the write-back of the input sheets fail as well (#211).
'   gTestFault = "restore-write", "restore-verify" and "restore-rollback"
'   do the same in RestoreParameters for Params (#357).
'   tests/modules/TEST_MainState.bas uses the first four and the restore
'   faults;
'   tests/modules/TEST_Reconcile.bas uses "recalc", "reconcile" and the
'   evidence faults; tests/modules/TEST_Import.bas uses "cleanup",
'   "evidence-write", "staging-write" and the commit faults.
'   gLastMessage keeps the last message passed to Notify (a failure
'   message is not kept), so that a test can read it in silent mode.
'   gKeepProtection lists, between bars, sheets whose protection
'   ProtectWorkbook leaves as it is, so that a test can make a sheet
'   unwritable for the macros too (#421); TEST_MainState uses it.
'   Production code never sets gTestFault or gKeepProtection.
'
' KNOWN DEVIATION
'   This standard module lives in src/workbook. See the known deviations in
'   docs/REPOSITORY_STRUCTURE.md.
'
' COMPATIBILITY
'   Excel VBA; no references beyond the defaults.
'
' UPDATED
'   2026-11-09
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

    'Recalculations tried before an operation is refused because Excel did
    'not report it finished.
        Private Const RECALC_ATTEMPTS As Long = 3

    'Title of every message box: the project name.
        Private Const MSG_TITLE     As String = "K-SACCR"

    'A CSV check's message. A message box shows 1,024 characters at most,
    'so the message is kept shorter and lists fewer findings when it must.
        Private Const CSV_MESSAGE_FINDINGS  As Long = 10      'Findings listed at most
        Private Const CSV_MESSAGE_WIDTH     As Long = 120     'Characters of a finding's line
        Private Const CSV_MESSAGE_COLUMN    As Long = 32      'Characters of its column name
        Private Const CSV_MESSAGE_PATH      As Long = 200     'Characters of the file path, from its end
        Private Const CSV_MESSAGE_LENGTH    As Long = 1000    'Characters of the whole message

    'Staging block formats (#208), as the template's input cells: font
    'size, the grey fill of the CSV line column, the input fill of the
    'values and the input border colour, as RGB values.
        Private Const STAGING_FONT_SIZE     As Long = 9           'Points
        Private Const STAGING_LINE_FILL     As Long = 16250871    'RGB(247, 247, 247)
        Private Const STAGING_VALUE_FILL    As Long = 13431551    'RGB(255, 242, 204)
        Private Const STAGING_BORDER        As Long = 12566463    'RGB(191, 191, 191)
        Private Const EXCEL_CELL_CHARS      As Long = 32767       'Characters a cell holds

    'Staging validation (#209): the fill of a cell with an error, Excel's
    'red of a bad value, and the findings listed on Checks and marked at
    'most; the summary line gives the full count.
        Private Const STAGING_ERROR_FILL    As Long = 13551615    'RGB(255, 199, 206)
        Private Const STAGING_LISTED_MAX    As Long = 10000       'Findings listed and marked
    'The code of the finding of a run that would stop at Params, and the
    'table columns of its row searched for its reference.
        Private Const STAGING_ROW_SEARCH    As Long = 16

    'Restore regulatory parameters (#357): the overwritten values its
    'confirmation lists at most; a last line counts the others.
        Private Const RESTORE_LISTED        As Long = 12

    'Reconciliations of bank figures (#45).
        Private Const RECON_NS      As Long = 1    'BankNettingSets against Results
        Private Const RECON_TR      As Long = 2    'BankTrades against TradeCalc

'------------------------------------------------------------------------------
' TYPES
'------------------------------------------------------------------------------
    'One finding of a staging validation (#209), as Checks lists it.
        Private Type StagedFinding
            SheetName       As String     'Sheet of its cell
            RowNumber       As Long       'Row of its cell; 0 for none
            CellAddress     As String     'Its cell, as A1; "" for none
            Code            As String     'Stable finding code
            FieldName       As String     'Structured field context
            Message         As String     'Code, column, message and CSV line
            Marked          As Boolean    'The cell is staged and gets the error fill
        End Type

'------------------------------------------------------------------------------
' MODULE STATE
'------------------------------------------------------------------------------
    'gSilent switches messages to the Immediate window and errors to the
    'caller. gTestFault and gLastMessage are the test seams described above.
    'gQuietLog drops the messages of silent mode instead of printing them,
    'for the test suites (TEST_Suite.Test_RunAll, TEST_MainState,
    'TEST_Reconcile and TEST_Import), whose output would overflow the
    'Immediate window.
    'mRunning refuses a second operation while one is in progress.
        Public gSilent      As Boolean    'Silent mode for automation and tests
        Public gTestFault   As String     'Test seam: "", "operation", "cleanup", "recalc", "recalc-error", "reconcile", "evidence-*", "staging-write", "commit-*" or "restore-*"
        Public gQuietLog    As Boolean    'Test seam: silent-mode messages not printed
        Public gLastMessage As String     'Test seam: the last message passed to Notify
        Public gKeepProtection As String  'Test seam: sheets, between bars, ProtectWorkbook leaves as they are
        Private mRunning    As Boolean    'True while an operation is in progress
    'The navigation bar was hidden when a save began, and is hidden again
    'after it, or by the restore scheduled at mBarRestoreAt when the save
    'never ends, as a cancelled Save As (#620).
        Private mBarHiddenForSave As Boolean
        Private mBarRestoreAt     As Date      'When the restore was scheduled; 0 for none
        Private mBarRestoreName   As String    'The procedure it was scheduled with
    'The reconciliations of the last operation (#45), for its message and
    'result line: NONE, DONE, FAILED or UNAVAILABLE, and the message lines.
        Private mReconState      As String    'Netting-set reconciliation outcome
        Private mTradeReconState As String    'Trade reconciliation outcome
        Private mReconLine       As String    'Lines for the run message; "" for none


'
'------------------------------------------------------------------------------
'
'                              WORKBOOK PROTECTION
'
'------------------------------------------------------------------------------
'

Public Sub Auto_Open()
'
'==============================================================================
'                                  Auto_Open
'------------------------------------------------------------------------------
' PURPOSE
'   Run by Excel when a user opens the workbook with macros enabled:
'   protect every sheet again for the user interface only (#421).
'
' UPDATED
'   2026-11-03
'==============================================================================
'

'------------------------------------------------------------------------------
' PROTECT
'------------------------------------------------------------------------------
        ProtectWorkbook

End Sub


Public Sub ProtectWorkbook()
'
'==============================================================================
'                               ProtectWorkbook
'------------------------------------------------------------------------------
' PURPOSE
'   Protect every worksheet and the workbook structure, without a
'   password (#421). A sheet is protected for the user interface only, so
'   that the macros write it while the user changes only its unlocked
'   cells (docs/workbook/protection.json), with the options that
'   PROTECT_SORTING, PROTECT_FILTERING, PROTECT_INSERTING_ROWS and
'   PROTECT_OUTLINING give it. Excel saves neither UserInterfaceOnly nor
'   EnableOutlining with the file, so this runs when the workbook opens
'   (Auto_Open) and when each operation starts (ExecuteOperation,
'   LoadStaging, CheckStaging and CommitStaging): a sheet unprotected for
'   maintenance is protected again then. A sheet already protected for
'   the user interface only is left as it is.
'   The workbook's Saved flag is kept: protecting again is not a change
'   of the file, so opening the workbook with macros enabled does not ask
'   to save it on closing (#620).
'
' USAGE
'   Also listed in the Macros dialog, to protect the workbook again at
'   once after maintenance.
'
' ERROR POLICY
'   Never raises. A sheet, or the structure, protected with a password is
'   left as it is, and so is a sheet the test seam gKeepProtection lists.
'
' UPDATED
'   2026-11-07
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Sheet       As Worksheet    'Each worksheet
    Dim WasSaved    As Boolean      'The workbook's Saved flag before

'------------------------------------------------------------------------------
' PROTECT
'------------------------------------------------------------------------------
        WasSaved = ThisWorkbook.Saved
        For Each Sheet In ThisWorkbook.Worksheets
            ProtectSheet Sheet
        Next Sheet
        On Error GoTo Failed
        If Not ThisWorkbook.ProtectStructure Then
            ThisWorkbook.Protect Structure:=True, Windows:=False
        End If

Failed:
        ThisWorkbook.Saved = WasSaved

End Sub


Private Sub ProtectSheet( _
    ByVal Sheet As Worksheet)
'
'==============================================================================
'                                 ProtectSheet
'------------------------------------------------------------------------------
' PURPOSE
'   Protect one worksheet for the user interface only, with its options,
'   unless it already is; let the user open and close its row groups when
'   PROTECT_OUTLINING lists it.
'
' ERROR POLICY
'   Never raises: a sheet protected with a password stays as it is, and
'   so does one that gKeepProtection lists.
'
' UPDATED
'   2026-11-03
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Listed  As String    'The sheet's name between bars, as the lists hold it

'------------------------------------------------------------------------------
' PROTECT
'------------------------------------------------------------------------------
        On Error GoTo Failed
        Listed = "|" & Sheet.Name & "|"
        If InStr(1, gKeepProtection, Listed, vbBinaryCompare) > 0 Then
            Exit Sub
        End If
        If Not (Sheet.ProtectContents And Sheet.ProtectionMode) Then
            Sheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True, _
                          AllowSorting:=(InStr(1, PROTECT_SORTING, Listed, vbBinaryCompare) > 0), _
                          AllowFiltering:=(InStr(1, PROTECT_FILTERING, Listed, vbBinaryCompare) > 0), _
                          AllowInsertingRows:=(InStr(1, PROTECT_INSERTING_ROWS, Listed, vbBinaryCompare) > 0)
        End If
        If InStr(1, PROTECT_OUTLINING, Listed, vbBinaryCompare) > 0 Then
            Sheet.EnableOutlining = True
        End If

Failed:

End Sub


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
'   show Results and report the economic trades and legs used, the total
'   EAD of each regime and the number of errors and warnings. A run that
'   stops shows Checks.
'
' USAGE
'   Called by Run on the K-SACCR Ribbon tab.
'
' UPDATED
'   2026-11-06
'==============================================================================
'

'------------------------------------------------------------------------------
' RUN
'------------------------------------------------------------------------------
        ExecuteOperation OP_RUN
        RibbonStateChanged

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
'   trade_reconciliation, codes and cleanup=PASS or FAIL, separated by "; ".
'   codes lists distinct ERROR/WARNING codes in encounter order, comma-separated;
'   it is empty when the current run has none.
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
'   2026-11-04
'==============================================================================
'

'------------------------------------------------------------------------------
' RUN SILENTLY
'------------------------------------------------------------------------------
        RunSACCR_Silent = SilentOperation(OP_RUN)

End Function


Public Function ValidateInputs_Silent() As String
'
'==============================================================================
'                            ValidateInputs_Silent
'------------------------------------------------------------------------------
' PURPOSE
'   Run ValidateInputs unattended, as RunSACCR_Silent runs RunSACCR.
'   tools\Check-SACCR-Protection.ps1 calls it through Application.Run.
'
' RETURNS
'   The result line of ReportOutcome: RESULT=OK or STOPPED,
'   operation=validate, and the fields RunSACCR_Silent returns.
'
' ERROR POLICY
'   As RunSACCR_Silent.
'
' UPDATED
'   2026-11-04
'==============================================================================
'
'------------------------------------------------------------------------------
' VALIDATE SILENTLY
'------------------------------------------------------------------------------
        ValidateInputs_Silent = SilentOperation(OP_VALIDATE)

End Function


Public Function ClearOutputs_Silent() As String
'
'==============================================================================
'                             ClearOutputs_Silent
'------------------------------------------------------------------------------
' PURPOSE
'   Run ClearOutputs unattended, as RunSACCR_Silent runs RunSACCR.
'   tools\Check-SACCR-Protection.ps1 calls it through Application.Run.
'
' RETURNS
'   The result line of ReportOutcome: RESULT=OK, operation=clear and
'   cleanup=PASS or FAIL.
'
' ERROR POLICY
'   As RunSACCR_Silent.
'
' UPDATED
'   2026-11-04
'==============================================================================
'
'------------------------------------------------------------------------------
' CLEAR SILENTLY
'------------------------------------------------------------------------------
        ClearOutputs_Silent = SilentOperation(OP_CLEAR)

End Function


Private Function SilentOperation( _
    ByVal Operation As Long) _
    As String
'
'==============================================================================
'                               SilentOperation
'------------------------------------------------------------------------------
' PURPOSE
'   Run one operation with gSilent set, for RunSACCR_Silent,
'   ValidateInputs_Silent and ClearOutputs_Silent.
'
' INPUTS
'   Operation: OP_RUN, OP_VALIDATE or OP_CLEAR.
'
' RETURNS
'   The result line of ExecuteOperation.
'
' STATE OWNERSHIP
'   Sets gSilent for the operation and restores its previous value
'   afterwards, also when the operation raises.
'
' ERROR POLICY
'   Re-raises the operation's error after gSilent has been restored.
'
' UPDATED
'   2026-11-04
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
        SilentOperation = ExecuteOperation(Operation)
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
'   Validate inputs on the K-SACCR Ribbon tab (Run menu), or the Macros
'   dialog.
'
' UPDATED
'   2026-10-06
'==============================================================================
'

'------------------------------------------------------------------------------
' VALIDATE
'------------------------------------------------------------------------------
        ExecuteOperation OP_VALIDATE
        RibbonStateChanged

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
'   Clear outputs on the K-SACCR Ribbon tab (Run menu), or the Macros
'   dialog.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' CLEAR
'------------------------------------------------------------------------------
        ExecuteOperation OP_CLEAR
        RibbonStateChanged

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
'   reconciliation sheet was changed after the run (#315), or a staging
'   validation rewrote Checks (#209); the evidence export refuses such
'   outputs.
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
'   2026-10-28
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
'     =ResultsStatusText(SACCR_RunInputs, NettingSets!$C:$T, Trades!$C:$AC,
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
'   2026-10-22
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
'        BankNettingSets!$1:$1048576, NettingSets!$C:$T, Trades!$C:$AC,
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
'   2026-10-22
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
'        SACCR_RunInputs, BankTrades!$1:$1048576, NettingSets!$C:$T,
'        Trades!$C:$AC, Params!$1:$1048576)
'
' INPUTS
'   StoredRecon: the value of SACCR_TradeReconInputs; the others as
'   ReconciliationStatusText, with BankTrades as BankRows.
'
' RETURNS
'   As ReconStatus.
'
' UPDATED
'   2026-10-22
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
                          " and press Run SA-CCR (on the K-SACCR tab)."
        ElseIf StoredFingerprint = ReconFingerprint(ReconKind) Then
            ReconStatus = "Reconciliation current (bank figures and inputs " & StoredFingerprint & ")"
        Else
            ReconStatus = "OUT OF DATE - bank figures or inputs changed since the last run; press Run SA-CCR (on the K-SACCR tab)."
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
'   Export evidence on the K-SACCR Ribbon tab, or the Macros dialog.
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
        Notify "Evidence exported, " & LineCount & " lines:" & vbLf & vbLf & CStr(ChosenFile), vbInformation
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
'   Tell why no evidence report may be written now: Checks shows a
'   staging validation (#209), the results are not those of the current
'   inputs, an output cell was changed after the run,
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
'   2026-11-04
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
    'A staging validation rewrites Checks (#209); say so rather than call
    'it an edit.
        If ChecksHasImportLines() Then
            EvidenceProblem = "Checks shows a staging validation, not the checks of the last run; press Run " & _
                              "SA-CCR first"
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
    'The output sheets are protected for the user interface only, and a
    'user may unprotect one: ResultsStatus has compared them with what the
    'run wrote, which a run that recorded nothing cannot be.
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
'   Write the evidence report to FilePath through WriteTextFile.
'
' RETURNS
'   The number of lines written.
'
' ERROR POLICY
'   Raises the error of WriteTextFile.
'
' UPDATED
'   2026-10-23
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ReportText   As String    'The report

'------------------------------------------------------------------------------
' WRITE
'------------------------------------------------------------------------------
        ReportText = EvidenceText()
        WriteTextFile FilePath, ReportText
        WriteEvidence = (Len(ReportText) - Len(Replace(ReportText, vbCrLf, ""))) \ 2

End Function


Private Sub WriteTextFile( _
    ByVal FilePath As String, _
    ByVal FileText As String)
'
'==============================================================================
'                                WriteTextFile
'------------------------------------------------------------------------------
' PURPOSE
'   Write FileText to FilePath as UTF-8 without a byte-order mark,
'   replacing any file of that name, hidden ones included; a read-only
'   file is left as it is and the write fails. Used by the evidence export,
'   the CSV template export and the import specification export.
'
' METHOD
'   The text is written whole to a new file beside FilePath first. Only
'   then is an older file moved aside, the new one renamed to FilePath and
'   the older one deleted, so a failed write leaves the older file as it
'   was (#312). A failed replacement puts the older file back with its
'   attributes. The replacement is as atomic as the file system's rename,
'   no more.
'
' ERROR POLICY
'   Raises the first error when the file cannot be written or replaced,
'   after closing the file number, deleting the new file and putting an
'   older file back; an error in that cleanup does not replace the first
'   one. An older file that cannot be deleted once replaced is left beside
'   the new one.
'
' TEST SEAM
'   gTestFault = "evidence-open", "evidence-write" and "evidence-replace"
'   raise ERR_INJECTED_FAULT here, whichever export calls it.
'
' UPDATED
'   2026-10-23
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim TextBytes()         As Byte       'The text's UTF-8 bytes
    Dim FileNumber          As Integer    'File number; 0 while none is open
    Dim NewPath             As String     'The file before it replaces FilePath; "" once renamed
    Dim OldPath             As String     'An older FilePath while it is moved aside; "" otherwise
    Dim OldAttributes       As Long       'Its attributes, put back if the replacement fails
    Dim AttributesCleared   As Boolean    'Its attributes were cleared
    Dim ErrorNumber         As Long       'First error; 0 on success
    Dim ErrorText           As String     'Its description

'------------------------------------------------------------------------------
' WRITE THE NEW FILE
'------------------------------------------------------------------------------
        TextBytes = CORE_Export.Utf8Bytes(FileText)
        On Error GoTo Failed
        If FileExists(FilePath) Then
            If (GetAttr(FilePath) And vbReadOnly) <> 0 Then
                Err.Raise ERR_EXPORT_WRITE, "M_Main.WriteTextFile", "the file is read-only"
            End If
        End If
        NewPath = SiblingPath(FilePath, "new")
        If gTestFault = "evidence-open" Then
            Err.Raise ERR_INJECTED_FAULT, "M_Main.WriteTextFile", "Injected failure before the file was opened."
        End If
        FileNumber = FreeFile
        Open NewPath For Binary Access Write As #FileNumber
        Put #FileNumber, , TextBytes
        If gTestFault = "evidence-write" Then
            Err.Raise ERR_INJECTED_FAULT, "M_Main.WriteTextFile", "Injected failure while the file was written."
        End If
        If LOF(FileNumber) <> UBound(TextBytes) - LBound(TextBytes) + 1 Then
            Err.Raise ERR_EXPORT_WRITE, "M_Main.WriteTextFile", "the file was not written whole"
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
            Err.Raise ERR_INJECTED_FAULT, "M_Main.WriteTextFile", "Injected failure while the older file was replaced."
        End If
        Name NewPath As FilePath
        NewPath = ""

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
            Err.Raise ErrorNumber, "M_Main.WriteTextFile", "Could not write " & FilePath & ": " & ErrorText
        End If
        Exit Sub

'------------------------------------------------------------------------------
' HANDLE ERROR
'------------------------------------------------------------------------------
Failed:
        ErrorNumber = Err.Number
        ErrorText = Err.Description
        Resume CleanUp

End Sub


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
'   2026-10-23
'==============================================================================
'
        DefaultEvidenceName = InWorkbookFolder("K-SACCR_evidence_" & Format$(Now, "yyyymmdd-hhnnss") & ".csv")

End Function


Private Function InWorkbookFolder( _
    ByVal FileName As String) _
    As String
'
'==============================================================================
'                              InWorkbookFolder
'------------------------------------------------------------------------------
' PURPOSE
'   FileName in the workbook's folder, for a save dialog to offer; FileName
'   alone when the workbook has no folder the dialog can start from.
'
' UPDATED
'   2026-10-23
'==============================================================================
'
        InWorkbookFolder = FileName
    'A workbook opened from OneDrive or SharePoint has a URL as its path,
    'which the save dialog cannot start from.
        If Len(ThisWorkbook.Path) > 0 And InStr(ThisWorkbook.Path, "://") = 0 Then
            InWorkbookFolder = ThisWorkbook.Path & Application.PathSeparator & FileName
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
'                             CSV TEMPLATE EXPORT
'
'------------------------------------------------------------------------------
'

Public Sub ExportNettingSetsCsvTemplate()
'
'==============================================================================
'                         ExportNettingSetsCsvTemplate
'------------------------------------------------------------------------------
' PURPOSE
'   Write the blank NettingSets CSV template of the input contract (#206)
'   to a file chosen by the user: its header line only.
'
' USAGE
'   NettingSets template on the K-SACCR Ribbon tab (Templates menu), or
'   the Macros dialog.
'
' ERROR POLICY
'   As ExportCsvTemplate; never raises.
'
' UPDATED
'   2026-10-23
'==============================================================================
'
        ExportCsvTemplate "NettingSets"

End Sub


Public Sub ExportTradesCsvTemplate()
'
'==============================================================================
'                           ExportTradesCsvTemplate
'------------------------------------------------------------------------------
' PURPOSE
'   Write the blank Trades CSV template of the input contract (#206) to a
'   file chosen by the user: its header line only.
'
' USAGE
'   Trades template on the K-SACCR Ribbon tab (Templates menu), or the
'   Macros dialog.
'
' ERROR POLICY
'   As ExportCsvTemplate; never raises.
'
' UPDATED
'   2026-10-23
'==============================================================================
'
        ExportCsvTemplate "Trades"

End Sub


Public Function ExportCsvTemplate_Silent( _
    ByVal Kind As String, _
    ByVal FilePath As String) _
    As String
'
'==============================================================================
'                          ExportCsvTemplate_Silent
'------------------------------------------------------------------------------
' PURPOSE
'   The CSV template exports for automation and tests: write the blank
'   template of Kind ("NettingSets" or "Trades") to FilePath without
'   asking.
'
' RETURNS
'   "RESULT=OK; file=..." or "RESULT=REFUSED; reason=...".
'
' ERROR POLICY
'   A file that cannot be written raises to the caller.
'
' UPDATED
'   2026-10-23
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim TemplateText   As String    'The template; "" for an unknown kind

'------------------------------------------------------------------------------
' CHECK AND WRITE
'------------------------------------------------------------------------------
        TemplateText = CsvTemplateText(Kind)
        If Len(TemplateText) = 0 Then
            ExportCsvTemplate_Silent = "RESULT=REFUSED; reason=unknown template kind " & Kind
        Else
            WriteTextFile FilePath, TemplateText
            ExportCsvTemplate_Silent = "RESULT=OK; file=" & FilePath
        End If

End Function


Private Sub ExportCsvTemplate( _
    ByVal Kind As String)
'
'==============================================================================
'                              ExportCsvTemplate
'------------------------------------------------------------------------------
' PURPOSE
'   Ask for a file, offering the template's file name in the workbook's
'   folder, and write the blank template of Kind to it through
'   WriteTextFile. A cancelled dialog writes nothing.
'
' ERROR POLICY
'   Reports a write failure in a message, or in the Immediate window in
'   silent mode; never raises.
'
' UPDATED
'   2026-10-23
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ChosenFile   As Variant    'Chosen file, or False when cancelled

'------------------------------------------------------------------------------
' ASK AND WRITE
'------------------------------------------------------------------------------
        On Error GoTo Failed
        ChosenFile = Application.GetSaveAsFilename(InitialFileName:=InWorkbookFolder(CsvTemplateFile(Kind)), _
                                               FileFilter:="CSV files (*.csv), *.csv", _
                                               Title:="Export " & Kind & " CSV template")
        If VarType(ChosenFile) = vbBoolean Then
            Exit Sub
        End If
        WriteTextFile CStr(ChosenFile), CsvTemplateText(Kind)
        Notify Kind & " CSV template exported:" & vbLf & vbLf & CStr(ChosenFile), vbInformation
        Exit Sub

Failed:
        Notify Kind & " CSV template not exported: " & Err.Description, vbCritical

End Sub


Private Function CsvTemplateText( _
    ByVal Kind As String) _
    As String
'
'==============================================================================
'                               CsvTemplateText
'------------------------------------------------------------------------------
' PURPOSE
'   The blank CSV template of Kind: its header line and a line feed, as in
'   docs/csv/templates; "" for an unknown kind.
'
' UPDATED
'   2026-10-23
'==============================================================================
'
        Select Case Kind
            Case "NettingSets"
                CsvTemplateText = CSV_NS_HEADER & vbLf
            Case "Trades"
                CsvTemplateText = CSV_TR_HEADER & vbLf
        End Select

End Function


Private Function CsvTemplateFile( _
    ByVal Kind As String) _
    As String
'
'==============================================================================
'                               CsvTemplateFile
'------------------------------------------------------------------------------
' PURPOSE
'   The file name of the blank CSV template of Kind; "" for an unknown
'   kind.
'
' UPDATED
'   2026-10-23
'==============================================================================
'
        Select Case Kind
            Case "NettingSets"
                CsvTemplateFile = CSV_NS_FILE
            Case "Trades"
                CsvTemplateFile = CSV_TR_FILE
        End Select

End Function


'
'------------------------------------------------------------------------------
'
'                         IMPORT SPECIFICATION EXPORT
'
'------------------------------------------------------------------------------
'

Public Sub ExportImportSpecification()
'
'==============================================================================
'                          ExportImportSpecification
'------------------------------------------------------------------------------
' PURPOSE
'   Save the CSV import specification (#473) to a file chosen by the user:
'   one HTML page, with no script and no external file, that describes
'   the NettingSets and Trades files and gives an example of each
'   (CORE_ImportSpec). A file that exists is replaced only once the user
'   confirms it; the user may then open the page.
'
' USAGE
'   Import specification on the K-SACCR Ribbon tab (Templates menu), the
'   Export import specifications button on Import_NettingSets and
'   Import_Trades, or the Macros dialog.
'
' ERROR POLICY
'   Reports a failure in a message, or in the Immediate window in silent
'   mode; never raises. A cancelled dialog or a refused replacement writes
'   nothing, and a failed write leaves an older file as it was
'   (WriteTextFile). The workbook is never changed, so the export works
'   with the protection on and whatever the inputs hold.
'
' UPDATED
'   2026-11-03
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ChosenFile   As Variant    'Chosen file, or False when cancelled
    Dim FilePath     As String     'Its path

'------------------------------------------------------------------------------
' ASK, CONFIRM AND WRITE
'------------------------------------------------------------------------------
        On Error GoTo Failed
        If mRunning Then
            Notify "An SA-CCR operation is already running.", vbExclamation
            Exit Sub
        End If
        ChosenFile = Application.GetSaveAsFilename(InitialFileName:=InWorkbookFolder(IMPORT_SPEC_FILE), _
                                               FileFilter:="Web pages (*.html), *.html", _
                                               Title:="Export import specifications")
        If VarType(ChosenFile) = vbBoolean Then
            Exit Sub
        End If
        FilePath = CStr(ChosenFile)
        If FileExists(FilePath) Then
            If MsgBox(FilePath & vbLf & vbLf & "already exists. Replace it?", _
                      vbYesNo + vbExclamation + vbDefaultButton2, MSG_TITLE) <> vbYes Then
                Exit Sub
            End If
        End If
        WriteTextFile FilePath, CORE_ImportSpec.ImportSpecificationText()

'------------------------------------------------------------------------------
' REPORT AND OFFER TO OPEN
'------------------------------------------------------------------------------
        If gSilent Then
            Notify "Import specifications exported:" & vbLf & vbLf & FilePath, vbInformation
        ElseIf MsgBox("Import specifications exported:" & vbLf & vbLf & FilePath & vbLf & vbLf & _
                      "Open the page now?", vbYesNo + vbInformation, MSG_TITLE) = vbYes Then
            OpenExportedPage FilePath
        End If
        Exit Sub

Failed:
        Notify "Import specifications not exported: " & Err.Description, vbCritical

End Sub


Public Function ExportImportSpecification_Silent( _
    ByVal FilePath As String, _
    Optional ByVal ReplaceExisting As Boolean = False) _
    As String
'
'==============================================================================
'                       ExportImportSpecification_Silent
'------------------------------------------------------------------------------
' PURPOSE
'   The import specification export for automation and tests: write the
'   page to FilePath without asking. A file that exists is replaced only
'   with ReplaceExisting.
'
' RETURNS
'   "RESULT=OK; bytes=N; file=..." or
'   "RESULT=REFUSED; reason=the file exists; file=...".
'
' ERROR POLICY
'   A file that cannot be written raises to the caller, as WriteTextFile
'   raises it, and an older file is left as it was.
'
' UPDATED
'   2026-11-03
'==============================================================================
'
        If FileExists(FilePath) And Not ReplaceExisting Then
            ExportImportSpecification_Silent = "RESULT=REFUSED; reason=the file exists; file=" & FilePath
            Exit Function
        End If
        WriteTextFile FilePath, CORE_ImportSpec.ImportSpecificationText()
        ExportImportSpecification_Silent = "RESULT=OK; bytes=" & IMPORT_SPEC_BYTES & "; file=" & FilePath

End Function


Private Sub OpenExportedPage( _
    ByVal FilePath As String)
'
'==============================================================================
'                               OpenExportedPage
'------------------------------------------------------------------------------
' PURPOSE
'   Open an exported page with the program Windows uses for HTML files.
'
' ERROR POLICY
'   Reports a page that cannot be opened in a message; never raises. The
'   page stays exported.
'
' UPDATED
'   2026-11-03
'==============================================================================
'
        On Error GoTo Failed
        ThisWorkbook.FollowHyperlink Address:=FilePath
        Exit Sub

Failed:
        Notify "The page was exported but could not be opened: " & Err.Description & vbLf & vbLf & FilePath, _
               vbExclamation

End Sub


'
'------------------------------------------------------------------------------
'
'                                 CSV CHECK
'
'------------------------------------------------------------------------------
'

Public Sub CheckTradesCsv()
'
'==============================================================================
'                               CheckTradesCsv
'------------------------------------------------------------------------------
' PURPOSE
'   Check a Trades CSV file chosen by the user against the CSV input
'   contract (#207) and report the outcome. Nothing is written to the
'   workbook.
'
' USAGE
'   Run from the Macros dialog (Alt+F8).
'
' ERROR POLICY
'   As CheckCsvFile; never raises.
'
' UPDATED
'   2026-10-23
'==============================================================================
'
        CheckCsvFile "Trades"

End Sub


Public Sub CheckNettingSetsCsv()
'
'==============================================================================
'                             CheckNettingSetsCsv
'------------------------------------------------------------------------------
' PURPOSE
'   Check a NettingSets CSV file chosen by the user against the CSV input
'   contract (#207) and report the outcome. Nothing is written to the
'   workbook.
'
' USAGE
'   Run from the Macros dialog (Alt+F8).
'
' ERROR POLICY
'   As CheckCsvFile; never raises.
'
' UPDATED
'   2026-10-23
'==============================================================================
'
        CheckCsvFile "NettingSets"

End Sub


Public Function ReadCsv_Silent( _
    ByVal Kind As String, _
    ByVal FilePath As String) _
    As String
'
'==============================================================================
'                               ReadCsv_Silent
'------------------------------------------------------------------------------
' PURPOSE
'   The CSV check for automation and tests: read FilePath as a file of
'   Kind ("Trades" or "NettingSets") without asking, and write nothing.
'
' RETURNS
'   "RESULT=ACCEPTED; records=N; file=..." or
'   "RESULT=REJECTED; findings=N; first=CODE|line|record|column; file=...".
'
' ERROR POLICY
'   An unknown kind and a file that cannot be read raise ERR_CSV_READ to
'   the caller.
'
' UPDATED
'   2026-10-23
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Result      As CsvReadResult    'What the reader found
    Dim Values()    As String           'The records of an accepted file
    Dim Lines()     As Long             'Their lines
    Dim Findings()  As CsvFinding       'The findings

'------------------------------------------------------------------------------
' READ
'------------------------------------------------------------------------------
        ReadCsvFile FilePath, Kind, Result, Values, Lines, Findings
        If Result.Accepted Then
            ReadCsv_Silent = "RESULT=ACCEPTED; records=" & Result.RecordCount & "; file=" & FilePath
        Else
            ReadCsv_Silent = "RESULT=REJECTED; findings=" & Result.FindingCount & _
                             "; first=" & FindingText(Findings(1)) & "; file=" & FilePath
        End If

End Function


Private Sub CheckCsvFile( _
    ByVal Kind As String)
'
'==============================================================================
'                                CheckCsvFile
'------------------------------------------------------------------------------
' PURPOSE
'   Ask for a CSV file, read it as a file of Kind and report whether it
'   follows the contract: the number of records of an accepted file, or
'   the number of findings and the first of them, up to
'   CSV_MESSAGE_FINDINGS (CsvCheckMessage). A cancelled dialog does
'   nothing; nothing is ever written to the workbook.
'
' ERROR POLICY
'   Reports a file that cannot be read in a message, or in the Immediate
'   window in silent mode; never raises.
'
' UPDATED
'   2026-11-04
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ChosenFile  As Variant          'Chosen file, or False when cancelled
    Dim Result      As CsvReadResult    'What the reader found
    Dim Values()    As String           'The records of an accepted file
    Dim Lines()     As Long             'Their lines
    Dim Findings()  As CsvFinding       'The findings

'------------------------------------------------------------------------------
' ASK AND READ
'------------------------------------------------------------------------------
        On Error GoTo Failed
        ChosenFile = Application.GetOpenFilename(FileFilter:="CSV files (*.csv), *.csv", _
                                                 Title:="Check " & Kind & " CSV file")
        If VarType(ChosenFile) = vbBoolean Then
            Exit Sub
        End If
        ReadCsvFile CStr(ChosenFile), Kind, Result, Values, Lines, Findings
        Notify CsvCheckMessage(CStr(ChosenFile), Result, Findings), _
               IIf(Result.Accepted, vbInformation, vbExclamation)
        Exit Sub

Failed:
        Notify Kind & " CSV file not checked: " & Err.Description, vbCritical

End Sub


Private Function CsvCheckMessage( _
    ByVal FilePath As String, _
    ByRef Result As CsvReadResult, _
    ByRef Findings() As CsvFinding, _
    Optional ByVal Note As String = "") _
    As String
'
'==============================================================================
'                               CsvCheckMessage
'------------------------------------------------------------------------------
' PURPOSE
'   The outcome of a CSV check or load for the user: an accepted file's
'   record count, or a rejected file's findings, the first
'   CSV_MESSAGE_FINDINGS of them with their line, record and column, then
'   how many more there are, Note, and the button that exports the import
'   specification, where the codes and rules are explained.
'
' INPUTS
'   Note: what was done with the file, in a sentence or two; "" for a
'   check, which writes nothing.
'
' METHOD
'   A long path is shown by its end, a long column name and a long line
'   are cut, and a finding is listed only while the whole message stays
'   within CSV_MESSAGE_LENGTH characters, so a message box never cuts the
'   count of the findings left or the pointer to the import
'   specification.
'
' UPDATED
'   2026-11-04
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Head        As String    'The path and the verdict
    Dim Body        As String    'The findings listed
    Dim Tail        As String    'Note and the pointer to the import specification
    Dim MoreText    As String    'The longest "... and N more." line
    Dim LineText    As String    'One finding's line
    Dim Shown       As Long      'Findings listed
    Dim f           As Long      'Finding counter

'------------------------------------------------------------------------------
' COMPOSE
'------------------------------------------------------------------------------
        Head = CutText(FilePath, CSV_MESSAGE_PATH, True) & vbLf & vbLf
        If Result.Accepted Then
            If Len(Note) = 0 Then
                Note = "It was only checked: nothing was written to the workbook."
            End If
            CsvCheckMessage = Head & "The file follows the " & Result.Kind & _
                              " CSV contract: " & Result.RecordCount & " records." & vbLf & Note
            Exit Function
        End If
        Head = Head & "The file does not follow the " & Result.Kind & " CSV contract: " & _
               Result.FindingCount & IIf(Result.FindingCount = 1, " finding.", " findings.") & vbLf
        If Len(Note) > 0 Then
            Tail = vbLf & vbLf & Note
        End If
        Tail = Tail & vbLf & vbLf & "Codes and rules: press Export import specifications on " & _
               "Import_NettingSets or Import_Trades."
        MoreText = vbLf & "... and " & Result.FindingCount & " more."
        For f = 1 To Result.FindingCount
            If Shown = CSV_MESSAGE_FINDINGS Then
                Exit For
            End If
            With Findings(f)
                LineText = "Line " & .LineNumber & ", record " & .RecordNumber & _
                           IIf(Len(.ColumnName) > 0, ", " & CutText(.ColumnName, CSV_MESSAGE_COLUMN, False), "") & _
                           ": " & .Code & ": " & .Message
            End With
            LineText = vbLf & CutText(LineText, CSV_MESSAGE_WIDTH, False)
            If Len(Head) + Len(Body) + Len(LineText) + Len(MoreText) + Len(Tail) > CSV_MESSAGE_LENGTH Then
                Exit For
            End If
            Body = Body & LineText
            Shown = Shown + 1
        Next f
        If Shown < Result.FindingCount Then
            Body = Body & vbLf & "... and " & (Result.FindingCount - Shown) & " more."
        End If
        CsvCheckMessage = Head & Body & Tail

End Function


Private Function CutText( _
    ByVal Text As String, _
    ByVal MaxLength As Long, _
    ByVal KeepEnd As Boolean) _
    As String
'
'==============================================================================
'                                   CutText
'------------------------------------------------------------------------------
' PURPOSE
'   Text as it is when it has at most MaxLength characters; otherwise cut to
'   MaxLength with "..." where the text was cut: at the start when KeepEnd is
'   True, as for a path, and at the end otherwise.
'
' UPDATED
'   2026-10-23
'==============================================================================
'
        If Len(Text) <= MaxLength Then
            CutText = Text
        ElseIf KeepEnd Then
            CutText = "..." & Right$(Text, MaxLength - 3)
        Else
            CutText = Left$(Text, MaxLength - 3) & "..."
        End If

End Function


'
'------------------------------------------------------------------------------
'
'                                 CSV STAGING
'
'------------------------------------------------------------------------------
'

Public Sub ImportNettingSetsCsv()
'
'==============================================================================
'                             ImportNettingSetsCsv
'------------------------------------------------------------------------------
' PURPOSE
'   Load a NettingSets CSV file chosen by the user into Import_NettingSets
'   (#208), for checking and correction before it replaces NettingSets.
'
' USAGE
'   NettingSets file on the K-SACCR Ribbon tab (Import CSV menu), the
'   Import CSV button on Import_NettingSets, or the Macros dialog.
'
' ERROR POLICY
'   As ImportCsv; never raises.
'
' UPDATED
'   2026-10-26
'==============================================================================
'
        ImportCsv "NettingSets"
        RibbonStateChanged

End Sub


Public Sub ImportTradesCsv()
'
'==============================================================================
'                               ImportTradesCsv
'------------------------------------------------------------------------------
' PURPOSE
'   Load a Trades CSV file chosen by the user into Import_Trades (#208),
'   for checking and correction before it replaces Trades.
'
' USAGE
'   Trades file on the K-SACCR Ribbon tab (Import CSV menu), the Import
'   CSV button on Import_Trades, or the Macros dialog.
'
' ERROR POLICY
'   As ImportCsv; never raises.
'
' UPDATED
'   2026-10-26
'==============================================================================
'
        ImportCsv "Trades"
        RibbonStateChanged

End Sub


Public Function ImportCsv_Silent( _
    ByVal Kind As String, _
    ByVal FilePath As String, _
    Optional ByVal ReplaceEdits As Boolean = False) _
    As String
'
'==============================================================================
'                               ImportCsv_Silent
'------------------------------------------------------------------------------
' PURPOSE
'   The staging load for automation and tests: read FilePath as a file of
'   Kind ("Trades" or "NettingSets") and load its records into the staging
'   sheet without asking. A staging sheet with changes since its last load
'   is replaced only when ReplaceEdits is True.
'
' RETURNS
'   "RESULT=LOADED; records=N; findings=K; file=..." when the records were
'   loaded, K being the findings about values to correct on the sheet;
'   "RESULT=REJECTED; findings=N; first=CODE|line|record|column; file=..."
'   when a structural finding stopped the load; "RESULT=REFUSED;
'   reason=...; file=..." when the sheet has changes and ReplaceEdits is
'   False. The sheet is changed only when the result is LOADED.
'
' ERROR POLICY
'   An unknown kind and a file that cannot be read raise ERR_CSV_READ. An
'   SA-CCR operation in progress raises ERR_RUN_ACTIVE, and a staging
'   sheet or name that is missing, and a value longer than a cell holds,
'   raise ERR_STAGING_WRITE, before the sheet changes; an error while the
'   sheet is written is raised after the sheet has been put back as it
'   was, as LoadStaging.
'
' UPDATED
'   2026-11-04
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Result      As CsvReadResult    'What the reader found
    Dim Values()    As String           'The records of a readable file
    Dim Lines()     As Long             'Their lines
    Dim Findings()  As CsvFinding       'The findings

'------------------------------------------------------------------------------
' READ AND LOAD
'------------------------------------------------------------------------------
        ReadCsvFile FilePath, Kind, Result, Values, Lines, Findings, 0, True
        If Not Result.Readable Then
            ImportCsv_Silent = "RESULT=REJECTED; findings=" & Result.FindingCount & _
                               "; first=" & FindingText(Findings(1)) & "; file=" & FilePath
            Exit Function
        End If
        If Not ReplaceEdits Then
            If StagingHasEdits(Kind) Then
                ImportCsv_Silent = "RESULT=REFUSED; reason=" & StagingSheetOf(Kind) & _
                                   " has changes since its last load; file=" & FilePath
                Exit Function
            End If
        End If
        LoadStaging Kind, FilePath, Result, Values, Lines
        ImportCsv_Silent = "RESULT=LOADED; records=" & Result.RecordCount & _
                           "; findings=" & Result.FindingCount & "; file=" & FilePath

End Function


Public Sub StagingEdited( _
    ByVal Kind As String, _
    ByVal Target As Range)
'
'==============================================================================
'                                StagingEdited
'------------------------------------------------------------------------------
' PURPOSE
'   Show EDITED in a staging sheet's status cell when a cell below its
'   header row changes. Called by the Worksheet_Change handler of
'   shImportNettingSets and shImportTrades. The status is for the user:
'   StagingHasEdits compares the block with the fingerprint of the last
'   load, so a change made with events off is found as well.
'
' INPUTS
'   Kind: "NettingSets" or "Trades". Target: the changed cells.
'
' ERROR POLICY
'   Contains every error: an event handler must not stop the user's edit.
'
' UPDATED
'   2026-10-26
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim HeaderCell  As Range    'The CSV line header cell
    Dim StatusCell  As Range    'The status cell

'------------------------------------------------------------------------------
' MARK
'------------------------------------------------------------------------------
    'Writing the status cell changes the sheet again; that change is above
    'the header row and returns at once.
        On Error GoTo Failed
        Set HeaderCell = NamedCell(StagingHeaderName(Kind))
        If HeaderCell Is Nothing Then
            Exit Sub
        End If
        If Not HeaderCell.Worksheet Is Target.Worksheet Then
            Exit Sub
        End If
        If Target.Row + Target.Rows.Count - 1 <= HeaderCell.Row Then
            Exit Sub
        End If
    'A sheet protected as saved, not for the user interface only, is
    'protected again so that the status can be written (#421); one
    'unprotected for maintenance stays so.
        Set StatusCell = NamedCell(StagingPrefixOf(Kind) & "_Status")
        If Not StatusCell Is Nothing Then
            If StatusCell.Worksheet.ProtectContents Then
                ProtectSheet StatusCell.Worksheet
            End If
            If CStr(StatusCell.Value2) <> STAGING_EDITED Then
                StatusCell.Value2 = STAGING_EDITED
            End If
        End If
        RibbonStateChanged

Failed:

End Sub


Public Function StagingFormulaCell( _
    ByVal Kind As String) _
    As String
'
'==============================================================================
'                              StagingFormulaCell
'------------------------------------------------------------------------------
' PURPOSE
'   The address of the first cell of a staging block that holds a formula,
'   or "" when none does (#208). A value pasted as a formula would be
'   calculated rather than kept as written, so staging validation and
'   commit (#209, #211) refuse such a block, whatever the events did; the
'   load checks its own writing with it, and StagingHasEdits counts one as
'   a change (#468).
'
' ERROR POLICY
'   Raises ERR_STAGING_WRITE when the staging sheet or its names are
'   missing.
'
' UPDATED
'   2026-10-29
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Sheet       As Worksheet    'The staging sheet
    Dim HeaderRow   As Long         'Its header row
    Dim LastRow     As Long         'Last row of its used range
    Dim Block       As Range        'The rows below the header
    Dim Found       As Range        'Its formula cells; Nothing for none

'------------------------------------------------------------------------------
' SEARCH
'------------------------------------------------------------------------------
    'HasFormula is False when no cell has one and Null when some do;
    'SpecialCells raises when it finds none.
        HeaderRow = StagingHeaderRow(Kind, Sheet)
        LastRow = UsedLastRow(Sheet)
        If LastRow <= HeaderRow Then
            Exit Function
        End If
        Set Block = TableRange(Sheet, HeaderRow + 1, 1, LastRow, StagingWidth(Kind))
        If Not IsNull(Block.HasFormula) Then
            If Not Block.HasFormula Then
                Exit Function
            End If
        End If
        On Error Resume Next
        Set Found = Block.SpecialCells(xlCellTypeFormulas)
        On Error GoTo 0
        If Not Found Is Nothing Then
            StagingFormulaCell = Found.Cells(1, 1).Address(False, False)
        End If

End Function


Private Sub ImportCsv( _
    ByVal Kind As String)
'
'==============================================================================
'                                  ImportCsv
'------------------------------------------------------------------------------
' PURPOSE
'   Ask for a CSV file, read it as a file of Kind and load its records into
'   the staging sheet as text, with the CSV line of each, then show the
'   outcome and the sheet. A structural finding loads nothing; findings
'   about values are listed and stay to be corrected on the sheet. When
'   the sheet has changes since its last load, the user is asked before
'   they are replaced. A cancelled dialog or a No changes nothing.
'
' ERROR POLICY
'   Reports a file that cannot be read or loaded in a message, or in the
'   Immediate window in silent mode; never raises. The staging sheet is
'   put back as it was when its writing fails (LoadStaging).
'
' UPDATED
'   2026-10-26
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim SheetName   As String           'The staging sheet
    Dim ChosenFile  As Variant          'Chosen file, or False when cancelled
    Dim Result      As CsvReadResult    'What the reader found
    Dim Values()    As String           'The records of a readable file
    Dim Lines()     As Long             'Their lines
    Dim Findings()  As CsvFinding       'The findings
    Dim Note        As String           'What was loaded

'------------------------------------------------------------------------------
' ASK, READ AND LOAD
'------------------------------------------------------------------------------
        On Error GoTo Failed
        SheetName = StagingSheetOf(Kind)
        ChosenFile = Application.GetOpenFilename(FileFilter:="CSV files (*.csv), *.csv", _
                                                 Title:="Import " & Kind & " CSV file into " & SheetName)
        If VarType(ChosenFile) = vbBoolean Then
            Exit Sub
        End If
        ReadCsvFile CStr(ChosenFile), Kind, Result, Values, Lines, Findings, 0, True
        If Not Result.Readable Then
            Notify CsvCheckMessage(CStr(ChosenFile), Result, Findings, _
                                   "Nothing was loaded: " & SheetName & " is unchanged."), vbExclamation
            Exit Sub
        End If
        If StagingHasEdits(Kind) Then
            If Not ReplaceConfirmed(Kind, CStr(ChosenFile)) Then
                Exit Sub
            End If
        End If
        LoadStaging Kind, CStr(ChosenFile), Result, Values, Lines
        If Result.Accepted Then
            Note = "Loaded into " & SheetName & " as text." & vbLf & "Not validated yet; " & Kind & " is unchanged."
        Else
            Note = Result.RecordCount & " records loaded into " & SheetName & " as text: correct the findings there." & _
                   vbLf & Kind & " is unchanged."
        End If
        ThisWorkbook.Worksheets(SheetName).Activate
        Notify CsvCheckMessage(CStr(ChosenFile), Result, Findings, Note), _
               IIf(Result.Accepted, vbInformation, vbExclamation)
        Exit Sub

Failed:
        Notify Kind & " CSV import: " & Err.Description, vbCritical

End Sub


Private Function ReplaceConfirmed( _
    ByVal Kind As String, _
    ByVal NewFile As String) _
    As Boolean
'
'==============================================================================
'                               ReplaceConfirmed
'------------------------------------------------------------------------------
' PURPOSE
'   Ask whether the changes on a staging sheet may be replaced by the
'   records of NewFile. No is the default answer; silent mode never
'   replaces them.
'
' UPDATED
'   2026-10-26
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim LoadedFrom  As String    'The file of the last load; "" for none
    Dim Question    As String    'The question

'------------------------------------------------------------------------------
' ASK
'------------------------------------------------------------------------------
        If gSilent Then
            Exit Function
        End If
        LoadedFrom = CStr(StagingMetaCell(Kind, 1).Value2)
        If Len(LoadedFrom) > 0 Then
            Question = StagingSheetOf(Kind) & " has changes since it was loaded from " & _
                       CutText(LoadedFrom, CSV_MESSAGE_PATH, True) & "."
        Else
            Question = StagingSheetOf(Kind) & " has rows that were not loaded from a file."
        End If
        Question = Question & vbLf & vbLf & "Replace them with the records of " & _
                   CutText(NewFile, CSV_MESSAGE_PATH, True) & "?"
        ReplaceConfirmed = (MsgBox(Question, vbYesNo + vbExclamation + vbDefaultButton2, MSG_TITLE) = vbYes)

End Function


Private Sub LoadStaging( _
    ByVal Kind As String, _
    ByVal FilePath As String, _
    ByRef Result As CsvReadResult, _
    ByRef Values() As String, _
    ByRef Lines() As Long)
'
'==============================================================================
'                                 LoadStaging
'------------------------------------------------------------------------------
' PURPOSE
'   Replace the block below a staging sheet's header row with the records
'   of a readable file: the CSV line of each in the first column, its
'   values as text after it. Then record the file, the time, the record
'   count and LOADED in the metadata cells, and the fingerprint of the
'   block in STAGED_*_NAME.
'
' METHOD
'   The new block is built in memory before the sheet is touched; a value
'   longer than a cell holds stops the load there. The old block, its
'   formulas with their formats and the metadata are kept, then cleared
'   and written over; on an error they are written back. The data cells
'   are formatted as text before their values are written, so that nothing
'   is converted, and the block is then checked for formulas. The
'   fingerprint is stored last, so a load that fails leaves the earlier
'   one.
'
' STATE OWNERSHIP
'   Captures calculation mode, events and screen updating, switches events
'   off so that the change handler does not mark the load as an edit, and
'   restores them on every path. Sets mRunning while it writes, so that no
'   SA-CCR operation starts on a half-written sheet, and refuses to start
'   while one runs.
'
' ERROR POLICY
'   Raises ERR_RUN_ACTIVE while an operation runs, and ERR_STAGING_WRITE
'   when the sheet or its names are missing or a value is longer than a
'   cell holds (EXCEL_CELL_CHARS), before anything is changed.
'   Any later error is raised after the old block and metadata have been
'   written back; a failure to write them back, or to restore a setting,
'   is added to its description. A setting not restored after a
'   successful load raises ERR_CLEANUP_FAILED.
'
' TEST SEAM
'   gTestFault = "staging-write" raises ERR_INJECTED_FAULT once the new
'   block has been written, before the metadata.
'
' UPDATED
'   2026-10-26
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Sheet               As Worksheet       'The staging sheet
    Dim FirstRow            As Long            'First row below its header
    Dim Width               As Long            'CSV line column and the contract's columns
    Dim MetaCells(1 To 4)   As Range           'File, load time, record count, status
    Dim NewBlock()          As Variant         'The records as they will be written
    Dim OldBlock            As Variant         'The block before the load
    Dim OldLast             As Long            'Its last row with a value; below FirstRow for none
    Dim OldMeta(1 To 4)     As Variant         'The metadata before the load
    Dim FormulaAddresses()  As String          'Cells of the old block that hold a formula
    Dim FormulaTexts()      As String          'Their formulas
    Dim FormulaFormats()    As String          'Their number formats
    Dim FormulaCount        As Long            'How many there are
    Dim Spec                As CsvSpec         'The file's contract
    Dim Columns()           As CsvColumn       'Its columns, for a message
    Dim Rules()             As CsvRule         'Its rules
    Dim Cleared             As Boolean         'The old block has been cleared
    Dim FormulaAt           As String          'A cell that became a formula
    Dim SavedCalculation    As XlCalculation   'Calculation mode before the load
    Dim SavedEvents         As Boolean         'EnableEvents before the load
    Dim SavedScreen         As Boolean         'ScreenUpdating before the load
    Dim StateCaptured       As Boolean         'True once all three were read
    Dim ErrorNumber         As Long            'Primary error number; 0 for none
    Dim ErrorSource         As String          'Primary error source
    Dim ErrorDescription    As String          'Primary error description
    Dim RestoreDetail       As String          'Why the old block was not written back
    Dim CleanupDetails      As String          'Settings not restored
    Dim r                   As Long            'Record
    Dim c                   As Long            'Contract column
    Dim m                   As Long            'Metadata cell

'------------------------------------------------------------------------------
' PREPARE
'------------------------------------------------------------------------------
        If mRunning Then
            Err.Raise ERR_RUN_ACTIVE, "M_Main.LoadStaging", "An SA-CCR operation is already running."
        End If
        FirstRow = StagingHeaderRow(Kind, Sheet) + 1
        Width = StagingWidth(Kind)
        For m = 1 To 4
            Set MetaCells(m) = StagingMetaCell(Kind, m)
        Next m
        If Result.RecordCount > 0 Then
            If Not CsvSpecOf(Kind, Spec, Columns, Rules) Then
                Err.Raise ERR_CSV_READ, "M_Main.LoadStaging", "unknown CSV kind '" & Kind & "'"
            End If
            ReDim NewBlock(1 To Result.RecordCount, 1 To Width)
            For r = 1 To Result.RecordCount
                NewBlock(r, 1) = Lines(r)
                For c = 1 To Width - 1
                    If Len(Values(r, c)) > EXCEL_CELL_CHARS Then
                        Err.Raise ERR_STAGING_WRITE, "M_Main.LoadStaging", _
                                  "The " & Columns(c).ColumnName & " value on line " & Lines(r) & " has " & _
                                  Len(Values(r, c)) & " characters; a cell holds " & EXCEL_CELL_CHARS & _
                                  ". Nothing was loaded."
                    End If
                    NewBlock(r, c + 1) = Values(r, c)
                Next c
            Next r
        End If
        mRunning = True
        ProtectWorkbook
        On Error GoTo Failed
        SavedCalculation = Application.Calculation
        SavedEvents = Application.EnableEvents
        SavedScreen = Application.ScreenUpdating
        StateCaptured = True
        Application.ScreenUpdating = False
        Application.EnableEvents = False
        Application.Calculation = xlCalculationManual

'------------------------------------------------------------------------------
' KEEP THE OLD BLOCK, THEN WRITE THE NEW ONE
'------------------------------------------------------------------------------
        OldLast = LastStagingRow(Sheet, FirstRow, Width)
        If OldLast >= FirstRow Then
            OldBlock = TableRange(Sheet, FirstRow, 1, OldLast, Width).Value2
            KeepFormulas TableRange(Sheet, FirstRow, 1, OldLast, Width), FormulaAddresses, FormulaTexts, _
                         FormulaFormats, FormulaCount
        End If
        For m = 1 To 4
            OldMeta(m) = MetaCells(m).Value2
        Next m
        Cleared = True
        ClearStagingBlock Sheet, FirstRow, Width
        If Result.RecordCount > 0 Then
            WriteStagingBlock Sheet, FirstRow, Width, NewBlock
        End If
        If gTestFault = "staging-write" Then
            Err.Raise ERR_INJECTED_FAULT, "M_Main.LoadStaging", "Injected staging write failure."
        End If
        FormulaAt = StagingFormulaCell(Kind)
        If Len(FormulaAt) > 0 Then
            Err.Raise ERR_STAGING_WRITE, "M_Main.LoadStaging", _
                      "A value was written as a formula in " & Sheet.Name & "!" & FormulaAt & "."
        End If
        MetaCells(1).Value2 = FilePath
        MetaCells(2).Value2 = Format$(Now, "yyyy-mm-dd hh:mm:ss")
        MetaCells(3).Value2 = CStr(Result.RecordCount)
        MetaCells(4).Value2 = STAGING_LOADED
        ThisWorkbook.Names.Add Name:=StagedNameOf(Kind), RefersTo:="=""" & StagingFingerprint(Kind) & """", _
                               Visible:=False

'------------------------------------------------------------------------------
' PUT BACK ON FAILURE, RESTORE EXCEL
'------------------------------------------------------------------------------
CleanUp:
        On Error GoTo 0
        If ErrorNumber <> 0 And Cleared Then
            RestoreDetail = RestoreStaging(Sheet, FirstRow, Width, OldLast, OldBlock, MetaCells, OldMeta, _
                                           FormulaAddresses, FormulaTexts, FormulaFormats, FormulaCount)
        End If
        If StateCaptured Then
            CleanupDetails = RestoreSettings(SavedCalculation, SavedEvents, SavedScreen)
        End If
        mRunning = False
        If ErrorNumber <> 0 Then
            If Len(RestoreDetail) > 0 Then
                ErrorDescription = ErrorDescription & " " & Sheet.Name & " could not be put back: " & RestoreDetail
            End If
            If Len(CleanupDetails) > 0 Then
                ErrorDescription = ErrorDescription & " Excel settings not restored: " & CleanupDetails
            End If
            Err.Raise ErrorNumber, ErrorSource, ErrorDescription
        End If
        If Len(CleanupDetails) > 0 Then
            Err.Raise ERR_CLEANUP_FAILED, "M_Main.LoadStaging", _
                      "The records were loaded, but Excel settings were not restored: " & CleanupDetails
        End If
        Exit Sub

'------------------------------------------------------------------------------
' HANDLE ERROR
'------------------------------------------------------------------------------
Failed:
        ErrorNumber = Err.Number
        ErrorSource = Err.Source
        ErrorDescription = Err.Description
        Resume CleanUp

End Sub


Private Function RestoreStaging( _
    ByVal Sheet As Worksheet, _
    ByVal FirstRow As Long, _
    ByVal Width As Long, _
    ByVal OldLast As Long, _
    ByRef OldBlock As Variant, _
    ByRef MetaCells() As Range, _
    ByRef OldMeta() As Variant, _
    ByRef FormulaAddresses() As String, _
    ByRef FormulaTexts() As String, _
    ByRef FormulaFormats() As String, _
    ByVal FormulaCount As Long) _
    As String
'
'==============================================================================
'                                RestoreStaging
'------------------------------------------------------------------------------
' PURPOSE
'   Write back the block and the metadata a failed load kept, with the
'   staging formats, then the formulas the block held, each with its own
'   number format, so that a formula pasted there is still found by
'   StagingFormulaCell.
'
' RETURNS
'   Empty on success; otherwise what failed.
'
' ERROR POLICY
'   Contains its own error and returns it as text.
'
' UPDATED
'   2026-10-26
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim m   As Long    'Metadata cell
    Dim i   As Long    'Formula cell

'------------------------------------------------------------------------------
' WRITE BACK
'------------------------------------------------------------------------------
        On Error GoTo Failed
        ClearStagingBlock Sheet, FirstRow, Width
        If OldLast >= FirstRow Then
            WriteStagingBlock Sheet, FirstRow, Width, OldBlock
        End If
        For i = 1 To FormulaCount
            With Sheet.Range(FormulaAddresses(i))
                .NumberFormat = FormulaFormats(i)
                .Formula = FormulaTexts(i)
            End With
        Next i
        For m = 1 To 4
            MetaCells(m).Value2 = OldMeta(m)
        Next m
        Exit Function

'------------------------------------------------------------------------------
' REPORT FAILURE
'------------------------------------------------------------------------------
Failed:
        RestoreStaging = CStr(Err.Number) & " / " & Err.Description

End Function


Private Sub KeepFormulas( _
    ByVal Block As Range, _
    ByRef FormulaAddresses() As String, _
    ByRef FormulaTexts() As String, _
    ByRef FormulaFormats() As String, _
    ByRef FormulaCount As Long)
'
'==============================================================================
'                                 KeepFormulas
'------------------------------------------------------------------------------
' PURPOSE
'   Keep the address, formula and number format of every cell of Block
'   that holds a formula, for RestoreStaging: a values-only copy would put
'   a formula back as its result.
'
' UPDATED
'   2026-10-26
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Found     As Range    'The formula cells; Nothing for none
    Dim OneCell   As Range    'One of them

'------------------------------------------------------------------------------
' KEEP
'------------------------------------------------------------------------------
    'HasFormula is False when no cell has one; SpecialCells raises when it
    'finds none.
        FormulaCount = 0
        If Not IsNull(Block.HasFormula) Then
            If Not Block.HasFormula Then
                Exit Sub
            End If
        End If
        On Error Resume Next
        Set Found = Block.SpecialCells(xlCellTypeFormulas)
        On Error GoTo 0
        If Found Is Nothing Then
            Exit Sub
        End If
        ReDim FormulaAddresses(1 To Found.Count)
        ReDim FormulaTexts(1 To Found.Count)
        ReDim FormulaFormats(1 To Found.Count)
        For Each OneCell In Found.Cells
            FormulaCount = FormulaCount + 1
            FormulaAddresses(FormulaCount) = OneCell.Address
            FormulaTexts(FormulaCount) = OneCell.Formula
            FormulaFormats(FormulaCount) = OneCell.NumberFormat
        Next OneCell

End Sub


Private Sub ClearStagingBlock( _
    ByVal Sheet As Worksheet, _
    ByVal FirstRow As Long, _
    ByVal Width As Long)
'
'==============================================================================
'                              ClearStagingBlock
'------------------------------------------------------------------------------
' PURPOSE
'   Clear the rows of a staging block, values and formats, down to the
'   last row of the used range, and leave its data columns formatted as
'   text and unlocked, as the template has them, so that a value typed
'   there later is kept as written and can be typed on the protected
'   sheet (#420); the CSV line column stays locked.
'
' UPDATED
'   2026-11-03
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim LastRow   As Long    'Last row of the used range

'------------------------------------------------------------------------------
' CLEAR
'------------------------------------------------------------------------------
        LastRow = UsedLastRow(Sheet)
        If LastRow < FirstRow Then
            Exit Sub
        End If
        With TableRange(Sheet, FirstRow, 1, LastRow, Width)
            .Clear
            .Font.Size = STAGING_FONT_SIZE
        End With
        With TableRange(Sheet, FirstRow, 2, LastRow, Width)
            .NumberFormat = "@"
            .Locked = False
        End With

End Sub


Private Sub WriteStagingBlock( _
    ByVal Sheet As Worksheet, _
    ByVal FirstRow As Long, _
    ByVal Width As Long, _
    ByRef Block As Variant)
'
'==============================================================================
'                              WriteStagingBlock
'------------------------------------------------------------------------------
' PURPOSE
'   Format the rows of Block on a staging sheet and write it: the CSV line
'   column as a number on a grey fill, locked, the data columns as text on
'   the input fill, unlocked (#420), with the input borders. The text
'   format is set before the values are written, so that Excel converts
'   none of them.
'
' UPDATED
'   2026-11-03
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim LastRow   As Long    'Last row of the block

'------------------------------------------------------------------------------
' FORMAT AND WRITE
'------------------------------------------------------------------------------
        LastRow = FirstRow + UBound(Block, 1) - LBound(Block, 1)
        With TableRange(Sheet, FirstRow, 1, LastRow, 1)
            .NumberFormat = "0"
            .Interior.Color = STAGING_LINE_FILL
            .Locked = True
        End With
        With TableRange(Sheet, FirstRow, 2, LastRow, Width)
            .NumberFormat = "@"
            .Interior.Color = STAGING_VALUE_FILL
            .Locked = False
        End With
        With TableRange(Sheet, FirstRow, 1, LastRow, Width)
            .Font.Size = STAGING_FONT_SIZE
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
            .Borders.Color = STAGING_BORDER
            .Value2 = Block
        End With

End Sub


Private Function LastInputRow( _
    ByVal Sheet As Worksheet, _
    ByVal FirstRow As Long, _
    ByVal Width As Long) _
    As Long
'
'==============================================================================
'                                 LastInputRow
'------------------------------------------------------------------------------
' PURPOSE
'   The last row of an input block with a value or a formula in one of
'   its columns, or FirstRow - 1 when none has either (#512).
'   LastStagingRow reads values only, so a formula that returns empty
'   text below the data would escape a commit: it would be neither kept,
'   cleared nor read back, and could show a value once Excel recalculates.
'
' UPDATED
'   2026-11-03
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim LastRow   As Long     'Last row of the used range
    Dim Found     As Range    'The formula cells; Nothing for none
    Dim Area      As Range    'One area of them
    Dim AreaLast  As Long     'Its last row

'------------------------------------------------------------------------------
' FIND
'------------------------------------------------------------------------------
    'HasFormula is False when no cell has one and Null when some do;
    'SpecialCells raises when it finds none.
        LastInputRow = LastStagingRow(Sheet, FirstRow, Width)
        LastRow = UsedLastRow(Sheet)
        If LastRow < FirstRow Then
            Exit Function
        End If
        With TableRange(Sheet, FirstRow, 1, LastRow, Width)
            If Not IsNull(.HasFormula) Then
                If Not .HasFormula Then
                    Exit Function
                End If
            End If
            On Error Resume Next
            Set Found = .SpecialCells(xlCellTypeFormulas)
            On Error GoTo 0
        End With
        If Found Is Nothing Then
            Exit Function
        End If
        For Each Area In Found.Areas
            AreaLast = Area.Row + Area.Rows.Count - 1
            If AreaLast > LastInputRow Then
                LastInputRow = AreaLast
            End If
        Next Area

End Function


Private Function LastStagingRow( _
    ByVal Sheet As Worksheet, _
    ByVal FirstRow As Long, _
    ByVal Width As Long) _
    As Long
'
'==============================================================================
'                                LastStagingRow
'------------------------------------------------------------------------------
' PURPOSE
'   The last row of a staging block with a value in any of its columns, or
'   FirstRow - 1 when it has none.
'
' UPDATED
'   2026-10-26
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim LastRow   As Long       'Last row of the used range
    Dim Grid      As Variant    'The block's values
    Dim r         As Long       'Row of Grid
    Dim c         As Long       'Column of Grid

'------------------------------------------------------------------------------
' SCAN FROM THE BOTTOM
'------------------------------------------------------------------------------
        LastStagingRow = FirstRow - 1
        LastRow = UsedLastRow(Sheet)
        If LastRow < FirstRow Then
            Exit Function
        End If
        Grid = TableRange(Sheet, FirstRow, 1, LastRow, Width).Value2
        For r = UBound(Grid, 1) To 1 Step -1
            For c = 1 To Width
                If Not IsEmpty(Grid(r, c)) Then
                    If Len(CStr(Grid(r, c))) > 0 Then
                        LastStagingRow = FirstRow + r - 1
                        Exit Function
                    End If
                End If
            Next c
        Next r

End Function


Private Function StagingHasEdits( _
    ByVal Kind As String) _
    As Boolean
'
'==============================================================================
'                               StagingHasEdits
'------------------------------------------------------------------------------
' PURPOSE
'   Tell whether a staging block differs from the one the last load wrote,
'   or, before any load, whether it holds anything: by its fingerprint,
'   whatever its status cell shows. A load never writes a formula, so a
'   formula in the block is a change, also when its value is the text it
'   replaced, which the fingerprint of the values cannot tell (#468).
'
' ERROR POLICY
'   As StagingFingerprint and StagingFormulaCell.
'
' UPDATED
'   2026-10-29
'==============================================================================
'
        StagingHasEdits = (StagingFingerprint(Kind) <> StoredText(StagedNameOf(Kind)))
        If Not StagingHasEdits Then
            StagingHasEdits = (Len(StagingFormulaCell(Kind)) > 0)
        End If

End Function


Private Function StagingFingerprint( _
    ByVal Kind As String) _
    As String
'
'==============================================================================
'                              StagingFingerprint
'------------------------------------------------------------------------------
' PURPOSE
'   Fingerprint a staging block, the CSV line column included, from the
'   row below its header to its last row with a value (LastStagingRow);
'   "" for an empty block. A cell holding an empty text reads as a blank
'   cell, as LastStagingRow and ReadStaging read it: an import writes an
'   empty CSV field as an empty text, and clearing that cell must not
'   read as an edit.
'
' ERROR POLICY
'   Raises ERR_STAGING_WRITE when the staging sheet or its names are
'   missing.
'
' UPDATED
'   2026-11-02
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Sheet       As Worksheet    'The staging sheet
    Dim FirstRow    As Long         'First row below its header
    Dim LastRow     As Long         'Last row with a value
    Dim Width       As Long         'Columns, the CSV line column included
    Dim Grid        As Variant      'The block's values
    Dim Pieces()    As String       'One text per cell
    Dim r           As Long         'Row of Grid
    Dim c           As Long         'Column of Grid
    Dim k           As Long         'Position in Pieces

'------------------------------------------------------------------------------
' FINGERPRINT
'------------------------------------------------------------------------------
    'Chr$(31) separates cells; the width fixes where rows end.
        FirstRow = StagingHeaderRow(Kind, Sheet) + 1
        Width = StagingWidth(Kind)
        LastRow = LastStagingRow(Sheet, FirstRow, Width)
        If LastRow < FirstRow Then
            Exit Function
        End If
        Grid = TableRange(Sheet, FirstRow, 1, LastRow, Width).Value2
        ReDim Pieces(1 To (LastRow - FirstRow + 1) * Width)
        For r = 1 To LastRow - FirstRow + 1
            For c = 1 To Width
                k = k + 1
                Pieces(k) = StagingToken(Grid(r, c))
            Next c
        Next r
        StagingFingerprint = TextHash(Join(Pieces, Chr$(31)))

End Function


Private Function StagingToken( _
    ByVal Content As Variant) _
    As String
'
'==============================================================================
'                                 StagingToken
'------------------------------------------------------------------------------
' PURPOSE
'   One staging cell as text for StagingFingerprint, with its type: ""
'   for a blank cell or an empty text, T for a text, B for a logical, E
'   for an error, N for a number.
'
' UPDATED
'   2026-11-02
'==============================================================================
'
        Select Case VarType(Content)
            Case vbEmpty
                StagingToken = ""
            Case vbString
                If Len(Content) > 0 Then
                    StagingToken = "T" & Content
                End If
            Case vbBoolean
                StagingToken = IIf(Content, "B1", "B0")
            Case vbError
                StagingToken = "E" & CStr(Content)
            Case Else
                StagingToken = "N" & CStr(Content)
        End Select

End Function


Private Function StagingHeaderRow( _
    ByVal Kind As String, _
    ByRef Sheet As Worksheet) _
    As Long
'
'==============================================================================
'                               StagingHeaderRow
'------------------------------------------------------------------------------
' PURPOSE
'   The header row of a staging sheet, from its header name, and the
'   sheet.
'
' ERROR POLICY
'   Raises ERR_STAGING_WRITE when the name is missing or not on the
'   staging sheet's CSV line column.
'
' UPDATED
'   2026-10-26
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim HeaderCell   As Range    'The CSV line header cell

'------------------------------------------------------------------------------
' LOCATE
'------------------------------------------------------------------------------
        Set HeaderCell = NamedCell(StagingHeaderName(Kind))
        If HeaderCell Is Nothing Then
            Err.Raise ERR_STAGING_WRITE, "M_Main.StagingHeaderRow", _
                      "The name " & StagingHeaderName(Kind) & " is missing: build the workbook from the current template."
        End If
        If HeaderCell.Worksheet.Name <> StagingSheetOf(Kind) Or HeaderCell.Column <> 1 + COL_OFFSET Then
            Err.Raise ERR_STAGING_WRITE, "M_Main.StagingHeaderRow", _
                      "The name " & StagingHeaderName(Kind) & " is not in column C of " & StagingSheetOf(Kind) & "."
        End If
        Set Sheet = HeaderCell.Worksheet
        StagingHeaderRow = HeaderCell.Row

End Function


Private Function StagingMetaCell( _
    ByVal Kind As String, _
    ByVal Index As Long) _
    As Range
'
'==============================================================================
'                               StagingMetaCell
'------------------------------------------------------------------------------
' PURPOSE
'   A metadata cell of a staging sheet: 1 the file, 2 the load time, 3 the
'   record count, 4 the status.
'
' ERROR POLICY
'   Raises ERR_STAGING_WRITE when its name is missing.
'
' UPDATED
'   2026-10-26
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim NameText   As String    'The cell's name

'------------------------------------------------------------------------------
' LOCATE
'------------------------------------------------------------------------------
        NameText = StagingPrefixOf(Kind) & "_" & Choose(Index, "File", "LoadedAt", "Records", "Status")
        Set StagingMetaCell = NamedCell(NameText)
        If StagingMetaCell Is Nothing Then
            Err.Raise ERR_STAGING_WRITE, "M_Main.StagingMetaCell", _
                      "The name " & NameText & " is missing: build the workbook from the current template."
        End If

End Function


Private Function StagingWidth( _
    ByVal Kind As String) _
    As Long
'
'==============================================================================
'                                 StagingWidth
'------------------------------------------------------------------------------
' PURPOSE
'   The columns of a staging block: the CSV line and the contract's.
'
' ERROR POLICY
'   Raises ERR_CSV_READ for an unknown kind.
'
' UPDATED
'   2026-10-26
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Spec        As CsvSpec      'The file's contract
    Dim Columns()   As CsvColumn    'Its columns
    Dim Rules()     As CsvRule      'Its rules

'------------------------------------------------------------------------------
' COUNT
'------------------------------------------------------------------------------
        If Not CsvSpecOf(Kind, Spec, Columns, Rules) Then
            Err.Raise ERR_CSV_READ, "M_Main.StagingWidth", "unknown CSV kind '" & Kind & "'"
        End If
        StagingWidth = Spec.ColumnCount + 1

End Function


Private Function StagingSheetOf( _
    ByVal Kind As String) _
    As String
'
'==============================================================================
' PURPOSE
'   The staging sheet of a CSV kind.
'==============================================================================
'
        StagingSheetOf = IIf(IsNettingSetsKind(Kind), SH_IMPORT_NS, SH_IMPORT_TR)

End Function


Private Function StagingHeaderName( _
    ByVal Kind As String) _
    As String
'
'==============================================================================
' PURPOSE
'   The header-row name of a staging sheet.
'==============================================================================
'
        StagingHeaderName = IIf(IsNettingSetsKind(Kind), IMPORT_NS_HEADER_NAME, IMPORT_TR_HEADER_NAME)

End Function


Private Function StagingPrefixOf( _
    ByVal Kind As String) _
    As String
'
'==============================================================================
' PURPOSE
'   The prefix of a staging sheet's metadata names.
'==============================================================================
'
        StagingPrefixOf = IIf(IsNettingSetsKind(Kind), IMPORT_NS_PREFIX, IMPORT_TR_PREFIX)

End Function


Private Function StagedNameOf( _
    ByVal Kind As String) _
    As String
'
'==============================================================================
' PURPOSE
'   The hidden name holding the fingerprint of a staging sheet's last load.
'==============================================================================
'
        StagedNameOf = IIf(IsNettingSetsKind(Kind), STAGED_NS_NAME, STAGED_TR_NAME)

End Function


Private Function IsNettingSetsKind( _
    ByVal Kind As String) _
    As Boolean
'
'==============================================================================
' PURPOSE
'   True for "NettingSets", False for "Trades"; raises ERR_CSV_READ for any
'   other kind.
'==============================================================================
'
        If Kind <> "NettingSets" And Kind <> "Trades" Then
            Err.Raise ERR_CSV_READ, "M_Main.IsNettingSetsKind", "unknown CSV kind '" & Kind & "'"
        End If
        IsNettingSetsKind = (Kind = "NettingSets")

End Function


'
'------------------------------------------------------------------------------
'
'                              STAGING VALIDATION
'
'------------------------------------------------------------------------------
'

Public Sub ValidateStaging()
'
'==============================================================================
'                               ValidateStaging
'------------------------------------------------------------------------------
' PURPOSE
'   Check the rows of both staging sheets (#209) and list the findings on
'   Checks with the origin IMPORT, each linked to its cell, which is marked
'   on its staging sheet. Nothing is copied to NettingSets or Trades.
'
' USAGE
'   Validate staging on the K-SACCR Ribbon tab, the Validate staging
'   button on Import_NettingSets and Import_Trades, or the Macros dialog.
'
' ERROR POLICY
'   Reports an error in a message, or in the Immediate window in silent
'   mode; never raises.
'
' UPDATED
'   2026-10-28
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ResultLine  As String    'What the validation found
    Dim Message     As String    'Its message

'------------------------------------------------------------------------------
' VALIDATE AND REPORT
'------------------------------------------------------------------------------
        On Error GoTo Failed
        ResultLine = CheckStaging(Message)
        RibbonStateChanged
        If Left$(ResultLine, Len("RESULT=REFUSED")) <> "RESULT=REFUSED" And Not gSilent Then
            GetSheet(SH_CHECKS).Activate
        End If
        Notify Message, IIf(Left$(ResultLine, Len("RESULT=VALID")) = "RESULT=VALID", vbInformation, vbExclamation)
        Exit Sub

Failed:
        Notify "Staging validation: " & Err.Description, vbCritical

End Sub


Public Function ValidateStaging_Silent() As String
'
'==============================================================================
'                            ValidateStaging_Silent
'------------------------------------------------------------------------------
' PURPOSE
'   The staging validation for automation and tests, as ValidateStaging
'   without a message.
'
' RETURNS
'   "RESULT=VALID; errors=0; nettingsets=N; trades=M" when no row has an
'   error, N and M being the rows of Import_NettingSets and Import_Trades;
'   "RESULT=INVALID; errors=K; nettingsets=N; trades=M; first=CODE|Sheet!A1"
'   otherwise, the first finding with its cell, or "-" for none; and
'   "RESULT=REFUSED; reason=..." when nothing was checked: nothing staged,
'   or a formula in a staging block.
'
' ERROR POLICY
'   Raises the error of CheckStaging, ERR_RUN_ACTIVE while an operation
'   runs among them.
'
' UPDATED
'   2026-10-28
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Message     As String    'The message, not shown

'------------------------------------------------------------------------------
' VALIDATE
'------------------------------------------------------------------------------
        ValidateStaging_Silent = CheckStaging(Message)

End Function


Private Function CheckStaging( _
    ByRef Message As String) _
    As String
'
'==============================================================================
'                                 CheckStaging
'------------------------------------------------------------------------------
' PURPOSE
'   Validate the portfolio a commit would make (#209): the staged rows of
'   each staging sheet that holds any, and the live sheet where none are
'   staged. Then rewrite Checks with the findings, mark their cells and
'   set each staging sheet's status to VALID or INVALID.
'
' METHOD
'   0. The workbook is recalculated first, as a run recalculates it, so
'      that a formula input is judged by its current result (#519).
'   1. Each staging block is read as text, as CheckStagedRecords reads
'      it: each field, the rules between fields and the keys.
'   2. The netting sets are those of Import_NettingSets when it holds
'      rows, or else those of NettingSets. CheckPortfolio checks the
'      staged trades against them; when only netting sets are staged, it
'      checks the netting set of each leg on Trades against them, and a
'      finding is on the Trades cell.
'   3. What a run reads first, the layout, Params and its two tables, is
'      read as a run reads it (CORE_Engine.LoadRunContext). CheckContext
'      judges the staged trades against it: the reporting date, the FX
'      table, the factor table and the regime of their netting set; with
'      only netting sets staged, the factor and regime rules on each leg
'      of Trades. When a run would stop there, nothing that depends on it
'      is judged and that finding, on the cell the run names, keeps every
'      staging sheet that holds rows from VALID.
'   4. Checks is rewritten: a first INFO line with the counts, then one
'      ERROR line per finding with the sheet, the row, a link to the cell
'      in Reference, the code, column, message and CSV line in Message,
'      and the origin IMPORT; at most STAGING_LISTED_MAX of them.
'   5. The marks of the last validation are taken off both staging
'      sheets, then each staged cell with a finding is marked. Live cells
'      are never marked or changed.
'   6. The status of each staging sheet that holds rows is set last, and
'      a validation without a finding stores the fingerprint of what it
'      judged for StagingValidationCurrent; any other empties it.
'   The fingerprints and counters of the last run are left as they are. Checks no
'   longer shows what the run wrote, so current outputs read as EDITED
'   and the evidence export refuses until the next run (#244).
'
' RETURNS
'   The result line of ValidateStaging_Silent, and Message for the user.
'
' STATE OWNERSHIP
'   Captures calculation mode, events and screen updating, switches them
'   off while it writes and restores them on every path. Sets mRunning
'   while it writes and refuses to start while an operation runs.
'
' ERROR POLICY
'   Raises ERR_STAGING_WRITE when a staging sheet or its names are
'   missing, and ERR_RUN_ACTIVE while an operation runs, before anything
'   is written; ERR_RECALC_PENDING when Excel does not finish
'   recalculating, before anything is judged or written;
'   ERR_OUTPUT_LAYOUT when the NettingSets or Trades table
'   it needs cannot be located, before Checks is written. A later error
'   is raised after Excel's settings have been restored; Checks may then
'   hold part of the findings, and no status is changed. A setting not
'   restored raises ERR_CLEANUP_FAILED.
'
' UPDATED
'   2026-11-03
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Kind                As Variant          'Staging kind
    Dim FormulaAt           As String           'A staged cell holding a formula
    Dim NsSheet             As Worksheet        'Import_NettingSets
    Dim TrSheet             As Worksheet        'Import_Trades
    Dim NsFirst             As Long             'Its first row below the header
    Dim TrFirst             As Long             'Its first row below the header
    Dim NsValues()          As String           'Staged netting sets, in contract order
    Dim TrValues()          As String           'Staged legs, in contract order
    Dim NsLines()           As Long             'Their CSV lines
    Dim TrLines()           As Long             'Their CSV lines
    Dim NsRows              As Long             'Staged netting sets
    Dim TrRows              As Long             'Staged legs
    Dim NsErrors            As Long             'Findings that keep Import_NettingSets from VALID
    Dim TrErrors            As Long             'Findings that keep Import_Trades from VALID
    Dim IDs()               As String           'The netting sets the legs may refer to
    Dim Regimes()           As String           'The regime override of each; blank for none
    Dim IDCount             As Long             'How many there are
    Dim NsSource            As String           'The sheet they come from
    Dim Context             As RunContext       'What a run reads from Params
    Dim ProblemSheet        As String           'Where a run would stop instead: the sheet
    Dim ProblemReference    As String           'Its parameter code, key or ID
    Dim ProblemRow          As Long             'Its row; 0 for none
    Dim ProblemMessage      As String           'Why
    Dim Result              As CsvReadResult    'What a check found
    Dim Findings()          As CsvFinding       'Its findings
    Dim Found()             As StagedFinding    'Every finding, for Checks
    Dim FoundCount          As Long             'How many there are
    Dim Summary             As String           'The first Checks line
    Dim Dialog              As String           'The same, in blocks, for the message
    Dim SavedCalculation    As XlCalculation    'Calculation mode before the validation
    Dim SavedEvents         As Boolean          'EnableEvents before the validation
    Dim SavedScreen         As Boolean          'ScreenUpdating before the validation
    Dim StateCaptured       As Boolean          'True once all three were read
    Dim ErrorNumber         As Long             'Primary error number; 0 for none
    Dim ErrorSource         As String           'Primary error source
    Dim ErrorDescription    As String           'Primary error description
    Dim CleanupDetails      As String           'Settings not restored

'------------------------------------------------------------------------------
' REFUSE
'------------------------------------------------------------------------------
    'A formula would be judged by its result, not by what was entered.
        If mRunning Then
            Err.Raise ERR_RUN_ACTIVE, "M_Main.CheckStaging", "An SA-CCR operation is already running."
        End If
        For Each Kind In Array("NettingSets", "Trades")
            FormulaAt = StagingFormulaCell(CStr(Kind))
            If Len(FormulaAt) > 0 Then
                FormulaAt = StagingSheetOf(CStr(Kind)) & "!" & FormulaAt
                CheckStaging = "RESULT=REFUSED; reason=" & FormulaAt & " holds a formula"
                Message = "Nothing was validated: " & FormulaAt & " holds a formula." & vbLf & vbLf & _
                          "Enter the value itself, as text."
                Exit Function
            End If
        Next Kind
        NsRows = ReadStaging("NettingSets", NsSheet, NsFirst, NsValues, NsLines)
        TrRows = ReadStaging("Trades", TrSheet, TrFirst, TrValues, TrLines)
        If NsRows = 0 And TrRows = 0 Then
            CheckStaging = "RESULT=REFUSED; reason=nothing is staged"
            Message = "Nothing was validated: " & SH_IMPORT_NS & " and " & SH_IMPORT_TR & " hold no rows." & _
                      vbLf & vbLf & "Import a CSV file first."
            Exit Function
        End If
        mRunning = True
        ProtectWorkbook
        On Error GoTo Failed
        SavedCalculation = Application.Calculation
        SavedEvents = Application.EnableEvents
        SavedScreen = Application.ScreenUpdating
        StateCaptured = True
        Application.ScreenUpdating = False
        Application.EnableEvents = False
        Application.Calculation = xlCalculationManual
    'A formula on Params, NettingSets or Trades is judged by its current
    'result, as a run judges it (#519).
        If Not InputsRecalculated() Then
            Err.Raise ERR_RECALC_PENDING, "M_Main.CheckStaging", _
                      "Excel did not finish recalculating the inputs, so nothing was validated."
        End If

'------------------------------------------------------------------------------
' EACH RECORD
'------------------------------------------------------------------------------
        ReDim Found(1 To 16)
        If NsRows > 0 Then
            CheckStagedRecords "NettingSets", NsValues, NsLines, NsRows, Result, Findings
            AddCsvFindings "NettingSets", NsSheet, NsFirst, Result, Findings, Found, FoundCount
            NsErrors = Result.FindingCount
        End If
        If TrRows > 0 Then
            CheckStagedRecords "Trades", TrValues, TrLines, TrRows, Result, Findings
            AddCsvFindings "Trades", TrSheet, TrFirst, Result, Findings, Found, FoundCount
            TrErrors = Result.FindingCount
        End If

'------------------------------------------------------------------------------
' BETWEEN RECORDS
'------------------------------------------------------------------------------
        If NsRows > 0 Then
            NsSource = SH_IMPORT_NS
            StagedKeys NsValues, NsRows, IDs, Regimes, IDCount
        Else
            NsSource = SH_NS
            LiveKeys IDs, Regimes, IDCount
        End If
        If TrRows > 0 Then
            CheckPortfolio TrValues, TrLines, TrRows, IDs, IDCount, Result, Findings, NsSource
            AddCsvFindings "Trades", TrSheet, TrFirst, Result, Findings, Found, FoundCount
            TrErrors = TrErrors + Result.FindingCount
        Else
            NsErrors = NsErrors + CheckLiveTrades(IDs, IDCount, Found, FoundCount)
        End If

'------------------------------------------------------------------------------
' AGAINST PARAMS
'------------------------------------------------------------------------------
    'What a run reads before the input sheets, read as a run reads it. When
    'a run would stop there, nothing that depends on it is judged, and no
    'staging sheet can be VALID.
        If CORE_Engine.LoadRunContext(Context, ProblemSheet, ProblemReference, ProblemRow, ProblemMessage) Then
            If TrRows > 0 Then
                CheckContext TrValues, TrLines, TrRows, Context, IDs, Regimes, IDCount, False, Result, Findings
                AddCsvFindings "Trades", TrSheet, TrFirst, Result, Findings, Found, FoundCount
                TrErrors = TrErrors + Result.FindingCount
            Else
                NsErrors = NsErrors + CheckLiveContext(Context, IDs, Regimes, IDCount, Found, FoundCount)
            End If
        Else
            AddParamsFinding ProblemSheet, ProblemReference, ProblemRow, ProblemMessage, Found, FoundCount
            If NsRows > 0 Then
                NsErrors = NsErrors + 1
            End If
            If TrRows > 0 Then
                TrErrors = TrErrors + 1
            End If
        End If

'------------------------------------------------------------------------------
' WRITE CHECKS, MARK THE CELLS, SET THE STATUS
'------------------------------------------------------------------------------
        Summary = StagingSummary(NsRows, NsErrors, TrRows, TrErrors, NsSource, FoundCount, Dialog)
        WriteStagingChecks Found, FoundCount, Summary
        UnmarkStaging "NettingSets"
        UnmarkStaging "Trades"
        MarkStaging Found, FoundCount
        If NsRows > 0 Then
            StagingMetaCell("NettingSets", 4).Value2 = IIf(NsErrors = 0, STAGING_VALID, STAGING_INVALID)
        End If
        If TrRows > 0 Then
            StagingMetaCell("Trades", 4).Value2 = IIf(TrErrors = 0, STAGING_VALID, STAGING_INVALID)
        End If
        RememberStagingValidation (FoundCount = 0)

'------------------------------------------------------------------------------
' RESTORE EXCEL AND REPORT
'------------------------------------------------------------------------------
CleanUp:
        On Error GoTo 0
        If StateCaptured Then
            CleanupDetails = RestoreSettings(SavedCalculation, SavedEvents, SavedScreen)
        End If
        mRunning = False
        If ErrorNumber <> 0 Then
            If Len(CleanupDetails) > 0 Then
                ErrorDescription = ErrorDescription & " Excel settings not restored: " & CleanupDetails
            End If
            Err.Raise ErrorNumber, ErrorSource, ErrorDescription
        End If
        If Len(CleanupDetails) > 0 Then
            Err.Raise ERR_CLEANUP_FAILED, "M_Main.CheckStaging", _
                      "The staging sheets were validated, but Excel settings were not restored: " & CleanupDetails
        End If
        CheckStaging = "RESULT=" & IIf(FoundCount = 0, "VALID", "INVALID") & "; errors=" & FoundCount & _
                       "; nettingsets=" & NsRows & "; trades=" & TrRows
        If FoundCount > 0 Then
            CheckStaging = CheckStaging & "; first=" & Found(1).Code & "|" & _
                           IIf(Len(Found(1).CellAddress) > 0, Found(1).SheetName & "!" & Found(1).CellAddress, "-")
        End If
        Message = Dialog
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


Private Function ReadStaging( _
    ByVal Kind As String, _
    ByRef Sheet As Worksheet, _
    ByRef FirstRow As Long, _
    ByRef Values() As String, _
    ByRef Lines() As Long) _
    As Long
'
'==============================================================================
'                                 ReadStaging
'------------------------------------------------------------------------------
' PURPOSE
'   Read a staging block down to its last row with a value: the CSV line
'   of each row, and its values as text in contract order.
'
' RETURNS
'   The number of rows, with Sheet, FirstRow, Values and Lines; 0 for an
'   empty block.
'
' METHOD
'   A value is read as Value2 holds it: text as it is, a number as CStr
'   writes it and an error value as Excel shows it, so that a value Excel
'   converted is judged as converted. The CSV line is a whole number
'   above 0, or 0 for anything else, a row added by hand.
'
' ERROR POLICY
'   As StagingHeaderRow.
'
' UPDATED
'   2026-10-28
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Width       As Long       'CSV line column and the contract's columns
    Dim LastRow     As Long       'Last row with a value
    Dim Grid        As Variant    'The block's values
    Dim RowCount    As Long       'Rows read
    Dim r           As Long       'Row
    Dim c           As Long       'Column of Grid

'------------------------------------------------------------------------------
' READ
'------------------------------------------------------------------------------
        FirstRow = StagingHeaderRow(Kind, Sheet) + 1
        Width = StagingWidth(Kind)
        LastRow = LastStagingRow(Sheet, FirstRow, Width)
        If LastRow < FirstRow Then
            Exit Function
        End If
        RowCount = LastRow - FirstRow + 1
        Grid = TableRange(Sheet, FirstRow, 1, LastRow, Width).Value2
        ReDim Values(1 To RowCount, 1 To Width - 1)
        ReDim Lines(1 To RowCount)
        For r = 1 To RowCount
            If VarType(Grid(r, 1)) = vbDouble Then
                If Grid(r, 1) >= 1 And Grid(r, 1) <= 2147483647# And Grid(r, 1) = Int(Grid(r, 1)) Then
                    Lines(r) = CLng(Grid(r, 1))
                End If
            End If
            For c = 2 To Width
                Values(r, c - 1) = StagedText(Grid(r, c))
            Next c
        Next r
        ReadStaging = RowCount

End Function


Private Function StagedText( _
    ByVal Value As Variant) _
    As String
'
'==============================================================================
' PURPOSE
'   A staged cell's value as text: Empty as "", an error value as Excel
'   shows it (CStr raises on one), anything else as CStr writes it.
'==============================================================================
'
        If IsError(Value) Then
            Select Case Value
                Case CVErr(xlErrNA)
                    StagedText = "#N/A"
                Case CVErr(xlErrDiv0)
                    StagedText = "#DIV/0!"
                Case CVErr(xlErrName)
                    StagedText = "#NAME?"
                Case CVErr(xlErrNull)
                    StagedText = "#NULL!"
                Case CVErr(xlErrNum)
                    StagedText = "#NUM!"
                Case CVErr(xlErrRef)
                    StagedText = "#REF!"
                Case Else
                    StagedText = "#VALUE!"
            End Select
        ElseIf Not IsEmpty(Value) Then
            StagedText = CStr(Value)
        End If

End Function


Private Sub StagedKeys( _
    ByRef Values() As String, _
    ByVal RowCount As Long, _
    ByRef IDs() As String, _
    ByRef Regimes() As String, _
    ByRef IDCount As Long)
'
'==============================================================================
' PURPOSE
'   The netting set IDs of the staged netting sets that have one, and
'   their regime overrides.
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim KeyColumn       As Long    'Contract column of the netting set ID
    Dim RegimeColumn    As Long    'Contract column of the regime override
    Dim r               As Long    'Row

'------------------------------------------------------------------------------
' COLLECT
'------------------------------------------------------------------------------
        KeyColumn = ContractColumn("NettingSets", "netting_set_id")
        RegimeColumn = ContractColumn("NettingSets", "regime_override")
        IDCount = 0
        ReDim IDs(1 To RowCount)
        ReDim Regimes(1 To RowCount)
        For r = 1 To RowCount
            If Len(Values(r, KeyColumn)) > 0 Then
                IDCount = IDCount + 1
                IDs(IDCount) = Values(r, KeyColumn)
                Regimes(IDCount) = Values(r, RegimeColumn)
            End If
        Next r

End Sub


Private Sub LiveKeys( _
    ByRef IDs() As String, _
    ByRef Regimes() As String, _
    ByRef IDCount As Long)
'
'==============================================================================
'                                   LiveKeys
'------------------------------------------------------------------------------
' PURPOSE
'   The netting set IDs on NettingSets, trimmed as the engine reads them
'   (SafeStr), and their regime overrides as it reads them (UTxt), for
'   staged legs when no netting sets are staged.
'
' ERROR POLICY
'   Raises ERR_OUTPUT_LAYOUT when the NettingSets table cannot be located
'   through its header name.
'
' UPDATED
'   2026-10-31
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Sheet       As Worksheet    'NettingSets
    Dim TopRow      As Long         'Its first data row
    Dim LastRow     As Long         'Its last row with a value
    Dim Grid        As Variant      'The columns from ID to regime override
    Dim r           As Long         'Row of Grid

'------------------------------------------------------------------------------
' COLLECT
'------------------------------------------------------------------------------
        IDCount = 0
        ReDim IDs(1 To 1)
        ReDim Regimes(1 To 1)
        TopRow = LiveFirstRow(SH_NS, NM_NS_HEADER)
        Set Sheet = GetSheet(SH_NS)
        LastRow = LastDataRowAny(Sheet, TopRow, NS_NCOLS, 0)
        If LastRow < TopRow Then
            Exit Sub
        End If
        Grid = TableRange(Sheet, TopRow, NS_ID, LastRow + 1, NS_REGIME).Value2
        ReDim IDs(1 To LastRow - TopRow + 1)
        ReDim Regimes(1 To LastRow - TopRow + 1)
        For r = 1 To LastRow - TopRow + 1
            If Len(SafeStr(Grid(r, 1))) > 0 Then
                IDCount = IDCount + 1
                IDs(IDCount) = SafeStr(Grid(r, 1))
                Regimes(IDCount) = UTxt(Grid(r, 1 + NS_REGIME - NS_ID))
            End If
        Next r

End Sub


Public Function StagingValidationCurrent() As Boolean
'
'==============================================================================
'                           StagingValidationCurrent
'------------------------------------------------------------------------------
' PURPOSE
'   Tell whether the last staging validation found no error and nothing it
'   judged has changed since (#209): both staging blocks, NettingSets,
'   Trades and Params, the FX and supervisory-factor tables included. A
'   commit (#210, #211) refuses otherwise, so that it never relies on a
'   validation of other inputs.
'
' RETURNS
'   True when the fingerprint stored by the last VALID validation matches
'   the inputs now and no staging block holds a formula. False when the
'   last validation was not VALID, there was none, the inputs changed,
'   even with events off, or the fingerprint cannot be taken because a
'   staging sheet or a Params name is missing.
'
' ERROR POLICY
'   Contains every error, which reads as not current.
'
' UPDATED
'   2026-10-31
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Stored      As String    'The fingerprint of the last VALID validation

'------------------------------------------------------------------------------
' COMPARE
'------------------------------------------------------------------------------
        On Error GoTo NotCurrent
        Stored = StoredText(STAGING_VALIDATED_NAME)
        If Len(Stored) = 0 Then
            Exit Function
        End If
        If Len(StagingFormulaCell("NettingSets")) > 0 Or Len(StagingFormulaCell("Trades")) > 0 Then
            Exit Function
        End If
        StagingValidationCurrent = (StagingValidationFingerprint() = Stored)
        Exit Function

NotCurrent:
        StagingValidationCurrent = False

End Function


Public Sub RefreshStagingStatus()
'
'==============================================================================
'                             RefreshStagingStatus
'------------------------------------------------------------------------------
' PURPOSE
'   Make the status cells of the staging sheets tell whether the last
'   staging validation still applies (#210). The change handlers mark only
'   edits made on a staging sheet with events on; a change to Params,
'   NettingSets or Trades, or an edit made with events off, would leave
'   VALID on view. When what the last validation judged has changed since
'   (its fingerprint, STAGING_CHECKED_NAME, no longer matches, or a
'   staging block holds a formula), a sheet that reads VALID reads OUT OF
'   DATE; when it matches again, a sheet that reads OUT OF DATE reads
'   VALID. Every other status is left as it is. The status is for the
'   user only: the commit (#211) checks StagingValidationCurrent itself,
'   never the status text.
'
' USAGE
'   Called by the Worksheet_Activate handlers of shImportNettingSets and
'   shImportTrades.
'
' ERROR POLICY
'   Contains every error: an event handler must not stop the user. A
'   status that cannot be refreshed is left as it was.
'
' UPDATED
'   2026-11-02
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Checked     As String     'Fingerprint of the last validation
    Dim Matches     As Boolean    'What it judged is unchanged
    Dim Events      As Boolean    'EnableEvents before the refresh
    Dim Kind        As Variant    'NettingSets, then Trades
    Dim StatusCell  As Range      'A status cell
    Dim Status      As String     'What it reads

'------------------------------------------------------------------------------
' COMPARE
'------------------------------------------------------------------------------
        If mRunning Then
            Exit Sub
        End If
        Events = Application.EnableEvents
        On Error GoTo Failed
        Checked = StoredText(STAGING_CHECKED_NAME)
        If Len(Checked) = 0 Then
            Exit Sub
        End If
        If Len(StagingFormulaCell("NettingSets")) = 0 And Len(StagingFormulaCell("Trades")) = 0 Then
            Matches = (ContainedValidationFingerprint() = Checked)
        End If

'------------------------------------------------------------------------------
' REFRESH
'------------------------------------------------------------------------------
    'A sheet protected as saved is protected again for the user interface
    'only, as in StagingEdited.
        Application.EnableEvents = False
        For Each Kind In Array("NettingSets", "Trades")
            Set StatusCell = NamedCell(StagingPrefixOf(CStr(Kind)) & "_Status")
            If Not StatusCell Is Nothing Then
                If StatusCell.Worksheet.ProtectContents Then
                    ProtectSheet StatusCell.Worksheet
                End If
                Status = CStr(StatusCell.Value2)
                If Matches And Status = STAGING_OUT_OF_DATE Then
                    StatusCell.Value2 = STAGING_VALID
                ElseIf Not Matches And Status = STAGING_VALID Then
                    StatusCell.Value2 = STAGING_OUT_OF_DATE
                End If
            End If
        Next Kind
        RibbonStateChanged

Failed:
        Application.EnableEvents = Events

End Sub


Private Function ContainedValidationFingerprint() As String
'
'==============================================================================
'                        ContainedValidationFingerprint
'------------------------------------------------------------------------------
' PURPOSE
'   StagingValidationFingerprint, or "" when it cannot be taken because a
'   staging sheet or a Params name is missing.
'
' UPDATED
'   2026-11-02
'==============================================================================
'
        On Error GoTo Failed
        ContainedValidationFingerprint = StagingValidationFingerprint()
        Exit Function

Failed:
        ContainedValidationFingerprint = ""

End Function


Public Sub CommitImport()
'
'==============================================================================
'                                 CommitImport
'------------------------------------------------------------------------------
' PURPOSE
'   Replace NettingSets and Trades with the staged rows of a VALID staging
'   validation (#211), after a confirmation that names each input sheet
'   with its rows now and the staged rows that replace them, or says it is
'   left as it is.
'
' USAGE
'   Commit import on the K-SACCR Ribbon tab, the Commit import button on
'   Import_NettingSets and Import_Trades, or the Macros dialog. Automation
'   calls CommitStaging_Silent.
'
' ERROR POLICY
'   Reports an error in a message, or in the Immediate window in silent
'   mode; never raises. Without a current VALID validation, or in silent
'   mode, where no confirmation can be given, nothing is committed.
'
' UPDATED
'   2026-11-03
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ResultLine  As String    'What the commit returned
    Dim Message     As String    'Its message

'------------------------------------------------------------------------------
' CONFIRM, COMMIT AND REPORT
'------------------------------------------------------------------------------
        On Error GoTo Failed
        If mRunning Then
            Notify "An SA-CCR operation is already running.", vbExclamation
            Exit Sub
        End If
        If Not InputsRecalculated() Then
            Notify RecalcRefusedText(), vbExclamation
            Exit Sub
        End If
        If Not StagingValidationCurrent() Then
            Notify CommitRefusedText(), vbExclamation
            Exit Sub
        End If
        If gSilent Then
            Notify "Nothing was committed." & vbLf & vbLf & "A commit is confirmed in a dialog, which silent mode " & _
                   "does not show. Automation calls CommitStaging_Silent.", vbExclamation
            Exit Sub
        End If
        If MsgBox(CommitQuestion(), vbOKCancel + vbExclamation + vbDefaultButton2, MSG_TITLE) <> vbOK Then
            Exit Sub
        End If
        ResultLine = CommitStaging(Message)
        RibbonStateChanged
        If Left$(ResultLine, Len("RESULT=COMMITTED")) = "RESULT=COMMITTED" Then
            Notify Message, vbInformation
        ElseIf InStr(1, ResultLine, "rollback=FAILED", vbBinaryCompare) > 0 Then
            Notify Message, vbCritical
        Else
            Notify Message, vbExclamation
        End If
        Exit Sub

Failed:
        Notify "Commit import: " & Err.Description, vbCritical

End Sub


Private Function InputsRecalculated() As Boolean
'
'==============================================================================
'                              InputsRecalculated
'------------------------------------------------------------------------------
' PURPOSE
'   Bring every formula whose precedents changed up to date, as a run does
'   before it reads its inputs (#195), so that staging validation and
'   commit judge a formula on Params, NettingSets or Trades by its current
'   result, not by the one Excel kept in manual calculation (#519).
'   Calculate refreshes no external data connection.
'
' RETURNS
'   True when Excel finished recalculating; False when it did not, for
'   example after Esc. Excel can report the calculation unfinished just
'   after Calculate returns, also with nothing left to interrupt it: an
'   owner certification run of 2026-11-05 refused two runs this way. So
'   Calculate is tried up to RECALC_ATTEMPTS times, and only a
'   recalculation that never reports done returns False.
'
' UPDATED
'   2026-11-05
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Attempt     As Long    'Recalculations tried

'------------------------------------------------------------------------------
' RECALCULATE
'------------------------------------------------------------------------------
        For Attempt = 1 To RECALC_ATTEMPTS
            Application.Calculate
            If Application.CalculationState = xlDone Then
                InputsRecalculated = True
                Exit Function
            End If
        Next Attempt

End Function


Private Function RecalcRefusedText() As String
'
'==============================================================================
' PURPOSE
'   The message of a commit refused because Excel did not finish
'   recalculating the inputs (#519).
'==============================================================================
'
        RecalcRefusedText = "Nothing was committed." & vbLf & vbLf & "Excel did not finish recalculating the " & _
                            "inputs." & vbLf & vbLf & "Press F9, then Commit import again."

End Function


Private Function CommitRefusedText() As String
'
'==============================================================================
'                              CommitRefusedText
'------------------------------------------------------------------------------
' PURPOSE
'   The message of a commit refused for want of a current VALID staging
'   validation (#211), for CommitImport and CommitStaging.
'
' UPDATED
'   2026-11-03
'==============================================================================
'
        CommitRefusedText = "Nothing was committed." & vbLf & vbLf & "The staged rows have no VALID validation " & _
                            "that still matches them, the input sheets and Params." & vbLf & vbLf & _
                            "Run Validate staging, correct any finding, then commit."

End Function


Private Function CommitQuestion() As String
'
'==============================================================================
'                                CommitQuestion
'------------------------------------------------------------------------------
' PURPOSE
'   The confirmation of CommitImport (#211): for each input sheet, its
'   rows now and the staged rows that replace them, or that it is left as
'   it is because its staging sheet holds no rows; what follows a commit;
'   and the question.
'
' UPDATED
'   2026-11-03
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Kind        As Variant      'NettingSets, then Trades
    Dim Staged      As Worksheet    'Its staging sheet
    Dim StagedFirst As Long         'Its first row below the header
    Dim StagedRows  As Long         'Its staged rows
    Dim LiveName    As String       'Its input sheet
    Dim LiveFirst   As Long         'Its first data row
    Dim LiveRows    As Long         'Its rows now
    Dim Text        As String       'The question

'------------------------------------------------------------------------------
' COUNT AND ASK
'------------------------------------------------------------------------------
        Text = "Commit replaces the input sheets with the staged rows of the VALID validation:" & vbLf
        For Each Kind In Array("NettingSets", "Trades")
            StagedFirst = StagingHeaderRow(CStr(Kind), Staged) + 1
            StagedRows = LastStagingRow(Staged, StagedFirst, StagingWidth(CStr(Kind))) - StagedFirst + 1
            LiveName = IIf(IsNettingSetsKind(CStr(Kind)), SH_NS, SH_TRADES)
            LiveFirst = TableFirstRow(LiveName)
            LiveRows = LastStagingRow(GetSheet(LiveName), LiveFirst, _
                                      IIf(IsNettingSetsKind(CStr(Kind)), NS_NCOLS, TR_NCOLS)) - LiveFirst + 1
            If StagedRows > 0 Then
                Text = Text & vbLf & "- " & LiveName & ": its " & Format$(LiveRows, "#,##0") & _
                       IIf(LiveRows = 1, " row", " rows") & " replaced by the " & Format$(StagedRows, "#,##0") & _
                       IIf(StagedRows = 1, " staged row", " staged rows") & " of " & Staged.Name
            Else
                Text = Text & vbLf & "- " & LiveName & ": left as it is; " & Staged.Name & " holds no rows"
            End If
        Next Kind
        CommitQuestion = Text & vbLf & vbLf & "Results will read OUT OF DATE until the next Run SA-CCR. Undo cannot " & _
                         "take the commit back; the staged rows stay on the staging sheets." & vbLf & vbLf & "Commit?"

End Function


Public Function CommitStaging_Silent() As String
'
'==============================================================================
'                             CommitStaging_Silent
'------------------------------------------------------------------------------
' PURPOSE
'   The commit of the staging sheets into the input sheets (#211), for
'   automation and tests, without a confirmation or a message.
'
' RETURNS
'   "RESULT=COMMITTED; nettingsets=N; trades=M", N and M being the rows
'   written to NettingSets and Trades, 0 for a sheet left as it was;
'   "RESULT=REFUSED; reason=..." when nothing was written; or
'   "RESULT=FAILED; error=...; rollback=..." when a write failed, with
'   rollback=NONE when nothing had been written, VERIFIED when the input
'   sheets were put back and read back as they were, and FAILED: ... when
'   they could not be.
'
' ERROR POLICY
'   Raises ERR_RUN_ACTIVE while an operation runs, and ERR_CLEANUP_FAILED
'   when the commit was done but Excel settings were not restored.
'
' UPDATED
'   2026-11-02
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Message     As String    'The message, not shown

'------------------------------------------------------------------------------
' COMMIT
'------------------------------------------------------------------------------
        CommitStaging_Silent = CommitStaging(Message)

End Function


Private Function CommitStaging( _
    ByRef Message As String) _
    As String
'
'==============================================================================
'                                CommitStaging
'------------------------------------------------------------------------------
' PURPOSE
'   Replace the input sheets with the staged rows (#211): NettingSets with
'   the rows of Import_NettingSets when it holds any, Trades with those of
'   Import_Trades when it holds any. A staging sheet without rows leaves
'   its input sheet as it is: nothing is ever erased for lack of staged
'   rows. Each replacement is whole: the staged rows are written from the
'   first data row, and the input cells of the rows below them that held
'   a value are cleared.
'
' METHOD
'   1. Recalculate the workbook, as a run does, and refuse when Excel does
'      not finish (#519); then refuse unless StagingValidationCurrent is
'      True: the last staging validation was VALID and nothing it judged,
'      staging blocks, input sheets and Params, has changed since, also
'      with events off or in a formula's result. The status text is never
'      read.
'   2. Convert each staged value as a value typed by hand would be:
'      numbers and integers to numbers, as the reader reads them (Val),
'      dates to Excel dates, every other type to text written with a
'      leading apostrophe, so that leading zeros and text that looks like
'      a number, a date or a formula stay as staged. Blank stays blank.
'   3. Keep the values and formulas of every input cell the commit may
'      touch, down to the last row with a value or a formula, whatever
'      the formula returns (LastInputRow, #512), a formula cell found as
'      one (FormulasOnly), then write and read back each sheet; a value
'      read back that differs from the one written, or a value or formula
'      left below the committed rows, is a failure.
'   4. On any failure, write back the kept cells of every input sheet the
'      commit touched and read them back; the result tells whether that
'      was verified, apart from the error itself.
'   5. On success each committed staging sheet reads COMMITTED and the
'      stored VALID validation is emptied, so that the same rows are not
'      committed twice without a new validation. Results, Checks and the
'      run and reconciliation fingerprints are left as they are: Results
'      reads OUT OF DATE until the next run, as after any edit.
'
' ERROR POLICY
'   See CommitStaging_Silent. gTestFault = "commit-write" raises
'   ERR_INJECTED_FAULT once the first input sheet is written;
'   "commit-verify" makes the read-back fail; "commit-rollback" raises it
'   there too and makes the write-back fail as well.
'
' UPDATED
'   2026-11-03
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Kinds(1 To 2)       As String           'NettingSets, Trades
    Dim LiveSheet(1 To 2)   As Worksheet        'The input sheets
    Dim LiveFirst(1 To 2)   As Long             'Their first data rows
    Dim InputWidth(1 To 2)       As Long             'Their input columns
    Dim NewRows(1 To 2)     As Long             'Staged rows to write; 0 leaves the sheet
    Dim OldLast(1 To 2)     As Long             'Last row with a value before the commit
    Dim KeptLast(1 To 2)    As Long             'Last row of the kept cells
    Dim NewBlock(1 To 2)    As Variant          'The typed values to write
    Dim KeptValues(1 To 2)  As Variant          'Value2 of the kept cells
    Dim KeptFormulas(1 To 2) As Variant         'Formula of the kept cells
    Dim Touched(1 To 2)     As Boolean          'The commit may have changed the sheet
    Dim StatusSet(1 To 2)   As Boolean          'The staging sheet was set to COMMITTED
    Dim OldStatus(1 To 2)   As String           'What it read before
    Dim StagedSheet         As Worksheet        'A staging sheet, from ReadStaging
    Dim StagedFirst         As Long             'Its first row, from ReadStaging
    Dim Values()            As String           'Its rows as text
    Dim Lines()             As Long             'Their CSV lines, unused
    Dim k                   As Long             'Kind
    Dim SavedCalculation    As XlCalculation    'Calculation mode before the commit
    Dim SavedEvents         As Boolean          'EnableEvents before the commit
    Dim SavedScreen         As Boolean          'ScreenUpdating before the commit
    Dim StateCaptured       As Boolean          'True once all three were read
    Dim Committed           As Boolean          'Every sheet was written and read back
    Dim ErrorNumber         As Long             'Primary error number; 0 for none
    Dim ErrorDescription    As String           'Primary error description
    Dim Rollback            As String           'NONE, VERIFIED or FAILED: details
    Dim CleanupDetails      As String           'Settings not restored

'------------------------------------------------------------------------------
' REFUSE A COMMIT THAT NO CURRENT VALIDATION ALLOWS
'------------------------------------------------------------------------------
        If mRunning Then
            Err.Raise ERR_RUN_ACTIVE, "M_Main.CommitStaging", "An SA-CCR operation is already running."
        End If
    'A formula input whose precedents changed since the validation must
    'show its new result before the fingerprints are compared (#519).
        If Not InputsRecalculated() Then
            CommitStaging = "RESULT=REFUSED; reason=Excel did not finish recalculating"
            Message = RecalcRefusedText()
            Exit Function
        End If
        If Not StagingValidationCurrent() Then
            CommitStaging = "RESULT=REFUSED; reason=no current VALID staging validation"
            Message = CommitRefusedText()
            Exit Function
        End If

'------------------------------------------------------------------------------
' PLAN: TYPED VALUES AND THE CELLS TO KEEP
'------------------------------------------------------------------------------
        Kinds(1) = "NettingSets"
        Kinds(2) = "Trades"
        mRunning = True
        ProtectWorkbook
        On Error GoTo Failed
        For k = 1 To 2
            NewRows(k) = ReadStaging(Kinds(k), StagedSheet, StagedFirst, Values, Lines)
            If NewRows(k) > 0 Then
                NewBlock(k) = TypedBlock(Kinds(k), Values, NewRows(k))
                Set LiveSheet(k) = GetSheet(IIf(k = 1, SH_NS, SH_TRADES))
                LiveFirst(k) = TableFirstRow(LiveSheet(k).Name)
                InputWidth(k) = IIf(k = 1, NS_NCOLS, TR_NCOLS)
                OldLast(k) = LastInputRow(LiveSheet(k), LiveFirst(k), InputWidth(k))
                KeptLast(k) = LiveFirst(k) + NewRows(k) - 1
                If OldLast(k) > KeptLast(k) Then
                    KeptLast(k) = OldLast(k)
                End If
                With TableRange(LiveSheet(k), LiveFirst(k), 1, KeptLast(k), InputWidth(k))
                    KeptValues(k) = .Value2
                    KeptFormulas(k) = FormulasOnly(.Cells)
                End With
            End If
        Next k
        If NewRows(1) = 0 And NewRows(2) = 0 Then
            mRunning = False
            CommitStaging = "RESULT=REFUSED; reason=nothing is staged"
            Message = "Nothing was committed: " & SH_IMPORT_NS & " and " & SH_IMPORT_TR & " hold no rows."
            Exit Function
        End If

'------------------------------------------------------------------------------
' WRITE AND READ BACK
'------------------------------------------------------------------------------
        SavedCalculation = Application.Calculation
        SavedEvents = Application.EnableEvents
        SavedScreen = Application.ScreenUpdating
        StateCaptured = True
        Application.ScreenUpdating = False
        Application.EnableEvents = False
        Application.Calculation = xlCalculationManual
        For k = 1 To 2
            If NewRows(k) > 0 Then
                Touched(k) = True
                TableRange(LiveSheet(k), LiveFirst(k), 1, LiveFirst(k) + NewRows(k) - 1, InputWidth(k)).Value = NewBlock(k)
                If OldLast(k) >= LiveFirst(k) + NewRows(k) Then
                    TableRange(LiveSheet(k), LiveFirst(k) + NewRows(k), 1, OldLast(k), InputWidth(k)).ClearContents
                End If
                If gTestFault = "commit-write" Or gTestFault = "commit-rollback" Then
                    Err.Raise ERR_INJECTED_FAULT, "M_Main.CommitStaging", "Injected commit write failure."
                End If
            End If
        Next k
        For k = 1 To 2
            If NewRows(k) > 0 Then
                CheckWritten LiveSheet(k), LiveFirst(k), InputWidth(k), NewBlock(k), NewRows(k), OldLast(k)
            End If
        Next k
        ThisWorkbook.Names.Add Name:=STAGING_VALIDATED_NAME, RefersTo:="=""""", Visible:=False
        For k = 1 To 2
            If NewRows(k) > 0 Then
                OldStatus(k) = CStr(StagingMetaCell(Kinds(k), 4).Value2)
                StatusSet(k) = True
                StagingMetaCell(Kinds(k), 4).Value2 = STAGING_COMMITTED
            End If
        Next k
        Committed = True

'------------------------------------------------------------------------------
' RESTORE EXCEL AND REPORT
'------------------------------------------------------------------------------
CleanUp:
        On Error GoTo 0
        If StateCaptured Then
            CleanupDetails = RestoreSettings(SavedCalculation, SavedEvents, SavedScreen)
        End If
        mRunning = False
        If Not Committed Then
            CommitStaging = "RESULT=FAILED; error=" & ErrorDescription & "; rollback=" & Rollback
            Message = "Nothing was committed." & vbLf & vbLf & ErrorDescription & " (error " & ErrorNumber & ")."
            If Left$(Rollback, 6) = "FAILED" Then
                Message = Message & vbLf & vbLf & "The input sheets could not be put back as they were: " & _
                          Mid$(Rollback, 9) & ". Check NettingSets and Trades before any run."
            ElseIf Rollback = "VERIFIED" Then
                Message = Message & vbLf & vbLf & "The input sheets were put back as they were and checked."
            End If
            If Len(CleanupDetails) > 0 Then
                Message = Message & vbLf & vbLf & "Excel settings were not restored: " & CleanupDetails
            End If
            Exit Function
        End If
        If Len(CleanupDetails) > 0 Then
            Err.Raise ERR_CLEANUP_FAILED, "M_Main.CommitStaging", _
                      "The staged rows were committed, but Excel settings were not restored: " & CleanupDetails
        End If
        CommitStaging = "RESULT=COMMITTED; nettingsets=" & NewRows(1) & "; trades=" & NewRows(2)
        Message = "Committed." & vbLf & vbLf & SH_NS & ": " & _
                  IIf(NewRows(1) > 0, CountText(NewRows(1), "row") & " written", "unchanged") & vbLf & SH_TRADES & ": " & _
                  IIf(NewRows(2) > 0, CountText(NewRows(2), "row") & " written", "unchanged") & vbLf & vbLf & _
                  "Results reads OUT OF DATE until the next Run SA-CCR."
        Exit Function

Failed:
        ErrorNumber = Err.Number
        ErrorDescription = Err.Description
        Rollback = "NONE"
        For k = 1 To 2
            If Touched(k) Then
                Rollback = ""
            End If
        Next k
        If Len(Rollback) = 0 Then
            Rollback = PutLiveBack(LiveSheet, LiveFirst, InputWidth, KeptLast, KeptValues, KeptFormulas, Touched)
        End If
        PutStatusesBack Kinds, StatusSet, OldStatus
        Resume CleanUp

End Function


Private Function TypedBlock( _
    ByVal Kind As String, _
    ByRef Values() As String, _
    ByVal RowCount As Long) _
    As Variant
'
'==============================================================================
'                                  TypedBlock
'------------------------------------------------------------------------------
' PURPOSE
'   The staged rows of Kind as the values a commit writes (#211), by the
'   type of each contract column: a number or an integer as the reader
'   reads it (Val), a date as an Excel date, any other type as text with
'   a leading apostrophe, and blank as Empty. The contract's columns are
'   the input sheet's, in the same order (check_csv_contract.py).
'
' ERROR POLICY
'   Raises ERR_CSV_READ when the contract has no Kind file, and
'   ERR_STAGING_WRITE for a date that cannot be read, which a VALID
'   validation rules out.
'
' UPDATED
'   2026-11-04
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Spec        As CsvSpec            'The contract's file
    Dim Columns()   As CsvColumn          'Its columns
    Dim Rules()     As CsvRule            'Its rules, unused
    Dim Block()     As Variant            'The values
    Dim Serial      As Double             'A date as a serial number
    Dim r           As Long               'Row
    Dim c           As Long               'Column

'------------------------------------------------------------------------------
' CONVERT
'------------------------------------------------------------------------------
        If Not CsvSpecOf(Kind, Spec, Columns, Rules) Then
            Err.Raise ERR_CSV_READ, "M_Main.TypedBlock", "the contract has no " & Kind & " file"
        End If
        ReDim Block(1 To RowCount, 1 To Spec.ColumnCount)
        For r = 1 To RowCount
            For c = 1 To Spec.ColumnCount
                If Len(Values(r, c)) > 0 Then
                    Select Case Columns(c).Kind
                        Case "number", "integer"
                            Block(r, c) = Val(Values(r, c))
                        Case "date"
                            Serial = ToSerial(Values(r, c))
                            If Serial < 0# Then
                                Err.Raise ERR_STAGING_WRITE, "M_Main.TypedBlock", _
                                          Kind & " row " & r & ": '" & Values(r, c) & "' is not a date"
                            End If
                            Block(r, c) = CDate(Serial)
                        Case Else
                            Block(r, c) = "'" & Values(r, c)
                    End Select
                End If
            Next c
        Next r
        TypedBlock = Block

End Function


Private Sub CheckWritten( _
    ByVal Sheet As Worksheet, _
    ByVal FirstRow As Long, _
    ByVal Width As Long, _
    ByRef Block As Variant, _
    ByVal RowCount As Long, _
    ByVal OldLast As Long)
'
'==============================================================================
'                                 CheckWritten
'------------------------------------------------------------------------------
' PURPOSE
'   Read back what a commit wrote to an input sheet (#211): each cell must
'   hold the value written, a text without its apostrophe and a date as
'   its serial number, and no input cell below the written rows may hold
'   a value.
'
' ERROR POLICY
'   Raises ERR_STAGING_WRITE at the first difference; gTestFault =
'   "commit-verify" raises it at once.
'
' UPDATED
'   2026-11-02
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Grid        As Variant    'The cells read back, Value2
    Dim Expected    As Variant    'What a cell should hold
    Dim r           As Long       'Row
    Dim c           As Long       'Column

'------------------------------------------------------------------------------
' READ BACK
'------------------------------------------------------------------------------
        If gTestFault = "commit-verify" Then
            Err.Raise ERR_STAGING_WRITE, "M_Main.CheckWritten", "Injected read-back difference on " & Sheet.Name
        End If
        Grid = TableRange(Sheet, FirstRow, 1, FirstRow + RowCount - 1, Width).Value2
        For r = 1 To RowCount
            For c = 1 To Width
                Expected = Block(r, c)
                If VarType(Expected) = vbString Then
                    Expected = Mid$(Expected, 2)
                ElseIf VarType(Expected) = vbDate Then
                    Expected = CDbl(Expected)
                End If
                If Not SameCellValue(Grid(r, c), Expected) Then
                    Err.Raise ERR_STAGING_WRITE, "M_Main.CheckWritten", Sheet.Name & "!" & _
                              TableCell(Sheet, FirstRow + r - 1, c).Address(False, False) & " does not read as written"
                End If
            Next c
        Next r
        If LastInputRow(Sheet, FirstRow, Width) <> FirstRow + RowCount - 1 Then
            Err.Raise ERR_STAGING_WRITE, "M_Main.CheckWritten", Sheet.Name & ": a row below the committed rows " & _
                      "still holds a value or a formula"
        End If

End Sub


Private Function SameCellValue( _
    ByVal Found As Variant, _
    ByVal Expected As Variant) _
    As Boolean
'
'==============================================================================
'                                SameCellValue
'------------------------------------------------------------------------------
' PURPOSE
'   Whether a cell's Value2 is the value expected: both blank, both the
'   same text, both the same number, both the same logical value or both
'   the same error. A blank cell and an empty text are alike.
'
' UPDATED
'   2026-11-02
'==============================================================================
'
        If IsEmpty(Found) Or (VarType(Found) = vbString And Len(CStr(Found)) = 0 And Not IsError(Found)) Then
            SameCellValue = IsEmpty(Expected) Or (VarType(Expected) = vbString And Len(CStr(Expected)) = 0)
        ElseIf IsError(Found) Or IsError(Expected) Then
            If IsError(Found) And IsError(Expected) Then
                SameCellValue = (CStr(Found) = CStr(Expected))
            End If
        ElseIf VarType(Found) <> VarType(Expected) Then
            If IsPlainNumber(Found) And IsPlainNumber(Expected) Then
                SameCellValue = (CDbl(Found) = CDbl(Expected))
            End If
        Else
            SameCellValue = (Found = Expected)
        End If

End Function


Private Function IsPlainNumber( _
    ByVal Content As Variant) _
    As Boolean
'
'==============================================================================
'                                IsPlainNumber
'------------------------------------------------------------------------------
' PURPOSE
'   Whether Content is a number held as a number: not a text, a logical
'   value, a blank or an error.
'
' UPDATED
'   2026-11-02
'==============================================================================
'
        Select Case VarType(Content)
            Case vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal, vbByte, vbDate
                IsPlainNumber = True
        End Select

End Function


Private Sub PutStatusesBack( _
    ByRef Kinds() As String, _
    ByRef StatusSet() As Boolean, _
    ByRef OldStatus() As String)
'
'==============================================================================
'                               PutStatusesBack
'------------------------------------------------------------------------------
' PURPOSE
'   Give back their previous text to the status cells a failed commit had
'   already set to COMMITTED (#211), so that an interrupted commit never
'   reads COMMITTED.
'
' ERROR POLICY
'   Contains every error.
'
' UPDATED
'   2026-11-02
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim k   As Long    'Kind

'------------------------------------------------------------------------------
' PUT BACK
'------------------------------------------------------------------------------
        On Error Resume Next
        For k = LBound(StatusSet) To UBound(StatusSet)
            If StatusSet(k) Then
                StagingMetaCell(Kinds(k), 4).Value2 = OldStatus(k)
            End If
        Next k

End Sub


Private Function FormulasOnly( _
    ByVal Target As Range) _
    As Variant
'
'==============================================================================
'                                 FormulasOnly
'------------------------------------------------------------------------------
' PURPOSE
'   The formulas of Target's formula cells, by row and column within
'   Target, and "" for every other cell (#211). Formula returns the value
'   of a constant cell, so a text that starts with "=" would be taken for
'   a formula; the formula cells are found with SpecialCells instead. A
'   single cell is read by itself: SpecialCells on one cell searches the
'   whole sheet (#357).
'
' UPDATED
'   2026-11-06
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Found       As Range      'The formula cells; Nothing for none
    Dim Area        As Range      'One area of them
    Dim Cell        As Range      'One of them
    Dim Out()       As Variant    'The result

'------------------------------------------------------------------------------
' FIND
'------------------------------------------------------------------------------
        ReDim Out(1 To Target.Rows.Count, 1 To Target.Columns.Count)
        If Target.Cells.Count = 1 Then
            If Target.HasFormula Then
                Out(1, 1) = Target.Formula
            End If
            FormulasOnly = Out
            Exit Function
        End If
        If Not IsNull(Target.HasFormula) Then
            If Not Target.HasFormula Then
                FormulasOnly = Out
                Exit Function
            End If
        End If
        On Error Resume Next
        Set Found = Target.SpecialCells(xlCellTypeFormulas)
        On Error GoTo 0
        If Not Found Is Nothing Then
            For Each Area In Found.Areas
                For Each Cell In Area.Cells
                    Out(Cell.Row - Target.Row + 1, Cell.Column - Target.Column + 1) = Cell.Formula
                Next Cell
            Next Area
        End If
        FormulasOnly = Out

End Function


Private Function PutLiveBack( _
    ByRef LiveSheet() As Worksheet, _
    ByRef LiveFirst() As Long, _
    ByRef InputWidth() As Long, _
    ByRef KeptLast() As Long, _
    ByRef KeptValues() As Variant, _
    ByRef KeptFormulas() As Variant, _
    ByRef Touched() As Boolean) _
    As String
'
'==============================================================================
'                                 PutLiveBack
'------------------------------------------------------------------------------
' PURPOSE
'   Write back the kept input cells of each input sheet a failed commit
'   touched (#211) and read them back: a formula as its formula, a text
'   with a leading apostrophe so that it stays text, any other value as
'   it was.
'
' RETURNS
'   "VERIFIED" when every kept cell reads back as it was; otherwise
'   "FAILED: " and what failed, for each sheet.
'
' ERROR POLICY
'   Contains every error. gTestFault = "commit-rollback" makes the write
'   fail.
'
' UPDATED
'   2026-11-02
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Problems    As String     'What failed
    Dim k           As Long       'Kind

'------------------------------------------------------------------------------
' WRITE BACK EACH TOUCHED SHEET
'------------------------------------------------------------------------------
        For k = LBound(Touched) To UBound(Touched)
            If Touched(k) Then
                Problems = Problems & PutSheetBack(LiveSheet(k), LiveFirst(k), InputWidth(k), KeptLast(k), _
                                                   KeptValues(k), KeptFormulas(k))
            End If
        Next k
        If Len(Problems) = 0 Then
            PutLiveBack = "VERIFIED"
        Else
            PutLiveBack = "FAILED: " & Mid$(Problems, 3)
        End If

End Function


Private Function PutSheetBack( _
    ByVal Sheet As Worksheet, _
    ByVal FirstRow As Long, _
    ByVal Width As Long, _
    ByVal LastRow As Long, _
    ByRef KeptValues As Variant, _
    ByRef KeptFormulas As Variant) _
    As String
'
'==============================================================================
'                                 PutSheetBack
'------------------------------------------------------------------------------
' PURPOSE
'   Write back the kept cells of one input sheet and read them back, for
'   PutLiveBack.
'
' RETURNS
'   "" when every cell reads back as it was; otherwise "; " and what
'   failed.
'
' ERROR POLICY
'   Contains every error.
'
' UPDATED
'   2026-11-06
'==============================================================================
'
        On Error GoTo Failed
        PutSheetBack = PutRangeBack(TableRange(Sheet, FirstRow, 1, LastRow, Width), KeptValues, KeptFormulas)
        Exit Function

Failed:
        PutSheetBack = "; " & Sheet.Name & ": " & Err.Description

End Function


Private Function PutRangeBack( _
    ByVal Target As Range, _
    ByRef KeptValues As Variant, _
    ByRef KeptFormulas As Variant) _
    As String
'
'==============================================================================
'                                 PutRangeBack
'------------------------------------------------------------------------------
' PURPOSE
'   Write back the kept cells of a block and read them back: a formula as
'   its formula, a text with a leading apostrophe so that it stays text,
'   any other value as it was. Used by a failed commit (#211) and a failed
'   restore of the regulatory parameters (#357).
'
' INPUTS
'   Target: the block, as it was kept.
'   KeptValues: its Value2, by row and column, also for a single cell.
'   KeptFormulas: its formulas, as FormulasOnly gives them.
'
' RETURNS
'   "" when every cell reads back as it was; otherwise "; " and what
'   failed.
'
' ERROR POLICY
'   Contains every error. gTestFault = "commit-rollback" or
'   "restore-rollback" makes the write fail.
'
' UPDATED
'   2026-11-06
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Restore()   As Variant    'What is written back
    Dim Grid        As Variant    'What is read back, Value2
    Dim Formulas    As Variant    'What is read back, Formula
    Dim Text        As String     'A kept formula
    Dim RowCount    As Long       'Rows of the block
    Dim ColCount    As Long       'Columns of the block
    Dim r           As Long       'Row
    Dim c           As Long       'Column

'------------------------------------------------------------------------------
' WRITE BACK
'------------------------------------------------------------------------------
        On Error GoTo Failed
        RowCount = Target.Rows.Count
        ColCount = Target.Columns.Count
        ReDim Restore(1 To RowCount, 1 To ColCount)
        For r = 1 To RowCount
            For c = 1 To ColCount
                Text = CStr(KeptFormulas(r, c))
                If Left$(Text, 1) = "=" Then
                    Restore(r, c) = Text
                ElseIf VarType(KeptValues(r, c)) = vbString Then
                    Restore(r, c) = "'" & KeptValues(r, c)
                Else
                    Restore(r, c) = KeptValues(r, c)
                End If
            Next c
        Next r
        If gTestFault = "commit-rollback" Or gTestFault = "restore-rollback" Then
            Err.Raise ERR_INJECTED_FAULT, "M_Main.PutRangeBack", "Injected write-back failure."
        End If
        Target.Value = Restore

'------------------------------------------------------------------------------
' READ BACK
'------------------------------------------------------------------------------
        Grid = BlockValues(Target)
        Formulas = FormulasOnly(Target)
        For r = 1 To RowCount
            For c = 1 To ColCount
                Text = CStr(KeptFormulas(r, c))
                If Left$(Text, 1) = "=" Then
                    If CStr(Formulas(r, c)) <> Text Then
                        PutRangeBack = "; " & Target.Worksheet.Name & "!" & Target.Cells(r, c).Address(False, False) & _
                                       " does not hold its formula again"
                        Exit Function
                    End If
                ElseIf Len(CStr(Formulas(r, c))) > 0 Or Not SameCellValue(Grid(r, c), KeptValues(r, c)) Then
                    PutRangeBack = "; " & Target.Worksheet.Name & "!" & Target.Cells(r, c).Address(False, False) & _
                                   " does not hold its value again"
                    Exit Function
                End If
            Next c
        Next r
        Exit Function

Failed:
        PutRangeBack = "; " & Target.Worksheet.Name & ": " & Err.Description

End Function


Private Function BlockValues( _
    ByVal Target As Range) _
    As Variant
'
'==============================================================================
'                                 BlockValues
'------------------------------------------------------------------------------
' PURPOSE
'   Target's Value2 by row and column, also for a single cell, whose
'   Value2 Excel gives as a value rather than an array.
'
' UPDATED
'   2026-11-06
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim One(1 To 1, 1 To 1) As Variant    'A single cell's value

'------------------------------------------------------------------------------
' READ
'------------------------------------------------------------------------------
        If Target.Cells.Count = 1 Then
            One(1, 1) = Target.Value2
            BlockValues = One
        Else
            BlockValues = Target.Value2
        End If

End Function


Private Function StagingValidationFingerprint() As String
'
'==============================================================================
'                         StagingValidationFingerprint
'------------------------------------------------------------------------------
' PURPOSE
'   Fingerprint what a staging validation judges (#209): everything a run
'   reads (CORE_Engine.InputFingerprint), which covers NettingSets, Trades
'   and Params with its two tables, and both staging blocks. The live
'   sheets are all included, also the one a staged block replaces: an
'   edit there makes a validation out of date rather than risk missing
'   one.
'
' ERROR POLICY
'   Raises the errors of InputFingerprint and StagingFingerprint.
'
' UPDATED
'   2026-10-31
'==============================================================================
'
        StagingValidationFingerprint = TextHash(Join(Array("SV1", CORE_Engine.InputFingerprint(), _
                                                StagingFingerprint("NettingSets"), StagingFingerprint("Trades")), Chr$(29)))

End Function


Private Sub RememberStagingValidation( _
    ByVal Valid As Boolean)
'
'==============================================================================
'                          RememberStagingValidation
'------------------------------------------------------------------------------
' PURPOSE
'   Store the fingerprint of a VALID staging validation in the hidden name
'   STAGING_VALIDATED_NAME, for StagingValidationCurrent; any other result
'   empties it, so that an earlier VALID validation cannot be relied on.
'   The fingerprint of every completed validation, whatever its result,
'   goes to STAGING_CHECKED_NAME, for RefreshStagingStatus (#210); "" when
'   it cannot be taken.
'
' UPDATED
'   2026-11-02
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Fingerprint     As String    'What is stored for a VALID validation; "" for none
    Dim Checked         As String    'What is stored for any validation

'------------------------------------------------------------------------------
' STORE
'------------------------------------------------------------------------------
        If Valid Then
            Fingerprint = StagingValidationFingerprint()
            Checked = Fingerprint
        Else
            Checked = ContainedValidationFingerprint()
        End If
        ThisWorkbook.Names.Add Name:=STAGING_VALIDATED_NAME, RefersTo:="=""" & Fingerprint & """", Visible:=False
        ThisWorkbook.Names.Add Name:=STAGING_CHECKED_NAME, RefersTo:="=""" & Checked & """", Visible:=False

End Sub


Private Function CheckLiveContext( _
    ByRef Context As RunContext, _
    ByRef IDs() As String, _
    ByRef Regimes() As String, _
    ByVal IDCount As Long, _
    ByRef Found() As StagedFinding, _
    ByRef FoundCount As Long) _
    As Long
'
'==============================================================================
'                               CheckLiveContext
'------------------------------------------------------------------------------
' PURPOSE
'   When only netting sets are staged, judge each leg on Trades with the
'   factor and regime rules against them (#209), because a staged regime
'   override can change a leg's outcome. The legs are read as the engine
'   reads them (UTxt). A finding is on the leg's Asset class or Sub-class
'   cell on Trades, which is not marked or changed. The other rules do not
'   depend on the netting sets and are left to the run.
'
' RETURNS
'   The number of findings added to Found.
'
' ERROR POLICY
'   Raises ERR_OUTPUT_LAYOUT when the Trades table cannot be located
'   through its header name.
'
' UPDATED
'   2026-10-31
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Sheet       As Worksheet        'Trades
    Dim TopRow      As Long             'Its first data row
    Dim LastRow     As Long             'Its last row with a value
    Dim Grid        As Variant          'Its columns from Trade ID to Sub-class
    Dim RowCount    As Long             'Rows read
    Dim Spec        As CsvSpec          'The Trades contract
    Dim Columns()   As CsvColumn        'Its columns
    Dim Rules()     As CsvRule          'Its rules, unused here
    Dim Values()    As String           'The legs, with only those columns
    Dim Lines()     As Long             'No CSV line: 0 for every leg
    Dim ColID       As Long             'Contract column of the Trade ID
    Dim Result      As CsvReadResult    'What CheckContext found
    Dim Findings()  As CsvFinding       'Its findings
    Dim RowNumber   As Long             'Sheet row of a finding
    Dim SheetColumn As Long             'Its column on Trades
    Dim r           As Long             'Row or finding

'------------------------------------------------------------------------------
' READ THE LEGS
'------------------------------------------------------------------------------
        TopRow = LiveFirstRow(SH_TRADES, NM_TR_HEADER)
        Set Sheet = GetSheet(SH_TRADES)
        LastRow = LastDataRowAny(Sheet, TopRow, TR_NCOLS, TR_COMMENT)
        If LastRow < TopRow Then
            Exit Function
        End If
        RowCount = LastRow - TopRow + 1
        Grid = TableRange(Sheet, TopRow, TR_ID, LastRow, TR_SUB).Value2
        If Not CsvSpecOf("Trades", Spec, Columns, Rules) Then
            Err.Raise ERR_CSV_READ, "M_Main.CheckLiveContext", "the contract has no Trades file"
        End If
        ColID = ContractColumn("Trades", "trade_id")
        ReDim Values(1 To RowCount, 1 To Spec.ColumnCount)
        ReDim Lines(1 To RowCount)
        For r = 1 To RowCount
            Values(r, ColID) = SafeStr(Grid(r, 1 + TR_ID - TR_ID))
            Values(r, ContractColumn("Trades", "netting_set_id")) = UTxt(Grid(r, 1 + TR_NS - TR_ID))
            Values(r, ContractColumn("Trades", "asset_class")) = UTxt(Grid(r, 1 + TR_AC - TR_ID))
            Values(r, ContractColumn("Trades", "sub_class")) = UTxt(Grid(r, 1 + TR_SUB - TR_ID))
        Next r

'------------------------------------------------------------------------------
' CHECK
'------------------------------------------------------------------------------
        CheckContext Values, Lines, RowCount, Context, IDs, Regimes, IDCount, True, Result, Findings
        For r = 1 To Result.FindingCount
            RowNumber = TopRow + Findings(r).RecordNumber - 1
            SheetColumn = IIf(Findings(r).ColumnName = "sub_class", TR_SUB, TR_AC)
            AddStagedFinding Found, FoundCount, SH_TRADES, RowNumber, _
                             TableCell(Sheet, RowNumber, SheetColumn).Address(False, False), Findings(r).Code, _
                             Findings(r).Code & ": " & Findings(r).ColumnName & ": " & Findings(r).Message & _
                             " (leg " & Values(Findings(r).RecordNumber, ColID) & " on " & SH_TRADES & ")", False, Findings(r).ColumnName
        Next r
        CheckLiveContext = Result.FindingCount

End Function


Private Sub AddParamsFinding( _
    ByVal SheetName As String, _
    ByVal Reference As String, _
    ByVal RowNumber As Long, _
    ByVal Message As String, _
    ByRef Found() As StagedFinding, _
    ByRef FoundCount As Long)
'
'==============================================================================
'                               AddParamsFinding
'------------------------------------------------------------------------------
' PURPOSE
'   Add the finding of a run that would stop before the input sheets,
'   as CORE_Engine.LoadRunContext reports it (#209). Its cell is the one
'   named by the parameter code it names, or else the cell of its row
'   that holds its reference, or else the first cell of its row; with
'   neither a name nor a row it has no cell. It is never marked.
'
' UPDATED
'   2026-10-31
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Target      As Range        'Its cell; Nothing for none
    Dim Sheet       As Worksheet    'The sheet of its row
    Dim Text        As String       'Its message
    Dim c           As Long         'Column of its row

'------------------------------------------------------------------------------
' LOCATE
'------------------------------------------------------------------------------
        If Len(Reference) > 0 Then
            Set Target = NamedCell(Reference)
        End If
        If Not Target Is Nothing Then
            If Target.Worksheet.Name <> SheetName Or Target.Cells.Count <> 1 Then
                Set Target = Nothing
            End If
        End If
        If Target Is Nothing And RowNumber > 0 And Len(SheetName) > 0 Then
            On Error Resume Next
            Set Sheet = GetSheet(SheetName)
            On Error GoTo 0
            If Not Sheet Is Nothing Then
                Set Target = TableCell(Sheet, RowNumber, 1)
                For c = 1 To STAGING_ROW_SEARCH
                    If Len(Reference) > 0 And StrComp(SafeStr(TableCell(Sheet, RowNumber, c).Value2), Reference, _
                                                      vbTextCompare) = 0 Then
                        Set Target = TableCell(Sheet, RowNumber, c)
                        Exit For
                    End If
                Next c
            End If
        End If

'------------------------------------------------------------------------------
' ADD
'------------------------------------------------------------------------------
        Text = STAGING_PARAMS_CODE & ": "
        If Len(Reference) > 0 Then
            Text = Text & Reference & ": "
        End If
        Text = Text & Message & " - a run would stop here, so nothing that depends on Params was checked"
        If Target Is Nothing Then
            AddStagedFinding Found, FoundCount, SheetName, RowNumber, "", STAGING_PARAMS_CODE, Text, False, Reference
        Else
            AddStagedFinding Found, FoundCount, Target.Worksheet.Name, Target.Row, Target.Address(False, False), _
                             STAGING_PARAMS_CODE, Text, False, Reference
        End If

End Sub


Private Function CheckLiveTrades( _
    ByRef IDs() As String, _
    ByVal IDCount As Long, _
    ByRef Found() As StagedFinding, _
    ByRef FoundCount As Long) _
    As Long
'
'==============================================================================
'                               CheckLiveTrades
'------------------------------------------------------------------------------
' PURPOSE
'   When only netting sets are staged, check that each leg on Trades
'   refers to one of them, as CheckPortfolio checks a staged leg; a
'   finding is on the leg's Netting set cell on Trades, which is not
'   marked or changed.
'
' RETURNS
'   The number of findings added to Found.
'
' ERROR POLICY
'   Raises ERR_OUTPUT_LAYOUT when the Trades table cannot be located
'   through its header name.
'
' UPDATED
'   2026-10-28
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Sheet       As Worksheet        'Trades
    Dim TopRow      As Long             'Its first data row
    Dim LastRow     As Long             'Its last row with a value
    Dim Grid        As Variant          'Its Trade ID and Netting set columns
    Dim RowCount    As Long             'Rows read
    Dim Spec        As CsvSpec          'The Trades contract
    Dim Columns()   As CsvColumn        'Its columns
    Dim Rules()     As CsvRule          'Its rules, unused here
    Dim Values()    As String           'The legs, with only those two columns
    Dim Lines()     As Long             'No CSV line: 0 for every leg
    Dim ColID       As Long             'Contract column of the Trade ID
    Dim ColNS       As Long             'Contract column of the netting set
    Dim Result      As CsvReadResult    'What CheckPortfolio found
    Dim Findings()  As CsvFinding       'Its findings
    Dim RowNumber   As Long             'Sheet row of a finding
    Dim r           As Long             'Row or finding

'------------------------------------------------------------------------------
' READ THE LEGS
'------------------------------------------------------------------------------
        TopRow = LiveFirstRow(SH_TRADES, NM_TR_HEADER)
        Set Sheet = GetSheet(SH_TRADES)
        LastRow = LastDataRowAny(Sheet, TopRow, TR_NCOLS, TR_COMMENT)
        If LastRow < TopRow Then
            Exit Function
        End If
        RowCount = LastRow - TopRow + 1
        Grid = TableRange(Sheet, TopRow, TR_ID, LastRow, TR_NS).Value2
        If Not CsvSpecOf("Trades", Spec, Columns, Rules) Then
            Err.Raise ERR_CSV_READ, "M_Main.CheckLiveTrades", "the contract has no Trades file"
        End If
        ColID = ContractColumn("Trades", "trade_id")
        ColNS = ContractColumn("Trades", "netting_set_id")
        ReDim Values(1 To RowCount, 1 To Spec.ColumnCount)
        ReDim Lines(1 To RowCount)
        For r = 1 To RowCount
            Values(r, ColID) = SafeStr(Grid(r, 1))
            Values(r, ColNS) = SafeStr(Grid(r, 1 + TR_NS - TR_ID))
        Next r

'------------------------------------------------------------------------------
' CHECK
'------------------------------------------------------------------------------
        CheckPortfolio Values, Lines, RowCount, IDs, IDCount, Result, Findings, SH_IMPORT_NS
        For r = 1 To Result.FindingCount
            RowNumber = TopRow + Findings(r).RecordNumber - 1
            AddStagedFinding Found, FoundCount, SH_TRADES, RowNumber, _
                             TableCell(Sheet, RowNumber, TR_NS).Address(False, False), Findings(r).Code, _
                             Findings(r).Code & ": " & Findings(r).ColumnName & ": " & Findings(r).Message & _
                             " (leg " & Values(Findings(r).RecordNumber, ColID) & " on " & SH_TRADES & ")", False, Findings(r).ColumnName
        Next r
        CheckLiveTrades = Result.FindingCount

End Function


Private Function LiveFirstRow( _
    ByVal SheetName As String, _
    ByVal HeaderName As String) _
    As Long
'
'==============================================================================
' PURPOSE
'   The first data row of NettingSets or Trades; raises ERR_OUTPUT_LAYOUT
'   when its header name cannot be located.
'==============================================================================
'
        LiveFirstRow = TableFirstRow(SheetName)
        If LiveFirstRow = 0 Then
            Err.Raise ERR_OUTPUT_LAYOUT, "M_Main.LiveFirstRow", _
                      "The " & SheetName & " table cannot be located through the workbook name " & HeaderName & _
                      ": build the workbook from the current template."
        End If

End Function


Private Sub AddCsvFindings( _
    ByVal Kind As String, _
    ByVal Sheet As Worksheet, _
    ByVal FirstRow As Long, _
    ByRef Result As CsvReadResult, _
    ByRef Findings() As CsvFinding, _
    ByRef Found() As StagedFinding, _
    ByRef FoundCount As Long)
'
'==============================================================================
'                                AddCsvFindings
'------------------------------------------------------------------------------
' PURPOSE
'   Add the findings about a staging block to Found, each on its cell: the
'   row of its record and the column it names, or the CSV line column when
'   it names none.
'
' UPDATED
'   2026-10-28
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim RowNumber   As Long      'Sheet row of a finding
    Dim Column      As Long      'Its table column
    Dim Text        As String    'Its message
    Dim f           As Long      'Finding

'------------------------------------------------------------------------------
' ADD
'------------------------------------------------------------------------------
        For f = 1 To Result.FindingCount
            With Findings(f)
                RowNumber = FirstRow + .RecordNumber - 1
                Column = 1
                Text = .Code & ": "
                If Len(.ColumnName) > 0 Then
                    Column = 1 + ContractColumn(Kind, .ColumnName)
                    Text = Text & .ColumnName & ": "
                End If
                Text = Text & .Message
                If .LineNumber > 0 Then
                    Text = Text & " (CSV line " & .LineNumber & ")"
                Else
                    Text = Text & " (no CSV line: a row added here)"
                End If
                AddStagedFinding Found, FoundCount, Sheet.Name, RowNumber, _
                                 TableCell(Sheet, RowNumber, Column).Address(False, False), .Code, Text, True, .ColumnName
            End With
        Next f

End Sub


Private Sub AddStagedFinding( _
    ByRef Found() As StagedFinding, _
    ByRef FoundCount As Long, _
    ByVal SheetName As String, _
    ByVal RowNumber As Long, _
    ByVal CellAddress As String, _
    ByVal Code As String, _
    ByVal Message As String, _
    ByVal Marked As Boolean, _
    ByVal FieldName As String)
'
'==============================================================================
' PURPOSE
'   Append one finding to Found, doubling the array when full.
'==============================================================================
'
        FoundCount = FoundCount + 1
        If FoundCount > UBound(Found) Then
            ReDim Preserve Found(1 To 2 * UBound(Found))
        End If
        With Found(FoundCount)
            .SheetName = SheetName
            .RowNumber = RowNumber
            .CellAddress = CellAddress
            .Code = Code
            .FieldName = FieldName
            .Message = Message
            .Marked = Marked
        End With

End Sub


Private Function ContractColumn( _
    ByVal Kind As String, _
    ByVal ColumnName As String) _
    As Long
'
'==============================================================================
' PURPOSE
'   The position of a CSV column in its file's contract; raises
'   ERR_CSV_READ for a column the contract does not have.
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Spec        As CsvSpec      'The file's contract
    Dim Columns()   As CsvColumn    'Its columns
    Dim Rules()     As CsvRule      'Its rules
    Dim c           As Long         'Column

'------------------------------------------------------------------------------
' SEARCH
'------------------------------------------------------------------------------
        If CsvSpecOf(Kind, Spec, Columns, Rules) Then
            For c = 1 To Spec.ColumnCount
                If Columns(c).ColumnName = ColumnName Then
                    ContractColumn = c
                    Exit Function
                End If
            Next c
        End If
        Err.Raise ERR_CSV_READ, "M_Main.ContractColumn", "no column '" & ColumnName & "' in the " & Kind & " contract"

End Function


Private Function StagingSummary( _
    ByVal NsRows As Long, _
    ByVal NsErrors As Long, _
    ByVal TrRows As Long, _
    ByVal TrErrors As Long, _
    ByVal NsSource As String, _
    ByVal FoundCount As Long, _
    ByRef DialogText As String) _
    As String
'
'==============================================================================
' PURPOSE
'   The first Checks line of a staging validation: when, what was checked
'   and what was found; and in DialogText the same facts for its message,
'   one block per line, so that the message box does not wrap one long
'   sentence.
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Stamp   As String    'Date and time of the validation
    Dim Text    As String    'The Checks line
    Dim Dialog  As String    'The message, in blocks

'------------------------------------------------------------------------------
' WRITE BOTH
'------------------------------------------------------------------------------
    'Checks gets one line; the message the same facts, one block each.
        Stamp = Format$(Now, "yyyy-mm-dd hh:mm:ss")
        Text = "Staging validation " & Stamp & ": "
        Dialog = "Staging validation " & Stamp & vbLf
        If NsRows > 0 Then
            Text = Text & SH_IMPORT_NS & " " & NsRows & " rows, " & NsErrors & " errors"
            Dialog = Dialog & vbLf & SH_IMPORT_NS & ": " & CountText(NsRows, "row") & ", " & CountText(NsErrors, "error")
            If TrRows = 0 Then
                Text = Text & ", with the legs on " & SH_TRADES
                Dialog = Dialog & vbLf & "    checked with the legs on " & SH_TRADES
            End If
        End If
        If TrRows > 0 Then
            If NsRows > 0 Then
                Text = Text & "; "
            End If
            Text = Text & SH_IMPORT_TR & " " & TrRows & " rows, " & TrErrors & " errors, against the netting sets on " & _
                   NsSource
            Dialog = Dialog & vbLf & SH_IMPORT_TR & ": " & CountText(TrRows, "row") & ", " & CountText(TrErrors, "error") & _
                     vbLf & "    checked against the netting sets on " & NsSource
        End If
        Text = Text & ". "
        If FoundCount = 0 Then
            Text = Text & "No issues found."
            Dialog = Dialog & vbLf & vbLf & "No issues found. Commit import replaces the input sheets with these rows."
        Else
            If FoundCount > STAGING_LISTED_MAX Then
                Text = Text & "The first " & Format$(STAGING_LISTED_MAX, "#,##0") & " of " & Format$(FoundCount, "#,##0") & _
                       " findings are listed below."
                Dialog = Dialog & vbLf & vbLf & "The first " & Format$(STAGING_LISTED_MAX, "#,##0") & " of " & _
                         Format$(FoundCount, "#,##0") & " findings are listed on Checks."
            Else
                Dialog = Dialog & vbLf & vbLf & CountText(FoundCount, "finding") & " listed on Checks."
            End If
            Dialog = Dialog & vbLf & "Each line links to its cell, which is marked on its staging sheet."
        End If
        StagingSummary = Text & " Nothing was copied to " & SH_NS & " or " & SH_TRADES & _
                         "; Run SA-CCR writes the checks of a run here again."
        DialogText = Dialog & vbLf & vbLf & "Nothing was copied to " & SH_NS & " or " & SH_TRADES & "." & vbLf & _
                     "Run SA-CCR writes the checks of a run on Checks again."

End Function


Private Function CountText( _
    ByVal Count As Long, _
    ByVal Noun As String) _
    As String
'
'==============================================================================
'                                  CountText
'------------------------------------------------------------------------------
' PURPOSE
'   A count with its noun, singular for 1 and plural otherwise, for the
'   messages: "1 row", "0 errors", "1,250 rows".
'
' UPDATED
'   2026-11-03
'==============================================================================
'
        CountText = Format$(Count, "#,##0") & " " & Noun & IIf(Count = 1, "", "s")

End Function


Private Sub WriteStagingChecks( _
    ByRef Found() As StagedFinding, _
    ByVal FoundCount As Long, _
    ByVal Summary As String)
'
'==============================================================================
'                              WriteStagingChecks
'------------------------------------------------------------------------------
' PURPOSE
'   Rewrite Checks with a staging validation: the summary as an INFO line,
'   then each finding, up to STAGING_LISTED_MAX, as an ERROR line in bold
'   with its sheet, row, a HYPERLINK formula to its cell in Reference, its
'   message and the origin IMPORT. A formula, unlike a hyperlink object,
'   goes with the value when the next run clears Checks.
'
' UPDATED
'   2026-10-28
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ChecksSheet   As Worksheet    'Checks
    Dim TopRow        As Long         'Its first data row
    Dim Listed        As Long         'Findings listed
    Dim Block()       As Variant      'The lines
    Dim i             As Long         'Finding

'------------------------------------------------------------------------------
' WRITE
'------------------------------------------------------------------------------
        Set ChecksSheet = GetSheet(SH_CHECKS)
        TopRow = OutputFirstRow(SH_CHECKS)
        ClearChecksBlock ChecksSheet, TopRow
        Listed = FoundCount
        If Listed > STAGING_LISTED_MAX Then
            Listed = STAGING_LISTED_MAX
        End If
        ReDim Block(1 To Listed + 1, 1 To CK_NCOLS)
        Block(1, 1) = SEV_INFO
        Block(1, 5) = Summary
        Block(1, CK_ORIGIN_COL) = ORIGIN_IMPORT
        Block(1, CK_CODE_COL) = DIAG_IMPORT_SUMMARY
        For i = 1 To Listed
            With Found(i)
                Block(i + 1, 1) = SEV_ERROR
                Block(i + 1, 2) = .SheetName
                If .RowNumber > 0 Then
                    Block(i + 1, 3) = .RowNumber
                End If
                If Len(.CellAddress) > 0 Then
                    Block(i + 1, 4) = "=HYPERLINK(""#'" & .SheetName & "'!" & .CellAddress & """,""" & .CellAddress & """)"
                End If
                Block(i + 1, 5) = .Message
                Block(i + 1, CK_ORIGIN_COL) = ORIGIN_IMPORT
                Block(i + 1, CK_CODE_COL) = .Code
                Block(i + 1, CK_FIELD_COL) = .FieldName
            End With
        Next i
        WriteBlock ChecksSheet, TopRow, Block, Listed + 1, CK_NCOLS
        If Listed > 0 Then
            TableRange(ChecksSheet, TopRow + 1, 1, TopRow + Listed, 1).Font.Bold = True
        End If

End Sub


Private Sub UnmarkStaging( _
    ByVal Kind As String)
'
'==============================================================================
'                                UnmarkStaging
'------------------------------------------------------------------------------
' PURPOSE
'   Take the marks of the last staging validation off a staging sheet:
'   every cell below the header with the error fill gets the fill of its
'   column back. Other fills are left as they are.
'
' UPDATED
'   2026-10-28
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Sheet       As Worksheet    'The staging sheet
    Dim FirstRow    As Long         'First row below its header
    Dim LastRow     As Long         'Last row of its used range
    Dim c           As Long         'Column

'------------------------------------------------------------------------------
' UNMARK
'------------------------------------------------------------------------------
        FirstRow = StagingHeaderRow(Kind, Sheet) + 1
        LastRow = UsedLastRow(Sheet)
        If LastRow < FirstRow Then
            Exit Sub
        End If
        Recolour TableRange(Sheet, FirstRow, 1, LastRow, 1), STAGING_LINE_FILL
        For c = 2 To StagingWidth(Kind)
            Recolour TableRange(Sheet, FirstRow, c, LastRow, c), STAGING_VALUE_FILL
        Next c

End Sub


Private Sub Recolour( _
    ByVal Target As Range, _
    ByVal NewFill As Long)
'
'==============================================================================
' PURPOSE
'   Give NewFill to the cells of a one-column range that have the error
'   fill. A range of one fill is read at once; a range of mixed fills is
'   split in halves, so that a few marks among many rows cost few reads.
'   Mixed fills are found by Interior.ColorIndex, which is Null for them:
'   Interior.Color may give a number instead of Null.
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Fill    As Variant    'The range's fill; Null when mixed
    Dim Half    As Long       'Rows of its first half

'------------------------------------------------------------------------------
' RECOLOUR
'------------------------------------------------------------------------------
        Fill = Null
        If Not IsNull(Target.Interior.ColorIndex) Then
            Fill = Target.Interior.Color
        End If
        If IsNull(Fill) Then
            Half = Target.Rows.Count \ 2
            Recolour Target.Resize(Half), NewFill
            Recolour Target.Offset(Half).Resize(Target.Rows.Count - Half), NewFill
        ElseIf CLng(Fill) = STAGING_ERROR_FILL Then
            Target.Interior.Color = NewFill
        End If

End Sub


Private Sub MarkStaging( _
    ByRef Found() As StagedFinding, _
    ByVal FoundCount As Long)
'
'==============================================================================
' PURPOSE
'   Give the error fill to the staged cell of each finding listed on
'   Checks; a finding on a live sheet is not marked.
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim i   As Long    'Finding

'------------------------------------------------------------------------------
' MARK
'------------------------------------------------------------------------------
        For i = 1 To FoundCount
            If i > STAGING_LISTED_MAX Then
                Exit For
            End If
            If Found(i).Marked Then
                GetSheet(Found(i).SheetName).Range(Found(i).CellAddress).Interior.Color = STAGING_ERROR_FILL
            End If
        Next i

End Sub


Private Function ChecksHasImportLines() As Boolean
'
'==============================================================================
' PURPOSE
'   Whether Checks holds the lines of a staging validation (origin
'   IMPORT), for the reason the evidence export gives.
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ChecksSheet   As Worksheet    'Checks
    Dim TopRow        As Long         'Its first data row
    Dim LastRow       As Long         'Its last line
    Dim Origins       As Variant      'The Origin column
    Dim r             As Long         'Line

'------------------------------------------------------------------------------
' SEARCH
'------------------------------------------------------------------------------
        Set ChecksSheet = GetSheet(SH_CHECKS)
        TopRow = OutputFirstRow(SH_CHECKS)
        LastRow = LastDataRow(ChecksSheet, TopRow, 1)
        If LastRow < TopRow Then
            Exit Function
        End If
        Origins = TableRange(ChecksSheet, TopRow, CK_ORIGIN_COL, LastRow + 1, CK_ORIGIN_COL).Value2
        For r = 1 To LastRow - TopRow + 1
            If SafeStr(Origins(r, 1)) = ORIGIN_IMPORT Then
                ChecksHasImportLines = True
                Exit Function
            End If
        Next r

End Function


'
'------------------------------------------------------------------------------
'
'                         REGULATORY PARAMETERS (#357)
'
'------------------------------------------------------------------------------
'

Public Sub RestoreRegulatoryParameters()
'
'==============================================================================
'                         RestoreRegulatoryParameters
'------------------------------------------------------------------------------
' PURPOSE
'   Put the reference parameters on Params and the supervisory-factor
'   table back to the regulatory set REF_SET_ID (#357), after a
'   confirmation that lists what will be overwritten, with Cancel as the
'   default. The reporting date, the reporting currency, the default
'   regime, the FX table and the netting sets' alpha overrides are inputs
'   and are not changed.
'
' USAGE
'   Restore parameters on the K-SACCR Ribbon tab, or the Macros dialog.
'   Automation calls RestoreRegulatoryParameters_Silent.
'
' ERROR POLICY
'   Reports an error in a message, or in the Immediate window in silent
'   mode; never raises. In silent mode, where no confirmation can be
'   given, nothing is restored.
'
' UPDATED
'   2026-11-06
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Changes     As String    'What would be overwritten, one line each
    Dim Problem     As String    'Why nothing can be restored; "" when it can
    Dim Count       As Long      'Cells or rows that differ from the set
    Dim ResultLine  As String    'What the restore returned
    Dim Message     As String    'Its message

'------------------------------------------------------------------------------
' CONFIRM, RESTORE AND REPORT
'------------------------------------------------------------------------------
        On Error GoTo Failed
        If mRunning Then
            Notify "An SA-CCR operation is already running.", vbExclamation
            Exit Sub
        End If
        If gSilent Then
            Notify "Nothing was restored." & vbLf & vbLf & "A restore is confirmed in a dialog, which silent mode " & _
                   "does not show. Automation calls RestoreRegulatoryParameters_Silent.", vbExclamation
            Exit Sub
        End If
        Count = RestoreDifferences(Changes, Problem)
        If Len(Problem) > 0 Then
            Notify "Nothing was restored." & vbLf & vbLf & Problem, vbExclamation
            Exit Sub
        End If
        If Count = 0 Then
            Notify UnchangedText(), vbInformation
            Exit Sub
        End If
        If MsgBox(RestoreQuestion(Changes, Count), vbOKCancel + vbExclamation + vbDefaultButton2, MSG_TITLE) <> vbOK Then
            Exit Sub
        End If
        ResultLine = RestoreParameters(Message)
        RibbonStateChanged
        If Left$(ResultLine, Len("RESULT=RESTORED")) = "RESULT=RESTORED" Or _
           Left$(ResultLine, Len("RESULT=UNCHANGED")) = "RESULT=UNCHANGED" Then
            Notify Message, vbInformation
        ElseIf InStr(1, ResultLine, "rollback=FAILED", vbBinaryCompare) > 0 Then
            Notify Message, vbCritical
        Else
            Notify Message, vbExclamation
        End If
        Exit Sub

Failed:
        Notify "Restore regulatory parameters: " & Err.Description, vbCritical

End Sub


Public Function RestoreRegulatoryParameters_Silent() As String
'
'==============================================================================
'                      RestoreRegulatoryParameters_Silent
'------------------------------------------------------------------------------
' PURPOSE
'   Restore the regulatory parameters (#357), for automation and tests,
'   without a confirmation or a message.
'
' RETURNS
'   "RESULT=RESTORED; parameters=N; factor_rows=M", N and M being the
'   Params values and factor-table rows that differed and were written;
'   "RESULT=UNCHANGED" when Params already holds the set;
'   "RESULT=REFUSED; reason=..." when nothing was written; or
'   "RESULT=FAILED; error=...; rollback=..." when a write failed, with
'   rollback=NONE when nothing had been written, VERIFIED when Params was
'   put back and read back as it was, and FAILED: ... when it could not
'   be.
'
' ERROR POLICY
'   Raises ERR_RUN_ACTIVE while an operation runs, and ERR_CLEANUP_FAILED
'   when the restore was done but Excel settings were not restored.
'
' UPDATED
'   2026-11-06
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Message     As String    'The message, not shown

'------------------------------------------------------------------------------
' RESTORE
'------------------------------------------------------------------------------
        RestoreRegulatoryParameters_Silent = RestoreParameters(Message)

End Function


Private Function RestoreParameters( _
    ByRef Message As String) _
    As String
'
'==============================================================================
'                              RestoreParameters
'------------------------------------------------------------------------------
' PURPOSE
'   Write the regulatory set REF_SET_ID back to Params (#357): each value
'   of CORE_Engine.RegulatoryParameterCodes, and the rows of
'   CORE_Engine.RegulatoryFactorRows into the factor table, below its
'   header, in the columns from the key to the regimes. The Reference
'   column, the other parameters, the FX table and every other sheet are
'   not touched.
'
' METHOD
'   1. Find the cells (RestoreTargets) and refuse, writing nothing, when
'      a Params name is missing or the factor table has a key below the
'      regulatory rows. When every cell already holds its value, write
'      nothing and say so.
'   2. Keep the values and formulas of the cells that differ, write the
'      set into them through the macros' protection, with events off and
'      calculation manual, and read each cell back.
'   3. On any failure, write the kept cells back and read them back; the
'      result tells whether that was verified, apart from the error
'      itself.
'   4. Results, Checks and the run fingerprints are left as they are:
'      Results reads OUT OF DATE until the next run, as after any edit,
'      and the evidence export refuses until then. The status cells are
'      recalculated before Excel settings are restored, so that they read
'      OUT OF DATE in manual calculation too.
'
' ERROR POLICY
'   See RestoreRegulatoryParameters_Silent. gTestFault = "restore-write"
'   raises ERR_INJECTED_FAULT once the first cell is written;
'   "restore-verify" makes the read-back fail; "restore-rollback" raises
'   it there too and makes the write-back fail as well.
'
' UPDATED
'   2026-11-06
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Targets()           As Range            'The cells of the set: parameters, then factor rows
    Dim Expected()          As Variant          'What each holds in the set, by row and column
    Dim TargetCount         As Long             'Parameters and factor rows
    Dim ParameterCount      As Long             'Of which parameters
    Dim Problem             As String           'Why nothing can be restored
    Dim Differs()           As Boolean          'The target does not hold the set
    Dim KeptValues()        As Variant          'Value2 of each differing target
    Dim KeptFormulas()      As Variant          'Its formulas
    Dim Touched()           As Boolean          'The target may have been written
    Dim ChangedParameters   As Long             'Parameters that differed
    Dim ChangedRows         As Long             'Factor rows that differed
    Dim k                   As Long             'Target
    Dim SavedCalculation    As XlCalculation    'Calculation mode before the restore
    Dim SavedEvents         As Boolean          'EnableEvents before the restore
    Dim SavedScreen         As Boolean          'ScreenUpdating before the restore
    Dim StateCaptured       As Boolean          'True once all three were read
    Dim Restored            As Boolean          'Every target was written and read back
    Dim ErrorNumber         As Long             'Primary error number; 0 for none
    Dim ErrorDescription    As String           'Primary error description
    Dim Rollback            As String           'NONE, VERIFIED or FAILED: details
    Dim CleanupDetails      As String           'Settings not restored

'------------------------------------------------------------------------------
' PLAN
'------------------------------------------------------------------------------
        If mRunning Then
            Err.Raise ERR_RUN_ACTIVE, "M_Main.RestoreParameters", "An SA-CCR operation is already running."
        End If
        TargetCount = RestoreTargets(Targets, Expected, ParameterCount, Problem)
        If Len(Problem) > 0 Then
            RestoreParameters = "RESULT=REFUSED; reason=" & Problem
            Message = "Nothing was restored." & vbLf & vbLf & Problem
            Exit Function
        End If
        ReDim Differs(1 To TargetCount)
        ReDim KeptValues(1 To TargetCount)
        ReDim KeptFormulas(1 To TargetCount)
        ReDim Touched(1 To TargetCount)
        For k = 1 To TargetCount
            Differs(k) = Not HoldsExpected(Targets(k), Expected(k))
            If Differs(k) Then
                If k <= ParameterCount Then
                    ChangedParameters = ChangedParameters + 1
                Else
                    ChangedRows = ChangedRows + 1
                End If
            End If
        Next k
        If ChangedParameters + ChangedRows = 0 Then
            RestoreParameters = "RESULT=UNCHANGED"
            Message = UnchangedText()
            Exit Function
        End If

'------------------------------------------------------------------------------
' KEEP, WRITE AND READ BACK
'------------------------------------------------------------------------------
        mRunning = True
        ProtectWorkbook
        On Error GoTo Failed
        For k = 1 To TargetCount
            If Differs(k) Then
                KeptValues(k) = BlockValues(Targets(k))
                KeptFormulas(k) = FormulasOnly(Targets(k))
            End If
        Next k
        SavedCalculation = Application.Calculation
        SavedEvents = Application.EnableEvents
        SavedScreen = Application.ScreenUpdating
        StateCaptured = True
        Application.ScreenUpdating = False
        Application.EnableEvents = False
        Application.Calculation = xlCalculationManual
        For k = 1 To TargetCount
            If Differs(k) Then
                Touched(k) = True
                Targets(k).Value = Expected(k)
                If gTestFault = "restore-write" Or gTestFault = "restore-rollback" Then
                    Err.Raise ERR_INJECTED_FAULT, "M_Main.RestoreParameters", "Injected restore write failure."
                End If
            End If
        Next k
        If gTestFault = "restore-verify" Then
            Err.Raise ERR_PARAMS_WRITE, "M_Main.RestoreParameters", "Injected read-back difference on " & SH_PARAMS
        End If
        For k = 1 To TargetCount
            If Differs(k) Then
                If Not HoldsExpected(Targets(k), Expected(k)) Then
                    Err.Raise ERR_PARAMS_WRITE, "M_Main.RestoreParameters", SH_PARAMS & "!" & _
                              Targets(k).Address(False, False) & " does not read as written"
                End If
            End If
        Next k
        Restored = True

'------------------------------------------------------------------------------
' RESTORE EXCEL AND REPORT
'------------------------------------------------------------------------------
    'The status cells are recalculated first: in manual calculation they
    'would keep the value computed before the restore and still report the
    'last run as current.
CleanUp:
        On Error GoTo 0
        If StateCaptured Then
            RefreshResultsStatus
            CleanupDetails = RestoreSettings(SavedCalculation, SavedEvents, SavedScreen)
        End If
        mRunning = False
        If Not Restored Then
            RestoreParameters = "RESULT=FAILED; error=" & ErrorDescription & "; rollback=" & Rollback
            Message = "Nothing was restored." & vbLf & vbLf & ErrorDescription & " (error " & ErrorNumber & ")."
            If Left$(Rollback, 6) = "FAILED" Then
                Message = Message & vbLf & vbLf & SH_PARAMS & " could not be put back as it was: " & _
                          Mid$(Rollback, 9) & ". Check " & SH_PARAMS & " before any run."
            ElseIf Rollback = "VERIFIED" Then
                Message = Message & vbLf & vbLf & SH_PARAMS & " was put back as it was and checked."
            End If
            If Len(CleanupDetails) > 0 Then
                Message = Message & vbLf & vbLf & "Excel settings were not restored: " & CleanupDetails
            End If
            Exit Function
        End If
        If Len(CleanupDetails) > 0 Then
            Err.Raise ERR_CLEANUP_FAILED, "M_Main.RestoreParameters", _
                      "The regulatory parameters were restored, but Excel settings were not restored: " & CleanupDetails
        End If
        RestoreParameters = "RESULT=RESTORED; parameters=" & ChangedParameters & "; factor_rows=" & ChangedRows
        Message = "Restored the regulatory parameter set " & REF_SET_ID & "." & vbLf & vbLf & _
                  "Parameters changed: " & ChangedParameters & vbLf & _
                  "Factor-table rows changed: " & ChangedRows & vbLf & vbLf & _
                  "The reporting date, currency, regime, FX rates and the netting sets' alpha overrides " & _
                  "were not changed. Results reads OUT OF DATE until the next Run SA-CCR."
        Exit Function

Failed:
        ErrorNumber = Err.Number
        ErrorDescription = Err.Description
        Rollback = "NONE"
        For k = 1 To TargetCount
            If Touched(k) Then
                Rollback = ""
            End If
        Next k
        If Len(Rollback) = 0 Then
            For k = 1 To TargetCount
                If Touched(k) Then
                    Rollback = Rollback & PutRangeBack(Targets(k), KeptValues(k), KeptFormulas(k))
                End If
            Next k
            If Len(Rollback) = 0 Then
                Rollback = "VERIFIED"
            Else
                Rollback = "FAILED: " & Mid$(Rollback, 3)
            End If
        End If
        Resume CleanUp

End Function


Private Function RestoreTargets( _
    ByRef Targets() As Range, _
    ByRef Expected() As Variant, _
    ByRef ParameterCount As Long, _
    ByRef Problem As String) _
    As Long
'
'==============================================================================
'                                RestoreTargets
'------------------------------------------------------------------------------
' PURPOSE
'   The cells Restore regulatory parameters writes (#357), and what each
'   holds in the regulatory set: first the cell of each code of
'   CORE_Engine.RegulatoryParameterCodes, then one row of the factor
'   table, key to regimes, for each row of CORE_Engine.RegulatoryFactorRows,
'   from the row below the table's header. A factor row is written as
'   Params shows it: blank fields blank, blank regimes as BOTH, the
'   numbers as numbers.
'
' RETURNS
'   The number of targets, ParameterCount of them parameters; 0 with
'   Problem set when a Params name is missing or not on Params, the
'   header is not where the name says, or a factor row stands below the
'   regulatory rows: a key right below them, which the loader would read
'   as part of the table, or, from that row down to the row above the FX
'   table, a row with a number in the factor column that is not a
'   parameter row, which the loader rejects as a row after the end
'   (ValuesAfterEnd). The restore would leave either, and the next run
'   would not be the template's.
'
' UPDATED
'   2026-11-06
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ParamsSheet     As Worksheet    'Params
    Dim Codes           As Variant      'Parameter codes
    Dim FactorRows      As Variant      'The set's factor rows
    Dim Fields          As Variant      'One row's fields
    Dim Cell            As Range        'A parameter cell
    Dim HeaderRow       As Long         'Header row of the factor table
    Dim KeyColumn       As Long         'Its key column, as TableCell counts
    Dim RowValues()     As Variant      'One row as written
    Dim BelowKey        As Variant      'The key cell below the regulatory rows
    Dim FirstBelow      As Long         'Row below the regulatory rows
    Dim LastBelow       As Long         'Last row searched for a factor row
    Dim r               As Long         'Row searched
    Dim k               As Long         'Target
    Dim i               As Long         'Code or row
    Dim c               As Long         'Field

'------------------------------------------------------------------------------
' PARAMETERS
'------------------------------------------------------------------------------
        Set ParamsSheet = GetSheet(SH_PARAMS)
        Codes = CORE_Engine.RegulatoryParameterCodes()
        FactorRows = CORE_Engine.RegulatoryFactorRows()
        ParameterCount = UBound(Codes) - LBound(Codes) + 1
        ReDim Targets(1 To ParameterCount + UBound(FactorRows) - LBound(FactorRows) + 1)
        ReDim Expected(1 To UBound(Targets))
        For i = LBound(Codes) To UBound(Codes)
            Set Cell = NamedCell(CStr(Codes(i)))
            If Cell Is Nothing Then
                Problem = "the workbook name '" & Codes(i) & "' is missing or no longer refers to a cell on " & SH_PARAMS
                Exit Function
            ElseIf Cell.Worksheet.Name <> ParamsSheet.Name Then
                Problem = "the workbook name '" & Codes(i) & "' must refer to a cell on " & SH_PARAMS
                Exit Function
            End If
            k = k + 1
            Set Targets(k) = Cell.Cells(1, 1)
            Expected(k) = CORE_Engine.RegulatoryParameterValue(CStr(Codes(i)))
        Next i

'------------------------------------------------------------------------------
' FACTOR TABLE
'------------------------------------------------------------------------------
        HeaderRow = NamedRow(NM_SF_TABLE)
        If HeaderRow = 0 Then
            Problem = "the workbook name '" & NM_SF_TABLE & "' is missing or no longer refers to a cell on " & SH_PARAMS
            Exit Function
        End If
        KeyColumn = NamedTableCol(NM_SF_TABLE)
        If NamedCell(NM_SF_TABLE).Worksheet.Name <> ParamsSheet.Name Or _
           UCase$(SafeStr(TableCell(ParamsSheet, HeaderRow, KeyColumn).Value2)) <> UCase$(HDR_SF) Then
            Problem = "the factor table's header " & HDR_SF & " is not where the workbook name '" & NM_SF_TABLE & "' says"
            Exit Function
        End If
        FirstBelow = HeaderRow + UBound(FactorRows) - LBound(FactorRows) + 2
        BelowKey = TableCell(ParamsSheet, FirstBelow, KeyColumn).Value2
        If Len(Trim$(SafeStr(BelowKey))) > 0 Then
            Problem = "the factor table holds a row below its " & UBound(FactorRows) - LBound(FactorRows) + 1 & _
                      " regulatory rows (" & SafeStr(BelowKey) & "); remove it by hand after a maintenance unlock"
            Exit Function
        End If
    'From that row down, as CORE_Engine.LoadSFTable searches from the blank
    'key that ends the table: to the row above the FX table when that
    'follows, otherwise to the last used row.
        LastBelow = UsedLastRow(ParamsSheet)
        If NamedRow(NM_FX_TABLE) > HeaderRow Then
            LastBelow = NamedRow(NM_FX_TABLE) - 1
        End If
        For r = FirstBelow To LastBelow
            If IsNum(TableCell(ParamsSheet, r, KeyColumn + 3).Value) And _
               Not IsListedCode(SafeStr(TableCell(ParamsSheet, r, ParamsCodeCol()).Value)) Then
                Problem = "the factor table holds a row below its " & UBound(FactorRows) - LBound(FactorRows) + 1 & _
                          " regulatory rows (" & SH_PARAMS & " row " & r & "); remove it by hand after a " & _
                          "maintenance unlock"
                Exit Function
            End If
        Next r
        For i = LBound(FactorRows) To UBound(FactorRows)
            Fields = Split(FactorRows(i), "|")
            ReDim RowValues(1 To 1, 1 To 8)
            For c = 0 To 7
                Select Case c
                    Case 3, 4, 5
                        If Len(Fields(c)) > 0 Then
                            RowValues(1, c + 1) = Val(Fields(c))
                        End If
                    Case 7
                        RowValues(1, c + 1) = IIf(Len(Fields(c)) = 0, "BOTH", Fields(c))
                    Case Else
                        If Len(Fields(c)) > 0 Then
                            RowValues(1, c + 1) = CStr(Fields(c))
                        End If
                End Select
            Next c
            k = k + 1
            Set Targets(k) = TableRange(ParamsSheet, HeaderRow + 1 + i - LBound(FactorRows), KeyColumn, _
                                        HeaderRow + 1 + i - LBound(FactorRows), KeyColumn + 7)
            Expected(k) = RowValues
        Next i
        RestoreTargets = k

End Function


Private Function IsListedCode( _
    ByVal Candidate As String) _
    As Boolean
'
'==============================================================================
'                                 IsListedCode
'------------------------------------------------------------------------------
' PURPOSE
'   Whether Candidate is a Params code of CORE_Engine.ParameterCodes,
'   ignoring case, so that RestoreTargets tells a parameter row from a
'   factor row, as the engine's loader does (#357).
'
' UPDATED
'   2026-11-06
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Code    As Variant    'A supported code

'------------------------------------------------------------------------------
' SEARCH
'------------------------------------------------------------------------------
        For Each Code In CORE_Engine.ParameterCodes()
            If StrComp(Candidate, CStr(Code), vbTextCompare) = 0 Then
                IsListedCode = True
                Exit Function
            End If
        Next Code

End Function


Private Function HoldsExpected( _
    ByVal Target As Range, _
    ByVal Expected As Variant) _
    As Boolean
'
'==============================================================================
'                                HoldsExpected
'------------------------------------------------------------------------------
' PURPOSE
'   Whether a restore target holds its value in the regulatory set (#357):
'   a constant equal to it in every cell, compared as SameCellValue does,
'   with no formula.
'
' INPUTS
'   Target: one parameter cell, or one factor row.
'   Expected: a value for a cell, an array by row and column for a row.
'
' UPDATED
'   2026-11-06
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Grid        As Variant    'Target's Value2
    Dim Formulas    As Variant    'Its formulas
    Dim c           As Long       'Column

'------------------------------------------------------------------------------
' COMPARE
'------------------------------------------------------------------------------
        Grid = BlockValues(Target)
        Formulas = FormulasOnly(Target)
        If IsArray(Expected) Then
            For c = 1 To Target.Columns.Count
                If Len(CStr(Formulas(1, c))) > 0 Or Not SameCellValue(Grid(1, c), Expected(1, c)) Then
                    Exit Function
                End If
            Next c
            HoldsExpected = True
        Else
            HoldsExpected = Len(CStr(Formulas(1, 1))) = 0 And SameCellValue(Grid(1, 1), Expected)
        End If

End Function


Private Function RestoreDifferences( _
    ByRef Changes As String, _
    ByRef Problem As String) _
    As Long
'
'==============================================================================
'                              RestoreDifferences
'------------------------------------------------------------------------------
' PURPOSE
'   What a restore would overwrite (#357), for its confirmation: a line
'   for each parameter that differs from the set, with its value now and
'   in the set, and for each factor row that differs, by its key in the
'   set. At most RESTORE_LISTED lines are listed; a last line counts the
'   others.
'
' RETURNS
'   The number of parameters and rows that differ; 0 with Problem set
'   when the restore would be refused.
'
' UPDATED
'   2026-11-06
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Targets()       As Range       'The cells of the set
    Dim Expected()      As Variant     'What each holds in the set
    Dim TargetCount     As Long        'Their number
    Dim ParameterCount  As Long        'Of which parameters
    Dim Codes           As Variant     'Parameter codes
    Dim FactorRows      As Variant     'The set's factor rows
    Dim Listed          As Long        'Lines listed
    Dim ChangeLine      As String      'One line
    Dim k               As Long        'Target

'------------------------------------------------------------------------------
' COMPARE
'------------------------------------------------------------------------------
        TargetCount = RestoreTargets(Targets, Expected, ParameterCount, Problem)
        If Len(Problem) > 0 Then
            Exit Function
        End If
        Codes = CORE_Engine.RegulatoryParameterCodes()
        FactorRows = CORE_Engine.RegulatoryFactorRows()
        For k = 1 To TargetCount
            If Not HoldsExpected(Targets(k), Expected(k)) Then
                RestoreDifferences = RestoreDifferences + 1
                If k <= ParameterCount Then
                    ChangeLine = Codes(LBound(Codes) + k - 1) & ": " & ShownValue(Targets(k)) & " -> " & CStr(Expected(k))
                Else
                    ChangeLine = "Factor-table row " & Split(FactorRows(LBound(FactorRows) + k - ParameterCount - 1), "|")(0) & _
                           " (" & SH_PARAMS & " row " & Targets(k).Row & ")"
                End If
                If Listed < RESTORE_LISTED Then
                    Listed = Listed + 1
                    Changes = Changes & vbLf & "  " & ChangeLine
                End If
            End If
        Next k
        If RestoreDifferences > Listed Then
            Changes = Changes & vbLf & "  and " & Format$(RestoreDifferences - Listed, "#,##0") & " more"
        End If

End Function


Private Function ShownValue( _
    ByVal Cell As Range) _
    As String
'
'==============================================================================
'                                  ShownValue
'------------------------------------------------------------------------------
' PURPOSE
'   A parameter cell's content for the confirmation of a restore (#357):
'   its value, "blank", or "a formula" for a formula cell.
'
' UPDATED
'   2026-11-06
'==============================================================================
'
        If Cell.HasFormula Then
            ShownValue = "a formula"
        ElseIf IsError(Cell.Value2) Then
            ShownValue = CStr(Cell.Text)
        ElseIf Len(SafeStr(Cell.Value2)) = 0 Then
            ShownValue = "blank"
        Else
            ShownValue = CStr(Cell.Value2)
        End If

End Function


Private Function RestoreQuestion( _
    ByVal Changes As String, _
    ByVal Count As Long) _
    As String
'
'==============================================================================
'                               RestoreQuestion
'------------------------------------------------------------------------------
' PURPOSE
'   The confirmation of a restore (#357): what will be overwritten, what
'   is left as it is, and the effect on Results.
'
' UPDATED
'   2026-11-06
'==============================================================================
'
        RestoreQuestion = "Restore the regulatory parameter set " & REF_SET_ID & "?" & vbLf & vbLf & _
                          "These values on " & SH_PARAMS & " will be overwritten (" & Count & "):" & Changes & _
                          vbLf & vbLf & "Not changed: the reporting date, currency and regime, the FX rates and " & _
                          "the netting sets' alpha overrides." & vbLf & vbLf & _
                          "Results will read OUT OF DATE until the next Run SA-CCR."

End Function


Private Function UnchangedText() As String
'
'==============================================================================
'                                UnchangedText
'------------------------------------------------------------------------------
' PURPOSE
'   The message of a restore that finds nothing to change (#357).
'
' UPDATED
'   2026-11-06
'==============================================================================
'
        UnchangedText = "Nothing was changed: the reference parameters and the factor table on " & SH_PARAMS & _
                        " already hold the regulatory parameter set " & REF_SET_ID & "."

End Function


'
'------------------------------------------------------------------------------
'
'                              RIBBON STATE (#216)
'
'------------------------------------------------------------------------------
'

Public Function OperationRunning() As Boolean
'
'==============================================================================
'                               OperationRunning
'------------------------------------------------------------------------------
' PURPOSE
'   Whether an operation is in progress, for the Ribbon's enabled states
'   (M_Ribbon.RibbonGetEnabled, #216). Read only: the flag mRunning stays
'   this module's.
'
' UPDATED
'   2026-11-06
'==============================================================================
'
        OperationRunning = mRunning

End Function


Public Sub RibbonStateChanged()
'
'==============================================================================
'                             RibbonStateChanged
'------------------------------------------------------------------------------
' PURPOSE
'   Ask the K-SACCR tab to read its states and labels again once Excel is
'   idle (M_Ribbon.ScheduleRefresh, #216), after a macro that can change
'   Results or a staging status, whether the Ribbon, a staging button,
'   the Macros dialog or a change handler started it, after showing or
'   hiding the navigation bar (ShowNavigationBar, #620), and after any cell
'   edit (ThisWorkbook.Workbook_SheetChange): an input edit recalculates
'   Results C5, which decides Export evidence. Nothing is scheduled in
'   silent mode or while Excel is hidden, as under automation, which may
'   close the workbook at once: Excel would reopen it to run the scheduled
'   refresh, and a hidden Excel shows no Ribbon.
'
' ERROR POLICY
'   Never raises: the refresh is a convenience.
'
' UPDATED
'   2026-11-07
'==============================================================================
'
        On Error GoTo Skipped
        If gSilent Or Not Application.Visible Then
            Exit Sub
        End If
        M_Ribbon.ScheduleRefresh

Skipped:

End Sub


'
'------------------------------------------------------------------------------
'
'                            NAVIGATION BAR (#620)
'
'------------------------------------------------------------------------------
'

Public Function NavigationBarShown() As Boolean
'
'==============================================================================
'                              NavigationBarShown
'------------------------------------------------------------------------------
' PURPOSE
'   Whether the navigation bar of #615 is shown, read from the workbook:
'   column A of the first worksheet is not hidden. The Ribbon's toggle
'   reads it, so its state survives a VBA state reset.
'
' ERROR POLICY
'   Never raises; True when the sheet cannot be read.
'
' UPDATED
'   2026-11-07
'==============================================================================
'
        On Error GoTo Failed
        NavigationBarShown = Not ThisWorkbook.Worksheets(1).Columns(1).Hidden
        Exit Function

Failed:
        NavigationBarShown = True

End Function


Public Sub ShowNavigationBar( _
    ByVal Shown As Boolean)
'
'==============================================================================
'                              ShowNavigationBar
'------------------------------------------------------------------------------
' PURPOSE
'   Show or hide the navigation bar on every sheet at once (#620): column
'   A and the bar's two shape groups. The sheets stay protected; the
'   change is for this session only, since a save shows the bar again
'   (NavigationBarBeforeSave). The tab is then refreshed
'   (RibbonStateChanged), so the Ribbon's toggle shows the state the
'   workbook holds also when the change came from the Macros dialog.
'
' USAGE
'   The Navigation bar toggle of the K-SACCR Ribbon tab
'   (M_Ribbon.RibbonToggle). ToggleNavigationBar flips the state from the
'   Macros dialog; automation calls SetNavigationBar_Silent.
'
' ERROR POLICY
'   Reports an error in a message, or in the Immediate window in silent
'   mode; never raises.
'
' UPDATED
'   2026-11-07
'==============================================================================
'
        On Error GoTo Failed
        If mRunning Then
            Notify "An SA-CCR operation is already running.", vbExclamation
        Else
            ApplyNavigationBar Shown
        End If
        RibbonStateChanged
        Exit Sub

Failed:
        RibbonStateChanged
        Notify "Navigation bar: " & Err.Description, vbExclamation

End Sub


Public Sub ToggleNavigationBar()
'
'==============================================================================
'                             ToggleNavigationBar
'------------------------------------------------------------------------------
' PURPOSE
'   Hide the navigation bar when it is shown and show it when it is
'   hidden (ShowNavigationBar), from the Macros dialog (#620).
'
' UPDATED
'   2026-11-07
'==============================================================================
'
        ShowNavigationBar Not NavigationBarShown()

End Sub


Public Function SetNavigationBar_Silent( _
    ByVal Shown As Boolean) _
    As String
'
'==============================================================================
'                           SetNavigationBar_Silent
'------------------------------------------------------------------------------
' PURPOSE
'   Show or hide the navigation bar on every sheet (#620), for automation
'   and tests, without a message. The tab is then refreshed
'   (RibbonStateChanged) when Excel is visible, so that the Ribbon's
'   toggle follows.
'
' RETURNS
'   "RESULT=SHOWN; sheets=N" or "RESULT=HIDDEN; sheets=N", N being the
'   worksheets changed; "RESULT=REFUSED; reason=..." while an operation
'   runs; "RESULT=FAILED; error=..." when a sheet could not be changed.
'
' ERROR POLICY
'   Never raises.
'
' UPDATED
'   2026-11-07
'==============================================================================
'
        If mRunning Then
            SetNavigationBar_Silent = "RESULT=REFUSED; reason=an SA-CCR operation is running"
            Exit Function
        End If
        On Error GoTo Failed
        SetNavigationBar_Silent = IIf(Shown, "RESULT=SHOWN", "RESULT=HIDDEN") & "; sheets=" & ApplyNavigationBar(Shown)
        RibbonStateChanged
        Exit Function

Failed:
        SetNavigationBar_Silent = "RESULT=FAILED; error=" & Err.Description
        RibbonStateChanged

End Function


Public Sub NavigationBarBeforeSave()
'
'==============================================================================
'                           NavigationBarBeforeSave
'------------------------------------------------------------------------------
' PURPOSE
'   Before a save, show the navigation bar when it is hidden, and note it,
'   so that the saved file always shows the bar, also to a user who opens
'   it with macros disabled (#620). ThisWorkbook.Workbook_BeforeSave calls
'   it. The Saved flag is kept.
'   Excel raises Workbook_AfterSave only after a save that happened: a
'   cancelled Save As dialog raises none. So the restore is also scheduled
'   (Application.OnTime) for when Excel is idle again, after the save or
'   the dialog; NavigationBarAfterSave cancels it when it runs first.
'
' ERROR POLICY
'   Never raises: a bar left shown is hidden again by the toggle.
'
' UPDATED
'   2026-11-07
'==============================================================================
'
        On Error GoTo Failed
        mBarHiddenForSave = Not NavigationBarShown()
        If mBarHiddenForSave Then
            ApplyNavigationBar True
            mBarRestoreAt = Now
            mBarRestoreName = BarRestoreProcedure()
            Application.OnTime mBarRestoreAt, mBarRestoreName
        End If

Failed:

End Sub


Public Sub NavigationBarAfterSave()
'
'==============================================================================
'                           NavigationBarAfterSave
'------------------------------------------------------------------------------
' PURPOSE
'   After a save, hide the navigation bar again when
'   NavigationBarBeforeSave showed it (#620), and cancel the restore it
'   scheduled. ThisWorkbook.Workbook_AfterSave calls it, and so does that
'   scheduled restore when no save followed, as after a cancelled Save As.
'   The Saved flag is kept, so a saved workbook stays saved.
'
' ERROR POLICY
'   Never raises.
'
' UPDATED
'   2026-11-07
'==============================================================================
'
        CancelBarRestore
        On Error GoTo Failed
        If mBarHiddenForSave Then
            ApplyNavigationBar False
        End If

Failed:
        mBarHiddenForSave = False

End Sub


Private Function BarRestoreProcedure() As String
'
'==============================================================================
'                             BarRestoreProcedure
'------------------------------------------------------------------------------
' PURPOSE
'   The workbook-qualified name of NavigationBarAfterSave for
'   Application.OnTime, with an apostrophe in the workbook name doubled, as
'   Excel requires (#620).
'
' UPDATED
'   2026-11-07
'==============================================================================
'
        BarRestoreProcedure = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!M_Main.NavigationBarAfterSave"

End Function


Private Sub CancelBarRestore()
'
'==============================================================================
'                              CancelBarRestore
'------------------------------------------------------------------------------
' PURPOSE
'   Cancel the restore NavigationBarBeforeSave scheduled, when one is
'   pending, so that it neither runs twice nor runs after a Save As has
'   renamed the workbook (#620). It is cancelled with the name it was
'   scheduled with, which a Save As does not change.
'
' ERROR POLICY
'   Never raises: a restore that already ran cannot be cancelled, and that
'   is the expected case when the restore itself calls this.
'
' UPDATED
'   2026-11-07
'==============================================================================
'
        On Error GoTo Done
        If mBarRestoreAt <> 0 Then
            Application.OnTime mBarRestoreAt, mBarRestoreName, , False
        End If

Done:
        mBarRestoreAt = 0
        mBarRestoreName = vbNullString

End Sub


Private Function ApplyNavigationBar( _
    ByVal Shown As Boolean) _
    As Long
'
'==============================================================================
'                             ApplyNavigationBar
'------------------------------------------------------------------------------
' PURPOSE
'   Show or hide column A and the shape groups NAV_GROUP_TOP and
'   NAV_GROUP_REST of every worksheet in one pass, with screen updating
'   off, so the layout changes at once (#620). ProtectWorkbook runs first,
'   so the macros change the protected sheets through their user-interface
'   protection. The workbook's Saved flag and screen updating are kept.
'
' RETURNS
'   The number of worksheets changed.
'
' ERROR POLICY
'   Re-raises the first error after screen updating and the Saved flag are
'   restored.
'
' UPDATED
'   2026-11-07
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Sheet               As Worksheet    'Each worksheet
    Dim Item                As Shape        'Each shape of a sheet
    Dim WasSaved            As Boolean      'The workbook's Saved flag before
    Dim SavedScreen         As Boolean      'Screen updating before
    Dim Changed             As Long         'Worksheets changed
    Dim ErrorNumber         As Long         'The first error's number
    Dim ErrorDescription    As String       'The first error's description

'------------------------------------------------------------------------------
' SHOW OR HIDE
'------------------------------------------------------------------------------
        WasSaved = ThisWorkbook.Saved
        SavedScreen = Application.ScreenUpdating
        On Error GoTo Failed
        Application.ScreenUpdating = False
        ProtectWorkbook
        For Each Sheet In ThisWorkbook.Worksheets
            Sheet.Columns(1).Hidden = Not Shown
            For Each Item In Sheet.Shapes
                If Item.Name = NAV_GROUP_TOP Or Item.Name = NAV_GROUP_REST Then
                    Item.Visible = IIf(Shown, msoTrue, msoFalse)
                End If
            Next Item
            Changed = Changed + 1
        Next Sheet
        Application.ScreenUpdating = SavedScreen
        ThisWorkbook.Saved = WasSaved
        ApplyNavigationBar = Changed
        Exit Function

Failed:
        ErrorNumber = Err.Number
        ErrorDescription = Err.Description
        Application.ScreenUpdating = SavedScreen
        ThisWorkbook.Saved = WasSaved
        Err.Raise ErrorNumber, "M_Main.ApplyNavigationBar", ErrorDescription

End Function


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
'   The result line built by ReportOutcome; empty when the operation raised
'   an error or was refused because another was running.
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
'   2026-11-05
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
        ProtectWorkbook
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
            If Not InputsRecalculated() Or gTestFault = "recalc" Then
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
'   2026-10-28
'==============================================================================
'

'------------------------------------------------------------------------------
' CLEAR
'------------------------------------------------------------------------------
        ClearOutputBlock GetSheet(SH_TRADECALC), OutputFirstRow(SH_TRADECALC), TC_NCOLS
        ClearOutputBlock GetSheet(SH_BUCKETS), OutputFirstRow(SH_BUCKETS), BK_NCOLS
        ClearOutputBlock GetSheet(SH_HEDGING), OutputFirstRow(SH_HEDGING), HS_NCOLS
        ClearOutputBlock GetSheet(SH_RESULTS), OutputFirstRow(SH_RESULTS), RS_NCOLS
        ClearChecksBlock GetSheet(SH_CHECKS), OutputFirstRow(SH_CHECKS)
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
'   reconciliation=DONE; trade_reconciliation=NONE; codes=TR-MATURED;
'   cleanup=PASS"; codes, the distinct diagnostic codes of the run, comes
'   before cleanup for a run and a validation, not for a clear (#81).
'
' ERROR POLICY
'   In silent mode a cleanup failure raises ERR_CLEANUP_FAILED, with the
'   result line in its description; otherwise it is added to the message.
'
' UPDATED
'   2026-11-09
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
    'A completed run shows Results, and its message points to Checks whenever
    'there are errors or warnings, since it can still have excluded trades. A
    'stopped run and a validation show Checks. ShowSheet never raises.
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
                    ShowSheet SH_RESULTS
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
                    ShowSheet SH_CHECKS
                End If
            Case OP_VALIDATE
                ShowSheet SH_CHECKS
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
        If Operation <> OP_CLEAR Then ResultLine = ResultLine & "; codes=" & CORE_Engine.DiagnosticCodes
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


Private Sub ShowSheet(ByVal SheetName As String)
'
'==============================================================================
'                                  ShowSheet
'------------------------------------------------------------------------------
' PURPOSE
'   Bring the sheet an operation leaves the user on to the front: Results
'   after a completed run, Checks after a stopped run or a validation.
'
' INPUTS
'   SheetName: the sheet's name, SH_RESULTS or SH_CHECKS.
'
' STATE OWNERSHIP
'   Activates the sheet; does nothing in silent mode, where no one is
'   looking and the window may be hidden.
'
' ERROR POLICY
'   Best effort: activation fails when the window or the sheet is hidden.
'   That failure is contained, because showing the sheet is a convenience
'   and must not replace the operation's outcome.
'
' UPDATED
'   2026-11-06
'==============================================================================
'

'------------------------------------------------------------------------------
' SHOW
'------------------------------------------------------------------------------
        If gSilent Then
            Exit Sub
        End If
        On Error GoTo NotShown
        ThisWorkbook.Worksheets(SheetName).Activate
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
'   2026-10-21
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
    'ERR_RUN_ACTIVE and ERR_CHECKS_WRITE are shown with their description
    'alone; any other error, the project's own included, with its number.
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
