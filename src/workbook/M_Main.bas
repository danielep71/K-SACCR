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
'   thrown by it would (#204); gTestFault = "reconcile-rows" makes the reconciliation need every row of its
'   sheet and one more (#718); gTestFault = "reconcile" raises
'   ERR_INJECTED_FAULT in a reconciliation that has bank figures (#318);
'   gTestFault = "evidence-open", "evidence-write" and "evidence-replace"
'   raise ERR_INJECTED_FAULT in WriteTextFile, which writes the evidence
'   report, the CSV templates and the import specification, before the new file is opened, while it
'   is written and while the older file is replaced (#312).
'   gTestFault = "staging-write" raises ERR_INJECTED_FAULT in LoadStaging
'   once the new staging block has been written (#208).
'   gTestFault = "bank-write" raises ERR_INJECTED_FAULT in LoadBank once
'   the bank records have been written (#718).
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
'   2026-11-22
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

    'The formatted run report (#259): the kinds of the columns of each table
    'and the measures of a reconciliation that are not amounts, as
    'docs/workbook/report.json lists them (tools/test_report_spec.py).
        Private Const RPT_TOTALS_HEADERS      As String = "Regime|V (sum MtM)|C (VM + NICA)|RC|AddOn IR|AddOn FX|" & _
                                                          "AddOn CR|AddOn EQ|AddOn CO|AddOn OT|AddOn aggregate|PFE|EAD|Status"
        Private Const RPT_TOTALS_KINDS        As String = "text|amount|amount|amount|amount|amount|amount|amount|" & _
                                                          "amount|amount|amount|amount|amount|text"
        Private Const RPT_RESULTS_KINDS       As String = "text|text|text|text|count|count|amount|amount|amount|" & _
                                                          "amount|amount|amount|amount|amount|amount|amount|ratio|" & _
                                                          "amount|ratio|amount|amount|amount|text|amount|amount|" & _
                                                          "amount|ratio|text"
        Private Const RPT_CHECKS_KINDS        As String = "text|text|count|text|text|text|text|text"
        Private Const RPT_RECON_KINDS         As String = "text|count|text|text|text|text|measure|measure|measure|" & _
                                                          "ratio|text|text"
        Private Const RPT_TRADE_RECON_KINDS   As String = "text|count|text|text|text|text|measure|measure|measure|" & _
                                                          "ratio|text|text|text"
        Private Const RPT_RECON_RATIOS        As String = "Multiplier"
        Private Const RPT_TRADE_RECON_RATIOS  As String = "M|Supervisory duration|Supervisory delta|MF unmargined|" & _
                                                          "MF margined"
    'The look of the report: the font, the header fill and the tab colors of
    'the workbook's sheets (docs/workbook/sheet-style.json), as BGR values.
        Private Const RPT_FONT                As String = "Aptos"       'Font of every cell
        Private Const RPT_HEADER_FILL         As Long = &H64381F        'Header fill 1F3864, with white bold text
        Private Const RPT_TAB_REFERENCE       As Long = &H64381F        'Tab 1F3864: Run
        Private Const RPT_TAB_OUTPUT          As Long = &HB6752E        'Tab 2E75B6: every other section
    'The pages of a PDF report: A4 landscape, 841.89 by 595.28 points, with
    'the margins every sheet is given and the printable area they leave,
    'in which the Charts sheet is laid out; the pictures of the charts.
        Private Const RPT_MARGIN_SIDE         As Double = 50.4          'Left and right margins, 0.7 inch
        Private Const RPT_MARGIN_END          As Double = 54#           'Top and bottom margins, 0.75 inch
        Private Const RPT_PRINT_WIDTH         As Double = 741#          'Page width less the side margins
        Private Const RPT_PRINT_HEIGHT        As Double = 487#          'Page height less the end margins
        Private Const RPT_CHART_GAP           As Double = 12#           'Points around and between pictures
        Private Const RPT_CHART_ZOOM          As Long = 300             'Window zoom the charts are pictured at
    'What the PDF keeps of long tables, as report.json gives it (owner's
    'decision of 2026-11-26, after a real portfolio gave 163 pages).
        Private Const RPT_PDF_TRADE_DIFFERENCES As Long = 50            'Largest differences of the trade legs
        Private Const RPT_PDF_CHECKS_PER_CODE   As Long = 10            'Rows of Checks of each code

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
        Public gTestFault   As String     'Test seam: "", "operation", "cleanup", "recalc", "recalc-error", "reconcile", "reconcile-rows", "evidence-*", "staging-write", "commit-*" or "restore-*"
        Public gQuietLog    As Boolean    'Test seam: silent-mode messages not printed
        Public gLastMessage As String     'Test seam: the last message passed to Notify
        Public gKeepProtection As String  'Test seam: sheets, between bars, ProtectWorkbook leaves as they are
        Public gTestEcbAnswer  As String  'Test seam: the answer of the ECB with gTestFault = "ecb-answer" (#745)
        Private mRunning    As Boolean    'True while an operation is in progress
    'The lookup of closing prices in progress (#721): the formulas are in the
    'scratch column and the prices are awaited. Excel hands a request of
    'STOCKHISTORY back only when no macro runs, so the button writes the
    'formulas, returns, and PollClosingPrices, scheduled each second with
    'Application.OnTime, reads the prices when they have come.
        Private mPriceActive    As Boolean    'Formulas written, prices awaited
        Private mPriceSheet     As Worksheet  'Params
        Private mPriceKeyColumn As Long       'Key column of the FX table
        Private mPriceFirstRow  As Long       'First row of the table
        Private mPriceLastRow   As Long       'Last row of the table
        Private mPriceCount     As Long       'Rows with a pair
        Private mPriceRows()    As Long       'The rows with a pair
        Private mPricePairs()   As String     'The pair each price is asked for
        Private mPriceKept()    As Variant    'The rate each row held before
        Private mPriceKeptDate  As Variant    'The date of the prices before
        Private mPriceKeptSource As Variant   'The source of the rates before (#745)
        Private mPriceKeptPrint As Variant    'The fingerprint of the table as downloaded before (#745)
        Private mPriceReportRow As Long       'The row of the reporting currency; 0 for none (#745)
        Private mPriceReportKept As Variant   'Its rate before, set to 1 with the prices
        Private mPriceAsOf      As Variant    'The reporting date asked for
        Private mPriceMode      As XlCalculation 'The caller's calculation mode
        Private mPriceDeadline  As Date       'When to stop waiting
        Private mPriceStarted   As Date       'When the formulas were written
        Private mPriceWritten   As Boolean    'FinishPriceLookup began to write rates
        Private mPollAt         As Date       'When the next poll is scheduled; 0 for none
        Private mPollName       As String     'The procedure it was scheduled with
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
    'The progress of a long operation in the status bar (#807): the
    'operation ProgressStart named, "" when none shows its progress, the
    'text last written, so that a step that changes nothing is not written,
    'and the step, for ProgressDetail.
        Private mProgressOperation As String  'Operation in progress; "" for none
        Private mProgressText      As String  'Text of the status bar; "" for none
        Private mProgressDone      As Long    'Steps done
        Private mProgressTotal     As Long    'Steps in all


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
'   protect every sheet again for the user interface only (#421), and give
'   every sheet a zoom of 100% (#726).
'
' UPDATED
'   2026-11-20
'==============================================================================
'

'------------------------------------------------------------------------------
' PROTECT
'------------------------------------------------------------------------------
        ProtectWorkbook
        ApplyOpeningView

End Sub


Public Sub ApplyOpeningView()
'
'==============================================================================
'                              ApplyOpeningView
'------------------------------------------------------------------------------
' PURPOSE
'   Give every visible sheet of this workbook a zoom of 100%, so that a copy
'   saved with another zoom opens as the template does (#726). The sheet that
'   was active stays active. Events are off meanwhile, so that no sheet
'   refreshes its status because it was visited, and the Saved flag is kept.
'
' ERROR POLICY
'   Never raises and never stops the opening: nothing is done when this
'   workbook is not the active one, or has no window.
'
' UPDATED
'   2026-11-20
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Sheet          As Worksheet    'Each worksheet
    Dim Original       As Object       'The sheet that was active
    Dim WasSaved       As Boolean      'The workbook's Saved flag before
    Dim SavedEvents    As Boolean      'Application.EnableEvents before
    Dim SavedScreen    As Boolean      'Application.ScreenUpdating before

'------------------------------------------------------------------------------
' ZOOM
'------------------------------------------------------------------------------
        On Error Resume Next
        If Not Application.ActiveWorkbook Is ThisWorkbook Then Exit Sub
        If ThisWorkbook.Windows.Count = 0 Then Exit Sub
        WasSaved = ThisWorkbook.Saved
        SavedEvents = Application.EnableEvents
        SavedScreen = Application.ScreenUpdating
        Set Original = ThisWorkbook.ActiveSheet
        Application.EnableEvents = False
        Application.ScreenUpdating = False
        For Each Sheet In ThisWorkbook.Worksheets
            If Sheet.Visible = xlSheetVisible Then
                Sheet.Activate
                ThisWorkbook.Windows(1).Zoom = 100
            End If
        Next Sheet
        If Not Original Is Nothing Then Original.Activate
        Application.ScreenUpdating = SavedScreen
        Application.EnableEvents = SavedEvents
        ThisWorkbook.Saved = WasSaved
        Err.Clear

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
'   PROTECT_OUTLINING lists it. A sheet of PROTECT_NONE is never protected:
'   one found protected without a password is unprotected (#720).
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
        If InStr(1, PROTECT_NONE, Listed, vbBinaryCompare) > 0 Then
            If Sheet.ProtectContents Then
                Sheet.Unprotect
            End If
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
'   trade_reconciliation, year_convention, calendar, holiday_table,
'   fx_source, codes and cleanup=PASS or FAIL, separated by "; ".
'   codes lists distinct ERROR/WARNING codes in encounter order, comma-separated;
'   it is empty when the current run has none.
'   incomplete counts netting sets whose EAD was withheld; total_ead_crr
'   and total_ead_bcbs are the totals of each regime's VALID netting sets,
'   0 when it has none (#198); inputs is the fingerprint of the inputs
'   the results were calculated from, empty when there are none.
'   year_convention, calendar and holiday_table say how the dates became
'   years (#721), fx_source where the FX rates came from (#745), as
'   CORE_Engine.FxSourceText writes it.
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
'   2026-11-22
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
'   Clear outputs on the K-SACCR Ribbon tab (Calculation group), the
'   button of an output sheet, or the Macros dialog.
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


Private Sub ReapplyOutputFilters()
'
'==============================================================================
'                            ReapplyOutputFilters
'------------------------------------------------------------------------------
' PURPOSE
'   After an operation has written the output tables, or the validation of a
'   staging sheet has written Checks, apply again the filter the user set on
'   the header of Results, HedgingSets, Buckets, TradeCalc and Checks (#754):
'   the rows written into hidden rows would otherwise stay as they were,
'   shown or hidden, whatever the rule says. A sheet with no
'   filter set, or without the buttons, is left as it is; no operation sets
'   or clears a filter.
'
' ERROR POLICY
'   Never raises: the rows stay as written, and the user can apply the filter
'   again from the buttons.
'
' UPDATED
'   2026-11-22
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim SheetNames  As Variant    'The five sheets
    Dim i           As Long       'Sheet
    Dim Sheet       As Worksheet  'One of them

'------------------------------------------------------------------------------
' APPLY AGAIN
'------------------------------------------------------------------------------
        On Error Resume Next
        SheetNames = Split(OUTPUT_FILTER_SHEETS, "|")
        For i = LBound(SheetNames) To UBound(SheetNames)
            Set Sheet = Nothing
            Set Sheet = GetSheet(CStr(SheetNames(i)))
            If Not Sheet Is Nothing Then
                If Sheet.AutoFilterMode Then
                    If Sheet.FilterMode Then
                        Sheet.AutoFilter.ApplyFilter
                    End If
                End If
            End If
        Next i

End Sub


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
            If gTestFault = "reconcile-rows" Then
                RowsToWrite = ReconSheet.Rows.Count
            End If
        'One row per figure, 11 per trade leg: a sheet holds 1,048,576 rows. A reconciliation that does not
        'fit is not made, and the message says what it needs (#718).
            If ReconFirstRow + RowsToWrite - 1 > ReconSheet.Rows.Count Then
                Problem = "it needs " & RowsToWrite & " rows from row " & ReconFirstRow & " and " & ReconSheet.Name & _
                          " holds " & ReconSheet.Rows.Count & ": use fewer bank rows"
                ClearReconciliation "Reconciliation not run " & Format$(Now, "yyyy-mm-dd hh:mm:ss") & ": " & Problem & ".", _
                                    ReconKind
                AddReconLine ReconName & " not run: " & Problem & "."
                ReconcileAfterRun = "FAILED"
                Exit Function
            End If
            WriteBlock ReconSheet, ReconFirstRow, ReconRows, RowsToWrite, ReconColsOf(ReconKind)
            FormatColumns ReconSheet, ReconFirstRow, RowsToWrite, "@|0|@|@|@|@|#,##0.00|#,##0.00|" & _
                          "#,##0.00|0.0000%|@|@" & IIf(ReconKind = RECON_NS, "", "|@")
        'The rows of a filter the user set stay filtered: the rule is applied to the new rows. The
        'template's filter ends at row 100,000: a reconciliation that goes beyond it gets a new filter
        'over all its rows, without the criteria of the old one (#720).
            If ReconSheet.AutoFilterMode Then
                If ReconFirstRow + RowsToWrite - 1 > RECON_FILTER_LAST_ROW Then
                    ReconSheet.AutoFilterMode = False
                    TableRange(ReconSheet, ReconFirstRow - 1, 1, ReconFirstRow + RowsToWrite - 1, ReconColsOf(ReconKind)).AutoFilter
                ElseIf ReconSheet.FilterMode Then
                    On Error Resume Next
                    ReconSheet.AutoFilter.ApplyFilter
                    On Error GoTo Failed
                End If
            End If
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


Private Sub WriteDashboard()
'
'==============================================================================
'                                WriteDashboard
'------------------------------------------------------------------------------
' PURPOSE
'   After a completed run and its reconciliations, write every block of
'   the Dashboard (#675): read Results, TradeCalc, Checks, NettingSets and
'   the two reconciliation sheets, let CORE_Engine compute the blocks,
'   and write each into its workbook name.
'
' ERROR POLICY
'   A workbook without the Dashboard sheet, built from an older template,
'   is left alone. A block name that is missing or of another size than
'   its values is ERR_OUTPUT_LAYOUT. Every block is computed before one is
'   written; any error clears every block, best effort, and is raised
'   again, so the sheet never shows part of a run.
'
' UPDATED
'   2026-11-23
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim BlockList   As Variant    'Block names
    Dim Values      As Variant    'The blocks' values, in BlockList order
    Dim RunInfo     As Variant    'Date, time, currency, regime, class, inputs
    Dim Index       As Long       'Block position
    Dim ErrNumber   As Long       'Error captured before the cleanup
    Dim ErrSource   As String     'Its source
    Dim ErrText     As String     'Its description

'------------------------------------------------------------------------------
' COMPUTE EVERY BLOCK
'------------------------------------------------------------------------------
        If Not SheetExists(SH_DASHBOARD) Then Exit Sub
        On Error GoTo Failed
        RunInfo = Array(GetParam(PRM_ASOF), Now, GetParam(PRM_REPCCY), GetParam(PRM_REGIME), _
                        CORE_Engine.RunClass, CORE_Engine.RunInputs)
        Values = CORE_Engine.DashboardBlocks( _
                     DashboardRows(SH_RESULTS, RS_NCOLS), DashboardRows(SH_TRADECALC, TC_NCOLS), _
                     DashboardRows(SH_CHECKS, CK_NCOLS), DashboardRows(SH_NS, NS_NCOLS), _
                     DashboardRows(SH_RECON, RC_RESULT), DashboardRows(SH_TRADE_RECON, RC_RESULT), RunInfo)

'------------------------------------------------------------------------------
' WRITE
'------------------------------------------------------------------------------
        BlockList = CORE_Engine.DashboardBlockNames()
        For Index = LBound(BlockList) To UBound(BlockList)
            DashboardRange(CStr(BlockList(Index)), Values(Index)).Value = Values(Index)
        Next Index
        Exit Sub

'------------------------------------------------------------------------------
' CLEAR AND RAISE AGAIN
'------------------------------------------------------------------------------
Failed:
        ErrNumber = Err.Number
        ErrSource = Err.Source
        ErrText = Err.Description
        ClearDashboard
        Err.Raise ErrNumber, ErrSource, "Dashboard not written: " & ErrText

End Sub


Private Sub ClearDashboard()
'
'==============================================================================
'                                ClearDashboard
'------------------------------------------------------------------------------
' PURPOSE
'   Empty the value cells of every Dashboard block (#675); the labels,
'   headers, formats and charts stay. Every operation starts with this, so
'   that no Dashboard outlives the results it summarised.
'
' ERROR POLICY
'   Best effort, and never raises: a workbook without the sheet or a name
'   is left as it is.
'
' UPDATED
'   2026-11-23
'==============================================================================
'
    Dim BlockList   As Variant    'Block names
    Dim Index       As Long       'Block position
        On Error Resume Next
        BlockList = CORE_Engine.DashboardBlockNames()
        For Index = LBound(BlockList) To UBound(BlockList)
            ThisWorkbook.Names(CStr(BlockList(Index))).RefersToRange.ClearContents
        Next Index
        Err.Clear

End Sub


Private Function DashboardRows( _
    ByVal SheetName As String, _
    ByVal ColumnCount As Long) _
    As Variant
'
'==============================================================================
'                                DashboardRows
'------------------------------------------------------------------------------
' PURPOSE
'   The rows of a table for the Dashboard, from its first row down to the
'   last with a value in its first column, as a 1-based two-dimensional
'   array; Empty when there is none. A reconciliation sheet missing from
'   an older workbook reads as Empty.
'
' UPDATED
'   2026-11-23
'==============================================================================
'
    Dim Sheet      As Worksheet    'The table's sheet
    Dim FirstRow   As Long         'First data row
    Dim LastRow    As Long         'Last data row
    Dim Values     As Variant      'The cells
    Dim OneCell(1 To 1, 1 To 1) As Variant    'A one-cell table as an array
        DashboardRows = Empty
        If (SheetName = SH_RECON Or SheetName = SH_TRADE_RECON) And Not SheetExists(SheetName) Then Exit Function
        Set Sheet = GetSheet(SheetName)
        FirstRow = OutputFirstRow(SheetName)
        LastRow = LastDataRow(Sheet, FirstRow, 1)
        If LastRow < FirstRow Then Exit Function
        Values = TableRange(Sheet, FirstRow, 1, LastRow, ColumnCount).Value2
        If Not IsArray(Values) Then
            OneCell(1, 1) = Values
            Values = OneCell
        End If
        DashboardRows = Values

End Function


Private Function SheetExists( _
    ByVal SheetName As String) _
    As Boolean
'
'==============================================================================
'                                 SheetExists
'------------------------------------------------------------------------------
' PURPOSE
'   Tell whether the workbook has a worksheet of this name.
'
' UPDATED
'   2026-11-23
'==============================================================================
'
    Dim Sheet   As Worksheet    'The sheet; Nothing when absent
        On Error Resume Next
        Set Sheet = ThisWorkbook.Worksheets(SheetName)
        On Error GoTo 0
        SheetExists = Not (Sheet Is Nothing)

End Function


Private Function DashboardRange( _
    ByVal BlockName As String, _
    ByVal Values As Variant) _
    As Range
'
'==============================================================================
'                                DashboardRange
'------------------------------------------------------------------------------
' PURPOSE
'   The value cells of a Dashboard block, through its workbook name,
'   checked to be on the Dashboard and of the size of Values.
'
' ERROR POLICY
'   ERR_OUTPUT_LAYOUT when the name is missing, elsewhere or of another
'   size.
'
' UPDATED
'   2026-11-23
'==============================================================================
'
    Dim Target   As Range    'The name's cells
        On Error Resume Next
        Set Target = ThisWorkbook.Names(BlockName).RefersToRange
        On Error GoTo 0
        If Target Is Nothing Then
            Err.Raise ERR_OUTPUT_LAYOUT, "M_Main.DashboardRange", _
                      "Workbook name '" & BlockName & "' is missing - rebuild the workbook from the template."
        End If
        If Target.Worksheet.Name <> SH_DASHBOARD Or Target.Rows.Count <> UBound(Values, 1) Or _
           Target.Columns.Count <> UBound(Values, 2) Then
            Err.Raise ERR_OUTPUT_LAYOUT, "M_Main.DashboardRange", _
                      "Workbook name '" & BlockName & "' does not cover " & UBound(Values, 1) & " x " & _
                      UBound(Values, 2) & " cells of the " & SH_DASHBOARD & " sheet - rebuild the workbook from the template."
        End If
        Set DashboardRange = Target

End Function


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
'                              RUN REPORT EXPORT
'
'------------------------------------------------------------------------------
'

Public Sub ExportReport()

'
'------------------------------------------------------------------------------
'                    WRITE A FORMATTED REPORT OF THE LAST RUN
'------------------------------------------------------------------------------
' PURPOSE
'   Writes a formatted report of the last run (#259), as a PDF or an XLSX
'   chosen by the user in the Save As dialog
'
' WHY THIS EXISTS
'   A reviewer reads results rather than parses them: the report opens
'   with the Dashboard and gives the evidence behind it. The CSV of Export
'   evidence stays the machine-readable record (docs/workbook/REPORT.md)
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   - Refuses, with the reason, in the cases of Export evidence and when
'     the Dashboard was changed after the run (ReportProblem)
'   - Asks for the file; the extension, .pdf or .xlsx, decides the format
'   - Writes the report (WriteReport), showing its progress in the status
'     bar (ProgressStart, #807), and says where
'
' ERROR POLICY
'   Reports a refusal or a failure in a message, or in the Immediate
'   window in silent mode; never raises
'
' DEPENDENCIES
'   ReportProblem, ReportFileFormat, WriteReport, DefaultReportName,
'   ProgressStart, Notify
'
' NOTES
'   Runs from Export report on the K-SACCR Ribbon tab (M_Ribbon), or the
'   Macros dialog
'
' UPDATED
'   2026-11-27
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "M_Main.ExportReport"

    Dim Problem         As String     'Why the report cannot be written; "" when it can
    Dim ChosenFile      As Variant    'Chosen file, or False when cancelled
    Dim FileFormat      As String     'PDF or XLSX; "" for another extension
    Dim Written         As String     'What WriteReport wrote

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' CHECK THE RUN
'------------------------------------------------------------------------------
    'The report describes the run the evidence CSV would describe, or none
        Problem = ReportProblem()
        If VBA.Len(Problem) > 0 Then
            Notify "Report not exported: " & Problem & ".", vbExclamation
            Exit Sub
        End If

'------------------------------------------------------------------------------
' ASK FOR THE FILE
'------------------------------------------------------------------------------
    'The extension of the chosen file decides the format
        ChosenFile = Application.GetSaveAsFilename(InitialFileName:=DefaultReportName(), _
                                                   FileFilter:="PDF files (*.pdf), *.pdf, " & _
                                                               "Excel workbooks (*.xlsx), *.xlsx", _
                                                   Title:="Export report")
        If VarType(ChosenFile) = vbBoolean Then
            Exit Sub
        End If
        FileFormat = ReportFileFormat(CStr(ChosenFile))
        If VBA.Len(FileFormat) = 0 Then
            Notify "Report not exported: the file name must end in .pdf or .xlsx.", vbExclamation
            Exit Sub
        End If

'------------------------------------------------------------------------------
' WRITE
'------------------------------------------------------------------------------
    'Write the report, its progress in the status bar, and say where;
    'the message gives the status bar back
        ProgressStart "Report"
        Written = WriteReport(CStr(ChosenFile), FileFormat = "PDF")
        Notify "Report exported as " & FileFormat & " (" & Written & "):" & vbLf & vbLf & CStr(ChosenFile), _
               vbInformation

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
ErrorHandler:
    'Say what failed; WriteReport has put everything back, and the
    'message gives the status bar back
        Notify "Report not exported: " & Err.Description, vbCritical

End Sub


Public Function ExportReport_Silent( _
    ByVal FilePath As String) _
    As String
'
'------------------------------------------------------------------------------
'              WRITE A FORMATTED REPORT OF THE LAST RUN, UNATTENDED
'------------------------------------------------------------------------------
' PURPOSE
'   Writes the formatted report of the last run (#259) to FilePath
'   without asking, for automation and tests
'
' WHY THIS EXISTS
'   The macro entry point of the report, as ExportEvidence_Silent is that
'   of the evidence CSV
'
' INPUTS
'   FilePath
'     The file to write, ending in .pdf or .xlsx (any case); an existing
'     file is replaced, unless it is read-only
'
' RETURNS
'   "RESULT=OK; format=PDF|XLSX; sheets=N; charts=N; page_setup=...;
'   file=..." or "RESULT=REFUSED; reason=..."
'
' BEHAVIOR
'   Refuses another extension and the cases of ReportProblem, then
'   calls WriteReport
'
' ERROR POLICY
'   A report that cannot be written raises to the caller, after
'   WriteReport has put everything back
'
' DEPENDENCIES
'   ReportFileFormat, ReportProblem, WriteReport
'
' NOTES
'   None
'
' UPDATED
'   2026-11-26
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "M_Main.ExportReport_Silent"

    Dim FileFormat      As String    'PDF or XLSX; "" for another extension
    Dim Problem         As String    'Why the report cannot be written; "" when it can

'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'Only the two formats of the report
        FileFormat = ReportFileFormat(FilePath)
        If VBA.Len(FileFormat) = 0 Then
            ExportReport_Silent = "RESULT=REFUSED; reason=the file name must end in .pdf or .xlsx"
            Exit Function
        End If
    'The run the evidence CSV would describe, or none
        Problem = ReportProblem()
        If VBA.Len(Problem) > 0 Then
            ExportReport_Silent = "RESULT=REFUSED; reason=" & Problem
            Exit Function
        End If

'------------------------------------------------------------------------------
' WRITE
'------------------------------------------------------------------------------
    'WriteReport raises to the caller
        ExportReport_Silent = "RESULT=OK; format=" & FileFormat & "; " & _
                              WriteReport(FilePath, FileFormat = "PDF") & "; file=" & FilePath

End Function


Private Function ReportFileFormat( _
    ByVal FilePath As String) _
    As String
'
'------------------------------------------------------------------------------
'                     TELL THE FORMAT OF A REPORT FILE
'------------------------------------------------------------------------------
' PURPOSE
'   Returns "PDF" or "XLSX" from the extension of a file name, any case;
'   "" for another extension
'
' WHY THIS EXISTS
'   The extension the user types or chooses decides the format (#259)
'
' INPUTS
'   FilePath
'     A file name or path
'
' RETURNS
'   "PDF", "XLSX" or ""
'
' BEHAVIOR
'   Compares the last characters of the name
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   None
'
' NOTES
'   None
'
' UPDATED
'   2026-11-26
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'The extension, compared without case
        If VBA.LCase$(VBA.Right$(FilePath, 4)) = ".pdf" Then
            ReportFileFormat = "PDF"
        ElseIf VBA.LCase$(VBA.Right$(FilePath, 5)) = ".xlsx" Then
            ReportFileFormat = "XLSX"
        Else
            ReportFileFormat = vbNullString
        End If

End Function


Private Function DefaultReportName() As String

'
'------------------------------------------------------------------------------
'                       OFFER A NAME FOR A RUN REPORT
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the file name offered for a report: K-SACCR_report_ with the
'   date and time, as a PDF, in the workbook's folder when it has one
'
' WHY THIS EXISTS
'   The Save As dialog of Export report starts from it (#259)
'
' INPUTS
'   None
'
' RETURNS
'   The file name or path
'
' BEHAVIOR
'   As DefaultEvidenceName, with the name and extension of a report
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   InWorkbookFolder
'
' NOTES
'   None
'
' UPDATED
'   2026-11-26
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'The date and time keep reports of different runs apart
        DefaultReportName = InWorkbookFolder("K-SACCR_report_" & VBA.Format$(VBA.Now, "yyyymmdd-hhnnss") & ".pdf")

End Function


Private Function WriteReport( _
    ByVal FilePath As String, _
    ByVal AsPdf As Boolean) _
    As String
'
'------------------------------------------------------------------------------
'                BUILD THE RUN REPORT IN A NEW WORKBOOK AND SAVE IT
'------------------------------------------------------------------------------
' PURPOSE
'   Builds the formatted report of the last run in a new, temporary
'   workbook, one sheet per section of docs/workbook/report.json, and
'   saves it as an XLSX or exports it as a PDF to FilePath
'
' WHY THIS EXISTS
'   The report must leave the K-SACCR workbook as it is (#259): no sheet,
'   name, cell or setting of it changes, so its results stay current
'
' INPUTS
'   FilePath
'     The file to write; an existing file is replaced, unless it is
'     read-only
'   AsPdf
'     True for a PDF, with the charts and the reconciliation rows that
'     are not MATCH; False for an XLSX with every row
'
' RETURNS
'   "sheets=N; charts=N; page_setup=OK" or, when the page setup of a PDF
'   could not be applied, "page_setup=FAILED (...)"
'
' BEHAVIOR
'   - Refuses a read-only target before anything is built, and an export
'     while an operation runs
'   - Holds the operation guard (mRunning) until the end, on success and
'     on failure, so that no Run or Clear outputs starts while the charts
'     are pictured (Codex on #805)
'   - Turns off screen updating, alerts and events
'   - Gives the new workbook the font and the vertical alignment of the
'     K-SACCR sheets (docs/workbook/sheet-style.json)
'   - Writes Run, Dashboard, Charts (PDF only), Totals, Results, Checks,
'     and the two reconciliations when they were made, saying each in
'     the status bar when the caller shows the progress (ProgressShow,
'     #807)
'   - Hides the gridlines, saves or exports beside the target, closes
'     the new workbook and replaces the target (ReplaceReportFile)
'
' ERROR POLICY
'   Raises the first error after closing the new workbook unsaved,
'   deleting its file and putting back the settings and the active sheet;
'   an older target is left as it was
'
' DEPENDENCIES
'   AddReportSheet, WriteReportRun, WriteReportDashboard,
'   WriteReportCharts, WriteReportTotals, WriteReportResults,
'   WriteReportChecks, WriteReportReconciliation,
'   ApplyReportPageSetup, HideReportGridlines, ReportSiblingPath,
'   ReplaceReportFile, ProgressShow
'
' NOTES
'   gTestFault "report-save" raises before the file is saved, and
'   "report-replace" while the older file is replaced
'
' UPDATED
'   2026-11-27
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME         As String = "M_Main.WriteReport"

    Dim SavedScreen         As Boolean      'ScreenUpdating before the report
    Dim SavedAlerts         As Boolean      'DisplayAlerts before the report
    Dim SavedEvents         As Boolean      'EnableEvents before the report
    Dim SavedBook           As Workbook     'Active workbook before the report
    Dim SavedSheet          As Object       'Active sheet before the report
    Dim StateCaptured       As Boolean      'True once the settings were read
    Dim Guarded             As Boolean      'True once this export holds the operation guard
    Dim ReportBook          As Workbook     'The new workbook; Nothing once closed
    Dim ReportSheet         As Worksheet    'A sheet of the report
    Dim NewPath             As String       'The file before it replaces FilePath; "" once renamed
    Dim SheetCount          As Long         'Sheets written
    Dim ChartCount          As Long         'Charts pictured
    Dim ChartZoom           As Long         'Print zoom of the Charts sheet
    Dim SetupProblems       As String       'Page settings that could not be applied
    Dim Footer              As String       'Footer of every page of a PDF
    Dim TitleRow            As Long         'Row a sheet repeats on its pages; 0 for none
    Dim kind                As Long         'Reconciliation
    Dim Included(RECON_NS To RECON_TR) As Boolean 'The reconciliation was made and is reported
    Dim Steps               As Long         'Sections to write, and the save
    Dim Step                As Long         'Steps done
    Dim ErrorNumber         As Long         'First error; 0 on success
    Dim ErrorSource         As String       'Its source
    Dim ErrorText           As String       'Its description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        On Error GoTo ErrorHandler
    'A read-only target fails before anything is built
        If FileExists(FilePath) Then
            If (GetAttr(FilePath) And vbReadOnly) <> 0 Then
                Err.Raise ERR_EXPORT_WRITE, PROC_NAME, "the file is read-only"
            End If
        End If
    'One operation at a time: the export yields to Excel while it pictures
    'the charts, and a Run or Clear outputs then would mix two runs
        If mRunning Then
            Err.Raise ERR_RUN_ACTIVE, PROC_NAME, "An SA-CCR operation is already running."
        End If
        mRunning = True
        Guarded = True
    'Capture the settings the report changes
        SavedScreen = Application.ScreenUpdating
        SavedAlerts = Application.DisplayAlerts
        SavedEvents = Application.EnableEvents
        Set SavedBook = Application.ActiveWorkbook
        If Not SavedBook Is Nothing Then
            Set SavedSheet = SavedBook.ActiveSheet
        End If
        StateCaptured = True
        Application.ScreenUpdating = False
        Application.DisplayAlerts = False
        Application.EnableEvents = False

'------------------------------------------------------------------------------
' BUILD THE REPORT
'------------------------------------------------------------------------------
    'A new workbook with one sheet, named after the run's own
        Set ReportBook = Application.Workbooks.Add(xlWBATWorksheet)
        ReportBook.BuiltinDocumentProperties("Author").Value = MSG_TITLE
        ReportBook.BuiltinDocumentProperties("Title").Value = "K-SACCR run report"
    'The font and the vertical alignment of the workbook's sheets
        ReportBook.Styles("Normal").Font.Name = RPT_FONT
        ReportBook.Styles("Normal").VerticalAlignment = xlCenter
    'One sheet per section, in the order of report.json; the pages of a
    'PDF are set up as each sheet is written
        Footer = "K-SACCR report - run " & ReportRunTime() & " - inputs " & CORE_Engine.LastRunInputs() & _
                 " - page &P of &N"
    'The steps the status bar counts: each section and the save
        Steps = IIf(AsPdf, 7, 6)
        For kind = RECON_NS To RECON_TR
            If ReconciliationSheetsPresent(kind) Then
                Included(kind) = (VBA.Len(StoredText(InputsNameOf(kind))) > 0)
            End If
            If Included(kind) Then Steps = Steps + 1
        Next kind
        ProgressShow Step, Steps, "Run"
        Set ReportSheet = AddReportSheet(ReportBook, "Run")
        TitleRow = WriteReportRun(ReportSheet)
        If AsPdf Then SetupProblems = SetupProblems & ApplyReportPageSetup(ReportSheet, Footer, TitleRow)
        Step = Step + 1
        ProgressShow Step, Steps, "Dashboard"
        Set ReportSheet = AddReportSheet(ReportBook, "Dashboard")
        TitleRow = WriteReportDashboard(ReportSheet, AsPdf)
        If AsPdf Then SetupProblems = SetupProblems & ApplyReportPageSetup(ReportSheet, Footer, TitleRow)
        If AsPdf Then
            Step = Step + 1
            ProgressShow Step, Steps, "Charts"
            Set ReportSheet = AddReportSheet(ReportBook, "Charts")
            ChartCount = WriteReportCharts(ReportSheet, ChartZoom)
            SetupProblems = SetupProblems & ApplyReportPageSetup(ReportSheet, Footer, 0, ChartZoom)
        End If
        Step = Step + 1
        ProgressShow Step, Steps, "Totals"
        Set ReportSheet = AddReportSheet(ReportBook, "Totals")
        TitleRow = WriteReportTotals(ReportSheet, AsPdf)
        If AsPdf Then SetupProblems = SetupProblems & ApplyReportPageSetup(ReportSheet, Footer, TitleRow)
        Step = Step + 1
        ProgressShow Step, Steps, "Results"
        Set ReportSheet = AddReportSheet(ReportBook, "Results")
        TitleRow = WriteReportResults(ReportSheet, AsPdf)
        If AsPdf Then SetupProblems = SetupProblems & ApplyReportPageSetup(ReportSheet, Footer, TitleRow)
        Step = Step + 1
        ProgressShow Step, Steps, "Checks"
        Set ReportSheet = AddReportSheet(ReportBook, "Checks")
        TitleRow = WriteReportChecks(ReportSheet, AsPdf)
        If AsPdf Then SetupProblems = SetupProblems & ApplyReportPageSetup(ReportSheet, Footer, TitleRow)
        For kind = RECON_NS To RECON_TR
            If Included(kind) Then
                Step = Step + 1
                ProgressShow Step, Steps, IIf(kind = RECON_NS, "Reconciliation", "TradeReconciliation")
                Set ReportSheet = AddReportSheet(ReportBook, _
                                                 IIf(kind = RECON_NS, "Reconciliation", "TradeReconciliation"))
                TitleRow = WriteReportReconciliation(ReportSheet, kind, AsPdf)
                If AsPdf Then SetupProblems = SetupProblems & ApplyReportPageSetup(ReportSheet, Footer, TitleRow)
            End If
        Next kind
        SheetCount = ReportBook.Worksheets.Count

'------------------------------------------------------------------------------
' SAVE AND REPLACE
'------------------------------------------------------------------------------
    'Save beside the target, then replace it
        If gTestFault = "report-save" Then
            Err.Raise ERR_INJECTED_FAULT, PROC_NAME, "Injected failure before the report was saved."
        End If
        Step = Step + 1
        ProgressShow Step, Steps, IIf(AsPdf, "saving the PDF", "saving the workbook")
        HideReportGridlines ReportBook
        NewPath = ReportSiblingPath(FilePath)
        If AsPdf Then
            ReportBook.ExportAsFixedFormat Type:=xlTypePDF, Filename:=NewPath, Quality:=xlQualityStandard, _
                                           IncludeDocProperties:=False, IgnorePrintAreas:=False, _
                                           OpenAfterPublish:=False
        Else
            ReportBook.SaveAs Filename:=NewPath, FileFormat:=xlOpenXMLWorkbook
        End If
        ReportBook.Close SaveChanges:=False
        Set ReportBook = Nothing
        ReplaceReportFile NewPath, FilePath
        NewPath = vbNullString

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'What was written
        WriteReport = "sheets=" & SheetCount & "; charts=" & ChartCount & "; page_setup=" & _
                      IIf(VBA.Len(SetupProblems) = 0, "OK", "FAILED (" & SetupProblems & ")")

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Reached on success and, through the handler, after an error; a
    'cleanup error never replaces the first
CleanUp:
        On Error Resume Next
        If Not ReportBook Is Nothing Then
            ReportBook.Close SaveChanges:=False
        End If
        If VBA.Len(NewPath) > 0 Then
            If FileExists(NewPath) Then Kill NewPath
        End If
        If StateCaptured Then
            If Not SavedBook Is Nothing Then SavedBook.Activate
            If Not SavedSheet Is Nothing Then SavedSheet.Activate
            Application.EnableEvents = SavedEvents
            Application.DisplayAlerts = SavedAlerts
            Application.ScreenUpdating = SavedScreen
        End If
        If Guarded Then mRunning = False
        On Error GoTo 0
        If ErrorNumber <> 0 Then
            Err.Raise ErrorNumber, ErrorSource, "Could not write the report " & FilePath & ": " & ErrorText
        End If
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
ErrorHandler:
    'Keep the first error for the caller
        ErrorNumber = Err.Number
        ErrorSource = Err.Source
        ErrorText = Err.Description
        Resume CleanUp

End Function


Private Function AddReportSheet( _
    ByVal ReportBook As Workbook, _
    ByVal SheetName As String) _
    As Worksheet
'
'------------------------------------------------------------------------------
'                       ADD A SHEET TO THE RUN REPORT
'------------------------------------------------------------------------------
' PURPOSE
'   Returns a sheet of the report named SheetName: the workbook's first
'   sheet when it has none named yet, else a new sheet after the last
'
' WHY THIS EXISTS
'   A new workbook has one sheet; each section of the report adds one
'
' INPUTS
'   ReportBook
'     The report's workbook
'   SheetName
'     The section's sheet name (report.json)
'
' RETURNS
'   The sheet
'
' BEHAVIOR
'   - Uses the first sheet for Run, then adds sheets at the end
'   - Gives the tab the color of the workbook's sheets of its role: Run
'     that of the reference sheets, the others that of the output sheets
'     (docs/workbook/sheet-style.json)
'
' ERROR POLICY
'   Raises an error of Excel
'
' DEPENDENCIES
'   None
'
' NOTES
'   None
'
' UPDATED
'   2026-11-26
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim NewSheet   As Worksheet    'The sheet

'------------------------------------------------------------------------------
' ADD
'------------------------------------------------------------------------------
    'The first section takes the sheet the workbook was created with
        If SheetName = "Run" Then
            Set NewSheet = ReportBook.Worksheets(1)
        Else
            Set NewSheet = ReportBook.Worksheets.Add(After:=ReportBook.Worksheets(ReportBook.Worksheets.Count))
        End If
        NewSheet.Name = SheetName
    'The tab color of the workbook's sheets of the same role
        If SheetName = "Run" Then
            NewSheet.Tab.Color = RPT_TAB_REFERENCE
        Else
            NewSheet.Tab.Color = RPT_TAB_OUTPUT
        End If
        Set AddReportSheet = NewSheet

End Function


Private Function WriteReportRun( _
    ByVal ReportSheet As Worksheet) _
    As Long
'
'------------------------------------------------------------------------------
'                        WRITE THE IDENTITY OF THE RUN
'------------------------------------------------------------------------------
' PURPOSE
'   Writes the Run sheet of the report: what identifies the run, as the
'   RUN lines of the evidence CSV give it
'
' WHY THIS EXISTS
'   The first section of the report (report.json, section run)
'
' INPUTS
'   ReportSheet
'     The report's Run sheet
'
' RETURNS
'   0: the sheet is short and repeats no row on its pages
'
' BEHAVIOR
'   - One row per item: run summary, status, reporting date, currency,
'     default regime, parameter set, day count, calendar, closing days,
'     FX source, fingerprints, the state of each reconciliation and the
'     units
'   - The reporting date is a date with the format yyyy-mm-dd, which
'     reads the same in every locale
'
' ERROR POLICY
'   Raises an error of Excel
'
' DEPENDENCIES
'   ResultsStatus, StoredText, ReconciliationSheetsPresent, InputsNameOf,
'   CORE_Engine, WriteReportTable
'
' NOTES
'   A reporting date that is not a number is written as it is
'
' UPDATED
'   2026-11-26
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Items(1 To 15, 1 To 2)   As Variant    'Label and value of each row
    Dim ReportingDate            As Variant    'Reporting date of Params (Value2)
    Dim HeaderRow                As Long       'Row of the table's headers
    Dim kind                     As Long       'Reconciliation
    Dim State                    As String     'State of a reconciliation

'------------------------------------------------------------------------------
' READ THE RUN
'------------------------------------------------------------------------------
    'The items of the RUN lines of the evidence CSV, in their order
        ReportingDate = NamedCell(PRM_ASOF).Value2
        Items(1, 1) = "Run summary"
        Items(1, 2) = GetSheet(SH_RESULTS).Range(RUNINFO_CELL).Value2
        Items(2, 1) = "Status of the results"
        Items(2, 2) = ResultsStatus()
        Items(3, 1) = "Reporting date"
        Items(3, 2) = ReportingDate
        Items(4, 1) = "Reporting currency"
        Items(4, 2) = NamedCell(PRM_REPCCY).Value2
        Items(5, 1) = "Default regime"
        Items(5, 2) = NamedCell(PRM_REGIME).Value2
        Items(6, 1) = "Regulatory parameter set"
        Items(6, 2) = REF_SET_ID
        Items(7, 1) = "Days of the year"
        Items(7, 2) = CORE_Engine.YearConventionText()
        Items(8, 1) = "Calendar"
        Items(8, 2) = CORE_Engine.CalendarText()
        Items(9, 1) = "Closing-day table"
        Items(9, 2) = CORE_Engine.HolidayTableText()
        Items(10, 1) = "FX rate source"
        Items(10, 2) = CORE_Engine.FxSourceText()
        Items(11, 1) = "Input fingerprint"
        Items(11, 2) = CORE_Engine.LastRunInputs()
        Items(12, 1) = "Output fingerprint"
        Items(12, 2) = StoredText(RUN_OUTPUTS_NAME)
    'A reconciliation is CURRENT when made, else NOT MADE or UNAVAILABLE
        For kind = RECON_NS To RECON_TR
            If Not ReconciliationSheetsPresent(kind) Then
                State = "UNAVAILABLE"
            ElseIf VBA.Len(StoredText(InputsNameOf(kind))) > 0 Then
                State = "CURRENT"
            Else
                State = "NOT MADE"
            End If
            Items(12 + kind, 1) = IIf(kind = RECON_NS, "Reconciliation of netting sets", "Reconciliation of trade legs")
            Items(12 + kind, 2) = State
        Next kind
        Items(15, 1) = "Units"
        Items(15, 2) = "Amounts in " & CORE_Export.ReportCellText("text", Items(4, 2)) & _
                       ", unscaled, unless a table says otherwise"

'------------------------------------------------------------------------------
' WRITE
'------------------------------------------------------------------------------
    'A table of text, then the reporting date as a date
        HeaderRow = WriteReportTable(ReportSheet, 1, "K-SACCR run report", Array(), Array("Item", "Value"), _
                                     VBA.Split("text|text", "|"), Items, True, 0, vbNullString)
        If VarType(ReportingDate) = vbDouble Then
            ReportSheet.Cells(HeaderRow + 3, 2).NumberFormat = "yyyy-mm-dd"
            ReportSheet.Cells(HeaderRow + 3, 2).Value2 = CDbl(ReportingDate)
            ReportSheet.Cells(HeaderRow + 3, 2).HorizontalAlignment = xlLeft
        End If
        FitReportColumns ReportSheet
        WriteReportRun = 0

End Function


Private Function WriteReportTable( _
    ByVal ReportSheet As Worksheet, _
    ByVal TopRow As Long, _
    ByVal Title As String, _
    ByVal Notes As Variant, _
    ByVal Headers As Variant, _
    ByVal Kinds As Variant, _
    ByVal Data As Variant, _
    ByVal AsPdf As Boolean, _
    ByVal MeasureColumn As Long, _
    ByVal RatioMeasures As String) _
    As Long
'
'------------------------------------------------------------------------------
'                        WRITE ONE TABLE OF THE REPORT
'------------------------------------------------------------------------------
' PURPOSE
'   Writes a title, its notes, a header row and the rows of a table on a
'   sheet of the report, each column in the format of its kind, and
'   returns the header row
'
' WHY THIS EXISTS
'   Every section of the report but the Dashboard is one or two such
'   tables (report.json); the kinds decide the formats (REPORT.md)
'
' INPUTS
'   ReportSheet
'     The sheet to write on
'   TopRow
'     Row of the title
'   Title
'     The title, in bold
'   Notes
'     A 1-D array of lines written under the title; blank ones are skipped
'   Headers
'     A 1-D array of the column headers
'   Kinds
'     A 1-D array of the kind of each column, as many as Headers: amount,
'     ratio, count, date, datetime, text, or measure for a figure whose
'     format follows the Measure of its row
'   Data
'     The rows, a 2-D array with bounds from 1 and as many columns as
'     Headers; Empty for none
'   AsPdf
'     True to write text in the PDF's invariant format, False to write
'     the values with Excel number formats
'   MeasureColumn
'     Column of Data holding the measure of a row, for kind measure; 0
'     when no column has it
'   RatioMeasures
'     The measures that are not amounts, separated by "|"
'
' RETURNS
'   The row of the headers; the rows follow it
'
' BEHAVIOR
'   - The header row in white bold on the header fill of the K-SACCR
'     tables (RPT_HEADER_FILL)
'   - XLSX: each column gets the number format of its kind before the
'     values are written, text columns "@", so that a text is never read
'     as a number or a formula; a text starting with =, +, - or @ in
'     another column is written with an apostrophe
'   - PDF: every value is written as text (CORE_Export.ReportCellText),
'     numbers aligned to the right
'
' ERROR POLICY
'   Raises an error of Excel, and ERR_REPORT_VALUE through ReportCellText
'
' DEPENDENCIES
'   CORE_Export.ReportCellText, CORE_Export.ReportMeasureKind,
'   ReportNumberFormat, ReportSafeValue
'
' NOTES
'   The whole block is written in one assignment, so thousands of rows
'   take no longer than a few
'
' UPDATED
'   2026-11-26
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim HeaderRow      As Long       'Row of the headers
    Dim ColumnCount    As Long       'Columns of the table
    Dim RowCount       As Long       'Rows of data
    Dim Block          As Range      'The data cells
    Dim Values         As Variant    'What is written in them
    Dim ColumnKind     As String     'Kind of a column
    Dim CellKind       As String     'Kind of a cell
    Dim RowBase        As Long       'Lower bound of the rows of Data, less one
    Dim ColumnBase     As Long       'Lower bound of the columns of Data, less one
    Dim i              As Long       'Row of data
    Dim c              As Long       'Column
    Dim NoteIndex      As Long       'Note

'------------------------------------------------------------------------------
' WRITE THE TITLE AND THE HEADERS
'------------------------------------------------------------------------------
    'The title, its notes, then the headers as text
        ReportSheet.Cells(TopRow, 1).Value2 = Title
        ReportSheet.Cells(TopRow, 1).Font.Bold = True
        HeaderRow = TopRow + 1
        If IsArray(Notes) Then
            For NoteIndex = LBound(Notes) To UBound(Notes)
                If VBA.Len(CORE_Export.ReportCellText("text", Notes(NoteIndex))) > 0 Then
                    ReportSheet.Cells(HeaderRow, 1).NumberFormat = "@"
                    ReportSheet.Cells(HeaderRow, 1).Value2 = CORE_Export.ReportCellText("text", Notes(NoteIndex))
                    ReportSheet.Cells(HeaderRow, 1).Font.Italic = True
                    HeaderRow = HeaderRow + 1
                End If
            Next NoteIndex
        End If
        ColumnCount = UBound(Headers) - LBound(Headers) + 1
        With ReportSheet.Range(ReportSheet.Cells(HeaderRow, 1), ReportSheet.Cells(HeaderRow, ColumnCount))
            .NumberFormat = "@"
            .Value2 = Headers
            .Font.Bold = True
            .Font.Color = vbWhite
            .Interior.Color = RPT_HEADER_FILL
        End With
        WriteReportTable = HeaderRow
        If Not IsArray(Data) Then Exit Function

'------------------------------------------------------------------------------
' PREPARE THE VALUES
'------------------------------------------------------------------------------
    'One assignment for the block; the kind decides each column's format
        RowCount = UBound(Data, 1) - LBound(Data, 1) + 1
        RowBase = LBound(Data, 1) - 1
        ColumnBase = LBound(Data, 2) - 1
        Set Block = ReportSheet.Range(ReportSheet.Cells(HeaderRow + 1, 1), _
                                      ReportSheet.Cells(HeaderRow + RowCount, ColumnCount))
        ReDim Values(1 To RowCount, 1 To ColumnCount)
        For c = 1 To ColumnCount
            ColumnKind = CStr(Kinds(LBound(Kinds) + c - 1))
            For i = 1 To RowCount
                CellKind = ColumnKind
                If ColumnKind = "measure" Then
                    CellKind = CORE_Export.ReportMeasureKind(CORE_Export.ReportCellText("text", _
                                   Data(RowBase + i, ColumnBase + MeasureColumn)), RatioMeasures)
                End If
                If AsPdf Then
                    Values(i, c) = CORE_Export.ReportCellText(CellKind, Data(RowBase + i, ColumnBase + c))
                ElseIf ColumnKind = "text" Then
                    Values(i, c) = Data(RowBase + i, ColumnBase + c)
                Else
                    Values(i, c) = ReportSafeValue(Data(RowBase + i, ColumnBase + c))
                End If
            Next i
        Next c

'------------------------------------------------------------------------------
' WRITE THE VALUES
'------------------------------------------------------------------------------
    'Formats first, so that Excel reads nothing into the values
        For c = 1 To ColumnCount
            ColumnKind = CStr(Kinds(LBound(Kinds) + c - 1))
            If AsPdf Then
                Block.Columns(c).NumberFormat = "@"
                If ColumnKind <> "text" Then
                    Block.Columns(c).HorizontalAlignment = xlRight
                End If
            ElseIf ColumnKind = "measure" Then
                Block.Columns(c).NumberFormat = ReportNumberFormat("amount")
            Else
                Block.Columns(c).NumberFormat = ReportNumberFormat(ColumnKind)
            End If
        Next c
        Block.Value2 = Values
    'In the XLSX a figure of a measure that is not an amount is a ratio
        If Not AsPdf And MeasureColumn > 0 Then
            For c = 1 To ColumnCount
                If CStr(Kinds(LBound(Kinds) + c - 1)) = "measure" Then
                    For i = 1 To RowCount
                        If CORE_Export.ReportMeasureKind(CORE_Export.ReportCellText("text", _
                               Data(RowBase + i, ColumnBase + MeasureColumn)), RatioMeasures) = "ratio" Then
                            Block.Cells(i, c).NumberFormat = ReportNumberFormat("ratio")
                        End If
                    Next i
                End If
            Next c
        End If

End Function


Private Function ReportNumberFormat( _
    ByVal Kind As String) _
    As String
'
'------------------------------------------------------------------------------
'                    GIVE THE EXCEL FORMAT OF A KIND OF COLUMN
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the number format of a column of the XLSX report, in Excel's
'   invariant codes (report.json, formats)
'
' WHY THIS EXISTS
'   The XLSX writes numbers as numbers and dates as dates (#259)
'
' INPUTS
'   Kind
'     amount, ratio, count, date, datetime or text
'
' RETURNS
'   The format; "@" for text and for an unknown kind
'
' BEHAVIOR
'   One format per kind
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   None
'
' NOTES
'   tools/test_report_spec.py checks these formats against report.json
'
' UPDATED
'   2026-11-26
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'The formats of report.json
        Select Case Kind
            Case "amount"
                ReportNumberFormat = "#,##0.00"
            Case "ratio"
                ReportNumberFormat = "0.000000"
            Case "count"
                ReportNumberFormat = "0"
            Case "date"
                ReportNumberFormat = "yyyy-mm-dd"
            Case "datetime"
                ReportNumberFormat = "yyyy-mm-dd hh:mm:ss"
            Case Else
                ReportNumberFormat = "@"
        End Select

End Function


Private Function ReportSafeValue( _
    ByVal Value As Variant) _
    As Variant
'
'------------------------------------------------------------------------------
'               KEEP A TEXT FROM BEING READ AS A FORMULA IN A REPORT
'------------------------------------------------------------------------------
' PURPOSE
'   Returns a value to write into a cell of a number column of the XLSX
'   report: a text starting with =, +, - or @ gets an apostrophe in front,
'   so that Excel keeps it as text
'
' WHY THIS EXISTS
'   The XLSX holds values only, never a formula (#259)
'
' INPUTS
'   Value
'     A cell value (Value2)
'
' RETURNS
'   The value, or the text with an apostrophe
'
' BEHAVIOR
'   Leaves numbers, blanks, logical and error values as they are
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   None
'
' NOTES
'   None
'
' UPDATED
'   2026-11-26
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Only a text can become a formula
        ReportSafeValue = Value
        If VarType(Value) = vbString Then
            If VBA.Len(Value) > 0 Then
                If VBA.InStr(1, "=+-@", VBA.Left$(Value, 1), vbBinaryCompare) > 0 Then
                    ReportSafeValue = "'" & Value
                End If
            End If
        End If

End Function


Private Function WriteReportDashboard( _
    ByVal ReportSheet As Worksheet, _
    ByVal AsPdf As Boolean) _
    As Long
'
'------------------------------------------------------------------------------
'                     COPY THE DASHBOARD'S TABLES INTO THE REPORT
'------------------------------------------------------------------------------
' PURPOSE
'   Writes every block of the Dashboard's data table on the report's
'   Dashboard sheet, as the sheet shows it: its title with the unit, its
'   column headers, its row labels and its values
'
' WHY THIS EXISTS
'   The report opens with the Dashboard (#259, #675); its blocks are the
'   contract of docs/workbook/dashboard.json
'
' INPUTS
'   ReportSheet
'     The report's Dashboard sheet
'   AsPdf
'     True to write the values as text in the format of their cell, False
'     to write them with that format
'
' RETURNS
'   0: the blocks have headers of their own, so no row is repeated
'
' BEHAVIOR
'   - For each block, in the order of CORE_Engine.DashboardBlockNames,
'     copies the cells from its title row, two rows above the values, to
'     its last row, and from its label column to its last column; a block
'     without labels, a ranking, starts at its values
'   - XLSX: the values with the number format of each cell; PDF: their
'     text in that format (CORE_Export.ReportFormatText)
'   - The title row in bold, the header row in white bold on the header
'     fill of the K-SACCR tables (RPT_HEADER_FILL)
'
' ERROR POLICY
'   Raises an error of Excel, and ERR_OUTPUT_LAYOUT when a block's name is
'   missing
'
' DEPENDENCIES
'   CORE_Engine.DashboardBlockNames, CORE_Export.ReportFormatText,
'   ReportSafeValue, GetSheet
'
' NOTES
'   No formula is copied: the title's unit is the value its formula shows
'
' UPDATED
'   2026-11-26
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME       As String = "M_Main.WriteReportDashboard"

    Dim BlockNames        As Variant      'The blocks, in order
    Dim SourceSheet       As Worksheet    'The Dashboard sheet
    Dim ValueCells        As Range        'A block's values
    Dim Source            As Range        'A block with its title, headers and labels
    Dim Target            As Range        'Where the block goes in the report
    Dim Values            As Variant      'The cells of the block
    Dim FirstColumn       As Long         'First column copied
    Dim TopRow            As Long         'Next free row of the report
    Dim n                 As Long         'Block
    Dim i                 As Long         'Row of a block
    Dim c                 As Long         'Column of a block

'------------------------------------------------------------------------------
' WRITE
'------------------------------------------------------------------------------
    'A title, then each block with a blank row after it
        ReportSheet.Cells(1, 1).Value2 = "Dashboard"
        ReportSheet.Cells(1, 1).Font.Bold = True
        ReportSheet.Cells(2, 1).Value2 = "Amounts on the scale named in each table's title"
        ReportSheet.Cells(2, 1).Font.Italic = True
        TopRow = 4
        Set SourceSheet = GetSheet(SH_DASHBOARD)
        BlockNames = CORE_Engine.DashboardBlockNames()
        For n = LBound(BlockNames) To UBound(BlockNames)
            If Not NameExists(CStr(BlockNames(n))) Then
                Err.Raise ERR_OUTPUT_LAYOUT, PROC_NAME, "The Dashboard has no block " & CStr(BlockNames(n))
            End If
            Set ValueCells = ThisWorkbook.Names(CStr(BlockNames(n))).RefersToRange
        'A ranking has no label column: it starts at its values
            FirstColumn = ValueCells.Column - 1
            Set Source = SourceSheet.Range(SourceSheet.Cells(ValueCells.Row - 2, FirstColumn), _
                                           SourceSheet.Cells(ValueCells.Row + ValueCells.Rows.Count - 1, FirstColumn))
            If Application.WorksheetFunction.CountA(Source) = 0 Then
                FirstColumn = ValueCells.Column
            End If
            Set Source = SourceSheet.Range(SourceSheet.Cells(ValueCells.Row - 2, FirstColumn), _
                                           SourceSheet.Cells(ValueCells.Row + ValueCells.Rows.Count - 1, _
                                                             ValueCells.Column + ValueCells.Columns.Count - 1))
            Values = Source.Value2
            Set Target = ReportSheet.Cells(TopRow, 1).Resize(Source.Rows.Count, Source.Columns.Count)
            For i = 1 To Source.Rows.Count
                For c = 1 To Source.Columns.Count
                    If AsPdf Then
                        Target.Cells(i, c).NumberFormat = "@"
                        Target.Cells(i, c).Value2 = CORE_Export.ReportFormatText(Values(i, c), _
                                                        CStr(Source.Cells(i, c).NumberFormat))
                        If VarType(Values(i, c)) = vbDouble Then
                            Target.Cells(i, c).HorizontalAlignment = xlRight
                        End If
                    ElseIf VarType(Values(i, c)) = vbString Then
                        Target.Cells(i, c).NumberFormat = "@"
                        Target.Cells(i, c).Value2 = Values(i, c)
                    Else
                        Target.Cells(i, c).NumberFormat = Source.Cells(i, c).NumberFormat
                        Target.Cells(i, c).Value2 = ReportSafeValue(Values(i, c))
                    End If
                Next c
            Next i
            Target.Rows(1).Font.Bold = True
            Target.Rows(2).Font.Bold = True
            Target.Rows(2).Font.Color = vbWhite
            Target.Rows(2).Interior.Color = RPT_HEADER_FILL
            TopRow = TopRow + Source.Rows.Count + 1
        Next n
        FitReportColumns ReportSheet
        WriteReportDashboard = 0

End Function


Private Function NameExists( _
    ByVal NameText As String) _
    As Boolean
'
'------------------------------------------------------------------------------
'                     TELL WHETHER A WORKBOOK NAME EXISTS
'------------------------------------------------------------------------------
' PURPOSE
'   Returns True when the K-SACCR workbook has a name NameText that
'   refers to cells
'
' WHY THIS EXISTS
'   A Dashboard block missing from an older workbook is reported by name
'
' INPUTS
'   NameText
'     The workbook name
'
' RETURNS
'   True or False
'
' BEHAVIOR
'   Reads the name's range; an error means it has none
'
' ERROR POLICY
'   Never raises
'
' DEPENDENCIES
'   None
'
' NOTES
'   None
'
' UPDATED
'   2026-11-26
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim NamedRange   As Range    'The name's cells

'------------------------------------------------------------------------------
' READ
'------------------------------------------------------------------------------
    'An error of the one statement means no such name
        On Error Resume Next
        Set NamedRange = ThisWorkbook.Names(NameText).RefersToRange
        NameExists = (Err.Number = 0) And Not NamedRange Is Nothing
        On Error GoTo 0

End Function


Private Function WriteReportCharts( _
    ByVal ReportSheet As Worksheet, _
    ByRef PrintZoom As Long) _
    As Long
'
'------------------------------------------------------------------------------
'                    PICTURE THE DASHBOARD'S CHARTS IN THE REPORT
'------------------------------------------------------------------------------
' PURPOSE
'   Places a picture of each chart of the Dashboard sheet on the report's
'   Charts sheet, in rows that never cross a page, and returns how many
'
' WHY THIS EXISTS
'   The PDF shows the Dashboard's charts as the sheet shows them (#259);
'   a picture, unlike a chart, reads nothing from the K-SACCR workbook.
'   Excel draws a chart only where it shows it: exported from a sheet it
'   does not show, with screen updating off, a chart came out empty, its
'   frame and title alone. Exported at the window's 100% zoom a chart is
'   pictured at screen resolution, which the PDF showed blurred, and a
'   row of charts could be cut by a page break (owner's verification of
'   2026-11-26)
'
' INPUTS
'   ReportSheet
'     The report's Charts sheet
'
'   PrintZoom
'     Receives the print zoom of the sheet, in percent, at which its rows
'     of pictures fit the page width (LayOutReportCharts)
'
' RETURNS
'   The number of charts pictured
'
' BEHAVIOR
'   - Shows the Dashboard sheet, with screen updating on and the window
'     zoomed to RPT_CHART_ZOOM, while its charts are pictured, so that
'     each picture has three times the screen's resolution
'   - For each chart: says which in the status bar (ProgressDetail),
'     scrolls the window to it (ShowReportChart), exports it as a PNG file
'     to the temporary folder, places the picture at the chart's size and
'     deletes the file
'   - Puts back the Dashboard's zoom, scroll position and screen updating,
'     and activates the report again, on success and on failure
'   - Lays the pictures out in rows, a new page before a row that would
'     cross one (LayOutReportCharts)
'
' ERROR POLICY
'   Raises an error of Excel after deleting the picture file and putting
'   the window and screen updating back
'
' DEPENDENCIES
'   GetSheet, FileExists, ReportTempFolder, ReportSourceWindow,
'   SetReportWindowZoom, ShowReportChart, ScrollReportWindow,
'   LayOutReportCharts, ProgressDetail
'
' NOTES
'   The pictures are embedded, not linked to the files. WriteReport runs
'   with events off, so showing the Dashboard runs none of its handlers;
'   its selection is not changed
'
' UPDATED
'   2026-11-27
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "M_Main.WriteReportCharts"

    Dim SourceSheet     As Worksheet       'The Dashboard sheet
    Dim SourceWindow    As Window          'The K-SACCR window; Nothing when there is none
    Dim SavedScreen     As Boolean         'ScreenUpdating before the charts
    Dim SavedRow        As Long            'Scroll row of the Dashboard before the charts
    Dim SavedColumn     As Long            'Its scroll column
    Dim SavedZoom       As Variant         'Its zoom; Empty when it could not be read
    Dim Shown           As Boolean         'True once the Dashboard is shown
    Dim ChartFrame      As ChartObject     'A chart of the Dashboard
    Dim PictureFile     As String          'Its picture file; "" once deleted
    Dim Pictures        As Collection      'The pictures placed, in the charts' order
    Dim ChartCount      As Long            'Charts pictured
    Dim ErrorNumber     As Long            'First error; 0 on success
    Dim ErrorSource     As String          'Its source
    Dim ErrorText       As String          'Its description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        On Error GoTo ErrorHandler
    'The title of the sheet
        ReportSheet.Cells(1, 1).Value2 = "Dashboard charts"
        ReportSheet.Cells(1, 1).Font.Bold = True
        Set Pictures = New Collection

'------------------------------------------------------------------------------
' SHOW THE DASHBOARD
'------------------------------------------------------------------------------
    'Excel draws a chart only on a sheet it shows, with screen updating on
        Set SourceSheet = GetSheet(SH_DASHBOARD)
        SavedScreen = Application.ScreenUpdating
        Shown = True
        Application.ScreenUpdating = True
        ThisWorkbook.Activate
        SourceSheet.Activate
    'Keep the window, its scroll position and its zoom, to put them back;
    'a zoomed window pictures each chart at a higher resolution
        Set SourceWindow = ReportSourceWindow(SavedRow, SavedColumn)
        SavedZoom = SetReportWindowZoom(SourceWindow, RPT_CHART_ZOOM)

'------------------------------------------------------------------------------
' PICTURE THE CHARTS
'------------------------------------------------------------------------------
    'Show, export, place, delete
        For Each ChartFrame In SourceSheet.ChartObjects
            ChartCount = ChartCount + 1
            ProgressDetail "Charts: chart " & ChartCount & " of " & SourceSheet.ChartObjects.Count
            ShowReportChart SourceWindow, ChartFrame
            PictureFile = ReportTempFolder() & Application.PathSeparator & "K-SACCR_report_chart_" & ChartCount & ".png"
            If FileExists(PictureFile) Then Kill PictureFile
            ChartFrame.Chart.Export Filename:=PictureFile, FilterName:="PNG"
            Pictures.Add ReportSheet.Shapes.AddPicture(PictureFile, msoFalse, msoTrue, 0#, 0#, _
                                                       ChartFrame.Width, ChartFrame.Height)
            Kill PictureFile
            PictureFile = vbNullString
        Next ChartFrame
        WriteReportCharts = ChartCount

'------------------------------------------------------------------------------
' PUT THE DASHBOARD BACK
'------------------------------------------------------------------------------
    'The window's zoom and scroll position, screen updating, then the
    'report in front
        SetReportWindowZoom SourceWindow, SavedZoom
        ScrollReportWindow SourceWindow, SavedRow, SavedColumn
        Application.ScreenUpdating = SavedScreen
        Shown = False
        ReportSheet.Parent.Activate

'------------------------------------------------------------------------------
' LAY OUT THE PICTURES
'------------------------------------------------------------------------------
    'Rows of pictures, none across a page, at the zoom that fits the width
        PrintZoom = LayOutReportCharts(ReportSheet, Pictures)

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Leave before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
ErrorHandler:
    'Capture the first error
        ErrorNumber = Err.Number
        ErrorSource = Err.Source
        ErrorText = Err.Description
    'Delete the picture file and put the window and screen updating back
        On Error Resume Next
        If VBA.Len(PictureFile) > 0 Then
            If FileExists(PictureFile) Then Kill PictureFile
        End If
        If Shown Then
            SetReportWindowZoom SourceWindow, SavedZoom
            ScrollReportWindow SourceWindow, SavedRow, SavedColumn
            Application.ScreenUpdating = SavedScreen
        End If
        ReportSheet.Parent.Activate
    'Raise the first error
        On Error GoTo 0
        Err.Raise ErrorNumber, ErrorSource, "Chart " & ChartCount & ": " & ErrorText

End Function


Private Function LayOutReportCharts( _
    ByVal ReportSheet As Worksheet, _
    ByVal Pictures As Collection) _
    As Long
'
'------------------------------------------------------------------------------
'                 LAY OUT THE CHART PICTURES ON WHOLE PAGES
'------------------------------------------------------------------------------
' PURPOSE
'   Places the pictures of the Charts sheet in rows, left to right in the
'   charts' order, starts a new page before a row that would cross one,
'   and returns the print zoom at which the rows fit the page width
'
' WHY THIS EXISTS
'   A row of charts cut by a page break showed half charts on two pages
'   (owner's verification of 2026-11-26, #259). Excel ignores manual page
'   breaks on a sheet printed to fit a number of pages, so the Charts
'   sheet is printed at a zoom computed here instead
'
' INPUTS
'   ReportSheet
'     The report's Charts sheet
'
'   Pictures
'     The pictures, as WriteReportCharts placed them
'
' RETURNS
'   The print zoom in percent, from 10 to 100
'
' BEHAVIOR
'   - A row holds pictures while their widths, with RPT_CHART_GAP between
'     them, fit twice the widest picture: two wide charts, or three
'     narrow ones
'   - The zoom is the printable width (RPT_PRINT_WIDTH) over the widest
'     row with its margins, rounded down; a page then holds
'     RPT_PRINT_HEIGHT at that zoom
'   - Each row starts on a row of cells, from row 3; a row that would end
'     beyond its page starts a new page there (AddReportPageBreak),
'     unless it is the first row of the page
'
' ERROR POLICY
'   Raises an error of Excel; a page break Excel refuses is skipped
'
' DEPENDENCIES
'   AddReportPageBreak
'
' NOTES
'   The pictures keep their size: the zoom scales the whole sheet, as
'   the fit to one page wide did
'
' UPDATED
'   2026-11-26
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim RowLimit        As Double     'Widest a row of pictures may be
    Dim WidestRow       As Double     'Widest row laid out
    Dim RowWidth        As Double     'Width of the row being filled
    Dim RowOf()         As Long       'Row of pictures of each picture
    Dim RowCount        As Long       'Rows of pictures
    Dim Zoom            As Long       'Print zoom in percent
    Dim PageHeight      As Double     'Points of the sheet a page holds
    Dim PageTop         As Double     'Top of the current page
    Dim CellRow         As Long       'Row of cells the next row of pictures starts on
    Dim RowTop          As Double     'Top of the row of pictures
    Dim RowHeight       As Double     'Height of the row of pictures
    Dim LeftEdge        As Double     'Left of the next picture
    Dim PictureIndex    As Long       'Picture
    Dim r               As Long       'Row of pictures

'------------------------------------------------------------------------------
' FORM THE ROWS
'------------------------------------------------------------------------------
    'Nothing to lay out without pictures
        LayOutReportCharts = 100
        If Pictures.Count = 0 Then Exit Function
    'A row is at most as wide as two of the widest pictures
        For PictureIndex = 1 To Pictures.Count
            If 2# * Pictures(PictureIndex).Width + RPT_CHART_GAP > RowLimit Then
                RowLimit = 2# * Pictures(PictureIndex).Width + RPT_CHART_GAP
            End If
        Next PictureIndex
    'Fill each row left to right while the pictures fit
        ReDim RowOf(1 To Pictures.Count)
        RowCount = 1
        For PictureIndex = 1 To Pictures.Count
            If RowWidth > 0# And RowWidth + RPT_CHART_GAP + Pictures(PictureIndex).Width > RowLimit Then
                RowCount = RowCount + 1
                RowWidth = 0#
            End If
            If RowWidth > 0# Then RowWidth = RowWidth + RPT_CHART_GAP
            RowWidth = RowWidth + Pictures(PictureIndex).Width
            RowOf(PictureIndex) = RowCount
            If RowWidth > WidestRow Then WidestRow = RowWidth
        Next PictureIndex

'------------------------------------------------------------------------------
' FIND THE ZOOM
'------------------------------------------------------------------------------
    'The widest row with a gap on each side fits the printable width
        Zoom = Int(100# * RPT_PRINT_WIDTH / (WidestRow + 2# * RPT_CHART_GAP))
        If Zoom > 100 Then Zoom = 100
        If Zoom < 10 Then Zoom = 10
        PageHeight = RPT_PRINT_HEIGHT * 100# / Zoom
        LayOutReportCharts = Zoom

'------------------------------------------------------------------------------
' PLACE THE ROWS
'------------------------------------------------------------------------------
    'Each row on a row of cells; a new page before a row that would cross one
        PageTop = 0#
        CellRow = 3
        PictureIndex = 1
        For r = 1 To RowCount
            RowTop = ReportSheet.Rows(CellRow).Top
            RowHeight = 0#
            LeftEdge = RPT_CHART_GAP
            Do While PictureIndex <= Pictures.Count
                If RowOf(PictureIndex) <> r Then Exit Do
                Pictures(PictureIndex).Left = LeftEdge
                Pictures(PictureIndex).Top = RowTop
                LeftEdge = LeftEdge + Pictures(PictureIndex).Width + RPT_CHART_GAP
                If Pictures(PictureIndex).Height > RowHeight Then RowHeight = Pictures(PictureIndex).Height
                PictureIndex = PictureIndex + 1
            Loop
            If RowTop + RowHeight - PageTop > PageHeight And RowTop > PageTop Then
                AddReportPageBreak ReportSheet, CellRow
                PageTop = RowTop
            End If
        'The next row starts on the first row of cells below this one and its gap
            Do While ReportSheet.Rows(CellRow).Top < RowTop + RowHeight + RPT_CHART_GAP
                CellRow = CellRow + 1
            Loop
        Next r

End Function


Private Function SetReportWindowZoom( _
    ByVal SourceWindow As Window, _
    ByVal Zoom As Variant) _
    As Variant
'
'------------------------------------------------------------------------------
'                    ZOOM THE K-SACCR WINDOW FOR THE CHARTS
'------------------------------------------------------------------------------
' PURPOSE
'   Sets the zoom of the window and returns the zoom it had, so that it
'   can be put back
'
' WHY THIS EXISTS
'   Excel pictures a chart at the window's zoom: at RPT_CHART_ZOOM the
'   pictures of the PDF have three times the screen's resolution (#259)
'
' INPUTS
'   SourceWindow
'     The window; Nothing to do nothing
'
'   Zoom
'     The zoom in percent; Empty to do nothing
'
' RETURNS
'   The zoom the window had; Empty when there is no window or its zoom
'   could not be read
'
' BEHAVIOR
'   Reads the zoom, then sets the new one
'
' ERROR POLICY
'   Best effort: never raises; a window that keeps its zoom gives pictures
'   at the screen's resolution
'
' DEPENDENCIES
'   None
'
' NOTES
'   None
'
' UPDATED
'   2026-11-26
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'The zoom is an aid, never a failure
        On Error Resume Next
        SetReportWindowZoom = Empty

'------------------------------------------------------------------------------
' ZOOM
'------------------------------------------------------------------------------
    'Keep the zoom the window had, then set the new one
        If SourceWindow Is Nothing Then Exit Function
        If IsEmpty(Zoom) Then Exit Function
        SetReportWindowZoom = SourceWindow.Zoom
        SourceWindow.Zoom = Zoom

End Function


Private Sub AddReportPageBreak( _
    ByVal ReportSheet As Worksheet, _
    ByVal CellRow As Long)
'
'------------------------------------------------------------------------------
'                     START A NEW PAGE OF THE CHARTS SHEET
'------------------------------------------------------------------------------
' PURPOSE
'   Adds a manual page break above a row of the Charts sheet (#259)
'
' WHY THIS EXISTS
'   A row of charts starts a new page rather than being cut by one
'
' INPUTS
'   ReportSheet
'     The report's Charts sheet
'
'   CellRow
'     The row the new page starts with
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Adds the break with HPageBreaks.Add
'
' ERROR POLICY
'   Best effort: never raises; a break Excel refuses, as without a
'   printer, leaves the page to Excel
'
' DEPENDENCIES
'   None
'
' NOTES
'   None
'
' UPDATED
'   2026-11-26
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'A page break is an aid, never a failure
        On Error Resume Next

'------------------------------------------------------------------------------
' ADD
'------------------------------------------------------------------------------
    'The new page starts with CellRow
        ReportSheet.HPageBreaks.Add Before:=ReportSheet.Rows(CellRow)

End Sub


Private Function ReportSourceWindow( _
    ByRef SavedRow As Long, _
    ByRef SavedColumn As Long) _
    As Window
'
'------------------------------------------------------------------------------
'                  FIND THE K-SACCR WINDOW AND ITS SCROLL POSITION
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the first window of the K-SACCR workbook and gives its scroll
'   row and column, so that WriteReportCharts can put them back (#259)
'
' WHY THIS EXISTS
'   The charts are scrolled into view to be drawn; the window must end
'   as the user left it
'
' INPUTS
'   SavedRow
'     Receives the scroll row; 0 when it cannot be read
'
'   SavedColumn
'     Receives the scroll column; 0 when it cannot be read
'
' RETURNS
'   The window; Nothing when there is none or it cannot be read
'
' BEHAVIOR
'   Reads ThisWorkbook.Windows(1)
'
' ERROR POLICY
'   Best effort: never raises; without a window the charts are exported
'   where they are
'
' DEPENDENCIES
'   None
'
' NOTES
'   None
'
' UPDATED
'   2026-11-26
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim SourceWindow    As Window    'The window found

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'A window is an aid, never a failure
        On Error GoTo Unreadable

'------------------------------------------------------------------------------
' READ
'------------------------------------------------------------------------------
    'The window and where it is scrolled
        SavedRow = 0
        SavedColumn = 0
        Set SourceWindow = ThisWorkbook.Windows(1)
        SavedRow = SourceWindow.ScrollRow
        SavedColumn = SourceWindow.ScrollColumn
        Set ReportSourceWindow = SourceWindow

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Leave before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Unreadable:
    'No window: nothing to scroll or put back
        Set ReportSourceWindow = Nothing

End Function


Private Sub ScrollReportWindow( _
    ByVal SourceWindow As Window, _
    ByVal TopRow As Long, _
    ByVal LeftColumn As Long)
'
'------------------------------------------------------------------------------
'                  SCROLL THE K-SACCR WINDOW TO A ROW AND COLUMN
'------------------------------------------------------------------------------
' PURPOSE
'   Scrolls the window so that TopRow and LeftColumn are its first
'   scrolling row and column (#259)
'
' WHY THIS EXISTS
'   ShowReportChart brings a chart into view, and WriteReportCharts puts
'   the window back where the user left it
'
' INPUTS
'   SourceWindow
'     The window; Nothing to do nothing
'
'   TopRow
'     The row; 0 to do nothing
'
'   LeftColumn
'     The column; 0 to do nothing
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Sets ScrollRow and ScrollColumn
'
' ERROR POLICY
'   Best effort: never raises; Excel keeps the frozen rows and columns in
'   view whatever is asked
'
' DEPENDENCIES
'   None
'
' NOTES
'   None
'
' UPDATED
'   2026-11-26
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Scrolling is an aid, never a failure
        On Error Resume Next

'------------------------------------------------------------------------------
' SCROLL
'------------------------------------------------------------------------------
    'Only a window, a row and a column that are known
        If SourceWindow Is Nothing Then Exit Sub
        If TopRow < 1 Or LeftColumn < 1 Then Exit Sub
        SourceWindow.ScrollRow = TopRow
        SourceWindow.ScrollColumn = LeftColumn

End Sub


Private Sub ShowReportChart( _
    ByVal SourceWindow As Window, _
    ByVal ChartFrame As ChartObject)
'
'------------------------------------------------------------------------------
'                    BRING A CHART OF THE DASHBOARD INTO VIEW
'------------------------------------------------------------------------------
' PURPOSE
'   Scrolls the K-SACCR window to a chart of the Dashboard and lets Excel
'   draw it, before WriteReportCharts exports it
'
' WHY THIS EXISTS
'   Excel draws a chart only where it shows it; a chart below the window
'   could be exported empty (#259)
'
' INPUTS
'   SourceWindow
'     The window of the K-SACCR workbook; Nothing when there is none
'
'   ChartFrame
'     A chart of the Dashboard
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Scrolls to the chart's top left cell, refreshes the chart and lets
'   Excel draw it (DoEvents)
'
' ERROR POLICY
'   Best effort: a window that cannot scroll, as under automation with
'   Excel hidden, leaves the export to draw the chart
'
' DEPENDENCIES
'   ScrollReportWindow
'
' NOTES
'   WriteReportCharts puts the scroll position back
'
' UPDATED
'   2026-11-26
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Scrolling is an aid, never a failure
        On Error Resume Next

'------------------------------------------------------------------------------
' SCROLL AND DRAW
'------------------------------------------------------------------------------
    'Bring the chart's top left cell into view
        ScrollReportWindow SourceWindow, ChartFrame.TopLeftCell.Row, ChartFrame.TopLeftCell.Column
    'Draw the chart before it is exported
        ChartFrame.Chart.Refresh
        DoEvents

End Sub


Private Function ReportTempFolder() As String

'
'------------------------------------------------------------------------------
'                    FIND A FOLDER FOR THE REPORT'S PICTURES
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the temporary folder of the user, or the workbook's folder
'   when there is none
'
' WHY THIS EXISTS
'   A chart is pictured through a file that is deleted at once (#259)
'
' INPUTS
'   None
'
' RETURNS
'   A folder path without a final separator
'
' BEHAVIOR
'   Reads the TEMP variable
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   None
'
' NOTES
'   None
'
' UPDATED
'   2026-11-26
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'TEMP, or the workbook's folder
        ReportTempFolder = VBA.Environ$("TEMP")
        If VBA.Len(ReportTempFolder) = 0 Then
            ReportTempFolder = ThisWorkbook.Path
        End If

End Function


Private Function ReadReportTable( _
    ByVal SheetName As String, _
    ByVal ColumnCount As Long, _
    ByRef Headers As Variant) _
    As Variant
'
'------------------------------------------------------------------------------
'                     READ AN OUTPUT TABLE FOR THE REPORT
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the rows of an output table, as the evidence CSV reads them,
'   and its headers
'
' WHY THIS EXISTS
'   The report copies Results, Checks and the reconciliations as the CSV
'   does (#259): from the header row down to the last row with a value in
'   the first column
'
' INPUTS
'   SheetName
'     The output sheet
'   ColumnCount
'     Its columns (RS_NCOLS, CK_NCOLS, RC_NCOLS, TRC_NCOLS)
'   Headers
'     Written: a 1-D array of the headers, from 1
'
' RETURNS
'   A 2-D array of the rows, from 1; Empty when the table has none
'
' BEHAVIOR
'   Reads the header row above OutputFirstRow and the rows below it
'
' ERROR POLICY
'   Raises an error of Excel or of the sheet's layout
'
' DEPENDENCIES
'   GetSheet, OutputFirstRow, LastDataRow, TableRange
'
' NOTES
'   None
'
' UPDATED
'   2026-11-26
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim TableSheet   As Worksheet    'The sheet
    Dim HeaderRow    As Long         'Its header row
    Dim LastRow      As Long         'Its last row
    Dim HeaderRange  As Variant      'The header row's values
    Dim Names()      As Variant      'The headers
    Dim c            As Long         'Column

'------------------------------------------------------------------------------
' READ
'------------------------------------------------------------------------------
    'Headers, then the rows when there are any
        Set TableSheet = GetSheet(SheetName)
        HeaderRow = OutputFirstRow(SheetName) - 1
        HeaderRange = TableRange(TableSheet, HeaderRow, 1, HeaderRow, ColumnCount).Value2
        ReDim Names(1 To ColumnCount)
        For c = 1 To ColumnCount
            Names(c) = HeaderRange(1, c)
        Next c
        Headers = Names
        LastRow = LastDataRow(TableSheet, HeaderRow + 1, 1)
        If LastRow > HeaderRow Then
            ReadReportTable = TableRange(TableSheet, HeaderRow + 1, 1, LastRow, ColumnCount).Value2
        Else
            ReadReportTable = Empty
        End If

End Function


Private Function WriteReportTotals( _
    ByVal ReportSheet As Worksheet, _
    ByVal AsPdf As Boolean) _
    As Long
'
'------------------------------------------------------------------------------
'                     WRITE THE TOTALS AND STATUS PER REGIME
'------------------------------------------------------------------------------
' PURPOSE
'   Writes the Totals sheet of the report: the TOTAL rows of Results, one
'   per regime, with the columns of report.json (section totals)
'
' WHY THIS EXISTS
'   The report gives the totals and status per regime before the netting
'   sets (#259)
'
' INPUTS
'   ReportSheet
'     The report's Totals sheet
'   AsPdf
'     As WriteReportTable
'
' RETURNS
'   The header row, repeated on every page of a PDF
'
' BEHAVIOR
'   Reads Results, keeps its rows TOTAL CRR and TOTAL BCBS, the IDs the
'   engine reserves, and the columns named in RPT_TOTALS_HEADERS, found by
'   their headers; a netting set named TOTAL CLIENT is not a total
'
' ERROR POLICY
'   Raises ERR_OUTPUT_LAYOUT when a header is missing from Results
'
' DEPENDENCIES
'   ReadReportTable, WriteReportTable, FitReportColumns
'
' NOTES
'   None
'
' UPDATED
'   2026-11-26
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME       As String = "M_Main.WriteReportTotals"

    Dim Headers           As Variant    'Headers of Results
    Dim TableRows         As Variant    'Rows of Results
    Dim Wanted            As Variant    'Headers of the totals
    Dim SourceColumns()   As Long       'Column of Results of each header of the totals
    Dim Totals            As Variant    'The TOTAL rows
    Dim TotalCount        As Long       'TOTAL rows found
    Dim ResultId          As String     'NettingSetID of a row of Results
    Dim r                 As Long       'Row of Results
    Dim c                 As Long       'Column of the totals
    Dim k                 As Long       'Column of Results

'------------------------------------------------------------------------------
' READ
'------------------------------------------------------------------------------
    'The columns of the totals, by their headers on Results
        TableRows = ReadReportTable(SH_RESULTS, RS_NCOLS, Headers)
        Wanted = VBA.Split(RPT_TOTALS_HEADERS, "|")
        ReDim SourceColumns(LBound(Wanted) To UBound(Wanted))
        For c = LBound(Wanted) To UBound(Wanted)
            For k = 1 To RS_NCOLS
                If CORE_Export.ReportCellText("text", Headers(k)) = Wanted(c) Then Exit For
            Next k
            If k > RS_NCOLS Then
                Err.Raise ERR_OUTPUT_LAYOUT, PROC_NAME, "Results has no column " & Wanted(c)
            End If
            SourceColumns(c) = k
        Next c
    'The TOTAL rows, in their order
        If IsArray(TableRows) Then
            ReDim Totals(1 To UBound(TableRows, 1), 1 To UBound(Wanted) - LBound(Wanted) + 1)
            For r = 1 To UBound(TableRows, 1)
                ResultId = CORE_Export.ReportCellText("text", TableRows(r, 1))
                If ResultId = "TOTAL " & RG_CRR Or ResultId = "TOTAL " & RG_BCBS Then
                    TotalCount = TotalCount + 1
                    For c = LBound(Wanted) To UBound(Wanted)
                        Totals(TotalCount, c - LBound(Wanted) + 1) = TableRows(r, SourceColumns(c))
                    Next c
                End If
            Next r
        End If
        If TotalCount = 0 Then
            Totals = Empty
        Else
            Totals = TopRows(Totals, TotalCount)
        End If

'------------------------------------------------------------------------------
' WRITE
'------------------------------------------------------------------------------
    'One table
        WriteReportTotals = WriteReportTable(ReportSheet, 1, "Totals and status per regime", _
                         Array("Amounts in " & ReportCurrencyText() & ", unscaled."), Wanted, _
                         VBA.Split(RPT_TOTALS_KINDS, "|"), Totals, AsPdf, 0, vbNullString)
        FitReportColumns ReportSheet

End Function


Private Function TopRows( _
    ByVal Block As Variant, _
    ByVal RowCount As Long) _
    As Variant
'
'------------------------------------------------------------------------------
'                      KEEP THE FIRST ROWS OF A 2-D ARRAY
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the first RowCount rows of a 2-D array with bounds from 1
'
' WHY THIS EXISTS
'   ReDim Preserve can change only the last bound, and the report fills
'   a block before it knows how many of its rows it needs
'
' INPUTS
'   Block
'     A 2-D array with bounds from 1
'   RowCount
'     Rows to keep, 1 or more and no more than Block has
'
' RETURNS
'   A 2-D array of RowCount rows and the columns of Block
'
' BEHAVIOR
'   Copies the cells
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   None
'
' NOTES
'   None
'
' UPDATED
'   2026-11-26
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Result   As Variant    'The rows kept
    Dim i        As Long       'Row
    Dim c        As Long       'Column

'------------------------------------------------------------------------------
' COPY
'------------------------------------------------------------------------------
    'The first rows, every column
        ReDim Result(1 To RowCount, 1 To UBound(Block, 2))
        For i = 1 To RowCount
            For c = 1 To UBound(Block, 2)
                Result(i, c) = Block(i, c)
            Next c
        Next i
        TopRows = Result

End Function


Private Function ReportCurrencyText() As String

'
'------------------------------------------------------------------------------
'                      NAME THE CURRENCY OF THE AMOUNTS
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the reporting currency of Params as text, for the notes of
'   the report's tables
'
' WHY THIS EXISTS
'   Every amount of the report states its unit (#259)
'
' INPUTS
'   None
'
' RETURNS
'   The currency code, such as EUR
'
' BEHAVIOR
'   Reads ReportingCcy
'
' ERROR POLICY
'   Raises an error of the workbook's layout
'
' DEPENDENCIES
'   NamedCell, CORE_Export.ReportCellText
'
' NOTES
'   None
'
' UPDATED
'   2026-11-26
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'The code as written on Params
        ReportCurrencyText = CORE_Export.ReportCellText("text", NamedCell(PRM_REPCCY).Value2)

End Function


Private Function WriteReportResults( _
    ByVal ReportSheet As Worksheet, _
    ByVal AsPdf As Boolean) _
    As Long
'
'------------------------------------------------------------------------------
'                      WRITE THE RESULTS PER NETTING SET
'------------------------------------------------------------------------------
' PURPOSE
'   Writes the Results sheet of the report: every row of Results, its
'   netting sets and its TOTAL rows, with its 28 columns
'
' WHY THIS EXISTS
'   The report gives the result of each netting set (#259)
'
' INPUTS
'   ReportSheet
'     The report's Results sheet
'   AsPdf
'     As WriteReportTable
'
' RETURNS
'   The header row, repeated on every page of a PDF
'
' BEHAVIOR
'   Copies the table with the kinds of RPT_RESULTS_KINDS
'
' ERROR POLICY
'   Raises an error of Excel
'
' DEPENDENCIES
'   ReadReportTable, WriteReportTable, FitReportColumns
'
' NOTES
'   None
'
' UPDATED
'   2026-11-26
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Headers     As Variant    'Headers of Results
    Dim TableRows   As Variant    'Rows of Results

'------------------------------------------------------------------------------
' WRITE
'------------------------------------------------------------------------------
    'The table as it is
        TableRows = ReadReportTable(SH_RESULTS, RS_NCOLS, Headers)
        WriteReportResults = WriteReportTable(ReportSheet, 1, "Results per netting set", _
                         Array("Amounts in " & ReportCurrencyText() & ", unscaled."), Headers, _
                         VBA.Split(RPT_RESULTS_KINDS, "|"), TableRows, AsPdf, 0, vbNullString)
        FitReportColumns ReportSheet

End Function


Private Function WriteReportChecks( _
    ByVal ReportSheet As Worksheet, _
    ByVal AsPdf As Boolean) _
    As Long
'
'------------------------------------------------------------------------------
'                           WRITE THE CHECKS BY CODE
'------------------------------------------------------------------------------
' PURPOSE
'   Writes the Checks sheet of the report: a row per code with its
'   highest severity and its rows, then the rows of Checks grouped by
'   code: every row in the XLSX, at most RPT_PDF_CHECKS_PER_CODE of each
'   code in the PDF
'
' WHY THIS EXISTS
'   The owner's decisions of 2026-11-26 on Checks by code (#259): with a
'   real portfolio Checks can hold a row per trade
'
' INPUTS
'   ReportSheet
'     The report's Checks sheet
'   AsPdf
'     As WriteReportTable
'
' RETURNS
'   The header row of the rows of Checks, repeated on every page of a PDF
'
' BEHAVIOR
'   - Orders the rows with CORE_Export.ReportCodeOrder
'   - PDF: keeps the first RPT_PDF_CHECKS_PER_CODE rows of each code and
'     says how many are left out (CORE_Export.ReportChecksNote)
'   - Writes the two tables
'
' ERROR POLICY
'   Raises an error of Excel
'
' DEPENDENCIES
'   ReadReportTable, CORE_Export.ReportCodeOrder,
'   CORE_Export.ReportCellText, CORE_Export.ReportChecksNote,
'   WriteReportTable, FitReportColumns, TopRows
'
' NOTES
'   The severity is the first column of Checks and the code the seventh;
'   the summary by code always counts every row. The rows of a code are
'   together in ReportCodeOrder's order, so a change of the code's text
'   starts the next code
'
' UPDATED
'   2026-11-26
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Headers     As Variant    'Headers of Checks
    Dim TableRows   As Variant    'Rows of Checks
    Dim Summary     As Variant    'A row per code
    Dim Order       As Variant    'Rows in the order of the report
    Dim Ordered     As Variant    'The rows written, in that order
    Dim KeptCount   As Long       'Rows written
    Dim LeftOut     As Long       'Rows a PDF leaves out
    Dim Code        As String     'Code of a row
    Dim LastCode    As String     'Code of the row before
    Dim OfCode      As Long       'Rows of the current code so far
    Dim HeaderRow   As Long       'Header row of the summary
    Dim i           As Long       'Row of the report
    Dim c           As Long       'Column

'------------------------------------------------------------------------------
' ORDER THE ROWS
'------------------------------------------------------------------------------
    'Codes by severity, then code; each code's rows as on the sheet; a PDF
    'keeps the first of each code
        TableRows = ReadReportTable(SH_CHECKS, CK_NCOLS, Headers)
        Order = CORE_Export.ReportCodeOrder(TableRows, 1, CK_CODE_COL, Summary)
        If IsArray(TableRows) Then
            ReDim Ordered(1 To UBound(TableRows, 1), 1 To CK_NCOLS)
            For i = 1 To UBound(TableRows, 1)
                Code = VBA.Trim$(CORE_Export.ReportCellText("text", TableRows(Order(i), CK_CODE_COL)))
                If i = 1 Or Code <> LastCode Then
                    OfCode = 0
                    LastCode = Code
                End If
                OfCode = OfCode + 1
                If AsPdf And OfCode > RPT_PDF_CHECKS_PER_CODE Then
                    LeftOut = LeftOut + 1
                Else
                    KeptCount = KeptCount + 1
                    For c = 1 To CK_NCOLS
                        Ordered(KeptCount, c) = TableRows(Order(i), c)
                    Next c
                End If
            Next i
            Ordered = TopRows(Ordered, KeptCount)
        Else
            Ordered = Empty
        End If

'------------------------------------------------------------------------------
' WRITE
'------------------------------------------------------------------------------
    'The summary, then the rows, with what a PDF leaves out
        HeaderRow = WriteReportTable(ReportSheet, 1, "Checks by code", Array(), Array("Code", "Severity", "Rows"), _
                                     VBA.Split("text|text|count", "|"), Summary, AsPdf, 0, vbNullString)
        WriteReportChecks = WriteReportTable(ReportSheet, HeaderRow + RowsOf(Summary) + 2, "Checks, grouped by code", _
                                             Array(CORE_Export.ReportChecksNote(RPT_PDF_CHECKS_PER_CODE, LeftOut)), _
                                             Headers, VBA.Split(RPT_CHECKS_KINDS, "|"), Ordered, AsPdf, 0, _
                                             vbNullString)
        FitReportColumns ReportSheet

End Function


Private Function RowsOf( _
    ByVal Block As Variant) _
    As Long
'
'------------------------------------------------------------------------------
'                         COUNT THE ROWS OF A 2-D ARRAY
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the rows of a 2-D array; 0 when it is not an array
'
' WHY THIS EXISTS
'   A table of the report with no row is Empty (#259)
'
' INPUTS
'   Block
'     A 2-D array, or Empty
'
' RETURNS
'   The number of rows
'
' BEHAVIOR
'   Reads the bounds of the first dimension
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   None
'
' NOTES
'   None
'
' UPDATED
'   2026-11-26
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Empty has no row
        If IsArray(Block) Then
            RowsOf = UBound(Block, 1) - LBound(Block, 1) + 1
        Else
            RowsOf = 0
        End If

End Function


Private Function WriteReportReconciliation( _
    ByVal ReportSheet As Worksheet, _
    ByVal ReconKind As Long, _
    ByVal AsPdf As Boolean) _
    As Long
'
'------------------------------------------------------------------------------
'                      WRITE A RECONCILIATION OF THE RUN
'------------------------------------------------------------------------------
' PURPOSE
'   Writes a reconciliation sheet of the report: the summary of the
'   reconciliation and its rows. The XLSX has every row; the PDF the rows
'   of netting sets that are not MATCH, or the largest differences of the
'   trade legs, with a line saying what is left out
'
' WHY THIS EXISTS
'   The owner's decisions of 2026-11-26 on the reconciliation rows (#259):
'   a real portfolio gave a PDF of 163 pages, most of them trade legs
'   not in the bank's file
'
' INPUTS
'   ReportSheet
'     The report's sheet of the reconciliation
'   ReconKind
'     RECON_NS or RECON_TR
'   AsPdf
'     As WriteReportTable
'
' RETURNS
'   The header row, repeated on every page of a PDF
'
' BEHAVIOR
'   - Reads the summary line of the sheet and its table
'   - PDF, netting sets: keeps the rows whose Result is not MATCH
'   - PDF, trade legs: keeps the DIFFERENCE rows with the largest absolute
'     difference, at most RPT_PDF_TRADE_DIFFERENCES (LargestDifferences)
'   - The figures take the format of their measure (report.json,
'     measure_ratios)
'
' ERROR POLICY
'   Raises an error of Excel
'
' DEPENDENCIES
'   ReadReportTable, ReconSheetOf, LargestDifferences,
'   CORE_Export.ReportOmittedNote, CORE_Export.ReportDifferencesNote,
'   WriteReportTable, FitReportColumns
'
' NOTES
'   The measure is the sixth column, the absolute difference the ninth and
'   the result the eleventh in both tables
'
' UPDATED
'   2026-11-26
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const MEASURE_COLUMN     As Long = 6      'Measure, in both tables
    Const DIFFERENCE_COLUMN  As Long = 9      'Absolute difference, in both tables
    Const RESULT_COLUMN      As Long = 11     'Result, in both tables

    Dim SheetName            As String        'The reconciliation sheet
    Dim ColumnCount          As Long          'Its columns
    Dim Headers              As Variant       'Its headers
    Dim TableRows            As Variant       'Its rows
    Dim Kept                 As Variant       'The rows written
    Dim KeptCount            As Long          'Rows written
    Dim Omitted              As Long          'MATCH rows left out of a PDF
    Dim Differences          As Long          'DIFFERENCE rows of the trade legs
    Dim LeftNote             As String        'What a PDF leaves out
    Dim Summary              As String        'The summary line of the sheet
    Dim Title                As String        'Title of the table
    Dim Notes                As Variant       'Lines under the title
    Dim Kinds                As Variant       'Kind of each column
    Dim i                    As Long          'Row
    Dim c                    As Long          'Column

'------------------------------------------------------------------------------
' READ
'------------------------------------------------------------------------------
    'The summary line and the table
        SheetName = ReconSheetOf(ReconKind)
        ColumnCount = IIf(ReconKind = RECON_NS, RC_NCOLS, TRC_NCOLS)
        Summary = CORE_Export.ReportCellText("text", GetSheet(SheetName).Range(RECON_INFO_CELL).Value2)
        TableRows = ReadReportTable(SheetName, ColumnCount, Headers)
        Kept = TableRows

'------------------------------------------------------------------------------
' CHOOSE THE ROWS OF A PDF
'------------------------------------------------------------------------------
    'Trade legs: the largest differences, the rest counted by the summary
        If AsPdf And ReconKind = RECON_TR Then
            Kept = LargestDifferences(TableRows, RESULT_COLUMN, DIFFERENCE_COLUMN, RPT_PDF_TRADE_DIFFERENCES, _
                                      Differences, KeptCount)
            LeftNote = CORE_Export.ReportDifferencesNote(KeptCount, Differences, RowsOf(TableRows) - KeptCount)
    'Netting sets: the rows that are not MATCH
        ElseIf AsPdf And IsArray(TableRows) Then
            ReDim Kept(1 To UBound(TableRows, 1), 1 To ColumnCount)
            For i = 1 To UBound(TableRows, 1)
                If VBA.UCase$(VBA.Trim$(CORE_Export.ReportCellText("text", TableRows(i, RESULT_COLUMN)))) = "MATCH" Then
                    Omitted = Omitted + 1
                Else
                    KeptCount = KeptCount + 1
                    For c = 1 To ColumnCount
                        Kept(KeptCount, c) = TableRows(i, c)
                    Next c
                End If
            Next i
            If KeptCount = 0 Then
                Kept = Empty
            Else
                Kept = TopRows(Kept, KeptCount)
            End If
            LeftNote = CORE_Export.ReportOmittedNote(Omitted)
        End If

'------------------------------------------------------------------------------
' WRITE
'------------------------------------------------------------------------------
    'The summary, the units, what is left out, then the rows
        Title = IIf(ReconKind = RECON_NS, "Reconciliation of netting sets", "Reconciliation of trade legs")
        Notes = Array(Summary, "Amounts in units of " & ReportCurrencyText() & _
                      "; the bank's figures multiplied by the scale of the settings; ratios are not scaled.", _
                      LeftNote)
        Kinds = VBA.Split(IIf(ReconKind = RECON_NS, RPT_RECON_KINDS, RPT_TRADE_RECON_KINDS), "|")
        WriteReportReconciliation = WriteReportTable(ReportSheet, 1, Title, Notes, Headers, Kinds, Kept, AsPdf, _
                                                     MEASURE_COLUMN, IIf(ReconKind = RECON_NS, RPT_RECON_RATIOS, _
                                                                         RPT_TRADE_RECON_RATIOS))
        FitReportColumns ReportSheet

End Function


Private Function LargestDifferences( _
    ByVal TableRows As Variant, _
    ByVal ResultColumn As Long, _
    ByVal DifferenceColumn As Long, _
    ByVal Limit As Long, _
    ByRef Differences As Long, _
    ByRef KeptCount As Long) _
    As Variant
'
'------------------------------------------------------------------------------
'              CHOOSE THE LARGEST DIFFERENCES OF A RECONCILIATION
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the rows whose Result is DIFFERENCE with the largest absolute
'   difference, at most Limit of them, largest first; ties keep their
'   order on the sheet
'
' WHY THIS EXISTS
'   The PDF shows the largest differences of the trade legs (#259, owner's
'   decision of 2026-11-26); tools/report.py (pdf_trade_rows) is the
'   reference
'
' INPUTS
'   TableRows
'     The rows of the reconciliation, a 2-D array from 1; Empty for none
'   ResultColumn
'     Column of the Result
'   DifferenceColumn
'     Column of the absolute difference
'   Limit
'     Rows kept at most
'   Differences
'     Receives the number of DIFFERENCE rows
'   KeptCount
'     Receives the number of rows returned
'
' RETURNS
'   A 2-D array of the rows kept, from 1; Empty when none
'
' BEHAVIOR
'   Picks, Limit times, the remaining DIFFERENCE row of the largest
'   absolute difference, the first on the sheet among equals: a pass per
'   row kept, so thousands of rows take no time to speak of
'
' ERROR POLICY
'   None: a difference that is not a number counts as 0
'
' DEPENDENCIES
'   CORE_Export.ReportCellText, CORE_Util.ToDbl
'
' NOTES
'   None
'
' UPDATED
'   2026-11-26
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim RowCount        As Long       'Rows of the table
    Dim ColumnCount     As Long       'Its columns
    Dim Candidate()     As Boolean    'A DIFFERENCE row not yet kept
    Dim Size()          As Double     'Its absolute difference
    Dim Kept            As Variant    'The rows kept
    Dim Best            As Long       'Row of the largest difference left
    Dim i               As Long       'Row
    Dim c               As Long       'Column

'------------------------------------------------------------------------------
' FIND THE DIFFERENCES
'------------------------------------------------------------------------------
    'Nothing to keep without rows
        Differences = 0
        KeptCount = 0
        LargestDifferences = Empty
        If Not IsArray(TableRows) Then Exit Function
        RowCount = UBound(TableRows, 1)
        ColumnCount = UBound(TableRows, 2)
    'Mark the DIFFERENCE rows and their size
        ReDim Candidate(1 To RowCount)
        ReDim Size(1 To RowCount)
        For i = 1 To RowCount
            If VBA.UCase$(VBA.Trim$(CORE_Export.ReportCellText("text", TableRows(i, ResultColumn)))) = _
                    "DIFFERENCE" Then
                Candidate(i) = True
                Size(i) = VBA.Abs(CORE_Util.ToDbl(TableRows(i, DifferenceColumn), 0#))
                Differences = Differences + 1
            End If
        Next i
        If Differences = 0 Or Limit < 1 Then Exit Function

'------------------------------------------------------------------------------
' KEEP THE LARGEST
'------------------------------------------------------------------------------
    'The largest left, the first among equals, until the limit
        ReDim Kept(1 To IIf(Differences < Limit, Differences, Limit), 1 To ColumnCount)
        Do While KeptCount < UBound(Kept, 1)
            Best = 0
            For i = 1 To RowCount
                If Candidate(i) Then
                    If Best = 0 Then
                        Best = i
                    ElseIf Size(i) > Size(Best) Then
                        Best = i
                    End If
                End If
            Next i
            KeptCount = KeptCount + 1
            For c = 1 To ColumnCount
                Kept(KeptCount, c) = TableRows(Best, c)
            Next c
            Candidate(Best) = False
        Loop
        LargestDifferences = Kept

End Function


Private Sub FitReportColumns( _
    ByVal ReportSheet As Worksheet)
'
'------------------------------------------------------------------------------
'                      FIT THE COLUMNS OF A REPORT SHEET
'------------------------------------------------------------------------------
' PURPOSE
'   Fits each used column of a report sheet to its contents, at most 60
'   characters wide; a narrowed column wraps its text and its rows grow to
'   show it whole
'
' WHY THIS EXISTS
'   A long message must neither widen a page beyond reading nor be cut off
'   by the next column in the PDF (#259)
'
' INPUTS
'   ReportSheet
'     A sheet of the report
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   - AutoFit, then a cap of 60 on a wider column
'   - A capped column wraps its cells, except a title or a note: a cell of
'     column A alone on its row, which overflows to the right as it is
'   - The rows are then fitted to their wrapped text
'
' ERROR POLICY
'   Raises an error of Excel
'
' DEPENDENCIES
'   None
'
' NOTES
'   None
'
' UPDATED
'   2026-11-26
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const MAX_WIDTH   As Double = 60#    'Characters

    Dim UsedColumn    As Range           'A used column
    Dim ColumnCell    As Range           'A cell of a capped column
    Dim AnyCapped     As Boolean         'A column was narrowed

'------------------------------------------------------------------------------
' FIT
'------------------------------------------------------------------------------
    'Fit, then cap and wrap
        ReportSheet.UsedRange.Columns.AutoFit
        For Each UsedColumn In ReportSheet.UsedRange.Columns
            If UsedColumn.ColumnWidth > MAX_WIDTH Then
                UsedColumn.ColumnWidth = MAX_WIDTH
                AnyCapped = True
                If UsedColumn.Column > 1 Then
                    UsedColumn.WrapText = True
                Else
                    For Each ColumnCell In UsedColumn.Cells
                        If Application.WorksheetFunction.CountA(ColumnCell.EntireRow) > 1 Then
                            ColumnCell.WrapText = True
                        End If
                    Next ColumnCell
                End If
            End If
        Next UsedColumn
    'Rows as high as their wrapped text
        If AnyCapped Then
            ReportSheet.UsedRange.Rows.AutoFit
        End If

End Sub


Private Function ApplyReportPageSetup( _
    ByVal ReportSheet As Worksheet, _
    ByVal Footer As String, _
    ByVal TitleRow As Long, _
    Optional ByVal PrintZoom As Long = 0) _
    As String
'
'------------------------------------------------------------------------------
'                     SET UP THE PAGES OF A PDF REPORT SHEET
'------------------------------------------------------------------------------
' PURPOSE
'   Sets a sheet of the PDF to A4 landscape, one page wide, with the
'   header row of its table repeated and the report's footer; returns
'   what could not be set
'
' WHY THIS EXISTS
'   The PDF's pages (report.json, pdf); a machine whose printer driver
'   refuses a setting still writes the report, and says so
'
' INPUTS
'   ReportSheet
'     A sheet of the report
'   Footer
'     The footer, with &P and &N for the page and the pages
'   TitleRow
'     The header row repeated on every page; 0 for none
'   PrintZoom
'     The print zoom in percent, for a sheet with its own page breaks
'     (the Charts sheet); 0, the default, to fit one page wide
'
' RETURNS
'   "" when every setting was applied, else "sheet: error; "
'
' BEHAVIOR
'   - Communication with the printer is paused while the settings are
'     written
'   - The margins are RPT_MARGIN_SIDE and RPT_MARGIN_END, which
'     LayOutReportCharts lays the charts out for
'   - One page wide, or the zoom given: Excel ignores manual page breaks
'     on a sheet printed to fit a number of pages
'
' ERROR POLICY
'   Never raises: a failure is returned
'
' DEPENDENCIES
'   None
'
' NOTES
'   PrintCommunication exists from Excel 2010 on
'
' UPDATED
'   2026-11-26
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' SET UP THE PAGES
'------------------------------------------------------------------------------
    'A4 landscape, one page wide, the footer
        Application.PrintCommunication = False
        With ReportSheet.PageSetup
            .Orientation = xlLandscape
            .PaperSize = xlPaperA4
            .LeftMargin = RPT_MARGIN_SIDE
            .RightMargin = RPT_MARGIN_SIDE
            .TopMargin = RPT_MARGIN_END
            .BottomMargin = RPT_MARGIN_END
            If PrintZoom > 0 Then
                .Zoom = PrintZoom
            Else
                .Zoom = False
                .FitToPagesWide = 1
                .FitToPagesTall = False
            End If
            If TitleRow > 0 Then
                .PrintTitleRows = "$" & TitleRow & ":$" & TitleRow
            End If
            .CenterFooter = Footer
        End With
        Application.PrintCommunication = True
        ApplyReportPageSetup = vbNullString

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
ErrorHandler:
    'Report the failure and let the printer settings through again
        ApplyReportPageSetup = ReportSheet.Name & ": " & Err.Description & "; "
        On Error Resume Next
        Application.PrintCommunication = True
        On Error GoTo 0

End Function


Private Function ReportRunTime() As String

'
'------------------------------------------------------------------------------
'                     READ THE TIME OF THE RUN FOR THE FOOTER
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the date and time of the last run, as the run summary on
'   Results writes it after "Last run "
'
' WHY THIS EXISTS
'   The footer of every page of the PDF names the run (#259)
'
' INPUTS
'   None
'
' RETURNS
'   Such as 2026-11-26 12:48:55; "" when the summary has none
'
' BEHAVIOR
'   Reads the 19 characters after "Last run " at the start of Results C4
'
' ERROR POLICY
'   Raises an error of Excel
'
' DEPENDENCIES
'   GetSheet, CORE_Export.ReportCellText
'
' NOTES
'   None
'
' UPDATED
'   2026-11-26
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Summary   As String    'The run summary

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'The time follows "Last run "
        Summary = CORE_Export.ReportCellText("text", GetSheet(SH_RESULTS).Range(RUNINFO_CELL).Value2)
        If VBA.Left$(Summary, 9) = "Last run " Then
            ReportRunTime = VBA.Mid$(Summary, 10, 19)
        Else
            ReportRunTime = vbNullString
        End If

End Function


Private Sub HideReportGridlines( _
    ByVal ReportBook As Workbook)
'
'------------------------------------------------------------------------------
'                    HIDE THE GRIDLINES OF THE RUN REPORT
'------------------------------------------------------------------------------
' PURPOSE
'   Hides the gridlines of every sheet of the report, as the workbook's
'   sheets have none (docs/workbook/sheet-style.json)
'
' WHY THIS EXISTS
'   An XLSX report opens as the K-SACCR sheets look (#259); a PDF never
'   prints gridlines
'
' INPUTS
'   ReportBook
'     The report's workbook
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Turns off the gridlines of each sheet's view in the report's window,
'   without activating a sheet
'
' ERROR POLICY
'   Best effort: a window without sheet views, as under automation, keeps
'   its gridlines, which change no value
'
' DEPENDENCIES
'   None
'
' NOTES
'   WorksheetView needs Excel 2013 or later, as the workbook does
'
' UPDATED
'   2026-11-26
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim View    As WorksheetView    'The view of a sheet in the report's window

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'The gridlines are a matter of look, never a failure
        On Error Resume Next

'------------------------------------------------------------------------------
' HIDE
'------------------------------------------------------------------------------
    'Every sheet's view in the report's window, without activating a sheet
        For Each View In ReportBook.Windows(1).SheetViews
            View.DisplayGridlines = False
        Next View

End Sub


Private Function ReportSiblingPath( _
    ByVal FilePath As String) _
    As String
'
'------------------------------------------------------------------------------
'                FIND A NAME BESIDE THE REPORT WITH ITS EXTENSION
'------------------------------------------------------------------------------
' PURPOSE
'   Returns a file name in the folder of FilePath that does not exist yet
'   and ends with the same extension, such as "saccr-new1-report.pdf"
'
' WHY THIS EXISTS
'   The report is saved beside its target and then renamed, so that a
'   failure leaves an older file as it was (#259). Excel saves an XLSX
'   only under a name ending in .xlsx
'
' INPUTS
'   FilePath
'     The target
'
' RETURNS
'   The new name, with its folder
'
' BEHAVIOR
'   Puts "saccr-new" and a counter in front of the file name
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   FileExists
'
' NOTES
'   None
'
' UPDATED
'   2026-11-26
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim SlashAt     As Long      'Position of the last separator; 0 when none
    Dim Folder      As String    'The folder, with its separator
    Dim FileName    As String    'The file name
    Dim n           As Long      'Counter that makes the name unique
    Dim Candidate   As String    'Name tried

'------------------------------------------------------------------------------
' FIND A FREE NAME
'------------------------------------------------------------------------------
    'Same folder, same extension
        SlashAt = VBA.InStrRev(FilePath, Application.PathSeparator, -1, vbBinaryCompare)
        Folder = VBA.Left$(FilePath, SlashAt)
        FileName = VBA.Mid$(FilePath, SlashAt + 1)
        Do
            n = n + 1
            Candidate = Folder & "saccr-new" & n & "-" & FileName
        Loop While FileExists(Candidate)
        ReportSiblingPath = Candidate

End Function


Private Sub ReplaceReportFile( _
    ByVal NewPath As String, _
    ByVal FilePath As String)
'
'------------------------------------------------------------------------------
'                   REPLACE A REPORT FILE BY THE ONE JUST SAVED
'------------------------------------------------------------------------------
' PURPOSE
'   Renames the saved report NewPath to FilePath, replacing an older
'   file, hidden ones included
'
' WHY THIS EXISTS
'   A failed export leaves an older report as it was, as the evidence
'   CSV does (#312)
'
' INPUTS
'   NewPath
'     The report just saved
'   FilePath
'     The target
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   - Moves an older target aside, renames the new file and deletes the
'     older one
'   - On a failure puts the older file back with its attributes
'
' ERROR POLICY
'   Raises the first error after putting the older file back; the new
'   file is deleted by the caller
'
' DEPENDENCIES
'   FileExists, SiblingPath
'
' NOTES
'   gTestFault "report-replace" raises while the older file is aside
'
' UPDATED
'   2026-11-26
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME           As String = "M_Main.ReplaceReportFile"

    Dim OldPath               As String     'The older file while it is aside; "" otherwise
    Dim OldAttributes         As Long       'Its attributes, put back on a failure
    Dim AttributesCleared     As Boolean    'Its attributes were cleared
    Dim ErrorNumber           As Long       'First error; 0 on success
    Dim ErrorSource           As String     'Its source
    Dim ErrorText             As String     'Its description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' REPLACE
'------------------------------------------------------------------------------
    'Move the older file aside, then rename the new one
        If FileExists(FilePath) Then
            OldPath = SiblingPath(FilePath, "old")
            OldAttributes = GetAttr(FilePath) And (vbReadOnly Or vbHidden Or vbSystem Or vbArchive)
            SetAttr FilePath, vbNormal
            AttributesCleared = True
            Name FilePath As OldPath
        End If
        If gTestFault = "report-replace" Then
            Err.Raise ERR_INJECTED_FAULT, PROC_NAME, "Injected failure while the older report was replaced."
        End If
        Name NewPath As FilePath
    'The older file is no longer needed
        If VBA.Len(OldPath) > 0 Then
            Kill OldPath
        End If

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
ErrorHandler:
    'Put the older file back, then raise the first error
        ErrorNumber = Err.Number
        ErrorSource = Err.Source
        ErrorText = Err.Description
        On Error Resume Next
        If VBA.Len(OldPath) > 0 Then
            If Not FileExists(FilePath) Then Name OldPath As FilePath
        End If
        If AttributesCleared Then SetAttr FilePath, OldAttributes
        On Error GoTo 0
        Err.Raise ErrorNumber, ErrorSource, ErrorText

End Sub


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
'   identity, with the day count, the calendar and the source of the FX
'   rates), PARAMS, SF_TABLE, FX_TABLE and HOLIDAY_TABLE (the parameter
'   snapshot),
'   RESULTS, RECONCILIATION, TRADE_RECONCILIATION, CHECKS and LIMITATIONS.
'   Everything is read from the workbook, so the report is the same after
'   the workbook is closed and reopened.
'
' UPDATED
'   2026-11-22
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
    'How the dates became years and where the rates came from (#721, #745):
    'the inputs are those of the run, so these are the run's.
        Report = Report & RunLine("year_convention", CORE_Engine.YearConventionText())
        Report = Report & RunLine("calendar", CORE_Engine.CalendarText())
        Report = Report & RunLine("holiday_table", CORE_Engine.HolidayTableText())
        Report = Report & RunLine("fx_source", CORE_Engine.FxSourceText())
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
        Report = Report & HolidayTableLines()

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


Private Function HolidayTableLines() As String
'
'==============================================================================
'                             HolidayTableLines
'------------------------------------------------------------------------------
' PURPOSE
'   The Holiday_Table of Params as the run reads it (#721): its header row,
'   then every one of its HOLIDAY_ROWS rows with a closing day, also those
'   below a blank row, as the engine reads them all.
'
' ERROR POLICY
'   Raises when the name does not locate a header row.
'
' UPDATED
'   2026-11-20
'==============================================================================
'

    Dim ParamsSheet   As Worksheet    'Params
    Dim HeaderRow     As Long         'Header row
    Dim KeyColumn     As Long         'First column of the table
    Dim Block         As Variant      'Header and rows
    Dim r             As Long         'Row read

        Set ParamsSheet = GetSheet(SH_PARAMS)
        HeaderRow = NamedRow(NM_HOLIDAY_TABLE)
        KeyColumn = NamedTableCol(NM_HOLIDAY_TABLE)
        If HeaderRow = 0 Then
            Err.Raise ERR_PARAMS_LAYOUT, "M_Main.HolidayTableLines", "The name " & NM_HOLIDAY_TABLE & " is missing."
        End If
        Block = TableRange(ParamsSheet, HeaderRow, KeyColumn, HeaderRow + HOLIDAY_ROWS, KeyColumn + 1).Value2
        HolidayTableLines = CORE_Export.CsvLine("HOLIDAY_TABLE", Array(Block(1, 1), Block(1, 2)))
        For r = 2 To HOLIDAY_ROWS + 1
            If IsError(Block(r, 1)) Or Not IsBlankCell(Block(r, 1)) Then
                HolidayTableLines = HolidayTableLines & CORE_Export.CsvLine("HOLIDAY_TABLE", Array(Block(r, 1), Block(r, 2)))
            End If
        Next r

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
'------------------------------------------------------------------------------
'                    REMEMBER THE OUTPUTS OF A COMPLETED RUN
'------------------------------------------------------------------------------
' PURPOSE
'   Stores the fingerprint of the output sheets when a run has completed,
'   its reconciliations included, and that of the Dashboard's data table
'
' WHY THIS EXISTS
'   The evidence export and the report refuse an output edited after the
'   run (#45); the report also copies the Dashboard, which the fingerprint
'   of the outputs does not cover (#259, Codex on #805)
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Writes RUN_OUTPUTS_NAME (OutputFingerprint) and RUN_DASHBOARD_NAME
'   (DashboardFingerprint) as hidden workbook names
'
' ERROR POLICY
'   Best effort: when a fingerprint cannot be taken or stored both are
'   emptied, and the exports refuse until the next run
'
' DEPENDENCIES
'   OutputFingerprint, DashboardFingerprint, ForgetRunOutputs
'
' NOTES
'   The run writes the Dashboard before it calls this procedure
'
' UPDATED
'   2026-11-26
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Empty both names on any error
        On Error GoTo Failed

'------------------------------------------------------------------------------
' STORE
'------------------------------------------------------------------------------
    'The outputs the evidence CSV copies, then the Dashboard
        ThisWorkbook.Names.Add Name:=RUN_OUTPUTS_NAME, RefersTo:="=""" & OutputFingerprint() & """", _
                               Visible:=False
        ThisWorkbook.Names.Add Name:=RUN_DASHBOARD_NAME, RefersTo:="=""" & DashboardFingerprint() & """", _
                               Visible:=False

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Leave before the error handler
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Failed:
    'Leave nothing that a later export could take for the run's
        ForgetRunOutputs

End Sub


Private Sub ForgetRunOutputs()
'
'------------------------------------------------------------------------------
'                    FORGET THE OUTPUTS OF THE LAST RUN
'------------------------------------------------------------------------------
' PURPOSE
'   Empties the stored fingerprints of the output sheets and of the
'   Dashboard
'
' WHY THIS EXISTS
'   Every operation starts by forgetting the last run, so that the exports
'   refuse until a run completes again
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Writes RUN_OUTPUTS_NAME and RUN_DASHBOARD_NAME empty
'
' ERROR POLICY
'   Contains any error
'
' DEPENDENCIES
'   None
'
' NOTES
'   None
'
' UPDATED
'   2026-11-26
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'An empty name is all that is needed; an error leaves it as it is
        On Error Resume Next

'------------------------------------------------------------------------------
' FORGET
'------------------------------------------------------------------------------
    'Both fingerprints empty
        ThisWorkbook.Names.Add Name:=RUN_OUTPUTS_NAME, RefersTo:="=""""", Visible:=False
        ThisWorkbook.Names.Add Name:=RUN_DASHBOARD_NAME, RefersTo:="=""""", Visible:=False

End Sub


Private Function DashboardFingerprint() As String
'
'------------------------------------------------------------------------------
'                    FINGERPRINT THE DASHBOARD'S DATA TABLE
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the fingerprint of the values of every block of the
'   Dashboard's data table, the figures the report copies
'
' WHY THIS EXISTS
'   A figure of the Dashboard changed after the run, with the sheet
'   unprotected, would otherwise reach the report beside the run's
'   fingerprints and a CURRENT status (Codex on #805, #259)
'
' INPUTS
'   None
'
' RETURNS
'   The fingerprint (TextHash); "" when the Dashboard has no block
'
' BEHAVIOR
'   Joins the evidence text of each value of each block, in the order of
'   CORE_Engine.DashboardBlockNames, and fingerprints it; a block whose
'   name is missing adds its name, so that it cannot go unnoticed
'
' ERROR POLICY
'   Raises an error of Excel
'
' DEPENDENCIES
'   CORE_Engine.DashboardBlockNames, CORE_Export.CsvField, NameExists,
'   TextHash
'
' NOTES
'   The status cell of the Dashboard, a formula, is not a block and is not
'   fingerprinted: it turns out of date by itself when the inputs change
'
' UPDATED
'   2026-11-26
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim BlockNames      As Variant    'The blocks, in order
    Dim Values          As Variant    'The values of a block
    Dim Text            As String     'The text fingerprinted
    Dim n               As Long       'Block
    Dim i               As Long       'Row of a block
    Dim c               As Long       'Column of a block

'------------------------------------------------------------------------------
' FINGERPRINT
'------------------------------------------------------------------------------
    'Every value of every block, a separator between them
        BlockNames = CORE_Engine.DashboardBlockNames()
        For n = LBound(BlockNames) To UBound(BlockNames)
            If NameExists(CStr(BlockNames(n))) Then
                Values = ThisWorkbook.Names(CStr(BlockNames(n))).RefersToRange.Value2
                If IsArray(Values) Then
                    For i = LBound(Values, 1) To UBound(Values, 1)
                        For c = LBound(Values, 2) To UBound(Values, 2)
                            Text = Text & Chr$(31) & CORE_Export.CsvField(Values(i, c))
                        Next c
                    Next i
                Else
                    Text = Text & Chr$(31) & CORE_Export.CsvField(Values)
                End If
            Else
                Text = Text & Chr$(29) & "missing " & CStr(BlockNames(n))
            End If
            Text = Text & Chr$(29)
        Next n
        If VBA.Len(Text) = 0 Then
            DashboardFingerprint = vbNullString
        Else
            DashboardFingerprint = TextHash(Text)
        End If

End Function


Private Function ReportProblem() As String
'
'------------------------------------------------------------------------------
'                    TELL WHY THE REPORT CANNOT BE WRITTEN
'------------------------------------------------------------------------------
' PURPOSE
'   Returns why the formatted report of the run cannot be written; ""
'   when it can
'
' WHY THIS EXISTS
'   The report refuses in the cases of Export evidence (EvidenceProblem)
'   and also when the Dashboard, which it copies and the evidence CSV does
'   not, was changed after the run (#259, Codex on #805)
'
' INPUTS
'   None
'
' RETURNS
'   The reason, as EvidenceProblem gives it; "" when the report can be
'   written
'
' BEHAVIOR
'   - EvidenceProblem first
'   - Then the stored fingerprint of the Dashboard against the Dashboard as
'     it is: missing when the run recorded none, as for a workbook run
'     before this check, different when a figure was changed
'
' ERROR POLICY
'   Raises an error of Excel
'
' DEPENDENCIES
'   EvidenceProblem, StoredText, DashboardFingerprint
'
' NOTES
'   None
'
' UPDATED
'   2026-11-26
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Stored      As String    'The Dashboard's fingerprint at the run

'------------------------------------------------------------------------------
' CHECK
'------------------------------------------------------------------------------
    'The rules of the evidence export
        ReportProblem = EvidenceProblem()
        If VBA.Len(ReportProblem) > 0 Then Exit Function
    'The Dashboard as the run wrote it
        Stored = StoredText(RUN_DASHBOARD_NAME)
        If VBA.Len(Stored) = 0 Then
            ReportProblem = "the run recorded no Dashboard; press Run SA-CCR first"
        ElseIf Stored <> DashboardFingerprint() Then
            ReportProblem = "a figure on the Dashboard was changed after the run; press Run SA-CCR first"
        End If

End Function


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


Public Sub ExportBankNettingSetsCsvTemplate()
'
'==============================================================================
'                         ExportBankNettingSetsCsvTemplate
'------------------------------------------------------------------------------
' PURPOSE
'   Write the blank BankNettingSets CSV template of the input contract
'   (#718) to a file chosen by the user: its header line only.
'
' USAGE
'   BankNettingSets template on the K-SACCR Ribbon tab (Templates menu), the
'   Export CSV template button on BankNettingSets, or the Macros dialog.
'
' ERROR POLICY
'   As ExportCsvTemplate; never raises.
'
' UPDATED
'   2026-11-20
'==============================================================================
'
        ExportCsvTemplate "BankNettingSets"

End Sub


Public Sub ExportBankTradesCsvTemplate()
'
'==============================================================================
'                           ExportBankTradesCsvTemplate
'------------------------------------------------------------------------------
' PURPOSE
'   Write the blank BankTrades CSV template of the input contract (#718) to
'   a file chosen by the user: its header line only.
'
' USAGE
'   BankTrades template on the K-SACCR Ribbon tab (Templates menu), the
'   Export CSV template button on BankTrades, or the Macros dialog.
'
' ERROR POLICY
'   As ExportCsvTemplate; never raises.
'
' UPDATED
'   2026-11-20
'==============================================================================
'
        ExportCsvTemplate "BankTrades"

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
            Case "BankNettingSets"
                CsvTemplateText = CSV_BN_HEADER & vbLf
            Case "BankTrades"
                CsvTemplateText = CSV_BT_HEADER & vbLf
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
            Case "BankNettingSets"
                CsvTemplateFile = CSV_BN_FILE
            Case "BankTrades"
                CsvTemplateFile = CSV_BT_FILE
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
'------------------------------------------------------------------------------
'                 CHECK A CSV FILE AGAINST ITS INPUT CONTRACT
'------------------------------------------------------------------------------
' PURPOSE
'   Asks for a CSV file, reads it as a file of Kind and reports whether it
'   follows the contract: the number of records of an accepted file, or
'   the number of findings and the first of them, up to
'   CSV_MESSAGE_FINDINGS (CsvCheckMessage)
'
' WHY THIS EXISTS
'   A file can be checked before it is imported; nothing is ever written
'   to the workbook
'
' INPUTS
'   Kind
'     The kind of file, such as "Trades" or "NettingSets"
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   - A cancelled dialog does nothing
'   - Shows "reading and checking" in the status bar while the file is
'     read (#807); the message gives the status bar back
'
' ERROR POLICY
'   Reports a file that cannot be read in a message, or in the Immediate
'   window in silent mode; never raises
'
' DEPENDENCIES
'   ReadCsvFile (CORE_CsvReader), CsvCheckMessage, ProgressStart,
'   ProgressShow, Notify
'
' NOTES
'   Runs from CheckTradesCsv and the other Check macros
'
' UPDATED
'   2026-11-27
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "M_Main.CheckCsvFile"

    Dim ChosenFile      As Variant          'Chosen file, or False when cancelled
    Dim Result          As CsvReadResult    'What the reader found
    Dim Values()        As String           'The records of an accepted file
    Dim Lines()         As Long             'Their lines
    Dim Findings()      As CsvFinding       'The findings

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' ASK AND READ
'------------------------------------------------------------------------------
    'Ask for the file; a cancelled dialog does nothing
        ChosenFile = Application.GetOpenFilename(FileFilter:="CSV files (*.csv), *.csv", _
                                                 Title:="Check " & Kind & " CSV file")
        If VarType(ChosenFile) = vbBoolean Then
            Exit Sub
        End If
    'Read and check it, its progress in the status bar
        ProgressStart "Check " & Kind
        ProgressShow 0, 1, "reading and checking " & CutText(CStr(ChosenFile), CSV_MESSAGE_PATH, True)
        ReadCsvFile CStr(ChosenFile), Kind, Result, Values, Lines, Findings

'------------------------------------------------------------------------------
' REPORT
'------------------------------------------------------------------------------
    'Say what the file holds; the message gives the status bar back
        Notify CsvCheckMessage(CStr(ChosenFile), Result, Findings), _
               IIf(Result.Accepted, vbInformation, vbExclamation)

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
ErrorHandler:
    'Say what failed
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


Public Sub ImportBankNettingSetsCsv()
'
'==============================================================================
'                           ImportBankNettingSetsCsv
'------------------------------------------------------------------------------
' PURPOSE
'   Load a BankNettingSets CSV file chosen by the user onto BankNettingSets
'   (#718): the bank's figures replace those of the sheet when the whole
'   file follows the CSV contract, and nothing changes otherwise.
'
' USAGE
'   BankNettingSets file on the K-SACCR Ribbon tab (Import CSV menu), the
'   Import CSV button on BankNettingSets, or the Macros dialog.
'
' ERROR POLICY
'   As ImportBankCsv; never raises.
'
' UPDATED
'   2026-11-20
'==============================================================================
'
        ImportBankCsv "BankNettingSets"
        RibbonStateChanged

End Sub


Public Sub ImportBankTradesCsv()
'
'==============================================================================
'                             ImportBankTradesCsv
'------------------------------------------------------------------------------
' PURPOSE
'   Load a BankTrades CSV file chosen by the user onto BankTrades (#718),
'   as ImportBankNettingSetsCsv does for BankNettingSets.
'
' USAGE
'   BankTrades file on the K-SACCR Ribbon tab (Import CSV menu), the Import
'   CSV button on BankTrades, or the Macros dialog.
'
' ERROR POLICY
'   As ImportBankCsv; never raises.
'
' UPDATED
'   2026-11-20
'==============================================================================
'
        ImportBankCsv "BankTrades"
        RibbonStateChanged

End Sub


Public Function ImportBankCsv_Silent( _
    ByVal Kind As String, _
    ByVal FilePath As String, _
    Optional ByVal ReplaceExisting As Boolean = False) _
    As String
'
'==============================================================================
'                             ImportBankCsv_Silent
'------------------------------------------------------------------------------
' PURPOSE
'   The bank-figure load for automation and tests (#718): read FilePath as a
'   file of Kind ("BankNettingSets" or "BankTrades") and write its records
'   onto the bank sheet without asking. A sheet that already holds bank
'   figures is replaced only when ReplaceExisting is True.
'
' RETURNS
'   "RESULT=LOADED; records=N; file=..." when the records were written;
'   "RESULT=REJECTED; findings=N; first=CODE|line|record|column; file=..."
'   when the file does not follow the contract, a file with more records
'   than the sheet holds included; "RESULT=REFUSED; reason=...; file=..."
'   when the file has no record, or the sheet holds figures and
'   ReplaceExisting is False. The sheet is changed only when the result is
'   LOADED.
'
' ERROR POLICY
'   An unknown kind and a file that cannot be read raise ERR_CSV_READ. An
'   SA-CCR operation in progress raises ERR_RUN_ACTIVE, and a bank sheet
'   or name that is missing, and a value longer than a cell holds, raise
'   ERR_STAGING_WRITE, before the sheet changes; an error while the sheet
'   is written is raised after the sheet has been put back as it was.
'
' UPDATED
'   2026-11-20
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ReconKind   As Long             'The reconciliation the sheet belongs to
    Dim Result      As CsvReadResult    'What the reader found
    Dim Values()    As String           'The records of a readable file
    Dim Lines()     As Long             'Their lines
    Dim Findings()  As CsvFinding       'The findings

'------------------------------------------------------------------------------
' READ AND LOAD
'------------------------------------------------------------------------------
        ReconKind = BankReconKind(Kind)
        ReadCsvFile FilePath, Kind, Result, Values, Lines, Findings, 0, True
        If Not Result.Accepted Then
            ImportBankCsv_Silent = "RESULT=REJECTED; findings=" & Result.FindingCount & _
                                   "; first=" & FindingText(Findings(1)) & "; file=" & FilePath
            Exit Function
        End If
        If Result.RecordCount = 0 Then
            ImportBankCsv_Silent = "RESULT=REFUSED; reason=the file has no record; file=" & FilePath
            Exit Function
        End If
        If Not ReplaceExisting Then
            If BankRowsEntered(ReconKind) Then
                ImportBankCsv_Silent = "RESULT=REFUSED; reason=" & BankSheetOf(ReconKind) & _
                                       " holds bank figures; file=" & FilePath
                Exit Function
            End If
        End If
        LoadBank ReconKind, Result, Values
        ImportBankCsv_Silent = "RESULT=LOADED; records=" & Result.RecordCount & "; file=" & FilePath

End Function


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
'------------------------------------------------------------------------------
'                    IMPORT A CSV FILE INTO ITS STAGING SHEET
'------------------------------------------------------------------------------
' PURPOSE
'   Asks for a CSV file, reads it as a file of Kind and loads its records
'   into the staging sheet as text, with the CSV line of each, then shows
'   the outcome and the sheet
'
' WHY THIS EXISTS
'   The records are loaded as they are written in the file, so that the
'   findings about values can be corrected on the sheet before they are
'   validated (docs/IMPORT.md)
'
' INPUTS
'   Kind
'     The kind of file, such as "Trades" or "NettingSets"
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   - A cancelled dialog changes nothing
'   - Shows the progress in the status bar by phases (#807): reading and
'     checking the file, then writing its records
'   - A structural finding loads nothing; findings about values are
'     listed and stay to be corrected on the sheet
'   - When the sheet has changes since its last load, asks before they
'     are replaced; a No changes nothing
'
' ERROR POLICY
'   Reports a file that cannot be read or loaded in a message, or in the
'   Immediate window in silent mode; never raises. The staging sheet is
'   put back as it was when its writing fails (LoadStaging)
'
' DEPENDENCIES
'   StagingSheetOf, ReadCsvFile (CORE_CsvReader), CsvCheckMessage,
'   StagingHasEdits, ReplaceConfirmed, LoadStaging, CutText,
'   ProgressStart, ProgressShow, ProgressEnd, Notify
'
' NOTES
'   Every message gives the status bar back (Notify); the question gets it
'   back first (ProgressEnd)
'
' UPDATED
'   2026-11-27
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "M_Main.ImportCsv"

    Dim SheetName       As String           'The staging sheet
    Dim ChosenFile      As Variant          'Chosen file, or False when cancelled
    Dim Result          As CsvReadResult    'What the reader found
    Dim Values()        As String           'The records of a readable file
    Dim Lines()         As Long             'Their lines
    Dim Findings()      As CsvFinding       'The findings
    Dim Note            As String           'What was loaded

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' ASK AND READ
'------------------------------------------------------------------------------
    'Ask for the file; a cancelled dialog changes nothing
        SheetName = StagingSheetOf(Kind)
        ChosenFile = Application.GetOpenFilename(FileFilter:="CSV files (*.csv), *.csv", _
                                                 Title:="Import " & Kind & " CSV file into " & SheetName)
        If VarType(ChosenFile) = vbBoolean Then
            Exit Sub
        End If
    'Read and check it, its progress in the status bar
        ProgressStart "Import " & Kind
        ProgressShow 0, 2, "reading and checking " & CutText(CStr(ChosenFile), CSV_MESSAGE_PATH, True)
        ReadCsvFile CStr(ChosenFile), Kind, Result, Values, Lines, Findings, 0, True
    'A file that cannot be read loads nothing
        If Not Result.Readable Then
            Notify CsvCheckMessage(CStr(ChosenFile), Result, Findings, _
                                   "Nothing was loaded: " & SheetName & " is unchanged."), vbExclamation
            Exit Sub
        End If

'------------------------------------------------------------------------------
' CONFIRM
'------------------------------------------------------------------------------
    'Changes made on the sheet are replaced only when the user says so
        If StagingHasEdits(Kind) Then
            ProgressEnd
            If Not ReplaceConfirmed(Kind, CStr(ChosenFile)) Then
                Exit Sub
            End If
            ProgressStart "Import " & Kind
        End If

'------------------------------------------------------------------------------
' LOAD
'------------------------------------------------------------------------------
    'Write the records to the staging sheet
        ProgressShow 1, 2, "writing " & Result.RecordCount & " records to " & SheetName
        LoadStaging Kind, CStr(ChosenFile), Result, Values, Lines
    'Say what was loaded and what is left to do
        If Result.Accepted Then
            Note = "Loaded into " & SheetName & " as text." & vbLf & "Not validated yet; " & Kind & " is unchanged."
        Else
            Note = Result.RecordCount & " records loaded into " & SheetName & " as text: correct the findings there." & _
                   vbLf & Kind & " is unchanged."
        End If
        ThisWorkbook.Worksheets(SheetName).Activate
        Notify CsvCheckMessage(CStr(ChosenFile), Result, Findings, Note), _
               IIf(Result.Accepted, vbInformation, vbExclamation)

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
ErrorHandler:
    'Say what failed; the message gives the status bar back
        Notify Kind & " CSV import: " & Err.Description, vbCritical

End Sub


Private Sub ImportBankCsv( _
    ByVal Kind As String)
'
'------------------------------------------------------------------------------
'                     IMPORT A BANK CSV FILE INTO ITS SHEET
'------------------------------------------------------------------------------
' PURPOSE
'   The interactive load of a bank file (#718): asks for the file, reads
'   it, says what is wrong with it when it does not follow the contract,
'   asks before replacing figures that are on the sheet, and writes the
'   records
'
' WHY THIS EXISTS
'   The bank figures are compared with the run's (docs/IMPORT.md); the
'   file is accepted or rejected as a whole, so a record is never dropped
'   or truncated and a file with more records than the sheet holds is
'   rejected
'
' INPUTS
'   Kind
'     BankNettingSets or BankTrades
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   - A cancelled dialog changes nothing
'   - Shows the progress in the status bar by phases (#807): reading and
'     checking the file, then writing its records
'   - A file that is not accepted, or holds no record, loads nothing
'   - Asks before replacing the figures on the sheet, except in silent
'     mode; a No changes nothing
'
' ERROR POLICY
'   Reported to the user, never raised
'
' DEPENDENCIES
'   BankReconKind, BankSheetOf, ReadCsvFile (CORE_CsvReader),
'   CsvCheckMessage, CutText, BankRowsEntered, LoadBank, ProgressStart,
'   ProgressShow, ProgressEnd, Notify
'
' NOTES
'   Every message gives the status bar back (Notify); the question gets it
'   back first (ProgressEnd)
'
' UPDATED
'   2026-11-27
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "M_Main.ImportBankCsv"

    Dim ReconKind       As Long             'The reconciliation the sheet belongs to
    Dim SheetName       As String           'The bank sheet
    Dim ChosenFile      As Variant          'Chosen file, or False when cancelled
    Dim Result          As CsvReadResult    'What the reader found
    Dim Values()        As String           'The records of a readable file
    Dim Lines()         As Long             'Their lines
    Dim Findings()      As CsvFinding       'The findings

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' ASK AND READ
'------------------------------------------------------------------------------
    'Ask for the file; a cancelled dialog changes nothing
        ReconKind = BankReconKind(Kind)
        SheetName = BankSheetOf(ReconKind)
        ChosenFile = Application.GetOpenFilename(FileFilter:="CSV files (*.csv), *.csv", _
                                                 Title:="Import " & Kind & " CSV file into " & SheetName)
        If VarType(ChosenFile) = vbBoolean Then
            Exit Sub
        End If
    'Read and check it, its progress in the status bar
        ProgressStart "Import " & Kind
        ProgressShow 0, 2, "reading and checking " & CutText(CStr(ChosenFile), CSV_MESSAGE_PATH, True)
        ReadCsvFile CStr(ChosenFile), Kind, Result, Values, Lines, Findings, 0, True
    'The file is accepted as a whole, or loads nothing
        If Not Result.Accepted Then
            Notify CsvCheckMessage(CStr(ChosenFile), Result, Findings, _
                                   "Nothing was loaded: " & SheetName & " is unchanged."), vbExclamation
            Exit Sub
        End If
        If Result.RecordCount = 0 Then
            Notify CutText(CStr(ChosenFile), CSV_MESSAGE_PATH, True) & vbLf & vbLf & _
                   "The file holds no record. Nothing was loaded: " & SheetName & " is unchanged.", vbExclamation
            Exit Sub
        End If

'------------------------------------------------------------------------------
' CONFIRM
'------------------------------------------------------------------------------
    'Figures on the sheet are replaced only when the user says so
        If BankRowsEntered(ReconKind) Then
            If Not gSilent Then
                ProgressEnd
                If MsgBox(SheetName & " holds bank figures." & vbLf & vbLf & "Replace them with the " & _
                          Result.RecordCount & " records of " & CutText(CStr(ChosenFile), CSV_MESSAGE_PATH, True) & "?", _
                          vbYesNo + vbExclamation + vbDefaultButton2, MSG_TITLE) <> vbYes Then
                    Exit Sub
                End If
                ProgressStart "Import " & Kind
            End If
        End If

'------------------------------------------------------------------------------
' LOAD
'------------------------------------------------------------------------------
    'Write the records to the bank sheet and say so
        ProgressShow 1, 2, "writing " & Result.RecordCount & " records to " & SheetName
        LoadBank ReconKind, Result, Values
        ThisWorkbook.Worksheets(SheetName).Activate
        Notify CsvCheckMessage(CStr(ChosenFile), Result, Findings, _
                               "Loaded into " & SheetName & ": " & Result.RecordCount & " records." & vbLf & _
                               "The amount scale and the tolerances of the sheet are unchanged; press Run SA-CCR to reconcile."), _
               vbInformation

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
ErrorHandler:
    'Say what failed; the message gives the status bar back
        Notify Kind & " CSV import: " & Err.Description, vbCritical

End Sub


Private Function BankReconKind( _
    ByVal Kind As String) _
    As Long
'
'==============================================================================
'                               BankReconKind
'------------------------------------------------------------------------------
' PURPOSE
'   The reconciliation (RECON_NS or RECON_TR) of the bank file Kind.
'
' ERROR POLICY
'   Raises ERR_CSV_READ for a kind that is not a bank file.
'
' UPDATED
'   2026-11-20
'==============================================================================
'
        Select Case Kind
            Case "BankNettingSets"
                BankReconKind = RECON_NS
            Case "BankTrades"
                BankReconKind = RECON_TR
            Case Else
                Err.Raise ERR_CSV_READ, "M_Main.BankReconKind", "unknown bank CSV kind '" & Kind & "'"
        End Select

End Function


Private Sub LoadBank( _
    ByVal ReconKind As Long, _
    ByRef Result As CsvReadResult, _
    ByRef Values() As String)
'
'==============================================================================
'                                  LoadBank
'------------------------------------------------------------------------------
' PURPOSE
'   Write the records of an accepted bank file onto the bank sheet (#718),
'   from the first row below the header: the identifiers, the regime and
'   the comment as text, the figures as numbers, a blank figure as an empty
'   cell. The figures and comments of the sheet are cleared first. The
'   amount scale and the tolerances, above the table, are not touched.
'
' ERROR POLICY
'   Raises ERR_RUN_ACTIVE while an operation runs, ERR_STAGING_WRITE for a
'   missing bank header and a value longer than a cell holds, both before
'   the sheet changes. An error while writing puts the sheet back as it
'   was and is raised; settings that could not be restored are added to
'   its description.
'
' UPDATED
'   2026-11-20
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Kind                As String          'BankNettingSets or BankTrades
    Dim Sheet               As Worksheet       'The bank sheet
    Dim HeaderRow           As Long            'Its header row
    Dim FirstRow            As Long            'First row below the header
    Dim Width               As Long            'Columns of the table
    Dim Spec                As CsvSpec         'The file's contract
    Dim Columns()           As CsvColumn       'Its columns
    Dim Rules()             As CsvRule         'Its rules
    Dim NewBlock()          As Variant         'The records as they will be written
    Dim OldBlock            As Variant         'The block before the load, as formulas
    Dim OldLast             As Long            'Its last used row; below FirstRow for none
    Dim LastRow             As Long            'Last row written or cleared
    Dim Cleared             As Boolean         'The old block has been cleared
    Dim SavedCalculation    As XlCalculation   'Calculation mode before the load
    Dim SavedEvents         As Boolean         'EnableEvents before the load
    Dim SavedScreen         As Boolean         'ScreenUpdating before the load
    Dim StateCaptured       As Boolean         'True once all three were read
    Dim ErrorNumber         As Long            'Primary error number; 0 for none
    Dim ErrorSource         As String          'Primary error source
    Dim ErrorDescription    As String          'Primary error description
    Dim CleanupDetails      As String          'Settings not restored
    Dim r                   As Long            'Record
    Dim c                   As Long            'Column

'------------------------------------------------------------------------------
' PREPARE
'------------------------------------------------------------------------------
        If mRunning Then
            Err.Raise ERR_RUN_ACTIVE, "M_Main.LoadBank", "An SA-CCR operation is already running."
        End If
        Kind = IIf(ReconKind = RECON_NS, "BankNettingSets", "BankTrades")
        HeaderRow = BankHeaderRow(ReconKind)
        If HeaderRow = 0 Then
            Err.Raise ERR_STAGING_WRITE, "M_Main.LoadBank", _
                      "The name " & HeaderNameOf(ReconKind) & " is missing: build the workbook from the current template."
        End If
        Set Sheet = GetSheet(BankSheetOf(ReconKind))
        FirstRow = HeaderRow + 1
        Width = BankColsOf(ReconKind)
        If Not CsvSpecOf(Kind, Spec, Columns, Rules) Then
            Err.Raise ERR_CSV_READ, "M_Main.LoadBank", "unknown CSV kind '" & Kind & "'"
        End If
        If Spec.ColumnCount <> Width Then
            Err.Raise ERR_STAGING_WRITE, "M_Main.LoadBank", _
                      "The CSV contract has " & Spec.ColumnCount & " columns and " & Sheet.Name & " has " & Width & "."
        End If
        ReDim NewBlock(1 To Result.RecordCount, 1 To Width)
        For r = 1 To Result.RecordCount
            For c = 1 To Width
                If Len(Values(r, c)) > EXCEL_CELL_CHARS Then
                    Err.Raise ERR_STAGING_WRITE, "M_Main.LoadBank", _
                              "The " & Columns(c).ColumnName & " value of record " & r & " has " & Len(Values(r, c)) & _
                              " characters; a cell holds " & EXCEL_CELL_CHARS & ". Nothing was loaded."
                End If
                If Len(Values(r, c)) = 0 Then
                    NewBlock(r, c) = Empty
                ElseIf Columns(c).Kind = "number" Then
                    NewBlock(r, c) = Val(Values(r, c))    'Val reads the "." of the file in every locale
                Else
                    NewBlock(r, c) = Values(r, c)
                End If
            Next c
        Next r

'------------------------------------------------------------------------------
' WRITE
'------------------------------------------------------------------------------
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

        OldLast = UsedLastRow(Sheet)
        If OldLast >= FirstRow Then
            OldBlock = TableRange(Sheet, FirstRow, 1, OldLast, Width).Formula
        End If
        LastRow = FirstRow + Result.RecordCount - 1
        If OldLast > LastRow Then
            LastRow = OldLast
        End If
        Cleared = True
        TableRange(Sheet, FirstRow, 1, LastRow, Width).ClearContents
    'The template formats the first rows of the sheet; a longer file is written
    'with the format of the first row in every column (#718), text as text.
        For c = 1 To Width
            If Columns(c).Kind <> "number" Then
                TableRange(Sheet, FirstRow, c, FirstRow + Result.RecordCount - 1, c).NumberFormat = "@"
            Else
                TableRange(Sheet, FirstRow, c, FirstRow + Result.RecordCount - 1, c).NumberFormat = _
                    TableCell(Sheet, FirstRow, c).NumberFormat
            End If
        Next c
        TableRange(Sheet, FirstRow, 1, FirstRow + Result.RecordCount - 1, Width).Value2 = NewBlock
        If gTestFault = "bank-write" Then
            Err.Raise ERR_INJECTED_FAULT, "M_Main.LoadBank", "Injected bank write failure."
        End If

CleanUp:
        On Error GoTo 0
        If ErrorNumber <> 0 And Cleared Then
            On Error Resume Next
            TableRange(Sheet, FirstRow, 1, LastRow, Width).ClearContents
            If OldLast >= FirstRow Then
                TableRange(Sheet, FirstRow, 1, OldLast, Width).Formula = OldBlock
            End If
            If Err.Number <> 0 Then
                ErrorDescription = ErrorDescription & " " & Sheet.Name & " could not be put back: " & Err.Description
            End If
            On Error GoTo 0
        End If
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
            Err.Raise ERR_CLEANUP_FAILED, "M_Main.LoadBank", _
                      "The records were loaded, but Excel settings were not restored: " & CleanupDetails
        End If
        Exit Sub

Failed:
        ErrorNumber = Err.Number
        ErrorSource = Err.Source
        ErrorDescription = Err.Description
        Resume CleanUp

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
        ReapplyOutputFilters

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
'                  CLOSING PRICES OF THE FX TABLE (#721)
'
'------------------------------------------------------------------------------
'

Public Sub DownloadClosingPrices()
'
'==============================================================================
'                            DownloadClosingPrices
'------------------------------------------------------------------------------
' PURPOSE
'   Look up the rate of each currency of the FX table on the reporting date
'   from the source chosen in FX_Source and write it as a value in the rate
'   cell of the currency, after a confirmation with Cancel as the default,
'   and show the source and the date of the rates beside the button. The
'   sources are the closing price of the pair with STOCKHISTORY (#721) and
'   the euro reference rate of the ECB (#745). The pair is built from the
'   reporting currency and the code of the row; no formula, no data request
'   and no connection is left in the workbook.
'
' USAGE
'   Download prices on the Params sheet, or the Macros dialog. Automation
'   calls DownloadClosingPrices_Silent.
'
' METHOD
'   Closing price: the macro writes the formulas into a hidden scratch
'   column and returns, as Excel hands the requests of STOCKHISTORY back
'   only when no macro runs. PollClosingPrices, scheduled each second,
'   writes the rates when the prices have come, and the message is shown
'   then. ECB reference rate: LookUpEcbRates asks the ECB data portal over
'   HTTPS and writes the rates in the same call. The status bar shows the
'   wait.
'
' ERROR POLICY
'   Reports an error in a message, or in the Immediate window in silent
'   mode; never raises. STOCKHISTORY needs Microsoft 365 and a connection,
'   the ECB a connection to its data portal: without them no rate is
'   changed. In silent mode, where no confirmation can be given, nothing is
'   looked up.
'
' UPDATED
'   2026-11-22
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Pairs       As Long      'Currencies with a pair
    Dim Problem     As String    'Why nothing was looked up
    Dim Source      As String    'FX_SOURCE_CLOSING or FX_SOURCE_ECB
    Dim Priced      As Long      'Rates an ECB lookup wrote
    Dim Without     As String    'Currencies the ECB gave no rate for
    Dim Fixing      As String    'The day of the ECB rates, when not the reporting date

'------------------------------------------------------------------------------
' CONFIRM AND START THE LOOKUP
'------------------------------------------------------------------------------
        On Error GoTo Failed
        If mRunning Then
            Notify "An SA-CCR operation is already running.", vbExclamation
            Exit Sub
        End If
        If gSilent Then
            Notify "Nothing was looked up." & vbLf & vbLf & "A lookup is confirmed in a dialog, which silent mode " & _
                   "does not show. Automation calls DownloadClosingPrices_Silent.", vbExclamation
            Exit Sub
        End If
        Pairs = ClosingPricePairs()
        If Pairs = 0 Then
            Notify "No currency of the FX table has a pair: nothing to look up.", vbExclamation
            Exit Sub
        End If
        Source = PriceSource(vbNullString, Problem)
        If Len(Problem) > 0 Then
            Notify "Nothing was looked up." & vbLf & vbLf & Problem, vbExclamation
            Exit Sub
        End If
        If MsgBox(PriceQuestion(Pairs, Source), vbOKCancel + vbQuestion + vbDefaultButton2, MSG_TITLE) <> vbOK Then
            Exit Sub
        End If

    'The ECB: one request, answered within this call
        If Source = FX_SOURCE_ECB Then
            Application.Cursor = xlWait
            Application.StatusBar = "Downloading the ECB reference rates of " & ClosingPriceDate() & "..."
            LookUpEcbRates Priced, Without, Problem, Fixing
            Application.Cursor = xlDefault
            Application.StatusBar = False
            RibbonStateChanged
            ReportPrices Priced, Without, Problem, Source, Fixing
            Exit Sub
        End If

    'STOCKHISTORY: the formulas now, the prices from the timer
        If Not StartPriceLookup(Problem) Then
            ReportPrices 0, vbNullString, Problem, Source
            Exit Sub
        End If
        Application.StatusBar = "Downloading the prices of " & ClosingPriceDate() & "..."
        SchedulePricePoll
        Exit Sub

Failed:
        If Source = FX_SOURCE_ECB Then
            Application.Cursor = xlDefault
            Application.StatusBar = False
        End If
        Notify "Download prices: " & Err.Description, vbCritical

End Sub


Public Sub PollClosingPrices()
'
'==============================================================================
'                              PollClosingPrices
'------------------------------------------------------------------------------
' PURPOSE
'   Look at the scratch cells of the lookup that DownloadClosingPrices
'   started, one second after the last look (#721): while a price is still
'   awaited and the time is not up, schedule the next look and let Excel
'   work; otherwise write the rates, show the result and end the lookup.
'   Application.OnTime calls it, so that Excel is idle between the looks
'   and hands the requests of STOCKHISTORY back.
'
' ERROR POLICY
'   Reports an error in a message; the lookup ends, the rates are put back.
'
' UPDATED
'   2026-11-21
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ResultLine  As String    'What the lookup returned
    Dim Priced      As Long      'Rates written
    Dim Without     As String    'Currencies with no price
    Dim Problem     As String    'Why no rate was changed
    Dim ErrText     As String    'The description of an error

'------------------------------------------------------------------------------
' LOOK
'------------------------------------------------------------------------------
        mPollAt = 0
        If Not mPriceActive Then
            Exit Sub
        End If
        On Error GoTo Failed
        If PricesPending() And Now < mPriceDeadline Then
            Application.StatusBar = "Downloading the prices of " & ClosingPriceDate() & "... " & _
                                    Format$((Now - mPriceStarted) * 86400#, "0") & " s"
            SchedulePricePoll
            Exit Sub
        End If
        ResultLine = FinishPriceLookup(Priced, Without, Problem)
        RibbonStateChanged
        ReportPrices Priced, Without, Problem, FX_SOURCE_CLOSING
        Exit Sub

Failed:
        ErrText = Err.Description
        AbandonPriceLookup
        Notify "Download closing prices: " & ErrText, vbCritical

End Sub


Public Sub CancelPriceLookup()
'
'==============================================================================
'                              CancelPriceLookup
'------------------------------------------------------------------------------
' PURPOSE
'   End a lookup of closing prices that is waiting for its data, without
'   changing a rate, when the workbook is saved or closed (#721), so that
'   no price formula is saved and no scheduled look reopens the workbook.
'   ThisWorkbook calls it. Does nothing when no lookup is waiting.
'
' METHOD
'   Emptying the scratch column with events on would raise
'   Workbook_SheetChange, whose Ribbon refresh is scheduled with
'   Application.OnTime and could reopen the workbook once it is closed. So
'   events are off while the lookup ends, and put back as they were. The
'   Ribbon was disabled while the lookup ran and the workbook can stay open
'   (a save, or a close whose save prompt is cancelled), so it is
'   invalidated afterwards, directly and not through a scheduled callback.
'
' ERROR POLICY
'   Never raises.
'
' UPDATED
'   2026-11-21
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim EventsOn    As Boolean   'Events were enabled

'------------------------------------------------------------------------------
' END THE LOOKUP
'------------------------------------------------------------------------------
        If mPriceActive Then
            On Error Resume Next
            EventsOn = Application.EnableEvents
            Application.EnableEvents = False
            AbandonPriceLookup
            Application.EnableEvents = EventsOn
            M_Ribbon.RefreshRibbon
        End If

End Sub


Private Sub ReportPrices( _
    ByVal Priced As Long, _
    ByVal Without As String, _
    ByVal Problem As String, _
    ByVal Source As String, _
    Optional ByVal Fixing As String = vbNullString)
'
'==============================================================================
'                               ReportPrices
'------------------------------------------------------------------------------
' PURPOSE
'   Tell the user the result of a lookup of rates (#721, #745): why no rate
'   was changed, or how many rates were set from which source and for which
'   currencies there was no rate. Fixing is the day of the ECB rates when it
'   is not the reporting date, a day without a fixing.
'
' UPDATED
'   2026-11-22
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim What        As String    'The kind of rate, as the message names it

'------------------------------------------------------------------------------
' TELL
'------------------------------------------------------------------------------
        If Len(Problem) > 0 Then
            Notify "No rate was changed." & vbLf & vbLf & Problem, vbExclamation
            Exit Sub
        End If
        What = IIf(Source = FX_SOURCE_ECB, "ECB reference rate", "closing price")
        Notify Priced & " rate" & IIf(Priced = 1, "", "s") & " set to the " & What & " of " & ClosingPriceDate() & _
               IIf(Len(Fixing) > 0, " (the fixing of " & Fixing & ", the last one on or before that day)", "") & _
               IIf(Len(Without) > 0, "; no " & What & " for " & Replace(Without, ",", ", ") & _
                   ", whose rate stays as it was.", "."), _
               IIf(Len(Without) > 0, vbExclamation, vbInformation)

End Sub


Private Function PriceQuestion( _
    ByVal Pairs As Long, _
    ByVal Source As String) _
    As String
'
'==============================================================================
'                              PriceQuestion
'------------------------------------------------------------------------------
' PURPOSE
'   The confirmation of Download prices, which names the source (#745) and
'   says what the lookup needs and sends.
'
' UPDATED
'   2026-11-22
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Currencies  As String    'How many currencies, in words

'------------------------------------------------------------------------------
' BUILD
'------------------------------------------------------------------------------
        Currencies = Pairs & " currenc" & IIf(Pairs = 1, "y", "ies")
        If Source = FX_SOURCE_ECB Then
            PriceQuestion = "Replace the rates of " & Currencies & " with the ECB reference rates of " & _
                            ClosingPriceDate() & "?" & vbLf & vbLf & _
                            "The rates now in the table are overwritten and Undo cannot bring them back. The rates are " & _
                            "asked of the ECB data portal over HTTPS, with the currency codes and the dates only; it needs " & _
                            "a connection and takes up to " & PRICE_WAIT_SECONDS & " seconds."
        Else
            PriceQuestion = "Replace the rates of " & Currencies & " with the closing prices of " & _
                            ClosingPriceDate() & "?" & vbLf & vbLf & _
                            "The rates now in the table are overwritten and Undo cannot bring them back. It needs " & _
                            "Microsoft 365 and a connection, and takes up to " & PRICE_WAIT_SECONDS & " seconds."
        End If

End Function


Private Function PriceSource( _
    ByVal Given As String, _
    ByRef Problem As String) _
    As String
'
'==============================================================================
'                               PriceSource
'------------------------------------------------------------------------------
' PURPOSE
'   The source of the rates a lookup uses (#745): Given when it is not
'   blank, else the cell FX_Source on Params; blank is the default,
'   FX_SOURCE_CLOSING. Case and surrounding spaces are ignored.
'
' RETURNS
'   FX_SOURCE_CLOSING or FX_SOURCE_ECB. "" with Problem set when the source
'   is neither, or the name FX_Source is missing.
'
' UPDATED
'   2026-11-22
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Cell        As Range     'FX_Source
    Dim Text        As String    'The source as given

'------------------------------------------------------------------------------
' READ
'------------------------------------------------------------------------------
        Problem = vbNullString
        Text = Trim$(Given)
        If Len(Text) = 0 Then
            Set Cell = NamedCell(NM_FX_SOURCE)
            If Cell Is Nothing Then
                Problem = "the workbook name '" & NM_FX_SOURCE & "' is missing."
                Exit Function
            End If
            Text = Trim$(SafeStr(Cell.Cells(1, 1).Value2))
        End If
        Select Case LCase$(Text)
            Case vbNullString, LCase$(FX_SOURCE_CLOSING)
                PriceSource = FX_SOURCE_CLOSING
            Case LCase$(FX_SOURCE_ECB)
                PriceSource = FX_SOURCE_ECB
            Case Else
                Problem = "the source of the rates is '" & Text & "': choose " & FX_SOURCE_CLOSING & " or " & _
                          FX_SOURCE_ECB & " beside Download prices on " & SH_PARAMS & "."
        End Select

End Function


Public Function DownloadClosingPrices_Silent( _
    Optional ByVal Source As String = vbNullString) _
    As String
'
'==============================================================================
'                         DownloadClosingPrices_Silent
'------------------------------------------------------------------------------
' PURPOSE
'   Look up the rates of the FX table and write them (#721), for automation
'   and tests, without a confirmation or a message, from Source, or from
'   the source in FX_Source when Source is blank (#745).
'
' RETURNS
'   "RESULT=PRICED; rates=N; without_price=M" (and "; codes=..." with the
'   currencies that got no rate, whose rate stays; from the ECB, then
'   "; source=ECB" and "; fixing=yyyy-mm-dd" when the fixing is of an
'   earlier day), or "RESULT=REFUSED; reason=..." when no rate was
'   changed, for example without STOCKHISTORY, without a connection, or
'   when no rate came for any currency.
'
' ERROR POLICY
'   Raises ERR_RUN_ACTIVE while an operation runs. gTestFault =
'   "prices-unavailable" makes the lookup refuse as an Excel that
'   cannot take the formula does; "prices-name-error" writes =#NAME? as
'   an Excel that takes it but has no STOCKHISTORY would evaluate it;
'   "prices-fixed" writes the price 2 for every pair, and no price (#N/A)
'   for the first one, and "prices-none" no price (#N/A) for any pair, in
'   place of the formulas of STOCKHISTORY. The faults of the ECB are in
'   EcbAnswer and LookUpEcbRates.
'
' UPDATED
'   2026-11-22
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Priced      As Long      'Rates written
    Dim Without     As String    'Currencies with no price
    Dim Problem     As String    'Why no rate was changed
    Dim Fixing      As String    'The day of the ECB rates, when not the reporting date
    Dim Chosen      As String    'The source used

'------------------------------------------------------------------------------
' LOOK UP
'------------------------------------------------------------------------------
        If mRunning Then
            Err.Raise ERR_RUN_ACTIVE, "M_Main.DownloadClosingPrices_Silent", "An SA-CCR operation is already running."
        End If
        Chosen = PriceSource(Source, Problem)
        If Len(Problem) > 0 Then
            DownloadClosingPrices_Silent = "RESULT=REFUSED; reason=" & Problem
        ElseIf Chosen = FX_SOURCE_ECB Then
            DownloadClosingPrices_Silent = LookUpEcbRates(Priced, Without, Problem, Fixing)
        Else
            DownloadClosingPrices_Silent = LookUpClosingPrices(Priced, Without, Problem)
        End If

End Function


Private Function ClosingPair( _
    ByVal ParamsSheet As Worksheet, _
    ByVal Row As Long, _
    ByVal KeyColumn As Long) _
    As String
'
'==============================================================================
'                               ClosingPair
'------------------------------------------------------------------------------
' PURPOSE
'   The currency pair of the FX-table row, the reporting currency over the
'   currency of the row, such as EUR/USD, in capitals (#721). Empty for the
'   row of the reporting currency. It is built from the code, so that a
'   currency appended below the table has one too.
'
' UPDATED
'   2026-11-21
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Code        As String    'The currency of the row
    Dim Reporting   As String    'The reporting currency

'------------------------------------------------------------------------------
' BUILD
'------------------------------------------------------------------------------
        Code = UCase$(Trim$(SafeStr(TableCell(ParamsSheet, Row, KeyColumn).Value2)))
        Reporting = UCase$(Trim$(SafeStr(NamedCell(PRM_REPCCY).Value2)))
        If Len(Code) = 0 Or Len(Reporting) = 0 Or Code = Reporting Then
            Exit Function
        End If
        ClosingPair = Reporting & "/" & Code

End Function


Private Function IsReportingRow( _
    ByVal ParamsSheet As Worksheet, _
    ByVal Row As Long, _
    ByVal KeyColumn As Long) _
    As Boolean
'
'==============================================================================
'                               IsReportingRow
'------------------------------------------------------------------------------
' PURPOSE
'   Whether the FX-table row is that of the reporting currency, whose rate
'   Download prices sets to 1 with the others (#745). Case and surrounding
'   spaces are ignored, as in ClosingPair.
'
' UPDATED
'   2026-11-22
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Code        As String    'The currency of the row
    Dim Reporting   As String    'The reporting currency

'------------------------------------------------------------------------------
' COMPARE
'------------------------------------------------------------------------------
        Code = UCase$(Trim$(SafeStr(TableCell(ParamsSheet, Row, KeyColumn).Value2)))
        Reporting = UCase$(Trim$(SafeStr(NamedCell(PRM_REPCCY).Value2)))
        IsReportingRow = (Len(Code) > 0 And Code = Reporting)

End Function


Private Function ClosingPriceDate() As String
'
'==============================================================================
'                             ClosingPriceDate
'------------------------------------------------------------------------------
' PURPOSE
'   The reporting date, the day of the closing prices, as yyyy-mm-dd (#721).
'
' UPDATED
'   2026-11-21
'==============================================================================
'

'------------------------------------------------------------------------------
' READ
'------------------------------------------------------------------------------
        On Error GoTo NoDate
        ClosingPriceDate = Format$(CDate(NamedCell(PRM_ASOF).Value2), "yyyy-mm-dd")
        Exit Function

NoDate:
        ClosingPriceDate = "the reporting date"

End Function


Private Function ClosingPricePairs() As Long
'
'==============================================================================
'                             ClosingPricePairs
'------------------------------------------------------------------------------
' PURPOSE
'   How many rows of the FX table have a currency pair (ClosingPair), the
'   rows whose closing price can be looked up (#721). The table runs down
'   to the first blank key.
'
' UPDATED
'   2026-11-21
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ParamsSheet   As Worksheet    'Params
    Dim HeaderRow     As Long         'Header row of the FX table
    Dim KeyColumn     As Long         'Its key column
    Dim r             As Long         'Row

'------------------------------------------------------------------------------
' COUNT
'------------------------------------------------------------------------------
        Set ParamsSheet = GetSheet(SH_PARAMS)
        HeaderRow = NamedRow(NM_FX_TABLE)
        If HeaderRow = 0 Then
            Exit Function
        End If
        KeyColumn = NamedTableCol(NM_FX_TABLE)
        r = HeaderRow + 1
        Do While Not IsBlankCell(TableCell(ParamsSheet, r, KeyColumn).Value)
            If Len(ClosingPair(ParamsSheet, r, KeyColumn)) > 0 Then
                ClosingPricePairs = ClosingPricePairs + 1
            End If
            r = r + 1
        Loop

End Function


Private Function LookUpClosingPrices( _
    ByRef Priced As Long, _
    ByRef Without As String, _
    ByRef Problem As String) _
    As String
'
'==============================================================================
'                            LookUpClosingPrices
'------------------------------------------------------------------------------
' PURPOSE
'   Look up the closing prices and write the rates in one call, for
'   automation and tests (#721): StartPriceLookup, a wait that lets
'   DoEvents run for PRICE_WAIT_SECONDS at most, and FinishPriceLookup.
'   Excel hands the requests of STOCKHISTORY back only when no macro runs,
'   so with real data this wait ends without prices: the button does not
'   use it, it starts the lookup and PollClosingPrices ends it.
'
' ERROR POLICY
'   A refusal is returned in the result line, not raised. An error puts the
'   rates back as they were and is raised again.
'
' UPDATED
'   2026-11-21
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Deadline    As Double    'When to stop waiting
    Dim ErrNumber   As Long      'The error to raise again
    Dim ErrText     As String    'Its description

'------------------------------------------------------------------------------
' START, WAIT AND FINISH
'------------------------------------------------------------------------------
        If Not StartPriceLookup(Problem) Then
            LookUpClosingPrices = "RESULT=REFUSED; reason=" & Problem
            Exit Function
        End If
        On Error GoTo Failed
        Deadline = Timer + PRICE_WAIT_SECONDS
        Do While PricesPending() And Timer < Deadline
            DoEvents
        Loop
        LookUpClosingPrices = FinishPriceLookup(Priced, Without, Problem)
        Exit Function

Failed:
        ErrNumber = Err.Number
        ErrText = Err.Description
        AbandonPriceLookup
        Err.Raise ErrNumber, "M_Main.LookUpClosingPrices", ErrText

End Function


Public Sub ReportingCurrencyChanged( _
    ByVal Sh As Object, _
    ByVal Target As Range)
'
'==============================================================================
'                          ReportingCurrencyChanged
'------------------------------------------------------------------------------
' PURPOSE
'   When the user changes the reporting currency, offer to rebase the FX
'   table on it (#725): the rates are units of each currency per 1 unit of
'   the reporting currency, so after a change they are still those of the
'   old one, and a run refuses a table whose reporting currency is not 1.
'   ThisWorkbook.Workbook_SheetChange calls it for every edit.
'
' METHOD
'   Nothing is asked when the edit is not the reporting currency of
'   Params, when an operation runs, in silent mode, when the new currency
'   is not in the FX table or its rate is not a number above zero, or
'   when its rate is already 1. Otherwise a question, with No as the
'   default, names the currency and its rate; Yes runs RebaseFxTable.
'
' ERROR POLICY
'   Never raises: an edit must never fail because of it.
'
' UPDATED
'   2026-11-22
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Cell        As Range     'ReportingCcy
    Dim Code        As String    'The new reporting currency
    Dim Rate        As Double    'Its rate in the FX table
    Dim Problem     As String    'Why nothing was rebased
    Dim ResultLine  As String    'What the rebase returned

'------------------------------------------------------------------------------
' ASK
'------------------------------------------------------------------------------
        On Error GoTo Done
        If mRunning Or gSilent Then
            Exit Sub
        End If
        Set Cell = NamedCell(PRM_REPCCY)
        If Cell Is Nothing Then
            Exit Sub
        End If
        If Not Sh Is Cell.Worksheet Then
            Exit Sub
        End If
        If Application.Intersect(Target, Cell) Is Nothing Then
            Exit Sub
        End If
        Code = UCase$(Trim$(SafeStr(Cell.Cells(1, 1).Value2)))
        If Not Code Like "[A-Z][A-Z][A-Z]" Then
            Exit Sub
        End If
        Rate = FxTableRate(Code)
        If Rate <= 0# Or Abs(Rate - 1#) <= 0.000000000001 Then
            Exit Sub
        End If
        If MsgBox("The reporting currency is now " & Code & ", whose rate in the FX table is " & CStr(Rate) & "." & _
                  vbLf & vbLf & "Rebase the FX table on " & Code & "? Every rate is divided by " & CStr(Rate) & _
                  ", so that " & Code & " is 1 and every other rate is in units per 1 " & Code & _
                  ". Undo cannot bring the rates back." & vbLf & vbLf & _
                  "Without it the next run stops, until the rates are rebased or downloaded again.", _
                  vbYesNo + vbQuestion + vbDefaultButton2, MSG_TITLE) <> vbYes Then
            Exit Sub
        End If
        ResultLine = RebaseFxTable(Code, Problem)
        If Len(Problem) > 0 Then
            Notify "The FX table was not rebased." & vbLf & vbLf & Problem, vbExclamation
        Else
            Notify RebaseMessage(ResultLine, Code), vbInformation
        End If

Done:

End Sub


Public Function RebaseFxTable_Silent( _
    Optional ByVal Code As String = vbNullString) _
    As String
'
'==============================================================================
'                            RebaseFxTable_Silent
'------------------------------------------------------------------------------
' PURPOSE
'   Rebase the FX table on Code, or on the reporting currency of Params
'   when Code is blank, without a question or a message (#725), for
'   automation and tests. See RebaseFxTable.
'
' RETURNS
'   "RESULT=REBASED; currency=XXX; divisor=D; rates=N[; not_rebased=...]",
'   "RESULT=UNCHANGED; currency=XXX" when its rate is already 1, or
'   "RESULT=REFUSED; reason=...".
'
' ERROR POLICY
'   Raises ERR_RUN_ACTIVE while an operation runs.
'
' UPDATED
'   2026-11-22
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Problem     As String    'Why nothing was rebased

'------------------------------------------------------------------------------
' REBASE
'------------------------------------------------------------------------------
        If mRunning Then
            Err.Raise ERR_RUN_ACTIVE, "M_Main.RebaseFxTable_Silent", "An SA-CCR operation is already running."
        End If
        If Len(Trim$(Code)) = 0 Then
            Code = SafeStr(NamedCell(PRM_REPCCY).Value2)
        End If
        RebaseFxTable_Silent = RebaseFxTable(UCase$(Trim$(Code)), Problem)

End Function


Private Function RebaseFxTable( _
    ByVal Code As String, _
    ByRef Problem As String) _
    As String
'
'==============================================================================
'                                RebaseFxTable
'------------------------------------------------------------------------------
' PURPOSE
'   Express every rate of the FX table in units per 1 unit of Code (#725):
'   each rate is divided by the rate of Code, which becomes exactly 1. The
'   rates are relative, so the conversion of every amount is unchanged.
'   The date and the source of the rates stay: they are the same quotes.
'   So does the record of the download (#745): a table as downloaded is
'   still as downloaded once rebased, and its fingerprint is taken again;
'   a table edited after the download stays edited.
'
' METHOD
'   Work out every new rate, then write them with events off, all or none.
'   A rate that is a formula, or not a number above zero, is left as it is
'   and listed in not_rebased; when the rate of Code itself is a formula,
'   the rebase is refused, as it would replace it by 1. The divisor in the
'   result is written with Str$, with a decimal point whatever the locale.
'
' RETURNS
'   See RebaseFxTable_Silent. Problem is set when the result is REFUSED.
'
' ERROR POLICY
'   A refusal is returned, not raised. An error puts the rates back and is
'   raised again.
'
' UPDATED
'   2026-11-22
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ParamsSheet   As Worksheet    'Params
    Dim HeaderRow     As Long         'Header row of the FX table
    Dim KeyColumn     As Long         'Its key column
    Dim LastRow       As Long         'Its last row
    Dim r             As Long         'Row
    Dim RateCell      As Range        'A rate
    Dim Divisor       As Double       'The rate of Code
    Dim Kept()        As Variant      'Each rate before
    Dim NewRates()    As Variant      'Each rate after; Empty when it stays
    Dim Rebased       As Long         'Rates divided
    Dim Skipped       As String       'Currencies whose rate stays
    Dim EventsOn      As Boolean      'Events were enabled
    Dim Started       As Boolean      'The writing began
    Dim AsDownloaded  As Boolean      'The table was as Download prices wrote it
    Dim ErrNumber     As Long         'The error to raise again
    Dim ErrText       As String       'Its description

'------------------------------------------------------------------------------
' WORK OUT
'------------------------------------------------------------------------------
        Problem = vbNullString
        If Not Code Like "[A-Z][A-Z][A-Z]" Then
            Problem = "the currency '" & Code & "' is not a three-letter code."
            RebaseFxTable = "RESULT=REFUSED; reason=" & Problem
            Exit Function
        End If
        Divisor = FxTableRate(Code)
        If Divisor = 0# Then
            Problem = Code & " is not in the FX table, or its rate is not a number above zero."
            RebaseFxTable = "RESULT=REFUSED; reason=" & Problem
            Exit Function
        End If
        If Abs(Divisor - 1#) <= 0.000000000001 Then
            RebaseFxTable = "RESULT=UNCHANGED; currency=" & Code
            Exit Function
        End If
        Set ParamsSheet = GetSheet(SH_PARAMS)
        HeaderRow = NamedRow(NM_FX_TABLE)
        KeyColumn = NamedTableCol(NM_FX_TABLE)
        LastRow = HeaderRow
        Do While Not IsBlankCell(TableCell(ParamsSheet, LastRow + 1, KeyColumn).Value)
            LastRow = LastRow + 1
        Loop
        ReDim Kept(HeaderRow + 1 To LastRow)
        ReDim NewRates(HeaderRow + 1 To LastRow)
        For r = HeaderRow + 1 To LastRow
            Set RateCell = TableCell(ParamsSheet, r, KeyColumn + FX_COL_RATE)
            Kept(r) = RateCell.Formula
            If UCase$(Trim$(SafeStr(TableCell(ParamsSheet, r, KeyColumn).Value2))) = Code Then
                If RateCell.HasFormula Then
                    Problem = "the rate of " & Code & " is a formula, which a rebase would replace by 1: write " & _
                              "its rate as a number first."
                    RebaseFxTable = "RESULT=REFUSED; reason=" & Problem
                    Exit Function
                End If
                NewRates(r) = 1#
            ElseIf RateCell.HasFormula Or VarType(RateCell.Value2) <> vbDouble Then
                Skipped = Skipped & IIf(Len(Skipped) > 0, ",", "") & UCase$(Trim$(SafeStr(TableCell(ParamsSheet, r, KeyColumn).Value2)))
            ElseIf RateCell.Value2 <= 0# Then
                Skipped = Skipped & IIf(Len(Skipped) > 0, ",", "") & UCase$(Trim$(SafeStr(TableCell(ParamsSheet, r, KeyColumn).Value2)))
            Else
                NewRates(r) = RateCell.Value2 / Divisor
            End If
        Next r

'------------------------------------------------------------------------------
' WRITE
'------------------------------------------------------------------------------
        If Not NamedCell(NM_FX_PRICE_FP) Is Nothing Then
            AsDownloaded = Len(SafeStr(NamedCell(NM_FX_PRICE_FP).Value2)) > 0 And _
                           SafeStr(NamedCell(NM_FX_PRICE_FP).Value2) = CORE_Engine.FxTableFingerprint()
        End If
        On Error GoTo Failed
        mRunning = True
        EventsOn = Application.EnableEvents
        Application.EnableEvents = False
        Started = True
        For r = HeaderRow + 1 To LastRow
            If Not IsEmpty(NewRates(r)) Then
                TableCell(ParamsSheet, r, KeyColumn + FX_COL_RATE).Value2 = NewRates(r)
                Rebased = Rebased + 1
                If gTestFault = "rebase-write" Then
                    Err.Raise ERR_INJECTED_FAULT, "M_Main.RebaseFxTable", "Injected failure while the FX table is rebased."
                End If
            End If
        Next r
        If AsDownloaded Then
            PutFxPrint CORE_Engine.FxTableFingerprint()
        End If
        Application.EnableEvents = EventsOn
        mRunning = False
        RebaseFxTable = "RESULT=REBASED; currency=" & Code & "; divisor=" & Trim$(Str$(Divisor)) & "; rates=" & Rebased & _
                        IIf(Len(Skipped) > 0, "; not_rebased=" & Skipped, "")
        Exit Function

Failed:
        ErrNumber = Err.Number
        ErrText = Err.Description
        On Error Resume Next
        If Started Then
            For r = HeaderRow + 1 To LastRow
                TableCell(ParamsSheet, r, KeyColumn + FX_COL_RATE).Formula = Kept(r)
            Next r
            Application.EnableEvents = EventsOn
        End If
        mRunning = False
        On Error GoTo 0
        Err.Raise ErrNumber, "M_Main.RebaseFxTable", ErrText

End Function


Private Sub PutFxPrint( _
    ByVal Fingerprint As Variant)
'
'==============================================================================
'                                  PutFxPrint
'------------------------------------------------------------------------------
' PURPOSE
'   Write a fingerprint of the FX table into the hidden cell
'   FX_PriceFingerprint (#745) as text, with an apostrophe, so that Excel
'   never reads one such as "00012E45" as a number; an empty one empties
'   the cell.
'
' ERROR POLICY
'   Raises when the name is missing; the callers check it first.
'
' UPDATED
'   2026-11-22
'==============================================================================
'
        If Len(SafeStr(Fingerprint)) = 0 Then
            NamedCell(NM_FX_PRICE_FP).Value2 = vbNullString
        Else
            NamedCell(NM_FX_PRICE_FP).Value2 = "'" & SafeStr(Fingerprint)
        End If

End Sub


Private Function FxTableRate( _
    ByVal Code As String) _
    As Double
'
'==============================================================================
'                                 FxTableRate
'------------------------------------------------------------------------------
' PURPOSE
'   The rate of Code in the FX table (#725): 0 when the table has no such
'   currency, or its rate is not a number above zero.
'
' UPDATED
'   2026-11-22
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ParamsSheet   As Worksheet    'Params
    Dim HeaderRow     As Long         'Header row of the FX table
    Dim KeyColumn     As Long         'Its key column
    Dim r             As Long         'Row
    Dim Found         As Variant      'A rate

'------------------------------------------------------------------------------
' FIND
'------------------------------------------------------------------------------
        Set ParamsSheet = GetSheet(SH_PARAMS)
        HeaderRow = NamedRow(NM_FX_TABLE)
        If HeaderRow = 0 Then
            Exit Function
        End If
        KeyColumn = NamedTableCol(NM_FX_TABLE)
        r = HeaderRow + 1
        Do While Not IsBlankCell(TableCell(ParamsSheet, r, KeyColumn).Value)
            If UCase$(Trim$(SafeStr(TableCell(ParamsSheet, r, KeyColumn).Value2))) = Code Then
                Found = TableCell(ParamsSheet, r, KeyColumn + FX_COL_RATE).Value2
                If VarType(Found) = vbDouble Then
                    If Found > 0# Then
                        FxTableRate = Found
                    End If
                End If
                Exit Function
            End If
            r = r + 1
        Loop

End Function


Private Function RebaseMessage( _
    ByVal ResultLine As String, _
    ByVal Code As String) _
    As String
'
'==============================================================================
'                                RebaseMessage
'------------------------------------------------------------------------------
' PURPOSE
'   The message after a rebase of the FX table (#725), with the rates that
'   stayed as they were.
'
' UPDATED
'   2026-11-22
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim At          As Long      'Where not_rebased starts

'------------------------------------------------------------------------------
' BUILD
'------------------------------------------------------------------------------
        RebaseMessage = "The FX table is now in units per 1 " & Code & "; " & Code & " is 1."
        At = InStr(1, ResultLine, "; not_rebased=", vbBinaryCompare)
        If At > 0 Then
            RebaseMessage = RebaseMessage & vbLf & vbLf & "Not rebased, as their rate is a formula or not a number " & _
                            "above zero: " & Replace(Mid$(ResultLine, At + Len("; not_rebased=")), ",", ", ") & _
                            ". Check them before the next run."
        End If

End Function


Private Function LookUpEcbRates( _
    ByRef Priced As Long, _
    ByRef Without As String, _
    ByRef Problem As String, _
    ByRef Fixing As String) _
    As String
'
'==============================================================================
'                               LookUpEcbRates
'------------------------------------------------------------------------------
' PURPOSE
'   Look up the ECB euro foreign exchange reference rates on the reporting
'   date and write them as the rates of the FX table, all or none (#745).
'   For each currency the latest fixing on or before the reporting date is
'   used, ECB_LOOKBACK_DAYS back at most. The ECB quotes every currency
'   against the euro: with another reporting currency the rate is crossed
'   through the euro, the currency per euro over the reporting currency per
'   euro, and the euro is 1 over the latter. A currency the ECB gave no rate
'   for keeps its rate and is listed in Without. The row of the reporting
'   currency takes 1, the rate a run requires of it.
'
' METHOD
'   1. Find the rows with a pair and note their rates, and the date and
'      source of the rates, to put them back.
'   2. Ask the ECB (EcbAnswer) for the currencies of the table that are
'      three-letter codes, the euro aside, and the reporting currency
'      unless it is the euro; read the answer (EcbLatestRates).
'   3. Work out every rate; refuse when the reporting date, the reporting
'      currency or a code was changed while the answer was awaited, which
'      DoEvents allows; then write the rates, 1 for the reporting currency,
'      the reporting date into
'      NM_FX_PRICE_DATE and FX_SOURCE_ECB into NM_FX_PRICE_SRC.
'   mRunning is set while the answer is awaited, so that no other operation
'   starts.
'
' RETURNS
'   "RESULT=PRICED; rates=N; without_price=M[; codes=...]; source=ECB
'   [; fixing=yyyy-mm-dd]" or "RESULT=REFUSED; reason=...". Fixing is the
'   day of the fixing used when it is not the reporting date.
'
' ERROR POLICY
'   A refusal is returned, not raised, and changes nothing: no connection,
'   an HTTP error, a time-out, an answer that is not the CSV expected, no
'   rate for the reporting currency, none for any currency, or an input
'   changed while the answer was awaited. An error
'   puts the rates, the date and the source back and is raised again.
'   gTestFault = "ecb-write" raises ERR_INJECTED_FAULT after the first rate
'   is written.
'
' UPDATED
'   2026-11-22
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ParamsSheet   As Worksheet    'Params
    Dim HeaderRow     As Long         'Header row of the FX table
    Dim KeyColumn     As Long         'Its key column
    Dim r             As Long         'Row
    Dim i             As Long         'Row with a pair
    Dim k             As Long         'Currency asked for
    Dim Count         As Long         'Rows with a pair
    Dim PairRows()    As Long         'The rows with a pair
    Dim Codes()       As String       'Their currencies
    Dim Kept()        As Variant      'The rate each row held
    Dim KeptDate      As Variant      'The date of the rates before
    Dim KeptSource    As Variant      'Their source before
    Dim KeptPrint     As Variant      'The fingerprint of the table as downloaded before
    Dim Reporting     As String       'The reporting currency
    Dim AsOfValue     As Variant      'The reporting date as the cell holds it
    Dim AsOf          As Date         'The reporting date
    Dim Asked()       As String       'The currencies asked of the ECB
    Dim AskedCount    As Long         'How many
    Dim Answer        As String       'What the ECB answered
    Dim PerEuro()     As Double       'The rate of each currency asked, per euro; 0 for none
    Dim Days()        As Date         'The day of each
    Dim ReportingPer  As Double       'Units of the reporting currency per euro
    Dim CodePer       As Double       'Units of a currency per euro
    Dim NewRates()    As Double       'The rate of each row; 0 for none
    Dim Latest        As Date         'The latest fixing used
    Dim Written       As Boolean      'A rate was written
    Dim Changed       As Boolean      'An input changed while the answer was awaited
    Dim ReportRow     As Long         'The row of the reporting currency; 0 for none
    Dim ReportKept    As Variant      'Its rate before
    Dim ErrNumber     As Long         'The error to raise again
    Dim ErrText       As String       'Its description

'------------------------------------------------------------------------------
' FIND THE ROWS
'------------------------------------------------------------------------------
        Priced = 0
        Without = vbNullString
        Problem = vbNullString
        Fixing = vbNullString
        Set ParamsSheet = GetSheet(SH_PARAMS)
        HeaderRow = NamedRow(NM_FX_TABLE)
        If HeaderRow = 0 Then
            Problem = "the workbook name '" & NM_FX_TABLE & "' is missing or no longer refers to a cell on " & SH_PARAMS & "."
            LookUpEcbRates = "RESULT=REFUSED; reason=" & Problem
            Exit Function
        End If
        If NamedCell(NM_FX_PRICE_DATE) Is Nothing Or NamedCell(NM_FX_PRICE_SRC) Is Nothing Or _
           NamedCell(NM_FX_PRICE_FP) Is Nothing Then
            Problem = "the workbook name '" & IIf(NamedCell(NM_FX_PRICE_DATE) Is Nothing, NM_FX_PRICE_DATE, _
                      IIf(NamedCell(NM_FX_PRICE_SRC) Is Nothing, NM_FX_PRICE_SRC, NM_FX_PRICE_FP)) & "' is missing."
            LookUpEcbRates = "RESULT=REFUSED; reason=" & Problem
            Exit Function
        End If
        Reporting = UCase$(Trim$(SafeStr(NamedCell(PRM_REPCCY).Value2)))
        AsOfValue = NamedCell(PRM_ASOF).Value2
        If Not Reporting Like "[A-Z][A-Z][A-Z]" Then
            Problem = "the reporting currency '" & Reporting & "' is not a three-letter code."
        ElseIf VarType(AsOfValue) <> vbDouble Then
            Problem = "the reporting date is not a date."
        End If
        If Len(Problem) > 0 Then
            LookUpEcbRates = "RESULT=REFUSED; reason=" & Problem
            Exit Function
        End If
        AsOf = CDate(Int(AsOfValue))
        KeyColumn = NamedTableCol(NM_FX_TABLE)
        r = HeaderRow + 1
        Do While Not IsBlankCell(TableCell(ParamsSheet, r, KeyColumn).Value)
            If Len(ClosingPair(ParamsSheet, r, KeyColumn)) > 0 Then
                Count = Count + 1
                ReDim Preserve PairRows(1 To Count)
                ReDim Preserve Codes(1 To Count)
                ReDim Preserve Kept(1 To Count)
                PairRows(Count) = r
                Codes(Count) = UCase$(Trim$(SafeStr(TableCell(ParamsSheet, r, KeyColumn).Value2)))
                Kept(Count) = TableCell(ParamsSheet, r, KeyColumn + FX_COL_RATE).Value2
            ElseIf ReportRow = 0 And IsReportingRow(ParamsSheet, r, KeyColumn) Then
                ReportRow = r
                ReportKept = TableCell(ParamsSheet, r, KeyColumn + FX_COL_RATE).Value2
            End If
            r = r + 1
        Loop
        If Count = 0 Then
            Problem = "no currency of the FX table has a pair."
            LookUpEcbRates = "RESULT=REFUSED; reason=" & Problem
            Exit Function
        End If
        KeptDate = NamedCell(NM_FX_PRICE_DATE).Value2
        KeptSource = NamedCell(NM_FX_PRICE_SRC).Value2
        KeptPrint = NamedCell(NM_FX_PRICE_FP).Value2

'------------------------------------------------------------------------------
' ASK THE ECB
'------------------------------------------------------------------------------
    'Only three-letter codes go into the request: nothing else leaves the workbook
        ReDim Asked(1 To Count + 1)
        If Reporting <> "EUR" Then
            AskedCount = 1
            Asked(1) = Reporting
        End If
        For i = 1 To Count
            If Codes(i) Like "[A-Z][A-Z][A-Z]" And Codes(i) <> "EUR" And IndexOfCode(Asked, AskedCount, Codes(i)) = 0 Then
                AskedCount = AskedCount + 1
                Asked(AskedCount) = Codes(i)
            End If
        Next i
        If AskedCount = 0 Then
            Problem = "no currency of the FX table is a code the ECB could publish a rate for."
            LookUpEcbRates = "RESULT=REFUSED; reason=" & Problem
            Exit Function
        End If
        ReDim Preserve Asked(1 To AskedCount)
        On Error GoTo Failed
        mRunning = True
        Answer = EcbAnswer(EcbRequestUrl(Asked, AsOf), Problem)
        If Len(Problem) = 0 Then
            If Not EcbLatestRates(Answer, Asked, AsOf, PerEuro, Days, Problem) Then
                If Len(Problem) = 0 Then
                    Problem = "the answer of the ECB data portal could not be read."
                End If
            End If
        End If
        If Len(Problem) > 0 Then
            mRunning = False
            LookUpEcbRates = "RESULT=REFUSED; reason=" & Problem
            Exit Function
        End If

'------------------------------------------------------------------------------
' WORK OUT THE RATES
'------------------------------------------------------------------------------
        ReportingPer = 1#
        If Reporting <> "EUR" Then
            ReportingPer = PerEuro(IndexOfCode(Asked, AskedCount, Reporting))
            Latest = Days(IndexOfCode(Asked, AskedCount, Reporting))
            If ReportingPer <= 0# Then
                mRunning = False
                Problem = "the ECB published no reference rate of the reporting currency " & Reporting & " from " & _
                          Format$(AsOf - ECB_LOOKBACK_DAYS, "yyyy-mm-dd") & " to " & Format$(AsOf, "yyyy-mm-dd") & _
                          ", so no rate can be crossed through the euro."
                LookUpEcbRates = "RESULT=REFUSED; reason=" & Problem
                Exit Function
            End If
        End If
        ReDim NewRates(1 To Count)
        For i = 1 To Count
            CodePer = 0#
            If Codes(i) = "EUR" Then
                CodePer = 1#
            Else
                k = IndexOfCode(Asked, AskedCount, Codes(i))
                If k > 0 Then
                    CodePer = PerEuro(k)
                    If CodePer > 0# And Days(k) > Latest Then
                        Latest = Days(k)
                    End If
                End If
            End If
            If CodePer > 0# Then
                NewRates(i) = CodePer / ReportingPer
                Priced = Priced + 1
            Else
                Without = Without & IIf(Len(Without) > 0, ",", "") & Codes(i)
            End If
        Next i
        If Priced = 0 Then
            mRunning = False
            Without = vbNullString
            Problem = "the ECB published no reference rate for any currency of the FX table from " & _
                      Format$(AsOf - ECB_LOOKBACK_DAYS, "yyyy-mm-dd") & " to " & Format$(AsOf, "yyyy-mm-dd") & "."
            LookUpEcbRates = "RESULT=REFUSED; reason=" & Problem
            Exit Function
        End If

'------------------------------------------------------------------------------
' WRITE THE RATES
'------------------------------------------------------------------------------
    'Params stays editable while the answer is awaited: the inputs must still be those asked for
        Changed = (NamedCell(PRM_ASOF).Value2 <> AsOfValue) Or _
                  (UCase$(Trim$(SafeStr(NamedCell(PRM_REPCCY).Value2))) <> Reporting)
        For i = 1 To Count
            If UCase$(Trim$(SafeStr(TableCell(ParamsSheet, PairRows(i), KeyColumn).Value2))) <> Codes(i) Then
                Changed = True
            End If
        Next i
        If ReportRow > 0 Then
            If Not IsReportingRow(ParamsSheet, ReportRow, KeyColumn) Then
                Changed = True
            End If
        End If
        If Changed Then
            mRunning = False
            Priced = 0
            Without = vbNullString
            Problem = "the reporting date, the reporting currency or a currency code was changed while the rates came in, " & _
                      "so no rate was changed."
            LookUpEcbRates = "RESULT=REFUSED; reason=" & Problem
            Exit Function
        End If
        For i = 1 To Count
            If NewRates(i) > 0# Then
                Written = True
                TableCell(ParamsSheet, PairRows(i), KeyColumn + FX_COL_RATE).Value2 = NewRates(i)
                If gTestFault = "ecb-write" Then
                    Err.Raise ERR_INJECTED_FAULT, "M_Main.LookUpEcbRates", "Injected failure while the ECB rates are written."
                End If
            End If
        Next i
        If ReportRow > 0 Then
            Written = True
            TableCell(ParamsSheet, ReportRow, KeyColumn + FX_COL_RATE).Value2 = 1#
        End If
        NamedCell(NM_FX_PRICE_DATE).Value2 = AsOfValue
        NamedCell(NM_FX_PRICE_SRC).Value2 = FX_SOURCE_ECB
        PutFxPrint CORE_Engine.FxTableFingerprint()
        mRunning = False
        If Latest <> AsOf And Latest > 0 Then
            Fixing = Format$(Latest, "yyyy-mm-dd")
        End If
        LookUpEcbRates = "RESULT=PRICED; rates=" & Priced & "; without_price=" & _
                         IIf(Len(Without) > 0, UBound(Split(Without, ",")) + 1, 0) & _
                         IIf(Len(Without) > 0, "; codes=" & Without, "") & "; source=ECB" & _
                         IIf(Len(Fixing) > 0, "; fixing=" & Fixing, "")
        Exit Function

Failed:
        ErrNumber = Err.Number
        ErrText = Err.Description
        On Error Resume Next
        If Written Then
            For i = 1 To Count
                TableCell(ParamsSheet, PairRows(i), KeyColumn + FX_COL_RATE).Value2 = Kept(i)
            Next i
            If ReportRow > 0 Then
                TableCell(ParamsSheet, ReportRow, KeyColumn + FX_COL_RATE).Value2 = ReportKept
            End If
            NamedCell(NM_FX_PRICE_DATE).Value2 = KeptDate
            NamedCell(NM_FX_PRICE_SRC).Value2 = KeptSource
            PutFxPrint KeptPrint
        End If
        mRunning = False
        On Error GoTo 0
        Err.Raise ErrNumber, "M_Main.LookUpEcbRates", ErrText

End Function


Public Function EcbRequestUrl( _
    ByRef Codes() As String, _
    ByVal AsOf As Date) _
    As String
'
'==============================================================================
'                                EcbRequestUrl
'------------------------------------------------------------------------------
' PURPOSE
'   The HTTPS address that asks the ECB data portal for the daily euro
'   reference rates of Codes from ECB_LOOKBACK_DAYS before AsOf to AsOf,
'   as CSV with the data columns only (#745). It carries the currency codes
'   and the two dates, nothing else.
'
' UPDATED
'   2026-11-22
'==============================================================================
'
        EcbRequestUrl = ECB_EXR_URL & Join(Codes, "+") & ".EUR.SP00.A?startPeriod=" & _
                        Format$(AsOf - ECB_LOOKBACK_DAYS, "yyyy-mm-dd") & "&endPeriod=" & Format$(AsOf, "yyyy-mm-dd") & _
                        "&format=csvdata&detail=dataonly"

End Function


Private Function EcbAnswer( _
    ByVal Url As String, _
    ByRef Problem As String) _
    As String
'
'==============================================================================
'                                  EcbAnswer
'------------------------------------------------------------------------------
' PURPOSE
'   Send the request Url to the ECB data portal and return its answer, the
'   text of a CSV file (#745). HTTP 404 is the portal's answer when none of
'   the currencies asked has a series: it is an empty answer, no rate.
'
' METHOD
'   MSXML2.XMLHTTP, by late binding, sends the request through the
'   connection and the proxy settings of Windows, as a browser does, and
'   asynchronously, so that the wait can end after PRICE_WAIT_SECONDS. The
'   request asks the portal itself, not a copy kept by Windows. No web
'   query and no connection is left in the workbook.
'
' RETURNS
'   The answer; "" with Problem set when there is none: no connection, a
'   time-out, an HTTP status other than 200 and 404, or an address that is
'   not HTTPS.
'
' ERROR POLICY
'   Never raises. gTestFault = "ecb-answer" returns gTestEcbAnswer, and so
'   does "ecb-write", whose failure LookUpEcbRates injects later, and
'   "ecb-edit", which also moves the reporting date a day on, as a user
'   could while the answer is awaited;
'   "ecb-offline", "ecb-http" and "ecb-timeout" give the problem of no
'   connection, of HTTP 503 and of a time-out; none of them sends anything.
'
' UPDATED
'   2026-11-22
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Request     As Object    'MSXML2.XMLHTTP
    Dim Deadline    As Date      'When to stop waiting
    Dim Status      As Long      'The HTTP status
    Dim Failure     As String    'What the request reported

'------------------------------------------------------------------------------
' FAULTS
'------------------------------------------------------------------------------
        Problem = vbNullString
        Select Case gTestFault
            Case "ecb-answer", "ecb-write"
                EcbAnswer = gTestEcbAnswer
                Exit Function
            Case "ecb-edit"
                EcbAnswer = gTestEcbAnswer
                NamedCell(PRM_ASOF).Value2 = NamedCell(PRM_ASOF).Value2 + 1
                Exit Function
            Case "ecb-offline"
                Problem = EcbUnreachable("the server name could not be resolved")
                Exit Function
            Case "ecb-http"
                Problem = "the ECB data portal answered HTTP 503 (Service Unavailable). Try again later."
                Exit Function
            Case "ecb-timeout"
                Problem = "the ECB data portal did not answer within " & PRICE_WAIT_SECONDS & " seconds. " & _
                          "Check the connection and try again."
                Exit Function
        End Select
        If LCase$(Left$(Url, 8)) <> "https://" Then
            Problem = "the address of the ECB data portal is not HTTPS, so nothing was sent."
            Exit Function
        End If

'------------------------------------------------------------------------------
' SEND AND WAIT
'------------------------------------------------------------------------------
        On Error Resume Next
        Set Request = CreateObject("MSXML2.XMLHTTP.6.0")
        If Request Is Nothing Then
            Err.Clear
            On Error GoTo 0
            Problem = "this Excel cannot send a web request (MSXML2.XMLHTTP is missing), so no rate was looked up."
            Exit Function
        End If
        Request.Open "GET", Url, True
        Request.setRequestHeader "Cache-Control", "no-cache"
        Request.setRequestHeader "If-Modified-Since", "Sat, 01 Jan 2000 00:00:00 GMT"
        Request.send
        If Err.Number <> 0 Then
            Failure = Err.Description
            Err.Clear
            On Error GoTo 0
            Problem = EcbUnreachable(Failure)
            Exit Function
        End If
        Deadline = Now + TimeSerial(0, 0, PRICE_WAIT_SECONDS)
        Do While Request.readyState <> 4
            If Now >= Deadline Then
                Request.abort
                Err.Clear
                On Error GoTo 0
                Problem = "the ECB data portal did not answer within " & PRICE_WAIT_SECONDS & " seconds. " & _
                          "Check the connection and try again."
                Exit Function
            End If
            DoEvents
        Loop
        Status = Request.Status
        If Err.Number <> 0 Then
            Failure = Err.Description
            Err.Clear
            On Error GoTo 0
            Problem = EcbUnreachable(Failure)
            Exit Function
        End If

'------------------------------------------------------------------------------
' READ THE STATUS
'------------------------------------------------------------------------------
    'Windows reports a request that did not reach the server as a status of 0, or 12000 and above
        Select Case Status
            Case 200
                EcbAnswer = Request.responseText
            Case 404
                EcbAnswer = vbNullString
            Case 0, 12000 To 12999
                Problem = EcbUnreachable(IIf(Status = 0, "no answer", "network error " & Status))
            Case Else
                Problem = "the ECB data portal answered HTTP " & Status & _
                          IIf(Len(SafeStr(Request.statusText)) > 0, " (" & SafeStr(Request.statusText) & ")", "") & _
                          ". Try again later."
        End Select
        If Err.Number <> 0 And Len(Problem) = 0 Then
            EcbAnswer = vbNullString
            Problem = EcbUnreachable(Err.Description)
        End If
        Err.Clear
        On Error GoTo 0

End Function


Private Function EcbUnreachable( _
    ByVal Failure As String) _
    As String
'
'==============================================================================
'                               EcbUnreachable
'------------------------------------------------------------------------------
' PURPOSE
'   The problem of a request that did not reach the ECB data portal (#745),
'   with what Windows reported.
'
' UPDATED
'   2026-11-22
'==============================================================================
'
        EcbUnreachable = "the ECB data portal could not be reached (" & Trim$(Replace(Failure, vbCrLf, " ")) & _
                         "). Check the connection, and that the network lets Excel reach data-api.ecb.europa.eu."

End Function


Public Function EcbLatestRates( _
    ByVal Answer As String, _
    ByRef Codes() As String, _
    ByVal AsOf As Date, _
    ByRef PerEuro() As Double, _
    ByRef Days() As Date, _
    ByRef Problem As String) _
    As Boolean
'
'==============================================================================
'                               EcbLatestRates
'------------------------------------------------------------------------------
' PURPOSE
'   Read an answer of the ECB data portal, a CSV file of daily euro
'   reference rates, and keep for each of Codes the latest rate on or
'   before AsOf (#745), as published: units of the currency per euro.
'
' INPUTS
'   Answer: the text of the CSV file; blank when the ECB published nothing.
'   Codes: the currencies asked for, upper case, from index 1.
'
' METHOD
'   The header names the columns: CURRENCY, TIME_PERIOD and OBS_VALUE are
'   needed, CURRENCY_DENOM is read when present and must be EUR. Fields may
'   be quoted, and every line has as many fields as the header, so that a
'   decimal comma is never read as two fields. A rate without a value is
'   no rate; a currency that was not asked for is skipped; a day after AsOf
'   is not used. Dates are read
'   digit by digit and numbers with Val, so that the Windows locale does
'   not matter.
'
' RETURNS
'   True with PerEuro and Days, of the bounds of Codes, set: 0 for a
'   currency with no rate. False with Problem set when the answer is not
'   the CSV expected, or a line is malformed: then nothing is used.
'
' UPDATED
'   2026-11-22
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Lines       As Variant   'The lines of the answer
    Dim Fields      As Variant   'The fields of a line
    Dim n           As Long      'Line
    Dim HeaderLine  As Long      'The line of the header; -1 before it is found
    Dim CodeCol     As Long      'Column CURRENCY; -1 for none
    Dim DenomCol    As Long      'Column CURRENCY_DENOM; -1 for none
    Dim DayCol      As Long      'Column TIME_PERIOD; -1 for none
    Dim ValueCol    As Long      'Column OBS_VALUE; -1 for none
    Dim HeaderLast  As Long      'The last field of the header
    Dim c           As Long      'Column
    Dim k           As Long      'Currency
    Dim DayText     As String    'A date as text
    Dim ValueText   As String    'A rate as text
    Dim Day         As Date      'A date
    Dim Missing     As String    'A column the header lacks

'------------------------------------------------------------------------------
' HEADER
'------------------------------------------------------------------------------
        Problem = vbNullString
        ReDim PerEuro(LBound(Codes) To UBound(Codes))
        ReDim Days(LBound(Codes) To UBound(Codes))
        Lines = Split(Replace(Answer, vbCr, vbNullString), vbLf)
        HeaderLine = -1
        For n = LBound(Lines) To UBound(Lines)
            If Len(Trim$(Lines(n))) > 0 Then
                HeaderLine = n
                Exit For
            End If
        Next n
        If HeaderLine < 0 Then
            EcbLatestRates = True
            Exit Function
        End If
        Fields = CsvLineFields(CStr(Lines(HeaderLine)))
        CodeCol = -1
        DenomCol = -1
        DayCol = -1
        ValueCol = -1
        If Not IsEmpty(Fields) Then
            HeaderLast = UBound(Fields)
            For c = LBound(Fields) To UBound(Fields)
                Select Case UCase$(Trim$(Fields(c)))
                    Case "CURRENCY"
                        CodeCol = c
                    Case "CURRENCY_DENOM"
                        DenomCol = c
                    Case "TIME_PERIOD"
                        DayCol = c
                    Case "OBS_VALUE"
                        ValueCol = c
                End Select
            Next c
        End If
        If CodeCol < 0 Then
            Missing = "CURRENCY"
        ElseIf DayCol < 0 Then
            Missing = "TIME_PERIOD"
        ElseIf ValueCol < 0 Then
            Missing = "OBS_VALUE"
        End If
        If Len(Missing) > 0 Then
            Problem = "the answer of the ECB data portal is not the CSV expected: its first line has no column " & Missing & "."
            Exit Function
        End If

'------------------------------------------------------------------------------
' RATES
'------------------------------------------------------------------------------
        For n = HeaderLine + 1 To UBound(Lines)
            If Len(Trim$(Lines(n))) > 0 Then
                Fields = CsvLineFields(CStr(Lines(n)))
                If IsEmpty(Fields) Then
                    Problem = EcbMalformed(n + 1, "a quote is not closed")
                    Exit Function
                End If
                If UBound(Fields) <> HeaderLast Then
                    Problem = EcbMalformed(n + 1, "it has " & (UBound(Fields) + 1) & " fields, the header " & (HeaderLast + 1))
                    Exit Function
                End If
                If DenomCol >= 0 Then
                    If UCase$(Trim$(Fields(DenomCol))) <> "EUR" Then
                        Problem = EcbMalformed(n + 1, "the rate is not against the euro")
                        Exit Function
                    End If
                End If
                k = IndexOfCode(Codes, UBound(Codes), UCase$(Trim$(Fields(CodeCol))))
                If k > 0 Then
                    DayText = Trim$(Fields(DayCol))
                    If Not DayText Like "####-##-##" Then
                        Problem = EcbMalformed(n + 1, "the date '" & DayText & "' is not yyyy-mm-dd")
                        Exit Function
                    End If
                    Day = DateSerial(CInt(Left$(DayText, 4)), CInt(Mid$(DayText, 6, 2)), CInt(Right$(DayText, 2)))
                    If Format$(Day, "yyyy-mm-dd") <> DayText Then
                        Problem = EcbMalformed(n + 1, "the date '" & DayText & "' does not exist")
                        Exit Function
                    End If
                    ValueText = Trim$(Fields(ValueCol))
                    If Len(ValueText) > 0 And UCase$(ValueText) <> "NAN" Then
                        If Not IsDecimalText(ValueText) Then
                            Problem = EcbMalformed(n + 1, "the rate '" & ValueText & "' is not a number")
                            Exit Function
                        End If
                        If Val(ValueText) <= 0# Then
                            Problem = EcbMalformed(n + 1, "the rate '" & ValueText & "' is not above zero")
                            Exit Function
                        End If
                        If Day <= AsOf And Day > Days(k) Then
                            PerEuro(k) = Val(ValueText)
                            Days(k) = Day
                        End If
                    End If
                End If
            End If
        Next n
        EcbLatestRates = True

End Function


Private Function EcbMalformed( _
    ByVal LineNumber As Long, _
    ByVal What As String) _
    As String
'
'==============================================================================
'                                EcbMalformed
'------------------------------------------------------------------------------
' PURPOSE
'   The problem of a malformed line of the answer of the ECB (#745).
'
' UPDATED
'   2026-11-22
'==============================================================================
'
        EcbMalformed = "line " & LineNumber & " of the answer of the ECB data portal is malformed: " & What & _
                       ", so no rate was used."

End Function


Private Function IndexOfCode( _
    ByRef Codes() As String, _
    ByVal Last As Long, _
    ByVal Code As String) _
    As Long
'
'==============================================================================
'                                 IndexOfCode
'------------------------------------------------------------------------------
' PURPOSE
'   The index of Code among Codes, from index 1 to Last; 0 when it is not
'   there or Last is 0 (#745).
'
' UPDATED
'   2026-11-22
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim i           As Long      'Index

'------------------------------------------------------------------------------
' SEARCH
'------------------------------------------------------------------------------
        For i = 1 To Last
            If Codes(i) = Code Then
                IndexOfCode = i
                Exit Function
            End If
        Next i

End Function


Private Function IsDecimalText( _
    ByVal Text As String) _
    As Boolean
'
'==============================================================================
'                                IsDecimalText
'------------------------------------------------------------------------------
' PURPOSE
'   Whether Text is digits with at most one decimal point between digits,
'   as the ECB writes a rate (#745): no sign, no exponent, no separator.
'
' UPDATED
'   2026-11-22
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim i           As Long      'Character
    Dim Ch          As String    'One character
    Dim Points      As Long      'Decimal points seen

'------------------------------------------------------------------------------
' CHECK
'------------------------------------------------------------------------------
        If Len(Text) = 0 Or Left$(Text, 1) = "." Or Right$(Text, 1) = "." Then
            Exit Function
        End If
        For i = 1 To Len(Text)
            Ch = Mid$(Text, i, 1)
            If Ch = "." Then
                Points = Points + 1
            ElseIf Ch < "0" Or Ch > "9" Then
                Exit Function
            End If
        Next i
        IsDecimalText = (Points <= 1)

End Function


Private Function CsvLineFields( _
    ByVal LineText As String) _
    As Variant
'
'==============================================================================
'                                CsvLineFields
'------------------------------------------------------------------------------
' PURPOSE
'   The fields of one line of a CSV file, comma-separated, a field possibly
'   in double quotes with "" for a quote inside (#745).
'
' RETURNS
'   A String array from 0; Empty when a quote is not closed.
'
' UPDATED
'   2026-11-22
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Result()    As String    'The fields
    Dim Count       As Long      'Fields found
    Dim Current     As String    'The field being read
    Dim Quoted      As Boolean   'Inside quotes
    Dim i           As Long      'Character
    Dim Ch          As String    'One character

'------------------------------------------------------------------------------
' SPLIT
'------------------------------------------------------------------------------
        ReDim Result(0 To 0)
        i = 1
        Do While i <= Len(LineText)
            Ch = Mid$(LineText, i, 1)
            If Quoted Then
                If Ch = """" Then
                    If Mid$(LineText, i + 1, 1) = """" Then
                        Current = Current & """"
                        i = i + 1
                    Else
                        Quoted = False
                    End If
                Else
                    Current = Current & Ch
                End If
            ElseIf Ch = """" Then
                Quoted = True
            ElseIf Ch = "," Then
                ReDim Preserve Result(0 To Count)
                Result(Count) = Current
                Count = Count + 1
                Current = vbNullString
            Else
                Current = Current & Ch
            End If
            i = i + 1
        Loop
        If Quoted Then
            Exit Function
        End If
        ReDim Preserve Result(0 To Count)
        Result(Count) = Current
        CsvLineFields = Result

End Function


Private Function StartPriceLookup( _
    ByRef Problem As String) _
    As Boolean
'
'==============================================================================
'                             StartPriceLookup
'------------------------------------------------------------------------------
' PURPOSE
'   Start a lookup of closing prices (#721): write the formula of each row
'   that has a pair (ClosingPair) into the hidden scratch column, and keep
'   what FinishPriceLookup needs.
'
' METHOD
'   1. Find the rows with a pair and the row of the reporting currency,
'      and note their rates and the date of the prices, to put them back.
'   2. Set calculation automatic and write the formulas through Formula2
'      by late binding, so that the code compiles in an Excel without it;
'      an Excel that refuses it, or a fault injected, is a refusal.
'   From here mRunning is set, so that no other operation, and no second
'   lookup, starts while the prices are awaited: every way out releases it.
'
' RETURNS
'   True when the lookup started. False with Problem set when it was
'   refused, nothing changed.
'
' ERROR POLICY
'   A refusal is returned, not raised. An error puts everything back and is
'   raised again.
'
' UPDATED
'   2026-11-21
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim HeaderRow     As Long         'Header row of the FX table
    Dim r             As Long         'Row
    Dim i             As Long         'Row with a pair
    Dim PriceCell     As Range        'A scratch cell
    Dim Formula       As String       'The formula written into a scratch cell
    Dim Refused       As Boolean      'The formula could not be written
    Dim ErrNumber     As Long         'The error to raise again
    Dim ErrText       As String       'Its description

'------------------------------------------------------------------------------
' FIND THE ROWS
'------------------------------------------------------------------------------
        Problem = vbNullString
        Set mPriceSheet = GetSheet(SH_PARAMS)
        HeaderRow = NamedRow(NM_FX_TABLE)
        If HeaderRow = 0 Then
            Problem = "the workbook name '" & NM_FX_TABLE & "' is missing or no longer refers to a cell on " & SH_PARAMS & "."
            Exit Function
        End If
        mPriceKeyColumn = NamedTableCol(NM_FX_TABLE)
        mPriceFirstRow = HeaderRow + 1
        mPriceLastRow = HeaderRow
        Do While Not IsBlankCell(TableCell(mPriceSheet, mPriceLastRow + 1, mPriceKeyColumn).Value)
            mPriceLastRow = mPriceLastRow + 1
        Loop
        ReDim mPriceRows(1 To mPriceLastRow - HeaderRow + 1)
        ReDim mPricePairs(1 To mPriceLastRow - HeaderRow + 1)
        ReDim mPriceKept(1 To mPriceLastRow - HeaderRow + 1)
        mPriceCount = 0
        On Error Resume Next
        mPriceKeptDate = NamedCell(NM_FX_PRICE_DATE).Value2
        If Err.Number <> 0 Then
            Err.Clear
            On Error GoTo 0
            Problem = "the workbook name '" & NM_FX_PRICE_DATE & "' is missing."
            Exit Function
        End If
        mPriceKeptSource = NamedCell(NM_FX_PRICE_SRC).Value2
        If Err.Number <> 0 Then
            Err.Clear
            On Error GoTo 0
            Problem = "the workbook name '" & NM_FX_PRICE_SRC & "' is missing."
            Exit Function
        End If
        mPriceKeptPrint = NamedCell(NM_FX_PRICE_FP).Value2
        If Err.Number <> 0 Then
            Err.Clear
            On Error GoTo 0
            Problem = "the workbook name '" & NM_FX_PRICE_FP & "' is missing."
            Exit Function
        End If
        On Error GoTo 0
        mPriceReportRow = 0
        For r = HeaderRow + 1 To mPriceLastRow
            If Len(ClosingPair(mPriceSheet, r, mPriceKeyColumn)) > 0 Then
                mPriceCount = mPriceCount + 1
                mPriceRows(mPriceCount) = r
                mPricePairs(mPriceCount) = ClosingPair(mPriceSheet, r, mPriceKeyColumn)
                mPriceKept(mPriceCount) = TableCell(mPriceSheet, r, mPriceKeyColumn + FX_COL_RATE).Value2
            ElseIf mPriceReportRow = 0 And IsReportingRow(mPriceSheet, r, mPriceKeyColumn) Then
                mPriceReportRow = r
                mPriceReportKept = TableCell(mPriceSheet, r, mPriceKeyColumn + FX_COL_RATE).Value2
            End If
        Next r
        If mPriceCount = 0 Then
            Problem = "no currency of the FX table has a pair."
            Exit Function
        End If

'------------------------------------------------------------------------------
' WRITE THE FORMULAS
'------------------------------------------------------------------------------
        On Error GoTo Failed
        mRunning = True
        mPriceActive = True
        mPriceWritten = False
        mPriceAsOf = NamedCell(PRM_ASOF).Value2
        mPriceMode = Application.Calculation
        Application.Calculation = xlCalculationAutomatic
        ClearPriceScratch mPriceSheet, mPriceKeyColumn, mPriceFirstRow, mPriceLastRow
        On Error Resume Next
        For i = 1 To mPriceCount
            Refused = (gTestFault = "prices-unavailable")
            If Not Refused Then
                Set PriceCell = TableCell(mPriceSheet, mPriceRows(i), mPriceKeyColumn + FX_COL_PRICE)
                Formula = PriceFormula(mPricePairs(i))
                If gTestFault = "prices-name-error" Then
                    Formula = "=#NAME?"
                ElseIf gTestFault = "prices-fixed" Then
                    Formula = IIf(i = 1, "=NA()", "=2")
                ElseIf gTestFault = "prices-none" Then
                    Formula = "=NA()"
                End If
                CallByName PriceCell, "Formula2", VbLet, Formula
                Refused = (Err.Number <> 0)
            End If
            If Refused Then
                Exit For
            End If
        Next i
        If Refused Then
            Err.Clear
            On Error GoTo Failed
            AbandonPriceLookup
            Problem = "STOCKHISTORY is not available in this Excel (Microsoft 365 is needed), so no price was looked up."
            Exit Function
        End If
        On Error GoTo Failed
        mPriceStarted = Now
        mPriceDeadline = Now + TimeSerial(0, 0, PRICE_WAIT_SECONDS)
        StartPriceLookup = True
        Exit Function

Failed:
        ErrNumber = Err.Number
        ErrText = Err.Description
        AbandonPriceLookup
        Err.Raise ErrNumber, "M_Main.StartPriceLookup", ErrText

End Function


Private Function FinishPriceLookup( _
    ByRef Priced As Long, _
    ByRef Without As String, _
    ByRef Problem As String) _
    As String
'
'==============================================================================
'                             FinishPriceLookup
'------------------------------------------------------------------------------
' PURPOSE
'   End the lookup StartPriceLookup started (#721): read the scratch cells,
'   write each price above zero into its rate cell, the reporting date into
'   the hidden cell NM_FX_PRICE_DATE and the source into NM_FX_PRICE_SRC,
'   which the sheet shows beside the button (#745). A row with no price
'   keeps its rate and is listed in Without; the row of the reporting
'   currency takes 1, the rate a run requires of it.
'
' METHOD
'   Refuse, changing no rate, when a cell evaluated to #NAME? (an Excel
'   that takes the formula but has no STOCKHISTORY), when the reporting
'   date, the reporting currency or a code was changed while the data came
'   in, or when no price came for any currency, in which case the rates and
'   the date stay and Problem says what the first cell showed. Then empty
'   the scratch column and put the calculation mode back.
'
' RETURNS
'   "RESULT=PRICED; rates=N; without_price=M[; codes=...]" or
'   "RESULT=REFUSED; reason=...".
'
' ERROR POLICY
'   A refusal is returned, not raised. An error puts the rates back as they
'   were and is raised again.
'
' UPDATED
'   2026-11-21
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim i             As Long         'Row with a pair
    Dim PriceCell     As Range        'A scratch cell
    Dim Found         As Variant      'What the formula returned, Empty when it is no price
    Dim Unavailable   As Boolean      'A formula evaluated to #NAME?, STOCKHISTORY is unknown
    Dim Changed       As Boolean      'An input changed while the data came in
    Dim TimedOut      As Boolean      'A price was still awaited
    Dim FirstState    As String       'What the first scratch cell returned
    Dim CalcText      As String       'The calculation state
    Dim Code          As String       'A currency
    Dim ErrNumber     As Long         'The error to raise again
    Dim ErrText       As String       'Its description

'------------------------------------------------------------------------------
' CHECK WHAT CAME
'------------------------------------------------------------------------------
        Priced = 0
        Without = vbNullString
        Problem = vbNullString
        On Error GoTo Failed
        TimedOut = PricesPending()
        Select Case Application.CalculationState
            Case xlDone
                CalcText = "done"
            Case xlCalculating
                CalcText = "calculating"
            Case Else
                CalcText = "pending"
        End Select
        FirstState = ScratchState(TableCell(mPriceSheet, mPriceRows(1), mPriceKeyColumn + FX_COL_PRICE).Value2)
    'An Excel that takes the formula but does not know STOCKHISTORY returns #NAME?
        For i = 1 To mPriceCount
            Found = TableCell(mPriceSheet, mPriceRows(i), mPriceKeyColumn + FX_COL_PRICE).Value2
            If IsError(Found) Then
                If ErrorNumber(Found) = 2029 Then
                    Unavailable = True
                End If
            End If
        Next i
    'The reporting date and the pairs must still be the ones asked for
        Changed = (NamedCell(PRM_ASOF).Value2 <> mPriceAsOf)
        For i = 1 To mPriceCount
            If ClosingPair(mPriceSheet, mPriceRows(i), mPriceKeyColumn) <> mPricePairs(i) Then
                Changed = True
            End If
        Next i
        If mPriceReportRow > 0 Then
            If Not IsReportingRow(mPriceSheet, mPriceReportRow, mPriceKeyColumn) Then
                Changed = True
            End If
        End If
        If Unavailable Or Changed Then
            If Unavailable Then
                Problem = "STOCKHISTORY is not available in this Excel (Microsoft 365 is needed), so no price was looked up."
            Else
                Problem = "the reporting date, the reporting currency or a currency code was changed while the prices came in, so no rate was changed."
            End If
            AbandonPriceLookup
            FinishPriceLookup = "RESULT=REFUSED; reason=" & Problem
            Exit Function
        End If

'------------------------------------------------------------------------------
' WRITE THE RATES
'------------------------------------------------------------------------------
        For i = 1 To mPriceCount
            Set PriceCell = TableCell(mPriceSheet, mPriceRows(i), mPriceKeyColumn + FX_COL_PRICE)
            Found = Empty
            If VarType(PriceCell.Value2) = vbDouble Then
                If PriceCell.Value2 > 0# Then
                    Found = PriceCell.Value2
                End If
            End If
            If IsEmpty(Found) Then
                Code = UCase$(Trim$(SafeStr(TableCell(mPriceSheet, mPriceRows(i), mPriceKeyColumn).Value2)))
                Without = Without & IIf(Len(Without) > 0, ",", "") & Code
            Else
                mPriceWritten = True
                TableCell(mPriceSheet, mPriceRows(i), mPriceKeyColumn + FX_COL_RATE).Value2 = Found
                Priced = Priced + 1
            End If
        Next i

    'No price at all: no rate and no date changes
        If Priced = 0 Then
            If TimedOut Then
                Problem = "no price came within " & PRICE_WAIT_SECONDS & " seconds (the first, " & mPricePairs(1) & _
                          ", showed " & FirstState & "; the calculation was " & CalcText & "). Check the connection and try again."
            Else
                Problem = "no price came for any currency (the first, " & mPricePairs(1) & ", returned " & FirstState & _
                          "). Check the connection and that this Excel can look up prices."
            End If
            AbandonPriceLookup
            FinishPriceLookup = "RESULT=REFUSED; reason=" & Problem
            Exit Function
        End If

    'The reporting currency is 1 unit of itself, as a run requires
        If mPriceReportRow > 0 Then
            TableCell(mPriceSheet, mPriceReportRow, mPriceKeyColumn + FX_COL_RATE).Value2 = 1#
        End If

    'The date and the source of the rates, shown beside the button
        NamedCell(NM_FX_PRICE_DATE).Value2 = mPriceAsOf
        NamedCell(NM_FX_PRICE_SRC).Value2 = FX_SOURCE_CLOSING
        PutFxPrint CORE_Engine.FxTableFingerprint()
        ReleasePriceLookup True
        FinishPriceLookup = "RESULT=PRICED; rates=" & Priced & "; without_price=" & _
                            IIf(Len(Without) > 0, UBound(Split(Without, ",")) + 1, 0) & _
                            IIf(Len(Without) > 0, "; codes=" & Without, "")
        Exit Function

Failed:
        ErrNumber = Err.Number
        ErrText = Err.Description
        AbandonPriceLookup
        Err.Raise ErrNumber, "M_Main.FinishPriceLookup", ErrText

End Function


Private Sub AbandonPriceLookup()
'
'==============================================================================
'                             AbandonPriceLookup
'------------------------------------------------------------------------------
' PURPOSE
'   End a lookup of closing prices that did not succeed (#721): when
'   FinishPriceLookup had begun to write rates, put them and the date and
'   source of the prices back as they were; then empty the scratch column
'   and release the lookup. A rate is not put back when none was written, so that an edit
'   the user made while the prices were awaited stays.
'
' ERROR POLICY
'   Never raises.
'
' UPDATED
'   2026-11-21
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim i           As Long      'Row with a pair

'------------------------------------------------------------------------------
' PUT BACK AND RELEASE
'------------------------------------------------------------------------------
        On Error Resume Next
        If mPriceActive And mPriceWritten Then
            For i = 1 To mPriceCount
                TableCell(mPriceSheet, mPriceRows(i), mPriceKeyColumn + FX_COL_RATE).Value2 = mPriceKept(i)
            Next i
            If mPriceReportRow > 0 Then
                TableCell(mPriceSheet, mPriceReportRow, mPriceKeyColumn + FX_COL_RATE).Value2 = mPriceReportKept
            End If
            NamedCell(NM_FX_PRICE_DATE).Value2 = mPriceKeptDate
            NamedCell(NM_FX_PRICE_SRC).Value2 = mPriceKeptSource
            PutFxPrint mPriceKeptPrint
        End If
        ReleasePriceLookup True

End Sub


Private Sub ReleasePriceLookup( _
    ByVal Clear As Boolean)
'
'==============================================================================
'                             ReleasePriceLookup
'------------------------------------------------------------------------------
' PURPOSE
'   Leave the state of a lookup of closing prices (#721): cancel the next
'   look, empty the scratch column when Clear is True, put the calculation
'   mode back, clear the status bar and release mRunning.
'
' ERROR POLICY
'   Never raises.
'
' UPDATED
'   2026-11-21
'==============================================================================
'
        On Error Resume Next
        CancelPricePoll
        If mPriceActive Then
            If Clear Then
                ClearPriceScratch mPriceSheet, mPriceKeyColumn, mPriceFirstRow, mPriceLastRow
            End If
            Application.Calculation = mPriceMode
        End If
        mPriceActive = False
        mRunning = False
        Application.StatusBar = False

End Sub


Private Sub SchedulePricePoll()
'
'==============================================================================
'                             SchedulePricePoll
'------------------------------------------------------------------------------
' PURPOSE
'   Schedule PollClosingPrices one second from now (#721), with the
'   workbook-qualified name, an apostrophe in the workbook name doubled, as
'   Excel requires.
'
' UPDATED
'   2026-11-21
'==============================================================================
'
        mPollAt = Now + TimeSerial(0, 0, 1)
        mPollName = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!M_Main.PollClosingPrices"
        Application.OnTime mPollAt, mPollName

End Sub


Private Sub CancelPricePoll()
'
'==============================================================================
'                              CancelPricePoll
'------------------------------------------------------------------------------
' PURPOSE
'   Cancel the look SchedulePricePoll scheduled, when one is pending, so
'   that it does not run after the lookup ended or reopen a closed
'   workbook (#721).
'
' ERROR POLICY
'   Never raises: a look that already ran cannot be cancelled, and that is
'   the expected case when the look itself ends the lookup.
'
' UPDATED
'   2026-11-21
'==============================================================================
'
        On Error GoTo Done
        If mPollAt <> 0 Then
            Application.OnTime mPollAt, mPollName, , False
        End If

Done:
        mPollAt = 0
        mPollName = vbNullString

End Sub


Private Function PriceFormula( _
    ByVal Pair As String) _
    As String
'
'==============================================================================
'                               PriceFormula
'------------------------------------------------------------------------------
' PURPOSE
'   The formula that returns the last closing price up to the reporting date
'   (seven days back at most) of a pair such as EUR/USD, which STOCKHISTORY
'   takes as EUR:USD (#721).
'
' UPDATED
'   2026-11-21
'==============================================================================
'

'------------------------------------------------------------------------------
' BUILD
'------------------------------------------------------------------------------
        PriceFormula = "=" & PriceTerm(Replace(Pair, "/", ":"))

End Function


Private Function PriceTerm( _
    ByVal Pair As String) _
    As String
'
'==============================================================================
'                                PriceTerm
'------------------------------------------------------------------------------
' PURPOSE
'   One lookup of PriceFormula: the last row of the history of Pair up to the
'   reporting date, without headers, the closing price (#721).
'
' UPDATED
'   2026-11-21
'==============================================================================
'

'------------------------------------------------------------------------------
' BUILD
'------------------------------------------------------------------------------
        PriceTerm = "LET(r,STOCKHISTORY(""" & Pair & """,AsOfDate-7,AsOfDate,0,0,1),INDEX(r,ROWS(r)))"

End Function


Private Function PricesPending() As Boolean
'
'==============================================================================
'                              PricesPending
'------------------------------------------------------------------------------
' PURPOSE
'   True while a scratch cell still waits for its data (#721): a request of
'   STOCKHISTORY that has not come back shows #BUSY!, and the calculation can
'   be reported done before it has. A cell is pending when it is empty or
'   holds an Excel error that is not one of the final ones (#N/A, #VALUE!,
'   #NAME?, #CONNECT!, ...), so the wait does not depend on the number that
'   Excel gives to the busy state.
'
' UPDATED
'   2026-11-21
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim i           As Long      'Pair
    Dim Value       As Variant   'What a scratch cell holds

'------------------------------------------------------------------------------
' CHECK
'------------------------------------------------------------------------------
        For i = 1 To mPriceCount
            Value = TableCell(mPriceSheet, mPriceRows(i), mPriceKeyColumn + FX_COL_PRICE).Value2
            If IsError(Value) Then
                If Not IsFinalError(ErrorNumber(Value)) Then
                    PricesPending = True
                    Exit Function
                End If
            ElseIf IsEmpty(Value) Then
                PricesPending = True
                Exit Function
            End If
        Next i

End Function


Private Function ErrorNumber( _
    ByVal Value As Variant) _
    As Long
'
'==============================================================================
'                               ErrorNumber
'------------------------------------------------------------------------------
' PURPOSE
'   The number of an Excel error value, such as 2042 for #N/A, read from its
'   text ("Error 2042") so that no Error variant is compared (#721).
'
' UPDATED
'   2026-11-21
'==============================================================================
'

'------------------------------------------------------------------------------
' READ
'------------------------------------------------------------------------------
        ErrorNumber = CLng(Val(Mid$(CStr(Value), 7)))

End Function


Private Function IsFinalError( _
    ByVal Number As Long) _
    As Boolean
'
'==============================================================================
'                               IsFinalError
'------------------------------------------------------------------------------
' PURPOSE
'   True for an Excel error that a lookup ends with (#NULL!, #DIV/0!,
'   #VALUE!, #REF!, #NAME?, #NUM!, #N/A, #SPILL!, #CONNECT!, #BLOCKED!,
'   #UNKNOWN!, #FIELD!, #CALC!), as opposed to the busy state of a request
'   that has not come back (#721).
'
' UPDATED
'   2026-11-21
'==============================================================================
'

'------------------------------------------------------------------------------
' CHECK
'------------------------------------------------------------------------------
        Select Case Number
            Case 2000, 2007, 2015, 2023, 2029, 2036, 2042, 2045 To 2050
                IsFinalError = True
        End Select

End Function


Private Function ScratchState( _
    ByVal Value As Variant) _
    As String
'
'==============================================================================
'                               ScratchState
'------------------------------------------------------------------------------
' PURPOSE
'   Say what a scratch cell returned, for the message when no price came
'   (#721): the Excel error, or the kind of value.
'
' UPDATED
'   2026-11-21
'==============================================================================
'

'------------------------------------------------------------------------------
' BUILD
'------------------------------------------------------------------------------
        If IsError(Value) Then
            Select Case ErrorNumber(Value)
                Case 2000
                    ScratchState = "#NULL!"
                Case 2007
                    ScratchState = "#DIV/0!"
                Case 2015
                    ScratchState = "#VALUE!"
                Case 2023
                    ScratchState = "#REF!"
                Case 2029
                    ScratchState = "#NAME?"
                Case 2036
                    ScratchState = "#NUM!"
                Case 2042
                    ScratchState = "#N/A"
                Case 2045
                    ScratchState = "#SPILL!"
                Case 2046
                    ScratchState = "#CONNECT!"
                Case 2047
                    ScratchState = "#BLOCKED!"
                Case 2048
                    ScratchState = "#UNKNOWN!"
                Case 2049
                    ScratchState = "#FIELD!"
                Case 2050
                    ScratchState = "#CALC!"
                Case Else
                    ScratchState = "the busy state of a request (error " & ErrorNumber(Value) & "), the data had not come"
            End Select
        ElseIf IsEmpty(Value) Then
            ScratchState = "an empty cell"
        ElseIf VarType(Value) = vbString Then
            ScratchState = "text"
        Else
            ScratchState = "a value that is not a price"
        End If

End Function


Private Sub ClearPriceScratch( _
    ByVal ParamsSheet As Worksheet, _
    ByVal KeyColumn As Long, _
    ByVal FirstRow As Long, _
    ByVal LastRow As Long)
'
'==============================================================================
'                            ClearPriceScratch
'------------------------------------------------------------------------------
' PURPOSE
'   Empty the hidden scratch column that the price formulas of a lookup use,
'   from FirstRow to LastRow (#721), so that no formula and no data request
'   stays in the workbook.
'
' UPDATED
'   2026-11-21
'==============================================================================
'

'------------------------------------------------------------------------------
' EMPTY
'------------------------------------------------------------------------------
        If LastRow >= FirstRow Then
            TableRange(ParamsSheet, FirstRow, KeyColumn + FX_COL_PRICE, LastRow, KeyColumn + FX_COL_PRICE).ClearContents
        End If

End Sub

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
'------------------------------------------------------------------------------
'                 HIDE THE BAR WHEN SHOWN, SHOW IT WHEN HIDDEN
'------------------------------------------------------------------------------
' PURPOSE
'   Hides the navigation bar when it is shown and shows it when it is
'   hidden, on every sheet (ShowNavigationBar)
'
' WHY THIS EXISTS
'   The macro of the two triangles at the far left of B1 of every sheet
'   (#777), which can also be run from the Macros dialog (#620)
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Reads the state from the workbook (NavigationBarShown) and asks
'   ShowNavigationBar for the other one
'
' ERROR POLICY
'   As ShowNavigationBar: an error is reported in a message, never raised
'
' DEPENDENCIES
'   NavigationBarShown, ShowNavigationBar
'
' NOTES
'   tools/navigation_bar.py links the triangles to this name, as
'   [0]!ToggleNavigationBar; renaming it would break them
'
' UPDATED
'   2026-11-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "M_Main.ToggleNavigationBar"

'------------------------------------------------------------------------------
' TOGGLE
'------------------------------------------------------------------------------
    'Ask for the state the bar is not in
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
'------------------------------------------------------------------------------
'                   SHOW OR HIDE THE BAR ON EVERY WORKSHEET
'------------------------------------------------------------------------------
' PURPOSE
'   Shows or hides column A, the shape groups NAV_GROUP_TOP and
'   NAV_GROUP_REST and the triangle of B1 that hides the bar, and does the
'   opposite with the triangle that shows it, on every worksheet in one pass
'
' WHY THIS EXISTS
'   The Ribbon's Navigation bar toggle (#620), the triangles of B1 (#777)
'   and the save handlers change the bar the same way, with the layout
'   changing at once on every sheet
'
' INPUTS
'   Shown
'     True to show the bar, False to hide it
'
' RETURNS
'   The number of worksheets changed
'
' BEHAVIOR
'   - Turns screen updating off and runs ProtectWorkbook, so the macros
'     change the protected sheets through their user-interface protection
'   - On every worksheet: column A is hidden when the bar is not shown; the
'     two groups and NAV_TOGGLE_HIDE are visible when it is, and
'     NAV_TOGGLE_SHOW when it is not
'   - Restores screen updating and the workbook's Saved flag
'
' ERROR POLICY
'   Re-raises the first error after screen updating and the Saved flag are
'   restored
'
' DEPENDENCIES
'   ProtectWorkbook, CORE_Config (NAV_GROUP_TOP, NAV_GROUP_REST,
'   NAV_TOGGLE_HIDE, NAV_TOGGLE_SHOW)
'
' NOTES
'   A sheet without the triangles, as in a workbook built before #777, has
'   only its column and groups changed
'
' UPDATED
'   2026-11-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME           As String = "M_Main.ApplyNavigationBar"

    Dim Sheet                 As Worksheet    'Each worksheet
    Dim Item                  As Shape        'Each shape of a sheet
    Dim WasSaved              As Boolean      'The workbook's Saved flag before
    Dim SavedScreen           As Boolean      'Screen updating before
    Dim Changed               As Long         'Worksheets changed
    Dim ErrorNumber           As Long         'The first error's number
    Dim ErrorDescription      As String       'The first error's description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Keep the Saved flag and screen updating, and change the sheets at once
        WasSaved = ThisWorkbook.Saved
        SavedScreen = Application.ScreenUpdating
        On Error GoTo ErrorHandler
        Application.ScreenUpdating = False
    'Protect the sheets for the user interface only, so the macros can change them
        ProtectWorkbook

'------------------------------------------------------------------------------
' SHOW OR HIDE
'------------------------------------------------------------------------------
    'Column A, the two groups and the triangle that hides the bar follow the
    'state; the triangle that shows it does the opposite
        For Each Sheet In ThisWorkbook.Worksheets
            Sheet.Columns(1).Hidden = Not Shown
            For Each Item In Sheet.Shapes
                Select Case Item.Name
                    Case NAV_GROUP_TOP, NAV_GROUP_REST, NAV_TOGGLE_HIDE
                        Item.Visible = IIf(Shown, msoTrue, msoFalse)
                    Case NAV_TOGGLE_SHOW
                        Item.Visible = IIf(Shown, msoFalse, msoTrue)
                    Case Else
                End Select
            Next Item
            Changed = Changed + 1
        Next Sheet

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Restore screen updating and the Saved flag, and return the count
        Application.ScreenUpdating = SavedScreen
        ThisWorkbook.Saved = WasSaved
        ApplyNavigationBar = Changed
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
ErrorHandler:
    'Capture the error before the cleanup resets it
        ErrorNumber = Err.Number
        ErrorDescription = Err.Description
    'Restore screen updating and the Saved flag
        Application.ScreenUpdating = SavedScreen
        ThisWorkbook.Saved = WasSaved
    'Raise the first error again
        Err.Raise ErrorNumber, PROC_NAME, ErrorDescription

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
    'Every operation first withdraws both reconciliations (#45), the
    'fingerprint of the output sheets and the Dashboard (#675), before
    'anything that can stop it, the recalculation included; only a run that
    'completes makes new ones, so a run that fails leaves no Dashboard.
        mReconState = "NONE"
        mTradeReconState = "NONE"
        mReconLine = ""
        ForgetRunOutputs
        ClearDashboard
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
                    WriteDashboard
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
        ReapplyOutputFilters

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
'   completed run. The Dashboard is recalculated whole: its status, title
'   and unit formulas read the blocks the operation wrote or emptied (#675).
'
' ERROR POLICY
'   Best effort: a status that cannot be recalculated here is refreshed by
'   the next recalculation, and the outcome of the operation is unchanged.
'
' UPDATED
'   2026-11-23
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
        If SheetExists(SH_DASHBOARD) Then
            GetSheet(SH_DASHBOARD).Calculate
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
'   reconciliation=DONE; trade_reconciliation=NONE; year_convention=365;
'   calendar=TABLE; holiday_table=empty; fx_source=Manual;
'   codes=TR-MATURED; cleanup=PASS"; codes, the distinct diagnostic codes
'   of the run, comes before cleanup for a run and a validation, not for a
'   clear (#81).
'
' ERROR POLICY
'   In silent mode a cleanup failure raises ERR_CLEANUP_FAILED, with the
'   result line in its description; otherwise it is added to the message.
'
' UPDATED
'   2026-11-22
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
                     "; trade_reconciliation=" & mTradeReconState & _
                     "; year_convention=" & CORE_Engine.YearConventionText() & _
                     "; calendar=" & CORE_Engine.CalendarText() & _
                     "; holiday_table=" & CORE_Engine.HolidayTableText() & _
                     "; fx_source=" & CORE_Engine.FxSourceText()
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
'------------------------------------------------------------------------------
'                          SHOW A MESSAGE TO THE USER
'------------------------------------------------------------------------------
' PURPOSE
'   Shows a message to the user, or prints it in silent mode unless
'   gQuietLog is set
'
' WHY THIS EXISTS
'   Every message of the workbook goes through one place, so that silent
'   mode and the tests can capture it
'
' INPUTS
'   Message
'     The text to show
'   IconStyle
'     Message-box icon, for example vbInformation or vbCritical
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   - Gives the status bar back to Excel first, when an operation shows
'     its progress there (ProgressEnd, #807)
'
' ERROR POLICY
'   None of its own; MsgBox errors propagate
'
' DEPENDENCIES
'   ProgressEnd, MsgBox, gSilent, gQuietLog, gLastMessage
'
' NOTES
'   None
'
' UPDATED
'   2026-11-27
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' SHOW OR PRINT
'------------------------------------------------------------------------------
    'The message ends the operation, and its progress
        ProgressEnd
    'Remember the message for the tests, then show or print it
        gLastMessage = Message
        If gSilent Then
            If Not gQuietLog Then
                Debug.Print Message
            End If
        Else
            MsgBox Message, IconStyle, MSG_TITLE
        End If

End Sub


Private Sub ProgressStart( _
    ByVal Operation As String)
'
'------------------------------------------------------------------------------
'                     SHOW THE PROGRESS OF AN OPERATION
'------------------------------------------------------------------------------
' PURPOSE
'   Starts showing the progress of a long operation in the status bar
'   (#807), under the name Operation
'
' WHY THIS EXISTS
'   Excel freezes while a long import or report runs; the status bar
'   says what it is doing
'
' INPUTS
'   Operation
'     The operation, such as "Report" or "Import Trades"
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   - Does nothing in silent mode: automation and tests show nothing
'   - ProgressShow then writes each step, ProgressEnd gives the status bar
'     back to Excel; Notify calls ProgressEnd
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   gSilent
'
' NOTES
'   None
'
' UPDATED
'   2026-11-27
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' START
'------------------------------------------------------------------------------
    'Nothing is shown in silent mode
        If gSilent Then Exit Sub
    'Name the operation; nothing is written until its first step
        mProgressOperation = Operation
        mProgressText = vbNullString
        mProgressDone = 0
        mProgressTotal = 0

End Sub


Private Sub ProgressShow( _
    ByVal Done As Long, _
    ByVal Total As Long, _
    ByVal Detail As String)
'
'------------------------------------------------------------------------------
'                     WRITE A STEP OF AN OPERATION'S PROGRESS
'------------------------------------------------------------------------------
' PURPOSE
'   Writes the progress of the operation ProgressStart named to the status
'   bar: Done of Total steps, and what is being done (CORE_Util.ProgressText)
'
' WHY THIS EXISTS
'   The procedures of an operation say their own steps, while only the
'   interactive caller decides whether progress is shown
'
' INPUTS
'   Done
'     Steps done
'   Total
'     Steps in all; 0 for unknown, which shows no bar
'   Detail
'     What is being done
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   - Does nothing when no progress was started, so that the procedures of
'     an operation can show their steps whoever calls them
'   - Writes the status bar only when its text changes
'   - Keeps the step, for ProgressDetail
'
' ERROR POLICY
'   Best effort: a status bar that cannot be written is left as it is
'
' DEPENDENCIES
'   CORE_Util.ProgressText, Application.StatusBar
'
' NOTES
'   No DoEvents: the status bar repaints without yielding to Excel
'
' UPDATED
'   2026-11-27
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Text            As String         'The text of the status bar

'------------------------------------------------------------------------------
' WRITE
'------------------------------------------------------------------------------
    'Only an operation that shows its progress
        If VBA.Len(mProgressOperation) = 0 Then Exit Sub
    'Keep the step for ProgressDetail
        mProgressDone = Done
        mProgressTotal = Total
    'Only a text that changed
        Text = CORE_Util.ProgressText(mProgressOperation, Done, Total, Detail)
        If Text = mProgressText Then Exit Sub
        On Error Resume Next
        Application.StatusBar = Text
        On Error GoTo 0
        mProgressText = Text

End Sub


Private Sub ProgressDetail( _
    ByVal Detail As String)
'
'------------------------------------------------------------------------------
'                    SAY WHAT THE CURRENT STEP IS DOING
'------------------------------------------------------------------------------
' PURPOSE
'   Writes a new detail for the step ProgressShow wrote last, such as the
'   chart being pictured, without moving the bar (#807)
'
' WHY THIS EXISTS
'   A step can be long, as the charts of a PDF; its detail moves while
'   the bar stays
'
' INPUTS
'   Detail
'     What is being done
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   - Does nothing when no progress was started (ProgressShow)
'
' ERROR POLICY
'   As ProgressShow
'
' DEPENDENCIES
'   ProgressShow
'
' NOTES
'   None
'
' UPDATED
'   2026-11-27
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' WRITE
'------------------------------------------------------------------------------
    'The step stays; the detail changes
        ProgressShow mProgressDone, mProgressTotal, Detail

End Sub


Private Sub ProgressEnd()

'
'------------------------------------------------------------------------------
'                   GIVE THE STATUS BAR BACK TO EXCEL
'------------------------------------------------------------------------------
' PURPOSE
'   Ends the progress ProgressStart began and gives the status bar back to
'   Excel (#807)
'
' WHY THIS EXISTS
'   Excel keeps a text written to the status bar until it is given back
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   - Does nothing when no progress was started, so that a status bar
'     another procedure writes, as Download prices, is left as it is
'   - Called before every message (Notify) and every question, so that a
'     dialog never shows over a status bar that still runs
'
' ERROR POLICY
'   Best effort: a status bar that cannot be reset is left as it is
'
' DEPENDENCIES
'   Application.StatusBar
'
' NOTES
'   None
'
' UPDATED
'   2026-11-27
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' END
'------------------------------------------------------------------------------
    'Only an operation that showed its progress
        If VBA.Len(mProgressOperation) = 0 Then Exit Sub
    'Give the status bar back to Excel
        mProgressOperation = vbNullString
        mProgressText = vbNullString
        mProgressDone = 0
        mProgressTotal = 0
        On Error Resume Next
        Application.StatusBar = False
        On Error GoTo 0

End Sub
