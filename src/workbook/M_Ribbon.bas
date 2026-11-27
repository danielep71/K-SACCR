Attribute VB_Name = "M_Ribbon"
'==============================================================================
' MODULE: M_Ribbon
'------------------------------------------------------------------------------
' PURPOSE
'   Answer the callbacks of the workbook's K-SACCR Ribbon tab (#213, #215):
'   run the existing M_Main macro of a command, activate the sheet of a
'   Go to entry, and give the state of the controls. The contract is
'   docs/workbook/RIBBON.md; the RibbonX source is src/workbook/customUI14.xml.
'
' PUBLIC SURFACE
'   RibbonOnLoad: keep the Ribbon object, so that the tab can be refreshed,
'   and its pointer in a hidden workbook name, so that a VBA state reset
'   does not lose it (#216).
'   RibbonAction: run the macro mapped to a command control.
'   RibbonGoTo: activate the sheet mapped to a Go to entry.
'   RibbonGetEnabled: whether a command can be clicked, by the state table
'   of the contract (#216).
'   RibbonGetLabel: the text of a Status label, read from a status cell.
'   RibbonToggle and RibbonGetPressed: the Navigation bar toggle, which
'   shows or hides the navigation bar of every sheet, and its state, read
'   from the workbook (#620).
'   RibbonMacro, RibbonSheet, ResultsLabelFor and EnabledFor: the pure maps
'   behind them, for tests: IRibbonControl cannot be created from VBA.
'   PointerText and PointerFromText: the pure format of the stored pointer.
'   RefreshRibbon: refresh the tab; ScheduleRefresh: refresh it once Excel
'   is idle, after an operation and its recalculation. RibbonLoaded:
'   whether the Ribbon object is held or can be recovered, from the
'   Immediate window or a test (?M_Ribbon.RibbonLoaded()).
'   HidePointer and ShowPointer: take the stored pointer out of the
'   workbook before a save and put it back after, for ThisWorkbook.
'   The callbacks have the signatures Office requires and are called by
'   Excel by name. Like the M_Main macros, they are a host-only surface and
'   are not listed in docs/PUBLIC_API.txt, which lists src/modules only.
'
' DEPENDENCIES
'   M_Main for every command and its busy flag (OperationRunning);
'   CORE_Util for the sheet and named-cell lookups; CORE_Config for the
'   sheet names, the Results status cell and the staging name prefixes;
'   RtlMoveMemory of kernel32, part of Windows, to rebuild the Ribbon
'   object from its pointer. Nothing is installed or registered.
'
' STATE OWNERSHIP
'   Owns mRibbon, the Ribbon object given to RibbonOnLoad, and the hidden
'   workbook name RIBBON_POINTER_NAME, which holds its pointer and the
'   window handle of this Excel session as text (#216). An unhandled error
'   or a code edit resets VBA and clears mRibbon; CurrentRibbon then
'   rebuilds it from the pointer, but only when the window handle is this
'   session's. The name is written when Excel loads the tab, with the
'   workbook's Saved flag kept, so opening the workbook does not mark it
'   changed; HidePointer removes it before every save and ShowPointer puts
'   it back after, so the file never holds a pointer of an earlier
'   session. Also owns mInCallback, which keeps a callback from running
'   inside another, and mHeldPointer, the pointer during a save.
'
' ERROR POLICY
'   No callback raises to Excel: Excel would show its own error and stop
'   calling the callback. RibbonAction and RibbonGoTo report a failure in a
'   message box; RibbonGetLabel shows "unavailable"; RibbonGetEnabled
'   answers True, as it does when a status cannot be read or a callback
'   runs inside another: a control is never disabled on a guess. The
'   macros keep their own error handling and every gate (staging
'   validation, commit, export eligibility), whatever the Ribbon shows.
'   Without a stored pointer, or with one of another session, the Ribbon
'   is simply not refreshed until the workbook is reopened.
'
' COMPATIBILITY
'   Excel VBA; IRibbonUI and IRibbonControl come from the Microsoft Office
'   Object Library, one of the workbook's default references.
'
' UPDATED
'   2026-11-07
'
' AUTHOR
'   Daniele Penza
'==============================================================================

'------------------------------------------------------------------------------
' MODULE SETTINGS
'------------------------------------------------------------------------------
    'Require explicit declarations. Excel calls the callbacks by name, so
    'this module is not private.
    Option Explicit

'------------------------------------------------------------------------------
' WINDOWS API
'------------------------------------------------------------------------------
    'Copy bytes between two memory locations (kernel32, part of Windows):
    'CurrentRibbon uses it to turn a stored pointer back into the Ribbon
    'object (#216). PtrSafe and LongPtr serve 64-bit and 32-bit Excel alike.
        Private Declare PtrSafe Sub RtlMoveMemory Lib "kernel32" ( _
            ByRef Destination As Any, _
            ByRef Source As Any, _
            ByVal Length As LongPtr)

'------------------------------------------------------------------------------
' MODULE CONSTANTS
'------------------------------------------------------------------------------
    'Title of every message box: the project name.
        Private Const MSG_TITLE     As String = "K-SACCR"
    'The hidden workbook name that holds the Ribbon's pointer (#216).
        Private Const RIBBON_POINTER_NAME   As String = "ksRibbonPointer"
    'The id of the K-SACCR tab in customUI14.xml, selected when the workbook opens (#726).
        Private Const RIBBON_TAB_ID         As String = "ksTab"
    'Seconds after the load of the Ribbon at which the tab is selected once more (#726).
        Private Const TAB_RETRY_SECONDS     As Long = 3

'------------------------------------------------------------------------------
' MODULE STATE
'------------------------------------------------------------------------------
        Private mRibbon             As IRibbonUI    'The Ribbon, from RibbonOnLoad
        Private mInCallback         As Boolean      'A state callback is running
        Private mHeldPointer        As String       'The stored pointer while a save runs
        Private mTabAt              As Date         'When the next selection of the tab is scheduled
        Private mTabName            As String       'The procedure it was scheduled with

'------------------------------------------------------------------------------
' PUBLIC PROCEDURES
'------------------------------------------------------------------------------

Public Sub RibbonOnLoad( _
    ByVal Ribbon As IRibbonUI)
'
'==============================================================================
'                                RibbonOnLoad
'------------------------------------------------------------------------------
' PURPOSE
'   Keep the Ribbon object Excel passes when it loads the tab, and store
'   its pointer with this session's window handle in the hidden name
'   RIBBON_POINTER_NAME, without marking the workbook changed (#216), and
'   open the tab K-SACCR (#726). An add-in that loads its own tab after
'   this one (think-cell) takes the selection away, so the tab is selected
'   again once, a moment later, by ActivateRibbonTab.
'
' ERROR POLICY
'   Selecting the tab is best effort: a failure never stops the load.
'
' UPDATED
'   2026-11-21
'==============================================================================
'
        Set mRibbon = Ribbon
        WritePointer PointerText(ObjPtr(Ribbon), Application.Hwnd)
        On Error Resume Next
        Ribbon.ActivateTab RIBBON_TAB_ID
        On Error GoTo 0
        ScheduleTabActivation

End Sub


Public Sub ActivateRibbonTab()
'
'==============================================================================
'                             ActivateRibbonTab
'------------------------------------------------------------------------------
' PURPOSE
'   Select the tab K-SACCR once more, TAB_RETRY_SECONDS after the Ribbon
'   loaded, for an add-in that took the selection when its own tab loaded
'   (#726). Application.OnTime runs it, once. It selects the tab only while
'   this workbook is the active one, so that another open workbook keeps its
'   own tab. Only the load of the workbook starts it: a sheet change or a tab
'   the user chooses after that moment is left alone. Excel cannot say which
'   tab is selected, so a tab the user chose within those seconds is replaced.
'
' ERROR POLICY
'   Never raises: the tab stays where it is when it cannot be selected.
'
' UPDATED
'   2026-11-21
'==============================================================================
'
        mTabAt = 0
        On Error GoTo Done
        If mRibbon Is Nothing Then GoTo Done
        If Not Application.ActiveWorkbook Is ThisWorkbook Then GoTo Done
        mRibbon.ActivateTab RIBBON_TAB_ID

Done:

End Sub


Public Sub CancelTabActivation()
'
'==============================================================================
'                            CancelTabActivation
'------------------------------------------------------------------------------
' PURPOSE
'   Cancel the selection of the tab that is scheduled, when there is one,
'   so that Application.OnTime does not reopen a workbook that was closed
'   before it ran (#726). ThisWorkbook.Workbook_BeforeClose calls it.
'
' ERROR POLICY
'   Never raises: a selection that already ran cannot be cancelled, and that
'   is the expected case.
'
' UPDATED
'   2026-11-21
'==============================================================================
'
        On Error Resume Next
        If mTabAt <> 0 Then
            Application.OnTime mTabAt, mTabName, , False
        End If
        On Error GoTo 0
        mTabAt = 0
        mTabName = vbNullString

End Sub


Private Sub ScheduleTabActivation()
'
'==============================================================================
'                           ScheduleTabActivation
'------------------------------------------------------------------------------
' PURPOSE
'   Schedule ActivateRibbonTab TAB_RETRY_SECONDS from now. The
'   workbook name is quoted, with an apostrophe in it doubled, as Excel
'   requires (#726).
'
' ERROR POLICY
'   Never raises: when it cannot be scheduled, the tab is not selected again.
'
' UPDATED
'   2026-11-21
'==============================================================================
'
        On Error GoTo Failed
        mTabAt = Now + TimeSerial(0, 0, TAB_RETRY_SECONDS)
        mTabName = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!M_Ribbon.ActivateRibbonTab"
        Application.OnTime mTabAt, mTabName
        Exit Sub

Failed:
        mTabAt = 0
        mTabName = vbNullString

End Sub


Public Sub RibbonAction( _
    ByVal Control As IRibbonControl)
'
'------------------------------------------------------------------------------
'                      RUN THE MACRO OF A COMMAND CONTROL
'------------------------------------------------------------------------------
' PURPOSE
'   Runs the M_Main macro mapped to a command control of the K-SACCR tab
'   (RibbonMacro), then schedules a refresh of the tab
'
' WHY THIS EXISTS
'   The Ribbon reaches the existing macros and holds no business logic:
'   each macro keeps its protection, busy guard, error handling and every
'   gate (docs/workbook/RIBBON.md). The refresh lets the Status labels and
'   the enabled states follow the operation
'
' INPUTS
'   Control
'     The control Excel passes to the onAction callback
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   - Finds the macro mapped to the control (RibbonMacro)
'   - Calls it by name, one Case per macro, so that the compiler checks
'     every call and tools/test_ribbon_contract.py can read the calls
'   - Schedules the refresh of the tab once Excel is idle (ScheduleRefresh)
'
' ERROR POLICY
'   The macros report their own outcome and never raise. A control mapped
'   to no macro, a mapped macro this procedure does not call, and an error
'   are reported in a message box, and nothing runs. Never raises to Excel
'
' DEPENDENCIES
'   RibbonMacro, ScheduleRefresh; the M_Main macros of the command table
'   of docs/workbook/RIBBON.md
'
' NOTES
'   Export report (#259) runs M_Main.ExportReport, the macro of the
'   formatted run report
'
' UPDATED
'   2026-11-26
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "M_Ribbon.RibbonAction"

    Dim Macro           As String    'The macro mapped to the control

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Report any error in a message box, never to Excel
        On Error GoTo Failed

'------------------------------------------------------------------------------
' RUN THE MACRO
'------------------------------------------------------------------------------
    'Find the macro mapped to the control
        Macro = RibbonMacro(Control.Id)
    'Call it by name
        Select Case Macro
            Case "RunSACCR"
                M_Main.RunSACCR
            Case "ValidateInputs"
                M_Main.ValidateInputs
            Case "ClearOutputs"
                M_Main.ClearOutputs
            Case "ImportNettingSetsCsv"
                M_Main.ImportNettingSetsCsv
            Case "ImportTradesCsv"
                M_Main.ImportTradesCsv
            Case "ImportBankNettingSetsCsv"
                M_Main.ImportBankNettingSetsCsv
            Case "ImportBankTradesCsv"
                M_Main.ImportBankTradesCsv
            Case "ValidateStaging"
                M_Main.ValidateStaging
            Case "CommitImport"
                M_Main.CommitImport
            Case "ExportNettingSetsCsvTemplate"
                M_Main.ExportNettingSetsCsvTemplate
            Case "ExportTradesCsvTemplate"
                M_Main.ExportTradesCsvTemplate
            Case "ExportBankNettingSetsCsvTemplate"
                M_Main.ExportBankNettingSetsCsvTemplate
            Case "ExportBankTradesCsvTemplate"
                M_Main.ExportBankTradesCsvTemplate
            Case "ExportImportSpecification"
                M_Main.ExportImportSpecification
            Case "ExportEvidence"
                M_Main.ExportEvidence
            Case "ExportReport"
                M_Main.ExportReport
            Case "RestoreRegulatoryParameters"
                M_Main.RestoreRegulatoryParameters
            Case vbNullString
                VBA.MsgBox "The Ribbon control " & Control.Id & " runs no command.", vbExclamation, MSG_TITLE
            Case Else
                'Mapped, but not a macro this module calls: a contract error
                VBA.MsgBox "This command is not available in this workbook.", vbInformation, MSG_TITLE
        End Select

'------------------------------------------------------------------------------
' REFRESH THE TAB
'------------------------------------------------------------------------------
    'Let the Status labels and the enabled states follow the operation
        ScheduleRefresh

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Leave before the error handler
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Failed:
    'Report the failure in a message box
        VBA.MsgBox "The Ribbon command could not be run: " & Err.Description, vbExclamation, MSG_TITLE

End Sub


Public Sub RibbonGoTo( _
    ByVal Control As IRibbonControl)
'
'==============================================================================
'                                 RibbonGoTo
'------------------------------------------------------------------------------
' PURPOSE
'   Activate the sheet of this workbook whose CodeName RibbonSheet maps to
'   a Go to entry, so that a renamed sheet tab cannot break the entry.
'
' ERROR POLICY
'   A sheet this workbook does not have is reported in a message box, and
'   nothing changes; another workbook is never touched.
'
' UPDATED
'   2026-11-05
'==============================================================================
'
    Dim Target      As String       'CodeName of the sheet
    Dim Sheet       As Worksheet    'A sheet of this workbook
    Dim Found       As Worksheet    'The sheet with that CodeName

        On Error GoTo Failed
        Target = RibbonSheet(Control.Id)
        If Len(Target) > 0 Then
            For Each Sheet In ThisWorkbook.Worksheets
                If StrComp(Sheet.CodeName, Target, vbTextCompare) = 0 Then
                    Set Found = Sheet
                    Exit For
                End If
            Next Sheet
        End If
        If Found Is Nothing Then
            MsgBox "This workbook has no sheet for the Ribbon entry " & Control.Id & ".", vbExclamation, MSG_TITLE
        Else
            Found.Activate
        End If
        Exit Sub

Failed:
        MsgBox "The sheet could not be opened: " & Err.Description, vbExclamation, MSG_TITLE

End Sub


Public Sub RibbonGetEnabled( _
    ByVal Control As IRibbonControl, _
    ByRef ReturnedVal As Variant)
'
'==============================================================================
'                              RibbonGetEnabled
'------------------------------------------------------------------------------
' PURPOSE
'   Whether a command can be clicked, by the state table of
'   docs/workbook/RIBBON.md (EnabledFor, #216). It reads only the busy flag
'   of M_Main, Results C5 and the two staging status cells; it never
'   calculates or writes. The macros refuse what they must in any case.
'
' ERROR POLICY
'   Never raises. A status that cannot be read, an error, or a call while
'   another state callback runs gives True.
'
' UPDATED
'   2026-11-06
'==============================================================================
'
        If mInCallback Then
            ReturnedVal = True
            Exit Sub
        End If
        mInCallback = True
        On Error GoTo Failed
        ReturnedVal = EnabledFor(Control.Id, M_Main.OperationRunning(), ResultsLabel(), _
                                 StagingLabel(IMPORT_NS_PREFIX), StagingLabel(IMPORT_TR_PREFIX))
        mInCallback = False
        Exit Sub

Failed:
        ReturnedVal = True
        mInCallback = False

End Sub


Public Sub RibbonGetLabel( _
    ByVal Control As IRibbonControl, _
    ByRef ReturnedVal As Variant)
'
'==============================================================================
'                               RibbonGetLabel
'------------------------------------------------------------------------------
' PURPOSE
'   The text of a Status label. It reads only a status cell: Results C5
'   for the results, D7 of each staging sheet through its name; it never
'   calculates or writes.
'
' ERROR POLICY
'   A staging status name that is missing gives "unavailable" after the
'   label's prefix; an error, or a call while another state callback runs,
'   gives "Status unavailable"; nothing is raised.
'
' UPDATED
'   2026-11-09
'==============================================================================
'
        If mInCallback Then
            ReturnedVal = "Status unavailable"
            Exit Sub
        End If
        mInCallback = True
        On Error GoTo Failed
        Select Case Control.Id
            Case "ksStatusResults"
                ReturnedVal = "Results: " & ResultsLabel()
            Case "ksStatusNettingSets"
                ReturnedVal = "NettingSets staging: " & StagingLabel(IMPORT_NS_PREFIX)
            Case "ksStatusTrades"
                ReturnedVal = "Trades staging: " & StagingLabel(IMPORT_TR_PREFIX)
            Case Else
                ReturnedVal = vbNullString
        End Select
        mInCallback = False
        Exit Sub

Failed:
        ReturnedVal = "Status unavailable"
        mInCallback = False

End Sub


Public Sub RibbonToggle( _
    ByVal Control As IRibbonControl, _
    ByVal Pressed As Boolean)
'
'==============================================================================
'                                RibbonToggle
'------------------------------------------------------------------------------
' PURPOSE
'   Answer a toggle button: Navigation bar (ksNavigationBar) shows the
'   navigation bar of every sheet when pressed and hides it when released
'   (M_Main.ShowNavigationBar, #620), which then refreshes the tab, so
'   that the toggle shows the state the workbook holds.
'
' ERROR POLICY
'   M_Main.ShowNavigationBar reports its own failure. An unknown toggle is
'   reported in a message box, and nothing changes.
'
' UPDATED
'   2026-11-07
'==============================================================================
'
        On Error GoTo Failed
        Select Case Control.Id
            Case "ksNavigationBar"
                M_Main.ShowNavigationBar Pressed
            Case Else
                MsgBox "The Ribbon toggle " & Control.Id & " changes nothing.", vbExclamation, MSG_TITLE
                ScheduleRefresh
        End Select
        Exit Sub

Failed:
        MsgBox "The Ribbon toggle could not be used: " & Err.Description, vbExclamation, MSG_TITLE

End Sub


Public Sub RibbonGetPressed( _
    ByVal Control As IRibbonControl, _
    ByRef ReturnedVal As Variant)
'
'==============================================================================
'                              RibbonGetPressed
'------------------------------------------------------------------------------
' PURPOSE
'   Whether a toggle button is pressed: Navigation bar is pressed while the
'   navigation bar is shown, as the workbook holds it
'   (M_Main.NavigationBarShown, #620), so the state survives a VBA state
'   reset. It never writes.
'
' ERROR POLICY
'   Never raises; an unknown toggle, or a state that cannot be read, gives
'   True.
'
' UPDATED
'   2026-11-07
'==============================================================================
'
        On Error GoTo Failed
        Select Case Control.Id
            Case "ksNavigationBar"
                ReturnedVal = M_Main.NavigationBarShown()
            Case Else
                ReturnedVal = True
        End Select
        Exit Sub

Failed:
        ReturnedVal = True

End Sub


Public Function RibbonMacro( _
    ByVal ControlId As String) _
    As String
'
'------------------------------------------------------------------------------
'                        FIND THE MACRO OF A COMMAND
'------------------------------------------------------------------------------
' PURPOSE
'   Gives the M_Main macro of a command control, as the command table of
'   docs/workbook/RIBBON.md gives it
'
' WHY THIS EXISTS
'   RibbonAction runs it, and tests check it: IRibbonControl cannot be
'   created from VBA, so the map is a pure function.
'   tools/test_ribbon_contract.py checks it against the contract
'
' INPUTS
'   ControlId
'     The id of a control of src/workbook/customUI14.xml
'
' RETURNS
'   The name of the macro; "" for any other control
'
' BEHAVIOR
'   One Case per control, in the order of the command table
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   None
'
' NOTES
'   ksExportReport runs ExportReport, the formatted run report (#259)
'
' UPDATED
'   2026-11-26
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' MAP THE CONTROL
'------------------------------------------------------------------------------
    'The macro of each command control; none for any other control
        Select Case ControlId
            Case "ksRun"
                RibbonMacro = "RunSACCR"
            Case "ksRunMenuRun"
                RibbonMacro = "RunSACCR"
            Case "ksValidateInputs"
                RibbonMacro = "ValidateInputs"
            Case "ksClearOutputs"
                RibbonMacro = "ClearOutputs"
            Case "ksImportNettingSets"
                RibbonMacro = "ImportNettingSetsCsv"
            Case "ksImportTrades"
                RibbonMacro = "ImportTradesCsv"
            Case "ksImportBankNettingSets"
                RibbonMacro = "ImportBankNettingSetsCsv"
            Case "ksImportBankTrades"
                RibbonMacro = "ImportBankTradesCsv"
            Case "ksValidateStaging"
                RibbonMacro = "ValidateStaging"
            Case "ksCommitImport"
                RibbonMacro = "CommitImport"
            Case "ksTemplateNettingSets"
                RibbonMacro = "ExportNettingSetsCsvTemplate"
            Case "ksTemplateTrades"
                RibbonMacro = "ExportTradesCsvTemplate"
            Case "ksTemplateBankNettingSets"
                RibbonMacro = "ExportBankNettingSetsCsvTemplate"
            Case "ksTemplateBankTrades"
                RibbonMacro = "ExportBankTradesCsvTemplate"
            Case "ksImportSpecification"
                RibbonMacro = "ExportImportSpecification"
            Case "ksExportEvidence"
                RibbonMacro = "ExportEvidence"
            Case "ksExportReport"
                RibbonMacro = "ExportReport"
            Case "ksRestoreParameters"
                RibbonMacro = "RestoreRegulatoryParameters"
            Case Else
                RibbonMacro = vbNullString
        End Select

End Function


Public Function RibbonSheet( _
    ByVal ControlId As String) _
    As String
'
'==============================================================================
'                                 RibbonSheet
'------------------------------------------------------------------------------
' PURPOSE
'   The CodeName of the sheet a Go to entry opens, as the navigation table
'   of docs/workbook/RIBBON.md gives it; "" for any other control.
'   tools/test_ribbon_contract.py checks this map against the contract.
'
' UPDATED
'   2026-11-05
'==============================================================================
'
        Select Case ControlId
            Case "ksGo"
                RibbonSheet = "shHome"
            Case "ksGoHome"
                RibbonSheet = "shHome"
            Case "ksGoReadme"
                RibbonSheet = "shReadme"
            Case "ksGoParams"
                RibbonSheet = "shParams"
            Case "ksGoNettingSets"
                RibbonSheet = "shNettingSets"
            Case "ksGoTrades"
                RibbonSheet = "shTrades"
            Case "ksGoImportNettingSets"
                RibbonSheet = "shImportNettingSets"
            Case "ksGoImportTrades"
                RibbonSheet = "shImportTrades"
            Case "ksGoDashboard"
                RibbonSheet = "shDashboard"
            Case "ksGoResults"
                RibbonSheet = "shResults"
            Case "ksGoHedgingSets"
                RibbonSheet = "shHedgingSets"
            Case "ksGoBuckets"
                RibbonSheet = "shBuckets"
            Case "ksGoTradeCalc"
                RibbonSheet = "shTradeCalc"
            Case "ksGoChecks"
                RibbonSheet = "shChecks"
            Case "ksGoBankNettingSets"
                RibbonSheet = "shBankNettingSets"
            Case "ksGoReconciliation"
                RibbonSheet = "shReconciliation"
            Case "ksGoBankTrades"
                RibbonSheet = "shBankTrades"
            Case "ksGoTradeReconciliation"
                RibbonSheet = "shTradeReconciliation"
            Case "ksGoReconSummary"
                RibbonSheet = "shReconSummary"
            Case "ksGoTestCatalogue"
                RibbonSheet = "shTestCatalogue"
            Case "ksGoFormulaTest"
                RibbonSheet = "shTests"
            Case "ksGoBaselVsCRR"
                RibbonSheet = "shBaselVsCRR"
            Case Else
                RibbonSheet = vbNullString
        End Select

End Function


Public Function ResultsLabelFor( _
    ByVal StatusText As String) _
    As String
'
'==============================================================================
'                               ResultsLabelFor
'------------------------------------------------------------------------------
' PURPOSE
'   The Results status word for the Ribbon from the text of Results C5,
'   which M_Main.ResultsStatusText writes: CURRENT, OUT OF DATE or NONE;
'   UNKNOWN for any other text.
'
' UPDATED
'   2026-11-05
'==============================================================================
'
        If Left$(StatusText, 16) = "Inputs unchanged" Then
            ResultsLabelFor = "CURRENT"
        ElseIf Left$(StatusText, 11) = "OUT OF DATE" Then
            ResultsLabelFor = "OUT OF DATE"
        ElseIf Left$(StatusText, 10) = "No results" Then
            ResultsLabelFor = "NONE"
        Else
            ResultsLabelFor = "UNKNOWN"
        End If

End Function


Public Function EnabledFor( _
    ByVal ControlId As String, _
    ByVal Busy As Boolean, _
    ByVal ResultsState As String, _
    ByVal NettingSetsStatus As String, _
    ByVal TradesStatus As String) _
    As Boolean
'
'------------------------------------------------------------------------------
'                     DECIDE WHETHER A CONTROL CAN BE CLICKED
'------------------------------------------------------------------------------
' PURPOSE
'   Gives the state table of docs/workbook/RIBBON.md (#216): whether a
'   control can be clicked, from the busy flag, the Results state and the
'   two staging statuses
'
' WHY THIS EXISTS
'   RibbonGetEnabled answers Excel with it, and TEST_MainState checks it:
'   IRibbonControl cannot be created from VBA, so the table is a pure
'   function. The states help the user; the macros enforce the same rules
'   again, whatever the Ribbon shows
'
' INPUTS
'   ControlId
'     The id of a control of src/workbook/customUI14.xml
'
'   Busy
'     True while an operation of M_Main runs (OperationRunning)
'
'   ResultsState
'     The Results state as ResultsLabelFor gives it: CURRENT, OUT OF DATE,
'     NONE or UNKNOWN
'
'   NettingSetsStatus
'     The status of Import_NettingSets, or "unavailable"
'
'   TradesStatus
'     The status of Import_Trades, or "unavailable"
'
' RETURNS
'   True when the control can be clicked
'
' BEHAVIOR
'   - While an operation runs, no control is enabled
'   - Validate staging needs a staging sheet that holds rows: a status
'     other than EMPTY
'   - Commit import needs a staging sheet that reads VALID
'   - Export evidence and Export report (#259) need Results CURRENT: not
'     OUT OF DATE and not NONE
'   - Every other control, and an unknown one, is enabled
'
' ERROR POLICY
'   None. A status that could not be read ("unavailable", "UNKNOWN")
'   enables the control: a control is never disabled on a guess, and the
'   macro decides
'
' DEPENDENCIES
'   CORE_Config.STAGING_EMPTY, CORE_Config.STAGING_VALID
'
' NOTES
'   The two exports refuse in the same cases (M_Main.EvidenceProblem), so
'   they share one rule
'
' UPDATED
'   2026-11-26
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' WHILE AN OPERATION RUNS
'------------------------------------------------------------------------------
    'Enable nothing while an operation runs
        If Busy Then
            EnabledFor = False
            Exit Function
        End If

'------------------------------------------------------------------------------
' APPLY THE STATE TABLE
'------------------------------------------------------------------------------
    'Decide by the control
        Select Case ControlId
            Case "ksValidateStaging"
                EnabledFor = NettingSetsStatus <> STAGING_EMPTY Or TradesStatus <> STAGING_EMPTY
            Case "ksCommitImport"
                EnabledFor = NettingSetsStatus = STAGING_VALID Or TradesStatus = STAGING_VALID Or _
                             NettingSetsStatus = "unavailable" Or TradesStatus = "unavailable"
            Case "ksExportEvidence", "ksExportReport"
                EnabledFor = ResultsState <> "OUT OF DATE" And ResultsState <> "NONE"
            Case Else
                EnabledFor = True
        End Select

End Function


Public Function PointerText( _
    ByVal Pointer As LongPtr, _
    ByVal WindowHandle As Long) _
    As String
'
'==============================================================================
'                                 PointerText
'------------------------------------------------------------------------------
' PURPOSE
'   The text RIBBON_POINTER_NAME holds (#216): the Ribbon's pointer and
'   the window handle of the Excel session that loaded it, as
'   "pointer|handle" in decimal.
'
' UPDATED
'   2026-11-06
'==============================================================================
'
        PointerText = CStr(Pointer) & "|" & CStr(WindowHandle)

End Function


Public Function PointerFromText( _
    ByVal Text As String, _
    ByVal WindowHandle As Long) _
    As LongPtr
'
'==============================================================================
'                               PointerFromText
'------------------------------------------------------------------------------
' PURPOSE
'   The pointer of a text PointerText wrote, when it was written in the
'   Excel session with WindowHandle (#216).
'
' RETURNS
'   0 when the text is not two decimal integers joined by "|", the handle
'   is another session's, or the pointer is 0: no pointer is ever used
'   unless it is this session's. A minus sign is accepted: on 32-bit
'   Excel an address with its high bit set is a negative LongPtr, which
'   PointerText writes so.
'
' UPDATED
'   2026-11-06
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Parts   As Variant    'Pointer and handle

'------------------------------------------------------------------------------
' PARSE
'------------------------------------------------------------------------------
        Parts = Split(Text, "|")
        If UBound(Parts) <> 1 Then
            Exit Function
        End If
        If Not IsInteger(CStr(Parts(0))) Or Not IsInteger(CStr(Parts(1))) Then
            Exit Function
        End If
        If CStr(Parts(1)) <> CStr(WindowHandle) Then
            Exit Function
        End If
        PointerFromText = CLngPtr(Val(Parts(0)))

End Function

Public Sub RefreshRibbon()

'
'------------------------------------------------------------------------------
'                           REFRESH THE K-SACCR TAB
'------------------------------------------------------------------------------
' PURPOSE
'   Asks Excel to call the state callbacks of the K-SACCR tab again, while
'   this workbook is the active one
'
' WHY THIS EXISTS
'   The tab's buttons are enabled or disabled from the workbook's state
'   (#216). Excel looks for the callbacks, RibbonGetEnabled among them, in
'   the active workbook: invalidated while another workbook is active, as
'   a refresh scheduled with OnTime can be, Excel stops on "Cannot run the
'   macro 'RibbonGetEnabled'" and stays busy until the message is closed
'   (#773)
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   - Holds or recovers the Ribbon object (CurrentRibbon), and does nothing
'     without it: tab not loaded, or no pointer of this session
'   - Stores the pointer again when a cancelled save left it out
'   - Invalidates the Ribbon only while this workbook is active; otherwise
'     leaves it, as Workbook_Activate refreshes it when the workbook is
'     active again
'
' ERROR POLICY
'   Never raises: an error ends the refresh, which the next one repeats
'
' DEPENDENCIES
'   CurrentRibbon, StoredPointerText, WritePointer, PointerText;
'   ThisWorkbook.Workbook_Activate and Workbook_SheetActivate call it
'
' NOTES
'   ActivateRibbonTab keeps to the active workbook in the same way
'
' UPDATED
'   2026-11-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME   As String = "M_Ribbon.RefreshRibbon"

    Dim Ribbon        As IRibbonUI    'The Ribbon, held or recovered

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'End the refresh on any error
        On Error GoTo Failed

'------------------------------------------------------------------------------
' FIND THE RIBBON
'------------------------------------------------------------------------------
    'Leave without the Ribbon object
        Set Ribbon = CurrentRibbon()
        If Ribbon Is Nothing Then Exit Sub
    'Store the pointer again when a cancelled save left it out
        If Len(StoredPointerText()) = 0 Then
            WritePointer PointerText(ObjPtr(Ribbon), Application.Hwnd)
        End If

'------------------------------------------------------------------------------
' REFRESH
'------------------------------------------------------------------------------
    'Invalidate only while this workbook is active, where Excel finds the
    'callbacks
        If Not Application.ActiveWorkbook Is ThisWorkbook Then Exit Sub
        Ribbon.Invalidate

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Leave before the error handler
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Failed:
    'Leave the refresh to the next one
        Exit Sub

End Sub

Public Sub ScheduleRefresh()
'
'==============================================================================
'                               ScheduleRefresh
'------------------------------------------------------------------------------
' PURPOSE
'   Refresh the tab once Excel is idle. Excel asks for the labels as soon as
'   the tab is invalidated; inside an operation, before Excel has
'   recalculated the status formula in Results C5, the labels would show
'   the state before the operation. Application.OnTime runs RefreshRibbon
'   of this workbook when the operation and its recalculation are over. The
'   workbook name is quoted, with an apostrophe in it doubled, as Excel
'   requires in a workbook-qualified macro name ('O''Brien.xlsm'!...).
'
' ERROR POLICY
'   Never raises; if the refresh cannot be scheduled, the tab is refreshed
'   at once.
'
' UPDATED
'   2026-11-06
'==============================================================================
'
        On Error GoTo Immediate
        Application.OnTime Now, "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!M_Ribbon.RefreshRibbon"
        Exit Sub

Immediate:
        RefreshRibbon

End Sub


Public Function RibbonLoaded() As Boolean
'
'==============================================================================
'                                RibbonLoaded
'------------------------------------------------------------------------------
' PURPOSE
'   Whether this session holds the Ribbon object or can recover it, so
'   that RefreshRibbon can refresh the tab. False when Excel loaded the
'   tab before macros were enabled.
'
' UPDATED
'   2026-11-06
'==============================================================================
'
        On Error Resume Next
        RibbonLoaded = Not CurrentRibbon() Is Nothing

End Function


Public Sub HidePointer()
'
'==============================================================================
'                                 HidePointer
'------------------------------------------------------------------------------
' PURPOSE
'   Before a save, keep the stored pointer in memory and remove the name
'   RIBBON_POINTER_NAME, so that the file never holds a pointer (#216).
'   ThisWorkbook.Workbook_BeforeSave calls it. The Saved flag is kept.
'
' ERROR POLICY
'   Never raises.
'
' UPDATED
'   2026-11-06
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim WasSaved    As Boolean    'The workbook's Saved flag before

'------------------------------------------------------------------------------
' REMOVE
'------------------------------------------------------------------------------
        On Error GoTo Failed
        mHeldPointer = StoredPointerText()
        If Len(mHeldPointer) > 0 Then
            WasSaved = ThisWorkbook.Saved
            ThisWorkbook.Names(RIBBON_POINTER_NAME).Delete
            ThisWorkbook.Saved = WasSaved
        End If

Failed:

End Sub


Public Sub ShowPointer()
'
'==============================================================================
'                                 ShowPointer
'------------------------------------------------------------------------------
' PURPOSE
'   After a save, put back the pointer HidePointer removed, or the one of
'   the Ribbon object still held (#216). ThisWorkbook.Workbook_AfterSave
'   calls it. The Saved flag is kept, so a saved workbook stays saved.
'
' ERROR POLICY
'   Never raises.
'
' UPDATED
'   2026-11-06
'==============================================================================
'
        On Error GoTo Failed
        If Not mRibbon Is Nothing Then
            WritePointer PointerText(ObjPtr(mRibbon), Application.Hwnd)
        ElseIf Len(mHeldPointer) > 0 Then
            WritePointer mHeldPointer
        End If
        mHeldPointer = vbNullString

Failed:

End Sub

'------------------------------------------------------------------------------
' PRIVATE PROCEDURES
'------------------------------------------------------------------------------

Private Function StagingLabel( _
    ByVal Prefix As String) _
    As String
'
'==============================================================================
'                                StagingLabel
'------------------------------------------------------------------------------
' PURPOSE
'   The status of a staging sheet, as its named status cell (D7) shows it.
'
' INPUTS
'   Prefix: IMPORT_NS_PREFIX or IMPORT_TR_PREFIX.
'
' UPDATED
'   2026-11-05
'==============================================================================
'
    Dim StatusCell  As Range    'The staging sheet's status cell

        Set StatusCell = CORE_Util.NamedCell(Prefix & "_Status")
        If StatusCell Is Nothing Then
            StagingLabel = "unavailable"
        Else
            StagingLabel = CORE_Util.SafeStr(StatusCell.Value2)
        End If

End Function


Private Function ResultsLabel() As String
'
'==============================================================================
'                                 ResultsLabel
'------------------------------------------------------------------------------
' PURPOSE
'   The Results state as ResultsLabelFor gives it, from Results C5.
'
' UPDATED
'   2026-11-06
'==============================================================================
'
        ResultsLabel = ResultsLabelFor(CORE_Util.SafeStr(CORE_Util.GetSheet(SH_RESULTS).Range(RESULTS_STATUS_CELL).Value2))

End Function


Private Function CurrentRibbon() As IRibbonUI
'
'==============================================================================
'                                CurrentRibbon
'------------------------------------------------------------------------------
' PURPOSE
'   The Ribbon object: the one RibbonOnLoad kept or, after a VBA state
'   reset, the one rebuilt from the pointer RIBBON_POINTER_NAME holds,
'   when that pointer is this session's (#216).
'
' RETURNS
'   Nothing when there is neither.
'
' UPDATED
'   2026-11-06
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Pointer     As LongPtr    'The stored pointer; 0 for none
    Dim Recovered   As Object     'The object at that pointer

'------------------------------------------------------------------------------
' RECOVER
'------------------------------------------------------------------------------
        If mRibbon Is Nothing Then
            Pointer = PointerFromText(StoredPointerText(), Application.Hwnd)
            If Pointer <> 0 Then
                Set Recovered = ObjectAt(Pointer)
                Set mRibbon = Recovered
            End If
        End If
        Set CurrentRibbon = mRibbon

End Function


Private Function ObjectAt( _
    ByVal Pointer As LongPtr) _
    As Object
'
'==============================================================================
'                                   ObjectAt
'------------------------------------------------------------------------------
' PURPOSE
'   The object at a pointer: the pointer is copied into an object
'   variable, which is then set to the result, taking a reference of its
'   own, and cleared by copying a zero back, so that the borrowed copy is
'   never released (#216).
'
' UPDATED
'   2026-11-06
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Borrowed    As Object     'The pointer as an object, not counted
    Dim Zero        As LongPtr    'Clears Borrowed without a release

'------------------------------------------------------------------------------
' COPY
'------------------------------------------------------------------------------
        RtlMoveMemory Borrowed, Pointer, LenB(Pointer)
        Set ObjectAt = Borrowed
        RtlMoveMemory Borrowed, Zero, LenB(Zero)

End Function


Private Function StoredPointerText() As String
'
'==============================================================================
'                              StoredPointerText
'------------------------------------------------------------------------------
' PURPOSE
'   The text RIBBON_POINTER_NAME holds; "" when the name is missing.
'
' UPDATED
'   2026-11-06
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Formula     As String    'The name's RefersTo: ="text"

'------------------------------------------------------------------------------
' READ
'------------------------------------------------------------------------------
        On Error GoTo Missing
        Formula = ThisWorkbook.Names(RIBBON_POINTER_NAME).RefersTo
        If Left$(Formula, 2) = "=""" And Right$(Formula, 1) = """" Then
            StoredPointerText = Mid$(Formula, 3, Len(Formula) - 3)
        End If

Missing:

End Function


Private Sub WritePointer( _
    ByVal Text As String)
'
'==============================================================================
'                                 WritePointer
'------------------------------------------------------------------------------
' PURPOSE
'   Store a pointer text in the hidden name RIBBON_POINTER_NAME, keeping
'   the workbook's Saved flag, so that storing it never asks the user to
'   save (#216).
'
' ERROR POLICY
'   Never raises; a pointer that cannot be stored is simply not recovered.
'
' UPDATED
'   2026-11-06
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim WasSaved    As Boolean    'The workbook's Saved flag before

'------------------------------------------------------------------------------
' STORE
'------------------------------------------------------------------------------
        On Error GoTo Failed
        WasSaved = ThisWorkbook.Saved
        ThisWorkbook.Names.Add Name:=RIBBON_POINTER_NAME, RefersTo:="=""" & Text & """", Visible:=False
        ThisWorkbook.Saved = WasSaved

Failed:

End Sub


Private Function IsInteger( _
    ByVal Text As String) _
    As Boolean
'
'==============================================================================
'                                  IsInteger
'------------------------------------------------------------------------------
' PURPOSE
'   Whether Text is a decimal integer as CStr writes one: one or more
'   digits, after a minus sign or not, and nothing else.
'
' UPDATED
'   2026-11-06
'==============================================================================
'
        If Left$(Text, 1) = "-" Then
            Text = Mid$(Text, 2)
        End If
        IsInteger = Len(Text) > 0 And Not Text Like "*[!0-9]*"

End Function
