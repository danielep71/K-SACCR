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

'------------------------------------------------------------------------------
' MODULE STATE
'------------------------------------------------------------------------------
        Private mRibbon             As IRibbonUI    'The Ribbon, from RibbonOnLoad
        Private mInCallback         As Boolean      'A state callback is running
        Private mHeldPointer        As String       'The stored pointer while a save runs

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
'   RIBBON_POINTER_NAME, without marking the workbook changed (#216).
'
' UPDATED
'   2026-11-06
'==============================================================================
'
        Set mRibbon = Ribbon
        WritePointer PointerText(ObjPtr(Ribbon), Application.Hwnd)

End Sub


Public Sub RibbonAction( _
    ByVal Control As IRibbonControl)
'
'==============================================================================
'                                RibbonAction
'------------------------------------------------------------------------------
' PURPOSE
'   Run the M_Main macro mapped to a command control (RibbonMacro), then
'   schedule a refresh of the tab so that the Status labels follow.
'
' ERROR POLICY
'   The macros report their own outcome and never raise. An unknown control
'   or a mapped macro this procedure does not call is reported in a
'   message box, and nothing runs.
'
' UPDATED
'   2026-11-06
'==============================================================================
'
    Dim Macro       As String    'The macro mapped to the control

        On Error GoTo Failed
        Macro = RibbonMacro(Control.Id)
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
            Case "ValidateStaging"
                M_Main.ValidateStaging
            Case "CommitImport"
                M_Main.CommitImport
            Case "ExportNettingSetsCsvTemplate"
                M_Main.ExportNettingSetsCsvTemplate
            Case "ExportTradesCsvTemplate"
                M_Main.ExportTradesCsvTemplate
            Case "ExportImportSpecification"
                M_Main.ExportImportSpecification
            Case "ExportEvidence"
                M_Main.ExportEvidence
            Case "RestoreRegulatoryParameters"
                M_Main.RestoreRegulatoryParameters
            Case ""
                MsgBox "The Ribbon control " & Control.Id & " runs no command.", vbExclamation, MSG_TITLE
            Case Else
                'Mapped, but not a macro this module calls: a contract error.
                MsgBox "This command is not available in this workbook.", vbInformation, MSG_TITLE
        End Select
        ScheduleRefresh
        Exit Sub

Failed:
        MsgBox "The Ribbon command could not be run: " & Err.Description, vbExclamation, MSG_TITLE

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
'   A cell that cannot be read gives "unavailable", and a call while
'   another state callback runs "Status unavailable"; nothing is raised.
'
' UPDATED
'   2026-11-06
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
'   (M_Main.ShowNavigationBar, #620). The tab is then refreshed, so that
'   the toggle shows the state the workbook holds.
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
        End Select
        ScheduleRefresh
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
'==============================================================================
'                                 RibbonMacro
'------------------------------------------------------------------------------
' PURPOSE
'   The M_Main macro of a command control, as the command table of
'   docs/workbook/RIBBON.md gives it; "" for any other control.
'   tools/test_ribbon_contract.py checks this map against the contract.
'
' UPDATED
'   2026-11-05
'==============================================================================
'
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
            Case "ksValidateStaging"
                RibbonMacro = "ValidateStaging"
            Case "ksCommitImport"
                RibbonMacro = "CommitImport"
            Case "ksTemplateNettingSets"
                RibbonMacro = "ExportNettingSetsCsvTemplate"
            Case "ksTemplateTrades"
                RibbonMacro = "ExportTradesCsvTemplate"
            Case "ksImportSpecification"
                RibbonMacro = "ExportImportSpecification"
            Case "ksExportEvidence"
                RibbonMacro = "ExportEvidence"
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
'==============================================================================
'                                  EnabledFor
'------------------------------------------------------------------------------
' PURPOSE
'   The state table of docs/workbook/RIBBON.md (#216): whether a control
'   can be clicked, from the busy flag, the Results state as
'   ResultsLabelFor gives it and the two staging statuses.
'
' RULES
'   While an operation runs, no control is enabled. Validate staging needs
'   a staging sheet that holds rows: a status other than EMPTY. Commit
'   import needs a staging sheet that reads VALID. Export evidence needs
'   Results CURRENT. A status that could not be read ("unavailable",
'   "UNKNOWN") enables the control: the macro decides. Every other
'   control, and an unknown one, is enabled.
'
' UPDATED
'   2026-11-06
'==============================================================================
'
        If Busy Then
            EnabledFor = False
            Exit Function
        End If
        Select Case ControlId
            Case "ksValidateStaging"
                EnabledFor = NettingSetsStatus <> STAGING_EMPTY Or TradesStatus <> STAGING_EMPTY
            Case "ksCommitImport"
                EnabledFor = NettingSetsStatus = STAGING_VALID Or TradesStatus = STAGING_VALID Or _
                             NettingSetsStatus = "unavailable" Or TradesStatus = "unavailable"
            Case "ksExportEvidence"
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
'==============================================================================
'                                RefreshRibbon
'------------------------------------------------------------------------------
' PURPOSE
'   Ask Excel to call the state callbacks again, with the Ribbon object
'   held or recovered after a VBA state reset (CurrentRibbon, #216). Also
'   stores the pointer again when a cancelled save left it out. Without
'   the object (tab not loaded, no pointer of this session) it does
'   nothing.
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
    Dim Ribbon  As IRibbonUI    'The Ribbon, held or recovered

'------------------------------------------------------------------------------
' REFRESH
'------------------------------------------------------------------------------
        On Error GoTo Failed
        Set Ribbon = CurrentRibbon()
        If Ribbon Is Nothing Then
            Exit Sub
        End If
        If Len(StoredPointerText()) = 0 Then
            WritePointer PointerText(ObjPtr(Ribbon), Application.Hwnd)
        End If
        Ribbon.Invalidate

Failed:

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
