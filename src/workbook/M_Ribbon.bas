Attribute VB_Name = "M_Ribbon"
'==============================================================================
' MODULE: M_Ribbon
'------------------------------------------------------------------------------
' PURPOSE
'   Answer the callbacks of the workbook's SA-CCR Ribbon tab (#213, #215):
'   run the existing M_Main macro of a command, activate the sheet of a
'   Go to entry, and give the state of the controls. The contract is
'   docs/workbook/RIBBON.md; the RibbonX source is src/workbook/customUI14.xml.
'
' PUBLIC SURFACE
'   RibbonOnLoad: keep the Ribbon object, so that the tab can be refreshed.
'   RibbonAction: run the macro mapped to a command control.
'   RibbonGoTo: activate the sheet mapped to a Go to entry.
'   RibbonGetEnabled: whether a command can be clicked; every command is
'   enabled until #216 adds the state table.
'   RibbonGetLabel: the text of a Status label, read from a status cell.
'   RibbonMacro, RibbonSheet and ResultsLabelFor: the pure maps behind
'   them, for tests: IRibbonControl cannot be created from VBA.
'   RefreshRibbon: refresh the tab after an operation.
'   The callbacks have the signatures Office requires and are called by
'   Excel by name. Like the M_Main macros, they are a host-only surface and
'   are not listed in docs/PUBLIC_API.txt, which lists src/modules only.
'
' DEPENDENCIES
'   M_Main for every command; CORE_Util for the sheet and named-cell
'   lookups; CORE_Config for the sheet names, the Results status cell and
'   the staging name prefixes.
'
' STATE OWNERSHIP
'   Owns mRibbon, the Ribbon object given to RibbonOnLoad. A VBA state reset
'   clears it; RefreshRibbon then does nothing until the workbook is
'   reopened (#216 keeps it across resets).
'
' ERROR POLICY
'   No callback raises to Excel: Excel would show its own error and stop
'   calling the callback. RibbonAction and RibbonGoTo report a failure in a
'   message box; RibbonGetLabel shows "unavailable"; RibbonGetEnabled
'   answers True. The macros keep their own error handling and every gate
'   (staging validation, commit, export eligibility), whatever the Ribbon
'   shows.
'
' COMPATIBILITY
'   Excel VBA; IRibbonUI and IRibbonControl come from the Microsoft Office
'   Object Library, one of the workbook's default references.
'
' UPDATED
'   2026-11-05
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
' MODULE CONSTANTS
'------------------------------------------------------------------------------
    'Title of every message box: the project name.
        Private Const MSG_TITLE     As String = "K-SACCR"

'------------------------------------------------------------------------------
' MODULE STATE
'------------------------------------------------------------------------------
        Private mRibbon             As IRibbonUI    'The Ribbon, from RibbonOnLoad

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
'   Keep the Ribbon object Excel passes when it loads the tab.
'
' UPDATED
'   2026-11-05
'==============================================================================
'
        Set mRibbon = Ribbon

End Sub


Public Sub RibbonAction( _
    ByVal Control As IRibbonControl)
'
'==============================================================================
'                                RibbonAction
'------------------------------------------------------------------------------
' PURPOSE
'   Run the M_Main macro mapped to a command control (RibbonMacro), then
'   refresh the tab so that the Status labels follow.
'
' ERROR POLICY
'   The macros report their own outcome and never raise. An unknown control
'   or a command whose macro this workbook does not have yet is reported in
'   a message box, and nothing runs.
'
' UPDATED
'   2026-11-05
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
            Case ""
                MsgBox "The Ribbon control " & Control.Id & " runs no command.", vbExclamation, MSG_TITLE
            Case Else
                'Mapped, but delivered by a later change (RestoreRegulatoryParameters, #357).
                MsgBox "This command is not available in this workbook yet.", vbInformation, MSG_TITLE
        End Select
        RefreshRibbon
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
'   Whether a command can be clicked. Every command is enabled until #216
'   adds the state table of docs/workbook/RIBBON.md; the macros refuse what
'   they must in any case.
'
' UPDATED
'   2026-11-05
'==============================================================================
'
        ReturnedVal = True

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
'   A cell that cannot be read gives "unavailable"; nothing is raised.
'
' UPDATED
'   2026-11-05
'==============================================================================
'
        On Error GoTo Failed
        Select Case Control.Id
            Case "ksStatusResults"
                ReturnedVal = "Results: " & ResultsLabelFor(CORE_Util.SafeStr( _
                              CORE_Util.GetSheet(SH_RESULTS).Range(RESULTS_STATUS_CELL).Value2))
            Case "ksStatusNettingSets"
                ReturnedVal = "NettingSets staging: " & StagingLabel(IMPORT_NS_PREFIX)
            Case "ksStatusTrades"
                ReturnedVal = "Trades staging: " & StagingLabel(IMPORT_TR_PREFIX)
            Case Else
                ReturnedVal = vbNullString
        End Select
        Exit Sub

Failed:
        ReturnedVal = "Status unavailable"

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


Public Sub RefreshRibbon()
'
'==============================================================================
'                                RefreshRibbon
'------------------------------------------------------------------------------
' PURPOSE
'   Ask Excel to call the state callbacks again. Without the Ribbon object
'   (macros reset, tab not loaded) it does nothing.
'
' ERROR POLICY
'   Never raises.
'
' UPDATED
'   2026-11-05
'==============================================================================
'
        On Error Resume Next
        If Not mRibbon Is Nothing Then mRibbon.Invalidate
        On Error GoTo 0

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
