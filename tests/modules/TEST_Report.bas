Attribute VB_Name = "TEST_Report"
'==============================================================================
' MODULE: TEST_Report
'------------------------------------------------------------------------------
' PURPOSE
'   Check the formatted run report (#259): the text the PDF writes for a
'   number, a ratio, a count and a date, on the examples of
'   docs/workbook/report.json, whatever the regional settings; Checks by
'   code; the refusals of Export evidence and a changed Dashboard; an XLSX
'   read back from its file, with values only, its sheets in order and the
'   figures of the run; a
'   PDF; and the failures that must leave an older report as it was. The
'   K-SACCR workbook must stay as it is: the results stay current.
'
' PUBLIC SURFACE
'   RunReportTests is the entry point. Option Private Module keeps it out
'   of the external workbook automation API.
'
' DEPENDENCIES
'   TEST_CaseRunner, which writes the inputs, records the checks and
'   restores the workbook; M_Main for the run, the report and the test
'   fault; CORE_Export for the text of the report; CORE_Util and
'   CORE_Config for the sheets, columns and constants.
'
' WORKSHEET SAFETY
'   As TEST_CaseRunner: inputs and outputs are rewritten during the run and
'   restored at the end. The report files are written to the temporary
'   folder and deleted. Use a development workbook.
'
' USAGE
'   Run TEST_Report.RunReportTests from the Immediate window.
'
' UPDATED
'   2026-11-26
'
' AUTHOR
'   Daniele Penza
'==============================================================================

'------------------------------------------------------------------------------
' MODULE SETTINGS
'------------------------------------------------------------------------------
    Option Explicit
    Option Private Module

'------------------------------------------------------------------------------
' MODULE CONSTANTS
'------------------------------------------------------------------------------
        Private Const VALUATION   As String = "2026-09-30"    'Valuation date of the cases
        Private Const CASES       As Long = 5                 'Cases in a complete run
        Private Const CHECKS      As Long = 55                'Checks in a complete run
        Private Const CLASS_TAG   As String = "illustrative"  'Reference class of every check
        Private Const SHEET_ORDER As String = "Run|Dashboard|Totals|Results|Checks*"


Public Sub RunReportTests()
'
'==============================================================================
'                                RunReportTests
'------------------------------------------------------------------------------
' PURPOSE
'   Run every case and report through TEST_CaseRunner.
'
' UPDATED
'   2026-11-26
'==============================================================================
'
        If Not TEST_CaseRunner.BeginSuite("TEST_Report", CASES, CHECKS) Then Exit Sub
        On Error GoTo Failed

        TEST_CaseRunner.BeginCase "report-text", "CRR", VALUATION, "EUR"
        CheckTexts

        TEST_CaseRunner.BeginCase "report-refusal", "CRR", VALUATION, "EUR"
        TEST_CaseRunner.LoadSavedInputs
        CheckRefusals

        TEST_CaseRunner.BeginCase "report-xlsx", "CRR", VALUATION, "EUR"
        TEST_CaseRunner.LoadSavedInputs
        CheckWorkbookReport

        TEST_CaseRunner.BeginCase "report-pdf", "CRR", VALUATION, "EUR"
        TEST_CaseRunner.LoadSavedInputs
        CheckPdfReport

        TEST_CaseRunner.BeginCase "report-failures", "CRR", VALUATION, "EUR"
        TEST_CaseRunner.LoadSavedInputs
        CheckFailures

        M_Main.gTestFault = ""
        TEST_CaseRunner.EndSuite ""
        Exit Sub

Failed:
        M_Main.gTestFault = ""
        TEST_CaseRunner.EndSuite "unexpected error " & Err.Number & ": " & Err.Description

End Sub


'
'------------------------------------------------------------------------------
'
'                                    CASES
'
'------------------------------------------------------------------------------
'

Private Sub CheckTexts()
'
'==============================================================================
'                                  CheckTexts
'------------------------------------------------------------------------------
' PURPOSE
'   The text of the PDF: the examples of docs/workbook/report.json, which
'   tools/report.py also checks (tools/test_report_spec.py holds this list
'   to that file); the format of a reconciled figure; a value in the
'   format of its cell; the line of the rows left out; Checks by code.
'   26 checks.
'
' UPDATED
'   2026-11-26
'==============================================================================
'
    Dim CheckTable(1 To 5, 1 To 8)   As Variant    'Rows of Checks
    Dim Summary                      As Variant    'A row per code
    Dim Order                        As Variant    'Rows in the order of the report

    'The examples of report.json, one check each.
        ExpectText "amount-grouped", "amount", 1234567.891, "1,234,567.89"
        ExpectText "amount-half-up", "amount", 0.005, "0.01"
        ExpectText "amount-no-negative-zero", "amount", -0.004, "0.00"
        ExpectText "amount-half-away", "amount", -2.675, "-2.68"
        ExpectText "amount-carry", "amount", 999.995, "1,000.00"
        ExpectText "amount-zero", "amount", 0#, "0.00"
        ExpectText "amount-ead", "amount", 50640180.761782, "50,640,180.76"
        ExpectText "ratio", "ratio", 0.9585858586, "0.958586"
        ExpectText "ratio-padded", "ratio", 1.4, "1.400000"
        ExpectText "count", "count", 14#, "14"
        ExpectText "date", "date", 46295#, "2026-09-30"
        ExpectText "date-first", "date", 61#, "1900-03-01"
        ExpectText "datetime", "datetime", 46352.5340277778, "2026-11-26 12:49:00"
        ExpectText "datetime-midnight", "datetime", 46352.999994213, "2026-11-27 00:00:00"
        ExpectText "datetime-morning", "datetime", 46352.25, "2026-11-26 06:00:00"
        ExpectText "text-in-a-number-column", "amount", "n/a", "n/a"

    'A multiplier is a ratio, an EAD an amount.
        TEST_CaseRunner.ExpectTrue "measure-ratio", _
            CORE_Export.ReportMeasureKind("Multiplier", "Multiplier") = "ratio", "Multiplier", CLASS_TAG
        TEST_CaseRunner.ExpectTrue "measure-amount", _
            CORE_Export.ReportMeasureKind("EAD", "Multiplier") = "amount", "EAD", CLASS_TAG

    'A value in the format of its cell, separators invariant.
        ExpectFormat "format-amount", 1234.5678, "#,##0.00", "1,234.57"
        ExpectFormat "format-percent", 0.0045, "0.0%", "0.5%"
        ExpectFormat "format-date", 46295#, "yyyy-mm-dd", "2026-09-30"
        ExpectFormat "format-general", 0.25, "General", "0.25"

    'The line of the rows a PDF leaves out.
        TEST_CaseRunner.ExpectTrue "omitted-none", CORE_Export.ReportOmittedNote(0) = "", _
            CORE_Export.ReportOmittedNote(0), CLASS_TAG
        TEST_CaseRunner.ExpectTrue "omitted-three", CORE_Export.ReportOmittedNote(3) = _
            "3 rows with the result MATCH are left out; the workbook version of the report has every row.", _
            CORE_Export.ReportOmittedNote(3), CLASS_TAG

    'Checks by code: ERROR codes first, a blank code is (none), each code's
    'rows in their order.
        SetCheck CheckTable, 1, "WARNING", "TR-LATE"
        SetCheck CheckTable, 2, "ERROR", "TR-MATURED"
        SetCheck CheckTable, 3, "INFO", "INF-NOTE"
        SetCheck CheckTable, 4, "ERROR", "TR-LATE"
        SetCheck CheckTable, 5, "WARNING", ""
        Order = CORE_Export.ReportCodeOrder(CheckTable, 1, CK_CODE_COL, Summary)
        TEST_CaseRunner.ExpectTrue "codes", SummaryText(Summary) = _
            "TR-LATE/ERROR/2|TR-MATURED/ERROR/1|(none)/WARNING/1|INF-NOTE/INFO/1", SummaryText(Summary), CLASS_TAG
        TEST_CaseRunner.ExpectTrue "rows", Order(1) = 1 And Order(2) = 4 And Order(3) = 2 And Order(4) = 5 And _
            Order(5) = 3, Order(1) & "," & Order(2) & "," & Order(3) & "," & Order(4) & "," & Order(5), CLASS_TAG

End Sub


Private Sub CheckRefusals()
'
'==============================================================================
'                                CheckRefusals
'------------------------------------------------------------------------------
' PURPOSE
'   The report refuses another extension, and an output cell edited after
'   the run, as Export evidence does; and a figure of the Dashboard changed
'   after the run, which the evidence CSV does not copy (Codex on #805).
'   3 checks.
'
' UPDATED
'   2026-11-26
'==============================================================================
'
    Dim Result        As String    'Result line
    Dim OutputCell    As Range     'An output cell
    Dim SavedOutput   As Variant   'Its value
    Dim BlockNames    As Variant   'The Dashboard's blocks
    Dim DashboardCell As Range     'A figure of the Dashboard
    Dim SavedFigure   As Variant   'Its value

        Result = M_Main.RunSACCR_Silent()
        Application.Calculate
        Result = M_Main.ExportReport_Silent(ReportPath("txt"))
        TEST_CaseRunner.ExpectTrue "extension", Left$(Result, 14) = "RESULT=REFUSED" And _
            InStr(1, Result, ".pdf or .xlsx", vbBinaryCompare) > 0, Result, CLASS_TAG
        Set OutputCell = TableCell(GetSheet(SH_RESULTS), OutputFirstRow(SH_RESULTS), RS_EAD)
        SavedOutput = OutputCell.Value2
        OutputCell.Value2 = ToDbl(SavedOutput, 0#) + 1#
        Application.Calculate
        Result = M_Main.ExportReport_Silent(ReportPath("xlsx"))
        OutputCell.Value2 = SavedOutput
        TEST_CaseRunner.ExpectTrue "edited-output", Left$(Result, 14) = "RESULT=REFUSED" And _
            InStr(1, Result, "output cell", vbBinaryCompare) > 0, Result, CLASS_TAG
    'A figure of the Dashboard changed after the run, then put back.
        BlockNames = CORE_Engine.DashboardBlockNames()
        Set DashboardCell = ThisWorkbook.Names(CStr(BlockNames(LBound(BlockNames)))).RefersToRange.Cells(1, 1)
        SavedFigure = DashboardCell.Value2
        If VarType(SavedFigure) = vbDouble Then
            DashboardCell.Value2 = SavedFigure + 1#
        Else
            DashboardCell.Value2 = SafeStr(SavedFigure) & "x"
        End If
        Application.Calculate
        Result = M_Main.ExportReport_Silent(ReportPath("xlsx"))
        DashboardCell.Value2 = SavedFigure
        TEST_CaseRunner.ExpectTrue "edited-dashboard", Left$(Result, 14) = "RESULT=REFUSED" And _
            InStr(1, Result, "Dashboard", vbBinaryCompare) > 0, Result, CLASS_TAG

End Sub


Private Sub CheckWorkbookReport()
'
'==============================================================================
'                             CheckWorkbookReport
'------------------------------------------------------------------------------
' PURPOSE
'   An XLSX report of the demonstration run, read back from its file: the
'   sheets in order, values only, no link and no macro, the figures of the
'   run; the K-SACCR workbook left as it was. 15 checks. The suite runs in
'   manual calculation, so each run is followed by a recalculation, as a user
'   presses F9: the report refuses while one is pending.
'
' UPDATED
'   2026-11-26
'==============================================================================
'
    Dim Result          As String       'Result line
    Dim FilePath        As String       'The report
    Dim BookCount       As Long         'Workbooks open before the export
    Dim Report          As Workbook     'The report, opened read-only
    Dim SheetNames      As String       'Its sheets, in order
    Dim ReportSheet     As Worksheet    'A sheet of it
    Dim HasFormula      As Boolean      'A sheet holds a formula
    Dim Found           As Range        'A cell found
    Dim CheckRows       As Long         'Rows of Checks on the workbook
    Dim CountedRows     As Double       'Rows counted by the summary
    Dim r               As Long         'Row
    Dim TotalEad        As Variant      'EAD of TOTAL CRR on Results

        Result = M_Main.RunSACCR_Silent()
        Application.Calculate
        FilePath = ReportPath("xlsx")
        BookCount = Application.Workbooks.Count
        Result = M_Main.ExportReport_Silent(FilePath)
        TEST_CaseRunner.ExpectTrue "exported", Left$(Result, 9) = "RESULT=OK" And _
            InStr(1, Result, "format=XLSX", vbBinaryCompare) > 0, Result, CLASS_TAG
        TEST_CaseRunner.ExpectTrue "results-current", M_Main.ResultsStatus() = "CURRENT", _
            M_Main.ResultsStatus(), CLASS_TAG
        TEST_CaseRunner.ExpectTrue "active-workbook", Application.ActiveWorkbook Is ThisWorkbook, _
            Application.ActiveWorkbook.Name, CLASS_TAG
        TEST_CaseRunner.ExpectTrue "no-workbook-left-open", Application.Workbooks.Count = BookCount, _
            Application.Workbooks.Count & " open, " & BookCount & " before", CLASS_TAG

    'Without a file the eleven checks that read it fail, so that the count holds.
        If Left$(Result, 9) <> "RESULT=OK" Then
            FailChecks "sheets|values-only|no-link|xlsx|fingerprint|fingerprint-value|reporting-date|" & _
                       "reporting-date-value|total-ead|checks-summary|no-limitations", Result
            Exit Sub
        End If

        Set Report = Application.Workbooks.Open(Filename:=FilePath, ReadOnly:=True, UpdateLinks:=0)
        For Each ReportSheet In Report.Worksheets
            SheetNames = SheetNames & IIf(Len(SheetNames) = 0, "", "|") & ReportSheet.Name
            If Not IsNull(ReportSheet.UsedRange.HasFormula) Then
                If ReportSheet.UsedRange.HasFormula Then HasFormula = True
            Else
                HasFormula = True
            End If
        Next ReportSheet
        TEST_CaseRunner.ExpectTrue "sheets", SheetNames Like SHEET_ORDER, SheetNames, CLASS_TAG
        TEST_CaseRunner.ExpectTrue "values-only", Not HasFormula, "a sheet holds a formula", CLASS_TAG
        TEST_CaseRunner.ExpectTrue "no-link", IsEmpty(Report.LinkSources(xlExcelLinks)), "links found", CLASS_TAG
        TEST_CaseRunner.ExpectTrue "xlsx", Report.FileFormat = xlOpenXMLWorkbook And Not Report.HasVBProject, _
            "format " & Report.FileFormat, CLASS_TAG

    'The run's figures.
        Set Found = Report.Worksheets("Run").Columns(1).Find(What:="Input fingerprint", LookAt:=xlWhole)
        TEST_CaseRunner.ExpectTrue "fingerprint", Not Found Is Nothing, "no fingerprint row", CLASS_TAG
        If Not Found Is Nothing Then
            TEST_CaseRunner.ExpectTrue "fingerprint-value", _
                SafeStr(Found.Offset(0, 1).Value2) = CORE_Engine.LastRunInputs(), _
                SafeStr(Found.Offset(0, 1).Value2), CLASS_TAG
        Else
            TEST_CaseRunner.ExpectTrue "fingerprint-value", False, "no fingerprint row", CLASS_TAG
        End If
        Set Found = Report.Worksheets("Run").Columns(1).Find(What:="Reporting date", LookAt:=xlWhole)
        TEST_CaseRunner.ExpectTrue "reporting-date", Not Found Is Nothing, "no reporting date row", CLASS_TAG
        If Not Found Is Nothing Then
            TEST_CaseRunner.ExpectTrue "reporting-date-value", VarType(Found.Offset(0, 1).Value2) = vbDouble And _
                Found.Offset(0, 1).NumberFormat = "yyyy-mm-dd", SafeStr(Found.Offset(0, 1).Value2), CLASS_TAG
        Else
            TEST_CaseRunner.ExpectTrue "reporting-date-value", False, "no reporting date row", CLASS_TAG
        End If
        TotalEad = TotalOf("TOTAL " & RG_CRR)
        Set Found = Report.Worksheets("Results").Columns(1).Find(What:="TOTAL " & RG_CRR, LookAt:=xlWhole)
        If Found Is Nothing Then
            TEST_CaseRunner.ExpectTrue "total-ead", False, "no TOTAL CRR row", CLASS_TAG
        Else
            TEST_CaseRunner.ExpectTrue "total-ead", Found.Offset(0, RS_EAD - 1).Value2 = TotalEad, _
                SafeStr(Found.Offset(0, RS_EAD - 1).Value2) & " against " & SafeStr(TotalEad), CLASS_TAG
        End If
        CheckRows = LastDataRow(GetSheet(SH_CHECKS), OutputFirstRow(SH_CHECKS), 1) - OutputFirstRow(SH_CHECKS) + 1
        With Report.Worksheets("Checks")
            r = 3
            Do While Len(SafeStr(.Cells(r, 1).Value2)) > 0
                CountedRows = CountedRows + ToDbl(.Cells(r, 3).Value2, 0#)
                r = r + 1
            Loop
        End With
        TEST_CaseRunner.ExpectTrue "checks-summary", CountedRows = CheckRows, _
            CountedRows & " rows counted, " & CheckRows & " on Checks", CLASS_TAG
    'The report has no Limitations sheet (owner's verification of 2026-11-26)
        TEST_CaseRunner.ExpectTrue "no-limitations", _
            InStr(1, "|" & SheetNames & "|", "|Limitations|", vbBinaryCompare) = 0, SheetNames, CLASS_TAG
        Report.Close SaveChanges:=False
        Kill FilePath

End Sub


Private Sub CheckPdfReport()
'
'==============================================================================
'                               CheckPdfReport
'------------------------------------------------------------------------------
' PURPOSE
'   A PDF report of the demonstration run: a PDF file, the Dashboard's
'   charts pictured, the K-SACCR workbook left as it was. The pages are
'   checked by the owner (REPORT.md). 4 checks.
'
' UPDATED
'   2026-11-26
'==============================================================================
'
    Dim Result     As String    'Result line
    Dim FilePath   As String    'The report
    Dim Charts     As Long      'Charts of the Dashboard

        Result = M_Main.RunSACCR_Silent()
        Application.Calculate
        FilePath = ReportPath("pdf")
        Charts = GetSheet(SH_DASHBOARD).ChartObjects.Count
        Result = M_Main.ExportReport_Silent(FilePath)
        Debug.Print "TEST_Report PDF: " & Result
        TEST_CaseRunner.ExpectTrue "exported", Left$(Result, 9) = "RESULT=OK" And _
            InStr(1, Result, "format=PDF", vbBinaryCompare) > 0, Result, CLASS_TAG
        TEST_CaseRunner.ExpectTrue "pdf-file", Left$(FileStart(FilePath), 5) = "%PDF-", FileStart(FilePath), CLASS_TAG
        TEST_CaseRunner.ExpectTrue "charts", InStr(1, Result, "; charts=" & Charts & ";", vbBinaryCompare) > 0, _
            Result & " (" & Charts & " charts on the Dashboard)", CLASS_TAG
        TEST_CaseRunner.ExpectTrue "results-current", M_Main.ResultsStatus() = "CURRENT", _
            M_Main.ResultsStatus(), CLASS_TAG
        If Len(Dir$(FilePath)) > 0 Then Kill FilePath

End Sub


Private Sub CheckFailures()
'
'==============================================================================
'                                CheckFailures
'------------------------------------------------------------------------------
' PURPOSE
'   A report that fails before it is saved, or while an older report is
'   replaced, leaves the older one as it was and no file beside it; a
'   read-only report is not replaced. 7 checks.
'
' UPDATED
'   2026-11-26
'==============================================================================
'
    Dim Result        As String    'Result line
    Dim FilePath      As String    'The older report
    Dim Fault         As Variant   'Each injected failure
    Dim ErrorNumber   As Long      'Error the export raised

        Result = M_Main.RunSACCR_Silent()
        Application.Calculate
        FilePath = ReportPath("xlsx")
        For Each Fault In Array("report-save", "report-replace")
            WriteOldReport FilePath
            M_Main.gTestFault = CStr(Fault)
            ErrorNumber = 0
            On Error Resume Next
            Result = M_Main.ExportReport_Silent(FilePath)
            ErrorNumber = Err.Number
            On Error GoTo 0
            M_Main.gTestFault = ""
            TEST_CaseRunner.ExpectTrue CStr(Fault) & "-raised", ErrorNumber = ERR_INJECTED_FAULT, _
                "error " & ErrorNumber, CLASS_TAG
            TEST_CaseRunner.ExpectTrue CStr(Fault) & "-older-kept", FileStart(FilePath) = "older report", _
                FileStart(FilePath), CLASS_TAG
            TEST_CaseRunner.ExpectTrue CStr(Fault) & "-nothing-left", _
                Len(Dir$(ReportFolder() & "\saccr-*", vbHidden)) = 0, _
                Dir$(ReportFolder() & "\saccr-*", vbHidden), CLASS_TAG
        Next Fault

    'A read-only report is refused and kept.
        WriteOldReport FilePath
        SetAttr FilePath, vbReadOnly
        ErrorNumber = 0
        On Error Resume Next
        Result = M_Main.ExportReport_Silent(FilePath)
        ErrorNumber = Err.Number
        On Error GoTo 0
        SetAttr FilePath, vbNormal
        TEST_CaseRunner.ExpectTrue "read-only", _
            ErrorNumber = ERR_EXPORT_WRITE And FileStart(FilePath) = "older report", _
            "error " & ErrorNumber & ", " & FileStart(FilePath), CLASS_TAG
        Kill FilePath

End Sub


'
'------------------------------------------------------------------------------
'
'                                   HELPERS
'
'------------------------------------------------------------------------------
'

Private Sub ExpectText( _
    ByVal Label As String, _
    ByVal Kind As String, _
    ByVal Value As Variant, _
    ByVal Expected As String)
'
'==============================================================================
' PURPOSE
'   One check of CORE_Export.ReportCellText.
'==============================================================================
'
    Dim Found   As String    'The text written

        Found = CORE_Export.ReportCellText(Kind, Value)
        TEST_CaseRunner.ExpectTrue Label, Found = Expected, Found & " instead of " & Expected, CLASS_TAG

End Sub


Private Sub FailChecks( _
    ByVal Labels As String, _
    ByVal Detail As String)
'
'==============================================================================
' PURPOSE
'   Record as failed each check of Labels (separated by "|") that could not
'   run, so that the suite's count of checks holds.
'==============================================================================
'
    Dim Label   As Variant    'Each check

        For Each Label In Split(Labels, "|")
            TEST_CaseRunner.ExpectTrue CStr(Label), False, "not run: " & Detail, CLASS_TAG
        Next Label

End Sub


Private Sub ExpectFormat( _
    ByVal Label As String, _
    ByVal Value As Variant, _
    ByVal NumberFormat As String, _
    ByVal Expected As String)
'
'==============================================================================
' PURPOSE
'   One check of CORE_Export.ReportFormatText.
'==============================================================================
'
    Dim Found   As String    'The text written

        Found = CORE_Export.ReportFormatText(Value, NumberFormat)
        TEST_CaseRunner.ExpectTrue Label, Found = Expected, Found & " instead of " & Expected, CLASS_TAG

End Sub


Private Sub SetCheck( _
    ByRef CheckTable() As Variant, _
    ByVal RowNumber As Long, _
    ByVal Severity As String, _
    ByVal Code As String)
'
'==============================================================================
' PURPOSE
'   Fill one row of a Checks table: its severity and its code.
'==============================================================================
'
        CheckTable(RowNumber, 1) = Severity
        CheckTable(RowNumber, CK_CODE_COL) = Code

End Sub


Private Function SummaryText( _
    ByVal Summary As Variant) _
    As String
'
'==============================================================================
' PURPOSE
'   The summary of Checks by code as one line: code/severity/rows|...
'==============================================================================
'
    Dim k      As Long      'Code
    Dim Line   As String    'The line

        If Not IsArray(Summary) Then Exit Function
        For k = 1 To UBound(Summary, 1)
            Line = Line & IIf(k > 1, "|", "") & Summary(k, 1) & "/" & Summary(k, 2) & "/" & Summary(k, 3)
        Next k
        SummaryText = Line

End Function


Private Function TotalOf( _
    ByVal TotalId As String) _
    As Variant
'
'==============================================================================
' PURPOSE
'   The EAD of a TOTAL row of Results; Empty when there is none.
'==============================================================================
'
    Dim r          As Long    'Row
    Dim LastRow    As Long    'Last row of Results

        LastRow = LastDataRow(GetSheet(SH_RESULTS), OutputFirstRow(SH_RESULTS), RS_ID)
        For r = OutputFirstRow(SH_RESULTS) To LastRow
            If SafeStr(TableCell(GetSheet(SH_RESULTS), r, RS_ID).Value2) = TotalId Then
                TotalOf = TableCell(GetSheet(SH_RESULTS), r, RS_EAD).Value2
                Exit Function
            End If
        Next r
        TotalOf = Empty

End Function


Private Function ReportFolder() As String
'
'==============================================================================
' PURPOSE
'   The temporary folder, or the workbook's folder when there is none.
'==============================================================================
'
        ReportFolder = Environ$("TEMP")
        If Len(ReportFolder) = 0 Then
            ReportFolder = ThisWorkbook.Path
        End If

End Function


Private Function ReportPath( _
    ByVal Extension As String) _
    As String
'
'==============================================================================
' PURPOSE
'   A report file of the tests in the temporary folder, deleted first.
'==============================================================================
'
        ReportPath = ReportFolder() & Application.PathSeparator & "K-SACCR_report_test." & Extension
        If Len(Dir$(ReportPath, vbHidden)) > 0 Then
            SetAttr ReportPath, vbNormal
            Kill ReportPath
        End If

End Function


Private Sub WriteOldReport( _
    ByVal FilePath As String)
'
'==============================================================================
' PURPOSE
'   Write an older report: a file whose text is "older report".
'==============================================================================
'
    Dim FileNumber   As Integer    'File number

        If Len(Dir$(FilePath, vbHidden)) > 0 Then
            SetAttr FilePath, vbNormal
            Kill FilePath
        End If
        FileNumber = FreeFile
        Open FilePath For Binary Access Write As #FileNumber
        Put #FileNumber, , "older report"
        Close #FileNumber

End Sub


Private Function FileStart( _
    ByVal FilePath As String) _
    As String
'
'==============================================================================
' PURPOSE
'   The first twelve bytes of a file as text; "" when there is no file.
'==============================================================================
'
    Dim FileNumber   As Integer    'File number
    Dim Bytes()      As Byte       'The bytes

        If Len(Dir$(FilePath, vbHidden)) = 0 Then Exit Function
        FileNumber = FreeFile
        Open FilePath For Binary Access Read As #FileNumber
        If LOF(FileNumber) > 0 Then
            ReDim Bytes(0 To IIf(LOF(FileNumber) < 12, LOF(FileNumber), 12) - 1)
            Get #FileNumber, , Bytes
            FileStart = StrConv(Bytes, vbUnicode)
        End If
        Close #FileNumber

End Function
