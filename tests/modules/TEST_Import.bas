Attribute VB_Name = "TEST_Import"
'==============================================================================
' MODULE: TEST_Import
'------------------------------------------------------------------------------
' PURPOSE
'   Check the CSV staging load (#208), M_Main.ImportCsv_Silent: a file
'   whose findings are all about values is loaded into Import_NettingSets
'   or Import_Trades as text, each record with its CSV line, a blank
'   record left out and a quoted line break kept; leading zeros and text
'   that starts like a formula stay as written and no cell becomes a
'   formula; the metadata and the stored fingerprint follow the load. A
'   structural finding, a missing file or an unknown kind loads nothing.
'   A change since the last load is found with events off and refuses a
'   new load unless replacement is asked for; with events on, the change
'   handler shows EDITED. A shorter file clears the rows below it. A load
'   that fails while writing puts the sheet back as it was, a formula in
'   it included. A value longer than a cell holds loads nothing. A formula
'   in the block is found (M_Main.StagingFormulaCell). The live inputs,
'   the run fingerprints and the other staging sheet are never changed.
'
' PUBLIC SURFACE
'   RunImportTests is the entry point. Option Private Module keeps it out
'   of the external workbook automation API.
'
' DEPENDENCIES
'   M_Main, with its test seam gTestFault, for the load; CORE_Engine for
'   BlockText and InputFingerprint; CORE_Util for the cells and TextHash;
'   CORE_Export for the UTF-8 bytes; CORE_Config for the sheet and name
'   constants, the CSV header lines and the error numbers; TEST_Suite.Say
'   for the log lines. Needs a workbook built from the current template,
'   with the two staging sheets. Every record is synthetic.
'
' WORKSHEET SAFETY
'   The cases load files into Import_NettingSets and Import_Trades and
'   edit their cells, then put both sheets back: their blocks and metadata
'   as values, the data columns formatted as text, and the hidden
'   fingerprint names, also after an unexpected error. The files are
'   written to the TEMP folder (K-SACCR_import_test.csv) and deleted. No
'   input or output sheet is touched.
'
' STATE OWNERSHIP
'   Restores the caller's calculation mode, events, M_Main.gSilent and
'   M_Main.gQuietLog; sets M_Main.gTestFault back to "", also after an
'   unexpected error.
'
' ERROR POLICY
'   An unexpected error is recorded against its case and the remaining
'   cases still run.
'
' USAGE
'   Import with the production modules, compile, then run
'   TEST_Import.RunImportTests from the Immediate window. It ends with
'   RESULT=PASS; cases=...; checks=...; failures=0.
'
' UPDATED
'   2026-10-26
'
' AUTHOR
'   Daniele Penza
'==============================================================================

'------------------------------------------------------------------------------
' MODULE SETTINGS
'------------------------------------------------------------------------------
    'Require explicit declarations; keep the tests out of the external API.
    Option Explicit
    Option Private Module

'------------------------------------------------------------------------------
' MODULE CONSTANTS
'------------------------------------------------------------------------------
        Private Const EXPECTED_CASES    As Long = 10                           'Cases in a complete run
        Private Const EXPECTED_CHECKS   As Long = 52                           'Checks in a complete run
        Private Const TEST_FILE         As String = "K-SACCR_import_test.csv"  'Test file in TEMP
        Private Const MISSING_FILE      As String = "K-SACCR_import_missing.csv"    'A file never written
        Private Const KIND_NS           As Long = 1                            'Import_NettingSets
        Private Const KIND_TR           As Long = 2                            'Import_Trades
    'Synthetic records: two netting sets, and a trade leg of NS-T1 whose
    'ID, notional and comment each case chooses.
        Private Const NS_ROW_1          As String = "NS-T1;CPTY A;Y;N;1;N;N;;0;0;0;0;;;;;"
        Private Const NS_ROW_2          As String = "NS-T2;CPTY B;N;N;1;N;N;;0;0;0;0;;;;;"
    'Table columns of the staging block: the CSV line, then the contract's.
        Private Const COL_LINE          As Long = 1     'CSV line
        Private Const COL_ID            As Long = 2     'trade_id or netting_set_id
        Private Const COL_NS_REMARGIN   As Long = 6     'remargin_period_business_days
        Private Const COL_TR_NOTIONAL   As Long = 12    'notional_amount
        Private Const COL_TR_COMMENT    As Long = 29    'comment

'------------------------------------------------------------------------------
' MODULE STATE
'------------------------------------------------------------------------------
        Private mCaseCount      As Long      'Cases started
        Private mCheckCount     As Long      'Checks evaluated
        Private mFailureCount   As Long      'Failed checks and unexpected errors
        Private mCurrentCase    As String    'Case being run, for messages
        Private mTestFile       As String    'Path of the test file
        Private mNsLoaded       As String    'Import_NettingSets block after the first case
    'Each staging sheet as the suite found it, put back at the end.
        Private mSavedBlock(1 To 2)  As Variant    'Block values; Empty for none
        Private mSavedRows(1 To 2)   As Long       'Its rows
        Private mSavedMeta(1 To 2)   As Variant    'The four metadata values
        Private mSavedName(1 To 2)   As String     'RefersTo of the fingerprint name
        Private mHadName(1 To 2)     As Boolean    'The fingerprint name existed


'
'------------------------------------------------------------------------------
'
'                                 ENTRY POINT
'
'------------------------------------------------------------------------------
'

Public Sub RunImportTests()
'
'==============================================================================
'                                RunImportTests
'------------------------------------------------------------------------------
' PURPOSE
'   Run every case and print one CASE line per case, each failed check,
'   and a final RESULT line.
'
' UPDATED
'   2026-10-26
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim CallerCalculation   As XlCalculation    'Caller's calculation mode
    Dim CallerEvents        As Boolean          'Caller's EnableEvents
    Dim CallerSilent        As Boolean          'Caller's M_Main.gSilent
    Dim CallerQuiet         As Boolean          'Caller's M_Main.gQuietLog
    Dim LiveBefore          As String           'Live inputs and run names before the cases
    Dim TempFolder          As String           'Folder of the test file

'------------------------------------------------------------------------------
' RUN
'------------------------------------------------------------------------------
        mCaseCount = 0
        mCheckCount = 0
        mFailureCount = 0
        TEST_Suite.Say "SACCR CSV IMPORT TESTS"
        If NamedCell(IMPORT_NS_HEADER_NAME) Is Nothing Or NamedCell(IMPORT_TR_HEADER_NAME) Is Nothing Then
            Fail "suite", "Import_NettingSets or Import_Trades is missing: build the workbook from the current template"
            TEST_Suite.Say "RESULT=FAIL; cases=0; checks=0; failures=" & mFailureCount
            Exit Sub
        End If
        CallerCalculation = Application.Calculation
        CallerEvents = Application.EnableEvents
        CallerSilent = M_Main.gSilent
        CallerQuiet = M_Main.gQuietLog
        M_Main.gSilent = True
        M_Main.gQuietLog = True
        TempFolder = Environ$("TEMP")
        If Len(TempFolder) = 0 Then
            TempFolder = ThisWorkbook.Path
        End If
        mTestFile = TempFolder & Application.PathSeparator & TEST_FILE
        LiveBefore = LiveInputsText()
        SaveStaging KIND_NS
        SaveStaging KIND_TR

        CaseNettingSetsAccepted
        CaseTradesValueFindings
        CaseStructuralRejected
        CaseMissingFile
        CaseEdits
        CaseWriteFailure
        CaseCellLimit
        CaseFormulaCell
        CaseHeaderOnly
        CaseLiveSheets LiveBefore

        PutStagingBack KIND_NS
        PutStagingBack KIND_TR
        RemoveFile mTestFile
        M_Main.gTestFault = ""
        Application.Calculation = CallerCalculation
        Application.EnableEvents = CallerEvents
        M_Main.gSilent = CallerSilent
        M_Main.gQuietLog = CallerQuiet
        If mCaseCount <> EXPECTED_CASES Then
            Fail "suite", "expected " & EXPECTED_CASES & " cases, ran " & mCaseCount
        End If
        If mCheckCount <> EXPECTED_CHECKS Then
            Fail "suite", "expected " & EXPECTED_CHECKS & " checks, ran " & mCheckCount
        End If
        TEST_Suite.Say "RESULT=" & IIf(mFailureCount = 0, "PASS", "FAIL") & _
                    "; cases=" & mCaseCount & "; checks=" & mCheckCount & _
                    "; failures=" & mFailureCount

End Sub


'
'------------------------------------------------------------------------------
'
'                                    CASES
'
'------------------------------------------------------------------------------
'

Private Sub CaseNettingSetsAccepted()
'
'==============================================================================
' PURPOSE
'   An accepted NettingSets file: two records, on lines 2 and 3, as text,
'   with the file, time, count, status and fingerprint of the load.
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ResultLine   As String    'What the load returned

'------------------------------------------------------------------------------
' LOAD
'------------------------------------------------------------------------------
        On Error GoTo Unexpected
        BeginCase "import.netting-sets-accepted"
        WriteCsv CSV_NS_HEADER & vbLf & NS_ROW_1 & vbLf & NS_ROW_2 & vbLf
        ResultLine = M_Main.ImportCsv_Silent("NettingSets", mTestFile, True)
        CheckText "result", ResultLine, "RESULT=LOADED; records=2; findings=0; file=" & mTestFile
        CheckText "status", Meta(KIND_NS, 4), STAGING_LOADED
        CheckText "records", Meta(KIND_NS, 3), "2"
        CheckText "file", Meta(KIND_NS, 1), mTestFile
        Check Len(Meta(KIND_NS, 2)) = 19, "load time '" & Meta(KIND_NS, 2) & "'"
        Check CellValue(KIND_NS, 1, COL_LINE) = 2 And CellValue(KIND_NS, 2, COL_LINE) = 3, "CSV lines"
        CheckText "first ID", CStr(CellValue(KIND_NS, 1, COL_ID)), "NS-T1"
        Check VarType(CellValue(KIND_NS, 1, COL_NS_REMARGIN)) = vbString, "a number field was not kept as text"
        Check IsEmpty(CellValue(KIND_NS, 3, COL_LINE)) And IsEmpty(CellValue(KIND_NS, 3, COL_ID)), "a third row"
        CheckText "stored fingerprint", NameText(STAGED_NS_NAME), "=""" & TextHash(BlockState(KIND_NS)) & """"
        mNsLoaded = BlockState(KIND_NS)
        Exit Sub

Unexpected:
        Fail mCurrentCase, "unexpected error " & Err.Number & ": " & Err.Description

End Sub


Private Sub CaseTradesValueFindings()
'
'==============================================================================
' PURPOSE
'   A CRLF Trades file with findings about values only: a notional that
'   is not a number, a comment that starts like a formula and a blank
'   record. The three other records are loaded, on lines 2, 4 and 6, the
'   second with a quoted line break; leading zeros and the formula-like
'   comment stay as written, and Import_NettingSets is left as it was.
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ResultLine   As String     'What the load returned
    Dim Comment      As String     'The comment with a line break

'------------------------------------------------------------------------------
' LOAD
'------------------------------------------------------------------------------
        On Error GoTo Unexpected
        BeginCase "import.trades-value-findings"
        WriteCsv CSV_TR_HEADER & vbCrLf & TradeRow("00123", "abc", "=1+1") & vbCrLf & vbCrLf & _
                 TradeRow("T3", "1000", """two" & vbCrLf & "lines""") & vbCrLf & TradeRow("T4", "2000", "") & vbCrLf
        ResultLine = M_Main.ImportCsv_Silent("Trades", mTestFile, True)
        CheckText "result", ResultLine, "RESULT=LOADED; records=3; findings=3; file=" & mTestFile
        Check CellValue(KIND_TR, 1, COL_LINE) = 2 And CellValue(KIND_TR, 2, COL_LINE) = 4 And _
              CellValue(KIND_TR, 3, COL_LINE) = 6, "CSV lines"
        CheckText "leading zeros", CStr(CellValue(KIND_TR, 1, COL_ID)), "00123"
        CheckText "notional that is not a number", CStr(CellValue(KIND_TR, 1, COL_TR_NOTIONAL)), "abc"
        CheckText "formula-like comment", CStr(CellValue(KIND_TR, 1, COL_TR_COMMENT)), "=1+1"
        Check Not StagingCell(KIND_TR, 1, COL_TR_COMMENT).HasFormula, "the comment became a formula"
        Comment = CStr(CellValue(KIND_TR, 2, COL_TR_COMMENT))
        Check Left$(Comment, 3) = "two" And Right$(Comment, 5) = "lines" And InStr(Comment, vbLf) > 0, _
              "the quoted line break was not kept: '" & Comment & "'"
        CheckText "formula search", M_Main.StagingFormulaCell("Trades"), ""
        CheckText "status", Meta(KIND_TR, 4), STAGING_LOADED
        Check IsEmpty(CellValue(KIND_TR, 4, COL_LINE)), "a fourth row"
        Check BlockState(KIND_NS) = mNsLoaded, "Import_NettingSets changed"
        Exit Sub

Unexpected:
        Fail mCurrentCase, "unexpected error " & Err.Number & ": " & Err.Description

End Sub


Private Sub CaseStructuralRejected()
'
'==============================================================================
' PURPOSE
'   A comma-separated file has a structural finding: nothing is loaded and
'   the staging sheet keeps its block and metadata.
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ResultLine   As String    'What the load returned
    Dim Before       As String    'The block before
    Dim MetaBefore   As String    'The metadata before

'------------------------------------------------------------------------------
' LOAD
'------------------------------------------------------------------------------
        On Error GoTo Unexpected
        BeginCase "import.structural-rejected"
        Before = BlockState(KIND_TR)
        MetaBefore = MetaText(KIND_TR)
        WriteCsv Replace(CSV_TR_HEADER, ";", ",") & vbLf & Replace(TradeRow("T1", "1", ""), ";", ",") & vbLf
        ResultLine = M_Main.ImportCsv_Silent("Trades", mTestFile, True)
        CheckText "result", ResultLine, "RESULT=REJECTED; findings=1; first=CSV-DELIMITER|1|0|; file=" & mTestFile
        Check BlockState(KIND_TR) = Before, "the block changed"
        CheckText "metadata", MetaText(KIND_TR), MetaBefore
        Exit Sub

Unexpected:
        Fail mCurrentCase, "unexpected error " & Err.Number & ": " & Err.Description

End Sub


Private Sub CaseMissingFile()
'
'==============================================================================
' PURPOSE
'   A missing file and an unknown kind raise ERR_CSV_READ and load
'   nothing.
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ResultLine    As String    'What the load returned
    Dim Before        As String    'The block before
    Dim MissingPath   As String    'A file that does not exist
    Dim ErrorNumber   As Long      'Error raised by the load

'------------------------------------------------------------------------------
' LOAD
'------------------------------------------------------------------------------
        On Error GoTo Unexpected
        BeginCase "import.missing-file"
        Before = BlockState(KIND_TR)
        MissingPath = Left$(mTestFile, Len(mTestFile) - Len(TEST_FILE)) & MISSING_FILE
        RemoveFile MissingPath
        On Error Resume Next
        ResultLine = M_Main.ImportCsv_Silent("Trades", MissingPath, True)
        ErrorNumber = Err.Number
        On Error GoTo Unexpected
        Check ErrorNumber = ERR_CSV_READ, "missing file: error " & ErrorNumber
        Check BlockState(KIND_TR) = Before, "missing file: the block changed"
        ErrorNumber = 0
        On Error Resume Next
        ResultLine = M_Main.ImportCsv_Silent("Other", mTestFile, True)
        ErrorNumber = Err.Number
        On Error GoTo Unexpected
        Check ErrorNumber = ERR_CSV_READ, "unknown kind: error " & ErrorNumber
        Exit Sub

Unexpected:
        Fail mCurrentCase, "unexpected error " & Err.Number & ": " & Err.Description

End Sub


Private Sub CaseEdits()
'
'==============================================================================
' PURPOSE
'   A change made with events off leaves the status alone, but is found
'   by the fingerprint: a load is refused and the change kept. With events
'   on, a change shows EDITED. A load that may replace changes loads a
'   one-record file and clears the rows below it.
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ResultLine   As String     'What the load returned
    Dim Saved        As Boolean    'EnableEvents before the case

'------------------------------------------------------------------------------
' EDIT AND LOAD
'------------------------------------------------------------------------------
        On Error GoTo Unexpected
        BeginCase "import.edits"
        Saved = Application.EnableEvents
        Application.EnableEvents = False
        StagingCell(KIND_TR, 1, COL_ID).Value2 = "EDITED-1"
        CheckText "status with events off", Meta(KIND_TR, 4), STAGING_LOADED
        WriteCsv CSV_TR_HEADER & vbLf & TradeRow("T9", "5", "") & vbLf
        ResultLine = M_Main.ImportCsv_Silent("Trades", mTestFile)
        CheckText "refused", ResultLine, "RESULT=REFUSED; reason=" & SH_IMPORT_TR & _
                  " has changes since its last load; file=" & mTestFile
        CheckText "the change is kept", CStr(CellValue(KIND_TR, 1, COL_ID)), "EDITED-1"
        Application.EnableEvents = True
        StagingCell(KIND_TR, 2, COL_ID).Value2 = "EDITED-2"
        CheckText "status after a change", Meta(KIND_TR, 4), STAGING_EDITED
        ResultLine = M_Main.ImportCsv_Silent("Trades", mTestFile, True)
        CheckText "replaced", ResultLine, "RESULT=LOADED; records=1; findings=0; file=" & mTestFile
        CheckText "status after the load", Meta(KIND_TR, 4), STAGING_LOADED
        CheckText "first ID", CStr(CellValue(KIND_TR, 1, COL_ID)), "T9"
        Check IsEmpty(CellValue(KIND_TR, 2, COL_LINE)) And IsEmpty(CellValue(KIND_TR, 2, COL_ID)) And _
              IsEmpty(CellValue(KIND_TR, 3, COL_ID)), "the old rows were not cleared"
        CheckText "records", Meta(KIND_TR, 3), "1"
        Application.EnableEvents = Saved
        Exit Sub

Unexpected:
        Fail mCurrentCase, "unexpected error " & Err.Number & ": " & Err.Description
        Application.EnableEvents = Saved

End Sub


Private Sub CaseWriteFailure()
'
'==============================================================================
' PURPOSE
'   A load that fails once the new block is written puts the block, the
'   metadata and the fingerprint back, and restores events.
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ResultLine    As String     'What the load returned
    Dim Before        As String     'The block before
    Dim MetaBefore    As String     'The metadata before
    Dim NameBefore    As String     'The fingerprint name before
    Dim EventsBefore  As Boolean    'EnableEvents before
    Dim ErrorNumber   As Long       'Error raised by the load

'------------------------------------------------------------------------------
' LOAD WITH A FAILURE
'------------------------------------------------------------------------------
        On Error GoTo Unexpected
        BeginCase "import.write-failure"
        Before = BlockState(KIND_TR)
        MetaBefore = MetaText(KIND_TR)
        NameBefore = NameText(STAGED_TR_NAME)
        EventsBefore = Application.EnableEvents
        WriteCsv CSV_TR_HEADER & vbLf & TradeRow("T7", "7", "") & vbLf & TradeRow("T8", "8", "") & vbLf
        M_Main.gTestFault = "staging-write"
        On Error Resume Next
        ResultLine = M_Main.ImportCsv_Silent("Trades", mTestFile, True)
        ErrorNumber = Err.Number
        On Error GoTo Unexpected
        M_Main.gTestFault = ""
        Check ErrorNumber = ERR_INJECTED_FAULT, "error " & ErrorNumber
        Check BlockState(KIND_TR) = Before, "the block was not put back"
        CheckText "metadata", MetaText(KIND_TR), MetaBefore
        CheckText "fingerprint name", NameText(STAGED_TR_NAME), NameBefore
        Check Application.EnableEvents = EventsBefore, "events were not restored"
        Exit Sub

Unexpected:
        Fail mCurrentCase, "unexpected error " & Err.Number & ": " & Err.Description
        M_Main.gTestFault = ""

End Sub


Private Sub CaseCellLimit()
'
'==============================================================================
' PURPOSE
'   A value longer than a cell holds (32,767 characters) raises
'   ERR_STAGING_WRITE and loads nothing.
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ResultLine    As String    'What the load returned
    Dim Before        As String    'The block before
    Dim ErrorNumber   As Long      'Error raised by the load

'------------------------------------------------------------------------------
' LOAD
'------------------------------------------------------------------------------
        On Error GoTo Unexpected
        BeginCase "import.cell-limit"
        Before = BlockState(KIND_TR)
        WriteCsv CSV_TR_HEADER & vbLf & TradeRow("TL", "1", String$(40000, "x")) & vbLf
        On Error Resume Next
        ResultLine = M_Main.ImportCsv_Silent("Trades", mTestFile, True)
        ErrorNumber = Err.Number
        On Error GoTo Unexpected
        Check ErrorNumber = ERR_STAGING_WRITE, "error " & ErrorNumber
        Check BlockState(KIND_TR) = Before, "the block changed"
        Exit Sub

Unexpected:
        Fail mCurrentCase, "unexpected error " & Err.Number & ": " & Err.Description

End Sub


Private Sub CaseFormulaCell()
'
'==============================================================================
' PURPOSE
'   A formula in a staging block is found by M_Main.StagingFormulaCell; a
'   load that fails while writing puts it back as a formula; once it is
'   removed, none is found.
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Target        As Range      'The cell given a formula
    Dim Saved         As Variant    'Its value
    Dim Events        As Boolean    'EnableEvents before the case
    Dim ResultLine    As String     'What the load returned
    Dim ErrorNumber   As Long       'Error raised by the load

'------------------------------------------------------------------------------
' WRITE A FORMULA, FAIL A LOAD, REMOVE IT
'------------------------------------------------------------------------------
        On Error GoTo Unexpected
        BeginCase "import.formula-cell"
        Events = Application.EnableEvents
        Application.EnableEvents = False
        Set Target = StagingCell(KIND_TR, 1, COL_TR_NOTIONAL)
        Saved = Target.Value2
        Target.NumberFormat = "General"
        Target.Formula = "=1+1"
        CheckText "formula found", M_Main.StagingFormulaCell("Trades"), Target.Address(False, False)
        WriteCsv CSV_TR_HEADER & vbLf & TradeRow("T6", "6", "") & vbLf
        M_Main.gTestFault = "staging-write"
        On Error Resume Next
        ResultLine = M_Main.ImportCsv_Silent("Trades", mTestFile, True)
        ErrorNumber = Err.Number
        On Error GoTo Unexpected
        M_Main.gTestFault = ""
        Check ErrorNumber = ERR_INJECTED_FAULT, "failed load: error " & ErrorNumber
        CheckText "formula after the failed load", M_Main.StagingFormulaCell("Trades"), Target.Address(False, False)
        Target.NumberFormat = "@"
        Target.Value2 = Saved
        CheckText "formula removed", M_Main.StagingFormulaCell("Trades"), ""
        Application.EnableEvents = Events
        Exit Sub

Unexpected:
        Fail mCurrentCase, "unexpected error " & Err.Number & ": " & Err.Description
        M_Main.gTestFault = ""
        Application.EnableEvents = Events

End Sub


Private Sub CaseHeaderOnly()
'
'==============================================================================
' PURPOSE
'   A file with only the header loads no record: the block is cleared and
'   the stored fingerprint is empty.
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ResultLine   As String    'What the load returned

'------------------------------------------------------------------------------
' LOAD
'------------------------------------------------------------------------------
        On Error GoTo Unexpected
        BeginCase "import.header-only"
        WriteCsv CSV_NS_HEADER & vbLf
        ResultLine = M_Main.ImportCsv_Silent("NettingSets", mTestFile, True)
        CheckText "result", ResultLine, "RESULT=LOADED; records=0; findings=0; file=" & mTestFile
        Check Len(BlockState(KIND_NS)) = 0, "the block was not cleared"
        CheckText "records", Meta(KIND_NS, 3), "0"
        CheckText "stored fingerprint", NameText(STAGED_NS_NAME), "="""""
        Exit Sub

Unexpected:
        Fail mCurrentCase, "unexpected error " & Err.Number & ": " & Err.Description

End Sub


Private Sub CaseLiveSheets( _
    ByVal LiveBefore As String)
'
'==============================================================================
' PURPOSE
'   After every load, the live inputs and Params and the stored run and
'   reconciliation fingerprints are as they were.
'==============================================================================
'
        On Error GoTo Unexpected
        BeginCase "import.live-sheets"
        CheckText "live inputs and run fingerprints", LiveInputsText(), LiveBefore
        Exit Sub

Unexpected:
        Fail mCurrentCase, "unexpected error " & Err.Number & ": " & Err.Description

End Sub


'
'------------------------------------------------------------------------------
'
'                                   HELPERS
'
'------------------------------------------------------------------------------
'

Private Function TradeRow( _
    ByVal TradeId As String, _
    ByVal Notional As String, _
    ByVal Comment As String) _
    As String
'
'==============================================================================
' PURPOSE
'   A Trades record of NS-T1: an interest-rate swap leg with the given ID,
'   notional and comment, as CSV text.
'==============================================================================
'
        TradeRow = TradeId & ";NS-T1;IR;;EUR;Linear;Long;;;;" & Notional & _
                   ";EUR;0;;;2031-01-02;2031-01-02;;;;;;;;;;;" & Comment

End Function


Private Sub WriteCsv( _
    ByVal Text As String)
'
'==============================================================================
' PURPOSE
'   Write Text to the test file as UTF-8, replacing it.
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Bytes()      As Byte       'The text's UTF-8 bytes
    Dim FileNumber   As Integer    'File number

'------------------------------------------------------------------------------
' WRITE
'------------------------------------------------------------------------------
        Bytes = CORE_Export.Utf8Bytes(Text)
        RemoveFile mTestFile
        FileNumber = FreeFile
        Open mTestFile For Binary Access Write As #FileNumber
        Put #FileNumber, 1, Bytes
        Close #FileNumber

End Sub


Private Sub RemoveFile( _
    ByVal FilePath As String)
'
'==============================================================================
' PURPOSE
'   Delete a test file when it exists.
'
' ERROR POLICY
'   Contained.
'==============================================================================
'
        On Error Resume Next
        If Len(FilePath) > 0 Then
            If Len(Dir$(FilePath, vbHidden)) > 0 Then
                SetAttr FilePath, vbNormal
                Kill FilePath
            End If
        End If
        Err.Clear

End Sub


Private Function StagingSheet( _
    ByVal KindIndex As Long) _
    As Worksheet
'
'==============================================================================
' PURPOSE
'   Import_NettingSets or Import_Trades.
'==============================================================================
'
        Set StagingSheet = ThisWorkbook.Worksheets(IIf(KindIndex = KIND_NS, SH_IMPORT_NS, SH_IMPORT_TR))

End Function


Private Function FirstDataRow( _
    ByVal KindIndex As Long) _
    As Long
'
'==============================================================================
' PURPOSE
'   The first row below a staging sheet's header row.
'==============================================================================
'
        FirstDataRow = NamedCell(IIf(KindIndex = KIND_NS, IMPORT_NS_HEADER_NAME, IMPORT_TR_HEADER_NAME)).Row + 1

End Function


Private Function BlockWidth( _
    ByVal KindIndex As Long) _
    As Long
'
'==============================================================================
' PURPOSE
'   The columns of a staging block: the CSV line and the header's fields.
'==============================================================================
'
        BlockWidth = UBound(Split(IIf(KindIndex = KIND_NS, CSV_NS_HEADER, CSV_TR_HEADER), ";")) + 2

End Function


Private Function StagingCell( _
    ByVal KindIndex As Long, _
    ByVal DataRow As Long, _
    ByVal Column As Long) _
    As Range
'
'==============================================================================
' PURPOSE
'   A cell of a staging block: data row 1 is the first below the header,
'   table column 1 the CSV line.
'==============================================================================
'
        Set StagingCell = TableCell(StagingSheet(KindIndex), FirstDataRow(KindIndex) + DataRow - 1, Column)

End Function


Private Function CellValue( _
    ByVal KindIndex As Long, _
    ByVal DataRow As Long, _
    ByVal Column As Long) _
    As Variant
'
'==============================================================================
' PURPOSE
'   The Value2 of a cell of a staging block.
'==============================================================================
'
        CellValue = StagingCell(KindIndex, DataRow, Column).Value2

End Function


Private Function BlockState( _
    ByVal KindIndex As Long) _
    As String
'
'==============================================================================
' PURPOSE
'   A staging block as text, as M_Main fingerprints it; "" when empty.
'==============================================================================
'
        BlockState = CORE_Engine.BlockText(StagingSheet(KindIndex), FirstDataRow(KindIndex), BlockWidth(KindIndex))

End Function


Private Function Meta( _
    ByVal KindIndex As Long, _
    ByVal Index As Long) _
    As String
'
'==============================================================================
' PURPOSE
'   A metadata value of a staging sheet: 1 file, 2 load time, 3 record
'   count, 4 status.
'==============================================================================
'
        Meta = CStr(NamedCell(IIf(KindIndex = KIND_NS, IMPORT_NS_PREFIX, IMPORT_TR_PREFIX) & "_" & _
                              Choose(Index, "File", "LoadedAt", "Records", "Status")).Value2)

End Function


Private Function MetaText( _
    ByVal KindIndex As Long) _
    As String
'
'==============================================================================
' PURPOSE
'   The four metadata values of a staging sheet, as one text.
'==============================================================================
'
        MetaText = Meta(KindIndex, 1) & "|" & Meta(KindIndex, 2) & "|" & Meta(KindIndex, 3) & "|" & Meta(KindIndex, 4)

End Function


Private Function NameText( _
    ByVal NameValue As String) _
    As String
'
'==============================================================================
' PURPOSE
'   What a workbook name refers to; "" when it does not exist.
'==============================================================================
'
        On Error Resume Next
        NameText = ThisWorkbook.Names(NameValue).RefersTo
        Err.Clear

End Function


Private Function LiveInputsText() As String
'
'==============================================================================
' PURPOSE
'   The fingerprint of the live inputs and Params, and the stored run,
'   output and reconciliation fingerprints, as one text.
'==============================================================================
'
        LiveInputsText = CORE_Engine.InputFingerprint() & "|" & NameText("SACCR_RunInputs") & "|" & _
                         NameText(RUN_OUTPUTS_NAME) & "|" & NameText(RECON_INPUTS_NAME) & "|" & _
                         NameText(TRADE_RECON_INPUTS_NAME)

End Function


Private Sub SaveStaging( _
    ByVal KindIndex As Long)
'
'==============================================================================
' PURPOSE
'   Keep a staging sheet's block, metadata and fingerprint name, to put
'   them back at the end.
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Sheet      As Worksheet    'The staging sheet
    Dim FirstRow   As Long         'Its first data row
    Dim LastRow    As Long         'Last row of its used range

'------------------------------------------------------------------------------
' KEEP
'------------------------------------------------------------------------------
        Set Sheet = StagingSheet(KindIndex)
        FirstRow = FirstDataRow(KindIndex)
        LastRow = UsedLastRow(Sheet)
        mSavedRows(KindIndex) = 0
        mSavedBlock(KindIndex) = Empty
        If LastRow >= FirstRow Then
            mSavedRows(KindIndex) = LastRow - FirstRow + 1
            mSavedBlock(KindIndex) = TableRange(Sheet, FirstRow, 1, LastRow, BlockWidth(KindIndex)).Value2
        End If
        mSavedMeta(KindIndex) = Array(Meta(KindIndex, 1), Meta(KindIndex, 2), Meta(KindIndex, 3), Meta(KindIndex, 4))
        mSavedName(KindIndex) = NameText(IIf(KindIndex = KIND_NS, STAGED_NS_NAME, STAGED_TR_NAME))
        mHadName(KindIndex) = Len(mSavedName(KindIndex)) > 0

End Sub


Private Sub PutStagingBack( _
    ByVal KindIndex As Long)
'
'==============================================================================
' PURPOSE
'   Put back what SaveStaging kept: the block as values, its data columns
'   formatted as text, the metadata and the fingerprint name.
'
' ERROR POLICY
'   A failure is recorded as a suite failure.
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Sheet       As Worksheet    'The staging sheet
    Dim FirstRow    As Long         'Its first data row
    Dim LastRow     As Long         'Last row of its used range
    Dim Width       As Long         'Block columns
    Dim NameValue   As String       'The fingerprint name
    Dim m           As Long         'Metadata value
    Dim Events      As Boolean      'EnableEvents before

'------------------------------------------------------------------------------
' PUT BACK
'------------------------------------------------------------------------------
        On Error GoTo Failed
        Events = Application.EnableEvents
        Application.EnableEvents = False
        Set Sheet = StagingSheet(KindIndex)
        FirstRow = FirstDataRow(KindIndex)
        Width = BlockWidth(KindIndex)
        LastRow = UsedLastRow(Sheet)
        If LastRow >= FirstRow Then
            TableRange(Sheet, FirstRow, 1, LastRow, Width).Clear
            TableRange(Sheet, FirstRow, 2, LastRow, Width).NumberFormat = "@"
        End If
        If mSavedRows(KindIndex) > 0 Then
            TableRange(Sheet, FirstRow, 2, FirstRow + mSavedRows(KindIndex) - 1, Width).NumberFormat = "@"
            TableRange(Sheet, FirstRow, 1, FirstRow + mSavedRows(KindIndex) - 1, Width).Value2 = mSavedBlock(KindIndex)
        End If
        For m = 1 To 4
            NamedCell(IIf(KindIndex = KIND_NS, IMPORT_NS_PREFIX, IMPORT_TR_PREFIX) & "_" & _
                      Choose(m, "File", "LoadedAt", "Records", "Status")).Value2 = mSavedMeta(KindIndex)(m - 1)
        Next m
        NameValue = IIf(KindIndex = KIND_NS, STAGED_NS_NAME, STAGED_TR_NAME)
        If mHadName(KindIndex) Then
            ThisWorkbook.Names.Add Name:=NameValue, RefersTo:=mSavedName(KindIndex), Visible:=False
        ElseIf Len(NameText(NameValue)) > 0 Then
            ThisWorkbook.Names(NameValue).Delete
        End If
        Application.EnableEvents = Events
        Exit Sub

Failed:
        Fail "suite", "staging sheet " & KindIndex & " not put back: " & Err.Number & ": " & Err.Description
        Application.EnableEvents = Events

End Sub


Private Sub CheckText( _
    ByVal CheckLabel As String, _
    ByVal Actual As String, _
    ByVal Expected As String)
'
'==============================================================================
' PURPOSE
'   Check that a text equals the expected one.
'==============================================================================
'
        Check Actual = Expected, CheckLabel & ": expected " & Expected & ", found " & Actual

End Sub


Private Sub BeginCase( _
    ByVal CaseName As String)
'
'==============================================================================
' PURPOSE
'   Count a case and print its CASE line.
'==============================================================================
'
        mCaseCount = mCaseCount + 1
        mCurrentCase = CaseName
        TEST_Suite.Say "CASE=" & CaseName

End Sub


Private Sub Check( _
    ByVal Passed As Boolean, _
    ByVal Detail As String)
'
'==============================================================================
' PURPOSE
'   Count one check and record it when it fails.
'==============================================================================
'
        mCheckCount = mCheckCount + 1
        If Not Passed Then
            Fail mCurrentCase, Detail
        End If

End Sub


Private Sub Fail( _
    ByVal CaseName As String, _
    ByVal Detail As String)
'
'==============================================================================
' PURPOSE
'   Count a failure and print it.
'==============================================================================
'
        mFailureCount = mFailureCount + 1
        TEST_Suite.Say "FAILURE=" & CaseName & ": " & Detail

End Sub
