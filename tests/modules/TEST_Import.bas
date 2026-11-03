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
'   in the block is found (M_Main.StagingFormulaCell), and a formula
'   whose value is the text it replaced is a change that refuses a new
'   load (#468). The live inputs,
'   the run fingerprints and the other staging sheet are never changed.
'   The staging validation (#209), M_Main.ValidateStaging_Silent, lists
'   each finding on Checks with the origin IMPORT and a link to its cell,
'   marks the cell and sets the status; once corrected, the marks go and
'   the status is VALID. With only netting sets staged, the legs on
'   Trades are checked against them and no live cell is marked. The
'   evidence export then refuses, and the run fingerprints are unchanged.
'   A formula in a staging block, or nothing staged, refuses the
'   validation and leaves Checks as it was. An Excel error value in a
'   staged cell is judged by the text Excel shows for it, without a VBA
'   error (#476). Against Params: a matured leg, a currency without a
'   rate and an OT leg under a BCBS override are findings on their cells;
'   a reporting date cleared on Params is one finding, on its cell there,
'   and no sheet is VALID; with only netting sets staged, a staged
'   override applies to the legs on Trades. The counters of the last run
'   are unchanged. M_Main.StagingValidationCurrent is True after a VALID
'   validation only while nothing it judged has changed. Validating the
'   same rows again gives the same lines and marks, without duplicates;
'   M_Main.RefreshStagingStatus shows OUT OF DATE on a VALID sheet once
'   what the last validation judged has changed, and VALID again once it
'   matches (#210). M_Main.CommitStaging_Silent refuses without a current
'   VALID validation; a failed write or read-back puts the input sheets
'   back and reports it verified, and a failed put-back is reported apart
'   from the error; a commit replaces each input sheet whose staging sheet
'   holds rows, with typed values and leading zeros kept, and leaves the
'   other as it was (#211). The suite puts the input sheets back.
'
' PUBLIC SURFACE
'   RunImportTests is the entry point. Option Private Module keeps it out
'   of the external workbook automation API.
'
' DEPENDENCIES
'   M_Main, with its test seam gTestFault, for the load, the validation
'   and the evidence export, and StagingValidationCurrent; CORE_Engine
'   for BlockText, InputFingerprint and the run counters; CORE_Util for the cells and TextHash;
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
'   written to the TEMP folder (K-SACCR_import_test.csv) and deleted.
'   The validation cases rewrite Checks; its lines are put back as
'   values, an ERROR line in bold, so that the run's outputs read as they
'   did, and the hidden name of the last VALID validation is put back.
'   Two cases change one Params value each, AsOfDate and Alpha, and put
'   it back as it was, also after an unexpected error; no other input
'   cell is touched, and import.live-sheets checks the live inputs and
'   the run names. Only then does the last case run SA-CCR (#211): it
'   puts the input sheets back and leaves a new run of them, CURRENT, or
'   no results when there were none before.
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
'   2026-10-31
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
        Private Const EXPECTED_CASES    As Long = 31                           'Cases in a complete run
        Private Const EXPECTED_CHECKS   As Long = 172                          'Checks in a complete run
        Private Const TEST_FILE         As String = "K-SACCR_import_test.csv"  'Test file in TEMP
        Private Const MISSING_FILE      As String = "K-SACCR_import_missing.csv"    'A file never written
        Private Const KIND_NS           As Long = 1                            'Import_NettingSets
        Private Const KIND_TR           As Long = 2                            'Import_Trades
    'Synthetic records: two netting sets, a third with the regime override
    'BCBS, and a trade leg of NS-T1 whose ID, notional and comment each
    'case chooses.
        Private Const NS_ROW_1          As String = "NS-T1;CPTY A;Y;N;1;N;N;;0;0;0;0;;;;;"
        Private Const NS_ROW_2          As String = "NS-T2;CPTY B;N;N;1;N;N;;0;0;0;0;;;;;"
        Private Const NS_ROW_BCBS       As String = "NS-T3;CPTY C;N;N;1;N;N;;0;0;0;0;;BCBS;;;"
    'Table columns of the staging block: the CSV line, then the contract's.
        Private Const COL_LINE          As Long = 1     'CSV line
        Private Const COL_ID            As Long = 2     'trade_id or netting_set_id
        Private Const COL_TR_NS         As Long = 3     'netting_set_id of a leg
        Private Const COL_TR_CLASS      As Long = 4     'asset_class
        Private Const COL_NS_REMARGIN   As Long = 6     'remargin_period_business_days
        Private Const COL_TR_NOTIONAL   As Long = 12    'notional_amount
        Private Const COL_TR_NCCY       As Long = 13    'notional_currency
        Private Const COL_TR_MATURITY   As Long = 18    'maturity_date
        Private Const COL_TR_COMMENT    As Long = 29    'comment
    'Fills of a staged value and of a cell with an error, as M_Main sets
    'them (STAGING_VALUE_FILL and STAGING_ERROR_FILL).
        Private Const VALUE_FILL        As Long = 13431551    'RGB(255, 242, 204)
        Private Const ERROR_FILL        As Long = 13551615    'RGB(255, 199, 206)
    'The reason the evidence export gives while Checks shows a staging
    'validation.
        Private Const EVIDENCE_REFUSED  As String = "RESULT=REFUSED; reason=Checks shows a staging validation, " & _
                                                    "not the checks of the last run; press Run SA-CCR first"

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
    'Checks as the suite found it, put back at the end.
        Private mSavedChecks         As Variant    'Its lines; Empty for none
        Private mSavedChecksRows     As Long       'Their number
        Private mChecksBefore        As String     'Its lines as text, compared at the end
    'The fingerprints of the last VALID staging validation and of the last
    'staging validation, put back at the end.
        Private mSavedValidated(1 To 2)  As String     'RefersTo of STAGING_VALIDATED_NAME, STAGING_CHECKED_NAME
        Private mHadValidated(1 To 2)    As Boolean    'The name existed
    'The input cells of NettingSets and Trades as the suite found them, put
    'back after the commit cases (#211).
        Private mSavedLive(1 To 2)       As Variant    'Their Value2 and formulas, as PutBackLive writes them
        Private mSavedLiveLast(1 To 2)   As Long       'Last row with a value; first row - 1 for none


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
'   2026-10-31
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
    Dim i                   As Long             'Validation name

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
        SaveChecks
        SaveLive
        For i = 1 To 2
            mSavedValidated(i) = NameText(ValidationName(i))
            mHadValidated(i) = Len(mSavedValidated(i)) > 0
        Next i

        CaseNettingSetsAccepted
        CaseTradesValueFindings
        CaseStructuralRejected
        CaseMissingFile
        CaseEdits
        CaseWriteFailure
        CaseCellLimit
        CaseFormulaCell
        CaseFormulaEqualValue
        CaseHeaderOnly
        CaseValidateStaged
        CaseValidateRepeated
        CaseValidateCorrected
        CaseValidateEvidence
        CaseValidateLiveTrades
        CaseValidateRefused
        CaseValidateErrorValues
        CaseValidateContext
        CaseValidateParams
        CaseValidateCurrent
        CaseValidateLiveContext
        CaseValidateOutOfDate
        CaseCommitRefused
        CaseCommitFailures
        CaseCommitRollbackFailed
        CaseCommitBoth
        CaseCommitTradesOnly
        CaseCommitTrailingFormula
        CaseValidateRecalc
        PutBackLive
        PutChecksBack
        CaseLiveSheets LiveBefore
        CaseCommitEquivalentRun

        PutStagingBack KIND_NS
        PutStagingBack KIND_TR
        PutValidatedBack
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
        CheckText "stored fingerprint", NameText(STAGED_NS_NAME), "=""" & StagedFingerprint(KIND_NS) & """"
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
'   is not a number, a comment that starts like a formula, and a blank
'   record twice, as an empty line and as delimiters only. The three
'   other records are loaded, on lines 2, 5 and 7, the second with a
'   quoted line break; neither blank record is (#467). Leading zeros and
'   the formula-like comment stay as written, and Import_NettingSets is
'   left as it was.
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
                 String$(UBound(Split(CSV_TR_HEADER, ";")), ";") & vbCrLf & _
                 TradeRow("T3", "1000", """two" & vbCrLf & "lines""") & vbCrLf & TradeRow("T4", "2000", "") & vbCrLf
        ResultLine = M_Main.ImportCsv_Silent("Trades", mTestFile, True)
        CheckText "result", ResultLine, "RESULT=LOADED; records=3; findings=4; file=" & mTestFile
        Check CellValue(KIND_TR, 1, COL_LINE) = 2 And CellValue(KIND_TR, 2, COL_LINE) = 5 And _
              CellValue(KIND_TR, 3, COL_LINE) = 7, "CSV lines"
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
'   one-record file and clears the rows below it; the CSV line column
'   stays locked and the staged and cleared values unlocked (#420).
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
        Check StagingCell(KIND_TR, 1, COL_LINE).Locked And Not StagingCell(KIND_TR, 1, COL_ID).Locked And _
              Not StagingCell(KIND_TR, 2, COL_ID).Locked, "the CSV line must stay locked, the staged and cleared values unlocked"
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


Private Sub CaseFormulaEqualValue()
'
'==============================================================================
' PURPOSE
'   A staged ID replaced, with events off, by a formula that returns the
'   same text leaves the values as loaded, yet is a change (#468): a new
'   load is refused and the formula kept. Once the value is typed back as
'   text, a new load goes ahead.
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ResultLine   As String     'What the load returned
    Dim Target       As Range      'The cell given a formula
    Dim Events       As Boolean    'EnableEvents before the case

'------------------------------------------------------------------------------
' LOAD, REPLACE A VALUE WITH AN EQUAL FORMULA, LOAD AGAIN
'------------------------------------------------------------------------------
        On Error GoTo Unexpected
        BeginCase "import.formula-equal-value"
        Events = Application.EnableEvents
        WriteCsv CSV_TR_HEADER & vbLf & TradeRow("T1", "1000", "") & vbLf
        M_Main.ImportCsv_Silent "Trades", mTestFile, True
        Application.EnableEvents = False
        Set Target = StagingCell(KIND_TR, 1, COL_ID)
        Target.NumberFormat = "General"
        Target.Formula = "=""T1"""
        Application.EnableEvents = Events
        CheckText "the formula's value", CStr(Target.Value2), "T1"
        ResultLine = M_Main.ImportCsv_Silent("Trades", mTestFile)
        CheckText "refused", ResultLine, "RESULT=REFUSED; reason=" & SH_IMPORT_TR & _
                  " has changes since its last load; file=" & mTestFile
        Check Target.HasFormula, "the formula was not kept"
        Application.EnableEvents = False
        Target.NumberFormat = "@"
        Target.Value2 = "T1"
        Application.EnableEvents = Events
        ResultLine = M_Main.ImportCsv_Silent("Trades", mTestFile)
        CheckText "loaded once typed back", ResultLine, "RESULT=LOADED; records=1; findings=0; file=" & mTestFile
        Exit Sub

Unexpected:
        Fail mCurrentCase, "unexpected error " & Err.Number & ": " & Err.Description
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


Private Sub CaseValidateStaged()
'
'==============================================================================
' PURPOSE
'   Two staged netting sets and three staged legs, one in an unknown
'   netting set and one with a notional that is not a number: the
'   validation is INVALID with two findings. Checks starts with an INFO
'   line, then one ERROR line per finding with the staging sheet, row, a
'   link to the cell and the origin IMPORT; both cells are marked, no
'   other; Import_NettingSets is VALID and Import_Trades INVALID.
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ResultLine    As String    'What the validation returned
    Dim Notional      As Range     'The notional that is not a number
    Dim NettingSet    As Range     'The unknown netting set

'------------------------------------------------------------------------------
' LOAD AND VALIDATE
'------------------------------------------------------------------------------
        On Error GoTo Unexpected
        BeginCase "validate.staged-rows"
        WriteCsv CSV_NS_HEADER & vbLf & NS_ROW_1 & vbLf & NS_ROW_2 & vbLf
        M_Main.ImportCsv_Silent "NettingSets", mTestFile, True
        WriteCsv CSV_TR_HEADER & vbLf & TradeRow("T1", "1000", "") & vbLf & _
                 Replace(TradeRow("T2", "1000", ""), ";NS-T1;", ";NS-T9;") & vbLf & TradeRow("T3", "abc", "") & vbLf
        M_Main.ImportCsv_Silent "Trades", mTestFile, True
        Set Notional = StagingCell(KIND_TR, 3, COL_TR_NOTIONAL)
        Set NettingSet = StagingCell(KIND_TR, 2, COL_TR_NS)
        ResultLine = M_Main.ValidateStaging_Silent()
        CheckText "result", ResultLine, "RESULT=INVALID; errors=2; nettingsets=2; trades=3; first=CSV-TYPE|" & _
                  SH_IMPORT_TR & "!" & Notional.Address(False, False)
        CheckText "Origin header", SafeStr(TableCell(GetSheet(SH_CHECKS), TableHeaderRow(SH_CHECKS), CK_ORIGIN_COL).Value2), _
                  "Origin"
        CheckText "summary line", ChecksValue(1, 1) & "|" & ChecksValue(1, CK_ORIGIN_COL), SEV_INFO & "|" & ORIGIN_IMPORT
        CheckText "first finding", ChecksValue(2, 1) & "|" & ChecksValue(2, 2) & "|" & ChecksValue(2, 3) & "|" & _
                  ChecksValue(2, 4) & "|" & ChecksValue(2, CK_ORIGIN_COL), SEV_ERROR & "|" & SH_IMPORT_TR & "|" & _
                  Notional.Row & "|" & Notional.Address(False, False) & "|" & ORIGIN_IMPORT
        Check ChecksCell(2, 4).HasFormula And InStr(1, ChecksCell(2, 4).Formula, "HYPERLINK", vbTextCompare) > 0, _
              "the Reference cell is not a link: " & ChecksCell(2, 4).Formula
        Check Left$(ChecksValue(2, 5), 26) = "CSV-TYPE: notional_amount:" And InStr(ChecksValue(2, 5), "(CSV line 4)") > 0, _
              "first message: " & ChecksValue(2, 5)
        CheckText "second finding", Left$(ChecksValue(3, 5), 23) & "|" & ChecksValue(3, 4), _
                  "CSV-NETTING-SET-UNKNOWN|" & NettingSet.Address(False, False)
        Check Len(ChecksValue(4, 1)) = 0, "a fourth Checks line: " & ChecksValue(4, 5)
        Check Notional.Interior.Color = ERROR_FILL And NettingSet.Interior.Color = ERROR_FILL, "the cells are not marked"
        Check StagingCell(KIND_TR, 1, COL_TR_NOTIONAL).Interior.Color = VALUE_FILL, "a cell without a finding is marked"
        CheckText "statuses", Meta(KIND_NS, 4) & "|" & Meta(KIND_TR, 4), STAGING_VALID & "|" & STAGING_INVALID
        Exit Sub

Unexpected:
        Fail mCurrentCase, "unexpected error " & Err.Number & ": " & Err.Description

End Sub


Private Sub CaseValidateRepeated()
'
'==============================================================================
' PURPOSE
'   The same staged rows validated again (#210): the same result, the
'   same two ERROR lines on Checks and no third, the same two cells
'   marked, and the same statuses.
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ResultLine    As String    'What the validation returned

'------------------------------------------------------------------------------
' VALIDATE AGAIN
'------------------------------------------------------------------------------
        On Error GoTo Unexpected
        BeginCase "validate.repeated"
        ResultLine = M_Main.ValidateStaging_Silent()
        CheckText "result", ResultLine, "RESULT=INVALID; errors=2; nettingsets=2; trades=3; first=CSV-TYPE|" & _
                  SH_IMPORT_TR & "!" & StagingCell(KIND_TR, 3, COL_TR_NOTIONAL).Address(False, False)
        Check ChecksValue(3, 1) = SEV_ERROR And Len(ChecksValue(4, 1)) = 0, "Checks lines: " & ChecksValue(4, 5)
        Check StagingCell(KIND_TR, 3, COL_TR_NOTIONAL).Interior.Color = ERROR_FILL And _
              StagingCell(KIND_TR, 2, COL_TR_NS).Interior.Color = ERROR_FILL And _
              StagingCell(KIND_TR, 1, COL_TR_NOTIONAL).Interior.Color = VALUE_FILL, "the marks changed"
        CheckText "statuses", Meta(KIND_NS, 4) & "|" & Meta(KIND_TR, 4), STAGING_VALID & "|" & STAGING_INVALID
        Exit Sub

Unexpected:
        Fail mCurrentCase, "unexpected error " & Err.Number & ": " & Err.Description

End Sub


Private Sub CaseValidateCorrected()
'
'==============================================================================
' PURPOSE
'   The two cells corrected on the sheet, with events on: Import_Trades
'   shows EDITED, then the validation is VALID; the marks are gone,
'   Checks has its INFO line only, without a link, and the cell that held
'   one has its column's font again, not the link style Excel gave it;
'   both sheets are VALID. A netting set is found whatever its case.
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ResultLine    As String     'What the validation returned
    Dim Events        As Boolean    'EnableEvents before the case
    Dim Model         As Range      'Last cell of the Reference column, with its font

'------------------------------------------------------------------------------
' CORRECT AND VALIDATE
'------------------------------------------------------------------------------
        On Error GoTo Unexpected
        BeginCase "validate.corrected"
        Events = Application.EnableEvents
        Application.EnableEvents = True
        StagingCell(KIND_TR, 3, COL_TR_NOTIONAL).Value2 = "1000"
        StagingCell(KIND_TR, 2, COL_TR_NS).Value2 = "ns-t2"
        Application.EnableEvents = Events
        CheckText "status after the corrections", Meta(KIND_TR, 4), STAGING_EDITED
        ResultLine = M_Main.ValidateStaging_Silent()
        CheckText "result", ResultLine, "RESULT=VALID; errors=0; nettingsets=2; trades=3"
        Check StagingCell(KIND_TR, 3, COL_TR_NOTIONAL).Interior.Color = VALUE_FILL And _
              StagingCell(KIND_TR, 2, COL_TR_NS).Interior.Color = VALUE_FILL, "a mark was left"
        CheckText "statuses", Meta(KIND_NS, 4) & "|" & Meta(KIND_TR, 4), STAGING_VALID & "|" & STAGING_VALID
        Check InStr(ChecksValue(1, 5), "No issues found.") > 0 And Len(ChecksValue(2, 1)) = 0, _
              "Checks: " & ChecksValue(1, 5) & " / " & ChecksValue(2, 5)
        Check Not ChecksCell(2, 4).HasFormula, "a link was left on Checks"
        Set Model = TableCell(GetSheet(SH_CHECKS), GetSheet(SH_CHECKS).Rows.Count, 4)
        With ChecksCell(2, 4).Font
            Check .Name = Model.Font.Name And .Size = Model.Font.Size And .Color = Model.Font.Color And _
                  .Underline = Model.Font.Underline, "the link's font was left: " & .Name & " " & .Size
        End With
        Exit Sub

Unexpected:
        Fail mCurrentCase, "unexpected error " & Err.Number & ": " & Err.Description
        Application.EnableEvents = Events

End Sub


Private Sub CaseValidateEvidence()
'
'==============================================================================
' PURPOSE
'   While Checks shows a staging validation, the evidence export refuses
'   and writes no file: its lines are not the checks of the last run.
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ResultLine    As String    'What the export returned

'------------------------------------------------------------------------------
' EXPORT
'------------------------------------------------------------------------------
        On Error GoTo Unexpected
        BeginCase "validate.evidence-refused"
        RemoveFile mTestFile
        ResultLine = M_Main.ExportEvidence_Silent(mTestFile)
        CheckText "export", ResultLine, EVIDENCE_REFUSED
        Check Len(Dir$(mTestFile)) = 0, "the export wrote a file"
        Exit Sub

Unexpected:
        Fail mCurrentCase, "unexpected error " & Err.Number & ": " & Err.Description

End Sub


Private Sub CaseValidateLiveTrades()
'
'==============================================================================
' PURPOSE
'   With netting sets staged and no legs, each leg on Trades is checked
'   against the staged netting sets: as many findings as legs whose
'   netting set is neither NS-T1 nor NS-T2, each on its Trades cell, which
'   is not marked; Import_NettingSets is INVALID when there is one.
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ResultLine    As String    'What the validation returned
    Dim Outside       As Long      'Legs on Trades outside the staged netting sets
    Dim FirstCell     As String    'Trades cell of the first of them
    Dim FillBefore    As Long      'Its fill before the validation

'------------------------------------------------------------------------------
' VALIDATE
'------------------------------------------------------------------------------
        On Error GoTo Unexpected
        BeginCase "validate.live-trades"
        WriteCsv CSV_TR_HEADER & vbLf
        M_Main.ImportCsv_Silent "Trades", mTestFile, True
        Outside = LegsOutside("NS-T1|NS-T2", FirstCell)
        If Outside > 0 Then
            FillBefore = GetSheet(SH_TRADES).Range(FirstCell).Interior.Color
        End If
        ResultLine = M_Main.ValidateStaging_Silent()
        If Outside = 0 Then
            CheckText "result", ResultLine, "RESULT=VALID; errors=0; nettingsets=2; trades=0"
            CheckText "status", Meta(KIND_NS, 4), STAGING_VALID
            Check Len(ChecksValue(2, 1)) = 0, "a finding without a leg outside"
        Else
            CheckText "result", ResultLine, "RESULT=INVALID; errors=" & Outside & "; nettingsets=2; trades=0; " & _
                      "first=CSV-NETTING-SET-UNKNOWN|" & SH_TRADES & "!" & FirstCell
            CheckText "status", Meta(KIND_NS, 4), STAGING_INVALID
            CheckText "first finding", ChecksValue(2, 2) & "|" & ChecksValue(2, 4) & "|" & _
                      ChecksValue(2, CK_ORIGIN_COL) & "|" & GetSheet(SH_TRADES).Range(FirstCell).Interior.Color, _
                      SH_TRADES & "|" & FirstCell & "|" & ORIGIN_IMPORT & "|" & FillBefore
        End If
        Exit Sub

Unexpected:
        Fail mCurrentCase, "unexpected error " & Err.Number & ": " & Err.Description

End Sub


Private Sub CaseValidateRefused()
'
'==============================================================================
' PURPOSE
'   A formula in a staging block refuses the validation, naming its cell;
'   so does nothing staged. Checks is left as it was.
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ResultLine    As String     'What the validation returned
    Dim Before        As String     'Checks before
    Dim Target        As Range      'The cell given a formula
    Dim Events        As Boolean    'EnableEvents before the case

'------------------------------------------------------------------------------
' REFUSE
'------------------------------------------------------------------------------
        On Error GoTo Unexpected
        BeginCase "validate.refused"
        Before = ChecksText()
        Events = Application.EnableEvents
        Application.EnableEvents = False
        Set Target = StagingCell(KIND_NS, 1, COL_NS_REMARGIN)
        Target.NumberFormat = "General"
        Target.Formula = "=1"
        ResultLine = M_Main.ValidateStaging_Silent()
        CheckText "formula", ResultLine, "RESULT=REFUSED; reason=" & SH_IMPORT_NS & "!" & Target.Address(False, False) & _
                  " holds a formula"
        Target.NumberFormat = "@"
        Target.Value2 = "1"
        Application.EnableEvents = Events
        WriteCsv CSV_NS_HEADER & vbLf
        M_Main.ImportCsv_Silent "NettingSets", mTestFile, True
        ResultLine = M_Main.ValidateStaging_Silent()
        CheckText "nothing staged", ResultLine, "RESULT=REFUSED; reason=nothing is staged"
        CheckText "Checks unchanged", ChecksText(), Before
        Exit Sub

Unexpected:
        Fail mCurrentCase, "unexpected error " & Err.Number & ": " & Err.Description
        Application.EnableEvents = Events

End Sub


Private Sub CaseValidateErrorValues()
'
'==============================================================================
' PURPOSE
'   Excel error values in staged cells, as a paste of values can leave
'   them (#476): on T1, #N/A in the notional, a number column, and
'   #DIV/0! in the comment, a text column; on T2, #VALUE! in the notional
'   and #N/A in the comment. The validation runs without a VBA error and
'   is INVALID with two findings, each notional's CSV-TYPE naming its
'   error; the comments are judged as the text Excel shows.
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ResultLine    As String     'What the validation returned
    Dim Notional      As Range      'The notional of T1, holding #N/A
    Dim Notional2     As Range      'The notional of T2, holding #VALUE!
    Dim Events        As Boolean    'EnableEvents before the case
    Dim ErrorNumber   As Long       'Error raised by the validation

'------------------------------------------------------------------------------
' STAGE, PUT ERROR VALUES IN, VALIDATE
'------------------------------------------------------------------------------
        On Error GoTo Unexpected
        BeginCase "validate.error-values"
        Events = Application.EnableEvents
        WriteCsv CSV_NS_HEADER & vbLf & NS_ROW_1 & vbLf & NS_ROW_2 & vbLf
        M_Main.ImportCsv_Silent "NettingSets", mTestFile, True
        WriteCsv CSV_TR_HEADER & vbLf & TradeRow("T1", "1000", "") & vbLf & TradeRow("T2", "1000", "") & vbLf
        M_Main.ImportCsv_Silent "Trades", mTestFile, True
        Application.EnableEvents = False
        Set Notional = StagingCell(KIND_TR, 1, COL_TR_NOTIONAL)
        Set Notional2 = StagingCell(KIND_TR, 2, COL_TR_NOTIONAL)
        Notional.Value = CVErr(xlErrNA)
        StagingCell(KIND_TR, 1, COL_TR_COMMENT).Value = CVErr(xlErrDiv0)
        Notional2.Value = CVErr(xlErrValue)
        StagingCell(KIND_TR, 2, COL_TR_COMMENT).Value = CVErr(xlErrNA)
        Application.EnableEvents = Events
        Check IsError(Notional.Value2) And IsError(StagingCell(KIND_TR, 1, COL_TR_COMMENT).Value2) And _
              IsError(Notional2.Value2) And IsError(StagingCell(KIND_TR, 2, COL_TR_COMMENT).Value2), _
              "the cells do not hold error values"
        On Error Resume Next
        ResultLine = M_Main.ValidateStaging_Silent()
        ErrorNumber = Err.Number
        On Error GoTo Unexpected
        Check ErrorNumber = 0, "the validation raised error " & ErrorNumber
        CheckText "result", ResultLine, "RESULT=INVALID; errors=2; nettingsets=2; trades=2; first=CSV-TYPE|" & _
                  SH_IMPORT_TR & "!" & Notional.Address(False, False)
        Check InStr(1, ChecksValue(2, 5), "#N/A", vbBinaryCompare) > 0, "the finding does not name #N/A: " & ChecksValue(2, 5)
        Check InStr(1, ChecksValue(3, 5), "#VALUE!", vbBinaryCompare) > 0, "the finding does not name #VALUE!: " & ChecksValue(3, 5)
        CheckText "status", Meta(KIND_TR, 4), STAGING_INVALID
        Exit Sub

Unexpected:
        Fail mCurrentCase, "unexpected error " & Err.Number & ": " & Err.Description
        Application.EnableEvents = Events

End Sub


Private Sub CaseValidateContext()
'
'==============================================================================
' PURPOSE
'   Staged legs judged against Params (#209): T1 is valid; T2 has
'   matured, its maturity not after the reporting date; T3 is in SEK,
'   which has no FX rate; T4 is an OT leg of NS-T3, whose override is
'   BCBS. The validation is INVALID with three findings, in that order,
'   each on its staged cell, which is marked; Import_NettingSets is VALID
'   and Import_Trades INVALID; the counters of the last run are as they
'   were; and an INVALID validation does not read as current.
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ResultLine      As String    'What the validation returned
    Dim Maturity        As Range     'Maturity of T2
    Dim NotionalCcy     As Range     'Notional currency of T3
    Dim AssetClass      As Range     'Asset class of T4
    Dim CountersBefore  As String    'Errors, warnings and trades read of the last run

'------------------------------------------------------------------------------
' LOAD AND VALIDATE
'------------------------------------------------------------------------------
        On Error GoTo Unexpected
        BeginCase "validate.context"
        WriteCsv CSV_NS_HEADER & vbLf & NS_ROW_1 & vbLf & NS_ROW_2 & vbLf & NS_ROW_BCBS & vbLf
        M_Main.ImportCsv_Silent "NettingSets", mTestFile, True
        WriteCsv CSV_TR_HEADER & vbLf & TradeRow("T1", "1000", "") & vbLf & _
                 Replace(TradeRow("T2", "1000", ""), ";2031-01-02;2031-01-02;", ";2026-01-02;2026-01-02;") & vbLf & _
                 Replace(TradeRow("T3", "1000", ""), ";1000;EUR;", ";1000;SEK;") & vbLf & _
                 Replace(TradeRow("T4", "1000", ""), ";NS-T1;IR;;EUR;", ";NS-T3;OT;;LONGEVITY;") & vbLf
        M_Main.ImportCsv_Silent "Trades", mTestFile, True
        Set Maturity = StagingCell(KIND_TR, 2, COL_TR_MATURITY)
        Set NotionalCcy = StagingCell(KIND_TR, 3, COL_TR_NCCY)
        Set AssetClass = StagingCell(KIND_TR, 4, COL_TR_CLASS)
        CountersBefore = RunCounters()
        ResultLine = M_Main.ValidateStaging_Silent()
        CheckText "result", ResultLine, "RESULT=INVALID; errors=3; nettingsets=3; trades=4; first=CSV-CONTEXT-MATURED|" & _
                  SH_IMPORT_TR & "!" & Maturity.Address(False, False)
        CheckText "findings", FindingCode(2) & "|" & ChecksValue(2, 4) & "|" & FindingCode(3) & "|" & ChecksValue(3, 4) & _
                  "|" & FindingCode(4) & "|" & ChecksValue(4, 4), "CSV-CONTEXT-MATURED|" & Maturity.Address(False, False) & _
                  "|CSV-CONTEXT-CURRENCY|" & NotionalCcy.Address(False, False) & "|CSV-CONTEXT-REGIME|" & _
                  AssetClass.Address(False, False)
        Check Len(ChecksValue(5, 1)) = 0, "a fifth Checks line: " & ChecksValue(5, 5)
        Check Maturity.Interior.Color = ERROR_FILL And NotionalCcy.Interior.Color = ERROR_FILL And _
              AssetClass.Interior.Color = ERROR_FILL, "the cells are not marked"
        CheckText "statuses", Meta(KIND_NS, 4) & "|" & Meta(KIND_TR, 4), STAGING_VALID & "|" & STAGING_INVALID
        CheckText "counters of the last run", RunCounters(), CountersBefore
        Check Not M_Main.StagingValidationCurrent(), "an INVALID validation reads as current"
        Exit Sub

Unexpected:
        Fail mCurrentCase, "unexpected error " & Err.Number & ": " & Err.Description

End Sub


Private Sub CaseValidateParams()
'
'==============================================================================
' PURPOSE
'   With the reporting date on Params cleared, a run would stop at Params
'   (#209): the validation is INVALID with one finding, CSV-CONTEXT-
'   PARAMS, which names AsOfDate and links to its cell on Params, which is
'   not marked; neither staging sheet is VALID. The date is put back, also
'   after an unexpected error.
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ResultLine      As String     'What the validation returned
    Dim AsOfCell        As Range      'The reporting date on Params
    Dim SavedFormula    As String     'What it held
    Dim FillBefore      As Long       'Its fill
    Dim Cleared         As Boolean    'It was cleared and must be put back

'------------------------------------------------------------------------------
' CLEAR THE DATE AND VALIDATE
'------------------------------------------------------------------------------
        On Error GoTo Unexpected
        BeginCase "validate.params-invalid"
        WriteCsv CSV_NS_HEADER & vbLf & NS_ROW_1 & vbLf & NS_ROW_2 & vbLf
        M_Main.ImportCsv_Silent "NettingSets", mTestFile, True
        WriteCsv CSV_TR_HEADER & vbLf & TradeRow("T1", "1000", "") & vbLf
        M_Main.ImportCsv_Silent "Trades", mTestFile, True
        Set AsOfCell = NamedCell(PRM_ASOF)
        SavedFormula = AsOfCell.Formula
        FillBefore = AsOfCell.Interior.Color
        Cleared = True
        AsOfCell.ClearContents
        ResultLine = M_Main.ValidateStaging_Silent()
        AsOfCell.Formula = SavedFormula
        Cleared = False
        CheckText "result", ResultLine, "RESULT=INVALID; errors=1; nettingsets=2; trades=1; first=CSV-CONTEXT-PARAMS|" & _
                  SH_PARAMS & "!" & AsOfCell.Address(False, False)
        CheckText "finding", FindingCode(2) & "|" & ChecksValue(2, 2) & "|" & ChecksValue(2, 4) & "|" & _
                  AsOfCell.Interior.Color, "CSV-CONTEXT-PARAMS|" & SH_PARAMS & "|" & AsOfCell.Address(False, False) & _
                  "|" & FillBefore
        Check InStr(1, ChecksValue(2, 5), PRM_ASOF, vbBinaryCompare) > 0, "the finding does not name " & PRM_ASOF & _
              ": " & ChecksValue(2, 5)
        CheckText "statuses", Meta(KIND_NS, 4) & "|" & Meta(KIND_TR, 4), STAGING_INVALID & "|" & STAGING_INVALID
        Exit Sub

Unexpected:
        Fail mCurrentCase, "unexpected error " & Err.Number & ": " & Err.Description
        If Cleared Then
            AsOfCell.Formula = SavedFormula
        End If

End Sub


Private Sub CaseValidateCurrent()
'
'==============================================================================
' PURPOSE
'   M_Main.StagingValidationCurrent (#209): True after a VALID validation;
'   False once a staged cell changes, even with events off, and True
'   again once it is put back; False while a Params value is changed; and
'   False after an INVALID validation, even once its cause is put back.
'   The Params value is put back, also after an unexpected error.
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ResultLine      As String     'What a validation returned
    Dim Events          As Boolean    'EnableEvents before the case
    Dim AlphaCell       As Range      'Alpha on Params
    Dim SavedFormula    As String     'What it held
    Dim Changed         As Boolean    'It was changed and must be put back

'------------------------------------------------------------------------------
' VALIDATE, CHANGE AND PUT BACK
'------------------------------------------------------------------------------
        On Error GoTo Unexpected
        BeginCase "validate.current"
        Events = Application.EnableEvents
        WriteCsv CSV_NS_HEADER & vbLf & NS_ROW_1 & vbLf & NS_ROW_2 & vbLf
        M_Main.ImportCsv_Silent "NettingSets", mTestFile, True
        WriteCsv CSV_TR_HEADER & vbLf & TradeRow("T1", "1000", "") & vbLf
        M_Main.ImportCsv_Silent "Trades", mTestFile, True
        ResultLine = M_Main.ValidateStaging_Silent()
        CheckText "result", ResultLine, "RESULT=VALID; errors=0; nettingsets=2; trades=1"
        Check M_Main.StagingValidationCurrent(), "a VALID validation does not read as current"
        Application.EnableEvents = False
        StagingCell(KIND_TR, 1, COL_TR_COMMENT).Value2 = "edited"
        Check Not M_Main.StagingValidationCurrent(), "an edit with events off reads as current"
        StagingCell(KIND_TR, 1, COL_TR_COMMENT).ClearContents
        Check M_Main.StagingValidationCurrent(), "the edit put back does not read as current"
        Set AlphaCell = NamedCell(PRM_ALPHA)
        SavedFormula = AlphaCell.Formula
        Changed = True
        AlphaCell.Value2 = 1.5
        Check Not M_Main.StagingValidationCurrent(), "a changed Params value reads as current"
        AlphaCell.Formula = SavedFormula
        Changed = False
        StagingCell(KIND_TR, 1, COL_TR_NOTIONAL).Value2 = "abc"
        ResultLine = M_Main.ValidateStaging_Silent()
        StagingCell(KIND_TR, 1, COL_TR_NOTIONAL).Value2 = "1000"
        Application.EnableEvents = Events
        Check Not M_Main.StagingValidationCurrent(), "after an INVALID validation, its cause put back, it reads as current"
        Exit Sub

Unexpected:
        Fail mCurrentCase, "unexpected error " & Err.Number & ": " & Err.Description
        If Changed Then
            AlphaCell.Formula = SavedFormula
        End If
        Application.EnableEvents = Events

End Sub


Private Sub CaseValidateLiveContext()
'
'==============================================================================
' PURPOSE
'   With only netting sets staged, a staged regime override applies to the
'   legs on Trades (#209): CRR-OTHER staged with the override BCBS makes
'   each of its OT legs a CSV-CONTEXT-REGIME finding on its Asset class
'   cell on Trades, which is not marked; the other legs are outside the
'   staged netting sets, as in CaseValidateLiveTrades.
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ResultLine    As String    'What the validation returned
    Dim Outside       As Long      'Legs on Trades outside the staged netting set
    Dim FirstCell     As String    'Netting set cell of the first of them, unused
    Dim OtLegs        As Long      'OT legs of CRR-OTHER on Trades
    Dim OtCell        As String    'Asset class cell of the first of them
    Dim FillBefore    As Long      'Its fill before the validation
    Dim FoundLine     As Long      'The Checks line of its finding; 0 for none
    Dim i             As Long      'Checks line

'------------------------------------------------------------------------------
' VALIDATE
'------------------------------------------------------------------------------
        On Error GoTo Unexpected
        BeginCase "validate.live-context"
        WriteCsv CSV_NS_HEADER & vbLf & "CRR-OTHER;CPTY D;N;N;1;N;N;;0;0;0;0;;BCBS;;;" & vbLf
        M_Main.ImportCsv_Silent "NettingSets", mTestFile, True
        WriteCsv CSV_TR_HEADER & vbLf
        M_Main.ImportCsv_Silent "Trades", mTestFile, True
        Outside = LegsOutside("CRR-OTHER", FirstCell)
        OtLegs = LegsOfClass("CRR-OTHER", "OT", OtCell)
        Check OtLegs > 0, "Trades has no OT leg of CRR-OTHER: build the workbook from the current template"
        If OtLegs > 0 Then
            FillBefore = GetSheet(SH_TRADES).Range(OtCell).Interior.Color
        End If
        ResultLine = M_Main.ValidateStaging_Silent()
        CheckText "result", Left$(ResultLine, InStr(ResultLine & "; first=", "; first=") - 1), _
                  "RESULT=INVALID; errors=" & (Outside + OtLegs) & "; nettingsets=1; trades=0"
        For i = 2 To Outside + OtLegs + 1
            If FindingCode(i) = "CSV-CONTEXT-REGIME" And ChecksValue(i, 4) = OtCell Then
                FoundLine = i
                Exit For
            End If
        Next i
        Check FoundLine > 0 And ChecksValue(FoundLine, 2) = SH_TRADES, "no CSV-CONTEXT-REGIME finding on " & _
              SH_TRADES & "!" & OtCell
        If OtLegs > 0 Then
            Check GetSheet(SH_TRADES).Range(OtCell).Interior.Color = FillBefore, "the Trades cell is marked"
        Else
            Check False, "no OT leg whose cell could be checked"
        End If
        Exit Sub

Unexpected:
        Fail mCurrentCase, "unexpected error " & Err.Number & ": " & Err.Description

End Sub


Private Sub CaseValidateOutOfDate()
'
'==============================================================================
' PURPOSE
'   M_Main.RefreshStagingStatus, which runs when a staging sheet is
'   activated (#210): after a VALID validation both sheets read VALID; a
'   changed Params value or an edit made with events off makes them read
'   OUT OF DATE, and putting it back makes them read VALID again. After
'   an INVALID validation the statuses stay as it set them until
'   something changes: then the sheet that read VALID reads OUT OF DATE,
'   and the one that reads EDITED keeps it. The Params value is put back,
'   also after an unexpected error.
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ResultLine      As String     'What a validation returned
    Dim Events          As Boolean    'EnableEvents before the case
    Dim AlphaCell       As Range      'Alpha on Params
    Dim SavedFormula    As String     'What it held
    Dim Changed         As Boolean    'It was changed and must be put back

'------------------------------------------------------------------------------
' VALIDATE, CHANGE, REFRESH
'------------------------------------------------------------------------------
        On Error GoTo Unexpected
        BeginCase "validate.out-of-date"
        Events = Application.EnableEvents
        WriteCsv CSV_NS_HEADER & vbLf & NS_ROW_1 & vbLf & NS_ROW_2 & vbLf
        M_Main.ImportCsv_Silent "NettingSets", mTestFile, True
        WriteCsv CSV_TR_HEADER & vbLf & TradeRow("T1", "1000", "") & vbLf
        M_Main.ImportCsv_Silent "Trades", mTestFile, True
        ResultLine = M_Main.ValidateStaging_Silent()
        CheckText "result", ResultLine, "RESULT=VALID; errors=0; nettingsets=2; trades=1"
        M_Main.RefreshStagingStatus
        CheckText "after the validation", Meta(KIND_NS, 4) & "|" & Meta(KIND_TR, 4), STAGING_VALID & "|" & STAGING_VALID
        Set AlphaCell = NamedCell(PRM_ALPHA)
        SavedFormula = AlphaCell.Formula
        Changed = True
        AlphaCell.Value2 = 1.5
        M_Main.RefreshStagingStatus
        CheckText "Params changed", Meta(KIND_NS, 4) & "|" & Meta(KIND_TR, 4), STAGING_OUT_OF_DATE & "|" & STAGING_OUT_OF_DATE
        AlphaCell.Formula = SavedFormula
        Changed = False
        M_Main.RefreshStagingStatus
        CheckText "Params put back", Meta(KIND_NS, 4) & "|" & Meta(KIND_TR, 4), STAGING_VALID & "|" & STAGING_VALID
        Application.EnableEvents = False
        StagingCell(KIND_TR, 1, COL_TR_COMMENT).Value2 = "edited"
        M_Main.RefreshStagingStatus
        CheckText "edit with events off", Meta(KIND_NS, 4) & "|" & Meta(KIND_TR, 4), STAGING_OUT_OF_DATE & "|" & STAGING_OUT_OF_DATE
        StagingCell(KIND_TR, 1, COL_TR_COMMENT).ClearContents
        M_Main.RefreshStagingStatus
        CheckText "edit put back", Meta(KIND_NS, 4) & "|" & Meta(KIND_TR, 4), STAGING_VALID & "|" & STAGING_VALID
        StagingCell(KIND_TR, 1, COL_TR_NOTIONAL).Value2 = "abc"
        ResultLine = M_Main.ValidateStaging_Silent()
        M_Main.RefreshStagingStatus
        CheckText "after an INVALID validation", Meta(KIND_NS, 4) & "|" & Meta(KIND_TR, 4), STAGING_VALID & "|" & STAGING_INVALID
        Application.EnableEvents = True
        StagingCell(KIND_TR, 1, COL_TR_NOTIONAL).Value2 = "1000"
        Application.EnableEvents = Events
        M_Main.RefreshStagingStatus
        CheckText "its cause corrected", Meta(KIND_NS, 4) & "|" & Meta(KIND_TR, 4), STAGING_OUT_OF_DATE & "|" & STAGING_EDITED
        Exit Sub

Unexpected:
        Fail mCurrentCase, "unexpected error " & Err.Number & ": " & Err.Description
        If Changed Then
            AlphaCell.Formula = SavedFormula
        End If
        Application.EnableEvents = Events

End Sub


Private Sub CaseCommitRefused()
'
'==============================================================================
' PURPOSE
'   M_Main.CommitStaging_Silent refuses without a current VALID staging
'   validation (#211): after a load without a validation, and after a
'   VALID validation followed by an edit made with events off. The input
'   sheets and the run's names are unchanged.
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ResultLine    As String     'What a call returned
    Dim Events        As Boolean    'EnableEvents before the case
    Dim Before        As String     'Input sheets and run names before

'------------------------------------------------------------------------------
' REFUSE
'------------------------------------------------------------------------------
        On Error GoTo Unexpected
        BeginCase "commit.refused"
        Events = Application.EnableEvents
        WriteCsv CSV_NS_HEADER & vbLf & NS_ROW_1 & vbLf & NS_ROW_2 & vbLf
        M_Main.ImportCsv_Silent "NettingSets", mTestFile, True
        WriteCsv CSV_TR_HEADER & vbLf & TradeRow("T1", "1000", "") & vbLf
        M_Main.ImportCsv_Silent "Trades", mTestFile, True
        Before = LiveInputsText()
        ResultLine = M_Main.CommitStaging_Silent()
        CheckText "without a validation", ResultLine, "RESULT=REFUSED; reason=no current VALID staging validation"
        ResultLine = M_Main.ValidateStaging_Silent()
        CheckText "validation", ResultLine, "RESULT=VALID; errors=0; nettingsets=2; trades=1"
        Application.EnableEvents = False
        StagingCell(KIND_TR, 1, COL_TR_COMMENT).Value2 = "edited"
        ResultLine = M_Main.CommitStaging_Silent()
        StagingCell(KIND_TR, 1, COL_TR_COMMENT).ClearContents
        Application.EnableEvents = Events
        CheckText "after an edit with events off", ResultLine, "RESULT=REFUSED; reason=no current VALID staging validation"
        Check LiveInputsText() = Before, "the input sheets or the run's names changed"
        Exit Sub

Unexpected:
        Fail mCurrentCase, "unexpected error " & Err.Number & ": " & Err.Description
        Application.EnableEvents = Events

End Sub


Private Sub CaseCommitFailures()
'
'==============================================================================
' PURPOSE
'   A commit that fails once NettingSets is written, or whose read-back
'   finds a difference, puts both input sheets back and reports
'   rollback=VERIFIED (#211): the inputs and the run's names are as they
'   were, a counterparty entered as the text =1+1 is text again, not a
'   formula, both staging sheets still read VALID and the validation is
'   still current. Continues from CaseCommitRefused; the suite then puts
'   the input sheets back and validates again.
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ResultLine    As String     'What a call returned
    Dim Before        As String     'Input sheets and run names before
    Dim Events        As Boolean    'EnableEvents before the case
    Dim Counterparty  As Range      'The first counterparty on NettingSets

'------------------------------------------------------------------------------
' FAIL AND PUT BACK
'------------------------------------------------------------------------------
        On Error GoTo Unexpected
        BeginCase "commit.failures"
        Events = Application.EnableEvents
        Application.EnableEvents = False
        Set Counterparty = TableCell(GetSheet(SH_NS), TableFirstRow(SH_NS), NS_CPTY)
        Counterparty.Value = "'=1+1"
        Application.EnableEvents = Events
        ResultLine = M_Main.ValidateStaging_Silent()
        CheckText "validation", ResultLine, "RESULT=VALID; errors=0; nettingsets=2; trades=1"
        Before = LiveInputsText()
        M_Main.gTestFault = "commit-write"
        ResultLine = M_Main.CommitStaging_Silent()
        M_Main.gTestFault = ""
        CheckText "write failure", ResultLine, "RESULT=FAILED; error=Injected commit write failure.; rollback=VERIFIED"
        Check LiveInputsText() = Before, "write failure: the input sheets were not put back"
        CheckText "statuses", Meta(KIND_NS, 4) & "|" & Meta(KIND_TR, 4), STAGING_VALID & "|" & STAGING_VALID
        Check M_Main.StagingValidationCurrent(), "write failure: the validation no longer reads as current"
        M_Main.gTestFault = "commit-verify"
        ResultLine = M_Main.CommitStaging_Silent()
        M_Main.gTestFault = ""
        CheckText "read-back failure", ResultLine, _
                  "RESULT=FAILED; error=Injected read-back difference on " & SH_NS & "; rollback=VERIFIED"
        Check LiveInputsText() = Before, "read-back failure: the input sheets were not put back"
        Check Not Counterparty.HasFormula And SafeStr(Counterparty.Value2) = "=1+1", _
              "the text =1+1 did not stay text: " & Counterparty.Formula
        PutBackLive
        ResultLine = M_Main.ValidateStaging_Silent()
        CheckText "validation again", ResultLine, "RESULT=VALID; errors=0; nettingsets=2; trades=1"
        Exit Sub

Unexpected:
        M_Main.gTestFault = ""
        Application.EnableEvents = Events
        Fail mCurrentCase, "unexpected error " & Err.Number & ": " & Err.Description
        PutBackLive

End Sub


Private Sub CaseCommitRollbackFailed()
'
'==============================================================================
' PURPOSE
'   A commit whose write-back fails too reports rollback=FAILED with what
'   failed, apart from the error itself, and leaves both staging sheets
'   VALID, not COMMITTED (#211). The suite then puts the input sheets
'   back itself.
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ResultLine    As String    'What the call returned
    Dim Before        As String    'Input sheets and run names before

'------------------------------------------------------------------------------
' FAIL TWICE
'------------------------------------------------------------------------------
        On Error GoTo Unexpected
        BeginCase "commit.rollback-failed"
        Before = LiveInputsText()
        M_Main.gTestFault = "commit-rollback"
        ResultLine = M_Main.CommitStaging_Silent()
        M_Main.gTestFault = ""
        CheckText "result", ResultLine, "RESULT=FAILED; error=Injected commit write failure.; rollback=FAILED: " & _
                  SH_NS & ": Injected write-back failure."
        CheckText "statuses", Meta(KIND_NS, 4) & "|" & Meta(KIND_TR, 4), STAGING_VALID & "|" & STAGING_VALID
        PutBackLive
        Check LiveInputsText() = Before, "the suite could not put the input sheets back"
        Exit Sub

Unexpected:
        M_Main.gTestFault = ""
        Fail mCurrentCase, "unexpected error " & Err.Number & ": " & Err.Description

End Sub


Private Sub CaseCommitBoth()
'
'==============================================================================
' PURPOSE
'   Two staged netting sets and three staged legs, validated, then
'   committed (#211): NettingSets and Trades hold exactly the staged rows,
'   the rows below cleared; an ID with leading zeros and a comment that
'   looks like a number stay text, a notional is a number and a maturity
'   an Excel date. Both staging sheets read COMMITTED, the validation is
'   no longer current and a second commit is refused. Before the commit,
'   the Commit import button in silent mode, where no confirmation can be
'   given, commits nothing and says so. The input sheets stay protected,
'   the committed cells unlocked and the headers locked (#422). The suite
'   then puts the input sheets back.
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ResultLine    As String       'What a call returned
    Dim Before        As String       'Input sheets and run names before
    Dim NsSheet       As Worksheet    'NettingSets
    Dim TrSheet       As Worksheet    'Trades
    Dim NsFirst       As Long         'Its first data row
    Dim TrFirst       As Long         'Its first data row
    Dim Value         As Variant      'A cell read back

'------------------------------------------------------------------------------
' VALIDATE AND COMMIT
'------------------------------------------------------------------------------
        On Error GoTo Unexpected
        BeginCase "commit.both"
        WriteCsv CSV_NS_HEADER & vbLf & NS_ROW_1 & vbLf & NS_ROW_2 & vbLf
        M_Main.ImportCsv_Silent "NettingSets", mTestFile, True
        WriteCsv CSV_TR_HEADER & vbLf & TradeRow("00123", "1000", "007") & vbLf & TradeRow("T2", "2500.5", "") & vbLf & _
                 TradeRow("T3", "1000", "") & vbLf
        M_Main.ImportCsv_Silent "Trades", mTestFile, True
        ResultLine = M_Main.ValidateStaging_Silent()
        CheckText "validation", ResultLine, "RESULT=VALID; errors=0; nettingsets=2; trades=3"
        Before = LiveInputsText()
        M_Main.CommitImport
        Check InStr(1, M_Main.gLastMessage, "A commit is confirmed in a dialog, which silent mode does not show", _
                    vbBinaryCompare) > 0, "the button in silent mode: " & M_Main.gLastMessage
        Check LiveInputsText() = Before And M_Main.StagingValidationCurrent(), "the button in silent mode changed something"
        ResultLine = M_Main.CommitStaging_Silent()
        CheckText "commit", ResultLine, "RESULT=COMMITTED; nettingsets=2; trades=3"

'------------------------------------------------------------------------------
' THE INPUT SHEETS
'------------------------------------------------------------------------------
        Set NsSheet = GetSheet(SH_NS)
        Set TrSheet = GetSheet(SH_TRADES)
        NsFirst = TableFirstRow(SH_NS)
        TrFirst = TableFirstRow(SH_TRADES)
        Check LastLiveRow(NsSheet, NsFirst, NS_NCOLS) = NsFirst + 1, "NettingSets: not exactly two rows"
        CheckText "first netting set", CStr(TableCell(NsSheet, NsFirst, NS_ID).Value2), "NS-T1"
        Value = TableCell(TrSheet, TrFirst, TR_ID).Value2
        Check SameText(Value, "00123"), "the ID with leading zeros: " & CStr(Value)
        Value = TableCell(TrSheet, TrFirst + 1, TR_NOTIONAL).Value2
        Check SameNumber(Value, 2500.5), "the notional: " & CStr(Value)
        Value = TableCell(TrSheet, TrFirst, TR_MAT).Value2
        Check SameNumber(Value, CDbl(DateSerial(2031, 1, 2))), "the maturity: " & CStr(Value)
        Value = TableCell(TrSheet, TrFirst, TR_COMMENT).Value2
        Check SameText(Value, "007"), "the comment: " & CStr(Value)
        Check LastLiveRow(TrSheet, TrFirst, TR_NCOLS) = TrFirst + 2, "Trades: not exactly three rows"
        Check NsSheet.ProtectContents And TrSheet.ProtectContents And Not TableCell(NsSheet, NsFirst, NS_ID).Locked And _
              Not TableCell(TrSheet, TrFirst + 2, TR_COMMENT).Locked And TableCell(TrSheet, TrFirst - 1, TR_ID).Locked, _
              "the committed input sheets are not protected as the editability matrix says"

'------------------------------------------------------------------------------
' AFTER THE COMMIT
'------------------------------------------------------------------------------
        CheckText "statuses", Meta(KIND_NS, 4) & "|" & Meta(KIND_TR, 4), STAGING_COMMITTED & "|" & STAGING_COMMITTED
        Check Not M_Main.StagingValidationCurrent(), "the validation still reads as current"
        ResultLine = M_Main.CommitStaging_Silent()
        CheckText "second commit", ResultLine, "RESULT=REFUSED; reason=no current VALID staging validation"
        PutBackLive
        Check LiveInputsText() = Before, "the suite could not put the input sheets back"
        Exit Sub

Unexpected:
        Fail mCurrentCase, "unexpected error " & Err.Number & ": " & Err.Description
        PutBackLive

End Sub


Private Sub CaseCommitTradesOnly()
'
'==============================================================================
' PURPOSE
'   Only a leg staged, on the first netting set of NettingSets, and
'   Import_NettingSets empty (#211): the commit writes Trades only and
'   leaves NettingSets exactly as it was, never erasing it for lack of
'   staged rows; Import_Trades reads COMMITTED. The suite then puts
'   Trades back.
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ResultLine      As String    'What a call returned
    Dim Before          As String    'Input sheets and run names before
    Dim NsBefore        As String    'NettingSets before, as text
    Dim LiveSet         As String    'The first netting set of NettingSets

'------------------------------------------------------------------------------
' STAGE, VALIDATE AND COMMIT
'------------------------------------------------------------------------------
        On Error GoTo Unexpected
        BeginCase "commit.trades-only"
        LiveSet = CStr(TableCell(GetSheet(SH_NS), TableFirstRow(SH_NS), NS_ID).Value2)
        WriteCsv CSV_NS_HEADER & vbLf
        M_Main.ImportCsv_Silent "NettingSets", mTestFile, True
        WriteCsv CSV_TR_HEADER & vbLf & Replace(TradeRow("T1", "1000", ""), ";NS-T1;", ";" & LiveSet & ";") & vbLf
        M_Main.ImportCsv_Silent "Trades", mTestFile, True
        ResultLine = M_Main.ValidateStaging_Silent()
        CheckText "validation", ResultLine, "RESULT=VALID; errors=0; nettingsets=0; trades=1"
        Before = LiveInputsText()
        NsBefore = CORE_Engine.BlockText(GetSheet(SH_NS), TableFirstRow(SH_NS), NS_NCOLS)
        ResultLine = M_Main.CommitStaging_Silent()
        CheckText "commit", ResultLine, "RESULT=COMMITTED; nettingsets=0; trades=1"
        Check CORE_Engine.BlockText(GetSheet(SH_NS), TableFirstRow(SH_NS), NS_NCOLS) = NsBefore, "NettingSets changed"
        CheckText "Import_Trades status", Meta(KIND_TR, 4), STAGING_COMMITTED
        PutBackLive
        Check LiveInputsText() = Before, "the suite could not put the input sheets back"
        Exit Sub

Unexpected:
        Fail mCurrentCase, "unexpected error " & Err.Number & ": " & Err.Description
        PutBackLive

End Sub


Private Sub CaseCommitTrailingFormula()
'
'==============================================================================
' PURPOSE
'   A formula in an input cell of Trades below the data, which returns
'   empty text until the first leg's ID becomes T1, is inside the
'   commit's extent (#512): the commit clears it with the old rows and
'   reads it back, so no row outside the committed data can appear once
'   Excel recalculates. The suite then puts Trades back.
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ResultLine    As String       'What a call returned
    Dim Before        As String       'Input sheets and run names before
    Dim TrSheet       As Worksheet    'Trades
    Dim TrFirst       As Long         'Its first data row
    Dim FormulaCell   As Range        'The formula below the data

'------------------------------------------------------------------------------
' STAGE, VALIDATE AND COMMIT
'------------------------------------------------------------------------------
        On Error GoTo Unexpected
        BeginCase "commit.trailing-formula"
        Set TrSheet = GetSheet(SH_TRADES)
        TrFirst = TableFirstRow(SH_TRADES)
        Before = LiveInputsText()
        Set FormulaCell = TableCell(TrSheet, LastLiveRow(TrSheet, TrFirst, TR_NCOLS) + 3, TR_COMMENT)
        FormulaCell.Formula = "=IF(" & TableCell(TrSheet, TrFirst, TR_ID).Address(False, False) & _
                              "=""T1"",""ID-9"","""")"
        WriteCsv CSV_NS_HEADER & vbLf & NS_ROW_1 & vbLf & NS_ROW_2 & vbLf
        M_Main.ImportCsv_Silent "NettingSets", mTestFile, True
        WriteCsv CSV_TR_HEADER & vbLf & TradeRow("T1", "1000", "") & vbLf & TradeRow("T2", "2000", "") & vbLf
        M_Main.ImportCsv_Silent "Trades", mTestFile, True
        ResultLine = M_Main.ValidateStaging_Silent()
        ResultLine = M_Main.CommitStaging_Silent()
        CheckText "commit", ResultLine, "RESULT=COMMITTED; nettingsets=2; trades=2"
        Check Not FormulaCell.HasFormula And IsEmpty(FormulaCell.Value2), _
              "the formula below the data was left in place: " & FormulaCell.Formula
        Check LastLiveRow(TrSheet, TrFirst, TR_NCOLS) = TrFirst + 1, "Trades: not exactly two rows"
        PutBackLive
        Check LiveInputsText() = Before, "the suite could not put the input sheets back"
        Exit Sub

Unexpected:
        Fail mCurrentCase, "unexpected error " & Err.Number & ": " & Err.Description
        On Error Resume Next
        If Not FormulaCell Is Nothing Then FormulaCell.ClearContents
        PutBackLive

End Sub


Private Sub CaseCommitEquivalentRun()
'
'==============================================================================
' PURPOSE
'   Committed rows run as the same rows typed by hand (#211): two netting
'   sets and two legs, validated VALID and committed, then run; the run
'   finds no error, as the staging validation found none. The same rows
'   typed into the emptied input sheets, numbers as numbers, dates as
'   dates, text as text, give the same run line: the same input
'   fingerprint, counts, findings and total EADs. The input sheets are
'   then put back and run again, so that the workbook ends with a run of
'   its own inputs, CURRENT, or with no results when it had none. It runs
'   after import.live-sheets, which checks that no other case touches the
'   run.
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ResultLine    As String       'What a call returned
    Dim Committed     As String       'Run line after the commit
    Dim Typed         As String       'Run line after typing the same rows
    Dim SavedInputs   As String       'Input fingerprint before
    Dim HadRun        As Boolean      'Results existed before the case

'------------------------------------------------------------------------------
' COMMIT AND RUN, THEN TYPE AND RUN
'------------------------------------------------------------------------------
        On Error GoTo Unexpected
        BeginCase "commit.equivalent-run"
        SavedInputs = CORE_Engine.InputFingerprint()
        HadRun = Len(NameText(RUN_INPUTS_NAME)) > 0
        WriteCsv CSV_NS_HEADER & vbLf & NS_ROW_1 & vbLf & NS_ROW_2 & vbLf
        M_Main.ImportCsv_Silent "NettingSets", mTestFile, True
        WriteCsv CSV_TR_HEADER & vbLf & TradeRow("T1", "1000", "") & vbLf & TradeRow("T2", "2500.5", "") & vbLf
        M_Main.ImportCsv_Silent "Trades", mTestFile, True
        ResultLine = M_Main.ValidateStaging_Silent()
        CheckText "validation", ResultLine, "RESULT=VALID; errors=0; nettingsets=2; trades=2"
        ResultLine = M_Main.CommitStaging_Silent()
        CheckText "commit", ResultLine, "RESULT=COMMITTED; nettingsets=2; trades=2"
        Committed = M_Main.RunSACCR_Silent()
        Check Left$(Committed, 9) = "RESULT=OK" And InStr(1, Committed, "; errors=0;", vbBinaryCompare) > 0, _
              "the run on the committed rows: " & Committed
        TypeRows SH_NS, NS_NCOLS, Array(NS_ROW_1, NS_ROW_2)
        TypeRows SH_TRADES, TR_NCOLS, Array(TradeRow("T1", "1000", ""), TradeRow("T2", "2500.5", ""))
        Typed = M_Main.RunSACCR_Silent()
        Check Typed = Committed, "typed by hand: " & Typed & " | committed: " & Committed
        PutBackLive
        PutRunBack HadRun
        Check CORE_Engine.InputFingerprint() = SavedInputs And _
              M_Main.ResultsStatus() = IIf(HadRun, "CURRENT", "NONE"), _
              "inputs and results after the case: " & M_Main.ResultsStatus()
        Exit Sub

Unexpected:
        Fail mCurrentCase, "unexpected error " & Err.Number & ": " & Err.Description
        PutBackLive
        PutRunBack HadRun

End Sub


Private Sub TypeRows( _
    ByVal SheetName As String, _
    ByVal Width As Long, _
    ByVal Rows As Variant)
'
'==============================================================================
' PURPOSE
'   Empty the input rows of an input sheet and type CSV records into it as
'   a user would, field by field into the columns of the contract order:
'   a date as yyyy-mm-dd becomes a date, a number a number, anything else
'   text, and a blank field stays blank (#211).
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Sheet       As Worksheet    'The input sheet
    Dim FirstRow    As Long         'Its first data row
    Dim LastRow     As Long         'Its last row with a value or formula
    Dim Fields      As Variant      'One record's fields
    Dim Cell        As Range        'The cell of a field
    Dim r           As Long         'Record
    Dim c           As Long         'Field

'------------------------------------------------------------------------------
' TYPE
'------------------------------------------------------------------------------
        Set Sheet = GetSheet(SheetName)
        FirstRow = TableFirstRow(SheetName)
        LastRow = LastLiveRow(Sheet, FirstRow, Width)
        If LastRow >= FirstRow Then
            TableRange(Sheet, FirstRow, 1, LastRow, Width).ClearContents
        End If
        For r = LBound(Rows) To UBound(Rows)
            Fields = Split(CStr(Rows(r)), ";")
            For c = 0 To UBound(Fields)
                Set Cell = TableCell(Sheet, FirstRow + r - LBound(Rows), c + 1)
                If Len(Fields(c)) = 0 Then
                    'Left blank.
                ElseIf Fields(c) Like "####-##-##" Then
                    Cell.Value = DateSerial(CLng(Left$(Fields(c), 4)), CLng(Mid$(Fields(c), 6, 2)), CLng(Right$(Fields(c), 2)))
                ElseIf Fields(c) Like "*[!0-9.-]*" Then
                    Cell.Value2 = Fields(c)
                Else
                    Cell.Value2 = Val(Fields(c))
                End If
            Next c
        Next r

End Sub


Private Sub PutRunBack( _
    ByVal HadRun As Boolean)
'
'==============================================================================
' PURPOSE
'   Once CaseCommitEquivalentRun has put the input sheets back, run SA-CCR
'   on them when the workbook had results before the case, or clear the
'   outputs when it had none, so that the outputs, Checks and run names
'   agree again: CURRENT, or NONE.
'
' ERROR POLICY
'   Contained: a failure is reported as a test failure of the case.
'==============================================================================
'
        On Error GoTo Failed
        If HadRun Then
            M_Main.RunSACCR_Silent
        Else
            M_Main.ClearOutputs
        End If
        Exit Sub

Failed:
        Fail mCurrentCase, "could not put the run back: " & Err.Description

End Sub


Private Sub CaseValidateRecalc()
'
'==============================================================================
' PURPOSE
'   In manual calculation, a Params formula whose precedent changed is
'   judged by its new result (#519). With AsOfDate a formula on a free
'   Params cell, moved after the maturity of the staged leg and not
'   recalculated, the validation finds the leg matured; once the date is
'   back and the validation VALID, the same change, still not
'   recalculated, makes the commit refuse. AsOfDate, the free cell and
'   the calculation mode are then put back; the input sheets are not
'   written.
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ResultLine      As String           'What a call returned
    Dim DateCell        As Range            'AsOfDate
    Dim Precedent       As Range            'A free Params cell its formula reads
    Dim SavedValue      As Variant          'AsOfDate before the case
    Dim SavedFormula    As String           'Its formula, when it has one
    Dim HadFormula      As Boolean          'It had one
    Dim SavedCalc       As XlCalculation    'Calculation mode before the case

'------------------------------------------------------------------------------
' STAGE, CHANGE THE DATE, VALIDATE AND COMMIT
'------------------------------------------------------------------------------
        On Error GoTo Unexpected
        BeginCase "validate.recalc"
        SavedCalc = Application.Calculation
        Set DateCell = NamedCell("AsOfDate")
        Set Precedent = DateCell.Offset(0, 8)
        SavedValue = DateCell.Value2
        HadFormula = DateCell.HasFormula
        SavedFormula = DateCell.Formula
        WriteCsv CSV_NS_HEADER & vbLf & NS_ROW_1 & vbLf & NS_ROW_2 & vbLf
        M_Main.ImportCsv_Silent "NettingSets", mTestFile, True
        WriteCsv CSV_TR_HEADER & vbLf & TradeRow("T1", "1000", "") & vbLf
        M_Main.ImportCsv_Silent "Trades", mTestFile, True
        Application.Calculation = xlCalculationManual
        Precedent.Value2 = CDbl(DateSerial(2026, 9, 30))
        DateCell.Formula = "=" & Precedent.Address(False, False)
        Application.Calculate
        Precedent.Value2 = CDbl(DateSerial(2031, 6, 30))
        ResultLine = M_Main.ValidateStaging_Silent()
        Check Left$(ResultLine, 14) = "RESULT=INVALID" And InStr(1, ResultLine, "CSV-CONTEXT-MATURED", _
              vbBinaryCompare) > 0, "the validation judged the reporting date Excel had kept: " & ResultLine
        Precedent.Value2 = CDbl(DateSerial(2026, 9, 30))
        Application.Calculate
        ResultLine = M_Main.ValidateStaging_Silent()
        CheckText "validation", ResultLine, "RESULT=VALID; errors=0; nettingsets=2; trades=1"
        Precedent.Value2 = CDbl(DateSerial(2031, 6, 30))
        ResultLine = M_Main.CommitStaging_Silent()
        CheckText "commit", ResultLine, "RESULT=REFUSED; reason=no current VALID staging validation"
        PutDateBack DateCell, Precedent, SavedValue, SavedFormula, HadFormula, SavedCalc
        Exit Sub

Unexpected:
        Fail mCurrentCase, "unexpected error " & Err.Number & ": " & Err.Description
        PutDateBack DateCell, Precedent, SavedValue, SavedFormula, HadFormula, SavedCalc

End Sub


Private Sub PutDateBack( _
    ByVal DateCell As Range, _
    ByVal Precedent As Range, _
    ByVal SavedValue As Variant, _
    ByVal SavedFormula As String, _
    ByVal HadFormula As Boolean, _
    ByVal SavedCalc As XlCalculation)
'
'==============================================================================
' PURPOSE
'   Give AsOfDate back its value or formula, empty the free cell of
'   CaseValidateRecalc and restore the calculation mode.
'
' ERROR POLICY
'   Contained: a failure is reported as a test failure of the case.
'==============================================================================
'
        On Error GoTo Failed
        If Not DateCell Is Nothing Then
            If HadFormula Then
                DateCell.Formula = SavedFormula
            Else
                DateCell.Value2 = SavedValue
            End If
        End If
        If Not Precedent Is Nothing Then
            Precedent.ClearContents
        End If
        Application.Calculation = SavedCalc
        Exit Sub

Failed:
        Fail mCurrentCase, "could not put AsOfDate back: " & Err.Description
        Application.Calculation = SavedCalc

End Sub


Private Sub CaseLiveSheets( _
    ByVal LiveBefore As String)
'
'==============================================================================
' PURPOSE
'   After every load and validation, the live inputs and Params and the
'   stored run and reconciliation fingerprints are as they were, and so
'   is Checks once put back.
'==============================================================================
'
        On Error GoTo Unexpected
        BeginCase "import.live-sheets"
        CheckText "live inputs and run fingerprints", LiveInputsText(), LiveBefore
        CheckText "Checks put back", ChecksText(), mChecksBefore
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


Private Function StagedFingerprint( _
    ByVal KindIndex As Long) _
    As String
'
'==============================================================================
' PURPOSE
'   The fingerprint M_Main stores for a staging block, computed here on
'   its own: the rows up to the last with a value, each cell as a typed
'   token, a blank cell and an empty text alike, joined by Chr$(31); ""
'   when the block is empty.
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Sheet       As Worksheet    'The staging sheet
    Dim FirstRow    As Long         'First data row
    Dim LastRow     As Long         'Last row with a value
    Dim Width       As Long         'Columns, the CSV line column included
    Dim Grid        As Variant      'The block's values
    Dim Pieces()    As String       'One token per cell
    Dim r           As Long         'Row of Grid
    Dim c           As Long         'Column of Grid
    Dim k           As Long         'Position in Pieces
    Dim RowFilled   As Boolean      'The row holds a value
    Dim Cell        As Variant      'A cell's value

'------------------------------------------------------------------------------
' COMPUTE
'------------------------------------------------------------------------------
        Set Sheet = StagingSheet(KindIndex)
        FirstRow = FirstDataRow(KindIndex)
        Width = BlockWidth(KindIndex)
        For LastRow = UsedLastRow(Sheet) To FirstRow Step -1
            RowFilled = False
            For c = 1 To Width
                Cell = TableCell(Sheet, LastRow, c).Value2
                If IsError(Cell) Then
                    RowFilled = True
                ElseIf Len(CStr(Cell)) > 0 Then
                    RowFilled = True
                End If
            Next c
            If RowFilled Then
                Exit For
            End If
        Next LastRow
        If LastRow < FirstRow Then
            Exit Function
        End If
        Grid = TableRange(Sheet, FirstRow, 1, LastRow, Width).Value2
        ReDim Pieces(1 To (LastRow - FirstRow + 1) * Width)
        For r = 1 To LastRow - FirstRow + 1
            For c = 1 To Width
                k = k + 1
                Select Case VarType(Grid(r, c))
                    Case vbEmpty
                    Case vbString
                        If Len(Grid(r, c)) > 0 Then
                            Pieces(k) = "T" & Grid(r, c)
                        End If
                    Case vbBoolean
                        Pieces(k) = IIf(Grid(r, c), "B1", "B0")
                    Case vbError
                        Pieces(k) = "E" & CStr(Grid(r, c))
                    Case Else
                        Pieces(k) = "N" & CStr(Grid(r, c))
                End Select
            Next c
        Next r
        StagedFingerprint = TextHash(Join(Pieces, Chr$(31)))

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


Private Sub PutValidatedBack()
'
'==============================================================================
' PURPOSE
'   Put back the fingerprints of the last VALID staging validation and of
'   the last staging validation as the suite found them, or remove a name
'   when there was none.
'
' ERROR POLICY
'   A failure is recorded as a suite failure.
'==============================================================================
'
    Dim i   As Long    'Name

        On Error GoTo Failed
        For i = 1 To 2
            If mHadValidated(i) Then
                ThisWorkbook.Names.Add Name:=ValidationName(i), RefersTo:=mSavedValidated(i), Visible:=False
            ElseIf Len(NameText(ValidationName(i))) > 0 Then
                ThisWorkbook.Names(ValidationName(i)).Delete
            End If
        Next i
        Exit Sub

Failed:
        Fail "suite", "the staging validation fingerprint not put back: " & Err.Number & ": " & Err.Description

End Sub


Private Function LegsOutside( _
    ByVal NettingSets As String, _
    ByRef FirstCell As String) _
    As Long
'
'==============================================================================
' PURPOSE
'   The legs on Trades with a Trade ID and a netting set that is not one
'   of NettingSets, separated by "|", compared trimmed and in upper case
'   as the engine compares them; FirstCell is the Netting set cell of the
'   first of them.
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Sheet       As Worksheet    'Trades
    Dim TopRow      As Long         'Its first data row
    Dim LastRow     As Long         'Its last row with a value
    Dim Known       As String       'The netting sets, between bars
    Dim NettingSet  As String       'A leg's netting set
    Dim r           As Long         'Row

'------------------------------------------------------------------------------
' COUNT
'------------------------------------------------------------------------------
        Set Sheet = GetSheet(SH_TRADES)
        TopRow = TableFirstRow(SH_TRADES)
        LastRow = LastDataRowAny(Sheet, TopRow, TR_NCOLS, TR_COMMENT)
        Known = "|" & UCase$(NettingSets) & "|"
        For r = TopRow To LastRow
            NettingSet = UTxt(TableCell(Sheet, r, TR_NS).Value2)
            If Len(SafeStr(TableCell(Sheet, r, TR_ID).Value2)) > 0 And Len(NettingSet) > 0 Then
                If InStr(1, Known, "|" & NettingSet & "|", vbBinaryCompare) = 0 Then
                    LegsOutside = LegsOutside + 1
                    If LegsOutside = 1 Then
                        FirstCell = TableCell(Sheet, r, TR_NS).Address(False, False)
                    End If
                End If
            End If
        Next r

End Function


Private Function LegsOfClass( _
    ByVal NettingSet As String, _
    ByVal AssetClass As String, _
    ByRef FirstCell As String) _
    As Long
'
'==============================================================================
' PURPOSE
'   The legs on Trades of one netting set and asset class, and the Asset
'   class cell of the first of them.
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Sheet       As Worksheet    'Trades
    Dim TopRow      As Long         'Its first data row
    Dim LastRow     As Long         'Its last row with a value
    Dim r           As Long         'Row

'------------------------------------------------------------------------------
' COUNT
'------------------------------------------------------------------------------
        Set Sheet = GetSheet(SH_TRADES)
        TopRow = TableFirstRow(SH_TRADES)
        LastRow = LastDataRowAny(Sheet, TopRow, TR_NCOLS, TR_COMMENT)
        For r = TopRow To LastRow
            If Len(SafeStr(TableCell(Sheet, r, TR_ID).Value2)) > 0 And _
               UTxt(TableCell(Sheet, r, TR_NS).Value2) = UCase$(NettingSet) And _
               UTxt(TableCell(Sheet, r, TR_AC).Value2) = UCase$(AssetClass) Then
                LegsOfClass = LegsOfClass + 1
                If LegsOfClass = 1 Then
                    FirstCell = TableCell(Sheet, r, TR_AC).Address(False, False)
                End If
            End If
        Next r

End Function


Private Function FindingCode( _
    ByVal LineNumber As Long) _
    As String
'
'==============================================================================
' PURPOSE
'   The code a Checks line's message starts with, before its first ":".
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Text    As String    'The message

'------------------------------------------------------------------------------
' READ
'------------------------------------------------------------------------------
        Text = ChecksValue(LineNumber, 5)
        If InStr(Text, ":") > 0 Then
            FindingCode = Left$(Text, InStr(Text, ":") - 1)
        End If

End Function


Private Function RunCounters() As String
'
'==============================================================================
' PURPOSE
'   The counters the last run leaves readable: its ERROR and WARNING
'   messages and the trades it read.
'==============================================================================
'
        RunCounters = CORE_Engine.ErrorCount & "|" & CORE_Engine.WarningCount & "|" & CORE_Engine.TradesRead

End Function


Private Function ChecksCell( _
    ByVal LineNumber As Long, _
    ByVal Column As Long) _
    As Range
'
'==============================================================================
' PURPOSE
'   A cell of the Checks table: line 1 is its first data row.
'==============================================================================
'
        Set ChecksCell = TableCell(GetSheet(SH_CHECKS), OutputFirstRow(SH_CHECKS) + LineNumber - 1, Column)

End Function


Private Function ChecksValue( _
    ByVal LineNumber As Long, _
    ByVal Column As Long) _
    As String
'
'==============================================================================
' PURPOSE
'   The value of a cell of the Checks table, as text.
'==============================================================================
'
        ChecksValue = SafeStr(ChecksCell(LineNumber, Column).Value2)

End Function


Private Function ChecksText() As String
'
'==============================================================================
' PURPOSE
'   The Checks table as text, as M_Main fingerprints an output block.
'==============================================================================
'
        ChecksText = CORE_Engine.BlockText(GetSheet(SH_CHECKS), OutputFirstRow(SH_CHECKS), CK_NCOLS)

End Function


Private Sub SaveChecks()
'
'==============================================================================
' PURPOSE
'   Keep the lines of Checks, to put them back at the end.
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Sheet      As Worksheet    'Checks
    Dim FirstRow   As Long         'Its first data row
    Dim LastRow    As Long         'Last row of its used range

'------------------------------------------------------------------------------
' KEEP
'------------------------------------------------------------------------------
        Set Sheet = GetSheet(SH_CHECKS)
        FirstRow = OutputFirstRow(SH_CHECKS)
        LastRow = UsedLastRow(Sheet)
        mSavedChecksRows = 0
        mSavedChecks = Empty
        If LastRow >= FirstRow Then
            mSavedChecksRows = LastRow - FirstRow + 1
            mSavedChecks = TableRange(Sheet, FirstRow, 1, LastRow + 1, CK_NCOLS).Value2
        End If
        mChecksBefore = ChecksText()

End Sub


Private Sub PutChecksBack()
'
'==============================================================================
' PURPOSE
'   Put back the lines SaveChecks kept, as values, the severity of an
'   ERROR line in bold as the engine writes it.
'
' ERROR POLICY
'   A failure is recorded as a suite failure.
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Sheet      As Worksheet    'Checks
    Dim FirstRow   As Long         'Its first data row
    Dim r          As Long         'Line

'------------------------------------------------------------------------------
' PUT BACK
'------------------------------------------------------------------------------
        On Error GoTo Failed
        Set Sheet = GetSheet(SH_CHECKS)
        FirstRow = OutputFirstRow(SH_CHECKS)
        ClearChecksBlock Sheet, FirstRow
        If mSavedChecksRows > 0 Then
            TableRange(Sheet, FirstRow, 1, FirstRow + mSavedChecksRows, CK_NCOLS).Value2 = mSavedChecks
            For r = 1 To mSavedChecksRows
                If SafeStr(mSavedChecks(r, 1)) = SEV_ERROR Then
                    TableCell(Sheet, FirstRow + r - 1, 1).Font.Bold = True
                End If
            Next r
        End If
        Exit Sub

Failed:
        Fail "suite", "Checks not put back: " & Err.Number & ": " & Err.Description

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


Private Function ValidationName( _
    ByVal Index As Long) _
    As String
'
'==============================================================================
' PURPOSE
'   The hidden names of the staging validation: 1 STAGING_VALIDATED_NAME,
'   2 STAGING_CHECKED_NAME.
'==============================================================================
'
        ValidationName = IIf(Index = 1, STAGING_VALIDATED_NAME, STAGING_CHECKED_NAME)

End Function


Private Sub SaveLive()
'
'==============================================================================
' PURPOSE
'   Keep the input cells of NettingSets and Trades, for PutBackLive: for
'   each formula cell its formula, for any other its text with a leading
'   apostrophe so that it is written back as text, even one that starts
'   with "=", or its value.
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Sheet       As Worksheet    'An input sheet
    Dim FirstRow    As Long         'Its first data row
    Dim Width       As Long         'Its input columns
    Dim Values      As Variant      'Value2 of its cells
    Dim Formulas    As Variant      'Formula of its cells
    Dim Kept()      As Variant      'What is kept
    Dim k           As Long         'Sheet
    Dim r           As Long         'Row
    Dim c           As Long         'Column

'------------------------------------------------------------------------------
' KEEP
'------------------------------------------------------------------------------
        For k = 1 To 2
            Set Sheet = GetSheet(IIf(k = 1, SH_NS, SH_TRADES))
            FirstRow = TableFirstRow(Sheet.Name)
            Width = IIf(k = 1, NS_NCOLS, TR_NCOLS)
            mSavedLiveLast(k) = LastLiveRow(Sheet, FirstRow, Width)
            mSavedLive(k) = Empty
            If mSavedLiveLast(k) >= FirstRow Then
                With TableRange(Sheet, FirstRow, 1, mSavedLiveLast(k), Width)
                    Values = .Value2
                    Formulas = FormulaCells(.Cells)
                End With
                ReDim Kept(1 To mSavedLiveLast(k) - FirstRow + 1, 1 To Width)
                For r = 1 To UBound(Kept, 1)
                    For c = 1 To Width
                        If Len(CStr(Formulas(r, c))) > 0 Then
                            Kept(r, c) = Formulas(r, c)
                        ElseIf VarType(Values(r, c)) = vbString Then
                            Kept(r, c) = "'" & Values(r, c)
                        Else
                            Kept(r, c) = Values(r, c)
                        End If
                    Next c
                Next r
                mSavedLive(k) = Kept
            End If
        Next k

End Sub


Private Sub PutBackLive()
'
'==============================================================================
' PURPOSE
'   Put the input cells of NettingSets and Trades back as SaveLive kept
'   them: clear the input cells down to the last row with a value now or
'   then, and write the kept cells.
'
' ERROR POLICY
'   A failure is recorded as a suite failure.
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Sheet       As Worksheet    'An input sheet
    Dim FirstRow    As Long         'Its first data row
    Dim Width       As Long         'Its input columns
    Dim LastRow     As Long         'Last row to clear
    Dim Events      As Boolean      'EnableEvents before
    Dim k           As Long         'Sheet

'------------------------------------------------------------------------------
' PUT BACK
'------------------------------------------------------------------------------
        On Error GoTo Failed
        Events = Application.EnableEvents
        Application.EnableEvents = False
        For k = 1 To 2
            Set Sheet = GetSheet(IIf(k = 1, SH_NS, SH_TRADES))
            FirstRow = TableFirstRow(Sheet.Name)
            Width = IIf(k = 1, NS_NCOLS, TR_NCOLS)
            LastRow = LastLiveRow(Sheet, FirstRow, Width)
            If mSavedLiveLast(k) > LastRow Then
                LastRow = mSavedLiveLast(k)
            End If
            If LastRow >= FirstRow Then
                TableRange(Sheet, FirstRow, 1, LastRow, Width).ClearContents
            End If
            If mSavedLiveLast(k) >= FirstRow Then
                TableRange(Sheet, FirstRow, 1, mSavedLiveLast(k), Width).Value = mSavedLive(k)
            End If
        Next k
        Application.EnableEvents = Events
        Exit Sub

Failed:
        Application.EnableEvents = Events
        Fail "suite", "the input sheets not put back: " & Err.Number & ": " & Err.Description

End Sub


Private Function LastLiveRow( _
    ByVal Sheet As Worksheet, _
    ByVal FirstRow As Long, _
    ByVal Width As Long) _
    As Long
'
'==============================================================================
' PURPOSE
'   The last row of an input sheet with a value or a formula in one of
'   its input columns, whatever the formula returns, or FirstRow - 1 when
'   none has either: the extent the commit replaces (#512), so that the
'   suite saves and puts back a formula below the data too.
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim LastRow   As Long       'Last row of the used range
    Dim Grid      As Variant    'The input cells
    Dim Formulas  As Variant    'Their formulas; "" for a constant
    Dim r         As Long       'Row of Grid
    Dim c         As Long       'Column of Grid

'------------------------------------------------------------------------------
' SCAN FROM THE BOTTOM
'------------------------------------------------------------------------------
        LastLiveRow = FirstRow - 1
        LastRow = UsedLastRow(Sheet)
        If LastRow < FirstRow Then
            Exit Function
        End If
        With TableRange(Sheet, FirstRow, 1, LastRow, Width)
            Grid = .Value2
            Formulas = FormulaCells(.Cells)
        End With
        For r = UBound(Grid, 1) To 1 Step -1
            For c = 1 To Width
                If Len(CStr(Formulas(r, c))) > 0 Then
                    LastLiveRow = FirstRow + r - 1
                    Exit Function
                ElseIf Not IsEmpty(Grid(r, c)) Then
                    If IsError(Grid(r, c)) Then
                        LastLiveRow = FirstRow + r - 1
                        Exit Function
                    ElseIf Len(CStr(Grid(r, c))) > 0 Then
                        LastLiveRow = FirstRow + r - 1
                        Exit Function
                    End If
                End If
            Next c
        Next r

End Function


Private Function SameNumber( _
    ByVal Value As Variant, _
    ByVal Expected As Double) _
    As Boolean
'
'==============================================================================
' PURPOSE
'   Whether a cell's Value2 is the number Expected, held as a number.
'==============================================================================
'
        If VarType(Value) = vbDouble Then
            SameNumber = (Value = Expected)
        End If

End Function


Private Function SameText( _
    ByVal Value As Variant, _
    ByVal Expected As String) _
    As Boolean
'
'==============================================================================
' PURPOSE
'   Whether a cell's Value2 is the text Expected, held as text.
'==============================================================================
'
        If VarType(Value) = vbString Then
            SameText = (Value = Expected)
        End If

End Function


Private Function FormulaCells( _
    ByVal Target As Range) _
    As Variant
'
'==============================================================================
' PURPOSE
'   The formula of each formula cell of Target, by row and column within
'   it, and "" for any other cell: found with SpecialCells, because
'   Formula returns the value of a constant cell.
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Found   As Range      'The formula cells; Nothing for none
    Dim Area    As Range      'One area of them
    Dim Cell    As Range      'One of them
    Dim Out()   As Variant    'The result

'------------------------------------------------------------------------------
' FIND
'------------------------------------------------------------------------------
        ReDim Out(1 To Target.Rows.Count, 1 To Target.Columns.Count)
        If Not IsNull(Target.HasFormula) Then
            If Not Target.HasFormula Then
                FormulaCells = Out
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
        FormulaCells = Out

End Function
