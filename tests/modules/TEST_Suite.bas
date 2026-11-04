Attribute VB_Name = "TEST_Suite"
'==============================================================================
' MODULE: TEST_Suite
'------------------------------------------------------------------------------
' PURPOSE
'   Run every test suite in one go and print only a summary: one line per
'   suite with its verdict, counts and duration, the failure lines of any
'   suite that failed, and a final RESULT line. The suites still print
'   their full logs when they are run on their own. Certify runs the
'   harness and every suite once each, on its own, and then Test_RunAll,
'   writing each report to its own log file of an evidence bundle, first
'   line CANDIDATE= (#535).
'
' PUBLIC SURFACE
'   Test_RunAll runs the suites. Say is the print routine of every suite,
'   of TEST_CaseRunner and of TEST_Benchmark: it writes to the Immediate
'   window, or during Test_RunAll records the RESULT and FAILURE lines
'   instead; while Certify runs, every printed line is also written to
'   the open log file. Certify and CertificationReady are called by
'   tools\Certify-SACCR.ps1 through Application.Run. Option Private
'   Module keeps them out of the external workbook automation API.
'
' DEPENDENCIES
'   CORE_Export.Utf8Bytes, to write the log files as UTF-8; TEST_Harness, TEST_Formulas, TEST_MainState, TEST_Cases,
'   TEST_InputValidation, TEST_Aggregation, TEST_Invariants,
'   TEST_Reconcile, TEST_CsvReader and TEST_Import, with TEST_CaseRunner
'   for TEST_Cases, TEST_InputValidation, TEST_Aggregation and
'   TEST_Invariants;
'   M_Main.gQuietLog, set while a suite runs, so that the macros'
'   silent-mode messages are not printed either.
'
' WORKSHEET SAFETY
'   As the suites it runs: they rewrite inputs, outputs and Params and
'   restore them. Use a development workbook.
'
' ERROR POLICY
'   An error raised by a suite is recorded as that suite's failure and the
'   remaining suites still run. Certify writes it to the suite's log as an
'   ERROR= line, without a RESULT line, and goes on; a log line it cannot
'   write is counted and reported in its result.
'
' USAGE
'   Import with the test modules, compile, then type Test_RunAll in the
'   Immediate window. For an evidence bundle, run tools\Certify-SACCR.ps1,
'   which calls Certify (docs/EXCEL_EVIDENCE.md).
'
' UPDATED
'   2026-11-04
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
        Private Const SUITE_COUNT   As Long = 10    'Suites run by Test_RunAll

'------------------------------------------------------------------------------
' MODULE STATE
'------------------------------------------------------------------------------
        Private mQuiet      As Boolean       'True while Test_RunAll runs a suite
        Private mResult     As String        'Last RESULT line of the running suite
        Private mFailures   As Collection    'FAILURE lines of the running suite
        Private mLogPath    As String        'Log file being written by Certify; "" for none
        Private mLogErrors  As Long          'Lines Certify could not write


'
'------------------------------------------------------------------------------
'
'                                 ENTRY POINT
'
'------------------------------------------------------------------------------
'

Public Sub Test_RunAll()
'
'==============================================================================
'                                 Test_RunAll
'------------------------------------------------------------------------------
' PURPOSE
'   Run the nine suites in turn and print the summary.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Suites         As Variant    'Entry points, as Module.Procedure
    Dim i              As Long       'Index into Suites
    Dim FailedCount    As Long       'Suites that did not pass
    Dim StartTime      As Single     'Start of the whole run
    Dim TotalSeconds   As Single     'Seconds for the whole run

'------------------------------------------------------------------------------
' RUN
'------------------------------------------------------------------------------
        Suites = SuiteEntries()
        Emit "SACCR TEST SUMMARY " & Format$(Now, "yyyy-mm-dd hh:nn:ss")
        StartTime = Timer
        For i = LBound(Suites) To UBound(Suites)
            If Not RunOneSuite(CStr(Suites(i))) Then
                FailedCount = FailedCount + 1
            End If
        Next i

'------------------------------------------------------------------------------
' REPORT
'------------------------------------------------------------------------------
        TotalSeconds = Timer - StartTime
        If TotalSeconds < 0 Then
            TotalSeconds = TotalSeconds + 86400!
        End If
        Emit "RESULT=" & IIf(FailedCount = 0, "PASS", "FAIL") & "; suites=" & SUITE_COUNT & _
                    "; passed=" & (SUITE_COUNT - FailedCount) & "; failed=" & FailedCount & _
                    "; seconds=" & Format$(TotalSeconds, "0.0")

End Sub


Public Sub Say( _
    ByVal Text As String)
'
'==============================================================================
'                                     Say
'------------------------------------------------------------------------------
' PURPOSE
'   Print one line of a suite's log. During Test_RunAll the line is not
'   printed: the RESULT line is kept, and the FAILURE= and FAILURE_DETAILS=
'   lines are collected for the summary; a count such as FAILURES= is not.
'   A printed line is also written to the log that Certify has open.
'
' UPDATED
'   2026-11-04
'==============================================================================
'
        If Not mQuiet Then
            Emit Text
        ElseIf Left$(Text, 7) = "RESULT=" Then
            mResult = Text
        ElseIf Left$(Text, 8) = "FAILURE=" Or Left$(Text, 16) = "FAILURE_DETAILS=" Then
            mFailures.Add Text
        End If

End Sub


Public Function Certify( _
    ByVal Folder As String, _
    ByVal Candidate As String) _
    As String
'
'==============================================================================
'                                   Certify
'------------------------------------------------------------------------------
' PURPOSE
'   Write the test logs of an Excel evidence bundle (#535). Run the harness
'   and every suite once each, on its own and in the order Test_RunAll
'   runs them, each into its own log file in Folder: harness.log for
'   TEST_Harness, then TEST_Formulas.log and so on. Then run Test_RunAll
'   into summary.log. Each log starts with the line CANDIDATE=Candidate,
'   written here, not typed, followed by every line the run prints, as the
'   Immediate window shows it. The files are UTF-8, with CRLF line ends.
'
' INPUTS
'   Folder: the bundle folder; it must exist.
'   Candidate: the full 40-character lowercase SHA of the checked-out
'      commit, as tools\Certify-SACCR.ps1 reads it from Git.
'
' RETURNS
'   "RESULT=DONE; logs=11; write_errors=0" once every run has ended, the
'   count of lines that could not be written last. "RESULT=REFUSED;
'   reason=..." when the candidate is not a full lowercase SHA, the folder
'   does not exist, or one of the logs exists already: a log is never
'   replaced, so that a failing report cannot be overwritten by a later
'   one.
'
' ERROR POLICY
'   An error raised by a suite is printed and written to its log as an
'   ERROR= line, and the next suite runs. A line that cannot be written is
'   counted in write_errors; the validator then rejects the log.
'
' UPDATED
'   2026-11-04
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Suites      As Variant    'Entry points, as Module.Procedure
    Dim Base        As String     'Folder with one path separator at its end
    Dim i           As Long       'Index into Suites

'------------------------------------------------------------------------------
' CHECK
'------------------------------------------------------------------------------
        If Not IsCandidate(Candidate) Then
            Certify = "RESULT=REFUSED; reason=the candidate is not a full lowercase commit SHA"
            Exit Function
        End If
        Base = Folder
        If Right$(Base, 1) <> Application.PathSeparator Then
            Base = Base & Application.PathSeparator
        End If
        If Not FolderExists(Base) Then
            Certify = "RESULT=REFUSED; reason=the folder does not exist: " & Folder
            Exit Function
        End If
        Suites = SuiteEntries()
        For i = LBound(Suites) To UBound(Suites)
            If FileExists(Base & LogNameOf(CStr(Suites(i)))) Then
                Certify = "RESULT=REFUSED; reason=a log exists already: " & LogNameOf(CStr(Suites(i)))
                Exit Function
            End If
        Next i
        If FileExists(Base & "summary.log") Then
            Certify = "RESULT=REFUSED; reason=a log exists already: summary.log"
            Exit Function
        End If

'------------------------------------------------------------------------------
' RUN EACH SUITE ON ITS OWN, THEN THE SUMMARY
'------------------------------------------------------------------------------
        mLogErrors = 0
        For i = LBound(Suites) To UBound(Suites)
            RunLogged Base & LogNameOf(CStr(Suites(i))), Candidate, CStr(Suites(i))
        Next i
        mLogPath = Base & "summary.log"
        WriteLogLine "CANDIDATE=" & Candidate
        Test_RunAll
        mLogPath = ""
        Certify = "RESULT=DONE; logs=" & (UBound(Suites) - LBound(Suites) + 2) & "; write_errors=" & mLogErrors

End Function


Public Function CertificationReady() As String
'
'==============================================================================
'                              CertificationReady
'------------------------------------------------------------------------------
' PURPOSE
'   Return "READY". tools\Certify-SACCR.ps1 calls it through
'   Application.Run before Certify: the call succeeds only when the
'   workbook's macros may run, so the script can ask the operator to enable
'   content instead of changing any macro setting.
'
' UPDATED
'   2026-11-04
'==============================================================================
'
        CertificationReady = "READY"

End Function


'
'------------------------------------------------------------------------------
'
'                                   HELPERS
'
'------------------------------------------------------------------------------
'

Private Function RunOneSuite( _
    ByVal EntryPoint As String) _
    As Boolean
'
'==============================================================================
'                                 RunOneSuite
'------------------------------------------------------------------------------
' PURPOSE
'   Run one suite quietly and print its summary line, followed by its
'   failure lines, or the error it raised, when it did not pass.
'
' RETURNS
'   True when the suite printed RESULT=PASS and raised no error.
'
' ERROR POLICY
'   An error raised by the suite is contained and reported as a failure.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim StartTime      As Single     'Start of the suite
    Dim SuiteSeconds   As Single     'Duration of the suite
    Dim RaisedError    As String     'Error raised by the suite; "" when none
    Dim Passed         As Boolean    'RESULT=PASS and no error
    Dim Detail         As String     'The RESULT line after its verdict
    Dim FailureLine    As Variant    'One failure line

'------------------------------------------------------------------------------
' RUN QUIETLY
'------------------------------------------------------------------------------
    'M_Main.gQuietLog also drops the messages the macros print in silent
    'mode, one per engine run.
        mResult = ""
        Set mFailures = New Collection
        mQuiet = True
        M_Main.gQuietLog = True
        StartTime = Timer
        On Error GoTo Raised
        RunEntry EntryPoint

Report:
        On Error GoTo 0
        mQuiet = False
        M_Main.gQuietLog = False
        SuiteSeconds = Timer - StartTime
        If SuiteSeconds < 0 Then
            SuiteSeconds = SuiteSeconds + 86400!
        End If

'------------------------------------------------------------------------------
' PRINT THE SUITE LINE
'------------------------------------------------------------------------------
    'The RESULT line without its verdict, which the first column shows.
        Passed = (Left$(mResult, 11) = "RESULT=PASS") And Len(RaisedError) = 0
        Detail = mResult
        If InStr(Detail, "; ") > 0 Then
            Detail = Mid$(Detail, InStr(Detail, "; ") + 2)
        ElseIf Len(Detail) = 0 Then
            Detail = "no RESULT line"
        End If
        Emit "  " & IIf(Passed, "PASS", "FAIL") & "  " & Left$(Split(EntryPoint, ".")(0) & Space$(22), 22) & _
                    Detail & "  (" & Format$(SuiteSeconds, "0.0") & " s)"
        If Not Passed Then
            For Each FailureLine In mFailures
                Emit "        " & FailureLine
            Next FailureLine
            If Len(RaisedError) > 0 Then
                Emit "        ERROR=" & RaisedError
            End If
        End If
        RunOneSuite = Passed
        Exit Function

'------------------------------------------------------------------------------
' HANDLE A RAISED ERROR
'------------------------------------------------------------------------------
Raised:
        RaisedError = Err.Number & ": " & Err.Description
        Resume Report

End Function


Private Function SuiteEntries() As Variant
'
'==============================================================================
'                                 SuiteEntries
'------------------------------------------------------------------------------
' PURPOSE
'   The entry points of the harness and the nine suites, in the order of
'   .github/excel-evidence-policy.json, as Module.Procedure. The policy
'   gate reads this list.
'
' UPDATED
'   2026-11-04
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Suites      As Variant    'Entry points, as Module.Procedure

'------------------------------------------------------------------------------
' LIST
'------------------------------------------------------------------------------
        Suites = Array("TEST_Harness.RunTests", "TEST_Formulas.RunFormulaTests", _
                       "TEST_MainState.RunMainStateTests", "TEST_Cases.RunCaseTests", _
                       "TEST_InputValidation.RunInputValidationTests", _
                       "TEST_Aggregation.RunAggregationTests", "TEST_Invariants.RunInvariantTests", _
                       "TEST_Reconcile.RunReconcileTests", "TEST_CsvReader.RunCsvReaderTests", _
                       "TEST_Import.RunImportTests")
        SuiteEntries = Suites

End Function


Private Sub RunEntry( _
    ByVal EntryPoint As String)
'
'==============================================================================
'                                   RunEntry
'------------------------------------------------------------------------------
' PURPOSE
'   Run one entry point of SuiteEntries. An error it raises goes to the
'   caller.
'
' UPDATED
'   2026-11-04
'==============================================================================
'
        Select Case EntryPoint
            Case "TEST_Harness.RunTests"
                TEST_Harness.RunTests
            Case "TEST_Formulas.RunFormulaTests"
                TEST_Formulas.RunFormulaTests
            Case "TEST_MainState.RunMainStateTests"
                TEST_MainState.RunMainStateTests
            Case "TEST_Cases.RunCaseTests"
                TEST_Cases.RunCaseTests
            Case "TEST_InputValidation.RunInputValidationTests"
                TEST_InputValidation.RunInputValidationTests
            Case "TEST_Aggregation.RunAggregationTests"
                TEST_Aggregation.RunAggregationTests
            Case "TEST_Invariants.RunInvariantTests"
                TEST_Invariants.RunInvariantTests
            Case "TEST_Reconcile.RunReconcileTests"
                TEST_Reconcile.RunReconcileTests
            Case "TEST_CsvReader.RunCsvReaderTests"
                TEST_CsvReader.RunCsvReaderTests
            Case "TEST_Import.RunImportTests"
                TEST_Import.RunImportTests
        End Select

End Sub


Private Sub RunLogged( _
    ByVal LogPath As String, _
    ByVal Candidate As String, _
    ByVal EntryPoint As String)
'
'==============================================================================
'                                  RunLogged
'------------------------------------------------------------------------------
' PURPOSE
'   Run one entry point in normal mode with its log open: the CANDIDATE=
'   line first, then every line the run prints. An error it raises is
'   printed and logged as an ERROR= line.
'
' UPDATED
'   2026-11-04
'==============================================================================
'
        mLogPath = LogPath
        WriteLogLine "CANDIDATE=" & Candidate
        On Error GoTo Raised
        RunEntry EntryPoint
        mLogPath = ""
        Exit Sub

Raised:
        Emit "ERROR=" & Err.Number & ": " & Err.Description
        mLogPath = ""

End Sub


Private Sub Emit( _
    ByVal Text As String)
'
'==============================================================================
'                                     Emit
'------------------------------------------------------------------------------
' PURPOSE
'   Print one line to the Immediate window and, while Certify has a log
'   open, append it to that log.
'
' UPDATED
'   2026-11-04
'==============================================================================
'
        Debug.Print Text
        If Len(mLogPath) > 0 Then
            WriteLogLine Text
        End If

End Sub


Private Sub WriteLogLine( _
    ByVal Text As String)
'
'==============================================================================
'                                 WriteLogLine
'------------------------------------------------------------------------------
' PURPOSE
'   Append one line, as UTF-8 with CRLF, to the log Certify has open.
'
' ERROR POLICY
'   Contained: a line that cannot be written is counted in mLogErrors.
'
' UPDATED
'   2026-11-04
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Buffer()     As Byte       'The line as UTF-8
    Dim FileNumber   As Integer    'Open file; 0 for none

'------------------------------------------------------------------------------
' APPEND
'------------------------------------------------------------------------------
        On Error GoTo Failed
        Buffer = CORE_Export.Utf8Bytes(Text & vbCrLf)
        FileNumber = FreeFile
        Open mLogPath For Binary Access Write As #FileNumber
        Put #FileNumber, LOF(FileNumber) + 1, Buffer
        Close #FileNumber
        Exit Sub

Failed:
        mLogErrors = mLogErrors + 1
        If FileNumber <> 0 Then
            On Error Resume Next
            Close #FileNumber
        End If

End Sub


Private Function LogNameOf( _
    ByVal EntryPoint As String) _
    As String
'
'==============================================================================
'                                   LogNameOf
'------------------------------------------------------------------------------
' PURPOSE
'   The log file of an entry point in an evidence bundle: harness.log for
'   the harness, the module name with .log for a suite.
'
' UPDATED
'   2026-11-04
'==============================================================================
'
        If EntryPoint = "TEST_Harness.RunTests" Then
            LogNameOf = "harness.log"
        Else
            LogNameOf = Split(EntryPoint, ".")(0) & ".log"
        End If

End Function


Private Function IsCandidate( _
    ByVal Candidate As String) _
    As Boolean
'
'==============================================================================
'                                  IsCandidate
'------------------------------------------------------------------------------
' PURPOSE
'   Whether Candidate is a full commit SHA: 40 lowercase hexadecimal
'   characters.
'
' UPDATED
'   2026-11-04
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim i           As Long    'Character position

'------------------------------------------------------------------------------
' CHECK
'------------------------------------------------------------------------------
        If Len(Candidate) <> 40 Then
            Exit Function
        End If
        For i = 1 To 40
            If InStr(1, "0123456789abcdef", Mid$(Candidate, i, 1), vbBinaryCompare) = 0 Then
                Exit Function
            End If
        Next i
        IsCandidate = True

End Function


Private Function FolderExists( _
    ByVal Path As String) _
    As Boolean
'
'==============================================================================
'                                 FolderExists
'------------------------------------------------------------------------------
' PURPOSE
'   Whether Path names an existing folder.
'
' ERROR POLICY
'   Contained: a path that cannot be read is not a folder.
'
' UPDATED
'   2026-11-04
'==============================================================================
'
        On Error GoTo Missing
        FolderExists = (GetAttr(Path) And vbDirectory) = vbDirectory

Missing:
End Function


Private Function FileExists( _
    ByVal Path As String) _
    As Boolean
'
'==============================================================================
'                                  FileExists
'------------------------------------------------------------------------------
' PURPOSE
'   Whether Path names an existing file.
'
' ERROR POLICY
'   Contained: a path that cannot be read is not a file.
'
' UPDATED
'   2026-11-04
'==============================================================================
'
        On Error GoTo Missing
        FileExists = (GetAttr(Path) And vbDirectory) = 0

Missing:
End Function
