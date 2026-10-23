Attribute VB_Name = "TEST_Suite"
'==============================================================================
' MODULE: TEST_Suite
'------------------------------------------------------------------------------
' PURPOSE
'   Run every test suite in one go and print only a summary: one line per
'   suite with its verdict, counts and duration, the failure lines of any
'   suite that failed, and a final RESULT line. The suites still print
'   their full logs when they are run on their own.
'
' PUBLIC SURFACE
'   Test_RunAll runs the suites. Say is the print routine of every suite,
'   of TEST_CaseRunner and of TEST_Benchmark: it writes to the Immediate
'   window, or during Test_RunAll records the RESULT and FAILURE lines
'   instead. Option Private Module keeps both out
'   of the external workbook automation API.
'
' DEPENDENCIES
'   TEST_Harness, TEST_Formulas, TEST_MainState, TEST_Cases,
'   TEST_InputValidation, TEST_Aggregation, TEST_Invariants,
'   TEST_Reconcile and TEST_CsvReader, with TEST_CaseRunner for TEST_Cases,
'   TEST_InputValidation, TEST_Aggregation and TEST_Invariants;
'   M_Main.gQuietLog, set while a suite runs, so that the macros'
'   silent-mode messages are not printed either.
'
' WORKSHEET SAFETY
'   As the suites it runs: they rewrite inputs, outputs and Params and
'   restore them. Use a development workbook.
'
' ERROR POLICY
'   An error raised by a suite is recorded as that suite's failure and the
'   remaining suites still run.
'
' USAGE
'   Import with the test modules, compile, then type Test_RunAll in the
'   Immediate window.
'
' UPDATED
'   2026-10-23
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
        Private Const SUITE_COUNT   As Long = 9    'Suites run by Test_RunAll

'------------------------------------------------------------------------------
' MODULE STATE
'------------------------------------------------------------------------------
        Private mQuiet      As Boolean       'True while Test_RunAll runs a suite
        Private mResult     As String        'Last RESULT line of the running suite
        Private mFailures   As Collection    'FAILURE lines of the running suite


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
        Suites = Array("TEST_Harness.RunTests", "TEST_Formulas.RunFormulaTests", _
                       "TEST_MainState.RunMainStateTests", "TEST_Cases.RunCaseTests", _
                       "TEST_InputValidation.RunInputValidationTests", _
                       "TEST_Aggregation.RunAggregationTests", "TEST_Invariants.RunInvariantTests", _
                       "TEST_Reconcile.RunReconcileTests", "TEST_CsvReader.RunCsvReaderTests")
        Debug.Print "SACCR TEST SUMMARY " & Format$(Now, "yyyy-mm-dd hh:nn:ss")
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
        Debug.Print "RESULT=" & IIf(FailedCount = 0, "PASS", "FAIL") & "; suites=" & SUITE_COUNT & _
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
'
' UPDATED
'   2026-10-17
'==============================================================================
'
        If Not mQuiet Then
            Debug.Print Text
        ElseIf Left$(Text, 7) = "RESULT=" Then
            mResult = Text
        ElseIf Left$(Text, 8) = "FAILURE=" Or Left$(Text, 16) = "FAILURE_DETAILS=" Then
            mFailures.Add Text
        End If

End Sub


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
        End Select

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
        Debug.Print "  " & IIf(Passed, "PASS", "FAIL") & "  " & Left$(Split(EntryPoint, ".")(0) & Space$(22), 22) & _
                    Detail & "  (" & Format$(SuiteSeconds, "0.0") & " s)"
        If Not Passed Then
            For Each FailureLine In mFailures
                Debug.Print "        " & FailureLine
            Next FailureLine
            If Len(RaisedError) > 0 Then
                Debug.Print "        ERROR=" & RaisedError
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
