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
'   Test_RunAll runs the suites. Say is the print routine of every suite:
'   it writes to the Immediate window, or during Test_RunAll records the
'   RESULT and FAILURE lines instead. Option Private Module keeps both out
'   of the external workbook automation API.
'
' DEPENDENCIES
'   TEST_Harness, TEST_Formulas, TEST_MainState, TEST_Cases,
'   TEST_InputValidation, TEST_Aggregation and TEST_Invariants, with
'   TEST_CaseRunner for the last four.
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
'   2026-10-12
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
        Private Const SUITE_COUNT   As Long = 7    'Suites run by Test_RunAll

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
'   Run the seven suites in turn and print the summary.
'
' UPDATED
'   2026-10-12
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim suites     As Variant    'Entry points, as Module.Procedure
    Dim i          As Long       'Index into suites
    Dim failed     As Long       'Suites that did not pass
    Dim t0         As Single     'Start of the whole run
    Dim total      As Single     'Seconds for the whole run

'------------------------------------------------------------------------------
' RUN
'------------------------------------------------------------------------------
        suites = Array("TEST_Harness.RunTests", "TEST_Formulas.RunFormulaTests", _
                       "TEST_MainState.RunMainStateTests", "TEST_Cases.RunCaseTests", _
                       "TEST_InputValidation.RunInputValidationTests", _
                       "TEST_Aggregation.RunAggregationTests", "TEST_Invariants.RunInvariantTests")
        Debug.Print "SACCR TEST SUMMARY " & Format$(Now, "yyyy-mm-dd hh:nn:ss")
        t0 = Timer
        For i = LBound(suites) To UBound(suites)
            If Not RunOneSuite(CStr(suites(i))) Then
                failed = failed + 1
            End If
        Next i

'------------------------------------------------------------------------------
' REPORT
'------------------------------------------------------------------------------
        total = Timer - t0
        If total < 0 Then
            total = total + 86400!
        End If
        Debug.Print "RESULT=" & IIf(failed = 0, "PASS", "FAIL") & "; suites=" & SUITE_COUNT & _
                    "; passed=" & (SUITE_COUNT - failed) & "; failed=" & failed & _
                    "; seconds=" & Format$(total, "0.0")

End Sub


Public Sub Say( _
    ByVal text As String)
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
'   2026-10-12
'==============================================================================
'
        If Not mQuiet Then
            Debug.Print text
        ElseIf Left$(text, 7) = "RESULT=" Then
            mResult = text
        ElseIf Left$(text, 8) = "FAILURE=" Or Left$(text, 16) = "FAILURE_DETAILS=" Then
            mFailures.Add text
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
    ByVal entryPoint As String) _
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
'   2026-10-12
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim t0        As Single     'Start of the suite
    Dim seconds   As Single     'Duration of the suite
    Dim raised    As String     'Error raised by the suite; "" when none
    Dim passed    As Boolean    'RESULT=PASS and no error
    Dim detail    As String     'The RESULT line after its verdict
    Dim failure   As Variant    'One failure line

'------------------------------------------------------------------------------
' RUN QUIETLY
'------------------------------------------------------------------------------
        mResult = ""
        Set mFailures = New Collection
        mQuiet = True
        t0 = Timer
        On Error GoTo Raised
        Select Case entryPoint
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
        End Select

Report:
        On Error GoTo 0
        mQuiet = False
        seconds = Timer - t0
        If seconds < 0 Then
            seconds = seconds + 86400!
        End If

'------------------------------------------------------------------------------
' PRINT THE SUITE LINE
'------------------------------------------------------------------------------
    'The RESULT line without its verdict, which the first column shows.
        passed = (Left$(mResult, 11) = "RESULT=PASS") And Len(raised) = 0
        detail = mResult
        If InStr(detail, "; ") > 0 Then
            detail = Mid$(detail, InStr(detail, "; ") + 2)
        ElseIf Len(detail) = 0 Then
            detail = "no RESULT line"
        End If
        Debug.Print "  " & IIf(passed, "PASS", "FAIL") & "  " & Left$(Split(entryPoint, ".")(0) & Space$(22), 22) & _
                    detail & "  (" & Format$(seconds, "0.0") & " s)"
        If Not passed Then
            For Each failure In mFailures
                Debug.Print "        " & failure
            Next failure
            If Len(raised) > 0 Then
                Debug.Print "        ERROR=" & raised
            End If
        End If
        RunOneSuite = passed
        Exit Function

'------------------------------------------------------------------------------
' HANDLE A RAISED ERROR
'------------------------------------------------------------------------------
Raised:
        raised = Err.Number & ": " & Err.Description
        Resume Report

End Function
