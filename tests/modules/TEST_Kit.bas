Attribute VB_Name = "TEST_Kit"
'==============================================================================
' MODULE: TEST_Kit
'------------------------------------------------------------------------------
' PURPOSE
'   The lifecycle every Excel suite shares (#536): begin the suite, begin a
'   case, check, fail, end the suite. It owns the case, check and failure
'   counts, the completeness check against the counts the suite expects,
'   the check that the suite left Excel and the test seams as it found
'   them, and the one SUITE, MODE and RESULT line format of every suite.
'   It refuses to start a suite on a workbook that an earlier run left
'   dirty, instead of testing on top of it.
'
' PUBLIC SURFACE
'   BeginSuite, BeginCase, Check, Fail, CurrentCase, CheckCount and
'   EndSuite for the suites and TEST_CaseRunner; Refusal and RunningSuite
'   for the suite that tests the refusal; ClearAbortedRun, a recovery
'   command for a workbook whose last run did not finish. Option Private
'   Module keeps them out of the external workbook automation API.
'
' DEPENDENCIES
'   TEST_Suite.Say for the log lines; M_Main and CORE_Engine for the test
'   seams it checks (gSilent, gQuietLog, gTestFault, gKeepProtection and
'   gEngineFault).
'
' WORKSHEET SAFETY
'   Writes one hidden workbook name, RUN_MARK, while a suite runs, and
'   deletes it at the end of the suite. A name left behind shows that a
'   run did not finish. Nothing else in the workbook is written.
'
' STATE OWNERSHIP
'   Owns the counters below and the settings BeginSuite captures. EndSuite
'   compares the settings and the seams with what BeginSuite found, puts
'   back any that differ and reports them as a failed restoration.
'
' REPORTING
'   A suite prints SUITE=<module>, then MODE=NORMAL or MODE=INJECTED_FAILURE,
'   then one CASE line per case and one FAILURE line per failed check, and
'   ends with RESULT=PASS or FAIL; cases=...; checks=...; failures=...;
'   restore=PASS, FAIL or NOT_RUN.
'
' USAGE
'   In a suite's entry point:
'       If Not TEST_Kit.BeginSuite("TEST_Name", EXPECTED_CASES, EXPECTED_CHECKS) Then Exit Sub
'       ...cases, each starting with TEST_Kit.BeginCase "case-name"...
'       TEST_Kit.EndSuite SuiteRestored
'
' UPDATED
'   2026-11-11
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
        Public Const ANY_CHECKS     As Long = -1                    'A suite whose check count is not fixed
        Private Const RUN_MARK      As String = "TEST_SuiteRunning" 'Hidden name: the suite running
        Private Const RECOVERY      As String = "close the workbook without saving it, or build it again " & _
                                                "from the template; once you have checked it, " & _
                                                "TEST_Kit.ClearAbortedRun removes the mark"

'------------------------------------------------------------------------------
' MODULE STATE
'------------------------------------------------------------------------------
    'The suite being run.
        Private mActive          As Boolean    'Between BeginSuite and EndSuite
        Private mSuite           As String     'Its module name
        Private mExpectedCases   As Long       'Cases a complete run has
        Private mExpectedChecks  As Long       'Checks a complete run has; ANY_CHECKS for any
        Private mCaseCount       As Long       'Cases started
        Private mCheckCount      As Long       'Checks evaluated
        Private mFailureCount    As Long       'Failed checks, unexpected errors and suite failures
        Private mCurrentCase     As String     'Case being run, for messages

    'What BeginSuite found, for EndSuite.
        Private mCalculation     As XlCalculation    'Calculation mode
        Private mEvents          As Boolean          'EnableEvents
        Private mScreen          As Boolean          'ScreenUpdating
        Private mAlerts          As Boolean          'DisplayAlerts
        Private mSilent          As Boolean          'M_Main.gSilent
        Private mQuiet           As Boolean          'M_Main.gQuietLog


'
'------------------------------------------------------------------------------
'
'                                SUITE CONTROL
'
'------------------------------------------------------------------------------
'

Public Function BeginSuite( _
    ByVal SuiteName As String, _
    ByVal ExpectedCases As Long, _
    ByVal ExpectedChecks As Long, _
    Optional ByVal Mode As String = "NORMAL") _
    As Boolean
'
'==============================================================================
'                                  BeginSuite
'------------------------------------------------------------------------------
' PURPOSE
'   Start a suite: refuse a dirty workbook, reset the counts, capture the
'   settings EndSuite compares, mark the workbook and print the SUITE and
'   MODE lines.
'
' INPUTS
'   SuiteName: the suite's module name, as the evidence policy lists it.
'   ExpectedCases, ExpectedChecks: the counts of a complete run; a
'      different count fails the suite. ANY_CHECKS for a suite whose
'      checks are counted by its evidence only.
'   Mode: NORMAL, or INJECTED_FAILURE for TEST_CaseRunner.InjectFailure.
'
' RETURNS
'   True when the suite may run. False when the workbook is dirty: the
'   refusal has been printed, ending with RESULT=FAIL and restore=NOT_RUN,
'   nothing has been changed, and the suite must not call EndSuite.
'
' UPDATED
'   2026-11-11
'==============================================================================
'
    Dim Refused    As String    'The refusal lines; "" when the workbook is clean
    Dim Item       As Variant   'One of them

        Refused = Refusal(SuiteName, RunningSuite(), mActive, TestSeams())
        If Len(Refused) > 0 Then
            For Each Item In Split(Refused, vbLf)
                TEST_Suite.Say CStr(Item)
            Next Item
            Exit Function
        End If
        mActive = True
        mSuite = SuiteName
        mExpectedCases = ExpectedCases
        mExpectedChecks = ExpectedChecks
        mCaseCount = 0
        mCheckCount = 0
        mFailureCount = 0
        mCurrentCase = "suite"
        mCalculation = Application.Calculation
        mEvents = Application.EnableEvents
        mScreen = Application.ScreenUpdating
        mAlerts = Application.DisplayAlerts
        mSilent = M_Main.gSilent
        mQuiet = M_Main.gQuietLog
        ThisWorkbook.Names.Add Name:=RUN_MARK, RefersTo:="=""" & SuiteName & """", Visible:=False
        TEST_Suite.Say "SUITE=" & SuiteName
        TEST_Suite.Say "MODE=" & Mode
        BeginSuite = True

End Function


Public Sub EndSuite( _
    Optional ByVal SuiteRestored As Boolean = True)
'
'==============================================================================
'                                   EndSuite
'------------------------------------------------------------------------------
' PURPOSE
'   End the suite: check that every expected case and check ran, that the
'   settings and test seams are as BeginSuite found them, remove the run
'   mark, and print the RESULT line.
'
' INPUTS
'   SuiteRestored: False when the suite could not put back what it
'      changed itself; it has printed the reason.
'
' ERROR POLICY
'   A setting found changed is put back, and the restoration fails. A mark
'   that cannot be removed fails the restoration too.
'
' UPDATED
'   2026-11-11
'==============================================================================
'
    Dim Restored   As Boolean    'Every restoration held

        If Not mActive Then
            Exit Sub
        End If
        If mCaseCount <> mExpectedCases Then
            Fail "suite", "expected " & mExpectedCases & " cases, ran " & mCaseCount
        End If
        If mExpectedChecks <> ANY_CHECKS And mCheckCount <> mExpectedChecks Then
            Fail "suite", "expected " & mExpectedChecks & " checks, ran " & mCheckCount
        End If
        Restored = SuiteRestored
        Restored = SettingsRestored() And Restored
        Restored = MarkRemoved() And Restored
        mActive = False
        TEST_Suite.Say "RESULT=" & IIf(mFailureCount = 0 And Restored, "PASS", "FAIL") & _
                    "; cases=" & mCaseCount & "; checks=" & mCheckCount & _
                    "; failures=" & mFailureCount & "; restore=" & IIf(Restored, "PASS", "FAIL")

End Sub


'
'------------------------------------------------------------------------------
'
'                                CASES AND CHECKS
'
'------------------------------------------------------------------------------
'

Public Sub BeginCase( _
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


Public Sub Check( _
    ByVal Passed As Boolean, _
    ByVal Detail As String, _
    Optional ByVal Context As String = "")
'
'==============================================================================
' PURPOSE
'   Count one check and record it when it fails, against Context, or the
'   current case when Context is empty.
'==============================================================================
'
        mCheckCount = mCheckCount + 1
        If Not Passed Then
            Fail IIf(Len(Context) = 0, mCurrentCase, Context), Detail
        End If

End Sub


Public Sub Fail( _
    ByVal Context As String, _
    ByVal Detail As String)
'
'==============================================================================
' PURPOSE
'   Count a failure that is not a check, such as an unexpected error, and
'   print it as FAILURE=Context: Detail.
'==============================================================================
'
        mFailureCount = mFailureCount + 1
        TEST_Suite.Say "FAILURE=" & Context & ": " & Detail

End Sub


Public Function CurrentCase() As String
'
'==============================================================================
' PURPOSE
'   The case being run, for a suite's own messages.
'==============================================================================
'
        CurrentCase = mCurrentCase

End Function


Public Function CheckCount() As Long
'
'==============================================================================
' PURPOSE
'   The checks evaluated so far in the running suite.
'==============================================================================
'
        CheckCount = mCheckCount

End Function


'
'------------------------------------------------------------------------------
'
'                                 DIRTY WORKBOOK
'
'------------------------------------------------------------------------------
'

Public Function Refusal( _
    ByVal SuiteName As String, _
    ByVal MarkedSuite As String, _
    ByVal SuiteActive As Boolean, _
    ByVal ArmedSeams As String) _
    As String
'
'==============================================================================
'                                   Refusal
'------------------------------------------------------------------------------
' PURPOSE
'   The lines a suite prints instead of running when the workbook is
'   dirty, separated by vbLf; "" when it is clean. BeginSuite passes the
'   workbook's own state; TEST_MainState passes a chosen one.
'
' INPUTS
'   SuiteName: the suite asked to start.
'   MarkedSuite: the suite named by the run mark; "" for no mark.
'   SuiteActive: a suite of this session has not ended.
'   ArmedSeams: the test seams left armed, as TestSeams lists them.
'
' UPDATED
'   2026-11-11
'==============================================================================
'
    Dim Reason   As String    'Why the workbook is dirty

        If Len(MarkedSuite) > 0 Then
            Reason = "the run of " & MarkedSuite & " did not finish, so the workbook may still hold " & _
                     "its test data: " & RECOVERY
        ElseIf SuiteActive Then
            Reason = "the suite " & mSuite & " is still running"
        ElseIf Len(ArmedSeams) > 0 Then
            Reason = "test seams are still armed (" & ArmedSeams & "): " & RECOVERY
        Else
            Exit Function
        End If
        Refusal = "SUITE=" & SuiteName & vbLf & "MODE=NORMAL" & vbLf & _
                  "FAILURE=suite: refused, the workbook is dirty: " & Reason & vbLf & _
                  "RESULT=FAIL; cases=0; checks=0; failures=1; restore=NOT_RUN"

End Function


Public Function RunningSuite() As String
'
'==============================================================================
' PURPOSE
'   The suite the run mark names; "" when there is no mark.
'==============================================================================
'
    Dim Mark   As Name    'The run mark

        On Error Resume Next
        Set Mark = ThisWorkbook.Names(RUN_MARK)
        On Error GoTo 0
        If Not Mark Is Nothing Then
            RunningSuite = Replace(Mid$(Mark.RefersTo, 3, Len(Mark.RefersTo) - 3), """""", """")
        End If

End Function


Public Sub ClearAbortedRun()
'
'==============================================================================
'                               ClearAbortedRun
'------------------------------------------------------------------------------
' PURPOSE
'   Recovery command for a workbook whose last run did not finish, once
'   it has been checked: remove the run mark, disarm the test seams and
'   forget the unfinished suite, so that the next suite may start. Type
'   TEST_Kit.ClearAbortedRun in the Immediate window.
'
' UPDATED
'   2026-11-11
'==============================================================================
'
        On Error Resume Next
        ThisWorkbook.Names(RUN_MARK).Delete
        On Error GoTo 0
        M_Main.gTestFault = ""
        M_Main.gKeepProtection = ""
        CORE_Engine.gEngineFault = ""
        mActive = False
        TEST_Suite.Say "RESULT=CLEARED; mark=" & IIf(Len(RunningSuite()) = 0, "NONE", "KEPT")

End Sub


'
'------------------------------------------------------------------------------
'
'                                   HELPERS
'
'------------------------------------------------------------------------------
'

Private Function TestSeams() As String
'
'==============================================================================
' PURPOSE
'   The test seams that are armed, as "name=value" separated by ", ";
'   "" when none is. A suite disarms every seam it arms, so an armed seam
'   outside a suite was left by a run that stopped.
'==============================================================================
'
    Dim Text   As String    'The seams found so far

        If Len(M_Main.gTestFault) > 0 Then Text = Text & ", gTestFault=" & M_Main.gTestFault
        If Len(M_Main.gKeepProtection) > 0 Then Text = Text & ", gKeepProtection=" & M_Main.gKeepProtection
        If Len(CORE_Engine.gEngineFault) > 0 Then Text = Text & ", gEngineFault=" & CORE_Engine.gEngineFault
        If Len(Text) > 0 Then TestSeams = Mid$(Text, 3)

End Function


Private Function SettingsRestored() As Boolean
'
'==============================================================================
' PURPOSE
'   Compare the settings and seams with what BeginSuite found; print each
'   difference as a failure of the restoration and put it back.
'==============================================================================
'
    Dim Changed   As String    'The differences found

        If Application.Calculation <> mCalculation Then
            Changed = Changed & ", calculation " & mCalculation & " -> " & Application.Calculation
            Application.Calculation = mCalculation
        End If
        If Application.EnableEvents <> mEvents Then
            Changed = Changed & ", events " & mEvents & " -> " & Application.EnableEvents
            Application.EnableEvents = mEvents
        End If
        If Application.ScreenUpdating <> mScreen Then
            Changed = Changed & ", screen updating " & mScreen & " -> " & Application.ScreenUpdating
            Application.ScreenUpdating = mScreen
        End If
        If Application.DisplayAlerts <> mAlerts Then
            Changed = Changed & ", display alerts " & mAlerts & " -> " & Application.DisplayAlerts
            Application.DisplayAlerts = mAlerts
        End If
        If M_Main.gSilent <> mSilent Then
            Changed = Changed & ", gSilent " & mSilent & " -> " & M_Main.gSilent
            M_Main.gSilent = mSilent
        End If
        If M_Main.gQuietLog <> mQuiet Then
            Changed = Changed & ", gQuietLog " & mQuiet & " -> " & M_Main.gQuietLog
            M_Main.gQuietLog = mQuiet
        End If
        If Len(TestSeams()) > 0 Then
            Changed = Changed & ", armed " & TestSeams()
            M_Main.gTestFault = ""
            M_Main.gKeepProtection = ""
            CORE_Engine.gEngineFault = ""
        End If
        If Len(Changed) > 0 Then
            Fail "restore", "left changed and put back: " & Mid$(Changed, 3)
        End If
        SettingsRestored = (Len(Changed) = 0)

End Function


Private Function MarkRemoved() As Boolean
'
'==============================================================================
' PURPOSE
'   Delete the run mark; a mark that stays fails the restoration.
'==============================================================================
'
        On Error Resume Next
        ThisWorkbook.Names(RUN_MARK).Delete
        On Error GoTo 0
        If Len(RunningSuite()) > 0 Then
            Fail "restore", "the run mark " & RUN_MARK & " could not be removed"
            Exit Function
        End If
        MarkRemoved = True

End Function
