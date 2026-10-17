Attribute VB_Name = "TEST_Harness"
'==============================================================================
' MODULE: TEST_Harness
'------------------------------------------------------------------------------
' PURPOSE
'   Run a deterministic, dependency-free regression suite and report complete,
'   machine-checkable evidence to the Immediate window. The cases exercise the
'   SA-CCR worksheet functions in SACCR_Formulas; the engine and the workbook
'   macros are tested by the other TEST_ suites, which TEST_Suite.Test_RunAll
'   runs together with this one.
'
' PUBLIC SURFACE
'   RunTests is the documented test entry point. RunTestsWithInjectedFailure
'   runs the same suite with one deliberately wrong expectation, to show what a
'   failing run reports. ResetTests is a recovery command for an interrupted
'   run. Option Private Module keeps all three out of the external workbook
'   automation API.
'
' DEPENDENCIES
'   SACCR_Formulas, TEST_Suite.Say for its log lines, and the built-in
'   VBA/Excel object models only. No external references, workbook fixture,
'   worksheet, donor project, or test framework.
'
' STATE OWNERSHIP
'   Owns private counters, failure text, suite-completeness state, and a re-entry
'   flag. Cleanup clears the re-entry flag; report state is retained until the
'   next run/reset so the final summary remains inspectable.
'
' ERROR POLICY
'   Expected worksheet error values are asserted. Unexpected runner errors
'   are recorded, cleanup runs, and the original number/source/description is
'   re-raised. Assertion failures raise one test-suite error after reporting.
'
' WORKSHEET SAFETY
'   Reads selected Application properties to report the environment and prove
'   calculation, display-alert, event, and screen-updating state did not change.
'   It never creates, selects, activates, edits, calculates, saves, or closes
'   workbook or worksheet state.
'
' TEST SEAM
'   Tests only the supported SACCR_Formulas surface listed in
'   docs/PUBLIC_API.txt. Core procedures are tested by the other suites,
'   without adding production API.
'
' COMPATIBILITY
'   Excel VBA on Windows; host evidence identifies the actual Office bitness
'   and VBA generation. No external framework or workbook fixture is required.
'
' USAGE
'   Import the required production modules first, then run
'   TEST_Harness.RunTests from the VBE Immediate window.
'
' UPDATED
'   2026-10-17
'
' AUTHOR
'   Daniele Penza
'==============================================================================

'------------------------------------------------------------------------------
' MODULE SETTINGS
'------------------------------------------------------------------------------
    'Require explicit declarations; preserve the configured component visibility.
    Option Explicit
    Option Private Module

'------------------------------------------------------------------------------
' MODULE CONSTANTS
'------------------------------------------------------------------------------
    'Keep harness error codes separate from the production failure contract;
    'the expected counts define when the suite is complete.
        Private Const TEST_ERROR_DIRTY_START   As Long = vbObjectError + 2060    'Refused re-entry or interrupted run
        Private Const TEST_ERROR_FAILURES      As Long = vbObjectError + 2061    'Failed or incomplete suite outcome
        Private Const EXPECTED_ASSERTIONS      As Long = 6                       'Required assertions for a complete run
        Private Const EXPECTED_CASES           As Long = 4                       'Required cases for a complete run

'------------------------------------------------------------------------------
' MODULE STATE
'------------------------------------------------------------------------------
    'Own counters and diagnostics for one run; cleanup releases the active
    'flag while report values remain available until the next reset.
        Private mCaseCount        As Long       'Cases started in the current run
        Private mAssertionCount   As Long       'Assertions evaluated in the current run
        Private mFailureCount     As Long       'Failures recorded, including runner errors
        Private mFailureDetails   As String     'Diagnostics retained until the next reset
        Private mInjectFailure    As Boolean    'True only inside RunTestsWithInjectedFailure
        Private mRunActive        As Boolean    'Re-entry guard; cleared by cleanup or reset
        Private mSuiteCompleted   As Boolean    'True only when both expected counts match


'
'------------------------------------------------------------------------------
'
'                              TEST ENTRY POINTS
'
'------------------------------------------------------------------------------
'

Public Sub RunTests()
'
'==============================================================================
'                                   RunTests
'------------------------------------------------------------------------------
' PURPOSE
'   Run all four cases and report six assertions with cleanup evidence.
'
' USAGE
'   Run TEST_Harness.RunTests from the VBE Immediate window.
'
' STATE OWNERSHIP
'   Reject an active run before resetting counters. Snapshot four Excel
'   properties for comparison; the harness never changes those properties.
'
' ERROR POLICY
'   Record unexpected runner failures, attempt cleanup, print the summary,
'   then re-raise the saved error. Other failed outcomes raise the suite error.
'   A dirty start raises immediately and does not run cleanup.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    'Keep the host snapshot separate from the saved runner error so cleanup
    'can be reported without losing the original failure.
    Dim CleanupDetail           As String           'Cleanup explanation included in the report
    Dim CleanupPassed           As Boolean          'True only when cleanup verification succeeds
    Dim InitialCalculation      As XlCalculation    'Calculation mode captured before the suite
    Dim InitialDisplayAlerts    As Boolean          'Display-alert setting captured before the suite
    Dim InitialEnableEvents     As Boolean          'Event setting captured before the suite
    Dim InitialScreenUpdating   As Boolean          'Screen-update setting captured before the suite
    Dim SavedDescription        As String           'Original diagnostic preserved through cleanup
    Dim SavedNumber             As Long             'Original error number for later re-raise
    Dim SavedSource             As String           'Original runner error source for re-raise
    Dim StateSnapshotValid      As Boolean          'True only after all four properties were read

'------------------------------------------------------------------------------
' GUARD ENTRY
'------------------------------------------------------------------------------
    'Refuse an active run before clearing its counters or failure details.
    'An interrupted execution must be reset explicitly by the caller.
        If mRunActive Then
            TEST_Suite.Say "RESULT=FAIL_DIRTY_START; cleanup=NOT_RUN"
            Err.Raise _
                TEST_ERROR_DIRTY_START, _
                "TEST_Harness.RunTests", _
                "A TEST_Harness run is already active. Run ResetTests after an interrupted execution."
        End If

'------------------------------------------------------------------------------
' INITIALIZE RUN
'------------------------------------------------------------------------------
    'Start with empty report state, then claim the run before executing
    'anything that can fail through the shared runner handler.
        ResetRun
        mRunActive = True
        On Error GoTo RunFailed

'------------------------------------------------------------------------------
' SNAPSHOT HOST STATE
'------------------------------------------------------------------------------
    'Read all four properties before marking the snapshot valid. If a read
    'fails, cleanup must report an unavailable snapshot rather than compare
    'uninitialized values or claim that host state was unchanged.
        InitialCalculation = Application.Calculation
        InitialDisplayAlerts = Application.DisplayAlerts
        InitialEnableEvents = Application.EnableEvents
        InitialScreenUpdating = Application.ScreenUpdating
        StateSnapshotValid = True

'------------------------------------------------------------------------------
' RUN SUITE
'------------------------------------------------------------------------------
    'Run cases in the evidence contract order; each case records unexpected
    'errors so later cases can still contribute to the report.
        PrintEnvironment
        TestReplacementCost
        TestMaturityFactor
        TestCdoDeltaError
        TestRepeatability

    'Require both counts so an early return cannot produce a complete PASS.
        mSuiteCompleted = _
            (mCaseCount = EXPECTED_CASES) And _
            (mAssertionCount = EXPECTED_ASSERTIONS)
        If Not mSuiteCompleted Then
            RecordFailure _
                "suite.completeness", _
                "expected cases=" & CStr(EXPECTED_CASES) & _
                    ", assertions=" & CStr(EXPECTED_ASSERTIONS) & _
                    "; observed cases=" & CStr(mCaseCount) & _
                    ", assertions=" & CStr(mAssertionCount)
        End If

'------------------------------------------------------------------------------
' CLEANUP AND REPORT
'------------------------------------------------------------------------------
CleanExit:
    'Disable the runner handler before cleanup and reporting. A later raise
    'must escape to the caller instead of re-entering the runner handler.
        On Error GoTo 0
        CleanupPassed = CleanupRun( _
            CleanupDetail, _
            StateSnapshotValid, _
            InitialCalculation, _
            InitialDisplayAlerts, _
            InitialEnableEvents, _
            InitialScreenUpdating)
        PrintSummary CleanupPassed, CleanupDetail

    'Propagate the original runner error only after cleanup and reporting.
        If SavedNumber <> 0 Then
            Err.Raise SavedNumber, SavedSource, SavedDescription
        End If

    'Make assertion, cleanup and completeness failures visible to callers
    'even when no unexpected runner error was saved.
        If mFailureCount <> 0 Or Not CleanupPassed Or Not mSuiteCompleted Then
            Err.Raise _
                TEST_ERROR_FAILURES, _
                "TEST_Harness.RunTests", _
                "Regression failed; review the Immediate window report."
        End If
        Exit Sub

'------------------------------------------------------------------------------
' HANDLE RUNNER ERROR
'------------------------------------------------------------------------------
RunFailed:
    'Save the escaping error before recording its diagnostic; Resume then
    'leaves the active handler through the common cleanup path.
        SavedNumber = Err.Number
        SavedSource = Err.Source
        SavedDescription = Err.Description
        RecordFailure _
            "runner.unexpected", _
            "error=" & CStr(SavedNumber) & _
                "; source=" & SavedSource & _
                "; description=" & SavedDescription
        Resume CleanExit

End Sub


Public Sub ResetTests()
'
'==============================================================================
'                                  ResetTests
'------------------------------------------------------------------------------
' PURPOSE
'   Recover harness-owned state after an interrupted execution.
'
' USAGE
'   Run only after the previous execution has stopped; then rerun the suite.
'
' SIDE EFFECTS
'   Clear counters, failure text, completion, the active-run flag and the
'   injected-failure flag. Print confirmation; do not alter Excel application
'   properties.
'
' UPDATED
'   2026-10-06
'==============================================================================
'

'------------------------------------------------------------------------------
' RESET
'------------------------------------------------------------------------------
    'Recover only harness-owned state after execution has stopped; the
    'confirmation distinguishes an explicit reset from a completed test run.
        ResetRun
        mInjectFailure = False
        TEST_Suite.Say "SACCR TESTS RESET; run_active=no; counters=0; injection=off"

End Sub


Public Sub RunTestsWithInjectedFailure()
'
'==============================================================================
'                         RunTestsWithInjectedFailure
'------------------------------------------------------------------------------
' PURPOSE
'   Demonstrate the failure path: run the full suite with one deliberately
'   wrong expectation in replacement-cost.exact.
'
' USAGE
'   Run TEST_Harness.RunTestsWithInjectedFailure from the VBE Immediate window.
'   Expected report: MODE=INJECTED_FAILURE and
'   RESULT=FAIL; completeness=COMPLETE; cases=4; assertions=6; failures=1;
'   cleanup=PASS, followed by the suite failure error.
'
' STATE OWNERSHIP
'   Sets the injection flag only for this run and always clears it again,
'   whether RunTests returns or raises.
'
' ERROR POLICY
'   Re-raise whatever RunTests raised, after clearing the flag. The expected
'   outcome is the suite failure error.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    'Keep the suite error so it survives clearing the injection flag.
    Dim SavedDescription   As String    'Suite error description for re-raise
    Dim SavedNumber        As Long      'Suite error number for re-raise
    Dim SavedSource        As String    'Suite error source for re-raise

'------------------------------------------------------------------------------
' RUN WITH INJECTION
'------------------------------------------------------------------------------
    'Enable the wrong expectation, then run the normal suite unchanged.
        mInjectFailure = True
        On Error GoTo RestoreMode

        RunTests

'------------------------------------------------------------------------------
' RESTORE MODE
'------------------------------------------------------------------------------
RestoreMode:
    'Capture any suite error before clearing the flag, then re-raise it so the
    'caller still sees the failed outcome.
        SavedNumber = Err.Number
        SavedSource = Err.Source
        SavedDescription = Err.Description
        On Error GoTo 0
        mInjectFailure = False
        If SavedNumber <> 0 Then
            Err.Raise SavedNumber, SavedSource, SavedDescription
        End If

End Sub


'
'------------------------------------------------------------------------------
'
'                               REGRESSION CASES
'
'------------------------------------------------------------------------------
'

Private Sub TestReplacementCost()
'
'==============================================================================
'                             TestReplacementCost
'------------------------------------------------------------------------------
' PURPOSE
'   Check unmargined and margined replacement cost [CRE52.10, CRE52.18] on
'   binary-exact inputs.
'
' ERROR POLICY
'   Record unexpected case errors and let the remaining suite continue.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    'Hold the expectation separately so injected-failure mode can change it.
    Dim ExpectedCost   As Double    'Expected unmargined replacement cost

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
    'Integer inputs keep both results exact, so equality is required without
    'a tolerance. Unmargined: max(100 - 30, 0) = 70. Margined:
    'max(10 - 20, 50 + 5 - 15, 0) = 40. In injected-failure mode the first
    'expectation is deliberately wrong, so that assertion must fail.
        On Error GoTo CaseFailed

        BeginCase "replacement-cost.exact"
        ExpectedCost = 70#
        If mInjectFailure Then ExpectedCost = 71#
        AssertEqualDouble _
            "SACCR_ReplacementCost(100, 30)", _
            ExpectedCost, _
            SACCR_Formulas.SACCR_ReplacementCost(100#, 30#)
        AssertEqualDouble _
            "SACCR_ReplacementCost(10, 20, margined, TH 50, MTA 5, NICA 15)", _
            40#, _
            SACCR_Formulas.SACCR_ReplacementCost(10#, 20#, True, 50#, 5#, 15#)
        Exit Sub

'------------------------------------------------------------------------------
' HANDLE CASE ERROR
'------------------------------------------------------------------------------
CaseFailed:
    'Record this case as failed without aborting the remaining cases.
        RecordUnexpectedCaseError "replacement-cost.exact"

End Sub


Private Sub TestMaturityFactor()
'
'==============================================================================
'                              TestMaturityFactor
'------------------------------------------------------------------------------
' PURPOSE
'   Check the margined and unmargined maturity factor [CRE52.48, CRE52.52]
'   against explicit absolute tolerances.
'
' ERROR POLICY
'   Record unexpected case errors and let the remaining suite continue.
'
' UPDATED
'   2026-10-06
'==============================================================================
'

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
    'Both results involve a square root, so they are compared with an
    'explicit bound. Margined, MPOR 10 business days: 1.5 * sqrt(10/250) = 0.3.
    'Unmargined, M = 0.5 years: sqrt(0.5).
        On Error GoTo CaseFailed

        BeginCase "maturity-factor.tolerance"
        AssertNear _
            "SACCR_MaturityFactor(margined, MPOR 10)", _
            0.3, _
            SACCR_Formulas.SACCR_MaturityFactor(0#, True, 10#), _
            0.000000000001
        AssertNear _
            "SACCR_MaturityFactor(0.5)", _
            0.707106781186548, _
            SACCR_Formulas.SACCR_MaturityFactor(0.5), _
            0.000000000001
        Exit Sub

'------------------------------------------------------------------------------
' HANDLE CASE ERROR
'------------------------------------------------------------------------------
CaseFailed:
    'Keep the unexpected failure associated with the tolerance case.
        RecordUnexpectedCaseError "maturity-factor.tolerance"

End Sub


Private Sub TestCdoDeltaError()
'
'==============================================================================
'                              TestCdoDeltaError
'------------------------------------------------------------------------------
' PURPOSE
'   Verify that an invalid tranche, attachment not below detachment, returns
'   the worksheet error #NUM! [CRE52.41].
'
' ERROR POLICY
'   The function signals invalid input with an error value, not a raised
'   error; a raised error is recorded as unexpected.
'
' UPDATED
'   2026-10-06
'==============================================================================
'

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
    'Attachment 0.5 above detachment 0.3 is not a valid tranche.
        On Error GoTo CaseFailed

        BeginCase "cdo-delta.invalid-tranche"
        AssertErrorValue _
            "SACCR_CDODelta(0.5, 0.3)", _
            xlErrNum, _
            SACCR_Formulas.SACCR_CDODelta(0.5, 0.3, True)
        Exit Sub

'------------------------------------------------------------------------------
' HANDLE CASE ERROR
'------------------------------------------------------------------------------
CaseFailed:
    'Treat a raised error as unexpected: the contract is an error value.
        RecordUnexpectedCaseError "cdo-delta.invalid-tranche"

End Sub


Private Sub TestRepeatability()
'
'==============================================================================
'                              TestRepeatability
'------------------------------------------------------------------------------
' PURPOSE
'   Verify identical option-delta results for two calls with the same inputs.
'
' ERROR POLICY
'   Record unexpected case errors and let the remaining suite continue.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    'Retain both results so the comparison observes two separate calls.
    Dim FirstResult    As Double    'Baseline result for the repeatability check
    Dim SecondResult   As Double    'Result of the identical second call

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
    'Call the function twice with identical inputs to detect a result that
    'depends on residual state from an earlier invocation. An error value
    'fails the conversion and is recorded as unexpected.
        On Error GoTo CaseFailed

        BeginCase "option-delta.repeatability"
        FirstResult = CDbl(SACCR_Formulas.SACCR_OptionDelta(100#, 100#, 1#, 0.2, True, True))
        SecondResult = CDbl(SACCR_Formulas.SACCR_OptionDelta(100#, 100#, 1#, 0.2, True, True))
        AssertEqualDouble "Repeated calls", FirstResult, SecondResult
        Exit Sub

'------------------------------------------------------------------------------
' HANDLE CASE ERROR
'------------------------------------------------------------------------------
CaseFailed:
    'Report the failed repeatability case and allow suite finalization.
        RecordUnexpectedCaseError "option-delta.repeatability"

End Sub


'
'------------------------------------------------------------------------------
'
'                          CASE AND ASSERTION HELPERS
'
'------------------------------------------------------------------------------
'

Private Sub BeginCase( _
    ByVal CaseName As String)
'
'==============================================================================
'                                  BeginCase
'------------------------------------------------------------------------------
' PURPOSE
'   Register a started case in both the counter and machine-readable log.
'
' INPUTS
'   CaseName: stable identifier used by the evidence contract.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' REGISTER CASE
'------------------------------------------------------------------------------
    'Count a case when it starts, even if its first assertion later fails;
    'the matching log record identifies which case was reached.
        mCaseCount = mCaseCount + 1
        TEST_Suite.Say "CASE=" & CaseName

End Sub


Private Sub AssertEqualDouble( _
    ByVal AssertionName As String, _
    ByVal Expected As Double, _
    ByVal Actual As Double)
'
'==============================================================================
'                              AssertEqualDouble
'------------------------------------------------------------------------------
' PURPOSE
'   Count one assertion and require exact Double equality.
'
' INPUTS
'   AssertionName identifies the check; Expected and Actual are compared.
'
' ERROR POLICY
'   Append a mismatch to the run report without raising an assertion error.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' ASSERT
'------------------------------------------------------------------------------
    'Count the assertion before evaluating it, then record any exact-value
    'mismatch without terminating the suite.
        mAssertionCount = mAssertionCount + 1
        If Actual <> Expected Then
            RecordFailure _
                AssertionName, _
                "expected=" & CStr(Expected) & "; actual=" & CStr(Actual)
        End If

End Sub


Private Sub AssertNear( _
    ByVal AssertionName As String, _
    ByVal Expected As Double, _
    ByVal Actual As Double, _
    ByVal Tolerance As Double)
'
'==============================================================================
'                                  AssertNear
'------------------------------------------------------------------------------
' PURPOSE
'   Count one assertion and apply an absolute Double tolerance.
'
' INPUTS
'   AssertionName identifies the check; Expected and Actual are compared.
'   Tolerance must be nonnegative and is inclusive at the boundary.
'
' ERROR POLICY
'   Record a negative tolerance or an excessive difference as one failure.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' ASSERT
'------------------------------------------------------------------------------
    'Use an absolute error bound. A negative tolerance is a failed test
    'contract; equality at the permitted bound remains a pass.
        mAssertionCount = mAssertionCount + 1
        If Tolerance < 0# Then
            RecordFailure AssertionName, "Tolerance must not be negative."
        ElseIf Abs(Actual - Expected) > Tolerance Then
            RecordFailure _
                AssertionName, _
                "expected=" & CStr(Expected) & _
                    "; actual=" & CStr(Actual) & _
                    "; tolerance=" & CStr(Tolerance)
        End If

End Sub


Private Sub AssertErrorValue( _
    ByVal AssertionName As String, _
    ByVal ExpectedCode As Long, _
    ByVal Actual As Variant)
'
'==============================================================================
'                               AssertErrorValue
'------------------------------------------------------------------------------
' PURPOSE
'   Count one assertion and require a specific worksheet error value.
'
' INPUTS
'   AssertionName identifies the check; ExpectedCode is an xlErr constant;
'   Actual is the value returned by the function under test.
'
' ERROR POLICY
'   Append a mismatch to the run report without raising an assertion error.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' ASSERT
'------------------------------------------------------------------------------
    'Check the type before comparing: comparing a number with an error value
    'raises instead of returning False.
        mAssertionCount = mAssertionCount + 1
        If Not IsError(Actual) Then
            RecordFailure _
                AssertionName, _
                "expected error " & CStr(ExpectedCode) & "; actual type=" & TypeName(Actual)
        ElseIf Actual <> CVErr(ExpectedCode) Then
            RecordFailure _
                AssertionName, _
                "expected error " & CStr(ExpectedCode) & "; actual=" & CStr(Actual)
        End If

End Sub


'
'------------------------------------------------------------------------------
'
'                              FAILURE RECORDING
'
'------------------------------------------------------------------------------
'

Private Sub RecordUnexpectedCaseError( _
    ByVal CaseName As String)
'
'==============================================================================
'                          RecordUnexpectedCaseError
'------------------------------------------------------------------------------
' PURPOSE
'   Preserve the active case error before recording its diagnostic fields.
'
' INPUTS
'   CaseName: stable case identifier; Err supplies the original failure.
'
' SIDE EFFECTS
'   Add a failure record without incrementing the assertion count.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    'Retain Err fields before any error-state reset or reporting call.
    Dim ErrorDescription   As String    'Case diagnostic captured before recording
    Dim ErrorNumber        As Long      'Case error number captured before recording
    Dim ErrorSource        As String    'Case error source captured before recording

'------------------------------------------------------------------------------
' CAPTURE ERROR
'------------------------------------------------------------------------------
    'Read the original diagnostic before On Error GoTo 0 can clear Err.
        ErrorNumber = Err.Number
        ErrorSource = Err.Source
        ErrorDescription = Err.Description
        On Error GoTo 0

'------------------------------------------------------------------------------
' RECORD DIAGNOSTIC
'------------------------------------------------------------------------------
    'Add the case-qualified failure without inventing a completed assertion
    'for an operation that raised before its check could run.
        RecordFailure _
            CaseName & ".unexpected", _
            "error=" & CStr(ErrorNumber) & _
                "; source=" & ErrorSource & _
                "; description=" & ErrorDescription

End Sub


Private Sub RecordFailure( _
    ByVal AssertionName As String, _
    ByVal Detail As String)
'
'==============================================================================
'                                RecordFailure
'------------------------------------------------------------------------------
' PURPOSE
'   Accumulate failure details for the final report.
'
' INPUTS
'   AssertionName identifies the failure; Detail supplies its diagnostic.
'
' STATE OWNERSHIP
'   Increment the failure count and append text to the owned report buffer.
'   Do not change case or assertion counts.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' RECORD FAILURE
'------------------------------------------------------------------------------
    'Retain every diagnostic in arrival order. A separator belongs only
    'between records so the final report has no leading empty entry.
        mFailureCount = mFailureCount + 1
        If Len(mFailureDetails) > 0 Then
            mFailureDetails = mFailureDetails & vbNewLine
        End If
        mFailureDetails = mFailureDetails & AssertionName & ": " & Detail

End Sub


'
'------------------------------------------------------------------------------
'
'                          CLEANUP AND HOST EVIDENCE
'
'------------------------------------------------------------------------------
'

Private Function CleanupRun( _
    ByRef CleanupDetail As String, _
    ByVal StateSnapshotValid As Boolean, _
    ByVal InitialCalculation As XlCalculation, _
    ByVal InitialDisplayAlerts As Boolean, _
    ByVal InitialEnableEvents As Boolean, _
    ByVal InitialScreenUpdating As Boolean) _
    As Boolean
'
'==============================================================================
'                                  CleanupRun
'------------------------------------------------------------------------------
' PURPOSE
'   Release the run flag and verify that observed Excel state is unchanged.
'
' INPUTS
'   StateSnapshotValid indicates whether all four initial properties were
'   read. The initial values belong to the current run only.
'
' RETURNS
'   Boolean success, and in the ByRef CleanupDetail the reason for the
'   final report.
'
' ERROR POLICY
'   Return False for an unavailable snapshot, changed state, or a cleanup
'   error. Report the reason; do not restore properties the harness never owned.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    'Track the combined comparison without taking ownership of host state.
    Dim ExcelStateUnchanged   As Boolean    'All four observed properties match the snapshot

'------------------------------------------------------------------------------
' VERIFY CLEANUP
'------------------------------------------------------------------------------
    'Release the re-entry flag first, even when no complete host snapshot
    'is available to verify the remaining cleanup contract.
        On Error GoTo CleanupFailed

        mRunActive = False
        If Not StateSnapshotValid Then
            CleanupDetail = "run flag cleared; Excel state snapshot unavailable"
            CleanupRun = False
            Exit Function
        End If

    'Compare only properties that were captured successfully. The harness
    'does not restore them because it never changed or owned them.
        ExcelStateUnchanged = _
            (Application.Calculation = InitialCalculation) And _
            (Application.DisplayAlerts = InitialDisplayAlerts) And _
            (Application.EnableEvents = InitialEnableEvents) And _
            (Application.ScreenUpdating = InitialScreenUpdating)

    'Return the combined result and retain a reason for either outcome.
        CleanupRun = ExcelStateUnchanged
        If ExcelStateUnchanged Then
            CleanupDetail = _
                "run flag cleared; verified Excel state unchanged " & _
                "(calculation/display alerts/events/screen updating)"
        Else
            CleanupDetail = "run flag cleared; Excel application state changed during test run"
        End If
        Exit Function

'------------------------------------------------------------------------------
' HANDLE CLEANUP ERROR
'------------------------------------------------------------------------------
CleanupFailed:
    'Describe a failed verification and return False. The runner retains
    'its own saved diagnostic for any later re-raise.
        CleanupDetail = _
            "cleanup error=" & CStr(Err.Number) & _
            "; description=" & Err.Description
        Err.Clear
        CleanupRun = False

End Function


Private Sub PrintEnvironment()
'
'==============================================================================
'                               PrintEnvironment
'------------------------------------------------------------------------------
' PURPOSE
'   Write the report preamble and current Excel environment.
'
' ERROR POLICY
'   Host-property read failures propagate to the runner error path.
'
' UPDATED
'   2026-10-06
'==============================================================================
'

'------------------------------------------------------------------------------
' REPORT ENVIRONMENT
'------------------------------------------------------------------------------
    'Write the preamble before the host fields so retained output can be
    'recognized and attributed to the environment that produced it.
        TEST_Suite.Say "SACCR TESTS"
        TEST_Suite.Say "MODE=" & IIf(mInjectFailure, "INJECTED_FAILURE", "NORMAL")
        TEST_Suite.Say "ENVIRONMENT=" & EnvironmentSummary()

End Sub


Private Function EnvironmentSummary() _
    As String
'
'==============================================================================
'                              EnvironmentSummary
'------------------------------------------------------------------------------
' PURPOSE
'   Describe the current Excel host and compiled VBA environment.
'
' RETURNS
'   Machine-readable host, version, operating system, Office bitness, and
'   VBA generation fields. Compile-time branches select the last two values.
'
' ERROR POLICY
'   Host-property read failures propagate to the runner error path.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    'Keep compile-time properties distinct from values read from Excel.
    Dim Bitness         As String    'Office bitness selected by conditional compilation
    Dim VbaGeneration   As String    'Compiled VBA generation reported as text

'------------------------------------------------------------------------------
' RESOLVE COMPILED ENVIRONMENT
'------------------------------------------------------------------------------
    'Describe the running Office/VBA build from compiler constants; do not
    'infer Office bitness from the Windows operating-system description.
#If Win64 Then
        Bitness = "64-bit"
#Else
        Bitness = "32-bit"
#End If

#If VBA7 Then
        VbaGeneration = "VBA7+"
#Else
        VbaGeneration = "legacy VBA"
#End If

'------------------------------------------------------------------------------
' BUILD ENVIRONMENT RECORD
'------------------------------------------------------------------------------
    'Combine live host properties with the compiled environment fields,
    'keeping the evidence parser keys and order stable.
        EnvironmentSummary = _
            "host=" & Application.Name & _
            "; version=" & Application.Version & _
            "; os=" & Application.OperatingSystem & _
            "; office=" & Bitness & _
            "; runtime=" & VbaGeneration

End Function


Private Sub PrintSummary( _
    ByVal CleanupPassed As Boolean, _
    ByVal CleanupDetail As String)
'
'==============================================================================
'                                 PrintSummary
'------------------------------------------------------------------------------
' PURPOSE
'   Report counts, cleanup, failure details, and the complete run verdict.
'
' INPUTS
'   CleanupPassed and CleanupDetail: the outcome of CleanupRun.
'
' STATE OWNERSHIP
'   Read the retained counters and completion flag without resetting them.
'   PASS requires zero failures, successful cleanup, and a complete suite.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    'Resolve one verdict before writing the machine-readable report.
    Dim Verdict   As String    'Final status derived from failures and cleanup

'------------------------------------------------------------------------------
' RESOLVE VERDICT
'------------------------------------------------------------------------------
    'A zero failure count is insufficient: cleanup must succeed and the
    'suite must have reached both required counts before reporting PASS.
        If mFailureCount = 0 And CleanupPassed And mSuiteCompleted Then
            Verdict = "PASS"
        Else
            Verdict = "FAIL"
        End If

'------------------------------------------------------------------------------
' REPORT
'------------------------------------------------------------------------------
    'Write the retained counts and cleanup first, then any diagnostics and
    'the final result record consumed by the evidence validator.
        TEST_Suite.Say "CASES=" & CStr(mCaseCount)
        TEST_Suite.Say "ASSERTIONS=" & CStr(mAssertionCount)
        TEST_Suite.Say "FAILURES=" & CStr(mFailureCount)
        TEST_Suite.Say "CLEANUP=" & IIf(CleanupPassed, "PASS", "FAIL") & _
            "; detail=" & CleanupDetail

    'Omit the diagnostic record on a clean run; retain all details on failure.
        If Len(mFailureDetails) > 0 Then
            TEST_Suite.Say "FAILURE_DETAILS=" & mFailureDetails
        End If

        TEST_Suite.Say _
            "RESULT=" & Verdict & _
            "; completeness=" & IIf(mSuiteCompleted, "COMPLETE", "INCOMPLETE") & _
            "; cases=" & CStr(mCaseCount) & _
            "; assertions=" & CStr(mAssertionCount) & _
            "; failures=" & CStr(mFailureCount) & _
            "; cleanup=" & IIf(CleanupPassed, "PASS", "FAIL")

End Sub


'
'------------------------------------------------------------------------------
'
'                              OWNED STATE RESET
'
'------------------------------------------------------------------------------
'

Private Sub ResetRun()
'
'==============================================================================
'                                   ResetRun
'------------------------------------------------------------------------------
' PURPOSE
'   Return all harness-owned state to its idle baseline.
'
' STATE OWNERSHIP
'   Clear counters, failure text, active-run state, and completion.
'   This helper never changes Excel application properties.
'
' UPDATED
'   2026-10-06
'==============================================================================
'

'------------------------------------------------------------------------------
' RESET
'------------------------------------------------------------------------------
    'Clear all harness-owned state together so no count, diagnostic, or
    'completion flag can leak from a previous execution into the next one.
        mCaseCount = 0
        mAssertionCount = 0
        mFailureCount = 0
        mFailureDetails = vbNullString
        mRunActive = False
        mSuiteCompleted = False

End Sub
