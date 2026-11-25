Attribute VB_Name = "TEST_Formulas"
'==============================================================================
' MODULE: TEST_Formulas
'------------------------------------------------------------------------------
' PURPOSE
'   Check the domain and error contract of the SACCR_* worksheet functions
'   when they are called directly, outside the engine and without the
'   worksheet's data validation (#42): every input outside its documented
'   domain gives the documented error value (#NUM!, #VALUE!, #REF! or the
'   error passed in) and never a runtime error or a misleading number;
'   an overflow gives #NUM!; valid inputs at and near the boundaries keep
'   their results; ranges, arrays, single values and one-element inputs
'   are handled alike; and the same results reach worksheet cells.
'
' PUBLIC SURFACE
'   RunFormulaTests is the entry point. Option Private Module keeps it out
'   of the external workbook automation API.
'
' DEPENDENCIES
'   SACCR_Formulas; M_Main.ProtectWorkbook, which protects the workbook
'   structure again when the last case could not; and TEST_Kit for the
'   suite lifecycle and the log lines (#536). Expected values are hand calculations from the formulas
'   in the functions' headers; they check the contract, not the regulation,
'   except in formula-test.independent, whose values the R package SACCR
'   computes (tools/rsaccr_formulas.R, #100).
'
' WORKSHEET SAFETY
'   The last case unprotects the workbook structure when it is protected,
'   adds a temporary worksheet, writes values and formulas on it, deletes
'   it again and protects the structure again, also after an unexpected
'   error. Nothing else in the workbook is written. Use a development
'   workbook.
'
' ERROR POLICY
'   A raised error is a failure of its case: the contract is an error
'   value, not a raised error. The remaining cases still run.
'
' USAGE
'   Import with the production modules, compile, then run
'   TEST_Formulas.RunFormulaTests from the Immediate window. It ends with
'   RESULT=PASS; cases=...; checks=...; failures=0; restore=PASS.
'
' REFERENCE
'   docs/PUBLIC_API.txt; CRE52.23, 52.34, 52.48, 52.52, 52.57, 52.61;
'   Delegated Regulation (EU) 2021/931 Art. 5.
'
' UPDATED
'   2026-11-23
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
        Private Const EXPECTED_CASES   As Long = 12                 'Cases in a complete run
    'Two checks of LongLong inputs run on 64-bit VBA only.
#If Win64 Then
        Private Const EXPECTED_CHECKS  As Long = 135                'Checks in a complete run
#Else
        Private Const EXPECTED_CHECKS  As Long = 133                'Checks in a complete run
#End If
        Private Const TOL              As Double = 0.000000001     'Absolute tolerance of numeric checks


'
'------------------------------------------------------------------------------
'
'                                 ENTRY POINT
'
'------------------------------------------------------------------------------
'

Public Sub RunFormulaTests()

'
'------------------------------------------------------------------------------
'                            RUN THE FORMULA TESTS
'------------------------------------------------------------------------------
' PURPOSE
'   Runs every case and prints one CASE line per case, each failed check,
'   and a final RESULT line
'
' WHY THIS EXISTS
'   The entry point of the suite for Test_RunAll and for a run by hand
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Starts the suite with its expected counts, runs the cases in order and
'   ends the suite, which prints the RESULT line
'
' ERROR POLICY
'   Each case contains its own errors as failures of the case
'
' DEPENDENCIES
'   TEST_Kit.BeginSuite, TEST_Kit.EndSuite and the Case procedures below
'
' NOTES
'   CaseEngineCalendar compares the engine's list of closing days with
'   SACCR_BusinessDays (#771)
'
' UPDATED
'   2026-11-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' START
'------------------------------------------------------------------------------
    'Stop when another suite is running
        If Not TEST_Kit.BeginSuite("TEST_Formulas", EXPECTED_CASES, EXPECTED_CHECKS) Then Exit Sub

'------------------------------------------------------------------------------
' RUN CASES
'------------------------------------------------------------------------------
    'Every case, in order
        CaseSupervisoryDuration
        CaseMaturityFactor
        CaseLambdaThreshold
        CaseIREffectiveNotional
        CaseMultiplier
        CaseReplacementCost
        CaseAggregationShapes
        CaseAggregationElements
        CaseBusinessDays
        CaseEngineCalendar
        CaseIndependentValues
        CaseWorksheetCells

'------------------------------------------------------------------------------
' REPORT
'------------------------------------------------------------------------------
    'Print the RESULT line
        TEST_Kit.EndSuite

End Sub


'
'------------------------------------------------------------------------------
'
'                                    CASES
'
'------------------------------------------------------------------------------
'

Private Sub CaseSupervisoryDuration()
'
'==============================================================================
' PURPOSE
'   SD: E before S and a negative floor are #NUM!; a negative S is floored
'   at 0; E equal to S gives 0, then the floor.
'==============================================================================
'
        On Error GoTo Unexpected
        TEST_Kit.BeginCase "supervisory-duration.domain"
        CheckError "E before S", xlErrNum, SACCR_SupervisoryDuration(2#, 1#)
        CheckError "E before S after the floor of S", xlErrNum, SACCR_SupervisoryDuration(-1#, -0.5)
        CheckError "negative floor", xlErrNum, SACCR_SupervisoryDuration(0#, 5#, -0.01)
        CheckNear "S = 0, E = 5", 4.4239843385719, SACCR_SupervisoryDuration(0#, 5#), 0.000000000001
        CheckNear "S = -1 floored at 0", 4.4239843385719, SACCR_SupervisoryDuration(-1#, 5#), 0.000000000001
        CheckNear "E = S, no floor", 0#, SACCR_SupervisoryDuration(1#, 1#), TOL
        CheckNear "E = S, floor 10/250", 0.04, SACCR_SupervisoryDuration(1#, 1#, 0.04), TOL
        Exit Sub
Unexpected:
        TEST_Kit.Fail TEST_Kit.CurrentCase, "unexpected error " & Err.Number & ": " & Err.Description
End Sub


Private Sub CaseMaturityFactor()
'
'==============================================================================
' PURPOSE
'   MF: a business-day year of 0 or less, an MPOR of 0 or less (margined)
'   and a negative floor (unmargined) are #NUM!; an argument the basis does
'   not use is not checked; a negative M with a floor of 0 gives 0; M below
'   the floor of 10 business days is floored and M above 1 capped; an
'   overflow of the floor in years is #NUM!.
'==============================================================================
'
        On Error GoTo Unexpected
        TEST_Kit.BeginCase "maturity-factor.domain"
        CheckError "BDPerYear 0, unmargined", xlErrNum, SACCR_MaturityFactor(0.5, False, 10#, 10#, 0#)
        CheckError "BDPerYear negative, margined", xlErrNum, SACCR_MaturityFactor(0#, True, 10#, 10#, -250#)
        CheckError "MPOR 0, margined", xlErrNum, SACCR_MaturityFactor(0#, True, 0#)
        CheckError "MPOR negative, margined", xlErrNum, SACCR_MaturityFactor(0#, True, -5#)
        CheckError "MinBD negative, unmargined", xlErrNum, SACCR_MaturityFactor(0.5, False, 10#, -1#)
        CheckNear "margined ignores a negative MinBD", 0.3, SACCR_MaturityFactor(0#, True, 10#, -1#), 0.000000000001
        CheckNear "unmargined ignores MPOR 0", Sqr(0.5), SACCR_MaturityFactor(0.5, False, 0#), 0.000000000001
        CheckNear "negative M, floor 0", 0#, SACCR_MaturityFactor(-1#, False, 10#, 0#), TOL
        CheckNear "M below the floor of 10 BD", Sqr(10# / 250#), SACCR_MaturityFactor(0.001), 0.000000000001
        CheckNear "M above 1 capped", 1#, SACCR_MaturityFactor(3#), TOL
        CheckError "overflow of MinBD / BDPerYear", xlErrNum, SACCR_MaturityFactor(0.5, False, 10#, 1E+300, 1E-300)
        Exit Sub
Unexpected:
        TEST_Kit.Fail TEST_Kit.CurrentCase, "unexpected error " & Err.Number & ": " & Err.Description
End Sub


Private Sub CaseLambdaThreshold()
'
'==============================================================================
' PURPOSE
'   CRR lambda: a threshold that is text or a Boolean is #VALUE!, a
'   negative one #NUM!, an error value is passed on; omitted or blank takes
'   the default; 0 is a valid threshold, also as a LongLong on 64-bit VBA.
'==============================================================================
'
        On Error GoTo Unexpected
        TEST_Kit.BeginCase "lambda-crr.threshold"
        CheckError "text threshold", xlErrValue, SACCR_LambdaCRR(-0.002, 0.001, True, "abc")
        CheckError "Boolean threshold", xlErrValue, SACCR_LambdaCRR(-0.002, 0.001, True, True)
        CheckError "negative threshold", xlErrNum, SACCR_LambdaCRR(-0.002, 0.001, True, -0.001)
        CheckError "error threshold passed on", xlErrNA, SACCR_LambdaCRR(-0.002, 0.001, True, CVErr(xlErrNA))
        CheckNear "IR, threshold omitted", 0.003, SACCR_LambdaCRR(-0.002, 0.001, True), TOL
        CheckNear "IR, threshold blank", 0.003, SACCR_LambdaCRR(-0.002, 0.001, True, Empty), TOL
        CheckNear "IR, threshold 0", 0.002, SACCR_LambdaCRR(-0.002, 0.001, True, 0#), TOL
        CheckNear "commodity, threshold omitted", 5.5, SACCR_LambdaCRR(-5#, -4#, False), TOL
#If Win64 Then
        CheckNear "IR, threshold a LongLong 0", 0.002, SACCR_LambdaCRR(-0.002, 0.001, True, CLngLng(0)), TOL
#End If
        Exit Sub
Unexpected:
        TEST_Kit.Fail TEST_Kit.CurrentCase, "unexpected error " & Err.Number & ": " & Err.Description
End Sub


Private Sub CaseIREffectiveNotional()
'
'==============================================================================
' PURPOSE
'   IR effective notional: a correlation outside 0 to 1 or three that do
'   not form a correlation matrix are #NUM!, not a zero effective
'   notional; a negative quadratic form within rounding is 0; the Basel
'   example 1 USD value is kept; all buckets 0 give 0 and an overflow is
'   #NUM!.
'==============================================================================
'
    Dim EffectiveNotional   As Variant    'A result checked for being a small number

        On Error GoTo Unexpected
        TEST_Kit.BeginCase "ir-effective-notional.correlations"
        CheckError "rho12 above 1", xlErrNum, SACCR_IREffectiveNotional(1#, 1#, 1#, 1.1, 0.7, 0.3)
        CheckError "rho13 below 0", xlErrNum, SACCR_IREffectiveNotional(1#, 1#, 1#, 0.7, 0.7, -0.1)
        CheckError "not a correlation matrix", xlErrNum, SACCR_IREffectiveNotional(1#, -1#, 1#, 1#, 1#, 0#)
        CheckNear "Basel example 1, USD", 59269.9634637104, _
                  SACCR_IREffectiveNotional(0#, -36253.8493844036, 78693.8680574733), 0.000001
        CheckNear "all buckets 0", 0#, SACCR_IREffectiveNotional(0#, 0#, 0#), TOL
        CheckError "overflow of D1 squared", xlErrNum, SACCR_IREffectiveNotional(1E+200, 0#, 0#)
    'With all correlations 1 the form is (D1 + D2 + D3)^2, here 0 in exact
    'arithmetic but slightly negative in double precision.
        EffectiveNotional = SACCR_IREffectiveNotional(100000.1, 2.2, -100002.3, 1#, 1#, 1#)
        TEST_Kit.Check Not IsError(EffectiveNotional), "rounding residual returned an error"
        If Not IsError(EffectiveNotional) Then
            TEST_Kit.Check CDbl(EffectiveNotional) >= 0# And CDbl(EffectiveNotional) < 0.01, _
                           "rounding residual gave " & CStr(EffectiveNotional)
        Else
            TEST_Kit.Check False, "rounding residual gave no number"
        End If
        Exit Sub
Unexpected:
        TEST_Kit.Fail TEST_Kit.CurrentCase, "unexpected error " & Err.Number & ": " & Err.Description
End Sub


Private Sub CaseMultiplier()
'
'==============================================================================
' PURPOSE
'   Multiplier: a negative add-on and a floor outside 0 to below 1 are
'   #NUM!; a zero add-on gives 1; a floor of 0 is valid; extreme V - C or a
'   tiny add-on give the floor or 1 without an overflow.
'==============================================================================
'
        On Error GoTo Unexpected
        TEST_Kit.BeginCase "multiplier.domain"
        CheckError "floor 1", xlErrNum, SACCR_Multiplier(-20#, 282#, 1#)
        CheckError "floor above 1", xlErrNum, SACCR_Multiplier(-20#, 282#, 1.5)
        CheckError "negative floor", xlErrNum, SACCR_Multiplier(-20#, 282#, -0.01)
        CheckError "negative add-on", xlErrNum, SACCR_Multiplier(-20#, -1#)
        CheckNear "zero add-on", 1#, SACCR_Multiplier(-20#, 0#), TOL
        CheckNear "floor 0", Exp(-0.5), SACCR_Multiplier(-1#, 1#, 0#), 0.000000000001
        CheckNear "V - C = -20, add-on 282", 0.965192681187195, SACCR_Multiplier(-20#, 282#), 0.000000000001
        CheckNear "very negative V - C gives the floor", 0.05, SACCR_Multiplier(-1E+300, 1#), TOL
        CheckNear "tiny add-on, V - C below 0: the floor, no overflow", 0.05, SACCR_Multiplier(-1E+300, 1E-300), TOL
        CheckNear "tiny add-on, V - C above 0: 1, no overflow", 1#, SACCR_Multiplier(1E+300, 1E-300), TOL
        Exit Sub
Unexpected:
        TEST_Kit.Fail TEST_Kit.CurrentCase, "unexpected error " & Err.Number & ": " & Err.Description
End Sub


Private Sub CaseReplacementCost()
'
'==============================================================================
' PURPOSE
'   RC: a negative threshold or MTA is #NUM! for a margined set and ignored
'   for an unmargined one, as the formula ignores them; NICA may have any
'   sign; an overflow of V - C is #NUM!.
'==============================================================================
'
        On Error GoTo Unexpected
        TEST_Kit.BeginCase "replacement-cost.domain"
        CheckError "margined, negative TH", xlErrNum, SACCR_ReplacementCost(5#, 10#, True, -1#, 10#, 2#)
        CheckError "margined, negative MTA", xlErrNum, SACCR_ReplacementCost(5#, 10#, True, 0#, -10#, 2#)
        CheckNear "unmargined ignores a negative TH", 70#, SACCR_ReplacementCost(100#, 30#, False, -1#, -1#), TOL
        CheckNear "margined, NICA negative", 12#, SACCR_ReplacementCost(5#, 10#, True, 0#, 10#, -2#), TOL
        CheckNear "margined, TH + MTA - NICA", 8#, SACCR_ReplacementCost(5#, 10#, True, 0#, 10#, 2#), TOL
        CheckError "overflow of V - C", xlErrNum, SACCR_ReplacementCost(1E+308, -1E+308)
        Exit Sub
Unexpected:
        TEST_Kit.Fail TEST_Kit.CurrentCase, "unexpected error " & Err.Number & ": " & Err.Description
End Sub

Private Sub CaseEngineCalendar()

'
'------------------------------------------------------------------------------
'               COMPARE THE ENGINE'S CALENDAR WITH THE FUNCTION
'------------------------------------------------------------------------------
' PURPOSE
'   Checks that the engine's list of closing days (#771) counts the business
'   days of every day of its window, and of intervals across it, as
'   SACCR_BusinessDays does
'
' WHY THIS EXISTS
'   On a business-day convention the engine counts on its own list, built
'   once from its own copy of the TARGET rule; a copy that drifted from the
'   function would change S, E, M and T without an error
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   - Every day from 2 January 1999 to 31 December 2100, one day at a time,
'     on TARGET and on TARGET with a table of closing days
'   - 2,000 intervals of up to 40 years, from a fixed sequence, on both
'   - Dates outside the window, and a reversed interval, are left uncounted
'
' ERROR POLICY
'   An unexpected error is a failure of the case
'
' DEPENDENCIES
'   CORE_Engine.WindowBusinessDays, SACCR_BusinessDays, TEST_Kit
'
' NOTES
'   The table holds a Monday, a Saturday, a TARGET day twice, a fraction of
'   a day, two days outside the window and is not sorted, as
'   tools/test_closing_list.py, which compares the same list with the
'   Python reference of the calendar
'
' UPDATED
'   2026-11-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "TEST_Formulas.CaseEngineCalendar"
    Const WINDOW_FIRST  As Long = 36161      '1 January 1999
    Const WINDOW_LAST   As Long = 73415      '31 December 2100

    Dim Table           As Variant    'Table of closing days
    Dim Holidays        As Variant    'Table of the calendar being compared
    Dim Calendar        As Long       '1 for TARGET, 2 for TARGET with the table
    Dim LastDay         As Long       'Last day of a one-day interval
    Dim FirstDay        As Long       'Day before the first day of an interval
    Dim Seed            As Long       'State of the fixed sequence
    Dim Pair            As Long       'Interval of the sequence
    Dim Wrong           As String     'First difference found
    Dim Differences     As Long       'Differences found

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Contain an unexpected error as a failure
        On Error GoTo Unexpected
        TEST_Kit.BeginCase "business-days.engine-calendar"
        Table = Array(CDbl(DateSerial(2030, 6, 3)), CDbl(DateSerial(2030, 6, 1)), _
            CDbl(DateSerial(2026, 12, 25)), CDbl(DateSerial(2026, 12, 25)), _
            CDbl(DateSerial(2024, 8, 15)) + 0.5, CDbl(DateSerial(1998, 6, 1)), _
            CDbl(DateSerial(2101, 3, 1)), CDbl(DateSerial(2099, 11, 2)), CDbl(DateSerial(2000, 1, 3)))

'------------------------------------------------------------------------------
' EVERY DAY OF THE WINDOW
'------------------------------------------------------------------------------
    'Each day alone, on TARGET and on TARGET with the table
        For Calendar = 1 To 2
            Holidays = Empty
            If Calendar = 2 Then Holidays = Table
            Differences = 0
            Wrong = vbNullString
            For LastDay = WINDOW_FIRST + 1 To WINDOW_LAST
                If Not SameCount(LastDay - 1, LastDay, Holidays) Then
                    Differences = Differences + 1
                    If Len(Wrong) = 0 Then Wrong = Format$(CDate(LastDay), "yyyy-mm-dd")
                End If
            Next LastDay
            TEST_Kit.Check Differences = 0, "every day, calendar " & Calendar & ": " & Differences & _
                " day(s) differ, the first " & Wrong
        Next Calendar

'------------------------------------------------------------------------------
' INTERVALS
'------------------------------------------------------------------------------
    'Intervals of up to 40 years from a fixed sequence, on both calendars
        Differences = 0
        Wrong = vbNullString
        Seed = 771
        For Pair = 1 To 2000
            Seed = (Seed * 1103 + 12345) Mod 1000003
            FirstDay = WINDOW_FIRST + (Seed Mod (WINDOW_LAST - WINDOW_FIRST + 1))
            Seed = (Seed * 1103 + 12345) Mod 1000003
            LastDay = FirstDay + (Seed Mod 14610)
            If LastDay > WINDOW_LAST Then LastDay = WINDOW_LAST
            For Calendar = 1 To 2
                Holidays = Empty
                If Calendar = 2 Then Holidays = Table
                If Not SameCount(FirstDay, LastDay, Holidays) Then
                    Differences = Differences + 1
                    If Len(Wrong) = 0 Then Wrong = FirstDay & " to " & LastDay & ", calendar " & Calendar
                End If
            Next Calendar
        Next Pair
        TEST_Kit.Check Differences = 0, "intervals: " & Differences & " differ, the first " & Wrong

'------------------------------------------------------------------------------
' OUTSIDE THE WINDOW
'------------------------------------------------------------------------------
    'Before 1999, after 2100 and a reversed interval are not counted
        TEST_Kit.Check IsEmpty(CORE_Engine.WindowBusinessDays(WINDOW_FIRST - 1, WINDOW_FIRST + 40, Empty)) And _
            IsEmpty(CORE_Engine.WindowBusinessDays(WINDOW_LAST - 40, WINDOW_LAST + 1, Empty)) And _
            IsEmpty(CORE_Engine.WindowBusinessDays(WINDOW_FIRST + 40, WINDOW_FIRST + 39, Empty)), _
            "a date outside the window or a reversed interval was counted"

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Leave before the error handler
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Unexpected:
    'Report the error as a failure of the case
        TEST_Kit.Fail "business-days.engine-calendar", PROC_NAME & ": unexpected error " & Err.Number & _
            ": " & Err.Description

End Sub

Private Function SameCount( _
    ByVal FirstDay As Long, _
    ByVal LastDay As Long, _
    ByVal Holidays As Variant) _
    As Boolean
'
'------------------------------------------------------------------------------
'                    COMPARE ONE INTERVAL OF BUSINESS DAYS
'------------------------------------------------------------------------------
' PURPOSE
'   Returns True when the engine's list and SACCR_BusinessDays give the
'   same business days in (FirstDay, LastDay] for a table of closing days
'
' UPDATED
'   2026-11-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Expected   As Variant    'SACCR_BusinessDays
    Dim Actual     As Variant    'The engine's count

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Both counts, numbers and equal
        If IsEmpty(Holidays) Then
            Expected = SACCR_BusinessDays(CDbl(FirstDay), CDbl(LastDay))
        Else
            Expected = SACCR_BusinessDays(CDbl(FirstDay), CDbl(LastDay), Holidays)
        End If
        Actual = CORE_Engine.WindowBusinessDays(CDbl(FirstDay), CDbl(LastDay), Holidays)
        If IsError(Expected) Or IsEmpty(Actual) Then Exit Function
        SameCount = (CDbl(Expected) = CDbl(Actual))

End Function

Private Sub CaseIndependentValues()
'
'==============================================================================
' PURPOSE
'   The FormulaTest checks the R package SACCR can compute (#100): each
'   worksheet function on the inputs of its FormulaTest row, against the
'   value the package gives, printed by tools/rsaccr_formulas.R (R 4.3.3,
'   SACCR 3.4, Trading 3.2; source IMPL-R-SACCR). Rows 24 and 25 check the
'   CRR lambda through the delta it shifts, as the package does: a bought
'   call with P = -0.2% and K = 0.1%, then with P = 2% and K = 3%, T = 2,
'   volatility 50%. The rows the package cannot compute, or computes on
'   another rule (9, 11, 15, 17, 18, 26, 27, 28), are listed in
'   docs/methodology/README.md.
'==============================================================================
'
        On Error GoTo Unexpected
        TEST_Kit.BeginCase "formula-test.independent"
        CheckNear "row 7, SD, S=0, E=10", 7.86938680574733, SACCR_SupervisoryDuration(0#, 10#), 0.000000000001
        CheckNear "row 8, SD, S=1, E=11", 7.48559228240455, SACCR_SupervisoryDuration(1#, 11#), 0.000000000001
        CheckNear "row 10, MF unmargined, M=0.75", 0.866025403784439, SACCR_MaturityFactor(0.75), 0.000000000001
        CheckNear "row 12, MF margined, MPOR 10 BD", 0.3, SACCR_MaturityFactor(0#, True, 10#), 0.000000000001
        CheckNear "row 13, delta, bought put", -0.269395217710533, _
                  SACCR_OptionDelta(0.06, 0.05, 1#, 0.5, False, True), 0.000000000001
        CheckNear "row 14, delta, sold call", -0.379011345827803, _
                  SACCR_OptionDelta(100#, 110#, 0.5, 0.32, True, False), 0.000000000001
        CheckNear "row 16, CDO delta, protection bought", 5.74096754439682, _
                  SACCR_CDODelta(0.03, 0.06, True), 0.000000000001
        CheckNear "row 19, multiplier, V - C = -20", 0.965192681187194, SACCR_Multiplier(-20#, 282#), 0.000000000001
        CheckNear "row 20, multiplier, V - C = 10", 1#, SACCR_Multiplier(10#, 282#), 0.000000000001
        CheckNear "row 21, RC margined", 8#, SACCR_ReplacementCost(5#, 10#, True, 0#, 10#, 2#), 0.000000000001
        CheckNear "row 22, normal CDF(1.96)", 0.97500210485178, SACCR_NormCDF(1.96), 0.000000000001
        CheckNear "row 23, normal CDF(-4)", 3.16712418331199E-05, SACCR_NormCDF(-4#), 0.000000000000001
        CheckNear "row 24, delta with the CRR lambda", 5.40312538068138E-02, _
                  SACCR_OptionDelta(-0.002, 0.001, 2#, 0.5, True, True, SACCR_LambdaCRR(-0.002, 0.001, True)), _
                  0.000000000001
        CheckNear "row 25, delta with no shift", 0.412989758121121, _
                  SACCR_OptionDelta(0.02, 0.03, 2#, 0.5, True, True, SACCR_LambdaCRR(0.02, 0.03, True)), _
                  0.000000000001
        Exit Sub
Unexpected:
        TEST_Kit.Fail TEST_Kit.CurrentCase, "unexpected error " & Err.Number & ": " & Err.Description
End Sub


Private Sub CaseAggregationShapes()
'
'==============================================================================
' PURPOSE
'   Factor aggregation: single values, one-element arrays and vectors in
'   any orientation agree; different sizes, different 2-D shapes and arrays
'   of more than two dimensions are #REF!.
'==============================================================================
'
    Dim RowAddOns            As Variant    '1 x 2 add-ons
    Dim ColumnCorrelations   As Variant    '2 x 1 correlations
    Dim GridAddOns           As Variant    '2 x 2 add-ons
    Dim GridCorrelations     As Variant    '2 x 2 correlations
    Dim CubeAddOns           As Variant    '1 x 1 x 1 array

        On Error GoTo Unexpected
        TEST_Kit.BeginCase "factor-aggregation.shapes"
        ReDim RowAddOns(1 To 1, 1 To 2)
        RowAddOns(1, 1) = 100#
        RowAddOns(1, 2) = 50#
        ReDim ColumnCorrelations(1 To 2, 1 To 1)
        ColumnCorrelations(1, 1) = 0.5
        ColumnCorrelations(2, 1) = 0.8
        ReDim GridAddOns(1 To 2, 1 To 2)
        GridAddOns(1, 1) = 100#
        GridAddOns(1, 2) = 50#
        GridAddOns(2, 1) = -30#
        GridAddOns(2, 2) = 20#
        ReDim GridCorrelations(1 To 2, 1 To 2)
        GridCorrelations(1, 1) = 0.5
        GridCorrelations(1, 2) = 0.8
        GridCorrelations(2, 1) = 0.4
        GridCorrelations(2, 2) = 0.6
        ReDim CubeAddOns(1 To 1, 1 To 1, 1 To 1)
        CubeAddOns(1, 1, 1) = 100#

        CheckNear "single values", 100#, SACCR_FactorAggregation(100#, 0.5), TOL
        CheckNear "single value and one-element array", 100#, SACCR_FactorAggregation(100#, Array(0.5)), TOL
        CheckNear "one-dimensional arrays", 128.452325786651, _
                  SACCR_FactorAggregation(Array(100#, 50#), Array(0.5, 0.8)), 0.000000001
        CheckNear "row with column", 128.452325786651, SACCR_FactorAggregation(RowAddOns, ColumnCorrelations), 0.000000001
        CheckNear "2 x 2 with 2 x 2", 132.332913517386, SACCR_FactorAggregation(GridAddOns, GridCorrelations), 0.000000001
        CheckError "different sizes", xlErrRef, SACCR_FactorAggregation(Array(100#, 50#), Array(0.5))
        CheckError "single value with two correlations", xlErrRef, SACCR_FactorAggregation(100#, Array(0.5, 0.8))
        CheckError "2 x 2 with 1 x 4", xlErrRef, SACCR_FactorAggregation(GridAddOns, Array(0.5, 0.8, 0.4, 0.6))
        CheckError "three dimensions", xlErrRef, SACCR_FactorAggregation(CubeAddOns, 0.5)
        Exit Sub
Unexpected:
        TEST_Kit.Fail TEST_Kit.CurrentCase, "unexpected error " & Err.Number & ": " & Err.Description
End Sub


Private Sub CaseAggregationElements()
'
'==============================================================================
' PURPOSE
'   Factor aggregation: text and Booleans are #VALUE!, an error element is
'   passed on, a correlation outside 0 to 1 is #NUM!, a blank add-on counts
'   as 0 and a blank correlation is allowed only beside it; a negative
'   add-on is valid, and so is a LongLong add-on on 64-bit VBA.
'==============================================================================
'
        On Error GoTo Unexpected
        TEST_Kit.BeginCase "factor-aggregation.elements"
        CheckError "text add-on", xlErrValue, SACCR_FactorAggregation(Array(100#, "50"), Array(0.5, 0.8))
        CheckError "Boolean add-on", xlErrValue, SACCR_FactorAggregation(Array(100#, True), Array(0.5, 0.8))
        CheckError "text correlation", xlErrValue, SACCR_FactorAggregation(Array(100#, 50#), Array(0.5, "x"))
        CheckError "blank correlation beside an add-on", xlErrValue, _
                   SACCR_FactorAggregation(Array(100#, 50#), Array(0.5, Empty))
        CheckError "error add-on passed on", xlErrDiv0, _
                   SACCR_FactorAggregation(Array(100#, CVErr(xlErrDiv0)), Array(0.5, 0.8))
        CheckError "error correlation passed on", xlErrNA, _
                   SACCR_FactorAggregation(Array(100#, 50#), Array(CVErr(xlErrNA), 0.8))
        CheckError "correlation above 1", xlErrNum, SACCR_FactorAggregation(Array(100#, 50#), Array(0.5, 1.2))
        CheckError "negative correlation", xlErrNum, SACCR_FactorAggregation(Array(100#, 50#), Array(-0.1, 0.8))
        CheckNear "blank add-on and blank correlation", 100#, _
                  SACCR_FactorAggregation(Array(100#, Empty), Array(0.5, Empty)), TOL
        CheckNear "blank add-on beside a correlation", 100#, _
                  SACCR_FactorAggregation(Array(100#, Empty), Array(0.5, 0.8)), TOL
        CheckNear "negative add-on", 132.332913517386, _
                  SACCR_FactorAggregation(Array(100#, 50#, -30#, 20#), Array(0.5, 0.8, 0.4, 0.6)), 0.000000001
#If Win64 Then
        CheckNear "LongLong add-on", 128.452325786651, _
                  SACCR_FactorAggregation(Array(CLngLng(100), 50#), Array(0.5, 0.8)), 0.000000001
#End If
        Exit Sub
Unexpected:
        TEST_Kit.Fail TEST_Kit.CurrentCase, "unexpected error " & Err.Number & ": " & Err.Description
End Sub


Private Sub CaseBusinessDays()
'
'==============================================================================
' PURPOSE
'   SACCR_BusinessDays (#721) against the Python reference (tools/
'   business_calendar.py): published closing days of 2008, 2026 and 2038,
'   the business days of whole years, an interval counted day by day with a
'   second Easter method (Gauss), a table of closing days, dates before
'   1999 and the errors of the table.
'==============================================================================
'

    Dim Yr          As Long       'A year of the loop
    Dim Expected      As Variant    'Business days of the years 2024 to 2028
    Dim Mismatch      As String     'The first year or pair whose count differs
    Dim Extra         As Variant    'A table of closing days
    Dim Seed          As Double     'Generator of the pairs of dates
    Dim FirstDate     As Long       'Start of a pair
    Dim LastDate      As Long       'End of a pair
    Dim Fast          As Variant    'Count of the function
    Dim Slow          As Long       'Count made day by day
    Dim i             As Long       'Pair

        On Error GoTo Unexpected
        TEST_Kit.BeginCase "business-days.target-calendar"
        CheckNear "the same day", 0#, SACCR_BusinessDays(DateSerial(2026, 6, 30), DateSerial(2026, 6, 30)), TOL
        CheckNear "the next day", 1#, SACCR_BusinessDays(DateSerial(2026, 6, 30), DateSerial(2026, 7, 1)), TOL
        CheckNear "Friday to Monday", 1#, SACCR_BusinessDays(DateSerial(2026, 7, 3), DateSerial(2026, 7, 6)), TOL
        CheckNear "Saturday to Sunday", 0#, SACCR_BusinessDays(DateSerial(2026, 7, 4), DateSerial(2026, 7, 5)), TOL
        CheckNear "Easter 2026: Good Friday and Easter Monday are closed", 1#, _
                  SACCR_BusinessDays(DateSerial(2026, 4, 2), DateSerial(2026, 4, 7)), TOL
        CheckNear "Easter 2026 without TARGET", 3#, SACCR_BusinessDays(DateSerial(2026, 4, 2), DateSerial(2026, 4, 7), , False), TOL
        CheckNear "the earliest Easter in range, 2008", 1#, SACCR_BusinessDays(DateSerial(2008, 3, 20), DateSerial(2008, 3, 25)), TOL
        CheckNear "the latest Easter in range, 2038", 1#, SACCR_BusinessDays(DateSerial(2038, 4, 22), DateSerial(2038, 4, 27)), TOL
        CheckNear "Christmas 2026", 1#, SACCR_BusinessDays(DateSerial(2026, 12, 24), DateSerial(2026, 12, 28)), TOL
        Expected = Array(256#, 255#, 256#, 258#, 255#)
        For Yr = 2024 To 2028
            CheckNear "business days of " & Yr, CDbl(Expected(Yr - 2024)), _
                      SACCR_BusinessDays(DateSerial(Yr - 1, 12, 31), DateSerial(Yr, 12, 31)), TOL
        Next Yr
        CheckNear "business days of 2000", 255#, SACCR_BusinessDays(DateSerial(1999, 12, 31), DateSerial(2000, 12, 31)), TOL
        CheckNear "business days of 2001, with 31 December closed", 254#, _
                  SACCR_BusinessDays(DateSerial(2000, 12, 31), DateSerial(2001, 12, 31)), TOL
        CheckNear "ten years", 2561#, SACCR_BusinessDays(DateSerial(2026, 6, 30), DateSerial(2036, 6, 30)), TOL
        CheckNear "two years and a half", 482#, SACCR_BusinessDays(DateSerial(2026, 6, 30), DateSerial(2028, 5, 15)), TOL
        CheckNear "reversed interval", -482#, SACCR_BusinessDays(DateSerial(2028, 5, 15), DateSerial(2026, 6, 30)), TOL
        CheckNear "a time of day is ignored", 1#, SACCR_BusinessDays(46203.75, 46204.25), TOL

    'Every year from 2000 to 2099 against a count made day by day, with a
    'second Easter method.
        For Yr = 2000 To 2099
            Fast = SACCR_BusinessDays(DateSerial(Yr - 1, 12, 31), DateSerial(Yr, 12, 31))
            Slow = CountDayByDay(CLng(DateSerial(Yr - 1, 12, 31)), CLng(DateSerial(Yr, 12, 31)), Empty)
            If IsError(Fast) Then
                Mismatch = Mismatch & " " & Yr & ": error"
            ElseIf CLng(Fast) <> Slow Then
                Mismatch = Mismatch & " " & Yr & ": " & CLng(Fast) & " against " & Slow
            End If
        Next Yr
        TEST_Kit.Check Len(Mismatch) = 0, "years 2000 to 2099 against a count day by day:" & Mismatch

    'Pairs of dates from 2000 to 2032 with a table, against the same count.
        Extra = Array(CDbl(DateSerial(2026, 8, 15)), CDbl(DateSerial(2026, 12, 8)), CDbl(DateSerial(2027, 1, 6)), _
                      CDbl(DateSerial(2026, 8, 14)))
        Mismatch = ""
        Seed = 721#
        For i = 1 To 300
            Seed = NextSeed(Seed)
            FirstDate = CLng(DateSerial(2000, 1, 1)) + CLng(Seed Mod 12000#)
            Seed = NextSeed(Seed)
            LastDate = FirstDate + CLng(Seed Mod 1500#)
            Fast = SACCR_BusinessDays(FirstDate, LastDate, Extra)
            Slow = CountDayByDay(FirstDate, LastDate, Extra)
            If IsError(Fast) Then
                Mismatch = Mismatch & " " & FirstDate & "-" & LastDate & ": error"
            ElseIf CLng(Fast) <> Slow Then
                Mismatch = Mismatch & " " & FirstDate & "-" & LastDate & ": " & CLng(Fast) & " against " & Slow
            End If
        Next i
        TEST_Kit.Check Len(Mismatch) = 0, "300 pairs of dates with a table against a count day by day:" & Mismatch

    'The table: a Friday counted, a Saturday and a duplicate not, a table
    'alone, and no range of validity without the rule.
        CheckNear "TARGET and a table, Friday 14 and Saturday 15 August 2026", 1#, _
                  SACCR_BusinessDays(DateSerial(2026, 8, 13), DateSerial(2026, 8, 17), _
                                     Array(CDbl(DateSerial(2026, 8, 14)), CDbl(DateSerial(2026, 8, 15)), CDbl(DateSerial(2026, 8, 14)))), TOL
        CheckNear "TARGET alone, the same days", 2#, SACCR_BusinessDays(DateSerial(2026, 8, 13), DateSerial(2026, 8, 17)), TOL
        CheckNear "a table alone, Tuesday 8 December 2026", 1#, _
                  SACCR_BusinessDays(DateSerial(2026, 12, 7), DateSerial(2026, 12, 9), CDbl(DateSerial(2026, 12, 8)), False), TOL
        CheckNear "a table alone has no TARGET days", 2#, SACCR_BusinessDays(DateSerial(2026, 12, 24), DateSerial(2026, 12, 28), , False), TOL
        CheckNear "no range of validity without the rule", 5#, SACCR_BusinessDays(DateSerial(1990, 1, 1), DateSerial(1990, 1, 8), , False), TOL
        CheckNear "a VBA Date as the table", 1#, _
                  SACCR_BusinessDays(DateSerial(2026, 12, 7), DateSerial(2026, 12, 9), DateSerial(2026, 12, 8), False), TOL
        CheckNear "VBA Dates in an array", 1#, _
                  SACCR_BusinessDays(DateSerial(2026, 12, 7), DateSerial(2026, 12, 9), Array(DateSerial(2026, 12, 8), DateSerial(2026, 12, 8)), False), TOL
        CheckNear "a blank entry of the table is skipped", 1#, _
                  SACCR_BusinessDays(DateSerial(2026, 12, 7), DateSerial(2026, 12, 9), Array(Empty, CDbl(DateSerial(2026, 12, 8))), False), TOL

    'Domain.
        CheckError "an interval before 1999 under the rule", xlErrNum, SACCR_BusinessDays(DateSerial(1998, 12, 1), DateSerial(1999, 2, 1))
        CheckError "a Saturday of 1998 under the rule", xlErrNum, SACCR_BusinessDays(DateSerial(1998, 12, 26), DateSerial(1998, 12, 27))
        CheckError "a date before 1 March 1900", xlErrNum, SACCR_BusinessDays(1#, 100#, , False)
        CheckError "text in the table", xlErrValue, SACCR_BusinessDays(DateSerial(2026, 1, 1), DateSerial(2026, 2, 1), Array("not a date"))
        CheckError "an error value in the table", xlErrNA, SACCR_BusinessDays(DateSerial(2026, 1, 1), DateSerial(2026, 2, 1), Array(CVErr(xlErrNA)))
        Exit Sub
Unexpected:
        TEST_Kit.Fail TEST_Kit.CurrentCase, "unexpected error " & Err.Number & ": " & Err.Description
End Sub


Private Function NextSeed( _
    ByVal Seed As Double) _
    As Double
    'A linear congruential generator in Double arithmetic, so that no Long overflows.
        Dim Product   As Double    'Seed times the multiplier plus the increment

        Product = Seed * 1103515245# + 12345#
        NextSeed = Product - Int(Product / 2147483648#) * 2147483648#

End Function


Private Function CountDayByDay( _
    ByVal FirstDate As Long, _
    ByVal LastDate As Long, _
    ByVal Extra As Variant) _
    As Long
    'Business days in (FirstDate, LastDate] by testing each day, with the
    'closing days of TARGET from their own rule and the table.
        Dim Probe     As Long    'Probe tested
        Dim Total   As Long    'Business days so far
        Dim k       As Long    'Entry of the table
        Dim Listed  As Boolean 'The day is in the table

        For Probe = FirstDate + 1 To LastDate
            If Weekday(Probe, vbMonday) <= 5 And Not TargetClosed(Probe) Then
                Listed = False
                If IsArray(Extra) Then
                    For k = LBound(Extra) To UBound(Extra)
                        If CLng(Extra(k)) = Probe Then Listed = True
                    Next k
                End If
                If Not Listed Then Total = Total + 1
            End If
        Next Probe
        CountDayByDay = Total

End Function


Private Function TargetClosed( _
    ByVal Probe As Long) _
    As Boolean
    'A closing day of TARGET from 2000 on (and 31 December 2001), with Easter by Gauss's rule
    'and its two exceptions, valid for the years 1900 to 2099.
        Dim Y   As Long
        Dim a   As Long
        Dim b   As Long
        Dim c   As Long
        Dim k   As Long
        Dim p   As Long
        Dim q   As Long
        Dim m   As Long
        Dim n   As Long
        Dim d   As Long
        Dim e   As Long
        Dim Easter As Long    'Serial number of Easter Sunday

        Y = Year(Probe)
        If Month(Probe) = 1 And Day(Probe) = 1 Then TargetClosed = True
        If Month(Probe) = 5 And Day(Probe) = 1 Then TargetClosed = True
        If Month(Probe) = 12 And (Day(Probe) = 25 Or Day(Probe) = 26) Then TargetClosed = True
        If Y = 2001 And Month(Probe) = 12 And Day(Probe) = 31 Then TargetClosed = True
        a = Y Mod 19
        b = Y Mod 4
        c = Y Mod 7
        k = Y \ 100
        p = (13 + 8 * k) \ 25
        q = k \ 4
        m = (15 - p + k - q) Mod 30
        n = (4 + k - q) Mod 7
        d = (19 * a + m) Mod 30
        e = (2 * b + 4 * c + 6 * d + n) Mod 7
        If 22 + d + e > 31 Then
            If d + e - 9 = 26 Then
                Easter = CLng(DateSerial(Y, 4, 19))
            ElseIf d + e - 9 = 25 And d = 28 And e = 6 And a > 10 Then
                Easter = CLng(DateSerial(Y, 4, 18))
            Else
                Easter = CLng(DateSerial(Y, 4, d + e - 9))
            End If
        Else
            Easter = CLng(DateSerial(Y, 3, 22 + d + e))
        End If
        If Probe = Easter - 2 Or Probe = Easter + 1 Then TargetClosed = True

End Function


Private Sub CaseWorksheetCells()
'
'==============================================================================
' PURPOSE
'   The same contract in worksheet cells: formulas on a temporary sheet
'   show the error values, ranges with a blank and an error cell are read
'   as arrays, and a cell reference to a threshold is read as its value.
'==============================================================================
'
    Dim TempSheet      As Worksheet    'Temporary sheet; Nothing once deleted
    Dim CallerAlerts   As Boolean      'Caller's DisplayAlerts
    Dim CallerEvents   As Boolean      'Caller's EnableEvents
    Dim CallerActive   As Object       'Sheet active before the case
    Dim Structure      As Boolean      'The workbook structure was protected (#421)

        On Error GoTo Unexpected
        TEST_Kit.BeginCase "worksheet.cells"
        CallerAlerts = Application.DisplayAlerts
        CallerEvents = Application.EnableEvents
        Set CallerActive = ThisWorkbook.ActiveSheet
        Application.EnableEvents = False
        Structure = ThisWorkbook.ProtectStructure
        If Structure Then ThisWorkbook.Unprotect
        Set TempSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))

    'A1:A2 add-ons and B1:B2 correlations, with row 3 blank; C1:C2 a range
    'with an error; D1:D3 thresholds, D3 blank; E1:E10 the formulas.
        TempSheet.Range("A1").Value = 100#
        TempSheet.Range("A2").Value = 50#
        TempSheet.Range("B1").Value = 0.5
        TempSheet.Range("B2").Value = 0.8
        TempSheet.Range("C1").Value = 100#
        TempSheet.Range("C2").Formula = "=NA()"
        TempSheet.Range("D1").Value = "abc"
        TempSheet.Range("D2").Value = -0.001
        TempSheet.Range("E1").Formula = "=SACCR_FactorAggregation(A1:A3,B1:B3)"
        TempSheet.Range("E2").Formula = "=SACCR_FactorAggregation(A1:A2,B1:B3)"
        TempSheet.Range("E3").Formula = "=SACCR_FactorAggregation(C1:C2,B1:B2)"
        TempSheet.Range("E4").Formula = "=SACCR_FactorAggregation(A1,B1)"
        TempSheet.Range("E5").Formula = "=SACCR_Multiplier(-20,282,1)"
        TempSheet.Range("E6").Formula = "=SACCR_MaturityFactor(0.5,FALSE,10,10,0)"
        TempSheet.Range("E7").Formula = "=SACCR_LambdaCRR(-0.002,0.001,TRUE,D1)"
        TempSheet.Range("E8").Formula = "=SACCR_LambdaCRR(-0.002,0.001,TRUE,D2)"
        TempSheet.Range("E9").Formula = "=SACCR_LambdaCRR(-0.002,0.001,TRUE,D3)"
        TempSheet.Range("E10").Formula = "=SACCR_IREffectiveNotional(1,1,1,1,1,0)"
        TempSheet.Calculate

        CheckNear "cells: blank row in both ranges", 128.452325786651, TempSheet.Range("E1").Value, 0.000000001
        CheckError "cells: ranges of different sizes", xlErrRef, TempSheet.Range("E2").Value
        CheckError "cells: error cell passed on", xlErrNA, TempSheet.Range("E3").Value
        CheckNear "cells: single cells", 100#, TempSheet.Range("E4").Value, TOL
        CheckError "cells: multiplier floor 1", xlErrNum, TempSheet.Range("E5").Value
        CheckError "cells: maturity factor BDPerYear 0", xlErrNum, TempSheet.Range("E6").Value
        CheckError "cells: text threshold", xlErrValue, TempSheet.Range("E7").Value
        CheckError "cells: negative threshold", xlErrNum, TempSheet.Range("E8").Value
        CheckNear "cells: blank threshold takes the default", 0.003, TempSheet.Range("E9").Value, TOL
        CheckError "cells: not a correlation matrix", xlErrNum, TempSheet.Range("E10").Value
        RemoveTemporarySheet TempSheet, CallerAlerts, CallerEvents, CallerActive, Structure
        Exit Sub

Unexpected:
        TEST_Kit.Fail TEST_Kit.CurrentCase, "unexpected error " & Err.Number & ": " & Err.Description
        RemoveTemporarySheet TempSheet, CallerAlerts, CallerEvents, CallerActive, Structure

End Sub


Private Sub RemoveTemporarySheet( _
    ByVal TempSheet As Worksheet, _
    ByVal CallerAlerts As Boolean, _
    ByVal CallerEvents As Boolean, _
    ByVal CallerActive As Object, _
    ByVal Structure As Boolean)
'
'==============================================================================
' PURPOSE
'   Delete the temporary sheet of CaseWorksheetCells without a prompt and
'   give back the caller's settings and active sheet, and the protection
'   of the workbook structure when Structure says it had one (#421).
'
' ERROR POLICY
'   Contained: a failure is reported as a test failure of the current case.
'==============================================================================
'
        On Error GoTo Failed
        If Not TempSheet Is Nothing Then
            Application.DisplayAlerts = False
            TempSheet.Delete
        End If
        Application.DisplayAlerts = CallerAlerts
        Application.EnableEvents = CallerEvents
        If Structure Then ThisWorkbook.Protect Structure:=True, Windows:=False
        If Not CallerActive Is Nothing Then
            CallerActive.Activate
        End If
        Exit Sub

Failed:
        TEST_Kit.Fail TEST_Kit.CurrentCase, "could not remove the temporary sheet: " & Err.Description
        Application.DisplayAlerts = CallerAlerts
        Application.EnableEvents = CallerEvents
        If Structure Then M_Main.ProtectWorkbook

End Sub


'
'------------------------------------------------------------------------------
'
'                                   HELPERS
'
'------------------------------------------------------------------------------
'

Private Sub CheckError( _
    ByVal CheckName As String, _
    ByVal ExpectedCode As Long, _
    ByVal Actual As Variant)
'
'==============================================================================
' PURPOSE
'   Check that Actual is the worksheet error value ExpectedCode.
'==============================================================================
'
    'Compare error with error: comparing a number with an error value raises.
        If IsError(Actual) Then
            TEST_Kit.Check Actual = CVErr(ExpectedCode), CheckName & ": expected error " & ExpectedCode & _
                           ", found " & CStr(Actual)
        Else
            TEST_Kit.Check False, CheckName & ": expected error " & ExpectedCode & ", found " & DescribeResult(Actual)
        End If

End Sub


Private Sub CheckNear( _
    ByVal CheckName As String, _
    ByVal Expected As Double, _
    ByVal Actual As Variant, _
    ByVal Tolerance As Double)
'
'==============================================================================
' PURPOSE
'   Check that Actual is a number within Tolerance of Expected.
'==============================================================================
'
        If IsError(Actual) Or Not IsNumeric(Actual) Or VarType(Actual) = vbBoolean Or IsEmpty(Actual) Then
            TEST_Kit.Check False, CheckName & ": expected " & CStr(Expected) & ", found " & DescribeResult(Actual)
        Else
            TEST_Kit.Check Abs(CDbl(Actual) - Expected) <= Tolerance, CheckName & ": expected " & CStr(Expected) & _
                           ", found " & CStr(Actual)
        End If

End Sub


Private Function DescribeResult( _
    ByVal Outcome As Variant) _
    As String
'
'==============================================================================
' PURPOSE
'   Describe a result for a failure message.
'==============================================================================
'
        If IsError(Outcome) Then
            DescribeResult = CStr(Outcome)
        ElseIf IsEmpty(Outcome) Then
            DescribeResult = "(empty)"
        Else
            DescribeResult = CStr(Outcome)
        End If

End Function
