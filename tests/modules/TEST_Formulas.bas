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
'   their results; ranges,
'   arrays, single values and one-element inputs are handled alike; and
'   the same results reach worksheet cells.
'
' PUBLIC SURFACE
'   RunFormulaTests is the entry point. Option Private Module keeps it out
'   of the external workbook automation API.
'
' DEPENDENCIES
'   SACCR_Formulas, and TEST_Suite.Say for the log lines. Expected values
'   are hand calculations from the formulas in the functions' headers; they
'   check the contract, not the regulation.
'
' WORKSHEET SAFETY
'   The last case adds a temporary worksheet, writes values and formulas
'   on it and deletes it again, also after an unexpected error. Nothing
'   else in the workbook is written. Use a development workbook.
'
' ERROR POLICY
'   A raised error is a failure of its case: the contract is an error
'   value, not a raised error. The remaining cases still run.
'
' USAGE
'   Import with the production modules, compile, then run
'   TEST_Formulas.RunFormulaTests from the Immediate window. It ends with
'   RESULT=PASS; cases=...; checks=...; failures=0.
'
' REFERENCE
'   docs/PUBLIC_API.txt; CRE52.23, 52.34, 52.48, 52.52, 52.57, 52.61;
'   Delegated Regulation (EU) 2021/931 Art. 5.
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
        Private Const EXPECTED_CASES   As Long = 9                 'Cases in a complete run
    'Two checks of LongLong inputs run on 64-bit VBA only.
#If Win64 Then
        Private Const EXPECTED_CHECKS  As Long = 82                'Checks in a complete run
#Else
        Private Const EXPECTED_CHECKS  As Long = 80                'Checks in a complete run
#End If
        Private Const TOL              As Double = 0.000000001     'Absolute tolerance of numeric checks

'------------------------------------------------------------------------------
' MODULE STATE
'------------------------------------------------------------------------------
        Private mCaseCount      As Long      'Cases started
        Private mCheckCount     As Long      'Checks evaluated
        Private mFailureCount   As Long      'Failed checks and unexpected errors
        Private mCurrentCase    As String    'Case being run, for messages


'
'------------------------------------------------------------------------------
'
'                                 ENTRY POINT
'
'------------------------------------------------------------------------------
'

Public Sub RunFormulaTests()
'
'==============================================================================
'                               RunFormulaTests
'------------------------------------------------------------------------------
' PURPOSE
'   Run every case and print one CASE line per case, each failed check,
'   and a final RESULT line.
'
' UPDATED
'   2026-10-12
'==============================================================================
'
        mCaseCount = 0
        mCheckCount = 0
        mFailureCount = 0
        TEST_Suite.Say "SACCR FORMULA CONTRACT TESTS"

        CaseSupervisoryDuration
        CaseMaturityFactor
        CaseLambdaThreshold
        CaseIREffectiveNotional
        CaseMultiplier
        CaseReplacementCost
        CaseAggregationShapes
        CaseAggregationElements
        CaseWorksheetCells

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

Private Sub CaseSupervisoryDuration()
'
'==============================================================================
' PURPOSE
'   SD: E before S and a negative floor are #NUM!; a negative S is floored
'   at 0; E equal to S gives 0, then the floor.
'==============================================================================
'
        On Error GoTo Unexpected
        BeginCase "supervisory-duration.domain"
        CheckError "E before S", xlErrNum, SACCR_SupervisoryDuration(2#, 1#)
        CheckError "E before S after the floor of S", xlErrNum, SACCR_SupervisoryDuration(-1#, -0.5)
        CheckError "negative floor", xlErrNum, SACCR_SupervisoryDuration(0#, 5#, -0.01)
        CheckNear "S = 0, E = 5", 4.4239843385719, SACCR_SupervisoryDuration(0#, 5#), 0.000000000001
        CheckNear "S = -1 floored at 0", 4.4239843385719, SACCR_SupervisoryDuration(-1#, 5#), 0.000000000001
        CheckNear "E = S, no floor", 0#, SACCR_SupervisoryDuration(1#, 1#), TOL
        CheckNear "E = S, floor 10/250", 0.04, SACCR_SupervisoryDuration(1#, 1#, 0.04), TOL
        Exit Sub
Unexpected:
        Fail mCurrentCase, "unexpected error " & Err.Number & ": " & Err.Description
End Sub


Private Sub CaseMaturityFactor()
'
'==============================================================================
' PURPOSE
'   MF: a business-day year of 0 or less, an MPOR of 0 or less (margined)
'   and a negative floor (unmargined) are #NUM!; an argument the basis does
'   not use is not checked; a negative M with a floor of 0 gives 0.
'==============================================================================
'
        On Error GoTo Unexpected
        BeginCase "maturity-factor.domain"
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
        Fail mCurrentCase, "unexpected error " & Err.Number & ": " & Err.Description
End Sub


Private Sub CaseLambdaThreshold()
'
'==============================================================================
' PURPOSE
'   CRR lambda: a threshold that is text is #VALUE!, a negative one #NUM!,
'   an error value is passed on; omitted or blank takes the default; 0 is
'   a valid threshold.
'==============================================================================
'
        On Error GoTo Unexpected
        BeginCase "lambda-crr.threshold"
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
        Fail mCurrentCase, "unexpected error " & Err.Number & ": " & Err.Description
End Sub


Private Sub CaseIREffectiveNotional()
'
'==============================================================================
' PURPOSE
'   IR effective notional: a correlation outside 0 to 1 or three that do
'   not form a correlation matrix are #NUM!, not a zero effective
'   notional; a negative quadratic form within rounding is 0; the Basel
'   example 1 USD value is kept.
'==============================================================================
'
    Dim result   As Variant    'A result checked for being a small number

        On Error GoTo Unexpected
        BeginCase "ir-effective-notional.correlations"
        CheckError "rho12 above 1", xlErrNum, SACCR_IREffectiveNotional(1#, 1#, 1#, 1.1, 0.7, 0.3)
        CheckError "rho13 below 0", xlErrNum, SACCR_IREffectiveNotional(1#, 1#, 1#, 0.7, 0.7, -0.1)
        CheckError "not a correlation matrix", xlErrNum, SACCR_IREffectiveNotional(1#, -1#, 1#, 1#, 1#, 0#)
        CheckNear "Basel example 1, USD", 59269.9634637104, _
                  SACCR_IREffectiveNotional(0#, -36253.8493844036, 78693.8680574733), 0.000001
        CheckNear "all buckets 0", 0#, SACCR_IREffectiveNotional(0#, 0#, 0#), TOL
        CheckError "overflow of D1 squared", xlErrNum, SACCR_IREffectiveNotional(1E+200, 0#, 0#)
    'With all correlations 1 the form is (D1 + D2 + D3)^2, here 0 in exact
    'arithmetic but slightly negative in double precision.
        result = SACCR_IREffectiveNotional(100000.1, 2.2, -100002.3, 1#, 1#, 1#)
        Check Not IsError(result), "rounding residual returned an error"
        If Not IsError(result) Then
            Check CDbl(result) >= 0# And CDbl(result) < 0.01, "rounding residual gave " & CStr(result)
        Else
            Check False, "rounding residual gave no number"
        End If
        Exit Sub
Unexpected:
        Fail mCurrentCase, "unexpected error " & Err.Number & ": " & Err.Description
End Sub


Private Sub CaseMultiplier()
'
'==============================================================================
' PURPOSE
'   Multiplier: a negative add-on and a floor outside 0 to below 1 are
'   #NUM!; a zero add-on gives 1; a floor of 0 is valid.
'==============================================================================
'
        On Error GoTo Unexpected
        BeginCase "multiplier.domain"
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
        Fail mCurrentCase, "unexpected error " & Err.Number & ": " & Err.Description
End Sub


Private Sub CaseReplacementCost()
'
'==============================================================================
' PURPOSE
'   RC: a negative threshold or MTA is #NUM! for a margined set and ignored
'   for an unmargined one, as the formula ignores them; NICA may have any
'   sign.
'==============================================================================
'
        On Error GoTo Unexpected
        BeginCase "replacement-cost.domain"
        CheckError "margined, negative TH", xlErrNum, SACCR_ReplacementCost(5#, 10#, True, -1#, 10#, 2#)
        CheckError "margined, negative MTA", xlErrNum, SACCR_ReplacementCost(5#, 10#, True, 0#, -10#, 2#)
        CheckNear "unmargined ignores a negative TH", 70#, SACCR_ReplacementCost(100#, 30#, False, -1#, -1#), TOL
        CheckNear "margined, NICA negative", 12#, SACCR_ReplacementCost(5#, 10#, True, 0#, 10#, -2#), TOL
        CheckNear "margined, TH + MTA - NICA", 8#, SACCR_ReplacementCost(5#, 10#, True, 0#, 10#, 2#), TOL
        CheckError "overflow of V - C", xlErrNum, SACCR_ReplacementCost(1E+308, -1E+308)
        Exit Sub
Unexpected:
        Fail mCurrentCase, "unexpected error " & Err.Number & ": " & Err.Description
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
    Dim row2    As Variant    '1 x 2 add-ons
    Dim col2    As Variant    '2 x 1 correlations
    Dim grid    As Variant    '2 x 2 add-ons
    Dim gridR   As Variant    '2 x 2 correlations
    Dim cube    As Variant    '1 x 1 x 1 array

        On Error GoTo Unexpected
        BeginCase "factor-aggregation.shapes"
        ReDim row2(1 To 1, 1 To 2)
        row2(1, 1) = 100#
        row2(1, 2) = 50#
        ReDim col2(1 To 2, 1 To 1)
        col2(1, 1) = 0.5
        col2(2, 1) = 0.8
        ReDim grid(1 To 2, 1 To 2)
        grid(1, 1) = 100#
        grid(1, 2) = 50#
        grid(2, 1) = -30#
        grid(2, 2) = 20#
        ReDim gridR(1 To 2, 1 To 2)
        gridR(1, 1) = 0.5
        gridR(1, 2) = 0.8
        gridR(2, 1) = 0.4
        gridR(2, 2) = 0.6
        ReDim cube(1 To 1, 1 To 1, 1 To 1)
        cube(1, 1, 1) = 100#

        CheckNear "single values", 100#, SACCR_FactorAggregation(100#, 0.5), TOL
        CheckNear "single value and one-element array", 100#, SACCR_FactorAggregation(100#, Array(0.5)), TOL
        CheckNear "one-dimensional arrays", 128.452325786651, _
                  SACCR_FactorAggregation(Array(100#, 50#), Array(0.5, 0.8)), 0.000000001
        CheckNear "row with column", 128.452325786651, SACCR_FactorAggregation(row2, col2), 0.000000001
        CheckNear "2 x 2 with 2 x 2", 132.332913517386, SACCR_FactorAggregation(grid, gridR), 0.000000001
        CheckError "different sizes", xlErrRef, SACCR_FactorAggregation(Array(100#, 50#), Array(0.5))
        CheckError "single value with two correlations", xlErrRef, SACCR_FactorAggregation(100#, Array(0.5, 0.8))
        CheckError "2 x 2 with 1 x 4", xlErrRef, SACCR_FactorAggregation(grid, Array(0.5, 0.8, 0.4, 0.6))
        CheckError "three dimensions", xlErrRef, SACCR_FactorAggregation(cube, 0.5)
        Exit Sub
Unexpected:
        Fail mCurrentCase, "unexpected error " & Err.Number & ": " & Err.Description
End Sub


Private Sub CaseAggregationElements()
'
'==============================================================================
' PURPOSE
'   Factor aggregation: text and Booleans are #VALUE!, an error element is
'   passed on, a correlation outside 0 to 1 is #NUM!, a blank add-on counts
'   as 0 and a blank correlation is allowed only beside it.
'==============================================================================
'
        On Error GoTo Unexpected
        BeginCase "factor-aggregation.elements"
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
        Fail mCurrentCase, "unexpected error " & Err.Number & ": " & Err.Description
End Sub


Private Sub CaseWorksheetCells()
'
'==============================================================================
' PURPOSE
'   The same contract in worksheet cells: formulas on a temporary sheet
'   show the error values, ranges with a blank and an error cell are read
'   as arrays, and a cell reference to a threshold is read as its value.
'==============================================================================
'
    Dim ws            As Worksheet    'Temporary sheet; Nothing once deleted
    Dim callerAlerts  As Boolean      'Caller's DisplayAlerts
    Dim callerEvents  As Boolean      'Caller's EnableEvents
    Dim callerActive  As Object       'Sheet active before the case

        On Error GoTo Unexpected
        BeginCase "worksheet.cells"
        callerAlerts = Application.DisplayAlerts
        callerEvents = Application.EnableEvents
        Set callerActive = ThisWorkbook.ActiveSheet
        Application.EnableEvents = False
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))

    'A1:A3 add-ons, B1:B3 correlations, C1:C2 a range with an error.
        ws.Range("A1").Value = 100#
        ws.Range("A2").Value = 50#
        ws.Range("B1").Value = 0.5
        ws.Range("B2").Value = 0.8
        ws.Range("C1").Value = 100#
        ws.Range("C2").Formula = "=NA()"
        ws.Range("D1").Value = "abc"
        ws.Range("D2").Value = -0.001
        ws.Range("E1").Formula = "=SACCR_FactorAggregation(A1:A3,B1:B3)"
        ws.Range("E2").Formula = "=SACCR_FactorAggregation(A1:A2,B1:B3)"
        ws.Range("E3").Formula = "=SACCR_FactorAggregation(C1:C2,B1:B2)"
        ws.Range("E4").Formula = "=SACCR_FactorAggregation(A1,B1)"
        ws.Range("E5").Formula = "=SACCR_Multiplier(-20,282,1)"
        ws.Range("E6").Formula = "=SACCR_MaturityFactor(0.5,FALSE,10,10,0)"
        ws.Range("E7").Formula = "=SACCR_LambdaCRR(-0.002,0.001,TRUE,D1)"
        ws.Range("E8").Formula = "=SACCR_LambdaCRR(-0.002,0.001,TRUE,D2)"
        ws.Range("E9").Formula = "=SACCR_LambdaCRR(-0.002,0.001,TRUE,D3)"
        ws.Range("E10").Formula = "=SACCR_IREffectiveNotional(1,1,1,1,1,0)"
        ws.Calculate

        CheckNear "cells: blank row in both ranges", 128.452325786651, ws.Range("E1").Value, 0.000000001
        CheckError "cells: ranges of different sizes", xlErrRef, ws.Range("E2").Value
        CheckError "cells: error cell passed on", xlErrNA, ws.Range("E3").Value
        CheckNear "cells: single cells", 100#, ws.Range("E4").Value, TOL
        CheckError "cells: multiplier floor 1", xlErrNum, ws.Range("E5").Value
        CheckError "cells: maturity factor BDPerYear 0", xlErrNum, ws.Range("E6").Value
        CheckError "cells: text threshold", xlErrValue, ws.Range("E7").Value
        CheckError "cells: negative threshold", xlErrNum, ws.Range("E8").Value
        CheckNear "cells: blank threshold takes the default", 0.003, ws.Range("E9").Value, TOL
        CheckError "cells: not a correlation matrix", xlErrNum, ws.Range("E10").Value
        RemoveTemporarySheet ws, callerAlerts, callerEvents, callerActive
        Exit Sub

Unexpected:
        Fail mCurrentCase, "unexpected error " & Err.Number & ": " & Err.Description
        RemoveTemporarySheet ws, callerAlerts, callerEvents, callerActive

End Sub


Private Sub RemoveTemporarySheet( _
    ByVal ws As Worksheet, _
    ByVal callerAlerts As Boolean, _
    ByVal callerEvents As Boolean, _
    ByVal callerActive As Object)
'
'==============================================================================
' PURPOSE
'   Delete the temporary sheet of CaseWorksheetCells without a prompt and
'   give back the caller's settings and active sheet.
'
' ERROR POLICY
'   Contained: a failure is reported as a test failure of the current case.
'==============================================================================
'
        On Error GoTo Failed
        If Not ws Is Nothing Then
            Application.DisplayAlerts = False
            ws.Delete
        End If
        Application.DisplayAlerts = callerAlerts
        Application.EnableEvents = callerEvents
        If Not callerActive Is Nothing Then
            callerActive.Activate
        End If
        Exit Sub

Failed:
        Fail mCurrentCase, "could not remove the temporary sheet: " & Err.Description
        Application.DisplayAlerts = callerAlerts
        Application.EnableEvents = callerEvents

End Sub


'
'------------------------------------------------------------------------------
'
'                                   HELPERS
'
'------------------------------------------------------------------------------
'

Private Sub CheckError( _
    ByVal what As String, _
    ByVal expectedCode As Long, _
    ByVal actual As Variant)
'
'==============================================================================
' PURPOSE
'   Check that actual is the worksheet error value expectedCode.
'==============================================================================
'
    'Compare error with error: comparing a number with an error value raises.
        If IsError(actual) Then
            Check actual = CVErr(expectedCode), what & ": expected error " & expectedCode & _
                  ", found " & CStr(actual)
        Else
            Check False, what & ": expected error " & expectedCode & ", found " & DescribeResult(actual)
        End If

End Sub


Private Sub CheckNear( _
    ByVal what As String, _
    ByVal expected As Double, _
    ByVal actual As Variant, _
    ByVal tolerance As Double)
'
'==============================================================================
' PURPOSE
'   Check that actual is a number within tolerance of expected.
'==============================================================================
'
        If IsError(actual) Or Not IsNumeric(actual) Or VarType(actual) = vbBoolean Or IsEmpty(actual) Then
            Check False, what & ": expected " & CStr(expected) & ", found " & DescribeResult(actual)
        Else
            Check Abs(CDbl(actual) - expected) <= tolerance, what & ": expected " & CStr(expected) & _
                  ", found " & CStr(actual)
        End If

End Sub


Private Function DescribeResult( _
    ByVal v As Variant) _
    As String
'
'==============================================================================
' PURPOSE
'   Describe a result for a failure message.
'==============================================================================
'
        If IsError(v) Then
            DescribeResult = CStr(v)
        ElseIf IsEmpty(v) Then
            DescribeResult = "(empty)"
        Else
            DescribeResult = CStr(v)
        End If

End Function


Private Sub BeginCase( _
    ByVal caseName As String)
'
'==============================================================================
' PURPOSE
'   Count a case and print its CASE line.
'==============================================================================
'
        mCaseCount = mCaseCount + 1
        mCurrentCase = caseName
        TEST_Suite.Say "CASE=" & caseName

End Sub


Private Sub Check( _
    ByVal passed As Boolean, _
    ByVal detail As String)
'
'==============================================================================
' PURPOSE
'   Count one check and record it when it fails.
'==============================================================================
'
        mCheckCount = mCheckCount + 1
        If Not passed Then
            Fail mCurrentCase, detail
        End If

End Sub


Private Sub Fail( _
    ByVal caseName As String, _
    ByVal detail As String)
'
'==============================================================================
' PURPOSE
'   Count a failure and print it.
'==============================================================================
'
        mFailureCount = mFailureCount + 1
        TEST_Suite.Say "FAILURE=" & caseName & ": " & detail

End Sub
