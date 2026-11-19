Attribute VB_Name = "TEST_FormulaRandom"
'==============================================================================
' MODULE: TEST_FormulaRandom
'------------------------------------------------------------------------------
' PURPOSE
'   Run the comparison of the FormulaTest sheet at scale (#165): for each
'   range set of the sheet's range table, DRAWS random inputs from a
'   fixed seed, written to a test row of that set, plus deterministic
'   draws at and immediately around the branch points the table lists,
'   and the comparison of the sheet itself, |Function - Expected| /
'   max(1, |Expected|) at most FormulaTest_Tolerance, decides each one.
'   An error in Expected or Function is a failure, not a number. The
'   report names the seed and the first failing draw. The suite also
'   checks the two buttons' macros: the shipped inputs come back and
'   pass, a seed replays its draw, a range table that cannot give a valid
'   draw is refused with nothing written, nothing on the sheet is
'   volatile, and a perturbed Expected formula is caught.
'
' PUBLIC SURFACE
'   RunFormulaRandomTests is the entry point. Option Private Module keeps
'   it out of the external workbook automation API.
'
' DEPENDENCIES
'   M_FormulaTest, which draws; M_Main.ProtectWorkbook and M_Main.gSilent;
'   TEST_Kit for the suite lifecycle and the log lines (#536).
'
' WORKSHEET SAFETY
'   Writes the input cells, the seed cells, the range table and, for the
'   injected divergence, one Expected formula of the FormulaTest sheet,
'   and puts every one back, also after an unexpected error, with the
'   Excel settings; the restoration is reported apart from the
'   comparison (restore=). Use a development workbook.
'
' ERROR POLICY
'   A raised error is a failure of its case: the remaining cases still
'   run.
'
' USAGE
'   Import with the production modules, compile, then run
'   TEST_FormulaRandom.RunFormulaRandomTests from the Immediate window.
'   It ends with RESULT=PASS; cases=...; checks=...; failures=0;
'   restore=PASS.
'
' UPDATED
'   2026-11-19
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
        Private Const EXPECTED_CASES   As Long = 19            'Cases in a complete run
        Private Const EXPECTED_CHECKS  As Long = 47            'Checks in a complete run
        Private Const SUITE_SEED       As Long = 165000        'Seed of the first range set; the others follow it
        Private Const DRAWS            As Long = 1000          'Random draws of each range set
        Private Const BOUNDARY_STEP    As Double = 0.000000001 'Distance of a draw from a boundary, relative to it
        Private Const INJECTED_DRAWS   As Long = 40            'Draws made with a perturbed Expected
        Private Const FIRST_INPUT_COL  As Long = 5             'Column of Input 1
        Private Const LAST_INPUT_COL   As Long = 11            'Column of Input 7
        Private Const COL_EXPECTED     As Long = 12            'Column of Expected
        Private Const COL_RESULT       As Long = 15            'Column of Result
        Private Const INPUT_COUNT           As Long = 7             'Inputs of a test

'------------------------------------------------------------------------------
' MODULE STATE
'------------------------------------------------------------------------------
    'What the suite found, put back at the end.
        Private mSheet          As Worksheet     'The FormulaTest sheet
        Private mSavedInputs    As Variant       'Input cells: formulas or values
        Private mSavedIsFormula As Variant       'True where an input cell held a formula
        Private mSavedExpected  As Variant       'Expected formulas
        Private mSavedRanges    As Variant       'Range table values
        Private mSavedSeed      As Variant       'Seed cell
        Private mSavedDrawn     As Variant       'Cell of the seed shown
        Private mSavedTolerance As Variant       'Tolerance cell
        Private mCalculation    As XlCalculation 'Calculation mode
        Private mEvents         As Boolean       'EnableEvents
        Private mScreen         As Boolean       'ScreenUpdating
        Private mTests          As Long          'Tests on the sheet
        Private mFirstRow       As Long          'Worksheet row of the first test


'
'------------------------------------------------------------------------------
'
'                                 ENTRY POINT
'
'------------------------------------------------------------------------------
'

Public Sub RunFormulaRandomTests()
'
'==============================================================================
'                            RunFormulaRandomTests
'------------------------------------------------------------------------------
' PURPOSE
'   Run every case and print one CASE line per case, each failed check,
'   and a final RESULT line.
'
' UPDATED
'   2026-11-19
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Restored        As Boolean     'The sheet and Excel are as found

'------------------------------------------------------------------------------
' RUN
'------------------------------------------------------------------------------
        If Not TEST_Kit.BeginSuite("TEST_FormulaRandom", EXPECTED_CASES, EXPECTED_CHECKS) Then Exit Sub
        RunCases
        Restored = Restore()
        TEST_Kit.EndSuite Restored

End Sub


Private Sub RunCases()
'
'==============================================================================
'                                  RunCases
'------------------------------------------------------------------------------
' PURPOSE
'   Keep the sheet, run every case, and report an unexpected error as a
'   failure of the suite; the caller restores the sheet.
'
' UPDATED
'   2026-11-19
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Sets            As Variant     'Range sets
    Dim FirstIndex()    As Long        'First test of each
    Dim N               As Long        'Range set

'------------------------------------------------------------------------------
' RUN
'------------------------------------------------------------------------------
        On Error GoTo Failed
        M_Main.ProtectWorkbook
        M_FormulaTest.FormulaTestReload
        Save
        CaseRestoreDefaults
        CaseSeedReplay
        CaseRangesRefused
        CaseNoVolatile
        Sets = M_FormulaTest.FormulaTestFamilies(FirstIndex)
        For N = LBound(Sets) To UBound(Sets)
            CaseRange CStr(Sets(N)), FirstIndex(N), N
        Next N
        CaseInjectedDivergence CStr(Sets(LBound(Sets))), FirstIndex(LBound(Sets))
        Exit Sub

Failed:
        TEST_Kit.Fail "suite", "unexpected error " & Err.Number & ": " & Err.Description

End Sub


'
'------------------------------------------------------------------------------
'
'                                    CASES
'
'------------------------------------------------------------------------------
'

Private Sub CaseRestoreDefaults()
'
'==============================================================================
' PURPOSE
'   Restore defaults writes the shipped input sets, shows "defaults" as
'   the seed of the inputs, and all the tests pass.
'==============================================================================
'
    Dim Mismatch    As Long        'Input cells that differ from the table
    Dim Passed     As Long        'Tests that pass

        On Error GoTo Unexpected
        TEST_Kit.BeginCase "restore-defaults"
        M_FormulaTest.FormulaTestRandomize_Silent SUITE_SEED
        M_FormulaTest.FormulaTestRestore_Silent
        mSheet.Calculate
        Mismatch = DefaultsMismatch()
        TEST_Kit.Check Mismatch = 0, Mismatch & " input cells differ from the defaults table"
        Passed = PassingTests()
        TEST_Kit.Check Passed = mTests, Passed & " of " & mTests & " shipped tests pass"
        TEST_Kit.Check CStr(ThisWorkbook.Names("FormulaTest_DrawnSeed").RefersToRange.Value2) = "defaults", _
                       "the seed shown is not ""defaults"""
        Exit Sub
Unexpected:
        TEST_Kit.Fail TEST_Kit.CurrentCase, "unexpected error " & Err.Number & ": " & Err.Description
End Sub


Private Sub CaseSeedReplay()
'
'==============================================================================
' PURPOSE
'   Randomize inputs fills every input, writes the seed, and the same
'   seed gives the same inputs, typed in the seed cell or passed; another
'   seed, and the seed shown again, give other inputs.
'==============================================================================
'
    Dim First       As Variant     'Inputs of the first draw
    Dim Again       As Variant     'Inputs of a later draw
    Dim Result      As String      'Result line
    Dim Filled      As Long        'Input cells the shipped set uses that hold a value
    Dim Wanted      As Long        'Input cells the shipped set uses
    Dim Passed      As Long        'Drawn tests that pass

        On Error GoTo Unexpected
        TEST_Kit.BeginCase "seed-replay"
        Result = M_FormulaTest.FormulaTestRandomize_Silent(SUITE_SEED + 1)
        TEST_Kit.Check Left$(Result, Len("RESULT=RANDOMIZED")) = "RESULT=RANDOMIZED", Result
        First = InputValues()
        Filled = FilledCells(Wanted)
        TEST_Kit.Check Filled = Wanted And Wanted > 0, Filled & " of " & Wanted & " input cells hold a drawn value"
        Passed = PassingNow()
        TEST_Kit.Check Passed = mTests, Passed & " of " & mTests & " drawn tests pass"
        TEST_Kit.Check ThisWorkbook.Names("FormulaTest_Seed").RefersToRange.Value2 = SUITE_SEED + 1 And _
                       ThisWorkbook.Names("FormulaTest_DrawnSeed").RefersToRange.Value2 = SUITE_SEED + 1, "the seed cells do not show the seed"
        M_FormulaTest.FormulaTestRandomize_Silent SUITE_SEED + 2
        Again = InputValues()
        TEST_Kit.Check Not SameValues(First, Again), "another seed gave the same inputs"
        M_FormulaTest.FormulaTestRandomize_Silent SUITE_SEED + 1
        Again = InputValues()
        TEST_Kit.Check SameValues(First, Again), "the seed did not replay its inputs"
        M_FormulaTest.FormulaTestRandomize_Silent SUITE_SEED + 2
        ThisWorkbook.Names("FormulaTest_Seed").RefersToRange.Value2 = SUITE_SEED + 1
        M_FormulaTest.FormulaTestRandomize_Silent
        Again = InputValues()
        TEST_Kit.Check SameValues(First, Again), "a seed typed in the seed cell did not replay its inputs"
        M_FormulaTest.FormulaTestRandomize_Silent
        Again = InputValues()
        TEST_Kit.Check Not SameValues(First, Again) And ThisWorkbook.Names("FormulaTest_Seed").RefersToRange.Value2 <> SUITE_SEED + 1, _
                       "the seed unchanged did not give a new draw"
        Exit Sub
Unexpected:
        TEST_Kit.Fail TEST_Kit.CurrentCase, "unexpected error " & Err.Number & ": " & Err.Description
End Sub


Private Sub CaseRangesRefused()
'
'==============================================================================
' PURPOSE
'   A range table that cannot give a valid draw is refused, and nothing
'   is written: a Min above its Max, a limit outside the domain, P and K
'   that can never be above minus lambda, and an A that can never be
'   below D.
'==============================================================================
'
    Dim Table       As Range       'The range table
    Dim Before      As Variant     'The inputs before
    Dim DRow        As Long        'Row of D of the CDO range set
    Dim DSaved      As Variant     'D's Min and Max

        On Error GoTo Unexpected
        TEST_Kit.BeginCase "ranges-refused"
        Set Table = ThisWorkbook.Names("FormulaTest_Ranges").RefersToRange
        Before = InputValues()
        RefusedWith Table, "SD", 1, "is above Max", Before, "min above max", 20#, 10#
        RefusedWith Table, "DELTA", 3, "below the domain minimum", Before, "expiry at 0", 0#, 5#
        RefusedWith Table, "DELTA", 1, "above minus lambda", Before, "price never above 0", -2#, -1#
        DRow = TableRow(Table, "CDO", 2)
        DSaved = Array(Table.Cells(DRow, 5).Value2, Table.Cells(DRow, 6).Value2)
        Table.Cells(DRow, 5).Value2 = 0.01
        Table.Cells(DRow, 6).Value2 = 0.5
        RefusedWith Table, "CDO", 1, "A can never be below D", Before, "attachment above detachment", 0.6, 0.9
        Table.Cells(DRow, 5).Value2 = DSaved(0)
        Table.Cells(DRow, 6).Value2 = DSaved(1)
        M_FormulaTest.FormulaTestReload
        Exit Sub
Unexpected:
        TEST_Kit.Fail TEST_Kit.CurrentCase, "unexpected error " & Err.Number & ": " & Err.Description
End Sub


Private Sub CaseNoVolatile()
'
'==============================================================================
' PURPOSE
'   No formula on the sheet calls a volatile function.
'==============================================================================
'
    Dim Formulas    As Variant     'Every formula or value of the used range
    Dim R           As Long        'Row
    Dim C           As Long        'Column
    Dim Item        As Variant     'One volatile function
    Dim Found       As String      'The first one found

        On Error GoTo Unexpected
        TEST_Kit.BeginCase "no-volatile-function"
        Formulas = mSheet.UsedRange.Formula
        For R = 1 To UBound(Formulas, 1)
            For C = 1 To UBound(Formulas, 2)
                If Left$(CStr(Formulas(R, C)), 1) = "=" Then
                    For Each Item In Array("RAND(", "RANDBETWEEN(", "RANDARRAY(", "NOW(", "TODAY(", "OFFSET(", "INDIRECT(", "INFO(", "CELL(")
                        If InStr(1, UCase$(CStr(Formulas(R, C))), CStr(Item), vbBinaryCompare) > 0 And Len(Found) = 0 Then
                            Found = mSheet.UsedRange.Cells(R, C).Address(False, False) & " " & CStr(Formulas(R, C))
                        End If
                    Next Item
                End If
            Next C
        Next R
        TEST_Kit.Check Len(Found) = 0, "volatile function in " & Found
        Exit Sub
Unexpected:
        TEST_Kit.Fail TEST_Kit.CurrentCase, "unexpected error " & Err.Number & ": " & Err.Description
End Sub


Private Sub CaseRange( _
    ByVal Family As String, _
    ByVal FirstIndex As Long, _
    ByVal Ordinal As Long)
'
'==============================================================================
' PURPOSE
'   DRAWS random draws and the deterministic draws at the branch points
'   of one range set, each on the first test row that uses it: every one
'   passes the comparison of the sheet.
'==============================================================================
'
    Dim Seed        As Long        'Seed of this range set
    Dim State       As Double      'Generator state
    Dim Draw        As Long        'Draw
    Dim Inputs()    As Variant     'One draw
    Dim Problem     As String      'Why no draw was found
    Dim Detail      As String      'Why a draw failed
    Dim Failures    As Long        'Failed random draws
    Dim First       As String      'The first of them
    Dim SheetRow    As Long        'Worksheet row of the test

        On Error GoTo Unexpected
        TEST_Kit.BeginCase "random." & LCase$(Family)
        M_FormulaTest.FormulaTestReload
        SheetRow = M_FormulaTest.FormulaTestSheetRow(FirstIndex)
        Seed = SUITE_SEED + 100 + Ordinal
        State = M_FormulaTest.FormulaTestSeedState(Seed)
        For Draw = 1 To DRAWS
            ReDim Inputs(1 To INPUT_COUNT)
            If Not M_FormulaTest.FormulaTestDrawInputs(Family, State, Inputs, Problem) Then
                Failures = Failures + 1
                If Len(First) = 0 Then First = "draw " & Draw & ": " & Problem
                Exit For
            End If
            If Not RunDraw(SheetRow, Inputs, Detail) Then
                Failures = Failures + 1
                If Len(First) = 0 Then First = "draw " & Draw & ": " & Detail
            End If
        Next Draw
        TEST_Suite.Say "RANGE_SET=" & Family & "; seed=" & Seed & "; draws=" & DRAWS & "; failures=" & Failures
        TEST_Kit.Check Failures = 0, Family & " seed " & Seed & ", " & Failures & " of " & DRAWS & _
                       " random draws failed; first failing " & First
        BoundaryDraws Family, SheetRow, State
        Exit Sub
Unexpected:
        TEST_Kit.Fail TEST_Kit.CurrentCase, "unexpected error " & Err.Number & ": " & Err.Description
End Sub


Private Sub CaseInjectedDivergence( _
    ByVal Family As String, _
    ByVal Index As Long)
'
'==============================================================================
' PURPOSE
'   A perturbed Expected is caught: the draws of the first range set,
'   run with its Expected formula multiplied by 1.01, fail, the first
'   failing draw is reported, and the formula is put back.
'==============================================================================
'
    Dim SheetRow    As Long        'Worksheet row of the test
    Dim Target      As Range       'Its Expected cell
    Dim Original    As String      'The formula
    Dim State       As Double      'Generator state
    Dim Draw        As Long        'Draw
    Dim Inputs()    As Variant     'One draw
    Dim Problem     As String      'Why no draw was found
    Dim Detail      As String      'Why a draw failed
    Dim Failures    As Long        'Draws that failed
    Dim First       As String      'The first of them

        On Error GoTo Unexpected
        TEST_Kit.BeginCase "injected-divergence"
        SheetRow = M_FormulaTest.FormulaTestSheetRow(Index)
        Set Target = mSheet.Cells(SheetRow, COL_EXPECTED)
        Original = Target.Formula
        Target.Formula = "=(" & Mid$(Original, 2) & ")*1.01"
        State = M_FormulaTest.FormulaTestSeedState(SUITE_SEED + 500)
        For Draw = 1 To INJECTED_DRAWS
            ReDim Inputs(1 To INPUT_COUNT)
            If M_FormulaTest.FormulaTestDrawInputs(Family, State, Inputs, Problem) Then
                If Not RunDraw(SheetRow, Inputs, Detail) Then
                    Failures = Failures + 1
                    If Len(First) = 0 Then First = "draw " & Draw & ": " & Detail
                End If
            End If
        Next Draw
        Target.Formula = Original
        TEST_Kit.Check Failures = INJECTED_DRAWS, "a perturbed Expected failed " & Failures & " of " & INJECTED_DRAWS & " draws"
        TEST_Kit.Check Len(First) > 0 And InStr(1, First, "draw 1:", vbBinaryCompare) = 1, "the first failing draw was not reported: " & First
        TEST_Kit.Check Target.Formula = Original, "the Expected formula was not put back"
        Exit Sub
Unexpected:
        TEST_Kit.Fail TEST_Kit.CurrentCase, "unexpected error " & Err.Number & ": " & Err.Description
End Sub


'
'------------------------------------------------------------------------------
'
'                                   THE DRAWS
'
'------------------------------------------------------------------------------
'

Private Function RunDraw( _
    ByVal SheetRow As Long, _
    ByRef Inputs() As Variant, _
    ByRef Detail As String) _
    As Boolean
'
'==============================================================================
'                                   RunDraw
'------------------------------------------------------------------------------
' PURPOSE
'   Write a draw to a test row, calculate the row, and apply the sheet's
'   comparison: Result must read PASS. An error in Expected or Function,
'   or a Rel. diff that is not a number, is a failure.
'
' RETURNS
'   True when the draw passes; otherwise Detail names the inputs and the
'   two values.
'
' UPDATED
'   2026-11-19
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Block()     As Variant     'The row of inputs
    Dim K           As Long        'Input
    Dim Back        As Variant     'Expected, Function, Rel. diff and Result

'------------------------------------------------------------------------------
' WRITE AND CALCULATE
'------------------------------------------------------------------------------
        ReDim Block(1 To 1, 1 To INPUT_COUNT)
        For K = 1 To INPUT_COUNT
            Block(1, K) = Inputs(K)
        Next K
        mSheet.Range(mSheet.Cells(SheetRow, FIRST_INPUT_COL), mSheet.Cells(SheetRow, LAST_INPUT_COL)).Value2 = Block
        mSheet.Range(mSheet.Cells(SheetRow, COL_EXPECTED), mSheet.Cells(SheetRow, COL_RESULT)).Calculate
        Back = mSheet.Range(mSheet.Cells(SheetRow, COL_EXPECTED), mSheet.Cells(SheetRow, COL_RESULT)).Value2

'------------------------------------------------------------------------------
' COMPARE
'------------------------------------------------------------------------------
        If IsError(Back(1, 1)) Then
            Detail = "Expected is an error for " & InputsText(Inputs)
        ElseIf IsError(Back(1, 2)) Then
            Detail = "Function is an error for " & InputsText(Inputs) & ", Expected " & Trim$(Str$(Back(1, 1)))
        ElseIf IsError(Back(1, 3)) Then
            Detail = "Rel. diff is an error for " & InputsText(Inputs)
        ElseIf VarType(Back(1, 3)) = vbString Or CStr(Back(1, 4)) <> "PASS" Then
            Detail = "inputs " & InputsText(Inputs) & ", Expected " & Trim$(Str$(Back(1, 1))) & ", Function " & _
                     Trim$(Str$(Back(1, 2))) & ", Rel. diff " & CStr(Back(1, 3))
        Else
            RunDraw = True
        End If

End Function


Private Sub BoundaryDraws( _
    ByVal Family As String, _
    ByVal SheetRow As Long, _
    ByRef State As Double)
'
'==============================================================================
'                                BoundaryDraws
'------------------------------------------------------------------------------
' PURPOSE
'   Draws at and immediately around each branch point of the range
'   table: the boundary, and a step either side of it, in one input with
'   the others drawn. A value outside the domain, or one that breaks a
'   joint constraint, is skipped. For the replacement cost, the draws
'   put V - C at and around both TH + MTA - NICA and 0.
'
' UPDATED
'   2026-11-19
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim K           As Long        'Input
    Dim B           As Long        'Boundary
    Dim Side        As Long        'Step: -1, 0 or +1
    Dim Values()    As Double      'Boundaries of one input
    Dim Count       As Long        'How many
    Dim Inputs()    As Variant     'One draw
    Dim Problem     As String      'Why no draw was found
    Dim Detail      As String      'Why a draw failed
    Dim Made        As Long        'Boundary draws run
    Dim Failures    As Long        'Of them, failed
    Dim First       As String      'The first
    Dim Candidate   As Double      'The value tried
    Dim Distance    As Double      'Distance from the boundary

'------------------------------------------------------------------------------
' AROUND EACH BOUNDARY
'------------------------------------------------------------------------------
        For K = 1 To INPUT_COUNT
            Count = M_FormulaTest.FormulaTestBoundaries(Family, K, Values)
            For B = 1 To Count
                For Side = -1 To 1
                    ReDim Inputs(1 To INPUT_COUNT)
                    If M_FormulaTest.FormulaTestDrawInputs(Family, State, Inputs, Problem) Then
                        If VarType(Inputs(K)) = vbLong Then
                            Distance = 1#
                        Else
                            Distance = BOUNDARY_STEP * IIf(Abs(Values(B)) > 1#, Abs(Values(B)), 1#)
                        End If
                        Candidate = Values(B) + Side * Distance
                        If VarType(Inputs(K)) = vbLong Then
                            Inputs(K) = CLng(Candidate)
                        Else
                            Inputs(K) = Candidate
                        End If
                        If M_FormulaTest.FormulaTestInDomain(Family, K, Candidate) And M_FormulaTest.FormulaTestJointOk(Family, Inputs) Then
                            RunOne SheetRow, Inputs, Made, Failures, First
                        End If
                    End If
                Next Side
            Next B
        Next K

'------------------------------------------------------------------------------
' THE REPLACEMENT COST'S TWO KINKS
'------------------------------------------------------------------------------
        If Family = "RC" Then
            For Side = -1 To 1
                ReDim Inputs(1 To INPUT_COUNT)
                If M_FormulaTest.FormulaTestDrawInputs(Family, State, Inputs, Problem) Then
                    Inputs(3) = True
                    Inputs(1) = Inputs(2) + Inputs(4) + Inputs(5) - Inputs(6) + Side * BOUNDARY_STEP * 1000#
                    RunOne SheetRow, Inputs, Made, Failures, First
                    Inputs(1) = Inputs(2) + Side * BOUNDARY_STEP * 1000#
                    RunOne SheetRow, Inputs, Made, Failures, First
                End If
            Next Side
        End If
        TEST_Suite.Say "BOUNDARIES=" & Family & "; draws=" & Made & "; failures=" & Failures
        TEST_Kit.Check Failures = 0, Family & ", " & Failures & " of " & Made & " boundary draws failed; first failing " & First

End Sub


Private Sub RunOne( _
    ByVal SheetRow As Long, _
    ByRef Inputs() As Variant, _
    ByRef Made As Long, _
    ByRef Failures As Long, _
    ByRef First As String)
'
'==============================================================================
'                                    RunOne
'------------------------------------------------------------------------------
' PURPOSE
'   Run one boundary draw and count it.
'
' UPDATED
'   2026-11-19
'==============================================================================
'
    Dim Detail  As String     'Why it failed

        Made = Made + 1
        If Not RunDraw(SheetRow, Inputs, Detail) Then
            Failures = Failures + 1
            If Len(First) = 0 Then First = "boundary draw " & Made & ": " & Detail
        End If

End Sub


Private Function InputsText( _
    ByRef Inputs() As Variant) _
    As String
'
'==============================================================================
'                                 InputsText
'------------------------------------------------------------------------------
' PURPOSE
'   A draw as text, in the order of the inputs, for a failure's report.
'
' UPDATED
'   2026-11-19
'==============================================================================
'
    Dim K   As Long     'Input

        For K = 1 To INPUT_COUNT
            If Not IsEmpty(Inputs(K)) Then
                If VarType(Inputs(K)) = vbBoolean Then
                    InputsText = InputsText & IIf(Inputs(K), "TRUE", "FALSE") & "; "
                Else
                    InputsText = InputsText & Trim$(Str$(Inputs(K))) & "; "
                End If
            End If
        Next K

End Function


'
'------------------------------------------------------------------------------
'
'                                SHEET HELPERS
'
'------------------------------------------------------------------------------
'

Private Function InputValues() As Variant
'
'==============================================================================
'                                InputValues
'------------------------------------------------------------------------------
' PURPOSE
'   The input cells of every test, as values.
'
' UPDATED
'   2026-11-19
'==============================================================================
'
        InputValues = mSheet.Range(mSheet.Cells(mFirstRow, FIRST_INPUT_COL), _
                                   mSheet.Cells(mFirstRow + mTests - 1, LAST_INPUT_COL)).Value2

End Function


Private Function SameValues( _
    ByVal A As Variant, _
    ByVal B As Variant) _
    As Boolean
'
'==============================================================================
'                                SameValues
'------------------------------------------------------------------------------
' PURPOSE
'   Whether two blocks of values are the same, cell by cell, with the
'   same type.
'
' UPDATED
'   2026-11-19
'==============================================================================
'
    Dim R   As Long     'Row
    Dim C   As Long     'Column

        SameValues = True
        For R = 1 To UBound(A, 1)
            For C = 1 To UBound(A, 2)
                If VarType(A(R, C)) <> VarType(B(R, C)) Then
                    SameValues = False
                    Exit Function
                ElseIf Not IsEmpty(A(R, C)) Then
                    If A(R, C) <> B(R, C) Then
                        SameValues = False
                        Exit Function
                    End If
                End If
            Next C
        Next R

End Function


Private Function PassingTests() As Long
'
'==============================================================================
'                                PassingTests
'------------------------------------------------------------------------------
' PURPOSE
'   How many tests read PASS in Result.
'
' UPDATED
'   2026-11-19
'==============================================================================
'
    Dim R   As Long     'Test

        For R = 0 To mTests - 1
            If CStr(mSheet.Cells(mFirstRow + R, COL_RESULT).Value2) = "PASS" Then PassingTests = PassingTests + 1
        Next R

End Function


Private Function PassingNow() As Long
        mSheet.Calculate
        PassingNow = PassingTests()
End Function


Private Function DefaultsMismatch() As Long
'
'==============================================================================
'                              DefaultsMismatch
'------------------------------------------------------------------------------
' PURPOSE
'   How many input cells of the sheet differ from the defaults table:
'   a cell the table fills holds its value; one it leaves empty is empty
'   when the user may edit it.
'
' UPDATED
'   2026-11-19
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Table       As Variant     'The defaults table
    Dim R           As Long        'Test
    Dim K           As Long        'Input
    Dim Cell        As Range       'Its cell
    Dim Wanted      As Variant     'The default

'------------------------------------------------------------------------------
' COMPARE
'------------------------------------------------------------------------------
        Table = ThisWorkbook.Names("FormulaTest_Defaults").RefersToRange.Value2
        For R = 1 To mTests
            For K = 1 To INPUT_COUNT
                Set Cell = mSheet.Cells(mFirstRow + R - 1, FIRST_INPUT_COL + K - 1)
                Wanted = Table(R, 2 + K)
                If IsEmpty(Wanted) Then
                    If Not Cell.Locked And Not IsEmpty(Cell.Value2) Then DefaultsMismatch = DefaultsMismatch + 1
                ElseIf VarType(Cell.Value2) <> VarType(Wanted) Then
                    DefaultsMismatch = DefaultsMismatch + 1
                ElseIf Cell.Value2 <> Wanted Then
                    DefaultsMismatch = DefaultsMismatch + 1
                End If
            Next K
        Next R

End Function


Private Function FilledCells( _
    ByRef Wanted As Long) _
    As Long
'
'==============================================================================
'                                FilledCells
'------------------------------------------------------------------------------
' PURPOSE
'   How many input cells the shipped set uses hold a value, and how many
'   the shipped set uses.
'
' UPDATED
'   2026-11-19
'==============================================================================
'
    Dim Table       As Variant     'The defaults table
    Dim R           As Long        'Test
    Dim K           As Long        'Input

        Table = ThisWorkbook.Names("FormulaTest_Defaults").RefersToRange.Value2
        Wanted = 0
        For R = 1 To mTests
            For K = 1 To INPUT_COUNT
                If Not IsEmpty(Table(R, 2 + K)) Then
                    Wanted = Wanted + 1
                    If Not IsEmpty(mSheet.Cells(mFirstRow + R - 1, FIRST_INPUT_COL + K - 1).Value2) Then FilledCells = FilledCells + 1
                End If
            Next K
        Next R

End Function


Private Function TableRow( _
    ByVal Table As Range, _
    ByVal Family As String, _
    ByVal InputNo As Long) _
    As Long
'
'==============================================================================
'                                 TableRow
'------------------------------------------------------------------------------
' PURPOSE
'   The row of the range table for an input of a range set.
'
' UPDATED
'   2026-11-19
'==============================================================================
'
    Dim R   As Long     'Row of the table

        For R = 1 To Table.Rows.Count
            If CStr(Table.Cells(R, 1).Value2) = Family And Table.Cells(R, 3).Value2 = InputNo Then
                TableRow = R
                Exit Function
            End If
        Next R

End Function


Private Sub RefusedWith( _
    ByVal Table As Range, _
    ByVal Family As String, _
    ByVal InputNo As Long, _
    ByVal Expect As String, _
    ByVal Before As Variant, _
    ByVal Label As String, _
    ByVal NewMin As Double, _
    ByVal NewMax As Double)
'
'==============================================================================
'                                RefusedWith
'------------------------------------------------------------------------------
' PURPOSE
'   Set the Min and Max of one input of one range set, run Randomize
'   inputs, check that it is refused with a reason that contains Expect
'   and that no input changed, then put the two cells back.
'
' UPDATED
'   2026-11-19
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim R           As Long        'Row of the table
    Dim Saved       As Variant     'Min and Max before
    Dim Result      As String      'Result line

'------------------------------------------------------------------------------
' PATCH, RUN AND PUT BACK
'------------------------------------------------------------------------------
        R = TableRow(Table, Family, InputNo)
        Saved = Array(Table.Cells(R, 5).Value2, Table.Cells(R, 6).Value2)
        Table.Cells(R, 5).Value2 = NewMin
        Table.Cells(R, 6).Value2 = NewMax
        Result = M_FormulaTest.FormulaTestRandomize_Silent(SUITE_SEED + 3)
        Table.Cells(R, 5).Value2 = Saved(0)
        Table.Cells(R, 6).Value2 = Saved(1)
        M_FormulaTest.FormulaTestReload
        TEST_Kit.Check Left$(Result, Len("RESULT=REFUSED")) = "RESULT=REFUSED" And InStr(1, Result, Expect, vbBinaryCompare) > 0 And _
                       SameValues(Before, InputValues()), Label & ": " & Result

End Sub


'
'------------------------------------------------------------------------------
'
'                         SAVING AND RESTORING THE SHEET
'
'------------------------------------------------------------------------------
'

Private Sub Save()
'
'==============================================================================
'                                     Save
'------------------------------------------------------------------------------
' PURPOSE
'   Keep everything the suite writes, and the Excel settings, then
'   switch screen updating and events off and calculation to manual.
'
' UPDATED
'   2026-11-19
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim R           As Long        'Test
    Dim C           As Long        'Input
    Dim Cell        As Range       'Its cell
    Dim Saved()     As Variant     'Formulas or values of the inputs
    Dim IsFormula() As Variant     'Where there was a formula
    Dim Expected()  As Variant     'Expected formulas

'------------------------------------------------------------------------------
' KEEP
'------------------------------------------------------------------------------
        Set mSheet = ThisWorkbook.Names("FormulaTest_Tests").RefersToRange.Worksheet
        mTests = ThisWorkbook.Names("FormulaTest_Tests").RefersToRange.Rows.Count
        mFirstRow = ThisWorkbook.Names("FormulaTest_Tests").RefersToRange.Row
        ReDim Saved(1 To mTests, 1 To INPUT_COUNT)
        ReDim IsFormula(1 To mTests, 1 To INPUT_COUNT)
        ReDim Expected(1 To mTests, 1 To 1)
        For R = 1 To mTests
            For C = 1 To INPUT_COUNT
                Set Cell = mSheet.Cells(mFirstRow + R - 1, FIRST_INPUT_COL + C - 1)
                IsFormula(R, C) = Cell.HasFormula
                If Cell.HasFormula Then Saved(R, C) = Cell.Formula Else Saved(R, C) = Cell.Value2
            Next C
            Expected(R, 1) = mSheet.Cells(mFirstRow + R - 1, COL_EXPECTED).Formula
        Next R
        mSavedInputs = Saved
        mSavedIsFormula = IsFormula
        mSavedExpected = Expected
        mSavedRanges = ThisWorkbook.Names("FormulaTest_Ranges").RefersToRange.Value2
        mSavedSeed = ThisWorkbook.Names("FormulaTest_Seed").RefersToRange.Formula
        mSavedDrawn = ThisWorkbook.Names("FormulaTest_DrawnSeed").RefersToRange.Formula
        mSavedTolerance = ThisWorkbook.Names("FormulaTest_Tolerance").RefersToRange.Formula
        mCalculation = Application.Calculation
        mEvents = Application.EnableEvents
        mScreen = Application.ScreenUpdating
        Application.ScreenUpdating = False
        Application.EnableEvents = False
        Application.Calculation = xlCalculationManual

End Sub


Private Function Restore() As Boolean
'
'==============================================================================
'                                    Restore
'------------------------------------------------------------------------------
' PURPOSE
'   Put back the inputs (formulas where there were formulas), the
'   Expected formulas, the range table, the seed cells and the tolerance,
'   then the Excel settings, and read each one back.
'
' RETURNS
'   True when everything reads back as it was; the first difference is
'   printed as a FAILURE line.
'
' UPDATED
'   2026-11-19
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim R           As Long        'Test
    Dim C           As Long        'Input
    Dim Cell        As Range       'Its cell
    Dim Table       As Range       'The range table
    Dim Current        As Variant     'The table now
    Dim Wrong       As Long        'Cells that did not come back

'------------------------------------------------------------------------------
' PUT BACK
'------------------------------------------------------------------------------
        On Error GoTo Failed
        For R = 1 To mTests
            mSheet.Cells(mFirstRow + R - 1, COL_EXPECTED).Formula = mSavedExpected(R, 1)
            For C = 1 To INPUT_COUNT
                Set Cell = mSheet.Cells(mFirstRow + R - 1, FIRST_INPUT_COL + C - 1)
                If mSavedIsFormula(R, C) Then
                    Cell.Formula = mSavedInputs(R, C)
                ElseIf IsEmpty(mSavedInputs(R, C)) Then
                    Cell.ClearContents
                Else
                    Cell.Value2 = mSavedInputs(R, C)
                End If
            Next C
        Next R
        Set Table = ThisWorkbook.Names("FormulaTest_Ranges").RefersToRange
        Table.Value2 = mSavedRanges
        ThisWorkbook.Names("FormulaTest_Seed").RefersToRange.Formula = mSavedSeed
        ThisWorkbook.Names("FormulaTest_DrawnSeed").RefersToRange.Formula = mSavedDrawn
        ThisWorkbook.Names("FormulaTest_Tolerance").RefersToRange.Formula = mSavedTolerance
        mSheet.Calculate

'------------------------------------------------------------------------------
' READ BACK
'------------------------------------------------------------------------------
        For R = 1 To mTests
            If mSheet.Cells(mFirstRow + R - 1, COL_EXPECTED).Formula <> mSavedExpected(R, 1) Then Wrong = Wrong + 1
            For C = 1 To INPUT_COUNT
                Set Cell = mSheet.Cells(mFirstRow + R - 1, FIRST_INPUT_COL + C - 1)
                If mSavedIsFormula(R, C) Then
                    If Cell.Formula <> mSavedInputs(R, C) Then Wrong = Wrong + 1
                ElseIf VarType(Cell.Value2) <> VarType(mSavedInputs(R, C)) Then
                    Wrong = Wrong + 1
                ElseIf Not IsEmpty(Cell.Value2) Then
                    If Cell.Value2 <> mSavedInputs(R, C) Then Wrong = Wrong + 1
                End If
            Next C
        Next R
        Current = Table.Value2
        For R = 1 To UBound(Current, 1)
            For C = 1 To UBound(Current, 2)
                If VarType(Current(R, C)) <> VarType(mSavedRanges(R, C)) Then
                    Wrong = Wrong + 1
                ElseIf Not IsEmpty(Current(R, C)) Then
                    If Current(R, C) <> mSavedRanges(R, C) Then Wrong = Wrong + 1
                End If
            Next C
        Next R
        If ThisWorkbook.Names("FormulaTest_Seed").RefersToRange.Formula <> mSavedSeed Then Wrong = Wrong + 1
        If ThisWorkbook.Names("FormulaTest_DrawnSeed").RefersToRange.Formula <> mSavedDrawn Then Wrong = Wrong + 1
        If ThisWorkbook.Names("FormulaTest_Tolerance").RefersToRange.Formula <> mSavedTolerance Then Wrong = Wrong + 1
        M_FormulaTest.FormulaTestReload
        Application.Calculation = mCalculation
        Application.EnableEvents = mEvents
        Application.ScreenUpdating = mScreen
        If Wrong > 0 Then
            TEST_Kit.Fail "restore", Wrong & " cells of the FormulaTest sheet did not come back as they were"
        End If
        Restore = (Wrong = 0)
        Exit Function

Failed:
        Application.Calculation = mCalculation
        Application.EnableEvents = mEvents
        Application.ScreenUpdating = mScreen
        TEST_Kit.Fail "restore", "error " & Err.Number & ": " & Err.Description

End Function
