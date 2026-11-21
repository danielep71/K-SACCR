Attribute VB_Name = "M_FormulaTest"
'==============================================================================
' MODULE: M_FormulaTest
'------------------------------------------------------------------------------
' PURPOSE
'   Draw reproducible random inputs for the unit tests of the FormulaTest
'   sheet (#165) and put the shipped inputs back. Each input of each test
'   row is drawn within the range the sheet's range table gives it, from a
'   seed that is written next to the tolerance, so that any draw can be
'   replayed by entering its seed. Nothing on the sheet is volatile: the
'   draws are values written by a macro.
'
' PUBLIC SURFACE
'   RandomizeFormulaTestInputs and RestoreFormulaTestDefaults: the macros
'   of the two buttons on the FormulaTest sheet.
'   FormulaTestRandomize_Silent and FormulaTestRestore_Silent: the same
'   for automation and tests, returning a result line.
'   FormulaTestRangeProblem: why the range table cannot be used; "" when
'   it can.
'   FormulaTestReload, FormulaTestSeedState, FormulaTestNext,
'   FormulaTestFamilies, FormulaTestDrawInputs, FormulaTestJointOk,
'   FormulaTestBoundaries, FormulaTestInDomain, FormulaTestWriteInputs and
'   FormulaTestSheetRow: the draw itself, which the randomized test suite
'   (tests/modules/TEST_FormulaRandom.bas) runs at scale.
'
' LAYOUT CONTRACT
'   The sheet is found through workbook names, not addresses:
'   FormulaTest_Tests is the block of the 22 tests, from the Test column
'   to Result; its columns are Test, Arguments, Input 1 to Input 7,
'   Expected, Function, Rel. diff and Result.
'   FormulaTest_Ranges is the range table: Range set, Argument, Input,
'   Kind (number, integer or boolean), Min, Max, Domain min, Domain max,
'   Boundary 1 and Boundary 2. A boolean is drawn when its Min is FALSE
'   and its Max TRUE, and fixed when they are equal.
'   FormulaTest_Defaults holds the shipped inputs: Test, Range set and
'   Input 1 to Input 7, one row for each row of FormulaTest_Tests.
'   FormulaTest_Seed holds the seed, which the user may type, and
'   FormulaTest_DrawnSeed the seed of the inputs on the sheet, or
'   "defaults".
'
' RANDOM NUMBERS
'   The generator of Park and Miller (x := 16807 x mod 2147483647) in
'   Double arithmetic, which is exact here, not Rnd: a seed gives the same
'   draw in every Excel version and on 32-bit and 64-bit. The first
'   WARMUP values are dropped, because the first value of this generator
'   is nearly proportional to its seed.
'
' ERROR POLICY
'   A range table that cannot give a valid draw is refused with a message
'   and nothing is written: ranges outside a function's domain (#42),
'   Min above Max, no draw that meets the joint constraints (A below D,
'   P and K above minus lambda, E after S). A draw that meets a joint
'   constraint is searched for JOINT_TRIES times at most.
'
' DEPENDENCIES
'   M_Main.ProtectWorkbook and M_Main.gSilent.
'
' KNOWN DEVIATION
'   This standard module lives in src/workbook, as it holds the macros of
'   a sheet's buttons. See the known deviations in
'   docs/REPOSITORY_STRUCTURE.md.
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
    'Require explicit declarations. The macros must stay visible to the
    'sheet buttons, so this module is not private.
    Option Explicit

'------------------------------------------------------------------------------
' MODULE CONSTANTS
'------------------------------------------------------------------------------
    'Workbook names of the layout contract.
        Private Const NAME_TESTS        As String = "FormulaTest_Tests"
        Private Const NAME_RANGES       As String = "FormulaTest_Ranges"
        Private Const NAME_DEFAULTS     As String = "FormulaTest_Defaults"
        Private Const NAME_SEED         As String = "FormulaTest_Seed"
        Private Const NAME_DRAWN        As String = "FormulaTest_DrawnSeed"

    'Columns of the range table.
        Private Const RG_FAMILY         As Long = 1
        Private Const RG_ARGUMENT       As Long = 2
        Private Const RG_INPUT          As Long = 3
        Private Const RG_KIND           As Long = 4
        Private Const RG_MIN            As Long = 5
        Private Const RG_MAX            As Long = 6
        Private Const RG_DOMAIN_MIN     As Long = 7
        Private Const RG_DOMAIN_MAX     As Long = 8
        Private Const RG_BOUNDARY_1     As Long = 9
        Private Const RG_BOUNDARY_2     As Long = 10

    'Columns of the defaults table and of the test block.
        Private Const DF_FAMILY         As Long = 2
        Private Const DF_FIRST_INPUT    As Long = 3
        Private Const TS_FIRST_INPUT    As Long = 3

    'Inputs of a test.
        Private Const INPUT_COUNT       As Long = 7

    'Random numbers.
        Private Const GEN_MODULUS       As Double = 2147483647#
        Private Const GEN_MULTIPLIER    As Double = 16807#
        Private Const WARMUP            As Long = 8
        Private Const MAX_SEED          As Long = 2147483646
        Private Const JOINT_TRIES       As Long = 1000

    'Title of every message box: the project name.
        Private Const MSG_TITLE         As String = "K-SACCR"

'------------------------------------------------------------------------------
' MODULE STATE
'------------------------------------------------------------------------------
    'The two tables, read once and kept for the draws.
        Private mRanges     As Variant      'Range table values; Empty until read
        Private mDefaults   As Variant      'Defaults table values; Empty until read


'
'------------------------------------------------------------------------------
'
'                                BUTTON MACROS
'
'------------------------------------------------------------------------------
'

Public Sub RandomizeFormulaTestInputs()
'
'==============================================================================
'                          RandomizeFormulaTestInputs
'------------------------------------------------------------------------------
' PURPOSE
'   Draw new inputs for every test of the FormulaTest sheet and show the
'   outcome. The seed is the one typed next to the tolerance when it
'   differs from the seed of the inputs shown; with the seed unchanged, or
'   empty, a new seed is taken. The seed used is written back.
'
' USAGE
'   The Randomize inputs button of the FormulaTest sheet, or the Macros
'   dialog.
'
' ERROR POLICY
'   Reports an error in a message; never raises. Nothing is written when
'   the draw is refused.
'
' UPDATED
'   2026-11-19
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ResultLine  As String    'What the draw returned

'------------------------------------------------------------------------------
' DRAW AND REPORT
'------------------------------------------------------------------------------
        On Error GoTo Failed
        ResultLine = WithQuietExcel("randomize", 0)
        If Left$(ResultLine, Len("RESULT=RANDOMIZED")) = "RESULT=RANDOMIZED" Then
            Say "New inputs drawn for every test, from seed " & Field(ResultLine, "seed") & "." & vbLf & vbLf & _
                "Enter this seed in the Seed cell and press Randomize inputs again to replay the draw.", vbInformation
        Else
            Say "No inputs were changed." & vbLf & vbLf & Field(ResultLine, "reason"), vbExclamation
        End If
        Exit Sub

Failed:
        Say "Randomize inputs: " & Err.Description, vbCritical

End Sub


Public Sub RestoreFormulaTestDefaults()
'
'==============================================================================
'                          RestoreFormulaTestDefaults
'------------------------------------------------------------------------------
' PURPOSE
'   Write the shipped input sets back to every test of the FormulaTest
'   sheet and show the outcome.
'
' USAGE
'   The Restore defaults button of the FormulaTest sheet, or the Macros
'   dialog.
'
' ERROR POLICY
'   Reports an error in a message; never raises.
'
' UPDATED
'   2026-11-19
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ResultLine  As String    'What the restore returned

'------------------------------------------------------------------------------
' RESTORE AND REPORT
'------------------------------------------------------------------------------
        On Error GoTo Failed
        ResultLine = WithQuietExcel("restore", 0)
        If Left$(ResultLine, Len("RESULT=RESTORED")) = "RESULT=RESTORED" Then
            Say "The shipped inputs of the " & Field(ResultLine, "rows") & " tests are back.", vbInformation
        Else
            Say "No inputs were changed." & vbLf & vbLf & Field(ResultLine, "reason"), vbExclamation
        End If
        Exit Sub

Failed:
        Say "Restore defaults: " & Err.Description, vbCritical

End Sub


Private Function WithQuietExcel( _
    ByVal What As String, _
    ByVal Seed As Long) _
    As String
'
'==============================================================================
'                                WithQuietExcel
'------------------------------------------------------------------------------
' PURPOSE
'   Run the draw or the restore with screen updating and events off, and
'   put both back, also after an error.
'
' UPDATED
'   2026-11-19
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim WasScreen   As Boolean    'ScreenUpdating before
    Dim WasEvents   As Boolean    'EnableEvents before

'------------------------------------------------------------------------------
' RUN
'------------------------------------------------------------------------------
        WasScreen = Application.ScreenUpdating
        WasEvents = Application.EnableEvents
        Application.ScreenUpdating = False
        Application.EnableEvents = False
        On Error GoTo Failed
        If What = "randomize" Then
            WithQuietExcel = FormulaTestRandomize_Silent(Seed)
        Else
            WithQuietExcel = FormulaTestRestore_Silent()
        End If
        Application.EnableEvents = WasEvents
        Application.ScreenUpdating = WasScreen
        Exit Function

Failed:
        Application.EnableEvents = WasEvents
        Application.ScreenUpdating = WasScreen
        Err.Raise Err.Number, Err.Source, Err.Description

End Function


Private Sub Say( _
    ByVal Text As String, _
    ByVal Style As VbMsgBoxStyle)
'
'==============================================================================
'                                     Say
'------------------------------------------------------------------------------
' PURPOSE
'   Show a message, or print it to the Immediate window in silent mode.
'
' UPDATED
'   2026-11-19
'==============================================================================
'
        If M_Main.gSilent Then
            Debug.Print Replace(Text, vbLf, " ")
        Else
            MsgBox Text, Style, MSG_TITLE
        End If

End Sub


Private Function Field( _
    ByVal ResultLine As String, _
    ByVal Key As String) _
    As String
'
'==============================================================================
'                                    Field
'------------------------------------------------------------------------------
' PURPOSE
'   The value of "Key=" in a result line whose fields are separated by
'   semicolons and a space.
'
' UPDATED
'   2026-11-19
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Item    As Variant    'One field

'------------------------------------------------------------------------------
' FIND
'------------------------------------------------------------------------------
        For Each Item In Split(ResultLine, "; ")
            If Left$(CStr(Item), Len(Key) + 1) = Key & "=" Then
                Field = Mid$(CStr(Item), Len(Key) + 2)
                Exit Function
            End If
        Next Item

End Function


'
'------------------------------------------------------------------------------
'
'                          RANDOMIZE AND RESTORE
'
'------------------------------------------------------------------------------
'

Public Function FormulaTestRandomize_Silent( _
    Optional ByVal Seed As Long = 0) _
    As String
'
'==============================================================================
'                         FormulaTestRandomize_Silent
'------------------------------------------------------------------------------
' PURPOSE
'   Draw new inputs for every test, for automation and tests, and write
'   them as values with their seed.
'
' INPUTS
'   Seed: the seed to use, 1 to 2147483646. 0 takes the seed from the
'   sheet as RandomizeFormulaTestInputs does.
'
' RETURNS
'   "RESULT=RANDOMIZED; seed=N; rows=R; cells=C", or "RESULT=REFUSED;
'   reason=..." when nothing was written: a range table that cannot give
'   a valid draw, or a seed that is not a whole number in range.
'
' ERROR POLICY
'   Raises a VBA error when the layout contract is not met (a name is
'   missing). Every draw is made before any cell is written, so a refusal
'   leaves the sheet as it was.
'
' UPDATED
'   2026-11-19
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Problem     As String      'Why the draw is refused
    Dim Chosen      As Long        'Seed used
    Dim State       As Double      'Generator state
    Dim Tests       As Long        'Tests
    Dim Index       As Long        'Test being drawn
    Dim Draws()     As Variant     'The drawn inputs of each test
    Dim Inputs()    As Variant     'One test's inputs
    Dim Written     As Long        'Input cells written
    Dim K           As Long        'Input

'------------------------------------------------------------------------------
' CHECK THE TABLES AND CHOOSE THE SEED
'------------------------------------------------------------------------------
        M_Main.ProtectWorkbook
        FormulaTestReload
        Problem = FormulaTestRangeProblem()
        If Len(Problem) > 0 Then
            FormulaTestRandomize_Silent = "RESULT=REFUSED; reason=" & Problem
            Exit Function
        End If
        Chosen = Seed
        If Chosen = 0 Then
            Chosen = SeedFromSheet(Problem)
            If Len(Problem) > 0 Then
                FormulaTestRandomize_Silent = "RESULT=REFUSED; reason=" & Problem
                Exit Function
            End If
        End If
        If Chosen < 1 Or Chosen > MAX_SEED Then
            FormulaTestRandomize_Silent = "RESULT=REFUSED; reason=The seed must be a whole number from 1 to " & MAX_SEED & "."
            Exit Function
        End If

'------------------------------------------------------------------------------
' DRAW EVERY TEST BEFORE WRITING ANY
'------------------------------------------------------------------------------
        State = FormulaTestSeedState(Chosen)
        Tests = UBound(mDefaults, 1)
        ReDim Draws(1 To Tests)
        For Index = 1 To Tests
            ReDim Inputs(1 To INPUT_COUNT)
            If Not FormulaTestDrawInputs(CStr(mDefaults(Index, DF_FAMILY)), State, Inputs, Problem) Then
                FormulaTestRandomize_Silent = "RESULT=REFUSED; reason=" & Problem
                Exit Function
            End If
            Draws(Index) = Inputs
        Next Index

'------------------------------------------------------------------------------
' WRITE
'------------------------------------------------------------------------------
        For Index = 1 To Tests
            Inputs = Draws(Index)
            FormulaTestWriteInputs Index, Inputs
            For K = 1 To INPUT_COUNT
                If Not IsEmpty(Inputs(K)) Then Written = Written + 1
            Next K
        Next Index
        SeedCell().Value2 = Chosen
        DrawnCell().Value2 = Chosen
        TestsRange().Worksheet.Calculate
        FormulaTestRandomize_Silent = "RESULT=RANDOMIZED; seed=" & Chosen & "; rows=" & Tests & "; cells=" & Written

End Function


Public Function FormulaTestRestore_Silent() As String
'
'==============================================================================
'                          FormulaTestRestore_Silent
'------------------------------------------------------------------------------
' PURPOSE
'   Write the shipped inputs back to every test, for automation and
'   tests. An input cell the shipped set leaves empty is emptied when the
'   user may edit it. The seed cell is left as it is, and the seed of the
'   inputs shown becomes "defaults".
'
' RETURNS
'   "RESULT=RESTORED; rows=R; cells=C".
'
' ERROR POLICY
'   Raises a VBA error when the layout contract is not met.
'
' UPDATED
'   2026-11-19
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Index   As Long      'Test
    Dim K       As Long      'Input
    Dim Target  As Range     'Input cell
    Dim Written As Long      'Input cells written

'------------------------------------------------------------------------------
' WRITE
'------------------------------------------------------------------------------
        M_Main.ProtectWorkbook
        FormulaTestReload
        For Index = 1 To UBound(mDefaults, 1)
            For K = 1 To INPUT_COUNT
                Set Target = InputCell(Index, K)
                If Not IsEmpty(mDefaults(Index, DF_FIRST_INPUT + K - 1)) Then
                    Target.Value2 = mDefaults(Index, DF_FIRST_INPUT + K - 1)
                    Written = Written + 1
                ElseIf Not Target.Locked Then
                    Target.ClearContents
                End If
            Next K
        Next Index
        DrawnCell().Value2 = "defaults"
        TestsRange().Worksheet.Calculate
        FormulaTestRestore_Silent = "RESULT=RESTORED; rows=" & UBound(mDefaults, 1) & "; cells=" & Written

End Function


Private Function SeedFromSheet( _
    ByRef Problem As String) _
    As Long
'
'==============================================================================
'                                SeedFromSheet
'------------------------------------------------------------------------------
' PURPOSE
'   The seed to draw from: the one typed in the Seed cell when it differs
'   from the seed of the inputs shown, otherwise a new one.
'
' RETURNS
'   The seed; Problem says why there is none.
'
' UPDATED
'   2026-11-19
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Typed       As Variant     'The Seed cell
    Dim Shown       As Variant     'The seed of the inputs shown

'------------------------------------------------------------------------------
' CHOOSE
'------------------------------------------------------------------------------
        Problem = ""
        Typed = SeedCell().Value2
        Shown = DrawnCell().Value2
        If IsEmpty(Typed) Then
            SeedFromSheet = NewSeed()
        ElseIf IsError(Typed) Or Not IsNumeric(Typed) Or VarType(Typed) = vbString Then
            Problem = "The seed must be a whole number from 1 to " & MAX_SEED & "."
        ElseIf Typed <> Int(Typed) Or Typed < 1 Or Typed > MAX_SEED Then
            Problem = "The seed must be a whole number from 1 to " & MAX_SEED & "."
        ElseIf Not IsError(Shown) And IsNumeric(Shown) And VarType(Shown) <> vbString And Not IsEmpty(Shown) Then
            If Shown = Typed Then
                SeedFromSheet = NewSeed()
            Else
                SeedFromSheet = CLng(Typed)
            End If
        Else
            SeedFromSheet = CLng(Typed)
        End If

End Function


Private Function NewSeed() As Long
'
'==============================================================================
'                                   NewSeed
'------------------------------------------------------------------------------
' PURPOSE
'   A seed from the clock, 1 to 2147483646.
'
' UPDATED
'   2026-11-19
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Ticks   As Double     'Milliseconds since the date origin

'------------------------------------------------------------------------------
' TAKE
'------------------------------------------------------------------------------
        Ticks = Int(Date) * 86400000# + Int(Timer * 1000#)
        NewSeed = 1 + CLng(Ticks - Int(Ticks / MAX_SEED) * MAX_SEED)

End Function


'
'------------------------------------------------------------------------------
'
'                                 THE TABLES
'
'------------------------------------------------------------------------------
'

Public Sub FormulaTestReload()
'
'==============================================================================
'                              FormulaTestReload
'------------------------------------------------------------------------------
' PURPOSE
'   Read the range and defaults tables from the sheet again. The macros
'   do it when they start; a caller that edits a table in between calls
'   it too.
'
' ERROR POLICY
'   Raises a VBA error when a name is missing.
'
' UPDATED
'   2026-11-19
'==============================================================================
'
        mRanges = ThisWorkbook.Names(NAME_RANGES).RefersToRange.Value2
        mDefaults = ThisWorkbook.Names(NAME_DEFAULTS).RefersToRange.Value2

End Sub


Private Sub EnsureTables()
'
'==============================================================================
'                                 EnsureTables
'------------------------------------------------------------------------------
' PURPOSE
'   Read the tables when none has been read, or when the VBA state was
'   reset since.
'
' UPDATED
'   2026-11-19
'==============================================================================
'
        If IsEmpty(mRanges) Or IsEmpty(mDefaults) Then FormulaTestReload

End Sub


Private Function TestsRange() As Range
        Set TestsRange = ThisWorkbook.Names(NAME_TESTS).RefersToRange
End Function


Private Function SeedCell() As Range
        Set SeedCell = ThisWorkbook.Names(NAME_SEED).RefersToRange
End Function


Private Function DrawnCell() As Range
        Set DrawnCell = ThisWorkbook.Names(NAME_DRAWN).RefersToRange
End Function


Private Function InputCell( _
    ByVal Index As Long, _
    ByVal InputNo As Long) _
    As Range
'
'==============================================================================
'                                  InputCell
'------------------------------------------------------------------------------
' PURPOSE
'   The cell of input InputNo of test Index.
'
' UPDATED
'   2026-11-19
'==============================================================================
'
        Set InputCell = TestsRange().Cells(Index, TS_FIRST_INPUT + InputNo - 1)

End Function


Public Function FormulaTestSheetRow( _
    ByVal Index As Long) _
    As Long
'
'==============================================================================
'                             FormulaTestSheetRow
'------------------------------------------------------------------------------
' PURPOSE
'   The worksheet row of test Index, counted from 1.
'
' UPDATED
'   2026-11-19
'==============================================================================
'
        FormulaTestSheetRow = TestsRange().Row + Index - 1

End Function


Public Function FormulaTestFamilies( _
    ByRef FirstIndex() As Long) _
    As Variant
'
'==============================================================================
'                             FormulaTestFamilies
'------------------------------------------------------------------------------
' PURPOSE
'   The distinct range sets of the defaults table, in its order, as a
'   zero-based array of names; FirstIndex(n) is the first test, counted
'   from 1, that uses the range set n.
'
' UPDATED
'   2026-11-19
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Names()     As Variant     'The range sets found
    Dim Found       As Long        'How many
    Dim Index       As Long        'Test
    Dim N           As Long        'Range set already found

'------------------------------------------------------------------------------
' COLLECT
'------------------------------------------------------------------------------
        EnsureTables
        ReDim Names(0 To UBound(mDefaults, 1) - 1)
        ReDim FirstIndex(0 To UBound(mDefaults, 1) - 1)
        For Index = 1 To UBound(mDefaults, 1)
            For N = 0 To Found - 1
                If Names(N) = CStr(mDefaults(Index, DF_FAMILY)) Then Exit For
            Next N
            If N = Found Then
                Names(Found) = CStr(mDefaults(Index, DF_FAMILY))
                FirstIndex(Found) = Index
                Found = Found + 1
            End If
        Next Index
        ReDim Preserve Names(0 To Found - 1)
        ReDim Preserve FirstIndex(0 To Found - 1)
        FormulaTestFamilies = Names

End Function


Private Function RangeRow( _
    ByVal Family As String, _
    ByVal InputNo As Long) _
    As Long
'
'==============================================================================
'                                  RangeRow
'------------------------------------------------------------------------------
' PURPOSE
'   The row of the range table for an input of a range set; 0 for none.
'
' UPDATED
'   2026-11-19
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim R   As Long     'Row of the range table

'------------------------------------------------------------------------------
' FIND
'------------------------------------------------------------------------------
        EnsureTables
        For R = 1 To UBound(mRanges, 1)
            If CStr(mRanges(R, RG_FAMILY)) = Family And mRanges(R, RG_INPUT) = InputNo Then
                RangeRow = R
                Exit Function
            End If
        Next R

End Function


'
'------------------------------------------------------------------------------
'
'                            VALIDATING THE RANGES
'
'------------------------------------------------------------------------------
'

Public Function FormulaTestRangeProblem() As String
'
'==============================================================================
'                           FormulaTestRangeProblem
'------------------------------------------------------------------------------
' PURPOSE
'   Tell why the range table cannot give valid draws: a row that is not
'   a number in the right kind, Min above Max, a limit outside the
'   function's domain (#42), a test input with no range, or a range set
'   whose joint constraint no draw can meet.
'
' RETURNS
'   The first problem found, naming the range set and the input; "" when
'   the table is usable.
'
' UPDATED
'   2026-11-19
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim R           As Long        'Row of the range table
    Dim Index       As Long        'Test
    Dim K           As Long        'Input
    Dim Problem     As String      'Problem of one row
    Dim Family      As String      'Range set

'------------------------------------------------------------------------------
' EVERY ROW
'------------------------------------------------------------------------------
        EnsureTables
        For R = 1 To UBound(mRanges, 1)
            Problem = RowProblem(R)
            If Len(Problem) > 0 Then
                FormulaTestRangeProblem = Problem
                Exit Function
            End If
        Next R

'------------------------------------------------------------------------------
' EVERY INPUT OF EVERY TEST HAS A RANGE
'------------------------------------------------------------------------------
        For Index = 1 To UBound(mDefaults, 1)
            Family = CStr(mDefaults(Index, DF_FAMILY))
            For K = 1 To INPUT_COUNT
                If Not IsEmpty(mDefaults(Index, DF_FIRST_INPUT + K - 1)) And RangeRow(Family, K) = 0 Then
                    FormulaTestRangeProblem = "Range set " & Family & " has no range for input " & K & "."
                    Exit Function
                End If
            Next K
        Next Index

'------------------------------------------------------------------------------
' JOINT CONSTRAINTS CAN BE MET
'------------------------------------------------------------------------------
        FormulaTestRangeProblem = JointProblem()

End Function


Private Function RowProblem( _
    ByVal R As Long) _
    As String
'
'==============================================================================
'                                  RowProblem
'------------------------------------------------------------------------------
' PURPOSE
'   The problem of one row of the range table, as a sentence that names
'   it; "" when the row is valid.
'
' UPDATED
'   2026-11-19
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Where       As String      'Range set and input
    Dim Kind        As String      'Kind of the input
    Dim Lo          As Variant     'Min
    Dim Hi          As Variant     'Max

'------------------------------------------------------------------------------
' CHECK
'------------------------------------------------------------------------------
        Where = "Range set " & CStr(mRanges(R, RG_FAMILY)) & ", input " & CStr(mRanges(R, RG_INPUT)) & " (" & _
                CStr(mRanges(R, RG_ARGUMENT)) & "): "
        Kind = CStr(mRanges(R, RG_KIND))
        Lo = mRanges(R, RG_MIN)
        Hi = mRanges(R, RG_MAX)
        Select Case Kind
            Case "boolean"
                If VarType(Lo) <> vbBoolean Or VarType(Hi) <> vbBoolean Then
                    RowProblem = Where & "Min and Max must be TRUE or FALSE."
                ElseIf Lo And Not Hi Then
                    RowProblem = Where & "Min is TRUE and Max is FALSE."
                End If
            Case "number", "integer"
                RowProblem = NumberProblem(R, Where, Kind, Lo, Hi)
            Case Else
                RowProblem = Where & "the kind must be number, integer or boolean."
        End Select

End Function


Private Function NumberProblem( _
    ByVal R As Long, _
    ByVal Where As String, _
    ByVal Kind As String, _
    ByVal Lo As Variant, _
    ByVal Hi As Variant) _
    As String
'
'==============================================================================
'                                NumberProblem
'------------------------------------------------------------------------------
' PURPOSE
'   The problem of a number or integer row: Min and Max numbers with Min
'   not above Max, whole for an integer, and inside the domain.
'
' UPDATED
'   2026-11-19
'==============================================================================
'
        If Not IsPlainNumber(Lo) Or Not IsPlainNumber(Hi) Then
            NumberProblem = Where & "Min and Max must be numbers."
        ElseIf Lo > Hi Then
            NumberProblem = Where & "Min " & NumText(Lo) & " is above Max " & NumText(Hi) & "."
        ElseIf Kind = "integer" And (Lo <> Int(Lo) Or Hi <> Int(Hi)) Then
            NumberProblem = Where & "Min and Max must be whole numbers."
        ElseIf IsPlainNumber(mRanges(R, RG_DOMAIN_MIN)) And Lo < mRanges(R, RG_DOMAIN_MIN) Then
            NumberProblem = Where & "Min " & NumText(Lo) & " is below the domain minimum " & NumText(mRanges(R, RG_DOMAIN_MIN)) & "."
        ElseIf IsPlainNumber(mRanges(R, RG_DOMAIN_MAX)) And Hi > mRanges(R, RG_DOMAIN_MAX) Then
            NumberProblem = Where & "Max " & NumText(Hi) & " is above the domain maximum " & NumText(mRanges(R, RG_DOMAIN_MAX)) & "."
        Else
            NumberProblem = BoundaryProblem(R, Where, RG_BOUNDARY_1)
            If Len(NumberProblem) = 0 Then NumberProblem = BoundaryProblem(R, Where, RG_BOUNDARY_2)
        End If

End Function


Private Function BoundaryProblem( _
    ByVal R As Long, _
    ByVal Where As String, _
    ByVal Column As Long) _
    As String
'
'==============================================================================
'                               BoundaryProblem
'------------------------------------------------------------------------------
' PURPOSE
'   The problem of a boundary: a number inside the domain, or empty.
'
' UPDATED
'   2026-11-19
'==============================================================================
'
    Dim Value   As Variant     'The boundary

        Value = mRanges(R, Column)
        If IsEmpty(Value) Then
            Exit Function
        End If
        If Not IsPlainNumber(Value) Then
            BoundaryProblem = Where & "a boundary must be a number."
        ElseIf IsPlainNumber(mRanges(R, RG_DOMAIN_MIN)) And Value < mRanges(R, RG_DOMAIN_MIN) Then
            BoundaryProblem = Where & "the boundary " & NumText(Value) & " is below the domain minimum."
        ElseIf IsPlainNumber(mRanges(R, RG_DOMAIN_MAX)) And Value > mRanges(R, RG_DOMAIN_MAX) Then
            BoundaryProblem = Where & "the boundary " & NumText(Value) & " is above the domain maximum."
        End If

End Function


Private Function JointProblem() As String
'
'==============================================================================
'                                JointProblem
'------------------------------------------------------------------------------
' PURPOSE
'   Refuse ranges that no draw can satisfy together, instead of
'   searching for one: E must be able to follow S and 0 (SD), P and K
'   above minus lambda (option delta), A below D (CDO).
'
' UPDATED
'   2026-11-19
'==============================================================================
'
        JointProblem = FirstProblem(JointSd("SD"), JointSd("SDFLOOR"))
        If Len(JointProblem) = 0 Then JointProblem = FirstProblem(JointDelta("DELTA", False), JointDelta("DELTALAMBDA", True))
        If Len(JointProblem) = 0 Then JointProblem = JointCdo("CDO")

End Function


Private Function FirstProblem( _
    ByVal A As String, _
    ByVal B As String) _
    As String
        If Len(A) > 0 Then FirstProblem = A Else FirstProblem = B
End Function


Private Function Limit( _
    ByVal Family As String, _
    ByVal InputNo As Long, _
    ByVal Column As Long, _
    ByRef Value As Double) _
    As Boolean
'
'==============================================================================
'                                    Limit
'------------------------------------------------------------------------------
' PURPOSE
'   The Min or Max of an input of a range set; False when the range set
'   does not have the input.
'
' UPDATED
'   2026-11-19
'==============================================================================
'
    Dim R   As Long     'Row of the range table

        R = RangeRow(Family, InputNo)
        If R > 0 Then
            Value = CDbl(mRanges(R, Column))
            Limit = True
        End If

End Function


Private Function JointSd( _
    ByVal Family As String) _
    As String
    Dim SMin    As Double
    Dim EMax    As Double

        If Limit(Family, 1, RG_MIN, SMin) And Limit(Family, 2, RG_MAX, EMax) Then
            If EMax <= IIf(SMin > 0#, SMin, 0#) Then
                JointSd = "Range set " & Family & ": E can never be after S and after 0."
            End If
        End If

End Function


Private Function JointDelta( _
    ByVal Family As String, _
    ByVal WithLambda As Boolean) _
    As String
    Dim PMax    As Double
    Dim KMax    As Double
    Dim LMax    As Double

        If Limit(Family, 1, RG_MAX, PMax) And Limit(Family, 2, RG_MAX, KMax) Then
            If WithLambda Then
                If Not Limit(Family, 7, RG_MAX, LMax) Then Exit Function
            End If
            If PMax + LMax <= 0# Or KMax + LMax <= 0# Then
                JointDelta = "Range set " & Family & ": P and K can never be above minus lambda."
            End If
        End If

End Function


Private Function JointCdo( _
    ByVal Family As String) _
    As String
    Dim AMin    As Double
    Dim DMax    As Double

        If Limit(Family, 1, RG_MIN, AMin) And Limit(Family, 2, RG_MAX, DMax) Then
            If AMin >= DMax Then
                JointCdo = "Range set " & Family & ": A can never be below D."
            End If
        End If

End Function


Private Function IsPlainNumber( _
    ByVal Value As Variant) _
    As Boolean
        IsPlainNumber = Not IsError(Value) And Not IsEmpty(Value) And VarType(Value) <> vbString And VarType(Value) <> vbBoolean
        If IsPlainNumber Then IsPlainNumber = IsNumeric(Value)
End Function


Private Function NumText( _
    ByVal Value As Variant) _
    As String
        NumText = Trim$(Str$(CDbl(Value)))
End Function


'
'------------------------------------------------------------------------------
'
'                                  THE DRAW
'
'------------------------------------------------------------------------------
'

Public Function FormulaTestSeedState( _
    ByVal Seed As Long) _
    As Double
'
'==============================================================================
'                             FormulaTestSeedState
'------------------------------------------------------------------------------
' PURPOSE
'   The generator state for a seed, after dropping the first WARMUP
'   values.
'
' UPDATED
'   2026-11-19
'==============================================================================
'
    Dim State   As Double     'Generator state
    Dim I       As Long       'Value dropped

        State = CDbl(Seed)
        For I = 1 To WARMUP
            FormulaTestNext State
        Next I
        FormulaTestSeedState = State

End Function


Public Function FormulaTestNext( _
    ByRef State As Double) _
    As Double
'
'==============================================================================
'                                FormulaTestNext
'------------------------------------------------------------------------------
' PURPOSE
'   The next value of the generator, strictly between 0 and 1.
'
' UPDATED
'   2026-11-19
'==============================================================================
'
        State = State * GEN_MULTIPLIER
        State = State - Int(State / GEN_MODULUS) * GEN_MODULUS
        FormulaTestNext = State / GEN_MODULUS

End Function


Private Function DrawOne( _
    ByVal R As Long, _
    ByRef State As Double) _
    As Variant
'
'==============================================================================
'                                   DrawOne
'------------------------------------------------------------------------------
' PURPOSE
'   One value for the row R of the range table: a number between Min and
'   Max, a whole number from Min to Max, or a boolean.
'
' UPDATED
'   2026-11-19
'==============================================================================
'
    Dim U       As Double     'Uniform value
    Dim Lo      As Double     'Min
    Dim Hi      As Double     'Max

        U = FormulaTestNext(State)
        Select Case CStr(mRanges(R, RG_KIND))
            Case "boolean"
                If mRanges(R, RG_MIN) = mRanges(R, RG_MAX) Then
                    DrawOne = CBool(mRanges(R, RG_MIN))
                Else
                    DrawOne = (U >= 0.5)
                End If
            Case "integer"
                Lo = CDbl(mRanges(R, RG_MIN))
                Hi = CDbl(mRanges(R, RG_MAX))
                DrawOne = CLng(Lo + Int(U * (Hi - Lo + 1#)))
            Case Else
                Lo = CDbl(mRanges(R, RG_MIN))
                Hi = CDbl(mRanges(R, RG_MAX))
                DrawOne = Lo + U * (Hi - Lo)
        End Select

End Function


Public Function FormulaTestDrawInputs( _
    ByVal Family As String, _
    ByRef State As Double, _
    ByRef Inputs() As Variant, _
    ByRef Problem As String) _
    As Boolean
'
'==============================================================================
'                             FormulaTestDrawInputs
'------------------------------------------------------------------------------
' PURPOSE
'   Draw all the inputs of one test of a range set, again while the
'   joint constraints are not met, at most JOINT_TRIES times.
'
' INPUTS
'   Inputs: an array dimensioned 1 To 7, filled with the draw; an input
'   the range set has no range for is left Empty.
'
' RETURNS
'   True when a draw was found; otherwise Problem says so.
'
' UPDATED
'   2026-11-19
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Tries   As Long      'Draws made
    Dim K       As Long      'Input
    Dim R       As Long      'Row of the range table

'------------------------------------------------------------------------------
' DRAW
'------------------------------------------------------------------------------
        EnsureTables
        Do While Tries < JOINT_TRIES
            For K = 1 To INPUT_COUNT
                R = RangeRow(Family, K)
                If R = 0 Then
                    Inputs(K) = Empty
                Else
                    Inputs(K) = DrawOne(R, State)
                End If
            Next K
            If FormulaTestJointOk(Family, Inputs) Then
                FormulaTestDrawInputs = True
                Exit Function
            End If
            Tries = Tries + 1
        Loop
        Problem = "Range set " & Family & ": no draw met the joint constraints in " & JOINT_TRIES & " tries; " & _
                  "widen the ranges."

End Function


Public Function FormulaTestJointOk( _
    ByVal Family As String, _
    ByRef Inputs() As Variant) _
    As Boolean
'
'==============================================================================
'                             FormulaTestJointOk
'------------------------------------------------------------------------------
' PURPOSE
'   Whether a draw meets the joint constraints of its function (#42):
'   the end E after the start S and after 0 (SD), P and K above minus
'   lambda (option delta; lambda is 0 without the seventh input), the
'   attachment A below the detachment D (CDO).
'
' UPDATED
'   2026-11-19
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Lambda  As Double     'Lambda shift

'------------------------------------------------------------------------------
' CHECK
'------------------------------------------------------------------------------
        FormulaTestJointOk = True
        Select Case Family
            Case "SD", "SDFLOOR"
                FormulaTestJointOk = Inputs(2) > Inputs(1) And Inputs(2) > 0#
            Case "DELTA", "DELTALAMBDA"
                If Not IsEmpty(Inputs(7)) Then Lambda = Inputs(7)
                FormulaTestJointOk = Inputs(1) + Lambda > 0# And Inputs(2) + Lambda > 0#
            Case "CDO"
                FormulaTestJointOk = Inputs(1) < Inputs(2)
        End Select

End Function


Public Function FormulaTestBoundaries( _
    ByVal Family As String, _
    ByVal InputNo As Long, _
    ByRef Values() As Double) _
    As Long
'
'==============================================================================
'                             FormulaTestBoundaries
'------------------------------------------------------------------------------
' PURPOSE
'   The branch points the range table lists for an input of a range set.
'
' RETURNS
'   How many there are, 0 to 2; Values(1 To 2) holds them.
'
' UPDATED
'   2026-11-19
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim R   As Long     'Row of the range table

'------------------------------------------------------------------------------
' READ
'------------------------------------------------------------------------------
        ReDim Values(1 To 2)
        R = RangeRow(Family, InputNo)
        If R = 0 Then Exit Function
        If IsPlainNumber(mRanges(R, RG_BOUNDARY_1)) Then
            FormulaTestBoundaries = FormulaTestBoundaries + 1
            Values(FormulaTestBoundaries) = CDbl(mRanges(R, RG_BOUNDARY_1))
        End If
        If IsPlainNumber(mRanges(R, RG_BOUNDARY_2)) Then
            FormulaTestBoundaries = FormulaTestBoundaries + 1
            Values(FormulaTestBoundaries) = CDbl(mRanges(R, RG_BOUNDARY_2))
        End If

End Function


Public Function FormulaTestInDomain( _
    ByVal Family As String, _
    ByVal InputNo As Long, _
    ByVal Value As Double) _
    As Boolean
'
'==============================================================================
'                             FormulaTestInDomain
'------------------------------------------------------------------------------
' PURPOSE
'   Whether a value is inside the domain limits of the range table for an
'   input of a range set (#42); a limit that is empty does not bind.
'
' UPDATED
'   2026-11-19
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim R   As Long     'Row of the range table

'------------------------------------------------------------------------------
' CHECK
'------------------------------------------------------------------------------
        FormulaTestInDomain = True
        R = RangeRow(Family, InputNo)
        If R = 0 Then Exit Function
        If IsPlainNumber(mRanges(R, RG_DOMAIN_MIN)) Then
            If Value < mRanges(R, RG_DOMAIN_MIN) Then FormulaTestInDomain = False
        End If
        If IsPlainNumber(mRanges(R, RG_DOMAIN_MAX)) Then
            If Value > mRanges(R, RG_DOMAIN_MAX) Then FormulaTestInDomain = False
        End If

End Function


Public Sub FormulaTestWriteInputs( _
    ByVal Index As Long, _
    ByRef Inputs() As Variant)
'
'==============================================================================
'                            FormulaTestWriteInputs
'------------------------------------------------------------------------------
' PURPOSE
'   Write the inputs of a draw to the input cells of test Index, as
'   values. An input that is Empty is left as it is.
'
' UPDATED
'   2026-11-19
'==============================================================================
'
    Dim K   As Long     'Input

        For K = 1 To INPUT_COUNT
            If Not IsEmpty(Inputs(K)) Then
                InputCell(Index, K).Value2 = Inputs(K)
            End If
        Next K

End Sub
