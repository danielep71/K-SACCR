Attribute VB_Name = "TEST_Dashboard"
'==============================================================================
' MODULE: TEST_Dashboard
'------------------------------------------------------------------------------
' PURPOSE
'   Check the Dashboard (#675): after a run, every block equals the figure
'   computed here from Results, TradeCalc, Checks, NettingSets and the two
'   reconciliation sheets, as docs/workbook/DASHBOARD.md defines it, on the
'   demonstration portfolio, on a portfolio of both regimes with a netting
'   set in the money, one out of the money, one holding collateral and an
'   invalid one, and on an empty portfolio; the EAD of each regime equals
'   its TOTAL row on Results. Clear outputs empties every block, a run that
'   fails leaves none, and an edited input shows the sheet out of date.
'
' PUBLIC SURFACE
'   RunDashboardTests is the entry point. Option Private Module keeps it out
'   of the external workbook automation API.
'
' DEPENDENCIES
'   TEST_CaseRunner, which writes the inputs, records the checks and
'   restores the workbook; M_Main for the run, Clear outputs and the test
'   fault; CORE_Dashboard for the block names; CORE_Util and CORE_Config
'   for the sheets, columns and constants. The expected figures are
'   computed here from the sheets, by loops written apart from
'   CORE_Dashboard; the EAD comes from the TOTAL rows the engine writes.
'
' WORKSHEET SAFETY
'   As TEST_CaseRunner: inputs and outputs are rewritten during the run and
'   restored at the end. Use a development workbook.
'
' USAGE
'   Run TEST_Dashboard.RunDashboardTests from the Immediate window.
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
    Option Explicit
    Option Private Module

'------------------------------------------------------------------------------
' MODULE CONSTANTS
'------------------------------------------------------------------------------
        Private Const VALUATION   As String = "2026-09-30"    'Valuation date of the built cases
        Private Const CASES       As Long = 6                 'Cases in a complete run
        Private Const CHECKS      As Long = 64                'Checks in a complete run
        Private Const REL_TOL     As Double = 0.000000001     'Relative tolerance of a figure
        Private Const CLASS_TAG   As String = "illustrative"  'Reference class of every check

'------------------------------------------------------------------------------
' THE SHEETS, READ FOR ONE CHECK
'------------------------------------------------------------------------------
        Private mResults     As Variant    'Results rows, 1-based; Empty when none
        Private mLegs        As Variant    'TradeCalc rows
        Private mChecks      As Variant    'Checks rows
        Private mInputs      As Variant    'NettingSets rows
        Private mExposure    As Double     'Exposure divisor read from Dashboard_Scale
        Private mNotional    As Double     'Notional divisor read from Dashboard_Scale


Public Sub RunDashboardTests()
'
'==============================================================================
'                              RunDashboardTests
'------------------------------------------------------------------------------
' PURPOSE
'   Run every case and report through TEST_CaseRunner.
'
' UPDATED
'   2026-11-23
'==============================================================================
'
    Dim Result   As String    'Result line of a run

        If Not TEST_CaseRunner.BeginSuite("TEST_Dashboard", CASES, CHECKS) Then Exit Sub
        On Error GoTo Failed

    'The demonstration portfolio of the template: both regimes, invalid and
    'incomplete netting sets.
        TEST_CaseRunner.BeginCase "dashboard-demo", "CRR", VALUATION, "EUR"
        TEST_CaseRunner.LoadSavedInputs
        RunAndCheck

    'Both regimes and the path from notional to EAD. CRR: IN holds +600 and
    '-200 (V 400), OUT holds -300 (V -300), COLL is margined and holds 150 of
    'VM against +200 (RC 50); gross positive value 800, after netting 600,
    'RC 450. BCBS: B1 holds +100. BAD has the role MULTI and is invalid.
        TEST_CaseRunner.BeginCase "dashboard-paths", "CRR", VALUATION, "EUR"
        BuildPaths
        RunAndCheck
        TEST_CaseRunner.ExpectTrue "reading-crr", ReadingOf(1) = _
            "Netting removes 25.0% of gross positive value; collateral a further 18.8%.", ReadingOf(1), CLASS_TAG
        TEST_CaseRunner.ExpectTrue "reading-bcbs", ReadingOf(2) = _
            "Netting removes 0.0% of gross positive value; collateral a further 0.0%.", ReadingOf(2), CLASS_TAG
        TEST_CaseRunner.ExpectTrue "path-crr", Near(PathValue(2, 1) * mExposure, 800#) And _
            Near(PathValue(3, 1) * mExposure, 600#) And Near(PathValue(4, 1) * mExposure, 450#), _
            "gross " & PathValue(2, 1) * mExposure & ", netted " & PathValue(3, 1) * mExposure & _
            ", RC " & PathValue(4, 1) * mExposure, CLASS_TAG

    'A netting set with no trade: the run completes with nothing to show.
        TEST_CaseRunner.BeginCase "dashboard-empty", "CRR", VALUATION, "EUR"
        TEST_CaseRunner.AddNettingSet "NS1", "N", "N", "", "N", "N", "", "0", "0", "0", "0", ""
        RunAndCheck

    'Clear outputs empties every block.
        TEST_CaseRunner.BeginCase "dashboard-clear", "CRR", VALUATION, "EUR"
        TEST_CaseRunner.LoadSavedInputs
        Result = M_Main.RunSACCR_Silent()
        M_Main.ClearOutputs
        TEST_CaseRunner.ExpectTrue "empty", BlocksEmpty(), FirstFilled(), CLASS_TAG

    'A run that fails after the Dashboard was withdrawn leaves it empty.
        TEST_CaseRunner.BeginCase "dashboard-failed-run", "CRR", VALUATION, "EUR"
        TEST_CaseRunner.LoadSavedInputs
        Result = M_Main.RunSACCR_Silent()
        M_Main.gTestFault = "operation"
        Result = M_Main.RunSACCR_Silent()
        M_Main.gTestFault = ""
        TEST_CaseRunner.ExpectTrue "failed", Left$(Result, 9) <> "RESULT=OK", Result, CLASS_TAG
        TEST_CaseRunner.ExpectTrue "empty", BlocksEmpty(), FirstFilled(), CLASS_TAG

    'An edited input shows the Dashboard out of date, as Results.
        TEST_CaseRunner.BeginCase "dashboard-stale", "CRR", VALUATION, "EUR"
        TEST_CaseRunner.LoadSavedInputs
        Result = M_Main.RunSACCR_Silent()
        TableCell(GetSheet(SH_TRADES), TableFirstRow(SH_TRADES), TR_MTM).Value = _
            ToDbl(TableCell(GetSheet(SH_TRADES), TableFirstRow(SH_TRADES), TR_MTM).Value, 0#) + 1#
        GetSheet(SH_DASHBOARD).Range("C5").Calculate
        TEST_CaseRunner.ExpectTrue "status", Left$(SafeStr(GetSheet(SH_DASHBOARD).Range("C5").Value), 11) = _
            "OUT OF DATE", SafeStr(GetSheet(SH_DASHBOARD).Range("C5").Value), CLASS_TAG

        M_Main.gTestFault = ""
        TEST_CaseRunner.EndSuite ""
        Exit Sub

Failed:
        M_Main.gTestFault = ""
        TEST_CaseRunner.EndSuite "unexpected error " & Err.Number & ": " & Err.Description

End Sub


'
'------------------------------------------------------------------------------
'
'                                    CASES
'
'------------------------------------------------------------------------------
'

Private Sub BuildPaths()
'
'==============================================================================
'                                  BuildPaths
'------------------------------------------------------------------------------
' PURPOSE
'   The portfolio of the paths case: netting sets IN, OUT and COLL under
'   CRR, B1 under BCBS and BAD, invalid, with five-year EUR swaps.
'
' UPDATED
'   2026-11-23
'==============================================================================
'
        TEST_CaseRunner.AddNettingSet "IN", "N", "N", "", "N", "N", "", "0", "0", "0", "0", ""
        AddSwap "IN-1", "Long", "600"
        AddSwap "IN-2", "Short", "-200"
        AddSet 2, "OUT", "N", "N", RG_CRR, 0#
        AddSwap "OUT-1", "Short", "-300", "OUT"
        AddSet 3, "COLL", "Y", "N", RG_CRR, 150#
        AddSwap "COLL-1", "Long", "200", "COLL"
        AddSet 4, "B1", "N", "N", RG_BCBS, 0#
        AddSwap "B1-1", "Long", "100", "B1"
        AddSet 5, "BAD", "N", "MULTI", RG_CRR, 0#
        AddSwap "BAD-1", "Long", "50", "BAD"

End Sub


Private Sub AddSet( _
    ByVal DataRow As Long, _
    ByVal NettingSetKey As String, _
    ByVal MarginedFlag As String, _
    ByVal RoleCode As String, _
    ByVal Regime As String, _
    ByVal VmAmount As Double)
'
'==============================================================================
'                                    AddSet
'------------------------------------------------------------------------------
' PURPOSE
'   A netting set on a given input row, daily remargining when margined.
'
' UPDATED
'   2026-11-23
'==============================================================================
'
        TEST_CaseRunner.SetInputCell SH_NS, DataRow, NS_ID, NettingSetKey
        TEST_CaseRunner.SetInputCell SH_NS, DataRow, NS_CPTY, "Counterparty " & NettingSetKey
        TEST_CaseRunner.SetInputCell SH_NS, DataRow, NS_MARGINED, MarginedFlag
        TEST_CaseRunner.SetInputCell SH_NS, DataRow, NS_CLEARED, RoleCode
        If MarginedFlag = "Y" Then TEST_CaseRunner.SetInputCell SH_NS, DataRow, NS_FREQ, 1
        TEST_CaseRunner.SetInputCell SH_NS, DataRow, NS_VM, VmAmount
        TEST_CaseRunner.SetInputCell SH_NS, DataRow, NS_NICA, 0
        TEST_CaseRunner.SetInputCell SH_NS, DataRow, NS_TH, 0
        TEST_CaseRunner.SetInputCell SH_NS, DataRow, NS_MTA, 0
        TEST_CaseRunner.SetInputCell SH_NS, DataRow, NS_REGIME, Regime

End Sub


Private Sub AddSwap( _
    ByVal TradeId As String, _
    ByVal Direction As String, _
    ByVal MarketValue As String, _
    Optional ByVal NettingSetKey As String = "")
'
'==============================================================================
'                                   AddSwap
'------------------------------------------------------------------------------
' PURPOSE
'   A five-year EUR swap of notional 10,000 with its market value, moved to
'   another netting set when one is named.
'
' UPDATED
'   2026-11-23
'==============================================================================
'
    Dim Rows   As Long    'Trade rows written
        TEST_CaseRunner.AddTrade TradeId, "IR", "", "EUR", "Linear", Direction, "", "Standard", "", _
                                 "10000", MarketValue, "", "2031-09-30", "2031-09-30", "", "", "", "", "", ""
        If Len(NettingSetKey) > 0 Then
            Rows = LastDataRow(GetSheet(SH_TRADES), TableFirstRow(SH_TRADES), 1) - TableFirstRow(SH_TRADES) + 1
            TEST_CaseRunner.SetInputCell SH_TRADES, Rows, TR_NS, NettingSetKey
        End If

End Sub


Private Sub RunAndCheck()
'
'==============================================================================
'                                 RunAndCheck
'------------------------------------------------------------------------------
' PURPOSE
'   Run through M_Main, as the Ribbon does, then check every block.
'
' UPDATED
'   2026-11-23
'==============================================================================
'
    Dim Result   As String    'Result line of the run
        Result = M_Main.RunSACCR_Silent()
        TEST_CaseRunner.ExpectTrue "run", Left$(Result, 9) = "RESULT=OK", Result, CLASS_TAG
        CheckBlocks

End Sub


'
'------------------------------------------------------------------------------
'
'                                   CHECKS
'
'------------------------------------------------------------------------------
'

Private Sub CheckBlocks()
'
'==============================================================================
'                                 CheckBlocks
'------------------------------------------------------------------------------
' PURPOSE
'   One check per block: its cells against the figures computed here.
'
' UPDATED
'   2026-11-23
'==============================================================================
'
    Dim BlockList   As Variant    'Block names
    Dim Index   As Long       'Block
    Dim Detail  As String     'The first difference
        ReadSheets
        BlockList = CORE_Dashboard.BlockNames()
        For Index = LBound(BlockList) To UBound(BlockList)
            Detail = Compare(CStr(BlockList(Index)))
            TEST_CaseRunner.ExpectTrue CStr(BlockList(Index)), Len(Detail) = 0, Detail, CLASS_TAG
        Next Index

End Sub


Private Function Compare( _
    ByVal BlockName As String) _
    As String
'
'==============================================================================
'                                   Compare
'------------------------------------------------------------------------------
' PURPOSE
'   The first cell of a block that differs from its expected value, or ""
'   when none does.
'
' UPDATED
'   2026-11-23
'==============================================================================
'
    Dim Actual     As Variant    'The block's cells
    Dim Expected   As Variant    'The figures computed here
    Dim R          As Long       'Row
    Dim C          As Long       'Column
        Actual = ThisWorkbook.Names(BlockName).RefersToRange.Value2
        If Not IsArray(Actual) Then
            Dim One(1 To 1, 1 To 1) As Variant
            One(1, 1) = Actual
            Actual = One
        End If
        Expected = ExpectedBlock(BlockName)
        For R = 1 To UBound(Expected, 1)
            For C = 1 To UBound(Expected, 2)
                If Not Same(Actual(R, C), Expected(R, C)) Then
                    Compare = BlockName & "(" & R & "," & C & "): " & Shown(Actual(R, C)) & _
                              ", expected " & Shown(Expected(R, C))
                    Exit Function
                End If
            Next C
        Next R

End Function


Private Function Same( _
    ByVal Actual As Variant, _
    ByVal Expected As Variant) _
    As Boolean
'
'==============================================================================
'                                     Same
'------------------------------------------------------------------------------
' PURPOSE
'   Equal values: a number within REL_TOL, the same text, or both blank.
'   An expected "*" accepts any value that is not blank.
'
' UPDATED
'   2026-11-23
'==============================================================================
'
        If IsEmpty(Expected) Then
            Same = IsEmpty(Actual) Or SafeStr(Actual) = ""
        ElseIf VarType(Expected) = vbString Then
            If Expected = "*" Then
                Same = Not IsEmpty(Actual)
            Else
                Same = (SafeStr(Actual) = Expected)
            End If
        ElseIf IsNumeric(Actual) And Not IsEmpty(Actual) Then
            Same = Near(CDbl(Actual), CDbl(Expected))
        End If

End Function


Private Function Near( _
    ByVal Actual As Double, _
    ByVal Expected As Double) _
    As Boolean
'
'==============================================================================
'                                     Near
'------------------------------------------------------------------------------
' PURPOSE
'   Equal within REL_TOL of the larger of 1 and the expected value.
'
' UPDATED
'   2026-11-23
'==============================================================================
'
        Near = Abs(Actual - Expected) <= REL_TOL * IIf(Abs(Expected) > 1#, Abs(Expected), 1#)

End Function


Private Function Shown( _
    ByVal Value As Variant) _
    As String
'
'==============================================================================
'                                    Shown
'------------------------------------------------------------------------------
' PURPOSE
'   A value for a failure message.
'
' UPDATED
'   2026-11-23
'==============================================================================
'
        If IsEmpty(Value) Then
            Shown = "(blank)"
        Else
            Shown = SafeStr(Value)
        End If

End Function


Private Function BlocksEmpty() As Boolean
'
'==============================================================================
'                                 BlocksEmpty
'------------------------------------------------------------------------------
' PURPOSE
'   Tell whether every cell of every block is blank.
'
' UPDATED
'   2026-11-23
'==============================================================================
'
        BlocksEmpty = (Len(FirstFilled()) = 0)

End Function


Private Function FirstFilled() As String
'
'==============================================================================
'                                 FirstFilled
'------------------------------------------------------------------------------
' PURPOSE
'   The address of the first block cell that is not blank, or "".
'
' UPDATED
'   2026-11-23
'==============================================================================
'
    Dim BlockList   As Variant    'Block names
    Dim Index   As Long       'Block
    Dim Cell    As Range      'A block cell
        BlockList = CORE_Dashboard.BlockNames()
        For Index = LBound(BlockList) To UBound(BlockList)
            For Each Cell In ThisWorkbook.Names(CStr(BlockList(Index))).RefersToRange.Cells
                If Len(SafeStr(Cell.Value2)) > 0 Then
                    FirstFilled = CStr(BlockList(Index)) & " " & Cell.Address(False, False) & " = " & SafeStr(Cell.Value2)
                    Exit Function
                End If
            Next Cell
        Next Index

End Function


Private Function ReadingOf( _
    ByVal Regime As Long) _
    As String
'
'==============================================================================
'                                  ReadingOf
'------------------------------------------------------------------------------
' PURPOSE
'   The reading the Dashboard shows for a regime (1 CRR, 2 BCBS).
'
' UPDATED
'   2026-11-23
'==============================================================================
'
        ReadingOf = SafeStr(ThisWorkbook.Names("Dashboard_Reading").RefersToRange.Cells(Regime, 1).Value2)

End Function


Private Function PathValue( _
    ByVal StepRow As Long, _
    ByVal Regime As Long) _
    As Double
'
'==============================================================================
'                                  PathValue
'------------------------------------------------------------------------------
' PURPOSE
'   An amount of view 11 as the Dashboard shows it, in its unit.
'
' UPDATED
'   2026-11-23
'==============================================================================
'
        PathValue = ToDbl(ThisWorkbook.Names("Dashboard_NotionalToEAD").RefersToRange _
                          .Cells(StepRow, 2 * Regime - 1).Value2, 0#)
        mExposure = ToDbl(ThisWorkbook.Names("Dashboard_Scale").RefersToRange.Cells(1, 1).Value2, 1#)

End Function


'
'------------------------------------------------------------------------------
'
'                              EXPECTED FIGURES
'
'------------------------------------------------------------------------------
'

Private Sub ReadSheets()
'
'==============================================================================
'                                  ReadSheets
'------------------------------------------------------------------------------
' PURPOSE
'   Read the tables the figures come from, and the divisors the Dashboard
'   chose, which the checks of the units confirm separately.
'
' UPDATED
'   2026-11-23
'==============================================================================
'
        mResults = TableRows(SH_RESULTS, RS_NCOLS)
        mLegs = TableRows(SH_TRADECALC, TC_NCOLS)
        mChecks = TableRows(SH_CHECKS, CK_NCOLS)
        mInputs = TableRows(SH_NS, NS_NCOLS)
        mExposure = ToDbl(ThisWorkbook.Names("Dashboard_Scale").RefersToRange.Cells(1, 1).Value2, 1#)
        mNotional = ToDbl(ThisWorkbook.Names("Dashboard_Scale").RefersToRange.Cells(2, 1).Value2, 1#)
        If mExposure = 0# Then mExposure = 1#
        If mNotional = 0# Then mNotional = 1#

End Sub


Private Function TableRows( _
    ByVal SheetName As String, _
    ByVal ColumnCount As Long) _
    As Variant
'
'==============================================================================
'                                  TableRows
'------------------------------------------------------------------------------
' PURPOSE
'   A table's rows down to the last with a first column, 1-based; Empty
'   when there is none.
'
' UPDATED
'   2026-11-23
'==============================================================================
'
    Dim FirstRow   As Long       'First data row
    Dim LastRow    As Long       'Last data row
    Dim Values     As Variant    'The cells
        FirstRow = OutputFirstRow(SheetName)
        LastRow = LastDataRow(GetSheet(SheetName), FirstRow, 1)
        If LastRow < FirstRow Then
            TableRows = Empty
            Exit Function
        End If
        Values = TableRange(GetSheet(SheetName), FirstRow, 1, LastRow, ColumnCount).Value2
        If Not IsArray(Values) Then
            Dim One(1 To 1, 1 To 1) As Variant
            One(1, 1) = Values
            Values = One
        End If
        TableRows = Values

End Function


Private Function Size( _
    ByVal Table As Variant) _
    As Long
'
'==============================================================================
'                                     Size
'------------------------------------------------------------------------------
' PURPOSE
'   Rows of a table read by TableRows; 0 for Empty.
'
' UPDATED
'   2026-11-23
'==============================================================================
'
        If IsArray(Table) Then Size = UBound(Table, 1)

End Function


Private Function RegimeCode( _
    ByVal Regime As Long) _
    As String
'
'==============================================================================
'                                  RegimeCode
'------------------------------------------------------------------------------
' PURPOSE
'   CRR for 1, BCBS for 2.
'
' UPDATED
'   2026-11-23
'==============================================================================
'
        RegimeCode = IIf(Regime = 1, RG_CRR, RG_BCBS)

End Function


Private Function IsSet( _
    ByVal R As Long) _
    As Boolean
'
'==============================================================================
'                                    IsSet
'------------------------------------------------------------------------------
' PURPOSE
'   Tell whether a Results row is a netting set, not a TOTAL row.
'
' UPDATED
'   2026-11-23
'==============================================================================
'
    Dim RowId   As String    'The row's ID
        RowId = UCase$(Trim$(SafeStr(mResults(R, RS_ID))))
        IsSet = Not (RowId = "TOTAL" Or RowId = "TOTAL CRR" Or RowId = "TOTAL BCBS")

End Function


Private Function ValidIn( _
    ByVal R As Long, _
    ByVal Regime As Long) _
    As Boolean
'
'==============================================================================
'                                   ValidIn
'------------------------------------------------------------------------------
' PURPOSE
'   Tell whether a Results row is a VALID netting set of the regime.
'
' UPDATED
'   2026-11-23
'==============================================================================
'
        If Not IsSet(R) Then Exit Function
        ValidIn = (Trim$(SafeStr(mResults(R, RS_STATUS_COL))) = "VALID" And _
                   UCase$(Trim$(SafeStr(mResults(R, RS_REGIME)))) = RegimeCode(Regime))

End Function


Private Function TotalEAD( _
    ByVal Regime As Long) _
    As Double
'
'==============================================================================
'                                   TotalEAD
'------------------------------------------------------------------------------
' PURPOSE
'   The EAD of the regime's TOTAL row on Results; 0 when it has none.
'
' UPDATED
'   2026-11-23
'==============================================================================
'
    Dim R   As Long    'Results row
        For R = 1 To Size(mResults)
            If UCase$(Trim$(SafeStr(mResults(R, RS_ID)))) = "TOTAL " & RegimeCode(Regime) Then
                TotalEAD = ToDbl(mResults(R, RS_EAD), 0#)
            End If
        Next R

End Function


Private Function SetSum( _
    ByVal Regime As Long, _
    ByVal Column As Long, _
    Optional ByVal Test As String = "", _
    Optional ByVal Key As String = "") _
    As Double
'
'==============================================================================
'                                    SetSum
'------------------------------------------------------------------------------
' PURPOSE
'   A Results column summed over the valid netting sets of a regime; the
'   column 0 sums max(V, 0). Test SECTOR, ROLE or MARGINED keeps the sets
'   whose NettingSets sector, role (CCP split by QCCP) or Margined is Key.
'
' UPDATED
'   2026-11-23
'==============================================================================
'
    Dim R   As Long    'Results row
        For R = 1 To Size(mResults)
            If ValidIn(R, Regime) Then
                If Len(Test) = 0 Or SetKey(R, Test) = Key Then
                    If Column = 0 Then
                        If ToDbl(mResults(R, RS_V), 0#) > 0# Then SetSum = SetSum + ToDbl(mResults(R, RS_V), 0#)
                    Else
                        SetSum = SetSum + ToDbl(mResults(R, Column), 0#)
                    End If
                End If
            End If
        Next R

End Function


Private Function SetKey( _
    ByVal R As Long, _
    ByVal Test As String) _
    As String
'
'==============================================================================
'                                    SetKey
'------------------------------------------------------------------------------
' PURPOSE
'   The sector, role or Margined flag of a Results row, the first two
'   looked up on NettingSets by ID.
'
' UPDATED
'   2026-11-23
'==============================================================================
'
    Dim I      As Long      'NettingSets row
    Dim Found  As Long      'The first row of the ID
    Dim Value  As String    'The answer
        If Test = "MARGINED" Then
            SetKey = UCase$(Trim$(SafeStr(mResults(R, RS_MARGINED))))
            Exit Function
        End If
        For I = 1 To Size(mInputs)
            If UCase$(Trim$(SafeStr(mInputs(I, NS_ID)))) = UCase$(Trim$(SafeStr(mResults(R, RS_ID)))) Then
                Found = I
                Exit For
            End If
        Next I
        If Test = "SECTOR" Then
            Value = "NC"
            If Found > 0 Then
                Select Case UCase$(Trim$(SafeStr(mInputs(Found, NS_SECTOR))))
                    Case "CB", "GG", "CI", "OFC", "NFC", "HH"
                        Value = UCase$(Trim$(SafeStr(mInputs(Found, NS_SECTOR))))
                    Case Else
                End Select
            End If
        Else
            Value = "N"
            If Found > 0 Then
                Value = UCase$(Trim$(SafeStr(mInputs(Found, NS_CLEARED))))
                If Value = "" Then Value = "N"
                If Value = "CCP" Then
                    Value = IIf(UCase$(Trim$(SafeStr(mInputs(Found, NS_QCCP)))) = "Y", "CCP_Q", "CCP_NQ")
                End If
            End If
        End If
        SetKey = Value

End Function


Private Function LegTotal( _
    ByVal Regime As Long, _
    ByVal Positive As Boolean, _
    Optional ByVal Band As Long = 0, _
    Optional ByVal Code As String = "", _
    Optional ByVal CountOnly As Boolean = False) _
    As Double
'
'==============================================================================
'                                   LegTotal
'------------------------------------------------------------------------------
' PURPOSE
'   Over the TradeCalc legs that are OK in a valid netting set of the
'   regime: the notional, or with Positive the positive MtM, optionally of
'   one maturity band (1 under a year, 2 one to five, 3 over five) or one
'   asset class; with CountOnly the number of legs.
'
' UPDATED
'   2026-11-23
'==============================================================================
'
    Dim L        As Long      'Leg
    Dim R        As Long      'Results row
    Dim Used     As Boolean   'The leg's netting set is valid in the regime
    Dim M        As Double    'The leg's maturity
    Dim InBand   As Long      'Its band
        For L = 1 To Size(mLegs)
            If UCase$(Trim$(SafeStr(mLegs(L, TC_STATUS)))) = "OK" And _
               UCase$(Trim$(SafeStr(mLegs(L, TC_REGIME)))) = RegimeCode(Regime) Then
                Used = False
                For R = 1 To Size(mResults)
                    If ValidIn(R, Regime) Then
                        If UCase$(Trim$(SafeStr(mResults(R, RS_ID)))) = UCase$(Trim$(SafeStr(mLegs(L, TC_NS)))) Then Used = True
                    End If
                Next R
                M = ToDbl(mLegs(L, TC_M), 0#)
                InBand = IIf(M < 1#, 1, IIf(M <= 5#, 2, 3))
                If Used And (Band = 0 Or InBand = Band) And _
                   (Len(Code) = 0 Or UCase$(Trim$(SafeStr(mLegs(L, TC_AC)))) = Code) Then
                    If CountOnly Then
                        LegTotal = LegTotal + 1#
                    ElseIf Positive Then
                        If ToDbl(mLegs(L, TC_MTM), 0#) > 0# Then LegTotal = LegTotal + ToDbl(mLegs(L, TC_MTM), 0#)
                    Else
                        LegTotal = LegTotal + ToDbl(mLegs(L, TC_NOTIONAL), 0#)
                    End If
                End If
            End If
        Next L

End Function


Private Function ExpectedBlock( _
    ByVal BlockName As String) _
    As Variant
'
'==============================================================================
'                                ExpectedBlock
'------------------------------------------------------------------------------
' PURPOSE
'   The figures of a block, as DASHBOARD.md defines them, computed from the
'   sheets read by ReadSheets.
'
' UPDATED
'   2026-11-23
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Values()   As Variant    'The block
    Dim G          As Long       'Regime position
    Dim K          As Long       'Row
    Dim R          As Long       'Results, legs or Checks row
    Dim Keys       As Variant    'Row keys
    Dim Whole      As Double     'A regime's EAD
    Dim Steps(1 To 6) As Double  'View 11 of a regime

'------------------------------------------------------------------------------
' EACH BLOCK
'------------------------------------------------------------------------------
        Select Case BlockName
            Case "Dashboard_Run"
                ReDim Values(1 To 6, 1 To 1)
                Values(1, 1) = ToDbl(GetParam(PRM_ASOF), 0#)
                Values(2, 1) = "*"
                Values(3, 1) = Trim$(SafeStr(GetParam(PRM_REPCCY)))
                Values(4, 1) = UCase$(Trim$(SafeStr(GetParam(PRM_REGIME))))
                Values(5, 1) = IIf(CORE_Engine.RunClass = RUN_REGULATORY, RUN_REGULATORY & " " & REF_SET_ID, _
                                   CORE_Engine.RunClass)
                Values(6, 1) = CORE_Engine.RunInputs
            Case "Dashboard_Scale"
                ReDim Values(1 To 2, 1 To 2)
                Values(1, 1) = ScaleOf(LargestExposure())
                Values(1, 2) = UnitOf(ScaleOf(LargestExposure()))
                Values(2, 1) = ScaleOf(LargestNotional())
                Values(2, 2) = UnitOf(ScaleOf(LargestNotional()))
            Case "Dashboard_KeyFigures"
                ReDim Values(1 To 12, 1 To 2)
                For G = 1 To 2
                    Values(1, G) = TotalEAD(G) / mExposure
                    Values(2, G) = SetSum(G, RS_RC) / mExposure
                    Values(3, G) = SetSum(G, RS_PFE) / mExposure
                    For K = 4 To 11
                        Values(K, G) = 0
                    Next K
                    For R = 1 To Size(mResults)
                        If IsSet(R) And UCase$(Trim$(SafeStr(mResults(R, RS_REGIME)))) = RegimeCode(G) Then
                            Values(4, G) = Values(4, G) + 1
                            Select Case True
                                Case Trim$(SafeStr(mResults(R, RS_STATUS_COL))) = "VALID"
                                    Values(5, G) = Values(5, G) + 1
                                    If UCase$(Trim$(SafeStr(mResults(R, RS_CAP)))) = "Y" Then Values(11, G) = Values(11, G) + 1
                                Case Left$(SafeStr(mResults(R, RS_STATUS_COL)), 7) = "INVALID"
                                    Values(6, G) = Values(6, G) + 1
                                Case Left$(SafeStr(mResults(R, RS_STATUS_COL)), 10) = "INCOMPLETE"
                                    Values(7, G) = Values(7, G) + 1
                                Case Trim$(SafeStr(mResults(R, RS_STATUS_COL))) = "NO TRADES"
                                    Values(8, G) = Values(8, G) + 1
                                Case Else
                            End Select
                            If UCase$(Trim$(SafeStr(mResults(R, RS_MARGINED)))) = "Y" Then
                                Values(9, G) = Values(9, G) + 1
                            Else
                                Values(10, G) = Values(10, G) + 1
                            End If
                        End If
                    Next R
                    Values(12, G) = DistinctCounterparties(G)
                Next G
            Case "Dashboard_Counts"
                ReDim Values(1 To 8, 1 To 1)
                For K = 1 To 8
                    Values(K, 1) = 0
                Next K
                Values(1, 1) = Size(mLegs)
                For R = 1 To Size(mLegs)
                    If UCase$(Trim$(SafeStr(mLegs(R, TC_STATUS)))) = "OK" Then Values(2, 1) = Values(2, 1) + 1
                    If UCase$(Trim$(SafeStr(mLegs(R, TC_STATUS)))) = "EXCLUDED" Then Values(3, 1) = Values(3, 1) + 1
                Next R
                For R = 1 To Size(mResults)
                    If IsSet(R) Then
                        Select Case UCase$(Trim$(SafeStr(mResults(R, RS_REGIME))))
                            Case RG_CRR, RG_BCBS
                            Case Else
                                Values(4, 1) = Values(4, 1) + 1
                        End Select
                    End If
                Next R
                Values(5, 1) = DistinctCounterparties(0)
                For R = 1 To Size(mChecks)
                    If Trim$(SafeStr(mChecks(R, CK_CODE_COL))) <> "RUN-CLEAN" Then
                        Select Case UCase$(Trim$(SafeStr(mChecks(R, 1))))
                            Case SEV_ERROR
                                Values(6, 1) = Values(6, 1) + 1
                            Case SEV_WARN
                                Values(7, 1) = Values(7, 1) + 1
                            Case SEV_INFO
                                Values(8, 1) = Values(8, 1) + 1
                            Case Else
                        End Select
                    End If
                Next R
            Case "Dashboard_Sector", "Dashboard_ClearingRole"
                If BlockName = "Dashboard_Sector" Then
                    Keys = Array("CB", "GG", "CI", "OFC", "NFC", "HH", "NC")
                Else
                    Keys = Array("N", "CM", "CLIENT", "CCP_Q", "CCP_NQ")
                End If
                ReDim Values(1 To UBound(Keys) + 1, 1 To 2)
                For K = 0 To UBound(Keys)
                    For G = 1 To 2
                        Values(K + 1, G) = SetSum(G, RS_EAD, IIf(BlockName = "Dashboard_Sector", "SECTOR", "ROLE"), _
                                                  CStr(Keys(K))) / mExposure
                    Next G
                Next K
            Case "Dashboard_AddOn"
                ReDim Values(1 To 6, 1 To 2)
                For K = 1 To 6
                    For G = 1 To 2
                        Values(K, G) = SetSum(G, RS_ADDON_IR + K - 1) / mExposure
                    Next G
                Next K
            Case "Dashboard_Composition"
                ReDim Values(1 To 4, 1 To 2)
                For G = 1 To 2
                    Values(1, G) = SetSum(G, RS_RC) / mExposure
                    Values(2, G) = SetSum(G, RS_PFE) / mExposure
                    Values(3, G) = SetSum(G, RS_EAD, "MARGINED", "Y") / mExposure
                    Values(4, G) = SetSum(G, RS_EAD, "MARGINED", "N") / mExposure
                Next G
            Case "Dashboard_TopCRR", "Dashboard_TopBCBS"
                Values = TopExpected(IIf(BlockName = "Dashboard_TopCRR", 1, 2))
            Case "Dashboard_Concentration"
                ReDim Values(1 To 2, 1 To 2)
                For G = 1 To 2
                    Whole = SetSum(G, RS_EAD)
                    If Whole <> 0# Then
                        Values(1, G) = TopShare(G)
                        Values(2, G) = 0#
                        For R = 1 To Size(mResults)
                            If ValidIn(R, G) Then
                                Values(2, G) = Values(2, G) + (ToDbl(mResults(R, RS_EAD), 0#) / Whole) ^ 2
                            End If
                        Next R
                    End If
                Next G
            Case "Dashboard_Maturity"
                ReDim Values(1 To 3, 1 To 2)
                For K = 1 To 3
                    For G = 1 To 2
                        Values(K, G) = LegTotal(G, False, K) / mNotional
                    Next G
                Next K
            Case "Dashboard_AssetClass"
                Keys = Array("IR", "FX", "CR", "EQ", "CO", "OT")
                ReDim Values(1 To 6, 1 To 4)
                For K = 0 To 5
                    For G = 1 To 2
                        Values(K + 1, 2 * G - 1) = LegTotal(G, False, 0, CStr(Keys(K)), True)
                        Values(K + 1, 2 * G) = LegTotal(G, False, 0, CStr(Keys(K))) / mNotional
                    Next G
                Next K
            Case "Dashboard_DataQuality"
                ReDim Values(1 To 2, 1 To 1)
                Values(1, 1) = 0
                Values(2, 1) = 0
                For R = 1 To Size(mLegs)
                    If UCase$(Trim$(SafeStr(mLegs(R, TC_STATUS)))) = "EXCLUDED" Then Values(1, 1) = Values(1, 1) + 1
                Next R
                For R = 1 To Size(mResults)
                    If IsSet(R) And Trim$(SafeStr(mResults(R, RS_STATUS_COL))) <> "VALID" Then Values(2, 1) = Values(2, 1) + 1
                Next R
            Case "Dashboard_Codes"
                Values = CodesExpected()
            Case "Dashboard_Reconciliation"
                Values = ReconExpected()
            Case "Dashboard_NotionalToEAD"
                ReDim Values(1 To 6, 1 To 4)
                For G = 1 To 2
                    Steps(1) = LegTotal(G, False)
                    Steps(2) = LegTotal(G, True)
                    Steps(3) = SetSum(G, 0)
                    Steps(4) = SetSum(G, RS_RC)
                    Steps(5) = SetSum(G, RS_PFE)
                    Steps(6) = TotalEAD(G)
                    For K = 1 To 6
                        Values(K, 2 * G - 1) = Steps(K) / IIf(K = 1, mNotional, mExposure)
                        If Steps(1) <> 0# Then Values(K, 2 * G) = Steps(K) / Steps(1)
                    Next K
                Next G
            Case "Dashboard_Reading"
                ReDim Values(1 To 2, 1 To 1)
                For G = 1 To 2
                    Values(G, 1) = ReadingExpected(G)
                    If IsEmpty(Values(G, 1)) Then Values(G, 1) = Empty
                Next G
            Case Else
                ReDim Values(1 To 1, 1 To 1)
                Values(1, 1) = "unknown block"
        End Select
        ExpectedBlock = Values

End Function


Private Function DistinctCounterparties( _
    ByVal Regime As Long) _
    As Long
'
'==============================================================================
'                            DistinctCounterparties
'------------------------------------------------------------------------------
' PURPOSE
'   Distinct non-blank counterparty names, trimmed and ignoring case, of
'   the netting sets of a regime, or of all of them for 0.
'
' UPDATED
'   2026-11-23
'==============================================================================
'
    Dim Seen   As String    'Names seen, between separators
    Dim R      As Long      'Results row
    Dim Cpty   As String    'A name
        Seen = vbLf
        For R = 1 To Size(mResults)
            If IsSet(R) Then
                If Regime = 0 Or UCase$(Trim$(SafeStr(mResults(R, RS_REGIME)))) = RegimeCode(Regime) Then
                    Cpty = UCase$(Trim$(SafeStr(mResults(R, RS_CPTY))))
                    If Len(Cpty) > 0 And InStr(1, Seen, vbLf & Cpty & vbLf, vbBinaryCompare) = 0 Then
                        Seen = Seen & Cpty & vbLf
                        DistinctCounterparties = DistinctCounterparties + 1
                    End If
                End If
            End If
        Next R

End Function


Private Function ScaleOf( _
    ByVal Largest As Double) _
    As Double
'
'==============================================================================
'                                   ScaleOf
'------------------------------------------------------------------------------
' PURPOSE
'   The divisor of an amount: 1000 to the power of the largest k of 0 to 3
'   with 1000^k not above it.
'
' UPDATED
'   2026-11-23
'==============================================================================
'
    Dim K   As Long    'Power
        ScaleOf = 1#
        For K = 1 To 3
            If Largest >= 1000# ^ K Then ScaleOf = 1000# ^ K
        Next K

End Function


Private Function UnitOf( _
    ByVal Factor As Double) _
    As String
'
'==============================================================================
'                                    UnitOf
'------------------------------------------------------------------------------
' PURPOSE
'   The unit of a divisor in the reporting currency.
'
' UPDATED
'   2026-11-23
'==============================================================================
'
    Dim CcyCode   As String    'Reporting currency
        CcyCode = Trim$(SafeStr(GetParam(PRM_REPCCY)))
        UnitOf = CcyCode & Choose(Application.WorksheetFunction.Max(1, Round(Log(Factor) / Log(1000#)) + 1), _
                                   "", " thousand", " million", " billion")

End Function


Private Function LargestExposure() As Double
'
'==============================================================================
'                               LargestExposure
'------------------------------------------------------------------------------
' PURPOSE
'   The largest absolute exposure amount any block holds, over both
'   regimes, in units.
'
' UPDATED
'   2026-11-23
'==============================================================================
'
    Dim G        As Long        'Regime position
    Dim K        As Long        'Key
    Dim R        As Long        'Results row
    Dim Keys     As Variant     'Sector and role keys
    Dim Values   As Variant     'Candidate amounts
    Dim Item     As Variant     'One amount
        For G = 1 To 2
            Values = Array(TotalEAD(G), SetSum(G, RS_RC), SetSum(G, RS_PFE), SetSum(G, 0), LegTotal(G, True), _
                           SetSum(G, RS_EAD, "MARGINED", "Y"), SetSum(G, RS_EAD, "MARGINED", "N"))
            For Each Item In Values
                If Abs(Item) > LargestExposure Then LargestExposure = Abs(Item)
            Next Item
            For K = 0 To 5
                If Abs(SetSum(G, RS_ADDON_IR + K)) > LargestExposure Then LargestExposure = Abs(SetSum(G, RS_ADDON_IR + K))
            Next K
            Keys = Array("CB", "GG", "CI", "OFC", "NFC", "HH", "NC")
            For K = 0 To UBound(Keys)
                Item = SetSum(G, RS_EAD, "SECTOR", CStr(Keys(K)))
                If Abs(Item) > LargestExposure Then LargestExposure = Abs(Item)
            Next K
            Keys = Array("N", "CM", "CLIENT", "CCP_Q", "CCP_NQ")
            For K = 0 To UBound(Keys)
                Item = SetSum(G, RS_EAD, "ROLE", CStr(Keys(K)))
                If Abs(Item) > LargestExposure Then LargestExposure = Abs(Item)
            Next K
            For R = 1 To Size(mResults)
                If ValidIn(R, G) Then
                    If Abs(ToDbl(mResults(R, RS_EAD), 0#)) > LargestExposure Then LargestExposure = Abs(ToDbl(mResults(R, RS_EAD), 0#))
                End If
            Next R
        Next G

End Function


Private Function LargestNotional() As Double
'
'==============================================================================
'                               LargestNotional
'------------------------------------------------------------------------------
' PURPOSE
'   The largest absolute notional amount any block holds, over both regimes.
'
' UPDATED
'   2026-11-23
'==============================================================================
'
    Dim G       As Long       'Regime position
    Dim K       As Long       'Band or class
    Dim Keys    As Variant    'Asset classes
    Dim Item    As Double     'One amount
        Keys = Array("IR", "FX", "CR", "EQ", "CO", "OT")
        For G = 1 To 2
            Item = LegTotal(G, False)
            If Abs(Item) > LargestNotional Then LargestNotional = Abs(Item)
            For K = 1 To 3
                Item = LegTotal(G, False, K)
                If Abs(Item) > LargestNotional Then LargestNotional = Abs(Item)
            Next K
            For K = 0 To 5
                Item = LegTotal(G, False, 0, CStr(Keys(K)))
                If Abs(Item) > LargestNotional Then LargestNotional = Abs(Item)
            Next K
        Next G

End Function


Private Function TopExpected( _
    ByVal Regime As Long) _
    As Variant
'
'==============================================================================
'                                 TopExpected
'------------------------------------------------------------------------------
' PURPOSE
'   The ten largest valid netting sets of a regime by EAD, ties by ID in
'   binary order, by repeated selection of the largest not yet taken.
'
' UPDATED
'   2026-11-23
'==============================================================================
'
    Dim Values(1 To 10, 1 To 4)   As Variant    'The block
    Dim Taken()                   As Boolean    'Rows ranked already
    Dim Rank                      As Long       'Row of the block
    Dim R                         As Long       'Results row
    Dim Best                      As Long       'The next row
    Dim Whole                     As Double     'The regime's EAD
        Whole = SetSum(Regime, RS_EAD)
        If Size(mResults) > 0 Then ReDim Taken(1 To Size(mResults))
        For Rank = 1 To 10
            Best = 0
            For R = 1 To Size(mResults)
                If ValidIn(R, Regime) Then
                    If Not Taken(R) Then
                        If Best = 0 Then
                            Best = R
                        ElseIf ToDbl(mResults(R, RS_EAD), 0#) > ToDbl(mResults(Best, RS_EAD), 0#) Then
                            Best = R
                        ElseIf ToDbl(mResults(R, RS_EAD), 0#) = ToDbl(mResults(Best, RS_EAD), 0#) And _
                               StrComp(Trim$(SafeStr(mResults(R, RS_ID))), Trim$(SafeStr(mResults(Best, RS_ID))), _
                                       vbBinaryCompare) < 0 Then
                            Best = R
                        End If
                    End If
                End If
            Next R
            If Best = 0 Then Exit For
            Taken(Best) = True
            Values(Rank, 1) = Trim$(SafeStr(mResults(Best, RS_ID)))
            Values(Rank, 2) = Trim$(SafeStr(mResults(Best, RS_CPTY)))
            Values(Rank, 3) = ToDbl(mResults(Best, RS_EAD), 0#) / mExposure
            If Whole <> 0# Then Values(Rank, 4) = ToDbl(mResults(Best, RS_EAD), 0#) / Whole
        Next Rank
        TopExpected = Values

End Function


Private Function TopShare( _
    ByVal Regime As Long) _
    As Double
'
'==============================================================================
'                                   TopShare
'------------------------------------------------------------------------------
' PURPOSE
'   The share of the regime's EAD held by its ten largest netting sets.
'
' UPDATED
'   2026-11-23
'==============================================================================
'
    Dim Top     As Variant    'The ranking
    Dim Rank    As Long       'Row
        Top = TopExpected(Regime)
        For Rank = 1 To 10
            If Not IsEmpty(Top(Rank, 4)) Then TopShare = TopShare + Top(Rank, 4)
        Next Rank

End Function


Private Function CodesExpected() As Variant
'
'==============================================================================
'                                CodesExpected
'------------------------------------------------------------------------------
' PURPOSE
'   The five codes of most ERROR and WARNING rows on Checks; ties put ERROR
'   first, then the code in binary order.
'
' UPDATED
'   2026-11-23
'==============================================================================
'
    Dim Values(1 To 5, 1 To 3)   As Variant    'The block
    Dim Codes                    As String     'Codes ranked already, between separators
    Dim Rank                     As Long       'Row of the block
    Dim R                        As Long       'Checks row
    Dim Code                     As String     'A row's code
    Dim BestCode                 As String     'The next code
    Dim BestRows                 As Long       'Its rows
    Dim BestError                As Boolean    'It has an ERROR row
    Dim Rows                     As Long       'Rows of a code
    Dim HasError                 As Boolean    'A code has an ERROR row
        Codes = vbLf
        For Rank = 1 To 5
            BestCode = ""
            BestRows = 0
            BestError = False
            For R = 1 To Size(mChecks)
                Code = CheckCode(R)
                If Len(Code) > 0 And InStr(1, Codes, vbLf & Code & vbLf, vbBinaryCompare) = 0 Then
                    CountCode Code, Rows, HasError
                    If Rows > BestRows Or (Rows = BestRows And HasError And Not BestError) Or _
                       (Rows = BestRows And HasError = BestError And StrComp(Code, BestCode, vbBinaryCompare) < 0) Then
                        BestCode = Code
                        BestRows = Rows
                        BestError = HasError
                    End If
                End If
            Next R
            If BestRows = 0 Then Exit For
            Codes = Codes & BestCode & vbLf
            Values(Rank, 1) = BestCode
            Values(Rank, 2) = IIf(BestError, SEV_ERROR, SEV_WARN)
            Values(Rank, 3) = BestRows
        Next Rank
        CodesExpected = Values

End Function


Private Function CheckCode( _
    ByVal R As Long) _
    As String
'
'==============================================================================
'                                  CheckCode
'------------------------------------------------------------------------------
' PURPOSE
'   The code of an ERROR or WARNING row of Checks, (none) when blank; ""
'   for another severity.
'
' UPDATED
'   2026-11-23
'==============================================================================
'
        Select Case UCase$(Trim$(SafeStr(mChecks(R, 1))))
            Case SEV_ERROR, SEV_WARN
                CheckCode = Trim$(SafeStr(mChecks(R, CK_CODE_COL)))
                If Len(CheckCode) = 0 Then CheckCode = "(none)"
            Case Else
                CheckCode = ""
        End Select

End Function


Private Sub CountCode( _
    ByVal Code As String, _
    ByRef Rows As Long, _
    ByRef HasError As Boolean)
'
'==============================================================================
'                                  CountCode
'------------------------------------------------------------------------------
' PURPOSE
'   The ERROR and WARNING rows of a code, and whether one is an ERROR.
'
' UPDATED
'   2026-11-23
'==============================================================================
'
    Dim R   As Long    'Checks row
        Rows = 0
        HasError = False
        For R = 1 To Size(mChecks)
            If CheckCode(R) = Code Then
                Rows = Rows + 1
                If UCase$(Trim$(SafeStr(mChecks(R, 1)))) = SEV_ERROR Then HasError = True
            End If
        Next R

End Sub


Private Function ReconExpected() As Variant
'
'==============================================================================
'                                ReconExpected
'------------------------------------------------------------------------------
' PURPOSE
'   For the two reconciliation sheets: Y with rows, else N; then the rows
'   of MATCH, DIFFERENCE, MISSING IN CALCULATOR and any other result.
'
' UPDATED
'   2026-11-23
'==============================================================================
'
    Dim Values(1 To 5, 1 To 2)   As Variant    'The block
    Dim Table                    As Variant    'A sheet's rows
    Dim Column                   As Long       'Sheet
    Dim R                        As Long       'Row
    Dim K                        As Long       'Result row
        For Column = 1 To 2
            Table = TableRows(IIf(Column = 1, SH_RECON, SH_TRADE_RECON), RC_RESULT)
            Values(1, Column) = IIf(Size(Table) > 0, "Y", "N")
            For K = 2 To 5
                Values(K, Column) = 0
            Next K
            For R = 1 To Size(Table)
                Select Case UCase$(Trim$(SafeStr(Table(R, RC_RESULT))))
                    Case RES_MATCH
                        K = 2
                    Case RES_DIFF
                        K = 3
                    Case RES_MISSING
                        K = 4
                    Case Else
                        K = 5
                End Select
                Values(K, Column) = Values(K, Column) + 1
            Next R
        Next Column
        ReconExpected = Values

End Function


Private Function ReadingExpected( _
    ByVal Regime As Long) _
    As Variant
'
'==============================================================================
'                               ReadingExpected
'------------------------------------------------------------------------------
' PURPOSE
'   The reading of view 11 for a regime; Empty without a valid netting set.
'
' UPDATED
'   2026-11-23
'==============================================================================
'
    Dim R            As Long      'Results row
    Dim HasValid     As Boolean   'A valid netting set of the regime
    Dim Gross        As Double    'Gross positive value
    Dim Netted       As Double    'Value after netting
    Dim RC           As Double    'RC
    Dim Collateral   As Double    '(netted - RC) / gross
        For R = 1 To Size(mResults)
            If ValidIn(R, Regime) Then HasValid = True
        Next R
        If Not HasValid Then
            ReadingExpected = Empty
            Exit Function
        End If
        Gross = LegTotal(Regime, True)
        Netted = SetSum(Regime, 0)
        RC = SetSum(Regime, RS_RC)
        If Gross <= 0# Then
            ReadingExpected = "No trade has a positive market value."
            Exit Function
        End If
        Collateral = (Netted - RC) / Gross
        ReadingExpected = "Netting removes " & Percent((Gross - Netted) / Gross) & " of gross positive value; " & _
                          IIf(Collateral >= 0#, "collateral a further " & Percent(Collateral), _
                              "collateral and margin terms add " & Percent(-Collateral)) & "."

End Function


Private Function Percent( _
    ByVal Fraction As Double) _
    As String
'
'==============================================================================
'                                   Percent
'------------------------------------------------------------------------------
' PURPOSE
'   A non-negative fraction as a percentage to one decimal, halves up,
'   with a full stop: the tenths rounded from a Decimal.
'
' UPDATED
'   2026-11-23
'==============================================================================
'
    Dim Tenths   As Long    'The percentage in tenths
        Tenths = CLng(Int(CDec(Fraction) * CDec(1000) + CDec(0.5)))
        Percent = CStr(Tenths \ 10) & "." & CStr(Tenths Mod 10) & "%"

End Function
