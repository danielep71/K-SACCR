Attribute VB_Name = "TEST_Aggregation"
'==============================================================================
' MODULE: TEST_Aggregation
'------------------------------------------------------------------------------
' PURPOSE
'   Check that aggregation does not depend on the order of the trade rows
'   and offsets exactly where the regulation allows and nowhere else
'   (#37): the demo portfolio gives the same Results in reversed and
'   rotated row order; two opposite trades in one bucket offset exactly;
'   there is no offset across currencies, between interest rate and
'   inflation in one currency (#33), across credit entities or netting
'   sets, or between a standard and a basis hedging set.
'
' PUBLIC SURFACE
'   RunAggregationTests is the entry point. Option Private Module keeps it
'   out of the external workbook automation API.
'
' DEPENDENCIES
'   TEST_CaseRunner, which writes the inputs, runs the engine, records the
'   checks and restores the workbook. The offset checks compare outputs
'   with each other, so they need no expected value from outside; they are
'   illustrative. The order checks use the workbook's own inputs, so the
'   workbook must be built from the template.
'
' WORKSHEET SAFETY
'   As TEST_CaseRunner: inputs and outputs are rewritten during the run and
'   restored at the end. Use a development workbook.
'
' USAGE
'   Run TEST_Aggregation.RunAggregationTests from the Immediate window.
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
    Option Explicit
    Option Private Module

'------------------------------------------------------------------------------
' MODULE CONSTANTS
'------------------------------------------------------------------------------
        Private Const VALUATION   As String = "2026-09-30"    'Valuation date of the built cases
        Private Const CASES       As Long = 11                'Cases in a complete run
        Private Const CHECKS      As Long = 11                'Checks in a complete run
        Private Const REL_TOL     As Double = 0.000000001     'Relative tolerance for equal outputs
        Private Const IR_ADDON    As String = "add_on.interest_rate"
        Private Const CR_ADDON    As String = "add_on.credit"


'
'------------------------------------------------------------------------------
'
'                                 ENTRY POINT
'
'------------------------------------------------------------------------------
'

Public Sub RunAggregationTests()
'
'==============================================================================
'                             RunAggregationTests
'------------------------------------------------------------------------------
' PURPOSE
'   Run every case and print the TEST_CaseRunner report.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim OriginalRows    As Variant    'Results rows of the demo portfolio in its own order
    Dim IRSingleAddOn   As Double     'IR add-on of one swap
    Dim CRSingleAddOn   As Double     'Credit add-on of one credit default swap

'------------------------------------------------------------------------------
' ROW ORDER
'------------------------------------------------------------------------------
        TEST_CaseRunner.BeginSuite CASES, CHECKS
        On Error GoTo Failed

        TEST_CaseRunner.BeginCase "aggregation-demo-original-order", "CRR", VALUATION, "EUR"
        TEST_CaseRunner.LoadSavedInputs
        TEST_CaseRunner.RunCase
        OriginalRows = ResultsRows()
        TEST_CaseRunner.ExpectTrue "results", RowCount(OriginalRows) > 0, "no Results rows", "illustrative"

        TEST_CaseRunner.BeginCase "aggregation-demo-reversed-order", "CRR", VALUATION, "EUR"
        TEST_CaseRunner.LoadSavedInputs
        ReorderTrades 0
        TEST_CaseRunner.RunCase
        ExpectSameResults OriginalRows

        TEST_CaseRunner.BeginCase "aggregation-demo-rotated-order", "CRR", VALUATION, "EUR"
        TEST_CaseRunner.LoadSavedInputs
        ReorderTrades 7
        TEST_CaseRunner.RunCase
        ExpectSameResults OriginalRows

'------------------------------------------------------------------------------
' OFFSETS
'------------------------------------------------------------------------------
    'Each built case uses five-year trades of notional 10,000 in NS1; a
    'single trade gives the reference add-on.
        StartCase "aggregation-ir-single-swap"
        AddSwap "S1", "EUR", "Long", "Standard"
        TEST_CaseRunner.RunCase
        IRSingleAddOn = TEST_CaseRunner.OutputNumber("NS1", IR_ADDON)
        TEST_CaseRunner.ExpectTrue "add-on", IRSingleAddOn > 0#, "single swap add-on is " & IRSingleAddOn, "illustrative"

        StartCase "aggregation-ir-exact-offset"
        AddSwap "S1", "EUR", "Long", "Standard"
        AddSwap "S2", "EUR", "Short", "Standard"
        TEST_CaseRunner.RunCase
        ExpectNear "add-on", TEST_CaseRunner.OutputNumber("NS1", IR_ADDON), 0#, IRSingleAddOn

        StartCase "aggregation-ir-no-offset-across-currencies"
        AddSwap "S1", "EUR", "Long", "Standard"
        AddSwap "S2", "USD", "Short", "Standard"
        TEST_CaseRunner.RunCase
        ExpectNear "add-on", TEST_CaseRunner.OutputNumber("NS1", IR_ADDON), 2# * IRSingleAddOn, IRSingleAddOn

        StartCase "aggregation-ir-no-offset-with-inflation"
        AddSwap "S1", "EUR", "Long", "Standard"
        AddSwap "S2", "EUR-INFL", "Short", "Standard"
        TEST_CaseRunner.RunCase
        ExpectNear "add-on", TEST_CaseRunner.OutputNumber("NS1", IR_ADDON), 2# * IRSingleAddOn, IRSingleAddOn

        StartCase "aggregation-ir-no-offset-across-netting-sets"
        AddSwap "S1", "EUR", "Long", "Standard"
        AddSwap "S2", "EUR", "Short", "Standard"
        TEST_CaseRunner.SetInputCell SH_NS, 2, NS_ID, "NS2"
        TEST_CaseRunner.SetInputCell SH_TRADES, 2, TR_NS, "NS2"
        TEST_CaseRunner.RunCase
        ExpectNear "add-on", TEST_CaseRunner.OutputNumber("NS1", IR_ADDON) + _
                   TEST_CaseRunner.OutputNumber("NS2", IR_ADDON), 2# * IRSingleAddOn, IRSingleAddOn

    'A basis hedging set takes half the factor (BasisFactor 0.5) and never
    'offsets the standard one.
        StartCase "aggregation-ir-no-offset-across-hedging-sets"
        AddSwap "S1", "EUR", "Long", "Standard"
        AddSwap "S2", "EUR", "Short", "Basis"
        TEST_CaseRunner.RunCase
        ExpectNear "add-on", TEST_CaseRunner.OutputNumber("NS1", IR_ADDON), 1.5 * IRSingleAddOn, IRSingleAddOn

    'Credit (CRE52.61): two entities of opposite sign keep the
    'idiosyncratic part, sqrt(1 - rho^2) * sqrt(2) times one entity's
    'add-on with rho = 0.5, that is sqrt(1.5) times.
        StartCase "aggregation-cr-single-entity"
        AddCredit "C1", "FIRM A", "Long"
        TEST_CaseRunner.RunCase
        CRSingleAddOn = TEST_CaseRunner.OutputNumber("NS1", CR_ADDON)
        TEST_CaseRunner.ExpectTrue "add-on", CRSingleAddOn > 0#, "single entity add-on is " & CRSingleAddOn, "illustrative"

        StartCase "aggregation-cr-no-offset-across-entities"
        AddCredit "C1", "FIRM A", "Long"
        AddCredit "C2", "FIRM B", "Short"
        TEST_CaseRunner.RunCase
        ExpectNear "add-on", TEST_CaseRunner.OutputNumber("NS1", CR_ADDON), Sqr(1.5) * CRSingleAddOn, CRSingleAddOn

        TEST_CaseRunner.EndSuite ""
        Exit Sub

Failed:
        TEST_CaseRunner.EndSuite "unexpected error " & Err.Number & ": " & Err.Description

End Sub


'
'------------------------------------------------------------------------------
'
'                                   HELPERS
'
'------------------------------------------------------------------------------
'

Private Sub StartCase( _
    ByVal CaseId As String)
'
'==============================================================================
'                                  StartCase
'------------------------------------------------------------------------------
' PURPOSE
'   Start a CRR case with one unmargined netting set NS1, blank flags and
'   no collateral.
'
' UPDATED
'   2026-10-17
'==============================================================================
'
        TEST_CaseRunner.BeginCase CaseId, "CRR", VALUATION, "EUR"
        TEST_CaseRunner.AddNettingSet "NS1", "N", "N", "", "N", "N", "", "0", "0", "0", "0", ""

End Sub


Private Sub AddSwap( _
    ByVal TradeId As String, _
    ByVal CurrencyCode As String, _
    ByVal Direction As String, _
    ByVal Nature As String)
'
'==============================================================================
'                                   AddSwap
'------------------------------------------------------------------------------
' PURPOSE
'   Add a five-year interest-rate swap of notional 10,000 to NS1, in the
'   given currency's hedging set; a basis swap gets the label EURIBOR/ESTR.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim HedgingSetLabel   As String    'Basis hedging-set label; blank for a standard swap

'------------------------------------------------------------------------------
' ADD
'------------------------------------------------------------------------------
        If Nature = "Basis" Then
            HedgingSetLabel = "EURIBOR/ESTR"
        End If
        TEST_CaseRunner.AddTrade TradeId, "IR", "", CurrencyCode, "Linear", Direction, "", Nature, HedgingSetLabel, _
                                 "10000", "0", "", "2031-09-30", "2031-09-30", "", "", "", "", "", ""

End Sub


Private Sub AddCredit( _
    ByVal TradeId As String, _
    ByVal Reference As String, _
    ByVal Direction As String)
'
'==============================================================================
'                                  AddCredit
'------------------------------------------------------------------------------
' PURPOSE
'   Add a five-year single-name credit default swap on an AA entity, of
'   notional 10,000, to NS1.
'
' UPDATED
'   2026-10-17
'==============================================================================
'
        TEST_CaseRunner.AddTrade TradeId, "CR", "AA", Reference, "Linear", Direction, "", "Standard", "", _
                                 "10000", "0", "", "2031-09-30", "2031-09-30", "", "", "", "", "", ""

End Sub


Private Sub ExpectNear( _
    ByVal Label As String, _
    ByVal Actual As Double, _
    ByVal Expected As Double, _
    ByVal Magnitude As Double)
'
'==============================================================================
'                                  ExpectNear
'------------------------------------------------------------------------------
' PURPOSE
'   Record that Actual equals Expected within REL_TOL of Magnitude, the
'   size of the quantities compared (a zero expected value has no size of
'   its own).
'
' UPDATED
'   2026-10-17
'==============================================================================
'
        TEST_CaseRunner.ExpectTrue Label, Abs(Actual - Expected) <= REL_TOL * Abs(Magnitude), _
                                   "expected " & Expected & ", actual " & Actual, "illustrative"

End Sub


Private Sub ReorderTrades( _
    ByVal RotationRows As Long)
'
'==============================================================================
'                                ReorderTrades
'------------------------------------------------------------------------------
' PURPOSE
'   Rewrite the trade rows in another order: reversed when RotationRows is
'   0, otherwise rotated by RotationRows rows. Formulas are moved as they
'   are, so a trade keeps every input.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim TradesSheet      As Worksheet    'Trades sheet
    Dim LastRow          As Long         'Last used row
    Dim TradeRowCount    As Long         'Trade rows
    Dim TradeBlock       As Range        'Trade rows, table columns 1 to TR_NCOLS
    Dim SourceFormulas   As Variant      'Formulas in the original order
    Dim TargetFormulas   As Variant      'Formulas in the new order
    Dim r                As Long         'Row of TargetFormulas
    Dim SourceRow        As Long         'Row of SourceFormulas that goes to r
    Dim c                As Long         'Column
    Dim FirstRow         As Long         'First data row

'------------------------------------------------------------------------------
' REORDER
'------------------------------------------------------------------------------
        Set TradesSheet = GetSheet(SH_TRADES)
        FirstRow = TableFirstRow(SH_TRADES)
        LastRow = UsedLastRow(TradesSheet)
        TradeRowCount = LastRow - FirstRow + 1
        If TradeRowCount < 2 Then
            Exit Sub
        End If
        Set TradeBlock = TableRange(TradesSheet, FirstRow, 1, LastRow, TR_NCOLS)
        SourceFormulas = TradeBlock.Formula
        TargetFormulas = TradeBlock.Formula
        For r = 1 To TradeRowCount
            If RotationRows = 0 Then
                SourceRow = TradeRowCount - r + 1
            Else
                SourceRow = ((r - 1 + RotationRows) Mod TradeRowCount) + 1
            End If
            For c = 1 To TR_NCOLS
                TargetFormulas(r, c) = SourceFormulas(SourceRow, c)
            Next c
        Next r
        TradeBlock.Formula = TargetFormulas

End Sub


Private Function ResultsRows() As Variant
'
'==============================================================================
'                                 ResultsRows
'------------------------------------------------------------------------------
' PURPOSE
'   Return the Results rows of the last run, from the first data row to the
'   last used row, table columns 1 to RS_NCOLS; Empty when there are none.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ResultsSheet   As Worksheet    'Results sheet
    Dim LastRow        As Long         'Last used row

'------------------------------------------------------------------------------
' READ
'------------------------------------------------------------------------------
        Set ResultsSheet = GetSheet(SH_RESULTS)
        LastRow = UsedLastRow(ResultsSheet)
        If LastRow >= OutputFirstRow(SH_RESULTS) Then
            ResultsRows = TableRange(ResultsSheet, OutputFirstRow(SH_RESULTS), 1, LastRow, RS_NCOLS).Value
        End If

End Function


Private Function RowCount( _
    ByVal RowsBlock As Variant) _
    As Long
'
'==============================================================================
'                                   RowCount
'------------------------------------------------------------------------------
' PURPOSE
'   Number of rows in a ResultsRows array; 0 when it is Empty.
'
' UPDATED
'   2026-10-17
'==============================================================================
'
        If IsArray(RowsBlock) Then
            RowCount = UBound(RowsBlock, 1)
        End If

End Function


Private Sub ExpectSameResults( _
    ByVal Original As Variant)
'
'==============================================================================
'                              ExpectSameResults
'------------------------------------------------------------------------------
' PURPOSE
'   Record one check that the Results rows of the last run match Original
'   cell by cell: numbers within REL_TOL, other values exactly. Rows are
'   matched by netting-set ID, so the comparison does not rely on the
'   order of the Results rows either.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim CurrentRows     As Variant    'Results rows of the last run
    Dim r               As Long       'Row of Original
    Dim MatchRow        As Long       'Matching row of CurrentRows
    Dim c               As Long       'Column
    Dim OriginalValue   As Variant    'Original value
    Dim NewValue        As Variant    'New value
    Dim Detail          As String     'First difference; "" when none

'------------------------------------------------------------------------------
' COMPARE
'------------------------------------------------------------------------------
        CurrentRows = ResultsRows()
        If RowCount(CurrentRows) <> RowCount(Original) Then
            Detail = "Results has " & RowCount(CurrentRows) & " rows, expected " & RowCount(Original)
        End If
        For r = 1 To RowCount(Original)
            If Len(Detail) > 0 Then
                Exit For
            End If
            MatchRow = MatchingRow(CurrentRows, SafeStr(Original(r, 1)))
            If MatchRow = 0 Then
                Detail = "netting set " & SafeStr(Original(r, 1)) & " missing"
            Else
                For c = 1 To RS_NCOLS
                    OriginalValue = Original(r, c)
                    NewValue = CurrentRows(MatchRow, c)
                    If Not SameOutput(OriginalValue, NewValue) Then
                        Detail = SafeStr(Original(r, 1)) & " column " & c & ": " & SafeStr(OriginalValue) & _
                                 " became " & SafeStr(NewValue)
                        Exit For
                    End If
                Next c
            End If
        Next r
        TEST_CaseRunner.ExpectTrue "same-results", Len(Detail) = 0, Detail, "illustrative"

End Sub


Private Function MatchingRow( _
    ByVal RowsBlock As Variant, _
    ByVal RowId As String) _
    As Long
'
'==============================================================================
'                                 MatchingRow
'------------------------------------------------------------------------------
' PURPOSE
'   Position of the row of RowsBlock whose first column is RowId; 0 when
'   absent.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim r   As Long    'Row being compared

'------------------------------------------------------------------------------
' SEARCH
'------------------------------------------------------------------------------
        For r = 1 To RowCount(RowsBlock)
            If SafeStr(RowsBlock(r, 1)) = RowId Then
                MatchingRow = r
                Exit Function
            End If
        Next r

End Function


Private Function SameOutput( _
    ByVal FirstValue As Variant, _
    ByVal SecondValue As Variant) _
    As Boolean
'
'==============================================================================
'                                  SameOutput
'------------------------------------------------------------------------------
' PURPOSE
'   Compare two output cells: numbers within REL_TOL of their size, other
'   values exactly. Sums over trades in another order may differ in the
'   last bits.
'
' UPDATED
'   2026-10-17
'==============================================================================
'
        If IsError(FirstValue) Or IsError(SecondValue) Then
            SameOutput = IsError(FirstValue) And IsError(SecondValue)
        ElseIf IsNum(FirstValue) And IsNum(SecondValue) Then
            SameOutput = Abs(CDbl(FirstValue) - CDbl(SecondValue)) <= REL_TOL * Max2(1#, Abs(CDbl(FirstValue)))
        Else
            SameOutput = (SafeStr(FirstValue) = SafeStr(SecondValue))
        End If

End Function
