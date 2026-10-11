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
'   2026-10-12
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
'   2026-10-07
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim original   As Variant    'Results rows of the demo portfolio in its own order
    Dim irSingle   As Double     'IR add-on of one swap
    Dim crSingle   As Double     'Credit add-on of one credit default swap

'------------------------------------------------------------------------------
' ROW ORDER
'------------------------------------------------------------------------------
        TEST_CaseRunner.BeginSuite CASES, CHECKS
        On Error GoTo Failed

        TEST_CaseRunner.BeginCase "aggregation-demo-original-order", "CRR", VALUATION, "EUR"
        TEST_CaseRunner.LoadSavedInputs
        TEST_CaseRunner.RunCase
        original = ResultsRows()
        TEST_CaseRunner.ExpectTrue "results", RowCount(original) > 0, "no Results rows", "illustrative"

        TEST_CaseRunner.BeginCase "aggregation-demo-reversed-order", "CRR", VALUATION, "EUR"
        TEST_CaseRunner.LoadSavedInputs
        ReorderTrades 0
        TEST_CaseRunner.RunCase
        ExpectSameResults original

        TEST_CaseRunner.BeginCase "aggregation-demo-rotated-order", "CRR", VALUATION, "EUR"
        TEST_CaseRunner.LoadSavedInputs
        ReorderTrades 7
        TEST_CaseRunner.RunCase
        ExpectSameResults original

'------------------------------------------------------------------------------
' OFFSETS
'------------------------------------------------------------------------------
    'Each built case uses five-year trades of notional 10,000 in NS1; a
    'single trade gives the reference add-on.
        StartCase "aggregation-ir-single-swap"
        AddSwap "S1", "EUR", "Long", "Standard"
        TEST_CaseRunner.RunCase
        irSingle = TEST_CaseRunner.OutputNumber("NS1", IR_ADDON)
        TEST_CaseRunner.ExpectTrue "add-on", irSingle > 0#, "single swap add-on is " & irSingle, "illustrative"

        StartCase "aggregation-ir-exact-offset"
        AddSwap "S1", "EUR", "Long", "Standard"
        AddSwap "S2", "EUR", "Short", "Standard"
        TEST_CaseRunner.RunCase
        ExpectNear "add-on", TEST_CaseRunner.OutputNumber("NS1", IR_ADDON), 0#, irSingle

        StartCase "aggregation-ir-no-offset-across-currencies"
        AddSwap "S1", "EUR", "Long", "Standard"
        AddSwap "S2", "USD", "Short", "Standard"
        TEST_CaseRunner.RunCase
        ExpectNear "add-on", TEST_CaseRunner.OutputNumber("NS1", IR_ADDON), 2# * irSingle, irSingle

        StartCase "aggregation-ir-no-offset-with-inflation"
        AddSwap "S1", "EUR", "Long", "Standard"
        AddSwap "S2", "EUR-INFL", "Short", "Standard"
        TEST_CaseRunner.RunCase
        ExpectNear "add-on", TEST_CaseRunner.OutputNumber("NS1", IR_ADDON), 2# * irSingle, irSingle

        StartCase "aggregation-ir-no-offset-across-netting-sets"
        AddSwap "S1", "EUR", "Long", "Standard"
        AddSwap "S2", "EUR", "Short", "Standard"
        TEST_CaseRunner.SetInputCell SH_NS, 2, NS_ID, "NS2"
        TEST_CaseRunner.SetInputCell SH_TRADES, 2, TR_NS, "NS2"
        TEST_CaseRunner.RunCase
        ExpectNear "add-on", TEST_CaseRunner.OutputNumber("NS1", IR_ADDON) + _
                   TEST_CaseRunner.OutputNumber("NS2", IR_ADDON), 2# * irSingle, irSingle

    'A basis hedging set takes half the factor (BasisFactor 0.5) and never
    'offsets the standard one.
        StartCase "aggregation-ir-no-offset-across-hedging-sets"
        AddSwap "S1", "EUR", "Long", "Standard"
        AddSwap "S2", "EUR", "Short", "Basis"
        TEST_CaseRunner.RunCase
        ExpectNear "add-on", TEST_CaseRunner.OutputNumber("NS1", IR_ADDON), 1.5 * irSingle, irSingle

    'Credit (CRE52.61): two entities of opposite sign keep the
    'idiosyncratic part, sqrt(1 - rho^2) * sqrt(2) times one entity's
    'add-on with rho = 0.5, that is sqrt(1.5) times.
        StartCase "aggregation-cr-single-entity"
        AddCredit "C1", "FIRM A", "Long"
        TEST_CaseRunner.RunCase
        crSingle = TEST_CaseRunner.OutputNumber("NS1", CR_ADDON)
        TEST_CaseRunner.ExpectTrue "add-on", crSingle > 0#, "single entity add-on is " & crSingle, "illustrative"

        StartCase "aggregation-cr-no-offset-across-entities"
        AddCredit "C1", "FIRM A", "Long"
        AddCredit "C2", "FIRM B", "Short"
        TEST_CaseRunner.RunCase
        ExpectNear "add-on", TEST_CaseRunner.OutputNumber("NS1", CR_ADDON), Sqr(1.5) * crSingle, crSingle

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
    ByVal caseId As String)
'
'==============================================================================
'                                  StartCase
'------------------------------------------------------------------------------
' PURPOSE
'   Start a CRR case with one unmargined netting set NS1, blank flags and
'   no collateral.
'
' UPDATED
'   2026-10-07
'==============================================================================
'
        TEST_CaseRunner.BeginCase caseId, "CRR", VALUATION, "EUR"
        TEST_CaseRunner.AddNettingSet "NS1", "N", "N", "", "N", "N", "", "0", "0", "0", "0", ""

End Sub


Private Sub AddSwap( _
    ByVal tradeId As String, _
    ByVal currencyCode As String, _
    ByVal direction As String, _
    ByVal nature As String)
'
'==============================================================================
'                                   AddSwap
'------------------------------------------------------------------------------
' PURPOSE
'   Add a five-year interest-rate swap of notional 10,000 to NS1, in the
'   given currency's hedging set; a basis swap gets the label EURIBOR/ESTR.
'
' UPDATED
'   2026-10-07
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim hsLabel   As String    'Basis hedging-set label; blank for a standard swap

'------------------------------------------------------------------------------
' ADD
'------------------------------------------------------------------------------
        If nature = "Basis" Then
            hsLabel = "EURIBOR/ESTR"
        End If
        TEST_CaseRunner.AddTrade tradeId, "IR", "", currencyCode, "Linear", direction, "", nature, hsLabel, _
                                 "10000", "0", "", "2031-09-30", "2031-09-30", "", "", "", "", "", ""

End Sub


Private Sub AddCredit( _
    ByVal tradeId As String, _
    ByVal reference As String, _
    ByVal direction As String)
'
'==============================================================================
'                                  AddCredit
'------------------------------------------------------------------------------
' PURPOSE
'   Add a five-year single-name credit default swap on an AA entity, of
'   notional 10,000, to NS1.
'
' UPDATED
'   2026-10-07
'==============================================================================
'
        TEST_CaseRunner.AddTrade tradeId, "CR", "AA", reference, "Linear", direction, "", "Standard", "", _
                                 "10000", "0", "", "2031-09-30", "2031-09-30", "", "", "", "", "", ""

End Sub


Private Sub ExpectNear( _
    ByVal label As String, _
    ByVal actual As Double, _
    ByVal expected As Double, _
    ByVal magnitude As Double)
'
'==============================================================================
'                                  ExpectNear
'------------------------------------------------------------------------------
' PURPOSE
'   Record that actual equals expected within REL_TOL of magnitude, the
'   size of the quantities compared (a zero expected value has no size of
'   its own).
'
' UPDATED
'   2026-10-07
'==============================================================================
'
        TEST_CaseRunner.ExpectTrue label, Abs(actual - expected) <= REL_TOL * Abs(magnitude), _
                                   "expected " & expected & ", actual " & actual, "illustrative"

End Sub


Private Sub ReorderTrades( _
    ByVal shift As Long)
'
'==============================================================================
'                                ReorderTrades
'------------------------------------------------------------------------------
' PURPOSE
'   Rewrite the trade rows in another order: reversed when shift is 0,
'   otherwise rotated by shift rows. Formulas are moved as they are, so a
'   trade keeps every input.
'
' UPDATED
'   2026-10-12
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ws       As Worksheet    'Trades sheet
    Dim lastR    As Long         'Last used row
    Dim n        As Long         'Trade rows
    Dim block    As Range        'Trade rows, columns A to TR_NCOLS
    Dim src      As Variant      'Formulas in the original order
    Dim dst      As Variant      'Formulas in the new order
    Dim r        As Long         'Row of dst
    Dim from     As Long         'Row of src that goes to r
    Dim c        As Long         'Column

'------------------------------------------------------------------------------
' REORDER
'------------------------------------------------------------------------------
        Set ws = GetSheet(SH_TRADES)
        lastR = UsedLastRow(ws)
        n = lastR - FIRST_DATA_ROW + 1
        If n < 2 Then
            Exit Sub
        End If
        Set block = TableRange(ws, FIRST_DATA_ROW, 1, lastR, TR_NCOLS)
        src = block.Formula
        dst = block.Formula
        For r = 1 To n
            If shift = 0 Then
                from = n - r + 1
            Else
                from = ((r - 1 + shift) Mod n) + 1
            End If
            For c = 1 To TR_NCOLS
                dst(r, c) = src(from, c)
            Next c
        Next r
        block.Formula = dst

End Sub


Private Function ResultsRows() As Variant
'
'==============================================================================
'                                 ResultsRows
'------------------------------------------------------------------------------
' PURPOSE
'   Return the Results rows of the last run, from the first data row to the
'   last used row, columns A to RS_NCOLS; Empty when there are none.
'
' UPDATED
'   2026-10-12
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ws      As Worksheet    'Results sheet
    Dim lastR   As Long         'Last used row

'------------------------------------------------------------------------------
' READ
'------------------------------------------------------------------------------
        Set ws = GetSheet(SH_RESULTS)
        lastR = UsedLastRow(ws)
        If lastR >= FIRST_DATA_ROW Then
            ResultsRows = TableRange(ws, FIRST_DATA_ROW, 1, lastR, RS_NCOLS).Value
        End If

End Function


Private Function RowCount( _
    ByVal outputRows As Variant) _
    As Long
'
'==============================================================================
'                                   RowCount
'------------------------------------------------------------------------------
' PURPOSE
'   Number of rows in a ResultsRows array; 0 when it is Empty.
'
' UPDATED
'   2026-10-07
'==============================================================================
'
        If IsArray(outputRows) Then
            RowCount = UBound(outputRows, 1)
        End If

End Function


Private Sub ExpectSameResults( _
    ByVal original As Variant)
'
'==============================================================================
'                              ExpectSameResults
'------------------------------------------------------------------------------
' PURPOSE
'   Record one check that the Results rows of the last run match original
'   cell by cell: numbers within REL_TOL, other values exactly. Rows are
'   matched by netting-set ID, so the comparison does not rely on the
'   order of the Results rows either.
'
' UPDATED
'   2026-10-07
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim current   As Variant    'Results rows of the last run
    Dim r         As Long       'Row of original
    Dim q         As Long       'Matching row of current
    Dim c         As Long       'Column
    Dim a         As Variant    'Original value
    Dim b         As Variant    'New value
    Dim detail    As String     'First difference; "" when none

'------------------------------------------------------------------------------
' COMPARE
'------------------------------------------------------------------------------
        current = ResultsRows()
        If RowCount(current) <> RowCount(original) Then
            detail = "Results has " & RowCount(current) & " rows, expected " & RowCount(original)
        End If
        For r = 1 To RowCount(original)
            If Len(detail) > 0 Then
                Exit For
            End If
            q = MatchingRow(current, SafeStr(original(r, 1)))
            If q = 0 Then
                detail = "netting set " & SafeStr(original(r, 1)) & " missing"
            Else
                For c = 1 To RS_NCOLS
                    a = original(r, c)
                    b = current(q, c)
                    If Not SameOutput(a, b) Then
                        detail = SafeStr(original(r, 1)) & " column " & c & ": " & SafeStr(a) & _
                                 " became " & SafeStr(b)
                        Exit For
                    End If
                Next c
            End If
        Next r
        TEST_CaseRunner.ExpectTrue "same-results", Len(detail) = 0, detail, "illustrative"

End Sub


Private Function MatchingRow( _
    ByVal outputRows As Variant, _
    ByVal rowId As String) _
    As Long
'
'==============================================================================
'                                 MatchingRow
'------------------------------------------------------------------------------
' PURPOSE
'   Position of the row whose first column is rowId; 0 when absent.
'
' UPDATED
'   2026-10-07
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim r   As Long    'Row being compared

'------------------------------------------------------------------------------
' SEARCH
'------------------------------------------------------------------------------
        For r = 1 To RowCount(outputRows)
            If SafeStr(outputRows(r, 1)) = rowId Then
                MatchingRow = r
                Exit Function
            End If
        Next r

End Function


Private Function SameOutput( _
    ByVal a As Variant, _
    ByVal b As Variant) _
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
'   2026-10-07
'==============================================================================
'
        If IsError(a) Or IsError(b) Then
            SameOutput = IsError(a) And IsError(b)
        ElseIf IsNum(a) And IsNum(b) Then
            SameOutput = Abs(CDbl(a) - CDbl(b)) <= REL_TOL * Max2(1#, Abs(CDbl(a)))
        Else
            SameOutput = (SafeStr(a) = SafeStr(b))
        End If

End Function
