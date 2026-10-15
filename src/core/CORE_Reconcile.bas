Attribute VB_Name = "CORE_Reconcile"
'==============================================================================
' MODULE: CORE_Reconcile
'------------------------------------------------------------------------------
' PURPOSE
'   Compare a bank's own SA-CCR figures with the calculator's, one figure
'   at a time (#45): per netting set against Results, and per trade leg
'   against TradeCalc. Rows are matched by ID; netting sets of different
'   regimes are never compared; the bank's amounts are scaled to units;
'   every ID that cannot be compared is named with the reason.
'
'     difference          = calculator - bank
'     relative difference = difference / |bank|
'     MATCH               : |difference| <= absolute tolerance, or
'                           |relative difference| <= relative tolerance
'
'   Ratios and times (the multiplier, M, the supervisory duration, the
'   delta and the maturity factors) are not scaled, and only the relative
'   tolerance applies to them.
'
' PUBLIC SURFACE
'   None outside this VBA project. CompareNettingSets, CompareTrades,
'   HasBankRows, SettingsProblem, HeadersProblem, SummaryText and
'   TradeSummaryText, with the RCN_ count indexes, are Public for M_Main;
'   Option Private Module keeps them off the supported external surface.
'
' DEPENDENCIES
'   CORE_Config for the column layout and the result labels; CORE_Util for
'   text conversion and the Collection index (KeyIndex, KeyAdd).
'
' STATE OWNERSHIP
'   None. Every procedure works on its arguments only and touches no
'   workbook: M_Main reads the sheets and writes the rows.
'
' ERROR POLICY
'   An unusable bank row or figure is a result (INVALID INPUT), not a VBA
'   error. Unexpected errors are raised to the caller.
'
' COMPATIBILITY
'   Excel VBA; no references beyond the defaults.
'
' UPDATED
'   2026-10-15
'
' AUTHOR
'   Daniele Penza
'==============================================================================

'------------------------------------------------------------------------------
' MODULE SETTINGS
'------------------------------------------------------------------------------
    'Require explicit declarations; keep the comparison inside this project.
    Option Explicit
    Option Private Module

'------------------------------------------------------------------------------
' COUNT INDEXES
'------------------------------------------------------------------------------
    'Positions in the counts array that CompareNettingSets and CompareTrades
    'fill. A count that does not apply to one of them stays 0.
        Public Const RCN_SETS        As Long = 1     'Distinct IDs in the bank file
        Public Const RCN_COMPARED    As Long = 2     'Figures compared
        Public Const RCN_MATCH       As Long = 3     'Figures within a tolerance
        Public Const RCN_DIFF        As Long = 4     'Figures outside both tolerances
        Public Const RCN_MISSING     As Long = 5     'Bank IDs not on Results or TradeCalc
        Public Const RCN_NOTINBANK   As Long = 6     'Calculator IDs not in the bank file
        Public Const RCN_DUPLICATE   As Long = 7     'Bank rows of an ID listed more than once
        Public Const RCN_REGIME      As Long = 8     'Netting sets whose regimes differ
        Public Const RCN_WITHHELD    As Long = 9     'Netting sets without a VALID status on Results
        Public Const RCN_INVALID     As Long = 10    'Unusable bank rows and figures
        Public Const RCN_EXCLUDED    As Long = 11    'Trades EXCLUDED on TradeCalc
        Public Const RCN_NA          As Long = 12    'Bank figures with no calculator figure to compare
        Public Const RCN_COUNT       As Long = 12    'Number of counts

'------------------------------------------------------------------------------
' MODULE CONSTANTS
'------------------------------------------------------------------------------
    'What is compared: netting sets with Results, or trade legs with
    'TradeCalc.
        Private Const KIND_NS        As Long = 1     'Netting sets
        Private Const KIND_TR        As Long = 2     'Trade legs
    'The measures in the order of the bank columns from BN_EAD and from
    'BT_NOTIONAL. A measure in RATIOS_ is not scaled and only the relative
    'tolerance applies to it; the lists hold 0-based positions.
        Private Const NS_MEASURES    As String = "EAD|RC|PFE|AddOn aggregate|AddOn IR|AddOn FX|" & _
                                                 "AddOn CR|AddOn EQ|AddOn CO|AddOn OT|Multiplier"
        Private Const NS_RATIOS      As String = "|10|"
        Private Const TR_MEASURES    As String = "Notional|MtM|M|Supervisory duration|Adjusted notional d|" & _
                                                 "Supervisory delta|MF unmargined|MF margined|" & _
                                                 "Effective notional unmargined|Effective notional margined|" & _
                                                 "Stand-alone add-on"
        Private Const TR_RATIOS      As String = "|2|3|5|6|7|"
        Private Const MEASURE_COUNT  As Long = 11    'Measures compared per row, in both tables


'
'------------------------------------------------------------------------------
'
'                                 COMPARISON
'
'------------------------------------------------------------------------------
'

Public Function CompareNettingSets( _
    ByVal bank As Variant, _
    ByVal bankFirstRow As Long, _
    ByVal calc As Variant, _
    ByVal scaleFactor As Double, _
    ByVal absTol As Double, _
    ByVal relTol As Double, _
    ByRef counts() As Long) _
    As Variant
'
'==============================================================================
'                              CompareNettingSets
'------------------------------------------------------------------------------
' PURPOSE
'   Compare the bank's netting-set figures with the Results rows and
'   return the Reconciliation rows: first the bank rows in their order, one
'   row per figure compared or per reason a netting set is not compared,
'   then one row for each Results netting set that is not in the bank file.
'
' INPUTS
'   bank: the BankNettingSets rows, (row, 1 To BN_NCOLS), as Value2; Empty
'   when there are none. Rows with no value before the comment are skipped.
'   bankFirstRow: the sheet row of the first bank row, for the Bank row
'   column.
'   calc: the Results rows, (row, 1 To RS_NCOLS); Empty when there are
'   none. The reserved TOTAL rows are skipped.
'   scaleFactor: units per bank amount, greater than zero.
'   absTol, relTol: the tolerances, zero or more.
'   counts: receives the counts, (1 To RCN_COUNT).
'
' RETURNS
'   The rows, (row, 1 To RC_NCOLS); Empty when there are none.
'
' UPDATED
'   2026-10-15
'==============================================================================
'
        CompareNettingSets = CompareTable(KIND_NS, bank, bankFirstRow, calc, scaleFactor, absTol, relTol, counts)

End Function


Public Function CompareTrades( _
    ByVal bank As Variant, _
    ByVal bankFirstRow As Long, _
    ByVal calc As Variant, _
    ByVal scaleFactor As Double, _
    ByVal absTol As Double, _
    ByVal relTol As Double, _
    ByRef counts() As Long) _
    As Variant
'
'==============================================================================
'                                CompareTrades
'------------------------------------------------------------------------------
' PURPOSE
'   Compare the bank's trade-leg figures with the TradeCalc rows and
'   return the TradeReconciliation rows, in the layout of
'   CompareNettingSets: the trade's netting set and asset class stand in
'   the two regime columns (TRC_NS, TRC_AC).
'
' INPUTS
'   bank: the BankTrades rows, (row, 1 To BT_NCOLS), as Value2; Empty when
'   there are none.
'   calc: the TradeCalc rows, (row, 1 To TC_NCOLS); Empty when there are
'   none.
'   bankFirstRow, scaleFactor, absTol, relTol, counts: as
'   CompareNettingSets.
'
' RETURNS
'   The rows, (row, 1 To RC_NCOLS); Empty when there are none.
'
' UPDATED
'   2026-10-15
'==============================================================================
'
        CompareTrades = CompareTable(KIND_TR, bank, bankFirstRow, calc, scaleFactor, absTol, relTol, counts)

End Function


Public Function HasBankRows( _
    ByVal bank As Variant, _
    Optional ByVal commentCol As Long = BN_COMMENT) _
    As Boolean
'
'==============================================================================
'                                 HasBankRows
'------------------------------------------------------------------------------
' PURPOSE
'   Tell whether the bank rows hold any value before the comment column, so
'   that an empty bank table means no reconciliation rather than every
'   calculator row NOT IN BANK FILE.
'
' INPUTS
'   bank: as CompareNettingSets or CompareTrades.
'   commentCol: the comment column: BN_COMMENT or BT_COMMENT.
'
' UPDATED
'   2026-10-15
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim r   As Long    'Bank row

'------------------------------------------------------------------------------
' TEST
'------------------------------------------------------------------------------
        For r = 1 To RowCount(bank)
            If BankRowHasData(bank, r, commentCol) Then
                HasBankRows = True
                Exit Function
            End If
        Next r

End Function


Private Function CompareTable( _
    ByVal kind As Long, _
    ByRef bank As Variant, _
    ByVal bankFirstRow As Long, _
    ByRef calc As Variant, _
    ByVal scaleFactor As Double, _
    ByVal absTol As Double, _
    ByVal relTol As Double, _
    ByRef counts() As Long) _
    As Variant
'
'==============================================================================
'                                 CompareTable
'------------------------------------------------------------------------------
' PURPOSE
'   The comparison behind CompareNettingSets and CompareTrades: index the
'   calculator rows by ID, count the bank IDs, compare each bank row
'   (CompareNettingSetRow or CompareTradeRow), then list the calculator
'   rows that are not in the bank file.
'
' INPUTS
'   kind: KIND_NS or KIND_TR; the other arguments as CompareNettingSets.
'
' UPDATED
'   2026-10-15
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim calcIndex    As Collection    'Calculator ID to calculator row
    Dim calcUsed()   As Boolean       'The calculator row is in the bank file
    Dim bankIndex    As Collection    'Bank ID to position in bankSeen
    Dim bankSeen()   As Long          'Bank rows of each distinct ID
    Dim commentCol   As Long          'Comment column of the bank table
    Dim nCalc        As Long          'Calculator rows
    Dim nBank        As Long          'Bank rows
    Dim outArr()     As Variant       'Rows being built
    Dim nOut         As Long          'Rows filled
    Dim r            As Long          'Bank or calculator row
    Dim k            As Long          'Position of an ID
    Dim id           As String        'ID, upper case
    Dim result       As Variant       'The rows returned

'------------------------------------------------------------------------------
' INDEX THE CALCULATOR ROWS
'------------------------------------------------------------------------------
        ReDim counts(1 To RCN_COUNT)
        commentCol = IIf(kind = KIND_NS, BN_COMMENT, BT_COMMENT)
        nCalc = RowCount(calc)
        nBank = RowCount(bank)
        Set calcIndex = New Collection
        If nCalc > 0 Then
            ReDim calcUsed(1 To nCalc)
        End If
        For r = 1 To nCalc
            id = CalcId(kind, calc, r)
            If Len(id) > 0 Then
                If KeyIndex(calcIndex, id) = 0 Then
                    KeyAdd calcIndex, id, r
                End If
            End If
        Next r

'------------------------------------------------------------------------------
' COUNT THE BANK IDS
'------------------------------------------------------------------------------
    'An ID listed twice is not compared at all: which row is the bank's
    'figure cannot be told.
        Set bankIndex = New Collection
        If nBank > 0 Then
            ReDim bankSeen(1 To nBank)
        End If
        For r = 1 To nBank
            If BankRowHasData(bank, r, commentCol) Then
                id = UTxt(bank(r, 1))
                If Len(id) > 0 Then
                    k = KeyIndex(bankIndex, id)
                    If k = 0 Then
                        k = bankIndex.Count + 1
                        KeyAdd bankIndex, id, k
                    End If
                    bankSeen(k) = bankSeen(k) + 1
                End If
            End If
        Next r
        counts(RCN_SETS) = bankIndex.Count

'------------------------------------------------------------------------------
' COMPARE THE BANK ROWS
'------------------------------------------------------------------------------
    'At most one row per measure for each bank row, plus one row for each
    'calculator row.
        ReDim outArr(1 To nBank * MEASURE_COUNT + nCalc + 1, 1 To RC_NCOLS)
        For r = 1 To nBank
            If BankRowHasData(bank, r, commentCol) Then
                id = UTxt(bank(r, 1))
                If Len(id) = 0 Then
                    AddRow outArr, nOut, "", bankFirstRow + r - 1, _
                           IIf(kind = KIND_NS, UTxt(bank(r, BN_REGIME)), ""), "", "", "", _
                           Empty, Empty, RES_INVALID, IIf(kind = KIND_NS, "NettingSetID", "TradeID") & _
                           " missing: the row is not compared."
                    counts(RCN_INVALID) = counts(RCN_INVALID) + 1
                Else
                    k = KeyIndex(calcIndex, id)
                    If k > 0 Then
                        calcUsed(k) = True
                    End If
                    If kind = KIND_NS Then
                        CompareNettingSetRow bank, r, bankFirstRow + r - 1, calc, k, _
                                             bankSeen(KeyIndex(bankIndex, id)), _
                                             scaleFactor, absTol, relTol, outArr, nOut, counts
                    Else
                        CompareTradeRow bank, r, bankFirstRow + r - 1, calc, k, _
                                        bankSeen(KeyIndex(bankIndex, id)), _
                                        scaleFactor, absTol, relTol, outArr, nOut, counts
                    End If
                End If
            End If
        Next r

'------------------------------------------------------------------------------
' LIST THE CALCULATOR ROWS NOT IN THE BANK FILE
'------------------------------------------------------------------------------
    'With the EAD of a netting set, or the notional of a trade, for scale.
        For r = 1 To nCalc
            id = CalcId(kind, calc, r)
            If Len(id) > 0 Then
                If KeyIndex(calcIndex, id) = r And Not calcUsed(r) Then
                    If kind = KIND_NS Then
                        AddRow outArr, nOut, id, Empty, "", UTxt(calc(r, RS_REGIME)), _
                               SafeStr(calc(r, RS_STATUS_COL)), "EAD", Empty, NumberOrEmpty(calc(r, RS_EAD)), _
                               RES_NOTINBANK, "No row for this netting set on BankNettingSets."
                    Else
                        AddRow outArr, nOut, SafeStr(calc(r, TC_ID)), Empty, SafeStr(calc(r, TC_NS)), _
                               SafeStr(calc(r, TC_AC)), SafeStr(calc(r, TC_STATUS)), "Notional", Empty, _
                               NumberOrEmpty(calc(r, TC_NOTIONAL)), RES_NOTINBANK, _
                               "No row for this trade leg on BankTrades."
                    End If
                    counts(RCN_NOTINBANK) = counts(RCN_NOTINBANK) + 1
                End If
            End If
        Next r

'------------------------------------------------------------------------------
' RETURN THE ROWS FILLED
'------------------------------------------------------------------------------
        If nOut > 0 Then
            result = TrimRows(outArr, nOut)
        End If
        CompareTable = result

End Function


Private Sub CompareNettingSetRow( _
    ByRef bank As Variant, _
    ByVal r As Long, _
    ByVal sheetRow As Long, _
    ByRef calc As Variant, _
    ByVal calcRow As Long, _
    ByVal timesListed As Long, _
    ByVal scaleFactor As Double, _
    ByVal absTol As Double, _
    ByVal relTol As Double, _
    ByRef outArr() As Variant, _
    ByRef nOut As Long, _
    ByRef counts() As Long)
'
'==============================================================================
'                             CompareNettingSetRow
'------------------------------------------------------------------------------
' PURPOSE
'   Add the Reconciliation rows of one bank row with an ID: one row saying
'   why the netting set is not compared, or one row per figure entered.
'
' INPUTS
'   bank, r: the bank rows and this row's position; sheetRow: its sheet row.
'   calc, calcRow: the Results rows and the matching row; 0 when none.
'   timesListed: rows of this ID in the bank file.
'   scaleFactor, absTol, relTol: as CompareNettingSets.
'   outArr, nOut, counts: the rows and counts being built.
'
' UPDATED
'   2026-10-15
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim id            As String     'Netting-set ID, upper case
    Dim bankRegime    As String     'Regime given by the bank
    Dim calcRegime    As String     'Regime on Results
    Dim calcStatus    As String     'Status on Results

'------------------------------------------------------------------------------
' NETTING-SET CHECKS
'------------------------------------------------------------------------------
    'One reason per netting set, in this order: a figure of an ID listed
    'twice, of another regime or of a withheld netting set means nothing.
        id = UTxt(bank(r, BN_ID))
        bankRegime = UTxt(bank(r, BN_REGIME))
        If calcRow > 0 Then
            calcRegime = UTxt(calc(calcRow, RS_REGIME))
            calcStatus = SafeStr(calc(calcRow, RS_STATUS_COL))
        End If
        If timesListed > 1 Then
            AddRow outArr, nOut, id, sheetRow, bankRegime, calcRegime, calcStatus, "", Empty, Empty, _
                   RES_DUPLICATE, "Listed " & timesListed & " times on BankNettingSets: none of its rows is compared."
            counts(RCN_DUPLICATE) = counts(RCN_DUPLICATE) + 1
            Exit Sub
        End If
        If bankRegime <> RG_BCBS And bankRegime <> RG_CRR Then
            AddRow outArr, nOut, id, sheetRow, bankRegime, calcRegime, calcStatus, "", Empty, Empty, _
                   RES_INVALID, "Regime must be BCBS or CRR: the row is not compared."
            counts(RCN_INVALID) = counts(RCN_INVALID) + 1
            Exit Sub
        End If
        If calcRow = 0 Then
            AddRow outArr, nOut, id, sheetRow, bankRegime, "", "", "EAD", BankAmount(bank(r, BN_EAD), scaleFactor), _
                   Empty, RES_MISSING, "No netting set with this ID on Results."
            counts(RCN_MISSING) = counts(RCN_MISSING) + 1
            Exit Sub
        End If
        If bankRegime <> calcRegime Then
            AddRow outArr, nOut, id, sheetRow, bankRegime, calcRegime, calcStatus, "", Empty, Empty, _
                   RES_REGIME, "Bank regime " & bankRegime & ", Results regime " & calcRegime & _
                   ": figures of different regimes are not compared."
            counts(RCN_REGIME) = counts(RCN_REGIME) + 1
            Exit Sub
        End If
        If calcStatus <> "VALID" Then
            AddRow outArr, nOut, id, sheetRow, bankRegime, calcRegime, calcStatus, "", Empty, Empty, _
                   RES_WITHHELD, "No EAD on Results for this netting set: see its status and the Checks sheet."
            counts(RCN_WITHHELD) = counts(RCN_WITHHELD) + 1
            Exit Sub
        End If

'------------------------------------------------------------------------------
' COMPARE EACH FIGURE ENTERED
'------------------------------------------------------------------------------
        CompareFigures KIND_NS, bank, r, BN_EAD, id, sheetRow, bankRegime, calcRegime, calcStatus, _
                       calc, calcRow, scaleFactor, absTol, relTol, outArr, nOut, counts

End Sub


Private Sub CompareTradeRow( _
    ByRef bank As Variant, _
    ByVal r As Long, _
    ByVal sheetRow As Long, _
    ByRef calc As Variant, _
    ByVal calcRow As Long, _
    ByVal timesListed As Long, _
    ByVal scaleFactor As Double, _
    ByVal absTol As Double, _
    ByVal relTol As Double, _
    ByRef outArr() As Variant, _
    ByRef nOut As Long, _
    ByRef counts() As Long)
'
'==============================================================================
'                               CompareTradeRow
'------------------------------------------------------------------------------
' PURPOSE
'   Add the TradeReconciliation rows of one bank row with an ID: one row
'   saying why the trade leg is not compared, or one row per figure
'   entered. A trade carries the regime of its netting set, so there is no
'   regime check here.
'
' INPUTS
'   As CompareNettingSetRow, with the TradeCalc rows as calc.
'
' UPDATED
'   2026-10-15
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim id           As String    'TradeID as entered, trimmed
    Dim calcNs       As String    'Netting set on TradeCalc
    Dim calcAc       As String    'Asset class on TradeCalc
    Dim calcStatus   As String    'OK or EXCLUDED

'------------------------------------------------------------------------------
' TRADE CHECKS
'------------------------------------------------------------------------------
        id = SafeStr(bank(r, BT_ID))
        If calcRow > 0 Then
            calcNs = SafeStr(calc(calcRow, TC_NS))
            calcAc = SafeStr(calc(calcRow, TC_AC))
            calcStatus = SafeStr(calc(calcRow, TC_STATUS))
        End If
        If timesListed > 1 Then
            AddRow outArr, nOut, id, sheetRow, calcNs, calcAc, calcStatus, "", Empty, Empty, _
                   RES_DUPLICATE, "Listed " & timesListed & " times on BankTrades: none of its rows is compared."
            counts(RCN_DUPLICATE) = counts(RCN_DUPLICATE) + 1
            Exit Sub
        End If
        If calcRow = 0 Then
            AddRow outArr, nOut, id, sheetRow, "", "", "", "Notional", _
                   BankAmount(bank(r, BT_NOTIONAL), scaleFactor), Empty, RES_MISSING, _
                   "No trade leg with this ID on TradeCalc."
            counts(RCN_MISSING) = counts(RCN_MISSING) + 1
            Exit Sub
        End If
        If UCase$(calcStatus) = "EXCLUDED" Then
            AddRow outArr, nOut, id, sheetRow, calcNs, calcAc, calcStatus, "", Empty, Empty, _
                   RES_EXCLUDED, "Excluded on TradeCalc: " & Left$(SafeStr(calc(calcRow, TC_MESSAGES)), 250)
            counts(RCN_EXCLUDED) = counts(RCN_EXCLUDED) + 1
            Exit Sub
        End If

'------------------------------------------------------------------------------
' COMPARE EACH FIGURE ENTERED
'------------------------------------------------------------------------------
        CompareFigures KIND_TR, bank, r, BT_NOTIONAL, id, sheetRow, calcNs, calcAc, calcStatus, _
                       calc, calcRow, scaleFactor, absTol, relTol, outArr, nOut, counts

End Sub


Private Sub CompareFigures( _
    ByVal kind As Long, _
    ByRef bank As Variant, _
    ByVal r As Long, _
    ByVal firstCol As Long, _
    ByVal id As String, _
    ByVal sheetRow As Long, _
    ByVal col3 As String, _
    ByVal col4 As String, _
    ByVal calcStatus As String, _
    ByRef calc As Variant, _
    ByVal calcRow As Long, _
    ByVal scaleFactor As Double, _
    ByVal absTol As Double, _
    ByVal relTol As Double, _
    ByRef outArr() As Variant, _
    ByRef nOut As Long, _
    ByRef counts() As Long)
'
'==============================================================================
'                                CompareFigures
'------------------------------------------------------------------------------
' PURPOSE
'   Add one row per figure entered on a bank row: INVALID INPUT when it is
'   not a number, NOT APPLICABLE when the calculator has no such figure for
'   the row, otherwise MATCH or DIFFERENCE. A row with no figure gets one
'   INVALID INPUT row.
'
' INPUTS
'   kind: KIND_NS or KIND_TR; firstCol: the bank column of the first
'   measure; col3, col4: what the third and fourth output columns show
'   (the two regimes, or the trade's netting set and asset class); the
'   other arguments as CompareNettingSetRow.
'
' UPDATED
'   2026-10-15
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim measureNames   As Variant    'Measure names
    Dim m              As Long       'Measure, 0-based
    Dim v              As Variant    'Bank figure as entered
    Dim calcCell       As Variant    'Calculator figure as written
    Dim isRatio        As Boolean    'Not scaled; relative tolerance only
    Dim bankValue      As Double     'Bank figure, in units for an amount
    Dim calcValue      As Double     'Calculator figure
    Dim diff           As Double     'Calculator minus bank
    Dim rel            As Variant    'Relative difference; Empty when the bank figure is 0
    Dim isMatch        As Boolean    'Within a tolerance
    Dim entered        As Long       'Figures entered on the row

'------------------------------------------------------------------------------
' COMPARE
'------------------------------------------------------------------------------
        measureNames = Split(IIf(kind = KIND_NS, NS_MEASURES, TR_MEASURES), "|")
        For m = 0 To MEASURE_COUNT - 1
            v = bank(r, firstCol + m)
            If Not IsBlankBankCell(v) Then
                entered = entered + 1
                calcCell = calc(calcRow, CalcColumn(kind, m))
                isRatio = InStr(IIf(kind = KIND_NS, NS_RATIOS, TR_RATIOS), "|" & m & "|") > 0
                If Not IsNumberValue(v) Then
                    AddRow outArr, nOut, id, sheetRow, col3, col4, calcStatus, CStr(measureNames(m)), _
                           Empty, Empty, RES_INVALID, "Not a number: " & DescribeValue(v) & "."
                    counts(RCN_INVALID) = counts(RCN_INVALID) + 1
                ElseIf Not IsNumberValue(calcCell) Then
                    AddRow outArr, nOut, id, sheetRow, col3, col4, calcStatus, CStr(measureNames(m)), _
                           BankFigure(v, isRatio, scaleFactor), Empty, RES_NA, _
                           "The calculator has no " & measureNames(m) & " for this row."
                    counts(RCN_NA) = counts(RCN_NA) + 1
                Else
                    bankValue = BankFigure(v, isRatio, scaleFactor)
                    calcValue = CDbl(calcCell)
                    diff = calcValue - bankValue
                    If bankValue <> 0# Then
                        rel = diff / Abs(bankValue)
                    Else
                        rel = Empty
                    End If
                    If isRatio Then
                        isMatch = (diff = 0#)
                    Else
                        isMatch = (Abs(diff) <= absTol)
                    End If
                    If Not IsEmpty(rel) Then
                        isMatch = isMatch Or (Abs(rel) <= relTol)
                    End If
                    AddRow outArr, nOut, id, sheetRow, col3, col4, calcStatus, CStr(measureNames(m)), _
                           bankValue, calcValue, IIf(isMatch, RES_MATCH, RES_DIFF), "", diff, rel
                    counts(RCN_COMPARED) = counts(RCN_COMPARED) + 1
                    If isMatch Then
                        counts(RCN_MATCH) = counts(RCN_MATCH) + 1
                    Else
                        counts(RCN_DIFF) = counts(RCN_DIFF) + 1
                    End If
                End If
            End If
        Next m
        If entered = 0 Then
            AddRow outArr, nOut, id, sheetRow, col3, col4, calcStatus, "", Empty, Empty, _
                   RES_INVALID, "No bank figure entered on the row."
            counts(RCN_INVALID) = counts(RCN_INVALID) + 1
        End If

End Sub


'
'------------------------------------------------------------------------------
'
'                             SETTINGS AND HEADERS
'
'------------------------------------------------------------------------------
'

Public Function SettingsProblem( _
    ByVal scaleValue As Variant, _
    ByVal absValue As Variant, _
    ByVal relValue As Variant, _
    Optional ByVal scaleName As String = NM_BN_SCALE, _
    Optional ByVal absName As String = NM_BN_ABSTOL, _
    Optional ByVal relName As String = NM_BN_RELTOL) _
    As String
'
'==============================================================================
'                               SettingsProblem
'------------------------------------------------------------------------------
' PURPOSE
'   Check the amount scale and the tolerances read from a bank sheet.
'
' INPUTS
'   scaleValue, absValue, relValue: the values of the three named cells,
'   as Value2.
'   scaleName, absName, relName: their workbook names, for the message;
'   those of BankNettingSets by default.
'
' RETURNS
'   "" when the scale is a number greater than zero and both tolerances
'   are numbers of zero or more; otherwise what is wrong.
'
' UPDATED
'   2026-10-15
'==============================================================================
'
        If Not IsNumberValue(scaleValue) Then
            SettingsProblem = "the amount scale (" & scaleName & ") is not a number"
        ElseIf CDbl(scaleValue) <= 0# Then
            SettingsProblem = "the amount scale (" & scaleName & ") must be greater than zero"
        ElseIf Not IsNumberValue(absValue) Then
            SettingsProblem = "the absolute tolerance (" & absName & ") is not a number"
        ElseIf CDbl(absValue) < 0# Then
            SettingsProblem = "the absolute tolerance (" & absName & ") must not be negative"
        ElseIf Not IsNumberValue(relValue) Then
            SettingsProblem = "the relative tolerance (" & relName & ") is not a number"
        ElseIf CDbl(relValue) < 0# Then
            SettingsProblem = "the relative tolerance (" & relName & ") must not be negative"
        End If

End Function


Public Function HeadersProblem( _
    ByVal headers As Variant, _
    ByVal expected As String) _
    As String
'
'==============================================================================
'                                HeadersProblem
'------------------------------------------------------------------------------
' PURPOSE
'   Check that a header row holds the expected headers in order, ignoring
'   case, spaces and line breaks, so that a moved or inserted column is
'   found before any figure is read from the wrong column.
'
' INPUTS
'   headers: the header cells, (1, 1 To n), as Value2.
'   expected: the headers separated by "|"; an empty entry accepts any
'   header.
'
' RETURNS
'   "" when every header matches; otherwise the first that does not.
'
' UPDATED
'   2026-10-15
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim want   As Variant    'Expected headers
    Dim c      As Long       'Column, 0-based

'------------------------------------------------------------------------------
' COMPARE
'------------------------------------------------------------------------------
        want = Split(expected, "|")
        For c = 0 To UBound(want)
            If Len(want(c)) > 0 Then
                If HeaderKey(headers(1, c + 1)) <> HeaderKey(want(c)) Then
                    HeadersProblem = "column " & (c + 1) & " of the table is headed '" & SafeStr(headers(1, c + 1)) & _
                                     "', expected '" & want(c) & "'"
                    Exit Function
                End If
            End If
        Next c

End Function


Public Function SummaryText( _
    ByRef counts() As Long, _
    ByVal scaleFactor As Double, _
    ByVal absTol As Double, _
    ByVal relTol As Double) _
    As String
'
'==============================================================================
'                                 SummaryText
'------------------------------------------------------------------------------
' PURPOSE
'   Describe a netting-set reconciliation in one line for the
'   Reconciliation summary cell: the bank netting sets, the figures
'   compared and their results, every netting set not compared, the scale
'   and the tolerances.
'
' INPUTS
'   counts: as filled by CompareNettingSets.
'   scaleFactor, absTol, relTol: as used.
'
' UPDATED
'   2026-10-15
'==============================================================================
'
        SummaryText = "bank netting sets " & counts(RCN_SETS) & _
                      " | figures compared " & counts(RCN_COMPARED) & ": MATCH " & counts(RCN_MATCH) & _
                      ", DIFFERENCE " & counts(RCN_DIFF) & _
                      " | missing in calculator " & counts(RCN_MISSING) & _
                      " | not in bank file " & counts(RCN_NOTINBANK) & _
                      " | duplicate rows " & counts(RCN_DUPLICATE) & _
                      " | regime mismatch " & counts(RCN_REGIME) & _
                      " | withheld " & counts(RCN_WITHHELD) & _
                      " | invalid input " & counts(RCN_INVALID) & _
                      " | not applicable " & counts(RCN_NA) & _
                      " | scale " & CStr(scaleFactor) & _
                      " | tolerance " & CStr(absTol) & " or " & CStr(relTol * 100#) & "%"

End Function


Public Function TradeSummaryText( _
    ByRef counts() As Long, _
    ByVal scaleFactor As Double, _
    ByVal absTol As Double, _
    ByVal relTol As Double) _
    As String
'
'==============================================================================
'                               TradeSummaryText
'------------------------------------------------------------------------------
' PURPOSE
'   Describe a trade reconciliation in one line for the
'   TradeReconciliation summary cell, as SummaryText.
'
' INPUTS
'   counts: as filled by CompareTrades.
'   scaleFactor, absTol, relTol: as used.
'
' UPDATED
'   2026-10-15
'==============================================================================
'
        TradeSummaryText = "bank trade legs " & counts(RCN_SETS) & _
                           " | figures compared " & counts(RCN_COMPARED) & ": MATCH " & counts(RCN_MATCH) & _
                           ", DIFFERENCE " & counts(RCN_DIFF) & _
                           " | not applicable " & counts(RCN_NA) & _
                           " | missing in calculator " & counts(RCN_MISSING) & _
                           " | not in bank file " & counts(RCN_NOTINBANK) & _
                           " | duplicate rows " & counts(RCN_DUPLICATE) & _
                           " | excluded " & counts(RCN_EXCLUDED) & _
                           " | invalid input " & counts(RCN_INVALID) & _
                           " | scale " & CStr(scaleFactor) & _
                           " | tolerance " & CStr(absTol) & " or " & CStr(relTol * 100#) & "%"

End Function


'
'------------------------------------------------------------------------------
'
'                                   HELPERS
'
'------------------------------------------------------------------------------
'

Private Sub AddRow( _
    ByRef outArr() As Variant, _
    ByRef nOut As Long, _
    ByVal id As String, _
    ByVal bankRow As Variant, _
    ByVal col3 As String, _
    ByVal col4 As String, _
    ByVal calcStatus As String, _
    ByVal measure As String, _
    ByVal bankValue As Variant, _
    ByVal calcValue As Variant, _
    ByVal result As String, _
    ByVal note As String, _
    Optional ByVal diff As Variant, _
    Optional ByVal rel As Variant)
'
'==============================================================================
'                                    AddRow
'------------------------------------------------------------------------------
' PURPOSE
'   Append one output row. Blank fields are left Empty, so that the sheet
'   shows empty cells rather than zeros. col3 and col4 are the bank and
'   calculator regimes of a netting set, or the netting set and asset
'   class of a trade.
'
' UPDATED
'   2026-10-15
'==============================================================================
'
        nOut = nOut + 1
        outArr(nOut, RC_ID) = id
        outArr(nOut, RC_BANKROW) = bankRow
        If Len(col3) > 0 Then outArr(nOut, RC_BANKREGIME) = col3
        If Len(col4) > 0 Then outArr(nOut, RC_CALCREGIME) = col4
        If Len(calcStatus) > 0 Then outArr(nOut, RC_CALCSTATUS) = calcStatus
        If Len(measure) > 0 Then outArr(nOut, RC_MEASURE) = measure
        outArr(nOut, RC_BANK) = bankValue
        outArr(nOut, RC_CALC) = calcValue
        If Not IsMissing(diff) Then outArr(nOut, RC_DIFF) = diff
        If Not IsMissing(rel) Then outArr(nOut, RC_REL) = rel
        outArr(nOut, RC_RESULT) = result
        If Len(note) > 0 Then outArr(nOut, RC_NOTE) = note

End Sub


Private Function CalcId( _
    ByVal kind As Long, _
    ByRef calc As Variant, _
    ByVal r As Long) _
    As String
'
'==============================================================================
'                                    CalcId
'------------------------------------------------------------------------------
' PURPOSE
'   Return the ID of a calculator row, upper case; "" for a reserved TOTAL
'   row of Results, which is not a netting set.
'
' UPDATED
'   2026-10-15
'==============================================================================
'
        If kind = KIND_NS Then
            CalcId = UTxt(calc(r, RS_ID))
            If IsTotalRow(CalcId) Then
                CalcId = ""
            End If
        Else
            CalcId = UTxt(calc(r, TC_ID))
        End If

End Function


Private Function CalcColumn( _
    ByVal kind As Long, _
    ByVal m As Long) _
    As Long
'
'==============================================================================
'                                  CalcColumn
'------------------------------------------------------------------------------
' PURPOSE
'   Return the Results or TradeCalc column of measure m, 0-based in the
'   order of NS_MEASURES or TR_MEASURES.
'
' UPDATED
'   2026-10-15
'==============================================================================
'
        If kind = KIND_NS Then
            Select Case m
                Case 0
                    CalcColumn = RS_EAD
                Case 1
                    CalcColumn = RS_RC
                Case 2
                    CalcColumn = RS_PFE
                Case 3
                    CalcColumn = RS_ADDON
                Case 10
                    CalcColumn = RS_MULT
                Case Else
                    CalcColumn = RS_ADDON_IR + m - 4
            End Select
        Else
            CalcColumn = Choose(m + 1, TC_NOTIONAL, TC_MTM, TC_M, TC_SD, TC_ADJN, TC_DELTA, TC_MFU, TC_MFM, _
                                TC_ENU, TC_ENM, TC_ADDON)
        End If

End Function


Private Function BankFigure( _
    ByVal v As Variant, _
    ByVal isRatio As Boolean, _
    ByVal scaleFactor As Double) _
    As Double
'
'==============================================================================
'                                  BankFigure
'------------------------------------------------------------------------------
' PURPOSE
'   Return a bank figure that is a number in the calculator's units: an
'   amount times the scale, a ratio or time as entered.
'
' UPDATED
'   2026-10-15
'==============================================================================
'
        If isRatio Then
            BankFigure = CDbl(v)
        Else
            BankFigure = CDbl(v) * scaleFactor
        End If

End Function


Private Function BankAmount( _
    ByVal v As Variant, _
    ByVal scaleFactor As Double) _
    As Variant
'
'==============================================================================
'                                  BankAmount
'------------------------------------------------------------------------------
' PURPOSE
'   Return a bank amount scaled to units, or Empty when it is not a number.
'
' UPDATED
'   2026-10-15
'==============================================================================
'
        If IsNumberValue(v) Then
            BankAmount = CDbl(v) * scaleFactor
        End If

End Function


Private Function NumberOrEmpty( _
    ByVal v As Variant) _
    As Variant
'
'==============================================================================
'                                NumberOrEmpty
'------------------------------------------------------------------------------
' PURPOSE
'   Return a Results figure as a number, or Empty when the cell has none.
'
' UPDATED
'   2026-10-15
'==============================================================================
'
        If IsNumberValue(v) Then
            NumberOrEmpty = CDbl(v)
        End If

End Function


Private Function IsNumberValue( _
    ByVal v As Variant) _
    As Boolean
'
'==============================================================================
'                                IsNumberValue
'------------------------------------------------------------------------------
' PURPOSE
'   Tell whether a cell value is a number. Text that looks like a number is
'   not, because how it reads depends on the locale; a logical value is
'   not either.
'
' UPDATED
'   2026-10-15
'==============================================================================
'
        Select Case VarType(v)
            Case vbDouble, vbSingle, vbLong, vbInteger, vbCurrency, vbDecimal, vbByte
                IsNumberValue = True
        End Select

End Function


Private Function IsBlankBankCell( _
    ByVal v As Variant) _
    As Boolean
'
'==============================================================================
'                               IsBlankBankCell
'------------------------------------------------------------------------------
' PURPOSE
'   Tell whether a bank cell is empty or holds only spaces. Unlike
'   CORE_Util.IsBlankCell, an error value is not blank: it is an unusable
'   figure.
'
' UPDATED
'   2026-10-15
'==============================================================================
'
        If IsError(v) Then
            IsBlankBankCell = False
        Else
            IsBlankBankCell = (Len(SafeStr(v)) = 0)
        End If

End Function


Private Function BankRowHasData( _
    ByRef bank As Variant, _
    ByVal r As Long, _
    ByVal commentCol As Long) _
    As Boolean
'
'==============================================================================
'                                BankRowHasData
'------------------------------------------------------------------------------
' PURPOSE
'   Tell whether a bank row has a value in any column before the comment
'   column commentCol.
'
' UPDATED
'   2026-10-15
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim c   As Long    'Column

'------------------------------------------------------------------------------
' TEST
'------------------------------------------------------------------------------
        For c = 1 To commentCol - 1
            If Not IsBlankBankCell(bank(r, c)) Then
                BankRowHasData = True
                Exit Function
            End If
        Next c

End Function


Private Function DescribeValue( _
    ByVal v As Variant) _
    As String
'
'==============================================================================
'                                DescribeValue
'------------------------------------------------------------------------------
' PURPOSE
'   Describe a bank value that is not a number, for a note.
'
' UPDATED
'   2026-10-15
'==============================================================================
'
        If IsError(v) Then
            DescribeValue = "an error value"
        ElseIf VarType(v) = vbBoolean Then
            DescribeValue = "the logical value " & UCase$(CStr(v))
        Else
            DescribeValue = "the text '" & SafeStr(v) & "'"
        End If

End Function


Private Function HeaderKey( _
    ByVal v As Variant) _
    As String
'
'==============================================================================
'                                  HeaderKey
'------------------------------------------------------------------------------
' PURPOSE
'   Reduce a header to lower case without spaces or line breaks.
'
' UPDATED
'   2026-10-15
'==============================================================================
'
        HeaderKey = LCase$(Replace(Replace(Replace(SafeStr(v), " ", ""), vbCr, ""), vbLf, ""))

End Function


Private Function IsTotalRow( _
    ByVal id As String) _
    As Boolean
'
'==============================================================================
'                                  IsTotalRow
'------------------------------------------------------------------------------
' PURPOSE
'   Tell whether a Results ID is a TOTAL row (#198), which is not a netting
'   set: one of the IDs the engine reserves for them, so that a netting set
'   such as TOTAL DESK is still compared.
'
' UPDATED
'   2026-10-15
'==============================================================================
'
        IsTotalRow = (id = "TOTAL" Or id = "TOTAL " & RG_CRR Or id = "TOTAL " & RG_BCBS)

End Function


Private Function RowCount( _
    ByVal block As Variant) _
    As Long
'
'==============================================================================
'                                   RowCount
'------------------------------------------------------------------------------
' PURPOSE
'   Return the number of rows of a 2-D block; 0 when it is not an array.
'
' UPDATED
'   2026-10-15
'==============================================================================
'
        If IsArray(block) Then
            RowCount = UBound(block, 1) - LBound(block, 1) + 1
        End If

End Function


Private Function TrimRows( _
    ByRef outArr() As Variant, _
    ByVal nOut As Long) _
    As Variant
'
'==============================================================================
'                                   TrimRows
'------------------------------------------------------------------------------
' PURPOSE
'   Return the first nOut rows of outArr as a new array.
'
' UPDATED
'   2026-10-15
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim result()   As Variant    'The rows returned
    Dim r          As Long       'Row
    Dim c          As Long       'Column

'------------------------------------------------------------------------------
' COPY
'------------------------------------------------------------------------------
        ReDim result(1 To nOut, 1 To RC_NCOLS)
        For r = 1 To nOut
            For c = 1 To RC_NCOLS
                result(r, c) = outArr(r, c)
            Next c
        Next r
        TrimRows = result

End Function
