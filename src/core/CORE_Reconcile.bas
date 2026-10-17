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
'   tolerance applies to them: such a figure matches when it is equal or
'   within the relative tolerance.
'
' PUBLIC SURFACE
'   None outside this VBA project. CompareNettingSets, CompareTrades,
'   HasBankRows, SettingsProblem, HeadersProblem, SummaryText and
'   TradeSummaryText, with the RCN_ count indexes, are Public for M_Main
'   and TEST_Reconcile; Option Private Module keeps them off the supported
'   external surface.
'
' DEPENDENCIES
'   CORE_Config for the column layout, the result labels, the regime codes
'   and the workbook names of the settings; CORE_Util for text conversion
'   and the Collection index (KeyIndex, KeyAdd);
'   CORE_Export for numbers written the same in every locale.
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
'   2026-10-17
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
    ByVal Bank As Variant, _
    ByVal BankFirstRow As Long, _
    ByVal Calc As Variant, _
    ByVal ScaleFactor As Double, _
    ByVal AbsTolerance As Double, _
    ByVal RelTolerance As Double, _
    ByRef Counts() As Long) _
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
'   Bank: the BankNettingSets rows, (row, 1 To BN_NCOLS), as Value2; Empty
'   when there are none. Rows with no value before the comment are skipped.
'   BankFirstRow: the sheet row of the first bank row, for the Bank row
'   column.
'   Calc: the Results rows, (row, 1 To RS_NCOLS); Empty when there are
'   none. The reserved TOTAL rows are skipped.
'   ScaleFactor: units per bank amount, greater than zero.
'   AbsTolerance, RelTolerance: the tolerances, zero or more.
'   Counts: receives the counts, (1 To RCN_COUNT).
'
' RETURNS
'   The rows, (row, 1 To RC_NCOLS); Empty when there are none.
'
' UPDATED
'   2026-10-17
'==============================================================================
'
        CompareNettingSets = CompareTable(KIND_NS, Bank, BankFirstRow, Calc, ScaleFactor, AbsTolerance, RelTolerance, Counts)

End Function


Public Function CompareTrades( _
    ByVal Bank As Variant, _
    ByVal BankFirstRow As Long, _
    ByVal Calc As Variant, _
    ByVal ScaleFactor As Double, _
    ByVal AbsTolerance As Double, _
    ByVal RelTolerance As Double, _
    ByRef Counts() As Long) _
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
'   Bank: the BankTrades rows, (row, 1 To BT_NCOLS), as Value2; Empty when
'   there are none. Rows with no value before the comment are skipped.
'   Calc: the TradeCalc rows, (row, 1 To TC_NCOLS); Empty when there are
'   none.
'   BankFirstRow, ScaleFactor, AbsTolerance, RelTolerance, Counts: as
'   CompareNettingSets.
'
' RETURNS
'   The rows, (row, 1 To RC_NCOLS); Empty when there are none.
'
' UPDATED
'   2026-10-17
'==============================================================================
'
        CompareTrades = CompareTable(KIND_TR, Bank, BankFirstRow, Calc, ScaleFactor, AbsTolerance, RelTolerance, Counts)

End Function


Public Function HasBankRows( _
    ByVal Bank As Variant, _
    Optional ByVal CommentColumn As Long = BN_COMMENT) _
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
'   Bank: as CompareNettingSets or CompareTrades.
'   CommentColumn: the comment column: BN_COMMENT (the default) or
'   BT_COMMENT.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim r   As Long    'Bank row

'------------------------------------------------------------------------------
' TEST
'------------------------------------------------------------------------------
        For r = 1 To RowCount(Bank)
            If BankRowHasData(Bank, r, CommentColumn) Then
                HasBankRows = True
                Exit Function
            End If
        Next r

End Function


Private Function CompareTable( _
    ByVal Kind As Long, _
    ByRef Bank As Variant, _
    ByVal BankFirstRow As Long, _
    ByRef Calc As Variant, _
    ByVal ScaleFactor As Double, _
    ByVal AbsTolerance As Double, _
    ByVal RelTolerance As Double, _
    ByRef Counts() As Long) _
    As Variant
'
'==============================================================================
'                                 CompareTable
'------------------------------------------------------------------------------
' PURPOSE
'   The comparison behind CompareNettingSets and CompareTrades: index the
'   calculator rows by ID, count the bank IDs, compare each bank row
'   (CompareNettingSetRow or CompareTradeRow), then list the calculator
'   rows that are not in the bank file. A Trade ID listed more than once on
'   TradeCalc (repeated on Trades, where the engine excludes the later
'   rows) is not compared: which row is the calculator's leg cannot be
'   told.
'
' INPUTS
'   Kind: KIND_NS or KIND_TR; the other arguments as CompareNettingSets.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim CalcIndex           As Collection    'Calculator ID to calculator row
    Dim CalcUsed()          As Boolean       'The calculator row is in the bank file
    Dim CalcTimesListed()   As Long          'Rows of each calculator ID, by its first row
    Dim CalcListed          As Long          'Calculator rows of a bank ID
    Dim BankIndex           As Collection    'Bank ID to position in BankTimesListed
    Dim BankTimesListed()   As Long          'Bank rows of each distinct ID
    Dim CommentColumn       As Long          'Comment column of the bank table
    Dim CalcRowCount        As Long          'Calculator rows
    Dim BankRowCount        As Long          'Bank rows
    Dim OutRows()           As Variant       'Rows being built
    Dim OutCount            As Long          'Rows filled
    Dim r                   As Long          'Bank or calculator row
    Dim IdPosition          As Long          'Position of an ID
    Dim RowId               As String        'ID, upper case
    Dim Result              As Variant       'The rows returned

'------------------------------------------------------------------------------
' INDEX THE CALCULATOR ROWS
'------------------------------------------------------------------------------
        ReDim Counts(1 To RCN_COUNT)
        CommentColumn = IIf(Kind = KIND_NS, BN_COMMENT, BT_COMMENT)
        CalcRowCount = RowCount(Calc)
        BankRowCount = RowCount(Bank)
        Set CalcIndex = New Collection
        If CalcRowCount > 0 Then
            ReDim CalcUsed(1 To CalcRowCount)
            ReDim CalcTimesListed(1 To CalcRowCount)
        End If
        For r = 1 To CalcRowCount
            RowId = CalcId(Kind, Calc, r)
            If Len(RowId) > 0 Then
                IdPosition = KeyIndex(CalcIndex, RowId)
                If IdPosition = 0 Then
                    KeyAdd CalcIndex, RowId, r
                    CalcTimesListed(r) = 1
                Else
                    CalcTimesListed(IdPosition) = CalcTimesListed(IdPosition) + 1
                End If
            End If
        Next r

'------------------------------------------------------------------------------
' COUNT THE BANK IDS
'------------------------------------------------------------------------------
    'An ID listed twice is not compared at all: which row is the bank's
    'figure cannot be told.
        Set BankIndex = New Collection
        If BankRowCount > 0 Then
            ReDim BankTimesListed(1 To BankRowCount)
        End If
        For r = 1 To BankRowCount
            If BankRowHasData(Bank, r, CommentColumn) Then
                RowId = UTxt(Bank(r, 1))
                If Len(RowId) > 0 Then
                    IdPosition = KeyIndex(BankIndex, RowId)
                    If IdPosition = 0 Then
                        IdPosition = BankIndex.Count + 1
                        KeyAdd BankIndex, RowId, IdPosition
                    End If
                    BankTimesListed(IdPosition) = BankTimesListed(IdPosition) + 1
                End If
            End If
        Next r
        Counts(RCN_SETS) = BankIndex.Count

'------------------------------------------------------------------------------
' COMPARE THE BANK ROWS
'------------------------------------------------------------------------------
    'At most one row per measure for each bank row, plus one row for each
    'calculator row.
        ReDim OutRows(1 To BankRowCount * MEASURE_COUNT + CalcRowCount + 1, 1 To RC_NCOLS)
        For r = 1 To BankRowCount
            If BankRowHasData(Bank, r, CommentColumn) Then
                RowId = UTxt(Bank(r, 1))
                If Len(RowId) = 0 Then
                    AddRow OutRows, OutCount, "", BankFirstRow + r - 1, _
                           IIf(Kind = KIND_NS, UTxt(Bank(r, BN_REGIME)), ""), "", "", "", _
                           Empty, Empty, RES_INVALID, IIf(Kind = KIND_NS, "NettingSetID", "TradeID") & _
                           " missing: the row is not compared."
                    Counts(RCN_INVALID) = Counts(RCN_INVALID) + 1
                Else
                    IdPosition = KeyIndex(CalcIndex, RowId)
                    If IdPosition > 0 Then
                        CalcUsed(IdPosition) = True
                    End If
                    If Kind = KIND_NS Then
                        CompareNettingSetRow Bank, r, BankFirstRow + r - 1, Calc, IdPosition, _
                                             BankTimesListed(KeyIndex(BankIndex, RowId)), _
                                             ScaleFactor, AbsTolerance, RelTolerance, OutRows, OutCount, Counts
                    Else
                        CalcListed = 0
                        If IdPosition > 0 Then
                            CalcListed = CalcTimesListed(IdPosition)
                        End If
                        CompareTradeRow Bank, r, BankFirstRow + r - 1, Calc, IdPosition, _
                                        BankTimesListed(KeyIndex(BankIndex, RowId)), CalcListed, _
                                        ScaleFactor, AbsTolerance, RelTolerance, OutRows, OutCount, Counts
                    End If
                End If
            End If
        Next r

'------------------------------------------------------------------------------
' LIST THE CALCULATOR ROWS NOT IN THE BANK FILE
'------------------------------------------------------------------------------
    'With the EAD of a netting set, or the notional of a trade, for scale.
        For r = 1 To CalcRowCount
            RowId = CalcId(Kind, Calc, r)
            If Len(RowId) > 0 Then
                If KeyIndex(CalcIndex, RowId) = r And Not CalcUsed(r) Then
                    If Kind = KIND_TR And CalcTimesListed(r) > 1 Then
                        AddRow OutRows, OutCount, SafeStr(Calc(r, TC_ID)), Empty, SafeStr(Calc(r, TC_NS)), _
                               SafeStr(Calc(r, TC_AC)), "", "", Empty, Empty, RES_DUPLICATE, _
                               "Listed " & CalcTimesListed(r) & " times on TradeCalc (Trade ID repeated on Trades): " & _
                               "none of its rows is compared."
                        Counts(RCN_DUPLICATE) = Counts(RCN_DUPLICATE) + 1
                    ElseIf Kind = KIND_NS Then
                        AddRow OutRows, OutCount, RowId, Empty, "", UTxt(Calc(r, RS_REGIME)), _
                               SafeStr(Calc(r, RS_STATUS_COL)), "EAD", Empty, NumberOrEmpty(Calc(r, RS_EAD)), _
                               RES_NOTINBANK, "No row for this netting set on BankNettingSets."
                    Else
                        AddRow OutRows, OutCount, SafeStr(Calc(r, TC_ID)), Empty, SafeStr(Calc(r, TC_NS)), _
                               SafeStr(Calc(r, TC_AC)), SafeStr(Calc(r, TC_STATUS)), "Notional", Empty, _
                               NumberOrEmpty(Calc(r, TC_NOTIONAL)), RES_NOTINBANK, _
                               "No row for this trade leg on BankTrades."
                    End If
                    If Not (Kind = KIND_TR And CalcTimesListed(r) > 1) Then
                        Counts(RCN_NOTINBANK) = Counts(RCN_NOTINBANK) + 1
                    End If
                End If
            End If
        Next r

'------------------------------------------------------------------------------
' RETURN THE ROWS FILLED
'------------------------------------------------------------------------------
        If OutCount > 0 Then
            Result = TrimRows(OutRows, OutCount)
        End If
        CompareTable = Result

End Function


Private Sub CompareNettingSetRow( _
    ByRef Bank As Variant, _
    ByVal BankRow As Long, _
    ByVal SheetRow As Long, _
    ByRef Calc As Variant, _
    ByVal CalcRow As Long, _
    ByVal TimesListed As Long, _
    ByVal ScaleFactor As Double, _
    ByVal AbsTolerance As Double, _
    ByVal RelTolerance As Double, _
    ByRef OutRows() As Variant, _
    ByRef OutCount As Long, _
    ByRef Counts() As Long)
'
'==============================================================================
'                             CompareNettingSetRow
'------------------------------------------------------------------------------
' PURPOSE
'   Add the Reconciliation rows of one bank row with an ID: one row saying
'   why the netting set is not compared, or one row per figure entered.
'
' INPUTS
'   Bank, BankRow: the bank rows and this row's position; SheetRow: its
'   sheet row.
'   Calc, CalcRow: the Results rows and the matching row; 0 when none.
'   TimesListed: rows of this ID in the bank file.
'   ScaleFactor, AbsTolerance, RelTolerance: as CompareNettingSets.
'   OutRows, OutCount, Counts: the rows and counts being built.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim BankId       As String    'Netting-set ID, upper case
    Dim BankRegime   As String    'Regime given by the bank
    Dim CalcRegime   As String    'Regime on Results
    Dim CalcStatus   As String    'Status on Results

'------------------------------------------------------------------------------
' NETTING-SET CHECKS
'------------------------------------------------------------------------------
    'One reason per netting set, in this order: a figure of an ID listed
    'twice, of another regime or of a withheld netting set means nothing.
        BankId = UTxt(Bank(BankRow, BN_ID))
        BankRegime = UTxt(Bank(BankRow, BN_REGIME))
        If CalcRow > 0 Then
            CalcRegime = UTxt(Calc(CalcRow, RS_REGIME))
            CalcStatus = SafeStr(Calc(CalcRow, RS_STATUS_COL))
        End If
        If TimesListed > 1 Then
            AddRow OutRows, OutCount, BankId, SheetRow, BankRegime, CalcRegime, CalcStatus, "", Empty, Empty, _
                   RES_DUPLICATE, "Listed " & TimesListed & " times on BankNettingSets: none of its rows is compared."
            Counts(RCN_DUPLICATE) = Counts(RCN_DUPLICATE) + 1
            Exit Sub
        End If
        If BankRegime <> RG_BCBS And BankRegime <> RG_CRR Then
            AddRow OutRows, OutCount, BankId, SheetRow, BankRegime, CalcRegime, CalcStatus, "", Empty, Empty, _
                   RES_INVALID, "Regime must be BCBS or CRR: the row is not compared."
            Counts(RCN_INVALID) = Counts(RCN_INVALID) + 1
            Exit Sub
        End If
        If CalcRow = 0 Then
            AddRow OutRows, OutCount, BankId, SheetRow, BankRegime, "", "", "EAD", _
                   BankAmount(Bank(BankRow, BN_EAD), ScaleFactor), _
                   Empty, RES_MISSING, "No netting set with this ID on Results."
            Counts(RCN_MISSING) = Counts(RCN_MISSING) + 1
            Exit Sub
        End If
        If BankRegime <> CalcRegime Then
            AddRow OutRows, OutCount, BankId, SheetRow, BankRegime, CalcRegime, CalcStatus, "", Empty, Empty, _
                   RES_REGIME, "Bank regime " & BankRegime & ", Results regime " & CalcRegime & _
                   ": figures of different regimes are not compared."
            Counts(RCN_REGIME) = Counts(RCN_REGIME) + 1
            Exit Sub
        End If
        If CalcStatus <> "VALID" Then
            AddRow OutRows, OutCount, BankId, SheetRow, BankRegime, CalcRegime, CalcStatus, "", Empty, Empty, _
                   RES_WITHHELD, "No EAD on Results for this netting set: see its status and the Checks sheet."
            Counts(RCN_WITHHELD) = Counts(RCN_WITHHELD) + 1
            Exit Sub
        End If

'------------------------------------------------------------------------------
' COMPARE EACH FIGURE ENTERED
'------------------------------------------------------------------------------
        CompareFigures KIND_NS, Bank, BankRow, BN_EAD, BankId, SheetRow, BankRegime, CalcRegime, CalcStatus, _
                       Calc, CalcRow, ScaleFactor, AbsTolerance, RelTolerance, OutRows, OutCount, Counts

End Sub


Private Sub CompareTradeRow( _
    ByRef Bank As Variant, _
    ByVal BankRow As Long, _
    ByVal SheetRow As Long, _
    ByRef Calc As Variant, _
    ByVal CalcRow As Long, _
    ByVal TimesListed As Long, _
    ByVal CalcListed As Long, _
    ByVal ScaleFactor As Double, _
    ByVal AbsTolerance As Double, _
    ByVal RelTolerance As Double, _
    ByRef OutRows() As Variant, _
    ByRef OutCount As Long, _
    ByRef Counts() As Long)
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
'   As CompareNettingSetRow, with the TradeCalc rows as Calc; CalcListed:
'   the TradeCalc rows with this Trade ID, 0 when there is none.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim BankId           As String    'TradeID as entered, trimmed
    Dim CalcNettingSet   As String    'Netting set on TradeCalc
    Dim CalcAssetClass   As String    'Asset class on TradeCalc
    Dim CalcStatus       As String    'OK or EXCLUDED

'------------------------------------------------------------------------------
' TRADE CHECKS
'------------------------------------------------------------------------------
        BankId = SafeStr(Bank(BankRow, BT_ID))
        If CalcRow > 0 Then
            CalcNettingSet = SafeStr(Calc(CalcRow, TC_NS))
            CalcAssetClass = SafeStr(Calc(CalcRow, TC_AC))
            CalcStatus = SafeStr(Calc(CalcRow, TC_STATUS))
        End If
        If TimesListed > 1 Then
            AddRow OutRows, OutCount, BankId, SheetRow, CalcNettingSet, CalcAssetClass, CalcStatus, "", Empty, Empty, _
                   RES_DUPLICATE, "Listed " & TimesListed & " times on BankTrades: none of its rows is compared."
            Counts(RCN_DUPLICATE) = Counts(RCN_DUPLICATE) + 1
            Exit Sub
        End If
        If CalcListed > 1 Then
            AddRow OutRows, OutCount, BankId, SheetRow, CalcNettingSet, CalcAssetClass, "", "", Empty, Empty, RES_DUPLICATE, _
                   "Listed " & CalcListed & " times on TradeCalc (Trade ID repeated on Trades): none of its rows " & _
                   "is compared."
            Counts(RCN_DUPLICATE) = Counts(RCN_DUPLICATE) + 1
            Exit Sub
        End If
        If CalcRow = 0 Then
            AddRow OutRows, OutCount, BankId, SheetRow, "", "", "", "Notional", _
                   BankAmount(Bank(BankRow, BT_NOTIONAL), ScaleFactor), Empty, RES_MISSING, _
                   "No trade leg with this ID on TradeCalc."
            Counts(RCN_MISSING) = Counts(RCN_MISSING) + 1
            Exit Sub
        End If
        If UCase$(CalcStatus) = "EXCLUDED" Then
            AddRow OutRows, OutCount, BankId, SheetRow, CalcNettingSet, CalcAssetClass, CalcStatus, "", Empty, Empty, _
                   RES_EXCLUDED, "Excluded on TradeCalc: " & Left$(SafeStr(Calc(CalcRow, TC_MESSAGES)), 250)
            Counts(RCN_EXCLUDED) = Counts(RCN_EXCLUDED) + 1
            Exit Sub
        End If

'------------------------------------------------------------------------------
' COMPARE EACH FIGURE ENTERED
'------------------------------------------------------------------------------
        CompareFigures KIND_TR, Bank, BankRow, BT_NOTIONAL, BankId, SheetRow, CalcNettingSet, CalcAssetClass, CalcStatus, _
                       Calc, CalcRow, ScaleFactor, AbsTolerance, RelTolerance, OutRows, OutCount, Counts

End Sub


Private Sub CompareFigures( _
    ByVal Kind As Long, _
    ByRef Bank As Variant, _
    ByVal BankRow As Long, _
    ByVal FirstColumn As Long, _
    ByVal RowId As String, _
    ByVal SheetRow As Long, _
    ByVal ThirdColumn As String, _
    ByVal FourthColumn As String, _
    ByVal CalcStatus As String, _
    ByRef Calc As Variant, _
    ByVal CalcRow As Long, _
    ByVal ScaleFactor As Double, _
    ByVal AbsTolerance As Double, _
    ByVal RelTolerance As Double, _
    ByRef OutRows() As Variant, _
    ByRef OutCount As Long, _
    ByRef Counts() As Long)
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
'   Kind: KIND_NS or KIND_TR; FirstColumn: the bank column of the first
'   measure; ThirdColumn, FourthColumn: what the third and fourth output
'   columns show
'   (the two regimes, or the trade's netting set and asset class); the
'   other arguments as CompareNettingSetRow.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim MeasureNames         As Variant    'Measure names
    Dim m                    As Long       'Measure, 0-based
    Dim BankCell             As Variant    'Bank figure as entered
    Dim CalcCell             As Variant    'Calculator figure as written
    Dim IsRatio              As Boolean    'Not scaled; relative tolerance only
    Dim BankValue            As Double     'Bank figure, in units for an amount
    Dim CalcValue            As Double     'Calculator figure
    Dim Difference           As Double     'Calculator minus bank
    Dim RelativeDifference   As Variant    'Relative difference; Empty when the bank figure is 0
    Dim IsMatch              As Boolean    'Within a tolerance
    Dim EnteredCount         As Long       'Figures entered on the row

'------------------------------------------------------------------------------
' COMPARE
'------------------------------------------------------------------------------
        MeasureNames = Split(IIf(Kind = KIND_NS, NS_MEASURES, TR_MEASURES), "|")
        For m = 0 To MEASURE_COUNT - 1
            BankCell = Bank(BankRow, FirstColumn + m)
            If Not IsBlankBankCell(BankCell) Then
                EnteredCount = EnteredCount + 1
                CalcCell = Calc(CalcRow, CalcColumn(Kind, m))
                IsRatio = InStr(IIf(Kind = KIND_NS, NS_RATIOS, TR_RATIOS), "|" & m & "|") > 0
                If Not IsNumberValue(BankCell) Then
                    AddRow OutRows, OutCount, RowId, SheetRow, ThirdColumn, FourthColumn, CalcStatus, CStr(MeasureNames(m)), _
                           Empty, Empty, RES_INVALID, "Not a number: " & DescribeValue(BankCell) & "."
                    Counts(RCN_INVALID) = Counts(RCN_INVALID) + 1
                ElseIf Not IsNumberValue(CalcCell) Then
                    AddRow OutRows, OutCount, RowId, SheetRow, ThirdColumn, FourthColumn, CalcStatus, CStr(MeasureNames(m)), _
                           BankFigure(BankCell, IsRatio, ScaleFactor), Empty, RES_NA, _
                           "The calculator has no " & MeasureNames(m) & " for this row."
                    Counts(RCN_NA) = Counts(RCN_NA) + 1
                Else
                    BankValue = BankFigure(BankCell, IsRatio, ScaleFactor)
                    CalcValue = CDbl(CalcCell)
                    Difference = CalcValue - BankValue
                    If BankValue <> 0# Then
                        RelativeDifference = Difference / Abs(BankValue)
                    Else
                        RelativeDifference = Empty
                    End If
                'A ratio or time matches only when equal or within the relative
                'tolerance: the absolute tolerance is for amounts.
                    If IsRatio Then
                        IsMatch = (Difference = 0#)
                    Else
                        IsMatch = (Abs(Difference) <= AbsTolerance)
                    End If
                    If Not IsEmpty(RelativeDifference) Then
                        IsMatch = IsMatch Or (Abs(RelativeDifference) <= RelTolerance)
                    End If
                    AddRow OutRows, OutCount, RowId, SheetRow, ThirdColumn, FourthColumn, CalcStatus, CStr(MeasureNames(m)), _
                           BankValue, CalcValue, IIf(IsMatch, RES_MATCH, RES_DIFF), "", Difference, RelativeDifference
                    Counts(RCN_COMPARED) = Counts(RCN_COMPARED) + 1
                    If IsMatch Then
                        Counts(RCN_MATCH) = Counts(RCN_MATCH) + 1
                    Else
                        Counts(RCN_DIFF) = Counts(RCN_DIFF) + 1
                    End If
                End If
            End If
        Next m
        If EnteredCount = 0 Then
            AddRow OutRows, OutCount, RowId, SheetRow, ThirdColumn, FourthColumn, CalcStatus, "", Empty, Empty, _
                   RES_INVALID, "No bank figure entered on the row."
            Counts(RCN_INVALID) = Counts(RCN_INVALID) + 1
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
    ByVal ScaleValue As Variant, _
    ByVal AbsToleranceValue As Variant, _
    ByVal RelToleranceValue As Variant, _
    Optional ByVal ScaleName As String = NM_BN_SCALE, _
    Optional ByVal AbsToleranceName As String = NM_BN_ABSTOL, _
    Optional ByVal RelToleranceName As String = NM_BN_RELTOL) _
    As String
'
'==============================================================================
'                               SettingsProblem
'------------------------------------------------------------------------------
' PURPOSE
'   Check the amount scale and the tolerances read from a bank sheet.
'
' INPUTS
'   ScaleValue, AbsToleranceValue, RelToleranceValue: the values of the
'   three named cells, as Value2.
'   ScaleName, AbsToleranceName, RelToleranceName: their workbook names,
'   for the message;
'   those of BankNettingSets by default.
'
' RETURNS
'   "" when the scale is a number greater than zero and both tolerances
'   are numbers of zero or more; otherwise what is wrong.
'
' UPDATED
'   2026-10-17
'==============================================================================
'
        If Not IsNumberValue(ScaleValue) Then
            SettingsProblem = "the amount scale (" & ScaleName & ") is not a number"
        ElseIf CDbl(ScaleValue) <= 0# Then
            SettingsProblem = "the amount scale (" & ScaleName & ") must be greater than zero"
        ElseIf Not IsNumberValue(AbsToleranceValue) Then
            SettingsProblem = "the absolute tolerance (" & AbsToleranceName & ") is not a number"
        ElseIf CDbl(AbsToleranceValue) < 0# Then
            SettingsProblem = "the absolute tolerance (" & AbsToleranceName & ") must not be negative"
        ElseIf Not IsNumberValue(RelToleranceValue) Then
            SettingsProblem = "the relative tolerance (" & RelToleranceName & ") is not a number"
        ElseIf CDbl(RelToleranceValue) < 0# Then
            SettingsProblem = "the relative tolerance (" & RelToleranceName & ") must not be negative"
        End If

End Function


Public Function HeadersProblem( _
    ByVal Headers As Variant, _
    ByVal ExpectedText As String) _
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
'   Headers: the header cells, (1, 1 To n), as Value2.
'   ExpectedText: the headers separated by "|"; an empty entry accepts any
'   header.
'
' RETURNS
'   "" when every header matches; otherwise the first that does not.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ExpectedHeaders   As Variant    'Expected headers
    Dim c                 As Long       'Column, 0-based

'------------------------------------------------------------------------------
' COMPARE
'------------------------------------------------------------------------------
        ExpectedHeaders = Split(ExpectedText, "|")
        For c = 0 To UBound(ExpectedHeaders)
            If Len(ExpectedHeaders(c)) > 0 Then
                If HeaderKey(Headers(1, c + 1)) <> HeaderKey(ExpectedHeaders(c)) Then
                    HeadersProblem = "column " & (c + 1) & " of the table is headed '" & SafeStr(Headers(1, c + 1)) & _
                                     "', expected '" & ExpectedHeaders(c) & "'"
                    Exit Function
                End If
            End If
        Next c

End Function


Public Function SummaryText( _
    ByRef Counts() As Long, _
    ByVal ScaleFactor As Double, _
    ByVal AbsTolerance As Double, _
    ByVal RelTolerance As Double) _
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
'   Counts: as filled by CompareNettingSets.
'   ScaleFactor, AbsTolerance, RelTolerance: as used.
'
' UPDATED
'   2026-10-17
'==============================================================================
'
        SummaryText = "bank netting sets " & Counts(RCN_SETS) & _
                      " | figures compared " & Counts(RCN_COMPARED) & ": MATCH " & Counts(RCN_MATCH) & _
                      ", DIFFERENCE " & Counts(RCN_DIFF) & _
                      " | missing in calculator " & Counts(RCN_MISSING) & _
                      " | not in bank file " & Counts(RCN_NOTINBANK) & _
                      " | duplicate rows " & Counts(RCN_DUPLICATE) & _
                      " | regime mismatch " & Counts(RCN_REGIME) & _
                      " | withheld " & Counts(RCN_WITHHELD) & _
                      " | invalid input " & Counts(RCN_INVALID) & _
                      " | not applicable " & Counts(RCN_NA) & _
                      " | scale " & CORE_Export.NumberText(ScaleFactor) & _
                      " | tolerance " & CORE_Export.NumberText(AbsTolerance) & " or " & _
                      CORE_Export.NumberText(RelTolerance * 100#) & "%"

End Function


Public Function TradeSummaryText( _
    ByRef Counts() As Long, _
    ByVal ScaleFactor As Double, _
    ByVal AbsTolerance As Double, _
    ByVal RelTolerance As Double) _
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
'   Counts: as filled by CompareTrades.
'   ScaleFactor, AbsTolerance, RelTolerance: as used.
'
' UPDATED
'   2026-10-17
'==============================================================================
'
        TradeSummaryText = "bank trade legs " & Counts(RCN_SETS) & _
                           " | figures compared " & Counts(RCN_COMPARED) & ": MATCH " & Counts(RCN_MATCH) & _
                           ", DIFFERENCE " & Counts(RCN_DIFF) & _
                           " | not applicable " & Counts(RCN_NA) & _
                           " | missing in calculator " & Counts(RCN_MISSING) & _
                           " | not in bank file " & Counts(RCN_NOTINBANK) & _
                           " | duplicate rows " & Counts(RCN_DUPLICATE) & _
                           " | excluded " & Counts(RCN_EXCLUDED) & _
                           " | invalid input " & Counts(RCN_INVALID) & _
                           " | scale " & CORE_Export.NumberText(ScaleFactor) & _
                           " | tolerance " & CORE_Export.NumberText(AbsTolerance) & " or " & _
                           CORE_Export.NumberText(RelTolerance * 100#) & "%"

End Function


'
'------------------------------------------------------------------------------
'
'                                   HELPERS
'
'------------------------------------------------------------------------------
'

Private Sub AddRow( _
    ByRef OutRows() As Variant, _
    ByRef OutCount As Long, _
    ByVal RowId As String, _
    ByVal BankRow As Variant, _
    ByVal ThirdColumn As String, _
    ByVal FourthColumn As String, _
    ByVal CalcStatus As String, _
    ByVal Measure As String, _
    ByVal BankValue As Variant, _
    ByVal CalcValue As Variant, _
    ByVal ResultLabel As String, _
    ByVal Note As String, _
    Optional ByVal Difference As Variant, _
    Optional ByVal RelativeDifference As Variant)
'
'==============================================================================
'                                    AddRow
'------------------------------------------------------------------------------
' PURPOSE
'   Append one output row. Blank fields are left Empty, so that the sheet
'   shows empty cells rather than zeros. ThirdColumn and FourthColumn are
'   the bank and calculator regimes of a netting set, or the netting set
'   and asset class of a trade; Difference and RelativeDifference are left
'   Empty when they are not passed.
'
' UPDATED
'   2026-10-17
'==============================================================================
'
        OutCount = OutCount + 1
        OutRows(OutCount, RC_ID) = RowId
        OutRows(OutCount, RC_BANKROW) = BankRow
        If Len(ThirdColumn) > 0 Then OutRows(OutCount, RC_BANKREGIME) = ThirdColumn
        If Len(FourthColumn) > 0 Then OutRows(OutCount, RC_CALCREGIME) = FourthColumn
        If Len(CalcStatus) > 0 Then OutRows(OutCount, RC_CALCSTATUS) = CalcStatus
        If Len(Measure) > 0 Then OutRows(OutCount, RC_MEASURE) = Measure
        OutRows(OutCount, RC_BANK) = BankValue
        OutRows(OutCount, RC_CALC) = CalcValue
        If Not IsMissing(Difference) Then OutRows(OutCount, RC_DIFF) = Difference
        If Not IsMissing(RelativeDifference) Then OutRows(OutCount, RC_REL) = RelativeDifference
        OutRows(OutCount, RC_RESULT) = ResultLabel
        If Len(Note) > 0 Then OutRows(OutCount, RC_NOTE) = Note

End Sub


Private Function CalcId( _
    ByVal Kind As Long, _
    ByRef Calc As Variant, _
    ByVal CalcRow As Long) _
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
'   2026-10-17
'==============================================================================
'
        If Kind = KIND_NS Then
            CalcId = UTxt(Calc(CalcRow, RS_ID))
            If IsTotalRow(CalcId) Then
                CalcId = ""
            End If
        Else
            CalcId = UTxt(Calc(CalcRow, TC_ID))
        End If

End Function


Private Function CalcColumn( _
    ByVal Kind As Long, _
    ByVal Measure As Long) _
    As Long
'
'==============================================================================
'                                  CalcColumn
'------------------------------------------------------------------------------
' PURPOSE
'   Return the Results or TradeCalc column of a measure, given by its
'   0-based position in NS_MEASURES or TR_MEASURES.
'
' UPDATED
'   2026-10-17
'==============================================================================
'
        If Kind = KIND_NS Then
            Select Case Measure
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
                    CalcColumn = RS_ADDON_IR + Measure - 4
            End Select
        Else
            CalcColumn = Choose(Measure + 1, TC_NOTIONAL, TC_MTM, TC_M, TC_SD, TC_ADJN, TC_DELTA, TC_MFU, TC_MFM, _
                                TC_ENU, TC_ENM, TC_ADDON)
        End If

End Function


Private Function BankFigure( _
    ByVal Figure As Variant, _
    ByVal IsRatio As Boolean, _
    ByVal ScaleFactor As Double) _
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
'   2026-10-17
'==============================================================================
'
        If IsRatio Then
            BankFigure = CDbl(Figure)
        Else
            BankFigure = CDbl(Figure) * ScaleFactor
        End If

End Function


Private Function BankAmount( _
    ByVal Amount As Variant, _
    ByVal ScaleFactor As Double) _
    As Variant
'
'==============================================================================
'                                  BankAmount
'------------------------------------------------------------------------------
' PURPOSE
'   Return a bank amount scaled to units, or Empty when it is not a number.
'
' UPDATED
'   2026-10-17
'==============================================================================
'
        If IsNumberValue(Amount) Then
            BankAmount = CDbl(Amount) * ScaleFactor
        End If

End Function


Private Function NumberOrEmpty( _
    ByVal Content As Variant) _
    As Variant
'
'==============================================================================
'                                NumberOrEmpty
'------------------------------------------------------------------------------
' PURPOSE
'   Return a Results figure as a number, or Empty when the cell has none.
'
' UPDATED
'   2026-10-17
'==============================================================================
'
        If IsNumberValue(Content) Then
            NumberOrEmpty = CDbl(Content)
        End If

End Function


Private Function IsNumberValue( _
    ByVal Content As Variant) _
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
'   2026-10-17
'==============================================================================
'
        Select Case VarType(Content)
            Case vbDouble, vbSingle, vbLong, vbInteger, vbCurrency, vbDecimal, vbByte
                IsNumberValue = True
        End Select

End Function


Private Function IsBlankBankCell( _
    ByVal Content As Variant) _
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
'   2026-10-17
'==============================================================================
'
        If IsError(Content) Then
            IsBlankBankCell = False
        Else
            IsBlankBankCell = (Len(SafeStr(Content)) = 0)
        End If

End Function


Private Function BankRowHasData( _
    ByRef Bank As Variant, _
    ByVal BankRow As Long, _
    ByVal CommentColumn As Long) _
    As Boolean
'
'==============================================================================
'                                BankRowHasData
'------------------------------------------------------------------------------
' PURPOSE
'   Tell whether a bank row has a value in any column before the comment
'   column CommentColumn.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim c   As Long    'Column

'------------------------------------------------------------------------------
' TEST
'------------------------------------------------------------------------------
        For c = 1 To CommentColumn - 1
            If Not IsBlankBankCell(Bank(BankRow, c)) Then
                BankRowHasData = True
                Exit Function
            End If
        Next c

End Function


Private Function DescribeValue( _
    ByVal Content As Variant) _
    As String
'
'==============================================================================
'                                DescribeValue
'------------------------------------------------------------------------------
' PURPOSE
'   Describe a bank value that is not a number, for a note.
'
' UPDATED
'   2026-10-17
'==============================================================================
'
        If IsError(Content) Then
            DescribeValue = "an error value"
        ElseIf VarType(Content) = vbBoolean Then
            DescribeValue = "the logical value " & UCase$(CStr(Content))
        Else
            DescribeValue = "the text '" & SafeStr(Content) & "'"
        End If

End Function


Private Function HeaderKey( _
    ByVal Header As Variant) _
    As String
'
'==============================================================================
'                                  HeaderKey
'------------------------------------------------------------------------------
' PURPOSE
'   Reduce a header to lower case without spaces or line breaks.
'
' UPDATED
'   2026-10-17
'==============================================================================
'
        HeaderKey = LCase$(Replace(Replace(Replace(SafeStr(Header), " ", ""), vbCr, ""), vbLf, ""))

End Function


Private Function IsTotalRow( _
    ByVal RowId As String) _
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
'   2026-10-17
'==============================================================================
'
        IsTotalRow = (RowId = "TOTAL" Or RowId = "TOTAL " & RG_CRR Or RowId = "TOTAL " & RG_BCBS)

End Function


Private Function RowCount( _
    ByVal Block As Variant) _
    As Long
'
'==============================================================================
'                                   RowCount
'------------------------------------------------------------------------------
' PURPOSE
'   Return the number of rows of a 2-D block; 0 when it is not an array.
'
' UPDATED
'   2026-10-17
'==============================================================================
'
        If IsArray(Block) Then
            RowCount = UBound(Block, 1) - LBound(Block, 1) + 1
        End If

End Function


Private Function TrimRows( _
    ByRef OutRows() As Variant, _
    ByVal OutCount As Long) _
    As Variant
'
'==============================================================================
'                                   TrimRows
'------------------------------------------------------------------------------
' PURPOSE
'   Return the first OutCount rows of OutRows as a new array.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Result()   As Variant    'The rows returned
    Dim r          As Long       'Row
    Dim c          As Long       'Column

'------------------------------------------------------------------------------
' COPY
'------------------------------------------------------------------------------
        ReDim Result(1 To OutCount, 1 To RC_NCOLS)
        For r = 1 To OutCount
            For c = 1 To RC_NCOLS
                Result(r, c) = OutRows(r, c)
            Next c
        Next r
        TrimRows = Result

End Function
