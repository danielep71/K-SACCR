Attribute VB_Name = "CORE_Dashboard"
'==============================================================================
' MODULE: CORE_Dashboard
'------------------------------------------------------------------------------
' PURPOSE
'   Compute the Dashboard's blocks (#675): the figures of the last run,
'   block by block, for the sheet's data table, which its charts and titles
'   read. Every figure is a count, a sum, a share or a ranking of what the
'   run wrote on Results, TradeCalc and Checks, of the inputs of NettingSets
'   and of the results of the two reconciliation sheets. No SA-CCR figure
'   is calculated here.
'
'   docs/workbook/DASHBOARD.md defines every block, and tools/dashboard.py
'   is its Python reference. Amounts use the netting sets of a regime whose
'   Status is VALID, and the TradeCalc legs that are OK in one of them;
'   each regime has its own column and amounts of the two are never added.
'
' PUBLIC SURFACE
'   None outside this VBA project. DashboardBlocks and BlockNames are
'   Public for M_Main and the tests; Option Private Module keeps them off
'   the supported external surface.
'
' DEPENDENCIES
'   CORE_Config for the columns, regimes, labels and codes; CORE_Util for
'   text and number conversion, Max2 and the Collection index (KeyIndex,
'   KeyAdd).
'
' STATE OWNERSHIP
'   None in the workbook: M_Main reads the sheets, passes their rows and
'   writes the blocks, each a workbook name. The rows are kept in module
'   arrays for one call of DashboardBlocks and released after it.
'
' ERROR POLICY
'   Unexpected errors are raised to the caller, which writes no block
'   unless every block was computed.
'
' COMPATIBILITY
'   Excel VBA; no references beyond the defaults.
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
    'Require explicit declarations; keep the writer inside this project.
    Option Explicit
    Option Private Module

'------------------------------------------------------------------------------
' BLOCKS
'------------------------------------------------------------------------------
    'Workbook names of the blocks, as docs/workbook/dashboard.json lists them.
        Private Const BLOCK_LIST As String = "Dashboard_Run|Dashboard_Scale|Dashboard_KeyFigures|" & _
            "Dashboard_Counts|Dashboard_Sector|Dashboard_ClearingRole|Dashboard_AddOn|Dashboard_Composition|" & _
            "Dashboard_TopCRR|Dashboard_TopBCBS|Dashboard_Concentration|Dashboard_Maturity|Dashboard_AssetClass|" & _
            "Dashboard_DataQuality|Dashboard_Codes|Dashboard_Reconciliation|Dashboard_NotionalToEAD|" & _
            "Dashboard_Reading"
    'Row keys of the blocks that are read by key.
        Private Const SECTOR_LIST As String = "CB|GG|CI|OFC|NFC|HH|NC"          'Counterparty sectors; NC not classified
        Private Const ROLE_LIST   As String = "N|CM|CLIENT|CCP_Q|CCP_NQ"        'Clearing roles, CCP split by QCCP
        Private Const AC_LIST     As String = "IR|FX|CR|EQ|CO|OT"               'Asset classes, in Results order
        Private Const RANK_ROWS   As Long = 10                                 'Rows of a ranking of netting sets
        Private Const CODE_ROWS   As Long = 5                                  'Rows of the diagnostic codes
        Private Const CLEAN_CODE  As String = "RUN-CLEAN"                      'The one INFO row of a clean run
        Private Const NO_CODE     As String = "(none)"                         'A Checks row without a code

'------------------------------------------------------------------------------
' RUN RECORDS
'------------------------------------------------------------------------------
    'Results: one entry per netting-set row, TOTAL rows left out.
        Private mNsCount      As Long          'Netting-set rows
        Private mNsId()       As String        'NettingSetID
        Private mNsCpty()     As String        'Counterparty, trimmed
        Private mNsRegime()   As String        'Regime, upper case
        Private mNsMargined() As Boolean       'Margined Y
        Private mNsV()        As Double        'V, the sum of the trades' MtM
        Private mNsRC()       As Double        'RC
        Private mNsAddOn()    As Double        'Add-on per asset class, (row, 1 To 6)
        Private mNsPFE()      As Double        'PFE
        Private mNsEAD()      As Double        'EAD
        Private mNsCap()      As Boolean       'Cap applied Y
        Private mNsStatus()   As String        'Status
        Private mNsSector()   As String        'Counterparty sector on NettingSets, or NC
        Private mNsRole()     As String        'Clearing role on NettingSets, CCP split by QCCP
    'TradeCalc: one entry per leg.
        Private mLegCount     As Long          'Legs
        Private mLegNs()      As String        'NettingSetID
        Private mLegAC()      As String        'Asset class, upper case
        Private mLegStatus()  As String        'OK or EXCLUDED
        Private mLegNotional() As Double       'Notional, reporting currency
        Private mLegMtm()     As Double        'MtM, reporting currency
        Private mLegM()       As Double        'Maturity M, years
        Private mLegRegime()  As String        'Regime, upper case
        Private mLegUsed()    As Long          'Regime position (1 CRR, 2 BCBS) when used, else 0
    'Checks: severity and code of every row.
        Private mCheckCount   As Long          'Rows
        Private mCheckSev()   As String        'ERROR, WARNING or INFO
        Private mCheckCode()  As String        'Diagnostic code
    'NettingSets: the input rows, read for the sector and the clearing role.
        Private mInputs       As Variant       'NettingSets inputs, 1-based; Empty when none
    'The reconciliation sheets and the run.
        Private mRecon        As Variant       'Reconciliation rows; Empty when none
        Private mTradeRecon   As Variant       'TradeReconciliation rows; Empty when none
        Private mRunInfo      As Variant       'Date, time, currency, regime, class, inputs

'------------------------------------------------------------------------------
' SCALES
'------------------------------------------------------------------------------
        Private mExposureFactor As Double      'Divisor of the exposure amounts
        Private mNotionalFactor As Double      'Divisor of the notional amounts


'
'------------------------------------------------------------------------------
'
'                                 ENTRY POINTS
'
'------------------------------------------------------------------------------
'

Public Function DashboardBlocks( _
    ByVal ResultsRows As Variant, _
    ByVal LegRows As Variant, _
    ByVal CheckRows As Variant, _
    ByVal InputRows As Variant, _
    ByVal ReconRows As Variant, _
    ByVal TradeReconRows As Variant, _
    ByVal RunInfo As Variant) _
    As Variant
'
'==============================================================================
'                               DashboardBlocks
'------------------------------------------------------------------------------
' PURPOSE
'   Compute every block of the Dashboard, in the order of BlockNames.
'
' INPUTS
'   ResultsRows, LegRows, CheckRows, InputRows: the rows of Results,
'   TradeCalc, Checks and NettingSets, 1-based two-dimensional arrays from
'   table column 1; Empty for a table without rows.
'   ReconRows, TradeReconRows: the rows of Reconciliation and
'   TradeReconciliation, at least RC_RESULT columns wide; Empty for none.
'   RunInfo: Array(reporting date, run time, reporting currency, default
'   regime, run class, input fingerprint).
'
' RETURNS
'   A zero-based array of the blocks' values, each a 1-based
'   two-dimensional array of the block's size.
'
' UPDATED
'   2026-11-23
'==============================================================================
'
    Dim BlockList   As Variant    'Block names
    Dim Values()    As Variant    'One array per block, in BlockList order
    Dim Index       As Long       'Block position
        On Error GoTo Failed
        BlockList = BlockNames()
        ReDim Values(LBound(BlockList) To UBound(BlockList))
        LoadRun ResultsRows, LegRows, CheckRows, InputRows
        mRecon = ReconRows
        mTradeRecon = TradeReconRows
        mRunInfo = RunInfo
        SetScales
        For Index = LBound(BlockList) To UBound(BlockList)
            Values(Index) = BlockValues(CStr(BlockList(Index)))
        Next Index
        ReleaseRun
        DashboardBlocks = Values
        Exit Function
Failed:
        ReleaseRun
        Err.Raise Err.Number, Err.Source, Err.Description

End Function


Public Function BlockNames() As Variant
'
'==============================================================================
'                                  BlockNames
'------------------------------------------------------------------------------
' PURPOSE
'   The workbook names of the blocks, in the order of dashboard.json.
'
' UPDATED
'   2026-11-23
'==============================================================================
'
        BlockNames = Split(BLOCK_LIST, "|")

End Function


'
'------------------------------------------------------------------------------
'
'                                 READ THE RUN
'
'------------------------------------------------------------------------------
'

Private Sub LoadRun( _
    ByVal ResultsRows As Variant, _
    ByVal LegRows As Variant, _
    ByVal CheckRows As Variant, _
    ByVal InputRows As Variant)
'
'==============================================================================
'                                   LoadRun
'------------------------------------------------------------------------------
' PURPOSE
'   Keep the rows of Results, TradeCalc and Checks in the module arrays,
'   and give each netting set its sector and clearing role from NettingSets
'   and each leg the regime position of the valid netting set it is used
'   in.
'
' UPDATED
'   2026-11-23
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Block        As Variant      'A table's values
    Dim RowIndex     As Long         'Row in the block
    Dim ClassIndex   As Long         'Asset class
    Dim Inputs       As Collection   'NettingSets row of each ID, first one kept
    Dim Valid        As Collection   'Valid netting sets: regime position by ID
    Dim Position     As Long         'A lookup's result

'------------------------------------------------------------------------------
' RESULTS
'------------------------------------------------------------------------------
        ReleaseRun
        mInputs = InputRows
        Block = ResultsRows
        Set Inputs = InputIndex()
        Set Valid = New Collection
        For RowIndex = 1 To RowCount(Block)
            If Not IsTotalRow(UTxt(Block(RowIndex, RS_ID))) Then
                mNsCount = mNsCount + 1
                ReDim Preserve mNsId(1 To mNsCount)
                ReDim Preserve mNsCpty(1 To mNsCount)
                ReDim Preserve mNsRegime(1 To mNsCount)
                ReDim Preserve mNsMargined(1 To mNsCount)
                ReDim Preserve mNsV(1 To mNsCount)
                ReDim Preserve mNsRC(1 To mNsCount)
                ReDim Preserve mNsPFE(1 To mNsCount)
                ReDim Preserve mNsEAD(1 To mNsCount)
                ReDim Preserve mNsCap(1 To mNsCount)
                ReDim Preserve mNsStatus(1 To mNsCount)
                ReDim Preserve mNsSector(1 To mNsCount)
                ReDim Preserve mNsRole(1 To mNsCount)
                mNsId(mNsCount) = Trim$(SafeStr(Block(RowIndex, RS_ID)))
                mNsCpty(mNsCount) = Trim$(SafeStr(Block(RowIndex, RS_CPTY)))
                mNsRegime(mNsCount) = UTxt(Block(RowIndex, RS_REGIME))
                mNsMargined(mNsCount) = (UTxt(Block(RowIndex, RS_MARGINED)) = "Y")
                mNsV(mNsCount) = ToDbl(Block(RowIndex, RS_V), 0#)
                mNsRC(mNsCount) = ToDbl(Block(RowIndex, RS_RC), 0#)
                mNsPFE(mNsCount) = ToDbl(Block(RowIndex, RS_PFE), 0#)
                mNsEAD(mNsCount) = ToDbl(Block(RowIndex, RS_EAD), 0#)
                mNsCap(mNsCount) = (UTxt(Block(RowIndex, RS_CAP)) = "Y")
                mNsStatus(mNsCount) = Trim$(SafeStr(Block(RowIndex, RS_STATUS_COL)))
                Position = KeyIndex(Inputs, mNsId(mNsCount))
                mNsSector(mNsCount) = SectorOf(Position)
                mNsRole(mNsCount) = RoleOf(Position)
                If mNsStatus(mNsCount) = "VALID" And RegimePosition(mNsRegime(mNsCount)) > 0 Then
                    If KeyIndex(Valid, mNsId(mNsCount)) = 0 Then
                        KeyAdd Valid, mNsId(mNsCount), RegimePosition(mNsRegime(mNsCount))
                    End If
                End If
            End If
        Next RowIndex
        If mNsCount > 0 Then
            ReDim mNsAddOn(1 To mNsCount, 1 To AC_COUNT)
            Position = 0
            For RowIndex = 1 To RowCount(Block)
                If Not IsTotalRow(UTxt(Block(RowIndex, RS_ID))) Then
                    Position = Position + 1
                    For ClassIndex = 1 To AC_COUNT
                        mNsAddOn(Position, ClassIndex) = ToDbl(Block(RowIndex, RS_ADDON_IR + ClassIndex - 1), 0#)
                    Next ClassIndex
                End If
            Next RowIndex
        End If

'------------------------------------------------------------------------------
' TRADECALC
'------------------------------------------------------------------------------
        Block = LegRows
        mLegCount = RowCount(Block)
        If mLegCount > 0 Then
            ReDim mLegNs(1 To mLegCount)
            ReDim mLegAC(1 To mLegCount)
            ReDim mLegStatus(1 To mLegCount)
            ReDim mLegNotional(1 To mLegCount)
            ReDim mLegMtm(1 To mLegCount)
            ReDim mLegM(1 To mLegCount)
            ReDim mLegRegime(1 To mLegCount)
            ReDim mLegUsed(1 To mLegCount)
            For RowIndex = 1 To mLegCount
                mLegNs(RowIndex) = Trim$(SafeStr(Block(RowIndex, TC_NS)))
                mLegAC(RowIndex) = UTxt(Block(RowIndex, TC_AC))
                mLegStatus(RowIndex) = UTxt(Block(RowIndex, TC_STATUS))
                mLegNotional(RowIndex) = ToDbl(Block(RowIndex, TC_NOTIONAL), 0#)
                mLegMtm(RowIndex) = ToDbl(Block(RowIndex, TC_MTM), 0#)
                mLegM(RowIndex) = ToDbl(Block(RowIndex, TC_M), 0#)
                mLegRegime(RowIndex) = UTxt(Block(RowIndex, TC_REGIME))
                If mLegStatus(RowIndex) = "OK" Then
                    Position = KeyIndex(Valid, mLegNs(RowIndex))
                    If Position > 0 And Position = RegimePosition(mLegRegime(RowIndex)) Then
                        mLegUsed(RowIndex) = Position
                    End If
                End If
            Next RowIndex
        End If

'------------------------------------------------------------------------------
' CHECKS
'------------------------------------------------------------------------------
        Block = CheckRows
        mCheckCount = RowCount(Block)
        If mCheckCount > 0 Then
            ReDim mCheckSev(1 To mCheckCount)
            ReDim mCheckCode(1 To mCheckCount)
            For RowIndex = 1 To mCheckCount
                mCheckSev(RowIndex) = UTxt(Block(RowIndex, 1))
                mCheckCode(RowIndex) = Trim$(SafeStr(Block(RowIndex, CK_CODE_COL)))
            Next RowIndex
        End If

End Sub


Private Sub ReleaseRun()
'
'==============================================================================
'                                  ReleaseRun
'------------------------------------------------------------------------------
' PURPOSE
'   Empty the module arrays.
'
' UPDATED
'   2026-11-23
'==============================================================================
'
        mNsCount = 0
        Erase mNsId, mNsCpty, mNsRegime, mNsMargined, mNsV, mNsRC, mNsAddOn, mNsPFE, mNsEAD, mNsCap, _
              mNsStatus, mNsSector, mNsRole
        mLegCount = 0
        Erase mLegNs, mLegAC, mLegStatus, mLegNotional, mLegMtm, mLegM, mLegRegime, mLegUsed
        mCheckCount = 0
        Erase mCheckSev, mCheckCode
        mInputs = Empty
        mRecon = Empty
        mTradeRecon = Empty
        mRunInfo = Empty

End Sub


Private Function RowCount( _
    ByVal Block As Variant) _
    As Long
'
'==============================================================================
'                                   RowCount
'------------------------------------------------------------------------------
' PURPOSE
'   Rows of a table passed in; 0 for Empty.
'
' UPDATED
'   2026-11-23
'==============================================================================
'
        If IsArray(Block) Then RowCount = UBound(Block, 1)

End Function


Private Function InputIndex() As Collection
'
'==============================================================================
'                                  InputIndex
'------------------------------------------------------------------------------
' PURPOSE
'   The NettingSets row of each netting-set ID, the first one where an ID
'   is repeated (such a netting set is invalid and has no EAD).
'
' UPDATED
'   2026-11-23
'==============================================================================
'
    Dim Block      As Variant      'NettingSets inputs
    Dim RowIndex   As Long         'Row in the block
    Dim Lookup     As Collection   'Row by ID
        Set Lookup = New Collection
        Block = mInputs
        For RowIndex = 1 To RowCount(Block)
            If Len(Trim$(SafeStr(Block(RowIndex, NS_ID)))) > 0 Then
                If KeyIndex(Lookup, Trim$(SafeStr(Block(RowIndex, NS_ID)))) = 0 Then
                    KeyAdd Lookup, Trim$(SafeStr(Block(RowIndex, NS_ID))), RowIndex
                End If
            End If
        Next RowIndex
        Set InputIndex = Lookup

End Function


Private Function SectorOf( _
    ByVal InputRow As Long) _
    As String
'
'==============================================================================
'                                   SectorOf
'------------------------------------------------------------------------------
' PURPOSE
'   The counterparty sector of a NettingSets row: one of CB, GG, CI, OFC,
'   NFC and HH, or NC when it is blank, unknown or the row is not found.
'
' UPDATED
'   2026-11-23
'==============================================================================
'
    Dim Code   As String    'The sector entered
        SectorOf = "NC"
        If InputRow = 0 Then Exit Function
        Code = UTxt(mInputs(InputRow, NS_SECTOR))
        If Len(Code) > 0 And Code <> "NC" And InStr(1, "|" & SECTOR_LIST & "|", "|" & Code & "|", vbBinaryCompare) > 0 Then
            SectorOf = Code
        End If

End Function


Private Function RoleOf( _
    ByVal InputRow As Long) _
    As String
'
'==============================================================================
'                                    RoleOf
'------------------------------------------------------------------------------
' PURPOSE
'   The clearing role of a NettingSets row: N for a blank or N role, CCP_Q
'   or CCP_NQ for a CCP by its QCCP flag, else the role entered.
'
' UPDATED
'   2026-11-23
'==============================================================================
'
    Dim Code   As String    'The role entered
        RoleOf = "N"
        If InputRow = 0 Then Exit Function
        Code = UTxt(mInputs(InputRow, NS_CLEARED))
        Select Case Code
            Case "", "N"
                RoleOf = "N"
            Case "CCP"
                RoleOf = IIf(UTxt(mInputs(InputRow, NS_QCCP)) = "Y", "CCP_Q", "CCP_NQ")
            Case Else
                RoleOf = Code
        End Select

End Function


Private Function IsTotalRow( _
    ByVal RowId As String) _
    As Boolean
'
'==============================================================================
'                                  IsTotalRow
'------------------------------------------------------------------------------
' PURPOSE
'   Tell whether a Results ID is one of the TOTAL rows the engine reserves
'   (#198), which are not netting sets.
'
' UPDATED
'   2026-11-23
'==============================================================================
'
        IsTotalRow = (RowId = "TOTAL" Or RowId = "TOTAL " & RG_CRR Or RowId = "TOTAL " & RG_BCBS)

End Function


Private Function RegimePosition( _
    ByVal Regime As String) _
    As Long
'
'==============================================================================
'                                RegimePosition
'------------------------------------------------------------------------------
' PURPOSE
'   1 for CRR, 2 for BCBS, 0 otherwise: the column of a regime in a block.
'
' UPDATED
'   2026-11-23
'==============================================================================
'
        Select Case Regime
            Case RG_CRR
                RegimePosition = 1
            Case RG_BCBS
                RegimePosition = 2
            Case Else
                RegimePosition = 0
        End Select

End Function


'
'------------------------------------------------------------------------------
'
'                                   FIGURES
'
'------------------------------------------------------------------------------
'

Private Function IsValid( _
    ByVal NsIndex As Long, _
    ByVal Regime As Long) _
    As Boolean
'
'==============================================================================
'                                   IsValid
'------------------------------------------------------------------------------
' PURPOSE
'   Tell whether a netting set is VALID and of the regime at this position.
'
' UPDATED
'   2026-11-23
'==============================================================================
'
        IsValid = (mNsStatus(NsIndex) = "VALID" And RegimePosition(mNsRegime(NsIndex)) = Regime)

End Function


Private Function ValidSum( _
    ByVal Regime As Long, _
    ByVal Figure As String, _
    Optional ByVal Dimension As String = "", _
    Optional ByVal Key As String = "") _
    As Double
'
'==============================================================================
'                                   ValidSum
'------------------------------------------------------------------------------
' PURPOSE
'   Sum a figure (EAD, RC, PFE, VPOS for max(V, 0), or an asset class code
'   for its add-on) over the valid netting sets of a regime, optionally only
'   those of a sector, a clearing role or a margined flag (Dimension SECTOR,
'   ROLE or MARGINED, with Key the value, Y or N for MARGINED).
'
' UPDATED
'   2026-11-23
'==============================================================================
'
    Dim NsIndex   As Long      'Netting set
    Dim Total     As Double    'Running sum
    Dim Keep      As Boolean   'The filter's answer
        For NsIndex = 1 To mNsCount
            If IsValid(NsIndex, Regime) Then
                Select Case Dimension
                    Case ""
                        Keep = True
                    Case "SECTOR"
                        Keep = (mNsSector(NsIndex) = Key)
                    Case "ROLE"
                        Keep = (mNsRole(NsIndex) = Key)
                    Case "MARGINED"
                        Keep = (mNsMargined(NsIndex) = (Key = "Y"))
                    Case Else
                        Keep = False
                End Select
                If Keep Then Total = Total + NsFigure(NsIndex, Figure)
            End If
        Next NsIndex
        ValidSum = Total

End Function


Private Function NsFigure( _
    ByVal NsIndex As Long, _
    ByVal Figure As String) _
    As Double
'
'==============================================================================
'                                   NsFigure
'------------------------------------------------------------------------------
' PURPOSE
'   One figure of a netting set: EAD, RC, PFE, VPOS (max(V, 0)) or the
'   add-on of an asset class by its code.
'
' UPDATED
'   2026-11-23
'==============================================================================
'
        Select Case Figure
            Case "EAD"
                NsFigure = mNsEAD(NsIndex)
            Case "RC"
                NsFigure = mNsRC(NsIndex)
            Case "PFE"
                NsFigure = mNsPFE(NsIndex)
            Case "VPOS"
                NsFigure = IIf(mNsV(NsIndex) > 0#, mNsV(NsIndex), 0#)
            Case Else
                NsFigure = mNsAddOn(NsIndex, ACPosition(Figure))
        End Select

End Function


Private Function ACPosition( _
    ByVal Code As String) _
    As Long
'
'==============================================================================
'                                  ACPosition
'------------------------------------------------------------------------------
' PURPOSE
'   Position of an asset-class code in AC_LIST (1 for IR to 6 for OT); 0
'   for another code.
'
' UPDATED
'   2026-11-23
'==============================================================================
'
    Dim Codes   As Variant    'The codes
    Dim Index   As Long       'Position
        Codes = Split(AC_LIST, "|")
        For Index = 0 To UBound(Codes)
            If Codes(Index) = Code Then
                ACPosition = Index + 1
                Exit Function
            End If
        Next Index

End Function


Private Function LegSum( _
    ByVal Regime As Long, _
    ByVal Figure As String, _
    Optional ByVal Bucket As String = "", _
    Optional ByVal Code As String = "") _
    As Double
'
'==============================================================================
'                                    LegSum
'------------------------------------------------------------------------------
' PURPOSE
'   Sum the notional (NOTIONAL) or the positive MtM (MTMPOS) of the legs
'   used in the valid netting sets of a regime, optionally of one maturity
'   band (Bucket under_1y, 1y_5y or over_5y) or one asset class (Code).
'
' UPDATED
'   2026-11-23
'==============================================================================
'
    Dim LegIndex   As Long      'Leg
    Dim Total      As Double    'Running sum
        For LegIndex = 1 To mLegCount
            If mLegUsed(LegIndex) = Regime Then
                If (Len(Bucket) = 0 Or MaturityBand(mLegM(LegIndex)) = Bucket) And _
                   (Len(Code) = 0 Or mLegAC(LegIndex) = Code) Then
                    If Figure = "MTMPOS" Then
                        If mLegMtm(LegIndex) > 0# Then Total = Total + mLegMtm(LegIndex)
                    Else
                        Total = Total + mLegNotional(LegIndex)
                    End If
                End If
            End If
        Next LegIndex
        LegSum = Total

End Function


Private Function LegCount( _
    ByVal Regime As Long, _
    ByVal Code As String) _
    As Long
'
'==============================================================================
'                                   LegCount
'------------------------------------------------------------------------------
' PURPOSE
'   Legs used in the valid netting sets of a regime, of one asset class.
'
' UPDATED
'   2026-11-23
'==============================================================================
'
    Dim LegIndex   As Long    'Leg
        For LegIndex = 1 To mLegCount
            If mLegUsed(LegIndex) = Regime And mLegAC(LegIndex) = Code Then LegCount = LegCount + 1
        Next LegIndex

End Function


Private Function MaturityBand( _
    ByVal M As Double) _
    As String
'
'==============================================================================
'                                 MaturityBand
'------------------------------------------------------------------------------
' PURPOSE
'   The band of a maturity: under_1y (M < 1), 1y_5y (1 <= M <= 5) or
'   over_5y (M > 5).
'
' UPDATED
'   2026-11-23
'==============================================================================
'
        If M < 1# Then
            MaturityBand = "under_1y"
        ElseIf M <= 5# Then
            MaturityBand = "1y_5y"
        Else
            MaturityBand = "over_5y"
        End If

End Function


Private Sub SetScales()
'
'==============================================================================
'                                  SetScales
'------------------------------------------------------------------------------
' PURPOSE
'   Choose the divisor of each family of amounts from its largest absolute
'   amount over both regimes: 1, 1,000, 1,000,000 or 1,000,000,000, the
'   largest not above it (1 when every amount is 0).
'
' UPDATED
'   2026-11-23
'==============================================================================
'
    Dim Regime     As Long      'Regime position
    Dim Exposure   As Double    'Largest exposure amount
    Dim Notional   As Double    'Largest notional amount
    Dim Key        As Variant   'A row key
        For Regime = 1 To 2
            Exposure = Max2(Exposure, Abs(ValidSum(Regime, "EAD")))
            Exposure = Max2(Exposure, Abs(ValidSum(Regime, "RC")))
            Exposure = Max2(Exposure, Abs(ValidSum(Regime, "PFE")))
            Exposure = Max2(Exposure, Abs(ValidSum(Regime, "VPOS")))
            Exposure = Max2(Exposure, Abs(LegSum(Regime, "MTMPOS")))
            For Each Key In Split(AC_LIST, "|")
                Exposure = Max2(Exposure, Abs(ValidSum(Regime, CStr(Key))))
                Notional = Max2(Notional, Abs(LegSum(Regime, "NOTIONAL", "", CStr(Key))))
            Next Key
            For Each Key In Split(SECTOR_LIST, "|")
                Exposure = Max2(Exposure, Abs(ValidSum(Regime, "EAD", "SECTOR", CStr(Key))))
            Next Key
            For Each Key In Split(ROLE_LIST, "|")
                Exposure = Max2(Exposure, Abs(ValidSum(Regime, "EAD", "ROLE", CStr(Key))))
            Next Key
            Exposure = Max2(Exposure, Abs(ValidSum(Regime, "EAD", "MARGINED", "Y")))
            Exposure = Max2(Exposure, Abs(ValidSum(Regime, "EAD", "MARGINED", "N")))
            Exposure = Max2(Exposure, LargestEAD(Regime))
            Notional = Max2(Notional, Abs(LegSum(Regime, "NOTIONAL")))
            For Each Key In Array("under_1y", "1y_5y", "over_5y")
                Notional = Max2(Notional, Abs(LegSum(Regime, "NOTIONAL", CStr(Key))))
            Next Key
        Next Regime
        mExposureFactor = ScaleFactor(Exposure)
        mNotionalFactor = ScaleFactor(Notional)

End Sub


Private Function LargestEAD( _
    ByVal Regime As Long) _
    As Double
'
'==============================================================================
'                                  LargestEAD
'------------------------------------------------------------------------------
' PURPOSE
'   The largest absolute EAD of a valid netting set of a regime.
'
' UPDATED
'   2026-11-23
'==============================================================================
'
    Dim NsIndex   As Long    'Netting set
        For NsIndex = 1 To mNsCount
            If IsValid(NsIndex, Regime) Then LargestEAD = Max2(LargestEAD, Abs(mNsEAD(NsIndex)))
        Next NsIndex

End Function


Private Function ScaleFactor( _
    ByVal Largest As Double) _
    As Double
'
'==============================================================================
'                                 ScaleFactor
'------------------------------------------------------------------------------
' PURPOSE
'   1, 1E3, 1E6 or 1E9: the largest not above Largest, and 1 below 1,000.
'
' UPDATED
'   2026-11-23
'==============================================================================
'
        If Largest >= 1000000000# Then
            ScaleFactor = 1000000000#
        ElseIf Largest >= 1000000# Then
            ScaleFactor = 1000000#
        ElseIf Largest >= 1000# Then
            ScaleFactor = 1000#
        Else
            ScaleFactor = 1#
        End If

End Function


Private Function UnitText( _
    ByVal Factor As Double, _
    ByVal CcyCode As String) _
    As String
'
'==============================================================================
'                                   UnitText
'------------------------------------------------------------------------------
' PURPOSE
'   The unit of a scale: the currency, then thousand, million or billion.
'
' UPDATED
'   2026-11-23
'==============================================================================
'
        Select Case Factor
            Case 1000#
                UnitText = CcyCode & " thousand"
            Case 1000000#
                UnitText = CcyCode & " million"
            Case 1000000000#
                UnitText = CcyCode & " billion"
            Case Else
                UnitText = CcyCode
        End Select

End Function


Private Function ShareOf( _
    ByVal Part As Double, _
    ByVal Whole As Double) _
    As Variant
'
'==============================================================================
'                                   ShareOf
'------------------------------------------------------------------------------
' PURPOSE
'   Part / Whole; Empty when Whole is 0.
'
' UPDATED
'   2026-11-23
'==============================================================================
'
        If Whole = 0# Then
            ShareOf = Empty
        Else
            ShareOf = Part / Whole
        End If

End Function


Private Function PercentText( _
    ByVal Fraction As Double) _
    As String
'
'==============================================================================
'                                 PercentText
'------------------------------------------------------------------------------
' PURPOSE
'   A non-negative fraction as a percentage with one decimal, halves away
'   from zero, with a full stop in every locale: 0.3846 gives 38.5%.
'
' NOTES
'   The fraction is taken as a Decimal first, from its 15 significant
'   digits, so that 0.0045 is the half 0.45% and not 0.44999...%.
'
' UPDATED
'   2026-11-23
'==============================================================================
'
    Dim Tenths   As Variant    'The percentage in tenths, a Decimal
        Tenths = Int(CDec(Fraction) * 1000 + CDec(0.5))
        PercentText = CStr(Fix(Tenths / 10)) & "." & CStr(Tenths - Fix(Tenths / 10) * 10) & "%"

End Function


Private Function ReadingText( _
    ByVal Gross As Double, _
    ByVal Netted As Double, _
    ByVal RC As Double) _
    As String
'
'==============================================================================
'                                 ReadingText
'------------------------------------------------------------------------------
' PURPOSE
'   The one-line reading of view 11: how much of the gross positive value
'   close-out netting removes, and how much collateral removes further or,
'   when RC exceeds the value after netting, how much collateral and margin
'   terms add.
'
' UPDATED
'   2026-11-23
'==============================================================================
'
    Dim Collateral   As Double    '(netted - RC) / gross
        If Gross <= 0# Then
            ReadingText = "No trade has a positive market value."
            Exit Function
        End If
        ReadingText = "Netting removes " & PercentText((Gross - Netted) / Gross) & " of gross positive value"
        Collateral = (Netted - RC) / Gross
        If Collateral >= 0# Then
            ReadingText = ReadingText & "; collateral a further " & PercentText(Collateral) & "."
        Else
            ReadingText = ReadingText & "; collateral and margin terms add " & PercentText(-Collateral) & "."
        End If

End Function


'
'------------------------------------------------------------------------------
'
'                                    BLOCKS
'
'------------------------------------------------------------------------------
'

Private Function BlockValues( _
    ByVal BlockName As String) _
    As Variant
'
'==============================================================================
'                                 BlockValues
'------------------------------------------------------------------------------
' PURPOSE
'   The values of one block, a 1-based two-dimensional array of its size.
'
' UPDATED
'   2026-11-23
'==============================================================================
'
        Select Case BlockName
            Case "Dashboard_Run"
                BlockValues = RunBlock()
            Case "Dashboard_Scale"
                BlockValues = ScaleBlock()
            Case "Dashboard_KeyFigures"
                BlockValues = KeyFiguresBlock()
            Case "Dashboard_Counts"
                BlockValues = CountsBlock()
            Case "Dashboard_Sector"
                BlockValues = SplitBlock(SECTOR_LIST, "SECTOR")
            Case "Dashboard_ClearingRole"
                BlockValues = SplitBlock(ROLE_LIST, "ROLE")
            Case "Dashboard_AddOn"
                BlockValues = AddOnBlock()
            Case "Dashboard_Composition"
                BlockValues = CompositionBlock()
            Case "Dashboard_TopCRR"
                BlockValues = TopBlock(1)
            Case "Dashboard_TopBCBS"
                BlockValues = TopBlock(2)
            Case "Dashboard_Concentration"
                BlockValues = ConcentrationBlock()
            Case "Dashboard_Maturity"
                BlockValues = MaturityBlock()
            Case "Dashboard_AssetClass"
                BlockValues = AssetClassBlock()
            Case "Dashboard_DataQuality"
                BlockValues = DataQualityBlock()
            Case "Dashboard_Codes"
                BlockValues = CodesBlock()
            Case "Dashboard_Reconciliation"
                BlockValues = ReconciliationBlock()
            Case "Dashboard_NotionalToEAD"
                BlockValues = PathBlock()
            Case "Dashboard_Reading"
                BlockValues = ReadingBlock()
            Case Else
                Err.Raise ERR_OUTPUT_LAYOUT, "CORE_Dashboard.BlockValues", "Unknown block " & BlockName & "."
        End Select

End Function


Private Function RunBlock() As Variant
'
'==============================================================================
'                                   RunBlock
'------------------------------------------------------------------------------
' PURPOSE
'   Reporting date, run time, reporting currency, default regime,
'   parameter set and input fingerprint.
'
' UPDATED
'   2026-11-23
'==============================================================================
'
    Dim Values(1 To 6, 1 To 1)   As Variant    'The block
        Values(1, 1) = mRunInfo(0)
        Values(2, 1) = mRunInfo(1)
        Values(3, 1) = Trim$(SafeStr(mRunInfo(2)))
        Values(4, 1) = UTxt(mRunInfo(3))
        If SafeStr(mRunInfo(4)) = RUN_REGULATORY Then
            Values(5, 1) = RUN_REGULATORY & " " & REF_SET_ID
        Else
            Values(5, 1) = SafeStr(mRunInfo(4))
        End If
        Values(6, 1) = SafeStr(mRunInfo(5))
        RunBlock = Values

End Function


Private Function ScaleBlock() As Variant
'
'==============================================================================
'                                  ScaleBlock
'------------------------------------------------------------------------------
' PURPOSE
'   Factor and unit of the exposure and of the notional amounts.
'
' UPDATED
'   2026-11-23
'==============================================================================
'
    Dim Values(1 To 2, 1 To 2)   As Variant    'The block
    Dim CcyCode                 As String     'Reporting currency
        CcyCode = Trim$(SafeStr(mRunInfo(2)))
        Values(1, 1) = mExposureFactor
        Values(1, 2) = UnitText(mExposureFactor, CcyCode)
        Values(2, 1) = mNotionalFactor
        Values(2, 2) = UnitText(mNotionalFactor, CcyCode)
        ScaleBlock = Values

End Function


Private Function KeyFiguresBlock() As Variant
'
'==============================================================================
'                               KeyFiguresBlock
'------------------------------------------------------------------------------
' PURPOSE
'   Per regime: EAD, RC and PFE of the valid netting sets; netting sets in
'   total, valid, invalid, incomplete, with no trades, margined and
'   unmargined; valid ones with the cap applied; distinct counterparties.
'
' UPDATED
'   2026-11-23
'==============================================================================
'
    Dim Values(1 To 12, 1 To 2)   As Variant    'The block
    Dim Regime                    As Long       'Regime position
    Dim NsIndex                   As Long       'Netting set
    Dim Counts(4 To 11)           As Long       'Counts of rows 4 to 11
        For Regime = 1 To 2
            Erase Counts
            For NsIndex = 1 To mNsCount
                If RegimePosition(mNsRegime(NsIndex)) = Regime Then
                    Counts(4) = Counts(4) + 1
                    If mNsStatus(NsIndex) = "VALID" Then Counts(5) = Counts(5) + 1
                    If Left$(mNsStatus(NsIndex), 7) = "INVALID" Then Counts(6) = Counts(6) + 1
                    If Left$(mNsStatus(NsIndex), 10) = "INCOMPLETE" Then Counts(7) = Counts(7) + 1
                    If mNsStatus(NsIndex) = "NO TRADES" Then Counts(8) = Counts(8) + 1
                    If mNsMargined(NsIndex) Then Counts(9) = Counts(9) + 1 Else Counts(10) = Counts(10) + 1
                    If mNsStatus(NsIndex) = "VALID" And mNsCap(NsIndex) Then Counts(11) = Counts(11) + 1
                End If
            Next NsIndex
            Values(1, Regime) = ValidSum(Regime, "EAD") / mExposureFactor
            Values(2, Regime) = ValidSum(Regime, "RC") / mExposureFactor
            Values(3, Regime) = ValidSum(Regime, "PFE") / mExposureFactor
            For NsIndex = 4 To 11
                Values(NsIndex, Regime) = Counts(NsIndex)
            Next NsIndex
            Values(12, Regime) = Counterparties(Regime)
        Next Regime
        KeyFiguresBlock = Values

End Function


Private Function Counterparties( _
    ByVal Regime As Long) _
    As Long
'
'==============================================================================
'                                Counterparties
'------------------------------------------------------------------------------
' PURPOSE
'   Distinct non-blank counterparty names, compared trimmed and ignoring
'   case, of the netting sets of a regime; of every netting set for 0.
'
' UPDATED
'   2026-11-23
'==============================================================================
'
    Dim Seen      As Collection   'Names already counted
    Dim NsIndex   As Long         'Netting set
        Set Seen = New Collection
        For NsIndex = 1 To mNsCount
            If Regime = 0 Or RegimePosition(mNsRegime(NsIndex)) = Regime Then
                If Len(mNsCpty(NsIndex)) > 0 Then
                    If KeyIndex(Seen, mNsCpty(NsIndex)) = 0 Then
                        KeyAdd Seen, mNsCpty(NsIndex), 1
                        Counterparties = Counterparties + 1
                    End If
                End If
            End If
        Next NsIndex

End Function


Private Function CountsBlock() As Variant
'
'==============================================================================
'                                 CountsBlock
'------------------------------------------------------------------------------
' PURPOSE
'   Legs read, used and excluded; netting sets without a regime; distinct
'   counterparties; Checks rows by severity without the clean run's row.
'
' UPDATED
'   2026-11-23
'==============================================================================
'
    Dim Values(1 To 8, 1 To 1)   As Variant    'The block
    Dim Index                    As Long       'Leg, netting set or Checks row
        Values(1, 1) = mLegCount
        Values(2, 1) = 0
        Values(3, 1) = 0
        For Index = 1 To mLegCount
            If mLegStatus(Index) = "OK" Then Values(2, 1) = Values(2, 1) + 1
            If mLegStatus(Index) = "EXCLUDED" Then Values(3, 1) = Values(3, 1) + 1
        Next Index
        Values(4, 1) = 0
        For Index = 1 To mNsCount
            If RegimePosition(mNsRegime(Index)) = 0 Then Values(4, 1) = Values(4, 1) + 1
        Next Index
        Values(5, 1) = Counterparties(0)
        Values(6, 1) = 0
        Values(7, 1) = 0
        Values(8, 1) = 0
        For Index = 1 To mCheckCount
            If mCheckCode(Index) <> CLEAN_CODE Then
                Select Case mCheckSev(Index)
                    Case SEV_ERROR
                        Values(6, 1) = Values(6, 1) + 1
                    Case SEV_WARN
                        Values(7, 1) = Values(7, 1) + 1
                    Case SEV_INFO
                        Values(8, 1) = Values(8, 1) + 1
                    Case Else
                End Select
            End If
        Next Index
        CountsBlock = Values

End Function


Private Function SplitBlock( _
    ByVal Keys As String, _
    ByVal Dimension As String) _
    As Variant
'
'==============================================================================
'                                  SplitBlock
'------------------------------------------------------------------------------
' PURPOSE
'   EAD of the valid netting sets per regime, split by sector or clearing
'   role (views 1 and 2).
'
' UPDATED
'   2026-11-23
'==============================================================================
'
    Dim KeyList   As Variant    'Row keys
    Dim Values()  As Variant    'The block
    Dim Index     As Long       'Row
    Dim Regime    As Long       'Regime position
        KeyList = Split(Keys, "|")
        ReDim Values(1 To UBound(KeyList) + 1, 1 To 2)
        For Index = 0 To UBound(KeyList)
            For Regime = 1 To 2
                Values(Index + 1, Regime) = ValidSum(Regime, "EAD", Dimension, CStr(KeyList(Index))) / mExposureFactor
            Next Regime
        Next Index
        SplitBlock = Values

End Function


Private Function AddOnBlock() As Variant
'
'==============================================================================
'                                  AddOnBlock
'------------------------------------------------------------------------------
' PURPOSE
'   Add-on of the valid netting sets per asset class and regime (view 3).
'
' UPDATED
'   2026-11-23
'==============================================================================
'
    Dim Codes                    As Variant    'Asset classes
    Dim Values(1 To 6, 1 To 2)   As Variant    'The block
    Dim Index                    As Long       'Row
    Dim Regime                   As Long       'Regime position
        Codes = Split(AC_LIST, "|")
        For Index = 0 To 5
            For Regime = 1 To 2
                Values(Index + 1, Regime) = ValidSum(Regime, CStr(Codes(Index))) / mExposureFactor
            Next Regime
        Next Index
        AddOnBlock = Values

End Function


Private Function CompositionBlock() As Variant
'
'==============================================================================
'                               CompositionBlock
'------------------------------------------------------------------------------
' PURPOSE
'   RC, PFE, and EAD of the margined and of the unmargined valid netting
'   sets, per regime (view 4).
'
' UPDATED
'   2026-11-23
'==============================================================================
'
    Dim Values(1 To 4, 1 To 2)   As Variant    'The block
    Dim Regime                   As Long       'Regime position
        For Regime = 1 To 2
            Values(1, Regime) = ValidSum(Regime, "RC") / mExposureFactor
            Values(2, Regime) = ValidSum(Regime, "PFE") / mExposureFactor
            Values(3, Regime) = ValidSum(Regime, "EAD", "MARGINED", "Y") / mExposureFactor
            Values(4, Regime) = ValidSum(Regime, "EAD", "MARGINED", "N") / mExposureFactor
        Next Regime
        CompositionBlock = Values

End Function


Private Function Ranked( _
    ByVal Regime As Long) _
    As Long()
'
'==============================================================================
'                                    Ranked
'------------------------------------------------------------------------------
' PURPOSE
'   The valid netting sets of a regime by EAD, largest first; equal EADs by
'   NettingSetID in ordinal (binary) order. Returns indexes 1 To n, or an
'   array 0 To 0 holding 0 when there is none.
'
' UPDATED
'   2026-11-23
'==============================================================================
'
    Dim Order()   As Long      'Indexes, sorted in place
    Dim Count     As Long      'Valid netting sets
    Dim NsIndex   As Long      'Netting set
    Dim Outer     As Long      'Insertion position
    Dim Inner     As Long      'Comparison position
    Dim Moving    As Long      'The index being inserted
        ReDim Order(0 To 0)
        For NsIndex = 1 To mNsCount
            If IsValid(NsIndex, Regime) Then
                Count = Count + 1
                ReDim Preserve Order(0 To Count)
                Order(Count) = NsIndex
            End If
        Next NsIndex
        For Outer = 2 To Count
            Moving = Order(Outer)
            Inner = Outer - 1
            Do While Inner >= 1
                If Not Before(Moving, Order(Inner)) Then Exit Do
                Order(Inner + 1) = Order(Inner)
                Inner = Inner - 1
            Loop
            Order(Inner + 1) = Moving
        Next Outer
        Ranked = Order

End Function


Private Function Before( _
    ByVal NsA As Long, _
    ByVal NsB As Long) _
    As Boolean
'
'==============================================================================
'                                    Before
'------------------------------------------------------------------------------
' PURPOSE
'   Tell whether netting set First ranks before Second: a larger EAD, or an
'   equal EAD and a lower NettingSetID in binary order.
'
' UPDATED
'   2026-11-23
'==============================================================================
'
        If mNsEAD(NsA) <> mNsEAD(NsB) Then
            Before = (mNsEAD(NsA) > mNsEAD(NsB))
        Else
            Before = (StrComp(mNsId(NsA), mNsId(NsB), vbBinaryCompare) < 0)
        End If

End Function


Private Function TopBlock( _
    ByVal Regime As Long) _
    As Variant
'
'==============================================================================
'                                   TopBlock
'------------------------------------------------------------------------------
' PURPOSE
'   The ten largest valid netting sets of a regime: ID, counterparty, EAD
'   and share of the regime's EAD; blank rows below the last (view 5).
'
' UPDATED
'   2026-11-23
'==============================================================================
'
    Dim Values(1 To RANK_ROWS, 1 To 4)   As Variant    'The block
    Dim Order()                          As Long       'Ranked netting sets
    Dim Whole                            As Double     'The regime's EAD
    Dim Index                            As Long       'Rank
        Order = Ranked(Regime)
        Whole = ValidSum(Regime, "EAD")
        For Index = 1 To RANK_ROWS
            If Index > UBound(Order) Then Exit For
            Values(Index, 1) = mNsId(Order(Index))
            Values(Index, 2) = mNsCpty(Order(Index))
            Values(Index, 3) = mNsEAD(Order(Index)) / mExposureFactor
            Values(Index, 4) = ShareOf(mNsEAD(Order(Index)), Whole)
        Next Index
        TopBlock = Values

End Function


Private Function ConcentrationBlock() As Variant
'
'==============================================================================
'                              ConcentrationBlock
'------------------------------------------------------------------------------
' PURPOSE
'   Per regime: share of the ten largest valid netting sets in the EAD, and
'   the Herfindahl index over every valid netting set; blank without EAD
'   (view 6).
'
' UPDATED
'   2026-11-23
'==============================================================================
'
    Dim Values(1 To 2, 1 To 2)   As Variant    'The block
    Dim Order()                  As Long       'Ranked netting sets
    Dim Whole                    As Double     'The regime's EAD
    Dim Top                      As Double     'EAD of the ten largest
    Dim Index                    As Long       'Rank
    Dim Regime                   As Long       'Regime position
    Dim Herfindahl               As Double     'Sum of squared shares
        For Regime = 1 To 2
            Whole = ValidSum(Regime, "EAD")
            If Whole = 0# Then
                Values(1, Regime) = Empty
                Values(2, Regime) = Empty
            Else
                Order = Ranked(Regime)
                Top = 0#
                Herfindahl = 0#
                For Index = 1 To UBound(Order)
                    If Index <= RANK_ROWS Then Top = Top + mNsEAD(Order(Index))
                    Herfindahl = Herfindahl + (mNsEAD(Order(Index)) / Whole) ^ 2
                Next Index
                Values(1, Regime) = Top / Whole
                Values(2, Regime) = Herfindahl
            End If
        Next Regime
        ConcentrationBlock = Values

End Function


Private Function MaturityBlock() As Variant
'
'==============================================================================
'                                MaturityBlock
'------------------------------------------------------------------------------
' PURPOSE
'   Notional of the used legs by maturity band and regime (view 7).
'
' UPDATED
'   2026-11-23
'==============================================================================
'
    Dim Bands                    As Variant    'Band keys
    Dim Values(1 To 3, 1 To 2)   As Variant    'The block
    Dim Index                    As Long       'Row
    Dim Regime                   As Long       'Regime position
        Bands = Array("under_1y", "1y_5y", "over_5y")
        For Index = 0 To 2
            For Regime = 1 To 2
                Values(Index + 1, Regime) = LegSum(Regime, "NOTIONAL", CStr(Bands(Index))) / mNotionalFactor
            Next Regime
        Next Index
        MaturityBlock = Values

End Function


Private Function AssetClassBlock() As Variant
'
'==============================================================================
'                               AssetClassBlock
'------------------------------------------------------------------------------
' PURPOSE
'   Per asset class: used legs and their notional, CRR then BCBS (view 8).
'
' UPDATED
'   2026-11-23
'==============================================================================
'
    Dim Codes                    As Variant    'Asset classes
    Dim Values(1 To 6, 1 To 4)   As Variant    'The block
    Dim Index                    As Long       'Row
    Dim Regime                   As Long       'Regime position
        Codes = Split(AC_LIST, "|")
        For Index = 0 To 5
            For Regime = 1 To 2
                Values(Index + 1, 2 * Regime - 1) = LegCount(Regime, CStr(Codes(Index)))
                Values(Index + 1, 2 * Regime) = LegSum(Regime, "NOTIONAL", "", CStr(Codes(Index))) / mNotionalFactor
            Next Regime
        Next Index
        AssetClassBlock = Values

End Function


Private Function DataQualityBlock() As Variant
'
'==============================================================================
'                               DataQualityBlock
'------------------------------------------------------------------------------
' PURPOSE
'   Excluded legs, and netting sets whose EAD is withheld (view 9).
'
' UPDATED
'   2026-11-23
'==============================================================================
'
    Dim Values(1 To 2, 1 To 1)   As Variant    'The block
    Dim Index                    As Long       'Leg or netting set
        Values(1, 1) = 0
        Values(2, 1) = 0
        For Index = 1 To mLegCount
            If mLegStatus(Index) = "EXCLUDED" Then Values(1, 1) = Values(1, 1) + 1
        Next Index
        For Index = 1 To mNsCount
            If mNsStatus(Index) <> "VALID" Then Values(2, 1) = Values(2, 1) + 1
        Next Index
        DataQualityBlock = Values

End Function


Private Function CodesBlock() As Variant
'
'==============================================================================
'                                  CodesBlock
'------------------------------------------------------------------------------
' PURPOSE
'   The five codes with the most ERROR and WARNING rows on Checks: code,
'   ERROR when any of its rows is one, else WARNING, and the rows. More rows
'   first, then ERROR, then the code in binary order (view 9).
'
' UPDATED
'   2026-11-23
'==============================================================================
'
    Dim Values(1 To CODE_ROWS, 1 To 3)   As Variant      'The block
    Dim Lookup                           As Collection   'Position of each code
    Dim Codes()                          As String       'Distinct codes
    Dim Rows()                           As Long         'Rows per code
    Dim IsError()                        As Boolean      'An ERROR row per code
    Dim Count                            As Long         'Distinct codes
    Dim Index                            As Long         'Checks row, then rank
    Dim Position                         As Long         'A code's position
    Dim Best                             As Long         'The next code to rank
    Dim Taken()                          As Boolean      'Codes ranked already
    Dim Code                             As String       'The row's code
        Set Lookup = New Collection
        For Index = 1 To mCheckCount
            If mCheckSev(Index) = SEV_ERROR Or mCheckSev(Index) = SEV_WARN Then
                Code = mCheckCode(Index)
                If Len(Code) = 0 Then Code = NO_CODE
                Position = KeyIndex(Lookup, Code)
                If Position = 0 Then
                    Count = Count + 1
                    ReDim Preserve Codes(1 To Count)
                    ReDim Preserve Rows(1 To Count)
                    ReDim Preserve IsError(1 To Count)
                    Codes(Count) = Code
                    KeyAdd Lookup, Code, Count
                    Position = Count
                End If
                Rows(Position) = Rows(Position) + 1
                If mCheckSev(Index) = SEV_ERROR Then IsError(Position) = True
            End If
        Next Index
        If Count > 0 Then ReDim Taken(1 To Count)
        For Index = 1 To CODE_ROWS
            If Index > Count Then Exit For
            Best = 0
            For Position = 1 To Count
                If Not Taken(Position) Then
                    If Best = 0 Then
                        Best = Position
                    ElseIf Rows(Position) > Rows(Best) Then
                        Best = Position
                    ElseIf Rows(Position) = Rows(Best) Then
                        If IsError(Position) And Not IsError(Best) Then
                            Best = Position
                        ElseIf IsError(Position) = IsError(Best) Then
                            If StrComp(Codes(Position), Codes(Best), vbBinaryCompare) < 0 Then Best = Position
                        End If
                    End If
                End If
            Next Position
            Taken(Best) = True
            Values(Index, 1) = Codes(Best)
            Values(Index, 2) = IIf(IsError(Best), SEV_ERROR, SEV_WARN)
            Values(Index, 3) = Rows(Best)
        Next Index
        CodesBlock = Values

End Function


Private Function ReconciliationBlock() As Variant
'
'==============================================================================
'                             ReconciliationBlock
'------------------------------------------------------------------------------
' PURPOSE
'   For Reconciliation and TradeReconciliation: Y when the sheet holds a
'   comparison, else N; then its rows of MATCH, DIFFERENCE and MISSING IN
'   CALCULATOR, and of any other result (view 10).
'
' UPDATED
'   2026-11-23
'==============================================================================
'
    Dim Values(1 To 5, 1 To 2)   As Variant    'The block
    Dim Block                    As Variant    'A sheet's rows
    Dim Column                   As Long       'Sheet position
    Dim Index                    As Long       'Row
    Dim Result                   As String     'A row's result
        For Column = 1 To 2
            If Column = 1 Then Block = mRecon Else Block = mTradeRecon
            Values(1, Column) = IIf(RowCount(Block) > 0, "Y", "N")
            For Index = 2 To 5
                Values(Index, Column) = 0
            Next Index
            For Index = 1 To RowCount(Block)
                Result = UTxt(Block(Index, RC_RESULT))
                Select Case Result
                    Case RES_MATCH
                        Values(2, Column) = Values(2, Column) + 1
                    Case RES_DIFF
                        Values(3, Column) = Values(3, Column) + 1
                    Case RES_MISSING
                        Values(4, Column) = Values(4, Column) + 1
                    Case Else
                        Values(5, Column) = Values(5, Column) + 1
                End Select
            Next Index
        Next Column
        ReconciliationBlock = Values

End Function


Private Function PathBlock() As Variant
'
'==============================================================================
'                                  PathBlock
'------------------------------------------------------------------------------
' PURPOSE
'   Per regime, each with its share of the notional: the notional of the
'   used legs, their gross positive value, the value after close-out
'   netting (the sum of max(V, 0) over the valid netting sets), RC, PFE
'   and EAD (view 11).
'
' UPDATED
'   2026-11-23
'==============================================================================
'
    Dim Values(1 To 6, 1 To 4)   As Variant    'The block
    Dim Steps(1 To 6)            As Double     'The six amounts of a regime
    Dim Regime                   As Long       'Regime position
    Dim Index                    As Long       'Step
        For Regime = 1 To 2
            Steps(1) = LegSum(Regime, "NOTIONAL")
            Steps(2) = LegSum(Regime, "MTMPOS")
            Steps(3) = ValidSum(Regime, "VPOS")
            Steps(4) = ValidSum(Regime, "RC")
            Steps(5) = ValidSum(Regime, "PFE")
            Steps(6) = ValidSum(Regime, "EAD")
            For Index = 1 To 6
                Values(Index, 2 * Regime - 1) = Steps(Index) / IIf(Index = 1, mNotionalFactor, mExposureFactor)
                Values(Index, 2 * Regime) = ShareOf(Steps(Index), Steps(1))
            Next Index
        Next Regime
        PathBlock = Values

End Function


Private Function ReadingBlock() As Variant
'
'==============================================================================
'                                 ReadingBlock
'------------------------------------------------------------------------------
' PURPOSE
'   The reading of view 11 per regime; blank for a regime with no valid
'   netting set.
'
' UPDATED
'   2026-11-23
'==============================================================================
'
    Dim Values(1 To 2, 1 To 1)   As Variant    'The block
    Dim Regime                   As Long       'Regime position
    Dim NsIndex                  As Long       'Netting set
    Dim HasValid                 As Boolean    'The regime has a valid netting set
        For Regime = 1 To 2
            HasValid = False
            For NsIndex = 1 To mNsCount
                If IsValid(NsIndex, Regime) Then HasValid = True
            Next NsIndex
            If HasValid Then
                Values(Regime, 1) = ReadingText(LegSum(Regime, "MTMPOS"), ValidSum(Regime, "VPOS"), _
                                                ValidSum(Regime, "RC"))
            Else
                Values(Regime, 1) = ""
            End If
        Next Regime
        ReadingBlock = Values

End Function
