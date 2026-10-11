Attribute VB_Name = "CORE_Engine"
'==============================================================================
' MODULE: CORE_Engine
'------------------------------------------------------------------------------
' PURPOSE
'   Calculate SA-CCR exposure at default (EAD) for every netting set on the
'   input sheets, and write the intermediate and final figures to the output
'   sheets.
'
'     EAD = alpha * (RC + PFE)                                    [CRE52.1]
'     PFE = multiplier * aggregate add-on                         [CRE52.20]
'     aggregate add-on = sum of the asset-class add-ons           [CRE52.25]
'     trade effective notional = delta * d * MF
'
'   The EAD of a margined netting set is capped at the EAD of the same
'   netting set calculated as if it were unmargined.             [CRE52.2]
'
' REGIME LAYER
'   The regime is set on Params (CRR by default) and can be overridden per
'   netting set.
'   BCBS: the cap uses C = VM + NICA, with posted VM negative.    [CRE52.2]
'   CRR:  the cap follows Art. 274(3), with RC per Art. 275(1) and C = NICA
'         only, excluding VM posted or received (EBA Q&A 2023_6962).
'   CRR:  the other-risks class OT, supervisory factor 8%.       [Art. 280f]
'   CRR:  lambda for interest-rate and commodity options per Delegated
'         Regulation (EU) 2021/931 Art. 5.
'   CRR:  interest rate only with the bucket formula [Art. 280a(3)]; the
'         sum of absolute bucket values is a Basel alternative.
'   A row of the supervisory-factor table can be limited to one regime in
'   its Regimes column.
'
' FLOW
'   Calculate runs, in order:
'     ClearOutputSheets, ValidateSchema: clear old results, check the layout;
'     LoadParams, LoadSFTable, LoadFXTable, LoadNettingSets: read the inputs;
'     ProcessTrades: one row per trade, collected into buckets;
'     ComputeHedgingSets: buckets into hedging-set and asset-class add-ons;
'     ComputeNettingSets: RC, multiplier, PFE and EAD per netting set;
'     WriteBuckets, WriteHedgingSets, WriteChecks, WriteRunInfo: outputs.
'   Each trade, bucket and hedging set is stored once in an array, with a
'   Collection (see CORE_Util.KeyIndex) mapping its key to its array position.
'
' PUBLIC SURFACE
'   None outside this VBA project. Calculate, the run counters and the
'   run-fingerprint procedures (InputFingerprint, LastRunInputs,
'   ForgetRunInputs) are Public for M_Main; Option Private Module keeps them
'   off the supported external surface.
'
' DEPENDENCIES
'   CORE_Config for the layout, CORE_Util for conversions and sheet access, and
'   SACCR_Formulas for every regulatory formula.
'
' STATE OWNERSHIP
'   Owns the module state below. ResetState clears it at the start of each
'   run; the run counters stay readable until the next run. Reads the input
'   sheets and Params; writes the TradeCalc, Buckets, HedgingSets, Results
'   and Checks sheets. Changes no Excel application setting.
'
' ERROR POLICY
'   Invalid inputs are not VBA errors. They are logged as ERROR, WARNING or
'   INFO lines for the Checks sheet: an invalid trade is excluded and the
'   run continues; invalid parameters or netting sets stop the run.
'   LoadParams, WriteChecks and WriteRunInfo contain their own errors. A
'   Checks sheet that cannot be written withdraws the results and raises
'   ERR_CHECKS_WRITE (#36). Any other unexpected error during a run also
'   withdraws the results, is logged on Checks and in the run summary, and
'   is raised unchanged to M_Main (#36).
'
' KNOWN DEVIATION
'   This module lives in src/core but reads and writes worksheets, which the
'   repository structure reserves for src/workbook. See the known deviations
'   in docs/REPOSITORY_STRUCTURE.md.
'
' COMPATIBILITY
'   Excel VBA; no references beyond the defaults.
'
' UPDATED
'   2026-10-10
'
' AUTHOR
'   Daniele Penza
'==============================================================================

'------------------------------------------------------------------------------
' MODULE SETTINGS
'------------------------------------------------------------------------------
    'Require explicit declarations; keep the engine inside this project.
    Option Explicit
    Option Private Module

'------------------------------------------------------------------------------
' DATA STRUCTURES
'------------------------------------------------------------------------------
    'One netting set, as read from the NettingSets sheet. Amounts are in the
    'reporting currency.
    Private Type tNettingSet
        ID                As String     'Netting-set ID, upper case
        Counterparty      As String     'Counterparty name
        Margined          As Boolean    'Subject to a margin agreement
        ClearingRole      As String     'Client-clearing role: N, CM or CLIENT
        RemarginBD        As Double     'Remargining frequency, business days, at least 1
        LargeOrIlliquid   As Boolean    'Over 5,000 trades or illiquid collateral
        Disputes          As Boolean    'Margin disputes; doubles the MPOR
        MPOR              As Double     'Effective MPOR, business days
        MFMargined        As Double     'Margined maturity factor for the MPOR
        VM                As Double     'Net variation margin held (+) or posted (-)
        NICA              As Double     'Net independent collateral amount
        TH                As Double     'Threshold
        MTA               As Double     'Minimum transfer amount
        Alpha             As Double     'Alpha multiplier
        Regime            As String     'BCBS or CRR
        IsCRR             As Boolean    'Regime = CRR
        V                 As Double     'Sum of trade MtM
        Trades            As Long       'Number of valid trades
        Rejected          As Long       'Number of trades rejected by validation
        InputErrors       As Long       'Invalid netting-set fields; EAD withheld when > 0
    End Type

    'One row of the supervisory-factor table on Params.
    Private Type tSupervisory
        Key          As String    'Lookup key: IR, FX, OT, or class_subclass such as CR_AAA
        AssetClass   As String    'Asset-class code
        Category     As String    'Sub-class
        SF           As Double    'Supervisory factor
        Corr         As Double    'Correlation (credit, equity, commodity)
        Vol          As Double    'Supervisory option volatility
        Group        As String    'Commodity hedging set: ENERGY, METALS, ...
        Regimes      As String    'Blank (both regimes), BCBS or CRR
    End Type

    'One bucket: an IR maturity bucket, an FX currency pair, a credit or
    'equity entity, or a commodity type. Effective notionals are kept on
    'both bases: unmargined (U) and margined (M) maturity factor.
    Private Type tBucket
        NSIdx    As Long      'Netting set, position in mNS
        AC       As Long      'Asset class, an AC_ constant
        HSKey    As String    'Hedging-set key
        SubKey   As String    'Bucket key within the hedging set
        SF       As Double    'Supervisory factor after the basis or volatility multiplier
        Corr     As Double    'Correlation
        ENU      As Double    'Sum of delta * d * MF, unmargined MF
        ENM      As Double    'Sum of delta * d * MF, margined MF
        Trades   As Long      'Number of trades
        HSIdx    As Long      'Hedging set, position in mHS
    End Type

    'One hedging set, on both bases. IR uses the bucket totals D1 to D3;
    'FX and OT use EN; credit, equity and commodity use the systematic and
    'idiosyncratic sums of the single-factor model.
    Private Type tHedgingSet
        NSIdx    As Long      'Netting set, position in mNS
        AC       As Long      'Asset class
        Key      As String    'Hedging-set key
        D1U      As Double    'IR bucket under 1 year, unmargined
        D2U      As Double    'IR bucket 1 to 5 years, unmargined
        D3U      As Double    'IR bucket over 5 years, unmargined
        D1M      As Double    'IR bucket under 1 year, margined
        D2M      As Double    'IR bucket 1 to 5 years, margined
        D3M      As Double    'IR bucket over 5 years, margined
        SysU     As Double    'Sum of rho * add-on, unmargined
        IdioU    As Double    'Sum of (1 - rho^2) * add-on^2, unmargined
        SysM     As Double    'Sum of rho * add-on, margined
        IdioM    As Double    'Sum of (1 - rho^2) * add-on^2, margined
        ENU      As Double    'Effective notional, unmargined
        ENM      As Double    'Effective notional, margined
        AddOnU   As Double    'Hedging-set add-on, unmargined
        AddOnM   As Double    'Hedging-set add-on, margined
        Trades   As Long      'Number of trades
    End Type

'------------------------------------------------------------------------------
' MODULE STATE
'------------------------------------------------------------------------------
    'Each list is an array with a count of used elements and a Collection
    'mapping keys to positions. Arrays grow by doubling.
        Private mNS()         As tNettingSet     'Netting sets
        Private mNSCount      As Long            'Netting sets used
        Private mNSIndex      As Collection      'Netting-set ID to position

        Private mSF()         As tSupervisory    'Supervisory-factor table
        Private mSFCount      As Long            'Rows used
        Private mSFIndex      As Collection      'Key to position

        Private mFXRate()     As Double          'Units of reporting currency per unit of currency
        Private mFXCount      As Long            'Rates used
        Private mFXIndex      As Collection      'Currency code to position

        Private mBk()         As tBucket         'Buckets
        Private mBkCount      As Long            'Buckets used
        Private mBkIndex      As Collection      'Netting set # hedging set # bucket to position

        Private mHS()         As tHedgingSet     'Hedging sets
        Private mHSCount      As Long            'Hedging sets used
        Private mHSIndex      As Collection      'Netting set # hedging set to position

    'Sub-class of each credit or equity entity and commodity per netting set,
    'so that one reference cannot get two factors depending on row order.
        Private mRefClass()   As String          'Factor-table key of the reference, e.g. CR_AA
        Private mRefConflict() As Boolean        'A conflict was already reported
        Private mRefCount     As Long            'References used
        Private mRefIndex     As Collection      'Netting set | asset class | reference to position

        Private mAddOnU()     As Double          'Asset-class add-ons (netting set, asset class), unmargined
        Private mAddOnM()     As Double          'Asset-class add-ons (netting set, asset class), margined

    'Messages for the Checks sheet, one column per message, and the counts
    'reported to the user.
        Private mLog()        As Variant         '(field 1 to CK_NCOLS, message)
        Private mLogCount     As Long            'Messages logged
        Private mErrCount     As Long            'ERROR messages
        Private mWarnCount    As Long            'WARNING messages

        Private mTradesRead   As Long            'Trades with an ID
        Private mTradesUsed   As Long            'Trades included in the calculation
        Private mTotalEAD     As Double          'Sum of netting-set EADs, VALID sets only
        Private mIncomplete   As Long            'Netting sets whose EAD was withheld
        Private mParamError   As Boolean         'A parameter was present but not usable
        Private mRunInputs    As String          'Input fingerprint of this run; "" unless it completed

'------------------------------------------------------------------------------
' TEST SEAM
'------------------------------------------------------------------------------
    'Set only by TEST_MainState: "outputs" raises ERR_INJECTED_FAULT after
    'the first output sheets are written, to prove that a failed run
    'withdraws them (#36). "outputs+cleanup" also fails the TradeCalc and
    'Results clears with a different error. Empty in normal use.
        Public gEngineFault   As String

'------------------------------------------------------------------------------
' RUN PARAMETERS
'------------------------------------------------------------------------------
    'Read from Params by LoadParams; see CORE_Config for each code.
        Private pAsOf         As Double     'Reporting date, serial
        Private pRepCcy       As String     'Reporting currency
        Private pAlpha        As Double     'Default alpha
        Private pFloor        As Double     'Multiplier floor
        Private pDaysYear     As Double     'Calendar days per year
        Private pBDYear       As Double     'Business days per year
        Private pMinMatBD     As Double     'Floor on M, business days
        Private pSDFloorBD    As Double     'Floor on supervisory duration, business days
        Private pMPORBil      As Double     'MPOR floor, bilateral
        Private pMPORClr      As Double     'MPOR floor, cleared
        Private pMPORLarge    As Double     'MPOR floor, large or illiquid
        Private pBasisF       As Double     'Factor multiplier, basis transactions
        Private pVolF         As Double     'Factor multiplier, volatility transactions
        Private pRho12        As Double     'IR correlation, buckets 1 and 2
        Private pRho23        As Double     'IR correlation, buckets 2 and 3
        Private pRho13        As Double     'IR correlation, buckets 1 and 3
        Private pIRFull       As Boolean    'True: IR bucket formula; False: sum of absolutes
        Private pRegime       As String     'Default regime
        Private pLamThrIR     As Double     'CRR lambda threshold, IR options
        Private pLamThrCO     As Double     'CRR lambda threshold, commodity options


'
'------------------------------------------------------------------------------
'
'                         ENTRY POINT AND RUN COUNTERS
'
'------------------------------------------------------------------------------
'

Public Function Calculate( _
    ByVal writeOutputs As Boolean) _
    As Boolean
'
'==============================================================================
'                                  Calculate
'------------------------------------------------------------------------------
' PURPOSE
'   Run the engine once: load the inputs, calculate every trade, hedging set
'   and netting set, and write the outputs.
'
' INPUTS
'   writeOutputs: True writes every output sheet; False validates only and
'   writes the Checks sheet alone.
'
' RETURNS
'   True when the run completed. Individual trades may still have been
'   excluded; the Checks sheet lists them. False when the sheet layout is
'   not as expected, or a parameter, the factor or FX table, or the netting
'   sets could not be loaded.
'
' ERROR POLICY
'   Raises ERR_CHECKS_WRITE when the Checks sheet cannot be written, after
'   clearing the other output sheets: results whose errors and warnings
'   cannot be shown must not look valid, and the Checks sheet would still
'   hold the previous run's messages (#36). Any other error during the run
'   clears the output sheets too, so that a run stopped halfway cannot
'   leave tables from two runs, is logged on Checks and in the run summary,
'   and is raised with its original number, source and description, with
'   any output-cleanup or Checks-write diagnostics appended.
'
' STATE OWNERSHIP
'   Resets all module state, then fills it for this run.
'
' UPDATED
'   2026-10-06
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim t0               As Double     'Timer value at the start, for the run duration
    Dim checksFailure    As String     'Why the Checks sheet could not be written; "" if written
    Dim errNumber        As Long       'Unexpected error: number
    Dim errSource        As String     'Unexpected error: source
    Dim withdrawalDetails As String    'Failures by sheet, distinct from the primary error
    Dim withdrawalStatus As String     'Whether every output could be withdrawn
    Dim errDescription   As String     'Unexpected error: description

'------------------------------------------------------------------------------
' LOAD INPUTS
'------------------------------------------------------------------------------
    'Old outputs are cleared first, for validation too, so that a failed or
    'validation-only run can never leave earlier results looking current
    '(#36). Any loader that fails has logged why; skip to the Checks output.
        t0 = Timer
        ResetState
        On Error GoTo Failed
        ClearOutputSheets

        If Not ValidateSchema() Then GoTo Finish
        If Not LoadParams() Then GoTo Finish
        If Not LoadSFTable() Then GoTo Finish
        If Not LoadFXTable() Then GoTo Finish
        If Not LoadNettingSets() Then GoTo Finish

'------------------------------------------------------------------------------
' CALCULATE
'------------------------------------------------------------------------------
        ProcessTrades writeOutputs
        ComputeHedgingSets
        ComputeNettingSets writeOutputs
        If writeOutputs Then
            WriteBuckets
            If gEngineFault = "outputs" Or gEngineFault = "outputs+cleanup" Then
                Err.Raise ERR_INJECTED_FAULT, "CORE_Engine.Calculate", "Injected output failure."
            End If
            WriteHedgingSets
        End If
        Calculate = True

'------------------------------------------------------------------------------
' WRITE CHECKS AND RUN SUMMARY
'------------------------------------------------------------------------------
    'Reached on success and after a failed load. If the Checks sheet cannot
    'be written, the results are withdrawn and the failure raised.
Finish:
        If Not WriteChecks(checksFailure) Then
            Calculate = False
            withdrawalStatus = "results withdrawn"
            If Not WithdrawOutputs(withdrawalDetails) Then
                withdrawalStatus = "output cleanup incomplete; old results may remain - unprotect or repair outputs and rerun. " & _
                                   withdrawalDetails
            End If
            ForgetRunInputs
            WriteRunInfo Timer - t0, False, writeOutputs, "the Checks sheet could not be written (" & _
                         checksFailure & "); " & withdrawalStatus & ", and Checks may show an earlier run"
            On Error GoTo 0
            Err.Raise ERR_CHECKS_WRITE, "CORE_Engine.Calculate", "The Checks sheet could not be written (" & _
                      checksFailure & "); " & withdrawalStatus & ". Unprotect or repair the Checks " & _
                      "sheet and run again."
        End If

    'Only a completed run that wrote its outputs leaves results; remember
    'its inputs so that later edits show the results as out of date.
        If Calculate And writeOutputs Then
            mRunInputs = InputFingerprint()
            RememberRunInputs mRunInputs
        Else
            ForgetRunInputs
        End If
        WriteRunInfo Timer - t0, Calculate, writeOutputs, ""
        Exit Function

'------------------------------------------------------------------------------
' HANDLE UNEXPECTED ERROR
'------------------------------------------------------------------------------
    'Attempt withdrawal of everything written so far, and log the original
    'error on Checks and in the run summary, both best effort. Preserve its
    'number/source/text and append any secondary failure for M_Main to report.
Failed:
        errNumber = Err.Number
        errSource = Err.Source
        errDescription = Err.Description
        Calculate = False
        mRunInputs = ""
        'NONE means no valid run; protected output cells may still remain.
        ForgetRunInputs
        withdrawalStatus = "results withdrawn"
        If Not WithdrawOutputs(withdrawalDetails) Then
            withdrawalStatus = "output cleanup incomplete; old results may remain - unprotect or repair outputs and rerun. " & _
                               withdrawalDetails
        End If
        LogMsg SEV_ERROR, "", "", "Run stopped by error " & errNumber & ": " & errDescription & _
               " - " & withdrawalStatus & "."
        If Not WriteChecks(checksFailure) Then
            withdrawalStatus = withdrawalStatus & "; Checks could not be updated (" & checksFailure & _
                               "); Checks may show an earlier run."
        End If
        WriteRunInfo Timer - t0, False, writeOutputs, "error " & errNumber & ": " & errDescription & _
                     "; " & withdrawalStatus
        If Len(withdrawalDetails) > 0 Or Len(checksFailure) > 0 Then
            errDescription = errDescription & " " & withdrawalStatus & "."
        End If
        Err.Raise errNumber, errSource, errDescription

End Function


Public Property Get ErrorCount() As Long
'
'==============================================================================
'                                  ErrorCount
'------------------------------------------------------------------------------
' PURPOSE
'   Number of ERROR messages logged by the last run.
'
' UPDATED
'   2026-10-06
'==============================================================================
'
        ErrorCount = mErrCount

End Property


Public Property Get WarningCount() As Long
'
'==============================================================================
'                                 WarningCount
'------------------------------------------------------------------------------
' PURPOSE
'   Number of WARNING messages logged by the last run.
'
' UPDATED
'   2026-10-06
'==============================================================================
'
        WarningCount = mWarnCount

End Property


Public Property Get TotalEAD() As Double
'
'==============================================================================
'                                   TotalEAD
'------------------------------------------------------------------------------
' PURPOSE
'   Sum of the EADs of every reported netting set in the last run, in the
'   reporting currency.
'
' UPDATED
'   2026-10-06
'==============================================================================
'
        TotalEAD = mTotalEAD

End Property


Public Property Get TradesUsed() As Long
'
'==============================================================================
'                                  TradesUsed
'------------------------------------------------------------------------------
' PURPOSE
'   Number of trades included in the last run's calculation.
'
' UPDATED
'   2026-10-06
'==============================================================================
'
        TradesUsed = mTradesUsed

End Property


Public Property Get IncompleteCount() As Long
'
'==============================================================================
'                               IncompleteCount
'------------------------------------------------------------------------------
' PURPOSE
'   Number of netting sets in the last run whose EAD was withheld because
'   a trade was rejected or a netting-set field was invalid.
'
' UPDATED
'   2026-10-06
'==============================================================================
'
        IncompleteCount = mIncomplete

End Property


Public Property Get TradesRead() As Long
'
'==============================================================================
'                                  TradesRead
'------------------------------------------------------------------------------
' PURPOSE
'   Number of trade rows with an ID read by the last run.
'
' UPDATED
'   2026-10-06
'==============================================================================
'
        TradesRead = mTradesRead

End Property


Public Property Get RunInputs() As String
'
'==============================================================================
'                                  RunInputs
'------------------------------------------------------------------------------
' PURPOSE
'   Input fingerprint of the last run, "" unless it completed and wrote its
'   outputs. Identifies the inputs behind a result line (#36).
'
' UPDATED
'   2026-10-07
'==============================================================================
'
        RunInputs = mRunInputs

End Property


'
'------------------------------------------------------------------------------
'
'                           INITIALISATION AND LOG
'
'------------------------------------------------------------------------------
'

Private Sub ResetState()
'
'==============================================================================
'                                  ResetState
'------------------------------------------------------------------------------
' PURPOSE
'   Clear every list, counter and index before a run, and allocate the
'   arrays at a starting size.
'
' UPDATED
'   2026-10-06
'==============================================================================
'

'------------------------------------------------------------------------------
' RESET COUNTERS
'------------------------------------------------------------------------------
        mNSCount = 0
        mSFCount = 0
        mFXCount = 0
        mBkCount = 0
        mHSCount = 0
        mRefCount = 0
        mLogCount = 0
        mErrCount = 0
        mWarnCount = 0
        mTradesRead = 0
        mTradesUsed = 0
        mTotalEAD = 0#
        mIncomplete = 0
        mParamError = False
        mRunInputs = ""

'------------------------------------------------------------------------------
' RESET INDEXES AND ARRAYS
'------------------------------------------------------------------------------
    'The starting sizes are a guess; each array doubles when it fills up.
        Set mNSIndex = New Collection
        Set mSFIndex = New Collection
        Set mFXIndex = New Collection
        Set mBkIndex = New Collection
        Set mHSIndex = New Collection
        Set mRefIndex = New Collection
        ReDim mNS(1 To 16)
        ReDim mSF(1 To 32)
        ReDim mFXRate(1 To 32)
        ReDim mBk(1 To 64)
        ReDim mHS(1 To 32)
        ReDim mRefClass(1 To 32)
        ReDim mRefConflict(1 To 32)
        ReDim mLog(1 To CK_NCOLS, 1 To 64)

End Sub


Private Sub LogMsg( _
    ByVal severity As String, _
    ByVal sheetName As String, _
    ByVal ref As String, _
    ByVal msg As String, _
    Optional ByVal rowNum As Long = 0)
'
'==============================================================================
'                                    LogMsg
'------------------------------------------------------------------------------
' PURPOSE
'   Add one message for the Checks sheet and count errors and warnings.
'
' INPUTS
'   severity: SEV_ERROR, SEV_WARN or SEV_INFO.
'   sheetName: sheet the message refers to.
'   ref: the ID or parameter code concerned; may be empty.
'   msg: the message text.
'   rowNum: sheet row concerned; 0 leaves the Row column blank.
'
' UPDATED
'   2026-10-06
'==============================================================================
'

'------------------------------------------------------------------------------
' STORE MESSAGE
'------------------------------------------------------------------------------
    'mLog is (field, message) so that ReDim Preserve can grow its last
    'dimension.
        mLogCount = mLogCount + 1
        If mLogCount > UBound(mLog, 2) Then
            ReDim Preserve mLog(1 To CK_NCOLS, 1 To mLogCount * 2)
        End If
        mLog(1, mLogCount) = severity
        mLog(2, mLogCount) = sheetName
        If rowNum > 0 Then
            mLog(3, mLogCount) = rowNum
        Else
            mLog(3, mLogCount) = Empty
        End If
        mLog(4, mLogCount) = ref
        mLog(5, mLogCount) = msg

'------------------------------------------------------------------------------
' COUNT
'------------------------------------------------------------------------------
        If severity = SEV_ERROR Then
            mErrCount = mErrCount + 1
        End If
        If severity = SEV_WARN Then
            mWarnCount = mWarnCount + 1
        End If

End Sub


'
'------------------------------------------------------------------------------
'
'                                SCHEMA CHECK
'
'------------------------------------------------------------------------------
'

Private Function ValidateSchema() As Boolean
'
'==============================================================================
'                                ValidateSchema
'------------------------------------------------------------------------------
' PURPOSE
'   Check the workbook layout before any input is read (#35). The engine
'   reads every input by column number, so a column inserted, deleted or
'   moved would silently shift values into the wrong fields.
'
' RETURNS
'   True when the input sheets exist, every column the engine reads has its
'   expected header, no parameter is listed twice with different values and
'   no parameter name is broken. False otherwise, with an error logged for
'   each problem found.
'
' UPDATED
'   2026-10-07
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ws         As Worksheet    'Params sheet
    Dim ok         As Boolean      'No problem found so far
    Dim r0         As Long         'Header row of a Params table
    Dim lastR      As Long         'Last used row of Params
    Dim grid       As Variant      'Params columns A to C, rows 1 to lastR
    Dim codes      As Variant      'Every parameter code
    Dim i          As Long         'Index into codes
    Dim r          As Long         'Params row being compared
    Dim firstRow   As Long         'First Params row holding the code; 0 if none

'------------------------------------------------------------------------------
' INPUT SHEETS
'------------------------------------------------------------------------------
    'A missing output sheet has already raised in ClearOutputSheets.
        ok = SheetPresent(SH_PARAMS)
        ok = SheetPresent(SH_NS) And ok
        ok = SheetPresent(SH_TRADES) And ok
        If Not ok Then
            Exit Function
        End If

'------------------------------------------------------------------------------
' COLUMN HEADERS
'------------------------------------------------------------------------------
    'A missing factor or FX table header is reported by its loader.
        Set ws = GetSheet(SH_PARAMS)
        ok = HeadersMatch(GetSheet(SH_NS), HEADER_ROW, NS_HEADERS) And ok
        ok = HeadersMatch(GetSheet(SH_TRADES), HEADER_ROW, TR_HEADERS) And ok
        ok = HeadersMatch(ws, HEADER_ROW, PRM_HEADERS) And ok
        r0 = FindHeaderRow(ws, 1, HDR_SF)
        If r0 > 0 Then
            ok = HeadersMatch(ws, r0, SF_HEADERS) And ok
        End If
        r0 = FindHeaderRow(ws, 1, HDR_FX)
        If r0 > 0 Then
            ok = HeadersMatch(ws, r0, FX_HEADERS) And ok
        End If

'------------------------------------------------------------------------------
' PARAMETERS
'------------------------------------------------------------------------------
    'A parameter listed twice with the same value is harmless and warned
    'about; with different values it is ambiguous. A workbook name equal to
    'the code takes precedence over the Params row (CORE_Util.GetParam); a
    'broken one would silently fall back to the row.
        codes = ParameterCodes()
        lastR = UsedLastRow(ws)
        If lastR >= 1 Then
            grid = ws.Range(ws.Cells(1, 1), ws.Cells(lastR, PRM_VALUE_COL)).Value
        End If
        For i = LBound(codes) To UBound(codes)
            If BrokenName(codes(i)) Then
                ok = False
                LogMsg SEV_ERROR, SH_PARAMS, codes(i), "Workbook name '" & codes(i) & _
                       "' no longer refers to a cell (#REF!)."
            End If
            firstRow = 0
            For r = 1 To lastR
                If StrComp(SafeStr(grid(r, PRM_CODE_COL)), codes(i), vbTextCompare) = 0 Then
                    If firstRow = 0 Then
                        firstRow = r
                    ElseIf SameValue(grid(firstRow, PRM_VALUE_COL), grid(r, PRM_VALUE_COL)) Then
                        LogMsg SEV_WARN, SH_PARAMS, codes(i), "Parameter listed twice with the same value (rows " & _
                               firstRow & " and " & r & ").", r
                    Else
                        ok = False
                        LogMsg SEV_ERROR, SH_PARAMS, codes(i), "Parameter listed twice with different values (rows " & _
                               firstRow & " and " & r & ").", r
                    End If
                End If
            Next r
        Next i
        ValidateSchema = ok

End Function


Private Function ParameterCodes() As Variant
'
'==============================================================================
' PURPOSE
'   The parameter vocabulary shared by schema and table validation.
' RETURNS
'   An array of the supported parameter codes.
'==============================================================================
        ParameterCodes = Array(PRM_ASOF, PRM_REPCCY, PRM_ALPHA, PRM_FLOOR, PRM_DAYSYEAR, PRM_BDYEAR, _
                      PRM_MINMAT, PRM_SDFLOOR, PRM_MPOR_BIL, PRM_MPOR_CLR, PRM_MPOR_LARGE, _
                      PRM_BASIS, PRM_VOLF, PRM_RHO12, PRM_RHO23, PRM_RHO13, PRM_IRFULL, _
                      PRM_REGIME, PRM_LAMIR, PRM_LAMCO)
End Function


Private Function IsParameterCode(ByVal key As String) As Boolean
'
'==============================================================================
' PURPOSE
'   Distinguish parameter rows from FX rows, including below table gaps.
'==============================================================================
    Dim code As Variant    'A supported parameter code
        For Each code In ParameterCodes()
            If StrComp(key, CStr(code), vbTextCompare) = 0 Then
                IsParameterCode = True
                Exit Function
            End If
        Next code
End Function


Private Function SheetPresent( _
    ByVal sheetName As String) _
    As Boolean
'
'==============================================================================
'                                 SheetPresent
'------------------------------------------------------------------------------
' PURPOSE
'   Tell whether this workbook has a worksheet with the given tab name, and
'   log an error when it has not.
'
' UPDATED
'   2026-10-07
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ws   As Worksheet    'Worksheet being compared

'------------------------------------------------------------------------------
' SEARCH
'------------------------------------------------------------------------------
        For Each ws In ThisWorkbook.Worksheets
            If StrComp(ws.Name, sheetName, vbTextCompare) = 0 Then
                SheetPresent = True
                Exit Function
            End If
        Next ws
        LogMsg SEV_ERROR, sheetName, "", "Sheet '" & sheetName & "' not found."

End Function


Private Function HeadersMatch( _
    ByVal ws As Worksheet, _
    ByVal rowNum As Long, _
    ByVal expectedHeaders As String) _
    As Boolean
'
'==============================================================================
'                                 HeadersMatch
'------------------------------------------------------------------------------
' PURPOSE
'   Compare a header row with the expected headers, ignoring case and
'   surrounding spaces, and log an error for each header that differs.
'
' INPUTS
'   ws, rowNum: the sheet and its header row.
'   expectedHeaders: one of the _HEADERS constants in CORE_Config, starting in
'   column A; an empty entry is not checked.
'
' RETURNS
'   True when every checked header matches.
'
' UPDATED
'   2026-10-07
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim headers   As Variant    'Expected headers, from index 0 for column A
    Dim c         As Long       'Index into headers
    Dim found     As Variant    'Header cell value

'------------------------------------------------------------------------------
' COMPARE
'------------------------------------------------------------------------------
        HeadersMatch = True
        headers = Split(expectedHeaders, "|")
        For c = 0 To UBound(headers)
            If Len(headers(c)) > 0 Then
                found = ws.Cells(rowNum, c + 1).Value
                If StrComp(SafeStr(found), headers(c), vbTextCompare) <> 0 Then
                    HeadersMatch = False
                    LogMsg SEV_ERROR, ws.Name, ws.Cells(rowNum, c + 1).Address(False, False), _
                           "Header must be '" & headers(c) & "' (found " & DescribeValue(found) & _
                           ") - a column was inserted, deleted or moved, or the header renamed.", rowNum
                End If
            End If
        Next c

End Function


Private Function SameValue( _
    ByVal a As Variant, _
    ByVal b As Variant) _
    As Boolean
'
'==============================================================================
'                                  SameValue
'------------------------------------------------------------------------------
' PURPOSE
'   Compare two cell values: numbers exactly, text ignoring case and
'   surrounding spaces. An error value is never the same as anything.
'
' UPDATED
'   2026-10-07
'==============================================================================
'
        If IsError(a) Or IsError(b) Then
            SameValue = False
        ElseIf IsNum(a) And IsNum(b) Then
            SameValue = (CDbl(a) = CDbl(b))
        Else
            SameValue = (StrComp(SafeStr(a), SafeStr(b), vbTextCompare) = 0)
        End If

End Function


'
'------------------------------------------------------------------------------
'
'                                INPUT LOADING
'
'------------------------------------------------------------------------------
'

Private Function NumParam( _
    ByVal code As String, _
    ByVal dflt As Double) _
    As Double
'
'==============================================================================
'                                   NumParam
'------------------------------------------------------------------------------
' PURPOSE
'   Read a numeric parameter, falling back to its regulatory default.
'
' INPUTS
'   code: a PRM_ code from CORE_Config.
'   dflt: value used, with a warning, when the parameter is blank.
'
' RETURNS
'   The parameter value, or dflt.
'
' ERROR POLICY
'   A value that is present but not a number is an error and sets
'   mParamError, so that LoadParams stops the run (#35).
'
' UPDATED
'   2026-10-06
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim v   As Variant    'Raw parameter value

'------------------------------------------------------------------------------
' READ
'------------------------------------------------------------------------------
        v = GetParam(code)
        If IsNum(v) Then
            NumParam = CDbl(v)
        ElseIf IsBlankCell(v) And Not IsError(v) Then
            NumParam = dflt
            LogMsg SEV_WARN, SH_PARAMS, code, "Parameter missing - default " & CStr(dflt) & " used."
        Else
            NumParam = dflt
            mParamError = True
            LogMsg SEV_ERROR, SH_PARAMS, code, "Parameter must be a number (found " & DescribeValue(v) & ")."
        End If

End Function


Private Function LoadParams() As Boolean
'
'==============================================================================
'                                  LoadParams
'------------------------------------------------------------------------------
' PURPOSE
'   Read every run parameter from Params into the p* variables. Numeric
'   parameters that are missing take their regulatory default with a
'   warning.
'
' RETURNS
'   True when the parameters are usable. False, with an error logged, when
'   the reporting date or currency is missing, the regime is not BCBS or
'   CRR, or a day count is not positive.
'
' ERROR POLICY
'   An unexpected error while reading is logged as an error and returns
'   False.
'
' UPDATED
'   2026-10-06
'==============================================================================
'

'------------------------------------------------------------------------------
' REPORTING DATE AND CURRENCY
'------------------------------------------------------------------------------
    'Both are required: there is no sensible default.
        On Error GoTo Fail
        pAsOf = ToSerial(GetParam(PRM_ASOF))
        If pAsOf <= 0# Then
            LogMsg SEV_ERROR, SH_PARAMS, PRM_ASOF, "Reporting date missing or invalid."
            Exit Function
        End If
        pRepCcy = UTxt(GetParam(PRM_REPCCY))
        If Len(pRepCcy) = 0 Then
            LogMsg SEV_ERROR, SH_PARAMS, PRM_REPCCY, "Reporting currency missing."
            Exit Function
        End If

'------------------------------------------------------------------------------
' NUMERIC PARAMETERS
'------------------------------------------------------------------------------
    'Defaults are the regulatory values: CRE52 and CRR Art. 274 to 280f.
        pAlpha = NumParam(PRM_ALPHA, 1.4)
        If pAlpha <= 0# Then
            LogMsg SEV_ERROR, SH_PARAMS, PRM_ALPHA, "Alpha must be greater than zero."
            Exit Function
        End If
        pFloor = NumParam(PRM_FLOOR, 0.05)
        pDaysYear = NumParam(PRM_DAYSYEAR, 365#)
        pBDYear = NumParam(PRM_BDYEAR, 250#)
        pMinMatBD = NumParam(PRM_MINMAT, 10#)
        pSDFloorBD = NumParam(PRM_SDFLOOR, 10#)
        pMPORBil = NumParam(PRM_MPOR_BIL, 10#)
        pMPORClr = NumParam(PRM_MPOR_CLR, 5#)
        pMPORLarge = NumParam(PRM_MPOR_LARGE, 20#)
        pBasisF = NumParam(PRM_BASIS, 0.5)
        pVolF = NumParam(PRM_VOLF, 5#)
        pRho12 = NumParam(PRM_RHO12, 0.7)
        pRho23 = NumParam(PRM_RHO23, 0.7)
        pRho13 = NumParam(PRM_RHO13, 0.3)
        If Not TryBool(GetParam(PRM_IRFULL), True, pIRFull) Then
            LogMsg SEV_ERROR, SH_PARAMS, PRM_IRFULL, "Parameter must be TRUE or FALSE (found " & _
                   DescribeValue(GetParam(PRM_IRFULL)) & ")."
            Exit Function
        End If

'------------------------------------------------------------------------------
' REGIME
'------------------------------------------------------------------------------
    'A blank regime means CRR; anything other than BCBS or CRR stops the run.
        pRegime = UTxt(GetParam(PRM_REGIME))
        If Len(pRegime) = 0 Then
            pRegime = RG_CRR
            LogMsg SEV_WARN, SH_PARAMS, PRM_REGIME, "Regime missing - CRR assumed."
        ElseIf pRegime <> RG_BCBS And pRegime <> RG_CRR Then
            LogMsg SEV_ERROR, SH_PARAMS, PRM_REGIME, "Regime must be BCBS or CRR."
            Exit Function
        End If
        pLamThrIR = NumParam(PRM_LAMIR, 0.001)
        pLamThrCO = NumParam(PRM_LAMCO, 0.1)
        If mParamError Then
            Exit Function
        End If

'------------------------------------------------------------------------------
' DAY COUNTS
'------------------------------------------------------------------------------
    'Both are divisors in every year fraction.
        If pDaysYear <= 0# Or pBDYear <= 0# Then
            LogMsg SEV_ERROR, SH_PARAMS, PRM_DAYSYEAR, "Day-count parameters must be positive."
            Exit Function
        End If
        LoadParams = True
        Exit Function

'------------------------------------------------------------------------------
' HANDLE ERROR
'------------------------------------------------------------------------------
Fail:
        LogMsg SEV_ERROR, SH_PARAMS, "", "Cannot read parameters: " & Err.Description

End Function


Private Function LoadSFTable() As Boolean
'
'==============================================================================
'                                 LoadSFTable
'------------------------------------------------------------------------------
' PURPOSE
'   Read the supervisory-factor table on Params: the rows below the SF_Key
'   header, down to the first blank key.
'
' RETURNS
'   True when at least one row was read. False, with an error logged, when
'   the header is missing, the table is empty, a value is not a number, a
'   key is listed twice with different values, or a row with a factor lies
'   below the first blank key and would not be read (#35).
'
' STATE OWNERSHIP
'   Fills mSF, mSFCount and mSFIndex. A key listed twice with the same
'   values is warned about and the first row kept. "BOTH" in the Regimes
'   column is stored as blank.
'
' UPDATED
'   2026-10-07
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ws       As Worksheet      'Params sheet
    Dim r        As Long           'Row being read
    Dim r0       As Long           'Header row of the table
    Dim k        As String         'Key of the row, upper case
    Dim bad      As Boolean        'Some row is unusable or conflicts
    Dim rowBad   As Boolean        'This row has a value that is not a number
    Dim entry    As tSupervisory   'This row
    Dim idx      As Long           'Position of an earlier row with the key; 0 if none
    Dim lastR    As Long           'Last row searched for rows after the end
    Dim fxRow    As Long           'Header row of the FX table; 0 if absent

'------------------------------------------------------------------------------
' FIND THE TABLE
'------------------------------------------------------------------------------
        Set ws = GetSheet(SH_PARAMS)
        r0 = FindHeaderRow(ws, 1, HDR_SF)
        If r0 = 0 Then
            LogMsg SEV_ERROR, SH_PARAMS, HDR_SF, "Supervisory factor table not found (header '" & HDR_SF & "' in column A)."
            Exit Function
        End If

'------------------------------------------------------------------------------
' READ ROWS
'------------------------------------------------------------------------------
    'Columns: 1 key, 2 asset class, 3 category, 4 factor, 5 correlation,
    '6 option volatility, 7 commodity hedging set, 8 regimes.
        r = r0 + 1
        Do While Not IsBlankCell(ws.Cells(r, 1).Value)
            k = UTxt(ws.Cells(r, 1).Value)
            entry.Key = k
            entry.AssetClass = UTxt(ws.Cells(r, 2).Value)
            entry.Category = UTxt(ws.Cells(r, 3).Value)
            rowBad = Not TryDbl(ws.Cells(r, 4).Value, 0#, entry.SF)
            rowBad = Not TryDbl(ws.Cells(r, 5).Value, 0#, entry.Corr) Or rowBad
            rowBad = Not TryDbl(ws.Cells(r, 6).Value, 0#, entry.Vol) Or rowBad
            entry.Group = UTxt(ws.Cells(r, 7).Value)
            entry.Regimes = UTxt(ws.Cells(r, 8).Value)
            If entry.Regimes = "BOTH" Then
                entry.Regimes = ""
            End If
            If rowBad Then
                bad = True
                LogMsg SEV_ERROR, SH_PARAMS, k, "Supervisory factor, correlation and volatility must be numbers.", r
            End If

    'A key seen before: the same values are harmless, different values are
    'ambiguous and stop the run. A row with a value that is not a number has
    'already been reported.
            idx = KeyIndex(mSFIndex, k)
            If idx = 0 Then
                mSFCount = mSFCount + 1
                If mSFCount > UBound(mSF) Then
                    ReDim Preserve mSF(1 To mSFCount * 2)
                End If
                mSF(mSFCount) = entry
                KeyAdd mSFIndex, k, mSFCount
            ElseIf Not rowBad Then
                If SameFactorRow(mSF(idx), entry) Then
                    LogMsg SEV_WARN, SH_PARAMS, k, "Supervisory factor key listed twice with the same values - " & _
                           "first occurrence used.", r
                Else
                    bad = True
                    LogMsg SEV_ERROR, SH_PARAMS, k, "Supervisory factor key listed twice with different values.", r
                End If
            End If
            r = r + 1
        Loop

'------------------------------------------------------------------------------
' ROWS AFTER THE END
'------------------------------------------------------------------------------
    'The table ends at the first blank key. A row further down with a
    'factor would be dropped silently, so it is an error. The search stops
    'at the FX table when that follows.
        lastR = UsedLastRow(ws)
        fxRow = FindHeaderRow(ws, 1, HDR_FX)
        If fxRow > r Then
            lastR = fxRow - 1
        End If
        If ValuesAfterEnd(ws, r, lastR, 4, "Supervisory factor") Then
            bad = True
        End If

'------------------------------------------------------------------------------
' CHECK
'------------------------------------------------------------------------------
        If mSFCount = 0 Then
            LogMsg SEV_ERROR, SH_PARAMS, HDR_SF, "Supervisory factor table is empty."
            Exit Function
        End If
        If bad Then
            Exit Function
        End If
        LoadSFTable = True

End Function


Private Function SameFactorRow( _
    ByRef first As tSupervisory, _
    ByRef other As tSupervisory) _
    As Boolean
'
'==============================================================================
'                                SameFactorRow
'------------------------------------------------------------------------------
' PURPOSE
'   Tell whether two rows of the supervisory-factor table hold the same
'   values in every column the engine reads.
'
' UPDATED
'   2026-10-07
'==============================================================================
'
        SameFactorRow = (first.AssetClass = other.AssetClass) And (first.Category = other.Category) And _
                        (first.SF = other.SF) And (first.Corr = other.Corr) And (first.Vol = other.Vol) And _
                        (first.Group = other.Group) And (first.Regimes = other.Regimes)

End Function


Private Function LoadFXTable() As Boolean
'
'==============================================================================
'                                 LoadFXTable
'------------------------------------------------------------------------------
' PURPOSE
'   Read the FX table on Params: the rows below the FX_Ccy header, down to
'   the first blank currency. Column C holds the units of reporting currency
'   per one unit of the currency.
'
' RETURNS
'   True when the table was read. False, with an error logged, when the
'   header is missing, a currency is listed twice with different rates, or
'   a row with a rate lies below the first blank currency and would not be
'   read (#35).
'
' STATE OWNERSHIP
'   Fills mFXRate, mFXCount and mFXIndex. A rate that is missing or not
'   positive is warned about and the currency left out; a currency listed
'   twice with the same rate is warned about and the first row kept. The
'   reporting currency is added at rate 1 if absent, and warned about if
'   its rate is not 1.
'
' UPDATED
'   2026-10-07
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ws     As Worksheet    'Params sheet
    Dim r      As Long         'Row being read
    Dim r0     As Long         'Header row of the table
    Dim k      As String       'Currency code, upper case
    Dim rate   As Double       'Rate of the row; -1 when missing
    Dim idx    As Long         'Position of an earlier row with the currency; 0 if none
    Dim bad    As Boolean      'A currency conflicts or a rate would not be read
    Dim lastR  As Long         'Last row searched for rows after the end
    Dim sfRow  As Long         'Header row of the factor table; 0 if absent

'------------------------------------------------------------------------------
' FIND THE TABLE
'------------------------------------------------------------------------------
        Set ws = GetSheet(SH_PARAMS)
        r0 = FindHeaderRow(ws, 1, HDR_FX)
        If r0 = 0 Then
            LogMsg SEV_ERROR, SH_PARAMS, HDR_FX, "FX table not found (header '" & HDR_FX & "' in column A)."
            Exit Function
        End If

'------------------------------------------------------------------------------
' READ ROWS
'------------------------------------------------------------------------------
    'A currency seen before: the same rate is harmless, a different rate is
    'ambiguous and stops the run.
        r = r0 + 1
        Do While Not IsBlankCell(ws.Cells(r, 1).Value)
            k = UTxt(ws.Cells(r, 1).Value)
            rate = ToDbl(ws.Cells(r, 3).Value, -1#)
            idx = KeyIndex(mFXIndex, k)
            If rate <= 0# Then
                LogMsg SEV_WARN, SH_PARAMS, k, "FX rate missing or not positive - currency ignored.", r
            ElseIf idx = 0 Then
                mFXCount = mFXCount + 1
                If mFXCount > UBound(mFXRate) Then
                    ReDim Preserve mFXRate(1 To mFXCount * 2)
                End If
                mFXRate(mFXCount) = rate
                KeyAdd mFXIndex, k, mFXCount
            ElseIf mFXRate(idx) = rate Then
                LogMsg SEV_WARN, SH_PARAMS, k, "Currency listed twice with the same rate - first row used.", r
            Else
                bad = True
                LogMsg SEV_ERROR, SH_PARAMS, k, "Currency listed twice with different rates.", r
            End If
            r = r + 1
        Loop

'------------------------------------------------------------------------------
' ROWS AFTER THE END
'------------------------------------------------------------------------------
    'As for the factor table: a rate below the first blank currency would be
    'dropped silently. The search stops at the factor table when that
    'follows.
        lastR = UsedLastRow(ws)
        sfRow = FindHeaderRow(ws, 1, HDR_SF)
        If sfRow > r Then
            lastR = sfRow - 1
        End If
        If ValuesAfterEnd(ws, r, lastR, 3, "FX") Then
            bad = True
        End If
        If bad Then
            Exit Function
        End If

'------------------------------------------------------------------------------
' ENSURE THE REPORTING CURRENCY
'------------------------------------------------------------------------------
        If KeyIndex(mFXIndex, pRepCcy) = 0 Then
            mFXCount = mFXCount + 1
            If mFXCount > UBound(mFXRate) Then
                ReDim Preserve mFXRate(1 To mFXCount * 2)
            End If
            mFXRate(mFXCount) = 1#
            KeyAdd mFXIndex, pRepCcy, mFXCount
        ElseIf Abs(mFXRate(KeyIndex(mFXIndex, pRepCcy)) - 1#) > 0.0000001 Then
            LogMsg SEV_WARN, SH_PARAMS, pRepCcy, "Rate of the reporting currency is not 1."
        End If
        LoadFXTable = True

End Function


Private Function ValuesAfterEnd( _
    ByVal ws As Worksheet, _
    ByVal endRow As Long, _
    ByVal lastR As Long, _
    ByVal valueCol As Long, _
    ByVal tableName As String) _
    As Boolean
'
'==============================================================================
'                                ValuesAfterEnd
'------------------------------------------------------------------------------
' PURPOSE
'   Find rows of a Params table that its loader will not read: from the
'   blank key that ends the table down to lastR, any row with a number in
'   the table's value column. Each is logged as an error (#35).
'
' INPUTS
'   ws: Params sheet.
'   endRow: the row with the blank key that ended the table.
'   lastR: last row to search.
'   valueCol: the column that always holds a number in a table row:
'   4 (factor) or 3 (rate).
'   tableName: for the messages.
'
' RETURNS
'   True when such a row was found.
'
' UPDATED
'   2026-10-07
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim r     As Long      'Row being checked
    Dim k     As String    'Key in column A, upper case

'------------------------------------------------------------------------------
' SEARCH
'------------------------------------------------------------------------------
        For r = endRow To lastR
            k = UTxt(ws.Cells(r, 1).Value)
            If IsNum(ws.Cells(r, valueCol).Value) And Not (valueCol = PRM_VALUE_COL And IsParameterCode(k)) Then
                ValuesAfterEnd = True
                If Len(k) = 0 Then
                    LogMsg SEV_ERROR, SH_PARAMS, "", tableName & " row has a value but no key - not read.", r
                Else
                    LogMsg SEV_ERROR, SH_PARAMS, k, tableName & " row below a blank row - not read; " & _
                           "remove the blank row.", r
                End If
            End If
        Next r

End Function


Private Function FXRate( _
    ByVal ccy As String) _
    As Double
'
'==============================================================================
'                                    FXRate
'------------------------------------------------------------------------------
' PURPOSE
'   Look up the conversion rate of a currency.
'
' INPUTS
'   ccy: currency code, upper case; blank means the reporting currency.
'
' RETURNS
'   Units of reporting currency per one unit of ccy; -1 when unknown.
'
' UPDATED
'   2026-10-06
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim i   As Long    'Position of the currency in mFXRate; 0 when unknown

'------------------------------------------------------------------------------
' LOOK UP
'------------------------------------------------------------------------------
        If Len(ccy) = 0 Then
            ccy = pRepCcy
        End If
        i = KeyIndex(mFXIndex, ccy)
        If i = 0 Then
            FXRate = -1#
        Else
            FXRate = mFXRate(i)
        End If

End Function


Private Function LoadNettingSets() As Boolean
'
'==============================================================================
'                               LoadNettingSets
'------------------------------------------------------------------------------
' PURPOSE
'   Read every netting set from the NettingSets sheet, apply the defaults,
'   and work out the effective MPOR and margined maturity factor of each
'   margined netting set.
'
' RETURNS
'   True when at least one valid netting set was read. False, with an error
'   logged, otherwise.
'
' STATE OWNERSHIP
'   Fills mNS, mNSCount and mNSIndex, and sizes the asset-class add-on
'   arrays. A duplicate ID is an error and the later row is ignored.
'
' REFERENCE
'   MPOR floors CRE52.50 to CRE52.52 and CRR Art. 285(2) to (5); the
'   client-clearing floor CRE54.12 and CRR Art. 279c(1), configurable on
'   Params. VM on an unmargined set CRE52.2, CRE52.10 and CRR Art. 272(7),
'   (7a). Margin-agreement structures CRE52.74, CRE52.75 and CRR Art.
'   274(4), 275(3).
'
' UPDATED
'   2026-10-10
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ws        As Worksheet    'NettingSets sheet
    Dim lastR     As Long         'Last row with a netting-set ID
    Dim n         As Long         'Number of rows to read
    Dim i         As Long         'Row within data
    Dim rowNum    As Long         'Sheet row of data row i
    Dim data      As Variant      'Input block, (row, column)
    Dim id        As String       'Netting-set ID, upper case
    Dim floorBD   As Double       'MPOR floor before remargining, business days
    Dim mpor      As Double       'Effective MPOR, business days
    Dim ovr       As Double       'MPOR override entered; 0 when none
    Dim rg        As String       'Regime of the netting set
    Dim nsErrors  As Long         'Invalid fields found on the row

'------------------------------------------------------------------------------
' READ THE INPUT BLOCK
'------------------------------------------------------------------------------
        Set ws = GetSheet(SH_NS)
        lastR = LastDataRowAny(ws, FIRST_DATA_ROW, NS_NCOLS, 0)
        If lastR < FIRST_DATA_ROW Then
            LogMsg SEV_ERROR, SH_NS, "", "No netting sets defined."
            Exit Function
        End If
        n = lastR - FIRST_DATA_ROW + 1
        data = ws.Range(ws.Cells(FIRST_DATA_ROW, 1), ws.Cells(lastR, NS_NCOLS)).Value

'------------------------------------------------------------------------------
' READ EACH NETTING SET
'------------------------------------------------------------------------------
    'An empty row is skipped; a row with data but no ID is an error, so a
    'netting set cannot silently drop out (#35). A repeated ID is ignored and
    'makes the first set with that ID INVALID: which row's terms apply to
    'its trades is ambiguous, so its EAD is withheld.
        For i = 1 To n
            rowNum = FIRST_DATA_ROW + i - 1
            id = UTxt(data(i, NS_ID))
            If Len(id) = 0 Then
                If RowHasData(data, i, NS_NCOLS, 0) Then
                    LogMsg SEV_ERROR, SH_NS, "", "Row has netting-set data but no ID - row ignored.", rowNum
                End If
                GoTo NextRow
            End If
            If KeyIndex(mNSIndex, id) > 0 Then
                With mNS(KeyIndex(mNSIndex, id))
                    .InputErrors = .InputErrors + 1
                End With
                LogMsg SEV_ERROR, SH_NS, id, "Netting set ID listed twice - row ignored and EAD withheld.", rowNum
                GoTo NextRow
            End If
            mNSCount = mNSCount + 1
            If mNSCount > UBound(mNS) Then
                ReDim Preserve mNS(1 To mNSCount * 2)
            End If
            With mNS(mNSCount)

    'Plain fields. A blank field takes its default; a value that cannot be
    'read is an error that makes the netting set INVALID (#35). The
    'remargining frequency is at least 1 business day; a non-positive alpha
    'takes the Params value.
                nsErrors = 0
                .ID = id
                .Counterparty = SafeStr(data(i, NS_CPTY))
                .Margined = NsFlag(data(i, NS_MARGINED), False, "Margined", id, rowNum, nsErrors)
                .ClearingRole = NsCode(data(i, NS_CLEARED), CL_NONE, CL_NONE & "|" & CL_CM & "|" & CL_CLIENT, _
                                       "Client clearing role", id, rowNum, nsErrors)
                .RemarginBD = NsNumber(data(i, NS_FREQ), 1#, "Remargin frequency", id, rowNum, nsErrors)
                If .RemarginBD < 1# Then
                    .RemarginBD = 1#
                End If
                .LargeOrIlliquid = NsFlag(data(i, NS_LARGE), False, "Large or illiquid", id, rowNum, nsErrors)
                .Disputes = NsFlag(data(i, NS_DISPUTE), False, "Margin disputes", id, rowNum, nsErrors)
                .VM = NsNumber(data(i, NS_VM), 0#, "Net VM", id, rowNum, nsErrors)
                .NICA = NsNumber(data(i, NS_NICA), 0#, "NICA", id, rowNum, nsErrors)
                .TH = NsNumber(data(i, NS_TH), 0#, "Threshold", id, rowNum, nsErrors)
                .MTA = NsNumber(data(i, NS_MTA), 0#, "MTA", id, rowNum, nsErrors)
                .Alpha = NsNumber(data(i, NS_ALPHA), pAlpha, "Alpha override", id, rowNum, nsErrors)
                If .Alpha <= 0# Then
                    .Alpha = pAlpha
                End If

    'Regime: blank takes the Params default; any other value than BCBS or
    'CRR is an input error.
                rg = UTxt(data(i, NS_REGIME))
                If Len(rg) = 0 Then
                    rg = pRegime
                ElseIf rg <> RG_BCBS And rg <> RG_CRR Then
                    nsErrors = nsErrors + 1
                    LogMsg SEV_ERROR, SH_NS, id, "Regime override must be BCBS or CRR (found " & _
                           DescribeValue(data(i, NS_REGIME)) & ").", rowNum
                    rg = pRegime
                End If
                .Regime = rg
                .IsCRR = (rg = RG_CRR)

    'VM on an unmargined netting set. VM received means the counterparty
    'posts VM, so the set is margined [CRE52.2; CRR Art. 272(7)]. VM
    'posted is the one-way agreement in favour of the counterparty: under
    'BCBS the set is unmargined and C includes that VM with a negative sign
    '[CRE52.2, CRE52.10 footnote 2]; under CRR a one-way agreement is a
    'margin agreement [Art. 272(7a), 275(2)], so the set must be margined.
                If Not .Margined And .VM > 0# Then
                    nsErrors = nsErrors + 1
                    LogMsg SEV_ERROR, SH_NS, id, "VM received on an unmargined netting set: a counterparty " & _
                           "that posts VM makes the netting set margined - set Margined = Y.", rowNum
                ElseIf Not .Margined And .VM < 0# Then
                    If .IsCRR Then
                        nsErrors = nsErrors + 1
                        LogMsg SEV_ERROR, SH_NS, id, "VM posted on an unmargined CRR netting set: under CRR " & _
                               "a one-way margin agreement is a margin agreement (Art. 272(7a)) - set Margined = Y.", rowNum
                    Else
                        LogMsg SEV_INFO, SH_NS, id, "VM posted under a one-way margin agreement: unmargined " & _
                               "netting set, posted VM included in C with a negative sign (CRE52.10).", rowNum
                    End If
                End If

    'Margin-agreement structures the engine does not model are declared, so
    'that the netting set is rejected rather than calculated as one with a
    'single agreement of its own.
                Select Case NsCode(data(i, NS_STRUCT), "", MS_MIXED & "|" & MS_SHARED, _
                                   "Margin agreements", id, rowNum, nsErrors)
                    Case MS_MIXED
                        nsErrors = nsErrors + 1
                        LogMsg SEV_ERROR, SH_NS, id, "Several margin agreements, or margined and unmargined " & _
                               "trades, in one netting set are not supported (CRE52.74; CRR Art. 274(4)).", rowNum
                    Case MS_SHARED
                        nsErrors = nsErrors + 1
                        LogMsg SEV_ERROR, SH_NS, id, "One margin agreement over several netting sets is not " & _
                               "supported (CRE52.75; CRR Art. 275(3), 278(2)).", rowNum
                End Select
                .V = 0#
                .Trades = 0
                .Rejected = 0
                ovr = NsNumber(data(i, NS_MPOR), 0#, "MPOR override", id, rowNum, nsErrors)
                .InputErrors = nsErrors
                .MPOR = 0#
                .MFMargined = 0#

    'Effective MPOR of a margined netting set: the floor (client clearing
    'or bilateral, raised for large or illiquid sets) plus the remargining
    'period minus one day, doubled for disputes. The override, read and
    'validated above for every set, is accepted only if it is not below
    'that. The client-clearing floor applies to the clearing member facing
    'its client [CRE54.12; CRR Art. 279c(1)], and under CRR also to the
    'client facing its clearing member; under BCBS the client keeps the
    'bilateral floor [CRE54.8, CRE54.14].
                If .Margined Then
                    Select Case .ClearingRole
                        Case CL_CM
                            floorBD = pMPORClr
                        Case CL_CLIENT
                            If .IsCRR Then
                                floorBD = pMPORClr
                            Else
                                floorBD = pMPORBil
                                LogMsg SEV_INFO, SH_NS, id, "Client facing its clearing member under BCBS: " & _
                                       "bilateral MPOR floor; the client-clearing floor is for the clearing " & _
                                       "member (CRE54.12).", rowNum
                            End If
                        Case Else
                            floorBD = pMPORBil
                    End Select
                    If .LargeOrIlliquid Then
                        floorBD = Max2(floorBD, pMPORLarge)
                    End If
                    mpor = floorBD + .RemarginBD - 1#
                    If .Disputes Then
                        mpor = 2# * mpor
                    End If
                    If ovr > 0# Then
                        If ovr < mpor Then
                            LogMsg SEV_WARN, SH_NS, id, "MPOR override " & CStr(ovr) & _
                                   " BD is below the regulatory floor - " & CStr(mpor) & " BD applied.", rowNum
                        Else
                            mpor = ovr
                        End If
                    End If
                    .MPOR = mpor
                    .MFMargined = SACCR_MaturityFactor(0#, True, .MPOR, pMinMatBD, pBDYear)
                End If
            End With
            KeyAdd mNSIndex, id, mNSCount
NextRow:
        Next i

'------------------------------------------------------------------------------
' CHECK AND SIZE THE ADD-ON ARRAYS
'------------------------------------------------------------------------------
        If mNSCount = 0 Then
            LogMsg SEV_ERROR, SH_NS, "", "No valid netting sets."
            Exit Function
        End If
        ReDim mAddOnU(1 To mNSCount, 1 To AC_COUNT)
        ReDim mAddOnM(1 To mNSCount, 1 To AC_COUNT)
        LoadNettingSets = True

End Function


'
'------------------------------------------------------------------------------
'
'                           TRADE-LEVEL CALCULATION
'
'------------------------------------------------------------------------------
'

Private Function NsFlag( _
    ByVal v As Variant, _
    ByVal dflt As Boolean, _
    ByVal fieldName As String, _
    ByVal nsId As String, _
    ByVal rowNum As Long, _
    ByRef nsErrors As Long) _
    As Boolean
'
'==============================================================================
'                                    NsFlag
'------------------------------------------------------------------------------
' PURPOSE
'   Read a Y/N field of a netting set strictly: blank takes the default; an
'   unrecognised value is logged and counted as an input error (#35).
'
' UPDATED
'   2026-10-07
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim result   As Boolean    'Value read

'------------------------------------------------------------------------------
' READ
'------------------------------------------------------------------------------
        If Not TryBool(v, dflt, result) Then
            nsErrors = nsErrors + 1
            LogMsg SEV_ERROR, SH_NS, nsId, fieldName & " must be Y or N (found " & DescribeValue(v) & ").", rowNum
        End If
        NsFlag = result

End Function


Private Function NsCode( _
    ByVal v As Variant, _
    ByVal dflt As String, _
    ByVal allowed As String, _
    ByVal fieldName As String, _
    ByVal nsId As String, _
    ByVal rowNum As Long, _
    ByRef nsErrors As Long) _
    As String
'
'==============================================================================
'                                    NsCode
'------------------------------------------------------------------------------
' PURPOSE
'   Read a coded field of a netting set strictly: blank takes the default;
'   a value that is not one of the allowed codes, or an Excel error value,
'   is logged and counted as an input error (#35).
'
' INPUTS
'   allowed: the accepted codes, upper case, separated by "|".
'
' RETURNS
'   The code in upper case, or the default.
'
' UPDATED
'   2026-10-10
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim code   As String    'Value read, upper case

'------------------------------------------------------------------------------
' READ
'------------------------------------------------------------------------------
        code = UTxt(v)
        If IsError(v) Or (Len(code) > 0 And InStr(1, "|" & allowed & "|", "|" & code & "|") = 0) Then
            nsErrors = nsErrors + 1
            LogMsg SEV_ERROR, SH_NS, nsId, fieldName & " must be " & Replace(allowed, "|", ", ") & _
                   " or blank (found " & DescribeValue(v) & ").", rowNum
            NsCode = dflt
        ElseIf Len(code) = 0 Then
            NsCode = dflt
        Else
            NsCode = code
        End If

End Function


Private Function NsNumber( _
    ByVal v As Variant, _
    ByVal dflt As Double, _
    ByVal fieldName As String, _
    ByVal nsId As String, _
    ByVal rowNum As Long, _
    ByRef nsErrors As Long) _
    As Double
'
'==============================================================================
'                                   NsNumber
'------------------------------------------------------------------------------
' PURPOSE
'   Read a numeric field of a netting set strictly: blank takes the
'   default; text that is not a number is logged and counted as an input
'   error (#35).
'
' UPDATED
'   2026-10-07
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim result   As Double    'Value read

'------------------------------------------------------------------------------
' READ
'------------------------------------------------------------------------------
        If Not TryDbl(v, dflt, result) Then
            nsErrors = nsErrors + 1
            LogMsg SEV_ERROR, SH_NS, nsId, fieldName & " must be a number (found " & DescribeValue(v) & ").", rowNum
        End If
        NsNumber = result

End Function


Private Function TradeDate( _
    ByVal v As Variant, _
    ByVal fieldName As String, _
    ByRef ok As Boolean, _
    ByRef msg As String) _
    As Double
'
'==============================================================================
'                                  TradeDate
'------------------------------------------------------------------------------
' PURPOSE
'   Read a trade date: blank is "not given" (-1); a value that is not a
'   date makes the trade invalid instead of being treated as blank (#35).
'
' RETURNS
'   The date serial, or -1 when blank or unreadable.
'
' UPDATED
'   2026-10-07
'==============================================================================
'

'------------------------------------------------------------------------------
' READ
'------------------------------------------------------------------------------
        TradeDate = ToSerial(v)
        If TradeDate < 0# Then
            If IsError(v) Or Not IsBlankCell(v) Then
                AddErr ok, msg, fieldName & " is not a date (found " & DescribeValue(v) & ")"
            End If
            TradeDate = -1#
        End If

End Function


Private Function DescribeValue( _
    ByVal v As Variant) _
    As String
'
'==============================================================================
'                                DescribeValue
'------------------------------------------------------------------------------
' PURPOSE
'   Describe a cell value for a message: its text in quotes, or "an error
'   value" for #N/A and the like.
'
' UPDATED
'   2026-10-07
'==============================================================================
'
        If IsError(v) Then
            DescribeValue = "an error value"
        Else
            DescribeValue = "'" & SafeStr(v) & "'"
        End If

End Function


Private Sub AddErr( _
    ByRef ok As Boolean, _
    ByRef msg As String, _
    ByVal txt As String)
'
'==============================================================================
'                                    AddErr
'------------------------------------------------------------------------------
' PURPOSE
'   Mark the current trade as invalid and append the reason.
'
' INPUTS
'   ok: set to False.
'   msg: the trade's reasons so far; txt is appended after "; ".
'   txt: the reason.
'
' UPDATED
'   2026-10-06
'==============================================================================
'

'------------------------------------------------------------------------------
' APPEND
'------------------------------------------------------------------------------
        ok = False
        If Len(msg) > 0 Then
            msg = msg & "; "
        End If
        msg = msg & txt

End Sub


Private Sub AddWarn( _
    ByRef warn As String, _
    ByVal txt As String)
'
'==============================================================================
'                                   AddWarn
'------------------------------------------------------------------------------
' PURPOSE
'   Append a warning for the current trade; the trade stays valid.
'
' INPUTS
'   warn: the trade's warnings so far; txt is appended after "; ".
'   txt: the warning.
'
' UPDATED
'   2026-10-06
'==============================================================================
'

'------------------------------------------------------------------------------
' APPEND
'------------------------------------------------------------------------------
        If Len(warn) > 0 Then
            warn = warn & "; "
        End If
        warn = warn & txt

End Sub


Private Sub ProcessTrades( _
    ByVal writeOutputs As Boolean)
'
'==============================================================================
'                                ProcessTrades
'------------------------------------------------------------------------------
' PURPOSE
'   Validate and calculate every trade on the Trades sheet: classification,
'   supervisory factor, amounts in the reporting currency, time to maturity,
'   supervisory delta, adjusted notional, maturity factor and effective
'   notional. Each valid trade is added to its bucket; every trade, valid or
'   not, gets one TradeCalc row.
'
' INPUTS
'   writeOutputs: True writes the TradeCalc sheet.
'
' STATE OWNERSHIP
'   Adds to mBk through AddToBucket, adds MtM and trade counts to mNS, and
'   updates mTradesRead and mTradesUsed. An invalid trade is logged with all
'   its reasons and left out of the calculation.
'
' REFERENCE
'   CRE52.30 to CRE52.52 for the trade-level quantities; CRR Art. 277 to
'   280f for OT and the regime rules.
'
' UPDATED
'   2026-10-09
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    'Sheet access and output.
    Dim ws            As Worksheet    'Trades sheet
    Dim wsOut         As Worksheet    'TradeCalc sheet
    Dim lastR         As Long         'Last row with a trade ID
    Dim n             As Long         'Number of rows to read
    Dim i             As Long         'Row within data
    Dim rowNum        As Long         'Sheet row of data row i
    Dim data          As Variant      'Input block, (row, column)
    Dim outArr()      As Variant      'TradeCalc rows, (row, column)
    Dim nOut          As Long         'TradeCalc rows filled
    Dim seenIds       As Collection   'Trade IDs seen so far, mapped to their sheet row
    Dim firstRow      As Long         'Sheet row where a duplicate ID first appeared

    'Validation state of the current trade.
    Dim ok            As Boolean      'False once any error is found
    Dim msg           As String       'Error reasons, separated by "; "
    Dim warn          As String       'Warnings, separated by "; "

    'Classification, as entered.
    Dim tid           As String       'Trade ID
    Dim nsId          As String       'Netting-set ID, upper case
    Dim acTxt         As String       'Asset-class code, upper case
    Dim subCls        As String       'Sub-class, upper case
    Dim rf            As String       'Risk factor or reference, upper case
    Dim instType      As String       'LINEAR, OPTION or CDO
    Dim dirTxt        As String       'Direction text, upper case
    Dim optTxt        As String       'CALL or PUT
    Dim nature        As String       'STANDARD, BASIS or VOLATILITY
    Dim lbl           As String       'Basis or volatility hedging-set label

    'Classification, resolved.
    Dim nsIdx         As Long         'Netting set, position in mNS; 0 when unknown
    Dim ac            As Long         'Asset class, an AC_ constant; 0 when unknown
    Dim sfKey         As String       'Key into the supervisory-factor table
    Dim sfIdx         As Long         'Position in mSF; 0 when not found
    Dim isLong        As Boolean      'Long (bought) position
    Dim natFactor     As Double       'Factor multiplier: 1, basis or volatility
    Dim natTag        As String       'Hedging-set suffix: STD, BASIS:label or VOL:label

    'Amounts.
    Dim notional      As Double       'Notional, in its currency
    Dim nccy          As String       'Notional currency
    Dim fxN           As Double       'Rate of the notional currency
    Dim notionalRep   As Double       'Notional in the reporting currency
    Dim mtm           As Double       'MtM, in its currency
    Dim mccy          As String       'MtM currency
    Dim fxM           As Double       'Rate of the MtM currency
    Dim mtmRep        As Double       'MtM in the reporting currency

    'Dates as serials (-1 when missing) and as years from the reporting date.
    Dim dStart        As Double       'Start date
    Dim dEnd          As Double       'End date
    Dim dMat          As Double       'Maturity date
    Dim dExp          As Double       'Option expiry date
    Dim S             As Double       'Start S, years, at least 0
    Dim E             As Double       'End E, years
    Dim M             As Double       'Maturity M, years
    Dim T             As Double       'Option expiry T, years

    'Trade-level SA-CCR quantities.
    Dim sd            As Double       'Supervisory duration (IR and credit)
    Dim adjN          As Double       'Adjusted notional d
    Dim delta         As Double       'Supervisory delta
    Dim vDelta        As Variant      'Option delta as returned by SACCR_OptionDelta
    Dim mfU           As Double       'Maturity factor, unmargined
    Dim mfM           As Double       'Maturity factor, margined
    Dim enU           As Double       'Effective notional delta * d * MF, unmargined
    Dim enM           As Double       'Effective notional delta * d * MF, margined
    Dim hsKey         As String       'Hedging-set key
    Dim subKey        As String       'Bucket key within the hedging set
    Dim bucketNo      As Long         'IR maturity bucket 1, 2 or 3
    Dim sfEff         As Double       'Supervisory factor times natFactor
    Dim corrEff       As Double       'Correlation
    Dim pairSign      As Double       'FX: +1 pair kept, -1 inverted, 0 invalid

    'Option and CDO inputs.
    Dim P             As Double       'Underlying price
    Dim K             As Double       'Strike
    Dim lam           As Double       'Lambda shift applied
    Dim attA          As Double       'CDO attachment point
    Dim detD          As Double       'CDO detachment point
    Dim lamIn         As Variant      'Lambda as entered
    Dim lamOut        As Variant      'Lambda reported; Empty for non-options

'------------------------------------------------------------------------------
' READ THE INPUT BLOCK
'------------------------------------------------------------------------------
    'With no trades there is nothing to calculate; the old TradeCalc rows
    'are cleared so they cannot be mistaken for results.
        Set ws = GetSheet(SH_TRADES)
        lastR = LastDataRowAny(ws, FIRST_DATA_ROW, TR_NCOLS, TR_COMMENT)
        If lastR < FIRST_DATA_ROW Then
            LogMsg SEV_ERROR, SH_TRADES, "", "No trades found."
            If writeOutputs Then
                ClearOutputBlock GetSheet(SH_TRADECALC), FIRST_DATA_ROW, TC_NCOLS
            End If
            Exit Sub
        End If
        n = lastR - FIRST_DATA_ROW + 1
        data = ws.Range(ws.Cells(FIRST_DATA_ROW, 1), ws.Cells(lastR, TR_NCOLS)).Value
        ReDim outArr(1 To n, 1 To TC_NCOLS)
        nOut = 0

'------------------------------------------------------------------------------
' PROCESS EACH TRADE
'------------------------------------------------------------------------------
    'An empty row is skipped. A row with data but no trade ID is rejected,
    'and so is a repeated trade ID; either makes its netting set INCOMPLETE
    '(#35). Every per-trade value is reset so nothing leaks from the
    'previous trade into the output row.
        Set seenIds = New Collection
        For i = 1 To n
            rowNum = FIRST_DATA_ROW + i - 1
            tid = SafeStr(data(i, TR_ID))
            If Len(tid) = 0 Then
                If RowHasData(data, i, TR_NCOLS, TR_COMMENT) Then
                    mTradesRead = mTradesRead + 1
                    LogMsg SEV_ERROR, SH_TRADES, "", "Row has trade data but no Trade ID - trade rejected.", rowNum
                    nsIdx = KeyIndex(mNSIndex, UTxt(data(i, TR_NS)))
                    If nsIdx > 0 Then mNS(nsIdx).Rejected = mNS(nsIdx).Rejected + 1
                End If
                GoTo NextTrade
            End If
            mTradesRead = mTradesRead + 1
            ok = True
            msg = ""
            warn = ""
            S = 0#
            E = 0#
            M = 0#
            T = 0#
            sd = 0#
            adjN = 0#
            delta = 0#
            mfU = 0#
            mfM = 0#
            enU = 0#
            enM = 0#
            bucketNo = 0
            sfEff = 0#
            corrEff = 0#
            hsKey = ""
            subKey = ""
            sfIdx = 0
            notionalRep = 0#
            mtmRep = 0#
            pairSign = 1#
            lam = 0#
            lamOut = Empty

    '--- Classification ------------------------------------------------------
    'Netting set, asset class, risk factor, instrument, direction and
    'nature. Basis and volatility trades form their own hedging sets, one
    'per label, with the factor multiplied by BasisFactor or
    'VolatilityFactor [CRE52.46, CRE52.47].
            firstRow = KeyIndex(seenIds, tid)
            If firstRow > 0 Then
                AddErr ok, msg, "duplicate Trade ID (first used on row " & firstRow & ")"
            Else
                KeyAdd seenIds, tid, rowNum
            End If

            nsId = UTxt(data(i, TR_NS))
            nsIdx = KeyIndex(mNSIndex, nsId)
            If nsIdx = 0 Then AddErr ok, msg, "unknown netting set '" & nsId & "'"

            acTxt = UTxt(data(i, TR_AC))
            ac = ACIndex(acTxt)
            If ac = 0 Then AddErr ok, msg, "asset class must be IR, FX, CR, EQ, CO or OT"

            subCls = UTxt(data(i, TR_SUB))
            rf = UTxt(data(i, TR_RF))
            If Len(rf) = 0 Then AddErr ok, msg, "risk factor / reference missing"
            If InStr(rf, "|") > 0 Or InStr(rf, "#") > 0 Then
                AddErr ok, msg, "risk factor / reference must not contain '|' or '#'"
            End If
            If ac = AC_IR And Len(rf) > 0 And _
               Not (rf Like "[A-Z][A-Z][A-Z]" Or rf Like "[A-Z][A-Z][A-Z]-INFL") Then
                AddErr ok, msg, "interest-rate risk factor must be a 3-letter currency code such as EUR, " & _
                       "or EUR-INFL for inflation"
            End If

            instType = UTxt(data(i, TR_INSTR))
            If Len(instType) = 0 Then instType = "LINEAR"
            If instType <> "LINEAR" And instType <> "OPTION" And instType <> "CDO" Then
                AddErr ok, msg, "instrument type must be Linear, Option or CDO"
            End If
            If instType = "CDO" And ac <> AC_CR Then AddErr ok, msg, "CDO tranches are credit (CR) trades"

            dirTxt = UTxt(data(i, TR_DIR))
            Select Case dirTxt
                Case "LONG", "BUY", "BOUGHT", "L", "B"
                    isLong = True
                Case "SHORT", "SELL", "SOLD", "S"
                    isLong = False
                Case Else
                    AddErr ok, msg, "direction must be Long or Short"
            End Select

            nature = UTxt(data(i, TR_NATURE))
            If Len(nature) = 0 Then nature = "STANDARD"
            lbl = UTxt(data(i, TR_LABEL))
            If InStr(lbl, "|") > 0 Or InStr(lbl, "#") > 0 Then
                AddErr ok, msg, "hedging-set label must not contain '|' or '#'"
            End If
            Select Case nature
                Case "STANDARD"
                    natFactor = 1#
                    natTag = "STD"
                Case "BASIS"
                    natFactor = pBasisF
                    natTag = "BASIS:" & lbl
                    If Len(lbl) = 0 Then AddErr ok, msg, "basis trade needs a hedging-set label"
                Case "VOLATILITY"
                    natFactor = pVolF
                    natTag = "VOL:" & lbl
                    If Len(lbl) = 0 Then AddErr ok, msg, "volatility trade needs a hedging-set label"
                Case Else
                    AddErr ok, msg, "nature must be Standard, Basis or Volatility"
            End Select

    '--- Supervisory parameters ----------------------------------------------
    'IR, FX and OT have one factor each; the other classes are looked up by
    'class and sub-class, for example CR_AAA.
            If ac > 0 Then
                Select Case ac
                    Case AC_IR
                        sfKey = "IR"
                    Case AC_FX
                        sfKey = "FX"
                    Case AC_OT
                        sfKey = "OT"
                    Case Else
                        sfKey = acTxt & "_" & subCls
                End Select
                sfIdx = KeyIndex(mSFIndex, sfKey)
                If sfIdx = 0 Then AddErr ok, msg, "no supervisory factor for '" & sfKey & "' (check sub-class)"
            End If

    '--- Regime-specific availability ----------------------------------------
    'OT exists only under CRR, and a factor-table row limited to one regime
    'cannot be used by a netting set under the other.
            If nsIdx > 0 And ac = AC_OT Then
                If Not mNS(nsIdx).IsCRR Then
                    AddErr ok, msg, "asset class OT (other risks) exists only under CRR (Art. 277(1)(f), 280f); " & _
                           "under BCBS map the trade to the class of its primary risk driver (CRE52)"
                End If
            End If
            If nsIdx > 0 And sfIdx > 0 Then
                If Len(mSF(sfIdx).Regimes) > 0 And mSF(sfIdx).Regimes <> mNS(nsIdx).Regime Then
                    AddErr ok, msg, "'" & sfKey & "' is defined only for regime " & mSF(sfIdx).Regimes & _
                           " (netting set uses " & mNS(nsIdx).Regime & ")"
                End If
            End If

    '--- Amounts ---------------------------------------------------------------
    'The notional is unsigned; Direction gives the sign. A missing MtM is an
    'error, not an assumed 0 (#35); a blank MtM currency means the notional
    'currency. Both amounts are converted to the reporting currency.
            If IsNum(data(i, TR_NOTIONAL)) Then
                notional = CDbl(data(i, TR_NOTIONAL))
                If notional < 0# Then AddErr ok, msg, "notional must be positive (use Direction for the sign)"
            ElseIf IsBlankCell(data(i, TR_NOTIONAL)) And Not IsError(data(i, TR_NOTIONAL)) Then
                AddErr ok, msg, "notional missing"
            Else
                AddErr ok, msg, "notional must be a number (found " & DescribeValue(data(i, TR_NOTIONAL)) & ")"
            End If
            nccy = UTxt(data(i, TR_NCCY))
            fxN = FXRate(nccy)
            If fxN < 0# Then AddErr ok, msg, "no FX rate for notional currency '" & nccy & "'"
            If IsNum(data(i, TR_MTM)) Then
                mtm = CDbl(data(i, TR_MTM))
            ElseIf IsBlankCell(data(i, TR_MTM)) And Not IsError(data(i, TR_MTM)) Then
                mtm = 0#
                AddErr ok, msg, "MtM missing (enter 0 if the trade has no value)"
            Else
                mtm = 0#
                AddErr ok, msg, "MtM must be a number (found " & DescribeValue(data(i, TR_MTM)) & ")"
            End If
            mccy = UTxt(data(i, TR_MCCY))
            If Len(mccy) = 0 Then mccy = nccy
            fxM = FXRate(mccy)
            If fxM < 0# Then AddErr ok, msg, "no FX rate for MtM currency '" & mccy & "'"
            If ok Then
                notionalRep = notional * fxN
                mtmRep = mtm * fxM
            End If

    '--- Dates to year fractions -----------------------------------------------
    'A date that is present but cannot be read is an error (#35). Missing
    'dates are filled from each other: maturity from the end date,
    'then from the option expiry; the end date from the maturity. S, E and M
    'are years from the reporting date on the DaysPerYear basis, with S
    'floored at 0.
            dStart = TradeDate(data(i, TR_START), "start date", ok, msg)
            dEnd = TradeDate(data(i, TR_END), "end date", ok, msg)
            dMat = TradeDate(data(i, TR_MAT), "maturity date", ok, msg)
            dExp = TradeDate(data(i, TR_EXPIRY), "option expiry", ok, msg)
            If dMat < 0# Then dMat = dEnd
            If dMat < 0# Then dMat = dExp
            If dEnd < 0# Then dEnd = dMat
            If dMat < 0# Then
                AddErr ok, msg, "maturity / end date missing"
            Else
                M = (dMat - pAsOf) / pDaysYear
                E = (dEnd - pAsOf) / pDaysYear
                If dStart >= 0# Then
                    S = Max2(0#, (dStart - pAsOf) / pDaysYear)
                Else
                    S = 0#
                End If
                If M <= 0# Then
                    AddErr ok, msg, "trade has matured (maturity <= reporting date)"
                ElseIf (ac = AC_IR Or ac = AC_CR) And E <= S Then
                    AddErr ok, msg, "end date must be after start date"
                End If
            End If

    '--- Supervisory delta -------------------------------------------------------
    'Linear trades: +1 long, -1 short [CRE52.39]. Options: the option delta
    'with the lambda shift; under CRR the regulatory lambda replaces any
    'entered value for IR and commodity options [CRE52.40]. CDO tranches:
    'the tranche delta [CRE52.41].
            If ok Then
                Select Case instType
                    Case "LINEAR"
                        If isLong Then
                            delta = 1#
                        Else
                            delta = -1#
                        End If
                    Case "OPTION"
                        optTxt = UTxt(data(i, TR_OPT))
                        If optTxt <> "CALL" And optTxt <> "PUT" Then AddErr ok, msg, "option type must be Call or Put"
                        If dExp < 0# Then
                            AddErr ok, msg, "option expiry missing"
                        Else
                            T = (dExp - pAsOf) / pDaysYear
                            If T <= 0# Then AddErr ok, msg, "option has expired"
                        End If
                        If Not IsNum(data(i, TR_PRICE)) Or Not IsNum(data(i, TR_STRIKE)) Then
                            AddErr ok, msg, "option needs underlying price and strike"
                        End If
                        If ok Then
                            P = CDbl(data(i, TR_PRICE))
                            K = CDbl(data(i, TR_STRIKE))
                            lamIn = data(i, TR_LAMBDA)
                            If Not IsNum(lamIn) And Not (IsBlankCell(lamIn) And Not IsError(lamIn)) Then
                                AddErr ok, msg, "lambda must be a number (found " & DescribeValue(lamIn) & ")"
                            End If
                            If mNS(nsIdx).IsCRR And (ac = AC_IR Or ac = AC_CO) Then
                                'Delegated Regulation (EU) 2021/931 Art. 5, as amended by 2025/855.
                                If ac = AC_IR Then
                                    lam = SACCR_LambdaCRR(P, K, True, pLamThrIR)
                                Else
                                    lam = SACCR_LambdaCRR(P, K, False, pLamThrCO)
                                End If
                                If IsNum(lamIn) Then
                                    If Abs(CDbl(lamIn) - lam) > 0.000000001 Then
                                        AddWarn warn, "lambda input ignored under CRR - RTS lambda applied"
                                    End If
                                End If
                            Else
                                lam = ToDbl(lamIn, 0#)
                            End If
                            lamOut = lam
                            If P + lam <= 0# Or K + lam <= 0# Then
                                AddErr ok, msg, "price and strike must be > 0 after lambda shift (BCBS: enter lambda per CRE52.40 FAQ)"
                            ElseIf mSF(sfIdx).Vol <= 0# Then
                                AddErr ok, msg, "supervisory volatility missing for '" & sfKey & "'"
                            Else
                                vDelta = SACCR_OptionDelta(P, K, T, mSF(sfIdx).Vol, (optTxt = "CALL"), isLong, lam)
                                delta = CDbl(vDelta)
                            End If
                        End If
                    Case "CDO"
                        If Not IsNum(data(i, TR_ATTACH)) Or Not IsNum(data(i, TR_DETACH)) Then
                            AddErr ok, msg, "CDO tranche needs attachment and detachment points"
                        Else
                            attA = CDbl(data(i, TR_ATTACH))
                            detD = CDbl(data(i, TR_DETACH))
                            If attA < 0# Or detD > 1# Or attA >= detD Then
                                AddErr ok, msg, "attachment/detachment must satisfy 0 <= A < D <= 1"
                            Else
                                delta = CDbl(SACCR_CDODelta(attA, detD, isLong))
                            End If
                        End If
                End Select
            End If

    '--- Hedging set, adjusted notional and bucket -------------------------------
    'IR: d = notional * SD, hedging set per currency (risk factor), bucket
    'by end date: under 1 year, 1 to 5 years, over 5 years [CRE52.34,
    'CRE52.57(2)-(3)]; an inflation risk factor such as EUR-INFL is its own
    'hedging set [CRR Art. 277(4)(a), 277a(1)]. FX: d = notional, hedging
    'set per currency pair written in alphabetical order; an inverted pair
    'flips the delta [CRE52.58]. CR: d = notional * SD, one hedging set,
    'bucket per entity [CRE52.60, CRE52.64]. EQ: one hedging set, bucket per
    'entity [CRE52.66]. CO: hedging set per commodity group, bucket per
    'commodity [CRE52.70]. OT: hedging set per primary risk driver [CRR Art.
    '277a].
            If ok Then
                Select Case ac
                    Case AC_IR
                        sd = SACCR_SupervisoryDuration(S, E, pSDFloorBD / pBDYear)
                        adjN = notionalRep * sd
                        If E < 1# Then
                            bucketNo = 1
                            subKey = "<1Y"
                        ElseIf E <= 5# Then
                            bucketNo = 2
                            subKey = "1-5Y"
                        Else
                            bucketNo = 3
                            subKey = ">5Y"
                        End If
                        hsKey = "IR|" & rf & "|" & natTag
                    Case AC_FX
                        rf = NormalizePair(rf, pairSign)
                        If pairSign = 0# Then
                            AddErr ok, msg, "FX risk factor must be a currency pair such as EUR/USD"
                        End If
                        delta = delta * pairSign
                        adjN = notionalRep
                        hsKey = "FX|" & rf & "|" & natTag
                        subKey = rf
                    Case AC_CR
                        sd = SACCR_SupervisoryDuration(S, E, pSDFloorBD / pBDYear)
                        adjN = notionalRep * sd
                        hsKey = "CR|" & natTag
                        subKey = rf
                    Case AC_EQ
                        adjN = notionalRep
                        hsKey = "EQ|" & natTag
                        subKey = rf
                    Case AC_CO
                        adjN = notionalRep
                        If Len(mSF(sfIdx).Group) = 0 Then
                            AddErr ok, msg, "commodity hedging group missing for '" & sfKey & "'"
                        End If
                        hsKey = "CO|" & mSF(sfIdx).Group & "|" & natTag
                        subKey = rf
                    Case AC_OT
                        'Art. 277a: same hedging set only for an identical primary risk driver.
                        adjN = notionalRep
                        hsKey = "OT|" & rf & "|" & natTag
                        subKey = rf
                End Select
            End If

    '--- Maturity factor and effective notional ---------------------------------
    'The effective notional is kept on both bases: unmargined MF from the
    'trade's maturity, and margined MF from the netting set's MPOR. The
    'margined basis equals the unmargined one for an unmargined set.
            If ok Then
                sfEff = mSF(sfIdx).SF * natFactor
                corrEff = mSF(sfIdx).Corr
                mfU = SACCR_MaturityFactor(M, False, 0#, pMinMatBD, pBDYear)
                If mNS(nsIdx).Margined Then
                    mfM = mNS(nsIdx).MFMargined
                Else
                    mfM = mfU
                End If
                enU = delta * adjN * mfU
                enM = delta * adjN * mfM
                If ac = AC_CR Or ac = AC_EQ Or ac = AC_CO Then
                    CheckReferenceClass nsIdx, ac, rf, sfKey, tid, rowNum
                End If
                AddToBucket nsIdx, ac, hsKey, subKey, sfEff, corrEff, enU, enM, tid, rowNum
                mNS(nsIdx).V = mNS(nsIdx).V + mtmRep
                mNS(nsIdx).Trades = mNS(nsIdx).Trades + 1
                mTradesUsed = mTradesUsed + 1
            End If

            If Not ok Then LogMsg SEV_ERROR, SH_TRADES, tid, "Trade excluded: " & msg, rowNum
            If Not ok And nsIdx > 0 Then mNS(nsIdx).Rejected = mNS(nsIdx).Rejected + 1
            If Len(warn) > 0 Then LogMsg SEV_WARN, SH_TRADES, tid, warn, rowNum

    '--- TradeCalc row ------------------------------------------------------------
    'Columns: 1 ID, 2 netting set, 3 asset class, 4 status, 5 hedging set,
    '6 bucket, 7 effective SF, 8 notional, 9 MtM, 10 S, 11 E, 12 M, 13 T,
    '14 SD, 15 d, 16 delta, 17 MF unmargined, 18 MF margined, 19 effective
    'notional unmargined, 20 effective notional margined, 21 IR bucket,
    '22 stand-alone add-on, 23 lambda, 24 regime, 25 messages. Columns that
    'do not apply to the trade stay blank.
            nOut = nOut + 1
            outArr(nOut, 1) = tid
            outArr(nOut, 2) = nsId
            outArr(nOut, 3) = acTxt
            If ok Then
                outArr(nOut, 4) = "OK"
                outArr(nOut, 5) = hsKey
                outArr(nOut, 6) = subKey
                outArr(nOut, 7) = sfEff
                outArr(nOut, 8) = notionalRep
                outArr(nOut, 9) = mtmRep
                outArr(nOut, 10) = S
                outArr(nOut, 11) = E
                outArr(nOut, 12) = M
                If instType = "OPTION" Then outArr(nOut, 13) = T
                If ac = AC_IR Or ac = AC_CR Then outArr(nOut, 14) = sd
                outArr(nOut, 15) = adjN
                outArr(nOut, 16) = delta
                outArr(nOut, 17) = mfU
                If mNS(nsIdx).Margined Then outArr(nOut, 18) = mfM
                outArr(nOut, 19) = enU
                If mNS(nsIdx).Margined Then outArr(nOut, 20) = enM
                If ac = AC_IR Then outArr(nOut, 21) = bucketNo
                outArr(nOut, 22) = sfEff * Abs(enU)
                outArr(nOut, 23) = lamOut
                outArr(nOut, 24) = mNS(nsIdx).Regime
                outArr(nOut, 25) = warn
            Else
                outArr(nOut, 4) = "EXCLUDED"
                If nsIdx > 0 Then
                    outArr(nOut, 24) = mNS(nsIdx).Regime
                End If
                outArr(nOut, 25) = msg
            End If
NextTrade:
        Next i

'------------------------------------------------------------------------------
' WRITE TRADECALC
'------------------------------------------------------------------------------
        If writeOutputs Then
            Set wsOut = GetSheet(SH_TRADECALC)
            ClearOutputBlock wsOut, FIRST_DATA_ROW, TC_NCOLS
            WriteBlock wsOut, FIRST_DATA_ROW, outArr, nOut, TC_NCOLS
            FormatColumns wsOut, FIRST_DATA_ROW, nOut, _
                "@|@|@|@|@|@|0.00%|#,##0|#,##0|0.0000|0.0000|0.0000|0.0000|0.0000|#,##0|0.0000|0.0000|0.0000|#,##0|#,##0|0|#,##0|0.0000%|@|@"
        End If

End Sub


Private Function NormalizePair( _
    ByVal txt As String, _
    ByRef pairSgn As Double) _
    As String
'
'==============================================================================
'                                NormalizePair
'------------------------------------------------------------------------------
' PURPOSE
'   Write a currency pair in one canonical form, so that EUR/USD and
'   USD/EUR fall in the same hedging set.
'
' INPUTS
'   txt: the pair as entered, for example "EUR/USD", "EURUSD" or "eur-usd".
'
' RETURNS
'   The pair as "AAA/BBB" with the two codes in alphabetical order; txt
'   unchanged when it is not a valid pair.
'   pairSgn: +1 when the order was kept, -1 when it was inverted, 0 when txt
'   is not two different three-letter codes.
'
' UPDATED
'   2026-10-06
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim s    As String    'Pair without separators, upper case
    Dim c1   As String    'First currency
    Dim c2   As String    'Second currency

'------------------------------------------------------------------------------
' NORMALIZE
'------------------------------------------------------------------------------
    'Remove the separators "/", "-", " " and "." and expect six letters.
        s = UCase$(Replace(Replace(Replace(Replace(txt, "/", ""), "-", ""), " ", ""), ".", ""))
        If Len(s) <> 6 Then
            pairSgn = 0#
            NormalizePair = txt
            Exit Function
        End If
        c1 = Left$(s, 3)
        c2 = Right$(s, 3)
        If c1 = c2 Then
            pairSgn = 0#
            NormalizePair = txt
        ElseIf c1 < c2 Then
            pairSgn = 1#
            NormalizePair = c1 & "/" & c2
        Else
            pairSgn = -1#
            NormalizePair = c2 & "/" & c1
        End If

End Function


Private Sub CheckReferenceClass( _
    ByVal nsIdx As Long, _
    ByVal ac As Long, _
    ByVal rf As String, _
    ByVal sfKey As String, _
    ByVal tid As String, _
    ByVal rowNum As Long)
'
'==============================================================================
'                             CheckReferenceClass
'------------------------------------------------------------------------------
' PURPOSE
'   Make sure a credit or equity entity, or a commodity, has one sub-class
'   in its netting set. Its bucket takes one supervisory factor and
'   correlation; with two sub-classes the result would depend on which
'   trade comes first (#37).
'
' INPUTS
'   nsIdx, ac, rf: netting set, asset class and reference of the trade.
'   sfKey: the factor-table key the trade's sub-class gives, such as CR_AA.
'   tid, rowNum: trade ID and sheet row, for the message.
'
' STATE OWNERSHIP
'   Adds to mRefClass, mRefCount and mRefIndex. On the first conflict for a
'   reference, logs an error and adds an input error to the netting set,
'   which makes it INVALID with its EAD withheld whatever the row order.
'
' UPDATED
'   2026-10-07
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim key   As String    'Netting set | asset class | reference
    Dim k     As Long      'Position of the reference; 0 when new

'------------------------------------------------------------------------------
' COMPARE
'------------------------------------------------------------------------------
    'References contain no "|" (checked in ProcessTrades), so the key is
    'unambiguous.
        key = CStr(nsIdx) & "|" & CStr(ac) & "|" & rf
        k = KeyIndex(mRefIndex, key)
        If k = 0 Then
            mRefCount = mRefCount + 1
            If mRefCount > UBound(mRefClass) Then
                ReDim Preserve mRefClass(1 To mRefCount * 2)
                ReDim Preserve mRefConflict(1 To mRefCount * 2)
            End If
            mRefClass(mRefCount) = sfKey
            KeyAdd mRefIndex, key, mRefCount
        ElseIf mRefClass(k) <> sfKey And Not mRefConflict(k) Then
            mRefConflict(k) = True
            mNS(nsIdx).InputErrors = mNS(nsIdx).InputErrors + 1
            LogMsg SEV_ERROR, SH_TRADES, tid, "Reference '" & rf & "' is " & sfKey & " here but " & _
                   mRefClass(k) & " on another trade of the netting set - netting set INVALID, EAD withheld.", rowNum
        End If

End Sub


Private Sub AddToBucket( _
    ByVal nsIdx As Long, _
    ByVal ac As Long, _
    ByVal hsKey As String, _
    ByVal subKey As String, _
    ByVal sfEff As Double, _
    ByVal corr As Double, _
    ByVal enU As Double, _
    ByVal enM As Double, _
    ByVal tid As String, _
    ByVal rowNum As Long)
'
'==============================================================================
'                                 AddToBucket
'------------------------------------------------------------------------------
' PURPOSE
'   Add a trade's effective notionals to its bucket, creating the bucket on
'   its first trade.
'
' INPUTS
'   nsIdx, ac, hsKey, subKey: netting set, asset class, hedging set and
'      bucket of the trade; together they identify the bucket.
'   sfEff, corr: the trade's effective supervisory factor and correlation.
'   enU, enM: the trade's effective notional, unmargined and margined.
'   tid, rowNum: trade ID and sheet row, for messages.
'
' STATE OWNERSHIP
'   Adds to mBk, mBkCount and mBkIndex. Every trade of a bucket has the
'   same factor and correlation: the key holds the hedging set, which
'   includes the nature, and CheckReferenceClass makes a netting set whose
'   credit, equity or commodity reference has two sub-classes INVALID
'   (#37).
'
' UPDATED
'   2026-10-06
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim key   As String    'Bucket key: netting set # hedging set # bucket
    Dim b     As Long      'Position of the bucket in mBk

'------------------------------------------------------------------------------
' FIND OR CREATE THE BUCKET
'------------------------------------------------------------------------------
        key = CStr(nsIdx) & "#" & hsKey & "#" & subKey
        b = KeyIndex(mBkIndex, key)
        If b = 0 Then
            mBkCount = mBkCount + 1
            If mBkCount > UBound(mBk) Then
                ReDim Preserve mBk(1 To mBkCount * 2)
            End If
            b = mBkCount
            mBk(b).NSIdx = nsIdx
            mBk(b).AC = ac
            mBk(b).HSKey = hsKey
            mBk(b).SubKey = subKey
            mBk(b).SF = sfEff
            mBk(b).Corr = corr
            mBk(b).ENU = 0#
            mBk(b).ENM = 0#
            mBk(b).Trades = 0
            KeyAdd mBkIndex, key, b
        End If

'------------------------------------------------------------------------------
' ACCUMULATE
'------------------------------------------------------------------------------
        mBk(b).ENU = mBk(b).ENU + enU
        mBk(b).ENM = mBk(b).ENM + enM
        mBk(b).Trades = mBk(b).Trades + 1

End Sub


'
'------------------------------------------------------------------------------
'
'                     HEDGING-SET AND ASSET-CLASS ADD-ONS
'
'------------------------------------------------------------------------------
'

Private Sub ComputeHedgingSets()
'
'==============================================================================
'                              ComputeHedgingSets
'------------------------------------------------------------------------------
' PURPOSE
'   Collect the buckets into hedging sets, calculate each hedging-set
'   add-on on both bases, and sum them into the asset-class add-ons of each
'   netting set.
'
' STATE OWNERSHIP
'   Fills mHS, mHSCount and mHSIndex, links each bucket to its hedging set,
'   and adds to mAddOnU and mAddOnM.
'
' REFERENCE
'   IR CRE52.57; FX CRE52.59; credit CRE52.61; equity CRE52.66; commodity
'   CRE52.70; other risks CRR Art. 280f. The sum of absolute IR bucket
'   values is a Basel alternative (CRE52.57(5)); the CRR has only the
'   formula with offsets (Art. 280a(3)), so with IRBucketOffset FALSE a
'   CRR netting set with interest-rate trades is INVALID.
'
' UPDATED
'   2026-10-09
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim b           As Long       'Bucket, position in mBk
    Dim h           As Long       'Hedging set, position in mHS
    Dim key         As String     'Hedging-set key: netting set # hedging set
    Dim aU          As Double     'Entity or type add-on, unmargined
    Dim aM          As Double     'Entity or type add-on, margined
    Dim sfH         As Double     'Supervisory factor of an IR, FX or OT hedging set
    Dim refused()   As Boolean    'Per netting set: CRR set refused the IR alternative

'------------------------------------------------------------------------------
' COLLECT BUCKETS INTO HEDGING SETS
'------------------------------------------------------------------------------
    'IR buckets go into D1, D2 or D3. FX and OT add their effective
    'notionals. Credit, equity and commodity first turn each bucket into an
    'entity or type add-on (SF * EN), then add it to the systematic and
    'idiosyncratic sums of the single-factor model.
        For b = 1 To mBkCount
            key = CStr(mBk(b).NSIdx) & "#" & mBk(b).HSKey
            h = KeyIndex(mHSIndex, key)
            If h = 0 Then
                mHSCount = mHSCount + 1
                If mHSCount > UBound(mHS) Then
                    ReDim Preserve mHS(1 To mHSCount * 2)
                End If
                h = mHSCount
                mHS(h).NSIdx = mBk(b).NSIdx
                mHS(h).AC = mBk(b).AC
                mHS(h).Key = mBk(b).HSKey
                KeyAdd mHSIndex, key, h
            End If
            mBk(b).HSIdx = h
            mHS(h).Trades = mHS(h).Trades + mBk(b).Trades
            Select Case mBk(b).AC
                Case AC_IR
                    Select Case mBk(b).SubKey
                        Case "<1Y"
                            mHS(h).D1U = mHS(h).D1U + mBk(b).ENU
                            mHS(h).D1M = mHS(h).D1M + mBk(b).ENM
                        Case "1-5Y"
                            mHS(h).D2U = mHS(h).D2U + mBk(b).ENU
                            mHS(h).D2M = mHS(h).D2M + mBk(b).ENM
                        Case Else
                            mHS(h).D3U = mHS(h).D3U + mBk(b).ENU
                            mHS(h).D3M = mHS(h).D3M + mBk(b).ENM
                    End Select
                Case AC_FX, AC_OT
                    mHS(h).ENU = mHS(h).ENU + mBk(b).ENU
                    mHS(h).ENM = mHS(h).ENM + mBk(b).ENM
                Case Else
                    aU = mBk(b).SF * mBk(b).ENU
                    aM = mBk(b).SF * mBk(b).ENM
                    mHS(h).SysU = mHS(h).SysU + mBk(b).Corr * aU
                    mHS(h).IdioU = mHS(h).IdioU + (1# - mBk(b).Corr ^ 2) * aU * aU
                    mHS(h).SysM = mHS(h).SysM + mBk(b).Corr * aM
                    mHS(h).IdioM = mHS(h).IdioM + (1# - mBk(b).Corr ^ 2) * aM * aM
            End Select
        Next b

'------------------------------------------------------------------------------
' HEDGING-SET ADD-ONS
'------------------------------------------------------------------------------
    'IR: SF * effective notional across the buckets, with the bucket
    'formula or, when IRBucketOffset is False, the sum of absolute bucket
    'values. FX and OT: SF * |effective notional|. Credit, equity and
    'commodity: sqrt(systematic^2 + idiosyncratic). Each add-on is then
    'added to its netting set's asset-class total.
        If mNSCount > 0 Then
            ReDim refused(1 To mNSCount)
        End If
        For h = 1 To mHSCount
            Select Case mHS(h).AC
                Case AC_IR
                    sfH = BucketSF(h)
                    If Not pIRFull And mNS(mHS(h).NSIdx).IsCRR Then
                        If Not refused(mHS(h).NSIdx) Then
                            refused(mHS(h).NSIdx) = True
                            mNS(mHS(h).NSIdx).InputErrors = mNS(mHS(h).NSIdx).InputErrors + 1
                            LogMsg SEV_ERROR, SH_NS, mNS(mHS(h).NSIdx).ID, "IRBucketOffset is FALSE, but the " & _
                                   "CRR allows only the formula with offsets across maturity buckets " & _
                                   "(Art. 280a(3)) - netting set INVALID, EAD withheld."
                        End If
                    End If
                    If pIRFull Then
                        mHS(h).ENU = SACCR_IREffectiveNotional(mHS(h).D1U, mHS(h).D2U, mHS(h).D3U, pRho12, pRho23, pRho13)
                        mHS(h).ENM = SACCR_IREffectiveNotional(mHS(h).D1M, mHS(h).D2M, mHS(h).D3M, pRho12, pRho23, pRho13)
                    Else
                        mHS(h).ENU = Abs(mHS(h).D1U) + Abs(mHS(h).D2U) + Abs(mHS(h).D3U)
                        mHS(h).ENM = Abs(mHS(h).D1M) + Abs(mHS(h).D2M) + Abs(mHS(h).D3M)
                    End If
                    mHS(h).AddOnU = sfH * mHS(h).ENU
                    mHS(h).AddOnM = sfH * mHS(h).ENM
                Case AC_FX, AC_OT
                    sfH = BucketSF(h)
                    mHS(h).AddOnU = sfH * Abs(mHS(h).ENU)
                    mHS(h).AddOnM = sfH * Abs(mHS(h).ENM)
                Case Else
                    mHS(h).AddOnU = Sqr(mHS(h).SysU ^ 2 + mHS(h).IdioU)
                    mHS(h).AddOnM = Sqr(mHS(h).SysM ^ 2 + mHS(h).IdioM)
            End Select
            mAddOnU(mHS(h).NSIdx, mHS(h).AC) = mAddOnU(mHS(h).NSIdx, mHS(h).AC) + mHS(h).AddOnU
            mAddOnM(mHS(h).NSIdx, mHS(h).AC) = mAddOnM(mHS(h).NSIdx, mHS(h).AC) + mHS(h).AddOnM
        Next h

End Sub


Private Function IsFactorClass( _
    ByVal ac As Long) _
    As Boolean
'
'==============================================================================
'                                IsFactorClass
'------------------------------------------------------------------------------
' PURPOSE
'   Tell whether an asset class uses entity or type add-ons with
'   single-factor aggregation: credit, equity and commodity.
'
' UPDATED
'   2026-10-06
'==============================================================================
'

'------------------------------------------------------------------------------
' TEST
'------------------------------------------------------------------------------
        IsFactorClass = (ac = AC_CR Or ac = AC_EQ Or ac = AC_CO)

End Function


Private Function BucketSF( _
    ByVal h As Long) _
    As Double
'
'==============================================================================
'                                   BucketSF
'------------------------------------------------------------------------------
' PURPOSE
'   Return the supervisory factor of a hedging set from its first bucket.
'   Used for IR, FX and OT, where every bucket of a hedging set has the same
'   factor.
'
' INPUTS
'   h: hedging set, position in mHS.
'
' RETURNS
'   The factor of the first bucket linked to h; 0 when there is none.
'
' UPDATED
'   2026-10-06
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim b   As Long    'Bucket, position in mBk

'------------------------------------------------------------------------------
' SEARCH
'------------------------------------------------------------------------------
        For b = 1 To mBkCount
            If mBk(b).HSIdx = h Then
                BucketSF = mBk(b).SF
                Exit Function
            End If
        Next b

End Function


'
'------------------------------------------------------------------------------
'
'                             NETTING-SET RESULTS
'
'------------------------------------------------------------------------------
'

Private Sub ComputeNettingSets( _
    ByVal writeOutputs As Boolean)
'
'==============================================================================
'                              ComputeNettingSets
'------------------------------------------------------------------------------
' PURPOSE
'   Calculate RC, multiplier, PFE and EAD for every netting set whose
'   trades are all valid, apply the margined-EAD cap, and write the Results
'   sheet with one row per input netting set, a status and a total row.
'   A netting set with a rejected trade is INCOMPLETE: its EAD is withheld,
'   because a figure on the remaining trades would understate the exposure
'   (#36). One without trades is NO TRADES.
'
' INPUTS
'   writeOutputs: True writes the Results sheet.
'
' STATE OWNERSHIP
'   Adds each VALID EAD to mTotalEAD and counts INCOMPLETE sets in
'   mIncomplete. Only VALID sets enter the TOTAL row.
'
' REFERENCE
'   EAD CRE52.1; cap CRE52.2 and CRR Art. 274(3); RC CRE52.10, CRE52.18 and
'   CRR Art. 275(1); multiplier CRE52.23; collateral of the cap, NICA, per
'   CRE52.10 and EBA Q&A 2023_6962.
'
' UPDATED
'   2026-10-10
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    'Results columns (RS_NCOLS = 28): 1 ID, 2 counterparty, 3 regime,
    '4 margined, 5 trades, 6 MPOR, 7 V, 8 C, 9 RC, 10 to 15 add-ons IR, FX,
    'CR, EQ, CO, OT, 16 aggregate add-on, 17 multiplier, 18 PFE, 19 alpha,
    '20 EAD margined, 21 EAD cap, 22 EAD, 23 cap applied, 24 C (cap basis),
    '25 RC (cap basis), 26 add-on unmargined, 27 multiplier (cap basis),
    '28 status.
    Dim ws       As Worksheet    'Results sheet
    Dim k        As Long         'Netting set, position in mNS
    Dim a        As Long         'Asset class, or Results column in loops
    Dim outArr() As Variant      'Results rows, (row, column)
    Dim nOut     As Long         'Results rows filled
    Dim C        As Double       'Collateral C = VM + NICA
    Dim cCap     As Double       'Collateral used for the unmargined basis
    Dim addU     As Double       'Aggregate add-on, unmargined
    Dim addM     As Double       'Aggregate add-on, margined
    Dim rcU      As Double       'RC, unmargined basis
    Dim rcM      As Double       'RC, margined
    Dim mU       As Double       'Multiplier, unmargined basis
    Dim mM       As Double       'Multiplier, margined
    Dim pfeU     As Double       'PFE, unmargined basis
    Dim pfeM     As Double       'PFE, margined
    Dim eadU     As Double       'EAD, unmargined basis: the EAD or the cap
    Dim eadM     As Double       'EAD, margined, before the cap
    Dim ead      As Double       'EAD reported
    Dim tot(1 To RS_NCOLS) As Double    'Column totals for the TOTAL row

'------------------------------------------------------------------------------
' CALCULATE EACH NETTING SET
'------------------------------------------------------------------------------
        ReDim outArr(1 To mNSCount + 1, 1 To RS_NCOLS)
        nOut = 0
        For k = 1 To mNSCount
            If mNS(k).InputErrors > 0 Then
                mIncomplete = mIncomplete + 1
                LogMsg SEV_ERROR, SH_NS, mNS(k).ID, "EAD withheld: " & mNS(k).InputErrors & _
                       " invalid netting-set field(s)."
                nOut = nOut + 1
                WriteStatusRow outArr, nOut, k, "INVALID: " & mNS(k).InputErrors & " input error(s)"
                GoTo NextNS
            End If
            If mNS(k).Rejected > 0 Then
                mIncomplete = mIncomplete + 1
                LogMsg SEV_ERROR, SH_NS, mNS(k).ID, "EAD withheld: " & mNS(k).Rejected & _
                       " trade(s) rejected - see the Trades messages."
                nOut = nOut + 1
                WriteStatusRow outArr, nOut, k, "INCOMPLETE: " & mNS(k).Rejected & " of " & _
                               (mNS(k).Trades + mNS(k).Rejected) & " trade(s) rejected"
                GoTo NextNS
            End If
            If mNS(k).Trades = 0 Then
                LogMsg SEV_INFO, SH_NS, mNS(k).ID, "Netting set has no trades - no EAD."
                nOut = nOut + 1
                WriteStatusRow outArr, nOut, k, "NO TRADES"
                GoTo NextNS
            End If
            With mNS(k)

    'Collateral and aggregate add-ons on both bases.
                C = .VM + .NICA
                addU = 0#
                addM = 0#
                For a = 1 To AC_COUNT
                    addU = addU + mAddOnU(k, a)
                    addM = addM + mAddOnM(k, a)
                Next a

    'Unmargined basis: the EAD of an unmargined netting set, or the cap of
    'a margined one. The cap of a margined set uses NICA only, the
    'collateral of the same set without a margin agreement [CRE52.2,
    'CRE52.10; CRR Art. 274(3), 275(1); EBA Q&A 2023_6962]. An unmargined
    'set uses C: NICA, plus VM posted under a one-way agreement under BCBS.
                If .Margined Then
                    cCap = .NICA
                Else
                    cCap = C
                End If
                rcU = SACCR_ReplacementCost(.V, cCap, False)
                mU = SACCR_Multiplier(.V - cCap, addU, pFloor)
                pfeU = mU * addU
                eadU = .Alpha * (rcU + pfeU)

    'Margined basis, then the cap: the EAD is the lower of the two.
                If .Margined Then
                    rcM = SACCR_ReplacementCost(.V, C, True, .TH, .MTA, .NICA)
                    mM = SACCR_Multiplier(.V - C, addM, pFloor)
                    pfeM = mM * addM
                    eadM = .Alpha * (rcM + pfeM)
                    ead = Min2(eadM, eadU)
                Else
                    ead = eadU
                End If
                mTotalEAD = mTotalEAD + ead

    'Results row. A margined netting set shows its margined figures in
    'columns 9 to 18 and the cap figures in 21 and 24 to 27.
                nOut = nOut + 1
                outArr(nOut, 1) = .ID
                outArr(nOut, 2) = .Counterparty
                outArr(nOut, 3) = .Regime
                outArr(nOut, 4) = IIf(.Margined, "Y", "N")
                outArr(nOut, 5) = .Trades
                If .Margined Then
                    outArr(nOut, 6) = .MPOR
                End If
                outArr(nOut, 7) = .V
                outArr(nOut, 8) = C
                For a = 1 To AC_COUNT
                    If .Margined Then
                        outArr(nOut, 9 + a) = mAddOnM(k, a)
                    Else
                        outArr(nOut, 9 + a) = mAddOnU(k, a)
                    End If
                Next a
                If .Margined Then
                    outArr(nOut, 9) = rcM
                    outArr(nOut, 16) = addM
                    outArr(nOut, 17) = mM
                    outArr(nOut, 18) = pfeM
                    outArr(nOut, 20) = eadM
                    outArr(nOut, 21) = eadU
                    outArr(nOut, 23) = IIf(eadU < eadM, "Y", "N")
                    outArr(nOut, 24) = cCap
                    outArr(nOut, 25) = rcU
                    outArr(nOut, 26) = addU
                    outArr(nOut, 27) = mU
                Else
                    outArr(nOut, 9) = rcU
                    outArr(nOut, 16) = addU
                    outArr(nOut, 17) = mU
                    outArr(nOut, 18) = pfeU
                    outArr(nOut, 23) = "n/a"
                End If
                outArr(nOut, 19) = .Alpha
                outArr(nOut, 22) = ead
                outArr(nOut, RS_STATUS_COL) = "VALID"

    'Running totals for V, C, RC, the add-ons, PFE and EAD.
                tot(7) = tot(7) + .V
                tot(8) = tot(8) + C
                For a = 9 To 16
                    tot(a) = tot(a) + ToDbl(outArr(nOut, a))
                Next a
                tot(18) = tot(18) + ToDbl(outArr(nOut, 18))
                tot(22) = tot(22) + ead
            End With
NextNS:
        Next k

'------------------------------------------------------------------------------
' WRITE RESULTS
'------------------------------------------------------------------------------
    'The TOTAL row sums columns 7 to 18 except the multiplier (17), and the
    'EAD (22). It is written in bold.
        If writeOutputs Then
            Set ws = GetSheet(SH_RESULTS)
            ClearOutputBlock ws, FIRST_DATA_ROW, RS_NCOLS
            If nOut > 0 Then
                nOut = nOut + 1
                outArr(nOut, 1) = "TOTAL"
                For a = 7 To 18
                    If a <> 17 Then
                        outArr(nOut, a) = tot(a)
                    End If
                Next a
                outArr(nOut, 22) = tot(22)
                WriteBlock ws, FIRST_DATA_ROW, outArr, nOut, RS_NCOLS
                FormatColumns ws, FIRST_DATA_ROW, nOut, _
                    "@|@|@|@|0|0|#,##0|#,##0|#,##0|#,##0|#,##0|#,##0|#,##0|#,##0|#,##0|#,##0|0.0000|#,##0|0.00|#,##0|#,##0|#,##0|@|#,##0|#,##0|#,##0|0.0000|@"
                ws.Range(ws.Cells(FIRST_DATA_ROW + nOut - 1, 1), ws.Cells(FIRST_DATA_ROW + nOut - 1, RS_NCOLS)).Font.Bold = True
            End If
        End If

End Sub


'
'------------------------------------------------------------------------------
'
'                                OUTPUT WRITERS
'
'------------------------------------------------------------------------------
'

Private Sub WriteStatusRow( _
    ByRef outArr() As Variant, _
    ByVal r As Long, _
    ByVal k As Long, _
    ByVal status As String)
'
'==============================================================================
'                                WriteStatusRow
'------------------------------------------------------------------------------
' PURPOSE
'   Fill a Results row for a netting set that has no EAD: its identity, its
'   valid-trade count and its status. The exposure columns stay blank.
'
' INPUTS
'   outArr: the Results rows; r: the row to fill; k: the netting set,
'   position in mNS; status: the text for the status column.
'
' UPDATED
'   2026-10-06
'==============================================================================
'
        outArr(r, 1) = mNS(k).ID
        outArr(r, 2) = mNS(k).Counterparty
        outArr(r, 3) = mNS(k).Regime
        outArr(r, 4) = IIf(mNS(k).Margined, "Y", "N")
        outArr(r, 5) = mNS(k).Trades
        outArr(r, RS_STATUS_COL) = status

End Sub


Private Sub ClearOutputSheets()
'
'==============================================================================
'                              ClearOutputSheets
'------------------------------------------------------------------------------
' PURPOSE
'   Clear the TradeCalc, Buckets, HedgingSets and Results tables before a
'   run. Checks is rewritten by every run.
'
' UPDATED
'   2026-10-06
'==============================================================================
'
        ClearOutputBlock GetSheet(SH_TRADECALC), FIRST_DATA_ROW, TC_NCOLS
        ClearOutputBlock GetSheet(SH_BUCKETS), FIRST_DATA_ROW, BK_NCOLS
        ClearOutputBlock GetSheet(SH_HEDGING), FIRST_DATA_ROW, HS_NCOLS
        ClearOutputBlock GetSheet(SH_RESULTS), FIRST_DATA_ROW, RS_NCOLS

End Sub


'
'------------------------------------------------------------------------------
'
'                               RUN FINGERPRINT
'
'------------------------------------------------------------------------------
'

Public Function InputFingerprint() As String
'
'==============================================================================
'                               InputFingerprint
'------------------------------------------------------------------------------
' PURPOSE
'   Fingerprint everything a run reads: the NettingSets and Trades input
'   rows (without the unread Comment column) and Params columns A to H,
'   which hold the parameters and the factor and FX tables (#36).
'
' RETURNS
'   An 8-digit hexadecimal fingerprint (CORE_Util.TextHash). Equal
'   fingerprints mean the inputs are, all but certainly, unchanged.
'
' UPDATED
'   2026-10-07
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PARAMS_COLS As Long = 8    'Params columns A to H
    Dim parts(1 To 3)   As String    'Text of each input block

'------------------------------------------------------------------------------
' FINGERPRINT
'------------------------------------------------------------------------------
        parts(1) = BlockText(GetSheet(SH_NS), FIRST_DATA_ROW, NS_NCOLS)
        parts(2) = BlockText(GetSheet(SH_TRADES), FIRST_DATA_ROW, TR_COMMENT - 1)
        parts(3) = BlockText(GetSheet(SH_PARAMS), 1, PARAMS_COLS)
        InputFingerprint = TextHash(Join(parts, Chr$(29)))

End Function


Private Function BlockText( _
    ByVal ws As Worksheet, _
    ByVal firstRow As Long, _
    ByVal nCols As Long) _
    As String
'
'==============================================================================
'                                  BlockText
'------------------------------------------------------------------------------
' PURPOSE
'   Write a block of cells as one text, cell by cell, from firstRow to the
'   last row with a value. Blank rows below it are left out, so that a
'   used range that grew without new values does not change the text.
'
' UPDATED
'   2026-10-07
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim lastR    As Long        'Last used row
    Dim data     As Variant     'Cell values, Value2 (dates as serials)
    Dim nRows    As Long        'Rows up to the last with a value
    Dim pieces() As String      'One text per cell
    Dim r        As Long        'Row of data
    Dim c        As Long        'Column of data
    Dim k        As Long        'Position in pieces

'------------------------------------------------------------------------------
' READ
'------------------------------------------------------------------------------
        lastR = UsedLastRow(ws)
        If lastR < firstRow Then
            Exit Function
        End If
        data = ws.Range(ws.Cells(firstRow, 1), ws.Cells(lastR, nCols)).Value2
        For r = UBound(data, 1) To 1 Step -1
            If RowHasData(data, r, nCols, 0) Then
                nRows = r
                Exit For
            End If
        Next r
        If nRows = 0 Then
            Exit Function
        End If

'------------------------------------------------------------------------------
' WRITE AS TEXT
'------------------------------------------------------------------------------
    'Chr$(31) separates cells; the column count fixes where rows end.
        ReDim pieces(1 To nRows * nCols)
        For r = 1 To nRows
            For c = 1 To nCols
                k = k + 1
                If IsError(data(r, c)) Then
                    pieces(k) = "#ERR"
                Else
                    pieces(k) = CStr(data(r, c))
                End If
            Next c
        Next r
        BlockText = Join(pieces, Chr$(31))

End Function


Public Function LastRunInputs() As String
'
'==============================================================================
'                                LastRunInputs
'------------------------------------------------------------------------------
' PURPOSE
'   Return the input fingerprint stored by the last completed run, or ""
'   when no run's results are on the sheets (the name is empty or absent).
'
' ERROR POLICY
'   Contains the expected error of a missing name.
'
' UPDATED
'   2026-10-07
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim wbName   As Name    'The stored name, Nothing when absent

'------------------------------------------------------------------------------
' READ
'------------------------------------------------------------------------------
    'The name refers to a text constant: ="5A428560".
        On Error Resume Next
        Set wbName = ThisWorkbook.Names(RUN_INPUTS_NAME)
        Err.Clear
        On Error GoTo 0
        If Not wbName Is Nothing Then
            LastRunInputs = Replace(Replace(wbName.RefersTo, "=", ""), """", "")
        End If

End Function


Private Sub RememberRunInputs( _
    ByVal fingerprint As String)
'
'==============================================================================
'                              RememberRunInputs
'------------------------------------------------------------------------------
' PURPOSE
'   Store the input fingerprint of a completed run in a hidden workbook
'   name.
'
' ERROR POLICY
'   Best effort: a workbook whose names cannot be changed (protected
'   structure) still runs; its results are then reported as unknown.
'
' UPDATED
'   2026-10-07
'==============================================================================
'
        On Error GoTo Failed
        ThisWorkbook.Names.Add Name:=RUN_INPUTS_NAME, RefersTo:="=""" & fingerprint & """", Visible:=False
        Exit Sub

Failed:
        ForgetRunInputs

End Sub


Public Sub ForgetRunInputs()
'
'==============================================================================
'                               ForgetRunInputs
'------------------------------------------------------------------------------
' PURPOSE
'   Empty the stored fingerprint, because no run's results are on the
'   sheets: after a failed or validation-only run, and after Clear outputs.
'   The name is kept, empty, because the status formula on Results refers
'   to it.
'
' ERROR POLICY
'   Contains any error.
'
' UPDATED
'   2026-10-07
'==============================================================================
'
        On Error GoTo Failed
        ThisWorkbook.Names.Add Name:=RUN_INPUTS_NAME, RefersTo:="=""""", Visible:=False

Failed:

End Sub


Private Function WithdrawOutputs(ByRef failureDetails As String) As Boolean
'
'==============================================================================
'                               WithdrawOutputs
'------------------------------------------------------------------------------
' PURPOSE
'   Attempt every output independently after a run fails (#36, review #85).
' RETURNS
'   True only if all four clears succeeded. failureDetails is reset on entry
'   and contains every failed sheet, error number, source and description.
' ERROR POLICY
'   Each helper contains its own error; a failed clear never skips another.
'   The caller keeps the original run error and appends cleanup diagnostics.
' UPDATED
'   2026-10-08
'==============================================================================
'
        failureDetails = ""
        WithdrawOutputs = True
        If Not TryClearOutput(SH_TRADECALC, TC_NCOLS, failureDetails) Then WithdrawOutputs = False
        If Not TryClearOutput(SH_BUCKETS, BK_NCOLS, failureDetails) Then WithdrawOutputs = False
        If Not TryClearOutput(SH_HEDGING, HS_NCOLS, failureDetails) Then WithdrawOutputs = False
        If Not TryClearOutput(SH_RESULTS, RS_NCOLS, failureDetails) Then WithdrawOutputs = False
End Function


Private Function TryClearOutput( _
    ByVal sheetName As String, _
    ByVal nCols As Long, _
    ByRef failureDetails As String) As Boolean
'
'==============================================================================
'                                TryClearOutput
'------------------------------------------------------------------------------
' PURPOSE
'   Clear one output and capture its failure before any other operation.
' INPUTS
'   sheetName, nCols: output sheet and table width, starting at column A.
' RETURNS
'   True only when this output was cleared. Appends any failure to
'   failureDetails, preserving details from earlier sheets.
' ERROR POLICY
'   Contains this cleanup failure; the original run error remains primary.
' UPDATED
'   2026-10-08
'==============================================================================
    Dim errNumber As Long         'Cleanup error, captured before formatting
    Dim errSource As String       'Cleanup error source
    Dim errDescription As String  'Cleanup error description
        On Error GoTo Failed
        If gEngineFault = "outputs+cleanup" Then
            If sheetName = SH_TRADECALC Or sheetName = SH_RESULTS Then
                Err.Raise ERR_CLEANUP_FAILED, "CORE_Engine.TryClearOutput", "Injected output cleanup failure."
            End If
        End If
        ClearOutputBlock GetSheet(sheetName), FIRST_DATA_ROW, nCols
        TryClearOutput = True
        Exit Function
Failed:
        errNumber = Err.Number
        errSource = Err.Source
        errDescription = Err.Description
        If Len(failureDetails) > 0 Then failureDetails = failureDetails & "; "
        failureDetails = failureDetails & sheetName & " [error " & CStr(errNumber) & _
                         "; source=" & errSource & "]: " & errDescription
End Function


Private Sub WriteBuckets()
'
'==============================================================================
'                                 WriteBuckets
'------------------------------------------------------------------------------
' PURPOSE
'   Write one Buckets row per bucket.
'
' STATE OWNERSHIP
'   Clears and rewrites the Buckets sheet.
'
' UPDATED
'   2026-10-06
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ws      As Worksheet    'Buckets sheet
    Dim arr()   As Variant      'Output rows, (row, column)
    Dim b       As Long         'Bucket, position in mBk

'------------------------------------------------------------------------------
' BUILD ROWS
'------------------------------------------------------------------------------
    'Columns: 1 netting set, 2 asset class, 3 hedging set, 4 bucket,
    '5 trades, 6 SF, 7 correlation, 8 and 9 effective notional unmargined
    'and margined, 10 and 11 entity or type add-on unmargined and margined.
    'Correlation and add-ons apply to credit, equity and commodity only;
    'margined figures to margined netting sets only.
        Set ws = GetSheet(SH_BUCKETS)
        ClearOutputBlock ws, FIRST_DATA_ROW, BK_NCOLS
        If mBkCount = 0 Then
            Exit Sub
        End If
        ReDim arr(1 To mBkCount, 1 To BK_NCOLS)
        For b = 1 To mBkCount
            arr(b, 1) = mNS(mBk(b).NSIdx).ID
            arr(b, 2) = ACCode(mBk(b).AC)
            arr(b, 3) = mBk(b).HSKey
            arr(b, 4) = mBk(b).SubKey
            arr(b, 5) = mBk(b).Trades
            arr(b, 6) = mBk(b).SF
            If IsFactorClass(mBk(b).AC) Then arr(b, 7) = mBk(b).Corr
            arr(b, 8) = mBk(b).ENU
            arr(b, 9) = IIf(mNS(mBk(b).NSIdx).Margined, mBk(b).ENM, Empty)
            If IsFactorClass(mBk(b).AC) Then
                arr(b, 10) = mBk(b).SF * mBk(b).ENU
                If mNS(mBk(b).NSIdx).Margined Then arr(b, 11) = mBk(b).SF * mBk(b).ENM
            End If
        Next b

'------------------------------------------------------------------------------
' WRITE
'------------------------------------------------------------------------------
        WriteBlock ws, FIRST_DATA_ROW, arr, mBkCount, BK_NCOLS
        FormatColumns ws, FIRST_DATA_ROW, mBkCount, "@|@|@|@|0|0.00%|0%|#,##0|#,##0|#,##0|#,##0"

End Sub


Private Sub WriteHedgingSets()
'
'==============================================================================
'                               WriteHedgingSets
'------------------------------------------------------------------------------
' PURPOSE
'   Write one HedgingSets row per hedging set, on the basis that applies to
'   its netting set: margined figures for a margined set, unmargined
'   otherwise.
'
' STATE OWNERSHIP
'   Clears and rewrites the HedgingSets sheet.
'
' UPDATED
'   2026-10-06
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ws      As Worksheet    'HedgingSets sheet
    Dim arr()   As Variant      'Output rows, (row, column)
    Dim h       As Long         'Hedging set, position in mHS
    Dim mg      As Boolean      'Netting set of h is margined

'------------------------------------------------------------------------------
' BUILD ROWS
'------------------------------------------------------------------------------
    'Columns: 1 netting set, 2 asset class, 3 hedging set, 4 trades, 5 to 7
    'IR buckets D1 to D3, 8 effective notional (IR, FX, OT), 9 systematic
    'and 10 idiosyncratic part (credit, equity, commodity), 11 hedging-set
    'add-on, 12 add-on unmargined, 13 basis, 14 SF (IR, FX, OT).
        Set ws = GetSheet(SH_HEDGING)
        ClearOutputBlock ws, FIRST_DATA_ROW, HS_NCOLS
        If mHSCount = 0 Then
            Exit Sub
        End If
        ReDim arr(1 To mHSCount, 1 To HS_NCOLS)
        For h = 1 To mHSCount
            mg = mNS(mHS(h).NSIdx).Margined
            arr(h, 1) = mNS(mHS(h).NSIdx).ID
            arr(h, 2) = ACCode(mHS(h).AC)
            arr(h, 3) = mHS(h).Key
            arr(h, 4) = mHS(h).Trades
            If mHS(h).AC = AC_IR Then
                If mg Then
                    arr(h, 5) = mHS(h).D1M
                    arr(h, 6) = mHS(h).D2M
                    arr(h, 7) = mHS(h).D3M
                Else
                    arr(h, 5) = mHS(h).D1U
                    arr(h, 6) = mHS(h).D2U
                    arr(h, 7) = mHS(h).D3U
                End If
            End If
            If Not IsFactorClass(mHS(h).AC) Then
                If mg Then
                    arr(h, 8) = mHS(h).ENM
                Else
                    arr(h, 8) = mHS(h).ENU
                End If
            Else
                If mg Then
                    arr(h, 9) = mHS(h).SysM
                    arr(h, 10) = Sqr(mHS(h).IdioM)
                Else
                    arr(h, 9) = mHS(h).SysU
                    arr(h, 10) = Sqr(mHS(h).IdioU)
                End If
            End If
            If mg Then
                arr(h, 11) = mHS(h).AddOnM
            Else
                arr(h, 11) = mHS(h).AddOnU
            End If
            arr(h, 12) = mHS(h).AddOnU
            arr(h, 13) = IIf(mg, "Margined", "Unmargined")
            If Not IsFactorClass(mHS(h).AC) Then
                arr(h, 14) = BucketSF(h)
            End If
        Next h

'------------------------------------------------------------------------------
' WRITE
'------------------------------------------------------------------------------
        WriteBlock ws, FIRST_DATA_ROW, arr, mHSCount, HS_NCOLS
        FormatColumns ws, FIRST_DATA_ROW, mHSCount, "@|@|@|0|#,##0|#,##0|#,##0|#,##0|#,##0|#,##0|#,##0|#,##0|@|0.00%"

End Sub


Private Function WriteChecks( _
    ByRef failure As String) _
    As Boolean
'
'==============================================================================
'                                 WriteChecks
'------------------------------------------------------------------------------
' PURPOSE
'   Write the logged messages to the Checks sheet, errors in bold, or a
'   single "No issues found." line.
'
' RETURNS
'   True when the sheet was written. False when writing failed, with the
'   reason in failure; Calculate then withdraws the results (#36).
'
' STATE OWNERSHIP
'   Clears and rewrites the Checks sheet.
'
' ERROR POLICY
'   Contains any error and reports it through the result, never silently.
'
' UPDATED
'   2026-10-07
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ws      As Worksheet    'Checks sheet
    Dim arr()   As Variant      'Output rows, (message, field)
    Dim i       As Long         'Message
    Dim j       As Long         'Field

'------------------------------------------------------------------------------
' WRITE
'------------------------------------------------------------------------------
    'mLog is stored (field, message); it is transposed into sheet rows.
        On Error GoTo Failed
        failure = ""
        Set ws = GetSheet(SH_CHECKS)
        ClearOutputBlock ws, FIRST_DATA_ROW, CK_NCOLS
        If mLogCount = 0 Then
            ReDim arr(1 To 1, 1 To CK_NCOLS)
            arr(1, 1) = SEV_INFO
            arr(1, 5) = "No issues found."
            WriteBlock ws, FIRST_DATA_ROW, arr, 1, CK_NCOLS
            WriteChecks = True
            Exit Function
        End If
        ReDim arr(1 To mLogCount, 1 To CK_NCOLS)
        For i = 1 To mLogCount
            For j = 1 To CK_NCOLS
                arr(i, j) = mLog(j, i)
            Next j
        Next i
        WriteBlock ws, FIRST_DATA_ROW, arr, mLogCount, CK_NCOLS
        For i = 1 To mLogCount
            If arr(i, 1) = SEV_ERROR Then
                ws.Cells(FIRST_DATA_ROW + i - 1, 1).Font.Bold = True
            End If
        Next i
        WriteChecks = True
        Exit Function

'------------------------------------------------------------------------------
' REPORT FAILURE
'------------------------------------------------------------------------------
Failed:
        failure = "error " & Err.Number & ": " & Err.Description

End Function


Private Sub WriteRunInfo( _
    ByVal secs As Double, _
    ByVal completed As Boolean, _
    ByVal wroteOutputs As Boolean, _
    ByVal failure As String)
'
'==============================================================================
'                                 WriteRunInfo
'------------------------------------------------------------------------------
' PURPOSE
'   Write a one-line run summary to cell A2 of the Results sheet.
'
' INPUTS
'   secs: run duration in seconds.
'   completed: the result of Calculate.
'   wroteOutputs: False for a validation-only run, whose outputs were
'   cleared rather than written.
'   failure: why the run failed after it started writing, for example the
'   Checks sheet could not be written; "" otherwise.
'
' ERROR POLICY
'   Best effort: errors are ignored, because the summary is informative only.
'
' UPDATED
'   2026-10-06
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ws    As Worksheet    'Results sheet
    Dim txt   As String       'Summary line

'------------------------------------------------------------------------------
' WRITE
'------------------------------------------------------------------------------
        On Error Resume Next
        Set ws = GetSheet(SH_RESULTS)
        If Len(failure) > 0 Then
            txt = "Last run " & Format$(Now, "yyyy-mm-dd hh:mm:ss") & " FAILED - " & failure
        ElseIf Not wroteOutputs Then
            txt = "Last validation " & Format$(Now, "yyyy-mm-dd hh:mm:ss") & _
                  " | errors " & mErrCount & ", warnings " & mWarnCount & _
                  " | outputs cleared: press Run SA-CCR to calculate"
        ElseIf completed Then
            txt = "Last run " & Format$(Now, "yyyy-mm-dd hh:mm:ss") & _
                  " | reporting date " & Format$(CDate(pAsOf), "yyyy-mm-dd") & _
                  " | ccy " & pRepCcy & _
                  " | trades used " & mTradesUsed & " of " & mTradesRead & _
                  " | errors " & mErrCount & ", warnings " & mWarnCount & _
                  " | inputs " & mRunInputs & _
                  " | " & Format$(secs, "0.00") & " s"
            If mIncomplete > 0 Then
                txt = txt & " | EAD withheld for " & mIncomplete & " netting set(s)"
            End If
        Else
            txt = "Last run " & Format$(Now, "yyyy-mm-dd hh:mm:ss") & _
                  " FAILED - see Checks sheet (" & mErrCount & " errors)"
        End If
        ws.Range(RUNINFO_CELL).Value = txt

End Sub
