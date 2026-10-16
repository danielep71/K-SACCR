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
'   Both: the cap uses C = NICA only, the collateral of the same netting
'         set without a margin agreement.  [CRE52.2, CRE52.10; Art. 274(3),
'         275(1); EBA-QA-2023_6962]
'   BCBS: VM posted under a one-way agreement is part of C of an unmargined
'         netting set.                       [CRE52.10 footnote 2]
'   CRR:  such a netting set is margined and is rejected as unmargined.
'                                            [Art. 272(7a)]
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
'     WriteBuckets, WriteHedgingSets, WriteChecks, WriteRunInfo
'     (RunInfoText): outputs.
'   Each trade, bucket and hedging set is stored once in an array, with a
'   Collection (see CORE_Util.KeyIndex) mapping its key to its array position.
'
' PUBLIC SURFACE
'   None outside this VBA project. Calculate, the run counters (with the
'   per-regime totals TotalEAD, RegimeSetCount, RegimeValidCount,
'   RegimeWithheldCount and TotalStatusText) and the run-fingerprint
'   procedures (InputFingerprint, LastRunInputs, ForgetRunInputs, and
'   BlockText, which M_Main also uses for the reconciliation fingerprint)
'   and ParameterCodes, which it uses for the evidence report,
'   are Public for M_Main; Option Private Module keeps them off the
'   supported external surface.
'
' DEPENDENCIES
'   CORE_Config for the layout, CORE_Util for conversions and sheet access,
'   SACCR_Formulas for every regulatory formula, and CORE_Export.NumberText
'   for the run time in the run summary, with a decimal point in every
'   locale (#45).
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
'   2026-10-15
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
        Trades            As Long       'Number of valid trade legs
        Rejected          As Long       'Number of trade legs rejected by validation
        EconTrades        As Long       'Number of economic trades, each counted once (#38)
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

    'One trade while ProcessTrades works through it (#60): its inputs as
    'read, its validation state and its trade-level SA-CCR quantities. It is
    'reset for every trade. Amounts are in the reporting currency unless
    'stated; dates are serials, -1 when missing.
    Private Type tTrade
        ok            As Boolean    'False once any error is found
        msg           As String     'Error reasons, separated by "; "
        warn          As String     'Warnings, separated by "; "
        tid           As String     'Trade ID
        nsId          As String     'Netting-set ID, upper case
        acTxt         As String     'Asset-class code, upper case
        subCls        As String     'Sub-class, upper case
        rf            As String     'Risk factor or reference, upper case; FX: the pair in alphabetical order
        instType      As String     'LINEAR, OPTION or CDO
        isLong        As Boolean    'Long (bought) position
        optTxt        As String     'CALL or PUT
        nature        As String     'STANDARD, BASIS or VOLATILITY
        lbl           As String     'Basis or volatility hedging-set label
        nsIdx         As Long       'Netting set, position in mNS; 0 when unknown
        ac            As Long       'Asset class, an AC_ constant; 0 when unknown
        sfKey         As String     'Key into the supervisory-factor table
        sfIdx         As Long       'Position in mSF; 0 when not found
        natFactor     As Double     'Factor multiplier: 1, basis or volatility
        natTag        As String     'Hedging-set suffix: STD, BASIS:label or VOL:label
        notional      As Double     'Notional, in its currency
        nccy          As String     'Notional currency
        fxN           As Double     'Rate of the notional currency
        notionalRep   As Double     'Notional in the reporting currency
        mtm           As Double     'MtM, in its currency
        mccy          As String     'MtM currency
        fxM           As Double     'Rate of the MtM currency
        mtmRep        As Double     'MtM in the reporting currency
        dStart        As Double     'Start date
        dEnd          As Double     'End date
        dMat          As Double     'Maturity date
        dExp          As Double     'Option expiry date
        S             As Double     'Start S, years, at least 0
        E             As Double     'End E, years
        M             As Double     'Maturity M, years
        T             As Double     'Option expiry T, years
        sd            As Double     'Supervisory duration (IR and credit)
        adjN          As Double     'Adjusted notional d
        delta         As Double     'Supervisory delta
        mfU           As Double     'Maturity factor, unmargined
        mfM           As Double     'Maturity factor, margined
        enU           As Double     'Effective notional delta * d * MF, unmargined
        enM           As Double     'Effective notional delta * d * MF, margined
        hsKey         As String     'Hedging-set key
        subKey        As String     'Bucket key within the hedging set
        bucketNo      As Long       'IR maturity bucket 1, 2 or 3
        sfEff         As Double     'Supervisory factor times natFactor
        corrEff       As Double     'Correlation
        pairSign      As Double     'FX: +1 pair kept, -1 inverted, 0 invalid; 1 otherwise
        lam           As Double     'Lambda shift applied
        lamOut        As Variant    'Lambda reported; Empty for non-options
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
' MODULE CONSTANTS
'------------------------------------------------------------------------------
    'Opens the text of every input fingerprint; changed whenever what the
    'fingerprint covers changes, so an older stored one never matches.
        Private Const FP_FORMAT   As String = "FP2"    'Typed cells, headers, date system (#197)

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
        Private mRegEAD(1 To 2)       As Double  'Sum of VALID EADs per regime: 1 CRR, 2 BCBS (#198)
        Private mRegSets(1 To 2)      As Long    'Netting sets per regime
        Private mRegValid(1 To 2)     As Long    'VALID netting sets per regime
        Private mRegWithheld(1 To 2)  As Long    'INVALID or INCOMPLETE netting sets per regime
        Private mIncomplete   As Long            'Netting sets whose EAD was withheld
        Private mParamError   As Boolean         'A parameter was present but not usable
        Private mRunInputs    As String          'Input fingerprint of this run; "" unless it completed
        Private mRunClass     As String          'RUN_REGULATORY or RUN_WHATIF; "" until classified
        Private mEconRead     As Long            'Economic trades read: groups of legs (#38)
        Private mEconUsed     As Long            'Economic trades whose legs were all used
        Private mRowGroup()   As Long            'Economic trade of each Trades row; 0 for a row without ID
        Private mGrpError()   As String          'Why each economic trade is rejected; "" when it is not
        Private mGrpRejected() As Long           'Rejected legs of each economic trade
        Private mDeviations   As Long            'Parameters that differ from the regulatory set

'------------------------------------------------------------------------------
' TEST SEAM
'------------------------------------------------------------------------------
    'Set only by TEST_MainState: "outputs" raises ERR_INJECTED_FAULT after
    'the first output sheets are written, to prove that a failed run
    'withdraws them (#36). "outputs+cleanup" also fails the TradeCalc and
    'Results clears with a different error. "runinfo" fails the composing
    'of the run summary, which must not fail the run (#60). Empty in normal
    'use.
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
        ClassifyRun

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


Public Property Get TotalEAD( _
    ByVal regimeCode As String) _
    As Double
'
'==============================================================================
'                                   TotalEAD
'------------------------------------------------------------------------------
' PURPOSE
'   Sum of the EADs of the VALID netting sets of one regime in the last
'   run, in the reporting currency. EADs of different regimes are never
'   added together (#198).
'
' INPUTS
'   regimeCode: RG_CRR or RG_BCBS.
'
' UPDATED
'   2026-10-13
'==============================================================================
'
        TotalEAD = mRegEAD(RegimeIndex(regimeCode))

End Property


Public Property Get RegimeSetCount( _
    ByVal regimeCode As String) _
    As Long
'
'==============================================================================
'                                RegimeSetCount
'------------------------------------------------------------------------------
' PURPOSE
'   Netting sets of one regime in the last run, whatever their status.
'
' UPDATED
'   2026-10-13
'==============================================================================
'
        RegimeSetCount = mRegSets(RegimeIndex(regimeCode))

End Property


Public Property Get RegimeValidCount( _
    ByVal regimeCode As String) _
    As Long
'
'==============================================================================
'                               RegimeValidCount
'------------------------------------------------------------------------------
' PURPOSE
'   VALID netting sets of one regime in the last run: those whose EAD is
'   in TotalEAD.
'
' UPDATED
'   2026-10-13
'==============================================================================
'
        RegimeValidCount = mRegValid(RegimeIndex(regimeCode))

End Property


Public Property Get RegimeWithheldCount( _
    ByVal regimeCode As String) _
    As Long
'
'==============================================================================
'                             RegimeWithheldCount
'------------------------------------------------------------------------------
' PURPOSE
'   Netting sets of one regime in the last run whose EAD was withheld
'   (INVALID or INCOMPLETE).
'
' UPDATED
'   2026-10-13
'==============================================================================
'
        RegimeWithheldCount = mRegWithheld(RegimeIndex(regimeCode))

End Property


Public Function TotalStatusText( _
    ByVal regimeCode As String) _
    As String
'
'==============================================================================
'                               TotalStatusText
'------------------------------------------------------------------------------
' PURPOSE
'   Describe what a regime's total covers, as written in the status column
'   of its TOTAL row: for example "3 of 4 netting set(s) VALID; 1 withheld".
'   A NO TRADES netting set is counted but neither VALID nor withheld.
'
' UPDATED
'   2026-10-13
'==============================================================================
'
        TotalStatusText = RegimeValidCount(regimeCode) & " of " & RegimeSetCount(regimeCode) & _
                          " netting set(s) VALID; " & RegimeWithheldCount(regimeCode) & " withheld"

End Function


Public Function WithheldByRegimeText() As String
'
'==============================================================================
'                             WithheldByRegimeText
'------------------------------------------------------------------------------
' PURPOSE
'   Split the netting sets whose EAD is withheld by regime, for the run
'   summary and the end-of-run message: for example "BCBS 2" or "BCBS 1,
'   CRR 1". A regime without a withheld netting set is left out; the text
'   is empty when none is withheld (#275).
'
' UPDATED
'   2026-10-16
'==============================================================================
'
        If RegimeWithheldCount(RG_BCBS) > 0 Then
            WithheldByRegimeText = RG_BCBS & " " & RegimeWithheldCount(RG_BCBS)
        End If
        If RegimeWithheldCount(RG_CRR) > 0 Then
            If Len(WithheldByRegimeText) > 0 Then
                WithheldByRegimeText = WithheldByRegimeText & ", "
            End If
            WithheldByRegimeText = WithheldByRegimeText & RG_CRR & " " & RegimeWithheldCount(RG_CRR)
        End If

End Function


Private Function RegimeIndex( _
    ByVal regimeCode As String) _
    As Long
'
'==============================================================================
'                                 RegimeIndex
'------------------------------------------------------------------------------
' PURPOSE
'   Position of a regime in the per-regime totals: 1 CRR, 2 BCBS.
'
' UPDATED
'   2026-10-13
'==============================================================================
'
        If regimeCode = RG_BCBS Then
            RegimeIndex = 2
        Else
            RegimeIndex = 1
        End If

End Function


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


Public Property Get EconomicTradesRead() As Long
'
'==============================================================================
'                              EconomicTradesRead
'------------------------------------------------------------------------------
' PURPOSE
'   Number of economic trades read by the last run: legs sharing an
'   Economic trade ID count once, a row without one counts as a trade of
'   its own (#38). TradesRead counts the legs.
'
' UPDATED
'   2026-10-12
'==============================================================================
'
        EconomicTradesRead = mEconRead

End Property


Public Property Get EconomicTradesUsed() As Long
'
'==============================================================================
'                              EconomicTradesUsed
'------------------------------------------------------------------------------
' PURPOSE
'   Number of economic trades of the last run whose legs were all used
'   (#38). TradesUsed counts the legs.
'
' UPDATED
'   2026-10-12
'==============================================================================
'
        EconomicTradesUsed = mEconUsed

End Property


Public Property Get TradesRead() As Long
'
'==============================================================================
'                                  TradesRead
'------------------------------------------------------------------------------
' PURPOSE
'   Number of trade rows (legs) with an ID read by the last run.
'
' UPDATED
'   2026-10-06
'==============================================================================
'
        TradesRead = mTradesRead

End Property


Public Property Get RunClass() As String
'
'==============================================================================
'                                   RunClass
'------------------------------------------------------------------------------
' PURPOSE
'   Class of the last run (#41): RUN_REGULATORY when every regulatory
'   parameter equals the set REF_SET_ID, RUN_WHATIF otherwise; "" when the
'   run stopped before its inputs were read.
'
' UPDATED
'   2026-10-11
'==============================================================================
'
        RunClass = mRunClass

End Property


Public Property Get ParameterDeviations() As Long
'
'==============================================================================
'                             ParameterDeviations
'------------------------------------------------------------------------------
' PURPOSE
'   Number of parameters of the last run that differ from the regulatory
'   set, each listed on Checks as a "What-if:" warning (#41).
'
' UPDATED
'   2026-10-11
'==============================================================================
'
        ParameterDeviations = mDeviations

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
        Erase mRegEAD, mRegSets, mRegValid, mRegWithheld
        mIncomplete = 0
        mParamError = False
        mRunInputs = ""
        mRunClass = ""
        mDeviations = 0
        mEconRead = 0
        mEconUsed = 0

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
'   2026-10-13
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ws         As Worksheet    'Params sheet
    Dim ok         As Boolean      'No problem found so far
    Dim layoutOk   As Boolean      'Params names and headers are valid
    Dim codeCol    As Long         'Table column of the parameter codes
    Dim valueCol   As Long         'Table column of the parameter values
    Dim lastR      As Long         'Last used row of Params
    Dim grid       As Variant      'Parameter codes, rows 1 to lastR
    Dim values     As Variant      'Parameter values, rows 1 to lastR
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
    'The input sheets are located through workbook names: a missing or
    'broken name stops the run, and the headers are checked where the names
    'point.
        Set ws = GetSheet(SH_PARAMS)
        ok = InputHeadersMatch(SH_NS, NM_NS_HEADER, NS_HEADERS) And ok
        ok = InputHeadersMatch(SH_TRADES, NM_TR_HEADER, TR_HEADERS) And ok
        layoutOk = ParamsLayoutValid(ws)
        ok = layoutOk And ok
        If Not layoutOk Then
            Exit Function
        End If

'------------------------------------------------------------------------------
' PARAMETERS
'------------------------------------------------------------------------------
    'Every parameter is read through the workbook name equal to its code
    '(CORE_Util.GetParam), so a missing or broken name stops the run. A
    'code listed twice on Params with the same value is harmless and warned
    'about; with different values it is ambiguous.
        codes = ParameterCodes()
        codeCol = ParamsCodeCol()
        valueCol = ParamsValueCol()
        lastR = UsedLastRow(ws)
        If lastR >= 1 Then
            grid = TableRange(ws, 1, codeCol, lastR, codeCol).Value2
            values = TableRange(ws, 1, valueCol, lastR, valueCol).Value2
        End If
        For i = LBound(codes) To UBound(codes)
            If NamedCell(CStr(codes(i))) Is Nothing Then
                ok = False
                LogMsg SEV_ERROR, SH_PARAMS, codes(i), "Workbook name '" & codes(i) & _
                       "' is missing or no longer refers to a cell (#REF!); every parameter is read " & _
                       "through its name."
            End If
            firstRow = 0
            For r = 1 To lastR
                If StrComp(SafeStr(grid(r, 1)), codes(i), vbTextCompare) = 0 Then
                    If firstRow = 0 Then
                        firstRow = r
                    ElseIf SameValue(values(firstRow, 1), values(r, 1)) Then
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


Public Function ParameterCodes() As Variant
'
'==============================================================================
' PURPOSE
'   The parameter vocabulary shared by schema and table validation, and
'   by the evidence report of M_Main (#45).
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


Private Function ParamsLayoutValid( _
    ByVal ws As Worksheet) _
    As Boolean
'
'==============================================================================
'                              ParamsLayoutValid
'------------------------------------------------------------------------------
' PURPOSE
'   Check the workbook names that locate the Params layout: each must exist
'   and point to Params, the "Code" and "Value" headers must be on one row,
'   and the headers found there and at the factor and FX tables must be the
'   expected ones. An error is logged for each problem.
'
' RETURNS
'   True when Params can be read through its names.
'
' UPDATED
'   2026-10-11
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim layoutNames   As Variant    'The layout names
    Dim i             As Long       'Index into layoutNames
    Dim c             As Range      'Top-left cell of a name
    Dim ok            As Boolean    'No problem found so far

'------------------------------------------------------------------------------
' NAMES
'------------------------------------------------------------------------------
        ok = True
        layoutNames = Array(NM_PRM_CODE, NM_PRM_VALUE, NM_SF_TABLE, NM_FX_TABLE)
        For i = LBound(layoutNames) To UBound(layoutNames)
            Set c = NamedCell(CStr(layoutNames(i)))
            If c Is Nothing Then
                ok = False
                LogMsg SEV_ERROR, SH_PARAMS, CStr(layoutNames(i)), "Workbook name '" & layoutNames(i) & _
                       "' is missing or no longer refers to a cell - rebuild the workbook from the template."
            ElseIf c.Worksheet.Name <> ws.Name Then
                ok = False
                LogMsg SEV_ERROR, SH_PARAMS, CStr(layoutNames(i)), "Workbook name '" & layoutNames(i) & _
                       "' must refer to a cell on " & SH_PARAMS & "."
            End If
        Next i
        If Not ok Then
            Exit Function
        End If

'------------------------------------------------------------------------------
' HEADERS
'------------------------------------------------------------------------------
        If NamedRow(NM_PRM_CODE) <> NamedRow(NM_PRM_VALUE) Then
            ok = False
            LogMsg SEV_ERROR, SH_PARAMS, NM_PRM_VALUE, "Workbook names '" & NM_PRM_CODE & "' and '" & _
                   NM_PRM_VALUE & "' must be on the same row."
        End If
        ok = HeadersMatch(ws, NamedRow(NM_PRM_CODE), PRM_CODE_HEADER, ParamsCodeCol()) And ok
        ok = HeadersMatch(ws, NamedRow(NM_PRM_VALUE), PRM_VALUE_HEADER, ParamsValueCol()) And ok
        ok = HeadersMatch(ws, NamedRow(NM_SF_TABLE), SF_HEADERS, NamedTableCol(NM_SF_TABLE)) And ok
        ok = HeadersMatch(ws, NamedRow(NM_FX_TABLE), FX_HEADERS, NamedTableCol(NM_FX_TABLE)) And ok
        ParamsLayoutValid = ok

End Function


Private Function InputHeadersMatch( _
    ByVal sheetName As String, _
    ByVal headerName As String, _
    ByVal expectedHeaders As String) _
    As Boolean
'
'==============================================================================
'                              InputHeadersMatch
'------------------------------------------------------------------------------
' PURPOSE
'   Locate an input sheet's header row through the workbook name of its
'   first header cell and compare the headers there (HeadersMatch).
'
' INPUTS
'   sheetName: SH_NS or SH_TRADES.
'   headerName: NM_NS_HEADER or NM_TR_HEADER.
'   expectedHeaders: NS_HEADERS or TR_HEADERS.
'
' RETURNS
'   True when the name locates the header row and every header matches;
'   False, with an error logged, otherwise.
'
' UPDATED
'   2026-10-11
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim r   As Long    'Header row; 0 when the name cannot locate it

'------------------------------------------------------------------------------
' LOCATE AND COMPARE
'------------------------------------------------------------------------------
        r = TableHeaderRow(sheetName)
        If r = 0 Then
            LogMsg SEV_ERROR, sheetName, headerName, "Workbook name '" & headerName & _
                   "' is missing or no longer refers to the first header cell, in column C, of " & _
                   sheetName & " - rebuild the workbook from the template."
            Exit Function
        End If
        InputHeadersMatch = HeadersMatch(GetSheet(sheetName), r, expectedHeaders)

End Function


Private Function HeadersMatch( _
    ByVal ws As Worksheet, _
    ByVal rowNum As Long, _
    ByVal expectedHeaders As String, _
    Optional ByVal firstCol As Long = 1) _
    As Boolean
'
'==============================================================================
'                                 HeadersMatch
'------------------------------------------------------------------------------
' PURPOSE
'   Compare a header row with the expected headers, ignoring case, spaces
'   and line breaks, so that a header reworded only in its spacing or
'   wrapped onto two lines still matches, and log an error for each header
'   that differs.
'
' INPUTS
'   ws, rowNum: the sheet and its header row.
'   expectedHeaders: one of the _HEADERS constants in CORE_Config; an empty
'   entry is not checked.
'   firstCol: the table column of the first header, 1 by default.
'
' RETURNS
'   True when every checked header matches.
'
' UPDATED
'   2026-10-11
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim headers   As Variant    'Expected headers, from index 0 for table column 1
    Dim c         As Long       'Index into headers
    Dim found     As Variant    'Header cell value

'------------------------------------------------------------------------------
' COMPARE
'------------------------------------------------------------------------------
        HeadersMatch = True
        headers = Split(expectedHeaders, "|")
        For c = 0 To UBound(headers)
            If Len(headers(c)) > 0 Then
                found = TableCell(ws, rowNum, firstCol + c).Value
                If StrComp(HeaderKey(SafeStr(found)), HeaderKey(CStr(headers(c))), vbTextCompare) <> 0 Then
                    HeadersMatch = False
                    LogMsg SEV_ERROR, ws.Name, TableCell(ws, rowNum, firstCol + c).Address(False, False), _
                           "Header must be '" & headers(c) & "' (found " & DescribeValue(found) & _
                           ") - a column was inserted, deleted or moved, or the header renamed.", rowNum
                End If
            End If
        Next c

End Function


Private Function HeaderKey( _
    ByVal headerText As String) _
    As String
'
'==============================================================================
'                                  HeaderKey
'------------------------------------------------------------------------------
' PURPOSE
'   Strip spaces, non-breaking spaces, tabs and line breaks from a header,
'   for HeadersMatch.
'
' UPDATED
'   2026-10-11
'==============================================================================
'
        headerText = Replace(headerText, " ", "")
        headerText = Replace(headerText, Chr$(160), "")
        headerText = Replace(headerText, vbTab, "")
        headerText = Replace(headerText, vbCr, "")
        HeaderKey = Replace(headerText, vbLf, "")

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

Private Sub ParamInRange( _
    ByVal code As String, _
    ByVal amount As Double, _
    ByVal lowest As Double, _
    ByVal highest As Double, _
    ByVal highestAllowed As Boolean, _
    ByVal wholeOnly As Boolean)
'
'==============================================================================
'                                ParamInRange
'------------------------------------------------------------------------------
' PURPOSE
'   Check a numeric parameter against its range: at least lowest, and at
'   most highest (highestAllowed) or below it; a whole number when
'   wholeOnly. The ranges are those of the data validation on Params.
'
' STATE OWNERSHIP
'   Sets mParamError and logs an error naming the parameter when the value
'   is outside the range, so that LoadParams stops the run (#147).
'
' UPDATED
'   2026-10-11
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim inRange   As Boolean    'The value satisfies the range
    Dim rule      As String     'The range, in words

'------------------------------------------------------------------------------
' CHECK
'------------------------------------------------------------------------------
        If highestAllowed Then
            inRange = (amount >= lowest And amount <= highest)
            rule = "from " & CStr(lowest) & " to " & CStr(highest)
        Else
            inRange = (amount >= lowest And amount < highest)
            rule = "at least " & CStr(lowest) & " and below " & CStr(highest)
        End If
        If wholeOnly Then
            inRange = inRange And (amount = Int(amount))
            rule = "a whole number " & rule
        End If
        If Not inRange Then
            mParamError = True
            LogMsg SEV_ERROR, SH_PARAMS, code, code & " must be " & rule & " (found " & CStr(amount) & ")."
        End If

End Sub


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
'   the workbook uses the 1904 date system, the reporting date or currency
'   is missing, the regime is not BCBS or CRR, a numeric parameter is
'   outside the range of its data validation (ParamInRange), or the IR
'   bucket correlations do not form a valid correlation matrix (#147).
'
' ERROR POLICY
'   An unexpected error while reading is logged as an error and returns
'   False.
'
' UPDATED
'   2026-10-13
'==============================================================================
'

'------------------------------------------------------------------------------
' REPORTING DATE AND CURRENCY
'------------------------------------------------------------------------------
    'Both are required: there is no sensible default. The currency is three
    'letters A to Z; case and surrounding spaces are ignored. Only the format
    'is checked, not membership of ISO 4217 (#147). Dates are read in the
    '1900 date system: in a workbook set to the 1904 system, a date typed as
    'a plain number would be read 1,462 days early, so the run stops (#40).
        On Error GoTo Fail
        If ThisWorkbook.Date1904 Then
            LogMsg SEV_ERROR, SH_PARAMS, "", "The workbook uses the 1904 date system, which is not supported. " & _
                   "Clear 'Use 1904 date system' in File > Options > Advanced and check the dates."
            Exit Function
        End If
        pAsOf = ToSerial(GetParam(PRM_ASOF))
        If pAsOf <= 0# Then
            LogMsg SEV_ERROR, SH_PARAMS, PRM_ASOF, "Reporting date missing or not a date (enter a date, " & _
                   "or text in the form yyyy-mm-dd; found " & DescribeValue(GetParam(PRM_ASOF)) & ")."
            Exit Function
        End If
        pRepCcy = UTxt(GetParam(PRM_REPCCY))
        If Len(pRepCcy) = 0 Then
            LogMsg SEV_ERROR, SH_PARAMS, PRM_REPCCY, "Reporting currency missing."
            Exit Function
        End If
        If Not pRepCcy Like "[A-Z][A-Z][A-Z]" Then
            LogMsg SEV_ERROR, SH_PARAMS, PRM_REPCCY, "Reporting currency must be three letters A to Z, " & _
                   "e.g. EUR (found " & DescribeValue(GetParam(PRM_REPCCY)) & ")."
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
    'A blank regime means CRR; anything other than BCBS or CRR stops the run,
    'including an Excel error value, which is not read as blank (#194).
        pRegime = UTxt(GetParam(PRM_REGIME))
        If IsError(GetParam(PRM_REGIME)) Then
            LogMsg SEV_ERROR, SH_PARAMS, PRM_REGIME, "Regime must be BCBS or CRR (found an error value)."
            Exit Function
        ElseIf Len(pRegime) = 0 Then
            pRegime = RG_CRR
            LogMsg SEV_WARN, SH_PARAMS, PRM_REGIME, "Regime missing - CRR assumed."
        ElseIf pRegime <> RG_BCBS And pRegime <> RG_CRR Then
            LogMsg SEV_ERROR, SH_PARAMS, PRM_REGIME, "Regime must be BCBS or CRR."
            Exit Function
        End If
        pLamThrIR = NumParam(PRM_LAMIR, 0.001)
        pLamThrCO = NumParam(PRM_LAMCO, 0.1)

'------------------------------------------------------------------------------
' RANGES
'------------------------------------------------------------------------------
    'The ranges of the data validation on Params, checked again here because
    'validation does not stop pasted or VBA-written values (#147). A floor
    'of 1 would divide by zero in the multiplier; correlations that do not
    'form a correlation matrix would make the IR effective notional the
    'square root of a negative number.
        ParamInRange PRM_FLOOR, pFloor, 0#, 1#, False, False
        ParamInRange PRM_DAYSYEAR, pDaysYear, 360#, 366#, True, False
        ParamInRange PRM_BDYEAR, pBDYear, 200#, 366#, True, True
        ParamInRange PRM_MINMAT, pMinMatBD, 0#, 250#, True, True
        ParamInRange PRM_SDFLOOR, pSDFloorBD, 0#, 250#, True, True
        ParamInRange PRM_MPOR_BIL, pMPORBil, 1#, 250#, True, True
        ParamInRange PRM_MPOR_CLR, pMPORClr, 1#, 250#, True, True
        ParamInRange PRM_MPOR_LARGE, pMPORLarge, 1#, 250#, True, True
        ParamInRange PRM_BASIS, pBasisF, 0#, 1#, True, False
        ParamInRange PRM_VOLF, pVolF, 0#, 10#, True, False
        ParamInRange PRM_RHO12, pRho12, 0#, 1#, True, False
        ParamInRange PRM_RHO23, pRho23, 0#, 1#, True, False
        ParamInRange PRM_RHO13, pRho13, 0#, 1#, True, False
        ParamInRange PRM_LAMIR, pLamThrIR, 0#, 0.1, True, False
        ParamInRange PRM_LAMCO, pLamThrCO, 0#, 1#, True, False
        If Not mParamError Then
            If 1# - pRho12 * pRho12 - pRho23 * pRho23 - pRho13 * pRho13 _
               + 2# * pRho12 * pRho23 * pRho13 < -0.000000000001 Then
                mParamError = True
                LogMsg SEV_ERROR, SH_PARAMS, PRM_RHO12, "The IR bucket correlations " & PRM_RHO12 & ", " & _
                       PRM_RHO23 & " and " & PRM_RHO13 & " do not form a valid correlation matrix: " & _
                       "the IR effective notional would be undefined (CRE52.57)."
            End If
        End If
        If mParamError Then
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
'   factor or correlation is outside 0 to 1 or a volatility outside 0 to 5
'   (#147), a key is listed twice with different values, or a row with a
'   factor lies below the first blank key and would not be read (#35).
'
' STATE OWNERSHIP
'   Fills mSF, mSFCount and mSFIndex. A key listed twice with the same
'   values is warned about and the first row kept. "BOTH" in the Regimes
'   column is stored as blank.
'
' UPDATED
'   2026-10-13
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ws       As Worksheet      'Params sheet
    Dim r        As Long           'Row being read
    Dim r0       As Long           'Header row of the table
    Dim c0       As Long           'Table column of the keys
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
        r0 = NamedRow(NM_SF_TABLE)
        c0 = NamedTableCol(NM_SF_TABLE)
        If r0 = 0 Then
            LogMsg SEV_ERROR, SH_PARAMS, HDR_SF, "Supervisory factor table not found (workbook name '" & _
                   NM_SF_TABLE & "')."
            Exit Function
        End If

'------------------------------------------------------------------------------
' READ ROWS
'------------------------------------------------------------------------------
    'Columns: 1 key, 2 asset class, 3 category, 4 factor, 5 correlation,
    '6 option volatility, 7 commodity hedging set, 8 regimes.
        r = r0 + 1
        Do While Not IsBlankCell(TableCell(ws, r, c0).Value)
            k = UTxt(TableCell(ws, r, c0).Value)
            entry.Key = k
            entry.AssetClass = UTxt(TableCell(ws, r, c0 + 1).Value2)
            entry.Category = UTxt(TableCell(ws, r, c0 + 2).Value2)
            rowBad = Not TryDbl(TableCell(ws, r, c0 + 3).Value2, 0#, entry.SF)
            rowBad = Not TryDbl(TableCell(ws, r, c0 + 4).Value2, 0#, entry.Corr) Or rowBad
            rowBad = Not TryDbl(TableCell(ws, r, c0 + 5).Value2, 0#, entry.Vol) Or rowBad
            entry.Group = UTxt(TableCell(ws, r, c0 + 6).Value2)
            entry.Regimes = UTxt(TableCell(ws, r, c0 + 7).Value2)
            If entry.Regimes = "BOTH" Then
                entry.Regimes = ""
            End If
            If rowBad Then
                bad = True
                LogMsg SEV_ERROR, SH_PARAMS, k, "Supervisory factor, correlation and volatility must be numbers.", r
            ElseIf entry.SF < 0# Or entry.SF > 1# Or entry.Corr < 0# Or entry.Corr > 1# _
                   Or entry.Vol < 0# Or entry.Vol > 5# Then
                rowBad = True
                bad = True
                LogMsg SEV_ERROR, SH_PARAMS, k, "Supervisory factor and correlation must be from 0 to 1 and " & _
                       "the option volatility from 0 to 5 (found " & CStr(entry.SF) & ", " & CStr(entry.Corr) & _
                       ", " & CStr(entry.Vol) & ").", r
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
        fxRow = NamedRow(NM_FX_TABLE)
        If fxRow > r Then
            lastR = fxRow - 1
        End If
        If ValuesAfterEnd(ws, r, lastR, c0, c0 + 3, "Supervisory factor") Then
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
'   header is missing, a currency is listed twice with different rates, a
'   row with a rate lies below the first blank currency and would not be
'   read (#35), or the reporting currency's rate is not 1 (#41).
'
' STATE OWNERSHIP
'   Fills mFXRate, mFXCount and mFXIndex. A rate that is missing or not
'   positive is warned about and the currency left out; a currency listed
'   twice with the same rate is warned about and the first row kept. The
'   reporting currency is added at rate 1 if absent; a rate other than 1
'   is an error (#41).
'
' UPDATED
'   2026-10-13
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ws     As Worksheet    'Params sheet
    Dim r      As Long         'Row being read
    Dim r0     As Long         'Header row of the table
    Dim c0     As Long         'Table column of the currencies
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
        r0 = NamedRow(NM_FX_TABLE)
        c0 = NamedTableCol(NM_FX_TABLE)
        If r0 = 0 Then
            LogMsg SEV_ERROR, SH_PARAMS, HDR_FX, "FX table not found (workbook name '" & NM_FX_TABLE & "')."
            Exit Function
        End If

'------------------------------------------------------------------------------
' READ ROWS
'------------------------------------------------------------------------------
    'A currency seen before: the same rate is harmless, a different rate is
    'ambiguous and stops the run.
        r = r0 + 1
        Do While Not IsBlankCell(TableCell(ws, r, c0).Value)
            k = UTxt(TableCell(ws, r, c0).Value)
            rate = ToDbl(TableCell(ws, r, c0 + 2).Value2, -1#)
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
        sfRow = NamedRow(NM_SF_TABLE)
        If sfRow > r Then
            lastR = sfRow - 1
        End If
        If ValuesAfterEnd(ws, r, lastR, c0, c0 + 2, "FX") Then
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
            LogMsg SEV_ERROR, SH_PARAMS, pRepCcy, "Rate of the reporting currency must be 1 (found " & _
                   CStr(mFXRate(KeyIndex(mFXIndex, pRepCcy))) & "): rates are units of the reporting " & _
                   "currency per unit (#41)."
            Exit Function
        End If
        LoadFXTable = True

End Function


Private Function FormulaNumber( _
    ByVal result As Variant, _
    ByVal functionName As String) _
    As Double
'
'==============================================================================
'                                FormulaNumber
'------------------------------------------------------------------------------
' PURPOSE
'   Return the result of a SACCR_Formulas function as a number. The engine
'   checks its inputs before it calls these functions, so an error value
'   means the two disagree on a domain (#42): the run stops with an error
'   that names the function, not with a type mismatch.
'
' ERROR POLICY
'   Raises ERR_FORMULA_DOMAIN when result is an error value; Calculate
'   withdraws the results and reports it.
'
' UPDATED
'   2026-10-12
'==============================================================================
'
        If IsError(result) Then
            Err.Raise ERR_FORMULA_DOMAIN, "CORE_Engine." & functionName, functionName & " returned " & _
                      CStr(result) & " for inputs the engine had accepted (#42)."
        End If
        FormulaNumber = CDbl(result)

End Function


Private Function ValuesAfterEnd( _
    ByVal ws As Worksheet, _
    ByVal endRow As Long, _
    ByVal lastR As Long, _
    ByVal keyCol As Long, _
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
'   keyCol: the table column of the table's keys.
'   valueCol: the table column that always holds a number in a table row:
'   the factor or the rate.
'   A row whose Params code cell holds a parameter code is a parameter row,
'   not a table row, and is skipped.
'   tableName: for the messages.
'
' RETURNS
'   True when such a row was found.
'
' UPDATED
'   2026-10-11
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim r     As Long      'Row being checked
    Dim k     As String    'Key in the key column, upper case
    Dim pc    As Long      'Table column of the parameter codes

'------------------------------------------------------------------------------
' SEARCH
'------------------------------------------------------------------------------
        pc = ParamsCodeCol()
        For r = endRow To lastR
            k = UTxt(TableCell(ws, r, keyCol).Value)
            If IsNum(TableCell(ws, r, valueCol).Value) And Not IsParameterCode(UTxt(TableCell(ws, r, pc).Value)) Then
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


Private Sub ClassifyRun()
'
'==============================================================================
'                                 ClassifyRun
'------------------------------------------------------------------------------
' PURPOSE
'   Compare the run's regulatory parameters with the regulatory set
'   REF_SET_ID (#41): the numeric parameters on Params, every row of the
'   supervisory-factor table, and the alpha override of each netting set.
'   Each difference is logged as a "What-if:" warning; the run is
'   RUN_REGULATORY when there is none, RUN_WHATIF otherwise. DaysPerYear
'   and IRBucketOffset are conventions and are not compared.
'
' STATE OWNERSHIP
'   Sets mRunClass and mDeviations.
'
' REFERENCE
'   CRE52.72; CRR Art. 274(2), 278(3), 279b, 279c, 280 to 280f, 285;
'   CRR-RTS Art. 5. Values listed in docs/methodology/README.md.
'
' UPDATED
'   2026-10-11
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim partA     As Variant    'Regulatory factor rows, first half
    Dim partB     As Variant    'Regulatory factor rows, second half
    Dim refRows   As Collection 'Every regulatory factor row, by key
    Dim refRow    As Variant    'One row: key|class|category|SF|corr|vol|group|regimes
    Dim f         As Variant    'Fields of the row
    Dim idx       As Long       'Position of the key in mSF; 0 if absent
    Dim k         As Long       'Netting set or factor row

'------------------------------------------------------------------------------
' PARAMETERS
'------------------------------------------------------------------------------
        mDeviations = 0
        CompareParam PRM_ALPHA, pAlpha, REF_ALPHA
        CompareParam PRM_FLOOR, pFloor, REF_FLOOR
        CompareParam PRM_BDYEAR, pBDYear, REF_BDYEAR
        CompareParam PRM_MINMAT, pMinMatBD, REF_MINMAT
        CompareParam PRM_SDFLOOR, pSDFloorBD, REF_SDFLOOR
        CompareParam PRM_MPOR_BIL, pMPORBil, REF_MPOR_BIL
        CompareParam PRM_MPOR_CLR, pMPORClr, REF_MPOR_CLR
        CompareParam PRM_MPOR_LARGE, pMPORLarge, REF_MPOR_LARGE
        CompareParam PRM_BASIS, pBasisF, REF_BASIS
        CompareParam PRM_VOLF, pVolF, REF_VOLF
        CompareParam PRM_RHO12, pRho12, REF_RHO12
        CompareParam PRM_RHO23, pRho23, REF_RHO23
        CompareParam PRM_RHO13, pRho13, REF_RHO13
        CompareParam PRM_LAMIR, pLamThrIR, REF_LAMIR
        CompareParam PRM_LAMCO, pLamThrCO, REF_LAMCO

'------------------------------------------------------------------------------
' SUPERVISORY-FACTOR TABLE
'------------------------------------------------------------------------------
    'Fields: key, asset class, category, factor, correlation, volatility,
    'commodity hedging set, regimes (blank for both), as the engine stores
    'them. A row missing, different or added makes the run what-if.
        partA = Array( _
            "IR|IR||0.005||0.5||", _
            "FX|FX||0.04||0.15||", _
            "CR_AAA|CR|AAA|0.0038|0.5|1||", _
            "CR_AA|CR|AA|0.0038|0.5|1||", _
            "CR_A|CR|A|0.0042|0.5|1||", _
            "CR_BBB|CR|BBB|0.0054|0.5|1||", _
            "CR_BB|CR|BB|0.0106|0.5|1||", _
            "CR_B|CR|B|0.016|0.5|1||", _
            "CR_CCC|CR|CCC|0.06|0.5|1||", _
            "CR_CQS1|CR|CQS1|0.0038|0.5|1||", _
            "CR_CQS2|CR|CQS2|0.0042|0.5|1||", _
            "CR_CQS3|CR|CQS3|0.0054|0.5|1||", _
            "CR_CQS4|CR|CQS4|0.0106|0.5|1||")
        partB = Array( _
            "CR_CQS5|CR|CQS5|0.016|0.5|1||", _
            "CR_CQS6|CR|CQS6|0.06|0.5|1||", _
            "CR_IG_INDEX|CR|IG_INDEX|0.0038|0.8|0.8||", _
            "CR_SG_INDEX|CR|SG_INDEX|0.0106|0.8|0.8||", _
            "EQ_SINGLE|EQ|SINGLE|0.32|0.5|1.2||", _
            "EQ_INDEX|EQ|INDEX|0.2|0.8|0.75||", _
            "CO_ELECTRICITY|CO|ELECTRICITY|0.4|0.4|1.5|ENERGY|", _
            "CO_OIL_GAS|CO|OIL_GAS|0.18|0.4|0.7|ENERGY|", _
            "CO_METALS|CO|METALS|0.18|0.4|0.7|METALS|", _
            "CO_AGRICULTURAL|CO|AGRICULTURAL|0.18|0.4|0.7|AGRICULTURAL|", _
            "CO_OTHER|CO|OTHER|0.18|0.4|0.7|OTHER|", _
            "CO_CLIMATIC|CO|CLIMATIC|0.18|0.4|0.7|CLIMATIC|CRR", _
            "OT|OT||0.08||1.5||CRR")
        Set refRows = New Collection
        For Each refRow In partA
            refRows.Add refRow, Split(refRow, "|")(0)
        Next refRow
        For Each refRow In partB
            refRows.Add refRow, Split(refRow, "|")(0)
        Next refRow
        For Each refRow In refRows
            f = Split(refRow, "|")
            idx = KeyIndex(mSFIndex, CStr(f(0)))
            If idx = 0 Then
                Deviation SH_PARAMS, CStr(f(0)), "factor row " & f(0) & " is missing from the factor table."
            ElseIf mSF(idx).AssetClass <> f(1) Or mSF(idx).Category <> f(2) Or mSF(idx).Group <> f(6) _
                   Or mSF(idx).Regimes <> f(7) Or Not SameAmount(mSF(idx).SF, Val(f(3))) _
                   Or Not SameAmount(mSF(idx).Corr, Val(f(4))) Or Not SameAmount(mSF(idx).Vol, Val(f(5))) Then
                Deviation SH_PARAMS, CStr(f(0)), "factor row " & f(0) & " differs from the regulatory set " & _
                          "(factor " & CStr(mSF(idx).SF) & ", correlation " & CStr(mSF(idx).Corr) & _
                          ", volatility " & CStr(mSF(idx).Vol) & "; regulatory " & f(3) & ", " & _
                          IIf(Len(f(4)) = 0, "-", f(4)) & ", " & f(5) & ")."
            End If
        Next refRow
        For k = 1 To mSFCount
            If Not InCollection(refRows, mSF(k).Key) Then
                Deviation SH_PARAMS, mSF(k).Key, "factor row " & mSF(k).Key & " is not in the regulatory set."
            End If
        Next k

'------------------------------------------------------------------------------
' ALPHA OVERRIDES
'------------------------------------------------------------------------------
    'A set without an override takes the Params alpha, already compared.
        For k = 1 To mNSCount
            If Not SameAmount(mNS(k).Alpha, pAlpha) And Not SameAmount(mNS(k).Alpha, REF_ALPHA) Then
                Deviation SH_NS, mNS(k).ID, "alpha override " & CStr(mNS(k).Alpha) & " (regulatory " & _
                          CStr(REF_ALPHA) & ")."
            End If
        Next k
        mRunClass = IIf(mDeviations = 0, RUN_REGULATORY, RUN_WHATIF)

End Sub


Private Sub CompareParam( _
    ByVal code As String, _
    ByVal amount As Double, _
    ByVal regulatory As Double)
'
'==============================================================================
'                                CompareParam
'------------------------------------------------------------------------------
' PURPOSE
'   Log a "What-if:" deviation when a parameter differs from its value in
'   the regulatory set.
'
' UPDATED
'   2026-10-11
'==============================================================================
'
        If Not SameAmount(amount, regulatory) Then
            Deviation SH_PARAMS, code, code & " is " & CStr(amount) & " (regulatory " & CStr(regulatory) & ")."
        End If

End Sub


Private Sub Deviation( _
    ByVal sheetName As String, _
    ByVal key As String, _
    ByVal detail As String)
'
'==============================================================================
'                                  Deviation
'------------------------------------------------------------------------------
' PURPOSE
'   Count one difference from the regulatory set and list it on Checks as
'   a warning starting "What-if:".
'
' UPDATED
'   2026-10-11
'==============================================================================
'
        mDeviations = mDeviations + 1
        LogMsg SEV_WARN, sheetName, key, "What-if: " & detail

End Sub


Private Function SameAmount( _
    ByVal a As Double, _
    ByVal b As Double) _
    As Boolean
'
'==============================================================================
'                                  SameAmount
'------------------------------------------------------------------------------
' PURPOSE
'   Compare a parameter with its regulatory value, allowing for the binary
'   representation of decimals entered on the sheet.
'
' UPDATED
'   2026-10-11
'==============================================================================
'
        SameAmount = Abs(a - b) <= 0.000000001 * Max2(1#, Abs(b))

End Function


Private Function InCollection( _
    ByVal items As Collection, _
    ByVal key As String) _
    As Boolean
'
'==============================================================================
'                                 InCollection
'------------------------------------------------------------------------------
' PURPOSE
'   Tell whether a collection holds an item under a key.
'
' UPDATED
'   2026-10-11
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim probe   As Variant    'The item, when present

'------------------------------------------------------------------------------
' LOOK UP
'------------------------------------------------------------------------------
    'Item(key) raises when the key is absent.
        On Error Resume Next
        probe = items.Item(key)
        InCollection = (Err.Number = 0)
        Err.Clear
        On Error GoTo 0

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
'   arrays. A duplicate ID is an error and the later row is ignored, and so
'   is a row whose ID is reserved for a TOTAL row on Results (#198).
'
' REFERENCE
'   MPOR floors CRE52.50 to CRE52.52 and CRR Art. 285(2) to (5); the
'   client-clearing floor CRE54.12 and CRR Art. 279c(1), configurable on
'   Params. VM on an unmargined set CRE52.2, CRE52.10 and CRR Art. 272(7),
'   (7a). Margin-agreement structures CRE52.74, CRE52.75 and CRR Art.
'   274(4), 275(3).
'
' UPDATED
'   2026-10-13
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ws        As Worksheet    'NettingSets sheet
    Dim lastR     As Long         'Last row with a netting-set ID
    Dim topRow    As Long         'First data row, below the headers
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
        topRow = TableFirstRow(SH_NS)
        lastR = LastDataRowAny(ws, topRow, NS_NCOLS, 0)
        If lastR < topRow Then
            LogMsg SEV_ERROR, SH_NS, "", "No netting sets defined."
            Exit Function
        End If
        n = lastR - topRow + 1
        data = TableRange(ws, topRow, 1, lastR, NS_NCOLS).Value2

'------------------------------------------------------------------------------
' READ EACH NETTING SET
'------------------------------------------------------------------------------
    'An empty row is skipped; a row with data but no ID is an error, so a
    'netting set cannot silently drop out (#35). A repeated ID is ignored and
    'makes the first set with that ID INVALID: which row's terms apply to
    'its trades is ambiguous, so its EAD is withheld. TOTAL, TOTAL CRR and
    'TOTAL BCBS label the TOTAL rows on Results (#198), so a netting set
    'with one of them would be mistaken for a total; its row is ignored.
        For i = 1 To n
            rowNum = topRow + i - 1
            id = UTxt(data(i, NS_ID))
            If Len(id) = 0 Then
                If RowHasData(data, i, NS_NCOLS, 0) Then
                    LogMsg SEV_ERROR, SH_NS, "", "Row has netting-set data but no ID - row ignored.", rowNum
                End If
                GoTo NextRow
            End If
            If id = "TOTAL" Or id = "TOTAL " & RG_CRR Or id = "TOTAL " & RG_BCBS Then
                LogMsg SEV_ERROR, SH_NS, id, "Netting set ID " & id & " is reserved for a TOTAL row on Results - row ignored.", rowNum
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
    'read is an error that makes the netting set INVALID (#35), and so is a
    'negative threshold or MTA, a remargining frequency that is not a whole
    'number of business days from 0 to 10000, an alpha override at or below
    'zero, or an MPOR override that is not a whole number above zero
    '(#147). A remargining frequency of 0 is raised to 1 with a warning.
                nsErrors = 0
                .ID = id
                .Counterparty = SafeStr(data(i, NS_CPTY))
                .Margined = NsFlag(data(i, NS_MARGINED), False, "Margined", id, rowNum, nsErrors)
                .ClearingRole = NsCode(data(i, NS_CLEARED), CL_NONE, CL_NONE & "|" & CL_CM & "|" & CL_CLIENT, _
                                       "Client clearing role", id, rowNum, nsErrors)
                .RemarginBD = NsNumber(data(i, NS_FREQ), 1#, "Remargin frequency", id, rowNum, nsErrors)
                If .RemarginBD < 0# Or .RemarginBD > 10000# Or .RemarginBD <> Int(.RemarginBD) Then
                    nsErrors = nsErrors + 1
                    LogMsg SEV_ERROR, SH_NS, id, "Remargin frequency must be a whole number of business days " & _
                           "from 0 to 10000 (found " & CStr(.RemarginBD) & ").", rowNum
                    .RemarginBD = 1#
                ElseIf .RemarginBD < 1# Then
                    LogMsg SEV_WARN, SH_NS, id, "Remargin frequency " & CStr(.RemarginBD) & _
                           " BD is below 1 - daily remargining (1 BD) used.", rowNum
                    .RemarginBD = 1#
                End If
                .LargeOrIlliquid = NsFlag(data(i, NS_LARGE), False, "Large or illiquid", id, rowNum, nsErrors)
                .Disputes = NsFlag(data(i, NS_DISPUTE), False, "Margin disputes", id, rowNum, nsErrors)
                .VM = NsNumber(data(i, NS_VM), 0#, "Net VM", id, rowNum, nsErrors)
                .NICA = NsNumber(data(i, NS_NICA), 0#, "NICA", id, rowNum, nsErrors)
                .TH = NsNumber(data(i, NS_TH), 0#, "Threshold", id, rowNum, nsErrors)
                .MTA = NsNumber(data(i, NS_MTA), 0#, "MTA", id, rowNum, nsErrors)
                If .TH < 0# Then
                    nsErrors = nsErrors + 1
                    LogMsg SEV_ERROR, SH_NS, id, "Threshold TH must not be negative (found " & CStr(.TH) & ").", rowNum
                End If
                If .MTA < 0# Then
                    nsErrors = nsErrors + 1
                    LogMsg SEV_ERROR, SH_NS, id, "MTA must not be negative (found " & CStr(.MTA) & ").", rowNum
                End If
                .Alpha = NsNumber(data(i, NS_ALPHA), pAlpha, "Alpha override", id, rowNum, nsErrors)
                If .Alpha <= 0# Then
                    nsErrors = nsErrors + 1
                    LogMsg SEV_ERROR, SH_NS, id, "Alpha override must be greater than zero (found " & _
                           CStr(.Alpha) & ") - leave it blank to use the Params alpha.", rowNum
                    .Alpha = pAlpha
                End If

    'Regime: blank takes the Params default; any other value than BCBS or
    'CRR, or an Excel error value, is an input error (#194).
                rg = UTxt(data(i, NS_REGIME))
                If Len(rg) = 0 And Not IsError(data(i, NS_REGIME)) Then
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
                .EconTrades = 0
                ovr = NsNumber(data(i, NS_MPOR), 0#, "MPOR override", id, rowNum, nsErrors)
                If IsNum(data(i, NS_MPOR)) And (ovr <= 0# Or ovr <> Int(ovr)) Then
                    nsErrors = nsErrors + 1
                    LogMsg SEV_ERROR, SH_NS, id, "MPOR override must be a whole number of business days " & _
                           "greater than zero (found " & CStr(ovr) & ") - leave it blank for none.", rowNum
                End If
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
                    .MFMargined = FormulaNumber(SACCR_MaturityFactor(0#, True, .MPOR, pMinMatBD, pBDYear), "SACCR_MaturityFactor")
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


Private Function TradeText( _
    ByVal v As Variant, _
    ByVal fieldName As String, _
    ByRef ok As Boolean, _
    ByRef msg As String) _
    As String
'
'==============================================================================
'                                  TradeText
'------------------------------------------------------------------------------
' PURPOSE
'   Read an optional text field of a trade, upper case. A blank field is ""
'   and takes the caller's default; an Excel error value, such as #N/A from
'   a broken lookup, makes the trade invalid instead of being read as blank
'   and taking that default (#194).
'
' RETURNS
'   The trimmed upper-case text; "" when blank or an error value.
'
' UPDATED
'   2026-10-13
'==============================================================================
'

'------------------------------------------------------------------------------
' READ
'------------------------------------------------------------------------------
        If IsError(v) Then
            AddErr ok, msg, fieldName & " must be text or blank (found an error value)"
        End If
        TradeText = UTxt(v)

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


Private Sub CheckEconomicTrades( _
    ByVal data As Variant, _
    ByVal n As Long, _
    ByVal topRow As Long)
'
'==============================================================================
'                             CheckEconomicTrades
'------------------------------------------------------------------------------
' PURPOSE
'   Group the Trades rows into economic trades (#38) and find the groups
'   that cannot be used. Each row is a leg with a unique Trade ID; legs
'   sharing an Economic trade ID form one economic trade, and a row with a
'   blank Economic trade ID is a whole trade on its own. An economic trade
'   with legs is rejected, every leg with it, when:
'   - not exactly one leg has MtM leg = Y, or a leg that is not the MtM leg
'     has a MtM other than 0: the economic MtM must be counted once;
'   - its legs are in different netting sets (and so possibly different
'     regimes);
'   - two of its legs are in the same asset class and not both options:
'     only option structures may be split into options [CRE52 FAQ on
'     decomposition, CRE52.42(4), CRE52.43; CRR Art. 274(6)], while legs in
'     different asset classes map one transaction to several risk
'     categories [CRE52.28; CRR Art. 277(3)];
'   - a MtM leg flag is not Y, N or blank, or a whole trade is flagged N.
'
' INPUTS
'   data: the Trades block read by ProcessTrades, (row, column).
'   n: rows in data.
'   topRow: sheet row of data row 1, for the messages.
'
' STATE OWNERSHIP
'   Fills mRowGroup, mGrpError and mGrpRejected (zeroed) and sets
'   mEconRead; adds one economic trade per group to its netting set's
'   EconTrades.
'
' UPDATED
'   2026-10-12
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim groups      As Collection    'Group number by key: E|economic ID or L|trade ID
    Dim classes     As Collection    'Per group and asset class: index into acLegs
    Dim grpName()   As String        'Economic trade ID, or the Trade ID of a whole trade
    Dim grpNS()     As String        'Netting set of the first leg
    Dim grpLegs()   As Long          'Legs of the group
    Dim grpMtmLegs() As Long         'Legs flagged Y
    Dim grpGrouped() As Boolean      'The group has an Economic trade ID
    Dim acLegs()    As Long          'Legs in one asset class of one group
    Dim acLinear()  As Long          'Of those, legs that are not options
    Dim acGroup()   As Long          'Group of each asset-class entry
    Dim acName()    As String        'Asset class of each entry
    Dim nAc         As Long          'Asset-class entries
    Dim i           As Long          'Data row
    Dim g           As Long          'Group number
    Dim a           As Long          'Asset-class entry
    Dim tid         As String        'Trade ID of the row
    Dim econ        As String        'Economic trade ID of the row
    Dim groupKey    As String        'Key of the row's group
    Dim flag        As String        'MtM leg flag of the row
    Dim nsIdx       As Long          'Netting set of a new group; 0 if unknown

'------------------------------------------------------------------------------
' GROUP THE LEGS
'------------------------------------------------------------------------------
        ReDim mRowGroup(0 To n)
        ReDim mGrpError(0 To n)
        ReDim mGrpRejected(0 To n)
        ReDim grpName(0 To n)
        ReDim grpNS(0 To n)
        ReDim grpLegs(0 To n)
        ReDim grpMtmLegs(0 To n)
        ReDim grpGrouped(0 To n)
        ReDim acLegs(1 To n + 1)
        ReDim acLinear(1 To n + 1)
        ReDim acGroup(1 To n + 1)
        ReDim acName(1 To n + 1)
        Set groups = New Collection
        Set classes = New Collection
        mEconRead = 0
        nAc = 0
        For i = 1 To n
            tid = SafeStr(data(i, TR_ID))
            If Len(tid) > 0 Then
                econ = SafeStr(data(i, TR_ECON))
                If Len(econ) = 0 Then
                    groupKey = "L|" & UCase$(tid)
                Else
                    groupKey = "E|" & UCase$(econ)
                End If
                g = KeyIndex(groups, groupKey)
                If g = 0 Then
                    mEconRead = mEconRead + 1
                    g = mEconRead
                    KeyAdd groups, groupKey, g
                    grpName(g) = IIf(Len(econ) = 0, tid, econ)
                    grpGrouped(g) = (Len(econ) > 0)
                    grpNS(g) = UTxt(data(i, TR_NS))
                    nsIdx = KeyIndex(mNSIndex, grpNS(g))
                    If nsIdx > 0 Then
                        mNS(nsIdx).EconTrades = mNS(nsIdx).EconTrades + 1
                    End If
                ElseIf grpGrouped(g) And UTxt(data(i, TR_NS)) <> grpNS(g) Then
                    AddGroupError g, "its legs are in different netting sets (" & grpNS(g) & ", " & _
                                  UTxt(data(i, TR_NS)) & ")"
                End If
                mRowGroup(i) = g
                grpLegs(g) = grpLegs(g) + 1
                If IsError(data(i, TR_ECON)) Then
                    AddGroupError g, "Economic trade ID must be text or blank (found an error value on row " & _
                                  (topRow + i - 1) & ")"
                End If

    'The MtM leg flag: Y, N or blank (N for a leg, Y for a whole trade).
    'An error value is not read as blank.
                flag = UTxt(data(i, TR_MTMLEG))
                If IsError(data(i, TR_MTMLEG)) Then
                    AddGroupError g, "MtM leg must be Y or N (found an error value on row " & _
                                  (topRow + i - 1) & ")"
                ElseIf flag = "Y" Then
                    grpMtmLegs(g) = grpMtmLegs(g) + 1
                ElseIf flag = "N" Then
                    If Not grpGrouped(g) Then
                        AddGroupError g, "a trade without an Economic trade ID carries its own MtM, so " & _
                                      "MtM leg must be Y or blank"
                    End If
                ElseIf Len(flag) > 0 Then
                    AddGroupError g, "MtM leg must be Y or N (found " & DescribeValue(data(i, TR_MTMLEG)) & _
                                  " on row " & (topRow + i - 1) & ")"
                End If

    'Asset classes of the group's legs, for the decomposition rule.
                a = KeyIndex(classes, g & "|" & UTxt(data(i, TR_AC)))
                If a = 0 Then
                    nAc = nAc + 1
                    a = nAc
                    KeyAdd classes, g & "|" & UTxt(data(i, TR_AC)), a
                    acGroup(a) = g
                    acName(a) = UTxt(data(i, TR_AC))
                End If
                acLegs(a) = acLegs(a) + 1
                If UTxt(data(i, TR_INSTR)) <> "OPTION" Then
                    acLinear(a) = acLinear(a) + 1
                End If
            End If
        Next i

'------------------------------------------------------------------------------
' CHECK EACH ECONOMIC TRADE
'------------------------------------------------------------------------------
    'The MtM is counted once: exactly one MtM leg, and 0 on every other leg.
        For g = 1 To mEconRead
            If grpGrouped(g) And grpMtmLegs(g) <> 1 Then
                AddGroupError g, "exactly one leg must be the MtM leg (MtM leg = Y), found " & grpMtmLegs(g)
            End If
        Next g
        For i = 1 To n
            g = mRowGroup(i)
            If g > 0 Then
                If grpGrouped(g) And UTxt(data(i, TR_MTMLEG)) <> "Y" And IsNum(data(i, TR_MTM)) Then
                    If CDbl(data(i, TR_MTM)) <> 0# Then
                        AddGroupError g, "leg " & SafeStr(data(i, TR_ID)) & " is not the MtM leg, so its MtM " & _
                                      "must be 0 (found " & DescribeValue(data(i, TR_MTM)) & ")"
                    End If
                End If
            End If
        Next i

    'Legs in one asset class must all be options: a linear product is not
    'split into legs of the same class. A Trade ID repeated without an
    'Economic trade ID is not a group: ProcessTrades rejects the repeat.
        For a = 1 To nAc
            If grpGrouped(acGroup(a)) And acLegs(a) > 1 And acLinear(a) > 0 Then
                AddGroupError acGroup(a), "it has " & acLegs(a) & " legs in asset class " & acName(a) & _
                              " that are not all options: only an option structure may be split into " & _
                              "options (CRE52.42-43; CRR Art. 274(6)), and one risk category takes " & _
                              "one leg (CRR Art. 277(3))"
            End If
        Next a
        For g = 1 To mEconRead
            If Len(mGrpError(g)) > 0 Then
                mGrpError(g) = "economic trade " & grpName(g) & ": " & mGrpError(g)
            End If
        Next g

End Sub


Private Sub AddGroupError( _
    ByVal g As Long, _
    ByVal reason As String)
'
'==============================================================================
'                                AddGroupError
'------------------------------------------------------------------------------
' PURPOSE
'   Add a reason why economic trade g is rejected, once.
'
' UPDATED
'   2026-10-12
'==============================================================================
'
        If InStr(1, mGrpError(g), reason, vbBinaryCompare) = 0 Then
            If Len(mGrpError(g)) > 0 Then
                mGrpError(g) = mGrpError(g) & "; "
            End If
            mGrpError(g) = mGrpError(g) & reason
        End If

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
'   not, gets one TradeCalc row. The steps run in this order for each
'   trade, on one tTrade record: ReadTradeClass, ReadTradeAmounts,
'   ReadTradeDates, TradeDelta, TradeHedgingSet, UseTrade and
'   TradeCalcRow (#60).
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
'   2026-10-14
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ws        As Worksheet    'Trades sheet
    Dim wsOut     As Worksheet    'TradeCalc sheet
    Dim outTop    As Long         'First TradeCalc data row
    Dim lastR     As Long         'Last row with a trade ID
    Dim topRow    As Long         'First data row, below the headers
    Dim n         As Long         'Number of rows to read
    Dim i         As Long         'Row within data
    Dim rowNum    As Long         'Sheet row of data row i
    Dim data      As Variant      'Input block, (row, column)
    Dim outArr()  As Variant      'TradeCalc rows, (row, column)
    Dim nOut      As Long         'TradeCalc rows filled
    Dim seenIds   As Collection   'Trade IDs seen so far, mapped to their sheet row
    Dim nsIdx     As Long         'Netting set of a row without a trade ID
    Dim t         As tTrade       'The trade being processed
    Dim fresh     As tTrade       'Every field at its default, to reset t

'------------------------------------------------------------------------------
' READ THE INPUT BLOCK
'------------------------------------------------------------------------------
    'With no trades there is nothing to calculate; the old TradeCalc rows
    'are cleared so they cannot be mistaken for results.
        Set ws = GetSheet(SH_TRADES)
        topRow = TableFirstRow(SH_TRADES)
        lastR = LastDataRowAny(ws, topRow, TR_NCOLS, TR_COMMENT)
        If lastR < topRow Then
            LogMsg SEV_ERROR, SH_TRADES, "", "No trades found."
            If writeOutputs Then
                ClearOutputBlock GetSheet(SH_TRADECALC), OutputFirstRow(SH_TRADECALC), TC_NCOLS
            End If
            Exit Sub
        End If
        n = lastR - topRow + 1
    'Value2 gives the stored values: a Currency or Accounting format would
    'otherwise round a price or rate to four decimals, and dates come as
    'serials, which ToSerial reads (#196).
        data = TableRange(ws, topRow, 1, lastR, TR_NCOLS).Value2
        ReDim outArr(1 To n, 1 To TC_NCOLS)
        nOut = 0

'------------------------------------------------------------------------------
' PROCESS EACH TRADE
'------------------------------------------------------------------------------
    'An empty row is skipped. A row with data but no trade ID is rejected,
    'and so is a repeated trade ID; either makes its netting set INCOMPLETE
    '(#35). t starts from fresh for every trade, so nothing leaks from the
    'previous trade into the output row.
        CheckEconomicTrades data, n, topRow
        Set seenIds = New Collection
        For i = 1 To n
            rowNum = topRow + i - 1
            t = fresh
            t.tid = SafeStr(data(i, TR_ID))
            If Len(t.tid) = 0 Then
                If RowHasData(data, i, TR_NCOLS, TR_COMMENT) Then
                    mTradesRead = mTradesRead + 1
                    LogMsg SEV_ERROR, SH_TRADES, "", "Row has trade data but no Trade ID - trade rejected.", rowNum
                    nsIdx = KeyIndex(mNSIndex, UTxt(data(i, TR_NS)))
                    If nsIdx > 0 Then mNS(nsIdx).Rejected = mNS(nsIdx).Rejected + 1
                End If
                GoTo NextTrade
            End If
            mTradesRead = mTradesRead + 1
            t.ok = True
            t.pairSign = 1#
            ReadTradeClass t, data, i, rowNum, seenIds
            ReadTradeAmounts t, data, i
            ReadTradeDates t, data, i
            TradeDelta t, data, i
            TradeHedgingSet t
            UseTrade t, rowNum
            If Not t.ok Then LogMsg SEV_ERROR, SH_TRADES, t.tid, "Trade excluded: " & t.msg, rowNum
            If Not t.ok And t.nsIdx > 0 Then mNS(t.nsIdx).Rejected = mNS(t.nsIdx).Rejected + 1
            If Not t.ok Then mGrpRejected(mRowGroup(i)) = mGrpRejected(mRowGroup(i)) + 1
            If Len(t.warn) > 0 Then LogMsg SEV_WARN, SH_TRADES, t.tid, t.warn, rowNum
            nOut = nOut + 1
            TradeCalcRow t, outArr, nOut
NextTrade:
        Next i
        For i = 1 To mEconRead
            If mGrpRejected(i) = 0 Then
                mEconUsed = mEconUsed + 1
            End If
        Next i

'------------------------------------------------------------------------------
' WRITE TRADECALC
'------------------------------------------------------------------------------
        If writeOutputs Then
            Set wsOut = GetSheet(SH_TRADECALC)
            outTop = OutputFirstRow(SH_TRADECALC)
            ClearOutputBlock wsOut, outTop, TC_NCOLS
            WriteBlock wsOut, outTop, outArr, nOut, TC_NCOLS
            FormatColumns wsOut, outTop, nOut, _
                "@|@|@|@|@|@|0.00%|#,##0|#,##0|0.0000|0.0000|0.0000|0.0000|0.0000|#,##0|0.0000|0.0000|0.0000|#,##0|#,##0|0|#,##0|0.0000%|@|@"
        End If

End Sub


Private Sub ReadTradeClass( _
    ByRef t As tTrade, _
    ByRef data As Variant, _
    ByVal i As Long, _
    ByVal rowNum As Long, _
    ByVal seenIds As Collection)
'
'==============================================================================
'                                ReadTradeClass
'------------------------------------------------------------------------------
' PURPOSE
'   Read and check a trade's classification: a repeated Trade ID, netting
'   set, economic-trade errors, asset class, sub-class, risk factor,
'   instrument, direction and nature; then its supervisory-factor row and
'   the regime rules for OT and for regime-limited factor rows (#60).
'
' INPUTS
'   t: the trade being processed, from ProcessTrades.
'   data, i: the Trades input block and the trade's row in it.
'   rowNum: the trade's sheet row; seenIds: Trade IDs seen so far.
'
' UPDATED
'   2026-10-14
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim firstRow   As Long      'Sheet row where a duplicate ID first appeared
    Dim dirTxt     As String    'Direction text, upper case

'------------------------------------------------------------------------------
' CLASSIFICATION
'------------------------------------------------------------------------------
    'Netting set, asset class, risk factor, instrument, direction and
    'nature. Basis and volatility trades form their own hedging sets, one
    'per label, with the factor multiplied by BasisFactor or
    'VolatilityFactor [CRE52.46, CRE52.47].
        firstRow = KeyIndex(seenIds, t.tid)
        If firstRow > 0 Then
            AddErr t.ok, t.msg, "duplicate Trade ID (first used on row " & firstRow & ")"
        Else
            KeyAdd seenIds, t.tid, rowNum
        End If

        t.nsId = UTxt(data(i, TR_NS))
        t.nsIdx = KeyIndex(mNSIndex, t.nsId)
        If t.nsIdx = 0 Then AddErr t.ok, t.msg, "unknown netting set '" & t.nsId & "'"
        If Len(mGrpError(mRowGroup(i))) > 0 Then AddErr t.ok, t.msg, mGrpError(mRowGroup(i))

        t.acTxt = UTxt(data(i, TR_AC))
        t.ac = ACIndex(t.acTxt)
        If t.ac = 0 Then AddErr t.ok, t.msg, "asset class must be IR, FX, CR, EQ, CO or OT"

        t.subCls = TradeText(data(i, TR_SUB), "Sub-class", t.ok, t.msg)
        t.rf = UTxt(data(i, TR_RF))
        If Len(t.rf) = 0 Then AddErr t.ok, t.msg, "risk factor / reference missing"
        If InStr(t.rf, "|") > 0 Or InStr(t.rf, "#") > 0 Then
            AddErr t.ok, t.msg, "risk factor / reference must not contain '|' or '#'"
        End If
        If t.ac = AC_IR And Len(t.rf) > 0 And _
           Not (t.rf Like "[A-Z][A-Z][A-Z]" Or t.rf Like "[A-Z][A-Z][A-Z]-INFL") Then
            AddErr t.ok, t.msg, "interest-rate risk factor must be a 3-letter currency code such as EUR, " & _
                   "or EUR-INFL for inflation"
        End If

        t.instType = TradeText(data(i, TR_INSTR), "Instrument", t.ok, t.msg)
        If Len(t.instType) = 0 Then t.instType = "LINEAR"
        If t.instType <> "LINEAR" And t.instType <> "OPTION" And t.instType <> "CDO" Then
            AddErr t.ok, t.msg, "instrument type must be Linear, Option or CDO"
        End If
        If t.instType = "CDO" And t.ac <> AC_CR Then AddErr t.ok, t.msg, "CDO tranches are credit (CR) trades"

        dirTxt = UTxt(data(i, TR_DIR))
        Select Case dirTxt
            Case "LONG", "BUY", "BOUGHT", "L", "B"
                t.isLong = True
            Case "SHORT", "SELL", "SOLD", "S"
                t.isLong = False
            Case Else
                AddErr t.ok, t.msg, "direction must be Long or Short"
        End Select

        t.nature = TradeText(data(i, TR_NATURE), "Nature", t.ok, t.msg)
        If Len(t.nature) = 0 Then t.nature = "STANDARD"
        t.lbl = TradeText(data(i, TR_LABEL), "Hedging-set label", t.ok, t.msg)
        If InStr(t.lbl, "|") > 0 Or InStr(t.lbl, "#") > 0 Then
            AddErr t.ok, t.msg, "hedging-set label must not contain '|' or '#'"
        End If
        Select Case t.nature
            Case "STANDARD"
                t.natFactor = 1#
                t.natTag = "STD"
            Case "BASIS"
                t.natFactor = pBasisF
                t.natTag = "BASIS:" & t.lbl
                If Len(t.lbl) = 0 Then AddErr t.ok, t.msg, "basis trade needs a hedging-set label"
            Case "VOLATILITY"
                t.natFactor = pVolF
                t.natTag = "VOL:" & t.lbl
                If Len(t.lbl) = 0 Then AddErr t.ok, t.msg, "volatility trade needs a hedging-set label"
            Case Else
                AddErr t.ok, t.msg, "nature must be Standard, Basis or Volatility"
        End Select

    'IR, FX and OT have one factor each; the other classes are looked up by
    'class and sub-class, for example CR_AAA.
        If t.ac > 0 Then
            Select Case t.ac
                Case AC_IR
                    t.sfKey = "IR"
                Case AC_FX
                    t.sfKey = "FX"
                Case AC_OT
                    t.sfKey = "OT"
                Case Else
                    t.sfKey = t.acTxt & "_" & t.subCls
            End Select
            t.sfIdx = KeyIndex(mSFIndex, t.sfKey)
            If t.sfIdx = 0 Then AddErr t.ok, t.msg, "no supervisory factor for '" & t.sfKey & "' (check sub-class)"
        End If

    'OT exists only under CRR, and a factor-table row limited to one regime
    'cannot be used by a netting set under the other.
        If t.nsIdx > 0 And t.ac = AC_OT Then
            If Not mNS(t.nsIdx).IsCRR Then
                AddErr t.ok, t.msg, "asset class OT (other risks) exists only under CRR (Art. 277(1)(f), 280f); " & _
                       "under BCBS map the trade to the class of its primary risk driver (CRE52)"
            End If
        End If
        If t.nsIdx > 0 And t.sfIdx > 0 Then
            If Len(mSF(t.sfIdx).Regimes) > 0 And mSF(t.sfIdx).Regimes <> mNS(t.nsIdx).Regime Then
                AddErr t.ok, t.msg, "'" & t.sfKey & "' is defined only for regime " & mSF(t.sfIdx).Regimes & _
                       " (netting set uses " & mNS(t.nsIdx).Regime & ")"
            End If
        End If

End Sub


Private Sub ReadTradeAmounts( _
    ByRef t As tTrade, _
    ByRef data As Variant, _
    ByVal i As Long)
'
'==============================================================================
'                               ReadTradeAmounts
'------------------------------------------------------------------------------
' PURPOSE
'   Read a trade's notional and MtM with their currencies, and convert
'   both to the reporting currency when the trade is still valid (#60).
'
' INPUTS
'   t: the trade being processed, from ProcessTrades.
'   data, i: the Trades input block and the trade's row in it.
'
' UPDATED
'   2026-10-14
'==============================================================================
'

'------------------------------------------------------------------------------
' AMOUNTS
'------------------------------------------------------------------------------
    'The notional is unsigned; Direction gives the sign. A missing MtM is an
    'error, not an assumed 0 (#35); a blank notional currency means the
    'reporting currency, and a blank MtM currency the notional currency.
    'Both amounts are converted to the reporting currency.
        If IsNum(data(i, TR_NOTIONAL)) Then
            t.notional = CDbl(data(i, TR_NOTIONAL))
            If t.notional < 0# Then AddErr t.ok, t.msg, "notional must be positive (use Direction for the sign)"
        ElseIf IsBlankCell(data(i, TR_NOTIONAL)) And Not IsError(data(i, TR_NOTIONAL)) Then
            AddErr t.ok, t.msg, "notional missing"
        Else
            AddErr t.ok, t.msg, "notional must be a number (found " & DescribeValue(data(i, TR_NOTIONAL)) & ")"
        End If
        t.nccy = TradeText(data(i, TR_NCCY), "Notional currency", t.ok, t.msg)
        If Len(t.nccy) = 0 Then t.nccy = pRepCcy
        t.fxN = FXRate(t.nccy)
        If t.fxN < 0# Then AddErr t.ok, t.msg, "no FX rate for notional currency '" & t.nccy & "'"
        If IsNum(data(i, TR_MTM)) Then
            t.mtm = CDbl(data(i, TR_MTM))
        ElseIf IsBlankCell(data(i, TR_MTM)) And Not IsError(data(i, TR_MTM)) Then
            t.mtm = 0#
            AddErr t.ok, t.msg, "MtM missing (enter 0 if the trade has no value)"
        Else
            t.mtm = 0#
            AddErr t.ok, t.msg, "MtM must be a number (found " & DescribeValue(data(i, TR_MTM)) & ")"
        End If
        t.mccy = TradeText(data(i, TR_MCCY), "MtM currency", t.ok, t.msg)
        If Len(t.mccy) = 0 Then t.mccy = t.nccy
        t.fxM = FXRate(t.mccy)
        If t.fxM < 0# Then AddErr t.ok, t.msg, "no FX rate for MtM currency '" & t.mccy & "'"
        If t.ok Then
            t.notionalRep = t.notional * t.fxN
            t.mtmRep = t.mtm * t.fxM
        End If

End Sub


Private Sub ReadTradeDates( _
    ByRef t As tTrade, _
    ByRef data As Variant, _
    ByVal i As Long)
'
'==============================================================================
'                                ReadTradeDates
'------------------------------------------------------------------------------
' PURPOSE
'   Read a trade's start, end, maturity and expiry dates and turn them into
'   the years S, E and M from the reporting date (#60).
'
' INPUTS
'   t: the trade being processed, from ProcessTrades.
'   data, i: the Trades input block and the trade's row in it.
'
' UPDATED
'   2026-10-14
'==============================================================================
'

'------------------------------------------------------------------------------
' DATES
'------------------------------------------------------------------------------
    'A date that is present but cannot be read is an error (#35). Missing
    'dates are filled from each other: maturity from the end date,
    'then from the option expiry; the end date from the maturity. S, E and M
    'are years from the reporting date on the DaysPerYear basis, with S
    'floored at 0.
        t.dStart = TradeDate(data(i, TR_START), "start date", t.ok, t.msg)
        t.dEnd = TradeDate(data(i, TR_END), "end date", t.ok, t.msg)
        t.dMat = TradeDate(data(i, TR_MAT), "maturity date", t.ok, t.msg)
        t.dExp = TradeDate(data(i, TR_EXPIRY), "option expiry", t.ok, t.msg)
        If t.dMat < 0# Then t.dMat = t.dEnd
        If t.dMat < 0# Then t.dMat = t.dExp
        If t.dEnd < 0# Then t.dEnd = t.dMat
        If t.dMat < 0# Then
            AddErr t.ok, t.msg, "maturity / end date missing"
        Else
            t.M = (t.dMat - pAsOf) / pDaysYear
            t.E = (t.dEnd - pAsOf) / pDaysYear
            If t.dStart >= 0# Then
                t.S = Max2(0#, (t.dStart - pAsOf) / pDaysYear)
            Else
                t.S = 0#
            End If
            If t.M <= 0# Then
                AddErr t.ok, t.msg, "trade has matured (maturity <= reporting date)"
            ElseIf (t.ac = AC_IR Or t.ac = AC_CR) And t.E <= t.S Then
                AddErr t.ok, t.msg, "end date must be after start date"
            End If
        End If

    'An interest-rate or credit option takes S from its underlying
    '[CRE52.31(2), Table 1 of CRE52.32]: for a swaption the exercise date,
    'for a Bermudan swaption the earliest allowed exercise date
    '[CRE52.42(2)], for a cap or floor the start of its first period. A
    'blank start date would read as 0, the duration of a swap starting
    'today, so it is an error (#40).
        If t.instType = "OPTION" And (t.ac = AC_IR Or t.ac = AC_CR) And t.dStart < 0# Then
            If IsBlankCell(data(i, TR_START)) And Not IsError(data(i, TR_START)) Then
                AddErr t.ok, t.msg, "an interest-rate or credit option needs the start date S of its underlying " & _
                       "(swaption: the exercise date, the earliest one if Bermudan; cap or floor: the " & _
                       "start of the first period)"
            End If
        End If

End Sub


Private Sub TradeDelta( _
    ByRef t As tTrade, _
    ByRef data As Variant, _
    ByVal i As Long)
'
'==============================================================================
'                                  TradeDelta
'------------------------------------------------------------------------------
' PURPOSE
'   Compute the supervisory delta of a valid trade: linear, option (with
'   its expiry T and lambda shift) or CDO tranche (#60).
'
' INPUTS
'   t: the trade being processed, from ProcessTrades.
'   data, i: the Trades input block and the trade's row in it.
'
' UPDATED
'   2026-10-14
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim P        As Double     'Underlying price
    Dim K        As Double     'Strike
    Dim attA     As Double     'CDO attachment point
    Dim detD     As Double     'CDO detachment point
    Dim lamIn    As Variant    'Lambda as entered
    Dim vDelta   As Variant    'Option delta as returned by SACCR_OptionDelta

'------------------------------------------------------------------------------
' SUPERVISORY DELTA
'------------------------------------------------------------------------------
    'Linear trades: +1 long, -1 short [CRE52.39]. Options: the option delta
    'with the lambda shift; under CRR the regulatory lambda replaces any
    'entered value for IR and commodity options [CRE52.40]. CDO tranches:
    'the tranche delta [CRE52.41].
        If t.ok Then
            Select Case t.instType
                Case "LINEAR"
                    If t.isLong Then
                        t.delta = 1#
                    Else
                        t.delta = -1#
                    End If
                Case "OPTION"
                    t.optTxt = UTxt(data(i, TR_OPT))
                    If t.optTxt <> "CALL" And t.optTxt <> "PUT" Then AddErr t.ok, t.msg, "option type must be Call or Put"
                    If t.dExp < 0# Then
                        AddErr t.ok, t.msg, "option expiry missing"
                    Else
                        t.T = (t.dExp - pAsOf) / pDaysYear
                        If t.T <= 0# Then
                            AddErr t.ok, t.msg, "option has expired"
                        ElseIf t.dExp > t.dMat Then
                            AddErr t.ok, t.msg, "option expiry is after the maturity date (M is the latest " & _
                                   "date the contract may still be active)"
                        End If
                    End If
                    If Not IsNum(data(i, TR_PRICE)) Or Not IsNum(data(i, TR_STRIKE)) Then
                        AddErr t.ok, t.msg, "option needs underlying price and strike"
                    End If
                    If t.ok Then
                        P = CDbl(data(i, TR_PRICE))
                        K = CDbl(data(i, TR_STRIKE))
                        lamIn = data(i, TR_LAMBDA)
                        If Not IsNum(lamIn) And Not (IsBlankCell(lamIn) And Not IsError(lamIn)) Then
                            AddErr t.ok, t.msg, "lambda must be a number (found " & DescribeValue(lamIn) & ")"
                        End If
                        If mNS(t.nsIdx).IsCRR And (t.ac = AC_IR Or t.ac = AC_CO) Then
                            'Delegated Regulation (EU) 2021/931 Art. 5, as amended by 2025/855.
                            If t.ac = AC_IR Then
                                t.lam = FormulaNumber(SACCR_LambdaCRR(P, K, True, pLamThrIR), "SACCR_LambdaCRR")
                            Else
                                t.lam = FormulaNumber(SACCR_LambdaCRR(P, K, False, pLamThrCO), "SACCR_LambdaCRR")
                            End If
                            If IsNum(lamIn) Then
                                If Abs(CDbl(lamIn) - t.lam) > 0.000000001 Then
                                    AddWarn t.warn, "lambda input ignored under CRR - RTS lambda applied"
                                End If
                            End If
                        Else
                            t.lam = ToDbl(lamIn, 0#)
                        End If
                        t.lamOut = t.lam
                        If P + t.lam <= 0# Or K + t.lam <= 0# Then
                            AddErr t.ok, t.msg, "price and strike must be > 0 after lambda shift (BCBS: enter lambda per CRE52.40 FAQ)"
                        ElseIf mSF(t.sfIdx).Vol <= 0# Then
                            AddErr t.ok, t.msg, "supervisory volatility missing for '" & t.sfKey & "'"
                        Else
                            vDelta = SACCR_OptionDelta(P, K, t.T, mSF(t.sfIdx).Vol, (t.optTxt = "CALL"), t.isLong, t.lam)
                            t.delta = FormulaNumber(vDelta, "SACCR_OptionDelta")
                        End If
                    End If
                Case "CDO"
                    If Not IsNum(data(i, TR_ATTACH)) Or Not IsNum(data(i, TR_DETACH)) Then
                        AddErr t.ok, t.msg, "CDO tranche needs attachment and detachment points"
                    Else
                        attA = CDbl(data(i, TR_ATTACH))
                        detD = CDbl(data(i, TR_DETACH))
                        If attA < 0# Or detD > 1# Or attA >= detD Then
                            AddErr t.ok, t.msg, "attachment/detachment must satisfy 0 <= A < D <= 1"
                        Else
                            t.delta = FormulaNumber(SACCR_CDODelta(attA, detD, t.isLong), "SACCR_CDODelta")
                        End If
                    End If
            End Select
        End If

End Sub


Private Sub TradeHedgingSet( _
    ByRef t As tTrade)
'
'==============================================================================
'                               TradeHedgingSet
'------------------------------------------------------------------------------
' PURPOSE
'   Give a valid trade its adjusted notional d, hedging set and bucket, with
'   the FX pair checks (#60).
'
' INPUTS
'   t: the trade being processed, from ProcessTrades.
'
' UPDATED
'   2026-10-14
'==============================================================================
'

'------------------------------------------------------------------------------
' HEDGING SET, ADJUSTED NOTIONAL AND BUCKET
'------------------------------------------------------------------------------
    'IR: d = notional * SD, hedging set per currency (risk factor), bucket
    'by end date [CRE52.34, CRE52.57(2)-(3)]: under CRR up to and including
    '1 year, over 1 to 5 years, over 5 years [Art. 280a(3), Table 2]; under
    'BCBS under 1 year, 1 to 5 years, over 5 years, so an end date of
    'exactly 1 year is bucket 1 under CRR and bucket 2 under BCBS (#191). The
    'key of bucket 1 stays "<1Y" in both regimes. An inflation risk factor
    'such as EUR-INFL is its own hedging set [CRR Art. 277(4)(a), 277a(1)].
    'FX: d = notional, hedging set per currency pair written in alphabetical
    'order; an inverted pair flips the delta [CRE52.58]; both currencies of
    'the pair need an FX rate, and the notional is that of one of the pair's
    'legs (#40). An FX option must be entered in market quote order
    '(MarketPair): a call on one order and the put on the other give
    'different deltas, so the order must be one and the same for every
    'trade [CRE52.59 FAQ1] (#40). CR: d = notional * SD, one hedging set, bucket per entity
    '[CRE52.60, CRE52.64]. EQ: one hedging set, bucket per entity
    '[CRE52.66]. CO: hedging set per commodity group, bucket per commodity
    '[CRE52.70]. OT: hedging set per primary risk driver [CRR Art. 277a].
        If t.ok Then
            Select Case t.ac
                Case AC_IR
                    t.sd = FormulaNumber(SACCR_SupervisoryDuration(t.S, t.E, pSDFloorBD / pBDYear), "SACCR_SupervisoryDuration")
                    t.adjN = t.notionalRep * t.sd
                    If t.E < 1# Or (mNS(t.nsIdx).IsCRR And t.E <= 1#) Then
                        t.bucketNo = 1
                        t.subKey = "<1Y"
                    ElseIf t.E <= 5# Then
                        t.bucketNo = 2
                        t.subKey = "1-5Y"
                    Else
                        t.bucketNo = 3
                        t.subKey = ">5Y"
                    End If
                    t.hsKey = "IR|" & t.rf & "|" & t.natTag
                Case AC_FX
                    t.rf = NormalizePair(t.rf, t.pairSign)
                    If t.pairSign = 0# Then
                        AddErr t.ok, t.msg, "FX risk factor must be a currency pair such as EUR/USD"
                    ElseIf FXRate(Left$(t.rf, 3)) < 0# Then
                        AddErr t.ok, t.msg, "no FX rate for currency '" & Left$(t.rf, 3) & "' of the pair " & t.rf
                    ElseIf FXRate(Right$(t.rf, 3)) < 0# Then
                        AddErr t.ok, t.msg, "no FX rate for currency '" & Right$(t.rf, 3) & "' of the pair " & t.rf
                    ElseIf t.nccy <> Left$(t.rf, 3) And t.nccy <> Right$(t.rf, 3) Then
                        AddErr t.ok, t.msg, "notional currency '" & t.nccy & "' is not a currency of the pair " & t.rf
                    ElseIf t.instType = "OPTION" And t.pairSign * MarketSign(t.rf) < 0# Then
                        AddErr t.ok, t.msg, "FX option must be entered in market quote order as " & MarketPair(t.rf) & _
                               ", with price and strike in " & Right$(MarketPair(t.rf), 3) & " per " & _
                               Left$(MarketPair(t.rf), 3) & " (CRE52.59 FAQ1)"
                    End If
                    t.delta = t.delta * t.pairSign
                    t.adjN = t.notionalRep
                    t.hsKey = "FX|" & t.rf & "|" & t.natTag
                    t.subKey = t.rf
                Case AC_CR
                    t.sd = FormulaNumber(SACCR_SupervisoryDuration(t.S, t.E, pSDFloorBD / pBDYear), "SACCR_SupervisoryDuration")
                    t.adjN = t.notionalRep * t.sd
                    t.hsKey = "CR|" & t.natTag
                    t.subKey = t.rf
                Case AC_EQ
                    t.adjN = t.notionalRep
                    t.hsKey = "EQ|" & t.natTag
                    t.subKey = t.rf
                Case AC_CO
                    t.adjN = t.notionalRep
                    If Len(mSF(t.sfIdx).Group) = 0 Then
                        AddErr t.ok, t.msg, "commodity hedging group missing for '" & t.sfKey & "'"
                    End If
                    t.hsKey = "CO|" & mSF(t.sfIdx).Group & "|" & t.natTag
                    t.subKey = t.rf
                Case AC_OT
                    'Art. 277a: same hedging set only for an identical primary risk driver.
                    t.adjN = t.notionalRep
                    t.hsKey = "OT|" & t.rf & "|" & t.natTag
                    t.subKey = t.rf
            End Select
        End If

End Sub


Private Sub UseTrade( _
    ByRef t As tTrade, _
    ByVal rowNum As Long)
'
'==============================================================================
'                                   UseTrade
'------------------------------------------------------------------------------
' PURPOSE
'   Compute the maturity factors and effective notionals of a valid trade
'   and add it to its bucket, its netting set's MtM and the trade counts
'   (#60).
'
' INPUTS
'   t: the trade being processed, from ProcessTrades.
'   rowNum: the trade's sheet row.
'
' STATE OWNERSHIP
'   Adds to mBk through AddToBucket, to mNS and to mTradesUsed.
'
' UPDATED
'   2026-10-14
'==============================================================================
'

'------------------------------------------------------------------------------
' MATURITY FACTOR AND EFFECTIVE NOTIONAL
'------------------------------------------------------------------------------
    'The effective notional is kept on both bases: unmargined MF from the
    'trade's maturity, and margined MF from the netting set's MPOR. The
    'margined basis equals the unmargined one for an unmargined set.
        If t.ok Then
            t.sfEff = mSF(t.sfIdx).SF * t.natFactor
            t.corrEff = mSF(t.sfIdx).Corr
            t.mfU = FormulaNumber(SACCR_MaturityFactor(t.M, False, 0#, pMinMatBD, pBDYear), "SACCR_MaturityFactor")
            If mNS(t.nsIdx).Margined Then
                t.mfM = mNS(t.nsIdx).MFMargined
            Else
                t.mfM = t.mfU
            End If
            t.enU = t.delta * t.adjN * t.mfU
            t.enM = t.delta * t.adjN * t.mfM
            If t.ac = AC_CR Or t.ac = AC_EQ Or t.ac = AC_CO Then
                CheckReferenceClass t.nsIdx, t.ac, t.rf, t.sfKey, t.tid, rowNum
            End If
            AddToBucket t.nsIdx, t.ac, t.hsKey, t.subKey, t.sfEff, t.corrEff, t.enU, t.enM, t.tid, rowNum
            mNS(t.nsIdx).V = mNS(t.nsIdx).V + t.mtmRep
            mNS(t.nsIdx).Trades = mNS(t.nsIdx).Trades + 1
            mTradesUsed = mTradesUsed + 1
        End If

End Sub


Private Sub TradeCalcRow( _
    ByRef t As tTrade, _
    ByRef outArr() As Variant, _
    ByVal nOut As Long)
'
'==============================================================================
'                                 TradeCalcRow
'------------------------------------------------------------------------------
' PURPOSE
'   Fill row nOut of the TradeCalc block with a trade's results, or with its
'   rejection reasons (#60).
'
' INPUTS
'   t: the trade being processed, from ProcessTrades.
'   outArr, nOut: the TradeCalc block and the row to fill.
'
' UPDATED
'   2026-10-14
'==============================================================================
'

'------------------------------------------------------------------------------
' TRADECALC ROW
'------------------------------------------------------------------------------
    'Columns: 1 ID, 2 netting set, 3 asset class, 4 status, 5 hedging set,
    '6 bucket, 7 effective SF, 8 notional, 9 MtM, 10 S, 11 E, 12 M, 13 T,
    '14 SD, 15 d, 16 delta, 17 MF unmargined, 18 MF margined, 19 effective
    'notional unmargined, 20 effective notional margined, 21 IR bucket,
    '22 stand-alone add-on, 23 lambda, 24 regime, 25 messages. Columns that
    'do not apply to the trade stay blank.
        outArr(nOut, 1) = t.tid
        outArr(nOut, 2) = t.nsId
        outArr(nOut, 3) = t.acTxt
        If t.ok Then
            outArr(nOut, 4) = "OK"
            outArr(nOut, 5) = t.hsKey
            outArr(nOut, 6) = t.subKey
            outArr(nOut, 7) = t.sfEff
            outArr(nOut, 8) = t.notionalRep
            outArr(nOut, 9) = t.mtmRep
            outArr(nOut, 10) = t.S
            outArr(nOut, 11) = t.E
            outArr(nOut, 12) = t.M
            If t.instType = "OPTION" Then outArr(nOut, 13) = t.T
            If t.ac = AC_IR Or t.ac = AC_CR Then outArr(nOut, 14) = t.sd
            outArr(nOut, 15) = t.adjN
            outArr(nOut, 16) = t.delta
            outArr(nOut, 17) = t.mfU
            If mNS(t.nsIdx).Margined Then outArr(nOut, 18) = t.mfM
            outArr(nOut, 19) = t.enU
            If mNS(t.nsIdx).Margined Then outArr(nOut, 20) = t.enM
            If t.ac = AC_IR Then outArr(nOut, 21) = t.bucketNo
            outArr(nOut, 22) = t.sfEff * Abs(t.enU)
            outArr(nOut, 23) = t.lamOut
            outArr(nOut, 24) = mNS(t.nsIdx).Regime
            outArr(nOut, 25) = t.warn
        Else
            outArr(nOut, 4) = "EXCLUDED"
            If t.nsIdx > 0 Then
                outArr(nOut, 24) = mNS(t.nsIdx).Regime
            End If
            outArr(nOut, 25) = t.msg
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


Private Function MarketPair( _
    ByVal pair As String) _
    As String
'
'==============================================================================
'                                  MarketPair
'------------------------------------------------------------------------------
' PURPOSE
'   Write a currency pair "AAA/BBB" in market quote order: the currency
'   listed earlier in FX_QUOTE_ORDER first; a currency not listed after
'   every listed one; of two unlisted currencies, the alphabetical first
'   (#40).
'
' UPDATED
'   2026-10-13
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim c1   As String    'First currency of pair
    Dim c2   As String    'Second currency of pair
    Dim r1   As Long      'Rank of c1
    Dim r2   As Long      'Rank of c2

'------------------------------------------------------------------------------
' ORDER
'------------------------------------------------------------------------------
        c1 = Left$(pair, 3)
        c2 = Right$(pair, 3)
        r1 = QuoteRank(c1)
        r2 = QuoteRank(c2)
        If r1 < r2 Or (r1 = r2 And c1 < c2) Then
            MarketPair = c1 & "/" & c2
        Else
            MarketPair = c2 & "/" & c1
        End If

End Function


Private Function MarketSign( _
    ByVal pair As String) _
    As Double
'
'==============================================================================
'                                  MarketSign
'------------------------------------------------------------------------------
' PURPOSE
'   +1 when a pair written in alphabetical order (NormalizePair) is also
'   in market quote order, -1 when the market order is the reverse
'   (JPY/USD is quoted USD/JPY). Compared with the sign NormalizePair
'   gives, it tells whether a pair was entered in market order (#40).
'
' UPDATED
'   2026-10-13
'==============================================================================
'
        If MarketPair(pair) = pair Then
            MarketSign = 1#
        Else
            MarketSign = -1#
        End If

End Function


Private Function QuoteRank( _
    ByVal ccy As String) _
    As Long
'
'==============================================================================
'                                  QuoteRank
'------------------------------------------------------------------------------
' PURPOSE
'   Rank of a currency in FX_QUOTE_ORDER, lower first; a currency not
'   listed ranks after every listed one (#40).
'
' UPDATED
'   2026-10-13
'==============================================================================
'
        QuoteRank = InStr(1, " " & FX_QUOTE_ORDER & " ", " " & ccy & " ", vbBinaryCompare)
        If QuoteRank = 0 Then
            QuoteRank = Len(FX_QUOTE_ORDER) + 2
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
                        mHS(h).ENU = FormulaNumber(SACCR_IREffectiveNotional(mHS(h).D1U, mHS(h).D2U, mHS(h).D3U, _
                                     pRho12, pRho23, pRho13), "SACCR_IREffectiveNotional")
                        mHS(h).ENM = FormulaNumber(SACCR_IREffectiveNotional(mHS(h).D1M, mHS(h).D2M, mHS(h).D3M, _
                                     pRho12, pRho23, pRho13), "SACCR_IREffectiveNotional")
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
'   Adds each VALID EAD to the total of its regime (mRegEAD) and counts
'   the netting sets of each regime, VALID and withheld; counts INCOMPLETE
'   and INVALID sets in mIncomplete. Only VALID sets enter the totals, and
'   each regime has its own TOTAL row: EADs under CRR and BCBS are never
'   added together (#198).
'
' REFERENCE
'   EAD CRE52.1; cap CRE52.2 and CRR Art. 274(3); RC CRE52.10, CRE52.18 and
'   CRR Art. 275(1); multiplier CRE52.23; collateral of the cap, NICA, per
'   CRE52.10 and EBA-QA-2023_6962; trade-count warning CRE52.51(1) and
'   CRR Art. 285(3)(a).
'
' UPDATED
'   2026-10-13
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
    Dim ri       As Long         'Regime of the netting set: 1 CRR, 2 BCBS
    Dim firstTotal As Long       'Results row of the first TOTAL row
    Dim outTop   As Long         'First Results data row
    Dim tot(1 To 2, 1 To RS_NCOLS) As Double    'Column totals of each regime's TOTAL row

'------------------------------------------------------------------------------
' CALCULATE EACH NETTING SET
'------------------------------------------------------------------------------
        ReDim outArr(1 To mNSCount + 2, 1 To RS_NCOLS)
        nOut = 0
        For k = 1 To mNSCount
            ri = RegimeIndex(mNS(k).Regime)
            mRegSets(ri) = mRegSets(ri) + 1

    'More than 5,000 trades at any point in a quarter sets the 20-day MPOR
    'floor for the following quarter [CRE52.51(1); CRR Art. 285(3)(a)].
    'The workbook sees only today's count, so it warns when the declared
    'flag may be wrong instead of setting it.
            With mNS(k)
                If .Margined And Not .LargeOrIlliquid And .EconTrades > LARGE_NS_TRADES Then
                    LogMsg SEV_WARN, SH_NS, .ID, "Netting set has " & .EconTrades & _
                           " trades, more than " & LARGE_NS_TRADES & ": if it had more than " & _
                           LARGE_NS_TRADES & " during the previous quarter, set the large or illiquid " & _
                           "flag to Y (20-day MPOR floor)."
                End If
            End With
            If mNS(k).InputErrors > 0 Then
                mIncomplete = mIncomplete + 1
                mRegWithheld(ri) = mRegWithheld(ri) + 1
                LogMsg SEV_ERROR, SH_NS, mNS(k).ID, "EAD withheld: " & mNS(k).InputErrors & _
                       " invalid netting-set field(s)."
                nOut = nOut + 1
                WriteStatusRow outArr, nOut, k, "INVALID: " & mNS(k).InputErrors & " input error(s)"
                GoTo NextNS
            End If
            If mNS(k).Rejected > 0 Then
                mIncomplete = mIncomplete + 1
                mRegWithheld(ri) = mRegWithheld(ri) + 1
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
    'CRE52.10; CRR Art. 274(3), 275(1); EBA-QA-2023_6962]. An unmargined
    'set uses C: NICA, plus VM posted under a one-way agreement under BCBS.
                If .Margined Then
                    cCap = .NICA
                Else
                    cCap = C
                End If
                rcU = FormulaNumber(SACCR_ReplacementCost(.V, cCap, False), "SACCR_ReplacementCost")
                mU = FormulaNumber(SACCR_Multiplier(.V - cCap, addU, pFloor), "SACCR_Multiplier")
                pfeU = mU * addU
                eadU = .Alpha * (rcU + pfeU)

    'Margined basis, then the cap: the EAD is the lower of the two.
                If .Margined Then
                    rcM = FormulaNumber(SACCR_ReplacementCost(.V, C, True, .TH, .MTA, .NICA), "SACCR_ReplacementCost")
                    mM = FormulaNumber(SACCR_Multiplier(.V - C, addM, pFloor), "SACCR_Multiplier")
                    pfeM = mM * addM
                    eadM = .Alpha * (rcM + pfeM)
                    ead = Min2(eadM, eadU)
                Else
                    ead = eadU
                End If
                mRegEAD(ri) = mRegEAD(ri) + ead
                mRegValid(ri) = mRegValid(ri) + 1

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

    'Running totals of the regime for V, C, RC, the add-ons, PFE and EAD.
                tot(ri, 7) = tot(ri, 7) + .V
                tot(ri, 8) = tot(ri, 8) + C
                For a = 9 To 16
                    tot(ri, a) = tot(ri, a) + ToDbl(outArr(nOut, a))
                Next a
                tot(ri, 18) = tot(ri, 18) + ToDbl(outArr(nOut, 18))
                tot(ri, 22) = tot(ri, 22) + ead
            End With
NextNS:
        Next k

'------------------------------------------------------------------------------
' WRITE RESULTS
'------------------------------------------------------------------------------
    'One TOTAL row per regime with netting sets, CRR first: TOTAL CRR and
    'TOTAL BCBS (#198). Each sums, over the regime's VALID netting sets,
    'columns 7 to 18 except the multiplier (17), and the EAD (22); its
    'status column says how many netting sets it covers and how many were
    'withheld. There is no total across regimes. The rows are in bold.
        If writeOutputs Then
            Set ws = GetSheet(SH_RESULTS)
            outTop = OutputFirstRow(SH_RESULTS)
            ClearOutputBlock ws, outTop, RS_NCOLS
            If nOut > 0 Then
                firstTotal = nOut + 1
                For ri = 1 To 2
                    If mRegSets(ri) > 0 Then
                        nOut = nOut + 1
                        outArr(nOut, 1) = "TOTAL " & Choose(ri, RG_CRR, RG_BCBS)
                        outArr(nOut, 3) = Choose(ri, RG_CRR, RG_BCBS)
                        For a = 7 To 18
                            If a <> 17 Then
                                outArr(nOut, a) = tot(ri, a)
                            End If
                        Next a
                        outArr(nOut, 22) = tot(ri, 22)
                        outArr(nOut, RS_STATUS_COL) = TotalStatusText(Choose(ri, RG_CRR, RG_BCBS))
                    End If
                Next ri
                WriteBlock ws, outTop, outArr, nOut, RS_NCOLS
                FormatColumns ws, outTop, nOut, _
                    "@|@|@|@|0|0|#,##0|#,##0|#,##0|#,##0|#,##0|#,##0|#,##0|#,##0|#,##0|#,##0|0.0000|#,##0|0.00|#,##0|#,##0|#,##0|@|#,##0|#,##0|#,##0|0.0000|@"
                TableRange(ws, outTop + firstTotal - 1, 1, outTop + nOut - 1, RS_NCOLS).Font.Bold = True
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
        ClearOutputBlock GetSheet(SH_TRADECALC), OutputFirstRow(SH_TRADECALC), TC_NCOLS
        ClearOutputBlock GetSheet(SH_BUCKETS), OutputFirstRow(SH_BUCKETS), BK_NCOLS
        ClearOutputBlock GetSheet(SH_HEDGING), OutputFirstRow(SH_HEDGING), HS_NCOLS
        ClearOutputBlock GetSheet(SH_RESULTS), OutputFirstRow(SH_RESULTS), RS_NCOLS

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
'   Fingerprint everything a run reads: the workbook's date system, the
'   NettingSets and Trades headers (HeaderText) and input rows (without
'   the unread Comment column) and the Params inputs, located through the
'   workbook names wherever the layout has moved them (ParamsText) (#36,
'   #152). Each cell is written with its type (CellToken), so a number
'   and the same digits as text, or two error values, differ (#197).
'
' RETURNS
'   An 8-digit hexadecimal fingerprint (CORE_Util.TextHash). Equal
'   fingerprints mean the inputs are, all but certainly, unchanged.
'   FP_FORMAT opens the text, so a fingerprint stored before the format
'   changed reads as out of date rather than as current.
'
' ERROR POLICY
'   Raises ERR_PARAMS_LAYOUT when a Params name is missing, broken or not
'   on Params: a fingerprint of fixed columns could miss an edit, so none
'   is given.
'
' UPDATED
'   2026-10-13
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim parts(1 To 6)   As String    'Text of each input block

'------------------------------------------------------------------------------
' FINGERPRINT
'------------------------------------------------------------------------------
    'The run stops in a 1904 workbook (#185), and a renamed header stops
    'it too (#35), so both are part of what the results depend on.
        parts(1) = FP_FORMAT & "|1904=" & IIf(ThisWorkbook.Date1904, "1", "0")
        parts(2) = HeaderText(SH_NS, NS_NCOLS)
        parts(3) = BlockText(GetSheet(SH_NS), TableFirstRow(SH_NS), NS_NCOLS)
        parts(4) = HeaderText(SH_TRADES, TR_COMMENT - 1)
        parts(5) = BlockText(GetSheet(SH_TRADES), TableFirstRow(SH_TRADES), TR_COMMENT - 1)
        parts(6) = ParamsText()
        InputFingerprint = TextHash(Join(parts, Chr$(29)))

End Function


Private Function HeaderText( _
    ByVal sheetName As String, _
    ByVal nCols As Long) _
    As String
'
'==============================================================================
'                                  HeaderText
'------------------------------------------------------------------------------
' PURPOSE
'   Write the header row of the NettingSets or Trades block as one text,
'   each header as HeadersMatch compares it: without spaces, tabs and line
'   breaks (HeaderKey), in upper case. A header changed only in those ways
'   passes the schema check, so it must not make the results out of date;
'   any other change stops the run, so it must (#197). A header that is
'   not text is written as CellToken writes it.
'
' RETURNS
'   "" when the header row cannot be located; the run then stops on the
'   same broken name.
'
' UPDATED
'   2026-10-13
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ws        As Worksheet    'Input sheet
    Dim r         As Long         'Header row; 0 when it cannot be located
    Dim pieces()  As String       'One text per header
    Dim c         As Long         'Table column
    Dim v         As Variant      'Header cell value

'------------------------------------------------------------------------------
' WRITE
'------------------------------------------------------------------------------
        r = TableHeaderRow(sheetName)
        If r = 0 Then
            Exit Function
        End If
        Set ws = GetSheet(sheetName)
        ReDim pieces(1 To nCols)
        For c = 1 To nCols
            v = TableCell(ws, r, c).Value2
            If VarType(v) = vbString Then
                pieces(c) = "T" & UCase$(HeaderKey(CStr(v)))
            Else
                pieces(c) = CellToken(v)
            End If
        Next c
        HeaderText = Join(pieces, Chr$(31))

End Function


Private Function ParamsText() As String
'
'==============================================================================
'                                  ParamsText
'------------------------------------------------------------------------------
' PURPOSE
'   Write the Params inputs of a run as one text, located through the
'   workbook names as the run locates them: the value of each parameter's
'   named cell, the "Code" and "Value" columns (ValidateSchema reads them
'   for repeated codes), and the factor and FX tables from their header
'   rows (ParamsLayoutValid checks the headers) to their last row with a
'   value, or to the row above the other table when that follows
'   (LoadSFTable, LoadFXTable). Other cells, such as descriptions and
'   references, are left out.
'
' ERROR POLICY
'   Raises ERR_PARAMS_LAYOUT when a name is missing or broken, or when one
'   of the four layout names is not on Params.
'
' UPDATED
'   2026-10-13
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ws         As Worksheet    'Params sheet
    Dim nameList   As Variant      'Layout names, then parameter codes
    Dim i          As Long         'Index into nameList
    Dim c          As Range        'A name's top-left cell
    Dim pieces()   As String       'Parameters, then the blocks
    Dim k          As Long         'Position in pieces
    Dim sfRow      As Long         'Header row of the factor table
    Dim fxRow      As Long         'Header row of the FX table
    Dim sfLast     As Long         'Last row of the factor block; 0 = last used
    Dim fxLast     As Long         'Last row of the FX block; 0 = last used

'------------------------------------------------------------------------------
' CHECK THE LAYOUT NAMES
'------------------------------------------------------------------------------
        Set ws = GetSheet(SH_PARAMS)
        nameList = Array(NM_PRM_CODE, NM_PRM_VALUE, NM_SF_TABLE, NM_FX_TABLE)
        For i = LBound(nameList) To UBound(nameList)
            Set c = NamedCell(CStr(nameList(i)))
            If c Is Nothing Then
                Err.Raise ERR_PARAMS_LAYOUT, "CORE_Engine.InputFingerprint", "Workbook name '" & _
                          nameList(i) & "' is missing or no longer refers to a cell."
            ElseIf c.Worksheet.Name <> ws.Name Then
                Err.Raise ERR_PARAMS_LAYOUT, "CORE_Engine.InputFingerprint", "Workbook name '" & _
                          nameList(i) & "' must refer to a cell on " & SH_PARAMS & "."
            End If
        Next i

'------------------------------------------------------------------------------
' PARAMETERS
'------------------------------------------------------------------------------
    'Code=value of each named cell, in the order of ParameterCodes.
        nameList = ParameterCodes()
        ReDim pieces(1 To UBound(nameList) - LBound(nameList) + 5)
        For i = LBound(nameList) To UBound(nameList)
            Set c = NamedCell(CStr(nameList(i)))
            If c Is Nothing Then
                Err.Raise ERR_PARAMS_LAYOUT, "CORE_Engine.InputFingerprint", "Workbook name '" & _
                          nameList(i) & "' is missing or no longer refers to a cell."
            End If
            k = k + 1
            pieces(k) = nameList(i) & "=" & CellToken(c.Value2)
        Next i

'------------------------------------------------------------------------------
' COLUMNS AND TABLES
'------------------------------------------------------------------------------
    'A table starts at its header row, whose headers ParamsLayoutValid
    'checks, and stops above the other table when that follows it, as the
    'loaders stop; otherwise at its last row with a value.
        sfRow = NamedRow(NM_SF_TABLE)
        fxRow = NamedRow(NM_FX_TABLE)
        If fxRow > sfRow Then
            sfLast = fxRow - 1
        ElseIf sfRow > fxRow Then
            fxLast = sfRow - 1
        End If
        pieces(k + 1) = BlockText(ws, 1, 1, ParamsCodeCol())
        pieces(k + 2) = BlockText(ws, 1, 1, ParamsValueCol())
        pieces(k + 3) = BlockText(ws, sfRow, UBound(Split(SF_HEADERS, "|")) + 1, _
                                  NamedTableCol(NM_SF_TABLE), sfLast)
        pieces(k + 4) = BlockText(ws, fxRow, UBound(Split(FX_HEADERS, "|")) + 1, _
                                  NamedTableCol(NM_FX_TABLE), fxLast)
        ParamsText = Join(pieces, Chr$(30))

End Function


Public Function BlockText( _
    ByVal ws As Worksheet, _
    ByVal firstRow As Long, _
    ByVal nCols As Long, _
    Optional ByVal firstCol As Long = 1, _
    Optional ByVal lastRow As Long = 0) _
    As String
'
'==============================================================================
'                                  BlockText
'------------------------------------------------------------------------------
' PURPOSE
'   Write a block of cells as one text, cell by cell (CellToken), from
'   firstRow to the last row with a value. Blank rows below it are left
'   out, so that a used range that grew without new values does not
'   change the text. Public for M_Main, which fingerprints the bank
'   figures with it (#45).
'
' INPUTS
'   firstCol: the block's first table column, as TableCell counts it.
'   lastRow: the last row to read; 0 for the last used row.
'
' UPDATED
'   2026-10-15
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim lastR    As Long        'Last row read
    Dim block    As Variant     'Cell values, Value2 (dates as serials)
    Dim data     As Variant     'The same as a 2-D array, also for one cell
    Dim nRows    As Long        'Rows up to the last with a value
    Dim pieces() As String      'One text per cell
    Dim r        As Long        'Row of data
    Dim c        As Long        'Column of data
    Dim k        As Long        'Position in pieces

'------------------------------------------------------------------------------
' READ
'------------------------------------------------------------------------------
        If firstRow < 1 Then
            Exit Function
        End If
        lastR = UsedLastRow(ws)
        If lastRow > 0 And lastRow < lastR Then
            lastR = lastRow
        End If
        If lastR < firstRow Then
            Exit Function
        End If
        block = TableRange(ws, firstRow, firstCol, lastR, firstCol + nCols - 1).Value2
        If IsArray(block) Then
            data = block
        Else
            ReDim data(1 To 1, 1 To 1)
            data(1, 1) = block
        End If
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
                pieces(k) = CellToken(data(r, c))
            Next c
        Next r
        BlockText = Join(pieces, Chr$(31))

End Function


Private Function CellToken( _
    ByVal v As Variant) _
    As String
'
'==============================================================================
'                                  CellToken
'------------------------------------------------------------------------------
' PURPOSE
'   Write one cell value (Value2) for the fingerprint with a tag for its
'   type: N for a number, T for text, B for a logical value, E for an
'   error value with its code, and nothing for an empty cell. The number
'   47756 and the text '47756 are read differently since #185, so they
'   must not give the same text (#197). Formatting does not reach Value2,
'   so a formatting-only edit leaves the token unchanged.
'
' UPDATED
'   2026-10-13
'==============================================================================
'
        Select Case VarType(v)
            Case vbEmpty
                CellToken = ""
            Case vbString
                CellToken = "T" & v
            Case vbBoolean
                CellToken = IIf(v, "B1", "B0")
            Case vbError
                CellToken = "E" & CStr(v)
            Case Else
                CellToken = "N" & CStr(v)
        End Select

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
'   sheetName, nCols: output sheet and table width, from the first table
'   column.
' RETURNS
'   True only when this output was cleared. Appends any failure to
'   failureDetails, preserving details from earlier sheets.
' ERROR POLICY
'   Contains this cleanup failure; the original run error remains primary.
' UPDATED
'   2026-10-11
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
        ClearOutputBlock GetSheet(sheetName), OutputFirstRow(sheetName), nCols
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
    Dim topRow  As Long         'First Buckets data row
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
        topRow = OutputFirstRow(SH_BUCKETS)
        ClearOutputBlock ws, topRow, BK_NCOLS
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
        WriteBlock ws, topRow, arr, mBkCount, BK_NCOLS
        FormatColumns ws, topRow, mBkCount, "@|@|@|@|0|0.00%|0%|#,##0|#,##0|#,##0|#,##0"

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
    Dim topRow  As Long         'First HedgingSets data row
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
        topRow = OutputFirstRow(SH_HEDGING)
        ClearOutputBlock ws, topRow, HS_NCOLS
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
        WriteBlock ws, topRow, arr, mHSCount, HS_NCOLS
        FormatColumns ws, topRow, mHSCount, "@|@|@|0|#,##0|#,##0|#,##0|#,##0|#,##0|#,##0|#,##0|#,##0|@|0.00%"

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
'   2026-10-11
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ws      As Worksheet    'Checks sheet
    Dim topRow  As Long         'First Checks data row
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
        topRow = OutputFirstRow(SH_CHECKS)
        ClearOutputBlock ws, topRow, CK_NCOLS
        If mLogCount = 0 Then
            ReDim arr(1 To 1, 1 To CK_NCOLS)
            arr(1, 1) = SEV_INFO
            arr(1, 5) = "No issues found."
            WriteBlock ws, topRow, arr, 1, CK_NCOLS
            WriteChecks = True
            Exit Function
        End If
        ReDim arr(1 To mLogCount, 1 To CK_NCOLS)
        For i = 1 To mLogCount
            For j = 1 To CK_NCOLS
                arr(i, j) = mLog(j, i)
            Next j
        Next i
        WriteBlock ws, topRow, arr, mLogCount, CK_NCOLS
        For i = 1 To mLogCount
            If arr(i, 1) = SEV_ERROR Then
                TableCell(ws, topRow + i - 1, 1).Font.Bold = True
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
'   Write a one-line run summary (RunInfoText) to cell C4 (RUNINFO_CELL) of
'   the Results sheet.
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
'   Never raises: the summary is informative, and Calculate reports the
'   run's outcome itself, also from its error handler. A summary that
'   cannot be composed is replaced by a line saying so, never left empty
'   or partial. Only the write to C4 runs under On Error Resume Next: a
'   protected or missing Results sheet keeps its earlier summary, which
'   the run error or Checks then explains (#36, #60).
'
' UPDATED
'   2026-10-14
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim txt   As String    'Summary line

'------------------------------------------------------------------------------
' COMPOSE
'------------------------------------------------------------------------------
        On Error GoTo NoText
        If gEngineFault = "runinfo" Then
            Err.Raise ERR_INJECTED_FAULT, "CORE_Engine.WriteRunInfo", "Injected run-summary failure."
        End If
        txt = RunInfoText(secs, completed, wroteOutputs, failure)

'------------------------------------------------------------------------------
' WRITE
'------------------------------------------------------------------------------
WriteText:
        On Error Resume Next
        GetSheet(SH_RESULTS).Range(RUNINFO_CELL).Value = txt
        On Error GoTo 0
        Exit Sub

'------------------------------------------------------------------------------
' HANDLE A SUMMARY THAT CANNOT BE COMPOSED
'------------------------------------------------------------------------------
NoText:
        txt = "Last run " & Format$(Now, "yyyy-mm-dd hh:mm:ss") & " - run summary unavailable (error " & _
              Err.Number & ": " & Err.Description & "); see Checks"
        Resume WriteText

End Sub


Private Function RunInfoText( _
    ByVal secs As Double, _
    ByVal completed As Boolean, _
    ByVal wroteOutputs As Boolean, _
    ByVal failure As String) _
    As String
'
'==============================================================================
'                                 RunInfoText
'------------------------------------------------------------------------------
' PURPOSE
'   Compose the run summary written by WriteRunInfo: a failure, a
'   validation-only run, a completed run with its counts, run class, input
'   fingerprint and duration, or a run stopped by input errors.
'
' INPUTS
'   As WriteRunInfo.
'
' ERROR POLICY
'   Raises any error to WriteRunInfo, which writes a fallback line.
'
' UPDATED
'   2026-10-14
'==============================================================================
'
        If Len(failure) > 0 Then
            RunInfoText = "Last run " & Format$(Now, "yyyy-mm-dd hh:mm:ss") & " FAILED - " & failure
        ElseIf Not wroteOutputs Then
            RunInfoText = "Last validation " & Format$(Now, "yyyy-mm-dd hh:mm:ss") & _
                          " | errors " & mErrCount & ", warnings " & mWarnCount & _
                          RunClassText() & _
                          " | outputs cleared: press Run SA-CCR to calculate"
        ElseIf completed Then
            RunInfoText = "Last run " & Format$(Now, "yyyy-mm-dd hh:mm:ss") & _
                          " | reporting date " & Format$(CDate(pAsOf), "yyyy-mm-dd") & _
                          " | ccy " & pRepCcy & _
                          " | economic trades used " & mEconUsed & " of " & mEconRead & _
                          ", legs " & mTradesUsed & " of " & mTradesRead & _
                          " | errors " & mErrCount & ", warnings " & mWarnCount & _
                          RunClassText() & _
                          " | inputs " & mRunInputs & _
                          " | " & CORE_Export.NumberText(Round(secs, 2)) & " s"
            If mIncomplete > 0 Then
                RunInfoText = RunInfoText & " | EAD withheld for " & mIncomplete & " netting set(s): " & _
                              WithheldByRegimeText()
            End If
        Else
            RunInfoText = "Last run " & Format$(Now, "yyyy-mm-dd hh:mm:ss") & _
                          " FAILED - see Checks sheet (" & mErrCount & " errors)"
        End If

End Function


Private Function RunClassText() As String
'
'==============================================================================
'                                 RunClassText
'------------------------------------------------------------------------------
' PURPOSE
'   The run class for the run summary (#41): the regulatory set's identity,
'   or WHAT-IF with the number of parameters that differ from it; "" when
'   the run was not classified.
'
' UPDATED
'   2026-10-11
'==============================================================================
'
        If mRunClass = RUN_REGULATORY Then
            RunClassText = " | regulatory parameters " & REF_SET_ID
        ElseIf mRunClass = RUN_WHATIF Then
            RunClassText = " | WHAT-IF: " & mDeviations & " parameter(s) differ from " & REF_SET_ID
        End If

End Function
