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
'   RegimeWithheldCount, TotalStatusText and WithheldByRegimeText), the
'   run-fingerprint procedures (InputFingerprint, LastRunInputs,
'   ForgetRunInputs, and BlockText, which M_Main also uses for the
'   reconciliation fingerprint) and ParameterCodes, which M_Main uses for
'   the evidence report, are Public for M_Main and the tests; the test seam
'   gEngineFault is Public for the tests. Option Private Module keeps them
'   off the supported external surface.
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
'   2026-10-20
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
        ClearingRole      As String     'Client-clearing role: N, CM, CLIENT, CCP or MULTI
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
        MarginedSTM       As Boolean    'Margined, with a settled-to-market trade: INVALID (#360)
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
        stm           As Boolean    'Settled-to-market: settlement type STM (#360)
        dReset        As Double     'Next reset date; -1 when blank
        S             As Double     'Start S, years, at least 0
        E             As Double     'End E, years
        M             As Double     'Maturity M, years; for STM the time to the next reset date
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
    ByVal WriteOutputs As Boolean) _
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
'   WriteOutputs: True writes every output sheet; False validates only and
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
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim StartTime           As Double    'Timer value at the start, for the run duration
    Dim ChecksFailure       As String    'Why the Checks sheet could not be written; "" if written
    Dim ErrNumber           As Long      'Unexpected error: number
    Dim ErrSource           As String    'Unexpected error: source
    Dim WithdrawalDetails   As String    'Failures by sheet, distinct from the primary error
    Dim WithdrawalStatus    As String    'Whether every output could be withdrawn
    Dim ErrDescription      As String    'Unexpected error: description

'------------------------------------------------------------------------------
' LOAD INPUTS
'------------------------------------------------------------------------------
    'Old outputs are cleared first, for validation too, so that a failed or
    'validation-only run can never leave earlier results looking current
    '(#36). Any loader that fails has logged why; skip to the Checks output.
        StartTime = Timer
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
        ProcessTrades WriteOutputs
        ComputeHedgingSets
        ComputeNettingSets WriteOutputs
        If WriteOutputs Then
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
        If Not WriteChecks(ChecksFailure) Then
            Calculate = False
            WithdrawalStatus = "results withdrawn"
            If Not WithdrawOutputs(WithdrawalDetails) Then
                WithdrawalStatus = "output cleanup incomplete; old results may remain - unprotect or repair outputs and rerun. " & _
                                   WithdrawalDetails
            End If
            ForgetRunInputs
            WriteRunInfo Timer - StartTime, False, WriteOutputs, "the Checks sheet could not be written (" & _
                         ChecksFailure & "); " & WithdrawalStatus & ", and Checks may show an earlier run"
            On Error GoTo 0
            Err.Raise ERR_CHECKS_WRITE, "CORE_Engine.Calculate", "The Checks sheet could not be written (" & _
                      ChecksFailure & "); " & WithdrawalStatus & ". Unprotect or repair the Checks " & _
                      "sheet and run again."
        End If

    'Only a completed run that wrote its outputs leaves results; remember
    'its inputs so that later edits show the results as out of date.
        If Calculate And WriteOutputs Then
            mRunInputs = InputFingerprint()
            RememberRunInputs mRunInputs
        Else
            ForgetRunInputs
        End If
        WriteRunInfo Timer - StartTime, Calculate, WriteOutputs, ""
        Exit Function

'------------------------------------------------------------------------------
' HANDLE UNEXPECTED ERROR
'------------------------------------------------------------------------------
    'Attempt withdrawal of everything written so far, and log the original
    'error on Checks and in the run summary, both best effort. Preserve its
    'number/source/text and append any secondary failure for M_Main to report.
Failed:
        ErrNumber = Err.Number
        ErrSource = Err.Source
        ErrDescription = Err.Description
        Calculate = False
        mRunInputs = ""
    'With the fingerprint forgotten the results read NONE (no valid run),
    'although protected output cells may still hold old values.
        ForgetRunInputs
        WithdrawalStatus = "results withdrawn"
        If Not WithdrawOutputs(WithdrawalDetails) Then
            WithdrawalStatus = "output cleanup incomplete; old results may remain - unprotect or repair outputs and rerun. " & _
                               WithdrawalDetails
        End If
        LogMsg SEV_ERROR, "", "", "Run stopped by error " & ErrNumber & ": " & ErrDescription & _
               " - " & WithdrawalStatus & "."
        If Not WriteChecks(ChecksFailure) Then
            WithdrawalStatus = WithdrawalStatus & "; Checks could not be updated (" & ChecksFailure & _
                               "); Checks may show an earlier run."
        End If
        WriteRunInfo Timer - StartTime, False, WriteOutputs, "error " & ErrNumber & ": " & ErrDescription & _
                     "; " & WithdrawalStatus
        If Len(WithdrawalDetails) > 0 Or Len(ChecksFailure) > 0 Then
            ErrDescription = ErrDescription & " " & WithdrawalStatus & "."
        End If
        Err.Raise ErrNumber, ErrSource, ErrDescription

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
    ByVal RegimeCode As String) _
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
'   RegimeCode: RG_CRR or RG_BCBS.
'
' UPDATED
'   2026-10-17
'==============================================================================
'
        TotalEAD = mRegEAD(RegimeIndex(RegimeCode))

End Property


Public Property Get RegimeSetCount( _
    ByVal RegimeCode As String) _
    As Long
'
'==============================================================================
'                                RegimeSetCount
'------------------------------------------------------------------------------
' PURPOSE
'   Netting sets of one regime in the last run, whatever their status.
'
' UPDATED
'   2026-10-17
'==============================================================================
'
        RegimeSetCount = mRegSets(RegimeIndex(RegimeCode))

End Property


Public Property Get RegimeValidCount( _
    ByVal RegimeCode As String) _
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
'   2026-10-17
'==============================================================================
'
        RegimeValidCount = mRegValid(RegimeIndex(RegimeCode))

End Property


Public Property Get RegimeWithheldCount( _
    ByVal RegimeCode As String) _
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
'   2026-10-17
'==============================================================================
'
        RegimeWithheldCount = mRegWithheld(RegimeIndex(RegimeCode))

End Property


Public Function TotalStatusText( _
    ByVal RegimeCode As String) _
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
'   2026-10-17
'==============================================================================
'
        TotalStatusText = RegimeValidCount(RegimeCode) & " of " & RegimeSetCount(RegimeCode) & _
                          " netting set(s) VALID; " & RegimeWithheldCount(RegimeCode) & " withheld"

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
    ByVal RegimeCode As String) _
    As Long
'
'==============================================================================
'                                 RegimeIndex
'------------------------------------------------------------------------------
' PURPOSE
'   Position of a regime in the per-regime totals: 1 CRR, 2 BCBS.
'
' UPDATED
'   2026-10-17
'==============================================================================
'
        If RegimeCode = RG_BCBS Then
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
    ByVal Severity As String, _
    ByVal SheetName As String, _
    ByVal Reference As String, _
    ByVal Message As String, _
    Optional ByVal RowNumber As Long = 0)
'
'==============================================================================
'                                    LogMsg
'------------------------------------------------------------------------------
' PURPOSE
'   Add one message for the Checks sheet and count errors and warnings.
'
' INPUTS
'   Severity: SEV_ERROR, SEV_WARN or SEV_INFO.
'   SheetName: sheet the message refers to.
'   Reference: the ID or parameter code concerned; may be empty.
'   Message: the message text.
'   RowNumber: sheet row concerned; 0 leaves the Row column blank.
'
' UPDATED
'   2026-10-17
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
        mLog(1, mLogCount) = Severity
        mLog(2, mLogCount) = SheetName
        If RowNumber > 0 Then
            mLog(3, mLogCount) = RowNumber
        Else
            mLog(3, mLogCount) = Empty
        End If
        mLog(4, mLogCount) = Reference
        mLog(5, mLogCount) = Message

'------------------------------------------------------------------------------
' COUNT
'------------------------------------------------------------------------------
        If Severity = SEV_ERROR Then
            mErrCount = mErrCount + 1
        End If
        If Severity = SEV_WARN Then
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
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ParamsSheet   As Worksheet    'Params sheet
    Dim AllValid      As Boolean      'No problem found so far
    Dim LayoutValid   As Boolean      'Params names and headers are valid
    Dim CodeColumn    As Long         'Table column of the parameter codes
    Dim ValueColumn   As Long         'Table column of the parameter values
    Dim LastRow       As Long         'Last used row of Params
    Dim CodeCells     As Variant      'Parameter codes, rows 1 to LastRow
    Dim ValueCells    As Variant      'Parameter values, rows 1 to LastRow
    Dim Codes         As Variant      'Every parameter code
    Dim i             As Long         'Index into Codes
    Dim r             As Long         'Params row being compared
    Dim FirstRow      As Long         'First Params row holding the code; 0 if none

'------------------------------------------------------------------------------
' INPUT SHEETS
'------------------------------------------------------------------------------
    'A missing output sheet has already raised in ClearOutputSheets.
        AllValid = SheetPresent(SH_PARAMS)
        AllValid = SheetPresent(SH_NS) And AllValid
        AllValid = SheetPresent(SH_TRADES) And AllValid
        If Not AllValid Then
            Exit Function
        End If

'------------------------------------------------------------------------------
' COLUMN HEADERS
'------------------------------------------------------------------------------
    'The input sheets are located through workbook names: a missing or
    'broken name stops the run, and the headers are checked where the names
    'point.
        Set ParamsSheet = GetSheet(SH_PARAMS)
        AllValid = InputHeadersMatch(SH_NS, NM_NS_HEADER, NS_HEADERS) And AllValid
        AllValid = InputHeadersMatch(SH_TRADES, NM_TR_HEADER, TR_HEADERS) And AllValid
        LayoutValid = ParamsLayoutValid(ParamsSheet)
        AllValid = LayoutValid And AllValid
        If Not LayoutValid Then
            Exit Function
        End If

'------------------------------------------------------------------------------
' PARAMETERS
'------------------------------------------------------------------------------
    'Every parameter is read through the workbook name equal to its code
    '(CORE_Util.GetParam), so a missing or broken name stops the run. A
    'code listed twice on Params with the same value is harmless and warned
    'about; with different values it is ambiguous.
        Codes = ParameterCodes()
        CodeColumn = ParamsCodeCol()
        ValueColumn = ParamsValueCol()
        LastRow = UsedLastRow(ParamsSheet)
        If LastRow >= 1 Then
            CodeCells = TableRange(ParamsSheet, 1, CodeColumn, LastRow, CodeColumn).Value2
            ValueCells = TableRange(ParamsSheet, 1, ValueColumn, LastRow, ValueColumn).Value2
        End If
        For i = LBound(Codes) To UBound(Codes)
            If NamedCell(CStr(Codes(i))) Is Nothing Then
                AllValid = False
                LogMsg SEV_ERROR, SH_PARAMS, Codes(i), "Workbook name '" & Codes(i) & _
                       "' is missing or no longer refers to a cell (#REF!); every parameter is read " & _
                       "through its name."
            End If
            FirstRow = 0
            For r = 1 To LastRow
                If StrComp(SafeStr(CodeCells(r, 1)), Codes(i), vbTextCompare) = 0 Then
                    If FirstRow = 0 Then
                        FirstRow = r
                    ElseIf SameValue(ValueCells(FirstRow, 1), ValueCells(r, 1)) Then
                        LogMsg SEV_WARN, SH_PARAMS, Codes(i), "Parameter listed twice with the same value (rows " & _
                               FirstRow & " and " & r & ").", r
                    Else
                        AllValid = False
                        LogMsg SEV_ERROR, SH_PARAMS, Codes(i), "Parameter listed twice with different values (rows " & _
                               FirstRow & " and " & r & ").", r
                    End If
                End If
            Next r
        Next i
        ValidateSchema = AllValid

End Function


Public Function ParameterCodes() As Variant
'
'==============================================================================
' PURPOSE
'   The parameter vocabulary shared by the schema check (ValidateSchema),
'   the table loaders (IsParameterCode), the input fingerprint
'   (ParamsText) and the evidence report of M_Main (#45).
' RETURNS
'   An array of the supported parameter codes.
' UPDATED
'   2026-10-17
'==============================================================================
        ParameterCodes = Array(PRM_ASOF, PRM_REPCCY, PRM_ALPHA, PRM_FLOOR, PRM_DAYSYEAR, PRM_BDYEAR, _
                      PRM_MINMAT, PRM_SDFLOOR, PRM_MPOR_BIL, PRM_MPOR_CLR, PRM_MPOR_LARGE, _
                      PRM_BASIS, PRM_VOLF, PRM_RHO12, PRM_RHO23, PRM_RHO13, PRM_IRFULL, _
                      PRM_REGIME, PRM_LAMIR, PRM_LAMCO)
End Function


Private Function IsParameterCode(ByVal Candidate As String) As Boolean
'
'==============================================================================
' PURPOSE
'   Tell whether Candidate is a supported parameter code, ignoring case, so
'   that ValuesAfterEnd can tell a parameter row below a table from a
'   table row that would not be read.
' UPDATED
'   2026-10-17
'==============================================================================
    Dim Code As Variant    'A supported parameter code
        For Each Code In ParameterCodes()
            If StrComp(Candidate, CStr(Code), vbTextCompare) = 0 Then
                IsParameterCode = True
                Exit Function
            End If
        Next Code
End Function


Private Function SheetPresent( _
    ByVal SheetName As String) _
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
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim CandidateSheet   As Worksheet    'Worksheet being compared

'------------------------------------------------------------------------------
' SEARCH
'------------------------------------------------------------------------------
        For Each CandidateSheet In ThisWorkbook.Worksheets
            If StrComp(CandidateSheet.Name, SheetName, vbTextCompare) = 0 Then
                SheetPresent = True
                Exit Function
            End If
        Next CandidateSheet
        LogMsg SEV_ERROR, SheetName, "", "Sheet '" & SheetName & "' not found."

End Function


Private Function ParamsLayoutValid( _
    ByVal ParamsSheet As Worksheet) _
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
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim LayoutNames   As Variant    'The layout names
    Dim i             As Long       'Index into LayoutNames
    Dim NameCell      As Range      'Top-left cell of a name
    Dim AllValid      As Boolean    'No problem found so far

'------------------------------------------------------------------------------
' NAMES
'------------------------------------------------------------------------------
        AllValid = True
        LayoutNames = Array(NM_PRM_CODE, NM_PRM_VALUE, NM_SF_TABLE, NM_FX_TABLE)
        For i = LBound(LayoutNames) To UBound(LayoutNames)
            Set NameCell = NamedCell(CStr(LayoutNames(i)))
            If NameCell Is Nothing Then
                AllValid = False
                LogMsg SEV_ERROR, SH_PARAMS, CStr(LayoutNames(i)), "Workbook name '" & LayoutNames(i) & _
                       "' is missing or no longer refers to a cell - rebuild the workbook from the template."
            ElseIf NameCell.Worksheet.Name <> ParamsSheet.Name Then
                AllValid = False
                LogMsg SEV_ERROR, SH_PARAMS, CStr(LayoutNames(i)), "Workbook name '" & LayoutNames(i) & _
                       "' must refer to a cell on " & SH_PARAMS & "."
            End If
        Next i
        If Not AllValid Then
            Exit Function
        End If

'------------------------------------------------------------------------------
' HEADERS
'------------------------------------------------------------------------------
        If NamedRow(NM_PRM_CODE) <> NamedRow(NM_PRM_VALUE) Then
            AllValid = False
            LogMsg SEV_ERROR, SH_PARAMS, NM_PRM_VALUE, "Workbook names '" & NM_PRM_CODE & "' and '" & _
                   NM_PRM_VALUE & "' must be on the same row."
        End If
        AllValid = HeadersMatch(ParamsSheet, NamedRow(NM_PRM_CODE), PRM_CODE_HEADER, ParamsCodeCol()) And AllValid
        AllValid = HeadersMatch(ParamsSheet, NamedRow(NM_PRM_VALUE), PRM_VALUE_HEADER, ParamsValueCol()) And AllValid
        AllValid = HeadersMatch(ParamsSheet, NamedRow(NM_SF_TABLE), SF_HEADERS, NamedTableCol(NM_SF_TABLE)) And AllValid
        AllValid = HeadersMatch(ParamsSheet, NamedRow(NM_FX_TABLE), FX_HEADERS, NamedTableCol(NM_FX_TABLE)) And AllValid
        ParamsLayoutValid = AllValid

End Function


Private Function InputHeadersMatch( _
    ByVal SheetName As String, _
    ByVal HeaderName As String, _
    ByVal ExpectedHeaders As String) _
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
'   SheetName: SH_NS or SH_TRADES.
'   HeaderName: NM_NS_HEADER or NM_TR_HEADER.
'   ExpectedHeaders: NS_HEADERS or TR_HEADERS.
'
' RETURNS
'   True when the name locates the header row and every header matches;
'   False, with an error logged, otherwise.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim HeaderRow   As Long    'Header row; 0 when the name cannot locate it

'------------------------------------------------------------------------------
' LOCATE AND COMPARE
'------------------------------------------------------------------------------
        HeaderRow = TableHeaderRow(SheetName)
        If HeaderRow = 0 Then
            LogMsg SEV_ERROR, SheetName, HeaderName, "Workbook name '" & HeaderName & _
                   "' is missing or no longer refers to the first header cell, in column C, of " & _
                   SheetName & " - rebuild the workbook from the template."
            Exit Function
        End If
        InputHeadersMatch = HeadersMatch(GetSheet(SheetName), HeaderRow, ExpectedHeaders)

End Function


Private Function HeadersMatch( _
    ByVal HeaderSheet As Worksheet, _
    ByVal HeaderRow As Long, _
    ByVal ExpectedHeaders As String, _
    Optional ByVal FirstColumn As Long = 1) _
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
'   HeaderSheet, HeaderRow: the sheet and its header row.
'   ExpectedHeaders: one of the _HEADERS constants in CORE_Config; an empty
'   entry is not checked.
'   FirstColumn: the table column of the first header, 1 by default.
'
' RETURNS
'   True when every checked header matches.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Headers       As Variant    'Expected headers, from index 0 for table column 1
    Dim c             As Long       'Index into Headers
    Dim FoundHeader   As Variant    'Header cell value

'------------------------------------------------------------------------------
' COMPARE
'------------------------------------------------------------------------------
        HeadersMatch = True
        Headers = Split(ExpectedHeaders, "|")
        For c = 0 To UBound(Headers)
            If Len(Headers(c)) > 0 Then
                FoundHeader = TableCell(HeaderSheet, HeaderRow, FirstColumn + c).Value
                If StrComp(HeaderKey(SafeStr(FoundHeader)), HeaderKey(CStr(Headers(c))), vbTextCompare) <> 0 Then
                    HeadersMatch = False
                    LogMsg SEV_ERROR, HeaderSheet.Name, TableCell(HeaderSheet, HeaderRow, FirstColumn + c).Address(False, False), _
                           "Header must be '" & Headers(c) & "' (found " & DescribeValue(FoundHeader) & _
                           ") - a column was inserted, deleted or moved, or the header renamed.", HeaderRow
                End If
            End If
        Next c

End Function


Private Function HeaderKey( _
    ByVal RawHeader As String) _
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
'   2026-10-17
'==============================================================================
'
        RawHeader = Replace(RawHeader, " ", "")
        RawHeader = Replace(RawHeader, Chr$(160), "")
        RawHeader = Replace(RawHeader, vbTab, "")
        RawHeader = Replace(RawHeader, vbCr, "")
        HeaderKey = Replace(RawHeader, vbLf, "")

End Function


Private Function SameValue( _
    ByVal FirstValue As Variant, _
    ByVal SecondValue As Variant) _
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
'   2026-10-17
'==============================================================================
'
        If IsError(FirstValue) Or IsError(SecondValue) Then
            SameValue = False
        ElseIf IsNum(FirstValue) And IsNum(SecondValue) Then
            SameValue = (CDbl(FirstValue) = CDbl(SecondValue))
        Else
            SameValue = (StrComp(SafeStr(FirstValue), SafeStr(SecondValue), vbTextCompare) = 0)
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
    ByVal Code As String, _
    ByVal Amount As Double, _
    ByVal Lowest As Double, _
    ByVal Highest As Double, _
    ByVal HighestAllowed As Boolean, _
    ByVal WholeOnly As Boolean)
'
'==============================================================================
'                                ParamInRange
'------------------------------------------------------------------------------
' PURPOSE
'   Check a numeric parameter against its range: at least Lowest, and at
'   most Highest (HighestAllowed) or below it; a whole number when
'   WholeOnly. The ranges are those of the data validation on Params.
'
' STATE OWNERSHIP
'   Sets mParamError and logs an error naming the parameter when the value
'   is outside the range, so that LoadParams stops the run (#147).
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim InRange     As Boolean    'The value satisfies the range
    Dim RangeText   As String     'The range, in words

'------------------------------------------------------------------------------
' CHECK
'------------------------------------------------------------------------------
        If HighestAllowed Then
            InRange = (Amount >= Lowest And Amount <= Highest)
            RangeText = "from " & CStr(Lowest) & " to " & CStr(Highest)
        Else
            InRange = (Amount >= Lowest And Amount < Highest)
            RangeText = "at least " & CStr(Lowest) & " and below " & CStr(Highest)
        End If
        If WholeOnly Then
            InRange = InRange And (Amount = Int(Amount))
            RangeText = "a whole number " & RangeText
        End If
        If Not InRange Then
            mParamError = True
            LogMsg SEV_ERROR, SH_PARAMS, Code, Code & " must be " & RangeText & " (found " & CStr(Amount) & ")."
        End If

End Sub


Private Function NumParam( _
    ByVal Code As String, _
    ByVal DefaultValue As Double) _
    As Double
'
'==============================================================================
'                                   NumParam
'------------------------------------------------------------------------------
' PURPOSE
'   Read a numeric parameter, falling back to its regulatory default.
'
' INPUTS
'   Code: a PRM_ code from CORE_Config.
'   DefaultValue: value used, with a warning, when the parameter is blank.
'
' RETURNS
'   The parameter value, or DefaultValue.
'
' ERROR POLICY
'   A value that is present but not a number is an error and sets
'   mParamError, so that LoadParams stops the run (#35).
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim RawValue   As Variant    'Raw parameter value

'------------------------------------------------------------------------------
' READ
'------------------------------------------------------------------------------
        RawValue = GetParam(Code)
        If IsNum(RawValue) Then
            NumParam = CDbl(RawValue)
        ElseIf IsBlankCell(RawValue) And Not IsError(RawValue) Then
            NumParam = DefaultValue
            LogMsg SEV_WARN, SH_PARAMS, Code, "Parameter missing - default " & CStr(DefaultValue) & " used."
        Else
            NumParam = DefaultValue
            mParamError = True
            LogMsg SEV_ERROR, SH_PARAMS, Code, "Parameter must be a number (found " & DescribeValue(RawValue) & ")."
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
'   is missing or the currency is not three letters, alpha is not above
'   zero, a numeric parameter is not a number (NumParam) or is outside the
'   range of its data validation (ParamInRange), IRBucketOffset is not
'   TRUE or FALSE, the regime is not BCBS or CRR, or the IR bucket
'   correlations do not form a valid correlation matrix (#147).
'
' ERROR POLICY
'   An unexpected error while reading is logged as an error and returns
'   False.
'
' UPDATED
'   2026-10-17
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
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ParamsSheet       As Worksheet       'Params sheet
    Dim r                 As Long            'Row being read
    Dim HeaderRow         As Long            'Header row of the table
    Dim KeyColumn         As Long            'Table column of the keys
    Dim RowKey            As String          'Key of the row, upper case
    Dim TableInvalid      As Boolean         'Some row is unusable or conflicts
    Dim RowInvalid        As Boolean         'This row has a value that is not a number
    Dim FactorRow         As tSupervisory    'This row
    Dim EarlierPosition   As Long            'Position of an earlier row with the key; 0 if none
    Dim LastRow           As Long            'Last row searched for rows after the end
    Dim FXHeaderRow       As Long            'Header row of the FX table; 0 if absent

'------------------------------------------------------------------------------
' FIND THE TABLE
'------------------------------------------------------------------------------
        Set ParamsSheet = GetSheet(SH_PARAMS)
        HeaderRow = NamedRow(NM_SF_TABLE)
        KeyColumn = NamedTableCol(NM_SF_TABLE)
        If HeaderRow = 0 Then
            LogMsg SEV_ERROR, SH_PARAMS, HDR_SF, "Supervisory factor table not found (workbook name '" & _
                   NM_SF_TABLE & "')."
            Exit Function
        End If

'------------------------------------------------------------------------------
' READ ROWS
'------------------------------------------------------------------------------
    'Columns: 1 key, 2 asset class, 3 category, 4 factor, 5 correlation,
    '6 option volatility, 7 commodity hedging set, 8 regimes.
        r = HeaderRow + 1
        Do While Not IsBlankCell(TableCell(ParamsSheet, r, KeyColumn).Value)
            RowKey = UTxt(TableCell(ParamsSheet, r, KeyColumn).Value)
            FactorRow.Key = RowKey
            FactorRow.AssetClass = UTxt(TableCell(ParamsSheet, r, KeyColumn + 1).Value2)
            FactorRow.Category = UTxt(TableCell(ParamsSheet, r, KeyColumn + 2).Value2)
            RowInvalid = Not TryDbl(TableCell(ParamsSheet, r, KeyColumn + 3).Value2, 0#, FactorRow.SF)
            RowInvalid = Not TryDbl(TableCell(ParamsSheet, r, KeyColumn + 4).Value2, 0#, FactorRow.Corr) Or RowInvalid
            RowInvalid = Not TryDbl(TableCell(ParamsSheet, r, KeyColumn + 5).Value2, 0#, FactorRow.Vol) Or RowInvalid
            FactorRow.Group = UTxt(TableCell(ParamsSheet, r, KeyColumn + 6).Value2)
            FactorRow.Regimes = UTxt(TableCell(ParamsSheet, r, KeyColumn + 7).Value2)
            If FactorRow.Regimes = "BOTH" Then
                FactorRow.Regimes = ""
            End If
            If RowInvalid Then
                TableInvalid = True
                LogMsg SEV_ERROR, SH_PARAMS, RowKey, "Supervisory factor, correlation and volatility must be numbers.", r
            ElseIf FactorRow.SF < 0# Or FactorRow.SF > 1# Or FactorRow.Corr < 0# Or FactorRow.Corr > 1# _
                   Or FactorRow.Vol < 0# Or FactorRow.Vol > 5# Then
                RowInvalid = True
                TableInvalid = True
                LogMsg SEV_ERROR, SH_PARAMS, RowKey, "Supervisory factor and correlation must be from 0 to 1 and " & _
                       "the option volatility from 0 to 5 (found " & CStr(FactorRow.SF) & ", " & CStr(FactorRow.Corr) & _
                       ", " & CStr(FactorRow.Vol) & ").", r
            End If

    'A key seen before: the same values are harmless, different values are
    'ambiguous and stop the run. A row with a value that is not a number has
    'already been reported.
            EarlierPosition = KeyIndex(mSFIndex, RowKey)
            If EarlierPosition = 0 Then
                mSFCount = mSFCount + 1
                If mSFCount > UBound(mSF) Then
                    ReDim Preserve mSF(1 To mSFCount * 2)
                End If
                mSF(mSFCount) = FactorRow
                KeyAdd mSFIndex, RowKey, mSFCount
            ElseIf Not RowInvalid Then
                If SameFactorRow(mSF(EarlierPosition), FactorRow) Then
                    LogMsg SEV_WARN, SH_PARAMS, RowKey, "Supervisory factor key listed twice with the same values - " & _
                           "first occurrence used.", r
                Else
                    TableInvalid = True
                    LogMsg SEV_ERROR, SH_PARAMS, RowKey, "Supervisory factor key listed twice with different values.", r
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
        LastRow = UsedLastRow(ParamsSheet)
        FXHeaderRow = NamedRow(NM_FX_TABLE)
        If FXHeaderRow > r Then
            LastRow = FXHeaderRow - 1
        End If
        If ValuesAfterEnd(ParamsSheet, r, LastRow, KeyColumn, KeyColumn + 3, "Supervisory factor") Then
            TableInvalid = True
        End If

'------------------------------------------------------------------------------
' CHECK
'------------------------------------------------------------------------------
        If mSFCount = 0 Then
            LogMsg SEV_ERROR, SH_PARAMS, HDR_SF, "Supervisory factor table is empty."
            Exit Function
        End If
        If TableInvalid Then
            Exit Function
        End If
        LoadSFTable = True

End Function


Private Function SameFactorRow( _
    ByRef FirstEntry As tSupervisory, _
    ByRef OtherEntry As tSupervisory) _
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
'   2026-10-17
'==============================================================================
'
        SameFactorRow = (FirstEntry.AssetClass = OtherEntry.AssetClass) And (FirstEntry.Category = OtherEntry.Category) And _
                        (FirstEntry.SF = OtherEntry.SF) And (FirstEntry.Corr = OtherEntry.Corr) And (FirstEntry.Vol = OtherEntry.Vol) And _
                        (FirstEntry.Group = OtherEntry.Group) And (FirstEntry.Regimes = OtherEntry.Regimes)

End Function


Private Function LoadFXTable() As Boolean
'
'==============================================================================
'                                 LoadFXTable
'------------------------------------------------------------------------------
' PURPOSE
'   Read the FX table on Params: the rows below the FX_Ccy header, down to
'   the first blank currency. The third column of the table holds the
'   units of reporting currency per one unit of the currency.
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
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ParamsSheet       As Worksheet    'Params sheet
    Dim r                 As Long         'Row being read
    Dim HeaderRow         As Long         'Header row of the table
    Dim CurrencyColumn    As Long         'Table column of the currencies
    Dim CurrencyCode      As String       'Currency code, upper case
    Dim RowRate           As Double       'Rate of the row; -1 when missing
    Dim EarlierPosition   As Long         'Position of an earlier row with the currency; 0 if none
    Dim TableInvalid      As Boolean      'A currency conflicts or a rate would not be read
    Dim LastRow           As Long         'Last row searched for rows after the end
    Dim SFHeaderRow       As Long         'Header row of the factor table; 0 if absent

'------------------------------------------------------------------------------
' FIND THE TABLE
'------------------------------------------------------------------------------
        Set ParamsSheet = GetSheet(SH_PARAMS)
        HeaderRow = NamedRow(NM_FX_TABLE)
        CurrencyColumn = NamedTableCol(NM_FX_TABLE)
        If HeaderRow = 0 Then
            LogMsg SEV_ERROR, SH_PARAMS, HDR_FX, "FX table not found (workbook name '" & NM_FX_TABLE & "')."
            Exit Function
        End If

'------------------------------------------------------------------------------
' READ ROWS
'------------------------------------------------------------------------------
    'A currency seen before: the same rate is harmless, a different rate is
    'ambiguous and stops the run.
        r = HeaderRow + 1
        Do While Not IsBlankCell(TableCell(ParamsSheet, r, CurrencyColumn).Value)
            CurrencyCode = UTxt(TableCell(ParamsSheet, r, CurrencyColumn).Value)
            RowRate = ToDbl(TableCell(ParamsSheet, r, CurrencyColumn + 2).Value2, -1#)
            EarlierPosition = KeyIndex(mFXIndex, CurrencyCode)
            If RowRate <= 0# Then
                LogMsg SEV_WARN, SH_PARAMS, CurrencyCode, "FX rate missing or not positive - currency ignored.", r
            ElseIf EarlierPosition = 0 Then
                mFXCount = mFXCount + 1
                If mFXCount > UBound(mFXRate) Then
                    ReDim Preserve mFXRate(1 To mFXCount * 2)
                End If
                mFXRate(mFXCount) = RowRate
                KeyAdd mFXIndex, CurrencyCode, mFXCount
            ElseIf mFXRate(EarlierPosition) = RowRate Then
                LogMsg SEV_WARN, SH_PARAMS, CurrencyCode, "Currency listed twice with the same rate - first row used.", r
            Else
                TableInvalid = True
                LogMsg SEV_ERROR, SH_PARAMS, CurrencyCode, "Currency listed twice with different rates.", r
            End If
            r = r + 1
        Loop

'------------------------------------------------------------------------------
' ROWS AFTER THE END
'------------------------------------------------------------------------------
    'As for the factor table: a rate below the first blank currency would be
    'dropped silently. The search stops at the factor table when that
    'follows.
        LastRow = UsedLastRow(ParamsSheet)
        SFHeaderRow = NamedRow(NM_SF_TABLE)
        If SFHeaderRow > r Then
            LastRow = SFHeaderRow - 1
        End If
        If ValuesAfterEnd(ParamsSheet, r, LastRow, CurrencyColumn, CurrencyColumn + 2, "FX") Then
            TableInvalid = True
        End If
        If TableInvalid Then
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
    ByVal Result As Variant, _
    ByVal FunctionName As String) _
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
'   Raises ERR_FORMULA_DOMAIN when Result is an error value; Calculate
'   withdraws the results and reports it.
'
' UPDATED
'   2026-10-17
'==============================================================================
'
        If IsError(Result) Then
            Err.Raise ERR_FORMULA_DOMAIN, "CORE_Engine." & FunctionName, FunctionName & " returned " & _
                      CStr(Result) & " for inputs the engine had accepted (#42)."
        End If
        FormulaNumber = CDbl(Result)

End Function


Private Function ValuesAfterEnd( _
    ByVal ParamsSheet As Worksheet, _
    ByVal EndRow As Long, _
    ByVal LastRow As Long, _
    ByVal KeyColumn As Long, _
    ByVal ValueColumn As Long, _
    ByVal TableName As String) _
    As Boolean
'
'==============================================================================
'                                ValuesAfterEnd
'------------------------------------------------------------------------------
' PURPOSE
'   Find rows of a Params table that its loader will not read: from the
'   blank key that ends the table down to LastRow, any row with a number
'   in the table's value column. Each is logged as an error (#35). A row
'   whose Params code cell holds a parameter code is a parameter row, not
'   a table row, and is skipped.
'
' INPUTS
'   ParamsSheet: Params sheet.
'   EndRow: the row with the blank key that ended the table.
'   LastRow: last row to search.
'   KeyColumn: the table column of the table's keys.
'   ValueColumn: the table column that always holds a number in a table
'   row: the factor or the rate.
'   TableName: for the messages.
'
' RETURNS
'   True when such a row was found.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim r            As Long      'Row being checked
    Dim RowKey       As String    'Key in the key column, upper case
    Dim CodeColumn   As Long      'Table column of the parameter codes

'------------------------------------------------------------------------------
' SEARCH
'------------------------------------------------------------------------------
        CodeColumn = ParamsCodeCol()
        For r = EndRow To LastRow
            RowKey = UTxt(TableCell(ParamsSheet, r, KeyColumn).Value)
            If IsNum(TableCell(ParamsSheet, r, ValueColumn).Value) And Not IsParameterCode(UTxt(TableCell(ParamsSheet, r, CodeColumn).Value)) Then
                ValuesAfterEnd = True
                If Len(RowKey) = 0 Then
                    LogMsg SEV_ERROR, SH_PARAMS, "", TableName & " row has a value but no key - not read.", r
                Else
                    LogMsg SEV_ERROR, SH_PARAMS, RowKey, TableName & " row below a blank row - not read; " & _
                           "remove the blank row.", r
                End If
            End If
        Next r

End Function


Private Function FXRate( _
    ByVal CurrencyCode As String) _
    As Double
'
'==============================================================================
'                                    FXRate
'------------------------------------------------------------------------------
' PURPOSE
'   Look up the conversion rate of a currency.
'
' INPUTS
'   CurrencyCode: currency code, upper case; blank means the reporting
'   currency.
'
' RETURNS
'   Units of reporting currency per one unit of the currency; -1 when
'   unknown.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Position   As Long    'Position of the currency in mFXRate; 0 when unknown

'------------------------------------------------------------------------------
' LOOK UP
'------------------------------------------------------------------------------
        If Len(CurrencyCode) = 0 Then
            CurrencyCode = pRepCcy
        End If
        Position = KeyIndex(mFXIndex, CurrencyCode)
        If Position = 0 Then
            FXRate = -1#
        Else
            FXRate = mFXRate(Position)
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
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim RegulatoryRowsA   As Variant       'Regulatory factor rows, first half
    Dim RegulatoryRowsB   As Variant       'Regulatory factor rows, second half
    Dim RegulatoryRows    As Collection    'Every regulatory factor row, by key
    Dim RegulatoryRow     As Variant       'One row: key|class|category|SF|corr|vol|group|regimes
    Dim Fields            As Variant       'Fields of the row
    Dim Position          As Long          'Position of the key in mSF; 0 if absent
    Dim k                 As Long          'Netting set or factor row

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
        RegulatoryRowsA = Array( _
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
        RegulatoryRowsB = Array( _
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
        Set RegulatoryRows = New Collection
        For Each RegulatoryRow In RegulatoryRowsA
            RegulatoryRows.Add RegulatoryRow, Split(RegulatoryRow, "|")(0)
        Next RegulatoryRow
        For Each RegulatoryRow In RegulatoryRowsB
            RegulatoryRows.Add RegulatoryRow, Split(RegulatoryRow, "|")(0)
        Next RegulatoryRow
        For Each RegulatoryRow In RegulatoryRows
            Fields = Split(RegulatoryRow, "|")
            Position = KeyIndex(mSFIndex, CStr(Fields(0)))
            If Position = 0 Then
                Deviation SH_PARAMS, CStr(Fields(0)), "factor row " & Fields(0) & " is missing from the factor table."
            ElseIf mSF(Position).AssetClass <> Fields(1) Or mSF(Position).Category <> Fields(2) Or mSF(Position).Group <> Fields(6) _
                   Or mSF(Position).Regimes <> Fields(7) Or Not SameAmount(mSF(Position).SF, Val(Fields(3))) _
                   Or Not SameAmount(mSF(Position).Corr, Val(Fields(4))) Or Not SameAmount(mSF(Position).Vol, Val(Fields(5))) Then
                Deviation SH_PARAMS, CStr(Fields(0)), "factor row " & Fields(0) & " differs from the regulatory set " & _
                          "(factor " & CStr(mSF(Position).SF) & ", correlation " & CStr(mSF(Position).Corr) & _
                          ", volatility " & CStr(mSF(Position).Vol) & "; regulatory " & Fields(3) & ", " & _
                          IIf(Len(Fields(4)) = 0, "-", Fields(4)) & ", " & Fields(5) & ")."
            End If
        Next RegulatoryRow
        For k = 1 To mSFCount
            If Not InCollection(RegulatoryRows, mSF(k).Key) Then
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
    ByVal Code As String, _
    ByVal Amount As Double, _
    ByVal RegulatoryAmount As Double)
'
'==============================================================================
'                                CompareParam
'------------------------------------------------------------------------------
' PURPOSE
'   Log a "What-if:" deviation when a parameter differs from its value in
'   the regulatory set.
'
' UPDATED
'   2026-10-17
'==============================================================================
'
        If Not SameAmount(Amount, RegulatoryAmount) Then
            Deviation SH_PARAMS, Code, Code & " is " & CStr(Amount) & " (regulatory " & CStr(RegulatoryAmount) & ")."
        End If

End Sub


Private Sub Deviation( _
    ByVal SheetName As String, _
    ByVal Reference As String, _
    ByVal Detail As String)
'
'==============================================================================
'                                  Deviation
'------------------------------------------------------------------------------
' PURPOSE
'   Count one difference from the regulatory set and list it on Checks as
'   a warning starting "What-if:".
'
' UPDATED
'   2026-10-17
'==============================================================================
'
        mDeviations = mDeviations + 1
        LogMsg SEV_WARN, SheetName, Reference, "What-if: " & Detail

End Sub


Private Function SameAmount( _
    ByVal Amount As Double, _
    ByVal RegulatoryAmount As Double) _
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
'   2026-10-17
'==============================================================================
'
        SameAmount = Abs(Amount - RegulatoryAmount) <= 0.000000001 * Max2(1#, Abs(RegulatoryAmount))

End Function


Private Function InCollection( _
    ByVal Items As Collection, _
    ByVal ItemKey As String) _
    As Boolean
'
'==============================================================================
'                                 InCollection
'------------------------------------------------------------------------------
' PURPOSE
'   Tell whether a collection holds an item under a key.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Probe   As Variant    'The item, when present

'------------------------------------------------------------------------------
' LOOK UP
'------------------------------------------------------------------------------
    'Item(ItemKey) raises when the key is absent.
        On Error Resume Next
        Probe = Items.Item(ItemKey)
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
'   274(4), 275(3). A netting set facing a CCP is rejected: CRE54 and CRR
'   Art. 300-311 are not applied (#112); so is a multi-level client
'   structure, CRE54.14(3) and CRR Art. 305(4) (#359).
'
' UPDATED
'   2026-10-18
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim NettingSetsSheet   As Worksheet    'NettingSets sheet
    Dim LastRow            As Long         'Last row with a netting-set ID
    Dim TopRow             As Long         'First data row, below the headers
    Dim DataRowCount       As Long         'Number of rows to read
    Dim i                  As Long         'Row within InputBlock
    Dim RowNumber          As Long         'Sheet row of InputBlock row i
    Dim InputBlock         As Variant      'Input block, (row, column)
    Dim NettingSetRef      As String       'Netting-set ID, upper case
    Dim FloorBD            As Double       'MPOR floor before remargining, business days
    Dim EffectiveMPOR      As Double       'Effective MPOR, business days
    Dim MPOROverride       As Double       'MPOR override entered; 0 when none
    Dim RoleText           As String       'Clearing role named on Checks; empty for role N
    Dim RoleRule           As String       'Rule of that role's MPOR floor
    Dim NettingSetRegime   As String       'Regime of the netting set
    Dim FieldErrors        As Long         'Invalid fields found on the row

'------------------------------------------------------------------------------
' READ THE INPUT BLOCK
'------------------------------------------------------------------------------
        Set NettingSetsSheet = GetSheet(SH_NS)
        TopRow = TableFirstRow(SH_NS)
        LastRow = LastDataRowAny(NettingSetsSheet, TopRow, NS_NCOLS, 0)
        If LastRow < TopRow Then
            LogMsg SEV_ERROR, SH_NS, "", "No netting sets defined."
            Exit Function
        End If
        DataRowCount = LastRow - TopRow + 1
        InputBlock = TableRange(NettingSetsSheet, TopRow, 1, LastRow, NS_NCOLS).Value2

'------------------------------------------------------------------------------
' READ EACH NETTING SET
'------------------------------------------------------------------------------
    'An empty row is skipped; a row with data but no ID is an error, so a
    'netting set cannot silently drop out (#35). A repeated ID is ignored and
    'makes the first set with that ID INVALID: which row's terms apply to
    'its trades is ambiguous, so its EAD is withheld. TOTAL, TOTAL CRR and
    'TOTAL BCBS label the TOTAL rows on Results (#198), so a netting set
    'with one of them would be mistaken for a total; its row is ignored.
        For i = 1 To DataRowCount
            RowNumber = TopRow + i - 1
            NettingSetRef = UTxt(InputBlock(i, NS_ID))
            If Len(NettingSetRef) = 0 Then
                If RowHasData(InputBlock, i, NS_NCOLS, 0) Then
                    LogMsg SEV_ERROR, SH_NS, "", "Row has netting-set data but no ID - row ignored.", RowNumber
                End If
                GoTo NextRow
            End If
            If NettingSetRef = "TOTAL" Or NettingSetRef = "TOTAL " & RG_CRR Or NettingSetRef = "TOTAL " & RG_BCBS Then
                LogMsg SEV_ERROR, SH_NS, NettingSetRef, "Netting set ID " & NettingSetRef & " is reserved for a TOTAL row on Results - row ignored.", RowNumber
                GoTo NextRow
            End If
            If KeyIndex(mNSIndex, NettingSetRef) > 0 Then
                With mNS(KeyIndex(mNSIndex, NettingSetRef))
                    .InputErrors = .InputErrors + 1
                End With
                LogMsg SEV_ERROR, SH_NS, NettingSetRef, "Netting set ID listed twice - row ignored and EAD withheld.", RowNumber
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
                FieldErrors = 0
                .ID = NettingSetRef
                .Counterparty = SafeStr(InputBlock(i, NS_CPTY))
                .Margined = NsFlag(InputBlock(i, NS_MARGINED), False, "Margined", NettingSetRef, RowNumber, FieldErrors)
                .ClearingRole = NsCode(InputBlock(i, NS_CLEARED), CL_NONE, _
                                       CL_NONE & "|" & CL_CM & "|" & CL_CLIENT & "|" & CL_CCP & "|" & CL_MULTI, _
                                       "Client clearing role", NettingSetRef, RowNumber, FieldErrors)
                .RemarginBD = NsNumber(InputBlock(i, NS_FREQ), 1#, "Remargin frequency", NettingSetRef, RowNumber, FieldErrors)
                If .RemarginBD < 0# Or .RemarginBD > 10000# Or .RemarginBD <> Int(.RemarginBD) Then
                    FieldErrors = FieldErrors + 1
                    LogMsg SEV_ERROR, SH_NS, NettingSetRef, "Remargin frequency must be a whole number of business days " & _
                           "from 0 to 10000 (found " & CStr(.RemarginBD) & ").", RowNumber
                    .RemarginBD = 1#
                ElseIf .RemarginBD < 1# Then
                    LogMsg SEV_WARN, SH_NS, NettingSetRef, "Remargin frequency " & CStr(.RemarginBD) & _
                           " BD is below 1 - daily remargining (1 BD) used.", RowNumber
                    .RemarginBD = 1#
                End If
                .LargeOrIlliquid = NsFlag(InputBlock(i, NS_LARGE), False, "Large or illiquid", NettingSetRef, RowNumber, FieldErrors)
                .Disputes = NsFlag(InputBlock(i, NS_DISPUTE), False, "Margin disputes", NettingSetRef, RowNumber, FieldErrors)
                .VM = NsNumber(InputBlock(i, NS_VM), 0#, "Net VM", NettingSetRef, RowNumber, FieldErrors)
                .NICA = NsNumber(InputBlock(i, NS_NICA), 0#, "NICA", NettingSetRef, RowNumber, FieldErrors)
                .TH = NsNumber(InputBlock(i, NS_TH), 0#, "Threshold", NettingSetRef, RowNumber, FieldErrors)
                .MTA = NsNumber(InputBlock(i, NS_MTA), 0#, "MTA", NettingSetRef, RowNumber, FieldErrors)
                If .TH < 0# Then
                    FieldErrors = FieldErrors + 1
                    LogMsg SEV_ERROR, SH_NS, NettingSetRef, "Threshold TH must not be negative (found " & CStr(.TH) & ").", RowNumber
                End If
                If .MTA < 0# Then
                    FieldErrors = FieldErrors + 1
                    LogMsg SEV_ERROR, SH_NS, NettingSetRef, "MTA must not be negative (found " & CStr(.MTA) & ").", RowNumber
                End If
                .Alpha = NsNumber(InputBlock(i, NS_ALPHA), pAlpha, "Alpha override", NettingSetRef, RowNumber, FieldErrors)
                If .Alpha <= 0# Then
                    FieldErrors = FieldErrors + 1
                    LogMsg SEV_ERROR, SH_NS, NettingSetRef, "Alpha override must be greater than zero (found " & _
                           CStr(.Alpha) & ") - leave it blank to use the Params alpha.", RowNumber
                    .Alpha = pAlpha
                End If

    'Regime: blank takes the Params default; any other value than BCBS or
    'CRR, or an Excel error value, is an input error (#194).
                NettingSetRegime = UTxt(InputBlock(i, NS_REGIME))
                If Len(NettingSetRegime) = 0 And Not IsError(InputBlock(i, NS_REGIME)) Then
                    NettingSetRegime = pRegime
                ElseIf NettingSetRegime <> RG_BCBS And NettingSetRegime <> RG_CRR Then
                    FieldErrors = FieldErrors + 1
                    LogMsg SEV_ERROR, SH_NS, NettingSetRef, "Regime override must be BCBS or CRR (found " & _
                           DescribeValue(InputBlock(i, NS_REGIME)) & ").", RowNumber
                    NettingSetRegime = pRegime
                End If
                .Regime = NettingSetRegime
                .IsCRR = (NettingSetRegime = RG_CRR)

    'VM on an unmargined netting set. VM received means the counterparty
    'posts VM, so the set is margined [CRE52.2; CRR Art. 272(7)]. VM
    'posted is the one-way agreement in favour of the counterparty: under
    'BCBS the set is unmargined and C includes that VM with a negative sign
    '[CRE52.2, CRE52.10 footnote 2]; under CRR a one-way agreement is a
    'margin agreement [Art. 272(7a), 275(2)], so the set must be margined.
                If Not .Margined And .VM > 0# Then
                    FieldErrors = FieldErrors + 1
                    LogMsg SEV_ERROR, SH_NS, NettingSetRef, "VM received on an unmargined netting set: a counterparty " & _
                           "that posts VM makes the netting set margined - set Margined = Y.", RowNumber
                ElseIf Not .Margined And .VM < 0# Then
                    If .IsCRR Then
                        FieldErrors = FieldErrors + 1
                        LogMsg SEV_ERROR, SH_NS, NettingSetRef, "VM posted on an unmargined CRR netting set: under CRR " & _
                               "a one-way margin agreement is a margin agreement (Art. 272(7a)) - set Margined = Y.", RowNumber
                    Else
                        LogMsg SEV_INFO, SH_NS, NettingSetRef, "VM posted under a one-way margin agreement: unmargined " & _
                               "netting set, posted VM included in C with a negative sign (CRE52.10).", RowNumber
                    End If
                End If

    'A netting set facing a CCP, or a multi-level client structure, is
    'declared so that it is rejected rather than calculated as bilateral
    'or as an ordinary client-clearing set: the engine does not yet apply
    'the CCP rules, such as the horizon when the CCP retains VM (#112), nor
    'the conditions at each level of a multi-level structure (#359).
                Select Case .ClearingRole
                    Case CL_CCP
                        FieldErrors = FieldErrors + 1
                        LogMsg SEV_ERROR, SH_NS, NettingSetRef, "A netting set facing a CCP is not supported " & _
                               "(CRE54; CRR Art. 300-311).", RowNumber
                    Case CL_MULTI
                        FieldErrors = FieldErrors + 1
                        LogMsg SEV_ERROR, SH_NS, NettingSetRef, "A multi-level client structure is not supported " & _
                               "(CRE54.14(3); CRR Art. 305(4)).", RowNumber
                End Select

    'Margin-agreement structures the engine does not model are declared, so
    'that the netting set is rejected rather than calculated as one with a
    'single agreement of its own.
                Select Case NsCode(InputBlock(i, NS_STRUCT), "", MS_MIXED & "|" & MS_SHARED, _
                                   "Margin agreements", NettingSetRef, RowNumber, FieldErrors)
                    Case MS_MIXED
                        FieldErrors = FieldErrors + 1
                        LogMsg SEV_ERROR, SH_NS, NettingSetRef, "Several margin agreements, or margined and unmargined " & _
                               "trades, in one netting set are not supported (CRE52.74; CRR Art. 274(4)).", RowNumber
                    Case MS_SHARED
                        FieldErrors = FieldErrors + 1
                        LogMsg SEV_ERROR, SH_NS, NettingSetRef, "One margin agreement over several netting sets is not " & _
                               "supported (CRE52.75; CRR Art. 275(3), 278(2)).", RowNumber
                End Select
                .V = 0#
                .Trades = 0
                .Rejected = 0
                .EconTrades = 0
                MPOROverride = NsNumber(InputBlock(i, NS_MPOR), 0#, "MPOR override", NettingSetRef, RowNumber, FieldErrors)
                If IsNum(InputBlock(i, NS_MPOR)) And (MPOROverride <= 0# Or MPOROverride <> Int(MPOROverride)) Then
                    FieldErrors = FieldErrors + 1
                    LogMsg SEV_ERROR, SH_NS, NettingSetRef, "MPOR override must be a whole number of business days " & _
                           "greater than zero (found " & CStr(MPOROverride) & ") - leave it blank for none.", RowNumber
                End If
                .InputErrors = FieldErrors
                .MPOR = 0#
                .MFMargined = 0#

    'Effective MPOR of a margined netting set: the floor (client clearing
    'or bilateral, raised for large or illiquid sets) plus the remargining
    'period minus one day, doubled for disputes. The override, read and
    'validated above for every set, is accepted only if it is not below
    'that. The client-clearing floor applies to the clearing member facing
    'its client [CRE54.12; CRR Art. 279c(1)], and under CRR also to the
    'client facing its clearing member; under BCBS the client keeps the
    'bilateral floor, a conservative interpretation [CRE54.8, CRE54.14].
    'Every CM and CLIENT netting set gets an INFO line naming its role, the
    'floor and the rule applied, so that the evidence shows them (#359); it
    'is written after the large-or-illiquid adjustment, so that it names the
    'floor actually applied.
                If .Margined Then
                    RoleText = ""
                    RoleRule = ""
                    Select Case .ClearingRole
                        Case CL_CM
                            FloorBD = pMPORClr
                            RoleText = "Clearing member facing its client"
                            RoleRule = "CRE54.12; CRR Art. 279c(1)"
                        Case CL_CLIENT
                            If .IsCRR Then
                                FloorBD = pMPORClr
                                RoleText = "Client facing its clearing member under CRR"
                                RoleRule = "CRR Art. 279c(1)"
                            Else
                                FloorBD = pMPORBil
                                RoleText = "Client facing its clearing member under BCBS"
                                RoleRule = "a conservative interpretation: CRE54.8(2), CRE54.14-.17; the client-clearing " & _
                                           "floor of CRE54.12 is for the clearing member"
                            End If
                        Case Else
                            FloorBD = pMPORBil
                    End Select
                    If Len(RoleText) > 0 Then
                        If .LargeOrIlliquid And pMPORLarge > FloorBD Then
                            LogMsg SEV_INFO, SH_NS, NettingSetRef, RoleText & ": MPOR floor " & CStr(FloorBD) & " BD (" & _
                                   RoleRule & "), raised to " & CStr(pMPORLarge) & " BD by the large or illiquid flag " & _
                                   "(CRE52.51; CRR Art. 285(3)).", RowNumber
                        Else
                            LogMsg SEV_INFO, SH_NS, NettingSetRef, RoleText & ": MPOR floor " & CStr(FloorBD) & " BD (" & _
                                   RoleRule & ").", RowNumber
                        End If
                    End If
                    If .LargeOrIlliquid Then
                        FloorBD = Max2(FloorBD, pMPORLarge)
                    End If
                    EffectiveMPOR = FloorBD + .RemarginBD - 1#
                    If .Disputes Then
                        EffectiveMPOR = 2# * EffectiveMPOR
                    End If
                    If MPOROverride > 0# Then
                        If MPOROverride < EffectiveMPOR Then
                            LogMsg SEV_WARN, SH_NS, NettingSetRef, "MPOR override " & CStr(MPOROverride) & _
                                   " BD is below the regulatory floor - " & CStr(EffectiveMPOR) & " BD applied.", RowNumber
                        Else
                            EffectiveMPOR = MPOROverride
                        End If
                    End If
                    .MPOR = EffectiveMPOR
                    .MFMargined = FormulaNumber(SACCR_MaturityFactor(0#, True, .MPOR, pMinMatBD, pBDYear), "SACCR_MaturityFactor")
                ElseIf .ClearingRole = CL_CM Or .ClearingRole = CL_CLIENT Then
                    LogMsg SEV_INFO, SH_NS, NettingSetRef, "Client clearing role " & .ClearingRole & " on an unmargined " & _
                           "netting set: no MPOR floor applies; each trade's maturity factor uses its remaining " & _
                           "maturity (CRE52.48; CRR Art. 279c(1)(a)).", RowNumber
                End If
            End With
            KeyAdd mNSIndex, NettingSetRef, mNSCount
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
    ByVal RawValue As Variant, _
    ByVal DefaultValue As Boolean, _
    ByVal FieldName As String, _
    ByVal NettingSetRef As String, _
    ByVal RowNumber As Long, _
    ByRef FieldErrors As Long) _
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
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Result   As Boolean    'Value read

'------------------------------------------------------------------------------
' READ
'------------------------------------------------------------------------------
        If Not TryBool(RawValue, DefaultValue, Result) Then
            FieldErrors = FieldErrors + 1
            LogMsg SEV_ERROR, SH_NS, NettingSetRef, FieldName & " must be Y or N (found " & DescribeValue(RawValue) & ").", RowNumber
        End If
        NsFlag = Result

End Function


Private Function NsCode( _
    ByVal RawValue As Variant, _
    ByVal DefaultValue As String, _
    ByVal AllowedCodes As String, _
    ByVal FieldName As String, _
    ByVal NettingSetRef As String, _
    ByVal RowNumber As Long, _
    ByRef FieldErrors As Long) _
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
'   AllowedCodes: the accepted codes, upper case, separated by "|".
'
' RETURNS
'   The code in upper case, or the default.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Code   As String    'Value read, upper case

'------------------------------------------------------------------------------
' READ
'------------------------------------------------------------------------------
        Code = UTxt(RawValue)
        If IsError(RawValue) Or (Len(Code) > 0 And InStr(1, "|" & AllowedCodes & "|", "|" & Code & "|") = 0) Then
            FieldErrors = FieldErrors + 1
            LogMsg SEV_ERROR, SH_NS, NettingSetRef, FieldName & " must be " & Replace(AllowedCodes, "|", ", ") & _
                   " or blank (found " & DescribeValue(RawValue) & ").", RowNumber
            NsCode = DefaultValue
        ElseIf Len(Code) = 0 Then
            NsCode = DefaultValue
        Else
            NsCode = Code
        End If

End Function


Private Function NsNumber( _
    ByVal RawValue As Variant, _
    ByVal DefaultValue As Double, _
    ByVal FieldName As String, _
    ByVal NettingSetRef As String, _
    ByVal RowNumber As Long, _
    ByRef FieldErrors As Long) _
    As Double
'
'==============================================================================
'                                   NsNumber
'------------------------------------------------------------------------------
' PURPOSE
'   Read a numeric field of a netting set strictly: blank takes the
'   default; a value that is not a number, text or an Excel error value,
'   is logged and counted as an input error (#35).
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Result   As Double    'Value read

'------------------------------------------------------------------------------
' READ
'------------------------------------------------------------------------------
        If Not TryDbl(RawValue, DefaultValue, Result) Then
            FieldErrors = FieldErrors + 1
            LogMsg SEV_ERROR, SH_NS, NettingSetRef, FieldName & " must be a number (found " & DescribeValue(RawValue) & ").", RowNumber
        End If
        NsNumber = Result

End Function


Private Function TradeDate( _
    ByVal RawValue As Variant, _
    ByVal FieldName As String, _
    ByRef IsValid As Boolean, _
    ByRef Reasons As String) _
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
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' READ
'------------------------------------------------------------------------------
        TradeDate = ToSerial(RawValue)
        If TradeDate < 0# Then
            If IsError(RawValue) Or Not IsBlankCell(RawValue) Then
                AddErr IsValid, Reasons, FieldName & " is not a date (found " & DescribeValue(RawValue) & ")"
            End If
            TradeDate = -1#
        End If

End Function


Private Function TradeText( _
    ByVal RawValue As Variant, _
    ByVal FieldName As String, _
    ByRef IsValid As Boolean, _
    ByRef Reasons As String) _
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
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' READ
'------------------------------------------------------------------------------
        If IsError(RawValue) Then
            AddErr IsValid, Reasons, FieldName & " must be text or blank (found an error value)"
        End If
        TradeText = UTxt(RawValue)

End Function


Private Function DescribeValue( _
    ByVal Content As Variant) _
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
'   2026-10-17
'==============================================================================
'
        If IsError(Content) Then
            DescribeValue = "an error value"
        Else
            DescribeValue = "'" & SafeStr(Content) & "'"
        End If

End Function


Private Sub AddErr( _
    ByRef IsValid As Boolean, _
    ByRef Reasons As String, _
    ByVal Reason As String)
'
'==============================================================================
'                                    AddErr
'------------------------------------------------------------------------------
' PURPOSE
'   Mark the current trade as invalid and append the reason.
'
' INPUTS
'   IsValid: set to False.
'   Reasons: the trade's reasons so far; Reason is appended after "; ".
'   Reason: the reason.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' APPEND
'------------------------------------------------------------------------------
        IsValid = False
        If Len(Reasons) > 0 Then
            Reasons = Reasons & "; "
        End If
        Reasons = Reasons & Reason

End Sub


Private Sub AddWarn( _
    ByRef Warnings As String, _
    ByVal Warning As String)
'
'==============================================================================
'                                   AddWarn
'------------------------------------------------------------------------------
' PURPOSE
'   Append a warning for the current trade; the trade stays valid.
'
' INPUTS
'   Warnings: the trade's warnings so far; Warning is appended after "; ".
'   Warning: the warning.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' APPEND
'------------------------------------------------------------------------------
        If Len(Warnings) > 0 Then
            Warnings = Warnings & "; "
        End If
        Warnings = Warnings & Warning

End Sub


Private Sub CheckEconomicTrades( _
    ByVal InputBlock As Variant, _
    ByVal DataRowCount As Long, _
    ByVal TopRow As Long)
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
'   - a MtM leg flag is not Y, N or blank, or a whole trade is flagged N;
'   - its legs differ in settlement type (blank read as N) or next reset
'     date (#360).
'
' INPUTS
'   InputBlock: the Trades block read by ProcessTrades, (row, column).
'   DataRowCount: rows in InputBlock.
'   TopRow: sheet row of InputBlock row 1, for the messages.
'
' STATE OWNERSHIP
'   Fills mRowGroup, mGrpError and mGrpRejected (zeroed) and sets
'   mEconRead; adds one economic trade per group to its netting set's
'   EconTrades.
'
' UPDATED
'   2026-10-20
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Groups                 As Collection    'Group number by key: E|economic ID or L|trade ID
    Dim ClassEntries           As Collection    'Per group and asset class: index into ClassLegs
    Dim GroupName()            As String        'Economic trade ID, or the Trade ID of a whole trade
    Dim GroupNS()              As String        'Netting set of the first leg
    Dim GroupLegs()            As Long          'Legs of the group
    Dim GroupMtmLegs()         As Long          'Legs flagged Y
    Dim GroupHasEconomicID()   As Boolean       'The group has an Economic trade ID
    Dim GroupSettlement()      As String        'Settlement type of the first leg, blank read as N
    Dim GroupReset()           As Double        'Next reset date of the first leg; -1 when blank
    Dim ClassLegs()            As Long          'Legs in one asset class of one group
    Dim ClassLinearLegs()      As Long          'Of those, legs that are not options
    Dim ClassGroup()           As Long          'Group of each asset-class entry
    Dim ClassName()            As String        'Asset class of each entry
    Dim ClassCount             As Long          'Asset-class entries
    Dim i                      As Long          'Data row
    Dim g                      As Long          'Group number
    Dim a                      As Long          'Asset-class entry
    Dim TradeID                As String        'Trade ID of the row
    Dim EconomicID             As String        'Economic trade ID of the row
    Dim GroupKey               As String        'Key of the row's group
    Dim MtmLegFlag             As String        'MtM leg flag of the row
    Dim NSPosition             As Long          'Netting set of a new group; 0 if unknown
    Dim Settlement             As String        'Settlement type of the row, blank read as N

'------------------------------------------------------------------------------
' GROUP THE LEGS
'------------------------------------------------------------------------------
        ReDim mRowGroup(0 To DataRowCount)
        ReDim mGrpError(0 To DataRowCount)
        ReDim mGrpRejected(0 To DataRowCount)
        ReDim GroupName(0 To DataRowCount)
        ReDim GroupNS(0 To DataRowCount)
        ReDim GroupLegs(0 To DataRowCount)
        ReDim GroupMtmLegs(0 To DataRowCount)
        ReDim GroupHasEconomicID(0 To DataRowCount)
        ReDim GroupSettlement(0 To DataRowCount)
        ReDim GroupReset(0 To DataRowCount)
        ReDim ClassLegs(1 To DataRowCount + 1)
        ReDim ClassLinearLegs(1 To DataRowCount + 1)
        ReDim ClassGroup(1 To DataRowCount + 1)
        ReDim ClassName(1 To DataRowCount + 1)
        Set Groups = New Collection
        Set ClassEntries = New Collection
        mEconRead = 0
        ClassCount = 0
        For i = 1 To DataRowCount
            TradeID = SafeStr(InputBlock(i, TR_ID))
            If Len(TradeID) > 0 Then
                EconomicID = SafeStr(InputBlock(i, TR_ECON))
                Settlement = UTxt(InputBlock(i, TR_SETTLE))
                If Len(Settlement) = 0 Then Settlement = ST_NONE
                If Len(EconomicID) = 0 Then
                    GroupKey = "L|" & UCase$(TradeID)
                Else
                    GroupKey = "E|" & UCase$(EconomicID)
                End If
                g = KeyIndex(Groups, GroupKey)
                If g = 0 Then
                    mEconRead = mEconRead + 1
                    g = mEconRead
                    KeyAdd Groups, GroupKey, g
                    GroupName(g) = IIf(Len(EconomicID) = 0, TradeID, EconomicID)
                    GroupHasEconomicID(g) = (Len(EconomicID) > 0)
                    GroupSettlement(g) = Settlement
                    GroupReset(g) = ToSerial(InputBlock(i, TR_RESET))
                    GroupNS(g) = UTxt(InputBlock(i, TR_NS))
                    NSPosition = KeyIndex(mNSIndex, GroupNS(g))
                    If NSPosition > 0 Then
                        mNS(NSPosition).EconTrades = mNS(NSPosition).EconTrades + 1
                    End If
                ElseIf GroupHasEconomicID(g) Then
                    If UTxt(InputBlock(i, TR_NS)) <> GroupNS(g) Then
                        AddGroupError g, "its legs are in different netting sets (" & GroupNS(g) & ", " & _
                                      UTxt(InputBlock(i, TR_NS)) & ")"
                    End If

    'The legs decompose one contractual trade, so they share its settlement
    'arrangement; separately contracted trades have their own Economic
    'trade ID, or none, and may differ (#360).
                    If Settlement <> GroupSettlement(g) Then
                        AddGroupError g, "its legs have different settlement types (" & GroupSettlement(g) & _
                                      ", " & Settlement & ")"
                    ElseIf ToSerial(InputBlock(i, TR_RESET)) <> GroupReset(g) Then
                        AddGroupError g, "its legs have different next reset dates"
                    End If
                End If
                mRowGroup(i) = g
                GroupLegs(g) = GroupLegs(g) + 1
                If IsError(InputBlock(i, TR_ECON)) Then
                    AddGroupError g, "Economic trade ID must be text or blank (found an error value on row " & _
                                  (TopRow + i - 1) & ")"
                End If

    'The MtM leg flag: Y, N or blank (N for a leg, Y for a whole trade).
    'An error value is not read as blank.
                MtmLegFlag = UTxt(InputBlock(i, TR_MTMLEG))
                If IsError(InputBlock(i, TR_MTMLEG)) Then
                    AddGroupError g, "MtM leg must be Y or N (found an error value on row " & _
                                  (TopRow + i - 1) & ")"
                ElseIf MtmLegFlag = "Y" Then
                    GroupMtmLegs(g) = GroupMtmLegs(g) + 1
                ElseIf MtmLegFlag = "N" Then
                    If Not GroupHasEconomicID(g) Then
                        AddGroupError g, "a trade without an Economic trade ID carries its own MtM, so " & _
                                      "MtM leg must be Y or blank"
                    End If
                ElseIf Len(MtmLegFlag) > 0 Then
                    AddGroupError g, "MtM leg must be Y or N (found " & DescribeValue(InputBlock(i, TR_MTMLEG)) & _
                                  " on row " & (TopRow + i - 1) & ")"
                End If

    'Asset classes of the group's legs, for the decomposition rule.
                a = KeyIndex(ClassEntries, g & "|" & UTxt(InputBlock(i, TR_AC)))
                If a = 0 Then
                    ClassCount = ClassCount + 1
                    a = ClassCount
                    KeyAdd ClassEntries, g & "|" & UTxt(InputBlock(i, TR_AC)), a
                    ClassGroup(a) = g
                    ClassName(a) = UTxt(InputBlock(i, TR_AC))
                End If
                ClassLegs(a) = ClassLegs(a) + 1
                If UTxt(InputBlock(i, TR_INSTR)) <> "OPTION" Then
                    ClassLinearLegs(a) = ClassLinearLegs(a) + 1
                End If
            End If
        Next i

'------------------------------------------------------------------------------
' CHECK EACH ECONOMIC TRADE
'------------------------------------------------------------------------------
    'The MtM is counted once: exactly one MtM leg, and 0 on every other leg.
        For g = 1 To mEconRead
            If GroupHasEconomicID(g) And GroupMtmLegs(g) <> 1 Then
                AddGroupError g, "exactly one leg must be the MtM leg (MtM leg = Y), found " & GroupMtmLegs(g)
            End If
        Next g
        For i = 1 To DataRowCount
            g = mRowGroup(i)
            If g > 0 Then
                If GroupHasEconomicID(g) And UTxt(InputBlock(i, TR_MTMLEG)) <> "Y" And IsNum(InputBlock(i, TR_MTM)) Then
                    If CDbl(InputBlock(i, TR_MTM)) <> 0# Then
                        AddGroupError g, "leg " & SafeStr(InputBlock(i, TR_ID)) & " is not the MtM leg, so its MtM " & _
                                      "must be 0 (found " & DescribeValue(InputBlock(i, TR_MTM)) & ")"
                    End If
                End If
            End If
        Next i

    'Legs in one asset class must all be options: a linear product is not
    'split into legs of the same class. A Trade ID repeated without an
    'Economic trade ID is not a group: ProcessTrades rejects the repeat.
        For a = 1 To ClassCount
            If GroupHasEconomicID(ClassGroup(a)) And ClassLegs(a) > 1 And ClassLinearLegs(a) > 0 Then
                AddGroupError ClassGroup(a), "it has " & ClassLegs(a) & " legs in asset class " & ClassName(a) & _
                              " that are not all options: only an option structure may be split into " & _
                              "options (CRE52.42-43; CRR Art. 274(6)), and one risk category takes " & _
                              "one leg (CRR Art. 277(3))"
            End If
        Next a
        For g = 1 To mEconRead
            If Len(mGrpError(g)) > 0 Then
                mGrpError(g) = "economic trade " & GroupName(g) & ": " & mGrpError(g)
            End If
        Next g

End Sub


Private Sub AddGroupError( _
    ByVal GroupNumber As Long, _
    ByVal Reason As String)
'
'==============================================================================
'                                AddGroupError
'------------------------------------------------------------------------------
' PURPOSE
'   Add a reason why economic trade GroupNumber is rejected, once.
'
' UPDATED
'   2026-10-17
'==============================================================================
'
        If InStr(1, mGrpError(GroupNumber), Reason, vbBinaryCompare) = 0 Then
            If Len(mGrpError(GroupNumber)) > 0 Then
                mGrpError(GroupNumber) = mGrpError(GroupNumber) & "; "
            End If
            mGrpError(GroupNumber) = mGrpError(GroupNumber) & Reason
        End If

End Sub


Private Sub ProcessTrades( _
    ByVal WriteOutputs As Boolean)
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
'   ReadTradeDates, ReadTradeSettlement, TradeDelta, TradeHedgingSet,
'   UseTrade and TradeCalcRow (#60, #360).
'
' INPUTS
'   WriteOutputs: True writes the TradeCalc sheet.
'
' STATE OWNERSHIP
'   Adds to mBk through AddToBucket, adds MtM and trade counts to mNS, and
'   updates mTradesRead, mTradesUsed, mGrpRejected and mEconUsed;
'   CheckEconomicTrades fills the economic-trade arrays first. An invalid
'   trade is logged with all its reasons and left out of the calculation.
'
' REFERENCE
'   CRE52.30 to CRE52.52 for the trade-level quantities; CRR Art. 277 to
'   280f for OT and the regime rules.
'
' UPDATED
'   2026-10-20
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim TradesSheet      As Worksheet     'Trades sheet
    Dim TradeCalcSheet   As Worksheet     'TradeCalc sheet
    Dim OutputTop        As Long          'First TradeCalc data row
    Dim LastRow          As Long          'Last row with a trade ID
    Dim TopRow           As Long          'First data row, below the headers
    Dim DataRowCount     As Long          'Number of rows to read
    Dim i                As Long          'Row within InputBlock
    Dim RowNumber        As Long          'Sheet row of InputBlock row i
    Dim InputBlock       As Variant       'Input block, (row, column)
    Dim OutputBlock()    As Variant       'TradeCalc rows, (row, column)
    Dim OutputCount      As Long          'TradeCalc rows filled
    Dim SeenIDs          As Collection    'Trade IDs seen so far, mapped to their sheet row
    Dim NSPosition       As Long          'Netting set of a row without a trade ID
    Dim Trade            As tTrade        'The trade being processed
    Dim FreshTrade       As tTrade        'Every field at its default, to reset Trade

'------------------------------------------------------------------------------
' READ THE INPUT BLOCK
'------------------------------------------------------------------------------
    'With no trades there is nothing to calculate; the old TradeCalc rows
    'are cleared so they cannot be mistaken for results.
        Set TradesSheet = GetSheet(SH_TRADES)
        TopRow = TableFirstRow(SH_TRADES)
        LastRow = LastDataRowAny(TradesSheet, TopRow, TR_NCOLS, TR_COMMENT)
        If LastRow < TopRow Then
            LogMsg SEV_ERROR, SH_TRADES, "", "No trades found."
            If WriteOutputs Then
                ClearOutputBlock GetSheet(SH_TRADECALC), OutputFirstRow(SH_TRADECALC), TC_NCOLS
            End If
            Exit Sub
        End If
        DataRowCount = LastRow - TopRow + 1
    'Value2 gives the stored values: a Currency or Accounting format would
    'otherwise round a price or rate to four decimals, and dates come as
    'serials, which ToSerial reads (#196).
        InputBlock = TableRange(TradesSheet, TopRow, 1, LastRow, TR_NCOLS).Value2
        ReDim OutputBlock(1 To DataRowCount, 1 To TC_NCOLS)
        OutputCount = 0

'------------------------------------------------------------------------------
' PROCESS EACH TRADE
'------------------------------------------------------------------------------
    'An empty row is skipped. A row with data but no trade ID is rejected,
    'and so is a repeated trade ID; either makes its netting set INCOMPLETE
    '(#35). Trade starts from FreshTrade for every trade, so nothing leaks
    'from the previous trade into the output row.
        CheckEconomicTrades InputBlock, DataRowCount, TopRow
        Set SeenIDs = New Collection
        For i = 1 To DataRowCount
            RowNumber = TopRow + i - 1
            Trade = FreshTrade
            Trade.tid = SafeStr(InputBlock(i, TR_ID))
            If Len(Trade.tid) = 0 Then
                If RowHasData(InputBlock, i, TR_NCOLS, TR_COMMENT) Then
                    mTradesRead = mTradesRead + 1
                    LogMsg SEV_ERROR, SH_TRADES, "", "Row has trade data but no Trade ID - trade rejected.", RowNumber
                    NSPosition = KeyIndex(mNSIndex, UTxt(InputBlock(i, TR_NS)))
                    If NSPosition > 0 Then mNS(NSPosition).Rejected = mNS(NSPosition).Rejected + 1
                End If
                GoTo NextTrade
            End If
            mTradesRead = mTradesRead + 1
            Trade.ok = True
            Trade.pairSign = 1#
            ReadTradeClass Trade, InputBlock, i, RowNumber, SeenIDs
            ReadTradeAmounts Trade, InputBlock, i
            ReadTradeDates Trade, InputBlock, i
            ReadTradeSettlement Trade, InputBlock, i
            TradeDelta Trade, InputBlock, i
            TradeHedgingSet Trade
            UseTrade Trade, RowNumber
            If Not Trade.ok Then LogMsg SEV_ERROR, SH_TRADES, Trade.tid, "Trade excluded: " & Trade.msg, RowNumber
            If Not Trade.ok And Trade.nsIdx > 0 Then mNS(Trade.nsIdx).Rejected = mNS(Trade.nsIdx).Rejected + 1
            If Not Trade.ok Then mGrpRejected(mRowGroup(i)) = mGrpRejected(mRowGroup(i)) + 1
            If Len(Trade.warn) > 0 Then LogMsg SEV_WARN, SH_TRADES, Trade.tid, Trade.warn, RowNumber
            OutputCount = OutputCount + 1
            TradeCalcRow Trade, OutputBlock, OutputCount
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
        If WriteOutputs Then
            Set TradeCalcSheet = GetSheet(SH_TRADECALC)
            OutputTop = OutputFirstRow(SH_TRADECALC)
            ClearOutputBlock TradeCalcSheet, OutputTop, TC_NCOLS
            WriteBlock TradeCalcSheet, OutputTop, OutputBlock, OutputCount, TC_NCOLS
            FormatColumns TradeCalcSheet, OutputTop, OutputCount, _
                "@|@|@|@|@|@|0.00%|#,##0|#,##0|0.0000|0.0000|0.0000|0.0000|0.0000|#,##0|0.0000|0.0000|0.0000|#,##0|#,##0|0|#,##0|0.0000%|@|@"
        End If

End Sub


Private Sub ReadTradeClass( _
    ByRef Trade As tTrade, _
    ByRef InputBlock As Variant, _
    ByVal DataRow As Long, _
    ByVal RowNumber As Long, _
    ByVal SeenIDs As Collection)
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
'   Trade: the trade being processed, from ProcessTrades.
'   InputBlock, DataRow: the Trades input block and the trade's row in it.
'   RowNumber: the trade's sheet row; SeenIDs: Trade IDs seen so far.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim FirstRow        As Long      'Sheet row where a duplicate ID first appeared
    Dim DirectionText   As String    'Direction text, upper case

'------------------------------------------------------------------------------
' CLASSIFICATION
'------------------------------------------------------------------------------
    'Netting set, asset class, risk factor, instrument, direction and
    'nature. Basis and volatility trades form their own hedging sets, one
    'per label, with the factor multiplied by BasisFactor or
    'VolatilityFactor [CRE52.46, CRE52.47].
        FirstRow = KeyIndex(SeenIDs, Trade.tid)
        If FirstRow > 0 Then
            AddErr Trade.ok, Trade.msg, "duplicate Trade ID (first used on row " & FirstRow & ")"
        Else
            KeyAdd SeenIDs, Trade.tid, RowNumber
        End If

        Trade.nsId = UTxt(InputBlock(DataRow, TR_NS))
        Trade.nsIdx = KeyIndex(mNSIndex, Trade.nsId)
        If Trade.nsIdx = 0 Then AddErr Trade.ok, Trade.msg, "unknown netting set '" & Trade.nsId & "'"
        If Len(mGrpError(mRowGroup(DataRow))) > 0 Then AddErr Trade.ok, Trade.msg, mGrpError(mRowGroup(DataRow))

        Trade.acTxt = UTxt(InputBlock(DataRow, TR_AC))
        Trade.ac = ACIndex(Trade.acTxt)
        If Trade.ac = 0 Then AddErr Trade.ok, Trade.msg, "asset class must be IR, FX, CR, EQ, CO or OT"

        Trade.subCls = TradeText(InputBlock(DataRow, TR_SUB), "Sub-class", Trade.ok, Trade.msg)
        Trade.rf = UTxt(InputBlock(DataRow, TR_RF))
        If Len(Trade.rf) = 0 Then AddErr Trade.ok, Trade.msg, "risk factor / reference missing"
        If InStr(Trade.rf, "|") > 0 Or InStr(Trade.rf, "#") > 0 Then
            AddErr Trade.ok, Trade.msg, "risk factor / reference must not contain '|' or '#'"
        End If
        If Trade.ac = AC_IR And Len(Trade.rf) > 0 And _
           Not (Trade.rf Like "[A-Z][A-Z][A-Z]" Or Trade.rf Like "[A-Z][A-Z][A-Z]-INFL") Then
            AddErr Trade.ok, Trade.msg, "interest-rate risk factor must be a 3-letter currency code such as EUR, " & _
                   "or EUR-INFL for inflation"
        End If

        Trade.instType = TradeText(InputBlock(DataRow, TR_INSTR), "Instrument", Trade.ok, Trade.msg)
        If Len(Trade.instType) = 0 Then Trade.instType = "LINEAR"
        If Trade.instType <> "LINEAR" And Trade.instType <> "OPTION" And Trade.instType <> "CDO" Then
            AddErr Trade.ok, Trade.msg, "instrument type must be Linear, Option or CDO"
        End If
        If Trade.instType = "CDO" And Trade.ac <> AC_CR Then AddErr Trade.ok, Trade.msg, "CDO tranches are credit (CR) trades"

        DirectionText = UTxt(InputBlock(DataRow, TR_DIR))
        Select Case DirectionText
            Case "LONG", "BUY", "BOUGHT", "L", "B"
                Trade.isLong = True
            Case "SHORT", "SELL", "SOLD", "S"
                Trade.isLong = False
            Case Else
                AddErr Trade.ok, Trade.msg, "direction must be Long or Short"
        End Select

        Trade.nature = TradeText(InputBlock(DataRow, TR_NATURE), "Nature", Trade.ok, Trade.msg)
        If Len(Trade.nature) = 0 Then Trade.nature = "STANDARD"
        Trade.lbl = TradeText(InputBlock(DataRow, TR_LABEL), "Hedging-set label", Trade.ok, Trade.msg)
        If InStr(Trade.lbl, "|") > 0 Or InStr(Trade.lbl, "#") > 0 Then
            AddErr Trade.ok, Trade.msg, "hedging-set label must not contain '|' or '#'"
        End If
        Select Case Trade.nature
            Case "STANDARD"
                Trade.natFactor = 1#
                Trade.natTag = "STD"
            Case "BASIS"
                Trade.natFactor = pBasisF
                Trade.natTag = "BASIS:" & Trade.lbl
                If Len(Trade.lbl) = 0 Then AddErr Trade.ok, Trade.msg, "basis trade needs a hedging-set label"
            Case "VOLATILITY"
                Trade.natFactor = pVolF
                Trade.natTag = "VOL:" & Trade.lbl
                If Len(Trade.lbl) = 0 Then AddErr Trade.ok, Trade.msg, "volatility trade needs a hedging-set label"
            Case Else
                AddErr Trade.ok, Trade.msg, "nature must be Standard, Basis or Volatility"
        End Select

    'IR, FX and OT have one factor each; the other classes are looked up by
    'class and sub-class, for example CR_AAA.
        If Trade.ac > 0 Then
            Select Case Trade.ac
                Case AC_IR
                    Trade.sfKey = "IR"
                Case AC_FX
                    Trade.sfKey = "FX"
                Case AC_OT
                    Trade.sfKey = "OT"
                Case Else
                    Trade.sfKey = Trade.acTxt & "_" & Trade.subCls
            End Select
            Trade.sfIdx = KeyIndex(mSFIndex, Trade.sfKey)
            If Trade.sfIdx = 0 Then AddErr Trade.ok, Trade.msg, "no supervisory factor for '" & Trade.sfKey & "' (check sub-class)"
        End If

    'OT exists only under CRR, and a factor-table row limited to one regime
    'cannot be used by a netting set under the other.
        If Trade.nsIdx > 0 And Trade.ac = AC_OT Then
            If Not mNS(Trade.nsIdx).IsCRR Then
                AddErr Trade.ok, Trade.msg, "asset class OT (other risks) exists only under CRR (Art. 277(1)(f), 280f); " & _
                       "under BCBS map the trade to the class of its primary risk driver (CRE52)"
            End If
        End If
        If Trade.nsIdx > 0 And Trade.sfIdx > 0 Then
            If Len(mSF(Trade.sfIdx).Regimes) > 0 And mSF(Trade.sfIdx).Regimes <> mNS(Trade.nsIdx).Regime Then
                AddErr Trade.ok, Trade.msg, "'" & Trade.sfKey & "' is defined only for regime " & mSF(Trade.sfIdx).Regimes & _
                       " (netting set uses " & mNS(Trade.nsIdx).Regime & ")"
            End If
        End If

End Sub


Private Sub ReadTradeAmounts( _
    ByRef Trade As tTrade, _
    ByRef InputBlock As Variant, _
    ByVal DataRow As Long)
'
'==============================================================================
'                               ReadTradeAmounts
'------------------------------------------------------------------------------
' PURPOSE
'   Read a trade's notional and MtM with their currencies, and convert
'   both to the reporting currency when the trade is still valid (#60).
'
' INPUTS
'   Trade: the trade being processed, from ProcessTrades.
'   InputBlock, DataRow: the Trades input block and the trade's row in it.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' AMOUNTS
'------------------------------------------------------------------------------
    'The notional is unsigned; Direction gives the sign. A missing MtM is an
    'error, not an assumed 0 (#35); a blank notional currency means the
    'reporting currency, and a blank MtM currency the notional currency.
    'Both amounts are converted to the reporting currency.
        If IsNum(InputBlock(DataRow, TR_NOTIONAL)) Then
            Trade.notional = CDbl(InputBlock(DataRow, TR_NOTIONAL))
            If Trade.notional < 0# Then AddErr Trade.ok, Trade.msg, "notional must be positive (use Direction for the sign)"
        ElseIf IsBlankCell(InputBlock(DataRow, TR_NOTIONAL)) And Not IsError(InputBlock(DataRow, TR_NOTIONAL)) Then
            AddErr Trade.ok, Trade.msg, "notional missing"
        Else
            AddErr Trade.ok, Trade.msg, "notional must be a number (found " & DescribeValue(InputBlock(DataRow, TR_NOTIONAL)) & ")"
        End If
        Trade.nccy = TradeText(InputBlock(DataRow, TR_NCCY), "Notional currency", Trade.ok, Trade.msg)
        If Len(Trade.nccy) = 0 Then Trade.nccy = pRepCcy
        Trade.fxN = FXRate(Trade.nccy)
        If Trade.fxN < 0# Then AddErr Trade.ok, Trade.msg, "no FX rate for notional currency '" & Trade.nccy & "'"
        If IsNum(InputBlock(DataRow, TR_MTM)) Then
            Trade.mtm = CDbl(InputBlock(DataRow, TR_MTM))
        ElseIf IsBlankCell(InputBlock(DataRow, TR_MTM)) And Not IsError(InputBlock(DataRow, TR_MTM)) Then
            Trade.mtm = 0#
            AddErr Trade.ok, Trade.msg, "MtM missing (enter 0 if the trade has no value)"
        Else
            Trade.mtm = 0#
            AddErr Trade.ok, Trade.msg, "MtM must be a number (found " & DescribeValue(InputBlock(DataRow, TR_MTM)) & ")"
        End If
        Trade.mccy = TradeText(InputBlock(DataRow, TR_MCCY), "MtM currency", Trade.ok, Trade.msg)
        If Len(Trade.mccy) = 0 Then Trade.mccy = Trade.nccy
        Trade.fxM = FXRate(Trade.mccy)
        If Trade.fxM < 0# Then AddErr Trade.ok, Trade.msg, "no FX rate for MtM currency '" & Trade.mccy & "'"
        If Trade.ok Then
            Trade.notionalRep = Trade.notional * Trade.fxN
            Trade.mtmRep = Trade.mtm * Trade.fxM
        End If

End Sub


Private Sub ReadTradeDates( _
    ByRef Trade As tTrade, _
    ByRef InputBlock As Variant, _
    ByVal DataRow As Long)
'
'==============================================================================
'                                ReadTradeDates
'------------------------------------------------------------------------------
' PURPOSE
'   Read a trade's start, end, maturity and expiry dates and turn them into
'   the years S, E and M from the reporting date (#60).
'
' INPUTS
'   Trade: the trade being processed, from ProcessTrades.
'   InputBlock, DataRow: the Trades input block and the trade's row in it.
'
' UPDATED
'   2026-10-17
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
        Trade.dStart = TradeDate(InputBlock(DataRow, TR_START), "start date", Trade.ok, Trade.msg)
        Trade.dEnd = TradeDate(InputBlock(DataRow, TR_END), "end date", Trade.ok, Trade.msg)
        Trade.dMat = TradeDate(InputBlock(DataRow, TR_MAT), "maturity date", Trade.ok, Trade.msg)
        Trade.dExp = TradeDate(InputBlock(DataRow, TR_EXPIRY), "option expiry", Trade.ok, Trade.msg)
        If Trade.dMat < 0# Then Trade.dMat = Trade.dEnd
        If Trade.dMat < 0# Then Trade.dMat = Trade.dExp
        If Trade.dEnd < 0# Then Trade.dEnd = Trade.dMat
        If Trade.dMat < 0# Then
            AddErr Trade.ok, Trade.msg, "maturity / end date missing"
        Else
            Trade.M = (Trade.dMat - pAsOf) / pDaysYear
            Trade.E = (Trade.dEnd - pAsOf) / pDaysYear
            If Trade.dStart >= 0# Then
                Trade.S = Max2(0#, (Trade.dStart - pAsOf) / pDaysYear)
            Else
                Trade.S = 0#
            End If
            If Trade.M <= 0# Then
                AddErr Trade.ok, Trade.msg, "trade has matured (maturity <= reporting date)"
            ElseIf (Trade.ac = AC_IR Or Trade.ac = AC_CR) And Trade.E <= Trade.S Then
                AddErr Trade.ok, Trade.msg, "end date must be after start date"
            End If
        End If

    'An interest-rate or credit option takes S from its underlying
    '[CRE52.31(2), Table 1 of CRE52.32]: for a swaption the exercise date,
    'for a Bermudan swaption the earliest allowed exercise date
    '[CRE52.42(2)], for a cap or floor the start of its first period. A
    'blank start date would read as 0, the duration of a swap starting
    'today, so it is an error (#40).
        If Trade.instType = "OPTION" And (Trade.ac = AC_IR Or Trade.ac = AC_CR) And Trade.dStart < 0# Then
            If IsBlankCell(InputBlock(DataRow, TR_START)) And Not IsError(InputBlock(DataRow, TR_START)) Then
                AddErr Trade.ok, Trade.msg, "an interest-rate or credit option needs the start date S of its underlying " & _
                       "(swaption: the exercise date, the earliest one if Bermudan; cap or floor: the " & _
                       "start of the first period)"
            End If
        End If

End Sub


Private Sub ReadTradeSettlement( _
    ByRef Trade As tTrade, _
    ByRef InputBlock As Variant, _
    ByVal DataRow As Long)
'
'==============================================================================
'                             ReadTradeSettlement
'------------------------------------------------------------------------------
' PURPOSE
'   Read a trade's settlement type and next reset date (#360). A
'   settled-to-market trade (STM) settles its outstanding exposure on set
'   dates and resets its value to zero, so its remaining maturity M is the
'   time to the next reset date [CRE52.37(5); CRR Art. 279c(2)]. STM is the
'   user's attestation of these contractual facts; daily variation margin
'   alone is not STM. The engine checks the declaration and the dates.
'
' INPUTS
'   Trade: the trade being processed, after ReadTradeDates.
'   InputBlock, DataRow: the Trades input block and the trade's row in it.
'
' STATE OWNERSHIP
'   Sets MarginedSTM and adds one input error to a margined netting set
'   on its first STM trade: the set becomes INVALID.
'
' UPDATED
'   2026-10-20
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim SettlementType   As String    'Settlement type as entered, upper case

'------------------------------------------------------------------------------
' SETTLEMENT TYPE
'------------------------------------------------------------------------------
    'Blank and N are the ordinary treatment.
        SettlementType = TradeText(InputBlock(DataRow, TR_SETTLE), "Settlement type", Trade.ok, Trade.msg)
        If SettlementType = ST_STM Then
            Trade.stm = True
        ElseIf Len(SettlementType) > 0 And SettlementType <> ST_NONE Then
            AddErr Trade.ok, Trade.msg, "settlement type must be N, STM or blank (found " & _
                   DescribeValue(InputBlock(DataRow, TR_SETTLE)) & ")"
        End If

'------------------------------------------------------------------------------
' NEXT RESET DATE
'------------------------------------------------------------------------------
    'The MtM is the value after any settlement due on the reporting date,
    'so the next reset is the first one after it; a reset on the reporting
    'date would give the 10-day floor to a contract reset months later. The
    'reset is compared with the maturity as entered (or the end date or
    'expiry that fills it), before it replaces M; a final reset on the
    'maturity date is allowed. The floor of 10 business days still applies
    'in the maturity factor [CRE52.48; CRR Art. 279c(1)(a)].
        Trade.dReset = TradeDate(InputBlock(DataRow, TR_RESET), "next reset date", Trade.ok, Trade.msg)
        If Not Trade.stm Then
            If Trade.dReset >= 0# Then
                AddErr Trade.ok, Trade.msg, "a next reset date is given but the settlement type is not STM " & _
                       "(blank or N is the ordinary treatment, without a reset date)"
            End If
        ElseIf Trade.dReset < 0# Then
            If IsBlankCell(InputBlock(DataRow, TR_RESET)) And Not IsError(InputBlock(DataRow, TR_RESET)) Then
                AddErr Trade.ok, Trade.msg, "a settled-to-market trade (STM) needs its next reset date"
            End If
        ElseIf Trade.dReset <= pAsOf Then
            AddErr Trade.ok, Trade.msg, "next reset date must be after the reporting date (the MtM is the " & _
                   "value after that day's settlement)"
        ElseIf Trade.dMat >= 0# And Trade.dReset > Trade.dMat Then
            AddErr Trade.ok, Trade.msg, "next reset date is after the maturity date"
        Else
            Trade.M = (Trade.dReset - pAsOf) / pDaysYear
        End If

'------------------------------------------------------------------------------
' MARGINED NETTING SET
'------------------------------------------------------------------------------
    'A settled trade is unmargined (Basel FAQ to CRE52.31-32). In a margined
    'netting set the classification is inconsistent, or the set mixes
    'margined and unmargined trades, whose treatment [CRE52.74; CRR Art.
    '274(4)] is not implemented: the netting set is INVALID. This is a
    'limitation of the engine, not a regulatory prohibition.
        If Trade.stm And Trade.nsIdx > 0 Then
            With mNS(Trade.nsIdx)
                If .Margined And Not .MarginedSTM Then
                    .MarginedSTM = True
                    .InputErrors = .InputErrors + 1
                    LogMsg SEV_ERROR, SH_NS, .ID, "Trade " & Trade.tid & " is settled-to-market (STM) in a " & _
                           "margined netting set: STM requires unmargined treatment; mixed margined/" & _
                           "unmargined netting sets are not supported (CRE52.74; CRR Art. 274(4)) - netting " & _
                           "set INVALID, EAD withheld."
                End If
            End With
        End If

End Sub


Private Sub TradeDelta( _
    ByRef Trade As tTrade, _
    ByRef InputBlock As Variant, _
    ByVal DataRow As Long)
'
'==============================================================================
'                                  TradeDelta
'------------------------------------------------------------------------------
' PURPOSE
'   Compute the supervisory delta of a valid trade: linear, option (with
'   its expiry T and lambda shift) or CDO tranche (#60).
'
' INPUTS
'   Trade: the trade being processed, from ProcessTrades.
'   InputBlock, DataRow: the Trades input block and the trade's row in it.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Price         As Double     'Underlying price
    Dim Strike        As Double     'Strike
    Dim Attachment    As Double     'CDO attachment point
    Dim Detachment    As Double     'CDO detachment point
    Dim LambdaInput   As Variant    'Lambda as entered
    Dim DeltaResult   As Variant    'Option delta as returned by SACCR_OptionDelta

'------------------------------------------------------------------------------
' SUPERVISORY DELTA
'------------------------------------------------------------------------------
    'Linear trades: +1 long, -1 short [CRE52.39]. Options: the option delta
    'with the lambda shift; under CRR the regulatory lambda replaces any
    'entered value for IR and commodity options [CRE52.40]. CDO tranches:
    'the tranche delta [CRE52.41].
        If Trade.ok Then
            Select Case Trade.instType
                Case "LINEAR"
                    If Trade.isLong Then
                        Trade.delta = 1#
                    Else
                        Trade.delta = -1#
                    End If
                Case "OPTION"
                    Trade.optTxt = UTxt(InputBlock(DataRow, TR_OPT))
                    If Trade.optTxt <> "CALL" And Trade.optTxt <> "PUT" Then AddErr Trade.ok, Trade.msg, "option type must be Call or Put"
                    If Trade.dExp < 0# Then
                        AddErr Trade.ok, Trade.msg, "option expiry missing"
                    Else
                        Trade.T = (Trade.dExp - pAsOf) / pDaysYear
                        If Trade.T <= 0# Then
                            AddErr Trade.ok, Trade.msg, "option has expired"
                        ElseIf Trade.dExp > Trade.dMat Then
                            AddErr Trade.ok, Trade.msg, "option expiry is after the maturity date (M is the latest " & _
                                   "date the contract may still be active)"
                        End If
                    End If
                    If Not IsNum(InputBlock(DataRow, TR_PRICE)) Or Not IsNum(InputBlock(DataRow, TR_STRIKE)) Then
                        AddErr Trade.ok, Trade.msg, "option needs underlying price and strike"
                    End If
                    If Trade.ok Then
                        Price = CDbl(InputBlock(DataRow, TR_PRICE))
                        Strike = CDbl(InputBlock(DataRow, TR_STRIKE))
                        LambdaInput = InputBlock(DataRow, TR_LAMBDA)
                        If Not IsNum(LambdaInput) And Not (IsBlankCell(LambdaInput) And Not IsError(LambdaInput)) Then
                            AddErr Trade.ok, Trade.msg, "lambda must be a number (found " & DescribeValue(LambdaInput) & ")"
                        End If
                        If mNS(Trade.nsIdx).IsCRR And (Trade.ac = AC_IR Or Trade.ac = AC_CO) Then
                            'Delegated Regulation (EU) 2021/931 Art. 5, as amended by 2025/855.
                            If Trade.ac = AC_IR Then
                                Trade.lam = FormulaNumber(SACCR_LambdaCRR(Price, Strike, True, pLamThrIR), "SACCR_LambdaCRR")
                            Else
                                Trade.lam = FormulaNumber(SACCR_LambdaCRR(Price, Strike, False, pLamThrCO), "SACCR_LambdaCRR")
                            End If
                            If IsNum(LambdaInput) Then
                                If Abs(CDbl(LambdaInput) - Trade.lam) > 0.000000001 Then
                                    AddWarn Trade.warn, "lambda input ignored under CRR - RTS lambda applied"
                                End If
                            End If
                        Else
                            Trade.lam = ToDbl(LambdaInput, 0#)
                        End If
                        Trade.lamOut = Trade.lam
                        If Price + Trade.lam <= 0# Or Strike + Trade.lam <= 0# Then
                            AddErr Trade.ok, Trade.msg, "price and strike must be > 0 after lambda shift (BCBS: enter lambda per CRE52.40 FAQ)"
                        ElseIf mSF(Trade.sfIdx).Vol <= 0# Then
                            AddErr Trade.ok, Trade.msg, "supervisory volatility missing for '" & Trade.sfKey & "'"
                        Else
                            DeltaResult = SACCR_OptionDelta(Price, Strike, Trade.T, mSF(Trade.sfIdx).Vol, (Trade.optTxt = "CALL"), Trade.isLong, Trade.lam)
                            Trade.delta = FormulaNumber(DeltaResult, "SACCR_OptionDelta")
                        End If
                    End If
                Case "CDO"
                    If Not IsNum(InputBlock(DataRow, TR_ATTACH)) Or Not IsNum(InputBlock(DataRow, TR_DETACH)) Then
                        AddErr Trade.ok, Trade.msg, "CDO tranche needs attachment and detachment points"
                    Else
                        Attachment = CDbl(InputBlock(DataRow, TR_ATTACH))
                        Detachment = CDbl(InputBlock(DataRow, TR_DETACH))
                        If Attachment < 0# Or Detachment > 1# Or Attachment >= Detachment Then
                            AddErr Trade.ok, Trade.msg, "attachment/detachment must satisfy 0 <= A < D <= 1"
                        Else
                            Trade.delta = FormulaNumber(SACCR_CDODelta(Attachment, Detachment, Trade.isLong), "SACCR_CDODelta")
                        End If
                    End If
            End Select
        End If

End Sub


Private Sub TradeHedgingSet( _
    ByRef Trade As tTrade)
'
'==============================================================================
'                               TradeHedgingSet
'------------------------------------------------------------------------------
' PURPOSE
'   Give a valid trade its adjusted notional d, hedging set and bucket, with
'   the FX pair checks (#60).
'
' INPUTS
'   Trade: the trade being processed, from ProcessTrades.
'
' UPDATED
'   2026-10-17
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
    'trade [CRE52.59 FAQ1] (#40). CR: d = notional * SD, one hedging set,
    'bucket per entity [CRE52.60, CRE52.64]. EQ: one hedging set, bucket
    'per entity [CRE52.66]. CO: hedging set per commodity group, bucket per
    'commodity [CRE52.70]. OT: hedging set per primary risk driver [CRR
    'Art. 277a].
        If Trade.ok Then
            Select Case Trade.ac
                Case AC_IR
                    Trade.sd = FormulaNumber(SACCR_SupervisoryDuration(Trade.S, Trade.E, pSDFloorBD / pBDYear), "SACCR_SupervisoryDuration")
                    Trade.adjN = Trade.notionalRep * Trade.sd
                    If Trade.E < 1# Or (mNS(Trade.nsIdx).IsCRR And Trade.E <= 1#) Then
                        Trade.bucketNo = 1
                        Trade.subKey = "<1Y"
                    ElseIf Trade.E <= 5# Then
                        Trade.bucketNo = 2
                        Trade.subKey = "1-5Y"
                    Else
                        Trade.bucketNo = 3
                        Trade.subKey = ">5Y"
                    End If
                    Trade.hsKey = "IR|" & Trade.rf & "|" & Trade.natTag
                Case AC_FX
                    Trade.rf = NormalizePair(Trade.rf, Trade.pairSign)
                    If Trade.pairSign = 0# Then
                        AddErr Trade.ok, Trade.msg, "FX risk factor must be a currency pair such as EUR/USD"
                    ElseIf FXRate(Left$(Trade.rf, 3)) < 0# Then
                        AddErr Trade.ok, Trade.msg, "no FX rate for currency '" & Left$(Trade.rf, 3) & "' of the pair " & Trade.rf
                    ElseIf FXRate(Right$(Trade.rf, 3)) < 0# Then
                        AddErr Trade.ok, Trade.msg, "no FX rate for currency '" & Right$(Trade.rf, 3) & "' of the pair " & Trade.rf
                    ElseIf Trade.nccy <> Left$(Trade.rf, 3) And Trade.nccy <> Right$(Trade.rf, 3) Then
                        AddErr Trade.ok, Trade.msg, "notional currency '" & Trade.nccy & "' is not a currency of the pair " & Trade.rf
                    ElseIf Trade.instType = "OPTION" And Trade.pairSign * MarketSign(Trade.rf) < 0# Then
                        AddErr Trade.ok, Trade.msg, "FX option must be entered in market quote order as " & MarketPair(Trade.rf) & _
                               ", with price and strike in " & Right$(MarketPair(Trade.rf), 3) & " per " & _
                               Left$(MarketPair(Trade.rf), 3) & " (CRE52.59 FAQ1)"
                    End If
                    Trade.delta = Trade.delta * Trade.pairSign
                    Trade.adjN = Trade.notionalRep
                    Trade.hsKey = "FX|" & Trade.rf & "|" & Trade.natTag
                    Trade.subKey = Trade.rf
                Case AC_CR
                    Trade.sd = FormulaNumber(SACCR_SupervisoryDuration(Trade.S, Trade.E, pSDFloorBD / pBDYear), "SACCR_SupervisoryDuration")
                    Trade.adjN = Trade.notionalRep * Trade.sd
                    Trade.hsKey = "CR|" & Trade.natTag
                    Trade.subKey = Trade.rf
                Case AC_EQ
                    Trade.adjN = Trade.notionalRep
                    Trade.hsKey = "EQ|" & Trade.natTag
                    Trade.subKey = Trade.rf
                Case AC_CO
                    Trade.adjN = Trade.notionalRep
                    If Len(mSF(Trade.sfIdx).Group) = 0 Then
                        AddErr Trade.ok, Trade.msg, "commodity hedging group missing for '" & Trade.sfKey & "'"
                    End If
                    Trade.hsKey = "CO|" & mSF(Trade.sfIdx).Group & "|" & Trade.natTag
                    Trade.subKey = Trade.rf
                Case AC_OT
                    'Art. 277a: same hedging set only for an identical primary risk driver.
                    Trade.adjN = Trade.notionalRep
                    Trade.hsKey = "OT|" & Trade.rf & "|" & Trade.natTag
                    Trade.subKey = Trade.rf
            End Select
        End If

End Sub


Private Sub UseTrade( _
    ByRef Trade As tTrade, _
    ByVal RowNumber As Long)
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
'   Trade: the trade being processed, from ProcessTrades.
'   RowNumber: the trade's sheet row.
'
' STATE OWNERSHIP
'   Adds to mBk through AddToBucket, to mNS and to mTradesUsed, and to the
'   reference sub-classes through CheckReferenceClass.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' MATURITY FACTOR AND EFFECTIVE NOTIONAL
'------------------------------------------------------------------------------
    'The effective notional is kept on both bases: unmargined MF from the
    'trade's maturity, and margined MF from the netting set's MPOR. The
    'margined basis equals the unmargined one for an unmargined set.
        If Trade.ok Then
            Trade.sfEff = mSF(Trade.sfIdx).SF * Trade.natFactor
            Trade.corrEff = mSF(Trade.sfIdx).Corr
            Trade.mfU = FormulaNumber(SACCR_MaturityFactor(Trade.M, False, 0#, pMinMatBD, pBDYear), "SACCR_MaturityFactor")
            If mNS(Trade.nsIdx).Margined Then
                Trade.mfM = mNS(Trade.nsIdx).MFMargined
            Else
                Trade.mfM = Trade.mfU
            End If
            Trade.enU = Trade.delta * Trade.adjN * Trade.mfU
            Trade.enM = Trade.delta * Trade.adjN * Trade.mfM
            If Trade.ac = AC_CR Or Trade.ac = AC_EQ Or Trade.ac = AC_CO Then
                CheckReferenceClass Trade.nsIdx, Trade.ac, Trade.rf, Trade.sfKey, Trade.tid, RowNumber
            End If
            AddToBucket Trade.nsIdx, Trade.ac, Trade.hsKey, Trade.subKey, Trade.sfEff, Trade.corrEff, Trade.enU, Trade.enM, Trade.tid, RowNumber
            mNS(Trade.nsIdx).V = mNS(Trade.nsIdx).V + Trade.mtmRep
            mNS(Trade.nsIdx).Trades = mNS(Trade.nsIdx).Trades + 1
            mTradesUsed = mTradesUsed + 1
        End If

End Sub


Private Sub TradeCalcRow( _
    ByRef Trade As tTrade, _
    ByRef OutputBlock() As Variant, _
    ByVal OutputRow As Long)
'
'==============================================================================
'                                 TradeCalcRow
'------------------------------------------------------------------------------
' PURPOSE
'   Fill row OutputRow of the TradeCalc block with a trade's results, or
'   with its rejection reasons (#60).
'
' INPUTS
'   Trade: the trade being processed, from ProcessTrades.
'   OutputBlock, OutputRow: the TradeCalc block and the row to fill.
'
' UPDATED
'   2026-10-17
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
        OutputBlock(OutputRow, 1) = Trade.tid
        OutputBlock(OutputRow, 2) = Trade.nsId
        OutputBlock(OutputRow, 3) = Trade.acTxt
        If Trade.ok Then
            OutputBlock(OutputRow, 4) = "OK"
            OutputBlock(OutputRow, 5) = Trade.hsKey
            OutputBlock(OutputRow, 6) = Trade.subKey
            OutputBlock(OutputRow, 7) = Trade.sfEff
            OutputBlock(OutputRow, 8) = Trade.notionalRep
            OutputBlock(OutputRow, 9) = Trade.mtmRep
            OutputBlock(OutputRow, 10) = Trade.S
            OutputBlock(OutputRow, 11) = Trade.E
            OutputBlock(OutputRow, 12) = Trade.M
            If Trade.instType = "OPTION" Then OutputBlock(OutputRow, 13) = Trade.T
            If Trade.ac = AC_IR Or Trade.ac = AC_CR Then OutputBlock(OutputRow, 14) = Trade.sd
            OutputBlock(OutputRow, 15) = Trade.adjN
            OutputBlock(OutputRow, 16) = Trade.delta
            OutputBlock(OutputRow, 17) = Trade.mfU
            If mNS(Trade.nsIdx).Margined Then OutputBlock(OutputRow, 18) = Trade.mfM
            OutputBlock(OutputRow, 19) = Trade.enU
            If mNS(Trade.nsIdx).Margined Then OutputBlock(OutputRow, 20) = Trade.enM
            If Trade.ac = AC_IR Then OutputBlock(OutputRow, 21) = Trade.bucketNo
            OutputBlock(OutputRow, 22) = Trade.sfEff * Abs(Trade.enU)
            OutputBlock(OutputRow, 23) = Trade.lamOut
            OutputBlock(OutputRow, 24) = mNS(Trade.nsIdx).Regime
            OutputBlock(OutputRow, 25) = Trade.warn
        Else
            OutputBlock(OutputRow, 4) = "EXCLUDED"
            If Trade.nsIdx > 0 Then
                OutputBlock(OutputRow, 24) = mNS(Trade.nsIdx).Regime
            End If
            OutputBlock(OutputRow, 25) = Trade.msg
        End If

End Sub


Private Function NormalizePair( _
    ByVal PairText As String, _
    ByRef OrderSign As Double) _
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
'   PairText: the pair as entered, for example "EUR/USD", "EURUSD" or
'   "eur-usd".
'
' RETURNS
'   The pair as "AAA/BBB" with the two codes in alphabetical order;
'   PairText unchanged when it is not a valid pair.
'   OrderSign: +1 when the order was kept, -1 when it was inverted, 0 when
'   PairText is not two different three-letter codes.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Letters          As String    'Pair without separators, upper case
    Dim FirstCurrency    As String    'First currency
    Dim SecondCurrency   As String    'Second currency

'------------------------------------------------------------------------------
' NORMALIZE
'------------------------------------------------------------------------------
    'Remove the separators "/", "-", " " and "." and expect six letters.
        Letters = UCase$(Replace(Replace(Replace(Replace(PairText, "/", ""), "-", ""), " ", ""), ".", ""))
        If Len(Letters) <> 6 Then
            OrderSign = 0#
            NormalizePair = PairText
            Exit Function
        End If
        FirstCurrency = Left$(Letters, 3)
        SecondCurrency = Right$(Letters, 3)
        If FirstCurrency = SecondCurrency Then
            OrderSign = 0#
            NormalizePair = PairText
        ElseIf FirstCurrency < SecondCurrency Then
            OrderSign = 1#
            NormalizePair = FirstCurrency & "/" & SecondCurrency
        Else
            OrderSign = -1#
            NormalizePair = SecondCurrency & "/" & FirstCurrency
        End If

End Function


Private Function MarketPair( _
    ByVal Pair As String) _
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
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim FirstCurrency    As String    'First currency of pair
    Dim SecondCurrency   As String    'Second currency of pair
    Dim FirstRank        As Long      'Rank of FirstCurrency
    Dim SecondRank       As Long      'Rank of SecondCurrency

'------------------------------------------------------------------------------
' ORDER
'------------------------------------------------------------------------------
        FirstCurrency = Left$(Pair, 3)
        SecondCurrency = Right$(Pair, 3)
        FirstRank = QuoteRank(FirstCurrency)
        SecondRank = QuoteRank(SecondCurrency)
        If FirstRank < SecondRank Or (FirstRank = SecondRank And FirstCurrency < SecondCurrency) Then
            MarketPair = FirstCurrency & "/" & SecondCurrency
        Else
            MarketPair = SecondCurrency & "/" & FirstCurrency
        End If

End Function


Private Function MarketSign( _
    ByVal Pair As String) _
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
'   2026-10-17
'==============================================================================
'
        If MarketPair(Pair) = Pair Then
            MarketSign = 1#
        Else
            MarketSign = -1#
        End If

End Function


Private Function QuoteRank( _
    ByVal CurrencyCode As String) _
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
'   2026-10-17
'==============================================================================
'
        QuoteRank = InStr(1, " " & FX_QUOTE_ORDER & " ", " " & CurrencyCode & " ", vbBinaryCompare)
        If QuoteRank = 0 Then
            QuoteRank = Len(FX_QUOTE_ORDER) + 2
        End If

End Function


Private Sub CheckReferenceClass( _
    ByVal NSPosition As Long, _
    ByVal ClassIndex As Long, _
    ByVal Reference As String, _
    ByVal FactorKey As String, _
    ByVal TradeID As String, _
    ByVal RowNumber As Long)
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
'   NSPosition, ClassIndex, Reference: netting set, asset class and
'   reference of the trade.
'   FactorKey: the factor-table key the trade's sub-class gives, such as
'   CR_AA.
'   TradeID, RowNumber: trade ID and sheet row, for the message.
'
' STATE OWNERSHIP
'   Adds to mRefClass, mRefCount and mRefIndex. On the first conflict for a
'   reference, logs an error and adds an input error to the netting set,
'   which makes it INVALID with its EAD withheld whatever the row order.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ReferenceKey   As String    'Netting set | asset class | reference
    Dim Position       As Long      'Position of the reference; 0 when new

'------------------------------------------------------------------------------
' COMPARE
'------------------------------------------------------------------------------
    'References contain no "|" (checked in ReadTradeClass), so the key is
    'unambiguous.
        ReferenceKey = CStr(NSPosition) & "|" & CStr(ClassIndex) & "|" & Reference
        Position = KeyIndex(mRefIndex, ReferenceKey)
        If Position = 0 Then
            mRefCount = mRefCount + 1
            If mRefCount > UBound(mRefClass) Then
                ReDim Preserve mRefClass(1 To mRefCount * 2)
                ReDim Preserve mRefConflict(1 To mRefCount * 2)
            End If
            mRefClass(mRefCount) = FactorKey
            KeyAdd mRefIndex, ReferenceKey, mRefCount
        ElseIf mRefClass(Position) <> FactorKey And Not mRefConflict(Position) Then
            mRefConflict(Position) = True
            mNS(NSPosition).InputErrors = mNS(NSPosition).InputErrors + 1
            LogMsg SEV_ERROR, SH_TRADES, TradeID, "Reference '" & Reference & "' is " & FactorKey & " here but " & _
                   mRefClass(Position) & " on another trade of the netting set - netting set INVALID, EAD withheld.", RowNumber
        End If

End Sub


Private Sub AddToBucket( _
    ByVal NSPosition As Long, _
    ByVal ClassIndex As Long, _
    ByVal HedgingSetKey As String, _
    ByVal BucketKey As String, _
    ByVal EffectiveSF As Double, _
    ByVal Correlation As Double, _
    ByVal EffectiveNotionalU As Double, _
    ByVal EffectiveNotionalM As Double, _
    ByVal TradeID As String, _
    ByVal RowNumber As Long)
'
'==============================================================================
'                                 AddToBucket
'------------------------------------------------------------------------------
' PURPOSE
'   Add a trade's effective notionals to its bucket, creating the bucket on
'   its first trade.
'
' INPUTS
'   NSPosition, ClassIndex, HedgingSetKey, BucketKey: netting set, asset
'      class, hedging set and bucket of the trade; the netting set, hedging
'      set and bucket identify the bucket.
'   EffectiveSF, Correlation: the trade's effective supervisory factor and
'      correlation.
'   EffectiveNotionalU, EffectiveNotionalM: the trade's effective
'      notional, unmargined and margined.
'   TradeID, RowNumber: trade ID and sheet row; not used at present, as
'      nothing is logged here.
'
' STATE OWNERSHIP
'   Adds to mBk, mBkCount and mBkIndex. Every trade of a bucket has the
'   same factor and correlation: the key holds the hedging set, which
'   includes the nature, and CheckReferenceClass makes a netting set whose
'   credit, equity or commodity reference has two sub-classes INVALID
'   (#37).
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim FullKey    As String    'Bucket key: netting set # hedging set # bucket
    Dim Position   As Long      'Position of the bucket in mBk

'------------------------------------------------------------------------------
' FIND OR CREATE THE BUCKET
'------------------------------------------------------------------------------
        FullKey = CStr(NSPosition) & "#" & HedgingSetKey & "#" & BucketKey
        Position = KeyIndex(mBkIndex, FullKey)
        If Position = 0 Then
            mBkCount = mBkCount + 1
            If mBkCount > UBound(mBk) Then
                ReDim Preserve mBk(1 To mBkCount * 2)
            End If
            Position = mBkCount
            mBk(Position).NSIdx = NSPosition
            mBk(Position).AC = ClassIndex
            mBk(Position).HSKey = HedgingSetKey
            mBk(Position).SubKey = BucketKey
            mBk(Position).SF = EffectiveSF
            mBk(Position).Corr = Correlation
            mBk(Position).ENU = 0#
            mBk(Position).ENM = 0#
            mBk(Position).Trades = 0
            KeyAdd mBkIndex, FullKey, Position
        End If

'------------------------------------------------------------------------------
' ACCUMULATE
'------------------------------------------------------------------------------
        mBk(Position).ENU = mBk(Position).ENU + EffectiveNotionalU
        mBk(Position).ENM = mBk(Position).ENM + EffectiveNotionalM
        mBk(Position).Trades = mBk(Position).Trades + 1

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
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim b               As Long       'Bucket, position in mBk
    Dim h               As Long       'Hedging set, position in mHS
    Dim HedgingSetKey   As String     'Hedging-set key: netting set # hedging set
    Dim EntityAddOnU    As Double     'Entity or type add-on, unmargined
    Dim EntityAddOnM    As Double     'Entity or type add-on, margined
    Dim HedgingSetSF    As Double     'Supervisory factor of an IR, FX or OT hedging set
    Dim IRRefused()     As Boolean    'Per netting set: CRR set refused the IR alternative

'------------------------------------------------------------------------------
' COLLECT BUCKETS INTO HEDGING SETS
'------------------------------------------------------------------------------
    'IR buckets go into D1, D2 or D3. FX and OT add their effective
    'notionals. Credit, equity and commodity first turn each bucket into an
    'entity or type add-on (SF * EN), then add it to the systematic and
    'idiosyncratic sums of the single-factor model.
        For b = 1 To mBkCount
            HedgingSetKey = CStr(mBk(b).NSIdx) & "#" & mBk(b).HSKey
            h = KeyIndex(mHSIndex, HedgingSetKey)
            If h = 0 Then
                mHSCount = mHSCount + 1
                If mHSCount > UBound(mHS) Then
                    ReDim Preserve mHS(1 To mHSCount * 2)
                End If
                h = mHSCount
                mHS(h).NSIdx = mBk(b).NSIdx
                mHS(h).AC = mBk(b).AC
                mHS(h).Key = mBk(b).HSKey
                KeyAdd mHSIndex, HedgingSetKey, h
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
                    EntityAddOnU = mBk(b).SF * mBk(b).ENU
                    EntityAddOnM = mBk(b).SF * mBk(b).ENM
                    mHS(h).SysU = mHS(h).SysU + mBk(b).Corr * EntityAddOnU
                    mHS(h).IdioU = mHS(h).IdioU + (1# - mBk(b).Corr ^ 2) * EntityAddOnU * EntityAddOnU
                    mHS(h).SysM = mHS(h).SysM + mBk(b).Corr * EntityAddOnM
                    mHS(h).IdioM = mHS(h).IdioM + (1# - mBk(b).Corr ^ 2) * EntityAddOnM * EntityAddOnM
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
            ReDim IRRefused(1 To mNSCount)
        End If
        For h = 1 To mHSCount
            Select Case mHS(h).AC
                Case AC_IR
                    HedgingSetSF = BucketSF(h)
                    If Not pIRFull And mNS(mHS(h).NSIdx).IsCRR Then
                        If Not IRRefused(mHS(h).NSIdx) Then
                            IRRefused(mHS(h).NSIdx) = True
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
                    mHS(h).AddOnU = HedgingSetSF * mHS(h).ENU
                    mHS(h).AddOnM = HedgingSetSF * mHS(h).ENM
                Case AC_FX, AC_OT
                    HedgingSetSF = BucketSF(h)
                    mHS(h).AddOnU = HedgingSetSF * Abs(mHS(h).ENU)
                    mHS(h).AddOnM = HedgingSetSF * Abs(mHS(h).ENM)
                Case Else
                    mHS(h).AddOnU = Sqr(mHS(h).SysU ^ 2 + mHS(h).IdioU)
                    mHS(h).AddOnM = Sqr(mHS(h).SysM ^ 2 + mHS(h).IdioM)
            End Select
            mAddOnU(mHS(h).NSIdx, mHS(h).AC) = mAddOnU(mHS(h).NSIdx, mHS(h).AC) + mHS(h).AddOnU
            mAddOnM(mHS(h).NSIdx, mHS(h).AC) = mAddOnM(mHS(h).NSIdx, mHS(h).AC) + mHS(h).AddOnM
        Next h

End Sub


Private Function IsFactorClass( _
    ByVal ClassIndex As Long) _
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
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' TEST
'------------------------------------------------------------------------------
        IsFactorClass = (ClassIndex = AC_CR Or ClassIndex = AC_EQ Or ClassIndex = AC_CO)

End Function


Private Function BucketSF( _
    ByVal HedgingSetPosition As Long) _
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
'   HedgingSetPosition: hedging set, position in mHS.
'
' RETURNS
'   The factor of the first bucket linked to the hedging set; 0 when there
'   is none.
'
' UPDATED
'   2026-10-17
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
            If mBk(b).HSIdx = HedgingSetPosition Then
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
    ByVal WriteOutputs As Boolean)
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
'   WriteOutputs: True writes the Results sheet.
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
'   CRR Art. 285(3)(a). The status of a netting set facing a CCP or of a
'   multi-level client structure names the reason (#112, #359).
'
' UPDATED
'   2026-10-20
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
    Dim ResultsSheet       As Worksheet    'Results sheet
    Dim k                  As Long         'Netting set, position in mNS
    Dim a                  As Long         'Asset class, or Results column in loops
    Dim OutputBlock()      As Variant      'Results rows, (row, column)
    Dim OutputCount        As Long         'Results rows filled
    Dim Collateral         As Double       'Collateral C = VM + NICA
    Dim CapCollateral      As Double       'Collateral used for the unmargined basis
    Dim AggregateU         As Double       'Aggregate add-on, unmargined
    Dim AggregateM         As Double       'Aggregate add-on, margined
    Dim ReplacementCostU   As Double       'RC, unmargined basis
    Dim ReplacementCostM   As Double       'RC, margined
    Dim MultiplierU        As Double       'Multiplier, unmargined basis
    Dim MultiplierM        As Double       'Multiplier, margined
    Dim PFEU               As Double       'PFE, unmargined basis
    Dim PFEM               As Double       'PFE, margined
    Dim EADU               As Double       'EAD, unmargined basis: the EAD or the cap
    Dim EADM               As Double       'EAD, margined, before the cap
    Dim ReportedEAD        As Double       'EAD reported
    Dim RegimePosition     As Long         'Regime of the netting set: 1 CRR, 2 BCBS
    Dim FirstTotalRow      As Long         'Results row of the first TOTAL row
    Dim OutputTop          As Long         'First Results data row
    Dim StatusText         As String       'Status of an INVALID netting set
    Dim Totals(1 To 2, 1 To RS_NCOLS) As Double    'Column totals of each regime's TOTAL row

'------------------------------------------------------------------------------
' CALCULATE EACH NETTING SET
'------------------------------------------------------------------------------
        ReDim OutputBlock(1 To mNSCount + 2, 1 To RS_NCOLS)
        OutputCount = 0
        For k = 1 To mNSCount
            RegimePosition = RegimeIndex(mNS(k).Regime)
            mRegSets(RegimePosition) = mRegSets(RegimePosition) + 1

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
                mRegWithheld(RegimePosition) = mRegWithheld(RegimePosition) + 1
                LogMsg SEV_ERROR, SH_NS, mNS(k).ID, "EAD withheld: " & mNS(k).InputErrors & _
                       " invalid netting-set field(s)."
                Select Case mNS(k).ClearingRole
                    Case CL_CCP
                        StatusText = "INVALID: facing a CCP, not supported"
                    Case CL_MULTI
                        StatusText = "INVALID: multi-level client structure, not supported"
                    Case Else
                        If mNS(k).MarginedSTM Then
                            StatusText = "INVALID: STM requires unmargined treatment; mixed margined/" & _
                                         "unmargined netting sets are not supported"
                        Else
                            StatusText = ""
                        End If
                End Select
                If Len(StatusText) = 0 Then
                    StatusText = "INVALID: " & mNS(k).InputErrors & " input error(s)"
                ElseIf mNS(k).InputErrors > 1 Then
                    StatusText = StatusText & "; " & (mNS(k).InputErrors - 1) & " other input error(s)"
                End If
                OutputCount = OutputCount + 1
                WriteStatusRow OutputBlock, OutputCount, k, StatusText
                GoTo NextNS
            End If
            If mNS(k).Rejected > 0 Then
                mIncomplete = mIncomplete + 1
                mRegWithheld(RegimePosition) = mRegWithheld(RegimePosition) + 1
                LogMsg SEV_ERROR, SH_NS, mNS(k).ID, "EAD withheld: " & mNS(k).Rejected & _
                       " trade(s) rejected - see the Trades messages."
                OutputCount = OutputCount + 1
                WriteStatusRow OutputBlock, OutputCount, k, "INCOMPLETE: " & mNS(k).Rejected & " of " & _
                               (mNS(k).Trades + mNS(k).Rejected) & " trade(s) rejected"
                GoTo NextNS
            End If
            If mNS(k).Trades = 0 Then
                LogMsg SEV_INFO, SH_NS, mNS(k).ID, "Netting set has no trades - no EAD."
                OutputCount = OutputCount + 1
                WriteStatusRow OutputBlock, OutputCount, k, "NO TRADES"
                GoTo NextNS
            End If
            With mNS(k)

    'Collateral and aggregate add-ons on both bases.
                Collateral = .VM + .NICA
                AggregateU = 0#
                AggregateM = 0#
                For a = 1 To AC_COUNT
                    AggregateU = AggregateU + mAddOnU(k, a)
                    AggregateM = AggregateM + mAddOnM(k, a)
                Next a

    'Unmargined basis: the EAD of an unmargined netting set, or the cap of
    'a margined one. The cap of a margined set uses NICA only, the
    'collateral of the same set without a margin agreement [CRE52.2,
    'CRE52.10; CRR Art. 274(3), 275(1); EBA-QA-2023_6962]. An unmargined
    'set uses C: NICA, plus VM posted under a one-way agreement under BCBS.
                If .Margined Then
                    CapCollateral = .NICA
                Else
                    CapCollateral = Collateral
                End If
                ReplacementCostU = FormulaNumber(SACCR_ReplacementCost(.V, CapCollateral, False), "SACCR_ReplacementCost")
                MultiplierU = FormulaNumber(SACCR_Multiplier(.V - CapCollateral, AggregateU, pFloor), "SACCR_Multiplier")
                PFEU = MultiplierU * AggregateU
                EADU = .Alpha * (ReplacementCostU + PFEU)

    'Margined basis, then the cap: the EAD is the lower of the two.
                If .Margined Then
                    ReplacementCostM = FormulaNumber(SACCR_ReplacementCost(.V, Collateral, True, .TH, .MTA, .NICA), "SACCR_ReplacementCost")
                    MultiplierM = FormulaNumber(SACCR_Multiplier(.V - Collateral, AggregateM, pFloor), "SACCR_Multiplier")
                    PFEM = MultiplierM * AggregateM
                    EADM = .Alpha * (ReplacementCostM + PFEM)
                    ReportedEAD = Min2(EADM, EADU)
                Else
                    ReportedEAD = EADU
                End If
                mRegEAD(RegimePosition) = mRegEAD(RegimePosition) + ReportedEAD
                mRegValid(RegimePosition) = mRegValid(RegimePosition) + 1

    'Results row. A margined netting set shows its margined figures in
    'columns 9 to 18 and the cap figures in 21 and 24 to 27.
                OutputCount = OutputCount + 1
                OutputBlock(OutputCount, 1) = .ID
                OutputBlock(OutputCount, 2) = .Counterparty
                OutputBlock(OutputCount, 3) = .Regime
                OutputBlock(OutputCount, 4) = IIf(.Margined, "Y", "N")
                OutputBlock(OutputCount, 5) = .Trades
                If .Margined Then
                    OutputBlock(OutputCount, 6) = .MPOR
                End If
                OutputBlock(OutputCount, 7) = .V
                OutputBlock(OutputCount, 8) = Collateral
                For a = 1 To AC_COUNT
                    If .Margined Then
                        OutputBlock(OutputCount, 9 + a) = mAddOnM(k, a)
                    Else
                        OutputBlock(OutputCount, 9 + a) = mAddOnU(k, a)
                    End If
                Next a
                If .Margined Then
                    OutputBlock(OutputCount, 9) = ReplacementCostM
                    OutputBlock(OutputCount, 16) = AggregateM
                    OutputBlock(OutputCount, 17) = MultiplierM
                    OutputBlock(OutputCount, 18) = PFEM
                    OutputBlock(OutputCount, 20) = EADM
                    OutputBlock(OutputCount, 21) = EADU
                    OutputBlock(OutputCount, 23) = IIf(EADU < EADM, "Y", "N")
                    OutputBlock(OutputCount, 24) = CapCollateral
                    OutputBlock(OutputCount, 25) = ReplacementCostU
                    OutputBlock(OutputCount, 26) = AggregateU
                    OutputBlock(OutputCount, 27) = MultiplierU
                Else
                    OutputBlock(OutputCount, 9) = ReplacementCostU
                    OutputBlock(OutputCount, 16) = AggregateU
                    OutputBlock(OutputCount, 17) = MultiplierU
                    OutputBlock(OutputCount, 18) = PFEU
                    OutputBlock(OutputCount, 23) = "n/a"
                End If
                OutputBlock(OutputCount, 19) = .Alpha
                OutputBlock(OutputCount, 22) = ReportedEAD
                OutputBlock(OutputCount, RS_STATUS_COL) = "VALID"

    'Running totals of the regime for V, C, RC, the add-ons, PFE and EAD.
                Totals(RegimePosition, 7) = Totals(RegimePosition, 7) + .V
                Totals(RegimePosition, 8) = Totals(RegimePosition, 8) + Collateral
                For a = 9 To 16
                    Totals(RegimePosition, a) = Totals(RegimePosition, a) + ToDbl(OutputBlock(OutputCount, a))
                Next a
                Totals(RegimePosition, 18) = Totals(RegimePosition, 18) + ToDbl(OutputBlock(OutputCount, 18))
                Totals(RegimePosition, 22) = Totals(RegimePosition, 22) + ReportedEAD
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
        If WriteOutputs Then
            Set ResultsSheet = GetSheet(SH_RESULTS)
            OutputTop = OutputFirstRow(SH_RESULTS)
            ClearOutputBlock ResultsSheet, OutputTop, RS_NCOLS
            If OutputCount > 0 Then
                FirstTotalRow = OutputCount + 1
                For RegimePosition = 1 To 2
                    If mRegSets(RegimePosition) > 0 Then
                        OutputCount = OutputCount + 1
                        OutputBlock(OutputCount, 1) = "TOTAL " & Choose(RegimePosition, RG_CRR, RG_BCBS)
                        OutputBlock(OutputCount, 3) = Choose(RegimePosition, RG_CRR, RG_BCBS)
                        For a = 7 To 18
                            If a <> 17 Then
                                OutputBlock(OutputCount, a) = Totals(RegimePosition, a)
                            End If
                        Next a
                        OutputBlock(OutputCount, 22) = Totals(RegimePosition, 22)
                        OutputBlock(OutputCount, RS_STATUS_COL) = TotalStatusText(Choose(RegimePosition, RG_CRR, RG_BCBS))
                    End If
                Next RegimePosition
                WriteBlock ResultsSheet, OutputTop, OutputBlock, OutputCount, RS_NCOLS
                FormatColumns ResultsSheet, OutputTop, OutputCount, _
                    "@|@|@|@|0|0|#,##0|#,##0|#,##0|#,##0|#,##0|#,##0|#,##0|#,##0|#,##0|#,##0|0.0000|#,##0|0.00|#,##0|#,##0|#,##0|@|#,##0|#,##0|#,##0|0.0000|@"
                TableRange(ResultsSheet, OutputTop + FirstTotalRow - 1, 1, OutputTop + OutputCount - 1, RS_NCOLS).Font.Bold = True
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
    ByRef OutputBlock() As Variant, _
    ByVal OutputRow As Long, _
    ByVal NSPosition As Long, _
    ByVal Status As String)
'
'==============================================================================
'                                WriteStatusRow
'------------------------------------------------------------------------------
' PURPOSE
'   Fill a Results row for a netting set that has no EAD: its identity, its
'   valid-trade count and its status. The exposure columns stay blank.
'
' INPUTS
'   OutputBlock: the Results rows; OutputRow: the row to fill; NSPosition:
'   the netting set, position in mNS; Status: the text for the status
'   column.
'
' UPDATED
'   2026-10-17
'==============================================================================
'
        OutputBlock(OutputRow, 1) = mNS(NSPosition).ID
        OutputBlock(OutputRow, 2) = mNS(NSPosition).Counterparty
        OutputBlock(OutputRow, 3) = mNS(NSPosition).Regime
        OutputBlock(OutputRow, 4) = IIf(mNS(NSPosition).Margined, "Y", "N")
        OutputBlock(OutputRow, 5) = mNS(NSPosition).Trades
        OutputBlock(OutputRow, RS_STATUS_COL) = Status

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
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Parts(1 To 6)   As String    'Text of each input block

'------------------------------------------------------------------------------
' FINGERPRINT
'------------------------------------------------------------------------------
    'The run stops in a 1904 workbook (#185), and a renamed header stops
    'it too (#35), so both are part of what the results depend on.
        Parts(1) = FP_FORMAT & "|1904=" & IIf(ThisWorkbook.Date1904, "1", "0")
        Parts(2) = HeaderText(SH_NS, NS_NCOLS)
        Parts(3) = BlockText(GetSheet(SH_NS), TableFirstRow(SH_NS), NS_NCOLS)
        Parts(4) = HeaderText(SH_TRADES, TR_COMMENT - 1)
        Parts(5) = BlockText(GetSheet(SH_TRADES), TableFirstRow(SH_TRADES), TR_COMMENT - 1)
        Parts(6) = ParamsText()
        InputFingerprint = TextHash(Join(Parts, Chr$(29)))

End Function


Private Function HeaderText( _
    ByVal SheetName As String, _
    ByVal ColumnCount As Long) _
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
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim InputSheet    As Worksheet    'Input sheet
    Dim HeaderRow     As Long         'Header row; 0 when it cannot be located
    Dim Pieces()      As String       'One text per header
    Dim c             As Long         'Table column
    Dim HeaderValue   As Variant      'Header cell value

'------------------------------------------------------------------------------
' WRITE
'------------------------------------------------------------------------------
        HeaderRow = TableHeaderRow(SheetName)
        If HeaderRow = 0 Then
            Exit Function
        End If
        Set InputSheet = GetSheet(SheetName)
        ReDim Pieces(1 To ColumnCount)
        For c = 1 To ColumnCount
            HeaderValue = TableCell(InputSheet, HeaderRow, c).Value2
            If VarType(HeaderValue) = vbString Then
                Pieces(c) = "T" & UCase$(HeaderKey(CStr(HeaderValue)))
            Else
                Pieces(c) = CellToken(HeaderValue)
            End If
        Next c
        HeaderText = Join(Pieces, Chr$(31))

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
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ParamsSheet   As Worksheet    'Params sheet
    Dim NameList      As Variant      'Layout names, then parameter codes
    Dim i             As Long         'Index into NameList
    Dim TopLeftCell   As Range        'A name's top-left cell
    Dim Pieces()      As String       'Parameters, then the blocks
    Dim PieceCount    As Long         'Position in Pieces
    Dim SFHeaderRow   As Long         'Header row of the factor table
    Dim FXHeaderRow   As Long         'Header row of the FX table
    Dim SFLastRow     As Long         'Last row of the factor block; 0 = last used
    Dim FXLastRow     As Long         'Last row of the FX block; 0 = last used

'------------------------------------------------------------------------------
' CHECK THE LAYOUT NAMES
'------------------------------------------------------------------------------
        Set ParamsSheet = GetSheet(SH_PARAMS)
        NameList = Array(NM_PRM_CODE, NM_PRM_VALUE, NM_SF_TABLE, NM_FX_TABLE)
        For i = LBound(NameList) To UBound(NameList)
            Set TopLeftCell = NamedCell(CStr(NameList(i)))
            If TopLeftCell Is Nothing Then
                Err.Raise ERR_PARAMS_LAYOUT, "CORE_Engine.InputFingerprint", "Workbook name '" & _
                          NameList(i) & "' is missing or no longer refers to a cell."
            ElseIf TopLeftCell.Worksheet.Name <> ParamsSheet.Name Then
                Err.Raise ERR_PARAMS_LAYOUT, "CORE_Engine.InputFingerprint", "Workbook name '" & _
                          NameList(i) & "' must refer to a cell on " & SH_PARAMS & "."
            End If
        Next i

'------------------------------------------------------------------------------
' PARAMETERS
'------------------------------------------------------------------------------
    'Code=value of each named cell, in the order of ParameterCodes.
        NameList = ParameterCodes()
        ReDim Pieces(1 To UBound(NameList) - LBound(NameList) + 5)
        For i = LBound(NameList) To UBound(NameList)
            Set TopLeftCell = NamedCell(CStr(NameList(i)))
            If TopLeftCell Is Nothing Then
                Err.Raise ERR_PARAMS_LAYOUT, "CORE_Engine.InputFingerprint", "Workbook name '" & _
                          NameList(i) & "' is missing or no longer refers to a cell."
            End If
            PieceCount = PieceCount + 1
            Pieces(PieceCount) = NameList(i) & "=" & CellToken(TopLeftCell.Value2)
        Next i

'------------------------------------------------------------------------------
' COLUMNS AND TABLES
'------------------------------------------------------------------------------
    'A table starts at its header row, whose headers ParamsLayoutValid
    'checks, and stops above the other table when that follows it, as the
    'loaders stop; otherwise at its last row with a value.
        SFHeaderRow = NamedRow(NM_SF_TABLE)
        FXHeaderRow = NamedRow(NM_FX_TABLE)
        If FXHeaderRow > SFHeaderRow Then
            SFLastRow = FXHeaderRow - 1
        ElseIf SFHeaderRow > FXHeaderRow Then
            FXLastRow = SFHeaderRow - 1
        End If
        Pieces(PieceCount + 1) = BlockText(ParamsSheet, 1, 1, ParamsCodeCol())
        Pieces(PieceCount + 2) = BlockText(ParamsSheet, 1, 1, ParamsValueCol())
        Pieces(PieceCount + 3) = BlockText(ParamsSheet, SFHeaderRow, UBound(Split(SF_HEADERS, "|")) + 1, _
                                  NamedTableCol(NM_SF_TABLE), SFLastRow)
        Pieces(PieceCount + 4) = BlockText(ParamsSheet, FXHeaderRow, UBound(Split(FX_HEADERS, "|")) + 1, _
                                  NamedTableCol(NM_FX_TABLE), FXLastRow)
        ParamsText = Join(Pieces, Chr$(30))

End Function


Public Function BlockText( _
    ByVal SourceSheet As Worksheet, _
    ByVal FirstRow As Long, _
    ByVal ColumnCount As Long, _
    Optional ByVal FirstColumn As Long = 1, _
    Optional ByVal LastRowLimit As Long = 0) _
    As String
'
'==============================================================================
'                                  BlockText
'------------------------------------------------------------------------------
' PURPOSE
'   Write a block of cells as one text, cell by cell (CellToken), from
'   FirstRow to the last row with a value. Blank rows below it are left
'   out, so that a used range that grew without new values does not
'   change the text. Public for M_Main, which fingerprints the bank
'   figures with it (#45).
'
' INPUTS
'   SourceSheet, FirstRow, ColumnCount: the sheet, the block's first row
'   and its width in table columns; FirstRow below 1 gives "".
'   FirstColumn: the block's first table column, as TableCell counts it.
'   LastRowLimit: the last row to read; 0 for the last used row.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim LastRow          As Long       'Last row read
    Dim RawBlock         As Variant    'Cell values, Value2 (dates as serials)
    Dim Grid             As Variant    'The same as a 2-D array, also for one cell
    Dim FilledRowCount   As Long       'Rows up to the last with a value
    Dim Pieces()         As String     'One text per cell
    Dim r                As Long       'Row of Grid
    Dim c                As Long       'Column of Grid
    Dim PieceCount       As Long       'Position in Pieces

'------------------------------------------------------------------------------
' READ
'------------------------------------------------------------------------------
        If FirstRow < 1 Then
            Exit Function
        End If
        LastRow = UsedLastRow(SourceSheet)
        If LastRowLimit > 0 And LastRowLimit < LastRow Then
            LastRow = LastRowLimit
        End If
        If LastRow < FirstRow Then
            Exit Function
        End If
        RawBlock = TableRange(SourceSheet, FirstRow, FirstColumn, LastRow, FirstColumn + ColumnCount - 1).Value2
        If IsArray(RawBlock) Then
            Grid = RawBlock
        Else
            ReDim Grid(1 To 1, 1 To 1)
            Grid(1, 1) = RawBlock
        End If
        For r = UBound(Grid, 1) To 1 Step -1
            If RowHasData(Grid, r, ColumnCount, 0) Then
                FilledRowCount = r
                Exit For
            End If
        Next r
        If FilledRowCount = 0 Then
            Exit Function
        End If

'------------------------------------------------------------------------------
' WRITE AS TEXT
'------------------------------------------------------------------------------
    'Chr$(31) separates cells; the column count fixes where rows end.
        ReDim Pieces(1 To FilledRowCount * ColumnCount)
        For r = 1 To FilledRowCount
            For c = 1 To ColumnCount
                PieceCount = PieceCount + 1
                Pieces(PieceCount) = CellToken(Grid(r, c))
            Next c
        Next r
        BlockText = Join(Pieces, Chr$(31))

End Function


Private Function CellToken( _
    ByVal Content As Variant) _
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
'   2026-10-17
'==============================================================================
'
        Select Case VarType(Content)
            Case vbEmpty
                CellToken = ""
            Case vbString
                CellToken = "T" & Content
            Case vbBoolean
                CellToken = IIf(Content, "B1", "B0")
            Case vbError
                CellToken = "E" & CStr(Content)
            Case Else
                CellToken = "N" & CStr(Content)
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
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim StoredName   As Name    'The stored name, Nothing when absent

'------------------------------------------------------------------------------
' READ
'------------------------------------------------------------------------------
    'The name refers to a text constant: ="5A428560".
        On Error Resume Next
        Set StoredName = ThisWorkbook.Names(RUN_INPUTS_NAME)
        Err.Clear
        On Error GoTo 0
        If Not StoredName Is Nothing Then
            LastRunInputs = Replace(Replace(StoredName.RefersTo, "=", ""), """", "")
        End If

End Function


Private Sub RememberRunInputs( _
    ByVal Fingerprint As String)
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
'   structure) still runs, but its fingerprint is not stored, so its
'   results are not reported as current.
'
' UPDATED
'   2026-10-17
'==============================================================================
'
        On Error GoTo Failed
        ThisWorkbook.Names.Add Name:=RUN_INPUTS_NAME, RefersTo:="=""" & Fingerprint & """", Visible:=False
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


Private Function WithdrawOutputs(ByRef FailureDetails As String) As Boolean
'
'==============================================================================
'                               WithdrawOutputs
'------------------------------------------------------------------------------
' PURPOSE
'   Attempt every output independently after a run fails (#36, review #85).
' RETURNS
'   True only if all four clears succeeded. FailureDetails is reset on entry
'   and contains every failed sheet, error number, source and description.
' ERROR POLICY
'   Each helper contains its own error; a failed clear never skips another.
'   The caller keeps the original run error and appends cleanup diagnostics.
' UPDATED
'   2026-10-17
'==============================================================================
'
        FailureDetails = ""
        WithdrawOutputs = True
        If Not TryClearOutput(SH_TRADECALC, TC_NCOLS, FailureDetails) Then WithdrawOutputs = False
        If Not TryClearOutput(SH_BUCKETS, BK_NCOLS, FailureDetails) Then WithdrawOutputs = False
        If Not TryClearOutput(SH_HEDGING, HS_NCOLS, FailureDetails) Then WithdrawOutputs = False
        If Not TryClearOutput(SH_RESULTS, RS_NCOLS, FailureDetails) Then WithdrawOutputs = False
End Function


Private Function TryClearOutput( _
    ByVal SheetName As String, _
    ByVal ColumnCount As Long, _
    ByRef FailureDetails As String) As Boolean
'
'==============================================================================
'                                TryClearOutput
'------------------------------------------------------------------------------
' PURPOSE
'   Clear one output and capture its failure before any other operation.
' INPUTS
'   SheetName, ColumnCount: output sheet and table width, from the first
'   table column.
' RETURNS
'   True only when this output was cleared. Appends any failure to
'   FailureDetails, preserving details from earlier sheets.
' ERROR POLICY
'   Contains this cleanup failure; the original run error remains primary.
' UPDATED
'   2026-10-17
'==============================================================================
    Dim ErrNumber        As Long      'Cleanup error, captured before formatting
    Dim ErrSource        As String    'Cleanup error source
    Dim ErrDescription   As String    'Cleanup error description
        On Error GoTo Failed
        If gEngineFault = "outputs+cleanup" Then
            If SheetName = SH_TRADECALC Or SheetName = SH_RESULTS Then
                Err.Raise ERR_CLEANUP_FAILED, "CORE_Engine.TryClearOutput", "Injected output cleanup failure."
            End If
        End If
        ClearOutputBlock GetSheet(SheetName), OutputFirstRow(SheetName), ColumnCount
        TryClearOutput = True
        Exit Function
Failed:
        ErrNumber = Err.Number
        ErrSource = Err.Source
        ErrDescription = Err.Description
        If Len(FailureDetails) > 0 Then FailureDetails = FailureDetails & "; "
        FailureDetails = FailureDetails & SheetName & " [error " & CStr(ErrNumber) & _
                         "; source=" & ErrSource & "]: " & ErrDescription
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
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim BucketsSheet    As Worksheet    'Buckets sheet
    Dim TopRow          As Long         'First Buckets data row
    Dim OutputBlock()   As Variant      'Output rows, (row, column)
    Dim b               As Long         'Bucket, position in mBk

'------------------------------------------------------------------------------
' BUILD ROWS
'------------------------------------------------------------------------------
    'Columns: 1 netting set, 2 asset class, 3 hedging set, 4 bucket,
    '5 trades, 6 SF, 7 correlation, 8 and 9 effective notional unmargined
    'and margined, 10 and 11 entity or type add-on unmargined and margined.
    'Correlation and add-ons apply to credit, equity and commodity only;
    'margined figures to margined netting sets only.
        Set BucketsSheet = GetSheet(SH_BUCKETS)
        TopRow = OutputFirstRow(SH_BUCKETS)
        ClearOutputBlock BucketsSheet, TopRow, BK_NCOLS
        If mBkCount = 0 Then
            Exit Sub
        End If
        ReDim OutputBlock(1 To mBkCount, 1 To BK_NCOLS)
        For b = 1 To mBkCount
            OutputBlock(b, 1) = mNS(mBk(b).NSIdx).ID
            OutputBlock(b, 2) = ACCode(mBk(b).AC)
            OutputBlock(b, 3) = mBk(b).HSKey
            OutputBlock(b, 4) = mBk(b).SubKey
            OutputBlock(b, 5) = mBk(b).Trades
            OutputBlock(b, 6) = mBk(b).SF
            If IsFactorClass(mBk(b).AC) Then OutputBlock(b, 7) = mBk(b).Corr
            OutputBlock(b, 8) = mBk(b).ENU
            OutputBlock(b, 9) = IIf(mNS(mBk(b).NSIdx).Margined, mBk(b).ENM, Empty)
            If IsFactorClass(mBk(b).AC) Then
                OutputBlock(b, 10) = mBk(b).SF * mBk(b).ENU
                If mNS(mBk(b).NSIdx).Margined Then OutputBlock(b, 11) = mBk(b).SF * mBk(b).ENM
            End If
        Next b

'------------------------------------------------------------------------------
' WRITE
'------------------------------------------------------------------------------
        WriteBlock BucketsSheet, TopRow, OutputBlock, mBkCount, BK_NCOLS
        FormatColumns BucketsSheet, TopRow, mBkCount, "@|@|@|@|0|0.00%|0%|#,##0|#,##0|#,##0|#,##0"

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
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim HedgingSheet    As Worksheet    'HedgingSets sheet
    Dim TopRow          As Long         'First HedgingSets data row
    Dim OutputBlock()   As Variant      'Output rows, (row, column)
    Dim h               As Long         'Hedging set, position in mHS
    Dim IsMargined      As Boolean      'Netting set of h is margined

'------------------------------------------------------------------------------
' BUILD ROWS
'------------------------------------------------------------------------------
    'Columns: 1 netting set, 2 asset class, 3 hedging set, 4 trades, 5 to 7
    'IR buckets D1 to D3, 8 effective notional (IR, FX, OT), 9 systematic
    'and 10 idiosyncratic part (credit, equity, commodity), 11 hedging-set
    'add-on, 12 add-on unmargined, 13 basis, 14 SF (IR, FX, OT).
        Set HedgingSheet = GetSheet(SH_HEDGING)
        TopRow = OutputFirstRow(SH_HEDGING)
        ClearOutputBlock HedgingSheet, TopRow, HS_NCOLS
        If mHSCount = 0 Then
            Exit Sub
        End If
        ReDim OutputBlock(1 To mHSCount, 1 To HS_NCOLS)
        For h = 1 To mHSCount
            IsMargined = mNS(mHS(h).NSIdx).Margined
            OutputBlock(h, 1) = mNS(mHS(h).NSIdx).ID
            OutputBlock(h, 2) = ACCode(mHS(h).AC)
            OutputBlock(h, 3) = mHS(h).Key
            OutputBlock(h, 4) = mHS(h).Trades
            If mHS(h).AC = AC_IR Then
                If IsMargined Then
                    OutputBlock(h, 5) = mHS(h).D1M
                    OutputBlock(h, 6) = mHS(h).D2M
                    OutputBlock(h, 7) = mHS(h).D3M
                Else
                    OutputBlock(h, 5) = mHS(h).D1U
                    OutputBlock(h, 6) = mHS(h).D2U
                    OutputBlock(h, 7) = mHS(h).D3U
                End If
            End If
            If Not IsFactorClass(mHS(h).AC) Then
                If IsMargined Then
                    OutputBlock(h, 8) = mHS(h).ENM
                Else
                    OutputBlock(h, 8) = mHS(h).ENU
                End If
            Else
                If IsMargined Then
                    OutputBlock(h, 9) = mHS(h).SysM
                    OutputBlock(h, 10) = Sqr(mHS(h).IdioM)
                Else
                    OutputBlock(h, 9) = mHS(h).SysU
                    OutputBlock(h, 10) = Sqr(mHS(h).IdioU)
                End If
            End If
            If IsMargined Then
                OutputBlock(h, 11) = mHS(h).AddOnM
            Else
                OutputBlock(h, 11) = mHS(h).AddOnU
            End If
            OutputBlock(h, 12) = mHS(h).AddOnU
            OutputBlock(h, 13) = IIf(IsMargined, "Margined", "Unmargined")
            If Not IsFactorClass(mHS(h).AC) Then
                OutputBlock(h, 14) = BucketSF(h)
            End If
        Next h

'------------------------------------------------------------------------------
' WRITE
'------------------------------------------------------------------------------
        WriteBlock HedgingSheet, TopRow, OutputBlock, mHSCount, HS_NCOLS
        FormatColumns HedgingSheet, TopRow, mHSCount, "@|@|@|0|#,##0|#,##0|#,##0|#,##0|#,##0|#,##0|#,##0|#,##0|@|0.00%"

End Sub


Private Function WriteChecks( _
    ByRef Failure As String) _
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
'   reason in Failure; Calculate then withdraws the results (#36).
'
' STATE OWNERSHIP
'   Clears and rewrites the Checks sheet.
'
' ERROR POLICY
'   Contains any error and reports it through the result, never silently.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ChecksSheet     As Worksheet    'Checks sheet
    Dim TopRow          As Long         'First Checks data row
    Dim OutputBlock()   As Variant      'Output rows, (message, field)
    Dim i               As Long         'Message
    Dim j               As Long         'Field

'------------------------------------------------------------------------------
' WRITE
'------------------------------------------------------------------------------
    'mLog is stored (field, message); it is transposed into sheet rows.
        On Error GoTo Failed
        Failure = ""
        Set ChecksSheet = GetSheet(SH_CHECKS)
        TopRow = OutputFirstRow(SH_CHECKS)
        ClearOutputBlock ChecksSheet, TopRow, CK_NCOLS
        If mLogCount = 0 Then
            ReDim OutputBlock(1 To 1, 1 To CK_NCOLS)
            OutputBlock(1, 1) = SEV_INFO
            OutputBlock(1, 5) = "No issues found."
            WriteBlock ChecksSheet, TopRow, OutputBlock, 1, CK_NCOLS
            WriteChecks = True
            Exit Function
        End If
        ReDim OutputBlock(1 To mLogCount, 1 To CK_NCOLS)
        For i = 1 To mLogCount
            For j = 1 To CK_NCOLS
                OutputBlock(i, j) = mLog(j, i)
            Next j
        Next i
        WriteBlock ChecksSheet, TopRow, OutputBlock, mLogCount, CK_NCOLS
        For i = 1 To mLogCount
            If OutputBlock(i, 1) = SEV_ERROR Then
                TableCell(ChecksSheet, TopRow + i - 1, 1).Font.Bold = True
            End If
        Next i
        WriteChecks = True
        Exit Function

'------------------------------------------------------------------------------
' REPORT FAILURE
'------------------------------------------------------------------------------
Failed:
        Failure = "error " & Err.Number & ": " & Err.Description

End Function


Private Sub WriteRunInfo( _
    ByVal RunSeconds As Double, _
    ByVal Completed As Boolean, _
    ByVal WroteOutputs As Boolean, _
    ByVal Failure As String)
'
'==============================================================================
'                                 WriteRunInfo
'------------------------------------------------------------------------------
' PURPOSE
'   Write a one-line run summary (RunInfoText) to cell C4 (RUNINFO_CELL) of
'   the Results sheet.
'
' INPUTS
'   RunSeconds: run duration in seconds.
'   Completed: the result of Calculate.
'   WroteOutputs: False for a validation-only run, whose outputs were
'   cleared rather than written.
'   Failure: why the run failed after it started writing, for example the
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
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Summary   As String    'Summary line

'------------------------------------------------------------------------------
' COMPOSE
'------------------------------------------------------------------------------
        On Error GoTo NoText
        If gEngineFault = "runinfo" Then
            Err.Raise ERR_INJECTED_FAULT, "CORE_Engine.WriteRunInfo", "Injected run-summary failure."
        End If
        Summary = RunInfoText(RunSeconds, Completed, WroteOutputs, Failure)

'------------------------------------------------------------------------------
' WRITE
'------------------------------------------------------------------------------
WriteText:
        On Error Resume Next
        GetSheet(SH_RESULTS).Range(RUNINFO_CELL).Value = Summary
        On Error GoTo 0
        Exit Sub

'------------------------------------------------------------------------------
' HANDLE A SUMMARY THAT CANNOT BE COMPOSED
'------------------------------------------------------------------------------
NoText:
        Summary = "Last run " & Format$(Now, "yyyy-mm-dd hh:mm:ss") & " - run summary unavailable (error " & _
              Err.Number & ": " & Err.Description & "); see Checks"
        Resume WriteText

End Sub


Private Function RunInfoText( _
    ByVal RunSeconds As Double, _
    ByVal Completed As Boolean, _
    ByVal WroteOutputs As Boolean, _
    ByVal Failure As String) _
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
'   2026-10-17
'==============================================================================
'
        If Len(Failure) > 0 Then
            RunInfoText = "Last run " & Format$(Now, "yyyy-mm-dd hh:mm:ss") & " FAILED - " & Failure
        ElseIf Not WroteOutputs Then
            RunInfoText = "Last validation " & Format$(Now, "yyyy-mm-dd hh:mm:ss") & _
                          " | errors " & mErrCount & ", warnings " & mWarnCount & _
                          RunClassText() & _
                          " | outputs cleared: press Run SA-CCR to calculate"
        ElseIf Completed Then
            RunInfoText = "Last run " & Format$(Now, "yyyy-mm-dd hh:mm:ss") & _
                          " | reporting date " & Format$(CDate(pAsOf), "yyyy-mm-dd") & _
                          " | ccy " & pRepCcy & _
                          " | economic trades used " & mEconUsed & " of " & mEconRead & _
                          ", legs " & mTradesUsed & " of " & mTradesRead & _
                          " | errors " & mErrCount & ", warnings " & mWarnCount & _
                          RunClassText() & _
                          " | inputs " & mRunInputs & _
                          " | " & CORE_Export.NumberText(Round(RunSeconds, 2)) & " s"
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
