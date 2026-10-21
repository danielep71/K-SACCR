Attribute VB_Name = "TEST_CaseRunner"
'==============================================================================
' MODULE: TEST_CaseRunner
'------------------------------------------------------------------------------
' PURPOSE
'   Run the numerical test cases of tests/fixtures and tests/expected
'   against the engine. The generated module TEST_Cases calls these
'   procedures: for each case it writes the fixture into the input sheets,
'   runs the engine, and compares the output cells with the expected values
'   (methodology, running test cases; #44).
'
' PUBLIC SURFACE
'   InjectFailure, BeginSuite, BeginCase, LoadSavedInputs, SetParameter,
'   AddNettingSet, AddTrade, SetInputCell, PatchCell, PatchName, RunCase,
'   RunCaseExpectingStop, OutputNumber, ExpectNumber, ExpectText,
'   ExpectTrue, ExpectChecksLine and EndSuite, for TEST_Cases,
'   TEST_InputValidation, TEST_Aggregation, TEST_Invariants and
'   TEST_Benchmark. Option Private Module keeps them out of the external
'   workbook automation API.
'
' DEPENDENCIES
'   CORE_Engine.Calculate; CORE_Util for sheet access; CORE_Config for the
'   layout; TEST_Suite.Say for the log lines. The workbook must be built
'   from the template, because the engine reads its input sheets and
'   Params.
'
' WORKSHEET SAFETY
'   Unlike TEST_Harness, this module writes the NettingSets and Trades input
'   rows, the AsOfDate and ReportingCcy parameters, single cells through
'   PatchCell (Params values set by SetParameter, the FX rate set by
'   UnitFxRate), and every output sheet. BeginSuite saves the inputs and
'   parameters, and EndSuite writes them back and runs the engine again, so
'   the workbook ends as it started. A cell or workbook name changed with
'   PatchCell or PatchName is restored at the start of the next case and
'   by EndSuite. Use a development workbook.
'
' STATE OWNERSHIP
'   Owns the counters and saved inputs below. BeginSuite captures and
'   EndSuite restores calculation mode, events and screen updating.
'
' ERROR POLICY
'   A failed engine run is a failure of the case and a missing output a
'   failed check; the suite continues. Restoration steps run independently and a failed one
'   makes the suite fail with restore=FAIL.
'
' REPORTING
'   Each case prints a CASE line and each failed check a FAILURE line. The
'   summary counts results per reference class: illustrative results are
'   reported separately and are never evidence of correctness. A suite
'   started after InjectFailure prints MODE=INJECTED_FAILURE and must end
'   with RESULT=FAIL and exactly one failure.
'
' UPDATED
'   2026-10-21
'
' AUTHOR
'   Daniele Penza
'==============================================================================

'------------------------------------------------------------------------------
' MODULE SETTINGS
'------------------------------------------------------------------------------
    'Require explicit declarations; keep the runner out of the external API.
    Option Explicit
    Option Private Module

'------------------------------------------------------------------------------
' MODULE CONSTANTS
'------------------------------------------------------------------------------
    'Reference classes, as indexes into the per-class counters.
        Private Const CLASS_PUBLISHED      As Long = 0    'Printed in a registered source
        Private Const CLASS_INDEPENDENT    As Long = 1    'Independently derived and reviewed
        Private Const CLASS_ILLUSTRATIVE   As Long = 2    'Shows a behaviour; validates nothing

'------------------------------------------------------------------------------
' MODULE STATE
'------------------------------------------------------------------------------
    'Suite progress.
        Private mSuiteActive     As Boolean    'Between BeginSuite and EndSuite
        Private mExpectedCases   As Long       'Cases the calling suite will run
        Private mExpectedChecks  As Long       'Checks the calling suite will run
        Private mCaseCount       As Long       'Cases started
        Private mCheckCount      As Long       'Checks evaluated
        Private mFailureCount    As Long       'Failed checks, runs and restorations
        Private mPassed(0 To 2)  As Long       'Passed checks per reference class
        Private mFailed(0 To 2)  As Long       'Failed checks per reference class

    'The case being run.
        Private mCase            As String     'Case label: fixture.regime
        Private mRegime          As String     'CRR or BCBS
        Private mCurrency        As String     'Calculation currency
        Private mNettingSetId    As String     'The fixture's netting set
        Private mNextTradeRow    As Long       'Next free row on Trades

    'What BeginSuite saved, for EndSuite.
        Private mSavedCalculation   As XlCalculation    'Calculation mode
        Private mSavedEvents        As Boolean          'EnableEvents
        Private mSavedScreen        As Boolean          'ScreenUpdating
        Private mSavedNsAddress     As String           'NettingSets input block, "" if empty
        Private mSavedNs            As Variant          'Its formulas
        Private mSavedTrAddress     As String           'Trades input block, "" if empty
        Private mSavedTr            As Variant          'Its formulas
        Private mAsOfCell           As Range            'Cell holding AsOfDate
        Private mSavedAsOf          As Variant          'Its formula
        Private mCcyCell            As Range            'Cell holding ReportingCcy
        Private mSavedCcy           As Variant          'Its formula

    'What PatchCell and PatchName changed, in order, for RestorePatches.
        Private mPatchCount         As Long             'Changes recorded
        Private mPatchSheet()       As String           'Sheet of a cell; "" for a workbook name
        Private mPatchTarget()      As String           'Cell address, or the workbook name
        Private mPatchSaved()       As Variant          'Its formula, or what the name referred to

    'Failure injection, to prove that a wrong value is caught.
        Private mInjectArmed        As Boolean          'Set by InjectFailure for the next suite
        Private mInjectFailure      As Boolean          'This suite shifts its first numeric expectation
        Private mInjectedAt         As String           'Case and check it was applied to; "" before


'
'------------------------------------------------------------------------------
'
'                                SUITE CONTROL
'
'------------------------------------------------------------------------------
'

Public Sub InjectFailure()
'
'==============================================================================
'                                InjectFailure
'------------------------------------------------------------------------------
' PURPOSE
'   Demonstrate the failure path of any suite run through this module: the
'   next suite moves the expected value of its first numeric check by ten
'   times its tolerance plus one, so the real comparison must reject it.
'   The suite must then print MODE=INJECTED_FAILURE and an INJECTED line,
'   end with RESULT=FAIL and failures=1, and still restore the workbook.
'
' USAGE
'   On one line of the Immediate window, for example:
'   TEST_CaseRunner.InjectFailure: TEST_Cases.RunCaseTests
'
' STATE OWNERSHIP
'   Arms a flag that the next BeginSuite consumes and EndSuite clears, so
'   it never outlives one suite.
'
' UPDATED
'   2026-10-09
'==============================================================================
'
        mInjectArmed = True

End Sub


Public Sub BeginSuite( _
    ByVal ExpectedCases As Long, _
    ByVal ExpectedChecks As Long)
'
'==============================================================================
'                                  BeginSuite
'------------------------------------------------------------------------------
' PURPOSE
'   Save everything the cases will overwrite, then prepare Excel.
'
' INPUTS
'   ExpectedCases, ExpectedChecks: numbers of cases and checks the calling
'   suite will run; a different count fails the suite, so a stale or edited
'   suite module, such as TEST_Cases, cannot pass with checks missing.
'
' ERROR POLICY
'   Raises if a suite is already active or the inputs cannot be saved;
'   nothing has been changed at that point.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' GUARD AND RESET
'------------------------------------------------------------------------------
        If mSuiteActive Then
            Err.Raise ERR_RUN_ACTIVE, "TEST_CaseRunner.BeginSuite", "A case suite is already running."
        End If
        mExpectedCases = ExpectedCases
        mExpectedChecks = ExpectedChecks
        mInjectFailure = mInjectArmed
        mInjectArmed = False
        mInjectedAt = ""
        mCaseCount = 0
        mCheckCount = 0
        mFailureCount = 0
        mPatchCount = 0
        Erase mPassed
        Erase mFailed

'------------------------------------------------------------------------------
' SAVE INPUTS AND SETTINGS
'------------------------------------------------------------------------------
    'Formulas, not values, so that any formula in an input cell comes back.
        mSavedCalculation = Application.Calculation
        mSavedEvents = Application.EnableEvents
        mSavedScreen = Application.ScreenUpdating
        SaveBlock GetSheet(SH_NS), NS_NCOLS, mSavedNsAddress, mSavedNs
        SaveBlock GetSheet(SH_TRADES), TR_NCOLS, mSavedTrAddress, mSavedTr
        Set mAsOfCell = ParamCell(PRM_ASOF)
        mSavedAsOf = mAsOfCell.Formula
        Set mCcyCell = ParamCell(PRM_REPCCY)
        mSavedCcy = mCcyCell.Formula

'------------------------------------------------------------------------------
' PREPARE EXCEL
'------------------------------------------------------------------------------
        mSuiteActive = True
        Application.ScreenUpdating = False
        Application.EnableEvents = False
        Application.Calculation = xlCalculationManual
        TEST_Suite.Say "SACCR CASE TESTS"
        TEST_Suite.Say "MODE=" & IIf(mInjectFailure, "INJECTED_FAILURE", "NORMAL")

End Sub


Public Sub EndSuite( _
    ByVal AbortReason As String)
'
'==============================================================================
'                                   EndSuite
'------------------------------------------------------------------------------
' PURPOSE
'   Restore the inputs, parameters, outputs and Excel settings, and print
'   the summary and the RESULT line.
'
' INPUTS
'   AbortReason: why the calling suite stopped early; empty after a
'   complete run.
'
' ERROR POLICY
'   Each restoration step contains its own error; a failed one is printed
'   and makes the result FAIL.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Restored   As Boolean    'Every restoration step succeeded
    Dim Complete   As Boolean    'Every expected case ran

'------------------------------------------------------------------------------
' RESTORE
'------------------------------------------------------------------------------
    'The engine is run once more on the restored inputs so the output
    'sheets match them again.
        If Len(AbortReason) > 0 Then
            mFailureCount = mFailureCount + 1
            TEST_Suite.Say "FAILURE=suite: " & AbortReason
        End If
        Restored = True
        Restored = RestoreStep("inputs", 1) And Restored
        Restored = RestoreStep("parameters", 2) And Restored
        Restored = RestoreStep("outputs", 3) And Restored
        Restored = RestoreStep("calculation", 4) And Restored
        Restored = RestoreStep("events", 5) And Restored
        Restored = RestoreStep("screen updating", 6) And Restored
        mSuiteActive = False

'------------------------------------------------------------------------------
' REPORT
'------------------------------------------------------------------------------
        If mInjectFailure And Len(mInjectedAt) = 0 Then
            mFailureCount = mFailureCount + 1
            TEST_Suite.Say "FAILURE=suite: injected failure requested, but the suite ran no numeric check"
        End If
        mInjectFailure = False
        Complete = (mCaseCount = mExpectedCases) And (mCheckCount = mExpectedChecks)
        If Not Complete Then
            mFailureCount = mFailureCount + 1
            TEST_Suite.Say "FAILURE=suite: expected " & mExpectedCases & " cases and " & mExpectedChecks & _
                        " checks, ran " & mCaseCount & " and " & mCheckCount
        End If
        TEST_Suite.Say "PUBLISHED: passed=" & mPassed(CLASS_PUBLISHED) & "; failed=" & mFailed(CLASS_PUBLISHED)
        TEST_Suite.Say "INDEPENDENT: passed=" & mPassed(CLASS_INDEPENDENT) & "; failed=" & mFailed(CLASS_INDEPENDENT)
        TEST_Suite.Say "ILLUSTRATIVE: passed=" & mPassed(CLASS_ILLUSTRATIVE) & "; failed=" & _
                    mFailed(CLASS_ILLUSTRATIVE) & " (validates nothing)"
        TEST_Suite.Say "RESULT=" & IIf(mFailureCount = 0 And Restored And Complete, "PASS", "FAIL") & _
                    "; cases=" & mCaseCount & "; checks=" & mCheckCount & _
                    "; failures=" & mFailureCount & "; restore=" & IIf(Restored, "PASS", "FAIL")

End Sub


'
'------------------------------------------------------------------------------
'
'                                CASE BUILDING
'
'------------------------------------------------------------------------------
'

Public Sub BeginCase( _
    ByVal CaseId As String, _
    ByVal RegimeCode As String, _
    ByVal ValuationDate As String, _
    ByVal CalculationCurrency As String)
'
'==============================================================================
'                                  BeginCase
'------------------------------------------------------------------------------
' PURPOSE
'   Start a case: undo the previous case's patches, empty the input rows
'   and set the reporting date and currency of the fixture. Every amount
'   of a fixture is in its calculation currency, so that currency's rate
'   in the FX table is set to 1 for the case when it is not already: the
'   template's rates are per EUR, and a USD case would otherwise be
'   converted.
'
' INPUTS
'   CaseId: the fixture or case ID; the case label adds the regime.
'   RegimeCode: CRR or BCBS, applied as the netting set's regime override.
'   ValuationDate: YYYY-MM-DD.
'   CalculationCurrency: the fixture's calculation currency.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' RESET INPUTS
'------------------------------------------------------------------------------
        RestorePatches
        mCaseCount = mCaseCount + 1
        mCase = CaseId & "." & LCase$(RegimeCode)
        mRegime = RegimeCode
        mCurrency = CalculationCurrency
        mNettingSetId = ""
        mNextTradeRow = InputTop(SH_TRADES)
        ClearBlock GetSheet(SH_NS), NS_NCOLS
        ClearBlock GetSheet(SH_TRADES), TR_NCOLS
        mAsOfCell.Value = IsoDate(ValuationDate)
        mCcyCell.Value = CalculationCurrency
        UnitFxRate CalculationCurrency
        TEST_Suite.Say "CASE=" & mCase

End Sub


Public Sub LoadSavedInputs()
'
'==============================================================================
'                               LoadSavedInputs
'------------------------------------------------------------------------------
' PURPOSE
'   Put the workbook's own inputs, saved by BeginSuite, back on the input
'   sheets for the current case: its NettingSets and Trades rows and its
'   AsOfDate and ReportingCcy. Used by TEST_Aggregation to run the demo
'   portfolio in different row orders and by TEST_Invariants to check its
'   results.
'
' UPDATED
'   2026-10-07
'==============================================================================
'
        RestoreBlock GetSheet(SH_NS), NS_NCOLS, mSavedNsAddress, mSavedNs
        RestoreBlock GetSheet(SH_TRADES), TR_NCOLS, mSavedTrAddress, mSavedTr
        mAsOfCell.Formula = mSavedAsOf
        mCcyCell.Formula = mSavedCcy

End Sub


Public Sub AddNettingSet( _
    ByVal NettingSetKey As String, _
    ByVal MarginedFlag As String, _
    ByVal RoleCode As String, _
    ByVal RemarginDays As String, _
    ByVal LargeFlag As String, _
    ByVal DisputesFlag As String, _
    ByVal MporOverrideDays As String, _
    ByVal VmAmount As String, _
    ByVal NicaAmount As String, _
    ByVal ThresholdAmount As String, _
    ByVal MtaAmount As String, _
    ByVal AlphaFactor As String)
'
'==============================================================================
'                                AddNettingSet
'------------------------------------------------------------------------------
' PURPOSE
'   Write the fixture's netting set to the first NettingSets row, with the
'   case's regime as its override. The margin-agreement structure is left
'   blank: one agreement, or none, covering this netting set only.
'
' INPUTS
'   The fixture fields as text: Y or N for flags, N, CM or CLIENT for the
'   clearing role, numbers with a "." decimal point, empty for null.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim RowValues(1 To 1, 1 To NS_NCOLS)   As Variant    'The row to write

'------------------------------------------------------------------------------
' WRITE
'------------------------------------------------------------------------------
        mNettingSetId = NettingSetKey
        RowValues(1, NS_ID) = NettingSetKey
        RowValues(1, NS_CPTY) = "Test case " & mCase
        RowValues(1, NS_MARGINED) = MarginedFlag
        RowValues(1, NS_CLEARED) = RoleCode
        RowValues(1, NS_FREQ) = NumberOrBlank(RemarginDays)
        RowValues(1, NS_LARGE) = LargeFlag
        RowValues(1, NS_DISPUTE) = DisputesFlag
        RowValues(1, NS_MPOR) = NumberOrBlank(MporOverrideDays)
        RowValues(1, NS_VM) = NumberOrBlank(VmAmount)
        RowValues(1, NS_NICA) = NumberOrBlank(NicaAmount)
        RowValues(1, NS_TH) = NumberOrBlank(ThresholdAmount)
        RowValues(1, NS_MTA) = NumberOrBlank(MtaAmount)
        RowValues(1, NS_ALPHA) = NumberOrBlank(AlphaFactor)
        RowValues(1, NS_REGIME) = mRegime
        WriteRow GetSheet(SH_NS), InputTop(SH_NS), NS_NCOLS, RowValues

End Sub


Public Sub AddTrade( _
    ByVal TradeId As String, _
    ByVal AssetClassCode As String, _
    ByVal SubClass As String, _
    ByVal RiskFactor As String, _
    ByVal Instrument As String, _
    ByVal Direction As String, _
    ByVal OptionType As String, _
    ByVal Nature As String, _
    ByVal HedgingSetLabel As String, _
    ByVal NotionalAmount As String, _
    ByVal MarketValueAmount As String, _
    ByVal StartDate As String, _
    ByVal EndDate As String, _
    ByVal MaturityDate As String, _
    ByVal ExpiryDate As String, _
    ByVal UnderlyingPrice As String, _
    ByVal StrikePrice As String, _
    ByVal LambdaPrice As String, _
    ByVal AttachmentRate As String, _
    ByVal DetachmentRate As String)
'
'==============================================================================
'                                   AddTrade
'------------------------------------------------------------------------------
' PURPOSE
'   Write one trade of the fixture to the next Trades row, in the netting
'   set written by AddNettingSet. Amounts are in the calculation currency.
'
' INPUTS
'   The fixture fields as text, already in the Trades sheet's vocabulary
'   (IR, Linear, Long, ...): numbers with a "." decimal point, dates as
'   YYYY-MM-DD, empty for null.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim RowValues(1 To 1, 1 To TR_NCOLS)   As Variant    'The row to write

'------------------------------------------------------------------------------
' WRITE
'------------------------------------------------------------------------------
        RowValues(1, TR_ID) = TradeId
        RowValues(1, TR_NS) = mNettingSetId
        RowValues(1, TR_AC) = AssetClassCode
        RowValues(1, TR_SUB) = TextOrBlank(SubClass)
        RowValues(1, TR_RF) = RiskFactor
        RowValues(1, TR_INSTR) = Instrument
        RowValues(1, TR_DIR) = Direction
        RowValues(1, TR_OPT) = TextOrBlank(OptionType)
        RowValues(1, TR_NATURE) = Nature
        RowValues(1, TR_LABEL) = TextOrBlank(HedgingSetLabel)
        RowValues(1, TR_NOTIONAL) = NumberOrBlank(NotionalAmount)
        RowValues(1, TR_NCCY) = mCurrency
        RowValues(1, TR_MTM) = NumberOrBlank(MarketValueAmount)
        RowValues(1, TR_MCCY) = mCurrency
        RowValues(1, TR_START) = IsoDate(StartDate)
        RowValues(1, TR_END) = IsoDate(EndDate)
        RowValues(1, TR_MAT) = IsoDate(MaturityDate)
        RowValues(1, TR_EXPIRY) = IsoDate(ExpiryDate)
        RowValues(1, TR_PRICE) = NumberOrBlank(UnderlyingPrice)
        RowValues(1, TR_STRIKE) = NumberOrBlank(StrikePrice)
        RowValues(1, TR_LAMBDA) = NumberOrBlank(LambdaPrice)
        RowValues(1, TR_ATTACH) = NumberOrBlank(AttachmentRate)
        RowValues(1, TR_DETACH) = NumberOrBlank(DetachmentRate)
        WriteRow GetSheet(SH_TRADES), mNextTradeRow, TR_NCOLS, RowValues
        mNextTradeRow = mNextTradeRow + 1

End Sub


Public Sub SetInputCell( _
    ByVal SheetName As String, _
    ByVal DataRow As Long, _
    ByVal TableColumn As Long, _
    ByVal NewValue As Variant)
'
'==============================================================================
'                                 SetInputCell
'------------------------------------------------------------------------------
' PURPOSE
'   Overwrite one input cell after AddNettingSet or AddTrade, for invalid
'   inputs that a JSON fixture cannot express: a blank ID, a duplicate ID,
'   a typo in a flag. Used by TEST_InputValidation and TEST_Aggregation.
'
' INPUTS
'   SheetName: SH_NS or SH_TRADES.
'   DataRow: 1 for the first data row.
'   TableColumn: table column number, for example TR_ID.
'   NewValue: the value to write; Empty clears the cell.
'
' UPDATED
'   2026-10-17
'==============================================================================
'
        TableCell(GetSheet(SheetName), InputTop(SheetName) + DataRow - 1, TableColumn).Value = NewValue

End Sub


Public Sub PatchCell( _
    ByVal SheetName As String, _
    ByVal RowNumber As Long, _
    ByVal TableColumn As Long, _
    ByVal NewValue As Variant)
'
'==============================================================================
'                                  PatchCell
'------------------------------------------------------------------------------
' PURPOSE
'   Overwrite any cell outside the input rows, such as a header or a row of
'   a Params table, for one case. The cell is restored at the start of the
'   next case and by EndSuite. Used by TEST_InputValidation,
'   TEST_Invariants, SetParameter and UnitFxRate.
'
' INPUTS
'   SheetName: a SH_ constant.
'   RowNumber, TableColumn: the cell's sheet row and table column.
'   NewValue: the value to write; Empty clears the cell.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim TargetCell   As Range    'The cell

'------------------------------------------------------------------------------
' RECORD AND WRITE
'------------------------------------------------------------------------------
        Set TargetCell = TableCell(GetSheet(SheetName), RowNumber, TableColumn)
        RecordPatch SheetName, TargetCell.Address, TargetCell.Formula
        TargetCell.Value = NewValue

End Sub


Public Sub PatchName( _
    ByVal NameText As String, _
    ByVal NewRefersTo As String)
'
'==============================================================================
'                                  PatchName
'------------------------------------------------------------------------------
' PURPOSE
'   Change what a workbook name refers to, for one case, for example to
'   "=#REF!". Restored like PatchCell.
'
' INPUTS
'   NameText: an existing workbook name, such as a PRM_ code.
'   NewRefersTo: the new reference, as a formula.
'
' UPDATED
'   2026-10-17
'==============================================================================
'
        RecordPatch "", NameText, ThisWorkbook.Names(NameText).RefersTo
        ThisWorkbook.Names(NameText).RefersTo = NewRefersTo

End Sub


Public Sub SetParameter( _
    ByVal Code As String, _
    ByVal ValueText As String)
'
'==============================================================================
'                                 SetParameter
'------------------------------------------------------------------------------
' PURPOSE
'   Give one Params value for the current case, from a fixture's
'   "parameters" object, for example DaysPerYear 360 so that a published
'   maturity of 0.75 years can be entered as a date, or IRBucketOffset
'   FALSE. Restored like PatchCell.
'
' INPUTS
'   Code: a PRM_ code, whose workbook name locates the value cell.
'   ValueText: Y or N for a Boolean, written as TRUE or FALSE; otherwise
'   the number as text with a "." decimal point.
'
' ERROR POLICY
'   Raises ERR_TEST_SETUP when the parameter has no workbook name; the
'   suite then ends as failed.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ValueCell   As Range    'Value cell of the parameter

'------------------------------------------------------------------------------
' WRITE
'------------------------------------------------------------------------------
        Set ValueCell = ParamCell(Code)
        Select Case ValueText
            Case "Y"
                PatchCell SH_PARAMS, ValueCell.Row, ValueCell.Column - COL_OFFSET, True
            Case "N"
                PatchCell SH_PARAMS, ValueCell.Row, ValueCell.Column - COL_OFFSET, False
            Case Else
                PatchCell SH_PARAMS, ValueCell.Row, ValueCell.Column - COL_OFFSET, Val(ValueText)
        End Select

End Sub


Public Sub RunCase()
'
'==============================================================================
'                                   RunCase
'------------------------------------------------------------------------------
' PURPOSE
'   Run the engine on the case's inputs and write every output sheet.
'
' ERROR POLICY
'   A run that stops or raises is a failure of the case; its expected
'   outputs then fail as missing.
'
' UPDATED
'   2026-10-06
'==============================================================================
'

'------------------------------------------------------------------------------
' RUN
'------------------------------------------------------------------------------
        On Error GoTo Failed
        If Not CORE_Engine.Calculate(True) Then
            mFailureCount = mFailureCount + 1
            TEST_Suite.Say "FAILURE=" & mCase & ": engine run stopped; see the Checks sheet"
        End If
        Exit Sub

'------------------------------------------------------------------------------
' HANDLE ERROR
'------------------------------------------------------------------------------
Failed:
        mFailureCount = mFailureCount + 1
        TEST_Suite.Say "FAILURE=" & mCase & ": engine error " & Err.Number & ": " & Err.Description

End Sub


Public Sub RunCaseExpectingStop( _
    ByVal Label As String, _
    ByVal MessagePart As String, _
    ByVal ReferenceClass As String)
'
'==============================================================================
'                             RunCaseExpectingStop
'------------------------------------------------------------------------------
' PURPOSE
'   Run the engine on inputs it must refuse, and check as one result that
'   the run stopped and that the Checks sheet has an ERROR line naming the
'   reason.
'
' INPUTS
'   Label: the check's name in the report.
'   MessagePart: text the ERROR message must contain, ignoring case.
'   ReferenceClass: published, independent or illustrative.
'
' ERROR POLICY
'   A run that raises fails the check.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Completed   As Boolean    'Calculate returned True

'------------------------------------------------------------------------------
' RUN AND CHECK
'------------------------------------------------------------------------------
        On Error GoTo Failed
        Completed = CORE_Engine.Calculate(True)
        If Completed Then
            Record Label, ReferenceClass, False, "run completed; expected it to stop"
        Else
            Record Label, ReferenceClass, ChecksHasLine(SEV_ERROR, MessagePart), _
                   "run stopped, but no ERROR on Checks contains """ & MessagePart & """"
        End If
        Exit Sub

'------------------------------------------------------------------------------
' HANDLE ERROR
'------------------------------------------------------------------------------
Failed:
        Record Label, ReferenceClass, False, "engine error " & Err.Number & ": " & Err.Description

End Sub


'
'------------------------------------------------------------------------------
'
'                                 EXPECTATIONS
'
'------------------------------------------------------------------------------
'

Public Sub ExpectNumber( _
    ByVal Label As String, _
    ByVal TradeId As String, _
    ByVal Quantity As String, _
    ByVal ExpectedText As String, _
    ByVal AbsoluteTolerance As String, _
    ByVal RelativeTolerance As String, _
    ByVal ReferenceClass As String)
'
'==============================================================================
'                                 ExpectNumber
'------------------------------------------------------------------------------
' PURPOSE
'   Check a numeric output: |actual - expected| <= max(absolute,
'   relative * |expected|), as in docs/methodology/TEST_CASES.md. A
'   quantity that cannot be negative fails when it is, however small,
'   before the tolerance is applied.
'
' INPUTS
'   Label: catalogue ID or quantity, for the report.
'   TradeId: the trade for a trade-level quantity; empty for the netting
'   set.
'   Quantity: the quantity name.
'   ExpectedText, AbsoluteTolerance, RelativeTolerance: numbers as text.
'   ReferenceClass: published, independent or illustrative.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Actual                As Variant    'Output cell value
    Dim ExpectedValue         As Double     'Expected value
    Dim PermittedDifference   As Double     'Permitted difference
    Dim Detail                As String     'Diagnostic

'------------------------------------------------------------------------------
' COMPARE
'------------------------------------------------------------------------------
        ExpectedValue = Val(ExpectedText)
        PermittedDifference = Val(AbsoluteTolerance)
        If Val(RelativeTolerance) * Abs(ExpectedValue) > PermittedDifference Then
            PermittedDifference = Val(RelativeTolerance) * Abs(ExpectedValue)
        End If
        If mInjectFailure And Len(mInjectedAt) = 0 Then
            ExpectedValue = ExpectedValue + 10# * PermittedDifference + 1#
            mInjectedAt = mCase & " " & Label
            TEST_Suite.Say "INJECTED=" & mInjectedAt & ": expected value moved to " & Trim$(Str$(ExpectedValue))
        End If
        If Not ReadOutput(TradeId, Quantity, Actual, Detail) Then
            Record Label, ReferenceClass, False, Detail
        ElseIf IsError(Actual) Or IsEmpty(Actual) Or Not IsNumeric(Actual) Or VarType(Actual) = vbString Then
            Record Label, ReferenceClass, False, Quantity & " is not a number"
        ElseIf CDbl(Actual) < 0# And MustBeNonNegative(Quantity) Then
            Record Label, ReferenceClass, False, _
                   Quantity & " must not be negative, actual " & Trim$(Str$(CDbl(Actual)))
        Else
            Record Label, ReferenceClass, Abs(CDbl(Actual) - ExpectedValue) <= PermittedDifference, _
                   Quantity & " expected " & Trim$(Str$(ExpectedValue)) & ", actual " & _
                   Trim$(Str$(CDbl(Actual))) & ", tolerance " & Trim$(Str$(PermittedDifference))
        End If

End Sub


Public Function OutputNumber( _
    ByVal NettingSetKey As String, _
    ByVal Quantity As String) _
    As Double
'
'==============================================================================
'                                 OutputNumber
'------------------------------------------------------------------------------
' PURPOSE
'   Read a netting-set quantity from Results for any netting set of the
'   case, for checks that compare outputs with each other rather than with
'   a fixed expected value.
'
' INPUTS
'   NettingSetKey: the netting set's ID, which names its row on Results.
'   Quantity: a netting-set quantity name, as for ExpectNumber.
'
' ERROR POLICY
'   Raises ERR_TEST_SETUP when the row, the column or a number is missing;
'   the calling suite then fails.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ResultsSheet     As Worksheet    'Results sheet
    Dim QuantityColumn   As Long         'Output column; 0 when unknown
    Dim ResultRow        As Long         'Output row; 0 when not found
    Dim Content          As Variant      'Cell value

'------------------------------------------------------------------------------
' READ
'------------------------------------------------------------------------------
        Set ResultsSheet = GetSheet(SH_RESULTS)
        QuantityColumn = OutputColumn(Quantity, False)
        ResultRow = FindHeaderRow(ResultsSheet, 1, NettingSetKey)
        If QuantityColumn = 0 Or ResultRow < OutputFirstRow(SH_RESULTS) Then
            Err.Raise ERR_TEST_SETUP, "TEST_CaseRunner.OutputNumber", _
                      Quantity & " of " & NettingSetKey & " not found on Results."
        End If
        Content = TableCell(ResultsSheet, ResultRow, QuantityColumn).Value
        If Not IsNum(Content) Then
            Err.Raise ERR_TEST_SETUP, "TEST_CaseRunner.OutputNumber", _
                      Quantity & " of " & NettingSetKey & " is not a number (" & SafeStr(Content) & ")."
        End If
        OutputNumber = CDbl(Content)

End Function


Public Sub ExpectTrue( _
    ByVal Label As String, _
    ByVal Passed As Boolean, _
    ByVal Detail As String, _
    ByVal ReferenceClass As String)
'
'==============================================================================
'                                  ExpectTrue
'------------------------------------------------------------------------------
' PURPOSE
'   Record a check that the caller has evaluated itself, such as a relation
'   between two outputs.
'
' INPUTS
'   Label: the check's name in the report.
'   Passed: the result.
'   Detail: printed when the check fails.
'   ReferenceClass: published, independent or illustrative.
'
' UPDATED
'   2026-10-17
'==============================================================================
'
        Record Label, ReferenceClass, Passed, Detail

End Sub


Public Sub ExpectChecksLine( _
    ByVal Label As String, _
    ByVal Severity As String, _
    ByVal MessagePart As String, _
    ByVal ReferenceClass As String)
'
'==============================================================================
'                               ExpectChecksLine
'------------------------------------------------------------------------------
' PURPOSE
'   After RunCase, check that the Checks sheet has a line of the given
'   Severity (SEV_ERROR, SEV_WARN or SEV_INFO) whose message contains
'   MessagePart, ignoring case: for an ERROR, WARNING or INFO line that the
'   netting set's status alone does not show.
'
' UPDATED
'   2026-10-21
'==============================================================================
'
        Record Label, ReferenceClass, ChecksHasLine(Severity, MessagePart), _
               "Checks has no " & Severity & " line containing '" & MessagePart & "'"

End Sub


Public Sub ExpectText( _
    ByVal Label As String, _
    ByVal TradeId As String, _
    ByVal Quantity As String, _
    ByVal ExpectedText As String, _
    ByVal ReferenceClass As String)
'
'==============================================================================
'                                  ExpectText
'------------------------------------------------------------------------------
' PURPOSE
'   Check a text output for exact, case-sensitive equality.
'
' INPUTS
'   As ExpectNumber; ExpectedText is the exact text.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Actual   As Variant    'Output cell value
    Dim Detail   As String     'Diagnostic

'------------------------------------------------------------------------------
' COMPARE
'------------------------------------------------------------------------------
        If Not ReadOutput(TradeId, Quantity, Actual, Detail) Then
            Record Label, ReferenceClass, False, Detail
        ElseIf IsError(Actual) Then
            Record Label, ReferenceClass, False, Quantity & " is an error value"
        Else
            Record Label, ReferenceClass, StrComp(CStr(Actual), ExpectedText, vbBinaryCompare) = 0, _
                   Quantity & " expected """ & ExpectedText & """, actual """ & CStr(Actual) & """"
        End If

End Sub


'
'------------------------------------------------------------------------------
'
'                                   HELPERS
'
'------------------------------------------------------------------------------
'

Private Function ReadOutput( _
    ByVal TradeId As String, _
    ByVal Quantity As String, _
    ByRef Actual As Variant, _
    ByRef Detail As String) _
    As Boolean
'
'==============================================================================
'                                  ReadOutput
'------------------------------------------------------------------------------
' PURPOSE
'   Find the output cell of a quantity: on Results for the netting set, on
'   TradeCalc for a trade.
'
' RETURNS
'   True with Actual set to the cell value; False with Detail saying what
'   was missing.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim OutputSheet      As Worksheet    'Results or TradeCalc
    Dim RowId            As String       'Row ID to find
    Dim QuantityColumn   As Long         'Output column; 0 when unknown
    Dim OutputRow        As Long         'Output row; 0 when not found

'------------------------------------------------------------------------------
' LOCATE
'------------------------------------------------------------------------------
        If Len(TradeId) = 0 Then
            Set OutputSheet = GetSheet(SH_RESULTS)
            RowId = mNettingSetId
        Else
            Set OutputSheet = GetSheet(SH_TRADECALC)
            RowId = TradeId
        End If
        QuantityColumn = OutputColumn(Quantity, Len(TradeId) > 0)
        If QuantityColumn = 0 Then
            Detail = "no output column for " & Quantity
            Exit Function
        End If
        OutputRow = FindHeaderRow(OutputSheet, 1, RowId)
        If OutputRow < OutputFirstRow(OutputSheet.Name) Then
            Detail = RowId & " not found on " & OutputSheet.Name
            Exit Function
        End If
        Actual = TableCell(OutputSheet, OutputRow, QuantityColumn).Value
        ReadOutput = True

End Function


Private Function MustBeNonNegative( _
    ByVal Quantity As String) _
    As Boolean
'
'==============================================================================
'                              MustBeNonNegative
'------------------------------------------------------------------------------
' PURPOSE
'   Tell whether a quantity is non-negative by definition, so that a
'   negative result fails whatever the tolerance (TEST_CASES.md).
'   Supervisory delta carries the trade's sign and is excluded; adjusted
'   notional, never negative, is not listed either.
'
' UPDATED
'   2026-10-21
'==============================================================================
'
        Select Case Quantity
            Case "exposure_value", "replacement_cost", "potential_future_exposure", _
                 "aggregate_add_on", "multiplier", "margin_period_of_risk", _
                 "supervisory_factor", "lambda_shift"
                MustBeNonNegative = True
            Case Else
                MustBeNonNegative = (Left$(Quantity, 7) = "add_on.")
        End Select

End Function


Private Function OutputColumn( _
    ByVal Quantity As String, _
    ByVal TradeLevel As Boolean) _
    As Long
'
'==============================================================================
'                                 OutputColumn
'------------------------------------------------------------------------------
' PURPOSE
'   Map a quantity name to its column on Results or TradeCalc.
'
' RETURNS
'   The column number; 0 for an unknown quantity.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' MAP
'------------------------------------------------------------------------------
        If TradeLevel Then
            Select Case Quantity
                Case "trade_status"
                    OutputColumn = 4
                Case "hedging_set"
                    OutputColumn = 5
                Case "bucket"
                    OutputColumn = 6
                Case "supervisory_factor"
                    OutputColumn = 7
                Case "adjusted_notional"
                    OutputColumn = 15
                Case "supervisory_delta"
                    OutputColumn = 16
                Case "lambda_shift"
                    OutputColumn = 23
            End Select
        Else
            Select Case Quantity
                Case "margin_period_of_risk"
                    OutputColumn = 6
                Case "replacement_cost"
                    OutputColumn = 9
                Case "add_on.interest_rate"
                    OutputColumn = 10
                Case "add_on.foreign_exchange"
                    OutputColumn = 11
                Case "add_on.credit"
                    OutputColumn = 12
                Case "add_on.equity"
                    OutputColumn = 13
                Case "add_on.commodity"
                    OutputColumn = 14
                Case "add_on.other"
                    OutputColumn = 15
                Case "aggregate_add_on"
                    OutputColumn = 16
                Case "multiplier"
                    OutputColumn = 17
                Case "potential_future_exposure"
                    OutputColumn = 18
                Case "exposure_value"
                    OutputColumn = 22
                Case "cap_applied"
                    OutputColumn = 23
                Case "netting_set_status"
                    OutputColumn = RS_STATUS_COL
            End Select
        End If

End Function


Private Sub Record( _
    ByVal Label As String, _
    ByVal ReferenceClass As String, _
    ByVal Passed As Boolean, _
    ByVal Detail As String)
'
'==============================================================================
'                                    Record
'------------------------------------------------------------------------------
' PURPOSE
'   Count one check under its reference class and print it when it fails.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ClassIndex   As Long    'Reference-class index

'------------------------------------------------------------------------------
' COUNT
'------------------------------------------------------------------------------
        Select Case ReferenceClass
            Case "published"
                ClassIndex = CLASS_PUBLISHED
            Case "independent"
                ClassIndex = CLASS_INDEPENDENT
            Case Else
                ClassIndex = CLASS_ILLUSTRATIVE
        End Select
        mCheckCount = mCheckCount + 1
        If Passed Then
            mPassed(ClassIndex) = mPassed(ClassIndex) + 1
        Else
            mFailed(ClassIndex) = mFailed(ClassIndex) + 1
            mFailureCount = mFailureCount + 1
            TEST_Suite.Say "FAILURE=" & mCase & " " & Label & " (" & ReferenceClass & "): " & Detail
        End If

End Sub


Private Function ChecksHasLine( _
    ByVal Severity As String, _
    ByVal MessagePart As String) _
    As Boolean
'
'==============================================================================
'                                ChecksHasLine
'------------------------------------------------------------------------------
' PURPOSE
'   Tell whether the Checks sheet has a line of the given severity whose
'   message contains MessagePart, ignoring case.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ChecksSheet   As Worksheet    'Checks sheet
    Dim LastUsedRow   As Long         'Last used row
    Dim r             As Long         'Row being read

'------------------------------------------------------------------------------
' SEARCH
'------------------------------------------------------------------------------
    'Checks columns: 1 severity, 5 message.
        Set ChecksSheet = GetSheet(SH_CHECKS)
        LastUsedRow = UsedLastRow(ChecksSheet)
        For r = OutputFirstRow(SH_CHECKS) To LastUsedRow
            If SafeStr(TableCell(ChecksSheet, r, 1).Value) = Severity Then
                If InStr(1, SafeStr(TableCell(ChecksSheet, r, 5).Value), MessagePart, vbTextCompare) > 0 Then
                    ChecksHasLine = True
                    Exit Function
                End If
            End If
        Next r

End Function


Private Sub UnitFxRate( _
    ByVal CurrencyCode As String)
'
'==============================================================================
'                                  UnitFxRate
'------------------------------------------------------------------------------
' PURPOSE
'   Set the FX-table rate of the case's calculation currency to 1 through
'   PatchCell, so that it is restored with the other patches. A currency
'   missing from the table needs nothing: the engine adds the reporting
'   currency at rate 1.
'
' INPUTS
'   CurrencyCode: three-letter code, as in the first column of the FX
'   table.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ParamsSheet   As Worksheet    'Params sheet
    Dim CurrencyRow   As Long         'Row of the currency in the table; 0 when absent

'------------------------------------------------------------------------------
' PATCH
'------------------------------------------------------------------------------
        Set ParamsSheet = GetSheet(SH_PARAMS)
        CurrencyRow = FindHeaderRow(ParamsSheet, NamedTableCol(NM_FX_TABLE), CurrencyCode)
        If CurrencyRow > 0 Then
            If ToDbl(TableCell(ParamsSheet, CurrencyRow, NamedTableCol(NM_FX_TABLE) + 2).Value) <> 1# Then
                PatchCell SH_PARAMS, CurrencyRow, NamedTableCol(NM_FX_TABLE) + 2, 1#
            End If
        End If

End Sub


Private Sub RecordPatch( _
    ByVal SheetName As String, _
    ByVal Target As String, _
    ByVal Saved As Variant)
'
'==============================================================================
'                                 RecordPatch
'------------------------------------------------------------------------------
' PURPOSE
'   Remember a cell's formula or a workbook name's reference before
'   PatchCell or PatchName changes it.
'
' UPDATED
'   2026-10-17
'==============================================================================
'
        mPatchCount = mPatchCount + 1
        ReDim Preserve mPatchSheet(1 To mPatchCount)
        ReDim Preserve mPatchTarget(1 To mPatchCount)
        ReDim Preserve mPatchSaved(1 To mPatchCount)
        mPatchSheet(mPatchCount) = SheetName
        mPatchTarget(mPatchCount) = Target
        mPatchSaved(mPatchCount) = Saved

End Sub


Private Sub RestorePatches()
'
'==============================================================================
'                                RestorePatches
'------------------------------------------------------------------------------
' PURPOSE
'   Undo every recorded patch, newest first, so a cell patched twice gets
'   its original value back.
'
' ERROR POLICY
'   An error propagates: to the suite's handler from BeginCase, or to
'   RestoreStep from EndSuite. The patches are forgotten first, so a failed
'   restoration is reported once.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim i                 As Long    'Patch being undone
    Dim RecordedPatches   As Long    'Patches recorded

'------------------------------------------------------------------------------
' RESTORE
'------------------------------------------------------------------------------
        RecordedPatches = mPatchCount
        mPatchCount = 0
        For i = RecordedPatches To 1 Step -1
            If Len(mPatchSheet(i)) = 0 Then
                ThisWorkbook.Names(mPatchTarget(i)).RefersTo = mPatchSaved(i)
            Else
                GetSheet(mPatchSheet(i)).Range(mPatchTarget(i)).Formula = mPatchSaved(i)
            End If
        Next i

End Sub


Private Function RestoreStep( _
    ByVal StepName As String, _
    ByVal StepNumber As Long) _
    As Boolean
'
'==============================================================================
'                                 RestoreStep
'------------------------------------------------------------------------------
' PURPOSE
'   Run one restoration step of EndSuite with its own error handler.
'
' INPUTS
'   StepName: name for the report.
'   StepNumber: 1 inputs and patches, 2 parameters, 3 outputs,
'   4 calculation mode, 5 events, 6 screen updating. Each setting is its
'   own step, so one failed restoration cannot leave the others changed.
'
' RETURNS
'   True when the step succeeded.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' RESTORE
'------------------------------------------------------------------------------
        On Error GoTo Failed
        Select Case StepNumber
            Case 1
                RestorePatches
                RestoreBlock GetSheet(SH_NS), NS_NCOLS, mSavedNsAddress, mSavedNs
                RestoreBlock GetSheet(SH_TRADES), TR_NCOLS, mSavedTrAddress, mSavedTr
            Case 2
                mAsOfCell.Formula = mSavedAsOf
                mCcyCell.Formula = mSavedCcy
            Case 3
                CORE_Engine.Calculate True
            Case 4
                Application.Calculation = mSavedCalculation
            Case 5
                Application.EnableEvents = mSavedEvents
            Case 6
                Application.ScreenUpdating = mSavedScreen
        End Select
        RestoreStep = True
        Exit Function

'------------------------------------------------------------------------------
' REPORT FAILURE
'------------------------------------------------------------------------------
Failed:
        TEST_Suite.Say "FAILURE=restore " & StepName & ": " & Err.Number & ": " & Err.Description

End Function


Private Sub SaveBlock( _
    ByVal InputSheet As Worksheet, _
    ByVal ColumnCount As Long, _
    ByRef BlockAddress As String, _
    ByRef Formulas As Variant)
'
'==============================================================================
'                                  SaveBlock
'------------------------------------------------------------------------------
' PURPOSE
'   Save the formulas of an input block, from the first data row to the
'   last used row.
'
' RETURNS
'   BlockAddress: the block's address, "" when the sheet has no data rows.
'   Formulas: the block's formulas.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim LastUsedRow   As Long    'Last used row
    Dim TopRow        As Long    'First data row

'------------------------------------------------------------------------------
' SAVE
'------------------------------------------------------------------------------
        BlockAddress = ""
        Formulas = Empty
        TopRow = InputTop(InputSheet.Name)
        LastUsedRow = UsedLastRow(InputSheet)
        If LastUsedRow < TopRow Then
            Exit Sub
        End If
        With TableRange(InputSheet, TopRow, 1, LastUsedRow, ColumnCount)
            BlockAddress = .Address
            Formulas = .Formula
        End With

End Sub


Private Sub RestoreBlock( _
    ByVal InputSheet As Worksheet, _
    ByVal ColumnCount As Long, _
    ByVal BlockAddress As String, _
    ByVal Formulas As Variant)
'
'==============================================================================
'                                 RestoreBlock
'------------------------------------------------------------------------------
' PURPOSE
'   Clear the case rows and write the saved block back.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' RESTORE
'------------------------------------------------------------------------------
        ClearBlock InputSheet, ColumnCount
        If Len(BlockAddress) > 0 Then
            InputSheet.Range(BlockAddress).Formula = Formulas
        End If

End Sub


Private Sub ClearBlock( _
    ByVal InputSheet As Worksheet, _
    ByVal ColumnCount As Long)
'
'==============================================================================
'                                  ClearBlock
'------------------------------------------------------------------------------
' PURPOSE
'   Clear the values of an input block; formats and validation stay.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim LastUsedRow   As Long    'Last used row
    Dim TopRow        As Long    'First data row

'------------------------------------------------------------------------------
' CLEAR
'------------------------------------------------------------------------------
        TopRow = InputTop(InputSheet.Name)
        LastUsedRow = UsedLastRow(InputSheet)
        If LastUsedRow >= TopRow Then
            TableRange(InputSheet, TopRow, 1, LastUsedRow, ColumnCount).ClearContents
        End If

End Sub


Private Function InputTop( _
    ByVal SheetName As String) _
    As Long
'
'==============================================================================
'                                   InputTop
'------------------------------------------------------------------------------
' PURPOSE
'   Return the first data row of an input sheet, located through the
'   workbook name of its first header cell (CORE_Util.TableFirstRow).
'
' ERROR POLICY
'   Raises ERR_TEST_SETUP when the name cannot locate it; the suite then
'   ends as failed.
'
' UPDATED
'   2026-10-17
'==============================================================================
'
        InputTop = TableFirstRow(SheetName)
        If InputTop = 0 Then
            Err.Raise ERR_TEST_SETUP, "TEST_CaseRunner.InputTop", _
                      "The header of " & SheetName & " cannot be located through its workbook name - " & _
                      "rebuild the workbook from the template."
        End If

End Function


Private Sub WriteRow( _
    ByVal InputSheet As Worksheet, _
    ByVal RowNumber As Long, _
    ByVal ColumnCount As Long, _
    ByRef Values() As Variant)
'
'==============================================================================
'                                   WriteRow
'------------------------------------------------------------------------------
' PURPOSE
'   Write one input row from a (1 To 1, 1 To ColumnCount) array.
'
' UPDATED
'   2026-10-17
'==============================================================================
'
        TableRange(InputSheet, RowNumber, 1, RowNumber, ColumnCount).Value = Values

End Sub


Private Function ParamCell( _
    ByVal Code As String) _
    As Range
'
'==============================================================================
'                                  ParamCell
'------------------------------------------------------------------------------
' PURPOSE
'   Return the cell the engine reads a parameter from, as CORE_Util.GetParam
'   does: the top-left cell of the workbook name equal to its code.
'
' ERROR POLICY
'   Raises ERR_TEST_SETUP when the name is missing or broken.
'
' UPDATED
'   2026-10-17
'==============================================================================
'
        Set ParamCell = NamedCell(Code)
        If ParamCell Is Nothing Then
            Err.Raise ERR_TEST_SETUP, "TEST_CaseRunner.ParamCell", "Parameter " & Code & " has no workbook name."
        End If

End Function


Private Function NumberOrBlank( _
    ByVal InputText As String) _
    As Variant
'
'==============================================================================
'                                NumberOrBlank
'------------------------------------------------------------------------------
' PURPOSE
'   Convert a number written with a "." decimal point; empty means blank.
'   Val ignores the Windows locale, unlike CDbl.
'
' UPDATED
'   2026-10-17
'==============================================================================
'
        If Len(InputText) = 0 Then
            NumberOrBlank = Empty
        Else
            NumberOrBlank = Val(InputText)
        End If

End Function


Private Function TextOrBlank( _
    ByVal InputText As String) _
    As Variant
'
'==============================================================================
'                                 TextOrBlank
'------------------------------------------------------------------------------
' PURPOSE
'   Return text for a cell; empty means blank.
'
' UPDATED
'   2026-10-17
'==============================================================================
'
        If Len(InputText) = 0 Then
            TextOrBlank = Empty
        Else
            TextOrBlank = InputText
        End If

End Function


Private Function IsoDate( _
    ByVal InputText As String) _
    As Variant
'
'==============================================================================
'                                   IsoDate
'------------------------------------------------------------------------------
' PURPOSE
'   Convert YYYY-MM-DD to a date without depending on the locale; empty
'   means blank.
'
' UPDATED
'   2026-10-17
'==============================================================================
'
        If Len(InputText) = 0 Then
            IsoDate = Empty
        Else
            IsoDate = DateSerial(CInt(Left$(InputText, 4)), CInt(Mid$(InputText, 6, 2)), CInt(Right$(InputText, 2)))
        End If

End Function
