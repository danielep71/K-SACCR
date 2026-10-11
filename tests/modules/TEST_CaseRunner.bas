Attribute VB_Name = "TEST_CaseRunner"
'==============================================================================
' MODULE: TEST_CaseRunner
'------------------------------------------------------------------------------
' PURPOSE
'   Run the numerical test cases of tests/fixtures and tests/expected
'   against the engine. The generated module TEST_Cases calls these
'   procedures: for each case it writes the fixture into the input sheets,
'   runs the engine, and compares the output cells with the expected values
'   (methodology decision 7, #44).
'
' PUBLIC SURFACE
'   InjectFailure, BeginSuite, BeginCase, LoadSavedInputs, SetParameter,
'   AddNettingSet,
'   AddTrade, SetInputCell, PatchCell, PatchName, RunCase,
'   RunCaseExpectingStop,
'   OutputNumber, ExpectNumber, ExpectText, ExpectTrue and EndSuite, for
'   TEST_Cases, TEST_InputValidation, TEST_Aggregation and TEST_Invariants.
'   Option Private Module keeps them out of the external workbook automation
'   API.
'
' DEPENDENCIES
'   CORE_Engine.Calculate; CORE_Util for sheet access; CORE_Config for the layout.
'   The workbook must be built from the template, because the engine reads
'   its input sheets and Params.
'
' WORKSHEET SAFETY
'   Unlike TEST_Harness, this module writes the NettingSets and Trades input
'   rows, the AsOfDate and ReportingCcy parameters, and every output sheet.
'   BeginSuite saves the inputs and parameters and EndSuite writes them
'   back and recalculates, so the workbook ends as it started. A cell or
'   workbook name changed with PatchCell or PatchName is restored at the
'   start of the next case and by EndSuite. Use a development workbook.
'
' STATE OWNERSHIP
'   Owns the counters and saved inputs below. BeginSuite captures and
'   EndSuite restores calculation mode, events and screen updating.
'
' ERROR POLICY
'   A failed engine run or a missing output is a failed check, and the
'   suite continues. Restoration steps run independently and a failed one
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
'   2026-10-10
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
        Private mExpectedCases   As Long       'Cases TEST_Cases will run
        Private mExpectedChecks  As Long       'Checks TEST_Cases will run
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
    ByVal expectedCases As Long, _
    ByVal expectedChecks As Long)
'
'==============================================================================
'                                  BeginSuite
'------------------------------------------------------------------------------
' PURPOSE
'   Save everything the cases will overwrite, then prepare Excel.
'
' INPUTS
'   expectedCases, expectedChecks: numbers of cases and checks TEST_Cases
'   will run; a different count fails the suite, so a stale or edited
'   TEST_Cases module cannot pass with checks missing.
'
' ERROR POLICY
'   Raises if a suite is already active or the inputs cannot be saved;
'   nothing has been changed at that point.
'
' UPDATED
'   2026-10-09
'==============================================================================
'

'------------------------------------------------------------------------------
' GUARD AND RESET
'------------------------------------------------------------------------------
        If mSuiteActive Then
            Err.Raise ERR_RUN_ACTIVE, "TEST_CaseRunner.BeginSuite", "A case suite is already running."
        End If
        mExpectedCases = expectedCases
        mExpectedChecks = expectedChecks
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
        Debug.Print "SACCR CASE TESTS"
        Debug.Print "MODE=" & IIf(mInjectFailure, "INJECTED_FAILURE", "NORMAL")

End Sub


Public Sub EndSuite( _
    ByVal abortReason As String)
'
'==============================================================================
'                                   EndSuite
'------------------------------------------------------------------------------
' PURPOSE
'   Restore the inputs, parameters, outputs and Excel settings, and print
'   the summary and the RESULT line.
'
' INPUTS
'   abortReason: why TEST_Cases stopped early; empty after a complete run.
'
' ERROR POLICY
'   Each restoration step contains its own error; a failed one is printed
'   and makes the result FAIL.
'
' UPDATED
'   2026-10-09
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim restored   As Boolean    'Every restoration step succeeded
    Dim complete   As Boolean    'Every expected case ran

'------------------------------------------------------------------------------
' RESTORE
'------------------------------------------------------------------------------
    'The engine is run once more on the restored inputs so the output
    'sheets match them again.
        If Len(abortReason) > 0 Then
            mFailureCount = mFailureCount + 1
            Debug.Print "FAILURE=suite: " & abortReason
        End If
        restored = True
        restored = RestoreStep("inputs", 1) And restored
        restored = RestoreStep("parameters", 2) And restored
        restored = RestoreStep("outputs", 3) And restored
        restored = RestoreStep("calculation", 4) And restored
        restored = RestoreStep("events", 5) And restored
        restored = RestoreStep("screen updating", 6) And restored
        mSuiteActive = False

'------------------------------------------------------------------------------
' REPORT
'------------------------------------------------------------------------------
        If mInjectFailure And Len(mInjectedAt) = 0 Then
            mFailureCount = mFailureCount + 1
            Debug.Print "FAILURE=suite: injected failure requested, but the suite ran no numeric check"
        End If
        mInjectFailure = False
        complete = (mCaseCount = mExpectedCases) And (mCheckCount = mExpectedChecks)
        If Not complete Then
            mFailureCount = mFailureCount + 1
            Debug.Print "FAILURE=suite: expected " & mExpectedCases & " cases and " & mExpectedChecks & _
                        " checks, ran " & mCaseCount & " and " & mCheckCount
        End If
        Debug.Print "PUBLISHED: passed=" & mPassed(CLASS_PUBLISHED) & "; failed=" & mFailed(CLASS_PUBLISHED)
        Debug.Print "INDEPENDENT: passed=" & mPassed(CLASS_INDEPENDENT) & "; failed=" & mFailed(CLASS_INDEPENDENT)
        Debug.Print "ILLUSTRATIVE: passed=" & mPassed(CLASS_ILLUSTRATIVE) & "; failed=" & _
                    mFailed(CLASS_ILLUSTRATIVE) & " (validates nothing)"
        Debug.Print "RESULT=" & IIf(mFailureCount = 0 And restored And complete, "PASS", "FAIL") & _
                    "; cases=" & mCaseCount & "; checks=" & mCheckCount & _
                    "; failures=" & mFailureCount & "; restore=" & IIf(restored, "PASS", "FAIL")

End Sub


'
'------------------------------------------------------------------------------
'
'                                CASE BUILDING
'
'------------------------------------------------------------------------------
'

Public Sub BeginCase( _
    ByVal caseId As String, _
    ByVal regimeCode As String, _
    ByVal valuationDate As String, _
    ByVal calculationCurrency As String)
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
'   caseId: the fixture ID.
'   regimeCode: CRR or BCBS, applied as the netting set's regime override.
'   valuationDate: YYYY-MM-DD.
'   calculationCurrency: the fixture's calculation currency.
'
' UPDATED
'   2026-10-09
'==============================================================================
'

'------------------------------------------------------------------------------
' RESET INPUTS
'------------------------------------------------------------------------------
        RestorePatches
        mCaseCount = mCaseCount + 1
        mCase = caseId & "." & LCase$(regimeCode)
        mRegime = regimeCode
        mCurrency = calculationCurrency
        mNettingSetId = ""
        mNextTradeRow = FIRST_DATA_ROW
        ClearBlock GetSheet(SH_NS), NS_NCOLS
        ClearBlock GetSheet(SH_TRADES), TR_NCOLS
        mAsOfCell.Value = IsoDate(valuationDate)
        mCcyCell.Value = calculationCurrency
        UnitFxRate calculationCurrency
        Debug.Print "CASE=" & mCase

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
    ByVal nettingSetId As String, _
    ByVal marginedFlag As String, _
    ByVal roleCode As String, _
    ByVal remarginDays As String, _
    ByVal largeFlag As String, _
    ByVal disputesFlag As String, _
    ByVal mporOverrideDays As String, _
    ByVal vmAmount As String, _
    ByVal nicaAmount As String, _
    ByVal thresholdAmount As String, _
    ByVal mtaAmount As String, _
    ByVal alphaFactor As String)
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
'   2026-10-10
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim rowValues(1 To 1, 1 To NS_NCOLS)   As Variant    'The row to write

'------------------------------------------------------------------------------
' WRITE
'------------------------------------------------------------------------------
        mNettingSetId = nettingSetId
        rowValues(1, NS_ID) = nettingSetId
        rowValues(1, NS_CPTY) = "Test case " & mCase
        rowValues(1, NS_MARGINED) = marginedFlag
        rowValues(1, NS_CLEARED) = roleCode
        rowValues(1, NS_FREQ) = NumberOrBlank(remarginDays)
        rowValues(1, NS_LARGE) = largeFlag
        rowValues(1, NS_DISPUTE) = disputesFlag
        rowValues(1, NS_MPOR) = NumberOrBlank(mporOverrideDays)
        rowValues(1, NS_VM) = NumberOrBlank(vmAmount)
        rowValues(1, NS_NICA) = NumberOrBlank(nicaAmount)
        rowValues(1, NS_TH) = NumberOrBlank(thresholdAmount)
        rowValues(1, NS_MTA) = NumberOrBlank(mtaAmount)
        rowValues(1, NS_ALPHA) = NumberOrBlank(alphaFactor)
        rowValues(1, NS_REGIME) = mRegime
        WriteRow GetSheet(SH_NS), FIRST_DATA_ROW, NS_NCOLS, rowValues

End Sub


Public Sub AddTrade( _
    ByVal tradeId As String, _
    ByVal assetClassCode As String, _
    ByVal subClass As String, _
    ByVal riskFactor As String, _
    ByVal instrument As String, _
    ByVal direction As String, _
    ByVal optionType As String, _
    ByVal nature As String, _
    ByVal hedgingSetLabel As String, _
    ByVal notionalAmount As String, _
    ByVal marketValueAmount As String, _
    ByVal startDate As String, _
    ByVal endDate As String, _
    ByVal maturityDate As String, _
    ByVal expiryDate As String, _
    ByVal underlyingPrice As String, _
    ByVal strikePrice As String, _
    ByVal lambdaPrice As String, _
    ByVal attachmentRate As String, _
    ByVal detachmentRate As String)
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
'   2026-10-06
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim rowValues(1 To 1, 1 To TR_NCOLS)   As Variant    'The row to write

'------------------------------------------------------------------------------
' WRITE
'------------------------------------------------------------------------------
        rowValues(1, TR_ID) = tradeId
        rowValues(1, TR_NS) = mNettingSetId
        rowValues(1, TR_AC) = assetClassCode
        rowValues(1, TR_SUB) = TextOrBlank(subClass)
        rowValues(1, TR_RF) = riskFactor
        rowValues(1, TR_INSTR) = instrument
        rowValues(1, TR_DIR) = direction
        rowValues(1, TR_OPT) = TextOrBlank(optionType)
        rowValues(1, TR_NATURE) = nature
        rowValues(1, TR_LABEL) = TextOrBlank(hedgingSetLabel)
        rowValues(1, TR_NOTIONAL) = NumberOrBlank(notionalAmount)
        rowValues(1, TR_NCCY) = mCurrency
        rowValues(1, TR_MTM) = NumberOrBlank(marketValueAmount)
        rowValues(1, TR_MCCY) = mCurrency
        rowValues(1, TR_START) = IsoDate(startDate)
        rowValues(1, TR_END) = IsoDate(endDate)
        rowValues(1, TR_MAT) = IsoDate(maturityDate)
        rowValues(1, TR_EXPIRY) = IsoDate(expiryDate)
        rowValues(1, TR_PRICE) = NumberOrBlank(underlyingPrice)
        rowValues(1, TR_STRIKE) = NumberOrBlank(strikePrice)
        rowValues(1, TR_LAMBDA) = NumberOrBlank(lambdaPrice)
        rowValues(1, TR_ATTACH) = NumberOrBlank(attachmentRate)
        rowValues(1, TR_DETACH) = NumberOrBlank(detachmentRate)
        WriteRow GetSheet(SH_TRADES), mNextTradeRow, TR_NCOLS, rowValues
        mNextTradeRow = mNextTradeRow + 1

End Sub


Public Sub SetInputCell( _
    ByVal sheetName As String, _
    ByVal dataRow As Long, _
    ByVal col As Long, _
    ByVal newValue As Variant)
'
'==============================================================================
'                                 SetInputCell
'------------------------------------------------------------------------------
' PURPOSE
'   Overwrite one input cell after AddNettingSet or AddTrade, for invalid
'   inputs that a JSON fixture cannot express: a blank ID, a duplicate ID,
'   a typo in a flag. Used by TEST_InputValidation.
'
' INPUTS
'   sheetName: SH_NS or SH_TRADES.
'   dataRow: 1 for the first data row.
'   col: column number, for example TR_ID.
'   newValue: the value to write; Empty clears the cell.
'
' UPDATED
'   2026-10-07
'==============================================================================
'
        GetSheet(sheetName).Cells(FIRST_DATA_ROW + dataRow - 1, col).Value = newValue

End Sub


Public Sub PatchCell( _
    ByVal sheetName As String, _
    ByVal rowNum As Long, _
    ByVal col As Long, _
    ByVal newValue As Variant)
'
'==============================================================================
'                                  PatchCell
'------------------------------------------------------------------------------
' PURPOSE
'   Overwrite any cell outside the input rows, such as a header or a row of
'   a Params table, for one case. The cell is restored at the start of the
'   next case and by EndSuite. Used by TEST_InputValidation.
'
' INPUTS
'   sheetName: a SH_ constant.
'   rowNum, col: the cell's sheet row and column.
'   newValue: the value to write; Empty clears the cell.
'
' UPDATED
'   2026-10-07
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim target   As Range    'The cell

'------------------------------------------------------------------------------
' RECORD AND WRITE
'------------------------------------------------------------------------------
        Set target = GetSheet(sheetName).Cells(rowNum, col)
        RecordPatch sheetName, target.Address, target.Formula
        target.Value = newValue

End Sub


Public Sub PatchName( _
    ByVal nameText As String, _
    ByVal newRefersTo As String)
'
'==============================================================================
'                                  PatchName
'------------------------------------------------------------------------------
' PURPOSE
'   Change what a workbook name refers to, for one case, for example to
'   "=#REF!". Restored like PatchCell.
'
' INPUTS
'   nameText: an existing workbook name, such as a PRM_ code.
'   newRefersTo: the new reference, as a formula.
'
' UPDATED
'   2026-10-07
'==============================================================================
'
        RecordPatch "", nameText, ThisWorkbook.Names(nameText).RefersTo
        ThisWorkbook.Names(nameText).RefersTo = newRefersTo

End Sub


Public Sub SetParameter( _
    ByVal code As String, _
    ByVal valueText As String)
'
'==============================================================================
'                                 SetParameter
'------------------------------------------------------------------------------
' PURPOSE
'   Give one Params value for the current case, from a fixture's
'   "parameters" object, for example DaysPerYear 360 so that a published
'   maturity of 0.75 years can be entered as a date. Restored like
'   PatchCell.
'
' INPUTS
'   code: a PRM_ code in column A of Params.
'   valueText: the number as text with a "." decimal point.
'
' ERROR POLICY
'   Raises ERR_TEST_SETUP when the code is not on Params; the suite then
'   ends as failed.
'
' UPDATED
'   2026-10-09
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim r   As Long    'Params row of the code; 0 when absent

'------------------------------------------------------------------------------
' WRITE
'------------------------------------------------------------------------------
        r = FindHeaderRow(GetSheet(SH_PARAMS), PRM_CODE_COL, code)
        If r = 0 Then
            Err.Raise ERR_TEST_SETUP, "TEST_CaseRunner.SetParameter", _
                      "'" & code & "' not found in column A of Params."
        End If
        PatchCell SH_PARAMS, r, PRM_VALUE_COL, Val(valueText)

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
            Debug.Print "FAILURE=" & mCase & ": engine run stopped; see the Checks sheet"
        End If
        Exit Sub

'------------------------------------------------------------------------------
' HANDLE ERROR
'------------------------------------------------------------------------------
Failed:
        mFailureCount = mFailureCount + 1
        Debug.Print "FAILURE=" & mCase & ": engine error " & Err.Number & ": " & Err.Description

End Sub


Public Sub RunCaseExpectingStop( _
    ByVal label As String, _
    ByVal messagePart As String, _
    ByVal referenceClass As String)
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
'   label: the check's name in the report.
'   messagePart: text the ERROR message must contain, ignoring case.
'   referenceClass: published, independent or illustrative.
'
' ERROR POLICY
'   A run that raises fails the check.
'
' UPDATED
'   2026-10-07
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim completed   As Boolean    'Calculate returned True

'------------------------------------------------------------------------------
' RUN AND CHECK
'------------------------------------------------------------------------------
        On Error GoTo Failed
        completed = CORE_Engine.Calculate(True)
        If completed Then
            Record label, referenceClass, False, "run completed; expected it to stop"
        Else
            Record label, referenceClass, ChecksHasError(messagePart), _
                   "run stopped, but no ERROR on Checks contains """ & messagePart & """"
        End If
        Exit Sub

'------------------------------------------------------------------------------
' HANDLE ERROR
'------------------------------------------------------------------------------
Failed:
        Record label, referenceClass, False, "engine error " & Err.Number & ": " & Err.Description

End Sub


'
'------------------------------------------------------------------------------
'
'                                 EXPECTATIONS
'
'------------------------------------------------------------------------------
'

Public Sub ExpectNumber( _
    ByVal label As String, _
    ByVal tradeId As String, _
    ByVal quantity As String, _
    ByVal expected As String, _
    ByVal absoluteTolerance As String, _
    ByVal relativeTolerance As String, _
    ByVal referenceClass As String)
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
'   label: catalogue ID or quantity, for the report.
'   tradeId: the trade for a trade-level quantity; empty for the netting set.
'   quantity: the quantity name.
'   expected, absoluteTolerance, relativeTolerance: numbers as text.
'   referenceClass: published, independent or illustrative.
'
' UPDATED
'   2026-10-09
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim actual      As Variant    'Output cell value
    Dim expect      As Double     'Expected value
    Dim bound       As Double     'Permitted difference
    Dim detail      As String     'Diagnostic

'------------------------------------------------------------------------------
' COMPARE
'------------------------------------------------------------------------------
        expect = Val(expected)
        bound = Val(absoluteTolerance)
        If Val(relativeTolerance) * Abs(expect) > bound Then
            bound = Val(relativeTolerance) * Abs(expect)
        End If
        If mInjectFailure And Len(mInjectedAt) = 0 Then
            expect = expect + 10# * bound + 1#
            mInjectedAt = mCase & " " & label
            Debug.Print "INJECTED=" & mInjectedAt & ": expected value moved to " & Trim$(Str$(expect))
        End If
        If Not ReadOutput(tradeId, quantity, actual, detail) Then
            Record label, referenceClass, False, detail
        ElseIf IsError(actual) Or IsEmpty(actual) Or Not IsNumeric(actual) Or VarType(actual) = vbString Then
            Record label, referenceClass, False, quantity & " is not a number"
        ElseIf CDbl(actual) < 0# And MustBeNonNegative(quantity) Then
            Record label, referenceClass, False, _
                   quantity & " must not be negative, actual " & Trim$(Str$(CDbl(actual)))
        Else
            Record label, referenceClass, Abs(CDbl(actual) - expect) <= bound, _
                   quantity & " expected " & Trim$(Str$(expect)) & ", actual " & _
                   Trim$(Str$(CDbl(actual))) & ", tolerance " & Trim$(Str$(bound))
        End If

End Sub


Public Function OutputNumber( _
    ByVal nettingSetId As String, _
    ByVal quantity As String) _
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
'   nettingSetId: the netting set's row on Results.
'   quantity: a netting-set quantity name, as for ExpectNumber.
'
' ERROR POLICY
'   Raises ERR_TEST_SETUP when the row, the column or a number is missing;
'   the calling suite then fails.
'
' UPDATED
'   2026-10-07
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ws    As Worksheet    'Results sheet
    Dim col   As Long         'Output column; 0 when unknown
    Dim r     As Long         'Output row; 0 when not found
    Dim v     As Variant      'Cell value

'------------------------------------------------------------------------------
' READ
'------------------------------------------------------------------------------
        Set ws = GetSheet(SH_RESULTS)
        col = OutputColumn(quantity, False)
        r = FindHeaderRow(ws, 1, nettingSetId)
        If col = 0 Or r < FIRST_DATA_ROW Then
            Err.Raise ERR_TEST_SETUP, "TEST_CaseRunner.OutputNumber", _
                      quantity & " of " & nettingSetId & " not found on Results."
        End If
        v = ws.Cells(r, col).Value
        If Not IsNum(v) Then
            Err.Raise ERR_TEST_SETUP, "TEST_CaseRunner.OutputNumber", _
                      quantity & " of " & nettingSetId & " is not a number (" & SafeStr(v) & ")."
        End If
        OutputNumber = CDbl(v)

End Function


Public Sub ExpectTrue( _
    ByVal label As String, _
    ByVal passed As Boolean, _
    ByVal detail As String, _
    ByVal referenceClass As String)
'
'==============================================================================
'                                  ExpectTrue
'------------------------------------------------------------------------------
' PURPOSE
'   Record a check that the caller has evaluated itself, such as a relation
'   between two outputs.
'
' INPUTS
'   label: the check's name in the report.
'   passed: the result.
'   detail: printed when the check fails.
'   referenceClass: published, independent or illustrative.
'
' UPDATED
'   2026-10-07
'==============================================================================
'
        Record label, referenceClass, passed, detail

End Sub


Public Sub ExpectText( _
    ByVal label As String, _
    ByVal tradeId As String, _
    ByVal quantity As String, _
    ByVal expected As String, _
    ByVal referenceClass As String)
'
'==============================================================================
'                                  ExpectText
'------------------------------------------------------------------------------
' PURPOSE
'   Check a text output for exact, case-sensitive equality.
'
' INPUTS
'   As ExpectNumber; expected is the exact text.
'
' UPDATED
'   2026-10-06
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim actual   As Variant    'Output cell value
    Dim detail   As String     'Diagnostic

'------------------------------------------------------------------------------
' COMPARE
'------------------------------------------------------------------------------
        If Not ReadOutput(tradeId, quantity, actual, detail) Then
            Record label, referenceClass, False, detail
        ElseIf IsError(actual) Then
            Record label, referenceClass, False, quantity & " is an error value"
        Else
            Record label, referenceClass, StrComp(CStr(actual), expected, vbBinaryCompare) = 0, _
                   quantity & " expected """ & expected & """, actual """ & CStr(actual) & """"
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
    ByVal tradeId As String, _
    ByVal quantity As String, _
    ByRef actual As Variant, _
    ByRef detail As String) _
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
'   True with actual set to the cell value; False with detail saying what
'   was missing.
'
' UPDATED
'   2026-10-06
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ws    As Worksheet    'Results or TradeCalc
    Dim rowId As String       'Row ID to find
    Dim col   As Long         'Output column; 0 when unknown
    Dim r     As Long         'Output row; 0 when not found

'------------------------------------------------------------------------------
' LOCATE
'------------------------------------------------------------------------------
        If Len(tradeId) = 0 Then
            Set ws = GetSheet(SH_RESULTS)
            rowId = mNettingSetId
        Else
            Set ws = GetSheet(SH_TRADECALC)
            rowId = tradeId
        End If
        col = OutputColumn(quantity, Len(tradeId) > 0)
        If col = 0 Then
            detail = "no output column for " & quantity
            Exit Function
        End If
        r = FindHeaderRow(ws, 1, rowId)
        If r < FIRST_DATA_ROW Then
            detail = rowId & " not found on " & ws.Name
            Exit Function
        End If
        actual = ws.Cells(r, col).Value
        ReadOutput = True

End Function


Private Function MustBeNonNegative( _
    ByVal quantity As String) _
    As Boolean
'
'==============================================================================
'                              MustBeNonNegative
'------------------------------------------------------------------------------
' PURPOSE
'   Tell whether a quantity is non-negative by definition, so that a
'   negative result fails whatever the tolerance (TEST_CASES.md).
'   Supervisory delta and adjusted notional carry a sign and are excluded.
'
' UPDATED
'   2026-10-06
'==============================================================================
'
        Select Case quantity
            Case "exposure_value", "replacement_cost", "potential_future_exposure", _
                 "aggregate_add_on", "multiplier", "margin_period_of_risk", _
                 "supervisory_factor", "lambda_shift"
                MustBeNonNegative = True
            Case Else
                MustBeNonNegative = (Left$(quantity, 7) = "add_on.")
        End Select

End Function


Private Function OutputColumn( _
    ByVal quantity As String, _
    ByVal tradeLevel As Boolean) _
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
'   2026-10-06
'==============================================================================
'

'------------------------------------------------------------------------------
' MAP
'------------------------------------------------------------------------------
        If tradeLevel Then
            Select Case quantity
                Case "trade_status"
                    OutputColumn = 4
                Case "hedging_set"
                    OutputColumn = 5
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
            Select Case quantity
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
    ByVal label As String, _
    ByVal referenceClass As String, _
    ByVal passed As Boolean, _
    ByVal detail As String)
'
'==============================================================================
'                                    Record
'------------------------------------------------------------------------------
' PURPOSE
'   Count one check under its reference class and print it when it fails.
'
' UPDATED
'   2026-10-06
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim cls   As Long    'Reference-class index

'------------------------------------------------------------------------------
' COUNT
'------------------------------------------------------------------------------
        Select Case referenceClass
            Case "published"
                cls = CLASS_PUBLISHED
            Case "independent"
                cls = CLASS_INDEPENDENT
            Case Else
                cls = CLASS_ILLUSTRATIVE
        End Select
        mCheckCount = mCheckCount + 1
        If passed Then
            mPassed(cls) = mPassed(cls) + 1
        Else
            mFailed(cls) = mFailed(cls) + 1
            mFailureCount = mFailureCount + 1
            Debug.Print "FAILURE=" & mCase & " " & label & " (" & referenceClass & "): " & detail
        End If

End Sub


Private Function ChecksHasError( _
    ByVal messagePart As String) _
    As Boolean
'
'==============================================================================
'                                ChecksHasError
'------------------------------------------------------------------------------
' PURPOSE
'   Tell whether the Checks sheet has an ERROR line whose message contains
'   messagePart, ignoring case.
'
' UPDATED
'   2026-10-07
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ws      As Worksheet    'Checks sheet
    Dim lastR   As Long         'Last used row
    Dim r       As Long         'Row being read

'------------------------------------------------------------------------------
' SEARCH
'------------------------------------------------------------------------------
    'Checks columns: 1 severity, 5 message.
        Set ws = GetSheet(SH_CHECKS)
        lastR = UsedLastRow(ws)
        For r = FIRST_DATA_ROW To lastR
            If SafeStr(ws.Cells(r, 1).Value) = SEV_ERROR Then
                If InStr(1, SafeStr(ws.Cells(r, 5).Value), messagePart, vbTextCompare) > 0 Then
                    ChecksHasError = True
                    Exit Function
                End If
            End If
        Next r

End Function


Private Sub UnitFxRate( _
    ByVal currencyCode As String)
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
'   currencyCode: three-letter code, as in column A of the FX table.
'
' UPDATED
'   2026-10-09
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ws   As Worksheet    'Params sheet
    Dim r    As Long         'Row of the currency in column A; 0 when absent

'------------------------------------------------------------------------------
' PATCH
'------------------------------------------------------------------------------
        Set ws = GetSheet(SH_PARAMS)
        r = FindHeaderRow(ws, PRM_CODE_COL, currencyCode)
        If r > 0 Then
            If ToDbl(ws.Cells(r, PRM_VALUE_COL).Value) <> 1# Then
                PatchCell SH_PARAMS, r, PRM_VALUE_COL, 1#
            End If
        End If

End Sub


Private Sub RecordPatch( _
    ByVal sheetName As String, _
    ByVal target As String, _
    ByVal saved As Variant)
'
'==============================================================================
'                                 RecordPatch
'------------------------------------------------------------------------------
' PURPOSE
'   Remember a cell's formula or a workbook name's reference before
'   PatchCell or PatchName changes it.
'
' UPDATED
'   2026-10-07
'==============================================================================
'
        mPatchCount = mPatchCount + 1
        ReDim Preserve mPatchSheet(1 To mPatchCount)
        ReDim Preserve mPatchTarget(1 To mPatchCount)
        ReDim Preserve mPatchSaved(1 To mPatchCount)
        mPatchSheet(mPatchCount) = sheetName
        mPatchTarget(mPatchCount) = target
        mPatchSaved(mPatchCount) = saved

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
'   2026-10-07
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim i         As Long    'Patch being undone
    Dim nPatches  As Long    'Patches recorded

'------------------------------------------------------------------------------
' RESTORE
'------------------------------------------------------------------------------
        nPatches = mPatchCount
        mPatchCount = 0
        For i = nPatches To 1 Step -1
            If Len(mPatchSheet(i)) = 0 Then
                ThisWorkbook.Names(mPatchTarget(i)).RefersTo = mPatchSaved(i)
            Else
                GetSheet(mPatchSheet(i)).Range(mPatchTarget(i)).Formula = mPatchSaved(i)
            End If
        Next i

End Sub


Private Function RestoreStep( _
    ByVal stepName As String, _
    ByVal stepNumber As Long) _
    As Boolean
'
'==============================================================================
'                                 RestoreStep
'------------------------------------------------------------------------------
' PURPOSE
'   Run one restoration step of EndSuite with its own error handler.
'
' INPUTS
'   stepName: name for the report.
'   stepNumber: 1 inputs and patches, 2 parameters, 3 outputs,
'   4 calculation mode, 5 events, 6 screen updating. Each setting is its
'   own step, so one failed restoration cannot leave the others changed.
'
' RETURNS
'   True when the step succeeded.
'
' UPDATED
'   2026-10-06
'==============================================================================
'

'------------------------------------------------------------------------------
' RESTORE
'------------------------------------------------------------------------------
        On Error GoTo Failed
        Select Case stepNumber
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
        Debug.Print "FAILURE=restore " & stepName & ": " & Err.Number & ": " & Err.Description

End Function


Private Sub SaveBlock( _
    ByVal ws As Worksheet, _
    ByVal nCols As Long, _
    ByRef blockAddress As String, _
    ByRef formulas As Variant)
'
'==============================================================================
'                                  SaveBlock
'------------------------------------------------------------------------------
' PURPOSE
'   Save the formulas of an input block, from the first data row to the
'   last used row.
'
' RETURNS
'   blockAddress: the block's address, "" when the sheet has no data rows.
'   formulas: the block's formulas.
'
' UPDATED
'   2026-10-06
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim lastR   As Long    'Last used row

'------------------------------------------------------------------------------
' SAVE
'------------------------------------------------------------------------------
        blockAddress = ""
        formulas = Empty
        lastR = UsedLastRow(ws)
        If lastR < FIRST_DATA_ROW Then
            Exit Sub
        End If
        With ws.Range(ws.Cells(FIRST_DATA_ROW, 1), ws.Cells(lastR, nCols))
            blockAddress = .Address
            formulas = .Formula
        End With

End Sub


Private Sub RestoreBlock( _
    ByVal ws As Worksheet, _
    ByVal nCols As Long, _
    ByVal blockAddress As String, _
    ByVal formulas As Variant)
'
'==============================================================================
'                                 RestoreBlock
'------------------------------------------------------------------------------
' PURPOSE
'   Clear the case rows and write the saved block back.
'
' UPDATED
'   2026-10-06
'==============================================================================
'

'------------------------------------------------------------------------------
' RESTORE
'------------------------------------------------------------------------------
        ClearBlock ws, nCols
        If Len(blockAddress) > 0 Then
            ws.Range(blockAddress).Formula = formulas
        End If

End Sub


Private Sub ClearBlock( _
    ByVal ws As Worksheet, _
    ByVal nCols As Long)
'
'==============================================================================
'                                  ClearBlock
'------------------------------------------------------------------------------
' PURPOSE
'   Clear the values of an input block; formats and validation stay.
'
' UPDATED
'   2026-10-06
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim lastR   As Long    'Last used row

'------------------------------------------------------------------------------
' CLEAR
'------------------------------------------------------------------------------
        lastR = UsedLastRow(ws)
        If lastR >= FIRST_DATA_ROW Then
            ws.Range(ws.Cells(FIRST_DATA_ROW, 1), ws.Cells(lastR, nCols)).ClearContents
        End If

End Sub


Private Sub WriteRow( _
    ByVal ws As Worksheet, _
    ByVal rowNum As Long, _
    ByVal nCols As Long, _
    ByRef values() As Variant)
'
'==============================================================================
'                                   WriteRow
'------------------------------------------------------------------------------
' PURPOSE
'   Write one input row from a (1 To 1, 1 To nCols) array.
'
' UPDATED
'   2026-10-06
'==============================================================================
'
        ws.Range(ws.Cells(rowNum, 1), ws.Cells(rowNum, nCols)).Value = values

End Sub


Private Function ParamCell( _
    ByVal code As String) _
    As Range
'
'==============================================================================
'                                  ParamCell
'------------------------------------------------------------------------------
' PURPOSE
'   Return the cell the engine reads a parameter from, the same way as
'   CORE_Util.GetParam: the workbook name first, then the Params table.
'
' ERROR POLICY
'   Raises when the parameter is in neither place.
'
' UPDATED
'   2026-10-06
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim r          As Range   'Range of the workbook name, if any
    Dim paramRow   As Long    'Row of the code on Params

'------------------------------------------------------------------------------
' FIND
'------------------------------------------------------------------------------
        On Error Resume Next
        Set r = ThisWorkbook.Names(code).RefersToRange
        Err.Clear
        On Error GoTo 0
        If Not r Is Nothing Then
            Set ParamCell = r.Cells(1, 1)
            Exit Function
        End If
        paramRow = FindHeaderRow(GetSheet(SH_PARAMS), PRM_CODE_COL, code)
        If paramRow = 0 Then
            Err.Raise ERR_TEST_SETUP, "TEST_CaseRunner.ParamCell", "Parameter " & code & " not found."
        End If
        Set ParamCell = GetSheet(SH_PARAMS).Cells(paramRow, PRM_VALUE_COL)

End Function


Private Function NumberOrBlank( _
    ByVal inputText As String) _
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
'   2026-10-06
'==============================================================================
'
        If Len(inputText) = 0 Then
            NumberOrBlank = Empty
        Else
            NumberOrBlank = Val(inputText)
        End If

End Function


Private Function TextOrBlank( _
    ByVal inputText As String) _
    As Variant
'
'==============================================================================
'                                 TextOrBlank
'------------------------------------------------------------------------------
' PURPOSE
'   Return text for a cell; empty means blank.
'
' UPDATED
'   2026-10-06
'==============================================================================
'
        If Len(inputText) = 0 Then
            TextOrBlank = Empty
        Else
            TextOrBlank = inputText
        End If

End Function


Private Function IsoDate( _
    ByVal inputText As String) _
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
'   2026-10-06
'==============================================================================
'
        If Len(inputText) = 0 Then
            IsoDate = Empty
        Else
            IsoDate = DateSerial(CInt(Left$(inputText, 4)), CInt(Mid$(inputText, 6, 2)), CInt(Right$(inputText, 2)))
        End If

End Function
