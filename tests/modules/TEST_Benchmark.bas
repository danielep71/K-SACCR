Attribute VB_Name = "TEST_Benchmark"
'==============================================================================
' MODULE: TEST_Benchmark
'------------------------------------------------------------------------------
' PURPOSE
'   Measure how long the engine takes on large synthetic portfolios, and
'   check that it uses every leg, for the supported portfolio size (#291):
'   by default portfolios of 1,000 and 10,000 legs, on request others up to
'   100,000, and always last a portfolio of 1,000 legs whose Trades sheet
'   has its used range inflated by a stray comment 100,000 rows below the
'   first data row, to show that the run time follows the legs and not the
'   used range. Each scenario prints a BENCH line with the seconds taken to
'   write the inputs and to run the engine.
'
' PORTFOLIO
'   The legs are spread over netting sets of 500 legs, alternately
'   unmargined and margined (daily remargining, VM, NICA and MTA, not large
'   or illiquid), all under CRR, valued on 2026-09-30 in EUR. Every 20 legs
'   repeat one mix, with references, maturities, notionals and market
'   values cycling without randomness: seven interest-rate swaps in EUR and
'   USD ending in all three maturity buckets; three FX forwards and one FX
'   option on EUR/USD, EUR/GBP, EUR/JPY and EUR/CHF; three single-name
'   credit default swaps on 50 entities rated AAA to CCC; an equity forward
'   on a single name, one on an index and an option on a single name; and
'   three commodity forwards on oil and gas and on metals. A reference
'   always has the same sub-class, no leg belongs to an economic trade, and
'   every leg is valid, so the engine must use them all.
'
' PUBLIC SURFACE
'   RunBenchmark is the entry point. Option Private Module keeps it out of
'   the external workbook automation API. It is not a test suite:
'   Test_RunAll does not run it and the Excel evidence policy does not list
'   it, because its timings depend on the machine.
'
' DEPENDENCIES
'   TEST_CaseRunner saves the inputs, starts each case, records the checks
'   and restores the workbook. The inputs are written here as one block
'   per sheet, with the columns and value types of
'   TEST_CaseRunner.AddNettingSet and AddTrade: writing 10,000 rows one by
'   one would measure the writing instead of the engine. The engine is run
'   through CORE_Engine.Calculate directly, so that its time is measured
'   alone and whether it completed is a check of its own. TEST_Suite.Say
'   prints the lines. The workbook must be built from the template: the
'   portfolio uses the factor-table categories and the FX currencies of its
'   Params sheet.
'
' WORKSHEET SAFETY
'   As TEST_CaseRunner: inputs and outputs are rewritten during the run and
'   restored at the end. The stray comment's row is deleted at the end of
'   its scenario and, after an unexpected error, before the workbook is
'   restored, so the Trades used range shrinks back to its extent before
'   the comment. Use a development workbook.
'
' STATE OWNERSHIP
'   Owns mStrayRow, the row of the stray comment while it exists.
'   TEST_CaseRunner.BeginSuite captures, and EndSuite restores, calculation
'   mode, events and screen updating; this module changes no other
'   setting.
'
' ERROR POLICY
'   Sizes that cannot be read print a BENCH_REJECTED line and change
'   nothing. An unexpected error ends the run through
'   TEST_CaseRunner.EndSuite, which restores the inputs, outputs and Excel
'   settings and reports RESULT=FAIL.
'
' USAGE
'   In the Immediate window, TEST_Benchmark.RunBenchmark runs 1,000 and
'   10,000 legs; give the sizes as text for others, for example
'   TEST_Benchmark.RunBenchmark "1000,10000,50000".
'
' UPDATED
'   2026-10-16
'
' AUTHOR
'   Daniele Penza
'==============================================================================

'------------------------------------------------------------------------------
' MODULE SETTINGS
'------------------------------------------------------------------------------
    'Require explicit declarations; keep the benchmark out of the external API.
    Option Explicit
    Option Private Module

'------------------------------------------------------------------------------
' MODULE CONSTANTS
'------------------------------------------------------------------------------
    'Scenarios.
        Private Const VALUATION       As String = "2026-09-30"    'Valuation date of every scenario
        Private Const CALC_CCY        As String = "EUR"           'Reporting currency; every amount is in it
        Private Const MAX_LEGS        As Long = 100000            'Largest scenario accepted
        Private Const MAX_SCENARIOS   As Long = 10                'Most sizes accepted in one run

    'Portfolio.
        Private Const LEGS_PER_SET    As Long = 500    'Legs per netting set
        Private Const MIX_LENGTH      As Long = 20     'Legs in one repetition of the mix
        Private Const CREDIT_NAMES    As Long = 50     'Credit entities
        Private Const EQUITY_NAMES    As Long = 40     'Single-name equities

    'Inflated used range.
        Private Const USED_RANGE_LEGS   As Long = 1000      'Legs of the used-range scenario
        Private Const STRAY_OFFSET      As Long = 100000    'Rows from the first Trades row to the stray comment
        Private Const STRAY_TEXT        As String = "Benchmark: inflated used range"    'The stray comment

    'Checks. They compare the engine's counts with the portfolio written,
    'so they validate no figure and are reported as illustrative.
        Private Const CHECKS_PER_CASE   As Long = 6     'Checks of every scenario
        Private Const STRAY_CHECKS      As Long = 2     'Further checks of the used-range scenario
        Private Const REFERENCE_CLASS   As String = "illustrative"    'Reference class of every check

'------------------------------------------------------------------------------
' MODULE STATE
'------------------------------------------------------------------------------
        Private mStrayRow   As Long    'Sheet row of the stray comment; 0 when there is none
        Private mUsedBefore As Long    'Last row of the Trades used range before the stray comment


'
'------------------------------------------------------------------------------
'
'                                 ENTRY POINT
'
'------------------------------------------------------------------------------
'

Public Sub RunBenchmark( _
    Optional ByVal sizesText As String = "1000,10000")
'
'==============================================================================
'                                 RunBenchmark
'------------------------------------------------------------------------------
' PURPOSE
'   Run one scenario per size, then the used-range scenario, and print the
'   TEST_CaseRunner report.
'
' INPUTS
'   sizesText: the legs of each scenario, whole numbers from 1 to MAX_LEGS
'   separated by commas, at most MAX_SCENARIOS of them; 1,000 and 10,000
'   when omitted.
'
' ERROR POLICY
'   Sizes that cannot be read print a BENCH_REJECTED line, and nothing is
'   changed. Any other error ends the suite as failed: the stray comment is
'   removed first, then EndSuite restores the workbook.
'
' UPDATED
'   2026-10-16
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim sizes()       As Long      'Legs of each size scenario
    Dim reason        As String    'Why sizesText was rejected; cleanup failure
    Dim scenarios     As Long      'Size scenarios and the used-range one
    Dim listed        As String    'Sizes as run, for the log
    Dim i             As Long      'Index into sizes
    Dim abortReason   As String    'Unexpected error, for EndSuite

'------------------------------------------------------------------------------
' VALIDATE
'------------------------------------------------------------------------------
        If Not ParseSizes(sizesText, sizes, reason) Then
            TEST_Suite.Say "BENCH_REJECTED=" & reason
            Exit Sub
        End If
        scenarios = UBound(sizes) + 1
        For i = 1 To UBound(sizes)
            If i > 1 Then
                listed = listed & ","
            End If
            listed = listed & sizes(i)
        Next i
        mStrayRow = 0

'------------------------------------------------------------------------------
' RUN
'------------------------------------------------------------------------------
        TEST_CaseRunner.BeginSuite scenarios, scenarios * CHECKS_PER_CASE + STRAY_CHECKS
        On Error GoTo Failed
        TEST_Suite.Say "BENCH sizes=" & listed & "; used_range_legs=" & USED_RANGE_LEGS & _
                       "; legs_per_netting_set=" & LEGS_PER_SET
        For i = 1 To UBound(sizes)
            RunScenario "bench-" & sizes(i), sizes(i), False
        Next i
        RunScenario "bench-" & USED_RANGE_LEGS & "-usedrange", USED_RANGE_LEGS, True
        TEST_CaseRunner.EndSuite ""
        Exit Sub

'------------------------------------------------------------------------------
' HANDLE ERROR
'------------------------------------------------------------------------------
    'The error is copied before the stray comment is removed, which resets
    'Err; EndSuite then restores the workbook and reports the failure.
Failed:
        abortReason = "unexpected error " & Err.Number & ": " & Err.Description
        reason = ClearStray()
        If Len(reason) > 0 Then
            abortReason = abortReason & "; " & reason
        End If
        TEST_CaseRunner.EndSuite abortReason

End Sub


'
'------------------------------------------------------------------------------
'
'                                  SCENARIOS
'
'------------------------------------------------------------------------------
'

Private Sub RunScenario( _
    ByVal caseId As String, _
    ByVal legs As Long, _
    ByVal inflateUsedRange As Boolean)
'
'==============================================================================
'                                 RunScenario
'------------------------------------------------------------------------------
' PURPOSE
'   Write a portfolio, time the writing and the engine run, record the
'   checks and print the BENCH line. In the used-range scenario a stray
'   comment is written STRAY_OFFSET rows below the first Trades row before
'   the run, and its row deleted after it.
'
' INPUTS
'   caseId: the case label.
'   legs: trade legs, 1 to MAX_LEGS.
'   inflateUsedRange: True for the used-range scenario.
'
' ERROR POLICY
'   An unexpected error propagates to RunBenchmark, which removes the
'   stray comment.
'
' UPDATED
'   2026-10-16
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim nSets          As Long       'Netting sets
    Dim t0             As Double     'Timer at the start of the writing
    Dim writeSeconds   As Double     'Seconds to write the inputs
    Dim runSeconds     As Double     'Seconds of the engine run
    Dim completed      As Boolean    'CORE_Engine.Calculate returned True
    Dim detail         As String     'Why the run did not complete; cleanup failure
    Dim usedLast       As Long       'Last row of the Trades used range before the run
    Dim legsRead       As Long       'Legs the engine read
    Dim legsUsed       As Long       'Legs the engine used
    Dim validSets      As Long       'VALID CRR netting sets
    Dim withheld       As Long       'Netting sets whose EAD was withheld
    Dim errorLines     As Long       'ERROR lines on Checks
    Dim resultsRows    As Long       'Rows on Results
    Dim tradeRows      As Long       'Rows on TradeCalc
    Dim lastId         As String     'ID of the last Results row
    Dim lastTradeId    As String     'ID of the last TradeCalc row
    Dim report         As String     'BENCH line

'------------------------------------------------------------------------------
' WRITE THE INPUTS
'------------------------------------------------------------------------------
    'One netting set per LEGS_PER_SET legs, the last one possibly smaller.
        nSets = (legs + LEGS_PER_SET - 1) \ LEGS_PER_SET
        TEST_CaseRunner.BeginCase caseId, RG_CRR, VALUATION, CALC_CCY
        t0 = Timer
        WriteNettingSets nSets
        WriteTrades legs
        writeSeconds = SecondsSince(t0)
        If inflateUsedRange Then
            mUsedBefore = UsedLastRow(GetSheet(SH_TRADES))
            mStrayRow = FirstInputRow(SH_TRADES) + STRAY_OFFSET
            TableCell(GetSheet(SH_TRADES), mStrayRow, TR_COMMENT).Value = STRAY_TEXT
            usedLast = UsedLastRow(GetSheet(SH_TRADES))
            TEST_CaseRunner.ExpectTrue "used-range", usedLast >= mStrayRow, "the Trades used range ends on row " & _
                                       usedLast & ", above the stray comment on row " & mStrayRow, REFERENCE_CLASS
        End If

'------------------------------------------------------------------------------
' RUN THE ENGINE
'------------------------------------------------------------------------------
        completed = RunEngine(runSeconds, detail)

'------------------------------------------------------------------------------
' CHECK
'------------------------------------------------------------------------------
    'Every leg read and used, every netting set VALID, no ERROR on Checks,
    'one Results row per netting set and the TOTAL CRR row, one TradeCalc
    'row per leg.
        legsRead = CORE_Engine.TradesRead
        legsUsed = CORE_Engine.TradesUsed
        validSets = CORE_Engine.RegimeValidCount(RG_CRR)
        withheld = CORE_Engine.IncompleteCount
        errorLines = CORE_Engine.ErrorCount
        resultsRows = OutputRows(SH_RESULTS, RS_ID, lastId)
        tradeRows = OutputRows(SH_TRADECALC, TC_ID, lastTradeId)
        TEST_CaseRunner.ExpectTrue "completed", completed, detail, REFERENCE_CLASS
        TEST_CaseRunner.ExpectTrue "legs-used", (legsRead = legs) And (legsUsed = legs), _
                                   "read " & legsRead & " and used " & legsUsed & " of " & legs & " legs", _
                                   REFERENCE_CLASS
        TEST_CaseRunner.ExpectTrue "netting-sets-valid", (validSets = nSets) And (withheld = 0), _
                                   validSets & " of " & nSets & " netting sets VALID, " & withheld & " withheld", _
                                   REFERENCE_CLASS
        TEST_CaseRunner.ExpectTrue "no-errors", errorLines = 0, errorLines & " ERROR line(s) on Checks", _
                                   REFERENCE_CLASS
        TEST_CaseRunner.ExpectTrue "results-rows", (resultsRows = nSets + 1) And (lastId = "TOTAL " & RG_CRR), _
                                   resultsRows & " Results rows ending with '" & lastId & "', expected " & _
                                   (nSets + 1) & " ending with 'TOTAL " & RG_CRR & "'", REFERENCE_CLASS
        TEST_CaseRunner.ExpectTrue "tradecalc-rows", tradeRows = legs, _
                                   tradeRows & " TradeCalc rows, expected " & legs, REFERENCE_CLASS

'------------------------------------------------------------------------------
' REPORT
'------------------------------------------------------------------------------
        report = "BENCH case=" & caseId & "; legs=" & legs & "; netting_sets=" & nSets & _
                 "; used=" & legsUsed & "; withheld=" & withheld & "; errors=" & errorLines & _
                 "; warnings=" & CORE_Engine.WarningCount
        report = report & "; tradecalc_rows=" & tradeRows & "; results_rows=" & resultsRows
        If inflateUsedRange Then
            report = report & "; used_last_row=" & usedLast
        End If
        report = report & "; write_s=" & Format$(writeSeconds, "0.00") & "; run_s=" & Format$(runSeconds, "0.00")
        TEST_Suite.Say report

'------------------------------------------------------------------------------
' REMOVE THE STRAY COMMENT
'------------------------------------------------------------------------------
        If inflateUsedRange Then
            detail = ClearStray()
            TEST_CaseRunner.ExpectTrue "used-range-restored", Len(detail) = 0, detail, REFERENCE_CLASS
        End If

End Sub


Private Function RunEngine( _
    ByRef runSeconds As Double, _
    ByRef detail As String) _
    As Boolean
'
'==============================================================================
'                                  RunEngine
'------------------------------------------------------------------------------
' PURPOSE
'   Run the engine once on the inputs written, writing every output sheet,
'   and time the run.
'
' RETURNS
'   True when the run completed. runSeconds: the duration in seconds, also
'   when the run stopped or raised. detail: why the run did not complete;
'   empty when it did.
'
' ERROR POLICY
'   An error raised by the engine is contained and returned in detail; the
'   engine has already withdrawn its outputs.
'
' UPDATED
'   2026-10-16
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim t0   As Double    'Timer at the start of the run

'------------------------------------------------------------------------------
' RUN
'------------------------------------------------------------------------------
        detail = ""
        t0 = Timer
        On Error GoTo Failed
        RunEngine = CORE_Engine.Calculate(True)
        runSeconds = SecondsSince(t0)
        If Not RunEngine Then
            detail = "engine run stopped; see the Checks sheet"
        End If
        Exit Function

'------------------------------------------------------------------------------
' HANDLE ERROR
'------------------------------------------------------------------------------
Failed:
        runSeconds = SecondsSince(t0)
        detail = "engine error " & Err.Number & ": " & Err.Description

End Function


Private Function ClearStray() As String
'
'==============================================================================
'                                  ClearStray
'------------------------------------------------------------------------------
' PURPOSE
'   Remove the stray comment of the used-range scenario, if one was
'   written, and give the Trades used range back its extent before the
'   comment. Clearing the cell is not enough: Excel keeps the used range
'   at the stray row, so later scans and EndSuite would still cover it.
'   The stray row is deleted instead (it holds nothing else, far below
'   the template's formatted rows) and UsedRange is read again, which
'   makes Excel recompute it.
'
' RETURNS
'   Empty on success or when there is nothing to remove; otherwise why the
'   comment could not be removed or the used range did not shrink back.
'   EndSuite still clears a remaining comment with the other Trades rows.
'
' ERROR POLICY
'   Contains its own error, so that it can run on the error path.
'
' UPDATED
'   2026-10-16
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ws        As Worksheet    'Trades
    Dim usedNow   As Long         'Last row of the Trades used range after the deletion

'------------------------------------------------------------------------------
' REMOVE
'------------------------------------------------------------------------------
        If mStrayRow = 0 Then
            Exit Function
        End If
        On Error GoTo Failed
        Set ws = GetSheet(SH_TRADES)
        TableCell(ws, mStrayRow, TR_COMMENT).EntireRow.Delete
        mStrayRow = 0
        usedNow = UsedLastRow(ws)
        If usedNow > mUsedBefore Then
            ClearStray = "the Trades used range still ends on row " & usedNow & _
                         " after the stray row was deleted; it ended on row " & mUsedBefore & " before"
        End If
        Exit Function

'------------------------------------------------------------------------------
' HANDLE ERROR
'------------------------------------------------------------------------------
Failed:
        ClearStray = "stray comment on Trades row " & mStrayRow & " not removed: error " & _
                     Err.Number & ": " & Err.Description

End Function


'
'------------------------------------------------------------------------------
'
'                                  PORTFOLIO
'
'------------------------------------------------------------------------------
'

Private Sub WriteNettingSets( _
    ByVal nSets As Long)
'
'==============================================================================
'                               WriteNettingSets
'------------------------------------------------------------------------------
' PURPOSE
'   Write the netting sets as one block from the first NettingSets row,
'   with the value types TEST_CaseRunner.AddNettingSet writes. Odd ones are
'   unmargined, without collateral; even ones are margined with daily
'   remargining, VM held, NICA and an MTA, and none is large or illiquid,
'   in dispute or client cleared.
'
' INPUTS
'   nSets: netting sets, at least 1.
'
' UPDATED
'   2026-10-16
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim block()   As Variant    'NettingSets rows, (netting set, column)
    Dim s         As Long       'Netting set, from 1
    Dim topRow    As Long       'First NettingSets data row

'------------------------------------------------------------------------------
' BUILD
'------------------------------------------------------------------------------
    'Unused cells stay Empty, as the blanks AddNettingSet writes.
        ReDim block(1 To nSets, 1 To NS_NCOLS)
        For s = 1 To nSets
            block(s, NS_ID) = NettingSetId(s)
            block(s, NS_CPTY) = "Benchmark counterparty " & s
            block(s, NS_CLEARED) = CL_NONE
            block(s, NS_LARGE) = "N"
            block(s, NS_DISPUTE) = "N"
            block(s, NS_REGIME) = RG_CRR
            If (s Mod 2) = 0 Then
                block(s, NS_MARGINED) = "Y"
                block(s, NS_FREQ) = 1#
                block(s, NS_VM) = 100000# * (s Mod 3)
                block(s, NS_NICA) = 50000#
                block(s, NS_TH) = 0#
                block(s, NS_MTA) = 10000#
            Else
                block(s, NS_MARGINED) = "N"
                block(s, NS_VM) = 0#
                block(s, NS_NICA) = 0#
                block(s, NS_TH) = 0#
                block(s, NS_MTA) = 0#
            End If
        Next s

'------------------------------------------------------------------------------
' WRITE
'------------------------------------------------------------------------------
        topRow = FirstInputRow(SH_NS)
        TableRange(GetSheet(SH_NS), topRow, 1, topRow + nSets - 1, NS_NCOLS).Value = block

End Sub


Private Sub WriteTrades( _
    ByVal legs As Long)
'
'==============================================================================
'                                 WriteTrades
'------------------------------------------------------------------------------
' PURPOSE
'   Write the legs as one block from the first Trades row, with the value
'   types TEST_CaseRunner.AddTrade writes.
'
' INPUTS
'   legs: trade legs, 1 to MAX_LEGS.
'
' UPDATED
'   2026-10-16
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim block()   As Variant    'Trades rows, (leg, column)
    Dim i         As Long       'Leg, from 1
    Dim topRow    As Long       'First Trades data row

'------------------------------------------------------------------------------
' BUILD AND WRITE
'------------------------------------------------------------------------------
        ReDim block(1 To legs, 1 To TR_NCOLS)
        For i = 1 To legs
            FillLeg block, i
        Next i
        topRow = FirstInputRow(SH_TRADES)
        TableRange(GetSheet(SH_TRADES), topRow, 1, topRow + legs - 1, TR_NCOLS).Value = block

End Sub


Private Sub FillLeg( _
    ByRef block() As Variant, _
    ByVal i As Long)
'
'==============================================================================
'                                   FillLeg
'------------------------------------------------------------------------------
' PURPOSE
'   Fill row i of the Trades block with leg i of the mix described in the
'   module header. Amounts are in EUR: notionals of 1 to 9 million, market
'   values from -15,000 to 15,000. The economic-trade, lambda and CDO
'   columns and the comment stay Empty.
'
' INPUTS
'   block: the Trades rows, (leg, column), from WriteTrades.
'   i: the leg, from 1; netting set (i - 1) \ LEGS_PER_SET + 1.
'
' UPDATED
'   2026-10-16
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim j        As Long    'Leg, from 0
    Dim slot     As Long    'Position in the mix, 0 to MIX_LENGTH - 1
    Dim cycle    As Long    'Repetitions of the mix before this leg
    Dim entity   As Long    'Credit or equity entity, from 0

'------------------------------------------------------------------------------
' COMMON FIELDS
'------------------------------------------------------------------------------
        j = i - 1
        slot = j Mod MIX_LENGTH
        cycle = j \ MIX_LENGTH
        block(i, TR_ID) = "B" & Format$(i, "000000")
        block(i, TR_NS) = NettingSetId(j \ LEGS_PER_SET + 1)
        block(i, TR_INSTR) = "Linear"
        block(i, TR_DIR) = IIf(((j \ 3) Mod 2) = 0, "Long", "Short")
        block(i, TR_NATURE) = "Standard"
        block(i, TR_NOTIONAL) = 1000000# * (1 + (j Mod 9))
        block(i, TR_NCCY) = CALC_CCY
        block(i, TR_MTM) = 1000# * ((j Mod 7) - 3) * (1 + (j Mod 5))
        block(i, TR_MCCY) = CALC_CCY

'------------------------------------------------------------------------------
' ASSET-CLASS FIELDS
'------------------------------------------------------------------------------
    'Interest-rate swaps end from 6 months to 20 years, never exactly on the
    'bucket edges of 1 and 5 years. FX pairs all hold EUR, the notional
    'currency, and are in market quote order, as an FX option requires. A
    'credit, equity or commodity reference is always given the same
    'sub-class, as the engine requires within a netting set.
        Select Case slot
            Case 0 To 6
                block(i, TR_AC) = "IR"
                block(i, TR_RF) = IIf(((cycle + slot) Mod 2) = 0, "EUR", "USD")
                SetMaturity block, i, Choose(((cycle * 7 + slot) Mod 9) + 1, 6, 9, 18, 30, 42, 84, 120, 180, 240), _
                            True
            Case 7 To 9
                block(i, TR_AC) = "FX"
                block(i, TR_RF) = FxPair(cycle + slot)
                SetMaturity block, i, Choose(((cycle + slot) Mod 5) + 1, 3, 6, 12, 18, 24), False
            Case 10
                block(i, TR_AC) = "FX"
                block(i, TR_RF) = FxPair(cycle)
                SetOption block, i, cycle, Choose((cycle Mod 4) + 1, 1.16, 0.74, 172#, 0.93)
            Case 11 To 13
                entity = (cycle * 3 + slot - 11) Mod CREDIT_NAMES
                block(i, TR_AC) = "CR"
                block(i, TR_SUB) = Choose((entity Mod 7) + 1, "AAA", "AA", "A", "BBB", "BB", "B", "CCC")
                block(i, TR_RF) = "CR NAME " & Format$(entity + 1, "00")
                SetMaturity block, i, Choose(((cycle + slot) Mod 5) + 1, 12, 36, 60, 84, 120), True
            Case 14
                block(i, TR_AC) = "EQ"
                block(i, TR_SUB) = "SINGLE"
                block(i, TR_RF) = "EQ NAME " & Format$((cycle * 2) Mod EQUITY_NAMES + 1, "00")
                SetMaturity block, i, Choose((cycle Mod 4) + 1, 3, 6, 12, 24), False
            Case 15
                block(i, TR_AC) = "EQ"
                block(i, TR_SUB) = "INDEX"
                block(i, TR_RF) = "EQ INDEX " & (cycle Mod 3 + 1)
                SetMaturity block, i, Choose(((cycle + 1) Mod 4) + 1, 3, 6, 12, 24), False
            Case 16
                block(i, TR_AC) = "EQ"
                block(i, TR_SUB) = "SINGLE"
                block(i, TR_RF) = "EQ NAME " & Format$((cycle * 2 + 1) Mod EQUITY_NAMES + 1, "00")
                SetOption block, i, cycle, 100#
            Case 17, 19
                block(i, TR_AC) = "CO"
                block(i, TR_SUB) = "OIL_GAS"
                block(i, TR_RF) = Choose(((cycle + slot) Mod 3) + 1, "BRENT", "WTI", "NATURAL GAS")
                SetMaturity block, i, Choose(((cycle + slot) Mod 4) + 1, 3, 6, 12, 24), False
            Case Else
                block(i, TR_AC) = "CO"
                block(i, TR_SUB) = "METALS"
                block(i, TR_RF) = Choose((cycle Mod 3) + 1, "GOLD", "SILVER", "COPPER")
                SetMaturity block, i, Choose(((cycle + 2) Mod 4) + 1, 3, 6, 12, 24), False
        End Select

End Sub


Private Sub SetMaturity( _
    ByRef block() As Variant, _
    ByVal i As Long, _
    ByVal months As Long, _
    ByVal withEndDate As Boolean)
'
'==============================================================================
'                                 SetMaturity
'------------------------------------------------------------------------------
' PURPOSE
'   Give leg i a maturity date months after the valuation date, and the
'   same end date when withEndDate is True (swaps and credit default
'   swaps). The start date stays Empty: the leg starts on the valuation
'   date.
'
' INPUTS
'   block, i: as FillLeg.
'   months: months to maturity, 1 or more.
'   withEndDate: True to write the end date too.
'
' UPDATED
'   2026-10-16
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim maturity   As Date    'Maturity date

'------------------------------------------------------------------------------
' WRITE
'------------------------------------------------------------------------------
        maturity = MonthsAfterValuation(months)
        block(i, TR_MAT) = maturity
        If withEndDate Then
            block(i, TR_END) = maturity
        End If

End Sub


Private Sub SetOption( _
    ByRef block() As Variant, _
    ByVal i As Long, _
    ByVal cycle As Long, _
    ByVal price As Double)
'
'==============================================================================
'                                  SetOption
'------------------------------------------------------------------------------
' PURPOSE
'   Make leg i a European option on its underlying: a call or a put, with
'   the underlying price, a strike 5% below, at or 5% above it, and an
'   expiry 3, 6 or 12 months after the valuation date, which is also its
'   maturity. The lambda shift stays Empty: it applies to interest-rate and
'   commodity options only.
'
' INPUTS
'   block, i: as FillLeg.
'   cycle: the leg's repetition of the mix, which sets the type, strike and
'   expiry.
'   price: the underlying price P, > 0.
'
' UPDATED
'   2026-10-16
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim expiry   As Date    'Expiry and maturity date

'------------------------------------------------------------------------------
' WRITE
'------------------------------------------------------------------------------
        expiry = MonthsAfterValuation(Choose(((cycle + 1) Mod 3) + 1, 3, 6, 12))
        block(i, TR_INSTR) = "Option"
        block(i, TR_OPT) = IIf((cycle Mod 2) = 0, "Call", "Put")
        block(i, TR_PRICE) = price
        block(i, TR_STRIKE) = price * Choose((cycle Mod 3) + 1, 0.95, 1#, 1.05)
        block(i, TR_EXPIRY) = expiry
        block(i, TR_MAT) = expiry

End Sub


'
'------------------------------------------------------------------------------
'
'                                   HELPERS
'
'------------------------------------------------------------------------------
'

Private Function ParseSizes( _
    ByVal sizesText As String, _
    ByRef sizes() As Long, _
    ByRef reason As String) _
    As Boolean
'
'==============================================================================
'                                  ParseSizes
'------------------------------------------------------------------------------
' PURPOSE
'   Read the sizes given to RunBenchmark.
'
' INPUTS
'   sizesText: whole numbers separated by commas; spaces around them are
'   ignored.
'
' RETURNS
'   True with sizes (1 To count) filled when there are 1 to MAX_SCENARIOS
'   sizes, each a whole number from 1 to MAX_LEGS; otherwise False with
'   reason saying what is wrong.
'
' UPDATED
'   2026-10-16
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim parts()   As String    'Items between the commas
    Dim nSizes    As Long      'Items
    Dim i         As Long      'Size, from 1
    Dim item      As String    'One item, trimmed
    Dim rule      As String    'What a valid argument is

'------------------------------------------------------------------------------
' READ
'------------------------------------------------------------------------------
        rule = "give 1 to " & MAX_SCENARIOS & " whole numbers of legs from 1 to " & MAX_LEGS & _
               ", separated by commas, such as ""1000,10000,50000"""
        parts = Split(sizesText, ",")
        nSizes = UBound(parts) - LBound(parts) + 1
        If nSizes < 1 Or nSizes > MAX_SCENARIOS Then
            reason = rule & " (found """ & sizesText & """)"
            Exit Function
        End If
        ReDim sizes(1 To nSizes)
        For i = 1 To nSizes
            item = Trim$(parts(LBound(parts) + i - 1))
            If Len(item) = 0 Or Len(item) > 6 Or item Like "*[!0-9]*" Then
                reason = rule & " (found """ & item & """)"
                Exit Function
            End If
            sizes(i) = CLng(item)
            If sizes(i) < 1 Or sizes(i) > MAX_LEGS Then
                reason = rule & " (found " & sizes(i) & ")"
                Exit Function
            End If
        Next i
        ParseSizes = True

End Function


Private Function FirstInputRow( _
    ByVal sheetName As String) _
    As Long
'
'==============================================================================
'                                FirstInputRow
'------------------------------------------------------------------------------
' PURPOSE
'   Return the first data row of an input sheet, located as
'   TEST_CaseRunner locates it, through the workbook name of its first
'   header cell (CORE_Util.TableFirstRow).
'
' ERROR POLICY
'   Raises ERR_TEST_SETUP when the name cannot locate it; the run then ends
'   as failed.
'
' UPDATED
'   2026-10-16
'==============================================================================
'
        FirstInputRow = TableFirstRow(sheetName)
        If FirstInputRow = 0 Then
            Err.Raise ERR_TEST_SETUP, "TEST_Benchmark.FirstInputRow", _
                      "The header of " & sheetName & " cannot be located through its workbook name - " & _
                      "rebuild the workbook from the template."
        End If

End Function


Private Function OutputRows( _
    ByVal sheetName As String, _
    ByVal idCol As Long, _
    ByRef lastId As String) _
    As Long
'
'==============================================================================
'                                  OutputRows
'------------------------------------------------------------------------------
' PURPOSE
'   Count the rows of an output table, from its first data row to its last
'   row with an ID.
'
' INPUTS
'   sheetName: SH_RESULTS or SH_TRADECALC.
'   idCol: the table column of the row ID.
'
' RETURNS
'   The number of rows, 0 when there are none. lastId: the ID of the last
'   row; empty when there are none.
'
' UPDATED
'   2026-10-16
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ws         As Worksheet    'Output sheet
    Dim firstRow   As Long         'First data row
    Dim lastRow    As Long         'Last row with an ID

'------------------------------------------------------------------------------
' COUNT
'------------------------------------------------------------------------------
        Set ws = GetSheet(sheetName)
        firstRow = OutputFirstRow(sheetName)
        lastRow = LastDataRow(ws, firstRow, idCol)
        OutputRows = lastRow - firstRow + 1
        lastId = ""
        If lastRow >= firstRow Then
            lastId = SafeStr(TableCell(ws, lastRow, idCol).Value)
        End If

End Function


Private Function NettingSetId( _
    ByVal s As Long) _
    As String
'
'==============================================================================
'                                 NettingSetId
'------------------------------------------------------------------------------
' PURPOSE
'   ID of netting set s, from 1: BENCH-001, BENCH-002 and so on.
'
' UPDATED
'   2026-10-16
'==============================================================================
'
        NettingSetId = "BENCH-" & Format$(s, "000")

End Function


Private Function FxPair( _
    ByVal index As Long) _
    As String
'
'==============================================================================
'                                    FxPair
'------------------------------------------------------------------------------
' PURPOSE
'   One of the four currency pairs of the portfolio, in market quote order:
'   EUR/USD, EUR/GBP, EUR/JPY or EUR/CHF for index 0, 1, 2 or 3 modulo 4.
'   Each currency is in the template's FX table.
'
' UPDATED
'   2026-10-16
'==============================================================================
'
        FxPair = Choose((index Mod 4) + 1, "EUR/USD", "EUR/GBP", "EUR/JPY", "EUR/CHF")

End Function


Private Function MonthsAfterValuation( _
    ByVal months As Long) _
    As Date
'
'==============================================================================
'                             MonthsAfterValuation
'------------------------------------------------------------------------------
' PURPOSE
'   The date months calendar months after VALUATION, on the same day or
'   the last day of a shorter month.
'
' UPDATED
'   2026-10-16
'==============================================================================
'
        MonthsAfterValuation = DateAdd("m", months, DateSerial(CInt(Left$(VALUATION, 4)), _
                                                               CInt(Mid$(VALUATION, 6, 2)), _
                                                               CInt(Right$(VALUATION, 2))))

End Function


Private Function SecondsSince( _
    ByVal startTime As Double) _
    As Double
'
'==============================================================================
'                                 SecondsSince
'------------------------------------------------------------------------------
' PURPOSE
'   Seconds elapsed since startTime, a value of Timer, also across
'   midnight, when Timer starts again from 0.
'
' UPDATED
'   2026-10-16
'==============================================================================
'
        SecondsSince = Timer - startTime
        If SecondsSince < 0# Then
            SecondsSince = SecondsSince + 86400#
        End If

End Function
