Attribute VB_Name = "TEST_Benchmark"
'==============================================================================
' MODULE: TEST_Benchmark
'------------------------------------------------------------------------------
' PURPOSE
'   Measure how long the engine takes on large synthetic portfolios, and
'   check that it uses every leg, for the supported portfolio size (#291):
'   by default portfolios of 1,000 and 10,000 legs, on request others up to
'   1,000,000, and always last a portfolio of 1,000 legs whose Trades sheet
'   has its used range inflated by a stray comment 100,000 rows below the
'   end of its used range, to measure what an inflated used range adds to
'   the run. Each scenario prints a BENCH line with the seconds taken to
'   write the inputs and to run the engine.
'
' PORTFOLIO
'   The legs are spread over netting sets of 500 legs, alternately
'   unmargined and margined (daily remargining, VM, NICA and MTA, not large
'   or illiquid), all under CRR, valued on 2026-09-30 in EUR. Every 21 legs
'   repeat one mix, with references, maturities, notionals and market
'   values cycling without randomness: seven interest-rate swaps in EUR and
'   USD ending in all three maturity buckets; three FX forwards and one FX
'   option on EUR/USD, EUR/GBP, EUR/JPY and EUR/CHF; three single-name
'   credit default swaps on 50 entities rated AAA to CCC; an equity forward
'   on a single name, one on an index and an option on a single name;
'   three commodity forwards on oil and gas and on metals; and one
'   other-risks (OT) forward, a class only CRR has, on one of three
'   indices, long or short. A reference always has the same sub-class, no
'   leg belongs to an economic trade, and every leg is valid, so the
'   engine must use them all.
'
' PUBLIC SURFACE
'   RunBenchmark is the entry point for the engine; RunCsvBenchmark times
'   the CSV reader on synthetic files of up to the contract's record limit
'   and beyond (#207). Option Private Module keeps both out of
'   the external workbook automation API. Neither is a test suite:
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
'   Params sheet. RunCsvBenchmark uses CORE_CsvReader, CORE_Export for the
'   UTF-8 bytes and the CSV header lines of CORE_Config, and no sheet.
'
' WORKSHEET SAFETY
'   As TEST_CaseRunner: inputs and outputs are rewritten during the run and
'   restored at the end. The stray comment's row is deleted at the end of
'   its scenario and, after an unexpected error, before the workbook is
'   restored, so the Trades used range shrinks back to its extent before
'   the comment. After every scenario, and before the workbook is restored
'   after an unexpected error, the rows written below each sheet's extent
'   before the run are deleted, so a large portfolio leaves no used range
'   behind. Use a development workbook.
'
' STATE OWNERSHIP
'   Owns mStrayRow and mUsedBefore, the stray comment's row and the end of
'   the Trades used range before it, and mBaseRows and mHaveBase, each
'   sheet's last used row before the run. TEST_CaseRunner.BeginSuite
'   captures, and EndSuite restores, calculation mode, events and screen
'   updating; this module changes no other setting.
'
' ERROR POLICY
'   Sizes that cannot be read print a BENCH_REJECTED line and change
'   nothing. An unexpected error removes the stray comment and the rows
'   added below each sheet's extent, then ends the run through
'   TEST_CaseRunner.EndSuite, which restores the inputs, outputs and Excel
'   settings and reports RESULT=FAIL.
'
' USAGE
'   In the Immediate window, TEST_Benchmark.RunBenchmark runs 1,000 and
'   10,000 legs; give the sizes as text for others, for example
'   TEST_Benchmark.RunBenchmark "1000,10000,50000".
'   TEST_Benchmark.RunCsvBenchmark reads Trades files of 10,000 and
'   100,000 records and a NettingSets file of 10,000, writing each to the
'   TEMP folder under a name not already taken and deleting it; give
'   other sizes the same way.
'
' UPDATED
'   2026-10-23
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
        Private Const MAX_LEGS        As Long = 1000000           'Largest scenario; within Excel's 1,048,576 rows
        Private Const MAX_SCENARIOS   As Long = 10                'Most sizes accepted in one run

    'Portfolio.
        Private Const LEGS_PER_SET    As Long = 500    'Legs per netting set
        Private Const MIX_LENGTH      As Long = 21     'Legs in one repetition of the mix
        Private Const CREDIT_NAMES    As Long = 50     'Credit entities
        Private Const EQUITY_NAMES    As Long = 40     'Single-name equities
        Private Const OTHER_INDICES   As Long = 3      'Other-risks indices

    'Inflated used range.
        Private Const USED_RANGE_LEGS   As Long = 1000      'Legs of the used-range scenario
        Private Const STRAY_OFFSET      As Long = 100000    'Rows from the end of the Trades used range to the comment
        Private Const STRAY_TEXT        As String = "Benchmark: inflated used range"    'The stray comment

    'Checks. They compare the engine's counts with the portfolio written,
    'so they validate no figure and are reported as illustrative.
        Private Const CHECKS_PER_CASE   As Long = 7     'Checks of every scenario
        Private Const STRAY_CHECKS      As Long = 2     'Further checks of the used-range scenario
        Private Const REFERENCE_CLASS   As String = "illustrative"    'Reference class of every check

    'CSV reader (RunCsvBenchmark). The limits are the contract's, in
    'docs/csv/contract.json.
        Private Const CSV_BENCH_FILE                As String = "K-SACCR_csv_benchmark"    'File in TEMP, before its number
        Private Const CSV_BENCH_NETTING_SETS        As Long = 10000     'Records of the NettingSets file
        Private Const CSV_BENCH_SETS                As Long = 200       'Netting sets the trades refer to
        Private Const CSV_BENCH_TRADES_LIMIT        As Long = 100000    'Trades records per file
        Private Const CSV_BENCH_NETTING_SETS_LIMIT  As Long = 10000     'NettingSets records per file

'------------------------------------------------------------------------------
' MODULE STATE
'------------------------------------------------------------------------------
        Private mStrayRow   As Long    'Sheet row of the stray comment; 0 when there is none
        Private mUsedBefore As Long    'Last row of the Trades used range before the stray comment
        Private mBaseRows() As Long    'Last used row of each sheet of ExtentSheets before the run
        Private mHaveBase   As Boolean    'mBaseRows has been recorded


'
'------------------------------------------------------------------------------
'
'                                 ENTRY POINT
'
'------------------------------------------------------------------------------
'

Public Sub RunBenchmark( _
    Optional ByVal SizesText As String = "1000,10000")
'
'==============================================================================
'                                 RunBenchmark
'------------------------------------------------------------------------------
' PURPOSE
'   Run one scenario per size, then the used-range scenario, and print the
'   TEST_CaseRunner report.
'
' INPUTS
'   SizesText: the legs of each scenario, whole numbers from 1 to MAX_LEGS
'   separated by commas, at most MAX_SCENARIOS of them; 1,000 and 10,000
'   when omitted.
'
' ERROR POLICY
'   Sizes that cannot be read print a BENCH_REJECTED line, and nothing is
'   changed. Any other error ends the suite as failed: the stray comment
'   and the rows added below each sheet's extent are removed first, then
'   EndSuite restores the workbook.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Sizes()         As Long      'Legs of each size scenario
    Dim Reason          As String    'Why SizesText was rejected; cleanup failure
    Dim ScenarioCount   As Long      'Size scenarios and the used-range one
    Dim SizesListed     As String    'Sizes as run, for the log
    Dim i               As Long      'Index into Sizes
    Dim AbortReason     As String    'Unexpected error, for EndSuite

'------------------------------------------------------------------------------
' VALIDATE
'------------------------------------------------------------------------------
        If Not ParseSizes(SizesText, Sizes, Reason) Then
            TEST_Suite.Say "BENCH_REJECTED=" & Reason
            Exit Sub
        End If
        ScenarioCount = UBound(Sizes) + 1
        For i = 1 To UBound(Sizes)
            If i > 1 Then
                SizesListed = SizesListed & ","
            End If
            SizesListed = SizesListed & Sizes(i)
        Next i
        mStrayRow = 0
        RecordExtents

'------------------------------------------------------------------------------
' RUN
'------------------------------------------------------------------------------
        TEST_CaseRunner.BeginSuite ScenarioCount, ScenarioCount * CHECKS_PER_CASE + STRAY_CHECKS
        On Error GoTo Failed
        TEST_Suite.Say "BENCH sizes=" & SizesListed & "; used_range_legs=" & USED_RANGE_LEGS & _
                       "; legs_per_netting_set=" & LEGS_PER_SET
        For i = 1 To UBound(Sizes)
            RunScenario "bench-" & Sizes(i), Sizes(i), False
        Next i
        RunScenario "bench-" & USED_RANGE_LEGS & "-usedrange", USED_RANGE_LEGS, True
        TEST_CaseRunner.EndSuite ""
        Exit Sub

'------------------------------------------------------------------------------
' HANDLE ERROR
'------------------------------------------------------------------------------
    'The error is copied before the stray comment and the added rows are
    'removed, which resets Err; EndSuite then restores the workbook and
    'reports the failure.
Failed:
        AbortReason = "unexpected error " & Err.Number & ": " & Err.Description
        Reason = ClearStray()
        If Len(Reason) > 0 Then
            AbortReason = AbortReason & "; " & Reason
        End If
        Reason = ShrinkSheets()
        If Len(Reason) > 0 Then
            AbortReason = AbortReason & "; " & Reason
        End If
        TEST_CaseRunner.EndSuite AbortReason

End Sub


'
'------------------------------------------------------------------------------
'
'                                  SCENARIOS
'
'------------------------------------------------------------------------------
'

Private Sub RunScenario( _
    ByVal CaseId As String, _
    ByVal LegCount As Long, _
    ByVal InflateUsedRange As Boolean)
'
'==============================================================================
'                                 RunScenario
'------------------------------------------------------------------------------
' PURPOSE
'   Write a portfolio, time the writing and the engine run, record the
'   checks and print the BENCH line, then delete the rows written below
'   each sheet's extent before the run. In the used-range scenario a stray
'   comment is written STRAY_OFFSET rows below the end of the Trades used
'   range before the run, and its row deleted after it.
'
' INPUTS
'   CaseId: the case label.
'   LegCount: trade legs, 1 to MAX_LEGS.
'   InflateUsedRange: True for the used-range scenario.
'
' ERROR POLICY
'   An unexpected error propagates to RunBenchmark, which removes the
'   stray comment and the added rows.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim NettingSetCount     As Long       'Netting sets
    Dim StartTime           As Double     'Timer at the start of the writing
    Dim WriteSeconds        As Double     'Seconds to write the inputs
    Dim RunSeconds          As Double     'Seconds of the engine run
    Dim Completed           As Boolean    'CORE_Engine.Calculate returned True
    Dim Detail              As String     'Why the run did not complete; cleanup failure
    Dim UsedRangeEnd        As Long       'End of the Trades used range with the stray comment
    Dim LegsRead            As Long       'Legs the engine read
    Dim LegsUsed            As Long       'Legs the engine used
    Dim ValidSets           As Long       'VALID CRR netting sets
    Dim WithheldSets        As Long       'Netting sets whose EAD was withheld
    Dim ErrorLines          As Long       'ERROR lines on Checks
    Dim ResultsRowCount     As Long       'Rows on Results
    Dim TradeCalcRowCount   As Long       'Rows on TradeCalc
    Dim LastId              As String     'ID of the last Results row
    Dim LastTradeId         As String     'ID of the last TradeCalc row
    Dim BenchLine           As String     'BENCH line

'------------------------------------------------------------------------------
' WRITE THE INPUTS
'------------------------------------------------------------------------------
    'One netting set per LEGS_PER_SET legs, the last one possibly smaller.
        NettingSetCount = (LegCount + LEGS_PER_SET - 1) \ LEGS_PER_SET
        TEST_CaseRunner.BeginCase CaseId, RG_CRR, VALUATION, CALC_CCY
        StartTime = Timer
        WriteNettingSets NettingSetCount
        WriteTrades LegCount
        WriteSeconds = SecondsSince(StartTime)
        If InflateUsedRange Then
            mUsedBefore = UsedLastRow(GetSheet(SH_TRADES))
            mStrayRow = mUsedBefore + STRAY_OFFSET
            TableCell(GetSheet(SH_TRADES), mStrayRow, TR_COMMENT).Value = STRAY_TEXT
            UsedRangeEnd = UsedLastRow(GetSheet(SH_TRADES))
            TEST_CaseRunner.ExpectTrue "used-range", UsedRangeEnd >= mStrayRow, "the Trades used range ends on row " & _
                                       UsedRangeEnd & ", above the stray comment on row " & mStrayRow, REFERENCE_CLASS
        End If

'------------------------------------------------------------------------------
' RUN THE ENGINE
'------------------------------------------------------------------------------
        Completed = RunEngine(RunSeconds, Detail)

'------------------------------------------------------------------------------
' CHECK
'------------------------------------------------------------------------------
    'The run completed, every leg read and used, every netting set VALID,
    'no ERROR on Checks, one Results row per netting set and the TOTAL CRR
    'row, one TradeCalc row per leg.
        LegsRead = CORE_Engine.TradesRead
        LegsUsed = CORE_Engine.TradesUsed
        ValidSets = CORE_Engine.RegimeValidCount(RG_CRR)
        WithheldSets = CORE_Engine.IncompleteCount
        ErrorLines = CORE_Engine.ErrorCount
        ResultsRowCount = OutputRows(SH_RESULTS, RS_ID, LastId)
        TradeCalcRowCount = OutputRows(SH_TRADECALC, TC_ID, LastTradeId)
        TEST_CaseRunner.ExpectTrue "completed", Completed, Detail, REFERENCE_CLASS
        TEST_CaseRunner.ExpectTrue "legs-used", (LegsRead = LegCount) And (LegsUsed = LegCount), _
                                   "read " & LegsRead & " and used " & LegsUsed & " of " & LegCount & " legs", _
                                   REFERENCE_CLASS
        TEST_CaseRunner.ExpectTrue "netting-sets-valid", (ValidSets = NettingSetCount) And (WithheldSets = 0), _
                                   ValidSets & " of " & NettingSetCount & " netting sets VALID, " & WithheldSets & " withheld", _
                                   REFERENCE_CLASS
        TEST_CaseRunner.ExpectTrue "no-errors", ErrorLines = 0, ErrorLines & " ERROR line(s) on Checks", _
                                   REFERENCE_CLASS
        TEST_CaseRunner.ExpectTrue "results-rows", (ResultsRowCount = NettingSetCount + 1) And (LastId = "TOTAL " & RG_CRR), _
                                   ResultsRowCount & " Results rows ending with '" & LastId & "', expected " & _
                                   (NettingSetCount + 1) & " ending with 'TOTAL " & RG_CRR & "'", REFERENCE_CLASS
        TEST_CaseRunner.ExpectTrue "tradecalc-rows", TradeCalcRowCount = LegCount, _
                                   TradeCalcRowCount & " TradeCalc rows, expected " & LegCount, REFERENCE_CLASS

'------------------------------------------------------------------------------
' REPORT
'------------------------------------------------------------------------------
        BenchLine = "BENCH case=" & CaseId & "; legs=" & LegCount & "; netting_sets=" & NettingSetCount & _
                 "; used=" & LegsUsed & "; withheld=" & WithheldSets & "; errors=" & ErrorLines & _
                 "; warnings=" & CORE_Engine.WarningCount
        BenchLine = BenchLine & "; tradecalc_rows=" & TradeCalcRowCount & "; results_rows=" & ResultsRowCount
        If InflateUsedRange Then
            BenchLine = BenchLine & "; used_last_row=" & UsedRangeEnd
        End If
        BenchLine = BenchLine & "; write_s=" & Format$(WriteSeconds, "0.00") & "; run_s=" & Format$(RunSeconds, "0.00")
        TEST_Suite.Say BenchLine

'------------------------------------------------------------------------------
' REMOVE THE STRAY COMMENT
'------------------------------------------------------------------------------
        If InflateUsedRange Then
            Detail = ClearStray()
            TEST_CaseRunner.ExpectTrue "used-range-restored", Len(Detail) = 0, Detail, REFERENCE_CLASS
        End If

'------------------------------------------------------------------------------
' SHRINK THE SHEETS
'------------------------------------------------------------------------------
    'Rows the scenario wrote below the sheets' extent before the run are
    'deleted, so that the next scenario, and the workbook after EndSuite,
    'start from that extent rather than from a used range left by a
    'larger portfolio.
        Detail = ShrinkSheets()
        TEST_CaseRunner.ExpectTrue "extent-restored", Len(Detail) = 0, Detail, REFERENCE_CLASS

End Sub


Private Function RunEngine( _
    ByRef RunSeconds As Double, _
    ByRef Detail As String) _
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
'   True when the run completed. RunSeconds: the duration in seconds, also
'   when the run stopped or raised. Detail: why the run did not complete;
'   empty when it did.
'
' ERROR POLICY
'   An error raised by the engine is contained and returned in Detail; the
'   engine has already withdrawn its outputs.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim StartTime   As Double    'Timer at the start of the run

'------------------------------------------------------------------------------
' RUN
'------------------------------------------------------------------------------
        Detail = ""
        StartTime = Timer
        On Error GoTo Failed
        RunEngine = CORE_Engine.Calculate(True)
        RunSeconds = SecondsSince(StartTime)
        If Not RunEngine Then
            Detail = "engine run stopped; see the Checks sheet"
        End If
        Exit Function

'------------------------------------------------------------------------------
' HANDLE ERROR
'------------------------------------------------------------------------------
Failed:
        RunSeconds = SecondsSince(StartTime)
        Detail = "engine error " & Err.Number & ": " & Err.Description

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
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim TradesSheet    As Worksheet    'Trades
    Dim UsedRangeEnd   As Long         'Last row of the Trades used range after the deletion

'------------------------------------------------------------------------------
' REMOVE
'------------------------------------------------------------------------------
        If mStrayRow = 0 Then
            Exit Function
        End If
        On Error GoTo Failed
        Set TradesSheet = GetSheet(SH_TRADES)
        TableCell(TradesSheet, mStrayRow, TR_COMMENT).EntireRow.Delete
        mStrayRow = 0
        UsedRangeEnd = UsedLastRow(TradesSheet)
        If UsedRangeEnd > mUsedBefore Then
            ClearStray = "the Trades used range still ends on row " & UsedRangeEnd & _
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


Private Function ExtentSheets() As Variant
'
'==============================================================================
'                                 ExtentSheets
'------------------------------------------------------------------------------
' PURPOSE
'   Name the sheets whose used range a scenario can extend: the two input
'   sheets and the output sheets whose rows grow with the portfolio.
'
' UPDATED
'   2026-10-17
'==============================================================================
'
        ExtentSheets = Array(SH_TRADES, SH_NS, SH_TRADECALC, SH_RESULTS, SH_HEDGING, SH_BUCKETS, SH_CHECKS)

End Function


Private Sub RecordExtents()
'
'==============================================================================
'                                RecordExtents
'------------------------------------------------------------------------------
' PURPOSE
'   Record the last used row of each sheet of ExtentSheets before the
'   benchmark writes anything, for ShrinkSheets.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim SheetNames   As Variant    'ExtentSheets
    Dim k            As Long       'Index into SheetNames

'------------------------------------------------------------------------------
' RECORD
'------------------------------------------------------------------------------
        SheetNames = ExtentSheets()
        ReDim mBaseRows(LBound(SheetNames) To UBound(SheetNames))
        For k = LBound(SheetNames) To UBound(SheetNames)
            mBaseRows(k) = UsedLastRow(GetSheet(CStr(SheetNames(k))))
        Next k
        mHaveBase = True

End Sub


Private Function ShrinkSheets() As String
'
'==============================================================================
'                                 ShrinkSheets
'------------------------------------------------------------------------------
' PURPOSE
'   Give each sheet of ExtentSheets back the used range it had before the
'   benchmark: delete the rows below its recorded last row, which hold only
'   benchmark rows (the template's own rows are above it), and read
'   UsedRange again so that Excel recomputes it. Clearing contents is not
'   enough: Excel keeps the used range of cleared cells.
'
' RETURNS
'   Empty when every sheet is back to its recorded extent or none was
'   recorded; otherwise which sheets are not, or the error raised.
'
' ERROR POLICY
'   Contains its own error, so that it can run on the error path.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim SheetNames     As Variant      'ExtentSheets
    Dim k              As Long         'Index into SheetNames
    Dim TargetSheet    As Worksheet    'Sheet being shrunk
    Dim UsedRangeEnd   As Long         'Its last used row

'------------------------------------------------------------------------------
' SHRINK
'------------------------------------------------------------------------------
        On Error GoTo Failed
        If Not mHaveBase Then
            Exit Function
        End If
        SheetNames = ExtentSheets()
        For k = LBound(SheetNames) To UBound(SheetNames)
            Set TargetSheet = GetSheet(CStr(SheetNames(k)))
            UsedRangeEnd = UsedLastRow(TargetSheet)
            If UsedRangeEnd > mBaseRows(k) Then
                TargetSheet.Range(TargetSheet.Rows(mBaseRows(k) + 1), TargetSheet.Rows(UsedRangeEnd)).Delete
                UsedRangeEnd = UsedLastRow(TargetSheet)
                If UsedRangeEnd > mBaseRows(k) Then
                    ShrinkSheets = ShrinkSheets & IIf(Len(ShrinkSheets) > 0, "; ", "") & SheetNames(k) & _
                                   " used range ends on row " & UsedRangeEnd & ", not " & mBaseRows(k)
                End If
            End If
        Next k
        Exit Function

'------------------------------------------------------------------------------
' HANDLE ERROR
'------------------------------------------------------------------------------
Failed:
        ShrinkSheets = "sheet extents not restored: error " & Err.Number & ": " & Err.Description

End Function


'
'------------------------------------------------------------------------------
'
'                                  PORTFOLIO
'
'------------------------------------------------------------------------------
'

Private Sub WriteNettingSets( _
    ByVal NettingSetCount As Long)
'
'==============================================================================
'                               WriteNettingSets
'------------------------------------------------------------------------------
' PURPOSE
'   Write the netting sets as one block from the first NettingSets row,
'   with the value types TEST_CaseRunner.AddNettingSet writes. Odd ones are
'   unmargined, without collateral; even ones are margined with daily
'   remargining, VM held of 0, 100,000 or 200,000, NICA and an MTA, and
'   none is large or illiquid, in dispute or client cleared.
'
' INPUTS
'   NettingSetCount: netting sets, at least 1.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Block()   As Variant    'NettingSets rows, (netting set, column)
    Dim s         As Long       'Netting set, from 1
    Dim TopRow    As Long       'First NettingSets data row

'------------------------------------------------------------------------------
' BUILD
'------------------------------------------------------------------------------
    'Unused cells stay Empty, as the blanks AddNettingSet writes.
        ReDim Block(1 To NettingSetCount, 1 To NS_NCOLS)
        For s = 1 To NettingSetCount
            Block(s, NS_ID) = NettingSetId(s)
            Block(s, NS_CPTY) = "Benchmark counterparty " & s
            Block(s, NS_CLEARED) = CL_NONE
            Block(s, NS_LARGE) = "N"
            Block(s, NS_DISPUTE) = "N"
            Block(s, NS_REGIME) = RG_CRR
            If (s Mod 2) = 0 Then
                Block(s, NS_MARGINED) = "Y"
                Block(s, NS_FREQ) = 1#
                Block(s, NS_VM) = 100000# * (s Mod 3)
                Block(s, NS_NICA) = 50000#
                Block(s, NS_TH) = 0#
                Block(s, NS_MTA) = 10000#
            Else
                Block(s, NS_MARGINED) = "N"
                Block(s, NS_VM) = 0#
                Block(s, NS_NICA) = 0#
                Block(s, NS_TH) = 0#
                Block(s, NS_MTA) = 0#
            End If
        Next s

'------------------------------------------------------------------------------
' WRITE
'------------------------------------------------------------------------------
        TopRow = FirstInputRow(SH_NS)
        TableRange(GetSheet(SH_NS), TopRow, 1, TopRow + NettingSetCount - 1, NS_NCOLS).Value = Block

End Sub


Private Sub WriteTrades( _
    ByVal LegCount As Long)
'
'==============================================================================
'                                 WriteTrades
'------------------------------------------------------------------------------
' PURPOSE
'   Write the legs as one block from the first Trades row, with the value
'   types TEST_CaseRunner.AddTrade writes.
'
' INPUTS
'   LegCount: trade legs, 1 to MAX_LEGS.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Block()   As Variant    'Trades rows, (leg, column)
    Dim i         As Long       'Leg, from 1
    Dim TopRow    As Long       'First Trades data row

'------------------------------------------------------------------------------
' BUILD AND WRITE
'------------------------------------------------------------------------------
        ReDim Block(1 To LegCount, 1 To TR_NCOLS)
        For i = 1 To LegCount
            FillLeg Block, i
        Next i
        TopRow = FirstInputRow(SH_TRADES)
        TableRange(GetSheet(SH_TRADES), TopRow, 1, TopRow + LegCount - 1, TR_NCOLS).Value = Block

End Sub


Private Sub FillLeg( _
    ByRef Block() As Variant, _
    ByVal Leg As Long)
'
'==============================================================================
'                                   FillLeg
'------------------------------------------------------------------------------
' PURPOSE
'   Fill row Leg of the Trades block with that leg of the mix described in
'   the module header. Amounts are in EUR: notionals of 1 to 9 million, market
'   values from -15,000 to 15,000. The economic-trade, lambda and CDO
'   columns and the comment stay Empty.
'
' INPUTS
'   Block: the Trades rows, (leg, column), from WriteTrades.
'   Leg: the leg, from 1; netting set (Leg - 1) \ LEGS_PER_SET + 1.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim LegOffset   As Long    'Leg, from 0
    Dim Slot        As Long    'Position in the mix, 0 to MIX_LENGTH - 1
    Dim Cycle       As Long    'Repetitions of the mix before this leg
    Dim Entity      As Long    'Credit or equity entity, from 0

'------------------------------------------------------------------------------
' COMMON FIELDS
'------------------------------------------------------------------------------
        LegOffset = Leg - 1
        Slot = LegOffset Mod MIX_LENGTH
        Cycle = LegOffset \ MIX_LENGTH
        Block(Leg, TR_ID) = "B" & Format$(Leg, "000000")
        Block(Leg, TR_NS) = NettingSetId(LegOffset \ LEGS_PER_SET + 1)
        Block(Leg, TR_INSTR) = "Linear"
        Block(Leg, TR_DIR) = IIf(((LegOffset \ 3) Mod 2) = 0, "Long", "Short")
        Block(Leg, TR_NATURE) = "Standard"
        Block(Leg, TR_NOTIONAL) = 1000000# * (1 + (LegOffset Mod 9))
        Block(Leg, TR_NCCY) = CALC_CCY
        Block(Leg, TR_MTM) = 1000# * ((LegOffset Mod 7) - 3) * (1 + (LegOffset Mod 5))
        Block(Leg, TR_MCCY) = CALC_CCY

'------------------------------------------------------------------------------
' ASSET-CLASS FIELDS
'------------------------------------------------------------------------------
    'Interest-rate swaps end from 6 months to 20 years, never exactly on the
    'bucket edges of 1 and 5 years. FX pairs all hold EUR, the notional
    'currency, and are in market quote order, as an FX option requires. A
    'credit, equity or commodity reference is always given the same
    'sub-class, as the engine requires within a netting set. An other-risks
    'leg has no sub-class; its index is its hedging set (Art. 277a CRR), and
    'cycling the indices every 3 mixes while the direction alternates every
    'mix puts long and short legs in each.
        Select Case Slot
            Case 0 To 6
                Block(Leg, TR_AC) = "IR"
                Block(Leg, TR_RF) = IIf(((Cycle + Slot) Mod 2) = 0, "EUR", "USD")
                SetMaturity Block, Leg, Choose(((Cycle * 7 + Slot) Mod 9) + 1, 6, 9, 18, 30, 42, 84, 120, 180, 240), _
                            True
            Case 7 To 9
                Block(Leg, TR_AC) = "FX"
                Block(Leg, TR_RF) = FxPair(Cycle + Slot)
                SetMaturity Block, Leg, Choose(((Cycle + Slot) Mod 5) + 1, 3, 6, 12, 18, 24), False
            Case 10
                Block(Leg, TR_AC) = "FX"
                Block(Leg, TR_RF) = FxPair(Cycle)
                SetOption Block, Leg, Cycle, Choose((Cycle Mod 4) + 1, 1.16, 0.74, 172#, 0.93)
            Case 11 To 13
                Entity = (Cycle * 3 + Slot - 11) Mod CREDIT_NAMES
                Block(Leg, TR_AC) = "CR"
                Block(Leg, TR_SUB) = Choose((Entity Mod 7) + 1, "AAA", "AA", "A", "BBB", "BB", "B", "CCC")
                Block(Leg, TR_RF) = "CR NAME " & Format$(Entity + 1, "00")
                SetMaturity Block, Leg, Choose(((Cycle + Slot) Mod 5) + 1, 12, 36, 60, 84, 120), True
            Case 14
                Block(Leg, TR_AC) = "EQ"
                Block(Leg, TR_SUB) = "SINGLE"
                Block(Leg, TR_RF) = "EQ NAME " & Format$((Cycle * 2) Mod EQUITY_NAMES + 1, "00")
                SetMaturity Block, Leg, Choose((Cycle Mod 4) + 1, 3, 6, 12, 24), False
            Case 15
                Block(Leg, TR_AC) = "EQ"
                Block(Leg, TR_SUB) = "INDEX"
                Block(Leg, TR_RF) = "EQ INDEX " & (Cycle Mod 3 + 1)
                SetMaturity Block, Leg, Choose(((Cycle + 1) Mod 4) + 1, 3, 6, 12, 24), False
            Case 16
                Block(Leg, TR_AC) = "EQ"
                Block(Leg, TR_SUB) = "SINGLE"
                Block(Leg, TR_RF) = "EQ NAME " & Format$((Cycle * 2 + 1) Mod EQUITY_NAMES + 1, "00")
                SetOption Block, Leg, Cycle, 100#
            Case 17, 19
                Block(Leg, TR_AC) = "CO"
                Block(Leg, TR_SUB) = "OIL_GAS"
                Block(Leg, TR_RF) = Choose(((Cycle + Slot) Mod 3) + 1, "BRENT", "WTI", "NATURAL GAS")
                SetMaturity Block, Leg, Choose(((Cycle + Slot) Mod 4) + 1, 3, 6, 12, 24), False
            Case 18
                Block(Leg, TR_AC) = "CO"
                Block(Leg, TR_SUB) = "METALS"
                Block(Leg, TR_RF) = Choose((Cycle Mod 3) + 1, "GOLD", "SILVER", "COPPER")
                SetMaturity Block, Leg, Choose(((Cycle + 2) Mod 4) + 1, 3, 6, 12, 24), False
            Case Else
                Block(Leg, TR_AC) = "OT"
                Block(Leg, TR_RF) = Choose((Cycle Mod OTHER_INDICES) + 1, "LONGEVITY INDEX", "CAT LOSS INDEX", _
                                           "WEATHER INDEX")
                SetMaturity Block, Leg, Choose(((Cycle + 1) Mod 5) + 1, 6, 12, 24, 36, 60), False
        End Select

End Sub


Private Sub SetMaturity( _
    ByRef Block() As Variant, _
    ByVal Leg As Long, _
    ByVal Months As Long, _
    ByVal WithEndDate As Boolean)
'
'==============================================================================
'                                 SetMaturity
'------------------------------------------------------------------------------
' PURPOSE
'   Give leg Leg a maturity date Months after the valuation date, and the
'   same end date when WithEndDate is True (swaps and credit default
'   swaps). The start date stays Empty: the leg starts on the valuation
'   date.
'
' INPUTS
'   Block, Leg: as FillLeg.
'   Months: months to maturity, 1 or more.
'   WithEndDate: True to write the end date too.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim MaturityDate   As Date    'Maturity date

'------------------------------------------------------------------------------
' WRITE
'------------------------------------------------------------------------------
        MaturityDate = MonthsAfterValuation(Months)
        Block(Leg, TR_MAT) = MaturityDate
        If WithEndDate Then
            Block(Leg, TR_END) = MaturityDate
        End If

End Sub


Private Sub SetOption( _
    ByRef Block() As Variant, _
    ByVal Leg As Long, _
    ByVal Cycle As Long, _
    ByVal Price As Double)
'
'==============================================================================
'                                  SetOption
'------------------------------------------------------------------------------
' PURPOSE
'   Make leg Leg a European option on its underlying: a call or a put, with
'   the underlying price, a strike 5% below, at or 5% above it, and an
'   expiry 3, 6 or 12 months after the valuation date, which is also its
'   maturity. The lambda shift stays Empty: it applies to interest-rate and
'   commodity options only.
'
' INPUTS
'   Block, Leg: as FillLeg.
'   Cycle: the leg's repetition of the mix, which sets the type, strike and
'   expiry.
'   Price: the underlying price P, > 0.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ExpiryDate   As Date    'Expiry and maturity date

'------------------------------------------------------------------------------
' WRITE
'------------------------------------------------------------------------------
        ExpiryDate = MonthsAfterValuation(Choose(((Cycle + 1) Mod 3) + 1, 3, 6, 12))
        Block(Leg, TR_INSTR) = "Option"
        Block(Leg, TR_OPT) = IIf((Cycle Mod 2) = 0, "Call", "Put")
        Block(Leg, TR_PRICE) = Price
        Block(Leg, TR_STRIKE) = Price * Choose((Cycle Mod 3) + 1, 0.95, 1#, 1.05)
        Block(Leg, TR_EXPIRY) = ExpiryDate
        Block(Leg, TR_MAT) = ExpiryDate

End Sub


'
'------------------------------------------------------------------------------
'
'                                   HELPERS
'
'------------------------------------------------------------------------------
'

Private Function ParseSizes( _
    ByVal SizesText As String, _
    ByRef Sizes() As Long, _
    ByRef Reason As String) _
    As Boolean
'
'==============================================================================
'                                  ParseSizes
'------------------------------------------------------------------------------
' PURPOSE
'   Read the sizes given to RunBenchmark.
'
' INPUTS
'   SizesText: whole numbers separated by commas; spaces around them are
'   ignored.
'
' RETURNS
'   True with Sizes (1 To the number of sizes) filled when there are 1 to
'   MAX_SCENARIOS sizes, each a whole number from 1 to MAX_LEGS; otherwise
'   False with Reason saying what is wrong.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Parts()     As String    'Items between the commas
    Dim SizeCount   As Long      'Items
    Dim i           As Long      'Size, from 1
    Dim SizeText    As String    'One item, trimmed
    Dim ValidForm   As String    'What a valid argument is

'------------------------------------------------------------------------------
' READ
'------------------------------------------------------------------------------
        ValidForm = "give 1 to " & MAX_SCENARIOS & " whole numbers of legs from 1 to " & MAX_LEGS & _
               ", separated by commas, such as ""1000,10000,50000"""
        Parts = Split(SizesText, ",")
        SizeCount = UBound(Parts) - LBound(Parts) + 1
        If SizeCount < 1 Or SizeCount > MAX_SCENARIOS Then
            Reason = ValidForm & " (found """ & SizesText & """)"
            Exit Function
        End If
        ReDim Sizes(1 To SizeCount)
        For i = 1 To SizeCount
            SizeText = Trim$(Parts(LBound(Parts) + i - 1))
            If Len(SizeText) = 0 Or Len(SizeText) > Len(CStr(MAX_LEGS)) Or SizeText Like "*[!0-9]*" Then
                Reason = ValidForm & " (found """ & SizeText & """)"
                Exit Function
            End If
            Sizes(i) = CLng(SizeText)
            If Sizes(i) < 1 Or Sizes(i) > MAX_LEGS Then
                Reason = ValidForm & " (found " & Sizes(i) & ")"
                Exit Function
            End If
        Next i
        ParseSizes = True

End Function


Private Function FirstInputRow( _
    ByVal SheetName As String) _
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
'   2026-10-17
'==============================================================================
'
        FirstInputRow = TableFirstRow(SheetName)
        If FirstInputRow = 0 Then
            Err.Raise ERR_TEST_SETUP, "TEST_Benchmark.FirstInputRow", _
                      "The header of " & SheetName & " cannot be located through its workbook name - " & _
                      "rebuild the workbook from the template."
        End If

End Function


Private Function OutputRows( _
    ByVal SheetName As String, _
    ByVal IdColumn As Long, _
    ByRef LastId As String) _
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
'   SheetName: SH_RESULTS or SH_TRADECALC.
'   IdColumn: the table column of the row ID.
'
' RETURNS
'   The number of rows, 0 when there are none. LastId: the ID of the last
'   row; empty when there are none.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim OutputSheet   As Worksheet    'Output sheet
    Dim FirstRow      As Long         'First data row
    Dim LastRow       As Long         'Last row with an ID

'------------------------------------------------------------------------------
' COUNT
'------------------------------------------------------------------------------
        Set OutputSheet = GetSheet(SheetName)
        FirstRow = OutputFirstRow(SheetName)
        LastRow = LastDataRow(OutputSheet, FirstRow, IdColumn)
        OutputRows = LastRow - FirstRow + 1
        LastId = ""
        If LastRow >= FirstRow Then
            LastId = SafeStr(TableCell(OutputSheet, LastRow, IdColumn).Value)
        End If

End Function


Private Function NettingSetId( _
    ByVal SetNumber As Long) _
    As String
'
'==============================================================================
'                                 NettingSetId
'------------------------------------------------------------------------------
' PURPOSE
'   ID of netting set SetNumber, from 1: BENCH-001, BENCH-002 and so on.
'
' UPDATED
'   2026-10-17
'==============================================================================
'
        NettingSetId = "BENCH-" & Format$(SetNumber, "000")

End Function


Private Function FxPair( _
    ByVal PairIndex As Long) _
    As String
'
'==============================================================================
'                                    FxPair
'------------------------------------------------------------------------------
' PURPOSE
'   One of the four currency pairs of the portfolio, in market quote order:
'   EUR/USD, EUR/GBP, EUR/JPY or EUR/CHF for PairIndex 0, 1, 2 or 3
'   modulo 4.
'   Each currency is in the template's FX table.
'
' UPDATED
'   2026-10-17
'==============================================================================
'
        FxPair = Choose((PairIndex Mod 4) + 1, "EUR/USD", "EUR/GBP", "EUR/JPY", "EUR/CHF")

End Function


Private Function MonthsAfterValuation( _
    ByVal Months As Long) _
    As Date
'
'==============================================================================
'                             MonthsAfterValuation
'------------------------------------------------------------------------------
' PURPOSE
'   The date Months calendar months after VALUATION, on the same day or
'   the last day of a shorter month.
'
' UPDATED
'   2026-10-17
'==============================================================================
'
        MonthsAfterValuation = DateAdd("m", Months, DateSerial(CInt(Left$(VALUATION, 4)), _
                                                               CInt(Mid$(VALUATION, 6, 2)), _
                                                               CInt(Right$(VALUATION, 2))))

End Function


Private Function SecondsSince( _
    ByVal StartTime As Double) _
    As Double
'
'==============================================================================
'                                 SecondsSince
'------------------------------------------------------------------------------
' PURPOSE
'   Seconds elapsed since StartTime, a value of Timer, also across
'   midnight, when Timer starts again from 0.
'
' UPDATED
'   2026-10-17
'==============================================================================
'
        SecondsSince = Timer - StartTime
        If SecondsSince < 0# Then
            SecondsSince = SecondsSince + 86400#
        End If

End Function


'
'------------------------------------------------------------------------------
'
'                                 CSV READER
'
'------------------------------------------------------------------------------
'

Public Sub RunCsvBenchmark( _
    Optional ByVal SizesText As String = "10000,100000")
'
'==============================================================================
'                               RunCsvBenchmark
'------------------------------------------------------------------------------
' PURPOSE
'   Measure how long CORE_CsvReader takes to read synthetic CSV files
'   (#207): one Trades file per size, then a NettingSets file of
'   CSV_BENCH_NETTING_SETS records. Each prints a BENCH_CSV line with the
'   file's bytes, the seconds taken to build and write it and the seconds
'   taken to read it, and checks the outcome: a file within the record
'   limit is accepted whole; a larger one is rejected for its record
'   count. A final BENCH_CSV_RESULT line says whether every outcome was
'   as expected. No sheet is touched, and no file that was there before:
'   each file gets a name not already taken.
'
' INPUTS
'   SizesText: the Trades records of each file, as for RunBenchmark;
'   10,000 and 100,000, the contract's limit, when omitted.
'
' ERROR POLICY
'   Sizes that cannot be read print a BENCH_REJECTED line. An unexpected
'   error is printed as BENCH_CSV_RESULT=FAIL, after CsvScenario has
'   closed and deleted its file.
'
' UPDATED
'   2026-10-23
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Sizes()     As Long       'Records of each Trades file
    Dim Reason      As String     'Why SizesText was rejected
    Dim AllAsExpected As Boolean  'Every outcome was as expected
    Dim Folder      As String     'Folder of the benchmark files
    Dim i           As Long       'Index into Sizes

'------------------------------------------------------------------------------
' RUN
'------------------------------------------------------------------------------
        If Not ParseSizes(SizesText, Sizes, Reason) Then
            TEST_Suite.Say "BENCH_REJECTED=" & Reason
            Exit Sub
        End If
        Folder = Environ$("TEMP")
        If Len(Folder) = 0 Then
            Folder = ThisWorkbook.Path
        End If
        On Error GoTo Failed
        AllAsExpected = True
        For i = 1 To UBound(Sizes)
            AllAsExpected = CsvScenario(Folder, "Trades", Sizes(i)) And AllAsExpected
        Next i
        AllAsExpected = CsvScenario(Folder, "NettingSets", CSV_BENCH_NETTING_SETS) And AllAsExpected
        TEST_Suite.Say "BENCH_CSV_RESULT=" & IIf(AllAsExpected, "PASS", "FAIL")
        Exit Sub

Failed:
        TEST_Suite.Say "BENCH_CSV_RESULT=FAIL; unexpected error " & Err.Number & ": " & Err.Description

End Sub


Private Function CsvScenario( _
    ByVal Folder As String, _
    ByVal Kind As String, _
    ByVal RecordCount As Long) _
    As Boolean
'
'==============================================================================
'                                 CsvScenario
'------------------------------------------------------------------------------
' PURPOSE
'   Write a synthetic file of RecordCount records of Kind to Folder, under
'   a name not already taken, read it with CORE_CsvReader.ReadCsvFile,
'   print a BENCH_CSV line and delete the file.
'
' RETURNS
'   True when the outcome is the expected one: accepted with every record
'   within the contract's record limit, rejected for its record count
'   above it.
'
' ERROR POLICY
'   Raises the first error after closing the file number and deleting
'   the file, if it was created; an error in that cleanup does not
'   replace the first one.
'
' UPDATED
'   2026-10-23
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Records()   As String           'Header and records, one line each
    Dim Bytes()     As Byte             'The file
    Dim TargetFile  As String           'Its path
    Dim n           As Long             'Counter that makes the name unique
    Dim FileNumber  As Integer          'File number; 0 while none is open
    Dim Created     As Boolean          'The file was created
    Dim ErrorNumber As Long             'First error
    Dim ErrorSource As String           'Its source
    Dim ErrorText   As String           'Its description
    Dim StartTime   As Double           'Timer at the start of a step
    Dim BuildSeconds As Double          'Seconds to build and write the file
    Dim ReadSeconds As Double           'Seconds to read it
    Dim Result      As CsvReadResult    'What the reader found
    Dim Values()    As String           'The records read
    Dim Lines()     As Long             'Their lines
    Dim Findings()  As CsvFinding       'The findings
    Dim Limit       As Long             'The contract's record limit
    Dim Outcome     As String           'What the reader returned
    Dim r           As Long             'Record counter

'------------------------------------------------------------------------------
' BUILD AND WRITE
'------------------------------------------------------------------------------
        StartTime = Timer
        ReDim Records(0 To RecordCount)
        Records(0) = IIf(Kind = "Trades", CSV_TR_HEADER, CSV_NS_HEADER)
        For r = 1 To RecordCount
            Records(r) = CsvRecord(Kind, r)
        Next r
        Bytes = CORE_Export.Utf8Bytes(Join(Records, vbLf) & vbLf)
        Erase Records
        Do
            n = n + 1
            TargetFile = Folder & Application.PathSeparator & CSV_BENCH_FILE & "_" & n & ".csv"
        Loop While Len(Dir$(TargetFile, vbHidden Or vbSystem)) > 0
        On Error GoTo Failed
        FileNumber = FreeFile
        Open TargetFile For Binary Access Write As #FileNumber
        Created = True
        Put #FileNumber, 1, Bytes
        Close #FileNumber
        FileNumber = 0
        BuildSeconds = SecondsSince(StartTime)

'------------------------------------------------------------------------------
' READ AND REPORT
'------------------------------------------------------------------------------
        StartTime = Timer
        ReadCsvFile TargetFile, Kind, Result, Values, Lines, Findings
        ReadSeconds = SecondsSince(StartTime)
        Limit = IIf(Kind = "Trades", CSV_BENCH_TRADES_LIMIT, CSV_BENCH_NETTING_SETS_LIMIT)
        If Result.Accepted Then
            Outcome = "ACCEPTED " & Result.RecordCount
            CsvScenario = (RecordCount <= Limit And Result.RecordCount = RecordCount)
        Else
            Outcome = "REJECTED " & FindingText(Findings(1))
            CsvScenario = (RecordCount > Limit And Result.FindingCount = 1 And _
                           Findings(1).Code = "CSV-TOO-MANY-RECORDS")
        End If
        TEST_Suite.Say "BENCH_CSV kind=" & Kind & "; records=" & RecordCount & _
                       "; bytes=" & (UBound(Bytes) + 1) & _
                       "; build_seconds=" & Format$(BuildSeconds, "0.00") & _
                       "; read_seconds=" & Format$(ReadSeconds, "0.00") & _
                       "; outcome=" & Outcome & "; " & IIf(CsvScenario, "PASS", "FAIL")
        Kill TargetFile
        Exit Function

'------------------------------------------------------------------------------
' HANDLE ERROR
'------------------------------------------------------------------------------
Failed:
        ErrorNumber = Err.Number
        ErrorSource = Err.Source
        ErrorText = Err.Description
        On Error Resume Next
        If FileNumber <> 0 Then
            Close #FileNumber
        End If
        If Created Then
            Kill TargetFile
        End If
        On Error GoTo 0
        Err.Raise ErrorNumber, ErrorSource, ErrorText

End Function


Private Function CsvRecord( _
    ByVal Kind As String, _
    ByVal Index As Long) _
    As String
'
'==============================================================================
'                                  CsvRecord
'------------------------------------------------------------------------------
' PURPOSE
'   One synthetic record of Kind, valid under the contract and unique by
'   its key: for Trades an interest-rate swap leg of one of
'   CSV_BENCH_SETS netting sets with a comment, for NettingSets a margined
'   netting set with a counterparty name.
'
' UPDATED
'   2026-10-23
'==============================================================================
'
        If Kind = "Trades" Then
            CsvRecord = "T" & Index & ";NS-" & (Index Mod CSV_BENCH_SETS) & ";IR;;EUR;Linear;" & _
                        IIf(Index Mod 2 = 0, "Long", "Short") & ";;;;" & (1000000 + Index) & _
                        ";EUR;1250.5;EUR;2026-01-02;2031-01-02;2031-01-02;;;;;;;;;;;Synthetic leg " & Index
        Else
            CsvRecord = "NS-" & Index & ";Counterparty " & Index & ";Y;N;1;N;N;;0;0;0;0;;;;;"
        End If

End Function
