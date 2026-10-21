Attribute VB_Name = "TEST_Invariants"
'==============================================================================
' MODULE: TEST_Invariants
'------------------------------------------------------------------------------
' PURPOSE
'   Check relations that every result must satisfy, whatever the inputs
'   (#44): EAD = alpha * (RC + PFE); RC is never negative and follows the
'   margined or unmargined formula; the multiplier follows its formula,
'   stays between the floor and 1, and is 1 when V - C >= 0; PFE is the
'   multiplier times the aggregate add-on; a margined netting set's EAD is
'   the lower of the margined EAD and the unmargined cap; withheld netting
'   sets show no figures; each regime has a TOTAL row (TOTAL CRR, TOTAL
'   BCBS) that holds, for each of V, C, RC, the six asset-class add-ons,
'   the aggregate add-on, PFE and EAD, the sum over that regime's VALID
'   netting sets, says in its status how many are VALID and withheld, and
'   holds nothing else (#198). Before these relations it checks that every
'   figure they read is a number and that the demo portfolio covers every
'   branch they test. It also checks that a second run gives identical
'   outputs; that with the IR sum of absolute bucket values a CRR netting
'   set with interest-rate trades becomes INVALID, while every other
'   netting set VALID before stays VALID and has an IR add-on not below
'   the bucket formula, and at least one has a higher one; and that a
'   netting set with a zero add-on is calculated without a division
'   failure. Two checks corrupt copies of the outputs and confirm that the
'   TOTAL and IR relations then fail and name the corrupted column or
'   netting set (#103, #104).
'
' PUBLIC SURFACE
'   RunInvariantTests is the entry point. Option Private Module keeps it
'   out of the external workbook automation API.
'
' DEPENDENCIES
'   TEST_CaseRunner, which writes the inputs, runs the engine, records the
'   checks and restores the workbook; the CORE_Util readers (GetParam,
'   ToDbl, SafeStr, FindHeaderRow, TableCell, ...) and the CORE_Config
'   sheet names, column indexes and parameter codes. The relations are
'   recomputed here
'   from the Results columns and the NettingSets inputs, without calling
'   the engine's formulas. They compare outputs with each other, so they
'   need no expected value from outside and are reported as illustrative.
'   The demo cases use the workbook's own inputs, so the workbook must be
'   built from the template.
'
' WORKSHEET SAFETY
'   As TEST_CaseRunner: inputs, the IRBucketOffset parameter and the
'   output sheets are rewritten during the run and restored at the end.
'   Use a development workbook.
'
' USAGE
'   Run TEST_Invariants.RunInvariantTests from the Immediate window.
'
' REFERENCE
'   CRE52.1 and CRR Art. 274(2) (EAD), CRE52.10 to 52.19 and CRR Art. 275
'   (RC), CRE52.20 (PFE), CRE52.23 and CRR Art. 278(3) (multiplier),
'   CRE52.25 (aggregate add-on), CRE52.2 and CRR Art. 274(3) (cap of a margined netting set),
'   CRE52.57 and CRR Art. 280a(3) (IR bucket formula, the only one the
'   CRR allows).
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
    Option Explicit
    Option Private Module

'------------------------------------------------------------------------------
' MODULE CONSTANTS
'------------------------------------------------------------------------------
    'Suite size and tolerance.
        Private Const VALUATION   As String = "2026-09-30"    'Valuation date of the built case
        Private Const CASES       As Long = 4                 'Cases in a complete run
        Private Const CHECKS      As Long = 23                'Checks in a complete run
        Private Const REL_TOL     As Double = 0.000000001     'Relative tolerance, on max(1, |expected|)

    'Results columns, as written by CORE_Engine.
        Private Const COL_REGIME      As Long = 3     'CRR or BCBS
        Private Const COL_MARGINED    As Long = 4     'Y or N
        Private Const COL_V           As Long = 7     'V, sum of market values
        Private Const COL_C           As Long = 8     'C = VM + NICA
        Private Const COL_RC          As Long = 9     'Replacement cost
        Private Const COL_IR_ADDON    As Long = 10    'IR add-on; FX, CR, EQ, CO and OT follow
        Private Const COL_OT_ADDON    As Long = 15    'OT add-on, the last asset class
        Private Const COL_AGGREGATE   As Long = 16    'Aggregate add-on
        Private Const COL_MULTIPLIER  As Long = 17    'Multiplier
        Private Const COL_PFE         As Long = 18    'Potential future exposure
        Private Const COL_ALPHA       As Long = 19    'Alpha
        Private Const COL_EAD_MARGIN  As Long = 20    'EAD margined, before the cap
        Private Const COL_EAD_CAP     As Long = 21    'EAD on the unmargined basis: the cap
        Private Const COL_EAD         As Long = 22    'EAD reported
        Private Const COL_CAP_FLAG    As Long = 23    'Y, N or n/a
        Private Const COL_C_CAP       As Long = 24    'C on the cap basis
        Private Const COL_RC_CAP      As Long = 25    'RC on the cap basis
        Private Const COL_ADDON_CAP   As Long = 26    'Aggregate add-on, unmargined
        Private Const COL_MULT_CAP    As Long = 27    'Multiplier on the cap basis

'------------------------------------------------------------------------------
' MODULE STATE
'------------------------------------------------------------------------------
        Private mFloor   As Double    'Multiplier floor from Params, decimal


'
'------------------------------------------------------------------------------
'
'                                 ENTRY POINT
'
'------------------------------------------------------------------------------
'

Public Sub RunInvariantTests()
'
'==============================================================================
'                              RunInvariantTests
'------------------------------------------------------------------------------
' PURPOSE
'   Run every case and print the TEST_CaseRunner report.
'
' ERROR POLICY
'   An unexpected error ends the suite through EndSuite, which restores
'   the workbook and reports the run as failed.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim DemoRows       As Variant    'Results rows of the first demo run
    Dim TradeRows      As Variant    'TradeCalc rows of the first demo run
    Dim HedgeRows      As Variant    'HedgingSets rows of the first demo run
    Dim BucketRows     As Variant    'Buckets rows of the first demo run
    Dim AbsoluteRows   As Variant    'Results rows with the sum of absolute bucket values
    Dim Detail         As String     'First violation; "" when none

'------------------------------------------------------------------------------
' RELATIONS ON THE DEMO PORTFOLIO
'------------------------------------------------------------------------------
        TEST_CaseRunner.BeginSuite CASES, CHECKS
        On Error GoTo Failed

        TEST_CaseRunner.BeginCase "invariants-demo", "CRR", VALUATION, "EUR"
        TEST_CaseRunner.LoadSavedInputs
        TEST_CaseRunner.RunCase
        mFloor = ToDbl(GetParam(PRM_FLOOR), -1#)
        DemoRows = SheetRows(SH_RESULTS, RS_NCOLS)
        TradeRows = SheetRows(SH_TRADECALC, TC_NCOLS)
        HedgeRows = SheetRows(SH_HEDGING, HS_NCOLS)
        BucketRows = SheetRows(SH_BUCKETS, BK_NCOLS)
        RecordRelation "numbers", NumberGap(DemoRows)
        RecordRelation "coverage", CoverageGap(DemoRows)
        RecordRelation "aggregate-add-on", AggregateGap(DemoRows)
        RecordRelation "replacement-cost", ReplacementCostGap(DemoRows)
        RecordRelation "multiplier", MultiplierGap(DemoRows)
        RecordRelation "potential-future-exposure", PfeGap(DemoRows)
        RecordRelation "exposure-value", ExposureGap(DemoRows)
        RecordRelation "unmargined-cap", CapGap(DemoRows)
        RecordRelation "withheld", WithheldGap(DemoRows)
        RecordRelation "total", TotalGap(DemoRows)
        RecordRelation "total-catches-corruption", TotalCorruptionGap(DemoRows)

'------------------------------------------------------------------------------
' REPEATABILITY
'------------------------------------------------------------------------------
        TEST_CaseRunner.BeginCase "invariants-demo-repeat", "CRR", VALUATION, "EUR"
        TEST_CaseRunner.LoadSavedInputs
        TEST_CaseRunner.RunCase
        RecordRelation "same-results", _
                       IdenticalGap(DemoRows, SheetRows(SH_RESULTS, RS_NCOLS), SH_RESULTS)
        RecordRelation "same-trade-calc", _
                       IdenticalGap(TradeRows, SheetRows(SH_TRADECALC, TC_NCOLS), SH_TRADECALC)
        RecordRelation "same-hedging-sets", _
                       IdenticalGap(HedgeRows, SheetRows(SH_HEDGING, HS_NCOLS), SH_HEDGING)
        RecordRelation "same-buckets", _
                       IdenticalGap(BucketRows, SheetRows(SH_BUCKETS, BK_NCOLS), SH_BUCKETS)

'------------------------------------------------------------------------------
' IR BUCKET OFFSET
'------------------------------------------------------------------------------
    'The sum of absolute bucket values can never be below the bucket
    'formula, whose cross terms have correlations below 1 [CRE52.57]. The
    'CRR allows only the formula [Art. 280a(3)], so a CRR netting set with
    'interest-rate trades must become INVALID.
        TEST_CaseRunner.BeginCase "invariants-demo-ir-sum-of-absolutes", "CRR", VALUATION, "EUR"
        TEST_CaseRunner.LoadSavedInputs
        TEST_CaseRunner.PatchCell SH_PARAMS, ParamsRow(PRM_IRFULL), ParamsValueCol(), False
        TEST_CaseRunner.RunCase
        AbsoluteRows = SheetRows(SH_RESULTS, RS_NCOLS)
        Detail = IrOffsetGap(DemoRows, AbsoluteRows, HedgeRows)
        RecordRelation "ir-add-on-not-lower", Detail
        TEST_CaseRunner.ExpectTrue "ir-add-on-higher", Len(Detail) = 0 And _
                                   IrHigherCount(DemoRows, AbsoluteRows) > 0, _
                                   "no netting set has a higher IR add-on without the bucket formula", _
                                   "illustrative"
        RecordRelation "ir-catches-invalid-set", IrInvalidSetGap(DemoRows, AbsoluteRows, HedgeRows)

'------------------------------------------------------------------------------
' ZERO ADD-ON
'------------------------------------------------------------------------------
    'Two opposite five-year swaps offset exactly, so the add-on is zero.
    'V - C = 60 > 0, so the multiplier is 1 and EAD = 1.4 * 60 = 84.
        TEST_CaseRunner.BeginCase "invariants-zero-add-on", "CRR", VALUATION, "EUR"
        TEST_CaseRunner.AddNettingSet "NS1", "N", "N", "", "N", "N", "", "0", "0", "0", "0", ""
        TEST_CaseRunner.AddTrade "S1", "IR", "", "EUR", "Linear", "Long", "", "Standard", "", _
                                 "10000", "100", "", "2031-09-30", "2031-09-30", "", "", "", "", "", ""
        TEST_CaseRunner.AddTrade "S2", "IR", "", "EUR", "Linear", "Short", "", "Standard", "", _
                                 "10000", "-40", "", "2031-09-30", "2031-09-30", "", "", "", "", "", ""
        TEST_CaseRunner.RunCase
        TEST_CaseRunner.ExpectText "status", "", "netting_set_status", "VALID", "illustrative"
        TEST_CaseRunner.ExpectNumber "aggregate-add-on", "", "aggregate_add_on", "0", "0.000001", "0", _
                                     "illustrative"
        TEST_CaseRunner.ExpectNumber "multiplier", "", "multiplier", "1", "0.000000001", "0", "illustrative"
        TEST_CaseRunner.ExpectNumber "potential-future-exposure", "", "potential_future_exposure", "0", _
                                     "0.000001", "0", "illustrative"
        TEST_CaseRunner.ExpectNumber "exposure-value", "", "exposure_value", "84", "0.000001", _
                                     "0.000000001", "illustrative"

        TEST_CaseRunner.EndSuite ""
        Exit Sub

'------------------------------------------------------------------------------
' HANDLE ERROR
'------------------------------------------------------------------------------
Failed:
        TEST_CaseRunner.EndSuite "unexpected error " & Err.Number & ": " & Err.Description

End Sub


'
'------------------------------------------------------------------------------
'
'                                  RELATIONS
'
'------------------------------------------------------------------------------
'

Private Function NumberGap( _
    ByVal ResultRows As Variant) _
    As String
'
'==============================================================================
'                                  NumberGap
'------------------------------------------------------------------------------
' PURPOSE
'   Check that every figure the other relations read is a number on each
'   VALID row: columns 7 to 19 and 22, and for a margined netting set also
'   20, 21 and 24 to 27. Without this, a blank cell would read as zero.
'
' RETURNS
'   The first missing figure; "" when all are present.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim r     As Long    'Results row
    Dim col   As Long    'Results column

'------------------------------------------------------------------------------
' CHECK
'------------------------------------------------------------------------------
        For r = 1 To RowCount(ResultRows)
            If IsValidRow(ResultRows, r) Then
                For col = COL_V To COL_EAD
                    If col <= COL_ALPHA Or col = COL_EAD Or IsMarginedRow(ResultRows, r) Then
                        If Not IsNum(ResultRows(r, col)) Then
                            NumberGap = RowLabel(ResultRows, r) & " column " & col & " is not a number: " & _
                                        SafeStr(ResultRows(r, col))
                            Exit Function
                        End If
                    End If
                Next col
                If IsMarginedRow(ResultRows, r) Then
                    For col = COL_C_CAP To COL_MULT_CAP
                        If Not IsNum(ResultRows(r, col)) Then
                            NumberGap = RowLabel(ResultRows, r) & " column " & col & " is not a number: " & _
                                        SafeStr(ResultRows(r, col))
                            Exit Function
                        End If
                    Next col
                End If
            End If
        Next r

End Function


Private Function CoverageGap( _
    ByVal ResultRows As Variant) _
    As String
'
'==============================================================================
'                                 CoverageGap
'------------------------------------------------------------------------------
' PURPOSE
'   Check that the demo portfolio exercises every branch the relations
'   test: VALID netting sets that are unmargined and margined, under CRR
'   and BCBS, a margined set where the cap applies and one where it does
'   not, and a multiplier below 1.
'
' RETURNS
'   The first branch with no netting set; "" when all are covered.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim r                       As Long       'Results row
    Dim HasUnmargined           As Boolean    'An unmargined netting set
    Dim HasMargined             As Boolean    'A margined netting set
    Dim HasCrr                  As Boolean    'A CRR netting set
    Dim HasBcbs                 As Boolean    'A BCBS netting set
    Dim HasCapped               As Boolean    'A margined netting set where the cap applies
    Dim HasUncapped             As Boolean    'A margined netting set where it does not
    Dim HasMultiplierBelowOne   As Boolean    'A multiplier below 1

'------------------------------------------------------------------------------
' SCAN
'------------------------------------------------------------------------------
        For r = 1 To RowCount(ResultRows)
            If IsValidRow(ResultRows, r) Then
                If IsMarginedRow(ResultRows, r) Then
                    HasMargined = True
                    If SafeStr(ResultRows(r, COL_CAP_FLAG)) = "Y" Then
                        HasCapped = True
                    ElseIf SafeStr(ResultRows(r, COL_CAP_FLAG)) = "N" Then
                        HasUncapped = True
                    End If
                Else
                    HasUnmargined = True
                End If
                If SafeStr(ResultRows(r, COL_REGIME)) = "CRR" Then
                    HasCrr = True
                ElseIf SafeStr(ResultRows(r, COL_REGIME)) = "BCBS" Then
                    HasBcbs = True
                End If
                If ToDbl(ResultRows(r, COL_MULTIPLIER), 1#) < 1# Then
                    HasMultiplierBelowOne = True
                End If
            End If
        Next r

'------------------------------------------------------------------------------
' REPORT
'------------------------------------------------------------------------------
        If Not HasUnmargined Then
            CoverageGap = "no VALID unmargined netting set"
        ElseIf Not HasMargined Then
            CoverageGap = "no VALID margined netting set"
        ElseIf Not HasCrr Then
            CoverageGap = "no VALID CRR netting set"
        ElseIf Not HasBcbs Then
            CoverageGap = "no VALID BCBS netting set"
        ElseIf Not HasCapped Then
            CoverageGap = "no margined netting set where the cap applies"
        ElseIf Not HasUncapped Then
            CoverageGap = "no margined netting set where the cap does not apply"
        ElseIf Not HasMultiplierBelowOne Then
            CoverageGap = "no multiplier below 1"
        End If

End Function


Private Function AggregateGap( _
    ByVal ResultRows As Variant) _
    As String
'
'==============================================================================
'                                 AggregateGap
'------------------------------------------------------------------------------
' PURPOSE
'   Check that the aggregate add-on is the sum of the six asset-class
'   add-ons and that none of them is negative [CRE52.25].
'
' RETURNS
'   The first violation; "" when none.
'
' UPDATED
'   2026-10-21
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim r          As Long      'Results row
    Dim col        As Long      'Asset-class add-on column
    Dim AddOnSum   As Double    'Sum of the asset-class add-ons

'------------------------------------------------------------------------------
' CHECK
'------------------------------------------------------------------------------
        For r = 1 To RowCount(ResultRows)
            If IsValidRow(ResultRows, r) Then
                AddOnSum = 0#
                For col = COL_IR_ADDON To COL_OT_ADDON
                    If ToDbl(ResultRows(r, col)) < 0# Then
                        AggregateGap = RowLabel(ResultRows, r) & " add-on in column " & col & " is negative"
                        Exit Function
                    End If
                    AddOnSum = AddOnSum + ToDbl(ResultRows(r, col))
                Next col
                If Not Near(ToDbl(ResultRows(r, COL_AGGREGATE)), AddOnSum) Then
                    AggregateGap = RowLabel(ResultRows, r) & " aggregate add-on " & _
                                   ToDbl(ResultRows(r, COL_AGGREGATE)) & ", sum of asset classes " & AddOnSum
                    Exit Function
                End If
            End If
        Next r

End Function


Private Function ReplacementCostGap( _
    ByVal ResultRows As Variant) _
    As String
'
'==============================================================================
'                              ReplacementCostGap
'------------------------------------------------------------------------------
' PURPOSE
'   Check RC against its formula: max(V - C, 0) for an unmargined netting
'   set, max(V - C, TH + MTA - NICA, 0) for a margined one, with TH, MTA
'   and NICA read from the NettingSets inputs. RC is therefore never
'   negative.
'
' RETURNS
'   The first violation; "" when none.
'
' REFERENCE
'   CRE52.10, CRE52.18; CRR Art. 275(1) and (2).
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim r          As Long      'Results row
    Dim Expected   As Double    'RC from the formula
    Dim Actual     As Double    'RC on Results

'------------------------------------------------------------------------------
' CHECK
'------------------------------------------------------------------------------
        For r = 1 To RowCount(ResultRows)
            If IsValidRow(ResultRows, r) Then
                Expected = Max2(ToDbl(ResultRows(r, COL_V)) - ToDbl(ResultRows(r, COL_C)), 0#)
                If IsMarginedRow(ResultRows, r) Then
                    Expected = Max2(Expected, NettingSetInput(RowLabel(ResultRows, r), NS_TH) + _
                                              NettingSetInput(RowLabel(ResultRows, r), NS_MTA) - _
                                              NettingSetInput(RowLabel(ResultRows, r), NS_NICA))
                End If
                Actual = ToDbl(ResultRows(r, COL_RC))
                If Actual < 0# Or Not Near(Actual, Expected) Then
                    ReplacementCostGap = RowLabel(ResultRows, r) & " RC " & Actual & _
                                         ", formula gives " & Expected
                    Exit Function
                End If
            End If
        Next r

End Function


Private Function MultiplierGap( _
    ByVal ResultRows As Variant) _
    As String
'
'==============================================================================
'                                MultiplierGap
'------------------------------------------------------------------------------
' PURPOSE
'   Check the multiplier: between the floor and 1, equal to 1 when
'   V - C >= 0, and equal to its formula when the aggregate add-on is
'   positive. With a zero add-on the formula is undefined and only the
'   bounds are checked; PFE is zero whatever the multiplier. The floor
'   read from Params must itself lie strictly between 0 and 1.
'
' RETURNS
'   The first violation; "" when none.
'
' REFERENCE
'   CRE52.23; CRR Art. 278(3).
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim r             As Long      'Results row
    Dim VLessC        As Double    'V - C
    Dim AddOnAmount   As Double    'Aggregate add-on
    Dim Actual        As Double    'Multiplier on Results

'------------------------------------------------------------------------------
' CHECK
'------------------------------------------------------------------------------
        If mFloor <= 0# Or mFloor >= 1# Then
            MultiplierGap = "MultiplierFloor on Params is not between 0 and 1"
            Exit Function
        End If
        For r = 1 To RowCount(ResultRows)
            If IsValidRow(ResultRows, r) Then
                VLessC = ToDbl(ResultRows(r, COL_V)) - ToDbl(ResultRows(r, COL_C))
                AddOnAmount = ToDbl(ResultRows(r, COL_AGGREGATE))
                Actual = ToDbl(ResultRows(r, COL_MULTIPLIER))
                MultiplierGap = OneMultiplierGap(RowLabel(ResultRows, r), VLessC, AddOnAmount, Actual)
                If Len(MultiplierGap) > 0 Then
                    Exit Function
                End If
            End If
        Next r

End Function


Private Function PfeGap( _
    ByVal ResultRows As Variant) _
    As String
'
'==============================================================================
'                                    PfeGap
'------------------------------------------------------------------------------
' PURPOSE
'   Check that PFE (potential future exposure) is the multiplier times
'   the aggregate add-on [CRE52.20].
'
' RETURNS
'   The first violation; "" when none.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim r          As Long      'Results row
    Dim Expected   As Double    'Multiplier * aggregate add-on

'------------------------------------------------------------------------------
' CHECK
'------------------------------------------------------------------------------
        For r = 1 To RowCount(ResultRows)
            If IsValidRow(ResultRows, r) Then
                Expected = ToDbl(ResultRows(r, COL_MULTIPLIER)) * ToDbl(ResultRows(r, COL_AGGREGATE))
                If Not Near(ToDbl(ResultRows(r, COL_PFE)), Expected) Then
                    PfeGap = RowLabel(ResultRows, r) & " PFE " & ToDbl(ResultRows(r, COL_PFE)) & _
                             ", multiplier * add-on " & Expected
                    Exit Function
                End If
            End If
        Next r

End Function


Private Function ExposureGap( _
    ByVal ResultRows As Variant) _
    As String
'
'==============================================================================
'                                 ExposureGap
'------------------------------------------------------------------------------
' PURPOSE
'   Check EAD = alpha * (RC + PFE): the reported EAD of an unmargined
'   netting set, the margined EAD before the cap of a margined one. Alpha
'   must be positive and the EAD non-negative.
'
' RETURNS
'   The first violation; "" when none.
'
' REFERENCE
'   CRE52.1; CRR Art. 274(2).
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim r            As Long      'Results row
    Dim AlphaValue   As Double    'Alpha of the netting set
    Dim Expected     As Double    'Alpha * (RC + PFE)
    Dim Actual       As Double    'The EAD it must equal
    Dim EadColumn    As Long      'Column of that EAD

'------------------------------------------------------------------------------
' CHECK
'------------------------------------------------------------------------------
        For r = 1 To RowCount(ResultRows)
            If IsValidRow(ResultRows, r) Then
                AlphaValue = ToDbl(ResultRows(r, COL_ALPHA))
                If IsMarginedRow(ResultRows, r) Then
                    EadColumn = COL_EAD_MARGIN
                Else
                    EadColumn = COL_EAD
                End If
                Expected = AlphaValue * (ToDbl(ResultRows(r, COL_RC)) + ToDbl(ResultRows(r, COL_PFE)))
                Actual = ToDbl(ResultRows(r, EadColumn))
                If AlphaValue <= 0# Then
                    ExposureGap = RowLabel(ResultRows, r) & " alpha " & AlphaValue & " is not positive"
                    Exit Function
                ElseIf Actual < 0# Or ToDbl(ResultRows(r, COL_EAD)) < 0# Then
                    ExposureGap = RowLabel(ResultRows, r) & " EAD is negative"
                    Exit Function
                ElseIf Not Near(Actual, Expected) Then
                    ExposureGap = RowLabel(ResultRows, r) & " EAD in column " & EadColumn & " is " & Actual & _
                                  ", alpha * (RC + PFE) " & Expected
                    Exit Function
                End If
            End If
        Next r

End Function


Private Function CapGap( _
    ByVal ResultRows As Variant) _
    As String
'
'==============================================================================
'                                    CapGap
'------------------------------------------------------------------------------
' PURPOSE
'   Check the cap of a margined netting set: its unmargined-basis figures
'   follow the RC and multiplier formulas with C replaced by the cap-basis
'   collateral, which is NICA in both regimes; the cap is
'   alpha * (RC + multiplier * unmargined add-on); the reported EAD is the
'   lower of the margined EAD and the cap; the flag is Y exactly when the
'   cap is lower. An unmargined netting set shows n/a and no cap.
'
' RETURNS
'   The first violation; "" when none.
'
' REFERENCE
'   CRE52.2, CRE52.10; CRR Art. 274(2) and (3); EBA-QA-2023_6962.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim r                    As Long      'Results row
    Dim SetId                As String    'Netting-set ID
    Dim CapCollateral        As Double    'Expected cap-basis collateral
    Dim CapReplacementCost   As Double    'Expected cap-basis RC
    Dim CapAddOn             As Double    'Unmargined aggregate add-on
    Dim CapMultiplier        As Double    'Cap-basis multiplier on Results
    Dim ExpectedCap          As Double    'Expected cap
    Dim MarginedEad          As Double    'Margined EAD before the cap
    Dim ReportedCap          As Double    'Cap on Results
    Dim CapFlag              As String    'Cap-applied flag on Results

'------------------------------------------------------------------------------
' CHECK
'------------------------------------------------------------------------------
        For r = 1 To RowCount(ResultRows)
            If IsValidRow(ResultRows, r) Then
                SetId = RowLabel(ResultRows, r)
                CapFlag = SafeStr(ResultRows(r, COL_CAP_FLAG))
                If Not IsMarginedRow(ResultRows, r) Then
                    If CapFlag <> "n/a" Or Not IsEmpty(ResultRows(r, COL_EAD_CAP)) Then
                        CapGap = SetId & " is unmargined but shows a cap"
                        Exit Function
                    End If
                Else
                    CapCollateral = NettingSetInput(SetId, NS_NICA)
                    CapReplacementCost = Max2(ToDbl(ResultRows(r, COL_V)) - CapCollateral, 0#)
                    CapAddOn = ToDbl(ResultRows(r, COL_ADDON_CAP))
                    CapMultiplier = ToDbl(ResultRows(r, COL_MULT_CAP))
                    ExpectedCap = ToDbl(ResultRows(r, COL_ALPHA)) * (CapReplacementCost + CapMultiplier * CapAddOn)
                    MarginedEad = ToDbl(ResultRows(r, COL_EAD_MARGIN))
                    ReportedCap = ToDbl(ResultRows(r, COL_EAD_CAP))
                    If Not Near(ToDbl(ResultRows(r, COL_C_CAP)), CapCollateral) Then
                        CapGap = SetId & " cap-basis C " & ToDbl(ResultRows(r, COL_C_CAP)) & ", expected " & CapCollateral
                    ElseIf Not Near(ToDbl(ResultRows(r, COL_RC_CAP)), CapReplacementCost) Then
                        CapGap = SetId & " cap-basis RC " & ToDbl(ResultRows(r, COL_RC_CAP)) & _
                                 ", expected " & CapReplacementCost
                    ElseIf CapAddOn < 0# Then
                        CapGap = SetId & " unmargined add-on is negative"
                    ElseIf Not Near(ReportedCap, ExpectedCap) Then
                        CapGap = SetId & " cap " & ReportedCap & ", alpha * (RC + PFE) on the cap basis " & ExpectedCap
                    ElseIf Not Near(ToDbl(ResultRows(r, COL_EAD)), Min2(MarginedEad, ReportedCap)) Then
                        CapGap = SetId & " EAD " & ToDbl(ResultRows(r, COL_EAD)) & ", lower of margined " & _
                                 MarginedEad & " and cap " & ReportedCap & " is " & Min2(MarginedEad, ReportedCap)
                    ElseIf CapFlag <> IIf(ReportedCap < MarginedEad, "Y", "N") Then
                        CapGap = SetId & " cap flag " & CapFlag & " with margined EAD " & MarginedEad & _
                                 " and cap " & ReportedCap
                    Else
                        CapGap = OneMultiplierGap(SetId & " (cap basis)", ToDbl(ResultRows(r, COL_V)) - CapCollateral, _
                                                  CapAddOn, CapMultiplier)
                    End If
                    If Len(CapGap) > 0 Then
                        Exit Function
                    End If
                End If
            End If
        Next r

End Function


Private Function WithheldGap( _
    ByVal ResultRows As Variant) _
    As String
'
'==============================================================================
'                                 WithheldGap
'------------------------------------------------------------------------------
' PURPOSE
'   Check that a netting set that is not VALID (INVALID or INCOMPLETE,
'   whose EAD is withheld, or NO TRADES, which has none) shows no figures
'   from MPOR to the cap-basis multiplier, so that no partial result can be
'   read as an EAD.
'
' RETURNS
'   The first violation; "" when none.
'
' UPDATED
'   2026-10-21
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim r     As Long    'Results row
    Dim col   As Long    'Results column

'------------------------------------------------------------------------------
' CHECK
'------------------------------------------------------------------------------
        For r = 1 To RowCount(ResultRows)
            If Not IsValidRow(ResultRows, r) And Not IsTotalRow(ResultRows, r) Then
                For col = 6 To COL_MULT_CAP
                    If Not IsEmpty(ResultRows(r, col)) Then
                        WithheldGap = RowLabel(ResultRows, r) & " (" & _
                                      SafeStr(ResultRows(r, RS_STATUS_COL)) & ") shows " & _
                                      SafeStr(ResultRows(r, col)) & " in column " & col
                        Exit Function
                    End If
                Next col
            End If
        Next r

End Function


Private Function TotalGap( _
    ByVal ResultRows As Variant) _
    As String
'
'==============================================================================
'                                   TotalGap
'------------------------------------------------------------------------------
' PURPOSE
'   Check the TOTAL rows (#198). They follow the netting-set rows: one per
'   regime that has netting sets, CRR first, labelled "TOTAL CRR" and
'   "TOTAL BCBS", with the regime in the regime column. In each, every
'   column the engine totals (V, C, RC, the six asset-class add-ons, the
'   aggregate add-on, PFE and EAD) holds the sum over the regime's VALID
'   netting sets; the status column says how many netting sets of the
'   regime are VALID and how many were withheld (INVALID or INCOMPLETE);
'   every other column is blank. Withheld netting sets show no figures
'   (WithheldGap), so they add nothing. There is no total across regimes.
'
' RETURNS
'   The first violation, naming the column as "(column n)"; "" when none.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim r                As Long       'Results row
    Dim col              As Long       'Results column
    Dim LastRow          As Long       'Last Results row
    Dim FirstTotalRow    As Long       'First TOTAL row; LastRow + 1 when none
    Dim Regimes          As Variant    'The regimes, CRR first
    Dim g                As Long       'Index into Regimes
    Dim RegimeCode       As String     'Regime of the TOTAL row being checked
    Dim TotalRow         As Long       'Results row of that TOTAL row
    Dim SetCount         As Long       'Netting sets of the regime
    Dim ValidCount       As Long       'Its VALID netting sets
    Dim WithheldCount    As Long       'Its INVALID or INCOMPLETE netting sets
    Dim Status           As String     'Status of a netting-set row
    Dim ExpectedStatus   As String     'Expected status text of the TOTAL row
    Dim Sums(1 To RS_NCOLS) As Double    'Sum of each totalled column over the regime's VALID rows

'------------------------------------------------------------------------------
' LOCATE THE TOTAL ROWS
'------------------------------------------------------------------------------
        LastRow = RowCount(ResultRows)
        If LastRow = 0 Then
            TotalGap = "Results is empty"
            Exit Function
        End If
        FirstTotalRow = LastRow + 1
        For r = 1 To LastRow
            If IsTotalRow(ResultRows, r) Then
                FirstTotalRow = r
                Exit For
            End If
        Next r
        For r = FirstTotalRow To LastRow
            If Not IsTotalRow(ResultRows, r) Then
                TotalGap = "netting set " & RowLabel(ResultRows, r) & " follows a TOTAL row"
                Exit Function
            End If
        Next r

'------------------------------------------------------------------------------
' CHECK EACH REGIME
'------------------------------------------------------------------------------
        Regimes = Array(RG_CRR, RG_BCBS)
        TotalRow = FirstTotalRow
        For g = LBound(Regimes) To UBound(Regimes)
            RegimeCode = CStr(Regimes(g))
            SetCount = 0
            ValidCount = 0
            WithheldCount = 0
            Erase Sums
            For r = 1 To FirstTotalRow - 1
                If SafeStr(ResultRows(r, COL_REGIME)) = RegimeCode Then
                    SetCount = SetCount + 1
                    Status = SafeStr(ResultRows(r, RS_STATUS_COL))
                    If Status = "VALID" Then
                        ValidCount = ValidCount + 1
                        For col = COL_V To COL_EAD
                            If Len(TotalColumnName(col)) > 0 Then
                                Sums(col) = Sums(col) + ToDbl(ResultRows(r, col))
                            End If
                        Next col
                    ElseIf Left$(Status, 7) = "INVALID" Or Left$(Status, 10) = "INCOMPLETE" Then
                        WithheldCount = WithheldCount + 1
                    End If
                End If
            Next r
            If SetCount > 0 Then
                If TotalRow > LastRow Then
                    TotalGap = "no TOTAL " & RegimeCode & " row for " & SetCount & " " & RegimeCode & " netting set(s)"
                    Exit Function
                End If
                If RowLabel(ResultRows, TotalRow) <> "TOTAL " & RegimeCode Then
                    TotalGap = "expected TOTAL " & RegimeCode & ", found " & RowLabel(ResultRows, TotalRow)
                    Exit Function
                End If
                ExpectedStatus = ValidCount & " of " & SetCount & " netting set(s) VALID; " & WithheldCount & " withheld"
                For col = 2 To RS_NCOLS
                    If Len(TotalColumnName(col)) > 0 Then
                        If Not IsNum(ResultRows(TotalRow, col)) Then
                            TotalGap = "TOTAL " & RegimeCode & " " & TotalColumnName(col) & " (column " & col & _
                                       ") is not a number: " & SafeStr(ResultRows(TotalRow, col))
                            Exit Function
                        End If
                        If Not Near(ToDbl(ResultRows(TotalRow, col)), Sums(col)) Then
                            TotalGap = "TOTAL " & RegimeCode & " " & TotalColumnName(col) & " (column " & col & ") " & _
                                       ToDbl(ResultRows(TotalRow, col)) & ", sum of VALID " & RegimeCode & " rows " & Sums(col)
                            Exit Function
                        End If
                    ElseIf col = COL_REGIME Then
                        If SafeStr(ResultRows(TotalRow, col)) <> RegimeCode Then
                            TotalGap = "TOTAL " & RegimeCode & " regime (column " & col & ") holds " & _
                                       SafeStr(ResultRows(TotalRow, col))
                            Exit Function
                        End If
                    ElseIf col = RS_STATUS_COL Then
                        If SafeStr(ResultRows(TotalRow, col)) <> ExpectedStatus Then
                            TotalGap = "TOTAL " & RegimeCode & " status (column " & col & ") is """ & _
                                       SafeStr(ResultRows(TotalRow, col)) & """, expected """ & ExpectedStatus & """"
                            Exit Function
                        End If
                    ElseIf Not IsEmpty(ResultRows(TotalRow, col)) Then
                        TotalGap = "TOTAL " & RegimeCode & " (column " & col & ") holds " & SafeStr(ResultRows(TotalRow, col)) & _
                                   " where it should be blank"
                        Exit Function
                    End If
                Next col
                TotalRow = TotalRow + 1
            End If
        Next g
        If TotalRow <= LastRow Then
            TotalGap = "unexpected row " & RowLabel(ResultRows, TotalRow) & " after the TOTAL rows"
        End If

End Function


Private Function IsTotalRow( _
    ByVal ResultRows As Variant, _
    ByVal RowIndex As Long) _
    As Boolean
'
'==============================================================================
'                                  IsTotalRow
'------------------------------------------------------------------------------
' PURPOSE
'   Tell whether a Results row is a TOTAL row: its label is TOTAL CRR or
'   TOTAL BCBS, or TOTAL as written before #198. The engine rejects these
'   as netting-set IDs, so a netting set such as TOTALENERGIES-1 is not
'   taken for a total.
'
' UPDATED
'   2026-10-17
'==============================================================================
'
    Dim Label   As String    'Label of the row

        Label = RowLabel(ResultRows, RowIndex)
        IsTotalRow = (Label = "TOTAL" Or Label = "TOTAL " & RG_CRR Or Label = "TOTAL " & RG_BCBS)

End Function


Private Function TotalCorruptionGap( _
    ByVal ResultRows As Variant) _
    As String
'
'==============================================================================
'                              TotalCorruptionGap
'------------------------------------------------------------------------------
' PURPOSE
'   Show that TotalGap catches a wrong TOTAL row (#104). On a copy of the
'   Results rows, each totalled column of the last TOTAL row is in turn
'   moved away from its sum and then blanked, and each other column,
'   including its regime and status, is filled with 1. TotalGap must fail
'   every time and name the corrupted column.
'
' RETURNS
'   The first corruption TotalGap missed or attributed to another column;
'   "" when it caught them all.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim col           As Long       'Results column being corrupted
    Dim TotalRow      As Long       'The TOTAL row
    Dim ShownFigure   As Double     'TOTAL figure before the corruption
    Dim pass          As Long       '1 moves the figure, 2 blanks it
    Dim Mutated       As Variant    'Copy of the rows with one cell corrupted
    Dim Detail        As String     'TotalGap on the corrupted copy

'------------------------------------------------------------------------------
' CORRUPT EACH COLUMN
'------------------------------------------------------------------------------
        Detail = TotalGap(ResultRows)
        If Len(Detail) > 0 Then
            TotalCorruptionGap = "TOTAL fails before any corruption: " & Detail
            Exit Function
        End If
        TotalRow = RowCount(ResultRows)
        For col = 2 To RS_NCOLS
            For pass = 1 To 2
                Mutated = ResultRows
                If Len(TotalColumnName(col)) = 0 Then
                    Mutated(TotalRow, col) = 1#
                ElseIf pass = 1 Then
                    ShownFigure = ToDbl(ResultRows(TotalRow, col))
                    Mutated(TotalRow, col) = ShownFigure + 1# + Abs(ShownFigure) * 0.001
                Else
                    Mutated(TotalRow, col) = Empty
                End If
                Detail = TotalGap(Mutated)
                If InStr(Detail, "(column " & col & ")") = 0 Then
                    TotalCorruptionGap = "TOTAL (column " & col & ") " & _
                                         IIf(IsEmpty(Mutated(TotalRow, col)), "blanked", _
                                             "set to " & SafeStr(Mutated(TotalRow, col))) & ": " & _
                                         IIf(Len(Detail) = 0, "not caught", "reported as " & Detail)
                    Exit Function
                End If
                If Len(TotalColumnName(col)) = 0 Then
                    Exit For
                End If
            Next pass
        Next col

End Function


Private Function IdenticalGap( _
    ByVal FirstRows As Variant, _
    ByVal SecondRows As Variant, _
    ByVal SheetName As String) _
    As String
'
'==============================================================================
'                                 IdenticalGap
'------------------------------------------------------------------------------
' PURPOSE
'   Check that two runs on the same inputs wrote exactly the same cells:
'   same row count, same values, numbers equal to the last bit.
'
' INPUTS
'   FirstRows, SecondRows: the sheet's output block from each run.
'   SheetName: for the report.
'
' RETURNS
'   The first difference; "" when none.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim r            As Long       'Row
    Dim col          As Long       'Column
    Dim FirstCell    As Variant    'First run's cell
    Dim SecondCell   As Variant    'Second run's cell
    Dim IsSame       As Boolean    'The two cells are equal

'------------------------------------------------------------------------------
' COMPARE
'------------------------------------------------------------------------------
        If RowCount(FirstRows) = 0 Then
            IdenticalGap = SheetName & " is empty"
            Exit Function
        End If
        If RowCount(SecondRows) <> RowCount(FirstRows) Then
            IdenticalGap = SheetName & " has " & RowCount(SecondRows) & " rows, first run " & _
                           RowCount(FirstRows)
            Exit Function
        End If
        For r = 1 To RowCount(FirstRows)
            For col = LBound(FirstRows, 2) To UBound(FirstRows, 2)
                FirstCell = FirstRows(r, col)
                SecondCell = SecondRows(r, col)
                If IsError(FirstCell) Or IsError(SecondCell) Then
                    IsSame = IsError(FirstCell) And IsError(SecondCell)
                ElseIf VarType(FirstCell) <> VarType(SecondCell) Then
                    IsSame = False
                Else
                    IsSame = (FirstCell = SecondCell)
                End If
                If Not IsSame Then
                    IdenticalGap = SheetName & " row " & r & " column " & col & ": " & SafeStr(FirstCell) & _
                                   " became " & SafeStr(SecondCell)
                    Exit Function
                End If
            Next col
        Next r

End Function


Private Function IrOffsetGap( _
    ByVal BucketFormulaRows As Variant, _
    ByVal SumOfAbsoluteRows As Variant, _
    ByVal HedgeRows As Variant) _
    As String
'
'==============================================================================
'                                 IrOffsetGap
'------------------------------------------------------------------------------
' PURPOSE
'   For every netting set VALID with the bucket formula, check the run
'   with the sum of absolute bucket values (#103). A CRR netting set with
'   an IR hedging set must be INVALID there, since the CRR allows only the
'   formula [Art. 280a(3)]. Every other set must stay VALID with an IR
'   add-on not lower. A set that is missing, or withheld when it should
'   not be, fails, so another set's higher add-on cannot hide it.
'
' INPUTS
'   BucketFormulaRows: Results rows with the bucket formula.
'   SumOfAbsoluteRows: Results rows with the sum of absolute values.
'   HedgeRows: HedgingSets rows with the bucket formula.
'
' RETURNS
'   The first violation; "" when none.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim r               As Long      'Row of BucketFormulaRows
    Dim MatchRow        As Long      'Matching row of SumOfAbsoluteRows
    Dim WithOffset      As Double    'IR add-on with the bucket formula
    Dim WithoutOffset   As Double    'IR add-on with the sum of absolutes
    Dim SetId           As String    'Netting set
    Dim SecondStatus    As String    'Its status in the second run

'------------------------------------------------------------------------------
' CHECK
'------------------------------------------------------------------------------
        For r = 1 To RowCount(BucketFormulaRows)
            If IsValidRow(BucketFormulaRows, r) Then
                SetId = RowLabel(BucketFormulaRows, r)
                MatchRow = MatchingRow(SumOfAbsoluteRows, SetId)
                If MatchRow = 0 Then
                    IrOffsetGap = SetId & " missing from the second run"
                    Exit Function
                End If
                SecondStatus = SafeStr(SumOfAbsoluteRows(MatchRow, RS_STATUS_COL))
                If SafeStr(BucketFormulaRows(r, COL_REGIME)) = "CRR" And HasIrHedgingSet(HedgeRows, SetId) Then
                    If Left$(SecondStatus, 7) <> "INVALID" Then
                        IrOffsetGap = SetId & " (CRR, interest-rate trades) is " & SecondStatus & _
                                      " with the sum of absolutes, which Art. 280a(3) does not allow"
                        Exit Function
                    End If
                ElseIf Not IsValidRow(SumOfAbsoluteRows, MatchRow) Then
                    IrOffsetGap = SetId & " is " & SecondStatus & " in the second run, VALID in the first"
                    Exit Function
                Else
                    WithOffset = ToDbl(BucketFormulaRows(r, COL_IR_ADDON))
                    WithoutOffset = ToDbl(SumOfAbsoluteRows(MatchRow, COL_IR_ADDON))
                    If WithoutOffset < WithOffset And Not Near(WithoutOffset, WithOffset) Then
                        IrOffsetGap = SetId & " IR add-on " & WithoutOffset & " without the bucket formula, " & _
                                      WithOffset & " with it"
                        Exit Function
                    End If
                End If
            End If
        Next r

End Function


Private Function IrHigherCount( _
    ByVal BucketFormulaRows As Variant, _
    ByVal SumOfAbsoluteRows As Variant) _
    As Long
'
'==============================================================================
'                                IrHigherCount
'------------------------------------------------------------------------------
' PURPOSE
'   Count the netting sets whose IR add-on is strictly higher with the sum
'   of absolute bucket values, to show that the parameter took effect.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim r               As Long      'Row of BucketFormulaRows
    Dim MatchRow        As Long      'Matching row of SumOfAbsoluteRows
    Dim WithOffset      As Double    'IR add-on with the bucket formula
    Dim WithoutOffset   As Double    'IR add-on with the sum of absolutes

'------------------------------------------------------------------------------
' COUNT
'------------------------------------------------------------------------------
        For r = 1 To RowCount(BucketFormulaRows)
            If IsValidRow(BucketFormulaRows, r) Then
                MatchRow = MatchingRow(SumOfAbsoluteRows, RowLabel(BucketFormulaRows, r))
                If MatchRow > 0 Then
                    WithOffset = ToDbl(BucketFormulaRows(r, COL_IR_ADDON))
                    WithoutOffset = ToDbl(SumOfAbsoluteRows(MatchRow, COL_IR_ADDON))
                    If WithoutOffset > WithOffset And Not Near(WithoutOffset, WithOffset) Then
                        IrHigherCount = IrHigherCount + 1
                    End If
                End If
            End If
        Next r

End Function


Private Function IrInvalidSetGap( _
    ByVal BucketFormulaRows As Variant, _
    ByVal SumOfAbsoluteRows As Variant, _
    ByVal HedgeRows As Variant) _
    As String
'
'==============================================================================
'                               IrInvalidSetGap
'------------------------------------------------------------------------------
' PURPOSE
'   Show that IrOffsetGap catches a netting set that becomes invalid
'   (#103). On a copy of the second run, the first netting set VALID in
'   both runs whose withdrawal still leaves another set with a higher IR
'   add-on is marked INVALID, with its figures blanked as the engine
'   withholds them. IrOffsetGap must then fail and name that netting set.
'
' RETURNS
'   The violation; "" when IrOffsetGap caught it.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim r          As Long       'Row of BucketFormulaRows
    Dim MatchRow   As Long       'Matching row of the second run
    Dim col        As Long       'Results column blanked
    Dim Mutated    As Variant    'Copy of the second run with one set withheld
    Dim SetId      As String     'Netting set made INVALID
    Dim Detail     As String     'IrOffsetGap on the corrupted copy

'------------------------------------------------------------------------------
' WITHHOLD ONE NETTING SET
'------------------------------------------------------------------------------
        For r = 1 To RowCount(BucketFormulaRows)
            If IsValidRow(BucketFormulaRows, r) Then
                SetId = RowLabel(BucketFormulaRows, r)
                MatchRow = MatchingRow(SumOfAbsoluteRows, SetId)
                If MatchRow > 0 Then
                    If Not IsValidRow(SumOfAbsoluteRows, MatchRow) Then
                        MatchRow = 0
                    End If
                End If
                If MatchRow > 0 Then
                    Mutated = SumOfAbsoluteRows
                    For col = 6 To COL_MULT_CAP
                        Mutated(MatchRow, col) = Empty
                    Next col
                    Mutated(MatchRow, RS_STATUS_COL) = "INVALID"
                    If IrHigherCount(BucketFormulaRows, Mutated) > 0 Then
                        Detail = IrOffsetGap(BucketFormulaRows, Mutated, HedgeRows)
                        If InStr(Detail, SetId & " is INVALID") = 0 Then
                            IrInvalidSetGap = SetId & " made INVALID in the second run: " & _
                                              IIf(Len(Detail) = 0, "not caught", "reported as " & Detail)
                        End If
                        Exit Function
                    End If
                End If
            End If
        Next r
        IrInvalidSetGap = "no VALID netting set can be withheld while another keeps a higher IR add-on"

End Function


'
'------------------------------------------------------------------------------
'
'                                   HELPERS
'
'------------------------------------------------------------------------------
'

Private Sub RecordRelation( _
    ByVal Label As String, _
    ByVal Detail As String)
'
'==============================================================================
'                                RecordRelation
'------------------------------------------------------------------------------
' PURPOSE
'   Record a relation check: it passes when no violation was found.
'
' UPDATED
'   2026-10-17
'==============================================================================
'
        TEST_CaseRunner.ExpectTrue Label, Len(Detail) = 0, Detail, "illustrative"

End Sub


Private Function OneMultiplierGap( _
    ByVal Label As String, _
    ByVal VLessC As Double, _
    ByVal AddOnAmount As Double, _
    ByVal Actual As Double) _
    As String
'
'==============================================================================
'                               OneMultiplierGap
'------------------------------------------------------------------------------
' PURPOSE
'   Check one multiplier against floor <= m <= 1, m = 1 when V - C >= 0,
'   and m = min(1, floor + (1 - floor) * exp((V - C) / (2 * (1 - floor) *
'   add-on))) when the add-on is positive.
'
' INPUTS
'   Label: netting set, for the report.
'   VLessC: V - C in the calculation currency.
'   AddOnAmount: aggregate add-on, >= 0.
'   Actual: the multiplier on Results.
'
' RETURNS
'   The violation; "" when none.
'
' REFERENCE
'   CRE52.23; CRR Art. 278(3).
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Exponent   As Double    '(V - C) / (2 * (1 - floor) * add-on)
    Dim Expected   As Double    'Multiplier from the formula

'------------------------------------------------------------------------------
' CHECK
'------------------------------------------------------------------------------
        If Actual < mFloor - REL_TOL Or Actual > 1# + REL_TOL Then
            OneMultiplierGap = Label & " multiplier " & Actual & " outside [" & mFloor & ", 1]"
        ElseIf VLessC >= 0# And Not Near(Actual, 1#) Then
            OneMultiplierGap = Label & " multiplier " & Actual & " with V - C = " & VLessC & " >= 0"
        ElseIf AddOnAmount > 0# Then
            Exponent = VLessC / (2# * (1# - mFloor) * AddOnAmount)
            If Exponent >= 0# Then
                Expected = 1#
            ElseIf Exponent < -700# Then
                Expected = mFloor
            Else
                Expected = Min2(1#, mFloor + (1# - mFloor) * Exp(Exponent))
            End If
            If Not Near(Actual, Expected) Then
                OneMultiplierGap = Label & " multiplier " & Actual & ", formula gives " & Expected
            End If
        End If

End Function


Private Function NettingSetInput( _
    ByVal SetId As String, _
    ByVal InputColumn As Long) _
    As Double
'
'==============================================================================
'                               NettingSetInput
'------------------------------------------------------------------------------
' PURPOSE
'   Read one amount from the netting set's row on NettingSets; a blank
'   cell is 0, as the engine reads it.
'
' ERROR POLICY
'   Raises ERR_TEST_SETUP when the netting set has no input row; the suite
'   then ends as failed.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim NettingSetsSheet   As Worksheet    'NettingSets sheet
    Dim InputRow           As Long         'Input row; 0 when absent

'------------------------------------------------------------------------------
' READ
'------------------------------------------------------------------------------
        Set NettingSetsSheet = GetSheet(SH_NS)
        InputRow = FindHeaderRow(NettingSetsSheet, NS_ID, SetId)
        If InputRow < TableFirstRow(SH_NS) Then
            Err.Raise ERR_TEST_SETUP, "TEST_Invariants.NettingSetInput", _
                      SetId & " not found on " & SH_NS & "."
        End If
        NettingSetInput = ToDbl(TableCell(NettingSetsSheet, InputRow, InputColumn).Value)

End Function


Private Function ParamsRow( _
    ByVal ParamCode As String) _
    As Long
'
'==============================================================================
'                                  ParamsRow
'------------------------------------------------------------------------------
' PURPOSE
'   Find the Params row whose code column holds a parameter code.
'
' ERROR POLICY
'   Raises ERR_TEST_SETUP when the code is absent; the suite then ends as
'   failed.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim FoundRow   As Long    'Row found; 0 when absent

'------------------------------------------------------------------------------
' FIND
'------------------------------------------------------------------------------
        FoundRow = FindHeaderRow(GetSheet(SH_PARAMS), ParamsCodeCol(), ParamCode)
        If FoundRow = 0 Then
            Err.Raise ERR_TEST_SETUP, "TEST_Invariants.ParamsRow", _
                      "'" & ParamCode & "' not found in column C of Params."
        End If
        ParamsRow = FoundRow

End Function


Private Function SheetRows( _
    ByVal SheetName As String, _
    ByVal ColumnCount As Long) _
    As Variant
'
'==============================================================================
'                                  SheetRows
'------------------------------------------------------------------------------
' PURPOSE
'   Read an output sheet's data block, from the first data row to the last
'   used row, as a (row, column) array.
'
' RETURNS
'   The values; Empty when the sheet has no data rows.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim OutputSheet   As Worksheet    'Output sheet
    Dim LastUsedRow   As Long         'Last used row

'------------------------------------------------------------------------------
' READ
'------------------------------------------------------------------------------
        Set OutputSheet = GetSheet(SheetName)
        LastUsedRow = UsedLastRow(OutputSheet)
        If LastUsedRow >= OutputFirstRow(SheetName) Then
            SheetRows = TableRange(OutputSheet, OutputFirstRow(SheetName), 1, LastUsedRow, ColumnCount).Value
        End If

End Function


Private Function RowCount( _
    ByVal ResultRows As Variant) _
    As Long
'
'==============================================================================
'                                   RowCount
'------------------------------------------------------------------------------
' PURPOSE
'   Number of rows in a block read by SheetRows; 0 when it is empty.
'
' UPDATED
'   2026-10-17
'==============================================================================
'
        If IsArray(ResultRows) Then
            RowCount = UBound(ResultRows, 1)
        End If

End Function


Private Function RowLabel( _
    ByVal ResultRows As Variant, _
    ByVal RowIndex As Long) _
    As String
'
'==============================================================================
'                                    RowLabel
'------------------------------------------------------------------------------
' PURPOSE
'   The ID in the first column of a Results row.
'
' UPDATED
'   2026-10-17
'==============================================================================
'
        RowLabel = SafeStr(ResultRows(RowIndex, 1))

End Function


Private Function IsValidRow( _
    ByVal ResultRows As Variant, _
    ByVal RowIndex As Long) _
    As Boolean
'
'==============================================================================
'                                  IsValidRow
'------------------------------------------------------------------------------
' PURPOSE
'   Tell whether a Results row is a netting set with an EAD (status VALID).
'
' UPDATED
'   2026-10-17
'==============================================================================
'
        IsValidRow = (SafeStr(ResultRows(RowIndex, RS_STATUS_COL)) = "VALID")

End Function


Private Function IsMarginedRow( _
    ByVal ResultRows As Variant, _
    ByVal RowIndex As Long) _
    As Boolean
'
'==============================================================================
'                                IsMarginedRow
'------------------------------------------------------------------------------
' PURPOSE
'   Tell whether a Results row is a margined netting set.
'
' UPDATED
'   2026-10-17
'==============================================================================
'
        IsMarginedRow = (SafeStr(ResultRows(RowIndex, COL_MARGINED)) = "Y")

End Function


Private Function MatchingRow( _
    ByVal ResultRows As Variant, _
    ByVal RowKey As String) _
    As Long
'
'==============================================================================
'                                 MatchingRow
'------------------------------------------------------------------------------
' PURPOSE
'   Find the row of a block whose first column holds RowKey.
'
' RETURNS
'   The row; 0 when absent.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim r   As Long    'Row being compared

'------------------------------------------------------------------------------
' FIND
'------------------------------------------------------------------------------
        For r = 1 To RowCount(ResultRows)
            If RowLabel(ResultRows, r) = RowKey Then
                MatchingRow = r
                Exit Function
            End If
        Next r

End Function


Private Function HasIrHedgingSet( _
    ByVal HedgeRows As Variant, _
    ByVal SetId As String) _
    As Boolean
'
'==============================================================================
'                               HasIrHedgingSet
'------------------------------------------------------------------------------
' PURPOSE
'   Tell whether a netting set has an interest-rate hedging set on
'   HedgingSets (column 1 netting set, column 2 asset class).
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim r   As Long    'HedgingSets row

'------------------------------------------------------------------------------
' FIND
'------------------------------------------------------------------------------
        For r = 1 To RowCount(HedgeRows)
            If SafeStr(HedgeRows(r, 1)) = SetId And SafeStr(HedgeRows(r, 2)) = ACCode(AC_IR) Then
                HasIrHedgingSet = True
                Exit Function
            End If
        Next r

End Function


Private Function TotalColumnName( _
    ByVal ResultsColumn As Long) _
    As String
'
'==============================================================================
'                               TotalColumnName
'------------------------------------------------------------------------------
' PURPOSE
'   Name a Results column the engine sums into the TOTAL row: V, C, RC,
'   the six asset-class add-ons, the aggregate add-on, PFE and EAD, as
'   CORE_Engine.ComputeNettingSets writes them.
'
' RETURNS
'   The name; "" for a column the TOTAL row leaves blank.
'
' UPDATED
'   2026-10-17
'==============================================================================
'
        Select Case ResultsColumn
            Case COL_V:           TotalColumnName = "V"
            Case COL_C:           TotalColumnName = "C"
            Case COL_RC:          TotalColumnName = "RC"
            Case COL_IR_ADDON:    TotalColumnName = "IR add-on"
            Case 11:              TotalColumnName = "FX add-on"
            Case 12:              TotalColumnName = "CR add-on"
            Case 13:              TotalColumnName = "EQ add-on"
            Case 14:              TotalColumnName = "CO add-on"
            Case COL_OT_ADDON:    TotalColumnName = "OT add-on"
            Case COL_AGGREGATE:   TotalColumnName = "aggregate add-on"
            Case COL_PFE:         TotalColumnName = "PFE"
            Case COL_EAD:         TotalColumnName = "EAD"
        End Select

End Function


Private Function Near( _
    ByVal Actual As Double, _
    ByVal Expected As Double) _
    As Boolean
'
'==============================================================================
'                                     Near
'------------------------------------------------------------------------------
' PURPOSE
'   Compare two figures: |actual - expected| <= REL_TOL * max(1,
'   |expected|), which allows for the order of floating-point sums.
'
' UPDATED
'   2026-10-17
'==============================================================================
'
        Near = Abs(Actual - Expected) <= REL_TOL * Max2(1#, Abs(Expected))

End Function
