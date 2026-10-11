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
'   sets show no figures; the TOTAL row adds up. It also checks that a
'   second run gives identical outputs, that the IR sum of absolute bucket
'   values is never below the bucket formula, and that a netting set with
'   a zero add-on is calculated without a division failure.
'
' PUBLIC SURFACE
'   RunInvariantTests is the entry point. Option Private Module keeps it
'   out of the external workbook automation API.
'
' DEPENDENCIES
'   TEST_CaseRunner, which writes the inputs, runs the engine, records the
'   checks and restores the workbook. The relations are recomputed here
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
'   CRE52.1 and CRR Art. 274(2) (EAD), CRE52.10 to 52.20 and CRR Art. 275
'   (RC), CRE52.23 and CRR Art. 278(3) (multiplier), CRE52.24 (aggregate
'   add-on), CRE52.2 and CRR Art. 274(3) (cap of a margined netting set),
'   CRE52.57 and CRR Art. 280a (IR bucket formula).
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
    Option Explicit
    Option Private Module

'------------------------------------------------------------------------------
' MODULE CONSTANTS
'------------------------------------------------------------------------------
    'Suite size and tolerance.
        Private Const VALUATION   As String = "2026-09-30"    'Valuation date of the built case
        Private Const CASES       As Long = 4                 'Cases in a complete run
        Private Const CHECKS      As Long = 21                'Checks in a complete run
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
'   2026-10-09
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim original     As Variant    'Results rows of the first demo run
    Dim tradeRows    As Variant    'TradeCalc rows of the first demo run
    Dim hedgeRows    As Variant    'HedgingSets rows of the first demo run
    Dim bucketRows   As Variant    'Buckets rows of the first demo run
    Dim detail       As String     'First violation; "" when none

'------------------------------------------------------------------------------
' RELATIONS ON THE DEMO PORTFOLIO
'------------------------------------------------------------------------------
        TEST_CaseRunner.BeginSuite CASES, CHECKS
        On Error GoTo Failed

        TEST_CaseRunner.BeginCase "invariants-demo", "CRR", VALUATION, "EUR"
        TEST_CaseRunner.LoadSavedInputs
        TEST_CaseRunner.RunCase
        mFloor = ToDbl(GetParam(PRM_FLOOR), -1#)
        original = SheetRows(SH_RESULTS, RS_NCOLS)
        tradeRows = SheetRows(SH_TRADECALC, TC_NCOLS)
        hedgeRows = SheetRows(SH_HEDGING, HS_NCOLS)
        bucketRows = SheetRows(SH_BUCKETS, BK_NCOLS)
        RecordRelation "numbers", NumberGap(original)
        RecordRelation "coverage", CoverageGap(original)
        RecordRelation "aggregate-add-on", AggregateGap(original)
        RecordRelation "replacement-cost", ReplacementCostGap(original)
        RecordRelation "multiplier", MultiplierGap(original)
        RecordRelation "potential-future-exposure", PfeGap(original)
        RecordRelation "exposure-value", ExposureGap(original)
        RecordRelation "unmargined-cap", CapGap(original)
        RecordRelation "withheld", WithheldGap(original)
        RecordRelation "total", TotalGap(original)

'------------------------------------------------------------------------------
' REPEATABILITY
'------------------------------------------------------------------------------
        TEST_CaseRunner.BeginCase "invariants-demo-repeat", "CRR", VALUATION, "EUR"
        TEST_CaseRunner.LoadSavedInputs
        TEST_CaseRunner.RunCase
        RecordRelation "same-results", _
                       IdenticalGap(original, SheetRows(SH_RESULTS, RS_NCOLS), SH_RESULTS)
        RecordRelation "same-trade-calc", _
                       IdenticalGap(tradeRows, SheetRows(SH_TRADECALC, TC_NCOLS), SH_TRADECALC)
        RecordRelation "same-hedging-sets", _
                       IdenticalGap(hedgeRows, SheetRows(SH_HEDGING, HS_NCOLS), SH_HEDGING)
        RecordRelation "same-buckets", _
                       IdenticalGap(bucketRows, SheetRows(SH_BUCKETS, BK_NCOLS), SH_BUCKETS)

'------------------------------------------------------------------------------
' IR BUCKET OFFSET
'------------------------------------------------------------------------------
    'The sum of absolute bucket values can never be below the bucket
    'formula, whose cross terms have correlations below 1 [CRE52.57].
        TEST_CaseRunner.BeginCase "invariants-demo-ir-sum-of-absolutes", "CRR", VALUATION, "EUR"
        TEST_CaseRunner.LoadSavedInputs
        TEST_CaseRunner.PatchCell SH_PARAMS, ParamsRow(PRM_IRFULL), PRM_VALUE_COL, False
        TEST_CaseRunner.RunCase
        detail = IrOffsetGap(original, SheetRows(SH_RESULTS, RS_NCOLS))
        RecordRelation "ir-add-on-not-lower", detail
        TEST_CaseRunner.ExpectTrue "ir-add-on-higher", Len(detail) = 0 And _
                                   IrHigherCount(original, SheetRows(SH_RESULTS, RS_NCOLS)) > 0, _
                                   "no netting set has a higher IR add-on without the bucket formula", _
                                   "illustrative"

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
    ByVal outputRows As Variant) _
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
'   2026-10-09
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
        For r = 1 To RowCount(outputRows)
            If IsValidRow(outputRows, r) Then
                For col = COL_V To COL_EAD
                    If col <= COL_ALPHA Or col = COL_EAD Or IsMarginedRow(outputRows, r) Then
                        If Not IsNum(outputRows(r, col)) Then
                            NumberGap = RowLabel(outputRows, r) & " column " & col & " is not a number: " & _
                                        SafeStr(outputRows(r, col))
                            Exit Function
                        End If
                    End If
                Next col
                If IsMarginedRow(outputRows, r) Then
                    For col = COL_C_CAP To COL_MULT_CAP
                        If Not IsNum(outputRows(r, col)) Then
                            NumberGap = RowLabel(outputRows, r) & " column " & col & " is not a number: " & _
                                        SafeStr(outputRows(r, col))
                            Exit Function
                        End If
                    Next col
                End If
            End If
        Next r

End Function


Private Function CoverageGap( _
    ByVal outputRows As Variant) _
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
'   2026-10-09
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim r          As Long       'Results row
    Dim unmarg     As Boolean    'An unmargined netting set
    Dim marg       As Boolean    'A margined netting set
    Dim crr        As Boolean    'A CRR netting set
    Dim bcbs       As Boolean    'A BCBS netting set
    Dim capped     As Boolean    'A margined netting set where the cap applies
    Dim uncapped   As Boolean    'A margined netting set where it does not
    Dim belowOne   As Boolean    'A multiplier below 1

'------------------------------------------------------------------------------
' SCAN
'------------------------------------------------------------------------------
        For r = 1 To RowCount(outputRows)
            If IsValidRow(outputRows, r) Then
                If IsMarginedRow(outputRows, r) Then
                    marg = True
                    If SafeStr(outputRows(r, COL_CAP_FLAG)) = "Y" Then
                        capped = True
                    ElseIf SafeStr(outputRows(r, COL_CAP_FLAG)) = "N" Then
                        uncapped = True
                    End If
                Else
                    unmarg = True
                End If
                If SafeStr(outputRows(r, COL_REGIME)) = "CRR" Then
                    crr = True
                ElseIf SafeStr(outputRows(r, COL_REGIME)) = "BCBS" Then
                    bcbs = True
                End If
                If ToDbl(outputRows(r, COL_MULTIPLIER), 1#) < 1# Then
                    belowOne = True
                End If
            End If
        Next r

'------------------------------------------------------------------------------
' REPORT
'------------------------------------------------------------------------------
        If Not unmarg Then
            CoverageGap = "no VALID unmargined netting set"
        ElseIf Not marg Then
            CoverageGap = "no VALID margined netting set"
        ElseIf Not crr Then
            CoverageGap = "no VALID CRR netting set"
        ElseIf Not bcbs Then
            CoverageGap = "no VALID BCBS netting set"
        ElseIf Not capped Then
            CoverageGap = "no margined netting set where the cap applies"
        ElseIf Not uncapped Then
            CoverageGap = "no margined netting set where the cap does not apply"
        ElseIf Not belowOne Then
            CoverageGap = "no multiplier below 1"
        End If

End Function


Private Function AggregateGap( _
    ByVal outputRows As Variant) _
    As String
'
'==============================================================================
'                                 AggregateGap
'------------------------------------------------------------------------------
' PURPOSE
'   Check that the aggregate add-on is the sum of the six asset-class
'   add-ons and that none of them is negative [CRE52.24].
'
' RETURNS
'   The first violation; "" when none.
'
' UPDATED
'   2026-10-09
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim r       As Long      'Results row
    Dim col     As Long      'Asset-class add-on column
    Dim added   As Double    'Sum of the asset-class add-ons

'------------------------------------------------------------------------------
' CHECK
'------------------------------------------------------------------------------
        For r = 1 To RowCount(outputRows)
            If IsValidRow(outputRows, r) Then
                added = 0#
                For col = COL_IR_ADDON To COL_OT_ADDON
                    If ToDbl(outputRows(r, col)) < 0# Then
                        AggregateGap = RowLabel(outputRows, r) & " add-on in column " & col & " is negative"
                        Exit Function
                    End If
                    added = added + ToDbl(outputRows(r, col))
                Next col
                If Not Near(ToDbl(outputRows(r, COL_AGGREGATE)), added) Then
                    AggregateGap = RowLabel(outputRows, r) & " aggregate add-on " & _
                                   ToDbl(outputRows(r, COL_AGGREGATE)) & ", sum of asset classes " & added
                    Exit Function
                End If
            End If
        Next r

End Function


Private Function ReplacementCostGap( _
    ByVal outputRows As Variant) _
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
'   2026-10-09
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim r          As Long      'Results row
    Dim expected   As Double    'RC from the formula
    Dim actual     As Double    'RC on Results

'------------------------------------------------------------------------------
' CHECK
'------------------------------------------------------------------------------
        For r = 1 To RowCount(outputRows)
            If IsValidRow(outputRows, r) Then
                expected = Max2(ToDbl(outputRows(r, COL_V)) - ToDbl(outputRows(r, COL_C)), 0#)
                If IsMarginedRow(outputRows, r) Then
                    expected = Max2(expected, NettingSetInput(RowLabel(outputRows, r), NS_TH) + _
                                              NettingSetInput(RowLabel(outputRows, r), NS_MTA) - _
                                              NettingSetInput(RowLabel(outputRows, r), NS_NICA))
                End If
                actual = ToDbl(outputRows(r, COL_RC))
                If actual < 0# Or Not Near(actual, expected) Then
                    ReplacementCostGap = RowLabel(outputRows, r) & " RC " & actual & _
                                         ", formula gives " & expected
                    Exit Function
                End If
            End If
        Next r

End Function


Private Function MultiplierGap( _
    ByVal outputRows As Variant) _
    As String
'
'==============================================================================
'                                MultiplierGap
'------------------------------------------------------------------------------
' PURPOSE
'   Check the multiplier: between the floor and 1, equal to 1 when
'   V - C >= 0, and equal to its formula when the aggregate add-on is
'   positive. With a zero add-on the formula is undefined and only the
'   bounds are checked; PFE is zero whatever the multiplier.
'
' RETURNS
'   The first violation; "" when none.
'
' REFERENCE
'   CRE52.23; CRR Art. 278(3).
'
' UPDATED
'   2026-10-09
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim r             As Long      'Results row
    Dim vLessC        As Double    'V - C
    Dim addOnAmount   As Double    'Aggregate add-on
    Dim actual        As Double    'Multiplier on Results

'------------------------------------------------------------------------------
' CHECK
'------------------------------------------------------------------------------
        If mFloor <= 0# Or mFloor >= 1# Then
            MultiplierGap = "MultiplierFloor on Params is not between 0 and 1"
            Exit Function
        End If
        For r = 1 To RowCount(outputRows)
            If IsValidRow(outputRows, r) Then
                vLessC = ToDbl(outputRows(r, COL_V)) - ToDbl(outputRows(r, COL_C))
                addOnAmount = ToDbl(outputRows(r, COL_AGGREGATE))
                actual = ToDbl(outputRows(r, COL_MULTIPLIER))
                MultiplierGap = OneMultiplierGap(RowLabel(outputRows, r), vLessC, addOnAmount, actual)
                If Len(MultiplierGap) > 0 Then
                    Exit Function
                End If
            End If
        Next r

End Function


Private Function PfeGap( _
    ByVal outputRows As Variant) _
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
'   2026-10-09
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim r          As Long      'Results row
    Dim expected   As Double    'Multiplier * aggregate add-on

'------------------------------------------------------------------------------
' CHECK
'------------------------------------------------------------------------------
        For r = 1 To RowCount(outputRows)
            If IsValidRow(outputRows, r) Then
                expected = ToDbl(outputRows(r, COL_MULTIPLIER)) * ToDbl(outputRows(r, COL_AGGREGATE))
                If Not Near(ToDbl(outputRows(r, COL_PFE)), expected) Then
                    PfeGap = RowLabel(outputRows, r) & " PFE " & ToDbl(outputRows(r, COL_PFE)) & _
                             ", multiplier * add-on " & expected
                    Exit Function
                End If
            End If
        Next r

End Function


Private Function ExposureGap( _
    ByVal outputRows As Variant) _
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
'   2026-10-09
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim r          As Long      'Results row
    Dim alphaVal   As Double    'Alpha of the netting set
    Dim expected   As Double    'alpha * (RC + PFE)
    Dim actual     As Double    'The EAD it must equal
    Dim col        As Long      'Column of that EAD

'------------------------------------------------------------------------------
' CHECK
'------------------------------------------------------------------------------
        For r = 1 To RowCount(outputRows)
            If IsValidRow(outputRows, r) Then
                alphaVal = ToDbl(outputRows(r, COL_ALPHA))
                If IsMarginedRow(outputRows, r) Then
                    col = COL_EAD_MARGIN
                Else
                    col = COL_EAD
                End If
                expected = alphaVal * (ToDbl(outputRows(r, COL_RC)) + ToDbl(outputRows(r, COL_PFE)))
                actual = ToDbl(outputRows(r, col))
                If alphaVal <= 0# Then
                    ExposureGap = RowLabel(outputRows, r) & " alpha " & alphaVal & " is not positive"
                    Exit Function
                ElseIf actual < 0# Or ToDbl(outputRows(r, COL_EAD)) < 0# Then
                    ExposureGap = RowLabel(outputRows, r) & " EAD is negative"
                    Exit Function
                ElseIf Not Near(actual, expected) Then
                    ExposureGap = RowLabel(outputRows, r) & " EAD in column " & col & " is " & actual & _
                                  ", alpha * (RC + PFE) " & expected
                    Exit Function
                End If
            End If
        Next r

End Function


Private Function CapGap( _
    ByVal outputRows As Variant) _
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
'   CRE52.2, CRE52.10; CRR Art. 274(2) and (3); EBA Q&A 2023_6962.
'
' UPDATED
'   2026-10-10
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim r          As Long      'Results row
    Dim id         As String    'Netting-set ID
    Dim cCap       As Double    'Expected cap-basis collateral
    Dim rcCap      As Double    'Expected cap-basis RC
    Dim addOnCap   As Double    'Unmargined aggregate add-on
    Dim multCap    As Double    'Cap-basis multiplier on Results
    Dim capEad     As Double    'Expected cap
    Dim eadMarg    As Double    'Margined EAD before the cap
    Dim eadCap     As Double    'Cap on Results
    Dim flag       As String    'Cap-applied flag on Results

'------------------------------------------------------------------------------
' CHECK
'------------------------------------------------------------------------------
        For r = 1 To RowCount(outputRows)
            If IsValidRow(outputRows, r) Then
                id = RowLabel(outputRows, r)
                flag = SafeStr(outputRows(r, COL_CAP_FLAG))
                If Not IsMarginedRow(outputRows, r) Then
                    If flag <> "n/a" Or Not IsEmpty(outputRows(r, COL_EAD_CAP)) Then
                        CapGap = id & " is unmargined but shows a cap"
                        Exit Function
                    End If
                Else
                    cCap = NettingSetInput(id, NS_NICA)
                    rcCap = Max2(ToDbl(outputRows(r, COL_V)) - cCap, 0#)
                    addOnCap = ToDbl(outputRows(r, COL_ADDON_CAP))
                    multCap = ToDbl(outputRows(r, COL_MULT_CAP))
                    capEad = ToDbl(outputRows(r, COL_ALPHA)) * (rcCap + multCap * addOnCap)
                    eadMarg = ToDbl(outputRows(r, COL_EAD_MARGIN))
                    eadCap = ToDbl(outputRows(r, COL_EAD_CAP))
                    If Not Near(ToDbl(outputRows(r, COL_C_CAP)), cCap) Then
                        CapGap = id & " cap-basis C " & ToDbl(outputRows(r, COL_C_CAP)) & ", expected " & cCap
                    ElseIf Not Near(ToDbl(outputRows(r, COL_RC_CAP)), rcCap) Then
                        CapGap = id & " cap-basis RC " & ToDbl(outputRows(r, COL_RC_CAP)) & _
                                 ", expected " & rcCap
                    ElseIf addOnCap < 0# Then
                        CapGap = id & " unmargined add-on is negative"
                    ElseIf Not Near(eadCap, capEad) Then
                        CapGap = id & " cap " & eadCap & ", alpha * (RC + PFE) on the cap basis " & capEad
                    ElseIf Not Near(ToDbl(outputRows(r, COL_EAD)), Min2(eadMarg, eadCap)) Then
                        CapGap = id & " EAD " & ToDbl(outputRows(r, COL_EAD)) & ", lower of margined " & _
                                 eadMarg & " and cap " & eadCap & " is " & Min2(eadMarg, eadCap)
                    ElseIf flag <> IIf(eadCap < eadMarg, "Y", "N") Then
                        CapGap = id & " cap flag " & flag & " with margined EAD " & eadMarg & _
                                 " and cap " & eadCap
                    Else
                        CapGap = OneMultiplierGap(id & " (cap basis)", ToDbl(outputRows(r, COL_V)) - cCap, _
                                                  addOnCap, multCap)
                    End If
                    If Len(CapGap) > 0 Then
                        Exit Function
                    End If
                End If
            End If
        Next r

End Function


Private Function WithheldGap( _
    ByVal outputRows As Variant) _
    As String
'
'==============================================================================
'                                 WithheldGap
'------------------------------------------------------------------------------
' PURPOSE
'   Check that a netting set whose EAD is withheld (INVALID, INCOMPLETE or
'   NO TRADES) shows no figures from MPOR to the cap-basis multiplier, so
'   that no partial result can be read as an EAD.
'
' RETURNS
'   The first violation; "" when none.
'
' UPDATED
'   2026-10-09
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
        For r = 1 To RowCount(outputRows)
            If Not IsValidRow(outputRows, r) And RowLabel(outputRows, r) <> "TOTAL" Then
                For col = 6 To COL_MULT_CAP
                    If Not IsEmpty(outputRows(r, col)) Then
                        WithheldGap = RowLabel(outputRows, r) & " (" & _
                                      SafeStr(outputRows(r, RS_STATUS_COL)) & ") shows " & _
                                      SafeStr(outputRows(r, col)) & " in column " & col
                        Exit Function
                    End If
                Next col
            End If
        Next r

End Function


Private Function TotalGap( _
    ByVal outputRows As Variant) _
    As String
'
'==============================================================================
'                                   TotalGap
'------------------------------------------------------------------------------
' PURPOSE
'   Check that the TOTAL row is the last row and holds the sums of RC,
'   PFE and EAD over the VALID netting sets.
'
' RETURNS
'   The first violation; "" when none.
'
' UPDATED
'   2026-10-09
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim r        As Long      'Results row
    Dim last     As Long      'The TOTAL row
    Dim sumRc    As Double    'Sum of RC
    Dim sumPfe   As Double    'Sum of PFE
    Dim sumEad   As Double    'Sum of EAD

'------------------------------------------------------------------------------
' CHECK
'------------------------------------------------------------------------------
        last = RowCount(outputRows)
        If last = 0 Then
            TotalGap = "Results is empty"
            Exit Function
        End If
        If RowLabel(outputRows, last) <> "TOTAL" Then
            TotalGap = "the last Results row is " & RowLabel(outputRows, last) & ", not TOTAL"
            Exit Function
        End If
        For r = 1 To last - 1
            If IsValidRow(outputRows, r) Then
                sumRc = sumRc + ToDbl(outputRows(r, COL_RC))
                sumPfe = sumPfe + ToDbl(outputRows(r, COL_PFE))
                sumEad = sumEad + ToDbl(outputRows(r, COL_EAD))
            End If
        Next r
        If Not Near(ToDbl(outputRows(last, COL_RC)), sumRc) Then
            TotalGap = "TOTAL RC " & ToDbl(outputRows(last, COL_RC)) & ", sum " & sumRc
        ElseIf Not Near(ToDbl(outputRows(last, COL_PFE)), sumPfe) Then
            TotalGap = "TOTAL PFE " & ToDbl(outputRows(last, COL_PFE)) & ", sum " & sumPfe
        ElseIf Not Near(ToDbl(outputRows(last, COL_EAD)), sumEad) Then
            TotalGap = "TOTAL EAD " & ToDbl(outputRows(last, COL_EAD)) & ", sum " & sumEad
        End If

End Function


Private Function IdenticalGap( _
    ByVal firstRows As Variant, _
    ByVal secondRows As Variant, _
    ByVal sheetName As String) _
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
'   firstRows, secondRows: the sheet's output block from each run.
'   sheetName: for the report.
'
' RETURNS
'   The first difference; "" when none.
'
' UPDATED
'   2026-10-09
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim r      As Long       'Row
    Dim col    As Long       'Column
    Dim a      As Variant    'First run's cell
    Dim b      As Variant    'Second run's cell
    Dim same   As Boolean    'The two cells are equal

'------------------------------------------------------------------------------
' COMPARE
'------------------------------------------------------------------------------
        If RowCount(firstRows) = 0 Then
            IdenticalGap = sheetName & " is empty"
            Exit Function
        End If
        If RowCount(secondRows) <> RowCount(firstRows) Then
            IdenticalGap = sheetName & " has " & RowCount(secondRows) & " rows, first run " & _
                           RowCount(firstRows)
            Exit Function
        End If
        For r = 1 To RowCount(firstRows)
            For col = LBound(firstRows, 2) To UBound(firstRows, 2)
                a = firstRows(r, col)
                b = secondRows(r, col)
                If IsError(a) Or IsError(b) Then
                    same = IsError(a) And IsError(b)
                ElseIf VarType(a) <> VarType(b) Then
                    same = False
                Else
                    same = (a = b)
                End If
                If Not same Then
                    IdenticalGap = sheetName & " row " & r & " column " & col & ": " & SafeStr(a) & _
                                   " became " & SafeStr(b)
                    Exit Function
                End If
            Next col
        Next r

End Function


Private Function IrOffsetGap( _
    ByVal bucketFormulaRows As Variant, _
    ByVal sumOfAbsoluteRows As Variant) _
    As String
'
'==============================================================================
'                                 IrOffsetGap
'------------------------------------------------------------------------------
' PURPOSE
'   Check that for every netting set VALID in both runs, the IR add-on
'   with the sum of absolute bucket values is not below the IR add-on
'   with the bucket formula.
'
' RETURNS
'   The first violation; "" when none.
'
' UPDATED
'   2026-10-09
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim r            As Long      'Row of bucketFormulaRows
    Dim q            As Long      'Matching row of sumOfAbsoluteRows
    Dim withOff      As Double    'IR add-on with the bucket formula
    Dim withoutOff   As Double    'IR add-on with the sum of absolutes

'------------------------------------------------------------------------------
' CHECK
'------------------------------------------------------------------------------
        For r = 1 To RowCount(bucketFormulaRows)
            If IsValidRow(bucketFormulaRows, r) Then
                q = MatchingRow(sumOfAbsoluteRows, RowLabel(bucketFormulaRows, r))
                If q = 0 Then
                    IrOffsetGap = RowLabel(bucketFormulaRows, r) & " missing from the second run"
                    Exit Function
                End If
                If IsValidRow(sumOfAbsoluteRows, q) Then
                    withOff = ToDbl(bucketFormulaRows(r, COL_IR_ADDON))
                    withoutOff = ToDbl(sumOfAbsoluteRows(q, COL_IR_ADDON))
                    If withoutOff < withOff And Not Near(withoutOff, withOff) Then
                        IrOffsetGap = RowLabel(bucketFormulaRows, r) & " IR add-on " & withoutOff & _
                                      " without the bucket formula, " & withOff & " with it"
                        Exit Function
                    End If
                End If
            End If
        Next r

End Function


Private Function IrHigherCount( _
    ByVal bucketFormulaRows As Variant, _
    ByVal sumOfAbsoluteRows As Variant) _
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
'   2026-10-09
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim r            As Long      'Row of bucketFormulaRows
    Dim q            As Long      'Matching row of sumOfAbsoluteRows
    Dim withOff      As Double    'IR add-on with the bucket formula
    Dim withoutOff   As Double    'IR add-on with the sum of absolutes

'------------------------------------------------------------------------------
' COUNT
'------------------------------------------------------------------------------
        For r = 1 To RowCount(bucketFormulaRows)
            If IsValidRow(bucketFormulaRows, r) Then
                q = MatchingRow(sumOfAbsoluteRows, RowLabel(bucketFormulaRows, r))
                If q > 0 Then
                    withOff = ToDbl(bucketFormulaRows(r, COL_IR_ADDON))
                    withoutOff = ToDbl(sumOfAbsoluteRows(q, COL_IR_ADDON))
                    If withoutOff > withOff And Not Near(withoutOff, withOff) Then
                        IrHigherCount = IrHigherCount + 1
                    End If
                End If
            End If
        Next r

End Function


'
'------------------------------------------------------------------------------
'
'                                   HELPERS
'
'------------------------------------------------------------------------------
'

Private Sub RecordRelation( _
    ByVal label As String, _
    ByVal detail As String)
'
'==============================================================================
'                                RecordRelation
'------------------------------------------------------------------------------
' PURPOSE
'   Record a relation check: it passes when no violation was found.
'
' UPDATED
'   2026-10-09
'==============================================================================
'
        TEST_CaseRunner.ExpectTrue label, Len(detail) = 0, detail, "illustrative"

End Sub


Private Function OneMultiplierGap( _
    ByVal label As String, _
    ByVal vLessC As Double, _
    ByVal addOnAmount As Double, _
    ByVal actual As Double) _
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
'   label: netting set, for the report.
'   vLessC: V - C in the calculation currency.
'   addOnAmount: aggregate add-on, >= 0.
'   actual: the multiplier on Results.
'
' RETURNS
'   The violation; "" when none.
'
' REFERENCE
'   CRE52.23; CRR Art. 278(3).
'
' UPDATED
'   2026-10-09
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim exponent   As Double    '(V - C) / (2 * (1 - floor) * add-on)
    Dim expected   As Double    'Multiplier from the formula

'------------------------------------------------------------------------------
' CHECK
'------------------------------------------------------------------------------
        If actual < mFloor - REL_TOL Or actual > 1# + REL_TOL Then
            OneMultiplierGap = label & " multiplier " & actual & " outside [" & mFloor & ", 1]"
        ElseIf vLessC >= 0# And Not Near(actual, 1#) Then
            OneMultiplierGap = label & " multiplier " & actual & " with V - C = " & vLessC & " >= 0"
        ElseIf addOnAmount > 0# Then
            exponent = vLessC / (2# * (1# - mFloor) * addOnAmount)
            If exponent >= 0# Then
                expected = 1#
            ElseIf exponent < -700# Then
                expected = mFloor
            Else
                expected = Min2(1#, mFloor + (1# - mFloor) * Exp(exponent))
            End If
            If Not Near(actual, expected) Then
                OneMultiplierGap = label & " multiplier " & actual & ", formula gives " & expected
            End If
        End If

End Function


Private Function NettingSetInput( _
    ByVal nettingSetId As String, _
    ByVal col As Long) _
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
'   2026-10-09
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ws   As Worksheet    'NettingSets sheet
    Dim r    As Long         'Input row; 0 when absent

'------------------------------------------------------------------------------
' READ
'------------------------------------------------------------------------------
        Set ws = GetSheet(SH_NS)
        r = FindHeaderRow(ws, NS_ID, nettingSetId)
        If r < FIRST_DATA_ROW Then
            Err.Raise ERR_TEST_SETUP, "TEST_Invariants.NettingSetInput", _
                      nettingSetId & " not found on " & SH_NS & "."
        End If
        NettingSetInput = ToDbl(ws.Cells(r, col).Value)

End Function


Private Function ParamsRow( _
    ByVal key As String) _
    As Long
'
'==============================================================================
'                                  ParamsRow
'------------------------------------------------------------------------------
' PURPOSE
'   Find the Params row whose column A holds a parameter code.
'
' ERROR POLICY
'   Raises ERR_TEST_SETUP when the code is absent; the suite then ends as
'   failed.
'
' UPDATED
'   2026-10-09
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim r   As Long    'Row found; 0 when absent

'------------------------------------------------------------------------------
' FIND
'------------------------------------------------------------------------------
        r = FindHeaderRow(GetSheet(SH_PARAMS), PRM_CODE_COL, key)
        If r = 0 Then
            Err.Raise ERR_TEST_SETUP, "TEST_Invariants.ParamsRow", _
                      "'" & key & "' not found in column A of Params."
        End If
        ParamsRow = r

End Function


Private Function SheetRows( _
    ByVal sheetName As String, _
    ByVal columnCount As Long) _
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
'   2026-10-09
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ws      As Worksheet    'Output sheet
    Dim lastR   As Long         'Last used row

'------------------------------------------------------------------------------
' READ
'------------------------------------------------------------------------------
        Set ws = GetSheet(sheetName)
        lastR = UsedLastRow(ws)
        If lastR >= FIRST_DATA_ROW Then
            SheetRows = ws.Range(ws.Cells(FIRST_DATA_ROW, 1), ws.Cells(lastR, columnCount)).Value
        End If

End Function


Private Function RowCount( _
    ByVal outputRows As Variant) _
    As Long
'
'==============================================================================
'                                   RowCount
'------------------------------------------------------------------------------
' PURPOSE
'   Number of rows in a block read by SheetRows; 0 when it is empty.
'
' UPDATED
'   2026-10-09
'==============================================================================
'
        If IsArray(outputRows) Then
            RowCount = UBound(outputRows, 1)
        End If

End Function


Private Function RowLabel( _
    ByVal outputRows As Variant, _
    ByVal r As Long) _
    As String
'
'==============================================================================
'                                    RowLabel
'------------------------------------------------------------------------------
' PURPOSE
'   The ID in column A of a Results row.
'
' UPDATED
'   2026-10-09
'==============================================================================
'
        RowLabel = SafeStr(outputRows(r, 1))

End Function


Private Function IsValidRow( _
    ByVal outputRows As Variant, _
    ByVal r As Long) _
    As Boolean
'
'==============================================================================
'                                  IsValidRow
'------------------------------------------------------------------------------
' PURPOSE
'   Tell whether a Results row is a netting set with an EAD (status VALID).
'
' UPDATED
'   2026-10-09
'==============================================================================
'
        IsValidRow = (SafeStr(outputRows(r, RS_STATUS_COL)) = "VALID")

End Function


Private Function IsMarginedRow( _
    ByVal outputRows As Variant, _
    ByVal r As Long) _
    As Boolean
'
'==============================================================================
'                                IsMarginedRow
'------------------------------------------------------------------------------
' PURPOSE
'   Tell whether a Results row is a margined netting set.
'
' UPDATED
'   2026-10-09
'==============================================================================
'
        IsMarginedRow = (SafeStr(outputRows(r, COL_MARGINED)) = "Y")

End Function


Private Function MatchingRow( _
    ByVal outputRows As Variant, _
    ByVal rowKey As String) _
    As Long
'
'==============================================================================
'                                 MatchingRow
'------------------------------------------------------------------------------
' PURPOSE
'   Find the row of a block whose column A holds rowKey.
'
' RETURNS
'   The row; 0 when absent.
'
' UPDATED
'   2026-10-09
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim r   As Long    'Row being compared

'------------------------------------------------------------------------------
' FIND
'------------------------------------------------------------------------------
        For r = 1 To RowCount(outputRows)
            If RowLabel(outputRows, r) = rowKey Then
                MatchingRow = r
                Exit Function
            End If
        Next r

End Function


Private Function Near( _
    ByVal actual As Double, _
    ByVal expected As Double) _
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
'   2026-10-09
'==============================================================================
'
        Near = Abs(actual - expected) <= REL_TOL * Max2(1#, Abs(expected))

End Function
