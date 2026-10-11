Attribute VB_Name = "TEST_InputValidation"
'==============================================================================
' MODULE: TEST_InputValidation
'------------------------------------------------------------------------------
' PURPOSE
'   Check that malformed inputs are rejected instead of silently changing
'   the trade population, the netting-set terms or the parameters (#35):
'   trade rows with data but no ID, duplicate trade and netting-set IDs, a
'   missing MtM, an unreadable date, netting-set fields that cannot be
'   read, an Excel error value (#N/A) in a trade, netting set or parameter,
'   a moved or renamed column, a broken parameter name, a parameter,
'   supervisory factor or currency listed twice with different values, and
'   a factor or rate the table loader would not read because of a blank
'   row or key; and, for aggregation (#37), a credit reference given two
'   sub-classes in either order, a basis trade without a hedging-set label,
'   a reserved character in a reference, and an interest-rate risk factor
'   that is not a currency; and, for the CRR scope (#33), the sum of
'   absolute IR bucket values in a CRR netting set and an IR risk factor
'   with a suffix other than -INFL; and, for the netting-set terms (#39),
'   VM on an unmargined netting set, a clearing role or margin-agreement
'   structure that is not a code, and the structures the engine does not
'   model. A blank optional field, a duplicate with the same values, an
'   inflation risk factor such as EUR-INFL, the sum of absolutes in a BCBS
'   netting set and VM posted on an unmargined BCBS netting set must still
'   be accepted. The same section checks the MPOR of each clearing role and
'   regime, and with disputes or a large netting set and weekly
'   remargining, against the rules recorded in the methodology.
'
' PUBLIC SURFACE
'   RunInputValidationTests is the entry point. Option Private Module keeps
'   it out of the external workbook automation API.
'
' DEPENDENCIES
'   TEST_CaseRunner, which writes the inputs, runs the engine, restores the
'   workbook and reports. These inputs cannot be expressed as JSON
'   fixtures, so they are written here directly, as TEST_CASES.md allows
'   for invalid inputs. The table cases find their rows on Params by key
'   (CO_OTHER, OT, USD, JPY, CHF), so they need the template's Params
'   sheet.
'
' WORKSHEET SAFETY
'   As TEST_CaseRunner: inputs, outputs, and the headers, Params cells and
'   workbook name a case patches are rewritten during the run and
'   restored. Use a development workbook.
'
' USAGE
'   Run TEST_InputValidation.RunInputValidationTests from the Immediate
'   window.
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
        Private Const VALUATION   As String = "2026-09-30"    'Valuation date of every case
        Private Const CASES       As Long = 53                'Cases in a complete run
        Private Const CHECKS      As Long = 60                'Checks in a complete run


'
'------------------------------------------------------------------------------
'
'                                 ENTRY POINT
'
'------------------------------------------------------------------------------
'

Public Sub RunInputValidationTests()
'
'==============================================================================
'                           RunInputValidationTests
'------------------------------------------------------------------------------
' PURPOSE
'   Run every case and print the TEST_CaseRunner report.
'
' UPDATED
'   2026-10-07
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim r   As Long    'Params row a table case patches

'------------------------------------------------------------------------------
' RUN CASES
'------------------------------------------------------------------------------
        TEST_CaseRunner.BeginSuite CASES, CHECKS
        On Error GoTo Failed

    'Trade rows.
        StartCase "trade-middle-row-without-id", "N"
        AddSwap "T1"
        AddSwap "T2"
        AddSwap "T3"
        TEST_CaseRunner.SetInputCell SH_TRADES, 2, TR_ID, Empty
        ExpectStatus "INCOMPLETE: 1 of 3 trade(s) rejected"

        StartCase "trade-trailing-row-without-id", "N"
        AddSwap "T1"
        AddSwap "T2"
        TEST_CaseRunner.SetInputCell SH_TRADES, 2, TR_ID, Empty
        ExpectStatus "INCOMPLETE: 1 of 2 trade(s) rejected"

        StartCase "trade-duplicate-id", "N"
        AddSwap "T1"
        AddSwap "T1"
        ExpectStatus "INCOMPLETE: 1 of 2 trade(s) rejected"

        StartCase "trade-missing-mtm", "N"
        AddSwap "T1"
        TEST_CaseRunner.SetInputCell SH_TRADES, 1, TR_MTM, Empty
        ExpectStatus "INCOMPLETE: 1 of 1 trade(s) rejected"

        StartCase "trade-unreadable-date", "N"
        AddSwap "T1"
        TEST_CaseRunner.SetInputCell SH_TRADES, 1, TR_MAT, "next year"
        ExpectStatus "INCOMPLETE: 1 of 1 trade(s) rejected"

    'Netting-set fields.
        StartCase "netting-set-unrecognised-flag", "N"
        TEST_CaseRunner.SetInputCell SH_NS, 1, NS_MARGINED, "Maybe"
        AddSwap "T1"
        ExpectStatus "INVALID: 1 input error(s)"

        StartCase "netting-set-non-numeric-amount", "N"
        TEST_CaseRunner.SetInputCell SH_NS, 1, NS_NICA, "abc"
        AddSwap "T1"
        ExpectStatus "INVALID: 1 input error(s)"

        StartCase "netting-set-unknown-regime", "N"
        TEST_CaseRunner.SetInputCell SH_NS, 1, NS_REGIME, "EU"
        AddSwap "T1"
        ExpectStatus "INVALID: 1 input error(s)"

        StartCase "netting-set-blank-flags-take-defaults", ""
        AddSwap "T1"
        ExpectStatus "VALID"

        StartCase "netting-set-duplicate-id", "N"
        AddSwap "T1"
        TEST_CaseRunner.SetInputCell SH_NS, 2, NS_ID, "NS1"
        ExpectStatus "INVALID: 1 input error(s)"

    'Excel error values are invalid, never blank.
        StartCase "trade-error-value-mtm", "N"
        AddSwap "T1"
        TEST_CaseRunner.SetInputCell SH_TRADES, 1, TR_MTM, CVErr(xlErrNA)
        ExpectStatus "INCOMPLETE: 1 of 1 trade(s) rejected"

        StartCase "netting-set-error-value-amount", "N"
        TEST_CaseRunner.SetInputCell SH_NS, 1, NS_NICA, CVErr(xlErrNA)
        AddSwap "T1"
        ExpectStatus "INVALID: 1 input error(s)"

        StartCase "params-error-value", "N"
        AddSwap "T1"
        TEST_CaseRunner.PatchCell SH_PARAMS, ParamsRow(PRM_ALPHA), PRM_VALUE_COL, CVErr(xlErrNA)
        ExpectStop "Parameter must be a number"

    'Sheet layout: a header that differs means a column moved.
        StartCase "layout-trades-header-renamed", "N"
        AddSwap "T1"
        TEST_CaseRunner.PatchCell SH_TRADES, HEADER_ROW, TR_NOTIONAL, "Nominal"
        ExpectStop "Header must be 'Notional'"

        StartCase "layout-netting-sets-column-shifted", "N"
        AddSwap "T1"
        TEST_CaseRunner.PatchCell SH_NS, HEADER_ROW, NS_NICA, "Threshold TH"
        ExpectStop "Header must be 'NICA'"

        StartCase "layout-params-value-header-renamed", "N"
        AddSwap "T1"
        TEST_CaseRunner.PatchCell SH_PARAMS, HEADER_ROW, PRM_VALUE_COL, "Amount"
        ExpectStop "Header must be 'Value'"

        StartCase "layout-broken-parameter-name", "N"
        AddSwap "T1"
        TEST_CaseRunner.PatchName PRM_ALPHA, "=#REF!"
        ExpectStop "no longer refers to a cell"

    'Duplicates: the same values are accepted, different values stop the run.
        StartCase "params-duplicate-different-value", "N"
        AddSwap "T1"
        AddParamRow PRM_ALPHA, 1.2
        ExpectStop "Parameter listed twice with different values"

        StartCase "params-alpha-zero", "N"
        AddSwap "T1"
        TEST_CaseRunner.PatchCell SH_PARAMS, ParamsRow(PRM_ALPHA), PRM_VALUE_COL, 0#
        ExpectStop "Alpha must be greater than zero"

        StartCase "params-alpha-negative", "N"
        AddSwap "T1"
        TEST_CaseRunner.PatchCell SH_PARAMS, ParamsRow(PRM_ALPHA), PRM_VALUE_COL, -1#
        ExpectStop "Alpha must be greater than zero"

        StartCase "params-duplicate-below-fx", "N"
        AddSwap "T1"
        r = UsedLastRow(GetSheet(SH_PARAMS)) + 2
        TEST_CaseRunner.PatchCell SH_PARAMS, r, PRM_CODE_COL, PRM_ALPHA
        TEST_CaseRunner.PatchCell SH_PARAMS, r, PRM_VALUE_COL, ParamsValue(PRM_ALPHA, PRM_VALUE_COL)
        ExpectStatus "VALID"

        StartCase "params-duplicate-same-value", "N"
        AddSwap "T1"
        AddParamRow PRM_ALPHA, ParamsValue(PRM_ALPHA, PRM_VALUE_COL)
        ExpectStatus "VALID"

    'CO_OTHER becomes a second CO_METALS row: it differs in category and
    'hedging set, unless those are patched too.
        StartCase "factor-table-duplicate-different-values", "N"
        AddSwap "T1"
        r = ParamsRow("CO_OTHER")
        TEST_CaseRunner.PatchCell SH_PARAMS, r, 1, "CO_METALS"
        ExpectStop "Supervisory factor key listed twice with different values"

        StartCase "factor-table-duplicate-same-values", "N"
        AddSwap "T1"
        r = ParamsRow("CO_OTHER")
        TEST_CaseRunner.PatchCell SH_PARAMS, r, 1, "CO_METALS"
        TEST_CaseRunner.PatchCell SH_PARAMS, r, 3, "METALS"
        TEST_CaseRunner.PatchCell SH_PARAMS, r, 7, "METALS"
        ExpectStatus "VALID"

    'CHF becomes a second USD row, at CHF's rate or at USD's.
        StartCase "fx-table-duplicate-different-rates", "N"
        AddSwap "T1"
        r = ParamsRow("CHF")
        TEST_CaseRunner.PatchCell SH_PARAMS, r, 1, "USD"
        ExpectStop "Currency listed twice with different rates"

        StartCase "fx-table-duplicate-same-rate", "N"
        AddSwap "T1"
        r = ParamsRow("CHF")
        TEST_CaseRunner.PatchCell SH_PARAMS, r, 1, "USD"
        TEST_CaseRunner.PatchCell SH_PARAMS, r, 3, ParamsValue("USD", 3)
        ExpectStatus "VALID"

    'Gaps: a table ends at its first blank key, so a row below a blank row,
    'or a value without a key, would not be read.
        StartCase "factor-table-blank-row-inside", "N"
        AddSwap "T1"
        ClearParamsRow ParamsRow("CO_OTHER"), 8
        ExpectStop "below a blank row"

        StartCase "factor-table-value-without-key", "N"
        AddSwap "T1"
        TEST_CaseRunner.PatchCell SH_PARAMS, ParamsRow("OT"), 1, Empty
        ExpectStop "has a value but no key"

        StartCase "fx-table-blank-row-inside", "N"
        AddSwap "T1"
        ClearParamsRow ParamsRow("JPY"), 3
        ExpectStop "below a blank row"

        StartCase "fx-table-rate-without-currency", "N"
        AddSwap "T1"
        TEST_CaseRunner.PatchCell SH_PARAMS, ParamsRow("CHF"), 1, Empty
        ExpectStop "has a value but no key"

    'Aggregation (#37): one reference, one sub-class, whatever the order.
        StartCase "aggregation-reference-two-sub-classes", "N"
        AddCredit "C1", "AA", "FIRM A"
        AddCredit "C2", "BBB", "FIRM A"
        ExpectStatus "INVALID: 1 input error(s)"

        StartCase "aggregation-reference-two-sub-classes-reversed", "N"
        AddCredit "C2", "BBB", "FIRM A"
        AddCredit "C1", "AA", "FIRM A"
        ExpectStatus "INVALID: 1 input error(s)"

        StartCase "aggregation-basis-without-label", "N"
        TEST_CaseRunner.AddTrade "B1", "IR", "", "EUR", "Linear", "Long", "", "Basis", "", _
                                 "10000", "0", "", "2031-09-30", "2031-09-30", "", "", "", "", "", ""
        ExpectStatus "INCOMPLETE: 1 of 1 trade(s) rejected"

        StartCase "aggregation-reserved-character", "N"
        AddCredit "C1", "AA", "FIRM|A"
        ExpectStatus "INCOMPLETE: 1 of 1 trade(s) rejected"

        StartCase "aggregation-ir-risk-factor-not-currency", "N"
        TEST_CaseRunner.AddTrade "T1", "IR", "", "EURO", "Linear", "Long", "", "Standard", "", _
                                 "10000", "0", "", "2031-09-30", "2031-09-30", "", "", "", "", "", ""
        ExpectStatus "INCOMPLETE: 1 of 1 trade(s) rejected"

    'Interest-rate aggregation (#33): the CRR has only the formula with
    'offsets across maturity buckets [Art. 280a(3)]; Basel also allows the
    'sum of absolute bucket values [CRE52.57(5)].
        StartCase "params-ir-sum-of-absolutes-crr", "N"
        AddSwap "T1"
        TEST_CaseRunner.PatchCell SH_PARAMS, ParamsRow(PRM_IRFULL), PRM_VALUE_COL, False
        ExpectStatus "INVALID: 1 input error(s)"

        StartCase "params-ir-sum-of-absolutes-bcbs", "N"
        TEST_CaseRunner.SetInputCell SH_NS, 1, NS_REGIME, "BCBS"
        AddSwap "T1"
        TEST_CaseRunner.PatchCell SH_PARAMS, ParamsRow(PRM_IRFULL), PRM_VALUE_COL, False
        ExpectStatus "VALID"

    'Inflation is interest rate, entered as a currency code with -INFL
    '[CRR Art. 277(4)(a), 277a(1)]; any other suffix is rejected.
        StartCase "trade-ir-inflation-risk-factor", "N"
        TEST_CaseRunner.AddTrade "T1", "IR", "", "EUR-INFL", "Linear", "Long", "", "Standard", "", _
                                 "10000", "0", "", "2031-09-30", "2031-09-30", "", "", "", "", "", ""
        ExpectStatus "VALID"

        StartCase "trade-ir-risk-factor-unknown-suffix", "N"
        TEST_CaseRunner.AddTrade "T1", "IR", "", "EUR-CPI", "Linear", "Long", "", "Standard", "", _
                                 "10000", "0", "", "2031-09-30", "2031-09-30", "", "", "", "", "", ""
        ExpectStatus "INCOMPLETE: 1 of 1 trade(s) rejected"

    'Collateral (#39). VM received means the counterparty posts VM, so the
    'set is margined [CRE52.2; CRR Art. 272(7)]. VM posted on an
    'unmargined set is a one-way agreement: BCBS includes it in C with a
    'negative sign [CRE52.10 footnote 2], so RC = 30 - (-1000); under CRR
    'the set is margined [Art. 272(7a)].
        StartCase "netting-set-unmargined-vm-received", "N"
        TEST_CaseRunner.SetInputCell SH_NS, 1, NS_VM, 1000#
        AddSwap "T1"
        ExpectStatus "INVALID: 1 input error(s)"

        StartCase "netting-set-unmargined-vm-received-bcbs", "N"
        TEST_CaseRunner.SetInputCell SH_NS, 1, NS_REGIME, "BCBS"
        TEST_CaseRunner.SetInputCell SH_NS, 1, NS_VM, 1000#
        AddSwap "T1"
        ExpectStatus "INVALID: 1 input error(s)"

        StartCase "netting-set-unmargined-vm-posted-crr", "N"
        TEST_CaseRunner.SetInputCell SH_NS, 1, NS_VM, -1000#
        AddSwap "T1"
        ExpectStatus "INVALID: 1 input error(s)"

        StartCase "netting-set-unmargined-vm-posted-bcbs", "N"
        TEST_CaseRunner.SetInputCell SH_NS, 1, NS_REGIME, "BCBS"
        TEST_CaseRunner.SetInputCell SH_NS, 1, NS_VM, -1000#
        AddSwap "T1"
        ExpectStatus "VALID"
        TEST_CaseRunner.ExpectNumber "rc", "", "replacement_cost", "1030", "0.000001", "0", "illustrative"

    'Client clearing role (#39): the 5-day floor for the clearing member in
    'both regimes [CRE54.12; CRR Art. 279c(1)], for the client only under
    'CRR; the BCBS client keeps the 10-day floor [CRE54.8, CRE54.14].
        StartCase "netting-set-clearing-role-not-a-code", "N"
        TEST_CaseRunner.SetInputCell SH_NS, 1, NS_CLEARED, "Y"
        AddSwap "T1"
        ExpectStatus "INVALID: 1 input error(s)"

        StartMarginedCase "netting-set-clearing-member-crr", "CRR", "CM", "1", "N", "N"
        ExpectMPOR 5

        StartMarginedCase "netting-set-clearing-member-bcbs", "BCBS", "CM", "1", "N", "N"
        ExpectMPOR 5

        StartMarginedCase "netting-set-clearing-client-crr", "CRR", "CLIENT", "1", "N", "N"
        ExpectMPOR 5

        StartMarginedCase "netting-set-clearing-client-bcbs", "BCBS", "CLIENT", "1", "N", "N"
        ExpectMPOR 10

    'MPOR rules recorded in #39: disputes double F + N - 1, which is the
    'BCBS rule [CRE52.50-.51] and above the CRR minimum [Art. 285(4)-(5)];
    'a large or illiquid set takes 20 + N - 1 [Art. 285(3), (5)], also
    'under BCBS [CRE52.51].
        StartMarginedCase "netting-set-mpor-disputes-weekly", "CRR", "N", "5", "N", "Y"
        ExpectMPOR 28

        StartMarginedCase "netting-set-mpor-large-weekly-bcbs", "BCBS", "N", "5", "Y", "N"
        ExpectMPOR 24

    'Margin-agreement structures the engine does not model (#39) are
    'declared and rejected [CRE52.74, CRE52.75; CRR Art. 274(4), 275(3)].
        StartCase "netting-set-margin-agreements-mixed", "N"
        TEST_CaseRunner.SetInputCell SH_NS, 1, NS_STRUCT, "MIXED"
        AddSwap "T1"
        ExpectStatus "INVALID: 1 input error(s)"

        StartCase "netting-set-margin-agreements-shared", "N"
        TEST_CaseRunner.SetInputCell SH_NS, 1, NS_STRUCT, "SHARED"
        AddSwap "T1"
        ExpectStatus "INVALID: 1 input error(s)"

        StartCase "netting-set-margin-agreements-not-a-code", "N"
        TEST_CaseRunner.SetInputCell SH_NS, 1, NS_STRUCT, "TWO"
        AddSwap "T1"
        ExpectStatus "INVALID: 1 input error(s)"

        TEST_CaseRunner.EndSuite ""
        Exit Sub

Failed:
        TEST_CaseRunner.EndSuite "unexpected error " & Err.Number & ": " & Err.Description

End Sub


'
'------------------------------------------------------------------------------
'
'                                   HELPERS
'
'------------------------------------------------------------------------------
'

Private Sub StartCase( _
    ByVal caseId As String, _
    ByVal flagValue As String)
'
'==============================================================================
'                                  StartCase
'------------------------------------------------------------------------------
' PURPOSE
'   Start a CRR case with one unmargined netting set NS1. flagValue is
'   written to every Y/N field: "N", or "" to leave them blank.
'
' UPDATED
'   2026-10-07
'==============================================================================
'
        TEST_CaseRunner.BeginCase "input-" & caseId, "CRR", VALUATION, "EUR"
        TEST_CaseRunner.AddNettingSet "NS1", flagValue, flagValue, "", flagValue, flagValue, "", _
                                 "0", "0", "0", "0", ""

End Sub


Private Sub StartMarginedCase( _
    ByVal caseId As String, _
    ByVal regimeCode As String, _
    ByVal roleCode As String, _
    ByVal remarginDays As String, _
    ByVal largeFlag As String, _
    ByVal disputesFlag As String)
'
'==============================================================================
'                              StartMarginedCase
'------------------------------------------------------------------------------
' PURPOSE
'   Start a case with one margined netting set NS1, without collateral,
'   threshold or MTA, in the given regime, and add a swap.
'
' UPDATED
'   2026-10-10
'==============================================================================
'
        TEST_CaseRunner.BeginCase "input-" & caseId, regimeCode, VALUATION, "EUR"
        TEST_CaseRunner.AddNettingSet "NS1", "Y", roleCode, remarginDays, largeFlag, disputesFlag, "", _
                                 "0", "0", "0", "0", ""
        AddSwap "T1"

End Sub


Private Sub ExpectMPOR( _
    ByVal expectedDays As Long)
'
'==============================================================================
'                                  ExpectMPOR
'------------------------------------------------------------------------------
' PURPOSE
'   Run the case and check that the netting set is VALID with the expected
'   MPOR in business days.
'
' UPDATED
'   2026-10-10
'==============================================================================
'
        ExpectStatus "VALID"
        TEST_CaseRunner.ExpectNumber "mpor", "", "margin_period_of_risk", CStr(expectedDays), "0", "0", _
                                     "illustrative"

End Sub


Private Sub AddSwap( _
    ByVal tradeId As String)
'
'==============================================================================
'                                   AddSwap
'------------------------------------------------------------------------------
' PURPOSE
'   Add a valid five-year EUR interest-rate swap to NS1.
'
' UPDATED
'   2026-10-07
'==============================================================================
'
        TEST_CaseRunner.AddTrade tradeId, "IR", "", "EUR", "Linear", "Long", "", "Standard", "", _
                            "10000", "30", "", "2031-09-30", "2031-09-30", "", "", "", "", "", ""

End Sub


Private Sub ExpectStop( _
    ByVal messagePart As String)
'
'==============================================================================
'                                  ExpectStop
'------------------------------------------------------------------------------
' PURPOSE
'   Run the case and check that the run stopped with an ERROR on Checks
'   containing messagePart.
'
' UPDATED
'   2026-10-07
'==============================================================================
'
        TEST_CaseRunner.RunCaseExpectingStop "stop", messagePart, "illustrative"

End Sub


Private Function ParamsRow( _
    ByVal key As String) _
    As Long
'
'==============================================================================
'                                  ParamsRow
'------------------------------------------------------------------------------
' PURPOSE
'   Return the Params row whose column A holds key: a parameter code, a
'   supervisory-factor key or a currency.
'
' ERROR POLICY
'   Raises ERR_TEST_SETUP when the key is not on Params, so that a changed
'   template fails the suite instead of patching the wrong row.
'
' UPDATED
'   2026-10-07
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim r   As Long    'Row found; 0 when absent

'------------------------------------------------------------------------------
' FIND
'------------------------------------------------------------------------------
        r = FindHeaderRow(GetSheet(SH_PARAMS), 1, key)
        If r = 0 Then
            Err.Raise ERR_TEST_SETUP, "TEST_InputValidation.ParamsRow", "'" & key & "' not found in column A of Params."
        End If
        ParamsRow = r

End Function


Private Function ParamsValue( _
    ByVal key As String, _
    ByVal col As Long) _
    As Variant
'
'==============================================================================
'                                 ParamsValue
'------------------------------------------------------------------------------
' PURPOSE
'   Return a value from the Params row whose column A holds key.
'
' UPDATED
'   2026-10-07
'==============================================================================
'
        ParamsValue = GetSheet(SH_PARAMS).Cells(ParamsRow(key), col).Value

End Function


Private Sub AddParamRow( _
    ByVal code As String, _
    ByVal newValue As Variant)
'
'==============================================================================
'                                 AddParamRow
'------------------------------------------------------------------------------
' PURPOSE
'   Add a second Params row for a parameter in the blank row directly below
'   the parameter list (LambdaThresholdCO is the last), where a user would
'   add one. A separate case covers a parameter row below the FX table.
'
' ERROR POLICY
'   Raises ERR_TEST_SETUP when that row is not blank, so that a changed
'   template fails the suite instead of overwriting a row.
'
' UPDATED
'   2026-10-07
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim r   As Long    'The new row

'------------------------------------------------------------------------------
' WRITE
'------------------------------------------------------------------------------
        r = ParamsRow(PRM_LAMCO) + 1
        If Not IsBlankCell(GetSheet(SH_PARAMS).Cells(r, PRM_CODE_COL).Value) Or _
           Not IsBlankCell(GetSheet(SH_PARAMS).Cells(r, PRM_VALUE_COL).Value) Then
            Err.Raise ERR_TEST_SETUP, "TEST_InputValidation.AddParamRow", "Params row " & r & " is not blank."
        End If
        TEST_CaseRunner.PatchCell SH_PARAMS, r, PRM_CODE_COL, code
        TEST_CaseRunner.PatchCell SH_PARAMS, r, PRM_VALUE_COL, newValue

End Sub


Private Sub ClearParamsRow( _
    ByVal rowNum As Long, _
    ByVal nCols As Long)
'
'==============================================================================
'                                ClearParamsRow
'------------------------------------------------------------------------------
' PURPOSE
'   Blank columns A to nCols of one Params row, as if a blank row had been
'   inserted in a table there. Restored with the case's other patches.
'
' UPDATED
'   2026-10-07
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim c   As Long    'Column being cleared

'------------------------------------------------------------------------------
' CLEAR
'------------------------------------------------------------------------------
        For c = 1 To nCols
            TEST_CaseRunner.PatchCell SH_PARAMS, rowNum, c, Empty
        Next c

End Sub


Private Sub AddCredit( _
    ByVal tradeId As String, _
    ByVal subClass As String, _
    ByVal reference As String)
'
'==============================================================================
'                                  AddCredit
'------------------------------------------------------------------------------
' PURPOSE
'   Add a five-year single-name credit default swap, protection bought, to
'   NS1.
'
' UPDATED
'   2026-10-07
'==============================================================================
'
        TEST_CaseRunner.AddTrade tradeId, "CR", subClass, reference, "Linear", "Long", "", "Standard", "", _
                                 "10000", "0", "", "2031-09-30", "2031-09-30", "", "", "", "", "", ""

End Sub


Private Sub ExpectStatus( _
    ByVal expected As String)
'
'==============================================================================
'                                 ExpectStatus
'------------------------------------------------------------------------------
' PURPOSE
'   Run the case and check the netting set's status on Results.
'
' UPDATED
'   2026-10-07
'==============================================================================
'
        TEST_CaseRunner.RunCase
        TEST_CaseRunner.ExpectText "status", "", "netting_set_status", expected, "illustrative"

End Sub
