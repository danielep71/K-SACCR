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
'   a moved or renamed column, a broken parameter or layout name, a parameter,
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
'   remargining, against the rules recorded in the methodology, and the
'   warning for a margined netting set of more than 5,000 trades whose
'   large or illiquid flag is N. Values outside the ranges of the data
'   validation (#147), written by VBA so that validation does not stop
'   them, must also be rejected: a multiplier floor of 1, a correlation
'   above 1, correlations that do not form a correlation matrix, an MPOR
'   floor of 0, a fractional business-day count, a negative supervisory
'   factor, a negative threshold or MTA, a remargining frequency or MPOR
'   override that is not a whole number of business days in range (0.5
'   and 10001 rejected, 10000 accepted), a reporting currency that is not
'   three letters A to Z, and an alpha override at or below zero; a
'   remargining frequency of 0 is raised to daily with a warning. A run
'   with the regulatory parameter set is REGULATORY, and one with a
'   different alpha, a zero factor or an alpha override is WHAT-IF with
'   the difference on Checks (#41); a DaysPerYear of 360 keeps it
'   regulatory; a reporting-currency FX rate other than 1 stops the run.
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
'   2026-10-11
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
        Private Const CASES       As Long = 97                'Cases in a complete run
        Private Const CHECKS      As Long = 129                'Checks in a complete run


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
'   2026-10-11
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim r   As Long    'Params row a table case patches
    Dim i   As Long    'Trade number in the large netting set

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
        TEST_CaseRunner.PatchCell SH_PARAMS, ParamsRow(PRM_ALPHA), ParamsValueCol(), CVErr(xlErrNA)
        ExpectStop "Parameter must be a number"

    'Sheet layout: a header that differs means a column moved.
        StartCase "layout-trades-header-renamed", "N"
        AddSwap "T1"
        TEST_CaseRunner.PatchCell SH_TRADES, TableHeaderRow(SH_TRADES), TR_NOTIONAL, "Nominal"
        ExpectStop "Header must be 'Notional'"

        StartCase "layout-netting-sets-column-shifted", "N"
        AddSwap "T1"
        TEST_CaseRunner.PatchCell SH_NS, TableHeaderRow(SH_NS), NS_NICA, "Threshold TH"
        ExpectStop "Header must be 'NICA'"

        StartCase "layout-params-value-header-renamed", "N"
        AddSwap "T1"
        TEST_CaseRunner.PatchCell SH_PARAMS, ParamsHeaderRow(), ParamsValueCol(), "Amount"
        ExpectStop "Header must be 'Value'"

        StartCase "layout-broken-parameter-name", "N"
        AddSwap "T1"
        TEST_CaseRunner.PatchName PRM_ALPHA, "=#REF!"
        ExpectStop "no longer refers to a cell"

    'Params is located through workbook names; a broken one stops the run.
        StartCase "layout-params-name-broken", "N"
        AddSwap "T1"
        TEST_CaseRunner.PatchName NM_PRM_CODE, "=#REF!"
        ExpectStop "'" & NM_PRM_CODE & "' is missing or no longer refers to a cell"

    'NettingSets and Trades are located through the workbook name of their
    'first header cell; a broken one stops the run.
        StartCase "layout-trades-name-broken", "N"
        AddSwap "T1"
        TEST_CaseRunner.PatchName NM_TR_HEADER, "=#REF!"
        ExpectStop "'" & NM_TR_HEADER & "' is missing or no longer refers"

        StartCase "layout-netting-sets-name-broken", "N"
        AddSwap "T1"
        TEST_CaseRunner.PatchName NM_NS_HEADER, "=#REF!"
        ExpectStop "'" & NM_NS_HEADER & "' is missing or no longer refers"

    'Duplicates: the same values are accepted, different values stop the run.
        StartCase "params-duplicate-different-value", "N"
        AddSwap "T1"
        AddParamRow PRM_ALPHA, 1.2
        ExpectStop "Parameter listed twice with different values"

        StartCase "params-alpha-zero", "N"
        AddSwap "T1"
        TEST_CaseRunner.PatchCell SH_PARAMS, ParamsRow(PRM_ALPHA), ParamsValueCol(), 0#
        ExpectStop "Alpha must be greater than zero"

        StartCase "params-alpha-negative", "N"
        AddSwap "T1"
        TEST_CaseRunner.PatchCell SH_PARAMS, ParamsRow(PRM_ALPHA), ParamsValueCol(), -1#
        ExpectStop "Alpha must be greater than zero"

    'Ranges of the data validation, checked by the engine for values that
    'bypass it (#147): a floor of 1 would divide by zero, correlations that
    'are no correlation matrix would make the IR effective notional undefined.
        StartCase "params-floor-one", "N"
        AddSwap "T1"
        TEST_CaseRunner.PatchCell SH_PARAMS, ParamsRow(PRM_FLOOR), ParamsValueCol(), 1#
        ExpectStop PRM_FLOOR & " must be at least"

        StartCase "params-correlation-above-one", "N"
        AddSwap "T1"
        TEST_CaseRunner.PatchCell SH_PARAMS, ParamsRow(PRM_RHO13), ParamsValueCol(), 1.5
        ExpectStop PRM_RHO13 & " must be from"

        StartCase "params-correlations-not-a-matrix", "N"
        AddSwap "T1"
        TEST_CaseRunner.PatchCell SH_PARAMS, ParamsRow(PRM_RHO12), ParamsValueCol(), 0.95
        ExpectStop "do not form a valid correlation matrix"

        StartCase "params-mpor-floor-zero", "N"
        AddSwap "T1"
        TEST_CaseRunner.PatchCell SH_PARAMS, ParamsRow(PRM_MPOR_BIL), ParamsValueCol(), 0#
        ExpectStop PRM_MPOR_BIL & " must be a whole number"

        StartCase "params-business-days-not-whole", "N"
        AddSwap "T1"
        TEST_CaseRunner.PatchCell SH_PARAMS, ParamsRow(PRM_BDYEAR), ParamsValueCol(), 250.5
        ExpectStop PRM_BDYEAR & " must be a whole number"

        StartCase "factor-table-negative-factor", "N"
        AddSwap "T1"
        TEST_CaseRunner.PatchCell SH_PARAMS, ParamsRow("EQ_INDEX"), TableCol(NM_SF_TABLE, 4), -0.2
        ExpectStop "Supervisory factor and correlation must be from 0 to 1"

        StartCase "params-duplicate-below-fx", "N"
        AddSwap "T1"
        r = UsedLastRow(GetSheet(SH_PARAMS)) + 2
        TEST_CaseRunner.PatchCell SH_PARAMS, r, ParamsCodeCol(), PRM_ALPHA
        TEST_CaseRunner.PatchCell SH_PARAMS, r, ParamsValueCol(), ParamsValue(PRM_ALPHA, ParamsValueCol())
        ExpectStatus "VALID"

        StartCase "params-duplicate-same-value", "N"
        AddSwap "T1"
        AddParamRow PRM_ALPHA, ParamsValue(PRM_ALPHA, ParamsValueCol())
        ExpectStatus "VALID"

    'CO_OTHER becomes a second CO_METALS row: it differs in category and
    'hedging set, unless those are patched too.
        StartCase "factor-table-duplicate-different-values", "N"
        AddSwap "T1"
        r = ParamsRow("CO_OTHER")
        TEST_CaseRunner.PatchCell SH_PARAMS, r, TableCol(NM_SF_TABLE, 1), "CO_METALS"
        ExpectStop "Supervisory factor key listed twice with different values"

        StartCase "factor-table-duplicate-same-values", "N"
        AddSwap "T1"
        r = ParamsRow("CO_OTHER")
        TEST_CaseRunner.PatchCell SH_PARAMS, r, TableCol(NM_SF_TABLE, 1), "CO_METALS"
        TEST_CaseRunner.PatchCell SH_PARAMS, r, TableCol(NM_SF_TABLE, 3), "METALS"
        TEST_CaseRunner.PatchCell SH_PARAMS, r, TableCol(NM_SF_TABLE, 7), "METALS"
        ExpectStatus "VALID"

    'CHF becomes a second USD row, at CHF's rate or at USD's.
        StartCase "fx-table-duplicate-different-rates", "N"
        AddSwap "T1"
        r = ParamsRow("CHF")
        TEST_CaseRunner.PatchCell SH_PARAMS, r, TableCol(NM_FX_TABLE, 1), "USD"
        ExpectStop "Currency listed twice with different rates"

        StartCase "fx-table-duplicate-same-rate", "N"
        AddSwap "T1"
        r = ParamsRow("CHF")
        TEST_CaseRunner.PatchCell SH_PARAMS, r, TableCol(NM_FX_TABLE, 1), "USD"
        TEST_CaseRunner.PatchCell SH_PARAMS, r, TableCol(NM_FX_TABLE, 3), ParamsValue("USD", TableCol(NM_FX_TABLE, 3))
        ExpectStatus "VALID"

    'Gaps: a table ends at its first blank key, so a row below a blank row,
    'or a value without a key, would not be read.
        StartCase "factor-table-blank-row-inside", "N"
        AddSwap "T1"
        ClearParamsRow ParamsRow("CO_OTHER"), TableCol(NM_SF_TABLE, 1), 8
        ExpectStop "below a blank row"

        StartCase "factor-table-value-without-key", "N"
        AddSwap "T1"
        TEST_CaseRunner.PatchCell SH_PARAMS, ParamsRow("OT"), TableCol(NM_SF_TABLE, 1), Empty
        ExpectStop "has a value but no key"

        StartCase "fx-table-blank-row-inside", "N"
        AddSwap "T1"
        ClearParamsRow ParamsRow("JPY"), TableCol(NM_FX_TABLE, 1), 3
        ExpectStop "below a blank row"

        StartCase "fx-table-rate-without-currency", "N"
        AddSwap "T1"
        TEST_CaseRunner.PatchCell SH_PARAMS, ParamsRow("CHF"), TableCol(NM_FX_TABLE, 1), Empty
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
        TEST_CaseRunner.PatchCell SH_PARAMS, ParamsRow(PRM_IRFULL), ParamsValueCol(), False
        ExpectStatus "INVALID: 1 input error(s)"

        StartCase "params-ir-sum-of-absolutes-bcbs", "N"
        TEST_CaseRunner.SetInputCell SH_NS, 1, NS_REGIME, "BCBS"
        AddSwap "T1"
        TEST_CaseRunner.PatchCell SH_PARAMS, ParamsRow(PRM_IRFULL), ParamsValueCol(), False
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

    'A margined netting set over 5,000 trades whose large or illiquid flag
    'is N stays VALID with a warning: the 20-day floor depends on the count
    'in the previous quarter [CRE52.51(1); CRR Art. 285(3)(a)].
        StartMarginedCase "netting-set-over-5000-trades", "CRR", "N", "1", "N", "N"
        For i = 2 To LARGE_NS_TRADES + 1
            AddSwap "T" & i
        Next i
        ExpectStatus "VALID"
        TEST_CaseRunner.ExpectChecksLine "warning", SEV_WARN, "more than " & LARGE_NS_TRADES, "illustrative"

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

    'Netting-set terms outside their range (#147): negative TH, MTA or
    'remargining frequency, and an alpha or MPOR override at or below zero,
    'make the set INVALID; a frequency below 1 day is raised to 1 with a
    'warning.
        StartCase "netting-set-negative-threshold", "N"
        TEST_CaseRunner.SetInputCell SH_NS, 1, NS_TH, -100#
        AddSwap "T1"
        ExpectStatus "INVALID: 1 input error(s)"

        StartCase "netting-set-negative-mta", "N"
        TEST_CaseRunner.SetInputCell SH_NS, 1, NS_MTA, -100#
        AddSwap "T1"
        ExpectStatus "INVALID: 1 input error(s)"

        StartCase "netting-set-alpha-override-zero", "N"
        TEST_CaseRunner.SetInputCell SH_NS, 1, NS_ALPHA, 0#
        AddSwap "T1"
        ExpectStatus "INVALID: 1 input error(s)"

        StartMarginedCase "netting-set-mpor-override-negative", "CRR", "N", "1", "N", "N"
        TEST_CaseRunner.SetInputCell SH_NS, 1, NS_MPOR, -5#
        ExpectStatus "INVALID: 1 input error(s)"

        StartMarginedCase "netting-set-remargin-negative", "CRR", "N", "-1", "N", "N"
        ExpectStatus "INVALID: 1 input error(s)"

        StartMarginedCase "netting-set-remargin-not-whole", "CRR", "N", "1.5", "N", "N"
        ExpectStatus "INVALID: 1 input error(s)"

        StartMarginedCase "netting-set-mpor-override-not-whole", "CRR", "N", "1", "N", "N"
        TEST_CaseRunner.SetInputCell SH_NS, 1, NS_MPOR, 10.5
        ExpectStatus "INVALID: 1 input error(s)"

        StartMarginedCase "netting-set-remargin-zero", "CRR", "N", "0", "N", "N"
        ExpectMPOR 10
        TEST_CaseRunner.ExpectChecksLine "warning", SEV_WARN, "daily remargining (1 BD) used", "illustrative"

    'Remargining boundaries (#147): half a day is not a whole number, 10000
    'is the highest accepted frequency (MPOR 10 + 10000 - 1), 10001 is above
    'it.
        StartMarginedCase "netting-set-remargin-half-day", "CRR", "N", "0.5", "N", "N"
        ExpectStatus "INVALID: 1 input error(s)"

        StartMarginedCase "netting-set-remargin-upper-bound", "CRR", "N", "10000", "N", "N"
        ExpectMPOR 10009

        StartMarginedCase "netting-set-remargin-above-bound", "CRR", "N", "10001", "N", "N"
        ExpectStatus "INVALID: 1 input error(s)"

    'Regulatory or what-if run (#41): a run whose parameters, factor table
    'and alpha overrides equal the regulatory set is REGULATORY and names
    'the set in the run summary; any difference makes it WHAT-IF and is
    'listed on Checks. DaysPerYear is a convention. The reporting
    'currency's FX rate must be 1.
        StartCase "run-class-regulatory", "N"
        AddSwap "T1"
        ExpectStatus "VALID"
        ExpectRunClass RUN_REGULATORY
        TEST_CaseRunner.ExpectTrue "run-summary", _
                                   InStr(SafeStr(GetSheet(SH_RESULTS).Range(RUNINFO_CELL).Value), REF_SET_ID) > 0, _
                                   "run summary does not name " & REF_SET_ID, "illustrative"

        StartCase "run-class-days-per-year-convention", "N"
        AddSwap "T1"
        TEST_CaseRunner.PatchCell SH_PARAMS, ParamsRow(PRM_DAYSYEAR), ParamsValueCol(), 360#
        ExpectStatus "VALID"
        ExpectRunClass RUN_REGULATORY

        StartCase "run-class-what-if-alpha", "N"
        AddSwap "T1"
        TEST_CaseRunner.PatchCell SH_PARAMS, ParamsRow(PRM_ALPHA), ParamsValueCol(), 1.5
        ExpectStatus "VALID"
        ExpectRunClass RUN_WHATIF
        TEST_CaseRunner.ExpectChecksLine "what-if", SEV_WARN, "What-if: " & PRM_ALPHA & " is", "illustrative"

        StartCase "run-class-what-if-zero-factor", "N"
        AddSwap "T1"
        TEST_CaseRunner.PatchCell SH_PARAMS, ParamsRow("EQ_INDEX"), TableCol(NM_SF_TABLE, 4), 0#
        ExpectStatus "VALID"
        ExpectRunClass RUN_WHATIF
        TEST_CaseRunner.ExpectChecksLine "what-if", SEV_WARN, "What-if: factor row EQ_INDEX differs", "illustrative"

        StartCase "run-class-what-if-alpha-override", "N"
        TEST_CaseRunner.SetInputCell SH_NS, 1, NS_ALPHA, 1.2
        AddSwap "T1"
        ExpectStatus "VALID"
        ExpectRunClass RUN_WHATIF
        TEST_CaseRunner.ExpectChecksLine "what-if", SEV_WARN, "What-if: alpha override", "illustrative"

        StartCase "fx-reporting-currency-rate-not-one", "N"
        AddSwap "T1"
        TEST_CaseRunner.PatchCell SH_PARAMS, ParamsRow("EUR"), TableCol(NM_FX_TABLE, 3), 0.9
        ExpectStop "Rate of the reporting currency must be 1"

    'Reporting currency format (#147): three letters A to Z, case and
    'surrounding spaces ignored. Written by VBA, so the worksheet validation
    'does not stop them; the format is checked, not ISO membership.
        StartReportingCcyCase "reporting-currency-digit", "E1R"
        ExpectStop "Reporting currency must be three letters"

        StartReportingCcyCase "reporting-currency-punctuation", "EU-"
        ExpectStop "Reporting currency must be three letters"

        StartReportingCcyCase "reporting-currency-four-letters", "EURO"
        ExpectStop "Reporting currency must be three letters"

        StartReportingCcyCase "reporting-currency-two-letters", "EU"
        ExpectStop "Reporting currency must be three letters"

        StartReportingCcyCase "reporting-currency-blank", ""
        ExpectStop "Reporting currency missing"

        StartReportingCcyCase "reporting-currency-lower-case", " eur "
        ExpectStatus "VALID"

    'Economic trades and legs (#38). Legs sharing an Economic trade ID form
    'one economic trade; exactly one leg, the MtM leg, carries its MtM and
    'every other leg holds 0. Legs in different asset classes map one
    'transaction to several risk categories [CRE52.28; CRR Art. 277(3)];
    'legs in one asset class must be options [CRE52.42-43; CRR Art.
    '274(6)]. Worked example, illustrative: an equity forward with an FX
    'driver, L1 EQ single name 10,000 with the MtM 30, L2 EUR/USD 10,000
    'with MtM 0. Add-ons 32% x 10,000 = 3,200 and 4% x 10,000 = 400, RC =
    '30, multiplier 1, EAD = 1.4 x (30 + 3,600) = 5,082. Entered as two
    'trades each with MtM 30, V would be 60 and EAD 5,124.
        StartCase "legs-two-risk-categories", "N"
        AddEquityForward "L1", "30"
        AddFxForward "L2", "0"
        SetLeg 1, "E1", "Y"
        SetLeg 2, "E1", "N"
        ExpectStatus "VALID"
        TEST_CaseRunner.ExpectNumber "replacement-cost", "", "replacement_cost", "30", "0.000001", "0.000000001", _
                                     "illustrative"
        TEST_CaseRunner.ExpectNumber "add-on-equity", "", "add_on.equity", "3200", "0.000001", "0.000000001", _
                                     "illustrative"
        TEST_CaseRunner.ExpectNumber "add-on-fx", "", "add_on.foreign_exchange", "400", "0.000001", _
                                     "0.000000001", "illustrative"
        TEST_CaseRunner.ExpectNumber "exposure-value", "", "exposure_value", "5082", "0.000001", "0.000000001", _
                                     "illustrative"
        TEST_CaseRunner.ExpectTrue "counts", CORE_Engine.EconomicTradesRead = 1 And CORE_Engine.TradesRead = 2 _
                                   And CORE_Engine.EconomicTradesUsed = 1 And CORE_Engine.TradesUsed = 2, _
                                   "economic trades " & CORE_Engine.EconomicTradesUsed & " of " & _
                                   CORE_Engine.EconomicTradesRead & ", legs " & CORE_Engine.TradesUsed & " of " & _
                                   CORE_Engine.TradesRead, "illustrative"

        StartCase "legs-two-mtm-legs", "N"
        AddEquityForward "L1", "30"
        AddFxForward "L2", "0"
        SetLeg 1, "E1", "Y"
        SetLeg 2, "E1", "Y"
        ExpectStatus "INCOMPLETE: 2 of 2 trade(s) rejected"
        TEST_CaseRunner.ExpectChecksLine "two-mtm-legs", SEV_ERROR, "exactly one leg must be the MtM leg", _
                                         "illustrative"

        StartCase "legs-repeated-mtm", "N"
        AddEquityForward "L1", "30"
        AddFxForward "L2", "30"
        SetLeg 1, "E1", "Y"
        SetLeg 2, "E1", "N"
        ExpectStatus "INCOMPLETE: 2 of 2 trade(s) rejected"
        TEST_CaseRunner.ExpectChecksLine "repeated-mtm", SEV_ERROR, "is not the MtM leg, so its MtM must be 0", _
                                         "illustrative"

        StartCase "legs-no-mtm-leg", "N"
        AddEquityForward "L1", "30"
        AddFxForward "L2", "0"
        SetLeg 1, "E1", "N"
        SetLeg 2, "E1", "N"
        ExpectStatus "INCOMPLETE: 2 of 2 trade(s) rejected"

        StartCase "legs-different-netting-sets", "N"
        TEST_CaseRunner.SetInputCell SH_NS, 2, NS_ID, "NS2"
        AddEquityForward "L1", "30"
        AddFxForward "L2", "0"
        SetLeg 1, "E1", "Y"
        SetLeg 2, "E1", "N"
        TEST_CaseRunner.SetInputCell SH_TRADES, 2, TR_NS, "NS2"
        ExpectStatus "INCOMPLETE: 1 of 1 trade(s) rejected"
        TEST_CaseRunner.ExpectChecksLine "netting-sets", SEV_ERROR, "its legs are in different netting sets", _
                                         "illustrative"

        StartCase "legs-linear-product-split", "N"
        AddSwap "L1"
        AddSwap "L2"
        TEST_CaseRunner.SetInputCell SH_TRADES, 2, TR_MTM, 0#
        SetLeg 1, "E1", "Y"
        SetLeg 2, "E1", "N"
        ExpectStatus "INCOMPLETE: 2 of 2 trade(s) rejected"
        TEST_CaseRunner.ExpectChecksLine "linear-split", SEV_ERROR, "that are not all options", "illustrative"

        StartCase "legs-option-structure", "N"
        AddEquityCall "L1", "Long", "100", "30"
        AddEquityCall "L2", "Short", "110", "0"
        SetLeg 1, "E1", "Y"
        SetLeg 2, "E1", "N"
        ExpectStatus "VALID"
        TEST_CaseRunner.ExpectTrue "counts", CORE_Engine.EconomicTradesRead = 1 And CORE_Engine.TradesRead = 2, _
                                   "economic trades " & CORE_Engine.EconomicTradesRead & ", legs " & _
                                   CORE_Engine.TradesRead, "illustrative"

        StartCase "legs-whole-trade-flagged-n", "N"
        AddSwap "T1"
        TEST_CaseRunner.SetInputCell SH_TRADES, 1, TR_MTMLEG, "N"
        ExpectStatus "INCOMPLETE: 1 of 1 trade(s) rejected"

        StartCase "legs-mtm-flag-not-a-code", "N"
        AddSwap "T1"
        SetLeg 1, "E1", "X"
        ExpectStatus "INCOMPLETE: 1 of 1 trade(s) rejected"

    'An error value in either column is rejected, not read as blank: as
    'blank, an Economic trade ID error would make a leg a whole trade
    'and an MtM leg error would pass as a non-MtM leg.
        StartCase "legs-economic-id-error", "N"
        AddSwap "T1"
        TEST_CaseRunner.SetInputCell SH_TRADES, 1, TR_ECON, CVErr(xlErrNA)
        ExpectStatus "INCOMPLETE: 1 of 1 trade(s) rejected"
        TEST_CaseRunner.ExpectChecksLine "economic-id-error", SEV_ERROR, _
                                         "Economic trade ID must be text or blank (found an error value", _
                                         "illustrative"

        StartCase "legs-mtm-flag-error", "N"
        AddEquityForward "L1", "30"
        AddFxForward "L2", "0"
        SetLeg 1, "E1", "Y"
        SetLeg 2, "E1", "N"
        TEST_CaseRunner.SetInputCell SH_TRADES, 2, TR_MTMLEG, CVErr(xlErrNA)
        ExpectStatus "INCOMPLETE: 2 of 2 trade(s) rejected"
        TEST_CaseRunner.ExpectChecksLine "mtm-flag-error", SEV_ERROR, _
                                         "MtM leg must be Y or N (found an error value", "illustrative"

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


Private Sub ExpectRunClass( _
    ByVal expected As String)
'
'==============================================================================
'                                ExpectRunClass
'------------------------------------------------------------------------------
' PURPOSE
'   Check the class of the case's run: RUN_REGULATORY or RUN_WHATIF (#41).
'
' UPDATED
'   2026-10-11
'==============================================================================
'
        TEST_CaseRunner.ExpectTrue "run-class", CORE_Engine.RunClass = expected, _
                                   "run class " & CORE_Engine.RunClass & ", expected " & expected, "illustrative"

End Sub


Private Sub StartReportingCcyCase( _
    ByVal caseId As String, _
    ByVal code As String)
'
'==============================================================================
'                            StartReportingCcyCase
'------------------------------------------------------------------------------
' PURPOSE
'   Start a CRR case with one swap and the reporting currency written as
'   code, bypassing the worksheet validation.
'
' UPDATED
'   2026-10-12
'==============================================================================
'
        StartCase caseId, "N"
        AddSwap "T1"
        TEST_CaseRunner.PatchCell SH_PARAMS, ParamsRow(PRM_REPCCY), ParamsValueCol(), code

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
'   Return the Params row that holds key: a parameter code in the code
'   column, a supervisory-factor key or a currency in its table's first
'   column, as the workbook names locate them.
'
' ERROR POLICY
'   Raises ERR_TEST_SETUP when the key is not on Params, so that a changed
'   template fails the suite instead of patching the wrong row.
'
' UPDATED
'   2026-10-11
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim r      As Long       'Row found; 0 when absent
    Dim cols   As Variant    'Columns searched, as table columns
    Dim i      As Long       'Index into cols

'------------------------------------------------------------------------------
' FIND
'------------------------------------------------------------------------------
        cols = Array(ParamsCodeCol(), TableCol(NM_SF_TABLE, 1), TableCol(NM_FX_TABLE, 1))
        For i = LBound(cols) To UBound(cols)
            r = FindHeaderRow(GetSheet(SH_PARAMS), CLng(cols(i)), key)
            If r > 0 Then
                Exit For
            End If
        Next i
        If r = 0 Then
            Err.Raise ERR_TEST_SETUP, "TEST_InputValidation.ParamsRow", "'" & key & "' not found on Params."
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
'   Return a value from the Params row that holds key (see ParamsRow), at
'   table column col.
'
' UPDATED
'   2026-10-11
'==============================================================================
'
        ParamsValue = TableCell(GetSheet(SH_PARAMS), ParamsRow(key), col).Value

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
'   2026-10-11
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
        If Not IsBlankCell(TableCell(GetSheet(SH_PARAMS), r, ParamsCodeCol()).Value) Or _
           Not IsBlankCell(TableCell(GetSheet(SH_PARAMS), r, ParamsValueCol()).Value) Then
            Err.Raise ERR_TEST_SETUP, "TEST_InputValidation.AddParamRow", "Params row " & r & " is not blank."
        End If
        TEST_CaseRunner.PatchCell SH_PARAMS, r, ParamsCodeCol(), code
        TEST_CaseRunner.PatchCell SH_PARAMS, r, ParamsValueCol(), newValue

End Sub


Private Sub ClearParamsRow( _
    ByVal rowNum As Long, _
    ByVal firstCol As Long, _
    ByVal nCols As Long)
'
'==============================================================================
'                                ClearParamsRow
'------------------------------------------------------------------------------
' PURPOSE
'   Blank nCols table columns from firstCol of one Params row, as if a blank
'   row had been inserted in a table there. Restored with the case's other
'   patches.
'
' UPDATED
'   2026-10-11
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim c   As Long    'Column being cleared

'------------------------------------------------------------------------------
' CLEAR
'------------------------------------------------------------------------------
        For c = firstCol To firstCol + nCols - 1
            TEST_CaseRunner.PatchCell SH_PARAMS, rowNum, c, Empty
        Next c

End Sub


Private Function TableCol( _
    ByVal nameText As String, _
    ByVal k As Long) _
    As Long
'
'==============================================================================
'                                   TableCol
'------------------------------------------------------------------------------
' PURPOSE
'   Return the table column of column k of a Params lookup table, counted
'   from the top-left cell of its workbook name (1 for the keys).
'
' UPDATED
'   2026-10-11
'==============================================================================
'
        TableCol = NamedTableCol(nameText) + k - 1

End Function


Private Sub AddEquityForward( _
    ByVal tradeId As String, _
    ByVal marketValue As String)
'
'==============================================================================
'                               AddEquityForward
'------------------------------------------------------------------------------
' PURPOSE
'   Add a five-year single-name equity forward, long, notional 10,000, to
'   NS1 (#38).
'
' UPDATED
'   2026-10-12
'==============================================================================
'
        TEST_CaseRunner.AddTrade tradeId, "EQ", "SINGLE", "ACME", "Linear", "Long", "", "Standard", "", _
                                 "10000", marketValue, "", "2031-09-30", "2031-09-30", "", "", "", "", "", ""

End Sub


Private Sub AddFxForward( _
    ByVal tradeId As String, _
    ByVal marketValue As String)
'
'==============================================================================
'                                AddFxForward
'------------------------------------------------------------------------------
' PURPOSE
'   Add a five-year EUR/USD forward, long, notional 10,000, to NS1 (#38).
'
' UPDATED
'   2026-10-12
'==============================================================================
'
        TEST_CaseRunner.AddTrade tradeId, "FX", "", "EUR/USD", "Linear", "Long", "", "Standard", "", _
                                 "10000", marketValue, "", "2031-09-30", "2031-09-30", "", "", "", "", "", ""

End Sub


Private Sub AddEquityCall( _
    ByVal tradeId As String, _
    ByVal direction As String, _
    ByVal strikePrice As String, _
    ByVal marketValue As String)
'
'==============================================================================
'                                AddEquityCall
'------------------------------------------------------------------------------
' PURPOSE
'   Add a one-year single-name equity call, notional 10,000, underlying
'   price 100, to NS1: two of them with different strikes make a call
'   spread (#38).
'
' UPDATED
'   2026-10-12
'==============================================================================
'
        TEST_CaseRunner.AddTrade tradeId, "EQ", "SINGLE", "ACME", "Option", direction, "Call", "Standard", "", _
                                 "10000", marketValue, "", "2027-09-30", "2027-09-30", "2027-09-30", "100", _
                                 strikePrice, "", "", ""

End Sub


Private Sub SetLeg( _
    ByVal dataRow As Long, _
    ByVal economicTradeId As String, _
    ByVal mtmLegFlag As String)
'
'==============================================================================
'                                    SetLeg
'------------------------------------------------------------------------------
' PURPOSE
'   Make a trade row a leg: its Economic trade ID and its MtM leg flag (#38).
'
' UPDATED
'   2026-10-12
'==============================================================================
'
        TEST_CaseRunner.SetInputCell SH_TRADES, dataRow, TR_ECON, economicTradeId
        TEST_CaseRunner.SetInputCell SH_TRADES, dataRow, TR_MTMLEG, mtmLegFlag

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
