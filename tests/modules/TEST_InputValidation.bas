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
'   For the input conventions (#40): a text date is read only as
'   yyyy-mm-dd, a workbook in the 1904 date system stops the run, the IR
'   maturity buckets split at exactly 1 and 5 years with the regime's edge
'   at 1 year, which changes the IR add-on of offsetting swaps (#191), a
'   trade or option ending on the reporting date is rejected, an option
'   expiring after its maturity is rejected, both currencies of an FX pair
'   need an FX rate, the notional must be in one of them, and a pair
'   entered the other way round offsets the same pair entered in order.
'   An Excel error value in an optional text field of a trade, in a
'   netting set's regime or in the Params regime is an input error, not
'   a blank that takes a default (#194), and a Currency number format
'   does not round an input (#196). Results has one TOTAL row per regime,
'   after the netting sets, with the number of VALID and withheld netting
'   sets it covers (#198).
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
'   2026-10-13
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
        Private Const CASES       As Long = 126               'Cases in a complete run
        Private Const CHECKS      As Long = 190               'Checks in a complete run


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
'   2026-10-13
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim r        As Long       'Params row a table case patches
    Dim i        As Long       'Trade number in the large netting set
    Dim fields   As Variant    'Trades column and field name, in pairs
    Dim f        As Long       'Index of a column in fields
    Dim savedFormat As String  'Number format of a cell a case changes

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

    'Dates (#40). Text is read as a date only in the ISO form yyyy-mm-dd,
    'whatever the regional settings; 30/09/2031, a serial written as text
    'and an impossible day are rejected. The text dates are written with a
    'leading apostrophe, so that Excel keeps them as text. ISO text gives
    'the same adjusted notional as a date: 10,000 x (1 - exp(-0.05 x
    '1826/365)) / 0.05.
        StartCase "date-text-forms", "N"
        TEST_CaseRunner.ExpectTrue "to-serial", _
            CORE_Util.ToSerial("2031-09-30") = CDbl(DateSerial(2031, 9, 30)) And _
            CORE_Util.ToSerial(" 2031-09-30 ") = CDbl(DateSerial(2031, 9, 30)) And _
            CORE_Util.ToSerial(DateSerial(2031, 9, 30)) = CDbl(DateSerial(2031, 9, 30)) And _
            CORE_Util.ToSerial(47756#) = 47756# And CORE_Util.ToSerial("47756") = -1# And _
            CORE_Util.ToSerial("30/09/2031") = -1# And CORE_Util.ToSerial("09/30/2031") = -1# And _
            CORE_Util.ToSerial("2031-9-30") = -1# And CORE_Util.ToSerial("2031-02-30") = -1# And _
            CORE_Util.ToSerial("30 September 2031") = -1#, _
            "ToSerial read a text date other than yyyy-mm-dd, or misread one", "illustrative"

        StartCase "date-iso-text", "N"
        AddSwap "T1"
        TEST_CaseRunner.SetInputCell SH_TRADES, 1, TR_END, "'2031-09-30"
        TEST_CaseRunner.SetInputCell SH_TRADES, 1, TR_MAT, "'2031-09-30"
        ExpectStatus "VALID"
        TEST_CaseRunner.ExpectNumber "adjusted-notional", "T1", "adjusted_notional", "44261.1789321023", _
                                     "0.000001", "0.000000001", "illustrative"

        StartCase "date-locale-text", "N"
        AddSwap "T1"
        TEST_CaseRunner.SetInputCell SH_TRADES, 1, TR_END, "'30/09/2031"
        TEST_CaseRunner.SetInputCell SH_TRADES, 1, TR_MAT, "'30/09/2031"
        ExpectStatus "INCOMPLETE: 1 of 1 trade(s) rejected"
        TEST_CaseRunner.ExpectChecksLine "not-a-date", SEV_ERROR, "maturity date is not a date", "illustrative"

        StartCase "reporting-date-iso-text", "N"
        AddSwap "T1"
        TEST_CaseRunner.PatchCell SH_PARAMS, ParamsRow(PRM_ASOF), ParamsValueCol(), "'2026-09-30"
        ExpectStatus "VALID"

        StartCase "reporting-date-locale-text", "N"
        AddSwap "T1"
        TEST_CaseRunner.PatchCell SH_PARAMS, ParamsRow(PRM_ASOF), ParamsValueCol(), "'30/09/2026"
        ExpectStop "Reporting date missing or not a date"

    'The 1904 date system stops the run; the workbook is set back at once.
        StartCase "date-system-1904", "N"
        AddSwap "T1"
        ExpectStopIn1904 "1904 date system"

    'Dates on the reporting date and the IR maturity buckets (#40, #191).
    'E is the end date in years of 365 days from 2026-09-30: 364 days is
    'under 1 year, 365 days exactly 1 year, 1,825 days exactly 5 years and
    '1,826 days (2031-09-30) over 5 years. Under CRR bucket 1 includes
    'exactly 1 year [Art. 280a(3), Table 2: > 0 and <= 1]; under BCBS it
    'does not [CRE52.57(3): less than one year]. Exactly 5 years is in
    'bucket 2 in both. A trade or option ending on the reporting date is
    'rejected; a start on it is S = 0.
        StartCase "ir-bucket-boundaries-crr", "N"
        AddBucketSwaps
        ExpectStatus "VALID"
        TEST_CaseRunner.ExpectText "bucket-364", "B364", "bucket", "<1Y", "illustrative"
        TEST_CaseRunner.ExpectText "bucket-365", "B365", "bucket", "<1Y", "illustrative"
        TEST_CaseRunner.ExpectText "bucket-1825", "B1825", "bucket", "1-5Y", "illustrative"
        TEST_CaseRunner.ExpectText "bucket-1826", "B1826", "bucket", ">5Y", "illustrative"

        StartCase "ir-bucket-boundaries-bcbs", "N"
        TEST_CaseRunner.SetInputCell SH_NS, 1, NS_REGIME, "BCBS"
        AddBucketSwaps
        ExpectStatus "VALID"
        TEST_CaseRunner.ExpectText "bucket-364", "B364", "bucket", "<1Y", "illustrative"
        TEST_CaseRunner.ExpectText "bucket-365", "B365", "bucket", "1-5Y", "illustrative"
        TEST_CaseRunner.ExpectText "bucket-1825", "B1825", "bucket", "1-5Y", "illustrative"
        TEST_CaseRunner.ExpectText "bucket-1826", "B1826", "bucket", ">5Y", "illustrative"

    'The edge changes the add-on, not only the label (#191). A long swap
    'ending at exactly 1 year and a short one ending at 2 years, notional
    '10,000 each, MF 1: D1 = 10,000 x SD(1) = 9,754.12 and D2 = -10,000 x
    'SD(2) = -19,032.52, SD(E) = (1 - exp(-0.05 E)) / 0.05. Under CRR they
    'are in buckets 1 and 2: add-on = 0.5% x sqrt(D1^2 + D2^2 + 1.4 D1 D2)
    '= 70.2631. Under BCBS both are in bucket 2: 0.5% x |D1 + D2| = 46.3920.
        StartCase "ir-bucket-offset-crr", "N"
        AddOffsetSwaps
        ExpectStatus "VALID"
        TEST_CaseRunner.ExpectNumber "add-on-ir", "", "add_on.interest_rate", "70.2630670897333", _
                                     "0.000001", "0.000000001", "illustrative"

        StartCase "ir-bucket-offset-bcbs", "N"
        TEST_CaseRunner.SetInputCell SH_NS, 1, NS_REGIME, "BCBS"
        AddOffsetSwaps
        ExpectStatus "VALID"
        TEST_CaseRunner.ExpectNumber "add-on-ir", "", "add_on.interest_rate", "46.3920064647545", _
                                     "0.000001", "0.000000001", "illustrative"

        StartCase "trade-maturing-on-reporting-date", "N"
        AddSwapTo "T1", VALUATION
        ExpectStatus "INCOMPLETE: 1 of 1 trade(s) rejected"
        TEST_CaseRunner.ExpectChecksLine "matured", SEV_ERROR, "trade has matured", "illustrative"

        StartCase "trade-starting-on-reporting-date", "N"
        AddSwap "T1"
        TEST_CaseRunner.SetInputCell SH_TRADES, 1, TR_START, DateSerial(2026, 9, 30)
        ExpectStatus "VALID"
        TEST_CaseRunner.ExpectNumber "adjusted-notional", "T1", "adjusted_notional", "44261.1789321023", _
                                     "0.000001", "0.000000001", "illustrative"

        StartCase "option-expiring-on-reporting-date", "N"
        AddEquityCall "T1", "Long", "100", "30"
        TEST_CaseRunner.SetInputCell SH_TRADES, 1, TR_EXPIRY, DateSerial(2026, 9, 30)
        ExpectStatus "INCOMPLETE: 1 of 1 trade(s) rejected"
        TEST_CaseRunner.ExpectChecksLine "expired", SEV_ERROR, "option has expired", "illustrative"

        StartCase "option-expiry-after-maturity", "N"
        AddEquityCall "T1", "Long", "100", "30"
        TEST_CaseRunner.SetInputCell SH_TRADES, 1, TR_EXPIRY, DateSerial(2028, 9, 30)
        ExpectStatus "INCOMPLETE: 1 of 1 trade(s) rejected"
        TEST_CaseRunner.ExpectChecksLine "expiry-after-maturity", SEV_ERROR, _
                                         "option expiry is after the maturity date", "illustrative"

    'Currency pairs (#40): both currencies need an FX rate on Params, and
    'the notional is in one of them. USD/EUR long is EUR/USD short in the
    'EUR/USD hedging set, so it offsets a EUR/USD long of the same
    'notional: FX add-on 0.
        StartCase "fx-pair-unknown-currency", "N"
        AddFxForwardOn "T1", "EUR/XYZ", "Long"
        ExpectStatus "INCOMPLETE: 1 of 1 trade(s) rejected"
        TEST_CaseRunner.ExpectChecksLine "unknown-currency", SEV_ERROR, _
                                         "no FX rate for currency 'XYZ' of the pair EUR/XYZ", "illustrative"

        StartCase "fx-notional-outside-pair", "N"
        AddFxForwardOn "T1", "USD/JPY", "Long"
        ExpectStatus "INCOMPLETE: 1 of 1 trade(s) rejected"
        TEST_CaseRunner.ExpectChecksLine "notional-outside-pair", SEV_ERROR, _
                                         "notional currency 'EUR' is not a currency of the pair JPY/USD", _
                                         "illustrative"

        StartCase "fx-reversed-pair-linear", "N"
        AddFxForwardOn "T1", "EUR/USD", "Long"
        AddFxForwardOn "T2", "USDEUR", "Long"
        ExpectStatus "VALID"
        TEST_CaseRunner.ExpectText "hedging-set", "T2", "hedging_set", "FX|EUR/USD|STD", "illustrative"
        TEST_CaseRunner.ExpectNumber "delta", "T2", "supervisory_delta", "-1", "0", "0", "illustrative"
        TEST_CaseRunner.ExpectNumber "add-on-fx", "", "add_on.foreign_exchange", "0", "0.000001", "0", _
                                     "illustrative"

    'An Excel error value, such as #N/A from a broken lookup, in an optional
    'text field is an input error, not a blank that takes the field's
    'default (#194): Instrument would become Linear, Nature Standard, the
    'notional currency the reporting currency, the MtM currency the
    'notional currency, and a netting-set or Params regime the default.
        fields = Array(TR_INSTR, "Instrument", TR_NATURE, "Nature", TR_SUB, "Sub-class", _
                       TR_LABEL, "Hedging-set label", TR_NCCY, "Notional currency", TR_MCCY, "MtM currency")
        For f = LBound(fields) To UBound(fields) Step 2
            StartCase "text-error-" & LCase$(Replace(CStr(fields(f + 1)), " ", "-")), "N"
            AddSwap "T1"
            TEST_CaseRunner.SetInputCell SH_TRADES, 1, CLng(fields(f)), CVErr(xlErrNA)
            ExpectStatus "INCOMPLETE: 1 of 1 trade(s) rejected"
            TEST_CaseRunner.ExpectChecksLine "error-value", SEV_ERROR, _
                                             CStr(fields(f + 1)) & " must be text or blank (found an error value)", _
                                             "illustrative"
        Next f

        StartCase "text-error-netting-set-regime", "N"
        TEST_CaseRunner.SetInputCell SH_NS, 1, NS_REGIME, CVErr(xlErrNA)
        AddSwap "T1"
        ExpectStatus "INVALID: 1 input error(s)"
        TEST_CaseRunner.ExpectChecksLine "error-value", SEV_ERROR, _
                                         "Regime override must be BCBS or CRR (found an error value", "illustrative"

        StartCase "text-error-params-regime", "N"
        AddSwap "T1"
        TEST_CaseRunner.PatchCell SH_PARAMS, ParamsRow(PRM_REGIME), ParamsValueCol(), CVErr(xlErrNA)
        ExpectStop "Regime must be BCBS or CRR (found an error value)"

    'Number formats do not change what is read (#196): with a Currency
    'format, Range.Value rounds to four decimals. A one-year single-name
    'equity call, price 0.012345, strike 0.0123, volatility 120%, has delta
    'N(d1) = 0.72676 with d1 = (ln(P/K) + 0.5 x 1.2^2) / 1.2; read as
    '0.0123 the price would give 0.72575. Both formats must give 0.72676.
    'The cell formats are put back at the end of the case.
        StartCase "number-format-general", "N"
        AddSmallOption
        ExpectStatus "VALID"
        TEST_CaseRunner.ExpectNumber "delta", "T1", "supervisory_delta", "0.726760030216713", _
                                     "0.000001", "0", "illustrative"

        StartCase "number-format-currency", "N"
        AddSmallOption
        r = TableFirstRow(SH_TRADES)
        savedFormat = TableCell(GetSheet(SH_TRADES), r, TR_PRICE).NumberFormat
        TableCell(GetSheet(SH_TRADES), r, TR_PRICE).NumberFormat = "$#,##0.00"
        ExpectStatus "VALID"
        TEST_CaseRunner.ExpectNumber "delta", "T1", "supervisory_delta", "0.726760030216713", _
                                     "0.000001", "0", "illustrative"
        TableCell(GetSheet(SH_TRADES), r, TR_PRICE).NumberFormat = savedFormat

    'Totals per regime (#198): one TOTAL row for each regime with netting
    'sets, CRR first, after the netting-set rows; each sums the regime's
    'VALID netting sets and says how many it covers and how many were
    'withheld. NS1 (CRR) and NS2 (BCBS) hold one swap each; NS3 (CRR) holds
    'a swap without MtM, so its EAD is withheld.
        StartCase "totals-per-regime", "N"
        TEST_CaseRunner.SetInputCell SH_NS, 2, NS_ID, "NS2"
        TEST_CaseRunner.SetInputCell SH_NS, 2, NS_REGIME, "BCBS"
        TEST_CaseRunner.SetInputCell SH_NS, 3, NS_ID, "NS3"
        TEST_CaseRunner.SetInputCell SH_NS, 3, NS_REGIME, "CRR"
        AddSwap "T1"
        AddSwap "T2"
        AddSwap "T3"
        TEST_CaseRunner.SetInputCell SH_TRADES, 2, TR_NS, "NS2"
        TEST_CaseRunner.SetInputCell SH_TRADES, 3, TR_NS, "NS3"
        TEST_CaseRunner.SetInputCell SH_TRADES, 3, TR_MTM, Empty
        TEST_CaseRunner.RunCase
        ExpectTotal RG_CRR, "1 of 2 netting set(s) VALID; 1 withheld", "NS1"
        ExpectTotal RG_BCBS, "1 of 1 netting set(s) VALID; 0 withheld", "NS2"
        TEST_CaseRunner.ExpectTrue "totals-last", ResultsRow("TOTAL CRR") = ResultsRow("NS3") + 1 And _
                                   ResultsRow("TOTAL BCBS") = ResultsRow("TOTAL CRR") + 1 And _
                                   ResultsRow("TOTAL") = 0, _
                                   "the TOTAL rows are not TOTAL CRR then TOTAL BCBS after the netting sets", _
                                   "illustrative"

        StartCase "totals-single-regime", "N"
        AddSwap "T1"
        TEST_CaseRunner.RunCase
        ExpectTotal RG_CRR, "1 of 1 netting set(s) VALID; 0 withheld", "NS1"
        TEST_CaseRunner.ExpectTrue "one-total", ResultsRow("TOTAL BCBS") = 0 And ResultsRow("TOTAL") = 0, _
                                   "a TOTAL row other than TOTAL CRR was written", "illustrative"

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
'   2026-10-12
'==============================================================================
'
        AddSwapTo tradeId, "2031-09-30"

End Sub


Private Sub AddSwapTo( _
    ByVal tradeId As String, _
    ByVal endDate As String)
'
'==============================================================================
'                                  AddSwapTo
'------------------------------------------------------------------------------
' PURPOSE
'   Add a EUR interest-rate swap to NS1 starting on the reporting date and
'   ending on endDate (yyyy-mm-dd), which is also its maturity (#40).
'
' UPDATED
'   2026-10-12
'==============================================================================
'
        TEST_CaseRunner.AddTrade tradeId, "IR", "", "EUR", "Linear", "Long", "", "Standard", "", _
                            "10000", "30", "", endDate, endDate, "", "", "", "", "", ""

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


Private Sub AddBucketSwaps()
'
'==============================================================================
'                                AddBucketSwaps
'------------------------------------------------------------------------------
' PURPOSE
'   Add four EUR swaps to NS1 ending 364, 365, 1,825 and 1,826 days after
'   the reporting date, at the IR maturity-bucket edges (#40, #191).
'
' UPDATED
'   2026-10-13
'==============================================================================
'
        AddSwapTo "B364", "2027-09-29"
        AddSwapTo "B365", "2027-09-30"
        AddSwapTo "B1825", "2031-09-29"
        AddSwapTo "B1826", "2031-09-30"

End Sub


Private Sub AddOffsetSwaps()
'
'==============================================================================
'                                AddOffsetSwaps
'------------------------------------------------------------------------------
' PURPOSE
'   Add to NS1 a long EUR swap ending exactly 1 year after the reporting
'   date and a short one ending 2 years after it (730 days), notional
'   10,000 each (#191).
'
' UPDATED
'   2026-10-13
'==============================================================================
'
        AddSwapTo "L365", "2027-09-30"
        AddSwapTo "S730", "2028-09-29"
        TEST_CaseRunner.SetInputCell SH_TRADES, 2, TR_DIR, "Short"

End Sub


Private Sub AddSmallOption()
'
'==============================================================================
'                                AddSmallOption
'------------------------------------------------------------------------------
' PURPOSE
'   Add to NS1 a bought one-year single-name equity call, notional 10,000,
'   underlying price 0.012345 and strike 0.0123, values with more than
'   four decimals (#196).
'
' UPDATED
'   2026-10-13
'==============================================================================
'
        TEST_CaseRunner.AddTrade "T1", "EQ", "SINGLE", "ACME", "Option", "Long", "Call", "Standard", "", _
                                 "10000", "0", "", "2027-09-30", "2027-09-30", "2027-09-30", "0.012345", _
                                 "0.0123", "", "", ""

End Sub


Private Sub AddFxForwardOn( _
    ByVal tradeId As String, _
    ByVal pair As String, _
    ByVal direction As String)
'
'==============================================================================
'                                AddFxForwardOn
'------------------------------------------------------------------------------
' PURPOSE
'   Add a five-year FX forward on pair as written, notional 10,000 in the
'   case currency (EUR), MtM 0, to NS1 (#40).
'
' UPDATED
'   2026-10-12
'==============================================================================
'
        TEST_CaseRunner.AddTrade tradeId, "FX", "", pair, "Linear", direction, "", "Standard", "", _
                                 "10000", "0", "", "2031-09-30", "2031-09-30", "", "", "", "", "", ""

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


Private Sub ExpectStopIn1904( _
    ByVal messagePart As String)
'
'==============================================================================
'                               ExpectStopIn1904
'------------------------------------------------------------------------------
' PURPOSE
'   Set the workbook to the 1904 date system, run the case expecting it to
'   stop with messagePart, and set the workbook back to the 1900 system
'   (#40). The case's dates shift by 1,462 days while the setting is on;
'   they return when it is cleared.
'
' ERROR POLICY
'   The setting is cleared before any error is raised again, so that a
'   failure cannot leave the workbook in the 1904 system.
'
' UPDATED
'   2026-10-12
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim errNumber   As Long      'Error raised while the setting was on
    Dim errText     As String    'Its description

'------------------------------------------------------------------------------
' RUN IN THE 1904 SYSTEM
'------------------------------------------------------------------------------
        If ThisWorkbook.Date1904 Then
            Err.Raise ERR_TEST_SETUP, "TEST_InputValidation.ExpectStopIn1904", _
                      "The workbook already uses the 1904 date system."
        End If
        ThisWorkbook.Date1904 = True
        On Error Resume Next
        ExpectStop messagePart
        errNumber = Err.Number
        errText = Err.Description
        On Error GoTo 0
        ThisWorkbook.Date1904 = False
        If errNumber <> 0 Then
            Err.Raise errNumber, "TEST_InputValidation.ExpectStopIn1904", errText
        End If

End Sub


Private Function ResultsRow( _
    ByVal label As String) _
    As Long
'
'==============================================================================
'                                  ResultsRow
'------------------------------------------------------------------------------
' PURPOSE
'   Find a Results row by the exact text of its first column: a netting
'   set's ID or a TOTAL row's label (#198).
'
' RETURNS
'   The sheet row; 0 when no row has that label.
'
' UPDATED
'   2026-10-13
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ws   As Worksheet    'Results sheet
    Dim r    As Long         'Row being read

'------------------------------------------------------------------------------
' FIND
'------------------------------------------------------------------------------
        Set ws = GetSheet(SH_RESULTS)
        r = FIRST_DATA_ROW
        Do While Not IsBlankCell(TableCell(ws, r, 1).Value)
            If SafeStr(TableCell(ws, r, 1).Value) = label Then
                ResultsRow = r
                Exit Function
            End If
            r = r + 1
        Loop

End Function


Private Sub ExpectTotal( _
    ByVal regimeCode As String, _
    ByVal statusText As String, _
    ByVal nettingSetId As String)
'
'==============================================================================
'                                 ExpectTotal
'------------------------------------------------------------------------------
' PURPOSE
'   Check the TOTAL row of a regime (#198): it exists, its status is
'   statusText, its EAD equals the EAD of the regime's one VALID netting
'   set nettingSetId, and CORE_Engine.TotalEAD reports the same figure.
'
' UPDATED
'   2026-10-13
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ws       As Worksheet    'Results sheet
    Dim tr       As Long         'Row of the TOTAL row
    Dim nr       As Long         'Row of the netting set
    Dim ok       As Boolean      'All conditions hold
    Dim detail   As String       'Diagnostic

'------------------------------------------------------------------------------
' CHECK
'------------------------------------------------------------------------------
        Set ws = GetSheet(SH_RESULTS)
        tr = ResultsRow("TOTAL " & regimeCode)
        nr = ResultsRow(nettingSetId)
        If tr = 0 Or nr = 0 Then
            detail = "TOTAL " & regimeCode & " or " & nettingSetId & " is missing from Results"
        Else
            ok = (SafeStr(TableCell(ws, tr, RS_STATUS_COL).Value) = statusText) And _
                 (Abs(ToDbl(TableCell(ws, tr, 22).Value) - ToDbl(TableCell(ws, nr, 22).Value)) < 0.000001) And _
                 (Abs(CORE_Engine.TotalEAD(regimeCode) - ToDbl(TableCell(ws, nr, 22).Value)) < 0.000001)
            detail = "TOTAL " & regimeCode & " status """ & SafeStr(TableCell(ws, tr, RS_STATUS_COL).Value) & _
                     """, EAD " & TableCell(ws, tr, 22).Value & "; expected """ & statusText & """ and the EAD of " & _
                     nettingSetId & ", " & TableCell(ws, nr, 22).Value
        End If
        TEST_CaseRunner.ExpectTrue "total-" & LCase$(regimeCode), ok, detail, "illustrative"

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
