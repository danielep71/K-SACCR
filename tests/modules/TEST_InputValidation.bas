Attribute VB_Name = "TEST_InputValidation"
'==============================================================================
' MODULE: TEST_InputValidation
'------------------------------------------------------------------------------
' PURPOSE
'   Check that malformed inputs are rejected instead of silently changing the
'   trade population, the netting-set terms or the parameters (#35): trade
'   rows with data but no ID, duplicate trade and netting-set IDs, a missing
'   MtM, an unreadable date, netting-set fields that cannot be read, an Excel
'   error value (#N/A) in a trade, netting set or parameter, a moved or
'   renamed column, a broken parameter or layout name, a parameter,
'   supervisory factor or currency listed twice with different values, and a
'   factor or rate the table loader would not read because of a blank row or
'   key; and, for aggregation (#37), a credit reference given two sub-classes
'   in either order, a basis trade without a hedging-set label, a reserved
'   character in a reference, and an interest-rate risk factor that is not a
'   currency; and, for the CRR scope (#33), the sum of absolute IR bucket
'   values in a CRR netting set and an IR risk factor with a suffix other
'   than -INFL; and, for the netting-set terms (#39), VM on an unmargined
'   netting set, a clearing role or margin-agreement structure that is not a
'   code, and the structures the engine does not model, including a netting
'   set declared as facing a CCP (#112). A blank optional field, a duplicate
'   with the same values, an inflation risk factor such as EUR-INFL, the sum
'   of absolutes in a BCBS netting set and VM posted on an unmargined BCBS
'   netting set must still be accepted. The same section checks the MPOR of
'   each clearing role and regime, also with a large flag, disputes or an
'   MPOR override (#359), and with disputes or a large netting set and weekly
'   remargining, against the rules recorded in the methodology, and the
'   warning for a margined netting set of more than 5,000 trades whose large
'   or illiquid flag is N. Values outside the ranges of the data validation
'   (#147), written by VBA so that validation does not stop them, must also
'   be rejected: a multiplier floor of 1, a correlation above 1, correlations
'   that do not form a correlation matrix, an MPOR floor of 0, a fractional
'   business-day count, a negative supervisory factor, a negative threshold
'   or MTA, a remargining frequency or MPOR override that is not a whole
'   number of business days in range (0.5 and 10001 rejected, 10000
'   accepted), a reporting currency that is not three letters A to Z, and an
'   alpha override at or below zero; a remargining frequency of 0 is raised
'   to daily with a warning. A run with the regulatory parameter set is
'   REGULATORY, and one with a different alpha, a zero factor or an alpha
'   override is WHAT-IF with the difference on Checks (#41); a DaysPerYear of
'   360 keeps it regulatory; a reporting-currency FX rate other than 1 stops
'   the run. For the input conventions (#40): a text date is read only as
'   yyyy-mm-dd, a workbook in the 1904 date system stops the run, the IR
'   maturity buckets split at exactly 1 and 5 years with the regime's edge at
'   1 year, which changes the IR add-on of offsetting swaps (#191), a trade
'   or option ending on the reporting date is rejected, an option expiring
'   after its maturity is rejected, both currencies of an FX pair need an FX
'   rate, the notional must be in one of them, and a pair entered the other
'   way round offsets the same pair entered in order. An Excel error value in
'   an optional text field of a trade, in a netting set's regime or in the
'   Params regime is an input error, not a blank that takes a default (#194),
'   and a Currency or Accounting number format does not round an input
'   (#196). A swaption needs the start date of its underlying and a Bermudan
'   one takes its earliest and latest exercise dates (#40), and an FX option
'   must be entered in market quote order (#40). For economic trades (#38):
'   legs sharing an Economic trade ID form one trade with exactly one MtM
'   leg, legs in different risk categories are accepted, legs in one asset
'   class must be options, and an error value in either leg column is
'   rejected. For settled-to-market trades (#360): the next reset date must
'   be after the reporting date and not after the maturity, a reset on the
'   maturity is accepted, blank or N with a reset date, STM without one,
'   another code and a reset that is not a date are rejected, an STM trade
'   in a margined netting set makes it INVALID, and the legs of one
'   economic trade share the settlement type and reset date. Results has
'   one TOTAL row per regime, after the netting sets, with the number of
'   VALID and withheld netting sets it covers, and a netting set may not
'   take the label of a TOTAL row (#198).
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
'   2026-10-20
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
        Private Const CASES       As Long = 163               'Cases in a complete run
        Private Const CHECKS      As Long = 267               'Checks in a complete run


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
'   2026-10-18
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim PatchedRow    As Long       'Params or Trades row a case patches
    Dim i             As Long       'Trade number in the large netting set
    Dim TextFields    As Variant    'Trades column and field name, in pairs
    Dim f             As Long       'Index of a column in TextFields
    Dim SavedFormat   As String     'Number format of a cell a case changes

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
        PatchedRow = UsedLastRow(GetSheet(SH_PARAMS)) + 2
        TEST_CaseRunner.PatchCell SH_PARAMS, PatchedRow, ParamsCodeCol(), PRM_ALPHA
        TEST_CaseRunner.PatchCell SH_PARAMS, PatchedRow, ParamsValueCol(), ParamsValue(PRM_ALPHA, ParamsValueCol())
        ExpectStatus "VALID"

        StartCase "params-duplicate-same-value", "N"
        AddSwap "T1"
        AddParamRow PRM_ALPHA, ParamsValue(PRM_ALPHA, ParamsValueCol())
        ExpectStatus "VALID"

    'CO_OTHER becomes a second CO_METALS row: it differs in category and
    'hedging set, unless those are patched too.
        StartCase "factor-table-duplicate-different-values", "N"
        AddSwap "T1"
        PatchedRow = ParamsRow("CO_OTHER")
        TEST_CaseRunner.PatchCell SH_PARAMS, PatchedRow, TableCol(NM_SF_TABLE, 1), "CO_METALS"
        ExpectStop "Supervisory factor key listed twice with different values"

        StartCase "factor-table-duplicate-same-values", "N"
        AddSwap "T1"
        PatchedRow = ParamsRow("CO_OTHER")
        TEST_CaseRunner.PatchCell SH_PARAMS, PatchedRow, TableCol(NM_SF_TABLE, 1), "CO_METALS"
        TEST_CaseRunner.PatchCell SH_PARAMS, PatchedRow, TableCol(NM_SF_TABLE, 3), "METALS"
        TEST_CaseRunner.PatchCell SH_PARAMS, PatchedRow, TableCol(NM_SF_TABLE, 7), "METALS"
        ExpectStatus "VALID"

    'CHF becomes a second USD row, at CHF's rate or at USD's.
        StartCase "fx-table-duplicate-different-rates", "N"
        AddSwap "T1"
        PatchedRow = ParamsRow("CHF")
        TEST_CaseRunner.PatchCell SH_PARAMS, PatchedRow, TableCol(NM_FX_TABLE, 1), "USD"
        ExpectStop "Currency listed twice with different rates"

        StartCase "fx-table-duplicate-same-rate", "N"
        AddSwap "T1"
        PatchedRow = ParamsRow("CHF")
        TEST_CaseRunner.PatchCell SH_PARAMS, PatchedRow, TableCol(NM_FX_TABLE, 1), "USD"
        TEST_CaseRunner.PatchCell SH_PARAMS, PatchedRow, TableCol(NM_FX_TABLE, 3), _
                                  ParamsValue("USD", TableCol(NM_FX_TABLE, 3))
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
    'CRR; the BCBS client keeps the 10-day floor [CRE54.8, CRE54.14]. Each
    'CM and CLIENT netting set has an INFO line naming its role, floor and
    'rule, also when it is unmargined and no floor applies (#359).
        StartCase "netting-set-clearing-role-not-a-code", "N"
        TEST_CaseRunner.SetInputCell SH_NS, 1, NS_CLEARED, "Y"
        AddSwap "T1"
        ExpectStatus "INVALID: 1 input error(s)"

    'A netting set declared as facing a CCP is rejected, margined or not,
    'rather than calculated with the bilateral floors, and its status names
    'the reason; another invalid field is counted after it (#112).
        StartCase "netting-set-clearing-role-ccp", "N"
        TEST_CaseRunner.SetInputCell SH_NS, 1, NS_CLEARED, CL_CCP
        AddSwap "T1"
        ExpectStatus "INVALID: facing a CCP, not supported"
        TEST_CaseRunner.ExpectChecksLine "error", SEV_ERROR, "facing a CCP is not supported", "illustrative"

        StartMarginedCase "netting-set-clearing-role-ccp-margined", "CRR", CL_CCP, "1", "N", "N"
        TEST_CaseRunner.SetInputCell SH_NS, 1, NS_TH, -1#
        ExpectStatus "INVALID: facing a CCP, not supported; 1 other input error(s)"

    'So is a declared multi-level client structure, whose conditions apply
    'at every level and are not modelled (#359).
        StartCase "netting-set-clearing-role-multi-level", "N"
        TEST_CaseRunner.SetInputCell SH_NS, 1, NS_CLEARED, CL_MULTI
        AddSwap "T1"
        ExpectStatus "INVALID: multi-level client structure, not supported"
        TEST_CaseRunner.ExpectChecksLine "error", SEV_ERROR, "multi-level client structure is not supported", "illustrative"

        StartMarginedCase "netting-set-clearing-member-crr", "CRR", "CM", "1", "N", "N"
        ExpectMPOR 5
        TEST_CaseRunner.ExpectChecksLine "info", SEV_INFO, "Clearing member facing its client: MPOR floor 5 BD", "illustrative"

        StartMarginedCase "netting-set-clearing-member-bcbs", "BCBS", "CM", "1", "N", "N"
        ExpectMPOR 5
        TEST_CaseRunner.ExpectChecksLine "info", SEV_INFO, "Clearing member facing its client: MPOR floor 5 BD", "illustrative"

        StartMarginedCase "netting-set-clearing-client-crr", "CRR", "CLIENT", "1", "N", "N"
        ExpectMPOR 5
        TEST_CaseRunner.ExpectChecksLine "info", SEV_INFO, "Client facing its clearing member under CRR: MPOR floor 5 BD", "illustrative"

        StartMarginedCase "netting-set-clearing-client-bcbs", "BCBS", "CLIENT", "1", "N", "N"
        ExpectMPOR 10
        TEST_CaseRunner.ExpectChecksLine "info", SEV_INFO, "Client facing its clearing member under BCBS: MPOR floor 10 BD", "illustrative"

        StartCase "netting-set-clearing-member-unmargined", "N"
        TEST_CaseRunner.SetInputCell SH_NS, 1, NS_CLEARED, CL_CM
        AddSwap "T1"
        ExpectStatus "VALID"
        TEST_CaseRunner.ExpectChecksLine "info", SEV_INFO, "on an unmargined netting set: no MPOR floor applies", "illustrative"

    'The client-clearing floors combined with the other MPOR rules (#359).
    'One large or illiquid flag sets 20 days for the clearing member in both
    'regimes, a conservative limitation: CRR Art. 304(3)(c) would let it
    'disregard the trade-count floor towards its clients (decision 3 of the
    'clearing coverage). Disputes double the client's floor, 5 under CRR and
    '10 under BCBS [Art. 285(4); CRE52.51(3)]. An override below the
    'clearing member's floor is replaced by the floor with a warning; one
    'above it is applied.
        StartMarginedCase "netting-set-clearing-member-large-crr", "CRR", "CM", "1", "Y", "N"
        ExpectMPOR 20
        TEST_CaseRunner.ExpectChecksLine "info", SEV_INFO, "MPOR floor 5 BD (CRE54.12; CRR Art. 279c(1)), raised to 20 BD", "illustrative"

        StartMarginedCase "netting-set-clearing-member-large-bcbs", "BCBS", "CM", "1", "Y", "N"
        ExpectMPOR 20
        TEST_CaseRunner.ExpectChecksLine "info", SEV_INFO, "MPOR floor 5 BD (CRE54.12; CRR Art. 279c(1)), raised to 20 BD", "illustrative"

        StartMarginedCase "netting-set-clearing-client-disputes-crr", "CRR", "CLIENT", "1", "N", "Y"
        ExpectMPOR 10

        StartMarginedCase "netting-set-clearing-client-disputes-bcbs", "BCBS", "CLIENT", "1", "N", "Y"
        ExpectMPOR 20

        StartMarginedCase "netting-set-clearing-member-override-below-floor", "CRR", "CM", "1", "N", "N"
        TEST_CaseRunner.SetInputCell SH_NS, 1, NS_MPOR, 3#
        ExpectMPOR 5
        TEST_CaseRunner.ExpectChecksLine "warning", SEV_WARN, "below the regulatory floor", "illustrative"

        StartMarginedCase "netting-set-clearing-member-override-above-floor", "CRR", "CM", "1", "N", "N"
        TEST_CaseRunner.SetInputCell SH_NS, 1, NS_MPOR, 7#
        ExpectMPOR 7

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

    'Netting-set terms outside their range (#147): a negative TH or MTA, a
    'negative or fractional remargining frequency or MPOR override and an
    'alpha override of zero make the set INVALID; a frequency of 0 is
    'raised to daily (1 business day) with a warning.
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

    'Settled-to-market trades (#360). The next reset date must be after the
    'reporting date and not after the maturity as entered; a reset on the
    'maturity date is accepted. Blank or N with a reset date, STM without
    'one, another code and a reset that is not a date are rejected. Lower
    'case stm is read as STM; a daily reset gives M = 1/365, floored at 10
    'business days in the maturity factor: sqrt(10/250) = 0.2.
        StartCase "stm-reset-on-reporting-date", "N"
        AddEquityForward "T1", "0"
        SetSettlement 1, "STM", DateSerial(2026, 9, 30)
        ExpectStatus "INCOMPLETE: 1 of 1 trade(s) rejected"
        TEST_CaseRunner.ExpectChecksLine "reset-on-date", SEV_ERROR, _
                                         "next reset date must be after the reporting date", "illustrative"

        StartCase "stm-reset-before-reporting-date", "N"
        AddEquityForward "T1", "0"
        SetSettlement 1, "STM", DateSerial(2026, 9, 29)
        ExpectStatus "INCOMPLETE: 1 of 1 trade(s) rejected"

        StartCase "stm-reset-after-maturity", "N"
        AddEquityForward "T1", "0"
        SetSettlement 1, "STM", DateSerial(2031, 10, 1)
        ExpectStatus "INCOMPLETE: 1 of 1 trade(s) rejected"
        TEST_CaseRunner.ExpectChecksLine "reset-after-maturity", SEV_ERROR, _
                                         "next reset date is after the maturity date", "illustrative"

        StartCase "stm-reset-on-maturity", "N"
        AddEquityForward "T1", "0"
        SetSettlement 1, "STM", DateSerial(2031, 9, 30)
        ExpectStatus "VALID"

    'With M blank, the reset is compared with the end date that fills it.
        StartCase "stm-reset-after-end-date", "N"
        AddEquityForward "T1", "0"
        TEST_CaseRunner.SetInputCell SH_TRADES, 1, TR_MAT, Empty
        SetSettlement 1, "STM", DateSerial(2031, 10, 1)
        ExpectStatus "INCOMPLETE: 1 of 1 trade(s) rejected"

        StartCase "stm-without-reset", "N"
        AddEquityForward "T1", "0"
        SetSettlement 1, "STM", Empty
        ExpectStatus "INCOMPLETE: 1 of 1 trade(s) rejected"
        TEST_CaseRunner.ExpectChecksLine "no-reset", SEV_ERROR, "needs its next reset date", "illustrative"

        StartCase "settlement-blank-with-reset", "N"
        AddEquityForward "T1", "0"
        SetSettlement 1, Empty, DateSerial(2026, 10, 30)
        ExpectStatus "INCOMPLETE: 1 of 1 trade(s) rejected"
        TEST_CaseRunner.ExpectChecksLine "blank-with-reset", SEV_ERROR, _
                                         "a next reset date is given but the settlement type is not STM", _
                                         "illustrative"

        StartCase "settlement-n-with-reset", "N"
        AddEquityForward "T1", "0"
        SetSettlement 1, ST_NONE, DateSerial(2026, 10, 30)
        ExpectStatus "INCOMPLETE: 1 of 1 trade(s) rejected"

        StartCase "settlement-type-not-a-code", "N"
        AddEquityForward "T1", "0"
        SetSettlement 1, "X", Empty
        ExpectStatus "INCOMPLETE: 1 of 1 trade(s) rejected"
        TEST_CaseRunner.ExpectChecksLine "not-a-code", SEV_ERROR, "settlement type must be N, STM or blank", _
                                         "illustrative"

        StartCase "settlement-type-error", "N"
        AddEquityForward "T1", "0"
        SetSettlement 1, CVErr(xlErrNA), DateSerial(2026, 10, 1)
        ExpectStatus "INCOMPLETE: 1 of 1 trade(s) rejected"
        TEST_CaseRunner.ExpectChecksLine "type-error", SEV_ERROR, _
                                         "Settlement type must be text or blank (found an error value)", "illustrative"

        StartCase "stm-reset-not-a-date", "N"
        AddEquityForward "T1", "0"
        SetSettlement 1, "STM", "soon"
        ExpectStatus "INCOMPLETE: 1 of 1 trade(s) rejected"
        TEST_CaseRunner.ExpectChecksLine "reset-not-a-date", SEV_ERROR, "next reset date is not a date", _
                                         "illustrative"

        StartCase "stm-daily-lower-case", "N"
        AddEquityForward "T1", "0"
        SetSettlement 1, "stm", DateSerial(2026, 10, 1)
        ExpectStatus "VALID"
        TEST_CaseRunner.ExpectNumber "remaining-maturity", "T1", "remaining_maturity", "0.0027397260273972603", _
                                     "0.000000000001", "0", "illustrative"
        TEST_CaseRunner.ExpectNumber "maturity-factor", "T1", "maturity_factor", "0.2", "0.000000000001", "0", _
                                     "illustrative"

    'A settled trade is unmargined (Basel FAQ to CRE52.31-32): in a margined
    'netting set, whether all its trades are STM (an inconsistent
    'classification) or only some (a mixed set, whose CRE52.74 treatment is
    'not implemented), the netting set is INVALID, and another invalid field
    'is counted after the reason.
        StartMarginedCase "stm-margined-all-crr", "CRR", "N", "1", "N", "N"
        SetSettlement 1, "STM", DateSerial(2026, 10, 1)
        ExpectStatus "INVALID: STM requires unmargined treatment; mixed margined/unmargined netting sets are " & _
                     "not supported"
        TEST_CaseRunner.ExpectChecksLine "margined-stm", SEV_ERROR, "STM requires unmargined treatment", _
                                         "illustrative"

        StartMarginedCase "stm-margined-all-bcbs", "BCBS", "N", "1", "N", "N"
        SetSettlement 1, "STM", DateSerial(2026, 10, 1)
        ExpectStatus "INVALID: STM requires unmargined treatment; mixed margined/unmargined netting sets are " & _
                     "not supported"

        StartMarginedCase "stm-margined-mixed", "CRR", "N", "1", "N", "N"
        AddSwap "T2"
        SetSettlement 1, "STM", DateSerial(2026, 10, 1)
        ExpectStatus "INVALID: STM requires unmargined treatment; mixed margined/unmargined netting sets are " & _
                     "not supported"

        StartMarginedCase "stm-margined-other-error", "CRR", "N", "1", "N", "N"
        TEST_CaseRunner.SetInputCell SH_NS, 1, NS_TH, -1#
        SetSettlement 1, "STM", DateSerial(2026, 10, 1)
        ExpectStatus "INVALID: STM requires unmargined treatment; mixed margined/unmargined netting sets are " & _
                     "not supported; 1 other input error(s)"

    'The legs of one economic trade share its settlement arrangement, blank
    'and N alike; separately contracted trades may differ.
        StartCase "legs-settlement-blank-and-n", "N"
        AddEquityCall "L1", "Long", "100", "30"
        AddEquityCall "L2", "Short", "110", "0"
        SetLeg 1, "E1", "Y"
        SetLeg 2, "E1", "N"
        SetSettlement 2, ST_NONE, Empty
        ExpectStatus "VALID"

        StartCase "legs-settlement-differ", "N"
        AddEquityCall "L1", "Long", "100", "30"
        AddEquityCall "L2", "Short", "110", "0"
        SetLeg 1, "E1", "Y"
        SetLeg 2, "E1", "N"
        SetSettlement 1, "STM", DateSerial(2026, 10, 30)
        ExpectStatus "INCOMPLETE: 2 of 2 trade(s) rejected"
        TEST_CaseRunner.ExpectChecksLine "settlement-differ", SEV_ERROR, "its legs have different settlement types", _
                                         "illustrative"

        StartCase "legs-reset-differ", "N"
        AddEquityCall "L1", "Long", "100", "30"
        AddEquityCall "L2", "Short", "110", "0"
        SetLeg 1, "E1", "Y"
        SetLeg 2, "E1", "N"
        SetSettlement 1, "STM", DateSerial(2026, 10, 30)
        SetSettlement 2, "STM", DateSerial(2026, 11, 30)
        ExpectStatus "INCOMPLETE: 2 of 2 trade(s) rejected"
        TEST_CaseRunner.ExpectChecksLine "reset-differ", SEV_ERROR, "its legs have different next reset dates", _
                                         "illustrative"

        StartCase "separate-trades-settlement-differ", "N"
        AddEquityForward "T1", "0"
        AddSwap "T2"
        SetSettlement 1, "STM", DateSerial(2026, 10, 1)
        ExpectStatus "VALID"

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

    'Option dates of a swaption (#40): S is the start of the underlying, so
    'an IR option without a start date is rejected instead of taking S = 0.
    'A Bermudan swaption exercisable yearly from 2027-09-30 to 2029-09-30
    'into a swap ending 2036-09-30 takes S = 1 (the earliest exercise date)
    'for the duration and T = 3.0027 (the latest) for the delta
    '[CRE52.42(1)-(2)]: d = 10000 * (exp(-0.05) - exp(-0.05 * 10.0082)) /
    '0.05 = 68989.59, not the 78743.71 that S = 0 would give, and with P =
    'K the delta is N(0.5 * 0.5 * sqrt(3.0027)) = 0.66757.
        StartCase "ir-option-start-required", "N"
        AddSwaption "T1", ""
        ExpectStatus "INCOMPLETE: 1 of 1 trade(s) rejected"
        TEST_CaseRunner.ExpectChecksLine "start-required", SEV_ERROR, _
                                         "option needs the start date S of its underlying", "illustrative"

        StartCase "ir-option-bermudan-dates", "N"
        AddSwaption "T1", "2027-09-30"
        ExpectStatus "VALID"
        TEST_CaseRunner.ExpectNumber "bermudan-duration", "T1", "adjusted_notional", "68989.5945505363", _
                                     "0.000001", "0.000000001", "illustrative"
        TEST_CaseRunner.ExpectNumber "bermudan-delta", "T1", "supervisory_delta", "0.667569030490699", _
                                     "0.000001", "0", "illustrative"

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

    'FX option quote order (#40): an FX option must be entered in market
    'quote order (FX_QUOTE_ORDER), because a call on EUR/USD and the put on
    'USD/EUR give different deltas (CRE52.59 FAQ1). A one-year EUR/USD call
    'at the money has delta N(0.5 * 0.15) = 0.529893. A USD/JPY call, in
    'market order but not in alphabetical order, sits in the JPY/USD
    'hedging set with its delta reversed: -N((ln(148.5 / 150) + 0.5 *
    '0.15^2) / 0.15) = -0.503191.
        StartCase "fx-option-reversed-pair", "N"
        AddFxOption "T1", "USD/EUR", "1.1", "1.1", "EUR"
        ExpectStatus "INCOMPLETE: 1 of 1 trade(s) rejected"
        TEST_CaseRunner.ExpectChecksLine "quote-order", SEV_ERROR, _
                                         "FX option must be entered in market quote order as EUR/USD", _
                                         "illustrative"

        StartCase "fx-option-market-order", "N"
        AddFxOption "T1", "EUR/USD", "1.1", "1.1", "EUR"
        ExpectStatus "VALID"
        TEST_CaseRunner.ExpectNumber "delta", "T1", "supervisory_delta", "0.529892644052895", _
                                     "0.000001", "0", "illustrative"

        StartCase "fx-option-market-order-not-alphabetical", "N"
        AddFxOption "T1", "USD/JPY", "148.5", "150", "USD"
        ExpectStatus "VALID"
        TEST_CaseRunner.ExpectText "hedging-set", "T1", "hedging_set", "FX|JPY/USD|STD", "illustrative"
        TEST_CaseRunner.ExpectNumber "delta", "T1", "supervisory_delta", "-0.503190610987961", _
                                     "0.000001", "0", "illustrative"

    'An Excel error value, such as #N/A from a broken lookup, in an optional
    'text field is an input error, not a blank that takes the field's
    'default (#194): Instrument would become Linear, Nature Standard, the
    'notional currency the reporting currency, the MtM currency the
    'notional currency, and a netting-set or Params regime the default.
        TextFields = Array(TR_INSTR, "Instrument", TR_NATURE, "Nature", TR_SUB, "Sub-class", _
                       TR_LABEL, "Hedging-set label", TR_NCCY, "Notional currency", TR_MCCY, "MtM currency")
        For f = LBound(TextFields) To UBound(TextFields) Step 2
            StartCase "text-error-" & LCase$(Replace(CStr(TextFields(f + 1)), " ", "-")), "N"
            AddSwap "T1"
            TEST_CaseRunner.SetInputCell SH_TRADES, 1, CLng(TextFields(f)), CVErr(xlErrNA)
            ExpectStatus "INCOMPLETE: 1 of 1 trade(s) rejected"
            TEST_CaseRunner.ExpectChecksLine "error-value", SEV_ERROR, _
                                             CStr(TextFields(f + 1)) & " must be text or blank (found an error value)", _
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
    '0.0123 the price would give 0.72575. General, Currency and Accounting
    'must all give 0.72676. The Currency and Accounting cases also check
    'that Range.Value returns the VBA Currency type for the formatted cell,
    'so that each is a format that would have rounded the input. The cell
    'formats are put back at the end of each case.
        StartCase "number-format-general", "N"
        AddSmallOption
        ExpectStatus "VALID"
        TEST_CaseRunner.ExpectNumber "delta", "T1", "supervisory_delta", "0.726760030216713", _
                                     "0.000001", "0", "illustrative"

        StartCase "number-format-currency", "N"
        AddSmallOption
        PatchedRow = TableFirstRow(SH_TRADES)
        SavedFormat = TableCell(GetSheet(SH_TRADES), PatchedRow, TR_PRICE).NumberFormat
        TableCell(GetSheet(SH_TRADES), PatchedRow, TR_PRICE).NumberFormat = "$#,##0.00"
        ExpectStatus "VALID"
        TEST_CaseRunner.ExpectNumber "delta", "T1", "supervisory_delta", "0.726760030216713", _
                                     "0.000001", "0", "illustrative"
        TEST_CaseRunner.ExpectTrue "value-type", _
                                   VarType(TableCell(GetSheet(SH_TRADES), PatchedRow, TR_PRICE).Value) = vbCurrency, _
                                   "the Currency format did not make Range.Value return Currency", "illustrative"
        TableCell(GetSheet(SH_TRADES), PatchedRow, TR_PRICE).NumberFormat = SavedFormat

        StartCase "number-format-accounting", "N"
        AddSmallOption
        PatchedRow = TableFirstRow(SH_TRADES)
        SavedFormat = TableCell(GetSheet(SH_TRADES), PatchedRow, TR_PRICE).NumberFormat
        TableCell(GetSheet(SH_TRADES), PatchedRow, TR_PRICE).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
        ExpectStatus "VALID"
        TEST_CaseRunner.ExpectNumber "delta", "T1", "supervisory_delta", "0.726760030216713", _
                                     "0.000001", "0", "illustrative"
        TEST_CaseRunner.ExpectTrue "value-type", _
                                   VarType(TableCell(GetSheet(SH_TRADES), PatchedRow, TR_PRICE).Value) = vbCurrency, _
                                   "the Accounting format did not make Range.Value return Currency", "illustrative"
        TableCell(GetSheet(SH_TRADES), PatchedRow, TR_PRICE).NumberFormat = SavedFormat

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

    'A netting set may not take the label of a TOTAL row: its row is
    'ignored with an error, and TOTAL CRR still covers NS1 alone.
        StartCase "netting-set-reserved-id", "N"
        TEST_CaseRunner.SetInputCell SH_NS, 2, NS_ID, "total crr"
        TEST_CaseRunner.SetInputCell SH_NS, 2, NS_REGIME, "CRR"
        AddSwap "T1"
        TEST_CaseRunner.RunCase
        TEST_CaseRunner.ExpectChecksLine "reserved-id", SEV_ERROR, "is reserved for a TOTAL row", "illustrative"
        ExpectTotal RG_CRR, "1 of 1 netting set(s) VALID; 0 withheld", "NS1"

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
    ByVal CaseId As String, _
    ByVal FlagValue As String)
'
'==============================================================================
'                                  StartCase
'------------------------------------------------------------------------------
' PURPOSE
'   Start a CRR case with one unmargined netting set NS1. FlagValue is
'   written to the margined, clearing-role, large-or-illiquid and disputes
'   fields: "N", or "" to leave them blank.
'
' INPUTS
'   CaseId: the case ID; the log prints it behind "input-", which
'   tools/check_excel_evidence.py reads from the BeginCase line below.
'
' UPDATED
'   2026-10-17
'==============================================================================
'
        TEST_CaseRunner.BeginCase "input-" & CaseId, "CRR", VALUATION, "EUR"
        TEST_CaseRunner.AddNettingSet "NS1", FlagValue, FlagValue, "", FlagValue, FlagValue, "", _
                                 "0", "0", "0", "0", ""

End Sub


Private Sub StartMarginedCase( _
    ByVal CaseId As String, _
    ByVal RegimeCode As String, _
    ByVal RoleCode As String, _
    ByVal RemarginDays As String, _
    ByVal LargeFlag As String, _
    ByVal DisputesFlag As String)
'
'==============================================================================
'                              StartMarginedCase
'------------------------------------------------------------------------------
' PURPOSE
'   Start a case with one margined netting set NS1, without collateral,
'   threshold or MTA, in RegimeCode, with the clearing role, remargining
'   frequency and large-or-illiquid and disputes flags given, and add a
'   swap.
'
' UPDATED
'   2026-10-17
'==============================================================================
'
        TEST_CaseRunner.BeginCase "input-" & CaseId, RegimeCode, VALUATION, "EUR"
        TEST_CaseRunner.AddNettingSet "NS1", "Y", RoleCode, RemarginDays, LargeFlag, DisputesFlag, "", _
                                 "0", "0", "0", "0", ""
        AddSwap "T1"

End Sub


Private Sub ExpectMPOR( _
    ByVal ExpectedDays As Long)
'
'==============================================================================
'                                  ExpectMPOR
'------------------------------------------------------------------------------
' PURPOSE
'   Run the case and check that the netting set is VALID with the expected
'   MPOR in business days.
'
' UPDATED
'   2026-10-17
'==============================================================================
'
        ExpectStatus "VALID"
        TEST_CaseRunner.ExpectNumber "mpor", "", "margin_period_of_risk", CStr(ExpectedDays), "0", "0", _
                                     "illustrative"

End Sub


Private Sub AddSwap( _
    ByVal TradeId As String)
'
'==============================================================================
'                                   AddSwap
'------------------------------------------------------------------------------
' PURPOSE
'   Add a valid five-year EUR interest-rate swap to NS1.
'
' UPDATED
'   2026-10-17
'==============================================================================
'
        AddSwapTo TradeId, "2031-09-30"

End Sub


Private Sub AddSwapTo( _
    ByVal TradeId As String, _
    ByVal EndDate As String)
'
'==============================================================================
'                                  AddSwapTo
'------------------------------------------------------------------------------
' PURPOSE
'   Add a EUR interest-rate swap to NS1 starting on the reporting date and
'   ending on EndDate (yyyy-mm-dd), which is also its maturity (#40).
'
' UPDATED
'   2026-10-17
'==============================================================================
'
        TEST_CaseRunner.AddTrade TradeId, "IR", "", "EUR", "Linear", "Long", "", "Standard", "", _
                            "10000", "30", "", EndDate, EndDate, "", "", "", "", "", ""

End Sub


Private Sub ExpectRunClass( _
    ByVal ExpectedClass As String)
'
'==============================================================================
'                                ExpectRunClass
'------------------------------------------------------------------------------
' PURPOSE
'   Check the class of the case's run: RUN_REGULATORY or RUN_WHATIF (#41).
'
' UPDATED
'   2026-10-17
'==============================================================================
'
        TEST_CaseRunner.ExpectTrue "run-class", CORE_Engine.RunClass = ExpectedClass, _
                                   "run class " & CORE_Engine.RunClass & ", expected " & ExpectedClass, "illustrative"

End Sub


Private Sub StartReportingCcyCase( _
    ByVal CaseId As String, _
    ByVal CurrencyCode As String)
'
'==============================================================================
'                            StartReportingCcyCase
'------------------------------------------------------------------------------
' PURPOSE
'   Start a CRR case with one swap and the reporting currency written as
'   CurrencyCode, bypassing the worksheet validation.
'
' UPDATED
'   2026-10-17
'==============================================================================
'
        StartCase CaseId, "N"
        AddSwap "T1"
        TEST_CaseRunner.PatchCell SH_PARAMS, ParamsRow(PRM_REPCCY), ParamsValueCol(), CurrencyCode

End Sub


Private Sub ExpectStop( _
    ByVal MessagePart As String)
'
'==============================================================================
'                                  ExpectStop
'------------------------------------------------------------------------------
' PURPOSE
'   Run the case and check that the run stopped with an ERROR on Checks
'   containing MessagePart.
'
' UPDATED
'   2026-10-17
'==============================================================================
'
        TEST_CaseRunner.RunCaseExpectingStop "stop", MessagePart, "illustrative"

End Sub


Private Function ParamsRow( _
    ByVal ParamsKey As String) _
    As Long
'
'==============================================================================
'                                  ParamsRow
'------------------------------------------------------------------------------
' PURPOSE
'   Return the Params row that holds ParamsKey: a parameter code in the code
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
    Dim FoundRow        As Long       'Row found; 0 when absent
    Dim SearchColumns   As Variant    'Columns searched, as table columns
    Dim i               As Long       'Index into SearchColumns

'------------------------------------------------------------------------------
' FIND
'------------------------------------------------------------------------------
        SearchColumns = Array(ParamsCodeCol(), TableCol(NM_SF_TABLE, 1), TableCol(NM_FX_TABLE, 1))
        For i = LBound(SearchColumns) To UBound(SearchColumns)
            FoundRow = FindHeaderRow(GetSheet(SH_PARAMS), CLng(SearchColumns(i)), ParamsKey)
            If FoundRow > 0 Then
                Exit For
            End If
        Next i
        If FoundRow = 0 Then
            Err.Raise ERR_TEST_SETUP, "TEST_InputValidation.ParamsRow", "'" & ParamsKey & "' not found on Params."
        End If
        ParamsRow = FoundRow

End Function


Private Function ParamsValue( _
    ByVal ParamsKey As String, _
    ByVal ColumnNumber As Long) _
    As Variant
'
'==============================================================================
'                                 ParamsValue
'------------------------------------------------------------------------------
' PURPOSE
'   Return a value from the Params row that holds ParamsKey (see
'   ParamsRow), at table column ColumnNumber.
'
' UPDATED
'   2026-10-11
'==============================================================================
'
        ParamsValue = TableCell(GetSheet(SH_PARAMS), ParamsRow(ParamsKey), ColumnNumber).Value

End Function


Private Sub AddParamRow( _
    ByVal ParamCode As String, _
    ByVal NewValue As Variant)
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
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim NewRow   As Long    'The new row

'------------------------------------------------------------------------------
' WRITE
'------------------------------------------------------------------------------
        NewRow = ParamsRow(PRM_LAMCO) + 1
        If Not IsBlankCell(TableCell(GetSheet(SH_PARAMS), NewRow, ParamsCodeCol()).Value) Or _
           Not IsBlankCell(TableCell(GetSheet(SH_PARAMS), NewRow, ParamsValueCol()).Value) Then
            Err.Raise ERR_TEST_SETUP, "TEST_InputValidation.AddParamRow", "Params row " & NewRow & " is not blank."
        End If
        TEST_CaseRunner.PatchCell SH_PARAMS, NewRow, ParamsCodeCol(), ParamCode
        TEST_CaseRunner.PatchCell SH_PARAMS, NewRow, ParamsValueCol(), NewValue

End Sub


Private Sub ClearParamsRow( _
    ByVal RowNumber As Long, _
    ByVal FirstColumn As Long, _
    ByVal ColumnCount As Long)
'
'==============================================================================
'                                ClearParamsRow
'------------------------------------------------------------------------------
' PURPOSE
'   Blank ColumnCount table columns from FirstColumn of one Params row, as
'   if a blank row had been inserted in a table there. Restored with the
'   case's other patches.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim c   As Long    'Column being cleared

'------------------------------------------------------------------------------
' CLEAR
'------------------------------------------------------------------------------
        For c = FirstColumn To FirstColumn + ColumnCount - 1
            TEST_CaseRunner.PatchCell SH_PARAMS, RowNumber, c, Empty
        Next c

End Sub


Private Function TableCol( _
    ByVal NameText As String, _
    ByVal ColumnIndex As Long) _
    As Long
'
'==============================================================================
'                                   TableCol
'------------------------------------------------------------------------------
' PURPOSE
'   Return the table column of column ColumnIndex of a Params lookup
'   table, counted from the top-left cell of its workbook name (1 for the
'   keys).
'
' UPDATED
'   2026-10-11
'==============================================================================
'
        TableCol = NamedTableCol(NameText) + ColumnIndex - 1

End Function


Private Sub AddEquityForward( _
    ByVal TradeId As String, _
    ByVal MarketValue As String)
'
'==============================================================================
'                               AddEquityForward
'------------------------------------------------------------------------------
' PURPOSE
'   Add a five-year single-name equity forward, long, notional 10,000, to
'   NS1 (#38).
'
' UPDATED
'   2026-10-17
'==============================================================================
'
        TEST_CaseRunner.AddTrade TradeId, "EQ", "SINGLE", "ACME", "Linear", "Long", "", "Standard", "", _
                                 "10000", MarketValue, "", "2031-09-30", "2031-09-30", "", "", "", "", "", ""

End Sub


Private Sub AddFxForward( _
    ByVal TradeId As String, _
    ByVal MarketValue As String)
'
'==============================================================================
'                                AddFxForward
'------------------------------------------------------------------------------
' PURPOSE
'   Add a five-year EUR/USD forward, long, notional 10,000, to NS1 (#38).
'
' UPDATED
'   2026-10-17
'==============================================================================
'
        TEST_CaseRunner.AddTrade TradeId, "FX", "", "EUR/USD", "Linear", "Long", "", "Standard", "", _
                                 "10000", MarketValue, "", "2031-09-30", "2031-09-30", "", "", "", "", "", ""

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


Private Sub AddSwaption( _
    ByVal TradeId As String, _
    ByVal StartDate As String)
'
'==============================================================================
'                                 AddSwaption
'------------------------------------------------------------------------------
' PURPOSE
'   Add a bought EUR payer swaption on 10,000, struck at the money (P = K
'   = 5%), with its latest exercise date on 2029-09-30 and the underlying
'   swap ending on 2036-09-30; StartDate is the start of the underlying,
'   "" for none (#40).
'
' UPDATED
'   2026-10-17
'==============================================================================
'
        TEST_CaseRunner.AddTrade TradeId, "IR", "", "EUR", "Option", "Long", "Call", "Standard", "", _
                                 "10000", "0", StartDate, "2036-09-30", "2036-09-30", "2029-09-30", "0.05", _
                                 "0.05", "", "", ""

End Sub


Private Sub AddFxOption( _
    ByVal TradeId As String, _
    ByVal Pair As String, _
    ByVal Price As String, _
    ByVal Strike As String, _
    ByVal NotionalCurrency As String)
'
'==============================================================================
'                                 AddFxOption
'------------------------------------------------------------------------------
' PURPOSE
'   Add a bought FX call on 10,000 of NotionalCurrency on Pair as given,
'   expiring and maturing on 2027-09-30, one year after the valuation
'   date (#40). The notional currency is written on Trades row 1, so the
'   option must be the case's first trade.
'
' UPDATED
'   2026-10-17
'==============================================================================
'
        TEST_CaseRunner.AddTrade TradeId, "FX", "", Pair, "Option", "Long", "Call", "Standard", "", _
                                 "10000", "0", "", "", "2027-09-30", "2027-09-30", Price, Strike, "", "", ""
        TEST_CaseRunner.SetInputCell SH_TRADES, 1, TR_NCCY, NotionalCurrency

End Sub


Private Sub AddFxForwardOn( _
    ByVal TradeId As String, _
    ByVal Pair As String, _
    ByVal Direction As String)
'
'==============================================================================
'                                AddFxForwardOn
'------------------------------------------------------------------------------
' PURPOSE
'   Add a five-year FX forward on Pair as written, notional 10,000 in the
'   case currency (EUR), MtM 0, to NS1 (#40).
'
' UPDATED
'   2026-10-17
'==============================================================================
'
        TEST_CaseRunner.AddTrade TradeId, "FX", "", Pair, "Linear", Direction, "", "Standard", "", _
                                 "10000", "0", "", "2031-09-30", "2031-09-30", "", "", "", "", "", ""

End Sub


Private Sub AddEquityCall( _
    ByVal TradeId As String, _
    ByVal Direction As String, _
    ByVal StrikePrice As String, _
    ByVal MarketValue As String)
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
'   2026-10-17
'==============================================================================
'
        TEST_CaseRunner.AddTrade TradeId, "EQ", "SINGLE", "ACME", "Option", Direction, "Call", "Standard", "", _
                                 "10000", MarketValue, "", "2027-09-30", "2027-09-30", "2027-09-30", "100", _
                                 StrikePrice, "", "", ""

End Sub


Private Sub SetLeg( _
    ByVal DataRow As Long, _
    ByVal EconomicTradeId As String, _
    ByVal MtmLegFlag As String)
'
'==============================================================================
'                                    SetLeg
'------------------------------------------------------------------------------
' PURPOSE
'   Make a trade row a leg: its Economic trade ID and its MtM leg flag (#38).
'
' UPDATED
'   2026-10-17
'==============================================================================
'
        TEST_CaseRunner.SetInputCell SH_TRADES, DataRow, TR_ECON, EconomicTradeId
        TEST_CaseRunner.SetInputCell SH_TRADES, DataRow, TR_MTMLEG, MtmLegFlag

End Sub


Private Sub SetSettlement( _
    ByVal DataRow As Long, _
    ByVal SettlementType As Variant, _
    ByVal NextResetDate As Variant)
'
'==============================================================================
'                                SetSettlement
'------------------------------------------------------------------------------
' PURPOSE
'   Write a trade row's settlement type and next reset date (#360). Empty
'   leaves a cell blank; a date is passed as a Date, so that Excel does not
'   read text in the regional format.
'
' UPDATED
'   2026-10-20
'==============================================================================
'
        TEST_CaseRunner.SetInputCell SH_TRADES, DataRow, TR_SETTLE, SettlementType
        TEST_CaseRunner.SetInputCell SH_TRADES, DataRow, TR_RESET, NextResetDate

End Sub


Private Sub AddCredit( _
    ByVal TradeId As String, _
    ByVal SubClass As String, _
    ByVal ReferenceName As String)
'
'==============================================================================
'                                  AddCredit
'------------------------------------------------------------------------------
' PURPOSE
'   Add a five-year single-name credit default swap, protection bought, to
'   NS1.
'
' UPDATED
'   2026-10-17
'==============================================================================
'
        TEST_CaseRunner.AddTrade TradeId, "CR", SubClass, ReferenceName, "Linear", "Long", "", "Standard", "", _
                                 "10000", "0", "", "2031-09-30", "2031-09-30", "", "", "", "", "", ""

End Sub


Private Sub ExpectStopIn1904( _
    ByVal MessagePart As String)
'
'==============================================================================
'                               ExpectStopIn1904
'------------------------------------------------------------------------------
' PURPOSE
'   Set the workbook to the 1904 date system, run the case expecting it to
'   stop with MessagePart, and set the workbook back to the 1900 system
'   (#40). The case's dates shift by 1,462 days while the setting is on;
'   they return when it is cleared.
'
' ERROR POLICY
'   The setting is cleared before any error is raised again, so that a
'   failure cannot leave the workbook in the 1904 system.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ErrorNumber   As Long      'Error raised while the setting was on
    Dim ErrorText     As String    'Its description

'------------------------------------------------------------------------------
' RUN IN THE 1904 SYSTEM
'------------------------------------------------------------------------------
        If ThisWorkbook.Date1904 Then
            Err.Raise ERR_TEST_SETUP, "TEST_InputValidation.ExpectStopIn1904", _
                      "The workbook already uses the 1904 date system."
        End If
        ThisWorkbook.Date1904 = True
        On Error Resume Next
        ExpectStop MessagePart
        ErrorNumber = Err.Number
        ErrorText = Err.Description
        On Error GoTo 0
        ThisWorkbook.Date1904 = False
        If ErrorNumber <> 0 Then
            Err.Raise ErrorNumber, "TEST_InputValidation.ExpectStopIn1904", ErrorText
        End If

End Sub


Private Function ResultsRow( _
    ByVal FirstColumnText As String) _
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
    Dim ResultsSheet   As Worksheet    'Results sheet
    Dim r              As Long         'Row being read

'------------------------------------------------------------------------------
' FIND
'------------------------------------------------------------------------------
        Set ResultsSheet = GetSheet(SH_RESULTS)
        r = OutputFirstRow(SH_RESULTS)
        Do While Not IsBlankCell(TableCell(ResultsSheet, r, 1).Value)
            If SafeStr(TableCell(ResultsSheet, r, 1).Value) = FirstColumnText Then
                ResultsRow = r
                Exit Function
            End If
            r = r + 1
        Loop

End Function


Private Sub ExpectTotal( _
    ByVal RegimeCode As String, _
    ByVal StatusText As String, _
    ByVal ExpectedNettingSet As String)
'
'==============================================================================
'                                 ExpectTotal
'------------------------------------------------------------------------------
' PURPOSE
'   Check the TOTAL row of a regime (#198): it exists, its status is
'   StatusText, its EAD (column 22) equals the EAD of the regime's one
'   VALID netting set ExpectedNettingSet, and CORE_Engine.TotalEAD reports
'   the same figure.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ResultsSheet    As Worksheet    'Results sheet
    Dim TotalRow        As Long         'Row of the TOTAL row
    Dim NettingSetRow   As Long         'Row of the netting set
    Dim AllHold         As Boolean      'All conditions hold
    Dim Detail          As String       'Diagnostic

'------------------------------------------------------------------------------
' CHECK
'------------------------------------------------------------------------------
        Set ResultsSheet = GetSheet(SH_RESULTS)
        TotalRow = ResultsRow("TOTAL " & RegimeCode)
        NettingSetRow = ResultsRow(ExpectedNettingSet)
        If TotalRow = 0 Or NettingSetRow = 0 Then
            Detail = "TOTAL " & RegimeCode & " or " & ExpectedNettingSet & " is missing from Results"
        Else
            AllHold = (SafeStr(TableCell(ResultsSheet, TotalRow, RS_STATUS_COL).Value) = StatusText) And _
                      (Abs(ToDbl(TableCell(ResultsSheet, TotalRow, 22).Value) - _
                           ToDbl(TableCell(ResultsSheet, NettingSetRow, 22).Value)) < 0.000001) And _
                      (Abs(CORE_Engine.TotalEAD(RegimeCode) - _
                           ToDbl(TableCell(ResultsSheet, NettingSetRow, 22).Value)) < 0.000001)
            Detail = "TOTAL " & RegimeCode & " status """ & _
                     SafeStr(TableCell(ResultsSheet, TotalRow, RS_STATUS_COL).Value) & _
                     """, EAD " & TableCell(ResultsSheet, TotalRow, 22).Value & "; expected """ & StatusText & _
                     """ and the EAD of " & ExpectedNettingSet & ", " & _
                     TableCell(ResultsSheet, NettingSetRow, 22).Value
        End If
        TEST_CaseRunner.ExpectTrue "total-" & LCase$(RegimeCode), AllHold, Detail, "illustrative"

End Sub


Private Sub ExpectStatus( _
    ByVal ExpectedStatus As String)
'
'==============================================================================
'                                 ExpectStatus
'------------------------------------------------------------------------------
' PURPOSE
'   Run the case and check the netting set's status on Results.
'
' UPDATED
'   2026-10-17
'==============================================================================
'
        TEST_CaseRunner.RunCase
        TEST_CaseRunner.ExpectText "status", "", "netting_set_status", ExpectedStatus, "illustrative"

End Sub
