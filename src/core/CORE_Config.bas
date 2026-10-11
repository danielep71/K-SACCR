Attribute VB_Name = "CORE_Config"
'==============================================================================
' MODULE: CORE_Config
'------------------------------------------------------------------------------
' PURPOSE
'   Hold every layout constant the SA-CCR engine depends on: sheet names,
'   header and first data rows, the column of each input field, the width of
'   each output table, parameter codes, regime and asset-class codes, the
'   severity labels written to the Checks sheet, the expected column
'   headers, and the error numbers raised by the workbook macros. When a
'   sheet layout changes, it changes here and nowhere else.
'
' PUBLIC SURFACE
'   None outside this VBA project. The constants are Public for in-project use
'   by CORE_Util, CORE_Engine and M_Main; Option Private Module keeps them off the
'   supported external surface.
'
' DEPENDENCIES
'   None. Every value is a literal.
'
' STATE OWNERSHIP
'   Constants only; no state.
'
' ERROR POLICY
'   Not applicable.
'
' COMPATIBILITY
'   Excel VBA; no references beyond the defaults.
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
    'Require explicit declarations; keep the constants inside this project.
    Option Explicit
    Option Private Module

'------------------------------------------------------------------------------
' SHEET NAMES
'------------------------------------------------------------------------------
    'Worksheet tab names, as the engine looks them up in ThisWorkbook.
        Public Const SH_README      As String = "README"         'Read-me; not used by the engine
        Public Const SH_PARAMS      As String = "Params"         'Parameters, factor and FX tables
        Public Const SH_NS          As String = "NettingSets"    'Netting-set inputs
        Public Const SH_TRADES      As String = "Trades"         'Trade inputs
        Public Const SH_RESULTS     As String = "Results"        'One row per netting set, and a total
        Public Const SH_HEDGING     As String = "HedgingSets"    'One row per hedging set
        Public Const SH_BUCKETS     As String = "Buckets"        'One row per bucket
        Public Const SH_TRADECALC   As String = "TradeCalc"      'One row per trade
        Public Const SH_CHECKS      As String = "Checks"         'Errors, warnings, information

'------------------------------------------------------------------------------
' COMMON LAYOUT
'------------------------------------------------------------------------------
    'Every sheet has two narrow margin columns (A:B) and two margin rows
    '(1:2). Every table starts in column C. The output tables have their
    'headers on row 6 and data from row 7. The input tables are located
    'through workbook names, so their rows can move: NettingSets and Trades
    'through the name of their first header cell (NM_NS_HEADER,
    'NM_TR_HEADER), with data from the next row (CORE_Util.TableHeaderRow
    'and TableFirstRow), and Params as described under PARAMS SHEET. Column
    'numbers in this
    'module (NS_, TR_, RS_ and the others) count from the first table
    'column; CORE_Util.TableCell and TableRange add COL_OFFSET. The run
    'summary is written to cell C4 of the Results sheet.
        Public Const COL_OFFSET       As Long = 2         'Margin columns left of every table
        Public Const HEADER_ROW       As Long = 6         'Row holding the column headers
        Public Const FIRST_DATA_ROW   As Long = 7         'First row of data below the headers
        Public Const RUNINFO_CELL     As String = "C4"    'Cell receiving the run summary line
        Public Const NM_NS_HEADER     As String = "NettingSets_Header"    'NettingSets header cell "NettingSetID"
        Public Const NM_TR_HEADER     As String = "Trades_Header"         'Trades header cell "TradeID"

'------------------------------------------------------------------------------
' PARAMS SHEET
'------------------------------------------------------------------------------
    'The Params sheet is located through workbook names, not fixed rows
    'and columns, so its layout can change: NM_PRM_CODE and NM_PRM_VALUE
    'name the "Code" and "Value" header cells of the parameter rows, and
    'the top-left cells of NM_SF_TABLE and NM_FX_TABLE are the header cells
    'of the two lookup tables (CORE_Util.ParamsCodeCol and the related
    'functions). Each parameter is a row with its code under "Code" and its
    'value under "Value", read through the workbook name equal to its code
    '(CORE_Util.GetParam). BD = business days; SD = supervisory duration;
    'SF = supervisory factor; CO = commodity.
        Public Const NM_PRM_CODE      As String = "Params_Code"            'Header cell "Code"
        Public Const NM_PRM_VALUE     As String = "Params_Value"           'Header cell "Value"
        Public Const NM_SF_TABLE      As String = "SF_Table"               'Factor table, from its header
        Public Const NM_FX_TABLE      As String = "FX_Table"               'FX table, from its header
        Public Const PRM_ASOF         As String = "AsOfDate"               'Reporting date
        Public Const PRM_REPCCY       As String = "ReportingCcy"           'Reporting currency
        Public Const PRM_ALPHA        As String = "Alpha"                  'Alpha
        Public Const PRM_FLOOR        As String = "MultiplierFloor"        'Multiplier floor
        Public Const PRM_DAYSYEAR     As String = "DaysPerYear"            'Days per year for S, E, M, T
        Public Const PRM_BDYEAR       As String = "BusinessDaysPerYear"    'Business days per year
        Public Const PRM_MINMAT       As String = "MinMaturityBD"          'Floor on M, BD
        Public Const PRM_SDFLOOR      As String = "SDFloorBD"              'Floor on SD, BD
        Public Const PRM_MPOR_BIL     As String = "MPORFloorBilateral"     'MPOR floor, bilateral
        Public Const PRM_MPOR_CLR     As String = "MPORFloorCleared"       'MPOR floor, cleared
        Public Const PRM_MPOR_LARGE   As String = "MPORFloorLarge"         'MPOR floor, large/illiquid
        Public Const PRM_BASIS        As String = "BasisFactor"            'SF multiplier, basis
        Public Const PRM_VOLF         As String = "VolatilityFactor"       'SF multiplier, volatility
        Public Const PRM_RHO12        As String = "IRCorrBucket12"         'IR correlation, buckets 1-2
        Public Const PRM_RHO23        As String = "IRCorrBucket23"         'IR correlation, buckets 2-3
        Public Const PRM_RHO13        As String = "IRCorrBucket13"         'IR correlation, buckets 1-3
        Public Const PRM_IRFULL       As String = "IRBucketOffset"         'True: bucket formula
        Public Const PRM_REGIME       As String = "Regime"                 'Default regime
        Public Const PRM_LAMIR        As String = "LambdaThresholdIR"      'CRR lambda threshold, IR
        Public Const PRM_LAMCO        As String = "LambdaThresholdCO"      'CRR lambda threshold, CO

'------------------------------------------------------------------------------
' RUN FINGERPRINT
'------------------------------------------------------------------------------
    'Hidden workbook name holding the input fingerprint of the last
    'completed run, so that results can be recognised as out of date (#36).
    'It always exists and is empty when no results exist, so that the
    'status formula on Results, which refers to it, recalculates when a run
    'changes it.
        Public Const RUN_INPUTS_NAME       As String = "SACCR_RunInputs"    'Empty when no results exist
        Public Const RESULTS_STATUS_CELL   As String = "C5"                 'Results: status formula

'------------------------------------------------------------------------------
' REGIMES
'------------------------------------------------------------------------------
    'Regime codes accepted on Params and as a netting-set override.
        Public Const RG_BCBS   As String = "BCBS"    'Basel framework, CRE52
        Public Const RG_CRR    As String = "CRR"     'EU Capital Requirements Regulation

'------------------------------------------------------------------------------
' REGULATORY PARAMETER SET
'------------------------------------------------------------------------------
    'The values a regulatory run uses (#41), as listed in the methodology:
    'CRE52 with CRE52.72, and the consolidated CRR of 26.06.2026 with
    'Delegated Regulation (EU) 2021/931 as amended by 2025/855, checked on
    '2026-10-09. A run whose parameters, factor table and alpha overrides
    'all equal this set is REGULATORY and names REF_SET_ID; any difference
    'makes it WHAT-IF (CORE_Engine.ClassifyRun). DaysPerYear and
    'IRBucketOffset are conventions and do not change the class.
        Public Const REF_SET_ID       As String = "CRE52-CRR-2026-10-09"    'Identity of the set
        Public Const RUN_REGULATORY   As String = "REGULATORY"              'Every parameter as the set
        Public Const RUN_WHATIF       As String = "WHAT-IF"                 'Some parameter differs
        Public Const REF_ALPHA        As Double = 1.4       'CRE52.1; CRR Art. 274(2)
        Public Const REF_FLOOR        As Double = 0.05      'CRE52.23; CRR Art. 278(3)
        Public Const REF_BDYEAR       As Double = 250       'CRE52.48; CRR Art. 279c(1)
        Public Const REF_MINMAT       As Double = 10        'CRE52.48; CRR Art. 279c(1)(a)
        Public Const REF_SDFLOOR      As Double = 10        'CRE52.34; CRR Art. 279b(1)(a)
        Public Const REF_MPOR_BIL     As Double = 10        'CRE52.50; CRR Art. 285(2)(b)
        Public Const REF_MPOR_CLR     As Double = 5         'CRE54.12; CRR Art. 279c(1)(b)
        Public Const REF_MPOR_LARGE   As Double = 20        'CRE52.50; CRR Art. 285(3)
        Public Const REF_BASIS        As Double = 0.5       'CRE52.46; CRR Art. 280
        Public Const REF_VOLF         As Double = 5         'CRE52.47; CRR Art. 280
        Public Const REF_RHO12        As Double = 0.7       'CRE52.57; CRR Art. 280a(3)
        Public Const REF_RHO23        As Double = 0.7       'CRE52.57; CRR Art. 280a(3)
        Public Const REF_RHO13        As Double = 0.3       'CRE52.57; CRR Art. 280a(3)
        Public Const REF_LAMIR        As Double = 0.001     'CRR-RTS Art. 5(2)
        Public Const REF_LAMCO        As Double = 0.1       'CRR-RTS Art. 5(2)

'------------------------------------------------------------------------------
' NETTING-SET CODES
'------------------------------------------------------------------------------
    'Client-clearing role of the bank; blank is N.
        Public Const CL_NONE     As String = "N"         'Not a client-clearing netting set
        Public Const CL_CM       As String = "CM"        'Clearing member facing its client
        Public Const CL_CLIENT   As String = "CLIENT"    'Client facing its clearing member
    'Margin-agreement structure; blank is one agreement, or none, covering
    'this netting set only. Both codes are declared so that the netting set
    'is rejected: the engine does not model them.
        Public Const MS_MIXED    As String = "MIXED"     'Several agreements, or margined and unmargined trades
        Public Const MS_SHARED   As String = "SHARED"    'One agreement over several netting sets
    'Trades in a netting set above which the 20-day MPOR floor applies in the
    'following quarter [CRE52.51(1); CRR Art. 285(3)(a)].
        Public Const LARGE_NS_TRADES As Long = 5000

'------------------------------------------------------------------------------
' PARAMS TABLE HEADERS
'------------------------------------------------------------------------------
    'Header text of the first column of each lookup table, at the top-left
    'cell of its workbook name. The table runs down to the first blank key.
        Public Const HDR_SF   As String = "SF_Key"    'Supervisory-factor table
        Public Const HDR_FX   As String = "FX_Ccy"    'FX-rate table

'------------------------------------------------------------------------------
' NETTINGSETS INPUT COLUMNS
'------------------------------------------------------------------------------
    'Column number of each field on the NettingSets sheet.
        Public Const NS_ID         As Long = 1     'Netting-set ID
        Public Const NS_CPTY       As Long = 2     'Counterparty
        Public Const NS_MARGINED   As Long = 3     'Margined (Y/N)
        Public Const NS_CLEARED    As Long = 4     'Client-clearing role: N, CM or CLIENT
        Public Const NS_FREQ       As Long = 5     'Remargining frequency, business days
        Public Const NS_LARGE      As Long = 6     'Over 5,000 trades or illiquid collateral (Y/N)
        Public Const NS_DISPUTE    As Long = 7     'Margin disputes (Y/N)
        Public Const NS_MPOR       As Long = 8     'MPOR override, business days
        Public Const NS_VM         As Long = 9     'Net variation margin held (+) or posted (-)
        Public Const NS_NICA       As Long = 10    'Net independent collateral amount
        Public Const NS_TH         As Long = 11    'Threshold
        Public Const NS_MTA        As Long = 12    'Minimum transfer amount
        Public Const NS_ALPHA      As Long = 13    'Alpha override
        Public Const NS_REGIME     As Long = 14    'Regime override; blank uses Params
        Public Const NS_STRUCT     As Long = 15    'Margin-agreement structure: blank, MIXED or SHARED
        Public Const NS_NCOLS      As Long = 15    'Number of input columns

'------------------------------------------------------------------------------
' TRADES INPUT COLUMNS
'------------------------------------------------------------------------------
    'Column number of each field on the Trades sheet.
        Public Const TR_ID         As Long = 1     'Trade ID
        Public Const TR_NS         As Long = 2     'Netting-set ID
        Public Const TR_AC         As Long = 3     'Asset class: IR, FX, CR, EQ, CO or OT
        Public Const TR_SUB        As Long = 4     'Sub-class, for example AAA or SINGLE
        Public Const TR_RF         As Long = 5     'Risk factor or reference
        Public Const TR_INSTR      As Long = 6     'Instrument: Linear, Option or CDO
        Public Const TR_DIR        As Long = 7     'Direction: Long or Short
        Public Const TR_OPT        As Long = 8     'Option type: Call or Put
        Public Const TR_NATURE     As Long = 9     'Nature: Standard, Basis or Volatility
        Public Const TR_LABEL      As Long = 10    'Basis or volatility hedging-set label
        Public Const TR_NOTIONAL   As Long = 11    'Notional, in the notional currency
        Public Const TR_NCCY       As Long = 12    'Notional currency
        Public Const TR_MTM        As Long = 13    'Mark-to-market value, in the MtM currency
        Public Const TR_MCCY       As Long = 14    'MtM currency; blank uses the notional currency
        Public Const TR_START      As Long = 15    'Start date (S)
        Public Const TR_END        As Long = 16    'End date (E)
        Public Const TR_MAT        As Long = 17    'Maturity date (M)
        Public Const TR_EXPIRY     As Long = 18    'Option expiry date (T)
        Public Const TR_PRICE      As Long = 19    'Underlying price P
        Public Const TR_STRIKE     As Long = 20    'Strike K
        Public Const TR_LAMBDA     As Long = 21    'Lambda shift
        Public Const TR_ATTACH     As Long = 22    'CDO attachment point A
        Public Const TR_DETACH     As Long = 23    'CDO detachment point D
        Public Const TR_COMMENT    As Long = 24    'Free-text comment, not read
        Public Const TR_NCOLS      As Long = 24    'Number of input columns

'------------------------------------------------------------------------------
' EXPECTED HEADERS
'------------------------------------------------------------------------------
    'Header text of every column the engine reads, in column order and
    'separated by "|"; an empty entry is a column the engine does not read.
    'The engine reads by column number, so a header that differs (ignoring
    'case, spaces and line breaks) means a column was inserted, deleted or
    'moved, and the run stops (#35). The Params, factor and FX headers are
    'checked from the cells that their workbook names locate.
        Public Const NS_HEADERS As String = "NettingSetID|Counterparty|Margined (Y/N)|" & _
            "Client clearing role (N / CM / CLIENT)|Remargin frequency N (BD)|" & _
            ">5,000 trades or illiquid collateral (Y/N)|Margin disputes (Y/N)|" & _
            "MPOR override (BD)|Net VM held (+) / posted (-)|NICA|Threshold TH|MTA|" & _
            "Alpha override|Regime override (blank = Params)|" & _
            "Margin agreements (blank / MIXED / SHARED)"
        Public Const TR_HEADERS As String = "TradeID|NettingSetID|Asset class|Sub-class|" & _
            "Risk factor / reference|Instrument|Direction|Option type|Nature|" & _
            "Basis / vol hedging-set label|Notional|Notional ccy|MtM|MtM ccy|" & _
            "Start date (S)|End date (E)|Maturity date (M)|Option expiry (T)|" & _
            "Underlying price P|Strike K|Lambda shift|Attachment A|Detachment D|"
        Public Const PRM_CODE_HEADER As String = "Code"
        Public Const PRM_VALUE_HEADER As String = "Value"
        Public Const SF_HEADERS As String = "SF_Key|Asset class|Category|Supervisory factor|" & _
            "Correlation|Supervisory option vol|Commodity hedging set|Regimes"
        Public Const FX_HEADERS As String = "FX_Ccy||Units of reporting ccy per 1 unit"

'------------------------------------------------------------------------------
' OUTPUT TABLE WIDTHS
'------------------------------------------------------------------------------
    'Number of columns the engine clears and writes on each output sheet.
        Public Const TC_NCOLS   As Long = 25    'TradeCalc
        Public Const BK_NCOLS   As Long = 11    'Buckets
        Public Const HS_NCOLS   As Long = 14    'HedgingSets
        Public Const RS_NCOLS   As Long = 28    'Results
        Public Const CK_NCOLS   As Long = 5     'Checks
        Public Const RS_STATUS_COL  As Long = 28    'Results: VALID, INCOMPLETE or NO TRADES

'------------------------------------------------------------------------------
' ASSET CLASSES
'------------------------------------------------------------------------------
    'Internal asset-class numbers, used as array indexes. OT (other risks)
    'exists only under CRR Art. 277(1)(f) and Art. 280f.
        Public Const AC_IR      As Long = 1    'Interest rate
        Public Const AC_FX      As Long = 2    'Foreign exchange
        Public Const AC_CR      As Long = 3    'Credit
        Public Const AC_EQ      As Long = 4    'Equity
        Public Const AC_CO      As Long = 5    'Commodity
        Public Const AC_OT      As Long = 6    'Other risks, CRR only
        Public Const AC_COUNT   As Long = 6    'Number of asset classes

'------------------------------------------------------------------------------
' ERROR NUMBERS
'------------------------------------------------------------------------------
    'Errors raised by the workbook macros, in the project range
    'vbObjectError + 2048 to + 4095. The regression harness uses its own
    'numbers from + 2060.
        Public Const ERR_RUN_ACTIVE       As Long = vbObjectError + 2049    'An operation is already running
        Public Const ERR_CLEANUP_FAILED   As Long = vbObjectError + 2050    'Excel settings were not restored
        Public Const ERR_INJECTED_FAULT   As Long = vbObjectError + 2051    'Test seam: injected failure
        Public Const ERR_TEST_SETUP       As Long = vbObjectError + 2052    'Test runner: workbook not set up
        Public Const ERR_CHECKS_WRITE     As Long = vbObjectError + 2053    'Checks sheet could not be written

'------------------------------------------------------------------------------
' SEVERITY LABELS
'------------------------------------------------------------------------------
    'Severity written in the first column of the Checks sheet.
        Public Const SEV_ERROR   As String = "ERROR"      'Input rejected or run stopped
        Public Const SEV_WARN    As String = "WARNING"    'Input accepted with an assumption
        Public Const SEV_INFO    As String = "INFO"       'Information only
