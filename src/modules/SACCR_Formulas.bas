Attribute VB_Name = "SACCR_Formulas"
'==============================================================================
' MODULE: SACCR_Formulas
'------------------------------------------------------------------------------
' PURPOSE
'   Provide the SA-CCR building blocks as public functions. The engine calls
'   them, and they can also be entered in worksheet cells as user-defined
'   functions, so every intermediate figure can be checked independently.
'
' PUBLIC SURFACE
'   The ten SACCR_* functions, listed in docs/PUBLIC_API.txt. Their names,
'   arguments, defaults and results are a contract: changing one is a public
'   API change.
'
' DEPENDENCIES
'   VBA runtime only. SACCR_NormCDF is implemented here so that results do
'   not depend on WorksheetFunction, the Excel version or the locale.
'
' STATE OWNERSHIP
'   Stateless. Results depend only on the arguments.
'
' ERROR POLICY
'   Each function checks the domain of its inputs, stated in its INPUTS
'   section, and returns a worksheet error value instead of raising or
'   returning a misleading number (#42): #NUM! for a number outside its
'   domain, #VALUE! for a value that is not a number where one is needed,
'   #REF! for ranges of different sizes or shapes, and the error itself for
'   an error value in a range. A cell then shows the error and the rest of
'   the sheet still calculates. A VBA caller receives the same error value
'   in the Variant result and tests it with IsError; CORE_Engine stops the
'   run when it receives one. The documented floors (S below 0, M below the
'   minimum, SD below its floor) are applied, not rejected. An overflow
'   beyond the range of a Double also gives #NUM!. SACCR_NormCDF accepts any
'   real number and keeps its Double result.
'
' REFERENCE
'   BCBS CRE52 (paragraph numbers in brackets); EU CRR Art. 274 to 280f and
'   Commission Delegated Regulation (EU) 2021/931 as amended.
'
' COMPATIBILITY
'   Excel VBA; no references beyond the defaults.
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
    'Require explicit declarations. This module is public: its functions are
    'available to worksheet cells.
    Option Explicit

'------------------------------------------------------------------------------
' MODULE CONSTANTS
'------------------------------------------------------------------------------
    'VarType of a LongLong, vbLongLong; only 64-bit VBA produces it. The
    'value is used instead of the name so that 32-bit VBA compiles too.
        Private Const VT_LONGLONG As Long = 20


'
'------------------------------------------------------------------------------
'
'                             STATISTICAL FUNCTION
'
'------------------------------------------------------------------------------
'

Public Function SACCR_NormCDF( _
    ByVal x As Double) _
    As Double
'
'==============================================================================
'                                SACCR_NormCDF
'------------------------------------------------------------------------------
' PURPOSE
'   Standard normal cumulative distribution function N(x).
'
' INPUTS
'   x: any real number.
'
' RETURNS
'   N(x), between 0 and 1. Absolute error below 1E-14.
'
' REFERENCE
'   Hart (1968) double-precision algorithm, as published by G. West (2005),
'   "Better approximations to cumulative normal functions".
'
' UPDATED
'   2026-10-06
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim xAbs   As Double    '|x|
    Dim ex     As Double    'exp(-x^2 / 2)
    Dim num    As Double    'Numerator polynomial, central region
    Dim den    As Double    'Denominator polynomial, central region
    Dim b      As Double    'Continued fraction, tail region

'------------------------------------------------------------------------------
' LOWER-TAIL PROBABILITY OF |x|
'------------------------------------------------------------------------------
    'Compute N(-|x|). Beyond 37 it is below 1E-298 and is taken as 0. Up to
    '7.07 (10 / sqrt(2)) a rational approximation is used; beyond, a
    'continued fraction.
        xAbs = Abs(x)
        If xAbs > 37# Then
            SACCR_NormCDF = 0#
        Else
            ex = Exp(-xAbs * xAbs / 2#)
            If xAbs < 7.07106781186547 Then
                num = 3.52624965998911E-02 * xAbs + 0.700383064443688
                num = num * xAbs + 6.37396220353165
                num = num * xAbs + 33.912866078383
                num = num * xAbs + 112.079291497871
                num = num * xAbs + 221.213596169931
                num = num * xAbs + 220.206867912376
                den = 8.83883476483184E-02 * xAbs + 1.75566716318264
                den = den * xAbs + 16.064177579207
                den = den * xAbs + 86.7807322029461
                den = den * xAbs + 296.564248779674
                den = den * xAbs + 637.333633378831
                den = den * xAbs + 793.826512519948
                den = den * xAbs + 440.413735824752
                SACCR_NormCDF = ex * num / den
            Else
                b = xAbs + 0.65
                b = xAbs + 4# / b
                b = xAbs + 3# / b
                b = xAbs + 2# / b
                b = xAbs + 1# / b
                SACCR_NormCDF = ex / b / 2.506628274631
            End If
        End If

'------------------------------------------------------------------------------
' REFLECT FOR POSITIVE x
'------------------------------------------------------------------------------
    'By symmetry, N(x) = 1 - N(-x).
        If x > 0# Then
            SACCR_NormCDF = 1# - SACCR_NormCDF
        End If

End Function


'
'------------------------------------------------------------------------------
'
'                          TRADE-LEVEL BUILDING BLOCKS
'
'------------------------------------------------------------------------------
'

Public Function SACCR_SupervisoryDuration( _
    ByVal S As Double, _
    ByVal E As Double, _
    Optional ByVal FloorYears As Double = 0#) _
    As Variant
'
'==============================================================================
'                          SACCR_SupervisoryDuration
'------------------------------------------------------------------------------
' PURPOSE
'   Supervisory duration of an interest-rate or credit trade:
'     SD = (exp(-0.05 * S) - exp(-0.05 * E)) / 0.05
'
' INPUTS
'   S: start of the trade in years from the reporting date; a negative value
'      (already started) is floored at 0.
'   E: end of the trade in years from the reporting date; not before S
'      after its floor.
'   FloorYears: lower bound on SD in years, 0 or more; 0 (no floor) by
'      default. Ten business days is 10 / 250.
'
' RETURNS
'   SD in years, at least FloorYears; #NUM! when E is before S (after the
'   floor) or FloorYears is negative.
'
' REFERENCE
'   CRE52.34; CRR Art. 279b.
'
' UPDATED
'   2026-10-12
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim sd   As Double    'Supervisory duration, years

'------------------------------------------------------------------------------
' VALIDATE AND CALCULATE
'------------------------------------------------------------------------------
        On Error GoTo Overflowed
        If S < 0# Then
            S = 0#
        End If
        If E < S Or FloorYears < 0# Then
            SACCR_SupervisoryDuration = CVErr(xlErrNum)
            Exit Function
        End If
        sd = (Exp(-0.05 * S) - Exp(-0.05 * E)) / 0.05
        If sd < FloorYears Then
            sd = FloorYears
        End If
        SACCR_SupervisoryDuration = sd
        Exit Function

'------------------------------------------------------------------------------
' HANDLE OVERFLOW
'------------------------------------------------------------------------------
    'A result or an intermediate value beyond the range of a Double is
    'not a number a cell can show: #NUM!, as Excel gives, not a raised
    'error (#42).
Overflowed:
        SACCR_SupervisoryDuration = CVErr(xlErrNum)

End Function


Public Function SACCR_MaturityFactor( _
    ByVal M As Double, _
    Optional ByVal Margined As Boolean = False, _
    Optional ByVal MPOR_BD As Double = 10#, _
    Optional ByVal MinBD As Double = 10#, _
    Optional ByVal BDPerYear As Double = 250#) _
    As Variant
'
'==============================================================================
'                             SACCR_MaturityFactor
'------------------------------------------------------------------------------
' PURPOSE
'   Maturity factor MF of a trade:
'     unmargined: MF = sqrt(min(max(M, MinBD / BDPerYear), 1))
'     margined:   MF = 1.5 * sqrt(MPOR_BD / BDPerYear)
'
' INPUTS
'   M: remaining maturity in years; ignored when Margined is True.
'   Margined: True for a margined netting set.
'   MPOR_BD: margin period of risk in business days, above 0; 10 by
'      default; used only when Margined is True.
'   MinBD: floor on M in business days, 0 or more; 10 by default; used only
'      when Margined is False.
'   BDPerYear: business days per year, above 0; 250 by default.
'
' RETURNS
'   The maturity factor, 0 or more; #NUM! when an input it uses is outside
'   its domain.
'
' REFERENCE
'   CRE52.48 (unmargined); CRE52.52 (margined).
'
' UPDATED
'   2026-10-12
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim mEff   As Double    'Maturity after the floor and the one-year cap

'------------------------------------------------------------------------------
' VALIDATE
'------------------------------------------------------------------------------
        On Error GoTo Overflowed
    'A denominator of 0 or less, a negative MPOR under the square root, or a
    'negative floor that would let a negative M reach the square root.
        If BDPerYear <= 0# Then
            SACCR_MaturityFactor = CVErr(xlErrNum)
            Exit Function
        End If
        If Margined Then
            If MPOR_BD <= 0# Then
                SACCR_MaturityFactor = CVErr(xlErrNum)
                Exit Function
            End If
        ElseIf MinBD < 0# Then
            SACCR_MaturityFactor = CVErr(xlErrNum)
            Exit Function
        End If

'------------------------------------------------------------------------------
' CALCULATE
'------------------------------------------------------------------------------
        If Margined Then
            SACCR_MaturityFactor = 1.5 * Sqr(MPOR_BD / BDPerYear)
        Else
            mEff = M
            If mEff < MinBD / BDPerYear Then
                mEff = MinBD / BDPerYear
            End If
            If mEff > 1# Then
                mEff = 1#
            End If
            SACCR_MaturityFactor = Sqr(mEff)
        End If
        Exit Function

'------------------------------------------------------------------------------
' HANDLE OVERFLOW
'------------------------------------------------------------------------------
    'A result or an intermediate value beyond the range of a Double is
    'not a number a cell can show: #NUM!, as Excel gives, not a raised
    'error (#42).
Overflowed:
        SACCR_MaturityFactor = CVErr(xlErrNum)

End Function


Public Function SACCR_OptionDelta( _
    ByVal P As Double, _
    ByVal K As Double, _
    ByVal T As Double, _
    ByVal Vol As Double, _
    ByVal IsCall As Boolean, _
    ByVal IsBought As Boolean, _
    Optional ByVal Lambda As Double = 0#) _
    As Variant
'
'==============================================================================
'                              SACCR_OptionDelta
'------------------------------------------------------------------------------
' PURPOSE
'   Supervisory delta of an option:
'     d1 = (ln((P + L) / (K + L)) + 0.5 * Vol^2 * T) / (Vol * sqrt(T))
'     bought call +N(d1); sold call -N(d1); bought put -N(-d1); sold put +N(-d1)
'
' INPUTS
'   P: price of the underlying; P + Lambda > 0.
'   K: strike price; K + Lambda > 0.
'   T: time to the latest exercise date in years; T > 0.
'   Vol: supervisory option volatility as a decimal; Vol > 0.
'   IsCall: True for a call, False for a put.
'   IsBought: True for a bought (long) option, False for a sold one.
'   Lambda: shift L for negative prices or rates; 0 by default.
'
' RETURNS
'   The delta, between -1 and 1; #NUM! when an input is outside its domain.
'
' REFERENCE
'   CRE52.40; CRR Art. 279a; lambda per Delegated Regulation (EU) 2021/931.
'
' UPDATED
'   2026-10-06
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim d1    As Double    'Black-Scholes d1 with the lambda shift
    Dim dlt   As Double    'Delta before the sign of the position

'------------------------------------------------------------------------------
' VALIDATE
'------------------------------------------------------------------------------
        On Error GoTo Overflowed
    'The logarithm and the square root need strictly positive arguments.
        If T <= 0# Or Vol <= 0# Or (P + Lambda) <= 0# Or (K + Lambda) <= 0# Then
            SACCR_OptionDelta = CVErr(xlErrNum)
            Exit Function
        End If

'------------------------------------------------------------------------------
' CALCULATE
'------------------------------------------------------------------------------
    'Delta of a bought option first; a sold option takes the opposite sign.
        d1 = (Log((P + Lambda) / (K + Lambda)) + 0.5 * Vol * Vol * T) / (Vol * Sqr(T))
        If IsCall Then
            dlt = SACCR_NormCDF(d1)
        Else
            dlt = -SACCR_NormCDF(-d1)
        End If
        If Not IsBought Then
            dlt = -dlt
        End If
        SACCR_OptionDelta = dlt
        Exit Function

'------------------------------------------------------------------------------
' HANDLE OVERFLOW
'------------------------------------------------------------------------------
    'A result or an intermediate value beyond the range of a Double is
    'not a number a cell can show: #NUM!, as Excel gives, not a raised
    'error (#42).
Overflowed:
        SACCR_OptionDelta = CVErr(xlErrNum)

End Function


Public Function SACCR_LambdaCRR( _
    ByVal P As Double, _
    ByVal K As Double, _
    ByVal IsInterestRate As Boolean, _
    Optional ByVal Threshold As Variant) _
    As Variant
'
'==============================================================================
'                               SACCR_LambdaCRR
'------------------------------------------------------------------------------
' PURPOSE
'   Lambda shift of an option under the EU rules, determined per option:
'     interest rate: lambda = max(threshold - min(P, K), 0); threshold 0.10%
'     commodity:     lambda = max(-(1 + threshold) * min(P, K), 0); threshold 0.1
'   Under Basel the shift is left to judgement (CRE52.40 FAQ) and is entered
'   on the trade instead.
'
' INPUTS
'   P: price of the underlying.
'   K: strike price.
'   IsInterestRate: True for an option mapped to interest rate, False for one
'      mapped to commodity.
'   Threshold: optional threshold as a decimal, a number of 0 or more; when
'      omitted or blank, 0.001 for interest rate and 0.1 for commodity.
'
' RETURNS
'   Lambda, zero or positive; #VALUE! when Threshold is not a number, #NUM!
'   when it is negative, and the error itself when it is an error value.
'
' REFERENCE
'   Commission Delegated Regulation (EU) 2021/931 Art. 5, as amended by
'   Commission Delegated Regulation (EU) 2025/855.
'
' UPDATED
'   2026-10-12
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim thr      As Double     'Threshold applied
    Dim given    As Variant    'Threshold as given; Empty when omitted or blank
    Dim lowest   As Double     'min(P, K)
    Dim lam      As Double     'Lambda before the floor at zero

'------------------------------------------------------------------------------
' VALIDATE THE THRESHOLD
'------------------------------------------------------------------------------
        On Error GoTo Overflowed
    'A cell reference arrives as a Range: its value is read. An omitted or
    'blank threshold takes the default below.
        If Not IsMissing(Threshold) Then
            given = CellValue(Threshold)
            If IsError(given) Then
                SACCR_LambdaCRR = given
                Exit Function
            ElseIf Not IsEmpty(given) Then
                If Not IsNumberValue(given) Then
                    SACCR_LambdaCRR = CVErr(xlErrValue)
                    Exit Function
                ElseIf CDbl(given) < 0# Then
                    SACCR_LambdaCRR = CVErr(xlErrNum)
                    Exit Function
                End If
            End If
        End If

'------------------------------------------------------------------------------
' TAKE THE LOWER OF PRICE AND STRIKE
'------------------------------------------------------------------------------
        lowest = P
        If K < lowest Then
            lowest = K
        End If

'------------------------------------------------------------------------------
' CALCULATE
'------------------------------------------------------------------------------
        If IsInterestRate Then
            If IsEmpty(given) Then
                thr = 0.001
            Else
                thr = CDbl(given)
            End If
            lam = thr - lowest
        Else
            If IsEmpty(given) Then
                thr = 0.1
            Else
                thr = CDbl(given)
            End If
            lam = -(1# + thr) * lowest
        End If
        If lam < 0# Then
            lam = 0#
        End If
        SACCR_LambdaCRR = lam
        Exit Function

'------------------------------------------------------------------------------
' HANDLE OVERFLOW
'------------------------------------------------------------------------------
    'A result or an intermediate value beyond the range of a Double is
    'not a number a cell can show: #NUM!, as Excel gives, not a raised
    'error (#42).
Overflowed:
        SACCR_LambdaCRR = CVErr(xlErrNum)

End Function


Public Function SACCR_CDODelta( _
    ByVal A As Double, _
    ByVal D As Double, _
    ByVal IsLong As Boolean) _
    As Variant
'
'==============================================================================
'                                SACCR_CDODelta
'------------------------------------------------------------------------------
' PURPOSE
'   Supervisory delta of a CDO tranche:
'     delta = +/- 15 / ((1 + 14 * A) * (1 + 14 * D))
'   positive for purchased protection (long), negative for sold protection.
'
' INPUTS
'   A: attachment point as a decimal; A >= 0.
'   D: detachment point as a decimal; A < D <= 1.
'   IsLong: True for purchased protection.
'
' RETURNS
'   The delta; #NUM! when the tranche is not valid.
'
' REFERENCE
'   CRE52.41.
'
' UPDATED
'   2026-10-06
'==============================================================================
'

'------------------------------------------------------------------------------
' VALIDATE
'------------------------------------------------------------------------------
        On Error GoTo Overflowed
        If A < 0# Or D > 1# Or A >= D Then
            SACCR_CDODelta = CVErr(xlErrNum)
            Exit Function
        End If

'------------------------------------------------------------------------------
' CALCULATE
'------------------------------------------------------------------------------
        SACCR_CDODelta = IIf(IsLong, 1#, -1#) * 15# / ((1# + 14# * A) * (1# + 14# * D))
        Exit Function

'------------------------------------------------------------------------------
' HANDLE OVERFLOW
'------------------------------------------------------------------------------
    'A result or an intermediate value beyond the range of a Double is
    'not a number a cell can show: #NUM!, as Excel gives, not a raised
    'error (#42).
Overflowed:
        SACCR_CDODelta = CVErr(xlErrNum)

End Function


'
'------------------------------------------------------------------------------
'
'                        HEDGING-SET AND NETTING-SET FORMULAS
'
'------------------------------------------------------------------------------
'

Public Function SACCR_IREffectiveNotional( _
    ByVal D1 As Double, _
    ByVal D2 As Double, _
    ByVal D3 As Double, _
    Optional ByVal Rho12 As Double = 0.7, _
    Optional ByVal Rho23 As Double = 0.7, _
    Optional ByVal Rho13 As Double = 0.3) _
    As Variant
'
'==============================================================================
'                          SACCR_IREffectiveNotional
'------------------------------------------------------------------------------
' PURPOSE
'   Effective notional of an interest-rate hedging set across the three
'   maturity buckets:
'     EN = sqrt(D1^2 + D2^2 + D3^2 + 2*r12*D1*D2 + 2*r23*D2*D3 + 2*r13*D1*D3)
'   The default correlations give the regulatory cross terms 1.4, 1.4, 0.6.
'
' INPUTS
'   D1, D2, D3: signed effective notionals of the buckets under one year,
'      one to five years and over five years.
'   Rho12, Rho23, Rho13: bucket correlations as decimals, each from 0 to 1,
'      that together form a valid correlation matrix (the range and test
'      CORE_Engine.LoadParams applies to the IRCorrBucket parameters).
'
' RETURNS
'   The effective notional, zero or positive; #NUM! when the correlations
'   are out of range or do not form a valid correlation matrix.
'
' REFERENCE
'   CRE52.57.
'
' UPDATED
'   2026-10-12
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim q          As Double    'Quadratic form under the square root
    Dim termSize   As Double    '(|D1| + |D2| + |D3|)^2, the size of the terms

'------------------------------------------------------------------------------
' VALIDATE
'------------------------------------------------------------------------------
        On Error GoTo Overflowed
    'The determinant of the 3x3 correlation matrix must not be negative;
    'the tolerance is the one CORE_Engine.LoadParams uses.
        If Rho12 < 0# Or Rho12 > 1# Or Rho23 < 0# Or Rho23 > 1# Or Rho13 < 0# Or Rho13 > 1# Then
            SACCR_IREffectiveNotional = CVErr(xlErrNum)
            Exit Function
        End If
        If 1# - Rho12 * Rho12 - Rho23 * Rho23 - Rho13 * Rho13 _
           + 2# * Rho12 * Rho23 * Rho13 < -0.000000000001 Then
            SACCR_IREffectiveNotional = CVErr(xlErrNum)
            Exit Function
        End If

'------------------------------------------------------------------------------
' CALCULATE
'------------------------------------------------------------------------------
    'With a valid matrix the quadratic form is not negative; a negative
    'value within rounding of its terms is taken as 0, anything below is
    'an error rather than a zero effective notional.
        q = D1 * D1 + D2 * D2 + D3 * D3 _
            + 2# * Rho12 * D1 * D2 + 2# * Rho23 * D2 * D3 + 2# * Rho13 * D1 * D3
        If q < 0# Then
            termSize = (Abs(D1) + Abs(D2) + Abs(D3)) ^ 2
            If -q > 0.000000000001 * termSize Then
                SACCR_IREffectiveNotional = CVErr(xlErrNum)
                Exit Function
            End If
            q = 0#
        End If
        SACCR_IREffectiveNotional = Sqr(q)
        Exit Function

'------------------------------------------------------------------------------
' HANDLE OVERFLOW
'------------------------------------------------------------------------------
    'A result or an intermediate value beyond the range of a Double is
    'not a number a cell can show: #NUM!, as Excel gives, not a raised
    'error (#42).
Overflowed:
        SACCR_IREffectiveNotional = CVErr(xlErrNum)

End Function


Public Function SACCR_Multiplier( _
    ByVal VminusC As Double, _
    ByVal AddOn As Double, _
    Optional ByVal FloorPct As Double = 0.05) _
    As Variant
'
'==============================================================================
'                               SACCR_Multiplier
'------------------------------------------------------------------------------
' PURPOSE
'   PFE multiplier of a netting set:
'     m = min(1, F + (1 - F) * exp((V - C) / (2 * (1 - F) * AddOn)))
'
' INPUTS
'   VminusC: V - C, value of the trades less the collateral held.
'   AddOn: aggregate add-on of the netting set, 0 or more; when 0, m is 1.
'   FloorPct: floor F as a decimal, at least 0 and below 1; 0.05 by default.
'
' RETURNS
'   The multiplier, between FloorPct and 1; #NUM! when AddOn is negative or
'   FloorPct is outside its domain.
'
' REFERENCE
'   CRE52.23.
'
' UPDATED
'   2026-10-12
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim m             As Double    'Multiplier
    Dim z             As Double    'Exponent of the formula
    Dim denominator   As Double    '2 * (1 - F) * AddOn

'------------------------------------------------------------------------------
' VALIDATE
'------------------------------------------------------------------------------
        On Error GoTo Overflowed
    'A floor of 1 or more divides by zero or inverts the formula.
        If AddOn < 0# Or FloorPct < 0# Or FloorPct >= 1# Then
            SACCR_Multiplier = CVErr(xlErrNum)
            Exit Function
        End If

'------------------------------------------------------------------------------
' HANDLE A ZERO ADD-ON
'------------------------------------------------------------------------------
        If AddOn = 0# Then
            SACCR_Multiplier = 1#
            Exit Function
        End If

'------------------------------------------------------------------------------
' CALCULATE
'------------------------------------------------------------------------------
    'A V - C of 0 or more always gives 1 after the cap. An exponent below
    '-700 makes Exp underflow, so the result is the floor; the comparison is
    'made before the division, which a very small add-on would overflow.
        denominator = 2# * (1# - FloorPct) * AddOn
        If VminusC >= 0# Then
            m = 1#
        ElseIf VminusC < -700# * denominator Then
            m = FloorPct
        Else
            z = VminusC / denominator
            m = FloorPct + (1# - FloorPct) * Exp(z)
        End If
        If m > 1# Then
            m = 1#
        End If
        SACCR_Multiplier = m
        Exit Function

'------------------------------------------------------------------------------
' HANDLE OVERFLOW
'------------------------------------------------------------------------------
    'A result or an intermediate value beyond the range of a Double is
    'not a number a cell can show: #NUM!, as Excel gives, not a raised
    'error (#42).
Overflowed:
        SACCR_Multiplier = CVErr(xlErrNum)

End Function


Public Function SACCR_ReplacementCost( _
    ByVal V As Double, _
    ByVal C As Double, _
    Optional ByVal Margined As Boolean = False, _
    Optional ByVal TH As Double = 0#, _
    Optional ByVal MTA As Double = 0#, _
    Optional ByVal NICA As Double = 0#) _
    As Variant
'
'==============================================================================
'                            SACCR_ReplacementCost
'------------------------------------------------------------------------------
' PURPOSE
'   Replacement cost RC of a netting set:
'     unmargined: RC = max(V - C, 0)
'     margined:   RC = max(V - C, TH + MTA - NICA, 0)
'
' INPUTS
'   V: value of the trades in the netting set.
'   C: haircut value of the net collateral held.
'   Margined: True for a margined netting set.
'   TH: threshold, 0 or more; MTA: minimum transfer amount, 0 or more;
'      NICA: net independent collateral amount, any sign. Used only when
'      Margined is True.
'
' RETURNS
'   RC, zero or positive; #NUM! when a margined set has a negative TH or
'   MTA.
'
' REFERENCE
'   CRE52.10 (unmargined); CRE52.18 (margined).
'
' UPDATED
'   2026-10-12
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim rc   As Double    'Replacement cost

'------------------------------------------------------------------------------
' VALIDATE
'------------------------------------------------------------------------------
        On Error GoTo Overflowed
        If Margined And (TH < 0# Or MTA < 0#) Then
            SACCR_ReplacementCost = CVErr(xlErrNum)
            Exit Function
        End If

'------------------------------------------------------------------------------
' CALCULATE
'------------------------------------------------------------------------------
        rc = V - C
        If Margined Then
            If TH + MTA - NICA > rc Then
                rc = TH + MTA - NICA
            End If
        End If
        If rc < 0# Then
            rc = 0#
        End If
        SACCR_ReplacementCost = rc
        Exit Function

'------------------------------------------------------------------------------
' HANDLE OVERFLOW
'------------------------------------------------------------------------------
    'A result or an intermediate value beyond the range of a Double is
    'not a number a cell can show: #NUM!, as Excel gives, not a raised
    'error (#42).
Overflowed:
        SACCR_ReplacementCost = CVErr(xlErrNum)

End Function


Public Function SACCR_FactorAggregation( _
    ByVal AddOns As Variant, _
    ByVal Rhos As Variant) _
    As Variant
'
'==============================================================================
'                           SACCR_FactorAggregation
'------------------------------------------------------------------------------
' PURPOSE
'   Single-factor aggregation used for credit, equity and commodity:
'     AddOn = sqrt((sum rho_k * A_k)^2 + sum (1 - rho_k^2) * A_k^2)
'
' INPUTS
'   AddOns: the entity or commodity-type add-ons A_k, any sign; a range, an
'      array or a single number. A blank add-on counts as 0.
'   Rhos: the correlations rho_k as decimals, each from 0 to 1; as many as
'      AddOns. A blank correlation is allowed only beside a blank add-on.
'   A single number and a one-element range or array are the same input.
'   Two vectors (one row, one column, or one-dimensional) are paired in
'   order whatever their orientation; any other range or array must have
'   the same rows and columns as the other input.
'
' RETURNS
'   The aggregated add-on; #REF! when AddOns and Rhos differ in size or
'   shape, #VALUE! when an element is text, a Boolean or another value
'   that is not a number, #NUM! when a correlation is outside 0 to 1, and
'   the first error value found in either input.
'
' REFERENCE
'   CRE52.61 (credit), CRE52.66 (equity), CRE52.70 (commodity).
'
' UPDATED
'   2026-10-12
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim va       As Variant    'Add-ons as a 1-based list
    Dim vr       As Variant    'Correlations as a 1-based list
    Dim aRows    As Long       'Rows of AddOns; 0 for one dimension
    Dim aCols    As Long       'Columns of AddOns
    Dim rRows    As Long       'Rows of Rhos; 0 for one dimension
    Dim rCols    As Long       'Columns of Rhos
    Dim i        As Long       'Element counter
    Dim sys      As Double     'Systematic part: sum rho_k * A_k
    Dim idio     As Double     'Idiosyncratic part: sum (1 - rho_k^2) * A_k^2
    Dim av       As Double     'Add-on A_k
    Dim rv       As Double     'Correlation rho_k

'------------------------------------------------------------------------------
' FLATTEN AND COMPARE THE SHAPES
'------------------------------------------------------------------------------
        On Error GoTo Overflowed
        va = FlatList(CellValue(AddOns), aRows, aCols)
        vr = FlatList(CellValue(Rhos), rRows, rCols)
        If aRows < 0 Or rRows < 0 Then
            SACCR_FactorAggregation = CVErr(xlErrRef)
            Exit Function
        End If
        If UBound(va) <> UBound(vr) Then
            SACCR_FactorAggregation = CVErr(xlErrRef)
            Exit Function
        End If
        If Not (IsVectorShape(aRows, aCols) And IsVectorShape(rRows, rCols)) Then
            If aRows <> rRows Or aCols <> rCols Then
                SACCR_FactorAggregation = CVErr(xlErrRef)
                Exit Function
            End If
        End If

'------------------------------------------------------------------------------
' VALIDATE THE ELEMENTS
'------------------------------------------------------------------------------
    'Errors first, so that an error in a range is passed on unchanged.
        For i = 1 To UBound(va)
            If IsError(va(i)) Then
                SACCR_FactorAggregation = va(i)
                Exit Function
            ElseIf IsError(vr(i)) Then
                SACCR_FactorAggregation = vr(i)
                Exit Function
            End If
        Next i
        For i = 1 To UBound(va)
            If IsEmpty(va(i)) Then
                If Not IsEmpty(vr(i)) And Not IsNumberValue(vr(i)) Then
                    SACCR_FactorAggregation = CVErr(xlErrValue)
                    Exit Function
                End If
            ElseIf Not IsNumberValue(va(i)) Or Not IsNumberValue(vr(i)) Then
                SACCR_FactorAggregation = CVErr(xlErrValue)
                Exit Function
            End If
            If IsNumberValue(vr(i)) Then
                If CDbl(vr(i)) < 0# Or CDbl(vr(i)) > 1# Then
                    SACCR_FactorAggregation = CVErr(xlErrNum)
                    Exit Function
                End If
            End If
        Next i

'------------------------------------------------------------------------------
' AGGREGATE
'------------------------------------------------------------------------------
        For i = 1 To UBound(va)
            If Not IsEmpty(va(i)) Then
                av = CDbl(va(i))
                rv = CDbl(vr(i))
                sys = sys + rv * av
                idio = idio + (1# - rv * rv) * av * av
            End If
        Next i
        SACCR_FactorAggregation = Sqr(sys * sys + idio)
        Exit Function

'------------------------------------------------------------------------------
' HANDLE OVERFLOW
'------------------------------------------------------------------------------
    'A result or an intermediate value beyond the range of a Double is
    'not a number a cell can show: #NUM!, as Excel gives, not a raised
    'error (#42).
Overflowed:
        SACCR_FactorAggregation = CVErr(xlErrNum)

End Function


'
'------------------------------------------------------------------------------
'
'                              INPUT HELPERS
'
'------------------------------------------------------------------------------
'

Private Function CellValue( _
    ByVal v As Variant) _
    As Variant
'
'==============================================================================
'                                  CellValue
'------------------------------------------------------------------------------
' PURPOSE
'   Return the value of a worksheet argument: the Value2 of a Range (a
'   2-D array, or one value for a single cell), so that a cell's number
'   format does not convert it (#196); anything else unchanged.
'
' UPDATED
'   2026-10-21
'==============================================================================
'
        If IsObject(v) Then
            If TypeOf v Is Range Then
                CellValue = v.Value2
                Exit Function
            End If
        End If
        CellValue = v

End Function


Private Function IsNumberValue( _
    ByVal v As Variant) _
    As Boolean
'
'==============================================================================
'                                IsNumberValue
'------------------------------------------------------------------------------
' PURPOSE
'   Tell whether a value is a number, including a LongLong in 64-bit VBA:
'   not text, a Boolean, blank or an error, even when VBA would convert it.
'
' UPDATED
'   2026-10-12
'==============================================================================
'
        Select Case VarType(v)
            Case vbDouble, vbSingle, vbInteger, vbLong, vbCurrency, vbDecimal, vbByte, VT_LONGLONG
                IsNumberValue = True
        End Select

End Function


Private Function FlatList( _
    ByVal v As Variant, _
    ByRef nRows As Long, _
    ByRef nCols As Long) _
    As Variant
'
'==============================================================================
'                                   FlatList
'------------------------------------------------------------------------------
' PURPOSE
'   Return the elements of a value, an array or a 2-D array as a 1-based
'   list, row by row, with its shape: rows and columns of a 2-D array, 0
'   rows and the element count for one dimension, 1 by 1 for one value.
'   An array without bounds or with more than two dimensions gives -1 rows
'   and a one-element list, for the caller to reject.
'
' UPDATED
'   2026-10-12
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim list()   As Variant    'The elements
    Dim r        As Long       'Row
    Dim c        As Long       'Column
    Dim k        As Long       'Position in list

'------------------------------------------------------------------------------
' FLATTEN
'------------------------------------------------------------------------------
        If Not IsArray(v) Then
            nRows = 1
            nCols = 1
            ReDim list(1 To 1)
            list(1) = v
        ElseIf ArrayDimensions(v) = 0 Or ArrayDimensions(v) > 2 Then
            nRows = -1
            nCols = 0
            ReDim list(1 To 1)
        ElseIf ArrayDimensions(v) = 1 Then
            nRows = 0
            nCols = UBound(v) - LBound(v) + 1
            ReDim list(1 To nCols)
            For c = LBound(v) To UBound(v)
                k = k + 1
                list(k) = v(c)
            Next c
        Else
            nRows = UBound(v, 1) - LBound(v, 1) + 1
            nCols = UBound(v, 2) - LBound(v, 2) + 1
            ReDim list(1 To nRows * nCols)
            For r = LBound(v, 1) To UBound(v, 1)
                For c = LBound(v, 2) To UBound(v, 2)
                    k = k + 1
                    list(k) = v(r, c)
                Next c
            Next r
        End If
        FlatList = list

End Function


Private Function ArrayDimensions( _
    ByVal v As Variant) _
    As Long
'
'==============================================================================
'                               ArrayDimensions
'------------------------------------------------------------------------------
' PURPOSE
'   Return the number of dimensions of an array, up to 3: 0 for a dynamic
'   array without bounds, 3 for three or more.
'
' ERROR POLICY
'   Contains the expected error of UBound on a missing dimension.
'
' UPDATED
'   2026-10-12
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim probe   As Long    'Upper bound of the probed dimension

'------------------------------------------------------------------------------
' PROBE
'------------------------------------------------------------------------------
        On Error GoTo Done
        probe = UBound(v, 1)
        ArrayDimensions = 1
        probe = UBound(v, 2)
        ArrayDimensions = 2
        probe = UBound(v, 3)
        ArrayDimensions = 3

Done:
        Err.Clear

End Function


Private Function IsVectorShape( _
    ByVal nRows As Long, _
    ByVal nCols As Long) _
    As Boolean
'
'==============================================================================
'                                IsVectorShape
'------------------------------------------------------------------------------
' PURPOSE
'   Tell whether a shape from FlatList is a vector: one dimension, one row
'   or one column.
'
' UPDATED
'   2026-10-12
'==============================================================================
'
        IsVectorShape = (nRows <= 1 Or nCols = 1)

End Function
