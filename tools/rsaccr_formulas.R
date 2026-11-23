# Compute, with the CRAN package SACCR and its Trading classes, the values of
# the FormulaTest checks that the package can compute (#100). They are the
# independent values of the case formula-test.independent in
# tests/modules/TEST_Formulas.bas, which calls the SACCR_* worksheet functions
# on the same inputs.
#
# Usage: Rscript tools/rsaccr_formulas.R
#
# Each line is: FormulaTest row, check, value with 17 significant digits.
# The package is called, not copied: it is GPL-3 and this repository is
# MPL-2.0. Rows the package cannot compute, or computes on another rule, are
# listed in docs/methodology/README.md (Independent values of the FormulaTest
# checks).

suppressPackageStartupMessages({
    library(SACCR)
    library(Trading)
})

swap <- function(S, E) Trading::IRDSwap(Notional = 1, MtM = 0, Si = S, Ei = E, BuySell = "Buy", Currency = "EUR")

swaption <- function(T, P, K, type) {
    # For a swaption the package takes the option's expiry as S.
    Trading::IRDSwaption(Notional = 1, MtM = 0, Si = T, Ei = T + 1, BuySell = "Buy", OptionType = type,
                         UnderlyingPrice = P, StrikePrice = K, Currency = "EUR")
}

margined <- Trading::CSA(ID = "csa", thres_cpty = 0, MTA_cpty = 10, IM_cpty = 0, remargin_freq = 1,
                         Values_type = "Actual")
collateral <- list(Trading::Collateral(ID = "vm", csa_id = "csa", Amount = 8, type = "VariationMargin"),
                   Trading::Collateral(ID = "nica", csa_id = "csa", Amount = 2, type = "ICA"))

values <- list(
    list(7, "SD, S=0, E=10", swap(0, 10)$CalcSupervDuration()),
    list(8, "SD, S=1, E=11", swap(1, 11)$CalcSupervDuration()),
    list(10, "MF unmargined, M=0.75", swap(0, 0.75)$CalcMaturityFactor()),
    # remargin_freq 1 gives the package's MPOR of 10 + 1 - 1 = 10 business days.
    list(12, "MF margined, MPOR=10 BD", margined$CalcMF(simplified = FALSE)),
    list(13, "delta, bought put P=6% K=5% T=1 vol=50%", swaption(1, 0.06, 0.05, "Put")$CalcSupervDelta(0.5)),
    list(14, "delta, sold call P=100 K=110 T=0.5 vol=32%",
         Trading::EquityOptionSingle(Notional = 1, MtM = 0, Si = 0, Ei = 0.5, BuySell = "Sell", OptionType = "Call",
                                     UnderlyingPrice = 100, StrikePrice = 110)$CalcSupervDelta(0.32)),
    list(16, "CDO delta, 3-6% tranche, protection bought",
         Trading::CDOTranche(Notional = 1, MtM = 0, Si = 0, Ei = 5, BuySell = "Buy", cdo_attach_point = 0.03,
                             cdo_detach_point = 0.06)$CalcSupervDelta()),
    list(19, "multiplier, V-C=-20, add-on=282", SACCR:::CalcPFE(-20, Addon_Aggregate = 282, simplified = FALSE) / 282),
    list(20, "multiplier, V-C=10, add-on=282", SACCR:::CalcPFE(10, Addon_Aggregate = 282, simplified = FALSE) / 282),
    list(21, "RC margined, V=5, VM=8, NICA=2, TH=0, MTA=10",
         SACCR:::CalcRC(list(list(MtM = 5)), margined, collateral, simplified = FALSE)$RC),
    # The package's deltas use R's normal distribution.
    list(22, "normal CDF(1.96)", pnorm(1.96)),
    list(23, "normal CDF(-4)", pnorm(-4)),
    # The package shifts by its CRR lambda when P and K have opposite signs.
    list(24, "delta with the CRR lambda, bought call P=-0.2% K=0.1% T=2 vol=50%",
         swaption(2, -0.002, 0.001, "Call")$CalcSupervDelta(0.5)),
    list(25, "delta with no shift, bought call P=2% K=3% T=2 vol=50%",
         swaption(2, 0.02, 0.03, "Call")$CalcSupervDelta(0.5))
)

cat(sprintf("R %s, SACCR %s, Trading %s\n", getRversion(), packageVersion("SACCR"), packageVersion("Trading")))
for (v in values) cat(sprintf("%d\t%s\t%.17g\n", v[[1]], v[[2]], v[[3]]))
