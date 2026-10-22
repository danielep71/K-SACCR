# Run the CRAN package SACCR on the netting sets written by
# tools/rsaccr_crosscheck.py and write its results as JSON (#352).
#
# Usage: Rscript tools/rsaccr_run.R <inputs.json> <results.json>
#
# Each netting set is one counterparty, run first without a CSA. A margined
# netting set is then run again with its CSA and collateral; the package's own
# cap ignores NICA, so the script rebuilds the cap of the margined EAD from the
# unmargined run with C = NICA, as the engine does. The package is called, not
# copied: it is GPL-3 and this repository is MPL-2.0.

suppressPackageStartupMessages({
    library(SACCR)
    library(jsonlite)
})

args <- commandArgs(trailingOnly = TRUE)
if (length(args) != 2) stop("usage: Rscript tools/rsaccr_run.R <inputs.json> <results.json>")
inputs <- fromJSON(args[[1]], simplifyVector = FALSE)

make_trade <- function(trade) {
    constructor <- getExportedValue("Trading", trade$class)
    do.call(constructor, trade$args)
}

make_csa <- function(csa) {
    Trading::CSA(ID = csa$id, thres_cpty = csa$threshold, thres_PO = 0, MTA_cpty = csa$mta, MTA_PO = 0,
                 IM_cpty = csa$nica, IM_PO = 0, mpor_days = csa$mpor_days, remargin_freq = csa$remargin_freq,
                 Counterparty = csa$counterparty, Currency = as.list(unlist(csa$currencies)),
                 TradeGroups = as.list(unlist(csa$trade_groups)), Values_type = "Actual")
}

make_collateral <- function(coll) {
    Trading::Collateral(ID = coll$id, Amount = coll$amount, csa_id = coll$csa_id, type = coll$type)
}

# runExampleCalcs returns a list holding one tree per CSA and one for the
# unmargined trades; a netting set here fills exactly one of them.
first_tree <- function(result) {
    for (item in result) if (inherits(item, "Node")) return(item)
    stop("the package returned no exposure tree")
}

# The package stores the supervisory factor on the hedging set or reference
# entity above a trade, so each trade takes the nearest one above it.
trade_details <- function(tree) {
    details <- list()
    walk <- function(node, factor) {
        if (!is.null(node$supervisory_factor)) factor <- node$supervisory_factor
        if (!is.null(node$exposure_details)) {
            e <- node$exposure_details
            details[[node$name]] <<- list(adjusted_notional = e$AdjNotional, supervisory_delta = e$superv_delta,
                                          maturity_factor = e$maturity_factor,
                                          effective_notional = e$effective_notional, supervisory_factor = factor)
        }
        for (child in node$children) walk(child, factor)
    }
    walk(tree, NULL)
    details
}

summarise <- function(tree) {
    rc <- tree$`Replacement Cost`
    classes <- list()
    for (node in tree$`Asset Classes`$children) classes[[node$name]] <- node$addon
    list(V = rc$V, V_C = rc$V_C, RC = rc$RC, add_on = tree$addon, add_on_by_class = classes,
         multiplier = if (tree$addon > 0) tree$PFE / tree$addon else 1, PFE = tree$PFE, EAD = tree$EAD,
         trades = trade_details(tree))
}

results <- list()
for (ns in inputs$netting_sets) {
    trades <- lapply(ns$trades, make_trade)
    unmargined <- summarise(first_tree(SACCR:::runExampleCalcs(trades, list(), list())))
    if (is.null(ns$csa)) {
        summary <- unmargined
    } else {
        csa <- make_csa(ns$csa)
        colls <- lapply(ns$collateral, make_collateral)
        summary <- summarise(first_tree(SACCR:::runExampleCalcs(trades, list(csa), colls)))
        # The package caps the margined EAD with an unmargined EAD that
        # ignores NICA; the engine's cap keeps C = NICA (CRR Art. 274(3),
        # 275(1); CRE52.2, CRE52.10). The cap is rebuilt here that way.
        nica <- if (is.null(ns$csa$nica)) 0 else ns$csa$nica
        margined_EAD <- CalcEAD(summary$RC, summary$PFE)
        cap_PFE <- CalcPFE(unmargined$V - nica, unmargined$V, unmargined$add_on, simplified = FALSE)
        summary$unmargined_EAD <- CalcEAD(max(unmargined$V - nica, 0), cap_PFE)
        summary$margined_EAD <- margined_EAD
        summary$EAD <- min(margined_EAD, summary$unmargined_EAD)
        summary$cap_applied <- summary$unmargined_EAD < margined_EAD
        summary$maturity_factor_margined <- csa$CalcMF(simplified = FALSE)
        summary$mpor_business_days <- 10 + ns$csa$remargin_freq - 1
    }
    summary$fixture <- ns$fixture
    summary$netting_set <- ns$id
    results[[length(results) + 1]] <- summary
}

versions <- list(R = paste(R.version$major, R.version$minor, sep = "."))
for (p in c("SACCR", "Trading", "data.tree", "jsonlite")) versions[[p]] <- as.character(packageVersion(p))
write_json(list(versions = versions, results = results), args[[2]], auto_unbox = TRUE, digits = NA, pretty = TRUE)
