"""The ORE probe of #100: its inputs and the checks it makes on ORE's reports, without ORE."""
from __future__ import annotations

import tempfile
import unittest
import xml.etree.ElementTree as ET
from pathlib import Path

import ore_saccr_probe as probe

# The reports of ORE 1.8.17 on the probe's portfolio, trimmed to the rows the checks read.
SACCR = """#NettingSet,AssetClass,HedgingSet,AddOn,NPV,IndependentAmountHeld,InitialMargin,VariationMargin,\
ThresholdAmount,MinimumTransferAmount,RC,Multiplier,PFE,EAD,RW,CC
NS_CR,All,All,0.000000,0.000000,0.000000,0.000000,0.000000,0.000000,0.000000,0.000000,1.000000,0.000000,0.000000,\
1.000000,0.000000
NS_FXO,All,All,26351.770718,84031.905702,0.000000,0.000000,0.000000,0.000000,0.000000,84031.905702,1.000000,\
26351.770718,154537.146988,1.000000,154537.146988
NS_IR,All,All,-14371.019023,-124097.216172,0.000000,0.000000,0.000000,0.000000,0.000000,0.000000,1.000000,\
-14371.019023,-20119.426632,1.000000,-20119.426632
NS_IR,IR,All,-14371.019023,,,,,,,,,,,,
"""
DETAIL = """#TradeId,TradeType,NettingSet,NPV,AssetClass,HedgingSet,HedgingSubset,Bucket,Qualifier,Currency,delta,d,MF,M,\
S,E,T,SD,CurrentPrice,NumNominalFlows,Price,Strike,Volatility
FX_OPT,FxOption,NS_FXO,74211.04,FX,EURUSD,,,EURUSD,EUR,0.6216,1059755.1789,1.0000000,1.0030,#N/A,#N/A,1.0030,#N/A,\
0.871601,#N/A,0.8716,0.8333,0.1500
IR_A,Swap,NS_IR,469558.62,IR,EUR,,3,EUR-EURIBOR-6M,EUR,1.0000,39123492.7208,1.0000000,5.0660,0.5710,5.0660,#N/A,4,\
#N/A,#N/A,#N/A,#N/A,0.5000
IR_B,Swap,NS_IR,-579152.51,IR,EUR,,2,EUR-EURIBOR-6M,EUR,-1.0000,41997696.5254,1.0000000,2.0660,0.5710,2.0660,#N/A,1,\
#N/A,#N/A,#N/A,#N/A,0.5000
"""
NPV = """#TradeId,TradeType,Maturity,MaturityTime,NPV,NpvCurrency,NPV(Base),BaseCurrency,Notional,NotionalCurrency,\
Notional(Base),NettingSet,CounterParty
CR_CDS,CreditDefaultSwap,2021-03-22,5.123550,-505389.997647,EUR,-505389.997647,EUR,10000000.00,EUR,10000000.00,NS_CR,\
CPTY_A
FX_OPT,FxOption,2017-02-06,1.003002,74211.039383,EUR,74211.039383,EUR,1000000.00,EUR,1000000.00,NS_FXO,CPTY_A
"""


class ProbeTest(unittest.TestCase):
    def reports(self, saccr: str = SACCR, detail: str = DETAIL) -> list[tuple[str, bool, str]]:
        with tempfile.TemporaryDirectory() as folder:
            paths = []
            for name, text in (("saccr.csv", saccr), ("saccr_detail.csv", detail), ("npv.csv", NPV)):
                path = Path(folder) / name
                path.write_text(text, encoding="utf-8")
                paths.append(path)
            return probe.checks(*(probe.read_report(p) for p in paths))

    def test_recorded_reports_show_every_defect(self) -> None:
        results = self.reports()
        self.assertEqual(len(results), 4)
        self.assertTrue(all(defect for _, defect, _ in results), results)

    def test_basel_figures_are_not_taken_for_defects(self) -> None:
        # The CRE52.57 add-on of the two swaps, the CRE52.40 delta and an NPV in EUR.
        saccr = SACCR.replace("-14371.019023,,,", "157648.259718,,,").replace("84031.905702", "74211.039383")
        detail = DETAIL.replace("EUR,0.6216,", "EUR,0.6459,")
        results = dict((name, defect) for name, defect, _ in self.reports(saccr, detail))
        self.assertFalse(results["interest-rate add-on without buckets, signed"])
        self.assertFalse(results["option delta with volatility and time swapped"])
        self.assertFalse(results["netting-set NPV in USD, not the base currency"])

    def test_inputs_are_well_formed(self) -> None:
        portfolio = ET.fromstring(probe.portfolio())
        self.assertEqual([t.get("id") for t in portfolio], ["IR_A", "IR_B", "CR_CDS", "FX_OPT"])
        netting = ET.fromstring(probe.netting())
        self.assertEqual([n.findtext("ActiveCSAFlag") for n in netting], ["false"] * 3)
        setup = ET.fromstring(probe.parameters(Path("Examples")))
        self.assertEqual(setup.findtext("Setup/Parameter[@name='asofDate']"), "2016-02-05")


if __name__ == "__main__":
    unittest.main()
