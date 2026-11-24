#!/usr/bin/env python3
"""Show why ORE's SA-CCR analytic is not registered as a source (#100).

Open Source Risk Engine (ORE) 1.8.17 has an SA-CCR analytic. This script
writes a small unmargined portfolio, runs ORE on it through its Python
package and checks four results against the Basel formulas:

1. the interest-rate add-on of a currency: ORE adds the signed effective
   notionals with no maturity buckets (CRE52.57), so a net payer gives a
   negative add-on and a negative exposure value;
2. credit: ORE has no SA-CCR treatment of a credit default swap, which it
   leaves out of the netting set, and its calculator leaves the credit
   add-on at zero (CRE52.60-.64);
3. the supervisory delta of an option: ORE swaps the volatility and the
   time in the formula (CRE52.40);
4. the netting set's value: ORE reports it in USD whatever the base
   currency, so with a EUR base RC and the multiplier mix two currencies.

Usage::

    python tools/ore_saccr_probe.py --examples <ORE Examples folder> --out <folder>
    python tools/ore_saccr_probe.py --out <folder> --no-run

``--examples`` is the ``Examples`` folder of the ORE source (OpenSourceRisk/Engine
on GitHub, v1.8.17): the script uses its ``Input`` market, curve and pricing
files and the counterparty file of ``CreditRisk/Input/SA-CCR``. ORE runs with
``import ORE`` (the PyPI package ``open-source-risk-engine``). ``--no-run``
reads the reports of an earlier run in ``<folder>/Output``. Exit status is 0
when every check finds the result recorded in docs/methodology/README.md, 1
otherwise. Not a gate: CI has no ORE.
"""

from __future__ import annotations

import argparse
import csv
import math
import os
import shutil
import sys
from pathlib import Path
from statistics import NormalDist

BASE = "EUR"
EUR_USD = 1.132337  # FX/RATE/EUR/USD of Examples/Input/market_20160205_flat.txt
TOLERANCE = 1e-6
NETTING_SETS = ("NS_IR", "NS_CR", "NS_FXO")


def swap(trade_id: str, netting_set: str, pay_floating: bool, end: str, notional: int) -> str:
    """A EUR fixed-for-EURIBOR swap starting on 1 March 2016."""
    def leg(kind: str, payer: bool, tenor: str, body: str) -> str:
        return (f"<LegData><LegType>{kind}</LegType><Payer>{str(payer).lower()}</Payer><Currency>EUR</Currency>"
                f"<Notionals><Notional>{notional}</Notional></Notionals>{body}<ScheduleData><Rules>"
                f"<StartDate>20160301</StartDate><EndDate>{end}</EndDate><Tenor>{tenor}</Tenor>"
                "<Calendar>TARGET</Calendar><Convention>MF</Convention><TermConvention>MF</TermConvention>"
                "<Rule>Forward</Rule></Rules></ScheduleData></LegData>")
    fixed = leg("Fixed", not pay_floating, "1Y", "<DayCounter>30/360</DayCounter><PaymentConvention>F</PaymentConvention>"
                "<FixedLegData><Rates><Rate>0.01</Rate></Rates></FixedLegData>")
    floating = leg("Floating", pay_floating, "6M", "<DayCounter>A360</DayCounter><PaymentConvention>MF</PaymentConvention>"
                   "<FloatingLegData><Index>EUR-EURIBOR-6M</Index><Spreads><Spread>0</Spread></Spreads>"
                   "<IsInArrears>false</IsInArrears><FixingDays>2</FixingDays></FloatingLegData>")
    return (f'<Trade id="{trade_id}"><TradeType>Swap</TradeType><Envelope><CounterParty>CPTY_A</CounterParty>'
            f"<NettingSetId>{netting_set}</NettingSetId><AdditionalFields/></Envelope>"
            f"<SwapData>{fixed}{floating}</SwapData></Trade>")


def portfolio() -> str:
    """Two EUR swaps in two maturity buckets, net payer; a CDS; a bought FX call."""
    cds = ('<Trade id="CR_CDS"><TradeType>CreditDefaultSwap</TradeType><Envelope><CounterParty>CPTY_A</CounterParty>'
           "<NettingSetId>NS_CR</NettingSetId><AdditionalFields/></Envelope><CreditDefaultSwapData>"
           "<IssuerId>CPTY_C</IssuerId><CreditCurveId>CPTY_C</CreditCurveId><SettlesAccrual>Y</SettlesAccrual>"
           "<ProtectionPaymentTime>atDefault</ProtectionPaymentTime><LegData><LegType>Fixed</LegType>"
           "<Payer>true</Payer><Currency>EUR</Currency><Notionals><Notional>10000000</Notional></Notionals>"
           "<DayCounter>A360</DayCounter><PaymentConvention>F</PaymentConvention><FixedLegData><Rates>"
           "<Rate>0.01</Rate></Rates></FixedLegData><ScheduleData><Rules><StartDate>20160320</StartDate>"
           "<EndDate>20210320</EndDate><Tenor>3M</Tenor><Calendar>TARGET</Calendar><Convention>F</Convention>"
           "<Rule>CDS2015</Rule></Rules></ScheduleData></LegData></CreditDefaultSwapData></Trade>")
    option = ('<Trade id="FX_OPT"><TradeType>FxOption</TradeType><Envelope><CounterParty>CPTY_A</CounterParty>'
              "<NettingSetId>NS_FXO</NettingSetId><AdditionalFields/></Envelope><FxOptionData><OptionData>"
              "<LongShort>Long</LongShort><OptionType>Call</OptionType><Style>European</Style>"
              "<Settlement>Cash</Settlement><ExerciseDates><ExerciseDate>2017-02-06</ExerciseDate></ExerciseDates>"
              "</OptionData><BoughtCurrency>USD</BoughtCurrency><BoughtAmount>1200000</BoughtAmount>"
              "<SoldCurrency>EUR</SoldCurrency><SoldAmount>1000000</SoldAmount></FxOptionData></Trade>")
    trades = [swap("IR_A", "NS_IR", False, "20210301", 10_000_000), swap("IR_B", "NS_IR", True, "20180301", 30_000_000),
              cds, option]
    return '<?xml version="1.0"?>\n<Portfolio>\n' + "\n".join(trades) + "\n</Portfolio>\n"


def netting() -> str:
    sets = "".join(f"<NettingSet><NettingSetId>{n}</NettingSetId><ActiveCSAFlag>false</ActiveCSAFlag></NettingSet>"
                   for n in NETTING_SETS)
    return f'<?xml version="1.0"?>\n<NettingSetDefinitions>{sets}</NettingSetDefinitions>\n'


def parameters(examples: Path) -> str:
    inputs = (examples / "Input").resolve()
    files = {"marketDataFile": "market_20160205_flat.txt", "fixingDataFile": "fixings_20160205.txt",
             "curveConfigFile": "curveconfig.xml", "conventionsFile": "conventions.xml",
             "marketConfigFile": "todaysmarket.xml", "pricingEnginesFile": "pricingengine.xml"}
    setup = "".join(f'<Parameter name="{k}">{inputs / v}</Parameter>' for k, v in files.items())
    markets = "".join(f'<Parameter name="{k}">{v}</Parameter>' for k, v in
                      (("lgmcalibration", "collateral_inccy"), ("fxcalibration", "xois_eur"), ("pricing", "xois_eur"),
                       ("simulation", "xois_eur"), ("sensitivity", "xois_eur")))
    return ('<?xml version="1.0"?>\n<ORE><Setup><Parameter name="asofDate">2016-02-05</Parameter>'
            '<Parameter name="inputPath">Input</Parameter><Parameter name="outputPath">Output</Parameter>'
            '<Parameter name="logFile">log.txt</Parameter><Parameter name="logMask">31</Parameter>'
            f'{setup}<Parameter name="implyTodaysFixings">N</Parameter>'
            '<Parameter name="portfolioFile">portfolio.xml</Parameter>'
            '<Parameter name="counterpartyFile">counterparty.xml</Parameter>'
            '<Parameter name="observationModel">Disable</Parameter><Parameter name="nThreads">1</Parameter>'
            f"</Setup><Markets>{markets}</Markets><Analytics>"
            f'<Analytic type="npv"><Parameter name="active">Y</Parameter><Parameter name="baseCurrency">{BASE}'
            '</Parameter><Parameter name="outputFileName">npv.csv</Parameter></Analytic>'
            '<Analytic type="saccr"><Parameter name="active">Y</Parameter><Parameter name="simmVersion">2.1</Parameter>'
            '<Parameter name="csaFile">netting.xml</Parameter></Analytic></Analytics></ORE>\n')


def write_inputs(examples: Path, out: Path) -> None:
    folder = out / "Input"
    folder.mkdir(parents=True, exist_ok=True)
    (folder / "portfolio.xml").write_text(portfolio(), encoding="utf-8")
    (folder / "netting.xml").write_text(netting(), encoding="utf-8")
    (folder / "ore.xml").write_text(parameters(examples), encoding="utf-8")
    shutil.copyfile(examples / "CreditRisk" / "Input" / "SA-CCR" / "counterparty.xml", folder / "counterparty.xml")


def run_ore(out: Path) -> None:
    import ORE  # type: ignore[import-not-found]  # noqa: PLC0415 - optional, outside CI
    here = Path.cwd()
    try:
        os.chdir(out)
        params = ORE.Parameters()
        params.fromFile("Input/ore.xml")
        ORE.OREApp(params).run()
    finally:
        os.chdir(here)


def read_report(path: Path) -> list[dict[str, str]]:
    """An ORE CSV report, whose first header starts with '#'."""
    with path.open(encoding="utf-8", newline="") as handle:
        rows = list(csv.reader(handle))
    header = [h.lstrip("#") for h in rows[0]]
    return [dict(zip(header, row, strict=False)) for row in rows[1:]]


def bcbs_delta(price: float, strike: float, years: float, vol: float) -> float:
    """CRE52.40: delta of a bought call."""
    return NormalDist().cdf((math.log(price / strike) + 0.5 * vol * vol * years) / (vol * math.sqrt(years)))


def checks(saccr: list[dict[str, str]], detail: list[dict[str, str]],
           npv: list[dict[str, str]]) -> list[tuple[str, bool, str]]:
    """Each check: its name, whether ORE shows the recorded defect, and the figures."""
    total = {(r["NettingSet"], r["AssetClass"]): r for r in saccr if r["HedgingSet"] == "All"}
    trades = {r["TradeId"]: r for r in detail}
    found: list[tuple[str, bool, str]] = []

    # 1. Interest rate: buckets 1-3 with the CRE52.57 correlations, absolute value of the result.
    buckets = {"1": 0.0, "2": 0.0, "3": 0.0}
    for trade_id in ("IR_A", "IR_B"):
        row = trades[trade_id]
        buckets[row["Bucket"]] += float(row["delta"]) * float(row["d"]) * float(row["MF"])
    d1, d2, d3 = buckets["1"], buckets["2"], buckets["3"]
    basel = 0.005 * math.sqrt(d1 * d1 + d2 * d2 + d3 * d3 + 1.4 * d1 * d2 + 1.4 * d2 * d3 + 0.6 * d1 * d3)
    ore = float(total[("NS_IR", "IR")]["AddOn"])
    found.append(("interest-rate add-on without buckets, signed", ore < 0 and abs(ore - basel) > 1,
                  f"ORE {ore:.2f}, CRE52.57 {basel:.2f}, ORE EAD {total[('NS_IR', 'All')]['EAD']}"))

    # 2. Credit: a CDS has an adjusted notional and a positive add-on; ORE leaves it out.
    credit = total.get(("NS_CR", "All"), {})
    left_out = "CR_CDS" not in trades and float(credit.get("EAD", "nan")) == 0.0
    found.append(("credit default swap left out, credit add-on zero", left_out,
                  f"CDS in the trade report: {'CR_CDS' in trades}; netting set EAD {credit.get('EAD')}"))

    # 3. Option delta: ORE's figure against CRE52.40 and against the formula with vol and T swapped.
    option = trades["FX_OPT"]
    price, strike, years, vol = (float(option[k]) for k in ("Price", "Strike", "T", "Volatility"))
    swapped = NormalDist().cdf((math.log(price / strike) + 0.5 * years * years * vol) / (years * math.sqrt(vol)))
    delta = float(option["delta"])
    found.append(("option delta with volatility and time swapped",
                  abs(delta - swapped) < 1e-4 and abs(delta - bcbs_delta(price, strike, years, vol)) > 1e-3,
                  f"ORE {delta}, CRE52.40 {bcbs_delta(price, strike, years, vol):.4f}, swapped {swapped:.4f}"))

    # 4. Netting-set value: the trades' NPV in the base currency, against ORE's netting-set NPV.
    base = sum(float(r["NPV(Base)"]) for r in npv if r["NettingSet"] == "NS_FXO")
    reported = float(total[("NS_FXO", "All")]["NPV"])
    found.append(("netting-set NPV in USD, not the base currency", abs(reported - base * EUR_USD) < 1e-2 * abs(base),
                  f"ORE {reported:.2f}, trades in {BASE} {base:.2f}, ratio {reported / base:.6f}"))
    return found


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__.split("\n\n", 1)[0])
    parser.add_argument("--examples", type=Path, help="the Examples folder of the ORE 1.8.17 source")
    parser.add_argument("--out", type=Path, required=True, help="folder for the inputs and ORE's reports")
    parser.add_argument("--no-run", action="store_true", help="read the reports of an earlier run")
    options = parser.parse_args(argv)
    if not options.no_run:
        if options.examples is None:
            parser.error("--examples is required unless --no-run")
        write_inputs(options.examples, options.out)
        run_ore(options.out)
    reports = options.out / "Output"
    results = checks(read_report(reports / "saccr.csv"), read_report(reports / "saccr_detail.csv"),
                     read_report(reports / "npv.csv"))
    for name, defect, figures in results:
        print(f"{'FOUND' if defect else 'NOT FOUND':9}  {name}: {figures}")
    return 0 if all(defect for _, defect, _ in results) else 1


if __name__ == "__main__":
    sys.exit(main())
