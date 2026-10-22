#!/usr/bin/env python3
"""Check the numerical test cases in tests/fixtures and tests/expected against
docs/methodology/TEST_CASES.md, without running Excel."""
from __future__ import annotations

import json
import math
import re
import sys
from datetime import date
from pathlib import Path
from typing import Any

from _gatelib import parse_report_args, run_gate

FIXTURES = "tests/fixtures"
EXPECTED = "tests/expected"
REGISTER = "docs/methodology/README.md"
CATEGORIES = {"nominal", "boundary", "invalid", "regime"}
REGIMES = {"CRR", "BCBS"}
ASSET_CLASSES = ("interest_rate", "foreign_exchange", "credit", "equity", "commodity", "other")
CLASSES = {"published", "independent", "illustrative"}
CASE_ID = re.compile(r"^[a-z0-9]+(?:-[a-z0-9]+)*$")
CURRENCY = re.compile(r"^[A-Z]{3}$")
RULE = re.compile(r"^(?:CRR(?:-RTS)?\.[0-9]+[a-z]?(?:\.[0-9]+)*|CRE52(?:\.[0-9]+)?|CRE54\.[0-9]+)$")
ERROR_NAME = re.compile(r"^SACCR_ERROR_[A-Z0-9_]+$")
CATALOGUE = re.compile(r"^T[0-9]{2}$")
# A registered EBA Q&A is cited by its source ID, not in prose (#129). The
# scan covers the expected files, the VBA source and the methodology.
QA_SOURCE_ID = re.compile(r"^EBA-QA-([0-9]{4}_[0-9]+)$")
CITATION_FILES = ("tests/expected/*.json", "src/**/*.bas", "src/**/*.cls", "tests/modules/*.bas",
                  "docs/methodology/README.md", "docs/methodology/TEST_CASES.md")

# Field name -> (kind, nullable). The kind follows the unit suffix of the name;
# the two business-day fields have the engine's ranges for their column on
# the NettingSets sheet (CORE_Engine.LoadNettingSets, #147, #161).
NETTING_SET_FIELDS: dict[str, tuple[str, bool]] = {
    "id": ("text", False),
    "margined": ("bool", False),
    "clearing_role": ("clearing_role", False),
    "remargin_period_business_days": ("remargin_business_days", True),
    "large_or_illiquid": ("bool", False),
    "margin_disputes": ("bool", False),
    "mpor_override_business_days": ("mpor_business_days", True),
    "variation_margin_net_amount": ("number", False),
    "independent_collateral_net_amount": ("number", False),
    "threshold_amount": ("non_negative", False),
    "minimum_transfer_amount": ("non_negative", False),
    "alpha_factor": ("positive", True),
    "qccp": ("bool", True),
    "ccp_retains_variation_margin": ("bool", True),
}
TRADE_FIELDS: dict[str, tuple[str, bool]] = {
    "id": ("text", False),
    "asset_class": ("asset_class", False),
    "sub_class": ("text", True),
    "risk_factor": ("text", False),
    "instrument": ("instrument", False),
    "direction": ("direction", False),
    "option_type": ("option_type", True),
    "nature": ("nature", False),
    "hedging_set_label": ("text", True),
    "notional_amount": ("non_negative", False),
    "market_value_amount": ("number", False),
    "start_date": ("date", True),
    "end_date": ("date", True),
    "maturity_date": ("date", True),
    "option_expiry_date": ("date", True),
    "underlying_price": ("number", True),
    "strike_price": ("number", True),
    "lambda_price": ("number", True),
    "attachment_rate": ("number", True),
    "detachment_rate": ("number", True),
    "settlement_type": ("settlement_type", False),
    "next_reset_date": ("date", True),
    "description": ("text", True),
}
# Params codes a fixture may set for its own run -> kind of the value, with
# the engine's range for that parameter (CORE_Engine.LoadParams).
PARAMETERS: dict[str, str] = {
    "DaysPerYear": "days_per_year",
    "IRBucketOffset": "bool",
}
ENUMS = {
    "asset_class": set(ASSET_CLASSES),
    "instrument": {"linear", "option", "cdo"},
    "direction": {"long", "short"},
    "option_type": {"call", "put"},
    "nature": {"standard", "basis", "volatility"},
    "clearing_role": {"none", "clearing_member", "client", "ccp"},
    "settlement_type": {"ordinary", "settled_to_market"},
}

NETTING_SET_QUANTITIES = {
    "exposure_value", "replacement_cost", "potential_future_exposure", "multiplier",
    "aggregate_add_on", "margin_period_of_risk", "cap_applied", "netting_set_status",
    *(f"add_on.{name}" for name in ASSET_CLASSES),
}
TRADE_QUANTITIES = {
    "adjusted_notional", "supervisory_delta", "supervisory_factor", "lambda_shift",
    "hedging_set", "bucket", "trade_status", "maturity_factor", "maturity_factor_margined",
    "remaining_maturity",
}
TEXT_QUANTITIES = {"cap_applied", "hedging_set", "bucket", "trade_status", "netting_set_status"}
REFERENCE_FIELDS = {"class", "source", "locator", "derivation", "derived_by", "reviewed_by", "date"}
REQUIRED_REFERENCE = {
    "published": ("source", "locator"),
    "independent": ("source", "locator", "derivation", "derived_by"),
    "illustrative": ("derivation",),
}


def is_number(value: object) -> bool:
    return (isinstance(value, int) and not isinstance(value, bool)) or (
        isinstance(value, float) and math.isfinite(value))


def is_whole_number(value: object) -> bool:
    if isinstance(value, float):
        return math.isfinite(value) and value.is_integer()
    return is_number(value)


def reject_json_constant(value: str) -> None:
    raise ValueError(f"non-finite JSON constant {value}")


def finite_json_float(value: str) -> float:
    number = float(value)
    if not math.isfinite(number):
        raise ValueError(f"non-finite JSON number {value}")
    return number


def valid_date(value: object) -> bool:
    if not isinstance(value, str) or not re.fullmatch(r"\d{4}-\d{2}-\d{2}", value):
        return False
    try:
        date.fromisoformat(value)
    except ValueError:
        return False
    return True


def check_value(where: str, kind: str, value: object) -> list[str]:
    """Return a finding when value does not match its field kind."""
    checks = {
        "text": lambda v: isinstance(v, str) and bool(v.strip()),
        "bool": lambda v: isinstance(v, bool),
        "number": is_number,
        "non_negative": lambda v: is_number(v) and v >= 0,
        "positive": lambda v: is_number(v) and v > 0,
        # Remargining: a whole number from 0 to 10000; the engine raises 0 to 1.
        "remargin_business_days": lambda v: is_whole_number(v) and 0 <= v <= 10000,
        # MPOR override: a whole number above 0; null, not 0, means none.
        "mpor_business_days": lambda v: is_whole_number(v) and v >= 1,
        "days_per_year": lambda v: is_number(v) and 360 <= v <= 366,
        "date": valid_date,
    }
    if kind in ENUMS:
        ok = value in ENUMS[kind]
    else:
        ok = checks[kind](value)
    return [] if ok else [f"{where}: invalid {kind} value {value!r}"]


def check_fields(where: str, record: object, spec: dict[str, tuple[str, bool]]) -> list[str]:
    """Every field of the record must be declared, present and well typed."""
    if not isinstance(record, dict):
        return [f"{where}: must be an object"]
    findings = [f"{where}: unknown field '{key}'" for key in sorted(set(record) - set(spec))]
    findings += [f"{where}: missing field '{key}'" for key in sorted(set(spec) - set(record))]
    for key, (kind, nullable) in spec.items():
        if key in record and not (nullable and record[key] is None):
            findings += check_value(f"{where}.{key}", kind, record[key])
    return findings


def load(path: Path, findings: list[str], name: str) -> dict[str, Any] | None:
    try:
        data = json.loads(path.read_text(encoding="utf-8"), parse_constant=reject_json_constant,
                          parse_float=finite_json_float)
    except (UnicodeDecodeError, ValueError) as error:
        findings.append(f"{name}: not valid UTF-8 JSON ({error})")
        return None
    if not isinstance(data, dict):
        findings.append(f"{name}: top level must be an object")
        return None
    return data


def check_fixture(name: str, stem: str, data: dict[str, Any]) -> list[str]:
    expected_keys = {"schema_version", "kind", "id", "synthetic", "description", "category",
                     "calculation_currency", "valuation_date", "parameters", "netting_set", "trades"}
    findings = [f"{name}: unknown field '{k}'" for k in sorted(set(data) - expected_keys)]
    findings += [f"{name}: missing field '{k}'" for k in sorted(expected_keys - set(data))]
    if data.get("schema_version") != 1 or data.get("kind") != "saccr-fixture":
        findings.append(f"{name}: schema_version must be 1 and kind 'saccr-fixture'")
    if data.get("id") != stem or not CASE_ID.match(stem):
        findings.append(f"{name}: id must equal the lowercase kebab-case file name")
    if data.get("synthetic") is not True:
        findings.append(f"{name}: synthetic must be true")
    findings += check_value(f"{name}.description", "text", data.get("description"))
    if data.get("category") not in CATEGORIES:
        findings.append(f"{name}: category must be one of {sorted(CATEGORIES)}")
    if not CURRENCY.match(str(data.get("calculation_currency"))):
        findings.append(f"{name}: calculation_currency must be a three-letter code")
    findings += check_value(f"{name}.valuation_date", "date", data.get("valuation_date"))
    findings += check_parameters(f"{name}.parameters", data.get("parameters", {}))
    findings += check_fields(f"{name}.netting_set", data.get("netting_set"), NETTING_SET_FIELDS)
    trades = data.get("trades")
    if not isinstance(trades, list) or not trades:
        return findings + [f"{name}: trades must be a non-empty array"]
    ids = [str(t.get("id")) for t in trades if isinstance(t, dict)]
    findings += [f"{name}: duplicate trade id '{i}'" for i in sorted({i for i in ids if ids.count(i) > 1})]
    for index, trade in enumerate(trades):
        findings += check_fields(f"{name}.trades[{index}]", trade, TRADE_FIELDS)
    return findings


def iso_date(value: object) -> date | None:
    """The date of a well-formed field, or None; a malformed one is reported by check_fields."""
    return date.fromisoformat(value) if isinstance(value, str) and valid_date(value) else None


def input_errors(data: dict[str, Any]) -> tuple[dict[str, list[str]], list[str]]:
    """The settlement and CCP inputs the engine rejects, as the engine checks them:
    per trade (the trade is EXCLUDED; CORE_Engine.ReadTradeSettlement, #360) and
    per netting set (it is INVALID; CORE_Engine.LoadNettingSets, #112). Each field
    is validated alone by check_fields; these rules relate fields to each other."""
    trade_errors: dict[str, list[str]] = {}
    set_errors: list[str] = []
    netting_set = data.get("netting_set")
    netting_set = netting_set if isinstance(netting_set, dict) else {}
    margined = netting_set.get("margined") is True
    valuation = iso_date(data.get("valuation_date"))
    trades = data.get("trades")
    for trade in trades if isinstance(trades, list) else []:
        if not isinstance(trade, dict):
            continue
        errors: list[str] = []
        settled = trade.get("settlement_type") == "settled_to_market"
        reset_value = trade.get("next_reset_date")
        reset = iso_date(reset_value)
        # M is the maturity date, or else the end date, or else the option expiry.
        maturity = next((d for d in (iso_date(trade.get(k)) for k in
                                     ("maturity_date", "end_date", "option_expiry_date")) if d), None)
        if not settled:
            if reset_value is not None:
                errors.append("next_reset_date is given but settlement_type is not settled_to_market")
        elif reset_value is None:
            errors.append("a settled_to_market trade needs next_reset_date")
        elif reset is not None and valuation is not None and reset <= valuation:
            errors.append("next_reset_date must be after valuation_date")
        elif reset is not None and maturity is not None and reset > maturity:
            errors.append("next_reset_date is after the maturity (maturity_date, else end_date, "
                          "else option_expiry_date)")
        if errors:
            trade_errors[str(trade.get("id"))] = errors
        if settled and margined:
            set_errors.append(f"trade {trade.get('id')} is settled_to_market in a margined netting set")
    ccp = netting_set.get("clearing_role") == "ccp"
    for field in ("qccp", "ccp_retains_variation_margin"):
        if field not in netting_set:
            continue
        if ccp and netting_set[field] is None:
            set_errors.append(f"{field} is required for clearing_role ccp")
        elif not ccp and netting_set[field] is not None:
            set_errors.append(f"{field} applies only to clearing_role ccp and must be null")
    if ccp and netting_set.get("ccp_retains_variation_margin") is True and not margined:
        set_errors.append("ccp_retains_variation_margin needs a margined netting set")
    return trade_errors, set_errors


def check_declared_rejections(case: str, data: dict[str, Any],
                              expected: list[tuple[str, list[dict[str, Any]]]]) -> list[str]:
    """A fixture holds valid inputs, except where it tests a rejection: every
    expected file of the fixture must then declare the outcome the engine
    gives, trade_status EXCLUDED for the trade or a netting_set_status
    starting with INVALID. A rejected input without that declaration is a
    malformed fixture, whose figures would not describe the calculation (#401)."""
    trade_errors, set_errors = input_errors(data)
    findings: list[str] = []
    for name, outputs in expected:
        excluded = {o.get("trade") for o in outputs
                    if o.get("quantity") == "trade_status" and o.get("text") == "EXCLUDED"}
        invalid = any(o.get("quantity") == "netting_set_status" and str(o.get("text", "")).startswith("INVALID")
                      for o in outputs)
        for trade, errors in sorted(trade_errors.items()):
            if trade not in excluded:
                findings += [f"{FIXTURES}/{case}.json: trade {trade}: {e}; the engine rejects it, but {name} "
                             f"does not expect trade_status EXCLUDED" for e in errors]
        if not invalid:
            findings += [f"{FIXTURES}/{case}.json: {e}; the engine makes the netting set INVALID, but {name} "
                         f"does not expect a netting_set_status starting with INVALID" for e in set_errors]
    return findings


def check_parameters(where: str, parameters: object) -> list[str]:
    """Params values a case runs with instead of the template's; {} for none."""
    if not isinstance(parameters, dict):
        return [f"{where}: must be an object"]
    findings = [f"{where}: unknown parameter '{k}'" for k in sorted(set(parameters) - set(PARAMETERS))]
    for key in sorted(set(parameters) & set(PARAMETERS)):
        findings += check_value(f"{where}.{key}", PARAMETERS[key], parameters[key])
    return findings


def quantity_unit(output: dict[str, Any], fixture: dict[str, Any]) -> str:
    """Resolve the documented unit, including the lambda trade's underlying."""
    quantity = output.get("quantity")
    if quantity in {"multiplier", "supervisory_delta", "supervisory_factor", "maturity_factor",
                    "maturity_factor_margined"}:
        return "1"
    if quantity == "remaining_maturity":
        return "years"
    if quantity == "margin_period_of_risk":
        return "business_days"
    if quantity == "lambda_shift":
        trade = next((t for t in fixture.get("trades", [])
                      if isinstance(t, dict) and t.get("id") == output.get("trade")), {})
        return "1" if trade.get("asset_class") == "interest_rate" else "price"
    return str(fixture.get("calculation_currency"))


def check_measure(where: str, output: dict[str, Any], fixture: dict[str, Any]) -> list[str]:
    """An output carries exactly one of value (with unit and tolerance), text or error."""
    forms = [key for key in ("value", "text", "error") if key in output]
    if len(forms) != 1:
        return [f"{where}: needs exactly one of value, text or error"]
    quantity, form = output.get("quantity"), forms[0]
    if quantity in TEXT_QUANTITIES and form != "text":
        return [f"{where}: {quantity} is a text quantity"]
    if quantity not in TEXT_QUANTITIES and form == "text":
        return [f"{where}: {quantity} is not a text quantity"]
    if form == "text":
        stray = {"unit", "tolerance"} & set(output)
        return [f"{where}: a text result has no unit or tolerance"] if stray else check_value(
            f"{where}.text", "text", output["text"])
    if form == "error":
        stray = {"unit", "tolerance"} & set(output)
        findings = [f"{where}: an error result has no unit or tolerance"] if stray else []
        if not ERROR_NAME.match(str(output["error"])):
            findings.append(f"{where}: error must name a SACCR_ERROR_* constant")
        return findings
    findings = check_value(f"{where}.value", "number", output["value"])
    unit = output.get("unit")
    required_unit = quantity_unit(output, fixture)
    if unit != required_unit:
        findings.append(f"{where}: unit must be {required_unit} for {quantity}")
    tolerance = output.get("tolerance")
    if not isinstance(tolerance, dict) or set(tolerance) != {"absolute", "relative"} or not all(
            is_number(v) and v >= 0 for v in tolerance.values()):
        findings.append(f"{where}: tolerance needs non-negative 'absolute' and 'relative'")
    return findings


def check_reference(where: str, reference: object, register: set[str]) -> list[str]:
    if not isinstance(reference, dict):
        return [f"{where}: reference must be an object"]
    findings = [f"{where}: unknown reference field '{k}'" for k in sorted(set(reference) - REFERENCE_FIELDS)]
    kind = reference.get("class")
    if kind not in CLASSES:
        return findings + [f"{where}: class must be one of {sorted(CLASSES)}"]
    for key in REQUIRED_REFERENCE[str(kind)]:
        value = reference.get(key)
        if not isinstance(value, str) or not value.strip():
            findings.append(f"{where}: a {kind} reference needs '{key}'")
    source = reference.get("source")
    if source is not None and source not in register:
        findings.append(f"{where}: source '{source}' is not in the source register")
    if "date" in reference and not valid_date(reference["date"]):
        findings.append(f"{where}: date must be YYYY-MM-DD")
    return findings


def check_output(where: str, output: object, fixture: dict[str, Any], register: set[str]) -> list[str]:
    if not isinstance(output, dict):
        return [f"{where}: must be an object"]
    allowed = {"quantity", "rule", "trade", "value", "unit", "tolerance", "text", "error",
               "reference", "catalogue"}
    findings = [f"{where}: unknown field '{k}'" for k in sorted(set(output) - allowed)]
    quantity = output.get("quantity")
    if quantity not in NETTING_SET_QUANTITIES | TRADE_QUANTITIES:
        findings.append(f"{where}: unknown quantity {quantity!r}")
    if not RULE.match(str(output.get("rule"))):
        findings.append(f"{where}: rule must look like CRR.275.1, CRR-RTS.5, CRE52.23 or CRE54.12")
    trade_ids = {t.get("id") for t in fixture.get("trades", []) if isinstance(t, dict)}
    if quantity in TRADE_QUANTITIES and output.get("trade") not in trade_ids:
        findings.append(f"{where}: {quantity} needs 'trade' naming a trade of the fixture")
    if quantity in NETTING_SET_QUANTITIES and "trade" in output:
        findings.append(f"{where}: {quantity} is a netting-set quantity and takes no 'trade'")
    # TradeCalc reports the margined maturity factor separately; the
    # quantity is the unmargined one, so it is checked only where it applies.
    if quantity == "maturity_factor" and (fixture.get("netting_set") or {}).get("margined"):
        findings.append(f"{where}: maturity_factor is the unmargined maturity factor; the netting set is margined")
    if quantity == "maturity_factor_margined" and not (fixture.get("netting_set") or {}).get("margined"):
        findings.append(f"{where}: maturity_factor_margined needs a margined netting set")
    if "catalogue" in output and not CATALOGUE.match(str(output["catalogue"])):
        findings.append(f"{where}: catalogue must look like T01")
    findings += check_measure(where, output, fixture)
    return findings + check_reference(f"{where}.reference", output.get("reference"), register)


def check_expected(name: str, stem: str, data: dict[str, Any], fixtures: dict[str, dict[str, Any]],
                   register: set[str]) -> tuple[list[str], list[dict[str, Any]]]:
    expected_keys = {"schema_version", "kind", "fixture", "regime", "outputs"}
    findings = [f"{name}: unknown field '{k}'" for k in sorted(set(data) - expected_keys)]
    if data.get("schema_version") != 1 or data.get("kind") != "saccr-expected":
        findings.append(f"{name}: schema_version must be 1 and kind 'saccr-expected'")
    case, _, regime = stem.rpartition(".")
    if data.get("fixture") != case or case not in fixtures:
        findings.append(f"{name}: fixture must name tests/fixtures/{case}.json")
    if data.get("regime") not in REGIMES or str(data.get("regime")).lower() != regime:
        findings.append(f"{name}: regime must be CRR or BCBS and match the file name")
    outputs = data.get("outputs")
    if not isinstance(outputs, list) or not outputs:
        return findings + [f"{name}: outputs must be a non-empty array"], []
    fixture = fixtures.get(case, {})
    for index, output in enumerate(outputs):
        findings += check_output(f"{name}.outputs[{index}]", output, fixture, register)
    return findings, [o for o in outputs if isinstance(o, dict)]


def read_register(root: Path) -> set[str]:
    """Source IDs from the source-register table of the methodology README."""
    text = (root / REGISTER).read_text(encoding="utf-8")
    section = text.split('<a id="source-register"></a>', 1)[-1].split("\n## ", 2)
    table = section[1] if len(section) > 1 else ""
    return set(re.findall(r"^\| `([A-Z0-9_-]+)` \|", table, re.M))


def join_lines(lines: list[str], vba: bool) -> tuple[str, list[int]]:
    """Join the lines of a file with single spaces, so that a citation
    wrapped over two lines is found (#145), and give the line number of each
    character. In VBA the comment mark that starts a line is left out, so
    that a citation continued on the next comment line reads as one."""
    parts: list[str] = []
    owner: list[int] = []
    for number, line in enumerate(lines, 1):
        content = line.strip()
        if vba and content.startswith("'"):
            content = content[1:].strip()
        parts.append(content)
        owner.extend([number] * (len(content) + 1))
    return " ".join(parts), owner


def check_citations(root: Path, register: set[str]) -> list[str]:
    """Flag a registered EBA Q&A cited in prose instead of by its source ID,
    in the short form (EBA Q&A 2023_6962) or the full name (EBA Single
    Rulebook Q&A 2023_6962) (#145). The source register's row that defines
    the ID is the one place the full name belongs, and only that row is
    exempt."""
    numbers = {m.group(1): sid for sid in sorted(register) if (m := QA_SOURCE_ID.match(sid))}
    if not numbers:
        return []
    alternatives = "|".join(map(re.escape, numbers))
    prose = re.compile(r"\bEBA\s+(?:Single\s+Rulebook\s+)?Q&A\s+(" + alternatives + r")(?![0-9])")
    defining_row = re.compile(r"^\| `EBA-QA-(" + alternatives + r")` \|")
    findings = []
    for pattern in CITATION_FILES:
        for path in sorted(root.glob(pattern)):
            text = path.read_bytes().decode("utf-8", errors="replace")
            in_register = path.relative_to(root).as_posix() == REGISTER
            lines = text.splitlines()
            joined, owner = join_lines(lines, path.suffix.lower() in (".bas", ".cls"))
            for match in prose.finditer(joined):
                number = owner[match.start()]
                defined = defining_row.match(lines[number - 1]) if in_register else None
                if defined and defined.group(1) == match.group(1) and owner[match.end() - 1] == number:
                    continue
                findings.append(f"{path.relative_to(root).as_posix()}:{number}: cite EBA Q&A "
                                f"{match.group(1)} by its source ID {numbers[match.group(1)]}")
    return findings


def check_cases(case_outputs: dict[str, list[dict[str, Any]]], fixtures: set[str]) -> list[str]:
    findings = [f"{FIXTURES}/{c}.json: no expected file" for c in sorted(fixtures - set(case_outputs))]
    catalogue: dict[str, str] = {}
    for case, outputs in sorted(case_outputs.items()):
        classes = {o.get("reference", {}).get("class") for o in outputs if isinstance(o.get("reference"), dict)}
        if case.startswith("illustrative-") and classes - {"illustrative"}:
            findings.append(f"{case}: an illustrative- case may hold only illustrative values")
        if not case.startswith("illustrative-") and classes == {"illustrative"}:
            findings.append(f"{case}: a case whose values are all illustrative must start with illustrative-")
        for output in outputs:
            test = output.get("catalogue")
            if isinstance(test, str):
                if test in catalogue:
                    findings.append(f"{case}: catalogue {test} also used by {catalogue[test]}")
                catalogue[test] = case
    return findings


def run_check(root: Path) -> dict[str, Any]:
    findings: list[str] = []
    register = read_register(root)
    if not register:
        findings.append(f"{REGISTER}: source register table not found")
    fixtures: dict[str, dict[str, Any]] = {}
    for path in sorted((root / FIXTURES).glob("*.json")):
        name = f"{FIXTURES}/{path.name}"
        data = load(path, findings, name)
        if data is not None:
            findings += check_fixture(name, path.stem, data)
            fixtures[path.stem] = data
    case_outputs: dict[str, list[dict[str, Any]]] = {}
    case_files: dict[str, list[tuple[str, list[dict[str, Any]]]]] = {}
    expected_files = sorted((root / EXPECTED).glob("*.json"))
    for path in expected_files:
        name = f"{EXPECTED}/{path.name}"
        data = load(path, findings, name)
        if data is not None:
            more, outputs = check_expected(name, path.stem, data, fixtures, register)
            findings += more
            case = path.stem.rpartition(".")[0]
            case_outputs.setdefault(case, []).extend(outputs)
            case_files.setdefault(case, []).append((name, outputs))
    for case, files in sorted(case_files.items()):
        if case in fixtures:
            findings += check_declared_rejections(case, fixtures[case], files)
    findings += check_cases(case_outputs, set(fixtures))
    findings += check_citations(root, register)
    outputs = [o for group in case_outputs.values() for o in group]
    counts = {kind: sum(1 for o in outputs if isinstance(o.get("reference"), dict)
                        and o["reference"].get("class") == kind) for kind in sorted(CLASSES)}
    return {"schema_version": 1, "status": "fail" if findings else "pass",
            "fixtures": len(fixtures), "expected_files": len(expected_files),
            "outputs": len(outputs), "classes": counts, "findings": findings}


def markdown_report(report: dict[str, Any]) -> str:
    classes = ", ".join(f"{k} {v}" for k, v in report["classes"].items())
    lines = [f"Test cases: {report['status'].upper()} ({report['fixtures']} fixture(s), "
             f"{report['expected_files']} expected file(s), {report['outputs']} output(s): {classes})"]
    lines.extend(report["findings"])
    return "\n".join(lines) + "\n"


def main() -> int:
    options = parse_report_args(sys.argv[1:], description=__doc__)
    return run_gate(options, build=lambda: run_check(options.root.resolve()),
                    markdown=markdown_report, errors=(OSError, ValueError))


if __name__ == "__main__":
    raise SystemExit(main())
