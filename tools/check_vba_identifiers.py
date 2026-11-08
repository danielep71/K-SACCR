"""Advisory VBA identifier inventory; unknown names are not compiler verdicts."""
from __future__ import annotations

import json
import sys
from pathlib import Path
from typing import Any

from _gatelib import parse_report_args, run_gate, tracked_files
from check_vba_calls import KEYWORDS, STATEMENT_KEYWORDS, Module, Proc, parse_module
from check_vba_conditionals import logical_units, reachable_sources
from vba_lexer import Tok, tokenize

# Explicit language/runtime names, never a blanket vb*/xl* or capitalization exemption.
BUILTINS = frozenset("""
choose setattr objptr lenb vbdecimal vbmsgboxstyle boolean byte integer long longlong longptr single double currency decimal date string
variant object any collection debug err me vba excel office scripting
abs array asc ascw atn cbool cbyte ccur cdate cdbl cdec chr chrw cint clng clnglng clngptr
cos csng cstr cvar cvdate cverr dateadd datediff datepart dateserial datevalue day dir doevents
environ eof exp fileattr filedatetime filelen fix format freefile getallsettings getattr
getobject hex hour iif inputbox instr instrrev int isarray isdate isempty iserror ismissing
isnull isnumeric isobject join lbound lcase left len loadpicture loc lof log ltrim mid minute
month monthname msgbox now oct replace rgb right rnd round rtrim second sgn shell sin space
split sqr str strcomp strconv string strreverse switch tan time timeserial timer timevalue
trim typename ubound ucase val vartype weekday weekdayname year createobject callbyname
vbcr vbcrlf vblf vbtab vbnullchar vbnullstring vbtrue vbfalse vbuseDefault vbobjecterror
vbbinarycompare vbtextcompare vbuppercase vblowercase vbpropercase vbunicode vbfromunicode
vberror vbboolean vbbyte vbinteger vblong vbsingle vbdouble vbcurrency vbdate vbstring
vbobject vbvariant vbarray vbempty vbnull vbnormal vbreadonly vbhidden vbsystem vbdirectory
vbarchive vbokonly vbokcancel vbabortretryignore vbyesnocancel vbyesno vbretrycancel vbcritical
vbquestion vbexclamation vbinformation vbdefaultbutton1 vbdefaultbutton2 vbdefaultbutton3
vbok vbcancel vbabort vbretry vbignore vbyes vbno vbmethod vbget vblet vbset vbnewline
access append binary input output random read write shared encoding
""".lower().split())
# Host object roots/types: also used by the dependency gate to detect core -> Excel.
HOST_OBJECTS = frozenset("""
msgbox inputbox doevents sendkeys appactivate shell beep excel office application thisworkbook activeworkbook activesheet activecell selection workbooks worksheets
sheets range cells rows columns names workbook worksheet chart charts window windows
shape xlcalculation worksheetfunction evaluate union intersect usedrange iribboncontrol iribbonui
""".split())
HOST_CONSTANTS = frozenset("""
xlcalculationautomatic xlcalculationmanual xlcalculationsemiautomatic xldone xlup xldown
xltoleft xltoright xlnone xlcontinuous xlthin xlmedium xlthick xlcenter xlleft xlright
xltop xlbottom xlwhole xlpart xlbyrows xlbycolumns xlnext xlprevious xlvalues xlformulas
xlcelltypeconstants xlcelltypeformulas xlcelltypelastcell xlcelltypevisible xlerrvalue
xlerrna xlerrnum xlerrdiv0 xlerrref xlerrname xlerrnull xlworksheet xlwbatworksheet
xlopenxmlworkbook xlopenxmlworkbookmacroenabled xlfileformat xlveryhidden xlsheetveryhidden
xlsheethidden xlsheetvisible xla1 xlr1c1 xlascending xldescending xlyes xlno xlguess
xlfiltercopy xlfilterinplace xlsolid xlpatternnone xlvalidalertstop xlvalidatecustom
xlbetween xlqualitystandard xltypepdf msofiledialogfilepicker msofiledialogsaveas
msoautomationsecurityforcedisable msotrue msofalse
""".split())
KNOWN_NAMES = BUILTINS | HOST_OBJECTS | HOST_CONSTANTS | KEYWORDS | STATEMENT_KEYWORDS


def references(tokens: list[Tok], *, include_host_statements: bool = False) -> list[Tok]:
    """Root names only; member names, named-argument keys and labels are not variables."""
    result = []
    brackets = 0
    for index, token in enumerate(tokens):
        if token.text == "[":
            if not brackets:
                result.append(Tok("host", "Evaluate"))
            brackets += 1
        elif token.text == "]" and brackets:
            brackets -= 1
            continue
        if brackets:
            continue
        previous = tokens[index - 1] if index else None
        following = tokens[index + 1] if index + 1 < len(tokens) else None
        if not token.is_id() or (token.low in KEYWORDS | STATEMENT_KEYWORDS and
                                 not (include_host_statements and token.low in HOST_OBJECTS)):
            continue
        if previous and previous.is_op(".", "!"):
            continue
        if following and following.is_op(":=", ":"):
            continue
        if previous and previous.is_kw("goto", "gosub", "resume"):
            continue
        result.append(token)
    return result


def visible_names(modules: list[Module], current: Module) -> set[str]:
    names = set(current.members)
    for module in modules:
        names.add(module.name.casefold())
        names.update(name for name, member in module.members.items()
                     if member.visibility != "private" and
                     (module.kind == "standard" or member.kind in {"enum", "enum_member"}))
    return names


def procedure_findings(module: Module, proc: Proc, visible: set[str],
                       header: list[Tok]) -> list[dict[str, Any]]:
    known = visible | set(proc.locals) | KNOWN_NAMES | {proc.name.casefold()}
    findings = []
    for line, tokens in [(proc.line, header), *proc.statements]:
        # Rem after a colon is a comment, not an identifier use.
        if tokens and tokens[0].is_kw("rem"):
            continue
        for token in references(tokens):
            if token.low not in known:
                findings.append({"path": module.path, "procedure": proc.name, "line": line,
                                 "identifier": token.text, "code": "VBA-ID-001",
                                 "message": "Name is not declared or in the known VBA/Excel inventory."})
    return findings


def analyze(sources: dict[str, str]) -> dict[str, Any]:
    variants: dict[str, dict[str, str]] = {}
    failures: list[dict[str, Any]] = []
    for path, source in sources.items():
        reachable, errors = reachable_sources(path, source, comment_aware=True)
        failures.extend(errors)
        for environment, text in reachable.items():
            variants.setdefault(environment, {})[path] = text
    findings: list[dict[str, Any]] = []
    for environment, files in variants.items():
        modules = [parse_module(path, source, comment_aware=True) for path, source in files.items()]
        for module in modules:
            failures.extend(module.failures)
            visible = visible_names(modules, module)
            headers = {line: tokenize(code) for line, _, _, code
                       in logical_units(files[module.path].splitlines(), comment_aware=True)}
            for proc in module.procs:
                for item in procedure_findings(module, proc, visible, headers[proc.line]):
                    findings.append({**item, "environment": environment})
    return {"status": "fail" if failures else "pass", "advisory": True,
            "findings": findings, "analysis_failures": failures,
            "scope": "Root identifiers only; external/late-bound members and dynamic expressions require Excel compilation."}


def project_sources(root: Path) -> dict[str, dict[str, str]]:
    paths = tracked_files(root)
    files = {p: (root / p).read_bytes().decode("cp1252") for p in sorted(paths)
             if Path(p).suffix.lower() in {".bas", ".cls", ".frm"}}
    # Production must never resolve an undeclared name through a test module.
    return {"production": {p: s for p, s in files.items() if p.startswith("src/")},
            "development": {p: s for p, s in files.items() if p.startswith(("src/", "tests/"))},
            "examples": {p: s for p, s in files.items() if p.startswith(("src/", "examples/"))}}


def run_check(root: Path) -> dict[str, Any]:
    projects = {name: analyze(files) for name, files in project_sources(root).items()}
    return {"schema_version": 1, "status": "fail" if any(p["status"] == "fail" for p in projects.values()) else "pass",
            "advisory": True, "projects": projects}


def markdown_report(report: dict[str, Any]) -> str:
    return "## VBA identifiers (advisory)\n\n```json\n" + json.dumps(report, indent=2) + "\n```\n"


def main() -> int:
    options = parse_report_args(sys.argv[1:], description=__doc__)
    return run_gate(options,
                    build=lambda: run_check(options.root), markdown=markdown_report,
                    errors=(OSError, UnicodeError, ValueError, RuntimeError))


if __name__ == "__main__":
    raise SystemExit(main())
