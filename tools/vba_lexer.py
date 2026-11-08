"""Shared lexical primitives for VBA gates (not a VBA compiler).

Legacy callers choose whether strings are preserved or masked; retaining these
profiles keeps their established findings and physical line numbers unchanged.
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field

TOKEN_RE = re.compile(
    r"""
    (?P<ws>\s+)
  | (?P<str>"(?:[^"]|"")*")
  | (?P<date>\#\s*\d[\d/\-:. ]*(?:[AaPp][Mm])?\s*\#)
  | (?P<num>&[HhOo][0-9A-Fa-f]+&?|\d+(?:\.\d+)?(?:[Ee][+-]?\d+)?[%&!#@^]?)
  | (?P<op>:=|<>|<=|>=|[-+*/\\^&=<>(),.:;!])
  | (?P<id>[A-Za-z_][A-Za-z0-9_]*[$%&@]?)
  | (?P<other>.)
    """,
    re.VERBOSE,
)


@dataclass(frozen=True)
class Tok:
    kind: str
    text: str
    # A "." after whitespace that follows the first identifier of a statement: "Check .Name = a, m"
    # passes a With member to Check, while "Check.Name = a" assigns a member of Check (#498).
    spaced: bool = field(default=False, compare=False)

    @property
    def low(self) -> str:
        return self.text.casefold().rstrip("$%&@")

    def is_id(self) -> bool:
        return self.kind == "id"

    def is_op(self, *values: str) -> bool:
        return self.kind == "op" and self.text in values

    def is_kw(self, *words: str) -> bool:
        return self.kind == "id" and self.low in words

    def continues_chain(self) -> bool:
        """True for the ``.`` or ``!`` of a member access, false for a spaced ``.`` that starts a With member."""
        return self.is_op("!") or (self.is_op(".") and not self.spaced)


def tokenize(code: str) -> list[Tok]:
    """Tokens without whitespace; a "." keeps whether whitespace came before it."""
    tokens = []
    spaced = False
    for match in TOKEN_RE.finditer(code):
        kind = match.lastgroup or "other"
        if kind == "ws":
            spaced = True
            continue
        tokens.append(Tok(kind, match.group(), spaced and match.group() == "."))
        spaced = False
    return tokens


def strip_vba(raw: str, *, mask_strings: bool = False) -> str:
    result: list[str] = []
    in_string = False
    index = 0
    while index < len(raw):
        char = raw[index]
        if char == '"':
            if in_string and index + 1 < len(raw) and raw[index + 1] == '"':
                result.extend('  ' if mask_strings else '""')
                index += 2
                continue
            in_string = not in_string
            result.append(' ' if mask_strings else char)
            index += 1
            continue
        if char == "'" and not in_string:
            break
        result.append(' ' if mask_strings and in_string else char)
        index += 1
    text = "".join(result)
    if re.match(r"^\s*Rem(?:\s|$)", text, re.IGNORECASE):
        return ""
    return text.rstrip()


def is_date_delimiter(code: str, index: int, in_date: bool) -> bool:
    if in_date:
        return True
    # An adjacent identifier/numeric token owns its # type suffix.
    previous = code[index - 1] if index else ""
    if previous and (previous.isalnum() or previous in "_.)]"):
        return False
    # A literal needs a closing delimiter; e.g. Print #1 is not a date.
    return "#" in code[index + 1:]


def split_statements(code: str) -> list[str]:
    """Split colons outside VBA strings, preserving named arguments and Rem comments."""
    statements: list[str] = []
    start = 0
    in_string = False
    in_date = False
    index = 0
    while index < len(code):
        char = code[index]
        if char == '"' and not in_date:
            if in_string and code[index:index + 2] == '""':
                index += 2
                continue
            in_string = not in_string
        elif char == "#" and not in_string and is_date_delimiter(code, index, in_date):
            in_date = not in_date
        elif char == ":" and not (in_string or in_date) and code[index:index + 2] != ":=":
            statement = code[start:index].strip()
            if re.match(r"^Rem(?:\s|$)", statement, re.I):
                return statements
            statements.append(statement)
            start = index + 1
        index += 1
    statement = code[start:].strip()
    if not re.match(r"^Rem(?:\s|$)", statement, re.I):
        statements.append(statement)
    return statements


def logical_lines(lines: list[str], *, mask_strings: bool = False) -> list[tuple[int, int, str]]:
    result: list[tuple[int, int, str]] = []
    buffer: list[str] = []
    start = 0
    for number, raw in enumerate(lines, 1):
        code = strip_vba(raw, mask_strings=mask_strings)
        if not buffer:
            start = number
        if re.search(r"\s_\s*$", code):
            buffer.append(re.sub(r"\s_\s*$", " ", code))
            continue
        buffer.append(code)
        result.append((start, number, " ".join(item.strip() for item in buffer)))
        buffer.clear()
    if buffer:
        result.append((start, len(lines), " ".join(item.strip() for item in buffer)))
    return result


def non_ascii_outside_strings(text: str) -> list[tuple[int, int]]:
    """Physical locations, including comments (quotes in a comment are not literals)."""
    findings: list[tuple[int, int]] = []
    continued_comment = False
    for line, raw in enumerate(text.splitlines(), 1):
        quoted = False
        comment = continued_comment
        statement_start = True
        for index, char in enumerate(raw):
            if not comment:
                if not quoted and (char == "'" or
                                   (statement_start and re.match(r"Rem(?:\s|$)", raw[index:], re.I))):
                    comment = True
                elif char == '"':
                    quoted = not quoted
            if ord(char) > 127 and (comment or not quoted):
                findings.append((line, index + 1))
            if char == ":" and not quoted:
                statement_start = True
            elif not char.isspace() and not (statement_start and char.isdigit()):
                statement_start = False
        continued_comment = comment and raw.rstrip().endswith(" _")
    return findings
