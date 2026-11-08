#!/usr/bin/env python3
"""Check tracked Markdown destinations and GitHub-style anchors without network access."""
from __future__ import annotations

import argparse
import html
import json
import re
import subprocess
import unicodedata
from pathlib import Path
from urllib.parse import unquote, urlsplit


def prose(text: str) -> str:
    """Blank code and comments while retaining line offsets."""
    text = re.sub(r"<!--.*?-->", lambda m: "\n" * m[0].count("\n"), text, flags=re.S)
    lines = text.splitlines(keepends=True)
    fence = ""
    for i, line in enumerate(lines):
        mark = re.match(r"^\s{0,3}(`{3,}|~{3,})", line)
        if mark:
            token = mark[1]
            if not fence:
                fence = token
            elif token[0] == fence[0] and len(token) >= len(fence):
                fence = ""
            lines[i] = "\n" if line.endswith("\n") else ""
        elif fence or line.startswith("    ") or line.startswith("\t"):
            lines[i] = "\n" if line.endswith("\n") else ""
    return "".join(lines)


def slug(value: str) -> str:
    value = html.unescape(re.sub(r"<[^>]*>", "", value)).lower()
    value = re.sub(r"\[([^]]+)\]\([^)]*\)", r"\1", value)
    return "".join(c for c in value if c in " -_" or unicodedata.category(c)[0] in "LNM").replace(" ", "-")


def anchors(text: str) -> set[str]:
    clean = prose(text)
    found = set(re.findall(r'<a\b[^>]*(?:id|name)=["\']([^"\']+)', clean, re.I))
    counts: dict[str, int] = {}
    lines = clean.splitlines()
    for i, line in enumerate(lines):
        match = re.match(r"^\s{0,3}#{1,6}\s+(.+?)(?:\s+#+\s*)?$", line)
        value = match[1] if match else ""
        if i + 1 < len(lines) and re.fullmatch(r"\s{0,3}(?:=+|-+)\s*", lines[i + 1]) and line.strip():
            value = line.strip()
        if value:
            base = slug(value)
            candidate = base
            while candidate in counts:
                counts[base] = counts.get(base, 0) + 1
                candidate = f"{base}-{counts[base]}"
            counts[candidate] = 0
            found.add(candidate)
    return found


def links(text: str) -> list[tuple[int, str]]:
    clean = re.sub(r"(`+)([^`]|(?!\1)`)*?\1", "", prose(text))
    definitions = {m[1].strip().casefold(): m[2].strip("<>") for m in re.finditer(
        r"^\s{0,3}\[([^]]+)\]:\s*(<[^>]+>|\S+)", clean, re.M)}
    result = []
    for number, line in enumerate(clean.splitlines(), 1):
        if re.match(r"^\s{0,3}\[[^]]+\]:", line):
            continue
        for match in re.finditer(r"!?\[([^]]+)\](?:\((<[^>]*>|(?:[^()\s]|\([^()]*\))+)(?:\s+[\"'][^\n]*?[\"'])?\)|\[([^]]*)\])?", line):
            label, inline, reference = match.groups()
            if inline is not None:
                result.append((number, html.unescape(inline.strip("<>"))))
            else:
                key = (reference or label).strip().casefold()
                if key in definitions:
                    result.append((number, definitions[key]))
                elif reference is not None:
                    result.append((number, "MISSING_REFERENCE:" + key))
        for match in re.finditer(r'<(?:a|img)\b[^>]*(?:href|src)=["\']([^"\']+)', line, re.I):
            result.append((number, html.unescape(match[1])))
    return result


def tracked_markdown(root: Path) -> list[Path]:
    output = subprocess.check_output(["git", "ls-files", "-z"], cwd=root).decode()
    return [root / name for name in output.split("\0") if name.lower().endswith(".md")]


def check(root: Path, files: list[Path]) -> list[dict[str, str | int]]:
    findings: list[dict[str, str | int]] = []
    cache: dict[Path, set[str]] = {}
    for source in files:
        for line, link in links(source.read_text(encoding="utf-8-sig")):
            parsed = urlsplit(link)
            if parsed.scheme and not link.startswith("MISSING_REFERENCE:") or parsed.netloc:
                continue
            target = (source.parent / unquote(parsed.path)).resolve() if parsed.path else source.resolve()
            reason = ""
            if link.startswith("MISSING_REFERENCE:"):
                reason = "undefined reference"
            elif not target.is_relative_to(root.resolve()):
                reason = "path leaves repository"
            elif not target.exists():
                reason = "missing target"
            elif target.is_file() and target.stat().st_size == 0:
                reason = "empty target"
            elif parsed.fragment and target.suffix.lower() == ".md":
                if target not in cache:
                    cache[target] = anchors(target.read_text(encoding="utf-8-sig"))
                if unquote(parsed.fragment) not in cache[target]:
                    reason = "missing fragment"
            if reason:
                findings.append({"file": source.relative_to(root).as_posix(), "line": line,
                                 "link": link, "reason": reason})
    return findings


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", type=Path, default=Path(__file__).resolve().parents[1])
    parser.add_argument("--output", type=Path)
    args = parser.parse_args()
    root = args.root.resolve()
    findings = check(root, tracked_markdown(root))
    for item in findings:
        print(f"{item['file']}:{item['line']}: {item['reason']}: {item['link']}")
    if args.output:
        args.output.parent.mkdir(parents=True, exist_ok=True)
        args.output.write_text(json.dumps(findings, indent=2) + "\n", encoding="utf-8")
    print(f"local-links: {'FAIL' if findings else 'PASS'} ({len(findings)} findings)")
    return int(bool(findings))


if __name__ == "__main__":
    raise SystemExit(main())
