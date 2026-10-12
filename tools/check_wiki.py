#!/usr/bin/env python3
"""Validate Wiki sources, export an exact-source bundle, or compare a Wiki checkout."""
from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
from pathlib import Path, PurePosixPath
from typing import Any

from _gatelib import git_text, run_gate, tracked_files

WIKI = "docs/wiki"
CATALOGUE = f"{WIKI}/catalogue.json"
NOTICE = "> **Guide, not policy:**"


def require(condition: Any, message: str) -> None:
    if not condition:
        raise ValueError(message)


def read_json(path: Path) -> dict[str, Any]:
    value = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(value, dict):
        raise ValueError("JSON root must be an object")
    return value


def page_rows(data: dict[str, Any]) -> list[dict[str, Any]]:
    pages = data.get("pages")
    if not isinstance(pages, list) or not pages:
        raise ValueError("page order is required")
    result: list[dict[str, Any]] = []
    for page in pages:
        require(isinstance(page, dict), "page entries must be objects")
        result.append(page)
    return result


def repository_name(root: Path) -> str:
    candidate = os.environ.get("GITHUB_REPOSITORY", "").strip()
    if not candidate:
        remote = git_text(root, "remote", "get-url", "origin", check=True).stdout.strip()
        match = re.search(r"github\.com(?::|/)([^/]+/[^/]+?)(?:\.git)?$", remote)
        require(match is not None, "origin is not a supported GitHub URL")
        assert match is not None
        candidate = match[1]
    require(re.fullmatch(r"[A-Za-z0-9_.-]+/[A-Za-z0-9_.-]+", candidate) is not None,
            "invalid GitHub repository name")
    return candidate


def sidebar(data: dict[str, Any]) -> str:
    lines = ["# SA-CCR Benchmark", ""]
    lines.extend(f'- [{p["title"]}]({p["name"]}.md)' for p in page_rows(data))
    lines += ["", NOTICE + " versioned repository authorities linked from each page govern.", ""]
    return "\n".join(lines)


def validate_catalogue(root: Path, data: dict[str, Any], files: set[str]) -> None:
    require(data.get("schema_version") == 1, "unsupported Wiki catalogue")
    pages = page_rows(data)
    names = [str(p.get("name", "")) for p in pages]
    require(names[0] == "Home", "Home must be the first Wiki page")
    require(len(set(names)) == len(names), "duplicate Wiki page names")
    require(all(re.fullmatch(r"[A-Z][A-Za-z0-9-]*", n) for n in names), "unsafe Wiki page name")
    expected = {f"{WIKI}/{n}.md" for n in names} | {f"{WIKI}/_Sidebar.md"}
    if f"{WIKI}/_Footer.md" in files:
        expected.add(f"{WIKI}/_Footer.md")
    actual = {p for p in files if p.startswith(WIKI + "/") and p.endswith(".md")}
    require(expected == actual, "Wiki catalogue differs from source Markdown set: "
            + ", ".join(sorted(expected ^ actual)))
    for page in pages:
        name = str(page.get("name", ""))
        title = page.get("title")
        authority = str(page.get("authority", ""))
        require(isinstance(title, str) and bool(title.strip()), f"missing page title: {name}")
        require(authority in files, f"missing page authority: {name}")
        text = (root / WIKI / f"{name}.md").read_text(encoding="utf-8")
        notices = [line for line in text.splitlines() if NOTICE in line]
        require(len(notices) == 1, f"page must contain one authority notice: {name}")
        targets = re.findall(r"\]\(([^\s)]+)\)", notices[0])
        wanted = (root / authority).resolve()
        require(any((root / WIKI / t.partition("#")[0]).resolve() == wanted for t in targets),
                f"authority notice does not link registered authority: {name}")


def render_page(text: str, page: str, names: set[str], repository: str, sha: str,
                provenance: bool = True) -> str:
    def link(match: re.Match[str]) -> str:
        target = match[1]
        if target.startswith(("https://", "http://", "mailto:", "#")):
            return match[0]
        path, separator, fragment = target.partition("#")
        if path.endswith(".md") and path[:-3] in names:
            return "](" + path[:-3] + (separator + fragment if separator else "") + ")"
        parts: list[str] = []
        for part in (PurePosixPath(WIKI) / path).parts:
            if part == "..":
                require(bool(parts), "Wiki link escapes repository")
                parts.pop()
            elif part != ".":
                parts.append(part)
        suffix = separator + fragment if separator else ""
        return f'](https://github.com/{repository}/blob/{sha}/{"/".join(parts)}{suffix})'
    rendered = re.sub(r"\]\(([^\s)]+)\)", link, text)
    if not provenance:
        return rendered.rstrip() + "\n"
    url = f"https://github.com/{repository}/blob/{sha}/{WIKI}/{page}"
    return rendered.rstrip() + f"\n\n---\nPublished from [{sha}]({url}).\n"


def bundle(root: Path, data: dict[str, Any], sha: str, files: set[str]) -> dict[str, bytes]:
    require(re.fullmatch(r"[0-9a-f]{40}", sha) is not None, "source SHA must be a full commit")
    repository = repository_name(root)
    names = {str(p["name"]) for p in page_rows(data)}
    result: dict[str, bytes] = {}
    # The footer is published when tracked, the rule validate_catalogue applies.
    # It is shown under every page, which already carries its own provenance line.
    special = {"_Sidebar"}
    if f"{WIKI}/_Footer.md" in files:
        special.add("_Footer")
    for name in sorted(names | special):
        page = name + ".md"
        source = (root / WIKI / page).read_text(encoding="utf-8")
        rendered = render_page(source, page, names, repository, sha, provenance=name != "_Footer")
        result[page] = rendered.encode("utf-8")
    record = {"schema_version": 1, "repository": repository, "source_sha": sha,
              "pages": {p: hashlib.sha256(b).hexdigest() for p, b in sorted(result.items())}}
    result["Wiki-Source.json"] = (json.dumps(record, indent=2, sort_keys=True) + "\n").encode()
    return result


def compare_bundle(expected: dict[str, bytes], destination: Path) -> list[str]:
    require(destination.is_dir() and not destination.is_symlink(), "published Wiki unavailable")
    paths = [p for p in destination.rglob("*") if ".git" not in p.relative_to(destination).parts]
    require(not any(p.is_symlink() for p in paths), "published Wiki contains a symlink")
    actual = {p.relative_to(destination).as_posix() for p in paths if p.is_file()}
    findings = [f"published path set differs: {p}" for p in sorted(actual ^ set(expected))]
    findings += [f"published bytes differ: {p}" for p in sorted(actual & set(expected))
                 if (destination / p).read_bytes() != expected[p]]
    return findings


def write_bundle(expected: dict[str, bytes], destination: Path, root: Path) -> None:
    require(not destination.exists(), "export destination must not exist")
    require(not destination.resolve().is_relative_to(root.resolve()),
            "export must be outside the source repository")
    destination.mkdir(parents=True)
    for path, content in expected.items():
        (destination / path).write_bytes(content)


def build_report(args: argparse.Namespace) -> dict[str, Any]:
    root = args.root.resolve()
    findings: list[str] = []
    try:
        files = tracked_files(root)
        data = read_json(root / CATALOGUE)
        validate_catalogue(root, data, files)
        expected_sidebar = sidebar(data)
        side = root / WIKI / "_Sidebar.md"
        if args.write_sidebar:
            side.write_text(expected_sidebar, encoding="utf-8", newline="\n")
        elif side.read_text(encoding="utf-8") != expected_sidebar:
            findings.append("_Sidebar.md is stale; use --write-sidebar")
        if args.export_dir or args.published_dir:
            require(not findings, "fix Wiki source findings before publication")
            sha = git_text(root, "rev-parse", "HEAD", check=True).stdout.strip()
            require(not git_text(root, "status", "--porcelain", check=True).stdout,
                    "Wiki publication requires a clean committed source tree")
            expected = bundle(root, data, sha, files)
            if args.export_dir:
                write_bundle(expected, args.export_dir, root)
            if args.published_dir:
                findings.extend(compare_bundle(expected, args.published_dir))
    except (ValueError, OSError, KeyError, TypeError, RuntimeError) as error:
        findings.append(str(error))
    return {"status": "fail" if findings else "pass", "findings": sorted(findings),
            "scope_note": "Checks Wiki source/navigation and exact published bytes; "
                          "it does not validate SA-CCR methodology or Excel execution."}


def markdown(report: dict[str, Any]) -> str:
    findings = report["findings"]
    require(isinstance(findings, list), "invalid findings report")
    lines = ["# Wiki contract", "", f'Result: {str(report["status"]).upper()}', ""]
    lines += [f"- {item}" for item in findings]
    lines += ["", str(report["scope_note"]), ""]
    return "\n".join(lines)


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", type=Path, default=Path.cwd())
    parser.add_argument("--output", type=Path)
    parser.add_argument("--summary", type=Path)
    mode = parser.add_mutually_exclusive_group()
    mode.add_argument("--write-sidebar", action="store_true")
    mode.add_argument("--export-dir", type=Path)
    mode.add_argument("--published-dir", type=Path)
    args = parser.parse_args()
    return run_gate(args, build=lambda: build_report(args), markdown=markdown, errors=(OSError,))


if __name__ == "__main__":
    raise SystemExit(main())
