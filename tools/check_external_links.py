#!/usr/bin/env python3
"""Read-only external URL checks; optionally maintain one marked GitHub issue."""
from __future__ import annotations

import argparse
import concurrent.futures
import json
import os
import re
import time
import urllib.error
import urllib.request
from pathlib import Path
from typing import Any
from urllib.parse import urldefrag

from check_links import links, tracked_markdown

MARKER = "<!-- saccr-external-link-audit -->"


def inspect(url: str) -> tuple[str, str]:
    outcomes = []
    for attempt in range(3):
        try:
            request = urllib.request.Request(url, headers={"User-Agent": "K-SACCR-link-audit/1.0"})
            with urllib.request.urlopen(request, timeout=10) as response:
                code = response.status
            if 200 <= code < 400:
                return "ok", str(code)
            outcomes.append(str(code))
        except urllib.error.HTTPError as error:
            outcomes.append(str(error.code))
        except (OSError, ValueError) as error:
            outcomes.append(type(error).__name__)
        if attempt < 2:
            time.sleep(0.2 * (attempt + 1))
    status = "broken" if all(code in {"404", "410"} for code in outcomes) else "inconclusive"
    return status, ", ".join(outcomes)


def api(repository: str, method: str, suffix: str, data: Any = None) -> Any:
    request = urllib.request.Request(
        f"https://api.github.com/repos/{repository}{suffix}", method=method,
        data=None if data is None else json.dumps(data).encode(),
        headers={"Authorization": "Bearer " + os.environ["GITHUB_TOKEN"],
                 "Accept": "application/vnd.github+json", "Content-Type": "application/json",
                 "X-GitHub-Api-Version": "2022-11-28"})
    with urllib.request.urlopen(request, timeout=30) as response:
        return json.load(response)


def publish(repository: str, body: str, failed: bool) -> None:
    candidates = []
    for page in range(1, 101):
        items = api(repository, "GET", f"/issues?state=all&per_page=100&page={page}")
        candidates += [item for item in items if "pull_request" not in item
                       and MARKER in (item.get("body") or "")]
        if len(items) < 100:
            break
    else:
        raise RuntimeError("Refusing partial issue scan")
    if candidates:
        selected = min(candidates, key=lambda item: item["number"])
        api(repository, "PATCH", f"/issues/{selected['number']}",
            {"body": body, "state": "open" if failed else "closed"})
    elif failed:
        governance = (Path(__file__).resolve().parents[1] / "docs/GOVERNANCE.md").read_text(encoding="utf-8")
        active = re.search(r"The active release branch is `release/([^`]+)`", governance)
        if not active:
            raise RuntimeError("Cannot identify active milestone")
        milestones = api(repository, "GET", "/milestones?state=open&per_page=100")
        matching = [m for m in milestones if m["title"].split(" ")[0] == "v" + active[1]]
        if len(matching) != 1:
            raise RuntimeError("Cannot identify one active milestone")
        api(repository, "POST", "/issues", {"title": "External documentation links need review",
            "body": body, "milestone": matching[0]["number"], "labels": ["documentation", "P3"], "assignees": ["danielep71"]})


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", type=Path, default=Path(__file__).resolve().parents[1])
    parser.add_argument("--output", type=Path, default=Path("test-results/external-links.md"))
    parser.add_argument("--publish", action="store_true")
    args = parser.parse_args()
    sources: dict[str, list[str]] = {}
    for path in tracked_markdown(args.root):
        for line, url in links(path.read_text(encoding="utf-8-sig")):
            if url.startswith(("https://", "http://")):
                sources.setdefault(urldefrag(url)[0], []).append(f"{path.relative_to(args.root)}:{line}")
    with concurrent.futures.ThreadPoolExecutor(max_workers=12) as pool:
        results = dict(zip(sources, pool.map(inspect, sources)))
    failures = {url: result for url, result in results.items() if result[0] != "ok"}
    body = MARKER + "\n## External link audit\n\n"
    body += "Three attempts per failing URL. Repeated 404/410 means broken; other failures are inconclusive (rate limits, access restrictions or transient errors).\n\n"
    for url, (status, detail) in failures.items():
        body += f"- **{status}**: {url} — {detail}\n  Sources: {', '.join(sources[url])}\n"
    if not failures:
        body += "All checked URLs responded successfully.\n"
    body += "\n## Dependencies\n\nNone. Review inconclusive URLs manually before changing documentation.\n"
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(body, encoding="utf-8")
    print(f"external-links: {len(sources)} URLs, {len(failures)} require review")
    if args.publish:
        publish(os.environ["GITHUB_REPOSITORY"], body, bool(failures))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
