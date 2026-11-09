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
from urllib.parse import quote, unquote, urldefrag

from check_links import links, tracked_markdown

MARKER = "<!-- saccr-external-link-audit -->"
USER_AGENT = "K-SACCR-link-audit/1.0"
API = "https://api.github.com"
# A github.com repository page: owner, repository (an optional .git dropped), then
# the path inside the repository; the query string, as in an issues search, is
# dropped. The site's own pages do not match.
REPOSITORY_PAGE = re.compile(r"^https://github\.com/([A-Za-z0-9][A-Za-z0-9-]*)/([A-Za-z0-9_.-]+?)"
                             r"(?:\.git)?(?:/([^?]*))?(?:\?.*)?$")
# A github.com user or organization page.
ACCOUNT_PAGE = re.compile(r"^https://github\.com/([A-Za-z0-9][A-Za-z0-9-]*)/?$")
# Listing pages verified as the repository itself when nothing follows them in
# the path: those every repository has, and those of an optional feature, which
# count only when the repository's API record has the feature on. A path below
# one, such as a workflow run, a discussion or a wiki page, names a resource of
# its own and stays on the web check, and so do the pages not listed here.
REPOSITORY_PAGES = {"stargazers", "watchers", "forks", "pulls", "pulse", "graphs", "network",
                    "tags", "branches", "releases", "commits", "activity"}
FEATURE_PAGES = {"issues": "has_issues", "labels": "has_issues", "milestones": "has_issues",
                 "wiki": "has_wiki", "discussions": "has_discussions", "projects": "has_projects"}


def api_targets(url: str) -> list[str]:
    """The API requests that verify a github.com repository page, in the order to try.

    GitHub answers unauthenticated scripted requests to its web pages with 429
    (rate limit) and, for some pages, 403 or 404, although a browser opens them
    (#649), so a repository page is verified through the API instead: the
    repository for a listing page (with its feature on, for an optional
    feature's page), a commit listing, a comparison, an issue
    or pull request, a milestone, a release, or a file or folder at a ref; a
    user or organization page is
    verified as the account. A ref may hold slashes, as
    release/1.3.0, so a blob or tree page gives one candidate per split of the
    path into ref and file, longest ref last; the first that answers wins. An
    empty list means the page is left to the web check.
    """
    account = ACCOUNT_PAGE.match(url)
    if account:
        return [f"/users/{account[1]}"]
    match = REPOSITORY_PAGE.match(url)
    if not match:
        return []
    owner, name, rest = match.groups()
    base = f"/repos/{owner}/{name}"
    # Split on the URL's own slashes, then decode each part, so that an escape in
    # the link (My%20File.md, a ref with %2F) is encoded once, not twice, below.
    parts = [unquote(part) for part in (rest or "").split("/") if part]
    if not parts:
        return [base]
    head, tail = parts[0], parts[1:]
    # A resource named by a number, or the latest release, is checked only when
    # nothing follows it: a page below it, such as pull/7/commits/<sha>, names a
    # resource of its own and stays on the web check.
    if head in ("issues", "pull") and len(tail) == 1 and tail[0].isdigit():
        return [f"{base}/issues/{tail[0]}"]
    if head == "milestone" and len(tail) == 1 and tail[0].isdigit():
        return [f"{base}/milestones/{tail[0]}"]
    if head == "releases" and tail == ["latest"]:
        return [f"{base}/releases/latest"]
    # A ref, a tag or a range may hold slashes, as release/1.3.0: the whole
    # rest of the path is the name. A download names an asset, which only the
    # web request checks.
    if head == "commits" and tail:
        return [f"{base}/commits?sha={quote('/'.join(tail), safe='')}&per_page=1"]
    if head == "releases" and len(tail) >= 2 and tail[0] == "tag":
        return [f"{base}/releases/tags/{quote('/'.join(tail[1:]), safe='')}"]
    if head in ("blob", "tree", "raw") and tail:
        # The root of a ref is /contents, without a trailing slash, which the API rejects.
        return [f"{base}/contents{'/' + quote('/'.join(tail[i:])) if tail[i:] else ''}"
                f"?ref={quote('/'.join(tail[:i]), safe='')}"
                for i in range(1, len(tail) + 1)]
    if head == "compare" and tail:
        return [f"{base}/compare/{quote('/'.join(tail), safe='.')}"]
    if (head in REPOSITORY_PAGES or head in FEATURE_PAGES) and not tail:
        return [base]
    return []


def required_feature(url: str) -> str | None:
    """The repository field an optional feature's listing page needs on, if it is one."""
    match = REPOSITORY_PAGE.match(url)
    parts = [part for part in (match[3] or "").split("/") if part] if match else []
    return FEATURE_PAGES.get(parts[0]) if len(parts) == 1 else None


def api_headers() -> dict[str, str]:
    headers = {"User-Agent": USER_AGENT, "Accept": "application/vnd.github+json",
               "X-GitHub-Api-Version": "2022-11-28"}
    if os.environ.get("GITHUB_TOKEN"):
        headers["Authorization"] = "Bearer " + os.environ["GITHUB_TOKEN"]
    return headers


def fetch(url: str, headers: dict[str, str], feature: str | None = None) -> tuple[str, str]:
    """Three attempts at one URL: ok, broken (404 or 410 every time) or inconclusive.

    With a feature, a JSON answer counts as ok only when that field is true; a
    feature turned off is reported as such, not as ok.
    """
    outcomes = []
    for attempt in range(3):
        try:
            request = urllib.request.Request(url, headers=headers)
            with urllib.request.urlopen(request, timeout=10) as response:
                code = response.status
                record = json.load(response) if feature and 200 <= code < 300 else {}
            if feature and 200 <= code < 300 and record.get(feature) is not True:
                return "inconclusive", f"{code} with {feature} off"
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


def inspect(url: str) -> tuple[str, str]:
    """Verify one URL: a github.com repository page through the API first, then the web."""
    notes = []
    feature = required_feature(url)
    for target in api_targets(url):
        status, detail = fetch(API + target, api_headers(), feature)
        if status == "ok":
            return "ok", "api " + detail
        notes.append("api " + detail)
    status, detail = fetch(url, {"User-Agent": USER_AGENT})
    return status, "; ".join(notes + [detail])


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
