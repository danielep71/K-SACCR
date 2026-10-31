"""Tests for the workflow and Dependabot invariants (tools/check_workflows.py).

PyYAML is required. Without it these tests are skipped locally and fail in
GitHub Actions, where tools/requirements-quality-ci.txt installs it.
"""

from __future__ import annotations

import contextlib
import copy
import io
import os
import re
import unittest
from collections.abc import Callable
from pathlib import Path
from typing import Any

try:
    import check_workflows as cw
except ModuleNotFoundError as error:  # pragma: no cover - depends on the environment
    if error.name != "yaml":
        raise
    if os.environ.get("GITHUB_ACTIONS"):
        raise RuntimeError("PyYAML is required in GitHub Actions; install "
                           "tools/requirements-quality-ci.txt before check.py") from error
    raise unittest.SkipTest("PyYAML is not installed") from error

ROOT = Path(__file__).resolve().parents[1]
GOVERNANCE = (ROOT / "docs" / "GOVERNANCE.md").read_text(encoding="utf-8")
BUG_FORM = (ROOT / ".github" / "ISSUE_TEMPLATE" / "bug.yml").read_text(encoding="utf-8")
WORKFLOWS = cw.load_workflows()
DEPENDABOT = cw.load_yaml(ROOT / ".github" / "dependabot.yml")


def mutated(change: Callable[[dict[str, Any]], None]) -> dict[str, Any]:
    workflows = copy.deepcopy(WORKFLOWS)
    change(workflows)
    return workflows


def job(workflows: dict[str, Any], name: str, job_id: str) -> dict[str, Any]:
    found: dict[str, Any] = workflows[name]["jobs"][job_id]
    return found


class RepositoryTests(unittest.TestCase):
    def test_repository_meets_every_invariant(self) -> None:
        self.assertEqual(cw.all_findings(), [])


class ViolationTests(unittest.TestCase):
    """Every check reports a deliberately introduced violation."""

    def assertFinds(self, findings: list[str], fragment: str) -> None:
        self.assertTrue(any(fragment in finding for finding in findings), findings)

    def test_unpinned_action(self) -> None:
        def change(w: dict[str, Any]) -> None:
            job(w, "labels-drift.yml", "fixtures")["steps"][0]["uses"] = "actions/checkout@v7"
        self.assertFinds(cw.check_pins(mutated(change)), "not pinned")

    def test_checkout_persisting_credentials(self) -> None:
        def change(w: dict[str, Any]) -> None:
            job(w, "daily-traffic.yml", "traffic")["steps"][0]["with"]["persist-credentials"] = True
        self.assertFinds(cw.check_pins(mutated(change)), "persist-credentials")

        def remove(w: dict[str, Any]) -> None:
            del job(w, "static-checks.yml", "repository-integrity")["steps"][0]["with"]["persist-credentials"]
        self.assertFinds(cw.check_pins(mutated(remove)), "persist-credentials")

    def test_missing_permissions_and_timeout(self) -> None:
        def change(w: dict[str, Any]) -> None:
            del w["issue-triage.yml"]["permissions"]
            del job(w, "labels-sync.yml", "validate")["timeout-minutes"]
        findings = cw.check_limits(mutated(change))
        self.assertFinds(findings, "issue-triage.yml: no top-level permissions")
        self.assertFinds(findings, "labels-sync.yml/validate: no timeout-minutes")

    def test_write_permission_outside_the_list(self) -> None:
        def change(w: dict[str, Any]) -> None:
            job(w, "labels-sync.yml", "reconcile")["permissions"]["contents"] = "write"
            job(w, "issue-triage.yml", "audit")["permissions"]["issues"] = "write"
            w["labels-drift.yml"]["permissions"] = "write-all"
        findings = cw.check_write_permissions(mutated(change))
        self.assertFinds(findings, "labels-sync.yml/reconcile: write permission ['contents']")
        self.assertFinds(findings, "issue-triage.yml/audit: write permission ['issues']")
        self.assertFinds(findings, "labels-drift.yml: workflow-level write")

    def test_integrity_job(self) -> None:
        cases: list[tuple[Callable[[dict[str, Any]], None], str]] = [
            (lambda w: job(w, "static-checks.yml", "repository-integrity").__setitem__("if", "false"), "'if'"),
            (lambda w: job(w, "static-checks.yml", "repository-integrity").__setitem__(
                "strategy", {"matrix": {"python": ["3.10"]}}), "'strategy'"),
            (lambda w: job(w, "static-checks.yml", "repository-integrity").__setitem__(
                "name", "Integrity"), "name is not"),
            (lambda w: cw.events(w["static-checks.yml"])["pull_request"].__setitem__(
                "branches", ["main"]), "pull_request branches"),
            (lambda w: cw.events(w["static-checks.yml"])["push"].__setitem__(
                "paths", ["src/**"]), "push has a filter"),
        ]
        for change, fragment in cases:
            with self.subTest(fragment=fragment):
                self.assertFinds(cw.check_integrity_job(mutated(change), GOVERNANCE), fragment)
        self.assertFinds(cw.check_integrity_job(WORKFLOWS, GOVERNANCE.replace(
            "**Repository integrity** required", "**Integrity** required")), "GOVERNANCE")

    def test_label_sync_outside_the_default_branch(self) -> None:
        def change(w: dict[str, Any]) -> None:
            job(w, "labels-sync.yml", "reconcile")["if"] = "github.event_name != 'pull_request'"
        self.assertFinds(cw.check_label_sync(mutated(change)), "default branch")

    def test_milestone_workflow(self) -> None:
        name = "pr-milestones.yml"
        cases: list[tuple[Callable[[dict[str, Any]], None], str]] = [
            (lambda w: cw.events(w[name]).__setitem__("pull_request", None), "events changed"),
            (lambda w: w[name].__setitem__("permissions", {"contents": "read"}), "top-level permissions"),
            (lambda w: job(w, name, "reconcile")["permissions"].__setitem__("issues", "write"),
             "pull-requests: write only"),
            (lambda w: job(w, name, "reconcile").__setitem__("if", "true"), "condition changed"),
            (lambda w: job(w, name, "reconcile")["steps"].insert(
                0, {"uses": "actions/checkout@" + "0" * 40}), "one inline run step"),
            (lambda w: job(w, name, "reconcile")["steps"][0].__setitem__(
                "run", job(w, name, "reconcile")["steps"][0]["run"] + "\necho ${{ github.head_ref }}\n"),
             "contains '${{'"),
        ]
        for change, fragment in cases:
            with self.subTest(fragment=fragment):
                self.assertFinds(cw.check_milestones(mutated(change)), fragment)

    def test_dependabot(self) -> None:
        cases: list[tuple[Callable[[dict[str, Any]], None], str]] = [
            (lambda d: d["updates"].append(copy.deepcopy(d["updates"][0])), "2 blocks for 'github-actions'"),
            (lambda d: d["updates"].pop(), "0 blocks for 'pip'"),
            (lambda d: d["updates"][0].__setitem__("milestone", 3), "sets a milestone"),
            (lambda d: d["updates"][1].__setitem__("assignees", ["someone"]), "assignees"),
            (lambda d: d["updates"][0].__setitem__("labels", ["ci", "P2"]), "not ['P3']"),
            (lambda d: d["updates"][1].__setitem__("labels", ["ci", "P3", "P2"]), "not ['P3']"),
        ]
        for change, fragment in cases:
            with self.subTest(fragment=fragment):
                config = copy.deepcopy(DEPENDABOT)
                change(config)
                self.assertTrue(any(fragment in f for f in cw.check_dependabot(config, BUG_FORM)),
                                cw.check_dependabot(config, BUG_FORM))


class FakeApi:
    """A GitHub API stand-in for the milestone script: pull requests and issues by number."""

    def __init__(self, items: dict[int, dict[str, Any]]) -> None:
        self.items = items
        self.calls: list[tuple[str, str]] = []

    def __call__(self, method: str, path: str, data: Any = None) -> Any:
        self.calls.append((method, path))
        number = re.search(r"/issues/(\d+)$", path)
        if number:
            item = self.items[int(number.group(1))]
            if method == "PATCH":
                item["milestone"] = data["milestone"]
            return dict(item)
        # Read the page from page=, not from per_page=, which matches first.
        page = int(re.search(r"[?&]page=(\d+)", path).group(1))  # type: ignore[union-attr]
        pulls = [dict(item, number=n) for n, item in sorted(self.items.items()) if "pull_request" in item]
        return pulls[(page - 1) * 100:page * 100]


class MilestoneScriptTests(unittest.TestCase):
    def setUp(self) -> None:
        namespace: dict[str, Any] = {"__name__": "pr_milestones"}
        exec(compile(cw.milestone_script(WORKFLOWS["pr-milestones.yml"]), "pr-milestones", "exec"),
             namespace)
        script: Callable[[Any, str, dict[str, Any]], None] = namespace["reconcile"]

        def quiet(request: Any, repository: str, event: dict[str, Any]) -> None:
            with contextlib.redirect_stdout(io.StringIO()):
                script(request, repository, event)
        self.reconcile = quiet

    def test_pull_request_event_clears_its_milestone(self) -> None:
        api = FakeApi({7: {"pull_request": {}, "milestone": {"title": "v1.1.0"}}})
        self.reconcile(api, "owner/repo", {"pull_request": {}, "number": 7})
        self.assertIsNone(api.items[7]["milestone"])
        self.assertEqual(api.calls, [("GET", "/repos/owner/repo/issues/7"),
                                     ("PATCH", "/repos/owner/repo/issues/7")])

    def test_scan_clears_every_page(self) -> None:
        items: dict[int, dict[str, Any]] = {n: {"pull_request": {}, "milestone": None} for n in range(1, 151)}
        for n in (3, 120, 150):
            items[n]["milestone"] = {"title": "v1.1.0"}
        api = FakeApi(items)
        self.reconcile(api, "owner/repo", {})
        self.assertTrue(all(item["milestone"] is None for item in items.values()))
        pages = [path for method, path in api.calls if "/pulls?" in path]
        self.assertEqual(len(pages), 2)
        self.assertEqual(sum(1 for method, _ in api.calls if method == "PATCH"), 3)

    def test_an_issue_is_never_changed(self) -> None:
        api = FakeApi({9: {"milestone": {"title": "v1.1.0"}}})
        with self.assertRaisesRegex(ValueError, "Refusing to change an issue"):
            self.reconcile(api, "owner/repo", {"pull_request": {}, "number": 9})
        self.assertEqual(api.items[9]["milestone"], {"title": "v1.1.0"})
        self.assertNotIn("PATCH", [method for method, _ in api.calls])


if __name__ == "__main__":
    unittest.main()
