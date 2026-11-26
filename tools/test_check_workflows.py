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
import subprocess
import tempfile
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
            (lambda w: cw.events(w["static-checks.yml"]).__setitem__(
                "pull_request", {"branches": ["main"]}), "pull_request has a filter"),
            (lambda w: cw.events(w["static-checks.yml"]).pop("pull_request"), "pull_request is missing"),
            (lambda w: cw.events(w["static-checks.yml"])["push"].__setitem__(
                "branches", ["main"]), "push branches"),
            (lambda w: cw.events(w["static-checks.yml"])["push"].__setitem__(
                "paths", ["src/**"]), "push has a filter"),
        ]
        for change, fragment in cases:
            with self.subTest(fragment=fragment):
                self.assertFinds(cw.check_integrity_job(mutated(change), GOVERNANCE), fragment)
        self.assertFinds(cw.check_integrity_job(WORKFLOWS, GOVERNANCE.replace(
            "**Repository integrity** required", "**Integrity** required")), "GOVERNANCE")

    def test_label_sync_outside_the_default_branch(self) -> None:
        guard = job(WORKFLOWS, "labels-sync.yml", "reconcile")["if"]
        for condition in ("github.event_name != 'pull_request'", guard.replace("&&", "||")):
            with self.subTest(condition=condition):
                def change(w: dict[str, Any]) -> None:
                    job(w, "labels-sync.yml", "reconcile")["if"] = condition
                self.assertFinds(cw.check_label_sync(mutated(change)), "default branch")

    def test_yaml_extension_is_checked(self) -> None:
        with tempfile.TemporaryDirectory() as folder:
            (Path(folder) / "extra.yaml").write_text(
                "on: push\njobs:\n  build:\n    steps:\n      - uses: actions/checkout@v7\n",
                encoding="utf-8")
            workflows = cw.load_workflows(Path(folder))
        self.assertEqual(list(workflows), ["extra.yaml"])
        self.assertFinds(cw.check_pins(workflows), "not pinned")
        self.assertFinds(cw.check_limits(workflows), "extra.yaml: no top-level permissions")

    def test_milestone_workflow(self) -> None:
        name = "pr-milestones.yml"
        cases: list[tuple[Callable[[dict[str, Any]], None], str]] = [
            (lambda w: cw.events(w[name]).__setitem__("pull_request", None), "events changed"),
            (lambda w: w[name].__setitem__("permissions", {"contents": "read"}), "top-level permissions"),
            (lambda w: job(w, name, "reconcile")["permissions"].__setitem__("issues", "write"),
             "permissions must be"),
            (lambda w: job(w, name, "reconcile")["permissions"].pop("issues"), "permissions must be"),
            (lambda w: job(w, name, "reconcile")["permissions"].__setitem__("contents", "read"),
             "permissions must be"),
            (lambda w: w[name]["concurrency"].__setitem__("group", "pr-milestones"), "concurrency"),
            (lambda w: w[name]["concurrency"].__setitem__(
                "group", "pr-milestones-${{ github.event.pull_request.number }}"), "concurrency"),
            (lambda w: w[name]["concurrency"].__setitem__("cancel-in-progress", True), "concurrency"),
            (lambda w: w[name].pop("concurrency"), "concurrency"),
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

    def test_self_hosted_job(self) -> None:
        name, quick = "excel-quick.yml", "quick"

        def events_of(w: dict[str, Any]) -> dict[str, Any]:
            found: dict[str, Any] = cw.events(w[name])
            return found

        cases: list[tuple[Callable[[dict[str, Any]], None], str]] = [
            (lambda w: job(w, "labels-sync.yml", "validate").__setitem__("runs-on", ["self-hosted", "excel"]),
             "labels-sync.yml/validate: runs on a self-hosted runner"),
            (lambda w: job(w, "labels-sync.yml", "validate").__setitem__("runs-on", "${{ inputs.runner }}"),
             "labels-sync.yml/validate: runs on a self-hosted runner"),
            (lambda w: events_of(w).__setitem__("pull_request_target", {"branches": ["main"]}),
             "event 'pull_request_target'"),
            (lambda w: events_of(w).__setitem__("push", {"branches": ["main"]}), "event 'push'"),
            (lambda w: events_of(w).__setitem__("pull_request", {}), "no branches filter"),
            (lambda w: w[name].__setitem__("permissions", {"contents": "read"}), "top-level permissions"),
            (lambda w: w[name]["concurrency"].__setitem__("cancel-in-progress", True), "cancel-in-progress: false"),
            (lambda w: w[name].pop("concurrency"), "cancel-in-progress: false"),
            (lambda w: job(w, name, quick).__setitem__("if", "github.event_name == 'pull_request'"),
             "condition does not limit"),
            (lambda w: job(w, name, quick).pop("if"), "condition does not limit"),
            (lambda w: job(w, name, quick)["permissions"].__setitem__("contents", "write"),
             "permissions must be contents: read"),
            (lambda w: job(w, name, quick)["permissions"].__setitem__("issues", "read"),
             "permissions must be contents: read"),
            (lambda w: job(w, name, quick).__setitem__("timeout-minutes", 600), "timeout-minutes must be at most"),
            (lambda w: job(w, name, quick).pop("timeout-minutes"), "timeout-minutes must be at most"),
            (lambda w: job(w, name, quick)["steps"][1].__setitem__(
                "run", "echo ${{ github.event.pull_request.head.ref }}"), "a run step contains"),
            (lambda w: job(w, name, quick)["steps"][1]["env"].__setitem__("TOKEN", "${{ secrets.PAT }}"),
             "uses a secret"),
            (lambda w: job(w, name, quick)["steps"][1].__setitem__("shell", "powershell"), "-ExecutionPolicy Bypass"),
            (lambda w: job(w, name, quick)["steps"][1].pop("shell"), "-ExecutionPolicy Bypass"),
        ]
        for change, fragment in cases:
            with self.subTest(fragment=fragment):
                self.assertFinds(cw.check_self_hosted(mutated(change)), fragment)

        def remove(w: dict[str, Any]) -> None:
            del w[name]
        self.assertFinds(cw.check_self_hosted(mutated(remove)), "excel-quick.yml is missing")

    def test_self_hosted_write_permission_is_listed(self) -> None:
        def change(w: dict[str, Any]) -> None:
            job(w, "excel-quick.yml", "quick")["permissions"]["issues"] = "write"
        self.assertFinds(cw.check_write_permissions(mutated(change)), "excel-quick.yml/quick: write permission ['issues']")

    def test_triage_runtimes_are_installed_and_exact(self) -> None:
        for name, job_id in (("issue-triage.yml", "audit"), (cw.INTEGRITY_FILE, cw.INTEGRITY_JOB)):
            for action, key in (("setup-node", "node-version"), ("setup-python", "python-version")):
                with self.subTest(name=name, action=action):
                    def change(w: dict[str, Any]) -> None:
                        step = next(s for s in job(w, name, job_id)["steps"] if action in s.get("uses", ""))
                        step["with"][key] = "latest"
                    self.assertFinds(cw.check_triage_runtimes(mutated(change)), "requires one")
                    def remove(w: dict[str, Any]) -> None:
                        configured = job(w, name, job_id)["steps"]
                        configured[:] = [s for s in configured if action not in s.get("uses", "")]
                    self.assertFinds(cw.check_triage_runtimes(mutated(remove)), "requires one")

    def test_supply_chain_violations(self) -> None:
        sources = {name: (ROOT / ".github/workflows" / name).read_text(encoding="utf-8") for name in WORKFLOWS}
        self.assertEqual(cw.check_supply_chain(WORKFLOWS, sources, DEPENDABOT, GOVERNANCE), [])
        self.assertFinds(cw.check_supply_chain({}, {}, DEPENDABOT, GOVERNANCE), "compatibility tests")
        bad_sources = {"extra.yml": "uses: owner/action@" + "0" * 40}
        self.assertFinds(cw.check_supply_chain({}, bad_sources, DEPENDABOT, GOVERNANCE), "version comment")
        config = copy.deepcopy(DEPENDABOT)
        config["updates"][0]["target-branch"] = "release/obsolete"
        self.assertFinds(cw.check_supply_chain({}, {}, config, GOVERNANCE), "target-branch")
        workflows: dict[str, Any]
        for command in ("pip install package", "python -m pip install package", "pip3 install package"):
            workflows = {"extra.yml": {"jobs": {"test": {"steps": [{"run": command}]}}}}
            self.assertFinds(cw.check_supply_chain(workflows, {}, DEPENDABOT, GOVERNANCE), "require-hashes")
        for action, key in (("setup-python", "python-version"), ("setup-node", "node-version")):
            workflows = {"extra.yml": {"jobs": {"test": {"steps": [
                {"uses": f"actions/{action}@" + "0" * 40, "with": {key: "latest"}}]}}}}
            self.assertFinds(cw.check_supply_chain(workflows, {}, DEPENDABOT, GOVERNANCE), "inconsistent")

    def test_dependabot(self) -> None:
        cases: list[tuple[Callable[[dict[str, Any]], None], str]] = [
            (lambda d: d["updates"].append(copy.deepcopy(d["updates"][0])), "2 blocks for 'github-actions'"),
            (lambda d: d["updates"].pop(), "0 blocks for 'pip'"),
            (lambda d: d["updates"][0].__setitem__("milestone", 3), "sets a milestone"),
            (lambda d: d["updates"][1].__setitem__("assignees", ["someone"]), "assignees"),
            (lambda d: d["updates"][0].__setitem__("labels", ["ci", "P2"]), "not ['P3']"),
            (lambda d: d["updates"][1].__setitem__("labels", ["ci", "P3", "P2"]), "not ['P3']"),
            (lambda d: d["updates"][0].pop("cooldown"), "github-actions: cooldown default-days is not 14"),
            (lambda d: d["updates"][1]["cooldown"].__setitem__("default-days", 7), "pip: cooldown default-days is not 14"),
        ]
        for change, fragment in cases:
            with self.subTest(fragment=fragment):
                config = copy.deepcopy(DEPENDABOT)
                change(config)
                self.assertTrue(any(fragment in f for f in cw.check_dependabot(config, BUG_FORM)),
                                cw.check_dependabot(config, BUG_FORM))


class TrafficAppTests(unittest.TestCase):
    def assertCredentialScope(self, workflow: dict[str, Any]) -> None:
        candidate = copy.deepcopy(workflow)
        traffic = candidate["jobs"]["traffic"]
        self.assertEqual(traffic["environment"], "analytics")
        all_steps = traffic["steps"]
        preflight = next(s for s in all_steps if s.get("name") == "Check the traffic App configuration")
        token = next(s for s in all_steps if s.get("id") == "app-token")
        fetch = next(s for s in all_steps if s.get("name") == "Fetch GitHub traffic API data")
        self.assertLess(all_steps.index(preflight), all_steps.index(token))
        self.assertLess(all_steps.index(token), all_steps.index(fetch))
        self.assertEqual(preflight.pop("env"), {
            "CLIENT_ID": "${{ vars.TRAFFIC_APP_CLIENT_ID }}",
            "HAS_KEY": "${{ secrets.TRAFFIC_APP_PRIVATE_KEY != '' }}",
        })
        self.assertEqual(token["uses"],
                         "actions/create-github-app-token@f418cff399f7aef4c2f0ad47e25ca7a6abbdd6fc")
        self.assertEqual(token.pop("with"), {
            "client-id": "${{ vars.TRAFFIC_APP_CLIENT_ID }}",
            "private-key": "${{ secrets.TRAFFIC_APP_PRIVATE_KEY }}",
            "permission-administration": "read",
        })
        self.assertEqual(fetch["env"].pop("TRAFFIC_TOKEN"), "${{ steps.app-token.outputs.token }}")
        # Inspect every remaining field, including workflow/job env and outputs.
        # Only the three exact bindings removed above may reference credentials.
        remaining = repr(candidate)
        for reference in ("secrets.", "secrets[", "steps.app-token.outputs", "steps['app-token']",
                          'steps["app-token"]', "toJSON(secrets)", "toJSON(steps)"):
            self.assertNotIn(reference, remaining)
        for step in all_steps:
            if step is not fetch:
                self.assertNotIn("TRAFFIC_TOKEN", repr(step))
        self.assertNotIn("GITHUB_ENV", fetch["run"])
        self.assertNotIn("GITHUB_OUTPUT", fetch["run"])

    def test_traffic_app_credentials_stay_in_their_steps(self) -> None:
        self.assertCredentialScope(WORKFLOWS["daily-traffic.yml"])

    def test_traffic_credential_leaks_are_rejected(self) -> None:
        mutations: list[Callable[[dict[str, Any]], None]] = [
            lambda w: w.__setitem__("env", {"KEY": "${{ secrets.TRAFFIC_APP_PRIVATE_KEY }}"}),
            lambda w: w["jobs"]["traffic"].__setitem__(
                "env", {"TOKEN": "${{ steps.app-token.outputs.token }}"}),
            lambda w: w["jobs"]["traffic"]["steps"].append(
                {"run": "echo unsafe", "env": {"KEY": "${{ secrets.TRAFFIC_APP_PRIVATE_KEY }}"}}),
            lambda w: w["jobs"]["traffic"]["steps"].append(
                {"run": "echo unsafe", "env": {"TOKEN": "${{ steps.app-token.outputs.token }}"}}),
        ]
        for index, change in enumerate(mutations):
            with self.subTest(index=index):
                workflow = copy.deepcopy(WORKFLOWS["daily-traffic.yml"])
                change(workflow)
                with self.assertRaises(AssertionError):
                    self.assertCredentialScope(workflow)

    def test_traffic_configuration_preflight(self) -> None:
        traffic_steps = WORKFLOWS["daily-traffic.yml"]["jobs"]["traffic"]["steps"]
        preflight = next(s for s in traffic_steps if s.get("name") == "Check the traffic App configuration")
        for client, has_key, expected in (("", "false", 1), ("", "true", 1),
                                          ("test-client", "false", 1), ("test-client", "true", 0)):
            with self.subTest(client=client, has_key=has_key):
                result = subprocess.run(["bash", "-c", preflight["run"]],
                                        env={**os.environ, "CLIENT_ID": client, "HAS_KEY": has_key},
                                        capture_output=True, text=True, check=False)
                self.assertEqual(result.returncode, expected)
                if expected:
                    self.assertIn("analytics environment", result.stdout)
                    self.assertIn("::error::", result.stdout)


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
        script: Callable[[Any, str, dict[str, Any]], list[tuple[int, str]]] = namespace["reconcile"]
        self.run_all: Callable[[Any, str, dict[str, Any], Callable[[str], None]], None] = namespace["run"]

        def quiet(request: Any, repository: str, event: dict[str, Any]) -> list[tuple[int, str]]:
            with contextlib.redirect_stdout(io.StringIO()):
                return script(request, repository, event)
        self.reconcile = quiet

    def test_pull_request_event_clears_its_milestone(self) -> None:
        api = FakeApi({7: {"pull_request": {}, "milestone": {"title": "v1.1.0"}}})
        self.reconcile(api, "owner/repo", {"pull_request": {}, "number": 7})
        self.assertIsNone(api.items[7]["milestone"])
        # The pull request is read again after the write, for what GitHub keeps.
        self.assertEqual(api.calls, [("GET", "/repos/owner/repo/issues/7"),
                                     ("PATCH", "/repos/owner/repo/issues/7"),
                                     ("GET", "/repos/owner/repo/issues/7")])

    def test_a_milestone_that_stays_after_the_write_is_an_error(self) -> None:
        class Sticky(FakeApi):
            def __call__(self, method: str, path: str, data: Any = None) -> Any:
                if method == "PATCH":  # The response says it is gone; the object keeps it.
                    self.calls.append((method, path))
                    return {"milestone": None}
                return super().__call__(method, path, data)

        api = Sticky({4: {"pull_request": {}, "milestone": {"title": "v1.3.0"}}})
        with self.assertRaisesRegex(ValueError, "PR #4 still has a milestone after the change"):
            self.reconcile(api, "owner/repo", {"pull_request": {}, "number": 4})

    def test_the_summary_lists_each_removal_and_treats_the_title_as_data(self) -> None:
        api = FakeApi({3: {"pull_request": {}, "milestone": {"title": "v1|2\n# x\\"}},
                       5: {"pull_request": {}, "milestone": {"title": "v1.3.0"}},
                       6: {"pull_request": {}, "milestone": None}})
        written: list[str] = []
        with contextlib.redirect_stdout(io.StringIO()):
            self.run_all(api, "owner/repo", {}, written.append)
        text = "".join(written)
        self.assertIn("## Keep milestones issue-only", text)
        self.assertIn("Scope: every pull request.", text)
        self.assertIn("| #3 | v1\\|2 # x\\\\ |", text)  # pipe escaped, line break and backslash neutralised
        self.assertIn("| #5 | v1.3.0 |", text)
        self.assertNotIn("#6", text)
        quiet_api = FakeApi({8: {"pull_request": {}, "milestone": None}})
        written.clear()
        with contextlib.redirect_stdout(io.StringIO()):
            self.run_all(quiet_api, "owner/repo", {"pull_request": {}, "number": 8}, written.append)
        self.assertIn("Scope: pull request #8.", "".join(written))
        self.assertIn("nothing was removed", "".join(written))

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
