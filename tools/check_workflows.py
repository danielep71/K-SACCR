#!/usr/bin/env python3
"""Check the invariants of the GitHub workflows and the Dependabot configuration.

The files are read with a real YAML parser (PyYAML, hash-locked in
tools/requirements-quality-ci.txt). Each check returns a list of findings;
an empty list means the repository meets it. test_check_workflows.py runs
every check against the repository and against deliberate violations.
"""

from __future__ import annotations

import re
import sys
from collections.abc import Callable, Mapping
from pathlib import Path
from typing import Any

import yaml  # type: ignore[import-untyped]

ROOT = Path(__file__).resolve().parents[1]
WORKFLOWS = ROOT / ".github" / "workflows"

PINNED = re.compile(r"^[A-Za-z0-9_.-]+/[A-Za-z0-9_./-]+@[0-9a-f]{40}$")
DEFAULT_BRANCH_GUARD = "github.ref == format('refs/heads/{0}', github.event.repository.default_branch)"
# The only write permissions a job may hold: (workflow file, job id) -> scopes.
WRITE_ALLOWED: dict[tuple[str, str], set[str]] = {
    ("labels-sync.yml", "reconcile"): {"issues"},
    ("daily-traffic.yml", "traffic"): {"contents", "issues"},
    ("pr-milestones.yml", "reconcile"): {"pull-requests"},
}
INTEGRITY_FILE = "static-checks.yml"
INTEGRITY_JOB = "repository-integrity"
INTEGRITY_NAME = "Repository integrity"
INTEGRITY_BRANCHES = ["main", "release/**"]
MILESTONE_FILE = "pr-milestones.yml"
MILESTONE_EVENTS: dict[str, Any] = {
    "push": {"branches": ["main"], "paths": [".github/workflows/pr-milestones.yml"]},
    "pull_request_target": {"types": ["opened", "reopened", "edited", "milestoned"]},
    "schedule": [{"cron": "17 3 * * *"}],
    "workflow_dispatch": None,
}
LABEL_SYNC_GUARD = "github.event_name != 'pull_request' && " + DEFAULT_BRANCH_GUARD
MILESTONE_GUARD = ("github.event_name == 'pull_request_target' || "
                   "github.event_name == 'schedule' || " + DEFAULT_BRANCH_GUARD)
# Every ecosystem whose versions the repository pins.
PINNED_ECOSYSTEMS = ("github-actions", "pip")
PRIORITIES = ("P1", "P2", "P3")

Docs = dict[str, Any]


def load_yaml(path: Path) -> Any:
    return yaml.safe_load(path.read_text(encoding="utf-8"))


def load_workflows(folder: Path = WORKFLOWS) -> Docs:
    # GitHub Actions runs both extensions.
    paths = sorted([*folder.glob("*.yml"), *folder.glob("*.yaml")])
    return {path.name: load_yaml(path) for path in paths}


def events(workflow: Mapping[Any, Any]) -> Any:
    # PyYAML reads the bare key `on` as the boolean True.
    return workflow.get("on", workflow.get(True))


def jobs(workflow: Mapping[str, Any]) -> dict[str, Any]:
    found = workflow.get("jobs")
    return found if isinstance(found, dict) else {}


def steps(job: Mapping[str, Any]) -> list[Any]:
    found = job.get("steps")
    return found if isinstance(found, list) else []


def check_pins(workflows: Docs) -> list[str]:
    """Every `uses:` is pinned to a 40-character SHA; every checkout persists no credentials."""
    findings = []
    for name, workflow in workflows.items():
        for job_id, job in jobs(workflow).items():
            uses = [job.get("uses")] + [step.get("uses") for step in steps(job)]
            for target in (u for u in uses if u is not None):
                if not PINNED.match(str(target)):
                    findings.append(f"{name}/{job_id}: '{target}' is not pinned to a 40-character SHA")
            for step in steps(job):
                if str(step.get("uses", "")).startswith("actions/checkout@"):
                    if (step.get("with") or {}).get("persist-credentials") is not False:
                        findings.append(f"{name}/{job_id}: actions/checkout without persist-credentials: false")
    return findings


def check_limits(workflows: Docs) -> list[str]:
    """Every workflow declares permissions and every job a timeout."""
    findings = []
    for name, workflow in workflows.items():
        if "permissions" not in workflow:
            findings.append(f"{name}: no top-level permissions")
        for job_id, job in jobs(workflow).items():
            if "timeout-minutes" not in job:
                findings.append(f"{name}/{job_id}: no timeout-minutes")
    return findings


def write_scopes(permissions: Any) -> set[str]:
    if permissions in ("write-all",):
        return {"*"}
    if isinstance(permissions, Mapping):
        return {str(scope) for scope, level in permissions.items() if level == "write"}
    return set()


def check_write_permissions(workflows: Docs) -> list[str]:
    """No write permission outside WRITE_ALLOWED; none at workflow level."""
    findings = []
    for name, workflow in workflows.items():
        top = write_scopes(workflow.get("permissions"))
        if top:
            findings.append(f"{name}: workflow-level write permission {sorted(top)}")
        for job_id, job in jobs(workflow).items():
            extra = write_scopes(job.get("permissions")) - WRITE_ALLOWED.get((name, job_id), set())
            if extra:
                findings.append(f"{name}/{job_id}: write permission {sorted(extra)} is not allowed")
    return findings


def check_integrity_job(workflows: Docs, governance: str) -> list[str]:
    """The required Repository integrity job always runs, on the documented branches."""
    workflow = workflows.get(INTEGRITY_FILE)
    if not isinstance(workflow, dict):
        return [f"{INTEGRITY_FILE} is missing"]
    findings = []
    job = jobs(workflow).get(INTEGRITY_JOB)
    if not isinstance(job, dict):
        return [f"{INTEGRITY_FILE}: job '{INTEGRITY_JOB}' is missing"]
    if job.get("name") != INTEGRITY_NAME:
        findings.append(f"{INTEGRITY_FILE}/{INTEGRITY_JOB}: name is not '{INTEGRITY_NAME}'")
    for key in ("if", "strategy"):
        if key in job:
            findings.append(f"{INTEGRITY_FILE}/{INTEGRITY_JOB}: has '{key}', so the required check can be skipped")
    on = events(workflow) or {}
    for event in ("pull_request", "push"):
        if ((on.get(event) or {}).get("branches")) != INTEGRITY_BRANCHES:
            findings.append(f"{INTEGRITY_FILE}: {event} branches are not {INTEGRITY_BRANCHES}")
        if (on.get(event) or {}).keys() - {"branches"}:
            findings.append(f"{INTEGRITY_FILE}: {event} has a filter besides branches")
    if f"**{INTEGRITY_NAME}** required" not in governance:
        findings.append(f"docs/GOVERNANCE.md does not name '{INTEGRITY_NAME}' as the required check")
    return findings


def check_label_sync(workflows: Docs) -> list[str]:
    """Live labels are reconciled from the default branch only."""
    job = jobs(workflows.get("labels-sync.yml") or {}).get("reconcile")
    if not isinstance(job, dict):
        return ["labels-sync.yml: job 'reconcile' is missing"]
    condition = " ".join(str(job.get("if", "")).split())
    if condition != LABEL_SYNC_GUARD:
        return ["labels-sync.yml/reconcile: condition does not limit it to the default branch"]
    return []


def milestone_script(workflow: Mapping[str, Any]) -> str:
    """The inline Python of the milestone workflow, between <<'PY' and PY."""
    run = str(steps(jobs(workflow)["reconcile"])[0].get("run", ""))
    match = re.search(r"<<'PY'\n(.*)\nPY\s*$", run, re.S)
    if not match:
        raise ValueError("pr-milestones.yml: no inline Python block")
    return match.group(1)


def check_milestones(workflows: Docs) -> list[str]:
    """The milestone workflow keeps its events, permissions and single inline step."""
    workflow = workflows.get(MILESTONE_FILE)
    if not isinstance(workflow, dict):
        return [f"{MILESTONE_FILE} is missing"]
    findings = []
    if events(workflow) != MILESTONE_EVENTS:
        findings.append(f"{MILESTONE_FILE}: events changed")
    if workflow.get("permissions") != {}:
        findings.append(f"{MILESTONE_FILE}: top-level permissions are not {{}}")
    if list(jobs(workflow)) != ["reconcile"]:
        return findings + [f"{MILESTONE_FILE}: jobs are not exactly ['reconcile']"]
    job = jobs(workflow)["reconcile"]
    if job.get("permissions") != {"pull-requests": "write"}:
        findings.append(f"{MILESTONE_FILE}/reconcile: permissions are not pull-requests: write only")
    if " ".join(str(job.get("if", "")).split()) != MILESTONE_GUARD:
        findings.append(f"{MILESTONE_FILE}/reconcile: condition changed")
    job_steps = steps(job)
    if len(job_steps) != 1 or "uses" in job_steps[0] or "run" not in job_steps[0]:
        findings.append(f"{MILESTONE_FILE}/reconcile: not exactly one inline run step")
    elif "${{" in str(job_steps[0]["run"]):
        findings.append(f"{MILESTONE_FILE}/reconcile: the inline script contains '${{{{'")
    return findings


def configured_assignee(bug_form: str) -> str:
    match = re.search(r"^assignees:\s*\[\s*\"?([^\"\],]+)\"?\s*\]", bug_form, re.M)
    return match.group(1).strip() if match else ""


def check_dependabot(config: Any, bug_form: str) -> list[str]:
    """One block per pinned ecosystem, assigned as bug.yml, with P3 and no milestone."""
    updates = (config or {}).get("updates") or []
    findings = []
    ecosystems = [str(block.get("package-ecosystem")) for block in updates]
    for ecosystem in PINNED_ECOSYSTEMS:
        if ecosystems.count(ecosystem) != 1:
            findings.append(f"dependabot.yml: {ecosystems.count(ecosystem)} blocks for '{ecosystem}', not 1")
    assignee = configured_assignee(bug_form)
    for block in updates:
        label = str(block.get("package-ecosystem"))
        if "milestone" in block:
            findings.append(f"dependabot.yml/{label}: sets a milestone")
        if not assignee or block.get("assignees") != [assignee]:
            findings.append(f"dependabot.yml/{label}: assignees are not ['{assignee}'] (bug.yml)")
        priorities = [p for p in block.get("labels") or [] if p in PRIORITIES]
        if priorities != ["P3"]:
            findings.append(f"dependabot.yml/{label}: priority labels are {priorities}, not ['P3']")
    return findings


def all_findings(root: Path = ROOT) -> list[str]:
    workflows = load_workflows(root / ".github" / "workflows")
    governance = (root / "docs" / "GOVERNANCE.md").read_text(encoding="utf-8")
    bug_form = (root / ".github" / "ISSUE_TEMPLATE" / "bug.yml").read_text(encoding="utf-8")
    checks: list[Callable[[], list[str]]] = [
        lambda: check_pins(workflows),
        lambda: check_limits(workflows),
        lambda: check_write_permissions(workflows),
        lambda: check_integrity_job(workflows, governance),
        lambda: check_label_sync(workflows),
        lambda: check_milestones(workflows),
        lambda: check_dependabot(load_yaml(root / ".github" / "dependabot.yml"), bug_form),
    ]
    return [finding for check in checks for finding in check()]


def main() -> int:
    findings = all_findings()
    for finding in findings:
        print(f"FAIL: {finding}")
    print(f"workflow-invariants: {'PASS' if not findings else 'FAIL'}")
    return 1 if findings else 0


if __name__ == "__main__":
    sys.exit(main())
