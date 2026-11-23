"""Tests for the issue triage rules (tools/check_issues.py)."""

from __future__ import annotations

import contextlib
import io
import json
import shutil
import tempfile
import unittest
import unittest.mock
from dataclasses import replace
from pathlib import Path

import check_issues
from check_issues import Issue

ROOT = Path(__file__).resolve().parents[1]
REPOSITORY = "example-org/sample-project"
BLOCKED_LABEL = check_issues.BLOCKED
START = "### 🗓️ Start date (Europe/Rome)\n\n2027-01-04\n\n"
READY = "## Dependencies and blocking status\n\n**Ready to start**: none.\n"


def issue(number: int, *labels: str, state: str = "open", body: str = READY,
          assignees: tuple[str, ...] = ("triage-person",), milestone: str | None = "v0.1.0",
          start: bool = True) -> Issue:
    """An issue that meets every rule; the start date every open issue needs is added unless the body has one."""
    if start and "start date" not in body.lower():
        body = START + body
    return Issue(number, state, labels or ("P3",), assignees, milestone, body)


def blocked_by(*numbers: int, heading: str = "## Dependencies and blocking status") -> str:
    lines = "\n".join(f"- #{n} — prerequisite" for n in numbers)
    return f"## Objective\n\nx\n\n{heading}\n\n**blocked**:\n\n{lines}\n\n## Tracking\n\n#99\n"


def rules(issues: list[Issue], states: dict[int, str] | None = None,
          pulls: frozenset[int] = frozenset()) -> list[tuple[int, str]]:
    known = states if states is not None else {i.number: i.state for i in issues}
    return [(f.number, f.rule)
            for f in check_issues.audit(issues, known, REPOSITORY, pulls=pulls, assignee="triage-person")]


class MetadataTests(unittest.TestCase):
    def test_compliant_issue_has_no_findings(self) -> None:
        self.assertEqual(rules([issue(1)]), [])

    def test_assignee_priority_and_milestone_are_required(self) -> None:
        self.assertEqual(rules([issue(1, assignees=())]), [(1, "assignee")])
        self.assertEqual(rules([issue(1, "documentation")]), [(1, "priority")])
        self.assertEqual(rules([issue(1, "P2", "P3")]), [(1, "priority")])
        self.assertEqual(rules([issue(1, milestone=None)]), [(1, "milestone")])

    def test_p1_is_reserved_for_bugs(self) -> None:
        self.assertEqual(rules([issue(1, "P1", "documentation")]), [(1, "p1-reserved")])
        self.assertEqual(rules([issue(1, "P1", "bug")]), [])

    def test_traffic_alerts_need_no_milestone_or_dependency_section(self) -> None:
        alert = Issue(5, "open", ("P3",), ("triage-person",), None, "Automated traffic alert.",
                      "Traffic alert: 2026-10-31", check_issues.TRAFFIC_ALERT_AUTHOR)
        self.assertEqual(rules([alert]), [])
        # The exemption needs both the workflow's author and its title.
        by_person = Issue(5, "open", ("P3",), ("triage-person",), None, "x",
                          "Traffic alert: 2026-10-31", "someone")
        self.assertEqual(rules([by_person]), [(5, "milestone"), (5, "start-date"), (5, "dependencies")])
        unassigned = Issue(5, "open", (), (), None, "x", "Traffic alert: 2026-10-31",
                           check_issues.TRAFFIC_ALERT_AUTHOR)
        self.assertEqual(rules([unassigned]), [(5, "assignee"), (5, "priority")])

    def test_closed_issues_are_not_audited(self) -> None:
        self.assertEqual(rules([issue(1, "documentation", state="closed", assignees=())]), [])

    def test_unknown_labels_are_reported_against_the_catalogue(self) -> None:
        findings = check_issues.audit([issue(1, "P3", "ready")], {1: "open"}, REPOSITORY,
                                      catalogue={"P3"}, assignee="triage-person")
        self.assertEqual([(f.number, f.rule) for f in findings], [(1, "catalogue")])


class PlanningTests(unittest.TestCase):
    def test_dates_and_milestone_state_from_api(self) -> None:
        parsed = check_issues.issue_from_api({"number": 1, "state": "open", "milestone": {
            "title": "v1", "state": "closed", "due_on": "2027-01-15T00:00:00Z"}})
        self.assertEqual(parsed.milestone_state, "closed")
        self.assertEqual(parsed.milestone_due, "2027-01-15T00:00:00Z")
        self.assertIn("milestone-closed", [f.rule for f in check_issues.planning_findings(parsed)])

    def test_target_boundaries_and_both_formats(self) -> None:
        for template in ("### 📅 Target date\n\n{}\n", "## 📅 Planning dates\n- **Target completion date:** {}\n"):
            for value, expected in (("2027-01-14", []), ("2027-01-15", []),
                                    ("2027-01-16", [(1, "target-after-milestone")])):
                with self.subTest(template=template, value=value):
                    candidate = replace(issue(1, body=READY + template.format(value)),
                                        milestone_state="open", milestone_due="2027-01-15T00:00:00Z")
                    self.assertEqual(rules([candidate]), expected)

    def test_optional_invalid_conflicting_and_unrelated_dates(self) -> None:
        for body in (READY, READY + "### Target date\n\n_No response_\n",
                     READY + "## Notes\nTarget date: 2099-01-01\n",
                     READY + "## Planning dates\n```\nTarget date: 2099-01-01\n```\n"):
            self.assertEqual(rules([issue(1, body=body)]), [])
        def due(body: str) -> Issue:
            return replace(issue(1, body=body), milestone_due="2027-12-31T00:00:00Z")

        for value in ("2027-02-30", "15/01/2027", "2027-1-15", "tomorrow"):
            self.assertEqual(rules([due(READY + "### Target date\n" + value)]), [(1, "target-date")])
        body = READY + "### Target date\n2027-01-15\n## Planning dates\n- Target date: 2027-01-16\n"
        self.assertEqual(rules([due(body)]), [(1, "target-date")])
        self.assertEqual(rules([due(body.replace("2027-01-16", "2027-01-15"))]), [])
        candidate = replace(issue(1, body=READY + "### Target date\n2027-01-15"), milestone_due="bad")
        self.assertEqual(rules([candidate]), [(1, "milestone-date")])
        self.assertEqual(rules([replace(candidate, state="closed", milestone_state="closed")]), [])


    def test_a_target_needs_a_milestone_with_a_due_date_to_be_checked_against(self) -> None:
        # The deadline is inclusive (docs/GOVERNANCE.md), so a target on the due date is fine; a milestone
        # with no due date cannot check any target, and says so.
        body = READY + "### Target date\n2027-01-15\n"
        self.assertEqual(rules([issue(1, body=body)]), [(1, "milestone-date")])
        self.assertEqual(rules([issue(1, body=READY)]), [])  # no target, nothing to check
        self.assertEqual(rules([issue(1, body=body, milestone=None)]), [(1, "milestone")])

    def test_the_start_date_is_required_valid_and_not_after_the_target(self) -> None:
        for body in (START + READY, "## Planning dates\n- Planned start date: `2027-01-04`\n" + READY,
                     "## Planning dates\n- **Start date:** 2027-01-04\n" + READY,
                     "### 🗓️ Start date\n\n2027-01-04\n" + READY):
            self.assertEqual(rules([issue(1, body=body, start=False)]), [], body)
        for body in (READY, "### 🗓️ Start date (Europe/Rome)\n\n_No response_\n" + READY,
                     "## Notes\nStart date: 2027-01-04\n" + READY,
                     "## Planning dates\n```\n- Start date: 2027-01-04\n```\n" + READY):
            self.assertEqual(rules([issue(1, body=body, start=False)]), [(1, "start-date")], body)
        for value in ("2027-02-30", "04/01/2027", "2027-1-4", "soon"):
            body = f"### 🗓️ Start date (Europe/Rome)\n\n{value}\n" + READY
            self.assertEqual(rules([issue(1, body=body, start=False)]), [(1, "start-date")], value)
        conflicting = START + "## Planning dates\n- Planned start date: 2027-01-05\n" + READY
        self.assertEqual(rules([issue(1, body=conflicting, start=False)]), [(1, "start-date")])
        late = replace(issue(1, body=START + READY + "### Target date\n2027-01-03\n"),
                       milestone_due="2027-12-31T00:00:00Z")
        self.assertEqual(rules([late]), [(1, "start-date")])
        same_day = replace(issue(1, body=START + READY + "### Target date\n2027-01-04\n"),
                           milestone_due="2027-12-31T00:00:00Z")
        self.assertEqual(rules([same_day]), [])
        self.assertEqual(rules([issue(1, "P3", body=READY, state="closed", start=False)]), [])


class DependencyTests(unittest.TestCase):
    def test_an_empty_dependency_section_is_a_finding(self) -> None:
        for body in ("## Dependencies and blocking status\n\n", "## Dependencies and blocking status\n\n_No response_\n",
                     "## Dependencies and blocking status\n   \n\n## Tracking\n\n#99\n"):
            self.assertEqual(rules([issue(1, body=body)]), [(1, "dependencies")], body)
        self.assertEqual(rules([issue(1, body="## Dependencies and blocking status\n\nNone\n")]), [])

    def test_blocked_issues_that_wait_for_each_other_are_a_cycle(self) -> None:
        a, b, c = (issue(n, "P3", BLOCKED_LABEL, body=blocked_by(m)) for n, m in ((1, 2), (2, 3), (3, 1)))
        self.assertEqual(rules([a, b, c]), [(1, "dependency-cycle"), (2, "dependency-cycle"), (3, "dependency-cycle")])
        findings = check_issues.cycle_findings([a, b, c], {1: "open", 2: "open", 3: "open"}, REPOSITORY)
        self.assertIn("#1, #2, #3", findings[0].message)
        # A chain is not a cycle, nor is a closed member.
        self.assertEqual(rules([a, b, issue(3, "P3", BLOCKED_LABEL, body=blocked_by(1), state="closed")]),
                         [(2, "blocked")])  # 2 waits only for a closed 3: it must be unblocked, and there is no cycle
        self.assertEqual(rules([issue(1, "P3", BLOCKED_LABEL, body=blocked_by(2)), issue(2)]), [])

    def test_related_issues_that_name_each_other_are_not_a_cycle(self) -> None:
        # A dependency section may name related issues: only a blocked issue waits for the ones it names.
        self.assertEqual(rules([issue(1, "P3", body=blocked_by(2)), issue(2, "P3", body=blocked_by(1))]), [])
        self.assertEqual(rules([issue(1, "P3", BLOCKED_LABEL, body=blocked_by(2)),
                                issue(2, "P3", body=blocked_by(1))]), [])

    def test_a_long_chain_does_not_exceed_the_recursion_limit(self) -> None:
        chain = [issue(n, "P3", BLOCKED_LABEL, body=blocked_by(n + 1)) for n in range(1, 3001)] + [issue(3001)]
        self.assertEqual(rules(chain), [])

    def test_blocked_by_an_open_prerequisite(self) -> None:
        self.assertEqual(rules([issue(2, "P3", BLOCKED_LABEL, body=blocked_by(1)), issue(1)]), [])

    def test_blocked_with_only_closed_prerequisites_must_be_unblocked(self) -> None:
        closed = issue(1, state="closed")
        self.assertEqual(rules([issue(2, "P3", BLOCKED_LABEL, body=blocked_by(1)), closed]),
                         [(2, "blocked")])

    def test_open_related_issue_does_not_require_blocked(self) -> None:
        # A dependency section may list related open issues; only blocked is checked.
        self.assertEqual(rules([issue(2, "P3", body=blocked_by(1)), issue(1)]), [])

    def test_external_condition_counts_as_a_blocker(self) -> None:
        body = "### 🔗 Dependencies and blocking status\n\nExternal condition: Excel host access\n"
        self.assertEqual(rules([issue(2, "P3", BLOCKED_LABEL, body=body)]), [])
        self.assertEqual(rules([issue(2, "P3", body=body)]), [(2, "blocked")])

    def test_references_outside_the_section_are_ignored(self) -> None:
        # blocked_by() also mentions #99 under Tracking; it must not count.
        self.assertEqual(rules([issue(2, "P3", BLOCKED_LABEL, body=blocked_by(1)), issue(1)],
                               {1: "open", 2: "open"}), [])

    def test_issue_urls_and_form_headings_are_recognised(self) -> None:
        body = ("### 🔗 Dependencies and blocking status\n\n"
                f"- https://github.com/{REPOSITORY}/issues/1\n")
        self.assertEqual(rules([issue(2, "P3", BLOCKED_LABEL, body=body), issue(1)]), [])

    def test_missing_section_and_unknown_reference(self) -> None:
        self.assertEqual(rules([issue(1, body="## Objective\n\nx\n")]), [(1, "dependencies")])
        self.assertEqual(rules([issue(2, "P3", BLOCKED_LABEL, body=blocked_by(7))]),
                         [(2, "dependencies"), (2, "blocked")])

    def test_self_reference_is_ignored(self) -> None:
        self.assertEqual(rules([issue(2, "P3", body=blocked_by(2))]), [])

    def test_named_pull_request_is_a_reference_not_a_prerequisite(self) -> None:
        body = "## Dependencies and blocking status\n\nNone. #1 was delivered in #22.\n"
        closed = issue(1, state="closed")
        self.assertEqual(rules([issue(2, "P3", body=body), closed], pulls=frozenset({22})), [])
        # A pull request never justifies blocked, and unknown numbers are still reported.
        self.assertEqual(rules([issue(2, "P3", BLOCKED_LABEL, body=blocked_by(22, 7))],
                               pulls=frozenset({22})),
                         [(2, "dependencies"), (2, "blocked")])



class ConfigTests(unittest.TestCase):
    def test_repository_configuration_passes(self) -> None:
        self.assertEqual(check_issues.config_findings(ROOT), [])

    def test_forms_require_target_field_and_keep_it_optional(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            shutil.copytree(ROOT / ".github/ISSUE_TEMPLATE", root / ".github/ISSUE_TEMPLATE")
            for name in check_issues.FORM_PRIORITIES:
                path = root / ".github/ISSUE_TEMPLATE" / name
                text = path.read_text(encoding="utf-8")
                field = text.split("id: target_date", 1)[1]
                self.assertIn('label: "📅 Target date"', field)
                self.assertIn("required: false", field)
                path.write_text(text.replace("id: target_date", "id: omitted"), encoding="utf-8")
            with unittest.mock.patch.object(check_issues, "catalogue_labels", return_value={
                    "bug", "enhancement", "documentation", "P1", "P2", "P3", "blocked"}):
                findings = check_issues.config_findings(root)
            self.assertEqual(sum("has no target_date field" in finding for finding in findings), 3)

    def test_an_invalid_manifest_is_rejected(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            shutil.copytree(ROOT / ".github/scripts", root / ".github/scripts")
            manifest = json.loads((ROOT / ".github/labels.json").read_text())
            manifest["overlays"] = {"selected": [{"name": "extra"}]}
            (root / ".github/labels.json").write_text(json.dumps(manifest))
            with self.assertRaisesRegex(ValueError, "root must contain exactly"):
                check_issues.catalogue_labels(root)

    def test_assignee_comes_from_configuration_not_repository_identity(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            forms = root / ".github/ISSUE_TEMPLATE"
            forms.mkdir(parents=True)
            (forms / "bug.yml").write_text('assignees: [triage-person]\n')
            self.assertEqual(check_issues.configured_assignee(root), "triage-person")
            (forms / "bug.yml").write_text('assignees: []\n')
            with self.assertRaises(ValueError):
                check_issues.configured_assignee(root)

    def test_flow_list(self) -> None:
        text = 'name: x\nlabels: ["bug", "P2"]\nassignees: [triage-person]\n'
        self.assertEqual(check_issues.flow_list(text, "labels"), ["bug", "P2"])
        self.assertEqual(check_issues.flow_list(text, "assignees"), ["triage-person"])
        self.assertEqual(check_issues.flow_list(text, "title"), [])


class MainTests(unittest.TestCase):
    def test_input_mode_reports_and_exits_nonzero_on_findings(self) -> None:
        items = [
            {"number": 1, "state": "open", "labels": [{"name": "P3"}],
             "assignees": [{"login": "triage-person"}], "milestone": None,
             "body": START + READY + "Delivered in #2, a pull request.\n"},
            {"number": 2, "state": "open", "labels": [], "assignees": [],
             "milestone": None, "body": READY, "pull_request": {}},
        ]
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "issues.json"
            source.write_text(json.dumps(items), encoding="utf-8")
            output = Path(directory) / "report.json"
            with contextlib.redirect_stdout(io.StringIO()) as stdout, \
                    unittest.mock.patch.object(check_issues, "configured_assignee",
                                               return_value="triage-person"):
                code = check_issues.main(["--root", str(ROOT), "--input", str(source),
                                          "--output", str(output), "--repo-name", REPOSITORY])
            self.assertEqual(code, 1)
            report = json.loads(output.read_text(encoding="utf-8"))
            self.assertEqual(report["status"], "fail")
            self.assertEqual([f["rule"] for f in report["findings"]], ["milestone"])
            self.assertIn("open issues audited: 1", stdout.getvalue())

    def test_offline_mode_requires_explicit_repository(self) -> None:
        with contextlib.redirect_stderr(io.StringIO()), self.assertRaises(SystemExit) as error:
            check_issues.parse_args(["--input", "issues.json"])
        self.assertEqual(error.exception.code, 2)

    def test_live_mode_requires_a_token(self) -> None:
        with contextlib.redirect_stderr(io.StringIO()), \
                unittest.mock.patch.dict("os.environ", {"GITHUB_TOKEN": ""}):
            self.assertEqual(check_issues.main(["--repository", REPOSITORY]), 2)


if __name__ == "__main__":
    unittest.main()
