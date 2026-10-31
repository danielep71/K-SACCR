# Repository governance

Owner and maintainer: **@danielep71**. This document is the single home for
the branch model, issue metadata and triage, the GitHub settings and the issue
labels. Contribution steps are in [`CONTRIBUTING.md`](../CONTRIBUTING.md); the
release runbook is in [`RELEASING.md`](../RELEASING.md).

<a id="branches"></a>

## Branches

The active release branch is `release/1.1.0`.

| Branch | Purpose and allowed work |
| --- | --- |
| `release/1.1.0` | Active integration branch for all development. Retained after each integration and brought up to date with `main` through a pull request. |
| `release/1.0.2` | Deleted after owner-authorized cleanup. Published history remains available through tag `v1.0.2` and its release artifacts; all former branch commits are retained in `main`. |
| `release/1.0.1` | Deleted after owner-authorized cleanup. Published history remains available through tag `v1.0.1` and its release artifacts; all former branch commits are retained in `main`. |
| `release/1.0.0` | Deleted after release. Published history remains available through tag `v1.0.0` and its release artifacts. |
| `main` | Default branch. Receives the active release branch only through an owner-requested PR merged with a merge commit. Nothing is committed to main directly. |
| Short-lived task branches | Every change: branch from the active release branch, open the PR against it and squash-merge it. Use a descriptive name such as `fix/<issue>-<slug>`, `docs/<slug>` or `chore/<slug>`. |
| `traffic-history` | Retained for the traffic-export automation; not a release branch or cleanup target. |
| Other `release/*` branches | Reserved for explicitly agreed future milestones; do not infer a new development target from a version number. |

The default branch remains `main`; this does not make it the active development
branch. Check the PR base explicitly. Nothing is committed directly to the
release branch or to main, and nothing is integrated, versioned, tagged or
released without the owner's explicit authorization. Closing a milestone does
not publish a tag, release or workbook, or authorize a merge into main. Never
force-push or reset shared branches.

When a new release branch is created, update this section, the
release-deletion row under [GitHub settings](#github-settings),
`target-branch` in `.github/dependabot.yml`, which Dependabot reads from
`main`, the `git switch` command under
[Get the source](../INSTALLATION.md#get-the-source) in `INSTALLATION.md`, and
the section of the current milestone in the
[README roadmap](../README.md#roadmap), which names the branch.

<a id="integrating-the-release-branch-into-main"></a>

### Integrating the release branch into main

Done only when the owner asks. This is an integration, not a release, and it is
not opened just to record progress.

1. Confirm the release branch is green and has no open review findings.
2. Open a PR from the release branch into `main`.
3. Merge it with a **merge commit**, not a squash, once the hosted checks are
   green, so provenance remains visible.
4. Keep the release branch and bring it up to date with `main` through a
   second PR, from `main` into the release branch, merged with a **merge
   commit** once **Repository integrity** is green. **Protect release
   branches** has no bypass, so the branch cannot be moved by a direct push;
   the PR carries only the integration merge commit, never new development.
   Do not delete/recreate the branch or weaken the separate
   **Prevent release branch deletion** ruleset for this operation. At an
   owner-approved milestone transition, retain the completed branch and
   create the next `release/X.Y.Z` from the integration commit, then update
   [Branches](#branches) and Dependabot's `target-branch`.
5. Verify that the release branch contains the integration merge commit and
   that CI passed on both branches.

<a id="ownership-authorization-and-review"></a>

## Ownership, authorization and review

The owner sets scope, priorities, exceptions and release decisions. Contributors
work within the authorized issue. Every change reaches the active release
branch through a PR. Inspect the diff and run the applicable checks before
pushing. An author's self-check is not an independent review.

No change to visibility, paid subscriptions, security-sensitive access or
main-branch scope is implied by an ordinary development task. Raise a concrete
exception when needed rather than bypassing a requirement. Keep unresolved
review findings and failed checks visible; do not call unavailable validation
successful.

The review checklist is in the
[pull request template](../.github/PULL_REQUEST_TEMPLATE.md); it is the only
one. Local validation (`python tools/check.py`) precedes every push; run
`python tools/check.py --ci` on the clean committed candidate, with
`--base <base-sha>` for a PR. Hosted checks must be green before a PR is
merged. Report static checks and real Excel runs separately. If a hosted run
fails after a merge, record the failure and fix or revert it through a new PR
before closing the issue. Never rewrite shared history to conceal the failed
candidate.

<a id="issue-metadata-and-completion"></a>

## Issue metadata and completion

Every issue has assignee **danielep71**, exactly one priority label
(**P1**, **P2** or **P3**; see [Labels](#labels)) and the milestone of the
release branch it targets. Milestones are for issues only: a pull request
links its issue and never carries a milestone. The **Keep milestones
issue-only** workflow
([`.github/workflows/pr-milestones.yml`](../.github/workflows/pr-milestones.yml))
removes one when a pull request is opened, reopened, edited or given a
milestone, and clears any left on open or closed pull requests every day at
03:17 UTC. Include objective, acceptance checklist, a
**Dependencies and blocking status** section and required evidence. The
section names each prerequisite as an issue link (`#12`) or on a line
starting `External condition:`, or says `None`; it may also name related
issues. `blocked` is set only while a named prerequisite is open or an
external condition holds.

Exception: traffic alert issues opened by the daily traffic export get the
assignee and **P3** automatically and no milestone; they are operational
analytics, not development work, and are closed once reviewed.

| Field | How it is set |
| --- | --- |
| Assignee `danielep71` | Set by every issue form |
| Type label (`bug`, `enhancement`, `documentation`) | Set by the form |
| Priority | Each form sets a default (`P2`, or `P3` for documentation); triage confirms or changes it |
| Milestone | **Manual.** GitHub issue forms cannot set a milestone |

Forms apply to web issues created through a form. `blank_issues_enabled: false`
hides ordinary blank issues, but users with Write, Maintain or Admin access
can still use the **Maintainers only** blank-issue option. Those issues, and
issues created through the API, CLI or another tool, bypass the forms and get
no metadata automatically. Nothing blocks an issue that lacks metadata: this is
a **manual rule backed by the triage check** below, not enforcement. The issue forms
and the pull-request template are read by GitHub from the default branch, so a
change to them takes effect after the next integration into `main`.

### Triage check

The read-only **Issue triage** workflow
([`.github/workflows/issue-triage.yml`](../.github/workflows/issue-triage.yml))
runs `tools/check_issues.py` on every issue event and on manual dispatch. It
audits every open issue:

- assigned to the person `bug.yml` assigns;
- exactly one of `P1`, `P2` and `P3`, and `P1` only together with `bug`;
- a milestone;
- only labels of `.github/labels.json`;
- a dependency section whose issue references exist;
- `blocked` only with an open prerequisite or an external condition, and
  always with an external condition.

Traffic alerts need no milestone or dependency section. The workflow never
changes an issue: a failed run lists the findings, and I resolve each by hand.
Issue events run it from `main`, so a change to it takes effect after the next
integration. To audit locally, with an export of the issues as GitHub REST
JSON:

```sh
python tools/check_issues.py --input issues.json --repo-name danielep71/K-SACCR
```

The searches below also cover closed issues. Run these searches at issue triage and before milestone closeout. Each must
return no results. Closed issues are included because the metadata rule applies
to every issue. Keep the missing-milestone search repository-wide: filtering by
milestone would hide orphans.

- [Issues without a milestone](https://github.com/danielep71/K-SACCR/issues?q=is%3Aissue%20no%3Amilestone%20-author%3Aapp%2Fgithub-actions), excluding traffic alerts, which the workflow opens as `github-actions`
- [Issues without an assignee](https://github.com/danielep71/K-SACCR/issues?q=is%3Aissue%20no%3Aassignee)
- [Issues without a priority label](https://github.com/danielep71/K-SACCR/issues?q=is%3Aissue%20-label%3AP1%20-label%3AP2%20-label%3AP3)
- Issues with more than one priority label: [P1 and P2](https://github.com/danielep71/K-SACCR/issues?q=is%3Aissue%20label%3AP1%20label%3AP2), [P1 and P3](https://github.com/danielep71/K-SACCR/issues?q=is%3Aissue%20label%3AP1%20label%3AP3), [P2 and P3](https://github.com/danielep71/K-SACCR/issues?q=is%3Aissue%20label%3AP2%20label%3AP3)

### Closing an issue

Close an issue only when every acceptance criterion has linked evidence: the
merged PR, the green check run and, for VBA, the Excel result. Record that
evidence in a closing comment. A blocked Excel run keeps the issue open. Each
milestone closes through its own closeout issue.

<a id="github-settings"></a>

## GitHub settings

Settings is authoritative; recheck it before relying on a row or changing a
configuration.

| Control | State and working response |
| --- | --- |
| Repository visibility | Public. |
| Pull request settings | Updating a pull request branch from its base is offered. Projects and the legacy downloads feature are off. |
| Branch rulesets | **Protect main** (the default branch): pull request required with no approvals, conversations resolved, merge commits only, **Repository integrity** required, up to date and enforced on branch creation, no force pushes or deletion, no bypass. **Protect release branches** (`release/**`): the same, except that squash and merge commits are allowed and the check is not enforced when a branch is created; task PRs are squash-merged and the PR that brings the branch up to date with `main` is merged with a merge commit. **Protect traffic history** (`traffic-history`): no force pushes or deletion, no bypass. The separate **Prevent release branch deletion** ruleset blocks deletion of matching release branches with no bypass. Keep following the review checklist for what rulesets cannot check. |
| Release deletion coverage | **Prevent release branch deletion** is active for `release/**`, with no branch exclusions and no bypass actors. It protects the active `release/1.1.0` and future matching branches. Historical `release/1.0.0`, `release/1.0.1` and `release/1.0.2` have been deleted; their published tags and artifacts remain. Any future historical-branch deletion requires separate owner authorization, commit-reachability and link checks, and immediate removal of temporary ruleset exceptions. |
| Tag ruleset | `v*` tags: updates, deletions and force pushes blocked, so a published tag cannot move. |
| Automatic head-branch deletion | Enabled for eligible merged PR branches. It does not remove release branches covered by **Prevent release branch deletion**; retain the active release branch after integration and bring it up to date with `main` through a pull request. |
| Wiki | Editable by collaborators only. Maintained directly in the Wiki repository, outside this repository's source and checks (owner decision); it links to files and images of this repository only on `main`. Authoritative documentation stays in versioned `docs/`. |
| Private vulnerability reporting | Enabled; the primary reporting channel in [`SECURITY.md`](../SECURITY.md). |
| Code scanning | CodeQL default setup is enabled for `actions` and `python` with the default query suite: it analyses pushes and pull requests and runs weekly. It covers the workflows and the Python tools; VBA is not a CodeQL language. It is not a required check in the rulesets; review its alerts under Security → Code scanning. |
| Secret scanning and push protection | Enabled. Push protection blocks a push that contains a detected secret; never bypass it for a real credential, remove the secret and rotate it instead. Non-provider patterns and validity checks are off. |
| Dependency graph, Dependabot alerts and security updates | Enabled. Review alerts rather than assuming a clean dependency graph proves application security. Security update pull requests target `main`; see [Dependency updates](../tools/README.md#dependency-updates). |
| Low-impact development-dependency alert auto-dismissal | Disabled by explicit owner decision. Keep those alerts visible. |
| Automatic Dependabot updates | See [Dependency updates](../tools/README.md#dependency-updates). Alerts and update PRs are separate features. |
| Automated review | The automated code review reviews every pull request; its review conversations must be resolved before merge (ruleset), like any other finding. It does not replace the owner's review. |
| Repository integrity CI | See [GitHub Actions](../tools/README.md#github-actions). Static checks only; no Excel execution. |
| Labels | See [Labels](#labels). |
| Pull-request milestones | Removed by the **Keep milestones issue-only** workflow; see [Issue metadata and completion](#issue-metadata-and-completion). |
| Traffic history | See [Traffic history](../tools/README.md#traffic-history). Analytics only, never CI. |

<a id="labels"></a>

## Labels

K-SACCR uses the 20-label core catalogue from
danielep71/EXCEL-VBA-PROJECT-TEMPLATE at commit
b903fe44ef6a032c4689870b83745afa1c22490d. Names, colours and descriptions are
defined in `.github/labels.json`, schema version 2: its root holds only
`schema_version`, `prune` and the `labels` list, with no profiles or
overlays. Its `prune: true` policy removes labels outside the catalogue during
reconciliation.

### When to use each label

Every issue carries exactly one priority label. `P1` is reserved for a
confirmed calculation bug and is used only together with `bug`; any other
defect, however urgent, is `P2`. Issues closed before this rule keep the
priority they were closed with.

| Label | Use when |
| --- | --- |
| `P1` | Confirmed calculation bug: wrong results for valid inputs (reserved for this case) |
| `P2` | Material correctness, robustness, or maintainability defect |
| `P3` | Non-blocking quality, testing, documentation, setup or technical-debt work |

The forms default to `P2` (bug, feature) and `P3` (documentation); Dependabot
update PRs and traffic alerts get `P3`. Triage confirms or changes it.

Type labels say what kind of work an issue or PR is. The forms set `bug`,
`enhancement` or `documentation`; add others where they apply.

| Label | Use for |
| --- | --- |
| `bug` | Confirmed defect in behavior, correctness, error handling, or a documented contract |
| `enhancement` | New capability or meaningful improvement to supported behavior |
| `documentation` | Documentation for users or for my repository maintenance |
| `tests` | Regression coverage, test harnesses, fixtures, or manual validation |
| `refactor` | Internal restructuring intended to preserve supported behavior |
| `behavior-change` | Intentional externally observable behavior change |
| `ci` | Automated quality checks, workflow maintenance, or release certification |
| `repository` | Repository structure, governance, metadata, tooling, or maintenance |
| `release` | Release preparation, certification, evidence, packaging, or publication |
| `security` | Security boundary, vulnerability handling, supply chain, or hardening; never for undisclosed vulnerabilities, which follow `SECURITY.md` |

Status and triage labels:

| Label | Use when |
| --- | --- |
| `blocked` | Cannot proceed until a named dependency is complete |
| `question` | I need further information or need to make a decision |
| `help wanted` | I welcome external investigation, implementation, validation, or review |
| `good first issue` | Well-scoped task suitable for a first contribution |
| `duplicate` | An existing issue or pull request already covers this work |
| `invalid` | Report is incomplete, unsupported, or not actionable as submitted |
| `wontfix` | I have decided not to implement or continue this work |

Labels do not assign priorities or milestones automatically.

### Label workflows

Both reconciliation scripts are the schema-2 versions of
danielep71/MyTemplate at commit f4e11c3e5ced16fb3d699a31a21e5e9d7ade7173,
unchanged. They read `.github/labels.json` and take no repository policy, so
the live catalogue is exactly the manifest's 20 labels. Every checkout in the
two workflows sets `persist-credentials: false`, so code under review never
receives Git credentials. The workflows are K-SACCR's own in two respects:
pull-request checks run only for pull requests into `main` or `release/**`.
The reconciliation job runs only on the default branch, for a push to `main`
or a manual run on `main`; a manual run on any other branch skips it.

- **Sync issue labels** (`labels-sync.yml`) validates pull requests into `main`
  or `release/**` without write permission. Pushes to main affecting its inputs,
  and manual runs on main, reconcile live labels and verify an exact match. Only the
  reconciliation job receives `issues: write`.
- **Detect issue-label drift** (`labels-drift.yml`) checks deterministic
  fixtures on pull requests and compares live labels daily at 05:17 UTC or on
  manual request. It is read-only and uploads JSON/Markdown evidence for 30
  days. Drift causes failure; detection does not silently repair it.
- Both workflows use full-SHA-pinned actions.

Changes merged into the release branch alone do not change the live labels:
they take effect when the release branch is integrated into `main`. Develop a
label-catalogue change on the release branch, validate it locally and through
the pull-request checks, and review removals and their issue associations
before integration. Do not modify main solely to trigger sync.

To update the catalogue, edit the manifest through a pull request. Review
removals carefully: deleting a label also removes its issue/PR associations.
Optional local checks, with Node.js:

```sh
node .github/scripts/labels-sync.mjs --self-test
node .github/scripts/labels-sync.mjs --mode validate
node .github/scripts/labels-drift.mjs --self-test
```
