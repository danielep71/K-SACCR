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
| `release/1.1.0` | Active integration branch for all development. Retained and fast-forwarded to `main` after each integration. |
| `release/1.0.2` | Closed: released as `v1.0.2`. Kept for history; takes no new work. |
| `release/1.0.1` | Closed: released as `v1.0.1`. Kept for history; takes no new work. |
| `release/1.0.0` | Deleted after release. Published history remains available through tag `v1.0.0` and its release artifacts. |
| `main` | Default branch. Receives the active release branch only through an owner-requested PR merged with a merge commit. Nothing is committed to main directly. |
| Short-lived task branches | Every change: branch from the active release branch, open the PR against it and squash-merge it. Use a descriptive name such as `fix/<issue>-<slug>`, `docs/<slug>` or `chore/<slug>`. |
| Other `release/*` branches | Reserved for explicitly agreed future milestones; do not infer a new development target from a version number. |

The default branch remains `main`; this does not make it the active development
branch. Check the PR base explicitly. Nothing is committed directly to the
release branch or to main, and nothing is integrated, versioned, tagged or
released without the owner's explicit authorization. Closing a milestone does
not publish a tag, release or workbook, or authorize a merge into main. Never
force-push or reset shared branches.

When a new release branch is created, update this section and
`target-branch` in `.github/dependabot.yml`, which Dependabot reads from
`main`.

<a id="integrating-the-release-branch-into-main"></a>

### Integrating the release branch into main

Done only when the owner asks. This is an integration, not a release, and it is
not opened just to record progress.

1. Confirm the release branch is green and has no open review findings.
2. Open a PR from the release branch into `main`.
3. Merge it with a **merge commit**, not a squash, once the hosted checks are
   green, so provenance remains visible.
4. Keep the release branch and fast-forward it to the integration merge
   commit on `main`. Check its current head first: it must be an ancestor of
   that commit. Use an expected-head check when updating the ref; if new work
   arrived after the integration, stop and reconcile it rather than resetting
   or force-pushing. The existing administrator bypass of **Protect release
   branches** permits this ref-only alignment; it does not authorize direct
   development commits. Do not delete/recreate the branch or weaken the
   separate **Prevent release branch deletion** ruleset for this operation.
   At an owner-approved milestone transition, retain the completed branch
   and create the next `release/X.Y.Z` from the integration commit, then update
   [Branches](#branches) and Dependabot's `target-branch`.
5. Verify both branches point to the merge commit and that CI passed on it.

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
release branch it targets. Include objective, acceptance checklist,
dependencies and required evidence.

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
a **manual rule backed by the triage check**, not enforcement. The issue forms
and the pull-request template are read by GitHub from the default branch, so a
change to them takes effect after the next integration into `main`.

### Triage check

Run these searches at issue triage and before milestone closeout. Each must
return no results. Closed issues are included because the metadata rule applies
to every issue. Keep the missing-milestone search repository-wide: filtering by
milestone would hide orphans.

- [Issues without a milestone](https://github.com/danielep71/VBA-SACCR-Toolkit/issues?q=is%3Aissue%20no%3Amilestone%20-author%3Aapp%2Fgithub-actions), excluding traffic alerts, which the workflow opens as `github-actions`
- [Issues without an assignee](https://github.com/danielep71/VBA-SACCR-Toolkit/issues?q=is%3Aissue%20no%3Aassignee)
- [Issues without a priority label](https://github.com/danielep71/VBA-SACCR-Toolkit/issues?q=is%3Aissue%20-label%3AP1%20-label%3AP2%20-label%3AP3)
- Issues with more than one priority label: [P1 and P2](https://github.com/danielep71/VBA-SACCR-Toolkit/issues?q=is%3Aissue%20label%3AP1%20label%3AP2), [P1 and P3](https://github.com/danielep71/VBA-SACCR-Toolkit/issues?q=is%3Aissue%20label%3AP1%20label%3AP3), [P2 and P3](https://github.com/danielep71/VBA-SACCR-Toolkit/issues?q=is%3Aissue%20label%3AP2%20label%3AP3)

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
| Branch rulesets | `main`: pull request required, merge commits only, conversations resolved, **Repository integrity** required and up to date, no force pushes or deletion; repository admins may bypass only through a pull request. **Protect release branches**: pull request required, squash only, conversations resolved, **Repository integrity** required, no force pushes or deletion, with an administrator bypass. The separate **Prevent release branch deletion** ruleset blocks deletion of matching release branches with no bypass; the administrator exception in the other ruleset does not override it. Keep following the review checklist for what rulesets cannot check. |
| Release deletion coverage | **Prevent release branch deletion** is active for `release/**`, with no branch exclusions and no bypass actors. It protects the retained historical branches (`release/1.0.1` and `release/1.0.2`), the active `release/1.1.0`, and future matching branches. The removed `release/1.0.0` exception does not recreate that branch. Retain historical release branches; recheck Settings before any separately authorized cleanup. |
| Tag ruleset | `v*` tags: updates, deletions and force pushes blocked, so a published tag cannot move. |
| Automatic head-branch deletion | Enabled for eligible merged PR branches. It does not remove release branches covered by **Prevent release branch deletion**; retain and fast-forward the active release branch after integration. |
| Wiki | Editable by collaborators only. Maintained directly in the Wiki repository, outside this repository's source and checks (owner decision); it links to files and images of this repository only on `main`. Authoritative documentation stays in versioned `docs/`. |
| Private vulnerability reporting | Enabled; the primary reporting channel in [`SECURITY.md`](../SECURITY.md). |
| Advanced code/secret scanning | Available for public repositories. Do not claim CodeQL, secret scanning or push protection are active unless Settings → Code security shows them enabled. |
| Dependency graph and Dependabot alerts | Enabled. Review alerts rather than assuming a clean dependency graph proves application security. |
| Low-impact development-dependency alert auto-dismissal | Disabled by explicit owner decision. Keep those alerts visible. |
| Automatic Dependabot updates | See [Dependency updates](../tools/README.md#dependency-updates). Alerts and update PRs are separate features. |
| Automated review | The automated code review reviews every pull request; its review conversations must be resolved before merge (ruleset), like any other finding. It does not replace the owner's review. |
| Repository integrity CI | See [GitHub Actions](../tools/README.md#github-actions). Static checks only; no Excel execution. |
| Labels | See [Labels](#labels). |
| Traffic history | See [Traffic history](../tools/README.md#traffic-history). Analytics only, never CI. |

<a id="labels"></a>

## Labels

K-SACCR uses the 20-label core catalogue from
danielep71/EXCEL-VBA-PROJECT-TEMPLATE at commit
b903fe44ef6a032c4689870b83745afa1c22490d. Names, colours and descriptions are
defined in `.github/labels.json`; its `prune: true` policy removes labels
outside the catalogue during reconciliation.

### When to use each label

Every issue carries exactly one priority label:

| Label | Use when |
| --- | --- |
| `P1` | Critical release blocker or serious correctness, data, or application-state risk |
| `P2` | Material correctness, robustness, or maintainability defect |
| `P3` | Non-blocking quality, testing, documentation, or technical-debt improvement |

The forms default to `P2` (bug, feature) and `P3` (documentation); Dependabot
update PRs and traffic alerts get `P3`. Triage confirms or changes it.

Type labels say what kind of work an issue or PR is. The forms set `bug`,
`enhancement` or `documentation`; add others where they apply.

| Label | Use for |
| --- | --- |
| `bug` | Confirmed defect in behavior, correctness, error handling, or a documented contract |
| `enhancement` | New capability or meaningful improvement to supported behavior |
| `documentation` | User-facing or maintainer-facing documentation |
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
| `question` | Further information or a maintainer decision is required |
| `help wanted` | Maintainer welcomes external investigation, implementation, validation, or review |
| `good first issue` | Well-scoped task suitable for a first contribution |
| `duplicate` | An existing issue or pull request already covers this work |
| `invalid` | Report is incomplete, unsupported, or not actionable as submitted |
| `wontfix` | Maintainers have decided not to implement or continue this work |

Labels do not assign priorities or milestones automatically.

### Label workflows

The label manifest and both reconciliation scripts are copied unchanged from
the template. The two workflows are copied from the same revision with two
integration adjustments: they pass `--policy .github/label-policy.json` and
watch that file instead of the template-wide repository profile, and every
checkout sets `persist-credentials: false`, so code under review never receives
Git credentials. The policy selects the generated application profile with no
domain overlays; all profile overlays are empty, so the result is exactly the
template's 20 core labels.

- **Sync issue labels** (`labels-sync.yml`) validates pull requests into `main`
  or `release/**` without write permission. Pushes to main affecting its inputs,
  and manual runs, reconcile live labels and verify an exact match. Only the
  reconciliation job receives `issues: write`.
- **Detect issue-label drift** (`labels-drift.yml`) checks deterministic
  fixtures on pull requests and compares live labels daily at 05:17 UTC or on
  manual request. It is read-only and uploads JSON/Markdown evidence for 30
  days. Drift causes failure; detection does not silently repair it.
- Both workflows use full-SHA-pinned actions.

Changes merged into the release branch alone do not trigger live sync. If an
authorized label-policy change is developed on the release branch, validate it
locally and via PR checks, review removals and their issue associations, then
use an explicitly requested manual sync on that reviewed release ref when live
reconciliation is needed. Do not modify main solely to trigger sync.

To update the catalogue, edit the manifest through a pull request. Review
removals carefully: deleting a label also removes its issue/PR associations.
Optional local checks, with Node.js:

```sh
node .github/scripts/labels-sync.mjs --policy .github/label-policy.json --self-test
node .github/scripts/labels-sync.mjs --policy .github/label-policy.json --mode validate
node .github/scripts/labels-drift.mjs --policy .github/label-policy.json --self-test
```
