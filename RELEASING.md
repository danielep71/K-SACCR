# 🚀 K-SACCR Release Guide

[![Release model: exact source](https://img.shields.io/badge/release-exact%20source-0969da)](#release-invariants)
[![Versioning: SemVer](https://img.shields.io/badge/versioning-SemVer-3f4551)](CHANGELOG.md#date-and-version-rules)
[![Status: pre-release](https://img.shields.io/badge/status-pre--release-6e7781)](#current-status)
[![Security policy](https://img.shields.io/badge/security-private-d73a49)](SECURITY.md)

This document is authoritative for the **maintainer integration and release
sequence**. Day-to-day contribution is owned by
[`CONTRIBUTING.md`](CONTRIBUTING.md); branch and review policy by
[`docs/GOVERNANCE.md`](docs/GOVERNANCE.md); version and changelog format by
[`CHANGELOG.md`](CHANGELOG.md#date-and-version-rules).

> [!IMPORTANT]
> Nothing is released, versioned, tagged or stamped unless the owner asks for it
> explicitly. Closing a milestone does not authorize a release, and unreleased
> work stays under `## [Unreleased]` in the changelog.

<a id="current-status"></a>

## 🧭 Current status

| Property | State |
| --- | --- |
| Released versions | None |
| Active branch | `release/1.0.0` |
| Version file | Not present; created when the first release is prepared ([versioning](#versioning)) |
| Release evidence tooling | Static checks (`tools/check.py`) and the Excel evidence validator (`tools/check_excel_evidence.py`) |
| Excel certification procedure | Harness `TEST_Harness.RunTests`; evidence bundle and validator in [`docs/EXCEL_EVIDENCE.md`](docs/EXCEL_EVIDENCE.md) |

Steps below marked *(to be defined)* are settled with the first distributed
artifact. Repository setup, milestone v0.0.1, closed without a release; its
baseline is recorded in [issue #13](https://github.com/danielep71/VBA-SACCR-Toolkit/issues/13).

## 🌿 Branch model

| Branch | Role |
| --- | --- |
| `release/<version>` | Active integration branch. Task PRs are squash-merged into it. |
| `main` | Default branch. Receives the release branch only through an owner-requested PR. |
| Task branches | One per change, from the release branch, deleted automatically after merge. |

<a id="versioning"></a>

## 🏷️ Versioning

| Item | Convention |
| --- | --- |
| Version | Semantic Versioning `MAJOR.MINOR.PATCH`; rules in [`CHANGELOG.md`](CHANGELOG.md#date-and-version-rules) |
| Release branch | `release/X.Y.Z`, without `v`, opened only by owner decision |
| Milestone | `vX.Y.Z - <title>`, matching the release branch |
| `VERSION` | Root file holding one `X.Y.Z` line. Created in the release-preparation PR, never before; afterwards it always names the newest released version |
| Changelog | Work stays under `## [Unreleased]` until release preparation is approved; the release PR moves it to `## [X.Y.Z] - YYYY-MM-DD` |
| Tag | Annotated `vX.Y.Z` on the certified commit on `main`; protected by the `v*` tag ruleset |
| GitHub Release | Created from the tag, with certified artifacts only |

`python tools/check.py` enforces the link between `VERSION` and the changelog:
once the changelog has a dated release, `VERSION` must exist and equal the
newest one; a `VERSION` without a dated release fails. Nothing is versioned,
tagged or released without the owner's explicit request.

<a id="release-invariants"></a>

## 🔒 Release invariants

A release is valid only when:

1. one exact candidate SHA is frozen and reviewable;
2. the changelog entry for the version is dated and matches the scope;
3. `python tools/check.py --ci` passes on that candidate, and the hosted
   **Repository integrity** check is green;
4. VBA compile, the regression harness and the release scenarios pass in Excel
   on that candidate;
5. any distributed artifact is built from and tested against that candidate;
6. an annotated `v*` tag targets the certified commit on `main`; and
7. post-publication checks pass.

If source changes after certification, the evidence is stale and must be
rerun. Never compensate by editing an already-tested artifact.

<a id="integrating-the-release-branch-into-main"></a>

## 🔀 Integrating the release branch into main

Done only when the owner asks. This is an integration, not a release.

1. Confirm the release branch is green and has no open review findings.
2. Open a PR from the release branch into `main`.
3. Merge it with a **merge commit**, not a squash, once CI is green.
4. Bring the release branch back level with `main`:
   - if the release branch still exists, fast-forward it to `main`;
   - if automatic branch deletion removed it, recreate it from `main` at the
     merge commit;
   - at a milestone closeout, open the next `release/X.Y.Z` branch from the
     merge commit instead, and point Dependabot's `target-branch` and the docs
     at it.
5. Verify both branches point to the merge commit and that CI passed on it.

Never force-push or reset either branch.

## 📦 Releasing a version

Performed only on the owner's explicit request.

### 1. Freeze and identify the candidate

Record the release-branch SHA, the previous tag (if any) and the intended
version. Verify that the branch rulesets are still active in Settings → Rules.

### 2. Finalize the changelog

Move the `[Unreleased]` entries into `## [X.Y.Z] - YYYY-MM-DD`, add the link
reference for the new version, keep an empty `[Unreleased]` section, and create
or update `VERSION` to `X.Y.Z` in the same PR.
`tools/check_source.py` enforces the heading format, real calendar dates, link
references, newest-first order and the `VERSION` match. Merge this through a task PR into the release branch.

### 3. Run the static gates

```sh
python tools/check.py --ci
```

Confirm the hosted **Repository integrity** check is green on the same SHA.

### 4. Certify in Excel

Import the exact candidate into a clean workbook, compile, run the full
regression harness and the release scenarios, following the manual procedure in
[`docs/EXCEL_EVIDENCE.md`](docs/EXCEL_EVIDENCE.md#manual-procedure). Validate the
bundle against the candidate SHA:

```sh
python tools/check_excel_evidence.py --candidate-sha FULL_SHA --evidence ../bundle/host.json
```

Only exit code 0 (`PASS`) certifies the candidate. The release scenarios are
recorded alongside it until the harness covers them.

### 5. Build artifacts *(to be defined)*

If a workbook or add-in is distributed, build it from the certified source,
test the packaged file and record its SHA-256 hash. No artifact is distributed
today.

### 6. Integrate into main

Follow [Integrating the release branch into main](#integrating-the-release-branch-into-main).
If the merge changes the source identity, repeat the certification on the
merged SHA.

### 7. Create the annotated tag

Tag only the certified commit on `main`, from the command line so the tag is
annotated:

```sh
git switch main
git pull --ff-only
git rev-parse HEAD
git tag -a vX.Y.Z -F ../tag-message-vX.Y.Z.txt
git cat-file -t vX.Y.Z        # must print: tag
git push origin vX.Y.Z
```

The tag message is the release's permanent record:

```text
vX.Y.Z - <short release title>

Certified at <full 40-character SHA>.
Excel: compile <result>; harness <passed>/<run>, Failed=0.
Fixes #<n>, #<n>.
Deviations: none.
```

`Deviations:` is mandatory: write `none`, or list each gate that was not run.
Copy totals verbatim from the run; never round or reuse earlier figures. Never
delete and recreate a published tag.

### 8. Publish and verify

Publish the GitHub Release from the tag, attach only certified artifacts with
their hashes, and verify that a fresh clone of the tag passes the static checks
and imports cleanly.

## 🧾 Evidence record

Retain at least: candidate SHA, previous tag, static-check result, Excel
version and bitness, compile and harness results, artifact names, sizes and
SHA-256 hashes, the integration PR, validation timestamps and any deviations.
A release note summarizes evidence; it does not replace it.

## 🧯 Recovery

- **Before tagging:** fix through a task PR, update the evidence and rerun the
  affected gates.
- **After tagging, before publishing:** pause and document the correction;
  prefer a new patch version if the tag may have propagated.
- **After publishing:** do not replace assets or move the tag. Publish a
  corrected patch, record the problem in the changelog, and use
  [`SECURITY.md`](SECURITY.md) for vulnerabilities.

## ☑️ Maintainer checklist

| # | Gate | Done |
| ---: | --- | :---: |
| 1 | Owner requested the release | ☐ |
| 2 | Scope frozen and diff reviewed | ☐ |
| 3 | Changelog finalized | ☐ |
| 4 | Static checks pass locally and in CI | ☐ |
| 5 | Excel certification passes | ☐ |
| 6 | Artifacts built, tested and hashed (if any) | ☐ |
| 7 | Release branch merged into `main` with a merge commit | ☐ |
| 8 | Release branch level with `main` again | ☐ |
| 9 | Annotated tag targets the certified `main` SHA | ☐ |
| 10 | GitHub Release published and verified | ☐ |

## 📚 Related documents

- [`CHANGELOG.md`](CHANGELOG.md) — history and version rules
- [`CONTRIBUTING.md`](CONTRIBUTING.md) — contribution workflow
- [`INSTALLATION.md`](INSTALLATION.md) — setup, import and validation record
- [`docs/GOVERNANCE.md`](docs/GOVERNANCE.md) — branches, issues and controls
- [`SECURITY.md`](SECURITY.md) — vulnerability handling
