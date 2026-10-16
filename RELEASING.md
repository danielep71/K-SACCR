# 🚀 K-SACCR Release Guide

This is the maintainer runbook for **releasing a version**. The branch model
and the integration of the release branch into `main` are in
[`docs/GOVERNANCE.md`](docs/GOVERNANCE.md#branches); day-to-day contribution is
in [`CONTRIBUTING.md`](CONTRIBUTING.md); version and changelog format are in
[`CHANGELOG.md`](CHANGELOG.md#date-and-version-rules).

> [!IMPORTANT]
> Nothing is released, versioned or tagged unless the owner asks for it
> explicitly. Closing a milestone does not authorize a release, and unreleased
> work stays under `## [Unreleased]` in the changelog.

<a id="current-status"></a>

## 🧭 Current status

| Property | State |
| --- | --- |
| Released versions | None |
| Active branch | See [`docs/GOVERNANCE.md`](docs/GOVERNANCE.md#branches) |
| Version file | Not present; created when the first release is prepared ([versioning](#versioning)) |
| Release evidence tooling | Static checks (`tools/check.py`) and the Excel evidence validator (`tools/check_excel_evidence.py`) |
| Excel certification | All eight suites listed in `.github/excel-evidence-policy.json`, recorded and validated as described in [`docs/EXCEL_EVIDENCE.md`](docs/EXCEL_EVIDENCE.md) |

Steps below marked *(to be defined)* are settled with the first distributed
artifact.

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
newest one; a `VERSION` without a dated release fails.

<a id="release-invariants"></a>

## 🔒 Release invariants

A release is valid only when:

1. one exact candidate SHA is frozen and reviewable;
2. the changelog entry for the version is dated and matches the scope;
3. `python tools/check.py --ci` passes on that candidate, and the hosted
   **Repository integrity** check is green;
4. VBA compiles and all eight suites listed in
   `.github/excel-evidence-policy.json` pass in Excel on that candidate,
   recorded and validated as described in
   [`docs/EXCEL_EVIDENCE.md`](docs/EXCEL_EVIDENCE.md);
5. any distributed artifact is built from and tested against that candidate;
6. an annotated `v*` tag targets the certified commit on `main`; and
7. post-publication checks pass.

If source changes after certification, the evidence is stale and must be
rerun. Never compensate by editing an already-tested artifact.

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
references, newest-first order and the `VERSION` match. Merge this through a
task PR into the release branch.

### 3. Check the social preview

Before the first release, check that the figures in the social preview image
still match its source notes in
[`docs/assets/README.md`](docs/assets/README.md#social-preview),
and update its "Under development" status as those notes describe.

### 4. Run the static gates

```sh
python tools/check.py --ci
```

Confirm the hosted **Repository integrity** check is green on the same SHA.

### 5. Certify in Excel

Import the exact candidate into a clean workbook, compile, and run all eight
suites listed in `.github/excel-evidence-policy.json`, each on its own,
following the [manual procedure](docs/EXCEL_EVIDENCE.md#manual-procedure).
Validate the bundle against the candidate SHA:

```sh
python tools/check_excel_evidence.py --candidate-sha FULL_SHA --evidence ../bundle/host.json
```

Only exit code 0 (`PASS`) certifies the candidate.

### 6. Build artifacts *(to be defined)*

If a workbook or add-in is distributed, build it from the certified source,
test the packaged file and record its SHA-256 hash. No artifact is distributed
today.

### 7. Integrate into main

Follow [integrating the release branch into main](docs/GOVERNANCE.md#integrating-the-release-branch-into-main).
If the merge changes the source identity, repeat the certification on the
merged SHA.

### 8. Create the annotated tag

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
Excel: compile <result>; suites <passed>/8, failed=0; evidence validator PASS.
Fixes #<n>, #<n>.
Deviations: none.
```

`Deviations:` is mandatory: write `none`, or list each gate that was not run.
Copy totals verbatim from the run; never round or reuse earlier figures. Never
delete and recreate a published tag.

### 9. Publish and verify

Publish the GitHub Release from the tag, attach only certified artifacts with
their hashes, and verify that a fresh clone of the tag passes the static checks
and imports cleanly.

## 🧾 Evidence record

Retain at least: candidate SHA, previous tag, static-check result, Excel
version and bitness, the compile result and the validated evidence bundle for
every suite, artifact names, sizes and SHA-256 hashes, the integration PR,
validation timestamps and any deviations. A release note summarizes evidence;
it does not replace it.

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
| 3 | Changelog and `VERSION` finalized | ☐ |
| 4 | Social preview checked (first release) | ☐ |
| 5 | Static checks pass locally and in CI | ☐ |
| 6 | Excel certification passes for all eight suites | ☐ |
| 7 | Artifacts built, tested and hashed (if any) | ☐ |
| 8 | Release branch merged into `main` with a merge commit | ☐ |
| 9 | Release branch level with `main` again | ☐ |
| 10 | Annotated tag targets the certified `main` SHA | ☐ |
| 11 | GitHub Release published and verified | ☐ |
