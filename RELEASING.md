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
| Released versions | 1.2.0 (2026-11-07), tag `v1.2.0`; 1.1.0 (2026-11-05), tag `v1.1.0`; 1.0.2 (2026-10-23), tag `v1.0.2`; 1.0.1 (2026-10-18), tag `v1.0.1`; 1.0.0 (2026-10-17), tag `v1.0.0` |
| Active branch | See [`docs/GOVERNANCE.md`](docs/GOVERNANCE.md#branches) |
| Version file | `VERSION`, holding `1.2.0` ([versioning](#versioning)) |
| Release evidence tooling | Static checks (`tools/check.py`) and the Excel evidence validator (`tools/check_excel_evidence.py`) |
| Excel certification | The harness and the nine suites listed in `.github/excel-evidence-policy.json`, recorded and validated as described in [`docs/EXCEL_EVIDENCE.md`](docs/EXCEL_EVIDENCE.md) |


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
4. VBA compiles and the harness and the nine suites listed in
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

The candidate is the merge commit that brings the release branch into `main`:
it is certified, the artifacts are built from it, and the tag is put on it, so
all three name one SHA.

### 1. Run the maintenance audit

Before the scope is fixed, check that GitHub tells the truth about the
work. The audit corrects issues, comments and other repository text; it
changes no code, merges no PR, changes no setting, deletes no branch and
publishes nothing. Read [`docs/GOVERNANCE.md`](docs/GOVERNANCE.md) and
[`CONTRIBUTING.md`](CONTRIBUTING.md) first, and go through every page of each
list, not only the first.

Judge each item by its evidence: the current code, the merged PRs, the CI and
Excel results, the linked discussions and the dependencies. A title, a label,
a ticked checkbox, a merged PR or a green check is not proof on its own that
every acceptance criterion is met.

1. **Closed issues in the release milestone and any other open milestone.**
   Check that each one's scope and acceptance criteria are really met: title,
   description, labels (priority, `blocked`), assignee, milestone, checkboxes,
   comments, linked PRs, dependencies and follow-ups. Correct stale metadata
   and text. Reopen an issue closed too early and say exactly what remains.
   Tell completed work apart from duplicates, rejected proposals and work
   deferred or abandoned on purpose, and close each with the matching reason.
   Never tick a criterion that has not been verified.
2. **Open issues.** Check that each one is still outstanding work, with the
   same fields. Bring stale descriptions and progress up to date and fix
   broken or obsolete links. Keep `blocked` only while a named prerequisite is
   still open or an external condition still holds, and remove it once the
   dependency is resolved. Close an issue only with the evidence that it is
   done; for a duplicate, name the canonical issue and carry over any
   requirement only the duplicate had.
3. **PR comments and review threads**, on open and closed PRs. Look for wrong
   claims, stale instructions, unresolved findings and misleading statements
   about the implementation, the tests or the review. Check review evidence
   against the commit it names, and tell a self-review, an automated review,
   CI and a manual test apart. Correct an editable comment where needed. A
   statement that was true when written stays; a later change is noted next
   to it. Never edit another contributor's words, or a bot's status summary,
   so that it implies an approval or a completion. Resolve a thread only when
   its concern is addressed or explicitly accepted, with the evidence; a
   finding that needs a code change is reported, not resolved.
4. **Notifications** of this repository, with their discussions and review
   threads. Handle what the audit covers. Mark a notification done only when
   nothing remains to do; leave an actionable or blocked one pending, saying
   what remains. Notifications of other repositories are not touched.
5. **Repository scope and wording.** Look through the documentation,
   configuration, templates, issues, PR descriptions and comments for text
   copied from another repository. Remove or correct references that do not
   apply, do not fit or are wrong. Keep real dependencies, source
   attribution, provenance, licence notices, comparisons and deliberate links
   to other repositories. In editable text written for the maintainer, prefer
   direct, impersonal wording, or the first person, to references in the third
   person such as "the owner" or "the user". Keep the records of authorship,
   responsibility and approval that are needed. Never invent a statement or
   an approval, and never change a quotation or another contributor's words.

Apply routine corrections that the evidence supports directly. Leave an item
unchanged when it is ambiguous, lacks evidence or needs a decision, and report
it. Avoid speculative changes and duplicate comments.

End with a report:

- what was audited and how much: issues, PRs and comments, and notifications
  examined;
- the changes, by category, with links and a short reason each;
- the issues reopened or closed, the dependencies corrected and the review
  threads resolved;
- the notifications marked done, and those left pending with what remains;
- the findings still open and the next actions recommended;
- what could not be examined, and why.

Keep verified facts apart from assumptions. If pagination, permissions or
missing evidence left something unexamined, the report says so and does not
claim a complete audit. A finding that needs code is fixed through a task PR
before step 2; an issue reopened for this release stays in its scope or is
moved out of the milestone by decision, never left behind silently.

### 2. Finalize the changelog

Move the `[Unreleased]` entries into `## [X.Y.Z] - YYYY-MM-DD`, add the link
reference for the new version, keep an empty `[Unreleased]` section, and create
or update `VERSION` to `X.Y.Z` in the same PR.
`tools/check_source.py` enforces the heading format, real calendar dates, link
references, newest-first order and the `VERSION` match. Merge this through a
task PR into the release branch.

### 3. Check the social preview

Check that the figures in the social preview image still match its source
notes in [`docs/assets/README.md`](docs/assets/README.md#social-preview), and
update the release it names as those notes describe.

### 4. Integrate into main and freeze the candidate

Follow [integrating the release branch into main](docs/GOVERNANCE.md#integrating-the-release-branch-into-main),
which leaves the release branch level with `main`. The merge commit on `main`
is the candidate: record its full SHA, the previous tag (if any) and the
intended version, and verify that the rulesets are still active in
Settings → Rules. Until the tag exists, the release branch takes only fixes
for a gate that failed on the candidate (steps 5 to 7). Such a fix records its
change in the dated `## [X.Y.Z]` section being released, not under
`[Unreleased]`; each is integrated again and the new merge commit becomes the
candidate.

### 5. Run the static gates

```sh
python tools/check.py --ci
```

Confirm the hosted **Repository integrity** check is green on the same SHA.

### 6. Certify in Excel

On a clean checkout of the candidate, run the one-command certification:

```powershell
git switch --detach FULL_SHA
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tools\Certify-SACCR.ps1
```

It builds a development workbook from the candidate, compiles it, runs the
harness and the nine suites listed in `.github/excel-evidence-policy.json`,
each on its own, and `Test_RunAll`, writing every log into a new bundle next
to the checkout. It then runs the
[manual protection checks](docs/workbook/WORKBOOK_PROTECTION.md#manual-checks)
on a production workbook built from the candidate, writes each outcome to
`protection-checks.txt` in the bundle, and finally writes `host.json` and
validates the bundle
([one-command certification](docs/EXCEL_EVIDENCE.md#one-command-certification);
the [manual procedure](docs/EXCEL_EVIDENCE.md#manual-procedure) remains the
fallback). Its exit code is 0 only when the bundle validates and the 14 checks
pass. To validate the bundle again later:

```sh
python tools/check_excel_evidence.py --candidate-sha FULL_SHA --evidence ../saccr-evidence-SHORTSHA/host.json
```

Only exit code 0 (`PASS`) certifies the candidate. If it fails, fix through a
task PR into the release branch, integrate again and certify the new merge
commit. The quick run of `tools/Verify-SACCR-Branch.ps1 -Quick` is a
pull-request check and writes no record: it never certifies a candidate.

<a id="ribbon-walk-through"></a>

**Then walk through the Ribbon by hand.** Certification cannot press the
controls of the K-SACCR tab or read its labels, so on the development
workbook of the bundle, opened in a visible Excel with macros enabled, check
the following and record each result with the release evidence
([Ribbon contract](docs/workbook/RIBBON.md)):

1. **Packaging.** The workbook opens on Home with no repair prompt and no
   prompt beyond Excel's own macro notification. The K-SACCR tab (**Alt, K,
   S**) shows its six groups: Calculation, CSV import, Evidence, Parameters,
   Status and Navigate.
2. **Commands.** Each command runs its macro: Run, and Run SA-CCR and Validate
   inputs in its menu; Clear outputs; Import CSV > NettingSets file... and
   Trades file...; Validate staging; Commit import; Templates >
   NettingSets template..., Trades template... and Import specification...;
   Export evidence; Restore parameters. A completed run opens Results,
   Validate inputs opens Checks, Clear outputs keeps the sheet.
3. **Go to.** The primary click opens Home, and every menu entry opens its
   sheet, in six sections in tab order, README last.
4. **Navigation bar.** The toggle hides column A and the bar on every sheet,
   and shows them again, without marking the workbook changed; a save while
   the bar is hidden still saves it shown.
5. **Enabled states and labels.** With both staging sheets `EMPTY`, Validate
   staging and Commit import are disabled; after an import Validate staging
   is enabled; Commit import only once a staging sheet reads `VALID`. Export
   evidence is disabled before a run and while Results reads `OUT OF DATE`,
   and enabled again when an input edit is undone. The three Status labels
   follow a load, an edit of a staged cell, a validation, a commit, a run, a
   clear and Restore parameters, also when started from the staging buttons.
6. **After a VBA reset.** Reset the project in the Visual Basic editor
   (Run > Reset), then run an operation: the labels and states still follow.
7. **Restore parameters.** After a maintenance unlock change `Alpha` and one
   factor: the confirmation lists both, Cancel changes nothing, OK restores
   them and Results reads `OUT OF DATE`; the next run is `REGULATORY`; a
   second restore says nothing changed.
8. **CSV workflow from the Ribbon.** Import both example files of the import
   specification, correct a staged value, validate until both sheets read
   `VALID`, commit and run: the run uses all 5 legs of the 3 netting sets.
   Commit import refuses after a staged edit, from the tab and from the
   staging button.
9. **Buttons.** Only the 8 buttons of `Import_NettingSets` and
   `Import_Trades` remain on the sheets.
10. **Saving.** Open the workbook and close it without changes: Excel asks
    nothing, apart from the prompt outside K-SACCR that the CHANGELOG records
    under #620. Save it and close it: the saved file holds no name
    `ksRibbonPointer`. In PowerShell this prints `False`:

    ```powershell
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip = [IO.Compression.ZipFile]::OpenRead((Resolve-Path .\K-SACCR_Certification.xlsm))
    (New-Object IO.StreamReader($zip.GetEntry("xl/workbook.xml").Open())).ReadToEnd() -match "ksRibbonPointer"
    $zip.Dispose()
    ```

The protection, the evidence export, the Home links with macros disabled and
every automation entry point are covered by the certification above and its
protection checks; repeat them by hand only when a check fails.

### 7. Build artifacts

From a clean checkout of the candidate, build:

- `K-SACCR_X.Y.Z.xlsm`: a **production** workbook, built as in
  [`INSTALLATION.md`](INSTALLATION.md#building-a-workbook-from-source),
  compiled, run once on the demonstration portfolio, cleared with
  **Clear outputs** and saved. Check it with the
  [source round trip](docs/EXCEL_EVIDENCE.md#source-round-trip) in production
  mode against the candidate SHA; only `RESULT=PASS` with
  `binding=VERIFIED` qualifies it;
- `K-SACCR_User_Manual_X.Y.Z.pdf`: built from `docs/user-manual/vX.Y.Z/` at
  the candidate, with the command in that folder's README;
- `SHA256SUMS.txt`: the SHA-256 of both files, in `sha256sum` format.

Check the **development** workbook certified in step 6 the same way: export
it and run the source round trip in development mode against the candidate
SHA. It is not published, but its round trip shows that the suites ran on
exactly the candidate's code, test modules included, among them
`TEST_Benchmark`, which the evidence policy does not run. The comparison
takes the expected components from `tools/vba-projects.json` at the
candidate. Only `RESULT=PASS` with `binding=VERIFIED` qualifies it.

Once attached to a published release, an asset is never replaced.

### 8. Create the annotated tag

Tag only the certified candidate, from the command line so the tag is
annotated:

```sh
git fetch origin
git tag -a vX.Y.Z FULL_SHA -F ../tag-message-vX.Y.Z.txt
git cat-file -t vX.Y.Z        # must print: tag
git push origin vX.Y.Z
```

The tag message is the release's permanent record:

```text
vX.Y.Z - <short release title>

Certified at <full 40-character SHA>.
Excel: compile <result>; suites <passed>/<suites>, failed=0; evidence validator PASS.
Fixes #<n>, #<n>.
Deviations: none.
```

`Deviations:` is mandatory: write `none`, or list each gate that was not run.
Copy totals verbatim from the run; never round or reuse earlier figures. Never
delete and recreate a published tag.

### 9. Publish and verify

Create the GitHub Release from the tag as a **draft**, attach the certified
artifacts and `SHA256SUMS.txt`, and only then publish it. Publishing starts
**Sign release assets** (`.github/workflows/release-signing.yml`). It
downloads the assets and refuses them unless `SHA256SUMS.txt` lists exactly
the other assets, every hash matches and no signature is attached yet. It
then signs each file with Sigstore keyless signing under the workflow's
GitHub identity, verifies each signature and attaches `<file>.sigstore.json`
for each: three bundles. Check that the run passed and that the release lists
the bundles. If it failed because an asset was attached after publication,
run the workflow by hand: **Actions → Sign release assets → Run workflow**,
from `main`, with the tag. Do not turn on GitHub's immutable releases: an
immutable release accepts no file after publication, so the signatures could
not be attached.

Verify that a fresh clone of the tag passes the static checks and imports
cleanly.

Then use the **published workbook as a user would**, on a copy downloaded
from the release page:

1. Check it against `SHA256SUMS.txt` and its signature, and unblock it, as
   [`INSTALLATION.md`](INSTALLATION.md#released-workbook) describes.
2. Open it and enable macros: it opens on Home and the K-SACCR tab is on
   the Ribbon. **Data → Edit Links** lists no external links, and
   **Formulas → Name Manager** shows no name with a `#REF!` reference.
3. **Validate inputs**, **Run SA-CCR** and **Clear outputs** from the K-SACCR
   Ribbon tab, then **Run SA-CCR** again. Save, close and reopen it: Results C5
   reads current.
4. **Export evidence**: the file is written. Unprotect Results (**Review →
   Unprotect Sheet**, a deliberate maintenance unlock on this copy), change an
   EAD cell and export again: the export is refused as `EDITED`. **Run
   SA-CCR**: Results is rebuilt and protected again.
5. **Export import specifications** on `Import_Trades`, save the page and open
   it. Copy its two examples into UTF-8 `.csv` files, **Import CSV** each on
   its staging sheet, **Validate staging** (both sheets read `VALID`),
   **Commit import** and **Run SA-CCR**: the run uses all 5 legs of the 3
   netting sets, with no error.
6. Close without saving, and record what was done and seen with the release
   evidence.

Once the release assets exist, update the website as
[`docs/WEBSITE.md`](docs/WEBSITE.md#maintenance-and-checks) describes: the
release named on the page, the download buttons and the `SHA256SUMS.txt` link.
The change goes through a task PR into the active release branch and then an
owner-authorized integration into `main`, from which Pages deploys `/docs`.
The website counts as updated only when that deployment has succeeded and the
public page names the new release and starts its downloads.

## 🧾 Evidence record

Retain at least: candidate SHA, previous tag, static-check result, Excel
version and bitness, the compile result and the validated evidence bundle for
every suite, the production and development round trips, the check of the
published workbook, artifact names, sizes and SHA-256 hashes, the signing run,
the integration PR,
validation timestamps and any deviations. A release note summarizes evidence;
it does not replace it.

## 🧯 Recovery

- **Before tagging:** fix through a task PR, update the evidence and rerun the
  affected gates.
- **After tagging, before publishing:** pause and document the correction;
  prefer a new patch version if the tag may have propagated.
- **After publishing:** do not replace assets or move the tag; the only
  files added are the signatures. Publish a
  corrected patch, record the problem in the changelog, and use
  [`SECURITY.md`](SECURITY.md) for vulnerabilities.

## ☑️ Maintainer checklist

| # | Gate | Done |
| ---: | --- | :---: |
| 1 | Owner requested the release | ☐ |
| 1a | Maintenance audit done and its report recorded; every finding fixed, moved out of the release or accepted | ☐ |
| 2 | Changelog and `VERSION` finalized | ☐ |
| 3 | Social preview checked | ☐ |
| 4 | Release branch merged into `main` with a merge commit, and level with `main` again | ☐ |
| 5 | Candidate frozen: the merge commit's SHA, diff reviewed | ☐ |
| 6 | Static checks pass locally and in CI on the candidate | ☐ |
| 7 | Excel certification passes for the harness and the nine policy suites on the candidate, and the manual protection checks pass | ☐ |
| 7a | [Ribbon walk-through](#ribbon-walk-through) done on the certified development workbook and recorded | ☐ |
| 8 | Artifacts built from the candidate, tested and hashed; production and development round trips `PASS` with `binding=VERIFIED` | ☐ |
| 9 | Annotated tag targets the candidate | ☐ |
| 10 | GitHub Release published, its assets signed by the signing workflow and verified, and its workbook used as a user would | ☐ |
| 11 | Website change integrated into `main`, Pages deployed, and the public page names and downloads the published release | ☐ |
