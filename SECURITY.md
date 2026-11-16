# K-SACCR Security Policy

## Supported versions

| Version | Supported |
| --- | --- |
| 1.2.0 (tag `v1.2.0`) | Yes |
| 1.1.0 (tag `v1.1.0`) | No: superseded by 1.2.0 |
| 1.0.2 (tag `v1.0.2`) | No: superseded by 1.1.0 |
| 1.0.1 (tag `v1.0.1`) | No: superseded by 1.0.2 |
| 1.0.0 (tag `v1.0.0`) | No: superseded by 1.0.1, which changes no calculation |
| Active release branch and `main` | Best effort |

The active release branch is named in
[`docs/GOVERNANCE.md`](docs/GOVERNANCE.md#branches). A fix to a released version
is published as a new patch release; a published release is never changed, apart
from the Sigstore signatures of its files, which the release workflow attaches.
Modified copies and unofficial mirrors are unsupported
unless the problem is reproduced in this repository. Reports must identify a
full commit SHA.

The [security assurance case](docs/ASSURANCE_CASE.md) states the security
requirements, threats and trust boundaries, and argues how the design and the
code meet them.

Every file of a release has a Sigstore signature, `<file>.sigstore.json`,
made by the repository's release workflow;
[`INSTALLATION.md`](INSTALLATION.md#released-workbook) shows how to verify it.

## Reporting a vulnerability

Do **not** disclose a suspected vulnerability in an issue, pull request, commit
message, sample workbook or screenshot.

Report it privately through GitHub: the repository's **Security** tab →
**Report a vulnerability**. If that is unavailable, email
**danielep71@gmail.com** with the subject **Private security report — K-SACCR**.

Include:

- the full commit SHA and the affected file, procedure or workflow;
- the Excel/Office build, bitness and Windows version, where relevant;
- the impact and minimal reproduction steps, with no real portfolio data; and
- any tested workaround.

## Response targets

K-SACCR is maintained by one person; targets are best effort, not a contractual
SLA: acknowledgement within 5 business days, an initial assessment within 10
business days after sufficient evidence, and an update at least every 14 days
while investigating. Remediation and disclosure timing are agreed with the
reporter according to severity.

## Scope

**In scope:** the exported VBA and the workbook template in `src/`, the test
modules in `tests/`, the tooling in [`tools/`](tools/README.md), and the GitHub
Actions workflows.

**Out of scope:** vulnerabilities in Excel, Office, Windows, GitHub, Python or
Node.js themselves; organization-controlled macro and deployment policy;
malicious VBA already trusted in the same Excel process; modified copies; and
ordinary defects without concrete security impact, such as an incorrect
exposure figure, which are reported as issues.

## Runtime risk surface

- **Workbook macros** read the input sheets, write the output sheets, show
  message boxes, and switch screen updating, events and calculation mode during
  a run, restoring them afterwards.
- **CSV check** (`M_Main.CheckTradesCsv` and `CheckNettingSetsCsv`) reads the
  file chosen in `Application.GetOpenFilename` up to the contract's 128 MiB
  limit and writes nothing; the size is checked before the file is read.
- **Import CSV** (`M_Main.ImportTradesCsv` and `ImportNettingSetsCsv`) reads
  the file chosen in `Application.GetOpenFilename` with the same reader and
  limit and writes its records as text to the staging sheet `Import_Trades`
  or `Import_NettingSets`; when that sheet has changes since its last load, it
  asks before replacing them. It writes no file.
- **Commit import** (`M_Main.CommitImport`) overwrites the input sheets
  `NettingSets` and `Trades` with the staged rows of a current `VALID`
  validation, only after a confirmation (OK or Cancel, Cancel by default)
  naming each sheet and the rows that replace its rows.
- **Export evidence** (`M_Main.ExportEvidence`) and **Export CSV template**
  (`M_Main.ExportTradesCsvTemplate` and `ExportNettingSetsCsvTemplate`) ask for
  a file name with `Application.GetSaveAsFilename` and write the evidence CSV
  or the blank template there. For these two exports an existing file of that
  name, hidden ones included, is deleted and replaced without a further prompt
  from the macro; a read-only file is left unchanged and the export fails.
- **Export import specifications** (`M_Main.ExportImportSpecification`) asks
  for a file name with `Application.GetSaveAsFilename`, asks before replacing
  an existing file of that name, and writes the import specification page
  there. When the user agrees, it opens the page with
  `ThisWorkbook.FollowHyperlink`, in the application Windows associates with
  `.html` files.
- **Test suites** and the benchmark `TEST_Benchmark` also write the input and
  output sheets. In `%TEMP%` (or the workbook's folder) `TEST_Reconcile` and
  `TEST_CsvReader` write and delete five CSV files and one,
  `TEST_Import` writes and deletes `K-SACCR_import_test.csv` and
  `K-SACCR_import_spec_test.html`, and `TEST_Benchmark.RunCsvBenchmark` writes
  each of its files as `K-SACCR_csv_benchmark_<n>.csv`, under a number not
  already taken, and deletes it once read. Run them only in a development
  workbook.
  `src/` and `tests/` contain no `Shell`, `CreateObject` or network access,
  and one `Declare`, described next.
- **Windows API.** `M_Ribbon` declares one function of Windows,
  `RtlMoveMemory` of `kernel32`, with `PtrSafe` and `LongPtr`. It copies the
  Ribbon object's pointer into an object variable to recover the K-SACCR tab
  after a VBA state reset (#216). The pointer is kept in the hidden workbook
  name `ksRibbonPointer` with the window handle of the Excel session, used only
  in that session, and removed before every save, so no file carries it.
  Nothing is installed, registered or loaded beyond `kernel32`, which every
  Windows process has.
- **Local tooling.** `tools/Sync-SACCR-VBA.ps1`, `tools/Export-SACCR-VBA.ps1`
  and the scripts that build workbooks through the first
  (`tools/Certify-SACCR.ps1`, `tools/Check-SACCR-Protection.ps1` and
  `tools/Verify-SACCR-Branch.ps1`) need Excel's **Trust access to the VBA
  project object model** setting, which none of them changes. `Certify-SACCR.ps1` and
  `Check-SACCR-Protection.ps1`, which it runs, open their workbooks following
  the Trust Center macro setting and ask the operator to enable content; the
  protection checks also open the production workbook with its macros
  disabled for this Excel session only. `Verify-SACCR-Branch.ps1 -Quick`, the
  pull-request check, and `Check-SACCR-Protection.ps1 -Unattended`, which it
  runs, instead enable the macros of the workbooks they build for their own
  Excel session only (`AutomationSecurity` Low), without touching the Trust
  Center. The
  `.cmd` launchers run PowerShell with `-ExecutionPolicy Bypass`, and
  `Sync-SACCR-VBA-Downloads.cmd` imports modules from a folder such as
  Downloads: use it only with files obtained from this repository.
- **Automation.** Workflow permissions and credentials are described in
  [`tools/README.md`](tools/README.md#github-actions). The
  [traffic export](tools/README.md#traffic-history) uses a dedicated GitHub App
  with only Administration: Read-only. Its client ID and private key are held
  in the `analytics` environment; only the token step receives the key and
  only the fetch step receives the short-lived installation token. Follow the
  linked setup and rotation procedure. Rotate the private key every 90 days,
  immediately on suspected exposure, and if anyone but the owner changes the
  workflow. Never disclose the key in an issue, PR, log or chat.

## Data and secrets

Test fixtures and examples are synthetic. Never commit, attach or send real
trades, portfolios, counterparty or client data, credentials, tokens or
internal paths. Excel files can hold sensitive material outside visible cells.
Revoke or rotate an exposed secret at once, before reporting.

## Safe harbor

The project will not initiate or recommend legal action for good-faith
research that follows this policy, stays within project-owned source, avoids
privacy violations and destructive actions, and is reported privately. This
does not authorize testing third-party systems; no bug bounty is offered.
