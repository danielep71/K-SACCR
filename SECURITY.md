# K-SACCR Security Policy

## Supported versions

| Version | Supported |
| --- | --- |
| 1.0.2 (tag `v1.0.2`) | Yes |
| 1.0.1 (tag `v1.0.1`) | No: superseded by 1.0.2 |
| 1.0.0 (tag `v1.0.0`) | No: superseded by 1.0.1, which changes no calculation |
| Active release branch and `main` | Best effort |

The active release branch is named in
[`docs/GOVERNANCE.md`](docs/GOVERNANCE.md#branches). A fix to a released version
is published as a new patch release; a published release is never changed. Modified copies and unofficial mirrors are unsupported
unless the problem is reproduced in this repository. Reports must identify a
full commit SHA.

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
- **Export evidence** (`M_Main.ExportEvidence`) asks for a file name with
  `Application.GetSaveAsFilename` and writes the evidence CSV there. An existing
  file of that name, hidden ones included, is deleted and replaced without a
  further prompt from the macro; a read-only file is left unchanged and the
  export fails.
- **Test suites** and the benchmark `TEST_Benchmark` also write the input and
  output sheets, and `TEST_Reconcile` writes and deletes three CSV files in
  `%TEMP%` (or the workbook's folder); run them only in a development workbook.
  `src/` and `tests/` contain no `Declare`, `Shell`, `CreateObject` or network
  access.
- **Local tooling.** `tools/Sync-SACCR-VBA.ps1` and `tools/Export-SACCR-VBA.ps1`
  need Excel's **Trust access to the VBA project object model** setting. The
  `.cmd` launchers run PowerShell with `-ExecutionPolicy Bypass`, and
  `Sync-SACCR-VBA-Downloads.cmd` imports modules from a folder such as
  Downloads: use it only with files obtained from this repository.
- **Automation.** Workflow permissions and credentials are described in
  [`tools/README.md`](tools/README.md#github-actions). The `TRAFFIC_TOKEN` of the
  [traffic export](tools/README.md#traffic-history) is rotated every 90 days,
  and immediately if that workflow is changed by anyone but the owner.

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
