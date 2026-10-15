<div align="center">

# 🔒 K-SACCR Security Policy

### Counterparty credit risk under SA-CCR, built in Excel/VBA

[![Reporting](https://img.shields.io/badge/Reporting-Private-d97706?style=for-the-badge)](#reporting-a-vulnerability)
[![Support](https://img.shields.io/badge/Support-Pre--release-6e7781?style=for-the-badge)](#supported-versions)
[![Scope](https://img.shields.io/badge/Scope-Source_%7C_Automation-0969da?style=for-the-badge)](#security-scope)
[![Disclosure](https://img.shields.io/badge/Disclosure-Coordinated-6f42c1?style=for-the-badge)](#coordinated-disclosure)

<br>

**Protect users · Minimize exposure · Preserve evidence · Coordinate disclosure**

</div>

---

This document is authoritative for **vulnerability scope, private reporting,
security triage, coordinated disclosure and safe harbor**. Contribution workflow
is owned by [`CONTRIBUTING.md`](CONTRIBUTING.md); the release sequence by
[`RELEASING.md`](RELEASING.md); the current GitHub security controls are
recorded in [`docs/GOVERNANCE.md`](docs/GOVERNANCE.md#current-github-controls-and-limitations).

> [!IMPORTANT]
> A security policy does not make macros, workbooks, add-ins or release artifacts
> inherently trustworthy. Establish provenance and apply organizational security
> controls before enabling executable content.

## 🧭 Security model

The project assumes Microsoft Excel, the operating system and the VBA runtime
are trusted; the user is authorized to run the project; macros are enabled
through an approved trust mechanism; and source comes from this repository.

These are trust boundaries, not guarantees. VBA projects running in the same
Excel process are not isolated security sandboxes.

<a id="supported-versions"></a>

## 📦 Supported versions

K-SACCR is **pre-release**. No version has been released, so no version carries
production security support.

| Source state | Security support |
| --- | --- |
| `release/1.0.0` (active development) | ⚠️ Best effort |
| `main` | ⚠️ Best effort |
| Modified copies or unofficial mirrors | ❌ Unsupported unless reproduced in official source |

This table will be revised when the first version is released. Reports must
identify a full commit SHA; relative descriptions such as "latest" are
insufficient.

<a id="reporting-a-vulnerability"></a>

## 📣 Reporting a vulnerability

Do **not** disclose a suspected vulnerability in an issue, pull request, commit
message, sample workbook, screenshot or release note.

GitHub private vulnerability reporting is not enabled for this repository.
Report by email to **danielep71@gmail.com** with the subject
**Private security report — K-SACCR**.

Include only the information needed to assess the issue:

| Evidence | Requested detail |
| --- | --- |
| Identity | Full commit SHA, component or procedure, affected file or workflow |
| Environment | Excel/Office build, bitness, Windows version and deployment model |
| Impact | Confidentiality, integrity, availability, execution or supply-chain consequence |
| Reproduction | Minimal steps using synthetic data |
| Exploitability | Preconditions, privileges, user interaction and persistence |
| Mitigation | Tested workaround or containment, if known |
| Evidence | Sanitized diagnostics, hashes or proof of concept |

Never send real portfolios, trades, netting sets, counterparty or collateral
data, client or employer workbooks, or personal data. Remove credentials,
internal paths, links, connections, document metadata, cached values and
unrelated content.

If a secret has been exposed, revoke or rotate it immediately before improving
the report.

## ⏱️ Response process

K-SACCR is maintained by one person; targets are best-effort, not a contractual
SLA.

| Stage | Target |
| --- | --- |
| Acknowledge | Within 5 business days |
| Initial scope and severity assessment | Within 10 business days after sufficient evidence |
| Active-investigation update | At least every 14 days |
| Remediation and disclosure | Proportionate to severity, exploitability and validation needs |

The normal path is reproduce → scope affected commits → contain risk → fix →
add regression evidence → validate in Excel → publish a correction when
appropriate.

## 🎯 Security issue or ordinary defect?

When uncertain, report privately. Security-relevant reports include credible
risk of:

- unintended code execution or trust-boundary crossing;
- unauthorized reading, modification, deletion or disclosure of data;
- persistent or exploitable loss of availability;
- credential, token, runner or automation compromise;
- a provenance or validation bypass that can present unsafe output as trusted;
  or
- a correctness defect deliberately exploitable to defeat an integrity boundary.

An incorrect exposure figure, compatibility problem, bounded performance
regression or documentation error is normally an ordinary bug, reported through
an issue, unless it creates concrete security impact.

<a id="security-scope"></a>

## 🛠️ Security scope

### In scope

- source in this repository, including the exported VBA and the workbook
  template `src/workbook/SACCR_Template.xlsx`;
- repository-owned validation tooling in [`tools/`](tools/README.md);
- GitHub Actions workflows, their permissions and pinned actions; and
- security or integrity behavior introduced by project code.

### Current risk surfaces

- **Runtime.** The engine imported from the prototype workbook reads its own
  input sheets and writes its own output sheets, shows message boxes, and
  switches screen updating, events and calculation mode during a run. It uses
  no files, network, native code (`Declare`), `Shell` or `CreateObject`. The
  harness only reads Excel settings. The other test suites write the input and
  output sheets and switch calculation, events and screen updating, restoring
  them afterwards, so run them in a development workbook.
- **Automation.** Every workflow checkout that can run code under review sets
  `persist-credentials: false`, so that code never receives Git credentials.
  The one exception is the daily traffic export, whose job runs only on the
  default branch, on a schedule or manual dispatch, executes no repository code,
  and keeps credentials to push its data to the `traffic-history` branch. Its
  `TRAFFIC_TOKEN`, a fine-grained token with `Administration: read` on this
  repository only, lives in the `analytics` environment; rotate it every 90
  days and immediately if the workflow is changed by anyone but the owner.
  The `analytics` environment separately allows only the `main` branch and no
  tags (configured 2026-10-06). Keep this restriction before adding its token;
  update the allowlist if the default branch changes. The workflow guard alone
  is not a security boundary against an edited workflow on another branch.
  The static-check and
  pull-request label jobs run with a read-only token. The label-sync workflow
  grants `issues: write` only to its reconciliation job, which runs on pushes to
  `main` and manual dispatch, never on pull requests. All actions are pinned to full
  commit SHAs; the CI quality tools install from hash-locked requirements, and the
  actionlint binary is checksum-verified. Dependabot proposes action updates for
  manual review. See [`tools/README.md`](tools/README.md#github-actions) and
  [`docs/LABELS.md`](docs/LABELS.md).
- **Local tooling.** The VBA synchronizer `tools/Sync-SACCR-VBA.ps1` needs
  Excel's **Trust access to the VBA project object model** setting. The `.cmd`
  launchers run PowerShell with `-ExecutionPolicy Bypass`. With `-SourceDir`
  the synchronizer imports modules from a folder such as Downloads, so use it
  only with files obtained from this repository. See
  [`tools/README.md`](tools/README.md#local-vba-synchronization).
- **Artifacts.** No workbook, add-in or other binary is distributed. Office
  packages are ignored by Git unless an exact path is re-included.

### Out of scope

- vulnerabilities in Microsoft Excel, Office, Windows, GitHub, Python, Node.js or
  VBA themselves;
- organization-controlled endpoint, macro, access or deployment policy;
- malicious VBA already trusted in the same Excel process;
- unrelated workbooks, add-ins, dependencies or infrastructure;
- modified copies, mirrors or historical snapshots;
- compromised user credentials not exposed by this project; and
- ordinary defects without concrete security impact.

Upstream vulnerabilities belong with the responsible vendor or platform.

<a id="data-and-secrets"></a>

## 🔐 Data and secret handling

Never commit, upload, log or attach:

- passwords, API keys, personal access tokens, signing keys, certificates or
  connection strings;
- real trades, netting sets, counterparties, collateral agreements, portfolio
  extracts or any client, employer or personal data;
- proprietary source, models, workbooks, production extracts or licensed data;
- internal URLs, machine-specific paths, environment dumps or unredacted
  screenshots; or
- exploit material beyond what is necessary to establish the issue.

Test fixtures and examples are synthetic. Excel files can contain sensitive
material outside visible cells, including document properties, names, hidden
sheets, VBA, cached values, queries, links and connections.

This document does not rely on GitHub secret scanning or code scanning; these
rules are enforced by review.
Repository secrets must use least privilege, remain unavailable to untrusted
pull-request code and be rotated after suspected exposure.

## 📦 Supply-chain boundary

Trusted source is limited to this repository. The release sequence is defined in
[`RELEASING.md`](RELEASING.md) and is not duplicated here.

Security-sensitive workflow changes require least-privilege permissions,
full-SHA action pins and explicit review. Do not run untrusted code on a
persistent credentialed Excel/Windows machine. Treat logs, screenshots,
workbooks, test artifacts and environment metadata as potentially sensitive.

## ✅ Safe-use guidance

Users should:

- preserve organization-approved macro security and deployment controls;
- obtain source only from this repository and know which commit it is;
- test with synthetic data in a controlled environment before using real data;
  and
- understand that K-SACCR output is a calculation aid. It is not regulatory
  approval, model validation or an authentication or authorization control.

<a id="coordinated-disclosure"></a>

## 📣 Coordinated disclosure

Avoid wider disclosure while exploitability is being assessed, a fix is being
prepared, users have not had reasonable time to update, or an exposed secret
remains valid.

The maintainer and reporter agree a plan based on severity, active
exploitation, remediation complexity, workarounds and validation time. The
maintainer may request a sanitized reproduction, more environment detail,
confirmation against a candidate fix or a reasonable embargo.

<a id="safe-harbor"></a>

## 🛡️ Good-faith research and safe harbor

Good-faith research is welcome when it:

- stays within project-owned source, artifacts and documented integrations;
- avoids privacy violations, destructive actions, persistence, social
  engineering and unnecessary access;
- stops after establishing the minimum required evidence;
- reports privately and promptly; and
- allows reasonable investigation and remediation time.

The project will not initiate or recommend legal action solely for research
conducted in good faith and consistently with this policy. This does not
authorize testing third-party systems or bind Microsoft, GitHub, an employer,
client or other third party.

No paid bug bounty is offered.

## 📚 Related documents

- [`CONTRIBUTING.md`](CONTRIBUTING.md) — contribution and review workflow
- [`CODE_OF_CONDUCT.md`](CODE_OF_CONDUCT.md) — participant behavior
- [`INSTALLATION.md`](INSTALLATION.md) — developer setup and safe import
- [`RELEASING.md`](RELEASING.md) — release sequence
- [`docs/GOVERNANCE.md`](docs/GOVERNANCE.md) — GitHub controls and review process

Conduct complaints and vulnerability reports are different: use the Code of
Conduct for participant behavior and this policy for software and security risk.

---

<div align="center">

### Security principle

**Trust deliberately · Run minimally · Protect secrets · Preserve evidence · Disclose responsibly**

<br>

Maintained by **Daniele Penza**

</div>
