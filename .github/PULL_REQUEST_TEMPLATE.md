<!--
  Target: the active release branch (see docs/GOVERNANCE.md#branches). Never
  main, unless the owner asked for an integration.
  Write NOT RUN or NOT APPLICABLE with a reason; never manufacture PASS
  evidence. Static checks are never Excel evidence. No real portfolio data.
  Report vulnerabilities privately through SECURITY.md.
-->

## Summary

<!-- The observable outcome and why it is needed, in two or three sentences. -->

Closes #

**Type of change:** <!-- bug fix / feature / calculation change / tests / refactor / documentation / tooling or CI / release -->

## Calculation impact

<!-- "None" if no computed figure can change. Otherwise: -->

- Regime(s): <!-- CRR / BCBS / both -->
- Cited source: <!-- e.g. CRR Article 275(1), CRE99 Example 1 -->
- Class of expected values: <!-- published / independent / illustrative, per docs/methodology/TEST_CASES.md -->

## Validation

- `python tools/check.py`: <!-- PASS / FAIL -->
- Excel `TEST_Suite.Test_RunAll`: <!-- the RESULT= line verbatim and the full commit SHA it ran on, or NOT RUN and the reason -->
- Evidence record (`tools/check_excel_evidence.py`): <!-- PASS / FAIL with the record's candidate SHA, or N/A -->

## Documentation

- [ ] Affected documentation is updated.
- [ ] `CHANGELOG.md` has an entry under `[Unreleased]`, or the change is not user-facing.

## Risks and what remains unverified

<!-- Principal failure modes, environments not tested (32-bit/64-bit, locale),
checks skipped and why. -->

## Review checklist

- [ ] The PR targets the active release branch and links its issue, which has an assignee, exactly one priority label and a milestone.
- [ ] Scope is focused; the diff contains no unrelated formatting, local artifacts or generated output.
- [ ] Calculation, public API and compatibility impact are stated; a calculation change cites its source.
- [ ] VBA changes were compiled and tested in Excel on the head SHA, or on an earlier head when later updates changed nothing in `src/`, `tests/modules/` or the template (CONTRIBUTING step 7); otherwise they are stated as not yet verified, and the PR is not merged until they are.
- [ ] Expected values that count as validation are `published` or `independent`, never produced by running this code.
- [ ] Changed or added suites update `.github/excel-evidence-policy.json` and the README counts together.
- [ ] Excel state changed by a procedure is restored on success and failure.
- [ ] No credentials, real portfolio or client data, or vulnerability detail.
- [ ] **Repository integrity** is green and every review conversation, including the automated code review, is resolved.
