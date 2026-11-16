# 🧾 Excel runs

Dated records of Excel runs that are not release certifications: the quick
run of `tools/Verify-SACCR-Branch.ps1 -Quick` before a pull request with VBA
or template changes is merged, when its result is worth keeping. A record
names the exact commit, the command, the host, each step's result and what
its pictures showed when looked at, and links the pictures kept in
[`docs/assets/screenshots/`](../assets/screenshots/). It is not validated by
`tools/check_excel_evidence.py` and is never release evidence: a release is
certified as [`RELEASING.md`](../../RELEASING.md) describes.

A record is written once, after the run, and not edited afterwards; a later
run gets its own record. The name is `YYYY-MM-DD-quick-SHORTSHA.md`.

| Date | Commit | Pull request | Result | Record |
| --- | --- | --- | --- | --- |
| 2026-11-15 | `369abf4` | #707 (#706, #708) | PASS, with the sheet review | [2026-11-15-quick-369abf4.md](2026-11-15-quick-369abf4.md) |
