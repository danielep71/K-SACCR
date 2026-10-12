# 🚧 Open Items and Limitations

> **Guide, not policy:** [Methodology](../methodology/README.md) owns supported scope; linked GitHub issues own current delivery status.

## Snapshot: October 12, 2026

The reviewed baseline is #166, commit `21978de6e783df697611be09290dc135d8edd065`. This page is a dated guide; issue state can change after publication.

### Current release work

| Issue | Remaining work |
| --- | --- |
| [#165](https://github.com/danielep71/VBA-SACCR-Toolkit/issues/165) | Randomised inputs for the formula tests, by button and in the harness |
| [#161](https://github.com/danielep71/VBA-SACCR-Toolkit/issues/161) | Align numerical fixture validation with whole-business-day input rules |
| [#160](https://github.com/danielep71/VBA-SACCR-Toolkit/issues/160) | Require an explicit VB_Name before accepting download-folder modules |
| [#159](https://github.com/danielep71/VBA-SACCR-Toolkit/issues/159) | Preserve cp1252 text when replacing workbook document modules |
| [#152](https://github.com/danielep71/VBA-SACCR-Toolkit/issues/152) | Track named Params ranges in the input fingerprint and stale-results dependencies |
| [#145](https://github.com/danielep71/VBA-SACCR-Toolkit/issues/145) | Detect full-name EBA Q&A prose citations while allowing source-register definitions |
| [#144](https://github.com/danielep71/VBA-SACCR-Toolkit/issues/144) | Ignore block-commented source-root declarations before selecting the active list |
| [#120](https://github.com/danielep71/VBA-SACCR-Toolkit/issues/120) | Define and validate digital-option collar sizing and cap before claiming support |
| [#60](https://github.com/danielep71/VBA-SACCR-Toolkit/issues/60) | Refactor the imported engine to the repository's architecture and house style |
| [#46](https://github.com/danielep71/VBA-SACCR-Toolkit/issues/46) | Certify and package the first usable Excel release |
| [#45](https://github.com/danielep71/VBA-SACCR-Toolkit/issues/45) | Add reproducible run evidence and bank-result reconciliation |
| [#42](https://github.com/danielep71/VBA-SACCR-Toolkit/issues/42) | Define safe public UDF domains and consistent error contracts |
| [#40](https://github.com/danielep71/VBA-SACCR-Toolkit/issues/40) | Specify and test date, FX-notional and option conventions |
| [#34](https://github.com/danielep71/VBA-SACCR-Toolkit/issues/34) | Integrate the prototype as reviewable VBA source and a reproducible workbook |

### Delivered baselines

Economic-trade grouping (#38/#156), automatic calibration classification (#41/#154), the FormulaTest layout (#162–#167), and reporting-currency format validation with remargining boundary tests (#147/#168, after the reviewed baseline) are delivered. The climatic factor is no longer labeled an unconfirmed placeholder. Remaining follow-up issues are not waived by closing these baselines.

## Operational and product limits

- Windows 64-bit Excel is the supported target; Mac and web Excel are unsupported.
- CRR numerical validation and independent second-source evidence are not claimed for v1.0.0.
- User mapping, input completeness, collateral eligibility/haircuts, supplied FX rates, and the applicable legal netting arrangements remain external responsibilities.
- Multiple/shared margin agreements are unsupported under the declared input contract.
- Generic option-leg support does not complete digital-option sizing/cap validation.
- Optional sold-option and sold-protection reductions are not automatically applied.
- Module synchronization does not migrate workbook layouts. SourceDir identity/encoding and Params fingerprint issues remain open.
- Invalid or incomplete netting sets have withheld results; do not use an apparently successful partial total.

## Later milestones

| Milestone | Planned work |
| --- | --- |
| v1.1.0 | Independent comparison (#100), diagnostic codes (#81), optional FX quote/freeze workflow (#157) |
| v1.2.0 | Simplified SA-CCR (#107) |
| v1.3.0 | Original Exposure Method (#111) |
| v1.4.0 | CCP exposures and default-fund treatment (#112) |
| v1.5.0 | SFT exposures (#113) |
| v1.6.0 | CVA (#114) |
| v1.7.0 | RWA (#115) |

Milestone assignment is planning, not implementation, release publication, or authorization for a new branch.
