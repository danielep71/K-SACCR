# 📚 Regulatory Sources

> **Guide, not policy:** [Methodology](../methodology/README.md#source-register) owns exact source versions, locators, and verification records.

## Source roles

EU implementation decisions follow the applicable CRR and delegated acts. EBA Q&A is interpretative material, not a substitute for those legal texts. Basel is a separate supported regime. A regulatory citation establishes a rule's source; it does not independently validate a numerical implementation.

## Registered baseline

The following is a summary of the repository's source register, not a fresh legal verification performed for this wiki review.

| ID | Registered source | Recorded version or role |
| --- | --- | --- |
| CRR | Regulation (EU) No 575/2013, as amended | Consolidated June 26, 2026, version 021.001 |
| CRR-RTS | Delegated Regulation (EU) 2021/931, amended by 2025/855 | Consolidated May 25, 2025, version 001.001 |
| EBA-QA-2023_6962 | Final EBA Q&A on collateral in the margined exposure cap | Published March 15, 2024 |
| BCBS | Basel Framework CRE52 | Effective January 1, 2023; registered update June 5, 2020 |
| BCBS-279 | SA-CCR standard and Annex 4a examples | March 2014, revised April 2014 |
| CRE99 | Basel Framework application guidance | Published numerical examples used by the fixture suite |

The register records checks on October 9–10, 2026, and identifies their scope. In particular, the RTS mapping articles are not recorded there as fully rechecked. Do not change that status to “verified” merely because a wiki page cites them.

Source copies, addresses, and hashes are in the [source archive](../methodology/sources/README.md). Use that register to locate the relevant official version.

## Evidence labels

- **Published:** an expected value printed in a registered source, with its locator, units, and rounding tolerance.
- **Illustrative:** an engineering example, unreviewed hand derivation, or value from an unreviewed comparator.
- **Independent:** reserved for an independently derived and reviewed value; no such values are claimed for v1.0.0.

A rejected Q&A is not regulatory authority. The old wiki's posted-VM scenario and “independent Python” badges must not be used as certification evidence.

See [Regulatory Basis](Regulatory-Basis.md), [Basel vs CRR](Basel-vs-CRR.md), and [Validation Approach](Validation-Approach.md).
