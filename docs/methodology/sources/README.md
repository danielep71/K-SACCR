# 📚 Regulatory sources

Copies of the official texts K-SACCR's methodology was checked against, so that
anyone can verify a citation without searching for the right version. They are
the public versions published by the EU and the Basel Committee; the
[source register](../README.md#source-register) says which rules rely on which
text.

These are documentation copies. The authentic EU texts are those published in
the Official Journal of the European Union; the authoritative Basel text is the
one on bis.org. A newer version is added as a new file next to the old one, with
its own row below; a file is never replaced.

The copies are not covered by the repository's licence; each stays under its
publisher's terms. The BIS terms and conditions
(<https://www.bis.org/terms_conditions.htm>) allow BIS material, including
the Basel Committee's chapters and BCBS 279, to be downloaded and
redistributed for non-commercial purposes without written permission.

Do not edit or reformat these copies, including `EBA_QA_2023_6962.md`: their
hashes are registered below, and a changed byte breaks the record. To check
them, run this from the repository root (it works on Windows too) and compare
with the table:

```sh
python -c "import glob,hashlib;[print(hashlib.sha256(open(f,'rb').read()).hexdigest(),f) for f in sorted(glob.glob('docs/methodology/sources/*')) if not f.endswith('README.md')]"
```

bis.org builds its chapter PDFs when they are downloaded, with the download
time as the creation date, so a new download of a Basel chapter does not give
the registered hash.

| File | Source ID | Document | Version | Retrieved from | Retrieved | SHA-256 |
| --- | --- | --- | --- | --- | --- | --- |
| `CRR_575-2013_consolidated_2026-06-26_extract.pdf` | `CRR` | Regulation (EU) No 575/2013, consolidated text: page 1, pages 379–417 (Articles 271–282 and the start of Article 283) and pages 421–423 (Article 285 and the start of Article 286) | 26.06.2026, version 021.001 | <https://eur-lex.europa.eu/legal-content/EN/TXT/PDF/?uri=CELEX:02013R0575-20260626> | 2026-10-10 | `11471b9b93be3e2c7b485bfac84d220dfdb65fdc0f8ebb363446c5bc91dae78d` |
| `CRR-RTS_2021-931_consolidated_2025-05-25.pdf` | `CRR-RTS` | Commission Delegated Regulation (EU) 2021/931, consolidated text including Delegated Regulation (EU) 2025/855 | 25.05.2025, version 001.001 | <https://eur-lex.europa.eu/legal-content/EN/TXT/PDF/?uri=CELEX:02021R0931-20250525> | 2026-10-10 | `f6a06f62b15f2bc261535cc6b269130955a2d0564b93b78e50eb81e674462f99` |
| `BCBS_CRE52_effective_2023-01-01.pdf` | `BCBS` | Basel Framework, CRE52 Standardised approach to counterparty credit risk, with FAQs | Effective 1 January 2023, last updated 5 June 2020 | <https://www.bis.org/committees/bcbs/basel-framework/81412/chapter.pdf> | 2026-10-10 | `6ce15b86f1c67692e16806ae76df6bd97c8f19820a72c7ceb7ddac78b509c111` |
| `BCBS_CRE99_effective_2023-01-01.pdf` | `CRE99` | Basel Framework, CRE99 Application guidance, with the SA-CCR worked examples | Effective 1 January 2023, published 27 March 2020 | <https://www.bis.org/committees/bcbs/basel-framework/81410/chapter.pdf> | 2026-10-10 | `f0d86acfe451699d7c3d783822e2372b180798b3e4fdd7b4018524ba8a937855` |
| `BCBS_CRE54_effective_2023-01-01.pdf` | `CRE54` | Basel Framework, CRE54 Capital requirements for bank exposures to central counterparties | Effective 1 January 2023, published 27 March 2020 | <https://www.bis.org/committees/bcbs/basel-framework/81405/chapter.pdf> | 2026-10-16 | `44d8c361e8d5d827d8ebcde630024f3b6ed0de6efdda09cb5f43e010e85895d4` |
| `CRR_575-2013_consolidated_2026-06-26_extract_CCP.pdf` | `CRR-CCP` | Regulation (EU) No 575/2013, consolidated text: page 1, page 22 (Article 4(1)(83)–(92)) and pages 441–449 (Section 9, Articles 300–311) | 26.06.2026, version 021.001 | <https://eur-lex.europa.eu/legal-content/EN/TXT/PDF/?uri=CELEX:02013R0575-20260626> | 2026-10-18 | `73acdbba784b73e9a6c47f48008523340a5818124f3b7b6f1416de784191e047` |
| `BCBS_279_revised_2014-04.pdf` | `BCBS-279` | BCBS, *The standardised approach for measuring counterparty credit risk exposures*, 37 pages, with Annex 4a (sample portfolios) | March 2014, revised April 2014 | <https://www.bis.org/publ/bcbs279.htm> | 2026-10-18 | `5f17fea239b28bc4cf152b11d3c44284b8f2dcc5e6c034a9e2f38dad1ede0b0f` |
| `EBA_QA_2023_6962.md` | `EBA-QA-2023_6962` | EBA Single Rulebook Q&A 2023_6962, exposure value cap for netting sets subject to a margin agreement | Final, published 15/03/2024 | <https://www.eba.europa.eu/single-rule-book-qa/qna/view/publicId/2023_6962> | 2026-10-10 | `9eea2133a5ae404a80ef8c234b764d5c64db94b22059c0810fa618babfc37abb` |

Both CRR extracts were cut from the full consolidated PDF (SHA-256
`6796302ed4213bcba4d4b9b17dc48aefcaf7ef586b875c30c8c71a7689288ecf`, 895 pages)
without changing the pages; the second, retrieved again on 2026-10-18, had the
same hash. Of the articles in the first extract, the methodology implements the
full SA-CCR of Articles 274–280f and relies on the definitions (Article 272) and
the margin period of risk (Article 285(2)–(5)). Articles 273a (conditions for
the simplified methods), 281 (simplified SA-CCR) and 282 (original exposure
method) are in the extract as context only; K-SACCR does not implement them.
The second extract holds the definitions of Article 4(1)(88)–(91) and Section 9
on exposures to central counterparties, which the
[clearing coverage](../CLEARING_COVERAGE.md) relies on; K-SACCR calculates
nothing from them yet. The consolidated text itself has no legal effect, as its
first page states.

BCBS 279 (2014) is the original SA-CCR standard, now part of the consolidated
Basel Framework; its Annex 4a prints the same worked examples as CRE99. The
copy was supplied by the owner from the BIS publication page; a fresh download
for comparing hashes was not possible here because bis.org answered with a
browser check. Every published value used by the test cases is still cited
from CRE99.
