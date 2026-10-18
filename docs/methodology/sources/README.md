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
| `EBA_QA_2023_6962.md` | `EBA-QA-2023_6962` | EBA Single Rulebook Q&A 2023_6962, exposure value cap for netting sets subject to a margin agreement | Final, published 15/03/2024 | <https://www.eba.europa.eu/single-rule-book-qa/qna/view/publicId/2023_6962> | 2026-10-10 | `9eea2133a5ae404a80ef8c234b764d5c64db94b22059c0810fa618babfc37abb` |

The CRR extract was cut from the full consolidated PDF (SHA-256
`6796302ed4213bcba4d4b9b17dc48aefcaf7ef586b875c30c8c71a7689288ecf`, 895 pages)
without changing the pages. Of these articles, the methodology implements the
full SA-CCR of Articles 274–280f and relies on the definitions (Article 272) and
the margin period of risk (Article 285(2)–(5)). Articles 273a (conditions for
the simplified methods), 281 (simplified SA-CCR) and 282 (original exposure
method) are in the extract as context only; K-SACCR does not implement them.
The consolidated text itself has no legal effect, as its first page states.

BCBS 279 (2014), whose Annex 4a prints the same worked examples as CRE99, is
cited but not copied. Its 37-page PDF (March 2014, revised April 2014) is
available from the BIS publication page
<https://www.bis.org/publ/bcbs279.htm>, but its notice reserves all rights and
allows only brief excerpts, so it is not reproduced here; every published value
used by the test cases is cited from CRE99.
