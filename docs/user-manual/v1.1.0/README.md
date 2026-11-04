# K-SACCR user manual, version 1.1.0

The source of the user manual for K-SACCR 1.1.0. The PDF built from it is
attached to the
[1.1.0 GitHub Release](https://github.com/danielep71/K-SACCR/releases/tag/v1.1.0);
it is not stored in this repository.

| File | Contents |
| --- | --- |
| [`K-SACCR_v1.1.0_User_Manual.md`](K-SACCR_v1.1.0_User_Manual.md) | The manual, in pandoc Markdown |
| `assets/data_flow.png`, `assets/evidence_gate.png`, `assets/rc_pfe.png` | The three diagrams it includes; `rc_pfe.png` is a copy of `docs/assets/SACCR_RC_PFE.png` |
| `header.tex` | The LaTeX header the PDF build needs (`\Needspace`, the link colour `KBlue`) |

The Markdown is written for pandoc and LaTeX: page-break commands
(`\Needspace`, `\newpage`, `\clearpage`) and display formulas in `\[ … \]`
appear as plain text when the file is viewed on GitHub. Read the PDF.

## Status

Released with 1.1.0 (tag `v1.1.0`, 5 November 2026). Started from the 1.0.2 manual and updated for the changes of 1.1.0:
the CSV import from end to end (the blank templates, the import
specification page, the staging sheets `Import_NettingSets` and
`Import_Trades` and their statuses, correcting staged values, Validate
staging and its findings on Checks, Commit import and the run after it, the
file format, limits and finding codes); the workbook protection (what each
sheet lets a user change, what is refused, and the maintenance unlock and
relock); the Origin column of Checks; the map of the 18 worksheets, the
buttons, the user journey, moving inputs from 1.0.2, the known limitations,
the checklists, the troubleshooting table, the test summary, the glossary
and the quick-reference card.

As before, the manual is written for someone who receives only the workbook
`K-SACCR_1.1.0.xlsm` and `SHA256SUMS.txt`: it does not mention the source
repository, building the workbook from source or the release process, and of
the testing it gives only the summary result. The release's PDF is built from
this folder at the tag, and its SHA-256 is in the release's `SHA256SUMS.txt`.
The 1.0.0, 1.0.1 and 1.0.2 manuals stay as released in
[`../v1.0.0/`](../v1.0.0/README.md), [`../v1.0.1/`](../v1.0.1/README.md) and
[`../v1.0.2/`](../v1.0.2/README.md).

## Building the PDF

Requires pandoc 3 and a XeLaTeX distribution with the `needspace` package
(for example TeX Live with `texlive-xetex` and `texlive-latex-extra`) and the
DejaVu fonts. From this folder:

```sh
pandoc K-SACCR_v1.1.0_User_Manual.md \
  --from markdown+tex_math_single_backslash \
  --pdf-engine=xelatex --toc --number-sections \
  -H header.tex -V geometry:margin=2cm -V fontsize=10pt \
  -V mainfont="DejaVu Sans" -V monofont="DejaVu Sans Mono" \
  --resource-path=. -o K-SACCR_User_Manual_1.1.0.pdf
```

Edit the Markdown, not the PDF. A later version of the manual goes in its own
folder (`docs/user-manual/v1.1.1/` and so on); this one stays as released.
