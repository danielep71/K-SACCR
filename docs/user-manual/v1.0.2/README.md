# K-SACCR user manual, version 1.0.2

The source of the user manual for K-SACCR 1.0.2. The PDF built from it is
attached to the 1.0.2 GitHub Release when that version is released; it is not
stored in this repository.

| File | Contents |
| --- | --- |
| [`K-SACCR_v1.0.2_User_Manual.md`](K-SACCR_v1.0.2_User_Manual.md) | The manual, in pandoc Markdown |
| `assets/data_flow.png`, `assets/evidence_gate.png`, `assets/rc_pfe.png` | The three diagrams it includes; `rc_pfe.png` is a copy of `docs/assets/SACCR_RC_PFE.png` |
| `header.tex` | The LaTeX header the PDF build needs (`\Needspace`, the link colour `KBlue`) |

The Markdown is written for pandoc and LaTeX: page-break commands
(`\Needspace`, `\newpage`, `\clearpage`) and display formulas in `\[ … \]`
appear as plain text when the file is viewed on GitHub. Read the PDF.

## Status

Prepared for 1.0.2, not yet released: the release date, the delivered file
names and the test counts are confirmed when 1.0.2 is released. Started from
the 1.0.1 manual and updated for the changes of 1.0.2: trade exposures to a
CCP and the two NettingSets columns they need, settled-to-market trades and
the two Trades columns before Comment, the role `MULTI`, the INFO lines on
Checks for client clearing and CCPs, the highlighted Results rows, the
migration from 1.0.1, and the worksheet-function list of section 10.2, whose
signatures no longer overflow the page (#379).

As before, the manual is written for someone who receives only the workbook
`K-SACCR_1.0.2.xlsm` and `SHA256SUMS.txt`: it does not mention the source
repository, building the workbook from source or the release process, and of
the testing it gives only the summary result. The 1.0.0 and 1.0.1 manuals stay
as released in [`../v1.0.0/`](../v1.0.0/README.md) and
[`../v1.0.1/`](../v1.0.1/README.md).

## Building the PDF

Requires pandoc 3 and a XeLaTeX distribution with the `needspace` package
(for example TeX Live with `texlive-xetex` and `texlive-latex-extra`) and the
DejaVu fonts. From this folder:

```sh
pandoc K-SACCR_v1.0.2_User_Manual.md \
  --from markdown+tex_math_single_backslash \
  --pdf-engine=xelatex --toc --number-sections \
  -H header.tex -V geometry:margin=2cm -V fontsize=10pt \
  -V mainfont="DejaVu Sans" -V monofont="DejaVu Sans Mono" \
  --resource-path=. -o K-SACCR_User_Manual_1.0.2.pdf
```

Edit the Markdown, not the PDF. A later version of the manual goes in its own
folder (`docs/user-manual/v1.0.3/` and so on); this one stays as released.
