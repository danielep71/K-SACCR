# K-SACCR user manual, version 1.0.1

The source of the user manual for K-SACCR 1.0.1. The PDF built from it is
attached to the
[1.0.1 GitHub Release](https://github.com/danielep71/VBA-SACCR-Toolkit/releases/tag/v1.0.1);
it is not stored in this repository.

| File | Contents |
| --- | --- |
| [`K-SACCR_v1.0.1_User_Manual.md`](K-SACCR_v1.0.1_User_Manual.md) | The manual, in pandoc Markdown |
| `assets/data_flow.png`, `assets/evidence_gate.png`, `assets/rc_pfe.png` | The three diagrams it includes; `rc_pfe.png` is a copy of `docs/assets/SACCR_RC_PFE.png` |
| `header.tex` | The LaTeX header the PDF build needs (`\Needspace`, the link colour `KBlue`) |

The Markdown is written for pandoc and LaTeX: page-break commands
(`\Needspace`, `\newpage`, `\clearpage`) and display formulas in `\[ … \]`
appear as plain text when the file is viewed on GitHub. Read the PDF.

## Status

Released with 1.0.1 (tag `v1.0.1`, 18 October 2026). The manual's
repository links point to `blob/v1.0.1/`. The release's PDF is built from this
folder at the tag, and its SHA-256 is in the release's `SHA256SUMS.txt`.
The 1.0.0 manual stays as released in [`../v1.0.0/`](../v1.0.0/README.md).

## Building the PDF

Requires pandoc 3 and a XeLaTeX distribution with the `needspace` package
(for example TeX Live with `texlive-xetex` and `texlive-latex-extra`) and the
DejaVu fonts. From this folder:

```sh
pandoc K-SACCR_v1.0.1_User_Manual.md \
  --from markdown+tex_math_single_backslash \
  --pdf-engine=xelatex --toc --number-sections \
  -H header.tex -V geometry:margin=2cm -V fontsize=10pt \
  -V mainfont="DejaVu Sans" -V monofont="DejaVu Sans Mono" \
  --resource-path=. -o K-SACCR_User_Manual_1.0.1.pdf
```

Edit the Markdown, not the PDF. A later version of the manual goes in its own
folder (`docs/user-manual/v1.0.2/` and so on); this one stays as released.
