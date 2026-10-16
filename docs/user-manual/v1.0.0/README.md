# K-SACCR user manual, version 1.0.0

The source of the user manual for K-SACCR 1.0.0. The PDF built from it is
attached to the GitHub Release; it is not stored in this repository.

| File | Contents |
| --- | --- |
| [`K-SACCR_v1.0.0_User_Manual.md`](K-SACCR_v1.0.0_User_Manual.md) | The manual, in pandoc Markdown |
| `assets/data_flow.png`, `assets/evidence_gate.png` | The two diagrams it includes |
| `header.tex` | The LaTeX header the PDF build needs (`\Needspace`, the link colour `KBlue`) |

The Markdown is written for pandoc and LaTeX: page-break commands
(`\Needspace`, `\newpage`, `\clearpage`) and display formulas in `\[ … \]`
appear as plain text when the file is viewed on GitHub. Read the PDF.

## Status

Documentation baseline 2026-10-16, source commit
`1152170c842d52acad166f93dbc84ea4f6546669` on the release branch. Nothing has
been released: the manual describes the release candidate and does not certify
it. Before the PDF is attached to the v1.0.0 release, as part of the
final-package work in
[#46](https://github.com/danielep71/VBA-SACCR-Toolkit/issues/46):

- the release-status box and the baseline commit are updated to the tagged
  release;
- the repository links that point to `blob/release/1.0.0/` are changed to
  `blob/v1.0.0/`;
- the PDF is rebuilt from this folder at the tag and its SHA-256 is recorded
  with the release assets.

## Building the PDF

Requires pandoc 3 and a XeLaTeX distribution with the `needspace` package
(for example TeX Live with `texlive-xetex` and `texlive-latex-extra`) and the
DejaVu fonts. From this folder:

```sh
pandoc K-SACCR_v1.0.0_User_Manual.md \
  --from markdown+tex_math_single_backslash \
  --pdf-engine=xelatex --toc --number-sections \
  -H header.tex -V geometry:margin=2cm -V fontsize=10pt \
  -V mainfont="DejaVu Sans" -V monofont="DejaVu Sans Mono" \
  --resource-path=. -o K-SACCR_User_Manual_1.0.0.pdf
```

Edit the Markdown, not the PDF. A later version of the manual goes in its own
folder (`docs/user-manual/v1.1.0/` and so on); this one stays as released.
