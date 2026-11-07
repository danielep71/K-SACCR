# K-SACCR user manual, version 1.2.0

The source of the user manual for K-SACCR 1.2.0. The PDF built from it is
attached to the
[1.2.0 GitHub Release](https://github.com/danielep71/K-SACCR/releases/tag/v1.2.0);
it is not stored in this repository.

| File | Contents |
| --- | --- |
| [`K-SACCR_v1.2.0_User_Manual.md`](K-SACCR_v1.2.0_User_Manual.md) | The manual, in pandoc Markdown |
| `assets/data_flow.png`, `assets/evidence_gate.png`, `assets/rc_pfe.png` | The three diagrams it includes, unchanged from 1.1.0; `rc_pfe.png` is a copy of `docs/assets/SACCR_RC_PFE.png` |
| `header.tex` | The LaTeX header the PDF build needs (`\Needspace`, the link colour `KBlue`), unchanged from 1.1.0 |

The Markdown is written for pandoc and LaTeX: page-break commands
(`\Needspace`, `\newpage`, `\clearpage`) and display formulas in `\[ … \]`
appear as plain text when the file is viewed on GitHub. Read the PDF.

## Status

Released with 1.2.0 (tag `v1.2.0`, 7 November 2026). Started from the 1.1.0
manual and updated for the changes of 1.2.0:

- the **K-SACCR** Ribbon tab, in a new chapter that replaces *Buttons and
  their exact roles*: the six groups, every control with its keytip, macro and
  effect, the enabled states, the three Status labels, the **Go to** menu,
  when the tab updates and what happens after a VBA error, and the four
  buttons that remain on each staging sheet;
- the removal of the 12 calculation and export buttons from the sheets,
  throughout: the walkthrough, the CSV import, the evidence export, the
  troubleshooting table and the quick-reference card now point to the tab;
- the `Home` sheet, on which the workbook opens, and the map of the 19
  worksheets in tab order, with `README` last; `Home` in the protection
  matrix;
- **Restore parameters** and its automation entry point
  `RestoreRegulatoryParameters_Silent`: a new section on its scope,
  confirmation, unchanged case, failure rollback and refusal, and its effect
  on `Results` and on the evidence export, with the Params, what-if,
  maintenance-unlock and freshness sections that refer to it;
- where a command leaves the user: `Results` after a completed run, `Checks`
  after a stopped run and **Validate inputs**;
- the one Windows function the workbook calls, `RtlMoveMemory`, in the
  supported environment;
- kept apart from the Ribbon changes, the two input-contract corrections: the
  new `NettingSets` clearing-role header *Client clearing role (N / CM /
  CLIENT / CCP; MULTI rejected)* and the fingerprint it changes
  (`inputs=119D94E2` for the demonstration portfolio), and the zero notional
  now accepted on every path, a negative one still rejected;
- the note on unlisted multi-name credit positions, which are not supported;
- *Moving inputs from 1.1.0* in place of *Moving inputs from 1.0.2*, the test
  summary, the known limitations, the checklists, the troubleshooting table,
  the delivered files, the glossary and the quick-reference card.

The calculation is unchanged, and so is every rule the manual states about
it. As before, the manual is written for someone who receives only the
workbook `K-SACCR_1.2.0.xlsm` and `SHA256SUMS.txt`: it does not mention the
source repository, building the workbook from source or the release process,
and of the testing it gives only the summary result. The release's PDF is built
from this folder at the release tag, and its SHA-256 is listed in the
release's `SHA256SUMS.txt`. The 1.0.0, 1.0.1, 1.0.2 and 1.1.0 manuals stay as
released in [`../v1.0.0/`](../v1.0.0/README.md),
[`../v1.0.1/`](../v1.0.1/README.md), [`../v1.0.2/`](../v1.0.2/README.md) and
[`../v1.1.0/`](../v1.1.0/README.md).

## Building the PDF

Requires pandoc 3 and a XeLaTeX distribution with the `needspace` package
(for example TeX Live with `texlive-xetex` and `texlive-latex-extra`) and the
DejaVu fonts. From this folder:

```sh
pandoc K-SACCR_v1.2.0_User_Manual.md \
  --from markdown+tex_math_single_backslash \
  --pdf-engine=xelatex --toc --number-sections \
  -H header.tex -V geometry:margin=2cm -V fontsize=10pt \
  -V mainfont="DejaVu Sans" -V monofont="DejaVu Sans Mono" \
  --resource-path=. -o K-SACCR_User_Manual_1.2.0.pdf
```

Edit the Markdown, not the PDF. A later version of the manual goes in its own
folder (`docs/user-manual/v1.2.1/` and so on); once 1.2.0 is released, this
one stays as released.
