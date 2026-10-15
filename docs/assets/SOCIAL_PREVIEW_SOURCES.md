# Social preview

`SACCR_social_preview.html` is the source for `SACCR_social_preview.png`, the
repository's social preview image (1280 x 640 px, under GitHub's 1 MB limit).
It is a promotional image; it does not certify the engine or replace the
methodology source-register decisions.

- **Figures:** the netting-set card shows CRE99 Example 5 in USD thousands:
  MPOR 14 business days (CRE99.83), aggregate add-on 1,401 (CRE99.95),
  multiplier 0.958 (CRE99.96) and EAD 1,879 (CRE99.97). PFE is
  0.958 x 1,401 = 1,342 and RC is 0, since 1.4 x (0 + 1,342) = 1,879. The strip
  shows the published EAD of Examples 1 to 5 (CRE99.40, .58, .76, .79, .97),
  the `published` values in `tests/expected/cre99-example-*.bcbs.json`.
- **Curve:** the PFE multiplier m = min(1, F + (1 - F) exp((V - C) /
  (2 (1 - F) AddOn))) with F = 5% (CRE52.23), plotted for (V - C) / AddOn from
  -4 to +1.
- **Name:** the title is the project name, K-SACCR; the method, SA-CCR, is
  named in the label beside the status.
- **Status:** the image carries "Under development" while no release exists;
  update it at the first release.

Regenerate by opening the HTML in a Chromium browser with a 1280 x 640 viewport
at device scale factor 1, after the Google Fonts (Barlow Condensed, Barlow,
JetBrains Mono) have loaded, and capturing the full page. Then upload the PNG
in **Settings > General > Social preview**; GitHub has no API for this setting.
