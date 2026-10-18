# Project website

The public K-SACCR landing page is maintained in [`index.html`](index.html).
Its purpose is to introduce the project, make its scope discoverable, and link
visitors to the authoritative repository documentation. It is not a browser
calculator or a software release.

## Publishing

After the normal task PR into the active release branch, integrate the reviewed
website change into `main` through an owner-authorized integration PR. In
**Settings > Pages**, use **Deploy from a branch**, branch **main**, folder
**/docs**. GitHub's built-in Pages deployment publishes this directory.
[`.nojekyll`](.nojekyll) disables Jekyll processing: the landing page needs no
package installation, build tool, or external font. Its only script, at the
end of the page, starts the release downloads in a hidden frame so the page
does not visibly pass through GitHub; the links work without it.

The public address is <https://danielep71.github.io/VBA-SACCR-Toolkit/>. Set
the repository's **About > Website** field to this address after verifying a
successful deployment. The source HTML uses the existing screenshots in
[`assets/`](assets/README.md); their provenance and update rules remain there.
Other documentation links point to the rendered files on GitHub because
Markdown is not converted to HTML by this static deployment.

## Maintenance and checks

- Keep the page's scope, release status, evidence classes, and installation
  links aligned with the root README and methodology. At each release, update
  the release named on the page and in the social preview together. The
  download buttons and the `SHA256SUMS.txt` link point straight at the release
  assets, whose URLs and file names carry the version: update all three and
  check that each still starts a download, with the `data-direct-download`
  attribute that the script uses.
- Do not copy transient test counts onto the landing page. Link the current
  evidence instead; illustrative results must not be called independent
  validation.
- Keep the canonical URL, Open Graph URL/image, and [`sitemap.xml`](sitemap.xml)
  aligned if the repository or hosting address changes. The sitemap contains
  only the landing page, not GitHub repository URLs or raw Markdown files.
- Run `python tools/check.py` before pushing and the normal committed/hosted
  gates before merge. No Excel execution is needed for website-only changes.
- Verify desktop and narrow-screen layout, keyboard navigation, local anchors,
  existing image paths, external destinations, and the live sitemap. The page
  must work without JavaScript.

## Search Console (separate setup)

Publishing does not submit the site to Google or guarantee indexing. Add the
site as a URL-prefix property in Google Search Console. Add the exact HTML
verification tag supplied by the owner to the page's `head` through the normal
PR workflow; never insert an invented token. After verification, the owner can
submit `sitemap.xml` and request indexing of the landing page. Do not put a
project-level `robots.txt` here: crawlers use the host-root `/robots.txt`.

The owner-supplied HTML verification tag is maintained in `index.html`. Keep
it in the page after verification so Google can recheck ownership. Its presence
does not establish that the owner has completed verification in Search Console.
