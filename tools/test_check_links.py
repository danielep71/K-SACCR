"""Offline link and external-response regression fixtures."""
from __future__ import annotations

import tempfile
import unittest
import urllib.error
from email.message import Message
from pathlib import Path
from unittest.mock import patch

import check_external_links as external
import check_links as local


class LinkTests(unittest.TestCase):
    def test_destinations_and_anchors(self) -> None:
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            target = root / "target.md"
            target.write_text('# Repeat\n# Repeat\n<a id="explicit"></a>\n', encoding="utf-8")
            (root / "empty.md").touch()
            source = root / "source.md"
            source.write_text('[ok](target.md#repeat-1) [ok](target.md#explicit)\n'
                              '[bad](missing.md) [bad](empty.md) [bad](../outside.md)\n'
                              '[bad](target.md#absent)\n```md\n[ignored](absent.md)\n```\n'
                              '`[ignored](absent.md)`\n[external](https://example.com)\n', encoding="utf-8")
            self.assertEqual([item['reason'] for item in local.check(root, [source])],
                             ['missing target', 'empty target', 'path leaves repository', 'missing fragment'])

    def test_code_labels_and_autolinks(self) -> None:
        self.assertEqual(local.links("[`missing.md`](missing.md#absent)"), [(1, "missing.md#absent")])
        self.assertEqual(local.links("<https://example.com>"), [(1, "https://example.com")])
        self.assertEqual(local.links("`<https://example.com>`"), [])

    def test_linked_badge_checks_both_destinations(self) -> None:
        self.assertEqual(local.links("![](missing.png)"), [(1, "missing.png")])
        self.assertEqual(local.links("[![Contributing](https://example.com/badge)](CONTRIBUTING.md#absent)"),
                         [(1, "CONTRIBUTING.md#absent"), (1, "https://example.com/badge")])

    def test_references_and_duplicate_collision(self) -> None:
        self.assertEqual(local.links('[a][ref]\n[ref]: target.md\n'), [(1, 'target.md')])
        self.assertEqual(local.anchors('# A\n# A\n# A-1\n'), {'a', 'a-1', 'a-1-1'})

    def test_external_retries_classify_uncertainty(self) -> None:
        for code, expected in ((404, 'broken'), (410, 'broken'), (429, 'inconclusive'),
                               (503, 'inconclusive'), (403, 'inconclusive')):
            with self.subTest(code=code), patch('urllib.request.urlopen',
                    side_effect=urllib.error.HTTPError('https://example.com', code, '', Message(), None)), \
                    patch('time.sleep'):
                self.assertEqual(external.inspect('https://example.com')[0], expected)

    def test_github_repository_pages_map_to_the_api(self) -> None:
        base = 'https://github.com/owner/repo'
        for page, targets in (
                ('', ['/repos/owner/repo']), ('.git', ['/repos/owner/repo']),
                ('/stargazers', ['/repos/owner/repo']),
                ('/issues?q=is%3Aissue%20label%3AP1', ['/repos/owner/repo']),
                ('/wiki/Page#why', ['/repos/owner/repo']), ('/compare/v1...HEAD', ['/repos/owner/repo']),
                ('/commits/main', ['/repos/owner/repo/commits?sha=main&per_page=1']),
                ('/commits/release/1.3.0', ['/repos/owner/repo/commits?sha=release%2F1.3.0&per_page=1']),
                ('/issues/100', ['/repos/owner/repo/issues/100']), ('/pull/7', ['/repos/owner/repo/issues/7']),
                ('/milestone/5', ['/repos/owner/repo/milestones/5']),
                ('/releases/latest', ['/repos/owner/repo/releases/latest']),
                ('/releases/tag/v1.2.0', ['/repos/owner/repo/releases/tags/v1.2.0']),
                ('/releases/download/v1.2.0/K-SACCR.xlsm', ['/repos/owner/repo/releases/tags/v1.2.0']),
                ('/blob/release/1.3.0/docs/A.md', ['/repos/owner/repo/contents/1.3.0/docs/A.md?ref=release',
                                                   '/repos/owner/repo/contents/docs/A.md?ref=release%2F1.3.0',
                                                   '/repos/owner/repo/contents/A.md?ref=release%2F1.3.0%2Fdocs',
                                                   '/repos/owner/repo/contents?ref=release%2F1.3.0%2Fdocs%2FA.md']),
                ('/tree/main', ['/repos/owner/repo/contents?ref=main']),
                ('/tree/release/1.3.0', ['/repos/owner/repo/contents/1.3.0?ref=release',
                                         '/repos/owner/repo/contents?ref=release%2F1.3.0']),
                ('/settings', []), ('/blob', [])):
            with self.subTest(page=page):
                self.assertEqual(external.api_targets(base + page), targets)
        self.assertEqual(external.api_targets('https://github.com/owner'), ['/users/owner'])
        for url in ('https://github.com/', 'https://example.com/owner/repo'):
            with self.subTest(url=url):
                self.assertEqual(external.api_targets(url), [])

    def test_github_repository_page_verified_through_the_api(self) -> None:
        class Response:
            status = 200

            def __enter__(self) -> 'Response':
                return self

            def __exit__(self, *_: object) -> None:
                return None

        def answer(request: urllib.request.Request, timeout: int) -> Response:
            seen.append((request.full_url, request.get_header('Authorization')))
            if request.full_url.startswith('https://api.github.com/'):
                if request.full_url.endswith('?ref=release'):
                    raise urllib.error.HTTPError(request.full_url, 404, '', Message(), None)
                return Response()
            raise urllib.error.HTTPError(request.full_url, 429, '', Message(), None)

        seen: list[tuple[str, str | None]] = []
        with patch('urllib.request.urlopen', side_effect=answer), patch('time.sleep'), \
                patch.dict('os.environ', {'GITHUB_TOKEN': 'token'}):
            self.assertEqual(external.inspect('https://github.com/owner/repo/stargazers'), ('ok', 'api 200'))
            self.assertEqual(external.inspect('https://github.com/owner/repo/blob/release/1.3.0/docs/A.md'),
                             ('ok', 'api 200'))
        self.assertEqual(seen[0], ('https://api.github.com/repos/owner/repo', 'Bearer token'))
        self.assertEqual([url for url, _ in seen[1:]],
                         ['https://api.github.com/repos/owner/repo/contents/1.3.0/docs/A.md?ref=release'] * 3
                         + ['https://api.github.com/repos/owner/repo/contents/docs/A.md?ref=release%2F1.3.0'])

    def test_github_repository_page_falls_back_to_the_web(self) -> None:
        def answer(request: urllib.request.Request, timeout: int) -> None:
            code = 404 if request.full_url.startswith('https://api.github.com/') else 429
            raise urllib.error.HTTPError(request.full_url, code, '', Message(), None)

        with patch('urllib.request.urlopen', side_effect=answer), patch('time.sleep'), \
                patch.dict('os.environ', {}, clear=True):
            self.assertEqual(external.inspect('https://github.com/apps/name'),
                             ('inconclusive', 'api 404, 404, 404; 429, 429, 429'))

    def test_report_updates_existing_issue(self) -> None:
        with patch.object(external, 'api', side_effect=[[
            {'number': 7, 'body': external.MARKER, 'state': 'closed'}], {}]) as api:
            external.publish('owner/repo', external.MARKER, True)
            self.assertEqual(api.call_args.args[1:3], ('PATCH', '/issues/7'))
            self.assertEqual(api.call_args.args[3]['state'], 'open')


if __name__ == '__main__':
    unittest.main()
