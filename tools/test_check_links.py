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

    def test_report_updates_existing_issue(self) -> None:
        with patch.object(external, 'api', side_effect=[[
            {'number': 7, 'body': external.MARKER, 'state': 'closed'}], {}]) as api:
            external.publish('owner/repo', external.MARKER, True)
            self.assertEqual(api.call_args.args[1:3], ('PATCH', '/issues/7'))
            self.assertEqual(api.call_args.args[3]['state'], 'open')


if __name__ == '__main__':
    unittest.main()
