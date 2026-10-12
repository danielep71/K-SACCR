#!/usr/bin/env python3
"""Offline tests for the Wiki publication contract."""
from __future__ import annotations

import tempfile
import unittest
from pathlib import Path
from typing import Any
from unittest.mock import patch

from check_wiki import bundle, compare_bundle, render_page, sidebar, validate_catalogue, write_bundle


def fixture(root: Path) -> tuple[dict[str, Any], set[str]]:
    (root / "docs/wiki").mkdir(parents=True)
    (root / "README.md").write_text("Authority", encoding="utf-8")
    (root / "docs/wiki/Home.md").write_text(
        "# Home\n\n> **Guide, not policy:** [README](../../README.md) governs.\n", encoding="utf-8")
    (root / "docs/wiki/_Sidebar.md").write_text("", encoding="utf-8")
    data: dict[str, Any] = {"schema_version": 1,
                            "pages": [{"name": "Home", "title": "Home", "authority": "README.md"}]}
    files = {"README.md", "docs/wiki/Home.md", "docs/wiki/_Sidebar.md", "docs/wiki/catalogue.json"}
    return data, files


class WikiTests(unittest.TestCase):
    def test_valid_catalogue(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            data, files = fixture(root)
            validate_catalogue(root, data, files)

    def test_optional_footer_is_managed(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            data, files = fixture(root)
            (root / "docs/wiki/_Footer.md").write_text("Footer", encoding="utf-8")
            files.add("docs/wiki/_Footer.md")
            validate_catalogue(root, data, files)
            with patch("check_wiki.repository_name", return_value="owner/repo"):
                result = bundle(root, data, "a" * 40, files)
            self.assertIn("_Footer.md", result)
            self.assertIn(b"_Footer.md", result["Wiki-Source.json"])
            self.assertNotIn(b"Published from", result["_Footer.md"])
            self.assertIn(b"Published from", result["Home.md"])

    def test_untracked_footer_is_not_published(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            data, files = fixture(root)
            (root / "docs/wiki/_Footer.md").write_text("Footer", encoding="utf-8")
            validate_catalogue(root, data, files)
            with patch("check_wiki.repository_name", return_value="owner/repo"):
                result = bundle(root, data, "a" * 40, files)
            self.assertNotIn("_Footer.md", result)

    def test_unregistered_page_fails(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            data, files = fixture(root)
            files.add("docs/wiki/Extra.md")
            with self.assertRaisesRegex(ValueError, "catalogue differs"):
                validate_catalogue(root, data, files)

    def test_sidebar_order(self) -> None:
        data: dict[str, Any] = {"pages": [{"name": "Home", "title": "Home"},
                                         {"name": "Next", "title": "Next"}]}
        text = sidebar(data)
        self.assertLess(text.index("Home.md"), text.index("Next.md"))

    def test_render_links(self) -> None:
        text = "[Next](Next.md) [Rule](../../README.md#status) [Web](https://example.com)"
        result = render_page(text, "Home.md", {"Home", "Next"}, "owner/repo", "a" * 40)
        self.assertIn("[Next](Next)", result)
        self.assertIn("/README.md#status", result)
        self.assertIn("[Web](https://example.com)", result)

    def test_compare_detects_drift(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            (root / "Home.md").write_bytes(b"changed")
            (root / "Extra.md").write_bytes(b"extra")
            self.assertEqual(len(compare_bundle({"Home.md": b"page", "Missing.md": b"x"}, root)), 3)

    def test_write_bundle_refuses_overwrite(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp) / "source"
            root.mkdir()
            destination = Path(tmp) / "export"
            write_bundle({"Home.md": b"page"}, destination, root)
            with self.assertRaisesRegex(ValueError, "must not exist"):
                write_bundle({"Home.md": b"changed"}, destination, root)


if __name__ == "__main__":
    unittest.main()
