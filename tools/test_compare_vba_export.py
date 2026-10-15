"""Tests for tools/compare_vba_export.py, without Excel (#34).

Each test writes an "export" from the repository source at HEAD, with the
CRLF line endings the VBE uses, changes it the way a faulty round trip
would, and runs the comparison.
"""
from __future__ import annotations

import io
import json
import subprocess
import tempfile
import unittest
from contextlib import redirect_stderr, redirect_stdout
from pathlib import Path

import compare_vba_export as tool

ROOT = Path(__file__).resolve().parent.parent
MANIFEST = json.loads((ROOT / "tools" / "vba-projects.json").read_text(encoding="utf-8"))


def source(path: str) -> bytes:
    return subprocess.run(["git", "-C", str(ROOT), "show", f"HEAD:{path}"],
                          capture_output=True, check=True).stdout


class CompareExportTests(unittest.TestCase):
    def setUp(self) -> None:
        temp = tempfile.TemporaryDirectory()
        self.addCleanup(temp.cleanup)
        self.folder = Path(temp.name)

    def export(self, mode: str) -> None:
        for path in [*MANIFEST["configurations"][mode]["modules"], *MANIFEST["document_modules"]]:
            (self.folder / Path(path).name).write_bytes(source(path).replace(b"\n", b"\r\n"))

    def run_tool(self, mode: str) -> tuple[int, str]:
        out, err = io.StringIO(), io.StringIO()
        with redirect_stdout(out), redirect_stderr(err):
            code = tool.main(["--root", str(ROOT), "--export", str(self.folder), "--mode", mode])
        return code, out.getvalue() + err.getvalue()

    def edit(self, name: str, old: bytes, new: bytes) -> None:
        path = self.folder / name
        data = path.read_bytes()
        self.assertIn(old, data)
        path.write_bytes(data.replace(old, new, 1))

    def test_exact_export_passes_in_both_modes(self) -> None:
        for mode in ("production", "development"):
            with self.subTest(mode=mode):
                for path in self.folder.iterdir():
                    path.unlink()
                self.export(mode)
                code, out = self.run_tool(mode)
                self.assertEqual(code, 0, out)
                self.assertIn("RESULT=PASS", out)

    def test_trailing_blank_lines_and_line_endings_are_ignored(self) -> None:
        self.export("production")
        path = self.folder / "CORE_Util.bas"
        path.write_bytes(path.read_bytes().replace(b"\r\n", b"\n") + b"\r\n\r\n")
        self.assertEqual(self.run_tool("production")[0], 0)

    def test_changed_code_fails_with_its_line(self) -> None:
        self.export("production")
        self.edit("M_Main.bas", b'"K-SACCR"', b'"SA-CCR"')
        code, out = self.run_tool("production")
        self.assertEqual(code, 1)
        self.assertIn("src/workbook/M_Main.bas: code differs at line", out)
        self.assertIn("'SA-CCR'", out.replace('"', "'"))

    def test_identifier_case_change_fails(self) -> None:
        # The VBE re-cases identifiers project-wide; the source must follow.
        self.export("production")
        self.edit("CORE_Util.bas", b"Option Explicit", b"option explicit")
        self.assertEqual(self.run_tool("production")[0], 1)

    def test_changed_attribute_fails(self) -> None:
        self.export("production")
        self.edit("shParams.cls", b"Attribute VB_Exposed = True", b"Attribute VB_Exposed = False")
        code, out = self.run_tool("production")
        self.assertEqual(code, 1)
        self.assertIn("src/workbook/shParams.cls: attributes differ", out)

    def test_missing_and_extra_components_fail(self) -> None:
        self.export("development")
        (self.folder / "SACCR_Formulas.bas").unlink()
        code, out = self.run_tool("production")
        self.assertEqual(code, 1)
        self.assertIn("missing from the export: SACCR_Formulas", out)
        self.assertIn("not in the source for this mode: TEST_Harness", out)

    def test_component_without_vb_name_is_an_error(self) -> None:
        self.export("production")
        (self.folder / "Stray.bas").write_bytes(b"Option Explicit\r\n")
        code, out = self.run_tool("production")
        self.assertEqual(code, 2)
        self.assertIn("Stray.bas: expected one Attribute VB_Name line, found 0", out)


    def test_userform_export_header_is_parsed(self) -> None:
        # A form export: VERSION 5.00, a title-case Begin ... End block with a
        # nested block, then the attributes and the code.
        raw = ("VERSION 5.00\r\n"
               "Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UF_Picker \r\n"
               "   Caption         =   \"Pick\"\r\n"
               "   Begin {0BA686C6-F7D3-101A-993E-0000C0EF6F5E} Ctl \r\n"
               "   End\r\n"
               "End\r\n"
               'Attribute VB_Name = "UF_Picker"\r\n'
               "Attribute VB_GlobalNameSpace = False\r\n"
               "Option Explicit\r\n").encode("cp1252")
        component = tool.parse(raw, "UF_Picker.frm")
        self.assertEqual(component.name, "UF_Picker")
        self.assertEqual(component.attributes,
                         ('Attribute VB_Name = "UF_Picker"', "Attribute VB_GlobalNameSpace = False"))
        self.assertEqual(component.body, ("Option Explicit",))


if __name__ == "__main__":
    unittest.main()
