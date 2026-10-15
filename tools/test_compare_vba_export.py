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

    def test_identifier_case_is_ignored_and_counted(self) -> None:
        # The VBE gives each name one spelling across the project, as it did
        # for the netting-set ID field in the first production round trip.
        self.export("production")
        self.edit("CORE_Engine.bas", b"        ID                As String", b"        id                As String")
        self.edit("CORE_Util.bas", b"Option Explicit", b"option explicit")
        code, out = self.run_tool("production")
        self.assertEqual(code, 0, out)
        self.assertIn("Re-cased:  2 line(s)", out)

    def test_case_change_in_a_string_or_comment_fails(self) -> None:
        for path, old, new in (("src/workbook/M_Main.bas", b'"K-SACCR"', b'"k-saccr"'),
                               ("src/core/CORE_Engine.bas", b"'Netting-set ID, upper case",
                                b"'netting-set ID, upper case")):
            with self.subTest(path=path):
                for item in self.folder.iterdir():
                    item.unlink()
                self.export("production")
                self.edit(Path(path).name, old, new)
                code, out = self.run_tool("production")
                self.assertEqual(code, 1, out)
                self.assertIn(f"{path}: code differs at line", out)

    def test_fold_case_keeps_strings_and_comments(self) -> None:
        self.assertEqual(tool.fold_case("    MsgBox \"It's OK\", , MSG_TITLE 'Show It"),
                         "    msgbox \"It's OK\", , msg_title 'Show It")
        self.assertEqual(tool.fold_case("    Rem Keep THIS"), "    rem Keep THIS")
        self.assertEqual(tool.fold_case("    X = 1: Rem Keep THIS"), "    x = 1: rem Keep THIS")
        self.assertEqual(tool.fold_case("Done: REM it's KEPT"), "done: rem it's KEPT")
        self.assertEqual(tool.fold_case("10 Rem Keep THIS"), "10 rem Keep THIS")
        self.assertEqual(tool.fold_case("    Remaining = Rem2"), "    remaining = rem2")
        self.assertEqual(tool.fold_case('    S = """A"" B"'), '    s = """A"" B"')

    def test_changed_attribute_of_a_standard_module_fails(self) -> None:
        self.export("production")
        self.edit("CORE_Util.bas", b'Attribute VB_Name = "CORE_Util"\r\n',
                  b'Attribute VB_Name = "CORE_Util"\r\nAttribute VB_Description = "x"\r\n')
        code, out = self.run_tool("production")
        self.assertEqual(code, 1)
        self.assertIn("src/core/CORE_Util.bas: attributes differ", out)

    def test_document_module_attributes_are_not_compared(self) -> None:
        # Excel writes them and the synchronizer replaces only the code. The
        # first production round trip exported the sheet modules without
        # VB_Base, VB_TemplateDerived and VB_Customizable.
        self.export("production")
        for line in (b'Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"\r\n',
                     b"Attribute VB_TemplateDerived = False\r\n", b"Attribute VB_Customizable = True\r\n"):
            self.edit("shParams.cls", line, b"")
        code, out = self.run_tool("production")
        self.assertEqual(code, 0, out)

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
