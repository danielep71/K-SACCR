"""Offline negative fixtures for the SACCR source gate."""
import shutil
import subprocess
import tempfile
import unittest
import zipfile
from pathlib import Path

import check_source
import check_vba_public_api as public_api

ROOT = Path(__file__).resolve().parents[1]
CHANGELOG = "# Changelog\n\n## [Unreleased]\n\n[Unreleased]: https://example.invalid\n"
MODULE = 'Attribute VB_Name = "SACCR_Test"\r\nOption Explicit\r\n\r\nPublic Sub Run()\r\nEnd Sub\r\n'


class SourceGateTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.git("init", "-q")
        shutil.copy(ROOT / ".gitattributes", self.root / ".gitattributes")
        self.write("CHANGELOG.md", CHANGELOG)
        # The synchronizer and the template are required (#137): by default the
        # synchronizer reads every documented VBA home and the template has no
        # CodeNames, so neither adds findings of its own.
        for home in check_source.VBA_HOMES:
            (self.root / home).mkdir(parents=True, exist_ok=True)
        self.write_sync(*(home.rstrip("/").replace("/", "\\") for home in check_source.VBA_HOMES))
        self.write_workbook("src/workbook/SACCR_Template.xlsx", {
            "[Content_Types].xml": "<Types/>",
            "_rels/.rels": '<Relationships Target="xl/workbook.xml"/>',
            "xl/workbook.xml": "<workbook/>"})

    def git(self, *args: str) -> str:
        return subprocess.run(["git", "-C", str(self.root), *args], check=True,
                              capture_output=True, text=True).stdout

    def write(self, path: str, text: str) -> None:
        target = self.root / path
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_bytes(text.encode("cp1252"))

    def findings(self) -> list[str]:
        self.git("add", "-A")
        findings: list[str] = check_source.run_check(self.root)["findings"]
        return findings

    def test_clean_component_passes_and_is_stored_lf(self) -> None:
        self.write("src/modules/SACCR_Test.bas", MODULE)
        self.assertEqual(self.findings(), [])
        self.assertEqual(self.git("ls-files", "--eol", "src/modules/SACCR_Test.bas").split()[0], "i/lf")

    def test_vb_name_must_match_filename(self) -> None:
        self.write("src/modules/M_Other.bas", MODULE)
        self.assertIn("src/modules/M_Other.bas: VB_Name must match the filename exactly", self.findings())

    def test_duplicate_component_name(self) -> None:
        self.write("src/modules/SACCR_Test.bas", MODULE)
        self.write("tests/SACCR_Test.bas", MODULE)
        self.assertTrue(any("duplicate component name" in f for f in self.findings()))

    def test_option_explicit_in_comment_does_not_count(self) -> None:
        self.write("src/modules/SACCR_Test.bas", MODULE.replace("Option Explicit", "' Option Explicit"))
        self.assertIn("src/modules/SACCR_Test.bas: missing Option Explicit", self.findings())

    def test_unclosed_string_literal_is_reported_with_its_line(self) -> None:
        body = '    x = "=""" & y & """""\r\n'
        self.write("src/modules/SACCR_Test.bas", MODULE.replace("End Sub", body + "End Sub"))
        self.assertIn("src/modules/SACCR_Test.bas:5: string literal not closed", self.findings())

    def test_unclosed_literal_after_a_quoted_sheet_name_is_reported(self) -> None:
        body = "    Debug.Print ['Sheet 1'!A1], \"open\r\n"
        self.write("src/modules/SACCR_Test.bas", MODULE.replace("End Sub", body + "End Sub"))
        self.assertIn("src/modules/SACCR_Test.bas:5: string literal not closed", self.findings())

    def test_quotes_in_strings_and_comments_are_not_literals(self) -> None:
        body = ('    x = "=""" & y & """" \' a "comment\r\n'
                '    x = "it\'s": Rem say "hi\r\n'
                '    \' a comment _\r\n'
                '    continued with a " quote\r\n'
                '100 Rem say "hi\r\n'
                '    Debug.Print [\'Sheet 1\'!A1], "x" \' "\r\n')
        self.write("src/modules/SACCR_Test.bas", MODULE.replace("End Sub", body + "End Sub"))
        self.assertEqual(self.findings(), [])

    def test_form_requires_tracked_frx_within_bounds(self) -> None:
        form = 'VERSION 5.00\r\nBegin {X} F_Test\r\n   OleObjectBlob = "F_Test.frx":0010\r\nEnd\r\n'
        form += MODULE.replace("SACCR_Test", "F_Test")
        self.write("src/forms/F_Test.frm", form)
        self.assertIn("src/forms/F_Test.frm: missing or unsafe form resource F_Test.frx", self.findings())
        (self.root / "src/forms/F_Test.frx").write_bytes(b"\0" * 8)
        self.assertIn("src/forms/F_Test.frm: resource offset is outside F_Test.frx", self.findings())
        (self.root / "src/forms/F_Test.frx").write_bytes(b"\0" * 32)
        self.assertEqual(self.findings(), [])

    def test_vba_outside_documented_locations_is_rejected(self) -> None:
        self.write("src/M_Test.bas", MODULE)
        self.write("misc/Old.frx", "x")
        found = self.findings()
        self.assertIn("src/M_Test.bas: VBA component outside the documented source locations", found)
        self.assertIn("misc/Old.frx: VBA component outside the documented source locations", found)

    def test_core_module_requires_option_private_module(self) -> None:
        core = MODULE.replace("SACCR_Test", "CORE_Test")
        self.write("src/core/CORE_Test.bas", core)
        self.assertIn("src/core/CORE_Test.bas: core modules must declare Option Private Module", self.findings())
        self.write("src/core/CORE_Test.bas", core.replace("Option Explicit", "Option Explicit\r\nOption Private Module"))
        self.assertEqual(self.findings(), [])

    def write_workbook(self, path: str, parts: dict[str, str]) -> None:
        target = self.root / path
        target.parent.mkdir(parents=True, exist_ok=True)
        with zipfile.ZipFile(target, "w") as package:
            for name, text in parts.items():
                package.writestr(name, text)

    def test_committed_workbook_without_metadata_passes(self) -> None:
        self.write_workbook("src/workbook/T.xlsx", {"[Content_Types].xml": "<Types/>",
                                                    "_rels/.rels": '<Relationships Target="xl/workbook.xml"/>',
                                                    "xl/workbook.xml": "<workbook/>"})
        self.assertEqual(self.findings(), [])

    def test_committed_workbook_with_metadata_or_vba_is_rejected(self) -> None:
        self.write_workbook("src/workbook/T.xlsx", {"[Content_Types].xml": "<Types/>",
                                                    "_rels/.rels": '<Relationship Target="docProps/core.xml"/>',
                                                    "docProps/core.xml": "<coreProperties/>",
                                                    "xl/vbaProject.bin": "x"})
        found = self.findings()
        self.assertIn("src/workbook/T.xlsx: must not contain docProps/core.xml", found)
        self.assertIn("src/workbook/T.xlsx: must not contain xl/vbaProject.bin", found)
        self.assertIn("src/workbook/T.xlsx: package still references document properties or a VBA project", found)

    def test_dangling_vba_relationship_is_rejected(self) -> None:
        for relationship_part in ("xl/_rels/workbook.xml.rels", "xl/worksheets/_rels/sheet1.xml.rels"):
            with self.subTest(part=relationship_part):
                self.write_workbook("src/workbook/T.xlsx", {
                    "[Content_Types].xml": "<Types/>",
                    "_rels/.rels": '<Relationships Target="xl/workbook.xml"/>',
                    relationship_part: '<Relationship Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/>',
                })
                self.assertTrue(any("still references" in f for f in self.findings()))

    def test_unreadable_workbook_is_rejected(self) -> None:
        self.write("src/workbook/T.xlsx", "not a zip")
        self.assertTrue(any("not a readable workbook package" in f for f in self.findings()))

    def write_sync(self, *directories: str) -> None:
        entries = ",\r\n".join(f'        (Join-Path $repoFullPath "{d}")' for d in directories)
        self.write("tools/Sync-SACCR-VBA.ps1", f"$sourceRoots = @(\r\n{entries}\r\n    )\r\n")

    def test_component_outside_sync_roots_is_rejected(self) -> None:
        self.write_sync("src\\modules")
        self.write("src/modules/SACCR_Test.bas", MODULE)
        self.write("tests/modules/TEST_Test.bas", MODULE.replace("SACCR_Test", "TEST_Test"))
        self.assertEqual(self.findings(), ["tests/modules/TEST_Test.bas: not synchronized; add its "
                                           "directory to $sourceRoots in tools/Sync-SACCR-VBA.ps1"])
        self.write_sync("src\\modules", "tests\\modules")
        self.assertEqual(self.findings(), [])

    def test_missing_sync_root_is_rejected(self) -> None:
        self.write_sync("src\\modules", "src\\nowhere")
        self.write("src/modules/SACCR_Test.bas", MODULE)
        self.assertEqual(self.findings(),
                         ["tools/Sync-SACCR-VBA.ps1: source directory src/nowhere/ does not exist"])

    def test_commented_sync_root_is_not_counted(self) -> None:
        self.write("tests/modules/TEST_Test.bas", MODULE.replace("SACCR_Test", "TEST_Test"))
        unsynchronized = ("tests/modules/TEST_Test.bas: not synchronized; add its directory to "
                          "$sourceRoots in tools/Sync-SACCR-VBA.ps1")
        for commented in ('        # (Join-Path $repoFullPath "tests\\modules")',
                          '        <# (Join-Path $repoFullPath "tests\\modules")\r\n        #>'):
            with self.subTest(entry=commented):
                self.write("tools/Sync-SACCR-VBA.ps1",
                           '$sourceRoots = @(\r\n        (Join-Path $repoFullPath "src\\modules"),'
                           f"\r\n{commented}\r\n    )\r\n")
                self.assertEqual(self.findings(), [unsynchronized])
        self.write("tools/Sync-SACCR-VBA.ps1",
                   '$sourceRoots = @(\r\n        (Join-Path $repoFullPath "tests\\modules")  # unit tests'
                   "\r\n    )\r\n")
        self.assertEqual(self.findings(), [])

    def test_commented_sync_declaration_is_not_the_active_list(self) -> None:
        # A complete list inside a block comment before an incomplete active
        # list must not hide the missing root (#144).
        self.write("src/modules/SACCR_Test.bas", MODULE)
        self.write("tests/modules/TEST_Test.bas", MODULE.replace("SACCR_Test", "TEST_Test"))
        unsynchronized = ("tests/modules/TEST_Test.bas: not synchronized; add its directory to "
                          "$sourceRoots in tools/Sync-SACCR-VBA.ps1")
        old = ('$sourceRoots = @(\r\n        (Join-Path $repoFullPath "src\\modules"),\r\n'
               '        (Join-Path $repoFullPath "tests\\modules")\r\n    )\r\n')
        active = '$sourceRoots = @(\r\n        (Join-Path $repoFullPath "src\\modules")\r\n    )\r\n'
        for before in (f"<#\r\n{old}#>\r\n", f"<# {old}#>\r\n",
                       "".join(f"# {line}\r\n" for line in old.split("\r\n") if line)):
            with self.subTest(before=before[:12]):
                self.write("tools/Sync-SACCR-VBA.ps1", before + active)
                self.assertEqual(self.findings(), [unsynchronized])
        # A commented declaration and no active list: no list is found.
        self.write("tools/Sync-SACCR-VBA.ps1", f"<#\r\n{old}#>\r\n")
        self.assertIn("tools/Sync-SACCR-VBA.ps1: $sourceRoots list not found", self.findings())
        # A complete active list after a block comment is accepted.
        self.write("tools/Sync-SACCR-VBA.ps1", "<# synchronizer\r\n   notes #>\r\n" + old)
        self.assertEqual(self.findings(), [])
        # A quoted "<#" before the list does not open a comment, also in a
        # string or here-string over several lines.
        for before in ('$label = "<# not a comment"\r\n',
                       '$label = "first line\r\n<# second line"\r\n',
                       "$help = @'\r\n<# text\r\n'@\r\n",
                       '$help = @"\r\n  <# text $x\r\n"@\r\n'):
            with self.subTest(string=before[:16]):
                self.write("tools/Sync-SACCR-VBA.ps1", before + old)
                self.assertEqual(self.findings(), [])
        # A declaration inside a here-string is text, not the active list.
        self.write("tools/Sync-SACCR-VBA.ps1", "$doc = @'\r\n" + old + "'@\r\n" + active)
        self.assertEqual(self.findings(), [unsynchronized])

    def test_backtick_escapes_in_sync_preamble(self) -> None:
        # In a double-quoted string a backtick escapes the next character, so
        # `" does not close the string and a following <# is text; `` is a
        # literal backtick and the quote after it does close the string. In a
        # single-quoted string a backtick is literal (#188).
        self.write("src/modules/SACCR_Test.bas", MODULE)
        self.write("tests/modules/TEST_Test.bas", MODULE.replace("SACCR_Test", "TEST_Test"))
        unsynchronized = ("tests/modules/TEST_Test.bas: not synchronized; add its directory to "
                          "$sourceRoots in tools/Sync-SACCR-VBA.ps1")
        old = ('$sourceRoots = @(\r\n        (Join-Path $repoFullPath "src\\modules"),\r\n'
               '        (Join-Path $repoFullPath "tests\\modules")\r\n    )\r\n')
        active = '$sourceRoots = @(\r\n        (Join-Path $repoFullPath "src\\modules")\r\n    )\r\n'
        for before in ('$x = "literal `"<#`""\r\n',
                       '$x = "`"# not a comment `"<# nor this"\r\n',
                       '$x = "first `"\r\n<# still text"\r\n'):
            with self.subTest(string=before[:24]):
                self.write("tools/Sync-SACCR-VBA.ps1", before + old)
                self.assertEqual(self.findings(), [])
        # After an escaped backtick the quote closes the string, so the <#
        # opens a real comment and the complete list inside it is not active;
        # the same holds after a single-quoted string ending in a backtick.
        for before in (f'$x = "C:\\path``" <#\r\n{old}#>\r\n',
                       f"$y = 'C:\\path`' <#\r\n{old}#>\r\n"):
            with self.subTest(string=before[:24]):
                self.write("tools/Sync-SACCR-VBA.ps1", before + active)
                self.assertEqual(self.findings(), [unsynchronized])

    def test_unsupported_sync_entry_is_rejected(self) -> None:
        self.write("tools/Sync-SACCR-VBA.ps1",
                   '$sourceRoots = @(\r\n        (Join-Path $repoFullPath "src\\modules"),\r\n'
                   "        (Get-Item $env:SOURCE)\r\n")
        found = self.findings()
        self.assertIn("tools/Sync-SACCR-VBA.ps1: unsupported $sourceRoots entry: (Get-Item $env:SOURCE)", found)
        self.assertIn("tools/Sync-SACCR-VBA.ps1: $sourceRoots list is not closed", found)

    def test_required_files_must_exist(self) -> None:
        for path in ("tools/Sync-SACCR-VBA.ps1", "src/workbook/SACCR_Template.xlsx"):
            with self.subTest(path=path):
                self.findings()
                (self.root / path).unlink()
                self.assertIn(f"{path}: tracked but missing from the working tree",
                              check_source.run_check(self.root)["findings"])
                self.assertIn(f"{path}: required file is not tracked", self.findings())
                self.setUp()

    def test_document_modules_match_template_sheets(self) -> None:
        self.write_workbook("src/workbook/SACCR_Template.xlsx", {
            "[Content_Types].xml": "<Types/>",
            "_rels/.rels": '<Relationships Target="xl/workbook.xml"/>',
            "xl/workbook.xml": '<workbook><workbookPr codeName="ThisWorkbook"/></workbook>',
            "xl/worksheets/sheet1.xml": '<worksheet><sheetPr codeName="shInputs"/></worksheet>'})
        document = 'VERSION 1.0 CLASS\r\nAttribute VB_Name = "{0}"\r\nOption Explicit\r\n'
        self.write("src/workbook/ThisWorkbook.cls", document.format("ThisWorkbook"))
        self.write("src/workbook/shOutputs.cls", document.format("shOutputs"))
        found = self.findings()
        self.assertIn("src/workbook/shOutputs.cls: no sheet or workbook with this CodeName in "
                      "src/workbook/SACCR_Template.xlsx", found)
        self.assertIn("src/workbook/SACCR_Template.xlsx: CodeName shInputs has no document module "
                      "src/workbook/shInputs.cls", found)
        (self.root / "src/workbook/shOutputs.cls").unlink()
        self.write("src/workbook/shInputs.cls", document.format("shInputs"))
        self.assertEqual(self.findings(), [])

    def test_module_names_carry_their_role_prefix(self) -> None:
        for path, name in (("src/core/Engine.bas", "Engine"), ("src/modules/Formulas.bas", "Formulas"),
                           ("tests/modules/Harness.bas", "Harness")):
            self.write(path, MODULE.replace("SACCR_Test", name).replace(
                "Option Explicit", "Option Explicit\r\nOption Private Module"))
        found = self.findings()
        self.assertIn("src/core/Engine.bas: standard modules in src/core/ must be named CORE_<subject>", found)
        self.assertIn("src/modules/Formulas.bas: standard modules in src/modules/ must be named SACCR_<subject>", found)
        self.assertIn("tests/modules/Harness.bas: standard modules in tests/ must be named TEST_<subject>", found)

    def test_crlf_in_index_is_rejected(self) -> None:
        (self.root / ".gitattributes").write_text("* -text\n")
        self.write("notes.txt", "a\r\nb\r\n")
        found = self.findings()
        self.assertIn("notes.txt: stored with CRLF in Git; renormalize to LF", found)

    def test_lone_cr_in_declared_text_is_rejected(self) -> None:
        self.write("notes.md", "a\rb\r")
        (self.root / "blob.dat").write_bytes(b"a\rb\r")
        found = self.findings()
        self.assertIn("notes.md: declared text but Git classifies the blob as non-text (lone CR?)", found)
        self.assertFalse(any(f.startswith("blob.dat") for f in found))

    def test_vba_without_crlf_checkout_attribute_is_rejected(self) -> None:
        (self.root / ".gitattributes").write_text("* text=auto eol=lf\n")
        self.write("src/modules/SACCR_Test.bas", MODULE.replace("\r\n", "\n"))
        self.assertIn("src/modules/SACCR_Test.bas: .gitattributes must check VBA source out as CRLF", self.findings())

    def test_changelog_requires_unreleased_first(self) -> None:
        self.write("CHANGELOG.md", "## [0.0.1] - 2026-01-01\n\n## [Unreleased]\n")
        found = self.findings()
        self.assertIn("CHANGELOG.md: the first version heading must be '## [Unreleased]'", found)
        self.assertIn("CHANGELOG.md: missing link reference for [0.0.1]", found)

    def test_changelog_release_heading_format(self) -> None:
        self.write("CHANGELOG.md", CHANGELOG + "\n## [v0.0.1] 2026-01-01\n\n[v0.0.1]: x\n")
        self.assertTrue(any("release heading must be" in f for f in self.findings()))

    def test_version_file_tracks_newest_release(self) -> None:
        self.assertEqual(self.findings(), [])
        self.write("VERSION", "0.1.0\n")
        self.assertIn("VERSION exists but CHANGELOG.md has no release heading", self.findings())
        released = CHANGELOG.replace("[Unreleased]: ", "## [0.2.0] - 2026-02-01\n\n## [0.1.0] - 2026-01-01\n\n"
                                     "[0.2.0]: x\n[0.1.0]: x\n[Unreleased]: ")
        self.write("CHANGELOG.md", released)
        self.assertIn("VERSION 0.1.0 differs from the newest CHANGELOG.md release [0.2.0]", self.findings())
        self.write("VERSION", "v0.2.0\n")
        self.assertIn("VERSION must contain one X.Y.Z line ending in a newline", self.findings())
        self.write("VERSION", "0.2.0\n")
        self.assertEqual(self.findings(), [])
        (self.root / "VERSION").unlink()
        self.assertIn("VERSION is missing; CHANGELOG.md releases [0.2.0]", self.findings())

    def test_changelog_releases_newest_first(self) -> None:
        for first, second in (("1.0.0] - 2026-01-01", "2.0.0] - 2026-02-01"),
                              ("2.0.0] - 2026-01-01", "1.0.0] - 2026-02-01")):
            self.write("CHANGELOG.md", CHANGELOG.replace("[Unreleased]: ", f"## [{first}\n\n## [{second}\n\n"
                                                         "[1.0.0]: x\n[2.0.0]: x\n[Unreleased]: "))
            self.write("VERSION", first.split("]")[0] + "\n")
            self.assertIn("CHANGELOG.md: releases must be listed newest first, by version and date",
                          self.findings())

    def test_changelog_release_date_must_exist(self) -> None:
        for bad in ("2026-02-31", "2026-99-99"):
            self.write("CHANGELOG.md", CHANGELOG + f"\n## [0.0.1] - {bad}\n\n[0.0.1]: x\n")
            self.assertIn(f"CHANGELOG.md: [0.0.1] date {bad} is not a calendar date", self.findings())
        self.write("CHANGELOG.md", CHANGELOG + "\n## [0.0.1] - 2028-02-29\n\n[0.0.1]: x\n")
        self.write("VERSION", "0.0.1\n")
        self.assertEqual(self.findings(), [])

    def test_missing_changelog(self) -> None:
        (self.root / "CHANGELOG.md").unlink()
        self.assertIn("CHANGELOG.md is missing", self.findings())


class PublicApiRoleTests(unittest.TestCase):
    FACADE = ('Attribute VB_Name = "Facade"\nOption Explicit\n'
              'Public Function Echo(ByVal value As Long) As Long\nEnd Function\n')
    MANIFEST = ["Facade\tFunction\tEcho",
                "# SIG\tFacade\tFunction\tEcho\tPublic Function Echo(ByVal value As Long) As Long"]

    def test_facade_declaration_listed_passes(self) -> None:
        self.assertEqual(public_api.fixture(self.FACADE, self.MANIFEST)["status"], "pass")

    def test_unlisted_facade_declaration_fails(self) -> None:
        self.assertEqual(public_api.fixture(self.FACADE, [])["status"], "fail")

    def test_core_declaration_cannot_be_listed(self) -> None:
        manifest = self.MANIFEST + ["Core\tFunction\tInternalOnly"]
        self.assertEqual(public_api.fixture(self.FACADE, manifest)["status"], "fail")

    def test_roles_follow_repository_structure(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            subprocess.run(["git", "init", "-q", str(root)], check=True)
            for path in ("src/modules/A.bas", "src/core/B.bas", "src/classes/C.cls",
                         "tests/modules/D.bas", "examples/modules/E.bas", "misc/F.bas"):
                (root / path).parent.mkdir(parents=True, exist_ok=True)
                (root / path).write_text("x")
            subprocess.run(["git", "-C", str(root), "add", "-A"], check=True)
            self.assertEqual(public_api.component_roles(root), {
                "src/modules/A.bas": "public", "src/core/B.bas": "internal",
                "src/classes/C.cls": "internal", "tests/modules/D.bas": "test",
                "examples/modules/E.bas": "example"})


if __name__ == "__main__":
    unittest.main()
