"""Tests of the reconciliation output sheets of the workbook template (#720).

The two sheets are never protected, carry an AutoFilter on their header row, show the amount
figures in one number format with a format of their own for the factor rows, and name the
difference column an absolute difference. Above the table they show summary tiles: rounded-rectangle
shapes whose values are linked to helper cells in hidden columns, the formulas of which follow the filter.
"""

from __future__ import annotations

import re
import unittest
import zipfile
from pathlib import Path
from xml.etree import ElementTree

from template_layout import formula_text, named_cell
from test_template_buttons import sheet_parts

ROOT = Path(__file__).resolve().parents[1]
TEMPLATE = ROOT / "src" / "workbook" / "SACCR_Template.xlsx"
MAIN = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
# Sheet: (last table column, the measures shown in the factor format)
SHEETS = {
    "Reconciliation": ("N", {"Multiplier"}),
    "TradeReconciliation": ("O", {"M", "Supervisory duration", "Supervisory delta", "MF unmargined", "MF margined"}),
}
FACTOR_FORMAT = "#,##0.000000"
LAST_SHEET_ROW = 1048576
HEADER, FIRST = 11, 12  # the header row and the first data row: four rows of tiles and a spacer sit above the header
XDR = "{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}"
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
CONFIG = (ROOT / "src" / "core" / "CORE_Config.bas").read_text(encoding="cp1252")
# Tile: (shape, label, linked cell, result counted by the cell or None, fill of the value)
TILES = (
    ("tile_LastReconciliation", "Last reconciliation", "R3", None, "DDEBF7"),
    ("tile_RowsShown", "Rows shown", "R2", None, "DDEBF7"),
    ("tile_FiguresCompared", "Figures compared", "T1", None, "DDEBF7"),
    ("tile_Match", "Match", "R4", "MATCH", "C6EFCE"),
    ("tile_Difference", "Difference", "R5", "DIFFERENCE", "FFC7CE"),
    ("tile_NotInBank", "Not in bank file", "T2", "NOT IN BANK FILE", "DDEBF7"),
    ("tile_MissingInCalculator", "Missing in calculator", "T3", "MISSING IN CALCULATOR", "DDEBF7"),
    ("tile_Duplicate", "Duplicate rows", "T4", "DUPLICATE", "DDEBF7"),
    ("tile_OtherResults", "Other results", "T5", None, "DDEBF7"),
    ("tile_Measure", "Measure shown", "X4", None, "DDEBF7"),
    ("tile_SumBank", "Sum of bank figures", "X5", None, "DDEBF7"),
    ("tile_SumCalc", "Sum of calc. figures", "Z1", None, "DDEBF7"),
    ("tile_SumAbsDiff", "Sum of abs. diff.", "Z2", None, "DDEBF7"),
    ("tile_MaxAbsDiff", "Largest abs. diff.", "Z3", None, "DDEBF7"),
)
# The measures of each sheet that are factors, not amounts: no sum is shown for them.
FACTORS = {"Reconciliation": ("Multiplier",),
           "TradeReconciliation": ("M", "Supervisory duration", "Supervisory delta", "MF unmargined", "MF margined")}
# The results the Other results tile counts, one helper cell each.
OTHERS = {"V1": "WITHHELD", "V2": "INVALID INPUT", "V3": "REGIME MISMATCH", "V4": "EXCLUDED", "V5": "NOT APPLICABLE"}


class ReconciliationSheetTests(unittest.TestCase):
    def setUp(self) -> None:
        self.package = zipfile.ZipFile(TEMPLATE)
        self.addCleanup(self.package.close)
        self.parts = sheet_parts(self.package)

    def sheet(self, name: str) -> ElementTree.Element:
        return ElementTree.fromstring(self.package.read(self.parts[name]))

    def test_the_sheets_are_saved_unprotected(self) -> None:
        for name in SHEETS:
            self.assertIsNone(self.sheet(name).find(MAIN + "sheetProtection"), name)

    def test_the_header_row_has_an_autofilter(self) -> None:
        for name, (last, _) in SHEETS.items():
            autofilter = self.sheet(name).find(MAIN + "autoFilter")
            self.assertIsNotNone(autofilter, name)
            assert autofilter is not None
            self.assertEqual(autofilter.attrib["ref"], f"C{HEADER}:{last}100000", name)
            workbook = self.package.read("xl/workbook.xml").decode("utf-8")
            self.assertRegex(workbook, rf'<definedName name="_xlnm\._FilterDatabase" localSheetId="\d+" hidden="1">'
                                       rf"{name}!\$C\${HEADER}:\${last}\$100000</definedName>")

    def test_the_factor_rows_have_their_own_format(self) -> None:
        styles = ElementTree.fromstring(self.package.read("xl/styles.xml"))
        dxfs = styles.find(MAIN + "dxfs")
        assert dxfs is not None
        for name, (_, measures) in SHEETS.items():
            rules = [rule for block in self.sheet(name).findall(MAIN + "conditionalFormatting")
                     if block.attrib["sqref"] == f"I{FIRST}:K1048576" for rule in block.findall(MAIN + "cfRule")]
            self.assertEqual(len(rules), 1, name)
            formula = rules[0].findtext(MAIN + "formula") or ""
            self.assertEqual(set(re.findall(rf'\$H{FIRST}="([^"]+)"', formula)), measures, name)
            dxf = list(dxfs)[int(rules[0].attrib["dxfId"])]
            numfmt = dxf.find(MAIN + "numFmt")
            assert numfmt is not None
            self.assertEqual(numfmt.attrib["formatCode"], FACTOR_FORMAT, name)

    def test_the_difference_column_is_an_absolute_difference(self) -> None:
        strings = self.package.read("xl/sharedStrings.xml").decode("utf-8")
        self.assertIn("<t>Absolute difference</t>", strings)
        self.assertNotIn("calculator - bank", strings)

    def test_the_macros_write_the_absolute_difference_and_one_amount_format(self) -> None:
        reconcile = (ROOT / "src" / "core" / "CORE_Reconcile.bas").read_text(encoding="utf-8")
        self.assertIn('"", Abs(Difference), RelativeDifference', reconcile)
        main = (ROOT / "src" / "workbook" / "M_Main.bas").read_text(encoding="utf-8")
        self.assertIn('"@|0|@|@|@|@|#,##0.00|#,##0.00|"', main)
        self.assertNotIn("0.00########", main)

    def test_a_longer_reconciliation_gets_a_filter_over_all_its_rows(self) -> None:
        config = (ROOT / "src" / "core" / "CORE_Config.bas").read_text(encoding="utf-8")
        match = re.search(r"RECON_FILTER_LAST_ROW\s+As Long = (\d+)", config)
        self.assertIsNotNone(match)
        assert match is not None
        self.assertEqual(int(match.group(1)), 100000)
        for name in SHEETS:
            autofilter = self.sheet(name).find(MAIN + "autoFilter")
            assert autofilter is not None
            self.assertTrue(autofilter.attrib["ref"].endswith(match.group(1)), name)
        main = (ROOT / "src" / "workbook" / "M_Main.bas").read_text(encoding="utf-8")
        self.assertIn("> RECON_FILTER_LAST_ROW Then", main)
        self.assertNotIn("FormatFactorRows", main)

    def drawing(self, name: str) -> ElementTree.Element:
        rels = self.package.read(self.parts[name].replace("worksheets/", "worksheets/_rels/") + ".rels").decode("utf-8")
        match = re.search(r'Target="\.\./drawings/(drawing\d+\.xml)"', rels)
        assert match is not None
        return ElementTree.fromstring(self.package.read("xl/drawings/" + match.group(1)))

    def helper_formulas(self, name: str) -> dict[str, str]:
        """The formula of each helper cell of rows 1 to 5, columns R, T, V, X and Z."""
        out: dict[str, str] = {}
        for cell in self.sheet(name).iter(MAIN + "c"):
            if re.fullmatch(r"[RTVXZ][1-5]", cell.attrib["r"]):
                out[cell.attrib["r"]] = cell.findtext(MAIN + "f") or ""
        return out

    def test_the_header_row_moved_down_to_make_room_for_the_tiles(self) -> None:
        workbook = self.package.read("xl/workbook.xml").decode("utf-8")
        for name in SHEETS:
            self.assertIn(f'<definedName name="{name}_Header">{name}!$C${HEADER}</definedName>', workbook)
            pane = self.sheet(name).find(f"{MAIN}sheetViews/{MAIN}sheetView/{MAIN}pane")
            assert pane is not None
            self.assertEqual((pane.attrib["ySplit"], pane.attrib["topLeftCell"]), (str(HEADER), f"D{FIRST}"), name)
            rows = {row.attrib["r"]: row for row in self.sheet(name).iter(MAIN + "row")}
            self.assertEqual([rows[str(r)].attrib["ht"] for r in (6, 7, 8, 9, 10)], ["22.5", "22.5", "22.5", "22.5", "6"], name)
            self.assertIsNotNone(rows[str(HEADER)].find(f'{MAIN}c[@r="C{HEADER}"]'), name)

    def test_every_tile_is_a_shape_linked_to_a_helper_cell(self) -> None:
        for name in SHEETS:
            groups = [anchor for anchor in self.drawing(name).iter(XDR + "oneCellAnchor")
                      if anchor.find(f"{XDR}grpSp/{XDR}nvGrpSpPr/{XDR}cNvPr") is not None
                      and anchor.find(f"{XDR}grpSp/{XDR}nvGrpSpPr/{XDR}cNvPr").attrib["name"] == "tile_Group"]  # type: ignore[union-attr]
            self.assertEqual(len(groups), 1, name)
            anchor = groups[0]
            # Right of the frozen columns, from the first tile row, never moved or resized with the cells.
            frozen = int(re.search(r'<pane xSplit="(\d+)"', self.package.read(self.parts[name]).decode("utf-8")).group(1))  # type: ignore[union-attr]
            self.assertGreaterEqual(int(anchor.findtext(f"{XDR}from/{XDR}col") or "0"), frozen, name)
            self.assertEqual(anchor.findtext(f"{XDR}from/{XDR}row"), "5", name)
            self.assertEqual(anchor.find(XDR + "clientData").attrib.get("fLocksWithSheet", "1"), "1", name)  # type: ignore[union-attr]  # 1 is the default
            shapes = {item.find(f"{XDR}nvSpPr/{XDR}cNvPr").attrib["name"]: item  # type: ignore[union-attr]
                      for item in anchor.iter(XDR + "sp")}
            self.assertEqual(len(shapes), 2 * len(TILES), name)
            for tile, label, cell, _, fill in TILES:
                value, caption = shapes[tile], shapes[tile + "_Label"]
                self.assertEqual("".join(item.text or "" for item in caption.iter(A + "t")), label, (name, tile))
                self.assertEqual(caption.attrib["textlink"], "", (name, tile))
                self.assertEqual(value.attrib["textlink"], f"${cell[0]}${cell[1:]}", (name, tile))
                for shape in (value, caption):
                    geometry = shape.find(f"{XDR}spPr/{A}prstGeom")
                    assert geometry is not None
                    self.assertEqual(geometry.attrib["prst"], "roundRect", (name, tile))
                    self.assertEqual(shape.attrib["macro"], "", (name, tile))
                color = value.find(f"{XDR}spPr/{A}solidFill/{A}srgbClr")
                assert color is not None
                self.assertEqual(color.attrib["val"], fill, (name, tile))
            ids = [item.attrib["id"] for item in self.drawing(name).iter(XDR + "cNvPr")]
            self.assertEqual(len(ids), len(set(ids)), name)

    def test_the_helper_cells_follow_the_filter_and_count_the_results_of_the_calculator(self) -> None:
        results = {value for value in re.findall(r'RES_\w+\s+As String = "([^"]+)"', CONFIG)}
        for name in SHEETS:
            formulas = self.helper_formulas(name)
            self.assertEqual(set(formulas), {"R1", "R2", "R3", "R4", "R5", "T1", "T2", "T3", "T4", "T5", "X1", "X2", "X3", "X4",
                                             "X5", "Z1", "Z2", "Z3"} | set(OTHERS), name)
            # To the last row of the sheet, not to the last row of the AutoFilter: a reconciliation of more rows
            # than RECON_FILTER_LAST_ROW gets a filter over all of them, and the tiles must count them too.
            self.assertEqual(formulas["R1"], f"COUNTA($M${FIRST}:$M${LAST_SHEET_ROW})", name)
            self.assertEqual(formulas["R2"], f"SUBTOTAL(103,$M${FIRST}:$M${LAST_SHEET_ROW})", name)
            self.assertEqual(formulas["T1"], "R4+R5", name)
            for cell, result in ({tile[2]: tile[3] for tile in TILES if tile[3]} | OTHERS).items():
                if cell == "T1":
                    continue
                self.assertIn(result, results, (name, cell))
                # SUBTOTAL(103) of the visible rows, one by one, times the rows of that result; none when the table is empty.
                self.assertTrue(formulas[cell].startswith("IF($R$1=0,0,SUMPRODUCT(SUBTOTAL(103,OFFSET("), (name, cell))
                self.assertTrue(formulas[cell].endswith(f'="{result}")))'), (name, cell))
            self.assertIn('LEFT($C$4,20)="Last reconciliation "', formulas["R3"])
            self.assertIn('"None"', formulas["T5"])
            workbook = self.package.read("xl/workbook.xml").decode("utf-8")
            self.assertIn(f'<definedName name="{name}_Tiles">{name}!$Q$1:$Z$5</definedName>', workbook)

    def test_the_figures_of_one_measure_follow_the_filter_and_are_not_summed_across_measures(self) -> None:
        for name in SHEETS:
            formulas = self.helper_formulas(name)
            visible = f"SUBTOTAL(103,OFFSET($M${FIRST},ROW(OFFSET($M${FIRST},0,0,$R$1,1))-ROW($M${FIRST}),0))"
            # The first visible measure, and whether the filter leaves that measure alone.
            self.assertIn(f"MATCH(1,{visible},0)", formulas["X1"], name)
            self.assertIn(f"SUMPRODUCT({visible}*(OFFSET($H${FIRST},0,0,$R$1,1)=$X$1))=$R$2", formulas["X2"], name)
            for measure in FACTORS[name]:
                self.assertIn(f'$X$1="{measure}"', formulas["X3"], (name, measure))
            self.assertEqual(formulas["X3"].count("$X$1="), len(FACTORS[name]), name)
            # A sum is shown for one amount measure only; with several, a prompt; for a factor, none.
            for cell, column in (("X5", "I"), ("Z1", "J"), ("Z2", "K")):
                self.assertIn("IF($X$2=0,", formulas[cell], (name, cell))
                self.assertIn('IF($X$3=1,"Factor",FIXED(SUMPRODUCT(', formulas[cell], (name, cell))
                self.assertIn(f"*OFFSET(${column}${FIRST},0,0,$R$1,1)),2)", formulas[cell], (name, cell))
            # The largest absolute difference is shown for a factor too, with six decimals instead of two.
            self.assertIn(f"SUMPRODUCT(MAX({visible}*(OFFSET($H${FIRST},0,0,$R$1,1)=$X$1)*OFFSET($K${FIRST},0,0,$R$1,1)))", formulas["Z3"], name)
            self.assertIn("IF($X$3=1,6,2)", formulas["Z3"], name)
            self.assertIn('"Filter one measure"', formulas["X5"], name)

    def test_the_helper_cells_sit_in_hidden_columns_beyond_the_table(self) -> None:
        for name, (last, _) in SHEETS.items():
            hidden = [(int(col.attrib["min"]), int(col.attrib["max"])) for col in self.sheet(name).iter(MAIN + "col")
                      if col.attrib.get("hidden") == "1"]
            self.assertEqual(hidden, [(17, 26)], name)  # Q to Z; the table ends at N or O
            self.assertLess(ord(last), ord("Q"), name)

    def test_the_sheets_are_listed_as_never_protected(self) -> None:
        config = (ROOT / "src" / "core" / "CORE_Config.bas").read_text(encoding="utf-8")
        match = re.search(r'PROTECT_NONE\s+As String = "([^"]*)"', config)
        self.assertIsNotNone(match)
        assert match is not None
        self.assertEqual({name for name in match.group(1).split("|") if name}, set(SHEETS))


class ReconSummaryTests(unittest.TestCase):
    """The ReconSummary sheet: one row per measure of each reconciliation sheet, following the filters (#720)."""

    FIRST_ROWS = {"Reconciliation": 8, "TradeReconciliation": 22}
    COLUMNS = "CDEFGHIJKL"

    def setUp(self) -> None:
        self.package = zipfile.ZipFile(TEMPLATE)
        self.addCleanup(self.package.close)
        self.parts = sheet_parts(self.package)
        self.strings = [re.sub(r"<[^>]+>", "", item) for item in
                        re.findall(r"<si>(.*?)</si>", self.package.read("xl/sharedStrings.xml").decode("utf-8"), re.S)]
        self.sheet = ElementTree.fromstring(self.package.read(self.parts["ReconSummary"]))

    def cell(self, ref: str) -> ElementTree.Element:
        found = self.sheet.find(f'.//{MAIN}c[@r="{ref}"]')
        assert found is not None, ref
        return found

    def text(self, ref: str) -> str:
        value = self.cell(ref).findtext(MAIN + "v") or ""
        return self.strings[int(value)] if self.cell(ref).attrib.get("t") == "s" else value

    def formula(self, ref: str) -> str:
        return formula_text(self.sheet, self.cell(ref))

    def source_measures(self, constant: str) -> list[str]:
        reconcile = (ROOT / "src" / "core" / "CORE_Reconcile.bas").read_text(encoding="utf-8")
        match = re.search(rf"{constant}\s+As String = ((?:\"[^\"]*\"(?: & _\s*)?)+)", reconcile)
        assert match is not None
        return "".join(re.findall(r'"([^"]*)"', match.group(1))).split("|")

    def test_the_sheet_follows_the_two_reconciliation_sheets_in_tab_order(self) -> None:
        names = re.findall(r'<sheet name="([^"]+)"', self.package.read("xl/workbook.xml").decode("utf-8"))
        self.assertEqual(names[names.index("TradeReconciliation") + 1], "ReconSummary")

    def test_it_lists_the_measures_of_the_reconciliation_code_in_its_order(self) -> None:
        for sheet, constant in (("Reconciliation", "NS_MEASURES"), ("TradeReconciliation", "TR_MEASURES")):
            first = self.FIRST_ROWS[sheet]
            measures = self.source_measures(constant)
            self.assertEqual(len(measures), 11)
            self.assertEqual([self.text(f"C{first + i}") for i in range(len(measures))], measures, sheet)

    def test_the_formulas_read_the_visible_rows_of_the_right_sheet(self) -> None:
        for sheet, first in self.FIRST_ROWS.items():
            for row in range(first, first + 11):
                visible = (f"SUBTOTAL(103,OFFSET({sheet}!$M${FIRST},ROW(OFFSET({sheet}!$M${FIRST},0,0,{sheet}!$R$1,1))"
                           f"-ROW({sheet}!$M${FIRST}),0))")
                measure = f"(OFFSET({sheet}!$H${FIRST},0,0,{sheet}!$R$1,1)=$C{row})"
                self.assertIn(f"IF({sheet}!$R$1=0,0,SUMPRODUCT({visible}*{measure}", self.formula(f"D{row}"), row)
                self.assertIn(f'*(OFFSET({sheet}!$M${FIRST},0,0,{sheet}!$R$1,1)="MATCH")', self.formula(f"E{row}"), row)
                self.assertIn('="DIFFERENCE")', self.formula(f"F{row}"), row)
                self.assertEqual(self.formula(f"G{row}"), f"D{row}-E{row}-F{row}", row)
                # Sums and differences count only the rows compared, so that the two sums cover the same rows.
                compared = (f'((OFFSET({sheet}!$M${FIRST},0,0,{sheet}!$R$1,1)="MATCH")'
                            f'+(OFFSET({sheet}!$M${FIRST},0,0,{sheet}!$R$1,1)="DIFFERENCE"))')
                self.assertIn(f"SUMPRODUCT(MAX({visible}*{measure}*{compared}*OFFSET({sheet}!$K${FIRST},0,0,"
                              f"{sheet}!$R$1,1)))", self.formula(f"K{row}"), row)
                self.assertTrue(self.formula(f"K{row}").startswith(f'IF(E{row}+F{row}=0,"",'), row)
                self.assertIn(f"*{measure}*{compared}*", self.formula(f"L{row}"), row)
                self.assertIn(f"+{FIRST - 1}", self.formula(f"L{row}"), row)

    def test_amounts_have_sums_and_factors_do_not(self) -> None:
        for sheet, first in self.FIRST_ROWS.items():
            for i, measure in enumerate(self.source_measures("NS_MEASURES" if sheet == "Reconciliation" else "TR_MEASURES")):
                row = first + i
                if measure in FACTORS[sheet]:
                    for column in "HIJ":
                        self.assertEqual(self.text(f"{column}{row}"), "-", (sheet, measure))
                        self.assertEqual(self.formula(f"{column}{row}"), "", (sheet, measure))
                else:
                    for column, source in zip("HIJ", "IJK"):
                        self.assertIn(f'="DIFFERENCE"))*OFFSET({sheet}!${source}${FIRST},0,0,{sheet}!$R$1,1)',
                                      self.formula(f"{column}{row}"), (sheet, measure))
                        self.assertTrue(self.formula(f"{column}{row}").startswith(f'IF(E{row}+F{row}=0,"",'),
                                        (sheet, measure))

    def test_home_links_to_every_other_sheet_in_tab_order(self) -> None:
        names = re.findall(r'<sheet name="([^"]+)"', self.package.read("xl/workbook.xml").decode("utf-8"))
        home = self.package.read(self.parts["Home"]).decode("utf-8")
        links = re.findall(r'<hyperlink ref="C(\d+)" location="\'([^\']+)\'!A1" display="([^"]+)"', home)
        self.assertEqual([name for _, name, _ in links], names[1:])
        self.assertTrue(all(name == shown for _, name, shown in links))
        rows = [int(row) for row, _, _ in links]
        self.assertEqual(rows, list(range(rows[0], rows[0] + len(rows))))

    def test_the_sheet_is_saved_protected_as_an_output_sheet(self) -> None:
        protection = self.sheet.find(MAIN + "sheetProtection")
        assert protection is not None
        self.assertEqual(protection.attrib.get("sheet"), "1")
        for cell in self.sheet.iter(MAIN + "c"):  # formulas only: no cell is unlocked
            if cell.find(MAIN + "f") is not None:
                self.assertTrue(formula_text(self.sheet, cell).strip(), cell.attrib.get("r"))
        styles = ElementTree.fromstring(self.package.read("xl/styles.xml"))
        xfs = list(styles.find(MAIN + "cellXfs"))  # type: ignore[arg-type]
        for cell in self.sheet.iter(MAIN + "c"):
            protection_of = xfs[int(cell.attrib.get("s", "0"))].find(MAIN + "protection")
            self.assertTrue(protection_of is None or protection_of.attrib.get("locked") != "0", cell.attrib["r"])


class SampleBankFigureTests(unittest.TestCase):
    """The template carries sample bank figures, so that a first run shows a reconciliation (#720)."""

    def setUp(self) -> None:
        self.package = zipfile.ZipFile(TEMPLATE)
        self.addCleanup(self.package.close)
        self.parts = sheet_parts(self.package)
        self.strings = [re.sub(r"<[^>]+>", "", item) for item in
                        re.findall(r"<si>(.*?)</si>", self.package.read("xl/sharedStrings.xml").decode("utf-8"), re.S)]

    def column(self, sheet: str, letter: str, first: int, last: int) -> list[str]:
        root = ElementTree.fromstring(self.package.read(self.parts[sheet]))
        out: list[str] = []
        for row in range(first, last + 1):
            cell = root.find(f'.//{MAIN}c[@r="{letter}{row}"]')
            value = "" if cell is None else cell.findtext(MAIN + "v") or ""
            out.append(self.strings[int(value)] if cell is not None and cell.attrib.get("t") == "s" and value else value)
        return out

    def first_row(self, name: str) -> int:
        """The first data row of a sheet: the row below the header its workbook name locates."""
        return named_cell(self.package, name).row + 1

    def test_the_bank_sheets_ship_without_figures(self) -> None:
        # The figures of the sample portfolio spoiled the reconciliation of a real one (#815): the
        # template ships both bank tables empty, below their header, with their settings.
        for sheet, header, last in (("BankNettingSets", "BankNettingSets_Header", "P"),
                                    ("BankTrades", "BankTrades_Header", "O")):
            first = self.first_row(header)
            for letter in "CDEFGHIJKLMNOP"[: "CDEFGHIJKLMNOP".index(last) + 1]:
                self.assertEqual([v for v in self.column(sheet, letter, first, first + 60) if v], [],
                                 f"{sheet} column {letter}")
            self.assertEqual(self.column(sheet, "D", 4, 6), ["1", "1", "0.01"], sheet)

if __name__ == "__main__":
    unittest.main()
