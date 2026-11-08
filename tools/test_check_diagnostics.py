"""Mutation coverage for missing, unknown, unused and retired diagnostic codes."""
from __future__ import annotations

import copy
import unittest

from check_diagnostics import CONFIG, check_sources


class DiagnosticTests(unittest.TestCase):
    def setUp(self) -> None:
        self.registry = {"codes": {"NS-FLAG": {"constant": "DIAG_NS_FLAG",
                         "status": "active", "description": "Y/N flag"}}}
        self.sources = {CONFIG: 'Public Const DIAG_NS_FLAG As String = "NS-FLAG"',
                        "engine.bas": 'Private Sub Validate()\n'
                        'LogMsg DIAG_NS_FLAG, "Margined", SEV_ERROR, "NS", "id", "bad"\nEnd Sub'}

    def test_valid_code_and_message_changes(self) -> None:
        self.assertEqual(check_sources(self.sources, self.registry), [])
        self.sources["engine.bas"] = self.sources["engine.bas"].replace('"bad"', '"new wording"')
        self.assertEqual(check_sources(self.sources, self.registry), [])

    def test_missing_and_unknown(self) -> None:
        for replacement in ("SEV_ERROR", "DIAG_UNKNOWN", '"NS-FLAG"'):
            with self.subTest(replacement=replacement):
                sources = dict(self.sources)
                sources["engine.bas"] = sources["engine.bas"].replace("DIAG_NS_FLAG", replacement)
                self.assertTrue(any("missing/unknown" in item for item in check_sources(sources, self.registry)))

    def test_comments_and_strings_are_not_usage(self) -> None:
        self.sources["engine.bas"] = '\' LogMsg DIAG_NS_FLAG\nvalue = "LogMsg DIAG_NS_FLAG"'
        self.assertEqual(check_sources(self.sources, self.registry), ["unused active code: DIAG_NS_FLAG"])

    def test_retired_kept_but_not_emitted(self) -> None:
        registry = copy.deepcopy(self.registry)
        registry["codes"]["NS-FLAG"]["status"] = "retired"
        self.assertTrue(any("retired code" in item for item in check_sources(self.sources, registry)))
        self.sources["engine.bas"] = ""
        self.assertEqual(check_sources(self.sources, registry), [])

    def test_registry_value_drift(self) -> None:
        self.sources[CONFIG] = self.sources[CONFIG].replace('"NS-FLAG"', '"NS-OTHER"')
        self.assertTrue(any("mismatch" in item for item in check_sources(self.sources, self.registry)))

    def test_continuation_inline_and_wrapper(self) -> None:
        self.sources["engine.bas"] = 'Private Sub Validate()\nIf bad Then LogMsg _\n DIAG_NS_FLAG, "field", SEV_ERROR\n'
        self.assertEqual(check_sources(self.sources, self.registry), [])
        self.sources["engine.bas"] += 'Private Sub AddErr()\nLogMsg DiagnosticCode, FieldName, SEV_ERROR\n'
        self.assertEqual(check_sources(self.sources, self.registry), [])


if __name__ == "__main__":
    unittest.main()
