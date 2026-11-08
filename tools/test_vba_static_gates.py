"""Valid/defective twins for the #538 lexical and static-analysis contracts."""
from __future__ import annotations

import contextlib
import io
import json
import tempfile
import unittest
from collections.abc import Callable
from pathlib import Path
from typing import Any

import check_source
import check_vba_calls as calls
import check_vba_conditionals as conditionals
import check_vba_dependencies as dependencies
import check_vba_identifiers as identifiers
import check_vba_jumps as jumps
import check_vba_public_api as public_api
from vba_lexer import logical_lines, non_ascii_outside_strings, strip_vba, tokenize


def module(name: str, body: str) -> str:
    return f'Attribute VB_Name = "{name}"\nOption Explicit\n{body}\n'


class LexicalTests(unittest.TestCase):
    def test_strings_comments_continuations_and_named_arguments(self) -> None:
        source = ['Call F("a:""b", _', ' x:=1) \' ignored']
        self.assertEqual(logical_lines(source), [(1, 2, 'Call F("a:""b", x:=1)')])
        self.assertEqual(strip_vba(' x = "a\'b" \' comment'), ' x = "a\'b"')
        self.assertEqual(strip_vba('Rem GoTo missing'), '')
        self.assertEqual(jumps.strip_vba('GoTo Done \' "bad"'), 'GoTo Done')
        self.assertEqual(public_api.split_statements('x = #12:30:00#: F x:=1'),
                         ['x = #12:30:00#', 'F x:=1'])
        self.assertEqual([t.text for t in tokenize('F .Name, F.Name')],
                         ['F', '.', 'Name', ',', 'F', '.', 'Name'])
        self.assertTrue(tokenize('F .Name')[1].spaced)
        self.assertFalse(tokenize('F.Name')[1].spaced)

    def test_existing_gate_fixture_results_remain_unchanged(self) -> None:
        for checker in (jumps, conditionals, public_api):
            with self.subTest(checker=checker.__name__), contextlib.redirect_stdout(io.StringIO()):
                self.assertEqual(checker.run_self_test(), 0)

    def test_ascii_literal_and_comment_twins(self) -> None:
        self.assertEqual(non_ascii_outside_strings('x = "caf\u00e9"\n\' cafe'), [])
        self.assertEqual(non_ascii_outside_strings('x = "a""caf\u00e9"'), [])
        for source in ('If True Then Rem "caf\u00e9"', 'If False Then x = 1 Else Rem "caf\u00e9"', '\' "caf\u00e9"', 'Rem "caf\u00e9"', '10 Rem "caf\u00e9"',
                       'x=1: Rem "caf\u00e9"', '\' continued _\n"caf\u00e9"', 'caf\u00e9 = 1'):
            with self.subTest(source=source):
                self.assertEqual(len(non_ascii_outside_strings(source)), 1)


    def test_bracketed_sheet_quotes_do_not_start_comments(self) -> None:
        source = "Debug.Print ['Sheet 1'!A1], \"caf\u00e9\""
        self.assertEqual(non_ascii_outside_strings(source), [])
        self.assertEqual(len(non_ascii_outside_strings(source + " ' caf\u00e9")), 1)
        source = "Debug.Print ['Rem Then Else'!A1], \"caf\u00e9\""
        self.assertEqual(non_ascii_outside_strings(source), [])

    def test_source_gate_enforces_ascii_outside_literals(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            path = "src/core/CORE_A.bas"
            (root / path).parent.mkdir(parents=True)
            source = module("CORE_A", 'Option Private Module\nPrivate Const Text As String = "caf\u00e9"')
            (root / path).write_bytes(source.encode("cp1252"))
            self.assertEqual(check_source.check_component(root, path, {path}, {}), [])
            (root / path).write_bytes((source + "' caf\u00e9\n").encode("cp1252"))
            findings = check_source.check_component(root, path, {path}, {})
            self.assertEqual(len(findings), 1)
            self.assertIn("non-ASCII outside string literal", findings[0])


class IdentifierTests(unittest.TestCase):
    def test_declared_and_misspelled_variable_twins(self) -> None:
        valid = module('CORE_A', '''Private saved As Long
Private Sub Run(ByVal count As Long)
Dim value As Long
value = count + saved + Len("ignored misspelling")
Debug.Print value
End Sub''')
        self.assertEqual(identifiers.analyze({'src/core/CORE_A.bas': valid})['findings'], [])
        invalid = valid.replace('value = count', 'value = coutn')
        report = identifiers.analyze({'src/core/CORE_A.bas': invalid})
        self.assertEqual(report['status'], 'pass')  # advisory
        self.assertEqual({f['identifier'] for f in report['findings']}, {'coutn'})

    def test_member_names_labels_named_arguments_and_comments_are_not_variables(self) -> None:
        source = module('CORE_A', '''Sub Run(ByVal obj As Object, ByVal index As Long)
Dim result As Variant
On Error GoTo Failed
result = obj.UnknownMember(index)
obj.Other arg:=index
With obj
.Other named:=index
End With
GoTo Done
Failed: Resume Next
Done:
' ignore undefined
Rem ignore too
End Sub''')
        self.assertEqual(identifiers.analyze({'src/core/CORE_A.bas': source})['findings'], [])

    def test_public_project_name_resolves_but_foreign_private_does_not(self) -> None:
        sources = {'src/core/CORE_A.bas': module('CORE_A', 'Sub Run()\nDebug.Print SharedValue\nEnd Sub'),
                   'src/core/CORE_B.bas': module('CORE_B', 'Public Const SharedValue As Long = 4')}
        self.assertEqual(identifiers.analyze(sources)['findings'], [])
        sources['src/core/CORE_B.bas'] = sources['src/core/CORE_B.bas'].replace('Public Const', 'Private Const')
        self.assertEqual({f['identifier'] for f in identifiers.analyze(sources)['findings']}, {'SharedValue'})

    def test_conditional_variants_do_not_hide_missing_declarations(self) -> None:
        source = module('CORE_A', '''Sub Run()
#If Win64 Then
Dim only64 As Long
#End If
Debug.Print only64
End Sub''')
        findings = identifiers.analyze({'src/core/CORE_A.bas': source})['findings']
        self.assertEqual({f['environment'] for f in findings}, {'vba6-win32', 'vba7-win32'})

    def test_parameter_types_and_defaults_are_checked(self) -> None:
        source = module('CORE_A', 'Sub Run(Optional x As MissingType = MissingDefault)\nEnd Sub')
        self.assertEqual({f['identifier'] for f in identifiers.analyze({'src/core/CORE_A.bas': source})['findings']},
                         {'MissingType', 'MissingDefault'})

    def test_unknown_host_name_is_not_exempt_by_prefix(self) -> None:
        source = module('CORE_A', 'Sub Run()\nDebug.Print xlTypo, vbTypo\nEnd Sub')
        self.assertEqual({f['identifier'] for f in identifiers.analyze({'src/core/CORE_A.bas': source})['findings']},
                         {'xlTypo', 'vbTypo'})


class DependencyTests(unittest.TestCase):
    def test_allowed_and_inverted_layer_twins(self) -> None:
        sources = {'src/core/CORE_A.bas': module('CORE_A', 'Public Sub Work()\nEnd Sub'),
                   'src/modules/SACCR_A.bas': module('SACCR_A', 'Sub Run()\nCORE_A.Work\nEnd Sub')}
        self.assertEqual(dependencies.analyze(sources, [])['status'], 'pass')
        sources['src/core/CORE_A.bas'] = module('CORE_A', 'Sub Work()\nSACCR_A.Run\nEnd Sub')
        self.assertEqual(dependencies.analyze(sources, [])['status'], 'fail')
        exception = {'source': 'src/core/CORE_A.bas', 'target': 'src/modules/SACCR_A.bas'}
        self.assertEqual(dependencies.analyze(sources, [exception])['status'], 'pass')
        sources['src/workbook/M_Main.bas'] = module('M_Main', 'Sub Other()\nCORE_A.Work\nEnd Sub')
        self.assertEqual(dependencies.analyze(sources, [exception])['status'], 'fail')

    def test_unqualified_reference_and_type_are_checked(self) -> None:
        for call in ('Call Run', 'Dim x As FacadeType'):
            with self.subTest(call=call):
                sources = {'src/core/CORE_A.bas': module('CORE_A', f'Sub Work()\n{call}\nEnd Sub'),
                           'src/modules/SACCR_A.bas': module('SACCR_A',
                           'Public Type FacadeType\nx As Long\nEnd Type\nPublic Sub Run()\nEnd Sub')}
                self.assertEqual(dependencies.analyze(sources, [])['status'], 'fail')

    def test_host_access_and_shadowed_name_twins(self) -> None:
        for code, expected in (('Dim r As Range', 'fail'), ('Application.Calculate', 'fail'),
                               ('MsgBox "hello"', 'fail'), ('SendKeys "x"', 'fail'), ('DoEvents', 'fail'),
                               ('Dim x As String: x = InputBox("prompt")', 'fail'),
                               ('Excel.Application.Calculate', 'fail'),
                               ('Dim Range As Range', 'fail'),
                               ('Dim columns As Long\ncolumns = 1', 'pass'),
                               ('Debug.Print "Range, Application.Calculate"', 'pass')):
            with self.subTest(code=code):
                source = module('CORE_A', f'Sub Work()\n{code}\nEnd Sub')
                self.assertEqual(dependencies.analyze({'src/core/CORE_A.bas': source}, [])['status'], expected)

    def test_udt_field_names_are_not_host_roots_but_types_are(self) -> None:
        for field, expected in (('Range(1 To 2) As Long', 'pass'),
                                ('Columns As Long', 'pass'), ('Items(1 To 2) As Worksheet', 'fail'),
                                ('Worksheet As Worksheet', 'fail'),
                                ('Worksheet(1 To 2) As Worksheet', 'fail')):
            with self.subTest(field=field):
                source = module('CORE_A', f'Public Type Data\n{field}\nEnd Type')
                self.assertEqual(dependencies.analyze({'src/core/CORE_A.bas': source}, [])['status'], expected)

    def test_production_to_tests_and_examples_to_core_fail(self) -> None:
        sources = {'src/core/CORE_A.bas': module('CORE_A', 'Sub Work()\nTEST_A.Help\nEnd Sub'),
                   'tests/modules/TEST_A.bas': module('TEST_A', 'Public Sub Help()\nEnd Sub')}
        self.assertEqual(dependencies.analyze(sources, [])['status'], 'fail')
        sources['src/core/CORE_A.bas'] = module('CORE_A', 'Sub Work()\nEnd Sub')
        self.assertEqual(dependencies.analyze(sources, [])['status'], 'pass')
        sources['examples/modules/EX_A.bas'] = module('EX_A', 'Sub Example()\nCORE_A.Work\nEnd Sub')
        self.assertEqual(dependencies.analyze(sources, [])['status'], 'fail')

    def test_conditional_errors_fail_closed(self) -> None:
        source = module('CORE_A', '#If UnknownFeature Then\nSub Work()\nEnd Sub\n#End If')
        checkers: tuple[Callable[[], dict[str, Any]], ...] = (lambda: dependencies.analyze({'src/core/CORE_A.bas': source}, []),
                        lambda: identifiers.analyze({'src/core/CORE_A.bas': source}))
        for checker in checkers:
            self.assertEqual(checker()['status'], 'fail')

    def test_rem_comments_cannot_create_references_or_declarations(self) -> None:
        source = module('CORE_A', '''Sub Work()
Dim n As Long
n = 1: Rem old code: Application.Calculate: Dim hidden As Long
10 Rem also: Excel.Application.Calculate
Done: Rem label comment: UnknownMember
Rem continued _
Application.Calculate
Debug.Print hidden
End Sub''')
        files = {'src/core/CORE_A.bas': source}
        self.assertEqual(dependencies.analyze(files, [])['status'], 'pass')
        self.assertEqual({f['identifier'] for f in identifiers.analyze(files)['findings']}, {'hidden'})

    def test_native_declaration_parameters_are_not_host_roots(self) -> None:
        for name in ('columns', 'worksheet', 'range'):
            source = module('CORE_A', f'Private Declare Sub Work Lib "example" (ByVal {name} As Long)')
            self.assertEqual(dependencies.analyze({'src/core/CORE_A.bas': source}, [])['status'], 'pass')
        source = module('CORE_A', 'Private Declare Sub Work Lib "example" (ByVal x As Worksheet)')
        self.assertEqual(dependencies.analyze({'src/core/CORE_A.bas': source}, [])['status'], 'fail')

    def test_event_parameter_names_are_not_host_dependencies(self) -> None:
        for type_name, expected in (('Long', 'pass'), ('Worksheet', 'fail')):
            source = module('CORE_A', f'Public Event Changed(ByVal columns As {type_name})')
            self.assertEqual(dependencies.analyze({'src/core/CORE_A.cls': source}, [])['status'], expected)

    def test_project_symbols_take_precedence_over_host_fallback(self) -> None:
        for name in ('Columns', 'Range', 'Worksheet'):
            sources = {'src/core/CORE_A.bas': module('CORE_A', f'Public Sub {name}()\nEnd Sub'),
                       'src/core/CORE_B.bas': module('CORE_B', f'Sub Run()\n{name}\nEnd Sub')}
            self.assertEqual(dependencies.analyze(sources, [])['status'], 'pass')
            sources['src/core/CORE_B.bas'] = module('CORE_B', f'Sub Run()\nDim x As {name}\nEnd Sub')
            self.assertEqual(dependencies.analyze(sources, [])['status'], 'fail')
        source = module('CORE_A', 'Private Type Worksheet\nx As Long\nEnd Type\nSub Run()\nDim x As Worksheet\nEnd Sub')
        self.assertEqual(dependencies.analyze({'src/core/CORE_A.bas': source}, [])['status'], 'pass')

    def test_public_object_module_enums_have_dependency_owners(self) -> None:
        source = module('ThisWorkbook', 'Public Enum DocumentState\nReady = 1\nEnd Enum')
        for use in ('Dim x As DocumentState', 'Dim x As Long: x = Ready'):
            sources = {'src/workbook/ThisWorkbook.cls': source,
                       'src/core/CORE_A.bas': module('CORE_A', f'Sub Run()\n{use}\nEnd Sub')}
            result = dependencies.analyze(sources, [])
            self.assertEqual(result['status'], 'fail')
            self.assertEqual({f['target'] for f in result['findings']}, {'src/workbook/ThisWorkbook.cls'})
            sources['tests/modules/TEST_A.bas'] = sources.pop('src/core/CORE_A.bas')
            self.assertEqual(dependencies.analyze(sources, [])['status'], 'pass')

    def test_quoted_sheet_references_cannot_hide_later_code(self) -> None:
        source = module('CORE_A', '''Sub Run()
Debug.Print ['Sheet 1'!A1]: Application.Calculate
Dim Evaluate As Long
Debug.Print ['Sheet 1'!A1:B2]: unknownVariable = 1
End Sub''')
        files = {'src/core/CORE_A.bas': source}
        findings = dependencies.analyze(files, [])['findings']
        self.assertEqual({f['target'] for f in findings}, {'host:application', 'host:evaluate'})
        self.assertEqual({f['identifier'] for f in identifiers.analyze(files)['findings']}, {'unknownVariable'})

    def test_parentheses_inside_sheet_names_are_opaque(self) -> None:
        for sheet in ('Sheet 1', 'Sheet (1', 'Sheet )1'):
            source = module('SACCR_A', f"Sub Run()\nDebug.Print ['{sheet}'!A1]: End Sub")
            files = {'src/modules/SACCR_A.bas': source}
            self.assertEqual(dependencies.analyze(files, [])['status'], 'pass')
            self.assertEqual(identifiers.analyze(files)['analysis_failures'], [])
            self.assertEqual(identifiers.analyze(files)['findings'], [])

    def test_qualified_ui_calls_fail_but_pure_runtime_calls_pass(self) -> None:
        for call, expected in (('VBA.MsgBox "hello"', 'fail'),
                               ('VBA.Interaction.MsgBox "hello"', 'fail'),
                               ('Interaction.MsgBox "hello"', 'fail'),
                               ('VBA.Interaction.InputBox$("prompt")', 'fail'),
                               ('Debug.Print VBA.Len("safe")', 'pass'),
                               ('Debug.Print VBA.Strings.Len("safe")', 'pass'),
                               ('Dim VBA As Object: VBA.MsgBox "late-bound"', 'pass')):
            with self.subTest(call=call):
                source = module('CORE_A', f'Sub Run()\n{call}\nEnd Sub')
                self.assertEqual(dependencies.analyze({'src/core/CORE_A.bas': source}, [])['status'], expected)

    def test_current_project_qualification_cannot_hide_dependencies(self) -> None:
        facade = module('SACCR_A', 'Public Type FacadeType\nx As Long\nEnd Type\nPublic Sub Run()\nEnd Sub')
        for use in ('VBAProject.SACCR_A.Run', 'VBAProject.Run', 'Dim x As VBAProject.FacadeType'):
            sources = {'src/modules/SACCR_A.bas': facade,
                       'src/core/CORE_A.bas': module('CORE_A', f'Sub Work()\n{use}\nEnd Sub')}
            self.assertEqual(dependencies.analyze(sources, [])['status'], 'fail')
            self.assertEqual(identifiers.analyze(sources)['findings'], [])
            sources['tests/modules/TEST_A.bas'] = sources.pop('src/core/CORE_A.bas')
            self.assertEqual(dependencies.analyze(sources, [])['status'], 'pass')

    def test_exception_manifest_rejects_blanket_exemptions(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            (root / 'tools').mkdir()
            entry = {'source': 'src/core/*', 'target': 'host:*', 'issue': '#220', 'reason': 'legacy'}
            (root / dependencies.POLICY).write_text(json.dumps({'schema_version': 1, 'exceptions': [entry]}))
            with self.assertRaises(ValueError):
                dependencies.read_exceptions(root)

    def test_calls_use_the_shared_tokenizer(self) -> None:
        self.assertIs(calls.tokenize, tokenize)


if __name__ == '__main__':
    unittest.main()
