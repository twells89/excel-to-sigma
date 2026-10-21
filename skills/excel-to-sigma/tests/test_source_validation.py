"""Source XML blank semantics are distinct from application route outcomes."""
import sys
import tempfile
import unittest
import xml.etree.ElementTree as ET
from pathlib import Path
from zipfile import ZipFile

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / 'scripts/lib'))
from source_validation import classify_blank, read_raw_validations, POLICY_VERSION
import presentation


class SourceBlankTests(unittest.TestCase):
    def test_typed_classification_and_raw_attempt(self):
        base = {'sourceType': 'decimal', 'operator': 'between', 'formula1': '0', 'formula2': '0.25',
                'rawAttributes': {'allowBlank': {'present': False, 'value': None}}}
        self.assertEqual(POLICY_VERSION, 2)
        self.assertEqual(classify_blank(base), 'SATISFIES_NUMERIC_RULE_AS_ZERO')
        self.assertEqual(classify_blank({**base, 'sourceType': 'whole', 'formula1': '1', 'formula2': '90'}),
                         'VIOLATES_NUMERIC_RULE_AS_ZERO')
        for kind in ('list', 'custom'):
            self.assertEqual(classify_blank({**base, 'sourceType': kind}), 'UNRESOLVED')
        for formula in ('A1', 'TRUE', '', True):
            self.assertEqual(classify_blank({**base, 'formula1': formula}), 'UNRESOLVED')
        self.assertEqual(classify_blank({**base, 'rawAttributes': {'allowBlank': {'present': True, 'value': '1'}}}),
                         'ALLOW_BLANK')
        self.assertEqual(classify_blank({**base, 'rawAttributes': {'allowBlank': {'present': True, 'value': '0'}}}),
                         'SATISFIES_NUMERIC_RULE_AS_ZERO')
        self.assertEqual(classify_blank({**base, 'operator': 'greaterThan', 'formula2': None}),
                         'VIOLATES_NUMERIC_RULE_AS_ZERO')
        self.assertEqual(classify_blank({**base, 'operator': 'lessThanOrEqual', 'formula2': None}),
                         'SATISFIES_NUMERIC_RULE_AS_ZERO')

    def test_fixture_xml_preserves_omission_namespace_part_and_hash(self):
        source = ROOT / 'tests/fixtures/fixture-research-model.xlsx'
        rows = read_raw_validations(source)
        self.assertEqual(len(rows), 1)
        row = rows[0]
        self.assertEqual(row['sheetPart'], 'xl/worksheets/sheet2.xml')
        self.assertEqual(row['namespace'], 'http://schemas.openxmlformats.org/spreadsheetml/2006/main')
        self.assertEqual(row['rawAttributes']['allowBlank'], {'present': False, 'value': None})
        self.assertEqual(row['rawAttributes']['showErrorMessage'], {'present': False, 'value': None})
        self.assertEqual(row['rawAttributes']['errorStyle'], {'present': False, 'value': None})
        self.assertEqual((row['formula1'], row['formula2']), ('0', '0.25'))
        self.assertEqual(len(row['ruleHash']), 64)
        self.assertEqual(len(row['sourceHash']), 64)

    def test_all_fifteen_fixture_rules_keep_authored_omission(self):
        files = sorted((ROOT / 'tests/fixtures').glob('fixture-*.xlsx'))
        rows = [row for source in files for row in read_raw_validations(source)]
        self.assertEqual(len(rows), 15)
        for row in rows:
            self.assertEqual(set(row['rawAttributes']), {'allowBlank', 'showErrorMessage', 'errorStyle'})
            self.assertTrue(all(value == {'present': False, 'value': None}
                                for value in row['rawAttributes'].values()))

    def test_explicit_true_false_and_alert_attributes_survive_xlsx_parser(self):
        original = ROOT / 'tests/fixtures/fixture-research-model.xlsx'
        for value, expected in (('1', 'ALLOW_BLANK'), ('0', 'SATISFIES_NUMERIC_RULE_AS_ZERO')):
            with self.subTest(value=value), tempfile.TemporaryDirectory() as directory:
                edited = Path(directory) / original.name
                with ZipFile(original) as source, ZipFile(edited, 'w') as destination:
                    for item in source.infolist():
                        data = source.read(item.filename)
                        if item.filename == 'xl/worksheets/sheet2.xml':
                            root = ET.fromstring(data)
                            element = root.find('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}dataValidation')
                            element.set('allowBlank', value)
                            element.set('showErrorMessage', '0')
                            element.set('errorStyle', 'warning')
                            data = ET.tostring(root, encoding='utf-8')
                        destination.writestr(item, data)
                rule = presentation.extract_presentation(edited)['validations'][0]
                self.assertEqual(rule['sourceXml']['rawAttributes']['allowBlank'],
                                 {'present': True, 'value': value})
                self.assertEqual(rule['sourceXml']['rawAttributes']['showErrorMessage'],
                                 {'present': True, 'value': '0'})
                self.assertEqual(rule['sourceXml']['rawAttributes']['errorStyle'],
                                 {'present': True, 'value': 'warning'})
                self.assertEqual(rule['blankSemantics']['classification'], expected)

    def test_referenced_numeric_bound_stays_unresolved(self):
        original = ROOT / 'tests/fixtures/fixture-research-model.xlsx'
        with tempfile.TemporaryDirectory() as directory:
            edited = Path(directory) / original.name
            with ZipFile(original) as source, ZipFile(edited, 'w') as destination:
                for item in source.infolist():
                    data = source.read(item.filename)
                    if item.filename == 'xl/worksheets/sheet2.xml':
                        root = ET.fromstring(data)
                        formula = root.find('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}dataValidation/{http://schemas.openxmlformats.org/spreadsheetml/2006/main}formula1')
                        formula.text = 'A1'
                        data = ET.tostring(root, encoding='utf-8')
                    destination.writestr(item, data)
            rule = presentation.extract_presentation(edited)['validations'][0]
            self.assertEqual(rule['status'], 'unresolved')
            self.assertIsNone(rule['emission'])
            self.assertEqual(rule['blankSemantics']['classification'], 'UNRESOLVED')

    def test_parser_keeps_blank_and_zero_distinct(self):
        source = ROOT / 'tests/fixtures/fixture-research-model.xlsx'
        rule = presentation.extract_presentation(source)['validations'][0]
        self.assertEqual(rule['blankSemantics']['classification'], 'SATISFIES_NUMERIC_RULE_AS_ZERO')
        self.assertEqual(rule['sourceXml']['rawAttributes']['allowBlank'], {'present': False, 'value': None})
        self.assertEqual(rule['sourceXml']['sheetPart'], 'xl/worksheets/sheet2.xml')
        self.assertIsNone(presentation.validation_accepts(rule, None))
        for value in ('', ' ', '0', 'text', False, True):
            self.assertFalse(presentation.validation_accepts(rule, value), repr(value))
        self.assertTrue(presentation.validation_accepts(rule, 0))

    def test_inventory_zero_outside_declared_range_but_route_unknown(self):
        source = ROOT / 'tests/fixtures/fixture-inventory-reorder.xlsx'
        rule = next(r for r in presentation.extract_presentation(source)['validations'] if r['sourceType'] == 'whole')
        self.assertEqual(rule['blankSemantics']['classification'], 'VIOLATES_NUMERIC_RULE_AS_ZERO')
        self.assertIsNone(presentation.validation_accepts(rule, None))
        self.assertFalse(presentation.validation_accepts(rule, 0))


if __name__ == '__main__':
    unittest.main()
