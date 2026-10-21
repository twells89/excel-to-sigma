"""Portable exact-DDL receipt when Sigma REST omits inherited table names."""
import ast
import copy
import gzip
import json
from pathlib import Path
import sys
import unittest

SCRIPTS = Path(__file__).resolve().parents[1] / 'scripts'
sys.path[:0] = [str(SCRIPTS), str(SCRIPTS / 'lib')]
from lib.native_exact_plan import digest
from lib.regional_grouped_model import (build_grouped_model, prepare_product_error_patch,
    confirm_product_error_patch, prepare_product_grouping_program)

FIXTURES = Path(__file__).resolve().parent / 'fixtures' / 'regional-grouped'


class InheritedColumnReceiptTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        inputs = json.loads(gzip.decompress((FIXTURES / 'input-plan.json.gz').read_bytes()))
        native = json.loads(gzip.decompress((FIXTURES / 'native-plan.json.gz').read_bytes()))
        cls.model = build_grouped_model(inputs, native)

    def setUp(self):
        self.current = json.loads((FIXTURES / 'saved-b4-readback.json').read_text())
        self.current['documentVersion'] += 1
        self.parent = next(e for e in self.current['contents']['elements']
                           if e.get('name') == 'Regional Sales Rows')
        self.parent['columns'].extend([
            {'id': 'assigned-product-key', 'name': 'Product Key', 'formula': 'Lower([Sales/Product Line])'},
            {'id': 'assigned-month-key', 'name': 'Month Key', 'formula': '[Sales/Month]'}])
        labels = [c['name'] for c in self.parent['columns']]
        labels[-3] = 'First Error Ordinal (1)'
        self.labels = labels
        self.child = {'id': 'assigned-product-rows', 'kind': 'table',
                      'name': 'Regional Product Rows',
                      'source': {'kind': 'table', 'elementId': self.parent['id']},
                      'columns': [{'id': 'child-' + str(i),
                                   'formula': f'[{self.parent["name"]}/{label}]'}
                                  for i, label in enumerate(labels)] +
                                 [{'id': 'assigned-error', 'name': 'Product Error Ordinal',
                                   'formula': '0'}]}
        self.current['contents']['elements'].append(self.child)
        self.scaffold = {'status': 'PRODUCT_ROWS_SCAFFOLD_PENDING_EXECUTION',
                         'sourceSha256': self.model['sourceSha256'],
                         'documentId': self.current['workbookId'],
                         'rowElementId': self.parent['id']}

    def ddl(self, element, labels, formulas, types=None):
        if types is None:
            types = ['text'] * len(labels)
        lines = [f'-- Workbook ID: "{self.current["workbookId"]}"',
                 f'-- Element ID: "{element["id"]}"',
                 f'CREATE TABLE "workbook"."{element["id"]}" (']
        for index, (column, label, formula, kind) in enumerate(zip(
                element['columns'], labels, formulas, types)):
            comma = ',' if index < len(labels) - 1 else ''
            lines.append(f'  "{column["id"]}" {kind}{comma} -- "{label}" | Formula: {formula}')
        return '\n'.join(lines + [');']) + '\n'

    def receipt(self):
        parent_formulas = [c['formula'] for c in self.parent['columns']]
        child_labels = self.labels + ['Product Error Ordinal']
        child_formulas = [c['formula'] for c in self.child['columns']]
        return {'documentId': self.current['workbookId'],
                'documentVersion': self.current['documentVersion'],
                'contentsHash': digest(self.current['contents']),
                'rowElementId': self.parent['id'], 'productElementId': self.child['id'],
                'parentDdl': self.ddl(self.parent, self.labels, parent_formulas),
                'productDdl': self.ddl(self.child, child_labels, child_formulas,
                                   ['text'] * len(self.labels) + ['number']),
                'reviewedParentNameDifferences': [{
                    'id': self.parent['columns'][-3]['id'],
                    'restName': 'First Error Ordinal',
                    'sdkName': 'First Error Ordinal (1)'}]}

    def patch(self, receipt=None):
        return prepare_product_error_patch(self.model, self.current, self.scaffold,
            row_element_id=self.parent['id'], product_element_id=self.child['id'],
            error_column_id='assigned-error', inherited_column_receipt=receipt)

    def test_strict_default_blocks_unnamed_and_complete_receipt_patches_only_error(self):
        with self.assertRaisesRegex(ValueError, 'BLOCKED'):
            self.patch()
        packet = self.patch(self.receipt())
        self.assertEqual(packet['ifMatchDocumentVersion'], self.current['documentVersion'])
        self.assertEqual(packet['contents']['elements'][-1]['columns'][-1]['formula'],
                         'If([Revenue Error] != "", [Source Ordinal], Null)')
        restored = copy.deepcopy(packet['contents'])
        restored['elements'][-1]['columns'][-1]['formula'] = '0'
        self.assertEqual(restored, self.current['contents'])

    def test_foreign_stale_incomplete_and_lineage_drift_rejected(self):
        base = self.receipt()
        variants = []
        for key, value in (('documentId', 'foreign'),
                           ('documentVersion', base['documentVersion'] + 1),
                           ('contentsHash', '0' * 64),
                           ('rowElementId', 'wrong-row'),
                           ('productElementId', 'wrong-product')):
            x = copy.deepcopy(base)
            x[key] = value
            variants.append(x)
        x = copy.deepcopy(base)
        x['productDdl'] = x['productDdl'].replace('"child-0"', '"wrong-id"')
        variants.append(x)
        x = copy.deepcopy(base)
        x['productDdl'] = x['productDdl'].replace('"child-0" text', '"child-0" number')
        variants.append(x)
        x = copy.deepcopy(base)
        x['productDdl'] = x['productDdl'].replace('[Regional Sales Rows/Group Key]', '0')
        variants.append(x)
        x = copy.deepcopy(base)
        x['reviewedParentNameDifferences'] = []
        variants.append(x)
        x = copy.deepcopy(base)
        x['productDdl'] = '\n'.join(x['productDdl'].splitlines()[:-2])
        variants.append(x)
        for index, value in enumerate(variants):
            with self.subTest(index=index), self.assertRaisesRegex(ValueError, 'BLOCKED'):
                self.patch(value)
        named = copy.deepcopy(self.current)
        child = named['contents']['elements'][-1]
        for col, label in zip(child['columns'][:-1], self.labels):
            col['name'] = label
        receipt = self.receipt()
        with self.assertRaisesRegex(ValueError, 'BLOCKED'):
            prepare_product_error_patch(self.model, named, self.scaffold,
                row_element_id=self.parent['id'], product_element_id=self.child['id'],
                error_column_id='assigned-error', inherited_column_receipt=receipt)

    def test_error_placeholder_type_and_complete_ddl_required(self):
        base = self.receipt()
        variants = []
        x = copy.deepcopy(base)
        x['productDdl'] = x['productDdl'].replace('"assigned-error" number', '"assigned-error" text')
        variants.append(x)
        x = copy.deepcopy(base)
        x['productDdl'] = x['productDdl'].replace('\n);', '\n  "hidden" object -- "Hidden" | Formula: 0\n);')
        variants.append(x)
        x = copy.deepcopy(base)
        x['productDdl'] += '\nCREATE TABLE "workbook"."foreign" ();'
        variants.append(x)
        for index, receipt in enumerate(variants):
            with self.subTest(index=index), self.assertRaisesRegex(ValueError, 'BLOCKED'):
                self.patch(receipt)

    def test_complete_receipt_and_column_order_required(self):
        base = self.receipt()
        variants = []
        for key in base:
            x = copy.deepcopy(base)
            x.pop(key)
            variants.append(x)
        x = copy.deepcopy(base)
        x['documentVersion'] = True
        variants.append(x)
        x = copy.deepcopy(base)
        lines = x['productDdl'].splitlines()
        lines[3], lines[4] = lines[4], lines[3]
        x['productDdl'] = '\n'.join(lines)
        variants.append(x)
        x = copy.deepcopy(base)
        x['productDdl'] = x['productDdl'].replace('"child-0"', '"child-1"')
        variants.append(x)
        for index, receipt in enumerate(variants):
            with self.subTest(index=index), self.assertRaisesRegex(ValueError, 'BLOCKED'):
                self.patch(receipt)

    def test_default_named_path_unchanged_and_receipt_is_specific_to_unnamed(self):
        named = copy.deepcopy(self.current)
        for col, name in zip(named['contents']['elements'][-1]['columns'][:-1], self.labels):
            col['name'] = name
        packet = prepare_product_error_patch(self.model, named, self.scaffold,
            row_element_id=self.parent['id'], product_element_id=self.child['id'],
            error_column_id='assigned-error')
        restored = copy.deepcopy(packet['contents'])
        restored['elements'][-1]['columns'][-1]['formula'] = '0'
        self.assertEqual(restored, named['contents'])
        receipt = self.receipt()
        receipt['contentsHash'] = digest(named['contents'])
        with self.assertRaisesRegex(ValueError, 'BLOCKED'):
            prepare_product_error_patch(self.model, named, self.scaffold,
                row_element_id=self.parent['id'], product_element_id=self.child['id'],
                error_column_id='assigned-error', inherited_column_receipt=receipt)

    def grouping_state(self):
        patch = self.patch(self.receipt())
        saved = copy.deepcopy(self.current)
        saved['documentVersion'] += 1
        saved['contents'] = patch['contents']
        confirmation = confirm_product_error_patch(patch, saved)
        self.current = saved
        self.parent = next(e for e in saved['contents']['elements'] if e['id'] == self.parent['id'])
        self.child = next(e for e in saved['contents']['elements'] if e['id'] == self.child['id'])
        ddl_receipt = self.receipt()
        sdk = {'documentId': saved['workbookId'], 'documentVersion': saved['documentVersion'],
               'contentsHash': digest(saved['contents']), 'elementId': self.child['id'],
               'columns': [{'id': c['id'], 'name': name, 'formula': c['formula']}
                           for c, name in zip(self.child['columns'],
                                              self.labels + ['Product Error Ordinal'])],
               'reviewedNameDifferences': []}
        return confirmation, ddl_receipt, sdk

    def group(self, confirmation, ddl_receipt=None, sdk_receipt=None):
        return prepare_product_grouping_program(self.model, self.current, confirmation,
            product_element_id=self.child['id'],
            page_id=self.current['contents']['pages'][0]['id'], grid_row=130,
            inherited_column_receipt=ddl_receipt, sdk_metadata_receipt=sdk_receipt)

    def test_grouping_uses_complete_saved_formula_and_actual_sdk_projection(self):
        confirmation, ddl_receipt, sdk = self.grouping_state()
        with self.assertRaisesRegex(ValueError, 'BLOCKED'):
            self.group(confirmation)
        with self.assertRaisesRegex(ValueError, 'BLOCKED'):
            self.group(confirmation, ddl_receipt)
        packet = self.group(confirmation, ddl_receipt, sdk)
        tree = ast.parse(packet['program'])
        metadata = next(n.test.comparators[0] for n in tree.body if isinstance(n, ast.If)
                        and isinstance(n.test, ast.Compare) and isinstance(n.test.left, ast.Name)
                        and n.test.left.id == 'actual')
        self.assertEqual(ast.literal_eval(metadata), sdk['columns'])
        for column in ['child-20', 'child-21', 'child-8', 'assigned-error']:
            self.assertIn('get_column("' + column + '")', packet['program'])
        self.assertEqual(packet['contentsHash'], digest(self.current['contents']))

    def test_grouping_rejects_stale_projection_or_ddl_or_confirmation(self):
        confirmation, ddl_receipt, sdk = self.grouping_state()
        variants = []
        for key, value in [('documentId', 'foreign'), ('documentVersion', sdk['documentVersion'] - 1),
                           ('contentsHash', '0' * 64), ('elementId', 'foreign')]:
            x = copy.deepcopy(sdk)
            x[key] = value
            variants.append(x)
        x = copy.deepcopy(sdk)
        x['columns'].reverse()
        variants.append(x)
        x = copy.deepcopy(sdk)
        x['columns'][-1]['formula'] = '0'
        variants.append(x)
        x = copy.deepcopy(sdk)
        x['columns'][0]['name'] = 'unreviewed'
        variants.append(x)
        for index, receipt in enumerate(variants):
            with self.subTest(index=index), self.assertRaisesRegex(ValueError, 'BLOCKED'):
                self.group(confirmation, ddl_receipt, receipt)
        wrong = copy.deepcopy(ddl_receipt)
        wrong['productDdl'] = wrong['productDdl'].replace(
            'Formula: If([Revenue Error] != "", [Source Ordinal], Null)', 'Formula: 0')
        with self.assertRaisesRegex(ValueError, 'BLOCKED'):
            self.group(confirmation, wrong, sdk)
        wrong = copy.deepcopy(confirmation)
        wrong['errorColumnId'] = 'foreign'
        with self.assertRaisesRegex(ValueError, 'BLOCKED'):
            self.group(wrong, ddl_receipt, sdk)


if __name__ == '__main__':
    unittest.main()
