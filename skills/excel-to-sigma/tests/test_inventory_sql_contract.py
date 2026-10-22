"""Offline SQL contract probe. This does not compile or execute Snowflake SQL."""
import re
import unittest
from pathlib import Path

SQL=(Path(__file__).resolve().parents[1]/
    'refs/inventory-isolated-integrations.sql').read_text()


def _match_count_expression():
    query=re.search(r'SELECT\s+COUNT\(\*\),\s*(.*?)\s+INTO\s+:V_PRIOR,\s*:V_MATCH',
        SQL,re.I|re.S)
    if not query:
        raise AssertionError('exact prior/match aggregate query missing')
    return query.group(1)


def _count_if_snowflake(predicate_values):
    # Snowflake COUNT_IF is NULL when no record satisfies the condition.
    matches=sum(value is True for value in predicate_values)
    return matches if matches else None


def _sql_not_equal_one(value):
    # Three-valued SQL logic: NULL <> 1 is UNKNOWN, not TRUE.
    return None if value is None else value != 1


def _retry_branch(existing, *, sku, quantity, batch):
    """Independent single-row branch model selected by the actual SQL aggregate."""
    expression=_match_count_expression()
    prior=[row for row in existing if row['operation']=='intent-P' and row['source_row']==4]
    other_batch=any(row['operation']=='intent-P' and row['batch']!=batch for row in existing)
    counted=_count_if_snowflake([
        row['sku']==sku and row['quantity']==quantity and row['batch']==batch
        for row in prior])
    if re.search(r'COALESCE\s*\(\s*COUNT_IF\s*\(',expression,re.I):
        counted=0 if counted is None else counted
    mismatch=(len(prior)==1 and _sql_not_equal_one(counted))
    # Snowflake Scripting IF treats UNKNOWN as false.
    if other_batch or len(prior)>1 or mismatch is True:
        return 'FAILURE'
    return 'RETRY_PRESENT' if len(prior)==1 else 'STAGED'


class InventorySqlContractTests(unittest.TestCase):
    def test_null_source_row_and_fingerprint_are_explicit_invalid_arguments(self):
        header=SQL.split('BEGIN TRANSACTION;',1)[0]
        self.assertRegex(header,r'(?i)P_SOURCE_ROW\s+IS\s+NULL')
        self.assertRegex(header,r'(?i)P_BATCH_FINGERPRINT\s+IS\s+NULL')

    def test_zero_match_aggregate_is_explicitly_normalized(self):
        expression=_match_count_expression()
        self.assertRegex(expression,r'(?is)COALESCE\s*\(\s*COUNT_IF\s*\(')
        self.assertRegex(expression,r'(?is),\s*0\s*\)\s*$')

    def test_same_intent_changed_sku_or_quantity_cannot_return_retry_present(self):
        old=[{'operation':'intent-P','source_row':4,'sku':'SKU-200',
              'quantity':25,'batch':'a'*64}]
        self.assertEqual(_retry_branch(old,sku='SKU-200',quantity=25,batch='a'*64),'RETRY_PRESENT')
        self.assertEqual(_retry_branch(old,sku='SKU-999',quantity=25,batch='a'*64),'FAILURE')
        self.assertEqual(_retry_branch(old,sku='SKU-200',quantity=50,batch='a'*64),'FAILURE')
        self.assertEqual(_retry_branch([],sku='SKU-200',quantity=25,batch='a'*64),'STAGED')

    def test_duplicate_or_other_batch_fails_even_with_matching_row(self):
        old={'operation':'intent-P','source_row':4,'sku':'SKU-200',
              'quantity':25,'batch':'a'*64}
        self.assertEqual(_retry_branch([old,old],sku='SKU-200',quantity=25,batch='a'*64),'FAILURE')
        self.assertEqual(_retry_branch([old],sku='SKU-200',quantity=25,batch='b'*64),'FAILURE')


if __name__=='__main__':
    unittest.main()
