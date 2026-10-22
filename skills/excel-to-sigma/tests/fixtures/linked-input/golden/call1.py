# excel-to-sigma linked input table - call 1 of 3: base + one-row-per-key grouped table
# target: new workbook 'Rate override inputs'
old_pages = sigma.workbook.list_pages()
data_page = sigma.workbook.add_page('Data')
data_page.set_visibility(False)
page = sigma.workbook.add_page('Rate overrides')
dm = sigma.data.source.data_model('ff3d2a9b-dd11-4dfd-b059-7b9dff60d1ed')
members = [t for t in dm.list_tables() if any([c.name() == 'Default Rate' for c in t.columns()])]
if len(members) != 1:
    raise RuntimeError('expected exactly one data model member with column ' + 'Default Rate' + ', found ' + str(len(members)))
base = data_page.add_table(members[0])
base.set_name('Rate overrides baseline source')
base.set_frame(1, 1, 24, 12)
plan_base = data_page.add_table(base.to_source())
plan_base.set_name('Rate overrides baseline')
plan_base.set_frame(1, 13, 24, 12)
g = plan_base.set_groupings([plan_base.get_column_by_name('Line Item')])[0]
g.set_calculation('Baseline Rate', sigma.formula.max(plan_base.get_column_by_name('Default Rate')))
for p in old_pages:
    p.delete()
print('dataPageId=' + data_page.id())
print('pageId=' + page.id())
print('planBaseId=' + plan_base.id())
