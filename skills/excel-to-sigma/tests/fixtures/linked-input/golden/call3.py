# excel-to-sigma linked input table - call 3 of 3: seed / update rows by key
# target: existing workbook (the id call 1 returned)
page = sigma.workbook.get_page('<PAGE ID>')
if not isinstance(page, Page):
    raise RuntimeError('re-fetch page')
lt = page.get_element('<LINKED TABLE ID>')
if not isinstance(lt, LinkedInputTable):
    raise RuntimeError('expected linked input table')
k0 = lt.get_column_by_name('Line Item')
v0 = lt.get_column_by_name('Override Rate')
if not isinstance(v0, LiteralEditableColumn):
    raise RuntimeError('expected editable column: Override Rate')
v1 = lt.get_column_by_name('Rationale')
if not isinstance(v1, LiteralEditableColumn):
    raise RuntimeError('expected editable column: Rationale')
lt.update_rows([
    ({k0: 'Subscription - Enterprise'}, {v0: 0.045, v1: 'spike test'}),
    ({k0: 'Hosting'}, {v0: 0.025, v1: 'spike test'}),
])
print('updatedKeys=2')
