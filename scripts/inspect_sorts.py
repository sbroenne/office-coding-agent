"""
Final transformation pass:
- Add sortMode to sort_pivot_field_labels/values
- Add filterType: 'manual' to apply_pivot_manual_filter
- Fix apply_pivot_label_filter (condition→labelCondition, value1→labelValue1, value2→labelValue2)
- clear_pivot_field_filters: remove existing filterType arg 
- Handle set_number_validation numberType→type conversion
- Handle custom_validation formula→customFormula
- Handle set_workbook_properties verify update
- Handle get_workbook_protection verify update
- Handle set_chart_legend_visibility arg renames
- Handle chart verify lambda updates
- Handle pivot verify lambda updates
- Remove dropped test blocks (5d, 5e, 5f, pivot 14, 19, 20, 21, 22-25)
- Add refresh_all_pivot_tables with pivotTableName
"""
import re

with open(r'D:\source\office-coding-agent\tests-e2e\src\test-taskpane.ts', 'r', encoding='utf-8') as f:
    content = f.read()

print(f'Input: {content.count(chr(10))} lines')

# ─── Fix sort_pivot_field_labels: add sortMode: 'labels' ───
# Pattern: action: 'sort',\n      pivotTableName: ...,\n      fieldName: ...,\n      sortBy: ...
# But NOT valuesHierarchyName - those are sort_pivot_field_values

# Looking at the sort calls, they now all have action: 'sort',
# sort_pivot_field_labels have sortBy directly after fieldName, no valuesHierarchyName
# sort_pivot_field_values have valuesHierarchyName

# For the label sort calls (no valuesHierarchyName before sort call close):
# PT_NAME version:
old_sort_label_1 = "{ action: 'sort',\n    { pivotTableName: PT_NAME, fieldName: 'Region', sortBy: 'Descending', sheetName: PIVOT_DST },"
# That pattern is wrong. Let me find the actual text.
idx = content.find("action: 'sort',\n    { pivotTableName: PT_NAME, fieldName: 'Region', sortBy: 'Descending',")
if idx >= 0:
    print(f'Found label sort PT_NAME at {idx}')
    print('Context:', repr(content[idx-20:idx+100]))

# Let me look at actual sort occurrences
for m in re.finditer(r"action: 'sort'[^}]+?\}", content, re.DOTALL):
    print('SORT MATCH:', repr(m.group()[:100]))

