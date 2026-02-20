"""
Fix remaining multi-line tool invocations in test-taskpane.ts.
Handles the pattern:
    configs,
    'old_name',
    { args
"""
import re

with open(r'D:\source\office-coding-agent\tests-e2e\src\test-taskpane.ts', 'r', encoding='utf-8') as f:
    content = f.read()

print(f'Input: {content.count(chr(10))} lines')

def replace_multiline_tool(text, old_configs, old_name, new_configs, new_name, new_action):
    """Replace multi-line: 
       old_configs,\n    'old_name',\n    { args
    → new_configs,\n    'new_name',\n    { action: 'new_action', args
    Also handles single-line variant.
    """
    # Multi-line pattern (may have varying indentation)
    # Use regex to handle different indentation levels
    pattern = re.compile(
        re.escape(old_configs) + r',\n(\s+)' + re.escape("'" + old_name + "'") + r',\n(\s+)\{',
        re.MULTILINE
    )
    def repl(m):
        indent1 = m.group(1)
        indent2 = m.group(2)
        if new_action:
            return f"{new_configs},\n{indent1}'{new_name}',\n{indent2}{{ action: '{new_action}',"
        else:
            return f"{new_configs},\n{indent1}'{new_name}',\n{indent2}{{"
    new_text, cnt = pattern.subn(repl, text)
    return new_text, cnt

# ─── Range tools ───
range_multiline_map = [
    ('set_range_values', 'set_values'),
    ('sort_range', 'sort'),
    ('delete_range', 'delete'),
    ('insert_range', 'insert'),
    ('replace_values', 'replace'),
    ('remove_duplicates', 'remove_duplicates'),
    ('set_range_formulas', 'set_formulas'),
    ('auto_fill_range', 'fill'),
]
for old, new_action in range_multiline_map:
    content, cnt = replace_multiline_tool(content, 'rangeConfigs', old, 'rangeConfigs', 'range', new_action)
    if cnt: print(f'  range/{old} → {new_action}: {cnt}')

# Range format tools (multi-line)
range_format_multiline = [
    ('format_range', 'format'),
    ('set_number_format', 'set_number_format'),
    ('set_cell_borders', 'set_borders'),
    ('set_hyperlink', 'set_hyperlink'),
]
for old, new_action in range_format_multiline:
    content, cnt = replace_multiline_tool(content, 'rangeConfigs', old, 'rangeFormatConfigs', 'range_format', new_action)
    if cnt: print(f'  range_format/{old} → {new_action}: {cnt}')

# auto_fit_columns / auto_fit_rows multi-line
pattern_cols = re.compile(r"rangeConfigs,\n(\s+)'auto_fit_columns',\n(\s+)\{", re.MULTILINE)
def repl_cols(m):
    return f"rangeFormatConfigs,\n{m.group(1)}'range_format',\n{m.group(2)}{{ action: 'auto_fit', fitTarget: 'columns',"
content, cnt = pattern_cols.subn(repl_cols, content)
print(f'  auto_fit_columns multi: {cnt}')

pattern_rows = re.compile(r"rangeConfigs,\n(\s+)'auto_fit_rows',\n(\s+)\{", re.MULTILINE)
def repl_rows(m):
    return f"rangeFormatConfigs,\n{m.group(1)}'range_format',\n{m.group(2)}{{ action: 'auto_fit', fitTarget: 'rows',"
content, cnt = pattern_rows.subn(repl_rows, content)
print(f'  auto_fit_rows multi: {cnt}')

# ─── Table tools ───
table_multiline = [
    ('add_table_rows', 'add_rows'),
    ('sort_table', 'sort'),
    ('filter_table', 'filter'),
    ('add_table_column', 'add_column'),
    ('delete_table_column', 'delete_column'),
    ('set_table_style', 'configure'),
    ('set_table_header_totals_visibility', 'configure'),
]
for old, new_action in table_multiline:
    content, cnt = replace_multiline_tool(content, 'tableConfigs', old, 'tableConfigs', 'table', new_action)
    if cnt: print(f'  table/{old} → {new_action}: {cnt}')

# ─── Chart tools ───
chart_multiline = [
    ('create_chart', 'create'),
    ('set_chart_title', 'configure'),
    ('set_chart_type', 'configure'),
    ('set_chart_data_source', 'configure'),
    ('set_chart_position', 'configure'),
    ('set_chart_legend_visibility', 'configure'),
    ('set_chart_axis_title', 'configure'),
    ('set_chart_axis_visibility', 'configure'),
    ('set_chart_series_filtered', 'configure'),
]
for old, new_action in chart_multiline:
    content, cnt = replace_multiline_tool(content, 'chartConfigs', old, 'chartConfigs', 'chart', new_action)
    if cnt: print(f'  chart/{old} → {new_action}: {cnt}')

# ─── Sheet tools ───
sheet_multiline = [
    ('rename_sheet', 'rename'),
    ('set_sheet_gridlines', 'set_gridlines'),
    ('set_sheet_headings', 'set_headings'),
    ('set_page_layout', 'set_page_layout'),
]
for old, new_action in sheet_multiline:
    content, cnt = replace_multiline_tool(content, 'sheetConfigs', old, 'sheetConfigs', 'sheet', new_action)
    if cnt: print(f'  sheet/{old} → {new_action}: {cnt}')

# ─── Workbook tools ───
workbook_multiline = [
    ('set_workbook_properties', 'set_properties'),
    ('define_named_range', 'define_named_range'),
]
for old, new_action in workbook_multiline:
    content, cnt = replace_multiline_tool(content, 'workbookConfigs', old, 'workbookConfigs', 'workbook', new_action)
    if cnt: print(f'  workbook/{old} → {new_action}: {cnt}')

# ─── Comment tools ───
comment_multiline = [
    ('add_comment', 'add'),
    ('edit_comment', 'edit'),
]
for old, new_action in comment_multiline:
    content, cnt = replace_multiline_tool(content, 'commentConfigs', old, 'commentConfigs', 'comment', new_action)
    if cnt: print(f'  comment/{old} → {new_action}: {cnt}')

# ─── Conditional format tools ───
cf_multiline = [
    ('list_conditional_formats', 'list'),
    ('clear_conditional_formats', 'clear'),
]
for old, new_action in cf_multiline:
    content, cnt = replace_multiline_tool(content, 'conditionalFormatConfigs', old, 'conditionalFormatConfigs', 'conditional_format', new_action)
    if cnt: print(f'  cf/{old} → {new_action}: {cnt}')

# CF add with type (multi-line)
cf_add_multiline = [
    ('add_color_scale', 'colorScale'),
    ('add_data_bar', 'dataBar'),
    ('add_cell_value_format', 'cellValue'),
    ('add_top_bottom_format', 'topBottom'),
    ('add_contains_text_format', 'containsText'),
    ('add_custom_format', 'custom'),
]
for old, cf_type in cf_add_multiline:
    pattern = re.compile(
        r"conditionalFormatConfigs,\n(\s+)'" + re.escape(old) + r"',\n(\s+)\{",
        re.MULTILINE
    )
    def make_repl(t):
        def repl(m):
            return f"conditionalFormatConfigs,\n{m.group(1)}'conditional_format',\n{m.group(2)}{{ action: 'add', type: '{t}',"
        return repl
    content, cnt = pattern.subn(make_repl(cf_type), content)
    if cnt: print(f'  cf/{old} → add/{cf_type}: {cnt}')

# ─── Data validation tools ───
dv_multiline = [
    ('get_data_validation', 'get'),
    ('clear_data_validation', 'clear'),
]
for old, new_action in dv_multiline:
    content, cnt = replace_multiline_tool(content, 'dataValidationConfigs', old, 'dataValidationConfigs', 'data_validation', new_action)
    if cnt: print(f'  dv/{old} → {new_action}: {cnt}')

# set_list_validation (multi-line with listValues already set by first script)
pattern_list = re.compile(r"dataValidationConfigs,\n(\s+)'set_list_validation',\n(\s+)\{", re.MULTILINE)
def repl_list(m):
    return f"dataValidationConfigs,\n{m.group(1)}'data_validation',\n{m.group(2)}{{ action: 'set', type: 'list',"
content, cnt = pattern_list.subn(repl_list, content)
print(f'  set_list_validation: {cnt}')

# set_date_validation
pattern_date = re.compile(r"dataValidationConfigs,\n(\s+)'set_date_validation',\n(\s+)\{", re.MULTILINE)
def repl_date(m):
    return f"dataValidationConfigs,\n{m.group(1)}'data_validation',\n{m.group(2)}{{ action: 'set', type: 'date',"
content, cnt = pattern_date.subn(repl_date, content)
print(f'  set_date_validation: {cnt}')

# set_text_length_validation
pattern_text = re.compile(r"dataValidationConfigs,\n(\s+)'set_text_length_validation',\n(\s+)\{", re.MULTILINE)
def repl_text(m):
    return f"dataValidationConfigs,\n{m.group(1)}'data_validation',\n{m.group(2)}{{ action: 'set', type: 'textLength',"
content, cnt = pattern_text.subn(repl_text, content)
print(f'  set_text_length_validation: {cnt}')

# set_custom_validation
pattern_custom = re.compile(r"dataValidationConfigs,\n(\s+)'set_custom_validation',\n(\s+)\{", re.MULTILINE)
def repl_custom(m):
    return f"dataValidationConfigs,\n{m.group(1)}'data_validation',\n{m.group(2)}{{ action: 'set', type: 'custom',"
content, cnt = pattern_custom.subn(repl_custom, content)
print(f'  set_custom_validation: {cnt}')

# set_number_validation - no type yet, handle numberType transformation below
pattern_num = re.compile(r"dataValidationConfigs,\n(\s+)'set_number_validation',\n(\s+)\{", re.MULTILINE)
def repl_num(m):
    return f"dataValidationConfigs,\n{m.group(1)}'data_validation',\n{m.group(2)}{{ action: 'set',"
content, cnt = pattern_num.subn(repl_num, content)
print(f'  set_number_validation: {cnt}')

# ─── Pivot table tools ───
pivot_multiline = [
    ('create_pivot_table', 'create'),
    ('delete_pivot_table', 'delete'),
    ('refresh_pivot_table', 'refresh'),
    ('add_pivot_field', 'add_field'),
    ('remove_pivot_field', 'remove_field'),
    ('set_pivot_layout', 'configure'),
    ('set_pivot_table_options', 'configure'),
    ('apply_pivot_manual_filter', 'filter'),
    ('sort_pivot_field_labels', 'sort'),
    ('sort_pivot_field_values', 'sort'),
    ('pivot_table_exists', 'list'),
    ('get_pivot_table_location', 'get_info'),
    ('get_pivot_table_source_info', 'get_info'),
    ('get_pivot_hierarchy_counts', 'get_info'),
    ('get_pivot_hierarchies', 'get_info'),
]
for old, new_action in pivot_multiline:
    content, cnt = replace_multiline_tool(content, 'pivotTableConfigs', old, 'pivotTableConfigs', 'pivot', new_action)
    if cnt: print(f'  pivot/{old} → {new_action}: {cnt}')

# apply_pivot_label_filter (multi-line)
pattern_plf = re.compile(r"pivotTableConfigs,\n(\s+)'apply_pivot_label_filter',\n(\s+)\{", re.MULTILINE)
def repl_plf(m):
    return f"pivotTableConfigs,\n{m.group(1)}'pivot',\n{m.group(2)}{{ action: 'filter', filterType: 'label',"
content, cnt = pattern_plf.subn(repl_plf, content)
print(f'  apply_pivot_label_filter: {cnt}')

# clear_pivot_field_filters (multi-line)
pattern_cpff = re.compile(r"pivotTableConfigs,\n(\s+)'clear_pivot_field_filters',\n(\s+)\{", re.MULTILINE)
def repl_cpff(m):
    return f"pivotTableConfigs,\n{m.group(1)}'pivot',\n{m.group(2)}{{ action: 'filter', filterType: 'clear',"
content, cnt = pattern_cpff.subn(repl_cpff, content)
print(f'  clear_pivot_field_filters: {cnt}')

print(f'\nAfter multi-line tool renames: {content.count(chr(10))} lines')

# ─── Additional arg renames that need to be done now ───

# sort_pivot_field_labels → add sortMode: 'labels' (already has action: 'sort',)
# sort_pivot_field_values → add sortMode: 'values' (already has action: 'sort',)
# Problem: both are now mapped to action: 'sort', and we need to distinguish
# The sort_field_labels calls have sortBy but no valuesHierarchyName
# The sort_field_values calls have valuesHierarchyName

# Let's add sortMode based on context
# For sort_pivot_field_labels: the calls have 'sortBy' but not 'valuesHierarchyName' in the same args
# For sort_pivot_field_values: the calls have 'valuesHierarchyName'

# A simpler approach: search for the specific call patterns
# sort_pivot_field_labels calls:
# PT_NAME: { action: 'sort', pivotTableName: PT_NAME, fieldName: 'Region', sortBy: 'Descending',
# PT_V: { action: 'sort', pivotTableName: PT_V, fieldName: 'Region', sortBy: 'Ascending',

# sort_pivot_field_values calls:
# PT_NAME: { action: 'sort', ... valuesHierarchyName: 'Sales',
# PT_V: { action: 'sort', ... valuesHierarchyName: 'Sales',

# Add sortMode: 'labels' where there's no valuesHierarchyName near the action: 'sort'
# This is complex. Let me use a regex that looks for sort calls

# Actually let me look at the exact text to understand the patterns better
idx_sort = content.find("action: 'sort',")
ctx = content[idx_sort-100:idx_sort+200]
print('Sort context 1:', repr(ctx[:200]))

with open(r'D:\source\office-coding-agent\tests-e2e\src\test-taskpane.ts', 'w', encoding='utf-8') as f:
    f.write(content)

print('File written (pass 2)')
