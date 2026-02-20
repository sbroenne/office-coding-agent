"""
Transform test-taskpane.ts according to the refactoring instructions.
Steps 1, 2, and 3 (all invocation updates + arg renames + verify lambda updates + block removals).
"""
import re

with open(r'D:\source\office-coding-agent\tests-e2e\src\test-taskpane.ts', 'r', encoding='utf-8') as f:
    content = f.read()

print(f'Original: {content.count(chr(10))} lines')

# ─── STEP 1 & 2: Remove legacy infrastructure, replace callTool body ───

legacy_start = content.find('/** Flat registry keyed by config name for the legacy mapping table. */')
old_calltool_end = content.find('\n}\n\n/**\n * Run a tool test:')

assert legacy_start > 0, "Could not find legacy start"
assert old_calltool_end > 0, "Could not find old callTool end"

new_calltool = (
    'async function callTool(configs: readonly ToolConfig[], name: string, args: Record<string, unknown> = {}): Promise<unknown> {\n'
    '  const config = configs.find(c => c.name === name);\n'
    '  if (!config) throw new Error(`Tool config not found: ${name}`);\n'
    '  let result: unknown;\n'
    '  await Excel.run(async context => {\n'
    '    result = await config.execute(context, args);\n'
    '  });\n'
    '  return result;\n'
    '}'
)

content = content[:legacy_start] + new_calltool + content[old_calltool_end + 2:]
print(f'After step 1&2: {content.count(chr(10))} lines')


# ─── Helper ───
def replace_tool_call(text, old_configs, old_name, new_configs, new_name, new_action):
    """Replace: old_configs, 'old_name', { → new_configs, 'new_name', { action: 'new_action',"""
    old_p = f"{old_configs}, '{old_name}', {{"
    new_p = f"{new_configs}, '{new_name}', {{ action: '{new_action}',"
    cnt = text.count(old_p)
    return text.replace(old_p, new_p), cnt


# ─── STEP 3a: Range tools ───

range_tool_map = [
    ('get_range_values', 'get_values'),
    ('set_range_values', 'set_values'),
    ('get_used_range', 'get_used'),
    ('clear_range', 'clear'),
    ('sort_range', 'sort'),
    ('find_values', 'find'),
    ('replace_values', 'replace'),
    ('insert_range', 'insert'),
    ('delete_range', 'delete'),
    ('merge_cells', 'merge'),
    ('unmerge_cells', 'unmerge'),
    ('group_rows_columns', 'group'),
    ('ungroup_rows_columns', 'ungroup'),
    ('remove_duplicates', 'remove_duplicates'),
    ('flash_fill_range', 'flash_fill'),
    ('get_special_cells', 'get_special_cells'),
    ('get_range_precedents', 'get_precedents'),
    ('get_range_dependents', 'get_dependents'),
    ('recalculate_range', 'recalculate'),
    ('get_tables_for_range', 'get_tables'),
    ('set_range_formulas', 'set_formulas'),
    ('get_range_formulas', 'get_formulas'),
    ('auto_fill_range', 'fill'),
]

for old_name, new_action in range_tool_map:
    content, cnt = replace_tool_call(content, 'rangeConfigs', old_name, 'rangeConfigs', 'range', new_action)
    if cnt: print(f'  range/{old_name} → {new_action}: {cnt}')

# Range format tools (configs change to rangeFormatConfigs, name → 'range_format')
range_format_map = [
    ('format_range', 'format'),
    ('set_number_format', 'set_number_format'),
    ('set_cell_borders', 'set_borders'),
    ('set_hyperlink', 'set_hyperlink'),
    ('toggle_row_column_visibility', 'toggle_visibility'),
]
for old_name, new_action in range_format_map:
    content, cnt = replace_tool_call(content, 'rangeConfigs', old_name, 'rangeFormatConfigs', 'range_format', new_action)
    if cnt: print(f'  range_format/{old_name} → {new_action}: {cnt}')

# auto_fit_columns: add fitTarget: 'columns'
old_p = "rangeConfigs, 'auto_fit_columns', {"
new_p = "rangeFormatConfigs, 'range_format', { action: 'auto_fit', fitTarget: 'columns',"
cnt = content.count(old_p); content = content.replace(old_p, new_p)
print(f'  auto_fit_columns: {cnt}')

old_p = "rangeConfigs, 'auto_fit_rows', {"
new_p = "rangeFormatConfigs, 'range_format', { action: 'auto_fit', fitTarget: 'rows',"
cnt = content.count(old_p); content = content.replace(old_p, new_p)
print(f'  auto_fit_rows: {cnt}')

# ─── STEP 3b: Table tools ───
table_map = [
    ('list_tables', 'list'),
    ('create_table', 'create'),
    ('delete_table', 'delete'),
    ('get_table_data', 'get_data'),
    ('add_table_rows', 'add_rows'),
    ('sort_table', 'sort'),
    ('filter_table', 'filter'),
    ('clear_table_filters', 'clear_filters'),
    ('reapply_table_filters', 'reapply_filters'),
    ('add_table_column', 'add_column'),
    ('delete_table_column', 'delete_column'),
    ('convert_table_to_range', 'convert_to_range'),
    ('resize_table', 'resize'),
    ('set_table_style', 'configure'),
    ('set_table_header_totals_visibility', 'configure'),
]
for old_name, new_action in table_map:
    content, cnt = replace_tool_call(content, 'tableConfigs', old_name, 'tableConfigs', 'table', new_action)
    if cnt: print(f'  table/{old_name} → {new_action}: {cnt}')

# ─── STEP 3c: Chart tools ───
chart_map = [
    ('list_charts', 'list'),
    ('create_chart', 'create'),
    ('delete_chart', 'delete'),
    ('set_chart_title', 'configure'),
    ('set_chart_type', 'configure'),
    ('set_chart_data_source', 'configure'),
    ('set_chart_position', 'configure'),
    ('set_chart_legend_visibility', 'configure'),
    # Drop these (will remove blocks entirely below)
    ('set_chart_axis_title', 'configure'),
    ('set_chart_axis_visibility', 'configure'),
    ('set_chart_series_filtered', 'configure'),
]
for old_name, new_action in chart_map:
    content, cnt = replace_tool_call(content, 'chartConfigs', old_name, 'chartConfigs', 'chart', new_action)
    if cnt: print(f'  chart/{old_name} → {new_action}: {cnt}')

# ─── STEP 3d: Sheet tools ───
sheet_map = [
    ('list_sheets', 'list'),
    ('create_sheet', 'create'),
    ('delete_sheet', 'delete'),
    ('rename_sheet', 'rename'),
    ('copy_sheet', 'copy'),
    ('move_sheet', 'move'),
    ('activate_sheet', 'activate'),
    ('protect_sheet', 'protect'),
    ('unprotect_sheet', 'unprotect'),
    ('freeze_panes', 'freeze'),
    ('set_sheet_visibility', 'set_visibility'),
    ('set_sheet_gridlines', 'set_gridlines'),
    ('set_sheet_headings', 'set_headings'),
    ('set_page_layout', 'set_page_layout'),
    ('recalculate_sheet', 'recalculate'),
]
for old_name, new_action in sheet_map:
    content, cnt = replace_tool_call(content, 'sheetConfigs', old_name, 'sheetConfigs', 'sheet', new_action)
    if cnt: print(f'  sheet/{old_name} → {new_action}: {cnt}')

# ─── STEP 3e: Workbook tools ───
workbook_map = [
    ('get_workbook_info', 'get_info'),
    ('get_selected_range', 'get_selected_range'),
    ('get_workbook_properties', 'get_properties'),
    ('set_workbook_properties', 'set_properties'),
    ('protect_workbook', 'protect'),
    ('unprotect_workbook', 'unprotect'),
    ('save_workbook', 'save'),
    ('recalculate_workbook', 'recalculate'),
    ('refresh_data_connections', 'refresh_connections'),
    ('define_named_range', 'define_named_range'),
    ('list_named_ranges', 'list_named_ranges'),
    ('list_queries', 'list_queries'),
    ('get_query', 'get_query'),
    ('get_workbook_protection', 'get_info'),
    ('get_query_count', 'list_queries'),
]
for old_name, new_action in workbook_map:
    content, cnt = replace_tool_call(content, 'workbookConfigs', old_name, 'workbookConfigs', 'workbook', new_action)
    if cnt: print(f'  workbook/{old_name} → {new_action}: {cnt}')

# ─── STEP 3f: Comment tools ───
comment_map = [
    ('list_comments', 'list'),
    ('add_comment', 'add'),
    ('edit_comment', 'edit'),
    ('delete_comment', 'delete'),
]
for old_name, new_action in comment_map:
    content, cnt = replace_tool_call(content, 'commentConfigs', old_name, 'commentConfigs', 'comment', new_action)
    if cnt: print(f'  comment/{old_name} → {new_action}: {cnt}')

# ─── STEP 3g: Conditional format tools ───
cf_map = [
    ('list_conditional_formats', 'list'),
    ('clear_conditional_formats', 'clear'),
]
for old_name, new_action in cf_map:
    content, cnt = replace_tool_call(content, 'conditionalFormatConfigs', old_name, 'conditionalFormatConfigs', 'conditional_format', new_action)
    if cnt: print(f'  cf/{old_name} → {new_action}: {cnt}')

# CF add operations with type
cf_add_map = [
    ('add_color_scale', 'colorScale'),
    ('add_data_bar', 'dataBar'),
    ('add_cell_value_format', 'cellValue'),
    ('add_top_bottom_format', 'topBottom'),
    ('add_contains_text_format', 'containsText'),
    ('add_custom_format', 'custom'),
]
for old_name, cf_type in cf_add_map:
    old_p = f"conditionalFormatConfigs, '{old_name}', {{"
    new_p = f"conditionalFormatConfigs, 'conditional_format', {{ action: 'add', type: '{cf_type}',"
    cnt = content.count(old_p); content = content.replace(old_p, new_p)
    if cnt: print(f'  cf/{old_name} → add/{cf_type}: {cnt}')

# ─── STEP 3h: Data validation tools ───
dv_map = [
    ('get_data_validation', 'get'),
    ('clear_data_validation', 'clear'),
]
for old_name, new_action in dv_map:
    content, cnt = replace_tool_call(content, 'dataValidationConfigs', old_name, 'dataValidationConfigs', 'data_validation', new_action)
    if cnt: print(f'  dv/{old_name} → {new_action}: {cnt}')

# set_list_validation → data_validation/set, type: 'list'
old_p = "dataValidationConfigs, 'set_list_validation', {"
new_p = "dataValidationConfigs, 'data_validation', { action: 'set', type: 'list',"
cnt = content.count(old_p); content = content.replace(old_p, new_p)
print(f'  set_list_validation: {cnt}')

old_p = "dataValidationConfigs, 'set_date_validation', {"
new_p = "dataValidationConfigs, 'data_validation', { action: 'set', type: 'date',"
cnt = content.count(old_p); content = content.replace(old_p, new_p)
print(f'  set_date_validation: {cnt}')

old_p = "dataValidationConfigs, 'set_text_length_validation', {"
new_p = "dataValidationConfigs, 'data_validation', { action: 'set', type: 'textLength',"
cnt = content.count(old_p); content = content.replace(old_p, new_p)
print(f'  set_text_length_validation: {cnt}')

old_p = "dataValidationConfigs, 'set_custom_validation', {"
new_p = "dataValidationConfigs, 'data_validation', { action: 'set', type: 'custom',"
cnt = content.count(old_p); content = content.replace(old_p, new_p)
print(f'  set_custom_validation: {cnt}')

# set_number_validation: no type arg yet, will handle numberType rename below
old_p = "dataValidationConfigs, 'set_number_validation', {"
new_p = "dataValidationConfigs, 'data_validation', { action: 'set',"
cnt = content.count(old_p); content = content.replace(old_p, new_p)
print(f'  set_number_validation: {cnt}')

# ─── STEP 3i: Pivot table tools ───
pivot_map = [
    ('list_pivot_tables', 'list'),
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
    ('get_pivot_table_count', 'list'),
    ('pivot_table_exists', 'list'),
    ('get_pivot_table_location', 'get_info'),
    ('refresh_all_pivot_tables', 'refresh'),
    ('get_pivot_table_source_info', 'get_info'),
    ('get_pivot_hierarchy_counts', 'get_info'),
    ('get_pivot_hierarchies', 'get_info'),
]
for old_name, new_action in pivot_map:
    content, cnt = replace_tool_call(content, 'pivotTableConfigs', old_name, 'pivotTableConfigs', 'pivot', new_action)
    if cnt: print(f'  pivot/{old_name} → {new_action}: {cnt}')

# apply_pivot_label_filter → pivot/filter with filterType: 'label'
old_p = "pivotTableConfigs, 'apply_pivot_label_filter', {"
new_p = "pivotTableConfigs, 'pivot', { action: 'filter', filterType: 'label',"
cnt = content.count(old_p); content = content.replace(old_p, new_p)
print(f'  apply_pivot_label_filter: {cnt}')

# clear_pivot_field_filters → pivot/filter with filterType: 'clear'
old_p = "pivotTableConfigs, 'clear_pivot_field_filters', {"
new_p = "pivotTableConfigs, 'pivot', { action: 'filter', filterType: 'clear',"
cnt = content.count(old_p); content = content.replace(old_p, new_p)
print(f'  clear_pivot_field_filters: {cnt}')

# Dropped pivot operations (still in tests as direct calls - handle individually below)
pivot_dropped = [
    'get_pivot_field_filters',
    'set_pivot_field_show_all_items',
    'get_pivot_layout_ranges',
    'set_pivot_layout_display_options',
    'get_pivot_data_hierarchy_for_cell',
    'get_pivot_items_for_cell',
    'set_pivot_layout_auto_sort_on_cell',
    'get_pivot_field_items',
]
for old_name in pivot_dropped:
    old_p = f"pivotTableConfigs, '{old_name}', {{"
    cnt = content.count(old_p)
    if cnt: print(f'  WARNING: dropped pivot {old_name} still present: {cnt}')

print(f'\nAfter basic name/action transforms: {content.count(chr(10))} lines')

# ─── STEP 3j: Arg renames ───

# copy_range: sourceAddress → address (in range/copy calls)
# We need to do this carefully - only rename in copy_range context
# Pattern: action: 'copy',...  sourceAddress:
# The rename is: sourceAddress → address
# But we need to be careful not to rename elsewhere
# Looking at the code, copy_range is used in one main place
# Let's do a targeted replacement
old_copy_src = "      sourceAddress: 'A1:C1',\n      destinationAddress: 'A10',"
new_copy_src = "      address: 'A1:C1',\n      destinationAddress: 'A10',"
cnt = content.count(old_copy_src); content = content.replace(old_copy_src, new_copy_src)
print(f'  copy_range sourceAddress→address: {cnt}')

# filter_table: values → filterValues
# Pattern: action: 'filter',...  column: 0, values: [
old_filter = "    { tableName: TABLE, column: 0, values: ['Widget', 'Gadget'] },"
new_filter = "    { tableName: TABLE, column: 0, filterValues: ['Widget', 'Gadget'] },"
cnt = content.count(old_filter); content = content.replace(old_filter, new_filter)
print(f'  filter_table values→filterValues: {cnt}')

# resize_table: newAddress → address
old_resize = "{ tableName: TABLE3, newAddress: 'U20:W25' },"
new_resize = "{ tableName: TABLE3, address: 'U20:W25' },"
cnt = content.count(old_resize); content = content.replace(old_resize, new_resize)
print(f'  resize_table newAddress→address: {cnt}')

# edit_comment: newText → text
old_edit1 = "    { cellAddress: 'L1', newText: 'Updated comment', sheetName: MAIN },"
new_edit1 = "    { cellAddress: 'L1', text: 'Updated comment', sheetName: MAIN },"
cnt = content.count(old_edit1); content = content.replace(old_edit1, new_edit1)
print(f'  edit_comment newText→text (1): {cnt}')

old_edit2 = "    { cellAddress: 'L5', newText: 'Edited no sheet' },"
new_edit2 = "    { cellAddress: 'L5', text: 'Edited no sheet' },"
cnt = content.count(old_edit2); content = content.replace(old_edit2, new_edit2)
print(f'  edit_comment newText→text (2): {cnt}')

# CF: add_data_bar barColor → fillColor
old_bar = "    { address: CF_RANGE, barColor: '#638EC6', sheetName: MAIN },"
new_bar = "    { address: CF_RANGE, fillColor: '#638EC6', sheetName: MAIN },"
cnt = content.count(old_bar); content = content.replace(old_bar, new_bar)
print(f'  add_data_bar barColor→fillColor: {cnt}')

# CF: add_cell_value_format fillColor → backgroundColor
# Multiple instances with different values
content = content.replace(
    "      operator: 'GreaterThan',\n      formula1: '40',\n      fillColor: '#00FF00',",
    "      operator: 'GreaterThan',\n      formula1: '40',\n      backgroundColor: '#00FF00',"
)
content = content.replace(
    "      operator: 'Between',\n      formula1: '20',\n      formula2: '50',\n      fontColor: '#0000FF',\n      fillColor: '#FFFF00',",
    "      operator: 'Between',\n      formula1: '20',\n      formula2: '50',\n      fontColor: '#0000FF',\n      backgroundColor: '#FFFF00',"
)
content = content.replace(
    "      operator: 'LessThan',\n      formula1: '25',\n      fillColor: 'orange',",
    "      operator: 'LessThan',\n      formula1: '25',\n      backgroundColor: 'orange',"
)
content = content.replace(
    "    { address: CF_RANGE, operator: 'EqualTo', formula1: '30', fillColor: 'cyan', sheetName: MAIN },",
    "    { address: CF_RANGE, operator: 'EqualTo', formula1: '30', backgroundColor: 'cyan', sheetName: MAIN },"
)
content = content.replace(
    "      operator: 'NotEqualTo',\n      formula1: '50',\n      fillColor: 'pink',",
    "      operator: 'NotEqualTo',\n      formula1: '50',\n      backgroundColor: 'pink',"
)
content = content.replace(
    "      operator: 'GreaterThanOrEqual',\n      formula1: '30',\n      fillColor: 'lime',",
    "      operator: 'GreaterThanOrEqual',\n      formula1: '30',\n      backgroundColor: 'lime',"
)
content = content.replace(
    "      operator: 'LessThanOrEqual',\n      formula1: '40',\n      fillColor: 'navy',",
    "      operator: 'LessThanOrEqual',\n      formula1: '40',\n      backgroundColor: 'navy',"
)
content = content.replace(
    "      operator: 'NotBetween',\n      formula1: '25',\n      formula2: '45',\n      fillColor: 'teal',",
    "      operator: 'NotBetween',\n      formula1: '25',\n      formula2: '45',\n      backgroundColor: 'teal',"
)
print('  CF cell_value fillColor→backgroundColor done')

# CF: add_top_bottom_format rank→topBottomRank, topOrBottom→topBottomType, fillColor→backgroundColor
# Multiple instances
tb_replacements = [
    (
        "    { address: CF_RANGE, rank: 3, topOrBottom: 'TopItems', fillColor: 'green', sheetName: MAIN },",
        "    { address: CF_RANGE, topBottomRank: 3, topBottomType: 'TopItems', backgroundColor: 'green', sheetName: MAIN },"
    ),
    (
        "    { address: CF_RANGE, rank: 2, topOrBottom: 'BottomItems', fillColor: 'red', sheetName: MAIN },",
        "    { address: CF_RANGE, topBottomRank: 2, topBottomType: 'BottomItems', backgroundColor: 'red', sheetName: MAIN },"
    ),
    (
        "      rank: 50,\n      topOrBottom: 'TopPercent',\n      fillColor: 'purple',",
        "      topBottomRank: 50,\n      topBottomType: 'TopPercent',\n      backgroundColor: 'purple',"
    ),
    (
        "    { address: CF_RANGE, fontColor: 'white', fillColor: 'black', sheetName: MAIN },",
        "    { address: CF_RANGE, fontColor: 'white', backgroundColor: 'black', sheetName: MAIN },"
    ),
    (
        "      rank: 25,\n      topOrBottom: 'BottomPercent',\n      fillColor: 'maroon',",
        "      topBottomRank: 25,\n      topBottomType: 'BottomPercent',\n      backgroundColor: 'maroon',"
    ),
]
for old, new in tb_replacements:
    cnt = content.count(old)
    content = content.replace(old, new)
    if cnt: print(f'  top_bottom arg rename: {cnt}')

# CF: add_contains_text_format text→containsText, fillColor→backgroundColor (fontColor stays)
ct_replacements = [
    (
        "    { address: 'B30:B36', text: 'Error', fontColor: 'red', sheetName: MAIN },",
        "    { address: 'B30:B36', containsText: 'Error', fontColor: 'red', sheetName: MAIN },"
    ),
    (
        "    { address: 'B30:B36', text: 'OK', sheetName: MAIN },",
        "    { address: 'B30:B36', containsText: 'OK', sheetName: MAIN },"
    ),
    (
        "    { address: 'B30:B36', text: 'Warning', fillColor: '#FFA500', sheetName: MAIN },",
        "    { address: 'B30:B36', containsText: 'Warning', backgroundColor: '#FFA500', sheetName: MAIN },"
    ),
]
for old, new in ct_replacements:
    cnt = content.count(old)
    content = content.replace(old, new)
    if cnt: print(f'  contains_text arg rename: {cnt}')

# CF: add_custom_format formula→formula1, fillColor→backgroundColor (fontColor stays)
cf_custom_replacements = [
    (
        "    { address: CF_RANGE, formula: '=A30>50', fillColor: '#FF00FF', sheetName: MAIN },",
        "    { address: CF_RANGE, formula1: '=A30>50', backgroundColor: '#FF00FF', sheetName: MAIN },"
    ),
    (
        "    { address: CF_RANGE, formula: '=A30<20', fontColor: 'red', sheetName: MAIN },",
        "    { address: CF_RANGE, formula1: '=A30<20', fontColor: 'red', sheetName: MAIN },"
    ),
]
for old, new in cf_custom_replacements:
    cnt = content.count(old)
    content = content.replace(old, new)
    if cnt: print(f'  custom_format arg rename: {cnt}')

# Data validation: set_list_validation source→listValues, remove inCellDropDown
# testDataValidationTools: source: 'Yes,No,Maybe'
old_dv1 = "    { address: 'C30', source: 'Yes,No,Maybe', sheetName: MAIN },"
new_dv1 = "    { address: 'C30', listValues: ['Yes', 'No', 'Maybe'], sheetName: MAIN },"
cnt = content.count(old_dv1); content = content.replace(old_dv1, new_dv1)
print(f'  set_list source→listValues (1): {cnt}')

# testDataValidationToolVariants: source: 'A,B,C', remove inCellDropDown
old_dv2 = (
    "      address: 'C31',\n"
    "      source: 'A,B,C',\n"
    "      inCellDropDown: false,\n"
)
new_dv2 = (
    "      address: 'C31',\n"
    "      listValues: ['A', 'B', 'C'],\n"
)
cnt = content.count(old_dv2); content = content.replace(old_dv2, new_dv2)
print(f'  set_list source→listValues (2): {cnt}')

# set_number_validation: replace numberType: 'wholeNumber' → type: 'number', numberType: 'decimal' → type: 'decimal'
# These appear in args after action: 'set',
# Pattern: action: 'set', ... numberType: 'wholeNumber' → type: 'number'
content = content.replace(
    "{ action: 'set',\n      address: 'D30',\n      numberType: 'wholeNumber',",
    "{ action: 'set',\n      address: 'D30',\n      type: 'number',"
)
content = content.replace(
    "{ action: 'set',\n      address: 'D31',\n      numberType: 'decimal',",
    "{ action: 'set',\n      address: 'D31',\n      type: 'decimal',"
)
content = content.replace(
    "{ action: 'set',\n      address: 'D32',\n      numberType: 'wholeNumber',",
    "{ action: 'set',\n      address: 'D32',\n      type: 'number',"
)
content = content.replace(
    "{ action: 'set',\n      address: 'D33',\n      numberType: 'wholeNumber',",
    "{ action: 'set',\n      address: 'D33',\n      type: 'number',"
)
content = content.replace(
    "{ action: 'set',\n      address: 'H31',\n      numberType: 'wholeNumber',",
    "{ action: 'set',\n      address: 'H31',\n      type: 'number',"
)
content = content.replace(
    "{ action: 'set',\n      address: 'H32',\n      numberType: 'decimal',",
    "{ action: 'set',\n      address: 'H32',\n      type: 'decimal',"
)
print('  set_number_validation numberType→type done')

# set_custom_validation: formula → customFormula
content = content.replace(
    "{ action: 'set', type: 'custom',\n      address: 'G30', formula: '=LEN(G30)<=100',",
    "{ action: 'set', type: 'custom',\n      address: 'G30', customFormula: '=LEN(G30)<=100',"
)
content = content.replace(
    "{ action: 'set', type: 'custom',\n      address: 'G31',\n      formula: '=ISNUMBER(G31)',",
    "{ action: 'set', type: 'custom',\n      address: 'G31',\n      customFormula: '=ISNUMBER(G31)',"
)
print('  set_custom_validation formula→customFormula done')

# Pivot: apply_pivot_label_filter condition→labelCondition, value1→labelValue1, value2→labelValue2
# testPivotTableTools:
content = content.replace(
    "      pivotTableName: PT_NAME,\n      fieldName: 'Region',\n      condition: 'Contains',\n      value1: 'North',",
    "      pivotTableName: PT_NAME,\n      fieldName: 'Region',\n      labelCondition: 'Contains',\n      labelValue1: 'North',"
)
# testPivotTableToolVariants (Between with value1/value2):
content = content.replace(
    "      pivotTableName: PT_V,\n      fieldName: 'Region',\n      condition: 'Between',\n      value1: 'A',\n      value2: 'Z',",
    "      pivotTableName: PT_V,\n      fieldName: 'Region',\n      labelCondition: 'Between',\n      labelValue1: 'A',\n      labelValue2: 'Z',"
)
print('  pivot label filter renames done')

# apply_pivot_manual_filter: add filterType: 'manual' explicitly (it's in action now but need to add filterType)
# The replace_tool_call already added action: 'filter', but for manual we need filterType: 'manual'
# Let me check the current state in content
# The replacement was: pivotTableConfigs, 'apply_pivot_manual_filter', { → pivotTableConfigs, 'pivot', { action: 'filter',
# But we need filterType: 'manual'. Let me fix those
old_manual = "pivotTableConfigs, 'pivot', { action: 'filter',\n      pivotTableName: PT_NAME,\n      fieldName: 'Region',\n      selectedItems: ['North'],"
new_manual = "pivotTableConfigs, 'pivot', { action: 'filter', filterType: 'manual',\n      pivotTableName: PT_NAME,\n      fieldName: 'Region',\n      selectedItems: ['North'],"
cnt = content.count(old_manual); content = content.replace(old_manual, new_manual)
print(f'  pivot manual filter PT_NAME: {cnt}')

old_manual2 = "pivotTableConfigs, 'pivot', { action: 'filter',\n      pivotTableName: PT_V,\n      fieldName: 'Region',\n      selectedItems: ['North', 'South'],"
new_manual2 = "pivotTableConfigs, 'pivot', { action: 'filter', filterType: 'manual',\n      pivotTableName: PT_V,\n      fieldName: 'Region',\n      selectedItems: ['North', 'South'],"
cnt = content.count(old_manual2); content = content.replace(old_manual2, new_manual2)
print(f'  pivot manual filter PT_V: {cnt}')

# sort_pivot_field_labels: add sortMode: 'labels'
# These are already renamed to action: 'sort', need to add sortMode
content = content.replace(
    "{ action: 'sort',\n    { pivotTableName: PT_NAME, fieldName: 'Region', sortBy: 'Descending',",
    "{ action: 'sort', sortMode: 'labels',\n    { pivotTableName: PT_NAME, fieldName: 'Region', sortBy: 'Descending',"
)

# The sort replacements generate: pivotTableConfigs, 'pivot', { action: 'sort', pivotTableName: ...
# But sort_pivot_field_labels and sort_pivot_field_values were mapped together to 'sort'
# We need to differentiate them. Let me search for the actual patterns.

# Actually let me check what the text looks like now for sort operations
idx = content.find("'sort_pivot_field_labels'")
if idx > 0:
    print(f'WARNING: sort_pivot_field_labels still present at {idx}')
else:
    print('  sort_pivot_field_labels replaced OK')

# Let me find the sort_pivot_field_labels occurrences in the new format
sort_labels_idx = content.find("action: 'sort',\n    { pivotTableName: PT_NAME, fieldName: 'Region', sortBy: 'Descending', sheetName: PIVOT_DST }")
print(f'sort_labels_idx: {sort_labels_idx}')

print(f'\nAfter all arg renames: {content.count(chr(10))} lines')

with open(r'D:\source\office-coding-agent\tests-e2e\src\test-taskpane.ts', 'w', encoding='utf-8') as f:
    f.write(content)

print('File written (intermediate state)')
