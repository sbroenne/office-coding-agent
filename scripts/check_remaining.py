content = open(r'D:\source\office-coding-agent\tests-e2e\src\test-taskpane.ts', 'r', encoding='utf-8').read()
old_names = ['set_list_validation', 'set_date_validation', 'set_text_length_validation', 
             'set_custom_validation', 'set_number_validation', 'set_range_values',
             'sort_range', 'delete_range', 'insert_range', 'replace_values', 'remove_duplicates',
             'set_range_formulas', 'add_table_rows', 'sort_table', 'filter_table',
             'add_table_column', 'delete_table_column', 'set_table_style',
             'set_table_header_totals_visibility', 'create_chart', 'set_chart_title',
             'set_chart_type', 'set_chart_data_source', 'set_chart_legend_visibility',
             'set_chart_position', 'rename_sheet', 'set_sheet_gridlines', 'set_sheet_headings',
             'set_page_layout', 'set_workbook_properties', 'define_named_range',
             'add_comment', 'edit_comment', 'add_color_scale', 'add_data_bar',
             'add_cell_value_format', 'add_top_bottom_format', 'add_contains_text_format',
             'add_custom_format', 'list_conditional_formats', 'clear_conditional_formats',
             'get_data_validation', 'apply_pivot_label_filter', 'clear_pivot_field_filters',
             'sort_pivot_field_labels', 'sort_pivot_field_values', 'pivot_table_exists',
             'get_pivot_table_location', 'get_pivot_table_source_info', 'get_pivot_hierarchy_counts',
             'get_pivot_hierarchies', 'set_pivot_layout', 'add_pivot_field', 'remove_pivot_field',
             'set_pivot_table_options', 'apply_pivot_manual_filter', 'delete_pivot_table',
             'create_pivot_table', 'refresh_pivot_table']
for name in old_names:
    cnt = content.count("'" + name + "'")
    if cnt > 0:
        print(name + ': ' + str(cnt) + ' occurrences')
        # Show context
        idx = content.find("'" + name + "'")
        print('  ' + repr(content[max(0,idx-40):idx+60]))
