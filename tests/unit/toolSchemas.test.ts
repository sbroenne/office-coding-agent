import { describe, it, expect } from 'vitest';
import { excelTools, allConfigs } from '@/tools';
import { createTools } from '@/tools/codegen';

/**
 * Tool schema tests — validate that every tool's Zod inputSchema
 * accepts valid inputs and rejects invalid ones.
 *
 * These DON'T execute the tools (no Excel.run needed), they only
 * exercise the Zod schemas that guard the AI function-call contract.
 *
 * Note: AI SDK wraps Zod schemas in `FlexibleSchema`, which hides
 * `.safeParse()` from TypeScript. At runtime the underlying Zod
 * schema is there, so we cast through `unknown` to access it.
 */

interface ZodLike {
  safeParse: (data: unknown) => { success: boolean; error?: any };
}

/** Cast an AI SDK FlexibleSchema to a plain Zod-like object for testing */
function asZod(schema: unknown): ZodLike {
  return schema as ZodLike;
}

/** All expected tool names after decomposition */
const ALL_TOOL_NAMES = [
  // range (31)
  'get_range_values',
  'set_range_values',
  'get_used_range',
  'clear_range',
  'format_range',
  'set_number_format',
  'auto_fit_columns',
  'auto_fit_rows',
  'set_range_formulas',
  'get_range_formulas',
  'sort_range',
  'auto_fill_range',
  'flash_fill_range',
  'get_special_cells',
  'get_range_precedents',
  'get_range_dependents',
  'recalculate_range',
  'get_tables_for_range',
  'copy_range',
  'find_values',
  'insert_range',
  'delete_range',
  'merge_cells',
  'unmerge_cells',
  'replace_values',
  'remove_duplicates',
  'set_hyperlink',
  'toggle_row_column_visibility',
  'group_rows_columns',
  'ungroup_rows_columns',
  'set_cell_borders',
  // table (15)
  'list_tables',
  'create_table',
  'add_table_rows',
  'get_table_data',
  'delete_table',
  'sort_table',
  'filter_table',
  'clear_table_filters',
  'add_table_column',
  'delete_table_column',
  'convert_table_to_range',
  'resize_table',
  'set_table_style',
  'set_table_header_totals_visibility',
  'reapply_table_filters',
  // chart (11)
  'list_charts',
  'create_chart',
  'delete_chart',
  'set_chart_title',
  'set_chart_type',
  'set_chart_data_source',
  'set_chart_position',
  'set_chart_legend_visibility',
  'set_chart_axis_title',
  'set_chart_axis_visibility',
  'set_chart_series_filtered',
  // sheet (16)
  'list_sheets',
  'create_sheet',
  'rename_sheet',
  'delete_sheet',
  'activate_sheet',
  'freeze_panes',
  'protect_sheet',
  'unprotect_sheet',
  'set_sheet_visibility',
  'copy_sheet',
  'move_sheet',
  'set_page_layout',
  'set_sheet_gridlines',
  'set_sheet_headings',
  'recalculate_sheet',
  // workbook (15)
  'get_workbook_info',
  'get_selected_range',
  'define_named_range',
  'list_named_ranges',
  'recalculate_workbook',
  'save_workbook',
  'get_workbook_properties',
  'set_workbook_properties',
  'get_workbook_protection',
  'protect_workbook',
  'unprotect_workbook',
  'refresh_data_connections',
  'list_queries',
  'get_query',
  'get_query_count',
  // comment (4)
  'add_comment',
  'list_comments',
  'edit_comment',
  'delete_comment',
  // conditional format (8 — decomposed from 3)
  'add_color_scale',
  'add_data_bar',
  'add_cell_value_format',
  'add_top_bottom_format',
  'add_contains_text_format',
  'add_custom_format',
  'list_conditional_formats',
  'clear_conditional_formats',
  // data validation (7 — decomposed from 3)
  'set_list_validation',
  'set_number_validation',
  'set_date_validation',
  'set_text_length_validation',
  'set_custom_validation',
  'get_data_validation',
  'clear_data_validation',
  // pivot table (16)
  'list_pivot_tables',
  'refresh_pivot_table',
  'get_pivot_table_source_info',
  'set_pivot_table_options',
  'delete_pivot_table',
  'create_pivot_table',
  'add_pivot_field',
  'set_pivot_layout',
  'get_pivot_field_filters',
  'clear_pivot_field_filters',
  'apply_pivot_label_filter',
  'sort_pivot_field_labels',
  'apply_pivot_manual_filter',
  'sort_pivot_field_values',
  'set_pivot_field_show_all_items',
  'remove_pivot_field',
] as const;

describe('tool schemas — structural', () => {
  it('excelTools record is non-empty and every tool has inputSchema and execute', () => {
    const names = Object.keys(excelTools);
    expect(names.length).toBeGreaterThan(0);

    for (const name of names) {
      const t = excelTools[name as keyof typeof excelTools];
      expect(t, `${name} is defined`).toBeDefined();
      expect(t.inputSchema, `${name} has inputSchema`).toBeDefined();
      expect(typeof t.execute, `${name} has execute fn`).toBe('function');
    }
  });

  it('excelTools contains exactly the expected tools', () => {
    const actual = Object.keys(excelTools).sort();
    const expected = [...ALL_TOOL_NAMES].sort();
    expect(actual).toEqual(expected);
  });

  it('every config produces a valid tool via createTools()', () => {
    for (const configs of allConfigs) {
      const tools = createTools(configs);
      for (const [name, tool] of Object.entries(tools)) {
        expect(tool, `${name} was created`).toBeDefined();
        expect(tool.inputSchema, `${name} has schema`).toBeDefined();
        expect(typeof tool.execute, `${name} has execute`).toBe('function');
      }
    }
  });
});

describe('tool schemas — range tools', () => {
  it('get_range_values accepts valid input', () => {
    const schema = excelTools.get_range_values.inputSchema;
    const result = asZod(schema).safeParse({ address: 'A1:C10' });
    expect(result.success).toBe(true);
  });

  it('get_range_values rejects missing address', () => {
    const schema = excelTools.get_range_values.inputSchema;
    const result = asZod(schema).safeParse({});
    expect(result.success).toBe(false);
  });

  it('set_range_values accepts address + 2D array', () => {
    const schema = excelTools.set_range_values.inputSchema;
    const result = asZod(schema).safeParse({
      address: 'A1:B2',
      values: [
        [1, 2],
        [3, 4],
      ],
    });
    expect(result.success).toBe(true);
  });

  it('set_range_values rejects missing values', () => {
    const schema = excelTools.set_range_values.inputSchema;
    const result = asZod(schema).safeParse({ address: 'A1:B2' });
    expect(result.success).toBe(false);
  });

  it('get_used_range accepts empty object', () => {
    const schema = excelTools.get_used_range.inputSchema;
    const result = asZod(schema).safeParse({});
    expect(result.success).toBe(true);
  });

  it('clear_range requires address', () => {
    const schema = excelTools.clear_range.inputSchema;
    expect(asZod(schema).safeParse({ address: 'A1:A10' }).success).toBe(true);
    expect(asZod(schema).safeParse({}).success).toBe(false);
  });

  it('format_range requires address', () => {
    const schema = excelTools.format_range.inputSchema;
    expect(asZod(schema).safeParse({ address: 'A1:B5', bold: true }).success).toBe(true);
    expect(asZod(schema).safeParse({}).success).toBe(false);
  });

  it('sort_range accepts column number and ascending flag', () => {
    const schema = excelTools.sort_range.inputSchema;
    expect(
      asZod(schema).safeParse({ address: 'A1:C10', column: 0, ascending: false }).success
    ).toBe(true);
  });

  it('copy_range requires source + destination', () => {
    const schema = excelTools.copy_range.inputSchema;
    expect(
      asZod(schema).safeParse({ sourceAddress: 'A1:B5', destinationAddress: 'D1' }).success
    ).toBe(true);
    expect(asZod(schema).safeParse({ sourceAddress: 'A1:B5' }).success).toBe(false);
  });

  it('replace_values requires find + replace', () => {
    const schema = excelTools.replace_values.inputSchema;
    expect(asZod(schema).safeParse({ find: 'foo', replace: 'bar' }).success).toBe(true);
    expect(asZod(schema).safeParse({ find: 'foo' }).success).toBe(false);
  });

  it('remove_duplicates requires address + columns', () => {
    const schema = excelTools.remove_duplicates.inputSchema;
    expect(asZod(schema).safeParse({ address: 'A1:D100', columns: ['0', '2'] }).success).toBe(true);
    expect(asZod(schema).safeParse({ address: 'A1:D100' }).success).toBe(false);
  });

  it('set_hyperlink requires address + url', () => {
    const schema = excelTools.set_hyperlink.inputSchema;
    expect(asZod(schema).safeParse({ address: 'A1', url: 'https://example.com' }).success).toBe(
      true
    );
    expect(asZod(schema).safeParse({ address: 'A1' }).success).toBe(false);
  });

  it('toggle_row_column_visibility requires address + hidden + target', () => {
    const schema = excelTools.toggle_row_column_visibility.inputSchema;
    expect(
      asZod(schema).safeParse({ address: 'A:C', hidden: true, target: 'columns' }).success
    ).toBe(true);
    expect(asZod(schema).safeParse({ address: 'A:C', hidden: true }).success).toBe(false);
  });

  it('group_rows_columns requires address', () => {
    const schema = excelTools.group_rows_columns.inputSchema;
    expect(asZod(schema).safeParse({ address: '3:5' }).success).toBe(true);
    expect(asZod(schema).safeParse({}).success).toBe(false);
  });

  it('ungroup_rows_columns requires address', () => {
    const schema = excelTools.ungroup_rows_columns.inputSchema;
    expect(asZod(schema).safeParse({ address: '3:5' }).success).toBe(true);
    expect(asZod(schema).safeParse({}).success).toBe(false);
  });

  it('set_cell_borders requires address + borderStyle + side', () => {
    const schema = excelTools.set_cell_borders.inputSchema;
    expect(
      asZod(schema).safeParse({ address: 'A1:C10', borderStyle: 'Thin', side: 'EdgeAll' }).success
    ).toBe(true);
    expect(asZod(schema).safeParse({ address: 'A1:C10', borderStyle: 'Thin' }).success).toBe(false);
  });

  it('auto_fill_range requires sourceAddress + destinationAddress', () => {
    const schema = excelTools.auto_fill_range.inputSchema;
    expect(
      asZod(schema).safeParse({ sourceAddress: 'A1:A2', destinationAddress: 'A1:A20' }).success
    ).toBe(true);
    expect(asZod(schema).safeParse({ sourceAddress: 'A1:A2' }).success).toBe(false);
  });

  it('flash_fill_range requires address', () => {
    const schema = excelTools.flash_fill_range.inputSchema;
    expect(asZod(schema).safeParse({ address: 'B2:B20' }).success).toBe(true);
    expect(asZod(schema).safeParse({}).success).toBe(false);
  });

  it('get_special_cells requires address + cellType', () => {
    const schema = excelTools.get_special_cells.inputSchema;
    expect(
      asZod(schema).safeParse({ address: 'A1:D100', cellType: 'Formulas', cellValueType: 'All' })
        .success
    ).toBe(true);
    expect(asZod(schema).safeParse({ address: 'A1:D100' }).success).toBe(false);
  });

  it('get_range_precedents requires address', () => {
    const schema = excelTools.get_range_precedents.inputSchema;
    expect(asZod(schema).safeParse({ address: 'D2:D10' }).success).toBe(true);
    expect(asZod(schema).safeParse({}).success).toBe(false);
  });

  it('get_range_dependents requires address', () => {
    const schema = excelTools.get_range_dependents.inputSchema;
    expect(asZod(schema).safeParse({ address: 'B2' }).success).toBe(true);
    expect(asZod(schema).safeParse({}).success).toBe(false);
  });

  it('recalculate_range requires address', () => {
    const schema = excelTools.recalculate_range.inputSchema;
    expect(asZod(schema).safeParse({ address: 'A1:Z100' }).success).toBe(true);
    expect(asZod(schema).safeParse({}).success).toBe(false);
  });

  it('get_tables_for_range requires address', () => {
    const schema = excelTools.get_tables_for_range.inputSchema;
    expect(asZod(schema).safeParse({ address: 'A1:H500', fullyContained: true }).success).toBe(
      true
    );
    expect(asZod(schema).safeParse({}).success).toBe(false);
  });
});

describe('tool schemas — sheet tools', () => {
  it('list_sheets accepts empty object', () => {
    const schema = excelTools.list_sheets.inputSchema;
    expect(asZod(schema).safeParse({}).success).toBe(true);
  });

  it('create_sheet requires a name', () => {
    const schema = excelTools.create_sheet.inputSchema;
    expect(asZod(schema).safeParse({ name: 'Sheet2' }).success).toBe(true);
    expect(asZod(schema).safeParse({}).success).toBe(false);
  });

  it('rename_sheet requires currentName and newName', () => {
    const schema = excelTools.rename_sheet.inputSchema;
    expect(asZod(schema).safeParse({ currentName: 'Sheet1', newName: 'Data' }).success).toBe(true);
    expect(asZod(schema).safeParse({ currentName: 'Sheet1' }).success).toBe(false);
  });

  it('freeze_panes requires name', () => {
    const schema = excelTools.freeze_panes.inputSchema;
    expect(asZod(schema).safeParse({ name: 'Sheet1', freezeAt: 'B3' }).success).toBe(true);
    expect(asZod(schema).safeParse({ name: 'Sheet1' }).success).toBe(true);
    expect(asZod(schema).safeParse({}).success).toBe(false);
  });

  it('protect_sheet requires name', () => {
    const schema = excelTools.protect_sheet.inputSchema;
    expect(asZod(schema).safeParse({ name: 'Sheet1' }).success).toBe(true);
    expect(asZod(schema).safeParse({ name: 'Sheet1', password: 'secret' }).success).toBe(true);
    expect(asZod(schema).safeParse({}).success).toBe(false);
  });

  it('unprotect_sheet requires name', () => {
    const schema = excelTools.unprotect_sheet.inputSchema;
    expect(asZod(schema).safeParse({ name: 'Sheet1' }).success).toBe(true);
    expect(asZod(schema).safeParse({}).success).toBe(false);
  });

  it('set_sheet_visibility requires name, optional visibility/tabColor', () => {
    const schema = excelTools.set_sheet_visibility.inputSchema;
    expect(asZod(schema).safeParse({ name: 'Sheet1', visibility: 'Hidden' }).success).toBe(true);
    expect(asZod(schema).safeParse({ name: 'Sheet1', tabColor: '#FF0000' }).success).toBe(true);
    expect(asZod(schema).safeParse({}).success).toBe(false);
  });

  it('copy_sheet requires name', () => {
    const schema = excelTools.copy_sheet.inputSchema;
    expect(asZod(schema).safeParse({ name: 'Sheet1' }).success).toBe(true);
    expect(asZod(schema).safeParse({ name: 'Sheet1', newName: 'Copy' }).success).toBe(true);
    expect(asZod(schema).safeParse({}).success).toBe(false);
  });

  it('move_sheet requires name + position', () => {
    const schema = excelTools.move_sheet.inputSchema;
    expect(asZod(schema).safeParse({ name: 'Sheet1', position: 2 }).success).toBe(true);
    expect(asZod(schema).safeParse({ name: 'Sheet1' }).success).toBe(false);
  });

  it('set_page_layout requires name, optional layout params', () => {
    const schema = excelTools.set_page_layout.inputSchema;
    expect(asZod(schema).safeParse({ name: 'Sheet1' }).success).toBe(true);
    expect(asZod(schema).safeParse({ name: 'Sheet1', orientation: 'Landscape' }).success).toBe(
      true
    );
    expect(asZod(schema).safeParse({}).success).toBe(false);
  });

  it('set_sheet_gridlines requires name + showGridlines', () => {
    const schema = excelTools.set_sheet_gridlines.inputSchema;
    expect(asZod(schema).safeParse({ name: 'Sheet1', showGridlines: true }).success).toBe(true);
    expect(asZod(schema).safeParse({ name: 'Sheet1' }).success).toBe(false);
  });

  it('set_sheet_headings requires name + showHeadings', () => {
    const schema = excelTools.set_sheet_headings.inputSchema;
    expect(asZod(schema).safeParse({ name: 'Sheet1', showHeadings: false }).success).toBe(true);
    expect(asZod(schema).safeParse({ name: 'Sheet1' }).success).toBe(false);
  });

  it('recalculate_sheet requires name, optional recalcType', () => {
    const schema = excelTools.recalculate_sheet.inputSchema;
    expect(asZod(schema).safeParse({ name: 'Sheet1' }).success).toBe(true);
    expect(asZod(schema).safeParse({ name: 'Sheet1', recalcType: 'Full' }).success).toBe(true);
    expect(asZod(schema).safeParse({}).success).toBe(false);
  });
});

describe('tool schemas — table tools', () => {
  it('add_table_column requires tableName, optional columnName', () => {
    const schema = excelTools.add_table_column.inputSchema;
    expect(asZod(schema).safeParse({ tableName: 'T1' }).success).toBe(true);
    expect(asZod(schema).safeParse({ tableName: 'T1', columnName: 'NewCol' }).success).toBe(true);
    expect(asZod(schema).safeParse({}).success).toBe(false);
  });

  it('delete_table_column requires tableName + columnName', () => {
    const schema = excelTools.delete_table_column.inputSchema;
    expect(asZod(schema).safeParse({ tableName: 'T1', columnName: 'ColA' }).success).toBe(true);
    expect(asZod(schema).safeParse({ tableName: 'T1' }).success).toBe(false);
  });

  it('convert_table_to_range requires tableName', () => {
    const schema = excelTools.convert_table_to_range.inputSchema;
    expect(asZod(schema).safeParse({ tableName: 'T1' }).success).toBe(true);
    expect(asZod(schema).safeParse({}).success).toBe(false);
  });

  it('resize_table requires tableName + newAddress', () => {
    const schema = excelTools.resize_table.inputSchema;
    expect(asZod(schema).safeParse({ tableName: 'T1', newAddress: 'A1:F100' }).success).toBe(true);
    expect(asZod(schema).safeParse({ tableName: 'T1' }).success).toBe(false);
  });

  it('set_table_style requires tableName + style', () => {
    const schema = excelTools.set_table_style.inputSchema;
    expect(asZod(schema).safeParse({ tableName: 'T1', style: 'TableStyleMedium2' }).success).toBe(
      true
    );
    expect(asZod(schema).safeParse({ tableName: 'T1' }).success).toBe(false);
  });

  it('set_table_header_totals_visibility requires tableName', () => {
    const schema = excelTools.set_table_header_totals_visibility.inputSchema;
    expect(asZod(schema).safeParse({ tableName: 'T1', showHeaders: true }).success).toBe(true);
    expect(asZod(schema).safeParse({ tableName: 'T1', showTotals: false }).success).toBe(true);
    expect(asZod(schema).safeParse({}).success).toBe(false);
  });

  it('reapply_table_filters requires tableName', () => {
    const schema = excelTools.reapply_table_filters.inputSchema;
    expect(asZod(schema).safeParse({ tableName: 'T1' }).success).toBe(true);
    expect(asZod(schema).safeParse({}).success).toBe(false);
  });
});

describe('tool schemas — chart tools', () => {
  it('list_charts accepts empty input', () => {
    const schema = excelTools.list_charts.inputSchema;
    expect(asZod(schema).safeParse({}).success).toBe(true);
  });

  it('create_chart requires dataRange and chartType', () => {
    const schema = excelTools.create_chart.inputSchema;
    expect(
      asZod(schema).safeParse({
        dataRange: 'A1:D10',
        chartType: 'ColumnClustered',
      }).success
    ).toBe(true);
    expect(asZod(schema).safeParse({ dataRange: 'A1:D10' }).success).toBe(false);
  });

  it('create_chart rejects invalid chart type', () => {
    const schema = excelTools.create_chart.inputSchema;
    expect(
      asZod(schema).safeParse({
        dataRange: 'A1:D10',
        chartType: 'InvalidType',
      }).success
    ).toBe(false);
  });

  it('set_chart_title requires chartName + title', () => {
    const schema = excelTools.set_chart_title.inputSchema;
    expect(asZod(schema).safeParse({ chartName: 'Chart1', title: 'Sales Report' }).success).toBe(
      true
    );
    expect(asZod(schema).safeParse({ chartName: 'Chart1' }).success).toBe(false);
  });

  it('set_chart_type requires chartName + chartType', () => {
    const schema = excelTools.set_chart_type.inputSchema;
    expect(asZod(schema).safeParse({ chartName: 'Chart1', chartType: 'Pie' }).success).toBe(true);
    expect(asZod(schema).safeParse({ chartName: 'Chart1' }).success).toBe(false);
  });

  it('set_chart_data_source requires chartName + dataRange', () => {
    const schema = excelTools.set_chart_data_source.inputSchema;
    expect(asZod(schema).safeParse({ chartName: 'Chart1', dataRange: 'B1:D20' }).success).toBe(
      true
    );
    expect(asZod(schema).safeParse({ chartName: 'Chart1' }).success).toBe(false);
  });

  it('set_chart_position requires chartName', () => {
    const schema = excelTools.set_chart_position.inputSchema;
    expect(asZod(schema).safeParse({ chartName: 'Chart1', left: 20, top: 30 }).success).toBe(true);
    expect(
      asZod(schema).safeParse({ chartName: 'Chart1', startCell: 'A1', endCell: 'H20' }).success
    ).toBe(true);
    expect(asZod(schema).safeParse({}).success).toBe(false);
  });

  it('set_chart_legend_visibility requires chartName + visible', () => {
    const schema = excelTools.set_chart_legend_visibility.inputSchema;
    expect(asZod(schema).safeParse({ chartName: 'Chart1', visible: true }).success).toBe(true);
    expect(asZod(schema).safeParse({ chartName: 'Chart1' }).success).toBe(false);
  });

  it('set_chart_axis_title requires chartName + axisType + title', () => {
    const schema = excelTools.set_chart_axis_title.inputSchema;
    expect(
      asZod(schema).safeParse({ chartName: 'Chart1', axisType: 'Value', title: 'Sales' }).success
    ).toBe(true);
    expect(asZod(schema).safeParse({ chartName: 'Chart1', axisType: 'Value' }).success).toBe(false);
  });

  it('set_chart_axis_visibility requires chartName + axisType + visible', () => {
    const schema = excelTools.set_chart_axis_visibility.inputSchema;
    expect(
      asZod(schema).safeParse({ chartName: 'Chart1', axisType: 'Category', visible: false }).success
    ).toBe(true);
    expect(asZod(schema).safeParse({ chartName: 'Chart1', axisType: 'Category' }).success).toBe(
      false
    );
  });

  it('set_chart_series_filtered requires chartName + seriesIndex + filtered', () => {
    const schema = excelTools.set_chart_series_filtered.inputSchema;
    expect(
      asZod(schema).safeParse({ chartName: 'Chart1', seriesIndex: 0, filtered: true }).success
    ).toBe(true);
    expect(asZod(schema).safeParse({ chartName: 'Chart1', seriesIndex: 0 }).success).toBe(false);
  });
});

describe('tool schemas — conditional format tools (decomposed)', () => {
  it('add_color_scale requires address, colors are optional', () => {
    const schema = excelTools.add_color_scale.inputSchema;
    expect(
      asZod(schema).safeParse({ address: 'A1:A10', minColor: '#FF0000', maxColor: '#00FF00' })
        .success
    ).toBe(true);
    expect(asZod(schema).safeParse({ address: 'A1:A10' }).success).toBe(true);
    expect(asZod(schema).safeParse({}).success).toBe(false);
  });

  it('add_data_bar requires address', () => {
    const schema = excelTools.add_data_bar.inputSchema;
    expect(asZod(schema).safeParse({ address: 'A1:A10' }).success).toBe(true);
    expect(asZod(schema).safeParse({}).success).toBe(false);
  });

  it('add_cell_value_format requires address + operator + formula1', () => {
    const schema = excelTools.add_cell_value_format.inputSchema;
    expect(
      asZod(schema).safeParse({
        address: 'A1:A10',
        operator: 'GreaterThan',
        formula1: '100',
        fillColor: '#FF0000',
      }).success
    ).toBe(true);
    expect(asZod(schema).safeParse({ address: 'A1:A10' }).success).toBe(false);
  });

  it('add_top_bottom_format requires address + topBottom + rank', () => {
    const schema = excelTools.add_top_bottom_format.inputSchema;
    expect(
      asZod(schema).safeParse({
        address: 'A1:A10',
        topBottom: 'top',
        rank: 5,
        fillColor: '#FFFF00',
      }).success
    ).toBe(true);
  });

  it('add_contains_text_format requires address + text + fillColor', () => {
    const schema = excelTools.add_contains_text_format.inputSchema;
    expect(
      asZod(schema).safeParse({
        address: 'A1:A10',
        text: 'Error',
        fillColor: '#FF0000',
      }).success
    ).toBe(true);
  });

  it('add_custom_format requires address + formula + fillColor', () => {
    const schema = excelTools.add_custom_format.inputSchema;
    expect(
      asZod(schema).safeParse({
        address: 'A1:A10',
        formula: '=A1>100',
        fillColor: '#00FF00',
      }).success
    ).toBe(true);
  });

  it('list_conditional_formats requires address', () => {
    const schema = excelTools.list_conditional_formats.inputSchema;
    expect(asZod(schema).safeParse({ address: 'A1:A10' }).success).toBe(true);
    expect(asZod(schema).safeParse({}).success).toBe(false);
  });

  it('clear_conditional_formats accepts optional address', () => {
    const schema = excelTools.clear_conditional_formats.inputSchema;
    expect(asZod(schema).safeParse({ address: 'A1:A10' }).success).toBe(true);
    expect(asZod(schema).safeParse({}).success).toBe(true);
  });
});

describe('tool schemas — data validation tools (decomposed)', () => {
  it('set_list_validation requires address + source', () => {
    const schema = excelTools.set_list_validation.inputSchema;
    expect(
      asZod(schema).safeParse({
        address: 'A1:A10',
        source: 'Yes,No,Maybe',
      }).success
    ).toBe(true);
    expect(asZod(schema).safeParse({ address: 'A1:A10' }).success).toBe(false);
  });

  it('set_number_validation requires address + numberType + operator + formula1', () => {
    const schema = excelTools.set_number_validation.inputSchema;
    expect(
      asZod(schema).safeParse({
        address: 'A1:A10',
        numberType: 'wholeNumber',
        operator: 'GreaterThan',
        formula1: '0',
      }).success
    ).toBe(true);
    expect(asZod(schema).safeParse({ address: 'A1:A10' }).success).toBe(false);
  });

  it('set_date_validation requires address + operator + formula1', () => {
    const schema = excelTools.set_date_validation.inputSchema;
    expect(
      asZod(schema).safeParse({
        address: 'A1:A10',
        operator: 'GreaterThan',
        formula1: '2024-01-01',
      }).success
    ).toBe(true);
    expect(asZod(schema).safeParse({ address: 'A1:A10' }).success).toBe(false);
  });

  it('set_text_length_validation requires address + operator + formula1', () => {
    const schema = excelTools.set_text_length_validation.inputSchema;
    expect(
      asZod(schema).safeParse({
        address: 'A1:A10',
        operator: 'LessThanOrEqualTo',
        formula1: '100',
      }).success
    ).toBe(true);
    expect(asZod(schema).safeParse({ address: 'A1:A10' }).success).toBe(false);
  });

  it('set_custom_validation requires address + formula', () => {
    const schema = excelTools.set_custom_validation.inputSchema;
    expect(
      asZod(schema).safeParse({
        address: 'A1:A10',
        formula: '=AND(A1>0, A1<100)',
      }).success
    ).toBe(true);
    expect(asZod(schema).safeParse({ address: 'A1:A10' }).success).toBe(false);
  });

  it('get_data_validation requires address', () => {
    const schema = excelTools.get_data_validation.inputSchema;
    expect(asZod(schema).safeParse({ address: 'F1:F10' }).success).toBe(true);
    expect(asZod(schema).safeParse({}).success).toBe(false);
  });

  it('clear_data_validation requires address', () => {
    const schema = excelTools.clear_data_validation.inputSchema;
    expect(asZod(schema).safeParse({ address: 'G1:G10' }).success).toBe(true);
    expect(asZod(schema).safeParse({}).success).toBe(false);
  });
});

describe('tool schemas — pivot table tools', () => {
  it('get_pivot_table_source_info requires pivotTableName', () => {
    const schema = excelTools.get_pivot_table_source_info.inputSchema;
    expect(asZod(schema).safeParse({ pivotTableName: 'PT1' }).success).toBe(true);
    expect(asZod(schema).safeParse({}).success).toBe(false);
  });

  it('set_pivot_table_options requires pivotTableName and accepts optional options', () => {
    const schema = excelTools.set_pivot_table_options.inputSchema;
    expect(asZod(schema).safeParse({ pivotTableName: 'PT1' }).success).toBe(true);
    expect(
      asZod(schema).safeParse({
        pivotTableName: 'PT1',
        allowMultipleFiltersPerField: true,
        useCustomSortLists: false,
        refreshOnOpen: true,
        enableDataValueEditing: false,
      }).success
    ).toBe(true);
    expect(asZod(schema).safeParse({ allowMultipleFiltersPerField: true }).success).toBe(false);
  });

  it('add_pivot_field requires pivotTableName + fieldName + fieldType', () => {
    const schema = excelTools.add_pivot_field.inputSchema;
    expect(
      asZod(schema).safeParse({
        pivotTableName: 'PT1',
        fieldName: 'Region',
        fieldType: 'row',
      }).success
    ).toBe(true);
    expect(asZod(schema).safeParse({ pivotTableName: 'PT1', fieldName: 'Region' }).success).toBe(
      false
    );
  });

  it('remove_pivot_field requires pivotTableName + fieldName + fieldType', () => {
    const schema = excelTools.remove_pivot_field.inputSchema;
    expect(
      asZod(schema).safeParse({
        pivotTableName: 'PT1',
        fieldName: 'Region',
        fieldType: 'row',
      }).success
    ).toBe(true);
    expect(asZod(schema).safeParse({ pivotTableName: 'PT1' }).success).toBe(false);
  });

  it('set_pivot_layout requires pivotTableName and accepts optional layout/display flags', () => {
    const schema = excelTools.set_pivot_layout.inputSchema;
    expect(
      asZod(schema).safeParse({
        pivotTableName: 'PT1',
        layoutType: 'Tabular',
        subtotalLocation: 'AtBottom',
        showFieldHeaders: true,
      }).success
    ).toBe(true);
    expect(asZod(schema).safeParse({ pivotTableName: 'PT1' }).success).toBe(true);
    expect(asZod(schema).safeParse({}).success).toBe(false);
  });

  it('get_pivot_field_filters requires pivotTableName + fieldName', () => {
    const schema = excelTools.get_pivot_field_filters.inputSchema;
    expect(asZod(schema).safeParse({ pivotTableName: 'PT1', fieldName: 'Region' }).success).toBe(
      true
    );
    expect(asZod(schema).safeParse({ pivotTableName: 'PT1' }).success).toBe(false);
  });

  it('clear_pivot_field_filters requires pivotTableName + fieldName and accepts optional filterType', () => {
    const schema = excelTools.clear_pivot_field_filters.inputSchema;
    expect(asZod(schema).safeParse({ pivotTableName: 'PT1', fieldName: 'Region' }).success).toBe(
      true
    );
    expect(
      asZod(schema).safeParse({
        pivotTableName: 'PT1',
        fieldName: 'Region',
        filterType: 'Label',
      }).success
    ).toBe(true);
    expect(asZod(schema).safeParse({ pivotTableName: 'PT1' }).success).toBe(false);
  });

  it('apply_pivot_label_filter requires pivotTableName + fieldName + condition + value1', () => {
    const schema = excelTools.apply_pivot_label_filter.inputSchema;
    expect(
      asZod(schema).safeParse({
        pivotTableName: 'PT1',
        fieldName: 'Region',
        condition: 'Contains',
        value1: 'North',
      }).success
    ).toBe(true);
    expect(
      asZod(schema).safeParse({
        pivotTableName: 'PT1',
        fieldName: 'Region',
        condition: 'Between',
        value1: 'A',
        value2: 'M',
      }).success
    ).toBe(true);
    expect(asZod(schema).safeParse({ pivotTableName: 'PT1', fieldName: 'Region' }).success).toBe(
      false
    );
  });

  it('sort_pivot_field_labels requires pivotTableName + fieldName + sortBy', () => {
    const schema = excelTools.sort_pivot_field_labels.inputSchema;
    expect(
      asZod(schema).safeParse({
        pivotTableName: 'PT1',
        fieldName: 'Region',
        sortBy: 'Descending',
      }).success
    ).toBe(true);
    expect(asZod(schema).safeParse({ pivotTableName: 'PT1', fieldName: 'Region' }).success).toBe(
      false
    );
  });

  it('apply_pivot_manual_filter requires pivotTableName + fieldName + selectedItems', () => {
    const schema = excelTools.apply_pivot_manual_filter.inputSchema;
    expect(
      asZod(schema).safeParse({
        pivotTableName: 'PT1',
        fieldName: 'Region',
        selectedItems: ['North', 'South'],
      }).success
    ).toBe(true);
    expect(asZod(schema).safeParse({ pivotTableName: 'PT1', fieldName: 'Region' }).success).toBe(
      false
    );
  });

  it('sort_pivot_field_values requires pivotTableName + fieldName + sortBy + valuesHierarchyName', () => {
    const schema = excelTools.sort_pivot_field_values.inputSchema;
    expect(
      asZod(schema).safeParse({
        pivotTableName: 'PT1',
        fieldName: 'Region',
        sortBy: 'Descending',
        valuesHierarchyName: 'Sales',
      }).success
    ).toBe(true);
    expect(
      asZod(schema).safeParse({
        pivotTableName: 'PT1',
        fieldName: 'Region',
        sortBy: 'Descending',
      }).success
    ).toBe(false);
  });

  it('set_pivot_field_show_all_items requires pivotTableName + fieldName + showAllItems', () => {
    const schema = excelTools.set_pivot_field_show_all_items.inputSchema;
    expect(
      asZod(schema).safeParse({
        pivotTableName: 'PT1',
        fieldName: 'Region',
        showAllItems: false,
      }).success
    ).toBe(true);
    expect(asZod(schema).safeParse({ pivotTableName: 'PT1', fieldName: 'Region' }).success).toBe(
      false
    );
  });
});

describe('tool schemas — workbook tools', () => {
  it('recalculate_workbook accepts empty or optional recalcType', () => {
    const schema = excelTools.recalculate_workbook.inputSchema;
    expect(asZod(schema).safeParse({}).success).toBe(true);
    expect(asZod(schema).safeParse({ recalcType: 'Full' }).success).toBe(true);
  });
});

describe('tool schemas — pivot table tools', () => {
  it('list_pivot_tables accepts optional sheetName', () => {
    const schema = excelTools.list_pivot_tables.inputSchema;
    expect(asZod(schema).safeParse({}).success).toBe(true);
    expect(asZod(schema).safeParse({ sheetName: 'Data' }).success).toBe(true);
  });

  it('refresh_pivot_table requires pivotTableName', () => {
    const schema = excelTools.refresh_pivot_table.inputSchema;
    expect(asZod(schema).safeParse({ pivotTableName: 'PT1' }).success).toBe(true);
    expect(asZod(schema).safeParse({}).success).toBe(false);
  });

  it('delete_pivot_table requires pivotTableName', () => {
    const schema = excelTools.delete_pivot_table.inputSchema;
    expect(asZod(schema).safeParse({ pivotTableName: 'PT1' }).success).toBe(true);
    expect(asZod(schema).safeParse({}).success).toBe(false);
  });

  it('create_pivot_table requires name + sourceAddress + destinationAddress + rowFields + valueFields', () => {
    const schema = excelTools.create_pivot_table.inputSchema;
    expect(
      asZod(schema).safeParse({
        name: 'SalesPivot',
        sourceAddress: 'A1:D100',
        destinationAddress: 'Sheet2!A1',
        rowFields: ['Region'],
        valueFields: ['Sales'],
      }).success
    ).toBe(true);
    expect(asZod(schema).safeParse({ name: 'PT1' }).success).toBe(false);
  });
});

describe('tool schemas — workbook tools', () => {
  it('get_workbook_info accepts empty object', () => {
    const schema = excelTools.get_workbook_info.inputSchema;
    expect(asZod(schema).safeParse({}).success).toBe(true);
  });

  it('get_selected_range accepts empty object', () => {
    const schema = excelTools.get_selected_range.inputSchema;
    expect(asZod(schema).safeParse({}).success).toBe(true);
  });

  it('define_named_range requires name + address', () => {
    const schema = excelTools.define_named_range.inputSchema;
    expect(asZod(schema).safeParse({ name: 'Revenue', address: 'B2:B100' }).success).toBe(true);
    expect(asZod(schema).safeParse({ name: 'Revenue' }).success).toBe(false);
  });

  it('list_queries accepts empty object', () => {
    const schema = excelTools.list_queries.inputSchema;
    expect(asZod(schema).safeParse({}).success).toBe(true);
  });

  it('get_query requires queryName', () => {
    const schema = excelTools.get_query.inputSchema;
    expect(asZod(schema).safeParse({ queryName: 'SalesQuery' }).success).toBe(true);
    expect(asZod(schema).safeParse({}).success).toBe(false);
  });

  it('get_query_count accepts empty object', () => {
    const schema = excelTools.get_query_count.inputSchema;
    expect(asZod(schema).safeParse({}).success).toBe(true);
  });

  it('save_workbook accepts empty or optional saveBehavior', () => {
    const schema = excelTools.save_workbook.inputSchema;
    expect(asZod(schema).safeParse({}).success).toBe(true);
    expect(asZod(schema).safeParse({ saveBehavior: 'Prompt' }).success).toBe(true);
  });

  it('get_workbook_properties accepts empty object', () => {
    const schema = excelTools.get_workbook_properties.inputSchema;
    expect(asZod(schema).safeParse({}).success).toBe(true);
  });

  it('set_workbook_properties accepts any supported property subset', () => {
    const schema = excelTools.set_workbook_properties.inputSchema;
    expect(asZod(schema).safeParse({ title: 'Quarterly Report' }).success).toBe(true);
    expect(asZod(schema).safeParse({ author: 'Analyst', revisionNumber: 2 }).success).toBe(true);
    expect(asZod(schema).safeParse({}).success).toBe(true);
  });

  it('get_workbook_protection accepts empty object', () => {
    const schema = excelTools.get_workbook_protection.inputSchema;
    expect(asZod(schema).safeParse({}).success).toBe(true);
  });

  it('protect_workbook and unprotect_workbook accept optional password', () => {
    const protectSchema = excelTools.protect_workbook.inputSchema;
    const unprotectSchema = excelTools.unprotect_workbook.inputSchema;
    expect(asZod(protectSchema).safeParse({}).success).toBe(true);
    expect(asZod(protectSchema).safeParse({ password: 'secret' }).success).toBe(true);
    expect(asZod(unprotectSchema).safeParse({}).success).toBe(true);
    expect(asZod(unprotectSchema).safeParse({ password: 'secret' }).success).toBe(true);
  });

  it('refresh_data_connections accepts empty object', () => {
    const schema = excelTools.refresh_data_connections.inputSchema;
    expect(asZod(schema).safeParse({}).success).toBe(true);
  });
});

describe('tool schemas — comment tools', () => {
  it('add_comment requires cellAddress + text', () => {
    const schema = excelTools.add_comment.inputSchema;
    expect(asZod(schema).safeParse({ cellAddress: 'A1', text: 'Review this' }).success).toBe(true);
    expect(asZod(schema).safeParse({ cellAddress: 'A1' }).success).toBe(false);
  });

  it('list_comments accepts optional sheetName', () => {
    const schema = excelTools.list_comments.inputSchema;
    expect(asZod(schema).safeParse({}).success).toBe(true);
    expect(asZod(schema).safeParse({ sheetName: 'Sheet1' }).success).toBe(true);
  });

  it('edit_comment requires cellAddress + newText', () => {
    const schema = excelTools.edit_comment.inputSchema;
    expect(asZod(schema).safeParse({ cellAddress: 'A1', newText: 'Updated' }).success).toBe(true);
    expect(asZod(schema).safeParse({ cellAddress: 'A1' }).success).toBe(false);
  });

  it('delete_comment requires cellAddress', () => {
    const schema = excelTools.delete_comment.inputSchema;
    expect(asZod(schema).safeParse({ cellAddress: 'A1' }).success).toBe(true);
    expect(asZod(schema).safeParse({}).success).toBe(false);
  });
});
