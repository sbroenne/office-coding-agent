import { describe, it, expect } from 'vitest';
import Ajv from 'ajv';
import { excelTools, allConfigs } from '@/tools';
import { createTools } from '@/tools/codegen';

/**
 * Tool schema tests — validate that every tool's JSON Schema parameters
 * accepts valid inputs and rejects invalid ones.
 *
 * These DON'T execute the tools (no Excel.run needed), they only
 * exercise the JSON schemas that guard the AI function-call contract.
 */

const ajv = new Ajv({ allErrors: true });

/** Validate data against a JSON Schema */
function validate(schema: unknown, data: unknown): { success: boolean } {
  const fn = ajv.compile(schema as object);
  return { success: !!fn(data) };
}

/** Look up an Excel tool by name */
const toolsByName = Object.fromEntries(excelTools.map(t => [t.name, t]));

/** All 10 consolidated tool names */
const ALL_TOOL_NAMES = [
  'range',
  'range_format',
  'table',
  'chart',
  'sheet',
  'workbook',
  'comment',
  'conditional_format',
  'data_validation',
  'pivot',
] as const;

describe('tool schemas — structural', () => {
  it('excelTools array is non-empty and every tool has parameters and handler', () => {
    expect(excelTools.length).toBeGreaterThan(0);
    for (const t of excelTools) {
      expect(t, `${t.name} is defined`).toBeDefined();
      expect(t.parameters, `${t.name} has parameters`).toBeDefined();
      expect(typeof t.handler, `${t.name} has handler fn`).toBe('function');
    }
  });

  it('excelTools contains exactly the expected tools', () => {
    const actual = excelTools.map(t => t.name).sort();
    const expected = [...ALL_TOOL_NAMES].sort();
    expect(actual).toEqual(expected);
  });

  it('every config produces a valid tool via createTools()', () => {
    for (const configs of allConfigs) {
      const tools = createTools(configs);
      for (const tool of tools) {
        expect(tool, `${tool.name} was created`).toBeDefined();
        expect(tool.parameters, `${tool.name} has parameters`).toBeDefined();
        expect(typeof tool.handler, `${tool.name} has handler`).toBe('function');
      }
    }
  });
});

describe('tool schemas — range', () => {
  it('requires action', () => {
    const schema = toolsByName.range.parameters;
    expect(validate(schema, {}).success).toBe(false);
    expect(validate(schema, { action: 'get_values', address: 'A1:C10' }).success).toBe(true);
  });

  it('accepts get_values with paging params', () => {
    const schema = toolsByName.range.parameters;
    expect(validate(schema, { action: 'get_values', address: 'A1:Z100', maxRows: 20, startRow: 21 }).success).toBe(true);
    expect(validate(schema, { action: 'get_values', address: 'A1:Z100', maxColumns: 5, startColumn: 6 }).success).toBe(true);
  });

  it('accepts get_used without address', () => {
    const schema = toolsByName.range.parameters;
    expect(validate(schema, { action: 'get_used' }).success).toBe(true);
    expect(validate(schema, { action: 'get_used', maxRows: 5 }).success).toBe(true);
  });

  it('accepts set_values with 2D array', () => {
    const schema = toolsByName.range.parameters;
    expect(validate(schema, { action: 'set_values', address: 'A1:B2', values: [[1, 2], [3, 4]] }).success).toBe(true);
  });

  it('accepts get_formulas / set_formulas', () => {
    const schema = toolsByName.range.parameters;
    expect(validate(schema, { action: 'get_formulas', address: 'A1:D10' }).success).toBe(true);
    expect(validate(schema, { action: 'set_formulas', address: 'D1', formulas: [['=SUM(A1:A10)']] }).success).toBe(true);
  });

  it('accepts sort with column', () => {
    const schema = toolsByName.range.parameters;
    expect(validate(schema, { action: 'sort', address: 'A1:C10', column: 0, ascending: false }).success).toBe(true);
  });

  it('accepts copy with destination', () => {
    const schema = toolsByName.range.parameters;
    expect(validate(schema, { action: 'copy', address: 'A1:B5', destinationAddress: 'D1' }).success).toBe(true);
  });

  it('accepts find', () => {
    const schema = toolsByName.range.parameters;
    expect(validate(schema, { action: 'find', searchText: 'hello' }).success).toBe(true);
  });

  it('accepts replace with find+replace params', () => {
    const schema = toolsByName.range.parameters;
    expect(validate(schema, { action: 'replace', find: 'foo', replace: 'bar' }).success).toBe(true);
  });

  it('accepts remove_duplicates with columns', () => {
    const schema = toolsByName.range.parameters;
    expect(validate(schema, { action: 'remove_duplicates', address: 'A1:D100', columns: ['0', '2'] }).success).toBe(true);
  });

  it('accepts merge/unmerge', () => {
    const schema = toolsByName.range.parameters;
    expect(validate(schema, { action: 'merge', address: 'A1:D1' }).success).toBe(true);
    expect(validate(schema, { action: 'unmerge', address: 'A1:D1' }).success).toBe(true);
  });

  it('accepts group/ungroup', () => {
    const schema = toolsByName.range.parameters;
    expect(validate(schema, { action: 'group', address: '3:5' }).success).toBe(true);
    expect(validate(schema, { action: 'ungroup', address: '3:5' }).success).toBe(true);
  });

  it('accepts insert/delete with shift', () => {
    const schema = toolsByName.range.parameters;
    expect(validate(schema, { action: 'insert', address: '3:5', shift: 'down' }).success).toBe(true);
    expect(validate(schema, { action: 'delete', address: '3:5', shift: 'up' }).success).toBe(true);
  });

  it('accepts get_special_cells with cellType', () => {
    const schema = toolsByName.range.parameters;
    expect(validate(schema, { action: 'get_special_cells', address: 'A1:D100', cellType: 'Formulas' }).success).toBe(true);
  });

  it('accepts get_precedents and get_dependents', () => {
    const schema = toolsByName.range.parameters;
    expect(validate(schema, { action: 'get_precedents', address: 'D2' }).success).toBe(true);
    expect(validate(schema, { action: 'get_dependents', address: 'B2' }).success).toBe(true);
  });
});

describe('tool schemas — range_format', () => {
  it('requires action', () => {
    const schema = toolsByName.range_format.parameters;
    expect(validate(schema, {}).success).toBe(false);
  });

  it('accepts format action', () => {
    const schema = toolsByName.range_format.parameters;
    expect(validate(schema, { action: 'format', address: 'A1:B5', bold: true }).success).toBe(true);
  });

  it('accepts set_number_format', () => {
    const schema = toolsByName.range_format.parameters;
    expect(validate(schema, { action: 'set_number_format', address: 'B2:B10', format: '#,##0.00' }).success).toBe(true);
  });

  it('accepts auto_fit', () => {
    const schema = toolsByName.range_format.parameters;
    expect(validate(schema, { action: 'auto_fit', fitTarget: 'columns' }).success).toBe(true);
    expect(validate(schema, { action: 'auto_fit', address: 'A1:C10', fitTarget: 'both' }).success).toBe(true);
  });

  it('accepts set_borders', () => {
    const schema = toolsByName.range_format.parameters;
    expect(validate(schema, { action: 'set_borders', address: 'A1:C10', borderStyle: 'Thin', side: 'EdgeAll' }).success).toBe(true);
  });

  it('accepts set_hyperlink', () => {
    const schema = toolsByName.range_format.parameters;
    expect(validate(schema, { action: 'set_hyperlink', address: 'A1', url: 'https://example.com' }).success).toBe(true);
  });

  it('accepts toggle_visibility', () => {
    const schema = toolsByName.range_format.parameters;
    expect(validate(schema, { action: 'toggle_visibility', address: 'A:C', hidden: true, target: 'columns' }).success).toBe(true);
  });
});

describe('tool schemas — sheet', () => {
  it('requires action', () => {
    const schema = toolsByName.sheet.parameters;
    expect(validate(schema, {}).success).toBe(false);
  });

  it('list accepts action only', () => {
    const schema = toolsByName.sheet.parameters;
    expect(validate(schema, { action: 'list' }).success).toBe(true);
  });

  it('create requires name', () => {
    const schema = toolsByName.sheet.parameters;
    expect(validate(schema, { action: 'create', name: 'NewSheet' }).success).toBe(true);
  });

  it('rename requires currentName + newName', () => {
    const schema = toolsByName.sheet.parameters;
    expect(validate(schema, { action: 'rename', currentName: 'Sheet1', newName: 'Data' }).success).toBe(true);
  });

  it('delete requires name', () => {
    const schema = toolsByName.sheet.parameters;
    expect(validate(schema, { action: 'delete', name: 'Sheet2' }).success).toBe(true);
  });

  it('freeze accepts optional freezeAt', () => {
    const schema = toolsByName.sheet.parameters;
    expect(validate(schema, { action: 'freeze', name: 'Sheet1', freezeAt: 'B3' }).success).toBe(true);
    expect(validate(schema, { action: 'freeze', name: 'Sheet1' }).success).toBe(true);
  });

  it('set_page_layout accepts orientation', () => {
    const schema = toolsByName.sheet.parameters;
    expect(validate(schema, { action: 'set_page_layout', name: 'Sheet1', orientation: 'Landscape' }).success).toBe(true);
  });

  it('set_gridlines requires showGridlines', () => {
    const schema = toolsByName.sheet.parameters;
    expect(validate(schema, { action: 'set_gridlines', name: 'Sheet1', showGridlines: false }).success).toBe(true);
  });
});

describe('tool schemas — workbook', () => {
  it('requires action', () => {
    const schema = toolsByName.workbook.parameters;
    expect(validate(schema, {}).success).toBe(false);
  });

  it('get_info accepts action only', () => {
    const schema = toolsByName.workbook.parameters;
    expect(validate(schema, { action: 'get_info' }).success).toBe(true);
  });

  it('set_properties accepts optional fields', () => {
    const schema = toolsByName.workbook.parameters;
    expect(validate(schema, { action: 'set_properties', title: 'My Report' }).success).toBe(true);
  });

  it('define_named_range requires name + address', () => {
    const schema = toolsByName.workbook.parameters;
    expect(validate(schema, { action: 'define_named_range', name: 'Sales', address: 'A1:D100' }).success).toBe(true);
  });

  it('get_query requires queryName', () => {
    const schema = toolsByName.workbook.parameters;
    expect(validate(schema, { action: 'get_query', queryName: 'SalesQuery' }).success).toBe(true);
  });
});

describe('tool schemas — table', () => {
  it('requires action', () => {
    const schema = toolsByName.table.parameters;
    expect(validate(schema, {}).success).toBe(false);
  });

  it('list accepts optional sheetName', () => {
    const schema = toolsByName.table.parameters;
    expect(validate(schema, { action: 'list' }).success).toBe(true);
  });

  it('create requires address', () => {
    const schema = toolsByName.table.parameters;
    expect(validate(schema, { action: 'create', address: 'A1:D10' }).success).toBe(true);
  });

  it('get_data requires tableName', () => {
    const schema = toolsByName.table.parameters;
    expect(validate(schema, { action: 'get_data', tableName: 'Sales' }).success).toBe(true);
  });

  it('sort requires tableName + column', () => {
    const schema = toolsByName.table.parameters;
    expect(validate(schema, { action: 'sort', tableName: 'T1', column: 0 }).success).toBe(true);
  });

  it('filter requires tableName + column + filterValues', () => {
    const schema = toolsByName.table.parameters;
    expect(validate(schema, { action: 'filter', tableName: 'T1', column: 0, filterValues: ['Yes'] }).success).toBe(true);
  });

  it('configure accepts style/showHeaders/showTotals', () => {
    const schema = toolsByName.table.parameters;
    expect(validate(schema, { action: 'configure', tableName: 'T1', style: 'TableStyleMedium2', showTotals: true }).success).toBe(true);
  });
});

describe('tool schemas — chart', () => {
  it('requires action', () => {
    const schema = toolsByName.chart.parameters;
    expect(validate(schema, {}).success).toBe(false);
  });

  it('list accepts action only', () => {
    const schema = toolsByName.chart.parameters;
    expect(validate(schema, { action: 'list' }).success).toBe(true);
  });

  it('create requires dataRange', () => {
    const schema = toolsByName.chart.parameters;
    expect(validate(schema, { action: 'create', dataRange: 'A1:D10', chartType: 'ColumnClustered' }).success).toBe(true);
  });

  it('delete requires chartName', () => {
    const schema = toolsByName.chart.parameters;
    expect(validate(schema, { action: 'delete', chartName: 'Chart1' }).success).toBe(true);
  });

  it('configure accepts title/type/position', () => {
    const schema = toolsByName.chart.parameters;
    expect(validate(schema, { action: 'configure', chartName: 'Chart1', title: 'Sales', chartType: 'Pie' }).success).toBe(true);
    expect(validate(schema, { action: 'configure', chartName: 'Chart1', left: 20, top: 30, width: 400, height: 300 }).success).toBe(true);
  });
});

describe('tool schemas — comment', () => {
  it('requires action', () => {
    const schema = toolsByName.comment.parameters;
    expect(validate(schema, {}).success).toBe(false);
  });

  it('list accepts action only', () => {
    const schema = toolsByName.comment.parameters;
    expect(validate(schema, { action: 'list' }).success).toBe(true);
  });

  it('add requires cellAddress + text', () => {
    const schema = toolsByName.comment.parameters;
    expect(validate(schema, { action: 'add', cellAddress: 'A1', text: 'Note' }).success).toBe(true);
  });

  it('edit requires cellAddress + text', () => {
    const schema = toolsByName.comment.parameters;
    expect(validate(schema, { action: 'edit', cellAddress: 'A1', text: 'Updated' }).success).toBe(true);
  });

  it('delete requires cellAddress', () => {
    const schema = toolsByName.comment.parameters;
    expect(validate(schema, { action: 'delete', cellAddress: 'A1' }).success).toBe(true);
  });
});

describe('tool schemas — conditional_format', () => {
  it('requires action', () => {
    const schema = toolsByName.conditional_format.parameters;
    expect(validate(schema, {}).success).toBe(false);
  });

  it('list accepts address', () => {
    const schema = toolsByName.conditional_format.parameters;
    expect(validate(schema, { action: 'list', address: 'A1:D20' }).success).toBe(true);
    expect(validate(schema, { action: 'list' }).success).toBe(true);
  });

  it('clear requires address', () => {
    const schema = toolsByName.conditional_format.parameters;
    expect(validate(schema, { action: 'clear', address: 'A1:D20' }).success).toBe(true);
  });

  it('add colorScale requires address + type + colors', () => {
    const schema = toolsByName.conditional_format.parameters;
    expect(validate(schema, { action: 'add', address: 'A1:A20', type: 'colorScale', minColor: '#FF0000', maxColor: '#00FF00' }).success).toBe(true);
  });

  it('add cellValue requires address + type + operator + formula1', () => {
    const schema = toolsByName.conditional_format.parameters;
    expect(validate(schema, { action: 'add', address: 'A1:A20', type: 'cellValue', operator: 'GreaterThan', formula1: '100' }).success).toBe(true);
  });
});

describe('tool schemas — data_validation', () => {
  it('requires action + address', () => {
    const schema = toolsByName.data_validation.parameters;
    expect(validate(schema, {}).success).toBe(false);
  });

  it('get requires address', () => {
    const schema = toolsByName.data_validation.parameters;
    expect(validate(schema, { action: 'get', address: 'A1:A100' }).success).toBe(true);
  });

  it('clear requires address', () => {
    const schema = toolsByName.data_validation.parameters;
    expect(validate(schema, { action: 'clear', address: 'A1:A100' }).success).toBe(true);
  });

  it('set list requires type + listValues', () => {
    const schema = toolsByName.data_validation.parameters;
    expect(validate(schema, { action: 'set', address: 'A1:A20', type: 'list', listValues: ['Yes', 'No'] }).success).toBe(true);
  });

  it('set number requires operator + formula1', () => {
    const schema = toolsByName.data_validation.parameters;
    expect(validate(schema, { action: 'set', address: 'B1:B20', type: 'number', operator: 'GreaterThan', formula1: '0' }).success).toBe(true);
  });
});

describe('tool schemas — pivot', () => {
  it('requires action', () => {
    const schema = toolsByName.pivot.parameters;
    expect(validate(schema, {}).success).toBe(false);
  });

  it('list accepts optional sheetName', () => {
    const schema = toolsByName.pivot.parameters;
    expect(validate(schema, { action: 'list' }).success).toBe(true);
  });

  it('create requires name + sourceAddress + destinationAddress', () => {
    const schema = toolsByName.pivot.parameters;
    expect(validate(schema, { action: 'create', name: 'PT1', sourceAddress: 'A1:D100', destinationAddress: 'F1', rowFields: ['Region'], valueFields: ['Sales'] }).success).toBe(true);
  });

  it('delete requires pivotTableName', () => {
    const schema = toolsByName.pivot.parameters;
    expect(validate(schema, { action: 'delete', pivotTableName: 'PT1' }).success).toBe(true);
  });

  it('add_field requires pivotTableName + fieldName + fieldType', () => {
    const schema = toolsByName.pivot.parameters;
    expect(validate(schema, { action: 'add_field', pivotTableName: 'PT1', fieldName: 'Category', fieldType: 'row' }).success).toBe(true);
  });

  it('configure accepts layout options', () => {
    const schema = toolsByName.pivot.parameters;
    expect(validate(schema, { action: 'configure', pivotTableName: 'PT1', layoutType: 'Tabular', showRowGrandTotals: false }).success).toBe(true);
  });

  it('filter accepts label/manual/clear types', () => {
    const schema = toolsByName.pivot.parameters;
    expect(validate(schema, { action: 'filter', pivotTableName: 'PT1', fieldName: 'Region', filterType: 'clear' }).success).toBe(true);
    expect(validate(schema, { action: 'filter', pivotTableName: 'PT1', fieldName: 'Region', filterType: 'manual', selectedItems: ['North', 'South'] }).success).toBe(true);
  });

  it('sort accepts labels mode', () => {
    const schema = toolsByName.pivot.parameters;
    expect(validate(schema, { action: 'sort', pivotTableName: 'PT1', fieldName: 'Region', sortBy: 'Ascending', sortMode: 'labels' }).success).toBe(true);
  });
});
