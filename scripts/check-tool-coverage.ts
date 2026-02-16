#!/usr/bin/env tsx
/**
 * Excel API Coverage Checker
 *
 * Parses @types/office-js to extract key Excel API methods/properties,
 * then compares against tool configs to flag uncovered capabilities.
 *
 * Usage: npx tsx scripts/check-tool-coverage.ts [--json] [--verbose]
 */

import * as fs from 'fs';
import * as path from 'path';

// ─── Configuration ─────────────────────────────────────────────────
// Key Excel API classes the AI assistant should be able to interact with.
// Each entry lists the class name and the methods/properties we consider
// actionable (i.e., things a user might ask the AI to do).

interface ApiCapability {
  /** Name of the Excel JS API class */
  className: string;
  /** Human-readable category */
  category: string;
  /** Methods and properties that represent user-facing capabilities */
  capabilities: {
    /** Method or property name from the API */
    name: string;
    /** What this capability does (for the report) */
    description: string;
    /** Priority: HIGH = common user request, MEDIUM = good to have, LOW = niche */
    priority: 'HIGH' | 'MEDIUM' | 'LOW';
    /** Tool names that cover this capability (matched against manifest) */
    coveredBy?: string[];
  }[];
}

// This is the authoritative list of Excel JS API capabilities we track.
// Add new entries here when new API features become available.
const API_CAPABILITIES: ApiCapability[] = [
  {
    className: 'Range',
    category: 'Range Operations',
    capabilities: [
      {
        name: 'values (get)',
        description: 'Read cell values from a range',
        priority: 'HIGH',
        coveredBy: ['get_range_values'],
      },
      {
        name: 'values (set)',
        description: 'Write cell values to a range',
        priority: 'HIGH',
        coveredBy: ['set_range_values'],
      },
      {
        name: 'getUsedRange()',
        description: 'Get the used range on a sheet',
        priority: 'HIGH',
        coveredBy: ['get_used_range'],
      },
      {
        name: 'clear()',
        description: 'Clear contents and formatting',
        priority: 'HIGH',
        coveredBy: ['clear_range'],
      },
      {
        name: 'format',
        description: 'Apply formatting (bold, color, etc.)',
        priority: 'HIGH',
        coveredBy: ['format_range'],
      },
      {
        name: 'numberFormat',
        description: 'Set number format codes',
        priority: 'HIGH',
        coveredBy: ['set_number_format'],
      },
      {
        name: 'getEntireColumn().format.autofitColumns()',
        description: 'Auto-fit column widths',
        priority: 'MEDIUM',
        coveredBy: ['auto_fit_columns'],
      },
      {
        name: 'getEntireRow().format.autofitRows()',
        description: 'Auto-fit row heights',
        priority: 'MEDIUM',
        coveredBy: ['auto_fit_rows'],
      },
      {
        name: 'formulas (get)',
        description: 'Read formulas from cells',
        priority: 'HIGH',
        coveredBy: ['get_range_formulas'],
      },
      {
        name: 'formulas (set)',
        description: 'Write formulas to cells',
        priority: 'HIGH',
        coveredBy: ['set_range_formulas'],
      },
      {
        name: 'sort.apply()',
        description: 'Sort a range',
        priority: 'HIGH',
        coveredBy: ['sort_range'],
      },
      {
        name: 'copyFrom()',
        description: 'Copy range to another location',
        priority: 'HIGH',
        coveredBy: ['copy_range'],
      },
      {
        name: 'find()',
        description: 'Find values in a range',
        priority: 'HIGH',
        coveredBy: ['find_values'],
      },
      {
        name: 'replaceAll()',
        description: 'Find and replace values',
        priority: 'HIGH',
        coveredBy: ['replace_values'],
      },
      {
        name: 'insert()',
        description: 'Insert cells, shifting others',
        priority: 'MEDIUM',
        coveredBy: ['insert_range'],
      },
      {
        name: 'delete()',
        description: 'Delete cells, shifting others',
        priority: 'MEDIUM',
        coveredBy: ['delete_range'],
      },
      {
        name: 'merge()',
        description: 'Merge cells',
        priority: 'MEDIUM',
        coveredBy: ['merge_cells'],
      },
      {
        name: 'unmerge()',
        description: 'Unmerge cells',
        priority: 'MEDIUM',
        coveredBy: ['unmerge_cells'],
      },
      {
        name: 'removeDuplicates()',
        description: 'Remove duplicate rows',
        priority: 'HIGH',
        coveredBy: ['remove_duplicates'],
      },
      {
        name: 'hyperlink',
        description: 'Set/get hyperlinks on cells',
        priority: 'HIGH',
        coveredBy: ['set_hyperlink'],
      },
      {
        name: 'rowHidden / columnHidden',
        description: 'Hide/show rows and columns',
        priority: 'HIGH',
        coveredBy: ['toggle_row_column_visibility'],
      },
      {
        name: 'group()',
        description: 'Group rows or columns for outlining',
        priority: 'MEDIUM',
        coveredBy: ['group_rows_columns'],
      },
      {
        name: 'ungroup()',
        description: 'Ungroup rows or columns',
        priority: 'MEDIUM',
        coveredBy: ['ungroup_rows_columns'],
      },
      {
        name: 'showGroupDetails()',
        description: 'Expand/collapse grouped rows/columns',
        priority: 'LOW',
      },
      {
        name: 'dataValidation',
        description: 'Set data validation rules',
        priority: 'HIGH',
        coveredBy: [
          'set_list_validation',
          'set_number_validation',
          'set_date_validation',
          'set_text_length_validation',
          'set_custom_validation',
        ],
      },
      {
        name: 'conditionalFormats.add()',
        description: 'Add conditional formatting',
        priority: 'HIGH',
        coveredBy: [
          'add_color_scale',
          'add_data_bar',
          'add_cell_value_format',
          'add_top_bottom_format',
          'add_contains_text_format',
          'add_custom_format',
        ],
      },
      {
        name: 'format.borders',
        description: 'Set cell borders',
        priority: 'MEDIUM',
        coveredBy: ['set_cell_borders'],
      },
      {
        name: 'format.columnWidth / rowHeight',
        description: 'Set specific column width or row height',
        priority: 'LOW',
      },
      { name: 'format.wrapText', description: 'Enable/disable text wrapping', priority: 'LOW' },
      {
        name: 'getSpecialCells()',
        description: 'Get cells with blanks, formulas, constants, etc.',
        priority: 'LOW',
      },
      { name: 'calculate()', description: 'Force range recalculation', priority: 'LOW' },
      {
        name: 'getDirectPrecedents()',
        description: 'Get cells that a formula depends on',
        priority: 'LOW',
      },
      {
        name: 'getDirectDependents()',
        description: 'Get cells that depend on this cell',
        priority: 'LOW',
      },
      { name: 'autoFill()', description: 'Auto-fill a pattern across cells', priority: 'MEDIUM' },
      { name: 'flashFill()', description: 'Flash fill based on pattern', priority: 'LOW' },
    ],
  },
  {
    className: 'Worksheet',
    category: 'Sheet Management',
    capabilities: [
      {
        name: 'worksheets.items',
        description: 'List all worksheets',
        priority: 'HIGH',
        coveredBy: ['list_sheets'],
      },
      {
        name: 'worksheets.add()',
        description: 'Create a new worksheet',
        priority: 'HIGH',
        coveredBy: ['create_sheet'],
      },
      {
        name: 'name (set)',
        description: 'Rename a worksheet',
        priority: 'HIGH',
        coveredBy: ['rename_sheet'],
      },
      {
        name: 'delete()',
        description: 'Delete a worksheet',
        priority: 'HIGH',
        coveredBy: ['delete_sheet'],
      },
      {
        name: 'activate()',
        description: 'Switch to a worksheet',
        priority: 'HIGH',
        coveredBy: ['activate_sheet'],
      },
      {
        name: 'freezePanes',
        description: 'Freeze rows/columns for scrolling',
        priority: 'HIGH',
        coveredBy: ['freeze_panes'],
      },
      {
        name: 'protection.protect()',
        description: 'Protect a worksheet',
        priority: 'HIGH',
        coveredBy: ['protect_sheet'],
      },
      {
        name: 'protection.unprotect()',
        description: 'Unprotect a worksheet',
        priority: 'HIGH',
        coveredBy: ['unprotect_sheet'],
      },
      {
        name: 'visibility',
        description: 'Show/hide/very-hide a worksheet',
        priority: 'HIGH',
        coveredBy: ['set_sheet_visibility'],
      },
      {
        name: 'tabColor',
        description: 'Set worksheet tab color',
        priority: 'MEDIUM',
        coveredBy: ['set_sheet_visibility'],
      },
      {
        name: 'copy()',
        description: 'Copy a worksheet',
        priority: 'HIGH',
        coveredBy: ['copy_sheet'],
      },
      {
        name: 'position',
        description: 'Move/reorder worksheets',
        priority: 'MEDIUM',
        coveredBy: ['move_sheet'],
      },
      { name: 'showGridlines', description: 'Toggle gridline visibility', priority: 'LOW' },
      {
        name: 'showHeadings',
        description: 'Toggle row/column heading visibility',
        priority: 'LOW',
      },
      {
        name: 'pageLayout',
        description: 'Page setup (orientation, margins, paper size)',
        priority: 'MEDIUM',
        coveredBy: ['set_page_layout'],
      },
      {
        name: 'getCell()',
        description: 'Get a specific cell by row/column index',
        priority: 'LOW',
      },
      {
        name: 'horizontalPageBreaks / verticalPageBreaks',
        description: 'Manage page breaks',
        priority: 'LOW',
      },
    ],
  },
  {
    className: 'Table',
    category: 'Table Operations',
    capabilities: [
      {
        name: 'tables.items',
        description: 'List all tables',
        priority: 'HIGH',
        coveredBy: ['list_tables'],
      },
      {
        name: 'tables.add()',
        description: 'Create a table',
        priority: 'HIGH',
        coveredBy: ['create_table'],
      },
      {
        name: 'rows.add()',
        description: 'Add rows to a table',
        priority: 'HIGH',
        coveredBy: ['add_table_rows'],
      },
      {
        name: 'getRange().values',
        description: 'Read table data',
        priority: 'HIGH',
        coveredBy: ['get_table_data'],
      },
      {
        name: 'delete()',
        description: 'Delete a table',
        priority: 'HIGH',
        coveredBy: ['delete_table'],
      },
      {
        name: 'sort.apply()',
        description: 'Sort a table',
        priority: 'HIGH',
        coveredBy: ['sort_table'],
      },
      {
        name: 'columns[].filter',
        description: 'Filter table columns',
        priority: 'HIGH',
        coveredBy: ['filter_table'],
      },
      {
        name: 'autoFilter.clearCriteria()',
        description: 'Clear all table filters',
        priority: 'HIGH',
        coveredBy: ['clear_table_filters'],
      },
      {
        name: 'columns.add()',
        description: 'Add a column to a table',
        priority: 'MEDIUM',
        coveredBy: ['add_table_column'],
      },
      {
        name: 'columns.delete()',
        description: 'Delete a table column',
        priority: 'MEDIUM',
        coveredBy: ['delete_table_column'],
      },
      {
        name: 'convertToRange()',
        description: 'Convert table back to plain range',
        priority: 'MEDIUM',
        coveredBy: ['convert_table_to_range'],
      },
      { name: 'resize()', description: 'Resize table to new range', priority: 'LOW' },
      { name: 'style', description: 'Apply table style', priority: 'LOW' },
      { name: 'showHeaders / showTotals', description: 'Toggle header/total row', priority: 'LOW' },
    ],
  },
  {
    className: 'Chart',
    category: 'Chart Operations',
    capabilities: [
      {
        name: 'charts.items',
        description: 'List all charts',
        priority: 'HIGH',
        coveredBy: ['list_charts'],
      },
      {
        name: 'charts.add()',
        description: 'Create a chart',
        priority: 'HIGH',
        coveredBy: ['create_chart'],
      },
      {
        name: 'delete()',
        description: 'Delete a chart',
        priority: 'HIGH',
        coveredBy: ['delete_chart'],
      },
      {
        name: 'title',
        description: 'Set chart title',
        priority: 'MEDIUM',
        coveredBy: ['set_chart_title'],
      },
      {
        name: 'chartType',
        description: 'Change chart type',
        priority: 'MEDIUM',
        coveredBy: ['set_chart_type'],
      },
      {
        name: 'setData()',
        description: 'Change chart data source',
        priority: 'MEDIUM',
        coveredBy: ['set_chart_data_source'],
      },
      { name: 'legend', description: 'Configure chart legend', priority: 'LOW' },
      { name: 'axes', description: 'Configure chart axes', priority: 'LOW' },
      { name: 'series', description: 'Manage chart data series', priority: 'LOW' },
      { name: 'setPosition()', description: 'Move/resize chart', priority: 'LOW' },
    ],
  },
  {
    className: 'Workbook',
    category: 'Workbook Operations',
    capabilities: [
      {
        name: 'worksheets + properties',
        description: 'Get workbook info',
        priority: 'HIGH',
        coveredBy: ['get_workbook_info'],
      },
      {
        name: 'getSelectedRange()',
        description: 'Get the currently selected range',
        priority: 'HIGH',
        coveredBy: ['get_selected_range'],
      },
      {
        name: 'names.add()',
        description: 'Define a named range',
        priority: 'MEDIUM',
        coveredBy: ['define_named_range'],
      },
      {
        name: 'names.items',
        description: 'List named ranges',
        priority: 'MEDIUM',
        coveredBy: ['list_named_ranges'],
      },
      {
        name: 'properties',
        description: 'Get/set workbook properties (title, author)',
        priority: 'LOW',
      },
      { name: 'save()', description: 'Save the workbook', priority: 'LOW' },
      {
        name: 'application.calculate()',
        description: 'Force full recalculation',
        priority: 'LOW',
        coveredBy: ['recalculate_workbook'],
      },
    ],
  },
  {
    className: 'Comment',
    category: 'Comments',
    capabilities: [
      {
        name: 'comments.add()',
        description: 'Add a comment to a cell',
        priority: 'HIGH',
        coveredBy: ['add_comment'],
      },
      {
        name: 'comments.items',
        description: 'List comments',
        priority: 'HIGH',
        coveredBy: ['list_comments'],
      },
      {
        name: 'content (set)',
        description: 'Edit a comment',
        priority: 'HIGH',
        coveredBy: ['edit_comment'],
      },
      {
        name: 'delete()',
        description: 'Delete a comment',
        priority: 'HIGH',
        coveredBy: ['delete_comment'],
      },
      { name: 'replies.add()', description: 'Add a threaded reply', priority: 'LOW' },
      { name: 'resolved', description: 'Mark comment as resolved', priority: 'LOW' },
    ],
  },
  {
    className: 'PivotTable',
    category: 'Pivot Tables',
    capabilities: [
      {
        name: 'pivotTables.items',
        description: 'List pivot tables',
        priority: 'HIGH',
        coveredBy: ['list_pivot_tables'],
      },
      {
        name: 'pivotTables.add()',
        description: 'Create a pivot table',
        priority: 'HIGH',
        coveredBy: ['create_pivot_table'],
      },
      {
        name: 'refresh()',
        description: 'Refresh a pivot table',
        priority: 'HIGH',
        coveredBy: ['refresh_pivot_table'],
      },
      {
        name: 'delete()',
        description: 'Delete a pivot table',
        priority: 'HIGH',
        coveredBy: ['delete_pivot_table'],
      },
      {
        name: 'rowHierarchies.add/remove()',
        description: 'Add/remove row fields',
        priority: 'MEDIUM',
        coveredBy: ['add_pivot_field', 'remove_pivot_field'],
      },
      {
        name: 'dataHierarchies.add/remove()',
        description: 'Add/remove data (value) fields',
        priority: 'MEDIUM',
        coveredBy: ['add_pivot_field', 'remove_pivot_field'],
      },
      {
        name: 'columnHierarchies.add/remove()',
        description: 'Add/remove column fields',
        priority: 'LOW',
      },
      {
        name: 'filterHierarchies.add/remove()',
        description: 'Add/remove filter fields',
        priority: 'LOW',
      },
      {
        name: 'layout',
        description: 'Set pivot table layout (compact, outline, tabular)',
        priority: 'LOW',
      },
    ],
  },
  {
    className: 'Shape',
    category: 'Shapes & Objects',
    capabilities: [
      { name: 'shapes.addGeometricShape()', description: 'Add a shape', priority: 'LOW' },
      { name: 'shapes.addTextBox()', description: 'Add a text box', priority: 'LOW' },
      { name: 'shapes.addImage()', description: 'Add an image', priority: 'LOW' },
      { name: 'delete()', description: 'Delete a shape', priority: 'LOW' },
      { name: 'left/top/width/height', description: 'Position and resize shapes', priority: 'LOW' },
    ],
  },
];

// ─── Manifest Loading ──────────────────────────────────────────────

interface ManifestTool {
  name: string;
  description: string;
  params: Record<string, unknown>;
}

function loadManifest(): ManifestTool[] {
  const manifestPath = path.resolve(__dirname, '../src/tools/tools-manifest.json');
  const data = JSON.parse(fs.readFileSync(manifestPath, 'utf-8'));
  return data.tools;
}

// ─── Analysis ──────────────────────────────────────────────────────

interface CoverageResult {
  totalCapabilities: number;
  covered: number;
  uncovered: number;
  coveragePercent: number;
  byPriority: Record<string, { total: number; covered: number; percent: number }>;
  byCategory: {
    category: string;
    total: number;
    covered: number;
    percent: number;
    gaps: { name: string; description: string; priority: string }[];
  }[];
  toolCount: number;
}

function analyze(verbose: boolean): CoverageResult {
  const manifest = loadManifest();
  const toolNames = new Set(manifest.map(t => t.name));

  let totalCapabilities = 0;
  let covered = 0;
  const priorityCounts: Record<string, { total: number; covered: number }> = {
    HIGH: { total: 0, covered: 0 },
    MEDIUM: { total: 0, covered: 0 },
    LOW: { total: 0, covered: 0 },
  };

  const categories: CoverageResult['byCategory'] = [];

  for (const api of API_CAPABILITIES) {
    const gaps: { name: string; description: string; priority: string }[] = [];
    let catTotal = 0;
    let catCovered = 0;

    for (const cap of api.capabilities) {
      totalCapabilities++;
      catTotal++;
      priorityCounts[cap.priority].total++;

      const isCovered = cap.coveredBy?.some(t => toolNames.has(t)) ?? false;
      if (isCovered) {
        covered++;
        catCovered++;
        priorityCounts[cap.priority].covered++;
      } else {
        gaps.push({ name: cap.name, description: cap.description, priority: cap.priority });
      }
    }

    categories.push({
      category: api.category,
      total: catTotal,
      covered: catCovered,
      percent: Math.round((catCovered / catTotal) * 100),
      gaps,
    });
  }

  const uncovered = totalCapabilities - covered;

  return {
    totalCapabilities,
    covered,
    uncovered,
    coveragePercent: Math.round((covered / totalCapabilities) * 100),
    byPriority: Object.fromEntries(
      Object.entries(priorityCounts).map(([p, c]) => [
        p,
        { total: c.total, covered: c.covered, percent: Math.round((c.covered / c.total) * 100) },
      ])
    ),
    byCategory: categories,
    toolCount: toolNames.size,
  };
}

// ─── Reporting ─────────────────────────────────────────────────────

function printReport(result: CoverageResult, verbose: boolean): void {
  console.log('\n╔══════════════════════════════════════════════════════════════╗');
  console.log('║           Excel Tool Coverage Report                       ║');
  console.log('╚══════════════════════════════════════════════════════════════╝\n');

  console.log(`Tools in manifest: ${result.toolCount}`);
  console.log(`API capabilities tracked: ${result.totalCapabilities}`);
  console.log(
    `Coverage: ${result.covered}/${result.totalCapabilities} (${result.coveragePercent}%)\n`
  );

  // Priority breakdown
  console.log('─── Coverage by Priority ─────────────────────────────────────');
  for (const [priority, data] of Object.entries(result.byPriority)) {
    const bar = makeBar(data.percent);
    console.log(`  ${priority.padEnd(7)} ${bar} ${data.covered}/${data.total} (${data.percent}%)`);
  }
  console.log();

  // Category breakdown
  console.log('─── Coverage by Category ─────────────────────────────────────');
  for (const cat of result.byCategory) {
    const bar = makeBar(cat.percent);
    console.log(
      `  ${cat.category.padEnd(22)} ${bar} ${cat.covered}/${cat.total} (${cat.percent}%)`
    );
    if (verbose || cat.gaps.some(g => g.priority === 'HIGH')) {
      for (const gap of cat.gaps) {
        if (verbose || gap.priority === 'HIGH') {
          const icon = gap.priority === 'HIGH' ? '🔴' : gap.priority === 'MEDIUM' ? '🟡' : '⚪';
          console.log(`    ${icon} ${gap.description} (${gap.name})`);
        }
      }
    }
  }
  console.log();

  // Summary of HIGH priority gaps
  const highGaps = result.byCategory.flatMap(c => c.gaps.filter(g => g.priority === 'HIGH'));
  if (highGaps.length > 0) {
    console.log('─── HIGH Priority Gaps ───────────────────────────────────────');
    for (const gap of highGaps) {
      console.log(`  🔴 ${gap.description} (${gap.name})`);
    }
    console.log();
  }

  if (verbose) {
    const medGaps = result.byCategory.flatMap(c => c.gaps.filter(g => g.priority === 'MEDIUM'));
    if (medGaps.length > 0) {
      console.log('─── MEDIUM Priority Gaps ─────────────────────────────────────');
      for (const gap of medGaps) {
        console.log(`  🟡 ${gap.description} (${gap.name})`);
      }
      console.log();
    }
  }
}

function makeBar(percent: number): string {
  const width = 20;
  const filled = Math.round((percent / 100) * width);
  return '█'.repeat(filled) + '░'.repeat(width - filled);
}

// ─── Main ──────────────────────────────────────────────────────────

const args = process.argv.slice(2);
const jsonMode = args.includes('--json');
const verbose = args.includes('--verbose') || args.includes('-v');

const result = analyze(verbose);

if (jsonMode) {
  console.log(JSON.stringify(result, null, 2));
} else {
  printReport(result, verbose);

  // Exit with error if any HIGH priority gaps remain
  const highGaps = result.byCategory.flatMap(c => c.gaps.filter(g => g.priority === 'HIGH'));
  if (highGaps.length > 0) {
    console.log(`⚠️  ${highGaps.length} HIGH priority gap(s) remaining.\n`);
    process.exit(1);
  } else {
    console.log('✅ All HIGH priority capabilities are covered!\n');
  }
}
