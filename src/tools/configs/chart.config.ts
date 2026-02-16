/**
 * Chart tool configs — 6 tools for creating and managing charts.
 *
 * Fixes applied (from tool audit):
 *   - list_charts: description fixed to say "on a worksheet" (not "in the workbook"),
 *     and sheetName param description clarified
 */

import type { ToolConfig } from '../codegen';
import { getSheet } from '../codegen';

export const chartConfigs: readonly ToolConfig[] = [
  {
    name: 'list_charts',
    description:
      "List all charts on a worksheet. Returns each chart's name, type, and title. Uses the active sheet if no sheet is specified.",
    params: {
      sheetName: {
        type: 'string',
        required: false,
        description: 'Worksheet name. Uses active sheet if omitted.',
      },
    },
    execute: async (context, args) => {
      const sheet = getSheet(context, args.sheetName as string | undefined);
      const charts = sheet.charts;
      charts.load('items');
      await context.sync();

      for (const chart of charts.items) {
        chart.load(['name', 'chartType']);
        chart.title.load('text');
      }
      await context.sync();

      const result = charts.items.map(chart => ({
        name: chart.name,
        chartType: chart.chartType,
        title: chart.title?.text ?? '',
      }));
      return { charts: result, count: result.length };
    },
  },

  {
    name: 'create_chart',
    description:
      'Create a new chart from a data range and place it on the same worksheet. The data range should include headers for proper axis labels and legend.',
    params: {
      dataRange: {
        type: 'string',
        description: 'Range address of the source data (e.g., "A1:D10")',
      },
      chartType: {
        type: 'string',
        description: 'Type of chart to create',
        enum: [
          'ColumnClustered',
          'ColumnStacked',
          'BarClustered',
          'BarStacked',
          'Line',
          'LineMarkers',
          'Pie',
          'Doughnut',
          'Area',
          'XYScatter',
        ],
      },
      title: { type: 'string', required: false, description: 'Optional chart title' },
      sheetName: {
        type: 'string',
        required: false,
        description: 'Optional worksheet name for the source data.',
      },
    },
    execute: async (context, args) => {
      const sheet = getSheet(context, args.sheetName as string | undefined);
      const range = sheet.getRange(args.dataRange as string);
      const chartType = args.chartType as string;
      const title = args.title as string | undefined;
      const chart = sheet.charts.add(chartType as Excel.ChartType, range, Excel.ChartSeriesBy.auto);
      if (title) chart.title.text = title;
      chart.load(['name', 'chartType']);
      await context.sync();
      return {
        name: chart.name,
        chartType: chart.chartType,
        title: title ?? '',
        dataRange: args.dataRange,
      };
    },
  },

  {
    name: 'delete_chart',
    description: 'Delete a chart from the worksheet.',
    params: {
      chartName: { type: 'string', description: 'Name of the chart to delete (from list_charts)' },
      sheetName: {
        type: 'string',
        required: false,
        description: 'Optional worksheet name where the chart is located.',
      },
    },
    execute: async (context, args) => {
      const sheet = getSheet(context, args.sheetName as string | undefined);
      const chart = sheet.charts.getItem(args.chartName as string);
      chart.delete();
      await context.sync();
      return { deleted: args.chartName };
    },
  },

  // ─── Chart Properties ────────────────────────────────────

  {
    name: 'set_chart_title',
    description: 'Set or change the title of a chart.',
    params: {
      chartName: { type: 'string', description: 'Name of the chart' },
      title: { type: 'string', description: 'New title text' },
      sheetName: {
        type: 'string',
        required: false,
        description: 'Optional worksheet name.',
      },
    },
    execute: async (context, args) => {
      const sheet = getSheet(context, args.sheetName as string | undefined);
      const chart = sheet.charts.getItem(args.chartName as string);
      chart.title.text = args.title as string;
      chart.load('name');
      await context.sync();
      return { chartName: args.chartName, title: args.title };
    },
  },

  {
    name: 'set_chart_type',
    description: 'Change the chart type (e.g., from column to pie).',
    params: {
      chartName: { type: 'string', description: 'Name of the chart' },
      chartType: {
        type: 'string',
        description: 'New chart type',
        enum: [
          'ColumnClustered',
          'ColumnStacked',
          'BarClustered',
          'BarStacked',
          'Line',
          'LineMarkers',
          'Pie',
          'Doughnut',
          'Area',
          'XYScatter',
        ],
      },
      sheetName: {
        type: 'string',
        required: false,
        description: 'Optional worksheet name.',
      },
    },
    execute: async (context, args) => {
      const sheet = getSheet(context, args.sheetName as string | undefined);
      const chart = sheet.charts.getItem(args.chartName as string);
      chart.chartType = args.chartType as Excel.ChartType;
      chart.load('chartType');
      await context.sync();
      return { chartName: args.chartName, chartType: chart.chartType };
    },
  },

  {
    name: 'set_chart_data_source',
    description: 'Change the data range that a chart is based on.',
    params: {
      chartName: { type: 'string', description: 'Name of the chart' },
      dataRange: {
        type: 'string',
        description: 'New data range address (e.g., "B1:D20")',
      },
      sheetName: {
        type: 'string',
        required: false,
        description: 'Optional worksheet name.',
      },
    },
    execute: async (context, args) => {
      const sheet = getSheet(context, args.sheetName as string | undefined);
      const chart = sheet.charts.getItem(args.chartName as string);
      const range = sheet.getRange(args.dataRange as string);
      chart.setData(range);
      await context.sync();
      return { chartName: args.chartName, dataRange: args.dataRange, updated: true };
    },
  },
];
