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

  {
    name: 'set_chart_position',
    description:
      'Position and size a chart. Use startCell/endCell to anchor by range, or left/top/width/height for point-based positioning.',
    params: {
      chartName: { type: 'string', description: 'Name of the chart' },
      startCell: {
        type: 'string',
        required: false,
        description: 'Optional top-left anchor cell address',
      },
      endCell: {
        type: 'string',
        required: false,
        description: 'Optional bottom-right anchor cell address',
      },
      left: { type: 'number', required: false, description: 'Left position in points' },
      top: { type: 'number', required: false, description: 'Top position in points' },
      width: { type: 'number', required: false, description: 'Width in points' },
      height: { type: 'number', required: false, description: 'Height in points' },
      sheetName: {
        type: 'string',
        required: false,
        description: 'Optional worksheet name.',
      },
    },
    execute: async (context, args) => {
      const sheet = getSheet(context, args.sheetName as string | undefined);
      const chart = sheet.charts.getItem(args.chartName as string);

      const startCell = args.startCell as string | undefined;
      const endCell = args.endCell as string | undefined;
      if (startCell) {
        chart.setPosition(startCell, endCell);
      }
      if (args.left !== undefined) chart.left = args.left as number;
      if (args.top !== undefined) chart.top = args.top as number;
      if (args.width !== undefined) chart.width = args.width as number;
      if (args.height !== undefined) chart.height = args.height as number;

      chart.load(['name', 'left', 'top', 'width', 'height']);
      await context.sync();
      return {
        chartName: chart.name,
        left: chart.left,
        top: chart.top,
        width: chart.width,
        height: chart.height,
      };
    },
  },

  {
    name: 'set_chart_legend_visibility',
    description: 'Show or hide the chart legend.',
    params: {
      chartName: { type: 'string', description: 'Name of the chart' },
      visible: { type: 'boolean', description: 'Legend visibility' },
      position: {
        type: 'string',
        required: false,
        description: 'Optional legend position',
        enum: ['Top', 'Bottom', 'Left', 'Right', 'Corner'],
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
      chart.legend.visible = args.visible as boolean;
      if (args.position) {
        chart.legend.position = args.position as Excel.ChartLegendPosition;
      }
      chart.legend.load(['visible', 'position']);
      await context.sync();
      return {
        chartName: args.chartName,
        visible: chart.legend.visible,
        position: chart.legend.position,
      };
    },
  },

  {
    name: 'set_chart_axis_title',
    description: 'Set the axis title text for a chart axis and make the title visible.',
    params: {
      chartName: { type: 'string', description: 'Name of the chart' },
      axisType: {
        type: 'string',
        description: 'Axis type to update',
        enum: ['Category', 'Value', 'Series'],
      },
      title: { type: 'string', description: 'Axis title text' },
      axisGroup: {
        type: 'string',
        required: false,
        description: 'Optional axis group',
        enum: ['Primary', 'Secondary'],
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
      const axis = chart.axes.getItem(
        args.axisType as Excel.ChartAxisType,
        args.axisGroup as Excel.ChartAxisGroup | undefined
      );
      axis.title.visible = true;
      axis.title.text = args.title as string;
      axis.title.load(['visible', 'text']);
      await context.sync();
      return {
        chartName: args.chartName,
        axisType: args.axisType,
        axisGroup: args.axisGroup ?? 'Primary',
        title: axis.title.text,
        titleVisible: axis.title.visible,
      };
    },
  },

  {
    name: 'set_chart_axis_visibility',
    description: 'Show or hide a chart axis.',
    params: {
      chartName: { type: 'string', description: 'Name of the chart' },
      axisType: {
        type: 'string',
        description: 'Axis type to update',
        enum: ['Category', 'Value', 'Series'],
      },
      visible: { type: 'boolean', description: 'Axis visibility' },
      axisGroup: {
        type: 'string',
        required: false,
        description: 'Optional axis group',
        enum: ['Primary', 'Secondary'],
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
      const axis = chart.axes.getItem(
        args.axisType as Excel.ChartAxisType,
        args.axisGroup as Excel.ChartAxisGroup | undefined
      );
      axis.visible = args.visible as boolean;
      axis.load('visible');
      await context.sync();
      return {
        chartName: args.chartName,
        axisType: args.axisType,
        axisGroup: args.axisGroup ?? 'Primary',
        visible: axis.visible,
      };
    },
  },

  {
    name: 'set_chart_series_filtered',
    description: 'Set whether an individual chart series is filtered (hidden) by index.',
    params: {
      chartName: { type: 'string', description: 'Name of the chart' },
      seriesIndex: { type: 'number', description: 'Zero-based series index' },
      filtered: { type: 'boolean', description: 'True to hide series, false to show series' },
      sheetName: {
        type: 'string',
        required: false,
        description: 'Optional worksheet name.',
      },
    },
    execute: async (context, args) => {
      const sheet = getSheet(context, args.sheetName as string | undefined);
      const chart = sheet.charts.getItem(args.chartName as string);
      const series = chart.series.getItemAt(args.seriesIndex as number);
      series.filtered = args.filtered as boolean;
      series.load(['name', 'filtered']);
      await context.sync();
      return {
        chartName: args.chartName,
        seriesIndex: args.seriesIndex,
        seriesName: series.name,
        filtered: series.filtered,
      };
    },
  },
];
