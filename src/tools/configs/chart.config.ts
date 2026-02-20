/**
 * Chart tool configs — 1 tool (chart) with actions: list, create, delete, configure.
 */

import type { ToolConfig } from '../codegen';
import { getSheet } from '../codegen';

export const chartConfigs: readonly ToolConfig[] = [
  {
    name: 'chart',
    description:
      'Manage charts. Use action "list" to list charts on a sheet, "create" to make a new chart, "delete" to remove one, or "configure" to update title/type/data/position/legend of an existing chart.',
    params: {
      action: {
        type: 'string',
        description: 'Operation to perform',
        enum: ['list', 'create', 'delete', 'configure'],
      },
      chartName: {
        type: 'string',
        required: false,
        description: 'Chart name. Required for delete/configure.',
      },
      // create params
      chartType: {
        type: 'string',
        required: false,
        description: 'Chart type for create/configure.',
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
      dataRange: {
        type: 'string',
        required: false,
        description:
          'Data range address (e.g. "A1:D20"). Required for create; used in configure to change data source.',
      },
      name: {
        type: 'string',
        required: false,
        description: 'Name to assign to the new chart (create action).',
      },
      seriesBy: {
        type: 'string',
        required: false,
        description: 'How to interpret series. Default "Auto".',
        enum: ['Auto', 'Columns', 'Rows'],
      },
      // configure / position params
      title: { type: 'string', required: false, description: 'Chart title text (configure).' },
      left: { type: 'number', required: false, description: 'Left position in points.' },
      top: { type: 'number', required: false, description: 'Top position in points.' },
      width: { type: 'number', required: false, description: 'Width in points.' },
      height: { type: 'number', required: false, description: 'Height in points.' },
      startCell: { type: 'string', required: false, description: 'Top-left anchor cell.' },
      endCell: { type: 'string', required: false, description: 'Bottom-right anchor cell.' },
      legendVisible: { type: 'boolean', required: false, description: 'Show/hide legend.' },
      legendPosition: {
        type: 'string',
        required: false,
        description: 'Legend position.',
        enum: ['Top', 'Bottom', 'Left', 'Right', 'Corner'],
      },
      sheetName: { type: 'string', required: false, description: 'Optional worksheet name.' },
    },
    execute: async (context, args) => {
      const action = args.action as string;
      const sheet = getSheet(context, args.sheetName as string | undefined);

      if (action === 'list') {
        const charts = sheet.charts;
        charts.load('items');
        await context.sync();
        for (const c of charts.items) {
          c.load(['name', 'chartType', 'left', 'top', 'width', 'height']);
        }
        await context.sync();
        const result = charts.items.map(c => ({
          name: c.name,
          chartType: c.chartType,
          left: c.left,
          top: c.top,
          width: c.width,
          height: c.height,
        }));
        return { charts: result, count: result.length };
      }

      if (action === 'create') {
        const range = sheet.getRange(args.dataRange as string);
        const seriesBy = ((args.seriesBy as string) ?? 'Auto') as Excel.ChartSeriesBy;
        const chart = sheet.charts.add(
          (args.chartType as Excel.ChartType) ?? 'ColumnClustered',
          range,
          seriesBy
        );
        if (args.name) chart.name = args.name as string;
        chart.load(['name', 'chartType', 'left', 'top', 'width', 'height']);
        await context.sync();
        return {
          name: chart.name,
          chartType: chart.chartType,
          left: chart.left,
          top: chart.top,
          width: chart.width,
          height: chart.height,
        };
      }

      const chart = sheet.charts.getItem(args.chartName as string);

      if (action === 'delete') {
        chart.delete();
        await context.sync();
        return { chartName: args.chartName, deleted: true };
      }

      // configure
      if (args.title !== undefined) chart.title.text = args.title as string;
      if (args.chartType !== undefined) chart.chartType = args.chartType as Excel.ChartType;
      if (args.dataRange !== undefined) {
        const range = sheet.getRange(args.dataRange as string);
        chart.setData(range);
      }
      if (args.startCell !== undefined) {
        chart.setPosition(args.startCell as string, args.endCell as string | undefined);
      }
      if (args.left !== undefined) chart.left = args.left as number;
      if (args.top !== undefined) chart.top = args.top as number;
      if (args.width !== undefined) chart.width = args.width as number;
      if (args.height !== undefined) chart.height = args.height as number;
      if (args.legendVisible !== undefined) chart.legend.visible = args.legendVisible as boolean;
      if (args.legendPosition !== undefined)
        chart.legend.position = args.legendPosition as Excel.ChartLegendPosition;

      chart.load(['name', 'chartType', 'left', 'top', 'width', 'height']);
      await context.sync();
      return {
        chartName: chart.name,
        chartType: chart.chartType,
        left: chart.left,
        top: chart.top,
        width: chart.width,
        height: chart.height,
      };
    },
  },
];
