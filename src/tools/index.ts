import { createTools } from './codegen';
import {
  rangeConfigs,
  tableConfigs,
  chartConfigs,
  sheetConfigs,
  workbookConfigs,
  commentConfigs,
  conditionalFormatConfigs,
  dataValidationConfigs,
  pivotTableConfigs,
} from './configs';
import type { ToolConfig } from './codegen/types';
import type { ToolSet } from 'ai';
import type { OfficeHostApp } from '@/services/office/host';

/** All tool configs combined for manifest generation */
export const allConfigs: readonly (readonly ToolConfig[])[] = [
  rangeConfigs,
  tableConfigs,
  chartConfigs,
  sheetConfigs,
  workbookConfigs,
  commentConfigs,
  conditionalFormatConfigs,
  dataValidationConfigs,
  pivotTableConfigs,
];

/** All Excel tools combined into a single record for AI SDK */
export const excelTools: ToolSet = allConfigs.reduce<ToolSet>((acc, configs) => {
  const generatedTools = createTools(configs);
  return { ...acc, ...generatedTools };
}, {});

export const powerPointTools: ToolSet = {};

export function getToolsForHost(host: OfficeHostApp): ToolSet {
  switch (host) {
    case 'excel':
      return excelTools;
    case 'powerpoint':
      return powerPointTools;
    default:
      return {};
  }
}
