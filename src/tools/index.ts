import { createTools } from './codegen';
import {
  rangeConfigs,
  rangeFormatConfigs,
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

export { getGeneralTools, webFetchTool, createRunSubagentTool } from './general';

export const MAX_TOOLS_PER_REQUEST = 128;

/** All tool configs combined for manifest generation */
export const allConfigs: readonly (readonly ToolConfig[])[] = [
  rangeConfigs,
  rangeFormatConfigs,
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

function clampToolSet(toolSet: ToolSet, maxTools = MAX_TOOLS_PER_REQUEST): ToolSet {
  const entries = Object.entries(toolSet);
  if (entries.length <= maxTools) return toolSet;
  return Object.fromEntries(entries.slice(0, maxTools)) as ToolSet;
}

export function getToolsForHost(host: OfficeHostApp): ToolSet {
  switch (host) {
    case 'excel':
      return clampToolSet(excelTools);
    case 'powerpoint':
      return clampToolSet(powerPointTools);
    default:
      return {};
  }
}
