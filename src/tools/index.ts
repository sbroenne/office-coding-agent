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
import type { Tool } from '@github/copilot-sdk';
import type { OfficeHostApp } from '@/services/office/host';
import { powerPointTools } from './powerpoint';
import { wordTools } from './word';

export { webFetchTool } from './general';

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

/** All Excel tools combined into a single array for Copilot SDK */
export const excelTools: Tool[] = allConfigs.flatMap(configs => createTools(configs));

export { powerPointTools } from './powerpoint';
export { wordTools } from './word';

export function getToolsForHost(host: OfficeHostApp): Tool[] {
  switch (host) {
    case 'excel':
      return excelTools.slice(0, MAX_TOOLS_PER_REQUEST);
    case 'powerpoint':
      return powerPointTools.slice(0, MAX_TOOLS_PER_REQUEST);
    case 'word':
      return wordTools.slice(0, MAX_TOOLS_PER_REQUEST);
    default:
      return [];
  }
}
