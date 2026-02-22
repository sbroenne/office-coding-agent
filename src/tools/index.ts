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
import type { ToolConfig, ToolConfigBase } from './codegen/types';
import type { Tool } from '@github/copilot-sdk';
import type { OfficeHostApp } from '@/services/office/host';
import { powerPointTools, powerPointConfigs } from './powerpoint';
import { wordTools, wordConfigs } from './word';

export { webFetchTool } from './general';

export const MAX_TOOLS_PER_REQUEST = 128;

/** All Excel tool configs combined for manifest generation */
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

/** All tool configs across all hosts — for manifest generation */
export const allConfigsByHost: Record<string, readonly (readonly ToolConfigBase[])[]> = {
  excel: allConfigs,
  powerpoint: [powerPointConfigs],
  word: [wordConfigs],
};

/** All Excel tools combined into a single array for Copilot SDK */
export const excelTools: Tool[] = allConfigs.flatMap(configs => createTools(configs));

export { powerPointTools, powerPointConfigs } from './powerpoint';
export { wordTools, wordConfigs } from './word';

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
