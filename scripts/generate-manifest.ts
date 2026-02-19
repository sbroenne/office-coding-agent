/**
 * Generate tools-manifest.json from all tool configs.
 *
 * This runs via vitest to leverage the project's path resolution and TS config.
 * Run: npx vitest run scripts/generate-manifest.ts
 * Or:  npm run manifest
 */
import { writeFileSync } from 'node:fs';
import { resolve, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import { generateManifest } from '../src/tools/codegen/manifest';
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
} from '../src/tools/configs';

const __dirname = dirname(fileURLToPath(import.meta.url));

const manifest = generateManifest(
  rangeConfigs,
  tableConfigs,
  chartConfigs,
  sheetConfigs,
  workbookConfigs,
  commentConfigs,
  conditionalFormatConfigs,
  dataValidationConfigs,
  pivotTableConfigs
);

const outPath = resolve(__dirname, '..', 'src', 'tools', 'tools-manifest.json');
writeFileSync(outPath, JSON.stringify(manifest, null, 2) + '\n');

console.log(`Generated ${manifest.tools.length} tools -> ${outPath}`);
