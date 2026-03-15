/**
 * Generate manifest files for AI tool-schema testing.
 *
 * This runs via vitest to leverage the project's path resolution and TS config.
 * Run: npx vitest run scripts/generate-manifest.ts
 * Or:  npm run manifest
 */
import { mkdirSync, writeFileSync } from 'node:fs';
import { resolve, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import type { Tool } from '@github/copilot-sdk';
import { generateManifest } from '../src/tools/codegen/manifest';
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
} from '../src/tools/configs';
import { powerPointTools, wordTools, outlookTools } from '../src/tools';

const __dirname = dirname(fileURLToPath(import.meta.url));

interface EvalManifest {
  version: string;
  generatedAt: string;
  host: string;
  tools: Array<{
    name: string;
    description: string;
    inputSchema: Record<string, unknown>;
  }>;
}

function writeJson(path: string, content: unknown): void {
  mkdirSync(dirname(path), { recursive: true });
  writeFileSync(path, JSON.stringify(content, null, 2) + '\n');
}

function normalizeInputSchema(schema: unknown): Record<string, unknown> {
  if (schema && typeof schema === 'object') {
    const schemaObject = schema as Record<string, unknown>;
    const required = Array.isArray(schemaObject.required) ? schemaObject.required : [];
    return {
      ...schemaObject,
      type: 'object',
      properties:
        schemaObject.properties && typeof schemaObject.properties === 'object'
          ? schemaObject.properties
          : {},
      required,
    };
  }

  return { type: 'object', properties: {}, required: [] };
}

function createEvalManifest(host: string, tools: readonly Tool[]): EvalManifest {
  return {
    version: '1.0.0',
    generatedAt: new Date().toISOString(),
    host,
    tools: tools.map(tool => ({
      name: tool.name,
      description: tool.description ?? '',
      inputSchema: normalizeInputSchema(tool.parameters),
    })),
  };
}

const evalManifestDir = resolve(__dirname, '..', 'tests-aitest', 'manifests');

const excelManifest = generateManifest(
  rangeConfigs,
  rangeFormatConfigs,
  tableConfigs,
  chartConfigs,
  sheetConfigs,
  workbookConfigs,
  commentConfigs,
  conditionalFormatConfigs,
  dataValidationConfigs,
  pivotTableConfigs
);
const excelOutPath = resolve(evalManifestDir, 'excel-tools-manifest.json');
writeJson(excelOutPath, excelManifest);
console.log(`Generated ${excelManifest.tools.length} excel manifest tools -> ${excelOutPath}`);

const hostToolSets = {
  powerpoint: powerPointTools,
  word: wordTools,
  outlook: outlookTools,
} satisfies Record<string, readonly Tool[]>;

for (const [host, tools] of Object.entries(hostToolSets)) {
  const hostManifest = createEvalManifest(host, tools);
  const outPath = resolve(evalManifestDir, `${host}-tools-manifest.json`);
  writeJson(outPath, hostManifest);
  console.log(`Generated ${hostManifest.tools.length} ${host} eval tools -> ${outPath}`);
}
