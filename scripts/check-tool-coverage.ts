#!/usr/bin/env tsx
/**
 * Excel API Coverage Checker (type-driven)
 *
 * Reads Excel API members directly from @types/office-js and compares them
 * against an explicit coverage map that points API members to tool names.
 *
 * Coverage map file: scripts/tool-coverage-map.json
 *   {
 *     "Range.getUsedRange()": ["get_used_range"],
 *     "Worksheet.name": ["rename_sheet"]
 *   }
 *
 * Usage:
 *   npx tsx scripts/check-tool-coverage.ts [--json] [--verbose] [--strict] [--max-gaps=200]
 */

import * as fs from 'fs';
import * as path from 'path';
import { fileURLToPath } from 'url';
import { createRequire } from 'module';
import ts from 'typescript';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const require = createRequire(import.meta.url);

interface ManifestTool {
  name: string;
  description: string;
  params: Record<string, unknown>;
}

interface ExcelMember {
  typeName: string;
  memberName: string;
  memberKind: 'method' | 'property';
  signature: string;
}

type CoverageMap = Record<string, string[]>;

interface CoverageResult {
  manifestToolCount: number;
  excelTypeCount: number;
  totalMembers: number;
  coveredMembers: number;
  uncoveredMembers: number;
  coveragePercent: number;
  goldenMapEntryCount: number;
  generatedMapEntryCount: number;
  mapEntryCount: number;
  uncoveredByType: Record<string, string[]>;
}

const IGNORE_MEMBERS = new Set([
  'context',
  'isNullObject',
  'load',
  'set',
  'toJSON',
  'track',
  'untrack',
]);

const IGNORE_TYPE_SUFFIXES = ['LoadOptions', 'CollectionLoadOptions', 'UpdateData'];

function parseArgs() {
  const args = process.argv.slice(2);
  const json = args.includes('--json');
  const verbose = args.includes('--verbose') || args.includes('-v');
  const strict = args.includes('--strict');
  const maxGapsArg = args.find(arg => arg.startsWith('--max-gaps='));
  const maxGaps = maxGapsArg ? Number(maxGapsArg.split('=')[1]) : 200;

  return {
    json,
    verbose,
    strict,
    maxGaps: Number.isFinite(maxGaps) ? maxGaps : 200,
  };
}

function readJson<T>(filePath: string, fallback: T): T {
  if (!fs.existsSync(filePath)) return fallback;
  return JSON.parse(fs.readFileSync(filePath, 'utf-8')) as T;
}

function loadManifestTools(): ManifestTool[] {
  const manifestPath = path.resolve(__dirname, '../src/tools/tools-manifest.json');
  const data = readJson<{ tools?: ManifestTool[] }>(manifestPath, {});
  return data.tools ?? [];
}

function normalizeMap(map: CoverageMap): CoverageMap {
  return Object.fromEntries(
    Object.entries(map).map(([member, tools]) => [
      member,
      Array.isArray(tools) ? tools.filter(t => typeof t === 'string') : [],
    ])
  );
}

function mergeMaps(base: CoverageMap, overlay: CoverageMap): CoverageMap {
  const merged: CoverageMap = { ...base };

  for (const [signature, toolNames] of Object.entries(overlay)) {
    const existing = new Set(merged[signature] ?? []);
    for (const toolName of toolNames) existing.add(toolName);
    merged[signature] = [...existing].sort((a, b) => a.localeCompare(b));
  }

  return merged;
}

function loadCoverageMaps(): {
  goldenMap: CoverageMap;
  generatedMap: CoverageMap;
  mergedMap: CoverageMap;
} {
  const goldenPath = path.resolve(__dirname, './tool-coverage-golden-map.json');
  const generatedPath = path.resolve(__dirname, './tool-coverage-map.json');

  const goldenMap = normalizeMap(readJson<CoverageMap>(goldenPath, {}));
  const generatedMap = normalizeMap(readJson<CoverageMap>(generatedPath, {}));
  const mergedMap = mergeMaps(goldenMap, generatedMap);

  return { goldenMap, generatedMap, mergedMap };
}

function resolveOfficeTypesPath(): string {
  const pkgPath = require.resolve('@types/office-js/package.json');
  const pkgDir = path.dirname(pkgPath);
  return path.join(pkgDir, 'index.d.ts');
}

function isIgnoredType(typeName: string): boolean {
  return IGNORE_TYPE_SUFFIXES.some(suffix => typeName.endsWith(suffix));
}

function getMemberName(nodeName: ts.PropertyName | ts.BindingName | undefined): string | null {
  if (!nodeName) return null;
  if (ts.isIdentifier(nodeName) || ts.isStringLiteral(nodeName)) return nodeName.text;
  return null;
}

function extractMembersFromInterfaceOrClass(typeName: string, node: ts.Node): ExcelMember[] {
  const members: ExcelMember[] = [];
  const nodeMembers: readonly ts.TypeElement[] | readonly ts.ClassElement[] =
    ts.isInterfaceDeclaration(node) || ts.isClassDeclaration(node) ? node.members : [];

  for (const member of nodeMembers) {
    if (
      ts.isMethodSignature(member) ||
      ts.isMethodDeclaration(member) ||
      ts.isPropertySignature(member) ||
      ts.isPropertyDeclaration(member) ||
      ts.isGetAccessorDeclaration(member) ||
      ts.isSetAccessorDeclaration(member)
    ) {
      const memberName = getMemberName(member.name);
      if (!memberName || IGNORE_MEMBERS.has(memberName) || memberName.startsWith('_')) continue;

      const isMethod =
        ts.isMethodSignature(member) ||
        ts.isMethodDeclaration(member) ||
        ts.isGetAccessorDeclaration(member) ||
        ts.isSetAccessorDeclaration(member);

      members.push({
        typeName,
        memberName,
        memberKind: isMethod ? 'method' : 'property',
        signature: `${typeName}.${memberName}${isMethod ? '()' : ''}`,
      });
    }
  }

  return members;
}

function findExcelNamespace(sourceFile: ts.SourceFile): ts.ModuleBlock | null {
  const statements = sourceFile.statements;
  for (const stmt of statements) {
    if (!ts.isModuleDeclaration(stmt)) continue;
    if (stmt.name.getText(sourceFile) !== 'Excel') continue;
    if (stmt.body && ts.isModuleBlock(stmt.body)) return stmt.body;
  }
  return null;
}

function extractExcelMembers(): ExcelMember[] {
  const typesPath = resolveOfficeTypesPath();
  const sourceText = fs.readFileSync(typesPath, 'utf-8');
  const sourceFile = ts.createSourceFile(typesPath, sourceText, ts.ScriptTarget.Latest, true);

  const excelNamespace = findExcelNamespace(sourceFile);
  if (!excelNamespace) {
    throw new Error('Could not find `declare namespace Excel` in @types/office-js/index.d.ts.');
  }

  const members: ExcelMember[] = [];
  for (const stmt of excelNamespace.statements) {
    if (!(ts.isInterfaceDeclaration(stmt) || ts.isClassDeclaration(stmt))) continue;
    if (!stmt.name) continue;
    const typeName = stmt.name.text;
    if (!typeName || isIgnoredType(typeName)) continue;

    members.push(...extractMembersFromInterfaceOrClass(typeName, stmt));
  }

  members.sort((a, b) => a.signature.localeCompare(b.signature));
  return members;
}

function analyze(): CoverageResult {
  const tools = loadManifestTools();
  const toolNames = new Set(tools.map(tool => tool.name));
  const { goldenMap, generatedMap, mergedMap } = loadCoverageMaps();
  const members = extractExcelMembers();

  const uncoveredByType: Record<string, string[]> = {};
  let coveredMembers = 0;

  for (const member of members) {
    const mappedTools = mergedMap[member.signature] ?? [];
    const covered =
      mappedTools.length > 0 && mappedTools.every(toolName => toolNames.has(toolName));

    if (covered) {
      coveredMembers += 1;
      continue;
    }

    if (!uncoveredByType[member.typeName]) uncoveredByType[member.typeName] = [];
    uncoveredByType[member.typeName].push(member.signature);
  }

  const totalMembers = members.length;
  const uncoveredMembers = totalMembers - coveredMembers;
  const excelTypeCount = new Set(members.map(m => m.typeName)).size;

  return {
    manifestToolCount: toolNames.size,
    excelTypeCount,
    totalMembers,
    coveredMembers,
    uncoveredMembers,
    coveragePercent: totalMembers === 0 ? 0 : Math.round((coveredMembers / totalMembers) * 100),
    goldenMapEntryCount: Object.keys(goldenMap).length,
    generatedMapEntryCount: Object.keys(generatedMap).length,
    mapEntryCount: Object.keys(mergedMap).length,
    uncoveredByType,
  };
}

function printReport(result: CoverageResult, verbose: boolean, maxGaps: number): void {
  console.log('\n╔══════════════════════════════════════════════════════════════╗');
  console.log('║   Excel API Tool Coverage (Type-Driven from office-js)      ║');
  console.log('╚══════════════════════════════════════════════════════════════╝\n');

  console.log(`Tools in manifest: ${result.manifestToolCount}`);
  console.log(`Excel types discovered: ${result.excelTypeCount}`);
  console.log(`Excel members discovered: ${result.totalMembers}`);
  console.log(`Coverage map entries: ${result.mapEntryCount}`);
  console.log(`  - golden: ${result.goldenMapEntryCount}`);
  console.log(`  - generated: ${result.generatedMapEntryCount}`);
  console.log(
    `Coverage: ${result.coveredMembers}/${result.totalMembers} (${result.coveragePercent}%)\n`
  );

  const typeRows = Object.entries(result.uncoveredByType)
    .map(([typeName, gaps]) => ({ typeName, count: gaps.length }))
    .sort((a, b) => b.count - a.count);

  console.log('Top uncovered types:');
  for (const row of typeRows.slice(0, 20)) {
    console.log(`  - ${row.typeName}: ${row.count} uncovered member(s)`);
  }
  console.log();

  if (verbose) {
    let printed = 0;
    for (const [typeName, gaps] of Object.entries(result.uncoveredByType)) {
      if (printed >= maxGaps) break;
      console.log(`Uncovered in ${typeName}:`);
      for (const gap of gaps) {
        if (printed >= maxGaps) break;
        console.log(`  - ${gap}`);
        printed += 1;
      }
    }

    if (printed >= maxGaps) {
      console.log(`\n...truncated at ${maxGaps} gaps. Use --max-gaps=<n> to show more.`);
    }
  }
}

function main(): void {
  const options = parseArgs();
  const result = analyze();

  if (options.json) {
    console.log(JSON.stringify(result, null, 2));
  } else {
    printReport(result, options.verbose, options.maxGaps);
  }

  if (options.strict && result.uncoveredMembers > 0) {
    console.log(`\n❌ ${result.uncoveredMembers} uncovered Excel API member(s) remain.\n`);
    process.exit(1);
  }

  if (!options.json) {
    console.log(
      result.uncoveredMembers === 0
        ? '\n✅ Full Excel API member coverage achieved.\n'
        : `\n⚠️  ${result.uncoveredMembers} Excel API member(s) are currently uncovered.\n`
    );
  }
}

main();
