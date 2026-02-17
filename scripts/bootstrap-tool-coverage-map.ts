#!/usr/bin/env tsx

import * as fs from 'fs';
import * as path from 'path';
import ts from 'typescript';

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

const STOP_WORDS = new Set([
  'the',
  'a',
  'an',
  'to',
  'for',
  'and',
  'or',
  'of',
  'on',
  'in',
  'with',
  'from',
  'by',
  'all',
  'current',
  'active',
  'optional',
  'excel',
]);

const ACTION_TOKENS = new Set([
  'get',
  'set',
  'list',
  'add',
  'create',
  'delete',
  'remove',
  'clear',
  'update',
  'edit',
  'move',
  'copy',
  'toggle',
  'refresh',
  'convert',
  'find',
  'replace',
  'sort',
  'merge',
  'unmerge',
  'define',
  'protect',
  'unprotect',
  'rename',
]);

const TYPE_HINTS: Array<{ token: string; types: string[] }> = [
  { token: 'range', types: ['Range', 'RangeFormat', 'RangeSort', 'RangeAreas'] },
  { token: 'sheet', types: ['Worksheet', 'WorksheetCollection', 'PageLayout'] },
  { token: 'workbook', types: ['Workbook', 'Application'] },
  { token: 'table', types: ['Table', 'TableCollection', 'TableColumn', 'TableRow'] },
  { token: 'chart', types: ['Chart', 'ChartCollection', 'ChartAxis', 'ChartSeries'] },
  { token: 'comment', types: ['Comment', 'CommentCollection'] },
  {
    token: 'pivot',
    types: ['PivotTable', 'PivotTableCollection', 'PivotHierarchy', 'PivotField', 'PivotLayout'],
  },
  { token: 'validation', types: ['DataValidation'] },
  { token: 'conditional', types: ['ConditionalFormat', 'ConditionalFormatCollection'] },
  { token: 'format', types: ['RangeFormat', 'ConditionalFormat', 'NumberFormatInfo'] },
  { token: 'named', types: ['NamedItem', 'NamedItemCollection'] },
  { token: 'hyperlink', types: ['Range', 'RangeHyperlink'] },
  { token: 'pane', types: ['WorksheetFreezePanes'] },
  { token: 'panes', types: ['WorksheetFreezePanes'] },
  { token: 'border', types: ['Range', 'RangeFormat'] },
  { token: 'formula', types: ['Range'] },
  { token: 'comment', types: ['Comment', 'CommentCollection'] },
];

function readJson<T>(filePath: string, fallback: T): T {
  if (!fs.existsSync(filePath)) return fallback;
  return JSON.parse(fs.readFileSync(filePath, 'utf-8')) as T;
}

function writeJson(filePath: string, value: unknown): void {
  fs.writeFileSync(filePath, `${JSON.stringify(value, null, 2)}\n`, 'utf-8');
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

function findExcelNamespace(sourceFile: ts.SourceFile): ts.ModuleBlock | null {
  for (const stmt of sourceFile.statements) {
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
    throw new Error('Could not find Excel namespace in @types/office-js/index.d.ts');
  }

  const seen = new Set<string>();
  const members: ExcelMember[] = [];

  for (const stmt of excelNamespace.statements) {
    if (!(ts.isInterfaceDeclaration(stmt) || ts.isClassDeclaration(stmt))) continue;
    if (!stmt.name) continue;
    const typeName = stmt.name.text;
    if (!typeName || isIgnoredType(typeName)) continue;

    for (const member of stmt.members) {
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

        const signature = `${typeName}.${memberName}${isMethod ? '()' : ''}`;
        if (seen.has(signature)) continue;
        seen.add(signature);

        members.push({
          typeName,
          memberName,
          memberKind: isMethod ? 'method' : 'property',
          signature,
        });
      }
    }
  }

  members.sort((a, b) => a.signature.localeCompare(b.signature));
  return members;
}

function loadManifestTools(): ManifestTool[] {
  const manifestPath = path.resolve(__dirname, '../src/tools/tools-manifest.json');
  const manifest = readJson<{ tools?: ManifestTool[] }>(manifestPath, {});
  return manifest.tools ?? [];
}

function loadCoverageMap(): CoverageMap {
  const mapPath = path.resolve(__dirname, './tool-coverage-map.json');
  return readJson<CoverageMap>(mapPath, {});
}

function loadGoldenCoverageMap(): CoverageMap {
  const goldenMapPath = path.resolve(__dirname, './tool-coverage-golden-map.json');
  return readJson<CoverageMap>(goldenMapPath, {});
}

function tokenize(value: string): string[] {
  return value
    .replace(/([a-z])([A-Z])/g, '$1 $2')
    .split(/[^A-Za-z0-9]+/)
    .map(token => token.toLowerCase())
    .filter(token => token.length > 1 && !STOP_WORDS.has(token));
}

function inferTypeHints(tokens: Set<string>): Set<string> {
  const hints = new Set<string>();

  for (const hint of TYPE_HINTS) {
    if (tokens.has(hint.token)) {
      for (const typeName of hint.types) hints.add(typeName);
    }
  }

  return hints;
}

function actionToken(toolName: string): string {
  return toolName.split('_')[0] ?? '';
}

function getToolTokens(tool: ManifestTool): Set<string> {
  return new Set(
    tokenize(`${tool.name} ${tool.description}`).map(token =>
      token.endsWith('s') && token.length > 3 ? token.slice(0, -1) : token
    )
  );
}

function getSpecificToolTokens(toolTokens: Set<string>): string[] {
  return [...toolTokens].filter(token => !ACTION_TOKENS.has(token));
}

function inferTypeHintsFromToolName(toolName: string, currentHints: Set<string>): Set<string> {
  const hints = new Set(currentHints);
  const normalized = toolName.toLowerCase();

  if (normalized.includes('sheet')) {
    hints.add('Worksheet');
    hints.add('WorksheetCollection');
    hints.add('WorksheetFreezePanes');
  }
  if (normalized.includes('workbook')) {
    hints.add('Workbook');
    hints.add('Application');
  }
  if (normalized.includes('range') || normalized.includes('row') || normalized.includes('column')) {
    hints.add('Range');
    hints.add('RangeFormat');
  }
  if (normalized.includes('table')) {
    hints.add('Table');
    hints.add('TableCollection');
    hints.add('TableColumn');
    hints.add('TableRow');
  }
  if (normalized.includes('chart')) {
    hints.add('Chart');
    hints.add('ChartCollection');
    hints.add('ChartAxis');
    hints.add('ChartSeries');
  }
  if (normalized.includes('pivot')) {
    hints.add('PivotTable');
    hints.add('PivotTableCollection');
    hints.add('PivotHierarchy');
    hints.add('PivotField');
    hints.add('PivotLayout');
  }
  if (normalized.includes('comment')) {
    hints.add('Comment');
    hints.add('CommentCollection');
  }
  if (normalized.includes('validation')) {
    hints.add('DataValidation');
  }
  if (normalized.includes('conditional')) {
    hints.add('ConditionalFormat');
    hints.add('ConditionalFormatCollection');
  }
  if (normalized.includes('named')) {
    hints.add('NamedItem');
    hints.add('NamedItemCollection');
  }

  return hints;
}

function inferHardDomainHintsFromToolName(toolName: string): Set<string> {
  const normalized = toolName.toLowerCase();
  const hints = new Set<string>();

  if (normalized.includes('sheet')) {
    hints.add('Worksheet');
    hints.add('WorksheetCollection');
    hints.add('WorksheetFreezePanes');
    hints.add('PageLayout');
  }
  if (normalized.includes('workbook')) {
    hints.add('Workbook');
    hints.add('Application');
  }
  if (normalized.includes('range') || normalized.includes('row') || normalized.includes('column')) {
    hints.add('Range');
    hints.add('RangeFormat');
    hints.add('RangeSort');
    hints.add('RangeHyperlink');
  }
  if (normalized.includes('table')) {
    hints.add('Table');
    hints.add('TableCollection');
    hints.add('TableColumn');
    hints.add('TableRow');
  }
  if (normalized.includes('chart')) {
    hints.add('Chart');
    hints.add('ChartCollection');
    hints.add('ChartAxis');
    hints.add('ChartSeries');
  }
  if (normalized.includes('pivot')) {
    hints.add('PivotTable');
    hints.add('PivotTableCollection');
    hints.add('PivotHierarchy');
    hints.add('PivotField');
    hints.add('PivotLayout');
  }
  if (normalized.includes('comment')) {
    hints.add('Comment');
    hints.add('CommentCollection');
  }
  if (normalized.includes('validation')) {
    hints.add('DataValidation');
  }
  if (normalized.includes('conditional')) {
    hints.add('ConditionalFormat');
    hints.add('ConditionalFormatCollection');
  }
  if (normalized.includes('named')) {
    hints.add('NamedItem');
    hints.add('NamedItemCollection');
  }
  if (normalized.includes('hyperlink')) {
    hints.add('Range');
    hints.add('RangeHyperlink');
  }
  if (normalized.includes('freeze_panes') || normalized.includes('freeze')) {
    hints.add('WorksheetFreezePanes');
    hints.add('Worksheet');
  }
  if (normalized.includes('page_layout')) {
    hints.add('PageLayout');
    hints.add('Worksheet');
  }

  return hints;
}

function tokenMatches(toolToken: string, memberToken: string): boolean {
  if (toolToken === memberToken) return true;

  const aliases: Record<string, string[]> = {
    sheet: ['worksheet'],
    sheets: ['worksheet', 'worksheetcollection'],
    row: ['row'],
    column: ['column'],
    columns: ['column'],
    cell: ['cell', 'range'],
    cells: ['cell', 'range'],
  };

  const expanded = aliases[toolToken] ?? [];
  return expanded.includes(memberToken);
}

function memberTokenOverlap(tool: ManifestTool, member: ExcelMember): number {
  const toolTokens = getToolTokens(tool);
  const specific = getSpecificToolTokens(toolTokens);
  const memberTokens = new Set(tokenize(`${member.typeName} ${member.memberName}`));

  return specific.reduce((count, token) => {
    for (const memberToken of memberTokens) {
      if (tokenMatches(token, memberToken)) return count + 1;
    }
    return count;
  }, 0);
}

function isActionCompatible(op: string, member: ExcelMember): boolean {
  const name = member.memberName.toLowerCase();

  if (op === 'get' || op === 'list') {
    return member.memberKind === 'property' || name.startsWith('get') || name === 'items';
  }
  if (op === 'set' || op === 'update' || op === 'edit' || op === 'rename' || op === 'move') {
    if (op === 'rename') {
      return name === 'name' || name.startsWith('setname');
    }
    if (op === 'move') {
      return name === 'position' || name.includes('move');
    }
    return member.memberKind === 'property' || name.startsWith('set') || name.startsWith('update');
  }
  if (op === 'add' || op === 'create' || op === 'define') {
    return member.memberKind === 'method' && (name.startsWith('add') || name.startsWith('create'));
  }
  if (op === 'delete' || op === 'remove' || op === 'clear') {
    return (
      member.memberKind === 'method' &&
      (name.startsWith('delete') || name.startsWith('remove') || name.startsWith('clear'))
    );
  }
  if (op === 'refresh') {
    return member.memberKind === 'method' && name.startsWith('refresh');
  }
  if (op === 'copy') {
    return member.memberKind === 'method' && name.startsWith('copy');
  }
  if (op === 'sort') {
    return name.includes('sort') || name === 'apply';
  }
  if (op === 'find') {
    return name.startsWith('find');
  }
  if (op === 'replace') {
    return name.startsWith('replace');
  }
  if (op === 'merge' || op === 'unmerge') {
    return name.startsWith(op);
  }
  if (op === 'toggle') {
    return member.memberKind === 'property';
  }
  if (op === 'protect' || op === 'unprotect') {
    return name.includes('protect');
  }

  return true;
}

function isPlausibleMapping(tool: ManifestTool, member: ExcelMember): boolean {
  const toolTokens = getToolTokens(tool);
  const hardHints = inferHardDomainHintsFromToolName(tool.name);
  if (hardHints.size > 0 && !hardHints.has(member.typeName)) return false;

  const hints = inferTypeHintsFromToolName(tool.name, inferTypeHints(toolTokens));
  if (hints.size > 0 && !hints.has(member.typeName)) return false;

  const op = actionToken(tool.name);
  if (!isActionCompatible(op, member)) return false;

  const overlap = memberTokenOverlap(tool, member);
  const specificTokenCount = getSpecificToolTokens(toolTokens).length;

  if (specificTokenCount === 0) return false;
  if (hints.size === 0 && overlap < 2) return false;
  if (hints.size > 0 && overlap < 1) return false;

  return true;
}

function scoreMember(tool: ManifestTool, member: ExcelMember): number {
  const toolTokens = getToolTokens(tool);
  const memberTokens = tokenize(`${member.typeName} ${member.memberName}`);

  const hardHints = inferHardDomainHintsFromToolName(tool.name);
  if (hardHints.size > 0 && !hardHints.has(member.typeName)) return -100;

  const hints = inferTypeHintsFromToolName(tool.name, inferTypeHints(toolTokens));
  if (hints.size > 0 && !hints.has(member.typeName)) return -100;

  let score = 0;

  for (const token of memberTokens) {
    if (toolTokens.has(token)) score += 2;
  }

  const op = actionToken(tool.name);
  if (op === 'get' || op === 'list') {
    if (member.memberKind === 'property' || member.memberName.startsWith('get')) score += 1;
  }
  if (op === 'set' || op === 'update') {
    if (member.memberKind === 'property' || member.memberName.startsWith('set')) score += 1;
  }
  if (op === 'add' || op === 'create') {
    if (member.memberName.startsWith('add')) score += 2;
  }
  if (op === 'delete' || op === 'remove' || op === 'clear') {
    if (
      member.memberName.startsWith('delete') ||
      member.memberName.startsWith('remove') ||
      member.memberName.startsWith('clear')
    ) {
      score += 2;
    }
  }

  const exactMemberToken = member.memberName.toLowerCase();
  if (toolTokens.has(exactMemberToken)) score += 2;

  const overlap = memberTokenOverlap(tool, member);
  score += overlap * 3;

  if (!isActionCompatible(op, member)) score -= 6;

  return score;
}

function buildBootstrapMap(tools: ManifestTool[], members: ExcelMember[]): CoverageMap {
  const result: CoverageMap = {};
  const minScore = 9;

  for (const tool of tools) {
    let best: ExcelMember | null = null;
    let bestScore = -Infinity;

    for (const member of members) {
      const score = scoreMember(tool, member);
      if (score > bestScore) {
        bestScore = score;
        best = member;
      }
    }

    if (!best || bestScore < minScore) continue;
    if (!isPlausibleMapping(tool, best)) continue;

    result[best.signature] ??= [];
    if (!result[best.signature].includes(tool.name)) {
      result[best.signature].push(tool.name);
    }
  }

  for (const signature of Object.keys(result)) {
    result[signature].sort((a, b) => a.localeCompare(b));
  }

  return result;
}

function cleanCurrentMap(
  current: CoverageMap,
  toolsByName: Map<string, ManifestTool>,
  membersBySignature: Map<string, ExcelMember>
): CoverageMap {
  const cleaned: CoverageMap = {};

  for (const [signature, toolNames] of Object.entries(current)) {
    const member = membersBySignature.get(signature);
    if (!member) continue;

    const retained = toolNames.filter(toolName => {
      const tool = toolsByName.get(toolName);
      if (!tool) return false;
      return isPlausibleMapping(tool, member);
    });

    if (retained.length > 0) {
      cleaned[signature] = [...new Set(retained)].sort((a, b) => a.localeCompare(b));
    }
  }

  return cleaned;
}

function mergeCoverageMaps(current: CoverageMap, generated: CoverageMap): CoverageMap {
  const merged: CoverageMap = { ...current };

  for (const [signature, tools] of Object.entries(generated)) {
    const existing = new Set(merged[signature] ?? []);
    for (const tool of tools) existing.add(tool);
    merged[signature] = [...existing].sort((a, b) => a.localeCompare(b));
  }

  return Object.fromEntries(
    Object.entries(merged).sort(([left], [right]) => left.localeCompare(right))
  );
}

function removeGoldenCoveredSignatures(generated: CoverageMap, golden: CoverageMap): CoverageMap {
  const filtered: CoverageMap = {};
  const goldenSignatures = new Set(Object.keys(golden));

  for (const [signature, tools] of Object.entries(generated)) {
    if (goldenSignatures.has(signature)) continue;
    filtered[signature] = tools;
  }

  return filtered;
}

function main(): void {
  const mapPath = path.resolve(__dirname, './tool-coverage-map.json');
  const tools = loadManifestTools();
  const members = extractExcelMembers();
  const golden = loadGoldenCoverageMap();
  const current = loadCoverageMap();
  const toolsByName = new Map(tools.map(tool => [tool.name, tool]));
  const membersBySignature = new Map(members.map(member => [member.signature, member]));
  const cleaned = cleanCurrentMap(current, toolsByName, membersBySignature);
  const generated = buildBootstrapMap(tools, members);
  const filteredGenerated = removeGoldenCoveredSignatures(generated, golden);
  const merged = mergeCoverageMaps(cleaned, filteredGenerated);

  writeJson(mapPath, merged);

  const generatedCount = Object.keys(filteredGenerated).length;
  const cleanedCount = Object.keys(cleaned).length;
  const originalCount = Object.keys(current).length;
  const totalCount = Object.keys(merged).length;
  const goldenCount = Object.keys(golden).length;

  console.log(
    `Cleaned existing map entries: ${originalCount} -> ${cleanedCount} (removed ${
      originalCount - cleanedCount
    }).`
  );
  console.log(`Golden map entries (immutable baseline): ${goldenCount}.`);
  console.log(`Generated ${generatedCount} additive candidate mapping entries.`);
  console.log(`Generated map now contains ${totalCount} signature entries.`);
}

main();
