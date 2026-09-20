// @vitest-environment node
import { execFileSync, spawnSync } from 'node:child_process';
import { copyFileSync, mkdirSync, mkdtempSync, readFileSync, rmSync, symlinkSync } from 'node:fs';
import { createRequire } from 'node:module';
import os from 'node:os';
import path from 'node:path';
import { afterAll, beforeAll, describe, expect, it } from 'vitest';

const require = createRequire(import.meta.url);
const root = path.resolve(import.meta.dirname, '../..');
const runner = require.resolve('tsx/cli');
let fixture: string;

function runScript(name: string, ...args: string[]) {
  return execFileSync(process.execPath, [runner, path.join(fixture, 'scripts', name), ...args], {
    cwd: fixture,
    encoding: 'utf8',
    timeout: 30_000,
  });
}

describe('tool coverage scripts with the native TypeScript API', () => {
  beforeAll(() => {
    fixture = mkdtempSync(path.join(os.tmpdir(), 'office coverage '));
    mkdirSync(path.join(fixture, 'scripts'));
    mkdirSync(path.join(fixture, 'src/tools'), { recursive: true });
    symlinkSync(path.join(root, 'node_modules'), path.join(fixture, 'node_modules'), 'junction');
    for (const file of [
      'package.json',
      'scripts/check-tool-coverage.ts',
      'scripts/bootstrap-tool-coverage-map.ts',
      'scripts/tool-coverage-map.json',
      'scripts/tool-coverage-golden-map.json',
      'src/tools/tools-manifest.json',
    ]) {
      copyFileSync(path.join(root, file), path.join(fixture, file));
    }
  });

  afterAll(() => {
    if (fixture) rmSync(fixture, { recursive: true, force: true });
  });

  it('reads real Office declarations and preserves member filtering', () => {
    const report = JSON.parse(runScript('check-tool-coverage.ts', '--json'));
    expect(report.excelTypeCount).toBeGreaterThan(300);
    expect(report.totalMembers).toBeGreaterThan(2000);
    expect(report.coveredMembers + report.uncoveredMembers).toBe(report.totalMembers);
    expect(report.uncoveredByType.Range).toEqual(
      expect.arrayContaining(['Range.values', 'Range.getCell()', 'Range.address'])
    );
    expect(report.uncoveredByType.Range).not.toEqual(expect.arrayContaining(['Range.load()']));
    expect(Object.keys(report.uncoveredByType).some(name => name.endsWith('LoadOptions'))).toBe(
      false
    );
  });

  it('keeps strict mode failing on uncovered members rather than compiler API errors', () => {
    const result = spawnSync(
      process.execPath,
      [runner, path.join(fixture, 'scripts/check-tool-coverage.ts'), '--strict'],
      { cwd: fixture, encoding: 'utf8', timeout: 30_000 }
    );
    expect(result.error).toBeUndefined();
    expect(result.status).toBe(1);
    expect(result.stdout).toContain('uncovered Excel API member(s) remain');
    expect(result.stderr).toBe('');
  });

  it('bootstraps the same candidate mappings without changing the golden baseline', () => {
    const goldenPath = path.join(fixture, 'scripts/tool-coverage-golden-map.json');
    const golden = readFileSync(goldenPath, 'utf8');
    runScript('bootstrap-tool-coverage-map.ts');
    const mapPath = path.join(fixture, 'scripts/tool-coverage-map.json');
    const generated = readFileSync(mapPath, 'utf8');
    expect(JSON.parse(generated)).toEqual({
      'ChartSeries.getDimensionDataSourceType()': ['chart'],
      'CommentCollection.getItemByCell()': ['comment'],
      'ConditionalFormat.type': ['conditional_format'],
      'PivotLayout.getColumnLabelRange()': ['pivot'],
      'RangeFormat.columnWidth': ['range_format'],
      'TableColumn.getHeaderRowRange()': ['table'],
    });
    expect(readFileSync(goldenPath, 'utf8')).toBe(golden);
    runScript('bootstrap-tool-coverage-map.ts');
    expect(readFileSync(mapPath, 'utf8')).toBe(generated);
  });
});
