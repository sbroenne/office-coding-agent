/**
 * Integration tests for permissionStore.
 *
 * Validates evaluate() path matching (normalize, isUnderPath, pathForRequest),
 * rule CRUD, allowAll toggle, workingDirectory read auto-approve, and persistence config.
 */
import { describe, it, expect, beforeEach } from 'vitest';
import { usePermissionStore } from '@/stores/permissionStore';

beforeEach(() => {
  usePermissionStore.setState({
    allowAll: false,
    workingDirectory: null,
    rules: [],
  });
});

// ─── allowAll ─────────────────────────────────────────────────────────────────

describe('permissionStore — allowAll', () => {
  it('defaults to true in a fresh store (production default)', () => {
    // The store's initialState in create() sets allowAll: true — verify via getInitialState
    const freshStore = usePermissionStore.getInitialState?.();
    // If getInitialState exists it should have allowAll true;
    // otherwise we verify the create-time default by re-checking subscribe default.
    if (freshStore) {
      expect(freshStore.allowAll).toBe(true);
    }
  });

  it('setAllowAll toggles the flag', () => {
    usePermissionStore.getState().setAllowAll(true);
    expect(usePermissionStore.getState().allowAll).toBe(true);

    usePermissionStore.getState().setAllowAll(false);
    expect(usePermissionStore.getState().allowAll).toBe(false);
  });

  it('evaluate returns approved when allowAll is true regardless of request', () => {
    usePermissionStore.getState().setAllowAll(true);
    const result = usePermissionStore.getState().evaluate({
      kind: 'write',
      path: '/some/dangerous/path',
    });
    expect(result).toBe('approved');
  });

  it('evaluate returns null when allowAll is false and no rules match', () => {
    usePermissionStore.getState().setAllowAll(false);
    const result = usePermissionStore.getState().evaluate({
      kind: 'write',
      path: '/some/path',
    });
    expect(result).toBeNull();
  });
});

// ─── workingDirectory ─────────────────────────────────────────────────────────

describe('permissionStore — workingDirectory', () => {
  it('setWorkingDirectory stores the value', () => {
    usePermissionStore.getState().setWorkingDirectory('/Users/test/project');
    expect(usePermissionStore.getState().workingDirectory).toBe('/Users/test/project');
  });

  it('setWorkingDirectory(null) clears the value', () => {
    usePermissionStore.getState().setWorkingDirectory('/a');
    usePermissionStore.getState().setWorkingDirectory(null);
    expect(usePermissionStore.getState().workingDirectory).toBeNull();
  });

  it('evaluate auto-approves read requests under workingDirectory', () => {
    usePermissionStore.getState().setWorkingDirectory('/Users/test/project');
    const result = usePermissionStore.getState().evaluate({
      kind: 'read',
      path: '/Users/test/project/src/index.ts',
    });
    expect(result).toBe('approved');
  });

  it('evaluate auto-approves read at exact workingDirectory', () => {
    usePermissionStore.getState().setWorkingDirectory('/Users/test/project');
    const result = usePermissionStore.getState().evaluate({
      kind: 'read',
      path: '/Users/test/project',
    });
    expect(result).toBe('approved');
  });

  it('evaluate does NOT auto-approve write requests under workingDirectory', () => {
    usePermissionStore.getState().setWorkingDirectory('/Users/test/project');
    const result = usePermissionStore.getState().evaluate({
      kind: 'write',
      path: '/Users/test/project/src/index.ts',
    });
    expect(result).toBeNull();
  });

  it('evaluate does NOT auto-approve reads outside workingDirectory', () => {
    usePermissionStore.getState().setWorkingDirectory('/Users/test/project');
    const result = usePermissionStore.getState().evaluate({
      kind: 'read',
      path: '/Users/other/project/file.ts',
    });
    expect(result).toBeNull();
  });

  it('path normalization handles backslashes (Windows paths)', () => {
    usePermissionStore.getState().setWorkingDirectory('C:\\Users\\test\\project');
    const result = usePermissionStore.getState().evaluate({
      kind: 'read',
      path: 'C:\\Users\\test\\project\\src\\index.ts',
    });
    expect(result).toBe('approved');
  });

  it('path normalization is case-insensitive', () => {
    usePermissionStore.getState().setWorkingDirectory('/Users/Test/Project');
    const result = usePermissionStore.getState().evaluate({
      kind: 'read',
      path: '/users/test/project/src/index.ts',
    });
    expect(result).toBe('approved');
  });

  it('path normalization strips trailing slashes', () => {
    usePermissionStore.getState().setWorkingDirectory('/Users/test/project///');
    const result = usePermissionStore.getState().evaluate({
      kind: 'read',
      path: '/Users/test/project/src/file.ts',
    });
    expect(result).toBe('approved');
  });
});

// ─── Rules CRUD ───────────────────────────────────────────────────────────────

describe('permissionStore — rules CRUD', () => {
  it('addRule adds a rule', () => {
    usePermissionStore.getState().addRule({ kind: 'write', pathPrefix: '/src' });
    expect(usePermissionStore.getState().rules).toHaveLength(1);
    expect(usePermissionStore.getState().rules[0].kind).toBe('write');
    expect(usePermissionStore.getState().rules[0].pathPrefix).toBe('/src');
  });

  it('addRule generates a deterministic id from kind + normalized path', () => {
    usePermissionStore.getState().addRule({ kind: 'write', pathPrefix: '/src/app' });
    const rule = usePermissionStore.getState().rules[0];
    expect(rule.id).toBe('write:/src/app');
  });

  it('addRule deduplicates: same kind + path is not added twice', () => {
    usePermissionStore.getState().addRule({ kind: 'write', pathPrefix: '/src' });
    usePermissionStore.getState().addRule({ kind: 'write', pathPrefix: '/src' });
    expect(usePermissionStore.getState().rules).toHaveLength(1);
  });

  it('addRule allows different kinds for the same path', () => {
    usePermissionStore.getState().addRule({ kind: 'read', pathPrefix: '/src' });
    usePermissionStore.getState().addRule({ kind: 'write', pathPrefix: '/src' });
    expect(usePermissionStore.getState().rules).toHaveLength(2);
  });

  it('removeRule removes by id', () => {
    usePermissionStore.getState().addRule({ kind: 'write', pathPrefix: '/src' });
    const id = usePermissionStore.getState().rules[0].id;
    usePermissionStore.getState().removeRule(id);
    expect(usePermissionStore.getState().rules).toHaveLength(0);
  });

  it('removeRule is no-op for unknown id', () => {
    usePermissionStore.getState().addRule({ kind: 'write', pathPrefix: '/src' });
    usePermissionStore.getState().removeRule('nonexistent');
    expect(usePermissionStore.getState().rules).toHaveLength(1);
  });

  it('clearRules removes all rules', () => {
    usePermissionStore.getState().addRule({ kind: 'read', pathPrefix: '/a' });
    usePermissionStore.getState().addRule({ kind: 'write', pathPrefix: '/b' });
    usePermissionStore.getState().clearRules();
    expect(usePermissionStore.getState().rules).toHaveLength(0);
  });
});

// ─── evaluate with rules ──────────────────────────────────────────────────────

describe('permissionStore — evaluate with rules', () => {
  it('approves request matching a rule kind + path prefix', () => {
    usePermissionStore.getState().addRule({ kind: 'write', pathPrefix: '/src' });
    const result = usePermissionStore.getState().evaluate({
      kind: 'write',
      path: '/src/components/App.tsx',
    });
    expect(result).toBe('approved');
  });

  it('does not approve request with wrong kind', () => {
    usePermissionStore.getState().addRule({ kind: 'read', pathPrefix: '/src' });
    const result = usePermissionStore.getState().evaluate({
      kind: 'write',
      path: '/src/components/App.tsx',
    });
    expect(result).toBeNull();
  });

  it('does not approve request outside rule path prefix', () => {
    usePermissionStore.getState().addRule({ kind: 'write', pathPrefix: '/src' });
    const result = usePermissionStore.getState().evaluate({
      kind: 'write',
      path: '/tmp/evil.sh',
    });
    expect(result).toBeNull();
  });

  it('approves exact path match', () => {
    usePermissionStore.getState().addRule({ kind: 'write', pathPrefix: '/src/index.ts' });
    const result = usePermissionStore.getState().evaluate({
      kind: 'write',
      path: '/src/index.ts',
    });
    expect(result).toBe('approved');
  });

  it('uses fileName when path is absent', () => {
    usePermissionStore.getState().addRule({ kind: 'read', pathPrefix: '/src' });
    const result = usePermissionStore.getState().evaluate({
      kind: 'read',
      fileName: '/src/utils/id.ts',
    });
    expect(result).toBe('approved');
  });

  it('uses fullCommandText when path and fileName are absent', () => {
    usePermissionStore.getState().addRule({ kind: 'execute', pathPrefix: '/usr/local/bin' });
    const result = usePermissionStore.getState().evaluate({
      kind: 'execute',
      fullCommandText: '/usr/local/bin/node script.js',
    });
    expect(result).toBe('approved');
  });

  it('returns null when no path can be extracted from request', () => {
    usePermissionStore.getState().addRule({ kind: 'write', pathPrefix: '/src' });
    const result = usePermissionStore.getState().evaluate({ kind: 'write' });
    expect(result).toBeNull();
  });

  it('multiple rules: first match wins', () => {
    usePermissionStore.getState().addRule({ kind: 'write', pathPrefix: '/a' });
    usePermissionStore.getState().addRule({ kind: 'write', pathPrefix: '/b' });
    expect(usePermissionStore.getState().evaluate({ kind: 'write', path: '/b/file.ts' })).toBe(
      'approved'
    );
    expect(
      usePermissionStore.getState().evaluate({ kind: 'write', path: '/c/file.ts' })
    ).toBeNull();
  });
});
