/**
 * Integration test: mcpStatusStore — ephemeral Zustand store for per-server MCP state.
 *
 * Tests status transitions, log append/cap/clear, tool discovery, and clearAll.
 */
import { describe, it, expect, beforeEach } from 'vitest';
import { useMcpStatusStore } from '@/stores/mcpStatusStore';
import type { McpLogEntry } from '@/types';

beforeEach(() => {
  useMcpStatusStore.getState().clearAll();
});

describe('mcpStatusStore — status', () => {
  it('starts with empty servers map', () => {
    expect(useMcpStatusStore.getState().servers).toEqual({});
  });

  it('setStatus creates server entry if not present', () => {
    useMcpStatusStore.getState().setStatus('my-server', 'starting');
    const server = useMcpStatusStore.getState().servers['my-server'];
    expect(server).toBeDefined();
    expect(server.status).toBe('starting');
    expect(server.tools).toEqual([]);
    expect(server.logs).toEqual([]);
  });

  it('setStatus updates existing server status', () => {
    useMcpStatusStore.getState().setStatus('s1', 'starting');
    useMcpStatusStore.getState().setStatus('s1', 'connected');
    expect(useMcpStatusStore.getState().servers['s1'].status).toBe('connected');
  });

  it('setStatus stores error string', () => {
    useMcpStatusStore.getState().setStatus('s1', 'error', 'connection refused');
    const server = useMcpStatusStore.getState().servers['s1'];
    expect(server.status).toBe('error');
    expect(server.error).toBe('connection refused');
  });

  it('setStatus clears error when status changes to non-error', () => {
    useMcpStatusStore.getState().setStatus('s1', 'error', 'timeout');
    useMcpStatusStore.getState().setStatus('s1', 'connected');
    expect(useMcpStatusStore.getState().servers['s1'].error).toBeUndefined();
  });

  it('handles multiple servers independently', () => {
    useMcpStatusStore.getState().setStatus('a', 'connected');
    useMcpStatusStore.getState().setStatus('b', 'error', 'fail');
    expect(useMcpStatusStore.getState().servers['a'].status).toBe('connected');
    expect(useMcpStatusStore.getState().servers['b'].status).toBe('error');
    expect(useMcpStatusStore.getState().servers['b'].error).toBe('fail');
  });
});

describe('mcpStatusStore — logs', () => {
  const makeLog = (msg: string, level: 'info' | 'warn' | 'error' = 'info'): McpLogEntry => ({
    timestamp: '2025-01-01T00:00:00Z',
    level,
    message: msg,
  });

  it('addLog appends to server log', () => {
    useMcpStatusStore.getState().addLog('s1', makeLog('hello'));
    useMcpStatusStore.getState().addLog('s1', makeLog('world'));
    expect(useMcpStatusStore.getState().servers['s1'].logs).toHaveLength(2);
    expect(useMcpStatusStore.getState().servers['s1'].logs[0].message).toBe('hello');
    expect(useMcpStatusStore.getState().servers['s1'].logs[1].message).toBe('world');
  });

  it('addLog creates server entry if not present', () => {
    useMcpStatusStore.getState().addLog('new-server', makeLog('first'));
    expect(useMcpStatusStore.getState().servers['new-server']).toBeDefined();
    expect(useMcpStatusStore.getState().servers['new-server'].status).toBe('stopped');
  });

  it('addLog trims logs to MAX_LOGS_PER_SERVER (500)', () => {
    for (let i = 0; i < 510; i++) {
      useMcpStatusStore.getState().addLog('s1', makeLog(`log-${i}`));
    }
    const logs = useMcpStatusStore.getState().servers['s1'].logs;
    expect(logs.length).toBe(500);
    // Oldest should be trimmed — first remaining should be log-10
    expect(logs[0].message).toBe('log-10');
    expect(logs[499].message).toBe('log-509');
  });

  it('clearLogs clears logs for a specific server', () => {
    useMcpStatusStore.getState().addLog('s1', makeLog('a'));
    useMcpStatusStore.getState().addLog('s2', makeLog('b'));
    useMcpStatusStore.getState().clearLogs('s1');
    expect(useMcpStatusStore.getState().servers['s1'].logs).toEqual([]);
    expect(useMcpStatusStore.getState().servers['s2'].logs).toHaveLength(1);
  });

  it('clearLogs is no-op for unknown server', () => {
    const before = useMcpStatusStore.getState();
    useMcpStatusStore.getState().clearLogs('nonexistent');
    expect(useMcpStatusStore.getState()).toBe(before);
  });
});

describe('mcpStatusStore — tools', () => {
  it('setTools stores discovered tools', () => {
    const tools = [
      { name: 'search', description: 'Search things' },
      { name: 'create', description: 'Create things' },
    ];
    useMcpStatusStore.getState().setTools('s1', tools);
    expect(useMcpStatusStore.getState().servers['s1'].tools).toEqual(tools);
  });

  it('setTools replaces existing tools', () => {
    useMcpStatusStore.getState().setTools('s1', [{ name: 'old', description: 'old' }]);
    useMcpStatusStore.getState().setTools('s1', [{ name: 'new', description: 'new' }]);
    expect(useMcpStatusStore.getState().servers['s1'].tools).toHaveLength(1);
    expect(useMcpStatusStore.getState().servers['s1'].tools[0].name).toBe('new');
  });

  it('setTools creates server entry if not present', () => {
    useMcpStatusStore.getState().setTools('new-srv', []);
    expect(useMcpStatusStore.getState().servers['new-srv']).toBeDefined();
  });
});

describe('mcpStatusStore — clearAll', () => {
  it('clears all server state', () => {
    useMcpStatusStore.getState().setStatus('a', 'connected');
    useMcpStatusStore.getState().setStatus('b', 'error');
    useMcpStatusStore.getState().addLog('a', { timestamp: 't', level: 'info', message: 'hi' });
    useMcpStatusStore.getState().clearAll();
    expect(useMcpStatusStore.getState().servers).toEqual({});
  });
});
