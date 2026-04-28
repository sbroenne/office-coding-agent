import { describe, expect, it, vi } from 'vitest';

// @ts-expect-error copilotProxy is a Node .mjs module without TypeScript declarations.
const proxyModule = await import('../../src/copilotProxy.mjs');
const {
  addLoginHintToAuthorizationUrl,
  initiateMcpOAuthForSession,
  normalizeMcpEvent,
  normalizeMcpOAuthLoginHint,
  toMcpServerKey,
} = proxyModule;

describe('MCP OAuth proxy helpers', () => {
  it.each([
    ['Power BI MCP', 'power_bi_mcp'],
    [' powerbi ', 'powerbi'],
    ['Power  BI', 'power_bi'],
  ])('normalizes %s to %s', (name, expected) => {
    expect(toMcpServerKey(name)).toBe(expected);
  });

  it('initiates SDK MCP OAuth with the normalized server key', async () => {
    const login = vi
      .fn()
      .mockResolvedValue({ authorizationUrl: 'https://login.example/authorize' });
    const session = { rpc: { mcp: { oauth: { login } } } };

    const result = await initiateMcpOAuthForSession(session, 'Power BI MCP');

    expect(result).toEqual({
      status: 'success',
      authorizationUrl: 'https://login.example/authorize',
    });
    expect(login).toHaveBeenCalledWith({
      serverName: 'power_bi_mcp',
      forceReauth: false,
      clientName: 'office-coding-agent',
      callbackSuccessMessage: 'You can return to Office Coding Agent.',
    });
  });

  it('normalizes MCP OAuth login hints for Microsoft aliases', () => {
    expect(normalizeMcpOAuthLoginHint('janesmith')).toBe('janesmith@microsoft.com');
    expect(normalizeMcpOAuthLoginHint(' jane@contoso.com ')).toBe('jane@contoso.com');
    expect(normalizeMcpOAuthLoginHint('   ')).toBeUndefined();
  });

  it('adds login_hint to authorization URLs when a login hint is provided', () => {
    expect(
      addLoginHintToAuthorizationUrl('https://login.example/authorize?client_id=1', 'janesmith')
    ).toBe('https://login.example/authorize?client_id=1&login_hint=janesmith%40microsoft.com');
  });

  it('passes forceReauth and returns oauthAlias when a login hint is provided', async () => {
    const login = vi
      .fn()
      .mockResolvedValue({ authorizationUrl: 'https://login.example/authorize' });
    const session = { rpc: { mcp: { oauth: { login } } } };

    const result = await initiateMcpOAuthForSession(session, 'Power BI MCP', 'janesmith');

    expect(result).toEqual({
      status: 'success',
      authorizationUrl: 'https://login.example/authorize?login_hint=janesmith%40microsoft.com',
      oauthAlias: 'janesmith@microsoft.com',
    });
    expect(login).toHaveBeenCalledWith({
      serverName: 'power_bi_mcp',
      forceReauth: true,
      clientName: 'office-coding-agent',
      callbackSuccessMessage: 'You can return to Office Coding Agent.',
    });
  });

  it('returns a typed error when the SDK session has no MCP OAuth RPC', async () => {
    const result = await initiateMcpOAuthForSession({}, 'powerbi');

    expect(result).toEqual({
      status: 'error',
      message: 'Current Copilot SDK session does not support MCP OAuth.',
    });
  });

  it('normalizes SDK OAuth-required events into UI auth and status notifications', () => {
    expect(
      normalizeMcpEvent({
        type: 'mcp.oauth_required',
        data: {
          requestId: 'request-1',
          serverName: 'Power BI MCP',
          serverUrl: 'https://api.fabric.microsoft.com/v1/mcp/powerbi',
        },
      })
    ).toEqual([
      {
        method: 'mcp.oauth-required',
        params: {
          requestId: 'request-1',
          serverName: 'Power BI MCP',
          serverUrl: 'https://api.fabric.microsoft.com/v1/mcp/powerbi',
        },
      },
      {
        method: 'mcp.status',
        params: {
          server: 'Power BI MCP',
          status: 'needs-auth',
        },
      },
    ]);
  });

  it('fans out SDK MCP server load status notifications', () => {
    expect(
      normalizeMcpEvent({
        type: 'session.mcp_servers_loaded',
        data: {
          servers: [
            { name: 'powerbi', status: 'needs-auth' },
            { name: 'workiq', status: 'connected' },
          ],
        },
      })
    ).toEqual([
      {
        method: 'mcp.status',
        params: { server: 'powerbi', status: 'needs-auth', error: undefined },
      },
      {
        method: 'mcp.status',
        params: { server: 'workiq', status: 'connected', error: undefined },
      },
    ]);
  });
});
