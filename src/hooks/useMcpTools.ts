import { useEffect, useRef, useState } from 'react';
import { createMCPClient } from '@ai-sdk/mcp';
import type { ToolSet } from 'ai';
import type { McpServerConfig } from '@/types';

/**
 * Connects to a list of MCP servers and returns a merged ToolSet.
 *
 * - Re-connects whenever the server list meaningfully changes (name, url, or transport).
 * - Silently skips servers that fail to connect (logs a warning).
 * - Closes all connections on cleanup.
 */
export function useMcpTools(configs: McpServerConfig[]): ToolSet {
  const [tools, setTools] = useState<ToolSet>({});

  // Stable key: re-run effect only when connection parameters change.
  const configKey = configs.map(c => `${c.name}|${c.url}|${c.transport}`).join(',');

  // Ref lets the async load function see the latest configs even after re-renders.
  const configsRef = useRef(configs);
  configsRef.current = configs;

  useEffect(() => {
    const currentConfigs = configsRef.current;

    if (currentConfigs.length === 0) {
      setTools({});
      return;
    }

    let cancelled = false;
    const clients: { close: () => Promise<void> }[] = [];

    async function loadTools() {
      const merged: ToolSet = {};

      for (const config of currentConfigs) {
        try {
          const client = await createMCPClient({
            transport: { type: config.transport, url: config.url, headers: config.headers },
            name: config.name,
          });
          clients.push(client);
          const serverTools = await client.tools();
          Object.assign(merged, serverTools);
        } catch (err) {
          console.warn(`[MCP] Failed to connect to "${config.name}":`, err);
        }
      }

      if (!cancelled) {
        setTools(merged);
      }
    }

    void loadTools();

    return () => {
      cancelled = true;
      for (const c of clients) {
        void c.close();
      }
    };
  }, [configKey]);

  return tools;
}
