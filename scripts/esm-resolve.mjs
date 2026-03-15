// Custom ESM resolve hook for vscode-jsonrpc subpath imports.
// @github/copilot-sdk imports "vscode-jsonrpc/node" (extensionless)
// but vscode-jsonrpc 8.x has no exports map, so Node ESM resolution fails.
// This hook appends ".js" when needed.
import { register } from 'node:module';
register('./esm-resolve-hook.mjs', import.meta.url);
