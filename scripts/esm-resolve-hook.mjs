// Resolve hook: fix extensionless subpath imports for packages without exports maps.
// vscode-jsonrpc 8.x exposes "node.js" but @github/copilot-sdk imports "node" (no extension).
export async function resolve(specifier, context, nextResolve) {
  if (specifier === 'vscode-jsonrpc/node') {
    return nextResolve('vscode-jsonrpc/node.js', context);
  }
  return nextResolve(specifier, context);
}
