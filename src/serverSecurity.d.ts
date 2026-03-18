declare module '*/serverSecurity.mjs' {
  export function isAllowedOrigin(origin: string): boolean;
  export function isLoopbackAddress(address: string): boolean;
  export function isTrustedRequestOrigin(origin: string | undefined, remoteAddress: string | undefined): boolean;
  export function getBrowseRoots(): Promise<string[]>;
  export function isPathWithinRoot(rootPath: string, targetPath: string): boolean;
  export function resolveBrowsePath(requestedPath: string | undefined, allowedRoots: string[]): Promise<string>;
}
