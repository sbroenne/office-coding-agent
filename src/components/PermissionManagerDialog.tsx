import React, { useEffect, useMemo, useState } from 'react';
import { Shield, Folder, Trash2 } from 'lucide-react';
import {
  Dialog,
  DialogContent,
  DialogDescription,
  DialogHeader,
  DialogTitle,
  DialogTrigger,
} from '@/components/ui/dialog';
import { usePermissionStore } from '@/stores';

interface BrowseResponse {
  path: string;
  parent: string | null;
  dirs: string[];
  error?: string;
}

export const PermissionManagerDialog: React.FC = () => {
  const [open, setOpen] = useState(false);
  const [browseOpen, setBrowseOpen] = useState(false);
  const [browsePath, setBrowsePath] = useState<string>('');
  const [browseParent, setBrowseParent] = useState<string | null>(null);
  const [browseDirs, setBrowseDirs] = useState<string[]>([]);
  const [browseLoading, setBrowseLoading] = useState(false);

  const allowAll = usePermissionStore(s => s.allowAll);
  const workingDirectory = usePermissionStore(s => s.workingDirectory);
  const rules = usePermissionStore(s => s.rules);
  const setAllowAll = usePermissionStore(s => s.setAllowAll);
  const setWorkingDirectory = usePermissionStore(s => s.setWorkingDirectory);
  const removeRule = usePermissionStore(s => s.removeRule);
  const clearRules = usePermissionStore(s => s.clearRules);

  const sortedRules = useMemo(
    () => [...rules].sort((a, b) => a.pathPrefix.localeCompare(b.pathPrefix)),
    [rules]
  );

  const loadDir = async (pathValue?: string) => {
    setBrowseLoading(true);
    try {
      const query = pathValue ? `?path=${encodeURIComponent(pathValue)}` : '';
      const response = await fetch(`/api/browse${query}`);
      const data = (await response.json()) as BrowseResponse;
      if (data.error) return;
      setBrowsePath(data.path);
      setBrowseParent(data.parent);
      setBrowseDirs(data.dirs ?? []);
    } finally {
      setBrowseLoading(false);
    }
  };

  useEffect(() => {
    if (!open) return;
    if (workingDirectory) {
      void loadDir(workingDirectory);
      return;
    }
    void (async () => {
      const response = await fetch('/api/env');
      const data = (await response.json()) as { cwd?: string; home?: string };
      void loadDir(data.cwd ?? data.home);
    })();
  }, [open, workingDirectory]);

  return (
    <Dialog open={open} onOpenChange={setOpen}>
      <DialogTrigger asChild>
        <button
          className="inline-flex h-8 w-8 items-center justify-center rounded-md text-muted-foreground hover:bg-accent hover:text-accent-foreground transition-colors"
          aria-label="Permissions"
          title="Permissions"
        >
          <Shield className="size-4" />
        </button>
      </DialogTrigger>

      <DialogContent className="max-w-[560px]">
        <DialogHeader>
          <DialogTitle>Permissions</DialogTitle>
          <DialogDescription>
            Manage auto-approval behavior and saved permission rules.
          </DialogDescription>
        </DialogHeader>

        <div className="space-y-4 text-sm">
          <div className="flex items-center justify-between rounded-md border border-border p-3">
            <div>
              <div className="font-medium">Allow all</div>
              <div className="text-xs text-muted-foreground">
                Auto-approve all permission requests.
              </div>
            </div>
            <button
              type="button"
              onClick={() => setAllowAll(!allowAll)}
              className={`inline-flex rounded-md px-2 py-1 text-xs font-medium ${
                allowAll ? 'bg-primary text-primary-foreground' : 'bg-muted text-muted-foreground'
              }`}
            >
              {allowAll ? 'On' : 'Off'}
            </button>
          </div>

          <div className="rounded-md border border-border p-3">
            <div className="mb-1 font-medium">Working directory</div>
            <div className="text-xs text-muted-foreground break-all">
              {workingDirectory ?? 'Not set'}
            </div>
            <div className="mt-2 flex items-center gap-2">
              <button
                type="button"
                onClick={() => setBrowseOpen(v => !v)}
                className="inline-flex items-center gap-1 rounded-md border border-border px-2 py-1 text-xs hover:bg-accent"
              >
                <Folder className="size-3" /> Browse
              </button>
              {workingDirectory && (
                <button
                  type="button"
                  onClick={() => setWorkingDirectory(null)}
                  className="inline-flex rounded-md border border-border px-2 py-1 text-xs hover:bg-accent"
                >
                  Clear
                </button>
              )}
            </div>

            {browseOpen && (
              <div className="mt-2 rounded-md border border-border p-2">
                <div className="mb-2 flex items-center justify-between text-xs">
                  <span className="truncate pr-2">{browsePath || '(loading...)'}</span>
                  {browseParent && (
                    <button
                      type="button"
                      onClick={() => void loadDir(browseParent)}
                      className="rounded border border-border px-2 py-0.5 hover:bg-accent"
                    >
                      Up
                    </button>
                  )}
                </div>
                <div className="max-h-36 overflow-auto rounded border border-border">
                  {browseLoading ? (
                    <div className="px-2 py-1.5 text-xs text-muted-foreground">Loading…</div>
                  ) : browseDirs.length === 0 ? (
                    <div className="px-2 py-1.5 text-xs text-muted-foreground">
                      No subdirectories
                    </div>
                  ) : (
                    browseDirs.map(dir => (
                      <button
                        type="button"
                        key={dir}
                        onClick={() => void loadDir(`${browsePath}/${dir}`)}
                        className="flex w-full items-center gap-1 px-2 py-1 text-left text-xs hover:bg-accent"
                      >
                        <Folder className="size-3" /> {dir}
                      </button>
                    ))
                  )}
                </div>
                <div className="mt-2 flex justify-end">
                  <button
                    type="button"
                    onClick={() => {
                      setWorkingDirectory(browsePath || null);
                      setBrowseOpen(false);
                    }}
                    className="rounded-md bg-primary px-2 py-1 text-xs text-primary-foreground hover:opacity-90"
                  >
                    Select
                  </button>
                </div>
              </div>
            )}
          </div>

          <div className="rounded-md border border-border p-3">
            <div className="mb-2 flex items-center justify-between">
              <div className="font-medium">Saved rules</div>
              {sortedRules.length > 0 && (
                <button
                  type="button"
                  onClick={clearRules}
                  className="rounded-md border border-border px-2 py-1 text-xs hover:bg-accent"
                >
                  Clear all
                </button>
              )}
            </div>

            {sortedRules.length === 0 ? (
              <div className="text-xs text-muted-foreground">No saved rules.</div>
            ) : (
              <div className="max-h-40 space-y-1 overflow-auto">
                {sortedRules.map(rule => (
                  <div
                    key={rule.id}
                    className="flex items-center justify-between gap-2 rounded border border-border px-2 py-1"
                  >
                    <div className="min-w-0 text-xs">
                      <div className="font-medium">{rule.kind}</div>
                      <div className="truncate text-muted-foreground">{rule.pathPrefix}</div>
                    </div>
                    <button
                      type="button"
                      onClick={() => removeRule(rule.id)}
                      className="rounded p-1 text-muted-foreground hover:bg-accent hover:text-foreground"
                      aria-label="Remove rule"
                      title="Remove rule"
                    >
                      <Trash2 className="size-3" />
                    </button>
                  </div>
                ))}
              </div>
            )}
          </div>
        </div>
      </DialogContent>
    </Dialog>
  );
};
