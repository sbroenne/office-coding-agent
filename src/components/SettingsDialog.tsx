import React, { useState, useCallback } from 'react';
import {
  Settings,
  Plus,
  Trash2,
  Pencil,
  Check,
  X,
  Eye,
  EyeOff,
  ChevronDown,
  Globe,
  Loader2,
} from 'lucide-react';
import {
  Dialog,
  DialogContent,
  DialogDescription,
  DialogFooter,
  DialogHeader,
  DialogTitle,
  DialogTrigger,
} from '@/components/ui/dialog';
import { Button } from '@/components/ui/button';
import { Input } from '@/components/ui/input';
import { Label } from '@/components/ui/label';
import { Separator } from '@/components/ui/separator';
import * as Popover from '@radix-ui/react-popover';
import { useSettingsStore } from '@/stores';
import { validateModelDeployment } from '@/services/ai';
import { PROVIDER_ORDER, getProviderConfig } from '@/services/ai/providerConfig';
import { ModelManager } from './ModelManager';
import type { FoundryEndpoint, ProviderType } from '@/types';

/** Extract hostname from a resource URL for display */
function extractHost(url: string): string {
  try {
    return new URL(url).hostname;
  } catch {
    return url;
  }
}

interface EndpointFormData {
  displayName: string;
  resourceUrl: string;
  apiKey: string;
  providerType: ProviderType;
}

// ─── Connection Form (shared for add / edit) ───

interface ConnectionFormProps {
  title: string;
  formData: EndpointFormData;
  onChange: (data: EndpointFormData) => void;
  onSave: () => void;
  onCancel: () => void;
  saveDisabled?: boolean;
  saveLabel?: string;
  placeholders?: boolean;
  testEndpoint?: FoundryEndpoint;
  /** When true the provider selector is locked (editing an existing endpoint) */
  lockProvider?: boolean;
}

const ConnectionForm: React.FC<ConnectionFormProps> = ({
  title,
  formData,
  onChange,
  onSave,
  onCancel,
  saveDisabled,
  saveLabel = 'Save',
  placeholders = false,
  testEndpoint,
  lockProvider = false,
}) => {
  const [showApiKey, setShowApiKey] = useState(false);
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [providerMenuOpen, setProviderMenuOpen] = useState(false);

  const cfg = getProviderConfig(formData.providerType);
  const isAzure = formData.providerType === 'azure';

  const trimmedDisplayName = formData.displayName.trim();
  const trimmedResourceUrl = isAzure ? formData.resourceUrl.trim() : (cfg.baseUrl ?? '');
  const trimmedApiKey = formData.apiKey.trim();
  const hasRequiredFields = Boolean(trimmedDisplayName && trimmedResourceUrl && trimmedApiKey);

  const handleSave = useCallback(async () => {
    const url = isAzure ? formData.resourceUrl.trim().replace(/\/+$/, '') : (cfg.baseUrl ?? '');
    const key = trimmedApiKey;
    if (!hasRequiredFields) {
      setError('Display name and API key are required.');
      return;
    }

    setSaving(true);
    setError(null);

    try {
      const epConfig: FoundryEndpoint = {
        id: testEndpoint?.id ?? '__test_connection__',
        displayName: formData.displayName || 'Test',
        resourceUrl: url,
        authMethod: 'apiKey',
        apiKey: key,
        providerType: formData.providerType,
      };

      // For non-Azure, validate with the first known model for the provider.
      // For Azure, use an existing model or the default model ID.
      const state = useSettingsStore.getState();
      const existingModels = testEndpoint ? (state.endpointModels[testEndpoint.id] ?? []) : [];
      const knownModels = cfg.defaultModels;
      const testModelId = existingModels[0]?.id ?? knownModels[0] ?? state.defaultModelId;

      const ok = await validateModelDeployment(epConfig, testModelId);

      if (ok) {
        onSave();
      } else {
        setError('Could not connect. Check the URL and API key.');
      }
    } catch {
      setError('Could not connect. Check the URL and API key.');
    } finally {
      setSaving(false);
    }
  }, [
    cfg,
    formData.displayName,
    formData.providerType,
    formData.resourceUrl,
    hasRequiredFields,
    isAzure,
    onSave,
    testEndpoint,
    trimmedApiKey,
  ]);

  return (
    <div className="flex flex-col gap-3">
      <h3 className="text-sm font-semibold">{title}</h3>
      <p className="text-xs text-muted-foreground">All fields are required.</p>

      {/* ── Provider selector ── */}
      <div className="flex flex-col gap-1.5">
        <Label>Provider</Label>
        {lockProvider ? (
          <div className="rounded-md bg-secondary px-3 py-1.5 text-sm">{cfg.label}</div>
        ) : (
          <Popover.Root open={providerMenuOpen} onOpenChange={setProviderMenuOpen}>
            <Popover.Trigger asChild>
              <Button
                variant="secondary"
                size="sm"
                className="w-full justify-between"
                type="button"
              >
                <span>{cfg.label}</span>
                <ChevronDown className="size-3.5 opacity-50" />
              </Button>
            </Popover.Trigger>
            <Popover.Portal>
              <Popover.Content
                align="start"
                sideOffset={4}
                className="z-50 min-w-[200px] rounded-md border border-border bg-popover p-1 text-popover-foreground shadow-md"
              >
                {PROVIDER_ORDER.map(pt => {
                  const pcfg = getProviderConfig(pt);
                  return (
                    <button
                      key={pt}
                      onClick={() => {
                        const newUrl = pt === 'azure' ? '' : (pcfg.baseUrl ?? '');
                        const newName = pt === 'azure' ? '' : pcfg.label;
                        onChange({
                          ...formData,
                          providerType: pt,
                          resourceUrl: newUrl,
                          displayName: formData.displayName || newName,
                        });
                        setProviderMenuOpen(false);
                        setError(null);
                      }}
                      className="flex w-full items-center gap-2 rounded-sm px-2 py-1.5 text-sm hover:bg-accent hover:text-accent-foreground"
                    >
                      {pt === formData.providerType && <Check className="size-3.5" />}
                      {pt !== formData.providerType && <span className="w-3.5" />}
                      {pcfg.label}
                    </button>
                  );
                })}
              </Popover.Content>
            </Popover.Portal>
          </Popover.Root>
        )}
      </div>

      <div className="flex flex-col gap-1.5">
        <Label htmlFor="conn-name">Display Name</Label>
        <Input
          id="conn-name"
          value={formData.displayName}
          onChange={e => {
            onChange({ ...formData, displayName: e.target.value });
            setError(null);
          }}
          placeholder={placeholders ? (isAzure ? 'My AI Foundry Resource' : cfg.label) : undefined}
        />
        {!trimmedDisplayName && <p className="text-xs text-muted-foreground">Required</p>}
      </div>

      {/* ── Resource URL — only for Azure ── */}
      {isAzure && (
        <div className="flex flex-col gap-1.5">
          <Label htmlFor="conn-url">Resource URL</Label>
          <Input
            id="conn-url"
            value={formData.resourceUrl}
            onChange={e => {
              onChange({ ...formData, resourceUrl: e.target.value });
              setError(null);
            }}
            placeholder={placeholders ? 'https://your-resource.openai.azure.com' : undefined}
          />
          {!formData.resourceUrl.trim() ? (
            <p className="text-xs text-muted-foreground">Required</p>
          ) : (
            <p className="text-xs text-muted-foreground">
              e.g., https://my-resource.openai.azure.com
            </p>
          )}
        </div>
      )}

      {/* ── Fixed URL badge for non-Azure ── */}
      {!isAzure && cfg.baseUrl && (
        <div className="flex flex-col gap-1.5">
          <Label>API Endpoint</Label>
          <div className="flex items-center gap-2 rounded-md bg-secondary px-3 py-1.5">
            <Globe className="size-4 shrink-0 text-muted-foreground" />
            <span className="truncate text-sm text-muted-foreground">{cfg.baseUrl}</span>
          </div>
        </div>
      )}

      <div className="flex flex-col gap-1.5">
        <Label htmlFor="conn-key">API Key</Label>
        <div className="relative">
          <Input
            id="conn-key"
            type={showApiKey ? 'text' : 'password'}
            value={formData.apiKey}
            onChange={e => {
              onChange({ ...formData, apiKey: e.target.value });
              setError(null);
            }}
            placeholder={placeholders ? 'Enter API key' : undefined}
            className="pr-9"
          />
          <button
            type="button"
            onClick={() => setShowApiKey(prev => !prev)}
            className="absolute right-2 top-1/2 -translate-y-1/2 text-muted-foreground hover:text-foreground"
            aria-label={showApiKey ? 'Hide API key' : 'Show API key'}
          >
            {showApiKey ? <EyeOff className="size-4" /> : <Eye className="size-4" />}
          </button>
        </div>
        {!trimmedApiKey && <p className="text-xs text-muted-foreground">Required</p>}
      </div>

      {error && (
        <div
          className="rounded-md border border-destructive/30 bg-destructive/10 px-3 py-2 text-sm text-destructive"
          role="alert"
        >
          {error}
        </div>
      )}

      <div className="flex justify-end gap-2">
        <Button variant="secondary" size="sm" onClick={onCancel} disabled={saving}>
          Cancel
        </Button>
        <Button
          size="sm"
          onClick={() => void handleSave()}
          disabled={(saveDisabled ?? false) || !hasRequiredFields || saving}
        >
          {saving && <Loader2 className="size-4 animate-spin" />}
          {saving ? 'Saving…' : saveLabel}
        </Button>
      </div>
    </div>
  );
};

// ─── Main SettingsDialog ───

interface SettingsDialogProps {
  open?: boolean;
  onOpenChange?: (open: boolean) => void;
}

export const SettingsDialog: React.FC<SettingsDialogProps> = ({
  open: controlledOpen,
  onOpenChange: controlledOnOpenChange,
}) => {
  const [internalOpen, setInternalOpen] = useState(false);

  const open = controlledOpen ?? internalOpen;
  const setOpen = useCallback(
    (value: boolean) => {
      setInternalOpen(value);
      controlledOnOpenChange?.(value);
    },
    [controlledOnOpenChange]
  );

  // Form state
  const [editingNew, setEditingNew] = useState(false);
  const [editingConnection, setEditingConnection] = useState(false);
  const [formData, setFormData] = useState<EndpointFormData>({
    displayName: '',
    resourceUrl: '',
    apiKey: '',
    providerType: 'azure',
  });

  // Delete confirmation
  const [confirmingDelete, setConfirmingDelete] = useState(false);

  // Endpoint switcher
  const [switcherOpen, setSwitcherOpen] = useState(false);

  const {
    endpoints,
    activeEndpointId,
    addEndpoint,
    updateEndpoint,
    removeEndpoint,
    setActiveEndpoint,
  } = useSettingsStore();

  const activeEndpoint = endpoints.find(ep => ep.id === activeEndpointId);
  const hasMultipleEndpoints = endpoints.length > 1;

  // ── Handlers ──

  const handleStartAdd = useCallback(() => {
    setEditingNew(true);
    setEditingConnection(false);
    setFormData({ displayName: '', resourceUrl: '', apiKey: '', providerType: 'azure' });
  }, []);

  const handleSaveNew = useCallback(() => {
    if (!formData.displayName) return;
    const cfg = getProviderConfig(formData.providerType);
    const url =
      formData.providerType === 'azure'
        ? formData.resourceUrl.replace(/\/+$/, '')
        : (cfg.baseUrl ?? '');
    addEndpoint({
      displayName: formData.displayName,
      resourceUrl: url,
      authMethod: 'apiKey',
      apiKey: formData.apiKey,
      providerType: formData.providerType,
    });
    setEditingNew(false);
    setFormData({ displayName: '', resourceUrl: '', apiKey: '', providerType: 'azure' });
  }, [formData, addEndpoint]);

  const handleCancelAdd = useCallback(() => {
    setEditingNew(false);
    setFormData({ displayName: '', resourceUrl: '', apiKey: '', providerType: 'azure' });
  }, []);

  const handleStartEditConnection = useCallback(() => {
    if (!activeEndpoint) return;
    setEditingConnection(true);
    setEditingNew(false);
    setFormData({
      displayName: activeEndpoint.displayName,
      resourceUrl: activeEndpoint.resourceUrl,
      apiKey: activeEndpoint.apiKey ?? '',
      providerType: activeEndpoint.providerType ?? 'azure',
    });
  }, [activeEndpoint]);

  const handleSaveConnection = useCallback(() => {
    if (!activeEndpointId || !formData.displayName) return;
    const cfg = getProviderConfig(formData.providerType);
    const url =
      formData.providerType === 'azure'
        ? formData.resourceUrl.replace(/\/+$/, '')
        : (cfg.baseUrl ?? '');
    updateEndpoint(activeEndpointId, {
      displayName: formData.displayName,
      resourceUrl: url,
      apiKey: formData.apiKey,
      providerType: formData.providerType,
    });
    setEditingConnection(false);
    setFormData({ displayName: '', resourceUrl: '', apiKey: '', providerType: 'azure' });
  }, [activeEndpointId, formData, updateEndpoint]);

  const handleCancelEditConnection = useCallback(() => {
    setEditingConnection(false);
    setFormData({ displayName: '', resourceUrl: '', apiKey: '', providerType: 'azure' });
  }, []);

  const handleDeleteRequest = useCallback(() => {
    setConfirmingDelete(true);
  }, []);

  const handleDeleteConfirm = useCallback(() => {
    if (activeEndpointId) {
      removeEndpoint(activeEndpointId);
      setEditingConnection(false);
      setFormData({ displayName: '', resourceUrl: '', apiKey: '', providerType: 'azure' });
    }
    setConfirmingDelete(false);
  }, [activeEndpointId, removeEndpoint]);

  const handleDeleteCancel = useCallback(() => {
    setConfirmingDelete(false);
  }, []);

  const showConnectionSummary = activeEndpoint && !editingConnection && !editingNew;
  const showConnectionForm = editingConnection && activeEndpoint;
  const showAddForm = editingNew;
  const showModels = activeEndpoint && !editingNew;

  return (
    <Dialog open={open} onOpenChange={setOpen}>
      <DialogTrigger asChild>
        <button
          className="inline-flex h-8 items-center gap-1.5 rounded-md px-2 text-muted-foreground transition-colors hover:bg-accent hover:text-accent-foreground"
          aria-label="Settings"
          title="Settings"
        >
          <Settings className="size-4" />
          <span className="text-xs font-medium">Settings</span>
        </button>
      </DialogTrigger>

      <DialogContent className="max-w-[480px] max-h-[90vh] flex flex-col">
        <DialogHeader>
          <DialogTitle>Settings</DialogTitle>
          <DialogDescription>Configure model endpoints and runtime preferences.</DialogDescription>
        </DialogHeader>

        <div className="flex-1 overflow-y-auto space-y-4 pr-1">
          <div className="space-y-2 rounded-md border border-border bg-card p-3">
            <h3 className="text-xs font-semibold uppercase tracking-wide text-muted-foreground">
              Connection
            </h3>
            {/* ── Endpoint Switcher (only when multiple endpoints) ── */}
            {hasMultipleEndpoints && !editingNew && (
              <div className="flex items-center gap-2">
                <span className="text-xs text-muted-foreground">Active endpoint:</span>
                <Popover.Root open={switcherOpen} onOpenChange={setSwitcherOpen}>
                  <Popover.Trigger asChild>
                    <Button variant="secondary" size="sm" className="min-w-0 justify-between gap-1">
                      <span className="truncate">{activeEndpoint?.displayName ?? 'Select'}</span>
                      <ChevronDown className="size-3.5 shrink-0 opacity-50" />
                    </Button>
                  </Popover.Trigger>
                  <Popover.Portal>
                    <Popover.Content
                      align="start"
                      sideOffset={4}
                      className="z-50 min-w-[160px] rounded-md border border-border bg-popover p-1 text-popover-foreground shadow-md"
                    >
                      {endpoints.map(ep => (
                        <button
                          key={ep.id}
                          onClick={() => {
                            setActiveEndpoint(ep.id);
                            setEditingConnection(false);
                            setSwitcherOpen(false);
                          }}
                          className="flex w-full items-center gap-2 rounded-sm px-2 py-1.5 text-sm hover:bg-accent hover:text-accent-foreground"
                        >
                          {ep.id === activeEndpointId && <Check className="size-3.5" />}
                          {ep.id !== activeEndpointId && <span className="w-3.5" />}
                          {ep.displayName}
                        </button>
                      ))}
                    </Popover.Content>
                  </Popover.Portal>
                </Popover.Root>
              </div>
            )}

            {/* ── No endpoints yet ── */}
            {endpoints.length === 0 && !editingNew && (
              <div className="py-6 text-center text-sm text-muted-foreground">
                <p>No endpoints configured.</p>
                <Button className="mt-2" onClick={handleStartAdd}>
                  <Plus className="size-4" />
                  Add Endpoint
                </Button>
              </div>
            )}

            {/* ── Connection Summary (read-only view) ── */}
            {showConnectionSummary && (
              <div>
                <div className="flex items-center justify-between mb-1.5">
                  <h3 className="text-sm font-semibold">{activeEndpoint.displayName}</h3>
                  <Button
                    variant="secondary"
                    size="sm"
                    onClick={handleStartEditConnection}
                    aria-label="Edit endpoint"
                  >
                    <Pencil className="size-3.5" />
                    Edit
                  </Button>
                </div>
                <button
                  className="flex w-full flex-col gap-0.5 rounded-md bg-secondary p-2.5 text-left transition-colors hover:bg-secondary/80"
                  onClick={handleStartEditConnection}
                  title="Click to edit connection"
                >
                  <span className="text-xs text-muted-foreground">
                    {getProviderConfig(activeEndpoint.providerType ?? 'azure').label}
                  </span>
                  <div className="flex items-center gap-2">
                    <Globe className="size-4 shrink-0 text-muted-foreground" />
                    <span className="truncate text-sm" title={activeEndpoint.resourceUrl}>
                      {extractHost(activeEndpoint.resourceUrl)}
                    </span>
                  </div>
                </button>
              </div>
            )}

            {/* ── Edit Connection Form ── */}
            {showConnectionForm && (
              <div>
                <ConnectionForm
                  title="Edit Connection"
                  formData={formData}
                  onChange={setFormData}
                  onSave={handleSaveConnection}
                  onCancel={handleCancelEditConnection}
                  saveDisabled={!formData.displayName}
                  testEndpoint={activeEndpoint}
                  lockProvider
                />

                {/* Delete option */}
                {confirmingDelete ? (
                  <div className="mt-2 flex items-center gap-1.5">
                    <span className="text-xs text-destructive">
                      Delete this endpoint and all its models?
                    </span>
                    <Button
                      variant="ghost"
                      size="icon"
                      className="size-6 text-destructive hover:text-destructive"
                      onClick={handleDeleteConfirm}
                      aria-label="Confirm delete"
                    >
                      <Check className="size-3.5" />
                    </Button>
                    <Button
                      variant="ghost"
                      size="icon"
                      className="size-6"
                      onClick={handleDeleteCancel}
                      aria-label="Cancel delete"
                    >
                      <X className="size-3.5" />
                    </Button>
                  </div>
                ) : (
                  <Button
                    variant="ghost"
                    size="sm"
                    className="mt-2 text-destructive hover:text-destructive"
                    onClick={handleDeleteRequest}
                  >
                    <Trash2 className="size-3.5" />
                    Delete endpoint
                  </Button>
                )}
              </div>
            )}

            {/* ── Add New Endpoint Form ── */}
            {showAddForm && (
              <ConnectionForm
                title="Add Endpoint"
                formData={formData}
                onChange={setFormData}
                onSave={handleSaveNew}
                onCancel={handleCancelAdd}
                saveDisabled={!formData.displayName}
                placeholders
              />
            )}
          </div>

          {/* ── Models Section ── */}
          {showModels && (
            <div className="space-y-2 rounded-md border border-border bg-card p-3">
              <h3 className="text-xs font-semibold uppercase tracking-wide text-muted-foreground">
                Models
              </h3>
              <Separator className="my-1" />
              <ModelManager endpointId={activeEndpointId ?? ''} />
            </div>
          )}
        </div>

        <DialogFooter className="mt-2">
          <div className="flex flex-1 items-center">
            {endpoints.length > 0 && !editingNew && (
              <Button variant="secondary" size="sm" onClick={handleStartAdd}>
                <Plus className="size-3.5" />
                Add another endpoint
              </Button>
            )}
          </div>
          <Button variant="secondary" onClick={() => setOpen(false)}>
            Close
          </Button>
        </DialogFooter>
      </DialogContent>
    </Dialog>
  );
};
