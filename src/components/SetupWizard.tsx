import React, { useState, useCallback, useEffect } from 'react';
import { CheckCircle, XCircle, Plus, Trash2, Eye, EyeOff, Loader2, Info } from 'lucide-react';
import { Button } from '@/components/ui/button';
import { Input } from '@/components/ui/input';
import { Label } from '@/components/ui/label';
import { useSettingsStore } from '@/stores';
import { discoverModels, validateModelDeployment, invalidateClient } from '@/services/ai';
import { PROVIDER_ORDER, getProviderConfig } from '@/services/ai/providerConfig';
import type { DiscoveryResult } from '@/services/ai';
import type { ModelInfo, ModelProvider, ProviderType } from '@/types';

type Step = 'provider' | 'endpoint' | 'auth' | 'connecting' | 'models' | 'done';

/** Environment defaults injected at build time via DefinePlugin */
const ENV_ENDPOINT = process.env.AZURE_OPENAI_ENDPOINT ?? '';
const ENV_API_KEY = process.env.AZURE_OPENAI_API_KEY ?? '';

interface Props {
  onComplete: () => void;
}

/** Map a ProviderType to the ModelProvider label used for UI grouping */
function toModelProvider(pt: ProviderType): ModelProvider {
  switch (pt) {
    case 'anthropic':
      return 'Anthropic';
    case 'openai':
      return 'OpenAI';
    case 'mistral':
      return 'Mistral';
    case 'deepseek':
      return 'DeepSeek';
    case 'xai':
      return 'xAI';
    case 'azure':
    default:
      return 'Other';
  }
}

export const SetupWizard: React.FC<Props> = ({ onComplete }) => {
  const { addEndpoint, setModelsForEndpoint, setActiveEndpoint, defaultModelId } =
    useSettingsStore();

  const [step, setStep] = useState<Step>('provider');
  const [providerType, setProviderType] = useState<ProviderType>('azure');
  const [resourceUrl, setResourceUrl] = useState(ENV_ENDPOINT);
  const [displayName, setDisplayName] = useState('');
  const [apiKey, setApiKey] = useState(ENV_API_KEY);
  const [showApiKey, setShowApiKey] = useState(false);
  const [error, setError] = useState<string | null>(null);

  // Connection validation state
  const [connecting, setConnecting] = useState(false);
  const [discoveredModels, setDiscoveredModels] = useState<ModelInfo[]>([]);
  const [connectionError, setConnectionError] = useState<string | null>(null);

  // Manual / pre-populated model state
  const [manualModelName, setManualModelName] = useState('');
  const [manualModels, setManualModels] = useState<ModelInfo[]>([]);
  const [validatingModel, setValidatingModel] = useState(false);
  const [modelValidationError, setModelValidationError] = useState<string | null>(null);

  const config = getProviderConfig(providerType);

  // ─── Provider → next step ───
  const handleProviderNext = useCallback(() => {
    if (providerType === 'azure') {
      setStep('endpoint');
    } else {
      const cfg = getProviderConfig(providerType);
      setResourceUrl(cfg.baseUrl ?? '');
      setDisplayName(cfg.label);
      setStep('auth');
    }
  }, [providerType]);

  const handleEndpointNext = useCallback(() => {
    if (!resourceUrl.trim()) return;
    try {
      const url = new URL(resourceUrl.trim());
      setDisplayName(url.hostname.split('.')[0]);
    } catch {
      setDisplayName('My Endpoint');
    }
    setError(null);
    setStep('auth');
  }, [resourceUrl]);

  const handleFinish = useCallback(() => {
    if (!apiKey.trim()) return;
    setConnectionError(null);
    setDiscoveredModels([]);
    setStep('connecting');
  }, [apiKey]);

  // ─── Connection validation effect ───
  useEffect(() => {
    if (step !== 'connecting') return;
    let cancelled = false;

    async function validate() {
      setConnecting(true);
      setConnectionError(null);
      try {
        const cfg = getProviderConfig(providerType);
        const url =
          providerType === 'azure' ? resourceUrl.trim().replace(/\/+$/, '') : (cfg.baseUrl ?? '');

        const endpointConfig = {
          id: '__setup_validation__',
          displayName: displayName.trim() || cfg.label,
          resourceUrl: url,
          authMethod: 'apiKey' as const,
          apiKey: apiKey.trim(),
          providerType,
        };

        if (providerType !== 'azure') {
          // Non-Azure: skip discovery — use the provider's known default models
          if (cancelled) return;
          invalidateClient('__setup_validation__');
          const prePopulated: ModelInfo[] = cfg.defaultModels.map(id => ({
            id,
            name: id,
            ownedBy: 'system',
            provider: toModelProvider(providerType),
          }));
          setManualModels(prePopulated);
          setManualModelName('');
          setStep('models');
          return;
        }

        // Azure: run auto-discovery / manual fallback
        const discovery: DiscoveryResult = await discoverModels(endpointConfig, true);
        const { models, method } = discovery;
        if (cancelled) return;

        if (method === 'manual') {
          setManualModels([]);
          setManualModelName('');
          let defaultOk = false;
          try {
            defaultOk = await validateModelDeployment(endpointConfig, defaultModelId);
          } catch (err) {
            console.warn('[SetupWizard] Default model validation failed:', err);
          }
          if (!cancelled) {
            if (defaultOk) {
              setManualModels([
                { id: defaultModelId, name: defaultModelId, ownedBy: 'user', provider: 'OpenAI' },
              ]);
            }
            invalidateClient('__setup_validation__');
            setStep('models');
          }
          return;
        }

        if (models.length === 0) {
          setConnectionError(
            'Connected successfully, but no chat models were found on this endpoint. Please deploy a model and try again.'
          );
          setConnecting(false);
          return;
        }

        setDiscoveredModels(models);
        const endpointId = addEndpoint({
          displayName: endpointConfig.displayName,
          resourceUrl: endpointConfig.resourceUrl,
          authMethod: 'apiKey',
          apiKey: apiKey.trim(),
          providerType,
        });
        setModelsForEndpoint(endpointId, models);
        setActiveEndpoint(endpointId);
        invalidateClient('__setup_validation__');
        setStep('done');
      } catch (err) {
        if (cancelled) return;
        console.error('Connection validation failed:', err);
        setConnectionError(
          err instanceof Error
            ? err.message
            : 'Could not connect to the endpoint. Please check your URL and credentials.'
        );
      } finally {
        if (!cancelled) setConnecting(false);
      }
    }

    void validate();
    return () => {
      cancelled = true;
    };
  }, [
    step,
    displayName,
    resourceUrl,
    apiKey,
    providerType,
    addEndpoint,
    setModelsForEndpoint,
    setActiveEndpoint,
    defaultModelId,
  ]);

  // ─── Step: Provider selection ───
  if (step === 'provider') {
    return (
      <div className="flex h-screen flex-col overflow-hidden bg-background px-4 py-3 text-foreground">
        <div className="mb-4 mt-1 shrink-0">
          <h2 className="text-sm font-semibold">Choose Your AI Provider</h2>
          <p className="mt-1 text-xs text-muted-foreground">
            Select the provider you want to connect to.
          </p>
        </div>
        <div className="flex flex-1 flex-col gap-2 overflow-y-auto">
          {PROVIDER_ORDER.map(pt => {
            const cfg = getProviderConfig(pt);
            const isSelected = providerType === pt;
            return (
              <button
                key={pt}
                onClick={() => setProviderType(pt)}
                className={`flex flex-col gap-0.5 rounded-md border px-3 py-2.5 text-left transition-colors ${
                  isSelected
                    ? 'border-primary bg-primary/10 text-foreground'
                    : 'border-border bg-card hover:bg-accent'
                }`}
              >
                <span className="text-sm font-medium">{cfg.label}</span>
                {cfg.baseUrl ? (
                  <span className="text-xs text-muted-foreground">{cfg.baseUrl}</span>
                ) : (
                  <span className="text-xs text-muted-foreground">Custom resource URL</span>
                )}
              </button>
            );
          })}
        </div>
        <div className="flex shrink-0 justify-end gap-2 py-3">
          <Button onClick={handleProviderNext}>Next</Button>
        </div>
      </div>
    );
  }

  // ─── Step: Endpoint URL (Azure only) ───
  if (step === 'endpoint') {
    return (
      <div className="flex h-screen flex-col overflow-hidden bg-background px-4 py-3 text-foreground">
        <div className="mb-4 mt-1 shrink-0">
          <h2 className="text-sm font-semibold">Connect to Azure AI Foundry</h2>
          <p className="mt-1 text-xs text-muted-foreground">
            Enter your endpoint URL to get started.
          </p>
        </div>
        <div className="flex flex-1 flex-col gap-3 overflow-y-auto">
          <div className="flex flex-col gap-1.5">
            <Label htmlFor="resource-url">Resource URL</Label>
            <Input
              id="resource-url"
              value={resourceUrl}
              onChange={e => setResourceUrl(e.target.value)}
              placeholder="https://your-resource.openai.azure.com"
            />
            <p className="text-xs text-muted-foreground">
              e.g., https://my-resource.openai.azure.com
            </p>
          </div>
        </div>
        <div className="flex shrink-0 justify-end gap-2 py-3">
          <Button variant="secondary" onClick={() => setStep('provider')}>
            Back
          </Button>
          <Button onClick={handleEndpointNext} disabled={!resourceUrl.trim()}>
            Next
          </Button>
        </div>
      </div>
    );
  }

  // ─── Step: Auth Method ───
  if (step === 'auth') {
    return (
      <div className="flex h-screen flex-col overflow-hidden bg-background px-4 py-3 text-foreground">
        <div className="mb-4 mt-1 shrink-0">
          <h2 className="text-sm font-semibold">Authentication</h2>
          <p className="mt-1 text-xs text-muted-foreground">Enter your {config.label} API key.</p>
        </div>
        <div className="flex flex-1 flex-col gap-3 overflow-y-auto">
          {error && (
            <div className="rounded-md border border-red-300 bg-red-50 px-3 py-2 text-sm text-red-800 dark:border-red-700 dark:bg-red-900/30 dark:text-red-200">
              {error}
            </div>
          )}
          <div className="flex flex-col gap-1.5">
            <Label htmlFor="api-key">API Key</Label>
            <div className="relative">
              <Input
                id="api-key"
                type={showApiKey ? 'text' : 'password'}
                value={apiKey}
                onChange={e => setApiKey(e.target.value)}
                placeholder="Enter your API key"
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
          </div>
        </div>
        <div className="flex shrink-0 justify-end gap-2 py-3">
          <Button
            variant="secondary"
            onClick={() => setStep(providerType === 'azure' ? 'endpoint' : 'provider')}
          >
            Back
          </Button>
          <Button onClick={handleFinish} disabled={!apiKey.trim()}>
            Connect
          </Button>
        </div>
      </div>
    );
  }

  // ─── Step: Connecting / Validating ───
  if (step === 'connecting') {
    return (
      <div className="flex h-screen flex-col overflow-hidden bg-background px-4 py-3 text-foreground">
        <div className="mb-4 mt-1 shrink-0">
          <h2 className="text-sm font-semibold">Connecting...</h2>
          <p className="mt-1 text-xs text-muted-foreground">
            Verifying your endpoint and discovering available models.
          </p>
        </div>
        <div className="flex flex-1 flex-col items-center justify-center gap-3 text-center">
          {connecting && !connectionError && (
            <>
              <Loader2 className="size-10 animate-spin text-muted-foreground" />
              <p className="text-sm text-muted-foreground">Connecting to endpoint...</p>
            </>
          )}
          {connectionError && (
            <>
              <XCircle className="size-9 text-destructive" />
              <div className="rounded-md border border-red-300 bg-red-50 px-3 py-2 text-sm text-red-800 dark:border-red-700 dark:bg-red-900/30 dark:text-red-200">
                {connectionError}
              </div>
            </>
          )}
        </div>
        {connectionError && (
          <div className="flex shrink-0 justify-end gap-2 py-3">
            <Button variant="secondary" onClick={() => setStep('auth')}>
              Back
            </Button>
            <Button onClick={() => setStep('connecting')}>Retry</Button>
          </div>
        )}
      </div>
    );
  }

  // ─── Handler: Add a manually entered model ───
  const handleAddManualModel = async () => {
    const name = manualModelName.trim();
    if (!name) return;
    if (manualModels.some(m => m.id === name)) {
      setManualModelName('');
      return;
    }

    if (providerType === 'azure') {
      // Azure: validate the deployment name with the real API
      setValidatingModel(true);
      setModelValidationError(null);

      const endpointConfig = {
        id: '__setup_validation__',
        displayName: displayName.trim() || 'My Endpoint',
        resourceUrl: resourceUrl.trim().replace(/\/+$/, ''),
        authMethod: 'apiKey' as const,
        apiKey: apiKey.trim(),
        providerType,
      };

      const ok = await validateModelDeployment(endpointConfig, name);
      setValidatingModel(false);

      if (!ok) {
        setModelValidationError(
          `Model "${name}" could not be reached. Verify the deployment name and try again.`
        );
        return;
      }
    }

    setManualModels(prev => [
      ...prev,
      { id: name, name, ownedBy: 'user', provider: toModelProvider(providerType) },
    ]);
    setManualModelName('');
    setModelValidationError(null);
  };

  const handleModelsFinish = () => {
    if (manualModels.length === 0) return;
    setDiscoveredModels(manualModels);
    const cfg = getProviderConfig(providerType);
    const url =
      providerType === 'azure' ? resourceUrl.trim().replace(/\/+$/, '') : (cfg.baseUrl ?? '');
    const endpointId = addEndpoint({
      displayName: displayName.trim() || cfg.label,
      resourceUrl: url,
      authMethod: 'apiKey',
      apiKey: apiKey.trim(),
      providerType,
    });
    setModelsForEndpoint(endpointId, manualModels);
    setActiveEndpoint(endpointId);
    setStep('done');
  };

  // ─── Step: Model selection / manual entry ───
  if (step === 'models') {
    const isAzure = providerType === 'azure';
    return (
      <div className="flex h-screen flex-col overflow-hidden bg-background px-4 py-3 text-foreground">
        <div className="mb-4 mt-1 shrink-0">
          <h2 className="text-sm font-semibold">Select Models</h2>
          <p className="mt-1 text-xs text-muted-foreground">
            {isAzure
              ? 'Enter your model deployment names below.'
              : 'Choose the models you want to use. You can also add a custom model ID.'}
          </p>
        </div>
        <div className="flex flex-1 flex-col gap-3 overflow-y-auto">
          {isAzure && (
            <div className="flex gap-2 rounded-md border border-blue-200 bg-blue-50 px-3 py-2 text-sm text-blue-800 dark:border-blue-700 dark:bg-blue-900/30 dark:text-blue-200">
              <Info className="mt-0.5 size-4 shrink-0" />
              <div>
                <span className="font-medium">Tip: </span>
                Enter the exact deployment name from Azure AI Foundry (e.g., gpt-4.1, gpt-5.2-chat).
                Each model will be validated before being added.
              </div>
            </div>
          )}

          <div className="flex items-end gap-2">
            <div className="flex flex-1 flex-col gap-1.5">
              <Label htmlFor="manual-model">
                {isAzure ? 'Model deployment name' : 'Custom model ID (optional)'}
              </Label>
              <Input
                id="manual-model"
                value={manualModelName}
                onChange={e => {
                  setManualModelName(e.target.value);
                  setModelValidationError(null);
                }}
                placeholder={isAzure ? 'e.g., gpt-4.1' : 'e.g., gpt-4o-2024-11-20'}
                onKeyDown={e => {
                  if (e.key === 'Enter') void handleAddManualModel();
                }}
                disabled={validatingModel}
              />
            </div>
            <Button
              onClick={() => void handleAddManualModel()}
              disabled={!manualModelName.trim() || validatingModel}
            >
              {validatingModel ? (
                <Loader2 className="size-4 animate-spin" />
              ) : (
                <Plus className="size-4" />
              )}
              Add
            </Button>
          </div>

          {modelValidationError && (
            <div className="rounded-md border border-red-300 bg-red-50 px-3 py-2 text-sm text-red-800 dark:border-red-700 dark:bg-red-900/30 dark:text-red-200">
              {modelValidationError}
            </div>
          )}

          {manualModels.length > 0 && (
            <div className="flex flex-col gap-1.5">
              <p className="text-xs text-muted-foreground">
                {isAzure ? 'Added models:' : 'Selected models:'}
              </p>
              {manualModels.map(m => (
                <div
                  key={m.id}
                  className="flex items-center justify-between rounded-md bg-secondary px-3 py-1.5"
                >
                  <span className="text-sm">{m.id}</span>
                  <button
                    onClick={() => setManualModels(prev => prev.filter(x => x.id !== m.id))}
                    className="text-muted-foreground hover:text-destructive"
                    aria-label={`Remove ${m.id}`}
                  >
                    <Trash2 className="size-4" />
                  </button>
                </div>
              ))}
            </div>
          )}
        </div>
        <div className="flex shrink-0 justify-end gap-2 py-3">
          <Button variant="secondary" onClick={() => setStep('auth')}>
            Back
          </Button>
          <Button onClick={handleModelsFinish} disabled={manualModels.length === 0}>
            Finish
          </Button>
        </div>
      </div>
    );
  }

  // ─── Step: Done ───
  return (
    <div className="flex h-screen flex-col items-center justify-center gap-3 bg-background px-4 py-3 text-center text-foreground">
      <CheckCircle className="size-12 text-green-500" />
      <h2 className="text-sm font-semibold">You&apos;re all set!</h2>
      <p className="text-sm text-muted-foreground">
        Connected — found {discoveredModels.length} model
        {discoveredModels.length !== 1 ? 's' : ''}
      </p>
      <div className="flex flex-col items-center gap-0.5">
        {discoveredModels.map(m => (
          <span key={m.id} className="text-xs text-muted-foreground">
            {m.name}
          </span>
        ))}
      </div>
      <Button onClick={onComplete} className="mt-2">
        Start Chatting
      </Button>
    </div>
  );
};
