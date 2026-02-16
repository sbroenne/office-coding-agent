import React, { useState, useCallback } from 'react';
import { Plus, Trash2, Pencil, Check, X, Loader2 } from 'lucide-react';
import { Button } from '@/components/ui/button';
import { Input } from '@/components/ui/input';
import { Label } from '@/components/ui/label';
import { Badge } from '@/components/ui/badge';
import { Tooltip, TooltipTrigger, TooltipContent } from '@/components/ui/tooltip';
import { useSettingsStore } from '@/stores';
import { validateModelDeployment, inferProvider } from '@/services/ai';
import type { ModelInfo } from '@/types';

interface ModelManagerProps {
  endpointId: string;
}

export const ModelManager: React.FC<ModelManagerProps> = ({ endpointId }) => {
  const { getModelsForEndpoint, getActiveEndpoint, addModel, updateModel, removeModel } =
    useSettingsStore();

  const endpoint = getActiveEndpoint();
  const models = getModelsForEndpoint(endpointId);

  // Add model state
  const [adding, setAdding] = useState(false);
  const [newModelName, setNewModelName] = useState('');
  const [validating, setValidating] = useState(false);
  const [validationError, setValidationError] = useState<string | null>(null);

  // Edit model state
  const [editingId, setEditingId] = useState<string | null>(null);
  const [editDisplayName, setEditDisplayName] = useState('');

  const handleStartAdd = useCallback(() => {
    setAdding(true);
    setNewModelName('');
    setValidationError(null);
  }, []);

  const handleCancelAdd = useCallback(() => {
    setAdding(false);
    setNewModelName('');
    setValidationError(null);
  }, []);

  const handleAddModel = useCallback(async () => {
    const name = newModelName.trim();
    if (!name || !endpoint) return;

    if (models.some(m => m.id === name)) {
      setValidationError(`"${name}" is already added.`);
      return;
    }

    setValidating(true);
    setValidationError(null);

    const epConfig = { ...endpoint, id: endpoint.id };
    const ok = await validateModelDeployment(epConfig, name);
    setValidating(false);

    if (!ok) {
      setValidationError(`Could not reach "${name}". Check the deployment name is correct.`);
      return;
    }

    const model: ModelInfo = {
      id: name,
      name,
      ownedBy: 'user',
      provider: inferProvider(name),
    };

    addModel(endpointId, model);
    setAdding(false);
    setNewModelName('');
    setValidationError(null);
  }, [newModelName, endpoint, models, addModel, endpointId]);

  const handleStartEdit = useCallback((model: ModelInfo) => {
    setEditingId(model.id);
    setEditDisplayName(model.name);
  }, []);

  const handleSaveEdit = useCallback(() => {
    if (!editingId || !editDisplayName.trim()) return;
    updateModel(endpointId, editingId, { name: editDisplayName.trim() });
    setEditingId(null);
    setEditDisplayName('');
  }, [editingId, editDisplayName, updateModel, endpointId]);

  const handleCancelEdit = useCallback(() => {
    setEditingId(null);
    setEditDisplayName('');
  }, []);

  const handleRemove = useCallback(
    (modelId: string) => {
      removeModel(endpointId, modelId);
    },
    [removeModel, endpointId]
  );

  return (
    <div className="flex flex-col gap-2">
      {/* Header */}
      <div className="flex items-center justify-between">
        <h3 className="text-sm font-semibold">Models</h3>
        {!adding && (
          <Tooltip>
            <TooltipTrigger asChild>
              <Button variant="ghost" size="icon" className="size-7" onClick={handleStartAdd}>
                <Plus className="size-4" />
              </Button>
            </TooltipTrigger>
            <TooltipContent>Add model</TooltipContent>
          </Tooltip>
        )}
      </div>

      {/* Empty state */}
      {models.length === 0 && !adding && (
        <div className="py-3 text-center text-xs text-muted-foreground">
          <p>No models configured.</p>
          <button onClick={handleStartAdd} className="mt-1 text-xs text-primary hover:underline">
            Add a model
          </button>
        </div>
      )}

      {/* Model list */}
      {models.map(model =>
        editingId === model.id ? (
          <div
            key={model.id}
            className="flex items-center gap-1.5 rounded-md bg-secondary px-2 py-1"
          >
            <Input
              value={editDisplayName}
              onChange={e => setEditDisplayName(e.target.value)}
              onKeyDown={e => {
                if (e.key === 'Enter') handleSaveEdit();
                if (e.key === 'Escape') handleCancelEdit();
              }}
              className="h-7 flex-1 text-xs"
            />
            <Button
              variant="ghost"
              size="icon"
              className="size-6"
              onClick={handleSaveEdit}
              disabled={!editDisplayName.trim()}
            >
              <Check className="size-3.5" />
            </Button>
            <Button variant="ghost" size="icon" className="size-6" onClick={handleCancelEdit}>
              <X className="size-3.5" />
            </Button>
          </div>
        ) : (
          <div
            key={model.id}
            className="flex items-center justify-between rounded-md bg-secondary px-2 py-1"
          >
            <div className="flex min-w-0 items-center gap-1.5">
              <span className="truncate text-sm">{model.name}</span>
              {model.id !== model.name && (
                <span className="text-xs text-muted-foreground">({model.id})</span>
              )}
              <Badge variant="outline" className="text-[10px] px-1.5 py-0">
                {model.provider}
              </Badge>
            </div>
            <div className="flex shrink-0 gap-0.5">
              <Tooltip>
                <TooltipTrigger asChild>
                  <Button
                    variant="ghost"
                    size="icon"
                    className="size-6"
                    onClick={() => handleStartEdit(model)}
                  >
                    <Pencil className="size-3.5" />
                  </Button>
                </TooltipTrigger>
                <TooltipContent>Rename</TooltipContent>
              </Tooltip>
              <Tooltip>
                <TooltipTrigger asChild>
                  <Button
                    variant="ghost"
                    size="icon"
                    className="size-6"
                    onClick={() => handleRemove(model.id)}
                  >
                    <Trash2 className="size-3.5" />
                  </Button>
                </TooltipTrigger>
                <TooltipContent>Remove</TooltipContent>
              </Tooltip>
            </div>
          </div>
        )
      )}

      {/* Add model form */}
      {adding && (
        <div className="flex flex-col gap-2">
          <div className="flex items-end gap-1.5">
            <div className="flex flex-1 flex-col gap-1">
              <Label htmlFor="new-model" className="text-xs">
                Deployment name
              </Label>
              <Input
                id="new-model"
                value={newModelName}
                onChange={e => {
                  setNewModelName(e.target.value);
                  setValidationError(null);
                }}
                placeholder="e.g., gpt-5.2-chat"
                onKeyDown={e => {
                  if (e.key === 'Enter') void handleAddModel();
                  if (e.key === 'Escape') handleCancelAdd();
                }}
                disabled={validating}
                className="h-8 text-xs"
              />
            </div>
            <Button
              size="sm"
              onClick={() => void handleAddModel()}
              disabled={!newModelName.trim() || validating}
            >
              {validating ? (
                <Loader2 className="size-4 animate-spin" />
              ) : (
                <Plus className="size-4" />
              )}
              {validating ? 'Validating' : 'Add'}
            </Button>
            <Button variant="secondary" size="sm" onClick={handleCancelAdd}>
              Cancel
            </Button>
          </div>
          {validationError && (
            <div className="rounded-md border border-red-300 bg-red-50 px-2 py-1.5 text-xs text-red-800 dark:border-red-700 dark:bg-red-900/30 dark:text-red-200">
              {validationError}
            </div>
          )}
          <p className="text-xs text-muted-foreground">
            Enter the exact deployment name from Azure AI Foundry. The model will be validated
            before being added.
          </p>
        </div>
      )}
    </div>
  );
};
