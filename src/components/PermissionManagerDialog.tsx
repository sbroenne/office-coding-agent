import React from 'react';
import { Codicon } from '@/components/Codicon';

interface PermissionManagerPanelProps {
  approveAll: boolean;
  onSetApproveAll: (enabled: boolean) => void | Promise<void>;
  onResetSessionApprovals: () => void | Promise<void>;
}

export const PermissionManagerPanel: React.FC<PermissionManagerPanelProps> = ({
  approveAll,
  onSetApproveAll,
  onResetSessionApprovals,
}) => {
  return (
    <div className="space-y-4 p-3 text-sm">
      <div
        className="flex items-center justify-between rounded-[var(--vscode-cornerRadius-medium)] border p-3"
        style={{ borderColor: 'var(--vscode-widget-border)' }}
      >
        <div>
          <div className="font-medium">Allow all</div>
          <div className="text-xs" style={{ color: 'var(--vscode-descriptionForeground)' }}>
            Use the Copilot CLI session setting to auto-approve permission requests.
          </div>
        </div>
        <button
          type="button"
          onClick={() => void onSetApproveAll(!approveAll)}
          className="inline-flex rounded-[var(--vscode-cornerRadius-medium)] px-2 py-1 text-xs font-medium"
          style={{
            color: approveAll
              ? 'var(--vscode-button-foreground)'
              : 'var(--vscode-button-secondaryForeground)',
            backgroundColor: approveAll
              ? 'var(--vscode-button-background)'
              : 'var(--vscode-button-secondaryBackground)',
          }}
        >
          {approveAll ? 'On' : 'Off'}
        </button>
      </div>

      <div
        className="rounded-[var(--vscode-cornerRadius-medium)] border p-3"
        style={{ borderColor: 'var(--vscode-widget-border)' }}
      >
        <div className="mb-1 font-medium">Session approvals</div>
        <div className="mb-3 text-xs" style={{ color: 'var(--vscode-descriptionForeground)' }}>
          Approvals remembered for this Copilot session are owned by the SDK.
        </div>
        <button
          type="button"
          onClick={() => void onResetSessionApprovals()}
          className="inline-flex items-center gap-1 rounded-[var(--vscode-cornerRadius-medium)] border px-2 py-1 text-xs"
          style={{
            borderColor: 'var(--vscode-widget-border)',
            color: 'var(--vscode-foreground)',
            backgroundColor: 'transparent',
          }}
        >
          <Codicon name="discard" className="text-xs" />
          Reset session approvals
        </button>
      </div>
    </div>
  );
};
