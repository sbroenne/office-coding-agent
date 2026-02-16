import { useMemo } from 'react';
import { useChat, type UseChatHelpers } from '@ai-sdk/react';
import { ToolLoopAgent, DirectChatTransport, stepCountIs, type ToolSet } from 'ai';
import type { AzureOpenAIProvider } from '@ai-sdk/azure';

import { getToolsForHost } from '@/tools';
import { buildSkillContext } from '@/services/skills';
import { resolveActiveAgent } from '@/services/agents';
import { useSettingsStore } from '@/stores';
import { buildSystemPrompt } from '@/services/ai/systemPrompt';
import type { OfficeHostApp } from '@/services/office/host';

export type { UseChatHelpers };

export function useOfficeChat(
  provider: AzureOpenAIProvider | null,
  modelId: string | null,
  host: OfficeHostApp,
  tools?: ToolSet
) {
  const activeSkillNames = useSettingsStore(s => s.activeSkillNames);
  const activeAgentId = useSettingsStore(s => s.activeAgentId);

  const agent = useMemo(() => {
    if (!provider || !modelId) return null;

    const resolvedAgent = resolveActiveAgent(activeAgentId, host);
    const agentInstructions = resolvedAgent?.instructions ?? '';
    const skillContext = buildSkillContext(activeSkillNames ?? undefined);
    const instructions = `${buildSystemPrompt(host)}\n\n${agentInstructions}${skillContext}`;

    return new ToolLoopAgent({
      model: provider.chat(modelId),
      instructions,
      tools: tools ?? getToolsForHost(host),
      stopWhen: stepCountIs(10),
      maxRetries: 4,
    });
  }, [provider, modelId, host, tools, activeSkillNames, activeAgentId]);

  const transport = useMemo(() => {
    if (!agent) return null;
    return new DirectChatTransport({ agent });
  }, [agent]);

  return useChat({
    transport: transport ?? undefined,
    onError: error => {
      console.error('[useOfficeChat] Chat error:', error);
      console.error('[useOfficeChat] Error details:', {
        message: error.message,
        name: error.name,
        stack: error.stack,
      });
    },
  });
}
