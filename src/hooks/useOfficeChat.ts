import { useMemo } from 'react';
import { useChat, type UseChatHelpers } from '@ai-sdk/react';
import {
  ToolLoopAgent,
  DirectChatTransport,
  stepCountIs,
  type ToolSet,
  type LanguageModel,
} from 'ai';

import { getToolsForHost, getGeneralTools } from '@/tools';
import { buildSkillContext } from '@/services/skills';
import { resolveActiveMcpServers } from '@/services/mcp';
import { resolveActiveAgent } from '@/services/agents';
import { useSettingsStore } from '@/stores';
import { buildSystemPrompt } from '@/services/ai/systemPrompt';
import { normalizeChatErrorMessage } from '@/services/ai/chatErrorMessage';
import { useMcpTools } from './useMcpTools';
import type { OfficeHostApp } from '@/services/office/host';

export type { UseChatHelpers };

export function useOfficeChat(model: LanguageModel | null, host: OfficeHostApp, tools?: ToolSet) {
  const activeSkillNames = useSettingsStore(s => s.activeSkillNames);
  const importedSkills = useSettingsStore(s => s.importedSkills);
  const activeAgentId = useSettingsStore(s => s.activeAgentId);
  const importedMcpServers = useSettingsStore(s => s.importedMcpServers);
  const activeMcpServerNames = useSettingsStore(s => s.activeMcpServerNames);

  const activeMcpServers = useMemo(
    () => resolveActiveMcpServers(importedMcpServers, activeMcpServerNames),
    [importedMcpServers, activeMcpServerNames]
  );
  const mcpTools = useMcpTools(activeMcpServers);

  const agent = useMemo(() => {
    if (!model) return null;

    const resolvedAgent = resolveActiveAgent(activeAgentId, host);
    const agentInstructions = resolvedAgent?.instructions ?? '';
    const skillContext = buildSkillContext(activeSkillNames ?? undefined);
    const instructions = `${buildSystemPrompt(host)}\n\n${agentInstructions}${skillContext}`;

    const hostTools = tools ?? getToolsForHost(host);
    return new ToolLoopAgent({
      model,
      instructions,
      tools: { ...hostTools, ...getGeneralTools(model, hostTools), ...mcpTools },
      stopWhen: stepCountIs(10),
      maxRetries: 4,
    });
  }, [model, host, tools, activeSkillNames, importedSkills, activeAgentId, mcpTools]);

  const transport = useMemo(() => {
    if (!agent) return null;
    return new DirectChatTransport({ agent });
  }, [agent]);

  return useChat({
    transport: transport ?? undefined,
    onError: error => {
      error.message = normalizeChatErrorMessage(error.message);
      console.error('[useOfficeChat] Chat error:', error);
      console.error('[useOfficeChat] Error details:', {
        message: error.message,
        name: error.name,
        stack: error.stack,
      });
    },
  });
}
