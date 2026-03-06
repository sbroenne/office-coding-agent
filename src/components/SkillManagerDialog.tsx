import React from 'react';
import { Codicon } from '@/components/Codicon';
import { Button } from '@/components/ui/button';
import { getBundledSkills } from '@/services/skills';
import { downloadSkill } from '@/services/extensions/zipExportService';

export const SkillManagerPanel: React.FC = () => {
  const bundledSkills = getBundledSkills();

  return (
    <div className="space-y-3 p-3">
      {/* Bundled skills */}
      <div className="space-y-1">
        <p className="text-[11px] font-medium text-muted-foreground">Bundled (read-only)</p>
        {bundledSkills.length === 0 ? (
          <p className="text-xs text-muted-foreground">No bundled skills.</p>
        ) : (
          bundledSkills.map(skill => (
            <div
              key={`bundled-skill-${skill.metadata.name}`}
              className="flex items-center justify-between rounded-md border border-border px-2 py-1.5"
            >
              <div className="min-w-0">
                <p className="truncate text-sm font-medium">{skill.metadata.name}</p>
                <p className="truncate text-xs text-muted-foreground">
                  {skill.metadata.description}
                </p>
              </div>
              <Button
                variant="ghost"
                size="icon"
                className="size-7 shrink-0"
                onClick={() => downloadSkill(skill)}
                aria-label={`Download ${skill.metadata.name} as template`}
                title="Download as template"
              >
                <Codicon name="cloud-download" className="text-sm" />
              </Button>
            </div>
          ))
        )}
      </div>

      {/* Plugin skills */}
      <div className="space-y-1">
        <p className="text-[11px] font-medium text-muted-foreground">From plugins</p>
        <p className="text-xs text-muted-foreground">
          Custom skills are installed via the Copilot CLI Plugin Hub.
        </p>
      </div>
    </div>
  );
};
