import fs from 'node:fs/promises';
import path from 'node:path';
import JSZip from 'jszip';

const outputDir = path.resolve(process.cwd(), 'samples', 'extensions');

async function writeZip(filePath, builder) {
  const zip = new JSZip();
  await builder(zip);
  const content = await zip.generateAsync({ type: 'nodebuffer' });
  await fs.writeFile(filePath, content);
}

async function createSkillsSampleZip() {
  const filePath = path.join(outputDir, 'sample-skills.zip');

  await writeZip(filePath, async zip => {
    zip.file(
      'skills/security-review.md',
      `---
name: Security Review
description: Security checklist and threat-model guidance for coding tasks.
version: 1.0.0
tags:
  - security
  - review
---

When the user asks for a security review:
- Identify attack surfaces
- Call out auth, data protection, and secrets handling risks
- Suggest concrete mitigation steps
`
    );

    zip.file(
      'skills/excel-analytics.md',
      `---
name: Excel Analytics
description: Guidance for analytics workflows and structured output.
version: 1.0.0
tags:
  - excel
  - analytics
---

When solving Excel analytics tasks:
- Explain assumptions briefly
- Prefer clear table outputs
- Summarize findings and next actions
`
    );
  });

  return filePath;
}

async function createAgentsSampleZip() {
  const filePath = path.join(outputDir, 'sample-agents.zip');

  await writeZip(filePath, async zip => {
    zip.file(
      'agents/excel-data-analyst.md',
      `---
name: Excel Data Analyst
description: Focused on spreadsheet analysis and actionable summaries.
version: 1.0.0
hosts: [excel]
defaultForHosts: []
---

You are an Excel-focused data analyst agent.
- Prioritize concise findings and practical next steps.
- When uncertain, state assumptions explicitly.
- Keep recommendations implementation-ready.
`
    );

    zip.file(
      'agents/powerpoint-storyteller.md',
      `---
name: PowerPoint Storyteller
description: Helps structure slide narratives and key messages.
version: 1.0.0
hosts: [powerpoint]
defaultForHosts: []
---

You are a presentation-structure specialist.
- Organize ideas into clear narrative flow.
- Keep language concise and audience-oriented.
- Highlight one key takeaway per section.
`
    );
  });

  return filePath;
}

async function run() {
  await fs.mkdir(outputDir, { recursive: true });

  const [skillsZip, agentsZip] = await Promise.all([
    createSkillsSampleZip(),
    createAgentsSampleZip(),
  ]);

  console.log('Created extension sample ZIP files:');
  console.log(` - ${skillsZip}`);
  console.log(` - ${agentsZip}`);
}

run().catch(error => {
  console.error('Failed to generate extension sample ZIP files.');
  console.error(error instanceof Error ? error.message : String(error));
  process.exit(1);
});
