#!/usr/bin/env node
import fs from 'node:fs';
import path from 'node:path';

function parseArgs(argv) {
  const result = {};
  for (let index = 0; index < argv.length; index += 1) {
    const arg = argv[index];
    if (!arg.startsWith('--')) continue;
    const key = arg.slice(2);
    const value = argv[index + 1];
    if (!value || value.startsWith('--')) {
      result[key] = true;
      continue;
    }
    result[key] = value;
    index += 1;
  }
  return result;
}

const args = parseArgs(process.argv.slice(2));

const sourcePath = path.resolve(args.source ?? 'manifest.xml');
const outputPath = path.resolve(args.output ?? 'manifests/manifest.staging.xml');
const baseUrlInput = args.baseUrl ?? 'https://sbroenne.github.io/office-coding-agent';
const addinId = args.id ?? '4f0b6a89-6cba-4930-b102-47c7dc4d9c2d';
const displayName = args.displayName ?? 'Office Coding Agent (Staging)';
const commandsGroupLabel = args.groupLabel ?? 'AI Assistant STG';
const chatButtonLabel = args.chatLabel ?? 'AI Chat STG';

const baseUrl = baseUrlInput.replace(/\/$/, '');

if (!fs.existsSync(sourcePath)) {
  throw new Error(`Source manifest not found: ${sourcePath}`);
}

let xml = fs.readFileSync(sourcePath, 'utf8');

xml = xml.replace(/<Id>[\s\S]*?<\/Id>/, `<Id>${addinId}</Id>`);

xml = xml.replace(
  /<DisplayName\s+DefaultValue="[^"]*"\s*\/>/,
  `<DisplayName DefaultValue="${displayName}" />`
);

xml = xml.replace(
  /<bt:String\s+id="GetStarted.Title"\s+DefaultValue="[^"]*"\s*\/>/,
  `<bt:String id="GetStarted.Title" DefaultValue="${displayName}" />`
);

xml = xml.replace(
  /<bt:String\s+id="CommandsGroup.Label"\s+DefaultValue="[^"]*"\s*\/>/,
  `<bt:String id="CommandsGroup.Label" DefaultValue="${commandsGroupLabel}" />`
);

xml = xml.replace(
  /<bt:String\s+id="TaskpaneButton.Label"\s+DefaultValue="[^"]*"\s*\/>/,
  `<bt:String id="TaskpaneButton.Label" DefaultValue="${chatButtonLabel}" />`
);

xml = xml.replaceAll('https://localhost:3000', baseUrl);

fs.mkdirSync(path.dirname(outputPath), { recursive: true });
fs.writeFileSync(outputPath, xml, 'utf8');

console.log(`Generated staging manifest: ${outputPath}`);
console.log(`Base URL: ${baseUrl}`);
console.log(`Add-in ID: ${addinId}`);
console.log(`Display name: ${displayName}`);
console.log(`Group label: ${commandsGroupLabel}`);
console.log(`Chat label: ${chatButtonLabel}`);
