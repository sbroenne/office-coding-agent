#!/usr/bin/env node
/**
 * PowerPoint Agent Quality Test
 *
 * Connects to the Copilot proxy with mock PowerPoint tools,
 * sends a presentation creation prompt, and evaluates agent behavior
 * (tool ordering, layout variety, verification loop, code quality).
 *
 * Usage: node tests-pptx/test-agent.mjs
 * Requires: npm run dev on https://localhost:3000
 */

import WS from 'ws';

const SERVER_URL = 'wss://localhost:3000/api/copilot';
const TIMEOUT_MS = 120_000;

// ─── Mock Presentation State ───────────────────────────────────────────
let slides = [{ index: 0, text: '(empty slide)', shapes: [] }];

const mockHandlers = {
  get_selected_slides: () =>
    `Selected: Slide 1 (use slideIndex=0 in tools). Text: "(empty slide)"`,
  get_presentation_overview: () => {
    const lines = [`Presentation: ${slides.length} slide(s)`];
    for (const s of slides) lines.push(`  Slide ${s.index + 1}: ${s.text}`);
    return lines.join('\n');
  },
  get_presentation_content: (args) => {
    const s = slides[args?.slideIndex ?? 0];
    return s ? `Slide ${s.index + 1}: ${s.text}` : 'Slide not found.';
  },
  get_slide_image: (args) => {
    const idx = args?.slideIndex ?? 0;
    const s = slides[idx];
    if (!s) return 'Slide not found.';
    return [
      `[Visual inspection of Slide ${idx + 1}]`,
      s.shapes.length === 0
        ? '  (empty — no content)'
        : s.shapes.map((sh, i) => `  Shape ${i}: ${sh.type} "${sh.text}"`).join('\n'),
      `Layout: ${s.shapes.length < 3 ? 'Sparse' : 'Good density'}`,
      'No overlapping elements detected. Margins adequate.',
    ].join('\n');
  },
  get_slide_shapes: (args) => {
    const s = slides[args?.slideIndex ?? 0];
    if (!s || s.shapes.length === 0) return 'No shapes.';
    return s.shapes.map((sh, i) => `Shape ${i}: ${sh.type} "${sh.text}"`).join('\n');
  },
  get_slide_layouts: () =>
    ['Title Slide', 'Title and Content', 'Section Header', 'Two Content', 'Comparison', 'Title Only', 'Blank']
      .map((l, i) => `Layout ${i}: ${l}`).join('\n'),
  add_slide_from_code: (args) => {
    const code = args?.code || '';
    const replaceIdx = args?.replaceSlideIndex;
    const textCount = (code.match(/\.addText\(/g) || []).length;
    const tableCount = (code.match(/\.addTable\(/g) || []).length;
    const shapeCount = (code.match(/\.addShape\(/g) || []).length;
    const total = textCount + tableCount + shapeCount;
    const shapes = [];
    for (let i = 0; i < textCount; i++) shapes.push({ type: 'text', text: '(text)' });
    for (let i = 0; i < tableCount; i++) shapes.push({ type: 'table', text: '(table)' });
    for (let i = 0; i < shapeCount; i++) shapes.push({ type: 'shape', text: '' });
    const slideData = { index: replaceIdx != null ? replaceIdx : slides.length, text: `(generated: ${total} elements)`, shapes };
    if (replaceIdx != null && replaceIdx < slides.length) {
      slides[replaceIdx] = slideData;
      return `Replaced slide ${replaceIdx}. ${total} elements. Use get_slide_image to verify.`;
    }
    slides.push(slideData);
    return `Added slide ${slides.length - 1}. ${total} elements. Use get_slide_image to verify.`;
  },
  // Simple pass-through mocks
  set_presentation_content: () => 'Text box added.',
  clear_slide: () => 'Cleared.', update_slide_shape: () => 'Updated.',
  set_shape_text: () => 'Text set.', update_shape_style: () => 'Style updated.',
  move_resize_shape: () => 'Moved.', delete_shape: () => 'Deleted.',
  delete_slide: () => 'Deleted.', move_slide: () => 'Moved.',
  set_slide_background: () => 'Background set.', add_geometric_shape: () => 'Shape added.',
  add_line: () => 'Line added.', get_slide_notes: () => '(no notes)',
  set_slide_notes: () => 'Notes set.', duplicate_slide: () => 'Duplicated.',
  apply_slide_layout: () => 'Layout applied.', get_selected_shapes: () => '(no shapes selected)',
};

// ─── Tool schemas ──────────────────────────────────────────────────────
const toolDefs = [
  { name: 'get_selected_slides', description: 'Get currently selected slides', parameters: { type: 'object', properties: {}, required: [] } },
  { name: 'get_presentation_overview', description: 'Get overview of all slides', parameters: { type: 'object', properties: {}, required: [] } },
  { name: 'get_presentation_content', description: 'Get text content of a slide', parameters: { type: 'object', properties: { slideIndex: { type: 'number' } }, required: [] } },
  { name: 'get_slide_image', description: 'Get visual snapshot of a slide for verification', parameters: { type: 'object', properties: { slideIndex: { type: 'number' } }, required: [] } },
  { name: 'get_slide_shapes', description: 'List shapes on a slide', parameters: { type: 'object', properties: { slideIndex: { type: 'number' } }, required: [] } },
  { name: 'get_slide_layouts', description: 'List available slide layouts', parameters: { type: 'object', properties: {}, required: [] } },
  { name: 'add_slide_from_code', description: 'Create rich slide with PptxGenJS code', parameters: { type: 'object', properties: { code: { type: 'string' }, replaceSlideIndex: { type: 'number' } }, required: ['code'] } },
  { name: 'set_presentation_content', description: 'Add a text box', parameters: { type: 'object', properties: { slideIndex: { type: 'number' }, text: { type: 'string' } }, required: ['text'] } },
  { name: 'clear_slide', description: 'Clear slide', parameters: { type: 'object', properties: { slideIndex: { type: 'number' } }, required: [] } },
  { name: 'update_slide_shape', description: 'Update shape text', parameters: { type: 'object', properties: { slideIndex: { type: 'number' }, shapeIndex: { type: 'number' }, text: { type: 'string' } }, required: ['shapeIndex', 'text'] } },
  { name: 'set_shape_text', description: 'Set shape text', parameters: { type: 'object', properties: { slideIndex: { type: 'number' }, shapeIndex: { type: 'number' }, text: { type: 'string' } }, required: ['shapeIndex', 'text'] } },
  { name: 'update_shape_style', description: 'Update shape style', parameters: { type: 'object', properties: { slideIndex: { type: 'number' }, shapeIndex: { type: 'number' } }, required: ['shapeIndex'] } },
  { name: 'move_resize_shape', description: 'Move/resize shape', parameters: { type: 'object', properties: { slideIndex: { type: 'number' }, shapeIndex: { type: 'number' } }, required: ['shapeIndex'] } },
  { name: 'delete_shape', description: 'Delete shape', parameters: { type: 'object', properties: { slideIndex: { type: 'number' }, shapeIndex: { type: 'number' } }, required: ['shapeIndex'] } },
  { name: 'delete_slide', description: 'Delete slide', parameters: { type: 'object', properties: { slideIndex: { type: 'number' } }, required: ['slideIndex'] } },
  { name: 'move_slide', description: 'Reorder slides', parameters: { type: 'object', properties: { fromIndex: { type: 'number' }, toIndex: { type: 'number' } }, required: ['fromIndex', 'toIndex'] } },
  { name: 'set_slide_background', description: 'Set background color', parameters: { type: 'object', properties: { slideIndex: { type: 'number' }, color: { type: 'string' } }, required: ['slideIndex', 'color'] } },
  { name: 'add_geometric_shape', description: 'Add shape', parameters: { type: 'object', properties: { slideIndex: { type: 'number' } }, required: ['slideIndex'] } },
  { name: 'add_line', description: 'Add line', parameters: { type: 'object', properties: { slideIndex: { type: 'number' } }, required: ['slideIndex'] } },
  { name: 'get_slide_notes', description: 'Get notes', parameters: { type: 'object', properties: { slideIndex: { type: 'number' } }, required: [] } },
  { name: 'set_slide_notes', description: 'Set notes', parameters: { type: 'object', properties: { slideIndex: { type: 'number' }, notes: { type: 'string' } }, required: ['notes'] } },
  { name: 'duplicate_slide', description: 'Duplicate slide', parameters: { type: 'object', properties: { sourceSlideIndex: { type: 'number' } }, required: ['sourceSlideIndex'] } },
  { name: 'apply_slide_layout', description: 'Apply layout', parameters: { type: 'object', properties: { slideIndex: { type: 'number' }, layoutIndex: { type: 'number' } }, required: ['slideIndex', 'layoutIndex'] } },
  { name: 'get_selected_shapes', description: 'Get selected shapes', parameters: { type: 'object', properties: {}, required: [] } },
];

// ─── LSP / JSON-RPC ────────────────────────────────────────────────────
function lspFrame(obj) {
  const json = JSON.stringify(obj);
  return `Content-Length: ${Buffer.byteLength(json)}\r\n\r\n${json}`;
}

// ─── Main ──────────────────────────────────────────────────────────────
async function runTest() {
  console.log('╔══════════════════════════════════════════════════════════╗');
  console.log('║  PowerPoint Agent Quality Test                          ║');
  console.log('╚══════════════════════════════════════════════════════════╝\n');

  slides = [{ index: 0, text: '(empty slide)', shapes: [] }];

  const ws = new WS(SERVER_URL, { rejectUnauthorized: false });
  await new Promise((resolve, reject) => {
    ws.on('open', resolve);
    ws.on('error', () => reject(new Error('Cannot connect — is npm run dev running?')));
    setTimeout(() => reject(new Error('Connection timeout')), 5000);
  });
  console.log('✓ Connected\n');

  let msgId = 0;
  const pending = new Map();
  const toolCalls = [];
  let assistantText = '';
  let done = false;
  let buffer = '';

  function sendRpc(method, params) {
    const id = ++msgId;
    return new Promise((resolve, reject) => {
      pending.set(id, { resolve, reject });
      ws.send(lspFrame({ jsonrpc: '2.0', id, method, params }));
      setTimeout(() => {
        if (pending.has(id)) { pending.delete(id); reject(new Error(`Timeout: ${method}`)); }
      }, TIMEOUT_MS);
    });
  }

  function sendResponse(id, result) {
    ws.send(lspFrame({ jsonrpc: '2.0', id, result }));
  }

  function processBuffer() {
    while (true) {
      const headerEnd = buffer.indexOf('\r\n\r\n');
      if (headerEnd === -1) return;
      const header = buffer.slice(0, headerEnd);
      const lenMatch = /Content-Length:\s*(\d+)/i.exec(header);
      if (!lenMatch) { buffer = ''; return; }
      const len = parseInt(lenMatch[1], 10);
      const bodyStart = headerEnd + 4;
      if (buffer.length < bodyStart + len) return; // wait for more data
      const body = buffer.slice(bodyStart, bodyStart + len);
      buffer = buffer.slice(bodyStart + len);

      let msg;
      try { msg = JSON.parse(body); } catch { continue; }

      // RPC response
      if (msg.id != null && pending.has(msg.id)) {
        const p = pending.get(msg.id);
        pending.delete(msg.id);
        if (msg.error) p.reject(new Error(msg.error.message));
        else p.resolve(msg.result);
        continue;
      }

      // Tool call
      if (msg.method === 'tool.call') {
        const name = msg.params?.toolName;
        const args = msg.params?.arguments || {};
        const handler = mockHandlers[name];
        const result = handler ? handler(args) : `Unknown tool: ${name}`;
        toolCalls.push({ name, args });
        const shortArgs = JSON.stringify(args);
        console.log(`  🔧 ${name}(${shortArgs.length > 100 ? shortArgs.slice(0, 100) + '…' : shortArgs})`);
        sendResponse(msg.id, { result });
        continue;
      }

      // Events
      if (msg.method === 'session.event') {
        const evt = msg.params?.event;
        if (!evt) continue;
        if (evt.type === 'assistant.message_delta') assistantText += evt.data?.deltaContent || '';
        if (evt.type === 'assistant.message') assistantText += evt.data?.content || '';
        if (evt.type === 'session.idle') done = true;
      }
    }
  }

  ws.on('message', (data) => {
    buffer += Buffer.isBuffer(data) ? data.toString('utf8') : String(data);
    processBuffer();
  });

  // ─── Create session ────────────────────────────────────────────────
  console.log('Creating session with', toolDefs.length, 'tools...');
  const systemContent = `You are a PowerPoint assistant with direct presentation access.

## Rules
1. ALWAYS call get_selected_slides first.
2. Use get_presentation_overview to understand structure.
3. Use add_slide_from_code with PptxGenJS for rich slides.
4. ALWAYS verify with get_slide_image after creating/modifying EACH slide.
5. Iterate: create → verify → fix → re-verify.

## Layout Variety — MANDATORY
Never repeat the same layout for >2 consecutive slides. Use: title, bullets, two-column, three-column, stat callouts, quote, table.

## PptxGenJS Rules
- bold: true for headings. { bullet: true } for lists.
- Colors: 6-digit hex without #: "4472C4"
- Margins: x≥0.5, y≥0.5, x+w≤9.5, y+h≤7.0
- Fonts: title≥28pt, body≥16pt`;

  const sessionResult = await sendRpc('session.create', {
    sessionId: `pptx-test-${Date.now()}`,
    systemMessage: { mode: 'append', content: systemContent },
    tools: toolDefs,
  });
  console.log('✓ Session:', sessionResult.sessionId, '\n');

  // ─── Send prompt ───────────────────────────────────────────────────
  const prompt = `Create a 5-slide presentation about "The Future of AI":
1. Title slide (dark background)
2. What is AI? (bullets)
3. Applications (multi-column layout, 4+ examples)
4. Challenges vs. Opportunities (two-column)
5. Conclusion with quote
Verify each slide with get_slide_image.`;

  console.log('📤 Prompt sent\n─── Tool Calls ────────────────────────────────────────\n');
  await sendRpc('session.send', { sessionId: sessionResult.sessionId, prompt });

  // Wait for idle
  const start = Date.now();
  while (!done && Date.now() - start < TIMEOUT_MS) await new Promise(r => setTimeout(r, 500));

  // ─── Evaluate ──────────────────────────────────────────────────────
  console.log('\n─── Evaluation ────────────────────────────────────────\n');

  const names = toolCalls.map(t => t.name);
  const slideCodes = toolCalls.filter(t => t.name === 'add_slide_from_code').map(t => t.args.code || '');
  const imageChecks = names.filter(t => t === 'get_slide_image').length;
  const slideCreates = names.filter(t => t === 'add_slide_from_code').length;
  const replaces = toolCalls.filter(t => t.name === 'add_slide_from_code' && t.args.replaceSlideIndex != null).length;

  const checks = [
    ['get_selected_slides called first', names[0] === 'get_selected_slides'],
    ['get_presentation_overview called early', names.slice(0, 3).includes('get_presentation_overview')],
    ['Used add_slide_from_code', names.includes('add_slide_from_code')],
    [`Created ≥5 slides (got ${slideCreates})`, slideCreates >= 5],
    [`Verified with get_slide_image (${imageChecks}× for ${slideCreates} slides)`, imageChecks >= slideCreates],
    [`Iterated (${replaces} replace calls)`, replaces > 0],
    ['Uses bold: true', slideCodes.some(c => /bold:\s*true/.test(c))],
    ['Uses { bullet: true }', slideCodes.some(c => /bullet:\s*true/.test(c))],
    ['No unicode bullets', !slideCodes.some(c => /[•‣◦▪▸]/.test(c))],
    ['Hex colors without #', slideCodes.some(c => /color:\s*["'][0-9A-Fa-f]{6}["']/.test(c))],
    ['Title font ≥ 28pt', slideCodes.some(c => /fontSize:\s*(2[8-9]|[3-9]\d|\d{3})/.test(c))],
  ];

  let passed = 0;
  for (const [label, ok] of checks) {
    console.log(`${ok ? '✅' : '❌'} ${label}`);
    if (ok) passed++;
  }

  // Summary
  console.log(`\n─── Summary ───────────────────────────────────────────\n`);
  const counts = {};
  for (const t of toolCalls) counts[t.name] = (counts[t.name] || 0) + 1;
  for (const [name, count] of Object.entries(counts).sort((a, b) => b[1] - a[1]))
    console.log(`  ${String(count).padStart(3)}× ${name}`);
  console.log(`\n  Score: ${passed}/${checks.length} | Tools: ${toolCalls.length} | Slides: ${slides.length}`);

  // Show code snippets
  if (slideCodes.length > 0) {
    console.log(`\n─── Code Sample (Slide 1) ─────────────────────────────\n`);
    console.log(slideCodes[0].slice(0, 500));
    if (slideCodes[0].length > 500) console.log('  ...(truncated)');
  }

  try { await sendRpc('session.destroy', { sessionId: sessionResult.sessionId }); } catch {}
  ws.close();
  console.log('\n✓ Done');
  process.exit(passed >= checks.length * 0.6 ? 0 : 1);
}

runTest().catch(err => { console.error('Fatal:', err.message); process.exit(1); });
