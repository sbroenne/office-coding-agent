/**
 * Quick smoke-test: connect to the local proxy WebSocket and call models.list.
 * Run with: node scripts/test-copilot-ws.mjs
 */
import { WebSocket } from 'ws';
import https from 'node:https';

const agent = new https.Agent({ rejectUnauthorized: false });
const ws = new WebSocket('wss://localhost:3000/api/copilot', { agent });

function lspFrame(obj) {
  const body = JSON.stringify(obj);
  return `Content-Length: ${Buffer.byteLength(body)}\r\n\r\n${body}`;
}

const timeout = setTimeout(() => {
  console.error('TIMEOUT — no response in 20 s');
  ws.terminate();
  process.exit(1);
}, 20000);

ws.on('open', () => {
  console.log('WS connected — requesting models.list...');
  ws.send(lspFrame({ jsonrpc: '2.0', id: 1, method: 'models.list', params: {} }));
});

let buf = '';
ws.on('message', d => {
  buf += d.toString();
  const hEnd = buf.indexOf('\r\n\r\n');
  if (hEnd === -1) return;
  const m = buf.match(/Content-Length:\s*(\d+)/i);
  if (!m) return;
  const len = parseInt(m[1]);
  if (buf.length < hEnd + 4 + len) return;
  const msg = JSON.parse(buf.slice(hEnd + 4, hEnd + 4 + len));
  clearTimeout(timeout);
  ws.close();
  if (msg.error) {
    console.error('ERROR from proxy:', JSON.stringify(msg.error));
    process.exit(1);
  }
  if (msg.result?.models) {
    console.log('SUCCESS — available models:');
    msg.result.models.forEach(m => console.log(' ', m.id));
    process.exit(0);
  }
  console.log('Unexpected response:', JSON.stringify(msg));
  process.exit(1);
});

ws.on('error', e => {
  clearTimeout(timeout);
  console.error('WS error:', e.message);
  process.exit(1);
});
