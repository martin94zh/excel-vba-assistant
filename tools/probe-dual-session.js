/**
 * 双进程会话探针：验证 mcp-excel.exe 2.0.8 下，
 * 进程 A 用 MCP 打开的工作簿/会话，进程 B（另一个独立 server 进程）能否复用。
 *
 * 用法：node tools/probe-dual-session.js [exePath] [testFile]
 */
const { spawn } = require('child_process');
const fs = require('fs');
const path = require('path');

const exePath = process.argv[2] || path.join(__dirname, '..', 'build', 'mcp-2.0.8', 'server', 'mcp-excel.exe');
const testFile = process.argv[3] || path.join(__dirname, '..', 'build', 'mcp-2.0.8', 'probe-test.xlsm');

function makeClient(name) {
  const child = spawn(exePath, [], { stdio: ['pipe', 'pipe', 'pipe'] });
  const pending = new Map();
  let buffer = '';
  let nextId = 1;
  child.stdout.on('data', (d) => {
    buffer += d.toString();
    const lines = buffer.split('\n');
    buffer = lines.pop() || '';
    for (const line of lines) {
      if (!line.trim()) continue;
      try {
        const msg = JSON.parse(line);
        if (msg.id && pending.has(msg.id)) {
          const { resolve } = pending.get(msg.id);
          pending.delete(msg.id);
          resolve(msg);
        }
      } catch {}
    }
  });
  child.stderr.on('data', (d) => process.stderr.write(`[${name} stderr] ${d}`));
  const call = (method, params) => new Promise((resolve, reject) => {
    const id = nextId++;
    pending.set(id, { resolve });
    child.stdin.write(JSON.stringify({ jsonrpc: '2.0', id, method, params }) + '\n');
    setTimeout(() => {
      if (pending.has(id)) { pending.delete(id); reject(new Error(`${name} ${method} timeout (60s)`)); }
    }, 60000);
  });
  const client = {
    child,
    async init() {
      await call('initialize', {
        protocolVersion: '2024-11-05',
        capabilities: {},
        clientInfo: { name: 'probe-dual-session', version: '1.0.0' },
      });
      child.stdin.write(JSON.stringify({ jsonrpc: '2.0', method: 'notifications/initialized' }) + '\n');
    },
    async tool(toolName, args) {
      const r = await call('tools/call', { name: toolName, arguments: args });
      return r.result;
    },
  };
  void name;
  return client;
}

function textOf(result) {
  try {
    return (result?.content || []).map((c) => c.text || '').join('\n') || JSON.stringify(result);
  } catch {
    return JSON.stringify(result);
  }
}

async function createTestFile(A) {
  for (const action of ['create', 'create-empty']) {
    const r = await A.tool('file', { action, path: testFile });
    const t = textOf(r);
    console.log(`file(${action}):`, t.slice(0, 300));
    if (fs.existsSync(testFile)) return true;
  }
  return fs.existsSync(testFile);
}

async function main() {
  console.log(`exe = ${exePath}`);
  console.log(`testFile = ${testFile}\n`);

  const A = makeClient('A');
  await A.init();

  console.log('=== 进程 A（模拟插件的 exe）===');
  if (!fs.existsSync(testFile)) {
    await createTestFile(A);
  }
  const opened = await A.tool('file', { action: 'open', path: testFile, show: true });
  const openText = textOf(opened);
  console.log('A file(open, show:true):', openText.slice(0, 500));
  const sidMatch = openText.match(/"session_?[iI]d"\s*:\s*"([^"]+)"/);
  const sidA = sidMatch ? sidMatch[1] : null;
  console.log('A 取得 session_id =', sidA, '\n');

  const B = makeClient('B');
  await B.init();

  console.log('=== 进程 B（模拟 Trae 另起的 exe）===');
  const listed = await B.tool('file', { action: 'list' });
  console.log('B file(list):', textOf(listed).slice(0, 500), '\n');

  const bOpen = await B.tool('file', { action: 'open', path: testFile });
  const bOpenText = textOf(bOpen);
  console.log('B file(open) 同一文件:', bOpenText.slice(0, 500), '\n');

  if (sidA) {
    const bUse = await B.tool('worksheet', { action: 'list', session_id: sidA });
    console.log('B 携带 A 的 sessionId 调 worksheet(list):', textOf(bUse).slice(0, 400), '\n');
  }

  console.log('=== 清理 ===');
  if (sidA) {
    await A.tool('file', { action: 'close', session_id: sidA, save: false }).catch(() => {});
    console.log('A 已请求关闭会话');
  }
  A.child.kill();
  B.child.kill();
  try { if (fs.existsSync(testFile)) fs.unlinkSync(testFile); } catch {}
  process.exit(0);
}

main().catch((e) => {
  console.error('PROBE FAILED:', e && e.message ? e.message : e);
  process.exit(1);
});
