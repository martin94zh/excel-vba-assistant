const { spawn } = require('child_process');
const path = require('path');

const exePath = process.argv[2] || path.join(__dirname, '..', 'dist', 'mcp-excel.exe');
const child = spawn(exePath, [], { stdio: ['pipe', 'pipe', 'pipe'] });

let buffer = '';
child.stdout.on('data', (data) => {
  buffer += data.toString();
  const lines = buffer.split('\n');
  buffer = lines.pop() || '';
  for (const line of lines) {
    if (!line.trim()) continue;
    try {
      const msg = JSON.parse(line);
      if (msg.id === 1 && msg.result?.protocolVersion) {
        child.stdin.write(JSON.stringify({ jsonrpc: '2.0', id: 2, method: 'tools/list', params: {} }) + '\n');
      } else if (msg.id === 2 && msg.result?.tools) {
        const names = msg.result.tools.map((t) => t.name).sort();
        console.log(names.join('\n'));
        child.kill();
        process.exit(0);
      }
    } catch {}
  }
});

setTimeout(() => {
  child.stdin.write(JSON.stringify({
    jsonrpc: '2.0', id: 1, method: 'initialize',
    params: { protocolVersion: '2024-11-05', capabilities: {}, clientInfo: { name: 'test', version: '1.0.0' } }
  }) + '\n');
}, 500);

setTimeout(() => { child.kill(); process.exit(1); }, 5000);
