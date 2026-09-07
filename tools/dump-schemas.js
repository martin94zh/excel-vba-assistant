const { spawn } = require('child_process');
const fs = require('fs');
const path = require('path');

const exePath = process.argv[2] || path.join(__dirname, '..', 'dist', 'mcp-excel.exe');
const outPath = path.join(__dirname, '..', 'tools', 'schemas.json');
const child = spawn(exePath, [], { stdio: ['pipe', 'pipe', 'pipe'] });

let buffer = '';
let done = false;
child.stdout.on('data', (data) => {
  buffer += data.toString();
  const lines = buffer.split('\n');
  buffer = lines.pop() || '';
  for (const line of lines) {
    if (!line.trim() || done) continue;
    try {
      const msg = JSON.parse(line);
      if (msg.id === 1 && msg.result?.protocolVersion) {
        child.stdin.write(JSON.stringify({ jsonrpc: '2.0', id: 2, method: 'tools/list', params: {} }) + '\n');
      } else if (msg.id === 2 && msg.result?.tools) {
        fs.writeFileSync(outPath, JSON.stringify(msg.result.tools, null, 2), 'utf8');
        console.log(`Dumped ${msg.result.tools.length} tool schemas to ${outPath}`);
        done = true;
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

setTimeout(() => { child.kill(); process.exit(1); }, 10000);
