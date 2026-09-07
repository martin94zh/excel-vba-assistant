const { spawn } = require('child_process');
const path = require('path');

const exePath = process.argv[2] || path.join(__dirname, '..', 'dist', 'mcp-excel.exe');
const toolName = process.argv[3] || 'file';

const child = spawn(exePath, [], { stdio: ['pipe', 'pipe', 'pipe'] });

let buffer = '';
let initialized = false;

child.stdout.on('data', (data) => {
  buffer += data.toString();
  const lines = buffer.split('\n');
  buffer = lines.pop() || '';
  for (const line of lines) {
    if (!line.trim()) continue;
    try {
      const msg = JSON.parse(line);
      if (msg.id === 1 && msg.result?.protocolVersion) {
        initialized = true;
        child.stdin.write(JSON.stringify({ jsonrpc: '2.0', id: 2, method: 'tools/list', params: {} }) + '\n');
      } else if (msg.id === 2 && msg.result?.tools) {
        const tool = msg.result.tools.find((t) => t.name === toolName);
        if (tool) {
          console.log(JSON.stringify(tool, null, 2));
        } else {
          console.log(`Tool ${toolName} not found. Available:`, msg.result.tools.map((t) => t.name).join(', '));
        }
        child.kill();
        process.exit(0);
      }
    } catch (e) {
      // ignore
    }
  }
});

child.stderr.on('data', (data) => {
  // ignore banner
});

setTimeout(() => {
  child.stdin.write(JSON.stringify({
    jsonrpc: '2.0', id: 1, method: 'initialize',
    params: { protocolVersion: '2024-11-05', capabilities: {}, clientInfo: { name: 'test', version: '1.0.0' } }
  }) + '\n');
}, 500);

setTimeout(() => {
  child.kill();
  process.exit(1);
}, 5000);
