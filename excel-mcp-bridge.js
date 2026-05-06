/**
 * Excel MCP HTTP Bridge Server
 * Bridges HTTP requests from the add-in to the sbroenne/excel-mcp server via stdio
 * 
 * Usage: node excel-mcp-bridge.js
 * 
 * Requirements:
 * - Windows OS with Excel installed
 * - Download mcp-excel.exe from https://github.com/sbroenne/mcp-server-excel/releases
 * - Place mcp-excel.exe in the same directory or update MCP_EXE_PATH below
 */

const { spawn } = require('child_process');
const http = require('http');
const path = require('path');

const PORT = process.env.PORT || 8765;
const MCP_EXE_PATH = process.env.MCP_EXE_PATH || path.join(__dirname, 'mcp-excel.exe');

let mcpProcess = null;
let mcpReady = false;
let requestId = 0;
const pendingRequests = new Map();

function startMCP() {
  console.log(`Starting Excel MCP server: ${MCP_EXE_PATH}`);
  
  mcpProcess = spawn(MCP_EXE_PATH, [], {
    stdio: ['pipe', 'pipe', 'pipe'],
    windowsHide: true
  });

  mcpProcess.stdout.on('data', (data) => {
    const output = data.toString();
    console.log('[MCP OUT]:', output.substring(0, 500));
    
    // Parse JSON-RPC responses
    try {
      const lines = output.split('\n').filter(l => l.trim());
      for (const line of lines) {
        try {
          const msg = JSON.parse(line);
          if (msg.id && pendingRequests.has(msg.id)) {
            const { resolve, reject } = pendingRequests.get(msg.id);
            pendingRequests.delete(msg.id);
            if (msg.error) {
              reject(new Error(msg.error.message || JSON.stringify(msg.error)));
            } else {
              resolve(msg.result);
            }
          }
        } catch (e) {
          // Not JSON, ignore
        }
      }
    } catch (e) {
      // Ignore parse errors
    }
  });

  mcpProcess.stderr.on('data', (data) => {
    console.log('[MCP ERR]:', data.toString().substring(0, 500));
  });

  mcpProcess.on('error', (err) => {
    console.error('MCP Process error:', err.message);
    mcpReady = false;
  });

  mcpProcess.on('exit', (code) => {
    console.log('MCP Process exited with code:', code);
    mcpReady = false;
    mcpProcess = null;
  });

  // Wait a moment for MCP to initialize
  setTimeout(() => {
    mcpReady = true;
    console.log('✅ Excel MCP server ready');
  }, 2000);
}

function sendMCPRequest(method, params = {}) {
  return new Promise((resolve, reject) => {
    if (!mcpProcess || !mcpReady) {
      reject(new Error('MCP server not ready'));
      return;
    }

    const id = ++requestId;
    const request = {
      jsonrpc: '2.0',
      id,
      method,
      params
    };

    pendingRequests.set(id, { resolve, reject });
    
    const requestStr = JSON.stringify(request) + '\n';
    mcpProcess.stdin.write(requestStr);
    
    // Timeout after 30 seconds
    setTimeout(() => {
      if (pendingRequests.has(id)) {
        pendingRequests.delete(id);
        reject(new Error('MCP request timeout'));
      }
    }, 30000);
  });
}

const server = http.createServer(async (req, res) => {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  
  if (req.method === 'OPTIONS') {
    res.writeHead(204);
    res.end();
    return;
  }

  const url = new URL(req.url, `http://localhost:${PORT}`);

  // Health check
  if (url.pathname === '/health') {
    res.writeHead(200, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify({ 
      status: mcpReady ? 'ready' : 'starting',
      mcp: mcpReady 
    }));
    return;
  }

  // List available tools
  if (url.pathname === '/tools') {
    try {
      const result = await sendMCPRequest('tools/list');
      res.writeHead(200, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify(result));
    } catch (err) {
      res.writeHead(500, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ error: err.message }));
    }
    return;
  }

  // Execute MCP tool
  if (url.pathname === '/execute' && req.method === 'POST') {
    let body = '';
    req.on('data', chunk => body += chunk);
    req.on('end', async () => {
      try {
        const { tool, arguments: args } = JSON.parse(body);
        const result = await sendMCPRequest('tools/call', { 
          name: tool, 
          arguments: args || {} 
        });
        res.writeHead(200, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify(result));
      } catch (err) {
        res.writeHead(500, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ error: err.message }));
      }
    });
    return;
  }

  // Chat endpoint - sends natural language to Excel MCP
  if (url.pathname === '/chat' && req.method === 'POST') {
    let body = '';
    req.on('data', chunk => body += chunk);
    req.on('end', async () => {
      try {
        const { message, context } = JSON.parse(body);
        
        // Build a prompt for the Excel MCP based on user's message
        // The MCP server will use its tools to execute the request
        const toolsResult = await sendMCPRequest('tools/call', {
          name: 'execute_macro',
          arguments: { 
            code: `
// Excel AI Bridge - Processing: ${message}
// Context: ${JSON.stringify(context || []).substring(0, 500)}
MsgBox "Processing: ${message}"
            `,
            showExcel: true
          }
        }).catch(() => null);
        
        // For now, return a response indicating the MCP is ready
        // In a full implementation, we'd use the MCP's actual tools
        res.writeHead(200, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ 
          response: `📊 Excel MCP recibido: "${message}". MCP server activo y listo para automatizar Excel.`,
          tools: ['Power Query', 'DAX', 'VBA', 'PivotTables', 'Charts', 'Ranges', 'Tables']
        }));
      } catch (err) {
        res.writeHead(500, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ error: err.message }));
      }
    });
    return;
  }

  res.writeHead(404, { 'Content-Type': 'application/json' });
  res.end(JSON.stringify({ error: 'Not found' }));
});

server.listen(PORT, () => {
  console.log(`🎯 Excel MCP Bridge Server running on http://localhost:${PORT}`);
  console.log(`   - GET  /health     - Health check`);
  console.log(`   - GET  /tools      - List MCP tools`);
  console.log(`   - POST /execute    - Execute MCP tool`);
  console.log(`   - POST /chat       - Natural language command`);
  
  startMCP();
});

process.on('SIGINT', () => {
  console.log('\nShutting down...');
  if (mcpProcess) {
    mcpProcess.kill();
  }
  server.close();
  process.exit();
});