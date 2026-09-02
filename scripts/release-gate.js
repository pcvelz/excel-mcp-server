#!/usr/bin/env node
// Release gate: build output must boot as a real MCP server, report the version
// being released, and actually write a workbook. Run before anything is
// committed, tagged, pushed or published.
//
// Three releases shipped reporting the wrong version because the built server
// was only ever exercised after publishing. A plain shell pipe cannot do this
// check: feeding every message at once closes stdin, and the server exits
// before it handles the tool call, so the pipe reports a false failure. The
// pipe has to stay open, which is why this is a script.

const { spawn } = require('child_process');
const fs = require('fs');
const os = require('os');
const path = require('path');

const expected = require(path.join(__dirname, '..', 'package.json')).version;
const launcher = path.join(__dirname, '..', 'dist', 'launcher.js');
const workdir = fs.mkdtempSync(path.join(os.tmpdir(), 'release-gate-'));
const workbook = path.join(workdir, 'gate.xlsx');

const child = spawn(process.execPath, [launcher], { stdio: ['pipe', 'pipe', 'inherit'] });

const cleanup = () => {
  child.kill();
  fs.rmSync(workdir, { recursive: true, force: true });
};

const fail = (message) => {
  cleanup();
  console.error(`GATE FAILED: ${message}`);
  process.exit(1);
};

const timer = setTimeout(() => fail('the server did not finish the exchange within 60s'), 60000);

const send = (message) => child.stdin.write(`${JSON.stringify(message)}\n`);

const pending = new Map();
let buffer = '';

child.stdout.on('data', (chunk) => {
  buffer += chunk.toString();
  let newline;
  while ((newline = buffer.indexOf('\n')) >= 0) {
    const line = buffer.slice(0, newline);
    buffer = buffer.slice(newline + 1);
    if (!line.trim()) continue;
    let message;
    try {
      message = JSON.parse(line);
    } catch (error) {
      return fail(`unparseable line from the server: ${line.slice(0, 200)}`);
    }
    const handler = pending.get(message.id);
    if (handler) {
      pending.delete(message.id);
      handler(message);
    }
  }
});

child.on('error', (error) => fail(`could not start the server: ${error.message}`));

pending.set(1, (message) => {
  const reported = message.result && message.result.serverInfo && message.result.serverInfo.version;
  if (reported !== expected) {
    return fail(`the server reports ${reported}, but this release is ${expected}`);
  }
  send({ jsonrpc: '2.0', method: 'notifications/initialized' });
  send({
    jsonrpc: '2.0',
    id: 2,
    method: 'tools/call',
    params: {
      name: 'excel_write_to_sheet',
      arguments: {
        fileAbsolutePath: workbook,
        sheetName: 'Gate',
        newSheet: true,
        range: 'A1:B1',
        values: [['a', 'b']],
      },
    },
  });
});

pending.set(2, (message) => {
  const result = message.result || {};
  if (result.isError) {
    const text = (result.content || []).map((c) => c.text || '').join('');
    return fail(`excel_write_to_sheet returned an error: ${text.slice(0, 300)}`);
  }
  if (!fs.existsSync(workbook)) {
    return fail('excel_write_to_sheet reported success but wrote no workbook');
  }
  if (fs.statSync(workbook).size === 0) {
    return fail('the workbook it wrote is empty');
  }
  clearTimeout(timer);
  cleanup();
  console.log(`GATE PASSED: ${expected} builds, boots, reports its own version, and writes a workbook`);
  process.exit(0);
});

send({
  jsonrpc: '2.0',
  id: 1,
  method: 'initialize',
  params: {
    protocolVersion: '2024-11-05',
    capabilities: {},
    clientInfo: { name: 'release-gate', version: expected },
  },
});
