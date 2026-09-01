#!/usr/bin/env node
// Refuses to let a build whose stamped version disagrees with package.json
// reach npm. This is not hypothetical: npm's ignore-scripts is enabled on some
// machines, which silently skips prepublishOnly, so `npm publish` happily ships
// whatever dist/ was left over from an earlier build. Every release up to and
// including 0.16.1 went out reporting the PREVIOUS release's version that way.
// Run through `npm run release`, which builds first and then calls this.

const { spawn } = require('child_process');
const path = require('path');

const expected = require(path.join(__dirname, '..', 'package.json')).version;
const launcher = path.join(__dirname, '..', 'dist', 'launcher.js');

const child = spawn(process.execPath, [launcher], { stdio: ['pipe', 'pipe', 'inherit'] });

const handshake = {
  jsonrpc: '2.0',
  id: 1,
  method: 'initialize',
  params: {
    protocolVersion: '2024-11-05',
    capabilities: {},
    clientInfo: { name: 'verify-dist-version', version: expected },
  },
};

let buffer = '';
let settled = false;

const fail = (message) => {
  if (settled) return;
  settled = true;
  child.kill();
  console.error(`verify-dist-version: ${message}`);
  process.exit(1);
};

const timer = setTimeout(() => fail('the server did not answer initialize within 30s'), 30000);

child.stdout.on('data', (chunk) => {
  buffer += chunk.toString();
  const newline = buffer.indexOf('\n');
  if (newline < 0 || settled) return;

  clearTimeout(timer);
  let reported;
  try {
    reported = JSON.parse(buffer.slice(0, newline)).result.serverInfo.version;
  } catch (error) {
    return fail(`could not read serverInfo.version: ${error.message}`);
  }
  if (reported !== expected) {
    return fail(
      `dist/ reports ${reported} but package.json says ${expected}. ` +
        'Run "npm run build" before publishing.'
    );
  }
  settled = true;
  child.kill();
  console.log(`verify-dist-version: dist/ reports ${reported}, matching package.json`);
});

child.on('error', (error) => fail(`could not start the server: ${error.message}`));
child.stdin.write(`${JSON.stringify(handshake)}\n`);
