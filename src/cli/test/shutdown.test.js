'use strict';

const assert = require('node:assert/strict');
const { execFileSync, spawnSync } = require('node:child_process');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const { test } = require('node:test');
const { Connection } = require('../lib/connection');

const cli = path.join(__dirname, '..', 'bin', 'netoffice.js');
const port = Number(process.env.NETOFFICE_PORT || 50051);
const baseUrl = `http://127.0.0.1:${port}`;
const args = ['--port', String(port), '--timeout', '30000'];

function run(...command) {
  return execFileSync(process.execPath, [cli, ...command, ...args], {
    encoding: 'utf8', timeout: 35000
  }).trim();
}

test('shutdown preserves unsaved work and exits with saved presentations open', async t => {
  let existing;
  try {
    existing = await fetch(`${baseUrl}/json/version`, { signal: AbortSignal.timeout(750) });
  } catch { /* No add-in is serving this port. */ }
  if (existing?.ok) return t.skip('PowerPoint is already running; this test requires an isolated session');

  let documentId;
  let stopped = false;
  const directory = fs.mkdtempSync(path.join(os.tmpdir(), 'netoffice-shutdown-'));
  try {
    run('powerpoint', 'launch');
    const client = await Connection.connect(`ws://127.0.0.1:${port}/devtools/application`, Date.now() + 30000);
    let processId;
    try {
      processId = (await client.request('PowerPoint.getStatus', {}, Date.now() + 30000)).processId;
    } finally {
      client.close();
    }
    const created = await fetch(`${baseUrl}/json/new`, { method: 'PUT' });
    assert.equal(created.status, 200);
    documentId = (await created.json()).id;
    const rejected = spawnSync(process.execPath, [cli, 'powerpoint', 'shutdown', ...args], {
      encoding: 'utf8', timeout: 35000
    });
    assert.equal(rejected.status, 1);
    assert.match(rejected.stderr, /unsaved changes/i);
    process.kill(processId, 0); // The failed command did not shut down PowerPoint.
    const closed = await fetch(`${baseUrl}/json/close/${documentId}?force`, { method: 'PUT' });
    assert.equal(closed.status, 200);
    documentId = undefined;
    const saved = JSON.parse(execFileSync(process.execPath,
      [cli, 'presentation', 'new', '--path', 'Saved deck', ...args],
      { cwd: directory, encoding: 'utf8', timeout: 35000 }));
    documentId = saved.id;
    assert.equal(saved.slideCount, 0);
    assert.ok(fs.existsSync(path.join(directory, 'Saved deck.pptx')));
    const result = JSON.parse(run('powerpoint', 'shutdown'));
    assert.deepEqual(result, { processId, stopped: true });
    stopped = true;
    documentId = undefined;
    assert.throws(() => process.kill(processId, 0), { code: 'ESRCH' });
    assert.ok(fs.existsSync(path.join(directory, 'Saved deck.pptx')));
  } finally {
    if (documentId) await fetch(`${baseUrl}/json/close/${documentId}?force`, { method: 'PUT' }).catch(() => {});
    if (!stopped) spawnSync(process.execPath, [cli, 'powerpoint', 'shutdown', ...args], {
      encoding: 'utf8', timeout: 35000
    });
    fs.rmSync(directory, { recursive: true, force: true, maxRetries: 10, retryDelay: 250 });
  }
});

test('force discards changes in open presentations before quitting', async t => {
  let existing;
  try {
    existing = await fetch(`${baseUrl}/json/version`, { signal: AbortSignal.timeout(750) });
  } catch { /* No add-in is serving this port. */ }
  if (existing?.ok) return t.skip('PowerPoint is already running; this test requires an isolated session');

  const directory = fs.mkdtempSync(path.join(os.tmpdir(), 'netoffice-force-'));
  const documentIds = [];
  let stopped = false;
  try {
    run('powerpoint', 'launch');
    const client = await Connection.connect(`ws://127.0.0.1:${port}/devtools/application`, Date.now() + 30000);
    let processId;
    try {
      processId = (await client.request('PowerPoint.getStatus', {}, Date.now() + 30000)).processId;
      const invalid = await client.request('PowerPoint.prepareShutdown',
        { force: 'yes' }, Date.now() + 30000).then(() => null, error => error);
      assert.equal(invalid.code, -32602);
    } finally {
      client.close();
    }
    const saved = JSON.parse(execFileSync(process.execPath,
      [cli, 'presentation', 'new', '--path', 'Force deck', ...args],
      { cwd: directory, encoding: 'utf8', timeout: 35000 }));
    documentIds.push(saved.id);
    const original = fs.readFileSync(saved.url);
    const slide = JSON.parse(run('slide', 'add', '--target', saved.id, '--layout', '12'));
    assert.ok(slide.slideId > 0);
    const blank = await fetch(`${baseUrl}/json/new`, { method: 'PUT' });
    assert.equal(blank.status, 200);
    documentIds.push((await blank.json()).id);
    const refused = spawnSync(process.execPath, [cli, 'powerpoint', 'shutdown', ...args], {
      encoding: 'utf8', timeout: 35000
    });
    assert.equal(refused.status, 1);
    assert.match(refused.stderr, /unsaved changes/i);
    const forced = JSON.parse(run('powerpoint', 'shutdown', '--force'));
    stopped = true;
    assert.deepEqual(forced, { processId, stopped: true });
    assert.throws(() => process.kill(processId, 0), { code: 'ESRCH' });
    assert.deepEqual(fs.readFileSync(saved.url), original, 'unsaved slide must not be written to disk');
  } finally {
    if (!stopped) {
      for (const id of documentIds)
        await fetch(`${baseUrl}/json/close/${id}?force`, { method: 'PUT' }).catch(() => {});
      spawnSync(process.execPath, [cli, 'powerpoint', 'shutdown', ...args], {
        encoding: 'utf8', timeout: 35000
      });
    }
    fs.rmSync(directory, { recursive: true, force: true, maxRetries: 10, retryDelay: 250 });
  }
});
