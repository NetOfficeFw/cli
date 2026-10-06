'use strict';

const assert = require('node:assert/strict');
const { execFileSync, spawnSync } = require('node:child_process');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const { test } = require('node:test');

const cli = path.join(__dirname, '..', 'bin', 'netoffice.js');
const port = Number(process.env.NETOFFICE_PORT || 50051);

function run(cwd, ...args) {
  return JSON.parse(execFileSync(process.execPath, [cli, ...args, '--port', String(port)], {
    cwd, encoding: 'utf8', timeout: 30000
  }));
}

test('presentation close scopes force and presentation open reloads the saved file', async () => {
  const directory = fs.mkdtempSync(path.join(os.tmpdir(), 'netoffice-close-'));
  const created = [];
  try {
    execFileSync(process.execPath, [cli, 'powerpoint', 'launch', '--port', String(port)], {
      cwd: process.cwd(), encoding: 'utf8', timeout: 30000
    });
    const first = run(directory, 'presentation', 'new', '--path', 'Discard me');
    created.push(first.id);
    const other = run(directory, 'presentation', 'new', '--path', 'Keep me');
    created.push(other.id);
    const original = fs.readFileSync(first.url);

    run(directory, 'slide', 'add', '--target', first.id);
    assert.equal(run(directory, 'presentation', 'state', '--target', first.id).saved, false);
    const refused = spawnSync(process.execPath, [cli, 'presentation', 'close',
      '--target', first.id, '--port', String(port)], {
      cwd: directory, encoding: 'utf8', timeout: 30000
    });
    assert.equal(refused.status, 1);
    assert.match(refused.stderr, /unsaved/i);
    assert.equal(run(directory, 'presentation', 'state', '--target', first.id).id, first.id);
    assert.equal(run(directory, 'presentation', 'state', '--target', other.id).id, other.id);

    assert.deepEqual(run(directory, 'presentation', 'close', '--target', first.id, '--force'), {});
    assert.equal(run(directory, 'presentation', 'list').some(item => item.id === first.id), false);
    assert.equal(run(directory, 'presentation', 'state', '--target', other.id).id, other.id);
    assert.ok(fs.readFileSync(first.url).equals(original), 'forced close must not save edits');
    const missing = spawnSync(process.execPath, [cli, 'presentation', 'open',
      '--path', 'Missing.pptx', '--port', String(port)], {
      cwd: directory, encoding: 'utf8', timeout: 30000
    });
    assert.equal(missing.status, 1);
    assert.match(missing.stderr, /Open/i);
    assert.equal(run(directory, 'presentation', 'state', '--target', other.id).id, other.id);

    const reopened = run(directory, 'presentation', 'open', '--path', 'Discard me.pptx');
    created.push(reopened.id);
    assert.equal(reopened.title, first.name);
    assert.equal(path.normalize(reopened.url), path.normalize(first.url));
    assert.equal(run(directory, 'presentation', 'state', '--target', reopened.id).slideCount, 0);
    assert.equal(run(directory, 'presentation', 'state', '--target', other.id).id, other.id);
    assert.deepEqual(run(directory, 'presentation', 'close', '--target', reopened.id), {});

    assert.deepEqual(run(directory, 'presentation', 'close', '--target', other.id), {});
    assert.equal(run(directory, 'presentation', 'list').some(item => item.id === other.id), false);
  } finally {
    for (const id of created) {
      await fetch(`http://127.0.0.1:${port}/json/close/${id}?force`, { method: 'PUT' })
        .catch(() => {});
    }
    fs.rmSync(directory, { recursive: true, force: true, maxRetries: 10, retryDelay: 250 });
  }
});
