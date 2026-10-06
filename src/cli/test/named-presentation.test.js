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
  return execFileSync(process.execPath, [cli, ...args, '--port', String(port)], {
    cwd, encoding: 'utf8', timeout: 30000
  }).trim();
}

test('named presentation has the requested PowerPoint name and saved path', async () => {
  const directory = fs.mkdtempSync(path.join(os.tmpdir(), 'netoffice-named-'));
  let before = new Set();
  try {
    run(process.cwd(), 'powerpoint', 'launch');
    before = new Set(JSON.parse(run(directory, 'presentation', 'list')).map(item => item.id));
    const created = JSON.parse(run(directory, 'presentation', 'new', '--path', path.join(directory, 'Résumé 研究')));
    assert.equal(created.name, 'Résumé 研究.pptx');
    assert.equal(created.slideCount, 0);
    const target = JSON.parse(run(directory, 'presentation', 'list')).find(item => !before.has(item.id));
    assert.ok(target, 'new document target must appear');
    assert.equal(target.title, created.name);
    assert.equal(path.normalize(target.url), path.join(directory, 'Résumé 研究.pptx'));
    assert.ok(fs.existsSync(target.url));
    const state = JSON.parse(run(directory, 'presentation', 'state', '--target', target.id));
    assert.equal(state.name, created.name);
    assert.equal(state.saved, true);
    assert.equal(state.slideCount, 0);
    const added = JSON.parse(run(directory, 'slide', 'add', '--target', target.id, '--layout', '12'));
    const shape = JSON.parse(run(directory, 'shape', 'add', '--target', target.id,
      '--slide-id', String(added.slideId), '--type', '1', '--left', '10', '--top', '10',
      '--width', '100', '--height', '50'));
    run(directory, 'shape', 'text', '--target', target.id, '--slide-id', String(added.slideId),
      '--shape-id', String(shape.shapeId), '--text', 'Updated separately');
    const detail = JSON.parse(run(directory, 'slide', 'show', '--target', target.id,
      '--slide-id', String(added.slideId)));
    assert.ok(detail.slide.shapes.some(item =>
      item.shapeId === shape.shapeId && item.text === 'Updated separately'));

    const duplicate = spawnSync(process.execPath, [cli, 'presentation', 'new',
      '--path', path.join(directory, 'Résumé 研究'), '--port', String(port)], {
      cwd: directory, encoding: 'utf8', timeout: 30000
    });
    assert.equal(duplicate.status, 1);
    assert.match(duplicate.stderr, /already exists/i);
    const after = JSON.parse(run(directory, 'presentation', 'list'));
    assert.equal(after.filter(item => !before.has(item.id)).length, 1);
    const response = await fetch(`http://127.0.0.1:${port}/json/new`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ name: 'Created via HTTP', directory })
    });
    assert.equal(response.status, 200);
    const httpCreated = await response.json();
    assert.equal(httpCreated.name, 'Created via HTTP.pptx');
    assert.equal(path.normalize(httpCreated.url), path.join(directory, 'Created via HTTP.pptx'));
    assert.ok(fs.existsSync(httpCreated.url));
    assert.equal(httpCreated.slideCount, 0);
    const conflict = await fetch(`http://127.0.0.1:${port}/json/new`, {
      method: 'POST', headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ name: 'Created via HTTP', directory })
    });
    assert.equal(conflict.status, 409);
    assert.equal((await conflict.json()).error.code, -32005);
  } finally {
    const targets = await fetch(`http://127.0.0.1:${port}/json/list`)
      .then(response => response.json()).catch(() => []);
    for (const target of targets) {
      if (!before.has(target.id) && target.url &&
          path.normalize(target.url).startsWith(directory + path.sep)) {
        await fetch(`http://127.0.0.1:${port}/json/close/${target.id}?force`, {
          method: 'PUT'
        }).catch(() => {});
      }
    }
    fs.rmSync(directory, { recursive: true, force: true, maxRetries: 10, retryDelay: 250 });
  }
});
