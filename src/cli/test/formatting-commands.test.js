'use strict';

const assert = require('node:assert/strict');
const { execFileSync } = require('node:child_process');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const { test } = require('node:test');
const { Connection } = require('../lib/connection');

const port = Number(process.env.NETOFFICE_PORT || 50051);
const baseUrl = `http://127.0.0.1:${port}`;
const deadline = () => Date.now() + 30000;

async function http(method, pathname, body) {
  const response = await fetch(`${baseUrl}${pathname}`, body === undefined ? { method } : {
    method, headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(body)
  });
  return { status: response.status, body: await response.json() };
}

test('slide, master, and application commands resolve their scope over both transports', async () => {
  execFileSync(process.execPath, [path.join(__dirname, '..', 'bin', 'netoffice.js'),
    'powerpoint', 'launch', '--port', String(port), '--timeout', '30000'], { stdio: 'inherit' });
  const client = await Connection.connect(`ws://127.0.0.1:${port}/devtools/application`, deadline());
  const directory = fs.mkdtempSync(path.join(os.tmpdir(), 'netoffice-format-'));
  let targetId;
  try {
    const request = (method, params) => client.request(method, params, deadline());
    targetId = (await request('PowerPoint.newPresentation', { name: 'Formatting', directory })).id;
    const { slideId } = await request('PowerPoint.addSlide', { targetId, layout: 12 });
    const { shapeId } = await request('PowerPoint.createShape',
      { targetId, slideId, shapeType: 1, left: 10, top: 10, width: 200, height: 60 });
    await request('PowerPoint.setShapeText', { targetId, slideId, shapeId, text: 'Quarterly results' });

    // Character formatting reads back as separate runs.
    await request('PowerPoint.setShapeFont', { targetId, slideId, shapeId, start: 1, length: 9, bold: true, color: '#1F3864' });
    const state = await http('GET', `/json/${targetId}/slides/${slideId}/shapes/${shapeId}`);
    assert.equal(state.status, 200);
    const runs = state.body.paragraphs[0].runs;
    assert.deepEqual(runs.map(run => [run.text, run.font.bold, run.font.color]),
      [['Quarterly', true, '#1F3864'], [' results', false, runs[1].font.color]]);

    // Bounds change keeps the shape ID; the setter echoes accepted members.
    const moved = await http('PUT', `/json/${targetId}/slides/${slideId}/shapes/${shapeId}/bounds`, { left: 40 });
    assert.equal(moved.status, 200);
    assert.equal(moved.body.shapeId, shapeId);
    assert.equal(moved.body.left, 40);
    assert.equal(Math.round(moved.body.bounds.left), 40);

    // A setter that sets nothing, and invalid values, are rejected before mutation.
    assert.equal((await http('PUT', `/json/${targetId}/slides/${slideId}/shapes/${shapeId}/fill`, {})).status, 400);
    await assert.rejects(request('PowerPoint.setShapeFill', { targetId, slideId, shapeId, color: 'navy' }),
      error => error.code === -32602);

    // Master routes address the slide master; slide-only commands refuse it.
    const master = await http('POST', `/json/${targetId}/master/textboxes`, { left: 10, top: 500, width: 200, height: 20 });
    assert.equal(master.status, 200);
    assert.equal(master.body.master, true);
    const masterShapes = (await request('PowerPoint.getMasterState', { targetId })).shapes;
    assert.ok(masterShapes.some(shape => shape.shapeId === master.body.shapeId));
    assert.equal((await request('PowerPoint.getSlideState', { targetId, slideId })).slide.shapes.length, 1);
    await assert.rejects(request('PowerPoint.moveSlide', { targetId, master: true, index: 1 }),
      error => error.code === -32602);
    await assert.rejects(request('PowerPoint.setBackground', { targetId, slideId, master: true, color: '#000000' }),
      error => error.code === -32602);

    // Custom layouts are containers by index and slide sources by index or exact name.
    const customLayouts = (await request('PowerPoint.getLayouts', { targetId })).layouts;
    const layout = customLayouts[1];
    const logo = await http('POST', `/json/${targetId}/layouts/${layout.index}/shapes`,
      { shapeType: 1, left: 900, top: 10, width: 40, height: 20 });
    assert.equal(logo.status, 200);
    assert.equal(logo.body.customLayout, layout.index);
    const filled = await request('PowerPoint.setShapeFill',
      { targetId, customLayout: layout.index, shapeId: logo.body.shapeId, color: '#C9A227' });
    assert.equal(filled.customLayout, layout.index);
    const layoutState = await http('GET', `/json/${targetId}/layouts/${layout.index}`);
    assert.equal(layoutState.body.name, layout.name);
    assert.ok(layoutState.body.shapes.some(shape => shape.shapeId === logo.body.shapeId));
    const named = await request('PowerPoint.addSlide', { targetId, customLayoutName: layout.name });
    assert.equal(named.customLayout, layout.index);
    await assert.rejects(request('PowerPoint.addSlide', { targetId, customLayoutName: 'No such layout' }),
      error => error.code === -32004);
    await assert.rejects(request('PowerPoint.setShapeFill',
      { targetId, slideId, customLayout: layout.index, shapeId, color: '#000000' }), error => error.code === -32602);

    // Grouped shapes are reported recursively.
    const second = await request('PowerPoint.createShape',
      { targetId, slideId, shapeType: 1, left: 10, top: 100, width: 50, height: 50 });
    const group = await request('PowerPoint.groupShapes', { targetId, slideId, shapeIds: [shapeId, second.shapeId] });
    const detail = await request('PowerPoint.getSlideState', { targetId, slideId });
    const grouped = detail.slide.shapes.find(shape => shape.shapeId === group.shapeId);
    assert.deepEqual(grouped.groupItems.map(item => item.shapeId).sort(), [shapeId, second.shapeId].sort());

    // Freeforms are built from straight segments through the points; invalid input creates nothing.
    const before = (await request('PowerPoint.getSlideState', { targetId, slideId })).slide.shapes.length;
    const triangle = [{ x: 300, y: 100 }, { x: 400, y: 250 }, { x: 250, y: 220 }, { x: 300, y: 100 }];
    const freeform = await http('POST', `/json/${targetId}/slides/${slideId}/freeforms`, { points: triangle });
    assert.equal(freeform.status, 200);
    assert.equal(freeform.body.slideId, slideId);
    const drawn = (await request('PowerPoint.getShapeState', { targetId, slideId, shapeId: freeform.body.shapeId }));
    assert.equal(drawn.shapeType, 5);
    assert.deepEqual([drawn.bounds.left, drawn.bounds.top, drawn.bounds.width, drawn.bounds.height].map(Math.round),
      [250, 100, 150, 150]);
    for (const points of [undefined, [], [{ x: 1, y: 1 }], [{ x: 1, y: 1 }, { x: 1, y: 1 }],
      [{ x: 1, y: 1 }, { x: 'a', y: 2 }], [{ x: 1, y: 1 }, { x: 2 }], [{ x: 1, y: 1 }, [2, 3]],
      [{ x: 1, y: 1 }, { x: 20000, y: 2 }], 'x'])
      await assert.rejects(request('PowerPoint.addFreeform', { targetId, slideId, points }), error => error.code === -32602);
    assert.equal((await request('PowerPoint.getSlideState', { targetId, slideId })).slide.shapes.length, before + 1);

    // Application-scoped reads need no target.
    const layouts = await http('GET', '/json/smartart-layouts');
    assert.equal(layouts.status, 200);
    assert.ok(layouts.body.layouts.length > 0);

    const saved = await http('POST', `/json/${targetId}/presentation/save`, {});
    assert.equal(saved.status, 200);
    assert.equal(saved.body.saved, true);
  } finally {
    if (targetId) await http('PUT', `/json/close/${targetId}?force`).catch(() => {});
    client.close();
    fs.rmSync(directory, { recursive: true, force: true, maxRetries: 10, retryDelay: 250 });
  }
});
