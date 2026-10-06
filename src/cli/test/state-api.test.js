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
const socketUrl = `ws://127.0.0.1:${port}/devtools/application`;
const deadline = () => Date.now() + 30000;

async function getJson(pathname) {
  const response = await fetch(`${baseUrl}${pathname}`);
  const body = await response.json();
  return { status: response.status, body };
}

function controlSlideShow(action, url) {
  const command = action === 'start'
    ? `$ppt = [Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application'); $presentation = @($ppt.Presentations | Where-Object { $_.FullName -eq $env:NETOFFICE_TEST_PATH }) | Select-Object -First 1; if (-not $presentation) { throw 'Test presentation was not found' }; $null = $presentation.SlideShowSettings.Run()`
    : `$ppt = [Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application'); foreach ($window in @($ppt.SlideShowWindows)) { if ($window.Presentation.FullName -eq $env:NETOFFICE_TEST_PATH) { $window.View.Exit() } }`;
  execFileSync('powershell.exe', ['-NoProfile', '-NonInteractive', '-Command', command], {
    env: { ...process.env, NETOFFICE_TEST_PATH: url },
    stdio: 'ignore'
  });
}

function selectFirstShape(url) {
  const command = `$ppt = [Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application'); $presentation = @($ppt.Presentations | Where-Object { $_.FullName -eq $env:NETOFFICE_TEST_PATH }) | Select-Object -First 1; if (-not $presentation) { throw 'Test presentation was not found' }; $window = $presentation.Windows.Item(1); $window.Activate(); $window.View.GotoSlide(1); $presentation.Slides.Item(1).Shapes.Item(1).Select()`;
  execFileSync('powershell.exe', ['-NoProfile', '-NonInteractive', '-Command', command], {
    env: { ...process.env, NETOFFICE_TEST_PATH: url },
    stdio: 'ignore'
  });
}

async function request(client, method, params = {}) {
  return client.request(method, params, deadline());
}

test('PowerPoint presentation state is readable and controllable over HTTP and WebSocket', async t => {
  assert.ok(Number.isInteger(port) && port > 0 && port <= 65535, 'NETOFFICE_PORT must be a valid TCP port');
  execFileSync(process.execPath, [path.join(__dirname, '..', 'bin', 'netoffice.js'),
    'powerpoint', 'launch', '--port', String(port), '--timeout', '30000'], { stdio: 'inherit' });
  const client = await Connection.connect(socketUrl, deadline());
  const createdTargets = [];
  const directory = fs.mkdtempSync(path.join(os.tmpdir(), 'netoffice-state-'));
  try {
    const before = await getJson('/json/list');
    assert.equal(before.status, 200);
    const previousIds = new Set(before.body.map(target => target.id));
    const created = await request(client, 'PowerPoint.newPresentation',
      { name: 'State API test', directory });
    assert.equal(created.slideCount, 0);
    createdTargets.push(created.id);
    const after = await getJson('/json/list');
    assert.equal(after.status, 200);
    const target = after.body.find(candidate => candidate.id === created.id && !previousIds.has(candidate.id));
    assert.ok(target, 'new presentation should appear as a document target');
    const secondCreated = await request(client, 'PowerPoint.newPresentation',
      { name: 'Other deck', directory });
    assert.equal(secondCreated.slideCount, 0);
    createdTargets.push(secondCreated.id);
    const withSecond = await getJson('/json/list');
    const other = withSecond.body.find(candidate => candidate.id === secondCreated.id);
    assert.ok(other, 'second presentation should have an independent document target');
    for (const [presentation, text] of [[target, 'State API test title'], [other, 'Other deck title']]) {
      const added = await request(client, 'PowerPoint.addSlide',
        { targetId: presentation.id, layout: 12 });
      const shape = await request(client, 'PowerPoint.createShape', {
        targetId: presentation.id, slideId: added.slideId,
        shapeType: 1, left: 36, top: 24, width: 648, height: 54
      });
      await request(client, 'PowerPoint.setShapeText', {
        targetId: presentation.id, slideId: added.slideId, shapeId: shape.shapeId, text
      });
    }

    await t.test('reads remain scoped to the requested document while another is active', async () => {
      const first = await request(client, 'PowerPoint.getSlideState',
        { targetId: target.id, slideId: (await request(client, 'PowerPoint.getSlides',
          { targetId: target.id })).slides[0].slideId });
      const second = await request(client, 'PowerPoint.getSlideState',
        { targetId: other.id, slideId: (await request(client, 'PowerPoint.getSlides',
          { targetId: other.id })).slides[0].slideId });
      assert.ok(first.slide.shapes.some(shape => shape.text && shape.text.trim() === 'State API test title'));
      assert.ok(second.slide.shapes.some(shape => shape.text && shape.text.trim() === 'Other deck title'));
    });

    await t.test('presentation summary', async () => {
      const [http, websocket] = await Promise.all([
        getJson(`/json/${target.id}/presentation`),
        request(client, 'PowerPoint.getPresentationState', { targetId: target.id })
      ]);
      assert.equal(http.status, 200);
      assert.deepEqual(http.body, websocket);
      assert.equal(typeof http.body.name, 'string');
      assert.equal(typeof http.body.url, 'string');
      assert.equal(http.body.id, target.id);
      assert.equal(http.body.slideCount, 1);
      assert.equal(http.body.saved, false);
      assert.equal(http.body.readOnly, false);
    });

    const slidesHttp = await getJson(`/json/${target.id}/slides`);
    const slidesWs = await request(client, 'PowerPoint.getSlides', { targetId: target.id });
    assert.equal(slidesHttp.status, 200);
    assert.deepEqual(slidesHttp.body, slidesWs);
    assert.equal(slidesHttp.body.id, target.id);
    assert.equal(slidesHttp.body.slides.length, 1);
    const slide = slidesHttp.body.slides[0];
    assert.equal(slide.slideIndex, 1);
    assert.equal(typeof slide.name, 'string');
    assert.equal(typeof slide.hidden, 'boolean');
    assert.ok(Number.isInteger(slide.slideId) && slide.slideId > 0);

    let mutationSlideIds;

    await t.test('creates slides over HTTP and WebSocket', async () => {
      const httpResponse = await fetch(`${baseUrl}/json/${target.id}/slides`, {
        method: 'POST',
        headers: { 'content-type': 'application/json' },
        body: JSON.stringify({ layout: 12 })
      });
      const http = await httpResponse.json();
      assert.equal(httpResponse.status, 200);
      assert.equal(http.id, target.id);
      assert.equal(http.slideIndex, 2);
      assert.ok(Number.isInteger(http.slideId) && http.slideId > 0);

      const websocket = await request(client, 'PowerPoint.addSlide',
        { targetId: target.id, layout: 12 });
      assert.equal(websocket.id, target.id);
      assert.equal(websocket.slideIndex, 3);
      assert.ok(Number.isInteger(websocket.slideId) && websocket.slideId > 0);
      const slides = await request(client, 'PowerPoint.getSlides', { targetId: target.id });
      assert.ok(slides.slides.some(slide => slide.slideId === http.slideId));
      assert.ok(slides.slides.some(slide => slide.slideId === websocket.slideId));
      mutationSlideIds = [http.slideId, websocket.slideId];
    });

    await t.test('creates, edits, navigates, and removes tester content over both transports', async () => {
      const [httpSlideId, websocketSlideId] = mutationSlideIds;
      const mutateHttp = async (pathname, method, payload) => {
        const response = await fetch(`${baseUrl}${pathname}`, {
          method,
          ...(payload === undefined ? {} : {
            headers: { 'content-type': 'application/json' },
            body: JSON.stringify(payload)
          })
        });
        return { status: response.status, body: await response.json() };
      };
      const createShape = { shapeType: 1, left: 24, top: 36, width: 160, height: 48 };
      const httpShape = await mutateHttp(
        `/json/${target.id}/slides/${httpSlideId}/shapes`, 'POST',
        createShape);
      assert.equal(httpShape.status, 200);
      assert.equal(httpShape.body.slideId, httpSlideId);
      assert.ok(Number.isInteger(httpShape.body.shapeId));
      const websocketShape = await request(client, 'PowerPoint.createShape', {
        targetId: target.id, slideId: websocketSlideId, ...createShape
      });
      assert.equal(websocketShape.slideId, websocketSlideId);
      assert.ok(Number.isInteger(websocketShape.shapeId));

      const changedHttpText = await mutateHttp(
        `/json/${target.id}/slides/${httpSlideId}/shapes/${httpShape.body.shapeId}`,
        'PUT', { text: 'Updated through HTTP' });
      assert.equal(changedHttpText.status, 200);
      const changedWebSocketText = await request(client, 'PowerPoint.setShapeText', {
        targetId: target.id, slideId: websocketSlideId, shapeId: websocketShape.shapeId,
        text: 'Updated through WebSocket'
      });
      assert.equal(changedWebSocketText.text, 'Updated through WebSocket');
      for (const [slideId, expectedText] of [
        [httpSlideId, 'Updated through HTTP'],
        [websocketSlideId, 'Updated through WebSocket']
      ]) {
        const state = await request(client, 'PowerPoint.getSlideState', { targetId: target.id, slideId });
        assert.ok(state.slide.shapes.some(shape => shape.text === expectedText));
      }

      const currentHttp = await mutateHttp(`/json/${target.id}/view`, 'PUT', { slideId: httpSlideId });
      assert.equal(currentHttp.status, 200);
      assert.equal(currentHttp.body.slideId, httpSlideId);
      const currentWebSocket = await request(client, 'PowerPoint.setCurrentSlide', {
        targetId: target.id, slideId: slide.slideId
      });
      assert.equal(currentWebSocket.slideId, slide.slideId);
      const view = await request(client, 'PowerPoint.getViewState', { targetId: target.id });
      assert.ok(view.windows.some(window => window.currentSlideId === slide.slideId));

      const started = await mutateHttp(`/json/${target.id}/slide-show`, 'POST');
      assert.equal(started.status, 200);
      assert.equal(started.body.running, true);
      const next = await mutateHttp(
        `/json/${target.id}/slide-show/navigation`, 'POST', { action: 'next' });
      assert.equal(next.status, 200);
      assert.ok(next.body.windows.some(window => window.currentSlideId === httpSlideId));
      const goto = await request(client, 'PowerPoint.navigateSlideShow', {
        targetId: target.id, action: 'goto', slideId: websocketSlideId
      });
      assert.ok(goto.windows.some(window => window.currentSlideId === websocketSlideId));
      const previous = await request(client, 'PowerPoint.navigateSlideShow', {
        targetId: target.id, action: 'previous'
      });
      assert.ok(previous.windows.some(window => window.currentSlideId === httpSlideId));
      const stopped = await request(client, 'PowerPoint.stopSlideShow', { targetId: target.id });
      assert.equal(stopped.running, false);
      const startedByWebSocket = await request(client, 'PowerPoint.startSlideShow',
        { targetId: target.id });
      assert.equal(startedByWebSocket.running, true);
      const stoppedByHttp = await mutateHttp(`/json/${target.id}/slide-show`, 'DELETE');
      assert.equal(stoppedByHttp.status, 200);
      assert.equal(stoppedByHttp.body.running, false);

      const deletedHttpShape = await mutateHttp(
        `/json/${target.id}/slides/${httpSlideId}/shapes/${httpShape.body.shapeId}`, 'DELETE');
      assert.equal(deletedHttpShape.status, 200);
      const deletedWebSocketShape = await request(client, 'PowerPoint.deleteShape', {
        targetId: target.id, slideId: websocketSlideId, shapeId: websocketShape.shapeId
      });
      assert.equal(deletedWebSocketShape.deleted, true);
      const deletedHttpSlide = await mutateHttp(`/json/${target.id}/slides/${httpSlideId}`, 'DELETE');
      assert.equal(deletedHttpSlide.status, 200);
      const deletedWebSocketSlide = await request(client, 'PowerPoint.deleteSlide', {
        targetId: target.id, slideId: websocketSlideId
      });
      assert.equal(deletedWebSocketSlide.deleted, true);
      const remaining = await request(client, 'PowerPoint.getSlides', { targetId: target.id });
      assert.deepEqual(remaining.slides.map(item => item.slideId), [slide.slideId]);
    });

    await t.test('slide detail includes separately created shape and stable metadata', async () => {
      const [http, websocket] = await Promise.all([
        getJson(`/json/${target.id}/slides/${slide.slideId}`),
        request(client, 'PowerPoint.getSlideState', { targetId: target.id, slideId: slide.slideId })
      ]);
      assert.equal(http.status, 200);
      assert.deepEqual(http.body, websocket);
      assert.equal(http.body.id, target.id);
      assert.equal(http.body.slide.slideId, slide.slideId);
      const textShape = http.body.slide.shapes.find(shape => shape.text && shape.text.trim() === 'State API test title');
      assert.ok(textShape, 'created shape should expose its separately updated text');
      assert.ok(Number.isInteger(textShape.shapeType));
      assert.ok(Number.isInteger(textShape.shapeId) && textShape.shapeId > 0);
      assert.ok(Number.isInteger(textShape.zOrderPosition) && textShape.zOrderPosition > 0);
      for (const field of ['left', 'top', 'width', 'height'])
        assert.ok(Number.isFinite(textShape.bounds[field]), `shape bounds.${field} should be numeric`);
    });

    await t.test('editing view reports current slide and selected shape', async () => {
      const presentation = await request(client, 'PowerPoint.getPresentationState', { targetId: target.id });
      selectFirstShape(presentation.url);
      const slides = await request(client, 'PowerPoint.getSlides', { targetId: target.id });
      const slideState = await request(client, 'PowerPoint.getSlideState',
        { targetId: target.id, slideId: slides.slides[0].slideId });
      const textShape = slideState.slide.shapes.find(shape =>
        shape.text && shape.text.trim() === 'State API test title');
      assert.ok(textShape);
      const [http, websocket] = await Promise.all([
        getJson(`/json/${target.id}/view`),
        request(client, 'PowerPoint.getViewState', { targetId: target.id })
      ]);
      assert.equal(http.status, 200);
      assert.deepEqual(http.body, websocket);
      assert.equal(http.body.id, target.id);
      assert.equal(http.body.available, true);
      assert.ok(http.body.windows.some(window =>
        window.currentSlideId === slides.slides[0].slideId &&
        window.currentSlideIndex === 1 &&
        window.selection?.type === 2 &&
        window.selection.shapeIds?.includes(textShape.shapeId)));
    });

    await t.test('slide-show state identifies a running show and reports stopped state', async () => {
      const presentation = await request(client, 'PowerPoint.getPresentationState', { targetId: target.id });
      controlSlideShow('start', presentation.url);
      try {
        const [http, websocket] = await Promise.all([
          getJson(`/json/${target.id}/slide-show`),
          request(client, 'PowerPoint.getSlideShowState', { targetId: target.id })
        ]);
        assert.equal(http.status, 200);
        assert.deepEqual(http.body, websocket);
        assert.equal(http.body.id, target.id);
        assert.equal(http.body.running, true);
        assert.ok(http.body.windows.length > 0);
        assert.ok(http.body.windows.some(window =>
          window.currentSlideId === slide.slideId &&
          Number.isInteger(window.currentSlideIndex) &&
          Number.isInteger(window.currentShowPosition)));
      } finally {
        controlSlideShow('stop', presentation.url);
      }
      const [stoppedHttp, stoppedWs] = await Promise.all([
        getJson(`/json/${target.id}/slide-show`),
        request(client, 'PowerPoint.getSlideShowState', { targetId: target.id })
      ]);
      assert.equal(stoppedHttp.status, 200);
      assert.deepEqual(stoppedHttp.body, stoppedWs);
      assert.equal(stoppedHttp.body.running, false);
      assert.deepEqual(stoppedHttp.body.windows, []);
    });

    await t.test('invalid targets, IDs, and mutation parameters return documented errors', async () => {
      const invalidTarget = await getJson('/json/not-a-guid/slides');
      assert.equal(invalidTarget.status, 400);
      const missingTarget = await getJson('/json/00000000-0000-0000-0000-000000000000/slides');
      assert.equal(missingTarget.status, 404);
      const invalidSlideId = await getJson(`/json/${target.id}/slides/0`);
      assert.equal(invalidSlideId.status, 400);
      const unsupportedVerb = await fetch(`${baseUrl}/json/${target.id}/view`, { method: 'DELETE' });
      assert.equal(unsupportedVerb.status, 405);
      const invalidShape = await fetch(`${baseUrl}/json/${target.id}/slides/${slide.slideId}/shapes`, {
        method: 'POST',
        headers: { 'content-type': 'application/json' },
        body: JSON.stringify({ shapeType: 1, left: 0, top: 0, width: 0, height: 10 })
      });
      assert.equal(invalidShape.status, 400);
      const invalidShapeWs = await request(client, 'PowerPoint.createShape', {
        targetId: target.id, slideId: slide.slideId,
        shapeType: 1, left: 0, top: 0, width: 0, height: 10
      }).then(() => null, error => error);
      assert.equal(invalidShapeWs.code, -32602);
      const combined = {
        shapeType: 1, left: 24, top: 36, width: 160, height: 48, text: 'Use shape text'
      };
      const combinedHttp = await fetch(`${baseUrl}/json/${target.id}/slides/${slide.slideId}/shapes`, {
        method: 'POST', headers: { 'content-type': 'application/json' },
        body: JSON.stringify(combined)
      });
      assert.equal(combinedHttp.status, 400);
      const combinedWs = await request(client, 'PowerPoint.createShape', {
        targetId: target.id, slideId: slide.slideId, ...combined
      }).then(() => null, error => error);
      assert.equal(combinedWs.code, -32602);
      const unknownSlideId = slide.slideId + 100000;
      const missingSlide = await getJson(`/json/${target.id}/slides/${unknownSlideId}`);
      assert.equal(missingSlide.status, 404);
      const missingSlideWs = await request(client, 'PowerPoint.getSlideState',
        { targetId: target.id, slideId: unknownSlideId }).then(() => null, error => error);
      assert.ok(missingSlideWs);
      assert.equal(missingSlideWs.code, -32004);
      const malformedTarget = await request(client, 'PowerPoint.getSlides',
        { targetId: 'not-a-guid' }).then(() => null, error => error);
      assert.ok(malformedTarget);
      assert.equal(malformedTarget.code, -32602);
    });
  } finally {
    for (const id of createdTargets)
      await fetch(`${baseUrl}/json/close/${id}?force`, { method: 'PUT' }).catch(() => {});
    client.close();
    fs.rmSync(directory, { recursive: true, force: true, maxRetries: 10, retryDelay: 250 });
  }
});
