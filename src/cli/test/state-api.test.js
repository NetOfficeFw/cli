'use strict';

const assert = require('node:assert/strict');
const { execFileSync } = require('node:child_process');
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

function controlSlideShow(action, name) {
  const command = action === 'start'
    ? `$ppt = [Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application'); $presentation = @($ppt.Presentations | Where-Object { $_.Name -eq $env:NETOFFICE_TEST_DOCUMENT }) | Select-Object -First 1; if (-not $presentation) { throw 'Test presentation was not found' }; $null = $presentation.SlideShowSettings.Run()`
    : `$ppt = [Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application'); foreach ($window in @($ppt.SlideShowWindows)) { if ($window.Presentation.Name -eq $env:NETOFFICE_TEST_DOCUMENT) { $window.View.Exit() } }`;
  execFileSync('powershell.exe', ['-NoProfile', '-NonInteractive', '-Command', command], {
    env: { ...process.env, NETOFFICE_TEST_DOCUMENT: name },
    stdio: 'ignore'
  });
}

function selectTitleShape(name) {
  const command = `$ppt = [Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application'); $presentation = @($ppt.Presentations | Where-Object { $_.Name -eq $env:NETOFFICE_TEST_DOCUMENT }) | Select-Object -First 1; if (-not $presentation) { throw 'Test presentation was not found' }; $window = $presentation.Windows.Item(1); $window.Activate(); $window.View.GotoSlide(1); $presentation.Slides.Item(1).Shapes.Title.Select()`;
  execFileSync('powershell.exe', ['-NoProfile', '-NonInteractive', '-Command', command], {
    env: { ...process.env, NETOFFICE_TEST_DOCUMENT: name },
    stdio: 'ignore'
  });
}

async function request(client, method, params = {}) {
  return client.request(method, params, deadline());
}

test('PowerPoint presentation state is readable over HTTP and WebSocket', async t => {
  assert.ok(Number.isInteger(port) && port > 0 && port <= 65535, 'NETOFFICE_PORT must be a valid TCP port');
  execFileSync(process.execPath, [path.join(__dirname, '..', 'bin', 'netoffice.js'),
    'powerpoint', 'launch', '--port', String(port), '--timeout', '30000'], { stdio: 'inherit' });
  const client = await Connection.connect(socketUrl, deadline());
  const createdTargets = [];
  try {
    const before = await getJson('/json/list');
    assert.equal(before.status, 200);
    const previousIds = new Set(before.body.map(target => target.id));
    const created = await request(client, 'PowerPoint.newPresentation', { title: 'State API test title' });
    assert.equal(created.slideCount, 1);
    const after = await getJson('/json/list');
    assert.equal(after.status, 200);
    const target = after.body.find(candidate => !previousIds.has(candidate.id));
    assert.ok(target, 'new presentation should appear as a document target');
    createdTargets.push(target.id);
    const secondCreated = await request(client, 'PowerPoint.newPresentation', { title: 'Other deck title' });
    assert.equal(secondCreated.slideCount, 1);
    const withSecond = await getJson('/json/list');
    const other = withSecond.body.find(candidate =>
      !previousIds.has(candidate.id) && candidate.id !== target.id);
    assert.ok(other, 'second presentation should have an independent document target');
    createdTargets.push(other.id);

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

    await t.test('slide detail includes title shape and stable shape metadata', async () => {
      const [http, websocket] = await Promise.all([
        getJson(`/json/${target.id}/slides/${slide.slideId}`),
        request(client, 'PowerPoint.getSlideState', { targetId: target.id, slideId: slide.slideId })
      ]);
      assert.equal(http.status, 200);
      assert.deepEqual(http.body, websocket);
      assert.equal(http.body.id, target.id);
      assert.equal(http.body.slide.slideId, slide.slideId);
      const title = http.body.slide.shapes.find(shape => shape.text && shape.text.trim() === 'State API test title');
      assert.ok(title, 'title slide should expose its configured title text');
      assert.ok(Number.isInteger(title.shapeType));
      assert.ok(Number.isInteger(title.shapeId) && title.shapeId > 0);
      assert.ok(Number.isInteger(title.zOrderPosition) && title.zOrderPosition > 0);
      for (const field of ['left', 'top', 'width', 'height'])
        assert.ok(Number.isFinite(title.bounds[field]), `shape bounds.${field} should be numeric`);
    });

    await t.test('editing view reports current slide and selected shape', async () => {
      const presentation = await request(client, 'PowerPoint.getPresentationState', { targetId: target.id });
      selectTitleShape(presentation.name);
      const slides = await request(client, 'PowerPoint.getSlides', { targetId: target.id });
      const slideState = await request(client, 'PowerPoint.getSlideState',
        { targetId: target.id, slideId: slides.slides[0].slideId });
      const titleShape = slideState.slide.shapes.find(shape =>
        shape.text && shape.text.trim() === 'State API test title');
      assert.ok(titleShape);
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
        window.selection.shapeIds?.includes(titleShape.shapeId)));
    });

    await t.test('slide-show state identifies a running show and reports stopped state', async () => {
      const presentation = await request(client, 'PowerPoint.getPresentationState', { targetId: target.id });
      controlSlideShow('start', presentation.name);
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
        controlSlideShow('stop', presentation.name);
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

    await t.test('invalid targets and slide IDs return the documented errors', async () => {
      const invalidTarget = await getJson('/json/not-a-guid/slides');
      assert.equal(invalidTarget.status, 400);
      const missingTarget = await getJson('/json/00000000-0000-0000-0000-000000000000/slides');
      assert.equal(missingTarget.status, 404);
      const invalidSlideId = await getJson(`/json/${target.id}/slides/0`);
      assert.equal(invalidSlideId.status, 400);
      const unsupportedVerb = await fetch(`${baseUrl}/json/${target.id}/slides`, { method: 'POST' });
      assert.equal(unsupportedVerb.status, 405);
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
  }
});
