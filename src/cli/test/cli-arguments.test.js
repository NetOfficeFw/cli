'use strict';

const assert = require('node:assert/strict');
const { spawnSync } = require('node:child_process');
const path = require('node:path');
const { test } = require('node:test');
const { parseArguments } = require('../lib/arguments');

const cli = path.join(__dirname, '..', 'bin', 'netoffice.js');
const target = '00000000-0000-0000-0000-000000000001';

function rejected(args, message) {
  const result = spawnSync(process.execPath, [cli, ...args], { encoding: 'utf8' });
  assert.equal(result.status, 1);
  assert.match(result.stderr, message);
}

test('target-scoped CLI commands reject absent and malformed document IDs', () => {
  rejected(['slide', 'add'], /requires --target/);
  rejected(['slide', 'add', '--target', 'not-a-guid'], /canonical GUID/);
});

test('presentation close requires a target and permits force only as a single flag', () => {
  rejected(['presentation', 'close'], /requires --target/);
  rejected(['presentation', 'close', '--target', 'not-a-guid'], /canonical GUID/);
  assert.deepEqual(parseArguments(['presentation', 'close', '--target', target]).params,
    { targetId: target });
  assert.deepEqual(parseArguments(['presentation', 'close', '--target', target, '--force']).params,
    { targetId: target, force: true });
  rejected(['presentation', 'close', '--target', target, '--force', '--force'],
    /Repeated option: --force/);
});

test('shape creation requires complete, finite, positive geometry', () => {
  const base = ['shape', 'add', '--target', target, '--slide-id', '1', '--type', '1'];
  rejected(base, /requires --left/);
  rejected([...base, '--left', '0', '--top', '0', '--width', '0', '--height', '48'], /must be positive/);
  rejected([...base, '--left', '0', '--top', '0', '--width', 'NaN', '--height', '48'], /finite number/);
  rejected([...base, '--left', '0', '--top', '0', '--width', '50', '--height', '48',
    '--text', 'Separate operation'], /Unknown option: --text/);
});

test('slide-show navigation requires a stable slide ID only for goto', () => {
  rejected(['slideshow', 'goto', '--target', target], /requires --slide-id/);
  rejected(['slideshow', 'next', '--target', target, '--slide-id', '1'], /Unknown option: --slide-id/);
});

test('creation and opening require paths; legacy name and implicit title commands are rejected', () => {
  rejected(['presentation', 'new'], /requires --path/);
  rejected(['presentation', 'open'], /requires --path/);
  rejected(['presentation', 'open', '--path', ''], /non-empty file path/);
  assert.deepEqual(parseArguments(['presentation', 'open', '--path', 'Deck.pptx']).params,
    { path: 'Deck.pptx' });
  rejected(['presentation', 'new', '--name', 'Old behavior'], /Unknown option: --name/);
  rejected(['presentation', 'new', '--title', 'Old behavior'], /Unknown option: --title/);
  rejected(['slide', 'title', 'Old behavior'], /Unknown command/);
});

test('shutdown accepts an optional flag to discard unsaved presentations', () => {
  assert.deepEqual(parseArguments(['powerpoint', 'shutdown', '--force']).params, { force: true });
  rejected(['powerpoint', 'shutdown', '--force', '--force'], /Repeated option: --force/);
  rejected(['powerpoint', 'shutdown', '--target', target], /Unknown option: --target/);
  rejected(['presentation', 'new', '--path', 'Deck', '--force'], /Unknown option: --force/);
});

test('help lists the presentation command syntax', () => {
  const result = spawnSync(process.execPath, [cli, '--help'], { encoding: 'utf8' });
  assert.equal(result.status, 0);
  assert.ok(result.stdout.includes('netoffice presentation close --target <id> [--force]'));
  assert.ok(result.stdout.includes('netoffice presentation new --path <file>'));
  assert.ok(result.stdout.includes('netoffice presentation open --path <file>'));
});

test('slide-or-master commands require exactly one container and send master as a flag', () => {
  const base = ['shape', 'fill', '--target', target, '--shape-id', '3', '--color', '#112233'];
  rejected(base, /exactly one of --slide-id, --master, or --custom-layout/);
  rejected([...base, '--slide-id', '1', '--master'], /exactly one of --slide-id, --master, or --custom-layout/);
  rejected([...base, '--master', '--custom-layout', '2'], /exactly one of --slide-id, --master, or --custom-layout/);
  assert.deepEqual(parseArguments([...base, '--master']).params,
    { targetId: target, shapeId: 3, color: '#112233', master: true });
  assert.deepEqual(parseArguments([...base, '--custom-layout', '2']).params,
    { targetId: target, shapeId: 3, color: '#112233', customLayout: 2 });
  rejected(['slide', 'move', '--target', target, '--master', '--index', '1'], /Unknown option: --master/);
  rejected(['shape', 'animation', '--target', target, '--master', '--shape-id', '1', '--effect', '10'],
    /Unknown option: --master/);
});

test('typed values are validated before any request', () => {
  const shape = ['--target', target, '--slide-id', '1', '--shape-id', '2'];
  rejected(['shape', 'fill', ...shape, '--color', 'red'], /#RRGGBB/);
  rejected(['shape', 'font', ...shape, '--bold', 'yes'], /true or false/);
  rejected(['shape', 'group', '--target', target, '--slide-id', '1', '--shape-ids', '2,x'], /positive integer/);
  rejected(['chart', 'data', ...shape, '--categories', 'Q1', '--series', 'Revenue'], /<name>:<v,v,...>/);
  rejected(['slide', 'add', '--target', target, '--layout', '2', '--custom-layout', '3'], /only one of/);
  rejected(['slide', 'add', '--target', target, '--custom-layout', '3', '--custom-layout-name', 'Content'], /only one of/);
  rejected(['smartart', 'node', ...shape, '--text', 'x'], /exactly one of --index or --add/);
  assert.deepEqual(parseArguments(['chart', 'data', ...shape, '--categories', 'Q1, Q2',
    '--series', 'Rev: 1,2;Cost:3,-4']).params.series,
    [{ name: 'Rev', values: [1, 2] }, { name: 'Cost', values: [3, -4] }]);
});

test('theme colors are grouped and paths resolve to absolute', () => {
  assert.deepEqual(parseArguments(['presentation', 'colors', '--target', target,
    '--accent1', '#1f3864', '--followed-hyperlink', '#000000']).params,
    { targetId: target, colors: { accent1: '#1F3864', followedHyperlink: '#000000' } });
  assert.equal(parseArguments(['presentation', 'save', '--target', target, '--path', 'Copy.pptx']).params.path,
    path.resolve('Copy.pptx'));
  assert.deepEqual(parseArguments(['smartart', 'layouts']).params, {});
});
