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
