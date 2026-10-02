'use strict';

const usage = `Usage:
  netoffice powerpoint launch [--timeout <milliseconds>]
  netoffice presentation new [--title <text>] [--timeout <milliseconds>]
  netoffice slide title <text> [--slide <positive integer>] [--timeout <milliseconds>]
  netoffice --help

Commands connect to the native PowerPoint add-in at 127.0.0.1:50051.
Launch reuses a ready PowerPoint session or starts visible PowerPoint on Windows.
New creates one title slide in a new active presentation (default title: empty).
Slide title updates the active presentation (default slide: 1).
--timeout sets the total operation deadline in milliseconds (default: 10000).
Use -- before positional title text starting with a dash.
`;

function positiveInteger(value, option, maximum = Number.MAX_SAFE_INTEGER) {
  if (!/^[0-9]+$/.test(value || '') || !Number.isSafeInteger(Number(value)) || Number(value) < 1 || Number(value) > maximum) {
    throw new Error(`${option} requires a positive integer${maximum !== Number.MAX_SAFE_INTEGER ? ` no greater than ${maximum}` : ''}.`);
  }
  return Number(value);
}

function parseArguments(args) {
  if (args.length === 1 && (args[0] === '--help' || args[0] === '-h')) return { help: true };
  const command = `${args[0] || ''} ${args[1] || ''}`;
  if (!['powerpoint launch', 'presentation new', 'slide title'].includes(command)) {
    throw new Error('Expected powerpoint launch, presentation new, or slide title. Use --help for usage.');
  }
  const result = { command, timeout: 10000, title: '', slide: 1 };
  const seen = new Set();
  const positionals = [];
  let literal = false;
  for (let index = 2; index < args.length; index++) {
    const arg = args[index];
    if (!literal && arg === '--') { literal = true; continue; }
    if (!literal && (arg === '--help' || arg === '-h')) { result.help = true; continue; }
    if (!literal && arg.startsWith('-')) {
      const allowed = arg === '--timeout' || (arg === '--title' && command === 'presentation new') || (arg === '--slide' && command === 'slide title');
      if (!allowed) throw new Error(`Unknown option: ${arg}`);
      if (seen.has(arg)) throw new Error(`Repeated option: ${arg}`);
      seen.add(arg);
      if (++index >= args.length || args[index].startsWith('--')) throw new Error(`Missing value for ${arg}.`);
      const value = args[index];
      if (arg === '--title') result.title = value;
      else if (arg === '--slide') result.slide = positiveInteger(value, arg, 2147483647);
      else result.timeout = positiveInteger(value, arg, 2147483647);
    } else positionals.push(arg);
  }
  if (command === 'slide title') {
    if (positionals.length !== 1 && !(result.help && positionals.length === 0)) throw new Error('slide title requires exactly one text argument; quote text containing spaces.');
    result.title = positionals[0];
  } else if (positionals.length) throw new Error(`Unexpected argument: ${positionals[0]}`);
  return result;
}

module.exports = { parseArguments, usage };
