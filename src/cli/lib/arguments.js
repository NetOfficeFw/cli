'use strict';

const commands = {
  'powerpoint shutdown': ['PowerPoint.prepareShutdown', [], ['force']],
  'presentation open': ['PowerPoint.openPresentation', ['path'], []],
  'presentation state': ['PowerPoint.getPresentationState', ['target'], []],
  'presentation close': ['PowerPoint.closePresentation', ['target'], ['force']],
  'slide list': ['PowerPoint.getSlides', ['target'], []],
  'slide show': ['PowerPoint.getSlideState', ['target', 'slideId'], []],
  'view state': ['PowerPoint.getViewState', ['target'], []],
  'slideshow state': ['PowerPoint.getSlideShowState', ['target'], []],
  'slide add': ['PowerPoint.addSlide', ['target'], ['layout']],
  'shape add': ['PowerPoint.createShape', ['target', 'slideId', 'shapeType', 'left', 'top', 'width', 'height'], []],
  'shape text': ['PowerPoint.setShapeText', ['target', 'slideId', 'shapeId', 'text'], []],
  'shape delete': ['PowerPoint.deleteShape', ['target', 'slideId', 'shapeId'], []],
  'slide delete': ['PowerPoint.deleteSlide', ['target', 'slideId'], []],
  'slide current': ['PowerPoint.setCurrentSlide', ['target', 'slideId'], []],
  'slideshow start': ['PowerPoint.startSlideShow', ['target'], []],
  'slideshow stop': ['PowerPoint.stopSlideShow', ['target'], []],
  'slideshow next': ['PowerPoint.navigateSlideShow', ['target'], []],
  'slideshow previous': ['PowerPoint.navigateSlideShow', ['target'], []],
  'slideshow goto': ['PowerPoint.navigateSlideShow', ['target', 'slideId'], []]
};

const usage = `Usage: netoffice <command> [options]

PowerPoint:
  netoffice powerpoint launch
    Start visible PowerPoint and wait for the add-in; reuse an already-ready session.
  netoffice powerpoint shutdown [--force]
    Quit PowerPoint; refuse unsaved presentations unless --force discards them.

Presentations:
  netoffice presentation new --path <file>
    Create a blank .pptx at that relative or absolute path; append .pptx if omitted.
  netoffice presentation open --path <file>
    Open an existing presentation at that relative or absolute path; return its ID.
  netoffice presentation list
    List open presentations and their target IDs.
  netoffice presentation state --target <id>
    Read the deck's name, path, saved/read-only status, and slide count.
  netoffice presentation close --target <id> [--force]
    Close only that deck; refuse unsaved changes unless --force discards them.

Slides:
  netoffice slide list --target <id>
    List slides in order with stable slide IDs.
  netoffice slide show --target <id> --slide-id <id>
    Read one slide and its shapes; does not display it or start a slideshow.
  netoffice slide add --target <id> [--layout <1..12>]
    Append a slide with the chosen layout (default 12: blank); return its slide ID.
  netoffice slide current --target <id> --slide-id <id>
    Select a slide in the editing window; does not navigate a running slideshow.
  netoffice slide delete --target <id> --slide-id <id>
    Delete the identified slide from the open presentation.

Shapes:
  netoffice shape add --target <id> --slide-id <id> --type <auto-shape-type> --left <pt> --top <pt> --width <pt> --height <pt>
    Add an auto-shape at the given point coordinates and size; return its shape ID.
  netoffice shape text --target <id> --slide-id <id> --shape-id <id> --text <text>
    Replace all plain text in that shape; separate from shape creation.
  netoffice shape delete --target <id> --slide-id <id> --shape-id <id>
    Delete the identified shape from its slide.

Views and presentation mode:
  netoffice view state --target <id>
    Read editing windows, current slide, and selection.
  netoffice slideshow state --target <id>
    Read running slideshow windows and their current slides.
  netoffice slideshow start --target <id>
    Start the target presentation in slideshow mode.
  netoffice slideshow stop --target <id>
    Exit its running slideshow windows.
  netoffice slideshow next --target <id>
    Advance a running slideshow by one slide.
  netoffice slideshow previous --target <id>
    Move a running slideshow back by one slide.
  netoffice slideshow goto --target <id> --slide-id <id>
    Jump a running slideshow to the stable slide ID.

Options:
  --path <file>     Presentation new destination or existing presentation to open.
  --layout <1..12>  Slide layout for slide add (default 12: blank).
  --type <number>   PowerPoint MsoAutoShapeType (1..255) for shape add.
  --text <text>     Replacement plain text for shape text.
  --force           Shutdown: discard all decks; presentation close: discard only its target.
  --port <port>     Add-in port (default 50051); accepted by every command.
  --timeout <ms>    Total deadline (default 10000); accepted by every command.

IDs and units:
  --target <id>     Presentation ID from presentation list or presentation new.
  --slide-id <id>   Stable SlideID from slide list or slide add; not a slide index.
  --shape-id <id>   Shape ID from shape add or slide show.
  --left <pt>, --top <pt>       Shape position in points.
  --width <pt>, --height <pt>   Positive shape dimensions in points.

Slide layouts (--layout):
   1  Title                 2  Text
   3  Two-column text       4  Table
   5  Text + Chart          6  Chart + Text
   7  Organization chart    8  Chart
   9  Text + ClipArt       10  ClipArt + Text
  11  Title only           12  Blank (default)

Mutations are never retried. Successful operational commands except powerpoint launch print JSON.
Run netoffice --help to show this reference.
`;

const optionNames = {
  '--target': 'target', '--slide-id': 'slideId', '--shape-id': 'shapeId',
  '--layout': 'layout', '--type': 'shapeType', '--left': 'left', '--top': 'top',
  '--width': 'width', '--height': 'height', '--text': 'text', '--force': 'force',
  '--path': 'path'
};

function positiveInteger(value, option, maximum = Number.MAX_SAFE_INTEGER) {
  if (!/^[0-9]+$/.test(value || '') || !Number.isSafeInteger(Number(value)) || Number(value) < 1 || Number(value) > maximum) {
    throw new Error(`${option} requires a positive integer${maximum !== Number.MAX_SAFE_INTEGER ? ` no greater than ${maximum}` : ''}.`);
  }
  return Number(value);
}

function finiteNumber(value, option) {
  if (!/^-?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/.test(value || '') || !Number.isFinite(Number(value))) {
    throw new Error(`${option} requires a finite number.`);
  }
  return Number(value);
}

function parseArguments(args) {
  if (args.length === 1 && (args[0] === '--help' || args[0] === '-h')) return { help: true };
  const command = `${args[0] || ''} ${args[1] || ''}`;
  const legacy = ['powerpoint launch', 'presentation new'];
  const spec = commands[command];
  if (!legacy.includes(command) && command !== 'presentation list' && !spec) {
    throw new Error('Unknown command. Use --help for usage.');
  }
  const result = { command, timeout: 10000, port: 50051 };
  const seen = new Set();
  for (let index = 2; index < args.length; index++) {
    const arg = args[index];
    if (arg === '--help' || arg === '-h') { result.help = true; continue; }
    if (arg.startsWith('-')) {
      const field = optionNames[arg];
      const allowed = arg === '--timeout' || arg === '--port' ||
        (arg === '--path' && command === 'presentation new') ||
        (spec && field && (spec[1].includes(field) || spec[2].includes(field)));
      if (!allowed) throw new Error(`Unknown option: ${arg}`);
      if (seen.has(arg)) throw new Error(`Repeated option: ${arg}`);
      seen.add(arg);
      if (arg === '--force') { result.force = true; continue; }
      if (++index >= args.length || args[index].startsWith('--')) throw new Error(`Missing value for ${arg}.`);
      const value = args[index];
      if (field === 'path') {
        if (!value) throw new Error('--path requires a non-empty file path.');
        result.path = value;
      } else if (arg === '--port') result.port = positiveInteger(value, arg, 65535);
      else if (arg === '--timeout') result.timeout = positiveInteger(value, arg, 2147483647);
      else if (field === 'target') {
        if (!/^[a-fA-F0-9]{8}(?:-[a-fA-F0-9]{4}){3}-[a-fA-F0-9]{12}$/.test(value)) {
          throw new Error('--target requires a canonical GUID from presentation list.');
        }
        result.target = value;
      } else if (field === 'text') result.text = value;
      else if (['left', 'top', 'width', 'height'].includes(field)) result[field] = finiteNumber(value, arg);
      else result[field] = positiveInteger(value, arg,
        field === 'layout' ? 12 : field === 'shapeType' ? 255 : 2147483647);
    } else throw new Error(`Unexpected argument: ${arg}`);
  }
  if (command === 'presentation new' && !result.help && !Object.hasOwn(result, 'path')) {
    throw new Error('presentation new requires --path.');
  }
  if (spec && !result.help) {
    for (const field of spec[1]) {
      if (!Object.hasOwn(result, field)) throw new Error(`${command} requires --${field.replace(/[A-Z]/g, letter => `-${letter.toLowerCase()}`)}.`);
    }
    if ((Object.hasOwn(result, 'width') && result.width <= 0) ||
        (Object.hasOwn(result, 'height') && result.height <= 0)) {
      throw new Error('--width and --height must be positive.');
    }
    const params = {};
    for (const field of [...spec[1], ...spec[2]]) {
      if (Object.hasOwn(result, field)) params[field === 'target' ? 'targetId' : field] = result[field];
    }
    if (command.startsWith('slideshow ') && ['next', 'previous', 'goto'].includes(command.split(' ')[1])) {
      params.action = command.split(' ')[1];
    }
    result.method = spec[0];
    result.params = params;
  }
  return result;
}

module.exports = { parseArguments, usage };
