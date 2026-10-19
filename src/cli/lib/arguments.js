'use strict';

const path = require('node:path');

// Option types: target (GUID), id (positive int32), integer, number, positive (number > 0),
// text (any string), name (non-empty string), file (non-empty path kept as given),
// absolute (non-empty path resolved against the current directory), boolean (true|false),
// flag (no value), color (#RRGGBB), ids (comma-separated positive int32 list),
// list (comma-separated strings), series (Name:1,2;Other:3,4), points (x,y;x,y;... at least two).
const options = {
  '--target': ['target', 'target'], '--slide-id': ['slideId', 'id'], '--shape-id': ['shapeId', 'id'],
  '--master': ['master', 'flag'], '--force': ['force', 'flag'], '--add': ['add', 'flag'],
  '--path': ['path', 'file'], '--layout': ['layout', 'id'], '--custom-layout': ['customLayout', 'id'],
  '--custom-layout-name': ['customLayoutName', 'name'],
  '--type': ['shapeType', 'id'], '--left': ['left', 'number'], '--top': ['top', 'number'],
  '--width': ['width', 'positive'], '--height': ['height', 'positive'], '--text': ['text', 'text'],
  '--name': ['name', 'name'], '--index': ['index', 'id'],
  '--major': ['major', 'name'], '--minor': ['minor', 'name'],
  '--dark1': ['dark1', 'color'], '--light1': ['light1', 'color'], '--dark2': ['dark2', 'color'],
  '--light2': ['light2', 'color'], '--accent1': ['accent1', 'color'], '--accent2': ['accent2', 'color'],
  '--accent3': ['accent3', 'color'], '--accent4': ['accent4', 'color'], '--accent5': ['accent5', 'color'],
  '--accent6': ['accent6', 'color'], '--hyperlink': ['hyperlink', 'color'],
  '--followed-hyperlink': ['followedHyperlink', 'color'],
  '--color': ['color', 'color'], '--color2': ['color2', 'color'], '--gradient-style': ['gradientStyle', 'name'],
  '--transparency': ['transparency', 'number'], '--picture': ['picture', 'absolute'],
  '--follow-master': ['followMaster', 'boolean'], '--footer': ['footer', 'boolean'],
  '--slide-number': ['slideNumber', 'boolean'], '--date': ['date', 'boolean'], '--date-text': ['dateText', 'text'],
  '--effect': ['effect', 'integer'], '--duration': ['duration', 'number'],
  '--advance-on-click': ['advanceOnClick', 'boolean'], '--advance-after': ['advanceAfter', 'number'],
  '--hidden': ['hidden', 'boolean'], '--width-px': ['widthPx', 'id'], '--height-px': ['heightPx', 'id'],
  '--trigger': ['trigger', 'name'], '--delay': ['delay', 'number'],
  '--visible': ['visible', 'boolean'], '--weight': ['weight', 'number'], '--dash': ['dash', 'name'],
  '--blur': ['blur', 'number'], '--offset-x': ['offsetX', 'number'], '--offset-y': ['offsetY', 'number'],
  '--size': ['size', 'number'], '--radius': ['radius', 'number'], '--bevel': ['bevel', 'integer'],
  '--bevel-depth': ['bevelDepth', 'number'], '--bevel-inset': ['bevelInset', 'number'],
  '--depth': ['depth', 'number'], '--material': ['material', 'integer'], '--style': ['style', 'integer'],
  '--value': ['value', 'number'], '--degrees': ['degrees', 'number'], '--direction': ['direction', 'name'],
  '--order': ['order', 'name'], '--source-shape-id': ['sourceShapeId', 'id'],
  '--orientation': ['orientation', 'name'], '--start': ['start', 'id'], '--length': ['length', 'id'],
  '--bold': ['bold', 'boolean'], '--italic': ['italic', 'boolean'], '--underline': ['underline', 'boolean'],
  '--paragraph': ['paragraph', 'id'], '--alignment': ['alignment', 'name'],
  '--space-before': ['spaceBefore', 'number'], '--space-after': ['spaceAfter', 'number'],
  '--line-spacing': ['lineSpacing', 'number'], '--indent-level': ['indentLevel', 'id'],
  '--bullet': ['bullet', 'name'], '--bullet-character': ['bulletCharacter', 'name'],
  '--anchor': ['anchor', 'name'], '--margin-left': ['marginLeft', 'number'],
  '--margin-right': ['marginRight', 'number'], '--margin-top': ['marginTop', 'number'],
  '--margin-bottom': ['marginBottom', 'number'], '--autosize': ['autoSize', 'name'],
  '--word-wrap': ['wordWrap', 'boolean'], '--begin-x': ['beginX', 'number'], '--begin-y': ['beginY', 'number'],
  '--end-x': ['endX', 'number'], '--end-y': ['endY', 'number'], '--connector-type': ['connectorType', 'name'],
  '--begin-shape-id': ['beginShapeId', 'id'], '--begin-site': ['beginSite', 'id'],
  '--end-shape-id': ['endShapeId', 'id'], '--end-site': ['endSite', 'id'],
  '--shape-ids': ['shapeIds', 'ids'], '--relative-to-slide': ['relativeToSlide', 'boolean'],
  '--rows': ['rows', 'id'], '--columns': ['columns', 'id'], '--row': ['row', 'id'], '--column': ['column', 'id'],
  '--fill-color': ['fillColor', 'color'], '--font-color': ['fontColor', 'color'],
  '--font-size': ['fontSize', 'number'], '--chart-type': ['chartType', 'integer'],
  '--categories': ['categories', 'list'], '--series': ['series', 'series'], '--file': ['file', 'file'],
  '--points': ['points', 'points']
};

const themeColors = ['dark1', 'light1', 'dark2', 'light2', 'accent1', 'accent2', 'accent3', 'accent4',
  'accent5', 'accent6', 'hyperlink', 'followedHyperlink'];
const geometry = ['left', 'top', 'width', 'height'];

// method; required fields; optional fields. 'slide' requires exactly one of --slide-id, --master,
// or --custom-layout (a container selector, not the new slide's layout as on slide add).
const commands = {
  'powerpoint shutdown': ['PowerPoint.prepareShutdown', [], ['force']],
  'presentation new': ['PowerPoint.newPresentation', ['path'], []],
  'presentation open': ['PowerPoint.openPresentation', ['path'], []],
  'presentation state': ['PowerPoint.getPresentationState', ['target'], []],
  'presentation close': ['PowerPoint.closePresentation', ['target'], ['force']],
  'presentation save': ['PowerPoint.savePresentation', ['target'], ['path']],
  'presentation size': ['PowerPoint.setSlideSize', ['target', 'width', 'height'], []],
  'presentation theme': ['PowerPoint.applyTheme', ['target', 'path'], []],
  'presentation colors': ['PowerPoint.setThemeColors', ['target'], themeColors],
  'presentation fonts': ['PowerPoint.setThemeFonts', ['target'], ['major', 'minor']],
  'layout list': ['PowerPoint.getLayouts', ['target'], []],
  'master show': ['PowerPoint.getMasterState', ['target'], []],
  'layout show': ['PowerPoint.getLayoutState', ['target', 'customLayout'], []],
  'slide list': ['PowerPoint.getSlides', ['target'], []],
  'slide show': ['PowerPoint.getSlideState', ['target', 'slideId'], []],
  'slide add': ['PowerPoint.addSlide', ['target'], ['layout', 'customLayout', 'customLayoutName']],
  'slide current': ['PowerPoint.setCurrentSlide', ['target', 'slideId'], []],
  'slide delete': ['PowerPoint.deleteSlide', ['target', 'slideId'], []],
  'slide move': ['PowerPoint.moveSlide', ['target', 'slideId', 'index'], []],
  'slide name': ['PowerPoint.setSlideName', ['target', 'slideId', 'name'], []],
  'slide notes': ['PowerPoint.setSlideNotes', ['target', 'slideId', 'text'], []],
  'slide transition': ['PowerPoint.setSlideTransition', ['target', 'slideId'],
    ['effect', 'duration', 'advanceOnClick', 'advanceAfter']],
  'slide hidden': ['PowerPoint.setSlideHidden', ['target', 'slideId', 'hidden'], []],
  'slide duplicate': ['PowerPoint.duplicateSlide', ['target', 'slideId'], []],
  'slide export': ['PowerPoint.exportSlide', ['target', 'slideId', 'path'], ['widthPx', 'heightPx']],
  'slide background': ['PowerPoint.setBackground', ['target', 'slide'],
    ['color', 'color2', 'gradientStyle', 'transparency', 'picture', 'followMaster']],
  'slide footer': ['PowerPoint.setHeadersFooters', ['target', 'slide'],
    ['text', 'footer', 'slideNumber', 'date', 'dateText']],
  'shape add': ['PowerPoint.createShape', ['target', 'slide', 'shapeType', ...geometry], []],
  'shape text': ['PowerPoint.setShapeText', ['target', 'slide', 'shapeId', 'text'], []],
  'shape delete': ['PowerPoint.deleteShape', ['target', 'slide', 'shapeId'], []],
  'shape state': ['PowerPoint.getShapeState', ['target', 'slide', 'shapeId'], []],
  'shape fill': ['PowerPoint.setShapeFill', ['target', 'slide', 'shapeId'],
    ['visible', 'color', 'color2', 'gradientStyle', 'transparency', 'picture']],
  'shape line': ['PowerPoint.setShapeLine', ['target', 'slide', 'shapeId'],
    ['visible', 'color', 'weight', 'dash', 'transparency']],
  'shape shadow': ['PowerPoint.setShapeShadow', ['target', 'slide', 'shapeId'],
    ['visible', 'color', 'blur', 'offsetX', 'offsetY', 'transparency', 'size']],
  'shape glow': ['PowerPoint.setShapeGlow', ['target', 'slide', 'shapeId', 'radius'], ['color', 'transparency']],
  'shape softedge': ['PowerPoint.setShapeSoftEdge', ['target', 'slide', 'shapeId', 'radius'], []],
  'shape 3d': ['PowerPoint.setShapeThreeD', ['target', 'slide', 'shapeId'],
    ['bevel', 'bevelDepth', 'bevelInset', 'depth', 'material']],
  'shape style': ['PowerPoint.setShapeStyle', ['target', 'slide', 'shapeId', 'style'], []],
  'shape adjust': ['PowerPoint.setShapeAdjustment', ['target', 'slide', 'shapeId', 'index', 'value'], []],
  'shape rotation': ['PowerPoint.setShapeRotation', ['target', 'slide', 'shapeId', 'degrees'], []],
  'shape flip': ['PowerPoint.flipShape', ['target', 'slide', 'shapeId', 'direction'], []],
  'shape name': ['PowerPoint.setShapeName', ['target', 'slide', 'shapeId', 'name'], []],
  'shape bounds': ['PowerPoint.setShapeBounds', ['target', 'slide', 'shapeId'], geometry],
  'shape zorder': ['PowerPoint.setShapeZOrder', ['target', 'slide', 'shapeId', 'order'], []],
  'shape duplicate': ['PowerPoint.duplicateShape', ['target', 'slide', 'shapeId'], ['left', 'top']],
  'shape copy-format': ['PowerPoint.copyShapeFormat', ['target', 'slide', 'shapeId', 'sourceShapeId'], []],
  'shape font': ['PowerPoint.setShapeFont', ['target', 'slide', 'shapeId'],
    ['start', 'length', 'name', 'size', 'bold', 'italic', 'underline', 'color']],
  'shape paragraph': ['PowerPoint.setShapeParagraph', ['target', 'slide', 'shapeId'],
    ['paragraph', 'alignment', 'spaceBefore', 'spaceAfter', 'lineSpacing', 'indentLevel', 'bullet', 'bulletCharacter']],
  'shape textframe': ['PowerPoint.setShapeTextFrame', ['target', 'slide', 'shapeId'],
    ['anchor', 'marginLeft', 'marginRight', 'marginTop', 'marginBottom', 'autoSize', 'wordWrap']],
  'shape animation': ['PowerPoint.addAnimation', ['target', 'slideId', 'shapeId', 'effect'],
    ['trigger', 'duration', 'delay']],
  'shape group': ['PowerPoint.groupShapes', ['target', 'slide', 'shapeIds'], []],
  'shape ungroup': ['PowerPoint.ungroupShape', ['target', 'slide', 'shapeId'], []],
  'shape align': ['PowerPoint.alignShapes', ['target', 'slide', 'shapeIds', 'alignment'], ['relativeToSlide']],
  'shape distribute': ['PowerPoint.distributeShapes', ['target', 'slide', 'shapeIds', 'direction'], ['relativeToSlide']],
  'textbox add': ['PowerPoint.addTextbox', ['target', 'slide', ...geometry], ['orientation']],
  'line add': ['PowerPoint.addLine', ['target', 'slide', 'beginX', 'beginY', 'endX', 'endY'], []],
  'freeform add': ['PowerPoint.addFreeform', ['target', 'slide', 'points'], []],
  'connector add': ['PowerPoint.addConnector', ['target', 'slide', 'connectorType', 'beginX', 'beginY', 'endX', 'endY'], []],
  'connector connect': ['PowerPoint.connectConnector', ['target', 'slide', 'shapeId'],
    ['beginShapeId', 'beginSite', 'endShapeId', 'endSite']],
  'picture add': ['PowerPoint.addPicture', ['target', 'slide', 'path', 'left', 'top'], ['width', 'height']],
  'table add': ['PowerPoint.addTable', ['target', 'slide', 'rows', 'columns', ...geometry], []],
  'table cell': ['PowerPoint.setTableCell', ['target', 'slide', 'shapeId', 'row', 'column'],
    ['text', 'fillColor', 'fontColor', 'fontSize', 'bold']],
  'chart add': ['PowerPoint.addChart', ['target', 'slide', 'chartType', ...geometry], ['style']],
  'chart data': ['PowerPoint.setChartData', ['target', 'slide', 'shapeId', 'categories', 'series'], []],
  'chart title': ['PowerPoint.setChartTitle', ['target', 'slide', 'shapeId'], ['text', 'visible']],
  'smartart layouts': ['PowerPoint.getSmartArtLayouts', [], []],
  'smartart add': ['PowerPoint.addSmartArt', ['target', 'slide', 'layout', ...geometry], []],
  'smartart node': ['PowerPoint.setSmartArtNode', ['target', 'slide', 'shapeId', 'text'], ['index', 'add']],
  'view state': ['PowerPoint.getViewState', ['target'], []],
  'slideshow state': ['PowerPoint.getSlideShowState', ['target'], []],
  'slideshow start': ['PowerPoint.startSlideShow', ['target'], []],
  'slideshow stop': ['PowerPoint.stopSlideShow', ['target'], []],
  'slideshow next': ['PowerPoint.navigateSlideShow', ['target'], []],
  'slideshow previous': ['PowerPoint.navigateSlideShow', ['target'], []],
  'slideshow goto': ['PowerPoint.navigateSlideShow', ['target', 'slideId'], []]
};

// Commands handled without a single WebSocket method.
const special = ['powerpoint launch', 'presentation list', 'batch run'];

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
    Read the deck's name, path, saved/read-only status, slide count, and slide size.
  netoffice presentation save --target <id> [--path <file>]
    Save the deck; with --path save a copy as .pptx, .pptm, .potx, .ppsx, or .pdf and continue there.
  netoffice presentation size --target <id> --width <pt> --height <pt>
    Set the slide size (960x540 is 16:9, 720x540 is 4:3).
  netoffice presentation theme --target <id> --path <file>
    Apply a theme or template (.thmx, .potx, .pot, .pptx).
  netoffice presentation colors --target <id> [--dark1 <color>] [--light1 <color>] [--dark2 <color>] [--light2 <color>]
      [--accent1 <color>] ... [--accent6 <color>] [--hyperlink <color>] [--followed-hyperlink <color>]
    Change theme color slots; returns all twelve slots.
  netoffice presentation fonts --target <id> [--major <font>] [--minor <font>]
    Change the theme heading (major) and body (minor) Latin fonts.
  netoffice presentation close --target <id> [--force]
    Close only that deck; refuse unsaved changes unless --force discards them.

Masters and layouts:
  netoffice master show --target <id>
    Read the slide master and its shapes; edit them with --master in place of --slide-id.
  netoffice layout list --target <id>
    List the master's custom layouts by index and name.
  netoffice layout show --target <id> --custom-layout <index>
    Read one custom layout and its shapes; edit them with --custom-layout in place of --slide-id.

Slides:
  netoffice slide list --target <id>
    List slides in order with stable slide IDs.
  netoffice slide show --target <id> --slide-id <id>
    Read one slide and its shapes (groups include groupItems); does not display it.
  netoffice slide add --target <id> [--layout <1..12> | --custom-layout <index> | --custom-layout-name <name>]
    Append a slide with a legacy layout (default 12: blank) or a master custom layout by index or exact name;
    return its slide ID.
  netoffice slide current --target <id> --slide-id <id>
    Select a slide in the editing window; does not navigate a running slideshow.
  netoffice slide delete --target <id> --slide-id <id>
    Delete the identified slide from the open presentation.
  netoffice slide move --target <id> --slide-id <id> --index <n>
    Move the slide to a 1-based position.
  netoffice slide name --target <id> --slide-id <id> --name <name>
    Rename the slide.
  netoffice slide notes --target <id> --slide-id <id> --text <text>
    Replace the speaker notes.
  netoffice slide transition --target <id> --slide-id <id> [--effect <PpEntryEffect>] [--duration <s>]
      [--advance-on-click true|false] [--advance-after <s>]
    Set the slide transition (e.g. 3844 fade smoothly, 0 none) and advance timing.
  netoffice slide hidden --target <id> --slide-id <id> --hidden true|false
    Hide or show the slide in slideshows.
  netoffice slide duplicate --target <id> --slide-id <id>
    Duplicate the slide after itself; return duplicateSlideId.
  netoffice slide export --target <id> --slide-id <id> --path <file> [--width-px <n>] [--height-px <n>]
    Render the slide to .png, .jpg, .gif, .bmp, .tif, or .svg.
  netoffice slide background --target <id> (--slide-id <id> | --master) [--color <color>] [--color2 <color>]
      [--gradient-style <style>] [--transparency <0..1>] [--picture <file>] [--follow-master true|false]
    Set a solid, two-color gradient, or picture background, or follow the master again.
  netoffice slide footer --target <id> (--slide-id <id> | --master) [--text <text>] [--footer true|false]
      [--slide-number true|false] [--date true|false] [--date-text <text>]
    Set the footer text and the footer, slide number, and date visibility.

Shapes (every shape command accepts --master or --custom-layout <index> in place of --slide-id,
except shape animation):
  netoffice shape add --target <id> --slide-id <id> --type <auto-shape-type> --left <pt> --top <pt> --width <pt> --height <pt>
    Add an auto-shape at the given point coordinates and size; return its shape ID.
  netoffice shape text --target <id> --slide-id <id> --shape-id <id> --text <text>
    Replace all plain text in that shape; separate from shape creation.
  netoffice shape delete --target <id> --slide-id <id> --shape-id <id>
    Delete the identified shape from its slide.
  netoffice shape state --target <id> --slide-id <id> --shape-id <id>
    Read the shape's geometry, fill, line, text frame, paragraphs, and formatted text runs.
  netoffice shape fill ... --shape-id <id> [--visible true|false] [--color <color>] [--color2 <color>]
      [--gradient-style <style>] [--transparency <0..1>] [--picture <file>]
    Set a solid, gradient, or picture fill, or remove it with --visible false.
  netoffice shape line ... --shape-id <id> [--visible true|false] [--color <color>] [--weight <pt>] [--dash <style>] [--transparency <0..1>]
    Set the outline, or remove it with --visible false.
  netoffice shape shadow ... --shape-id <id> [--visible true|false] [--color <color>] [--blur <pt>]
      [--offset-x <pt>] [--offset-y <pt>] [--transparency <0..1>] [--size <percent>]
  netoffice shape glow ... --shape-id <id> --radius <pt> [--color <color>] [--transparency <0..1>]
  netoffice shape softedge ... --shape-id <id> --radius <pt>
  netoffice shape 3d ... --shape-id <id> [--bevel <MsoBevelType>] [--bevel-depth <pt>] [--bevel-inset <pt>]
      [--depth <pt>] [--material <MsoPresetMaterial>]
  netoffice shape style ... --shape-id <id> --style <1..42>
    Apply a built-in shape style preset.
  netoffice shape adjust ... --shape-id <id> --index <n> --value <number>
    Set an adjustment handle (corner radius, chevron depth, ...).
  netoffice shape rotation ... --shape-id <id> --degrees <deg>
  netoffice shape flip ... --shape-id <id> --direction horizontal|vertical
  netoffice shape name ... --shape-id <id> --name <name>
  netoffice shape bounds ... --shape-id <id> [--left <pt>] [--top <pt>] [--width <pt>] [--height <pt>]
    Move or resize the shape, keeping its ID.
  netoffice shape zorder ... --shape-id <id> --order front|back|forward|backward
  netoffice shape duplicate ... --shape-id <id> [--left <pt>] [--top <pt>]
    Duplicate the shape; return duplicateShapeId.
  netoffice shape copy-format ... --shape-id <id> --source-shape-id <id>
    Copy the source shape's formatting onto this shape.
  netoffice shape font ... --shape-id <id> [--start <n> --length <n>] [--name <font>] [--size <pt>]
      [--bold true|false] [--italic true|false] [--underline true|false] [--color <color>]
    Format all text, or the characters from 1-based --start.
  netoffice shape paragraph ... --shape-id <id> [--paragraph <n>] [--alignment left|center|right|justify]
      [--space-before <pt>] [--space-after <pt>] [--line-spacing <lines>] [--indent-level <1..9>]
      [--bullet none|bullet|number] [--bullet-character <char>]
    Format every paragraph, or only the 1-based --paragraph.
  netoffice shape textframe ... --shape-id <id> [--anchor top|middle|bottom] [--margin-left <pt>] [--margin-right <pt>]
      [--margin-top <pt>] [--margin-bottom <pt>] [--autosize none|shape-to-fit-text] [--word-wrap true|false]
  netoffice shape animation --target <id> --slide-id <id> --shape-id <id> --effect <MsoAnimEffect>
      [--trigger on-click|with-previous|after-previous] [--duration <s>] [--delay <s>]
    Append an entrance or emphasis effect (e.g. 10 fade, 2 appear) to the slide's animation sequence.
  netoffice shape group ... --shape-ids <id,id,...>
    Group the shapes; return the group's shape ID.
  netoffice shape ungroup ... --shape-id <id>
    Ungroup; return the member shape IDs.
  netoffice shape align ... --shape-ids <id,id,...> --alignment left|center|right|top|middle|bottom [--relative-to-slide true|false]
  netoffice shape distribute ... --shape-ids <id,id,...> --direction horizontal|vertical [--relative-to-slide true|false]

Content (each add accepts --slide-id <id>, --master, or --custom-layout <index> and returns a shape ID):
  netoffice textbox add --target <id> --slide-id <id> --left <pt> --top <pt> --width <pt> --height <pt> [--orientation horizontal|upward|downward|vertical]
  netoffice line add --target <id> --slide-id <id> --begin-x <pt> --begin-y <pt> --end-x <pt> --end-y <pt>
  netoffice freeform add --target <id> --slide-id <id> --points "<x>,<y>;<x>,<y>;..."
    Draw a freeform through the points (pt, at least two distinct) joined by straight segments, in order;
    repeat the first point last to close it. Returns the shape ID.
  netoffice connector add --target <id> --slide-id <id> --connector-type straight|elbow|curve --begin-x <pt> --begin-y <pt> --end-x <pt> --end-y <pt>
  netoffice connector connect --target <id> --slide-id <id> --shape-id <connector> [--begin-shape-id <id> [--begin-site <n>]] [--end-shape-id <id> [--end-site <n>]]
  netoffice picture add --target <id> --slide-id <id> --path <image> --left <pt> --top <pt> [--width <pt> --height <pt>]
    Insert a picture or SVG icon, embedded in the deck.
  netoffice table add --target <id> --slide-id <id> --rows <n> --columns <n> --left <pt> --top <pt> --width <pt> --height <pt>
  netoffice table cell ... --shape-id <table> --row <n> --column <n> [--text <text>] [--fill-color <color>]
      [--font-color <color>] [--font-size <pt>] [--bold true|false]
  netoffice chart add --target <id> --slide-id <id> --chart-type <XlChartType> --left <pt> --top <pt> --width <pt> --height <pt> [--style <n>]
    Add a chart (e.g. 51 clustered column, 4 line, 5 pie, 57 clustered bar).
  netoffice chart data ... --shape-id <chart> --categories <a,b,...> --series "<name>:<v,v,...>;<name>:<v,v,...>"
  netoffice chart title ... --shape-id <chart> [--text <text>] [--visible true|false]
  netoffice smartart layouts
    List SmartArt layouts by index, ID, name, and category; needs no target.
  netoffice smartart add --target <id> --slide-id <id> --layout <index> --left <pt> --top <pt> --width <pt> --height <pt>
  netoffice smartart node ... --shape-id <smartart> (--index <n> | --add) --text <text>
    Replace a node's text, or append a node with that text.

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

Batches:
  netoffice batch run --file <commands.json>
    Run a JSON array of commands, each an array of CLI arguments (["shape","add","--target",...]),
    over one connection in order, stopping at the first failure. An argument "$<n>.<field>"
    is replaced by that field of the result of earlier command n (0-based). Prints all results.

Options:
  --path <file>     Presentation, theme, picture, or export file; relative paths resolve against the current directory.
  --layout <n>      Slide layout for slide add (1..12, default 12: blank) or SmartArt layout index.
  --type <number>   PowerPoint MsoAutoShapeType (1..255) for shape add.
  --text <text>     Replacement plain text.
  --force           Shutdown: discard all decks; presentation close: discard only its target.
  --master          Address the slide master instead of --slide-id.
  --custom-layout <index>  On slide add: the new slide's layout. Elsewhere: address that custom
                    layout (index from layout list) instead of --slide-id.
  --port <port>     Add-in port (default 50051); accepted by every command.
  --timeout <ms>    Total deadline (default 10000); accepted by every command.

IDs, units, and values:
  --target <id>     Presentation ID from presentation list or presentation new.
  --slide-id <id>   Stable SlideID from slide list or slide add; not a slide index.
  --shape-id <id>   Shape ID from shape add, any content add, or slide show.
  --left <pt>, --top <pt>       Shape position in points.
  --width <pt>, --height <pt>   Positive dimensions in points.
  <color>           #RRGGBB.
  <style> (gradient)            horizontal, vertical, diagonal-up, diagonal-down, from-corner, from-center.
  <style> (dash)                solid, square-dot, round-dot, dash, dash-dot, dash-dot-dot, long-dash, long-dash-dot.

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

function parseValue(type, value, option) {
  switch (type) {
    case 'target':
      if (!/^[a-fA-F0-9]{8}(?:-[a-fA-F0-9]{4}){3}-[a-fA-F0-9]{12}$/.test(value)) {
        throw new Error('--target requires a canonical GUID from presentation list.');
      }
      return value;
    case 'id': return positiveInteger(value, option, 2147483647);
    case 'integer':
      if (!/^-?[0-9]+$/.test(value) || !Number.isSafeInteger(Number(value))) throw new Error(`${option} requires an integer.`);
      return Number(value);
    case 'number': case 'positive': return finiteNumber(value, option);
    case 'text': return value;
    case 'name': case 'file':
      if (!value) throw new Error(`${option} requires a non-empty value.`);
      return value;
    case 'absolute':
      if (!value) throw new Error(`${option} requires a non-empty file path.`);
      return path.resolve(value);
    case 'boolean':
      if (value !== 'true' && value !== 'false') throw new Error(`${option} requires true or false.`);
      return value === 'true';
    case 'color':
      if (!/^#[0-9a-fA-F]{6}$/.test(value)) throw new Error(`${option} requires a #RRGGBB color.`);
      return value.toUpperCase();
    case 'ids':
      return value.split(',').map(part => positiveInteger(part.trim(), option, 2147483647));
    case 'list':
      if (!value) throw new Error(`${option} requires a comma-separated list.`);
      return value.split(',').map(part => part.trim());
    case 'points': {
      const points = value.split(';').map((entry, index) => {
        const parts = entry.split(',');
        if (parts.length !== 2) throw new Error(`${option} requires "<x>,<y>" pairs separated by ";" (pair ${index + 1}).`);
        return { x: finiteNumber(parts[0].trim(), option), y: finiteNumber(parts[1].trim(), option) };
      });
      if (points.length < 2) throw new Error(`${option} requires at least two points.`);
      return points;
    }
    case 'series':
      return value.split(';').map(entry => {
        const separator = entry.lastIndexOf(':');
        if (separator < 1) throw new Error(`${option} requires "<name>:<v,v,...>" entries separated by ";".`);
        return {
          name: entry.slice(0, separator).trim(),
          values: entry.slice(separator + 1).split(',').map(part => finiteNumber(part.trim(), option))
        };
      });
    default: throw new Error(`Unsupported option type for ${option}.`);
  }
}

const flagName = field => `--${field.replace(/[A-Z]/g, letter => `-${letter.toLowerCase()}`)}`;

function parseArguments(args) {
  if (args.length === 1 && (args[0] === '--help' || args[0] === '-h')) return { help: true };
  const command = `${args[0] || ''} ${args[1] || ''}`;
  const spec = commands[command];
  if (!spec && !special.includes(command)) {
    throw new Error('Unknown command. Use --help for usage.');
  }
  const [method, required, optional] = spec || [undefined, command === 'batch run' ? ['file'] : [], []];
  const accepted = new Set([...required, ...optional]);
  if (accepted.has('slide')) { accepted.add('slideId'); accepted.add('master'); accepted.add('customLayout'); }
  const result = { command, timeout: 10000, port: 50051 };
  const seen = new Set();
  for (let index = 2; index < args.length; index++) {
    const arg = args[index];
    if (arg === '--help' || arg === '-h') { result.help = true; continue; }
    if (!arg.startsWith('-')) throw new Error(`Unexpected argument: ${arg}`);
    const option = options[arg];
    const allowed = arg === '--timeout' || arg === '--port' || (option && accepted.has(option[0]));
    if (!allowed) throw new Error(`Unknown option: ${arg}`);
    if (seen.has(arg)) throw new Error(`Repeated option: ${arg}`);
    seen.add(arg);
    if (option && option[1] === 'flag') { result[option[0]] = true; continue; }
    if (++index >= args.length || args[index].startsWith('--')) throw new Error(`Missing value for ${arg}.`);
    const value = args[index];
    if (arg === '--port') result.port = positiveInteger(value, arg, 65535);
    else if (arg === '--timeout') result.timeout = positiveInteger(value, arg, 2147483647);
    else if (option[0] === 'path' && !value) throw new Error('--path requires a non-empty file path.');
    else if (arg === '--layout' && command === 'slide add') result.layout = positiveInteger(value, arg, 12);
    else if (arg === '--type') result.shapeType = positiveInteger(value, arg, 255);
    else result[option[0]] = parseValue(option[1], value, arg);
  }
  if (result.help) return result;
  for (const field of required) {
    if (field === 'slide') {
      const selectors = [Object.hasOwn(result, 'slideId'), result.master === true, Object.hasOwn(result, 'customLayout')];
      if (selectors.filter(Boolean).length !== 1) {
        throw new Error(`${command} requires exactly one of --slide-id, --master, or --custom-layout.`);
      }
    } else if (!Object.hasOwn(result, field)) throw new Error(`${command} requires ${flagName(field)}.`);
  }
  if ((Object.hasOwn(result, 'width') && result.width <= 0) ||
      (Object.hasOwn(result, 'height') && result.height <= 0)) {
    throw new Error('--width and --height must be positive.');
  }
  if (command === 'slide add' &&
      ['layout', 'customLayout', 'customLayoutName'].filter(field => Object.hasOwn(result, field)).length > 1) {
    throw new Error('slide add accepts only one of --layout, --custom-layout, or --custom-layout-name.');
  }
  if (command === 'smartart node' && Object.hasOwn(result, 'index') === (result.add === true)) {
    throw new Error('smartart node requires exactly one of --index or --add.');
  }
  if (!method) return result;
  const params = {};
  for (const field of accepted) {
    if (!Object.hasOwn(result, field)) continue;
    if (themeColors.includes(field) && command === 'presentation colors') {
      params.colors = { ...params.colors, [field]: result[field] };
    } else if (field === 'target') params.targetId = result.target;
    else if (field === 'path' && command === 'presentation new') {
      const destination = path.parse(path.resolve(result.path));
      params.name = destination.base;
      params.directory = destination.dir;
    } else if (field === 'path' && command !== 'presentation open') params.path = path.resolve(result.path);
    else params[field] = result[field];
  }
  if (command.startsWith('slideshow ') && ['next', 'previous', 'goto'].includes(args[1])) {
    params.action = args[1];
  }
  result.method = method;
  result.params = params;
  return result;
}

module.exports = { parseArguments, usage };
