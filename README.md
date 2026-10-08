# NetOffice Automate

> NetOffice command line tool for running automation tests of Microsoft Office applications.

## PowerPoint addin

`src/addin` builds the native `NetOffice.Automate` COM addin
(`{6d274715-3f05-4505-aa63-2ae9df1e6881}`). Its `IDTExtensibility2`
connection starts an embedded [CivetWeb](https://github.com/civetweb/civetweb)
HTTP/WebSocket server on IPv4 loopback (default `127.0.0.1:50051`). CivetWeb
owns HTTP parsing, WebSocket framing, handshakes, and socket I/O; the addin
provides JSON handlers. Document commands are queued onto PowerPoint's owning
STA; network and command worker threads never call Office COM.
`Connect.h` and `Connect.cpp` follow the native ATL addin pattern:
`IDispatchImpl` uses Office's imported extensibility interface, and ATL manages
COM lifetime and the class factory.

Build from a Visual Studio developer shell with the v145 C++ toolset,
Windows SDK, ATL, and Office's `MSADDNDR.OLB` installed. The project searches
the 32-bit Click-to-Run Office `DESIGNER` directory by default. For a different
Office layout, pass `/p:OfficeExtensibilityDir="path to the directory containing MSADDNDR.OLB"`.
Install [vcpkg](https://github.com/microsoft/vcpkg) and run
`vcpkg integrate install` once. The manifest pins header-only nlohmann-json
through a registry baseline and CivetWeb 1.16 through a minimal overlay.
MSBuild installs the static C library with the dynamic CRT. TLS, compression,
scripting, CGI, static-file serving, and the C++ wrapper are disabled.
The overlay adds a pre-allocation 1 MiB WebSocket frame limit; the addin also
limits complete fragmented messages and HTTP JSON bodies to 1 MiB.

Measured Release DLL sizes with this configuration: Win32 **342,528 bytes**
(334 KiB), x64 **380,928 bytes** (372 KiB). CivetWeb is linked statically;
there is no separate CivetWeb, gRPC, protobuf, OpenSSL, or zlib runtime DLL.

```powershell
msbuild src/addin/addin.vcxproj /p:Configuration=Release /p:Platform=Win32
```

Build the active Visual Studio configuration once to generate `MSADDNDR.tlh`.
IntelliSense reads that generated header from the configuration's intermediate
directory; normal compilation imports `MSADDNDR.OLB`.

Use `Win32` for 32-bit PowerPoint or `x64` for 64-bit PowerPoint.
Close PowerPoint before changing registration. From this directory, using
PowerShell Core 7 on Windows:

```powershell
pwsh -NoProfile -File src/addin/Register-Addin.ps1 -Action Register -Architecture x86 -DllPath src/addin/build/Release_Win32/addin.dll -Port 50051
pwsh -NoProfile -File src/addin/Register-Addin.ps1 -Action Unregister -Architecture x86
```

For 64-bit PowerPoint, use `-Architecture x64` and
`src/addin/build/Release_x64/addin.dll`. DLL architecture is validated before
registration; PowerShell's own bitness does not select the registry view.
Registration is per-user (HKCU), requires no elevation, and sets PowerPoint's
`LoadBehavior` to DWORD `3` for startup activation. The DLL must remain at its
registered absolute path. Release builds require the matching Microsoft Visual
C++ runtime. Deregistration also works after the DLL has been removed.
Neither `DllRegisterServer` nor `DllUnregisterServer` is exported; use the script,
not `regsvr32`.

Registration and server startup require no administrator rights, UAC prompt,
HTTP.sys URL reservation, or firewall rule. `-Port` accepts `1..65535` and stores
the DWORD `ServerPort` under the PowerPoint addin's HKCU registration key.
Missing `ServerPort` defaults to `50051`; an invalid registry value fails startup.
Use a free port and restart PowerPoint after changing it. Only one addin connection
can own the selected endpoint; the server never falls back to a different port.

Rebuild the configuration whose DLL path is actually registered; rebuilding
Release does not update a registered Debug DLL. Close PowerPoint before replacing
the loaded DLL.

## NodeJS CLI

Requires Node.js 20 or newer. After building and registering the addin, install
and link the CLI from this directory:

```powershell
npm --prefix src/cli ci
Push-Location src/cli
npm link
Pop-Location
netoffice --help
netoffice powerpoint launch
netoffice presentation new --path ".\Quarterly report.pptx"
```

Without linking, use `node src/cli/bin/netoffice.js` with the same arguments.
`powerpoint launch` starts visible desktop PowerPoint on Windows and waits for
the addin's WebSocket JSON readiness response. An already-ready session is reused.
The document commands require that session to be running; they do not launch it
implicitly or fall back to external COM automation.
`powerpoint shutdown` validates that all open presentations are saved, then
requests `Application.Quit` from a separate process. Unsaved changes or a deck
without a saved path block the default command (`-32005`). Use
`netoffice powerpoint shutdown --force` to **discard changes in every open
presentation**, close them without saving, and quit PowerPoint. This does not
terminate a hung process; the add-in and Office must be responsive. The CLI
reports success only after the PowerPoint process exits.

`presentation new --path <file>` creates a blank presentation at a relative
(resolved against the CLI's current directory) or absolute path. The `.pptx`
suffix is appended if omitted; the parent directory must already exist and the
destination must not. `presentation open --path <file>` opens an existing deck
at a relative or absolute path without saving it. Both return a target ID.
Create a slide and edit its shapes with the independent commands below.

The lower-level `PowerPoint.newPresentation` WebSocket method and `POST /json/new`
still accept separate `name` and absolute `directory` fields; the CLI resolves
and splits `--path` before calling the WebSocket method.

Every new read and mutation is also a `netoffice` subcommand. Use
`netoffice presentation list` to find the target ID, then
`netoffice slide list --target <id>` to find stable slide IDs. New commands output JSON so scripts can
capture IDs with `ConvertFrom-Json`:

```powershell
$target = (netoffice presentation new --path ".\Quarterly report.pptx" | ConvertFrom-Json).id
netoffice presentation state --target $target
netoffice slide list --target $target
netoffice view state --target $target
netoffice slideshow state --target $target
$newSlide = netoffice slide add --target $target --layout 12 | ConvertFrom-Json
netoffice slide show --target $target --slide-id $newSlide.slideId
$shape = netoffice shape add --target $target --slide-id $newSlide.slideId `
  --type 1 --left 24 --top 36 --width 160 --height 48 | ConvertFrom-Json
netoffice shape text --target $target --slide-id $newSlide.slideId --shape-id $shape.shapeId --text "Updated"
netoffice slide current --target $target --slide-id $newSlide.slideId
netoffice slideshow start --target $target
netoffice slideshow goto --target $target --slide-id $newSlide.slideId
netoffice slideshow previous --target $target
netoffice slideshow next --target $target
netoffice slideshow stop --target $target
netoffice shape delete --target $target --slide-id $newSlide.slideId --shape-id $shape.shapeId
netoffice slide delete --target $target --slide-id $newSlide.slideId
```

`--type` is PowerPoint's numeric `MsoAutoShapeType` for `Shapes.AddShape`;
the returned read-only `shapeType` describes `Shape.Type` (`MsoShapeType`).
Target-scoped commands require `--target`; `netoffice --help` lists every
command's options, ID sources, point units, and all twelve slide layouts.
`presentation list` reads `/json/list`; the other new commands use the matching
WebSocket methods.
`presentation close --target <id>` closes only that presentation and leaves
PowerPoint and other decks open. Unsaved edits cause a conflict unless `--force`
is supplied, which discards only the target's unsaved edits.


All commands accept `--port <port>` (default `50051`) and
`--timeout <milliseconds>` (default `10000`, total operation deadline).
`--port` must match the registered `ServerPort`; `powerpoint launch` does not
reconfigure the addin. For example, register with `-Port 50123` and invoke
`netoffice powerpoint launch --port 50123`.

Failures exit nonzero with a readable protocol error and optional HRESULT details.
Only connection/readiness probes are retried; document mutations are never retried.
WebSocket timeout or disconnect cancels queued work. A COM operation already
started may complete: inspect PowerPoint before repeating a mutation.
Disconnecting the addin or closing PowerPoint stops the server and wakes pending
requests before joining workers.

`npm --prefix src/cli pack` creates a standalone npm package containing the
CLI and JSON connection code. Its only runtime dependency is `ws`; no protobuf
schema or code-generation step is needed.

## JSON protocol

The HTTP interface is based on the
[Chrome DevTools Protocol discovery interface](https://chromedevtools.github.io/devtools-protocol/),
with document targets rather than browser tabs. This is not a Chromium-compatible
debugging implementation. WebSocket request/response envelopes follow CDP's shape;
there is no `jsonrpc` member.

| Endpoint | Behavior |
| --- | --- |
| `GET /json/version` | Office application/protocol/runtime metadata and application WebSocket URL |
| `GET /json/list` or `/json` | Targets for all currently open presentation documents; empty array when none are open |
| `PUT /json/new?url=<encoded-path-or-link>` | Open a local file or OneDrive/SharePoint link; absent or empty `url` creates a blank presentation |
| `POST /json/new` with `{"name":"Report","directory":"C:\\Documents"}` | Create and save a named blank presentation; reject existing destination (HTTP 409) |
| `PUT /json/activate/{target}` | Activate the document and bring its window to the foreground |
| `PUT /json/close/{target}` | Close a saved document; HTTP 409 when it has unsaved changes |
| `PUT /json/close/{target}?force` | Discard changes and close the document without saving |
| `GET /json/{target}/presentation` | Targeted presentation name/path, saved state, read-only state, and slide count |
| `GET /json/{target}/slides` | Ordered slides with stable slide IDs, indices, names, and hidden state |
| `GET /json/{target}/slides/{slide-id}` | Slide metadata and shape IDs, z-order, type, bounds, and readable plain text |
| `GET /json/{target}/view` | Editing-window view and selection state, when available |
| `GET /json/{target}/slide-show` | Running slide-show windows and each current slide |
| `POST /json/{target}/slides` | Create a slide (optional PowerPoint layout) |
| `POST /json/{target}/slides/{slide-id}/shapes` | Create an auto-shape with bounds; use the returned shape ID for a separate text update |
| `PUT /json/{target}/slides/{slide-id}/shapes/{shape-id}` | Replace a shape's plain text |
| `DELETE /json/{target}/slides/{slide-id}/shapes/{shape-id}` | Delete a shape |
| `DELETE /json/{target}/slides/{slide-id}` | Delete a slide |
| `PUT /json/{target}/view` | Set the target presentation's current editing slide |
| `POST /json/{target}/slide-show` / `DELETE /json/{target}/slide-show` | Start / stop its slide show |
| `POST /json/{target}/slide-show/navigation` | Navigate next, previous, or to a slide during a show |
| `POST /json/rpc` | One JSON request and response; requires `Content-Type: application/json` |
| `WS /devtools/application` | Application-wide UTF-8 JSON requests and correlated responses |

HTTP responses use `application/json; charset=utf-8`. WebSocket JSON uses text
messages rather than an HTTP content type. With the default port, connect to
`ws://127.0.0.1:50051/devtools/application`.

`/json/protocol` is intentionally unavailable (HTTP 404).

### Application metadata and document targets

`/json/version` uses `Application` instead of CDP's `Browser` field. Its value
contains the actual Office application name, version, build, and process architecture.
`V8-Version` is `null` until QuickJS is integrated; no engine version is fabricated.
For example:

```json
{
  "Protocol-Version": "1.0",
  "Application": "Microsoft PowerPoint/16.0 (build 20527; x86)",
  "V8-Version": null,
  "webSocketDebuggerUrl": "ws://127.0.0.1:50051/devtools/application"
}
```

Each `/json/list` entry and successful `/json/new` response is a document descriptor:

```json
{
  "id": "d17fd90d-90f8-4a29-b204-49495b78d064",
  "type": "document",
  "title": "Quarterly report.pptx",
  "url": "C:\\Documents\\Quarterly report.pptx"
}
```

Target IDs are opaque GUIDs, stable while the document remains open in the same
addin connection, including after Save As. Closed/reopened documents get new IDs.
Documents opened or closed outside the HTTP API are reflected on the next list.
Never identify documents by a collection index or title. `url` is the Office
document's full location, or empty for a never-saved document.

The target concept is application-neutral: a presentation in PowerPoint, document
in Word, or workbook in Excel. Only the PowerPoint adapter is implemented here.
The WebSocket is application-wide; document descriptors do not advertise
unimplemented document-specific sockets.

### Creating, opening, activating, and closing

Percent-encode the entire `url` query parameter, including spaces, Unicode,
`+`, `#`, `&`, and any query parameters in a cloud link. The decoded value is
passed unchanged to PowerPoint's `Presentations.Open`; Office owns file loading
and OneDrive/SharePoint authentication. The addin does not download or rewrite links.
Cloud opening requires an accessible link and the appropriate Office sign-in.

```powershell
$base = 'http://127.0.0.1:50051'
$blank = Invoke-RestMethod -Method Put -Uri "$base/json/new"
$url = [Uri]::EscapeDataString('C:\Documents\Quarterly report.pptx')
$opened = Invoke-RestMethod -Method Put -Uri "$base/json/new?url=$url"
Invoke-RestMethod -Method Put -Uri "$base/json/activate/$($opened.id)"
Invoke-RestMethod -Method Put -Uri "$base/json/close/$($opened.id)"
# Explicitly discard unsaved changes:
Invoke-RestMethod -Method Put -Uri "$base/json/close/$($blank.id)?force"
```

`PUT /json/new` creates Office's ordinary unsaved blank presentation; `POST`
creates and saves a named blank `.pptx`. No slide or title is added implicitly.
Activate and close return HTTP 200 with `{}`. The API rejects unsaved close with
HTTP 409 without modifying the document's `Saved` state. The `force` parameter
is a presence flag: `?force`, `?force=1`, and `?force=false` all authorize discard.
Apart from named presentation creation, no operation automatically saves a document.

Malformed queries, duplicate `url`/`force` parameters, and invalid target IDs
return HTTP 400. Unknown or closed targets return HTTP 404; there is no fallback
to the active presentation. Unsupported verbs return HTTP 405 with `Allow`.
Deadline expiry is HTTP 504, server stopping is HTTP 503, and COM failures are HTTP 500.
Errors are JSON `{"error":{"code":...,"message":...,"data":...}}`, with optional
HRESULT details. HTTP target requests have a 10-second queue deadline; a started
Office call may outlive it. Inspect the document list before repeating a mutation.

### Reading presentation state

The new read routes are target-scoped; `{target}` is the opaque ID from `/json/list`.
Slide routes use PowerPoint's stable `SlideID`, not the current slide index.
Responses use camelCase JSON fields. Slide details contain top-level shape
summaries (`shapeId`, `zOrderPosition`, `name`, numeric `shapeType`, `bounds`,
plain `text` or `null`, and recursive `groupItems` for groups). Rich text is read
per shape with `shape state`; chart data and image contents are not represented.
View/selection fields are transient and may be
unavailable. `slide-show` reports live show windows and their current slides.
Each call is a best-effort COM read, not a transactionally consistent snapshot.

```powershell
$base = 'http://127.0.0.1:50051'
$target = 'd17fd90d-90f8-4a29-b204-49495b78d064'
$presentation = Invoke-RestMethod "$base/json/$target/presentation"
$slides = Invoke-RestMethod "$base/json/$target/slides"
$slideId = $slides.slides[0].slideId
$slide = Invoke-RestMethod "$base/json/$target/slides/$slideId"
$view = Invoke-RestMethod "$base/json/$target/view"
$show = Invoke-RestMethod "$base/json/$target/slide-show"
```

Matching application WebSocket methods are `PowerPoint.getPresentationState`,
`PowerPoint.getSlides`, `PowerPoint.getSlideState`, `PowerPoint.getViewState`,
and `PowerPoint.getSlideShowState`. Pass `{ "targetId": "<target>" }`; the
slide-detail method also requires `slideId`. HTTP and WebSocket return the same
result shape. Unknown targets or slide IDs fail rather than falling back to the
active presentation. The corresponding CLI commands print the same result JSON.

See [`docs/powerpoint-state-api.md`](docs/powerpoint-state-api.md) for full
response examples, enum notes, and COM state caveats.

### Mutating presentation state for tests

Mutation routes are target-scoped and operate on the live, unsaved document;
they never save automatically. Slide IDs are stable PowerPoint `SlideID`s.
Slide creation defaults to blank layout `12`; optional `layout` accepts `1..12`.
Shape creation requires `shapeType`, finite `left`/`top`, and positive
`width`/`height`. Use the returned `shapeId` with `shape text` to set or replace
plain text. Deleting slides/shapes changes the document in memory.

`PUT /json/{target}/view` accepts `{"slideId":259}`. Slideshow navigation
accepts `{"action":"next"}`, `{"action":"previous"}`, or
`{"action":"goto","slideId":259}`. Start/stop and navigation affect only show
windows associated with that target.

Matching WebSocket methods are `PowerPoint.addSlide`, `PowerPoint.createShape`,
`PowerPoint.setShapeText`, `PowerPoint.deleteShape`, `PowerPoint.deleteSlide`,
`PowerPoint.setCurrentSlide`, `PowerPoint.startSlideShow`,
`PowerPoint.stopSlideShow`, `PowerPoint.navigateSlideShow`, and
`PowerPoint.closePresentation`. Every method
requires `targetId`; slide and shape methods also require stable `slideId` and,
where applicable, `shapeId`. For disconnect-sensitive mutations, use WebSocket:
an abandoned HTTP call may still execute until its queue deadline.

### Design, formatting, and content commands

Each row of `src/addin/PowerPointCommands.h` is one granular command: its
WebSocket method, HTTP verb, scope, and route suffix. The scope fixes the route
prefix and the required IDs:

| Scope | Route | Required params |
| --- | --- | --- |
| Application | `/json/<suffix>` | none |
| Target | `/json/{target}/<suffix>` | `targetId` |
| Slide | `/json/{target}/slides/{slide-id}/<suffix>` | `targetId`, `slideId` |
| Slide, master, or layout | `…/slides/{slide-id}/<suffix>`, `/json/{target}/master/<suffix>`, or `/json/{target}/layouts/{layout-index}/<suffix>` | `targetId` and exactly one of `slideId`, `master: true`, or `customLayout` |
| Shape | `…/shapes/{shape-id}/<suffix>` below a slide, the master, or a layout | the above plus `shapeId` |

`customLayout` is the 1-based index from `layout list` (`SlideMaster.CustomLayouts`).
`shape add`, `shape text`, and `shape delete` also accept the master and layouts
(`…/master/shapes[/{shape-id}]`, `…/layouts/{layout-index}/shapes[/{shape-id}]`). Mutations use a JSON body (send
`{}` when there are no members); results contain the routing identity, any new
IDs, and the accepted parameters. Colors are `"#RRGGBB"`; booleans are JSON
booleans; paths are absolute. Setters reject requests that set nothing.
`presentation state` also reports `slideWidth`/`slideHeight`, slide details
report group members as `groupItems`, and `slide add` accepts `customLayout`
(an index from `layout list`) or `customLayoutName` (an exact layout name)
instead of `layout`; its result then includes the resolved `customLayout`.

| CLI | Method | HTTP |
| --- | --- | --- |
| `presentation save [--path]` | `savePresentation` | `POST …/presentation/save` |
| `presentation size` | `setSlideSize` | `PUT …/presentation/page-setup` |
| `presentation theme` | `applyTheme` | `PUT …/presentation/theme` |
| `presentation colors` | `setThemeColors` (`colors` object) | `PUT …/presentation/theme/colors` |
| `presentation fonts` | `setThemeFonts` | `PUT …/presentation/theme/fonts` |
| `layout list` | `getLayouts` | `GET …/layouts` |
| `master show` | `getMasterState` | `GET …/master` |
| `layout show --custom-layout` | `getLayoutState` | `GET …/layouts/{layout-index}` |
| `slide background` | `setBackground` | `PUT …/background` |
| `slide footer` | `setHeadersFooters` | `PUT …/headers-footers` |
| `slide move`, `name`, `notes`, `transition`, `hidden` | `moveSlide`, `setSlideName`, `setSlideNotes`, `setSlideTransition`, `setSlideHidden` | `PUT …/slides/{id}/position`, `name`, `notes`, `transition`, `hidden` |
| `slide duplicate`, `slide export` | `duplicateSlide`, `exportSlide` | `POST …/slides/{id}/duplicate`, `export` |
| `shape state` | `getShapeState` | `GET …/shapes/{id}` |
| `shape fill`, `line`, `shadow`, `glow`, `softedge`, `3d`, `style`, `adjust`, `rotation`, `name`, `bounds`, `zorder` | `setShapeFill`, `setShapeLine`, `setShapeShadow`, `setShapeGlow`, `setShapeSoftEdge`, `setShapeThreeD`, `setShapeStyle`, `setShapeAdjustment`, `setShapeRotation`, `setShapeName`, `setShapeBounds`, `setShapeZOrder` | `PUT …/shapes/{id}/fill`, `line`, `shadow`, `glow`, `soft-edge`, `three-d`, `style`, `adjustments`, `rotation`, `name`, `bounds`, `z-order` |
| `shape flip`, `duplicate`, `copy-format`, `ungroup`, `animation` | `flipShape`, `duplicateShape`, `copyShapeFormat`, `ungroupShape`, `addAnimation` | `POST …/shapes/{id}/flip`, `duplicate`, `format`, `ungroup`, `animations` |
| `shape font`, `paragraph`, `textframe` | `setShapeFont`, `setShapeParagraph`, `setShapeTextFrame` | `PUT …/shapes/{id}/text/font`, `text/paragraphs`, `text/frame` |
| `shape group`, `align`, `distribute` | `groupShapes`, `alignShapes`, `distributeShapes` | `POST …/groups`, `alignment`, `distribution` |
| `textbox add`, `line add`, `freeform add`, `connector add`, `picture add`, `table add`, `chart add`, `smartart add` | `addTextbox`, `addLine`, `addFreeform`, `addConnector`, `addPicture`, `addTable`, `addChart`, `addSmartArt` | `POST …/textboxes`, `lines`, `freeforms`, `connectors`, `pictures`, `tables`, `charts`, `smartart` |
| `connector connect`, `table cell`, `chart data`, `chart title`, `smartart node` | `connectConnector`, `setTableCell`, `setChartData`, `setChartTitle`, `setSmartArtNode` | `PUT …/shapes/{id}/connections`, `cells`, `chart/data`, `chart/title`, `smartart/nodes` |
| `smartart layouts` | `getSmartArtLayouts` | `GET /json/smartart-layouts` |

`freeform add` / `addFreeform` draws one freeform (`Shapes.BuildFreeform` →
`FreeformBuilder.AddNodes` → `ConvertToShape`) through `points`, a JSON array of
2–10000 `{"x": <pt>, "y": <pt>}` objects (−10000…10000, at least two distinct)
joined in order by straight segments. Repeat the first point last to close the
shape. The CLI takes `--points "x,y;x,y;…"`. The result carries `shapeId`, `name`,
and the accepted `points`.

`netoffice --help` lists every option, value, and enum name. For example:

```powershell
$t = (netoffice presentation open --path .\LifeInCorporation.pptx | ConvertFrom-Json).id
$s = (netoffice slide add --target $t | ConvertFrom-Json).slideId
$card = (netoffice shape add --target $t --slide-id $s --type 5 --left 66 --top 232 --width 195 --height 200 | ConvertFrom-Json).shapeId
netoffice shape fill --target $t --slide-id $s --shape-id $card --color '#FFFFFF'
netoffice shape shadow --target $t --slide-id $s --shape-id $card --blur 14 --offset-y 4 --transparency 0.85
netoffice shape text --target $t --slide-id $s --shape-id $card --text "Meetings`nStand-ups and syncs"
netoffice shape font --target $t --slide-id $s --shape-id $card --start 1 --length 8 --bold true --color '#1F3864'
netoffice slide export --target $t --slide-id $s --path .\preview.png
netoffice presentation save --target $t
```

`netoffice batch run --file steps.json` runs a JSON array of CLI argument
arrays over one connection, in order, and stops at the first failure. In any
argument, `$<n>.<field>` becomes that field of the result of step `n`
(0-based), e.g. `["shape","fill","--target","$0.id","--slide-id","$1.slideId","--shape-id","$2.shapeId","--color","#FFFFFF"]`.
The whole batch shares one `--timeout`; it prints every result, including the
completed ones before a failure. `powerpoint` commands and `presentation list`
cannot run in a batch.

### Application WebSocket commands


```json
{"id":1,"method":"PowerPoint.getStatus"}
{"id":1,"result":{"processId":1234}}

{"id":2,"method":"PowerPoint.newPresentation","params":{"name":"Quarterly report","directory":"C:\\Documents"},"timeoutMs":10000}
{"id":2,"result":{"id":"d17fd90d-90f8-4a29-b204-49495b78d064","name":"Quarterly report.pptx","url":"C:\\Documents\\Quarterly report.pptx","slideCount":0}}
{"id":12,"method":"PowerPoint.openPresentation","params":{"path":"C:\\Documents\\Existing.pptx"}}
{"id":12,"result":{"id":"a4d083ef-2e42-4c54-a690-f4885097017c","title":"Existing.pptx","url":"C:\\Documents\\Existing.pptx","type":"document"}}
{"id":5,"method":"PowerPoint.getPresentationState","params":{"targetId":"d17fd90d-90f8-4a29-b204-49495b78d064"}}
{"id":5,"result":{"id":"d17fd90d-90f8-4a29-b204-49495b78d064","name":"Quarterly report.pptx","url":"C:\\Documents\\Quarterly report.pptx","saved":true,"readOnly":false,"slideCount":0,"slideWidth":960,"slideHeight":540}}

{"id":3,"method":"PowerPoint.addSlide","params":{"targetId":"d17fd90d-90f8-4a29-b204-49495b78d064","layout":12}}
{"id":3,"result":{"id":"d17fd90d-90f8-4a29-b204-49495b78d064","slideId":256,"slideIndex":1}}
{"id":10,"method":"PowerPoint.createShape","params":{"targetId":"d17fd90d-90f8-4a29-b204-49495b78d064","slideId":256,"shapeType":1,"left":36,"top":24,"width":648,"height":54}}
{"id":10,"result":{"id":"d17fd90d-90f8-4a29-b204-49495b78d064","slideId":256,"shapeId":2}}
{"id":11,"method":"PowerPoint.setShapeText","params":{"targetId":"d17fd90d-90f8-4a29-b204-49495b78d064","slideId":256,"shapeId":2,"text":"Agenda"}}
{"id":11,"result":{"id":"d17fd90d-90f8-4a29-b204-49495b78d064","slideId":256,"shapeId":2,"text":"Agenda"}}

{"id":6,"method":"PowerPoint.getSlides","params":{"targetId":"d17fd90d-90f8-4a29-b204-49495b78d064"}}
{"id":6,"result":{"id":"d17fd90d-90f8-4a29-b204-49495b78d064","slides":[{"slideId":256,"slideIndex":1,"name":"Slide1","hidden":false}]}}

{"id":7,"method":"PowerPoint.getSlideState","params":{"targetId":"d17fd90d-90f8-4a29-b204-49495b78d064","slideId":256}}
{"id":7,"result":{"id":"d17fd90d-90f8-4a29-b204-49495b78d064","slide":{"slideId":256,"slideIndex":1,"name":"Slide1","hidden":false,"shapes":[{"shapeId":2,"zOrderPosition":1,"name":"Rectangle 1","shapeType":1,"text":"Agenda","bounds":{"left":36.0,"top":24.0,"width":648.0,"height":54.0}}]}}}

{"id":8,"method":"PowerPoint.getViewState","params":{"targetId":"d17fd90d-90f8-4a29-b204-49495b78d064"}}
{"id":8,"result":{"id":"d17fd90d-90f8-4a29-b204-49495b78d064","available":true,"windows":[{"viewType":1,"currentSlideId":256,"currentSlideIndex":1,"selection":{"type":0}}]}}

{"id":9,"method":"PowerPoint.getSlideShowState","params":{"targetId":"d17fd90d-90f8-4a29-b204-49495b78d064"}}
{"id":9,"result":{"id":"d17fd90d-90f8-4a29-b204-49495b78d064","running":false,"windows":[]}}


{"id":4,"method":"PowerPoint.unknown"}
{"id":4,"error":{"code":-32601,"message":"Unknown PowerPoint method"}}
```

`id` must be a nonnegative JavaScript-safe integer. `params`, when supplied,
must be an object. `newPresentation` requires a filename `name` and an existing
absolute `directory`; it saves a blank `.pptx` before returning. `POST /json/new`
accepts the same JSON body. Existing destinations return `-32005` (HTTP 409).
`PowerPoint.prepareShutdown` returns `{}` after checking all open presentations.
Its optional boolean `force` skips the unsaved-document guard; this method does
not quit or discard anything. The CLI rechecks process identity and document state
via external COM, closes each deck without saving when forced, then calls
`Application.Quit`. This avoids quitting inside the add-in's STA callback.
`timeoutMs` is an optional positive int32, default `10000`. Responses retain
the request ID; malformed JSON or an invalid ID produces `id: null`.
The state methods require a canonical `targetId`; `getSlideState` also requires a positive int32 `slideId`. Invalid parameters return `-32602`; unknown targets or slide IDs return `-32004`.

| Error code | Meaning |
| --- | --- |
| `-32700` | Malformed JSON |
| `-32600` | Invalid request envelope |
| `-32601` | Unknown method |
| `-32602` | Invalid method parameters, IDs, layout, or geometry |
| `-32000` | COM/internal failure or exhausted capacity |
| `-32001` | No document window or no running slide show for the target |
| `-32002` | Request deadline expired |
| `-32003` | Connection cancelled or addin stopping |
| `-32004` | Unknown or closed document target, or missing slide/shape ID |
| `-32005` | Conflict: existing filename, unsaved close, or unsaved shutdown |

Use WebSocket for disconnect-sensitive mutations. CivetWeb's released server
API cannot detect an HTTP client's disconnect while its handler waits for
the STA, so an abandoned POST or PUT may still execute until its queue deadline.
Server shutdown cancels queued work on both transports.

### Integration test

On Windows, with the native add-in built and registered and the CLI dependencies
installed, run the live PowerPoint presentation API integration test:
```powershell
npm --prefix src/cli run test:state-api
```

The test launches or reuses PowerPoint, creates two named test presentations,
checks reads and mutations over HTTP and WebSocket, then closes and removes them.
Run `npm --prefix src/cli run test:named` for live named creation,
`npm --prefix src/cli run test:close` for targeted close, force, and reopening,
`npm --prefix src/cli run test:formatting` for live design, master, and group
commands, and `npm --prefix src/cli run test:cli` for offline argument checks.
Run `npm --prefix src/cli run test:shutdown` separately in an isolated
PowerPoint session to check safe refusal, forced discard, and process exit.
Set `NETOFFICE_PORT` if the registered add-in uses a non-default port.

The endpoint is plaintext and unauthenticated, bound only to `127.0.0.1`.
No wildcard CORS headers are sent. Requests with an `Origin` must use exactly
`http://127.0.0.1:<port>` or `http://localhost:<port>`; duplicate or other origins
are rejected. Any local process able to reach the port can still issue commands.
This is not a multi-user security boundary and must not be exposed through a proxy.


## License

Source code is licensed under [MIT License](LICENSE.txt).
