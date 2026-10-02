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
npm --prefix src/cli link
netoffice --help
netoffice powerpoint launch
netoffice presentation new --title "Hello from NetOffice"
netoffice slide title "Updated title" --slide 1
```

Without linking, use `node src/cli/bin/netoffice.js` with the same arguments.
`powerpoint launch` starts visible desktop PowerPoint on Windows and waits for
the addin's WebSocket JSON readiness response. An already-ready session is reused.
The document commands require that session to be running; they do not launch it
implicitly or fall back to external COM automation.

`presentation new` creates an unsaved, active presentation with one title slide.
Without `--title`, the title is empty. `slide title` updates a title placeholder
on the active presentation; slide indexes are one-based and default to `1`.
Use `--` before positional title text starting with a dash.

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
| `PUT /json/activate/{target}` | Activate the document and bring its window to the foreground |
| `PUT /json/close/{target}` | Close a saved document; HTTP 409 when it has unsaved changes |
| `PUT /json/close/{target}?force` | Discard changes and close the document without saving |
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

HTTP `/json/new` creates Office's ordinary blank presentation with no added slides;
the CLI `presentation new` retains its one-title-slide behavior.
Activate and close return HTTP 200 with `{}`. The API rejects unsaved close with
HTTP 409 without modifying the document's `Saved` state. The `force` parameter
is a presence flag: `?force`, `?force=1`, and `?force=false` all authorize discard.
No operation automatically saves a document.

Malformed queries, duplicate `url`/`force` parameters, and invalid target IDs
return HTTP 400. Unknown or closed targets return HTTP 404; there is no fallback
to the active presentation. Unsupported verbs return HTTP 405 with `Allow`.
Deadline expiry is HTTP 504, shutdown is HTTP 503, and COM failures are HTTP 500.
Errors are JSON `{"error":{"code":...,"message":...,"data":...}}`, with optional
HRESULT details. HTTP target requests have a 10-second queue deadline; a started
Office call may outlive it. Inspect the document list before repeating a mutation.

### Application WebSocket commands


```json
{"id":1,"method":"PowerPoint.getStatus"}
{"id":1,"result":{"processId":1234}}

{"id":2,"method":"PowerPoint.newPresentation","params":{"title":"Hello"},"timeoutMs":10000}
{"id":2,"result":{"name":"Presentation1","slideCount":1}}

{"id":3,"method":"PowerPoint.setSlideTitle","params":{"slideIndex":1,"text":"Updated title"}}
{"id":3,"result":{}}

{"id":4,"method":"PowerPoint.unknown"}
{"id":4,"error":{"code":-32601,"message":"Unknown PowerPoint method"}}
```

`id` must be a nonnegative JavaScript-safe integer. `params`, when supplied,
must be an object. `newPresentation` accepts an optional string `title`;
`setSlideTitle` requires a positive int32 `slideIndex` and string `text`.
`timeoutMs` is an optional positive int32, default `10000`. Responses retain
the request ID; malformed JSON or an invalid ID produces `id: null`.

| Error code | Meaning |
| --- | --- |
| `-32700` | Malformed JSON |
| `-32600` | Invalid request envelope |
| `-32601` | Unknown method |
| `-32602` | Invalid method parameters or slide index |
| `-32000` | COM/internal failure or exhausted capacity |
| `-32001` | No active presentation or title placeholder |
| `-32002` | Request deadline expired |
| `-32003` | Connection cancelled or addin stopping |

Use WebSocket for disconnect-sensitive mutations. CivetWeb's released server
API cannot detect an HTTP client's disconnect while its handler waits for
the STA, so an abandoned POST or PUT may still execute until its queue deadline.
Server shutdown cancels queued work on both transports.

The endpoint is plaintext and unauthenticated, bound only to `127.0.0.1`.
No wildcard CORS headers are sent. Requests with an `Origin` must use exactly
`http://127.0.0.1:<port>` or `http://localhost:<port>`; duplicate or other origins
are rejected. Any local process able to reach the port can still issue commands.
This is not a multi-user security boundary and must not be exposed through a proxy.


## License

Source code is licensed under [MIT License](LICENSE.txt).
