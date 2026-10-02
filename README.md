# NetOffice Automate

> NetOffice command line tool for running automation tests of Microsoft Office applications.

## PowerPoint addin

`src/addin` builds the native `NetOffice.Automate` COM addin
(`{6d274715-3f05-4505-aa63-2ae9df1e6881}`). Its `IDTExtensibility2`
connection starts a gRPC server on `127.0.0.1:50051`. Document commands are
queued onto PowerPoint's owning STA; gRPC worker threads never call Office COM.
`Connect.h` and `Connect.cpp` follow the native ATL addin pattern:
`IDispatchImpl` uses Office's imported extensibility interface, and ATL manages
COM lifetime and the class factory.

Build from a Visual Studio developer shell with the v145 C++ toolset,
Windows SDK, ATL, and Office's `MSADDNDR.OLB` installed. The project searches
the 32-bit Click-to-Run Office `DESIGNER` directory by default. For a different
Office layout, pass `/p:OfficeExtensibilityDir="path to the directory containing MSADDNDR.OLB"`.
Install [vcpkg](https://github.com/microsoft/vcpkg) and run
`vcpkg integrate install` once. The manifest pins gRPC and protobuf through a
registry baseline. MSBuild installs static libraries with the dynamic CRT and
generates C++ protocol files from `src/proto/netoffice.proto` into `obj`.

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
pwsh -NoProfile -File src/addin/Register-Addin.ps1 -Action Register -Architecture x86 -DllPath src/addin/build/Release_Win32/addin.dll
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
the addin's gRPC readiness response. An already-ready session is reused.
The document commands require that session to be running; they do not launch it
implicitly or fall back to external COM automation.

`presentation new` creates an unsaved, active presentation with one title slide.
Without `--title`, the title is empty. `slide title` updates a title placeholder
on the active presentation; slide indexes are one-based and default to `1`.
Use `--` before positional title text starting with a dash.

All commands accept `--timeout <milliseconds>` (default `10000`). Failures exit
nonzero with a readable error; document mutations are not retried. If a request
times out after execution has started, inspect PowerPoint before repeating it.
There is no active document: `FAILED_PRECONDITION`; an out-of-range slide:
`INVALID_ARGUMENT`. Disconnecting the addin or closing PowerPoint stops the server.

The shared service exposes `GetStatus`, `NewPresentation`, and `SetSlideTitle`.
The endpoint is plaintext and unauthenticated, restricted to IPv4 loopback.
Any local process that can reach the port can issue these commands; it is not
intended as a multi-user security boundary. Only one addin connection can own
the fixed port at a time.

`npm --prefix src/cli pack` creates a standalone npm package including the
protocol schema; the prepack script copies it from the shared source.


## License

Source code is licensed under [MIT License](LICENSE.txt).
