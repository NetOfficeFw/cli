# NetOffice Automate

> NetOffice command line tool for running automation tests of Microsoft Office applications.

## PowerPoint addin

`src/addin` builds the minimal native `NetOffice.Automate` COM addin
(`{6d274715-3f05-4505-aa63-2ae9df1e6881}`). It implements
`IDTExtensibility2` connection callbacks only; there is no UI or automation.
`Connect.h` and `Connect.cpp` follow the native ATL addin pattern:
`IDispatchImpl` uses Office's imported extensibility interface, and ATL manages
COM lifetime and the class factory.

Build from a Visual Studio developer shell with the v145 C++ toolset,
Windows SDK, ATL, and Office's `MSADDNDR.OLB` installed. The project searches
the 32-bit Click-to-Run Office `DESIGNER` directory by default. For a different
Office layout, pass `/p:OfficeExtensibilityDir="path to the directory containing MSADDNDR.OLB"`.
No third-party library is required.

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


## License

Source code is licensed under [MIT License](LICENSE.txt).
