#requires -Version 7.0
[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [ValidateSet('Register', 'Unregister')]
    [string] $Action,

    # Must match PowerPoint, not the PowerShell process.
    [Parameter(Mandatory)]
    [ValidateSet('x86', 'x64')]
    [string] $Architecture,

    [string] $DllPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
if (-not $IsWindows) { throw 'Office addin registration requires Windows.' }

$progId = 'NetOffice.Automate'
$clsid = '{6d274715-3f05-4505-aa63-2ae9df1e6881}'
$classKey = "Software\Classes\CLSID\$clsid"
$progIdKey = "Software\Classes\$progId"
$addinKey = "Software\Microsoft\Office\PowerPoint\Addins\$progId"
$view = if ($Architecture -eq 'x86') {
    [Microsoft.Win32.RegistryView]::Registry32
} else {
    [Microsoft.Win32.RegistryView]::Registry64
}

if ($Action -eq 'Register') {
    if (-not $DllPath) { throw '-DllPath is required for Register.' }
    $DllPath = (Resolve-Path -LiteralPath $DllPath).ProviderPath
    $reader = [IO.BinaryReader]::new([IO.File]::OpenRead($DllPath))
    try {
        if ($reader.ReadUInt16() -ne 0x5A4D) { throw 'Not a Windows DLL.' }
        $reader.BaseStream.Position = 0x3C
        $peOffset = $reader.ReadInt32()
        $reader.BaseStream.Position = $peOffset
        if ($reader.ReadUInt32() -ne 0x00004550) { throw 'Invalid PE signature.' }
        $machine = $reader.ReadUInt16()
        $expected = if ($Architecture -eq 'x86') { 0x014C } else { 0x8664 }
        if ($machine -ne $expected) { throw "DLL architecture does not match $Architecture." }
        $reader.BaseStream.Position = $peOffset + 22
        if (($reader.ReadUInt16() -band 0x2000) -eq 0) { throw 'PE file is not a DLL.' }
    } finally {
        $reader.Dispose()
    }
}

$registry = [Microsoft.Win32.RegistryKey]::OpenBaseKey(
    [Microsoft.Win32.RegistryHive]::CurrentUser, $view)
try {
    if ($Action -eq 'Unregister') {
        # Only this addin's keys; safe to repeat, even if the DLL was removed.
        $registry.DeleteSubKeyTree($addinKey, $false)
        $registry.DeleteSubKeyTree($progIdKey, $false)
        $registry.DeleteSubKeyTree($classKey, $false)
    } else {
        $values = @{
            $classKey = @{ '' = $progId }
            "$classKey\ProgID" = @{ '' = $progId }
            "$classKey\InprocServer32" = @{ '' = $DllPath; ThreadingModel = 'Apartment' }
            $progIdKey = @{ '' = $progId }
            "$progIdKey\CLSID" = @{ '' = $clsid }
            $addinKey = @{
                FriendlyName = 'NetOffice Automate'
                Description = 'Minimal NetOffice PowerPoint COM addin'
                LoadBehavior = 3
            }
        }
        foreach ($path in $values.Keys) {
            $key = $registry.CreateSubKey($path)
            try {
                foreach ($name in $values[$path].Keys) {
                    $value = $values[$path][$name]
                    $kind = if ($value -is [int]) {
                        [Microsoft.Win32.RegistryValueKind]::DWord
                    } else {
                        [Microsoft.Win32.RegistryValueKind]::String
                    }
                    $key.SetValue($name, $value, $kind)
                }
            } finally {
                $key.Dispose()
            }
        }
    }
} finally {
    $registry.Dispose()
}
Write-Output "$Action completed for $progId ($Architecture, current user)."
