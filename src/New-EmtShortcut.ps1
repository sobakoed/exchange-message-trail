#Requires -Version 5.1
<#
.SYNOPSIS
    Makes the shortcut "Exchange Message Trail" with the EMT icon in the EMT folder.

.DESCRIPTION
    The shortcut starts EMT from this folder. It has the EMT icon and the application ID of the EMT window:
    a shortcut on the taskbar and the window of EMT are one taskbar button.
    The shortcut opens EMT.cmd of the same folder. It has no PowerShell command line: an antivirus can delete
    a shortcut that starts a hidden PowerShell directly.
    EMT makes the shortcut at the first start and after a move of the folder. A shortcut that was removed later
    (by the user or by an antivirus) stays removed: run this script to make it again.
    Copy the shortcut to the desktop or to the Start menu if necessary.

.PARAMETER Path
    A different folder for the shortcut. The default is the EMT folder.

.EXAMPLE
    powershell.exe -ExecutionPolicy Bypass -File .\src\New-EmtShortcut.ps1
#>
[CmdletBinding()]
param([string] $Path)

$ErrorActionPreference = 'Stop'
$root   = Split-Path -Parent $PSScriptRoot
$cmd    = Join-Path $root 'EMT.cmd'
$icon   = Join-Path $root 'gui\emt.ico'
foreach ($f in $cmd, $icon) { if (-not (Test-Path $f)) { throw "The file is missing: $f" } }

# The same ID as in ExchangeMessageTrail.ps1 (SetCurrentProcessExplicitAppUserModelID).
$appId = 'ExchangeMessageTrail'

if (-not ('EmtShortcut' -as [type])) {
    Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

public static class EmtShortcut {
    [ComImport, Guid("00021401-0000-0000-C000-000000000046")] class ShellLink { }

    [ComImport, InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("000214F9-0000-0000-C000-000000000046")]
    interface IShellLinkW {
        void GetPath([Out, MarshalAs(UnmanagedType.LPWStr)] System.Text.StringBuilder f, int cch, IntPtr fd, uint flags);
        void GetIDList(out IntPtr ppidl);
        void SetIDList(IntPtr pidl);
        void GetDescription([Out, MarshalAs(UnmanagedType.LPWStr)] System.Text.StringBuilder s, int cch);
        void SetDescription([MarshalAs(UnmanagedType.LPWStr)] string s);
        void GetWorkingDirectory([Out, MarshalAs(UnmanagedType.LPWStr)] System.Text.StringBuilder s, int cch);
        void SetWorkingDirectory([MarshalAs(UnmanagedType.LPWStr)] string s);
        void GetArguments([Out, MarshalAs(UnmanagedType.LPWStr)] System.Text.StringBuilder s, int cch);
        void SetArguments([MarshalAs(UnmanagedType.LPWStr)] string s);
        void GetHotkey(out short k);
        void SetHotkey(short k);
        void GetShowCmd(out int c);
        void SetShowCmd(int c);
        void GetIconLocation([Out, MarshalAs(UnmanagedType.LPWStr)] System.Text.StringBuilder s, int cch, out int i);
        void SetIconLocation([MarshalAs(UnmanagedType.LPWStr)] string s, int i);
        void SetRelativePath([MarshalAs(UnmanagedType.LPWStr)] string s, uint r);
        void Resolve(IntPtr hwnd, uint flags);
        void SetPath([MarshalAs(UnmanagedType.LPWStr)] string s);
    }

    [StructLayout(LayoutKind.Sequential, Pack = 4)]
    struct PROPERTYKEY { public Guid fmtid; public uint pid; }

    [StructLayout(LayoutKind.Explicit)]
    struct PROPVARIANT { [FieldOffset(0)] public ushort vt; [FieldOffset(8)] public IntPtr p; }

    [ComImport, InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("886D8EEB-8CF2-4446-8D02-CDBA1DBDCF99")]
    interface IPropertyStore {
        void GetCount(out uint c);
        void GetAt(uint i, out PROPERTYKEY k);
        void GetValue(ref PROPERTYKEY k, out PROPVARIANT v);
        void SetValue(ref PROPERTYKEY k, ref PROPVARIANT v);
        void Commit();
    }

    public static void Save(string file, string target, string args, string workDir, string icon, string description, string appId) {
        IShellLinkW link = (IShellLinkW)new ShellLink();
        link.SetPath(target);
        link.SetArguments(args);
        link.SetWorkingDirectory(workDir);
        link.SetIconLocation(icon, 0);
        link.SetDescription(description);
        link.SetShowCmd(7);    // SW_SHOWMINNOACTIVE: the window of EMT.cmd starts minimized and closes at once
        IPropertyStore store = (IPropertyStore)link;
        PROPERTYKEY key = new PROPERTYKEY();
        key.fmtid = new Guid("9F4C2855-9F79-4B39-A8D0-E1D42DE1D5F3"); key.pid = 5;    // System.AppUserModel.ID
        PROPVARIANT v = new PROPVARIANT();
        v.vt = 31;                                                                  // VT_LPWSTR
        v.p = Marshal.StringToCoTaskMemUni(appId);
        try { store.SetValue(ref key, ref v); store.Commit(); }
        finally { Marshal.FreeCoTaskMem(v.p); }
        ((IPersistFile)link).Save(file, true);
    }
}
'@
}

$folder = if ($Path) { $Path } else { $root }

if (-not (Test-Path $folder)) { New-Item -ItemType Directory -Path $folder | Out-Null }
$file = Join-Path $folder 'Exchange Message Trail.lnk'
[EmtShortcut]::Save($file, $cmd, '', $root, $icon, 'Exchange Message Trail: message tracking and delete for Exchange', $appId)
$file
