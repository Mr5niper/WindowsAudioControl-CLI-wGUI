# Technical Specifications: audioctl — Windows Audio CLI/GUI Utility (v1.4.3.3)
## 1. Introduction
`audioctl` is a Windows audio control utility offering:
- a command‑line interface (CLI) for scripting, automation, and diagnostics, and
- a Tkinter‑based graphical user interface (GUI) for interactive control.
It targets reliable control of Windows Core Audio endpoints (Render/Capture) using COM via `comtypes`, with targeted `ctypes` vtable calls on critical paths (e.g., IPropertyStore). The package is designed to run from source or as a PyInstaller single‑file executable.
Core capabilities:
- enumerate playback/recording endpoints with human‑friendly names and default flags,
- set default endpoints by role,
- read/set master volume and mute,
- enable/disable “Listen to this device” on capture devices,
- enable/disable “Audio Enhancements” (SysFX) using a vendor‑first approach with a learnable INI, and
- diagnostics and discovery tooling for Enhancements behavior.
This document describes what the tool is and how it works in v1.4.3.3 (not a change log).
---
## 2. Core Technologies & Dependencies
- Python 3.x
- comtypes  
  COM automation from Python, including generated wrappers for Core Audio interfaces.
- ctypes  
  Low‑level FFI for raw IPropertyStore vtable calls, PROPVARIANT handling, and helper DLLs (e.g., PropVariantClear).
- pycaw  
  Convenience interfaces for IMMDevice, IAudioEndpointVolume, and utilities used for enumeration helpers and diagnostics.
- tkinter  
  GUI (Treeview device listing, context menus, status UI).
- argparse, json, re, os, sys, time, io, warnings  
  CLI parsing, structured output, pattern matching, and utilities.
- winreg  
  MMDevices registry access (verification, diagnostics, vendor toggles, and Listen feature playback target).
- Windows Core Audio COM APIs  
  - IMMDeviceEnumerator, IMMDevice
  - IAudioEndpointVolume
  - IPropertyStore
  - PolicyConfig (Vista/Client variants, Fx‑capable for SysFX reads/writes and SetDefault)
- PyInstaller  
  Produces a single‑file executable; integrates version metadata and icon.
- External assets
  - audio.ico: application/executable icon
  - version.txt: PE version resource for the built executable
  - vendor_toggles.ini: append‑only INI used to store learned vendor toggles for Enhancements
---
## 3. Dependency Management (Virtual Environment)
Create and use a dedicated virtual environment:
a) Create:
```bash
python -m venv venv
```
b) Activate:
- CMD
  ```bat
  .\venv\Scripts\activate.bat
  ```
- PowerShell
  ```powershell
  .\venv\Scripts\Activate.ps1
  ```
c) Recreate the environment elsewhere:
```bash
pip freeze > requirements.txt
pip install -r requirements.txt
```
Example requirements:
```
altgraph==0.17.4
comtypes==1.4.13
packaging==25.0
pefile==2023.2.7
psutil==7.1.1
pycaw==20251023
pyinstaller==6.16.0
pyinstaller-hooks-contrib==2025.9
pywin32-ctypes==0.2.3
setuptools==80.9.0
wheel==0.45.1
```
---
## 4. Architectural Overview
Package layout (selected modules):
- audioctl/compat.py  
  Early compatibility shim and shared constants (must be importable before `pycaw/comtypes`).
- audioctl/devices.py  
  Core Windows audio operations (enumeration, selection, default routing, volume/mute, Listen, SysFX readers/writers, diagnostics).
- audioctl/vendor_db.py  
  Vendor‑first Enhancements toggle system (INI I/O, verification, and learn helpers).
- audioctl/gui.py  
  Tkinter GUI (device browser and context menu actions).
- audioctl/cli.py  
  CLI parser and command handlers.
- audioctl/logging_setup.py  
  Centralized logging: lazy initialization, crash hooks, and optional debug tracing.
- audioctl.py / audioctl/__main__.py  
  Console/package entrypoints to `audioctl.cli.main()`.
Behavior:
- If launched without arguments, the GUI opens.
- Otherwise, the requested CLI command executes in the current console.
---
## 5. Key Design Patterns & Rationale
### 5.1. COM interop strategy
- Use `comtypes` for convenience where stable (e.g., IAudioEndpointVolume).
- Use `ctypes` vtable calls for sensitive paths, with **interface definitions cached at module load** to prevent runtime GC crashes. This applies to `IPropertyStore`, `IPolicyConfigFx`, and `IPolicyConfigVista`.
- `comtypes.CoInitialize()` and `comtypes.CoUninitialize()` are called at the top‑level (CLI and GUI). The library's `atexit` hook correctly manages shutdown, fixing race conditions.
### 5.2. comtypes compatibility shim (compat.py)
This shim is critical for stability, especially in packaged executables. It ensures:
- `PROPVARIANT` and standard VT/VARIANT constants are available.
- **Crucially, it forces the import of `comtypes._post_coinit` modules.** This ensures the library's COM object cleanup code (`__del__` -> `Release`) is loaded upfront, preventing dynamic import timing issues that lead to crashes during garbage collection or shutdown.
### 5.3. Device enumeration and naming
- `list_devices()` uses `IMMDeviceEnumerator.EnumAudioEndpoints()` for Render and Capture.
- Friendly names are resolved via a one‑time `pycaw.AudioUtilities.GetAllDevices()` map per listing call.
- A hardened raw IPropertyStore reader exists as a fallback (using cached interface definitions).
- The tool sorts by name within each flow. CLI selection indices (`--index`) match GUI view order.
### 5.4. Default endpoint management
- `set_default_endpoint()` uses the cached `IPolicyConfigVista` interface to set defaults by role: `console`, `multimedia`, `communications`, or `all`.
### 5.5. “Listen to this device” (Capture)

- **Hybrid Write Path:**
  - The **"Listen to this device" checkbox** is set via a raw `IPropertyStore::SetValue` call to the property key identified by the GUID `{24dbb0fc-9311-4b3d-9cf0-18ff155639d4}` and a Property ID of `1`.
  - The **"Playback through this device" target** is set via a direct `winreg` write to the correct registry location:
    - **Key:** `HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\Capture\{endpointGUID}\Properties`
    - **Value Name:** `{24dbb0fc-9311-4b3d-9cf0-18ff155639d4},0`
    - **Type:** `REG_SZ` (string)
    - **Data:** The full device ID of the playback target, or an empty string `""` for "Default Playback Device".
- **Admin Rights:** Writing the playback target to `HKLM` may require Administrator privileges on some systems, though it often works from non-elevated prompts.
- **Read Path:** The checkbox status is read via `IPropertyStore GetValue`.
- **Verification:** COM read is preferred; a short registry poll on `HKCU` is used as a fallback for the checkbox status.
### 5.6. Volume and mute (Render/Capture)
- `IAudioEndpointVolume` via `IMMDevice.Activate`.
- Robust conversions (tuple vs out‑param) produce:
  - volume 0..100 %
  - mute True/False (None on read failure).
### 5.7. Enhancements (SysFX) — vendor‑first control and learn
- (This section remains accurate and is unchanged).
### 5.8. Logging, diagnostics, and debug (v1.4.3.3)
- (This section remains accurate and is unchanged).
---
## 6. CLI Command Reference
All mutating commands operate on active endpoints. Name selections use GUI order; if multiple matches exist, use `--index`.
- list  
  ```bash
  audioctl list [--all] [--json]
  ```
- set-default  
  ```bash
  audioctl set-default
    (--playback-id <ID> | --playback-name <NAME>) [--playback-role <console|multimedia|communications|all>]
    (--recording-id <ID> | --recording-name <NAME>) [--recording-role <console|multimedia|communications|all>]
    [--index <N>] [--regex]
  ```
- set-volume  
  ```bash
  audioctl set-volume
    (--id <ID> | --name <NAME>) [--flow <Render|Capture>]
    (--level <0..100> | --mute | --unmute)
    [--index <N>] [--regex]
  ```
- listen (Capture only)  
  ```bash
  audioctl listen
    (--id <ID> | --name <NAME>)
    (--enable | --disable)
    [--playback-target-id [<RenderEndpointID>]]
    [--playback-target-name [<RenderEndpointName>]]
    [--index <N>] [--regex]
  ```
  *Note: Providing `--playback-target-id` or `--playback-target-name` without a value sets the target to "Default Playback Device".*
- enhancements (vendor‑first SysFX)  
  - (This section remains accurate and is unchanged).
- diag-sysfx  
  - (This section remains accurate and is unchanged).
- diag-mmdevices  
  - (This section remains accurate and is unchanged).
- discover-enhancements  
  - (This section remains accurate and is unchanged).
- wait  
  - (This section remains accurate and is unchanged).
---
## 7. GUI Application Structure
The Tkinter GUI lists Render and Capture endpoints in two groups and mirrors CLI order/selection. Right‑click a device to:
- Set as Default (all roles),
- Set Volume…,
- Mute/Unmute,
- Toggle Listen (capture only),
- Enable/Disable Enhancements (when a vendor toggle is known),
- Learn Enhancements (append a vendor INI entry via guided steps).
Quality/stability:
- **Runtime Stability:** The GUI disables Python's cyclical garbage collector (`gc.disable()`) on startup to prevent unpredictable COM object cleanup during the event loop, which was a source of crashes.
- **Shutdown Stability:** The application uses `comtypes.CoInitialize()` and `comtypes.CoUninitialize()`, leveraging the library's `atexit` hook to ensure all COM objects are released cleanly before the COM library is shut down, preventing fatal exceptions on exit.
- **Layout:** The window auto-sizes to content, with a small height buffer added to prevent the scrollbar from appearing unnecessarily.
---
## 8. Build (PyInstaller)
There are two methods to build the single-file executable. Method B is recommended for consistency.

### Method A: Using Command-Line Arguments
This method is useful for one-off builds or scripting without relying on a `.spec` file.

```powershell
pyinstaller -F --noupx --clean --console --name audioctl --bootloader-ignore-signals --collect-all pycaw --collect-all comtypes --hidden-import comtypes.automation --hidden-import comtypes._post_coinit --hidden-import comtypes._post_coinit.unknwn --hidden-import comtypes._post_coinit.misc --icon audio.ico --add-data "audio.ico;." --version-file version.txt .\audioctl.py
```

### Method B: Using a `.spec` File (Recommended)
This method is simpler to run and guarantees all build options are correct.

**1. `audioctl.spec` File Content**  
Ensure your `audioctl.spec` file contains the following. This file is a Python script that provides the build configuration to PyInstaller. The absence of a `COLLECT` block at the end instructs PyInstaller to create a single-file executable.

```python
# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_all

# Corresponds to: --add-data "audio.ico;."
datas = [('audio.ico', '.')]
binaries = []

# Corresponds to all the --hidden-import flags
hiddenimports = [
    'comtypes.automation',
    'comtypes._post_coinit',
    'comtypes._post_coinit.unknwn',
    'comtypes._post_coinit.misc'
]

# Corresponds to: --collect-all pycaw
tmp_ret = collect_all('pycaw')
datas += tmp_ret[0]
binaries += tmp_ret[1]
hiddenimports += tmp_ret[2]

# Corresponds to: --collect-all comtypes
tmp_ret = collect_all('comtypes')
datas += tmp_ret[0]
binaries += tmp_ret[1]
hiddenimports += tmp_ret[2]

a = Analysis(
    ['audioctl.py'], # The main script
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    # Corresponds to: --name audioctl
    name='audioctl',
    debug=False,
    # Corresponds to: --bootloader-ignore-signals
    bootloader_ignore_signals=True,
    strip=False,
    # Corresponds to: --noupx
    upx=False,
    upx_exclude=[],
    runtime_tmpdir=None,
    # Corresponds to: --console
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    # Corresponds to: --version-file version.txt
    version='version.txt',
    # Corresponds to: --icon audio.ico
    icon=['audio.ico'],
)
```

**2. Build Command**  
With the `.spec` file saved, run this simple command:

```powershell
pyinstaller audioctl.spec --clean
```

### Version Resource (`version.txt`)
This file provides the version metadata embedded into the final `.exe` file's properties.

Example for v1.4.3.3:
```text
VSVersionInfo(
  ffi=FixedFileInfo(
    filevers=(1, 4, 3, 3),
    prodvers=(1, 4, 3, 3),
    mask=0x3f,
    flags=0x0,
    OS=0x40004,
    fileType=0x1,
    subtype=0x0,
    date=(0, 0)
  ),
  kids=[
    StringFileInfo([
      StringTable('040904E4', [
          StringStruct('CompanyName', 'Mr5niper5oft'),
          StringStruct('FileDescription', 'Windows Audio Control CLI'),
          StringStruct('FileVersion', '1.4.3.3'),
          StringStruct('InternalName', 'audioctl'),
          StringStruct('LegalCopyright', 'Copyright (c) 2025 Mr5niper5oft'),
          StringStruct('OriginalFilename', 'audioctl.exe'),
          StringStruct('ProductName', 'Windows Audio Control CLI'),
          StringStruct('ProductVersion', '1.4.3.3'),
        ]
      )
    ]),
    VarFileInfo([VarStruct('Translation', [1033, 1252])])
  ]
)
```
---
## 9. Diagnostics & Debug Tools
- Enable verbose debug logging:
  - CMD:
    ```bat
    set AUDIOCTL_DEBUG=1
    audioctl.exe
    ```
  - PowerShell:
    ```powershell
    $env:AUDIOCTL_DEBUG = 1
    .\audioctl.exe
    ```
- `diag-sysfx`: inspect live Windows (PropertyStore/COM) vs vendor state.
- `diag-mmdevices`: full MMDevices dump for an endpoint (HKCU/HKLM).
- `discover-enhancements`: produce TXT/JSON bundles and optional INI snippets.
- GUI “Print CLI commands”: echo CLI equivalents of actions.
---
## 10. Learn: Vendor Enhancements Discovery, INI, and Runtime
- INI location: `vendor_toggles.ini` next to the executable (override path via `--vendor-ini` in CLI).
- INI section template:
  ```
  [vendor_{GUID},pid]
  value_name = {GUID},pid
  dword_enable = 0
  dword_disable = 1
  hives = HKCU,HKLM
  flows = Render,Capture
  notes = optional
  ```
- Manual learn (CLI/GUI): guided A/B toggle captures; appends to INI on a reliable DWORD flip under FxProperties.
- Discovery (CLI): produces TXT/JSON report; can append a suggested INI snippet.
- Runtime toggles: the first applicable vendor entry (INI first, then built‑in) is written and verified; HKLM writes need admin.
---
## 11. Conclusion
`audioctl` v1.4.3.3 provides a dependable, scriptable, and interactive way to manage Windows audio endpoints. It incorporates significant stability fixes for COM object lifecycle management, resolving both runtime and shutdown crashes. It now correctly handles the "Listen to this device" playback target and offers more user-friendly CLI options. The utility blends `comtypes` convenience with targeted `ctypes` vtable calls and direct registry access for robust Listen/SysFX operations.
