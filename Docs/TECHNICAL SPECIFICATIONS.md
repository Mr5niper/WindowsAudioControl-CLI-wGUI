# Technical Specifications: audioctl — Windows Audio CLI/GUI Utility (v1.4.3.1)

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

This document describes what the tool is and how it works in v1.4.3.1 (not a change log).

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
  MMDevices registry access (verification, diagnostics, vendor toggles).
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
- Use `ctypes` vtable calls for sensitive or packaging‑fragile paths (IPropertyStore SetValue/GetValue; Listen and Disable_SysFx reads/writes).
- `CoInitialize()`/`CoUninitialize()` are called at the top‑level (CLI and GUI). Helpers assume COM is already initialized by the caller.

### 5.2. comtypes compatibility shim (compat.py)
Ensures for packaged and source runs:
- `PROPVARIANT` (alias to `tagPROPVARIANT` when needed),
- VT constants: `VT_LPWSTR`, `VT_BOOL`, `VT_UI2`, `VT_UI4`,
- `VARIANT_TRUE`/`VARIANT_FALSE`,
- Endpoint role constants, device state flags, and STGM flags.

### 5.3. Device enumeration and naming
- `list_devices()` uses `IMMDeviceEnumerator.EnumAudioEndpoints()` for Render and Capture.
- Friendly names are resolved via a one‑time `pycaw.AudioUtilities.GetAllDevices()` map per listing call; this avoids unnecessary raw IPropertyStore calls.
- A hardened raw IPropertyStore reader exists as a fallback (balanced AddRef/Release and careful PROPVARIANT parsing).
- The tool sorts by name within each flow. CLI selection indices (`--index`) match GUI view order.

### 5.4. Default endpoint management
- `set_default_endpoint()` uses PolicyConfig (Vista/Client) to set defaults by role: `console`, `multimedia`, `communications`, or `all`.
- The default‑setting path remains robust and logs diagnostics; the SysFX codepath internally uses an Fx‑capable PolicyConfig singleton for its own reads/writes.

### 5.5. “Listen to this device” (Capture)
- Write path: raw IPropertyStore `SetValue`/`Commit`:
  - `PKEY_LISTEN_ENABLE` `{24dbb0fc-9311-4b3d-9cf0-18ff155639d4}`, pid 1 (boolean)
  - `PKEY_LISTEN_PLAYBACKTHROUGH` same GUID, pid 2 (LPWSTR endpoint id; `""` means “Default Playback Device”)
- Read path: IPropertyStore `GetValue`.
- Verification: COM read preferred; short registry poll fallback under
  `HKCU\...\MMDevices\Audio\Capture\{endpointGUID}\(FxProperties|Properties)` with support for `REG_DWORD`, `REG_BINARY` (PROPVARIANT), and `REG_SZ`.

### 5.6. Volume and mute (Render/Capture)
- IAudioEndpointVolume via `IMMDevice.Activate`.
- Robust conversions (tuple vs out‑param) produce:
  - volume 0..100 %
  - mute True/False (None on read failure).

### 5.7. Enhancements (SysFX) — vendor‑first control and learn
- Read paths:
  - PropertyStore: `PKEY_AudioEndpoint_Disable_SysFx` (parse VT_BOOL/UI2/UI4),
  - PolicyConfig (Fx and normal stores).
- Vendor toggles (control path):
  - Device‑specific DWORDs under endpoint `FxProperties/Properties` using either:
    - learned INI entries (`vendor_toggles.ini`), or
    - built‑in vendor entries (e.g., Realtek/Waves).
  - Writes across configured hives; HKLM writes require admin.
  - Post‑write verification polls the same DWORD for a short window.
- Learn, discovery, and diagnostics:
  - Manual learn (CLI `enhancements --learn`, GUI “Learn Enhancements”):
    - You toggle Windows Enhancements ON/OFF on prompt.
    - The tool captures A/B registry snapshots, diffs them, and appends a DWORD entry to `vendor_toggles.ini` if a reliable flip is found under FxProperties.
  - Discovery (CLI `discover-enhancements`):
    - Produces TXT/JSON bundles with COM/PropertyStore/registry snapshots and diffs; can append a suggested INI snippet.
  - Diagnostics (CLI `diag-sysfx`/`diag-mmdevices`):
    - Inspect Windows’ live state and registry surfaces alongside the vendor entry’s current value.

### 5.8. Logging, diagnostics, and debug (v1.4.3.1)
- Lazy initialization:
  - Logging file and hooks are created on first write, not at import time.
  - File path resolves to the executable folder; if not writable, a temp folder is used.
- Hooks:
  - `sys.excepthook`, `sys.unraisablehook`,
  - optional `faulthandler` to the log file,
  - `sys.exit` wrapper logs exit codes,
  - Windows console control handler logs control events.
- Debug trace:
  - Set `AUDIOCTL_DEBUG=1` to include `[DBG pid=... tid=...]` lines (idempotently starts logging if needed).
  - `logging_setup` utilities safely handle repeated initialization and are robust to import‑order differences.

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
    [--playback-target-id <RenderEndpointID-or-empty>]
    [--index <N>] [--regex]
  ```

- enhancements (vendor‑first SysFX)  
  ```bash
  # Enable/disable via vendor toggle
  audioctl enhancements
    (--id <ID> | --name <NAME>) [--flow <Render|Capture>]
    (--enable | --disable)
    [--index <N>] [--regex]
    [--prefer-hklm] [--vendor-ini <path>]

  # Manual learn (append device-specific vendor entry to INI)
  audioctl enhancements
    (--id <ID> | --name <NAME>) [--flow <Render|Capture>]
    --learn
    [--index <N>] [--regex]
    [--vendor-ini <path>]
  ```

- diag-sysfx  
  ```bash
  audioctl diag-sysfx
    (--id <ID> | --name <NAME>) [--flow <Render|Capture>]
    [--index <N>] [--regex]
  ```

- diag-mmdevices  
  ```bash
  audioctl diag-mmdevices
    (--id <ID> | --name <NAME>) [--flow <Render|Capture>]
    [--index <N>] [--regex]
  ```

- discover-enhancements  
  ```bash
  audioctl discover-enhancements
    (--id <ID> | --name <NAME>) [--flow <Render|Capture>]
    [--index <N>] [--regex]
    [--output-dir <path>] [--ini-snippet <path>]
  ```

- wait  
  ```bash
  audioctl wait
    (--id <ID> | --name <NAME>) [--flow <Render|Capture>]
    [--timeout <seconds>] [--index <N>] [--regex]
  ```

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
- COM is initialized at startup and uninitialized at shutdown.
- Layout sizes adapt to content; the UI disables group headers from selection and actions.
- “Print CLI commands” echoes exact CLI equivalents of GUI actions.

---

## 8. Build (PyInstaller)
Use the new build command (PowerShell):
```powershell
pyinstaller -F --noupx --clean --console --name audioctl --collect-all pycaw --collect-all comtypes --hidden-import comtypes.automation --hidden-import comtypes._post_coinit --hidden-import comtypes._post_coinit.unknwn --hidden-import comtypes._post_coinit.misc --icon audio.ico --add-data "audio.ico;." --version-file version.txt .\audioctl.py
```
Notes:
- Ensure `audio.ico` and `version.txt` are present.
- The build collects `pycaw` and ensures `comtypes.automation` is embedded for the shim.

Version resource (example for v1.4.3.1):
```text
VSVersionInfo(
  ffi=FixedFileInfo(
    filevers=(1, 4, 3, 1),
    prodvers=(1, 4, 3, 1),
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
          StringStruct('FileVersion', '1.4.3.1'),
          StringStruct('InternalName', 'audioctl'),
          StringStruct('LegalCopyright', 'Copyright (c) 2025 Mr5niper5oft'),
          StringStruct('OriginalFilename', 'audioctl.exe'),
          StringStruct('ProductName', 'Windows Audio Control CLI'),
          StringStruct('ProductVersion', '1.4.3.1'),
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
`audioctl` v1.4.3.1 provides a dependable, scriptable, and interactive way to manage Windows audio endpoints. It blends `comtypes` convenience with targeted `ctypes` vtable calls for robust Listen/SysFX operations, uses a vendor‑first approach for Enhancements, and offers comprehensive diagnostics and learn tooling. The logging subsystem initializes lazily, with opt‑in debugging, and the packaging flow is streamlined with a single PyInstaller command.
