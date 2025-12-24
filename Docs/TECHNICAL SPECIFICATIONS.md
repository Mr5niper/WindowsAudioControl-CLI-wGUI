# Technical Specifications: audioctl (Windows Audio CLI/GUI Package)

## 1. Introduction
`audioctl` is a Python package that provides comprehensive control over Windows audio devices. It offers:
- a robust command-line interface (CLI) for scripting and automation, and
- a user-friendly graphical user interface (GUI) built with Tkinter for interactive management.

The package leverages low-level Windows COM APIs via `ctypes` and `comtypes` to achieve reliable, precise audio manipulation in both source and packaged (PyInstaller) deployments.

Core capabilities:
- enumerate playback/recording endpoints and surface human‑friendly names,
- set default endpoints by role,
- read/set volume and mute,
- enable/disable “Listen to this device” on capture devices,
- enable/disable “Audio Enhancements” (SysFX) using a vendor‑first approach with a learnable INI database,
- diagnostics and discovery tooling for Enhancements behavior.

This document describes what the tool is and how it works (architecture, dependencies, deployment). It is not a changelog.

---

## 2. Core Technologies & Dependencies
- Python 3.x
- comtypes  
  Python wrapper for Microsoft COM interfaces.
- ctypes  
  Low‑level FFI used where direct control and robustness are required (e.g., raw IPropertyStore vtable calls).
- pycaw  
  Convenience interface for Windows Core Audio APIs (MMDevice API, EndpointVolume API).
- tkinter  
  GUI (optional).
- argparse  
  CLI argument parsing.
- json, re  
  Structured output and pattern matching for device selection.
- os, sys, time, warnings, io, contextlib, tempfile, datetime, traceback, atexit, faulthandler, winreg, configparser, threading, gc  
  Standard libraries for system interaction, logging, registry access, vendor INI management, and runtime stability.
- Windows COM APIs  
  - IMMDeviceEnumerator: enumerate endpoints
  - IMMDevice: endpoint objects
  - IAudioEndpointVolume: volume/mute
  - IPropertyStore: device properties (“Listen”, Disable_SysFx)
  - IPolicyConfig/…Vista (+ local variant for SysFX Get/SetPropertyValue): default endpoints and Disable_SysFx reads/writes
- PyInstaller  
  Packaging into a standalone Windows executable.
- External Assets
  - audio.ico: icon for the executable and GUI windows
  - version.txt: executable version metadata
  - vendor_toggles.ini (optional; created/updated at runtime): append‑only database of vendor DWORD toggles for “Audio Enhancements”

---

## 3. Dependency Management (Virtual Environment)
Use a virtual environment (venv) to isolate dependencies.

a) Create:
```bash
python -m venv venv
```

b) Activate:
- Command Prompt:
  ```bat
  .\venv\Scripts\activate.bat
  ```
- PowerShell:
  ```powershell
  .\venv\Scripts\Activate.ps1
  ```

c) Manage packages:
```bash
pip freeze > requirements.txt
pip install -r requirements.txt
```

Sample requirements:
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

Note: The package also relies on standard modules (`configparser`, `winreg`, `faulthandler`, `gc`, etc.).

---

## 4. Architectural Overview
The package is split into cohesive modules:

- audioctl.compat  
  Early, global compatibility/shim utilities (must be imported before `pycaw`/`comtypes`).
- audioctl.devices  
  Core audio operations: enumeration, selection, default routing, volume/mute, “Listen” and SysFX helpers, plus stability measures around COM interactions.
- audioctl.vendor_db  
  Vendor‑first SysFX control (INI file I/O, vendor toggles, verification) and auto/manual learning helpers.
- audioctl.gui  
  Tkinter GUI wrapper (device browser and context menu actions).
- audioctl.cli  
  CLI entry point with commands: list, set-default, set-volume, listen, enhancements, diag-sysfx, diag-mmdevices, discover-enhancements, wait.
- audioctl.logging_setup  
  Centralized logging, crash handling, optional debug trace helpers.
- audioctl.py / audioctl.__main__.py  
  Simple bootstraps that call `audioctl.cli.main()`.

Behavior:
- With no arguments, the GUI launches; otherwise, the requested CLI command executes.

---

## 5. Key Design Patterns & Rationale

### 5.1. COM Interop, GC Management, and Stability
- comtypes for convenience (e.g., IAudioEndpointVolume); direct `ctypes` vtable calls for critical functionality (IPropertyStore) to avoid dynamic codegen/bundling pitfalls.
- Explicit CoInitialize/CoUninitialize at CLI/GUI entry points.
- Garbage collector (GC) is disabled during critical sections (e.g., for the duration of the GUI and the CLI run) to prevent `comtypes` finalizers from running `Release()` while raw pointers are in use. GC is re‑enabled on shutdown/end of CLI. Some hot paths (like device refresh) temporarily pause/resume GC.

### 5.2. comtypes Compatibility Shim (audioctl.compat)
A minimal shim ensures the presence of:
- PROPVARIANT (aliased to `tagPROPVARIANT` if necessary),
- VT_LPWSTR, VT_BOOL, VT_UI2, VT_UI4,
- VARIANT_TRUE/VARIANT_FALSE.
This stabilizes packaged deployments where `comtypes.automation` may not export all elements.

### 5.3. Device Enumeration and Friendly Names
- `list_devices()` uses IMMDeviceEnumerator.EnumAudioEndpoints() (avoiding fragile release paths).
- Friendly names are resolved primarily via a one‑time `pycaw.AudioUtilities.GetAllDevices()` name map (preferred for stability). A raw IPropertyStore path is retained as a hardened fallback, with explicit AddRef/Release and GC pausing if needed.

### 5.4. “Listen to this device”
- `set_listen_to_device_ps()` writes Listen enable and playback‑through target via raw IPropertyStore calls; `Commit` persists changes.
- `_get_listen_to_device_status_ps()` reads Listen enable via IPropertyStore.
- Registry verification probes `HKCU\...\MMDevices\Audio\Capture\{GUID}\(FxProperties|Properties)` for the AudioEndpointSettings GUID and supports REG_DWORD, REG_BINARY (PROPVARIANT), and REG_SZ.

### 5.5. Default Endpoint Management
- `_get_policy_config()` obtains an IPolicyConfig/…Vista interface via layered fallbacks.
- A stable, lazy singleton of the PolicyConfigFx client is used during runtime to keep a single COM object alive (both GUI and CLI route to this singleton).
- `set_default_endpoint()` sets defaults per role (or all roles) for active devices; clearer per‑role status reporting.

### 5.6. Volume and Mute Control
- `set_endpoint_mute()`, `get_endpoint_mute()`, `set_endpoint_volume()`, `get_endpoint_volume()` operate on `IAudioEndpointVolume` and handle comtypes out‑param/tuple variants.

### 5.7. Audio Enhancements (SysFX) — Vendor‑First Strategy
- Read paths:
  - PropertyStore read of `PKEY_AudioEndpoint_Disable_SysFx` (parsing VT_BOOL/UI2/UI4),
  - PolicyConfig COM reads for both FX and normal stores.
- Control path:
  - Vendor DWORD toggles under `...MMDevices\Audio\{Render|Capture}\{GUID}\FxProperties` using:
    - a small code‑maintained vendor list (e.g., Realtek/Waves), and
    - a user‑maintained, append‑only INI (vendor_toggles.ini) learned on your machine.
  - HKLM writes require Administrator.
- Learning/Discovery:
  - Manual learn (GUI/CLI) captures A/B snapshots and writes an INI entry when a reliable REG_DWORD flip is found under FxProperties.
  - Discovery (CLI) produces TXT/JSON reports and can append a suggested INI snippet.
- Policy:
  - Runtime Enhancements toggling is vendor‑only (no Windows fallback). If no vendor entry applies, the command fails with guidance to learn first.

### 5.8. Logging and Diagnostics
- Centralized log at `audioctl_gui.log` next to the executable (or temp fallback).
- Uncaught and unraisable exceptions, fault handler, atexit hooks, sys.exit wrapper, and Windows console control handlers provide robust diagnostics.
- Optional debug tracing (pid/tid) controlled by `AUDIOCTL_DEBUG=1` environment variable.

---

## 6. CLI Command Reference
Selection supports `--id` or `--name`, with consistent `--index` disambiguation by GUI order (name‑sorted within each flow: Render, Capture). Mutating commands operate only on active endpoints.

- list  
  ```bash
  audioctl list [--all] [--json]
  ```

- set-default  
  ```bash
  audioctl set-default
    (--playback-id <ID> | --playback-name <NAME>) [--playback-role <ROLE>]
    (--recording-id <ID> | --recording-name <NAME>) [--recording-role <ROLE>]
    [--index <N>] [--regex]
  ```
  Roles: console, multimedia, communications, all.

- set-volume  
  ```bash
  audioctl set-volume
    (--id <ID> | --name <NAME>) [--flow <Render|Capture>]
    (--level <0-100> | --mute | --unmute)
    [--index <N>] [--regex]
  ```

- listen  
  ```bash
  audioctl listen
    (--id <ID> | --name <NAME>)
    (--enable | --disable)
    [--playback-target-id <ID|''>] [--index <N>] [--regex]
  ```

- enhancements  
  ```bash
  audioctl enhancements
    (--id <ID> | --name <NAME>) [--flow <Render|Capture>]
    (--enable | --disable | --learn)
    [--index <N>] [--regex]
    [--prefer-hklm] [--vendor-ini <PATH>]
  ```

- diag-sysfx  
  ```bash
  audioctl diag-sysfx
    (--id <ID> | --name <NAME>) [--flow <Render|Capture>] [--index <N>] [--regex]
  ```

- diag-mmdevices  
  ```bash
  audioctl diag-mmdevices
    (--id <ID> | --name <NAME>) [--flow <Render|Capture>] [--index <N>] [--regex]
  ```

- discover-enhancements  
  ```bash
  audioctl discover-enhancements
    (--id <ID> | --name <NAME>) [--flow <Render|Capture>] [--index <N>] [--regex]
    [--output-dir <DIR>] [--ini-snippet <PATH>]
  ```

- wait  
  ```bash
  audioctl wait
    (--id <ID> | --name <NAME>)
    [--flow <Render|Capture>] [--timeout <SECONDS>] [--index <N>] [--regex]
  ```

---

## 7. GUI Application Structure
The Tkinter GUI provides an interactive view with context menu actions.

- Window title reflects the current version (e.g., “Audio Control v1.4.1.0 12-22-2025”).
- Device list is grouped (Playback/Recording), name‑sorted, and aligned with CLI indexing.
- Context menu actions:
  - Set as Default (all roles),
  - Set Volume,
  - Mute/Unmute,
  - Toggle Listen (capture only),
  - Enable/Disable Enhancements (vendor‑first; disabled if unavailable),
  - Learn Enhancements (guided manual learning).
- “Print CLI commands” echoes the equivalent CLI for actions performed in the GUI.
- Stability measures:
  - GC is disabled for the GUI lifetime; certain operations (like refresh) temporarily pause/resume COM‑sensitive work; GC is re‑enabled on shutdown.

---

## 8. Deployment (PyInstaller)

```powershell
pyinstaller -F --noupx --clean --onefile --console --name audioctl --collect-all pycaw --hidden-import comtypes.automation --icon audio.ico --add-data "audio.ico;." --version-file version.txt .\audioctl_1.3.0.17.py
```

Version file (example for 1.4.1.0):
```text
VSVersionInfo(
  ffi=FixedFileInfo(
    filevers=(1, 4, 1, 0),
    prodvers=(1, 4, 1, 0),
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
          StringStruct('FileVersion', '1.4.1.0'),
          StringStruct('InternalName', 'audioctl'),
          StringStruct('LegalCopyright', 'Copyright (c) 2025 Mr5niper5oft'),
          StringStruct('OriginalFilename', 'audioctl.exe'),
          StringStruct('ProductName', 'Windows Audio Control CLI'),
          StringStruct('ProductVersion', '1.4.1.0'),
        ]
      )
    ]),
    VarFileInfo([VarStruct('Translation', [1033, 1252])])
  ]
)
```

Notes:
- `vendor_toggles.ini` is created/updated at runtime next to the executable; you may ship a pre‑populated file via `--add-data "vendor_toggles.ini;."` if desired.

---

## 9. Conclusion
`audioctl` provides dependable Windows audio control via a cohesive CLI/GUI interface, robust COM interop (with targeted `ctypes` vtable usage and a `comtypes` compatibility shim), and runtime stability measures (GC control, a stable PolicyConfig singleton). The vendor‑first Enhancements (SysFX) system is learnable (`vendor_toggles.ini`) and diagnostics‑friendly, while logging and error handling aid production troubleshooting. The modular package layout supports maintainability and consistent packaging.
```
