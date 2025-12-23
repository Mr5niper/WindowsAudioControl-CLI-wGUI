
# Technical Specifications: audioctl.py — Windows Audio CLI/GUI Utility (v1.4.1.0)

## 1. Introduction

`audioctl.py` is a Windows audio control utility offering:
- A command‑line interface (CLI) for scripting, automation, and diagnostics.
- A Tkinter‑based graphical user interface (GUI) for interactive control.

It targets reliable control of Windows Core Audio endpoints (Render/Capture) using COM via `comtypes`, and low‑level `ctypes` where needed. The application is designed to run both from source and as a standalone PyInstaller executable, with care taken to keep COM usage robust and packaging‑friendly.

This document describes the current architecture, technologies, and behavior of version 1.4.1.0. It is a technical guide to the system as it exists in this release (not a change log).

---

## 2. Core Technologies & Dependencies

- Python 3.13.x

- comtypes  
  COM automation from Python, including generated wrappers for Core Audio interfaces.

- ctypes  
  Low‑level FFI for:
  - Raw IPropertyStore vtable calls (Listen/Enhancements paths, robust friendly‑name reads).
  - Direct PROPVARIANT handling and Windows helper calls (PropVariantClear, InitPropVariantFromX).

- pycaw  
  Convenience interfaces for IMMDevice, IAudioEndpointVolume, and utilities used for enumeration helpers and diagnostics.

- tkinter  
  GUI (Treeview device listing, context menus, status UI).

- argparse, json, re, os, sys, time, io, warnings  
  CLI parsing, structured output, matching, and utilities.

- winreg  
  Windows registry access (MMDevices keys for verification, diagnostics, and vendor toggles).

- Windows Core Audio COM APIs  
  - IMMDeviceEnumerator, IMMDevice
  - IAudioEndpointVolume
  - IPropertyStore
  - PolicyConfig (Vista/Client variants, Fx‑capable for SysFX and SetDefault)

- PyInstaller  
  Produces a single‑file executable, integrates version metadata and icon (via `.spec`).

- External Assets
  - audio.ico: application/executable icon
  - version.txt: PE version resource for the built executable

---

## 3. Dependency Management (Virtual Environment)

Use a dedicated virtual environment:

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

c) Capture the environment:
```bash
pip freeze > requirements.txt
```

Example requirements for this build:
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

Install on a new machine:
```bash
pip install -r requirements.txt
```

Note: exact versions depend on the build environment used for v1.4.1.0.

---

## 4. Architectural Overview

The application is a Python package (`audioctl/`) with:

- Entry points
  - audioctl.py (console entry to CLI/GUI)
  - audioctl/__main__.py (package entry)

- Core modules
  - audioctl/cli.py: CLI parser and command handlers
  - audioctl/gui.py: Tkinter GUI (device listing and actions)
  - audioctl/devices.py: Core Windows audio operations (COM/ctypes/pycaw)
  - audioctl/vendor_db.py: Vendor Enhancements toggle logic and “learn” helpers
  - audioctl/compat.py: comtypes/Windows compatibility shims, constants
  - audioctl/logging_setup.py: logging subsystem, fault/exit hooks, optional debug

Behavior:
- If launched without arguments, the GUI starts by default.
- With arguments, the CLI executes the requested command.

Key design points:
- COM is initialized at process entry (CLI and GUI) and released at exit.
- Sensitive COM operations use raw `ctypes` vtable calls where that’s proved more robust (Listen, PropertyStore, SysFX operations).
- CLI and GUI share core logic; the same enumeration and selection/order rules are applied (GUI‑order indices, consistent naming).

---

## 5. Key Design Patterns & Rationale

### 5.1. COM Interop Strategy
- `comtypes` for high‑level convenience where stable (e.g., IAudioEndpointVolume).
- `ctypes` raw vtable calls for critical operations (IPropertyStore Set/Get, PROPVARIANT) to:
  - Avoid dynamic wrapper friction when packaged.
  - Precisely control memory and lifetime.

COM lifecycle:
- `CoInitialize()` at CLI/GUI startup; `CoUninitialize()` at shutdown.
- Core helpers assume COM is already initialized by the caller.

### 5.2. Compatibility Shim (compat.py)
- Ensures `comtypes.automation` exposes:
  - PROPVARIANT alias to tagPROPVARIANT when missing.
  - VT_* constants (VT_LPWSTR, VT_BOOL, VT_UI2, VT_UI4).
  - VARIANT_TRUE/FALSE constants.
- Provides constants for endpoint flows (Render/Capture), roles, device states, and STGM flags.

### 5.3. Device Enumeration and Naming
Enumeration:
- Uses `IMMDeviceEnumerator.EnumAudioEndpoints` for Render and Capture, honoring:
  - `DEVICE_STATE_ACTIVE` (operations)
  - `DEVICE_STATE_ALL` (list `--all`)

Naming:
- `_safe_friendly_name_from_device(dev)` obtains `PKEY_Device_FriendlyName` (fallback `PKEY_Device_DeviceDesc`) via raw IPropertyStore `GetValue`, with:
  - Explicit PROPVARIANT clear
  - Balanced AddRef/Release on IPropertyStore
  - A short GC guard during the direct vtable call (to avoid destructor‑time Release races)

Ordering:
- Devices are sorted by friendly name within each flow (“GUI order”); this order is mirrored by CLI selection to keep indices consistent across UI/CLI.

### 5.4. Default Endpoint Management
- PolicyConfig (Fx‑capable) interface is used for `SetDefaultEndpoint`.
- A lightweight singleton of PolicyConfig is kept alive (process‑wide) to avoid repeated create/release cycles.
- GUI and CLI route their internal `_get_policy_config` to this stable provider.
- Roles supported: `console`, `multimedia`, `communications`, and `all` (sets all three).

### 5.5. “Listen to this device” (Capture)
- COM path: raw IPropertyStore `SetValue/Commit` of:
  - `PKEY_LISTEN_ENABLE` (`{24dbb0fc-9311-4b3d-9cf0-18ff155639d4}`, pid 1)
  - `PKEY_LISTEN_PLAYBACKTHROUGH` (same GUID, pid 2) to target a render endpoint or `""` for “Default Playback Device”
- Verification:
  - First via COM (GetValue)
  - Fallback via registry polling (`HKCU\...\MMDevices\Audio\Capture\{endpointGUID}\(FxProperties|Properties)`)
  - Boolean parsing supports `REG_DWORD`, `REG_BINARY` (PROPVARIANT), `REG_SZ`

### 5.6. Volume and Mute (Render/Capture)
- `IAudioEndpointVolume` via `IMMDevice.Activate`.
- Get/Set guard both tuple returns and out‑params (ctypes.byref fallbacks), producing integer % 0..100 and boolean mute values.

### 5.7. Enhancements (SysFX) — Vendor‑Aware Toggle & Learn (overview)
- Live status readers:
  - PropertyStore (Disable_SysFx via IPropertyStore)
  - PolicyConfig (Fx and normal stores) for Disable_SysFx
- Vendor toggles:
  - INI‑driven (user‑learned) and built‑in vendor entries (e.g., Realtek/Waves)
  - Toggle writes DWORDs in endpoint `FxProperties/Properties` under HKCU/HKLM as configured
- Learn modes: see Section 10 for full details (manual learn and discovery, INI format, and runtime application)

### 5.8. Logging, Diagnostics, and Debug
- `audioctl_gui.log` written next to the executable (fallback `%TEMP%\audioctl\audioctl_gui.log`)
- Global hooks:
  - `sys.excepthook`, `sys.unraisablehook`
  - `faulthandler` (where available)
  - `atexit` exit breadcrumbs
  - `sys.exit` wrapper logs exit codes
  - Console control handler logs Ctrl+C/Close/etc.
- Optional debug logging (`_dbg`):
  - Enabled via `AUDIOCTL_DEBUG=1` environment variable
  - Emits structured debug lines (timestamp, pid, tid) around COM operations, refresh, and actions

### 5.9. Stability Guards
- PolicyConfig kept as a singleton to reduce COM churn.
- Friendly‑name read uses AddRef/Release and a short GC guard during low‑level IPropertyStore calls.
- GUI and CLI wrap long enumerations with brief GC pauses to avoid destructor‑time Release races on sensitive drivers.

---

## 6. CLI Command Reference

Device selection semantics:
- Operates on active endpoints unless stated otherwise.
- When selecting by name and multiple matches exist, `--index` is interpreted in the same per‑flow order as the GUI (name‑sorted within each flow).
- `--regex` switches name matching to regular expressions.

- list  
  Lists devices in GUI order, optionally including disabled/disconnected and/or JSON output.
  ```bash
  audioctl list [--all] [--json]
  ```

- set-default  
  Sets default playback/recording device for one or more roles. Device must be active.
  ```bash
  audioctl set-default
    (--playback-id <ID> | --playback-name <NAME>) [--playback-role <ROLE>]
    (--recording-id <ID> | --recording-name <NAME>) [--recording-role <ROLE>]
    [--index <N>] [--regex]
  ```
  Roles: `console`, `multimedia`, `communications`, `all` (sets all three)

- set-volume  
  Sets master volume (0–100) or mutes/unmutes the target device (render or capture).
  ```bash
  audioctl set-volume
    (--id <ID> | --name <NAME>) [--flow <Render|Capture>]
    (--level <0..100> | --mute | --unmute)
    [--index <N>] [--regex]
  ```

- listen  
  Enables/disables “Listen to this device” for an active capture device; optional playback target endpoint ID or `""` for “Default Playback Device”.
  ```bash
  audioctl listen
    (--id <ID> | --name <NAME>)
    (--enable | --disable)
    [--playback-target-id <RenderEndpointID-or-empty>]
    [--index <N>] [--regex]
  ```

- enhancements — vendor toggles and learn  
  Enables/disables Enhancements (SysFX) using vendor toggles (if configured), or runs a manual learn to append an INI vendor entry for this endpoint.
  ```bash
  # Enable/Disable via vendor toggles
  audioctl enhancements
    (--id <ID> | --name <NAME>) [--flow <Render|Capture>]
    (--enable | --disable)
    [--index <N>] [--regex]
    [--prefer-hklm] [--vendor-ini <path>]

  # Manual learn (append an INI vendor entry)
  audioctl enhancements
    (--id <ID> | --name <NAME>) [--flow <Render|Capture>]
    --learn
    [--index <N>] [--regex]
    [--vendor-ini <path>]
  ```
  Notes:
  - `--learn` guides you to toggle Windows’ Enhancements UI, captures A/B, and appends a vendor section to `vendor_toggles.ini` (or `--vendor-ini` path).
  - After learn, `--enable/--disable` will write the device‑specific DWORD across the configured hives and verify.

- diag-sysfx (diagnostic)  
  Dumps Enhancements (SysFX) status from:
  - PropertyStore (live)
  - PolicyConfig (Fx/normal stores)
  - First applicable vendor entry (if any)
  ```bash
  audioctl diag-sysfx
    (--id <ID> | --name <NAME>) [--flow <Render|Capture>]
    [--index <N>] [--regex]
  ```

- diag-mmdevices (diagnostic)  
  Dumps all MMDevices registry values for an endpoint under HKCU/HKLM (FxProperties/Properties), useful for debugging drivers and Learn results.
  ```bash
  audioctl diag-mmdevices
    (--id <ID> | --name <NAME>) [--flow <Render|Capture>]
    [--index <N>] [--regex]
  ```

- discover-enhancements (diagnostic + learn helper)  
  Interactive A/B capture while you toggle Enhancements in Windows UI; saves TXT/JSON report and can write a suggested INI section.
  ```bash
  audioctl discover-enhancements
    (--id <ID> | --name <NAME>) [--flow <Render|Capture>]
    [--index <N>] [--regex]
    [--output-dir <path>] [--ini-snippet <path>]
  ```

- wait (diagnostic/automation)  
  Waits for a device to appear (become active) within timeout and prints its descriptor.
  ```bash
  audioctl wait
    (--id <ID> | --name <NAME>) [--flow <Render|Capture>]
    [--timeout <seconds>] [--index <N>] [--regex]
  ```

---

## 7. GUI Application Structure

The `AudioGUI` class provides:
- A Treeview listing in two groups: “Playback (Render)” and “Recording (Capture)”.
- Per‑flow name‑sorted ordering, consistent with CLI indexing.
- Context menu actions:
  - Set as Default (all roles)
  - Set Volume…
  - Mute/Unmute
  - Toggle Listen (capture only)
  - Enable/Disable Enhancements (when a vendor toggle is available)
  - Learn Enhancements (guided vendor learning; appends to INI — see Section 10)
- View options:
  - Show disabled/disconnected endpoints
  - Print CLI commands (echo the exact CLI for any action)
- Status bar updates and guarded error dialogs.

Stability and behavior notes:
- COM is initialized at startup and uninitialized on exit.
- Device refresh and critical COM vtable operations are guarded (short GC pauses, AddRef/Release balance) to avoid destructor‑time Release races on sensitive drivers.
- Icons for the root and dialogs are resolved via `resource_path("audio.ico")`, working both from source and in a frozen build.

---

## 8. Deployment Considerations (PyInstaller)

### 8.1. Build

Invoke:

```powershell
pyinstaller -F --noupx --clean --onefile --console --name audioctl --collect-all pycaw --hidden-import comtypes.automation --icon audio.ico --add-data "audio.ico;." --version-file version.txt .\audioctl.py
```

### 8.2. Version File (`version.txt`)

PE version resource for v1.4.1.0, example:

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
- `FileVersion`/`ProductVersion` reflect the application’s release (1.4.1.0).
- Language table `040904E4` corresponds to U.S. English, Unicode.

### 8.3. Output

On success:
- `dist\audioctl.exe` contains the single‑file binary.
- `build\` contains intermediates for inspection.

---

## 9. Diagnostics & Debug Tools

- AUDIOCTL_DEBUG (environment variable)  
  Set to `1` to enable verbose debug lines in `audioctl_gui.log`.
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

- diag-sysfx (CLI)  
  Dumps Enhancements state from PropertyStore, PolicyConfig (Fx/normal), and vendor toggle status. Useful to see “live” Windows state vs. vendor‑mapped state.

- diag-mmdevices (CLI)  
  Dumps all MMDevices values (HKCU/HKLM) under FxProperties/Properties for an endpoint to help identify vendor toggles or diagnose driver behavior.

- discover-enhancements (CLI)  
  Guided A/B snapshot and diff for Enhancements. Produces TXT/JSON bundle and can write a suggested INI section.

- wait (CLI)  
  Polls for device appearance; useful for scripting sequences where a device is hot‑plugged.

- GUI “Print CLI commands”  
  Option to echo exact CLI equivalents of actions taken via GUI.

---

## 10. Learn: Vendor Enhancements Discovery, INI, and Runtime Behavior

### 10.1. Purpose

Some drivers (e.g., Realtek/Waves) handle “Audio Enhancements” via vendor‑specific registry values rather than Windows’ Disable_SysFx property alone. Learn captures how your device actually toggles Enhancements and stores a small, safe mapping so future enable/disable operations work reliably on that device.

### 10.2. Where data is stored

- Default INI path: `vendor_toggles.ini` next to the executable (or package root when running from source).
- Override with `--vendor-ini <path>` (CLI). GUI Learn writes to the default path.

### 10.3. INI format

```
[vendor_{guid},pid]
value_name = {GUID},pid
dword_enable = 0
dword_disable = 1
hives = HKCU,HKLM
flows = Render,Capture
notes = Context about the device/flow and capture (optional)
```

- `value_name`: An MMDevices FxProperties/Properties key (e.g. `{1da5d803-d492-4edd-8c23-e0c0ffee7f0e},5`)
- `dword_enable`/`dword_disable`: DWORD values used for enable/disable.
- `hives`: Where to write (order is the precedence).
- `flows`: `Render`, `Capture`, or both.

### 10.4. How to Learn

- CLI manual learn:
  ```bash
  audioctl enhancements --id "<endpoint-id>" --flow Render --learn [--vendor-ini "<path>"]
  ```
  The tool prompts you to toggle Windows’ Enhancements, captures A/B snapshots, diffs MMDevices, and appends a vendor section if a single DWORD flip candidate is found.

- GUI learn:
  - Right‑click a device → “Learn Enhancements”
  - Follow prompts; the section is appended to the INI if a good candidate is found.

- CLI guided discovery (report + optional INI snippet):
  ```bash
  audioctl discover-enhancements \
    --id "<endpoint-id>" --flow Render \
    --output-dir "C:\out" \
    --ini-snippet "C:\path\to\vendor_toggles.ini"
  ```

Safety note:
- Learn persists a vendor mapping. Subsequent `enhancements --enable/--disable` will write this DWORD for the endpoint. Remove the INI section to undo.

### 10.5. Runtime toggle behavior

- On `audioctl enhancements --enable/--disable` (or GUI toggle):
  1) The first applicable vendor entry (INI first, then built‑in code vendors) is selected for the endpoint/flow.
  2) The tool writes the corresponding DWORD to the configured hives.
  3) It verifies by reading back the same entry (with short retries for consistency).

- Diagnostics:
  - `audioctl diag-sysfx` shows:
    - PropertyStore Enhancements state (live)
    - PolicyConfig (Fx and normal) state
    - Vendor toggle status (the applied vendor entry’s current reading)

### 10.6. Permissions and precedence

- HKCU writes: no admin required.
- HKLM writes: admin required.
- Write order follows the INI’s `hives` field for vendor entries set during Learn.

### 10.7. Example manual INI entry

```
[vendor_{1da5d803-d492-4edd-8c23-e0c0ffee7f0e},5]
value_name = {1da5d803-d492-4edd-8c23-e0c0ffee7f0e},5
dword_enable = 0
dword_disable = 1
hives = HKCU,HKLM
flows = Render,Capture
notes = Learned on "Speakers (Realtek(R) Audio)" (A=enabled, B=disabled)
```

After saving:
```bash
audioctl enhancements --name "Speakers" --flow Render --disable
```

### 10.8. Troubleshooting Learn

- No DWORD flip found:
  - Driver may use non‑DWORD or binary blob. Keep the TXT/JSON from `discover-enhancements` for deeper analysis.
- Verification timing:
  - Some drivers update asynchronously; the tool’s verification loop already retries briefly.

---

## 11. Conclusion

`audioctl.py` (v1.4.1.0) provides a robust, scriptable and interactive way to manage Windows audio endpoints. Its COM interaction approach blends `comtypes` convenience with targeted `ctypes` vtable calls for reliability in sensitive areas, and it is structured to work consistently from source and as a frozen executable. Diagnostics and optional debug logging make it practical to deploy and support in varied Windows environments, while the GUI and CLI share a consistent selection model and device ordering. Learn mode enables vendor‑aware Enhancements toggling that is reliable across different driver implementations.
