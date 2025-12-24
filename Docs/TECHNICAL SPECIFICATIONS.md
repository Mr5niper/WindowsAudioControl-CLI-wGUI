# Technical Specifications: audioctl.py — Windows Audio CLI Utility

## 1. Introduction
`audioctl.py` is a versatile Python script that provides comprehensive control over Windows audio devices. It offers both:
- a robust command-line interface (CLI) for scripting and automation, and
- a user-friendly graphical user interface (GUI) built with Tkinter for interactive management.

The script leverages low-level Windows COM APIs through `ctypes` and `comtypes` to achieve reliable and precise audio manipulation, addressing common challenges and stability issues encountered with higher-level abstractions, especially when packaged as a standalone executable.

Core capabilities include:
- listing and selecting playback/recording endpoints,
- setting default endpoints by role,
- reading/setting volume and mute,
- enabling/disabling “Listen to this device” on capture devices,
- enabling/disabling “Audio Enhancements” (SysFX) via a vendor-first strategy,
- diagnostics and discovery tooling for Enhancements behavior.

This document describes what the tool is and how it works, including its architecture, dependencies, and deployment considerations.

---

## 2. Core Technologies & Dependencies
- Python 3.x  
  The primary programming language.
- comtypes  
  Python wrapper for Microsoft COM interfaces.
- ctypes  
  Low-level FFI used where direct control and robustness are required (e.g., IPropertyStore vtable calls).
- pycaw  
  Convenience interface for Windows Core Audio APIs (MMDevice API, EndpointVolume API); some operations are implemented directly for robustness.
- tkinter  
  Standard GUI library (optional graphical interface).
- argparse  
  CLI argument parsing.
- json  
  Structured output for CLI and data exchange.
- re  
  Regular expressions for device name matching.
- os, sys, time, warnings, io, contextlib, tempfile, datetime, traceback, atexit, faulthandler, winreg, configparser  
  Standard libraries used for system interaction, logging, registry access, and vendor INI management.
- Windows COM (Component Object Model) APIs  
  - IMMDeviceEnumerator: enumerate audio endpoints
  - IMMDevice: individual endpoints
  - IAudioEndpointVolume: volume/mute control
  - IPropertyStore: device properties (e.g., “Listen”, Disable_SysFx)
  - IPolicyConfigVista / IPolicyConfig (plus a local interface for SysFX Get/SetPropertyValue): default devices, property reads/writes
- PyInstaller  
  Packaging into a standalone Windows executable (see Section 8).
- External Assets
  - audio.ico: icon for the executable and GUI windows
  - version.txt: executable version metadata
  - vendor_toggles.ini (optional, created/updated at runtime): append-only database of vendor DWORD toggles for “Audio Enhancements”

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

Note: The script also relies on standard-library modules such as `configparser`, `winreg`, and `faulthandler`.

---

## 4. Architectural Overview
The script is organized into:

a) Core Audio Logic  
Functions interacting directly with Windows COM APIs to query and control audio devices. Sensitive paths (e.g., IPropertyStore access) rely on raw `ctypes` vtable calls for maximum robustness in both source and packaged runs. This includes:
- endpoint enumeration and naming,
- default device selection,
- volume/mute control,
- “Listen to this device” reads/writes with verification, and
- “Audio Enhancements” (SysFX) vendor-first reads/writes with learnable driver recipes.

b) User Interfaces
- CLI via `argparse`, with commands:
  - list, set-default, set-volume, listen, enhancements, diag-sysfx, diag-mmdevices, discover-enhancements, wait
- GUI via Tkinter:
  - device browser with context menu actions (defaults, volume, mute, listen, enhancements),
  - guided “Learn Enhancements” assistant to produce vendor INI entries.

If run without arguments, the GUI launches; otherwise, the requested CLI command executes.

---

## 5. Key Design Patterns & Rationale

### 5.1. COM Interoperability and Robustness
- comtypes is used for convenience (e.g., IAudioEndpointVolume), but direct `ctypes` vtable calls are used for critical functionalities (IPropertyStore) to avoid dynamic codegen or bundling pitfalls.
- Explicit `CoInitialize`/`CoUninitialize` are called at CLI/GUI entry points. Low-level helpers assume COM is already initialized.

### 5.2. comtypes Compatibility Shim
A minimal shim at the top of the script ensures the presence of:
- PROPVARIANT (aliasing `tagPROPVARIANT` if needed),
- VT_LPWSTR, VT_BOOL, and for SysFX handling also VT_UI2 and VT_UI4,
- VARIANT_TRUE/VARIANT_FALSE constants.

This stabilizes packaging scenarios where `comtypes.automation` may not expose all elements.

### 5.3. Device Enumeration and Friendly Names
- `list_devices()` uses `IMMDeviceEnumerator.EnumAudioEndpoints()` to avoid potential `Release()` issues reported with higher-level helpers.
- `_safe_friendly_name_from_device()` reads `PKEY_Device_FriendlyName` (or falls back to DeviceDesc) via direct IPropertyStore access using `ctypes`, robustly handling PROPVARIANT layouts.

### 5.4. “Listen to this device”
- `set_listen_to_device_ps()`: writes Listen enable and playback-through target via raw IPropertyStore vtable calls (`SetValue`, `Commit`).
- `_get_listen_to_device_status_ps()`: reads Listen enable via `GetValue`.
- Registry verification: probes `HKCU\...\MMDevices\Audio\Capture\{GUID}\(FxProperties|Properties)` to confirm expected state, supporting `REG_DWORD`, `REG_BINARY` (PROPVARIANT), and `REG_SZ`.

### 5.5. Default Endpoint Management
- `_get_policy_config()` locates an `IPolicyConfig`/`IPolicyConfigVista` implementation using multiple strategies (pycaw helpers, `pycaw.policyconfig`, or a local minimal interface).
- `set_default_endpoint()` sets defaults per role (or all roles) for active devices, warning about potential admin requirements.

### 5.6. Volume and Mute Control
- `set_endpoint_mute()`, `get_endpoint_mute()`, `set_endpoint_volume()`, `get_endpoint_volume()` operate on `IAudioEndpointVolume`.
- Return values are handled robustly (direct returns or tuple/out-param variants) to accommodate comtypes behavior.

### 5.7. Audio Enhancements (SysFX) — Vendor-First Strategy
Because audio drivers differ in how they implement Enhancements:
- Read paths:
  - Windows PropertyStore read of `PKEY_AudioEndpoint_Disable_SysFx` (parsing VT_BOOL/UI2/UI4).
  - PolicyConfig COM path (readers) for FX store and normal store.
- Write (control) path:
  - Vendor-first, via DWORD toggles under `...MMDevices\Audio\{Render|Capture}\{GUID}\FxProperties` using:
    - a small list of code-maintained vendor entries, and
    - a user-maintained, append-only `vendor_toggles.ini` (learned on your machine) with entries like:
      ```
      [vendor_section_name]
      value_name = {GUID},pid
      dword_enable = 0
      dword_disable = 1
      hives = HKLM,HKCU
      flows = Render,Capture
      notes = optional
      ```
  - HKLM writes require Administrator rights.
- Learning:
  - Manual (GUI/CLI): guided A/B capture where you toggle Windows settings on prompt; writes an INI entry if a reliable REG_DWORD flip is detected under FxProperties.
  - Discovery (CLI): produces TXT/JSON reports and an optional INI snippet for later inclusion.
- Policy:
  - At runtime, Enhancements toggling is vendor-only (no automatic fallback to Windows paths). If no vendor entry applies, the command fails with guidance to learn first. This avoids ambiguous results across diverse driver implementations.

### 5.8. Error Handling and Logging
- Global `sys.excepthook` and `sys.unraisablehook` write to `audioctl_gui.log`.
- `faulthandler` integrates crash dumps into the same log (all threads).
- `atexit` hooks and a `sys.exit` wrapper add breadcrumbs.
- A console control handler logs control events (Ctrl+C/Break, logoff, shutdown).
- GUI callback errors are logged and shown via user-friendly dialogs.

---

## 6. CLI Command Reference
Selection supports `--id` or `--name`, with consistent `--index` disambiguation by GUI order (name-sorted within each flow: Render, Capture). Mutating commands operate only on active endpoints.

- list  
  ```bash
  audioctl list [--all] [--json]
  ```
  Lists endpoints (active by default, or all with `--all`). Output aligns with GUI order; `--json` emits structured data.

- set-default  
  ```bash
  audioctl set-default
    (--playback-id <ID> | --playback-name <NAME>) [--playback-role <ROLE>]
    (--recording-id <ID> | --recording-name <NAME>) [--recording-role <ROLE>]
    [--index <N>] [--regex]
  ```
  Roles: `console`, `multimedia`, `communications`, `all`. Target must be active.

- set-volume  
  ```bash
  audioctl set-volume
    (--id <ID> | --name <NAME>) [--flow <Render|Capture>]
    (--level <0-100> | --mute | --unmute)
    [--index <N>] [--regex]
  ```
  Adjusts master volume or toggles mute on an active device.

- listen  
  ```bash
  audioctl listen
    (--id <ID> | --name <NAME>)
    (--enable | --disable)
    [--playback-target-id <ID|''>] [--index <N>] [--regex]
  ```
  Enables/disables “Listen to this device” for a capture device. `--playback-target-id ""` selects the Default Playback Device.

- enhancements  
  ```bash
  audioctl enhancements
    (--id <ID> | --name <NAME>) [--flow <Render|Capture>]
    (--enable | --disable | --learn)
    [--index <N>] [--regex]
    [--prefer-hklm]
    [--vendor-ini <PATH>]
  ```
  Vendor-first toggling of “Audio Enhancements”. If unsupported for a device, instructs using `--learn` to create a vendor INI entry. `--prefer-hklm` attempts HKLM writes first (Admin required). `--vendor-ini` sets a custom INI path (default: next to the executable).

- diag-sysfx  
  ```bash
  audioctl diag-sysfx
    (--id <ID> | --name <NAME>) [--flow <Render|Capture>] [--index <N>] [--regex]
  ```
  Prints Enhancements state from Windows (PropertyStore/COM) and from the first applicable vendor entry.

- diag-mmdevices  
  ```bash
  audioctl diag-mmdevices
    (--id <ID> | --name <NAME>) [--flow <Render|Capture>] [--index <N>] [--regex]
  ```
  Dumps all MMDevices values for the endpoint across HKCU/HKLM and FxProperties/Properties (debugging/discovery aid).

- discover-enhancements  
  ```bash
  audioctl discover-enhancements
    (--id <ID> | --name <NAME>) [--flow <Render|Capture>] [--index <N>] [--regex]
    [--output-dir <DIR>] [--ini-snippet <PATH>]
  ```
  Captures A/B snapshots (Enabled vs Disabled) to produce TXT and JSON reports; optionally appends a suggested INI snippet.

- wait  
  ```bash
  audioctl wait
    (--id <ID> | --name <NAME>)
    [--flow <Render|Capture>] [--timeout <SECONDS>] [--index <N>] [--regex]
  ```
  Waits for a device to appear active and prints its details.

---

## 7. GUI Application Structure
The Tkinter-based GUI provides an interactive view of endpoints with a context menu for common actions.

- Initialization  
  Applies an icon, robust exception routing to logs and dialogs, and consistent Treeview styling. The title reflects the current version.

- Device Management  
  `refresh_devices()` populates groups “Playback (Render)” and “Recording (Capture)”. Devices are sorted by name in each group to match CLI indexing behavior.

- Interaction
  - Right-click opens the context menu; double-click also opens it.
  - Selection guards ensure only device rows (not group headers) are selectable.
  - F5 refreshes the list.

- Context Menu Actions
  - Set as Default (all roles): invokes default endpoint logic.
  - Set Volume: opens a small dialog with slider and numeric entry.
  - Mute/Unmute: toggles endpoint mute with safe defaults.
  - Toggle Listen (capture only): toggles via IPropertyStore and verification fallback.
  - Enable/Disable Enhancements: vendor-first toggle of SysFX. If no vendor method applies, the item is disabled.
  - Learn Enhancements: guided manual learning that captures A/B snapshots and appends a reliable DWORD-based vendor entry to `vendor_toggles.ini` when possible.

- Volume Dialog  
  A compact `tk.Toplevel` with validated entry and slider (0–100).

- Resource Paths and Icons  
  `resource_path()` ensures assets (audio.ico) are available in both source and packaged runs.

---

## 8. Deployment Considerations (PyInstaller)
Build command (typical):
```bash
pyinstaller -F --noupx --clean --onefile --console --name audioctl --collect-all pycaw --hidden-import comtypes.automation --icon audio.ico --add-data "audio.ico;." --version-file version.txt .\audioctl_1.3.0.17.py
```

Notes:
- `vendor_toggles.ini` is created/updated at runtime next to the executable; you can optionally ship a pre-populated file via `--add-data "vendor_toggles.ini;."`.

### 8.1. Version File (version.txt)
Use 1.3.0.17:
```text
VSVersionInfo(
  ffi=FixedFileInfo(
    filevers=(1, 3, 0, 17),
    prodvers=(1, 3, 0, 17),
    mask=0x3f,
    flags=0x0,
    OS=0x40004,
    fileType=0x1,
    subtype=0x0,
    date=(0, 0)
  ),
  kids=[
    StringFileInfo([
      StringTable(
        '040904B0',
        [
          StringStruct('CompanyName', ''),
          StringStruct('FileDescription', 'Windows Audio Control CLI'),
          StringStruct('FileVersion', '1.3.0.17'),
          StringStruct('InternalName', 'audioctl'),
          StringStruct('LegalCopyright', ''),
          StringStruct('OriginalFilename', 'audioctl.exe'),
          StringStruct('ProductName', 'Windows Audio Control CLI'),
          StringStruct('ProductVersion', '1.3.0.17'),
        ]
      )
    ]),
    VarFileInfo([VarStruct('Translation', [1033][1200])])
  ]
)
```

### 8.2. Output
The `audioctl.exe` executable is produced under `dist/` with intermediate artifacts under `build/`.

---

## 9. Conclusion
`audioctl.py` is a robust and flexible Windows audio control utility. It pairs dependable COM interop (with targeted `ctypes` vtable usage and a `comtypes` compatibility shim) with a cohesive CLI/GUI feature set. In addition to device enumeration, default selection, volume/mute, and Listen control, it incorporates a vendor-first system for controlling “Audio Enhancements (SysFX)” that is both learnable (via `vendor_toggles.ini`) and verifiable. The error handling and logging stack is designed to aid diagnostics in production environments, and the packaging guidance ensures consistent, reliable deployment.
```
