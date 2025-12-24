# Technical Specifications: audioctl.py — Windows Audio CLI Utility

## 1. Introduction
`audioctl.py` is a versatile Python script designed to provide comprehensive control over Windows audio devices. It offers both a robust command-line interface (CLI) for scripting and automation, and a user-friendly graphical user interface (GUI) built with Tkinter for interactive management. The script leverages low-level Windows COM APIs through `ctypes` and `comtypes` to achieve reliable and precise audio manipulation, addressing common challenges and stability issues encountered with higher-level abstractions, especially when packaged as a standalone executable.

New in 1.3.0.17:
- Audio Enhancements (SysFX) management via a vendor-first strategy with an external, append-only vendor INI database.
- Interactive discovery and diagnostics tools to identify how a device’s “Audio Enhancements” toggle is implemented by its driver.
- New CLI commands: enhancements, diag-sysfx, diag-mmdevices, discover-enhancements.
- GUI context menu additions: Enable/Disable Enhancements and Learn Enhancements.

This document details the script's core functionalities, underlying technologies, architectural decisions, and specific deployment considerations using PyInstaller.

---

## 2. Core Technologies & Dependencies
- Python 3.x  
  The primary programming language.
- comtypes  
  A Python package providing a lightweight Python wrapper around Microsoft COM interfaces. It simplifies interaction with many Windows APIs.
- ctypes  
  Python's foreign function interface library, used here for direct low-level interaction with Windows DLLs and COM interfaces, particularly where comtypes abstractions proved insufficient or for improved robustness.
- pycaw  
  A Python library building on comtypes to provide a more convenient interface for Windows Core Audio APIs (MMDevice API, EndpointVolume API). Note that this script implements several functionalities around or instead of pycaw’s direct methods due to stability and robustness concerns with comtypes in certain deployment scenarios (e.g., PyInstaller).
- tkinter  
  Python’s standard GUI library, used for the optional graphical interface.
- argparse  
  Python's standard library for parsing command-line arguments.
- json  
  For structured output in CLI and data exchange.
- re  
  Regular expressions for device name matching.
- os, sys, time, warnings, io, contextlib, tempfile, datetime, traceback, atexit, faulthandler, winreg, configparser  
  Standard Python libraries and Windows-specific modules for system interaction, logging, vendor INI management, and error handling.
- Windows COM (Component Object Model) APIs  
  The underlying technology used, including:
  - IMMDeviceEnumerator: For enumerating audio endpoint devices.
  - IMMDevice: For representing individual audio endpoint devices.
  - IAudioEndpointVolume: For controlling volume and mute state.
  - IPropertyStore: For reading/writing device properties, notably "Listen to this device" and reading Disable_SysFx.
  - IPolicyConfigVista / IPolicyConfig (and a local IPolicyConfig variant used for SysFX Get/SetPropertyValue): For setting default audio endpoints and reading/writing Disable_SysFx in FX/normal stores.
- PyInstaller  
  The tool used to package the script into a standalone Windows executable. Its specific configuration is detailed in Section 8.
- External Assets
  - audio.ico: The icon file used for the generated executable and Tkinter windows.
  - version.txt: A text file containing version information for the executable's metadata.
  - vendor_toggles.ini (optional, created/updated at runtime): An append-only database of learned vendor DWORD toggles used to control “Audio Enhancements” reliably across driver implementations.

---

## 3. Dependency Management (Virtual Environment)
The development and deployment of this script utilize a Python virtual environment (venv). This practice ensures that project dependencies are isolated from the system-wide Python installation, preventing conflicts and ensuring consistent environments across development, testing, and deployment.

To set up the environment and install dependencies:

a) Create a virtual environment (if not already done):
```bash
python -m venv venv
```

b) Activate the virtual environment:
- Windows (Command Prompt):
  ```bat
  .\venv\Scripts\activate.bat
  ```
- Windows (PowerShell):
  ```powershell
  .\venv\Scripts\Activate.ps1
  ```

c) Install dependencies:
To get a precise list of all installed packages and their versions within your active virtual environment, use the following command to generate a requirements.txt file:
```bash
pip freeze > requirements.txt
```

As of this document the requirements are:
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

To install these dependencies on another system or fresh environment:
```bash
pip install -r requirements.txt
```

Note: The exact list of packages and their versions was generated from a specific venv using pip freeze. The script additionally uses standard-library modules like configparser.

---

## 4. Architectural Overview
The script is structured into two main parts:

a) Core Audio Logic  
A set of functions that interact directly with Windows COM APIs to query and control audio devices. These functions are designed to be robust and handle potential comtypes/ctypes interactions gracefully. In 1.3.0.17 this includes a new “Audio Enhancements (SysFX)” subsystem implemented via:
- A vendor-first toggle strategy based on driver-specific DWORDs under MMDevices FxProperties.
- A learnable, append-only vendor INI database (vendor_toggles.ini).
- Diagnostic helpers that read Windows’s Disable_SysFx through PolicyConfig and PropertyStore for cross-checks.

b) User Interfaces
- Command-Line Interface (CLI)  
  Implemented using argparse, providing distinct commands (list, set-default, set-volume, listen, enhancements, diag-sysfx, diag-mmdevices, discover-enhancements, wait).
- Graphical User Interface (GUI)  
  A Tkinter application that provides an interactive way to view devices and perform common actions, invoking the same core audio logic. New context menu items provide Enhancements control and Learn Enhancements.

The script defaults to launching the GUI if no command-line arguments are provided; otherwise, it executes the specified CLI command.

---

## 5. Key Design Patterns & Rationale

### 5.1. COM Interoperability and Robustness
Interfacing with Windows COM APIs from Python, especially when bundled with PyInstaller, can be challenging due to comtypes sometimes exhibiting unexpected behavior or missing definitions. The script employs a layered approach:
- comtypes as Primary  
  For most standard operations (device enumeration, IAudioEndpointVolume), comtypes is used for its convenience and Pythonic wrappers.
- ctypes for Critical Paths  
  For highly sensitive or problematic areas, particularly the "Listen to this device" functionality and robust property store access (IPropertyStore), the script uses direct ctypes VTABLE calls.
- Rationale  
  This bypasses potential comtypes.gen issues (especially in bundled executables) and ensures maximum control and stability by manually defining COM interface structures and method prototypes. This is a critical design choice to enhance reliability.
- CoInitialize / CoUninitialize  
  Called explicitly in both the GUI and the CLI entry points. Low-level helpers assume COM is already initialized by their caller.

### 5.2. comtypes Compatibility Shim
A minimal shim is placed at the very top of the script.
- Purpose  
  To address known PyInstaller bundling issues where comtypes.automation might not correctly expose PROPVARIANT or standard VT_ constants (like VT_LPWSTR, VT_BOOL, VARIANT_TRUE/FALSE). In 1.3.0.17 the shim also ensures VT_UI2 and VT_UI4 are available for SysFX value handling.
- Mechanism  
  It checks for the existence of these attributes and defines them with their known integer values if missing. It also ensures _automation.PROPVARIANT points to _automation.tagPROPVARIANT if the former is not directly present.

### 5.3. Robust Device Enumeration and Naming
- list_devices()  
  Avoids AudioUtilities.GetAllDevices() due to reported comtypes Release() crashes in certain scenarios. Instead, it directly uses IMMDeviceEnumerator.EnumAudioEndpoints().
- _safe_friendly_name_from_device(dev)  
  Retrieves the device's friendly name (PKEY_Device_FriendlyName) or device description (PKEY_Device_DeviceDesc) via direct IMMDevice.OpenPropertyStore() and IPropertyStore::GetValue() calls using ctypes.

### 5.4. "Listen to this Device" Implementation (Advanced ctypes)
- set_listen_to_device_ps(capture_device_id, enable, render_device_id=None)  
  Directly writes Listen enable and playback-through target via IPropertyStore::SetValue and Commit using a raw ctypes VTABLE.
- _get_listen_to_device_status_ps(device_id)  
  Reads PKEY_LISTEN_ENABLE via IPropertyStore::GetValue.
- Registry Fallback and Verification  
  Probes HKCU\...\MMDevices\Audio\Capture\{GUID}\(FxProperties|Properties) for the AudioEndpointSettings GUID and verifies expected values using a short polling window.

Thread/apartment model: helpers do not call CoInitialize/CoUninitialize; they rely on initialization done by GUI/CLI entry.

### 5.5. Default Endpoint Management
- _get_policy_config()  
  Attempts to acquire an IPolicyConfig or IPolicyConfigVista COM interface via multiple fallback paths.
- set_default_endpoint(device_id, role)  
  Invokes SetDefaultEndpoint() on the acquired interface.
- Administrator Privileges  
  Often required; the script warns in both GUI and CLI when not elevated.

### 5.6. Volume and Mute Control
- set_endpoint_mute(), get_endpoint_mute(), set_endpoint_volume(), get_endpoint_volume()  
  Interact with IAudioEndpointVolume obtained by IMMDevice.Activate. Includes robust handling for comtypes tuple/out-param variants.

### 5.7. Audio Enhancements (SysFX) — Vendor-First Strategy (New)
Different audio drivers implement the “Audio Enhancements” toggle in varying, vendor-specific ways.
- Read/Write Paths
  - Windows-native readings:
    - PropertyStore read of Disable_SysFx (PKEY_AudioEndpoint_Disable_SysFx).
    - PolicyConfig COM path for reading/writing Disable_SysFx in both FX store and normal store (readers used for diagnostics).
  - Vendor toggles:
    - Preferred way to control Enhancements state for reliability.
    - A code-maintained list of known vendor DWORD toggles (currently includes a Realtek/Waves primary DWORD), plus a user-maintained vendor INI (vendor_toggles.ini).
- vendor_toggles.ini
  - Append-only database the tool can learn and write.
  - Format per entry:
    ```
    [vendor_section_name]
    value_name = {GUID},pid
    dword_enable = 0|1
    dword_disable = 0|1
    hives = HKLM,HKCU
    flows = Render,Capture
    notes = optional
    ```
  - Location: next to the EXE by default (same folder as audioctl.exe); created on first learn.
  - HKLM writes require Administrator privileges.
- Learn Modes
  - Manual (GUI and CLI): You flip Enhancements in Windows settings as prompted; the tool captures “Enabled” and “Disabled” snapshots, diffs MMDevices registry, and writes an INI entry when a reliable REG_DWORD flip under FxProperties is found.
  - Discover-Only (CLI report): Captures A/B snapshots for your manual flipping and writes TXT/JSON reports; can optionally append a suggested INI snippet.
- Operational Policy
  - The runtime Enhancements toggle is vendor-only (no automatic fallback to Windows paths). If no vendor entry applies, the command fails and instructs the user to learn first.
  - This design avoids false positives and inconsistent states caused by diverging Windows-vs-driver controls.

### 5.8. Comprehensive Error Handling and Logging
- Global sys.excepthook and sys.unraisablehook  
  Log unhandled and unraisable exceptions with tracebacks to audioctl_gui.log.
- faulthandler Integration  
  Dumps tracebacks on fatal crashes to the same log.
- atexit Hooks, sys.exit Hook, and Console Control Handlers  
  Provide additional breadcrumbs and shutdown diagnostics.
- Dedicated Log File  
  audioctl_gui.log in the application directory (or temp folder as fallback).
- GUI Error Dialogs  
  report_callback_exception logs and shows user-friendly error popups with the log path.

---

## 6. CLI Command Reference
The script exposes the following commands. All selection options support either --id or --name and share consistent disambiguation rules via --index in GUI order (sorted by name within each flow).

- list  
  Name-sorted within each flow (Render/Capture) to match the GUI. Supports human-readable and JSON output.
  ```bash
  audioctl list [--all] [--json]
  ```

- set-default  
  Sets the default playback/recording device for the specified roles. Requires the target device to be active.
  ```bash
  audioctl set-default
    (--playback-id <ID> | --playback-name <NAME>) [--playback-role <ROLE>]
    (--recording-id <ID> | --recording-name <NAME>) [--recording-role <ROLE>]
    [--index <N>] [--regex]
  ```
  Roles: console, multimedia, communications, all.

- set-volume  
  Sets master volume or mutes/unmutes the target device. Operates only on active endpoints.
  ```bash
  audioctl set-volume
    (--id <ID> | --name <NAME>) [--flow <FLOW>]
    (--level <0-100> | --mute | --unmute)
    [--index <N>] [--regex]
  ```

- listen  
  Enables or disables “Listen to this device” for an active capture device. Writes via IPropertyStore and verifies via COM, falling back to a short registry poll if needed. Use --playback-target-id "" to select the Default Playback Device.
  ```bash
  audioctl listen
    (--id <ID> | --name <NAME>)
    (--enable | --disable)
    [--playback-target-id <ID>] [--index <N>] [--regex]
  ```

- enhancements (New)  
  Enables or disables “Audio Enhancements” (SysFX) using vendor toggles. If no vendor toggle is available for the device, the command fails and suggests using --learn to create one. Exactly one action must be specified: --enable, --disable, or --learn (manual learning mode).
  ```bash
  audioctl enhancements
    (--id <ID> | --name <NAME>) [--flow <FLOW>]
    (--enable | --disable | --learn)
    [--index <N>] [--regex]
    [--prefer-hklm]             # prefer HKLM writes (Admin required)
    [--vendor-ini <PATH>]       # custom path for vendor_toggles.ini; default next to EXE
  ```
  Notes:
  - --learn triggers manual learning (you flip the Windows setting on prompts); on success, an INI entry is appended.

- diag-sysfx (New)  
  Prints live Enhancements state from Windows (COM/PropertyStore) and the first applicable vendor toggle entry (if any).
  ```bash
  audioctl diag-sysfx
    (--id <ID> | --name <NAME>) [--flow <FLOW>] [--index <N>] [--regex]
  ```

- diag-mmdevices (New)  
  Dumps all MMDevices values under HKCU/HKLM for the endpoint’s Render/Capture keys, FxProperties/Properties. Useful for debugging and discovery.
  ```bash
  audioctl diag-mmdevices
    (--id <ID> | --name <NAME>) [--flow <FLOW>] [--index <N>] [--regex]
  ```

- discover-enhancements (New)  
  Interactive A/B snapshot capture for Enhancements “Enabled” and “Disabled”, writing a human-readable TXT report and a JSON bundle, with an optional INI snippet.
  ```bash
  audioctl discover-enhancements
    (--id <ID> | --name <NAME>) [--flow <FLOW>] [--index <N>] [--regex]
    [--output-dir <DIR>]        # where TXT/JSON reports are written
    [--ini-snippet <PATH>]      # append a suggested INI section here
  ```

- wait  
  Waits for a device to become active (appear) within the given timeout and outputs its details if found.
  ```bash
  audioctl wait
    (--id <ID> | --name <NAME>)
    [--flow <FLOW>] [--timeout <SECONDS>] [--index <N>] [--regex]
  ```

General notes:
- Mutating commands operate only on active endpoints (DEVICE_STATE_ACTIVE).
- The list command’s --all option affects visibility, not mutating behavior.

---

## 7. GUI Application Structure
The AudioGUI class handles:

- Initialization (__init__)  
  Sets up the main window, Treeview, scrollbar, status bar, and context menu. Applies custom styling. Window title reflects v1.3.0.17 12-22-2025.

- Device Management (refresh_devices)  
  Populates the Treeview with devices, categorizing them into “Playback (Render)” and “Recording (Capture)” groups. The device list is name-sorted within each flow; CLI mirrors this order and per‑flow indices.

- User Interaction
  - Right-click (on_right_click)  
    Displays a context menu for the selected device, dynamically adjusting menu item labels and states.
  - Selection guards (on_left_click, on_select_change)  
    Ensure only device rows (not group headers) are selectable.
  - Double-click (on_double_click)  
    Opens the context menu for the item under the cursor.

- Context Menu Actions
  - Set as Default (on_set_default)  
    Sets the selected device as default for all roles; warns if not elevated.
  - Set Volume (on_set_volume)  
    Opens a small dialog for adjusting volume precisely (0–100).
  - Mute/Unmute (on_toggle_mute)  
    Toggles mute status; safe default is to show “Unmute” if unreadable.
  - Listen (on_toggle_listen)  
    Toggles “Listen to this device” for capture devices; uses property store and verification logic.
  - Enhancements (on_toggle_enhancements) — New  
    - If a vendor toggle applies, shows “Enable Enhancements” or “Disable Enhancements” based on current state (vendor read).
    - Executes a vendor-only toggle and displays status.
    - If no vendor toggle applies, the item is disabled and “Learn Enhancements” is available.
  - Learn Enhancements (on_learn_enhancements) — New  
    - Guided, manual process to capture snapshots A/B (Enabled/Disabled) and write vendor_toggles.ini with a new entry if a reliable DWORD flip is detected under FxProperties.
    - Warns that learned entries persist; HKLM writes require Admin.

- open_volume_dialog()  
  Creates a tk.Toplevel window for precise volume adjustment with a slider and entry box.

- resource_path() and iconbitmap()  
  Ensures audio.ico is set for the root and any dialogs in both source runs and PyInstaller builds.

---

## 8. Deployment Considerations (PyInstaller)
The script is deployed as a standalone Windows executable using PyInstaller. The build process integrates version metadata and the application icon.

### 8.1. PyInstaller Command
The enhancements subsystem does not require additional third-party packages; it uses only standard library modules (e.g., configparser) and COM/ctypes functionality. The same build command applies:

```bash
pyinstaller --clean --onefile --console --name audioctl \
  --collect-all pycaw \
  --collect-submodules comtypes \
  --hidden-import comtypes \
  --hidden-import comtypes.gen \
  --hidden-import comtypes.automation \
  --icon audio.ico \
  --add-data "audio.ico;." \
  --version-file version.txt \
  audioctl.py
```

Notes:
- vendor_toggles.ini is optional and created/updated at runtime next to the EXE (no need to include it at build time). If you ship a pre-populated vendor_toggles.ini, add it as data with --add-data accordingly.

### 8.2. Version File (version.txt)
Update to 1.3.0.17:
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

Key fields and their significance:
- filevers=(1, 3, 0, 17) / FileVersion  
  Specifies the binary file version.
- prodvers=(1, 3, 0, 17) / ProductVersion  
  Specifies the product version this file belongs to.

### 8.3. Output
Upon successful execution of the PyInstaller command, the audioctl.exe executable will be generated in the dist/ directory, ready for distribution. A build/ directory will also be created containing intermediate build files.

---

## 9. Conclusion
`audioctl.py` 1.3.0.17 expands beyond device enumeration, default routing, volume/mute, and “Listen to this device” by adding a robust, vendor-first system for managing “Audio Enhancements (SysFX)”. This design combines safe, low-level Windows readings with a learnable, append-only INI database to drive reliable toggles across driver ecosystems. The new CLI commands and GUI actions help diagnose, discover, and operate Enhancements state in a controlled way. The continued emphasis on explicit COM initialization, strategic use of raw ctypes VTABLE calls, and a comtypes compatibility shim ensures dependable behavior in both source and packaged deployments. Comprehensive logging and error handling remain in place to aid diagnostics in production environments.
```
