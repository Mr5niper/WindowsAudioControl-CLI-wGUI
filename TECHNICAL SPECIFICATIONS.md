# Technical Specifications: audioctl.py — Windows Audio Control Utility

## 1. Introduction

`audioctl.py` is a versatile Python script designed to provide comprehensive control over Windows audio devices. It offers both a robust command-line interface (CLI) for scripting and automation, and a user-friendly graphical user interface (GUI) built with Tkinter for interactive management. The script leverages low-level Windows COM APIs through `ctypes` and `comtypes` to achieve reliable and precise audio manipulation, addressing common challenges and stability issues encountered with higher-level abstractions, especially when packaged as a standalone executable.

This document details the script's core functionalities, underlying technologies, architectural decisions, and specific deployment considerations using PyInstaller.

---

## 2. Core Technologies & Dependencies

- Python 3.x  
  The primary programming language.

- `comtypes`  
  A Python package providing a lightweight Python wrapper around Microsoft COM interfaces. It simplifies interaction with many Windows APIs.

- `ctypes`  
  Python's foreign function interface library, used here for direct low-level interaction with Windows DLLs and COM interfaces, particularly where `comtypes` abstractions proved insufficient or problematic or for improved robustness.

- `pycaw`  
  A Python library building on `comtypes` to provide a more convenient interface for Windows Core Audio APIs (MMDevice API, EndpointVolume API). Note that this script implements several functionalities around or instead of `pycaw`'s direct methods due to stability and robustness concerns with `comtypes` in certain deployment scenarios (e.g., PyInstaller).

- `tkinter`  
  Python's standard GUI library, used for the optional graphical interface.

- `argparse`  
  Python's standard library for parsing command-line arguments.

- `json`  
  For structured output in CLI and data exchange.

- `re`  
  Regular expressions for device name matching.

- `os`, `sys`, `time`, `warnings`, `io`, `contextlib`, `tempfile`, `datetime`, `traceback`, `atexit`, `faulthandler`, `winreg`  
  Standard Python libraries and Windows-specific modules for system interaction, logging, and error handling.

- Windows COM (Component Object Model) APIs  
  The underlying technology used, including:
  - `IMMDeviceEnumerator`: For enumerating audio endpoint devices.
  - `IMMDevice`: For representing individual audio endpoint devices.
  - `IAudioEndpointVolume`: For controlling volume and mute state.
  - `IPropertyStore`: For reading/writing device properties, notably "Listen to this device."
  - `IPolicyConfigVista` (or similar `IPolicyConfig`): For setting default audio endpoints.

- PyInstaller  
  The tool used to package the script into a standalone Windows executable. Its specific configuration is detailed in Section 8.

- External Assets
  - `audio.ico`: The icon file used for the generated executable and Tkinter windows.
  - `version.txt`: A text file containing version information for the executable's metadata.

---

## 3. Dependency Management (Virtual Environment)

The development and deployment of this script utilize a Python virtual environment (`venv`). This practice ensures that project dependencies are isolated from the system-wide Python installation, preventing conflicts and ensuring consistent environments across development, testing, and deployment.

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

To get a precise list of all installed packages and their versions within your active virtual environment, use the following command to generate a `requirements.txt` file:

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

Note: The exact list of packages and their versions was generated from a specific `venv` using `pip freeze`.

---

## 4. Architectural Overview

The script is structured into two main parts:

a) Core Audio Logic  
A set of functions that interacts directly with Windows COM APIs to query and control audio devices. These functions are designed to be robust and handle potential `comtypes`/`ctypes` interactions gracefully.

b) User Interfaces
- Command-Line Interface (CLI)  
  Implemented using `argparse`, providing distinct commands (`list`, `set-default`, `set-volume`, `listen`, `wait`).
- Graphical User Interface (GUI)  
  A Tkinter application that provides an interactive way to view devices and perform common actions, invoking the same core audio logic.

The script defaults to launching the GUI if no command-line arguments are provided; otherwise, it executes the specified CLI command.

---

## 5. Key Design Patterns & Rationale

### 5.1. COM Interoperability and Robustness

Interfacing with Windows COM APIs from Python, especially when bundled with PyInstaller, can be challenging due to `comtypes` sometimes exhibiting unexpected behavior or missing definitions. The script employs a layered approach:

- `comtypes` as Primary  
  For most standard operations (device enumeration, `IAudioEndpointVolume`), `comtypes` is used for its convenience and Pythonic wrappers.

- `ctypes` for Critical Paths  
  For highly sensitive or problematic areas, particularly the "Listen to this device" functionality and robust property store access (`IPropertyStore`), the script reverts to direct `ctypes` VTABLE calls.

- Rationale  
  This bypasses potential `comtypes.gen` issues (especially in bundled executables) and ensures maximum control and stability by manually defining COM interface structures (`IPropertyStoreRaw`, `IPropertyStoreVTBL`) and method prototypes (`CALL`). This is a critical design choice to enhance reliability.

- `CoInitialize()` / `CoUninitialize()`  
  Called explicitly in both the GUI and the CLI entry points. Low-level helpers (e.g., Listen via IPropertyStore) assume COM is already initialized by their caller. This avoids per-call COM teardown and prevents benign comtypes finalizer warnings.

### 5.2. `comtypes` Compatibility Shim

A minimal shim is placed at the very top of the script.

- Purpose  
  To address known PyInstaller bundling issues where `comtypes.automation` might not correctly expose `PROPVARIANT` or standard `VT_` constants (like `VT_LPWSTR`, `VT_BOOL`, `VARIANT_TRUE`/`FALSE`).

- Mechanism  
  It checks for the existence of these attributes and defines them with their known integer values if missing. It also ensures `_automation.PROPVARIANT` points to `_automation.tagPROPVARIANT` if the former is not directly present.

- Rationale  
  Ensures foundational types and constants required for manual `PROPVARIANT` construction and parsing (used in the raw `ctypes` sections) are always available, even if `comtypes`'s dynamic generation fails or is incomplete in a bundled environment.

### 5.3. Robust Device Enumeration and Naming

- `list_devices()`  
  Avoids `AudioUtilities.GetAllDevices()` (a `pycaw` utility) due to reported `comtypes` `Release()` crashes in certain scenarios. Instead, it directly uses `IMMDeviceEnumerator.EnumAudioEndpoints()` for enumeration.

- `_safe_friendly_name_from_device(dev)`  
  Retrieves the device's friendly name (`PKEY_Device_FriendlyName`) or device description (`PKEY_Device_DeviceDesc`) via direct `IMMDevice.OpenPropertyStore()` and `IPropertyStore::GetValue()` calls using `ctypes`.

- Rationale  
  Provides robust retrieval of human-readable device names, bypassing potentially fragile `comtypes` property accessors or Python `UserWarning`s, and handles multiple `PROPVARIANT` memory layouts.

### 5.4. "Listen to this Device" Implementation (Advanced `ctypes`)

This is one of the most technically involved sections.

- `set_listen_to_device_ps(capture_device_id, enable, render_device_id=None)`
  - Direct `IPropertyStore` Access  
    Opens the device's `IPropertyStore` (with `STGM_WRITE` permission) using `IMMDevice.OpenPropertyStore()`.
  - Raw `ctypes` VTABLE Calls  
    Casts the underlying COM pointer to a `POINTER(IPropertyStoreRaw)` and manually calls methods like `SetValue()` and `Commit()` through the VTABLE.
  - `PROPERTYKEY`  
    Defines:
    - `PKEY_LISTEN_ENABLE` — GUID `{24dbb0fc-9311-4b3d-9cf0-18ff155639d4}`, PID `1`
    - `PKEY_LISTEN_PLAYBACKTHROUGH` — same GUID, PID `2`
  - `PROPVARIANT` Management  
    Utilizes Windows `propsys.dll` functions (`InitPropVariantFromBoolean`, `InitPropVariantFromString`) if available, or falls back to manual `PROPVARIANT` struct population. Uses `PropVariantClear` from `ole32.dll` for cleanup.

- `_get_listen_to_device_status_ps(device_id)`  
  Reads the `PKEY_LISTEN_ENABLE` property using raw `ctypes` `IPropertyStore::GetValue()`.

- Registry Fallback (`_read_listen_enable_from_registry`, `_verify_listen_via_registry`)
  - Purpose  
    Secondary verification mechanism when COM results are delayed or ambiguous.
  - Mechanism  
    Probes:
    ```
    HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\Capture\{GUID}\(FxProperties|Properties)
    ```
    for the AudioEndpointSettings GUID (`{24dbb0fc-9311-4b3d-9cf0-18ff155639d4}`), handling `REG_DWORD`, `REG_BINARY` (PROPVARIANT), and `REG_SZ`.
  - Rationale  
    Registry provides a persistent view to confirm effective state changes.

_Thread/apartment model: helpers do not call `CoInitialize`/`CoUninitialize`; they rely on initialization done by GUI/CLI entry._

### 5.5. Default Endpoint Management

- `_get_policy_config()`  
  Attempts to acquire an `IPolicyConfig` or `IPolicyConfigVista` COM interface.
  - Multi-tier Fallback  
    Tries `AudioUtilities.GetPolicyConfig`, then `pycaw.policyconfig`, and finally a locally defined minimal `IPolicyConfigVista` interface via `ctypes`.
  - Rationale  
    Ensures compatibility across environments and `pycaw` versions.
- `set_default_endpoint(device_id, role)`  
  Invokes `SetDefaultEndpoint()` via `IPolicyConfigVista`.
- Administrator Privileges  
  Often required; the script warns in both GUI and CLI when not elevated.

### 5.6. Volume and Mute Control

- `set_endpoint_mute()`, `get_endpoint_mute()`, `set_endpoint_volume()`, `get_endpoint_volume()`  
  Interact with `IAudioEndpointVolume` obtained by `IMMDevice.Activate`.

- Robust Return Value Handling  
  Handles both direct return values and out-params exposed as tuples by `comtypes`, with `ctypes.byref` fallbacks.

### 5.7. Comprehensive Error Handling and Logging

- Global `sys.excepthook`  
  Logs unhandled exceptions with tracebacks to `audioctl_gui.log`.

- `sys.unraisablehook` (Python 3.8+)  
  Logs unraisable exceptions.

- `faulthandler` Integration  
  Dumps tracebacks on fatal crashes to the same log.

- `atexit` Hooks  
  Logs normal exits and flushes/cleans up `faulthandler` file handle.

- `sys.exit` Hook  
  Logs `sys.exit(code)` calls.

- Windows Console Control Handler  
  Logs control events (Ctrl+C/Break, logoff, shutdown).

- Dedicated Log File  
  `audioctl_gui.log` in the application directory (or temp folder as fallback).

- GUI Error Dialogs  
  `report_callback_exception` logs and shows user-friendly error popups with the log path.

---

## 6. CLI Command Reference

The script exposes the following commands:

- list  
  The list output is name-sorted within each flow (Render/Capture) to match the GUI and uses the same per‑flow indices. Supports human-readable and JSON output.
  ```bash
  audioctl list [--all] [--json]
  ```

- set-default  
  Sets the default playback/recording device for the specified roles. Requires the target device to be active. When selecting by name, `--index` refers to the same per‑flow (GUI) order; providing `--flow` is recommended when disambiguating.
  ```bash
  audioctl set-default
    (--playback-id <ID> | --playback-name <NAME>) [--playback-role <ROLE>]
    (--recording-id <ID> | --recording-name <NAME>) [--recording-role <ROLE>]
    [--index <N>] [--regex]
  ```

- set-volume  
  Sets master volume or mutes/unmutes the target device. Operates only on active endpoints. When multiple matches exist, `--index` is interpreted in the GUI’s per‑flow order; providing `--flow` is recommended with `--name + --index`.
  ```bash
  audioctl set-volume
    (--id <ID> | --name <NAME>) [--flow <FLOW>]
    (--level <LEVEL> | --mute | --unmute)
    [--index <N>] [--regex]
  ```

- listen  
  Enables or disables “Listen to this device” for an active capture device. Writes via `IPropertyStore` and verifies via COM, falling back to a short registry poll if needed. Use `--playback-target-id ""` to select the Default Playback Device.
  ```bash
  audioctl listen
    (--id <ID> | --name <NAME>)
    (--enable | --disable)
    [--playback-target-id <ID>] [--index <N>] [--regex]
  ```

- wait  
  Waits for a device to become active (appear) within the given timeout and outputs its details if found.
  ```bash
  audioctl wait
    (--id <ID> | --name <NAME>)
    [--flow <FLOW>] [--timeout <SECONDS>] [--index <N>] [--regex]
  ```

_Note: Mutating commands operate only on active endpoints (`DEVICE_STATE_ACTIVE`). The list command’s `--all` option affects visibility, not mutating behavior._

---

## 7. GUI Application Structure

The `AudioGUI` class handles:

- Initialization (`__init__`)  
  Sets up the main window, Treeview, scrollbar, status bar, and context menu. Applies custom styling.

- Device Management (`refresh_devices`)  
  Populates the Treeview with devices, categorizing them into "Playback (Render)" and "Recording (Capture)" groups. The device list is name-sorted within each flow (Render/Capture). The CLI mirrors this order and the per‑flow indices, so GUI and CLI indexing are consistent.

- User Interaction
  - `on_right_click()`  
    Displays a context menu for the selected device, dynamically adjusting menu item labels and states.
  - `on_left_click()`, `on_select_change()`  
    Custom handlers for Treeview selection to ensure only device rows (not group headers) are selectable.
  - `on_double_click()`  
    Displays a context menu for the selected device, dynamically adjusting menu item labels and states.

- Action Handlers (`on_set_default`, `on_set_volume`, `on_toggle_mute`, `on_toggle_listen`)  
  - Retrieve the selected device.  
  - Construct and optionally print the equivalent CLI command (`maybe_print_cli`).  
  - Call the appropriate core audio logic functions (`set_default_endpoint`, `set_endpoint_volume`, `set_endpoint_mute`, `set_listen_to_device_ps`).  
  - Update the status bar and refresh the device list as needed.  
  - Wrap calls in `try-except` blocks to display user-friendly error messages via `messagebox`.

- `open_volume_dialog()`  
  Creates a `tk.Toplevel` window for precise volume adjustment with a slider and entry box.

- `resource_path()` and `iconbitmap()`  
  The `resource_path` helper function intelligently locates bundled assets, ensuring that `audio.ico` is correctly set as the icon for both the main `Tk` root window and any `tk.Toplevel` dialogs, whether the script runs from source or as a PyInstaller executable.

---

## 8. Deployment Considerations (PyInstaller)

The script is deployed as a standalone Windows executable using PyInstaller. The build process integrates version metadata and the application icon.

### 8.1. PyInstaller Command

The following command is used for building the executable:

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

Breakdown of arguments:

- `--clean`  
  Cleans PyInstaller's cache and temporary files before building.

- `--onefile`  
  Packages the application into a single executable file.

- `--console`  
  Creates a console window for the application. Important for viewing `stdout`/`stderr` output (including CLI commands and logged warnings/errors) even when the GUI is launched.

- `--name audioctl`  
  Sets the name of the executable to `audioctl.exe`.

- `--collect-all pycaw`  
  Automatically collects all modules related to the `pycaw` package.

- `--collect-submodules comtypes`  
  Ensures all submodules of `comtypes` are collected.

- `--hidden-import comtypes`, `--hidden-import comtypes.gen`, `--hidden-import comtypes.automation`  
  Explicitly include these `comtypes` modules to accommodate dynamic code generation that static analysis can miss, directly supporting the `comtypes` compatibility shim.

- `--icon audio.ico`  
  Specifies `audio.ico` as the icon for the generated `audioctl.exe`.

- `--add-data "audio.ico;."`  
  Includes `audio.ico` as a data file accessible at runtime. The `;.` places the icon at the root of the temporary execution folder, making it accessible via `resource_path("audio.ico")` in the GUI.

- `--version-file version.txt`  
  Integrates version information from `version.txt` into the executable's metadata.

- `audioctl.py`  
  The main script file to be bundled.

### 8.2. Version File (`version.txt`)

The `version.txt` file contains the `VSVersionInfo` structure, which populates the "Details" tab in the `audioctl.exe` file properties in Windows Explorer.

```text
VSVersionInfo(
  ffi=FixedFileInfo(
    filevers=(1, 2, 1, 0),
    prodvers=(1, 2, 1, 0),
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
          StringStruct('FileVersion', '1.2.1.0'),
          StringStruct('InternalName', 'audioctl'),
          StringStruct('LegalCopyright', ''),
          StringStruct('OriginalFilename', 'audioctl.exe'),
          StringStruct('ProductName', 'Windows Audio Control CLI'),
          StringStruct('ProductVersion', '1.2.1.0'),
        ]
      )
    ]),
    VarFileInfo([VarStruct('Translation', [1033][1200])])
  ]
)
```

Key fields and their significance:

- `filevers=(1, 2, 1, 0)` / `FileVersion`  
  Specifies the binary file version.

- `prodvers=(1, 2, 1, 0)` / `ProductVersion`  
  Specifies the product version this file belongs to.

- `CompanyName`  
  Identifies the organization that produced the application.

- `FileDescription`  
  A brief description of the file's purpose.

- `InternalName`  
  The internal name of the file (usually without extension).

- `LegalCopyright`  
  Copyright notice for the product.

- `OriginalFilename`  
  The original name of the file, typically used for comparison.

- `ProductName`  
  The name of the product this file is part of.

- `StringTable('040904B0', ...)`  
  Specifies language and character set (0409 is US English, 04B0 is Unicode).

- `VarStruct('Translation', [1033, 1200])`  
  Defines the translation information (1033 = US English, 1200 = Unicode codepage).

### 8.3. Output

Upon successful execution of the PyInstaller command, the `audioctl.exe` executable will be generated in the `dist/` directory, ready for distribution. A `build/` directory will also be created containing intermediate build files.

---

## 9. Conclusion

`audioctl.py` is a robust and flexible solution for Windows audio control. Its careful implementation of COM interoperability, including strategic use of raw `ctypes` VTABLE calls and a `comtypes` compatibility shim, addresses many common stability issues encountered with Python-COM interactions, particularly in packaged applications. The dual CLI/GUI interface provides broad utility for both automation and interactive use. Coupled with a disciplined virtual environment for dependency management and a thoroughly configured PyInstaller build process, this utility is designed for reliable deployment and operation in Windows environments. The comprehensive error handling and logging further enhance its reliability and debuggability in production.
```
