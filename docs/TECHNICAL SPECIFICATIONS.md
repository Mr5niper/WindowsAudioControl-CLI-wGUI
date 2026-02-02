# Technical Specifications & Engineering Guide

**Audioctl (Windows Audio Control CLI + GUI) v1.4.7.2**
<BR>
<BR>
**Last updated:** 2026-02-02  
**Audience:** Engineers maintaining or extending audioctl  
**Platforms:** Windows 10/11 (x64 assumed; x86 supported if dependencies permit)  
**Core dependencies:** Python 3.13.5, pycaw, comtypes, ctypes, winreg, tkinter (GUI), PyInstaller (optional)

This document describes the **entire program**, including architecture, module contracts, data formats, COM/GC stability constraints, registry touchpoints, INI schema, CLI/GUI behaviors, and engineering extension guidance.

---

## 1) System goals and non-goals

### 1.1 Goals
- Provide a **scriptable CLI** for repeatable Windows audio endpoint operations.
- Provide an optional **GUI** that uses the CLI (not direct COM calls) for consistent behavior.
- Support vendor-specific enhancements/effects toggles via a learned, persistent configuration database (`vendor_toggles.ini`).
- Offer read-only “query helpers” designed for fast, side-effect-free state retrieval.
- Prioritize **stability** in the face of COM lifetime hazards and driver variability.

### 1.2 Non-goals
- Cross-platform support (Windows-only by design).
- Per-application audio session control (this tool controls endpoints, not per-app mixers).
- Perfect vendor coverage without learning: vendor toggles are inherently driver-specific.

---

## 2) Architectural overview

### 2.1 Layer model
**Layer 1 - Entry / orchestration**
- `audioctl.py` (PyInstaller entrypoint)
- `audioctl/__main__.py` (`python -m audioctl`)

**Layer 2 - Public surface (contracts)**
- `audioctl/cli.py` (argparse + command dispatch + JSON contracts)

**Layer 3 - UI**
- `audioctl/gui.py` (Tkinter; subprocess calls to CLI only)

**Layer 4 - Low-level Windows integration**
- `audioctl/devices.py` (COM, PropertyStore vtables, registry snapshots, endpoint control)
- `audioctl/vendor_db.py` (INI schema, registry writes for vendor toggles, learning/merge logic)

**Layer 5 - Cross-cutting**
- `audioctl/compat.py` (comtypes shims, forced imports for PyInstaller bundling)
- `audioctl/logging_setup.py` (lazy logging, exception hooks, fault handler)

### 2.2 Primary data flows
- CLI command → device selection → low-level action (COM/registry) → JSON result.
- GUI event → CLI subprocess call → parse JSON → update UI/state cache.

### 2.3 Contracts (hard requirements)
1. JSON output keys are API.
2. Prompt strings used for GUI interactive learn are API (pattern-based).
3. `--index` ordering must match GUI order.
4. Mutating commands must remain active-only unless intentionally changed.
5. COM initialization must be correct and isolated; avoid global COM init patterns.

---

## 3) Repository layout & module specifications

### 3.1 `audioctl.py`
**Purpose:** Frozen EXE entrypoint and console attachment helper.  
**Responsibilities:**
- Attach to parent console on Windows so shells (notably PS 5.1) don’t spawn a new window.
- Call `audioctl.cli.main()` and propagate exit code.

**Technical constraints:**
- Must remain minimal.
- Must not import heavy modules before console attach attempt (best-effort).

### 3.2 `audioctl/cli.py`
**Purpose:** CLI command router and canonical JSON contract provider.

**Responsibilities:**
- argparse subcommands definition.
- Device selection logic and disambiguation behavior.
- Calls to `devices.py` and `vendor_db.py`.
- GUI auto-launch when invoked with no arguments.

**Required behaviors:**
- For mutating operations, select only active endpoints (`include_all=False`).
- On ambiguity, return exit code `4` and require `--index`.

**CLI exit codes:**
- `0` success
- `1` invalid args/runtime failure
- `3` not found / wait timeout
- `4` multiple matches / index required / index out of range
- `130` Ctrl-C

**Command surface (spec):**
- `list [--all] [--json]`
  - JSON: `{ "devices": [ {id,name,flow,state,isDefault}, ... ] }`

- `set-default`
  - Args: playback-id|playback-name, recording-id|recording-name, roles, regex, index
  - JSON: `{ "set": [ {flow, role, id, name}, ... ] }`

- `set-volume`
  - Args: --id|--name, optional --flow, exactly one of --level|--mute|--unmute, optional --index, --regex
  - JSON success:
    - `{ "volumeSet": {id,name,level} }` OR `{ "muteSet": {id,name,muted} }`

- `get-volume`
  - JSON: `{ id, name, flow, volume, muted }`

- `listen` (Capture only)
  - Args: capture selector + enable/disable + optional playback target selector
  - JSON: `{ "listenSet": { id, name, enabled, [verifiedBy] } }`

- `get-listen`
  - JSON: `{ id, name, flow, listenEnabled }`

- `enhancements`
  - Main: `--enable|--disable|--learn`
  - FX ops: `--list-fx|--learn-fx|--enable-fx|--disable-fx|--delete-fx`
  - JSON varies by operation:
    - `{ "enhancementsSet": { id, name, enabled, verifiedBy } }`
    - `{ "vendorLearned": {...} }` / `{ "vendorAvailable": {...} }`
    - `{ "fxLearned": {...} }`
    - `{ "fxSet": {...} }`
    - `{ "fxDeleted": {...} }`

- `get-enhancements`
  - JSON: `{ id, name, flow, enhancementsEnabled }` (vendor-only)

- `get-device-state`
  - JSON: `{ id, name, flow, volume, muted, listenEnabled, enhancementsEnabled, availableFX:[...] }`

- `diag-sysfx`, `diag-mmdevices`, `discover-enhancements`, `wait`
  - Diagnostic/discovery commands with JSON or mixed output.

- Hidden: `vendor-ini-append`
  - Internal elevated write helper for GUI deployments in protected directories.

### 3.3 `audioctl/gui.py`
**Purpose:** Tkinter GUI frontend that shells out to CLI as subprocess.

**Responsibilities:**
- Display device list grouped by flow (Render/Capture).
- Context menu actions:
  - set default
  - set volume
  - mute/unmute
  - listen toggle (capture)
  - enhancements toggle (vendor-only)
  - FX toggles (learned)
  - learn workflows (main + FX)
- Background device state caching using `get-device-state`.

**Technical constraints:**
- Must not import low-level COM device helpers directly (avoid long-lived COM state in GUI).
- Must tolerate CLI failure and display actionable errors.
- Must keep UI responsive:
  - incremental state population
  - non-blocking learn runner for long operations

**State cache specification:**
- `device_state_cache: Dict[device_id -> state]`
- State is expected to match `get-device-state` JSON shape.
- Cache is refreshed incrementally after list refresh.

### 3.4 `audioctl/devices.py`
**Purpose:** Low-level Windows integration layer.

#### 3.4.1 COM lifecycle subsystem
- `_com_context()` provides thread-local COM initialization with refcounting.

**Specification:**
- On first entry for a thread: call `comtypes.CoInitialize()`.
- On final exit for a thread: call `comtypes.CoUninitialize()`.

**Invariants:**
- Nested calls must not prematurely uninitialize COM.
- COM init/exit errors are best-effort and must not crash the process directly.

#### 3.4.2 Endpoint enumeration
**API:**
- `list_devices(include_all: bool) -> List[device]`
- `find_devices_by_selector(devices, dev_id, name_substr, flow, regex) -> List[device]`
- `_sort_and_tag_gui_indices(devices) -> {Render: [...], Capture: [...]} (mutates devices to add guiIndex)`

**Device object shape:**
```json
{
  "id": "<IMMDevice ID>",
  "name": "<friendly name>",
  "flow": "Render|Capture",
  "state": "active|disabled|notpresent|unplugged|...",
  "isDefault": {
    "console": true|false,
    "multimedia": true|false,
    "communications": true|false
  }
}
```

#### 3.4.3 Default endpoint management
**API:**
- `set_default_endpoint(device_id: str, role: str) -> None`
- Roles:
  - `console`, `multimedia`, `communications`, `all`

**Constraints:**
- Must refuse inactive endpoints.

#### 3.4.4 Volume/mute
**API:**
- `get_endpoint_volume(device_id) -> int|None` (0..100)
- `set_endpoint_volume(device_id, level_percent) -> bool`
- `get_endpoint_mute(device_id) -> bool|None`
- `set_endpoint_mute(device_id, mute_state: bool) -> bool`

**Normalization rules:**
- Volume is scalar float 0.0..1.0 → percent 0..100.
- Mute returns bool when readable; None only on unrecoverable variant mismatch.

#### 3.4.5 Listen feature
**Enable flag (PropertyStore):**
- PROPERTYKEY fmtid `{24dbb0fc-9311-4b3d-9cf0-18ff155639d4}`, pid `1`
- Written via IPropertyStore::SetValue + Commit.

**Routing target (registry, HKLM):**
- HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\Capture\{endpoint-guid}\Properties
- Value `{24dbb0fc-9311-4b3d-9cf0-18ff155639d4},0` (REG_SZ)
- Empty string means “Default Playback Device”.

**Verification:**
- COM readback preferred; registry fallback and polling permitted.

#### 3.4.6 SysFX (Disable_SysFx) helpers
**Disable_SysFx PROPERTYKEY:**
- fmtid `{E4870E26-3CC5-4CD2-BA46-CA0A9A70ED04}`, pid `2`
- Semantics:
  - 0 → Enhancements enabled
  - 1 → Enhancements disabled

**Read paths:**
- PolicyConfigFx (`GetPropertyValue` with bFxStore True/False)
- Endpoint PropertyStore (raw vtable)
- Registry scanning for discovery/diagnostics

**Write paths:**
- PolicyConfigFx (`SetPropertyValue`)
- Endpoint PropertyStore (`SetValue`+Commit)
- Registry writes (used for learning/testing; not always runtime)

#### 3.4.7 Registry snapshot & diff subsystem
**API:**
- `_dump_mmdevices_all_values(device_id) -> List[record]`
- `_diff_mmdevices_lists(before, after) -> diff_summary`
- `_collect_sysfx_snapshot(device_id) -> {time, com, propStore, registry}`
- `_generate_enh_discovery_report(target, snapA, snapB, diffs) -> str`

**Registry record shape:**
```json
{
  "hive": "HKCU|HKLM",
  "flow": "Render|Capture",
  "subkey": "FxProperties|Properties|FxProperties\\...\\...",
  "name": "<value name>",
  "type": <REG_* int>,
  "dataPreview": "<human preview>",
  "dataRaw": "<exact payload (int/str/hex)>"
}
```

**Diff summary shape:**
- `added`, `removed`, `changed`
- `dword_flips`: (0↔1 flips)
- `disable_sysfx_hits`: entries involving Disable_SysFx fmtid

#### 3.4.8 PropertyStore raw vtable subsystem
`_get_property_store_interfaces()` builds a cached ctypes vtable definition.

**Stability requirements:**
- Must remain cached at module scope.
- Calls using raw pointers should remain GC-guarded.
- PropVariantClear must be called on allocated PROPVARIANTs.

### 3.5 `audioctl/vendor_db.py`
**Purpose:** Vendor toggles persistence + application + learning.

#### 3.5.1 INI file location
`vendor_toggles.ini` default path:
- EXE directory if writable
- else %LOCALAPPDATA%\audioctl\vendor_toggles.ini

#### 3.5.2 INI cache
Cache key:
- absolute path
- file mtime

**Behavior:**
- If file missing: cache “missing” and return empty DB.
- If mtime unchanged: reuse parsed DB.

#### 3.5.3 Split DB model
The loader returns:
```python
{ "main": [ ... ], "fx": [ ... ] }
```

#### 3.5.4 MAIN entry schema (Enhancements)
Required fields:
- `value_name` (normalized lower)
- `dword_enable`, `dword_disable` (int; typically 0/1)
- `hives` list (HKCU/HKLM)
- `flows` list (Render/Capture)
- `devices` list (endpoint GUIDs)
- `subkey` (FxProperties or Properties)

**Applicability rule:**
- Must match endpoint GUID membership.
- Must match flow (if specified).
- Prefer entries that actually exist under HKCU for the endpoint (reduces false positives).

#### 3.5.5 FX schema (legacy single DWORD)
Fields:
- `type=fx`
- `fx_name`
- `value_name`
- `dword_enable`, `dword_disable`
- `devices` list required

#### 3.5.6 FX schema (multi_write)
Fields:
- `type=fx`
- `fx_name`
- `multi_write=1`
- `write_count=N`
- `decider_index` (1-based)
- `quorum_threshold` (0.50..0.95; default 0.60)
- Per write block i:
  - `write{i}_hive = HKCU|HKLM`
  - `write{i}_subkey = FxProperties|Properties`
  - `write{i}_name = {fmtid},pid`
  - `write{i}_type_enable/disable = REG_DWORD|REG_SZ|REG_BINARY`
  - `write{i}_enable/disable = payload`
  - optional: `write{i}_devices`

**write{i}_devices semantics:**
- missing => universal (applies to all devices in bucket)
- empty => applies to nobody
- list => applies only to listed GUIDs

#### 3.5.7 Vendor-only runtime policy
- `_apply_enhancements` is vendor-only.
- If no vendor entry applies: return failure (`no-vendor-method`).

#### 3.5.8 Fast read semantics (GUI support)
Fast read functions do:
- single registry probe
- optional alternate hive probe
- last-write-time tie-break if HKCU/HKLM disagree

For multi_write FX:
- may select a “best” write based on strength (FxProperties + DWORD favored)
- uses decider/quorum logic where appropriate

#### 3.5.9 Learning algorithms (spec)
**Enhancements learn:**
- A/B snapshot capture while user toggles UI.
- Diff to find stable DWORD flips.
- Write INI section and append device GUID.

**FX learn (multi-write):**
- Uses stability filtering:
  - multiple samples to eliminate registry noise
- Derives a write set from stable differences:
  - supports DWORD, SZ, and binary values
- Merge behavior:
  - match identity+payload: append GUID to write{i}_devices
  - new payload: append a new write block
  - conflict cleanup: ensure one GUID is not attached to two competing writes with same identity

### 3.6 `audioctl/compat.py`
**Purpose:** Load-order critical compatibility shim.

**Requirements:**
- Must be imported before any comtypes usage.
- Ensures:
  - PROPVARIANT alias exists
  - VT constants exist (VT_BOOL, VT_LPWSTR, etc.)
  - VARIANT_TRUE/FALSE exists
- Forces import of `comtypes._post_coinit` modules so PyInstaller bundles cleanup code.

### 3.7 `audioctl/logging_setup.py`
**Purpose:** Lazy log file initialization + crash breadcrumbs.

**Specification:**
- No filesystem writes at import time.
- On first log call:
  - resolve log path (prefer exe dir; else temp/appdata)
  - install:
    - sys.excepthook
    - sys.unraisablehook
    - Windows console control handler (best-effort)
  - enable faulthandler if available
  - wrap sys.exit to log breadcrumb
  - register atexit breadcrumbs and handle cleanup

---

## 4) CLI technical specification (device selection & ordering)

### 4.1 Selection algorithm (standard pattern)
Given a command with selectors:
1. Enumerate devices:
   - mutating commands: `list_devices(include_all=False)`
   - list command: `include_all` optional
2. Filter matches:
   - by ID exact match, OR
   - by name substring/regex match, optional flow filter
3. If multiple matches:
   - sort in GUI-order (per-flow name sort)
   - require --index unless command uses a special resolution helper
4. Choose target based on index or first match.
5. Execute operation and return JSON.

### 4.2 GUI-order sorting definition
Within each flow bucket:
- Sort by `name.lower()`
- assign `guiIndex` sequentially starting at 0

---

## 5) Registry touchpoints (technical map)

> Registry paths shown here are conceptual; do not copy/paste these as Python strings unless you escape properly.

### 5.1 MMDevices base
- HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\{Render|Capture}\{endpoint-guid}\...
- HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\{Render|Capture}\{endpoint-guid}\...

Common subkeys:
- `FxProperties`
- `Properties`
- nested subkeys under FxProperties for plugins/user state

### 5.2 Listen keys
- Enable flag (PropertyStore): `{24dbb0fc-...}, pid=1`
- Routing target (HKLM registry value): `{24dbb0fc-...},0` under capture endpoint `Properties`

### 5.3 SysFX keys
- Disable_SysFx: `{E4870E26-...},2` (value name format `{fmtid},pid`)

### 5.4 Vendor toggles
- Vendor-defined `{fmtid},pid` values under either FxProperties or Properties.
- Hive selection (HKCU vs HKLM) varies by driver and by learned preference.

---

## 6) COM interface specifications (what is used and how)

### 6.1 IMMDeviceEnumerator
Used for:
- enumerating endpoints
- resolving devices by ID
- getting default endpoint IDs

### 6.2 IAudioEndpointVolume
Used for:
- volume get/set
- mute get/set

### 6.3 PolicyConfig / PolicyConfigVista
Used for:
- `SetDefaultEndpoint(device_id, role)`

### 6.4 PolicyConfigFx
Used for:
- reading/writing endpoint properties with bFxStore awareness
- diagnostics and discovery around Disable_SysFx

### 6.5 IPropertyStore (raw vtable)
Used for:
- Listen enable flag
- Disable_SysFx property store read/write
- friendly name fallback reads

---

## 7) Performance characteristics and timeouts

### 7.1 CLI
- Most operations are single COM calls + lightweight processing.
- Some operations intentionally poll:
  - listen verification registry polling (short, bounded)
  - enhancements verification in some contexts (bounded)
- `wait` polls at 0.5s intervals until timeout.

### 7.2 GUI
- Device listing:
  - one CLI call (`list --json`)
- State population:
  - multiple CLI calls (`get-device-state`) spread over Tk event loop ticks
- Context menu refresh:
  - optional quick JSON probe with short timeout (~0.75s)

---

## 8) Extensibility guide (with technical constraints)

### 8.1 Adding a CLI command (spec checklist)
- Must follow selection algorithm conventions.
- Must return stable JSON with a predictable top-level key.
- Must preserve exit code conventions.
- Must not introduce global COM init in cli.py.

### 8.2 Adding a GUI feature
- Must call CLI via subprocess.
- Must not import devices.py for direct COM work.
- Must update cache or schedule refresh after state changes.
- Must keep UI responsive (avoid blocking calls without timeouts or threading).

### 8.3 Adding new vendor schema fields
- Update INI parser with backward compatibility:
  - tolerate missing keys
  - keep existing behaviors unchanged by default
- Update apply path and fast read path.
- Update learn/merge logic if the field changes how entries are scoped.

### 8.4 Editing prompt strings (strict warning)
If you change learn prompt text:
- update GUI prompt matchers in `LearnRunner` / interactive patterns
- keep compatibility if possible (match old and new patterns)

---

## 9) Testing specification (engineer checklist)

### 9.1 Source-mode tests
- `python -m compileall .`
- `python -c "import audioctl.cli, audioctl.devices, audioctl.vendor_db, audioctl.gui"`

### 9.2 CLI smoke tests
- `audioctl list --json`
- `audioctl get-device-state --id <id> --flow Render`
- `audioctl set-volume --id <id> --flow Render --level 10`
- `audioctl set-volume --id <id> --flow Render --mute`
- `audioctl get-volume --id <id> --flow Render`

### 9.3 GUI smoke tests
- Launch GUI (no args).
- Refresh devices.
- Right-click a render endpoint: check mute label, enhancements label, FX menu.
- Right-click a capture endpoint: check listen label.
- Run a learn workflow only if your change impacts learn paths.

### 9.4 Frozen EXE tests (PyInstaller)
- Clean build.
- Run EXE from dist output folder.
- Confirm GUI launches and at least `list --json` works.
- Confirm no missing module errors.

---

## 10) Common failure modes & engineering diagnosis

### 10.1 COM crashes / access violations
Typical cause:
- COM object released on wrong thread or during raw pointer usage.

Engineering response:
- verify GC guards remain around raw vtable calls
- avoid long-lived COM objects
- enable `AUDIOCTL_DEBUG=1` and review log breadcrumbs

### 10.2 “Vendor toggle failed” / “No vendor toggle available”
Cause:
- INI has no entry for that endpoint GUID, or driver uses different keys.

Response:
- use learn flows
- use discovery report and inspect diffs

### 10.3 GUI learn doesn’t progress
Cause:
- prompt strings changed or not matched
- subprocess stdout buffering changes

Response:
- inspect learn runner logs
- confirm patterns in GUI match prompts in CLI

---

## 11) Appendix: Data structure definitions (copy/paste friendly)

### 11.1 Device dict (list output)
```json
{
  "id": "{0.0.0.00000000}.{...}",
  "name": "Speakers (Realtek(R) Audio)",
  "flow": "Render",
  "state": "active",
  "isDefault": {
    "console": true,
    "multimedia": true,
    "communications": false
  }
}
```

### 11.2 get-device-state output
```json
{
  "id": "{...}",
  "name": "Speakers (Realtek(R) Audio)",
  "flow": "Render",
  "volume": 78,
  "muted": false,
  "listenEnabled": null,
  "enhancementsEnabled": true,
  "availableFX": [
    { "fx_name": "BassBoost", "state": true, "source": "ini" }
  ]
}
```

### 11.3 vendor_toggles.ini main entry (example)
```ini
[vendor_example]
value_name = {vendor-guid},5
dword_enable = 0
dword_disable = 1
hives = HKCU,HKLM
flows = Render,Capture
subkey = FxProperties
notes = Example learned toggle
devices = {endpoint-guid-1},{endpoint-guid-2}
```

### 11.4 vendor_toggles.ini FX multi_write entry (example)
```ini
[fx_example_bucket]
type = fx
fx_name = BassBoost
multi_write = 1
write_count = 2
decider_index = 1
quorum_threshold = 0.60

write1_hive = HKCU
write1_subkey = FxProperties
write1_name = {vendor-guid},7
write1_type_enable = REG_DWORD
write1_type_disable = REG_DWORD
write1_enable = 1
write1_disable = 0
write1_devices = {endpoint-guid-1}

write2_hive = HKLM
write2_subkey = Properties
write2_name = {other-guid},1
write2_type_enable = REG_BINARY
write2_type_disable = REG_BINARY
write2_enable = hex:01
write2_disable = hex:00
; write2_devices missing => universal within this bucket

devices = {endpoint-guid-1}
```
