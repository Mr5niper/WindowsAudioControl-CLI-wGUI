# Technical Specifications & Engineering Guide  
**Audioctl (Windows Audio Control CLI + GUI) v1.5.0.0**
<br><br>
**Audience:** Engineers maintaining or extending audioctl  
**Platforms:** Windows 10/11 (x64 assumed; x86 may work if dependencies permit)  
**Core dependencies:** Python 3.14.3, pycaw, comtypes, ctypes, winreg, tkinter (GUI), PyInstaller (required for build)

This document describes the **entire program**, including architecture, module contracts, data formats, COM/GC stability constraints, registry touchpoints, INI schema, CLI/GUI behaviors, and engineering extension guidance.

---

## 1) System goals and non-goals

### 1.1 Goals
- Provide a **scriptable CLI** for repeatable Windows audio endpoint operations.
- Provide an optional **GUI** that uses the CLI (not direct COM calls) for consistent behavior.
- Support vendor-specific enhancements/effects toggles via a learned, persistent configuration database (`vendor_toggles.ini`).
- Offer read-only “query helpers” designed for fast, side-effect-free state retrieval.
- Prioritize **stability** in the face of COM lifetime hazards and driver variability.
- Support **FX (per-effect) toggles** via legacy single-DWORD and **multi-write** registry strategies.

### 1.2 Non-goals
- Cross-platform support (Windows-only by design).
- Per-application audio session control (endpoints only; not per-app mixers).
- Perfect vendor coverage without learning (vendor toggles are inherently driver-specific).

---

## 2) Architectural overview

### 2.1 Layer model

**Layer 1 - Entry / orchestration**
- `audioctl.py` (PyInstaller entrypoint wrapper; Windows console attach)
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
1. JSON output keys and shapes are API.
2. Prompt strings used for GUI interactive learn flows are API (pattern-based).
3. `--index` ordering must match GUI order.
4. Mutating commands must remain active-only unless intentionally changed.
5. COM initialization must be isolated and correct; avoid global COM init patterns.
6. GUI must remain COM-free: no importing low-level device helpers for direct use.

---

## 3) Repository layout & module specifications

### 3.1 `audioctl.py`
**Purpose:** Frozen EXE entrypoint and console attachment helper.  
**Responsibilities:**
- On Windows: attempt `AttachConsole(ATTACH_PARENT_PROCESS)` to reuse parent console (esp. older PowerShell behavior).
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
- GUI auto-launch when invoked with no arguments (`argv is None` and no CLI args).

**Required behaviors:**
- Mutating operations select only active endpoints (`include_all=False`).
- On ambiguity, return exit code `4` and require `--index` (except FX disambiguation which is interactive by design).

**CLI exit codes (contract):**
- `0` success
- `1` invalid args/runtime failure
- `3` not found / wait timeout
- `4` multiple matches / index required / index out of range / interactive selection invalid
- `130` Ctrl-C

#### 3.2.1 Command surface (spec)

- `list [--all] [--json]`  
  - JSON:  
    ```json
    { "devices": [ {id,name,flow,state,isDefault,guiIndex}, ... ] }
    ```
  - Note: devices are tagged with `guiIndex` (0-based within each flow) after GUI-order sorting.

- `set-default`  
  - Args: `--playback-id|--playback-name`, `--recording-id|--recording-name`, roles, `--regex`, `--index`  
  - JSON:
    ```json
    { "set": [ { "flow": "Render|Capture", "role": "console|multimedia|communications|all", "id": "...", "name": "..." }, ... ] }
    ```
  - Warning emitted to stderr if not elevated: *"might require Administrator privileges..."*

- `set-volume`  
  - Args: `--id|--name`, optional `--flow`, exactly one of `--level|--mute|--unmute`, optional `--index`, `--regex`  
  - JSON success:
    - `{ "volumeSet": {id,name,level} }` OR `{ "muteSet": {id,name,muted} }`
  - `--json` exists but current behavior always prints JSON on success; errors go to stderr.

- `get-volume`  
  - JSON:
    ```json
    { "id": "...", "name": "...", "flow": "Render|Capture", "volume": 0-100|null, "muted": true|false|null }
    ```

- `listen` (Capture only)  
  - Args:
    - capture selector: `--id|--name` (+ `--index`, `--regex`)
    - exactly one: `--enable|--disable`
    - optional playback routing target:
      - `--playback-target-id [ID]` (flag without value means default playback device via `const=''`)
      - `--playback-target-name [NAME]` (flag without value means default playback device via `const=''`)
  - JSON:
    ```json
    { "listenSet": { "id": "...", "name": "...", "enabled": true|false|null, "verifiedBy": "com|registry" } }
    ```
    `verifiedBy` is present only when the setter returned failure but state was confirmed anyway (driver/Windows timing quirks).

- `get-listen`  
  - JSON:
    ```json
    { "id": "...", "name": "...", "flow": "Capture", "listenEnabled": true|false }
    ```
  - Note: **fast path is registry-first with COM fallback; final fallback forces boolean False** (no `null`), per `_read_listen_enable_fast()`.

- `enhancements`  
  - Exactly one operation is required:
    - Main: `--enable|--disable|--learn`
    - FX subsystem: `--list-fx|--learn-fx|--enable-fx|--disable-fx|--delete-fx`
  - JSON varies by operation:
    - Main toggle:
      ```json
      { "enhancementsSet": { "id": "...", "name": "...", "enabled": true|false|null, "verifiedBy": "..." } }
      ```
    - Learn main:
      ```json
      { "vendorLearned": {...} }
      ```
      or
      ```json
      { "vendorAvailable": {...} }
      ```
    - FX learn:
      ```json
      { "fxLearned": {...} }
      ```
    - FX toggle:
      ```json
      { "fxSet": {...} }
      ```
    - FX delete:
      ```json
      { "fxDeleted": {...} }
      ```

  - FX enable/disable disambiguation:
    - If multiple FX name matches, CLI uses **interactive selection via stdin** (returns exit `4` on invalid selection). The GUI relies on this behavior for some flows.

- `get-enhancements` (vendor-only)  
  - JSON:
    ```json
    { "id": "...", "name": "...", "flow": "Render|Capture", "enhancementsEnabled": true|false|null }
    ```
  - Note: If a vendor method exists but read is inconclusive, CLI **default-assumes enabled** (`true`) for UI friendliness.

- `get-device-state` (GUI aggregator)  
  - JSON:
    ```json
    {
      "id": "...",
      "name": "...",
      "flow": "Render|Capture",
      "volume": 0-100|null,
      "muted": true|false|null,
      "listenEnabled": true|false|null,
      "enhancementsEnabled": true|false|null,
      "availableFX": [
        { "fx_name": "...", "state": true|false|null, "source": "ini" }
      ]
    }
    ```
  - Defaulting behavior for UI friendliness:
    - if vendor/FX entry exists but state read is inconclusive → default `true` (enabled) in this aggregator.

- Diagnostics/discovery:
  - `diag-sysfx` (JSON summary of live Windows + vendor toggles)
  - `diag-mmdevices` (JSON dump of all endpoint MMDevices registry values)
  - `discover-enhancements` (interactive; prints report and writes TXT/JSON bundle)
  - `wait` (poll until active endpoint appears; JSON on success)

- Hidden (GUI support):
  - `vendor-ini-append --work <json>`: elevated helper for appending INI entries in protected locations (supports `"kind": "main"` and `"kind": "fx"` work orders).

### 3.3 `audioctl/gui.py`
**Purpose:** Tkinter GUI frontend that shells out to CLI as subprocess.

**Responsibilities:**
- Display device list grouped by flow (Render/Capture).
- Context menu actions:
  - set default
  - set volume
  - mute/unmute (deterministic: queries state then chooses mute/unmute)
  - listen enable/disable (Capture only; queries state then chooses)
  - enhancements enable/disable (vendor-only; disabled if no vendor method)
  - FX toggles (learned; presented as Enable/Disable depending on cached state)
  - learn workflows (main + FX) via CLI interactive mode
- Background device state caching using `get-device-state`.

**Technical constraints:**
- Must not import/use low-level COM device helpers directly.
- Must tolerate CLI failure and display actionable errors.
- Must keep UI responsive:
  - incremental state population (`get-device-state` one device per tick)
  - quick JSON probe for context menu freshness (`run_audioctl_quick_json` timeout ~0.75s)
  - interactive learn runner mechanisms (see below)

#### 3.3.1 Learn flows in GUI (important update)
v1.5.0.0 contains **two learn-driving mechanisms**:
1. `run_audioctl_interactive(...)`: line-buffered; shows messageboxes when prompt substrings appear; then writes newline to CLI stdin.
2. `LearnRunner` (non-blocking): character-by-character stream parsing; designed to avoid UI freezing and to detect prompts without newline terminators.

Both exist; the code currently uses the interactive pattern-based approach in `_learn_main_toggle_via_cli` / `_learn_fx_toggle_via_cli`. A separate non-blocking controller exists (`_open_main_learn_dialog`) for main learn.

**Prompt-string compatibility warning:** GUI prompt matching relies on CLI text like:
- `"Step 1: ... set 'Audio Enhancements' to ENABLED ..."`
- `"Step 2: ... set 'Audio Enhancements' to DISABLED ..."`
- FX learn prompts: `"Set the '<fx>' effect to ENABLED..."`, `"Now set ... to DISABLED..."`, `"Enable ... again (second pass)"`, `"Disable ... again (second pass)"`

Changing these strings requires updating GUI match logic.

#### 3.3.2 State cache specification
- `device_state_cache: Dict[device_id -> state_dict]`
- state_dict shape matches `get-device-state`.

Cache population:
- After `Refresh`, the GUI:
  - runs `audioctl list --json`
  - then schedules incremental `get-device-state` calls

Context menu:
- Uses cached state immediately.
- Then performs a fast one-shot `get-device-state` refresh right before showing menu (timeout-limited).

### 3.4 `audioctl/devices.py`
**Purpose:** Low-level Windows integration layer (COM + registry).

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
- `_sort_and_tag_gui_indices(devices) -> {Render: [...], Capture: [...]}`
  - mutates devices by adding `guiIndex`

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
  },
  "guiIndex": 0
}
```

**Naming strategy update:**  
`list_devices()` now prefers names from `AudioUtilities.GetAllDevices()` (name map) to reduce use of raw PropertyStore reads (stability).

#### 3.4.3 Default endpoint management
**API:**
- `set_default_endpoint(device_id: str, role: str) -> None`
- Roles: `console`, `multimedia`, `communications`, `all`

**Constraints:**
- Must refuse inactive endpoints (`_is_device_active` check).

#### 3.4.4 Volume/mute
**API:**
- `get_endpoint_volume(device_id) -> int|None` (0..100)
- `set_endpoint_volume(device_id, level_percent) -> bool`
- `get_endpoint_mute(device_id) -> bool|None`
- `set_endpoint_mute(device_id, mute_state: bool) -> bool`

**Normalization rules:**
- Volume scalar float 0.0..1.0 → percent 0..100.
- Mute returns `bool` when readable; may return `None` on unrecoverable variant mismatch.

#### 3.4.5 Listen feature
**Enable flag (PropertyStore):**
- PROPERTYKEY fmtid `{24dbb0fc-9311-4b3d-9cf0-18ff155639d4}`, pid `1`
- Written via raw `IPropertyStore::SetValue + Commit`

**Routing target (registry, HKLM):**
- `HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\Capture\{endpoint-guid}\Properties`
- Value name `{24dbb0fc-9311-4b3d-9cf0-18ff155639d4},0` (REG_SZ)
- Empty string means “Default Playback Device”.

**Verification behavior:**
- Setter returns True/False based on PropertyStore write success.
- CLI will verify via COM read first, then registry polling (`_verify_listen_via_registry`) if needed.

**Fast read update (important):**
- `_read_listen_enable_fast()`:
  1. COM property store read
  2. registry decode fallback
  3. final fallback returns `False` (prevents `null` in JSON).

#### 3.4.6 SysFX (Disable_SysFx) helpers
**Disable_SysFx PROPERTYKEY:**
- fmtid `{E4870E26-3CC5-4CD2-BA46-CA0A9A70ED04}`, pid `2`
- Semantics:
  - 0 → Enhancements enabled
  - 1 → Enhancements disabled

**Read paths (diagnostic/discovery):**
- PolicyConfigFx (`GetPropertyValue` with `bFxStore=True/False`)
- Endpoint PropertyStore (raw vtable)
- Registry scanning (snapshot tooling)

**Write paths (diagnostic/learn tooling):**
- PolicyConfigFx (`SetPropertyValue`)
- Endpoint PropertyStore (`SetValue` + `Commit`)
- Registry writes (`_set_enhancements_registry`) mainly as fallback/test tooling

**Runtime policy note:** Main enhancements toggling in normal operation is vendor-only via `vendor_db.py`. Windows Disable_SysFx setters are for diagnostics/learning, not runtime toggling.

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
- `dword_flips`: 0↔1 flips (strong vendor toggle candidates)
- `disable_sysfx_hits`: entries involving Disable_SysFx fmtid

#### 3.4.8 PropertyStore raw vtable subsystem
`_get_property_store_interfaces()` builds a cached ctypes vtable definition.

**Stability requirements:**
- Must remain cached at module scope.
- Calls using raw pointers should remain GC-guarded (GC temporarily disabled during vtable access).
- `PropVariantClear` must be called on allocated PROPVARIANTs.

### 3.5 `audioctl/vendor_db.py`
**Purpose:** Vendor toggles persistence + application + learning.

#### 3.5.1 INI file location
Default `vendor_toggles.ini` path:
- EXE directory if writable  
- else `%LOCALAPPDATA%\audioctl\vendor_toggles.ini`

#### 3.5.2 INI cache
Cache key:
- absolute path
- file mtime

**Behavior:**
- If file missing: cache “missing” and return empty DB.
- If mtime unchanged: reuse parsed DB.

#### 3.5.3 Split DB model
Loader returns:
```python
{ "main": [ ... ], "fx": [ ... ] }
```

#### 3.5.4 MAIN entry schema (Enhancements)
Required fields:
- `value_name` (normalized lower)
- `dword_enable`, `dword_disable` (int; typically 0/1)
- `hives` list (HKCU/HKLM; ordering is preference)
- `flows` list (Render/Capture)
- `devices` list (endpoint GUIDs)
- `subkey` (`FxProperties` or `Properties`) — learned scope

**Applicability rule (runtime-truth update):**
- v1.5.0.0 now relies heavily on **signature truth**:
  - `_find_first_vendor_entry(...)` selects the **first MAIN entry whose registry signature matches this endpoint now**.
  - This is **not** purely GUID membership-based at runtime; the registry value must exist and match known enable/disable values.
- This reduces false positives and supports dedupe across endpoints/drivers.

#### 3.5.5 FX schema (legacy single DWORD)
Fields:
- `type = fx`
- `fx_name`
- `value_name`
- `dword_enable`, `dword_disable`
- `devices` list (may exist, but runtime discovery can also use signature matching)

#### 3.5.6 FX schema (multi_write)
Fields:
- `type = fx`
- `fx_name`
- `device_name_pattern` (optional; used for readability/spoof/candidate filters)
- `multi_write = 1`
- `write_count = N`
- `decider_index` (1-based)
- `quorum_threshold` (0.50..0.95; default 0.60)
- Per write block i:
  - `write{i}_hive = HKCU|HKLM`
  - `write{i}_subkey = FxProperties|Properties`
  - `write{i}_name = {fmtid},pid`
  - `write{i}_type_enable/disable = REG_DWORD|REG_SZ|REG_BINARY`
  - `write{i}_enable/disable = payload`
  - optional: `write{i}_devices`

**write{i}_devices semantics (documented & implemented):**
- missing => universal within bucket
- empty => applies to nobody
- list => applies only to listed GUIDs

**Important behavior note (v1.5.0.0):**  
For **matching/reading/verifying/applying**, the code frequently treats signature as truth and may **ignore per-write devices gating** in some paths, instead applying only to values that actually exist for the endpoint GUID. In other words:
- It won’t “invent” missing keys.
- It may still apply a write if the value exists on the device and matches known enable/disable signatures.

#### 3.5.7 Vendor-only runtime policy
- `_apply_enhancements` is vendor-only.
- If no vendor entry applies: return failure (`no-vendor-method`).

#### 3.5.8 Fast read semantics (GUI support)
Fast reads do:
- single registry probe(s)
- optional alternate hive probe
- last-write-time tie-break if HKCU/HKLM disagree

For multi_write FX fast reads:
- selects a “best” write signal (FxProperties + DWORD favored)
- uses type-aware signature comparison

#### 3.5.9 Learning algorithms (spec)

**Enhancements learn (manual UI toggling):**
- A/B snapshot capture while user toggles UI.
- Diff to find stable DWORD flips (`dword_flips`).
- Write INI entry with learned `subkey` (FxProperties vs Properties).
- Dedupe:
  - if an identical main entry exists, append GUID to its `devices`
  - also maintain optional name→GUID buckets (`name_<hash>`, `guids_<hash>`) for readability

**FX learn (two-pass A/B/A2/B2):**
- Guided interactive capture:
  - Pass 1: A/B “prime” driver
  - Pass 2: A2/B2 authoritative pair written into INI
- Uses stability filtering and can generate multi-write sets (DWORD/SZ/BINARY).
- Merge behavior:
  - match identity+payload: append GUID to `write{i}_devices`
  - new payload: append a new write block scoped to this GUID
  - conflict cleanup prevents a GUID from attaching to competing writes with same identity but different payload

**FX delete (new in doc; present in code):**
- Removes GUID association for a specific FX bucket without deleting the whole bucket:
  - updates section `devices = ...`
  - updates or introduces `write{i}_devices` to exclude the GUID where needed
  - preserves write blocks even if they become scoped to nobody

### 3.6 `audioctl/compat.py`
**Purpose:** Load-order critical compatibility shim.

**Requirements:**
- Must be imported before any comtypes usage.
- Ensures:
  - `PROPVARIANT` alias exists (if only `tagPROPVARIANT` exists)
  - VT constants exist (`VT_BOOL`, `VT_LPWSTR`, `VT_UI2`, `VT_UI4`)
  - `VARIANT_TRUE` / `VARIANT_FALSE` exists
- Forces import of:
  - `comtypes._post_coinit`
  - `comtypes._post_coinit.unknwn`
  - (and other hiddenimports are included in PyInstaller spec/bat)

### 3.7 `audioctl/logging_setup.py`
**Purpose:** Lazy log file initialization + crash breadcrumbs.

**Specification:**
- No filesystem writes at import time.
- On first log call:
  - resolve log path (prefer exe dir; else temp dir)
  - install:
    - `sys.excepthook`
    - `sys.unraisablehook`
    - Windows console control handler (best-effort)
  - enable `faulthandler` to log file if possible (`all_threads=True`)
  - wrap `sys.exit` to record breadcrumb
  - register atexit breadcrumbs and close handles

---

## 4) CLI technical specification (device selection & ordering)

### 4.1 Selection algorithm (standard pattern)
Given a command with selectors:
1. Enumerate devices:
   - mutating commands: `list_devices(include_all=False)`
   - list command: optional `include_all`
2. Tag devices with GUI order (`guiIndex`):
   - sort by `name.lower()` within each flow bucket
3. Filter matches:
   - by ID exact match, OR
   - by name substring/regex match, optional flow filter
4. If multiple matches:
   - require `--index` (except some FX sub-ops which may prompt interactively)
5. Choose target based on index or first match.
6. Execute operation and return JSON.

### 4.2 GUI-order sorting definition
Within each flow bucket:
- Sort by `name.lower()`
- Assign `guiIndex` sequentially starting at 0

---

## 5) Registry touchpoints (technical map)

> Registry paths shown here are conceptual; do not copy/paste these as Python strings unless you escape properly.

### 5.1 MMDevices base
- `HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\{Render|Capture}\{endpoint-guid}\...`
- `HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\{Render|Capture}\{endpoint-guid}\...`

Common subkeys:
- `FxProperties`
- `Properties`
- Nested subkeys under FxProperties for plugins/user state

### 5.2 Listen keys
- Enable flag (PropertyStore): `{24dbb0fc-...}, pid=1`
- Routing target (HKLM registry value): `{24dbb0fc-...},0` under capture endpoint `Properties`

### 5.3 SysFX keys
- Disable_SysFx: `{E4870E26-...},2` (value name format `{fmtid},pid`)

### 5.4 Vendor toggles
- Vendor-defined `{fmtid},pid` values under either FxProperties or Properties.
- Hive selection (HKCU vs HKLM) varies by driver and learned preference.

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
- diagnostics/discovery around Disable_SysFx

### 6.5 IPropertyStore (raw vtable)
Used for:
- Listen enable flag
- Disable_SysFx property store read/write (diagnostic/learn)
- Friendly name fallback reads (rare; name map preferred)

---

## 7) Performance characteristics and timeouts

### 7.1 CLI
- Most operations are single COM calls + lightweight processing.
- Some operations intentionally poll:
  - listen verification registry polling (short, bounded)
  - vendor entry verification polling (bounded)
- `wait` polls at 0.5s intervals until timeout.

### 7.2 GUI
- Device listing:
  - one CLI call (`list --json`)
- State population:
  - multiple CLI calls (`get-device-state`) spread over Tk loop ticks
- Context menu refresh:
  - optional quick JSON probe with short timeout (~0.75s)

---

## 8) Extensibility guide (with technical constraints)

### 8.1 Adding a CLI command (spec checklist)
- Must follow selection algorithm conventions.
- Must return stable JSON with a predictable top-level key.
- Must preserve exit code conventions.
- Must not introduce global COM init in `cli.py`.

### 8.2 Adding a GUI feature
- Must call CLI via subprocess.
- Must not import `devices.py` for direct COM work.
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
- update GUI prompt matchers (`LearnRunner` and/or `run_audioctl_interactive` patterns)
- preserve backward compatibility where possible (match old + new patterns)

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
- `audioctl get-listen --id <capture-id>`
- `audioctl get-enhancements --id <id> --flow Render`

### 9.3 GUI smoke tests
- Launch GUI (no args).
- Refresh devices.
- Right-click a render endpoint: verify mute label, enhancements label, FX submenu state.
- Right-click a capture endpoint: verify listen label.
- Run learn workflow only if changes impact learn paths.

### 9.4 Frozen EXE tests (PyInstaller)
- Clean build.
- Run EXE from dist output folder.
- Confirm GUI launches and `list --json` works.
- Confirm no missing module errors related to `comtypes._post_coinit*`.
- Confirm console attach behavior when launched from an existing terminal.

---

## 10) Common failure modes & engineering diagnosis

### 10.1 COM crashes / access violations
Typical cause:
- COM object released on wrong thread or during raw pointer usage.

Engineering response:
- Verify GC guards remain around raw vtable calls.
- Avoid long-lived COM objects.
- Enable `AUDIOCTL_DEBUG=1` and review `audioctl_gui.log` breadcrumbs.
- Prefer name map path (`AudioUtilities.GetAllDevices`) over per-device PropertyStore name reads.

### 10.2 “Vendor toggle failed” / “No vendor toggle available”
Cause:
- No matching vendor entry signature exists for this endpoint at runtime.

Response:
- Run main learn:
  - `audioctl enhancements --id <id> --flow <Render|Capture> --learn`
- Use discovery tooling:
  - `audioctl discover-enhancements --id <id> --flow <...> --output-dir <...>`
- Inspect generated report and diffs (look for `dword_flips`).

### 10.3 GUI learn doesn’t progress
Cause:
- Prompt strings changed or not matched.
- Subprocess output buffering/prompt-without-newline.

Response:
- Confirm GUI patterns match CLI prompt strings.
- Use `LearnRunner` (char-wise reading) if prompts appear without newline.
- Inspect logs (`audioctl_gui.log`) for subprocess output and state transitions.

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
  },
  "guiIndex": 0
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
device_name_pattern = Speakers
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
