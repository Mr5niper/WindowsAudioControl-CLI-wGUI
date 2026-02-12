# ENGINEERING GUIDE (v1.5.0.0)
**audioctl — Windows Audio Control CLI + GUI (PyInstaller EXE)**

**Audience:** Engineers debugging, maintaining, or extending the codebase  
**Platforms:** Windows 10/11  
**Version:** 1.5.0.0  
**Build contract:** **Python 3.13.12 required**, **PyInstaller required**

This document is intentionally **not** a usage manual. The README covers commands, examples, and screenshots.  
This file focuses on how the program is built, how it works internally, where it breaks, and how to extend it safely.

---

## 1) Build & runtime contract (non-negotiable)

### 1.1 Python version
- Development/build is pinned to **Python 3.13.12**.
- `BUILD_EXE.bat` performs a hard preflight check:
  - build fails if `python --version` is not exactly `3.13.12`.

### 1.2 PyInstaller is required
This project is shipped/tested as a **frozen EXE**. Source-mode runs exist for development parity, but the supported artifact is:
- `dist\audioctl.exe`

Packaging concerns that are *part of the engineering contract*:
- COM-related dynamic imports must be bundled reliably.
- `comtypes._post_coinit*` modules must be present in the frozen app to avoid shutdown/finalization instability.
- `pycaw` and `comtypes` must be collected with PyInstaller (hooks miss things otherwise).

### 1.3 Entry targets and parity
There are three entry paths; they must behave the same:
- **Frozen EXE**: `audioctl.exe` (from `audioctl.py` as PyInstaller script target)
- **Module run**: `python -m audioctl` (via `audioctl/__main__.py`)
- **Direct**: `python audioctl.py`

The *real* entrypoint is always:
- `audioctl.cli.main()`

### 1.4 Build artifacts / metadata
- Windows version info comes from `version.txt` (filevers/prodvers 1.5.0.0).
- Icon `audio.ico` is embedded and also shipped as data.

---

## 2) Architecture: who owns what (and why)

### 2.1 CLI-first, GUI-as-client
The core design decision:
- **GUI never performs COM/registry operations directly.**
- GUI shells out to the CLI and consumes JSON.

Reason: COM object lifetime + GC timing + thread affinity is a real stability hazard in long-lived GUI processes.  
CLI calls are short-lived and wrap COM init/cleanup per operation.

### 2.2 Layering (actual dependency direction)
- `audioctl/cli.py` is the **public API router** (JSON contract + exit codes + prompts).
- `audioctl/gui.py` is a **subprocess client** of the CLI.
- `audioctl/devices.py` is the **low-level Windows engine** (COM + PropertyStore + registry snapshot tooling).
- `audioctl/vendor_db.py` is the **vendor/FX rule engine** (INI parsing/cache + registry writes + learn/merge/delete logic).
- `audioctl/compat.py` must run **before** any comtypes/pycaw usage.
- `audioctl/logging_setup.py` is global best-effort logging; must never break runtime.

### 2.3 Data flow (canonical)
GUI action → spawn CLI subprocess → CLI resolves device → devices/vendor_db performs operation → CLI prints JSON → GUI parses JSON → updates UI + cache

This is intentional “backend boundary” design. If you bypass it, you will reintroduce COM lifetime instability.

---

## 3) CLI/GUI compatibility contracts (things you must not casually break)

### 3.1 JSON shapes are API
The GUI depends on these commands and expects these keys (minimum):
- `list --json` → `{"devices":[...]}` and devices include `id,name,flow,state,isDefault` plus `guiIndex` tagging behavior.
- `get-device-state` → includes `volume,muted,listenEnabled,enhancementsEnabled,availableFX`.
- `enhancements --list-fx --json` shape (GUI uses it for menus in some paths).
- `listen` may include `verifiedBy` when verification fallback is used.

If you change JSON keys, you must update GUI parsing and any external scripting users.

### 3.2 Prompt strings are API (learn flows)
The GUI has code that matches CLI output text during learn workflows.
If you change the CLI prompt lines, the GUI can stop progressing.

Prompts the GUI relies on (representative):
- MAIN learn:
  - `set 'Audio Enhancements' to ENABLED`
  - `set 'Audio Enhancements' to DISABLED`
- FX learn:
  - `Set the '<fx>' effect to ENABLED`
  - `Now set the '<fx>' effect to DISABLED`
  - `Enable the '<fx>' effect again (second pass)`
  - `Disable the '<fx>' effect again (second pass)`

Rule: if you edit learn text, update the GUI matcher patterns at the same time.

### 3.3 `--index` ordering must match GUI ordering
The CLI uses “GUI order” as:
- sort by `name.lower()` **within each flow bucket**
- assign `guiIndex` 0..N per flow

Commands that accept `--index` use this same ordering to disambiguate. This is why “index N” is stable between CLI and GUI.

---

## 4) GUI internals (what actually happens)

### 4.1 How device list is built
On Refresh:
1. GUI runs:
   - `audioctl list --json` (plus `--all` if enabled)
2. GUI splits devices by `flow` into:
   - Render group
   - Capture group
3. GUI inserts group header rows + device rows into the Treeview.
4. GUI does *not* trust internal COM state—everything comes from CLI output.

Important: the GUI currently displays its own per-group index in the Treeview, but the CLI also tags `guiIndex`. The contract that matters for CLI disambiguation is the CLI’s GUI-order index semantics (name-sorted per flow).

### 4.2 device_state_cache: why it exists
The context menu labels depend on state:
- Mute vs Unmute
- Enable vs Disable Listen (Capture only)
- Enable vs Disable Enhancements (vendor-only)
- FX menu entries (Enable/Disable per FX)

Querying each of these by spawning a CLI process on every right click is slow and sometimes blocks on drivers.  
So the GUI maintains:

- `device_state_cache[device_id] = (get-device-state JSON)`

### 4.3 Incremental state population (non-blocking)
After refresh builds the Treeview:
- GUI schedules repeated calls (one device per tick):
  - `audioctl get-device-state --id <id> --flow <flow>`

This is intentionally incremental so the UI doesn’t freeze.

### 4.4 Context menu build logic (stale-cache resistant)
When you right-click:
1. Build menu immediately using cached state if present.
2. Then do a **single quick probe** before showing the menu:
   - `run_audioctl_quick_json(["get-device-state", ...], timeout=0.75)`
3. Update menu labels with that “just-in-time” state.

This is the compromise between:
- responsiveness (don’t block UI)
- correctness (reduce stale labels)

### 4.5 GUI actions: how state is kept consistent
After GUI performs a mutating action (mute/listen/enhancements/fx):
- It updates the relevant portion of the cached state immediately via `_ensure_device_state_entry`
- It does **not** wait for the background refresh to catch up before showing the updated label next time

This is why you must keep `get-device-state` shape stable: the cache is treated as the canonical model.

### 4.6 Learn workflows (what the GUI really does)
There are two mechanisms in the codebase:

**A) Messagebox-driven interactive runner (currently used for main learn + FX learn)**
- `run_audioctl_interactive(...)` spawns CLI with stdin/stdout pipes
- reads stdout line-by-line
- when a prompt substring appears, shows a dialog
- when user clicks OK, writes `\n` to CLI stdin to continue

**B) LearnRunner (non-blocking, char-wise streaming)**
- `LearnRunner` reads stdout/stderr character-by-character
- detects prompt text even if it does not end with newline
- can auto-confirm the CLI confirmation prompt by sending `I UNDERSTAND\n` once
- designed to avoid freezing the Tk thread

Engineering note:
- If learn prompts ever stop ending in newlines, mechanism (A) becomes unreliable and you should move learn flows to (B).

### 4.7 “Print CLI commands” suppression during learn
During learn workflows, GUI disables “Print CLI commands” to avoid spamming stdout while the CLI is also running interactive output.

---

## 5) COM and GC safety rules (the part that prevents crashes)

### 5.1 Never do long-lived COM in the GUI
GUI must stay “pure Tk” + subprocess calls.  
If you add a GUI feature and you import `devices.py` for direct COM calls, you will eventually get:
- hard crashes (access violations)
- unraisable exceptions at shutdown
- COM apartment/thread mismatch issues

### 5.2 `_com_context()` is the correct COM lifetime model
`devices.py` uses:
- thread-local refcounted `CoInitialize()` / `CoUninitialize()`

Rules:
- Every COM entrypoint should run inside `_com_context()`
- Nested helpers must not `CoUninitialize` early
- Errors are best-effort; return values indicate success/failure

### 5.3 Raw PropertyStore vtable calls must be GC-guarded
Several operations use raw COM vtables via ctypes for stability and layout control:
- Listen enable flag read/write
- Disable_SysFx PropertyStore read/write
- Friendly name fallback reads

These operations:
- temporarily disable the Python GC while raw pointers are in use
- clear PROPVARIANTs using `PropVariantClear`

If you touch these paths:
- keep the GC guard pattern
- keep interface definitions cached (see below)

### 5.4 Interface definition caching is mandatory
`devices.py` caches:
- PolicyConfigFx interface definitions
- PropertyStore interface bundle (`_PROPERTY_STORE_INTERFACES_CACHE`)
- PolicyConfig fallback interface definitions

The reason is not micro-optimization; it is crash prevention:
- repeatedly defining ctypes function prototypes and comtypes interfaces can trigger GC/finalizers at unsafe times in frozen builds

---

## 6) Registry model (what you should know when debugging)

### 6.1 Endpoint ID vs endpoint GUID
Windows endpoints are identified by a full IMMDevice ID, but MMDevices registry paths are keyed by the GUID embedded in that ID.

Example:
- Endpoint ID:
  - `{0.0.0.00000000}.{83a9be54-901e-4429-993b-c9088e3028a0}`
- Endpoint GUID used in registry paths / INI:
  - `{83a9be54-901e-4429-993b-c9088e3028a0}`

`devices.py` extracts this via:
- `_extract_endpoint_guid_from_device_id()`

### 6.2 MMDevices base paths
Conceptual:
- `HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\{Render|Capture}\{GUID}\{FxProperties|Properties}`
- `HKLM\...` same structure

Drivers may store toggles under either:
- `FxProperties`
- `Properties`
- nested plugin keys under FxProperties

### 6.3 Listen toggle touchpoints
- Enable flag: PropertyStore key `{24dbb0fc-9311-4b3d-9cf0-18ff155639d4}`, pid `1`
- Playback routing target: registry value name `{24dbb0fc-...},0` under **HKLM Capture ... Properties**
  - empty string means “Default Playback Device”
  - may require Admin

### 6.4 Disable_SysFx (diagnostic tooling)
- PROPERTYKEY fmtid `{E4870E26-3CC5-4CD2-BA46-CA0A9A70ED04}`, pid `2`
- Semantics: 0 enhancements ON, 1 enhancements OFF

In v1.5.0.0:
- used for diagnostics/discovery/learn tooling
- **not** used as runtime toggle mechanism (runtime is vendor-only)

---

## 7) Vendor toggle engine (how the “real toggles” work)

You have a dedicated Vendor Toggles guide; this section only covers code-level “what matters when extending/debugging”.

### 7.1 Vendor-only runtime policy
For MAIN enhancements toggling:
- if no applicable vendor method matches this endpoint, runtime returns failure (no Windows fallback)

This behavior lives in:
- `vendor_db._apply_enhancements()`

### 7.2 “Registry truth” applicability (critical v1.5 behavior)
MAIN selection is based on signature match:
- value exists at the learned `subkey`
- value equals either enable or disable payload

The code path:
- `_find_first_vendor_entry()` → `_main_entry_signature_applies()`

This matters because:
- INI “devices = ...” lists are not sufficient to guarantee correctness at runtime
- entries are treated as a rule library; the registry decides what applies

### 7.3 FX discovery is not only “devices list”
FX listing does:
1. include FX where endpoint GUID is explicitly listed in the section `devices = ...`
2. also include FX whose live registry signature matches (even if GUID not listed)

This is why FX can appear “magically” after driver changes even before the INI is updated.

### 7.4 Multi-write FX verification model
Multi-write FX state isn’t always readable from one value.
`_read_decider_state()` attempts:
- quorum voting across write blocks
- fallback “best signal” (prefers FxProperties + DWORD-like toggles)

The GUI uses a **fast** read variant for responsiveness:
- `_fast_read_vendor_entry_state()`

### 7.5 FX delete behavior (bucket-safe)
`--delete-fx` removes this device’s association from a bucket without deleting the bucket section (buckets may be shared across devices).  
This logic is in:
- `_delete_fx_for_guid()`

---

## 8) Logging & diagnostics (how to debug real bugs)

### 8.1 Logging model
Logging is lazy and best-effort:
- importing logging module must not do file I/O
- logging must not crash the program even if the filesystem is locked down

Log file name:
- `audioctl_gui.log`

Location:
- prefers EXE directory if writable
- else falls back to a temp directory (`%TEMP%\audioctl\audioctl_gui.log` style behavior)

### 8.2 Enable debug logs
Set:
- `AUDIOCTL_DEBUG=1`

This turns on `_dbg(...)` breadcrumbs especially useful for:
- COM pointer creation timing
- enumeration sequences
- suspicious failures that disappear when run under a debugger

### 8.3 Use discovery tools for vendor issues
If “Enhancements toggle failed” or “no vendor method”:
- `diag-mmdevices` to dump current registry reality for the endpoint
- `discover-enhancements` to produce A/B snapshots + diff report + JSON bundle

These tools are designed to create sharable artifacts when debugging user machines.

---

## 9) Extending audioctl safely (engineering checklists)

### 9.1 Adding a new CLI command
- Implement selection using the standard active-only path and GUI-order indices.
- Output a stable JSON object (single top-level key).
- Use exit codes consistently (0/1/3/4).
- Do not introduce global `CoInitialize` in `cli.py`.

### 9.2 Adding a new GUI feature
- Must call the CLI via subprocess (`run_audioctl` / `run_audioctl_quick_json`).
- Must be tolerant of timeouts and failures (do not freeze UI).
- If action changes state, update `device_state_cache` immediately (don’t rely on background refresh).

### 9.3 Changing learn workflows
- Treat CLI prompt strings as API. If changed:
  - update GUI match patterns
  - prefer `LearnRunner` if prompts might not end in newlines

### 9.4 Vendor DB schema changes
- Keep INI parsing backward compatible:
  - tolerate missing keys
  - default conservatively
- Update both:
  - “full read/verify” path
  - “fast read” path used by GUI

---

## 10) Common engineering failure modes (and what to do)

### 10.1 Frozen EXE works in source but fails when frozen
Likely cause:
- missing dynamic imports / hiddenimports for comtypes/pycaw modules

Fix:
- ensure `collect_all('pycaw')` and `collect_all('comtypes')` remain in build
- ensure `comtypes._post_coinit*` hiddenimports remain

### 10.2 Random crashes on exit / unraisable exceptions
Likely cause:
- COM Release during GC finalization with raw pointers in play
- late import of comtypes cleanup modules in frozen build

Fix:
- preserve compat forced imports (`comtypes._post_coinit*`)
- preserve GC guards around vtable calls
- keep interface definitions cached

### 10.3 GUI labels wrong (Mute/Listen/Enhancements/FX)
Likely cause:
- stale `device_state_cache` or quick probe timed out

Debug:
- run `get-device-state` manually for that endpoint
- increase quick probe timeout only if you accept UI lag
- confirm device ID stability and that you’re querying the correct flow

### 10.4 “Enhancements shows enabled when unknown”
This is an intentional UI policy in `get-device-state`/`get-enhancements`:
- if a vendor method exists but read is inconclusive, default-assume enabled (`true`)

If you change this, you must update GUI expectations (and you may make UI feel wrong on drivers that only store “disabled” markers).

---
