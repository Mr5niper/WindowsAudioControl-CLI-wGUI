# Technical Specifications: audioctl — Windows Audio CLI/GUI Utility (v1.4.7.1)

## 1. Introduction
audioctl is a Windows audio control utility with:
- CLI for scripting/automation/diagnostics
- Tkinter GUI for interactive control

Core features:
- Enumerate playback/recording endpoints with friendly names and default flags
- Set default endpoints by role
- Read/set master volume and mute
- Enable/disable “Listen to this device” (Capture) with explicit routing
- Enable/disable Audio Enhancements (SysFX) via vendor-first methods, including learned multi-write FX
- Diagnostics/discovery helpers
- Read-only query helpers for GUI/scripts (volume, mute, listen, enhancements, FX)

This document describes v1.4.7.1.

---

## 2. Dependencies
- Python 3.x
- comtypes
- ctypes
- pycaw
- tkinter (GUI)
- winreg
- PyInstaller (optional binary distribution)
- Assets: audio.ico, version.txt, vendor_toggles.ini (user-learned database)

---

## 3. COM Strategy & Stability
- Thread-local _com_context in devices.py:
  - Each low-level call enters/leaves a COM apartment using comtypes.CoInitialize()/CoUninitialize().
  - No global CoInitialize in CLI entry; GUI also avoids persistent singletons.
- Interface definition caching:
  - PropertyStore raw vtable types are defined once (_get_property_store_interfaces()).
  - PolicyConfigFx (bFxStore) and Vista variants are defined once or loaded from pycaw when available.
  - Prevents GC-time races while ctypes constructs vtables.
- comtypes post-coinit modules:
  - compat.py forces import of comtypes._post_coinit.* so PyInstaller bundles cleanup modules; avoids late import during shutdown.

---

## 4. Device Enumeration & Naming
- list_devices() enumerates IMMDevice endpoints for Render/Capture.
- Friendly names:
  - Prefer pycaw.AudioUtilities.GetAllDevices() map.
  - Fallback: hardened IPropertyStore string read (cached interface types; GC guard; AddRef/Release).
- Indices:
  - Name-sorted per-flow (Render/Capture); CLI and GUI share the same ordering for --index.

---

## 5. Default Endpoint Management
- set_default_endpoint():
  - Uses PolicyConfig (Vista/Client) to set console, multimedia, communications, or all.
  - Refuses inactive targets.
  - Per-role error reporting.

---

## 6. Volume & Mute
- IAudioEndpointVolume via IMMDevice.Activate.
- Robust tuple/out-parameter handling; volume normalized to 0..100; mute -> bool or None.

---

## 7. “Listen to this device” (Capture)
- Enable/disable checkbox:
  - IPropertyStore::SetValue/Commit on key {24dbb0fc-...}, pid=1.
- Routing target:
  - HKLM\...\MMDevices\Audio\Capture\{guid}\Properties
  - Value: {24dbb0fc-9311-4b3d-9cf0-18ff155639d4},0
  - REG_SZ = full render device ID, or empty string for “Default Playback Device”
  - May require Administrator.
- Read/verify:
  - Primary: IPropertyStore::GetValue
  - Fallback: robust HKCU scan for equivalent properties and types
  - Registry poll loop used for verification on ambiguous results

---

## 8. Enhancements (SysFX) — Vendor-first
- Runtime toggling is vendor-only:
  - Main on/off switch relies on vendor DWORDs learned from discovery (no Windows fallback in production toggles).
  - Built-ins can exist, but priority is INI (user-learned).
- INI database:
  - Default path: exe dir if writable; else %LOCALAPPDATA%\audioctl\vendor_toggles.ini
  - Cached by (path, mtime) to avoid reparsing on each call.
- Main entries (“main”):
  - value_name, dword_enable, dword_disable, hives (HKCU/HKLM), flows (Render/Capture), subkey (FxProperties/Properties, recorded during learn), devices (per-endpoint membership).
- FX entries (“fx”):
  - Single DWORD or multi-write blocks with write_count and write{i}_* keys.
  - write{i}_devices allows per-toggle scoping (None=universal, []=none, list=[guid…]).
  - decider_index and quorum_threshold control verification and fast state reads.
- Fast reads:
  - _fast_get_enhancements_state (main) and _fast_read_vendor_entry_state (fx/main):
    - Probe configured hive; if inconclusive, also probe alternate hive and compare last-write times to break ties.
    - For main entries, if inconclusive but the entry applies, default to True (enabled) for GUI readability.

---

## 9. Learning (Manual & FX)
- Manual learn (main toggle):
  - Captures registry A/B snapshots while you toggle Windows UI.
  - Writes a vendor section with recorded subkey location.
  - CLI/GUI flows both supported; GUI suppresses “Print CLI commands” during learn.
- Learn FX:
  - Two-pass A/B (A, B, then A2, B2) stabilizes driver behaviors that initialize on first toggle.
  - Stable maps and multi-write extraction produce robust write sets ordered by strength (FxProperties preferred; DWORD signals preferred).
  - Automatic per-toggle device scoping and conflict cleanup within the bucket.

---

## 10. Query Helpers
- get-volume: volume/muted
- get-listen: fast listenEnabled state (Capture)
- get-enhancements: vendor-only main state
- get-device-state: aggregated snapshot (volume, mute, listen, enhancements, availableFX + states)

---

## 11. GUI Overview
- Device list grouped by flow; GUI indices match CLI.
- Context menu:
  - Set Default (all roles)
  - Set Volume…
  - Mute/Unmute
  - Toggle Listen (Capture)
  - Enable/Disable Enhancements (when learned)
  - Enhancement Effects (FX toggles from INI)
  - Learn Enhancements (guided)
- Background state cache (get-device-state) for accurate labels and instant menus.
- Modal learn flows with safe suppression of CLI echoing.

---

## 12. Build Notes
- PyInstaller build instructions are in Docs/BUILD_EXE.md (unchanged here).

---

## 13. Exit Codes
- 0 success, 1 invalid/runtime error, 3 not found/timeout, 4 multiple matches need --index, 130 Ctrl-C

---

## 14. Conclusion
v1.4.7.1 introduces fast query helpers, a robust FX subsystem (list/learn/toggle/delete, multi-write), safer COM lifecycle via thread-local contexts, improved Listen routing (ID and name, empty flag for default device), and a richer GUI with FX actions and background state caching.
