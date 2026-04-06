# Vendor Toggles & Learning Engine Reference
**Target:** `audioctl/vendor_db.py`  
**Database:** `vendor_toggles.ini`  

This document explains the internal architecture of the `audioctl` vendor toggle system. Because hardware vendors (Realtek, Waves, AMD, etc.) implement audio enhancements differently—often ignoring standard Windows APIs—`audioctl` relies on a learned database of registry operations to manipulate device states directly.

---

## 1. System Architecture: "Registry Truth"

The core philosophy of the `vendor_db.py` engine is **Registry Truth**. 

The system does *not* blindly trust the `devices = {GUID}` lists stored in the INI file to determine if a toggle applies to a device. Instead, the INI file acts as a **rule library**.

When evaluating whether an endpoint supports a specific toggle, the engine performs a live signature match:
1. It reads the target registry key defined in the rule (e.g., `FxProperties\{fmtid},pid`).
2. It checks if the current live registry value exactly matches either the defined `enable` or `disable` payload.
3. If the signature matches, the rule is considered applicable to that endpoint, regardless of whether the endpoint's GUID is explicitly listed in the INI.

This allows `audioctl` to seamlessly support newly plugged-in devices if they share a driver architecture with an already-learned device.

---

## 2. The Learning Engine

The learning subsystem (`_learn_fx_and_write_ini` and `_learn_vendor_from_discovery_and_write_ini`) generates these rules by observing registry changes during user interaction.

### 2.1 Stability Filtering (Noise Reduction)
When learning specific Enhancement Effects (FX), the Windows registry generates massive amounts of noise (timestamp updates, VU meter values, etc.). To isolate the actual toggle, the engine uses stability filtering:
- It takes the baseline snapshot (A).
- It takes a primary target snapshot (B), plus several rapid successive samples.
- `_stable_registry_map()` compares these samples. Any registry key whose type or value fluctuates between the quick samples is aggressively dropped.
- The diff is then performed strictly between the stable A map and the stable B map.

### 2.2 Two-Pass Learning
Many complex audio drivers lazily initialize registry keys—meaning they only create the necessary DWORDs/Binary blobs *after* the user toggles the UI for the very first time. 
To handle this, the engine supports a Two-Pass workflow:
- **Pass 1 (Prime):** The user toggles the effect ON and OFF. The engine captures this, forcing the driver to build its registry structures.
- **Pass 2 (Authoritative):** The user toggles the effect ON and OFF again. The engine uses this second pair of snapshots (`snapA2`, `snapB2`) as the absolute truth for building the INI rule.

### 2.3 Deduplication and Merging
When a new toggle is learned, the engine tries to avoid creating duplicate sections:
- If the exact same registry keys and payloads are found, the engine simply appends the new device's GUID to the existing rule's bucket.
- For human readability, it auto-generates `name_<hash>` and `guids_<hash>` metadata keys so engineers can easily see which endpoints share a specific rule.

---

## 3. MAIN Enhancements (SysFX)

MAIN entries define the master "Audio Enhancements" switch. These are universally modeled as a single `REG_DWORD` flip.

**Example MAIN Entry:**
```ini
[vendor_{1da5d803-d492-4edd-8c23-e0c0ffee7f0e},5]
value_name = {1da5d803-d492-4edd-8c23-e0c0ffee7f0e},5
dword_enable = 0
dword_disable = 1
hives = HKCU,HKLM
flows = Render,Capture
subkey = FxProperties
devices = {5f914d04-dfb5-4161-b879-38318c6a3725}
```

*   **`subkey`:** The engine records exactly where the flip occurred (`FxProperties` or `Properties`) and strictly limits runtime reads/writes to that location.
*   **`hives`:** Determines the write order. `HKCU` is preferred, but `HKLM` is supported (and usually requires Administrator privileges).

---

## 4. FX Effects (Multi-Write & Quorum Logic)

While some FX are simple DWORD flips, advanced drivers (like Waves MaxxAudio) require writing multiple distinct registry values of varying types (`REG_DWORD`, `REG_SZ`, `REG_BINARY`) simultaneously to toggle a single UI checkbox.

**Example Multi-Write FX Entry:**
```ini
[fx_a771ed7da0a3fd04]
type = fx
fx_name = Power Management
multi_write = 1
write_count = 2
quorum_threshold = 0.60

write1_hive = HKLM
write1_subkey = Properties
write1_name = {24dbb0fc-9311-4b3d-9cf0-18ff155639d4},2
write1_type_enable = REG_BINARY
write1_type_disable = REG_BINARY
write1_enable = hex:0b,00,00,00,01,00,00,00,ff,ff,00,00
write1_disable = hex:0b,00,00,00,01,00,00,00,00,00,00,00
write1_devices = {5f914d04-dfb5-4161-b879-38318c6a3725}

write2_hive = HKCU
...
```

### 4.1 State Readback (Quorum Voting)
Because multi-write FX involve several keys, determining if the effect is currently "Enabled" or "Disabled" requires consensus. The engine uses `_read_decider_state()`:
1. It reads every target value.
2. It counts how many values perfectly match the `enable` payloads vs. the `disable` payloads.
3. If `votes / total_writes >= quorum_threshold` (default 60%), the state is declared.
4. If a quorum cannot be reached (e.g., conflicting manual edits), it falls back to evaluating the single most reliable key (preferring `FxProperties` DWORDs).

### 4.2 Per-Write Scoping & Device Profiles (`write{i}_devices`)
During learning, different devices might map the same effect to slightly different payload combinations. The engine merges these into a single `fx_name` bucket, using `write{i}_devices` to scope specific writes to specific GUIDs.

By grouping writes that share the same GUIDs, the engine dynamically builds distinct **Device Profiles** within a single effect bucket.

*   **Missing:** This write is "Universal" and is injected into *every* device profile in the bucket.
*   **List (`{guid}`):** This write belongs strictly to the profile(s) for the listed GUIDs.
*   **Empty:** Applies to nobody (explicitly disabled by conflict cleanup).

### 4.3 Profile-Aware Universal Spoofing & Fallback
When a newly connected device (whose GUID is not yet explicitly learned in the INI) attempts to use an FX toggle, the engine uses **Profile-Aware Universal Spoofing**:

1. **Registry Scoring:** The engine scans the live Windows registry and compares existing keys against all known Device Profiles in the bucket.
2. **Profile Adoption:** It calculates a "best-fit" score. The engine adopts the profile with the highest match ratio and safely applies *only* the writes belonging to that specific profile. 
3. **Universal Fallback:** If the device's registry is completely unrecognized (it matches no specific profile), the engine safely falls back to evaluating all writes collectively. This ensures backwards compatibility with older, legacy INI rules.

---

## 5. Safe Deletion

Because FX rules are bucketed and shared across multiple devices, the CLI command `--delete-fx` does **not** blindly delete the INI section. 

Instead, `_delete_fx_for_guid()` performs a safe unlinking:
1. It removes the target GUID from the master `devices =` list.
2. It scans every `write{i}_devices` scope and removes the GUID.
3. If a write was previously "Universal" (missing scope), the engine materializes the scope to explicitly include the remaining devices, effectively isolating and removing the target GUID without breaking the effect for other endpoints.

---

## 6. INI File Lifecycle

The database is parsed via Python's `configparser` and cached in memory using a `(path, mtime)` tuple. This allows the GUI to rapidly poll device states using `_fast_read_vendor_entry_state()` without triggering disk I/O on every tick, while instantly reloading if a manual edit or Learn command modifies the file.
