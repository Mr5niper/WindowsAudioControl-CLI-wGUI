# Vendor Toggles Configuration Guide (v1.4.7.2)

This guide explains **vendor_toggles.ini**: the database that teaches **audioctl** how to reliably toggle **Audio Enhancements (SysFX)** and individual **FX effects** on specific Windows audio endpoints.

Because vendors implement enhancements/effects differently (and sometimes don’t honor Windows’ generic SysFX switches), audioctl uses **learned registry-write rules** stored in this INI.

---

## 1) What is vendor_toggles.ini?

`vendor_toggles.ini` is a **per-machine/per-user configuration database** that describes:

- **Main Enhancements switch** (“Audio Enhancements” on/off)
- **Per-effect FX toggles** (e.g., `BassBoost`, `Loudness`)
  - including complex drivers that require **multiple registry writes** (multi-write FX)

audioctl reads this file at runtime to decide:
- whether Enhancements/FX are supported for a device
- how to read the current state quickly (fast probes)
- what to write to toggle state

---

## 2) Default location and permissions

audioctl chooses the INI path using this priority:

1. **Next to `audioctl.exe`** (or the module directory) **if writable**
2. Otherwise:
   - `%LOCALAPPDATA%\audioctl\vendor_toggles.ini`

Notes:
- If you install under **Program Files**, the EXE directory may not be writable without elevation.
- The tool will automatically fall back to a user-writable location when needed.

---

## 3) Do I need this file?

You need `vendor_toggles.ini` if you see errors like:

- “No vendor toggle available for this device. Use --learn to teach a vendor method.”

If your driver/device already has a known learned entry, audioctl can toggle Enhancements/FX immediately.

---

## 4) Quick Start (GUI)

### 4.1 Learn the main Enhancements toggle
1. Right-click your device → **Learn Enhancements**
2. Read the warning and continue
3. In Windows Sound settings for that device:
   - Set **Audio Enhancements = Enabled**, then click OK in the GUI to capture snapshot A
   - Set **Audio Enhancements = Disabled**, then click OK to capture snapshot B
4. audioctl writes/updates a vendor entry in `vendor_toggles.ini` and associates the endpoint GUID with it.

### 4.2 Learn a specific FX effect
1. Right-click your device → **Learn Enhancements**
2. Choose **A specific effect (FX)** and enter an effect name (example: `BassBoost`)
3. Follow the prompts (two-pass A/B) to stabilize driver behavior
4. The INI gets (or updates) an FX “bucket” section containing one or more write blocks scoped to your endpoint.

---

## 5) Quick Start (CLI)

### 5.1 Learn main Enhancements (manual UI)
```bash
audioctl enhancements --name "Speakers" --flow Render --learn
```

### 5.2 Learn an FX
```bash
audioctl enhancements --id "{ENDPOINT-ID}" --flow Render --learn-fx "BassBoost"
```

---

## 6) Key concept: endpoint GUID vs endpoint ID

Windows exposes audio endpoints with a full **endpoint ID** string (IMMDevice ID), but the registry locations audioctl uses are keyed by the **endpoint GUID** inside it.

- **Endpoint ID** example:
  - `{0.0.0.00000000}.{83a9be54-901e-4429-993b-c9088e3028a0}`
- **Endpoint GUID** used in `devices = ...` lists:
  - `{83a9be54-901e-4429-993b-c9088e3028a0}`

`vendor_toggles.ini` uses the **endpoint GUID** values.

---

## 7) INI formats

`vendor_toggles.ini` contains sections. Each section is either:
- a **MAIN** Enhancements toggle (default; “type” is implied), or
- an **FX** toggle (`type = fx`)

### Important parsing rule
Only `type = fx` is special.  
Any section **without** `type = fx` is treated as a **MAIN** toggle section.

---

## 8) MAIN entry format (Enhancements on/off)

A MAIN entry describes a single DWORD-style registry value that flips between two values when Enhancements are enabled/disabled.

```ini
[vendor_{GUID},pid]
value_name = {GUID},pid
dword_enable = 0
dword_disable = 1
hives = HKCU,HKLM
flows = Render,Capture
subkey = FxProperties
notes = optional text
devices = {endpoint-guid-1},{endpoint-guid-2}
```

Field meanings:
- `value_name`
  - Registry value name in MMDevices format: `{fmtid},pid`
- `dword_enable` / `dword_disable`
  - Values written for enable/disable (typically 0/1)
- `hives`
  - Which registry hives to try writing/reading (`HKCU`, `HKLM`)
  - Ordering matters for preference; HKLM usually requires Admin for writes
- `flows`
  - Which endpoint flow(s) this applies to: `Render`, `Capture`, or both
- `subkey`
  - The learned location:
    - `FxProperties` or `Properties`
  - audioctl reads/writes exactly where the driver actually uses the value
- `devices`
  - Required. Comma-separated endpoint GUIDs that this entry applies to.

Runtime rule:
- MAIN toggling is vendor-only: if your endpoint GUID is not listed in a MAIN entry, Enhancements toggling will not work until learned.

---

## 9) FX entry format (legacy single DWORD)

Some FX effects can be toggled via a single DWORD flip (simple driver behavior).

```ini
[fx_ffff111122223333]
type = fx
fx_name = Loudness
value_name = {vendor-guid},10
dword_enable = 1
dword_disable = 0
hives = HKCU,HKLM
flows = Render,Capture
notes = Some devices
devices = {guid-a},{guid-b}
```

Field meanings:
- `type = fx`
  - Marks this as an FX entry
- `fx_name`
  - The human name used by CLI/GUI (`--enable-fx`, etc.)
- `devices`
  - Required. Endpoints that have this FX effect learned.

---

## 10) FX entry format (multi-write)

Many modern drivers require **multiple writes** to toggle a single effect. Those are represented as multi-write FX entries.

```ini
[fx_abcd1234ef567890]
type = fx
fx_name = BassBoost
device_name_pattern = Speakers (optional, informational)
multi_write = 1

write_count = 3
decider_index = 1
quorum_threshold = 0.60

write1_hive = HKCU
write1_subkey = FxProperties
write1_name = {vendor-guid},7
write1_type_enable = REG_DWORD
write1_type_disable = REG_DWORD
write1_enable = 1
write1_disable = 0
write1_devices = {guid-a},{guid-b}   ; optional: if missing, “universal” within this bucket

write2_hive = HKLM
write2_subkey = Properties
write2_name = {other-guid},1
write2_type_enable = REG_BINARY
write2_type_disable = REG_BINARY
write2_enable = hex:01
write2_disable = hex:00
; write2_devices omitted => universal within this bucket

write3_hive = HKCU
write3_subkey = FxProperties
write3_name = {x-guid},2
write3_type_enable = REG_SZ
write3_type_disable = REG_SZ
write3_enable = Enabled
write3_disable = Disabled
write3_devices =                ; empty => applies to nobody (explicitly disabled)

notes = optional text
devices = {guid-a},{guid-b}
```

### 10.1 The FX “bucket” concept (`devices = ...`)
For FX sections, `devices = ...` is the **bucket membership list**:
- It is required for the FX section to be considered applicable at all.
- It defines which endpoint GUIDs this FX section is associated with.

### 10.2 Per-write scoping (`write{i}_devices`)
Each write block can optionally define its own `write{i}_devices` list.

Semantics:
- **Missing** `write{i}_devices`:
  - The write is **universal within this bucket** (applies to any endpoint GUID listed in the section-level `devices`).
- **Empty** `write{i}_devices =`:
  - The write applies to **nobody** (effectively disabled for all endpoints).
- **List** `write{i}_devices = {guid1},{guid2}`:
  - The write applies **only** to those endpoint GUIDs.

This scoping exists because different devices can share an FX name but require slightly different registry payloads. audioctl can store multiple variants in one bucket and scope them correctly.

### 10.3 Types and payloads
Multi-write supports:
- `REG_DWORD`
  - payload is an integer (e.g., `0` or `1`)
- `REG_SZ`
  - payload is a string (e.g., `Enabled`)
- `REG_BINARY`
  - payload is typically stored as:
    - `hex:aa,bb,cc` (preferred for readability), or
    - raw hex without prefix (also accepted)

---

## 11) Decider & quorum (verification and state reads)

Multi-write FX entries use two mechanisms to decide the current state:

- `decider_index` (1-based):
  - The “primary” write block used first for deciding state.
- `quorum_threshold` (default 0.60):
  - Fraction of applicable writes that must agree to confidently report True/False.

How it’s used:
- During verification after applying a toggle (did the change “stick”?)
- During fast state reads for GUI menu labels

Practical guidance:
- Keep `decider_index` pointed at the most stable/obvious signal (often a DWORD under FxProperties).
- Keep quorum around the default unless you have a reason to raise/lower it.

---

## 12) CLI operations related to vendor_toggles.ini

### 12.1 List FX for a device
```bash
audioctl enhancements --id "{ENDPOINT-ID}" --list-fx
audioctl enhancements --id "{ENDPOINT-ID}" --list-fx --json
```

### 12.2 Learn an FX
```bash
audioctl enhancements --id "{ENDPOINT-ID}" --flow Render --learn-fx "BassBoost"
```

### 12.3 Enable/disable an FX
```bash
audioctl enhancements --id "{ENDPOINT-ID}" --flow Capture --enable-fx "BassBoost"
audioctl enhancements --id "{ENDPOINT-ID}" --flow Capture --disable-fx "BassBoost"
```

### 12.4 Delete FX associations for this device
```bash
audioctl enhancements --id "{ENDPOINT-ID}" --delete-fx "BassBoost"
```

What delete does (important):
- Removes this device GUID from the FX bucket (`devices = ...`).
- Removes this GUID from any `write{i}_devices` lines that list it.
- If a write block was previously universal (no `write{i}_devices` line), delete may materialize device scoping for remaining devices so the removed GUID is explicitly excluded.
- The section is generally not deleted; it is left in place for other devices that still use it.

### 12.5 Learn main Enhancements (manual UI)
```bash
audioctl enhancements --name "Speakers" --flow Render --learn
```

---

## 13) Best practices

- Learn the **main Enhancements** toggle before learning FX if your driver ties effects to the global switch.
- Run elevated when needed for HKLM writes:
  - Some vendor toggles only work when written to HKLM.
- Keep effect names consistent and human-friendly:
  - Example: `BassBoost`, not `BB`.
- Avoid manual edits unless you understand the consequences:
  - Wrong hives/subkeys/payloads can cause toggles to fail or report incorrect state.

---

## 14) Troubleshooting

### 14.1 “No vendor toggle available”
- Learn first:
  ```bash
  audioctl enhancements --id "{ENDPOINT-ID}" --flow Render --learn
  ```

### 14.2 Toggle applied but state doesn’t change
- Driver may store state in a different hive/subkey than expected.
- Use discovery to capture a detailed report:
  ```bash
  audioctl discover-enhancements --id "{ENDPOINT-ID}" --flow Render --output-dir "."
  ```

### 14.3 INI permissions issues
- If the EXE directory isn’t writable, audioctl will use `%LOCALAPPDATA%\audioctl\vendor_toggles.ini`.
- Ensure you’re editing the correct INI file.

---

## 15) Discovery (advanced)

Discovery produces a TXT report and JSON bundle that shows:
- what registry keys changed
- which values flipped
- which candidates look like the real vendor toggle

```bash
audioctl discover-enhancements --id "{ENDPOINT-ID}" --flow Render --output-dir "."
```

Optionally append a suggested INI snippet:
```bash
audioctl discover-enhancements --id "{ENDPOINT-ID}" --flow Render --ini-snippet vendor_snippets.ini
```
