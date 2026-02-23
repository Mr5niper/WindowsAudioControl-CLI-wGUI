# Vendor Toggles Configuration Guide (v1.5.1.1)
This guide explains **vendor_toggles.ini**: the database that teaches **audioctl** how to reliably toggle **Audio Enhancements (SysFX)** and individual **FX effects** on specific Windows audio endpoints.

Because vendors implement enhancements/effects differently (and sometimes don’t honor Windows’ generic SysFX switches), audioctl uses **learned registry-write rules** stored in this INI.

This document reflects the behavior of `audioctl/vendor_db.py` in **v1.5.1.1**.
--------------------------------------------------------------------------------

## 1) What is vendor_toggles.ini? (v1.5.1.1)

`vendor_toggles.ini` is audioctl’s **INI-backed vendor toggle rule database**. It stores the learned rules that tell audioctl **exactly which MMDevices registry values to read and write** to control:

- **MAIN “Audio Enhancements” (SysFX) on/off** for an endpoint
- **Per-effect FX toggles** (e.g., `BassBoost`, `Loudness`)
  - including drivers that require **multiple registry writes of different types** to toggle one effect (multi-write FX)

What’s in the file:
- For MAIN: one learned DWORD flip (enable vs disable) at a specific learned location (`FxProperties` or `Properties`)
- For FX:
  - legacy single-DWORD effects, or
  - multi-write effects (several values, possibly mixing `REG_DWORD`, `REG_SZ`, `REG_BINARY`)

### v1.5.1.1 nuance: shared rules + “registry truth” discovery
In v1.5.1.1, the INI is best thought of as a **rule library**, not just “a list of devices”.

audioctl uses **live registry signature matching (“registry truth”)** to decide applicability:

- **MAIN Enhancements**
  - At runtime, audioctl selects the MAIN rule whose `value_name` exists on the endpoint (at the learned `subkey`) and whose current value equals either the learned enable or disable payload.
  - The `devices = ...` list is still stored/maintained (and may be used for candidate filtering/bookkeeping), but **the final decision is made by reading the registry**.

- **FX effects**
  - FX sections are stored as **buckets** (shared between devices).
  - audioctl will list/apply FX when the endpoint is explicitly in the bucket’s `devices = ...`, **and it can also discover FX by signature** (if the endpoint’s registry matches the learned enable/disable signatures) even if the GUID isn’t listed yet.

So the file describes:
- how to toggle (what to write, where, and in which hive(s))
- how to read/verify state (single value or quorum/decider logic)
- how to identify applicability (signature match against the endpoint’s live registry)

--------------------------------------------------------------------------------

## 2) Default location and permissions (v1.5.1.1)
**`vendor_toggles.ini` ships inside the project’s `dist/` folder**, so when you build/package the app, the INI ends up **next to the generated `audioctl.exe`**.

So the effective default is:
- **`<exe folder>\vendor_toggles.ini`**

### 2.2 Runtime path selection rule (what the code does)
`vendor_db._vendor_ini_default_path()` still follows this rule:

1. Use `vendor_toggles.ini` **next to the EXE/module** **if that directory is writable**
2. Otherwise fall back to:
   - `%LOCALAPPDATA%\audioctl\vendor_toggles.ini`

This means:
- If the EXE is launched from a writable folder (typical for a `dist/` folder you own), audioctl will use the INI **next to the EXE**.
- If the EXE is installed into a non-writable location (commonly `C:\Program Files\...`), audioctl will automatically switch to the user-writable copy under **LocalAppData**.

### 2.3 Permissions / elevation implications
- **HKCU writes** (current user) usually work without elevation.
- **HKLM writes** may require running audioctl **as Administrator**.
- Even if you run as Administrator, the INI path selection is still based on **directory writability**; if the EXE folder is writable, it will keep using the INI next to the EXE.

### 2.4 Practical guidance
- If you expect one global shared INI (for all users), placing it next to the EXE works best when the EXE folder is writable by all intended users.
- If you install under Program Files and want to keep the INI next to the EXE anyway, you must ensure that directory is writable (not recommended), otherwise audioctl will fall back to `%LOCALAPPDATA%` automatically.

--------------------------------------------------------------------------------

## 3) Do I need this file?
Yes, if you want audioctl to toggle Enhancements or FX reliably.

In v1.5.1.1 the runtime policy is **vendor-only**:
- If audioctl cannot find a matching vendor rule in `vendor_toggles.ini`, it returns:
  - `no-vendor-method`

Common symptom:
- “No vendor toggle available for this device. Use --learn to teach a vendor method.”

--------------------------------------------------------------------------------

## 4) Key concept: endpoint GUID vs endpoint ID
Windows exposes audio endpoints with a full **endpoint ID** string (IMMDevice ID), but the MMDevices registry paths audioctl uses are keyed by the **endpoint GUID** inside it.

- Endpoint ID example:
  - `{0.0.0.00000000}.{83a9be54-901e-4429-993b-c9088e3028a0}`
- Endpoint GUID used in INI lists:
  - `{83a9be54-901e-4429-993b-c9088e3028a0}`

`vendor_toggles.ini` uses the **endpoint GUID** values in `devices = ...` and in per-write device scoping lists.

--------------------------------------------------------------------------------

## 5) Registry location model (what the INI “targets”)
For a given endpoint GUID, audioctl targets MMDevices keys:

`HKCU/HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\{Render|Capture}\{GUID}\{FxProperties|Properties}`

Important:
- v1.5.1.1 records and uses the learned **subkey**:
  - `FxProperties` or `Properties`
- audioctl tries hard not to “guess” at runtime; it reads/writes exactly the learned location.

--------------------------------------------------------------------------------

## 6) New / notable behavior in v1.5.1.1 (what changed vs older guides)

### 6.1 Vendor-only runtime (“no Windows fallback”)
`_apply_enhancements()` is vendor-only:
- It selects a MAIN entry from the INI using **registry signature truth** (see below)
- If none match, it fails; it does not fall back to generic Windows toggles

### 6.2 Registry-truth selection for MAIN Enhancements
For MAIN toggling and support checks:
- audioctl chooses the applicable MAIN entry by reading the live registry and validating:
  - the `value_name` exists at the learned subkey for the endpoint, and
  - the current value equals either `enable` or `disable` payload
- This selection is done by `_find_first_vendor_entry()` using `_main_entry_signature_applies()`

Practical result:
- Even if the INI contains many MAIN entries, only the one whose signature matches the endpoint “wins”.

### 6.3 FX discovery uses “signature truth” (universal listing)
For FX listing (`--list-fx`) and FX matching:
- audioctl first includes FX buckets where the GUID is explicitly listed in the section’s `devices=...`
- then it also includes FX entries whose registry signature matches *this endpoint now*:
  - multi-write FX: `_fx_signature_matches_multi()`
  - legacy single-DWORD FX: `_legacy_value_matches_this_guid_now()`

This means FX can show up even if your GUID is not explicitly listed—if the registry signature matches.

### 6.4 Multi-write “quorum” verification (unchanged concept, emphasized by v1.5 series)
Multi-write FX state is not determined by a single value:
- it uses quorum logic (`quorum_threshold`, default 0.60) in `_read_decider_state()`
- if quorum can’t decide, it falls back to a “best signal” write (prefers FxProperties + DWORD 0/1)

### 6.5 Two-pass FX learning supported (authoritative second pass)
`_learn_fx_and_write_ini(..., snapA2, snapB2)` supports:
- Pass 1: prime/init driver (some vendors create keys only after first toggle)
- Pass 2 (snapA2/snapB2): authoritative capture for learned writes

The code uses the second pair if provided:
- `useA = snapA2 if provided else snapA`
- `useB = snapB2 if provided else snapB`

### 6.6 Device-name -> GUID “buckets” for INI readability (v1.5 series behavior)
When learning/merging, audioctl may write extra keys into the same section:
- `name_<8hex> = <friendly device name>`
- `guids_<8hex> = {guid1},{guid2}`

This is informational / dedupe-friendly metadata:
- it is NOT required for runtime toggling
- it helps humans see that multiple GUIDs correspond to the “same named device” across re-installs

--------------------------------------------------------------------------------

## 7) INI formats and parsing rules

### 7.1 Section types
Each INI section is either:
- MAIN Enhancements toggle (default): section does NOT need `type = main`
- FX toggle: must include `type = fx`

Only `type = fx` is treated specially.
Any section without `type = fx` is treated as a MAIN entry candidate.

### 7.2 Common fields used across entries
- `notes` is optional (kept for humans; ignored by logic)
- `flows` is usually `Render,Capture`
- `hives` ordering matters for write attempts and read preference

--------------------------------------------------------------------------------

## 8) MAIN entry format (Enhancements on/off)
A MAIN entry describes a single DWORD-style registry value that flips between two values when Enhancements are enabled/disabled.

Example:
```ini
[vendor_{fmtid},pid]
value_name = {fmtid},pid
dword_enable = 0
dword_disable = 1
hives = HKCU,HKLM
flows = Render,Capture
subkey = FxProperties
notes = optional text
devices = {endpoint-guid-1},{endpoint-guid-2}

; Optional readability metadata (may be auto-added by learning)
name_1a2b3c4d = Speakers (Realtek(R) Audio)
guids_1a2b3c4d = {endpoint-guid-1},{endpoint-guid-2}
```

Field meanings:
- `value_name`
  - Registry value name in MMDevices format: `{fmtid},pid`
  - Stored lowercased internally
- `dword_enable` / `dword_disable`
  - Values written for enable/disable (typically 0/1)
- `hives`
  - Which registry hives to attempt (`HKCU`, `HKLM`)
  - HKLM usually requires Administrator
  - Ordering can influence preference
- `flows`
  - Which endpoint flows this entry can apply to: `Render`, `Capture`, or both
- `subkey`
  - Learned location: `FxProperties` or `Properties`
  - v1.5.1.1 reads/writes exactly this learned scope
- `devices`
  - A list of endpoint GUIDs associated with this entry (used for candidate filtering in some paths)
  - NOTE: v1.5.1.1 also uses registry signature truth; support/apply ultimately depends on signature match for the endpoint

Runtime rule (MAIN):
- Support is determined by whether a MAIN entry’s **registry signature matches** the endpoint now.
- Apply uses that matching entry and writes the configured DWORD to the learned subkey in the configured hives.

--------------------------------------------------------------------------------

## 9) FX entry format (legacy single DWORD)
Some FX effects can be toggled via a single DWORD flip.

Example:
```ini
[fx_ffff111122223333]
type = fx
fx_name = Loudness
device_name_pattern = Speakers   ; optional, informational / used for spoof matching
value_name = {fmtid},10
dword_enable = 1
dword_disable = 0
hives = HKCU,HKLM
flows = Render,Capture
notes = Some devices
devices = {guid-a},{guid-b}
```

Key points (v1.5.1.1):
- Legacy FX toggling is treated similarly to MAIN toggling for writing:
  - it writes a single DWORD value
- Before applying a legacy FX, audioctl performs a *truth check*:
  - `_legacy_value_matches_this_guid_now()` must be true
  - it then finds the live subkey (`FxProperties` vs `Properties`) via `_legacy_find_live_subkey()`
  - this prevents writing unrelated values on devices where the key doesn’t exist / doesn’t match expected payloads

Discovery/listing:
- A legacy FX can be listed for a device if:
  - the GUID is in the FX section’s `devices`, OR
  - the live registry value exists for this endpoint and matches enable/disable (signature match)

--------------------------------------------------------------------------------

## 10) FX entry format (multi-write)
Many drivers require multiple registry writes (including different types like BINARY) to toggle a single effect.

Example:
```ini
[fx_abcd1234ef567890]
type = fx
fx_name = BassBoost
device_name_pattern = Speakers (optional)
multi_write = 1
write_count = 3
decider_index = 1
quorum_threshold = 0.60

write1_hive = HKCU
write1_subkey = FxProperties
write1_name = {fmtid},7
write1_type_enable = REG_DWORD
write1_type_disable = REG_DWORD
write1_enable = 1
write1_disable = 0
write1_devices = {guid-a},{guid-b}   ; optional scoping

write2_hive = HKLM
write2_subkey = Properties
write2_name = {other-fmtid},1
write2_type_enable = REG_BINARY
write2_type_disable = REG_BINARY
write2_enable = hex:01
write2_disable = hex:00
; write2_devices omitted => universal (see semantics below)

write3_hive = HKCU
write3_subkey = FxProperties
write3_name = {x-fmtid},2
write3_type_enable = REG_SZ
write3_type_disable = REG_SZ
write3_enable = Enabled
write3_disable = Disabled
write3_devices =         ; empty => applies to nobody (explicitly disabled)

notes = optional text
devices = {guid-a},{guid-b}

; Optional readability metadata (may be auto-added by learning)
name_1a2b3c4d = Speakers (Realtek(R) Audio)
guids_1a2b3c4d = {guid-a},{guid-b}
```

### 10.1 Section-level `devices = ...` (FX bucket membership)
For FX sections, `devices = ...` is the “bucket membership” union list:
- used as a fast association list
- however, v1.5.1.1 can also discover FX by signature even if the GUID is not listed

### 10.2 Per-write scoping: `write{i}_devices` semantics (important)
Each write block can optionally define its own device list.

Semantics in v1.5.1.1 loader:
- Missing `write{i}_devices` key:
  - `devices: None` in memory
  - Means “universal within this FX bucket”
- Present but empty: `write{i}_devices =`
  - `devices: []` in memory
  - Means “applies to nobody” (explicitly disabled)
- Present with list:
  - applies only to those GUIDs

Important runtime nuance in v1.5.1.1:
- For *matching/verification/reading* multi-write FX signatures, code intentionally ignores per-write gating:
  - it checks whether the registry values match the INI signatures for this endpoint
- For *applying* multi-write writes, it also intentionally ignores per-write gating:
  - but it will only write a value if that value exists in HKCU or HKLM for this endpoint
  - (it skips “inventing” new keys)

Per-write device lists mainly matter for:
- merging learned variants into one FX bucket
- delete behavior (removing a device from a bucket and scoping writes to exclude it)

### 10.3 Types and payload encoding (multi-write)
Supported write types in the INI:
- `REG_DWORD`: payload is an integer text (e.g., `0`, `1`)
- `REG_SZ`: payload is text (e.g., `Enabled`)
- `REG_BINARY`: payload stored as:
  - preferred: `hex:aa,bb,cc` (human-readable, diffable, lossless)
  - also accepted: raw hex `aabbcc`

--------------------------------------------------------------------------------

## 11) Decider & quorum (verification and state reads for multi-write FX)
Multi-write FX uses `_read_decider_state()`:

1) Quorum vote:
- Reads each write’s value (recorded hive first, then alternate hive)
- Only values that exist and match either enable or disable count as votes
- If `votes_true / votes_total >= quorum_threshold` (and the opposite doesn’t also meet threshold) => True
- If `votes_false / votes_total >= quorum_threshold` => False

2) Fallback “best signal”:
- If quorum can’t decide, it selects a “best” write to read:
  - prefers `FxProperties`
  - prefers DWORD 0/1 flips
- Reads recorded hive then alternate hive

Config fields:
- `decider_index`
  - Stored/parsed, but the v1.5.1.1 readback logic primarily uses quorum + best-signal scoring.
- `quorum_threshold`
  - Default: 0.60
  - Clamped internally to [0.50, 0.95]

--------------------------------------------------------------------------------

## 12) Fast state reads (GUI polling) vs “full” reads
There are two read paths:

### 12.1 “Full” reads (verification / truth)
- `_read_vendor_entry_state()`:
  - MAIN and legacy FX: reads the learned subkey; prefers HKCU then HKLM (per entry hives)
  - multi-write FX: uses `_read_decider_state()` quorum logic

### 12.2 FAST reads (no COM, minimal probes)
- `_fast_read_vendor_entry_state()`:
  - MAIN / legacy FX: reads HKCU and HKLM and tie-breaks disagreements by key last-write time
  - multi-write FX: picks a “best indicator write” and probes recorded hive then alternate; tie-break by last-write time

Used by:
- GUI labels/polling where speed matters

--------------------------------------------------------------------------------

## 13) Learn flows and what they write to the INI

### 13.1 Learn MAIN Enhancements (manual UI)
Manual learn (“discovery flow”) captures:
- Snapshot A: user sets Enhancements Enabled
- Snapshot B: user sets Enhancements Disabled
Then:
- diffs registry and looks for a candidate DWORD flip
- records the subkey where the flip occurred (`FxProperties` or `Properties`)
- appends a MAIN section (or dedupes by identical payload and appends GUID to existing section)
- also may write name/guid bucket metadata

### 13.2 Learn MAIN Enhancements (auto)
Auto learn tries to toggle programmatically (propstore + registry), captures A/B, and writes similarly.

### 13.3 Learn FX (two-pass capable)
FX learn builds multi-write sets using stability filtering:
- Builds “stable maps” so noisy registry values are excluded
- Writes a multi-write FX bucket if any writes were found
- Otherwise falls back to a single DWORD flip learning

Merge rules when learning into an existing FX bucket:
- If an identical write identity+payload already exists, it appends this GUID to that write’s `write{i}_devices`
- If payload is new, it appends a new write block scoped to this GUID
- It then removes this GUID from any conflicting write blocks that share identity but differ in payload (conflict cleanup)

--------------------------------------------------------------------------------

## 14) Delete FX associations for a device (bucket-safe delete)
`--delete-fx "<name>"` (implemented by `_delete_fx_for_guid`) does NOT usually delete the whole FX section.

It:
- removes this GUID from the section-level `devices = ...`
- for each `write{i}`:
  - if `write{i}_devices` exists, removes the GUID from that list
  - if `write{i}_devices` is missing (universal), it materializes scoping to the remaining bucket devices to exclude the removed GUID
  - if nothing remains, it may set `write{i}_devices =` (empty => applies to nobody)

Rationale:
- FX buckets can be shared across multiple devices; deleting the whole bucket would break other devices.

--------------------------------------------------------------------------------

## 15) CLI operations related to vendor_toggles.ini

### 15.1 Learn main Enhancements (manual UI)
```bash
audioctl enhancements --name "Speakers" --flow Render --learn
```

### 15.2 Learn an FX
```bash
audioctl enhancements --id "{ENDPOINT-ID}" --flow Render --learn-fx "BassBoost"
```

### 15.3 List FX for a device
```bash
audioctl enhancements --id "{ENDPOINT-ID}" --list-fx
audioctl enhancements --id "{ENDPOINT-ID}" --list-fx --json
```

### 15.4 Enable/disable an FX
```bash
audioctl enhancements --id "{ENDPOINT-ID}" --flow Capture --enable-fx "BassBoost"
audioctl enhancements --id "{ENDPOINT-ID}" --flow Capture --disable-fx "BassBoost"
```

### 15.5 Delete FX associations for this device
```bash
audioctl enhancements --id "{ENDPOINT-ID}" --delete-fx "BassBoost"
```

--------------------------------------------------------------------------------

## 16) Best practices (v1.5.1.1 behavior-aware)
- Learn MAIN Enhancements first if your vendor effects depend on the global switch.
- Run elevated if HKLM writes are required (common for some vendor drivers).
- Keep effect names stable and human-friendly (`BassBoost`, `Loudness`, etc.).
- Prefer not to hand-edit multi-write payloads unless you understand the driver behavior:
  - incorrect types/payloads can cause mismatched signatures, failed writes, or unstable state reads.
- If a driver uses `Properties` instead of `FxProperties`, ensure `subkey = Properties` is recorded; v1.5.1.1 will respect it.

--------------------------------------------------------------------------------

## 17) Troubleshooting

### 17.1 “no-vendor-method” / “No vendor toggle available”
- You must learn a method:
  ```bash
  audioctl enhancements --id "{ENDPOINT-ID}" --flow Render --learn
  ```

### 17.2 Toggle writes succeed but state reads as None / wrong
- Possible causes:
  - driver writes to a different subkey (`Properties` vs `FxProperties`)
  - driver uses multi-write and quorum can’t decide due to missing values
  - HKCU/HKLM disagreement (fast reads tie-break by last-write time)
- Use discovery to inspect diffs:
  ```bash
  audioctl discover-enhancements --id "{ENDPOINT-ID}" --flow Render --output-dir "."
  ```

### 17.3 INI not updating / permission errors
- If EXE directory is not writable, the real INI is likely:
  - `%LOCALAPPDATA%\audioctl\vendor_toggles.ini`
- Verify which file is being used before editing.

--------------------------------------------------------------------------------

## 18) Discovery (advanced)
Discovery produces reports (TXT/JSON) showing:
- what registry keys changed
- which values flipped
- candidate toggles

```bash
audioctl discover-enhancements --id "{ENDPOINT-ID}" --flow Render --output-dir "."
```

Optionally append a suggested INI snippet:
```bash
audioctl discover-enhancements --id "{ENDPOINT-ID}" --flow Render --ini-snippet vendor_snippets.ini
```
```
