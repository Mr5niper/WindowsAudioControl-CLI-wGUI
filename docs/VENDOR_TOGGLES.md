# Vendor Toggles Configuration Guide (v1.4.7.1)

## What is vendor_toggles.ini?
A database that teaches audioctl how to toggle Enhancements (SysFX) for your device(s). It supports:
- Main Enhancements switch (on/off)
- Per-effect “FX” toggles (e.g., “BassBoost”, “Loudness”), including multi-write setups

Default location
- Next to audioctl.exe if writable; otherwise %LOCALAPPDATA%\audioctl\vendor_toggles.ini

Do I need this?
- Yes if Enhancements toggling fails with “No vendor toggle available.”
- You can teach it via the GUI or CLI (learn flows).

---

Quick Start (GUI)
1. Right-click your device → “Learn Enhancements”
2. Read the warning; click OK
3. Set “Audio Enhancements” to ENABLED in Windows → click OK to capture A
4. Set to DISABLED → click OK to capture B
5. The tool writes a vendor section into vendor_toggles.ini for this device (per-device membership is recorded)

Learn a specific effect (FX) in GUI
- Right-click → “Learn Enhancements” → choose “A specific effect” and specify an FX name (e.g., “BassBoost”)
- Follow the prompts (two-pass A/B) to stabilize driver behavior
- The INI gets an FX bucket with one or more write{i}_* blocks, scoped to your device

Quick Start (CLI)
- Learn main toggle (manual UI):
  ```
  audioctl enhancements --name "Speakers" --flow Render --learn
  ```
- Learn FX:
  ```
  audioctl enhancements --id "{ENDPOINT-ID}" --flow Render --learn-fx "BassBoost"
  ```

---

INI Formats

Main entry (type implied = main)
```
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
- subkey records the learned location (FxProperties or Properties) to read/write exactly where the driver uses it.
- devices is required; entries only apply to listed endpoint GUIDs.

FX entry (multi-write)
```
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
write1_devices = {guid-a},{guid-b}   ; optional: if missing, applies to all; empty => applies to none

write2_hive = HKLM
write2_subkey = Properties
write2_name = {other-guid},1
write2_type_enable = REG_BINARY
write2_type_disable = REG_BINARY
write2_enable = hex:01
write2_disable = hex:00
; write2_devices omitted => universal

write3_hive = HKCU
write3_subkey = FxProperties
write3_name = {x-guid},2
write3_type_enable = REG_SZ
write3_type_disable = REG_SZ
write3_enable = Enabled
write3_disable = Disabled
write3_devices =
```

FX entry (single DWORD — legacy/simple)
```
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

Devices field semantics
- devices (section-level): union list of endpoint GUIDs that this bucket applies to.
- write{i}_devices: per-toggle scoping:
  - missing => universal (applies to any device in the bucket)
  - empty => applies to nobody
  - list => applies only to those GUIDs

Decider & quorum
- decider_index: primary toggle for verification/fast reads (1-based)
- quorum_threshold (default 0.60): fraction of applicable toggles that must agree for a definite True/False

---

CLI Operations

List FX for a device
```
audioctl enhancements --id "{ENDPOINT-ID}" --list-fx
audioctl enhancements --id "{ENDPOINT-ID}" --list-fx --json
```

Learn an FX
```
audioctl enhancements --id "{ENDPOINT-ID}" --flow Render --learn-fx "BassBoost"
```

Enable/Disable an FX
```
audioctl enhancements --id "{ENDPOINT-ID}" --flow Capture --enable-fx "BassBoost"
audioctl enhancements --id "{ENDPOINT-ID}" --flow Capture --disable-fx "BassBoost"
```

Delete FX associations for this device
```
audioctl enhancements --id "{ENDPOINT-ID}" --delete-fx "BassBoost"
```

Main Enhancements learn (manual UI)
```
audioctl enhancements --name "Speakers" --flow Render --learn
```

---

Best Practices
- Learn main before FX if the driver couples FX to the global switch.
- Use admin if your driver needs HKLM writes (prefer-hklm for main toggles).
- Keep effect names human-friendly but consistent (e.g., “BassBoost”, not “BB”).

Troubleshooting
- “No vendor toggle available”: Learn first.
- Toggle applied but not reflected: Values or locations may be inverted or driver-specific; run discover-enhancements and inspect reports.
- INI permissions: If exe directory isn’t writable, the tool uses %LOCALAPPDATA%\audioctl\vendor_toggles.ini.

Discovery (advanced)
```
audioctl discover-enhancements --id "{ENDPOINT-ID}" --flow Render --output-dir "."
```
- Writes TXT and JSON bundle (with registry lists and diffs).
- Optionally append a suggested INI snippet:
  ```
  audioctl discover-enhancements --id "{ENDPOINT-ID}" --flow Render --ini-snippet vendor_snippets.ini
  ```
