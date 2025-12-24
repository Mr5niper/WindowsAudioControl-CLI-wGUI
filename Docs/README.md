# Windows Audio Control CLI + GUI
## audioctl.exe (pycaw-based)
[List devices](#list-devices)
<BR>
[Set default endpoints](#set-default-devices-cli)
<BR>
[Adjust volume / mute](#set-volume-or-muteunmute-cli)
<BR>
[Listen to this device](#listen-to-this-device-capture-only)
<BR>
[Audio enhancements](#audio-enhancements-sysfx--vendor-toggles)
<BR>
[Wait for a device](#wait-for-a-device-to-appear)
<BR>

<BR>


# Quick Start

- Help:
  ```bash
  audioctl.exe -h
  ```

- No arguments (double‑click or run without args) launches the GUI:
  ```bash
  audioctl.exe
  ```

---

## Selectors (used by most commands)

- `--id "{endpoint-id}"` (exact endpoint ID)
- `--name "substring"` (case-insensitive substring match)
- `--regex` (treat `--name` as a regex)
- `--flow Render|Capture` (optional filter for playback/recording)
- `--index N` (0-based index to disambiguate when multiple matches)

When using the --name option, commands accept partial matches against device names or descriptions. If the specified text uniquely identifies exactly one playback or recording device, that device will be targeted and an index value is not required.

---

# List Devices

- Human-readable:
  ```bash
  audioctl.exe list
  ```

- Include disabled/disconnected:
  ```bash
  audioctl.exe list --all
  ```

- JSON output:
  ```bash
  audioctl.exe list --json
  ```

---

# Set Default Devices (CLI)

Use `audioctl set-default` to choose the system’s default playback (Render) and/or recording (Capture) endpoints. Targets must be active (`DEVICE_STATE_ACTIVE`). On some systems, you may need an elevated console (Run as Administrator).

Roles you can set:
- console
- multimedia
- communications
- all (applies to all three roles)

Defaults if a role is omitted:
- Playback role defaults to `all`.
- Recording role defaults to `communications`.

About Windows roles (console vs multimedia):
- On most Windows systems, “console” and “multimedia” effectively point to the same default device. If you set either one, you will typically see both flags update in the next `audioctl list` output. The tool still lets you set them independently—or set `all`—for completeness and for environments where they’re distinct.

Get IDs first (optional but recommended):
```bash
audioctl list --json
```

Disambiguation when multiple name matches:
- If your `--playback-name` or `--recording-name` matches more than one active device, the command will stop and print the matching candidates in GUI order with an index for each. Then rerun with `--index N` to choose the one you want.
- Example of the prompt:
  ```
  Multiple playback matches:
    [Render idx 0] Speakers (USB DAC)  id={...}  defaults=-
    [Render idx 1] Speakers (Realtek(R) Audio)  id={...}  defaults=multimedia,console
  Use --index to disambiguate.
  ```

Notes about JSON output:
- Successful commands print compact, single‑line JSON (no pretty‑printing). For pretty output, pipe to a formatter (e.g., `| jq .`).

---

## Examples

### 1) Set playback default by exact ID (multimedia role)
```bash
audioctl set-default --playback-id "{RENDER-ENDPOINT-ID}" --playback-role multimedia
```
Sample output (single line):
```json
{"set":[{"flow":"Render","role":"multimedia","id":"{RENDER-ENDPOINT-ID}","name":"Speakers (Realtek(R) Audio)"}]}
```

### 2) Set playback default by name (all roles); disambiguate by index
```bash
audioctl set-default --playback-name "Speakers" --playback-role all --index 0
```
Sample output:
```json
{"set":[{"flow":"Render","role":"all","id":"{RENDER-ENDPOINT-ID}","name":"Speakers (USB DAC)"}]}
```

### 3) Set recording default by name (default role = communications)
```bash
audioctl set-default --recording-name "USB Microphone"
```
Sample output:
```json
{"set":[{"flow":"Capture","role":"communications","id":"{CAPTURE-ENDPOINT-ID}","name":"USB Microphone"}]}
```

### 4) Set recording default by exact ID (communications role)
```bash
audioctl set-default --recording-id "{CAPTURE-ENDPOINT-ID}" --recording-role communications
```
Sample output:
```json
{"set":[{"flow":"Capture","role":"communications","id":"{CAPTURE-ENDPOINT-ID}","name":"Headset Mic"}]}
```

### 5) Set both playback and recording in one command
Playback to all roles; recording to communications:
```bash
audioctl set-default \
  --playback-id "{RENDER-ENDPOINT-ID}" --playback-role all \
  --recording-id "{CAPTURE-ENDPOINT-ID}" --recording-role communications
```
Sample output:
```json
{"set":[{"flow":"Render","role":"all","id":"{RENDER-ENDPOINT-ID}","name":"Speakers (Realtek(R) Audio)"},{"flow":"Capture","role":"communications","id":"{CAPTURE-ENDPOINT-ID}","name":"USB Microphone"}]}
```

### 6) Use regex name matching (with index if needed)
Targets the first Render device matching the regex; `--index` is the GUI‑order index among matches for that flow.
```bash
audioctl set-default --playback-name "Speakers.*Realtek" --regex --playback-role multimedia --index 0
```
Sample output:
```json
{"set":[{"flow":"Render","role":"multimedia","id":"{RENDER-ENDPOINT-ID}","name":"Speakers (Realtek(R) Audio)"}]}
```

### 7) Set a specific playback role (console) by name
(Useful when you want to explicitly set console, even though many systems treat console and multimedia the same.)
```bash
audioctl set-default --playback-name "Headphones" --playback-role console
```
Sample output:
```json
{"set":[{"flow":"Render","role":"console","id":"{RENDER-ENDPOINT-ID}","name":"Headphones"}]}
```

### 8) Set the playback communications role (e.g., a softphone/headset)
```bash
audioctl set-default --playback-name "USB Headset" --playback-role communications
```
Sample output:
```json
{"set":[{"flow":"Render","role":"communications","id":"{RENDER-ENDPOINT-ID}","name":"USB Headset"}]}
```

### 9) JSON result for automation (single line)
```bash
audioctl set-default --playback-id "{RENDER-ENDPOINT-ID}" --playback-role all
```
Output:
```json
{"set":[{"flow":"Render","role":"all","id":"{RENDER-ENDPOINT-ID}","name":"Speakers (Realtek(R) Audio)"}]}
```

Additional notes:
- If list output shows both console and multimedia toggled after setting only one, that is expected on most Windows builds—they commonly map to the same endpoint.
- “device not found (active only)”: verify the device is connected/enabled and visible in `audioctl list`.
- If multiple name matches occur, rerun with `--index N` using the indices printed in the disambiguation prompt.

---

# Set Volume or Mute/Unmute (CLI)

Use `audioctl set-volume` to change an endpoint’s master volume (0–100%) or toggle its mute state. Targets must be active (`DEVICE_STATE_ACTIVE`). On name selection, prefer adding `--flow` to disambiguate between Render (playback) and Capture (recording). If multiple matches remain, add `--index N` (0‑based, GUI‑order within the flow).

Rules:
- You must specify exactly one of: `--level`, `--mute`, or `--unmute`.
- `--level` is an integer 0–100 (values outside this range are clamped by the tool).
- Either `--id` or `--name` is required for selection (with optional `--regex` for pattern matching).

Command template:
```bash
audioctl set-volume
  (--id <ID> | --name <NAME>) [--flow <Render|Capture>]
  (--level <0..100> | --mute | --unmute)
  [--index <N>] [--regex]
```

Disambiguation behavior:
- If name selection matches more than one active device and you don’t pass `--index`, the tool returns:
  ```
  ERROR: multiple matches; specify --index
  ```
  Rerun the command with `--flow` and/or `--index N`. Use `audioctl list` (or `--json`) to see devices in the same GUI order.

Notes about JSON output:
- Successful commands print compact, single‑line JSON. Pipe to a formatter (e.g., `| jq .`) if you want pretty output.

---

## Examples

### 1) Set playback volume by name (Render, 30%)
```bash
audioctl set-volume --name "Speakers" --flow Render --level 30
```
Sample output:
```json
{"volumeSet":{"id":"{RENDER-ENDPOINT-ID}","name":"Speakers (Realtek(R) Audio)","level":30}}
```

### 2) Set recording volume by ID (Capture, 70%)
```bash
audioctl set-volume --id "{CAPTURE-ENDPOINT-ID}" --level 70
```
Sample output:
```json
{"volumeSet":{"id":"{CAPTURE-ENDPOINT-ID}","name":"USB Microphone","level":70}}
```

### 3) Mute a playback device by name (Render) with index disambiguation
If multiple “Speakers” exist, specify which one using `--index` (0‑based, GUI‑order for Render).
```bash
audioctl set-volume --name "Speakers" --flow Render --mute --index 0
```
Sample output:
```json
{"muteSet":{"id":"{RENDER-ENDPOINT-ID}","name":"Speakers (USB DAC)","muted":true}}
```

### 4) Unmute a recording device by ID (Capture)
```bash
audioctl set-volume --id "{CAPTURE-ENDPOINT-ID}" --unmute
```
Sample output:
```json
{"muteSet":{"id":"{CAPTURE-ENDPOINT-ID}","name":"USB Microphone","muted":false}}
```

### 5) Use regex name matching (Render)
```bash
audioctl set-volume --name "^Speakers .* Realtek\\(R\\) Audio$" --regex --flow Render --level 15
```
Sample output:
```json
{"volumeSet":{"id":"{RENDER-ENDPOINT-ID}","name":"Speakers (Realtek(R) Audio)","level":15}}
```

### 6) Name match without `--flow` and multiple results
If both Render and Capture have a device matching “USB”, you’ll see:
```
ERROR: multiple matches; specify --index
```
Disambiguate by flow and/or index:
```bash
audioctl set-volume --name "USB" --flow Capture --unmute --index 0
```
Sample output:
```json
{"muteSet":{"id":"{CAPTURE-ENDPOINT-ID}","name":"USB Microphone","muted":false}}
```

### 7) Error safeguard: `--level` cannot be combined with `--mute/--unmute`
Invalid example (will return an error):
```bash
audioctl set-volume --name "Speakers" --flow Render --level 20 --mute
```
Error:
```
ERROR: Cannot specify both --level and --mute/--unmute
```

---

Tips:
- Prefer `--id` for exact targeting (no disambiguation needed).
- If you get “device not found (active only)”, verify the endpoint is connected/enabled and visible in `audioctl list`.
- The command controls the endpoint’s master scalar level and mute state (not per‑app volumes).

---

# “Listen to this device” (Capture Only)

- Enable “Listen”:
  ```bash
  audioctl.exe listen --name "Microphone" --enable
  ```

- Disable “Listen”:
  ```bash
  audioctl.exe listen --name "Microphone" --disable
  ```

- Enable “Listen” to **Default Playback Device**:
  ```bash
  audioctl.exe listen --name "Microphone" --enable --playback-target-id ""
  ```

- Enable and route to a specific playback endpoint:
  ```bash
  audioctl.exe listen --name "Microphone" --enable --playback-target-id "{render-endpoint-id}"
  ```

- Disable/enable “Listen” by ID:
  ```bash
  audioctl.exe listen --id "{capture-endpoint-id}" --disable
  ```

- Notes:
  - Use `audioctl.exe list --json` to find endpoint IDs.
  - When multiple “Microphone” matches exist, add `--index N`.
  - When `--playback-target-id` is omitted, the current target is preserved; when enabling and no target exists, `""` (Default Playback Device) is used.
  - The tool verifies the change via COM and (if needed) the registry; JSON output may include `verifiedBy`.

---

# Audio Enhancements (SysFX) – Vendor Toggles

For devices where a vendor toggle has been configured or learned, you can control “Audio Enhancements” at the endpoint level.

### Enable / Disable Enhancements

- Enable enhancements:
  ```bash
  audioctl.exe enhancements --name "Speakers" --flow Render --enable
  ```

- Disable enhancements:
  ```bash
  audioctl.exe enhancements --name "Speakers" --flow Render --disable
  ```

- Target by ID:
  ```bash
  audioctl.exe enhancements --id "{endpoint-id}" --flow Render --disable
  ```

- If multiple matches:
  ```bash
  audioctl.exe enhancements --name "Speakers" --flow Render --enable --index 0
  ```

- If no vendor toggle is known for that endpoint, you’ll get:
  ```text
  ERROR: No vendor toggle available for this device. Use --learn to teach a vendor method.
  ```

### Learn a Vendor Toggle (Manual)

Use this when your driver controls “Enhancements” via its own registry DWORD and you want audioctl to use that DWORD in the future.

- Manual learn (you flip the Windows UI yourself):
  ```bash
  audioctl.exe enhancements --name "Speakers" --flow Render --learn
  ```

This will:

1. Ask you for a strong confirmation (since it writes `vendor_toggles.ini`).
2. Instruct you to:
   - Set Enhancements **Enabled** in Windows UI → snapshot A.
   - Set Enhancements **Disabled** in Windows UI → snapshot B.
3. Diff MMDevices registry and write an entry into `vendor_toggles.ini` (next to the EXE or in the working dir).
4. From then on, `enhancements --enable/--disable` for that endpoint will write that vendor DWORD.

### Diagnose Enhancements / SysFx

- Dump live Enhancements status and vendor toggle status for a device:
  ```bash
  audioctl.exe diag-sysfx --name "Speakers" --flow Render
  ```

Sample JSON fields:

- `enhancementsEnabled_live_propstore` – Windows property store view (Disable_SysFx).
- `enhancementsEnabled_live_com` – PolicyConfig COM view.
- `vendor_toggle_status` – First matched vendor entry and its current state.

### Discover Enhancements Behavior (for debugging / advanced)

This is a deeper discovery command for diagnosing how a driver maps Enhancements to registry / COM.

- Interactive discovery:
  ```bash
  audioctl.exe discover-enhancements --name "Speakers" --flow Render --output-dir "."
  ```

It will:

1. Ask you to set Enhancements **Enabled** → snapshot A.
2. Ask you to set Enhancements **Disabled** → snapshot B.
3. Diff MMDevices registry and write:
   - A detailed TXT report.
   - A JSON bundle with both snapshots and diffs.

Optional:

- Also write a suggested vendor INI snippet:
  ```bash
  audioctl.exe discover-enhancements --name "Speakers" --flow Render --output-dir "." --ini-snippet vendor_toggles_suggested.ini
  ```

---

# Wait for a Device to Appear

- Wait up to 60 seconds for a capture device by name:
  ```bash
  audioctl.exe wait --name "USB Microphone" --flow Capture --timeout 60
  ```

- Wait by ID:
  ```bash
  audioctl.exe wait --id "{endpoint-id}" --timeout 30
  ```

On success, prints JSON with the selected device; on timeout, exits with code 3.

---

## Selection Tips

- Prefer `--id` for exact targeting.
- With `--name`:
  - Only **active** devices are considered (for most commands).
  - When multiple matches exist, specify `--index N` (GUI-style index).
- `--flow` helps disambiguate when the same name appears for Render and Capture.

---

## JSON for Automation

Most commands print a JSON object on success, for scripting.

Examples:

- Volume set:
  ```bash
  audioctl.exe set-volume --name "Speakers" --flow Render --level 20
  ```
  Output:
  ```json
  {"volumeSet":{"id":"...","name":"Speakers ...","level":20}}
  ```

- Mute set:
  ```json
  {"muteSet":{"id":"...","name":"USB Mic","muted":true}}
  ```

- Listen set:
  ```json
  {"listenSet":{"id":"...","name":"Microphone","enabled":true,"verifiedBy":"com"}}
  ```

- Enhancements set:
  ```json
  {"enhancementsSet":{"id":"...","name":"Speakers","enabled":false,"verifiedBy":"vendor:realtek_waves_primary"}}
  ```

---

## Exit Codes

- `0` – success
- `1` – runtime failure or invalid option combo
- `3` – device not found or wait timeout
- `4` – multiple matches; provide `--index`
- `130` – Ctrl‑C interrupted

---

## Troubleshooting

- **Set-default fails**:
  - Try running in an elevated PowerShell / Command Prompt (Run as Administrator).

- **comtypes/automation errors in the EXE**:
  - The build already includes a compatibility shim at the top of `audioctl/compat.py`.
  - Rebuild with the provided PyInstaller command if you change Python or comtypes versions.

- **“Listen” seems unchanged**:
  - Use explicit default routing: `--playback-target-id ""` (Default Playback Device).
  - Check JSON output: look at `enabled` and `verifiedBy` (`"com"` or `"registry"`).

- **Enhancements don’t toggle**:
  - For that endpoint, there may be no vendor DWORD configured yet.
  - Use `diag-sysfx` and/or `discover-enhancements`, then `enhancements --learn`.

---

## GUI (Graphical User Interface)

- Launch the GUI:
  - Double-click `audioctl.exe` (no arguments), or run it without command-line options.


<img width="1367" height="332" alt="image" src="https://github.com/user-attachments/assets/de30018a-5995-4892-9b1c-1f99e58ce956" />


### Features

- Visual list of playback (Render) and recording (Capture) devices, grouped and sorted.
- “Show disabled/disconnected” toggle.
- Right-click a **device** for actions:
  - “Set as Default (all roles)”
  - “Set Volume…”
  - “Mute” / “Unmute”
  - “Toggle Listen (capture only)”
  - “Enable/Disable Enhancements” (when a vendor toggle is known)
  - “Learn Enhancements” (discover a vendor DWORD for that device)


    <img width="344" height="224" alt="image" src="https://github.com/user-attachments/assets/8745eae7-b9d8-415a-bb7a-3feebd739300" />


### Print CLI Commands

- Enable “Print CLI commands” in the top bar.  
  Every GUI action will print the equivalent `audioctl.exe` command to stdout (useful for learning the CLI and scripting).


<img width="418" height="71" alt="image" src="https://github.com/user-attachments/assets/f93e952f-a0a1-4a3a-ae9b-3610d0c2c966" />
    <br>
    <br>

![CLI Output](https://github.com/user-attachments/assets/1f52d28b-71fa-4343-abde-728b164d2a99)

### Refresh Devices

- Click “Refresh” or press `F5`.

### Notes

- Setting default devices via the GUI may require Administrator privileges; the GUI will warn when recommended.
- Group header rows (e.g., “Playback (Render)”) are not selectable and have no actions; right-click the actual device entries.

---

## Credits

This project uses:

- pycaw — <https://github.com/AndreMiras/pycaw>  
- comtypes — <https://github.com/enthought/comtypes>  

Windows GUIDs and PROPERTYKEY constants are Microsoft API identifiers and do not require attribution.

---

## Third‑Party License Texts

### pycaw — MIT License

```text
Copyright (c) Andre Miras and contributors

Permission is hereby granted, free of charge, to any person obtaining a copy
...
THE SOFTWARE IS PROVIDED “AS IS”, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
...
```

Source: <https://github.com/AndreMiras/pycaw>

---

### comtypes — MIT License

```text
Copyright (c) Thomas Heller, Enthought, Inc., and contributors

Permission is hereby granted, free of charge, to any person obtaining a copy
...
THE SOFTWARE IS PROVIDED “AS IS”, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
...
```

Source: <https://github.com/enthought/comtypes>

---

### PyInstaller License Notice (for binary distributions)

If you distribute a PyInstaller-built executable of this project, PyInstaller is licensed under the GNU General Public License v2 (GPL-2.0) with a special exception that permits using PyInstaller to build and distribute executables, regardless of your program’s license.

Full text: <https://github.com/pyinstaller/pyinstaller/blob/develop/COPYING.txt>  
Project page: <https://www.pyinstaller.org/>
