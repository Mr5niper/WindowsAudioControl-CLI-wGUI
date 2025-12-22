# Windows Audio Control CLI + GUI
## `audioctl.exe`

Windows audio control CLI/GUI tool (pycaw-based);
List devices
set default endpoints
adjust volume/mute
toggle “Listen to this device”
toggle “Audio Enhancements” for devices that support vendor toggles.

---

## Quick Start

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

---

## List Devices

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

- Tip: grab an ID for later use:
  ```bash
  audioctl.exe list --json
  ```

---

## Set Default Device(s)

- Set playback default for **all roles** by name:
  ```bash
  audioctl.exe set-default --playback-name "Speakers (Realtek)" --playback-role all
  ```

- Set recording default for **communications** role by name:
  ```bash
  audioctl.exe set-default --recording-name "Headset Mic"
  ```

- Target by ID:
  ```bash
  audioctl.exe set-default --playback-id "{render-endpoint-id}" --playback-role multimedia
  ```

- When multiple name matches:
  ```bash
  audioctl.exe set-default --playback-name "Speakers" --index 0
  ```

- Note: setting defaults may require running in an elevated context (Run as Administrator).

---

## Set Volume or Mute/Unmute

- Set volume to 30% for a playback device:
  ```bash
  audioctl.exe set-volume --name "Speakers" --flow Render --level 30
  ```

- Mute a capture device by ID:
  ```bash
  audioctl.exe set-volume --id "{capture-endpoint-id}" --mute
  ```

- Unmute by name (add `--index` when multiple matches):
  ```bash
  audioctl.exe set-volume --name "USB Mic" --flow Capture --unmute
  ```

---

## “Listen to this device” (Capture Only)

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

## Audio Enhancements (SysFX) – Vendor Toggles

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

---

## Diagnose Enhancements / SysFx

- Dump live Enhancements status and vendor toggle status for a device:
  ```bash
  audioctl.exe diag-sysfx --name "Speakers" --flow Render
  ```

Sample JSON fields:

- `enhancementsEnabled_live_propstore` – Windows property store view (Disable_SysFx).
- `enhancementsEnabled_live_com` – PolicyConfig COM view.
- `vendor_toggle_status` – First matched vendor entry and its current state.

---

## Discover Enhancements Behavior (for debugging / advanced)

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

## Wait for a Device to Appear

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
