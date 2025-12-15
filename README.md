# Windows Audio Control CLI wGUI
## Audioctl.exe

Windows audio control CLI/GUI tool (pycaw-based). List devices, set default endpoints, adjust volume/mute, and toggle “Listen to this device.”

---

## Quick Start

- Help:
  ```
  dist\audioctl.exe -h
  ```

### Selectors (used by most commands)
- `--id "{endpoint-id}"` (exact)
- `--name "substring"` (case-insensitive)
- `--regex` (treat `--name` as regex)
- `--flow Render|Capture` (when needed)
- `--index N` (0-based, used when multiple matches)

---

## List Devices

- Human-readable:
  ```
  dist\audioctl.exe list
  ```

- Include disabled/disconnected:
  ```
  dist\audioctl.exe list --all
  ```

- JSON output:
  ```
  dist\audioctl.exe list --json
  ```

- Tip: grab an ID for later use:
  ```
  dist\audioctl.exe list --json
  ```

---

## Set Default Device(s)

- Set playback default for all roles by name:
  ```
  dist\audioctl.exe set-default --playback-name "Speakers (Realtek)" --playback-role all
  ```

- Set recording default for communications role by name:
  ```
  dist\audioctl.exe set-default --recording-name "Headset Mic"
  ```

- Target by ID:
  ```
  dist\audioctl.exe set-default --playback-id "{render-endpoint-id}" --playback-role multimedia
  ```

- When multiple name matches:
  ```
  dist\audioctl.exe set-default --playback-name "Speakers" --index 0
  ```

- Note: setting defaults may require running an elevated PowerShell (Run as Administrator).

---

## Set Volume or Mute/Unmute

- Set volume to 30% for a playback device:
  ```
  dist\audioctl.exe set-volume --name "Speakers" --flow Render --level 30
  ```

- Mute a capture device by ID:
  ```
  dist\audioctl.exe set-volume --id "{capture-endpoint-id}" --mute
  ```

- Unmute by name (add `--index` when multiple matches):
  ```
  dist\audioctl.exe set-volume --name "USB Mic" --flow Capture --unmute
  ```

---

## “Listen to this device” (Capture Only)

- Enable “Listen” to Default Playback Device:
  ```
  dist\audioctl.exe listen --name "Microphone" --enable --playback-target-id ""
  ```

- Enable and route to a specific playback endpoint:
  ```
  dist\audioctl.exe listen --name "Microphone" --enable --playback-target-id "{render-endpoint-id}"
  ```

- Disable “Listen” by ID:
  ```
  dist\audioctl.exe listen --id "{capture-endpoint-id}" --disable
  ```

- Notes:
  - Use `dist\audioctl.exe list --json` to find endpoint IDs.
  - When multiple “Microphone” matches exist, add `--index N`.
  - When `--playback-target-id` is omitted, the current target is preserved; when enabling and no target exists, `""` (Default Playback Device) is used.

---

## Wait for a Device to Appear

- Wait up to 60 seconds for a capture device by name:
  ```
  dist\audioctl.exe wait --name "USB Microphone" --flow Capture --timeout 60
  ```

- Wait by ID:
  ```
  dist\audioctl.exe wait --id "{endpoint-id}" --timeout 30
  ```

---

## Selection Tips

- Prefer `--id` for exact targeting.
- With `--name`:
  - Active devices are searched first; if none match, all devices are considered.
  - When multiple matches exist, specify `--index N`.

---

## JSON for Automation

- On success, commands print JSON (e.g., `listenSet`, `volumeSet`, `muteSet`, `set`).

- Example:
  ```
  dist\audioctl.exe set-volume --name "Speakers" --flow Render --level 20
  ```
  Output:
  ```
  {"volumeSet":{"id":"...","name":"Speakers ...","level":20}}
  ```

---

## Exit Codes

- 0: success
- 1: runtime failure or invalid option combo
- 3: device not found or wait timeout
- 4: multiple matches; provide `--index`
- 130: Ctrl-C interrupted

---

## Troubleshooting

- Set-default fails:
  - Run PowerShell as Administrator.

- comtypes/automation errors in the EXE:
  - Ensure the comtypes shim is at the very top of `audioctl.py` and rebuild.

- “Listen” seems unchanged:
  - Use explicit default routing: `--playback-target-id ""`
  - Changes are verified via COM and registry, and `verifiedBy` is included in JSON.

---

## GUI (Graphical User Interface)

- Launch the GUI:
  - Double-click `dist\audioctl.exe` (no arguments), or run it without command-line options.

<img width="1367" height="332" alt="image" src="https://github.com/user-attachments/assets/bc739e2f-e488-429d-ad8d-3aeabf6a468a" />


- Features:
  - Visual list of playback and recording devices, grouped and sorted.
  - Toggle “Show disabled/disconnected.”
  - Right-click a device for actions:
    - “Set as Default (all roles)”
    - “Set Volume…”
    - “Mute” / “Unmute”
    - “Toggle Listen (capture only)”


      <img width="328" height="169" alt="image" src="https://github.com/user-attachments/assets/3488adbf-b5cc-475c-9efd-0bf2e47acfa5" />



- Print CLI Commands:
  - Enable “Print CLI commands” in the top bar. Actions performed in the GUI will print the equivalent `audioctl.exe` command to stdout (useful for learning and scripting).

    <img width="419" height="71" alt="image" src="https://github.com/user-attachments/assets/c1cd3a37-2da9-42a5-9020-a65de7f7a4f3" />
    <br>
    <br>
    <img width="1006" height="191" alt="image" src="https://github.com/user-attachments/assets/1f52d28b-71fa-4343-abde-728b164d2a99" />



- Refresh Devices:
  - Click “Refresh” or press `F5`.

- Notes:
  - Setting default devices via the GUI may require Administrator privileges; the GUI will prompt when recommended.
  - Group header rows (e.g., “Playback (Render)”) cannot be selected or acted upon; right-click device entries.

---

## Credits

This project uses:

- pycaw — https://github.com/AndreMiras/pycaw
- comtypes — https://github.com/enthought/comtypes

Windows GUIDs and PROPERTYKEY constants are Microsoft API identifiers and do not require attribution.

---

## Third‑Party License Texts

### pycaw — MIT License

Copyright (c) Andre Miras and contributors

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the “Software”), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in
all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED “AS IS”, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
THE SOFTWARE.

Source: https://github.com/AndreMiras/pycaw

---

### comtypes — MIT License

Copyright (c) Thomas Heller, Enthought, Inc., and contributors

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the “Software”), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in
all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED “AS IS”, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
THE SOFTWARE.

Source: https://github.com/enthought/comtypes

---

### PyInstaller License Notice (for binary distributions)

If you distribute a PyInstaller-built executable of this project, PyInstaller is licensed under the GNU General Public License v2 (GPL-2.0) with a special exception that permits using PyInstaller to build and distribute executables, regardless of your program’s license.

Full text: https://github.com/pyinstaller/pyinstaller/blob/develop/COPYING.txt

Project page: https://www.pyinstaller.org/

---
```
