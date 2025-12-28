# Windows Audio Control CLI + GUI
## audioctl.exe (pycaw-based)
[List devices](#list-devices)
<BR>
[Set default endpoints](#set-default-devices-cli)
<BR>
[Adjust volume / mute](#set-volume-or-muteunmute-cli)
<BR>
[Listen to this device](#listen-to-this-device-capture-only-cli)
<BR>
[Audio enhancements](#audio-enhancements-sysfx--vendor-toggles-cli)
<BR>
[Wait for a device](#wait-for-a-device-to-appear-cli)
<BR>
[Troubleshooting](#troubleshooting)
<BR>
[GUI (Graphical User Interface)](#gui-graphical-user-interface)
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
- `--flow Render|Capture` (optional filter for playback/recording device)
- `--index N` (Number-based index to disambiguate between multiple matches)
- `--playback-name "substring"` (Render-only name selector; for set-default)
- `--recording-name "substring"` (Capture-only name selector; for set-default)

When using the `--name`, `--playback-name`, or `--recording-name` options, commands accept partial matches against device names or descriptions. If the specified text uniquely identifies exactly one playback or recording device, that device will be targeted and an index value is not required.

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
- console — the default device most apps use.
- multimedia — often resolved to the same device as console on most systems.
- communications — used by calling/telephony apps.
- all (applies to all three roles)

Defaults if a role is omitted:
- Playback role defaults to `all`.
- Recording role defaults to `communications`.

About Windows roles (console vs multimedia):
- On most Windows systems, “console” and “multimedia” effectively point to the same default device. If you set either one, you will typically see both flags update in the next `audioctl list` output. The tool still lets you set them independently—or set `all`—for completeness and for environments where they’re distinct.

Most users want a device to be the default for everything. Use:
- `--playback-role all` for playback (Render)
- `--recording-role all` for recording (Capture)

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

## Examples (only using “all”)

### 1) Playback (Render) — by name
```bash
audioctl set-default --playback-name "Speakers" --playback-role all --index 0
```
Sample output:
```json
{"set":[{"flow":"Render","role":"all","id":"{RENDER-ENDPOINT-ID}","name":"Speakers (Realtek(R) Audio)"}]}
```

### 2) Playback (Render) — by ID
```bash
audioctl set-default --playback-id "{RENDER-ENDPOINT-ID}" --playback-role all
```
Sample output:
```json
{"set":[{"flow":"Render","role":"all","id":"{RENDER-ENDPOINT-ID}","name":"Speakers (USB DAC)"}]}
```

### 3) Recording (Capture) — by name
```bash
audioctl set-default --recording-name "USB Microphone" --recording-role all --index 0
```
Sample output:
```json
{"set":[{"flow":"Capture","role":"all","id":"{CAPTURE-ENDPOINT-ID}","name":"USB Microphone"}]}
```

### 4) Recording (Capture) — by ID
```bash
audioctl set-default --recording-id "{CAPTURE-ENDPOINT-ID}" --recording-role all
```
Sample output:
```json
{"set":[{"flow":"Capture","role":"all","id":"{CAPTURE-ENDPOINT-ID}","name":"Headset Mic"}]}
```

### 5) Playback + Recording — both “all” in one command
```bash
audioctl set-default \
  --playback-id "{RENDER-ENDPOINT-ID}"   --playback-role all \
  --recording-id "{CAPTURE-ENDPOINT-ID}" --recording-role all
```
Sample output:
```json
{"set":[{"flow":"Render","role":"all","id":"{RENDER-ENDPOINT-ID}","name":"Speakers (Realtek(R) Audio)"},{"flow":"Capture","role":"all","id":"{CAPTURE-ENDPOINT-ID}","name":"USB Microphone"}]}
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

# “Listen to this device” (Capture Only, CLI)

Use `audioctl listen` to enable or disable “Listen to this device” on a capture (recording) endpoint. The target must be active (`DEVICE_STATE_ACTIVE`). This command only operates on Capture devices; Render devices are not eligible.

You must specify exactly one of:
- `--enable`
- `--disable`

Optional playback routing:
- `--playback-target-id "{RENDER-ENDPOINT-ID}"` routes the monitored audio to a specific playback (Render) endpoint.
- `--playback-target-id ""` routes to the Windows “Default Playback Device.”
- If `--playback-target-id` is omitted:
  - The current routing target is preserved.
  - When enabling and no target exists, the tool sets it to `""` (Default Playback Device).

Command template:
```bash
audioctl listen
  (--id <CAPTURE-ENDPOINT-ID> | --name <NAME>)
  (--enable | --disable)
  [--playback-target-id <RenderEndpointID-or-empty>]
  [--index <N>] [--regex]
```

Disambiguation behavior:
- If name selection matches more than one active capture device and you don’t pass `--index`, the command prints the matching candidates in GUI order (with indices) and exits with:
  ```
  ERROR: multiple matches; specify --index
  ```
  Rerun with `--index N` (0‑based, GUI‑order for the Capture flow). Use `audioctl list` (or `--json`) to see the same order.

Notes about JSON output:
- Success prints compact, single‑line JSON.
- When a retry/verification path is used, the JSON may include a `verifiedBy` field (`"com"` or `"registry"`).

---

## Examples

### 1) Enable Listen for a microphone by name (preserve current playback target)
```bash
audioctl listen --name "Microphone" --enable
```
Sample output:
```json
{"listenSet":{"id":"{CAPTURE-ENDPOINT-ID}","name":"Microphone","enabled":true}}
```

### 2) Enable Listen and route to the Default Playback Device
Explicitly set the routing to the default output by passing an empty string.
```bash
audioctl listen --name "Microphone" --enable --playback-target-id ""
```
Sample output:
```json
{"listenSet":{"id":"{CAPTURE-ENDPOINT-ID}","name":"Microphone","enabled":true}}
```

### 3) Enable Listen and route to a specific playback endpoint by ID
Find the Render endpoint ID with `audioctl list --json` (flow="Render"), then:
```bash
audioctl listen --name "Microphone" --enable --playback-target-id "{RENDER-ENDPOINT-ID}"
```
Sample output:
```json
{"listenSet":{"id":"{CAPTURE-ENDPOINT-ID}","name":"Microphone","enabled":true}}
```

### 4) Disable Listen by exact ID
```bash
audioctl listen --id "{CAPTURE-ENDPOINT-ID}" --disable
```
Sample output:
```json
{"listenSet":{"id":"{CAPTURE-ENDPOINT-ID}","name":"USB Microphone","enabled":false}}
```

### 5) Use regex name matching with index disambiguation (Capture)
If multiple capture devices match “Mic”, choose one with `--index`:
```bash
audioctl listen --name "Mic" --regex --enable --index 0
```
Sample output:
```json
{"listenSet":{"id":"{CAPTURE-ENDPOINT-ID}","name":"Studio Mic","enabled":true}}
```

### 6) Example with verification source included
When the initial COM write needs verification, you may see:
```json
{"listenSet":{"id":"{CAPTURE-ENDPOINT-ID}","name":"Microphone","enabled":true,"verifiedBy":"com"}}
```
or:
```json
{"listenSet":{"id":"{CAPTURE-ENDPOINT-ID}","name":"Microphone","enabled":true,"verifiedBy":"registry"}}
```

---

Tips:
- This command only targets Capture devices. If you try a Render device, it won’t match.
- Use `audioctl list --json` to get both Capture (for `--id`) and Render endpoints (for `--playback-target-id`).
- If you see “device not found (active only)”, ensure the microphone is connected/enabled and visible in `audioctl list`.
- If the name matches multiple capture devices, rerun with `--index N` using the indices shown in the disambiguation prompt.

---

# Audio Enhancements (SysFX) – Vendor Toggles (CLI)

Use `audioctl enhancements` to enable or disable “Audio Enhancements” (SysFX) on an endpoint using a vendor‑first method. Many drivers (e.g., Realtek/Waves) expose Enhancements via device‑specific registry DWORDs under MMDevices; this tool writes those DWORDs (when known) and verifies the change.

Key points:
- Control is vendor‑only at runtime: if no vendor toggle is known for the target endpoint, the command fails and asks you to use `--learn` first.
- Windows’ `Disable_SysFx` is still read for diagnostics, not used to control state here.
- HKLM writes require Administrator privileges; use `--prefer-hklm` if your driver only honors HKLM.

You must specify exactly one of:
- `--enable`    (turn Enhancements ON)
- `--disable`   (turn Enhancements OFF)
- `--learn`     (guided manual learning to append a vendor entry to INI)

Command template:
```bash
# Toggle using vendor DWORDs (if known)
audioctl enhancements
  (--id <ENDPOINT-ID> | --name <NAME>) [--flow <Render|Capture>]
  (--enable | --disable)
  [--index <N>] [--regex]
  [--prefer-hklm]
  [--vendor-ini <PATH>]

# Learn a vendor DWORD (you toggle the Windows UI; the tool captures A/B and appends to INI)
audioctl enhancements
  (--id <ENDPOINT-ID> | --name <NAME>) [--flow <Render|Capture>]
  --learn
  [--index <N>] [--regex]
  [--vendor-ini <PATH>]
```

Disambiguation behavior:
- If name selection matches more than one active device and you don’t pass `--index`, the tool prints candidates in GUI order (with indices) and exits:
  ```
  ERROR: multiple matches; specify --index
  ```
  Rerun with `--index N` (0‑based, GUI order for that flow). Use `audioctl list` to see the same order.

Notes about JSON output:
- Success prints compact, single‑line JSON, e.g.:
  ```json
  {"enhancementsSet":{"id":"{...}","name":"Speakers (Realtek(R) Audio)","enabled":true,"verifiedBy":"vendor:realtek_waves_primary"}}
  ```
- `enabled` is the final Enhancements state (True/False).
- `verifiedBy` indicates how the tool confirmed the result (usually a vendor tag like `vendor:realtek_waves_primary`).

---

## Examples (Toggle)

### 1) Enable Enhancements by name (Render)
```bash
audioctl enhancements --name "Speakers" --flow Render --enable
```
Sample output:
```json
{"enhancementsSet":{"id":"{RENDER-ENDPOINT-ID}","name":"Speakers (Realtek(R) Audio)","enabled":true,"verifiedBy":"vendor:realtek_waves_primary"}}
```

### 2) Disable Enhancements by exact ID and prefer HKLM (Admin recommended)
```bash
audioctl enhancements --id "{RENDER-ENDPOINT-ID}" --disable --prefer-hklm
```
Sample output:
```json
{"enhancementsSet":{"id":"{RENDER-ENDPOINT-ID}","name":"Speakers (USB DAC)","enabled":false,"verifiedBy":"vendor:realtek_waves_primary"}}
```

### 3) Toggle Enhancements for a capture device (microphone)
```bash
audioctl enhancements --name "USB Microphone" --flow Capture --disable
```
Sample output:
```json
{"enhancementsSet":{"id":"{CAPTURE-ENDPOINT-ID}","name":"USB Microphone","enabled":false,"verifiedBy":"vendor:realtek_waves_primary"}}
```

---

## Learn a Vendor Toggle (Manual)

Use `--learn` when the device has no known vendor toggle. You will be prompted to change the Windows UI setting (Enhancements ON, then OFF) while the tool captures A/B and appends a vendor entry to the INI if a reliable DWORD flip is found.

```bash
audioctl enhancements --name "Speakers" --flow Render --learn --vendor-ini "C:\path\vendor_toggles.ini"
```

Sample success output:
```json
{"vendorLearned":{"id":"{RENDER-ENDPOINT-ID}","name":"Speakers (Realtek(R) Audio)","flow":"Render","iniPath":"C:\\path\\vendor_toggles.ini","section":"vendor_{1da5d803-d492-4edd-8c23-e0c0ffee7f0e},5","value_name":"{1da5d803-d492-4edd-8c23-e0c0ffee7f0e},5","dword_enable":0,"dword_disable":1}}
```

INI entry semantics:
- `value_name`: the MMDevices DWORD key (e.g., `{GUID},pid`) under FxProperties/Properties for this endpoint.
- `dword_enable`: DWORD value meaning “Enhancements ON” for this vendor key.
- `dword_disable`: DWORD value meaning “Enhancements OFF”.
- `hives`: default write order is `HKCU,HKLM` (changeable during learn or by editing the INI).
- You can omit `--vendor-ini` to use the default `vendor_toggles.ini` next to the executable.

---

## Diagnostics & Discovery (Optional, Read‑Only)

### View live Enhancements status and vendor state
```bash
audioctl diag-sysfx --name "Speakers" --flow Render
```
Sample output (pretty‑printed here for readability; real output is single‑line JSON):
```json
{
  "id":"{RENDER-ENDPOINT-ID}",
  "name":"Speakers (Realtek(R) Audio)",
  "flow":"Render",
  "enhancementsEnabled_live_propstore": true,
  "enhancementsEnabled_live_com": true,
  "vendor_toggle_status": {"realtek_waves_primary ({1da5d803...},5)": true}
}
```

### Dump all MMDevices values for an endpoint (HKCU/HKLM; FxProperties/Properties)
```bash
audioctl diag-mmdevices --id "{RENDER-ENDPOINT-ID}" --flow Render
```
This is useful to debug drivers or confirm what changed during learn.

### Interactive discovery with reports (TXT/JSON) and optional INI snippet
```bash
audioctl discover-enhancements --name "Speakers" --flow Render --output-dir "." --ini-snippet ".\vendor_snippets.ini"
```
- Prompts you to set Enhancements ON → snapshot A, then OFF → snapshot B.
- Writes a human‑readable TXT and a JSON bundle to `--output-dir`.
- Optionally appends a suggested INI section to the file you provide via `--ini-snippet`.

---

## Selection & Errors

- Prefer `--id` for exact targeting.
- With `--name`, also pass `--flow` when the same text appears for both Render and Capture.
- If more than one match remains and `--index` is omitted:
  ```
  ERROR: multiple matches; specify --index
  ```
- If no vendor toggle exists for the device:
  ```
  ERROR: No vendor toggle available for this device. Use --learn to teach a vendor method.
  ```

---

## Notes & Permissions

- HKCU writes: no admin needed.
- HKLM writes: require Administrator (use `--prefer-hklm` if your driver ignores HKCU).
- The tool verifies by reading back the same vendor DWORD (short retries) and returns the final state with a `verifiedBy` tag.
- `vendor_toggles.ini` lives next to `audioctl.exe` by default. Use `--vendor-ini <PATH>` to read/write a different file.

---

# Wait for a Device to Appear (CLI)

Use `audioctl wait` to block until a device becomes active (appears) or until a timeout expires. Matching is against active endpoints only (`DEVICE_STATE_ACTIVE`). When a match is found, the command prints a compact, single‑line JSON object and exits with code 0.

You must specify one of:
- `--id "{ENDPOINT-ID}"` (exact match), or
- `--name "substring"` (case‑insensitive; add `--regex` for regular expressions)

Optional:
- `--flow Render|Capture` to restrict the search to playback (Render) or recording (Capture)
- `--index N` if multiple active matches exist (0‑based, GUI‑order within the flow)
- `--timeout SECONDS` (default: 30)

Command template:
```bash
audioctl wait
  (--id <ENDPOINT-ID> | --name <NAME>) [--flow <Render|Capture>]
  [--timeout <seconds>] [--index <N>] [--regex]
```

Behavior:
- Polls every ~0.5s until the first match is found or the timeout elapses.
- If `--flow` is omitted, results from Render and Capture are combined, and the first in GUI order is chosen (or use `--index` to select).
- On success, prints `{"found": {...device...}}` as a single line and exits 0.
- On timeout, prints an error line and exits 3.

Disambiguation:
- If your selection matches more than one active device and you don’t pass `--index`, the command exits with:
  ```
  ERROR: multiple matches; specify --index
  ```
  Rerun with `--index N` (0‑based, GUI‑order for that flow), or make the `--name` more specific and/or add `--flow`.

Notes about JSON output:
- Success prints a compact, single‑line JSON object like:
  ```json
  {"found":{"id":"{...}","name":"Speakers (Realtek(R) Audio)","flow":"Render","state":"active","isDefault":{"console":true,"multimedia":true,"communications":false},"guiIndex":0}}
  ```

---

## Examples

### 1) Wait up to 60 seconds for a capture device by name
```bash
audioctl wait --name "USB Microphone" --flow Capture --timeout 60
```
Sample output (single line on success):
```json
{"found":{"id":"{CAPTURE-ENDPOINT-ID}","name":"USB Microphone","flow":"Capture","state":"active","isDefault":{"console":false,"multimedia":false,"communications":true},"guiIndex":0}}
```

### 2) Wait for an endpoint by exact ID (default timeout 30s)
```bash
audioctl wait --id "{RENDER-ENDPOINT-ID}"
```
Sample output:
```json
{"found":{"id":"{RENDER-ENDPOINT-ID}","name":"Speakers (Realtek(R) Audio)","flow":"Render","state":"active","isDefault":{"console":true,"multimedia":true,"communications":false},"guiIndex":0}}
```

### 3) Use regex matching (Render), with index disambiguation
If multiple Render devices match “Speakers”, choose one with `--index`:
```bash
audioctl wait --name "^Speakers" --regex --flow Render --timeout 20 --index 1
```
Sample output:
```json
{"found":{"id":"{RENDER-ENDPOINT-ID}","name":"Speakers (USB DAC)","flow":"Render","state":"active","isDefault":{"console":false,"multimedia":false,"communications":false},"guiIndex":1}}
```

### 4) Wait without `--flow` (Render+Capture combined)
If your name appears in both flows, you may need `--index` to pick which one:
```bash
audioctl wait --name "USB" --timeout 45 --index 0
```
Sample output:
```json
{"found":{"id":"{CAPTURE-ENDPOINT-ID}","name":"USB Microphone","flow":"Capture","state":"active","isDefault":{"console":false,"multimedia":false,"communications":true},"guiIndex":0}}
```

---

Exit codes:
- `0` – found a matching active device and printed JSON
- `3` – timeout expired without finding a match
- `4` – multiple matches; rerun with `--index`

Tips:
- Prefer `--id` when you know it (no disambiguation needed).
- With `--name`, add `--flow` when the same text appears for both Render and Capture.
- Use `audioctl list --json` to discover IDs and confirm device names/flows before waiting.

---

# General Exit Codes

- `0` – success
- `1` – runtime failure or invalid option combo
- `3` – device not found or wait timeout
- `4` – multiple matches; provide `--index`
- `130` – Ctrl‑C interrupted

---

# Troubleshooting

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

# GUI (Graphical User Interface)

- Launch the GUI:
  - Double-click `audioctl.exe` (no arguments), or run it without command-line options.
    <br>
    <img width="88" height="102" alt="image" src="https://github.com/user-attachments/assets/6e57a1d1-88c6-4a0d-9ff0-033e4d80452f" />

  <br>
  
<img width="1367" height="357" alt="image" src="https://github.com/user-attachments/assets/58d35c73-2cdd-44d4-9704-62817a97e558" />
<br>

### Features

- Visual list of playback (Render) and recording (Capture) devices, grouped and sorted.
- **“Show disabled/disconnected”** toggle.
- **Right-click** or **double-click** a device for actions:
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img width="344" height="224" alt="image" src="https://github.com/user-attachments/assets/8745eae7-b9d8-415a-bb7a-3feebd739300" />
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<br>

  - **“Set as Default (all roles)”**
  - “Set Volume…”
    <br>
    <img width="245" height="146" alt="image" src="https://github.com/user-attachments/assets/010de5dc-28ca-407e-affe-fba7b2161d38" />
    <br>

  - **“Mute” / “Unmute”**
  - **“Toggle Listen (capture only)”**
  - **“Enable/Disable Enhancements”** (when a vendor toggle is known)
  - **“Learn Enhancements”** (discover a vendor DWORD for that device)

### Print CLI Commands

- Enable **“Print CLI commands”** in the top bar.  
  Every GUI action will print the equivalent `audioctl.exe` command to stdout (useful for learning the CLI and scripting).


<img width="420" height="71" alt="image" src="https://github.com/user-attachments/assets/8fbc4be3-02aa-43aa-93d3-fe7835dc6bfd" />
    <br>
    <br>

![CLI Output](https://github.com/user-attachments/assets/1f52d28b-71fa-4343-abde-728b164d2a99)

### Refresh Devices

- Click **“Refresh”** or press `F5`.

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
