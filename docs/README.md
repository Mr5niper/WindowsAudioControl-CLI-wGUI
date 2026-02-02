# Windows Audio Control CLI + GUI
## audioctl.exe (pycaw/comtypes-based)

A Windows audio control utility with a scriptable CLI and an optional GUI. You can:

- **Automate common tasks:** list devices, set default playback/recording endpoints, adjust volume, mute/unmute, toggle “Listen to this device,” and control Enhancements.
- **Learn vendor toggles:** “Learn Enhancements” and “Learn FX” observe what Windows/drivers change in the registry for a specific device, then store those rules so audioctl can reproduce the same toggle later.
- **Control individual effects (FX):** once learned, per-effect toggles (e.g., BassBoost, Loudness) can be enabled/disabled on demand, even when the driver does not provide a direct API.
- **Query state safely:** fast, read-only commands return current volume, mute, listen, enhancements, and FX state for scripts, hotkeys, and status UIs.
- **Use the GUI when convenient:** the GUI provides context-aware right-click actions and guided learn workflows, while still using the CLI underneath as the source of truth.

audioctl is designed to be CLI-first for repeatability and automation; the GUI is a helper layer on top.

---
[Quick Start](#quick-start)
<BR>
[Command Map (Full Tree)](#command-map-full-tree)
<BR>
[List Devices](#list-devices)
<BR>
[Set Default Endpoints](#set-default-devices-cli)
<BR>
[Adjust Volume / Mute](#set-volume-or-muteunmute-cli)
<BR>
[Listen To This Device](#listen-to-this-device-capture-only-cli)
<BR>
[Audio Enhancements](#audio-enhancements-sysfx--vendor-toggles-cli)
<BR>
[Query Helpers (Read‑Only)](#query-helpers-readonly)
<BR>
[Diagnostic Discovery](#diagnostics--discovery-optional-readonly)
<BR>
[Wait For A Device](#wait-for-a-device-to-appear-cli)
<BR>
[General Exit Codes](#general-exit-codes)
<BR>
[Troubleshooting](#troubleshooting)
<BR>
[GUI (Graphical User Interface)](#gui-graphical-user-interface)
<BR>


---

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

# Command map (full tree)

```
audioctl
 │ 
 ├─ list [--all] [--json]
 │ 
 ├─ set-default
 │  ├─ --playback-id/--playback-name [--playback-role console|multimedia|communications|all]
 │  └─ --recording-id/--recording-name [--recording-role console|multimedia|communications|all]
 │     [--index] [--regex]
 │ 
 ├─ set-volume
 │  └─ (--id | --name) [--flow Render|Capture] (--level 0..100 | --mute | --unmute)
 │     [--index] [--regex]
 │ 
 ├─ listen (Capture only)
 │  └─ (--id | --name) (--enable | --disable)
 │     [--playback-target-id [<RenderID>]] [--playback-target-name [<RenderName>]]
 │     [--index] [--regex]
 │ 
 ├─ enhancements
 │  ├─ Main switch:
 │  │  └─ (--id | --name) [--flow] (--enable | --disable | --learn)
 │  │     [--index] [--regex] [--prefer-hklm] [--vendor-ini PATH]
 │  │ 
 │  └─ FX operations:
 │     ├─ --list-fx [--json]
 │     ├─ --learn-fx "FX_NAME"
 │     ├─ --enable-fx "FX_NAME" | --disable-fx "FX_NAME"
 │     └─ --delete-fx "FX_NAME"
 │ 
 ├─ get-volume [--id|--name] [--flow] [--index] [--regex]
 │ 
 ├─ get-listen [--id|--name] [--index] [--regex]          (Capture)
 │ 
 ├─ get-enhancements [--id|--name] [--flow] [--index] [--regex]
 │ 
 ├─ get-device-state [--id|--name] [--flow] [--index] [--regex] [--vendor-ini PATH]
 │ 
 ├─ diag-sysfx [--id|--name] [--flow] [--index] [--regex]
 │ 
 ├─ diag-mmdevices [--id|--name] [--flow] [--index] [--regex]
 │ 
 ├─ discover-enhancements [--id|--name] [--flow] [--index] [--regex]
 │  └─ [--output-dir DIR] [--ini-snippet FILE]
 │ 
 └─ wait (--id|--name) [--flow] [--timeout] [--index] [--regex]
```

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
- console - the default device most apps use.
- multimedia - often resolved to the same device as console on most systems.
- communications - used by calling/telephony apps.
- all (applies to all three roles)

Defaults if a role is omitted:
- Playback role defaults to `all`.
- Recording role defaults to `communications`.

About Windows roles (console vs multimedia):
- On most Windows systems, “console” and “multimedia” effectively point to the same default device. If you set either one, you will typically see both flags update in the next `audioctl list` output. The tool still lets you set them independently, or set `all` for completeness and for environments where they’re distinct.

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

## Examples (only using “all”)
### 1) Playback (Render) - by name
```bash
audioctl set-default --playback-name "Speakers" --playback-role all --index 0
```
Sample output:
```json
{"set":[{"flow":"Render","role":"all","id":"{RENDER-ENDPOINT-ID}","name":"Speakers (Realtek(R) Audio)"}]}
```

### 2) Playback (Render) - by ID
```bash
audioctl set-default --playback-id "{RENDER-ENDPOINT-ID}" --playback-role all
```
Sample output:
```json
{"set":[{"flow":"Render","role":"all","id":"{RENDER-ENDPOINT-ID}","name":"Speakers (USB DAC)"}]}
```

### 3) Recording (Capture) - by name
```bash
audioctl set-default --recording-name "USB Microphone" --recording-role all --index 0
```
Sample output:
```json
{"set":[{"flow":"Capture","role":"all","id":"{CAPTURE-ENDPOINT-ID}","name":"USB Microphone"}]}
```

### 4) Recording (Capture) - by ID
```bash
audioctl set-default --recording-id "{CAPTURE-ENDPOINT-ID}" --recording-role all
```
Sample output:
```json
{"set":[{"flow":"Capture","role":"all","id":"{CAPTURE-ENDPOINT-ID}","name":"Headset Mic"}]}
```

### 5) Playback + Recording - both “all” in one command
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
- If list output shows both console and multimedia toggled after setting only one, that is expected on most Windows builds, they commonly map to the same endpoint.
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
- `--playback-target-name` mirrors the above using a name match; passing the flag without a value selects “Default Playback Device.”
- If the target flag is omitted, the current routing is preserved.

Command template:
```bash
audioctl listen
  (--id <CAPTURE-ENDPOINT-ID> | --name <NAME>)
  (--enable | --disable)
  [--playback-target-id [<RenderEndpointID>]]
  [--playback-target-name [<RenderEndpointName>]]
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
```bash
audioctl listen --name "Microphone" --enable --playback-target-id
```
Sample output:
```json
{"listenSet":{"id":"{CAPTURE-ENDPOINT-ID}","name":"Microphone","enabled":true}}
```

### 3) Enable Listen and route to a specific playback endpoint by ID
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
```bash
audioctl listen --name "Mic" --regex --enable --index 0
```
Sample output:
```json
{"listenSet":{"id":"{CAPTURE-ENDPOINT-ID}","name":"Studio Mic","enabled":true}}
```

### 6) Example with verification source included
```json
{"listenSet":{"id":"{CAPTURE-ENDPOINT-ID}","name":"Microphone","enabled":true,"verifiedBy":"com"}}
```
or:
```json
{"listenSet":{"id":"{CAPTURE-ENDPOINT-ID}","name":"Microphone","enabled":true,"verifiedBy":"registry"}}
```

Tips:
- This command only targets Capture devices. If you try a Render device, it won’t match.
- Use `audioctl list --json` to get both Capture (for `--id`) and Render endpoints (for routing).
- If the name matches multiple capture devices, rerun with `--index N`.

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

## Enhancement Effects (FX)
Beyond the main switch, you can manage individual effects (per-device) after learning them.

- List learned effects for a device:
  ```bash
  audioctl enhancements --id "{ENDPOINT-ID}" --list-fx
  audioctl enhancements --id "{ENDPOINT-ID}" --list-fx --json
  ```
- Learn a specific effect (guided, two-pass A/B):
  ```bash
  audioctl enhancements --id "{ENDPOINT-ID}" --flow Render --learn-fx "BassBoost"
  ```
- Toggle an effect:
  ```bash
  audioctl enhancements --id "{ENDPOINT-ID}" --flow Capture --enable-fx "Loudness"
  audioctl enhancements --id "{ENDPOINT-ID}" --flow Capture --disable-fx "Loudness"
  ```
- Delete a learned effect association for this device:
  ```bash
  audioctl enhancements --id "{ENDPOINT-ID}" --delete-fx "BassBoost"
  ```

Notes about JSON output:
- Success prints compact, single‑line JSON, e.g.:
  ```json
  {"enhancementsSet":{"id":"{...}","name":"Speakers (Realtek(R) Audio)","enabled":true,"verifiedBy":"vendor:realtek_waves_primary"}}
  ```
- FX toggles:
  ```json
  {"fxSet":{"id":"{...}","name":"Speakers","fx_name":"BassBoost","enabled":true,"verifiedBy":"vendor-fx:multi:BassBoost"}}
  ```

---

# Query helpers (read‑only)
Fast, side‑effect‑free commands perfect for status bars, hotkeys, and scripts.

- Current volume/mute:
  ```bash
  audioctl get-volume [--id <ID> | --name <NAME>] [--flow Render|Capture] [--index <N>] [--regex]
  ```
  Returns:
  ```json
  {"id":"{...}","name":"...","flow":"Render","volume":78,"muted":false}
  ```

- Current “Listen” state (Capture):
  ```bash
  audioctl get-listen [--id <ID> | --name <NAME>] [--index <N>] [--regex]
  ```
  Returns:
  ```json
  {"id":"{...}","name":"...","flow":"Capture","listenEnabled":true}
  ```

- Current Enhancements state (vendor‑only):
  ```bash
  audioctl get-enhancements [--id <ID> | --name <NAME>] [--flow Render|Capture] [--index <N>] [--regex]
  ```
  Returns:
  ```json
  {"id":"{...}","name":"...","flow":"Render","enhancementsEnabled":true}
  ```

- Aggregated device state (for GUI/scripts):
  ```bash
  audioctl get-device-state [--id <ID> | --name <NAME>] [--flow Render|Capture] [--index <N>] [--regex] [--vendor-ini PATH]
  ```
  Returns:
  ```json
  {
    "id":"{...}",
    "name":"...",
    "flow":"Render",
    "volume":78,
    "muted":false,
    "listenEnabled":null,
    "enhancementsEnabled":true,
    "availableFX":[{"fx_name":"BassBoost","state":true,"source":"ini"}]
  }
  ```

---

# Diagnostics & Discovery (Optional, Read‑Only)
- View live Enhancements status and vendor state:
  ```bash
  audioctl diag-sysfx --name "Speakers" --flow Render
  ```
- Dump all MMDevices values for an endpoint (HKCU/HKLM; FxProperties/Properties):
  ```bash
  audioctl diag-mmdevices --id "{RENDER-ENDPOINT-ID}" --flow Render
  ```
- Interactive discovery with reports (TXT/JSON) and optional INI snippet:
  ```bash
  audioctl discover-enhancements --name "Speakers" --flow Render --output-dir "." --ini-snippet ".\vendor_snippets.ini"
  ```

---

# Wait for a Device to Appear (CLI)
Use `audioctl wait` to block until a device becomes active (appears) or until a timeout expires. Matching is against active endpoints only (`DEVICE_STATE_ACTIVE`). When a match is found, the command prints a compact, single‑line JSON object and exits with code 0.

- Template:
  ```bash
  audioctl wait
    (--id <ENDPOINT-ID> | --name <NAME>) [--flow <Render|Capture>]
    [--timeout <seconds>] [--index <N>] [--regex]
  ```
- Exit codes:
  - 0: found a matching active device
  - 3: timeout
  - 4: multiple matches (rerun with `--index`)
- Tips:
  - Prefer `--id` when you know it.
  - Add `--flow` when the name appears in both flows.

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
- Enhancements don’t toggle:
  - First try the toggle directly:
    ```
    audioctl enhancements --id "{ENDPOINT-ID}" --flow Render --enable
    ```
    If you get:
    ```
    ERROR: No vendor toggle available for this device. Use --learn to teach a vendor method.
    ```
    then proceed to learn.
  - Learn the vendor toggle (manual A/B):
    ```
    audioctl enhancements --id "{ENDPOINT-ID}" --flow Render --learn
    ```
    After a successful learn, re-run the toggle:
    ```
    audioctl enhancements --id "{ENDPOINT-ID}" --flow Render --disable
    ```
  - If learn fails or the result can’t be verified:
    - Inspect live state vs vendor with:
      ```
      audioctl diag-sysfx --id "{ENDPOINT-ID}" --flow Render
      ```
    - Capture ON/OFF snapshots and generate a report/snippet:
      ```
      audioctl discover-enhancements --id "{ENDPOINT-ID}" --flow Render --output-dir "." --ini-snippet ".\vendor_snippets.ini"
      ```
      Use the suggested INI snippet or share the TXT/JSON for analysis.
  - Permissions and hive preference:
    - Some drivers only honor HKLM. Re-run the toggle with:
      ```
      audioctl enhancements --id "{ENDPOINT-ID}" --flow Render --enable --prefer-hklm
      ```
      (Run an elevated console if HKLM writes are required.)
  - Confirm selection and INI placement:
    - Make sure you’re targeting the correct endpoint/flow (use `--flow` and `--index` if needed).
    - Verify `vendor_toggles.ini` is at the expected path (exe directory if writable, otherwise `%LOCALAPPDATA%\audioctl\vendor_toggles.ini`) and that your endpoint GUID is listed under `devices =` in the learned section.

---

# GUI (Graphical User Interface)
- Launch the GUI:
  - Double-click `audioctl.exe` (no arguments), or run it without command-line options.
  <br>
    <img width="98" height="98" alt="image" src="https://github.com/user-attachments/assets/a8bacaf0-71e9-45c6-bfa2-676647c0157c" />
  <br>
  
<img width="1367" height="357" alt="image" src="https://github.com/user-attachments/assets/efd4d7bd-ce8e-4be8-b352-9fdea7af92b8" />
<br>

### Features
- Visual list of playback (Render) and recording (Capture) devices, grouped and sorted.
- **“Show disabled/disconnected”** toggle.
- **Right-click** or **double-click** a device for actions:
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img width="356" height="237" alt="image" src="https://github.com/user-attachments/assets/460f6740-1eac-4f92-8afe-db89675661f8" />
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<br>

  - **“Set as Default (all roles)”**
  - “Set Volume…”
    <br>
    <img width="218" height="146" alt="image" src="https://github.com/user-attachments/assets/b58c8b9e-df50-411a-ae13-d341d61ce332" />
    <br>
  - **“Mute” / “Unmute”**
  - **“Toggle Listen (capture only)”**
  - **“Enable/Disable Enhancements”** (when a vendor toggle is known)
  - **“Enhancement Effects”** (per‑effect toggles learned from your INI)
    <img width="642" height="237" alt="image" src="https://github.com/user-attachments/assets/d79cf646-6fb0-43fc-b417-85c88885583e" />
  
  - **“Learn Enhancements”** (discover a vendor DWORD and/or effects for that device)<br>
    <img width="346" height="238" alt="image" src="https://github.com/user-attachments/assets/ebb033e9-89fb-4a34-a92a-165d443fdd7a" />


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
- pycaw - <https://github.com/AndreMiras/pycaw>  
- comtypes - <https://github.com/enthought/comtypes>  

Windows GUIDs and PROPERTYKEY constants are Microsoft API identifiers and do not require attribution.

---

## Third‑Party License Texts

### pycaw - MIT License
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

### comtypes - MIT License
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
```

