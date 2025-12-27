# Vendor Toggles Configuration Guide

## What is vendor_toggles.ini?

A configuration file that teaches `audioctl` how to toggle "Audio Enhancements" for your specific audio hardware. Different manufacturers store the enhancement on/off state in different registry locations.

**Location:** Same directory as `audioctl.exe`

## Do I Need This?

**NO** - if the built-in Realtek/Waves toggle works for your device.

**YES** - if:
- Enhancement toggle doesn't work for your device
- You get "No vendor toggle available" error

## Quick Start: Learning Your Device

### GUI Method (Easiest)
1. Right-click your device
2. Click **"Learn Enhancements"**
3. Read the warning, click OK
4. Set "Audio Enhancements" to **ENABLED** in Windows Sound settings
5. Click OK to capture first snapshot
6. Set "Audio Enhancements" to **DISABLED** in Windows Sound settings
7. Click OK to capture second snapshot
8. Done! Entry is automatically added to `vendor_toggles.ini`

### CLI Method
```bash
audioctl enhancements --name "Speakers" --flow Render --learn
```
Follow the interactive prompts (same as GUI).

### ⚠️ Important During Learn
- **Only** toggle "Audio Enhancements" for the target device
- **Don't** change any other audio settings
- **Don't** switch default devices
- **Don't** open other audio applications

## INI File Format

### Example Entry
```ini
[vendor_my_speakers]
value_name = {1da5d803-d492-4edd-8c23-e0c0ffee7f0e},5
dword_enable = 0
dword_disable = 1
hives = HKCU,HKLM
flows = Render,Capture
notes = Learned from Realtek speakers on 2025-12-26
```

### Field Reference

| Field | Required | Description |
|-------|----------|-------------|
| **value_name** | ✅ | Registry value name (GUID,PropertyID) |
| **dword_enable** | ✅ | Registry value when enabled (0 or 1) |
| **dword_disable** | ✅ | Registry value when disabled (0 or 1) |
| **hives** | ✅ | `HKCU` (user) and/or `HKLM` (system, needs Admin) |
| **flows** | ✅ | `Render` (playback) and/or `Capture` (recording) |
| **notes** | ❌ | Optional description |

### Hives Explained
- **HKCU** (HKEY_CURRENT_USER): Per-user, no admin needed
- **HKLM** (HKEY_LOCAL_MACHINE): System-wide, requires admin

**Recommended:** `hives = HKCU,HKLM` (tries user first, then system)

## Built-in Toggles

### Realtek/Waves (Most Common)
Works for most Realtek and Waves-based devices. No INI entry needed.

**Value:** `{1da5d803-d492-4edd-8c23-e0c0ffee7f0e},5`

## Testing Your Configuration

### Check if device has a vendor toggle:
```bash
audioctl diag-sysfx --name "Speakers" --flow Render
```

### Test toggling:
```bash
# Enable
audioctl enhancements --name "Speakers" --flow Render --enable

# Disable
audioctl enhancements --name "Speakers" --flow Render --disable
```

### Verify in Windows:
1. Right-click volume → **Sounds**
2. Select device → **Properties** → **Advanced** tab
3. Check "Audio Enhancements" setting changed

## Troubleshooting

### "No vendor toggle available"
**Solution:** Use `--learn` to teach the tool about your device.

### Toggle succeeds but Windows doesn't change
**Causes:**
- `dword_enable`/`dword_disable` values are backwards
- Wrong `value_name` for your device

**Solution:** Re-run `--learn` or use discovery mode (advanced):
```bash
audioctl discover-enhancements --name "Speakers" --flow Render
```

### "Permission denied" writing INI
**Solutions:**
- Run as Administrator
- Move `audioctl.exe` to a user-writable folder
- Manually edit `vendor_toggles.ini` as Administrator

### Changes require Administrator
**Cause:** INI specifies `hives = HKLM`

**Solution:** Change to `hives = HKCU` or `hives = HKCU,HKLM`

## Manual INI Editing

### Adding Entries
1. Open `vendor_toggles.ini` in Notepad
2. Copy an existing `[section]` or use the example above
3. Modify the values
4. Save and test

### Section Names
- Must start with `vendor_`
- Can contain letters, numbers, `_`, `-`, `{`, `}`, `,`
- Must be unique

**Examples:**
- `[vendor_realtek_speakers]`
- `[vendor_{1da5d803-d492-4edd-8c23-e0c0ffee7f0e},5]`

### Removing Entries
Delete the entire `[section]` and its lines.

### Temporarily Disabling
Add `#` at the start of each line:
```ini
# [vendor_my_device]
# value_name = {guid},pid
# ...
```

## Advanced: Discovery Mode

For detailed analysis and manual configuration:

```bash
audioctl discover-enhancements --name "Speakers" --flow Render
```

**Creates:**
- `enh-discovery_*.txt` - Human-readable report
- `enh-discovery_*.json` - Full registry snapshots

**With INI snippet:**
```bash
audioctl discover-enhancements --name "Speakers" --flow Render --ini-snippet vendor_toggles.ini
```

Follow the same prompts as `--learn`, but generates detailed reports for troubleshooting.

## Common Patterns

### Inverted Logic
Some devices use `1` for enabled instead of `0`:
```ini
dword_enable = 1
dword_disable = 0
```

### HKCU Only
For devices that don't need admin:
```ini
hives = HKCU
```

### Both Playback and Recording
One entry for both device types:
```ini
flows = Render,Capture
```

### Playback Only
```ini
flows = Render
```

## FAQ

**Q: Do I need separate entries for each device?**  
A: No. One entry works for all devices that use the same registry structure (usually all devices from the same manufacturer).

**Q: Will this work on other computers?**  
A: Yes, if they have the same audio hardware.

**Q: Is this safe?**  
A: Yes. Only writes 0 or 1 to audio configuration registry values.

**Q: Can I share my vendor_toggles.ini?**  
A: Yes! Others with the same hardware can use it.

**Q: What if learn fails?**  
A: Use `discover-enhancements` for detailed analysis, then manually create an INI entry.

## Example: Complete Learn Session

```
> audioctl enhancements --name "Speakers" --enable --learn

READ CAREFULLY
This Learn mode will capture two registry snapshots and write a vendor entry into:
  C:\Path\To\audioctl\vendor_toggles.ini

Type exactly: I UNDERSTAND
> I UNDERSTAND

Step 1: In Windows Sound settings, set 'Audio Enhancements' to ENABLED for this device.
When ready, press Enter to capture snapshot A...

Step 2: Now set 'Audio Enhancements' to DISABLED for the same device.
When ready, press Enter to capture snapshot B...

Learned vendor toggle and appended to: vendor_toggles.ini

Section: [vendor_{1da5d803-d492-4edd-8c23-e0c0ffee7f0e},5]
Value: {1da5d803-d492-4edd-8c23-e0c0ffee7f0e},5
Enabled=0, Disabled=1
```

## Getting Help

**Check logs:** `audioctl_gui.log` (same directory as executable)

**Run diagnostics:**
```bash
audioctl diag-sysfx --name "YourDevice"
```

**Generate discovery report:**
```bash
audioctl discover-enhancements --name "YourDevice" --flow Render
```

Include the discovery report JSON when reporting issues.
```
