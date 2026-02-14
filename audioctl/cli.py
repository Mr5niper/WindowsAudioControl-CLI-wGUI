# audioctl/cli.py
#
# This module defines the *public* CLI surface for audioctl.
# It is intentionally the single "command router" used by:
#   - humans on the command line,
#   - scripts/automation (JSON outputs),
#   - and the Tkinter GUI (which shells out to the CLI rather than importing
#     low-level COM helpers directly).
#
# Why the GUI uses the CLI instead of calling COM helpers:
#   - Stability: the low-level Windows audio operations involve COM (comtypes/pycaw)
#     and, in some cases, raw vtable calls. Keeping that logic inside a CLI boundary
#     reduces the chance of long-lived Tk event-loop interactions with COM lifetime
#     and garbage collection timing.
#   - Single source of truth: CLI JSON output is the stable contract; the GUI consumes
#     the same results that scripts do.
#
# IMPORTANT:
#   compat.py MUST be imported before anything that might touch comtypes/pycaw.
#   It installs small compatibility shims needed for certain comtypes builds and
#   ensures PyInstaller bundles critical COM cleanup modules.
# --- stdlib imports ---
import sys
import argparse
import json
import time
import io
import os
import re
from contextlib import redirect_stderr
# --- local/project imports (do these early; compat must run before comtypes usage) ---
# Import compat BEFORE any comtypes usage
from .compat import (
    ROLES, is_admin,
)
from .logging_setup import _log, _log_exc
# --- device helpers (CLI stays thin; COM/registry implementation lives in devices.py) ---
from .devices import (
    list_devices, find_devices_by_selector, _sort_and_tag_gui_indices,
    _pretty_matches_msg, _select_by_name_active_only,
    set_default_endpoint,
    set_endpoint_mute, get_endpoint_mute,
    get_endpoint_volume, set_endpoint_volume,
    set_listen_to_device_ps, _get_listen_to_device_status_ps,
    _verify_listen_via_registry,
    _dump_mmdevices_all_values,
    _reemit_non_error_stderr,
    _collect_sysfx_snapshot,
    _diff_mmdevices_lists,
    _generate_enh_discovery_report,
    _get_enhancements_status_propstore,
    _get_enhancements_status_com,
    _read_listen_enable_fast,
)
# --- vendor/INI helpers (vendor-first enhancements/FX; parsing and registry writing live in vendor_db.py) ---
from .vendor_db import (
    _vendor_ini_default_path,
    _enhancements_supported,
    _apply_enhancements,
    _learn_vendor_from_discovery_and_write_ini,
    _build_vendor_ini_snippet,
    _find_first_vendor_entry,
    _read_vendor_entry_state,
    _list_fx_for_device,
    _learn_fx_and_write_ini,
    _fast_get_enhancements_state,
    _fast_read_vendor_entry_state,
)

def _resolve_standard_target(args, flow_forced=None):
    """
    Standard logic to find a SINGLE target device while respecting global GUI indices.
    1. List active devices.
    2. Tag ALL of them with global guiIndex (so indices match 'audioctl list').
    3. Filter by selector (id/name/flow).
    4. Handle disambiguation (--index).
    Returns (device_dict, None) or (None, error_msg).
    """
    devices = list_devices(include_all=False)
    _sort_and_tag_gui_indices(devices)

    f = flow_forced if flow_forced else getattr(args, "flow", None)
    matches = find_devices_by_selector(devices, dev_id=args.id, name_substr=args.name, flow=f, regex=args.regex)

    if not matches:
        return None, "ERROR: device not found (active only)"

    if len(matches) > 1 and args.index is None:
        return None, _pretty_matches_msg("device", matches) + "\nUse --index to disambiguate."

    if args.index is not None:
        # Select by the tagged guiIndex
        target = next((d for d in matches if d.get("guiIndex") == args.index), None)
        if not target:
             valid = sorted([d.get("guiIndex") for d in matches if "guiIndex" in d])
             return None, f"ERROR: --index {args.index} does not match any found device. (Found indices: {valid})"
        return target, None

    return matches[0], None

def cmd_list(args):
    # list:
    #   - Enumerate endpoints via devices.list_devices().
    #   - By default includes ACTIVE endpoints only; --all includes disabled/unplugged/etc.
    #
    # Indexing note (cross-cutting behavior for many commands):
    #   We sort devices by name within each flow (Render/Capture) and attach guiIndex.
    #   This mirrors the GUI order and makes --index stable and predictable.
    #
    # Output:
    #   - --json: { "devices": [ {id,name,flow,state,isDefault,guiIndex}, ... ] }
    #   - otherwise: human-readable sections with [guiIndex] labels.
    #
    # Exit codes:
    #   0 success
    devices = list_devices(include_all=args.all)
    buckets = _sort_and_tag_gui_indices(devices)
    if args.json:
        print(json.dumps({"devices": devices}, indent=2))
        return 0
    print("--- Playback (Render) ---")
    for d in buckets["Render"]:
        flags = [k for k, v in d["isDefault"].items() if v]
        print(f"[{d['guiIndex']}] {d['name']}  id={d['id']}  defaults={','.join(flags) if flags else '-'}\n")
    print("\n--- Recording (Capture) ---")
    for d in buckets["Capture"]:
        flags = [k for k, v in d["isDefault"].items() if v]
        print(f"[{d['guiIndex']}] {d['name']}  id={d['id']}  defaults={','.join(flags) if flags else '-'}\n")
    return 0

def cmd_set_default(args):
    # set-default:
    #   Set default playback and/or recording endpoints using PolicyConfig.
    #
    # Selection:
    #   - Active-only (we intentionally refuse to target disabled/notpresent devices).
    #   - playback selector: --playback-id (exact) OR --playback-name (substring/regex),
    #     with --index interpreted as GUI-order within the Render flow.
    #   - recording selector: --recording-id OR --recording-name, with --index interpreted
    #     as GUI-order within the Capture flow.
    #
    # Roles:
    #   - console/multimedia/communications/all (see ROLES).
    #   - Playback defaults to "all" because most users want the same endpoint for all roles.
    #   - Recording defaults to "communications" historically; you can choose "all" explicitly.
    #
    # Output:
    #   JSON: {"set":[{"flow":"Render|Capture","role":"...","id":"...","name":"..."}]}
    #
    # Exit codes:
    #   0 success (or partial success; failures set exit_code=1 but still prints JSON)
    #   1 runtime failure setting one or more defaults
    #   3 not found (active-only)
    #   4 multiple matches / index required (selection ambiguity)
    if not is_admin():
        # "Might" is intentional: some systems allow PolicyConfig calls without elevation,
        # others require Administrator depending on policy/drivers.
        print("WARNING: 'set-default' might require Administrator privileges on this system.", file=sys.stderr)
    
    exit_code = 0
    results = {"set": []}
    if args.playback_id or args.playback_name:
        flow_name = "Render"
        if args.playback_id:
            matches = find_devices_by_selector(list_devices(include_all=False), dev_id=args.playback_id, flow=flow_name, regex=args.regex)
            if len(matches) == 0:
                print("ERROR: playback device not found (active only)", file=sys.stderr)
                return 3
            target = matches[0]
        else:
            # _select_by_name_active_only enforces GUI-order indexing and active-only selection.
            target, err = _select_by_name_active_only(flow_name, args.playback_name, args.index, args.regex)
            if err:
                print(err, file=sys.stderr)
                return 4 if "Multiple" in err else 3
        
        try:
            role = args.playback_role
            set_default_endpoint(target["id"], role)
            results["set"].append({"flow": "Render", "role": role, "id": target["id"], "name": target["name"]})
        except Exception as e:
            print(f"ERROR: failed to set playback default: {e}", file=sys.stderr)
            exit_code = 1
    if args.recording_id or args.recording_name:
        flow_name = "Capture"
        if args.recording_id:
            matches = find_devices_by_selector(list_devices(include_all=False), dev_id=args.recording_id, flow=flow_name, regex=args.regex)
            if len(matches) == 0:
                print("ERROR: recording device not found (active only)", file=sys.stderr)
                return 3
            target = matches[0]
        else:
            # _select_by_name_active_only enforces GUI-order indexing and active-only selection.
            target, err = _select_by_name_active_only(flow_name, args.recording_name, args.index, args.regex)
            if err:
                print(err, file=sys.stderr)
                return 4 if "Multiple" in err else 3
        
        try:
            role = args.recording_role
            set_default_endpoint(target["id"], role)
            results["set"].append({"flow": "Capture", "role": role, "id": target["id"], "name": target["name"]})
        except Exception as e:
            print(f"ERROR: failed to set recording default: {e}", file=sys.stderr)
            exit_code = 1
    print(json.dumps(results))
    return exit_code

def cmd_set_volume(args):
    # set-volume:
    #   Mutate endpoint state:
    #     - set master volume scalar (0..100) OR
    #     - mute/unmute
    #
    # Validation:
    #   - Exactly one of: --level, --mute, --unmute.
    #
    # Selection:
    #   - Active-only.
    #   - Match by --id (exact) or --name (substring/regex), optional --flow filter.
    #   - If ambiguous and --index not provided: exit 4.
    #
    # Indexing:
    #   When selecting among matches, we sort and tag guiIndex using the same logic
    #   as the GUI (name-sorted within each flow). This keeps --index stable.
    #
    # Output:
    #   - {"volumeSet": {...}} or {"muteSet": {...}} on success.
    #
    # Exit codes:
    #   0 success
    #   1 invalid args or failed to set
    #   3 not found (active-only)
    #   4 ambiguous / --index out of range
    if (args.mute or args.unmute) and args.level is not None:
        print("ERROR: Cannot specify both --level and --mute/--unmute", file=sys.stderr)
        return 1
    if not (args.mute or args.unmute or args.level is not None):
        print("ERROR: Must specify --level or --mute/--unmute", file=sys.stderr)
        return 1

    target, err = _resolve_standard_target(args)
    if err:
        print(err, file=sys.stderr)
        return 4 if "Multiple" in err or "index" in err else 3

    ok = False
    if args.mute:
        ok = set_endpoint_mute(target["id"], True)
        if ok:
            print(json.dumps({"muteSet": {"id": target["id"], "name": target["name"], "muted": True}}))
    elif args.unmute:
        ok = set_endpoint_mute(target["id"], False)
        if ok:
            print(json.dumps({"muteSet": {"id": target["id"], "name": target["name"], "muted": False}}))
    elif args.level is not None:
        ok = set_endpoint_volume(target["id"], args.level)
        if ok:
            print(json.dumps({"volumeSet": {"id": target["id"], "name": target["name"], "level": args.level}}))
    if not ok:
        print("ERROR: failed to set volume/mute", file=sys.stderr)
        return 1
    return 0

def cmd_get_volume(args):
    """
    Get current volume and mute status for a device.
    This is a read-only helper so the GUI (or scripts) can query state
    without using low-level code directly.
    """
    # get-volume:
    #   Read-only query used heavily by GUI/script polling.
    #
    # Selection:
    #   - Active-only.
    #   - Same selection/disambiguation rules as set-volume.
    #
    # Output JSON:
    #   { "id": "...", "name": "...", "flow": "Render|Capture", "volume": int|null, "muted": bool|null }
    #
    # Exit codes:
    #   0 success
    #   3 not found
    #   4 ambiguous / --index out of range
    target, err = _resolve_standard_target(args)
    if err:
        print(err, file=sys.stderr)
        return 4 if "Multiple" in err or "index" in err else 3

    vol = get_endpoint_volume(target["id"])
    muted = get_endpoint_mute(target["id"])
    # Normalize muted to a plain bool/null-like; don't let odd types leak
    # (Some COM wrappers can return non-bool truthy values; GUI/scripts want stable JSON.)
    if muted is not None:
        muted = bool(muted)
    result = {
        "id": target["id"],
        "name": target["name"],
        "flow": target["flow"],
        "volume": vol,
        "muted": muted,
    }
    print(json.dumps(result))
    return 0

def cmd_listen(args):
    # listen:
    #   Enable/disable "Listen to this device" for a Capture endpoint.
    #
    # High-level flow:
    #   1) Resolve optional playback routing target (Render endpoint) first.
    #      - --playback-target-id may be: None (don't change), "" (default device), or actual ID.
    #      - --playback-target-name overrides the ID (also supports flag-without-value -> const='' -> default device).
    #   2) Resolve the Capture device (active-only, GUI-order --index).
    #   3) Call set_listen_to_device_ps():
    #      - writes the checkbox state via IPropertyStore (COM),
    #      - and may write routing target via registry (HKLM; may require admin).
    #   4) Capture stderr and re-emit only non-error lines:
    #      COM/registry paths can emit warnings that are useful but noisy; we filter to
    #      preserve INFO lines while avoiding duplicate ERROR framing.
    #   5) Verification fallbacks:
    #      - read back via COM PropertyStore (exact),
    #      - if COM read is inconclusive, poll registry until it matches.
    #
    # "verifiedBy" meaning in JSON:
    #   If the setter returns a failure but we can confirm the state changed anyway,
    #   we still return success with verifiedBy="com" or "registry". This reflects
    #   real-world Windows behavior where the write may apply even when the call path
    #   signaled an error due to timing or driver quirks.
    #
    # Exit codes:
    #   0 success
    #   1 failed (and not verifiable)
    #   3 not found (active-only)
    #   4 ambiguous / --index out of range
    #
    # Resolve playback target. Start with ID if it was provided.
    render_device_id = args.playback_target_id
    # Handle --playback-target-name. This will override the ID if both are used.
    if args.playback_target_name is not None:
        # If the flag was used without a value, argparse sets it to const=''
        if args.playback_target_name == '':
            render_device_id = ''
        else:
            # A name was provided, so find the device ID (Render devices, active-only).
            render_target, err = _select_by_name_active_only("Render", args.playback_target_name, None, args.regex)
            if err:
                print(f"ERROR: Could not find playback target device: {err}", file=sys.stderr)
                return 3
            render_device_id = render_target["id"]

    # --- Discovery for Issue #34: Resolve target info ONLY if requested in command ---
    playback_target_name = None
    warning = None
    
    if render_device_id is not None:
        if render_device_id == '':
            playback_target_name = "Default Playback Device"
        else:
            # Resolve the friendly name for the provided ID
            all_renders = [d for d in list_devices(include_all=True) if d["flow"] == "Render"]
            match = next((d for d in all_renders if d["id"] == render_device_id), None)
            if match:
                playback_target_name = match["name"]
            else:
                # If ID is invalid, Windows uses default. We report fallback and add a warning.
                playback_target_name = "Default Playback Device (Fallback)"
                warning = f"Playback target ID '{render_device_id}' invalid; using default."

    # --- The rest of the function is for finding the CAPTURE device ---
    target, err = _resolve_standard_target(args, flow_forced="Capture")
    if err:
        print(err, file=sys.stderr)
        return 4 if "Multiple" in err or "index" in err else 3
        
    # --- Now, call the device function with the resolved render_device_id ---
    captured_stderr = io.StringIO()
    ok = False
    with redirect_stderr(captured_stderr):
        # render_device_id semantics:
        #   None => do not modify routing target
        #   ''   => route to "Default Playback Device"
        #   id   => route to that specific Render endpoint
        ok = set_listen_to_device_ps(target["id"], args.enable, render_device_id=render_device_id)
        
    stderr_output = captured_stderr.getvalue()

    # Function to build the JSON dictionary dynamically based on if a target was requested
    def build_payload(enabled_state, verified_by=None):
        payload = {"id": target["id"], "name": target["name"], "enabled": enabled_state}
        if verified_by:
            payload["verifiedBy"] = verified_by
        # Only add playback fields if they were part of the command
        if playback_target_name is not None:
            payload["playbackTargetId"] = render_device_id if render_device_id else "default"
            payload["playbackTargetName"] = playback_target_name
            if warning:
                payload["warning"] = warning
        return payload

    if not ok:
        # Even if the setter returned False, Windows/driver may have applied the change.
        # Verify via COM first (exact), then registry (robust) before calling it a failure.
        actual = _get_listen_to_device_status_ps(target["id"])
        if actual is not None and actual == args.enable:
            _reemit_non_error_stderr(stderr_output)
            print(json.dumps({"listenSet": build_payload(actual, "com")}))
            return 0
        verified, reg_state = _verify_listen_via_registry(target["id"], args.enable, timeout=3.0, interval=0.20)
        if verified or (reg_state is not None and reg_state == args.enable):
            _reemit_non_error_stderr(stderr_output)
            print(json.dumps({"listenSet": build_payload(reg_state, "registry")}))
            return 0
        sys.stderr.write(stderr_output)
        print(f"ERROR: failed to set 'Listen to this device' for '{target['name']}'.", file=sys.stderr)
        return 1
        
    # Setter returned success; still re-emit any useful INFO/WARN lines.
    _reemit_non_error_stderr(stderr_output)
    actual_enabled_state = _get_listen_to_device_status_ps(target["id"])
    if actual_enabled_state is None:
        verified, reg_state = _verify_listen_via_registry(target["id"], args.enable, timeout=3.0, interval=0.20)
        if verified or reg_state is not None:
            actual_enabled_state = reg_state
            
    print(json.dumps({"listenSet": build_payload(actual_enabled_state)}))
    return 0

def cmd_get_listen(args):
    target, err = _resolve_standard_target(args, flow_forced="Capture")
    if err:
        print(err, file=sys.stderr)
        return 4 if "Multiple" in err or "index" in err else 3

    state = _read_listen_enable_fast(target["id"])
    
    # Base JSON (Original structure)
    result = {
        "id": target["id"],
        "name": target["name"],
        "flow": target["flow"],
        "listenEnabled": state,
    }

    # Only add routing info if specifically requested in the get-listen command
    if args.playback_target_id is not None or args.playback_target_name is not None:
        import winreg
        from .devices import _extract_endpoint_guid_from_device_id
        
        current_target_id = "default"
        current_target_name = "Default Playback Device"
        
        guid = _extract_endpoint_guid_from_device_id(target["id"])
        if guid:
            key_path = rf"SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\Capture\{guid}\Properties"
            value_name = "{24dbb0fc-9311-4b3d-9cf0-18ff155639d4},0"
            try:
                key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key_path, 0, winreg.KEY_READ)
                val, _ = winreg.QueryValueEx(key, value_name)
                winreg.CloseKey(key)
                if val:
                    current_target_id = val
                    all_renders = [d for d in list_devices(include_all=True) if d["flow"] == "Render"]
                    match = next((d for d in all_renders if d["id"] == val), None)
                    if match:
                        current_target_name = match["name"]
            except OSError:
                pass
        
        result["playbackTargetId"] = current_target_id
        result["playbackTargetName"] = current_target_name

    print(json.dumps(result))
    return 0

def cmd_enhancements(args):
    # enhancements:
    #   Vendor-first control surface for:
    #     - Main "Audio Enhancements" on/off (SysFX) toggle (vendor-only at runtime)
    #     - FX operations: list/learn/enable/disable/delete learned per-effect toggles
    #
    # Constraint:
    #   Exactly one operation must be specified. This keeps the command unambiguous
    #   and predictable for GUI callers and scripts.
    #
    # Device selection:
    #   Active-only selection with consistent --index semantics (GUI-order within flow).
    #
    # Exit codes:
    #   0 success
    #   1 invalid args or toggle/learn failure
    #   3 not found
    #   4 ambiguous selection / out of range / user selection invalid (interactive disambiguation)
    #
    # Validation: exactly one operation
    ops = [
        bool(args.enable), bool(args.disable), bool(args.learn),
        bool(args.learn_fx), bool(args.enable_fx), bool(args.disable_fx), bool(args.list_fx),
        bool(getattr(args, "delete_fx", None)),
    ]
    if sum(ops) != 1:
        print("ERROR: specify exactly one of --enable, --disable, --learn, --learn-fx, --enable-fx, --disable-fx, or --list-fx", file=sys.stderr)
        return 1
    # Device selection (existing pattern)
    target, err = _resolve_standard_target(args)
    if err:
        print(err, file=sys.stderr)
        return 4 if "Multiple" in err or "index" in err else 3

    # === FX Operations ===
    # FX entries are learned/defined in vendor_toggles.ini and are per-device scoped.
    # They may be:
    #   - simple single-DWORD toggles, or
    #   - "multi-write" sequences (multiple registry values/types written together).
    if args.list_fx:
        # List FX available to this device as defined in vendor_toggles.ini.
        fx_list = _list_fx_for_device(
            target["id"], target["flow"],
            ini_path=getattr(args, "vendor_ini", None),
            device_name=target["name"],
        )
        fx_list = sorted(fx_list, key=lambda x: (x.get("fx_name") or "").lower())
        # JSON form is consumed by the GUI; includes per-FX state if readable.
        if getattr(args, "json", False):
            result = {
                "device": {"id": target["id"], "name": target["name"], "flow": target["flow"]},
                "availableFX": []
            }
            for fx in fx_list:
                entry = fx.get("entry")
                state = None
                try:
                    state = _read_vendor_entry_state(entry, target["id"], target["flow"])
                except Exception:
                    state = None
                result["availableFX"].append({
                    "fx_name": fx.get("fx_name"),
                    "state": state,
                    "source": "ini"
                })
            print(json.dumps(result))
            return 0
        # Human-readable listing (default CLI UX).
        print(f"Enhancement Effects for: {target['name']} ({target['flow']})")
        if not fx_list:
            print("  (none)")
            return 0
        for fx in fx_list:
            entry = fx.get("entry")
            try:
                st = _read_vendor_entry_state(entry, target["id"], target["flow"])
            except Exception:
                st = None
            state_txt = "Enabled" if st is True else "Disabled" if st is False else "Unknown"
            src = "ini"
            print(f"  - {fx.get('fx_name')}  [source={src}]  state={state_txt}")
        return 0
    if args.learn_fx:
        # learn-fx:
        #   Guided, interactive A/B snapshot capture for a specific effect.
        #
        # Two-pass model (important for real drivers):
        #   Many drivers create registry keys or stabilize state only after the first
        #   toggle. So we do:
        #     - Pass 1: A/B (prime/initialize)
        #     - Pass 2: A2/B2 (the authoritative pair we record in the INI)
        #
        # Timing:
        #   We allow a small settle delay before each snapshot to let the driver and
        #   property system commit changes. This is configurable via:
        #     AUDIOCTL_LEARN_FX_SETTLE (seconds)
        fx_name = args.learn_fx.strip()
        if not fx_name:
            print("ERROR: FX name cannot be empty", file=sys.stderr)
            return 1
        print(f"Learning FX '{fx_name}' for: {target['name']} ({target['flow']})")
        # Settling delay before each snapshot; override with AUDIOCTL_LEARN_FX_SETTLE (seconds)
        try:
            FX_SETTLE = float(os.environ.get("AUDIOCTL_LEARN_FX_SETTLE", "0.35"))
        except Exception:
            FX_SETTLE = 0.35
        # First pass (A/B) – used only to "prime" the driver
        print(f"Set the '{fx_name}' effect to ENABLED for this device.")
        input("When ready, press Enter to capture snapshot A... ")
        captured_stderr = io.StringIO()
        with redirect_stderr(captured_stderr):
            time.sleep(FX_SETTLE)
            snapA = _collect_sysfx_snapshot(target["id"])
        _reemit_non_error_stderr(captured_stderr.getvalue())
        print(f"Now set the '{fx_name}' effect to DISABLED for the same device.")
        input("When ready, press Enter to capture snapshot B... ")
        captured_stderr = io.StringIO()
        with redirect_stderr(captured_stderr):
            time.sleep(FX_SETTLE)
            snapB = _collect_sysfx_snapshot(target["id"])
        _reemit_non_error_stderr(captured_stderr.getvalue())
        # Second pass (A2/B2) – this pair is what we will record in the INI
        print(f"Enable the '{fx_name}' effect again (second pass).")
        input("When ready, press Enter to capture snapshot A2... ")
        captured_stderr = io.StringIO()
        with redirect_stderr(captured_stderr):
            time.sleep(FX_SETTLE)
            snapA2 = _collect_sysfx_snapshot(target["id"])
        _reemit_non_error_stderr(captured_stderr.getvalue())
        print(f"Disable the '{fx_name}' effect again (second pass).")
        input("When ready, press Enter to capture snapshot B2... ")
        captured_stderr = io.StringIO()
        with redirect_stderr(captured_stderr):
            time.sleep(FX_SETTLE)
            snapB2 = _collect_sysfx_snapshot(target["id"])
        _reemit_non_error_stderr(captured_stderr.getvalue())
        ok, info = _learn_fx_and_write_ini(
            target, fx_name, snapA, snapB,
            ini_path=getattr(args, "vendor_ini", None) or _vendor_ini_default_path(),
            prefer_hkcu=not args.prefer_hklm,
            snapA2=snapA2, snapB2=snapB2
        )
        if ok:
            print(json.dumps({"fxLearned": {
                "id": target["id"],
                "name": target["name"],
                "flow": target["flow"],
                "fx_name": fx_name,
                **info
            }}))
            return 0
        else:
            print(f"ERROR: FX learn failed: {info}", file=sys.stderr)
            return 1
    if args.enable_fx or args.disable_fx:
        # enable-fx/disable-fx:
        #   Toggle a learned effect for this device.
        #
        # Matching:
        #   - Default: case-insensitive substring match against fx_name.
        #   - With --regex: treat the provided FX selector as a regex.
        #
        # If multiple matches:
        #   - interactive disambiguation prompts user for an index on stdin.
        #
        # Output:
        #   {"fxSet":{"id","name","fx_name","enabled","verifiedBy"}}
        # verifiedBy indicates the vendor method used (INI entry type / multi-write decider).
        desired = (args.enable_fx or args.disable_fx or "").strip()
        if not desired:
            print("ERROR: FX name cannot be empty", file=sys.stderr)
            return 1
        fx_all = _list_fx_for_device(target["id"], target["flow"], ini_path=getattr(args, "vendor_ini", None))
        # Build matches by name using substring or regex, case-insensitive
        matches_fx = []
        if args.regex:
            try:
                pat = re.compile(desired, re.IGNORECASE)
            except re.error as e:
                print(f"ERROR: invalid regex for FX name: {e}", file=sys.stderr)
                return 1
            for fx in fx_all:
                if pat.search(fx.get("fx_name") or ""):
                    matches_fx.append(fx)
        else:
            for fx in fx_all:
                if desired.lower() in (fx.get("fx_name") or "").lower():
                    matches_fx.append(fx)
        if not matches_fx:
            print(f"ERROR: FX '{desired}' not found on this device. Use --list-fx to see available effects.", file=sys.stderr)
            return 1
        # If multiple matches, prompt user to choose
        if len(matches_fx) > 1:
            print("Multiple FX matches found:")
            for i, fx in enumerate(matches_fx):
                entry = fx.get("entry") or {}
                print(f"  [{i}] {fx.get('fx_name')}  [source=ini]")
            try:
                sel = input(f"Select index (0..{len(matches_fx)-1}): ").strip()
                idx = int(sel)
                if idx < 0 or idx >= len(matches_fx):
                    print("ERROR: selection out of range.", file=sys.stderr)
                    return 4
            except Exception:
                print("ERROR: invalid selection.", file=sys.stderr)
                return 4
            chosen_name = matches_fx[idx].get("fx_name") or desired
        else:
            chosen_name = matches_fx[0].get("fx_name") or desired
        from .vendor_db import _apply_fx
        enable = bool(args.enable_fx)
        ok, verified_by, state = _apply_fx(
            target["id"], target["flow"], chosen_name, enable,
            ini_path=getattr(args, "vendor_ini", None),
            device_name=target["name"],
        )
        if ok:
            print(json.dumps({
                "fxSet": {
                    "id": target["id"],
                    "name": target["name"],
                    "fx_name": chosen_name,
                    "enabled": state,
                    "verifiedBy": verified_by
                }
            }))
            return 0
        else:
            print(f"ERROR: FX '{chosen_name}' toggle failed for this device", file=sys.stderr)
            return 1
    if args.delete_fx:
        # delete-fx:
        #   Removes the association for this device GUID from the FX bucket in the INI.
        #
        # Matching/disambiguation follows the same rules as enable/disable-fx.
        #
        # Output:
        #   {"fxDeleted": {"id","name","fx_name", ...info...}}
        desired = (args.delete_fx or "").strip()
        if not desired:
            print("ERROR: FX name cannot be empty", file=sys.stderr)
            return 1
        fx_all = _list_fx_for_device(
            target["id"], target["flow"],
            ini_path=getattr(args, "vendor_ini", None),
            device_name=target["name"],
        )
        # Build matches by name using substring or regex, case-insensitive (same as enable/disable-fx)
        matches_fx = []
        if args.regex:
            try:
                pat = re.compile(desired, re.IGNORECASE)
            except re.error as e:
                print(f"ERROR: invalid regex for FX name: {e}", file=sys.stderr)
                return 1
            for fx in fx_all:
                if pat.search(fx.get("fx_name") or ""):
                    matches_fx.append(fx)
        else:
            for fx in fx_all:
                if desired.lower() in (fx.get("fx_name") or "").lower():
                    matches_fx.append(fx)
        if not matches_fx:
            print(f"ERROR: FX '{desired}' not found on this device. Use --list-fx to see available effects.", file=sys.stderr)
            return 1
        # If multiple matches, prompt user to choose (same style as enable-fx)
        if len(matches_fx) > 1:
            print("Multiple FX matches found:")
            for i, fx in enumerate(matches_fx):
                print(f"  [{i}] {fx.get('fx_name')}")
            try:
                sel = input(f"Select index (0..{len(matches_fx)-1}): ").strip()
                idx = int(sel)
                if idx < 0 or idx >= len(matches_fx):
                    print("ERROR: selection out of range.", file=sys.stderr)
                    return 4
            except Exception:
                print("ERROR: invalid selection.", file=sys.stderr)
                return 4
            chosen_name = matches_fx[idx].get("fx_name") or desired
        else:
            chosen_name = matches_fx[0].get("fx_name") or desired
        from .vendor_db import _delete_fx_for_guid
        ok, info = _delete_fx_for_guid(chosen_name, target["id"], ini_path=getattr(args, "vendor_ini", None))
        if ok:
            print(json.dumps({
                "fxDeleted": {
                    "id": target["id"],
                    "name": target["name"],
                    "fx_name": chosen_name,
                    **(info or {})
                }
            }))
            return 0
        else:
            msg = info if isinstance(info, str) else (info.get("reason") if isinstance(info, dict) else "unknown")
            print(f"ERROR: delete-fx failed: {msg}", file=sys.stderr)
            return 1
    # === Main toggle operations (vendor-only at runtime) ===
    if args.learn:
        # learn (main):
        #   Manual learning flow where the user toggles "Audio Enhancements" in the
        #   Windows UI while we capture registry snapshots.
        #
        # Retry logic:
        #   Some drivers create their vendor keys only after the first toggle. So if
        #   we cannot detect a clean DWORD flip, we:
        #     - check if a vendor entry is now available (vendorAvailable),
        #     - ask the user to toggle again and retry capture,
        #     - re-check after the retry.
        #
        # Output JSON:
        #   - {"vendorLearned": {...}} if we successfully wrote/updated an INI entry
        #   - {"vendorAvailable": {...}} if a usable vendor method exists but no new
        #     section was written by this learn attempt
        #
        # Exit codes:
        #   0 learn succeeded or vendor method available
        #   1 learn failed
        ok, info = _learn_vendor_from_discovery_and_write_ini(
            target,
            ini_path=getattr(args, "vendor_ini", None) or _vendor_ini_default_path(),
            prefer_hkcu=True
        )
        if ok:
            print(json.dumps({"vendorLearned": {"id": target["id"], "name": target["name"], "flow": target["flow"], **info}}, indent=2))
            return 0
        # No DWORD flip detected. Explain and check for vendor availability (INI or code fallback).
        print("INFO: No clean DWORD flip detected. Checking for vendor methods initialized by your toggle...", file=sys.stderr)
        vend_entry = _find_first_vendor_entry(target["id"], target["flow"], ini_path=getattr(args, "vendor_ini", None))
        if vend_entry:
            print(json.dumps({
                "vendorAvailable": {
                    "id": target["id"],
                    "name": target["name"],
                    "flow": target["flow"],
                    "vendor": vend_entry.get("name"),
                    "value_name": vend_entry.get("value_name"),
                    "note": "Device can be controlled via vendor method (INI)."
                }
            }, indent=2))
            return 0
        # Still not found. Guide user to try toggling again (second pass).
        print("INFO: This may be the first time this endpoint was toggled. The driver often creates keys only after the first toggle.", file=sys.stderr)
        print("INFO: Please toggle Enhancements again (Enable, then Disable) for this same device when prompted.", file=sys.stderr)
        # Second attempt
        ok2, info2 = _learn_vendor_from_discovery_and_write_ini(
            target,
            ini_path=getattr(args, "vendor_ini", None) or _vendor_ini_default_path(),
            prefer_hkcu=True
        )
        if ok2:
            print(json.dumps({"vendorLearned": {"id": target["id"], "name": target["name"], "flow": target["flow"], **info2}}, indent=2))
            return 0
        # Final vendor re-check after second user toggle
        vend_entry2 = _find_first_vendor_entry(target["id"], target["flow"], ini_path=getattr(args, "vendor_ini", None))
        if vend_entry2:
            print(json.dumps({
                "vendorAvailable": {
                    "id": target["id"],
                    "name": target["name"],
                    "flow": target["flow"],
                    "vendor": vend_entry2.get("name"),
                    "value_name": vend_entry2.get("value_name"),
                    "note": "Device can be controlled via vendor method (initialized by your toggles; INI)."
                }
            }, indent=2))
            return 0
        # Nothing found after retry
        print("ERROR: No DWORD flip found and no vendor method became available after a retry. Learn failed.", file=sys.stderr)
        return 1
    enable = True if args.enable else False
    # Runtime policy: vendor-only. If we don't have a learned/known vendor method,
    # we do not attempt Windows Disable_SysFx toggles here (those exist for diagnostics/learn).
    if not _enhancements_supported(target["id"], target["flow"]):
        print("ERROR: No vendor toggle available for this device. Use --learn to teach a vendor method.", file=sys.stderr)
        return 1
    ok, verified_by, state = _apply_enhancements(
        target["id"], target["flow"], enable,
        prefer_hklm=args.prefer_hklm,
        allow_universal_scan=False,
        vendor_ini_path=getattr(args, "vendor_ini", None)
    )
    if ok:
        print(json.dumps({"enhancementsSet": {"id": target["id"], "name": target["name"], "enabled": state, "verifiedBy": verified_by}}))
        return 0
    print("ERROR: vendor toggle failed.", file=sys.stderr)
    return 1

def cmd_get_enhancements(args):
    """
    Get current enhancements enabled/disabled state from vendor methods only.
    Returns JSON:
      {
        "id": "...",
        "name": "...",
        "flow": "Render"|"Capture",
        "enhancementsEnabled": true|false|null
      }
    """
    # get-enhancements:
    #   Vendor-only, fast query intended for GUI labels and scripts.
    #
    # Behavior note:
    #   If a vendor method exists for the device but the read is inconclusive,
    #   we "default-assume enabled" (True). This is primarily for UI friendliness:
    #   many drivers omit explicit disable markers when enabled, and treating "unknown"
    #   as disabled would make the UI feel wrong.
    #
    # Exit codes:
    #   0 success
    #   3 not found
    #   4 ambiguous / --index out of range
    target, err = _resolve_standard_target(args)
    if err:
        print(err, file=sys.stderr)
        return 4 if "Multiple" in err or "index" in err else 3

    # FAST vendor-only read: only if INI vendor is known for this device
    has_vendor = False
    try:
        has_vendor = _enhancements_supported(target["id"], target["flow"])
    except Exception:
        has_vendor = False
    if not has_vendor:
        state = None
    else:
        state = _fast_get_enhancements_state(target["id"], target["flow"])
        if state is None:
            # Default-assume enabled when indicators absent/inconclusive
            state = True
    print(json.dumps({
        "id": target["id"],
        "name": target["name"],
        "flow": target["flow"],
        "enhancementsEnabled": state,
    }))
    return 0

def cmd_get_device_state(args):
    """
    Return a combined view of device state for GUI:
      - volume & mute
      - listenEnabled (for capture)
      - enhancementsEnabled (vendor-only)
      - available FX with states
    JSON shape:
      {
        "id": "...",
        "name": "...",
        "flow": "Render"|"Capture",
        "volume": int or null,
        "muted": bool or null,
        "listenEnabled": bool or null,
        "enhancementsEnabled": bool or null,
        "availableFX": [
          { "fx_name": "...", "state": true|false|null, "source": "ini" }
        ]
      }
    """
    # get-device-state:
    #   One-stop, read-only call designed for the GUI so it doesn't need to call
    #   multiple commands for labels and menu state.
    #
    # Composition:
    #   - Volume/mute: COM read (accurate; no artificial sleeps).
    #   - Listen: fast registry probe (Capture only).
    #   - Enhancements: fast vendor INI-driven read (vendor-only).
    #   - FX list + per-FX states: INI-driven list + fast state reads.
    #
    # Defaults for UI friendliness:
    #   If vendor/FX state is inconclusive but an entry exists, we default-assume True
    #   (enabled). Many drivers store explicit "disabled" markers only.
    #
    # Exit codes:
    #   0 success
    #   3 not found
    #   4 ambiguous / --index out of range
    import time as _t
    target, err = _resolve_standard_target(args)
    if err:
        print(err, file=sys.stderr)
        return 4 if "Multiple" in err or "index" in err else 3

    dev_id = target["id"]
    flow   = target["flow"]
    # Volume & mute (no artificial sleeps, keep COM accuracy)
    vol = get_endpoint_volume(dev_id)
    muted = get_endpoint_mute(dev_id)
    if muted is not None:
        muted = bool(muted)
    # Listen (only meaningful for capture) – FAST registry probe
    listen_enabled = None
    if flow == "Capture":
        try:
            listen_enabled = _read_listen_enable_fast(dev_id)
        except Exception:
            listen_enabled = None
    # Enhancements (vendor-only) + FX using fast vendor_db helpers
    from .vendor_db import (
        _fast_get_enhancements_state,
        _list_fx_for_device,
        _fast_read_vendor_entry_state,
        _vendor_ini_default_path,
        _enhancements_supported,
    )
    # Determine effective INI path (for learned vendors/FX)
    ini_path = getattr(args, "vendor_ini", None)
    if not ini_path:
        try:
            ini_path = _vendor_ini_default_path()
        except Exception:
            ini_path = None
    # If no vendor toggle is known for this device yet, skip vendor lookups.
    has_vendor = False
    try:
        has_vendor = _enhancements_supported(dev_id, flow)
    except Exception:
        has_vendor = False
    # Enhancements (vendor-only, fast)
    if has_vendor:
        try:
            enh_enabled = _fast_get_enhancements_state(dev_id, flow)
        except Exception:
            enh_enabled = None
        if enh_enabled is None:
            # Default-assume enabled when indicator absent/inconclusive
            enh_enabled = True
    else:
        enh_enabled = None
    # FX list with fast states (ALWAYS list FX if present in INI; do not depend on MAIN support)
    available_fx = []
    try:
        fx_list = _list_fx_for_device(dev_id, flow, ini_path=ini_path, device_name=target["name"])
        fx_list = sorted(fx_list, key=lambda x: (x.get("fx_name") or "").lower())
        for fx in fx_list:
            entry = fx.get("entry")
            state = None
            try:
                state = _fast_read_vendor_entry_state(entry, dev_id, flow)
            except Exception:
                state = None
            if state is None:
                # Default-assume enabled when indicator absent/inconclusive
                state = True
            available_fx.append({
                "fx_name": fx.get("fx_name"),
                "state": state,
                "source": "ini",
            })
    except Exception:
        available_fx = []
    result = {
        "id": dev_id,
        "name": target["name"],
        "flow": flow,
        "volume": vol,
        "muted": muted,
        "listenEnabled": listen_enabled,
        "enhancementsEnabled": enh_enabled,
        "availableFX": available_fx,
    }
    print(json.dumps(result))
    return 0

def cmd_diag_sysfx(args):
    # diag-sysfx:
    #   Diagnostic command intended for troubleshooting and development.
    #   It compares:
    #     - live Windows enhancement state via PropertyStore (endpoint property)
    #     - live Windows enhancement state via PolicyConfig COM
    #     - learned vendor toggle state (if any INI entry exists)
    #
    # This helps explain mismatches between Windows-reported SysFX state and
    # vendor registry toggles a driver actually uses.
    #
    # Exit codes:
    #   0 success
    #   3 not found
    #   4 ambiguous / --index out of range
    target, err = _resolve_standard_target(args)
    if err:
        print(err, file=sys.stderr)
        return 4 if "Multiple" in err or "index" in err else 3

    live_win = _get_enhancements_status_propstore(target["id"])
    live_com = _get_enhancements_status_com(target["id"])
    vend_entry = _find_first_vendor_entry(target["id"], target["flow"], ini_path=None)
    vend_state = None
    vend_tag = "None Found"
    if vend_entry:
        vend_state = _read_vendor_entry_state(vend_entry, target["id"], target["flow"])
        vend_tag = f"{vend_entry['name']} ({vend_entry['value_name']})"
    print(json.dumps({
        "id": target["id"], "name": target["name"], "flow": target["flow"],
        "enhancementsEnabled_live_propstore": live_win,
        "enhancementsEnabled_live_com": live_com,
        "vendor_toggle_status": {vend_tag or "None Found": vend_state}
    }, indent=2))
    return 0

def cmd_discover_enhancements(args):
    # discover-enhancements:
    #   Interactive diagnostic/discovery workflow (manual UI toggling):
    #     - user sets Enhancements enabled -> snapshot A
    #     - user sets Enhancements disabled -> snapshot B
    #     - tool diffs registry values under MMDevices (HKCU/HKLM, FxProperties/Properties)
    #     - generates a human-readable report and a JSON bundle for debugging/sharing
    #     - optionally appends a suggested vendor INI snippet
    #
    # Output:
    #   - prints report text
    #   - writes TXT and JSON to output-dir (default cwd)
    #
    # Exit codes:
    #   0 success
    #   3 not found
    #   4 ambiguous / --index out of range
    target, err = _resolve_standard_target(args)
    if err:
        print(err, file=sys.stderr)
        return 4 if "Multiple" in err or "index" in err else 3

    print(f"Discovery target: {target['name']} [{target['id']}] ({target['flow']})")
    print("Step 1: In Windows Sound settings, set 'Audio Enhancements' to ENABLED for this device.")
    input("When ready, press Enter to capture snapshot A... ")
    snapA = _collect_sysfx_snapshot(target["id"])
    print("Step 2: Now set 'Audio Enhancements' to DISABLED for the same device.")
    input("When ready, press Enter to capture snapshot B... ")
    snapB = _collect_sysfx_snapshot(target["id"])
    diffs = _diff_mmdevices_lists(snapA.get("registry") or [], snapB.get("registry") or [])
    base_name = re.sub(r'[^A-Za-z0-9_.-]+', "_", f"enh-discovery_{target['flow']}_{target['name']}")
    from datetime import datetime
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_dir = args.output_dir or os.getcwd()
    try:
        os.makedirs(out_dir, exist_ok=True)
    except Exception:
        pass
    txt_path = os.path.join(out_dir, f"{base_name}_{stamp}.txt")
    json_path = os.path.join(out_dir, f"{base_name}_{stamp}.json")
    report_text = _generate_enh_discovery_report(target, snapA, snapB, diffs)
    try:
        with open(txt_path, "w", encoding="utf-8", errors="replace") as f:
            f.write(report_text)
    except Exception as e:
        print(f"ERROR: failed to write report: {e}", file=sys.stderr)
    bundle = {
        "device": target,
        "snapshotA": snapA,
        "snapshotB": snapB,
        "diffs": diffs,
    }
    try:
        with open(json_path, "w", encoding="utf-8", errors="replace") as f:
            json.dump(bundle, f, indent=2)
    except Exception as e:
        print(f"ERROR: failed to write JSON bundle: {e}", file=sys.stderr)
    if getattr(args, "ini_snippet", None):
        # Best-effort INI snippet generation for simple DWORD flips; meant as a hint
        # for manual integration into vendor_toggles.ini.
        snippet, picked = _build_vendor_ini_snippet(target, snapA, snapB, diffs)
        if snippet:
            try:
                with open(args.ini_snippet, "a", encoding="utf-8", errors="replace") as f:
                    f.write("\n" + snippet)
                print(f"\nSuggested vendor INI section appended to: {args.ini_snippet}")
                print("Snippet:\n" + snippet)
            except Exception as e:
                print(f"ERROR: failed to write INI snippet: {e}", file=sys.stderr)
        else:
            print("No suitable DWORD flip candidate found for INI snippet.", file=sys.stderr)
    print(report_text)
    print(f"\nSaved:")
    print(f"  TXT  -> {txt_path}")
    print(f"  JSON -> {json_path}")
    return 0

def cmd_wait(args):
    # wait:
    #   Poll for a device to appear as an ACTIVE endpoint until timeout.
    #   This is used by automation scripts that need to wait for a USB device,
    #   headset docking, driver restart, etc.
    #
    # Behavior:
    #   - checks only active endpoints (consistent with all mutating commands)
    #   - sleeps 0.5s between polls
    #
    # Exit codes:
    #   0 found (prints {"found": <device-dict>})
    #   3 timeout
    #   4 ambiguous / --index out of range
    deadline = time.time() + args.timeout
    while time.time() < deadline:
        devices = list_devices(include_all=False)
        _sort_and_tag_gui_indices(devices) # Ensure indices match 'list'
        matches = find_devices_by_selector(devices, dev_id=args.id, name_substr=args.name, flow=args.flow, regex=args.regex)
        if matches:
            if args.index is not None:
                # Check for exact index match
                target = next((d for d in matches if d.get("guiIndex") == args.index), None)
                if not target:
                    # If index not found, keep waiting? Or fail? 
                    # Given 'wait' semantics, we should wait until the specific index appears.
                    time.sleep(0.5)
                    continue
            else:
                target = matches[0]

            print(json.dumps({"found": target}))
            return 0
        time.sleep(0.5)
    print("ERROR: timeout waiting for device", file=sys.stderr)
    return 3

def build_parser():
    # Build the argparse tree.
    #
    # The GUI depends on these commands and their JSON shapes remaining stable,
    # because the GUI treats the CLI as its "backend".
    #
    # Subcommand groups:
    #   - device listing and selection helpers (list, wait)
    #   - default endpoint control (set-default)
    #   - endpoint volume/mute (set-volume, get-volume)
    #   - Listen checkbox control (listen, get-listen)
    #   - Enhancements vendor control + FX subsystem (enhancements, get-enhancements, get-device-state)
    #   - diagnostics and discovery (diag-sysfx, diag-mmdevices, discover-enhancements)
    #   - internal helper for elevation (vendor-ini-append)
    p = argparse.ArgumentParser(prog="audioctl", description="Windows audio control CLI (pycaw-based)")
    sub = p.add_subparsers(dest="cmd", required=True)
    # --- list ---
    p_list = sub.add_parser("list", help="List devices")
    p_list.add_argument("--all", action="store_true", help="Include disabled/disconnected")
    p_list.add_argument("--json", action="store_true")
    p_list.set_defaults(func=cmd_list)
    # --- set-default ---
    p_sd = sub.add_parser("set-default", help="Set default playback/recording devices (Admin might be required)")
    p_sd.add_argument("--playback-id")
    p_sd.add_argument("--playback-name")
    p_sd.add_argument("--playback-role", choices=list(ROLES.keys()), default="all")
    p_sd.add_argument("--playback-flow", choices=["Render"], help=argparse.SUPPRESS)
    p_sd.add_argument("--recording-id")
    p_sd.add_argument("--recording-name")
    p_sd.add_argument("--recording-role", choices=list(ROLES.keys()), default="communications")
    p_sd.add_argument("--recording-flow", choices=["Capture"], help=argparse.SUPPRESS)
    p_sd.add_argument("--index", type=int)
    p_sd.add_argument("--regex", action="store_true")
    p_sd.set_defaults(func=cmd_set_default)
    # --- volume/mute ---
    p_sv = sub.add_parser("set-volume", help="Set endpoint volume (render or capture) or mute/unmute")
    p_sv.add_argument("--id")
    p_sv.add_argument("--name")
    p_sv.add_argument("--flow", choices=["Render", "Capture"], help="Optional filter to disambiguate")
    p_sv.add_argument("--level", type=int, help="0-100 (for volume)")
    p_sv.add_argument("--mute", action="store_true", help="Mute the device")
    p_sv.add_argument("--unmute", action="store_true", help="Unmute the device")
    p_sv.add_argument("--index", type=int)
    p_sv.add_argument("--regex", action="store_true")
    p_sv.add_argument("--json", action="store_true", help="Output JSON on success (default) and minimized text on error")
    p_sv.set_defaults(func=cmd_set_volume)
    p_gv = sub.add_parser("get-volume", help="Get endpoint volume and mute state (render or capture)")
    p_gv.add_argument("--id")
    p_gv.add_argument("--name")
    p_gv.add_argument("--flow", choices=["Render", "Capture"], help="Optional filter to disambiguate")
    p_gv.add_argument("--index", type=int)
    p_gv.add_argument("--regex", action="store_true")
    p_gv.set_defaults(func=cmd_get_volume)
    # --- listen ---
    p_ls = sub.add_parser("listen", help="Enable/disable 'Listen to this device' (capture only)")
    p_ls.add_argument("--id", help="Device ID for the capture device.")
    p_ls.add_argument("--name", help="Substring of the device name for the capture device.")
    p_ls.add_argument("--enable", action="store_true", help="Enable 'Listen to this device'.")
    p_ls.add_argument("--disable", action="store_true", help="Disable 'Listen to this device'.")
    p_ls.add_argument("--playback-target-id", nargs='?', const='', default=None, help="Optional: Render endpoint ID to play through. Use without a value for 'Default Playback Device'.")
    p_ls.add_argument("--playback-target-name", nargs='?', const='', default=None, help="Optional: Render endpoint name to play through. Use without a value for 'Default Playback Device'.")
    p_ls.add_argument("--index", type=int)
    p_ls.add_argument("--regex", action="store_true")
    p_ls.add_argument("--json", action="store_true", help="Output JSON on success (default) and minimized text on error")
    p_ls.set_defaults(func=cmd_listen)
    p_gl = sub.add_parser("get-listen", help="Get 'Listen to this device' enabled/disabled state for a capture device")
    p_gl.add_argument("--id", help="Device ID for the capture device.")
    p_gl.add_argument("--name", help="Substring or regex of the device name")
    p_gl.add_argument("--index", type=int)
    p_gl.add_argument("--regex", action="store_true")
    p_gl.add_argument("--playback-target-id", nargs='?', const='', help="Request current playback routing target ID")
    p_gl.add_argument("--playback-target-name", nargs='?', const='', help="Request current playback routing target name")
    p_gl.set_defaults(func=cmd_get_listen)
    # --- enhancements / FX subsystem ---
    p_fx = sub.add_parser("enhancements", help="Enable/disable 'Audio Enhancements' (SysFX) on a device")
    p_fx.add_argument("--id", help="Endpoint ID")
    p_fx.add_argument("--name", help="Substring or regex of the endpoint name")
    p_fx.add_argument("--flow", choices=["Render", "Capture"], help="Optional filter to disambiguate")
    p_fx.add_argument("--enable", action="store_true", help="Enable audio enhancements")
    p_fx.add_argument("--disable", action="store_true", help="Disable audio enhancements")
    p_fx.add_argument("--index", type=int, help="GUI-order index among matches")
    p_fx.add_argument("--regex", action="store_true")
    p_fx.add_argument("--prefer-hklm", action="store_true",
                      help="For learned INI toggles that write registry values, prefer HKLM ordering (Admin typically required).")
    p_fx.add_argument("--vendor-ini", help="Path to vendor_toggles.ini (default: next to the EXE).")
    p_fx.add_argument("--learn", action="store_true",
                      help="Manual learn (you toggle Windows UI). Captures A/B and writes vendor INI; no Windows fallback is used at runtime.")
    # NEW FX ARGUMENTS
    p_fx.add_argument("--learn-fx", metavar="FX_NAME",
                      help="Learn a specific audio effect (e.g., BassBoost, Loudness)")
    p_fx.add_argument("--enable-fx", metavar="FX_NAME",
                      help="Enable a learned audio effect")
    p_fx.add_argument("--disable-fx", metavar="FX_NAME",
                      help="Disable a learned audio effect")
    p_fx.add_argument("--list-fx", action="store_true",
                      help="List all learned effects for this device")
    p_fx.add_argument("--json", action="store_true",
                      help="For --list-fx: output JSON instead of human-readable text")
    p_fx.add_argument("--delete-fx", metavar="FX_NAME",
                      help="Delete a learned audio effect for this device")
    p_fx.set_defaults(func=cmd_enhancements)
    # --- vendor-only state queries ---
    p_ge = sub.add_parser(
        "get-enhancements",
        help="Get current enhancements enabled/disabled state for a device (vendor-only)"
    )
    p_ge.add_argument("--id")
    p_ge.add_argument("--name")
    p_ge.add_argument("--flow", choices=["Render", "Capture"])
    p_ge.add_argument("--index", type=int)
    p_ge.add_argument("--regex", action="store_true")
    p_ge.set_defaults(func=cmd_get_enhancements)
    p_gds = sub.add_parser(
        "get-device-state",
        help="Get current state for a device (volume, mute, listen, enhancements, FX) for GUI"
    )
    p_gds.add_argument("--id")
    p_gds.add_argument("--name")
    p_gds.add_argument("--flow", choices=["Render", "Capture"])
    p_gds.add_argument("--index", type=int)
    p_gds.add_argument("--regex", action="store_true")
    p_gds.add_argument("--vendor-ini", help="Optional vendor INI path for FX lookup")
    p_gds.set_defaults(func=cmd_get_device_state)
    # --- diagnostics ---
    p_dx = sub.add_parser(
        "diag-sysfx",
        help="Dump live Enhancements state (COM, PropertyStore, vendor toggles)"
    )
    p_dx.add_argument("--id")
    p_dx.add_argument("--name")
    p_dx.add_argument("--flow", choices=["Render", "Capture"])
    p_dx.add_argument("--index", type=int)
    p_dx.add_argument("--regex", action="store_true")
    p_dx.set_defaults(func=cmd_diag_sysfx)
    p_dm = sub.add_parser("diag-mmdevices", help="Dump all MMDevices values for an endpoint (debug)")
    p_dm.add_argument("--id")
    p_dm.add_argument("--name")
    p_dm.add_argument("--flow", choices=["Render", "Capture"])
    p_dm.add_argument("--index", type=int)
    p_dm.add_argument("--regex", action="store_true")
    p_dm.set_defaults(func=cmd_diag_mmdevices)
    # --- discovery (interactive report generation) ---
    p_learn = sub.add_parser("discover-enhancements", help="Interactively learn how Enhancements toggles for a device")
    p_learn.add_argument("--id")
    p_learn.add_argument("--name")
    p_learn.add_argument("--flow", choices=["Render", "Capture"])
    p_learn.add_argument("--index", type=int)
    p_learn.add_argument("--regex", action="store_true")
    p_learn.add_argument("--output-dir", help="Where to write the TXT/JSON report (default: current directory)")
    p_learn.add_argument("--ini-snippet", help="Write a suggested vendor INI section to this path (append).")
    p_learn.set_defaults(func=cmd_discover_enhancements)
    # --- wait ---
    p_w = sub.add_parser("wait", help="Wait for device to appear")
    p_w.add_argument("--id")
    p_w.add_argument("--name")
    p_w.add_argument("--flow", choices=["Render", "Capture"])
    p_w.add_argument("--timeout", type=int, default=30)
    p_w.add_argument("--index", type=int)
    p_w.add_argument("--regex", action="store_true")
    p_w.set_defaults(func=cmd_wait)
    # Hidden helper: elevated INI append (used by GUI to write into Program Files)
    # The GUI can write a "work order" JSON file and then run this subcommand via
    # an elevated process to perform the actual append when permissions require it.
    p_vi = sub.add_parser("vendor-ini-append", help=argparse.SUPPRESS)
    p_vi.add_argument("--work", required=True, help=argparse.SUPPRESS)  # path to JSON work order
    p_vi.set_defaults(func=cmd_vendor_ini_append)
    return p

def cmd_diag_mmdevices(args):
    # diag-mmdevices:
    #   Debug dump of all MMDevices registry values for an endpoint.
    #
    # This is primarily used for learn/discovery troubleshooting:
    # it reveals the raw keys/values that changed (or didn't) so we can
    # determine how a driver encodes "enhancements" and FX toggles.
    #
    # Exit codes:
    #   0 success
    #   3 not found
    #   4 ambiguous / --index out of range
    target, err = _resolve_standard_target(args)
    if err:
        print(err, file=sys.stderr)
        return 4 if "Multiple" in err or "index" in err else 3

    dump = _dump_mmdevices_all_values(target["id"])
    print(json.dumps({"id": target["id"], "name": target["name"], "flow": target["flow"], "mmdevices": dump}, indent=2))
    return 0

def cmd_vendor_ini_append(args):
    # vendor-ini-append (hidden):
    #   Internal helper used by the GUI to write/append learned INI entries when
    #   the INI path is in a protected location (e.g., Program Files).
    #
    # The GUI writes a "work order" JSON file with the intended operation, then
    # launches this subcommand elevated. This keeps the normal GUI process
    # non-elevated while still supporting admin-required writes.
    #
    # Supported work kinds:
    #   - "main": append a main Enhancements vendor toggle section
    #   - "fx":   append an FX section
    #
    # Error handling:
    #   - PermissionError: user didn't elevate or path is protected
    #   - FileExistsError: treated as a benign "already exists" result
    #
    # Exit codes:
    #   0 success (including "exists")
    #   1 failure (bad work order / write failure)
    try:
        with open(args.work, "r", encoding="utf-8") as f:
            work = json.load(f)
    except Exception as e:
        print(f"ERROR: failed to read work file: {e}", file=sys.stderr)
        return 1
    kind = (work.get("kind") or "").lower()
    ini_path = work.get("ini_path")
    if not ini_path or not kind:
        print("ERROR: invalid work order (missing kind or ini_path)", file=sys.stderr)
        return 1
    try:
        if kind == "main":
            from .vendor_db import _append_vendor_ini_entry_if_missing
            section = work["section"]
            value_name = work["value_name"]
            dword_enable = int(work["dword_enable"])
            dword_disable = int(work["dword_disable"])
            flows = work.get("flows", "Render,Capture")
            hives = work.get("hives", "HKCU,HKLM")
            notes = work.get("notes", "")
            res = _append_vendor_ini_entry_if_missing(
                ini_path, section, value_name,
                dword_enable, dword_disable,
                flows=flows, hives=hives, notes=notes
            )
            print(json.dumps({"iniAppend": {"kind": "main", "result": res, "iniPath": ini_path, "section": section}}))
            return 0
        elif kind == "fx":
            from .vendor_db import _append_fx_ini_entry
            section = work["section"]
            fx_name = work["fx_name"]
            device_name = work["device_name"]
            value_name = work["value_name"]
            dword_enable = int(work["dword_enable"])
            dword_disable = int(work["dword_disable"])
            flows = work.get("flows", "Render,Capture")
            hives = work.get("hives", "HKCU,HKLM")
            notes = work.get("notes", "")
            _append_fx_ini_entry(
                ini_path, section, fx_name, device_name,
                value_name, dword_enable, dword_disable,
                flows=flows, hives=hives, notes=notes
            )
            print(json.dumps({"iniAppend": {"kind": "fx", "result": "appended", "iniPath": ini_path, "section": section}}))
            return 0
        else:
            print("ERROR: unknown work kind (expected 'main' or 'fx')", file=sys.stderr)
            return 1
    except PermissionError as e:
        print(f"ERROR: permission denied writing INI: {e}", file=sys.stderr)
        return 1
    except FileExistsError as e:
        print(json.dumps({"iniAppend": {"kind": kind, "result": "exists", "iniPath": ini_path}}))
        return 0
    except Exception as e:
        print(f"ERROR: failed to append INI: {e}", file=sys.stderr)
        return 1

def main(argv=None):
    # Entry point for CLI usage and for "double-click / no args" GUI launching.
    #
    # If invoked with no args (and argv is None), we try to start the GUI.
    # The import is lazy so that:
    #   - CLI usage does not require tkinter,
    #   - and import-time side effects in the GUI module don't affect CLI runs.
    if argv is None and len(sys.argv) <= 1:
        try:
            from .gui import launch_gui  # Lazy import only if we actually need the GUI
            return launch_gui()
        except Exception as e:
            print(f"ERROR: GUI failed to start: {e}", file=sys.stderr)
    # NOTE: Do NOT call CoInitialize/CoUninitialize here anymore.
    # Each low-level helper in devices.py performs its own COM init/cleanup.
    # This avoids COM apartment lifetime bugs when multiple commands run inside
    # one Python process (and makes the CLI safe when called repeatedly by the GUI).
    parser = build_parser()
    args = parser.parse_args(argv)
    # Centralized validation for mutually exclusive flag sets.
    # (We duplicate some checks inside commands too so direct calls remain safe.)
    if args.cmd == "listen":
        if args.enable and args.disable:
            print("ERROR: specify only one of --enable or --disable", file=sys.stderr)
            return 1
        if not args.enable and not args.disable:
            print("ERROR: specify --enable or --disable", file=sys.stderr)
            return 1
        args.enable = True if args.enable else False
    if args.cmd == "enhancements":
        ops_count = (
            int(bool(getattr(args, "enable", False))) +
            int(bool(getattr(args, "disable", False))) +
            int(bool(getattr(args, "learn", False))) +
            int(bool(getattr(args, "learn_fx", None))) +
            int(bool(getattr(args, "enable_fx", None))) +
            int(bool(getattr(args, "disable_fx", None))) +
            int(bool(getattr(args, "list_fx", False))) +
            int(bool(getattr(args, "delete_fx", None)))
        )
        if ops_count != 1:
            print("ERROR: specify exactly one of --enable, --disable, --learn, --learn-fx, --enable-fx, --disable-fx, or --list-fx", file=sys.stderr)
            return 1
    if args.cmd == "set-volume":
        if (args.mute or args.unmute) and args.level is not None:
            print("ERROR: Cannot specify both --level and --mute/--unmute", file=sys.stderr)
            return 1
        if not (args.mute or args.unmute or args.level is not None):
            print("ERROR: Must specify --level or --mute/--unmute", file=sys.stderr)
            return 1
    try:
        rc = args.func(args)
    except KeyboardInterrupt:
        rc = 130
    return rc
