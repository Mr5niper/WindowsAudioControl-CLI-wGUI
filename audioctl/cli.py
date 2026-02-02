# audioctl/cli.py
#
# This module defines the *public* CLI surface for audioctl:
# - Argument parsing (subcommands + flags)
# - Command routing (cmd_* functions)
# - Stable, script/GUI-friendly JSON outputs and exit codes
#
# Design note (important for stability):
# The Tkinter GUI intentionally talks to audioctl by launching these CLI commands
# as subprocesses instead of importing low-level COM/registry helpers directly.
# That separation keeps GUI state isolated from COM lifetime/GC finalizers, and it
# also guarantees that the GUI and scripts observe the exact same behavior and
# JSON contract.
#
# Compatibility note:
# compat.py MUST be imported before any comtypes/pycaw usage. It patches
# comtypes.automation symbols/VT constants in environments where PyInstaller or
# certain comtypes versions expose slightly different names. The devices/vendor
# modules depend on those shims being in place early.

# --- stdlib imports (CLI plumbing, JSON I/O, subprocess-safe helpers) ---
import sys
import argparse
import json
import time
import io
import os
import re
from contextlib import redirect_stderr

# --- local modules (import compat BEFORE any comtypes usage downstream) ---
from .compat import (
    ROLES, is_admin,
)
from .logging_setup import _log, _log_exc

# --- low-level device + vendor helpers ---
# Keep the CLI thin: all COM / registry / vtable details live in devices.py and vendor_db.py.
# We import specific helpers to keep dependencies explicit and outputs consistent.
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


def cmd_list(args):
    """
    List audio endpoints.

    Selection rules:
      - Active-only by default.
      - --all includes disabled/disconnected devices (DEVICE_STATE_ALL).

    Output:
      - Default (human): grouped Render/Capture sections, with GUI-order indices.
      - --json: {"devices": [...]} where each device dict includes flow/state/defaults.

    Exit codes:
      - 0 on success (listing itself should not fail under normal conditions).
    """
    devices = list_devices(include_all=args.all)

    # We compute "GUI-order" indices by sorting names within each flow (Render/Capture).
    # This makes --index stable and consistent between CLI and GUI.
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
    """
    Set default playback and/or recording endpoints via PolicyConfig.

    Selection rules:
      - Targets are active-only (safety: avoids setting defaults to disabled/unplugged endpoints).
      - Selection is either:
          * exact by --playback-id/--recording-id, or
          * by name substring/regex, disambiguated by --index in GUI order within the flow.

    Roles:
      - console/multimedia/communications correspond to Windows default endpoint roles.
      - "all" sets all three roles.
      - Defaults are intentionally different:
          * playback_role defaults to "all" (typical expectation)
          * recording_role defaults to "communications" (common telephony behavior)

    Output:
      - Always JSON on stdout: {"set": [{"flow","role","id","name"}, ...]}

    Exit codes:
      - 0 success
      - 1 runtime failure while setting defaults
      - 3 not found (active-only)
      - 4 multiple matches (need --index) / disambiguation errors
    """
    # Admin isn't strictly required everywhere, but on some systems PolicyConfig calls
    # (or their downstream effects) may be restricted, so we warn instead of hard failing.
    if not is_admin():
        print("WARNING: 'set-default' might require Administrator privileges on this system.", file=sys.stderr)

    exit_code = 0
    results = {"set": []}

    # Playback selection: by ID (exact) or by name (active-only, GUI order for --index).
    if args.playback_id or args.playback_name:
        flow_name = "Render"
        if args.playback_id:
            matches = find_devices_by_selector(list_devices(include_all=False), dev_id=args.playback_id, flow=flow_name, regex=args.regex)
            if len(matches) == 0:
                print("ERROR: playback device not found (active only)", file=sys.stderr)
                return 3
            target = matches[0]
        else:
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

    # Recording selection: same rules as playback, but Capture flow.
    if args.recording_id or args.recording_name:
        flow_name = "Capture"
        if args.recording_id:
            matches = find_devices_by_selector(list_devices(include_all=False), dev_id=args.recording_id, flow=flow_name, regex=args.regex)
            if len(matches) == 0:
                print("ERROR: recording device not found (active only)", file=sys.stderr)
                return 3
            target = matches[0]
        else:
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
    """
    Set endpoint master volume and/or mute state (active endpoints only).

    Selection rules:
      - Active-only endpoints (safety + avoids surprising behavior on disabled endpoints).
      - Device is selected by:
          * --id (exact) OR
          * --name (substring/regex), optionally narrowed by --flow.
      - If multiple matches remain, caller must provide --index.
      - --index is interpreted in GUI order within the flow (name-sorted), matching the GUI.

    Operation rules:
      - Exactly one operation must be requested:
          * --level 0..100 OR
          * --mute OR
          * --unmute

    Output:
      - Success: JSON:
          {"volumeSet": {...}} or {"muteSet": {...}}
      - Errors: stderr message and non-zero exit code.

    Exit codes:
      - 0 success
      - 1 invalid args or operation failure
      - 3 not found
      - 4 multiple matches / index out of range
    """
    # Mutually exclusive operations guard: prevents ambiguous changes.
    if (args.mute or args.unmute) and args.level is not None:
        print("ERROR: Cannot specify both --level and --mute/--unmute", file=sys.stderr)
        return 1
    if not (args.mute or args.unmute or args.level is not None):
        print("ERROR: Must specify --level or --mute/--unmute", file=sys.stderr)
        return 1

    devices = list_devices(include_all=False)
    matches = find_devices_by_selector(devices, dev_id=args.id, name_substr=args.name, flow=args.flow, regex=args.regex)
    if len(matches) == 0:
        print("ERROR: device not found (active only)", file=sys.stderr)
        return 3
    if len(matches) > 1 and args.index is None:
        print("ERROR: multiple matches; specify --index", file=sys.stderr)
        return 4

    # Cross-cutting convention:
    # We sort/tag matches in the same way the GUI displays them so that --index is stable
    # across CLI usage, GUI behavior, and any scripts that rely on indices.
    buckets = _sort_and_tag_gui_indices([d for d in matches])
    flow = args.flow or (matches[0]["flow"] if matches else None)
    ordered = (buckets.get(flow) or []) if args.flow else (buckets["Render"] + buckets["Capture"])

    if not ordered:
        print("ERROR: no target device found for the specified criteria", file=sys.stderr)
        return 4

    if args.index is not None:
        if args.index < 0 or args.index >= len(ordered):
            print(f"ERROR: --index out of range (0..{len(ordered)-1})", file=sys.stderr)
            return 4
        target = ordered[args.index]
    else:
        target = ordered[0]

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
    devices = list_devices(include_all=False)
    matches = find_devices_by_selector(
        devices,
        dev_id=args.id,
        name_substr=args.name,
        flow=args.flow,
        regex=args.regex,
    )
    if len(matches) == 0:
        print("ERROR: device not found (active only)", file=sys.stderr)
        return 3
    if len(matches) > 1 and args.index is None:
        print("ERROR: multiple matches; specify --index", file=sys.stderr)
        return 4

    # --index is interpreted in GUI order within a flow (see cmd_set_volume for rationale).
    buckets = _sort_and_tag_gui_indices([d for d in matches])
    flow = args.flow or (matches[0]["flow"] if matches else None)
    ordered = (buckets.get(flow) or []) if args.flow else (buckets["Render"] + buckets["Capture"])
    if not ordered:
        print("ERROR: no target device found for the specified criteria", file=sys.stderr)
        return 4

    if args.index is not None:
        if args.index < 0 or args.index >= len(ordered):
            print(f"ERROR: --index out of range (0..{len(ordered)-1})", file=sys.stderr)
            return 4
        target = ordered[args.index]
    else:
        target = ordered[0]

    vol = get_endpoint_volume(target["id"])
    muted = get_endpoint_mute(target["id"])

    # Normalize muted to a plain bool/null-like so odd COM return types don't leak into JSON.
    if muted is not None:
        muted = bool(muted)

    # Contract used by GUI and scripts.
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
    """
    Enable/disable "Listen to this device" for a Capture endpoint.

    Selection rules:
      - Capture devices only, active-only.
      - Select by --id or --name (substring/regex), disambiguate via --index in GUI order.

    Playback routing:
      - Optional render_device_id (routing target) can be provided by ID or name.
      - Special values:
          * None  => do not modify routing target (preserve current selection)
          * ''    => set routing target to "Default Playback Device" (Windows UI behavior)
          * '<id>'=> set routing target to a specific Render endpoint ID

    Implementation/verification notes:
      - Setter uses COM PropertyStore for the enable flag and registry write for the target.
      - We capture stderr from the setter and re-emit only non-error lines. This keeps normal
        output (JSON) clean while still surfacing informative diagnostic warnings.
      - If the setter reports failure, we may still return success if read-back shows Windows
        actually applied the change (common with timing/COM HRESULT oddities).

    Output:
      - Success JSON: {"listenSet": {"id","name","enabled", ["verifiedBy"]}}
        verifiedBy indicates which read-back path confirmed the state ("com" or "registry").

    Exit codes:
      - 0 success
      - 1 failed to set/verify
      - 3 not found
      - 4 multiple matches / index errors
    """
    # Resolve playback target first. We do this up front so capture selection and mutation
    # can be done once with the correct routing intent.
    render_device_id = args.playback_target_id

    # Handle --playback-target-name. This overrides ID if both were provided.
    # argparse uses const='' when the flag is present without a value; we treat that as
    # "Default Playback Device" (Windows convention uses an empty string for default).
    if args.playback_target_name is not None:
        if args.playback_target_name == '':
            render_device_id = ''
        else:
            render_target, err = _select_by_name_active_only("Render", args.playback_target_name, None, args.regex)
            if err:
                print(f"ERROR: Could not find playback target device: {err}", file=sys.stderr)
                return 3
            render_device_id = render_target["id"]

    # --- Find the CAPTURE device (active-only) ---
    devices = list_devices(include_all=False)
    matches = find_devices_by_selector(devices, dev_id=args.id, name_substr=args.name, flow="Capture", regex=args.regex)
    if len(matches) == 0:
        print("ERROR: capture device not found (active only)", file=sys.stderr)
        return 3
    if len(matches) > 1 and args.index is None:
        print("ERROR: multiple matches; specify --index", file=sys.stderr)
        return 4

    # Interpret --index in GUI order within Capture flow.
    buckets = _sort_and_tag_gui_indices(matches[:])
    ordered = buckets["Capture"]
    if not ordered:
        print("ERROR: no target device found for the specified criteria", file=sys.stderr)
        return 4

    if args.index is not None:
        if args.index < 0 or args.index >= len(ordered):
            print(f"ERROR: --index out of range (0..{len(ordered)-1})", file=sys.stderr)
            return 4
        target = ordered[args.index]
    else:
        target = ordered[0]

    # Call the low-level setter. It handles COM init/cleanup internally.
    # We capture stderr because COM/property-store operations can emit warnings that
    # shouldn't corrupt JSON output.
    captured_stderr = io.StringIO()
    ok = False
    with redirect_stderr(captured_stderr):
        ok = set_listen_to_device_ps(target["id"], args.enable, render_device_id=render_device_id)

    stderr_output = captured_stderr.getvalue()

    # If the setter returned False, attempt read-back verification anyway:
    # some drivers apply the property but the COM call can still signal failure,
    # or the write may succeed while returning non-zero HRESULT due to timing.
    if not ok:
        # First verification attempt: COM read-back.
        actual = _get_listen_to_device_status_ps(target["id"])
        if actual is not None and actual == args.enable:
            _reemit_non_error_stderr(stderr_output)
            print(json.dumps({"listenSet": {"id": target["id"], "name": target["name"], "enabled": actual, "verifiedBy": "com"}}))
            return 0

        # Second verification attempt: registry polling (robust, slower).
        verified, reg_state = _verify_listen_via_registry(target["id"], args.enable, timeout=3.0, interval=0.20)
        if verified or (reg_state is not None and reg_state == args.enable):
            _reemit_non_error_stderr(stderr_output)
            print(json.dumps({"listenSet": {"id": target["id"], "name": target["name"], "enabled": reg_state, "verifiedBy": "registry"}}))
            return 0

        # If we still can't verify, re-emit original stderr and fail.
        sys.stderr.write(stderr_output)
        print(f"ERROR: failed to set 'Listen to this device' for '{target['name']}'.", file=sys.stderr)
        return 1

    # Success path: re-emit any INFO/WARNING text (excluding "ERROR:" lines),
    # then attempt to read the actual state for reporting.
    _reemit_non_error_stderr(stderr_output)

    actual_enabled_state = _get_listen_to_device_status_ps(target["id"])
    if actual_enabled_state is None:
        # If COM read is inconclusive, try registry polling once (best-effort).
        verified, reg_state = _verify_listen_via_registry(target["id"], args.enable, timeout=3.0, interval=0.20)
        if verified or reg_state is not None:
            actual_enabled_state = reg_state

    print(json.dumps({"listenSet": {"id": target["id"], "name": target["name"], "enabled": actual_enabled_state}}))
    return 0


def cmd_get_listen(args):
    """
    Get current 'Listen to this device' enabled/disabled state for a capture device.

    Performance note:
      - This is intentionally a FAST probe for GUI polling.
      - It avoids expensive COM work and prefers a registry-backed read.

    Returns JSON:
      { "id": "...", "name": "...", "flow": "Capture", "listenEnabled": true|false|null }

    Exit codes:
      - 0 success
      - 3 not found
      - 4 multiple matches / index errors
    """
    devices = list_devices(include_all=False)
    matches = find_devices_by_selector(
        devices,
        dev_id=args.id,
        name_substr=args.name,
        flow="Capture",
        regex=args.regex,
    )
    if len(matches) == 0:
        print("ERROR: capture device not found (active only)", file=sys.stderr)
        return 3
    if len(matches) > 1 and args.index is None:
        print("ERROR: multiple matches; specify --index", file=sys.stderr)
        return 4

    # GUI-order index within Capture flow.
    buckets = _sort_and_tag_gui_indices(matches[:])
    ordered = buckets["Capture"]
    if not ordered:
        print("ERROR: no target device found for the specified criteria", file=sys.stderr)
        return 4

    if args.index is not None:
        if args.index < 0 or args.index >= len(ordered):
            print(f"ERROR: --index out of range (0..{len(ordered)-1})", file=sys.stderr)
            return 4
        target = ordered[args.index]
    else:
        target = ordered[0]

    # FAST single probe; no COM; no defaulting.
    state = _read_listen_enable_fast(target["id"])
    print(json.dumps({
        "id": target["id"],
        "name": target["name"],
        "flow": target["flow"],
        "listenEnabled": state,
    }))
    return 0


def cmd_enhancements(args):
    """
    Manage "Audio Enhancements" and per-effect FX toggles.

    This command is intentionally multi-mode, but constrained:
      - Exactly one operation must be requested (enforced here and again in main()).

    Device selection (for all operations except vendor-ini-append):
      - Active-only endpoints.
      - Select by --id or --name (+ optional --flow filter).
      - If multiple matches remain, require --index (GUI order within flow).

    Operations (exactly one):
      - Main enhancements switch (vendor-only at runtime):
          * --enable / --disable
          * --learn (interactive A/B capture; writes vendor_toggles.ini)
      - FX (per-effect toggles learned into vendor_toggles.ini):
          * --list-fx [--json]
          * --learn-fx FX_NAME (two-pass A/B then A2/B2)
          * --enable-fx FX_NAME / --disable-fx FX_NAME (substring/regex match + optional interactive disambiguation)
          * --delete-fx FX_NAME (remove per-device association for that FX bucket)

    Verification and "verifiedBy":
      - Vendor toggles verify by re-reading the learned value(s).
      - JSON outputs may include verifiedBy to explain the confirmation path
        (e.g., vendor section name, decider/quorum mechanism for multi-write FX).

    Exit codes:
      - 0 success
      - 1 invalid args or operation failure
      - 3 not found
      - 4 multiple matches / index errors
    """
    # Validation: exactly one operation.
    ops = [
        bool(args.enable), bool(args.disable), bool(args.learn),
        bool(args.learn_fx), bool(args.enable_fx), bool(args.disable_fx), bool(args.list_fx),
        bool(getattr(args, "delete_fx", None)),
    ]
    if sum(ops) != 1:
        print("ERROR: specify exactly one of --enable, --disable, --learn, --learn-fx, --enable-fx, --disable-fx, or --list-fx", file=sys.stderr)
        return 1

    # Device selection (active-only). Most CLI commands use this same pattern so
    # users get consistent behavior and consistent meaning for --index.
    devices = list_devices(include_all=False)
    matches = find_devices_by_selector(devices, dev_id=args.id, name_substr=args.name, flow=args.flow, regex=args.regex)
    if not matches:
        print("ERROR: device not found (active only)", file=sys.stderr)
        return 3
    if len(matches) > 1 and args.index is None:
        print(_pretty_matches_msg("device", matches), file=sys.stderr)
        return 4

    buckets = _sort_and_tag_gui_indices(matches[:])
    ordered = (buckets.get(args.flow) or []) if args.flow else (buckets["Render"] + buckets["Capture"])
    if not ordered:
        print("ERROR: no target device found for the specified criteria", file=sys.stderr)
        return 4
    target = ordered[args.index] if args.index is not None else ordered[0]

    # === FX Operations (INI-driven) ===

    if args.list_fx:
        # List learned FX entries available for this device (INI only).
        fx_list = _list_fx_for_device(target["id"], target["flow"],
                                      ini_path=getattr(args, "vendor_ini", None))
        fx_list = sorted(fx_list, key=lambda x: (x.get("fx_name") or "").lower())

        # --json returns a machine-friendly list with best-effort state reads.
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

        # Human-readable output is used mainly for interactive CLI sessions.
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
        # Learn a specific effect (FX_NAME). This flow is interactive by design:
        # the user toggles the effect in the Windows UI while we capture snapshots.

        fx_name = args.learn_fx.strip()
        if not fx_name:
            print("ERROR: FX name cannot be empty", file=sys.stderr)
            return 1

        print(f"Learning FX '{fx_name}' for: {target['name']} ({target['flow']})")

        # Drivers often initialize registry state lazily on first toggle. A two-pass
        # capture (A/B, then A2/B2) stabilizes learning by letting the driver "settle"
        # before we record the authoritative pair.
        #
        # AUDIOCTL_LEARN_FX_SETTLE controls the delay before each snapshot to reduce
        # transient writes that can pollute the diff.
        try:
            FX_SETTLE = float(os.environ.get("AUDIOCTL_LEARN_FX_SETTLE", "0.35"))
        except Exception:
            FX_SETTLE = 0.35

        # First pass (A/B) – primarily primes the driver; not necessarily recorded to INI.
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

        # Second pass (A2/B2) – recorded as the authoritative toggle pair.
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

        # vendor_db writes a per-effect entry (single DWORD or multi-write set) into vendor_toggles.ini.
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
        # Enable/disable a learned effect. Matching is flexible:
        # - substring match by default
        # - regex match when --regex is provided
        # If multiple effects match, we prompt interactively to choose one.
        desired = (args.enable_fx or args.disable_fx or "").strip()
        if not desired:
            print("ERROR: FX name cannot be empty", file=sys.stderr)
            return 1

        fx_all = _list_fx_for_device(target["id"], target["flow"], ini_path=getattr(args, "vendor_ini", None))

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

        # Interactive disambiguation if multiple names match the user's request.
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
            ini_path=getattr(args, "vendor_ini", None)
        )
        if ok:
            # verifiedBy identifies how we validated the toggle (e.g., vendor-fx:multi:* vs vendor-fx:*).
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
        # Remove association between an FX bucket and this device GUID in vendor_toggles.ini.
        # This does not necessarily delete the entire bucket (it may apply to other devices).
        desired = (args.delete_fx or "").strip()
        if not desired:
            print("ERROR: FX name cannot be empty", file=sys.stderr)
            return 1

        fx_all = _list_fx_for_device(target["id"], target["flow"], ini_path=getattr(args, "vendor_ini", None))

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

    # === Main Enhancements toggle operations (vendor-only at runtime) ===

    if args.learn:
        # Manual learn flow:
        # - User toggles Windows UI, we capture A/B registry snapshots
        # - We look for a simple DWORD flip and write a per-device vendor entry into the INI
        #
        # Important nuance: Some drivers create their control keys only after the first
        # time the user toggles Enhancements. We therefore do a second attempt if the
        # first pass doesn't detect a clean flip, and also check for newly available
        # vendor entries after the user has toggled at least once.
        ok, info = _learn_vendor_from_discovery_and_write_ini(
            target,
            ini_path=getattr(args, "vendor_ini", None) or _vendor_ini_default_path(),
            prefer_hkcu=True
        )
        if ok:
            print(json.dumps({"vendorLearned": {"id": target["id"], "name": target["name"], "flow": target["flow"], **info}}, indent=2))
            return 0

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

        print("INFO: This may be the first time this endpoint was toggled. The driver often creates keys only after the first toggle.", file=sys.stderr)
        print("INFO: Please toggle Enhancements again (Enable, then Disable) for this same device when prompted.", file=sys.stderr)

        ok2, info2 = _learn_vendor_from_discovery_and_write_ini(
            target,
            ini_path=getattr(args, "vendor_ini", None) or _vendor_ini_default_path(),
            prefer_hkcu=True
        )
        if ok2:
            print(json.dumps({"vendorLearned": {"id": target["id"], "name": target["name"], "flow": target["flow"], **info2}}, indent=2))
            return 0

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

        print("ERROR: No DWORD flip found and no vendor method became available after a retry. Learn failed.", file=sys.stderr)
        return 1

    enable = True if args.enable else False

    # Runtime policy: enhancements toggling is vendor-only. If we don't have a vendor
    # method (INI entry) for this endpoint, we refuse rather than falling back to
    # the Windows Disable_SysFx property (keeps behavior deterministic and reversible).
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
        # verifiedBy identifies which vendor entry was used (and possibly how it was verified).
        print(json.dumps({"enhancementsSet": {"id": target["id"], "name": target["name"], "enabled": state, "verifiedBy": verified_by}}))
        return 0

    print("ERROR: vendor toggle failed.", file=sys.stderr)
    return 1


def cmd_get_enhancements(args):
    """
    Get current enhancements enabled/disabled state from vendor methods only.

    Performance/behavior notes:
      - Vendor-only: we do not consult Windows Disable_SysFx here.
      - Fast path: uses INI-driven probes; intended for GUI polling.
      - UX choice: if INI indicates the device is supported but the state is inconclusive,
        we default-assume enabled=True (many drivers omit explicit "enabled" markers).

    Returns JSON:
      {
        "id": "...",
        "name": "...",
        "flow": "Render"|"Capture",
        "enhancementsEnabled": true|false|null
      }

    Exit codes:
      - 0 success
      - 3 not found
      - 4 multiple matches / index errors
    """
    devices = list_devices(include_all=False)
    matches = find_devices_by_selector(
        devices,
        dev_id=args.id,
        name_substr=args.name,
        flow=args.flow,
        regex=args.regex,
    )
    if not matches:
        print("ERROR: device not found (active only)", file=sys.stderr)
        return 3
    if len(matches) > 1 and args.index is None:
        print(_pretty_matches_msg("device", matches), file=sys.stderr)
        return 4

    buckets = _sort_and_tag_gui_indices(matches[:])
    ordered = (buckets.get(args.flow) or []) if args.flow else (buckets["Render"] + buckets["Capture"])
    if not ordered:
        print("ERROR: no target device found for the specified criteria", file=sys.stderr)
        return 4

    if args.index is not None:
        if args.index < 0 or args.index >= len(ordered):
            print(f"ERROR: --index out of range (0..{len(ordered)-1})", file=sys.stderr)
            return 4
        target = ordered[args.index]
    else:
        target = ordered[0]

    # FAST vendor-only read: only if an INI vendor is known for this endpoint.
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
            # UI-friendly default: if supported but ambiguous, assume enabled.
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
    Return a combined view of device state for GUI/scripts in a single round-trip.

    Purpose:
      - Avoid multiple subprocess calls per right-click menu.
      - Provide a stable JSON contract for status bars / automation.

    Data sources:
      - volume & mute: COM endpoint volume (accurate)
      - listenEnabled: fast registry probe (Capture only)
      - enhancementsEnabled: fast vendor read (INI-driven)
      - availableFX + states: fast vendor reads (INI-driven)

    UX choice:
      - For enhancements/FX, if the device/bucket is known but state is inconclusive,
        default-assume enabled=True to keep the GUI labels deterministic.

    Exit codes:
      - 0 success
      - 3 not found
      - 4 multiple matches / index errors
    """
    import time as _t
    devices = list_devices(include_all=False)
    matches = find_devices_by_selector(
        devices,
        dev_id=args.id,
        name_substr=args.name,
        flow=args.flow,
        regex=args.regex,
    )
    if not matches:
        print("ERROR: device not found (active only)", file=sys.stderr)
        return 3
    if len(matches) > 1 and args.index is None:
        print(_pretty_matches_msg("device", matches), file=sys.stderr)
        return 4

    buckets = _sort_and_tag_gui_indices(matches[:])
    ordered = (buckets.get(args.flow) or []) if args.flow else (buckets["Render"] + buckets["Capture"])
    if not ordered:
        print("ERROR: no target device found for the specified criteria", file=sys.stderr)
        return 4

    if args.index is not None:
        if args.index < 0 or args.index >= len(ordered):
            print(f"ERROR: --index out of range (0..{len(ordered)-1})", file=sys.stderr)
            return 4
        target = ordered[args.index]
    else:
        target = ordered[0]

    dev_id = target["id"]
    flow   = target["flow"]

    # Volume & mute: no artificial sleeps; keep COM accuracy and keep GUI responsive.
    vol = get_endpoint_volume(dev_id)
    muted = get_endpoint_mute(dev_id)
    if muted is not None:
        muted = bool(muted)

    # Listen is only meaningful for capture endpoints; do a fast registry-backed read.
    listen_enabled = None
    if flow == "Capture":
        try:
            listen_enabled = _read_listen_enable_fast(dev_id)
        except Exception:
            listen_enabled = None

    # Enhancements (vendor-only) + FX list. We keep these fast and INI-driven so the GUI
    # can poll safely without COM lifecycle complications.
    from .vendor_db import (
        _fast_get_enhancements_state,
        _list_fx_for_device,
        _fast_read_vendor_entry_state,
        _vendor_ini_default_path,
        _enhancements_supported,
    )

    # Determine effective INI path (GUI may override for portable use).
    ini_path = getattr(args, "vendor_ini", None)
    if not ini_path:
        try:
            ini_path = _vendor_ini_default_path()
        except Exception:
            ini_path = None

    # If the device has no known vendor toggle yet, skip vendor state reads entirely.
    has_vendor = False
    try:
        has_vendor = _enhancements_supported(dev_id, flow)
    except Exception:
        has_vendor = False

    # Enhancements (vendor-only, fast).
    if has_vendor:
        try:
            enh_enabled = _fast_get_enhancements_state(dev_id, flow)
        except Exception:
            enh_enabled = None
        if enh_enabled is None:
            # UI-friendly default when indicator absent/inconclusive.
            enh_enabled = True
    else:
        enh_enabled = None

    # FX list with fast state reads. We list FX if present in INI regardless of whether
    # the main enhancements toggle is supported (FX can be learned independently).
    available_fx = []
    try:
        fx_list = _list_fx_for_device(dev_id, flow, ini_path=ini_path)
        fx_list = sorted(fx_list, key=lambda x: (x.get("fx_name") or "").lower())
        for fx in fx_list:
            entry = fx.get("entry")
            state = None
            try:
                state = _fast_read_vendor_entry_state(entry, dev_id, flow)
            except Exception:
                state = None
            if state is None:
                # UI-friendly default when indicator absent/inconclusive.
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
    """
    Diagnostic helper: compare multiple views of Enhancements state.

    Purpose:
      - Show the "live" Windows perspective (PropertyStore + COM PolicyConfigFx)
      - Show the vendor toggle entry (INI) if one exists and what it reads now

    Output:
      - Pretty-printed JSON including:
          enhancementsEnabled_live_propstore
          enhancementsEnabled_live_com
          vendor_toggle_status

    Exit codes:
      - 0 success
      - 3 not found
      - 4 multiple matches / index errors
    """
    devices = list_devices(include_all=False)
    matches = find_devices_by_selector(devices, dev_id=args.id, name_substr=args.name, flow=args.flow, regex=args.regex)
    if not matches:
        print("ERROR: device not found (active only)", file=sys.stderr)
        return 3

    buckets = _sort_and_tag_gui_indices([d for d in matches])
    ordered = (buckets.get(args.flow) or []) if args.flow else (buckets["Render"] + buckets["Capture"])
    if not ordered:
        print("ERROR: no target device found for the specified criteria", file=sys.stderr)
        return 4

    if args.index is not None:
        if args.index < 0 or args.index >= len(ordered):
            print(f"ERROR: --index out of range (0..{len(ordered)-1})", file=sys.stderr)
            return 4
        target = ordered[args.index]
    else:
        target = ordered[0]

    # "live" readings are for diagnosis: they show what Windows thinks, not what our vendor-only
    # runtime toggle might read/write.
    live_win = _get_enhancements_status_propstore(target["id"])
    live_com = _get_enhancements_status_com(target["id"])

    # Vendor entry is INI-driven (if learned). This is what runtime toggles use.
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
    """
    Interactive discovery tool for Enhancements (SysFX) behavior.

    Purpose:
      - Capture two full snapshots (A=enabled, B=disabled)
      - Diff MMDevices registry values to identify likely toggle keys
      - Emit:
          * a human-readable TXT report
          * a JSON bundle containing snapshots + diff data
      - Optionally generate/append an INI snippet candidate for vendor_toggles.ini.

    Notes:
      - This is intended for debugging/learning and may include a lot of registry noise.
      - Normal runtime toggling is vendor-only and uses learned INI entries.

    Exit codes:
      - 0 success
      - 3 not found
      - 4 multiple matches / index errors
    """
    devices = list_devices(include_all=False)
    matches = find_devices_by_selector(devices, dev_id=args.id, name_substr=args.name, flow=args.flow, regex=args.regex)
    if not matches:
        print("ERROR: device not found (active only)", file=sys.stderr)
        return 3
    if len(matches) > 1 and args.index is None:
        print(_pretty_matches_msg("device", matches), file=sys.stderr)
        return 4

    buckets = _sort_and_tag_gui_indices(matches[:])
    ordered = (buckets.get(args.flow) or []) if args.flow else (buckets["Render"] + buckets["Capture"])
    if not ordered:
        print("ERROR: no target device found for the specified criteria", file=sys.stderr)
        return 4

    if args.index is not None:
        if args.index < 0 or args.index >= len(ordered):
            print(f"ERROR: --index out of range (0..{len(ordered)-1})", file=sys.stderr)
            return 4
        target = ordered[args.index]
    else:
        target = ordered[0]

    print(f"Discovery target: {target['name']} [{target['id']}] ({target['flow']})")
    print("Step 1: In Windows Sound settings, set 'Audio Enhancements' to ENABLED for this device.")
    input("When ready, press Enter to capture snapshot A... ")
    snapA = _collect_sysfx_snapshot(target["id"])
    print("Step 2: Now set 'Audio Enhancements' to DISABLED for the same device.")
    input("When ready, press Enter to capture snapshot B... ")
    snapB = _collect_sysfx_snapshot(target["id"])

    diffs = _diff_mmdevices_lists(snapA.get("registry") or [], snapB.get("registry") or [])

    # Write a timestamped report bundle for later analysis/sharing.
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

    # Optional helper: append a suggested vendor INI snippet if we found a clean DWORD flip candidate.
    if getattr(args, "ini_snippet", None):
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
    """
    Wait until a device becomes active (appears), or timeout.

    Selection rules:
      - Active-only matching (DEVICE_STATE_ACTIVE).
      - Uses same selectors: --id or --name (+ optional --flow, --regex).
      - If multiple matches occur, require --index (GUI order within flow).

    Output:
      - On success: {"found": <device_dict>} in JSON (single line).

    Exit codes:
      - 0 found
      - 3 timeout / not found
      - 4 multiple matches / index errors
    """
    deadline = time.time() + args.timeout
    while time.time() < deadline:
        devices = list_devices(include_all=False)
        matches = find_devices_by_selector(devices, dev_id=args.id, name_substr=args.name, flow=args.flow, regex=args.regex)
        if matches:
            buckets = _sort_and_tag_gui_indices(matches[:])
            flow = args.flow or (matches[0]["flow"] if matches else None)
            ordered = (buckets.get(flow) or []) if args.flow else (buckets["Render"] + buckets["Capture"])
            if not ordered:
                print("ERROR: no target device found for the specified criteria", file=sys.stderr)
                return 4
            if args.index is not None:
                if args.index < 0 or args.index >= len(ordered):
                    print(f"ERROR: --index out of range (0..{len(ordered)-1})", file=sys.stderr)
                    return 4
                target = ordered[args.index]
            else:
                target = ordered[0]
            print(json.dumps({"found": target}))
            return 0
        time.sleep(0.5)
    print("ERROR: timeout waiting for device", file=sys.stderr)
    return 3


def build_parser():
    """
    Build the top-level argparse parser.

    Contract note:
      - The GUI relies on these subcommands, flags, and JSON output shapes.
      - Keep flag semantics stable; treat changes here as API changes.

    Also defines one hidden helper subcommand:
      - vendor-ini-append: used by the GUI when it needs to write vendor_toggles.ini
        into a protected directory (e.g., Program Files). The GUI can launch an elevated
        process with a "work order" JSON describing the exact append operation.
    """
    p = argparse.ArgumentParser(prog="audioctl", description="Windows audio control CLI (pycaw-based)")
    sub = p.add_subparsers(dest="cmd", required=True)

    # --- Device enumeration ---
    p_list = sub.add_parser("list", help="List devices")
    p_list.add_argument("--all", action="store_true", help="Include disabled/disconnected")
    p_list.add_argument("--json", action="store_true")
    p_list.set_defaults(func=cmd_list)

    # --- Default device management ---
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

    # --- Volume/mute ---
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

    # --- Read-only helpers (GUI/scripts) ---
    p_gv = sub.add_parser("get-volume", help="Get endpoint volume and mute state (render or capture)")
    p_gv.add_argument("--id")
    p_gv.add_argument("--name")
    p_gv.add_argument("--flow", choices=["Render", "Capture"], help="Optional filter to disambiguate")
    p_gv.add_argument("--index", type=int)
    p_gv.add_argument("--regex", action="store_true")
    p_gv.set_defaults(func=cmd_get_volume)

    # --- Listen feature (Capture only) ---
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
    p_gl.set_defaults(func=cmd_get_listen)

    # --- Enhancements + FX (vendor-only runtime + learn/discovery tools) ---
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
    # FX args
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

    # --- Diagnostics ---
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

    p_learn = sub.add_parser("discover-enhancements", help="Interactively learn how Enhancements toggles for a device")
    p_learn.add_argument("--id")
    p_learn.add_argument("--name")
    p_learn.add_argument("--flow", choices=["Render", "Capture"])
    p_learn.add_argument("--index", type=int)
    p_learn.add_argument("--regex", action="store_true")
    p_learn.add_argument("--output-dir", help="Where to write the TXT/JSON report (default: current directory)")
    p_learn.add_argument("--ini-snippet", help="Write a suggested vendor INI section to this path (append).")
    p_learn.set_defaults(func=cmd_discover_enhancements)

    # --- Wait/poll helper ---
    p_w = sub.add_parser("wait", help="Wait for device to appear")
    p_w.add_argument("--id")
    p_w.add_argument("--name")
    p_w.add_argument("--flow", choices=["Render", "Capture"])
    p_w.add_argument("--timeout", type=int, default=30)
    p_w.add_argument("--index", type=int)
    p_w.add_argument("--regex", action="store_true")
    p_w.set_defaults(func=cmd_wait)

    # Hidden helper: elevated INI append (used by GUI to write into Program Files)
    p_vi = sub.add_parser("vendor-ini-append", help=argparse.SUPPRESS)
    p_vi.add_argument("--work", required=True, help=argparse.SUPPRESS)  # path to JSON work order
    p_vi.set_defaults(func=cmd_vendor_ini_append)

    return p


def cmd_diag_mmdevices(args):
    """
    Diagnostic helper: dump all MMDevices registry values for an endpoint.

    Purpose:
      - Provides raw data used by learn/discovery flows.
      - Useful for debugging vendor toggles when a driver uses non-DWORD data.

    Output:
      - Pretty JSON with "mmdevices": a list of (hive/flow/subkey/name/type/dataPreview/dataRaw).

    Exit codes:
      - 0 success
      - 3 not found
      - 4 multiple matches / index errors
    """
    devices = list_devices(include_all=False)
    matches = find_devices_by_selector(devices, dev_id=args.id, name_substr=args.name, flow=args.flow, regex=args.regex)
    if not matches:
        print("ERROR: device not found (active only)", file=sys.stderr)
        return 3
    if len(matches) > 1 and args.index is None:
        print("ERROR: multiple matches; specify --index", file=sys.stderr)
        return 4

    buckets = _sort_and_tag_gui_indices([d for d in matches])
    ordered = (buckets.get(args.flow) or []) if args.flow else (buckets["Render"] + buckets["Capture"])
    if not ordered:
        print("ERROR: no target device found for the specified criteria", file=sys.stderr)
        return 4

    if args.index is not None:
        if args.index < 0 or args.index >= len(ordered):
            print(f"ERROR: --index out of range (0..{len(ordered)-1})", file=sys.stderr)
            return 4
        target = ordered[args.index]
    else:
        target = ordered[0]

    dump = _dump_mmdevices_all_values(target["id"])
    print(json.dumps({"id": target["id"], "name": target["name"], "flow": target["flow"], "mmdevices": dump}, indent=2))
    return 0


def cmd_vendor_ini_append(args):
    """
    Hidden helper used by the GUI for elevated writes to vendor_toggles.ini.

    Why it exists:
      - When the INI lives next to the executable in Program Files, normal users can't write it.
      - The GUI can generate a "work order" JSON file and re-launch audioctl elevated to append
        exactly one vendor entry in a controlled way.

    Work order:
      - JSON file describing one of:
          * kind="main" (append main toggle section if missing)
          * kind="fx"   (append fx entry)

    Output:
      - JSON describing the append result.

    Exit codes:
      - 0 success (including "exists" cases that are not errors)
      - 1 invalid input or append failure
    """
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
        # This typically means the process wasn't actually elevated or the target directory is protected.
        print(f"ERROR: permission denied writing INI: {e}", file=sys.stderr)
        return 1

    except FileExistsError as e:
        # Some append helpers may signal "already exists" via exception; treat as success.
        print(json.dumps({"iniAppend": {"kind": kind, "result": "exists", "iniPath": ini_path}}))
        return 0

    except Exception as e:
        print(f"ERROR: failed to append INI: {e}", file=sys.stderr)
        return 1


def main(argv=None):
    """
    CLI entry point.

    Behavior:
      - If launched with no arguments, start the GUI (lazy import) as a convenience.
      - Otherwise parse args and dispatch to the selected cmd_* handler.

    COM lifecycle note:
      - We do NOT call CoInitialize/CoUninitialize here.
      - Each low-level helper in devices.py manages COM per call via a thread-local context.
        This avoids global COM lifetime issues and reduces GC/finalizer timing hazards.

    Exit codes:
      - Propagates cmd_* return codes.
      - 130 for KeyboardInterrupt (Ctrl-C).
    """
    # No args => GUI mode (keeps UX friendly for double-click / end-users).
    # Lazy import avoids importing Tkinter on CLI-only usage.
    if argv is None and len(sys.argv) <= 1:
        try:
            from .gui import launch_gui  # Lazy import only if we actually need the GUI
            return launch_gui()
        except Exception as e:
            print(f"ERROR: GUI failed to start: {e}", file=sys.stderr)

    parser = build_parser()
    args = parser.parse_args(argv)

    # Centralized validation for mutually exclusive flags.
    # We do this here (in addition to per-command checks) to keep behavior consistent
    # even if commands are invoked by the GUI with unexpected flag combinations.
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
