# audioctl/cli.py
import sys
import argparse
import json
import time
import io
import os
import warnings
import re
from contextlib import redirect_stderr
import comtypes
from .compat import (
    E_RENDER, E_CAPTURE,
    ROLES, DEVICE_STATE_ACTIVE, DEVICE_STATE_ALL, is_admin,
)
from .logging_setup import _log, _log_exc
from .devices import (
    list_devices, find_devices_by_selector, _sort_and_tag_gui_indices,
    enum_endpoints, get_default_ids,
    _is_device_active, _pretty_matches_msg, _select_by_name_active_only,
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
)
from .vendor_db import (
    _vendor_ini_default_path,
    _enhancements_supported,
    _apply_enhancements,
    _learn_vendor_from_discovery_and_write_ini,
    _learn_vendor_and_write_ini,
    _build_vendor_ini_snippet,
    _find_first_vendor_entry,
    _read_vendor_entry_state,
)
from .gui import launch_gui

def cmd_list(args):
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
    if not is_admin():
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

def cmd_listen(args):
    # Resolve playback target. Start with ID if it was provided.
    render_device_id = args.playback_target_id
    # Handle --playback-target-name. This will override the ID if both are used.
    if args.playback_target_name is not None:
        # If the flag was used without a value, argparse sets it to const=''
        if args.playback_target_name == '':
            render_device_id = ''
        else:
            # A name was provided, so find the device ID
            render_target, err = _select_by_name_active_only("Render", args.playback_target_name, None, args.regex)
            if err:
                print(f"ERROR: Could not find playback target device: {err}", file=sys.stderr)
                return 3
            render_device_id = render_target["id"]

    # --- The rest of the function is for finding the CAPTURE device ---
    devices = list_devices(include_all=False)
    matches = find_devices_by_selector(devices, dev_id=args.id, name_substr=args.name, flow="Capture", regex=args.regex)
    if len(matches) == 0:
        print("ERROR: capture device not found (active only)", file=sys.stderr)
        return 3
    if len(matches) > 1 and args.index is None:
        print("ERROR: multiple matches; specify --index", file=sys.stderr)
        return 4
    
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
        
    # --- Now, call the device function with the resolved render_device_id ---
    captured_stderr = io.StringIO()
    ok = False
    with redirect_stderr(captured_stderr):
        # render_device_id will be a string ID, '', or None
        ok = set_listen_to_device_ps(target["id"], args.enable, render_device_id=render_device_id)
        
    stderr_output = captured_stderr.getvalue()
    if not ok:
        actual = _get_listen_to_device_status_ps(target["id"])
        if actual is not None and actual == args.enable:
            _reemit_non_error_stderr(stderr_output)
            print(json.dumps({"listenSet": {"id": target["id"], "name": target["name"], "enabled": actual, "verifiedBy": "com"}}))
            return 0
        verified, reg_state = _verify_listen_via_registry(target["id"], args.enable, timeout=3.0, interval=0.20)
        if verified or (reg_state is not None and reg_state == args.enable):
            _reemit_non_error_stderr(stderr_output)
            print(json.dumps({"listenSet": {"id": target["id"], "name": target["name"], "enabled": reg_state, "verifiedBy": "registry"}}))
            return 0
        sys.stderr.write(stderr_output)
        print(f"ERROR: failed to set 'Listen to this device' for '{target['name']}'.", file=sys.stderr)
        return 1
        
    _reemit_non_error_stderr(stderr_output)
    actual_enabled_state = _get_listen_to_device_status_ps(target["id"])
    if actual_enabled_state is None:
        verified, reg_state = _verify_listen_via_registry(target["id"], args.enable, timeout=3.0, interval=0.20)
        if verified or reg_state is not None:
            actual_enabled_state = reg_state
            
    print(json.dumps({"listenSet": {"id": target["id"], "name": target["name"], "enabled": actual_enabled_state}}))
    return 0

def cmd_enhancements(args):
    devices = list_devices(include_all=False)
    matches = find_devices_by_selector(devices, dev_id=args.id, name_substr=args.name, flow=args.flow, regex=args.regex)
    if len(matches) == 0:
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
        
    target = ordered[args.index] if args.index is not None else ordered[0]

    if getattr(args, "learn", False):
        ok, info = _learn_vendor_from_discovery_and_write_ini(
            target,
            ini_path=getattr(args, "vendor_ini", None) or _vendor_ini_default_path(),
            prefer_hkcu=True
        )
        if ok:
            print(json.dumps({"vendorLearned": {"id": target["id"], "name": target["name"], "flow": target["flow"], **info}}, indent=2))
            return 0
        else:
            print(f"ERROR: learn failed: {info}", file=sys.stderr)
            return 1

    enable = True if args.enable else False
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

def cmd_diag_sysfx(args):
    devices = list_devices(include_all=False)
    matches = find_devices_by_selector(devices, dev_id=args.id, name_substr=args.name, flow=args.flow, regex=args.regex)
    if not matches:
        print("ERROR: device not found (active only)", file=sys.stderr)
        return 3
    
    buckets = _sort_and_tag_gui_indices([d for d in matches])
    ordered = (buckets.get(args.flow) or []) if args.flow else (buckets["Render"] + buckets["Capture"])
    target = ordered[args.index] if args.index is not None else ordered[0]

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
    target = ordered[args.index] if args.index is not None else ordered[0]

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
    p = argparse.ArgumentParser(prog="audioctl", description="Windows audio control CLI (pycaw-based)")
    sub = p.add_subparsers(dest="cmd", required=True)

    p_list = sub.add_parser("list", help="List devices")
    p_list.add_argument("--all", action="store_true", help="Include disabled/disconnected")
    p_list.add_argument("--json", action="store_true")
    p_list.set_defaults(func=cmd_list)

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

    p_sv = sub.add_parser("set-volume", help="Set endpoint volume (render or capture) or mute/unmute")
    p_sv.add_argument("--id")
    p_sv.add_argument("--name")
    p_sv.add_argument("--flow", choices=["Render", "Capture"], help="Optional filter to disambiguate")
    p_sv.add_argument("--level", type=int, help="0-100 (for volume)")
    p_sv.add_argument("--mute", action="store_true", help="Mute the device")
    p_sv.add_argument("--unmute", action="store_true", help="Unmute the device")
    p_sv.add_argument("--index", type=int)
    p_sv.add_argument("--regex", action="store_true")
    p_sv.set_defaults(func=cmd_set_volume)

    p_ls = sub.add_parser("listen", help="Enable/disable 'Listen to this device' (capture only)")
    p_ls.add_argument("--id", help="Device ID for the capture device.")
    p_ls.add_argument("--name", help="Substring of the device name for the capture device.")
    p_ls.add_argument("--enable", action="store_true", help="Enable 'Listen to this device'.")
    p_ls.add_argument("--disable", action="store_true", help="Disable 'Listen to this device'.")
    p_ls.add_argument("--playback-target-id", nargs='?', const='', default=None, help="Optional: Render endpoint ID to play through. Use without a value for 'Default Playback Device'.")
    p_ls.add_argument("--playback-target-name", nargs='?', const='', default=None, help="Optional: Render endpoint name to play through. Use without a value for 'Default Playback Device'.")
    p_ls.add_argument("--index", type=int)
    p_ls.add_argument("--regex", action="store_true")
    p_ls.set_defaults(func=cmd_listen)

    p_fx = sub.add_parser("enhancements", help="Enable/disable 'Audio Enhancements' (SysFX) on a device")
    p_fx.add_argument("--id", help="Endpoint ID")
    p_fx.add_argument("--name", help="Substring or regex of the endpoint name")
    p_fx.add_argument("--flow", choices=["Render", "Capture"], help="Optional filter to disambiguate")
    p_fx.add_argument("--enable", action="store_true", help="Enable audio enhancements")
    p_fx.add_argument("--disable", action="store_true", help="Disable audio enhancements")
    p_fx.add_argument("--index", type=int, help="GUI-order index among matches")
    p_fx.add_argument("--regex", action="store_true")
    p_fx.add_argument("--prefer-hklm", action="store_true",
                      help="When falling back to the registry, try HKLM first (Admin required to write).")
    p_fx.add_argument("--vendor-ini", help="Path to vendor_toggles.ini (default: next to the EXE).")
    p_fx.add_argument("--learn", action="store_true",
                      help="Manual learn (you toggle Windows UI). Captures A/B and writes vendor INI; no Windows fallback is used at runtime.")
    p_fx.set_defaults(func=cmd_enhancements)

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

    p_w = sub.add_parser("wait", help="Wait for device to appear")
    p_w.add_argument("--id")
    p_w.add_argument("--name")
    p_w.add_argument("--flow", choices=["Render", "Capture"])
    p_w.add_argument("--timeout", type=int, default=30)
    p_w.add_argument("--index", type=int)
    p_w.add_argument("--regex", action="store_true")
    p_w.set_defaults(func=cmd_wait)

    return p

def cmd_diag_mmdevices(args):
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

def main(argv=None):
    if argv is None and len(sys.argv) <= 1:
        try:
            return launch_gui()
        except Exception as e:
            print(f"ERROR: GUI failed to start: {e}", file=sys.stderr)

    try:
        comtypes.CoInitialize()
    except Exception:
        pass

    parser = build_parser()
    args = parser.parse_args(argv)

    if args.cmd == "listen":
        if args.enable and args.disable:
            print("ERROR: specify only one of --enable or --disable", file=sys.stderr)
            try:
                comtypes.CoUninitialize()
            except Exception:
                pass
            return 1
        if not args.enable and not args.disable:
            print("ERROR: specify --enable or --disable", file=sys.stderr)
            try:
                comtypes.CoUninitialize()
            except Exception:
                pass
            return 1
        args.enable = True if args.enable else False

    if args.cmd == "enhancements":
        trio = int(bool(args.enable)) + int(bool(args.disable)) + int(bool(args.learn))
        if trio != 1:
            print("ERROR: specify exactly one of --enable, --disable, or --learn", file=sys.stderr)
            try:
                comtypes.CoUninitialize()
            except Exception:
                pass
            return 1

    if args.cmd == "set-volume":
        if (args.mute or args.unmute) and args.level is not None:
            print("ERROR: Cannot specify both --level and --mute/--unmute", file=sys.stderr)
            try:
                comtypes.CoUninitialize()
            except Exception:
                pass
            return 1
        if not (args.mute or args.unmute or args.level is not None):
            print("ERROR: Must specify --level or --mute/--unmute", file=sys.stderr)
            try:
                comtypes.CoUninitialize()
            except Exception:
                pass
            return 1

    try:
        rc = args.func(args)
    except KeyboardInterrupt:
        rc = 130
    finally:
        try:
            comtypes.CoUninitialize()
        except Exception:
            pass
    return rc
