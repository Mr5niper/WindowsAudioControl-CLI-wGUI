# audioctl/vendor_db.py
import os
import re
import configparser
import time
import winreg
from .compat import _guid_from_parts, is_admin
from .logging_setup import _exe_dir
from .devices import (
    _extract_endpoint_guid_from_device_id,
    _read_enhancements_from_registry,
    _verify_enhancements_via_registry,
    _set_enhancements_registry,
    _get_enhancements_status_propstore,
    _set_enhancements_propstore,
    _get_enhancements_status_com,
    _collect_sysfx_snapshot,
    _diff_mmdevices_lists,
    _generate_enh_discovery_report,
    _short_settle,
    _verify_effect_only, # Used by auto-learn to confirm Windows changes
)
# Code vendor entries (Realtek, etc.)
_CODE_VENDOR_ENTRIES = [
    {
        "name": "realtek_waves_primary",
        "value_name": "{1da5d803-d492-4edd-8c23-e0c0ffee7f0e},5",
        "enable": 0,
        "disable": 1,
        "hives": ["HKLM","HKCU"],
        "flows": ["Render","Capture"],
        "notes": "Realtek/Waves primary DWORD toggle (0=enabled,1=disabled)"
    },
]
def _vendor_ini_default_path():
    try:
        return os.path.join(_exe_dir(), "vendor_toggles.ini")
    except Exception:
        return os.path.join(os.getcwd(), "vendor_toggles.ini")
def _load_vendor_db(ini_path=None):
    """
    Load vendor toggles from INI. Returns list of normalized entries.
    """
    path = ini_path or _vendor_ini_default_path()
    cfg = configparser.ConfigParser()
    entries = []
    try:
        if not os.path.exists(path):
            return entries
        cfg.read(path, encoding="utf-8")
    except Exception:
        return entries
        
    for sec in cfg.sections():
        try:
            value_name = cfg.get(sec, "value_name").strip().lower()
            en = int(cfg.get(sec, "dword_enable"))
            di = int(cfg.get(sec, "dword_disable"))
            if en not in (0,1) or di not in (0,1) or en == di:
                continue
            hives = [x.strip().upper() for x in cfg.get(sec, "hives", fallback="HKLM,HKCU").split(",") if x.strip()]
            flows = [x.strip().capitalize() for x in cfg.get(sec, "flows", fallback="Render,Capture").split(",") if x.strip()]
            notes = cfg.get(sec, "notes", fallback="")
            entries.append({
                "name": sec,
                "value_name": value_name,
                "enable": en,
                "disable": di,
                "hives": [h for h in hives if h in ("HKLM","HKCU")],
                "flows": [f for f in flows if f in ("Render","Capture")],
                "notes": notes,
            })
        except Exception:
            continue
    return entries
def _endpoint_fx_key(device_id, flow):
    guid = _extract_endpoint_guid_from_device_id(device_id)
    if not guid:
        return None, None
    flow_name = "Render" if str(flow).lower().startswith("r") else "Capture"
    key_path = rf"SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\{flow_name}\{guid}\FxProperties"
    return flow_name, key_path
def _vendor_entry_applies(entry, device_id, flow):
    """
    Return True if entry.flow matches and REG_DWORD value exists under any listed hive for this endpoint.
    """
    flow_name = "Render" if str(flow).lower().startswith("r") else "Capture"
    if entry.get("flows") and flow_name not in entry["flows"]:
        return False
    _f, key_path = _endpoint_fx_key(device_id, flow)
    if not key_path:
        return False
    for h in entry.get("hives", []):
        hive = winreg.HKEY_LOCAL_MACHINE if h == "HKLM" else winreg.HKEY_CURRENT_USER
        try:
            with winreg.OpenKey(hive, key_path, 0, winreg.KEY_READ) as key:
                try:
                    _, typ = winreg.QueryValueEx(key, entry["value_name"])
                    if typ == winreg.REG_DWORD:
                        return True
                except FileNotFoundError:
                    continue
        except OSError:
            continue
    return False
def _set_vendor_entry_state(entry, device_id, flow, enable):
    """
    Write vendor entry DWORD to desired value across its configured hives. Returns True if any write succeeded.
    """
    _f, key_path = _endpoint_fx_key(device_id, flow)
    if not key_path:
        return False
    desired = entry["enable"] if enable else entry["disable"]
    ok = False
    for h in entry.get("hives", []):
        hive = winreg.HKEY_LOCAL_MACHINE if h == "HKLM" else winreg.HKEY_CURRENT_USER
        try:
            with winreg.OpenKey(hive, key_path, 0, winreg.KEY_SET_VALUE) as key:
                winreg.SetValueEx(key, entry["value_name"], 0, winreg.REG_DWORD, int(desired))
                ok = True or ok
        except OSError:
            continue
    return ok
def _read_vendor_entry_state(entry, device_id, flow):
    """
    Return True if current DWORD equals 'enable' value, False if equals 'disable', None otherwise.
    """
    _f, key_path = _endpoint_fx_key(device_id, flow)
    if not key_path:
        return None
    configured = entry.get("hives") or []
    read_order = [h for h in ("HKCU", "HKLM") if h in configured]
    hive_map = {"HKCU": winreg.HKEY_CURRENT_USER, "HKLM": winreg.HKEY_LOCAL_MACHINE}
    for hname in read_order:
        hive = hive_map[hname]
        try:
            with winreg.OpenKey(hive, key_path, 0, winreg.KEY_READ) as key:
                try:
                    val, typ = winreg.QueryValueEx(key, entry["value_name"])
                    if typ == winreg.REG_DWORD:
                        v = int(val)
                        if v == entry["enable"]:
                            return True
                        if v == entry["disable"]:
                            return False
                except FileNotFoundError:
                    continue
        except OSError:
            continue
    return None
def _verify_vendor_entry(entry, device_id, flow, expected_enabled, timeout=2.5, interval=0.2, consecutive=2):
    """
    Poll the same vendor DWORD until it reflects expected_enabled for 'consecutive' reads or timeout.
    """
    end = time.time() + float(timeout)
    ok_streak = 0
    last = None
    while time.time() < end:
        st = _read_vendor_entry_state(entry, device_id, flow)
        last = st
        if st is not None and st == bool(expected_enabled):
            ok_streak += 1
            if ok_streak >= consecutive:
                return True, st
        else:
            ok_streak = 0
        time.sleep(interval)
    return False, last
def _try_vendor_first(device_id, flow, enable, ini_path=None):
    """
    Try INI vendor entries first, then code vendors.
    """
    # 1) INI vendors first
    vendor_db = _load_vendor_db(ini_path)
    for entry in vendor_db:
        try:
            if _vendor_entry_applies(entry, device_id, flow):
                wrote = _set_vendor_entry_state(entry, device_id, flow, enable)
                if wrote:
                    ok, st = _verify_vendor_entry(entry, device_id, flow, enable, timeout=2.5, interval=0.2, consecutive=2)
                    if ok:
                        return True, f"vendor:{entry['name']}", st
        except Exception:
            continue
    # 2) Code vendors next
    for entry in _CODE_VENDOR_ENTRIES:
        try:
            if _vendor_entry_applies(entry, device_id, flow):
                wrote = _set_vendor_entry_state(entry, device_id, flow, enable)
                if wrote:
                    ok, st = _verify_vendor_entry(entry, device_id, flow, enable, timeout=2.5, interval=0.2, consecutive=2)
                    if ok:
                        return True, f"vendor:{entry['name']}", st
        except Exception:
            continue
    return False, None, None
def _find_first_vendor_entry(device_id, flow, ini_path=None):
    """
    Read-only: return the first vendor entry that applies to this endpoint.
    """
    for entry in _load_vendor_db(ini_path):
        try:
            if _vendor_entry_applies(entry, device_id, flow):
                return entry
        except Exception:
            continue
    for entry in _CODE_VENDOR_ENTRIES:
        try:
            if _vendor_entry_applies(entry, device_id, flow):
                return entry
        except Exception:
            continue
    return None
def _sanitize_ini_section_name(value_name: str):
    # e.g. "{1da5d803-...},5" -> "vendor_{1da5d803-...},5"
    base = re.sub(r'[^A-Za-z0-9_,\-{}]+', "_", value_name)
    return f"vendor_{base}"
def _append_vendor_ini_entry_if_missing(ini_path, section_name, value_name, dword_enable, dword_disable,
                                        flows="Render,Capture", hives="HKCU,HKLM", notes=""):
    """
    Append a vendor INI section to ini_path only if it does not already exist.
    """
    cfg = configparser.ConfigParser()
    try:
        if os.path.exists(ini_path):
            cfg.read(ini_path, encoding="utf-8")
    except Exception:
        pass
    if cfg.has_section(section_name):
        return "exists"
    try:
        ini_dir = os.path.dirname(ini_path)
        if ini_dir:
            os.makedirs(ini_dir, exist_ok=True)
    except Exception:
        pass
    lines = []
    lines.append("")
    lines.append(f"[{section_name}]")
    lines.append(f"value_name = {str(value_name).strip().lower()}")
    lines.append(f"dword_enable = {int(dword_enable)}")
    lines.append(f"dword_disable = {int(dword_disable)}")
    lines.append(f"hives = {hives}")
    lines.append(f"flows = {flows}")
    if notes:
        lines.append(f"notes = {notes}")
    text = "\n".join(lines) + "\n"
    with open(ini_path, "a", encoding="utf-8", errors="replace") as f:
        f.write(text)
    return "appended"
def _build_vendor_ini_snippet(target, snapA, snapB, diffs, section_name=None):
    """
    Build a suggested vendor INI section based on DWORD flips observed.
    """
    cands = []
    for f in diffs.get("dword_flips", []):
        name = str(f.get("name",""))
        subkey = str(f.get("subkey",""))
        hive = str(f.get("hive",""))
        if not (name.startswith("{") and "}" in name and "," in name):
            continue
        if subkey != "FxProperties":
            continue
        before = int(f.get("before"))
        after  = int(f.get("after"))
        cands.append({
            "hive": hive, "name": name, "before": before, "after": after, "subkey": subkey
        })
        
    if not cands:
        return None, None
        
    cands.sort(key=lambda x: (0 if x["hive"]=="HKLM" else 1))
    pick = cands[0]
    
    dword_enable = pick["before"]
    dword_disable = pick["after"]
    
    if not section_name:
        base = re.sub(r'[^A-Za-z0-9_,\-{}]+', "_", pick["name"])
        section_name = f"vendor_{base}"
        
    flows_all = "Render,Capture"
    hives = "HKCU,HKLM"
    notes = f"Auto-learned from discovery on device '{target.get('name')}' ({target.get('flow')}). A=enabled,B=disabled."
    
    snippet = []
    snippet.append(f"[{section_name}]")
    snippet.append(f"value_name = {pick['name']}")
    snippet.append(f"dword_enable = {int(dword_enable)}")
    snippet.append(f"dword_disable = {int(dword_disable)}")
    snippet.append(f"hives = {hives}")
    snippet.append(f"flows = {flows_all}")
    snippet.append(f"notes = {notes}")
    
    return "\n".join(snippet) + "\n", pick
def _learn_vendor_from_discovery_and_write_ini(target, ini_path=None, prefer_hkcu=True):
    """
    Manual learn using discovery flow.
    """
    import sys
    dev_id = target["id"]
    flow   = target["flow"]
    name   = target["name"]
    ini_path = ini_path or _vendor_ini_default_path()
    # STERN WARNING + explicit confirmation
    warning = f"""
READ CAREFULLY
This Learn mode will capture two registry snapshots and write a vendor entry into:
  {ini_path}
From now on, future 'enhancements' commands for this device WILL WRITE registry values on this machine (HKCU/optional HKLM) to toggle Enhancements. This is persistent until you manually remove the learned section from vendor_toggles.ini.
Critical rules during Learn:
- Do NOT change any other audio settings.
- Do NOT switch default devices.
- Do NOT open other audio/control apps.
- Only toggle 'Audio Enhancements' for THIS device exactly when asked.
If you accept this and understand the risk, type exactly:
I UNDERSTAND
"""
    resp = input(warning + "\n> ").strip()
    if resp != "I UNDERSTAND":
        return False, "Learn aborted by user (confirmation not provided)."
    print(f"Manual learn target: {name} ({flow})")
    print("Step 1: In Windows Sound settings, set 'Audio Enhancements' to ENABLED for this device.")
    input("When ready, press Enter to capture snapshot A... ")
    snapA = _collect_sysfx_snapshot(dev_id)
    print("Step 2: Now set 'Audio Enhancements' to DISABLED for the same device.")
    input("When ready, press Enter to capture snapshot B... ")
    snapB = _collect_sysfx_snapshot(dev_id)
    diffs = _diff_mmdevices_lists(snapA.get("registry") or [], snapB.get("registry") or [])
    snippet, picked = _build_vendor_ini_snippet(target, snapA, snapB, diffs)
    if not picked:
        return False, "No suitable REG_DWORD flip found under FxProperties. Driver may use non-DWORD or a different location."
    value_name    = picked["name"]
    dword_enable  = int(picked["before"])
    dword_disable = int(picked["after"])
    section_name  = _sanitize_ini_section_name(value_name)
    notes = f"Auto-learned (manual UI) on '{name}' ({flow}). A=enabled,B=disabled."
    hives = "HKCU,HKLM" if prefer_hkcu else "HKLM,HKCU"
    try:
        res = _append_vendor_ini_entry_if_missing(
            ini_path, section_name, value_name,
            dword_enable, dword_disable,
            flows="Render,Capture", hives=hives, notes=notes
        )
    except PermissionError as e:
        return False, f"Permission denied writing INI: {ini_path}. Run as Administrator. {e}"
    except OSError as e:
        return False, f"Failed to write INI: {ini_path}. {e}"
    return True, {
        "iniPath": ini_path,
        "section": section_name,
        "value_name": value_name,
        "dword_enable": dword_enable,
        "dword_disable": dword_disable
    }
def _learn_vendor_and_write_ini(target, ini_path=None):
    """
    Auto-learn a vendor DWORD toggle for target {'id','name','flow'}.
    """
    import sys
    dev_id = target["id"]
    flow   = target["flow"]
    name   = target["name"]
    ini_path = ini_path or _vendor_ini_default_path()
    # STERN WARNING + explicit confirmation
    warning = f"""
READ CAREFULLY
This automatic Learn attempt will write a vendor entry into:
  {ini_path}
It will also try to toggle Windows paths programmatically during capture. Future 'enhancements' commands WILL WRITE registry values for this device until you remove the learned section.
Do NOT change any other audio settings or devices while this runs.
Type exactly: I UNDERSTAND
"""
    resp = input(warning + "\n> ").strip()
    if resp != "I UNDERSTAND":
        return False, "Learn-auto aborted by user (confirmation not provided)."
    orig = _get_enhancements_status_propstore(dev_id)
    if orig is None:
        orig = _get_enhancements_status_com(dev_id)
        
    try:
        _set_enhancements_propstore(dev_id, True)
    except Exception:
        pass
    try:
        _set_enhancements_registry(dev_id, True, prefer_hklm=is_admin())
    except Exception:
        pass
        
    _short_settle(0.3)
    snapA = _collect_sysfx_snapshot(dev_id)
    
    try:
        _set_enhancements_propstore(dev_id, False)
    except Exception:
        pass
    try:
        _set_enhancements_registry(dev_id, False, prefer_hklm=is_admin())
    except Exception:
        pass
        
    _short_settle(0.3)
    snapB = _collect_sysfx_snapshot(dev_id)
    
    diffs = _diff_mmdevices_lists(snapA.get("registry") or [], snapB.get("registry") or [])
    snippet, picked = _build_vendor_ini_snippet(target, snapA, snapB, diffs)
    if not picked:
        return False, "No suitable REG_DWORD flip found under FxProperties. Driver may use non-DWORD or a different location."
        
    value_name = picked["name"]
    dword_enable  = int(picked["before"])
    dword_disable = int(picked["after"])
    section_name  = _sanitize_ini_section_name(value_name)
    notes = f"Auto-learned on '{name}' ({flow}). A=enabled,B=disabled."
    
    try:
        res = _append_vendor_ini_entry_if_missing(
            ini_path, section_name, value_name,
            dword_enable, dword_disable,
            flows="Render,Capture", hives="HKCU,HKLM", notes=notes
        )
    except PermissionError as e:
        return False, f"Permission denied writing INI: {ini_path}. Run as Administrator. {e}"
    except OSError as e:
        return False, f"Failed to write INI: {ini_path}. {e}"
        
    try:
        if orig is True or orig is False:
            _apply_enhancements(dev_id, flow, orig, prefer_hklm=is_admin(), allow_universal_scan=False, vendor_ini_path=ini_path)
    except Exception:
        pass
        
    return True, {
        "iniPath": ini_path,
        "section": section_name,
        "value_name": value_name,
        "dword_enable": dword_enable,
        "dword_disable": dword_disable
    }
def _get_enhancements_status_any(device_id, flow):
    """
    Best-effort read for display (GUI/labels), vendor-only:
      1) Realtek/code vendors first, then INI vendors (HKCU preferred)
      2) If no vendor applies -> None
    Returns:
      True  -> enhancements enabled
      False -> enhancements disabled
      None  -> unknown (no vendor key applies)
    """
    try:
        vend = _find_first_vendor_entry(device_id, flow, ini_path=_vendor_ini_default_path())
        if vend:
            vs = _read_vendor_entry_state(vend, device_id, flow)
            if vs is True or vs is False:
                return vs
    except Exception:
        pass
    return None
def _apply_enhancements(device_id, flow, enable, prefer_hklm=False, allow_universal_scan=False, vendor_ini_path=None):
    """
    Vendor-only policy:
      1) Try vendor toggles: INI vendors first (user-learned), then built-in code vendors (e.g. Realtek).
      2) If no vendor match, return failure (no Windows fallback).
    """
    ok_v, tag_v, state_v = _try_vendor_first(device_id, flow, enable, ini_path=vendor_ini_path)
    if ok_v:
        return True, tag_v, state_v
    return False, "no-vendor-method", None
def _enhancements_supported(device_id, flow):
    """
    Returns True if any vendor entry applies:
      - INI vendors first (user-learned), then built-in code vendors.
    Returns False otherwise. No Windows checks.
    """
    try:
        vend = _find_first_vendor_entry(device_id, flow, ini_path=_vendor_ini_default_path())
        return True if vend else False
    except Exception:
        return False
