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
    _dump_mmdevices_all_values,
)
# --- Helpers for multi-write FX entries ---
def _reg_type_to_name(typ: int) -> str:
    if typ == winreg.REG_DWORD: return "REG_DWORD"
    if typ == winreg.REG_SZ:    return "REG_SZ"
    if typ == winreg.REG_BINARY:return "REG_BINARY"
    return f"REG_{typ}"
def _reg_name_to_type(name: str) -> int:
    nm = (name or "").strip().upper()
    if nm == "REG_DWORD":  return winreg.REG_DWORD
    if nm == "REG_SZ":     return winreg.REG_SZ
    if nm == "REG_BINARY": return winreg.REG_BINARY
    # Fallback; unsupported types will be ignored gracefully
    raise ValueError(f"Unsupported registry type: {name}")
def _format_bin_hex(data_hex_no_prefix: str) -> str:
    """Return 'hex:' form for INI readability from raw hex (no prefix)."""
    h = data_hex_no_prefix or ""
    return "hex:" + ",".join(h[i:i+2] for i in range(0, len(h), 2))
def _parse_bin_hex(text: str) -> bytes:
    """
    Accepts:
      - 'hex:aa,bb,cc' (preferred)
      - 'aabbcc' (raw hex without prefix)
    Returns bytes.
    """
    t = (text or "").strip().lower()
    if t.startswith("hex:"):
        t = t[4:]
    t = t.replace(",", "").replace(" ", "")
    if t == "":
        return b""
    return bytes.fromhex(t)
def _key_tuple(rec):
    return (str(rec.get("hive")), str(rec.get("flow")), str(rec.get("subkey")), str(rec.get("name")))
def _index_registry_list(lst):
    idx = {}
    for e in (lst or []):
        idx[_key_tuple(e)] = e
    return idx
def _build_fx_multiwrite_from_snapshots(target, snapA, snapB):
    """
    Build a comprehensive multi-write plan from two snapshots (A=enabled, B=disabled).
    Includes both FxProperties and Properties under HKCU/HKLM for the endpoint GUID.
    Each write contains enable/disable raw values and per-side types.
    Returns a list of write dicts: {
      'hive','subkey','name',
      'type_enable','type_disable',
      'enable','disable'   (value strings: int for DWORD, str for SZ, 'hex:..' for binary)
    }
    """
    writes = []
    A = _index_registry_list(snapA.get("registry") or [])
    B = _index_registry_list(snapB.get("registry") or [])
    all_keys = set(A.keys()) | set(B.keys())
    for k in sorted(all_keys):
        a = A.get(k); b = B.get(k)
        if not a or not b:
            # Changed existence (added/removed) – skip for now
            continue
        # Only consider our two subkeys
        sub = str(a.get("subkey") or "")
        if not (sub.startswith("FxProperties") or sub.startswith("Properties")):
            continue
        # Compare exact raw payloads
        type_a = a.get("type"); type_b = b.get("type")
        raw_a  = a.get("dataRaw"); raw_b  = b.get("dataRaw")
        if type_a == type_b and raw_a == raw_b:
            continue  # unchanged
        hive, flow, subkey, name = k
        def _encode_value(typ, raw):
            if typ == winreg.REG_DWORD:
                try: return int(raw)
                except Exception: return None
            if typ == winreg.REG_SZ:
                try: return str(raw)
                except Exception: return None
            if typ == winreg.REG_BINARY:
                # store as hex:aa,bb,... for readability
                return _format_bin_hex(str(raw or ""))
            # Unsupported types -> None
            return None
        v_enable = _encode_value(type_a, raw_a)
        v_disable= _encode_value(type_b, raw_b)
        if v_enable is None or v_disable is None:
            # Skip if we cannot encode (unknown type)
            continue
        writes.append({
            "hive": hive,  # "HKLM" or "HKCU"
            "subkey": subkey,  # "FxProperties" or "Properties"
            "name": name,      # "{fmtid},pid"
            "type_enable": _reg_type_to_name(type_a),
            "type_disable": _reg_type_to_name(type_b),
            "enable": v_enable,
            "disable": v_disable,
        })
    return writes
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
# Built-in FX vendor entries (used only if no learned INI entry exists for the same FX)
_CODE_FX_ENTRIES = [
    {
        "type": "fx",
        "fx_name": "Immediate mode",
        "device_name_pattern": "Microphone (Intel® Smart Sound Technology)",
        "value_name": "{4b361010-def7-43a1-a5dc-071d955b62f7},0",
        "enable": 1,
        "disable": 0,
        "hives": ["HKCU", "HKLM"],
        "flows": ["Render", "Capture"],
        "notes": "Intel SST default FX: Immediate mode"
    },
    {
        "type": "fx",
        "fx_name": "Noise Suppression",
        "device_name_pattern": "Microphone (Intel® Smart Sound Technology)",
        "value_name": "{911DFF54-0B79-4e96-B3DE-577F235B619B},0",
        "enable": 1,
        "disable": 0,
        "hives": ["HKCU", "HKLM"],
        "flows": ["Render", "Capture"],
        "notes": "Intel SST default FX: Noise Suppression"
    },
    {
        "type": "fx",
        "fx_name": "Beam Forming",
        "device_name_pattern": "Microphone (Intel® Smart Sound Technology)",
        "value_name": "{911DFF54-0B79-4e96-B3DE-577F235B619B},1",
        "enable": 1,
        "disable": 0,
        "hives": ["HKCU", "HKLM"],
        "flows": ["Render", "Capture"],
        "notes": "Intel SST default FX: Beam Forming"
    },
    {
        "type": "fx",
        "fx_name": "Acoustic Echo Cancellation",
        "device_name_pattern": "Microphone (Intel® Smart Sound Technology)",
        "value_name": "{911DFF54-0B79-4e96-B3DE-577F235B619B},2",
        "enable": 1,
        "disable": 0,
        "hives": ["HKCU", "HKLM"],
        "flows": ["Render", "Capture"],
        "notes": "Intel SST default FX: Acoustic Echo Cancellation"
    },
]
def _vendor_ini_default_path():
    """
    Return a default vendor_toggles.ini path:
    - Prefer next to the EXE (or module) if writable.
    - Otherwise, fall back to a user-writable location under %LOCALAPPDATA%\audioctl\vendor_toggles.ini.
    """
    def _is_writable_dir(path_dir):
        try:
            os.makedirs(path_dir, exist_ok=True)
            probe = os.path.join(path_dir, ".writetest")
            with open(probe, "w", encoding="utf-8") as _:
                pass
            os.remove(probe)
            return True
        except Exception:
            return False
    try:
        base = _exe_dir()
    except Exception:
        base = os.getcwd()
    preferred_dir = base
    preferred_path = os.path.join(preferred_dir, "vendor_toggles.ini")
    if _is_writable_dir(preferred_dir):
        return preferred_path
    # Fallback to user-writable location
    local = os.environ.get("LOCALAPPDATA")
    if not local:
        try:
            home = os.path.expanduser("~")
            local = os.path.join(home, "AppData", "Local")
        except Exception:
            local = None
    if not local:
        import tempfile
        local = tempfile.gettempdir()
    fallback_dir = os.path.join(local, "audioctl")
    try:
        os.makedirs(fallback_dir, exist_ok=True)
    except Exception:
        pass
    return os.path.join(fallback_dir, "vendor_toggles.ini")
def _load_vendor_db_split(ini_path=None):
    """Load vendor toggles from INI. Returns dict with 'main' and 'fx' lists."""
    path = ini_path or _vendor_ini_default_path()
    cfg = configparser.ConfigParser()
    entries = {"main": [], "fx": []}
    try:
        if not os.path.exists(path):
            return entries
        cfg.read(path, encoding="utf-8")
    except Exception:
        return entries
    for sec in cfg.sections():
        try:
            entry_type = cfg.get(sec, "type", fallback="main").strip().lower()
            notes = cfg.get(sec, "notes", fallback="")
            if entry_type == "fx":
                # FX entry: could be single-DWORD or multi-write
                fx_name = cfg.get(sec, "fx_name", fallback="").strip()
                devpat = cfg.get(sec, "device_name_pattern", fallback="").strip()
                if not fx_name or not devpat:
                    continue
                e = {
                    "name": sec,
                    "type": "fx",
                    "fx_name": fx_name,
                    "device_name_pattern": devpat,
                    "notes": notes or "",
                }
                multi_write = cfg.get(sec, "multi_write", fallback="0").strip()
                if multi_write in ("1", "true", "yes"):
                    # Multi-write form
                    write_count = int(cfg.get(sec, "write_count", fallback="0") or "0")
                    decider_index = int(cfg.get(sec, "decider_index", fallback="1") or "1")
                    quorum_text = cfg.get(sec, "quorum_threshold", fallback="0.60").strip()
                    try:
                        quorum_threshold = float(quorum_text)
                    except Exception:
                        quorum_threshold = 0.60
                    quorum_threshold = max(0.50, min(0.95, quorum_threshold))
                    if write_count <= 0:
                        continue
                    writes = []
                    for i in range(1, write_count + 1):
                        hive = cfg.get(sec, f"write{i}_hive").strip().upper()
                        subk = cfg.get(sec, f"write{i}_subkey").strip()
                        name = cfg.get(sec, f"write{i}_name").strip().lower()
                        t_en = cfg.get(sec, f"write{i}_type_enable").strip().upper()
                        t_di = cfg.get(sec, f"write{i}_type_disable").strip().upper()
                        v_en = cfg.get(sec, f"write{i}_enable").strip()
                        v_di = cfg.get(sec, f"write{i}_disable").strip()
                        writes.append({
                            "hive": hive,
                            "subkey": subk,
                            "name": name,
                            "type_enable": t_en,
                            "type_disable": t_di,
                            "enable": v_en,
                            "disable": v_di,
                        })
                    e["multi_write"] = True
                    e["writes"] = writes
                    e["decider_index"] = max(1, decider_index)
                    e["quorum_threshold"] = quorum_threshold
                    # Optional scoping for listing (not used for clicking)
                    e["flows"] = [x.strip().capitalize() for x in cfg.get(sec, "flows", fallback="Render,Capture").split(",") if x.strip()]
                    e["hives"] = [x.strip().upper() for x in cfg.get(sec, "hives", fallback="HKLM,HKCU").split(",") if x.strip()]
                else:
                    # Single-DWORD classic FX
                    value_name = cfg.get(sec, "value_name").strip().lower()
                    en = int(cfg.get(sec, "dword_enable"))
                    di = int(cfg.get(sec, "dword_disable"))
                    if en not in (0,1) or di not in (0,1) or en == di:
                        continue
                    e.update({
                        "value_name": value_name,
                        "enable": en,
                        "disable": di,
                        "hives": [x.strip().upper() for x in cfg.get(sec, "hives", fallback="HKLM,HKCU").split(",") if x.strip()],
                        "flows": [x.strip().capitalize() for x in cfg.get(sec, "flows", fallback="Render,Capture").split(",") if x.strip()],
                    })
                    e["multi_write"] = False
                entries["fx"].append(e)
            else:
                # MAIN entry (unchanged)
                value_name = cfg.get(sec, "value_name").strip().lower()
                en = int(cfg.get(sec, "dword_enable"))
                di = int(cfg.get(sec, "dword_disable"))
                if en not in (0, 1) or di not in (0, 1) or en == di:
                    continue
                hives = [x.strip().upper() for x in cfg.get(sec, "hives", fallback="HKLM,HKCU").split(",") if x.strip()]
                flows = [x.strip().capitalize() for x in cfg.get(sec, "flows", fallback="Render,Capture").split(",") if x.strip()]
                entry = {
                    "name": sec,
                    "type": "main",
                    "value_name": value_name,
                    "enable": en,
                    "disable": di,
                    "hives": [h for h in hives if h in ("HKLM", "HKCU")],
                    "flows": [f for f in flows if f in ("Render", "Capture")],
                    "notes": notes,
                }
                entries["main"].append(entry)
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
def _append_fx_ini_entry(ini_path, section_name, fx_name, device_name,
                         value_name, dword_enable, dword_disable,
                         flows, hives, notes):
    """Append FX entry to INI. Raises ValueError if section exists."""
    cfg = configparser.ConfigParser()
    try:
        if os.path.exists(ini_path):
            cfg.read(ini_path, encoding="utf-8")
    except Exception:
        pass
    if cfg.has_section(section_name):
        raise ValueError(f"Section {section_name} already exists in INI")
    # Ensure directory exists
    try:
        ini_dir = os.path.dirname(ini_path)
        if ini_dir:
            os.makedirs(ini_dir, exist_ok=True)
    except Exception:
        pass
    lines = [
        "",
        f"[{section_name}]",
        "type = fx",
        f"fx_name = {fx_name}",
        f"device_name_pattern = {device_name}",
        f"value_name = {value_name}",
        f"dword_enable = {dword_enable}",
        f"dword_disable = {dword_disable}",
        f"hives = {hives}",
        f"flows = {flows}",
        f"notes = {notes}",
    ]
    with open(ini_path, "a", encoding="utf-8", errors="replace") as f:
        f.write("\n".join(lines) + "\n")
def _append_fx_ini_entry_multi(ini_path, section_name, fx_name, device_name, writes, notes=""):
    """
    Append an FX multi-write section. Raises ValueError if section exists.
    Schema:
      [<section_name>]
      type = fx
      fx_name = <fx_name>
      device_name_pattern = <device_name>
      multi_write = 1
      write_count = N
      decider_index = 1
      quorum_threshold = 0.60
      write{i}_hive = HKLM|HKCU
      write{i}_subkey = FxProperties|Properties
      write{i}_name = {fmtid},pid
      write{i}_type_enable = REG_DWORD|REG_BINARY|REG_SZ
      write{i}_type_disable = REG_DWORD|REG_BINARY|REG_SZ
      write{i}_enable = <value>       ; int for DWORD, text for SZ, 'hex:..' for binary
      write{i}_disable = <value>
      ; optional: flows/hives to scope listing (not used for multi-write operations)
    """
    cfg = configparser.ConfigParser()
    try:
        if os.path.exists(ini_path):
            cfg.read(ini_path, encoding="utf-8")
    except Exception:
        pass
    if cfg.has_section(section_name):
        raise ValueError(f"Section {section_name} already exists in INI")
    try:
        ini_dir = os.path.dirname(ini_path)
        if ini_dir:
            os.makedirs(ini_dir, exist_ok=True)
    except Exception:
        pass
    lines = []
    lines.append("")
    lines.append(f"[{section_name}]")
    lines.append("type = fx")
    lines.append(f"fx_name = {fx_name}")
    lines.append(f"device_name_pattern = {device_name}")
    lines.append("multi_write = 1")
    lines.append(f"write_count = {len(writes)}")
    lines.append("decider_index = 1")  # default to first write as decider
    lines.append("quorum_threshold = 0.60")
    # Enumerate writes
    for i, w in enumerate(writes, 1):
        lines.append(f"write{i}_hive = {w['hive']}")
        lines.append(f"write{i}_subkey = {w['subkey']}")
        lines.append(f"write{i}_name = {w['name']}")
        lines.append(f"write{i}_type_enable = {w['type_enable']}")
        lines.append(f"write{i}_type_disable = {w['type_disable']}")
        lines.append(f"write{i}_enable = {w['enable']}")
        lines.append(f"write{i}_disable = {w['disable']}")
    if notes:
        lines.append(f"notes = {notes}")
    with open(ini_path, "a", encoding="utf-8", errors="replace") as f:
        f.write("\n".join(lines) + "\n")
def _read_vendor_entry_state(entry, device_id, flow):
    """
    Return True if current state equals 'enable' value, False if equals 'disable', None otherwise.
    Behavior:
      - For FX entries with multi_write=True:
          Uses _read_decider_state(), which:
            * Reads all recorded writes (HKCU/HKLM fallback)
            * Uses quorum_threshold (default 0.60, min 0.50 max 0.95)
            * Falls back to decider_index and then a 'best' write heuristic
      - For MAIN entries and legacy single-DWORD FX entries:
          Reads the configured DWORD under FxProperties (per hives) and
          compares to entry["enable"]/entry["disable"].
    """
    # Multi-write FX: use decider logic + quorum
    if entry.get("type") == "fx" and entry.get("multi_write"):
        return _read_decider_state(entry, device_id, flow)
    # MAIN (enhancements) or legacy single-DWORD FX
    _f, key_path = _endpoint_fx_key(device_id, flow)
    if not key_path:
        return None
    configured = entry.get("hives") or []
    # Prefer HKCU then HKLM unless caller provided a different order
    read_order = [h for h in ("HKCU", "HKLM") if h in configured] or configured
    hive_map = {
        "HKCU": winreg.HKEY_CURRENT_USER,
        "HKLM": winreg.HKEY_LOCAL_MACHINE,
    }
    for hname in read_order:
        hive = hive_map.get(hname)
        if hive is None:
            continue
        try:
            with winreg.OpenKey(hive, key_path, 0, winreg.KEY_READ) as key:
                try:
                    val, typ = winreg.QueryValueEx(key, entry["value_name"])
                except FileNotFoundError:
                    continue
        except OSError:
            continue
        # Only REG_DWORD is meaningful for these entries
        if typ != winreg.REG_DWORD:
            continue
        try:
            v = int(val)
        except Exception:
            continue
        if v == entry.get("enable"):
            return True
        if v == entry.get("disable"):
            return False
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
    Try MAIN vendor entries from INI first, then built-in code vendors.
    IMPORTANT: This function must NOT consider FX entries.
    """
    # 1) INI vendors (MAIN only)
    db = _load_vendor_db_split(ini_path)
    main_entries = db.get("main") or []
    for entry in main_entries:
        try:
            if _vendor_entry_applies(entry, device_id, flow):
                wrote = _set_vendor_entry_state(entry, device_id, flow, enable)
                if wrote:
                    ok, st = _verify_vendor_entry(entry, device_id, flow, enable, timeout=2.5, interval=0.2, consecutive=2)
                    if ok:
                        return True, f"vendor:{entry['name']}", st
        except Exception:
            continue
    # 2) Code vendors next (also MAIN by design)
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
    Read-only: return the first MAIN vendor entry that applies to this endpoint.
    FX entries are intentionally ignored here.
    """
    db = _load_vendor_db_split(ini_path)
    main_entries = db.get("main") or []
    for entry in main_entries:
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
            "hive": hive, "flow": None, "subkey": subkey, "name": name,
            "before": before, "after": after
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
def _collect_registry_samples(device_id, repeats=3, delay=0.15):
    """
    Collect several registry-only samples for the current device state to filter UI noise.
    repeats >= 1; delay in seconds between samples.
    """
    samples = []
    for i in range(max(1, int(repeats))):
        try:
            samples.append(_dump_mmdevices_all_values(device_id))
        except Exception:
            samples.append([])
        if i + 1 < repeats:
            _short_settle(delay)
    return samples
def _stable_registry_map(samples):
    """
    From a list of registry dumps (lists of rec dicts), build a stability map:
      key -> {'type': typ, 'value': dataRaw} only if the key's type and value are identical
      across ALL samples. Keys that change (type OR dataRaw) are dropped.
    Keys are tuples: (hive, flow, subkey, name).
    """
    if not samples:
        return {}
    counts = {}
    total = len(samples)
    for lst in samples:
        for rec in (lst or []):
            k = (str(rec.get("hive")), str(rec.get("flow")), str(rec.get("subkey")), str(rec.get("name")).lower())
            typ = rec.get("type")
            val = rec.get("dataRaw")
            if k not in counts:
                counts[k] = {"ok": True, "type": typ, "value": val, "seen": 1}
            else:
                same = (counts[k]["type"] == typ) and (counts[k]["value"] == val)
                if same:
                    counts[k]["seen"] += 1
                else:
                    counts[k]["ok"] = False
    out = {}
    for k, info in counts.items():
        if info["ok"] and info["seen"] == total:
            out[k] = {"type": info["type"], "value": info["value"]}
    return out
def _build_fx_multiwrite_from_stable_maps(target, stableA, stableB):
    """
    Build multi-write entries from stability-filtered maps:
    - Only include keys present in both maps
    - Type must match and value must differ
    - Encode values into INI-friendly forms (int/str/'hex:..')
    - Order results so the first write (decider) is the most stable-looking signal
    """
    writes = []
    both = set(stableA.keys()) & set(stableB.keys())
    for k in sorted(both):
        ta = stableA[k]["type"]; tb = stableB[k]["type"]
        va = stableA[k]["value"]; vb = stableB[k]["value"]
        if ta != tb or va == vb:
            continue
        hive, flow, subkey, name = k
        def _encode(typ, raw):
            if typ == winreg.REG_DWORD:
                try: return int(raw)
                except Exception: return None
            if typ == winreg.REG_SZ:
                try: return str(raw)
                except Exception: return None
            if typ == winreg.REG_BINARY:
                return _format_bin_hex(str(raw or ""))
            return None
        en = _encode(ta, va)
        di = _encode(tb, vb)
        if en is None or di is None:
            continue
        writes.append({
            "hive": hive,
            "subkey": subkey,
            "name": name,
            "type_enable": _reg_type_to_name(ta),
            "type_disable": _reg_type_to_name(tb),
            "enable": en,
            "disable": di,
        })
    # Prefer stable indicators first => decider_index=1 picks the best
    def _score(w):
        score = 0
        if str(w.get("subkey") or "").startswith("FxProperties"):
            score += 3
        te = (w["type_enable"] or "").upper()
        td = (w["type_disable"] or "").upper()
        if te == "REG_DWORD" and td == "REG_DWORD":
            score += 3
            try:
                if {int(w["enable"]), int(w["disable"])} == {0, 1}:
                    score += 2
            except Exception:
                pass
        if te == "REG_BINARY" and td == "REG_BINARY":
            score -= 1
        return score
    writes.sort(key=_score, reverse=True)
    return writes
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
      1) INI vendors first (user-learned), then built-in code vendors
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
def _endpoint_base_path(device_id, flow, subkey):
    guid = _extract_endpoint_guid_from_device_id(device_id)
    if not guid:
        return None
    flow_name = "Render" if str(flow).lower().startswith("r") else "Capture"
    return rf"SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\{flow_name}\{guid}\{subkey}"
def _perform_multi_writes(entry, device_id, flow, enable):
    """
    Write all configured values for enable/disable.
    Returns True if ALL writes succeeded; False otherwise.
    """
    ok_all = True
    for w in entry.get("writes") or []:
        hive_name = w["hive"].upper()
        hive = winreg.HKEY_LOCAL_MACHINE if hive_name == "HKLM" else winreg.HKEY_CURRENT_USER
        subk = w["subkey"]
        name = w["name"]
        base = _endpoint_base_path(device_id, flow, subk)
        if not base:
            ok_all = False
            continue
        # Choose per-side type and value
        tname = w["type_enable"] if enable else w["type_disable"]
        try:
            typ = _reg_name_to_type(tname)
        except Exception:
            ok_all = False
            continue
        val_text = w["enable"] if enable else w["disable"]
        # Parse value to native Python type
        try:
            if typ == winreg.REG_DWORD:
                data = int(val_text)
            elif typ == winreg.REG_SZ:
                data = str(val_text)
            elif typ == winreg.REG_BINARY:
                data = _parse_bin_hex(val_text)
            else:
                ok_all = False
                continue
        except Exception:
            ok_all = False
            continue
        try:
            with winreg.OpenKey(hive, base, 0, winreg.KEY_SET_VALUE) as key:
                if typ == winreg.REG_BINARY:
                    winreg.SetValueEx(key, name, 0, winreg.REG_BINARY, data)
                else:
                    winreg.SetValueEx(key, name, 0, typ, data)
        except OSError:
            ok_all = False
            continue
    return ok_all
def _read_decider_state(entry, device_id, flow):
    """
    Read current FX state based on the configured decider write item.
    Returns True (enabled) / False (disabled) / None (unknown).
    Uses:
      - quorum_threshold across all writes (default 0.60)
      - recorded decider_index
      - fallback 'best' write heuristic
    This version includes debug prints so we can see why state is Unknown.
    """
    # Multi-write only
    if not entry.get("multi_write"):
        print("DEBUG: _read_decider_state called on non-multi_write entry")
        return None
    writes = entry.get("writes") or []
    if not writes:
        print("DEBUG: _read_decider_state: no writes for entry", entry.get("fx_name"))
        return None
    fx_name = entry.get("fx_name", "<unknown FX>")
    quorum_threshold = float(entry.get("quorum_threshold", 0.60))
    quorum_threshold = max(0.50, min(0.95, quorum_threshold))
    print(f"DEBUG: _read_decider_state for FX '{fx_name}', quorum={quorum_threshold}, writes={len(writes)}")
    def _eq_expected(cur_val, cur_typ, exp_text, exp_typ):
        if cur_typ != exp_typ:
            return False
        try:
            if exp_typ == winreg.REG_DWORD:
                return int(cur_val) == int(exp_text)
            if exp_typ == winreg.REG_SZ:
                return str(cur_val) == str(exp_text)
            if exp_typ == winreg.REG_BINARY:
                return bytes(cur_val) == _parse_bin_hex(exp_text)
        except Exception:
            return False
        return False
    def _try_read_one(w, hive_name):
        hive = winreg.HKEY_LOCAL_MACHINE if hive_name == "HKLM" else winreg.HKEY_CURRENT_USER
        subk = (w.get("subkey") or "").strip()
        name = (w.get("name") or "").strip().lower()
        base = _endpoint_base_path(device_id, flow, subk)
        if not base:
            print(f"DEBUG: _try_read_one: no base path for subkey={subk}")
            return None
        try:
            with winreg.OpenKey(hive, base, 0, winreg.KEY_READ) as key:
                val, typ = winreg.QueryValueEx(key, name)
        except OSError as e:
            print(f"DEBUG: _try_read_one: OpenKey/QueryValueEx failed hive={hive_name} base={base} name={name} err={e}")
            return None
        print(f"DEBUG: _try_read_one hive={hive_name} base={base} name={name} type={typ} len={len(val) if isinstance(val, (bytes, bytearray)) else 'n/a'}")
        try:
            t_en = _reg_name_to_type(w.get("type_enable"))
            t_di = _reg_name_to_type(w.get("type_disable"))
        except Exception as e:
            print(f"DEBUG: _try_read_one: reg_type_to_name error: {e}")
            return None
        if _eq_expected(val, typ, w.get("enable"), t_en):
            print("DEBUG: _try_read_one -> matches ENABLE")
            return True
        if _eq_expected(val, typ, w.get("disable"), t_di):
            print("DEBUG: _try_read_one -> matches DISABLE")
            return False
        print("DEBUG: _try_read_one -> no match enable/disable")
        return None
    # 1) Quorum voting across all writes (recorded hive, then alternate)
    votes_true = 0
    votes_false = 0
    votes_total = 0
    for idx, w in enumerate(writes, 1):
        rec_hive = (w.get("hive") or "").upper()
        alt_hive = "HKCU" if rec_hive == "HKLM" else "HKLM"
        print(f"DEBUG: quorum pass write#{idx}, rec_hive={rec_hive}, alt_hive={alt_hive}")
        s = _try_read_one(w, rec_hive)
        if s is None:
            s = _try_read_one(w, alt_hive)
        if s is True:
            votes_true += 1
            votes_total += 1
        elif s is False:
            votes_false += 1
            votes_total += 1
    print(f"DEBUG: quorum votes: total={votes_total} true={votes_true} false={votes_false}")
    if votes_total > 0:
        frac_true = votes_true / float(votes_total)
        frac_false = votes_false / float(votes_total)
        print(f"DEBUG: quorum fractions: frac_true={frac_true:.2f} frac_false={frac_false:.2f}")
        if frac_true >= quorum_threshold and frac_false < quorum_threshold:
            print("DEBUG: quorum -> ENABLED")
            return True
        if frac_false >= quorum_threshold and frac_true < quorum_threshold:
            print("DEBUG: quorum -> DISABLED")
            return False
    # 2) Decider index fallback
    def _score_write_for_decider(w):
        score = 0
        if str((w.get("subkey") or "").strip()).startswith("FxProperties"):
            score += 10
        t_en = (w.get("type_enable") or "").upper()
        t_di = (w.get("type_disable") or "").upper()
        if t_en == "REG_DWORD" and t_di == "REG_DWORD":
            score += 5
            try:
                en_v = int(w.get("enable"))
                di_v = int(w.get("disable"))
                if {en_v, di_v} == {0, 1}:
                    score += 2
            except Exception:
                pass
        return score
    idx = max(1, int(entry.get("decider_index", 1)))
    if idx > len(writes):
        idx = 1
    decider = writes[idx - 1]
    rec_hive = (decider.get("hive") or "").upper()
    alt_hive = "HKCU" if rec_hive == "HKLM" else "HKLM"
    print(f"DEBUG: decider fallback idx={idx}, rec_hive={rec_hive}, alt_hive={alt_hive}")
    state = _try_read_one(decider, rec_hive)
    if state is not None:
        print(f"DEBUG: decider(rec) -> {state}")
        return state
    state = _try_read_one(decider, alt_hive)
    if state is not None:
        print(f"DEBUG: decider(alt) -> {state}")
        return state
    # 3) Best-scoring write fallback
    print("DEBUG: best-write fallback scan")
    for w in sorted(writes, key=_score_write_for_decider, reverse=True):
        rec_hive = (w.get("hive") or "").upper()
        alt_hive = "HKCU" if rec_hive == "HKLM" else "HKLM"
        state = _try_read_one(w, rec_hive)
        if state is not None:
            print(f"DEBUG: best-write(rec) -> {state}")
            return state
        state = _try_read_one(w, alt_hive)
        if state is not None:
            print(f"DEBUG: best-write(alt) -> {state}")
            return state
    print("DEBUG: _read_decider_state -> UNKNOWN")
    return None
def _dump_mmdevices_all_values_for_fx_learn(device_id):
    # Deprecated in favor of _dump_mmdevices_all_values from devices.py
    return _dump_mmdevices_all_values(device_id)
def _learn_fx_and_write_ini(target, fx_name, snapA, snapB, ini_path=None, prefer_hkcu=True):
    """
    Pure logic: Learn an FX toggle from two pre-captured snapshots.
    A = enabled, B = disabled.
    Writes a multi_write FX section capturing ALL changed values across Properties and FxProperties.
    Falls back to single-DWORD flip if nothing else changes.
    """
    import re
    ini_path = ini_path or _vendor_ini_default_path()
    # Build stability-filtered maps
    try:
        stableA = _stable_registry_map([snapA.get("registry") or []])
    except Exception as e:
        return False, f"Failed to process snapshot A: {e}"
    try:
        samplesB = _collect_registry_samples(target["id"], repeats=3, delay=0.18)
        samplesB.insert(0, snapB.get("registry") or [])
        stableB = _stable_registry_map(samplesB)
    except Exception as e:
        return False, f"Failed to process snapshot B: {e}"
    writes = _build_fx_multiwrite_from_stable_maps(target, stableA, stableB)
    safe_device_name = re.sub(r'[^A-Za-z0-9_\- ]+', '_', target["name"])
    safe_fx_name = re.sub(r'[^A-Za-z0-9_\-]+', '_', fx_name)
    section_name = f"{safe_device_name}::{safe_fx_name}"
    notes = f"Learned FX '{fx_name}' for '{target['name']}' ({target['flow']}); A=enabled, B=disabled. (stability-filtered)"
    if writes:
        try:
            _append_fx_ini_entry_multi(
                ini_path, section_name, fx_name, target["name"],
                writes, notes=notes
            )
            return True, {
                "iniPath": ini_path,
                "section": section_name,
                "fx_name": fx_name,
                "multi_write": True,
                "write_count": len(writes),
            }
        except ValueError as e:
            return False, str(e)
        except Exception as e:
            return False, f"Failed to write INI: {e}"
    # Fallback to legacy DWORD flip (previous behavior)
    try:
        diffs = _diff_mmdevices_lists(snapA.get("registry") or [], snapB.get("registry") or [])
    except Exception as e:
        return False, f"Diff failed: {e}"
    snippet, picked = _build_vendor_ini_snippet(target, snapA, snapB, diffs)
    if not picked:
        return False, "No suitable registry differences found to learn."
    value_name = picked["name"]
    dword_enable = int(picked["before"])
    dword_disable = int(picked["after"])
    notes2 = notes + " (single DWORD)"
    hives = "HKCU,HKLM" if prefer_hkcu else "HKLM,HKCU"
    try:
        _append_fx_ini_entry(
            ini_path, section_name, fx_name, target["name"],
            value_name, dword_enable, dword_disable,
            flows="Render,Capture", hives=hives, notes=notes2
        )
    except ValueError as e:
        return False, str(e)
    except Exception as e:
        return False, f"Failed to write INI: {e}"
    return True, {
        "iniPath": ini_path,
        "section": section_name,
        "fx_name": fx_name,
        "value_name": value_name,
        "dword_enable": dword_enable,
        "dword_disable": dword_disable
    }
def _list_fx_for_device(device_id, flow, ini_path=None):
    """List all available FX for a device (learned INI first, then code defaults). Returns [{'fx_name','entry'}]."""
    from .devices import list_devices
    db = _load_vendor_db_split(ini_path)
    devices = list_devices(include_all=False)
    device = next((d for d in devices if d["id"] == device_id), None)
    if not device:
        return []
    device_name = device["name"]
    flow_name = "Render" if str(flow).lower().startswith("r") else "Capture"
    available_fx = []
    seen = set()  # fx_name lowercase; learned wins
    # 1) Learned FX
    for entry in db["fx"]:
        if flow_name not in (entry.get("flows") or [flow_name]):
            continue
        pattern = entry.get("device_name_pattern", "")
        if pattern and pattern.lower() in device_name.lower():
            fx_name = (entry.get("fx_name") or "").strip()
            if fx_name and fx_name.lower() not in seen:
                seen.add(fx_name.lower())
                e = dict(entry)
                e["source"] = "ini"
                available_fx.append({
                    "fx_name": fx_name,
                    "entry": e
                })
    # 2) Code FX (only add if not already provided by learned)
    for entry in _CODE_FX_ENTRIES:
        if flow_name not in (entry.get("flows") or []):
            continue
        pattern = entry.get("device_name_pattern", "")
        if pattern and pattern.lower() in device_name.lower():
            fx_name = (entry.get("fx_name") or "").strip()
            if fx_name and fx_name.lower() not in seen:
                seen.add(fx_name.lower())
                e = dict(entry)
                e["source"] = "code"
                available_fx.append({
                    "fx_name": fx_name,
                    "entry": e
                })
    return available_fx
def _find_fx_for_device(device_id, flow, fx_name, ini_path=None):
    """Find FX entries matching device and effect name. Learned INI entries first, then built-in code entries."""
    from .devices import list_devices
    db = _load_vendor_db_split(ini_path)
    matches = []
    devices = list_devices(include_all=False)
    device = next((d for d in devices if d["id"] == device_id), None)
    if not device:
        return []
    device_name = device["name"]
    flow_name = "Render" if str(flow).lower().startswith("r") else "Capture"
    fx_lc = str(fx_name or "").strip().lower()
    # 1) Learned INI FX (highest precedence)
    for entry in db["fx"]:
        if (entry.get("fx_name", "").strip().lower() == fx_lc and
            flow_name in (entry.get("flows") or [flow_name])):  # default allow both if not provided
            pattern = entry.get("device_name_pattern", "")
            if pattern and pattern.lower() in device_name.lower():
                e = dict(entry)
                e["source"] = "ini"
                matches.append(e)
    # 2) Built-in FX defaults (code), appended after learned so learned wins
    for entry in _CODE_FX_ENTRIES:
        if (entry.get("fx_name", "").strip().lower() == fx_lc and
            flow_name in (entry.get("flows") or [])):
            pattern = entry.get("device_name_pattern", "")
            if pattern and pattern.lower() in device_name.lower():
                e = dict(entry)
                e["source"] = "code"
                matches.append(e)
    return matches
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
def _apply_fx(device_id, flow, fx_name, enable, ini_path=None):
    """
    Toggle a learned or built-in FX effect.
    - If the FX entry is multi_write, apply all configured writes for the target state.
    - Otherwise, legacy single-DWORD behavior.
    Returns (success, verified_by, final_state).
    """
    entries = _find_fx_for_device(device_id, flow, fx_name, ini_path)
    if not entries:
        return False, None, None
    entry = entries[0]
    if entry.get("multi_write"):
        wrote_all = _perform_multi_writes(entry, device_id, flow, enable)
        if not wrote_all:
            return False, None, None
        st = _read_decider_state(entry, device_id, flow)
        verified_by = f"vendor-fx:multi:{entry.get('fx_name','')}"
        return (st is not None and st == bool(enable)), verified_by if st is not None else None, st
    # Legacy single-DWORD FX
    wrote = _set_vendor_entry_state(entry, device_id, flow, enable)
    if not wrote:
        return False, None, None
    ok, state = _verify_vendor_entry(entry, device_id, flow, enable,
                                     timeout=2.5, interval=0.2, consecutive=2)
    src = entry.get("source", "ini")
    verified_by = f"vendor-fx:{'code:' if src=='code' else ''}{entry.get('fx_name','')}"
    return ok, verified_by if ok else None, state
