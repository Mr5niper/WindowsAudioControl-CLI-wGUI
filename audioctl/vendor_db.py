# audioctl\vendor_db.py
import os
import re
import configparser
import time
import winreg
from .compat import is_admin
from .logging_setup import _exe_dir
from .devices import (
    _extract_endpoint_guid_from_device_id,
    _set_enhancements_registry,
    _get_enhancements_status_propstore,
    _set_enhancements_propstore,
    _get_enhancements_status_com,
    _collect_sysfx_snapshot,
    _diff_mmdevices_lists,
    _short_settle,
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
# --- Lightweight vendor DB cache (path + mtime -> parsed data) ---
_VENDOR_DB_CACHE = {
    "path": None,
    "mtime": None,
    "data": {"main": [], "fx": []},
}
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
    """
    Load vendor toggles from INI. Returns dict with 'main' and 'fx' lists.
    Uses a lightweight cache keyed by (absolute path, mtime) so we don't
    re-parse or re-fail on a missing file for every CLI call.
    """
    global _VENDOR_DB_CACHE
    # Resolve path
    path = os.path.abspath(ini_path or _vendor_ini_default_path())
    # Check file existence & mtime up front
    try:
        st = os.stat(path)
        mtime = st.st_mtime
        exists = True
    except OSError:
        exists = False
        mtime = None
    # If file does not exist, cache and return empty DB
    if not exists:
        if (_VENDOR_DB_CACHE.get("path") == path and
                _VENDOR_DB_CACHE.get("mtime") is None):
            # Already know it's missing; reuse empty DB
            return _VENDOR_DB_CACHE.get("data") or {"main": [], "fx": []}
        _VENDOR_DB_CACHE["path"] = path
        _VENDOR_DB_CACHE["mtime"] = None
        _VENDOR_DB_CACHE["data"] = {"main": [], "fx": []}
        return _VENDOR_DB_CACHE["data"]
    # If path and mtime match cache, reuse parsed DB
    if (_VENDOR_DB_CACHE.get("path") == path and
            _VENDOR_DB_CACHE.get("mtime") == mtime):
        return _VENDOR_DB_CACHE.get("data") or {"main": [], "fx": []}
    # Otherwise parse INI fresh (same logic as before)
    cfg = configparser.ConfigParser()
    entries = {"main": [], "fx": []}
    try:
        cfg.read(path, encoding="utf-8")
    except Exception:
        # On read failure, cache empty DB so we don't hammer again
        _VENDOR_DB_CACHE["path"] = path
        _VENDOR_DB_CACHE["mtime"] = mtime
        _VENDOR_DB_CACHE["data"] = entries
        return entries
    for sec in cfg.sections():
        try:
            entry_type = cfg.get(sec, "type", fallback="main").strip().lower()
            notes = cfg.get(sec, "notes", fallback="")
            if entry_type == "fx":
                # FX entry: could be single-DWORD or multi-write (existing behavior)
                fx_name = cfg.get(sec, "fx_name", fallback="").strip()
                devpat = cfg.get(sec, "device_name_pattern", fallback="").strip()  # optional in new model
                if not fx_name:
                    continue
                devices_text = cfg.get(sec, "devices", fallback="").strip()
                devices = [x.strip().lower() for x in devices_text.split(",") if x.strip()]
                e = {
                    "name": sec,
                    "type": "fx",
                    "fx_name": fx_name,
                    "device_name_pattern": devpat,
                    "notes": notes or "",
                    "devices": devices,
                }
                multi_write = cfg.get(sec, "multi_write", fallback="0").strip()
                if multi_write in ("1", "true", "yes"):
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
                    e["flows"] = [x.strip().capitalize() for x in cfg.get(sec, "flows", fallback="Render,Capture").split(",") if x.strip()]
                    e["hives"] = [x.strip().upper() for x in cfg.get(sec, "hives", fallback="HKLM,HKCU").split(",") if x.strip()]
                else:
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
                        "multi_write": False,
                    })
                # Only keep FX sections that have at least one device GUID
                if e["devices"]:
                    entries["fx"].append(e)
            else:
                # MAIN entry (supports optional subkey)
                value_name = cfg.get(sec, "value_name").strip().lower()
                en = int(cfg.get(sec, "dword_enable"))
                di = int(cfg.get(sec, "dword_disable"))
                if en not in (0, 1) or di not in (0, 1) or en == di:
                    continue
                hives = [x.strip().upper() for x in cfg.get(sec, "hives", fallback="HKLM,HKCU").split(",") if x.strip()]
                flows = [x.strip().capitalize() for x in cfg.get(sec, "flows", fallback="Render,Capture").split(",") if x.strip()]
                devices_text = cfg.get(sec, "devices", fallback="").strip()
                devices = [x.strip().lower() for x in devices_text.split(",") if x.strip()]
                subkey_txt = cfg.get(sec, "subkey", fallback="FxProperties").strip()
                subkey_norm = "Properties" if subkey_txt.lower().startswith("prop") else "FxProperties"
                entry = {
                    "name": sec,
                    "type": "main",
                    "value_name": value_name,
                    "enable": en,
                    "disable": di,
                    "hives": [h for h in hives if h in ("HKLM", "HKCU")],
                    "flows": [f for f in flows if f in ("Render", "Capture")],
                    "devices": devices,
                    "notes": notes,
                    "subkey": subkey_norm,
                }
                if entry["devices"]:
                    entries["main"].append(entry)
        except Exception:
            continue
    # Update cache with newly parsed DB
    _VENDOR_DB_CACHE["path"] = path
    _VENDOR_DB_CACHE["mtime"] = mtime
    _VENDOR_DB_CACHE["data"] = entries
    return entries
def _endpoint_fx_key(device_id, flow):
    guid = _extract_endpoint_guid_from_device_id(device_id)
    if not guid:
        return None, None
    flow_name = "Render" if str(flow).lower().startswith("r") else "Capture"
    key_path = rf"SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\{flow_name}\{guid}\FxProperties"
    return flow_name, key_path
def _guid_of(device_id):
    g = _extract_endpoint_guid_from_device_id(device_id)
    return (g or "").strip().lower()
def _vendor_entry_applies(entry, device_id, flow):
    """
    Return True if this MAIN entry applies to this endpoint AND the configured value
    exists under HKCU for the endpoint (FxProperties or Properties).
    - Checks devices membership and flows
    - HKCU only (per your environment)
    - Probes both FxProperties and Properties for value_name
    """
    guid = _extract_endpoint_guid_from_device_id(device_id)
    if not guid:
        return False
    # Device membership
    devs = set((entry.get("devices") or []))
    if not devs or guid.lower() not in {d.lower() for d in devs}:
        return False
    # Flow membership
    flow_name = "Render" if str(flow).lower().startswith("r") else "Capture"
    if entry.get("flows") and flow_name not in entry["flows"]:
        return False
    value_name = (entry.get("value_name") or "").strip().lower()
    if not value_name:
        return False
    # HKCU only; try FxProperties, then Properties
    for sub in ("FxProperties", "Properties"):
        base = _endpoint_base_path(device_id, flow, sub)
        if not base:
            continue
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, base, 0, winreg.KEY_READ) as key:
                try:
                    _ = winreg.QueryValueEx(key, value_name)
                    return True
                except FileNotFoundError:
                    pass
        except OSError:
            pass
    return False
def _set_vendor_entry_state(entry, device_id, flow, enable):
    """
    Write vendor entry DWORD to desired value across configured hives.
    Uses MAIN 'subkey' (where it came from) exactly.
    """
    subkey = (entry.get("subkey") or "FxProperties").strip()
    base = _endpoint_base_path(device_id, flow, subkey)
    if not base:
        return False
    desired = entry["enable"] if enable else entry["disable"]
    ok = False
    for h in (entry.get("hives") or ["HKCU", "HKLM"]):
        hive = winreg.HKEY_LOCAL_MACHINE if h.upper() == "HKLM" else winreg.HKEY_CURRENT_USER
        try:
            with winreg.OpenKey(hive, base, 0, winreg.KEY_SET_VALUE) as key:
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
        "devices = ",
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
    lines.append("devices = ")
    with open(ini_path, "a", encoding="utf-8", errors="replace") as f:
        f.write("\n".join(lines) + "\n")
def _read_vendor_entry_state(entry, device_id, flow):
    """
    Return True if current state equals 'enable' value, False if equals 'disable', None otherwise.
    Behavior:
      - For FX entries with multi_write=True: uses _read_decider_state (unchanged).
      - For MAIN entries and legacy single-DWORD FX entries:
          Read exactly the learned scope:
            HKCU\...\{FxProperties|Properties}\value_name for THIS endpoint,
          fallback to HKLM only if HKCU read is not present (keeps old behavior harmlessly).
    """
    # Multi-write FX: use decider logic + quorum (unchanged)
    if entry.get("type") == "fx" and entry.get("multi_write"):
        return _read_decider_state(entry, device_id, flow)

    # MAIN (enhancements) or legacy single-DWORD FX
    val_name = (entry.get("value_name") or "").strip().lower()
    if not val_name:
        return None

    # Learned subkey (where it came from); default FxProperties if missing
    subkey = (entry.get("subkey") or "FxProperties").strip()
    base = _endpoint_base_path(device_id, flow, subkey)
    if not base:
        return None

    # Hive order: prefer HKCU, then HKLM (minimal, no enumeration)
    hive_order = []
    configured = entry.get("hives") or []
    if configured:
        # preserve INI order but make sure HKCU is first if present
        seen = set()
        for h in configured:
            h_up = h.strip().upper()
            if h_up in ("HKCU", "HKLM") and h_up not in seen:
                hive_order.append(h_up); seen.add(h_up)
    if not hive_order:
        hive_order = ["HKCU", "HKLM"]

    hive_map = {
        "HKCU": winreg.HKEY_CURRENT_USER,
        "HKLM": winreg.HKEY_LOCAL_MACHINE,
    }

    for hname in hive_order:
        hive = hive_map.get(hname)
        if hive is None:
            continue
        try:
            with winreg.OpenKey(hive, base, 0, winreg.KEY_READ) as key:
                try:
                    val, typ = winreg.QueryValueEx(key, val_name)
                except FileNotFoundError:
                    continue
        except OSError:
            continue

        # Only REG_DWORD is meaningful here
        if typ != winreg.REG_DWORD:
            continue
        try:
            v = int(val)
        except Exception:
            continue

        # Accept either naming in entry
        en = entry.get("enable")
        di = entry.get("disable")
        if en is None or di is None:
            try:
                en = int(entry.get("dword_enable"))
                di = int(entry.get("dword_disable"))
            except Exception:
                en = di = None

        if en is None or di is None:
            return None

        if v == int(en):
            return True
        if v == int(di):
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
    Try MAIN vendor entries from INI first.
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
    return False, None, None
def _find_first_vendor_entry(device_id, flow, ini_path=None):
    """
    Return the first MAIN vendor entry that BOTH lists this endpoint (devices membership)
    AND actually exists under HKCU for this endpoint (FxProperties/Properties).
    If none exist (rare), fall back to first membership-only entry.
    """
    db = _load_vendor_db_split(ini_path)
    guid = _extract_endpoint_guid_from_device_id(device_id)
    if not guid:
        return None
    main_entries = db.get("main") or []
    flow_name = "Render" if str(flow).lower().startswith("r") else "Capture"
    # 1) Prefer entries that actually exist in registry for this endpoint
    for entry in main_entries:
        if _vendor_entry_applies(entry, device_id, flow_name):
            return entry
    # 2) Fallback: first membership-only match (old behavior)
    for entry in main_entries:
        try:
            devs = set((entry.get("devices") or []))
            if devs and guid.lower() in {d.lower() for d in devs}:
                if not entry.get("flows") or flow_name in entry["flows"]:
                    return entry
        except Exception:
            continue
    return None
def _entries_identical_main(a, b):
    return (a.get("value_name","").strip().lower() == b.get("value_name","").strip().lower()
            and int(a.get("enable", -999)) == int(b.get("enable", -999))
            and int(a.get("disable", -999)) == int(b.get("disable", -999)))
def _entries_identical_fx(a, b):
    # For dedupe, ignore fx_name differences; rely on writes or value_name/dwords
    if a.get("multi_write") and b.get("multi_write"):
        def _wkey(w):
            return (
                (w.get("hive") or "").upper(),
                (w.get("subkey") or "").strip().lower(),  # normalize case here
                (w.get("name") or "").strip().lower(),
                (w.get("type_enable") or "").upper(),
                (w.get("type_disable") or "").upper(),
                str(w.get("enable") or ""),
                str(w.get("disable") or "")
            )
        wa = sorted([_wkey(w) for w in (a.get("writes") or [])])
        wb = sorted([_wkey(w) for w in (b.get("writes") or [])])
        if wa != wb:
            return False
        if str(a.get("decider_index", 1)) != str(b.get("decider_index", 1)):
            return False
        try:
            qa = float(a.get("quorum_threshold", 0.60))
            qb = float(b.get("quorum_threshold", 0.60))
        except Exception:
            return False
        return abs(qa - qb) < 1e-6
    # Legacy/single-DWORD FX: compare value_name + enable/disable values
    a_val = (a.get("value_name") or "").strip().lower()
    b_val = (b.get("value_name") or "").strip().lower()
    # Accept either enable/disable or dword_enable/dword_disable keys
    def _ed_pair(entry):
        if "enable" in entry or "disable" in entry:
            return str(entry.get("enable")).strip(), str(entry.get("disable")).strip()
        return str(entry.get("dword_enable")).strip(), str(entry.get("dword_disable")).strip()
    ae, ad = _ed_pair(a)
    be, bd = _ed_pair(b)
    return (a_val == b_val) and (ae == be) and (ad == bd)
import hashlib
def _norm_write_item(w):
    # Normalize a single write item for canonical identity
    return (
        (w.get("hive") or "").upper(),
        (w.get("subkey") or "").strip().lower(),
        (w.get("name") or "").strip().lower(),
        (w.get("type_enable") or "").upper(),
        (w.get("type_disable") or "").upper(),
        str(w.get("enable") or ""),
        str(w.get("disable") or ""),
    )
def _fx_canonical_key_from_writes(writes, decider_index, quorum_threshold):
    # Build a canonical tuple for multi-write FX
    nw = sorted((_norm_write_item(w) for w in (writes or [])))
    try:
        di = int(decider_index or 1)
    except Exception:
        di = 1
    try:
        qt = float(quorum_threshold or 0.60)
    except Exception:
        qt = 0.60
    # freeze into a tuple
    return ("fx-multi", tuple(nw), di, round(qt, 6))
def _fx_canonical_key_single(value_name, enable, disable):
    # Canonical tuple for legacy/single-DWORD FX
    return ("fx-single",
            (str(value_name or "").strip().lower(),),
            int(enable), int(disable))
def _canonical_section_name_from_key(key_tuple):
    # Stable section name: fx_ + first 16 hex of sha1 over repr(key_tuple)
    h = hashlib.sha1(repr(key_tuple).encode("utf-8", "replace")).hexdigest()[:16]
    return f"fx_{h}"
# --- FAST, SINGLE-PROBE READ HELPERS (no fallbacks, no COM) ---
def _fast_read_one(hive_name: str, base_path: str, value_name: str):
    """
    Single registry read. Returns (value, type) or (None, None).
    No recursion, no alternates.
    """
    if not base_path or not value_name:
        return (None, None)
    hive = winreg.HKEY_LOCAL_MACHINE if (hive_name or "").upper() == "HKLM" else winreg.HKEY_CURRENT_USER
    try:
        with winreg.OpenKey(hive, base_path, 0, winreg.KEY_READ) as key:
            val, typ = winreg.QueryValueEx(key, value_name)
            return (val, typ)
    except OSError:
        return (None, None)
def _value_equals(expected, expected_type_name, actual_val, actual_typ):
    """
    Type-aware equality check for single-probe comparisons.
    expected_type_name is one of REG_DWORD|REG_SZ|REG_BINARY (string).
    expected is the entry text (int for dword, text for sz, 'hex:..' for binary).
    """
    tname = (expected_type_name or "").upper()
    if tname == "REG_DWORD" and actual_typ == winreg.REG_DWORD:
        try:
            return int(actual_val) == int(expected)
        except Exception:
            return False
    if tname == "REG_SZ" and actual_typ == winreg.REG_SZ:
        try:
            return str(actual_val) == str(expected)
        except Exception:
            return False
    if tname == "REG_BINARY" and actual_typ == winreg.REG_BINARY:
        try:
            exp_bytes = _parse_bin_hex(expected)
            return bytes(actual_val) == exp_bytes
        except Exception:
            return False
    # Not comparable or wrong type
    return False
def _fast_read_vendor_entry_state(entry, device_id, flow):
    """
    FAST state read (True/False/None) driven by learned scope.
    - MAIN: read exactly value_name at HKCU\...\{subkey} for THIS endpoint.
    - FX multi-write: read only the decider write (first reliable change).
    """
    try:
        if entry.get("type") == "fx" and entry.get("multi_write"):
            # Read only the decider write (first reliable change)
            writes = entry.get("writes") or []
            if not writes:
                return None
            idx = max(1, int(entry.get("decider_index", 1)))
            if idx > len(writes):
                idx = 1
            w = writes[idx - 1]
            hive_name = (w.get("hive") or "HKCU").upper()
            subkey = (w.get("subkey") or "FxProperties").strip()
            val_name = (w.get("name") or "").strip().lower()
            base = _endpoint_base_path(device_id, flow, subkey)
            if not base:
                return None
            actual_val, actual_typ = _fast_read_one(hive_name, base, val_name)
            if actual_val is None:
                return None
            # Compare using recorded enable/disable types/values
            try:
                t_en = _reg_name_to_type(w.get("type_enable"))
                t_di = _reg_name_to_type(w.get("type_disable"))
            except Exception:
                return None
            if _value_equals(w.get("enable"), w.get("type_enable"), actual_val, actual_typ):
                return True
            if _value_equals(w.get("disable"), w.get("type_disable"), actual_val, actual_typ):
                return False
            return None
        # MAIN / legacy single-DWORD
        val_name = (entry.get("value_name") or "").strip().lower()
        subkey = (entry.get("subkey") or "FxProperties").strip()
        if not val_name:
            return None
        try:
            en_val = int(entry.get("enable"))
            di_val = int(entry.get("disable"))
        except Exception:
            try:
                en_val = int(entry.get("dword_enable"))
                di_val = int(entry.get("dword_disable"))
            except Exception:
                return None
        base = _endpoint_base_path(device_id, flow, subkey)
        if not base:
            return None
        actual_val, actual_typ = _fast_read_one("HKCU", base, val_name)
        if actual_val is None or actual_typ != winreg.REG_DWORD:
            return None
        v = int(actual_val)
        if v == en_val:
            return True
        if v == di_val:
            return False
        return None
    except Exception:
        return None
def _fast_get_enhancements_state(device_id, flow):
    """
    FAST state for MAIN enhancement: find the matching MAIN entry for this device and probe it.
    Returns True/False/None.
    """
    e = _find_first_vendor_entry(device_id, flow, ini_path=_vendor_ini_default_path())
    if not e:
        return None
    return _fast_read_vendor_entry_state(e, device_id, flow)
def _append_guid_to_section(ini_path, section_name, guid_lc):
    """
    Append guid_lc to the 'devices' line of [section_name] in-place.
    - Preserves comments and ordering.
    - If devices line is missing, insert one at the end of the section.
    - If the section doesn't exist, append a new section with just devices.
    """
    try:
        with open(ini_path, "r", encoding="utf-8", errors="replace") as f:
            lines = f.read().splitlines(keepends=True)
    except FileNotFoundError:
        lines = []
    sec_hdr = f"[{section_name}]"
    sec_start = None
    sec_end = len(lines)
    # Locate section
    for i, line in enumerate(lines):
        stripped = line.strip()
        if stripped.startswith("[") and stripped.endswith("]"):
            if sec_start is None:
                if stripped.lower() == sec_hdr.lower():
                    sec_start = i
            else:
                # first header after our section -> marks end
                sec_end = i
                break
        elif sec_start is not None and sec_end == len(lines):
            # keep scanning until we see next section header
            continue
    if sec_start is None:
        # Section doesn't exist: append new section at end
        new = []
        if lines and not lines[-1].endswith(("\n", "\r")):
            new.append("\n")
        new.append(f"{sec_hdr}\n")
        new.append(f"devices = {guid_lc}\n")
        lines.extend(new)
    else:
        # Section exists: find devices= line
        devices_idx = None
        guid_set = None
        for i in range(sec_start + 1, sec_end):
            m = re.match(r"^\s*devices\s*=\s*(.*)$", lines[i], flags=re.IGNORECASE)
            if m:
                devices_idx = i
                # Parse CSV list into set (lowercased, trimmed)
                existing = [x.strip().lower() for x in m.group(1).split(",") if x.strip()]
                guid_set = set(existing)
                break
        if guid_set is None:
            guid_set = set()
        if guid_lc not in guid_set:
            guid_set.add(guid_lc)
            new_value = ",".join(sorted(guid_set))
            new_line = f"devices = {new_value}\n"
            if devices_idx is not None:
                lines[devices_idx] = new_line
            else:
                # Insert before sec_end (end of section)
                insert_at = sec_end
                # If there’s no trailing newline before next header, ensure one
                if insert_at > 0 and not lines[insert_at - 1].endswith(("\n", "\r")):
                    lines.insert(insert_at, "\n")
                    insert_at += 1
                lines.insert(insert_at, new_line)
    with open(ini_path, "w", encoding="utf-8", errors="replace") as f:
        f.writelines(lines)
def _sanitize_ini_section_name(value_name: str):
    # e.g. "{1da5d803-...},5" -> "vendor_{1da5d803-...},5"
    base = re.sub(r'[^A-Za-z0-9_,\-{}]+', "_", value_name)
    return f"vendor_{base}"
def _append_vendor_ini_entry_if_missing(ini_path, section_name, value_name, dword_enable, dword_disable,
                                        flows="Render,Capture", hives="HKCU,HKLM", notes="", subkey="FxProperties"):
    """
    Append a vendor INI section to ini_path only if it does not already exist.
    Records 'subkey' so fast reads/writes hit the exact learned spot.
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
    subkey_norm = "Properties" if str(subkey or "").strip().lower().startswith("prop") else "FxProperties"
    lines = []
    lines.append("")
    lines.append(f"[{section_name}]")
    lines.append(f"value_name = {str(value_name).strip().lower()}")
    lines.append(f"dword_enable = {int(dword_enable)}")
    lines.append(f"dword_disable = {int(dword_disable)}")
    lines.append(f"hives = {hives}")
    lines.append(f"flows = {flows}")
    lines.append(f"subkey = {subkey_norm}")  # record learned scope
    if notes:
        lines.append(f"notes = {notes}")
    lines.append("devices = ")
    text = "\n".join(lines) + "\n"
    with open(ini_path, "a", encoding="utf-8", errors="replace") as f:
        f.write(text)
    return "appended"
def _build_vendor_ini_snippet(target, snapA, snapB, diffs, section_name=None):
    """
    Build a suggested vendor INI section based on DWORD flips observed.
    Records the actual subkey (FxProperties or Properties) where the flip occurred.
    """
    cands = []
    for f in diffs.get("dword_flips", []):
        name = str(f.get("name",""))
        subkey = str(f.get("subkey",""))
        hive = str(f.get("hive",""))
        if not (name.startswith("{") and "}" in name and "," in name):
            continue
        # Keep both FxProperties and Properties; pick the first reliable change
        before = int(f.get("before"))
        after  = int(f.get("after"))
        cands.append({
            "hive": hive, "flow": None, "subkey": subkey, "name": name,
            "before": before, "after": after
        })
    if not cands:
        return None, None
    # Prefer HKCU candidates; keep original order otherwise (first reliable)
    cands.sort(key=lambda x: (0 if x["hive"] == "HKCU" else 1))
    pick = cands[0]
    dword_enable  = int(pick["before"])
    dword_disable = int(pick["after"])
    picked_subkey = pick.get("subkey") if pick.get("subkey") in ("FxProperties", "Properties") else "FxProperties"
    if not section_name:
        base = re.sub(r'[^A-Za-z0-9_,\-{}]+', "_", pick["name"])
        section_name = f"vendor_{base}"
    notes = f"Auto-learned (manual UI) on '{target.get('name')}' ({target.get('flow')}). A=enabled,B=disabled."
    snippet = []
    snippet.append(f"[{section_name}]")
    snippet.append(f"value_name = {pick['name']}")
    snippet.append(f"dword_enable = {dword_enable}")
    snippet.append(f"dword_disable = {dword_disable}")
    snippet.append("hives = HKCU,HKLM")
    snippet.append("flows = Render,Capture")
    snippet.append(f"subkey = {picked_subkey}")  # where it came from
    snippet.append(f"notes = {notes}")
    snippet.append("devices = ")
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
    confirmed = os.environ.get("AUDIOCTL_LEARN_CONFIRMED", "0") == "1"
    if not confirmed:
        resp = input(warning + "\n> ").strip()
        if resp != "I UNDERSTAND":
            return False, "Learn aborted by user (confirmation not provided)."
    else:
        print("INFO: Learn confirmation skipped via AUDIOCTL_LEARN_CONFIRMED=1")
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
    guid_lc = _guid_of(dev_id)
    # Try to dedupe into existing identical section
    db = _load_vendor_db_split(ini_path)
    candidate = {
        "type": "main",
        "value_name": value_name.strip().lower(),
        "enable": dword_enable,
        "disable": dword_disable,
    }
    for e in (db.get("main") or []):
        if _entries_identical_main(e, candidate):
            _append_guid_to_section(ini_path, e.get("name"), guid_lc)
            return True, {
                "iniPath": ini_path,
                "section": e.get("name"),
                "value_name": value_name,
                "dword_enable": dword_enable,
                "dword_disable": dword_disable
            }
    section_name  = _sanitize_ini_section_name(value_name)
    notes = f"Auto-learned (manual UI) on '{name}' ({flow}). A=enabled,B=disabled."
    hives = "HKCU,HKLM" if prefer_hkcu else "HKLM,HKCU"
    try:
        res = _append_vendor_ini_entry_if_missing(
            ini_path, section_name, value_name,
            dword_enable, dword_disable,
            flows="Render,Capture", hives=hives, notes=notes,
            subkey=(picked.get("subkey") if picked else "FxProperties")  # pass learned subkey
        )
        # Ensure devices list exists and contains this GUID
        _append_guid_to_section(ini_path, section_name, guid_lc)
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
    confirmed = os.environ.get("AUDIOCTL_LEARN_CONFIRMED", "0") == "1"
    if not confirmed:
        resp = input(warning + "\n> ").strip()
        if resp != "I UNDERSTAND":
            return False, "Learn-auto aborted by user (confirmation not provided)."
    else:
        print("INFO: Learn confirmation skipped via AUDIOCTL_LEARN_CONFIRMED=1")
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
    guid_lc = _guid_of(dev_id)
    section_name  = _sanitize_ini_section_name(value_name)
    notes = f"Auto-learned on '{name}' ({flow}). A=enabled,B=disabled."
    try:
        res = _append_vendor_ini_entry_if_missing(
            ini_path, section_name, value_name,
            dword_enable, dword_disable,
            flows="Render,Capture", hives="HKCU,HKLM", notes=notes
        )
        _append_guid_to_section(ini_path, section_name, guid_lc)
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
    Best-effort read for display (GUI/labels), vendor-only.
    Returns True/False if a vendor entry applies and can be read,
    or None if no vendor applies or status cannot be determined.
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
    """
    # Multi-write only
    if not entry.get("multi_write"):
        return None
    writes = entry.get("writes") or []
    if not writes:
        return None
    quorum_threshold = float(entry.get("quorum_threshold", 0.60))
    quorum_threshold = max(0.50, min(0.95, quorum_threshold))
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
            return None
        try:
            with winreg.OpenKey(hive, base, 0, winreg.KEY_READ) as key:
                val, typ = winreg.QueryValueEx(key, name)
        except OSError:
            return None
        try:
            t_en = _reg_name_to_type(w.get("type_enable"))
            t_di = _reg_name_to_type(w.get("type_disable"))
        except Exception:
            return None
        if _eq_expected(val, typ, w.get("enable"), t_en):
            return True
        if _eq_expected(val, typ, w.get("disable"), t_di):
            return False
        return None
    # 1) Quorum voting across all writes (recorded hive, then alternate)
    votes_true = 0
    votes_false = 0
    votes_total = 0
    for w in writes:
        rec_hive = (w.get("hive") or "").upper()
        alt_hive = "HKCU" if rec_hive == "HKLM" else "HKLM"
        s = _try_read_one(w, rec_hive)
        if s is None:
            s = _try_read_one(w, alt_hive)
        if s is True:
            votes_true += 1
            votes_total += 1
        elif s is False:
            votes_false += 1
            votes_total += 1
    if votes_total > 0:
        frac_true = votes_true / float(votes_total)
        frac_false = votes_false / float(votes_total)
        if frac_true >= quorum_threshold and frac_false < quorum_threshold:
            return True
        if frac_false >= quorum_threshold and frac_true < quorum_threshold:
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
    state = _try_read_one(decider, rec_hive)
    if state is not None:
        return state
    state = _try_read_one(decider, alt_hive)
    if state is not None:
        return state
    # 3) Best-scoring write fallback
    for w in sorted(writes, key=_score_write_for_decider, reverse=True):
        rec_hive = (w.get("hive") or "").upper()
        alt_hive = "HKCU" if rec_hive == "HKLM" else "HKLM"
        state = _try_read_one(w, rec_hive)
        if state is not None:
            return state
        state = _try_read_one(w, alt_hive)
        if state is not None:
            return state
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
    guid_lc = _guid_of(target["id"])
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
    notes = f"Learned FX '{fx_name}' for '{target['name']}' ({target['flow']}); A=enabled, B=disabled. (stability-filtered)"
    if writes:
        # Dedupe against existing multi_write entries
        db = _load_vendor_db_split(ini_path)
        candidate = {
            "type": "fx",
            "multi_write": True,
            "writes": writes,
            "decider_index": 1,
            "quorum_threshold": 0.60,
        }
        for e in (db.get("fx") or []):
            if e.get("multi_write") and _entries_identical_fx(e, candidate):
                _append_guid_to_section(ini_path, e.get("name"), guid_lc)
                return True, {
                    "iniPath": ini_path,
                    "section": e.get("name"),
                    "fx_name": fx_name,
                    "multi_write": True,
                    "write_count": len(writes),
                }
        try:
            canon_key = _fx_canonical_key_from_writes(writes, 1, 0.60)
            section_name = _canonical_section_name_from_key(canon_key)
            _append_fx_ini_entry_multi(
                ini_path, section_name, fx_name, target["name"],
                writes, notes=notes
            )
            _append_guid_to_section(ini_path, section_name, guid_lc)
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
    # Try to dedupe into an identical single-DWORD FX section
    db = _load_vendor_db_split(ini_path)
    candidate = {
        "type": "fx",
        "multi_write": False,
        "value_name": value_name.strip().lower(),
        "enable": dword_enable,
        "disable": dword_disable,
    }
    for e in (db.get("fx") or []):
        if not e.get("multi_write") and _entries_identical_fx(e, candidate):
            _append_guid_to_section(ini_path, e.get("name"), guid_lc)
            return True, {
                "iniPath": ini_path,
                "section": e.get("name"),
                "fx_name": fx_name,
                "value_name": value_name,
                "dword_enable": dword_enable,
                "dword_disable": dword_disable
            }
    try:
        canon_key = _fx_canonical_key_single(value_name, dword_enable, dword_disable)
        section_name = _canonical_section_name_from_key(canon_key)
        _append_fx_ini_entry(
            ini_path, section_name, fx_name, target["name"],
            value_name, dword_enable, dword_disable,
            flows="Render,Capture", hives=hives, notes=notes2
        )
        _append_guid_to_section(ini_path, section_name, guid_lc)
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
    """List all available FX for a device (INI-only, devices membership). Returns [{'fx_name','entry'}]."""
    db = _load_vendor_db_split(ini_path)
    guid = _extract_endpoint_guid_from_device_id(device_id)
    if not guid:
        return []
    out = []
    for entry in db.get("fx") or []:
        devs = set((entry.get("devices") or []))
        if devs and guid.lower() in {d.lower() for d in devs}:
            e = dict(entry)
            e["source"] = "ini"
            out.append({"fx_name": entry.get("fx_name"), "entry": e})
    return out
def _find_fx_for_device(device_id, flow, fx_name, ini_path=None):
    """Find FX entries matching device and effect name. INI entries only, devices membership enforced."""
    db = _load_vendor_db_split(ini_path)
    guid = _extract_endpoint_guid_from_device_id(device_id)
    if not guid:
        return []
    fx_lc = str(fx_name or "").strip().lower()
    matches = []
    for entry in db.get("fx") or []:
        devs = set((entry.get("devices") or []))
        if not devs or guid.lower() not in {d.lower() for d in devs}:
            continue
        if (entry.get("fx_name", "").strip().lower() == fx_lc):
            e = dict(entry)
            e["source"] = "ini"
            matches.append(e)
    return matches
def _apply_enhancements(device_id, flow, enable, prefer_hklm=False, allow_universal_scan=False, vendor_ini_path=None):
    """
    Vendor-only policy:
      1) Try vendor toggles: INI vendors only (per-device).
      2) If no vendor match, return failure (no Windows fallback).
    """
    db = _load_vendor_db_split(vendor_ini_path)
    guid = _extract_endpoint_guid_from_device_id(device_id)
    if not guid:
        return False, "no-vendor-method", None
    for entry in db.get("main") or []:
        try:
            devs = set((entry.get("devices") or []))
            if not devs or guid.lower() not in {d.lower() for d in devs}:
                continue
            if entry.get("flows") and ("Render" if str(flow).lower().startswith("r") else "Capture") not in entry["flows"]:
                continue
            wrote = _set_vendor_entry_state(entry, device_id, flow, enable)
            if wrote:
                ok, st = _verify_vendor_entry(entry, device_id, flow, enable, timeout=2.5, interval=0.2, consecutive=2)
                if ok:
                    return True, f"vendor:{entry['name']}", st
        except Exception:
            continue
    return False, "no-vendor-method", None
def _enhancements_supported(device_id, flow):
    """
    Returns True if any vendor entry applies:
      - INI vendors only, per-device.
    Returns False otherwise. No Windows checks.
    """
    try:
        db = _load_vendor_db_split(_vendor_ini_default_path())
        guid = _extract_endpoint_guid_from_device_id(device_id)
        if not guid:
            return False
        for entry in db.get("main") or []:
            devs = set((entry.get("devices") or []))
            if devs and guid.lower() in {d.lower() for d in devs}:
                return True
    except Exception:
        return False
    return False
def _apply_fx(device_id, flow, fx_name, enable, ini_path=None):
    """
    Toggle a learned FX effect (INI-only).
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
    verified_by = f"vendor-fx:{entry.get('fx_name','')}"
    return ok, verified_by if ok else None, state
