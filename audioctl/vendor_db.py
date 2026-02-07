# audioctl/vendor_db.py
#
# Vendor toggles / vendor_toggles.ini control layer
# -------------------------------------------------
# This module is the INI-driven layer that decides how we *actually* toggle:
#   - The main "Audio Enhancements" switch (vendor-only at runtime)
#   - Per-effect "FX" toggles (learned effects), which may be:
#       * legacy single-DWORD toggles, or
#       * multi_write toggles (multiple registry values/types written together)
#
# It also owns:
#   - INI parsing + caching keyed by (absolute path, mtime)
#   - learn flows that derive toggle rules from registry snapshots
#   - fast (no COM) read helpers used by GUI polling
#
# devices.py provides:
#   - endpoint GUID extraction from device IDs
#   - raw MMDevices snapshot collection and diff helpers
# This module uses those raw snapshots/diffs to decide what to write/read for a
# specific driver/device, and persists that decision into vendor_toggles.ini.
import os
import re
import configparser
import time
import winreg
import hashlib
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
# Registry encoding helpers:
#   INI stores types symbolically (REG_DWORD/REG_SZ/REG_BINARY) and stores values in
#   an INI-friendly representation:
#     - DWORD: integer
#     - SZ:    string
#     - BINARY: "hex:aa,bb,cc" (readable, diffable, and lossless)
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
    # REG_BINARY values are stored in INI in "hex:aa,bb,cc" form (human-readable but exact).
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

# --- Device-name -> GUID bucket mapping (for INI readability; case-insensitive) ---
def _canon_device_name(name: str) -> str:
    """Canonicalize a friendly name for bucketing (case-insensitive)."""
    try:
        return (name or "").strip().casefold()
    except Exception:
        return (name or "").strip().lower()

def _name_bucket_id(name: str) -> str:
    """
    Stable bucket id derived from canonicalized name.
    Used to generate INI keys:
      name_<id>  = <original device name>
      guids_<id> = {guid1},{guid2}
    """
    canon = _canon_device_name(name)
    h = hashlib.sha1(canon.encode("utf-8", "replace")).hexdigest()[:8]
    return h

def _append_guid_to_name_bucket(ini_path: str, section_name: str, device_name: str, guid_lc: str):
    """
    Maintain per-section device name buckets:
      name_<id>  = <device_name>
      guids_<id> = {guid},{guid2}
    - bucket id derived from canonicalized device_name (case-insensitive)
    - one bucket can contain multiple GUIDs (same name reused across endpoints)
    - does not remove anything; only adds/updates in-place
    """
    if not ini_path or not section_name or not device_name or not guid_lc:
        return
    bid = _name_bucket_id(device_name)
    key_name = f"name_{bid}"
    key_guids = f"guids_{bid}"
    sec_hdr = f"[{section_name}]"

    try:
        with open(ini_path, "r", encoding="utf-8", errors="replace") as f:
            lines = f.read().splitlines(keepends=True)
    except FileNotFoundError:
        lines = []

    # Locate section
    sec_start = None
    sec_end = len(lines)
    for i, line in enumerate(lines):
        s = line.strip()
        if s.startswith("[") and s.endswith("]"):
            if sec_start is None:
                if s.lower() == sec_hdr.lower():
                    sec_start = i
            else:
                sec_end = i
                break

    if sec_start is None:
        # Section missing: create minimal section (best-effort)
        new = []
        if lines and not lines[-1].endswith(("\n", "\r")):
            new.append("\n")
        new.append(f"{sec_hdr}\n")
        new.append(f"{key_name} = {device_name}\n")
        new.append(f"{key_guids} = {guid_lc}\n")
        lines.extend(new)
        with open(ini_path, "w", encoding="utf-8", errors="replace") as f:
            f.writelines(lines)
        return

    # Find existing name_<id> and guids_<id>
    name_idx = None
    guids_idx = None
    existing_guids = None

    re_name = re.compile(rf"^\s*{re.escape(key_name)}\s*=\s*(.*)$", re.IGNORECASE)
    re_guids = re.compile(rf"^\s*{re.escape(key_guids)}\s*=\s*(.*)$", re.IGNORECASE)

    for i in range(sec_start + 1, sec_end):
        m = re_name.match(lines[i])
        if m:
            name_idx = i
            # keep first-seen display name; do not overwrite (human readability)
            break

    for i in range(sec_start + 1, sec_end):
        m = re_guids.match(lines[i])
        if m:
            guids_idx = i
            txt = (m.group(1) or "").strip()
            existing_guids = [x.strip().lower() for x in txt.split(",") if x.strip()]
            break

    # Ensure name_<id> exists
    if name_idx is None:
        insert_at = sec_end
        if insert_at > 0 and not lines[insert_at - 1].endswith(("\n", "\r")):
            lines.insert(insert_at, "\n")
            insert_at += 1
        lines.insert(insert_at, f"{key_name} = {device_name}\n")
        sec_end += 1
        if insert_at <= sec_end:
            sec_end += 0

    # Ensure guids_<id> contains guid
    if existing_guids is None:
        # create guids line
        insert_at = sec_end
        if insert_at > 0 and not lines[insert_at - 1].endswith(("\n", "\r")):
            lines.insert(insert_at, "\n")
            insert_at += 1
        lines.insert(insert_at, f"{key_guids} = {guid_lc}\n")
    else:
        if guid_lc.lower() not in {g.lower() for g in existing_guids}:
            existing_guids.append(guid_lc.lower())
            new_val = ",".join(sorted(set(existing_guids)))
            new_line = f"{key_guids} = {new_val}\n"
            if guids_idx is not None:
                lines[guids_idx] = new_line
            else:
                insert_at = sec_end
                if insert_at > 0 and not lines[insert_at - 1].endswith(("\n", "\r")):
                    lines.insert(insert_at, "\n")
                    insert_at += 1
                lines.insert(insert_at, new_line)

    with open(ini_path, "w", encoding="utf-8", errors="replace") as f:
        f.writelines(lines)

# --- Heuristic FX matching helpers (pattern + registry signature) ---
def _fx_pattern_match(entry: dict, device_name: str) -> bool:
    """
    Regex match against entry['device_name_pattern'] (case-insensitive).
    Returns False if no pattern, no name, or invalid regex.
    """
    pat = (entry.get("device_name_pattern") or "").strip()
    if not pat or not device_name:
        return False
    try:
        return re.search(pat, device_name, re.IGNORECASE) is not None
    except Exception:
        return False

def _fx_signature_matches_legacy(entry: dict, device_id: str, flow: str) -> bool:
    """
    Legacy (single DWORD) FX spoof verification:
      - read current DWORD from registry (prefer entry subkey if present, else try both)
      - confirm current value equals enable or disable payload
    """
    val_name = (entry.get("value_name") or "").strip().lower()
    if not val_name:
        return False
    try:
        en_val = int(entry.get("enable", entry.get("dword_enable")))
        di_val = int(entry.get("disable", entry.get("dword_disable")))
    except Exception:
        return False
    # Determine which subkeys to probe (try learned subkey first if present)
    subkeys = []
    subkey = (entry.get("subkey") or "").strip()
    if subkey:
        subkeys.append("Properties" if subkey.lower().startswith("prop") else "FxProperties")
    subkeys.extend(["FxProperties", "Properties"])
    seen = set()
    for subk in subkeys:
        if subk in seen:
            continue
        seen.add(subk)
        base = _endpoint_base_path(device_id, flow, subk)
        if not base:
            continue
        # Read both hives cheaply; if both readable and disagree, prefer newest key.
        cu_val, cu_typ = _fast_read_one("HKCU", base, val_name)
        lm_val, lm_typ = _fast_read_one("HKLM", base, val_name)
        cu_state = None
        lm_state = None
        if cu_typ == winreg.REG_DWORD:
            try:
                v = int(cu_val)
                if v == en_val or v == di_val:
                    cu_state = v
            except Exception:
                cu_state = None
        if lm_typ == winreg.REG_DWORD:
            try:
                v = int(lm_val)
                if v == en_val or v == di_val:
                    lm_state = v
            except Exception:
                lm_state = None
        if cu_state is not None and lm_state is None:
            return True
        if lm_state is not None and cu_state is None:
            return True
        if cu_state is not None and lm_state is not None:
            if cu_state == lm_state:
                return True
            # Disagree: prefer most recently written key (same heuristic as fast reads)
            tcu = _fast_key_lastwrite("HKCU", base)
            tlm = _fast_key_lastwrite("HKLM", base)
            if isinstance(tcu, int) and isinstance(tlm, int):
                return True  # both are valid signatures; whichever is newest is "real"
            return True
    return False

def _fx_signature_matches_multi(entry: dict, device_id: str, flow: str) -> bool:
    """
    Multi-write FX spoof verification:
      - for applicable writes (universal + write{i}_devices for this guid),
        check that current registry value matches either enable or disable payload
        with the correct type.
      - require quorum_threshold fraction to match to avoid ghost FX.
    """
    writes_all = entry.get("writes") or []
    if not writes_all:
        return False
    guid_lc = _guid_of(device_id)
    writes = [w for w in writes_all if _write_applies_to_guid(w, guid_lc)]
    if not writes:
        return False
    try:
        qt = float(entry.get("quorum_threshold", 0.60))
    except Exception:
        qt = 0.60
    qt = max(0.50, min(0.95, qt))
    ok = 0
    total = 0
    for w in writes:
        subk = (w.get("subkey") or "").strip() or "FxProperties"
        name = (w.get("name") or "").strip().lower()
        if not name:
            continue
        base = _endpoint_base_path(device_id, flow, subk)
        if not base:
            continue
        total += 1
        # Read both hives; accept match in either.
        cu_val, cu_typ = _fast_read_one("HKCU", base, name)
        lm_val, lm_typ = _fast_read_one("HKLM", base, name)
        # enable signature
        if _value_equals(w.get("enable"), w.get("type_enable"), cu_val, cu_typ) or \
           _value_equals(w.get("enable"), w.get("type_enable"), lm_val, lm_typ):
            ok += 1
            continue
        # disable signature
        if _value_equals(w.get("disable"), w.get("type_disable"), cu_val, cu_typ) or \
           _value_equals(w.get("disable"), w.get("type_disable"), lm_val, lm_typ):
            ok += 1
            continue
    if total <= 0:
        return False
    return (ok / float(total)) >= qt

def _fx_entry_spoof_applies(entry: dict, device_id: str, flow: str, device_name: str) -> bool:
    """
    True if:
      - device_name_pattern matches device_name AND
      - registry signature indicates this entry is likely applicable to this endpoint.
    """
    if not _fx_pattern_match(entry, device_name):
        return False
    if entry.get("multi_write"):
        return _fx_signature_matches_multi(entry, device_id, flow)
    return _fx_signature_matches_legacy(entry, device_id, flow)

def _fx_entry_signature_applies(entry: dict, device_id: str, flow: str) -> bool:
    """
    Signature-only applicability check for THIS device_id/flow (live registry read).
    This is the truth check. It does not use device name or GUID membership.
    """
    try:
        if entry.get("multi_write"):
            return _fx_signature_matches_multi(entry, device_id, flow)
        return _fx_signature_matches_legacy(entry, device_id, flow)
    except Exception:
        return False

def _main_entry_signature_applies(entry: dict, device_id: str, flow: str) -> bool:
    """
    Signature-only applicability check for MAIN enhancements toggle (live registry read).
    Requires that the entry's value_name exists for this endpoint and equals either
    enable or disable value.
    """
    try:
        val_name = (entry.get("value_name") or "").strip().lower()
        if not val_name:
            return False
        # learned subkey (FxProperties vs Properties)
        subkey = (entry.get("subkey") or "FxProperties").strip()
        base = _endpoint_base_path(device_id, flow, subkey)
        if not base:
            return False
        try:
            en_val = int(entry.get("enable", entry.get("dword_enable")))
            di_val = int(entry.get("disable", entry.get("dword_disable")))
        except Exception:
            return False
        # Read HKCU then HKLM (same policy as _read_vendor_entry_state)
        for hive_name in ("HKCU", "HKLM"):
            hive = winreg.HKEY_CURRENT_USER if hive_name == "HKCU" else winreg.HKEY_LOCAL_MACHINE
            try:
                with winreg.OpenKey(hive, base, 0, winreg.KEY_READ) as key:
                    val, typ = winreg.QueryValueEx(key, val_name)
            except OSError:
                continue
            if typ != winreg.REG_DWORD:
                continue
            try:
                v = int(val)
            except Exception:
                continue
            if v == en_val or v == di_val:
                return True
        return False
    except Exception:
        return False

def _fx_candidate_by_guid_or_pattern(entry: dict, guid_lc: str, device_name: str) -> bool:
    """
    Fast candidate filter:
      - GUID is in entry.devices OR
      - device_name_pattern matches (if device_name provided and pattern present)
    This is NOT truth; it only narrows candidates. Truth is signature check.
    """
    try:
        devs = {d.lower() for d in (entry.get("devices") or [])}
        if guid_lc and devs and guid_lc in devs:
            return True
        pat = (entry.get("device_name_pattern") or "").strip()
        if device_name and pat:
            try:
                return re.search(pat, device_name, re.IGNORECASE) is not None
            except Exception:
                return False
        return False
    except Exception:
        return False

def _key_tuple(rec):
    # Snapshot records are keyed by (hive, flow, subkey, name). That identity is
    # stable across snapshots and is what we diff when learning.
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
    Why multi-write exists:
      Some drivers toggle an effect by changing multiple keys and types (including REG_BINARY
      blobs). Reproducing the UI behavior requires replaying the set of writes, not just one
      DWORD.
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
# Cache key:
#   - absolute INI path
#   - os.stat().st_mtime
#
# Why:
#   - GUI calls "fast" read helpers frequently; re-parsing the INI each time is wasteful.
#   - When the file is missing, we cache mtime=None so we don't repeatedly hit the filesystem.
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
    Why:
      The EXE directory may be under Program Files (not writable without elevation).
      Learn flows need to append/update the INI, so we prefer a per-user writable location
      when the EXE directory is not writable.
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
    r"""
    Load vendor toggles from INI. Returns dict with 'main' and 'fx' lists.
    Uses a lightweight cache keyed by (absolute path, mtime) so we don't
    re-parse or re-fail on a missing file for every CLI call.
    INI schema summary (debugger-oriented):
      MAIN entry (type default is main):
        - [section]
        - value_name = "{fmtid},pid" (stored lowercase internally)
        - dword_enable / dword_disable (usually 0/1)
        - hives = HKCU,HKLM (ordering matters for write preference; HKLM may require admin)
        - flows = Render,Capture
        - subkey = FxProperties or Properties (learned location)
        - devices = comma-separated endpoint GUIDs (required)
      FX entry (legacy single DWORD):
        - type = fx
        - fx_name
        - value_name / dword_enable / dword_disable
        - hives / flows (optional filters)
        - devices list required
      FX entry (multi_write):
        - type = fx
        - fx_name
        - multi_write = 1
        - write_count = N
        - write{i}_hive = HKCU|HKLM
        - write{i}_subkey = FxProperties|Properties
        - write{i}_name = "{fmtid},pid"
        - write{i}_type_enable/disable = REG_DWORD|REG_SZ|REG_BINARY
        - write{i}_enable/disable = int/text/hex:... payload
        - write{i}_devices semantics:
            * missing => universal within bucket
            * empty   => applies to nobody
            * list    => applies only to those GUIDs
        - decider_index and quorum_threshold control verification/readback behavior
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
                        # NEW: optional per-toggle devices list
                        # Semantics:
                        #   - missing => universal within this FX bucket
                        #   - empty   => applies to nobody
                        #   - list    => applies only to those GUIDs
                        raw_devices = cfg.get(sec, f"write{i}_devices", fallback=None)
                        if raw_devices is None:
                            devs = None            # universal (applies to all)
                        else:
                            raw_devices = raw_devices.strip()
                            if not raw_devices:
                                devs = []          # explicit: applies to nobody
                            else:
                                devs = [x.strip().lower() for x in raw_devices.split(",") if x.strip()]
                        writes.append({
                            "hive": hive,
                            "subkey": subk,
                            "name": name,
                            "type_enable": t_en,
                            "type_disable": t_di,
                            "enable": v_en,
                            "disable": v_di,
                            "devices": devs,  # None=universal, []=none, list=[guids]
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
    # Registry base used for endpoint properties:
    # HKCU/HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\{Render|Capture}\{GUID}\FxProperties
    key_path = rf"SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\{flow_name}\{guid}\FxProperties"
    return flow_name, key_path

def _guid_of(device_id):
    g = _extract_endpoint_guid_from_device_id(device_id)
    return (g or "").strip().lower()

def _vendor_entry_applies(entry, device_id, flow):
    r"""
    Return True if this MAIN entry applies to this endpoint AND the configured value
    exists under HKCU for the endpoint (FxProperties or Properties).
    - Checks devices membership and flows
    - HKCU only (per your environment)
    - Probes both FxProperties and Properties for value_name
    Why we probe for existence:
      Some drivers only create the relevant value after the user toggles the setting
      once in the Windows UI. This avoids writing to a "learned but not initialized"
      value path.
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
    r"""
    Write vendor entry DWORD to desired value across configured hives.
    Uses MAIN 'subkey' (where it came from) exactly.
    Scope rule:
      We write to the learned location (FxProperties vs Properties) recorded in the INI
      so we don't guess at runtime.
    Admin note:
      HKLM writes typically require Administrator privileges.
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
    r"""Append FX entry to INI. Raises ValueError if section exists."""
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
    r"""
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
    Multi-write notes:
      - Some drivers require multiple writes to reproduce an effect toggle.
      - write{i}_devices is optional and implements per-write scoping:
          missing => universal within bucket
          empty   => applies to nobody
          list    => applies only to those GUIDs
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
        if "devices" in w and isinstance(w["devices"], list):
            lines.append(f"write{i}_devices = {','.join(sorted(set(x.lower() for x in w['devices'])))}")
    if notes:
        lines.append(f"notes = {notes}")
    lines.append("devices = ")
    with open(ini_path, "a", encoding="utf-8", errors="replace") as f:
        f.write("\n".join(lines) + "\n")

def _read_vendor_entry_state(entry, device_id, flow):
    r"""
    Return True if current state equals 'enable' value, False if equals 'disable', None otherwise.
    Behavior:
      - For FX entries with multi_write=True: uses _read_decider_state (unchanged).
      - For MAIN entries and legacy single-DWORD FX entries:
          Read exactly the learned scope:
            HKCU\...\{FxProperties|Properties}\value_name for THIS endpoint,
          fallback to HKLM only if HKCU read is not present.
    """
    # Multi-write FX: state is determined via decider/quorum logic because multiple
    # values can represent one effect state.
    if entry.get("type") == "fx" and entry.get("multi_write"):
        return _read_decider_state(entry, device_id, flow)
    # MAIN (enhancements) or legacy single-DWORD FX
    val_name = (entry.get("value_name") or "").strip().lower()
    if not val_name:
        return None
    # Where it came from (learned)
    subkey = (entry.get("subkey") or "FxProperties").strip()
    base = _endpoint_base_path(device_id, flow, subkey)
    if not base:
        return None
    # Prefer HKCU, then HKLM if HKCU missing
    hive_order = []
    configured = entry.get("hives") or []
    if configured:
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
    # Accept either key naming for enable/disable
    try:
        en = int(entry.get("enable"))
        di = int(entry.get("disable"))
    except Exception:
        try:
            en = int(entry.get("dword_enable"))
            di = int(entry.get("dword_disable"))
        except Exception:
            return None
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
        if typ != winreg.REG_DWORD:
            continue
        try:
            v = int(val)
        except Exception:
            continue
        if v == en:
            return True
        if v == di:
            return False
    return None

def _verify_vendor_entry(entry, device_id, flow, expected_enabled, timeout=2.5, interval=0.2, consecutive=2):
    """
    Poll the same vendor DWORD until it reflects expected_enabled for 'consecutive' reads or timeout.
    Why consecutive reads:
      Some drivers update multiple keys asynchronously; requiring the same answer
      multiple times avoids transient states being reported as final.
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
            # Fail-fast: still require GUID membership for selecting candidates quickly,
            # but truth is signature. This prevents false "supported" when the value
            # doesn't exist for this endpoint.
            if _vendor_entry_applies(entry, device_id, flow) and _main_entry_signature_applies(entry, device_id, flow):
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
        if _vendor_entry_applies(entry, device_id, flow_name) and _main_entry_signature_applies(entry, device_id, flow_name):
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

def _write_applies_to_guid(w, guid_lc: str) -> bool:
    """
    True if this toggle applies to guid_lc.
    devices is interpreted as:
      - None  => universal (applies to all)
      - []    => applies to nobody
      - list  => applies only to listed GUIDs
    """
    devs = w.get("devices", None)
    if devs is None:
        return True
    if isinstance(devs, list) and len(devs) == 0:
        return False
    try:
        return guid_lc in {x.strip().lower() for x in devs}
    except Exception:
        return False

# --- FAST, SINGLE-PROBE READ HELPERS (no fallbacks, no COM) ---
# These are used by the GUI to poll state quickly without COM calls.
# When HKCU and HKLM disagree, we use key last-write time as a heuristic to prefer
# the hive most recently updated by the driver.
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

def _fast_key_lastwrite(hive_name: str, base_path: str):
    """
    Return the registry key last-write time (as an integer) for the given hive/base.
    None if the key cannot be opened. The integer is a FILETIME-scale value; we
    only compare magnitudes between hives, no conversion needed.
    """
    hive = winreg.HKEY_LOCAL_MACHINE if (hive_name or "").upper() == "HKLM" else winreg.HKEY_CURRENT_USER
    try:
        with winreg.OpenKey(hive, base_path, 0, winreg.KEY_READ) as key:
            _, _, last = winreg.QueryInfoKey(key)
            try:
                return int(last)
            except Exception:
                return None
    except OSError:
        return None

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
    - MAIN: unchanged from your version (HKCU/HKLM with newer-key tie-break)
    - FX multi-write: pick the best applicable toggle (filtered by write{i}_devices),
      probe recorded hive then alternate; newer-key tie-break if both disagree.
    """
    try:
        # FX multi-write (decider-only quick read with heuristic, filtered by GUID)
        if entry.get("type") == "fx" and entry.get("multi_write"):
            all_writes = entry.get("writes") or []
            if not all_writes:
                return None
            guid_lc = _guid_of(device_id)
            writes = [w for w in all_writes if _write_applies_to_guid(w, guid_lc)]
            if not writes:
                return None
            # Score: prefer FxProperties, then REG_DWORD (0/1), else others
            # This keeps GUI state reads fast while still picking the most stable indicator.
            def _score(w):
                s = 0
                if str((w.get("subkey") or "")).strip().startswith("FxProperties"):
                    s += 10
                t_en = (w.get("type_enable") or "").upper()
                t_di = (w.get("type_disable") or "").upper()
                if t_en == "REG_DWORD" and t_di == "REG_DWORD":
                    s += 5
                    try:
                        if {int(w.get("enable")), int(w.get("disable"))} == {0, 1}:
                            s += 2
                    except Exception:
                        pass
                return s
            w = sorted(writes, key=_score, reverse=True)[0]
            rec_hive = (w.get("hive") or "HKCU").upper()
            alt_hive = "HKCU" if rec_hive == "HKLM" else "HKLM"
            subkey = (w.get("subkey") or "FxProperties").strip()
            val_name = (w.get("name") or "").strip().lower()
            base = _endpoint_base_path(device_id, flow, subkey)
            if not base:
                return None
            states, times = {}, {}
            for hn in (rec_hive, alt_hive):
                val, typ = _fast_read_one(hn, base, val_name)
                if val is None:
                    states[hn] = None
                    times[hn] = _fast_key_lastwrite(hn, base)
                    continue
                try:
                    _ = _reg_name_to_type(w.get("type_enable"))
                    _ = _reg_name_to_type(w.get("type_disable"))
                except Exception:
                    states[hn] = None
                    times[hn] = _fast_key_lastwrite(hn, base)
                    continue
                if _value_equals(w.get("enable"), w.get("type_enable"), val, typ):
                    states[hn] = True
                elif _value_equals(w.get("disable"), w.get("type_disable"), val, typ):
                    states[hn] = False
                else:
                    states[hn] = None
                times[hn] = _fast_key_lastwrite(hn, base)
            s_rec, s_alt = states.get(rec_hive), states.get(alt_hive)
            if s_rec is not None and s_alt is None:
                return s_rec
            if s_alt is not None and s_rec is None:
                return s_alt
            if s_rec is not None and s_alt is not None:
                if s_rec == s_alt:
                    return s_rec
                # Tie-break when both hives are readable but disagree:
                # prefer whichever key was written more recently.
                t_rec = times.get(rec_hive); t_alt = times.get(alt_hive)
                try:
                    if isinstance(t_rec, int) and isinstance(t_alt, int):
                        if t_alt > t_rec: return s_alt
                        if t_rec > t_alt: return s_rec
                        return s_rec
                except Exception:
                    pass
                return s_rec
            return None
        # MAIN / legacy single-DWORD (heuristic: compare hives, use newer key)
        val_name = (entry.get("value_name") or "").strip().lower()
        subkey = (entry.get("subkey") or "FxProperties").strip()
        if not val_name:
            return None
        # Enable/disable values (support legacy keys)
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
        # Allowed hives filter (default to both)
        configured = entry.get("hives") or []
        allowed = {str(h).strip().upper() for h in configured if isinstance(h, str)}
        if not allowed:
            allowed = {"HKCU", "HKLM"}
        # Read both (subject to allowed)
        state = {}      # hive -> True/False/None
        lastw = {}      # hive -> int or None
        for hname in ("HKCU", "HKLM"):
            if hname not in allowed:
                state[hname] = None
                lastw[hname] = None
                continue
            val, typ = _fast_read_one(hname, base, val_name)
            if val is None or typ != winreg.REG_DWORD:
                state[hname] = None
            else:
                try:
                    v = int(val)
                    if v == en_val:
                        state[hname] = True
                    elif v == di_val:
                        state[hname] = False
                    else:
                        state[hname] = None
                except Exception:
                    state[hname] = None
            lastw[hname] = _fast_key_lastwrite(hname, base)
        cu = state.get("HKCU")
        lm = state.get("HKLM")
        if cu is not None and lm is None:
            return cu
        if lm is not None and cu is None:
            return lm
        if cu is not None and lm is not None and cu == lm:
            return cu
        if cu is not None and lm is not None and cu != lm:
            # Tie-break when both hives are readable but disagree:
            # choose the most recently written key.
            tcu = lastw.get("HKCU")
            tlm = lastw.get("HKLM")
            try:
                if isinstance(tcu, int) and isinstance(tlm, int):
                    if tlm > tcu:
                        return lm
                    elif tcu > tlm:
                        return cu
                    else:
                        return cu  # tie -> HKCU
            except Exception:
                pass
            return cu
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

def _append_guid_to_write_devices(ini_path, section_name, write_index, guid_lc):
    """
    Ensure write{write_index}_devices contains guid_lc.
    If missing, create it. If empty, add guid_lc. Keeps list unique and sorted.
    """
    try:
        with open(ini_path, "r", encoding="utf-8", errors="replace") as f:
            lines = f.read().splitlines(keepends=True)
    except FileNotFoundError:
        return
    sec_hdr = f"[{section_name}]"
    sec_start = None
    sec_end = len(lines)
    for i, line in enumerate(lines):
        s = line.strip()
        if s.startswith("[") and s.endswith("]"):
            if sec_start is None:
                if s.lower() == sec_hdr.lower():
                    sec_start = i
            else:
                sec_end = i
                break
    if sec_start is None:
        return
    key_pat = re.compile(rf"^\s*write{write_index}_devices\s*=\s*(.*)$", re.IGNORECASE)
    devices_idx = None
    existing = None  # None means no line present
    for i in range(sec_start + 1, sec_end):
        m = key_pat.match(lines[i])
        if m:
            devices_idx = i
            txt = m.group(1).strip()
            if not txt:
                existing = []  # explicit none
            else:
                existing = [x.strip().lower() for x in txt.split(",") if x.strip()]
            break
    if existing is None:
        new_val = guid_lc.lower()
        new_line = f"write{write_index}_devices = {new_val}\n"
        insert_at = sec_end
        if insert_at > 0 and not lines[insert_at - 1].endswith(("\n", "\r")):
            lines.insert(insert_at, "\n"); insert_at += 1
        lines.insert(insert_at, new_line)
    else:
        if guid_lc.lower() not in existing:
            existing.append(guid_lc.lower())
            new_line = f"write{write_index}_devices = {','.join(sorted(set(existing)))}\n"
            if devices_idx is not None:
                lines[devices_idx] = new_line
    with open(ini_path, "w", encoding="utf-8", errors="replace") as f:
        f.writelines(lines)

def _learn_vendor_from_discovery_and_write_ini(target, ini_path=None, prefer_hkcu=True):
    """
    Manual learn using discovery flow.
    Safety/UX:
      Requires explicit confirmation ("I UNDERSTAND") unless AUDIOCTL_LEARN_CONFIRMED=1 is set
      (GUI uses the env var after showing its own warning dialog).
    Learn behavior:
      - Captures A/B snapshots (user sets enabled, then disabled)
      - Chooses a candidate DWORD flip and records the subkey where it occurred
      - Dedupes identical entries by appending this GUID to an existing section when possible
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
            # Also record device name -> guid bucket mapping (readability, dedupe by name)
            try:
                _append_guid_to_name_bucket(ini_path, e.get("name"), name, guid_lc)
            except Exception:
                pass
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
        # Also record device name -> guid bucket mapping (readability, dedupe by name)
        try:
            _append_guid_to_name_bucket(ini_path, section_name, name, guid_lc)
        except Exception:
            pass
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
        # Also record device name -> guid bucket mapping (readability, dedupe by name)
        try:
            _append_guid_to_name_bucket(ini_path, section_name, name, guid_lc)
        except Exception:
            pass
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
    # Base path pattern (for reference):
    # HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\Render\{GUID}\FxProperties
    return rf"SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\{flow_name}\{guid}\{subkey}"

def _perform_multi_writes(entry, device_id, flow, enable):
    """
    Write all applicable toggles (universal + for this GUID).
    Returns True if ALL applicable writes succeeded; False otherwise.
    Notes:
      - Multi-write FX can include HKLM writes; those may require Administrator privileges.
      - We treat any failed applicable write as a failure because partial application
        often produces inconsistent driver state.
    """
    guid_lc = _guid_of(device_id)
    ok_all = True
    for w in entry.get("writes") or []:
        if not _write_applies_to_guid(w, guid_lc):
            continue
        hive_name = (w.get("hive") or "").upper()
        hive = winreg.HKEY_LOCAL_MACHINE if hive_name == "HKLM" else winreg.HKEY_CURRENT_USER
        subk = (w.get("subkey") or "").strip()
        name = (w.get("name") or "").strip().lower()
        base = _endpoint_base_path(device_id, flow, subk)
        if not base:
            ok_all = False
            continue
        tname = w.get("type_enable") if enable else w.get("type_disable")
        try:
            typ = _reg_name_to_type(tname)
        except Exception:
            ok_all = False
            continue
        val_text = w.get("enable") if enable else w.get("disable")
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
    # Multi-write readback:
    # - First attempt quorum decision (fraction of applicable writes that agree).
    # - If quorum can't be reached, fall back to reading a "best" signal write (FxProperties, DWORD 0/1 preferred).
    if not entry.get("multi_write"):
        return None
    all_writes = entry.get("writes") or []
    if not all_writes:
        return None
    guid_lc = _guid_of(device_id)
    writes = [w for w in all_writes if _write_applies_to_guid(w, guid_lc)]
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
    votes_true = votes_false = votes_total = 0
    for w in writes:
        rec_hive = (w.get("hive") or "").upper()
        alt_hive = "HKCU" if rec_hive == "HKLM" else "HKLM"
        s = _try_read_one(w, rec_hive)
        if s is None:
            s = _try_read_one(w, alt_hive)
        if s is True:
            votes_true += 1; votes_total += 1
        elif s is False:
            votes_false += 1; votes_total += 1
    if votes_total > 0:
        if votes_true / votes_total >= quorum_threshold and votes_false / votes_total < quorum_threshold:
            return True
        if votes_false / votes_total >= quorum_threshold and votes_true / votes_total < quorum_threshold:
            return False
    def _score(w):
        s = 0
        if str((w.get("subkey") or "").strip()).startswith("FxProperties"):
            s += 10
        t_en = (w.get("type_enable") or "").upper()
        t_di = (w.get("type_disable") or "").upper()
        if t_en == "REG_DWORD" and t_di == "REG_DWORD":
            s += 5
            try:
                if {int(w.get("enable")), int(w.get("disable"))} == {0, 1}:
                    s += 2
            except Exception:
                pass
        return s
    for w in sorted(writes, key=_score, reverse=True):
        rec_hive = (w.get("hive") or "").upper()
        alt_hive = "HKCU" if rec_hive == "HKLM" else "HKLM"
        s = _try_read_one(w, rec_hive)
        if s is not None:
            return s
        s = _try_read_one(w, alt_hive)
        if s is not None:
            return s
    return None

def _dump_mmdevices_all_values_for_fx_learn(device_id):
    # Deprecated in favor of _dump_mmdevices_all_values from devices.py
    return _dump_mmdevices_all_values(device_id)

def _find_fx_bucket_section_name(ini_path, fx_name):
    cfg = configparser.ConfigParser()
    try:
        if os.path.exists(ini_path):
            cfg.read(ini_path, encoding="utf-8")
    except Exception:
        return None
    target = (fx_name or "").strip().lower()
    for sec in cfg.sections():
        try:
            if cfg.get(sec, "type", fallback="").strip().lower() != "fx":
                continue
            if cfg.get(sec, "fx_name", fallback="").strip().lower() == target:
                return sec
        except Exception:
            continue
    return None

def _canonical_fx_bucket_name(fx_name):
    import hashlib
    key = (fx_name or "").strip().lower()
    h = hashlib.sha1(key.encode("utf-8", "replace")).hexdigest()[:16]
    return f"fx_{h}"

def _append_new_write_to_section(ini_path, section_name, write_dict, guid_lc):
    """
    Append a new write{i}_* block (with write{i}_devices = {guid}) and bump write_count.
    write_dict keys: hive, subkey, name, type_enable, type_disable, enable, disable
    """
    try:
        with open(ini_path, "r", encoding="utf-8", errors="replace") as f:
            lines = f.read().splitlines(keepends=True)
    except FileNotFoundError:
        return
    sec_hdr = f"[{section_name}]"
    sec_start = None
    sec_end = len(lines)
    for i, line in enumerate(lines):
        s = line.strip()
        if s.startswith("[") and s.endswith("]"):
            if sec_start is None:
                if s.lower() == sec_hdr.lower():
                    sec_start = i
            else:
                sec_end = i
                break
    if sec_start is None:
        return
    wc_idx = None
    write_count = 0
    wc_pat = re.compile(r"^\s*write_count\s*=\s*(\d+)\s*$", re.IGNORECASE)
    for i in range(sec_start + 1, sec_end):
        m = wc_pat.match(lines[i])
        if m:
            wc_idx = i
            try:
                write_count = int(m.group(1))
            except Exception:
                write_count = 0
            break
    new_idx = write_count + 1 if write_count > 0 else 1
    insert_at = sec_end
    if insert_at > 0 and not lines[insert_at - 1].endswith(("\n", "\r")):
        lines.insert(insert_at, "\n"); insert_at += 1
    w = write_dict
    block = [
        f"write{new_idx}_hive = {w.get('hive')}\n",
        f"write{new_idx}_subkey = {w.get('subkey')}\n",
        f"write{new_idx}_name = {w.get('name')}\n",
        f"write{new_idx}_type_enable = {w.get('type_enable')}\n",
        f"write{new_idx}_type_disable = {w.get('type_disable')}\n",
        f"write{new_idx}_enable = {w.get('enable')}\n",
        f"write{new_idx}_disable = {w.get('disable')}\n",
        f"write{new_idx}_devices = {guid_lc}\n",
    ]
    for line in block:
        lines.insert(insert_at, line); insert_at += 1
    if wc_idx is not None:
        lines[wc_idx] = f"write_count = {new_idx}\n"
    else:
        lines.insert(sec_start + 1, f"write_count = {new_idx}\n")
    with open(ini_path, "w", encoding="utf-8", errors="replace") as f:
        f.writelines(lines)

def _delete_fx_for_guid(fx_name, device_id, ini_path=None):
    """
    Remove associations for 'fx_name' for the specific device GUID from vendor_toggles.ini:
      - Remove GUID from the section devices (union)
      - For each write{i}, remove GUID from write{i}_devices (creating write{i}_devices to exclude this GUID if the write was universal)
      - Keep write blocks even if they end up with empty devices (applies to nobody); runtime ignores them
      - If no devices remain in the bucket (section devices empty), we leave the section; the loader will ignore it
    Returns (True, info_dict) or (False, reason_str)
    Delete semantics (important):
      We remove GUID associations rather than deleting the whole FX bucket. Buckets
      can be shared between multiple devices and keep useful learned history.
      If a write block was universal (no write{i}_devices line), delete converts it
      into an explicit scoped list of remaining devices so the removed GUID is excluded.
      Empty write{i}_devices lines are preserved intentionally (they mean "applies to nobody").
    """
    ini_path = ini_path or _vendor_ini_default_path()
    guid_lc = _guid_of(device_id)
    if not guid_lc:
        return False, "bad-device-id"
    section = _find_fx_bucket_section_name(ini_path, fx_name)
    if not section:
        return False, "fx-bucket-not-found"
    try:
        with open(ini_path, "r", encoding="utf-8", errors="replace") as f:
            lines = f.read().splitlines(keepends=True)
    except FileNotFoundError:
        return False, "ini-not-found"
    sec_hdr = f"[{section}]"
    sec_start = None
    sec_end = len(lines)
    for i, line in enumerate(lines):
        s = line.strip()
        if s.startswith("[") and s.endswith("]"):
            if sec_start is None:
                if s.lower() == sec_hdr.lower():
                    sec_start = i
            else:
                sec_end = i
                break
    if sec_start is None:
        return False, "bucket-section-missing"
    # Section devices (union)
    devices_idx = None
    cur_devices = []
    dev_pat = re.compile(r"^\s*devices\s*=\s*(.*)$", re.IGNORECASE)
    for i in range(sec_start + 1, sec_end):
        m = dev_pat.match(lines[i])
        if m:
            devices_idx = i
            txt = (m.group(1) or "").strip()
            cur_devices = [x.strip().lower() for x in txt.split(",") if x.strip()]
            break
    # write_count
    wc_idx = None
    write_count = 0
    wc_pat = re.compile(r"^\s*write_count\s*=\s*(\d+)\s*$", re.IGNORECASE)
    for i in range(sec_start + 1, sec_end):
        m = wc_pat.match(lines[i])
        if m:
            wc_idx = i
            try:
                write_count = int(m.group(1))
            except Exception:
                write_count = 0
            break
    # Helpers for per-write devices
    def _get_write_devices(i_idx):
        pat = re.compile(rf"^\s*write{i_idx}_devices\s*=\s*(.*)$", re.IGNORECASE)
        for j in range(sec_start + 1, sec_end):
            m = pat.match(lines[j] or "")
            if m:
                txt = (m.group(1) or "").strip()
                devs = [x.strip().lower() for x in txt.split(",") if x.strip()] if txt else []
                return j, devs
        return None, None
    def _set_write_devices(i_idx, dev_list):
        pat = re.compile(rf"^\s*write{i_idx}_devices\s*=", re.IGNORECASE)
        if dev_list is None:
            line_txt = f"write{i_idx}_devices = \n"  # explicit none
        else:
            line_txt = f"write{i_idx}_devices = {','.join(sorted(set(d.lower() for d in dev_list)))}\n" if dev_list else f"write{i_idx}_devices = \n"
        # Try to replace if exists
        for j in range(sec_start + 1, sec_end):
            if pat.match(lines[j] or ""):
                lines[j] = line_txt
                return
        # Else insert after write{i}_disable or at sec_end
        after_pat = re.compile(rf"^\s*write{i_idx}_disable\s*=", re.IGNORECASE)
        insert_at = sec_end
        for j in range(sec_start + 1, sec_end):
            if after_pat.match(lines[j] or ""):
                insert_at = j + 1
                break
        if insert_at > 0 and not lines[insert_at - 1].endswith(("\n", "\r")):
            lines.insert(insert_at, "\n"); insert_at += 1
        lines.insert(insert_at, line_txt)
    writes_changed = 0
    # Remaining devices in bucket (minus target)
    remaining_bucket_devs = [d for d in cur_devices if d != guid_lc]
    # Scan writes (up to declared count or a generous cap)
    max_scan = write_count if write_count > 0 else 256
    for i_idx in range(1, max_scan + 1):
        # Does write{i} exist?
        hv_pat = re.compile(rf"^\s*write{i_idx}_hive\s*=", re.IGNORECASE)
        exists = any(hv_pat.match(lines[j] or "") for j in range(sec_start + 1, sec_end))
        if not exists:
            if write_count > 0 and i_idx > write_count:
                break
            continue
        w_idx, w_devs = _get_write_devices(i_idx)
        if w_idx is not None:
            if guid_lc in (w_devs or []):
                new_list = [x for x in (w_devs or []) if x != guid_lc]
                _set_write_devices(i_idx, new_list)
                writes_changed += 1
        else:
            # Universal write: scope to remaining devices (exclude target) or to none
            if remaining_bucket_devs:
                _set_write_devices(i_idx, remaining_bucket_devs)
                writes_changed += 1
            else:
                _set_write_devices(i_idx, [])  # applies to nobody
                writes_changed += 1
    # Update section devices (union)
    new_devices = [d for d in cur_devices if d != guid_lc]
    new_line = f"devices = {','.join(sorted(set(new_devices)))}\n" if new_devices else "devices = \n"
    if devices_idx is not None:
        lines[devices_idx] = new_line
    else:
        insert_at = sec_end
        if insert_at > 0 and not lines[insert_at - 1].endswith(("\n", "\r")):
            lines.insert(insert_at, "\n"); insert_at += 1
        lines.insert(insert_at, new_line)
    try:
        with open(ini_path, "w", encoding="utf-8", errors="replace") as f:
            f.writelines(lines)
    except Exception as e:
        return False, f"write-ini-failed: {e}"
    return True, {
        "iniPath": ini_path,
        "section": section,
        "removedGuid": guid_lc,
        "writesAffected": writes_changed,
        "remainingDevices": new_devices,
    }

def _learn_fx_and_write_ini(target, fx_name, snapA, snapB, ini_path=None, prefer_hkcu=True, snapA2=None, snapB2=None):
    """
    Learn an FX toggle from captured snapshots.
    Two-pass behavior:
      If snapA2/snapB2 is provided, that second pair is treated as authoritative.
      The first A/B is used to prime/initialize driver state (many drivers create keys
      or stabilize values only after first toggle).
    Multi-write merge behavior:
      - If an identical identity+payload already exists in the bucket, append this GUID
        to that write's write{i}_devices.
      - If a new payload is discovered, append a new write block scoped to this GUID.
      - Conflict cleanup ensures a GUID is not attached to two competing writes with the
        same identity but different payload.
    Change here: if a second A/B pass (snapA2/snapB2) is provided, use that
    pair as the authoritative input for building the write set (the first
    A/B is only for priming/initializing driver state). Otherwise, behave
    exactly as before.
    """
    import re
    ini_path = ini_path or _vendor_ini_default_path()
    guid_lc = _guid_of(target["id"])
    # Choose which pair to use for building the write set
    useA = snapA2 if isinstance(snapA2, dict) else snapA
    useB = snapB2 if isinstance(snapB2, dict) else snapB
    # Build stability-filtered maps EXACTLY like before, just with the chosen pair.
    # A side: single snapshot stability map (previous behavior)
    try:
        stableA = _stable_registry_map([useA.get("registry") or []])
    except Exception as e:
        return False, f"Failed to process snapshot A: {e}"
    # B side: B snapshot + a couple of quick samples (previous behavior)
    try:
        samplesB = _collect_registry_samples(target["id"], repeats=3, delay=0.18)
        samplesB.insert(0, useB.get("registry") or [])
        stableB = _stable_registry_map(samplesB)
    except Exception as e:
        return False, f"Failed to process snapshot B: {e}"
    # Compute diff-based multi-write set (previous behavior)
    writes = _build_fx_multiwrite_from_stable_maps(target, stableA, stableB)
    safe_device_name = re.sub(r'[^A-Za-z0-9_\- ]+', '_', target["name"])
    notes = f"Learned FX '{fx_name}' for '{target['name']}' ({target['flow']}); second A/B pass; stability-filtered"
    if writes:
        # Find or create the bucket for this fx_name
        bucket = _find_fx_bucket_section_name(ini_path, fx_name)
        if bucket is None:
            # First device for this fx_name: create bucket named by fx_name (stable)
            for w in writes:
                w.setdefault("devices", None)  # allow universal later if you want
            section_name = _canonical_fx_bucket_name(fx_name)
            # seed writes with write{i}_devices = this guid
            seed = []
            for w in writes:
                seed.append({
                    "hive": w.get("hive"),
                    "subkey": w.get("subkey"),
                    "name": w.get("name"),
                    "type_enable": w.get("type_enable"),
                    "type_disable": w.get("type_disable"),
                    "enable": w.get("enable"),
                    "disable": w.get("disable"),
                    "devices": [guid_lc],
                })
            _append_fx_ini_entry_multi(
                ini_path, section_name, fx_name, target["name"],
                seed, notes=notes
            )
            _append_guid_to_section(ini_path, section_name, guid_lc)
            # Also record device name -> guid bucket mapping (readability, dedupe by name)
            try:
                _append_guid_to_name_bucket(ini_path, section_name, target["name"], guid_lc)
            except Exception:
                pass
            return True, {
                "iniPath": ini_path,
                "section": section_name,
                "fx_name": fx_name,
                "multi_write": True,
                "write_count": len(seed),
            }
        # Merge into existing bucket
        db = _load_vendor_db_split(ini_path)
        current = None
        for e in (db.get("fx") or []):
            if e.get("name") == bucket:
                current = e
                break
        if current is None:
            return False, f"Bucket '{bucket}' not found."
        def _same_identity_payload(a, b):
            return (
                (a.get("hive","").upper() == b.get("hive","").upper()) and
                (str(a.get("subkey","")).strip().lower() == str(b.get("subkey","")).strip().lower()) and
                (str(a.get("name","")).strip().lower() == str(b.get("name","")).strip().lower()) and
                (str(a.get("type_enable","")).upper() == str(b.get("type_enable","")).upper()) and
                (str(a.get("type_disable","")).upper() == str(b.get("type_disable","")).upper()) and
                (str(a.get("enable","")).strip() == str(b.get("enable","")).strip()) and
                (str(a.get("disable","")).strip() == str(b.get("disable","")).strip())
            )
        # Add/append each learned toggle to the bucket and clean conflicts
        for lw in writes:
            idx_match = None
            for idx, cw in enumerate(current.get("writes") or [], start=1):
                if _same_identity_payload(lw, cw):
                    idx_match = idx
                    break
            if idx_match is not None:
                _append_guid_to_write_devices(ini_path, bucket, idx_match, guid_lc)
                _cleanup_conflicting_toggles(ini_path, bucket, guid_lc, idx_match, lw)
            else:
                new_w = {
                    "hive": lw.get("hive"),
                    "subkey": lw.get("subkey"),
                    "name": lw.get("name"),
                    "type_enable": lw.get("type_enable"),
                    "type_disable": lw.get("type_disable"),
                    "enable": lw.get("enable"),
                    "disable": lw.get("disable"),
                }
                _append_new_write_to_section(ini_path, bucket, new_w, guid_lc)
                new_idx = _find_write_index_by_payload(ini_path, bucket, new_w)
                if new_idx is not None:
                    _cleanup_conflicting_toggles(ini_path, bucket, guid_lc, new_idx, new_w)
        # Ensure bucket devices includes this GUID (for discovery)
        _append_guid_to_section(ini_path, bucket, guid_lc)
        # Also record device name -> guid bucket mapping (readability, dedupe by name)
        try:
            _append_guid_to_name_bucket(ini_path, bucket, target["name"], guid_lc)
        except Exception:
            pass
        return True, {
            "iniPath": ini_path,
            "section": bucket,
            "fx_name": fx_name,
            "multi_write": True,
            "write_count": None
        }
    # Fallback to legacy DWORD flip (previous behavior)
    try:
        diffs = _diff_mmdevices_lists((useA.get("registry") or []), (useB.get("registry") or []))
    except Exception as e:
        return False, f"Diff failed: {e}"
    snippet, picked = _build_vendor_ini_snippet(target, useA, useB, diffs)
    if not picked:
        return False, "No suitable registry differences found to learn."
    value_name = picked["name"]
    dword_enable = int(picked["before"])
    dword_disable = int(picked["after"])
    notes2 = notes + " (single DWORD)"
    hives = "HKCU,HKLM" if prefer_hkcu else "HKLM,HKCU"
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
            # Also record device name -> guid bucket mapping (readability, dedupe by name)
            try:
                _append_guid_to_name_bucket(ini_path, e.get("name"), target["name"], guid_lc)
            except Exception:
                pass
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
        # Also record device name -> guid bucket mapping (readability, dedupe by name)
        try:
            _append_guid_to_name_bucket(ini_path, section_name, target["name"], guid_lc)
        except Exception:
            pass
    except ValueError as e:
        msg = str(e or "").lower()
        if "already exists" in msg:
            try:
                _append_guid_to_section(ini_path, section_name, guid_lc)
                # Also record device name -> guid bucket mapping (readability, dedupe by name)
                try:
                    _append_guid_to_name_bucket(ini_path, section_name, target["name"], guid_lc)
                except Exception:
                    pass
                return True, {
                    "iniPath": ini_path,
                    "section": section_name,
                    "fx_name": fx_name,
                    "value_name": value_name,
                    "dword_enable": dword_enable,
                    "dword_disable": dword_disable
                }
            except Exception as e2:
                return False, f"Failed to append device to existing section '{section_name}': {e2}"
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

def _list_fx_for_device(device_id, flow, ini_path=None, device_name=None):
    """
    List all available FX for a device.
    Matching (read-only; does NOT modify INI):
      1) Direct GUID membership in section 'devices' (fast)
      2) If device_name provided: pattern match + registry signature match ("spoof")
    Returns [{'fx_name','entry'}]
    """
    db = _load_vendor_db_split(ini_path)
    guid = _extract_endpoint_guid_from_device_id(device_id)
    if not guid:
        return []
    guid_lc = guid.strip().lower()
    out = []
    for entry in db.get("fx") or []:
        # Fail-fast candidate narrowing: only consider entries that match GUID or pattern.
        if not _fx_candidate_by_guid_or_pattern(entry, guid_lc, device_name):
            continue
        # Truth check: signature must fit this endpoint right now.
        try:
            if _fx_entry_signature_applies(entry, device_id, flow):
                e = dict(entry)
                e["source"] = "ini"
                e["_matchedBy"] = "signature"
                out.append({"fx_name": entry.get("fx_name"), "entry": e})
        except Exception:
            continue
    return out

def _find_fx_for_device(device_id, flow, fx_name, ini_path=None, device_name=None):
    """
    Find FX entries matching device and effect name.
    Uses same match rules as _list_fx_for_device (GUID or spoof if device_name provided).
    """
    fx_lc = str(fx_name or "").strip().lower()
    if not fx_lc:
        return []
    lst = _list_fx_for_device(device_id, flow, ini_path=ini_path, device_name=device_name)
    matches = []
    for item in lst:
        if (item.get("fx_name") or "").strip().lower() == fx_lc:
            e = dict(item.get("entry") or {})
            e["source"] = "ini"
            matches.append(e)
    return matches

def _apply_enhancements(device_id, flow, enable, prefer_hklm=False, allow_universal_scan=False, vendor_ini_path=None):
    """
    Vendor-only policy:
      1) Try vendor toggles: INI vendors only (per-device).
      2) If no vendor match, return failure (no Windows fallback).
    Runtime contract:
      Returns (False, "no-vendor-method", None) when no applicable vendor entry is known.
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
            # Fail fast: only allow apply when signature fits THIS endpoint now.
            if not _main_entry_signature_applies(entry, device_id, flow):
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
    This is a membership check used by CLI/GUI to decide if a device has a learned vendor method.
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

def _apply_fx(device_id, flow, fx_name, enable, ini_path=None, device_name=None):
    """
    Toggle a learned FX effect (INI-only).
    - If the FX entry is multi_write, apply all configured writes for the target state.
    - Otherwise, legacy single-DWORD behavior.
    Returns (success, verified_by, final_state).
    Verification:
      - multi_write: read back state using decider/quorum logic after writing
      - legacy: poll the single learned DWORD until stable
    """
    entries = _find_fx_for_device(device_id, flow, fx_name, ini_path, device_name=device_name)
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
