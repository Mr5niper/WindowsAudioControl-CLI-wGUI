# audioctl/vendor_db.py
"""
vendor_db.py

This module implements audioctl's *vendor-first* (and in normal runtime, effectively
vendor-only) model for toggling:

- Main "Audio Enhancements" switch (SysFX enable/disable)
- Per-effect "FX" toggles (e.g., BassBoost/Loudness) learned from vendor registry behavior

Key idea:
Windows exposes a generic SysFX knob (Disable_SysFx), but many OEM/driver stacks
(Realtek/Waves, etc.) actually honor device-specific registry values under MMDevices.
audioctl learns those vendor values once and then drives them directly in future runs.

This file provides:
- Parsing vendor_toggles.ini into an internal "main" vs "fx" database
- A lightweight path+mtime cache so GUI polling and frequent CLI calls don't reparse INI
- Applying learned toggles (main and FX), including multi-write FX operations
- Learn flows that derive registry deltas (A/B snapshots) and write/merge INI entries
- Fast "single probe" state readers used by the GUI to render menu labels quickly
"""

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
# "Multi-write" exists because some drivers don't expose a single clean DWORD flip:
# enabling/disabling a single effect can toggle multiple registry values across:
#   - FxProperties and/or Properties
#   - HKCU and/or HKLM
# and sometimes those values are REG_BINARY blobs (PROPVARIANT-encoded).
# The learn flow captures the exact raw payload so we can reproduce it later.

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
      - 'hex:aa,bb,cc' (preferred in INI for readability)
      - 'aabbcc' (raw hex without prefix)

    Returns bytes suitable for winreg.SetValueEx(..., REG_BINARY, bytes_value).
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
    Build a multi-write plan by diffing two snapshots:
      - A is "enabled"
      - B is "disabled"

    Why this exists:
    Many effect toggles are *not* a single DWORD. A driver can flip multiple
    values (DWORD/SZ/BINARY) spread across FxProperties and Properties, and may
    even mirror across HKCU/HKLM. Multi-write lets us reproduce the complete set
    of changes rather than guessing.

    Output: list of writes (each write captures enable/disable payload + types):
      {
        'hive': 'HKCU'|'HKLM',
        'subkey': 'FxProperties'|'Properties'|...,
        'name': '{fmtid},pid' (registry value name under MMDevices Properties),
        'type_enable'/'type_disable': 'REG_DWORD'|'REG_SZ'|'REG_BINARY',
        'enable'/'disable': int | str | 'hex:..'
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

        # Only consider our two canonical MMDevices property containers.
        # (MMDevices can contain deeper subkeys, but our learned write model focuses on
        # the standard 'FxProperties' and 'Properties' roots.)
        sub = str(a.get("subkey") or "")
        if not (sub.startswith("FxProperties") or sub.startswith("Properties")):
            continue

        # Compare exact raw payloads, not human preview text.
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
                # Persist binary as 'hex:aa,bb,...' so the INI is diffable by humans.
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
# We cache INI loads by (absolute_path, mtime) because:
# - the GUI calls fast state readers frequently (context menu labels, polling)
# - reparsing the INI repeatedly is wasted work
# - missing INI files are common on first run; caching avoids repeated stat/read failures
_VENDOR_DB_CACHE = {
    "path": None,
    "mtime": None,
    "data": {"main": [], "fx": []},
}

def _vendor_ini_default_path():
    """
    Determine the default path for vendor_toggles.ini.

    Preference order:
    1) Next to the EXE (or package) when writable:
       - Best UX for portable, "single-folder" deployments (PyInstaller).
       - But common installs under Program Files are *not writable* without Admin.
    2) User-writable fallback:
       - %LOCALAPPDATA%\\audioctl\\vendor_toggles.ini

    This mirrors how Windows apps typically separate "program files" (read-only)
    from "user data" (writable).
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

    # Fallback to a per-user writable location.
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
    Parse vendor_toggles.ini into a split database:
      {
        "main": [ ... main toggle entries ... ],
        "fx":   [ ... per-effect entries ... ],
      }

    The parser is intentionally permissive: bad sections are skipped so one broken
    entry doesn't break the entire tool.

    --------------------------
    INI SCHEMA (concise)
    --------------------------

    MAIN entries (type defaults to "main"):
      [section_name]
      value_name    = {fmtid-guid},pid     ; the MMDevices property name (lowercased in code)
      dword_enable  = 0|1                  ; value to write when "enabled"
      dword_disable = 0|1                  ; value to write when "disabled"
      hives         = HKCU,HKLM            ; write order / allowed hives
      flows         = Render,Capture       ; endpoint flow(s) this entry applies to
      subkey        = FxProperties|Properties
                                           ; IMPORTANT: learned location (where driver actually uses it)
      devices       = {endpoint-guid},...  ; REQUIRED membership list (per-endpoint GUIDs)
      notes         = free text

    FX entries (type = fx) support two shapes:

    (A) Legacy single DWORD FX:
      [section_name]
      type          = fx
      fx_name       = BassBoost
      value_name    = {fmtid-guid},pid
      dword_enable  = 0|1
      dword_disable = 0|1
      hives         = HKCU,HKLM
      flows         = Render,Capture
      devices       = {endpoint-guid},...  ; REQUIRED membership list

    (B) Multi-write FX (for complex drivers):
      [section_name]
      type            = fx
      fx_name         = BassBoost
      multi_write     = 1
      write_count     = N
      decider_index   = 1                  ; which write is the primary indicator (1-based)
      quorum_threshold= 0.60               ; fraction of toggles that must agree to decide state

      write{i}_hive         = HKCU|HKLM
      write{i}_subkey       = FxProperties|Properties|... (under the endpoint GUID)
      write{i}_name         = {fmtid-guid},pid           ; registry value name (casefolded)
      write{i}_type_enable  = REG_DWORD|REG_SZ|REG_BINARY
      write{i}_type_disable = REG_DWORD|REG_SZ|REG_BINARY
      write{i}_enable       = <value>                   ; int for DWORD, string for SZ, 'hex:..' for BINARY
      write{i}_disable      = <value>

      write{i}_devices semantics (per-toggle scoping):
        - missing    => universal (applies to all devices in this FX bucket)
        - empty line => applies to nobody (explicitly disabled for all devices)
        - list       => applies only to those endpoint GUIDs

    This per-write scoping is important because different devices may share an FX bucket
    name but require different underlying registry toggles; scoping prevents "wrong"
    writes from being applied to a device.
    """
    global _VENDOR_DB_CACHE

    # Resolve path
    path = os.path.abspath(ini_path or _vendor_ini_default_path())

    # Cache key is (path, mtime). If file is missing, mtime=None is cached to avoid
    # repeated stat/read attempts on every call.
    try:
        st = os.stat(path)
        mtime = st.st_mtime
        exists = True
    except OSError:
        exists = False
        mtime = None

    if not exists:
        if (_VENDOR_DB_CACHE.get("path") == path and
                _VENDOR_DB_CACHE.get("mtime") is None):
            return _VENDOR_DB_CACHE.get("data") or {"main": [], "fx": []}
        _VENDOR_DB_CACHE["path"] = path
        _VENDOR_DB_CACHE["mtime"] = None
        _VENDOR_DB_CACHE["data"] = {"main": [], "fx": []}
        return _VENDOR_DB_CACHE["data"]

    if (_VENDOR_DB_CACHE.get("path") == path and
            _VENDOR_DB_CACHE.get("mtime") == mtime):
        return _VENDOR_DB_CACHE.get("data") or {"main": [], "fx": []}

    cfg = configparser.ConfigParser()
    entries = {"main": [], "fx": []}
    try:
        cfg.read(path, encoding="utf-8")
    except Exception:
        _VENDOR_DB_CACHE["path"] = path
        _VENDOR_DB_CACHE["mtime"] = mtime
        _VENDOR_DB_CACHE["data"] = entries
        return entries

    for sec in cfg.sections():
        try:
            entry_type = cfg.get(sec, "type", fallback="main").strip().lower()
            notes = cfg.get(sec, "notes", fallback="")

            if entry_type == "fx":
                # FX entry: either legacy single-DWORD or multi-write.
                # FX entries always require device membership (`devices` list) so we never
                # accidentally apply an effect toggle to an unrelated endpoint.
                fx_name = cfg.get(sec, "fx_name", fallback="").strip()
                devpat = cfg.get(sec, "device_name_pattern", fallback="").strip()  # informational
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

                    # Bound quorum so "one flaky toggle" can't dominate (too low),
                    # and "must be perfect" can't prevent state decisions (too high).
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

                        # Optional per-write scoping:
                        # - missing => universal (applies to all)
                        # - empty   => applies to nobody
                        # - list    => applies only to listed endpoint GUIDs
                        raw_devices = cfg.get(sec, f"write{i}_devices", fallback=None)
                        if raw_devices is None:
                            devs = None
                        else:
                            raw_devices = raw_devices.strip()
                            if not raw_devices:
                                devs = []
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
                            "devices": devs,
                        })

                    e["multi_write"] = True
                    e["writes"] = writes
                    e["decider_index"] = max(1, decider_index)
                    e["quorum_threshold"] = quorum_threshold
                    e["flows"] = [x.strip().capitalize() for x in cfg.get(sec, "flows", fallback="Render,Capture").split(",") if x.strip()]
                    e["hives"] = [x.strip().upper() for x in cfg.get(sec, "hives", fallback="HKLM,HKCU").split(",") if x.strip()]
                else:
                    # Legacy single DWORD FX: use the same "vendor entry" machinery as main.
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

                # Keep only FX sections that declare membership.
                if e["devices"]:
                    entries["fx"].append(e)

            else:
                # MAIN entry (enhancements on/off).
                # Note: we record which subkey we learned from (FxProperties vs Properties)
                # because some drivers store their authoritative DWORD under Properties, not FxProperties.
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

    _VENDOR_DB_CACHE["path"] = path
    _VENDOR_DB_CACHE["mtime"] = mtime
    _VENDOR_DB_CACHE["data"] = entries
    return entries

def _endpoint_fx_key(device_id, flow):
    guid = _extract_endpoint_guid_from_device_id(device_id)
    if not guid:
        return None, None
    flow_name = "Render" if str(flow).lower().startswith("r") else "Capture"
    # Base pattern:
    #   HK??\SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\<Flow>\<EndpointGuid>\FxProperties
    key_path = rf"SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\{flow_name}\{guid}\FxProperties"
    return flow_name, key_path

def _guid_of(device_id):
    g = _extract_endpoint_guid_from_device_id(device_id)
    return (g or "").strip().lower()

def _vendor_entry_applies(entry, device_id, flow):
    """
    Determine whether a MAIN entry should be considered applicable to an endpoint.

    "Applies" is intentionally stricter than "device GUID is listed":
    - membership: endpoint GUID must be in entry.devices
    - flow: Render/Capture must be allowed by entry.flows
    - existence: we probe HKCU for the value name under FxProperties/Properties

    Why probe existence:
    Some drivers only create their vendor keys after the user toggles the setting at least once.
    An INI entry might list a device GUID (learned earlier), but if the driver hasn't
    initialized keys yet (fresh profile/new endpoint), writing may do nothing. Probing
    reduces false-positive "supported" results and improves troubleshooting.

    Note: the probe is HKCU-only because per-user configuration is commonly stored there.
    """
    guid = _extract_endpoint_guid_from_device_id(device_id)
    if not guid:
        return False

    devs = set((entry.get("devices") or []))
    if not devs or guid.lower() not in {d.lower() for d in devs}:
        return False

    flow_name = "Render" if str(flow).lower().startswith("r") else "Capture"
    if entry.get("flows") and flow_name not in entry["flows"]:
        return False

    value_name = (entry.get("value_name") or "").strip().lower()
    if not value_name:
        return False

    # Probe both canonical containers; the INI "subkey" tells us where to write,
    # but existence checks are forgiving because drivers sometimes mirror values.
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
    Apply a vendor entry by writing its DWORD under the *learned* subkey.

    - Uses entry.subkey to hit the exact place the learn flow observed changes
      ("FxProperties" vs "Properties"). Some drivers only honor one of these.
    - Writes across entry.hives in order (HKCU and/or HKLM). HKLM writes may require
      elevation depending on system policy and install context.

    Returns True if at least one hive write succeeded.
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
            # HKLM can fail without Admin; HKCU can fail if key doesn't exist yet.
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

    # Ensure directory exists (can be LocalAppData fallback).
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
        # Membership list is filled later (append-guid helper); kept empty here.
        "devices = ",
    ]
    with open(ini_path, "a", encoding="utf-8", errors="replace") as f:
        f.write("\n".join(lines) + "\n")

def _append_fx_ini_entry_multi(ini_path, section_name, fx_name, device_name, writes, notes=""):
    """
    Append an FX multi-write section.

    Multi-write sections store *exact* enable/disable payloads per registry value,
    including REG_BINARY blobs. This allows robust reproduction of a driver's
    effect toggles that can't be represented as a simple DWORD flip.

    See _load_vendor_db_split() doc comment for schema and write{i}_devices semantics.
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
    """
    Read a vendor-controlled state as True/False/None.

    - Multi-write FX:
        Uses decider/quorum evaluation (_read_decider_state) because state might be
        represented by multiple keys and not all toggles are always readable.
    - MAIN entries and legacy (single-DWORD) FX:
        Reads the learned DWORD from the learned subkey in HKCU first, then HKLM.
        Returns True if it equals entry.enable, False if equals entry.disable.

    Note: this "slow" reader is used when correctness is preferred over speed.
    GUI fast polling uses _fast_read_vendor_entry_state instead.
    """
    if entry.get("type") == "fx" and entry.get("multi_write"):
        return _read_decider_state(entry, device_id, flow)

    val_name = (entry.get("value_name") or "").strip().lower()
    if not val_name:
        return None

    # Write/read location is remembered during learn.
    subkey = (entry.get("subkey") or "FxProperties").strip()
    base = _endpoint_base_path(device_id, flow, subkey)
    if not base:
        return None

    # Prefer the hives in the INI's declared order; default to HKCU then HKLM.
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
    Poll the vendor-controlled state until it reflects expected_enabled.

    Verification is intentionally "consecutive reads" instead of a single read:
    registry updates can lag behind UI/driver state, especially with MMDevices.
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
    Try to apply a MAIN toggle by scanning INI "main" entries.

    IMPORTANT: This function must NOT consider FX entries:
    - FX toggles are effect-specific and may have different keys/payloads than
      the main SysFX switch.
    """
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
    Return the first MAIN vendor entry that:
      1) lists the endpoint GUID in devices AND
      2) appears to exist in HKCU for this endpoint (probe),
    otherwise fall back to membership-only match.

    This "exists first" heuristic improves correctness on systems where drivers
    lazily create vendor keys only after first toggle.
    """
    db = _load_vendor_db_split(ini_path)
    guid = _extract_endpoint_guid_from_device_id(device_id)
    if not guid:
        return None
    main_entries = db.get("main") or []
    flow_name = "Render" if str(flow).lower().startswith("r") else "Capture"

    for entry in main_entries:
        if _vendor_entry_applies(entry, device_id, flow_name):
            return entry

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
    # Dedupe for FX is based on the "payload" (writes or value_name+dwords),
    # not the human fx_name label. Multiple effect names could theoretically map
    # to the same underlying vendor knob, but we treat fx_name as metadata.
    if a.get("multi_write") and b.get("multi_write"):
        def _wkey(w):
            return (
                (w.get("hive") or "").upper(),
                (w.get("subkey") or "").strip().lower(),
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

    a_val = (a.get("value_name") or "").strip().lower()
    b_val = (b.get("value_name") or "").strip().lower()

    def _ed_pair(entry):
        if "enable" in entry or "disable" in entry:
            return str(entry.get("enable")).strip(), str(entry.get("disable")).strip()
        return str(entry.get("dword_enable")).strip(), str(entry.get("dword_disable")).strip()

    ae, ad = _ed_pair(a)
    be, bd = _ed_pair(b)
    return (a_val == b_val) and (ae == be) and (ad == bd)

import hashlib

def _norm_write_item(w):
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
    # Canonical identity for a multi-write bucket: payload + verification tuning.
    nw = sorted((_norm_write_item(w) for w in (writes or [])))
    try:
        di = int(decider_index or 1)
    except Exception:
        di = 1
    try:
        qt = float(quorum_threshold or 0.60)
    except Exception:
        qt = 0.60
    return ("fx-multi", tuple(nw), di, round(qt, 6))

def _fx_canonical_key_single(value_name, enable, disable):
    return ("fx-single",
            (str(value_name or "").strip().lower(),),
            int(enable), int(disable))

def _canonical_section_name_from_key(key_tuple):
    # Stable section names prevent churn when merging learned FX across devices.
    h = hashlib.sha1(repr(key_tuple).encode("utf-8", "replace")).hexdigest()[:16]
    return f"fx_{h}"

def _write_applies_to_guid(w, guid_lc: str) -> bool:
    """
    Interpret write{i}_devices semantics for multi-write FX:
      - None  => universal (applies to all devices)
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
# These helpers are designed for GUI polling where latency matters.
# They intentionally avoid COM calls and avoid wide registry scans.

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
    Return registry key last-write timestamp for tie-breaking.

    Why we need this:
    Some drivers mirror values across HKCU/HKLM or update them asynchronously.
    It's possible to read conflicting values. The "newer" key often represents
    the authoritative (most recently written) state.
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
    Type-aware equality check used by fast FX comparisons.
    expected_type_name is a string (REG_DWORD|REG_SZ|REG_BINARY).
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
    return False

def _fast_read_vendor_entry_state(entry, device_id, flow):
    """
    FAST state read (True/False/None) with minimal I/O.

    Behavior:
    - FX multi_write:
        Pick the "best" applicable write (filtered by write{i}_devices) and compare
        recorded hive vs alternate hive. If both are readable but disagree, use the
        newer key last-write time as a tie-breaker.
    - MAIN / legacy single-DWORD:
        Read the learned value in allowed hives. If both are readable but disagree,
        prefer the hive with newer key last-write time. Tie defaults to HKCU.

    Why last-write tie-break exists:
    Many drivers keep both hives populated, but only one is actually honored.
    The most recently written value usually indicates the active control path.
    """
    try:
        # FX multi-write: decider-ish fast read (not full quorum), filtered by GUID.
        if entry.get("type") == "fx" and entry.get("multi_write"):
            all_writes = entry.get("writes") or []
            if not all_writes:
                return None

            guid_lc = _guid_of(device_id)
            writes = [w for w in all_writes if _write_applies_to_guid(w, guid_lc)]
            if not writes:
                return None

            # Prefer the most "signal-like" write:
            # - FxProperties over Properties
            # - REG_DWORD over REG_BINARY/REG_SZ
            # - 0/1 flips are strongest (common vendor boolean representation)
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
                # Disagree -> pick the newer key.
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

        # MAIN / legacy single-DWORD:
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

        configured = entry.get("hives") or []
        allowed = {str(h).strip().upper() for h in configured if isinstance(h, str)}
        if not allowed:
            allowed = {"HKCU", "HKLM"}

        state = {}
        lastw = {}

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
            # Disagree -> newer key wins (ties prefer HKCU).
            tcu = lastw.get("HKCU")
            tlm = lastw.get("HKLM")
            try:
                if isinstance(tcu, int) and isinstance(tlm, int):
                    if tlm > tcu:
                        return lm
                    elif tcu > tlm:
                        return cu
                    else:
                        return cu
            except Exception:
                pass
            return cu

        return None
    except Exception:
        return None

def _fast_get_enhancements_state(device_id, flow):
    """
    FAST state read for the MAIN enhancements toggle:
    - Find the first applicable MAIN entry in the INI
    - Probe the vendor DWORD using the fast hive tie-break logic

    Returns True/False/None.
    """
    e = _find_first_vendor_entry(device_id, flow, ini_path=_vendor_ini_default_path())
    if not e:
        return None
    return _fast_read_vendor_entry_state(e, device_id, flow)

def _append_guid_to_section(ini_path, section_name, guid_lc):
    """
    INI maintenance helper:
    - Ensures the endpoint GUID is listed in 'devices =' under [section_name].

    This is used for dedupe/merge:
    if we learn an identical toggle payload as an existing section, we reuse it and
    simply attach the new endpoint GUID to the devices membership list.
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
                sec_end = i
                break
        elif sec_start is not None and sec_end == len(lines):
            continue
    if sec_start is None:
        new = []
        if lines and not lines[-1].endswith(("\n", "\r")):
            new.append("\n")
        new.append(f"{sec_hdr}\n")
        new.append(f"devices = {guid_lc}\n")
        lines.extend(new)
    else:
        devices_idx = None
        guid_set = None
        for i in range(sec_start + 1, sec_end):
            m = re.match(r"^\s*devices\s*=\s*(.*)$", lines[i], flags=re.IGNORECASE)
            if m:
                devices_idx = i
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
                insert_at = sec_end
                if insert_at > 0 and not lines[insert_at - 1].endswith(("\n", "\r")):
                    lines.insert(insert_at, "\n")
                    insert_at += 1
                lines.insert(insert_at, new_line)
    with open(ini_path, "w", encoding="utf-8", errors="replace") as f:
        f.writelines(lines)

def _append_guid_to_write_devices(ini_path, section_name, write_index, guid_lc):
    """
    Add an endpoint GUID to a specific multi-write toggle's write{i}_devices list.

    This is the primary merge mechanism for multi-write FX buckets:
    if we discover a write block with identical identity+payload, we attach the GUID
    instead of duplicating the write.
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
    existing = None
    for i in range(sec_start + 1, sec_end):
        m = key_pat.match(lines[i])
        if m:
            devices_idx = i
            txt = m.group(1).strip()
            if not txt:
                existing = []
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

def _remove_guid_from_write_devices(ini_path, section_name, write_index, guid_lc):
    """
    Remove guid_lc from write{write_index}_devices.

    Important semantic:
    If the resulting list becomes empty, we keep the line present but empty:
      write{i}_devices =
    Empty means "applies to nobody" (explicitly disabled), which is *not* the same as
    removing the line (which would mean "universal" and apply to all devices).
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
    for i in range(sec_start + 1, sec_end):
        m = key_pat.match(lines[i])
        if not m:
            continue
        txt = m.group(1).strip()
        cur = [x.strip().lower() for x in txt.split(",") if x.strip()] if txt else []
        cur = [x for x in cur if x != guid_lc.lower()]
        new_line = f"write{write_index}_devices = {','.join(cur)}\n" if cur else f"write{write_index}_devices = \n"
        lines[i] = new_line
        break
    with open(ini_path, "w", encoding="utf-8", errors="replace") as f:
        f.writelines(lines)

def _find_write_index_by_payload(ini_path, section_name, w):
    """
    Helper for merge/debug:
    locate which write{i} in an existing bucket matches a given identity+payload.
    """
    db = _load_vendor_db_split(ini_path)
    target = None
    for e in (db.get("fx") or []):
        if e.get("name") == section_name:
            target = e
            break
    if not target:
        return None

    def _same(a, b):
        return (
            (a.get("hive","").upper() == b.get("hive","").upper()) and
            (str(a.get("subkey","")).strip().lower() == str(b.get("subkey","")).strip().lower()) and
            (str(a.get("name","")).strip().lower() == str(b.get("name","")).strip().lower()) and
            (str(a.get("type_enable","")).upper() == str(b.get("type_enable","")).upper()) and
            (str(a.get("type_disable","")).upper() == str(b.get("type_disable","")).upper()) and
            (str(a.get("enable","")).strip() == str(b.get("enable","")).strip()) and
            (str(a.get("disable","")).strip() == str(b.get("disable","")).strip())
        )
    for idx, cw in enumerate(target.get("writes") or [], start=1):
        if _same(cw, w):
            return idx
    return None

def _cleanup_conflicting_toggles(ini_path, section_name, guid_lc, keep_idx, keep_write):
    """
    Conflict cleanup during FX merge:

    It's possible for a bucket to contain multiple write blocks with the same identity:
      (hive, subkey, name)
    but different payload (enable/disable values). A single device GUID must not be
    attached to two conflicting payloads for the same identity, otherwise runtime
    could write contradictory values.

    This removes guid_lc from any "same identity" write blocks except keep_idx.
    """
    db = _load_vendor_db_split(ini_path)
    target = None
    for e in (db.get("fx") or []):
        if e.get("name") == section_name:
            target = e
            break
    if not target:
        return

    def _same_identity(a, b):
        return (
            (a.get("hive","").upper() == b.get("hive","").upper()) and
            (str(a.get("subkey","")).strip().lower() == str(b.get("subkey","")).strip().lower()) and
            (str(a.get("name","")).strip().lower() == str(b.get("name","")).strip().lower())
        )

    for idx, cw in enumerate(target.get("writes") or [], start=1):
        if idx == keep_idx:
            continue
        if _same_identity(cw, keep_write):
            _remove_guid_from_write_devices(ini_path, section_name, idx, guid_lc)

def _sanitize_ini_section_name(value_name: str):
    base = re.sub(r'[^A-Za-z0-9_,\-{}]+', "_", value_name)
    return f"vendor_{base}"

def _append_vendor_ini_entry_if_missing(ini_path, section_name, value_name, dword_enable, dword_disable,
                                        flows="Render,Capture", hives="HKCU,HKLM", notes="", subkey="FxProperties"):
    """
    Append a MAIN vendor section only if missing.

    We record 'subkey' (FxProperties vs Properties) because learn can detect the flip
    in either location; runtime reads/writes must target the authoritative one.
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
    lines.append(f"subkey = {subkey_norm}")
    if notes:
        lines.append(f"notes = {notes}")
    lines.append("devices = ")
    text = "\n".join(lines) + "\n"
    with open(ini_path, "a", encoding="utf-8", errors="replace") as f:
        f.write(text)
    return "appended"

def _build_vendor_ini_snippet(target, snapA, snapB, diffs, section_name=None):
    """
    Build a suggested MAIN vendor INI snippet from observed DWORD flips.

    We also record the subkey (FxProperties vs Properties) where the flip occurred,
    because that's the best signal for where the driver is reading/writing.
    """
    cands = []
    for f in diffs.get("dword_flips", []):
        name = str(f.get("name",""))
        subkey = str(f.get("subkey",""))
        hive = str(f.get("hive",""))
        if not (name.startswith("{") and "}" in name and "," in name):
            continue
        before = int(f.get("before"))
        after  = int(f.get("after"))
        cands.append({
            "hive": hive, "flow": None, "subkey": subkey, "name": name,
            "before": before, "after": after
        })
    if not cands:
        return None, None
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
    snippet.append(f"subkey = {picked_subkey}")
    snippet.append(f"notes = {notes}")
    snippet.append("devices = ")
    return "\n".join(snippet) + "\n", pick

def _collect_registry_samples(device_id, repeats=3, delay=0.15):
    """
    Collect several registry-only samples for the current device state.

    Why we sample:
    UI/driver toggles can generate unrelated MMDevices noise (timestamps, other props).
    Sampling multiple times and keeping only stable keys reduces false positives,
    especially for REG_BINARY blobs.
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
    Build a stability-filtered map from multiple registry dumps.

    Only keys that remain identical (type + dataRaw) across ALL samples are kept.
    This is used to reduce noise when learning FX payloads (drivers may update unrelated
    properties while UI is open, or while effects initialize).
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
      - include only keys present in both A and B
      - require same registry type but different payload
      - encode payload into INI-friendly form

    Ordering/scoring:
    We bias the write list so that write1 is a strong indicator:
      - FxProperties is preferred (common MMDevices vendor store)
      - REG_DWORD 0/1 flips are preferred (cleanest "boolean" representation)
      - REG_BINARY is deprioritized (more noise-prone and harder to reason about)
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
    Manual learn for MAIN enhancements toggle.

    Flow:
      1) Show stern warning because learning writes a *persistent* toggle to disk.
      2) User toggles Windows UI; tool captures A (enabled) and B (disabled) snapshots.
      3) Find simple DWORD flip candidates; build an INI section recording:
         - value_name
         - enable/disable values
         - *subkey* where flip was found (FxProperties vs Properties)
      4) Dedupe:
         If an identical toggle already exists, just append this endpoint GUID to that
         section's devices list instead of creating a new section.

    Safety:
      - Requires explicit confirmation text, unless AUDIOCTL_LEARN_CONFIRMED=1 is set.
        The GUI sets this env var because it shows its own warning dialog.
    """
    import sys
    dev_id = target["id"]
    flow   = target["flow"]
    name   = target["name"]
    ini_path = ini_path or _vendor_ini_default_path()

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

    # Dedupe strategy: if we already have an identical payload section, reuse it and
    # just attach this endpoint GUID. This keeps the INI from exploding with duplicates.
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
            subkey=(picked.get("subkey") if picked else "FxProperties")
        )
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
    (Legacy/optional) Auto-learn attempt:
    this path programmatically toggles Windows properties and then diffs registry.
    Kept for backward compatibility; the project favors manual learn for reliability.
    """
    import sys
    dev_id = target["id"]
    flow   = target["flow"]
    name   = target["name"]
    ini_path = ini_path or _vendor_ini_default_path()
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
        # Registry writes to HKLM may require Admin; this is why we check is_admin().
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
    Best-effort status read for display (vendor-only).
    Returns True/False/None.
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
    """
    Build the MMDevices base key path under a specific endpoint GUID.

    Pattern (under HKCU or HKLM):
      SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\MMDevices\\Audio\\<Render|Capture>\\{EndpointGuid}\\<subkey>

    subkey is typically "FxProperties" or "Properties".
    """
    guid = _extract_endpoint_guid_from_device_id(device_id)
    if not guid:
        return None
    flow_name = "Render" if str(flow).lower().startswith("r") else "Capture"
    return rf"SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\{flow_name}\{guid}\{subkey}"

def _perform_multi_writes(entry, device_id, flow, enable):
    """
    Apply a multi-write FX entry:
      - Writes every write{i} whose write{i}_devices scope includes this endpoint GUID.
      - Returns True only if *all applicable writes* succeed.

    HKLM writes may require Admin depending on system ACLs.
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
    """
    Determine state for a multi-write FX entry using quorum voting.

    Why decider/quorum:
    - Not all writes are always present/readable on all devices (or all hives).
    - Drivers may mirror or partially update values.
    We count "votes" from readable toggles and accept True/False if one side meets
    quorum_threshold.

    If quorum is inconclusive, we fall back to a "best signal" read (FxProperties,
    REG_DWORD preferred) to at least return a state when possible.
    """
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
    # FX buckets are keyed by fx_name (case-insensitive) and may contain multiple write blocks.
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
    # Stable section name based on fx_name, so multiple devices contribute to the same bucket.
    import hashlib
    key = (fx_name or "").strip().lower()
    h = hashlib.sha1(key.encode("utf-8", "replace")).hexdigest()[:16]
    return f"fx_{h}"

def _append_new_write_to_section(ini_path, section_name, write_dict, guid_lc):
    """
    Append a new write{i}_* block to an existing multi-write FX section and bump write_count.

    This is used when learning finds a new identity+payload that doesn't match any existing
    write block in the bucket.
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
    Remove an FX association for a specific device GUID without necessarily deleting the section.

    Why we don't delete the whole section:
    FX "buckets" are shared across multiple endpoints. Deleting the section would remove
    the effect for all devices that use it.

    Behavior:
    - Remove GUID from section-level 'devices' union list.
    - For each write{i}:
        - If write{i}_devices exists, remove GUID from that list.
        - If write{i}_devices is missing (universal), convert it to an explicit list
          of remaining devices to effectively "exclude" the removed GUID.
    - Keep empty write{i}_devices lines as explicit "applies to nobody".
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
            line_txt = f"write{i_idx}_devices = \n"
        else:
            line_txt = f"write{i_idx}_devices = {','.join(sorted(set(d.lower() for d in dev_list)))}\n" if dev_list else f"write{i_idx}_devices = \n"
        for j in range(sec_start + 1, sec_end):
            if pat.match(lines[j] or ""):
                lines[j] = line_txt
                return
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
    remaining_bucket_devs = [d for d in cur_devices if d != guid_lc]
    max_scan = write_count if write_count > 0 else 256
    for i_idx in range(1, max_scan + 1):
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
            # Universal write becomes explicitly scoped to remaining devices, so the
            # removed GUID is excluded without changing payloads for other devices.
            if remaining_bucket_devs:
                _set_write_devices(i_idx, remaining_bucket_devs)
                writes_changed += 1
            else:
                _set_write_devices(i_idx, [])
                writes_changed += 1
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
    Learn an FX toggle and persist it into vendor_toggles.ini.

    Two-pass A/B concept:
    Some drivers "initialize" keys on the first toggle (creating values or changing
    additional properties). To avoid learning one-time initialization noise:
      - Pass 1 (A/B) primes the driver.
      - Pass 2 (A2/B2) is treated as authoritative for extracting the write set.

    Merge strategy into an existing FX bucket:
      - If an identity+payload matches an existing write block, append this GUID to
        that block's write{i}_devices.
      - If it doesn't match, append a new write block.
      - Conflict cleanup ensures a GUID isn't attached to two blocks with the same
        identity (hive/subkey/name) but different payload.

    Returns (ok, info_or_error).
    """
    import re
    ini_path = ini_path or _vendor_ini_default_path()
    guid_lc = _guid_of(target["id"])

    useA = snapA2 if isinstance(snapA2, dict) else snapA
    useB = snapB2 if isinstance(snapB2, dict) else snapB

    try:
        stableA = _stable_registry_map([useA.get("registry") or []])
    except Exception as e:
        return False, f"Failed to process snapshot A: {e}"

    try:
        samplesB = _collect_registry_samples(target["id"], repeats=3, delay=0.18)
        samplesB.insert(0, useB.get("registry") or [])
        stableB = _stable_registry_map(samplesB)
    except Exception as e:
        return False, f"Failed to process snapshot B: {e}"

    writes = _build_fx_multiwrite_from_stable_maps(target, stableA, stableB)
    safe_device_name = re.sub(r'[^A-Za-z0-9_\- ]+', '_', target["name"])
    notes = f"Learned FX '{fx_name}' for '{target['name']}' ({target['flow']}); second A/B pass; stability-filtered"

    if writes:
        bucket = _find_fx_bucket_section_name(ini_path, fx_name)
        if bucket is None:
            # First time seeing this fx_name: create a canonical bucket and seed it
            # with write{i}_devices = this guid.
            for w in writes:
                w.setdefault("devices", None)
            section_name = _canonical_fx_bucket_name(fx_name)
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

        _append_guid_to_section(ini_path, bucket, guid_lc)
        return True, {
            "iniPath": ini_path,
            "section": bucket,
            "fx_name": fx_name,
            "multi_write": True,
            "write_count": None
        }

    # Fallback: if we couldn't build a robust multi-write set, fall back to a simple
    # DWORD flip candidate (legacy model). Some drivers still use a single DWORD.
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
        msg = str(e or "").lower()
        if "already exists" in msg:
            try:
                _append_guid_to_section(ini_path, section_name, guid_lc)
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
    Apply the MAIN enhancements toggle using vendor methods only.

    Runtime policy:
    - If the device does not have a learned vendor entry in vendor_toggles.ini,
      we return failure with "no-vendor-method" rather than falling back to Windows'
      Disable_SysFx. This keeps runtime behavior predictable and avoids toggling a
      knob the driver may ignore.

    Returns:
      (ok: bool, verified_by: str, final_state: bool|None)
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
    Return True if this endpoint GUID is listed by any MAIN vendor entry in the INI.

    Note: This is membership-only (does not probe existence). The CLI uses this to
    decide whether to offer "vendor-only reads" and whether a toggle is potentially
    available. Detailed applicability checks are done elsewhere.
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

    - Multi-write FX:
        Writes all applicable write{i} blocks, then verifies using decider/quorum state.
    - Legacy single-DWORD FX:
        Writes the vendor DWORD and verifies via polling.

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
    wrote = _set_vendor_entry_state(entry, device_id, flow, enable)
    if not wrote:
        return False, None, None
    ok, state = _verify_vendor_entry(entry, device_id, flow, enable,
                                     timeout=2.5, interval=0.2, consecutive=2)
    verified_by = f"vendor-fx:{entry.get('fx_name','')}"
    return ok, verified_by if ok else None, state
