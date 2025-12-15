# --- comtypes compatibility shim (MUST be at the VERY TOP, before ANY other imports including pycaw) ---
# This minimal shim addresses known PyInstaller bundling issues with comtypes.automation.
# It ensures core constants and a PROPVARIANT fallback are available when needed by pycaw.
# It should make other parts of pycaw (like 'list') work. The 'listen' functions will now
# use raw ctypes vtable calls for maximum robustness, bypassing comtypes.gen.propsys.
try:
    import comtypes.automation as _automation
    # Ensure PROPVARIANT alias exists, using tagPROPVARIANT as a common fallback.
    if not hasattr(_automation, "PROPVARIANT") and hasattr(_automation, "tagPROPVARIANT"):
        _automation.PROPVARIANT = _automation.tagPROPVARIANT
    # Ensure standard VT_ constants are present.
    if not hasattr(_automation, "VT_LPWSTR"):
        _automation.VT_LPWSTR = 31  # Value for VT_LPWSTR
    if not hasattr(_automation, "VT_BOOL"):
        _automation.VT_BOOL = 11   # Value for VT_BOOL
    # Ensure VARIANT_TRUE/FALSE for boolean properties (used in manual PROPVARIANT)
    if not hasattr(_automation, "VARIANT_TRUE"):
        _automation.VARIANT_TRUE = -1
    if not hasattr(_automation, "VARIANT_FALSE"):
        _automation.VARIANT_FALSE = 0
except Exception as e:
    import sys
    print(f"WARNING: Global comtypes compatibility shim failed during initial import: {e}", file=sys.stderr)
# ------------------------------------------------------------------------------------------------
import argparse
import ctypes
import json
import re
import sys
import time
import warnings
import io

# (os and sys are already imported globally by other parts of the script)

import os, sys # Explicitly ensure they are here for this snippet's context
def resource_path(name: str):
    # Works both when frozen (PyInstaller) and when running from source
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, name)

from contextlib import redirect_stderr
from ctypes import POINTER, byref
from comtypes import CLSCTX_ALL, CoCreateInstance, GUID, IUnknown, COMMETHOD, HRESULT, CoInitialize, CoUninitialize
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume, IMMDeviceEnumerator
from pycaw.constants import CLSID_MMDeviceEnumerator
DEVICE_STATE_ACTIVE = 0x00000001
DEVICE_STATE_ALL = 0x0000000F  # active | disabled | notpresent | unplugGED
E_RENDER = 0  # Playback
E_CAPTURE = 1  # Recording
E_CONSOLE = 0
E_MULTIMEDIA = 1
E_COMMUNICATIONS = 2
ROLES = {
    "console": E_CONSOLE,
    "multimedia": E_MULTIMEDIA,
    "communications": E_COMMUNICATIONS,
    "all": "all",
}
# Map Windows device state flags for cleaner output
DEVICE_STATES = {
    0x00000001: "active",
    0x00000002: "disabled",
    0x00000004: "notpresent",
    0x00000008: "unplugged"
}
def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except Exception:
        return False
# STGM flags used by OpenPropertyStore
STGM_READ  = 0x00000000
STGM_WRITE = 0x00000001
def set_listen_to_device_ps(capture_device_id, enable, render_device_id=None):
    """
    Enable/disable 'Listen to this device' using raw IPropertyStore vtable calls.
    Assumes COM is already initialized on this thread (GUI or CLI does that).
    """
    import ctypes
    from ctypes import POINTER, byref, wintypes
    import comtypes.automation as automation
    from comtypes import GUID, CLSCTX_ALL, CoCreateInstance
    from pycaw.constants import CLSID_MMDeviceEnumerator
    from pycaw.pycaw import IMMDeviceEnumerator
    try:
        HRESULT_T = wintypes.HRESULT
    except Exception:
        HRESULT_T = ctypes.c_long
    CALL = ctypes.WINFUNCTYPE if ctypes.sizeof(ctypes.c_void_p) == 4 else ctypes.CFUNCTYPE
    def _hrx(hr): return f"0x{ctypes.c_uint(hr).value:08X}"
    def _raw_ptr(p): return ctypes.cast(p, ctypes.c_void_p).value
    if hasattr(automation, "PROPVARIANT"):
        PROPVARIANT = automation.PROPVARIANT
    elif hasattr(automation, "tagPROPVARIANT"):
        PROPVARIANT = automation.tagPROPVARIANT
    else:
        class _PVU(ctypes.Union):
            _fields_ = [
                ("pwszVal", ctypes.c_wchar_p),
                ("boolVal", ctypes.c_short),
                ("punkVal", ctypes.c_void_p),
                ("ulVal", ctypes.c_ulong),
                ("uhVal", ctypes.c_ulonglong),
            ]
        class PROPVARIANT(ctypes.Structure):
            _anonymous_ = ("data",)
            _fields_ = [
                ("vt", ctypes.c_ushort),
                ("wReserved1", ctypes.c_ushort),
                ("wReserved2", ctypes.c_ushort),
                ("wReserved3", ctypes.c_ushort),
                ("data", _PVU),
            ]
    VT_BOOL = getattr(automation, "VT_BOOL", 11)
    VT_LPWSTR = getattr(automation, "VT_LPWSTR", 31)
    VARIANT_TRUE = -1
    VARIANT_FALSE = 0
    class PROPERTYKEY(ctypes.Structure):
        _fields_ = (("fmtid", GUID), ("pid", wintypes.DWORD))
    class IPropertyStoreRaw(ctypes.Structure):
        pass
    PIPS = POINTER(IPropertyStoreRaw)
    _POGUID_ = POINTER(GUID)
    _PPVoid_ = POINTER(ctypes.c_void_p)
    _PDWORD_ = POINTER(wintypes.DWORD)
    _PPROPERTYKEY_ = POINTER(PROPERTYKEY)
    _PPROPVARIANT_ = POINTER(PROPVARIANT)
    QueryInterfaceProto = CALL(HRESULT_T, PIPS, _POGUID_, _PPVoid_)
    AddRefProto         = CALL(ctypes.c_ulong, PIPS)
    ReleaseProto        = CALL(ctypes.c_ulong, PIPS)
    GetCountProto       = CALL(HRESULT_T, PIPS, _PDWORD_)
    GetAtProto          = CALL(HRESULT_T, PIPS, wintypes.DWORD, _PPROPERTYKEY_)
    GetValueProto       = CALL(HRESULT_T, PIPS, _PPROPERTYKEY_, _PPROPVARIANT_)
    SetValueProto       = CALL(HRESULT_T, PIPS, _PPROPERTYKEY_, _PPROPVARIANT_)
    CommitProto         = CALL(HRESULT_T, PIPS)
    class IPropertyStoreVTBL(ctypes.Structure):
        _fields_ = [
            ("QueryInterface", QueryInterfaceProto),
            ("AddRef", AddRefProto),
            ("Release", ReleaseProto),
            ("GetCount", GetCountProto),
            ("GetAt", GetAtProto),
            ("GetValue", GetValueProto),
            ("SetValue", SetValueProto),
            ("Commit", CommitProto),
        ]
    IPropertyStoreRaw._fields_ = [("lpVtbl", POINTER(IPropertyStoreVTBL))]
    propsys = ctypes.OleDLL("propsys.dll")
    ole32 = ctypes.OleDLL("ole32.dll")
    have_helpers = True
    try:
        InitPropVariantFromBoolean = propsys.InitPropVariantFromBoolean
        InitPropVariantFromBoolean.restype = HRESULT_T
        InitPropVariantFromBoolean.argtypes = (wintypes.BOOL, POINTER(PROPVARIANT))
        InitPropVariantFromString = propsys.InitPropVariantFromString
        InitPropVariantFromString.restype = HRESULT_T
        InitPropVariantFromString.argtypes = (wintypes.LPCWSTR, POINTER(PROPVARIANT))
    except (AttributeError, OSError):
        have_helpers = False
    PropVariantClear = ole32.PropVariantClear
    PropVariantClear.restype = HRESULT_T
    PropVariantClear.argtypes = (POINTER(PROPVARIANT),)
    CoTaskMemAlloc = ole32.CoTaskMemAlloc
    CoTaskMemAlloc.restype = ctypes.c_void_p
    CoTaskMemAlloc.argtypes = (ctypes.c_size_t,)
    def _pv_from_bool_local(value: bool):
        pv = PROPVARIANT()
        if have_helpers:
            hr = InitPropVariantFromBoolean(VARIANT_TRUE if value else VARIANT_FALSE, byref(pv))
            if hr != 0:
                raise OSError(f"InitPropVariantFromBoolean failed: {_hrx(hr)}")
        else:
            pv.vt = VT_BOOL
            try:
                pv.boolVal = VARIANT_TRUE if value else VARIANT_FALSE
            except AttributeError:
                pass
        return pv
    def _pv_from_string_local(s: str):
        pv = PROPVARIANT()
        s_val = s if s is not None else ""
        if have_helpers:
            hr = InitPropVariantFromString(s_val, byref(pv))
            if hr != 0:
                raise OSError(f"InitPropVariantFromString failed: {_hrx(hr)}")
        else:
            nbytes = (len(s_val) + 1) * ctypes.sizeof(ctypes.c_wchar)
            buf = ctypes.create_unicode_buffer(s_val)
            ptr = CoTaskMemAlloc(nbytes)
            if not ptr:
                raise MemoryError("CoTaskMemAlloc failed")
            ctypes.memmove(ptr, ctypes.addressof(buf), nbytes)
            pv.vt = VT_LPWSTR
            pv.pwszVal = ctypes.cast(ptr, ctypes.c_wchar_p)
        return pv
    PKEY_LISTEN_ENABLE = PROPERTYKEY(GUID("{24dbb0fc-9311-4b3d-9cf0-18ff155639d4}"), 1)
    PKEY_LISTEN_PLAYBACKTHROUGH = PROPERTYKEY(GUID("{24dbb0fc-9311-4b3d-9cf0-18ff155639d4}"), 2)
    def _get_string_prop_local(ps_iface_ptr, pkey_obj):
        pv = PROPVARIANT()
        try:
            hr = ps_iface_ptr.contents.lpVtbl.contents.GetValue(ps_iface_ptr, byref(pkey_obj), byref(pv))
            if hr != 0 or getattr(pv, "vt", 0) != VT_LPWSTR:
                return None
            def _read_ptr_or_str(val):
                if isinstance(val, str):
                    return val
                if val:
                    try:
                        return ctypes.wstring_at(val)
                    except Exception:
                        return None
                return None
            s = _read_ptr_or_str(getattr(pv, "pwszVal", None))
            if not s:
                val = getattr(pv, "value", None)
                if val is not None:
                    s = _read_ptr_or_str(getattr(val, "pwszVal", None))
            if not s:
                data = getattr(pv, "data", None)
                if data is not None:
                    s = _read_ptr_or_str(getattr(data, "pwszVal", None))
            return s
        finally:
            PropVariantClear(byref(pv))
    enumerator = None
    pv_enable = None
    pv_target = None
    try:
        enumerator = CoCreateInstance(CLSID_MMDeviceEnumerator, interface=IMMDeviceEnumerator, clsctx=CLSCTX_ALL)
        dev = enumerator.GetDevice(capture_device_id)
        ps_unknown = dev.OpenPropertyStore(STGM_WRITE)
        ps_ptr_val = _raw_ptr(ps_unknown)
        if not ps_ptr_val:
            raise OSError("OpenPropertyStore returned null pointer for IPropertyStore.")
        ps_iface = ctypes.cast(ctypes.c_void_p(ps_ptr_val), PIPS)
        current_target_id = _get_string_prop_local(ps_iface, PKEY_LISTEN_PLAYBACKTHROUGH)
        pv_enable = _pv_from_bool_local(bool(enable))
        hr = ps_iface.contents.lpVtbl.contents.SetValue(ps_iface, byref(PKEY_LISTEN_ENABLE), byref(pv_enable))
        if hr != 0:
            raise OSError(f"IPropertyStore::SetValue(enable) failed: {_hrx(hr)}")
        target_string_to_set = None
        if render_device_id is None:
            if enable and not current_target_id:
                target_string_to_set = ""
        else:
            target_string_to_set = render_device_id
        if target_string_to_set is not None:
            pv_target = _pv_from_string_local(target_string_to_set)
            hr = ps_iface.contents.lpVtbl.contents.SetValue(ps_iface, byref(PKEY_LISTEN_PLAYBACKTHROUGH), byref(pv_target))
            if hr != 0:
                raise OSError(f"IPropertyStore::SetValue(target) failed: {_hrx(hr)}")
        hr = ps_iface.contents.lpVtbl.contents.Commit(ps_iface)
        if hr != 0:
            raise OSError(f"IPropertyStore::Commit failed: {_hrx(hr)}")
        return True
    except Exception as e:
        print(f"ERROR: set_listen_to_device_ps failed for '{capture_device_id}': {e}", file=sys.stderr)
        return False
    finally:
        try:
            if pv_enable is not None:
                PropVariantClear(byref(pv_enable))
        except Exception:
            pass
        try:
            if pv_target is not None:
                PropVariantClear(byref(pv_target))
        except Exception:
            pass

def _get_listen_to_device_status_ps(device_id):
    """
    Reads the 'Listen to this device' enable flag using IPropertyStore::GetValue via raw vtable (ctypes).
    Assumes COM is already initialized on this thread.
    Returns True/False/None.
    """
    import ctypes
    from ctypes import POINTER, byref, wintypes
    import comtypes.automation as automation
    from comtypes import GUID, CLSCTX_ALL, CoCreateInstance
    from pycaw.constants import CLSID_MMDeviceEnumerator
    from pycaw.pycaw import IMMDeviceEnumerator
    try:
        HRESULT_T = wintypes.HRESULT
    except Exception:
        HRESULT_T = ctypes.c_long
    CALL = ctypes.WINFUNCTYPE if ctypes.sizeof(ctypes.c_void_p) == 4 else ctypes.CFUNCTYPE
    def _raw_ptr(p): return ctypes.cast(p, ctypes.c_void_p).value
    if hasattr(automation, "PROPVARIANT"):
        PROPVARIANT = automation.PROPVARIANT
    elif hasattr(automation, "tagPROPVARIANT"):
        PROPVARIANT = automation.tagPROPVARIANT
    else:
        class _PVU(ctypes.Union):
            _fields_ = [("pwszVal", ctypes.c_wchar_p), ("boolVal", ctypes.c_short)]
        class PROPVARIANT(ctypes.Structure):
            _anonymous_ = ("data",)
            _fields_ = [
                ("vt", ctypes.c_ushort),
                ("wReserved1", ctypes.c_ushort),
                ("wReserved2", ctypes.c_ushort),
                ("wReserved3", ctypes.c_ushort),
                ("data", _PVU),
            ]
    VT_BOOL = getattr(automation, "VT_BOOL", 11)
    VARIANT_FALSE = getattr(automation, "VARIANT_FALSE", 0)
    class PROPERTYKEY(ctypes.Structure):
        _fields_ = (("fmtid", GUID), ("pid", wintypes.DWORD))
    class IPropertyStoreRaw(ctypes.Structure):
        pass
    PIPS = POINTER(IPropertyStoreRaw)
    _POGUID_ = POINTER(GUID)
    _PPVoid_ = POINTER(ctypes.c_void_p)
    _PDWORD_ = POINTER(wintypes.DWORD)
    _PPROPERTYKEY_ = POINTER(PROPERTYKEY)
    _PPROPVARIANT_ = POINTER(PROPVARIANT)
    QueryInterfaceProto = CALL(HRESULT_T, PIPS, _POGUID_, _PPVoid_)
    AddRefProto         = CALL(ctypes.c_ulong, PIPS)
    ReleaseProto        = CALL(ctypes.c_ulong, PIPS)
    GetCountProto       = CALL(HRESULT_T, PIPS, _PDWORD_)
    GetAtProto          = CALL(HRESULT_T, PIPS, wintypes.DWORD, _PPROPERTYKEY_)
    GetValueProto       = CALL(HRESULT_T, PIPS, _PPROPERTYKEY_, _PPROPVARIANT_)
    SetValueProto       = CALL(HRESULT_T, PIPS, _PPROPERTYKEY_, _PPROPVARIANT_)
    CommitProto         = CALL(HRESULT_T, PIPS)
    class IPropertyStoreVTBL(ctypes.Structure):
        _fields_ = [
            ("QueryInterface", QueryInterfaceProto),
            ("AddRef", AddRefProto),
            ("Release", ReleaseProto),
            ("GetCount", GetCountProto),
            ("GetAt", GetAtProto),
            ("GetValue", GetValueProto),
            ("SetValue", SetValueProto),
            ("Commit", CommitProto),
        ]
    IPropertyStoreRaw._fields_ = [("lpVtbl", POINTER(IPropertyStoreVTBL))]
    ole32 = ctypes.OleDLL("ole32.dll")
    PropVariantClear = ole32.PropVariantClear
    PropVariantClear.restype = HRESULT_T
    PropVariantClear.argtypes = (POINTER(PROPVARIANT),)
    PKEY_LISTEN_ENABLE = PROPERTYKEY(GUID("{24dbb0fc-9311-4b3d-9cf0-18ff155639d4}"), 1)
    pv = PROPVARIANT()
    try:
        enumerator = CoCreateInstance(CLSID_MMDeviceEnumerator, interface=IMMDeviceEnumerator, clsctx=CLSCTX_ALL)
        dev = enumerator.GetDevice(device_id)
        ps_unknown = dev.OpenPropertyStore(STGM_READ)
        ps_ptr_val = _raw_ptr(ps_unknown)
        if not ps_ptr_val:
            return None
        ps_iface = ctypes.cast(ctypes.c_void_p(ps_ptr_val), PIPS)
        hr = ps_iface.contents.lpVtbl.contents.GetValue(ps_iface, byref(PKEY_LISTEN_ENABLE), byref(pv))
        if hr != 0:
            return False
        if pv.vt == VT_BOOL:
            return pv.boolVal != VARIANT_FALSE
        return False
    except Exception as e:
        print(f"WARNING: Failed to read listen status via COM for '{device_id}': {e}", file=sys.stderr)
        return None
    finally:
        try:
            PropVariantClear(byref(pv))
        except Exception:
            pass

def _extract_endpoint_guid_from_device_id(device_id: str):
    """
    Extract the endpoint GUID (with braces) from a device id like:
      "{0.0.1.00000000}.{83a9be54-901e-4429-993b-c9088e3028a0}"
    Returns "{83a9be54-901e-4429-993b-c9088e3028a0}" or None.
    """
    try:
        # Use raw string literal for regex pattern to avoid SyntaxWarning
        m = re.search(r'\.\{([0-9A-Fa-f-]+)\}$', device_id)
        if not m:
            return None
        return "{" + m.group(1) + "}"
    except Exception:
        return None
def _read_listen_enable_from_registry(device_id: str):
    r"""
    Robustly read the 'Listen to this device' enable state from MMDevices.
    - Scans HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\Capture\{GUID}\(FxProperties|Properties)
    - Finds values whose name starts with the AudioEndpointSettings GUID "{24dbb0fc-9311-4b3d-9cf0-18ff155639d4}"
    - Prefers pid 1 (",1") if present, but will also accept any VT_BOOL value under that GUID if needed
    - Parses REG_DWORD (0/1) or REG_BINARY PROPVARIANT (VT_BOOL)
    Returns True/False or None if not found.
    """
    try:
        import winreg
    except Exception:
        return None
    guid = _extract_endpoint_guid_from_device_id(device_id)
    if not guid:
        return None
    base = r"SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\Capture" + "\\" + guid
    guid_base = "{24dbb0fc-9311-4b3d-9cf0-18ff155639d4}".lower()
    def _parse_bool_from_reg(val, typ):
        # REG_DWORD: 0/1
        if typ == winreg.REG_DWORD:
            try:
                return bool(int(val))
            except Exception:
                return None
        # REG_BINARY PROPVARIANT with VT_BOOL at offset 0x000B, boolVal at offset 8
        if typ == winreg.REG_BINARY:
            try:
                b = bytes(val)
                if len(b) >= 10:
                    vt = int.from_bytes(b[0:2], "little", signed=False)
                    if vt == 0x000B:
                        bool16 = int.from_bytes(b[8:10], "little", signed=True)
                        return bool16 != 0
            except Exception:
                return None
        # Occasionally stored as REG_SZ "0"/"1"/"true"/"false"
        if typ == winreg.REG_SZ:
            try:
                s = str(val).strip().lower()
                if s in ("1", "true", "yes", "on"):
                    return True
                if s in ("0", "false", "no", "off"):
                    return False
            except Exception:
                return None
        return None
    preferred = None  # pid==1 match
    fallback_any = None  # any VT_BOOL under the same GUID
    for sub in ("FxProperties", "Properties"):
        try:
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, base + "\\" + sub, 0, winreg.KEY_READ)
        except OSError:
            continue
        try:
            i = 0
            while True:
                try:
                    name, val, typ = winreg.EnumValue(key, i)
                    i += 1
                except OSError:
                    break
                name_l = name.lower()
                # Only consider values for our AudioEndpointSettings GUID
                if not name_l.startswith(guid_base):
                    continue
                pid = None
                if "," in name_l:
                    try:
                        pid = int(name_l.split(",", 1)[1])
                    except Exception:
                        pid = None
                parsed = _parse_bool_from_reg(val, typ)
                if parsed is None:
                    continue
                if pid == 1:
                    preferred = parsed
                    # Early exit if we found the exact pid
                    break
                # Keep a fallback from the same GUID if no pid==1 found
                if fallback_any is None:
                    fallback_any = parsed
            if preferred is not None:
                break
        finally:
            try:
                winreg.CloseKey(key)
            except Exception:
                pass
    if preferred is not None:
        return preferred
    if fallback_any is not None:
        return fallback_any
    return None
def _verify_listen_via_registry(device_id: str, expected_enabled: bool, timeout=2.0, interval=0.15):
    """
    Poll the registry for up to 'timeout' seconds until the 'Listen' checkbox matches 'expected_enabled'.
    Returns (True, state) if verified, else (False, last_state_or_None).
    """
    deadline = time.time() + timeout
    last_state = None
    while time.time() < deadline:
        state = _read_listen_enable_from_registry(device_id)
        last_state = state
        if state is not None and state == expected_enabled:
            return True, state
        time.sleep(interval)
    return False, last_state
def _reemit_non_error_stderr(buf_text: str):
    """
    Re-emit only non-error lines (e.g., INFO) from captured stderr.
    Suppresses lines starting with 'ERROR:' (ignoring leading whitespace).
    """
    try:
        for line in buf_text.splitlines(True):
            if not line.lstrip().lower().startswith("error:"):
                sys.stderr.write(line)
    except Exception:
        pass
def enum_endpoints(flow, state_mask):
    enumerator = CoCreateInstance(CLSID_MMDeviceEnumerator, IMMDeviceEnumerator, CLSCTX_ALL)
    collection = enumerator.EnumAudioEndpoints(flow, state_mask)
    return enumerator, collection
# CORRECTED: get_default_ids (assigns None to role_name string key, not role_val numeric key)
def get_default_ids(enumerator):
    defaults = {"Render": {}, "Capture": {}}
    for flow_name, flow in [("Render", E_RENDER), ("Capture", E_CAPTURE)]:
        for role_name, role_val in [
            ("console", E_CONSOLE),
            ("multimedia", E_MULTIMEDIA),
            ("communications", E_COMMUNICATIONS),
        ]:
            try:
                dev = enumerator.GetDefaultAudioEndpoint(flow, role_val)
                defaults[flow_name][role_name] = dev.GetId()
            except Exception:
                defaults[flow_name][role_name] = None # Corrected: Assign to role_name, not role_val
    return defaults
def _safe_friendly_name_from_device(dev):
    """
    Read PKEY_Device_FriendlyName from an IMMDevice via IPropertyStore using raw vtable.
    Robustly handles PROPVARIANT layouts (pv.pwszVal, pv.value.pwszVal, pv.data.pwszVal),
    and when comtypes returns a Python str instead of a pointer. Falls back to
    PKEY_Device_DeviceDesc. Returns str or None on error.
    """
    try:
        import ctypes
        from ctypes import POINTER, byref, wintypes
        import comtypes.automation as automation
        # Calling convention (stdcall on 32-bit, cdecl on 64-bit)
        CALL = ctypes.WINFUNCTYPE if ctypes.sizeof(ctypes.c_void_p) == 4 else ctypes.CFUNCTYPE
        # PROPVARIANT type (prefer comtypes' implementation)
        if hasattr(automation, "PROPVARIANT"):
            PROPVARIANT = automation.PROPVARIANT
        elif hasattr(automation, "tagPROPVARIANT"):
            PROPVARIANT = automation.tagPROPVARIANT
        else:
            class _PVU(ctypes.Union):
                _fields_ = [
                    ("pwszVal", ctypes.c_wchar_p),
                    ("boolVal", ctypes.c_short),
                ]
            class PROPVARIANT(ctypes.Structure):
                _anonymous_ = ("data",)
                _fields_ = [
                    ("vt", ctypes.c_ushort),
                    ("wReserved1", ctypes.c_ushort),
                    ("wReserved2", ctypes.c_ushort),
                    ("wReserved3", ctypes.c_ushort),
                    ("data", _PVU),
                ]
        VT_LPWSTR = getattr(automation, "VT_LPWSTR", 31)
        class PROPERTYKEY(ctypes.Structure):
            _fields_ = (("fmtid", GUID), ("pid", wintypes.DWORD))
        # Raw IPropertyStore interface
        class IPropertyStoreRaw(ctypes.Structure):
            pass
        PIPS = POINTER(IPropertyStoreRaw)
        try:
            HRESULT_T = wintypes.HRESULT
        except Exception:
            HRESULT_T = ctypes.c_long
        # Only define the GetValue prototype to keep vtable minimal
        GetValueProto = CALL(HRESULT_T, PIPS, POINTER(PROPERTYKEY), POINTER(PROPVARIANT))
        class IPropertyStoreVTBL(ctypes.Structure):
            _fields_ = [
                ("QueryInterface", ctypes.c_void_p),  # 0
                ("AddRef", ctypes.c_void_p),          # 1
                ("Release", ctypes.c_void_p),         # 2
                ("GetCount", ctypes.c_void_p),        # 3
                ("GetAt", ctypes.c_void_p),           # 4
                ("GetValue", GetValueProto),          # 5
                ("SetValue", ctypes.c_void_p),        # 6
                ("Commit", ctypes.c_void_p),          # 7
            ]
        IPropertyStoreRaw._fields_ = [("lpVtbl", POINTER(IPropertyStoreVTBL))]
        # Open property store
        ps_unknown = dev.OpenPropertyStore(STGM_READ)
        ps_ptr_val = ctypes.cast(ps_unknown, ctypes.c_void_p).value
        if not ps_ptr_val:
            return None
        ps_iface = ctypes.cast(ctypes.c_void_p(ps_ptr_val), PIPS)
        # Keys
        PKEY_Device_FriendlyName = PROPERTYKEY(GUID("{a45c254e-df1c-4efd-8020-67d146a850e0}"), 14)
        PKEY_Device_DeviceDesc   = PROPERTYKEY(GUID("{a45c254e-df1c-4efd-8020-67d146a850e0}"), 2)
        # Clear helper
        ole32 = ctypes.OleDLL("ole32.dll")
        PropVariantClear = ole32.PropVariantClear
        PropVariantClear.restype = HRESULT_T
        PropVariantClear.argtypes = (POINTER(PROPVARIANT),)
        def _read_ptr_or_str(val):
            # Accept either a wide-char pointer or a Python str
            if isinstance(val, str):
                return val
            if val:
                try:
                    return ctypes.wstring_at(val)
                except Exception:
                    return None
            return None
        def _pv_read_lpwstr(pv):
            # Try common layouts in order
            try:
                s = _read_ptr_or_str(getattr(pv, "pwszVal", None))
                if s:
                    return s
            except Exception:
                pass
            try:
                val = getattr(pv, "value", None)
                if val is not None:
                    s = _read_ptr_or_str(getattr(val, "pwszVal", None))
                    if s:
                        return s
            except Exception:
                pass
            try:
                data = getattr(pv, "data", None)
                if data is not None:
                    s = _read_ptr_or_str(getattr(data, "pwszVal", None))
                    if s:
                        return s
            except Exception:
                pass
            return None
        def _get_string_prop(pkey):
            pv = PROPVARIANT()
            try:
                hr = ps_iface.contents.lpVtbl.contents.GetValue(ps_iface, byref(pkey), byref(pv))
                if hr == 0 and pv.vt == VT_LPWSTR:
                    s = _pv_read_lpwstr(pv)
                    if s:
                        return s.strip("\x00 ").strip()
            finally:
                try:
                    PropVariantClear(byref(pv))
                except Exception:
                    pass
            return None
        # Try FriendlyName, then DeviceDesc
        name = _get_string_prop(PKEY_Device_FriendlyName)
        if not name:
            name = _get_string_prop(PKEY_Device_DeviceDesc)
        if name:
            return name
    except Exception:
        pass
    # Fallback: return device ID if no name found
    try:
        return dev.GetId()
    except Exception:
        return None

def list_devices(include_all=False):
    """
    Returns list of devices with fields: id, name, flow, state, isDefault flags.
    """
    with warnings.catch_warnings():
        warnings.simplefilter("ignore", UserWarning)
        # Get defaults once (valid for both flows)
        enumerator_for_defaults = CoCreateInstance(CLSID_MMDeviceEnumerator, IMMDeviceEnumerator, CLSCTX_ALL)
        defaults = get_default_ids(enumerator_for_defaults)
        state_mask = DEVICE_STATE_ALL if include_all else DEVICE_STATE_ACTIVE
        out = []
        for flow_name, flow in [("Render", E_RENDER), ("Capture", E_CAPTURE)]:
            enumerator, coll = enum_endpoints(flow, state_mask)
            for i in range(coll.GetCount()):
                dev = coll.Item(i)
                dev_id = dev.GetId()
                name = _safe_friendly_name_from_device(dev) or dev_id
                state_str = "active"
                if include_all:
                    try:
                        st = dev.GetState()
                        if st != DEVICE_STATE_ACTIVE:
                            parts = [label for bit, label in DEVICE_STATES.items() if st & bit]
                            state_str = ",".join(parts) if parts else "unknown"
                    except Exception:
                        state_str = "unknown"
                is_default = {
                    "console": dev_id == defaults[flow_name]["console"],
                    "multimedia": dev_id == defaults[flow_name]["multimedia"],
                    "communications": dev_id == defaults[flow_name]["communications"],
                }
                out.append({
                    "id": dev_id,
                    "name": name,
                    "flow": flow_name,
                    "state": state_str,
                    "isDefault": is_default,
                })
        return out

def find_devices_by_selector(devices, dev_id=None, name_substr=None, flow=None, regex=False):
    """
    Returns list of devices matching selector.
    flow: "Render", "Capture" or None
    """
    if not dev_id and not name_substr:
        return []
    def match(d):
        if flow and d["flow"].lower() != flow.lower():
            return False
        if dev_id:
            return d["id"] == dev_id
        if name_substr:
            if regex:
                return re.search(name_substr, d["name"], re.IGNORECASE) is not None
            return name_substr.lower() in d["name"].lower()
        return False
    return [d for d in devices if match(d)]

def _sort_and_tag_gui_indices(devices):
    """
    Sort devices by name within each flow exactly like the GUI, and tag each
    item with d['guiIndex'] (0-based within its flow).
    Returns {'Render': [list], 'Capture': [list]} and also mutates 'devices'
    to include 'guiIndex' for convenience.
    """
    buckets = {"Render": [], "Capture": []}
    for d in devices:
        buckets[d["flow"]].append(d)
    for flow in buckets:
        buckets[flow].sort(key=lambda x: x["name"].lower())
        for i, d in enumerate(buckets[flow]):
            d["guiIndex"] = i
    return buckets

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

def _is_device_active(device_id):
    for flow in (E_RENDER, E_CAPTURE):
        try:
            _, coll = enum_endpoints(flow, DEVICE_STATE_ACTIVE)
            for i in range(coll.GetCount()):
                if coll.Item(i).GetId() == device_id:
                    return True
        except Exception:
            pass
    return False
def _pretty_matches_msg(label, matches):
    # Print a small list of candidates in GUI order to help the user pick
    buckets = _sort_and_tag_gui_indices(matches[:])  # shallow copy for safety
    lines = []
    for flow in ("Render", "Capture"):
        for d in buckets[flow]:
            flags = [k for k, v in d["isDefault"].items() if v]
            lines.append(f"  [{flow} idx {d.get('guiIndex','?')}] {d['name']}  id={d['id']}  defaults={','.join(flags) if flags else '-'}")
    return f"Multiple {label} matches:\n" + "\n".join(lines)
def _select_by_name_active_only(flow_name, name_text, index, regex):
    """
    Interpret --index using the same GUI order (sorted by name within flow).
    Only considers active devices. Returns (device_dict, None) or (None, error).
    """
    label = "playback" if flow_name == "Render" else "recording"
    active_devices = list_devices(include_all=False)
    # same flow only
    candidates = [d for d in active_devices if d["flow"] == flow_name]
    # GUI order and tag indices
    buckets = _sort_and_tag_gui_indices(candidates)
    ordered = buckets[flow_name]
    # name filter
    if name_text:
        if regex:
            pat = re.compile(name_text, re.IGNORECASE)
            ordered = [d for d in ordered if pat.search(d["name"])]
        else:
            ordered = [d for d in ordered if name_text.lower() in d["name"].lower()]
    if not ordered:
        return None, f"ERROR: {label} device not found (active only)"
    if index is None and len(ordered) > 1:
        return None, _pretty_matches_msg(label, ordered) + "\nUse --index to disambiguate."
    # If --index provided, use GUI index space; pick the item whose guiIndex == index
    if index is not None:
        for d in ordered:
            if d.get("guiIndex") == index:
                return d, None
        return None, f"ERROR: --index {index} does not match any active {label} device in GUI order."
    return ordered[0], None

def _get_policy_config():
    """
    Obtain a PolicyConfig COM interface that supports SetDefaultEndpoint.
    Tries AudioUtilities, then pycaw.policyconfig, then a built-in local COM definition (Vista interface).
    """
    # 1) Try any helper exposed by the installed pycaw AudioUtilities
    for name in ("GetPolicyConfig", "_get_policy_config", "get_policy_config"):
        try:
            getter = getattr(AudioUtilities, name, None)
            if getter:
                return getter()
        except Exception:
            pass
    # 2) Fallback: try to import pycaw.policyconfig (if present in your pycaw)
    try:
        from pycaw.policyconfig import IPolicyConfig, IPolicyConfigVista, CLSID_PolicyConfigClient
        try:
            return CoCreateInstance(CLSID_PolicyConfigClient, interface=IPolicyConfig, clsctx=CLSCTX_ALL)
        except Exception:
            return CoCreateInstance(CLSID_PolicyConfigClient, interface=IPolicyConfigVista, clsctx=CLSCTX_ALL)
    except Exception:
        pass  # Not available in your environment
    # 3) Final fallback: define the Vista interface locally and use it directly
    try:
        import ctypes
        from ctypes import wintypes
        # CLSID for PolicyConfigClient
        CLSID_PolicyConfigClient = GUID("{294935CE-F637-4E7C-A41B-AB255460B862}")
        # Minimal IPolicyConfigVista definition with correct order up to SetDefaultEndpoint.
        # We use c_void_p for parameters we won't call to avoid pulling extra types.
        class IPolicyConfigVista(IUnknown):
            _iid_ = GUID("{568B9108-44BF-40B4-9006-86AFE5B5A620}")
            _methods_ = (
                COMMETHOD([], HRESULT, 'GetMixFormat',
                          (['in'], wintypes.LPCWSTR, 'wszDeviceId'),
                          (['out'], ctypes.POINTER(ctypes.c_void_p), 'ppFormat')),
                COMMETHOD([], HRESULT, 'GetDeviceFormat',
                          (['in'], wintypes.LPCWSTR, 'wszDeviceId'),
                          (['in'], wintypes.BOOL, 'bDefault'),
                          (['out'], ctypes.POINTER(ctypes.c_void_p), 'ppFormat')),
                COMMETHOD([], HRESULT, 'SetDeviceFormat',
                          (['in'], wintypes.LPCWSTR, 'wszDeviceId'),
                          (['in'], ctypes.c_void_p, 'pEndpointFormat'),
                          (['in'], ctypes.c_void_p, 'mixFormat')),
                COMMETHOD([], HRESULT, 'GetProcessingPeriod',
                          (['in'], wintypes.LPCWSTR, 'wszDeviceId'),
                          (['in'], wintypes.BOOL, 'bDefault'),
                          (['out'], ctypes.POINTER(ctypes.c_longlong), 'pmftDefaultPeriod'),
                          (['out'], ctypes.POINTER(ctypes.c_longlong), 'pmftMinimumPeriod')),
                COMMETHOD([], HRESULT, 'SetProcessingPeriod',
                          (['in'], wintypes.LPCWSTR, 'wszDeviceId'),
                          (['in'], ctypes.POINTER(ctypes.c_longlong), 'pmftPeriod')),
                COMMETHOD([], HRESULT, 'GetShareMode',
                          (['in'], wintypes.LPCWSTR, 'wszDeviceId'),
                          (['out'], ctypes.POINTER(ctypes.c_void_p), 'pMode')),
                COMMETHOD([], HRESULT, 'SetShareMode',
                          (['in'], wintypes.LPCWSTR, 'wszDeviceId'),
                          (['in'], ctypes.c_void_p, 'mode')),
                COMMETHOD([], HRESULT, 'GetPropertyValue',
                          (['in'], wintypes.LPCWSTR, 'wszDeviceId'),
                          (['in'], ctypes.POINTER(ctypes.c_void_p), 'key'),
                          (['out'], ctypes.POINTER(ctypes.c_void_p), 'pv')),
                COMMETHOD([], HRESULT, 'SetPropertyValue',
                          (['in'], wintypes.LPCWSTR, 'wszDeviceId'),
                          (['in'], ctypes.POINTER(ctypes.c_void_p), 'key'),
                          (['in'], ctypes.POINTER(ctypes.c_void_p), 'pv')),
                # The one we actually need:
                COMMETHOD([], HRESULT, 'SetDefaultEndpoint',
                          (['in'], wintypes.LPCWSTR, 'wszDeviceId'),
                          (['in'], wintypes.DWORD, 'role')),
                COMMETHOD([], HRESULT, 'SetEndpointVisibility',
                          (['in'], wintypes.LPCWSTR, 'wszDeviceId'),
                          (['in'], wintypes.BOOL, 'bVisible')),
            )
        return CoCreateInstance(CLSID_PolicyConfigClient, interface=IPolicyConfigVista, clsctx=CLSCTX_ALL)
    except Exception as e:
        raise AttributeError("Audio policy config interface not available in this environment") from e

def set_default_endpoint(device_id, role):
    """
    role in ROLES keys or 'all'. Requires the device to be active.
    """
    if not _is_device_active(device_id):
        raise RuntimeError("Target device is not active; refusing to set default.")
    policy = _get_policy_config()
    def _call(rname, rval):
        try:
            policy.SetDefaultEndpoint(device_id, rval)
            return True, None
        except Exception as e:
            return False, e
    if role == "all":
        results = {}
        ok_all = True
        last_err = None
        for rname, rval in (("console", E_CONSOLE), ("multimedia", E_MULTIMEDIA), ("communications", E_COMMUNICATIONS)):
            ok, err = _call(rname, rval)
            results[rname] = ok
            if not ok:
                ok_all = False
                last_err = err
        if not ok_all:
            details = ", ".join([f"{k}={'ok' if v else 'fail'}" for k, v in results.items()])
            raise RuntimeError(f"SetDefaultEndpoint failed for roles: {details}. Underlying error: {last_err}")
    else:
        policy.SetDefaultEndpoint(device_id, ROLES[role])

def set_endpoint_mute(device_id, mute_state):
    for flow in (E_RENDER, E_CAPTURE):
        _, coll = enum_endpoints(flow, DEVICE_STATE_ACTIVE)  # ACTIVE ONLY
        for i in range(coll.GetCount()):
            dev = coll.Item(i)
            if dev.GetId() == device_id:
                try:
                    vol_iface = dev.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
                    vol = ctypes.cast(vol_iface, ctypes.POINTER(IAudioEndpointVolume))
                    vol.SetMute(mute_state, None)
                    return True
                except Exception:
                    return False
    return False

def get_endpoint_mute(device_id):
    for flow in (E_RENDER, E_CAPTURE):
        try:
            _, coll = enum_endpoints(flow, DEVICE_STATE_ACTIVE)  # ACTIVE ONLY
            for i in range(coll.GetCount()):
                dev = coll.Item(i)
                if dev.GetId() == device_id:
                    vol_iface = dev.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
                    vol = ctypes.cast(vol_iface, ctypes.POINTER(IAudioEndpointVolume))
                    try:
                        ret = vol.GetMute()
                        if isinstance(ret, tuple):
                            ret = ret[0]
                        return bool(ret)
                    except Exception:
                        try:
                            from ctypes import wintypes
                            b = wintypes.BOOL()
                            vol.GetMute(ctypes.byref(b))
                            return bool(b.value)
                        except Exception:
                            return None
        except Exception:
            continue
    return None

def get_endpoint_volume(device_id):
    for flow in (E_RENDER, E_CAPTURE):
        try:
            _, coll = enum_endpoints(flow, DEVICE_STATE_ACTIVE)  # ACTIVE ONLY
            for i in range(coll.GetCount()):
                dev = coll.Item(i)
                if dev.GetId() == device_id:
                    vol_iface = dev.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
                    vol = ctypes.cast(vol_iface, ctypes.POINTER(IAudioEndpointVolume))
                    try:
                        ret = vol.GetMasterVolumeLevelScalar()
                        if isinstance(ret, tuple):
                            ret = ret[0]
                        return max(0, min(100, int(round(float(ret) * 100.0))))
                    except Exception:
                        try:
                            f = ctypes.c_float()
                            vol.GetMasterVolumeLevelScalar(ctypes.byref(f))
                            return max(0, min(100, int(round(float(f.value) * 100.0))))
                        except Exception:
                            return None
        except Exception:
            continue
    return None

def set_endpoint_volume(device_id, level_percent):
    level = max(0.0, min(1.0, float(level_percent) / 100.0))
    for flow in (E_RENDER, E_CAPTURE):
        _, coll = enum_endpoints(flow, DEVICE_STATE_ACTIVE)  # ACTIVE ONLY
        for i in range(coll.GetCount()):
            dev = coll.Item(i)
            if dev.GetId() == device_id:
                try:
                    vol_iface = dev.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
                    vol = ctypes.cast(vol_iface, ctypes.POINTER(IAudioEndpointVolume))
                    vol.SetMasterVolumeLevelScalar(level, None)
                    return True
                except Exception:
                    return False
    return False

def cmd_set_default(args):
    if not is_admin():
        print("WARNING: 'set-default' might require Administrator privileges on this system.", file=sys.stderr)
    exit_code = 0
    results = {"set": []}
    # Playback (Render)
    if args.playback_id or args.playback_name:
        flow_name = "Render"
        if args.playback_id:
            # active only
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
    # Recording (Capture)
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
    devices = list_devices(include_all=False)  # ACTIVE ONLY
    matches = find_devices_by_selector(devices, dev_id=args.id, name_substr=args.name, flow=args.flow, regex=args.regex)
    if len(matches) == 0:
        print("ERROR: device not found (active only)", file=sys.stderr)
        return 3
    if len(matches) > 1 and args.index is None:
        print("ERROR: multiple matches; specify --index", file=sys.stderr)
        return 4
    # Sort like GUI to interpret --index consistently
    buckets = _sort_and_tag_gui_indices([d for d in matches])
    flow = args.flow or (matches[0]["flow"] if matches else None) # Handle empty matches
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
        if ok: print(json.dumps({"muteSet": {"id": target["id"], "name": target["name"], "muted": True}}))
    elif args.unmute:
        ok = set_endpoint_mute(target["id"], False)
        if ok: print(json.dumps({"muteSet": {"id": target["id"], "name": target["name"], "muted": False}}))
    elif args.level is not None:
        ok = set_endpoint_volume(target["id"], args.level)
        if ok: print(json.dumps({"volumeSet": {"id": target["id"], "name": target["name"], "level": args.level}}))
    if not ok:
        print("ERROR: failed to set volume/mute", file=sys.stderr)
        return 1
    return 0
    
def cmd_listen(args):
    devices = list_devices(include_all=False)  # ACTIVE ONLY
    matches = find_devices_by_selector(devices, dev_id=args.id, name_substr=args.name, flow="Capture", regex=args.regex)
    if len(matches) == 0:
        print("ERROR: capture device not found (active only)", file=sys.stderr)
        return 3
    if len(matches) > 1 and args.index is None:
        print("ERROR: multiple matches; specify --index", file=sys.stderr)
        return 4
    # Sort like GUI to interpret --index consistently
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
    captured_stderr = io.StringIO()
    ok = False
    with redirect_stderr(captured_stderr):
        # FIX: pass args.enable (a boolean), not the device dict
        ok = set_listen_to_device_ps(target["id"], args.enable, render_device_id=args.playback_target_id)
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

def cmd_wait(args):
    """
    Wait for an ACTIVE device to appear (by name substring or regex, optional flow) up to timeout seconds.
    """
    deadline = time.time() + args.timeout
    while time.time() < deadline:
        devices = list_devices(include_all=False)  # ACTIVE ONLY
        matches = find_devices_by_selector(devices, dev_id=args.id, name_substr=args.name, flow=args.flow, regex=args.regex)
        if matches:
            # Report in GUI order for predictability
            buckets = _sort_and_tag_gui_indices(matches[:])
            flow = args.flow or (matches[0]["flow"] if matches else None) # Handle empty matches
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
    # list
    p_list = sub.add_parser("list", help="List devices")
    p_list.add_argument("--all", action="store_true", help="Include disabled/disconnected")
    p_list.add_argument("--json", action="store_true") 
    p_list.set_defaults(func=cmd_list)
    # set-default
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
    # set-volume
    p_sv = sub.add_parser("set-volume", help="Set endpoint volume (render or capture) or mute/unmute")
    p_sv.add_argument("--id")
    p_sv.add_argument("--name")
    p_sv.add_argument("--flow", choices=["Render", "Capture"], help="Optional filter to disambiguate")
    p_sv.add_argument("--level", type=int, help="0-100 (for volume)") # Changed to optional
    p_sv.add_argument("--mute", action="store_true", help="Mute the device") # mute
    p_sv.add_argument("--unmute", action="store_true", help="Unmute the device") # unmute
    p_sv.add_argument("--index", type=int)
    p_sv.add_argument("--regex", action="store_true") 
    p_sv.set_defaults(func=cmd_set_volume)
    # listen
    p_ls = sub.add_parser("listen", help="Enable/disable 'Listen to this device' (capture only)")
    p_ls.add_argument("--id", help="Device ID for the capture device.")
    p_ls.add_argument("--name", help="Substring of the device name for the capture device.")
    p_ls.add_argument("--enable", action="store_true", help="Enable 'Listen to this device'.") 
    p_ls.add_argument("--disable", action="store_true", help="Disable 'Listen to this device'.") 
    p_ls.add_argument("--playback-target-id", help="Optional: Render endpoint ID to play through. Use '' for 'Default Playback Device'. If omitted, current target is preserved (or set to '' if enabling and no target).")
    p_ls.add_argument("--index", type=int)
    p_ls.add_argument("--regex", action="store_true") 
    p_ls.set_defaults(func=cmd_listen)
    # wait
    p_w = sub.add_parser("wait", help="Wait for device to appear")
    p_w.add_argument("--id")
    p_w.add_argument("--name")
    p_w.add_argument("--flow", choices=["Render", "Capture"])
    p_w.add_argument("--timeout", type=int, default=30)
    p_w.add_argument("--index", type=int)
    p_w.add_argument("--regex", action="store_true") 
    p_w.set_defaults(func=cmd_wait)
    return p
# =========================
# GUI WRAPPER (runs when script is double-clicked with no args)
# =========================
# --- Robust global exception logging + faulthandler (replace your current logging block with this) ---
import os, sys, traceback, datetime, tempfile
try:
    import faulthandler
except Exception:
    faulthandler = None
# Prefer the directory of the EXE (PyInstaller) or the script file
def _exe_dir():
    try:
        if getattr(sys, "frozen", False):  # PyInstaller/py2exe
            return os.path.dirname(sys.executable)
        return os.path.dirname(os.path.abspath(__file__))
    except Exception:
        return os.getcwd()
def _init_log_paths():
    base = _exe_dir()
    path = os.path.join(base, "audioctl_gui.log")
    # Try to ensure we can write here; if not, fallback to temp
    try:
        os.makedirs(base, exist_ok=True)
        test = os.path.join(base, ".writetest")
        with open(test, "w", encoding="utf-8") as _:
            pass
        os.remove(test)
        return base, path
    except Exception:
        # Fallback: temp (last resort)
        tdir = os.path.join(tempfile.gettempdir(), "audioctl")
        try:
            os.makedirs(tdir, exist_ok=True)
        except Exception:
            tdir = tempfile.gettempdir()
        return tdir, os.path.join(tdir, "audioctl_gui.log")
_LOG_DIR, _LOG_PATH = _init_log_paths()
# Record where we’re logging
try:
    with open(_LOG_PATH, "a", encoding="utf-8", errors="replace") as _:
        pass
    # First breadcrumb
    ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(_LOG_PATH, "a", encoding="utf-8", errors="replace") as f:
        f.write(f"[{ts}] logging to: {_LOG_PATH}\n")
except Exception:
    pass
def _log_path():
    return _LOG_PATH
def _log(msg):
    try:
        ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with open(_LOG_PATH, "a", encoding="utf-8", errors="replace") as f:
            f.write(f"[{ts}] {msg}\n")
    except Exception:
        pass
def _log_exc(prefix, exc_info=None):
    try:
        if exc_info is None:
            exc_info = sys.exc_info()
        tb = "".join(traceback.format_exception(*exc_info))
        _log(f"{prefix}\n{tb}")
    except Exception:
        pass
def _global_excepthook(exc_type, exc_value, exc_tb):
    _log_exc("UNCAUGHT EXCEPTION", (exc_type, exc_value, exc_tb))
    try:
        sys.__excepthook__(exc_type, exc_value, exc_tb)
    except Exception:
        pass
sys.excepthook = _global_excepthook
# Unraisable exceptions (Python 3.8+)
try:
    def _unraisable_hook(unraisable):
        _log(f"UNRAISABLE: {getattr(unraisable.exc_type, '__name__', str(unraisable.exc_type))}: "
             f"{unraisable.exc_value}\nObject: {unraisable.object!r}")
    sys.unraisablehook = _unraisable_hook
except Exception:
    pass
# Define _fh before the try so it exists for atexit close
_fh = None
try:
    if faulthandler:
        _fh = open(_LOG_PATH, "a", buffering=1)
        faulthandler.enable(file=_fh, all_threads=True)
except Exception:
    try:
        if faulthandler:
            faulthandler.enable(all_threads=True)
    except Exception:
        pass
# Log normal exits and sys.exit calls
import atexit
@atexit.register
def _on_atexit():
    _log("atexit: process exiting normally")
# Close the faulthandler file on exit (flushes crash dumps)
@atexit.register
def _close_log_handles():
    try:
        if faulthandler and _fh and not _fh.closed:
            _fh.flush()
            _fh.close()
    except Exception:
        pass
_orig_sys_exit = sys.exit
def _logged_sys_exit(code=0):
    try:
        _log(f"sys.exit invoked with code={code!r}")
    except Exception:
        pass
    raise SystemExit(code)
sys.exit = _logged_sys_exit
# Best-effort logging of console/logoff/shutdown control events (Windows)
try:
    import ctypes
    from ctypes import wintypes
    CTRL_C_EVENT = 0
    CTRL_BREAK_EVENT = 1
    CTRL_CLOSE_EVENT = 2
    CTRL_LOGOFF_EVENT = 5
    CTRL_SHUTDOWN_EVENT = 6
    _ConsoleCtrlHandler = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.DWORD)
    def _console_ctrl_handler(ctrl_type):
        try:
            _log(f"ConsoleCtrl event: {ctrl_type} "
                 f"({'CTRL_C' if ctrl_type==CTRL_C_EVENT else 'CTRL_BREAK' if ctrl_type==CTRL_BREAK_EVENT else 'CTRL_CLOSE' if ctrl_type==CTRL_CLOSE_EVENT else 'CTRL_LOGOFF' if ctrl_type==CTRL_LOGOFF_EVENT else 'CTRL_SHUTDOWN' if ctrl_type==CTRL_SHUTDOWN_EVENT else 'UNKNOWN'})")
        except Exception:
            pass
        return False  # allow default handling
    try:
        ctypes.windll.kernel32.SetConsoleCtrlHandler(_ConsoleCtrlHandler(_console_ctrl_handler), True)
    except Exception:
        pass
except Exception:
    pass
# --- End robust global exception logging + faulthandler ---
def launch_gui():
    import tkinter as tk
    from tkinter import ttk, messagebox
    try:
        CoInitialize()
    except Exception:
        pass
    
    _log("launch_gui: creating Tk root")
    root = tk.Tk()

    try:
        if sys.platform.startswith("win"):
            root.iconbitmap(resource_path("audio.ico"))
    except Exception:
        pass

    def _on_root_close():
        try:
            _log("WM_DELETE_WINDOW received: root close requested (user/system)")
        except Exception:
            pass
        root.destroy()
    try:
        root.protocol("WM_DELETE_WINDOW", _on_root_close)
    except Exception:
        pass
    # Log destroy of the root (helps see who triggered it)
    def _on_any_destroy(ev):
        try:
            if ev.widget == root:
                _log("Tk <Destroy> on root window")
        except Exception:
            pass
    try:
        root.bind("<Destroy>", _on_any_destroy, add="+")
    except Exception:
        pass
    # Route any Tk callback exception to our logger, and show a friendly dialog
    def _tk_report_callback_exception(exc, val, tb):
        try:
            _log_exc("TK CALLBACK EXCEPTION", (exc, val, tb))
        except Exception:
            pass
        try:
            messagebox.showerror("Unexpected error", f"{exc.__name__}: {val}\n\nDetails were written to:\n{_LOG_PATH}")
        except Exception: # Use _LOG_PATH here as well
            pass
    try:
        root.report_callback_exception = _tk_report_callback_exception
    except Exception:
        pass
    
    # tkfont is imported locally where needed, as per instruction to prevent issues
    class AudioGUI:
        def __init__(self, root):
            self.root = root
            self.root.title("Audio Control v1.2.1.0 10-27-2025")
            try:
                pass
            except Exception:
                pass
            # Style: try to make Treeview column headers bold (guarded)
            style = ttk.Style(self.root)
            try:
                for theme in ("vista", "xpnative", "clam", "default"):
                    if theme in style.theme_names():
                        style.theme_use(theme)
                        break
            except Exception:
                pass
            try: # Start of the new try block for font styling
                from tkinter import font as tkfont # Moved import here
                base_font = tkfont.nametofont(self.root.cget("font"))
                try:
                    heading_font = base_font.copy()
                    heading_font.configure(weight="bold")
                except Exception:
                    heading_font = base_font
                # This line may fail on some Tk builds; wrap it
                style.configure("Treeview.Heading", font=heading_font)
            except Exception:
                # Fallback: leave default heading font
                pass
            
            # Make headers non-interactive: no hover/press highlight and no click behavior
            try:
                style.configure("Treeview.Heading", relief="flat")
                style.map("Treeview.Heading", background=[], relief=[], foreground=[])
            except Exception:
                pass
            # Variables
            self.include_all = tk.BooleanVar(value=False)
            self.print_cmd = tk.BooleanVar(value=False)  # When checked, print exact CLI command for actions
            self.devices = []
            self.item_to_device = {}
            # Layout
            self.container = ttk.Frame(self.root, padding=10)
            self.container.pack(fill="both", expand=True)
            self.topbar = ttk.Frame(self.container)
            self.topbar.pack(fill="x", pady=(0, 8))
            refresh_btn = ttk.Button(self.topbar, text="Refresh", command=self.refresh_devices)
            refresh_btn.pack(side="left")
            ttk.Checkbutton(
                self.topbar,
                text="Show disabled/disconnected",
                variable=self.include_all,
                command=self.refresh_devices
            ).pack(side="left", padx=(10, 0))
            ttk.Checkbutton(
                self.topbar,
                text="Print CLI commands",
                variable=self.print_cmd
            ).pack(side="left", padx=(10, 0))
            if not is_admin():
                admin_lbl = ttk.Label(self.topbar, text="Note: Some actions may require Administrator", foreground="#CC6600")
                admin_lbl.pack(side="right")
            # Treeview with a visible tree column (#0) to get a caret on group rows
            columns = ("Index", "Name", "Flow", "Defaults", "ID")
            self.tree = ttk.Treeview(
                self.container,
                columns=columns,
                show="tree headings",   # keep showing the tree column (#0) but we removed the indicators above
                selectmode="browse",
                height=10
            )
            self.tree.heading("#0", text="")   # no label; purely for group nodes
            for col in columns:
                self.tree.heading(col, text=col)
            # Make the tree column right-aligned and a bit slimmer so its text is closer to the Index column
            self.tree.column("#0", width=120, minwidth=100, anchor="e", stretch=False)  # group column (right-aligned now)
            self.tree.column("Index", width=60, minwidth=50, anchor="e", stretch=False)
            self.tree.column("Name",  width=200, minwidth=200, anchor="w", stretch=True)
            self.tree.column("Flow",  width=80,  minwidth=80,  anchor="w", stretch=False)
            self.tree.column("Defaults", width=160, minwidth=160, anchor="w", stretch=False)
            self.tree.column("ID",    width=260, minwidth=240, anchor="w", stretch=False)
            self.tree["displaycolumns"] = ("Index", "Name", "Flow", "Defaults", "ID")
            self.tree.pack(fill="both", expand=True)
            # Scrollbar
            self.yscroll = ttk.Scrollbar(self.tree, orient="vertical", command=self.tree.yview)
            self.tree.configure(yscrollcommand=self.yscroll.set)
            self.yscroll.pack(side="right", fill="y")
            
            # Tag to style group rows (Playback/Recording) – use only safe options
            try:
                self.tree.tag_configure("group", foreground="#202020")  # no 'font' here
            except Exception:
                pass
            # Treeview.Item layout without the indicator element, removing carets/arrows
            try:
                style.layout("Treeview.Item", [
                    ('Treeitem.padding', {
                        'sticky': 'nswe',
                        'children': [
                            ('Treeitem.image', {'side': 'left', 'sticky': ''}),
                            ('Treeitem.focus', {'side': 'left', 'sticky': 'nswe', 'children': [
                                ('Treeitem.text', {'side': 'left', 'sticky': ''})
                            ]})
                        ]
                    })
                ])
            except Exception:
                pass
            # Status bar
            self.status = tk.StringVar(value="Ready")
            self.statusbar = ttk.Label(self.root, textvariable=self.status, anchor="w", padding=(10, 3))
            self.statusbar.pack(fill="x", side="bottom")
            
            # Popup menu (store index for both Mute and Listen items so relabeling is easy)
            self.menu = tk.Menu(self.root, tearoff=0)
            self.menu.add_command(label="Set as Default (all roles)", command=self.on_set_default)
            self.menu.add_separator()
            self.menu.add_command(label="Set Volume...", command=self.on_set_volume)
            # Single Mute/Unmute item (label adjusted dynamically)
            self.menu.add_command(label="Mute", command=self.on_toggle_mute)
            self.mute_menu_index = self.menu.index("end")  # numeric index of the Mute/Unmute item
            self.menu.add_separator()
            # Toggle Listen item (capture only)
            self.listen_menu_default_label = "Toggle Listen (capture only)"
            self.menu.add_command(label=self.listen_menu_default_label, command=self.on_toggle_listen)
            self.listen_menu_index = self.menu.index("end")  # numeric index of the Listen item
            # Bindings
            self.tree.bind("<Button-3>", self.on_right_click)           # Right click
            self.tree.bind("<ButtonRelease-1>", self.on_left_release)   # Select on left click
            self.tree.bind("<Double-1>", self.on_double_click)          # Double-click to open menu
            self.root.bind("<F5>", lambda e: self.refresh_devices())    # F5 to refresh
            # bindings for group row behavior
            self.tree.bind("<Button-1>", self.on_left_click, add="+")      # Prevent selecting group headers
            self.tree.bind("<<TreeviewSelect>>", self.on_select_change)    # Guard against keyboard selection of headers
            # Initial load
            self.refresh_devices()
            # Ensure a final fit once everything is drawn
            self.root.after_idle(self.adjust_layout_to_content)
        def is_group_row(self, iid):
            # Group headers are not in the item_to_device mapping
            return iid not in self.item_to_device
            
        def on_left_click(self, event):
            # Ignore clicks in the header area and prevent selecting group headers
            region = self.tree.identify_region(event.x, event.y)
            if region == "heading":
                return "break"
            iid = self.tree.identify_row(event.y)
            if not iid:
                return
            if self.is_group_row(iid):
                return "break"
        def on_select_change(self, event):
            # Prevent group headers from staying selected via keyboard navigation
            sel = self.tree.selection()
            if not sel:
                return
            iid = sel[0]
            if self.is_group_row(iid):
                # If group has children, move selection to first child; otherwise clear selection
                children = self.tree.get_children(iid)
                if children:
                    self.tree.selection_set(children[0])
                    self.tree.focus(children[0])
                else:
                    self.tree.selection_remove(iid)
        def set_status(self, text):
            self.status.set(text)
            # Also echo to console if any (useful for debugging)
            try:
                print(text)
            except Exception:
                pass
        def refresh_devices(self):
            try:
                self.devices = list_devices(include_all=self.include_all.get())
                self.item_to_device.clear()
                for item in self.tree.get_children():
                    self.tree.delete(item)
                # Split and sort: Playback (Render) first, Recording (Capture) second
                render_devs = sorted([d for d in self.devices if d["flow"] == "Render"], key=lambda x: x["name"].lower())
                capture_devs = sorted([d for d in self.devices if d["flow"] == "Capture"], key=lambda x: x["name"].lower())
                
                # Insert group headers (non-device rows) into tree column (#0); enable caret/expand
                grp_render = self.tree.insert("", "end", text="Playback (Render)", values=("", "", "", "", ""), open=True, tags=("group",))
                grp_capture = self.tree.insert("", "end", text="Recording (Capture)", values=("", "", "", "", ""), open=True, tags=("group",))
                
                # Helper to insert child rows under a group with per-group index
                def insert_group(parent, devs, flow_name):
                    for idx, d in enumerate(devs):
                        flags = [k for k, v in d["isDefault"].items() if v] 
                        defaults_txt = ", ".join(flags) if flags else "-"
                        # Store a copy with its display index so we can show it and generate CLI commands
                        d_copy = dict(d)
                        d_copy["_index"] = idx             # per-group index
                        d_copy["_group"] = flow_name       # "Render" or "Capture"
                        # Device rows have empty tree-column text (the caret lives on group rows)
                        iid = self.tree.insert(parent, "end", text="", values=(idx, d["name"], d["flow"], defaults_txt, d["id"]))
                        self.item_to_device[iid] = d_copy
                
                insert_group(grp_render, render_devs, "Render")
                insert_group(grp_capture, capture_devs, "Capture")
                self.set_status("Device list updated")
                # Resize columns, tree height, and window to fit content
                self.adjust_layout_to_content()
                self.root.after_idle(self.adjust_layout_to_content)
            except Exception as e:
                messagebox.showerror("Error", f"Failed to list devices:\n{e}")
                self.set_status("Failed to refresh devices")
        def adjust_layout_to_content(self):
            self.root.update_idletasks()
            try:
                from tkinter import font as tkfont
                tv_font_name = self.tree.cget("font") or "TkDefaultFont"
                tv_font = tkfont.nametofont(tv_font_name)
            except Exception:
                tv_font = None
            names = [d["name"] for d in self.devices] or ["Name"]
            defaults_list = []
            for d in self.devices:
                flags = [k for k, v in d["isDefault"].items() if v]
                defaults_list.append(", ".join(flags) if flags else "-")
            if not defaults_list:
                defaults_list = ["-"]
            ids = [d["id"] for d in self.devices] or ["ID"]
            longest_name = max(names, key=len)
            longest_defaults = max(defaults_list, key=len)
            longest_id = max(ids, key=len)
            # Group column (#0) width based on group labels
            group_labels = ["Playback (Render)", "Recording (Capture)"]
            longest_group = max(group_labels, key=len)
            render_count = sum(1 for d in self.devices if d["flow"] == "Render")
            capture_count = sum(1 for d in self.devices if d["flow"] == "Capture")
            max_index_value = max(render_count - 1, capture_count - 1, 0)
            pad = 32
            def measure(text, fallback):
                try:
                    return tv_font.measure(text) if tv_font else fallback
                except Exception:
                    return fallback
            # Keep it relatively narrow (we just need a right-aligned label area)
            group_w = max(100, min(180, measure(longest_group, 140) + 12))
            name_w    = max(240, min(700, max(measure(longest_name, 300), measure("Name", 60)) + pad))
            flow_w    = max(80, max(measure("Recording", 90), measure("Flow", 60)) + 30)
            defaults_w= max(160, min(480, max(measure(longest_defaults, 240), measure("Defaults", 100)) + pad))
            id_w      = max(240, min(560, max(measure(longest_id, 340), measure("ID", 60)) + pad))
            index_digits = max(2, len(str(max_index_value)))
            index_w = max(60, measure("9" * index_digits, 30) + 24)
            try:
                self.tree.column("#0", width=int(group_w), minwidth=140, anchor="w", stretch=False)
                self.tree.column("Index", width=int(index_w), minwidth=50, anchor="e", stretch=False)
                self.tree.column("Name",  width=int(name_w),  minwidth=200, anchor="w", stretch=True)
                self.tree.column("Flow",  width=int(flow_w),  minwidth=80,  anchor="w", stretch=False)
                self.tree.column("Defaults", width=int(defaults_w), minwidth=160, anchor="w", stretch=False)
                self.tree.column("ID",    width=int(id_w),    minwidth=240, anchor="w", stretch=False)
            except Exception:
                pass
            # rows = group headers (2) + number of devices
            rows = len(self.devices) + 2 if self.devices else 2
            self.tree.configure(height=min(max(rows, 6), 22))
            self.root.update_idletasks()
            try:
                sb_w = max(self.yscroll.winfo_reqwidth(), 16) if self.yscroll else 16
            except Exception:
                sb_w = 16
            total_cols = int(group_w + index_w + name_w + flow_w + defaults_w + id_w + sb_w + 40)
            desired_w = max(total_cols, self.container.winfo_reqwidth() + 10, 600)
            desired_h = max(self.root.winfo_reqheight(), 300)
            scr_w = self.root.winfo_screenwidth()
            scr_h = self.root.winfo_screenheight()
            margin = 80
            w = min(desired_w, scr_w - margin)
            h = min(desired_h, scr_h - margin)
            self.root.geometry(f"{int(w)}x{int(h)}")
            self.root.minsize(int(min(w, scr_w - margin)), int(min(h, scr_h - margin)))
        def maybe_print_cli(self, cmd_str: str):
            # Note: This prints to stdout. To see it when double-clicking a --windowed build,
            # run the CLI or build a --console variant.
            if self.print_cmd.get():
                try:
                    print(cmd_str)
                except Exception:
                    pass
        def get_selected_device(self):
            sel = self.tree.selection()
            if not sel:
                return None
            return self.item_to_device.get(sel[0])  # returns None for group headers, which is fine
            
        def show_menu_for_item(self, event, iid=None):
            try:
                if iid is None:
                    iid = self.tree.identify_row(event.y)
                if not iid:
                    return
                d = self.item_to_device.get(iid)
                # Do not select group headers; for devices, ensure they are selected
                if d:
                    self.tree.selection_set(iid)
                else:
                    # Make sure group headers are not highlighted
                    self.tree.selection_remove(iid)
                end_idx = self.menu.index("end")
                end_idx = end_idx if end_idx is not None else -1
                if not d:
                    # Group header: disable applicable menu entries (skip separators)
                    for i in range(end_idx + 1):
                        etype = self.menu.type(i)
                        if etype in ("command", "cascade", "checkbutton", "radiobutton"):
                            self.menu.entryconfig(i, state="disabled")
                    self.menu.tk_popup(event.x_root, event.y_root)
                    return
                # Device row: enable menu entries (skip separators)
                for i in range(end_idx + 1):
                    etype = self.menu.type(i)
                    if etype in ("command", "cascade", "checkbutton", "radiobutton"):
                        self.menu.entryconfig(i, state="normal")
                # Configure the Mute/Unmute entry by numeric index
                try:
                    muted = get_endpoint_mute(d["id"])
                except Exception:
                    muted = None
                if muted is True:
                    mute_label = "Unmute"
                elif muted is False:
                    mute_label = "Mute"
                else:
                    # Safe default if status can't be read: show "Unmute" so user never loses the ability to unmute
                    mute_label = "Unmute"
                self.menu.entryconfig(self.mute_menu_index, label=mute_label, state="normal")
                # Configure the Listen entry by numeric index
                if d["flow"] == "Capture":
                    try:
                        current = _get_listen_to_device_status_ps(d["id"])
                    except Exception:
                        current = None
                    if current is True:
                        label = "Disable Listen"
                    elif current is False:
                        label = "Enable Listen"
                    else:
                        label = self.listen_menu_default_label
                    self.menu.entryconfig(self.listen_menu_index, label=label, state="normal")
                else:
                    self.menu.entryconfig(self.listen_menu_index, label=self.listen_menu_default_label, state="disabled")
                self.menu.tk_popup(event.x_root, event.y_root)
            finally:
                try:
                    self.menu.grab_release()
                except Exception:
                    pass
        # Event handlers
        def on_right_click(self, event):
            iid = self.tree.identify_row(event.y)
            if not iid:
                return
            if not self.is_group_row(iid):
                # Only select device rows on right-click
                self.tree.selection_set(iid)
            self.show_menu_for_item(event, iid=iid)
        def on_left_release(self, event):
            iid = self.tree.identify_row(event.y)
            if not iid:
                return
            if self.is_group_row(iid):
                # Ensure group header is not highlighted
                self.tree.selection_remove(iid)
                return
            # Device rows: keep default behavior
            self.tree.selection_set(iid)
        def on_double_click(self, event):
            iid = self.tree.identify_row(event.y)
            if not iid:
                return
            if self.is_group_row(iid):
                # Toggle expand/collapse without selecting the group header
                self.tree.item(iid, open=not bool(self.tree.item(iid, "open")))
                # Ensure group header is not highlighted
                self.tree.selection_remove(iid)
                return
            # Device row: open context menu
            self.show_menu_for_item(event, iid=iid)
        # Actions
        def on_set_default(self):
            d = self.get_selected_device()
            if not d:
                return
            try:
                # Build CLI string to print if requested
                if d["flow"] == "Render":
                    cmd = f'audioctl set-default --playback-id "{d["id"]}" --playback-role all'
                else:
                    cmd = f'audioctl set-default --recording-id "{d["id"]}" --recording-role all'
                
                # Perform action
                if not is_admin():
                    if not messagebox.askyesno(
                        "Administrator recommended",
                        "Setting default device may require Administrator privileges on some systems.\n\nContinue?"
                    ):
                        return
                set_default_endpoint(d["id"], "all")
                self.maybe_print_cli(cmd) # Print CLI command
                self.set_status(f"Set default ({d['flow']}) device: {d['name']} (all roles)")
                self.refresh_devices()
            except Exception as e:
                messagebox.showerror("Error", f"Failed to set default:\n{e}")
                self.set_status("Failed to set default")
        
        def on_set_volume(self):
            d = self.get_selected_device()
            if not d:
                return
            try:
                level = self.open_volume_dialog(d["id"], d["name"])
                if level is None:
                    return
                ok = set_endpoint_volume(d["id"], level)
                if ok:
                    cmd = f'audioctl set-volume --id "{d["id"]}" --flow {d["flow"]} --level {level}'
                    self.maybe_print_cli(cmd)
                    self.set_status(f"Volume set to {level}% for: {d['name']}")
                else:
                    messagebox.showerror("Error", "Failed to set volume/mute")
                    self.set_status("Failed to set volume")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to set volume:\n{e}")
                self.set_status("Failed to set volume")
        def on_toggle_mute(self):
            d = self.get_selected_device()
            if not d:
                return
            try:
                current = get_endpoint_mute(d["id"])
                # If status can't be read, prefer to unmute (safe default)
                target = False if current is None else (not bool(current))
                ok = set_endpoint_mute(d["id"], target)
                if ok:
                    cmd = f'audioctl set-volume --id "{d["id"]}" --flow {d["flow"]} --{"mute" if target else "unmute"}'
                    self.maybe_print_cli(cmd)
                    self.set_status(f'{"Muted" if target else "Unmuted"}: {d["name"]}')
                else:
                    messagebox.showerror("Error", "Failed to change mute state")
                    self.set_status("Failed to change mute state")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to toggle mute:\n{e}")
                self.set_status("Failed to toggle mute")

        def on_toggle_listen(self):
            d = self.get_selected_device()
            if not d:
                return
            if d["flow"] != "Capture":
                messagebox.showinfo("Not a capture device", "Listen can only be toggled for capture (recording) devices.")
                return
            try:
                _log(f"Listen toggle requested for {d['name']} ({d['id']})")
                current = _get_listen_to_device_status_ps(d["id"])
                if current is None:
                    current = _read_listen_enable_from_registry(d["id"])
                enable = not bool(current)
                # Print intended CLI command before applying (for reproducibility)
                cmd = f'audioctl listen --id "{d["id"]}" --{"enable" if enable else "disable"}'
                self.maybe_print_cli(cmd)
                captured_stderr = io.StringIO()
                with redirect_stderr(captured_stderr):
                    ok = set_listen_to_device_ps(d["id"], enable, render_device_id=None)
                if not ok:
                    actual = _get_listen_to_device_status_ps(d["id"])
                    if actual is None:
                        verified, reg_state = _verify_listen_via_registry(d["id"], enable, timeout=3.0, interval=0.20)
                        actual = reg_state if verified or reg_state is not None else None
                else:
                    actual = _get_listen_to_device_status_ps(d["id"])
                    if actual is None:
                        verified, reg_state = _verify_listen_via_registry(d["id"], enable, timeout=3.0, interval=0.20)
                        actual = reg_state if verified or reg_state is not None else None
                if actual is None:
                    _log(f"Listen toggle result unknown for {d['name']} ({d['id']}); requested={enable}")
                    messagebox.showwarning("Listen status unknown", "Could not verify final 'Listen' state. It may still have applied.")
                    self.set_status(f"Listen toggle requested for: {d['name']}")
                else:
                    _log(f"Listen toggle result for {d['name']} ({d['id']}): final={actual}")
                    state_txt = "enabled" if actual else "disabled"
                    self.set_status(f"Listen {state_txt} for: {d['name']}")
            except Exception as e:
                _log(f"Listen toggle exception for {d['name']} ({d['id']}): {e!r}")
                messagebox.showerror("Error", f"Failed to toggle Listen:\n{e}")
                self.set_status("Failed to toggle Listen")
        def open_volume_dialog(self, device_id, device_name):
            top = tk.Toplevel(self.root)
            try:
                if sys.platform.startswith("win"):
                    top.iconbitmap(resource_path("audio.ico"))
            except Exception:
                pass
            top.title("Set Volume")
            top.transient(self.root)
            top.grab_set()
            top.resizable(False, False)
            frm = ttk.Frame(top, padding=12)
            frm.pack(fill="both", expand=True)
            ttk.Label(frm, text=device_name, anchor="w").grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 8))
            initial = get_endpoint_volume(device_id)
            if initial is None:
                initial = 50
            v = tk.IntVar(value=initial)
            syncing = {"entry": False, "scale": False}
            # 3-digit entry (0-100) on the left
            def _validate(P):
                if P == "":
                    return True
                if not P.isdigit():
                    return False
                if len(P) > 3:
                    return False
                try:
                    val = int(P)
                except Exception:
                    return False
                return 0 <= val <= 100
            vcmd = (top.register(_validate), "%P")
            entry = ttk.Entry(frm, width=3, textvariable=v, validate="key", validatecommand=vcmd, justify="right")
            entry.grid(row=1, column=0, sticky="w")
            ttk.Label(frm, text="%").grid(row=1, column=1, sticky="w", padx=(4, 12))
            # Slider 0-100 on the right
            def on_scale(valstr):
                if syncing["entry"]:
                    return
                try:
                    syncing["scale"] = True
                    v.set(int(float(valstr)))
                finally:
                    syncing["scale"] = False
            scale = ttk.Scale(frm, from_=0, to=100, orient="horizontal", command=on_scale)
            scale.set(initial)
            scale.grid(row=1, column=2, sticky="we")
            frm.columnconfigure(2, weight=1)
            # Keep entry and scale in sync (entry -> scale)
            def on_entry_change(*_):
                if syncing["scale"]:
                    return
                try:
                    syncing["entry"] = True
                    try:
                        scale.set(int(v.get()))
                    except Exception:
                        pass
                finally:
                    syncing["entry"] = False
            v.trace_add("write", on_entry_change)
            # Buttons
            btns = ttk.Frame(frm)
            btns.grid(row=2, column=0, columnspan=3, sticky="e", pady=(12, 0))
            result = {"value": None}
            def ok():
                try:
                    result["value"] = max(0, min(100, int(v.get())))
                except Exception:
                    result["value"] = None
                top.destroy()
            def cancel():
                result["value"] = None
                top.destroy()
            ttk.Button(btns, text="OK", command=ok).pack(side="right")
            ttk.Button(btns, text="Cancel", command=cancel).pack(side="right", padx=(0, 8))
            top.bind("<Return>", lambda e: ok())
            top.bind("<Escape>", lambda e: cancel())
            entry.focus_set()
            top.wait_window()
            return result["value"]
    # Create and run the UI
    gui = AudioGUI(root)
    _log("launch_gui: entering mainloop")
    try:
        root.mainloop()
    except Exception:
        _log_exc("MAINLOOP EXCEPTION")
    _log("launch_gui: mainloop exited")
    try:
        CoUninitialize()
    except Exception:
        pass
    return 0

def main(argv=None):
    # If launched without any CLI arguments (double-click), open GUI
    if argv is None and len(sys.argv) <= 1:
        try:
            return launch_gui()
        except Exception as e:
            print(f"ERROR: GUI failed to start: {e}", file=sys.stderr)
    # Ensure COM is initialized for CLI operations as well
    try:
        CoInitialize()
    except Exception:
        pass
    parser = build_parser()
    args = parser.parse_args(argv)
    if args.cmd == "listen":
        if args.enable and args.disable:
            print("ERROR: specify only one of --enable or --disable", file=sys.stderr)
            try:
                CoUninitialize()
            except Exception:
                pass
            return 1
        if not args.enable and not args.disable:
            print("ERROR: specify --enable or --disable", file=sys.stderr)
            try:
                CoUninitialize()
            except Exception:
                pass
            return 1
        args.enable = True if args.enable else False
    if args.cmd == "set-volume":
        if (args.mute or args.unmute) and args.level is not None:
            print("ERROR: Cannot specify both --level and --mute/--unmute", file=sys.stderr)
            try:
                CoUninitialize()
            except Exception:
                pass
            return 1
        if not (args.mute or args.unmute or args.level is not None):
            print("ERROR: Must specify --level or --mute/--unmute", file=sys.stderr)
            try:
                CoUninitialize()
            except Exception:
                pass
            return 1
    try:
        rc = args.func(args)
    except KeyboardInterrupt:
        rc = 130
    finally:
        try:
            CoUninitialize()
        except Exception:
            pass
    return rc

if __name__ == "__main__":
    sys.exit(main())