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
    if not hasattr(_automation, "VT_UI2"):
        _automation.VT_UI2 = 18
    if not hasattr(_automation, "VT_UI4"):
        _automation.VT_UI4 = 19
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
import os  # Explicitly ensure 'os' is available
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
# Universal scan defaults (bounded)
# Known vendor DWORD toggles: (guid, [pid,...]) => 0 = enabled, 1 = disabled
# NOTE: This list is now replaced by the internal _CODE_VENDOR_ENTRIES map below.
# _KNOWN_VENDOR_TOGGLES = [
#     ("{1da5d803-d492-4edd-8c23-e0c0ffee7f0e}", [5]),  # Realtek/Waves primary on your system
# ]
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
def _guid_from_parts(*parts):
    # Assemble a GUID string from parts to avoid embedding the exact literal
    # Example: _guid_from_parts("E4870E26", "-3CC5-4CD2-", "BA46-", "CA0A9A70ED04")
    return "{" + "".join(parts) + "}"
# =========================
# SysFX (Audio Enhancements) helpers (additive-only)
# =========================
def _define_policyconfig_fx_interfaces():
    """
    Define a local IPolicyConfig interface with GetPropertyValue/SetPropertyValue
    that take the bFxStore flag, using runtime-assembled GUID strings.
    """
    import ctypes
    from ctypes import POINTER
    from ctypes import wintypes
    from comtypes import IUnknown, GUID, COMMETHOD, HRESULT
    class PROPERTYKEY(ctypes.Structure):
        _fields_ = (("fmtid", GUID), ("pid", wintypes.DWORD))
    # Prefer comtypes' PROPVARIANT; fallback to a minimal struct if missing
    try:
        import comtypes.automation as automation
        PROPVARIANT = getattr(automation, "PROPVARIANT", getattr(automation, "tagPROPVARIANT"))
    except Exception:
        class _PVU(ctypes.Union):
            _fields_ = [
                ("boolVal", ctypes.c_short),
                ("uiVal", ctypes.c_ushort),
                ("ulVal", ctypes.c_ulong),
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
    # IPolicyConfig IID {F8679F50-850A-41CF-9C72-430F290290C8}
    _IID_PolicyConfig = GUID(_guid_from_parts("F8679F50", "-850A-41CF-", "9C72-", "430F290290C8"))
    class IPolicyConfigFx(IUnknown):
        _iid_ = _IID_PolicyConfig
        _methods_ = (
            COMMETHOD([], HRESULT, 'GetMixFormat',
                      (['in'], wintypes.LPCWSTR, 'wszDeviceId'),
                      (['out'], POINTER(ctypes.c_void_p), 'ppFormat')),
            COMMETHOD([], HRESULT, 'GetDeviceFormat',
                      (['in'], wintypes.LPCWSTR, 'wszDeviceId'),
                      (['in'], ctypes.c_int, 'bDefault'),
                      (['out'], POINTER(ctypes.c_void_p), 'ppFormat')),
            COMMETHOD([], HRESULT, 'ResetDeviceFormat',
                      (['in'], wintypes.LPCWSTR, 'wszDeviceId')),
            COMMETHOD([], HRESULT, 'SetDeviceFormat',
                      (['in'], wintypes.LPCWSTR, 'wszDeviceId'),
                      (['in'], ctypes.c_void_p, 'pEndpointFormat'),
                      (['in'], ctypes.c_void_p, 'mixFormat')),
            COMMETHOD([], HRESULT, 'GetProcessingPeriod',
                      (['in'], wintypes.LPCWSTR, 'wszDeviceId'),
                      (['in'], ctypes.c_int, 'bDefault'),
                      (['out'], POINTER(ctypes.c_longlong), 'pmftDefaultPeriod'),
                      (['out'], POINTER(ctypes.c_longlong), 'pmftMinimumPeriod')),
            COMMETHOD([], HRESULT, 'SetProcessingPeriod',
                      (['in'], wintypes.LPCWSTR, 'wszDeviceId'),
                      (['in'], POINTER(ctypes.c_longlong), 'pmftPeriod')),
            COMMETHOD([], HRESULT, 'GetShareMode',
                      (['in'], wintypes.LPCWSTR, 'wszDeviceId'),
                      (['out'], POINTER(ctypes.c_void_p), 'pMode')),
            COMMETHOD([], HRESULT, 'SetShareMode',
                      (['in'], wintypes.LPCWSTR, 'wszDeviceId'),
                      (['in'], ctypes.c_void_p, 'mode')),
            COMMETHOD([], HRESULT, 'GetPropertyValue',
                      (['in'], wintypes.LPCWSTR, 'pszDeviceName'),
                      (['in'], wintypes.BOOL, 'bFxStore'),
                      (['in'], POINTER(PROPERTYKEY), 'pKey'),
                      (['out'], POINTER(PROPVARIANT), 'pv')),
            COMMETHOD([], HRESULT, 'SetPropertyValue',
                      (['in'], wintypes.LPCWSTR, 'pszDeviceName'),
                      (['in'], wintypes.BOOL, 'bFxStore'),
                      (['in'], POINTER(PROPERTYKEY), 'pKey'),
                      (['in'], POINTER(PROPVARIANT), 'pv')),
            COMMETHOD([], HRESULT, 'SetDefaultEndpoint',
                      (['in'], wintypes.LPCWSTR, 'wszDeviceId'),
                      (['in'], wintypes.DWORD, 'role')),
            COMMETHOD([], HRESULT, 'SetEndpointVisibility',
                      (['in'], wintypes.LPCWSTR, 'wszDeviceId'),
                      (['in'], wintypes.BOOL, 'bVisible')),
        )
    # CLSID_PolicyConfigClient {870AF99C-171D-4F9E-AF0D-E63DF40C2BC9}
    CLSID_PolicyConfigClient = GUID(_guid_from_parts("870AF99C", "-171D-4F9E-", "AF0D-", "E63DF40C2BC9"))
    return IPolicyConfigFx, CLSID_PolicyConfigClient, PROPERTYKEY, PROPVARIANT
def _get_policy_config_fx():
    from comtypes import CoCreateInstance, CLSCTX_ALL
    IPolicyConfigFx, CLSID_PolicyConfigClient, PROPERTYKEY, PROPVARIANT = _define_policyconfig_fx_interfaces()
    return CoCreateInstance(CLSID_PolicyConfigClient, interface=IPolicyConfigFx, clsctx=CLSCTX_ALL)
def _pkey_disable_sysfx():
    # PKEY_AudioEndpoint_Disable_SysFx {E4870E26-3CC5-4CD2-BA46-CA0A9A70ED04}, pid 2
    from comtypes import GUID
    import ctypes
    from ctypes import wintypes
    IPolicyConfigFx, CLSID_PolicyConfigClient, PROPERTYKEY, PROPVARIANT = _define_policyconfig_fx_interfaces()
    g = _guid_from_parts("E4870E26", "-3CC5-4CD2-", "BA46-", "CA0A9A70ED04")
    return PROPERTYKEY(GUID(g), wintypes.DWORD(2))
def _parse_boolish_from_propvariant(pv):
    try:
        import comtypes.automation as automation
        VT_BOOL = getattr(automation, "VT_BOOL", 11)
    except Exception:
        VT_BOOL = 11
    VT_UI2 = 18
    VT_UI4 = 19
    try:
        vt = getattr(pv, "vt", 0)
    except Exception:
        return None
    try:
        if vt == VT_BOOL:
            val = getattr(pv, "boolVal", 0)
            return 0 if int(val) == 0 else 1
        if vt == VT_UI2:
            if hasattr(pv, "uiVal"):
                return 1 if int(pv.uiVal) != 0 else 0
            if hasattr(pv, "value") and hasattr(pv.value, "uiVal"):
                return 1 if int(pv.value.uiVal) != 0 else 0
        if vt == VT_UI4:
            if hasattr(pv, "ulVal"):
                return 1 if int(pv.ulVal) != 0 else 0
            if hasattr(pv, "value") and hasattr(pv.value, "ulVal"):
                return 1 if int(pv.value.ulVal) != 0 else 0
    except Exception:
        return None
    return None
def _set_boolish_in_propvariant(pv, zero_or_one):
    try:
        import comtypes.automation as automation
        VT_BOOL = getattr(automation, "VT_BOOL", 11)
    except Exception:
        VT_BOOL = 11
    VT_UI2 = 18
    VT_UI4 = 19
    v = 1 if bool(zero_or_one) else 0
    vt = getattr(pv, "vt", 0)
    try:
        if vt == VT_BOOL:
            setattr(pv, "boolVal", -1 if v else 0)
            return True
        if vt == VT_UI2:
            if hasattr(pv, "uiVal"):
                setattr(pv, "uiVal", v); return True
            if hasattr(pv, "value") and hasattr(pv.value, "uiVal"):
                setattr(pv.value, "uiVal", v); return True
        if vt == VT_UI4:
            if hasattr(pv, "ulVal"):
                setattr(pv, "ulVal", v); return True
            if hasattr(pv, "value") and hasattr(pv.value, "ulVal"):
                setattr(pv.value, "ulVal", v); return True
    except Exception:
        return False
    try:
        if hasattr(pv, "uiVal"):
            pv.uiVal = v; return True
    except Exception:
        pass
    try:
        if hasattr(pv, "ulVal"):
            pv.ulVal = v; return True
    except Exception:
        pass
    return False
def _get_enhancements_status_com(device_id):
    """
    Returns True if enhancements are enabled, False if disabled, or None if unknown.
    Tries both FX store (bFxStore=True) and normal store (bFxStore=False).
    """
    try:
        IPolicyConfigFx, CLSID_PolicyConfigClient, PROPERTYKEY, PROPVARIANT = _define_policyconfig_fx_interfaces()
        pkey = _pkey_disable_sysfx()
        pc = _get_policy_config_fx()
        for bfx in (True, False):
            pv = PROPVARIANT()
            try:
                pc.GetPropertyValue(device_id, bfx, byref(pkey), byref(pv))
                raw = _parse_boolish_from_propvariant(pv)  # Disable_SysFx: 0=enh on, 1=enh off
                if raw is None:
                    continue
                return False if raw == 1 else True
            except Exception:
                continue
        return None
    except Exception:
        return None
def _set_enhancements_com(device_id, enable):
    """
    Set Disable_SysFx to desired value in both stores (FX and normal).
    enable=True -> Disable_SysFx=0
    Returns True if any write succeeded.
    """
    try:
        IPolicyConfigFx, CLSID_PolicyConfigClient, PROPERTYKEY, PROPVARIANT = _define_policyconfig_fx_interfaces()
        pkey = _pkey_disable_sysfx()
        pc = _get_policy_config_fx()
        desired_disable = 0 if enable else 1
        ok_any = False
        for bfx in (True, False):
            try:
                pv = PROPVARIANT()
                # Read current (to get correct VT), ignore errors
                try:
                    pc.GetPropertyValue(device_id, bfx, byref(pkey), byref(pv))
                except Exception:
                    pass
                if not _set_boolish_in_propvariant(pv, desired_disable):
                    try:
                        pv.vt = 19  # VT_UI4
                        pv.ulVal = desired_disable
                    except Exception:
                        pass
                pc.SetPropertyValue(device_id, bfx, byref(pkey), byref(pv))
                ok_any = True
            except Exception:
                continue
        return ok_any
    except Exception:
        return False
def _read_enhancements_from_registry(device_id):
    r"""
    Read enhancements state (enabled/disabled) via registry.
    Returns True (enabled) / False (disabled) / None (unknown).
    Searches BOTH HKCU and HKLM under:
      ...\MMDevices\Audio\{Render|Capture}\{guid}\{FxProperties|Properties}
    Prefers ",2" if present (common pid for Disable_SysFx).
    """
    try:
        import winreg
    except Exception:
        return None
    guid = _extract_endpoint_guid_from_device_id(device_id)
    if not guid:
        return None
    fmtid = "{e4870e26-3cc5-4cd2-ba46-ca0a9a70ed04}".lower()
    def _parse_bool_from_reg(val, typ):
        # REG_DWORD: 0/1
        if typ == winreg.REG_DWORD:
            try:
                return bool(int(val))
            except Exception:
                return None
        if typ == winreg.REG_BINARY:
            try:
                b = bytes(val)
                if len(b) >= 12:
                    vt = int.from_bytes(b[0:2], "little", signed=False)
                    if vt == 0x000B:  # VT_BOOL
                        return int.from_bytes(b[8:10], "little", signed=True) != 0
                    if vt == 0x0013:  # VT_UI4
                        return int.from_bytes(b[8:12], "little", signed=False) != 0
            except Exception:
                return None
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
    hive_list = [
        (winreg.HKEY_CURRENT_USER,  "HKCU"),
        (winreg.HKEY_LOCAL_MACHINE, "HKLM"),
    ]
    preferred = None
    fallback = None
    found_preferred = False
    for hive, _hn in hive_list:
        for flow in ("Render", "Capture"):
            for sub in ("FxProperties", "Properties"):
                key_path = rf"SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\{flow}\{guid}\{sub}"
                try:
                    key = winreg.OpenKey(hive, key_path, 0, winreg.KEY_READ)
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
                        nl = name.lower()
                        if not nl.startswith(fmtid):
                            continue
                        parsed = _parse_bool_from_reg(val, typ)
                        if parsed is None:
                            continue
                        if nl.endswith(",2"):
                            preferred = parsed
                            found_preferred = True
                            break
                        if fallback is None:
                            fallback = parsed
                finally:
                    try:
                        winreg.CloseKey(key)
                    except Exception:
                        pass
                if found_preferred:
                    break
            if found_preferred:
                break
        if found_preferred:
            break
    # Registry stores Disable_SysFx: True means DISABLED; we return 'enabled' boolean.
    if preferred is not None:
        return False if preferred else True
    if fallback is not None:
        return False if fallback else True
    return None
def _set_enhancements_registry(device_id, enable, prefer_hklm=False):
    """
    Fallback: write Disable_SysFx to registry (DWORD 0/1) under BOTH
    FxProperties and Properties for Render/Capture. Tries HKCU then HKLM
    (or HKLM first if prefer_hklm=True). Returns True if any write succeeded.
    Note: HKLM writes require Administrator.
    """
    try:
        import winreg
    except Exception:
        return False
    guid = _extract_endpoint_guid_from_device_id(device_id)
    if not guid:
        return False
    # Value name: Disable_SysFx pid 2
    name = "{e4870e26-3cc5-4cd2-ba46-ca0a9a70ed04},2"
    desired_disable = 0 if enable else 1
    # Decide hive order
    hive_order = [
        (winreg.HKEY_CURRENT_USER,  "HKCU"),
        (winreg.HKEY_LOCAL_MACHINE, "HKLM"),
    ]
    if prefer_hklm:
        hive_order = [
            (winreg.HKEY_LOCAL_MACHINE, "HKLM"),
            (winreg.HKEY_CURRENT_USER,  "HKCU"),
        ]
    ok_any = False
    for hive, _hn in hive_order:
        for flow in ("Render", "Capture"):
            for sub in ("FxProperties", "Properties"):
                key_path = rf"SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\{flow}\{guid}\{sub}"
                try:
                    key = winreg.OpenKey(hive, key_path, 0, winreg.KEY_SET_VALUE)
                except OSError:
                    continue
                try:
                    winreg.SetValueEx(key, name, 0, winreg.REG_DWORD, int(desired_disable))
                    ok_any = True
                except OSError:
                    pass
                finally:
                    try: winreg.CloseKey(key)
                    except Exception: pass
    return ok_any
def _verify_enhancements_via_registry(device_id, expected_enabled, timeout=2.0, interval=0.15):
    import time as _time
    deadline = _time.time() + timeout
    last_state = None
    while _time.time() < deadline:
        try:
            state = _read_enhancements_from_registry(device_id)
        except Exception:
            state = None
        last_state = state
        if state is not None and state == expected_enabled:
            return True, state
        _time.sleep(interval)
    return False, last_state
def _dump_mmdevices_all_values(device_id):
    r"""
    Dump ALL values under BOTH hives for this endpoint:
      HKCU/HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\{Render|Capture}\{guid}\{FxProperties|Properties}
    Returns a JSON-serializable list of {hive, flow, subkey, name, type, dataPreview}.
    """
    try:
        import winreg
    except Exception:
        return {"error": "winreg unavailable"}
    guid = _extract_endpoint_guid_from_device_id(device_id)
    if not guid:
        return {"error": "bad endpoint id, cannot extract guid"}
    items = []
    roots = [
        (winreg.HKEY_CURRENT_USER,  "HKCU"),
        (winreg.HKEY_LOCAL_MACHINE, "HKLM"),
    ]
    for hive, hive_name in roots:
        for flow in ("Render", "Capture"):
            for sub in ("FxProperties", "Properties"):
                key_path = rf"SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\{flow}\{guid}\{sub}"
                try:
                    key = winreg.OpenKey(hive, key_path, 0, winreg.KEY_READ)
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
                        rec = {
                            "hive": hive_name,
                            "flow": flow,
                            "subkey": sub,
                            "name": name,
                            "type": typ,
                        }
                        try:
                            if typ == winreg.REG_DWORD:
                                rec["dataPreview"] = int(val)
                            elif typ == winreg.REG_SZ:
                                rec["dataPreview"] = str(val)
                            elif typ == winreg.REG_BINARY:
                                b = bytes(val)
                                rec["dataPreview"] = "hex:" + b[:16].hex() + (f"...({len(b)})" if len(b) > 16 else "")
                            else:
                                rec["dataPreview"] = f"<type {typ}>"
                        except Exception:
                            rec["dataPreview"] = "<unreadable>"
                        items.append(rec)
                finally:
                    try:
                        winreg.CloseKey(key)
                    except Exception:
                        pass
    return items
def _get_enhancements_status_propstore(device_id):
    """
    Read Disable_SysFx directly from the endpoint's IPropertyStore.
    Returns:
      True  -> enhancements enabled
      False -> enhancements disabled
      None  -> unknown/not present
    """
    try:
        import ctypes
        from ctypes import wintypes, byref, POINTER
        from comtypes import GUID
        # Ensure an HRESULT type for PropVariantClear signature
        try:
            HRESULT_T = wintypes.HRESULT
        except Exception:
            HRESULT_T = ctypes.c_long
        # PROPVARIANT
        try:
            import comtypes.automation as automation
            PROPVARIANT = getattr(automation, "PROPVARIANT", getattr(automation, "tagPROPVARIANT"))
        except Exception:
            class _PVU(ctypes.Union):
                _fields_ = [("boolVal", ctypes.c_short), ("uiVal", ctypes.c_ushort), ("ulVal", ctypes.c_ulong)]
            class PROPVARIANT(ctypes.Structure):
                _anonymous_ = ("data",)
                _fields_ = [
                    ("vt", ctypes.c_ushort),
                    ("wReserved1", ctypes.c_ushort),
                    ("wReserved2", ctypes.c_ushort),
                    ("wReserved3", ctypes.c_ushort),
                    ("data", _PVU),
                ]
        CALL = ctypes.WINFUNCTYPE if ctypes.sizeof(ctypes.c_void_p) == 4 else ctypes.CFUNCTYPE
        class PROPERTYKEY(ctypes.Structure):
            _fields_ = (("fmtid", GUID), ("pid", wintypes.DWORD))
        class IPropertyStoreRaw(ctypes.Structure):
            pass
        PIPS = POINTER(IPropertyStoreRaw)
        GetValueProto = CALL(HRESULT_T, PIPS, POINTER(PROPERTYKEY), POINTER(PROPVARIANT))
        class IPropertyStoreVTBL(ctypes.Structure):
            _fields_ = [
                ("QueryInterface", ctypes.c_void_p),
                ("AddRef", ctypes.c_void_p),
                ("Release", ctypes.c_void_p),
                ("GetCount", ctypes.c_void_p),
                ("GetAt", ctypes.c_void_p),
                ("GetValue", GetValueProto),
                ("SetValue", ctypes.c_void_p),
                ("Commit", ctypes.c_void_p),
            ]
        IPropertyStoreRaw._fields_ = [("lpVtbl", POINTER(IPropertyStoreVTBL))]
        enumerator = CoCreateInstance(CLSID_MMDeviceEnumerator, interface=IMMDeviceEnumerator, clsctx=CLSCTX_ALL)
        dev = enumerator.GetDevice(device_id)
        ps_unknown = dev.OpenPropertyStore(STGM_READ)
        ps_ptr_val = ctypes.cast(ps_unknown, ctypes.c_void_p).value
        if not ps_ptr_val:
            return None
        ps_iface = ctypes.cast(ctypes.c_void_p(ps_ptr_val), PIPS)
        pkey = PROPERTYKEY(GUID("{E4870E26-3CC5-4CD2-BA46-CA0A9A70ED04}"), wintypes.DWORD(2))
        pv = PROPVARIANT()
        hr = ps_iface.contents.lpVtbl.contents.GetValue(ps_iface, byref(pkey), byref(pv))
        if hr != 0:
            return None
        # Parse before clearing the PROPVARIANT
        raw = _parse_boolish_from_propvariant(pv)  # 0 = enh ON, 1 = enh OFF
        # Clear PROPVARIANT
        try:
            ole32 = ctypes.OleDLL("ole32.dll")
            PropVariantClear = getattr(ole32, "PropVariantClear", None)
            if PropVariantClear:
                PropVariantClear.restype = HRESULT_T
                PropVariantClear.argtypes = (ctypes.POINTER(PROPVARIANT),)
                PropVariantClear(byref(pv))
        except Exception:
            pass
        if raw is None:
            return None
        return False if raw == 1 else True
    except Exception:
        return None
def _set_enhancements_propstore(device_id, enable):
    """
    Write Disable_SysFx directly via IPropertyStore::SetValue + Commit.
    enable=True  -> Disable_SysFx=0
    enable=False -> Disable_SysFx=1
    Returns True on success.
    """
    try:
        import ctypes
        from ctypes import wintypes, POINTER, byref
        from comtypes import GUID
        # PROPVARIANT (prefer comtypes; fallback minimal)
        try:
            import comtypes.automation as automation
            PROPVARIANT = getattr(automation, "PROPVARIANT", getattr(automation, "tagPROPVARIANT"))
        except Exception:
            class _PVU(ctypes.Union):
                _fields_ = [("boolVal", ctypes.c_short), ("uiVal", ctypes.c_ushort), ("ulVal", ctypes.c_ulong)]
            class PROPVARIANT(ctypes.Structure):
                _anonymous_ = ("data",)
                _fields_ = [
                    ("vt", ctypes.c_ushort),
                    ("wReserved1", ctypes.c_ushort),
                    ("wReserved2", ctypes.c_ushort),
                    ("wReserved3", ctypes.c_ushort),
                    ("data", _PVU),
                ]
        CALL = ctypes.WINFUNCTYPE if ctypes.sizeof(ctypes.c_void_p) == 4 else ctypes.CFUNCTYPE
        class PROPERTYKEY(ctypes.Structure):
            _fields_ = (("fmtid", GUID), ("pid", wintypes.DWORD))
        class IPropertyStoreRaw(ctypes.Structure):
            pass
        PIPS = POINTER(IPropertyStoreRaw)
        try:
            HRESULT_T = wintypes.HRESULT
        except Exception:
            HRESULT_T = ctypes.c_long
        GetValueProto = CALL(HRESULT_T, PIPS, POINTER(PROPERTYKEY), POINTER(PROPVARIANT))
        SetValueProto = CALL(HRESULT_T, PIPS, POINTER(PROPERTYKEY), POINTER(PROPVARIANT))
        CommitProto   = CALL(HRESULT_T, PIPS)
        class IPropertyStoreVTBL(ctypes.Structure):
            _fields_ = [
                ("QueryInterface", ctypes.c_void_p),
                ("AddRef", ctypes.c_void_p),
                ("Release", ctypes.c_void_p),
                ("GetCount", ctypes.c_void_p),
                ("GetAt", ctypes.c_void_p),
                ("GetValue", GetValueProto),
                ("SetValue", SetValueProto),
                ("Commit", CommitProto),
            ]
        IPropertyStoreRaw._fields_ = [("lpVtbl", POINTER(IPropertyStoreVTBL))]
        # Open for WRITE
        enumerator = CoCreateInstance(CLSID_MMDeviceEnumerator, interface=IMMDeviceEnumerator, clsctx=CLSCTX_ALL)
        dev = enumerator.GetDevice(device_id)
        ps_unknown = dev.OpenPropertyStore(STGM_WRITE)
        ps_ptr_val = ctypes.cast(ps_unknown, ctypes.c_void_p).value
        if not ps_ptr_val:
            return False
        ps_iface = ctypes.cast(ctypes.c_void_p(ps_ptr_val), PIPS)
        # Key and value
        pkey = PROPERTYKEY(GUID("{E4870E26-3CC5-4CD2-BA46-CA0A9A70ED04}"), wintypes.DWORD(2))
        desired_disable = 0 if enable else 1
        pv = PROPVARIANT()
        # Prefer to read once to get an existing VT, then set; if that fails, force VT_UI4
        try:
            ps_iface.contents.lpVtbl.contents.GetValue(ps_iface, byref(pkey), byref(pv))
            if not _set_boolish_in_propvariant(pv, desired_disable):
                pv.vt = 19  # VT_UI4
                try:
                    pv.ulVal = desired_disable
                except Exception:
                    pass
        except Exception:
            pv = PROPVARIANT()
            pv.vt = 19  # VT_UI4
            try:
                pv.ulVal = desired_disable
            except Exception:
                pass
        hr = ps_iface.contents.lpVtbl.contents.SetValue(ps_iface, byref(pkey), byref(pv))
        if hr != 0:
            return False
        hr = ps_iface.contents.lpVtbl.contents.Commit(ps_iface)
        return hr == 0
    except Exception:
        return False
def _wait_for_propstore_sysfx(device_id, expected_enabled, timeout=1.5, interval=0.12):
    """
    Poll the endpoint's IPropertyStore for Disable_SysFx until it matches expected_enabled
    or timeout. Returns (True, state) on success; (False, last_state_or_None) on timeout.
    """
    import time
    last = None
    end = time.time() + float(timeout)
    while time.time() < end:
        state = _get_enhancements_status_propstore(device_id)
        last = state
        if state is not None and state == expected_enabled:
            return True, state
        time.sleep(interval)
    return False, last

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

def _collect_sysfx_snapshot(device_id):
    """
    Collects a full snapshot for discovering how 'Audio Enhancements' toggles on this device.
    Returns a dict with:
      {
        "time": "...",
        "com": {"fxStore": {...}, "normalStore": {...}},
        "propStore": {"enhEnabled": True/False/None},
        "registry": [ list from _dump_mmdevices_all_values(device_id) ]
      }
    """
    import datetime
    snap = {
        "time": datetime.datetime.now().isoformat(timespec="seconds"),
        "com": {},
        "propStore": {},
        "registry": [],
    }
    # COM (both stores)
    try:
        IPolicyConfigFx, CLSID_PolicyConfigClient, PROPERTYKEY, PROPVARIANT = _define_policyconfig_fx_interfaces()
        pkey = _pkey_disable_sysfx()
        pc = _get_policy_config_fx()
        from ctypes import byref
        for bfx, label in ((True, "fxStore"), (False, "normalStore")):
            pv = PROPVARIANT()
            rec = {}
            try:
                pc.GetPropertyValue(device_id, bfx, byref(pkey), byref(pv))
                raw = _parse_boolish_from_propvariant(pv)  # Disable_SysFx: 0=enh on, 1=off
                rec["rawDisable"] = raw
                rec["enhEnabled"] = (False if raw == 1 else True) if raw is not None else None
            except Exception as e:
                rec["error"] = str(e)
            snap["com"][label] = rec
    except Exception as e:
        snap["com"] = {"error": str(e)}
    # Property store (live)
    try:
        enh = _get_enhancements_status_propstore(device_id)
        snap["propStore"] = {"enhEnabled": enh}
    except Exception as e:
        snap["propStore"] = {"error": str(e)}
    # Registry (all values under MMDevices for this endpoint)
    try:
        snap["registry"] = _dump_mmdevices_all_values(device_id)
    except Exception as e:
        snap["registry"] = [{"error": str(e)}]
    return snap
def _mmdev_key_of(rec):
    return f"{rec.get('hive','?')}|{rec.get('flow','?')}|{rec.get('subkey','?')}|{rec.get('name','?')}"
def _normalize_preview(v):
    # Our dump uses ints for REG_DWORD, str for REG_SZ, and "hex:..." for REG_BINARY preview
    try:
        return (v if isinstance(v, (int, float)) else str(v)).strip() if isinstance(v, str) else v
    except Exception:
        return v
def _diff_mmdevices_lists(before_list, after_list):
    """
    Diff two mmdevices lists (as produced by _dump_mmdevices_all_values).
    Returns {
      "added": [...],
      "removed": [...],
      "changed": [...],  # type or dataPreview changed
      "dword_flips": [...],  # only REG_DWORD 0<->1 flips
      "disable_sysfx_hits": [...],  # entries whose name starts with {E4870E26...}
    }
    """
    import copy
    idxA = {_mmdev_key_of(e): e for e in (before_list or [])}
    idxB = {_mmdev_key_of(e): e for e in (after_list or [])}
    all_keys = set(idxA.keys()) | set(idxB.keys())
    added = []
    removed = []
    changed = []
    flips = []
    hits = []
    # pattern match for Disable_SysFx GUID prefix
    guid_disable = "{e4870e26-3cc5-4cd2-ba46-ca0a9a70ed04}"
    for k in sorted(all_keys):
        a = idxA.get(k)
        b = idxB.get(k)
        if a is None:
            added.append(b)
            if str(b.get("name", "")).lower().startswith(guid_disable):
                hits.append(b)
            continue
        if b is None:
            removed.append(a)
            if str(a.get("name", "")).lower().startswith(guid_disable):
                hits.append(a)
            continue
        # both present -> compare
        try:
            tA = a.get("type")
            tB = b.get("type")
            vA = _normalize_preview(a.get("dataPreview"))
            vB = _normalize_preview(b.get("dataPreview"))
            if (tA != tB) or (vA != vB):
                row = copy.deepcopy(a)
                row["typeAfter"] = tB
                row["dataPreviewAfter"] = vB
                changed.append(row)
            # DWORD flip special case
            try:
                if tA == 4 and tB == 4:
                    ia = int(a.get("dataPreview"))
                    ib = int(b.get("dataPreview"))
                    if (ia == 0 and ib == 1) or (ia == 1 and ib == 0):
                        parts = k.split("|", 3)
                        flips.append({
                            "hive": parts[0], "flow": parts[1], "subkey": parts[2], "name": parts[3],
                            "before": ia, "after": ib
                        })
            except Exception:
                pass
            # record Disable_SysFx hits
            if str(a.get("name", "")).lower().startswith(guid_disable) or str(b.get("name","")).lower().startswith(guid_disable):
                hits.append(b)
        except Exception:
            continue
    return {
        "added": added,
        "removed": removed,
        "changed": changed,
        "dword_flips": flips,
        "disable_sysfx_hits": hits,
    }
def _generate_enh_discovery_report(target, snapA, snapB, diffs):
    """
    Build a human-readable text report string from snapshots and diff.
    """
    import datetime
    lines = []
    lines.append("Audio Enhancements (SysFx) Discovery Report")
    lines.append("=" * 60)
    lines.append(f"Generated: {datetime.datetime.now().isoformat(timespec='seconds')}")
    lines.append(f"Device:    {target.get('name')} [{target.get('id')}]")
    lines.append(f"Flow:      {target.get('flow')}")
    lines.append("")
    def _fmt_bool(x):
        return "True" if x is True else ("False" if x is False else "None")
    # COM summary
    lines.append("COM (PolicyConfig) - Disable_SysFx (0=Enh ON, 1=OFF)")
    for label in ("fxStore", "normalStore"):
        A = snapA.get("com", {}).get(label, {})
        B = snapB.get("com", {}).get(label, {})
        lines.append(f"  {label:12} A: rawDisable={A.get('rawDisable')} -> enhEnabled={_fmt_bool(A.get('enhEnabled'))} "
                     f"| B: rawDisable={B.get('rawDisable')} -> enhEnabled={_fmt_bool(B.get('enhEnabled'))}")
    # PropStore summary
    Aps = snapA.get("propStore", {}).get("enhEnabled")
    Bps = snapB.get("propStore", {}).get("enhEnabled")
    lines.append(f"PropertyStore live: A.enhEnabled={_fmt_bool(Aps)}  |  B.enhEnabled={_fmt_bool(Bps)}")
    lines.append("")
    lines.append("Registry (MMDevices) diff summary")
    lines.append(f"  Added:   {len(diffs.get('added', []))}")
    lines.append(f"  Removed: {len(diffs.get('removed', []))}")
    lines.append(f"  Changed: {len(diffs.get('changed', []))}")
    lines.append(f"  DWORD flips (0<->1): {len(diffs.get('dword_flips', []))}")
    lines.append("")
    # Highlight Disable_SysFx entries if present
    ds_hits = [e for e in diffs.get("changed", []) if str(e.get('name','')).lower().startswith("{e4870e26-3cc5-4cd2-ba46-ca0a9a70ed04}")]
    if ds_hits:
        lines.append("Disable_SysFx registry entries that changed:")
        for e in ds_hits:
            lines.append(f"  {e.get('hive')}\\{e.get('flow')}\\{e.get('subkey')}\\{e.get('name')} "
                         f"{e.get('dataPreview')} -> {e.get('dataPreviewAfter')} (type {e.get('type')} -> {e.get('typeAfter')})")
        lines.append("")
    # Show boolean-like flips (strong candidates)
    flips = diffs.get("dword_flips", [])
    if flips:
        lines.append("Candidate toggle keys (REG_DWORD flips 0<->1):")
        for f in flips:
            lines.append(f"  {f['hive']}\\{f['flow']}\\{f['subkey']}\\{f['name']}  {f['before']} -> {f['after']}")
        lines.append("")
    else:
        lines.append("No DWORD 0/1 flips detected. Vendor may use non-DWORD or a different location.")
        lines.append("")
    # Next steps suggestion
    lines.append("Notes:")
    lines.append("- If COM/PropertyStore show A!=B, Windows honored Disable_SysFx and the existing setter is correct.")
    lines.append("- If COM/PropertyStore stay the same but a vendor REG_DWORD flips, that key is likely the real toggle.")
    lines.append("- If only REG_BINARY blobs changed, we may need to write that vendor-specific property.")
    lines.append("")
    return "\n".join(lines)
# --- Vendor-specific logic (formerly _KNOWN_VENDOR_TOGGLES and related functions) is removed/replaced ---
# --- Now using generic vendor INI system and hardcoded vendor list:
# =========================
# Vendor toggles: code list + INI list
# =========================
import configparser
# Code vendor entries (preferred first. Add more here over time.
# Schema per entry:
#  {
#    "name": "short label",
#    "value_name": "{GUID},pid",  # REG_DWORD under endpoint FxProperties
#    "enable": 0,                 # DWORD value when enhancements enabled
#    "disable": 1,                # DWORD value when enhancements disabled
#    "hives": ["HKLM","HKCU"],    # where to write; HKLM requires admin
#    "flows": ["Render","Capture"]
#  }
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
    Load vendor toggles from INI. Returns list of normalized entries (same schema as code list).
    INI schema:
      [vendor_name]
      value_name = {GUID},pid
      dword_enable = 0
      dword_disable = 1
      hives = HKLM,HKCU
      flows = Render,Capture
      notes = optional
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
    try:
        import winreg as H
    except Exception:
        return False
    flow_name = "Render" if str(flow).lower().startswith("r") else "Capture"
    if entry.get("flows") and flow_name not in entry["flows"]:
        return False
    _f, key_path = _endpoint_fx_key(device_id, flow)
    if not key_path:
        return False
    for h in entry.get("hives", []):
        hive = H.HKEY_LOCAL_MACHINE if h == "HKLM" else H.HKEY_CURRENT_USER
        try:
            with H.OpenKey(hive, key_path, 0, H.KEY_READ) as key:
                try:
                    _, typ = H.QueryValueEx(key, entry["value_name"])
                    if typ == H.REG_DWORD:
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
    try:
        import winreg as H
    except Exception:
        return False
    _f, key_path = _endpoint_fx_key(device_id, flow)
    if not key_path:
        return False
    desired = entry["enable"] if enable else entry["disable"]
    ok = False
    for h in entry.get("hives", []):
        hive = H.HKEY_LOCAL_MACHINE if h == "HKLM" else H.HKEY_CURRENT_USER
        try:
            with H.OpenKey(hive, key_path, 0, H.KEY_SET_VALUE) as key:
                H.SetValueEx(key, entry["value_name"], 0, H.REG_DWORD, int(desired))
                ok = True or ok
        except OSError:
            continue
    return ok
def _read_vendor_entry_state(entry, device_id, flow):
    """
    Return True if current DWORD equals 'enable' value, False if equals 'disable', None otherwise.
    Prefer HKCU over HKLM for reading (drivers often update HKCU at runtime).
    Only read hives that are listed in entry['hives'].
    """
    try:
        import winreg as H
    except Exception:
        return None

    _f, key_path = _endpoint_fx_key(device_id, flow)
    if not key_path:
        return None

    configured = entry.get("hives") or []
    read_order = [h for h in ("HKCU", "HKLM") if h in configured]
    hive_map = {"HKCU": H.HKEY_CURRENT_USER, "HKLM": H.HKEY_LOCAL_MACHINE}

    for hname in read_order:
        hive = hive_map[hname]
        try:
            with H.OpenKey(hive, key_path, 0, H.KEY_READ) as key:
                try:
                    val, typ = H.QueryValueEx(key, entry["value_name"])
                    if typ == H.REG_DWORD:
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
    import time
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
    Try INI vendor entries first (user-learned), then code vendors (e.g., Realtek).
    If a matching vendor entry applies:
      - write the vendor DWORD, then verify it via the same DWORD.
      - return (True, f"vendor:{name}", state) on success; else fall through.
    """
    # 1) INI vendors first (user-learned should take precedence)
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

    # 2) Code vendors next (fallback, e.g., Realtek/Waves)
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
    Priority: INI vendors first (user-learned), then code vendors (e.g., Realtek).
    Does not write anything.
    """
    # 1) INI vendors first (user-learned should supersede built-ins)
    for entry in _load_vendor_db(ini_path):
        try:
            if _vendor_entry_applies(entry, device_id, flow):
                return entry
        except Exception:
            continue

    # 2) Code vendors next (fallback, e.g., Realtek/Waves)
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
    - If file exists and section exists: do nothing, return "exists".
    - If file exists and section missing: append a new section at the end, return "appended".
    - If file does not exist: create it with the section, return "appended".
    Never rewrites or removes other sections/comments.
    """
    # 1) Check existence via ConfigParser
    cfg = configparser.ConfigParser()
    try:
        if os.path.exists(ini_path):
            cfg.read(ini_path, encoding="utf-8")
    except Exception:
        # If parsing fails, we will still append a section (best-effort)
        pass

    if cfg.has_section(section_name):
        return "exists"

    # 2) Ensure directory exists (if path has a directory component)
    try:
        ini_dir = os.path.dirname(ini_path)
        if ini_dir:
            os.makedirs(ini_dir, exist_ok=True)
    except Exception:
        pass

    # 3) Build the section text
    lines = []
    lines.append("")  # ensure a leading newline to separate from previous content
    lines.append(f"[{section_name}]")
    lines.append(f"value_name = {str(value_name).strip().lower()}")
    lines.append(f"dword_enable = {int(dword_enable)}")
    lines.append(f"dword_disable = {int(dword_disable)}")
    lines.append(f"hives = {hives}")
    lines.append(f"flows = {flows}")
    if notes:
        lines.append(f"notes = {notes}")
    text = "\n".join(lines) + "\n"

    # 4) Append atomically (best-effort)
    with open(ini_path, "a", encoding="utf-8", errors="replace") as f:
        f.write(text)

    return "appended"
def _learn_vendor_from_discovery_and_write_ini(target, ini_path=None, prefer_hkcu=True):
    """
    Manual learn using discovery flow:
      - Prompt user to set Enhancements ENABLED -> capture snapshot A
      - Prompt user to set Enhancements DISABLED -> capture snapshot B
      - Diff registry, pick REG_DWORD flip under FxProperties
      - Write/update INI (create if needed)
    No TXT/JSON artifacts are written. Returns (ok, info_dict|error_msg).
    """
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
    dword_enable  = int(picked["before"])   # A = Enabled
    dword_disable = int(picked["after"])    # B = Disabled
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
    Auto-learn a vendor DWORD toggle for target {'id','name','flow'}:
      - drive to Enabled (Windows path), snapshot A
      - drive to Disabled (Windows path), snapshot B
      - diff MMDevices and pick a REG_DWORD flip under FxProperties
      - write/update INI section (create file if needed)
      - restore original state
    Returns (ok, info_dict or error_msg).
    """
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

    # Remember original state to restore later (best-effort)
    orig = _get_enhancements_status_propstore(dev_id)
    if orig is None:
        orig = _get_enhancements_status_com(dev_id)
    # Force Enabled using Windows path only (avoid vendor-first interference while learning)
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
    # Force Disabled using Windows path only
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
    # Diff and build a candidate vendor INI entry
    diffs = _diff_mmdevices_lists(snapA.get("registry") or [], snapB.get("registry") or [])
    snippet, picked = _build_vendor_ini_snippet(target, snapA, snapB, diffs)
    if not picked:
        return False, "No suitable REG_DWORD flip found under FxProperties. Driver may use non-DWORD or a different location."
    # picked contains: {'hive','name','before','after','subkey'}
    value_name = picked["name"]
    dword_enable  = int(picked["before"])  # A = Enabled
    dword_disable = int(picked["after"])   # B = Disabled
    section_name  = _sanitize_ini_section_name(value_name)
    notes = f"Auto-learned on '{name}' ({flow}). A=enabled,B=disabled."
    # Write/update INI (requires write permission; Program Files needs Admin)
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
    # Restore original state (best effort), now vendor-first will also consider the new INI
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
def _build_vendor_ini_snippet(target, snapA, snapB, diffs, section_name=None):
    """
    Build a suggested vendor INI section based on DWORD flips observed between:
      A = Enhancements ENABLED (snapshotA)
      B = Enhancements DISABLED (snapshotB)
    We look for REG_DWORD 0/1 flips under FxProperties with name like "{GUID},pid".
    If multiple candidates, pick the first HKLM FxProperties flip; otherwise first candidate.
    Returns (text_snippet: str | None, picked: dict | None)
    """
    # Prefer flips under FxProperties and HKLM
    cands = []
    for f in diffs.get("dword_flips", []):
        name = str(f.get("name",""))
        subkey = str(f.get("subkey",""))
        hive = str(f.get("hive",""))
        # Must match {GUID},pid pattern
        if not (name.startswith("{") and "}" in name and "," in name):
            continue
        # Must be under FxProperties (where vendor toggles usually live)
        if subkey != "FxProperties":
            continue
        # A=Enabled, B=Disabled
        before = int(f.get("before"))
        after  = int(f.get("after"))
        cands.append({
            "hive": hive, "name": name, "before": before, "after": after, "subkey": subkey
        })
    if not cands:
        return None, None
    # Prefer HKLM first
    cands.sort(key=lambda x: (0 if x["hive"]=="HKLM" else 1))
    pick = cands[0]
    # Determine semantics (DWORD values when enhancements enabled/disabled)
    # Snapshot A (Enabled) -> "before"; Snapshot B (Disabled) -> "after"
    dword_enable = pick["before"]
    dword_disable = pick["after"]
    # Construct a section name if not provided
    if not section_name:
        base = re.sub(r'[^A-Za-z0-9_,\-{}]+', "_", pick["name"])
        section_name = f"vendor_{base}"
    # Build snippet (we’ll prefer HKCU,HKLM in the INI to match HKCU-first reads)
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
def _verify_effect_only(device_id, flow, expected_enabled, timeout=2.5, interval=0.2, consecutive=2):
    """
    Windows-only verification for fallback paths: require PropertyStore Disable_SysFx match expected.
    Returns (ok, verifiedBy, finalState).
    """
    import time
    want = True if expected_enabled else False
    ok_streak = 0
    last_state = None
    end = time.time() + float(timeout)
    while time.time() < end:
        cur = _get_enhancements_status_propstore(device_id)
        src = "windows-live(ps)"
        last_state = cur
        if cur is not None and cur == want:
            ok_streak += 1
            if ok_streak >= consecutive:
                return True, src, cur
        else:
            ok_streak = 0
        time.sleep(interval)
    return False, None, last_state

def _apply_enhancements(device_id, flow, enable, prefer_hklm=False, allow_universal_scan=False, vendor_ini_path=None):
    """
    Vendor-only policy:
      1) Try vendor toggles: INI vendors first (user-learned), then built-in code vendors (e.g. Realtek).
      2) If no vendor match, return failure (no Windows fallback).
    """
    # Vendor first (code -> ini)
    ok_v, tag_v, state_v = _try_vendor_first(device_id, flow, enable, ini_path=vendor_ini_path)
    if ok_v:
        return True, tag_v, state_v
    return False, "no-vendor-method", None

def _short_settle(sec=0.15):
    try:
        time.sleep(float(sec))
    except Exception:
        pass

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
        CLSID_PolicyConfigClient = GUID(_guid_from_parts("294935CE", "-F637-4E7C-", "A41B-", "AB255460B862"))
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
        try:
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
        except Exception:
            pass
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
def cmd_enhancements(args):
    devices = list_devices(include_all=False)  # ACTIVE ONLY
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
    # Learn mode (manual): user flips UI; we capture A/B and write INI
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
    # Normal enable/disable flow (vendor-only; no Windows fallback)
    enable = True if args.enable else False
    # Require vendor availability before toggle
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
    """
    Print live Enhancements state from:
      - COM/PolicyConfig (Disable_SysFx)
      - IPropertyStore (Disable_SysFx)
      - Vendor toggle (e.g., Realtek/Waves {1DA5D803...,5})
    """
    devices = list_devices(include_all=False)
    matches = find_devices_by_selector(devices, dev_id=args.id, name_substr=args.name, flow=args.flow, regex=args.regex)
    if not matches:
        print("ERROR: device not found (active only)", file=sys.stderr); return 3
    buckets = _sort_and_tag_gui_indices([d for d in matches])
    ordered = (buckets.get(args.flow) or []) if args.flow else (buckets["Render"] + buckets["Capture"])
    target = ordered[args.index] if args.index is not None else ordered[0]
    live_win = _get_enhancements_status_propstore(target["id"])
    live_com = _get_enhancements_status_com(target["id"])
    
    # Read-only vendor diagnostic: find first applicable vendor entry (no writes), then read its state
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
    Interactive discovery: capture snapshots with Enhancements Enabled and Disabled,
    diff them, and write a TXT + JSON report, optionally emitting an INI snippet.
    """
    # Resolve target (active only)
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
    # Prompt user to set to Enabled (UI)
    print(f"Discovery target: {target['name']} [{target['id']}] ({target['flow']})")
    print("Step 1: In Windows Sound settings, set 'Audio Enhancements' to ENABLED for this device.")
    input("When ready, press Enter to capture snapshot A... ")
    snapA = _collect_sysfx_snapshot(target["id"])
    # Prompt user to set to Disabled (UI)
    print("Step 2: Now set 'Audio Enhancements' to DISABLED for the same device.")
    input("When ready, press Enter to capture snapshot B... ")
    snapB = _collect_sysfx_snapshot(target["id"])
    # Diff registry portion broadly
    diffs = _diff_mmdevices_lists(snapA.get("registry") or [], snapB.get("registry") or [])
    # Build outputs
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
    
    # Optional: INI snippet for vendor database
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

    # Console summary
    print(report_text)
    print(f"\nSaved:")
    print(f"  TXT  -> {txt_path}")
    print(f"  JSON -> {json_path}")
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
    # enhancements
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
    # diag-sysfx
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
    # diag-mmdevices
    p_dm = sub.add_parser("diag-mmdevices", help="Dump all MMDevices values for an endpoint (debug)")
    p_dm.add_argument("--id")
    p_dm.add_argument("--name")
    p_dm.add_argument("--flow", choices=["Render", "Capture"])
    p_dm.add_argument("--index", type=int)
    p_dm.add_argument("--regex", action="store_true")
    p_dm.set_defaults(func=cmd_diag_mmdevices)
    # discover-enhancements
    p_learn = sub.add_parser("discover-enhancements", help="Interactively learn how Enhancements toggles for a device")
    p_learn.add_argument("--id")
    p_learn.add_argument("--name")
    p_learn.add_argument("--flow", choices=["Render", "Capture"])
    p_learn.add_argument("--index", type=int)
    p_learn.add_argument("--regex", action="store_true")
    p_learn.add_argument("--output-dir", help="Where to write the TXT/JSON report (default: current directory)")
    p_learn.add_argument("--ini-snippet", help="Write a suggested vendor INI section to this path (append).")
    p_learn.set_defaults(func=cmd_discover_enhancements)
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
def cmd_diag_mmdevices(args):
    devices = list_devices(include_all=False)
    matches = find_devices_by_selector(devices, dev_id=args.id, name_substr=args.name, flow=args.flow, regex=args.regex)
    if not matches:
        print("ERROR: device not found (active only)", file=sys.stderr); return 3
    if len(matches) > 1 and args.index is None:
        print("ERROR: multiple matches; specify --index", file=sys.stderr); return 4
    buckets = _sort_and_tag_gui_indices([d for d in matches])
    ordered = (buckets.get(args.flow) or []) if args.flow else (buckets["Render"] + buckets["Capture"])
    if not ordered:
        print("ERROR: no target device found for the specified criteria", file=sys.stderr); return 4
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
# Record where we’sre logging
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
# Define _fh before the try so it exists for atexi close
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
            self.root.title("Audio Control v1.3.0.17 12-22-2025")
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
            # Toggle Enhancements (SysFX) item (render or capture)
            self.enh_menu_default_label = "Enable Enhancements"  # never show "Toggle"
            self.menu.add_command(label=self.enh_menu_default_label, command=self.on_toggle_enhancements)
            self.enh_menu_index = self.menu.index("end")
            
            # Learn Enhancements (manual discovery -> write INI)
            self.menu.add_command(label="Learn Enhancements", command=self.on_learn_enhancements)
            self.learn_menu_index = self.menu.index("end")
            
            self._pending_enh = None  # remembers intended action for the selected device
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
                # Configure the Enhancements (SysFX) entry (Render and Capture) - vendor-only
                vend_available = False
                try:
                    vend_available = bool(_find_first_vendor_entry(d["id"], d["flow"], ini_path=_vendor_ini_default_path()))
                except Exception:
                    vend_available = False

                if vend_available:
                    try:
                        enh = _get_enhancements_status_any(d["id"], d["flow"])
                    except Exception:
                        enh = None
                    if enh is True:
                        enh_label = "Disable Enhancements"
                        target_enable_next = False
                    elif enh is False:
                        enh_label = "Enable Enhancements"
                        target_enable_next = True
                    else:
                        # Unknown vendor state -> pick a safe concrete action
                        enh_label = "Enable Enhancements"
                        target_enable_next = True
                    self._pending_enh = {"id": d["id"], "enable": target_enable_next}
                    self.menu.entryconfig(self.enh_menu_index, label=enh_label, state="normal")
                else:
                    # No vendor method -> disable Enhancements item; Learn remains available
                    self._pending_enh = None
                    self.menu.entryconfig(self.enh_menu_index, label=self.enh_menu_default_label, state="disabled")
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
        def on_toggle_enhancements(self):
            d = self.get_selected_device()
            if not d:
                return
            try:
                _log(f"Enhancements toggle requested for {d['name']} ({d['id']})")
                
                # Determine target state based on current state
                if getattr(self, "_pending_enh", None) and self._pending_enh.get("id") == d["id"]:
                    enable = bool(self._pending_enh["enable"])
                else:
                    current = _get_enhancements_status_any(d["id"], d["flow"])
                    enable = True if current is None else (not bool(current))
                
                # Check if supported via vendor method
                if not _enhancements_supported(d["id"], d["flow"]):
                    messagebox.showinfo("Not supported", "This endpoint does not have a configured vendor toggle for 'Audio Enhancements'. Use 'Learn Enhancements' first.")
                    self.set_status("Enhancements toggle failed: No vendor method.")
                    return

                # Print intended CLI for reproducibility (no --vendor-ini in GUI)
                self.maybe_print_cli(f'audioctl enhancements --id "{d["id"]}" --flow {d["flow"]} --{"enable" if enable else "disable"}')
                
                # Apply using the core logic (GUI uses default INI path, no universal scan)
                ok, verified_by, state = _apply_enhancements(
                    d["id"], d["flow"], enable, 
                    prefer_hklm=is_admin(), 
                    allow_universal_scan=False,
                    vendor_ini_path=_vendor_ini_default_path()
                )
                # Clear remembered action for next menu open
                self._pending_enh = None
                
                if ok and (state is None or state == enable):
                    state_txt = "enabled" if state else "disabled"
                    _log(f"Enhancements toggle result for {d['name']} ({d['id']}): final={state_txt} via {verified_by}")
                    
                    if verified_by.startswith("vendor"):
                        self.set_status(f"Enhancements {state_txt} for: {d['name']} (Vendor controlled)")
                    else:
                        # Should not happen with vendor-only path
                        self.set_status(f"Enhancements {state_txt} for: {d['name']} (Unknown success path)")
                else:
                    from tkinter import messagebox
                    messagebox.showwarning("Could not verify", "Vendor toggle applied but could not verify final state, or the toggle failed.")
                    self.set_status(f"Enhancements toggle requested for: {d['name']} (Verification failed)")
            except Exception as e:
                _log(f"Enhancements toggle exception for {d['name']} ({d['id']}): {e!r}")
                from tkinter import messagebox
                messagebox.showerror("Error", f"Failed to toggle Enhancements:\n{e}")
                self.set_status("Failed to toggle Enhancements")
        def on_learn_enhancements(self):
            d = self.get_selected_device()
            if not d:
                return
            try:
                from tkinter import messagebox
                ini_path = _vendor_ini_default_path()

                # STERN WARNING + explicit OK/Cancel
                warn_txt = (
                    "READ CAREFULLY\n\n"
                    "This Learn mode will capture two registry snapshots and write a vendor entry into:\n"
                    f"  {ini_path}\n\n"
                    "From now on, future 'Enhancements' commands for this device WILL WRITE registry values on this machine "
                    "(HKCU/optional HKLM) to toggle Enhancements. This is persistent until you manually remove the learned section.\n\n"
                    "Critical rules during Learn:\n"
                    "- Do NOT change any other audio settings.\n"
                    "- Do NOT switch default devices.\n"
                    "- Do NOT open other audio/control apps.\n"
                    "- Only toggle 'Audio Enhancements' for THIS device exactly when asked.\n\n"
                    "Click OK to continue, or Cancel to abort."
                )
                if not messagebox.askokcancel("Warning – Learn writes registry (persistent)", warn_txt):
                    self.set_status("Learn Enhancements: aborted by user")
                    return

                # Step 1 (Enabled)
                messagebox.showinfo(
                    "Learn Enhancements - Step 1",
                    "Set 'Audio Enhancements' to ENABLED for this device in Windows Sound settings.\n\nClick OK to capture snapshot A."
                )
                snapA = _collect_sysfx_snapshot(d["id"])

                # Step 2 (Disabled)
                messagebox.showinfo(
                    "Learn Enhancements - Step 2",
                    "Set 'Audio Enhancements' to DISABLED for the same device.\n\nClick OK to capture snapshot B."
                )
                snapB = _collect_sysfx_snapshot(d["id"])

                # Diff and pick candidate
                diffs = _diff_mmdevices_lists(snapA.get("registry") or [], snapB.get("registry") or [])
                snippet, picked = _build_vendor_ini_snippet(d, snapA, snapB, diffs)
                if not picked:
                    messagebox.showwarning("Learn Enhancements", "No suitable REG_DWORD flip found under FxProperties.\nThe driver may use non-DWORD or a different location.")
                    self.set_status("Learn Enhancements: no DWORD flip found")
                    return

                value_name    = picked["name"]
                dword_enable  = int(picked["before"])
                dword_disable = int(picked["after"])
                section_name  = _sanitize_ini_section_name(value_name)
                notes = f"Auto-learned (manual UI) on '{d['name']}' ({d['flow']}). A=enabled,B=disabled."

                # Append-only INI write (HKCU,HKLM preferred)
                try:
                    res = _append_vendor_ini_entry_if_missing(
                        ini_path, section_name, value_name,
                        dword_enable, dword_disable,
                        flows="Render,Capture", hives="HKCU,HKLM", notes=notes
                    )
                    if res == "exists":
                        messagebox.showinfo(
                            "Learn Enhancements",
                            f"Vendor section already exists:\n{ini_path}\n\nSection: [{section_name}]\nNo changes were made."
                        )
                        self.set_status("Learn Enhancements: entry already exists")
                    else:
                        messagebox.showinfo(
                            "Learn Enhancements",
                            f"Learned vendor toggle and appended to:\n{ini_path}\n\nSection: [{section_name}]\nValue: {value_name}\nEnabled={dword_enable}, Disabled={dword_disable}"
                        )
                        self.set_status("Learn Enhancements: vendor INI updated")
                except PermissionError:
                    messagebox.showerror(
                        "Permission denied",
                        f"Could not write INI at:\n{ini_path}\nRun as Administrator or choose a writable location."
                    )
                    self.set_status("Learn Enhancements: permission denied")
                except Exception as e:
                    messagebox.showerror("Error", f"Failed to write INI: {e}")
                    self.set_status("Learn Enhancements: write failed")
            except Exception as e:
                from tkinter import messagebox
                messagebox.showerror("Error", f"Learn Enhancements failed:\n{e}")
                self.set_status("Learn Enhancements: failed")
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
    if args.cmd == "enhancements":
        trio = int(bool(args.enable)) + int(bool(args.disable)) + int(bool(args.learn))
        if trio != 1:
            print("ERROR: specify exactly one of --enable, --disable, or --learn", file=sys.stderr)
            try: CoUninitialize()
            except Exception: pass
            return 1
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