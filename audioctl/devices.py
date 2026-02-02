# audioctl/devices.py
#
# "Engine room" for Windows audio control.
#
# This module owns the low-level Windows integration:
#   - Device enumeration (IMMDeviceEnumerator / MMDevice API)
#   - Default endpoint routing (PolicyConfig / PolicyConfigFx)
#   - Volume & mute (IAudioEndpointVolume)
#   - "Listen to this device" toggle (IPropertyStore) + routing target (registry)
#   - SysFX/"Audio Enhancements" status & control (PolicyConfigFx, IPropertyStore, registry)
#
# Guiding principles for stability (learned the hard way on real systems/drivers):
#   1) Each public helper manages its own COM apartment lifetime via _com_context().
#      The CLI/GUI should not hold a process-wide COM init just to call into this module.
#   2) Avoid persistent COM singletons. COM objects are thread-affine; keeping them around
#      across GUI events + Python GC can cause Release() to run on the wrong thread and
#      crash with intermittent access violations.
#   3) Cache *interface class definitions* (ctypes/comtypes vtables), not COM instances.
#      Dynamically creating ctypes CFUNCTYPE/COMMETHOD vtables while GC runs can trigger
#      comtypes finalizers mid-construction, producing rare but catastrophic crashes.
#
import re
import time
import warnings
import ctypes
import winreg
from ctypes import POINTER, byref, wintypes

# Import compat BEFORE comtypes/pycaw:
# compat.py patches comtypes.automation symbols/constants that PyInstaller builds
# sometimes miss, and it forces import of comtypes._post_coinit modules so their
# Release-time cleanup code is bundled and available at shutdown.
from .compat import (
    E_RENDER, E_CAPTURE,
    E_CONSOLE, E_MULTIMEDIA, E_COMMUNICATIONS,
    ROLES, DEVICE_STATE_ACTIVE, DEVICE_STATE_ALL, DEVICE_STATES,
    STGM_READ, STGM_WRITE, _guid_from_parts,
)

from comtypes import CLSCTX_ALL, CoCreateInstance, GUID, IUnknown, COMMETHOD, HRESULT
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume, IMMDeviceEnumerator
from pycaw.constants import CLSID_MMDeviceEnumerator
from .logging_setup import _log, _log_exc, _dbg

# Removed: from .vendor_db import ...
import comtypes.automation as automation
import copy
import threading
import comtypes

# --- COM apartment lifecycle (thread-local, reference-counted) ---
#
# Windows COM requires a thread to initialize an apartment (CoInitialize) before
# calling most COM APIs. Many helpers in this module can call other helpers
# (nested), so we use a per-thread reference count:
#   - First entry on a thread calls CoInitialize()
#   - Each nested entry increments a counter
#   - Final exit decrements to zero and calls CoUninitialize()
#
# Exceptions are swallowed intentionally. Some environments/drivers are finicky:
# failures here typically mean the caller is already in a COM-initialized context,
# or COM init is not required for a particular code path. Treat as best-effort.
_com_tls = threading.local()

def _com_enter():
    try:
        cnt = getattr(_com_tls, "count", 0)
        if cnt == 0:
            comtypes.CoInitialize()
        _com_tls.count = cnt + 1
    except Exception:
        pass

def _com_exit():
    try:
        cnt = getattr(_com_tls, "count", 0) - 1
        if cnt <= 0:
            _com_tls.count = 0
            try:
                comtypes.CoUninitialize()
            except Exception:
                pass
        else:
            _com_tls.count = cnt
    except Exception:
        pass

from contextlib import contextmanager

@contextmanager
def _com_context():
    # Context manager wrapper so callers can reliably scope COM init/teardown.
    _com_enter()
    try:
        yield
    finally:
        _com_exit()

# --- Cached PolicyConfigFx interface definitions (define once at import time) ---
#
# PolicyConfigFx is an "undocumented-ish" but widely used COM interface family for:
#   - SetDefaultEndpoint()
#   - reading/writing endpoint properties (GetPropertyValue/SetPropertyValue)
#     including FxStore vs normal store toggles (bFxStore flag).
#
# Key stability requirement: define ctypes/comtypes interface classes ONCE.
# Re-defining COMMETHOD() signatures repeatedly at runtime increases the chance
# that GC or comtypes finalizers run while ctypes is building vtables, which has
# historically caused intermittent access violations.
_POLICY_CONFIG_FX_DEFS = None

def _init_policyconfig_fx_defs_once():
    global _POLICY_CONFIG_FX_DEFS
    if _POLICY_CONFIG_FX_DEFS is not None:
        return

    class PROPERTYKEY(ctypes.Structure):
        # PROPERTYKEY = {fmtid (GUID), pid (DWORD)}
        # Windows uses "{fmtid},pid" pairs to address properties in endpoint stores.
        _fields_ = (("fmtid", GUID), ("pid", wintypes.DWORD))

    # Prefer comtypes.automation PROPVARIANT, with a small fallback.
    # PROPVARIANT is the standard Windows "variant value" container for property stores.
    try:
        PROPVARIANT = getattr(automation, "PROPVARIANT", getattr(automation, "tagPROPVARIANT"))
    except Exception:
        class _PVU(ctypes.Union):
            _fields_ = [
                ("boolVal", ctypes.c_short),
                ("uiVal", ctypes.c_ushort),
                ("ulVal", ctypes.c_ulong),
                ("pwszVal", ctypes.c_wchar_p),
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

    # IPolicyConfigFx IID {F8679F50-850A-41CF-9C72-430F290290C8}
    _IID_PolicyConfig = GUID(_guid_from_parts("F8679F50", "-850A-41CF-", "9C72-", "430F290290C8"))

    class IPolicyConfigFx(IUnknown):
        _iid_ = _IID_PolicyConfig
        _methods_ = (
            COMMETHOD([], HRESULT, 'GetMixFormat',
                      (['in'], wintypes.LPCWSTR, 'wszDeviceId'),
                      (['out'], POINTER(ctypes.c_void_p), 'ppFormat')),
            COMMETHOD([], HRESULT, 'GetDeviceFormat',
                      (['in'], wintypes.LPCWSTR, 'wszDeviceId'),
                      (['in'], wintypes.BOOL, 'bDefault'),
                      (['out'], POINTER(ctypes.c_void_p), 'ppFormat')),
            COMMETHOD([], HRESULT, 'SetDeviceFormat',
                      (['in'], wintypes.LPCWSTR, 'wszDeviceId'),
                      (['in'], ctypes.c_void_p, 'pEndpointFormat'),
                      (['in'], ctypes.c_void_p, 'mixFormat')),
            COMMETHOD([], HRESULT, 'GetProcessingPeriod',
                      (['in'], wintypes.LPCWSTR, 'wszDeviceId'),
                      (['in'], wintypes.BOOL, 'bDefault'),
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
            # NOTE: bFxStore variants we need:
            # Some properties exist in both the "FX store" and the normal property store.
            # Drivers vary; we probe/write both for better compatibility.
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

    _POLICY_CONFIG_FX_DEFS = (IPolicyConfigFx, CLSID_PolicyConfigClient, PROPERTYKEY, PROPVARIANT)

# Initialize once at import time to avoid runtime ctypes class construction hazards.
_init_policyconfig_fx_defs_once()

def _define_policyconfig_fx_interfaces():
    # Backward-compatible helper that now just returns the cached defs.
    _init_policyconfig_fx_defs_once()
    return _POLICY_CONFIG_FX_DEFS

# Global cache for PropertyStore interface definitions to avoid GC-related COM crashes
_PROPERTY_STORE_INTERFACES_CACHE = None

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
        import sys
        for line in buf_text.splitlines(True):
            if not line.lstrip().lower().startswith("error:"):
                sys.stderr.write(line)
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

def set_listen_to_device_ps(capture_device_id, enable, render_device_id=None):
    """
    Enable/disable 'Listen to this device' for a Capture endpoint.

    Mechanism:
      - Enable flag (checkbox): IPropertyStore SetValue/Commit on PKEY_LISTEN_ENABLE (pid=1)
      - Routing target (which Render device to play through): registry write of pid=0 string

    Inputs:
      capture_device_id: full MMDevice endpoint ID (Capture)
      enable: bool
      render_device_id:
        - None: keep existing routing
        - "": route to "Default Playback Device" (Windows convention)
        - "<Render endpoint id>": route to a specific playback device

    Side effects / permissions:
      - The checkbox is a per-user property store value; this usually works without admin.
      - The routing value is written under HKLM (machine-wide MMDevices), which may require admin.
        We use HKLM because it is what Windows UI / driver routing actually honors reliably;
        setting routing only through COM often reports success but doesn't apply the target.

    Returns:
      True on best-effort success of the enable toggle.
      False on failure (and prints a single-line "ERROR:" to stderr).
      (Routing write failures are logged as WARNING and do not force False.)
    """
    with _com_context():
        import sys, gc

        # Cached raw interface definitions: building these repeatedly is risky
        # (ctypes vtable construction + GC + comtypes finalizers can crash).
        interfaces = _get_property_store_interfaces()
        PROPVARIANT = interfaces["PROPVARIANT"]
        PROPERTYKEY = interfaces["PROPERTYKEY"]
        IPropertyStoreRaw = interfaces["IPropertyStoreRaw"]
        PIPS = interfaces["PIPS"]
        VT_BOOL = interfaces["VT_BOOL"]
        HRESULT_T = interfaces["HRESULT_T"]
        VARIANT_TRUE = interfaces["VARIANT_TRUE"]
        VARIANT_FALSE = interfaces["VARIANT_FALSE"]

        def _hrx(hr): return f"0x{ctypes.c_uint(hr).value:08X}"
        def _raw_ptr(p): return ctypes.cast(p, ctypes.c_void_p).value

        propsys = ctypes.OleDLL("propsys.dll")
        ole32 = ctypes.OleDLL("ole32.dll")

        have_helpers = True
        try:
            # Prefer Windows helper APIs to initialize PROPVARIANT correctly.
            InitPropVariantFromBoolean = propsys.InitPropVariantFromBoolean
            InitPropVariantFromBoolean.restype = HRESULT_T
            InitPropVariantFromBoolean.argtypes = (wintypes.BOOL, POINTER(PROPVARIANT))
        except (AttributeError, OSError):
            have_helpers = False

        PropVariantClear = ole32.PropVariantClear
        PropVariantClear.restype = HRESULT_T
        PropVariantClear.argtypes = (POINTER(PROPVARIANT),)

        def _pv_from_bool_local(value: bool):
            # Build a PROPVARIANT(bool). Some comtypes builds expose members differently,
            # so we keep this small and defensive.
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

        # PKEY_Listen_Enable:
        #   fmtid = {24dbb0fc-9311-4b3d-9cf0-18ff155639d4}
        #   pid   = 1
        # Semantics: boolean-like "Listen to this device" checkbox state.
        PKEY_LISTEN_ENABLE = PROPERTYKEY(GUID("{24dbb0fc-9311-4b3d-9cf0-18ff155639d4}"), 1)

        pv_enable = None
        try:
            pv_enable = _pv_from_bool_local(bool(enable))

            # --- GC guard ---
            # We disable GC only around the raw vtable pointer usage. The goal is to prevent
            # comtypes finalizers (__del__ -> Release) for unrelated COM objects from running
            # while we hold raw interface pointers. Those races can produce intermittent,
            # hard-to-reproduce access violations.
            gc_was_enabled = gc.isenabled()
            if gc_was_enabled:
                gc.disable()
            try:
                enumerator = CoCreateInstance(CLSID_MMDeviceEnumerator, interface=IMMDeviceEnumerator, clsctx=CLSCTX_ALL)
                dev = enumerator.GetDevice(capture_device_id)
                ps_unknown = dev.OpenPropertyStore(STGM_WRITE)
                ps_ptr_val = _raw_ptr(ps_unknown)
                if not ps_ptr_val:
                    raise OSError("OpenPropertyStore returned null pointer for IPropertyStore.")
                ps_iface = ctypes.cast(ctypes.c_void_p(ps_ptr_val), PIPS)
                hr = ps_iface.contents.lpVtbl.contents.SetValue(ps_iface, byref(PKEY_LISTEN_ENABLE), byref(pv_enable))
                if hr != 0:
                    raise OSError(f"IPropertyStore::SetValue(enable) failed: {_hrx(hr)}")
                hr = ps_iface.contents.lpVtbl.contents.Commit(ps_iface)
                if hr != 0:
                    raise OSError(f"IPropertyStore::Commit failed: {_hrx(hr)}")
            finally:
                if gc_was_enabled:
                    gc.enable()

            # Routing target:
            # Windows stores the playback target for "Listen" as a REG_SZ under HKLM, pid=0.
            # This value is separate from the enable flag, and (in practice) the registry is
            # what the OS/UI/driver honors for the routing target.
            if render_device_id is not None:
                guid = _extract_endpoint_guid_from_device_id(capture_device_id)
                if guid:
                    key_path = rf"SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\Capture\{guid}\Properties"
                    value_name = "{24dbb0fc-9311-4b3d-9cf0-18ff155639d4},0"
                    target_value = render_device_id if render_device_id else ""
                    try:
                        key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key_path, 0, winreg.KEY_SET_VALUE)
                        try:
                            winreg.SetValueEx(key, value_name, 0, winreg.REG_SZ, target_value)
                        finally:
                            winreg.CloseKey(key)
                    except OSError as e:
                        # Routing failure should not mask the checkbox toggle; callers may still
                        # want "Listen" enabled even if routing couldn't be updated without admin.
                        print(f"WARNING: Failed to set playback target (requires Admin): {e}", file=sys.stderr)

            return True
        except Exception as e:
            # CLI captures stderr for this call and uses it to decide whether to fall back
            # to registry verification or emit a user-visible failure. Keep message stable.
            print(f"ERROR: set_listen_to_device_ps failed for '{capture_device_id}': {e}", file=sys.stderr)
            return False
        finally:
            try:
                # Always clear PROPVARIANT memory once we are done with it.
                if pv_enable is not None:
                    PropVariantClear(byref(pv_enable))
            except Exception:
                pass

def _get_listen_to_device_status_ps(device_id):
    """
    Read the 'Listen to this device' enable flag using IPropertyStore::GetValue via raw vtable (ctypes).
    Returns True/False/None.

    Rationale:
      - The PropertyStore value is the authoritative "checkbox" state.
      - COM reads can fail on some drivers/environments; callers should be prepared to fall back
        to a registry probe when this returns None.
    """
    with _com_context():
        import sys, gc

        interfaces = _get_property_store_interfaces()
        PROPVARIANT = interfaces["PROPVARIANT"]
        PROPERTYKEY = interfaces["PROPERTYKEY"]
        PIPS = interfaces["PIPS"]
        VT_BOOL = interfaces["VT_BOOL"]
        HRESULT_T = interfaces["HRESULT_T"]
        VARIANT_FALSE = interfaces["VARIANT_FALSE"]

        ole32 = ctypes.OleDLL("ole32.dll")
        PropVariantClear = ole32.PropVariantClear
        PropVariantClear.restype = HRESULT_T
        PropVariantClear.argtypes = (POINTER(PROPVARIANT),)

        # Same Listen property as in the setter (pid=1).
        PKEY_LISTEN_ENABLE = PROPERTYKEY(GUID("{24dbb0fc-9311-4b3d-9cf0-18ff155639d4}"), 1)

        pv = PROPVARIANT()
        try:
            result = None

            # --- GC guard ---
            # Raw vtable pointer calls are the sensitive part. Disabling GC here avoids
            # finalizer-induced Release() calls running concurrently with our pointer use.
            gc_was_enabled = gc.isenabled()
            if gc_was_enabled:
                gc.disable()
            try:
                enumerator = CoCreateInstance(CLSID_MMDeviceEnumerator, interface=IMMDeviceEnumerator, clsctx=CLSCTX_ALL)
                dev = enumerator.GetDevice(device_id)
                ps_unknown = dev.OpenPropertyStore(STGM_READ)
                ps_ptr_val = ctypes.cast(ps_unknown, ctypes.c_void_p).value
                if not ps_ptr_val:
                    result = None
                else:
                    ps_iface = ctypes.cast(ctypes.c_void_p(ps_ptr_val), PIPS)
                    hr = ps_iface.contents.lpVtbl.contents.GetValue(ps_iface, byref(PKEY_LISTEN_ENABLE), byref(pv))
                    if hr != 0:
                        # Let caller fall back to registry probe.
                        result = None
                    else:
                        if getattr(pv, "vt", 0) == VT_BOOL:
                            try:
                                result = (pv.boolVal != VARIANT_FALSE)
                            except Exception:
                                result = None
                        else:
                            # Unexpected VT; treat as unknown.
                            result = None
            finally:
                if gc_was_enabled:
                    gc.enable()

            return result
        except Exception as e:
            # Keep as WARNING: this is an optional read helper and callers may succeed via registry fallback.
            print(f"WARNING: Failed to read listen status via COM for '{device_id}': {e}", file=sys.stderr)
            return None
        finally:
            try:
                PropVariantClear(byref(pv))
            except Exception:
                pass

def _read_listen_enable_fast(device_id: str):
    """
    Compatibility helper used by CLI/GUI: try COM-backed read first (authoritative),
    then fall back to a robust registry probe. Returns True/False/None.
    """
    state = _get_listen_to_device_status_ps(device_id)
    if state is None:
        state = _read_listen_enable_from_registry(device_id)
    return state

def _read_listen_enable_from_registry(device_id: str):
    r"""
    Robustly read the 'Listen to this device' enable state from MMDevices.

    Why this exists:
      - Some drivers or permission contexts make the COM PropertyStore read unreliable.
      - Registry representations vary by driver: the property may appear as REG_DWORD,
        a REG_BINARY PROPVARIANT blob, or even REG_SZ.

    Where we look:
      HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\Capture\{endpoint-guid}\{FxProperties|Properties}
      for values whose name starts with the Listen fmtid.

    Returns:
      True/False if a value can be interpreted; None if unknown/unavailable.
    """
    try:
        import sys
    except Exception:
        return None

    guid = _extract_endpoint_guid_from_device_id(device_id)
    if not guid:
        return None

    base = r"SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\Capture" + "\\" + guid
    guid_base = "{24dbb0fc-9311-4b3d-9cf0-18ff155639d4}".lower()

    def _parse_bool_from_reg(val, typ):
        # Support common encodings:
        #   - DWORD: 0/1
        #   - BINARY: PROPVARIANT with VT_BOOL payload
        #   - SZ: string booleans used by some vendor layers
        if typ == winreg.REG_DWORD:
            try:
                return bool(int(val))
            except Exception:
                return None
        if typ == winreg.REG_BINARY:
            try:
                b = bytes(val)
                # PROPVARIANT layout: vt at [0:2], bool at [8:10] for VT_BOOL
                if len(b) >= 10:
                    vt = int.from_bytes(b[0:2], "little", signed=False)
                    if vt == 0x000B:
                        bool16 = int.from_bytes(b[8:10], "little", signed=True)
                        return bool16 != 0
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

    preferred = None
    fallback_any = None

    # Search both FxProperties and Properties because drivers differ in where they surface values.
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
                # pid=1 is the canonical "enable" value; prefer it if present.
                if pid == 1:
                    preferred = parsed
                    break
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

    Used when:
      - COM write reported failure, or COM read was ambiguous (None),
        but the change may still have applied (driver timing).
    Returns (verified_ok, last_state).
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

# --- Enhancements Helpers (PropertyStore, Registry, COM helpers) ---

def _get_policy_config_fx():
    IPolicyConfigFx, CLSID_PolicyConfigClient, PROPERTYKEY, PROPVARIANT = _define_policyconfig_fx_interfaces()
    with _com_context():
        return CoCreateInstance(CLSID_PolicyConfigClient, interface=IPolicyConfigFx, clsctx=CLSCTX_ALL)

def _get_policy_config_fx_singleton():
    """
    Create a fresh PolicyConfig object each time - no singleton caching.

    Why:
      COM objects are apartment-thread-affine. Holding a persistent singleton across GUI
      operations allows Python GC to Release() it at arbitrary times / threads, which can
      crash in driver-specific COM implementations. We prefer short-lived instances scoped
      to an operation; interface class definitions remain cached to keep this fast and safe.
    """
    _dbg("Creating PolicyConfigFx COM object (fresh instance, not singleton)")
    try:
        IPolicyConfigFx, CLSID_PolicyConfigClient, PROPERTYKEY, PROPVARIANT = _define_policyconfig_fx_interfaces()
        with _com_context():
            pc = CoCreateInstance(CLSID_PolicyConfigClient, interface=IPolicyConfigFx, clsctx=CLSCTX_ALL)
        try:
            import ctypes
            ptr = ctypes.cast(pc, ctypes.c_void_p).value
            _dbg(f"PolicyConfigFx COM pointer = 0x{ptr:016X}")
        except Exception:
            pass
        return pc
    except Exception as e:
        _dbg(f"Failed to create PolicyConfigFx: {e}")
        return None

# Define PolicyConfig interfaces once at module load to avoid GC issues during dynamic class creation
_POLICY_CONFIG_INTERFACES_CACHE = None

def _get_policy_config_interfaces():
    """
    Get or create PolicyConfig interface definitions once and cache them.

    Motivation:
      pycaw versions differ: some ship PolicyConfig interfaces, others don't.
      We try pycaw first, and fall back to a local interface definition that is
      "good enough" for SetDefaultEndpoint.
    """
    global _POLICY_CONFIG_INTERFACES_CACHE

    if _POLICY_CONFIG_INTERFACES_CACHE is not None:
        return _POLICY_CONFIG_INTERFACES_CACHE

    # Try to import from pycaw first
    try:
        from pycaw.policyconfig import IPolicyConfig, IPolicyConfigVista, CLSID_PolicyConfigClient
        _POLICY_CONFIG_INTERFACES_CACHE = (IPolicyConfig, IPolicyConfigVista, CLSID_PolicyConfigClient)
        return _POLICY_CONFIG_INTERFACES_CACHE
    except Exception:
        pass

    # Define locally if pycaw doesn't have them
    CLSID_PolicyConfigClient = GUID(_guid_from_parts("294935CE", "-F637-4E7C-", "A41B-", "AB255460B862"))

    class IPolicyConfigVista(IUnknown):
        _iid_ = GUID("{568B9108-44BF-40B4-9006-86AFE5B5A620}")
        _methods_ = (
            COMMETHOD([], HRESULT, 'GetMixFormat', (['in'], wintypes.LPCWSTR, 'wszDeviceId'), (['out'], ctypes.POINTER(ctypes.c_void_p), 'ppFormat')),
            COMMETHOD([], HRESULT, 'GetDeviceFormat', (['in'], wintypes.LPCWSTR, 'wszDeviceId'), (['in'], wintypes.BOOL, 'bDefault'), (['out'], ctypes.POINTER(ctypes.c_void_p), 'ppFormat')),
            COMMETHOD([], HRESULT, 'SetDeviceFormat', (['in'], wintypes.LPCWSTR, 'wszDeviceId'), (['in'], ctypes.c_void_p, 'pEndpointFormat'), (['in'], ctypes.c_void_p, 'mixFormat')),
            COMMETHOD([], HRESULT, 'GetProcessingPeriod', (['in'], wintypes.LPCWSTR, 'wszDeviceId'), (['in'], wintypes.BOOL, 'bDefault'), (['out'], ctypes.POINTER(ctypes.c_longlong), 'pmftDefaultPeriod'), (['out'], ctypes.POINTER(ctypes.c_longlong), 'pmftMinimumPeriod')),
            COMMETHOD([], HRESULT, 'SetProcessingPeriod', (['in'], wintypes.LPCWSTR, 'wszDeviceId'), (['in'], ctypes.POINTER(ctypes.c_longlong), 'pmftPeriod')),
            COMMETHOD([], HRESULT, 'GetShareMode', (['in'], wintypes.LPCWSTR, 'wszDeviceId'), (['out'], ctypes.POINTER(ctypes.c_void_p), 'pMode')),
            COMMETHOD([], HRESULT, 'SetShareMode', (['in'], wintypes.LPCWSTR, 'wszDeviceId'), (['in'], ctypes.c_void_p, 'mode')),
            COMMETHOD([], HRESULT, 'GetPropertyValue', (['in'], wintypes.LPCWSTR, 'wszDeviceId'), (['in'], ctypes.POINTER(ctypes.c_void_p), 'key'), (['out'], ctypes.POINTER(ctypes.c_void_p), 'pv')),
            COMMETHOD([], HRESULT, 'SetPropertyValue', (['in'], wintypes.LPCWSTR, 'wszDeviceId'), (['in'], ctypes.POINTER(ctypes.c_void_p), 'key'), (['in'], ctypes.POINTER(ctypes.c_void_p), 'pv')),
            COMMETHOD([], HRESULT, 'SetDefaultEndpoint', (['in'], wintypes.LPCWSTR, 'wszDeviceId'), (['in'], wintypes.DWORD, 'role')),
            COMMETHOD([], HRESULT, 'SetEndpointVisibility', (['in'], wintypes.LPCWSTR, 'wszDeviceId'), (['in'], wintypes.BOOL, 'bVisible')),
        )

    # IPolicyConfig is typically the same as Vista for our purposes
    IPolicyConfig = IPolicyConfigVista

    _POLICY_CONFIG_INTERFACES_CACHE = (IPolicyConfig, IPolicyConfigVista, CLSID_PolicyConfigClient)
    return _POLICY_CONFIG_INTERFACES_CACHE

def _release_singletons_quiet():
    """
    No-op now since we don't keep singletons.
    Kept for backward compatibility in case it's called from cleanup code.
    """
    pass

def _get_property_store_interfaces():
    """
    Get or create IPropertyStore raw vtable interface definitions once and cache them.

    Why raw vtable (ctypes) instead of comtypes wrapper:
      - comtypes "dynamic" interface wrappers can vary across versions and can behave
        differently in frozen builds.
      - Using raw vtable calls gives us direct access to GetValue/SetValue/Commit and
        avoids some comtypes dispatch overhead and edge-case failures.
      - Most importantly, we can tightly control lifetime/GC behavior around these calls.

    Notes:
      - CALL uses WINFUNCTYPE on 32-bit and CFUNCTYPE on 64-bit. The calling convention
        must match the process architecture or calls will crash.
      - We keep VT_* constants around because PROPVARIANT interpretation depends on vt.
    """
    global _PROPERTY_STORE_INTERFACES_CACHE

    if _PROPERTY_STORE_INTERFACES_CACHE is not None:
        return _PROPERTY_STORE_INTERFACES_CACHE

    try:
        HRESULT_T = wintypes.HRESULT
    except Exception:
        HRESULT_T = ctypes.c_long

    # ABI detail:
    # - On 32-bit Windows, COM vtables use stdcall (WINFUNCTYPE).
    # - On 64-bit, the calling convention is effectively unified and CFUNCTYPE works.
    CALL = ctypes.WINFUNCTYPE if ctypes.sizeof(ctypes.c_void_p) == 4 else ctypes.CFUNCTYPE

    # PROPVARIANT: prefer comtypes' definition, fall back to a minimal structure for our needs.
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

    # vtable function prototypes (IUnknown + IPropertyStore members)
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

    _PROPERTY_STORE_INTERFACES_CACHE = {
        "PROPVARIANT": PROPVARIANT,
        "PROPERTYKEY": PROPERTYKEY,
        "IPropertyStoreRaw": IPropertyStoreRaw,
        "PIPS": PIPS,
        "VT_BOOL": VT_BOOL,
        "VT_LPWSTR": VT_LPWSTR,
        "HRESULT_T": HRESULT_T,
        # VARIANT_TRUE/FALSE are used when synthesizing VT_BOOL values.
        "VARIANT_TRUE": -1,
        "VARIANT_FALSE": 0,
    }

    return _PROPERTY_STORE_INTERFACES_CACHE

def _pkey_disable_sysfx():
    # PKEY_AudioEndpoint_Disable_SysFx:
    #   fmtid = {E4870E26-3CC5-4CD2-BA46-CA0A9A70ED04}
    #   pid   = 2
    #
    # Semantics are inverted:
    #   0 -> enhancements ON
    #   1 -> enhancements OFF
    IPolicyConfigFx, CLSID_PolicyConfigClient, PROPERTYKEY, PROPVARIANT = _define_policyconfig_fx_interfaces()
    g = _guid_from_parts("E4870E26", "-3CC5-4CD2-", "BA46-", "CA0A9A70ED04")
    return PROPERTYKEY(GUID(g), wintypes.DWORD(2))

def _parse_boolish_from_propvariant(pv):
    # Helper used for Disable_SysFx and similar values:
    # drivers may store as VT_BOOL, VT_UI2, or VT_UI4.
    VT_BOOL = getattr(automation, "VT_BOOL", 11)
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
    # Mirror of _parse_boolish_from_propvariant: update existing vt when possible
    # so we don't fight the driver's preferred type.
    VT_BOOL = getattr(automation, "VT_BOOL", 11)
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
    Read enhancements enabled/disabled via PolicyConfigFx (COM).

    PolicyConfigFx exposes the Disable_SysFx property via GetPropertyValue.
    We probe both:
      - bFxStore=True   (FX store)
      - bFxStore=False  (normal store)
    because drivers and Windows builds differ in where the authoritative value lives.

    Returns:
      True  => enhancements enabled (Disable_SysFx == 0)
      False => enhancements disabled (Disable_SysFx == 1)
      None  => unknown / read failed
    """
    try:
        with _com_context():
            IPolicyConfigFx, CLSID_PolicyConfigClient, PROPERTYKEY, PROPVARIANT = _define_policyconfig_fx_interfaces()
            pkey = _pkey_disable_sysfx()
            pc = _get_policy_config_fx_singleton()
            if pc is None:
                return None
            for bfx in (True, False):
                pv = PROPVARIANT()
                try:
                    pc.GetPropertyValue(device_id, bfx, byref(pkey), byref(pv))
                    raw = _parse_boolish_from_propvariant(pv)  # Disable_SysFx: 0=enh on, 1=off
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
    Set enhancements via PolicyConfigFx (COM) by writing Disable_SysFx.

    Behavior:
      enable=True  -> Disable_SysFx = 0
      enable=False -> Disable_SysFx = 1

    We attempt to write in both FX and normal stores, because some drivers only honor one.
    Returns True if any write succeeded.
    """
    try:
        with _com_context():
            IPolicyConfigFx, CLSID_PolicyConfigClient, PROPERTYKEY, PROPVARIANT = _define_policyconfig_fx_interfaces()
            pkey = _pkey_disable_sysfx()
            pc = _get_policy_config_fx_singleton()
            if pc is None:
                return False
            desired_disable = 0 if enable else 1
            ok_any = False
            for bfx in (True, False):
                try:
                    pv = PROPVARIANT()
                    # Read current first so we preserve VT if possible.
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

    Where we look:
      HKCU/HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\{Render|Capture}\{guid}\{FxProperties|Properties}

    Why scan both hives and both subkeys:
      - Some drivers persist per-user state under HKCU, others under HKLM.
      - Some place the relevant values under FxProperties, others under Properties.
      - The value name is stored as "{fmtid},pid" (string), which is the serialized PROPERTYKEY.

    Note:
      The registry stores Disable_SysFx (disabled flag), but this function returns "enhancementsEnabled".
    """
    guid = _extract_endpoint_guid_from_device_id(device_id)
    if not guid:
        return None
    fmtid = "{e4870e26-3cc5-4cd2-ba46-ca0a9a70ed04}".lower()

    def _parse_bool_from_reg(val, typ):
        # Registry can store the equivalent of PROPVARIANT in REG_BINARY blobs.
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
    fallback_any = None
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
                        # pid=2 is the canonical Disable_SysFx property.
                        if nl.endswith(",2"):
                            preferred = parsed
                            found_preferred = True
                            break
                        if fallback_any is None:
                            fallback_any = parsed
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
    if fallback_any is not None:
        return False if fallback_any else True
    return None

def _set_enhancements_registry(device_id, enable, prefer_hklm=False):
    """
    Write Disable_SysFx via registry (DWORD 0/1).

    This is primarily used for diagnostic/learning workflows and as an optional fallback path
    (depending on higher-level policy). We try both FxProperties and Properties under both flows
    because drivers are inconsistent, and we try HKCU/HKLM based on preference/permissions.

    Returns True if any write succeeded.
    """
    guid = _extract_endpoint_guid_from_device_id(device_id)
    if not guid:
        return False
    # Value name is the serialized PROPERTYKEY: "{fmtid},pid"
    name = "{e4870e26-3cc5-4cd2-ba46-ca0a9a70ed04},2"
    desired_disable = 0 if enable else 1

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
    # Polling loop used after registry writes to account for driver/UI propagation delay.
    deadline = time.time() + timeout
    last_state = None
    while time.time() < deadline:
        try:
            state = _read_enhancements_from_registry(device_id)
        except Exception:
            state = None
        last_state = state
        if state is not None and state == expected_enabled:
            return True, state
        time.sleep(interval)
    return False, last_state

def _dump_mmdevices_all_values(device_id):
    r"""
    Dump ALL values under BOTH hives for this endpoint.

    Purpose:
      - Discovery/learn tooling needs a complete picture of what changed when the user toggles
        Enhancements/FX in Windows UI.
      - We include 'dataRaw' in addition to 'dataPreview' so the learning system can replay
        vendor-specific binary payloads exactly (REG_BINARY), not just detect that "something changed".

    Output:
      List of records:
        { hive, flow, subkey, name, type, dataPreview, dataRaw }
      where:
        - hive is "HKCU" or "HKLM"
        - flow is "Render" or "Capture"
        - subkey is relative under the endpoint GUID (FxProperties, Properties, and nested subkeys)
        - name is the value name (often "{fmtid},pid")
    """
    guid = _extract_endpoint_guid_from_device_id(device_id)
    if not guid:
        return {"error": "bad endpoint id, cannot extract guid"}
    items = []
    roots = [
        (winreg.HKEY_CURRENT_USER,  "HKCU"),
        (winreg.HKEY_LOCAL_MACHINE, "HKLM"),
    ]

    def _enum_key_recursive(hive, hive_name, root_path, rel_subkey, flow):
        """
        Enumerate values at root_path and recurse into subkeys.

        We recurse because many drivers store effect settings under:
          FxProperties\{plugin-guid}\User\...
        and the learn flow needs to see those too.
        """
        # Enumerate values in current key
        try:
            key = winreg.OpenKey(hive, root_path, 0, winreg.KEY_READ)
        except OSError:
            return
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
                    "subkey": rel_subkey,    # relative path under endpoint GUID
                    "name": name,
                    "type": typ,
                }
                # dataPreview (human-friendly)
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

                # dataRaw (exact payload for learn/replay)
                try:
                    if typ == winreg.REG_DWORD:
                        rec["dataRaw"] = int(val)
                    elif typ == winreg.REG_SZ:
                        rec["dataRaw"] = str(val)
                    elif typ == winreg.REG_BINARY:
                        rec["dataRaw"] = bytes(val).hex()
                    else:
                        rec["dataRaw"] = None
                except Exception:
                    rec["dataRaw"] = None

                items.append(rec)
        finally:
            try:
                winreg.CloseKey(key)
            except Exception:
                pass

        # Recurse into subkeys
        try:
            key = winreg.OpenKey(hive, root_path, 0, winreg.KEY_READ)
        except OSError:
            return
        try:
            i = 0
            while True:
                try:
                    subname = winreg.EnumKey(key, i)
                    i += 1
                except OSError:
                    break
                next_rel = rel_subkey + "\\" + subname if rel_subkey else subname
                next_path = root_path + "\\" + subname
                _enum_key_recursive(hive, hive_name, next_path, next_rel, flow)
        finally:
            try:
                winreg.CloseKey(key)
            except Exception:
                pass

    for hive, hive_name in roots:
        for flow in ("Render", "Capture"):
            for first in ("FxProperties", "Properties"):
                base = rf"SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\{flow}\{guid}\{first}"
                _enum_key_recursive(hive, hive_name, base, first, flow)

    return items


def _mmdev_key_of(rec):
    # Identity key used for diffing registry snapshots.
    # We include hive/flow/subkey/name so we can detect changes across both HKCU/HKLM and both flows.
    return f"{rec.get('hive','?')}|{rec.get('flow','?')}|{rec.get('subkey','?')}|{rec.get('name','?')}"

def _normalize_preview(v):
    try:
        return (v if isinstance(v, (int, float)) else str(v)).strip() if isinstance(v, str) else v
    except Exception:
        return v

def _diff_mmdevices_lists(before_list, after_list):
    """
    Diff two MMDevices registry dumps (from _dump_mmdevices_all_values).

    Output fields:
      - added/removed/changed: record-level differences (by hive|flow|subkey|name)
      - dword_flips: specifically REG_DWORD values that flipped 0<->1 (strong toggle candidates)
      - disable_sysfx_hits: entries whose name starts with the Disable_SysFx fmtid

    The 'dword_flips' list is used by learning/discovery to propose an INI toggle:
    simple DWORD flips are the easiest to reproduce reliably.
    """
    idxA = {_mmdev_key_of(e): e for e in (before_list or [])}
    idxB = {_mmdev_key_of(e): e for e in (after_list or [])}
    all_keys = set(idxA.keys()) | set(idxB.keys())
    added = []
    removed = []
    changed = []
    flips = []
    hits = []
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

            # 0<->1 DWORD flips are the primary learn signal we look for.
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

def _get_enhancements_status_propstore(device_id):
    """
    Read Disable_SysFx directly from the endpoint's IPropertyStore.

    This path is used primarily for diagnostics and learning:
      - It reflects the live property store state even if registry is lagging or virtualized.
      - It does not depend on vendor INI toggles.

    GC is disabled only around raw vtable usage to prevent intermittent Release() races.
    """
    import sys, gc
    try:
        if not sys.platform.startswith("win"):
            return None
        with _com_context():
            interfaces = _get_property_store_interfaces()
            PROPVARIANT = interfaces["PROPVARIANT"]
            PROPERTYKEY = interfaces["PROPERTYKEY"]
            PIPS = interfaces["PIPS"]
            HRESULT_T = interfaces["HRESULT_T"]

            # Disable_SysFx pid=2 (0=enh ON, 1=OFF)
            pkey = PROPERTYKEY(GUID("{E4870E26-3CC5-4CD2-BA46-CA0A9A70ED04}"), wintypes.DWORD(2))
            pv = PROPVARIANT()
            result = None

            gc_was_enabled = gc.isenabled()
            if gc_was_enabled:
                gc.disable()
            try:
                enumerator = CoCreateInstance(CLSID_MMDeviceEnumerator, interface=IMMDeviceEnumerator, clsctx=CLSCTX_ALL)
                dev = enumerator.GetDevice(device_id)
                ps_unknown = dev.OpenPropertyStore(STGM_READ)
                ps_ptr_val = ctypes.cast(ps_unknown, ctypes.c_void_p).value
                if not ps_ptr_val:
                    result = None
                else:
                    ps_iface = ctypes.cast(ctypes.c_void_p(ps_ptr_val), PIPS)
                    hr = ps_iface.contents.lpVtbl.contents.GetValue(ps_iface, byref(pkey), byref(pv))
                    if hr == 0:
                        raw = _parse_boolish_from_propvariant(pv)  # 0 = enh ON, 1 = OFF
                        if raw is None:
                            result = None
                        else:
                            result = (False if raw == 1 else True)
                    else:
                        result = None
            finally:
                if gc_was_enabled:
                    gc.enable()

            # Clear PROPVARIANT after GC is re-enabled.
            try:
                ole32 = ctypes.OleDLL("ole32.dll")
                PropVariantClear = getattr(ole32, "PropVariantClear", None)
                if PropVariantClear:
                    PropVariantClear.restype = HRESULT_T
                    PropVariantClear.argtypes = (ctypes.POINTER(PROPVARIANT),)
                    PropVariantClear(byref(pv))
            except Exception:
                pass

            return result
    except Exception:
        return None

def _set_enhancements_propstore(device_id, enable):
    """
    Write Disable_SysFx directly via IPropertyStore::SetValue + Commit.

    Like _get_enhancements_status_propstore, this is mainly for diagnostics/learning.
    It is a direct Windows mechanism (not vendor-specific). Drivers may or may not honor it.

    Returns True if the property store commit succeeded.
    """
    import sys, gc
    try:
        if not sys.platform.startswith("win"):
            return False
        with _com_context():
            interfaces = _get_property_store_interfaces()
            PROPVARIANT = interfaces["PROPVARIANT"]
            PROPERTYKEY = interfaces["PROPERTYKEY"]
            PIPS = interfaces["PIPS"]
            HRESULT_T = interfaces["HRESULT_T"]

            pkey = PROPERTYKEY(GUID("{E4870E26-3CC5-4CD2-BA46-CA0A9A70ED04}"), wintypes.DWORD(2))
            desired_disable = 0 if enable else 1
            pv = PROPVARIANT()
            ok = False

            gc_was_enabled = gc.isenabled()
            if gc_was_enabled:
                gc.disable()
            try:
                enumerator = CoCreateInstance(CLSID_MMDeviceEnumerator, interface=IMMDeviceEnumerator, clsctx=CLSCTX_ALL)
                dev = enumerator.GetDevice(device_id)
                ps_unknown = dev.OpenPropertyStore(STGM_WRITE)
                ps_ptr_val = ctypes.cast(ps_unknown, ctypes.c_void_p).value
                if not ps_ptr_val:
                    ok = False
                else:
                    ps_iface = ctypes.cast(ctypes.c_void_p(ps_ptr_val), PIPS)
                    # Try to read the existing pv to preserve VT where possible.
                    try:
                        ps_iface.contents.lpVtbl.contents.GetValue(ps_iface, byref(pkey), byref(pv))
                        if not _set_boolish_in_propvariant(pv, desired_disable):
                            pv.vt = 19  # VT_UI4
                            try:
                                pv.ulVal = desired_disable
                            except Exception:
                                pass
                    except Exception:
                        # Build a fresh PV if GetValue failed.
                        pv = PROPVARIANT()
                        pv.vt = 19  # VT_UI4
                        try:
                            pv.ulVal = desired_disable
                        except Exception:
                            pass
                    hr = ps_iface.contents.lpVtbl.contents.SetValue(ps_iface, byref(pkey), byref(pv))
                    if hr == 0:
                        hr = ps_iface.contents.lpVtbl.contents.Commit(ps_iface)
                        ok = (hr == 0)
                    else:
                        ok = False
            finally:
                if gc_was_enabled:
                    gc.enable()

            # Clear PROPVARIANT after GC is re-enabled.
            try:
                ole32 = ctypes.OleDLL("ole32.dll")
                PropVariantClear = getattr(ole32, "PropVariantClear", None)
                if PropVariantClear:
                    PropVariantClear.restype = HRESULT_T
                    PropVariantClear.argtypes = (ctypes.POINTER(PROPVARIANT),)
                    PropVariantClear(byref(pv))
            except Exception:
                pass

            return ok
    except Exception:
        return False

def _wait_for_propstore_sysfx(device_id, expected_enabled, timeout=1.5, interval=0.12):
    """
    Poll the endpoint's IPropertyStore for Disable_SysFx until it matches expected_enabled or timeout.

    Used to handle driver/UI propagation lag during diagnostic/learning flows.
    """
    last = None
    end = time.time() + float(timeout)
    while time.time() < end:
        state = _get_enhancements_status_propstore(device_id)
        last = state
        if state is not None and state == expected_enabled:
            return True, state
        time.sleep(interval)
    return False, last

def _collect_sysfx_snapshot(device_id):
    """
    Collect a full snapshot for discovery/learning of Enhancements behavior.

    Contents:
      - time: timestamp
      - com: PolicyConfigFx reads (fxStore and normalStore)
      - propStore: direct IPropertyStore read
      - registry: full MMDevices dump (HKCU/HKLM; FxProperties/Properties; recursive)

    This is consumed by:
      - discover-enhancements (reporting)
      - learn flows (to identify stable toggle keys)
    """
    import datetime
    snap = {
        "time": datetime.datetime.now().isoformat(timespec="seconds"),
        "com": {},
        "propStore": {},
        "registry": [],
    }

    # COM (both stores). Guard GC to reduce Release() races while COM objects are in play.
    try:
        import gc
        gc_was_enabled = gc.isenabled()
        if gc_was_enabled:
            gc.disable()
        try:
            with _com_context():
                IPolicyConfigFx, CLSID_PolicyConfigClient, PROPERTYKEY, PROPVARIANT = _define_policyconfig_fx_interfaces()
                pkey = _pkey_disable_sysfx()
                pc = _get_policy_config_fx()
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
                del pc
        finally:
            if gc_was_enabled:
                gc.enable()
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

def _generate_enh_discovery_report(target, snapA, snapB, diffs):
    """
    Build a human-readable report from discovery snapshots (A=enabled, B=disabled).

    Used by `discover-enhancements` to:
      - show COM vs PropertyStore vs registry behavior in one place
      - highlight candidate DWORD flips for vendor learn/INI snippets
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

    lines.append("COM (PolicyConfig) - Disable_SysFx (0=Enh ON, 1=OFF)")
    for label in ("fxStore", "normalStore"):
        A = snapA.get("com", {}).get(label, {})
        B = snapB.get("com", {}).get(label, {})
        lines.append(f"  {label:12} A: rawDisable={A.get('rawDisable')} -> enhEnabled={_fmt_bool(A.get('enhEnabled'))} "
                     f"| B: rawDisable={B.get('rawDisable')} -> enhEnabled={_fmt_bool(B.get('enhEnabled'))}")

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

    ds_hits = [e for e in diffs.get("changed", []) if str(e.get('name','')).lower().startswith("{e4870e26-3cc5-4cd2-ba46-ca0a9a70ed04}")]
    if ds_hits:
        lines.append("Disable_SysFx registry entries that changed:")
        for e in ds_hits:
            lines.append(f"  {e.get('hive')}\\{e.get('flow')}\\{e.get('subkey')}\\{e.get('name')} "
                         f"{e.get('dataPreview')} -> {e.get('dataPreviewAfter')} (type {e.get('type')} -> {e.get('typeAfter')})")
        lines.append("")

    flips = diffs.get("dword_flips", [])
    if flips:
        lines.append("Candidate toggle keys (REG_DWORD flips 0<->1):")
        for f in flips:
            lines.append(f"  {f['hive']}\\{f['flow']}\\{f['subkey']}\\{f['name']}  {f['before']} -> {f['after']}")
        lines.append("")
    else:
        lines.append("No DWORD 0/1 flips detected. Vendor may use non-DWORD or a different location.")
        lines.append("")

    lines.append("Notes:")
    lines.append("- If COM/PropertyStore show A!=B, Windows honored Disable_SysFx and the existing setter is correct.")
    lines.append("- If COM/PropertyStore stay the same but a vendor REG_DWORD flips, that key is likely the real toggle.")
    lines.append("- If only REG_BINARY blobs changed, we may need to write that vendor-specific property.")
    lines.append("")
    return "\n".join(lines)

def _get_policy_config():
    """
    Obtain a PolicyConfig COM interface that supports SetDefaultEndpoint.

    We probe AudioUtilities helpers first because pycaw versions differ in how they expose PolicyConfig.
    If not available, fall back to our cached interface definitions (Vista variant is widely compatible).
    """
    # 1) Try any helper exposed by the installed pycaw AudioUtilities
    for name in ("GetPolicyConfig", "_get_policy_config", "get_policy_config"):
        try:
            getter = getattr(AudioUtilities, name, None)
            if getter:
                with _com_context():
                    return getter()
        except Exception:
            pass

    # 2) Use cached interface definitions (either from pycaw or our fallback)
    try:
        IPolicyConfig, IPolicyConfigVista, CLSID_PolicyConfigClient = _get_policy_config_interfaces()
        try:
            with _com_context():
                return CoCreateInstance(CLSID_PolicyConfigClient, interface=IPolicyConfig, clsctx=CLSCTX_ALL)
        except Exception:
            with _com_context():
                return CoCreateInstance(CLSID_PolicyConfigClient, interface=IPolicyConfigVista, clsctx=CLSCTX_ALL)
    except Exception as e:
        raise AttributeError("Audio policy config interface not available in this environment") from e

def set_default_endpoint(device_id, role):
    """
    Set the default audio endpoint for one or more Windows "roles".

    Inputs:
      device_id: exact endpoint ID (must be active)
      role: one of ROLES keys ("console", "multimedia", "communications") or "all"

    Safety:
      We refuse to act on inactive endpoints (disabled/unplugged/notpresent) to avoid
      setting defaults to devices Windows won't route to reliably.

    Failure modes:
      Raises RuntimeError if the device is inactive or if PolicyConfig calls fail.
    """
    _dbg(f"SetDefaultEndpoint start: id={device_id} role={role}")
    if not _is_device_active(device_id):
        _dbg("SetDefaultEndpoint abort: device not active")
        raise RuntimeError("Target device is not active; refusing to set default.")
    with _com_context():
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
                _dbg(f"SetDefaultEndpoint failed for some roles: {results}")
                details = ", ".join([f"{k}={'ok' if v else 'fail'}" for k, v in results.items()])
                raise RuntimeError(f"SetDefaultEndpoint failed for roles: {details}. Underlying error: {last_err}")
        else:
            policy.SetDefaultEndpoint(device_id, ROLES[role])
    _dbg("SetDefaultEndpoint done")

def _is_device_active(device_id):
    """
    Verify the endpoint is currently active (DEVICE_STATE_ACTIVE).

    Why we do this:
      Some Windows builds/drivers will accept SetDefaultEndpoint for inactive devices,
      but the system won't actually route audio correctly. Enforcing active-only makes
      automation predictable and avoids "default device is set to unplugged device" confusion.
    """
    with _com_context():
        for flow in (E_RENDER, E_CAPTURE):
            try:
                _, coll = enum_endpoints(flow, DEVICE_STATE_ACTIVE)
                for i in range(coll.GetCount()):
                    if coll.Item(i).GetId() == device_id:
                        return True
            except Exception:
                pass
    return False

def enum_endpoints(flow, state_mask):
    with _com_context():
        enumerator = CoCreateInstance(CLSID_MMDeviceEnumerator, IMMDeviceEnumerator, CLSCTX_ALL)
        collection = enumerator.EnumAudioEndpoints(flow, state_mask)
        return enumerator, collection

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
                defaults[flow_name][role_name] = None
    return defaults

def _friendly_names_by_id():
    """
    Build {device_id: FriendlyName} using pycaw objects.

    Why we prefer this approach:
      - It avoids the raw IPropertyStore path for every device enumeration, which is the most
        crash-prone area (raw pointers + COM finalizers).
      - AudioUtilities.GetAllDevices() is generally stable and gives us friendly names cheaply.
      - We keep the hardened PropertyStore fallback (_safe_friendly_name_from_device) available
        for environments where pycaw does not provide FriendlyName reliably.
    """
    names = {}
    try:
        with _com_context():
            for dev in AudioUtilities.GetAllDevices():
                try:
                    dev_id = getattr(dev, "id", None) or dev.GetId()
                except Exception:
                    continue
                try:
                    fn = getattr(dev, "FriendlyName", None)
                except Exception:
                    fn = None
                if dev_id and fn:
                    names[dev_id] = fn
    except Exception:
        pass
    return names

def _safe_friendly_name_from_device(dev):
    """
    Read PKEY_Device_FriendlyName from an IMMDevice via IPropertyStore using cached interfaces.

    This is the "hard mode" name reader. It exists because:
      - Some environments/drivers expose incomplete FriendlyName via pycaw.
      - We still need a stable fallback to show users meaningful device names.

    Stability measures:
      - We disable GC while using raw vtable pointers to avoid comtypes __del__/Release races.
      - We explicitly AddRef/Release on the property store COM object while we read properties.
        This keeps the COM refcount balanced even if Python GC runs later.
      - We PropVariantClear() every PROPVARIANT we touch to avoid leaking memory from pwszVal.
      - comtypes can expose pwszVal as a Python string or as a pointer depending on build;
        we try multiple access paths to interpret it safely.
    """
    _dbg("FriendlyName: enter (_safe_friendly_name_from_device)")
    try:
        import sys, gc
        if not sys.platform.startswith("win"):
            return None

        interfaces = _get_property_store_interfaces()
        PROPVARIANT = interfaces["PROPVARIANT"]
        PROPERTYKEY = interfaces["PROPERTYKEY"]
        PIPS = interfaces["PIPS"]
        VT_LPWSTR = interfaces["VT_LPWSTR"]
        HRESULT_T = interfaces["HRESULT_T"]

        # Pause GC so comtypes finalizers don't Release() something while we hold raw pointers.
        gc_was_enabled = gc.isenabled()
        if gc_was_enabled:
            try:
                gc.disable()
            except Exception:
                gc_was_enabled = False
        try:
            ps_unknown = dev.OpenPropertyStore(STGM_READ)

            # Balance COM refcount explicitly while we use the raw vtable.
            # This is defensive: even if GC runs later, we keep lifetime predictable here.
            did_addref = False
            try:
                try:
                    ps_unknown.AddRef()
                    did_addref = True
                except Exception:
                    pass

                ps_ptr_val = ctypes.cast(ps_unknown, ctypes.c_void_p).value
                if not ps_ptr_val:
                    return None

                _dbg(f"FriendlyName: IPropertyStore raw=0x{ps_ptr_val:016X} (AddRef before use)")

                ps_iface = ctypes.cast(ctypes.c_void_p(ps_ptr_val), PIPS)

                # PKEY_Device_FriendlyName fmtid=a45c254e-df1c-4efd-8020-67d146a850e0 pid=14
                PKEY_Device_FriendlyName = PROPERTYKEY(GUID("{a45c254e-df1c-4efd-8020-67d146a850e0}"), 14)
                # PKEY_Device_DeviceDesc pid=2 is a common fallback description string
                PKEY_Device_DeviceDesc   = PROPERTYKEY(GUID("{a45c254e-df1c-4efd-8020-67d146a850e0}"), 2)

                ole32 = ctypes.OleDLL("ole32.dll")
                PropVariantClear = ole32.PropVariantClear
                PropVariantClear.restype = HRESULT_T
                PropVariantClear.argtypes = (ctypes.POINTER(PROPVARIANT),)

                def _read_ptr_or_str(val):
                    # comtypes may give us either:
                    #   - a Python str (already marshaled)
                    #   - a pointer address (needs wstring_at)
                    if isinstance(val, str):
                        return val
                    if val:
                        try:
                            return ctypes.wstring_at(val)
                        except Exception:
                            return None
                    return None

                def _pv_read_lpwstr(pv):
                    # Try common layouts: pv.pwszVal, pv.value.pwszVal, pv.data.pwszVal
                    try:
                        s = _read_ptr_or_str(getattr(pv, "pwszVal", None))
                        if s: return s
                    except Exception: pass
                    try:
                        val = getattr(pv, "value", None)
                        if val is not None:
                            s = _read_ptr_or_str(getattr(val, "pwszVal", None))
                            if s: return s
                    except Exception: pass
                    try:
                        data = getattr(pv, "data", None)
                        if data is not None:
                            s = _read_ptr_or_str(getattr(data, "pwszVal", None))
                            if s: return s
                    except Exception: pass
                    return None

                def _get_string_prop(pkey):
                    pv = PROPVARIANT()
                    try:
                        hr = ps_iface.contents.lpVtbl.contents.GetValue(ps_iface, byref(pkey), byref(pv))
                        if hr == 0 and getattr(pv, "vt", 0) == VT_LPWSTR:
                            s = _pv_read_lpwstr(pv)
                            if s:
                                return s.strip("\x00 ").strip()
                    finally:
                        try:
                            PropVariantClear(byref(pv))
                        except Exception:
                            pass
                    return None

                name = _get_string_prop(PKEY_Device_FriendlyName)
                if not name:
                    name = _get_string_prop(PKEY_Device_DeviceDesc)
                if name:
                    _dbg(f"FriendlyName: got='{name}'")
                    return name
            finally:
                # Balance the AddRef above so the refcount is correct no matter when GC runs.
                if did_addref:
                    try:
                        ps_unknown.Release()
                    except Exception:
                        pass
        except Exception:
            pass
        finally:
            # Re-enable GC if it was enabled initially.
            if gc_was_enabled:
                try:
                    gc.enable()
                except Exception:
                    pass
            _dbg("FriendlyName: leave (released, GC re-enabled)")
    except Exception:
        pass

    # Fallbacks: ID or None (prefer returning something stable for display).
    try:
        return dev.GetId()
    except Exception:
        return None

def list_devices(include_all=False):
    """
    Enumerate playback (Render) and recording (Capture) endpoints.

    Output:
      List[dict] each with:
        id, name, flow, state, isDefault{console,multimedia,communications}

    Notes:
      - This returns an unsorted list. GUI-order sorting and per-flow indices are handled
        in higher-level helpers (_sort_and_tag_gui_indices in devices.py, invoked by CLI/GUI).
      - Friendly names are gathered via a pycaw GetAllDevices() map first for stability.
    """
    _dbg(f"list_devices: include_all={include_all}")
    with _com_context():
        with warnings.catch_warnings():
            warnings.simplefilter("ignore", UserWarning)

            enumerator_for_defaults = CoCreateInstance(CLSID_MMDeviceEnumerator, IMMDeviceEnumerator, CLSCTX_ALL)
            defaults = get_default_ids(enumerator_for_defaults)
            state_mask = DEVICE_STATE_ALL if include_all else DEVICE_STATE_ACTIVE

            # Prefer a single "safe" pass for names rather than raw property reads per device.
            name_map = _friendly_names_by_id()

            out = []

            for flow_name, flow in [("Render", E_RENDER), ("Capture", E_CAPTURE)]:
                _dbg(f"Enum flow={flow_name}")
                enumerator, coll = enum_endpoints(flow, state_mask)
                for i in range(coll.GetCount()):
                    dev = coll.Item(i)
                    dev_id = dev.GetId()

                    # Name map lookup avoids raw IPropertyStore reads in the hot path.
                    name = name_map.get(dev_id) or dev_id

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
            _dbg(f"list_devices: total={len(out)}")
            return out

def find_devices_by_selector(devices, dev_id=None, name_substr=None, flow=None, regex=False):
    """
    Returns list of devices matching selector.
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

    This is used across CLI and GUI to ensure --index behaves consistently.
    """
    buckets = {"Render": [], "Capture": []}
    for d in devices:
        buckets[d["flow"]].append(d)

    for flow in buckets:
        buckets[flow].sort(key=lambda x: x["name"].lower())
        for i, d in enumerate(buckets[flow]):
            d["guiIndex"] = i

    return buckets

def _pretty_matches_msg(label, matches):
    """
    Print a small list of candidates in GUI order to help the user pick.
    """
    buckets = _sort_and_tag_gui_indices(matches[:])
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

    candidates = [d for d in active_devices if d["flow"] == flow_name]

    buckets = _sort_and_tag_gui_indices(candidates)
    ordered = buckets[flow_name]

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

    if index is not None:
        for d in ordered:
            if d.get("guiIndex") == index:
                return d, None
        return None, f"ERROR: --index {index} does not match any active {label} device in GUI order."

    return ordered[0], None

def set_endpoint_mute(device_id, mute_state):
    """
    Set endpoint mute state via IAudioEndpointVolume.

    Returns True on success, False on failure (device not active or COM call failed).
    """
    with _com_context():
        for flow in (E_RENDER, E_CAPTURE):
            _, coll = enum_endpoints(flow, DEVICE_STATE_ACTIVE)
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
    """
    Read endpoint mute state via IAudioEndpointVolume.

    Note:
      comtypes/pycaw can return values either as a plain scalar or as a one-item tuple.
      We normalize to bool or None.
    """
    with _com_context():
        for flow in (E_RENDER, E_CAPTURE):
            _, coll = enum_endpoints(flow, DEVICE_STATE_ACTIVE)
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
                        # Some COM wrappers use out-parameters; try that shape as a fallback.
                        try:
                            from ctypes import wintypes
                            b = wintypes.BOOL()
                            vol.GetMute(ctypes.byref(b))
                            return bool(b.value)
                        except Exception:
                            return None
    return None

def get_endpoint_volume(device_id):
    """
    Read endpoint master volume scalar via IAudioEndpointVolume.

    Returns:
      0..100 integer percentage, or None if unavailable.

    Normalization:
      Windows returns scalar float (0.0..1.0). We clamp to [0,100] and round to int.
      Like mute, some wrappers return a tuple or require an out-parameter.
    """
    with _com_context():
        for flow in (E_RENDER, E_CAPTURE):
            _, coll = enum_endpoints(flow, DEVICE_STATE_ACTIVE)
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
    return None

def set_endpoint_volume(device_id, level_percent):
    """
    Set endpoint master volume via IAudioEndpointVolume.

    Input:
      level_percent: 0..100 (values outside are clamped)
    Returns:
      True on success, False on failure.
    """
    level = max(0.0, min(1.0, float(level_percent) / 100.0))
    with _com_context():
        for flow in (E_RENDER, E_CAPTURE):
            _, coll = enum_endpoints(flow, DEVICE_STATE_ACTIVE)
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

def _verify_effect_only(device_id, flow, expected_enabled, timeout=2.5, interval=0.2, consecutive=2):
    """
    Windows-only verification for fallback paths: require PropertyStore Disable_SysFx match expected.

    This helper is intentionally strict and doesn't interpret vendor toggles; it's used when
    a "Windows path" toggle was attempted and we need to confirm the live property store reflects it.

    Returns:
      (ok, verifiedBy, finalState)
    where:
      verifiedBy is a short label used by higher-level JSON outputs (e.g., "windows-live(ps)")
      to explain which mechanism observed the final state.
    """
    import time as _time
    want = True if expected_enabled else False
    ok_streak = 0
    last_state = None
    end = _time.time() + float(timeout)
    while _time.time() < end:
        cur = _get_enhancements_status_propstore(device_id)
        src = "windows-live(ps)"
        last_state = cur
        if cur is not None and cur == want:
            ok_streak += 1
            if ok_streak >= consecutive:
                return True, src, cur
        else:
            ok_streak = 0
        _time.sleep(interval)
    return False, None, last_state
