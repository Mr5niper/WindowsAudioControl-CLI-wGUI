# audioctl/devices.py
#
# This module is the "engine room" for Windows audio control.
# It contains the low-level COM + registry plumbing that the CLI and GUI build on.
#
# Responsibilities (high-level):
# - Enumerate audio endpoints (Render/Playback and Capture/Recording).
# - Set default endpoints via PolicyConfig (undocumented-ish Windows COM API).
# - Read/set master volume and mute via IAudioEndpointVolume.
# - Toggle "Listen to this device" for Capture endpoints:
#     * enable flag via IPropertyStore (COM PropertyStore)
#     * routing target via MMDevices registry (HKLM) because Windows honors that for routing.
# - Read/set SysFX/Enhancements state via PolicyConfigFx and endpoint PropertyStore
#   (primarily for diagnostics and learning flows).
#
# Guiding stability principles (why this code looks "paranoid"):
# - Every helper that touches COM manages its own COM init/teardown via _com_context().
#   The CLI does NOT do a global CoInitialize anymore.
# - Avoid persistent COM singletons: COM objects have thread affinity; keeping them alive
#   across GUI events + GC cycles can trigger intermittent access violations at shutdown
#   when comtypes finalizers call Release() on the wrong thread.
# - Cache *interface class definitions* (ctypes vtable layouts / comtypes interface classes),
#   not COM objects. Dynamic class creation (COMMETHOD/CFUNCTYPE) is GC-sensitive; caching
#   avoids "GC running during vtable construction" crash patterns.

import re
import time
import warnings
import ctypes
import winreg
from ctypes import POINTER, byref, wintypes

# Import compat BEFORE comtypes/pycaw:
# The compat module applies global comtypes.automation shims that must be in place
# before we import modules that use PROPVARIANT/VT_* constants.
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

# ---- COM lifecycle management ------------------------------------------------
# COM must be initialized per-thread. We also want to support nested calls:
# e.g. cmd_get_device_state -> get_endpoint_volume() + get_endpoint_mute() etc.
# If each helper called CoInitialize/CoUninitialize without reference counting,
# nested calls could uninitialize COM while still in use.
#
# We therefore keep a thread-local refcount. First enter initializes COM,
# last exit uninitializes COM.
#
# Note: we deliberately swallow exceptions here. In some environments (frozen exe,
# unusual COM configuration, shutdown edge cases), CoInitialize/CoUninitialize can
# raise. We treat COM init as best-effort and let the caller's COM operation be the
# real success/failure signal.
_com_tls = threading.local()

def _com_enter():
    try:
        cnt = getattr(_com_tls, "count", 0)
        if cnt == 0:
            comtypes.CoInitialize()
        _com_tls.count = cnt + 1
    except Exception:
        # Best-effort: avoid crashing the app just because COM init failed.
        pass

def _com_exit():
    try:
        cnt = getattr(_com_tls, "count", 0) - 1
        if cnt <= 0:
            _com_tls.count = 0
            try:
                comtypes.CoUninitialize()
            except Exception:
                # Best-effort: teardown failures are non-fatal; COM may already be down.
                pass
        else:
            _com_tls.count = cnt
    except Exception:
        pass

from contextlib import contextmanager

@contextmanager
def _com_context():
    """
    Context manager wrapper for thread-local COM apartment lifetime.

    Why:
    - COM must be initialized once per thread.
    - Many helpers call each other (nested), so we use a refcount.
    - We want COM init/teardown to be localized to the actual operation, rather than
      relying on CLI/GUI outer layers to keep COM alive.
    """
    _com_enter()
    try:
        yield
    finally:
        _com_exit()

# --- Cached PolicyConfigFx interface definitions (define once at import time) ---
# PolicyConfigFx is an undocumented-ish but widely used COM interface that exposes:
# - SetDefaultEndpoint
# - GetPropertyValue/SetPropertyValue for endpoint properties (including FxStore)
#
# We define the interface and associated structs ourselves because:
# - pycaw versions vary in what they ship
# - we need the bFxStore parameter for SysFX work
# - defining COMMETHOD() and ctypes structures dynamically inside a function can
#   interact badly with GC (historically causing intermittent access violations)
#
# Important: we cache only the *definitions* (classes/structs/IIDs), not COM instances.
_POLICY_CONFIG_FX_DEFS = None

def _init_policyconfig_fx_defs_once():
    global _POLICY_CONFIG_FX_DEFS
    if _POLICY_CONFIG_FX_DEFS is not None:
        return

    class PROPERTYKEY(ctypes.Structure):
        _fields_ = (("fmtid", GUID), ("pid", wintypes.DWORD))

    # Prefer comtypes.automation PROPVARIANT, with a small fallback. Some builds of
    # comtypes expose PROPVARIANT under tagPROPVARIANT; compat.py also tries to ensure
    # PROPVARIANT exists.
    try:
        PROPVARIANT = getattr(automation, "PROPVARIANT", getattr(automation, "tagPROPVARIANT"))
    except Exception:
        # Minimal fallback PROPVARIANT definition sufficient for the VT types we touch.
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
            # We need bFxStore variants:
            # - bFxStore=True targets "FxProperties" style storage
            # - bFxStore=False targets the normal property store
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

# Initialize once at import time: defines interface classes early, avoiding runtime
# class construction in the middle of other COM work (historically GC-sensitive).
_init_policyconfig_fx_defs_once()

def _define_policyconfig_fx_interfaces():
    # Backward-compatible helper that now just returns the cached defs
    _init_policyconfig_fx_defs_once()
    return _POLICY_CONFIG_FX_DEFS

# Global cache for PropertyStore interface definitions to avoid GC-related COM crashes
_PROPERTY_STORE_INTERFACES_CACHE = None

def _short_settle(sec=0.15):
    # Small delay used for "settling" driver state after registry/COM writes during learn/verify flows.
    try:
        time.sleep(float(sec))
    except Exception:
        pass

def _reemit_non_error_stderr(buf_text: str):
    """
    Re-emit only non-error lines (e.g., INFO) from captured stderr.
    Suppresses lines starting with 'ERROR:' (ignoring leading whitespace).

    Used by CLI flows that capture stderr around COM/property operations so we
    can keep helpful informational messages while suppressing known noisy errors
    on benign fallback paths.
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

    Many registry locations under MMDevices are keyed by this endpoint GUID,
    not by the full IMMDevice ID string.
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

    How it works:
    - The checkbox ("Listen to this device") is stored as a PropertyStore boolean:
        PROPERTYKEY fmtid={24dbb0fc-9311-4b3d-9cf0-18ff155639d4}, pid=1
      We write it via IPropertyStore::SetValue + Commit.
    - The routing playback target ("Playback through this device") is stored separately
      in MMDevices registry as a string:
        HKLM\...\MMDevices\Audio\Capture\{endpoint-guid}\Properties
        "{24dbb0fc-...},0" = render_device_id or "" (empty string => Default Playback Device)

      We use registry for routing because Windows/driver reliably honors this location.
      Setting routing purely via PropertyStore/COM has historically reported success while
      not changing actual routing.

    Inputs:
    - capture_device_id: full IMMDevice ID for the capture endpoint.
    - enable: bool, desired listen state.
    - render_device_id:
        * None => do not change routing
        * ""   => route to Default Playback Device
        * "<id>" => route to specific render endpoint ID

    Returns:
    - True if the enable toggle write succeeded (routing write best-effort)
    - False on failure (errors printed to stderr; CLI may use captured stderr)
    """
    with _com_context():
        import sys, gc

        # Get cached interface definitions (ctypes vtable layout + PROPVARIANT helpers).
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
            InitPropVariantFromBoolean = propsys.InitPropVariantFromBoolean
            InitPropVariantFromBoolean.restype = HRESULT_T
            InitPropVariantFromBoolean.argtypes = (wintypes.BOOL, POINTER(PROPVARIANT))
        except (AttributeError, OSError):
            # Some systems don't export helper creators; we can still construct a minimal PV.
            have_helpers = False

        PropVariantClear = ole32.PropVariantClear
        PropVariantClear.restype = HRESULT_T
        PropVariantClear.argtypes = (POINTER(PROPVARIANT),)

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
                    # Some PROPVARIANT variants expose fields differently; vt is still enough for many callers.
                    pass
            return pv

        # PKEY_LISTEN_ENABLE:
        # fmtid is the Listen-to-device property set; pid=1 corresponds to the enabled flag.
        PKEY_LISTEN_ENABLE = PROPERTYKEY(GUID("{24dbb0fc-9311-4b3d-9cf0-18ff155639d4}"), 1)

        pv_enable = None
        try:
            pv_enable = _pv_from_bool_local(bool(enable))

            # GC guard: while we hold raw COM pointers and call through a ctypes vtable,
            # we don't want comtypes finalizers to run on unrelated objects and call Release()
            # at an inconvenient time (intermittent access violations).
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

            # Routing target write:
            # - pid=0 is the target endpoint id string (render endpoint)
            # - this is stored under HKLM, and Windows UI/driver behavior is consistent here.
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
                        # Not fatal to the listen enable toggle; routing requires Admin on many machines.
                        print(f"WARNING: Failed to set playback target (requires Admin): {e}", file=sys.stderr)

            return True

        except Exception as e:
            print(f"ERROR: set_listen_to_device_ps failed for '{capture_device_id}': {e}", file=sys.stderr)
            return False

        finally:
            # Always clear PROPVARIANT to avoid leaking allocated memory/strings.
            try:
                if pv_enable is not None:
                    PropVariantClear(byref(pv_enable))
            except Exception:
                pass

def _get_listen_to_device_status_ps(device_id):
    """
    Read the "Listen to this device" enable flag via COM PropertyStore:
      PKEY fmtid={24dbb0fc-...}, pid=1

    Returns:
    - True/False when readable
    - None when unreadable (caller may fall back to registry read)

    This is intended as the authoritative read (same API Windows uses),
    but it can fail on some systems due to COM/property store quirks, so
    callers treat None as "unknown" and may use registry parsing as a fallback.
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

        PKEY_LISTEN_ENABLE = PROPERTYKEY(GUID("{24dbb0fc-9311-4b3d-9cf0-18ff155639d4}"), 1)

        pv = PROPVARIANT()
        try:
            result = None

            # GC guard: we call raw vtable methods through ctypes pointers, so we avoid
            # comtypes __del__/Release() running concurrently and destabilizing COM refcounts.
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
                        # allow caller to fall back to registry
                        result = None
                    else:
                        if getattr(pv, "vt", 0) == VT_BOOL:
                            try:
                                result = (pv.boolVal != VARIANT_FALSE)
                            except Exception:
                                result = None
                        else:
                            # unexpected VT -> let caller fall back
                            result = None
            finally:
                if gc_was_enabled:
                    gc.enable()

            return result

        except Exception as e:
            print(f"WARNING: Failed to read listen status via COM for '{device_id}': {e}", file=sys.stderr)
            return None

        finally:
            try:
                PropVariantClear(byref(pv))
            except Exception:
                pass

def _read_listen_enable_fast(device_id: str):
    """
    Compatibility helper used by CLI/GUI to quickly determine listen state.

    Strategy:
    1) COM PropertyStore read (authoritative when it works).
    2) Registry parsing fallback for resilience.
    """
    state = _get_listen_to_device_status_ps(device_id)
    if state is None:
        state = _read_listen_enable_from_registry(device_id)
    return state

def _read_listen_enable_from_registry(device_id: str):
    r"""
    Robustly read the 'Listen to this device' enable state from MMDevices.

    Why this exists:
    - Some devices/drivers expose the listen flag in slightly different places
      (FxProperties vs Properties) and sometimes as REG_BINARY (PROPVARIANT blob)
      rather than a plain DWORD.
    - COM reads can fail intermittently; registry reads are a pragmatic fallback.

    Returns:
    - True/False if a value can be parsed
    - None if not found/parseable
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
        # Drivers may store booleans as:
        # - REG_DWORD: 0/1
        # - REG_BINARY: PROPVARIANT serialization (VT_BOOL at offset)
        # - REG_SZ: "true"/"false" variants
        if typ == winreg.REG_DWORD:
            try:
                return bool(int(val))
            except Exception:
                return None
        if typ == winreg.REG_BINARY:
            try:
                b = bytes(val)
                if len(b) >= 10:
                    vt = int.from_bytes(b[0:2], "little", signed=False)
                    if vt == 0x000B:  # VT_BOOL
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

    # We check HKCU only here because the listen enable flag is typically per-user.
    # (Routing target uses HKLM and is handled elsewhere.)
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

                # Prefer pid=1 (the canonical "listen enabled" pid).
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
    Poll the registry until the 'Listen' checkbox matches expected_enabled or timeout.

    Used by the CLI "listen" command to provide a verification fallback when:
    - The setter reports failure but state actually changed, or
    - COM readback is inconclusive, or
    - Windows applies the change asynchronously.

    Returns:
    - (True, state) if verified
    - (False, last_state_or_None) if timed out
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
    # Creates a PolicyConfigFx instance for property access (FxStore + normal store).
    IPolicyConfigFx, CLSID_PolicyConfigClient, PROPERTYKEY, PROPVARIANT = _define_policyconfig_fx_interfaces()
    with _com_context():
        return CoCreateInstance(CLSID_PolicyConfigClient, interface=IPolicyConfigFx, clsctx=CLSCTX_ALL)

def _get_policy_config_fx_singleton():
    """
    Create a fresh PolicyConfig object each time - no singleton caching.

    Why:
    - COM objects must be released on the same thread/apartment they were created in.
    - Keeping a singleton alive across GUI operations increases the chance that GC
      will finalize it later from an unexpected context, leading to COM Release races
      and intermittent access violations.
    - Creating a short-lived instance per operation is safer and predictable.
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

    Why:
    - pycaw's packaging varies: some builds expose policyconfig helpers, others don't.
    - Defining COM interfaces dynamically inside a hot codepath is risky: ctypes/comtypes
      class construction can run during GC, and comtypes finalizers may try to Release()
      COM objects concurrently (historically causing access violations).
    - Caching interface *definitions* avoids repeated dynamic class creation.
    """
    global _POLICY_CONFIG_INTERFACES_CACHE

    if _POLICY_CONFIG_INTERFACES_CACHE is not None:
        return _POLICY_CONFIG_INTERFACES_CACHE

    # Try to import from pycaw first (preferred; matches pycaw's known-good definitions).
    try:
        from pycaw.policyconfig import IPolicyConfig, IPolicyConfigVista, CLSID_PolicyConfigClient
        _POLICY_CONFIG_INTERFACES_CACHE = (IPolicyConfig, IPolicyConfigVista, CLSID_PolicyConfigClient)
        return _POLICY_CONFIG_INTERFACES_CACHE
    except Exception:
        pass

    # Define locally if pycaw doesn't have them.
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

    # IPolicyConfig is typically the same as Vista for our purposes.
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
    Get or create IPropertyStore interface definitions once and cache them.

    Why a raw vtable (ctypes) instead of a comtypes wrapper:
    - We want stability across comtypes/pycaw versions; some wrapper paths trigger
      cleanup timing issues.
    - We want direct access to GetValue/SetValue/Commit on IPropertyStore with
      predictable calling conventions.
    - We aggressively control lifetime (GC guards) when using raw pointers.

    Notes:
    - We build a vtable layout (QueryInterface/AddRef/Release/GetValue/SetValue/Commit).
    - CALL uses WINFUNCTYPE on 32-bit and CFUNCTYPE on 64-bit; Windows x64 uses a
      uniform calling convention and WINFUNCTYPE is not appropriate there.
    - PROPVARIANT definitions vary across comtypes versions; we prefer automation.PROPVARIANT
      but have a minimal fallback.
    """
    global _PROPERTY_STORE_INTERFACES_CACHE

    if _PROPERTY_STORE_INTERFACES_CACHE is not None:
        return _PROPERTY_STORE_INTERFACES_CACHE

    try:
        HRESULT_T = wintypes.HRESULT
    except Exception:
        HRESULT_T = ctypes.c_long

    # Calling convention selection:
    # - On 32-bit Windows, stdcall (WINFUNCTYPE) is required.
    # - On 64-bit Windows, the ABI is unified, so CFUNCTYPE is used.
    CALL = ctypes.WINFUNCTYPE if ctypes.sizeof(ctypes.c_void_p) == 4 else ctypes.CFUNCTYPE

    # PROPVARIANT: prefer comtypes.automation; fallback if missing.
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

    # Keep VT_* constants around so we can interpret PROPVARIANT content without relying
    # on comtypes having every constant defined (compat.py patches these).
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
        # VARIANT_TRUE/FALSE semantics (used for VT_BOOL).
        "VARIANT_TRUE": -1,
        "VARIANT_FALSE": 0,
    }

    return _PROPERTY_STORE_INTERFACES_CACHE

def _pkey_disable_sysfx():
    # PKEY_AudioEndpoint_Disable_SysFx:
    # - fmtid {E4870E26-3CC5-4CD2-BA46-CA0A9A70ED04}
    # - pid 2
    #
    # Semantics:
    #   Disable_SysFx = 0 => Enhancements ON
    #   Disable_SysFx = 1 => Enhancements OFF
    #
    # This is the Windows "Audio Enhancements" switch backing store used in diagnostics
    # and learning, but *runtime toggling* is vendor-driven in vendor_db.py.
    IPolicyConfigFx, CLSID_PolicyConfigClient, PROPERTYKEY, PROPVARIANT = _define_policyconfig_fx_interfaces()
    g = _guid_from_parts("E4870E26", "-3CC5-4CD2-", "BA46-", "CA0A9A70ED04")
    return PROPERTYKEY(GUID(g), wintypes.DWORD(2))

def _parse_boolish_from_propvariant(pv):
    # Property stores sometimes expose booleans as VT_BOOL, VT_UI2, or VT_UI4.
    # We normalize to 0/1 or None so the higher-level logic can interpret it.
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
    # Mirror of _parse_boolish_from_propvariant: attempt to write 0/1 into an
    # existing PROPVARIANT while preserving its VT type when possible.
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
    # If we can't preserve VT cleanly, try a couple common fields as best-effort.
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
    Read enhancements state via PolicyConfigFx::GetPropertyValue.

    Details:
    - We probe both bFxStore=True and bFxStore=False because different drivers/Windows
      builds surface Disable_SysFx in different stores.
    - Returns True if enhancements are enabled, False if disabled, None if unknown.
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
    Write enhancements state via PolicyConfigFx::SetPropertyValue.

    Semantics:
    - enable=True  => Disable_SysFx=0
    - enable=False => Disable_SysFx=1

    We attempt both stores (bFxStore True/False) to maximize compatibility.
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
                    # Read current to preserve VT where possible; ignore failures.
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
    - Both HKCU and HKLM (drivers vary; HKLM often needs Admin for writes).
    - Both FxProperties and Properties (Windows/driver version differences).
    - Values are named like "{fmtid},pid" under the endpoint's MMDevices key.

    Returns:
    - True  if enhancements enabled
    - False if enhancements disabled
    - None  if unknown

    Note: Registry stores Disable_SysFx itself, where True/1 means "disabled".
    This function returns the inverted "enabled" boolean for human usage.
    """
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
        # Some drivers store as a serialized PROPVARIANT in REG_BINARY.
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
                        # pid=2 is the canonical Disable_SysFx pid.
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
    Fallback: write Disable_SysFx to registry (DWORD 0/1). Returns True if any write succeeded.

    Used mainly for diagnostics/learning and compatibility testing. Normal runtime toggling
    of enhancements is vendor-driven (see vendor_db.py), but we keep this path available for:
    - discovery flows
    - comparison/verification
    - environments where Windows honors registry writes directly

    prefer_hklm changes hive ordering; HKLM writes typically require Admin.
    """
    guid = _extract_endpoint_guid_from_device_id(device_id)
    if not guid:
        return False
    # Value name format: "{fmtid},pid"
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
    # Poll-based verifier used in some diagnostics/learning flows. Windows/driver may
    # apply property changes asynchronously, so we allow time for the registry view to settle.
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
    Dump ALL values under BOTH HKCU and HKLM for this endpoint.

    Why:
    - Used by discovery/learn flows to capture everything the driver changes when
      you toggle Enhancements or a specific FX.
    - Includes 'dataRaw' in addition to a preview so we can reproduce binary
      payloads exactly when learning a "multi-write" FX toggle.

    Scope:
    - Recurses under both FxProperties and Properties, including nested subkeys
      like FxProperties\{plugin-guid}\User.
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

        rel_subkey is the relative path under the endpoint GUID, e.g.:
          - 'FxProperties'
          - 'FxProperties\\{plugin-guid}\\User'
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
                    "subkey": rel_subkey,
                    "name": name,
                    "type": typ,
                }
                # dataPreview is a small/cheap representation suitable for reports.
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
                # dataRaw is the exact payload used by learning to replay writes.
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
    # Unique key used for diffing MMDevices entries across snapshots.
    # We include hive + flow + subkey + value name to avoid collisions.
    return f"{rec.get('hive','?')}|{rec.get('flow','?')}|{rec.get('subkey','?')}|{rec.get('name','?')}"

def _normalize_preview(v):
    try:
        return (v if isinstance(v, (int, float)) else str(v)).strip() if isinstance(v, str) else v
    except Exception:
        return v

def _diff_mmdevices_lists(before_list, after_list):
    """
    Diff two MMDevices dumps (as produced by _dump_mmdevices_all_values).

    Returns a dict containing:
    - added/removed/changed: record lists for report/debugging
    - dword_flips: specifically REG_DWORD values that flipped 0<->1 (strong toggle candidates)
    - disable_sysfx_hits: any entries involving the Disable_SysFx fmtid (for targeted reporting)

    This is a core building block for learn/discovery workflows. We intentionally keep it
    conservative and based on stable keys, since registry dumps can be noisy.
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

            # Specifically detect clean boolean flips (DWORD 0<->1).
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

    Why this exists:
    - Useful for diagnostics and discovery (live view of what Windows thinks).
    - Some environments can read from PropertyStore even when PolicyConfigFx calls fail.

    Stability:
    - Uses the cached raw vtable interface definitions.
    - GC is disabled while holding raw pointers to avoid comtypes Release races
      (historically intermittent access violations).
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

            # Clear PROPVARIANT after GC is re-enabled (safer for comtypes cleanup timing).
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

    Used mainly for discovery and testing. Runtime toggling is vendor-only.

    Stability:
    - GC disabled while calling through raw vtable pointers to prevent Release races.
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
                    # Try to read existing PV to preserve VT where possible.
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
    Poll the endpoint's IPropertyStore for Disable_SysFx until it matches expected_enabled.

    Used for verification in some fallback paths. Some drivers update the property
    asynchronously after a write.
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
    Collect a full snapshot for enhancement discovery/learning.

    Snapshot contains:
    - com: PolicyConfigFx reads from both fxStore and normal store
    - propStore: live IPropertyStore read for Disable_SysFx
    - registry: full MMDevices dump (HKCU/HKLM, FxProperties/Properties, recursive)

    Used by:
    - discover-enhancements (interactive report)
    - learn flows (to detect which registry keys flip)
    """
    import datetime
    snap = {
        "time": datetime.datetime.now().isoformat(timespec="seconds"),
        "com": {},
        "propStore": {},
        "registry": [],
    }

    # COM snapshot: we also GC-guard this section. PolicyConfig calls are comtypes-based,
    # but we still want to avoid unrelated comtypes finalizers running during this sensitive
    # period (this is part of the "defensive stability" strategy).
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
    Build a human-readable report string for discover-enhancements.

    This report is intended to:
    - summarize COM/PropertyStore views of Disable_SysFx
    - highlight candidate registry keys that flipped during the user's toggle
    - guide the user/developer toward a vendor DWORD that can be learned into the INI
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

    lines.append("Notes:")
    lines.append("- If COM/PropertyStore show A!=B, Windows honored Disable_SysFx and the existing setter is correct.")
    lines.append("- If COM/PropertyStore stay the same but a vendor REG_DWORD flips, that key is likely the real toggle.")
    lines.append("- If only REG_BINARY blobs changed, we may need to write that vendor-specific property.")
    lines.append("")
    return "\n".join(lines)

def _get_policy_config():
    """
    Obtain a PolicyConfig COM interface that supports SetDefaultEndpoint.

    Why the multi-strategy approach:
    - pycaw has shipped multiple variants over time:
      * AudioUtilities may expose a helper to get PolicyConfig
      * policyconfig module may or may not exist / may define different interfaces
    - We prefer pycaw's own helpers when available, otherwise we fall back to our cached
      interface definitions.

    Failure mode:
    - Raises AttributeError if no suitable PolicyConfig interface is available in this environment.
    """
    # 1) Try any helper exposed by the installed pycaw AudioUtilities.
    for name in ("GetPolicyConfig", "_get_policy_config", "get_policy_config"):
        try:
            getter = getattr(AudioUtilities, name, None)
            if getter:
                with _com_context():
                    return getter()
        except Exception:
            pass

    # 2) Use cached interface definitions (either from pycaw or our fallback).
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
    Set the system default endpoint for one or more roles.

    Inputs:
    - device_id: IMMDevice endpoint ID (must be active)
    - role: one of ROLES keys ("console"/"multimedia"/"communications") or "all"

    Safety:
    - We refuse to set defaults to inactive devices (disabled/unplugged/not present).
      This avoids "setting default to something you can't use", and some systems will
      error/crash when asked to set an unplugged endpoint.
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
            # Apply to all three roles; report partial failure with detail.
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
    Verify the endpoint is currently active.

    Why:
    - Prevent setting defaults to unplugged/disabled endpoints.
    - Some PolicyConfig implementations behave unpredictably when passed inactive endpoints.
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
    # Low-level IMMDeviceEnumerator wrapper; always called within COM context by callers.
    with _com_context():
        enumerator = CoCreateInstance(CLSID_MMDeviceEnumerator, IMMDeviceEnumerator, CLSCTX_ALL)
        collection = enumerator.EnumAudioEndpoints(flow, state_mask)
        return enumerator, collection

def get_default_ids(enumerator):
    # Read current default endpoints for each flow/role. Used by list_devices() to flag defaults.
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

    Why prefer this approach:
    - It avoids direct raw PropertyStore reads for every device, which can be sensitive
      to COM lifetime issues on some drivers.
    - pycaw's GetAllDevices typically returns friendly names safely in managed wrappers.

    If name mapping fails, callers fall back to device_id.
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

    This is a hardened fallback path when pycaw-friendly-name mapping isn't sufficient.

    Stability measures explained:
    - We pause GC while we hold raw COM pointers; comtypes finalizers can otherwise
      run mid-call and Release() objects unexpectedly, causing intermittent access
      violations in native code.
    - We AddRef/Release explicitly on the IPropertyStore COM object while using the
      raw vtable pointer, to keep the refcount stable regardless of Python object lifetime.
    - We always call PropVariantClear on the returned PROPVARIANT to avoid leaking memory.
    - comtypes may expose PROPVARIANT.pwszVal as either a pointer or a Python str depending
      on version; we handle multiple access patterns defensively.
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

        # Pause GC so comtypes __del__ won't run Release while we hold raw pointers.
        gc_was_enabled = gc.isenabled()
        if gc_was_enabled:
            try:
                gc.disable()
            except Exception:
                gc_was_enabled = False
        try:
            ps_unknown = dev.OpenPropertyStore(STGM_READ)

            # Balance COM refcount explicitly while we use the raw vtable.
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

                # PKEY_Device_FriendlyName:
                # fmtid {a45c254e-df1c-4efd-8020-67d146a850e0}, pid=14
                PKEY_Device_FriendlyName = PROPERTYKEY(GUID("{a45c254e-df1c-4efd-8020-67d146a850e0}"), 14)
                # Fallback device description:
                PKEY_Device_DeviceDesc   = PROPERTYKEY(GUID("{a45c254e-df1c-4efd-8020-67d146a850e0}"), 2)

                ole32 = ctypes.OleDLL("ole32.dll")
                PropVariantClear = ole32.PropVariantClear
                PropVariantClear.restype = HRESULT_T
                PropVariantClear.argtypes = (ctypes.POINTER(PROPVARIANT),)

                def _read_ptr_or_str(val):
                    if isinstance(val, str):
                        return val
                    if val:
                        try:
                            return ctypes.wstring_at(val)
                        except Exception:
                            return None
                    return None

                def _pv_read_lpwstr(pv):
                    # comtypes can expose pv.pwszVal in different ways; try several layouts.
                    try:
                        s = _read_ptr_or_str(getattr(pv, "pwszVal", None))
                        if s: return s
                    except Exception:
                        pass
                    try:
                        val = getattr(pv, "value", None)
                        if val is not None:
                            s = _read_ptr_or_str(getattr(val, "pwszVal", None))
                            if s: return s
                    except Exception:
                        pass
                    try:
                        data = getattr(pv, "data", None)
                        if data is not None:
                            s = _read_ptr_or_str(getattr(data, "pwszVal", None))
                            if s: return s
                    except Exception:
                        pass
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
                # Balance the AddRef we did above so refcount is correct no matter when GC runs.
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

    # Fallbacks: ID or None
    try:
        return dev.GetId()
    except Exception:
        return None

def list_devices(include_all=False):
    """
    Enumerate endpoints and return a list of device dicts:
      {id, name, flow, state, isDefault}

    Notes:
    - This function returns an unsorted list; GUI/CLI order parity is enforced by
      _sort_and_tag_gui_indices() in higher-level layers.
    - Friendly names are populated from a pycaw GetAllDevices() map first for stability.
      Raw PropertyStore name reads are available as a fallback, but we avoid doing that
      for every device by default because it's more GC/COM sensitive.
    """
    _dbg(f"list_devices: include_all={include_all}")
    with _com_context():
        with warnings.catch_warnings():
            warnings.simplefilter("ignore", UserWarning)

            enumerator_for_defaults = CoCreateInstance(CLSID_MMDeviceEnumerator, IMMDeviceEnumerator, CLSCTX_ALL)
            defaults = get_default_ids(enumerator_for_defaults)
            state_mask = DEVICE_STATE_ALL if include_all else DEVICE_STATE_ACTIVE

            # Build the name map once per call (safer than per-device property-store reads).
            name_map = _friendly_names_by_id()

            out = []

            for flow_name, flow in [("Render", E_RENDER), ("Capture", E_CAPTURE)]:
                _dbg(f"Enum flow={flow_name}")
                enumerator, coll = enum_endpoints(flow, state_mask)
                for i in range(coll.GetCount()):
                    dev = coll.Item(i)
                    dev_id = dev.GetId()

                    # Prefer name map lookup; fall back to raw ID if unknown.
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

    This is intentionally pure/filter-only (no COM). Higher layers decide how to
    disambiguate matches in GUI order.
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

    This is the single source of truth for CLI/GUI index parity.
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

    This is used by CLI commands that allow name selection + --index.
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
    Set mute for an endpoint using IAudioEndpointVolume.

    Returns True/False (no exception propagation). Callers treat False as failure.
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
    Read mute state using IAudioEndpointVolume.GetMute().

    Some comtypes/pycaw versions return:
    - a raw bool/int
    - or a 1-tuple (value,)
    - or require an out-parameter (ctypes BOOL)

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
    Read master volume scalar via IAudioEndpointVolume.GetMasterVolumeLevelScalar().

    Normalization:
    - Windows returns 0.0..1.0 float
    - We convert to integer percent 0..100 (clamped)
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
                        # Some builds require an out-parameter.
                        try:
                            f = ctypes.c_float()
                            vol.GetMasterVolumeLevelScalar(ctypes.byref(f))
                            return max(0, min(100, int(round(float(f.value) * 100.0))))
                        except Exception:
                            return None
    return None

def set_endpoint_volume(device_id, level_percent):
    """
    Set master volume scalar using IAudioEndpointVolume.SetMasterVolumeLevelScalar().

    Input:
    - level_percent: 0..100 (we clamp, then convert to 0.0..1.0 float)

    Returns True/False.
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
    Windows-only verification for fallback paths.

    We require the live PropertyStore Disable_SysFx to match the expected_enabled state
    for 'consecutive' reads to be confident (guards against transient/async updates).

    Returns:
    - (ok, verifiedBy, finalState)

    verifiedBy tags (used by higher layers in JSON):
    - "windows-live(ps)" indicates the Windows PropertyStore readback path.
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
