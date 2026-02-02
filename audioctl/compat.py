# audioctl/compat.py
"""
Compatibility / global shims that must be imported before pycaw/comtypes.

Why this module exists (and why import order is critical)
---------------------------------------------------------
This project relies on `comtypes` (via `pycaw`) to talk to Windows Core Audio COM
interfaces. In practice, different `comtypes` versions/builds (and especially
PyInstaller-frozen builds) can expose slightly different symbol sets under
`comtypes.automation`. If we import `pycaw`/`comtypes` first and *then* discover
missing aliases/constants, we can end up with:
- incorrect PROPVARIANT layouts,
- missing VT_* type constants used to interpret property values,
- and, in frozen builds, shutdown-time crashes triggered by late imports during
  COM cleanup.

So: import this module BEFORE any code that touches comtypes/pycaw.
"""

# --- comtypes compatibility shim (MUST be at the VERY TOP) ---
try:
    import comtypes.automation as _automation

    # In some comtypes builds, the PROPVARIANT type is exposed as tagPROPVARIANT
    # instead of PROPVARIANT. Our code (and many examples online) reference
    # PROPVARIANT, so we alias it to keep downstream code stable.
    # This matters because PropertyStore reads/writes return PROPVARIANTs and we
    # interpret `vt` + union members to determine bool/string/dword states.
    if not hasattr(_automation, "PROPVARIANT") and hasattr(_automation, "tagPROPVARIANT"):
        _automation.PROPVARIANT = _automation.tagPROPVARIANT

    # `VT_*` constants are used when decoding/encoding PROPVARIANT values:
    # - VT_LPWSTR for strings (friendly names, some properties)
    # - VT_BOOL / VT_UI2 / VT_UI4 for various boolean-ish flags seen in audio
    #   property stores and vendor keys.
    #
    # Some bundled environments omit these constants from comtypes.automation,
    # so we define the canonical numeric values used by Windows VARIANT.
    if not hasattr(_automation, "VT_LPWSTR"):
        _automation.VT_LPWSTR = 31
    if not hasattr(_automation, "VT_BOOL"):
        _automation.VT_BOOL = 11
    if not hasattr(_automation, "VT_UI2"):
        _automation.VT_UI2 = 18
    if not hasattr(_automation, "VT_UI4"):
        _automation.VT_UI4 = 19

    # COM VARIANT booleans are historically defined as:
    #   VARIANT_TRUE  = -1 (all bits set in a signed 16-bit value)
    #   VARIANT_FALSE = 0
    # This surprises Python developers, but it's the COM convention and some
    # helper APIs (InitPropVariantFromBoolean, etc.) expect exactly these values.
    if not hasattr(_automation, "VARIANT_TRUE"):
        _automation.VARIANT_TRUE = -1
    if not hasattr(_automation, "VARIANT_FALSE"):
        _automation.VARIANT_FALSE = 0

except Exception as e:
    # This shim is "best effort": if comtypes isn't importable (non-Windows,
    # missing dependency, etc.) we don't hard-fail at import time.
    import sys
    print(f"WARNING: Global comtypes compatibility shim failed during initial import: {e}", file=sys.stderr)

# --- NEW SECTION: Force import of post_coinit modules ---
# comtypes uses internal "_post_coinit" modules to finalize certain COM types
# and to provide correct cleanup behavior (e.g., __del__ -> Release()).
#
# Two key reasons to force these imports early:
#  1) Prevent "late import during finalization":
#     If these modules are first imported during interpreter shutdown or GC,
#     Python can be in a partially-torn-down state. That timing tends to
#     amplify COM apartment rules and can surface as noisy unraisable exceptions
#     or hard crashes in frozen builds.
#  2) PyInstaller bundling:
#     PyInstaller sometimes fails to detect dynamic/late imports. Importing
#     these modules at startup makes them visible to the bundler, ensuring the
#     frozen executable actually contains them, which reduces shutdown crashes.
try:
    import comtypes._post_coinit
    import comtypes._post_coinit.unknwn
except ImportError:
    # If comtypes isn't present (or on platforms where it doesn't apply),
    # we simply skip.
    pass

# ----------------------------------------------------------------
import ctypes
import sys


def is_admin():
    # Lightweight UAC elevation check used for user-facing warnings.
    # - On Windows, IsUserAnAdmin() returns non-zero if the process token is elevated.
    # - On non-Windows or restricted contexts (service sandbox, hardened policy),
    #   this call can raise; we treat that as "not admin" rather than failing.
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except Exception:
        return False


# STGM flags used by OpenPropertyStore.
# These are COM storage access flags: we use them when opening an endpoint's
# IPropertyStore for read-only queries (STGM_READ) or for property writes
# (STGM_WRITE) such as toggling "Listen to this device" and SysFX properties.
STGM_READ  = 0x00000000
STGM_WRITE = 0x00000001

# Endpoint flows & roles (Core Audio policy concepts):
# - Flow determines whether an endpoint is Playback (Render) or Recording (Capture).
# - Role determines which "default device" slot is being set:
#     console        -> general/default audio device for apps
#     multimedia     -> media playback role (often same as console on many systems)
#     communications -> telephony/voice apps (Teams/Zoom/etc.)
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

# Device state flags returned by IMMDevice::GetState / EnumAudioEndpoints masks.
# We default most mutating operations to ACTIVE only to avoid touching
# disconnected/disabled endpoints.
DEVICE_STATE_ACTIVE = 0x00000001
DEVICE_STATE_ALL    = 0x0000000F  # active | disabled | notpresent | unplugged
DEVICE_STATES = {
    0x00000001: "active",
    0x00000002: "disabled",
    0x00000004: "notpresent",
    0x00000008: "unplugged",
}


def _guid_from_parts(*parts: str) -> str:
    """
    Assemble a GUID string from parts to avoid embedding exact literals.

    Why do this:
    - It keeps long GUIDs visually grouped and easier to review.
    - It reduces the risk of accidental typos when copying GUIDs around.
    - It also avoids sprinkling full literal GUID strings throughout the code,
      which helps keep low-level identifiers centralized and recognizable.

    Example: _guid_from_parts("E4870E26", "-3CC5-4CD2-", "BA46-", "CA0A9A70ED04")
    """
    return "{" + "".join(parts) + "}"
