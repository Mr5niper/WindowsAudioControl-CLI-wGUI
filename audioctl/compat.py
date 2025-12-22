# audioctl/compat.py
"""
Compatibility / global shims that must be imported before pycaw/comtypes.
"""
# --- comtypes compatibility shim (MUST be at the VERY TOP, before ANY other imports including pycaw) ---
try:
    import comtypes.automation as _automation
    # Ensure PROPVARIANT alias exists, using tagPROPVARIANT as a common fallback.
    if not hasattr(_automation, "PROPVARIANT") and hasattr(_automation, "tagPROPVARIANT"):
        _automation.PROPVARIANT = _automation.tagPROPVARIANT
    # Ensure standard VT_ constants are present.
    if not hasattr(_automation, "VT_LPWSTR"):
        _automation.VT_LPWSTR = 31  # VT_LPWSTR
    if not hasattr(_automation, "VT_BOOL"):
        _automation.VT_BOOL = 11   # VT_BOOL
    if not hasattr(_automation, "VT_UI2"):
        _automation.VT_UI2 = 18
    if not hasattr(_automation, "VT_UI4"):
        _automation.VT_UI4 = 19
    # Ensure VARIANT_TRUE/FALSE for boolean properties
    if not hasattr(_automation, "VARIANT_TRUE"):
        _automation.VARIANT_TRUE = -1
    if not hasattr(_automation, "VARIANT_FALSE"):
        _automation.VARIANT_FALSE = 0
except Exception as e:
    import sys
    print(f"WARNING: Global comtypes compatibility shim failed during initial import: {e}", file=sys.stderr)
# ------------------------------------------------------------------------------------------------
import ctypes
import sys
def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except Exception:
        return False
# STGM flags used by OpenPropertyStore
STGM_READ  = 0x00000000
STGM_WRITE = 0x00000001
# Endpoint flows & roles
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
# Device state flags
DEVICE_STATE_ACTIVE = 0x00000001
DEVICE_STATE_ALL    = 0x0000000F  # active | disabled | notpresent | unplugGED
DEVICE_STATES = {
    0x00000001: "active",
    0x00000002: "disabled",
    0x00000004: "notpresent",
    0x00000008: "unplugged",
}
def _guid_from_parts(*parts: str) -> str:
    """
    Assemble a GUID string from parts to avoid embedding exact literals.
    Example: _guid_from_parts("E4870E26", "-3CC5-4CD2-", "BA46-", "CA0A9A70ED04")
    """
    return "{" + "".join(parts) + "}"