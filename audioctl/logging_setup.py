# audioctl/logging_setup.py
"""
Centralized logging + crash/exception breadcrumbs for both CLI and GUI.

Design goals (why this module looks "overbuilt"):
- **Lazy initialization**: importing this module should not perform file I/O.
  This matters for:
  - fast CLI startup (especially when used as a helper from scripts/hotkeys)
  - frozen/PyInstaller apps where early filesystem writes can fail or slow launch
  - avoiding side effects during module import order (important with COM tooling)
- **Robust path selection**: prefer logging next to the executable for "portable"
  distributions, but transparently fall back to a user-writable location if the
  exe directory is not writable (e.g., Program Files).
- **Best-effort diagnostics**: all logging helpers swallow exceptions so that
  logging itself never becomes the reason the app fails (or triggers recursion).
"""
import os
import sys
import traceback
import datetime
import tempfile
import threading

try:
    import faulthandler
except Exception:
    # faulthandler is optional; we treat it as a "nice to have" for native crashes.
    faulthandler = None

# Debug toggle (runtime)
# Controlled via env var so we can enable verbose logs without changing code
# (critical when debugging user machines / frozen builds).
_DEBUG = bool(int(os.environ.get("AUDIOCTL_DEBUG", "0")))

# Internal state (lazy init: no file I/O at import time)
# We intentionally do not resolve/create the log file until the first actual
# _log/_dbg call, so importing audioctl stays side-effect free.
_LOG_DIR = None
_LOG_PATH = None
_INITIALIZED = False
_FH = None            # faulthandler file handle (kept open for process lifetime)
_HOOKS_INSTALLED = False


def set_debug(on: bool = True):
    """Enable/disable debug logging at runtime (primarily for dev/testing)."""
    global _DEBUG
    _DEBUG = bool(on)
    try:
        _log(f"DEBUG {'enabled' if _DEBUG else 'disabled'}")
    except Exception:
        # Logging is best-effort; never let debug toggling break app flow.
        pass


def _exe_dir():
    """
    Return the directory we consider "app-local".

    Why:
    - For frozen apps (PyInstaller), sys.executable points to the bundled EXE.
      Putting logs next to it is convenient for portability.
    - For source runs, keep logs near this module to keep everything together.
    """
    try:
        if getattr(sys, "frozen", False):  # PyInstaller/py2exe
            return os.path.dirname(sys.executable)
        return os.path.dirname(os.path.abspath(__file__))
    except Exception:
        return os.getcwd()


def resource_path(name: str):
    """
    Locate a resource file bundled with the app (icon, etc).

    Why:
    - When frozen, PyInstaller extracts resources into sys._MEIPASS.
    - When not frozen, they live next to the source tree.
    """
    base = getattr(sys, "_MEIPASS", _exe_dir())
    return os.path.join(base, name)


def _resolve_log_path():
    """
    Decide where the log would live, but do not create it yet.

    Strategy:
    1) Prefer the executable/module directory (portable behavior).
    2) If that directory isn't writable (common under Program Files),
       fall back to a temp/appdata location.

    Note: we *probe* writability by creating a small temporary file. We still
    avoid creating the real log file here to keep import-time side effects low.
    """
    base = _exe_dir()
    path = os.path.join(base, "audioctl_gui.log")
    try:
        os.makedirs(base, exist_ok=True)
        # Probe writability without creating the real log
        test = os.path.join(base, ".writetest")
        with open(test, "w", encoding="utf-8") as _:
            pass
        os.remove(test)
        return base, path
    except Exception:
        # If the EXE dir isn't writable, use a user-writable fallback.
        # This makes the GUI usable even when installed in protected folders.
        tdir = os.path.join(tempfile.gettempdir(), "audioctl")
        try:
            os.makedirs(tdir, exist_ok=True)
        except Exception:
            # Absolute last resort: raw temp dir.
            tdir = tempfile.gettempdir()
        return tdir, os.path.join(tdir, "audioctl_gui.log")


def _ensure_resolved():
    """
    Resolve path variables, but do not touch the filesystem.

    Why:
    - Some callers (e.g., error dialogs) want the log path for display even if
      we never successfully write the log.
    """
    global _LOG_DIR, _LOG_PATH
    if _LOG_DIR is None or _LOG_PATH is None:
        _LOG_DIR, _LOG_PATH = _resolve_log_path()


def _global_excepthook(exc_type, exc_value, exc_tb):
    """
    sys.excepthook replacement.

    Why:
    - Captures uncaught exceptions (including those from Tk callbacks in some
      configurations) into the log file.
    - Still delegates to the original excepthook so default stderr printing
      remains intact when a console exists.
    """
    _log_exc("UNCAUGHT EXCEPTION", (exc_type, exc_value, exc_tb))
    try:
        sys.__excepthook__(exc_type, exc_value, exc_tb)
    except Exception:
        pass


def _unraisable_hook(unraisable):
    """
    Python 3.8+ hook for exceptions raised in __del__ and other "unraisable" places.

    Why:
    - comtypes and ctypes-based COM wrappers can raise exceptions during garbage
      collection/finalization. These are easy to miss, and are exactly the kind
      of data we want when chasing shutdown/GC-related COM crashes.
    """
    try:
        _log(f"UNRAISABLE: {getattr(unraisable.exc_type, '__name__', str(unraisable.exc_type))}: "
             f"{unraisable.exc_value}\nObject: {unraisable.object!r}")
    except Exception:
        pass


def _install_hooks_once():
    """
    Install exception/console hooks (idempotent). No file I/O here.

    Why:
    - Hooks should be installed only once even if modules are imported multiple
      times (e.g., in embedded Python or some test harnesses).
    - Installation is deferred until first log write so we don't modify global
      interpreter behavior during import unless logging is actually used.
    """
    global _HOOKS_INSTALLED
    if _HOOKS_INSTALLED:
        return

    # Uncaught exceptions -> log file.
    try:
        sys.excepthook = _global_excepthook
    except Exception:
        pass

    # Unraisable exceptions -> log file (Python 3.8+).
    try:
        sys.unraisablehook = _unraisable_hook
    except Exception:
        pass

    # Console control handler (Windows)
    # Why:
    # - Helps correlate "random exit" reports with actual console events
    #   (Ctrl+C, close, logoff/shutdown).
    # - Particularly useful when users run the frozen EXE from a terminal.
    try:
        import ctypes
        from ctypes import wintypes
        _ConsoleCtrlHandler = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.DWORD)

        def _console_ctrl_handler(ctrl_type):
            try:
                _log(f"ConsoleCtrl event: {ctrl_type}")
            except Exception:
                pass
            # Returning False tells Windows "we didn't fully handle it"; default
            # behavior still occurs (termination, etc).
            return False

        try:
            ctypes.windll.kernel32.SetConsoleCtrlHandler(_ConsoleCtrlHandler(_console_ctrl_handler), True)
        except Exception:
            pass
    except Exception:
        pass

    _HOOKS_INSTALLED = True


# sys.exit wrapper (installed lazily)
# Why:
# - Many program exits are "normal" but still useful when correlating logs
#   (e.g., sys.exit() from GUI close or fatal argument parsing).
_orig_sys_exit = sys.exit


def _logged_sys_exit(code=0):
    # Breadcrumb before raising SystemExit (which is how sys.exit works internally).
    try:
        _log(f"sys.exit invoked with code={code!r}")
    except Exception:
        pass
    raise SystemExit(code)


def _atexit_close_handles():
    """
    Best-effort cleanup for faulthandler file handle.

    Why:
    - Ensure log output is flushed even on normal interpreter shutdown.
    - Avoid leaving file descriptors open on long-running embedding scenarios.
    """
    try:
        global _FH
        if faulthandler and _FH and not _FH.closed:
            _FH.flush()
            _FH.close()
            _FH = None
    except Exception:
        pass


def _atexit_normal():
    # Breadcrumb that the interpreter reached atexit (useful when debugging crashes
    # where this line never appears).
    _log("atexit: process exiting normally")


def _ensure_init():
    """
    Initialize logging on first use (lazy):
    - Resolve and create the log file
    - Write the first breadcrumb
    - Install exception hooks and console handler
    - Enable faulthandler (if available)
    - Wrap sys.exit
    - Register atexit handlers

    Why this is lazy:
    - Prevent import-time file creation in CLI and frozen apps.
    - Avoid failures when the process lacks filesystem permissions until we
      actually need the log.
    - Minimize side effects during module import (import order matters when COM
      modules and PyInstaller shims are involved).
    """
    global _INITIALIZED, _FH
    if _INITIALIZED:
        return

    _ensure_resolved()

    # Create file and first breadcrumb (this is the first actual write).
    # If this fails (permissions, AV, corporate lockdown), we still keep running.
    try:
        ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with open(_LOG_PATH, "a", encoding="utf-8", errors="replace") as f:
            f.write(f"[{ts}] logging to: {_LOG_PATH}\n")
    except Exception:
        pass

    # Install global hooks only once; keep this out of import time.
    _install_hooks_once()

    # faulthandler to the log file if possible
    # Why:
    # - Native crashes (e.g., access violations from COM Release races) bypass
    #   Python exceptions entirely; faulthandler gives us a Python stack snapshot.
    # - all_threads=True is important because COM-related issues often happen on
    #   worker threads (GUI threads, subprocess I/O threads, etc).
    if faulthandler and _FH is None:
        try:
            _FH = open(_LOG_PATH, "a", buffering=1)
            faulthandler.enable(file=_FH, all_threads=True)
        except Exception:
            # Fallback: enable without a file (goes to stderr). Still helpful
            # when running from console.
            try:
                faulthandler.enable(all_threads=True)
            except Exception:
                pass

    # Wrap sys.exit once so exits become visible in logs.
    try:
        if sys.exit is not _logged_sys_exit:
            sys.exit = _logged_sys_exit
    except Exception:
        pass

    # atexit handlers give us "last breadcrumb wins" behavior on normal exit.
    try:
        import atexit
        atexit.register(_atexit_normal)
        atexit.register(_atexit_close_handles)
    except Exception:
        pass

    _INITIALIZED = True


def _log_path():
    # Return the would-be log path without forcing log file creation.
    _ensure_resolved()   # no file I/O here
    return _LOG_PATH


def _log(msg: str):
    """
    Append a timestamped line to the log.

    Best-effort by design:
    - This function must never raise. Logging failures (permissions, encoding
      errors, etc.) should not break the application or trigger recursive
      exception handling.
    """
    _ensure_init()       # creates file on first use
    try:
        ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with open(_LOG_PATH, "a", encoding="utf-8", errors="replace") as f:
            f.write(f"[{ts}] {msg}\n")
    except Exception:
        pass


def _log_exc(prefix: str, exc_info=None):
    """
    Log an exception traceback.

    Why:
    - Centralizes formatting so we get consistent tracebacks in the same file as
      our other breadcrumbs.
    - Swallows errors to avoid "logging the logging failure" recursion.
    """
    try:
        if exc_info is None:
            exc_info = sys.exc_info()
        tb = "".join(traceback.format_exception(*exc_info))
        _log(f"{prefix}\n{tb}")
    except Exception:
        pass


def _dbg(msg: str):
    """
    Debug logging (opt-in).

    Why we include pid/tid:
    - COM issues are often thread-affine (apartment model). Being able to see
      thread IDs alongside operations is crucial when debugging GC/Release races
      and cross-thread COM object usage.
    """
    if not _DEBUG:
        return
    try:
        _ensure_init()   # creates file on first debug write
        ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        tid = threading.get_ident()
        pid = os.getpid()
        with open(_LOG_PATH, "a", encoding="utf-8", errors="replace") as f:
            f.write(f"[{ts}] [DBG pid={pid} tid={tid}] {msg}\n")
    except Exception:
        pass


def enable_gc_debug():
    """
    Enable Python GC debug flags (development aid).

    Notes:
    - DEBUG_SAVEALL keeps unreachable objects in gc.garbage, which can increase
      memory usage substantially and change GC behavior. This should be used
      only when chasing GC/comtypes finalizer issues.
    """
    try:
        import gc
        gc.set_debug(gc.DEBUG_SAVEALL)
        _dbg("GC debug flags enabled (DEBUG_SAVEALL)")
    except Exception:
        pass


# Optional no-op; kept for compatibility
def init_logging_runtime(enable_faulthandler=True):
    # Historically this module exposed an "init" call; we now keep it as a
    # compatibility shim and simply force lazy init.
    _ensure_init()
