# audioctl/logging_setup.py
import os
import sys
import traceback
import datetime
import tempfile
import threading

try:
    import faulthandler
except Exception:
    faulthandler = None

# Debug toggle (runtime)
_DEBUG = bool(int(os.environ.get("AUDIOCTL_DEBUG", "0")))

# Internal state (lazy init: no file I/O at import time)
_LOG_DIR = None
_LOG_PATH = None
_INITIALIZED = False
_FH = None            # faulthandler file handle
_HOOKS_INSTALLED = False

def set_debug(on: bool = True):
    global _DEBUG
    _DEBUG = bool(on)
    try:
        _log(f"DEBUG {'enabled' if _DEBUG else 'disabled'}")
    except Exception:
        pass

def _exe_dir():
    try:
        if getattr(sys, "frozen", False):  # PyInstaller/py2exe
            return os.path.dirname(sys.executable)
        return os.path.dirname(os.path.abspath(__file__))
    except Exception:
        return os.getcwd()

def resource_path(name: str):
    # Works both when frozen (PyInstaller) and when running from source
    base = getattr(sys, "_MEIPASS", _exe_dir())
    return os.path.join(base, name)

def _resolve_log_path():
    """
    Decide where the log would live, but do not create it yet.
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
        tdir = os.path.join(tempfile.gettempdir(), "audioctl")
        try:
            os.makedirs(tdir, exist_ok=True)
        except Exception:
            tdir = tempfile.gettempdir()
        return tdir, os.path.join(tdir, "audioctl_gui.log")

def _ensure_resolved():
    """
    Resolve path variables, but do not touch the filesystem.
    """
    global _LOG_DIR, _LOG_PATH
    if _LOG_DIR is None or _LOG_PATH is None:
        _LOG_DIR, _LOG_PATH = _resolve_log_path()

def _global_excepthook(exc_type, exc_value, exc_tb):
    _log_exc("UNCAUGHT EXCEPTION", (exc_type, exc_value, exc_tb))
    try:
        sys.__excepthook__(exc_type, exc_value, exc_tb)
    except Exception:
        pass

def _unraisable_hook(unraisable):
    try:
        _log(f"UNRAISABLE: {getattr(unraisable.exc_type, '__name__', str(unraisable.exc_type))}: "
             f"{unraisable.exc_value}\nObject: {unraisable.object!r}")
    except Exception:
        pass

def _install_hooks_once():
    """
    Install exception/console hooks (idempotent). No file I/O here.
    """
    global _HOOKS_INSTALLED
    if _HOOKS_INSTALLED:
        return
    try:
        sys.excepthook = _global_excepthook
    except Exception:
        pass
    try:
        sys.unraisablehook = _unraisable_hook
    except Exception:
        pass

    # Console control handler (Windows)
    try:
        import ctypes
        from ctypes import wintypes
        _ConsoleCtrlHandler = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.DWORD)
        def _console_ctrl_handler(ctrl_type):
            try:
                _log(f"ConsoleCtrl event: {ctrl_type}")
            except Exception:
                pass
            return False
        try:
            ctypes.windll.kernel32.SetConsoleCtrlHandler(_ConsoleCtrlHandler(_console_ctrl_handler), True)
        except Exception:
            pass
    except Exception:
        pass

    _HOOKS_INSTALLED = True

# sys.exit wrapper (installed lazily)
_orig_sys_exit = sys.exit
def _logged_sys_exit(code=0):
    try:
        _log(f"sys.exit invoked with code={code!r}")
    except Exception:
        pass
    raise SystemExit(code)

def _atexit_close_handles():
    try:
        global _FH
        if faulthandler and _FH and not _FH.closed:
            _FH.flush()
            _FH.close()
            _FH = None
    except Exception:
        pass

def _atexit_normal():
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
    """
    global _INITIALIZED, _FH
    if _INITIALIZED:
        return

    _ensure_resolved()

    # Create file and first breadcrumb (this is the first actual write)
    try:
        ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with open(_LOG_PATH, "a", encoding="utf-8", errors="replace") as f:
            f.write(f"[{ts}] logging to: {_LOG_PATH}\n")
    except Exception:
        pass

    _install_hooks_once()

    # faulthandler to the log file if possible
    if faulthandler and _FH is None:
        try:
            _FH = open(_LOG_PATH, "a", buffering=1)
            faulthandler.enable(file=_FH, all_threads=True)
        except Exception:
            try:
                faulthandler.enable(all_threads=True)
            except Exception:
                pass

    # Wrap sys.exit once
    try:
        if sys.exit is not _logged_sys_exit:
            sys.exit = _logged_sys_exit
    except Exception:
        pass

    # atexit
    try:
        import atexit
        atexit.register(_atexit_normal)
        atexit.register(_atexit_close_handles)
    except Exception:
        pass

    _INITIALIZED = True

def _log_path():
    _ensure_resolved()   # no file I/O here
    return _LOG_PATH

def _log(msg: str):
    _ensure_init()       # creates file on first use
    try:
        ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with open(_LOG_PATH, "a", encoding="utf-8", errors="replace") as f:
            f.write(f"[{ts}] {msg}\n")
    except Exception:
        pass

def _log_exc(prefix: str, exc_info=None):
    try:
        if exc_info is None:
            exc_info = sys.exc_info()
        tb = "".join(traceback.format_exception(*exc_info))
        _log(f"{prefix}\n{tb}")
    except Exception:
        pass

def _dbg(msg: str):
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
    try:
        import gc
        gc.set_debug(gc.DEBUG_SAVEALL)
        _dbg("GC debug flags enabled (DEBUG_SAVEALL)")
    except Exception:
        pass

# Optional no-op; kept for compatibility
def init_logging_runtime(enable_faulthandler=True):
    _ensure_init()
