# audioctl/logging_setup.py
import os
import sys
import traceback
import datetime
import tempfile
try:
    import faulthandler
except Exception:
    faulthandler = None
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
def _init_log_paths():
    base = _exe_dir()
    path = os.path.join(base, "audioctl_gui.log")
    try:
        os.makedirs(base, exist_ok=True)
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
_LOG_DIR, _LOG_PATH = _init_log_paths()
def _log_path():
    return _LOG_PATH
def _log(msg: str):
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
# --- Global exception hooks and faulthandler setup ---
# Create/append the first breadcrumb
try:
    with open(_LOG_PATH, "a", encoding="utf-8", errors="replace") as f:
        ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        f.write(f"[{ts}] logging to: {_LOG_PATH}\n")
except Exception:
    pass
def _global_excepthook(exc_type, exc_value, exc_tb):
    _log_exc("UNCAUGHT EXCEPTION", (exc_type, exc_value, exc_tb))
    try:
        sys.__excepthook__(exc_type, exc_value, exc_tb)
    except Exception:
        pass
sys.excepthook = _global_excepthook
try:
    def _unraisable_hook(unraisable):
        _log(f"UNRAISABLE: {getattr(unraisable.exc_type, '__name__', str(unraisable.exc_type))}: "
             f"{unraisable.exc_value}\nObject: {unraisable.object!r}")
    sys.unraisablehook = _unraisable_hook
except Exception:
    pass
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
import atexit
@atexit.register
def _on_atexit():
    _log("atexit: process exiting normally")
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
# Console control handler (Windows)
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
        return False
    try:
        ctypes.windll.kernel32.SetConsoleCtrlHandler(_ConsoleCtrlHandler(_console_ctrl_handler), True)
    except Exception:
        pass
except Exception:
    pass