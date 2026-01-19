# audioctl.py (entrypoint used by PyInstaller)
# Attach to the parent console on Windows so PS 5.1 doesn't spawn a new window
import os
if os.name == "nt":
    try:
        import ctypes
        ctypes.windll.kernel32.AttachConsole(-1)  # ATTACH_PARENT_PROCESS
    except Exception:
        pass

from audioctl.cli import main

if __name__ == "__main__":
    import sys
    sys.exit(main())
