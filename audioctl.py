# audioctl.py
# -----------------------------------------------------------------------------
# PyInstaller entrypoint.
#
# This tiny wrapper exists so we have a stable, single-file "script" target for
# PyInstaller (and for direct `python audioctl.py` usage), while keeping the
# actual application implementation inside the `audioctl` package.
#
# It also contains a Windows-specific console attachment hack so that when the
# executable is launched from an existing console (especially PowerShell 5.1),
# we reuse that console instead of flashing/spawning a second console window.
# -----------------------------------------------------------------------------

import os

# On Windows, some launch paths (notably older PowerShell 5.1 behavior) can
# cause a console EXE to appear in a new console window even when invoked from
# an existing console. AttachConsole(ATTACH_PARENT_PROCESS) forces this process
# to bind to the parent console so stdout/stderr behave as expected for CLI use.
if os.name == "nt":
    try:
        import ctypes
        ctypes.windll.kernel32.AttachConsole(-1)  # ATTACH_PARENT_PROCESS
    except Exception:
        # Non-fatal: if we cannot attach (no parent console, permission limits,
        # or unusual host process), the CLI can still run normally.
        # We intentionally swallow this to avoid breaking GUI launch or scripts.
        pass

# Import the real entrypoint from the package so the CLI/GUI implementation
# lives in one place (audioctl/cli.py), regardless of whether we run from source
# (`python -m audioctl`) or as a frozen EXE.
from audioctl.cli import main

if __name__ == "__main__":
    import sys
    # Propagate the CLI exit code for automation/scripting (PowerShell, batch,
    # CI pipelines). main() returns an int consistent with the documented codes.
    sys.exit(main())
