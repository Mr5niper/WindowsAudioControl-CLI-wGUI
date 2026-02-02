# audioctl/__main__.py
# This module exists so `python -m audioctl` behaves like the built EXE entrypoint.
# It forwards execution into the shared CLI/GUI dispatcher (audioctl.cli.main),
# keeping a single source of truth for behavior across:
#   - `python -m audioctl` (module execution)
#   - `audioctl.py` (PyInstaller entry stub)
#   - `audioctl.exe` (frozen executable)

from .cli import main

if __name__ == "__main__":
    # Preserve standard CLI semantics:
    # - main() consumes sys.argv (unless argv is explicitly provided)
    # - sys.exit(main()) propagates the integer return code as the process exit code
    import sys
    sys.exit(main())
