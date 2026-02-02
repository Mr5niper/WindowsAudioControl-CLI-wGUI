# audioctl/__init__.py

# Package-level entrypoint export.
# We re-export `main` so callers can embed/invoke the CLI/GUI dispatcher
# programmatically (e.g., tests, "python -c", other tools) without relying on
# the module-as-script behavior in __main__.py.
from .cli import main

# Deliberately keep the public API surface tiny and stable.
# This avoids accidental reliance on internal helper modules (COM/registry code),
# which are implementation details and can change independently.
__all__ = ["main"]
