# audioctl/cmdline_fmt.py
#
# Small utility for formatting argv lists into a human-friendly command string.
#
# IMPORTANT:
# - This is for display/logging only.
# - Do NOT use this to execute subprocesses (always pass argv as a list).
#
# Why this exists:
# - shlex.join() is POSIX-oriented and will show single quotes on Windows.
# - On Windows we want CreateProcess-compatible quoting and double quotes.
#   subprocess.list2cmdline() implements the correct Windows quoting rules.
#
import os
import subprocess


def format_cmd_for_display(argv) -> str:
    """
    Format an argv list into a command string suitable for display/logging.

    Windows:
      - Uses subprocess.list2cmdline() (double quotes; CreateProcess rules).

    Non-Windows:
      - Uses shlex.join() when available.

    Args:
      argv: Iterable of arguments (typically a list[str]).

    Returns:
      A single string representing the command.
    """
    if argv is None:
        return ""

    # Normalize to strings; be defensive about None/non-str values.
    args = ["" if a is None else str(a) for a in argv]

    if os.name == "nt":
        return subprocess.list2cmdline(args)

    try:
        import shlex
        return shlex.join(args)
    except Exception:
        # Very old Python or unexpected types; last-resort fallback.
        return " ".join(args)
