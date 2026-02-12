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


def format_cmd_for_display(argv, *, force_quote_chars: str = "") -> str:
    """
    Format an argv list into a command string suitable for display/logging.

    Windows:
      - Uses subprocess.list2cmdline() (double quotes; CreateProcess rules).
      - If force_quote_chars is provided, force-quote any arg that contains
        whitespace, is empty, or contains any character in force_quote_chars.

    Non-Windows:
      - Uses shlex.join() when available.

    Args:
      argv: Iterable of arguments (typically a list[str]).
      force_quote_chars: Characters that should trigger forced quoting on Windows.

    Returns:
      A single string representing the command.
    """
    if argv is None:
        return ""

    # Normalize to strings; be defensive about None/non-str values.
    args = ["" if a is None else str(a) for a in argv]

    if os.name == "nt":
        if force_quote_chars:
            out = []
            for a in args:
                needs = (
                    a == ""
                    or any(ch.isspace() for ch in a)
                    or any(ch in a for ch in force_quote_chars)
                )
                if needs:
                    # Correct Windows quoting/escaping for a single arg
                    out.append(subprocess.list2cmdline([a]))
                else:
                    out.append(a)
            return " ".join(out)
        return subprocess.list2cmdline(args)

    try:
        import shlex

        return shlex.join(args)
    except Exception:
        # Very old Python or unexpected types; last-resort fallback.
        return " ".join(args)


def format_audioctl_cmd_for_display(args, *, frozen: bool = False, cross_shell: bool = True) -> str:
    """
    Format an `audioctl ...` command line for humans to copy/paste.

    If frozen and cross_shell=True on Windows:
      - prefix with .\\audioctl.exe so it works in both cmd.exe and PowerShell
        from the current directory
      - force-quote `{}` so PowerShell doesn't misparse device IDs
    """
    if frozen and os.name == "nt" and cross_shell:
        prefix = r".\audioctl.exe"
        force = "{}"
    else:
        prefix = "audioctl"
        force = ""
    return prefix + " " + format_cmd_for_display(args, force_quote_chars=force)



