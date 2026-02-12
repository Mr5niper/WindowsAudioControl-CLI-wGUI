# audioctl/cmdline_fmt.py
#
# Command-line formatting helpers (display only)
# ---------------------------------------------
# This module exists because we need to display/echo commands in a way that:
#   1) matches Windows expectations (double quotes, CreateProcess rules), and
#   2) is copy/paste friendly into both cmd.exe and PowerShell (PowerShell is picky
#      about device IDs and `{}` tokens unless quoted).
#
# IMPORTANT:
# - These helpers are for *display/logging only*.
# - Do NOT build a shell string and execute it.
#   Always execute subprocesses by passing argv as a list (Popen([...])).
#
# Design notes:
# - `shlex.join()` is POSIX-oriented and uses single quotes; that is confusing on
#   Windows and often not pasteable as-is.
# - On Windows we prefer `subprocess.list2cmdline()` which implements the proper
#   CreateProcess quoting rules and produces double quotes.
# - For "Print CLI commands" (GUI echo), we may intentionally add quotes even when
#   not strictly required by CreateProcess so the command pastes cleanly into
#   PowerShell (e.g., `--id "{...}"`).
import os
import subprocess


def format_cmd_for_display(argv) -> str:
    if argv is None:
        return ""
    args = ["" if a is None else str(a) for a in argv]
    if os.name == "nt":
        return subprocess.list2cmdline(args)
    try:
        import shlex
        return shlex.join(args)
    except Exception:
        return " ".join(args)


def format_audioctl_cmd_for_display(args, *, frozen: bool = False, cross_shell: bool = True) -> str:
    r"""
    Print-friendly audioctl command for copy/paste into cmd.exe OR PowerShell.

    Key behavior:
      - On Windows, prefix with .\\audioctl.exe when frozen.
      - When cross_shell=True, ALWAYS wrap the value after --id in double quotes
        (PowerShell-safe), even if it contains no spaces.
    """
    args = ["" if a is None else str(a) for a in (args or [])]

    if os.name == "nt" and cross_shell:
        prefix = (r".\audioctl.exe" if frozen else "audioctl")

        cooked = []
        i = 0
        while i < len(args):
            a = args[i]
            cooked.append(a)

            if a == "--id" and i + 1 < len(args):
                val = args[i + 1]
                # FORCE quotes even when not required by Windows parsing rules
                cooked.append(f'"{val}"')
                i += 2
                continue

            i += 1

        return prefix + (" " + " ".join(cooked) if cooked else "")

    return "audioctl " + format_cmd_for_display(args)


