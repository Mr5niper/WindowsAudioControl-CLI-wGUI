# Command-line formatting helpers (display only)
# ---------------------------------------------
# This module formats commands for display/logging purposes only.
#
# Goals:
#   1) Match Windows CreateProcess quoting rules (double quotes).
#   2) Be copy/paste friendly for both cmd.exe and PowerShell.
#
# IMPORTANT:
# - These helpers are for display/logging only.
# - Do NOT execute a constructed string.
#   Always execute subprocesses using argv lists (Popen([...])).
#
# Design notes:
# - We use subprocess.list2cmdline() to apply proper Windows quoting rules.
# - PowerShell treats `{}` specially (script blocks), so GUID-like tokens
#   must be double-quoted for safe paste.
# - Additional quoting here is for display compatibility, not execution.

import sys
import subprocess
import re

def format_cmd_for_display(argv) -> str:
    """
    Used by gui.py for the debug log.
    Standard Windows formatting then ensures IDs are double-quoted.
    """
    if argv is None: return ""
    args = ["" if a is None else str(a) for a in argv]
    
    # 1. Use standard Windows rules (handles spaces in FX/Names)
    cmd_str = subprocess.list2cmdline(args)
    
    # 2. Swap single quotes to double quotes for the "Look" you want
    cmd_str = cmd_str.replace("'", '"')

    # 3. Ensure {GUID} IDs are always quoted for PowerShell safety
    guid_pattern = r'({[0-9A-Fa-f.-]+}(?:\.{[0-9A-Fa-f.-]+})?)'
    return re.sub(r'(?<!")' + guid_pattern + r'(?!")', r'"\1"', cmd_str)

def format_audioctl_cmd_for_display(args, *, frozen: bool = False, cross_shell: bool = True) -> str:
    r"""
    Used by gui.py for the copy/paste console.
    Handles the .\audioctl.exe or python -m prefix.
    """
    args_list = ["" if a is None else str(a) for a in (args or [])]
    
    # Handle the prefix correctly based on frozen state
    prefix = r".\audioctl.exe" if frozen else f"{sys.executable} -m audioctl"

    # Reuse the clean logic
    cmd_str = format_cmd_for_display(args_list)

    return f"{prefix} {cmd_str}".strip()
    
    

