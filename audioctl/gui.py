# audioctl/gui.py
#
# GUI philosophy (important):
# ---------------------------
# This Tkinter GUI is intentionally a *thin front-end* over the CLI.
# It shells out to `audioctl` commands as subprocesses and consumes JSON.
#
# Why we do this instead of importing/using the COM/registry layer directly:
# - COM stability / lifetime: pycaw+comtypes and raw vtable PropertyStore calls are
#   sensitive to COM apartment lifetime and to Python GC timing. The CLI wraps each
#   low-level operation in its own COM init/cleanup, so the GUI can stay "pure Tk"
#   and avoid long-lived COM objects in the UI process.
# - Single source of truth: CLI behavior and JSON schemas are the contract used
#   both by scripts and the GUI. The GUI stays in sync by calling the same commands.
# - Frozen vs source mode: in a PyInstaller build `sys.executable` is the bundled
#   exe. In source mode we invoke `python -m audioctl`. Centralizing the subprocess
#   calls prevents a lot of environment-specific bugs.
# ---------------------------
import sys
import io
import time
import tkinter as tk
from tkinter import ttk, messagebox
from contextlib import redirect_stderr
import re
import subprocess
import json
import shlex
import os
from .logging_setup import resource_path, _log, _log_exc, _log_path
from .compat import is_admin
# GUI uses only CLI commands; it does not import low-level device helpers directly.
# This keeps COM usage and raw registry/vtable interactions confined to the CLI layer.
from .vendor_db import (
    _vendor_ini_default_path,
    _load_vendor_db_split,  # NEW
)
# --- BEGIN: Non-blocking Learn runner (main Enhancements) ---
import threading
def _build_cli_cmd(args_list):
    # Build the command line used to run the CLI.
    #
    # Frozen exe (PyInstaller):
    #   - sys.executable is the packaged audioctl.exe; we run it directly.
    #
    # Source mode:
    #   - run module form (`python -m audioctl`) to ensure imports resolve the same
    #     way as normal CLI usage and to avoid path issues.
    #
    # NOTE: Centralizing this logic is important because the GUI must behave the
    # same in both environments.
    # Same behavior as run_audioctl for building the command
    if getattr(sys, "frozen", False):
        return [sys.executable] + args_list
    else:
        return [sys.executable, "-m", "audioctl"] + args_list
class LearnRunner:
    # LearnRunner solves a very specific GUI problem:
    #
    # The CLI "learn" flows are interactive (they print a prompt and wait on stdin).
    # If we ran them synchronously in the Tk thread, the UI would freeze.
    #
    # So we:
    #   - run the CLI as a subprocess in the background,
    #   - stream its stdout/stderr without blocking,
    #   - detect prompt text and expose "Continue" buttons,
    #   - optionally auto-confirm the CLI safety prompt (I UNDERSTAND) when allowed.
    #
    # Implementation note:
    # We read the streams character-by-character (bufsize=0) because prompts may be
    # printed without newline terminators; line-based reads would miss them.
    PATTERN_CONFIRM = re.compile(r"I UNDERSTAND")  # literal appears in the prompt text
    # Snapshot prompt markers printed by the CLI during learn:
    # - "snapshot A" is the "enabled" capture point
    # - "snapshot B" is the "disabled" capture point
    PATTERN_A = re.compile(r"When ready, press Enter to capture snapshot A", re.IGNORECASE)
    PATTERN_B = re.compile(r"When ready, press Enter to capture snapshot B", re.IGNORECASE)
    def __init__(self, args_list, on_output, on_state, confirmed=False):
        self.args_list = args_list
        self.on_output = on_output or (lambda _t: None)
        self.on_state = on_state or (lambda _s: None)
        # confirmed=True means we set AUDIOCTL_LEARN_CONFIRMED=1 so the CLI skips its
        # "type I UNDERSTAND" prompt. The GUI shows its own warning UI instead.
        self.confirmed = bool(confirmed)
        self.proc = None
        # Internal state flags that tell the GUI which "Continue" button should be enabled.
        self._waiting_a = False
        self._waiting_b = False
        # Auto-confirm is only sent once (first prompt occurrence) to avoid spamming stdin.
        self._sent_confirm = False
        # Collected output is kept so we can parse the final JSON summary after the process exits.
        self.collected_out = []
        self.collected_err = []
    def start(self):
        # Start the subprocess in a way that supports interactive stdin.
        # We keep stdout/stderr piped so we can:
        #  - display learn progress,
        #  - detect prompts,
        #  - parse the final JSON object emitted by the CLI.
        env = os.environ.copy()
        if self.confirmed:
            env["AUDIOCTL_LEARN_CONFIRMED"] = "1"
        cmd = _build_cli_cmd(self.args_list)
        self.proc = subprocess.Popen(
            cmd,
            stdin=subprocess.PIPE,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            encoding="utf-8",
            errors="replace",
            bufsize=0,  # unbuffered -> we read char-wise
            env=env,
        )
        self.on_state("started")
        # Read stdout/stderr concurrently so neither pipe can block the child process.
        threading.Thread(target=self._read_stream, args=(self.proc.stdout, True), daemon=True).start()
        threading.Thread(target=self._read_stream, args=(self.proc.stderr, False), daemon=True).start()
        threading.Thread(target=self._waiter, daemon=True).start()
    def _waiter(self):
        # Wait for completion off the Tk thread; report a simple state back to GUI.
        rc = self.proc.wait()
        self.on_state("done" if rc == 0 else "error")
    def _read_stream(self, stream, is_stdout):
        # Read stream one character at a time so we can detect prompts even if the CLI
        # prints them without a trailing newline.
        buf = ""
        while True:
            ch = stream.read(1)
            if ch is None or ch == "":
                break
            buf += ch
            # Emit whole lines as they complete
            if "\n" in buf:
                lines = buf.split("\n")
                for ln in lines[:-1]:
                    self._handle_text(ln + "\n", is_stdout)
                buf = lines[-1]
            else:
                # still scan partial text for prompts that don't end with newline
                self._scan_for_prompts(buf)
        if buf:
            self._handle_text(buf, is_stdout)
    def _handle_text(self, text, is_stdout):
        # Store output for later parsing and forward it to the GUI's output callback.
        try:
            if is_stdout:
                self.collected_out.append(text)
            else:
                self.collected_err.append(text)
        except Exception:
            pass
        self.on_output(text)
        self._scan_for_prompts(text)
    def _scan_for_prompts(self, text):
        t = text if isinstance(text, str) else str(text or "")
        # Auto-confirm (only on first attempt)
        # This is the CLI safety gate in learn modes ("type I UNDERSTAND").
        # The GUI uses this to avoid blocking the workflow if the user already
        # confirmed via the GUI warning dialog.
        if (not self.confirmed) and (not self._sent_confirm) and self.PATTERN_CONFIRM.search(t):
            try:
                self.proc.stdin.write("I UNDERSTAND\n")
                self.proc.stdin.flush()
                self._sent_confirm = True
            except Exception:
                pass
            return
        # Snapshot A prompt
        if self.PATTERN_A.search(t):
            self._waiting_a = True
            self.on_state("waiting_snapshot_a")
            return
        # Snapshot B prompt
        if self.PATTERN_B.search(t):
            self._waiting_b = True
            self.on_state("waiting_snapshot_b")
            return
    def continue_snapshot_a(self):
        # Simulate pressing Enter at the "capture snapshot A" prompt.
        if self._waiting_a and self.proc and self.proc.stdin:
            try:
                self.proc.stdin.write("\n")
                self.proc.stdin.flush()
            except Exception:
                pass
            self._waiting_a = False
    def continue_snapshot_b(self):
        # Simulate pressing Enter at the "capture snapshot B" prompt.
        if self._waiting_b and self.proc and self.proc.stdin:
            try:
                self.proc.stdin.write("\n")
                self.proc.stdin.flush()
            except Exception:
                pass
            self._waiting_b = False
    def terminate(self):
        # Best-effort cancellation hook used when the user closes the learn UI.
        try:
            if self.proc and self.proc.poll() is None:
                self.proc.terminate()
        except Exception:
            pass
# --- END: Non-blocking Learn runner ---
def run_audioctl(args_list, capture_json=False, expect_ok=True):
    """
    Run 'audioctl' CLI as a subprocess.

    Why subprocess:
    - The CLI is our authority for COM/registry interaction; it manages COM init/cleanup per call.
    - The GUI stays stable by avoiding direct COM usage in the Tk process.

    args_list: list of strings, e.g. ["list", "--json"]
    capture_json: if True, parse stdout as JSON and return the object
    expect_ok: if True, raise RuntimeError on non-zero exit codes
    Returns:
      - If capture_json=False: (rc, stdout_text, stderr_text)
      - If capture_json=True: parsed JSON object
    """
    # Frozen build vs source:
    # - In a PyInstaller build, sys.executable is audioctl.exe and should be run directly.
    # - In source/dev, run `python -m audioctl` so imports resolve in-package.
    if getattr(sys, "frozen", False):
        exe = sys.executable
        cmd = [exe] + args_list
    else:
        exe = sys.executable
        cmd = [exe, "-m", "audioctl"] + args_list
    try:
        from .logging_setup import _dbg
        _dbg(f"GUI run_audioctl: {shlex.join(cmd)}")
    except Exception:
        pass
    # stdout/stderr are captured so the GUI can:
    # - parse JSON from stdout
    # - display meaningful error details from stderr
    proc = subprocess.Popen(
        cmd,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        encoding="utf-8",
        errors="replace"
    )
    out, err = proc.communicate()
    rc = proc.returncode
    if capture_json:
        try:
            data = json.loads(out or "{}")
        except Exception as e:
            if expect_ok:
                raise RuntimeError(
                    f"Failed to parse JSON from audioctl: {e}\nstdout={out!r}\nstderr={err!r}"
                )
            else:
                return None
        if expect_ok and rc != 0:
            raise RuntimeError(f"audioctl failed with rc={rc}: {err or out}")
        return data
    if expect_ok and rc != 0:
        raise RuntimeError(f"audioctl failed with rc={rc}: {err or out}")
    return rc, out, err
def run_audioctl_quick_json(args_list, timeout=0.75):
    """
    Fast one-shot CLI call with timeout.

    Used primarily for UI freshness (context menu labels). We do *not* want to
    block the UI waiting on a slow/hung driver call; instead we:
      - enforce a small timeout
      - return None on failure
    """
    if getattr(sys, "frozen", False):
        cmd = [sys.executable] + args_list
    else:
        cmd = [sys.executable, "-m", "audioctl"] + args_list
    try:
        p = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=float(timeout),
        )
    except Exception:
        return None
    if p.returncode != 0:
        return None
    try:
        return json.loads(p.stdout or "{}")
    except Exception:
        return None
def run_audioctl_interactive(args_list, prompt_patterns, expect_ok=True):
    """
    Run 'audioctl' CLI as a subprocess, line-by-line.

    This is used for interactive CLI flows where the CLI prints a prompt
    on stdout and waits for Enter on stdin. We:
      - read stdout line-by-line
      - when a known prompt substring appears, show a messagebox to the user
      - then write "\n" to the subprocess stdin to continue

    prompt_patterns: list of (substring, title, custom_message) tuples.
      - substring: text to look for in a stdout line.
      - title: window title for the messagebox.
      - custom_message: if not None, this is shown instead of the raw line.
    When a line matches, we show a messagebox and then send '\n' to stdin.
    Returns (rc, collected_stdout, collected_stderr).
    """
    if getattr(sys, "frozen", False):
        exe = sys.executable
        cmd = [exe] + args_list
    else:
        exe = sys.executable
        cmd = [exe, "-m", "audioctl"] + args_list
    try:
        from .logging_setup import _dbg
        _dbg("GUI run_audioctl_interactive: " + shlex.join(cmd))
    except Exception:
        pass
    # IMPORTANT: open stdin for writing.
    # Without stdin=PIPE we couldn't "press Enter" to advance prompts.
    proc = subprocess.Popen(
        cmd,
        stdin=subprocess.PIPE,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        encoding="utf-8",
        errors="replace",
        bufsize=1,  # line-buffered
    )
    collected_out = []
    collected_err = []
    # Read stdout line by line, handle prompts
    while True:
        line = proc.stdout.readline()
        if line == "":
            break  # EOF
        collected_out.append(line)
        handled = False
        for substring, title, custom_message in prompt_patterns:
            if substring in line:
                # Decide what to show
                msg_text = custom_message if custom_message is not None else line.strip()
                try:
                    messagebox.showinfo(title, msg_text)
                except Exception:
                    pass
                # Simulate pressing Enter (unblocks the CLI input()).
                try:
                    proc.stdin.write("\n")
                except Exception:
                    pass
                try:
                    proc.stdin.flush()
                except Exception:
                    pass
                handled = True
                break  # stop checking other patterns for this line
        # Only prompt-specific handling is needed; other lines are just collected.
    # Read remaining stdout/stderr
    remaining_out, err = proc.communicate()
    if remaining_out:
        collected_out.append(remaining_out)
    if err:
        collected_err.append(err)
    rc = proc.returncode
    out_text = "".join(collected_out)
    err_text = "".join(collected_err)
    if expect_ok and rc != 0:
        raise RuntimeError(f"audioctl interactive failed with rc={rc}: {err_text or out_text}")
    return rc, out_text, err_text
class AudioGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Mr5niper's Audio Control  v1.4.7.3  02-06-2026")
        # Style and theme
        style = ttk.Style(self.root)
        try:
            for theme in ("vista", "xpnative", "clam", "default"):
                if theme in style.theme_names():
                    style.theme_use(theme)
                    break
        except Exception:
            pass
        try:
            from tkinter import font as tkfont
            base_font = tkfont.nametofont(self.root.cget("font"))
            try:
                heading_font = base_font.copy()
                heading_font.configure(weight="bold")
            except Exception:
                heading_font = base_font
            style.configure("Treeview.Heading", font=heading_font)
        except Exception:
            pass
        try:
            style.configure("Treeview.Heading", relief="flat")
            style.map("Treeview.Heading", background=[], relief=[], foreground=[])
        except Exception:
            pass
        # Variables / core GUI state:
        # - include_all: whether to include disabled/disconnected endpoints in list output.
        # - print_cmd: whether to echo the equivalent CLI commands for actions.
        # - devices: current list from `audioctl list --json`.
        # - item_to_device: Treeview item id -> device dict (used to resolve selection).
        self.include_all = tk.BooleanVar(value=False)
        self.print_cmd = tk.BooleanVar(value=False)
        self.devices = []
        self.item_to_device = {}
        # NEW: cache for per-device state (get-device-state JSON)
        # Key: device_id, Value: state dict from CLI
        #
        # This keeps the context menu responsive: we can render "Mute vs Unmute",
        # "Enable vs Disable Listen", Enhancements state, and FX list without
        # a blocking subprocess call on every right click.
        self.device_state_cache = {}
        # NEW: flag to suppress auto-refresh while in our own modal workflows
        # (learn flows involve interactive prompts; refreshing during them is noisy and risky).
        self._in_modal_operation = False
        # Layout
        self.container = ttk.Frame(self.root, padding=10)
        self.container.pack(fill="both", expand=True)
        self.topbar = ttk.Frame(self.container)
        self.topbar.pack(fill="x", pady=(0, 8))
        refresh_btn = ttk.Button(self.topbar, text="Refresh", command=self.refresh_devices)
        refresh_btn.pack(side="left")
        ttk.Checkbutton(
            self.topbar,
            text="Show disabled/disconnected",
            variable=self.include_all,
            command=self.refresh_devices
        ).pack(side="left", padx=(10, 0))
        self.chk_print_cmd = ttk.Checkbutton(
            self.topbar,
            text="Print CLI commands",
            variable=self.print_cmd
        )
        self.chk_print_cmd.pack(side="left", padx=(10, 0))
        # Hard runtime suppression flag used during learn flows to keep console output clean.
        self._suppress_cli_prints = False
        if not is_admin():
            admin_lbl = ttk.Label(self.topbar, text="Note: Some actions may require Administrator", foreground="#CC6600")
            admin_lbl.pack(side="right")
        # Treeview
        # We insert group rows ("Playback", "Recording") and then device rows under them.
        # Group rows are treated as non-selectable to avoid mis-targeted actions.
        columns = ("Index", "Name", "Flow", "Defaults", "ID")
        self.tree = ttk.Treeview(
            self.container,
            columns=columns,
            show="tree headings",
            selectmode="browse",
            height=10
        )
        self.tree.heading("#0", text="")
        for col in columns:
            self.tree.heading(col, text=col)
        self.tree.column("#0", width=120, minwidth=100, anchor="e", stretch=False)
        self.tree.column("Index", width=60, minwidth=50, anchor="e", stretch=False)
        self.tree.column("Name",  width=200, minwidth=200, anchor="w", stretch=True)
        self.tree.column("Flow",  width=80,  minwidth=80,  anchor="w", stretch=False)
        self.tree.column("Defaults", width=160, minwidth=160, anchor="w", stretch=False)
        self.tree.column("ID",    width=260, minwidth=240, anchor="w", stretch=False)
        self.tree["displaycolumns"] = ("Index", "Name", "Flow", "Defaults", "ID")
        self.tree.pack(fill="both", expand=True)
        # Scrollbar
        self.yscroll = ttk.Scrollbar(self.tree, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=self.yscroll.set)
        self.yscroll.pack(side="right", fill="y")
        # Group tags
        try:
            self.tree.tag_configure("group", foreground="#202020")
        except Exception:
            pass
        # Remove indicator element (cosmetic): we don't want expand/collapse affordances
        # for group rows; the view is always grouped and open by default.
        try:
            style.layout("Treeview.Item", [
                ('Treeitem.padding', {
                    'sticky': 'nswe',
                    'children': [
                        ('Treeitem.image', {'side': 'left', 'sticky': ''}),
                        ('Treeitem.focus', {'side': 'left', 'sticky': 'nswe', 'children': [
                            ('Treeitem.text', {'side': 'left', 'sticky': ''})
                        ]})
                    ]
                })
            ])
        except Exception:
            pass
        # Status bar
        self.status = tk.StringVar(value="Ready")
        self.statusbar = ttk.Label(self.root, textvariable=self.status, anchor="w", padding=(10, 3))
        self.statusbar.pack(fill="x", side="bottom")
        # Context menu
        self.menu = tk.Menu(self.root, tearoff=0)
        self.menu.add_command(label="Set as Default (all roles)", command=self.on_set_default)
        self.menu.add_separator()
        self.menu.add_command(label="Set Volume...", command=self.on_set_volume)
        self.menu.add_command(label="Mute/Unmute", command=self.on_toggle_mute)
        # We store indices of dynamic items so we can update labels later with entryconfig()
        # without rebuilding the entire menu.
        self.mute_menu_index = self.menu.index("end")
        self.menu.add_separator()
        self.listen_menu_default_label = "Toggle Listen (capture only)"
        self.menu.add_command(label=self.listen_menu_default_label, command=self.on_toggle_listen)
        self.listen_menu_index = self.menu.index("end")
        self.enh_menu_default_label = "Enable Enhancements"
        self.menu.add_command(label=self.enh_menu_default_label, command=self.on_toggle_enhancements)
        self.enh_menu_index = self.menu.index("end")
        # Enhancement Effects cascade submenu (populated via CLI)
        self.fx_menu = tk.Menu(self.menu, tearoff=0)
        self.menu.add_cascade(label="Enhancement Effects", menu=self.fx_menu)
        self.fx_cascade_index = self.menu.index("end")
        # Learn Enhancements remains after the cascade
        self.menu.add_command(label="Learn Enhancements", command=self.on_learn_enhancements)
        self.learn_menu_index = self.menu.index("end")
        # Track dynamically added FX menu items
        self._dynamic_fx_menu_items = []
        # _pending_enh remembers what enhancements state/support we believed when building the menu.
        # This avoids confusing "toggle" logic if the user clicks after state changed.
        self._pending_enh = None
        # Prevent overlapping menu builds (avoid racing CLI calls)
        self._menu_build_in_progress = False
        # Bindings
        self.tree.bind("<Button-3>", self.on_right_click)
        self.tree.bind("<ButtonRelease-1>", self.on_left_release)
        self.tree.bind("<Double-1>", self.on_double_click)
        self.root.bind("<F5>", lambda e: self.refresh_devices())
        self.tree.bind("<Button-1>", self.on_left_click, add="+")
        self.tree.bind("<<TreeviewSelect>>", self.on_select_change)
        # Optional: auto-refresh when window regains focus
        # self.root.bind("<FocusIn>", self.on_focus_in)
        # Initial load
        self.refresh_devices()
        self.root.after_idle(self.adjust_layout_to_content)
    def is_group_row(self, iid):
        return iid not in self.item_to_device
    def on_left_click(self, event):
        # Prevent selecting group header rows and prevent header-click behaviors.
        region = self.tree.identify_region(event.x, event.y)
        if region == "heading":
            return "break"
        iid = self.tree.identify_row(event.y)
        if not iid:
            return
        if self.is_group_row(iid):
            return "break"
    def on_select_change(self, event):
        # If a group row gets selected (e.g., via keyboard navigation), redirect
        # selection to the first child device row, since only devices are actionable.
        sel = self.tree.selection()
        if not sel:
            return
        iid = sel[0]
        if self.is_group_row(iid):
            children = self.tree.get_children(iid)
            if children:
                self.tree.selection_set(children[0])
                self.tree.focus(children[0])
            else:
                self.tree.selection_remove(iid)
    def on_focus_in(self, event):
        """
        Optional: when the main window regains focus, refresh devices & state.
        Skips if we're already in a modal/learn operation to avoid recursion.
        """
        if getattr(self, "_in_modal_operation", False):
            return
        try:
            self.refresh_devices()
        except Exception:
            pass
    def set_status(self, text):
        self.status.set(text)
        try:
            print(text)
        except Exception:
            pass
    def refresh_devices(self):
        try:
            from .logging_setup import _dbg
            _dbg("GUI: refresh_devices begin")
            # Phase 1: pull the device list via CLI to keep the GUI COM-free.
            # Build CLI args: audioctl list [--all] --json
            args = ["list", "--json"]
            if self.include_all.get():
                args.insert(1, "--all")
            data = run_audioctl(args, capture_json=True, expect_ok=True)
            self.devices = data.get("devices", [])
            # Clear tree
            self.item_to_device.clear()
            for item in self.tree.get_children():
                self.tree.delete(item)
            # Split by flow for display
            render_devs = sorted(
                [d for d in self.devices if d["flow"] == "Render"],
                key=lambda x: x["name"].lower()
            )
            capture_devs = sorted(
                [d for d in self.devices if d["flow"] == "Capture"],
                key=lambda x: x["name"].lower()
            )
            grp_render = self.tree.insert(
                "", "end", text="Playback (Render)",
                values=("", "", "", "", ""), open=True, tags=("group",)
            )
            grp_capture = self.tree.insert(
                "", "end", text="Recording (Capture)",
                values=("", "", "", "", ""), open=True, tags=("group",)
            )
            def insert_group(parent, devs, flow_name):
                for idx, d in enumerate(devs):
                    flags = [k for k, v in d["isDefault"].items() if v]
                    defaults_txt = ", ".join(flags) if flags else "-"
                    d_copy = dict(d)
                    d_copy["_index"] = idx
                    d_copy["_group"] = flow_name
                    iid = self.tree.insert(
                        parent, "end", text="",
                        values=(idx, d["name"], d["flow"], defaults_txt, d["id"])
                    )
                    self.item_to_device[iid] = d_copy
            insert_group(grp_render, render_devs, "Render")
            insert_group(grp_capture, capture_devs, "Capture")
            self.set_status("Device list updated")
            self.adjust_layout_to_content()
            self.root.after_idle(self.adjust_layout_to_content)
            _dbg("GUI: refresh_devices end (list/tree built)")
            # Phase 2: background-fill state cache via get-device-state.
            # We do this incrementally (one device at a time) so the UI remains responsive.
            # NEW: start incremental background state population
            self._schedule_state_population()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to list devices:\n{e}")
            self.set_status("Failed to refresh devices")
    def _schedule_state_population(self):
        """
        Start or restart incremental population of device_state_cache.

        Why incremental:
        - get-device-state can do COM reads and registry probes; running it for every
          device synchronously would freeze the UI.
        - Tk's event loop stays responsive when we schedule one device per tick.

        Note: We clear the cache on each refresh so we never mix old state from devices
        that may have disappeared or changed.
        """
        # Reset cache and queue
        self.device_state_cache.clear()
        # Simple queue of (id, flow)
        self._state_queue = [(d["id"], d["flow"]) for d in self.devices]
        # Kick off first step
        self.root.after(10, self._populate_next_device_state)
    def _populate_next_device_state(self):
        """
        Process one device from the queue:
        - call get-device-state and cache it
        - reschedule until queue is empty
        """
        q = getattr(self, "_state_queue", None)
        if not q:
            self._state_queue = []
            return
        dev_id, flow = q.pop(0)
        try:
            st = run_audioctl(
                ["get-device-state", "--id", dev_id, "--flow", flow],
                capture_json=True,
                expect_ok=False,
            )
        except Exception:
            st = None
        if isinstance(st, dict):
            self.device_state_cache[dev_id] = st
        self.root.after(10, self._populate_next_device_state)
    def adjust_layout_to_content(self):
        self.root.update_idletasks()
        try:
            from tkinter import font as tkfont
            tv_font_name = self.tree.cget("font") or "TkDefaultFont"
            tv_font = tkfont.nametofont(tv_font_name)
        except Exception:
            tv_font = None
        names = [d["name"] for d in self.devices] or ["Name"]
        defaults_list = []
        for d in self.devices:
            flags = [k for k, v in d["isDefault"].items() if v]
            defaults_list.append(", ".join(flags) if flags else "-")
        if not defaults_list:
            defaults_list = ["-"]
        ids = [d["id"] for d in self.devices] or ["ID"]
        longest_name = max(names, key=len)
        longest_defaults = max(defaults_list, key=len)
        longest_id = max(ids, key=len)
        group_labels = ["Playback (Render)", "Recording (Capture)"]
        longest_group = max(group_labels, key=len)
        render_count = sum(1 for d in self.devices if d["flow"] == "Render")
        capture_count = sum(1 for d in self.devices if d["flow"] == "Capture")
        max_index_value = max(render_count - 1, capture_count - 1, 0)
        pad = 32
        def measure(text, fallback):
            try:
                return tv_font.measure(text) if tv_font else fallback
            except Exception:
                return fallback
        group_w = max(100, min(180, measure(longest_group, 140) + 12))
        name_w    = max(240, min(700, max(measure(longest_name, 300), measure("Name", 60)) + pad))
        flow_w    = max(80, max(measure("Recording", 90), measure("Flow", 60)) + 30)
        defaults_w= max(160, min(480, max(measure(longest_defaults, 240), measure("Defaults", 100)) + pad))
        id_w      = max(240, min(560, max(measure(longest_id, 340), measure("ID", 60)) + pad))
        index_digits = max(2, len(str(max_index_value)))
        index_w = max(60, measure("9" * index_digits, 30) + 24)
        try:
            self.tree.column("#0", width=int(group_w), minwidth=140, anchor="w", stretch=False)
            self.tree.column("Index", width=int(index_w), minwidth=50, anchor="e", stretch=False)
            self.tree.column("Name",  width=int(name_w),  minwidth=200, anchor="w", stretch=True)
            self.tree.column("Flow",  width=int(flow_w),  minwidth=80,  anchor="w", stretch=False)
            self.tree.column("Defaults", width=int(defaults_w), minwidth=160, anchor="w", stretch=False)
            self.tree.column("ID",    width=int(id_w),    minwidth=240, anchor="w", stretch=False)
        except Exception:
            pass
        rows = len(self.devices) + 4 if self.devices else 4
        self.tree.configure(height=min(max(rows, 6), 50))
        self.root.update_idletasks()
        try:
            sb_w = max(self.yscroll.winfo_reqwidth(), 16) if self.yscroll else 16
        except Exception:
            sb_w = 16
        total_cols = int(group_w + index_w + name_w + flow_w + defaults_w + id_w + sb_w + 40)
        desired_w = max(total_cols, self.container.winfo_reqwidth() + 10, 600)
        desired_h = max(self.root.winfo_reqheight(), 325)
        scr_w = self.root.winfo_screenwidth()
        scr_h = self.root.winfo_screenheight()
        margin = 80
        w = min(desired_w, scr_w - margin)
        h = min(desired_h, scr_h - margin)
        self.root.geometry(f"{int(w)}x{int(h)}")
        self.root.minsize(int(min(w, scr_w - margin)), int(min(h, scr_h - margin)))
    def maybe_print_cli(self, cmd_str: str):
        # Print the equivalent CLI command for a GUI action.
        #
        # We suppress printing during learn flows because learn is interactive and
        # already produces a lot of prompt output; echoing commands mid-flow is noisy
        # and can confuse users (especially when run from a console).
        # Hard suppression: never print while a learn flow is running
        if getattr(self, "_suppress_cli_prints", False):
            return
        if self.print_cmd.get():
            try:
                print(cmd_str)
            except Exception:
                pass
    def _suspend_print_cli_and_disable_checkbox(self):
        # During learn flows, we temporarily disable "Print CLI commands" and force it off.
        # This keeps interactive output clean and prevents command spam while the GUI is
        # driving stdin prompts.
        # Save current checkbox value
        try:
            self._prev_print_cmd_val = bool(self.print_cmd.get())
        except Exception:
            self._prev_print_cmd_val = False
        # Force printing OFF and disable the checkbox
        try:
            self.print_cmd.set(False)
        except Exception:
            pass
        try:
            if hasattr(self, "chk_print_cmd") and self.chk_print_cmd:
                self.chk_print_cmd.configure(state="disabled")
        except Exception:
            pass
        # HARD suppression (blocks any attempt to print via maybe_print_cli)
        self._suppress_cli_prints = True
    def _restore_print_cli_checkbox(self):
        # Restore checkbox and printing state after learn completes/cancels.
        # Restore checkbox and printing state
        try:
            self.print_cmd.set(bool(getattr(self, "_prev_print_cmd_val", False)))
        except Exception:
            pass
        try:
            if hasattr(self, "chk_print_cmd") and self.chk_print_cmd:
                self.chk_print_cmd.configure(state="normal")
        except Exception:
            pass
        # Lift hard suppression
        self._suppress_cli_prints = False
    def get_selected_device(self):
        sel = self.tree.selection()
        if not sel:
            return None
        return self.item_to_device.get(sel[0])
    def _ensure_device_state_entry(self, dev_id, flow):
        """
        Return a mutable state dict for a device in device_state_cache.

        Why:
        After we perform an action (mute/listen/enhancements/fx), we update the cache
        immediately so the next context menu open shows correct labels even before
        the next background refresh completes.
        """
        st = self.device_state_cache.get(dev_id)
        if not isinstance(st, dict):
            st = {
                "id": dev_id,
                "flow": flow,
                "volume": None,
                "muted": None,
                "listenEnabled": None,
                "enhancementsEnabled": None,
                "availableFX": [],
            }
            self.device_state_cache[dev_id] = st
        # Ensure availableFX is a list
        if not isinstance(st.get("availableFX"), list):
            st["availableFX"] = []
        return st
    def _refresh_menu_state_async(self, d):
        # This is an asynchronous menu refresh helper (not always used).
        # It exists to avoid blocking UI while still updating menu labels based on
        # the CLI's aggregated `get-device-state` command.
        # d: device dict for the selection
        def worker():
            try:
                st = run_audioctl(
                    ["get-device-state", "--id", d["id"], "--flow", d["flow"]],
                    capture_json=True,
                    expect_ok=False,
                )
            except Exception:
                st = None
            if not isinstance(st, dict):
                return
            def apply_updates():
                try:
                    # Cache
                    self.device_state_cache[d["id"]] = st
                    # Listen label (Capture only)
                    if d["flow"] == "Capture":
                        ls = st.get("listenEnabled", None)
                        listen_label = "Disable Listen" if ls is True else "Enable Listen" if ls is False else self.listen_menu_default_label
                        self.menu.entryconfig(self.listen_menu_index, label=listen_label, state="normal")
                    # Enhancements (vendor-only): if None, there is no learned vendor method and we disable the action.
                    enh = st.get("enhancementsEnabled", None)
                    if enh is True:
                        self.menu.entryconfig(self.enh_menu_index, label="Disable Enhancements", state="normal")
                        self._pending_enh = {"id": d["id"], "flow": d["flow"], "current": True, "supported": True}
                    elif enh is False:
                        self.menu.entryconfig(self.enh_menu_index, label="Enable Enhancements", state="normal")
                        self._pending_enh = {"id": d["id"], "flow": d["flow"], "current": False, "supported": True}
                    else:
                        self.menu.entryconfig(self.enh_menu_index, label="Enhancements (no vendor toggle)", state="disabled")
                        self._pending_enh = {"id": d["id"], "flow": d["flow"], "current": None, "supported": False}
                    # FX submenu: rebuild with latest states (INI-driven effect list with per-effect states)
                    try:
                        self.fx_menu.delete(0, "end")
                    except Exception:
                        pass
                    fx_list = st.get("availableFX") or []
                    if fx_list:
                        fx_list = sorted(fx_list, key=lambda x: (x.get("fx_name") or "").lower())
                        for fx in fx_list:
                            fx_name = fx.get("fx_name")
                            state_fx = fx.get("state")
                            if state_fx is True:
                                label = f"Disable {fx_name}"
                            elif state_fx is False:
                                label = f"Enable {fx_name}"
                            else:
                                label = f"Toggle {fx_name}"
                            def make_cmd(name, cur):
                                def cmd():
                                    self.on_toggle_fx_live(name, cur)
                                return cmd
                            self.fx_menu.add_command(label=label, command=make_cmd(fx_name, state_fx))
                        self.menu.entryconfig(self.fx_cascade_index, state="normal")
                    else:
                        self.fx_menu.add_command(label="No effects available", state="disabled")
                        self.menu.entryconfig(self.fx_cascade_index, state="disabled")
                except Exception:
                    pass
            self.root.after(0, apply_updates)
        threading.Thread(target=worker, daemon=True).start()
    def show_menu_for_item(self, event, iid=None):
        try:
            # Context menu rendering strategy:
            # 1) Use cached device_state_cache first so menu labels appear immediately.
            # 2) Then do a single fast `get-device-state` call (timeout-limited) to
            #    refresh labels right before showing the menu (reduces stale labels).
            #
            # Prevent overlapping menu builds caused by rapid clicks
            if self._menu_build_in_progress:
                return
            self._menu_build_in_progress = True
            if iid is None:
                iid = self.tree.identify_row(event.y)
            if not iid:
                return
    
            d = self.item_to_device.get(iid)
            if d:
                self.tree.selection_set(iid)
            else:
                self.tree.selection_remove(iid)
    
            end_idx = self.menu.index("end")
            end_idx = end_idx if end_idx is not None else -1
    
            if not d:
                for i in range(end_idx + 1):
                    etype = self.menu.type(i)
                    if etype in ("command", "cascade", "checkbutton", "radiobutton"):
                        self.menu.entryconfig(i, state="disabled")
                self.menu.tk_popup(event.x_root, event.y_root)
                return
    
            # Enable all standard menu items
            for i in range(end_idx + 1):
                etype = self.menu.type(i)
                if etype in ("command", "cascade", "checkbutton", "radiobutton"):
                    self.menu.entryconfig(i, state="normal")
    
            # Fetch cached full device state instead of calling CLI each time
            state = self.device_state_cache.get(d["id"])
    
            # Mute label: use combined state (muted)
            mute_label = "Mute/Unmute"
            if isinstance(state, dict):
                muted = state.get("muted", None)
                if muted is True:
                    mute_label = "Unmute"
                elif muted is False:
                    mute_label = "Mute"
            self.menu.entryconfig(self.mute_menu_index, label=mute_label, state="normal")
    
            # Listen label (Capture only): use combined state (listenEnabled)
            if d["flow"] == "Capture":
                listen_label = self.listen_menu_default_label
                if isinstance(state, dict):
                    ls = state.get("listenEnabled", None)
                    if ls is True:
                        listen_label = "Disable Listen"
                    elif ls is False:
                        listen_label = "Enable Listen"
                self.menu.entryconfig(self.listen_menu_index, label=listen_label, state="normal")
            else:
                self.menu.entryconfig(self.listen_menu_index, label=self.listen_menu_default_label, state="disabled")
    
            # Main enhancements toggle – vendor-only. If CLI reports None, we disable the action.
            # _pending_enh is stored so the click handler knows what state we based the label on.
            self._pending_enh = None
            enh_label = self.enh_menu_default_label
            current_enh_state = None
            enh_state_from_cli = None
            if isinstance(state, dict):
                enh_state_from_cli = state.get("enhancementsEnabled", None)
            if enh_state_from_cli is True:
                enh_label = "Disable Enhancements"
                enh_state_normalized = True
                enh_state_enabled = True
                enh_menu_state = "normal"
            elif enh_state_from_cli is False:
                enh_label = "Enable Enhancements"
                enh_state_normalized = False
                enh_state_enabled = True
                enh_menu_state = "normal"
            else:
                # No vendor toggle available (CLI reported null): disable menu item.
                enh_label = "Enhancements (no vendor toggle)"
                enh_state_normalized = None
                enh_state_enabled = False
                enh_menu_state = "disabled"
            # Remember state for click-time toggling
            self._pending_enh = {
                "id": d["id"],
                "flow": d["flow"],
                "current": enh_state_normalized,
                "supported": enh_state_enabled,
            }
            self.menu.entryconfig(self.enh_menu_index, label=enh_label, state=enh_menu_state)
    
            # Enhancement Effects submenu via combined state
            try:
                self.fx_menu.delete(0, "end")
            except Exception:
                pass
            fx_list = []
            if isinstance(state, dict):
                fx_list = state.get("availableFX", []) or []
            if fx_list:
                fx_list = sorted(fx_list, key=lambda x: (x.get("fx_name") or "").lower())
                for fx in fx_list:
                    fx_name = fx.get("fx_name")
                    state_fx = fx.get("state")
                    if state_fx is True:
                        label = f"Disable {fx_name}"
                    elif state_fx is False:
                        label = f"Enable {fx_name}"
                    else:
                        label = f"Toggle {fx_name}"
                    def make_fx_command(name, current_state):
                        def cmd():
                            self.on_toggle_fx_live(name, current_state)
                        return cmd
                    self.fx_menu.add_command(label=label, command=make_fx_command(fx_name, state_fx))
                self.menu.entryconfig(self.fx_cascade_index, state="normal")
            else:
                self.fx_menu.add_command(label="No effects available", state="disabled")
                self.menu.entryconfig(self.fx_cascade_index, state="disabled")
    
            # Refresh this device's state BEFORE posting the menu (single fast CLI call).
            # This reduces stale labels if the background cache hasn't populated yet or state changed.
            try:
                if d:
                    st = run_audioctl_quick_json(
                        ["get-device-state", "--id", d["id"], "--flow", d["flow"]],
                        timeout=0.75
                    )
                    if isinstance(st, dict):
                        # cache it
                        self.device_state_cache[d["id"]] = st
                        # Listen label (Capture only)
                        if d["flow"] == "Capture":
                            ls = st.get("listenEnabled", None)
                            listen_label = "Disable Listen" if ls is True else "Enable Listen" if ls is False else self.listen_menu_default_label
                            self.menu.entryconfig(self.listen_menu_index, label=listen_label, state="normal")
                        # Enhancements
                        enh = st.get("enhancementsEnabled", None)
                        if enh is True:
                            self.menu.entryconfig(self.enh_menu_index, label="Disable Enhancements", state="normal")
                            self._pending_enh = {"id": d["id"], "flow": d["flow"], "current": True, "supported": True}
                        elif enh is False:
                            self.menu.entryconfig(self.enh_menu_index, label="Enable Enhancements", state="normal")
                            self._pending_enh = {"id": d["id"], "flow": d["flow"], "current": False, "supported": True}
                        else:
                            self.menu.entryconfig(self.enh_menu_index, label="Enhancements (no vendor toggle)", state="disabled")
                            self._pending_enh = {"id": d["id"], "flow": d["flow"], "current": None, "supported": False}
                        # FX submenu (rebuild)
                        try:
                            self.fx_menu.delete(0, "end")
                        except Exception:
                            pass
                        fx_list = st.get("availableFX") or []
                        if fx_list:
                            fx_list = sorted(fx_list, key=lambda x: (x.get("fx_name") or "").lower())
                            for fx in fx_list:
                                fx_name = fx.get("fx_name")
                                state_fx = fx.get("state")
                                if state_fx is True:
                                    label = f"Disable {fx_name}"
                                elif state_fx is False:
                                    label = f"Enable {fx_name}"
                                else:
                                    label = f"Toggle {fx_name}"
                                def make_cmd(name, cur):
                                    def cmd():
                                        self.on_toggle_fx_live(name, cur)
                                    return cmd
                                self.fx_menu.add_command(label=label, command=make_cmd(fx_name, state_fx))
                            self.menu.entryconfig(self.fx_cascade_index, state="normal")
                        else:
                            self.fx_menu.add_command(label="No effects available", state="disabled")
                            self.menu.entryconfig(self.fx_cascade_index, state="disabled")
            except Exception:
                pass
            self.menu.tk_popup(event.x_root, event.y_root)
    
        except Exception as e:
            try:
                from .logging_setup import _log_exc, _log
                _log(f"Context menu error for selection: {e!r}")
                _log_exc("RIGHT-CLICK CONTEXT MENU EXCEPTION")
            except Exception:
                pass
            try:
                messagebox.showerror("Error", f"Failed to build menu:\n{e}")
            except Exception:
                pass
            self.set_status("Failed to build menu")
        finally:
            try:
                self.menu.grab_release()
            except Exception:
                pass
            # Allow next menu build
            self._menu_build_in_progress = False
    def on_right_click(self, event):
        iid = self.tree.identify_row(event.y)
        if not iid:
            return
        if not self.is_group_row(iid):
            self.tree.selection_set(iid)
        self.show_menu_for_item(event, iid=iid)
    def on_left_release(self, event):
        iid = self.tree.identify_row(event.y)
        if not iid:
            return
        if self.is_group_row(iid):
            self.tree.selection_remove(iid)
            return
        self.tree.selection_set(iid)
    def on_double_click(self, event):
        iid = self.tree.identify_row(event.y)
        if not iid:
            return
        if self.is_group_row(iid):
            self.tree.item(iid, open=not bool(self.tree.item(iid, "open")))
            self.tree.selection_remove(iid)
            return
        # Schedule menu build slightly later to avoid racing with other events
        def _open_menu_later():
            if not self.tree.exists(iid):
                return
            self.show_menu_for_item(event, iid=iid)
        self.root.after(50, _open_menu_later)
    def on_set_default(self):
        # set-default may require admin depending on system policy/driver.
        # After success, we refresh the list so default flags update immediately.
        d = self.get_selected_device()
        if not d:
            return
        try:
            from .logging_setup import _log, _dbg
            _dbg(f"GUI: on_set_default for id={d['id']} flow={d['flow']}")
            if d["flow"] == "Render":
                args = ["set-default", "--playback-id", d["id"], "--playback-role", "all"]
            else:
                args = ["set-default", "--recording-id", d["id"], "--recording-role", "all"]
            if not is_admin():
                if not messagebox.askyesno(
                    "Administrator recommended",
                    "Setting default device may require Administrator privileges on some systems.\n\nContinue?"
                ):
                    _log(f"GUI action: set-default cancelled id={d['id']} name={d['name']}")
                    return
            _log(f"GUI action: set-default start via CLI id={d['id']} name={d['name']} flow={d['flow']} roles=all")
            data = run_audioctl(args, capture_json=True, expect_ok=True)
            cmd_str = "audioctl " + " ".join(shlex.quote(a) for a in args)
            self.maybe_print_cli(cmd_str)
            self.set_status(f"Set default ({d['flow']}) device: {d['name']} (all roles)")
            self.refresh_devices()
            _log(f"GUI action: set-default success via CLI id={d['id']} name={d['name']} flow={d['flow']}")
            _dbg("GUI: on_set_default successful")
        except Exception as e:
            from .logging_setup import _log
            _log(f"GUI action: set-default failed via CLI id={d['id']} name={d['name']} error={e}")
            messagebox.showerror("Error", f"Failed to set default:\n{e}")
            self.set_status("Failed to set default")
    def on_set_volume(self):
        # Volume dialog uses a synced entry + slider so users can be precise but fast.
        # Initial value is read via `get-volume` to reflect current device state.
        d = self.get_selected_device()
        if not d:
            return
        try:
            from .logging_setup import _log
            level = self.open_volume_dialog(d["id"], d["name"])
            if level is None:
                _log(f"GUI action: set-volume cancelled id={d['id']} name={d['name']}")
                return
            _log(f"GUI action: set-volume start via CLI id={d['id']} name={d['name']} flow={d['flow']} level={level}")
            args = [
                "set-volume",
                "--id", d["id"],
                "--flow", d["flow"],
                "--level", str(level),
                "--json",
            ]
            rc, out, err = run_audioctl(args, capture_json=False, expect_ok=False)
            cmd_str = "audioctl " + " ".join(shlex.quote(a) for a in args)
            self.maybe_print_cli(cmd_str)
            if rc == 0:
                try:
                    info = json.loads(out or "{}")
                    vs = info.get("volumeSet") or {}
                    final = vs.get("level", level)
                except Exception:
                    final = level
                self.set_status(f"Volume set to {final}% for: {d['name']}")
                _log(f"GUI action: set-volume success via CLI id={d['id']} name={d['name']} level={final}")
                # Update cached volume for this device only so the menu stays accurate.
                st = self._ensure_device_state_entry(d["id"], d["flow"])
                st["volume"] = final
            else:
                _log(f"GUI action: set-volume failed via CLI id={d['id']} name={d['name']} rc={rc} err={err}")
                messagebox.showerror("Error", f"Failed to set volume/mute:\n{err or out}")
                self.set_status("Failed to set volume")
        except Exception as e:
            from .logging_setup import _log
            _log(f"GUI action: set-volume error via CLI id={d['id']} name={d['name']} err={e}")
            messagebox.showerror("Error", f"Failed to set volume:\n{e}")
            self.set_status("Failed to set volume")
    def on_toggle_mute(self):
        # Mute toggle flow:
        # - Query current mute state via `get-volume` so we know whether to send --mute or --unmute.
        # - This avoids ambiguous "toggle" behavior if we don't know the actual state.
        d = self.get_selected_device()
        if not d:
            return
        try:
            from .logging_setup import _log
            # Query current mute state via CLI
            muted = None
            try:
                data = run_audioctl(
                    ["get-volume", "--id", d["id"], "--flow", d["flow"]],
                    capture_json=True,
                    expect_ok=False,
                )
                if isinstance(data, dict):
                    muted = data.get("muted", None)
            except Exception:
                muted = None
            # Decide target action
            target_flag = "--unmute" if muted is True else "--mute"
            args = [
                "set-volume",
                "--id", d["id"],
                "--flow", d["flow"],
                target_flag,
                "--json",
            ]
            rc, out, err = run_audioctl(args, capture_json=False, expect_ok=False)
            cmd_str = "audioctl " + " ".join(shlex.quote(a) for a in args)
            self.maybe_print_cli(cmd_str)
            if rc == 0:
                try:
                    info = json.loads(out or "{}")
                    ms = info.get("muteSet") or {}
                    final = ms.get("muted", target_flag == "--mute")
                except Exception:
                    final = (target_flag == "--mute")
                self.set_status(f'{"Muted" if final else "Unmuted"}: {d["name"]}')
                _log(f"GUI action: toggle-mute success via CLI id={d['id']} name={d['name']} final={final}")
                # Update cached mute state for this device
                st = self._ensure_device_state_entry(d["id"], d["flow"])
                st["muted"] = bool(final)
            else:
                _log(f"GUI action: toggle-mute failed via CLI id={d['id']} name={d['name']} rc={rc} err={err}")
                messagebox.showerror("Error", f"Failed to change mute state:\n{err or out}")
                self.set_status("Failed to change mute state")
        except Exception as e:
            from .logging_setup import _log
            _log(f"GUI action: toggle-mute error via CLI id={d['id']} name={d['name']} err={e}")
            messagebox.showerror("Error", f"Failed to toggle mute:\n{e}")
            self.set_status("Failed to change mute state")
    def on_toggle_listen(self):
        # Listen is Capture-only by Windows design; Render devices cannot be "listened to".
        # We also query current state first to decide enable/disable deterministically.
        d = self.get_selected_device()
        if not d:
            return
        if d["flow"] != "Capture":
            messagebox.showinfo(
                "Not a capture device",
                "Listen can only be toggled for capture (recording) devices."
            )
            return
        try:
            from .logging_setup import _log
            _log(f"Listen toggle requested for {d['name']} ({d['id']})")
            # Query current 'Listen' state via CLI get-listen
            current = None
            try:
                data_ls = run_audioctl(
                    ["get-listen", "--id", d["id"]],
                    capture_json=True,
                    expect_ok=False,
                )
                if isinstance(data_ls, dict):
                    current = data_ls.get("listenEnabled", None)
            except Exception:
                current = None
            # Decide new target: invert if known, else ask the user
            if current is True:
                enable = False
            elif current is False:
                enable = True
            else:
                choice = messagebox.askyesno(
                    "Toggle Listen",
                    "Enable 'Listen to this device' for this capture device?\n\n"
                    "Yes = Enable\nNo = Disable"
                )
                enable = bool(choice)
            args = ["listen", "--id", d["id"], "--enable" if enable else "--disable", "--json"]
            rc, out, err = run_audioctl(args, capture_json=False, expect_ok=False)
            cmd = "audioctl " + " ".join(shlex.quote(a) for a in args)
            self.maybe_print_cli(cmd)
            if rc == 0:
                try:
                    info = json.loads(out or "{}")
                    final = (info.get("listenSet") or {}).get("enabled", enable)
                except Exception:
                    final = enable
                state_txt = "enabled" if final else "disabled"
                self.set_status(f"Listen {state_txt} for: {d['name']}")
                _log(f"Listen toggle result via CLI for {d['name']} ({d['id']}): final={final}")
                # Update cached listenEnabled state for this device
                st = self._ensure_device_state_entry(d["id"], d["flow"])
                st["listenEnabled"] = bool(final)
            else:
                _log(f"Listen toggle failed via CLI for {d['name']} ({d['id']}): rc={rc} err={err}")
                messagebox.showerror("Error", f"Failed to toggle Listen:\n{err or out}")
                self.set_status("Failed to toggle Listen")
        except Exception as e:
            from .logging_setup import _log
            _log(f"Listen toggle exception via CLI for {d['name']} ({d['id']}): {e!r}")
            messagebox.showerror("Error", f"Failed to toggle Listen:\n{e}")
            self.set_status("Failed to toggle Listen")
    def on_toggle_enhancements(self):
        # Enhancements toggling is vendor-only at runtime.
        # If the CLI reports enhancementsEnabled=None, it means no learned vendor method
        # applies to this device (so we disable and instruct the user to learn).
        #
        # We prefer _pending_enh (captured during menu build) so we toggle from a
        # consistent reference state even if things changed while the menu was open.
        d = self.get_selected_device()
        if not d:
            return
        try:
            from .logging_setup import _log
            _log(f"Enhancements toggle requested for {d['name']} ({d['id']})")
            # Use the state captured when the context menu was built, if available
            st = None
            supported = True
            pe = getattr(self, "_pending_enh", None)
            if pe and pe.get("id") == d["id"] and pe.get("flow") == d["flow"]:
                st = pe.get("current", None)
                supported = pe.get("supported", True)
            # If GUI knows there is no vendor toggle, do not call CLI and inform the user.
            if not supported:
                messagebox.showinfo(
                    "Enhancements not learned",
                    "No vendor method is available for 'Audio Enhancements' on this device yet.\n\n"
                    "Use 'Learn Enhancements' first, then try again."
                )
                self.set_status("Enhancements: no vendor toggle for this device")
                _log(f"Enhancements toggle aborted (no vendor toggle) for {d['name']} ({d['id']})")
                return
            # Decide target: invert current state if known; else, ask user
            if st is True:
                enable = False
            elif st is False:
                enable = True
            else:
                choice = messagebox.askyesno(
                    "Toggle Enhancements",
                    f"Enable 'Audio Enhancements' for this device?\n\n"
                    f"Device: {d['name']} ({d['flow']})\n\n"
                    "Yes = Enable\nNo = Disable"
                )
                enable = bool(choice)
            args = [
                "enhancements",
                "--id", d["id"],
                "--flow", d["flow"],
                "--enable" if enable else "--disable",
            ]
            data = run_audioctl(args, capture_json=True, expect_ok=False)
            cmd_str = "audioctl " + " ".join(shlex.quote(a) for a in args)
            self.maybe_print_cli(cmd_str)
            enh = data.get("enhancementsSet")
            if enh:
                state = enh.get("enabled", enable)
                state_txt = "enabled" if state else "disabled"
                # verifiedBy indicates which vendor method verified the change (e.g. vendor:<ini-section>).
                verified_by = enh.get("verifiedBy", "vendor")
                _log(f"Enhancements toggle via CLI for {d['name']} ({d['id']}): final={state_txt} via {verified_by}")
                self.set_status(f"Enhancements {state_txt} for: {d['name']}")
                # Update cached enhancementsEnabled for this device
                st = self._ensure_device_state_entry(d["id"], d["flow"])
                st["enhancementsEnabled"] = bool(state)
            else:
                _log(f"Enhancements toggle via CLI unverified/failed id={d['id']} name={d['name']} data={data}")
                messagebox.showwarning(
                    "Could not verify",
                    "Vendor toggle applied but could not verify final state, or the toggle failed."
                )
                self.set_status("Enhancements toggle requested (verification failed)")
        except RuntimeError as e:
            from .logging_setup import _log
            _log(f"Enhancements toggle error via CLI for {d['name']} ({d['id']}): {e}")
            messagebox.showerror("Error", str(e))
            self.set_status("Failed to toggle Enhancements")
        except Exception as e:
            from .logging_setup import _log
            _log(f"Enhancements toggle exception via CLI for {d['name']} ({d['id']}): {e!r}")
            messagebox.showerror("Error", f"Failed to toggle Enhancements:\n{e}")
            self.set_status("Failed to toggle Enhancements")
    def on_learn_enhancements(self):
        # Learn flow entrypoint:
        # This is a chooser that allows learning either:
        #  - the main Enhancements vendor toggle, or
        #  - a specific FX (effect) toggle
        #
        # Both ultimately delegate to the CLI interactive learn flows so the COM/registry
        # discovery logic stays in one place.
        d = self.get_selected_device()
        if not d:
            return
        
        try:
            # Choice Dialog
            choice_dialog = tk.Toplevel(self.root)
            choice_dialog.title("Learn Enhancements")
            choice_dialog.transient(self.root)
            choice_dialog.grab_set()
            choice_dialog.resizable(False, False)
            
            try:
                if sys.platform.startswith("win"):
                    choice_dialog.iconbitmap(resource_path("audio.ico"))
            except Exception:
                pass
            
            frm = ttk.Frame(choice_dialog, padding=20)
            frm.pack(fill="both", expand=True)
            
            ttk.Label(frm, text="What are you learning?", 
                      font=("", 10, "bold")).pack(pady=(0, 15))
            
            learn_type = tk.StringVar(value="main")
            fx_name_var = tk.StringVar()
            fx_entry_widget = None
            
            def on_radio_change():
                if learn_type.get() == "fx":
                    fx_entry_widget.config(state="normal")
                    fx_entry_widget.focus_set()
                else:
                    fx_entry_widget.config(state="disabled")
            
            rb1 = ttk.Radiobutton(
                frm, 
                text="The main 'Audio Enhancements' on/off switch",
                variable=learn_type,
                value="main",
                command=on_radio_change
            )
            rb1.pack(anchor="w", pady=5)
            
            rb2 = ttk.Radiobutton(
                frm,
                text="A specific effect (e.g., Bass Boost, Loudness):",
                variable=learn_type,
                value="fx",
                command=on_radio_change
            )
            rb2.pack(anchor="w", pady=5)
            
            fx_name_frame = ttk.Frame(frm)
            fx_name_frame.pack(anchor="w", padx=(30, 0), pady=(5, 15))
            
            ttk.Label(fx_name_frame, text="Effect name:").pack(side="left", padx=(0, 5))
            fx_entry_widget = ttk.Combobox(
                fx_name_frame,
                textvariable=fx_name_var,
                width=30,
                state="disabled",  # becomes "normal" when FX is selected
                values=self._load_fx_names_for_combo()
            )
            fx_entry_widget.pack(side="left")
            # Refresh the list each time the dropdown is opened (vendor_toggles.ini may have changed).
            try:
                fx_entry_widget.configure(postcommand=lambda: fx_entry_widget.configure(values=self._load_fx_names_for_combo()))
            except Exception:
                pass
            # Enable simple type-to-autofill
            self._setup_combobox_autocomplete(fx_entry_widget)
            
            result = {"proceed": False, "type": None, "fx_name": None}
            
            def on_ok():
                if learn_type.get() == "fx":
                    fx = fx_name_var.get().strip()
                    if not fx:
                        messagebox.showerror("Invalid Input", 
                                           "Please enter an effect name.", 
                                           parent=choice_dialog)
                        return
                    result["fx_name"] = fx
                result["type"] = learn_type.get()
                result["proceed"] = True
                choice_dialog.destroy()
            
            def on_cancel():
                result["proceed"] = False
                choice_dialog.destroy()
            
            btn_frame = ttk.Frame(frm)
            btn_frame.pack(anchor="e")
            
            ttk.Button(btn_frame, text="OK", command=on_ok).pack(side="right", padx=(5, 0))
            ttk.Button(btn_frame, text="Cancel", command=on_cancel).pack(side="right")
            
            choice_dialog.bind("<Return>", lambda e: on_ok())
            choice_dialog.bind("<Escape>", lambda e: on_cancel())
            
            choice_dialog.update_idletasks()
            x = self.root.winfo_x() + (self.root.winfo_width() - choice_dialog.winfo_width()) // 2
            y = self.root.winfo_y() + (self.root.winfo_height() - choice_dialog.winfo_height()) // 2
            choice_dialog.geometry(f"+{x}+{y}")
            
            choice_dialog.wait_window()
            
            if not result["proceed"]:
                from .logging_setup import _log as _ilog
                _ilog("GUI action: learn chooser cancelled by user")
                self.set_status("Learn: cancelled by user")
                return
            
            if result["type"] == "main":
                self._learn_main_toggle_via_cli(d)
            else:
                self._learn_fx_toggle_via_cli(d, result["fx_name"])
        
        except Exception as e:
            messagebox.showerror("Error", f"Learn failed:\n{e}")
            self.set_status("Learn failed")
    def _open_main_learn_dialog(self, d):
        self._in_modal_operation = True
        # This variant runs the learn flow non-blocking using LearnRunner:
        # - keep Tk alive (no frozen window),
        # - expose Continue(A)/Continue(B) buttons that write Enter to stdin,
        # - parse the final JSON emitted by the CLI to show a single success/failure popup.
        # Toplevel controller
        top = tk.Toplevel(self.root)
        top.title("Learn Enhancements (Non-blocking)")
        top.transient(self.root)
        top.grab_set()
        top.resizable(True, True)
        try:
            if sys.platform.startswith("win"):
                top.iconbitmap(resource_path("audio.ico"))
        except Exception:
            pass
        frm = ttk.Frame(top, padding=10)
        frm.pack(fill="both", expand=True)
        # Log area
        txt = tk.Text(frm, height=18, width=100, wrap="word")
        txt.grid(row=0, column=0, columnspan=4, sticky="nsew", pady=(0, 8))
        frm.rowconfigure(0, weight=1)
        frm.columnconfigure(0, weight=1)
        # Buttons
        btn_a = ttk.Button(frm, text="Continue (A)", state="disabled")
        btn_b = ttk.Button(frm, text="Continue (B)", state="disabled")
        btn_cancel = ttk.Button(frm, text="Cancel")
        btn_retry = ttk.Button(frm, text="Retry (skip confirmation)", state="disabled")
        btn_a.grid(row=1, column=0, sticky="w", padx=(0, 6))
        btn_b.grid(row=1, column=1, sticky="w", padx=(0, 6))
        btn_cancel.grid(row=1, column=2, sticky="e", padx=(0, 6))
        btn_retry.grid(row=1, column=3, sticky="e")
        # Single GUI confirmation (non-blocking design).
        # We warn once here; then LearnRunner can set AUDIOCTL_LEARN_CONFIRMED so
        # the CLI doesn't ask the user to type I UNDERSTAND again.
        try:
            ini_path = _vendor_ini_default_path()
        except Exception:
            ini_path = "<vendor_toggles.ini>"
        warn_txt = (
            "READ CAREFULLY\n\n"
            "This Learn mode will capture two registry snapshots and write a vendor entry via the CLI into:\n"
            f"  {ini_path}\n\n"
            "From now on, future 'Enhancements' commands for this device WILL WRITE registry values.\n\n"
            "During Learn:\n"
            "- Do NOT change other audio settings.\n"
            "- Do NOT switch devices.\n"
            "- Only toggle 'Audio Enhancements' for THIS device when prompted.\n\n"
            "Click OK to continue, or Cancel to abort."
        )
        if not messagebox.askokcancel("Warning – Learn writes registry (persistent)", warn_txt, parent=top):
            self._in_modal_operation = False
            try:
                top.destroy()
            except Exception:
                pass
            return
        # Runner plumbing
        args_list = ["enhancements", "--id", d["id"], "--flow", d["flow"], "--learn"]
        def append_log(s):
            # marshal back to Tk thread
            self.root.after(0, lambda: (txt.insert("end", s), txt.see("end")))
        def handle_state(st):
            def _apply():
                if st == "started":
                    btn_a.configure(state="disabled")
                    btn_b.configure(state="disabled")
                    btn_retry.configure(state="disabled")
                elif st == "waiting_snapshot_a":
                    btn_a.configure(state="normal")
                elif st == "waiting_snapshot_b":
                    btn_b.configure(state="normal")
                elif st in ("done", "error"):
                    # No more steps
                    btn_a.configure(state="disabled")
                    btn_b.configure(state="disabled")
                    btn_retry.configure(state="disabled")
                    # Parse final JSON (if any).
                    # The CLI emits a JSON object at the end with either:
                    #   - vendorLearned: learned and wrote/updated the INI
                    #   - vendorAvailable: a method exists but may not have written a new section
                    out_text = "".join(runner.collected_out)
                    info = None
                    try:
                        lines = out_text.splitlines()
                        for raw in reversed(lines):
                            line = raw.strip()
                            if not line or not line.startswith("{"):
                                continue
                            try:
                                cand = json.loads(line)
                            except Exception:
                                continue
                            if "vendorLearned" in cand or "vendorAvailable" in cand:
                                info = cand
                                break
                    except Exception:
                        info = None
                    # Show the single final popup (success or failure)
                    if st == "done" and isinstance(info, dict):
                        try:
                            if "vendorLearned" in info:
                                msg = info["vendorLearned"]
                                messagebox.showinfo(
                                    "Learn Enhancements",
                                    f"Learned vendor toggle.\n\nSection: {msg.get('section')}\nINI: {msg.get('iniPath')}"
                                )
                            elif "vendorAvailable" in info:
                                msg = info["vendorAvailable"]
                                messagebox.showinfo(
                                    "Learn Enhancements",
                                    f"Vendor method available.\n\nVendor: {msg.get('vendor')}\nValue: {msg.get('value_name')}"
                                )
                        except Exception:
                            pass
                        # Refresh on success so Enhancements actions become enabled immediately.
                        self.refresh_devices()
                    else:
                        try:
                            messagebox.showerror(
                                "Learn Enhancements",
                                "Learn failed or could not verify a vendor entry.\nSee the log text for details."
                            )
                        except Exception:
                            pass
                    # After the user dismisses the popup, close the controller window
                    self._in_modal_operation = False
                    try:
                        top.destroy()
                    except Exception:
                        pass
            self.root.after(0, _apply)
        # First attempt: auto-confirm when we see the prompt
        runner = LearnRunner(args_list, on_output=append_log, on_state=handle_state, confirmed=False)
        btn_a.configure(command=runner.continue_snapshot_a)
        btn_b.configure(command=runner.continue_snapshot_b)
        def do_cancel():
            # Clear modal flag and close
            self._in_modal_operation = False
            try:
                runner.terminate()
            except Exception:
                pass
            top.destroy()
        def do_retry():
            nonlocal runner  # move this up
            # Retry with env skip (no confirmation prompt)
            btn_retry.configure(state="disabled")
            try:
                runner.terminate()
            except Exception:
                pass
            # New runner, confirmed=True -> sets AUDIOCTL_LEARN_CONFIRMED=1
            new_runner = LearnRunner(args_list, on_output=append_log, on_state=handle_state, confirmed=True)
            btn_a.configure(command=new_runner.continue_snapshot_a, state="disabled")
            btn_b.configure(command=new_runner.continue_snapshot_b, state="disabled")
            # Start on next loop tick
            self.root.after(0, new_runner.start)
            # Rebind closures
            runner = new_runner
        btn_cancel.configure(command=do_cancel)
        btn_retry.configure(command=do_retry)
        # kick off
        runner.start()
        top.protocol("WM_DELETE_WINDOW", do_cancel)
    def _learn_main_toggle_via_cli(self, d):
        """
        Delegate 'Learn Enhancements' (main) to the CLI interactive flow:
          audioctl enhancements --id "<id>" --flow <Flow> --learn

        Key behavior:
        - GUI shows the warning once, then sets AUDIOCTL_LEARN_CONFIRMED=1 so the CLI
          does not request the "I UNDERSTAND" confirmation again.
        - run_audioctl_interactive watches stdout prompts and "presses Enter" for the user.
        - We parse the final JSON out of mixed stdout text (prompts + JSON) by scanning
          backwards for a balanced JSON object containing vendorLearned/vendorAvailable.
        """
        from .logging_setup import _log
        try:
            self._in_modal_operation = True
            self._suspend_print_cli_and_disable_checkbox()
            try:
                ini_path = _vendor_ini_default_path()
            except Exception:
                ini_path = "<vendor_toggles.ini>"
            warn_txt = (
                "READ CAREFULLY\n\n"
                "This Learn mode will capture registry snapshots and write a vendor entry via the CLI into:\n"
                f"  {ini_path}\n\n"
                "From now on, future 'Enhancements' commands for this device WILL WRITE registry values.\n\n"
                "During Learn:\n"
                "- Do NOT change other audio settings.\n"
                "- Do NOT switch devices.\n"
                "- Only toggle 'Audio Enhancements' for THIS device when prompted.\n\n"
                "Click OK to continue, or Cancel to abort."
            )
            if not messagebox.askokcancel("Warning – Learn writes registry (persistent)", warn_txt):
                self.set_status("Learn Enhancements: aborted by user")
                _log(f"GUI action: learn-main cancelled id={d['id']} name={d['name']}")
                return
            # Prompt patterns (repeats are naturally handled: same substrings will reappear)
            prompt_patterns = [
                (
                    "set 'Audio Enhancements' to ENABLED",
                    "Learn Enhancements – Step 1",
                    "In Windows Sound settings, set 'Audio Enhancements' to ENABLED for this device.\n\n"
                    "Click OK to continue."
                ),
                (
                    "set 'Audio Enhancements' to DISABLED",
                    "Learn Enhancements – Step 2",
                    "In Windows Sound settings, set 'Audio Enhancements' to DISABLED for this device.\n\n"
                    "Click OK to continue."
                ),
            ]
            args = [
                "enhancements",
                "--id", d["id"],
                "--flow", d["flow"],
                "--learn",
            ]
            # Temporarily set the env flag so CLI skips its own "I UNDERSTAND" input
            prev = os.environ.get("AUDIOCTL_LEARN_CONFIRMED")
            os.environ["AUDIOCTL_LEARN_CONFIRMED"] = "1"
            try:
                rc, out, err = run_audioctl_interactive(args, prompt_patterns, expect_ok=False)
            finally:
                if prev is None:
                    try:
                        del os.environ["AUDIOCTL_LEARN_CONFIRMED"]
                    except Exception:
                        pass
                else:
                    os.environ["AUDIOCTL_LEARN_CONFIRMED"] = prev
            # Parse final JSON: vendorLearned or vendorAvailable (robust even if prompt text and JSON share a line)
            def _extract_last_vendor_json(text: str):
                s = text or ""
                i = len(s) - 1
                while i >= 0:
                    # find a closing brace
                    while i >= 0 and s[i] != "}":
                        i -= 1
                    if i < 0:
                        break
                    end = i
                    depth = 1
                    i -= 1
                    # walk backwards to find matching opening brace
                    while i >= 0 and depth > 0:
                        ch = s[i]
                        if ch == "}":
                            depth += 1
                        elif ch == "{":
                            depth -= 1
                        i -= 1
                    if depth == 0:
                        start = i + 1
                        chunk = s[start:end + 1]
                        try:
                            obj = json.loads(chunk)
                            if isinstance(obj, dict) and ("vendorLearned" in obj or "vendorAvailable" in obj):
                                return obj
                        except Exception:
                            pass
                    # continue scanning earlier content
                    i = end - 1
                return None
            data = _extract_last_vendor_json(out)
            if not isinstance(data, dict):
                messagebox.showwarning(
                    "Learn Enhancements",
                    "CLI ran but did not report a learned vendor entry.\n"
                    "See console/log for details."
                )
                self.set_status("Learn Enhancements: CLI did not learn entry")
                _log(f"GUI action: learn-main unknown-cli-output id={d['id']} name={d['name']} out={out!r} err={err!r}")
                return
            if "vendorLearned" in data:
                info = data["vendorLearned"]
                msg = (
                    "Learned a vendor toggle via CLI.\n\n"
                    f"Section: {info.get('section')}\n"
                    f"Value:   {info.get('value_name')}\n"
                    f"INI:     {info.get('iniPath')}\n\n"
                    "You can now use 'Enable/Disable Enhancements' on this device."
                )
                messagebox.showinfo("Learn Enhancements", msg)
                self.set_status("Learn Enhancements: vendor learned via CLI")
                _log(f"GUI action: learn-main success via CLI interactive id={d['id']} name={d['name']} info={info}")
                self.refresh_devices()
            elif "vendorAvailable" in data:
                info = data["vendorAvailable"]
                msg = (
                    "A vendor method is available for this device.\n\n"
                    f"Vendor: {info.get('vendor')}\n"
                    f"Value:  {info.get('value_name')}\n\n"
                    "No new INI section was written, but this device can already be controlled."
                )
                messagebox.showinfo("Learn Enhancements", msg)
                self.set_status("Learn Enhancements: vendor already available (CLI)")
                _log(f"GUI action: learn-main vendor-available via CLI interactive id={d['id']} name={d['name']} info={info}")
                self.refresh_devices()
            else:
                messagebox.showwarning(
                    "Learn Enhancements",
                    "CLI ran but did not report a learned vendor entry.\n"
                    "See console/log for details."
                )
                self.set_status("Learn Enhancements: CLI did not learn entry")
                _log(f"GUI action: learn-main unknown-json via CLI interactive id={d['id']} name={d['name']} data={data}")
        finally:
            self._restore_print_cli_checkbox()
            self._in_modal_operation = False
    def _learn_fx_toggle_via_cli(self, d, fx_name):
        """
        Delegate FX learn to the existing interactive CLI flow:
          audioctl enhancements --learn-fx "<FX_NAME>" ...

        Why prompt ordering matters:
        FX learn is a two-pass A/B (A,B then A2,B2). We match the second-pass prompts
        first because those lines can appear later and are more specific; ordering
        reduces the chance of matching the wrong pattern when text is similar.
        """
        from .logging_setup import _log
        try:
            self._in_modal_operation = True
            self._suspend_print_cli_and_disable_checkbox()
            try:
                ini_path = _vendor_ini_default_path()
            except Exception:
                ini_path = "<vendor_toggles.ini>"
            warn_txt = (
                f"READ CAREFULLY\n\n"
                f"This Learn mode will use the CLI interactive flow for '{fx_name}'.\n"
                f"CLI will write a vendor FX entry into:\n  {ini_path}\n\n"
                "From now on, this effect may be controlled via registry writes.\n\n"
                "Critical rules during Learn:\n"
                f"- ONLY toggle the '{fx_name}' checkbox/effect.\n"
                "- Do NOT toggle the main 'Audio Enhancements' switch.\n"
                "- Do NOT change any other audio settings.\n"
                "- Do NOT switch devices.\n\n"
                "Click OK to continue, or Cancel to abort."
            )
            if not messagebox.askokcancel(f"Warning – Learn FX '{fx_name}'", warn_txt):
                self.set_status(f"Learn FX '{fx_name}': aborted by user")
                _log(f"GUI action: learn-fx cancelled id={d['id']} name={d['name']} fx={fx_name}")
                return
            # Prompts reordered and tightened so that second pass (A2/B2) is matched first
            prompt_patterns = [
                # Second pass first (most specific) – instruction lines only
                (
                    "Enable the",
                    f"Learn FX '{fx_name}' - Step 3",
                    f"Enable the '{fx_name}' effect again (second pass), then click OK to continue."
                ),
                (
                    "Disable the",
                    f"Learn FX '{fx_name}' - Step 4",
                    f"Disable the '{fx_name}' effect again (second pass), then click OK to continue."
                ),
                # First pass (instruction lines only)
                (
                    "effect to ENABLED for this device.",
                    f"Learn FX '{fx_name}' - Step 1",
                    f"ENABLE the '{fx_name}' effect for this device.\n"
                    "(Do NOT toggle the main 'Audio Enhancements' switch.)\n\n"
                    "Click OK, then the GUI will continue."
                ),
                (
                    "to DISABLED for the same device.",
                    f"Learn FX '{fx_name}' - Step 2",
                    f"DISABLE the '{fx_name}' effect for this device.\n\n"
                    "Click OK, then the GUI will continue."
                ),
            ]
            args = [
                "enhancements",
                "--id", d["id"],
                "--flow", d["flow"],
                "--learn-fx", fx_name,
            ]
            # Run interactive; it will pop dialog for each matched line and press Enter afterwards
            rc, out, err = run_audioctl_interactive(args, prompt_patterns, expect_ok=False)
            # Parse final fxLearned JSON if present (robust extraction even if concatenated with prompt text)
            fx_info = None
            try:
                lines = (out or "").splitlines()
                for raw in reversed(lines):
                    line = raw.strip()
                    if '"fxLearned"' not in line:
                        continue
                    start = line.find("{")
                    end = line.rfind("}")
                    if start == -1 or end == -1 or end <= start:
                        continue
                    json_text = line[start:end+1]
                    try:
                        data = json.loads(json_text)
                    except Exception:
                        continue
                    candidate = data.get("fxLearned")
                    if candidate:
                        fx_info = candidate
                        break
            except Exception:
                fx_info = None
            if fx_info:
                msg = (
                    f"Learned FX '{fx_name}' via CLI.\n\n"
                    f"Section: {fx_info.get('section')}\n"
                    f"FX:      {fx_info.get('fx_name')}\n"
                    f"INI:     {fx_info.get('iniPath')}\n\n"
                    "The effect will now appear (or be updated) in the context menu."
                )
                messagebox.showinfo(f"Learn FX '{fx_name}'", msg)
                self.set_status(f"Learn FX '{fx_name}': vendor INI updated")
                _log(f"GUI action: learn-fx success via CLI interactive id={d['id']} name={d['name']} fx={fx_name} info={fx_info}")
                self.refresh_devices()
            else:
                messagebox.showwarning(
                    f"Learn FX '{fx_name}'",
                    "CLI interactive flow completed but did not report a learned FX entry.\n"
                    "See console/log for details."
                )
                self.set_status(f"Learn FX '{fx_name}': no fxLearned JSON detected")
                _log(f"GUI action: learn-fx no-json id={d['id']} name={d['name']} fx={fx_name} out={out!r} err={err!r}")
        finally:
            self._restore_print_cli_checkbox()
            self._in_modal_operation = False
    def open_volume_dialog(self, device_id, device_name):
        top = tk.Toplevel(self.root)
        try:
            if sys.platform.startswith("win"):
                top.iconbitmap(resource_path("audio.ico"))
        except Exception:
            pass
        top.title("Set Volume")
        top.transient(self.root)
        top.grab_set()
        top.resizable(False, False)
        frm = ttk.Frame(top, padding=12)
        frm.pack(fill="both", expand=True)
        ttk.Label(frm, text=device_name, anchor="w").grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 8))
        # Query current volume via CLI so the dialog starts at the real device level.
        initial = 50
        try:
            data = run_audioctl(
                ["get-volume", "--id", device_id],
                capture_json=True, expect_ok=False
            )
            if isinstance(data, dict) and data.get("volume") is not None:
                initial = int(data["volume"])
        except Exception:
            pass
        v = tk.IntVar(value=initial)
        # Two-way sync between entry and slider must avoid feedback loops.
        syncing = {"entry": False, "scale": False}
        def _validate(P):
            if P == "":
                return True
            if not P.isdigit():
                return False
            if len(P) > 3:
                return False
            try:
                val = int(P)
            except Exception:
                return False
            return 0 <= val <= 100
        vcmd = (top.register(_validate), "%P")
        entry = ttk.Entry(frm, width=3, textvariable=v, validate="key", validatecommand=vcmd, justify="right")
        entry.grid(row=1, column=0, sticky="w")
        ttk.Label(frm, text="%").grid(row=1, column=1, sticky="w", padx=(4, 12))
        def on_scale(valstr):
            if syncing["entry"]:
                return
            try:
                syncing["scale"] = True
                v.set(int(float(valstr)))
            finally:
                syncing["scale"] = False
        scale = ttk.Scale(frm, from_=0, to=100, orient="horizontal", command=on_scale)
        scale.set(initial)
        scale.grid(row=1, column=2, sticky="we")
        frm.columnconfigure(2, weight=1)
        def on_entry_change(*_):
            if syncing["scale"]:
                return
            try:
                syncing["entry"] = True
                try:
                    scale.set(int(v.get()))
                except Exception:
                    pass
            finally:
                syncing["entry"] = False
        v.trace_add("write", on_entry_change)
        btns = ttk.Frame(frm)
        btns.grid(row=2, column=0, columnspan=3, sticky="e", pady=(12, 0))
        result = {"value": None}
        def ok():
            try:
                result["value"] = max(0, min(100, int(v.get())))
            except Exception:
                result["value"] = None
            top.destroy()
        def cancel():
            result["value"] = None
            top.destroy()
        ttk.Button(btns, text="OK", command=ok).pack(side="right")
        ttk.Button(btns, text="Cancel", command=cancel).pack(side="right", padx=(0, 8))
        top.bind("<Return>", lambda e: ok())
        top.bind("<Escape>", lambda e: cancel())
        entry.focus_set()
        top.wait_window()
        return result["value"]
    def on_toggle_fx_live(self, fx_name, current_state):
        """
        Toggle an FX via CLI, based on the state we saw when the menu was built.

        FX toggles are executed through `audioctl enhancements` sub-operations:
          --enable-fx / --disable-fx

        On success we update the cached FX state so the menu reflects the change
        immediately.
        current_state: True (enabled), False (disabled), or None (unknown).
        """
        d = self.get_selected_device()
        if not d:
            return
        try:
            from .logging_setup import _log
            if current_state is True:
                enable = False
            else:
                enable = True
            args = [
                "enhancements",
                "--id", d["id"],
                "--flow", d["flow"],
                "--enable-fx" if enable else "--disable-fx",
                fx_name,
            ]
            data = run_audioctl(args, capture_json=True, expect_ok=False)
            cmd_str = "audioctl " + " ".join(shlex.quote(a) for a in args)
            self.maybe_print_cli(cmd_str)
            fx_set = data.get("fxSet")
            if fx_set:
                state = fx_set.get("enabled", enable)
                state_txt = "enabled" if state else "disabled"
                self.set_status(f"{fx_name} {state_txt} for: {d['name']}")
                _log(f"GUI action: toggle-fx success via CLI id={d['id']} name={d['name']} fx={fx_name} final={state}")
                # Update cached FX state for this device
                st = self._ensure_device_state_entry(d["id"], d["flow"])
                fx_list = st.get("availableFX") or []
                # Find matching fx_name in availableFX and update its state
                for fx in fx_list:
                    if (fx.get("fx_name") or "").strip().lower() == fx_name.strip().lower():
                        fx["state"] = bool(state)
                        break
            else:
                messagebox.showwarning(
                    "FX Toggle Failed",
                    f"Could not toggle effect '{fx_name}'.\n"
                    "The effect may not be properly learned for this device."
                )
                self.set_status(f"Failed to toggle {fx_name}")
                _log(f"GUI action: toggle-fx via CLI failed id={d['id']} name={d['name']} fx={fx_name}")
        except RuntimeError as e:
            from .logging_setup import _log
            _log(f"GUI action: toggle-fx via CLI error id={d['id']} name={d['name']} fx={fx_name} err={e}")
            messagebox.showerror("Error", f"Failed to toggle {fx_name}:\n{e}")
            self.set_status(f"Error toggling {fx_name}")
        except Exception as e:
            from .logging_setup import _log
            messagebox.showerror("Error", f"Failed to toggle {fx_name}:\n{e}")
            self.set_status(f"Error toggling {fx_name}")
            _log(f"GUI action: toggle-fx via CLI exception id={d['id']} name={d['name']} fx={fx_name} err={e}")
    def _load_fx_names_for_combo(self):
        # Returns a sorted list of unique FX names from vendor_toggles.ini
        # Used only to suggest names for the learn dialog; it does not toggle anything.
        try:
            ini_path = _vendor_ini_default_path()
            db = _load_vendor_db_split(ini_path)
            names = sorted({(e.get("fx_name") or "").strip()
                            for e in (db.get("fx") or [])
                            if (e.get("fx_name") or "").strip()})
            return names
        except Exception:
            return []
def _setup_combobox_autocomplete(self, combo: ttk.Combobox):
    def on_keyrelease(ev):
        # 1. Ignore modifier keys (Shift, Caps_Lock, Control, etc.)
        if ev.keysym in ("Shift_L", "Shift_R", "Caps_Lock", "Control_L", "Control_R", "Alt_L", "Alt_R"):
            return

        # 2. Ignore backspace/delete so user can actually correct mistakes
        if ev.keysym in ("BackSpace", "Delete", "Left", "Right"):
            return

        text = combo.get()
        if not text: 
            return

        vals = list(combo.cget("values") or [])
        low = text.lower()
        
        for v in vals:
            s = str(v)
            if s.lower().startswith(low):
                # Set the full text but keep track of where the user was
                typed_len = len(text)
                combo.set(s)
                
                # Position cursor at the end of the TYPED part, 
                # and select (highlight) the suggested completion
                combo.icursor(typed_len)
                combo.select_range(typed_len, tk.END)
                break

    combo.bind("<KeyRelease>", on_keyrelease, add="+")
def launch_gui():
    # GUI bootstrap:
    # - loads icon (best-effort)
    # - installs safe shutdown hooks
    # - routes Tk callback exceptions into our log file so crashes are diagnosable
    #   even when launched without a console (double-clicked EXE).
    _log("launch_gui: creating Tk root")
    root = tk.Tk()
    try:
        if sys.platform.startswith("win"):
            root.iconbitmap(resource_path("audio.ico"))
    except Exception:
        pass
    gui = AudioGUI(root)
    def _on_root_close():
        try:
            _log("WM_DELETE_WINDOW received: root close requested (user/system)")
        except Exception:
            pass
        root.destroy()
    try:
        root.protocol("WM_DELETE_WINDOW", _on_root_close)
    except Exception:
        pass
    def _on_any_destroy(ev):
        try:
            if ev.widget == root:
                _log("Tk <Destroy> on root window")
        except Exception:
            pass
    try:
        root.bind("<Destroy>", _on_any_destroy, add="+")
    except Exception:
        pass
    def _tk_report_callback_exception(exc, val, tb):
        # Tk normally swallows callback exceptions; we log them and show a dialog with the log path.
        try:
            _log_exc("TK CALLBACK EXCEPTION", (exc, val, tb))
        except Exception:
            pass
        try:
            messagebox.showerror("Unexpected error", f"{exc.__name__}: {val}\n\nDetails were written to:\n{_log_path()}")
        except Exception:
            pass
    try:
        root.report_callback_exception = _tk_report_callback_exception
    except Exception:
        pass
    _log("launch_gui: entering mainloop")
    try:
        root.mainloop()
    except Exception:
        _log_exc("MAINLOOP EXCEPTION")
    _log("launch_gui: mainloop exited")
    return 0



