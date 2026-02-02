# audioctl/gui.py
# -----------------------------------------------------------------------------
# GUI philosophy / design constraints
# -----------------------------------------------------------------------------
# This GUI is intentionally a *thin front-end* over the CLI:
#   - The GUI shells out to `audioctl` CLI commands as subprocesses for *all*
#     real work (list devices, set default, volume/mute, listen, enhancements,
#     learn flows, etc.).
#   - The GUI avoids importing the low-level COM / registry manipulation code
#     directly (devices.py, pycaw, comtypes usage) to keep COM lifetime and
#     GC-sensitive behavior centralized in one place (the CLI helpers).
#
# Why this matters:
#   - COM stability: The project has a long history of shutdown / GC crashes
#     caused by COM objects being released from the wrong thread or at the wrong
#     time. Keeping the GUI "CLI-only" prevents accidental COM imports in the
#     Tkinter process that would reintroduce lifetime issues.
#   - Single source of truth: The CLI defines the supported behaviors and the
#     JSON schemas. The GUI only orchestrates and presents the results.
#   - Frozen EXE behavior: In a PyInstaller build, `sys.executable` is the EXE.
#     In source mode, we must run `python -m audioctl ...`. The helpers below
#     hide that difference so the GUI works both frozen and unfrozen.
# -----------------------------------------------------------------------------

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
from .vendor_db import (
    _vendor_ini_default_path,
    _load_vendor_db_split,  # NEW
)

# --- BEGIN: Non-blocking Learn runner (main Enhancements) ---
import threading


def _build_cli_cmd(args_list):
    # Build the subprocess command line in a way that works in:
    #   - Frozen mode (PyInstaller): sys.executable is the bundled audioctl.exe.
    #   - Source mode: we run "python -m audioctl ..." so imports and package
    #     resolution are correct.
    #
    # Keeping this logic in one place helps ensure *every* GUI call uses the same
    # target binary/module.
    if getattr(sys, "frozen", False):
        return [sys.executable] + args_list
    else:
        return [sys.executable, "-m", "audioctl"] + args_list


class LearnRunner:
    # Learn flows are interactive: the CLI prints prompts like
    # "When ready, press Enter to capture snapshot A...".
    #
    # LearnRunner exists to run those flows without blocking Tk's mainloop:
    #   - We stream stdout/stderr asynchronously.
    #   - We detect prompt text using regex patterns.
    #   - We expose "Continue" methods that write "\n" to stdin to simulate
    #     pressing Enter at the correct time.
    #
    # Additionally, the CLI's learn mode includes a stern confirmation prompt
    # that expects typing "I UNDERSTAND". The GUI shows its *own* warning UI,
    # so LearnRunner can auto-confirm that CLI prompt when needed.

    # Regex patterns are intentionally broad enough to survive minor wording
    # changes while still being specific to the learn prompts.
    PATTERN_CONFIRM = re.compile(r"I UNDERSTAND")  # literal appears in the prompt text
    PATTERN_A = re.compile(r"When ready, press Enter to capture snapshot A", re.IGNORECASE)
    PATTERN_B = re.compile(r"When ready, press Enter to capture snapshot B", re.IGNORECASE)

    def __init__(self, args_list, on_output, on_state, confirmed=False):
        self.args_list = args_list
        self.on_output = on_output or (lambda _t: None)
        self.on_state = on_state or (lambda _s: None)

        # "confirmed" means the GUI has already shown the warning and we should
        # skip (or auto-satisfy) the CLI's confirmation step.
        self.confirmed = bool(confirmed)

        self.proc = None

        # Internal flags that track which prompt we're waiting on. These gates
        # ensure we only send Enter when the CLI is actually expecting it.
        self._waiting_a = False
        self._waiting_b = False

        # Track whether we've already responded to the CLI confirmation prompt.
        self._sent_confirm = False

        # Collected output buffers are used by the GUI to parse final JSON once
        # the process exits (learn flows can interleave prompts + JSON).
        self.collected_out = []
        self.collected_err = []

    def start(self):
        # Learn is long-running. We must not freeze the GUI, so we start the CLI
        # as a subprocess and read it on background threads.
        env = os.environ.copy()
        if self.confirmed:
            # This env var is a contract with vendor_db.py: it tells the CLI learn
            # routine to skip the "type I UNDERSTAND" input prompt.
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

        # Separate threads for stdout and stderr keep the pipes draining; this
        # prevents deadlocks and lets us react to prompts as soon as they appear.
        threading.Thread(target=self._read_stream, args=(self.proc.stdout, True), daemon=True).start()
        threading.Thread(target=self._read_stream, args=(self.proc.stderr, False), daemon=True).start()
        threading.Thread(target=self._waiter, daemon=True).start()

    def _waiter(self):
        # Wait for the learn process to finish and then notify the GUI.
        rc = self.proc.wait()
        self.on_state("done" if rc == 0 else "error")

    def _read_stream(self, stream, is_stdout):
        # We read char-by-char (bufsize=0) to detect prompts that might not be
        # newline-terminated promptly. This improves responsiveness for the GUI
        # "Continue" buttons.
        buf = ""
        while True:
            ch = stream.read(1)
            if ch is None or ch == "":
                break
            buf += ch

            # Emit complete lines when possible.
            if "\n" in buf:
                lines = buf.split("\n")
                for ln in lines[:-1]:
                    self._handle_text(ln + "\n", is_stdout)
                buf = lines[-1]
            else:
                # Still scan partial fragments for prompt patterns.
                self._scan_for_prompts(buf)

        if buf:
            self._handle_text(buf, is_stdout)

    def _handle_text(self, text, is_stdout):
        # Persist output for later parsing (e.g., extracting a JSON result).
        try:
            if is_stdout:
                self.collected_out.append(text)
            else:
                self.collected_err.append(text)
        except Exception:
            pass

        # Forward to GUI (typically appends to a text widget).
        self.on_output(text)

        # Prompts can occur anywhere in output; scan continuously.
        self._scan_for_prompts(text)

    def _scan_for_prompts(self, text):
        t = text if isinstance(text, str) else str(text or "")

        # Auto-confirm (only on first attempt). This prevents the CLI from
        # blocking at stdin waiting for the confirmation phrase.
        if (not self.confirmed) and (not self._sent_confirm) and self.PATTERN_CONFIRM.search(t):
            try:
                self.proc.stdin.write("I UNDERSTAND\n")
                self.proc.stdin.flush()
                self._sent_confirm = True
            except Exception:
                pass
            return

        # Snapshot A prompt: signal GUI to enable "Continue (A)".
        if self.PATTERN_A.search(t):
            self._waiting_a = True
            self.on_state("waiting_snapshot_a")
            return

        # Snapshot B prompt: signal GUI to enable "Continue (B)".
        if self.PATTERN_B.search(t):
            self._waiting_b = True
            self.on_state("waiting_snapshot_b")
            return

    def continue_snapshot_a(self):
        # Simulate pressing Enter for the CLI prompt A.
        if self._waiting_a and self.proc and self.proc.stdin:
            try:
                self.proc.stdin.write("\n")
                self.proc.stdin.flush()
            except Exception:
                pass
            self._waiting_a = False

    def continue_snapshot_b(self):
        # Simulate pressing Enter for the CLI prompt B.
        if self._waiting_b and self.proc and self.proc.stdin:
            try:
                self.proc.stdin.write("\n")
                self.proc.stdin.flush()
            except Exception:
                pass
            self._waiting_b = False

    def terminate(self):
        # Best-effort cancellation for a running learn flow.
        try:
            if self.proc and self.proc.poll() is None:
                self.proc.terminate()
        except Exception:
            pass


# --- END: Non-blocking Learn runner ---


def run_audioctl(args_list, capture_json=False, expect_ok=True):
    """
    Run 'audioctl' CLI as a subprocess.

    Why subprocess instead of importing logic directly:
      - Keeps COM/pycaw/comtypes initialization strictly inside CLI helpers.
      - Avoids mixed COM apartments / destructor timing inside the Tk thread.
      - Ensures GUI behavior matches CLI behavior exactly (same selection rules,
        same JSON schemas, same verification logic).

    Frozen vs source:
      - Frozen: sys.executable is audioctl.exe, so we run `[exe] + args`.
      - Source: run `python -m audioctl ...` to execute package entrypoint.

    Output handling:
      - stdout is where the CLI prints JSON success payloads.
      - stderr is where the CLI prints human-readable errors/warnings.
      - We capture both so the GUI can show errors with context.
    """
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
        # Most GUI operations expect CLI JSON on success. We parse stdout into a
        # dict so callers can read structured fields (e.g., enhancementsSet).
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
    Fast one-shot CLI call with timeout. Returns parsed JSON or None on failure.

    This is used primarily for "menu freshness":
      - The context menu needs up-to-date labels (Mute vs Unmute, Enable vs
        Disable Enhancements, FX states).
      - Blocking too long during right-click is a bad UX; quick_json bounds the
        wait and fails closed (keep cached labels) if the system is slow.
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
    Run 'audioctl' CLI as a subprocess, line-by-line, and respond to prompts.

    This mode exists specifically for CLI learn flows:
      - The CLI prints human instructions and then blocks on input().
      - The GUI shows those instructions in messageboxes.
      - After the user clicks OK, the GUI sends '\n' to stdin to let the CLI
        proceed.

    Why line-by-line:
      - CLI prompts are printed in sequence and are user-facing text; matching
        by substring on each line is stable and simple.
      - We keep stdin open so we can "press Enter" programmatically.

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

    # IMPORTANT: stdin must be PIPE so we can answer input() prompts.
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

    # Read stdout line by line; on matches, prompt the user and send Enter.
    while True:
        line = proc.stdout.readline()
        if line == "":
            break  # EOF
        collected_out.append(line)

        for substring, title, custom_message in prompt_patterns:
            if substring in line:
                # Show a stable, user-friendly message instead of raw CLI line
                # when a custom message is provided.
                msg_text = custom_message if custom_message is not None else line.strip()
                try:
                    messagebox.showinfo(title, msg_text)
                except Exception:
                    pass

                # Simulate pressing Enter for the CLI input().
                try:
                    proc.stdin.write("\n")
                except Exception:
                    pass
                try:
                    proc.stdin.flush()
                except Exception:
                    pass
                break

    # Collect remaining output.
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
        # Tk root and baseline window identity.
        self.root = root
        self.root.title("Mr5niper's Audio Control  v1.4.7.2  02-01-2026")

        # Style/theme selection: best-effort. We avoid hard dependencies on a
        # specific theme because Tk installations vary.
        style = ttk.Style(self.root)
        try:
            for theme in ("vista", "xpnative", "clam", "default"):
                if theme in style.theme_names():
                    style.theme_use(theme)
                    break
        except Exception:
            pass

        # Try to make headings visually distinct (bold) without relying on a
        # specific font being available.
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

        # Flatten heading look if possible (cosmetic).
        try:
            style.configure("Treeview.Heading", relief="flat")
            style.map("Treeview.Heading", background=[], relief=[], foreground=[])
        except Exception:
            pass

        # Core GUI state:
        #   - include_all: include disabled/disconnected endpoints in list output.
        #   - print_cmd: optional UX aid to show the CLI command equivalent.
        self.include_all = tk.BooleanVar(value=False)
        self.print_cmd = tk.BooleanVar(value=False)

        # Device list as returned by `audioctl list --json`.
        self.devices = []

        # Mapping between Treeview items and device dicts. Group header rows are
        # not devices and are not added to this map.
        self.item_to_device = {}

        # Cache of per-device "live" state from `get-device-state`.
        # This cache lets us build context menus instantly without making slow
        # CLI calls on every right-click. It is refreshed incrementally after
        # every device list refresh.
        #
        # Key: endpoint id
        # Value: dict from CLI {volume, muted, listenEnabled, enhancementsEnabled, availableFX...}
        self.device_state_cache = {}

        # Guard flag used to suppress certain auto-refresh behaviors during
        # modal flows (especially Learn). Some learn operations prompt the user
        # multiple times and we do not want background refresh to interfere.
        self._in_modal_operation = False

        # Layout root containers
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

        # During learn flows we suppress CLI printing entirely; this avoids
        # confusing output (learn runs many subprocess commands and prompts).
        self._suppress_cli_prints = False

        if not is_admin():
            admin_lbl = ttk.Label(self.topbar, text="Note: Some actions may require Administrator", foreground="#CC6600")
            admin_lbl.pack(side="right")

        # Treeview device list:
        # We intentionally group devices under two parent rows:
        #   - Playback (Render)
        #   - Recording (Capture)
        #
        # Those group rows are "headers" and are not selectable / actionable.
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

        # Group tag for header rows (visual differentiation).
        try:
            self.tree.tag_configure("group", foreground="#202020")
        except Exception:
            pass

        # Remove the expand/collapse indicator element for a cleaner look.
        # We still use groups, but we don't want users to treat them like nodes.
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

        # Context menu:
        # We store menu indices because later we update labels and state with
        # entryconfig() based on live device state (mute/listen/enh/FX).
        self.menu = tk.Menu(self.root, tearoff=0)
        self.menu.add_command(label="Set as Default (all roles)", command=self.on_set_default)
        self.menu.add_separator()
        self.menu.add_command(label="Set Volume...", command=self.on_set_volume)
        self.menu.add_command(label="Mute/Unmute", command=self.on_toggle_mute)
        self.mute_menu_index = self.menu.index("end")

        self.menu.add_separator()
        self.listen_menu_default_label = "Toggle Listen (capture only)"
        self.menu.add_command(label=self.listen_menu_default_label, command=self.on_toggle_listen)
        self.listen_menu_index = self.menu.index("end")

        self.enh_menu_default_label = "Enable Enhancements"
        self.menu.add_command(label=self.enh_menu_default_label, command=self.on_toggle_enhancements)
        self.enh_menu_index = self.menu.index("end")

        # FX submenu is dynamically populated from get-device-state.availableFX.
        self.fx_menu = tk.Menu(self.menu, tearoff=0)
        self.menu.add_cascade(label="Enhancement Effects", menu=self.fx_menu)
        self.fx_cascade_index = self.menu.index("end")

        # Learn Enhancements launches the guided learn flow chooser (main vs FX).
        self.menu.add_command(label="Learn Enhancements", command=self.on_learn_enhancements)
        self.learn_menu_index = self.menu.index("end")

        # Track dynamically added FX menu items (kept for compatibility, and for
        # future cleanup patterns).
        self._dynamic_fx_menu_items = []

        # _pending_enh stores what we *think* the current enhancements state is
        # at menu build time, so the toggle handler can invert it without an
        # extra round-trip.
        self._pending_enh = None

        # Prevent overlapping menu builds when users click rapidly; without this
        # guard, concurrent calls could race and produce inconsistent labels.
        self._menu_build_in_progress = False

        # Bindings
        self.tree.bind("<Button-3>", self.on_right_click)
        self.tree.bind("<ButtonRelease-1>", self.on_left_release)
        self.tree.bind("<Double-1>", self.on_double_click)
        self.root.bind("<F5>", lambda e: self.refresh_devices())
        self.tree.bind("<Button-1>", self.on_left_click, add="+")
        self.tree.bind("<<TreeviewSelect>>", self.on_select_change)

        # Initial load
        self.refresh_devices()
        self.root.after_idle(self.adjust_layout_to_content)

    def is_group_row(self, iid):
        # Group rows are headers; only rows in item_to_device are actionable.
        return iid not in self.item_to_device

    def on_left_click(self, event):
        # Prevent header clicks from selecting/sorting (we don't implement sorting).
        region = self.tree.identify_region(event.x, event.y)
        if region == "heading":
            return "break"
        iid = self.tree.identify_row(event.y)
        if not iid:
            return
        # Prevent selecting group header rows.
        if self.is_group_row(iid):
            return "break"

    def on_select_change(self, event):
        # If a group row ever gets selected (e.g., by keyboard), redirect
        # selection to the first child device row for consistent UX.
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

        This is deliberately conservative: learn flows may open Windows settings
        and return focus frequently; we don't want to spam CLI calls or disrupt
        the UX mid-learn.
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
        # Two-stage refresh:
        #   1) Rebuild device list via `audioctl list --json` (authoritative).
        #   2) Populate device_state_cache incrementally by calling
        #      `get-device-state` for each device. This prevents right-click menu
        #      actions from needing to block on state probes.
        #
        # We clear the cache on each refresh because endpoint IDs can change and
        # we want state to correspond to the current list.
        try:
            from .logging_setup import _dbg
            _dbg("GUI: refresh_devices begin")

            # Build CLI args: audioctl list [--all] --json
            args = ["list", "--json"]
            if self.include_all.get():
                args.insert(1, "--all")

            data = run_audioctl(args, capture_json=True, expect_ok=True)
            self.devices = data.get("devices", [])

            # Clear tree UI and mapping.
            self.item_to_device.clear()
            for item in self.tree.get_children():
                self.tree.delete(item)

            # Split by flow for display. GUI order is name-sorted within flow to
            # match CLI index semantics across the project.
            render_devs = sorted(
                [d for d in self.devices if d["flow"] == "Render"],
                key=lambda x: x["name"].lower()
            )
            capture_devs = sorted(
                [d for d in self.devices if d["flow"] == "Capture"],
                key=lambda x: x["name"].lower()
            )

            # Group header rows.
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

                    # We store a copy so we can annotate GUI-local metadata
                    # without mutating the CLI's raw output.
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

            # Start background state population after the list is visible.
            self._schedule_state_population()

        except Exception as e:
            messagebox.showerror("Error", f"Failed to list devices:\n{e}")
            self.set_status("Failed to refresh devices")

    def _schedule_state_population(self):
        """
        Start or restart incremental population of device_state_cache.

        Why incremental:
          - `get-device-state` can involve COM reads (volume/mute) and registry
            reads (listen/enh/FX). Doing that for all devices at once would
            freeze the UI.
          - By processing one device per Tk "tick" (after()), we keep the UI
            responsive and allow right-click actions immediately (using partial
            cache as it fills).
        """
        # Reset cache and queue: the current device list is authoritative.
        self.device_state_cache.clear()

        # Queue of (id, flow) pairs to populate.
        self._state_queue = [(d["id"], d["flow"]) for d in self.devices]

        # Kick off first step.
        self.root.after(10, self._populate_next_device_state)

    def _populate_next_device_state(self):
        """
        Process one device from the queue:
          - Call CLI `get-device-state` (best-effort).
          - Cache the result if it parses as JSON.
          - Reschedule itself until the queue is empty.
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

        # Schedule next device. Keeping the delay small spreads the cost
        # naturally across UI frames.
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
        # Printing CLI commands is a UX feature for power users (helps learning
        # and scripting). We hard-suppress it during learn flows because learn
        # can emit many intermediate commands/prompts and would clutter output.
        if getattr(self, "_suppress_cli_prints", False):
            return
        if self.print_cmd.get():
            try:
                print(cmd_str)
            except Exception:
                pass

    def _suspend_print_cli_and_disable_checkbox(self):
        # During modal learn workflows, we:
        #   - force Print CLI Commands off
        #   - disable the checkbox to prevent toggling mid-flow
        #   - enable a "hard suppression" flag so no internal prints escape
        try:
            self._prev_print_cmd_val = bool(self.print_cmd.get())
        except Exception:
            self._prev_print_cmd_val = False

        try:
            self.print_cmd.set(False)
        except Exception:
            pass

        try:
            if hasattr(self, "chk_print_cmd") and self.chk_print_cmd:
                self.chk_print_cmd.configure(state="disabled")
        except Exception:
            pass

        self._suppress_cli_prints = True

    def _restore_print_cli_checkbox(self):
        # Restore the Print CLI Commands UI and behavior after a learn flow.
        try:
            self.print_cmd.set(bool(getattr(self, "_prev_print_cmd_val", False)))
        except Exception:
            pass
        try:
            if hasattr(self, "chk_print_cmd") and self.chk_print_cmd:
                self.chk_print_cmd.configure(state="normal")
        except Exception:
            pass
        self._suppress_cli_prints = False

    def get_selected_device(self):
        sel = self.tree.selection()
        if not sel:
            return None
        return self.item_to_device.get(sel[0])

    def _ensure_device_state_entry(self, dev_id, flow):
        """
        Ensure there is a mutable state dict for dev_id in device_state_cache.

        The GUI updates the cache opportunistically after actions succeed so the
        next context-menu build reflects the new state immediately (without
        waiting for the next background get-device-state pass).
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
        if not isinstance(st.get("availableFX"), list):
            st["availableFX"] = []
        return st

    def _refresh_menu_state_async(self, d):
        # Background refresh helper: fetch get-device-state and update menu labels.
        # This is used when we want to update state without blocking the UI thread.
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
                    self.device_state_cache[d["id"]] = st

                    # Listen state label is meaningful only for Capture devices.
                    if d["flow"] == "Capture":
                        ls = st.get("listenEnabled", None)
                        listen_label = "Disable Listen" if ls is True else "Enable Listen" if ls is False else self.listen_menu_default_label
                        self.menu.entryconfig(self.listen_menu_index, label=listen_label, state="normal")

                    # Enhancements: vendor-only. None means "no vendor toggle known".
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

                    # FX submenu: rebuild from availableFX list.
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
        # Context menu build strategy:
        #   1) Use cached device_state_cache for instant labels (fast UX).
        #   2) Do a quick bounded-time `get-device-state` call (quick_json) to
        #      reduce staleness (e.g., if Windows UI changed state externally).
        #   3) Build FX submenu from state.availableFX.
        #
        # We guard against overlapping builds because right-click/double-click
        # can trigger multiple events rapidly and menu.entryconfig is not
        # designed for concurrent mutation.
        try:
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
                # Group header rows: disable everything; no actions apply.
                for i in range(end_idx + 1):
                    etype = self.menu.type(i)
                    if etype in ("command", "cascade", "checkbutton", "radiobutton"):
                        self.menu.entryconfig(i, state="disabled")
                self.menu.tk_popup(event.x_root, event.y_root)
                return

            # Enable standard menu items for a real device row.
            for i in range(end_idx + 1):
                etype = self.menu.type(i)
                if etype in ("command", "cascade", "checkbutton", "radiobutton"):
                    self.menu.entryconfig(i, state="normal")

            # Use cached state first to avoid blocking menu display.
            state = self.device_state_cache.get(d["id"])

            # Mute label: derived from cached muted flag.
            mute_label = "Mute/Unmute"
            if isinstance(state, dict):
                muted = state.get("muted", None)
                if muted is True:
                    mute_label = "Unmute"
                elif muted is False:
                    mute_label = "Mute"
            self.menu.entryconfig(self.mute_menu_index, label=mute_label, state="normal")

            # Listen label: Capture-only.
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

            # Enhancements label: vendor-only; None => disable toggle action.
            self._pending_enh = None
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
                enh_label = "Enhancements (no vendor toggle)"
                enh_state_normalized = None
                enh_state_enabled = False
                enh_menu_state = "disabled"

            # Save decision for the click handler so we don't need to re-fetch.
            self._pending_enh = {
                "id": d["id"],
                "flow": d["flow"],
                "current": enh_state_normalized,
                "supported": enh_state_enabled,
            }
            self.menu.entryconfig(self.enh_menu_index, label=enh_label, state=enh_menu_state)

            # FX submenu from cached state (if any).
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

            # Final step: do a quick best-effort refresh to avoid stale labels.
            # If this fails (timeout/parse), we keep cached labels.
            try:
                st = run_audioctl_quick_json(
                    ["get-device-state", "--id", d["id"], "--flow", d["flow"]],
                    timeout=0.75
                )
                if isinstance(st, dict):
                    self.device_state_cache[d["id"]] = st

                    if d["flow"] == "Capture":
                        ls = st.get("listenEnabled", None)
                        listen_label = "Disable Listen" if ls is True else "Enable Listen" if ls is False else self.listen_menu_default_label
                        self.menu.entryconfig(self.listen_menu_index, label=listen_label, state="normal")

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

        # We delay slightly so selection state settles before we build menu.
        def _open_menu_later():
            if not self.tree.exists(iid):
                return
            self.show_menu_for_item(event, iid=iid)

        self.root.after(50, _open_menu_later)

    def on_set_default(self):
        # Sets system default endpoint via CLI PolicyConfig path.
        # We show an admin warning because some systems require elevation for
        # SetDefaultEndpoint (PolicyConfig behavior varies by Windows build/OEM).
        # On success, we refresh the device list so default flags update.
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
        # Volume set is a 2-step UX:
        #   1) open_volume_dialog reads current volume (get-volume) for a sane
        #      default slider position.
        #   2) set-volume writes the new level and we update cache immediately.
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

                # Cache update avoids stale menu labels until background refresh.
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
        # Toggle mute is state-driven:
        #   - We first query get-volume (muted flag) so we can choose --mute or
        #     --unmute explicitly. This avoids ambiguity and keeps behavior
        #     aligned with CLI semantics.
        d = self.get_selected_device()
        if not d:
            return
        try:
            from .logging_setup import _log

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
        # Listen is capture-only by Windows design. The CLI enforces this too,
        # but we gate it here for UX and clearer messaging.
        #
        # State-driven logic mirrors mute:
        #   - Query get-listen for current state.
        #   - Invert if known; otherwise ask the user.
        #   - Call `audioctl listen --enable/--disable`.
        #   - Update device_state_cache on success.
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
        # Enhancements are *vendor-only* at runtime in this project.
        # That means:
        #   - If the CLI reports enhancementsEnabled=None, there is no learned
        #     vendor method for this device and we must not try Windows fallback.
        #   - We use the menu-captured _pending_enh state (current + supported)
        #     to decide whether to call CLI and which direction to toggle.
        d = self.get_selected_device()
        if not d:
            return
        try:
            from .logging_setup import _log
            _log(f"Enhancements toggle requested for {d['name']} ({d['id']})")

            st = None
            supported = True
            pe = getattr(self, "_pending_enh", None)
            if pe and pe.get("id") == d["id"] and pe.get("flow") == d["flow"]:
                st = pe.get("current", None)
                supported = pe.get("supported", True)

            if not supported:
                messagebox.showinfo(
                    "Enhancements not learned",
                    "No vendor method is available for 'Audio Enhancements' on this device yet.\n\n"
                    "Use 'Learn Enhancements' first, then try again."
                )
                self.set_status("Enhancements: no vendor toggle for this device")
                _log(f"Enhancements toggle aborted (no vendor toggle) for {d['name']} ({d['id']})")
                return

            # Invert known state; otherwise ask user which direction to apply.
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
                # CLI returns {"enhancementsSet": {"enabled": ..., "verifiedBy": ...}}
                # verifiedBy indicates which vendor method confirmed the result
                # (e.g., "vendor:<section>"). GUI shows it only in logs.
                state = enh.get("enabled", enable)
                state_txt = "enabled" if state else "disabled"
                verified_by = enh.get("verifiedBy", "vendor")
                _log(f"Enhancements toggle via CLI for {d['name']} ({d['id']}): final={state_txt} via {verified_by}")
                self.set_status(f"Enhancements {state_txt} for: {d['name']}")

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
        # Learn is a guided workflow that writes vendor_toggles.ini entries.
        # The GUI provides a chooser because we support two learn modes:
        #   - main: learn the global "Audio Enhancements" switch vendor toggle
        #   - fx:   learn a specific effect (e.g., BassBoost) which may require
        #           a multi-write registry plan
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
            # Refresh the list each time the dropdown is opened (INI may change).
            try:
                fx_entry_widget.configure(postcommand=lambda: fx_entry_widget.configure(values=self._load_fx_names_for_combo()))
            except Exception:
                pass
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
        # This is an alternate non-blocking learn UI using LearnRunner.
        # It keeps the entire learn flow inside a Toplevel with a live log and
        # explicit Continue buttons. The key purpose is to avoid freezing the
        # main window while the CLI waits for user prompts.
        self._in_modal_operation = True
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

        txt = tk.Text(frm, height=18, width=100, wrap="word")
        txt.grid(row=0, column=0, columnspan=4, sticky="nsew", pady=(0, 8))
        frm.rowconfigure(0, weight=1)
        frm.columnconfigure(0, weight=1)

        btn_a = ttk.Button(frm, text="Continue (A)", state="disabled")
        btn_b = ttk.Button(frm, text="Continue (B)", state="disabled")
        btn_cancel = ttk.Button(frm, text="Cancel")
        btn_retry = ttk.Button(frm, text="Retry (skip confirmation)", state="disabled")
        btn_a.grid(row=1, column=0, sticky="w", padx=(0, 6))
        btn_b.grid(row=1, column=1, sticky="w", padx=(0, 6))
        btn_cancel.grid(row=1, column=2, sticky="e", padx=(0, 6))
        btn_retry.grid(row=1, column=3, sticky="e")

        # The GUI shows one consolidated warning. The CLI can also warn and ask
        # for "I UNDERSTAND"; in this path we can auto-confirm or set env to skip.
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
                    btn_a.configure(state="disabled")
                    btn_b.configure(state="disabled")
                    btn_retry.configure(state="disabled")

                    # Learn flows interleave prompts and JSON; we scan for the
                    # last meaningful JSON object.
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
                        self.refresh_devices()
                    else:
                        try:
                            messagebox.showerror(
                                "Learn Enhancements",
                                "Learn failed or could not verify a vendor entry.\nSee the log text for details."
                            )
                        except Exception:
                            pass

                    self._in_modal_operation = False
                    try:
                        top.destroy()
                    except Exception:
                        pass

            self.root.after(0, _apply)

        # First attempt: auto-confirm when we see the prompt.
        runner = LearnRunner(args_list, on_output=append_log, on_state=handle_state, confirmed=False)
        btn_a.configure(command=runner.continue_snapshot_a)
        btn_b.configure(command=runner.continue_snapshot_b)

        def do_cancel():
            self._in_modal_operation = False
            try:
                runner.terminate()
            except Exception:
                pass
            top.destroy()

        def do_retry():
            # Retry with env skip (no confirmation prompt). Useful if the CLI's
            # confirmation prompt behavior differs across versions.
            nonlocal runner
            btn_retry.configure(state="disabled")
            try:
                runner.terminate()
            except Exception:
                pass
            new_runner = LearnRunner(args_list, on_output=append_log, on_state=handle_state, confirmed=True)
            btn_a.configure(command=new_runner.continue_snapshot_a, state="disabled")
            btn_b.configure(command=new_runner.continue_snapshot_b, state="disabled")
            self.root.after(0, new_runner.start)
            runner = new_runner

        btn_cancel.configure(command=do_cancel)
        btn_retry.configure(command=do_retry)

        runner.start()
        top.protocol("WM_DELETE_WINDOW", do_cancel)

    def _learn_main_toggle_via_cli(self, d):
        """
        Learn Enhancements (main toggle) via the CLI interactive flow.

        Key behaviors:
          - GUI shows *one* warning, then sets AUDIOCTL_LEARN_CONFIRMED=1 so the
            CLI doesn't stop for its own confirmation prompt (no double warning).
          - Uses run_audioctl_interactive() to:
              * detect CLI instruction lines,
              * show message boxes,
              * send Enter to stdin after the user clicks OK.
          - The CLI may print prompts and JSON on the same output stream; we
            parse the *last* JSON object that contains vendorLearned/vendorAvailable.
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

            # Tell the CLI to skip its stdin "I UNDERSTAND" step because the GUI
            # already handled user confirmation.
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

            # Output may contain both prompt lines and JSON. We scan backward for
            # the last JSON object that includes a vendor result.
            def _extract_last_vendor_json(text: str):
                s = text or ""
                i = len(s) - 1
                while i >= 0:
                    while i >= 0 and s[i] != "}":
                        i -= 1
                    if i < 0:
                        break
                    end = i
                    depth = 1
                    i -= 1
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
        Learn an FX (per-effect) toggle via the CLI interactive flow.

        FX learn uses a two-pass A/B approach:
          - A/B "prime" the driver: some vendors create registry keys only after
            the first toggle.
          - A2/B2 are the authoritative snapshots used to build stable
            multi-write toggles in vendor_toggles.ini.

        Prompt pattern ordering matters:
          - We match second-pass instruction lines first (Enable/Disable again)
            because they can overlap with first-pass wording ("Enable the ...").
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

            # Prompts reordered and tightened so that second pass (A2/B2) is matched first.
            prompt_patterns = [
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

            rc, out, err = run_audioctl_interactive(args, prompt_patterns, expect_ok=False)

            # Extract fxLearned JSON from output. Learn can produce many lines;
            # we scan from the bottom for robustness.
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
        # Volume dialog design:
        #   - A 0-100 slider and a 3-digit entry box.
        #   - They stay in sync (with re-entrancy guards).
        #   - Initial value comes from CLI get-volume so the dialog reflects
        #     current system state.
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
        syncing = {"entry": False, "scale": False}

        def _validate(P):
            # Strict entry validation: keep it numeric and within 0..100.
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
            # Scale callback provides strings; update entry value safely.
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
            # Entry trace updates scale; guard avoids feedback loop.
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
        Toggle an FX via CLI.

        FX toggles are routed through the `enhancements` command:
          - `enhancements --enable-fx "<name>"` or `--disable-fx "<name>"`

        We decide direction based on the menu-captured current_state:
          - True  => user likely wants to disable
          - False => user likely wants to enable
          - None  => default to enable (safe "turn it on" action)

        On success, we update device_state_cache.availableFX for the chosen
        effect so labels stay correct immediately.
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

                st = self._ensure_device_state_entry(d["id"], d["flow"])
                fx_list = st.get("availableFX") or []
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
        # Provide autocomplete suggestions from vendor_toggles.ini so users can
        # reuse consistent FX names across devices.
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
        # Lightweight "starts-with" autocomplete for effect names; avoids adding
        # dependencies or complex UI logic.
        def on_keyrelease(ev):
            text = combo.get()
            if not text:
                return
            vals = list(combo.cget("values") or [])
            low = text.lower()
            for v in vals:
                s = (v or "")
                if s.lower().startswith(low):
                    combo.set(s)
                    try:
                        combo.icursor(len(text))
                        combo.select_range(len(text), tk.END)
                    except Exception:
                        pass
                    break

        combo.bind("<KeyRelease>", on_keyrelease, add="+")


def launch_gui():
    # GUI entrypoint used by CLI `main()` when no CLI args are provided.
    # It sets up:
    #   - Icon loading (best-effort; won't crash if missing)
    #   - Global Tk exception handler that writes to our log file
    #   - Shutdown hooks that log close events for debugging "silent exits"
    _log("launch_gui: creating Tk root")
    root = tk.Tk()
    try:
        if sys.platform.startswith("win"):
            root.iconbitmap(resource_path("audio.ico"))
    except Exception:
        pass

    gui = AudioGUI(root)

    def _on_root_close():
        # Explicit hook gives us a breadcrumb in logs when the user closes window.
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
        # Tkinter swallows exceptions by default; by overriding this we ensure
        # failures are logged and the user gets a useful pointer to the log file.
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
