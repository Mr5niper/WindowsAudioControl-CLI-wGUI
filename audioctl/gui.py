# audioctl/gui.py
import sys
import io
import time
import tkinter as tk
from tkinter import ttk, messagebox
from contextlib import redirect_stderr
import comtypes
import re
from .logging_setup import resource_path, _log, _log_exc, _log_path
from .compat import is_admin
from .devices import (
    list_devices, _sort_and_tag_gui_indices,
    get_endpoint_mute, set_endpoint_mute,
    get_endpoint_volume, set_endpoint_volume,
    set_default_endpoint,
    _get_listen_to_device_status_ps, _read_listen_enable_from_registry,
    _verify_listen_via_registry, set_listen_to_device_ps,
    _collect_sysfx_snapshot, _diff_mmdevices_lists,
    _reemit_non_error_stderr,
)
from .vendor_db import (
    _vendor_ini_default_path,
    _find_first_vendor_entry,
    _get_enhancements_status_any,
    _append_vendor_ini_entry_if_missing,
    _append_fx_ini_entry,             # ADDED
    _sanitize_ini_section_name,
    _build_vendor_ini_snippet,
    _apply_enhancements,
    _enhancements_supported
)
class AudioGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Audio Control v1.4.4.0 12-31-2025")
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
        # Variables
        self.include_all = tk.BooleanVar(value=False)
        self.print_cmd = tk.BooleanVar(value=False)
        self.devices = []
        self.item_to_device = {}
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
        ttk.Checkbutton(
            self.topbar,
            text="Print CLI commands",
            variable=self.print_cmd
        ).pack(side="left", padx=(10, 0))
        if not is_admin():
            admin_lbl = ttk.Label(self.topbar, text="Note: Some actions may require Administrator", foreground="#CC6600")
            admin_lbl.pack(side="right")
        # Treeview
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
        # Remove indicator element
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
        self.menu.add_command(label="Mute", command=self.on_toggle_mute)
        self.mute_menu_index = self.menu.index("end")
        self.menu.add_separator()
        self.listen_menu_default_label = "Toggle Listen (capture only)"
        self.menu.add_command(label=self.listen_menu_default_label, command=self.on_toggle_listen)
        self.listen_menu_index = self.menu.index("end")
        self.enh_menu_default_label = "Enable Enhancements"
        self.menu.add_command(label=self.enh_menu_default_label, command=self.on_toggle_enhancements)
        self.enh_menu_index = self.menu.index("end")
        # NEW: Enhancement Effects cascade submenu (populated dynamically)
        self.fx_menu = tk.Menu(self.menu, tearoff=0)
        self.menu.add_cascade(label="Enhancement Effects", menu=self.fx_menu)
        self.fx_cascade_index = self.menu.index("end")
        # Learn Enhancements remains after the cascade
        self.menu.add_command(label="Learn Enhancements", command=self.on_learn_enhancements)
        self.learn_menu_index = self.menu.index("end")
        # NEW: Track dynamically added FX menu items
        self._dynamic_fx_menu_items = []
        self._pending_enh = None
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
        return iid not in self.item_to_device
    def on_left_click(self, event):
        region = self.tree.identify_region(event.x, event.y)
        if region == "heading":
            return "break"
        iid = self.tree.identify_row(event.y)
        if not iid:
            return
        if self.is_group_row(iid):
            return "break"
    def on_select_change(self, event):
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
            self.devices = list_devices(include_all=self.include_all.get())
            _dbg("GUI: refresh_devices end")
            self.item_to_device.clear()
            for item in self.tree.get_children():
                self.tree.delete(item)
            render_devs = sorted([d for d in self.devices if d["flow"] == "Render"], key=lambda x: x["name"].lower())
            capture_devs = sorted([d for d in self.devices if d["flow"] == "Capture"], key=lambda x: x["name"].lower())
            grp_render = self.tree.insert("", "end", text="Playback (Render)", values=("", "", "", "", ""), open=True, tags=("group",))
            grp_capture = self.tree.insert("", "end", text="Recording (Capture)", values=("", "", "", "", ""), open=True, tags=("group",))
            def insert_group(parent, devs, flow_name):
                for idx, d in enumerate(devs):
                    flags = [k for k, v in d["isDefault"].items() if v]
                    defaults_txt = ", ".join(flags) if flags else "-"
                    d_copy = dict(d)
                    d_copy["_index"] = idx
                    d_copy["_group"] = flow_name
                    iid = self.tree.insert(parent, "end", text="", values=(idx, d["name"], d["flow"], defaults_txt, d["id"]))
                    self.item_to_device[iid] = d_copy
            insert_group(grp_render, render_devs, "Render")
            insert_group(grp_capture, capture_devs, "Capture")
            self.set_status("Device list updated")
            self.adjust_layout_to_content()
            self.root.after_idle(self.adjust_layout_to_content)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to list devices:\n{e}")
            self.set_status("Failed to refresh devices")
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
        if self.print_cmd.get():
            try:
                print(cmd_str)
            except Exception:
                pass
    def get_selected_device(self):
        sel = self.tree.selection()
        if not sel:
            return None
        return self.item_to_device.get(sel[0])
    def show_menu_for_item(self, event, iid=None):
        try:
            if iid is None:
                iid = self.tree.identify_row(event.y)
            if not iid:
                return
            d = self.item_to_device.get(iid)
            if d:
                self.tree.selection_set(iid)
            else:
                self.tree.selection_remove(iid)
            # Determine current end (not used to trim now; we rebuild only the FX submenu)
            end_idx = self.menu.index("end")
            end_idx = end_idx if end_idx is not None else -1
            if not d:
                # Disable actionable items when no device is selected
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
            # Mute label
            try:
                muted = get_endpoint_mute(d["id"])
            except Exception:
                muted = None
            if muted is True:
                mute_label = "Unmute"
            elif muted is False:
                mute_label = "Mute"
            else:
                mute_label = "Unmute"
            self.menu.entryconfig(self.mute_menu_index, label=mute_label, state="normal")
            # Listen label (Capture only)
            if d["flow"] == "Capture":
                try:
                    current = _get_listen_to_device_status_ps(d["id"])
                except Exception:
                    current = None
                if current is None:
                    current = _read_listen_enable_from_registry(d["id"])
                if current is True:
                    label = "Disable Listen"
                elif current is False:
                    label = "Enable Listen"
                else:
                    label = self.listen_menu_default_label
                self.menu.entryconfig(self.listen_menu_index, label=label, state="normal")
            else:
                self.menu.entryconfig(self.listen_menu_index, label=self.listen_menu_default_label, state="disabled")
            # Main enhancements toggle label/state
            vend_available = False
            try:
                vend_available = bool(_find_first_vendor_entry(d["id"], d["flow"], ini_path=_vendor_ini_default_path()))
            except Exception:
                vend_available = False
            if vend_available:
                try:
                    enh = _get_enhancements_status_any(d["id"], d["flow"])
                except Exception:
                    enh = None
                if enh is True:
                    enh_label = "Disable Enhancements"
                    target_enable_next = False
                elif enh is False:
                    enh_label = "Enable Enhancements"
                    target_enable_next = True
                else:
                    enh_label = "Enable Enhancements"
                    target_enable_next = True
                self._pending_enh = {"id": d["id"], "enable": target_enable_next}
                self.menu.entryconfig(self.enh_menu_index, label=enh_label, state="normal")
            else:
                self._pending_enh = None
                self.menu.entryconfig(self.enh_menu_index, label=self.enh_menu_default_label, state="disabled")
            # Rebuild Enhancement Effects submenu safely
            try:
                self.fx_menu.delete(0, "end")
            except Exception:
                # If the menu hasn't any items yet, ignore
                pass
            try:
                from .vendor_db import _list_fx_for_device, _read_vendor_entry_state
                fx_list = _list_fx_for_device(d["id"], d["flow"], ini_path=_vendor_ini_default_path())
                if fx_list:
                    for fx_info in fx_list:
                        fx_name = fx_info["fx_name"]
                        entry = fx_info["entry"]
                        try:
                            current_state = _read_vendor_entry_state(entry, d["id"], d["flow"])
                        except Exception:
                            current_state = None
                        if current_state is True:
                            label = f"Disable {fx_name}"
                        elif current_state is False:
                            label = f"Enable {fx_name}"
                        else:
                            label = f"Toggle {fx_name}"
                        def make_fx_command(fx_n):
                            def cmd():
                                self.on_toggle_fx_live(fx_n)
                            return cmd
                        self.fx_menu.add_command(label=label, command=make_fx_command(fx_name))
                    # Enable the cascade
                    self.menu.entryconfig(self.fx_cascade_index, state="normal")
                else:
                    self.fx_menu.add_command(label="No effects available", state="disabled")
                    self.menu.entryconfig(self.fx_cascade_index, state="disabled")
            except Exception as e:
                try:
                    from .logging_setup import _log
                    _log(f"Failed to build FX submenu for {d['name']}: {e}")
                except Exception:
                    pass
                self.fx_menu.add_command(label="Failed to load effects", state="disabled")
                self.menu.entryconfig(self.fx_cascade_index, state="disabled")
            # Show menu
            self.menu.tk_popup(event.x_root, event.y_root)
        finally:
            try:
                self.menu.grab_release()
            except Exception:
                pass
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
        self.show_menu_for_item(event, iid=iid)
    def on_set_default(self):
        d = self.get_selected_device()
        if not d:
            return
        try:
            from .logging_setup import _log, _dbg
            _dbg(f"GUI: on_set_default for id={d['id']} flow={d['flow']}")
            if d["flow"] == "Render":
                cmd = f'audioctl set-default --playback-id "{d["id"]}" --playback-role all'
            else:
                cmd = f'audioctl set-default --recording-id "{d["id"]}" --recording-role all'
            _log(f"GUI action: set-default start id={d['id']} name={d['name']} flow={d['flow']} roles=all")
            if not is_admin():
                if not messagebox.askyesno(
                    "Administrator recommended",
                    "Setting default device may require Administrator privileges on some systems.\n\nContinue?"
                ):
                    _log(f"GUI action: set-default cancelled id={d['id']} name={d['name']}")
                    return
            set_default_endpoint(d["id"], "all")
            self.maybe_print_cli(cmd)
            self.set_status(f"Set default ({d['flow']}) device: {d['name']} (all roles)")
            self.refresh_devices()
            _log(f"GUI action: set-default success id={d['id']} name={d['name']} flow={d['flow']}")
            _dbg(f"GUI: on_set_default successful")
        except Exception as e:
            from .logging_setup import _log
            _log(f"GUI action: set-default failed id={d['id']} name={d['name']} error={e}")
            messagebox.showerror("Error", f"Failed to set default:\n{e}")
            self.set_status("Failed to set default")
    def on_set_volume(self):
        d = self.get_selected_device()
        if not d:
            return
        try:
            from .logging_setup import _log
            level = self.open_volume_dialog(d["id"], d["name"])
            if level is None:
                _log(f"GUI action: set-volume cancelled id={d['id']} name={d['name']}")
                return
            _log(f"GUI action: set-volume start id={d['id']} name={d['name']} flow={d['flow']} level={level}")
            ok = set_endpoint_volume(d["id"], level)
            if ok:
                cmd = f'audioctl set-volume --id "{d["id"]}" --flow {d["flow"]} --level {level}'
                self.maybe_print_cli(cmd)
                self.set_status(f"Volume set to {level}% for: {d['name']}")
                _log(f"GUI action: set-volume success id={d['id']} name={d['name']} level={level}")
            else:
                _log(f"GUI action: set-volume failed id={d['id']} name={d['name']} level={level}")
                messagebox.showerror("Error", "Failed to set volume/mute")
                self.set_status("Failed to set volume")
        except Exception as e:
            from .logging_setup import _log
            _log(f"GUI action: set-volume error id={d['id']} name={d['name']} err={e}")
            messagebox.showerror("Error", f"Failed to set volume:\n{e}")
            self.set_status("Failed to set volume")
    def on_toggle_mute(self):
        d = self.get_selected_device()
        if not d:
            return
        try:
            from .logging_setup import _log
            current = get_endpoint_mute(d["id"])
            target = False if current is None else (not bool(current))
            _log(f"GUI action: toggle-mute start id={d['id']} name={d['name']} current={current} target={target}")
            ok = set_endpoint_mute(d["id"], target)
            if ok:
                cmd = f'audioctl set-volume --id "{d["id"]}" --flow {d["flow"]} --{"mute" if target else "unmute"}'
                self.maybe_print_cli(cmd)
                self.set_status(f'{"Muted" if target else "Unmuted"}: {d["name"]}')
                _log(f"GUI action: toggle-mute success id={d['id']} name={d['name']} final={target}")
            else:
                _log(f"GUI action: toggle-mute failed id={d['id']} name={d['name']} target={target}")
                messagebox.showerror("Error", "Failed to change mute state")
                self.set_status("Failed to change mute state")
        except Exception as e:
            from .logging_setup import _log
            _log(f"GUI action: toggle-mute error id={d['id']} name={d['name']} err={e}")
            messagebox.showerror("Error", f"Failed to toggle mute:\n{e}")
            self.set_status("Failed to toggle mute")
    def on_toggle_listen(self):
        d = self.get_selected_device()
        if not d:
            return
        if d["flow"] != "Capture":
            messagebox.showinfo("Not a capture device", "Listen can only be toggled for capture (recording) devices.")
            return
        try:
            _log(f"Listen toggle requested for {d['name']} ({d['id']})")
            current = _get_listen_to_device_status_ps(d["id"])
            if current is None:
                current = _read_listen_enable_from_registry(d["id"])
            enable = not bool(current)
            cmd = f'audioctl listen --id "{d["id"]}" --{"enable" if enable else "disable"}'
            self.maybe_print_cli(cmd)
            captured_stderr = io.StringIO()
            with redirect_stderr(captured_stderr):
                ok = set_listen_to_device_ps(d["id"], enable, render_device_id=None)
            if not ok:
                actual = _get_listen_to_device_status_ps(d["id"])
                if actual is None:
                    verified, reg_state = _verify_listen_via_registry(d["id"], enable, timeout=3.0, interval=0.20)
                    actual = reg_state if verified or reg_state is not None else None
            else:
                actual = _get_listen_to_device_status_ps(d["id"])
                if actual is None:
                    verified, reg_state = _verify_listen_via_registry(d["id"], enable, timeout=3.0, interval=0.20)
                    actual = reg_state if verified or reg_state is not None else None
            if actual is None:
                _log(f"Listen toggle result unknown for {d['name']} ({d['id']}); requested={enable}")
                messagebox.showwarning("Listen status unknown", "Could not verify final 'Listen' state. It may still have applied.")
                self.set_status(f"Listen toggle requested for: {d['name']}")
            else:
                _log(f"Listen toggle result for {d['name']} ({d['id']}): final={actual}")
                state_txt = "enabled" if actual else "disabled"
                self.set_status(f"Listen {state_txt} for: {d['name']}")
        except Exception as e:
            _log(f"Listen toggle exception for {d['name']} ({d['id']}): {e!r}")
            messagebox.showerror("Error", f"Failed to toggle Listen:\n{e}")
            self.set_status("Failed to toggle Listen")
    def on_toggle_enhancements(self):
        d = self.get_selected_device()
        if not d:
            return
        try:
            _log(f"Enhancements toggle requested for {d['name']} ({d['id']})")
            if getattr(self, "_pending_enh", None) and self._pending_enh.get("id") == d["id"]:
                enable = bool(self._pending_enh["enable"])
            else:
                current = _get_enhancements_status_any(d["id"], d["flow"])
                enable = True if current is None else (not bool(current))
            if not _enhancements_supported(d["id"], d["flow"]):
                from .logging_setup import _log as _ilog
                _ilog(f"GUI action: enhancements not-supported id={d['id']} name={d['name']} flow={d['flow']}")
                messagebox.showinfo("Not supported", "This endpoint does not have a configured vendor toggle for 'Audio Enhancements'. Use 'Learn Enhancements' first.")
                self.set_status("Enhancements toggle failed: No vendor method.")
                return
            self.maybe_print_cli(f'audioctl enhancements --id "{d["id"]}" --flow {d["flow"]} --{"enable" if enable else "disable"}')
            ok, verified_by, state = _apply_enhancements(
                d["id"], d["flow"], enable,
                prefer_hklm=is_admin(),
                allow_universal_scan=False,
                vendor_ini_path=_vendor_ini_default_path()
            )
            self._pending_enh = None
            if ok and (state is None or state == enable):
                state_txt = "enabled" if state else "disabled"
                _log(f"Enhancements toggle result for {d['name']} ({d['id']}): final={state_txt} via {verified_by}")
                if verified_by.startswith("vendor"):
                    self.set_status(f"Enhancements {state_txt} for: {d['name']} (Vendor controlled)")
                else:
                    self.set_status(f"Enhancements {state_txt} for: {d['name']} (Unknown success path)")
            else:
                from .logging_setup import _log as _ilog
                _ilog(f"GUI action: enhancements failed-or-unverified id={d['id']} name={d['name']} flow={d['flow']} requested_enable={enable} ok={ok} state={state} verified_by={verified_by}")
                messagebox.showwarning("Could not verify", "Vendor toggle applied but could not verify final state, or the toggle failed.")
                self.set_status(f"Enhancements toggle requested for: {d['name']} (Verification failed)")
        except Exception as e:
            _log(f"Enhancements toggle exception for {d['name']} ({d['id']}): {e!r}")
            messagebox.showerror("Error", f"Failed to toggle Enhancements:\n{e}")
            self.set_status("Failed to toggle Enhancements")
    def _run_elevated_vendor_ini_append(self, work_dict):
        """
        Elevate a tiny CLI helper to append to vendor_toggles.ini at a protected path.
        We write a temp JSON work order, then ShellExecute 'runas' audioctl vendor-ini-append --work <json>.
        """
        import tempfile, json, os, sys, ctypes
        try:
            # Write work order file
            tmp_dir = tempfile.gettempdir()
            work_path = os.path.join(tmp_dir, f"audioctl_work_{int(time.time())}.json")
            with open(work_path, "w", encoding="utf-8") as f:
                json.dump(work_dict, f, indent=2)
            # Build elevated command
            if getattr(sys, "frozen", False):
                exe = sys.executable
                params = f'vendor-ini-append --work "{work_path}"'
            else:
                exe = sys.executable
                params = f'-m audioctl vendor-ini-append --work "{work_path}"'
            # ShellExecuteW returns >32 on success
            ret = ctypes.windll.shell32.ShellExecuteW(None, "runas", exe, params, None, 1)
            if ret <= 32:
                messagebox.showerror("Elevation failed", "Could not start elevated helper. Please run as Administrator or choose a user-writable INI path.")
                return False
            messagebox.showinfo("Elevated write started", "An elevated helper was launched to write vendor_toggles.ini.\nIf you see a UAC prompt, click Yes.")
            return True
        except Exception as e:
            messagebox.showerror("Elevation error", f"Failed to start elevated helper:\n{e}")
            return False
    def on_learn_enhancements(self):
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
            fx_entry_widget = ttk.Entry(fx_name_frame, textvariable=fx_name_var, width=25)
            fx_entry_widget.pack(side="left")
            fx_entry_widget.config(state="disabled")
            
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
                self._learn_main_toggle(d)
            else:
                self._learn_fx_toggle(d, result["fx_name"])
        
        except Exception as e:
            messagebox.showerror("Error", f"Learn failed:\n{e}")
            self.set_status("Learn failed")
    def _learn_main_toggle(self, d):
        """Existing learn logic (main Enhancements switch)."""
        from .logging_setup import _log
        ini_path = _vendor_ini_default_path()
        warn_txt = (
            "READ CAREFULLY\n\n"
            "This Learn mode will capture two registry snapshots and write a vendor entry into:\n"
            f"  {ini_path}\n\n"
            "From now on, future 'Enhancements' commands for this device WILL WRITE registry values on this machine "
            "(HKCU/optional HKLM) to toggle Enhancements. This is persistent until you manually remove the learned section.\n\n"
            "Critical rules during Learn:\n"
            "- Do NOT change any other audio settings.\n"
            "- Do NOT switch default devices.\n"
            "- Do NOT open other audio/control apps.\n"
            "- Only toggle 'Audio Enhancements' for THIS device exactly when asked.\n\n"
            "Click OK to continue, or Cancel to abort."
        )
        if not messagebox.askokcancel("Warning – Learn writes registry (persistent)", warn_txt):
            self.set_status("Learn Enhancements: aborted by user")
            _log(f"GUI action: learn-main cancelled id={d['id']} name={d['name']}")
            return
        # Print CLI equivalent for learn
        self.maybe_print_cli(f'audioctl enhancements --id "{d["id"]}" --flow {d["flow"]} --learn')
        _log(f"GUI action: learn-main start id={d['id']} name={d['name']} flow={d['flow']} ini={ini_path}")
        messagebox.showinfo(
            "Learn Enhancements - Step 1",
            "Set 'Audio Enhancements' to ENABLED for this device in Windows Sound settings.\n\nClick OK to capture snapshot A."
        )
        captured_stderr = io.StringIO()
        with redirect_stderr(captured_stderr):
            snapA = _collect_sysfx_snapshot(d["id"])
        _reemit_non_error_stderr(captured_stderr.getvalue())
        messagebox.showinfo(
            "Learn Enhancements - Step 2",
            "Set 'Audio Enhancements' to DISABLED for the same device.\n\nClick OK to capture snapshot B."
        )
        captured_stderr = io.StringIO()
        with redirect_stderr(captured_stderr):
            snapB = _collect_sysfx_snapshot(d["id"])
        _reemit_non_error_stderr(captured_stderr.getvalue())
        diffs = _diff_mmdevices_lists(snapA.get("registry") or [], snapB.get("registry") or [])
        snippet, picked = _build_vendor_ini_snippet(d, snapA, snapB, diffs)
        if not picked:
            _log(f"GUI action: learn-main no-dword-flip id={d['id']} name={d['name']}")
            # First vendor availability check
            try:
                vend_entry = _find_first_vendor_entry(d["id"], d["flow"], ini_path=_vendor_ini_default_path())
            except Exception:
                vend_entry = None
            if vend_entry:
                # Vendor path is already available (INI or built-in) due to your toggle
                messagebox.showinfo(
                    "Enhancements vendor available",
                    "A vendor method was detected for this device after your toggle.\n\n"
                    "No INI write was needed. You can now use 'Enable/Disable Enhancements' from the menu."
                )
                self.set_status("Vendor method available (no INI required)")
                return
            # Guide user to try again (drivers often materialize keys only after the first toggle)
            retry = messagebox.askyesno(
                "Try again?",
                "No DWORD flip was detected.\n\n"
                "This may be the first time this endpoint was toggled. Drivers often create keys only after the first toggle.\n\n"
                "Would you like to try again? Please toggle 'Audio Enhancements' to ENABLED, then DISABLED when prompted."
            )
            if not retry:
                self.set_status("Learn Enhancements: user skipped retry")
                return
            # Second pass: capture snapshots again
            messagebox.showinfo(
                "Learn Enhancements - Retry Step 1",
                "Set 'Audio Enhancements' to ENABLED for this device.\n\nClick OK to capture snapshot A."
            )
            captured_stderr = io.StringIO()
            with redirect_stderr(captured_stderr):
                snapA2 = _collect_sysfx_snapshot(d["id"])
            _reemit_non_error_stderr(captured_stderr.getvalue())
            messagebox.showinfo(
                "Learn Enhancements - Retry Step 2",
                "Set 'Audio Enhancements' to DISABLED for the same device.\n\nClick OK to capture snapshot B."
            )
            captured_stderr = io.StringIO()
            with redirect_stderr(captured_stderr):
                snapB2 = _collect_sysfx_snapshot(d["id"])
            _reemit_non_error_stderr(captured_stderr.getvalue())
            diffs2 = _diff_mmdevices_lists(snapA2.get("registry") or [], snapB2.get("registry") or [])
            snippet2, picked2 = _build_vendor_ini_snippet(d, snapA2, snapB2, diffs2)
            if picked2:
                # Proceed with writing INI (same as original success path)
                value_name2 = picked2["name"]
                dword_enable2 = int(picked2["before"])
                dword_disable2 = int(picked2["after"])
                section_name2 = _sanitize_ini_section_name(value_name2)
                notes2 = f"Auto-learned (retry) on '{d['name']}' ({d['flow']}). A=enabled,B=disabled."
                try:
                    res = _append_vendor_ini_entry_if_missing(
                        _vendor_ini_default_path(), section_name2, value_name2,
                        dword_enable2, dword_disable2,
                        flows="Render,Capture", hives="HKCU,HKLM", notes=notes2
                    )
                    if res == "exists":
                        messagebox.showinfo(
                            "Learn Enhancements",
                            f"Vendor section already exists:\n{_vendor_ini_default_path()}\n\nSection: [{section_name2}]\nNo changes were made."
                        )
                        self.set_status("Learn Enhancements: entry already exists (retry)")
                    else:
                        messagebox.showinfo(
                            "Learn Enhancements",
                            f"Learned vendor toggle (retry) and appended to:\n{_vendor_ini_default_path()}\n\n"
                            f"Section: [{section_name2}]\nValue: {value_name2}\nEnabled={dword_enable2}, Disabled={dword_disable2}"
                        )
                        self.set_status("Learn Enhancements: vendor INI updated (retry)")
                except PermissionError:
                    # Elevation or fallback to user INI (same as original block)
                    if messagebox.askyesno(
                        "Permission denied",
                        f"Cannot write INI at:\n{_vendor_ini_default_path()}\n\n"
                        "Choose Yes to attempt elevation and write here.\n"
                        "Choose No to save into a user-writable location instead."
                    ):
                        work = {
                            "kind": "main",
                            "ini_path": _vendor_ini_default_path(),
                            "section": section_name2,
                            "value_name": value_name2,
                            "dword_enable": int(dword_enable2),
                            "dword_disable": int(dword_disable2),
                            "flows": "Render,Capture",
                            "hives": "HKCU,HKLM",
                            "notes": notes2,
                        }
                        ok = self._run_elevated_vendor_ini_append(work)
                        if ok:
                            self.set_status("Learn Enhancements: elevated write started (retry)")
                        else:
                            self.set_status("Learn Enhancements: elevation failed (retry)")
                    else:
                        user_ini = _vendor_ini_default_path()
                        try:
                            res = _append_vendor_ini_entry_if_missing(
                                user_ini, section_name2, value_name2,
                                dword_enable2, dword_disable2,
                                flows="Render,Capture", hives="HKCU,HKLM", notes=notes2
                            )
                            messagebox.showinfo("Saved to user INI", f"INI entry written to:\n{user_ini}\nSection: [{section_name2}]")
                            self.set_status("Learn Enhancements: saved to user INI (retry)")
                        except Exception as e2:
                            messagebox.showerror("Write failed", f"Could not write to user INI either:\n{e2}")
                            self.set_status("Learn Enhancements: write failed (retry)")
                except Exception as e:
                    messagebox.showerror("Error", f"Failed to write INI (retry): {e}")
                    self.set_status("Learn Enhancements: write failed (retry)")
                return  # retry path done
            # Still no DWORD flip after retry: vendor availability final check
            try:
                vend_entry3 = _find_first_vendor_entry(d["id"], d["flow"], ini_path=_vendor_ini_default_path())
            except Exception:
                vend_entry3 = None
            if vend_entry3:
                messagebox.showinfo(
                    "Enhancements vendor available",
                    "A vendor method is now available for this device after your toggles.\n\n"
                    "No INI write was needed. You can now use 'Enable/Disable Enhancements'."
                )
                self.set_status("Vendor method available (retry)")
                return
            # Nothing found
            messagebox.showwarning(
                "Learn Enhancements",
                "No suitable REG_DWORD flip found under FxProperties and no vendor method became available after retry."
            )
            self.set_status("Learn Enhancements: no DWORD flip and no vendor method after retry")
            return
        value_name = picked["name"]
        dword_enable = int(picked["before"])
        dword_disable = int(picked["after"])
        section_name = _sanitize_ini_section_name(value_name)
        notes = f"Auto-learned (manual UI) on '{d['name']}' ({d['flow']}). A=enabled,B=disabled."
        try:
            res = _append_vendor_ini_entry_if_missing(
                ini_path, section_name, value_name,
                dword_enable, dword_disable,
                flows="Render,Capture", hives="HKCU,HKLM", notes=notes
            )
            if res == "exists":
                messagebox.showinfo(
                    "Learn Enhancements",
                    f"Vendor section already exists:\n{ini_path}\n\nSection: [{section_name}]\nNo changes were made."
                )
                self.set_status("Learn Enhancements: entry already exists")
                _log(f"GUI action: learn-main exists id={d['id']} name={d['name']} section={section_name}")
            else:
                messagebox.showinfo(
                    "Learn Enhancements",
                    f"Learned vendor toggle and appended to:\n{ini_path}\n\nSection: [{section_name}]\nValue: {value_name}\nEnabled={dword_enable}, Disabled={dword_disable}"
                )
                self.set_status("Learn Enhancements: vendor INI updated")
                _log(f"GUI action: learn-main success id={d['id']} name={d['name']} section={section_name} value={value_name} enable={dword_enable} disable={dword_disable} ini={ini_path}")
        except PermissionError:
            # Offer elevation or fallback
            if messagebox.askyesno(
                "Permission denied",
                f"Cannot write INI at:\n{ini_path}\n\n"
                "Choose Yes to attempt elevation and write here.\n"
                "Choose No to save into a user-writable location instead."
            ):
                # Elevate just the write using a work order
                work = {
                    "kind": "main",
                    "ini_path": ini_path,
                    "section": section_name,
                    "value_name": value_name,
                    "dword_enable": int(dword_enable),
                    "dword_disable": int(dword_disable),
                    "flows": "Render,Capture",
                    "hives": "HKCU,HKLM",
                    "notes": notes,
                }
                ok = self._run_elevated_vendor_ini_append(work)
                if ok:
                    self.set_status("Learn Enhancements: elevated write started")
                    _log(f"GUI action: learn-main elevation-started id={d['id']} name={d['name']} section={section_name}")
                else:
                    self.set_status("Learn Enhancements: elevation failed")
                    _log(f"GUI action: learn-main elevation-failed id={d['id']} name={d['name']} section={section_name}")
            else:
                # Fallback to user-writable INI (use module-level import)
                user_ini = _vendor_ini_default_path()
                try:
                    res = _append_vendor_ini_entry_if_missing(
                        user_ini, section_name, value_name,
                        dword_enable, dword_disable,
                        flows="Render,Capture", hives="HKCU,HKLM", notes=notes
                    )
                    messagebox.showinfo("Saved to user INI", f"INI entry written to:\n{user_ini}\nSection: [{section_name}]")
                    self.set_status("Learn Enhancements: saved to user INI")
                    _log(f"GUI action: learn-main success-user-ini id={d['id']} name={d['name']} section={section_name} ini={user_ini}")
                except Exception as e2:
                    messagebox.showerror("Write failed", f"Could not write to user INI either:\n{e2}")
                    self.set_status("Learn Enhancements: write failed")
                    _log(f"GUI action: learn-main user-ini-failed id={d['id']} name={d['name']} err={e2}")
        except Exception as e:
            _log(f"GUI action: learn-main failed id={d['id']} name={d['name']} err={e}")
            messagebox.showerror("Error", f"Failed to write INI: {e}")
            self.set_status("Learn Enhancements: write failed")
    def _learn_fx_toggle(self, d, fx_name):
        """Learn a specific FX effect using GUI prompts and core logic."""
        from .vendor_db import _learn_fx_and_write_ini, _vendor_ini_default_path
        from .logging_setup import _log
        ini_path = _vendor_ini_default_path()
        warn_txt = (
            "READ CAREFULLY\n\n"
            f"This Learn mode will capture two registry snapshots for the effect '{fx_name}' "
            f"and write a vendor entry into:\n  {ini_path}\n\n"
            "From now on, this effect WILL BE CONTROLLABLE via registry writes.\n\n"
            "Critical rules during Learn:\n"
            f"- ONLY toggle the '{fx_name}' checkbox/effect\n"
            "- Do NOT toggle the main 'Audio Enhancements' switch\n"
            "- Do NOT change any other audio settings\n"
            "- Do NOT switch devices\n\n"
            "Click OK to continue, or Cancel to abort."
        )
        if not messagebox.askokcancel(f"Warning – Learn FX '{fx_name}'", warn_txt):
            self.set_status(f"Learn FX '{fx_name}': aborted by user")
            _log(f"GUI action: learn-fx cancelled id={d['id']} name={d['name']} fx={fx_name}")
            return
        # Print CLI equivalent for FX learn
        self.maybe_print_cli(f'audioctl enhancements --id "{d["id"]}" --flow {d["flow"]} --learn-fx "{fx_name}"')
        _log(f"GUI action: learn-fx start id={d['id']} name={d['name']} flow={d['flow']} fx={fx_name} ini={ini_path}")
        messagebox.showinfo(
            f"Learn FX '{fx_name}' - Step 1",
            f"ENABLE the '{fx_name}' effect for this device.\n"
            "(Do NOT toggle the main switch, only this specific effect)\n\n"
            "Click OK to capture snapshot A."
        )
        captured_stderr = io.StringIO()
        with redirect_stderr(captured_stderr):
            snapA = _collect_sysfx_snapshot(d["id"])
        _reemit_non_error_stderr(captured_stderr.getvalue())
        messagebox.showinfo(
            f"Learn FX '{fx_name}' - Step 2",
            f"DISABLE the '{fx_name}' effect for the same device.\n\n"
            "Click OK to capture snapshot B."
        )
        captured_stderr = io.StringIO()
        with redirect_stderr(captured_stderr):
            snapB = _collect_sysfx_snapshot(d["id"])
        _reemit_non_error_stderr(captured_stderr.getvalue())
        ok, info = _learn_fx_and_write_ini(
            d, fx_name, snapA, snapB,
            ini_path=ini_path, prefer_hkcu=True
        )
        if ok:
            try:
                messagebox.showinfo(
                    f"Learn FX '{fx_name}'",
                    f"Learned effect '{fx_name}' and appended to:\n{info.get('iniPath', ini_path)}\n\n"
                    f"Section: [{info.get('section')}]\n"
                    f"Value: {info.get('value_name')}\n"
                    f"Enabled={info.get('dword_enable')}, Disabled={info.get('dword_disable')}\n\n"
                    "The effect will now appear in the context menu."
                )
                self.set_status(f"Learn FX '{fx_name}': vendor INI updated")
                _log(f"GUI action: learn-fx success id={d['id']} name={d['name']} fx={fx_name} section={info.get('section')} value={info.get('value_name')}")
            except PermissionError:
                # Offer elevation or fallback
                section_name = info.get("section")
                value_name = info.get("value_name")
                dword_enable = int(info.get("dword_enable"))
                dword_disable = int(info.get("dword_disable"))
                notes = f"Learned FX '{fx_name}' for '{d['name']}' ({d['flow']})"
                if messagebox.askyesno(
                    "Permission denied",
                    f"Cannot write INI at:\n{ini_path}\n\n"
                    "Choose Yes to attempt elevation and write here.\n"
                    "Choose No to save into a user-writable location instead."
                ):
                    work = {
                        "kind": "fx",
                        "ini_path": ini_path,
                        "section": section_name,
                        "fx_name": fx_name,
                        "device_name": d["name"],
                        "value_name": value_name,
                        "dword_enable": dword_enable,
                        "dword_disable": dword_disable,
                        "flows": "Render,Capture",
                        "hives": "HKCU,HKLM",
                        "notes": notes,
                    }
                    ok2 = self._run_elevated_vendor_ini_append(work)
                    if ok2:
                        self.set_status(f"Learn FX '{fx_name}': elevated write started")
                        _log(f"GUI action: learn-fx elevation-started id={d['id']} name={d['name']} fx={fx_name} section={section_name}")
                    else:
                        self.set_status(f"Learn FX '{fx_name}': elevation failed")
                        _log(f"GUI action: learn-fx elevation-failed id={d['id']} name={d['name']} fx={fx_name} section={section_name}")
                else:
                    user_ini = _vendor_ini_default_path()
                    try:
                        _append_fx_ini_entry(
                            user_ini, section_name, fx_name, d["name"],
                            value_name, dword_enable, dword_disable,
                            flows="Render,Capture", hives="HKCU,HKLM", notes=notes
                        )
                        messagebox.showinfo("Saved to user INI", f"INI entry written to:\n{user_ini}\nSection: [{section_name}]")
                        self.set_status(f"Learn FX '{fx_name}': saved to user INI")
                        _log(f"GUI action: learn-fx success-user-ini id={d['id']} name={d['name']} fx={fx_name} section={section_name} ini={user_ini}")
                    except Exception as e2:
                        messagebox.showerror("Write failed", f"Could not write to user INI either:\n{e2}")
                        self.set_status(f"Learn FX '{fx_name}': write failed")
                        _log(f"GUI action: learn-fx user-ini-failed id={d['id']} name={d['name']} fx={fx_name} err={e2}")
        else:
            msg = str(info) if info else "Unknown failure (no details)."
            if "No suitable REG_DWORD flip" in msg:
                messagebox.showwarning(
                    f"Learn FX '{fx_name}'",
                    f"No suitable REG_DWORD flip was found for effect '{fx_name}'.\n"
                    "The driver may use a non-DWORD or a different location.\n\n"
                    f"Details: {msg}"
                )
                self.set_status(f"Learn FX '{fx_name}': no DWORD flip found")
                _log(f"GUI action: learn-fx no-dword-flip id={d['id']} name={d['name']} fx={fx_name}")
            else:
                messagebox.showerror("Error", f"FX learn failed:\n{msg}")
                self.set_status(f"Learn FX '{fx_name}': failed")
                _log(f"GUI action: learn-fx failed id={d['id']} name={d['name']} fx={fx_name} msg={msg}")
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
        initial = get_endpoint_volume(device_id)
        if initial is None:
            initial = 50
        v = tk.IntVar(value=initial)
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
    def on_toggle_fx(self, fx_name, enable):
        """Toggle an FX via high-level vendor_db function."""
        d = self.get_selected_device()
        if not d:
            return
        try:
            from .vendor_db import _apply_fx, _vendor_ini_default_path
            from .logging_setup import _log
            _log(f"GUI action: toggle-fx start id={d['id']} name={d['name']} flow={d['flow']} fx={fx_name} enable={bool(enable)}")
            ok, verified_by, state = _apply_fx(
                d["id"], d["flow"], fx_name, bool(enable),
                ini_path=_vendor_ini_default_path()
            )
            # Print the CLI equivalent for this toggle
            self.maybe_print_cli(f'audioctl enhancements --id "{d["id"]}" --flow {d["flow"]} --{"enable-fx" if enable else "disable-fx"} "{fx_name}"')
            if ok:
                state_txt = "enabled" if state else "disabled"
                self.set_status(f"{fx_name} {state_txt} for: {d['name']}")
                _log(f"GUI action: toggle-fx success id={d['id']} name={d['name']} fx={fx_name} final={state}")
            else:
                messagebox.showwarning(
                    "FX Toggle Failed",
                    f"Could not toggle effect '{fx_name}'.\n"
                    "The effect may not be properly learned for this device."
                )
                self.set_status(f"Failed to toggle {fx_name}")
                _log(f"GUI action: toggle-fx failed id={d['id']} name={d['name']} fx={fx_name}")
        except Exception as e:
            from .logging_setup import _log
            _log(f"GUI action: toggle-fx error id={d['id']} name={d['name']} fx={fx_name} err={e}")
            messagebox.showerror("Error", f"Failed to toggle {fx_name}:\n{e}")
            self.set_status(f"Error toggling {fx_name}")
    def on_toggle_fx_live(self, fx_name):
        """
        Toggle an FX at click-time by reading current state again, so the action is correct
        even if the state changed while the context menu was open.
        """
        d = self.get_selected_device()
        if not d:
            return
        try:
            from .vendor_db import _list_fx_for_device, _read_vendor_entry_state, _apply_fx, _vendor_ini_default_path
            from .logging_setup import _log
            entries = _list_fx_for_device(d["id"], d["flow"], ini_path=_vendor_ini_default_path())
            # Find the specific FX entry
            entry = None
            for e in entries:
                if str(e.get("fx_name") or "").strip().lower() == str(fx_name).strip().lower():
                    entry = e.get("entry")
                    break
            if not entry:
                messagebox.showwarning("FX Toggle", f"Effect '{fx_name}' is not currently learned for this device.")
                _log(f"GUI action: toggle-fx unavailable id={d['id']} name={d['name']} fx={fx_name}")
                return
            current = None
            try:
                current = _read_vendor_entry_state(entry, d["id"], d["flow"])
            except Exception:
                current = None
            # Decide the new desired state based on current live state
            if current is True:
                enable = False
            elif current is False:
                enable = True
            else:
                # Unknown -> choose enable as a safe default
                enable = True
            _log(f"GUI action: toggle-fx start id={d['id']} name={d['name']} fx={fx_name} current={current} target={enable}")
            ok, verified_by, state = _apply_fx(
                d["id"], d["flow"], fx_name, enable,
                ini_path=_vendor_ini_default_path()
            )
            # Print the CLI equivalent
            self.maybe_print_cli(f'audioctl enhancements --id "{d["id"]}" --flow {d["flow"]} --{"enable-fx" if enable else "disable-fx"} "{fx_name}"')
            if ok:
                state_txt = "enabled" if state else "disabled"
                self.set_status(f"{fx_name} {state_txt} for: {d['name']}")
                _log(f"GUI action: toggle-fx success id={d['id']} name={d['name']} fx={fx_name} final={state}")
            else:
                messagebox.showwarning(
                    "FX Toggle Failed",
                    f"Could not toggle effect '{fx_name}'.\n"
                    "It may have been removed or is not properly learned for this device."
                )
                self.set_status(f"Failed to toggle {fx_name}")
                _log(f"GUI action: toggle-fx failed id={d['id']} name={d['name']} fx={fx_name}")
        except Exception as e:
            from .logging_setup import _log
            messagebox.showerror("Error", f"Failed to toggle {fx_name}:\n{e}")
            self.set_status(f"Error toggling {fx_name}")
            _log(f"GUI action: toggle-fx error id={d['id']} name={d['name']} fx={fx_name} err={e}")
def launch_gui():
    try:
        # Use the library's own initializer
        comtypes.CoInitialize()
    except Exception:
        pass
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
    
    try:
        # Use the library's own uninitializer.
        # This correctly handles cleanup at process exit.
        comtypes.CoUninitialize()
    except Exception:
        pass
        
    return 0

