import queue
import tkinter as tk
from tkinter import messagebox, ttk

from .constants import (
    CONTROL_MODE_DRIVER,
    CONTROL_MODE_INSTANT,
    DEFAULT_SETTINGS,
    TRAY_COMMAND_DISABLE,
    TRAY_COMMAND_ENABLE,
    TRAY_COMMAND_EXIT,
    TRAY_COMMAND_REBOOT,
    TRAY_COMMAND_SHOW,
)
from .resources import resource_path
from .settings import (
    autostart_supported,
    save_settings,
    set_autostart_enabled,
    sync_settings_with_system,
)
from .system_control import (
    RESTART_REQUIRED,
    RESTART_USUALLY_NOT_REQUIRED,
    get_keyboard_control_context,
    is_admin,
    reboot_computer,
    reboot_computer_via_uac,
    set_keyboard_enabled,
    set_keyboard_enabled_via_uac,
)
from .tray import TrayIcon


class KeyboardControlApp(tk.Tk):
    def __init__(self, control_context, i18n, settings, launched_from_autostart=False):
        super().__init__()
        self.i18n = i18n
        self.t = i18n.t
        self.settings = dict(DEFAULT_SETTINGS)
        self.settings.update(settings or {})
        self.launched_from_autostart = launched_from_autostart
        self.control_context = dict(control_context or {})
        self.pending_driver_state = None
        self._tray_icon = None
        self._tray_queue = queue.Queue()
        self._closing = False
        self._device_vars = {}
        self._device_checkbuttons = []
        self._device_canvas_window = None

        self.status_var = tk.StringVar()
        self.subtitle_var = tk.StringVar()
        self.summary_var = tk.StringVar()
        self.mode_description_var = tk.StringVar()
        self.device_summary_var = tk.StringVar()
        self.device_caption_var = tk.StringVar()
        self.hint_var = tk.StringVar()
        self.language_var = tk.StringVar(
            value=self.t("settings.language", language=self.i18n.language_name)
        )
        self.autostart_var = tk.BooleanVar(value=bool(self.settings.get("autostart_enabled")))
        self.minimize_to_tray_var = tk.BooleanVar(
            value=bool(self.settings.get("start_minimized_to_tray"))
        )
        self.control_mode_var = tk.StringVar(value=self._resolve_initial_control_mode())

        self._configure_window()
        self._build_ui()
        self._rebuild_device_list()
        self._start_tray()
        self._update_control_state()
        self._update_settings_summary()

        self.protocol("WM_DELETE_WINDOW", self.hide_to_tray)
        self.after(150, self._poll_tray_commands)
        self.after(300, self._refresh_control_context)

        if self._active_state() is None and not self._should_start_hidden():
            self.after(250, self._show_unknown_state_warning)

        if self._should_start_hidden() and self._tray_icon is not None:
            self.after(200, self.hide_to_tray)

    def _configure_window(self):
        self.title(self.t("app.title"))
        self.geometry("900x560")
        self.minsize(820, 520)
        self.configure(bg="#eff3f8")

        try:
            self.iconbitmap(resource_path("img", "icon.ico"))
        except tk.TclError:
            pass

        style = ttk.Style(self)
        if "vista" in style.theme_names():
            style.theme_use("vista")
        elif "xpnative" in style.theme_names():
            style.theme_use("xpnative")

        style.configure("Page.TFrame", padding=16)
        style.configure("Card.TFrame", padding=14)
        style.configure("Title.TLabel", font=("Microsoft YaHei UI", 17, "bold"))
        style.configure("Subtitle.TLabel", foreground="#526070")
        style.configure("SectionTitle.TLabel", font=("Microsoft YaHei UI", 10, "bold"))
        style.configure("Body.TLabel", foreground="#223042")
        style.configure("Hint.TLabel", foreground="#5b6978")
        style.configure("StatusValue.TLabel", font=("Microsoft YaHei UI", 15, "bold"))
        style.configure("Primary.TButton", padding=(14, 9))
        style.configure("Device.TCheckbutton", font=("Microsoft YaHei UI", 10, "bold"))
        style.configure("DeviceMeta.TLabel", foreground="#506071")
        style.configure("DeviceId.TLabel", foreground="#6b7683")
        style.configure("DeviceList.Vertical.TScrollbar", arrowsize=12)

    def _build_ui(self):
        container = ttk.Frame(self, style="Page.TFrame")
        container.pack(fill=tk.BOTH, expand=True)
        container.columnconfigure(0, weight=1)
        container.rowconfigure(0, weight=1)

        notebook = ttk.Notebook(container)
        notebook.grid(row=0, column=0, sticky="nsew")

        self.control_tab = ttk.Frame(notebook, padding=16)
        self.settings_tab = ttk.Frame(notebook, padding=16)
        notebook.add(self.control_tab, text=self.t("tab.control"))
        notebook.add(self.settings_tab, text=self.t("tab.settings"))

        self._build_control_tab()
        self._build_settings_tab()

    def _build_control_tab(self):
        self.control_tab.columnconfigure(0, weight=3)
        self.control_tab.columnconfigure(1, weight=2)
        self.control_tab.rowconfigure(2, weight=1)

        header = ttk.Frame(self.control_tab, style="Card.TFrame")
        header.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 12))
        header.columnconfigure(0, weight=1)
        ttk.Label(header, text=self.t("control.title"), style="Title.TLabel").grid(
            row=0, column=0, sticky="w"
        )
        ttk.Label(header, textvariable=self.subtitle_var, style="Subtitle.TLabel").grid(
            row=1, column=0, sticky="w", pady=(6, 0)
        )

        status_frame = ttk.LabelFrame(self.control_tab, text=self.t("status.group"), padding=14)
        status_frame.grid(row=1, column=0, sticky="nsew", padx=(0, 8), pady=(0, 12))
        status_frame.columnconfigure(0, weight=1)
        ttk.Label(status_frame, text=self.t("status.label"), style="SectionTitle.TLabel").grid(
            row=0, column=0, sticky="w"
        )
        ttk.Label(status_frame, textvariable=self.status_var, style="StatusValue.TLabel").grid(
            row=1, column=0, sticky="w", pady=(8, 2)
        )
        ttk.Label(
            status_frame,
            textvariable=self.device_summary_var,
            style="Hint.TLabel",
            wraplength=420,
            justify=tk.LEFT,
        ).grid(row=2, column=0, sticky="w")

        mode_frame = ttk.LabelFrame(
            self.control_tab, text=self.t("control.mode_group"), padding=14
        )
        mode_frame.grid(row=1, column=1, sticky="nsew", pady=(0, 12))
        mode_frame.columnconfigure(0, weight=1)
        self.instant_mode_radio = ttk.Radiobutton(
            mode_frame,
            text=self.t("control.mode.instant"),
            value=CONTROL_MODE_INSTANT,
            variable=self.control_mode_var,
            command=self._on_control_mode_changed,
        )
        self.instant_mode_radio.grid(row=0, column=0, sticky="w")
        self.driver_mode_radio = ttk.Radiobutton(
            mode_frame,
            text=self.t("control.mode.driver"),
            value=CONTROL_MODE_DRIVER,
            variable=self.control_mode_var,
            command=self._on_control_mode_changed,
        )
        self.driver_mode_radio.grid(row=1, column=0, sticky="w", pady=(6, 0))
        ttk.Label(
            mode_frame,
            textvariable=self.mode_description_var,
            style="Hint.TLabel",
            wraplength=300,
            justify=tk.LEFT,
        ).grid(row=2, column=0, sticky="w", pady=(10, 0))

        device_frame = ttk.LabelFrame(
            self.control_tab, text=self.t("control.device_group"), padding=14
        )
        device_frame.grid(row=2, column=0, sticky="nsew", padx=(0, 8))
        device_frame.columnconfigure(0, weight=1)
        device_frame.rowconfigure(1, weight=1)
        ttk.Label(
            device_frame,
            textvariable=self.device_caption_var,
            style="Hint.TLabel",
            wraplength=460,
            justify=tk.LEFT,
        ).grid(row=0, column=0, sticky="w", pady=(0, 10))
        device_list_shell = ttk.Frame(device_frame)
        device_list_shell.grid(row=1, column=0, sticky="nsew")
        device_list_shell.columnconfigure(0, weight=1)
        device_list_shell.rowconfigure(0, weight=1)

        self.device_canvas = tk.Canvas(
            device_list_shell,
            highlightthickness=0,
            borderwidth=0,
            background=self.cget("bg"),
        )
        self.device_canvas.grid(row=0, column=0, sticky="nsew")

        self.device_scrollbar = ttk.Scrollbar(
            device_list_shell,
            orient=tk.VERTICAL,
            command=self.device_canvas.yview,
            style="DeviceList.Vertical.TScrollbar",
        )
        self.device_scrollbar.grid(row=0, column=1, sticky="ns", padx=(8, 0))
        self.device_canvas.configure(yscrollcommand=self.device_scrollbar.set)

        self.device_list_container = ttk.Frame(self.device_canvas)
        self.device_list_container.columnconfigure(0, weight=1)
        self._device_canvas_window = self.device_canvas.create_window(
            (0, 0),
            window=self.device_list_container,
            anchor="nw",
        )
        self.device_list_container.bind("<Configure>", self._on_device_list_configure)
        self.device_canvas.bind("<Configure>", self._on_device_canvas_configure)
        self.device_canvas.bind("<Enter>", self._bind_device_mousewheel)
        self.device_canvas.bind("<Leave>", self._unbind_device_mousewheel)

        action_frame = ttk.LabelFrame(
            self.control_tab, text=self.t("control.action_group"), padding=14
        )
        action_frame.grid(row=2, column=1, sticky="nsew")
        action_frame.columnconfigure(0, weight=1)
        action_frame.columnconfigure(1, weight=1)
        self.disable_button = ttk.Button(
            action_frame,
            text=self.t("button.disable"),
            style="Primary.TButton",
            command=lambda: self._apply_keyboard_state(False),
        )
        self.disable_button.grid(row=0, column=0, sticky="ew", padx=(0, 6), pady=(0, 10))
        self.enable_button = ttk.Button(
            action_frame,
            text=self.t("button.enable"),
            style="Primary.TButton",
            command=lambda: self._apply_keyboard_state(True),
        )
        self.enable_button.grid(row=0, column=1, sticky="ew", pady=(0, 10))
        self.refresh_button = ttk.Button(
            action_frame,
            text=self.t("button.refresh"),
            style="Primary.TButton",
            command=self._refresh_state_from_system,
        )
        self.refresh_button.grid(row=1, column=0, sticky="ew", padx=(0, 6))
        self.reboot_button = ttk.Button(
            action_frame,
            text=self.t("button.reboot"),
            style="Primary.TButton",
            command=self._request_reboot,
        )
        self.reboot_button.grid(row=1, column=1, sticky="ew")

        hint_frame = ttk.LabelFrame(self.control_tab, text=self.t("hint.group"), padding=14)
        hint_frame.grid(row=3, column=0, columnspan=2, sticky="ew", pady=(12, 0))
        hint_frame.columnconfigure(0, weight=1)
        ttk.Label(
            hint_frame,
            textvariable=self.hint_var,
            style="Body.TLabel",
            wraplength=760,
            justify=tk.LEFT,
        ).grid(row=0, column=0, sticky="nw")

    def _build_settings_tab(self):
        self.settings_tab.columnconfigure(0, weight=1)
        self.settings_tab.rowconfigure(2, weight=1)

        header = ttk.Frame(self.settings_tab, style="Card.TFrame")
        header.grid(row=0, column=0, sticky="ew", pady=(0, 14))
        header.columnconfigure(0, weight=1)
        ttk.Label(header, text=self.t("settings.title"), style="Title.TLabel").grid(
            row=0, column=0, sticky="w"
        )
        ttk.Label(header, text=self.t("settings.subtitle"), style="Subtitle.TLabel").grid(
            row=1, column=0, sticky="w", pady=(6, 0)
        )

        startup_frame = ttk.LabelFrame(
            self.settings_tab, text=self.t("settings.startup_group"), padding=18
        )
        startup_frame.grid(row=1, column=0, sticky="ew", pady=(0, 14))
        startup_frame.columnconfigure(0, weight=1)

        self.autostart_checkbox = ttk.Checkbutton(
            startup_frame,
            text=self.t("settings.autostart"),
            variable=self.autostart_var,
            command=self._refresh_startup_controls,
        )
        self.autostart_checkbox.grid(row=0, column=0, sticky="w")
        self.minimize_checkbox = ttk.Checkbutton(
            startup_frame,
            text=self.t("settings.minimize_to_tray"),
            variable=self.minimize_to_tray_var,
        )
        self.minimize_checkbox.grid(row=1, column=0, sticky="w", pady=(8, 0))
        ttk.Label(startup_frame, textvariable=self.language_var, style="Body.TLabel").grid(
            row=2, column=0, sticky="w", pady=(16, 0)
        )
        ttk.Label(
            startup_frame,
            text=self.t("settings.startup_hint"),
            wraplength=640,
            justify=tk.LEFT,
            style="Hint.TLabel",
        ).grid(row=3, column=0, sticky="w", pady=(10, 0))

        current_frame = ttk.LabelFrame(
            self.settings_tab, text=self.t("settings.current_group"), padding=18
        )
        current_frame.grid(row=2, column=0, sticky="nsew")
        current_frame.columnconfigure(0, weight=1)
        current_frame.rowconfigure(0, weight=1)
        ttk.Label(
            current_frame,
            textvariable=self.summary_var,
            wraplength=640,
            justify=tk.LEFT,
            style="Body.TLabel",
        ).grid(row=0, column=0, sticky="nw")
        self.save_button = ttk.Button(
            current_frame,
            text=self.t("settings.save"),
            style="Primary.TButton",
            command=self._save_settings,
        )
        self.save_button.grid(row=1, column=0, sticky="w", pady=(18, 0))

        self._refresh_startup_controls()

    def _resolve_initial_control_mode(self):
        preferred_mode = str(
            self.settings.get("preferred_control_mode", CONTROL_MODE_INSTANT)
        ).strip() or CONTROL_MODE_INSTANT
        instant_available = bool(self.control_context.get("instant_available"))
        if preferred_mode == CONTROL_MODE_INSTANT and instant_available:
            return CONTROL_MODE_INSTANT
        return CONTROL_MODE_DRIVER

    def _persist_settings_quietly(self):
        try:
            save_settings(self.settings)
        except OSError:
            pass

    def _current_mode(self):
        selected_mode = self.control_mode_var.get()
        if selected_mode == CONTROL_MODE_INSTANT and self.control_context.get("instant_available"):
            return CONTROL_MODE_INSTANT
        return CONTROL_MODE_DRIVER

    def _current_mode_label(self):
        return (
            self.t("control.mode.instant")
            if self._current_mode() == CONTROL_MODE_INSTANT
            else self.t("control.mode.driver")
        )

    def _available_groups(self):
        return list(self.control_context.get("instant_available_target_groups", []))

    def _selected_groups(self):
        selected = []
        if self._device_vars:
            for group in self._available_groups():
                variable = self._device_vars.get(group.get("group_key", ""))
                if variable and variable.get():
                    selected.append(group)
            return selected
        return list(self.control_context.get("instant_selected_target_groups", []))

    def _selected_target_ids(self):
        return [
            member_id
            for group in self._selected_groups()
            for member_id in group.get("member_instance_ids", [])
        ]

    def _apply_selected_targets_from_ui(self):
        selected_groups = self._selected_groups()
        selected_ids = [
            member_id
            for group in selected_groups
            for member_id in group.get("member_instance_ids", [])
        ]
        selected_lookup = {target_id.upper() for target_id in selected_ids}
        selected_devices = [
            device
            for device in self.control_context.get("instant_available_target_devices", [])
            if device.get("instance_id", "").upper() in selected_lookup
        ]
        self.control_context["instant_selected_target_groups"] = list(selected_groups)
        self.control_context["instant_target_ids"] = list(selected_ids)
        self.control_context["instant_target_devices"] = selected_devices
        self.settings["instant_target_ids"] = list(selected_ids)

    def _active_state(self):
        if self._current_mode() == CONTROL_MODE_INSTANT:
            selected_groups = self._selected_groups()
            if not selected_groups:
                return None

            group_states = []
            for group in selected_groups:
                members = list(group.get("member_devices", []))
                if not members:
                    group_states.append(None)
                    continue
                if all(self._device_is_marked_disabled(device) for device in members):
                    group_states.append(False)
                    continue
                if all(
                    device.get("present") and not self._device_is_marked_disabled(device)
                    for device in members
                ):
                    group_states.append(True)
                    continue
                group_states.append(None)

            if group_states and all(state is False for state in group_states):
                return False

            if group_states and all(state is True for state in group_states):
                return True

            return self.control_context.get("instant_enabled")

        if self.pending_driver_state is not None:
            return self.pending_driver_state

        return self.control_context.get("driver_enabled")

    def _device_is_marked_disabled(self, device):
        problem = str(device.get("problem", "") or "").upper()
        status = str(device.get("status", "") or "").upper()
        return "DISABLED" in problem or "(22)" in problem or (status == "ERROR" and bool(problem))

    def _device_kind_text(self, device):
        reason = str(device.get("reason", "") or "unknown")
        return self.t(f"control.device.kind.{reason}")

    def _device_restart_text(self, device):
        requirement = device.get("restart_requirement")
        if requirement == RESTART_REQUIRED:
            return self.t("control.device.restart.required")
        if requirement == RESTART_USUALLY_NOT_REQUIRED:
            return self.t("control.device.restart.usually_not_required")
        return self.t("control.device.restart.verify")

    def _device_meta_text(self, device):
        parts = [
            self._device_kind_text(device),
            self.t("control.device.restart.label", value=self._device_restart_text(device)),
        ]
        member_count = int(device.get("member_count", 0) or 0)
        if member_count > 1:
            parts.append(self.t("control.device.members", count=member_count))
        if not device.get("present", True):
            parts.append(self.t("control.device.not_present"))
        return " · ".join(parts)

    def _on_device_list_configure(self, _event=None):
        if hasattr(self, "device_canvas"):
            self.device_canvas.configure(scrollregion=self.device_canvas.bbox("all"))

    def _on_device_canvas_configure(self, event):
        if self._device_canvas_window is not None:
            self.device_canvas.itemconfigure(self._device_canvas_window, width=event.width)

    def _bind_device_mousewheel(self, _event=None):
        self.bind_all("<MouseWheel>", self._on_device_mousewheel)

    def _unbind_device_mousewheel(self, _event=None):
        self.unbind_all("<MouseWheel>")

    def _on_device_mousewheel(self, event):
        if not hasattr(self, "device_canvas"):
            return
        delta = event.delta
        if delta == 0:
            return
        self.device_canvas.yview_scroll(int(-delta / 120), "units")

    def _rebuild_device_list(self):
        if not hasattr(self, "device_list_container"):
            return

        for child in self.device_list_container.winfo_children():
            child.destroy()

        self._device_vars = {}
        self._device_checkbuttons = []

        available_groups = self._available_groups()
        selected_keys = {
            group.get("group_key", "")
            for group in self.control_context.get("instant_selected_target_groups", [])
        }

        if not available_groups:
            ttk.Label(
                self.device_list_container,
                text=self.t("control.device.empty"),
                style="Hint.TLabel",
                wraplength=440,
                justify=tk.LEFT,
            ).grid(row=0, column=0, sticky="w")
            return

        for row_index, group in enumerate(available_groups):
            group_key = group.get("group_key", "")
            row_frame = ttk.Frame(self.device_list_container)
            row_frame.grid(row=row_index, column=0, sticky="ew", pady=(0, 6))
            row_frame.columnconfigure(0, weight=1)

            variable = tk.BooleanVar(value=group_key in selected_keys)
            self._device_vars[group_key] = variable

            check = ttk.Checkbutton(
                row_frame,
                text=group.get("display_name") or group_key,
                variable=variable,
                command=self._on_device_selection_changed,
                style="Device.TCheckbutton",
            )
            check.grid(row=0, column=0, sticky="w")
            self._device_checkbuttons.append(check)

            ttk.Label(
                row_frame,
                text=self._device_meta_text(group),
                style="DeviceMeta.TLabel",
                wraplength=430,
                justify=tk.LEFT,
            ).grid(row=1, column=0, sticky="w", pady=(1, 0))

            member_ids = list(group.get("member_instance_ids", []))
            if len(member_ids) == 1:
                instance_text = self.t("control.device.instance", instance_id=member_ids[0])
            else:
                instance_text = self.t(
                    "control.device.instance_group",
                    count=len(member_ids),
                    instance_id=member_ids[0],
                )
            ttk.Label(
                row_frame,
                text=instance_text,
                style="DeviceId.TLabel",
                wraplength=430,
                justify=tk.LEFT,
            ).grid(row=2, column=0, sticky="w")

        self.update_idletasks()
        self._on_device_list_configure()

    def _refresh_control_context(self):
        refreshed = get_keyboard_control_context(self.settings.get("instant_target_ids"))
        self.control_context = dict(refreshed)
        self.settings["instant_target_ids"] = list(refreshed.get("instant_target_ids", []))

        if self.control_context.get("driver_enabled") == self.pending_driver_state:
            self.pending_driver_state = None

        if self.control_context.get("instant_available"):
            self.instant_mode_radio.state(["!disabled"])
        else:
            self.instant_mode_radio.state(["disabled"])
            self.control_mode_var.set(CONTROL_MODE_DRIVER)

        self.settings["preferred_control_mode"] = self._current_mode()
        self._persist_settings_quietly()
        self._rebuild_device_list()
        self._update_control_state()
        self._update_settings_summary()

    def _on_control_mode_changed(self):
        if self.control_mode_var.get() == CONTROL_MODE_INSTANT and not self.control_context.get("instant_available"):
            self.control_mode_var.set(CONTROL_MODE_DRIVER)
        self.settings["preferred_control_mode"] = self._current_mode()
        self._persist_settings_quietly()
        self._update_control_state()
        self._update_settings_summary()

    def _on_device_selection_changed(self):
        self._apply_selected_targets_from_ui()
        self._persist_settings_quietly()
        self._update_control_state()
        self._update_settings_summary()

    def _update_control_state(self):
        if self.control_context.get("instant_available"):
            self.instant_mode_radio.state(["!disabled"])
        else:
            self.instant_mode_radio.state(["disabled"])
            self.control_mode_var.set(CONTROL_MODE_DRIVER)

        self._apply_selected_targets_from_ui()

        current_mode = self._current_mode()
        effective_state = self._active_state()
        available_targets = self._available_groups()
        selected_groups = self._selected_groups()
        selected_count = len(selected_groups)
        total_count = len(available_targets)
        ignored_count = len(self.control_context.get("instant_ignored_devices", []))

        if current_mode == CONTROL_MODE_INSTANT:
            self.subtitle_var.set(self.t("control.subtitle.instant"))
            self.mode_description_var.set(self.t("control.visual_body.instant"))
            self.hint_var.set(self.t("hint.body.instant"))
            if total_count:
                self.device_caption_var.set(self.t("control.device.caption.instant"))
            else:
                self.device_caption_var.set(self.t("control.device.caption.empty"))
        else:
            self.subtitle_var.set(self.t("control.subtitle.driver"))
            self.mode_description_var.set(self.t("control.visual_body.driver"))
            self.hint_var.set(self.t("hint.body.driver"))
            if total_count:
                self.device_caption_var.set(self.t("control.device.caption.driver"))
            else:
                self.device_caption_var.set(self.t("control.device.caption.empty"))

        if total_count:
            self.device_summary_var.set(
                self.t(
                    "control.device.summary",
                    selected=selected_count,
                    total=total_count,
                    ignored=ignored_count,
                )
            )
        else:
            self.device_summary_var.set(self.t("control.mode.instant_unavailable"))

        if current_mode == CONTROL_MODE_INSTANT and total_count and selected_count == 0:
            status_text = self.t("status.no_selection")
        elif effective_state is None:
            status_text = self.t("status.unknown")
        else:
            status_text = self.t("status.enabled") if effective_state else self.t("status.disabled")

        if current_mode == CONTROL_MODE_DRIVER and self.pending_driver_state is not None:
            status_text = f"{status_text}{self.t('status.pending')}"

        self.status_var.set(status_text)

        can_operate = current_mode != CONTROL_MODE_INSTANT or selected_count > 0
        if not can_operate:
            self.disable_button.state(["disabled"])
            self.enable_button.state(["disabled"])
        elif effective_state is True:
            self.enable_button.state(["disabled"])
            self.disable_button.state(["!disabled"])
        elif effective_state is False:
            self.disable_button.state(["disabled"])
            self.enable_button.state(["!disabled"])
        else:
            self.disable_button.state(["!disabled"])
            self.enable_button.state(["!disabled"])

    def _update_settings_summary(self):
        if not autostart_supported():
            summary = self.t("settings.status.unsupported")
        elif not self.settings.get("autostart_enabled"):
            summary = self.t("settings.status.disabled")
        elif self.settings.get("start_minimized_to_tray"):
            summary = self.t("settings.status.enabled_minimized")
        else:
            summary = self.t("settings.status.enabled_visible")

        summary_lines = [summary, self.t("settings.mode.current", mode=self._current_mode_label())]
        if self._available_groups():
            summary_lines.append(
                self.t("settings.targets.current", count=len(self._selected_groups()))
            )

        self.summary_var.set("\n\n".join(summary_lines))
        self.language_var.set(self.t("settings.language", language=self.i18n.language_name))

    def _refresh_startup_controls(self):
        enabled = self.autostart_var.get()
        if autostart_supported():
            self.autostart_checkbox.state(["!disabled"])
            if enabled:
                self.minimize_checkbox.state(["!disabled"])
            else:
                self.minimize_checkbox.state(["disabled"])
        else:
            self.autostart_checkbox.state(["disabled"])
            self.minimize_checkbox.state(["disabled"])

    def _show_error(self, message):
        messagebox.showerror(self.t("dialog.error_title"), message)

    def _show_info(self, message):
        messagebox.showinfo(self.t("dialog.info_title"), message)

    def _show_warning(self, message):
        messagebox.showwarning(self.t("dialog.confirm_title"), message)

    def _show_unknown_state_warning(self):
        messagebox.showwarning(self.t("dialog.confirm_title"), self.t("warning.unknown_state"))

    def _refresh_state_from_system(self):
        self._refresh_control_context()

    def _schedule_state_refreshes(self):
        for delay_ms in (250, 900, 1800, 3200):
            self.after(delay_ms, self._refresh_state_from_system)

    def _normalize_operation_error(self, details):
        if details == "cancelled":
            return self.t("error.elevation_cancelled")
        if details == "no_internal_targets":
            return self.t("error.no_internal_targets")
        if details == "instant_requires_reboot":
            return self.t("error.instant_requires_reboot")
        if details == "instant_no_state_change":
            return self.t("error.instant_no_state_change")
        return details

    def _apply_keyboard_state(self, enabled):
        current_mode = self._current_mode()
        target_ids = self._selected_target_ids()

        if is_admin():
            success, details, resolved_target_ids = set_keyboard_enabled(
                enabled,
                control_mode=current_mode,
                target_ids=target_ids,
            )
        else:
            success, details, resolved_target_ids = set_keyboard_enabled_via_uac(
                enabled,
                control_mode=current_mode,
                target_ids=target_ids,
            )

        if not success:
            self._schedule_state_refreshes()
            normalized = self._normalize_operation_error(details)
            if current_mode == CONTROL_MODE_INSTANT and details == "instant_requires_reboot":
                self._show_warning(normalized)
                return
            error_title = (
                self.t("error.update_device")
                if current_mode == CONTROL_MODE_INSTANT
                else self.t("error.update_service")
            )
            self._show_error(f"{error_title}\n\n{normalized}")
            return

        if current_mode == CONTROL_MODE_DRIVER:
            self.pending_driver_state = enabled
        else:
            self.pending_driver_state = None

        if resolved_target_ids:
            self.settings["instant_target_ids"] = list(resolved_target_ids)
            self._persist_settings_quietly()

        self._refresh_control_context()
        self._schedule_state_refreshes()

    def _request_reboot(self):
        message = (
            self.t("dialog.reboot_pending")
            if self.pending_driver_state is not None
            else self.t("dialog.reboot_idle")
        )
        if not messagebox.askyesno(self.t("dialog.confirm_title"), message):
            return

        if is_admin():
            success, details = reboot_computer()
        else:
            success, details = reboot_computer_via_uac()

        if not success:
            normalized = self._normalize_operation_error(details)
            self._show_error(f"{self.t('error.reboot_failed')}\n\n{normalized}")

    def _save_settings(self):
        new_settings = {
            "autostart_enabled": bool(self.autostart_var.get()),
            "start_minimized_to_tray": bool(self.minimize_to_tray_var.get()),
            "preferred_control_mode": self._current_mode(),
            "instant_target_ids": list(self._selected_target_ids()),
        }

        if autostart_supported():
            success, details = set_autostart_enabled(
                new_settings["autostart_enabled"],
                new_settings["start_minimized_to_tray"],
            )
            if not success:
                self._show_error(self.t("error.save_settings", details=details))
                return

            self.settings = sync_settings_with_system(new_settings)
        else:
            try:
                save_settings(new_settings)
            except OSError as exc:
                self._show_error(self.t("error.save_settings", details=str(exc)))
                return
            self.settings.update(new_settings)

        self.settings["preferred_control_mode"] = new_settings["preferred_control_mode"]
        self.settings["instant_target_ids"] = list(new_settings["instant_target_ids"])
        self.autostart_var.set(bool(self.settings.get("autostart_enabled")))
        self.minimize_to_tray_var.set(bool(self.settings.get("start_minimized_to_tray")))
        self._refresh_startup_controls()
        self._update_settings_summary()
        self._show_info(self.t("settings.saved"))

    def _tray_labels(self):
        def compact_menu_label(text):
            return text.replace(" ", "", 1)

        return {
            "show": compact_menu_label(self.t("tray.show")),
            "disable": compact_menu_label(self.t("tray.disable")),
            "enable": compact_menu_label(self.t("tray.enable")),
            "reboot": compact_menu_label(self.t("tray.reboot")),
            "exit": compact_menu_label(self.t("tray.exit")),
        }

    def _start_tray(self):
        try:
            self._tray_icon = TrayIcon(
                tooltip=self.t("app.title"),
                command_queue=self._tray_queue,
                labels=self._tray_labels(),
                icon_path=resource_path("img", "icon.ico"),
            )
            self._tray_icon.start()
        except Exception:
            self._tray_icon = None

    def _poll_tray_commands(self):
        if self._closing:
            return

        try:
            while True:
                command = self._tray_queue.get_nowait()
                if command == TRAY_COMMAND_SHOW:
                    self.show_window()
                elif command == TRAY_COMMAND_DISABLE:
                    self._apply_keyboard_state(False)
                elif command == TRAY_COMMAND_ENABLE:
                    self._apply_keyboard_state(True)
                elif command == TRAY_COMMAND_REBOOT:
                    self._request_reboot()
                elif command == TRAY_COMMAND_EXIT:
                    self.exit_application()
        except queue.Empty:
            pass

        self.after(150, self._poll_tray_commands)

    def _should_start_hidden(self):
        return self.launched_from_autostart and bool(self.settings.get("start_minimized_to_tray"))

    def hide_to_tray(self):
        if self._tray_icon is None:
            self.exit_application()
            return
        self.withdraw()

    def show_window(self):
        self.deiconify()
        self.state("normal")
        self.after_idle(self.lift)
        self.after_idle(self.focus_force)

    def exit_application(self):
        if self._closing:
            return

        self._closing = True
        if self._tray_icon:
            self._tray_icon.stop()
            self._tray_icon = None
        self.destroy()
