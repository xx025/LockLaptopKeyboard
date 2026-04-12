import queue
import re

from PyQt5.QtCore import QTimer, Qt
from PyQt5.QtGui import QCloseEvent, QShowEvent
from PyQt5.QtWidgets import (
    QApplication,
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QMessageBox,
    QScrollArea,
    QSizePolicy,
    QVBoxLayout,
    QWidget,
)
from qfluentwidgets import (
    CardWidget,
    CheckBox,
    FluentIcon as FIF,
    FluentWindow,
    NavigationWidget,
    PrimaryPushButton,
    PushButton,
    RadioButton,
    Theme,
    setTheme,
)

from .constants import (
    CONTROL_MODE_DRIVER,
    CONTROL_MODE_INSTANT,
    DEFAULT_SETTINGS,
    THEME_DARK,
    THEME_LIGHT,
    THEME_SYSTEM,
    TRAY_COMMAND_DISABLE,
    TRAY_COMMAND_ENABLE,
    TRAY_COMMAND_EXIT,
    TRAY_COMMAND_REBOOT,
    TRAY_COMMAND_SHOW,
)
from .icons import app_icon, ensure_tray_icon_file
from .settings import (
    autostart_supported,
    normalize_theme_mode,
    resolve_theme_mode,
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


THEME_COLORS = {
    THEME_LIGHT: {
        "page_bg": "#f5f6f8",
        "nav_bg": "#fafbfc",
        "card_bg": "#ffffff",
        "card_alt_bg": "#f7f8fa",
        "tab_bg": "#eef3f8",
        "tab_selected_bg": "#e7f1fb",
        "border": "#e2e5ea",
        "text_primary": "#1c1f24",
        "text_secondary": "#5f6775",
        "text_muted": "#8690a0",
        "accent": "#0f6cbd",
        "status_enabled": "#107c10",
        "status_disabled": "#d13438",
        "status_unknown": "#a15c00",
    },
    THEME_DARK: {
        "page_bg": "#1f2228",
        "nav_bg": "#252a32",
        "card_bg": "#2b313a",
        "card_alt_bg": "#303742",
        "tab_bg": "#2d3847",
        "tab_selected_bg": "#244a71",
        "border": "#3a4350",
        "text_primary": "#eef2f7",
        "text_secondary": "#cbd3df",
        "text_muted": "#9aa6b7",
        "accent": "#5bbcff",
        "status_enabled": "#7fda71",
        "status_disabled": "#ff8f9c",
        "status_unknown": "#f0bf65",
    },
}


class KeyboardControlApp(FluentWindow):
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
        self._device_checkboxes = {}
        self._theme_colors = dict(THEME_COLORS[THEME_LIGHT])
        self._applied_theme_mode = None
        self._theme_mode = THEME_SYSTEM
        self._mode_change_guard = False
        self._theme_change_guard = False
        self._post_show_theme_synced = False

        self._build_ui()
        self._configure_window()
        self._set_control_mode(self._resolve_initial_control_mode())
        self._set_theme_mode(self._resolve_initial_theme_mode())
        self.autostart_checkbox.setChecked(bool(self.settings.get("autostart_enabled")))
        self.minimize_checkbox.setChecked(bool(self.settings.get("start_minimized_to_tray")))
        self._apply_theme()
        self._refresh_startup_controls()
        self._rebuild_device_list()
        self._update_control_state()
        self._update_settings_summary()
        self._start_tray()

        self._tray_timer = QTimer(self)
        self._tray_timer.timeout.connect(self._poll_tray_commands)
        self._tray_timer.start(150)

        self._theme_timer = QTimer(self)
        self._theme_timer.timeout.connect(self._poll_theme_changes)
        self._theme_timer.start(1200)

        QTimer.singleShot(300, self._refresh_control_context)

        if self._active_state() is None and not self._should_start_hidden():
            QTimer.singleShot(260, self._show_unknown_state_warning)
        if self._should_start_hidden() and self._tray_icon is not None:
            QTimer.singleShot(220, self.hide_to_tray)

    def mainloop(self):
        app = QApplication.instance()
        if app is None:
            raise RuntimeError("QApplication is not initialized.")
        self.show()
        return app.exec()

    def _create_card(self, name="panelCard", margin=(12, 10, 12, 10), spacing=8):
        card = CardWidget()
        card.setObjectName(name)
        layout = QVBoxLayout(card)
        layout.setContentsMargins(*margin)
        layout.setSpacing(spacing)
        return card, layout

    def _label(self, text="", name="", wrap=False):
        label = QLabel(text, self)
        if name:
            label.setObjectName(name)
        label.setWordWrap(wrap)
        label.setTextInteractionFlags(Qt.TextSelectableByMouse)
        return label

    def _clean_ui_text(self, text):
        normalized = str(text or "")
        normalized = re.sub(r"^[^\w\u4e00-\u9fffA-Za-z]+", "", normalized).strip()
        normalized = re.sub(r"\s{2,}", " ", normalized)
        return normalized

    def _ui_text(self, key, **kwargs):
        return self._clean_ui_text(self.t(key, **kwargs))

    def _build_ui(self):
        self.control_tab = QWidget(self)
        self.control_tab.setObjectName("controlTab")
        self.settings_tab = QWidget(self)
        self.settings_tab.setObjectName("settingsTab")
        self._build_control_tab()
        self._build_settings_tab()

        self.nav_top_spacer = NavigationWidget(False, self)
        self.nav_top_spacer.setObjectName("navTopSpacer")
        self.nav_top_spacer.setFixedHeight(48)
        self.navigationInterface.addWidget("app.topSpacer", self.nav_top_spacer)

        self.addSubInterface(self.control_tab, FIF.IOT, self._ui_text("tab.control"))
        self.addSubInterface(self.settings_tab, FIF.SETTING, self._ui_text("tab.settings"))
        self.switchTo(self.control_tab)

    def _sync_navigation_top_spacer(self):
        if not hasattr(self, "nav_top_spacer") or self.nav_top_spacer is None:
            return

        title_height = 0
        if hasattr(self, "systemTitleBarRect"):
            try:
                title_height = int(self.systemTitleBarRect().height())
            except Exception:
                title_height = 0

        if title_height <= 0 and hasattr(self, "titleBar"):
            try:
                title_height = int(self.titleBar.height())
            except Exception:
                title_height = 0

        self.nav_top_spacer.setFixedHeight(max(44, title_height + 12))

    def _build_control_tab(self):
        layout = QVBoxLayout(self.control_tab)
        layout.setContentsMargins(28, 22, 28, 20)
        layout.setSpacing(12)

        self.title_label = self._label(self._ui_text("control.title"), "pageTitleLabel")
        self.subtitle_label = self._label("", "pageSubtitleLabel", True)
        layout.addWidget(self.title_label)
        layout.addWidget(self.subtitle_label)

        summary_bar = QWidget(self.control_tab)
        summary_bar.setObjectName("summaryBar")
        summary_layout = QHBoxLayout(summary_bar)
        summary_layout.setContentsMargins(14, 10, 14, 10)
        summary_layout.setSpacing(18)

        status_wrap = QWidget(summary_bar)
        status_layout = QVBoxLayout(status_wrap)
        status_layout.setContentsMargins(0, 0, 0, 0)
        status_layout.setSpacing(4)
        status_layout.addWidget(self._label(self._ui_text("status.label"), "sectionLabel"))
        self.status_value_label = self._label("", "statusValueLabel")
        self.device_summary_label = self._label("", "hintLabel", True)
        status_layout.addWidget(self.status_value_label)
        status_layout.addWidget(self.device_summary_label)
        status_layout.addStretch(1)

        mode_wrap = QWidget(summary_bar)
        mode_layout = QVBoxLayout(mode_wrap)
        mode_layout.setContentsMargins(0, 0, 0, 0)
        mode_layout.setSpacing(4)
        mode_layout.addWidget(self._label(self._ui_text("control.mode_group"), "sectionLabel"))
        self.instant_mode_radio = RadioButton(self._ui_text("control.mode.instant"))
        self.driver_mode_radio = RadioButton(self._ui_text("control.mode.driver"))
        self.instant_mode_radio.toggled.connect(self._on_control_mode_changed)
        self.driver_mode_radio.toggled.connect(self._on_control_mode_changed)
        self.mode_description_label = self._label("", "hintLabel", True)
        self.mode_description_label.setMaximumHeight(24)
        mode_layout.addWidget(self.instant_mode_radio)
        mode_layout.addWidget(self.driver_mode_radio)
        mode_layout.addWidget(self.mode_description_label)
        mode_layout.addStretch(1)
        self.mode_description_label.hide()

        summary_layout.addWidget(status_wrap, stretch=3)
        summary_layout.addWidget(mode_wrap, stretch=2)
        layout.addWidget(summary_bar)

        body_layout = QHBoxLayout()
        body_layout.setContentsMargins(0, 0, 0, 0)
        body_layout.setSpacing(10)

        device_card, device_layout = self._create_card(margin=(12, 10, 12, 10), spacing=7)
        device_layout.addWidget(self._label(self._ui_text("control.device_group"), "sectionLabel"))
        self.device_caption_label = self._label("", "hintLabel", True)
        self.device_caption_label.setMaximumHeight(34)
        device_layout.addWidget(self.device_caption_label)

        self.device_scroll = QScrollArea(device_card)
        self.device_scroll.setObjectName("deviceScroll")
        self.device_scroll.setWidgetResizable(True)
        self.device_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.device_scroll.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.device_list_content = QWidget(self.device_scroll)
        self.device_list_content.setObjectName("deviceListContent")
        self.device_list_layout = QVBoxLayout(self.device_list_content)
        self.device_list_layout.setContentsMargins(2, 2, 2, 2)
        self.device_list_layout.setSpacing(6)
        self.device_scroll.setWidget(self.device_list_content)
        device_layout.addWidget(self.device_scroll, stretch=1)
        body_layout.addWidget(device_card, stretch=1)

        action_card, action_layout = self._create_card(margin=(12, 10, 12, 10), spacing=7)
        action_card.setFixedWidth(312)
        action_layout.addWidget(self._label(self._ui_text("control.action_group"), "sectionLabel"))
        grid = QGridLayout()
        grid.setHorizontalSpacing(7)
        grid.setVerticalSpacing(7)
        self.disable_button = PushButton(self._ui_text("button.disable"))
        self.enable_button = PrimaryPushButton(self._ui_text("button.enable"))
        self.refresh_button = PushButton(self._ui_text("button.refresh"))
        self.reboot_button = PushButton(self._ui_text("button.reboot"))
        self.disable_button.setObjectName("actionButton")
        self.enable_button.setObjectName("actionButton")
        self.refresh_button.setObjectName("actionButton")
        self.reboot_button.setObjectName("actionButton")
        self.disable_button.clicked.connect(lambda: self._apply_keyboard_state(False))
        self.enable_button.clicked.connect(lambda: self._apply_keyboard_state(True))
        self.refresh_button.clicked.connect(self._refresh_state_from_system)
        self.reboot_button.clicked.connect(self._request_reboot)
        grid.addWidget(self.disable_button, 0, 0)
        grid.addWidget(self.enable_button, 0, 1)
        grid.addWidget(self.refresh_button, 1, 0)
        grid.addWidget(self.reboot_button, 1, 1)
        action_layout.addLayout(grid)
        self.hint_body_label = self._label("", "hintLabel", True)
        self.hint_body_label.setMaximumHeight(24)
        action_layout.addWidget(self.hint_body_label)
        action_layout.addStretch(1)
        self.hint_body_label.hide()
        body_layout.addWidget(action_card, stretch=0)

        layout.addLayout(body_layout, stretch=1)

    def _build_settings_tab(self):
        layout = QVBoxLayout(self.settings_tab)
        layout.setContentsMargins(28, 22, 28, 20)
        layout.setSpacing(12)

        self.settings_title_label = self._label(self._ui_text("settings.title"), "pageTitleLabel")
        self.settings_subtitle_label = self._label(
            self._ui_text("settings.subtitle"),
            "pageSubtitleLabel",
            True,
        )
        layout.addWidget(self.settings_title_label)
        layout.addWidget(self.settings_subtitle_label)

        cards_layout = QHBoxLayout()
        cards_layout.setContentsMargins(0, 0, 0, 0)
        cards_layout.setSpacing(10)

        startup_card, startup_layout = self._create_card(margin=(12, 10, 12, 10), spacing=7)
        startup_layout.addWidget(self._label(self._ui_text("settings.startup_group"), "sectionLabel"))
        self.autostart_checkbox = CheckBox(self._ui_text("settings.autostart"))
        self.minimize_checkbox = CheckBox(self._ui_text("settings.minimize_to_tray"))
        self.autostart_checkbox.stateChanged.connect(self._refresh_startup_controls)
        self.language_label = self._label("", "bodyLabel")
        startup_layout.addWidget(self.autostart_checkbox)
        startup_layout.addWidget(self.minimize_checkbox)
        startup_layout.addWidget(self.language_label)
        self.startup_hint_label = self._label(self._ui_text("settings.startup_hint"), "hintLabel", True)
        startup_layout.addWidget(self.startup_hint_label)
        startup_layout.addStretch(1)
        self.startup_hint_label.hide()
        cards_layout.addWidget(startup_card, stretch=1)

        appearance_card, appearance_layout = self._create_card(margin=(12, 10, 12, 10), spacing=7)
        appearance_layout.addWidget(self._label(self._ui_text("settings.appearance_group"), "sectionLabel"))
        appearance_layout.addWidget(self._label(self._ui_text("settings.theme"), "bodyLabel"))
        self.theme_system_radio = RadioButton(self._ui_text("settings.theme.system"))
        self.theme_light_radio = RadioButton(self._ui_text("settings.theme.light"))
        self.theme_dark_radio = RadioButton(self._ui_text("settings.theme.dark"))
        self.theme_system_radio.toggled.connect(self._on_theme_mode_changed)
        self.theme_light_radio.toggled.connect(self._on_theme_mode_changed)
        self.theme_dark_radio.toggled.connect(self._on_theme_mode_changed)
        appearance_layout.addWidget(self.theme_system_radio)
        appearance_layout.addWidget(self.theme_light_radio)
        appearance_layout.addWidget(self.theme_dark_radio)
        self.appearance_hint_label = self._label(self._ui_text("settings.appearance_hint"), "hintLabel", True)
        appearance_layout.addWidget(self.appearance_hint_label)
        appearance_layout.addStretch(1)
        self.appearance_hint_label.hide()
        cards_layout.addWidget(appearance_card, stretch=1)

        layout.addLayout(cards_layout, stretch=1)

        self.save_button = PrimaryPushButton(self._ui_text("settings.save"))
        self.save_button.setObjectName("actionButton")
        self.save_button.clicked.connect(self._save_settings)
        footer_layout = QHBoxLayout()
        footer_layout.setContentsMargins(0, 0, 0, 0)
        footer_layout.addStretch(1)
        footer_layout.addWidget(self.save_button, alignment=Qt.AlignRight)
        layout.addLayout(footer_layout)

    def _switch_page(self, index):
        if not hasattr(self, "switchTo"):
            return
        self.switchTo(self.control_tab if index == 0 else self.settings_tab)

    def _update_nav_button_styles(self):
        return

    def _tray_keyboard_icon_path(self):
        return ensure_tray_icon_file()

    def _apply_keyboard_window_icon(self):
        icon = app_icon()
        if icon and not icon.isNull():
            app = QApplication.instance()
            if app is not None:
                app.setWindowIcon(icon)
            self.setWindowIcon(icon)

    def _configure_window(self):
        self.setWindowTitle(self._ui_text("app.title"))
        self.resize(1180, 800)
        self.setMinimumSize(1080, 720)

        if hasattr(self, "navigationInterface"):
            self.navigationInterface.setExpandWidth(220)
            self.navigationInterface.setMinimumExpandWidth(188)
            self.navigationInterface.setAcrylicEnabled(True)
            self.navigationInterface.setCollapsible(False)
            self.navigationInterface.setMenuButtonVisible(False)
            self.navigationInterface.setReturnButtonVisible(False)
            self.navigationInterface.setContentsMargins(0, 0, 0, 12)
            self.navigationInterface.setStyleSheet("QWidget#scrollWidget { margin-top: 0px; }")

        if hasattr(self, "setMicaEffectEnabled"):
            try:
                self.setMicaEffectEnabled(True)
            except Exception:
                pass

        self._sync_navigation_top_spacer()
        QTimer.singleShot(0, self._sync_navigation_top_spacer)
        self._apply_keyboard_window_icon()

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self._sync_navigation_top_spacer()

    def closeEvent(self, event: QCloseEvent):
        if self._closing:
            event.accept()
            return
        if self._tray_icon is None:
            self.exit_application()
            event.accept()
            return
        event.ignore()
        self.hide_to_tray()

    def _resolve_initial_control_mode(self):
        preferred_mode = str(
            self.settings.get("preferred_control_mode", CONTROL_MODE_INSTANT)
        ).strip() or CONTROL_MODE_INSTANT
        instant_available = bool(self.control_context.get("instant_available"))
        if preferred_mode == CONTROL_MODE_INSTANT and instant_available:
            return CONTROL_MODE_INSTANT
        return CONTROL_MODE_DRIVER

    def _resolve_initial_theme_mode(self):
        return normalize_theme_mode(self.settings.get("theme_mode"))

    def _set_control_mode(self, mode):
        self._mode_change_guard = True
        self.instant_mode_radio.setChecked(mode == CONTROL_MODE_INSTANT)
        self.driver_mode_radio.setChecked(mode != CONTROL_MODE_INSTANT)
        self._mode_change_guard = False

    def _set_theme_mode(self, theme_mode):
        self._theme_mode = normalize_theme_mode(theme_mode)
        self._theme_change_guard = True
        self.theme_system_radio.setChecked(self._theme_mode == THEME_SYSTEM)
        self.theme_light_radio.setChecked(self._theme_mode == THEME_LIGHT)
        self.theme_dark_radio.setChecked(self._theme_mode == THEME_DARK)
        self._theme_change_guard = False

    def _effective_theme_mode(self):
        return resolve_theme_mode(self._theme_mode)

    def _current_theme_label(self):
        selected_theme = normalize_theme_mode(self._theme_mode)
        selected_label = self.t(f"settings.theme.{selected_theme}")
        if selected_theme == THEME_SYSTEM:
            effective_label = self.t(f"settings.theme.{self._effective_theme_mode()}")
            return self.t(
                "settings.theme.current_system",
                theme=selected_label,
                effective=effective_label,
            )
        return self.t("settings.theme.current", theme=selected_label)

    def _apply_theme(self):
        effective_theme = self._effective_theme_mode()
        self._theme_colors = dict(THEME_COLORS[effective_theme])
        self._applied_theme_mode = effective_theme
        setTheme(Theme.DARK if effective_theme == THEME_DARK else Theme.LIGHT, save=False, lazy=False)
        if hasattr(self, "setMicaEffectEnabled"):
            try:
                self.setMicaEffectEnabled(effective_theme != THEME_DARK)
            except Exception:
                pass
        if hasattr(self, "navigationInterface"):
            self.navigationInterface.setAcrylicEnabled(effective_theme != THEME_DARK)
        colors = self._theme_colors

        self.setStyleSheet(
            f"""
            QWidget {{
                color: {colors["text_primary"]};
                font-family: "Segoe UI Variable Text", "Microsoft YaHei UI", "Segoe UI";
                font-size: 10.5pt;
            }}
            QWidget#controlTab, QWidget#settingsTab {{
                background: {colors["page_bg"]};
            }}
            QWidget#summaryBar {{
                background: {colors["page_bg"]};
                border: 1px solid {colors["tab_bg"]};
                border-radius: 10px;
            }}
            QFrame#panelCard {{
                background: {colors["page_bg"]};
                border: 1px solid {colors["tab_bg"]};
                border-radius: 10px;
            }}
            QLabel#pageTitleLabel {{
                font-size: 24pt;
                font-weight: 600;
                color: {colors["text_primary"]};
                padding: 0;
            }}
            QLabel#pageSubtitleLabel {{
                font-size: 11pt;
                color: {colors["text_secondary"]};
                padding: 0 0 2px 0;
            }}
            QLabel#sectionLabel {{
                font-size: 11pt;
                font-weight: 600;
                color: {colors["text_primary"]};
            }}
            QLabel#statusValueLabel {{
                font-size: 18pt;
                font-weight: 600;
            }}
            QLabel#bodyLabel {{
                color: {colors["text_primary"]};
            }}
            QLabel#hintLabel {{
                color: {colors["text_secondary"]};
                font-size: 10pt;
            }}
            QPushButton#actionButton {{
                min-height: 34px;
                border-radius: 7px;
                font-size: 10.5pt;
            }}
            QPushButton#actionButton:disabled {{
                color: {colors["text_muted"]};
            }}
            QCheckBox, QRadioButton {{
                spacing: 7px;
                min-height: 24px;
            }}
            QLabel#deviceSectionLabel {{
                color: {colors["accent"]};
                font-weight: 600;
                font-size: 10pt;
                margin-top: 4px;
            }}
            QLabel#deviceMetaLabel {{
                color: {colors["text_secondary"]};
                font-size: 10pt;
            }}
            QFrame#deviceRow {{
                background: {colors["page_bg"]};
                border: 1px solid {colors["tab_bg"]};
                border-radius: 8px;
            }}
            QFrame#deviceRow:hover {{
                background: {colors["tab_bg"]};
            }}
            QScrollArea#deviceScroll {{
                border: none;
                background: {colors["page_bg"]};
            }}
            QWidget#deviceListContent {{
                background: transparent;
            }}
            QScrollBar:vertical {{
                border: none;
                background: transparent;
                width: 10px;
                margin: 6px 2px 6px 0;
            }}
            QScrollBar::handle:vertical {{
                background: {colors["border"]};
                min-height: 24px;
                border-radius: 4px;
            }}
            QScrollBar::handle:vertical:hover {{
                background: {colors["text_muted"]};
            }}
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{
                border: none;
                background: transparent;
                height: 0px;
            }}
            """
        )
        self._apply_keyboard_window_icon()
        self._force_theme_repaint()

    def _force_theme_repaint(self):
        self.setUpdatesEnabled(False)
        try:
            self.style().unpolish(self)
            self.style().polish(self)
            for widget in self.findChildren(QWidget):
                widget.style().unpolish(widget)
                widget.style().polish(widget)
                widget.update()
        finally:
            self.setUpdatesEnabled(True)
        self.update()
        self.repaint()

    def _poll_theme_changes(self):
        if self._closing:
            return

        if normalize_theme_mode(self._theme_mode) == THEME_SYSTEM:
            effective_theme = self._effective_theme_mode()
            if effective_theme != self._applied_theme_mode:
                self._apply_theme()
                self._rebuild_device_list()
                self._update_control_state()
                self._update_settings_summary()

    def _persist_settings_quietly(self):
        try:
            save_settings(self.settings)
        except OSError:
            pass

    def _current_mode(self):
        if self.instant_mode_radio.isChecked() and self.control_context.get("instant_available"):
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
        if self._device_checkboxes:
            for group in self._available_groups():
                checkbox = self._device_checkboxes.get(group.get("group_key", ""))
                if checkbox is not None and checkbox.isChecked():
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

    def _group_section_key(self, group):
        reason = str(group.get("reason", "") or "").lower()
        if reason in {"acpi_ps2", "keyboard_controller", "converted_device"}:
            return "builtin"
        if reason in {"external_usb", "external_bluetooth", "external_hid", "hid_keyboard"}:
            return "external"
        if reason == "virtual_remote":
            return "virtual"
        return "other"

    def _group_state_key(self, group):
        members = list(group.get("member_devices", []))
        if not members:
            return "unknown"
        if all(self._device_is_marked_disabled(device) for device in members):
            return "disabled"
        if all(
            device.get("present") and not self._device_is_marked_disabled(device)
            for device in members
        ):
            return "enabled"
        return "unknown"

    def _group_state_text(self, group):
        return self.t(f"control.device.state.{self._group_state_key(group)}")

    def _device_fingerprint(self, group):
        member_ids = list(group.get("member_instance_ids", []))
        if not member_ids:
            return ""

        first_id = str(member_ids[0]).upper()
        match = re.search(r"VID_([0-9A-F]{4}).*PID_([0-9A-F]{4})", first_id)
        if match:
            return f"VID {match.group(1)} / PID {match.group(2)}"
        if "CONVERTEDDEVICE" in first_id:
            return "ConvertedDevice"
        if first_id.startswith("HID\\GVINPUT"):
            return "GVInput"
        if first_id.startswith("ACPI\\"):
            parts = first_id.split("\\")
            if len(parts) > 1:
                return parts[1]
        if first_id.startswith("HID\\"):
            parts = first_id.split("\\")
            if len(parts) > 1:
                return parts[1]
        return first_id

    def _device_primary_text(self, group):
        display_name = str(group.get("display_name", "") or "").strip()
        kind_text = self._device_kind_text(group)
        fingerprint = self._device_fingerprint(group)
        generic_names = {"", "HID KEYBOARD DEVICE", "HID KEYBOARD", "STANDARD PS/2 KEYBOARD"}

        if display_name.upper() in generic_names:
            if fingerprint:
                return f"{kind_text} ({fingerprint})"
            return kind_text
        return display_name

    def _device_restart_text(self, device):
        requirement = device.get("restart_requirement")
        if requirement == RESTART_REQUIRED:
            return self.t("control.device.restart.required")
        if requirement == RESTART_USUALLY_NOT_REQUIRED:
            return self.t("control.device.restart.usually_not_required")
        return self.t("control.device.restart.verify")

    def _device_meta_text(self, device):
        parts = [self._device_kind_text(device)]
        requirement = device.get("restart_requirement")
        if requirement in {RESTART_REQUIRED, None}:
            parts.append(self.t("control.device.restart.label", value=self._device_restart_text(device)))
        if not device.get("present", True):
            parts.append(self.t("control.device.not_present"))
        return "  |  ".join(parts)

    def _instant_mode_reboot_likely(self):
        groups = self._available_groups()
        if not groups:
            return False
        return not any(
            group.get("restart_requirement") == RESTART_USUALLY_NOT_REQUIRED for group in groups
        )

    def _clear_layout(self, layout):
        while layout.count():
            item = layout.takeAt(0)
            widget = item.widget()
            child_layout = item.layout()
            if widget is not None:
                widget.deleteLater()
            elif child_layout is not None:
                self._clear_layout(child_layout)

    def _rebuild_device_list(self):
        if not hasattr(self, "device_list_layout"):
            return

        self._clear_layout(self.device_list_layout)
        self._device_checkboxes = {}
        available_groups = self._available_groups()
        selected_keys = {
            group.get("group_key", "")
            for group in self.control_context.get("instant_selected_target_groups", [])
        }

        if not available_groups:
            self.device_list_layout.addWidget(self._label(self.t("control.device.empty"), "hintLabel", True))
            self.device_list_layout.addStretch(1)
            return

        grouped_sections = {"builtin": [], "external": [], "virtual": [], "other": []}
        for group in available_groups:
            grouped_sections[self._group_section_key(group)].append(group)

        for section_key in ("builtin", "external", "virtual", "other"):
            section_groups = grouped_sections.get(section_key, [])
            if not section_groups:
                continue

            self.device_list_layout.addWidget(
                self._label(self.t(f"control.device.section.{section_key}"), "deviceSectionLabel")
            )

            for group in section_groups:
                group_key = group.get("group_key", "")
                state_key = self._group_state_key(group)

                row = QFrame(self.device_list_content)
                row.setObjectName("deviceRow")
                row_layout = QVBoxLayout(row)
                row_layout.setContentsMargins(10, 8, 10, 8)
                row_layout.setSpacing(4)

                top = QHBoxLayout()
                top.setContentsMargins(0, 0, 0, 0)
                top.setSpacing(8)

                checkbox = CheckBox(self._device_primary_text(group), row)
                checkbox.setChecked(group_key in selected_keys)
                checkbox.stateChanged.connect(self._on_device_selection_changed)
                self._device_checkboxes[group_key] = checkbox
                top.addWidget(checkbox, stretch=1)

                state_label = QLabel(self._group_state_text(group), row)
                if state_key == "disabled":
                    state_color = self._theme_colors["status_disabled"]
                    checkbox.setStyleSheet(f"color: {state_color}; font-weight: 600;")
                    state_label.setStyleSheet(f"color: {state_color}; font-weight: 600;")
                elif state_key == "enabled":
                    state_color = self._theme_colors["status_enabled"]
                    state_label.setStyleSheet(f"color: {state_color}; font-weight: 600;")
                else:
                    state_color = self._theme_colors["status_unknown"]
                    state_label.setStyleSheet(f"color: {state_color}; font-weight: 600;")

                top.addWidget(state_label, alignment=Qt.AlignRight)
                row_layout.addLayout(top)

                meta = QLabel(self._device_meta_text(group), row)
                meta.setObjectName("deviceMetaLabel")
                meta.setWordWrap(True)
                row_layout.addWidget(meta)
                self.device_list_layout.addWidget(row)

        self.device_list_layout.addStretch(1)

    def _refresh_control_context(self):
        refreshed = get_keyboard_control_context(self.settings.get("instant_target_ids"))
        self.control_context = dict(refreshed)
        self.settings["instant_target_ids"] = list(refreshed.get("instant_target_ids", []))

        if self.control_context.get("driver_enabled") == self.pending_driver_state:
            self.pending_driver_state = None

        if not self.control_context.get("instant_available"):
            self._set_control_mode(CONTROL_MODE_DRIVER)

        self.settings["preferred_control_mode"] = self._current_mode()
        self._persist_settings_quietly()
        self._rebuild_device_list()
        self._update_control_state()
        self._update_settings_summary()

    def _on_control_mode_changed(self, _checked=None):
        if _checked is False:
            return
        if self._mode_change_guard:
            return
        if self.instant_mode_radio.isChecked() and not self.control_context.get("instant_available"):
            self._set_control_mode(CONTROL_MODE_DRIVER)

        self.settings["preferred_control_mode"] = self._current_mode()
        self._persist_settings_quietly()
        self._update_control_state()
        self._update_settings_summary()

    def _on_theme_mode_changed(self, _checked=None):
        if _checked is False:
            return
        if self._theme_change_guard:
            return

        if self.theme_system_radio.isChecked():
            self._theme_mode = THEME_SYSTEM
        elif self.theme_dark_radio.isChecked():
            self._theme_mode = THEME_DARK
        else:
            self._theme_mode = THEME_LIGHT

        self.settings["theme_mode"] = self._theme_mode
        self._apply_theme()
        self._rebuild_device_list()
        self._update_control_state()
        self._update_settings_summary()

    def showEvent(self, event: QShowEvent):
        super().showEvent(event)
        if self._post_show_theme_synced:
            return
        self._post_show_theme_synced = True
        QTimer.singleShot(0, self._apply_theme)

    def _on_device_selection_changed(self, _state=None):
        self._apply_selected_targets_from_ui()
        self._persist_settings_quietly()
        self._update_control_state()
        self._update_settings_summary()

    def _update_control_state(self):
        instant_available = bool(self.control_context.get("instant_available"))
        self.instant_mode_radio.setEnabled(instant_available)
        if not instant_available and self.instant_mode_radio.isChecked():
            self._set_control_mode(CONTROL_MODE_DRIVER)

        self._apply_selected_targets_from_ui()

        current_mode = self._current_mode()
        effective_state = self._active_state()
        available_targets = self._available_groups()
        selected_groups = self._selected_groups()
        selected_count = len(selected_groups)
        total_count = len(available_targets)
        ignored_count = len(self.control_context.get("instant_ignored_devices", []))
        reboot_likely = self._instant_mode_reboot_likely()

        if current_mode == CONTROL_MODE_INSTANT:
            if reboot_likely:
                self.subtitle_label.setText(self._ui_text("control.subtitle.instant_limited"))
            else:
                self.subtitle_label.setText(self._ui_text("control.subtitle.instant"))

            self.mode_description_label.setText("")
            if total_count:
                if reboot_likely:
                    self.device_caption_label.setText(self._ui_text("control.device.caption.instant_limited"))
                else:
                    self.device_caption_label.setText(self._ui_text("control.device.caption.instant"))
            else:
                self.device_caption_label.setText(self._ui_text("control.device.caption.empty"))
            self.hint_body_label.setText("")
        else:
            self.subtitle_label.setText(self._ui_text("control.subtitle.driver"))
            self.mode_description_label.setText("")
            if total_count:
                self.device_caption_label.setText(self._ui_text("control.device.caption.driver"))
            else:
                self.device_caption_label.setText(self._ui_text("control.device.caption.empty"))
            self.hint_body_label.setText("")

        if total_count:
            self.device_summary_label.setText(
                self._ui_text(
                    "control.device.summary",
                    selected=selected_count,
                    total=total_count,
                    ignored=ignored_count,
                )
            )
        else:
            self.device_summary_label.setText(self._ui_text("control.mode.instant_unavailable"))

        if current_mode == CONTROL_MODE_INSTANT and total_count and selected_count == 0:
            status_text = self._ui_text("status.no_selection")
            status_color = self._theme_colors["text_muted"]
        elif effective_state is None:
            status_text = self._ui_text("status.unknown")
            status_color = self._theme_colors["status_unknown"]
        elif effective_state:
            status_text = self._ui_text("status.enabled")
            status_color = self._theme_colors["status_enabled"]
        else:
            status_text = self._ui_text("status.disabled")
            status_color = self._theme_colors["status_disabled"]

        if current_mode == CONTROL_MODE_DRIVER and self.pending_driver_state is not None:
            status_text = f"{status_text}{self._ui_text('status.pending')}"

        self.status_value_label.setText(status_text)
        self.status_value_label.setStyleSheet(f"color: {status_color};")

        can_operate = current_mode != CONTROL_MODE_INSTANT or selected_count > 0
        if not can_operate:
            self.disable_button.setEnabled(False)
            self.enable_button.setEnabled(False)
        elif effective_state is True:
            self.disable_button.setEnabled(True)
            self.enable_button.setEnabled(False)
        elif effective_state is False:
            self.disable_button.setEnabled(False)
            self.enable_button.setEnabled(True)
        else:
            self.disable_button.setEnabled(True)
            self.enable_button.setEnabled(True)

    def _update_settings_summary(self):
        self.language_label.setText(
            self._ui_text("settings.language", language=self.i18n.language_name)
        )

    def _refresh_startup_controls(self, _state=None):
        enabled = self.autostart_checkbox.isChecked()
        if autostart_supported():
            self.autostart_checkbox.setEnabled(True)
            self.minimize_checkbox.setEnabled(enabled)
        else:
            self.autostart_checkbox.setEnabled(False)
            self.minimize_checkbox.setEnabled(False)

    def _show_error(self, message):
        QMessageBox.critical(self, self.t("dialog.error_title"), message)

    def _show_info(self, message):
        QMessageBox.information(self, self.t("dialog.info_title"), message)

    def _show_warning(self, message):
        QMessageBox.warning(self, self.t("dialog.confirm_title"), message)

    def _show_unknown_state_warning(self):
        QMessageBox.warning(
            self,
            self.t("dialog.confirm_title"),
            self.t("warning.unknown_state"),
        )

    def _refresh_state_from_system(self, _checked=None):
        self._refresh_control_context()

    def _schedule_state_refreshes(self):
        for delay_ms in (250, 900, 1800, 3200):
            QTimer.singleShot(delay_ms, self._refresh_state_from_system)

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

    def _request_reboot(self, _checked=None):
        message = (
            self.t("dialog.reboot_pending")
            if self.pending_driver_state is not None
            else self.t("dialog.reboot_idle")
        )
        result = QMessageBox.question(
            self,
            self.t("dialog.confirm_title"),
            message,
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )
        if result != QMessageBox.Yes:
            return

        if is_admin():
            success, details = reboot_computer()
        else:
            success, details = reboot_computer_via_uac()

        if not success:
            normalized = self._normalize_operation_error(details)
            self._show_error(f"{self.t('error.reboot_failed')}\n\n{normalized}")

    def _save_settings(self, _checked=None):
        new_settings = {
            "autostart_enabled": bool(self.autostart_checkbox.isChecked()),
            "start_minimized_to_tray": bool(self.minimize_checkbox.isChecked()),
            "preferred_control_mode": self._current_mode(),
            "theme_mode": normalize_theme_mode(self._theme_mode),
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
        self.settings["theme_mode"] = new_settings["theme_mode"]
        self.settings["instant_target_ids"] = list(new_settings["instant_target_ids"])
        self.autostart_checkbox.setChecked(bool(self.settings.get("autostart_enabled")))
        self.minimize_checkbox.setChecked(bool(self.settings.get("start_minimized_to_tray")))
        self._set_theme_mode(self.settings.get("theme_mode"))
        self._apply_theme()
        self._rebuild_device_list()
        self._update_control_state()
        self._refresh_startup_controls()
        self._update_settings_summary()
        self._show_info(self.t("settings.saved"))

    def _tray_labels(self):
        def compact_menu_label(text):
            return text.replace(" ", "", 1)

        return {
            "show": compact_menu_label(self._ui_text("tray.show")),
            "disable": compact_menu_label(self._ui_text("tray.disable")),
            "enable": compact_menu_label(self._ui_text("tray.enable")),
            "reboot": compact_menu_label(self._ui_text("tray.reboot")),
            "exit": compact_menu_label(self._ui_text("tray.exit")),
        }

    def _start_tray(self):
        try:
            self._tray_icon = TrayIcon(
                tooltip=self._ui_text("app.title"),
                command_queue=self._tray_queue,
                labels=self._tray_labels(),
                icon_path=self._tray_keyboard_icon_path(),
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

    def _should_start_hidden(self):
        return self.launched_from_autostart and bool(self.settings.get("start_minimized_to_tray"))

    def hide_to_tray(self):
        if self._tray_icon is None:
            self.exit_application()
            return
        self.hide()

    def show_window(self):
        self.show()
        state = self.windowState()
        self.setWindowState((state & ~Qt.WindowMinimized) | Qt.WindowActive)
        self.raise_()
        self.activateWindow()

    def exit_application(self):
        if self._closing:
            return

        self._closing = True
        if hasattr(self, "_tray_timer"):
            self._tray_timer.stop()
        if hasattr(self, "_theme_timer"):
            self._theme_timer.stop()
        if self._tray_icon:
            self._tray_icon.stop()
            self._tray_icon = None

        app = QApplication.instance()
        if app is not None:
            app.quit()
