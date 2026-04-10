import ctypes
import locale
import sys


MESSAGES = {
    "zh-CN": {
        "app.title": "⌨ 键盘控制工具",
        "app.fatal_error": "程序启动失败：\n\n{details}",
        "tab.control": "🏠 控制",
        "tab.settings": "⚙️ 设置",
        "control.title": "⌨ 键盘控制",
        "control.subtitle": "当前为驱动模式，修改后需要重启电脑生效。",
        "control.subtitle.instant": "当前为即时设备模式，成功时通常无需重启即可生效。",
        "control.subtitle.instant_limited": "当前为即时设备模式，但这台机器上检测到的键盘大多仍需要重启后才能真正生效。",
        "control.subtitle.driver": "当前为驱动模式，修改后需要重启电脑生效。",
        "control.action_group": "⚡ 快捷操作",
        "control.mode_group": "🧭 控制模式",
        "control.mode.instant": "⚡ 即时设备模式",
        "control.mode.driver": "🧱 驱动兼容模式",
        "control.device_group": "🗂 设备选择",
        "control.device.caption.instant": "勾选即时模式要操作的键盘项。一个物理键盘可能会合并多个底层子设备，默认只预选更像内置键盘的目标。",
        "control.device.caption.instant_limited": "这台机器上的键盘在即时模式下大多仍会要求重启。列表已按内置、外置和其它输入分组，先勾选目标，再决定是否继续操作。",
        "control.device.caption.driver": "当前是驱动模式，设备勾选暂不会生效；切回即时模式时会使用这里的选择。",
        "control.device.caption.empty": "当前没有可选的键盘设备。可以点击“刷新设备”，或切换到驱动模式。",
        "control.device.summary": "已选 {selected} / {total} 个键盘项，已忽略 {ignored} 个虚拟或不支持的目标。",
        "control.device.empty": "未检测到可选择的键盘设备。",
        "control.device.instance": "ID: {instance_id}",
        "control.device.instance_group": "包含 {count} 个子设备，例如：{instance_id}",
        "control.device.members": "{count} 个子设备",
        "control.device.selection.recommended": "默认推荐",
        "control.device.state.enabled": "当前启用",
        "control.device.state.disabled": "当前禁用",
        "control.device.state.unknown": "状态待确认",
        "control.device.section.builtin": "推荐的内置目标",
        "control.device.section.external": "外置键盘",
        "control.device.section.virtual": "虚拟 / 远程输入",
        "control.device.section.other": "其它输入目标",
        "control.device.kind.acpi_ps2": "内置 PS/2 / ACPI",
        "control.device.kind.keyboard_controller": "内置控制器",
        "control.device.kind.converted_device": "系统转换设备",
        "control.device.kind.external_usb": "外置 USB 键盘",
        "control.device.kind.external_bluetooth": "外置蓝牙键盘",
        "control.device.kind.external_hid": "外置 HID 键盘",
        "control.device.kind.hid_keyboard": "HID 键盘",
        "control.device.kind.virtual_remote": "虚拟 / 远程键盘",
        "control.device.kind.unknown": "其它键盘设备",
        "control.device.restart.label": "重启：{value}",
        "control.device.restart.required": "需要",
        "control.device.restart.usually_not_required": "通常不需要",
        "control.device.restart.verify": "待验证",
        "control.device.not_present": "当前未连接",
        "control.mode.instant_summary": "检测到 {count} 个可直接控制的键盘设备。",
        "control.mode.instant_unavailable": "暂未检测到可安全操作的键盘设备，将回退到驱动模式。",
        "control.mode.driver_summary": "通过修改 i8042prt 服务启动类型控制键盘，兼容性更广，但需要重启。",
        "control.visual_title": "🧩 当前模式",
        "control.visual_body": "当前仍是驱动模式。禁用或启用操作会修改 i8042prt 启动类型，并在下次重启后生效。",
        "control.visual_body.instant": (
            "即时模式会定位当前可控的键盘设备，"
            "并按下方勾选目标直接禁用或启用。是否需要重启请以设备列表中的提示为准。"
        ),
        "control.visual_body.instant_limited": (
            "这台机器上检测到的键盘在即时模式下大多仍会返回“需要重启”。"
            "即时模式更适合用来精确选择设备目标，而不是保证立刻切换。"
        ),
        "control.visual_body.driver": (
            "驱动模式会修改 i8042prt 启动类型，兼容旧机器或无法识别设备级目标的情况，"
            "但需要重启后才会真正生效。"
        ),
        "status.group": "📌 状态",
        "status.label": "📌 当前状态：",
        "status.enabled": "✅ 已启用",
        "status.disabled": "🚫 已禁用",
        "status.unknown": "❓ 状态未知（按已启用显示）",
        "status.no_selection": "⚪ 未选择设备",
        "status.pending": "（♻️ 已修改，重启后生效）",
        "status.fallback_enabled": "✅ 键盘已启用",
        "status.fallback_disabled": "🚫 键盘已禁用",
        "hint.group": "ℹ️ 说明",
        "hint.body": (
            "所有操作都需要重启电脑后才会生效。\n"
            "本工具默认使用 demand 方式重新启用键盘。\n\n"
            "本工具通过修改 i8042prt 服务启动类型来控制笔记本内置键盘，"
            "可能会影响部分 PS/2 设备，请自行评估风险。"
        ),
        "hint.body.instant": (
            "即时模式会展示内置与外置键盘，并继续忽略明显的虚拟输入设备。\n"
            "默认只会预选更像内置键盘的目标；如果你也想控制外置键盘，可以手动勾选。"
        ),
        "hint.body.instant_limited": (
            "从当前机器的返回结果看，内置和外置键盘大多都仍需要重启后才会真正切换。\n"
            "列表已经按内置、外置和虚拟输入分组，并补充了当前状态、重启提示和子设备数量，方便你更准确地选择目标。"
        ),
        "hint.body.driver": (
            "驱动模式适合作为兼容兜底方案。\n"
            "它通过修改 i8042prt 服务启动类型控制笔记本内置键盘，"
            "可能会影响部分 PS/2 设备，并且需要重启后生效。"
        ),
        "button.reboot": "🔄 重启电脑",
        "button.disable": "🚫 禁用键盘",
        "button.enable": "✅ 启用键盘",
        "button.refresh": "🔁 刷新设备",
        "warning.unknown_state": (
            "⚠️ 未能读取当前键盘状态，将按“已启用”显示。"
            "如果状态不一致，请手动确认后再操作。"
        ),
        "settings.title": "⚙️ 启动设置",
        "settings.subtitle": "这里可以控制程序在 Windows 启动后的行为。",
        "settings.startup_group": "🚀 开机启动",
        "settings.autostart": "开机自启",
        "settings.minimize_to_tray": "开机时最小化到托盘",
        "settings.language": "🌐 界面语言：跟随系统（{language}）",
        "settings.startup_hint": (
            "开机自启写入当前用户的 Windows 启动项。\n"
            "如果勾选“最小化到托盘”，系统启动后程序不会自动弹出主窗口。"
        ),
        "settings.appearance_group": "🎨 外观",
        "settings.theme": "主题",
        "settings.theme.system": "跟随系统",
        "settings.theme.light": "浅色",
        "settings.theme.dark": "深色",
        "settings.appearance_hint": "切换主题会立即预览，保存设置后下次启动继续使用。",
        "settings.current_group": "🧭 当前设置",
        "settings.status.unsupported": "当前平台不支持开机自启设置。",
        "settings.status.disabled": "当前未开启开机自启。",
        "settings.status.enabled_visible": "已开启开机自启，系统启动后会正常显示主窗口。",
        "settings.status.enabled_minimized": "已开启开机自启，系统启动后将直接最小化到托盘。",
        "settings.save": "💾 保存设置",
        "settings.saved": "设置已保存。",
        "settings.targets.current": "当前已选择 {count} 个即时模式键盘项。",
        "settings.theme.current": "当前主题：{theme}",
        "settings.theme.current_system": "当前主题：{theme}（正在使用{effective}）",
        "tray.show": "🪟 显示窗口",
        "tray.disable": "🚫 禁用键盘",
        "tray.enable": "✅ 启用键盘",
        "tray.reboot": "🔄 重启电脑",
        "tray.exit": "❌ 退出",
        "dialog.error_title": "错误",
        "dialog.info_title": "提示",
        "dialog.confirm_title": "确认",
        "dialog.reboot_pending": "修改需要重启电脑后才会生效，是否现在重启？",
        "dialog.reboot_idle": "当前没有新的修改，是否仍然立即重启电脑？",
        "error.update_service": "未能更新 i8042prt 服务启动类型。",
        "error.update_device": "未能更新键盘设备状态。",
        "error.no_internal_targets": "未检测到或未选择可操作的键盘设备。",
        "error.instant_requires_reboot": "系统已经接受了这次设备变更，但目标键盘当前仍未立即切换，通常需要重启后才会真正生效。",
        "error.instant_no_state_change": "设备命令已经执行，但没有检测到状态变化。这个键盘可能不支持即时切换，或系统仍在处理中。",
        "error.save_settings": "未能保存设置：\n\n{details}",
        "error.reboot_failed": "未能发起重启。",
        "error.elevation_cancelled": "管理员授权已取消。",
        "settings.mode.current": "当前默认控制模式：{mode}",
        "language.name": "简体中文",
    },
    "en-US": {
        "app.title": "⌨ Keyboard Control Utility",
        "app.fatal_error": "The application failed to start:\n\n{details}",
        "tab.control": "🏠 Control",
        "tab.settings": "⚙️ Settings",
        "control.title": "⌨ Keyboard Control",
        "control.subtitle": "Driver mode is active. Changes take effect after a reboot.",
        "control.subtitle.instant": "Instant device mode is active. Successful changes usually apply without a reboot.",
        "control.subtitle.instant_limited": "Instant device mode is active, but most detected keyboards on this machine still seem to require a reboot before the change really applies.",
        "control.subtitle.driver": "Driver mode is active. Changes take effect after a reboot.",
        "control.action_group": "⚡ Quick Actions",
        "control.mode_group": "🧭 Control Mode",
        "control.mode.instant": "⚡ Instant device mode",
        "control.mode.driver": "🧱 Driver compatibility mode",
        "control.device_group": "🗂 Device Selection",
        "control.device.caption.instant": "Select which keyboard entries instant mode should control. One physical keyboard may combine multiple low-level child devices, and only built-in-looking targets are preselected by default.",
        "control.device.caption.instant_limited": "Most keyboards on this machine still appear to require a reboot even in instant mode. The list is grouped into built-in, external, and other input targets so it is easier to pick the right device first.",
        "control.device.caption.driver": "Driver mode ignores device selection for now. These selections will be used again when you switch back to instant mode.",
        "control.device.caption.empty": "No selectable keyboard device is available right now. Try Refresh devices or switch to driver mode.",
        "control.device.summary": "{selected} of {total} keyboard entries selected. {ignored} virtual or unsupported targets are ignored.",
        "control.device.empty": "No selectable keyboard device was detected.",
        "control.device.instance": "ID: {instance_id}",
        "control.device.instance_group": "Contains {count} child devices, for example: {instance_id}",
        "control.device.members": "{count} child devices",
        "control.device.selection.recommended": "Recommended by default",
        "control.device.state.enabled": "Currently enabled",
        "control.device.state.disabled": "Currently disabled",
        "control.device.state.unknown": "State needs verification",
        "control.device.section.builtin": "Recommended built-in targets",
        "control.device.section.external": "External keyboards",
        "control.device.section.virtual": "Virtual / remote input",
        "control.device.section.other": "Other input targets",
        "control.device.kind.acpi_ps2": "Built-in PS/2 / ACPI",
        "control.device.kind.keyboard_controller": "Built-in controller",
        "control.device.kind.converted_device": "Converted system device",
        "control.device.kind.external_usb": "External USB keyboard",
        "control.device.kind.external_bluetooth": "External Bluetooth keyboard",
        "control.device.kind.external_hid": "External HID keyboard",
        "control.device.kind.hid_keyboard": "HID keyboard",
        "control.device.kind.virtual_remote": "Virtual / remote keyboard",
        "control.device.kind.unknown": "Other keyboard device",
        "control.device.restart.label": "Reboot: {value}",
        "control.device.restart.required": "Required",
        "control.device.restart.usually_not_required": "Usually not needed",
        "control.device.restart.verify": "Needs verification",
        "control.device.not_present": "Currently disconnected",
        "control.mode.instant_summary": "Detected {count} keyboard devices that can be controlled directly.",
        "control.mode.instant_unavailable": "No safe keyboard device target was detected, so the app will fall back to driver mode.",
        "control.mode.driver_summary": "Controls the keyboard by changing the i8042prt start type. More compatible, but requires a reboot.",
        "control.visual_title": "🧩 Current Mode",
        "control.visual_body": (
            "Driver mode is still active. Disabling or enabling updates the i8042prt start type "
            "and takes effect after the next reboot."
        ),
        "control.visual_body.instant": (
            "Instant mode tries to locate controllable keyboard devices "
            "and apply changes to the selected targets directly. Check the device list for the reboot expectation of each target."
        ),
        "control.visual_body.instant_limited": (
            "On this machine, most detected keyboards still report that a reboot is required even in instant mode. "
            "Instant mode is therefore more useful for precise target selection than for guaranteeing an immediate switch."
        ),
        "control.visual_body.driver": (
            "Driver mode changes the i8042prt start type. It is a wider compatibility fallback, "
            "but the change only applies after the next reboot."
        ),
        "status.group": "📌 Status",
        "status.label": "📌 Current status:",
        "status.enabled": "✅ Enabled",
        "status.disabled": "🚫 Disabled",
        "status.unknown": "❓ Unknown (shown as enabled)",
        "status.no_selection": "⚪ No device selected",
        "status.pending": " (♻️ Changed, pending reboot)",
        "status.fallback_enabled": "✅ Keyboard enabled",
        "status.fallback_disabled": "🚫 Keyboard disabled",
        "hint.group": "ℹ️ Notes",
        "hint.body": (
            "All changes require a reboot to take effect.\n"
            "This app uses the demand start type to re-enable the keyboard.\n\n"
            "The app controls the built-in keyboard by changing the i8042prt service start type. "
            "This may affect some PS/2 devices, so please review the risk before use."
        ),
        "hint.body.instant": (
            "Instant mode lists both built-in and external keyboards while still excluding obvious virtual input devices.\n"
            "Only built-in-looking targets are preselected by default. Select external keyboards manually if you also want to control them."
        ),
        "hint.body.instant_limited": (
            "Based on the responses from this machine, both built-in and external keyboards usually still need a reboot before the change really applies.\n"
            "The list is now grouped by built-in, external, and virtual input, and each entry shows the current state, reboot expectation, and child-device count."
        ),
        "hint.body.driver": (
            "Driver mode is kept as the compatibility fallback.\n"
            "It changes the i8042prt start type, which may affect some PS/2 devices and requires a reboot."
        ),
        "button.reboot": "🔄 Reboot computer",
        "button.disable": "🚫 Disable keyboard",
        "button.enable": "✅ Enable keyboard",
        "button.refresh": "🔁 Refresh devices",
        "warning.unknown_state": (
            "⚠️ The current keyboard state could not be read. "
            "It will be shown as enabled. Please verify before making changes."
        ),
        "settings.title": "⚙️ Startup Settings",
        "settings.subtitle": "Control how the app behaves after Windows starts.",
        "settings.startup_group": "🚀 Startup",
        "settings.autostart": "Start with Windows",
        "settings.minimize_to_tray": "Start minimized to tray",
        "settings.language": "🌐 Interface language follows the system ({language})",
        "settings.startup_hint": (
            "Autostart uses the current user's Windows Run entry.\n"
            "If 'Start minimized to tray' is enabled, the main window stays hidden on startup."
        ),
        "settings.appearance_group": "🎨 Appearance",
        "settings.theme": "Theme",
        "settings.theme.system": "Follow system",
        "settings.theme.light": "Light",
        "settings.theme.dark": "Dark",
        "settings.appearance_hint": "Theme changes preview immediately and will be kept after you save settings.",
        "settings.current_group": "🧭 Current Settings",
        "settings.status.unsupported": "Autostart settings are not supported on this platform.",
        "settings.status.disabled": "Autostart is currently disabled.",
        "settings.status.enabled_visible": "Autostart is enabled and the main window will be shown on startup.",
        "settings.status.enabled_minimized": "Autostart is enabled and the app will start minimized to the tray.",
        "settings.save": "💾 Save settings",
        "settings.saved": "Settings saved.",
        "settings.targets.current": "{count} instant-mode keyboard entries are currently selected.",
        "settings.theme.current": "Current theme: {theme}",
        "settings.theme.current_system": "Current theme: {theme} (using {effective})",
        "tray.show": "🪟 Show window",
        "tray.disable": "🚫 Disable keyboard",
        "tray.enable": "✅ Enable keyboard",
        "tray.reboot": "🔄 Reboot computer",
        "tray.exit": "❌ Exit",
        "dialog.error_title": "Error",
        "dialog.info_title": "Info",
        "dialog.confirm_title": "Confirm",
        "dialog.reboot_pending": "The change only takes effect after a reboot. Reboot now?",
        "dialog.reboot_idle": "No new change is pending. Reboot now anyway?",
        "error.update_service": "Could not update the i8042prt start type.",
        "error.update_device": "Could not update the keyboard device state.",
        "error.no_internal_targets": "No controllable keyboard device was detected or selected.",
        "error.instant_requires_reboot": (
            "Windows accepted the device change, but the target keyboard did not switch immediately. "
            "This device usually still requires a reboot."
        ),
        "error.instant_no_state_change": (
            "The device command completed, but no state change was detected. "
            "This keyboard may not support instant switching, or Windows is still processing it."
        ),
        "error.save_settings": "Could not save settings:\n\n{details}",
        "error.reboot_failed": "Could not start the reboot.",
        "error.elevation_cancelled": "Administrator permission was cancelled.",
        "settings.mode.current": "Current default control mode: {mode}",
        "language.name": "English",
    },
}


def normalize_language_tag(tag):
    normalized = (tag or "").replace("_", "-").lower()
    if normalized.startswith("zh"):
        return "zh-CN"
    return "en-US"


def detect_system_language():
    language_tag = ""

    if sys.platform == "win32":
        try:
            language_id = ctypes.windll.kernel32.GetUserDefaultUILanguage()
            language_tag = locale.windows_locale.get(language_id, "")
        except (AttributeError, OSError):
            language_tag = ""

    if not language_tag:
        for getter in (locale.getlocale, locale.getdefaultlocale):
            try:
                result = getter()
            except (ValueError, TypeError, AttributeError):
                continue

            if isinstance(result, tuple) and result and result[0]:
                language_tag = result[0]
                break

    return normalize_language_tag(language_tag)


class I18n:
    def __init__(self, language=None):
        self.language = normalize_language_tag(language or detect_system_language())

    @property
    def language_name(self):
        return self.t("language.name")

    def t(self, key, **kwargs):
        language_messages = MESSAGES.get(self.language, MESSAGES["en-US"])
        template = language_messages.get(key, MESSAGES["en-US"].get(key, key))
        return template.format(**kwargs)


def create_i18n(language=None):
    return I18n(language=language)
