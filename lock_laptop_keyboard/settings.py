import json
import subprocess
import sys

try:
    import winreg
except ImportError:
    winreg = None

from .constants import (
    AUTOSTART_FLAG,
    CONTROL_MODE_INSTANT,
    DEFAULT_SETTINGS,
    RUN_KEY_PATH,
    RUN_VALUE_NAME,
    SETTINGS_FILE_NAME,
)
from .resources import app_data_dir, entry_script_path, launcher_executable


def settings_file_path():
    return app_data_dir().joinpath(SETTINGS_FILE_NAME)


def load_settings():
    settings = dict(DEFAULT_SETTINGS)
    settings_path = settings_file_path()
    if not settings_path.exists():
        return settings

    try:
        loaded = json.loads(settings_path.read_text(encoding="utf-8"))
    except (OSError, ValueError, TypeError):
        return settings

    if not isinstance(loaded, dict):
        return settings

    settings.update(
        {
            "autostart_enabled": bool(loaded.get("autostart_enabled", settings["autostart_enabled"])),
            "start_minimized_to_tray": bool(
                loaded.get("start_minimized_to_tray", settings["start_minimized_to_tray"])
            ),
            "preferred_control_mode": str(
                loaded.get("preferred_control_mode", settings["preferred_control_mode"])
            )
            or CONTROL_MODE_INSTANT,
            "instant_target_ids": [
                str(item)
                for item in loaded.get("instant_target_ids", settings["instant_target_ids"])
                if isinstance(item, str) and item
            ],
        }
    )
    return settings


def save_settings(settings):
    settings_path = settings_file_path()
    settings_path.parent.mkdir(parents=True, exist_ok=True)
    settings_path.write_text(
        json.dumps(
            {
                "autostart_enabled": bool(settings.get("autostart_enabled", False)),
                "start_minimized_to_tray": bool(settings.get("start_minimized_to_tray", True)),
                "preferred_control_mode": str(
                    settings.get("preferred_control_mode", CONTROL_MODE_INSTANT)
                )
                or CONTROL_MODE_INSTANT,
                "instant_target_ids": [
                    str(item)
                    for item in settings.get("instant_target_ids", [])
                    if isinstance(item, str) and item
                ],
            },
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
    )


def autostart_supported():
    return sys.platform == "win32" and winreg is not None


def autostart_entry():
    if not autostart_supported():
        return None

    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, RUN_KEY_PATH, 0, winreg.KEY_READ) as key:
            value, _ = winreg.QueryValueEx(key, RUN_VALUE_NAME)
            return value
    except (FileNotFoundError, OSError):
        return None


def build_autostart_command(start_minimized_to_tray):
    if getattr(sys, "frozen", False):
        parts = [sys.executable]
    else:
        parts = [launcher_executable(), str(entry_script_path())]

    if start_minimized_to_tray:
        parts.append(AUTOSTART_FLAG)

    return subprocess.list2cmdline(parts)


def set_autostart_enabled(enabled, start_minimized_to_tray):
    if not autostart_supported():
        return False, "Autostart is not supported on this platform."

    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, RUN_KEY_PATH, 0, winreg.KEY_SET_VALUE) as key:
            if enabled:
                winreg.SetValueEx(
                    key,
                    RUN_VALUE_NAME,
                    0,
                    winreg.REG_SZ,
                    build_autostart_command(start_minimized_to_tray),
                )
            else:
                try:
                    winreg.DeleteValue(key, RUN_VALUE_NAME)
                except FileNotFoundError:
                    pass
        return True, ""
    except OSError as exc:
        return False, str(exc)


def sync_settings_with_system(settings):
    synced = dict(DEFAULT_SETTINGS)
    synced.update(settings)

    entry = autostart_entry()
    if entry is None:
        synced["autostart_enabled"] = False
    else:
        synced["autostart_enabled"] = True
        synced["start_minimized_to_tray"] = AUTOSTART_FLAG in entry

    try:
        save_settings(synced)
    except OSError:
        pass
    return synced
