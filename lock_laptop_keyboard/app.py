import argparse
import json
import sys
import traceback
import tkinter as tk
from tkinter import messagebox
from pathlib import Path

from .constants import (
    AUTOSTART_FLAG,
    CONTROL_MODE_DRIVER,
    CONTROL_MODE_FLAG,
    CONTROL_MODE_INSTANT,
    REBOOT_NOW_FLAG,
    RESULT_FILE_FLAG,
    SET_STATE_FLAG,
    TARGET_ID_FLAG,
)
from .i18n import create_i18n
from .settings import load_settings, sync_settings_with_system
from .system_control import (
    get_keyboard_control_context,
    reboot_computer,
    set_keyboard_enabled,
    set_windows_app_id,
)
from .ui import KeyboardControlApp


def build_parser():
    parser = argparse.ArgumentParser(add_help=False)
    parser.add_argument(AUTOSTART_FLAG, dest="autostart", action="store_true")
    parser.add_argument(SET_STATE_FLAG, dest="set_state", choices=("enable", "disable"))
    parser.add_argument(REBOOT_NOW_FLAG, dest="reboot_now", action="store_true")
    parser.add_argument(
        CONTROL_MODE_FLAG,
        dest="control_mode",
        choices=(CONTROL_MODE_INSTANT, CONTROL_MODE_DRIVER),
        default=CONTROL_MODE_DRIVER,
    )
    parser.add_argument(TARGET_ID_FLAG, dest="target_ids", action="append", default=[])
    parser.add_argument(RESULT_FILE_FLAG, dest="result_file")
    return parser


def _write_helper_result(result_file, payload):
    if not result_file:
        return

    Path(result_file).write_text(
        json.dumps(payload or {}, ensure_ascii=False),
        encoding="utf-8",
    )


def _run_helper_action(args):
    if args.set_state:
        success, details, resolved_target_ids = set_keyboard_enabled(
            args.set_state == "enable",
            control_mode=args.control_mode,
            target_ids=args.target_ids,
        )
        _write_helper_result(
            args.result_file,
            {
                "success": bool(success),
                "details": details or "",
                "resolved_target_ids": list(resolved_target_ids or []),
            },
        )
        return 0 if success else 1

    if args.reboot_now:
        success, _ = reboot_computer()
        _write_helper_result(args.result_file, {"success": bool(success)})
        return 0 if success else 1

    return None


def _show_fatal_error(i18n, details):
    try:
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror(
            i18n.t("app.title"),
            i18n.t("app.fatal_error", details=details),
        )
        root.destroy()
    except Exception:
        print(details, file=sys.stderr)


def main(argv=None):
    parser = build_parser()
    args = parser.parse_args(argv)

    helper_exit_code = _run_helper_action(args)
    if helper_exit_code is not None:
        return helper_exit_code

    i18n = create_i18n()

    try:
        set_windows_app_id()
        settings = sync_settings_with_system(load_settings())
        control_context = get_keyboard_control_context(settings.get("instant_target_ids"))
        app = KeyboardControlApp(
            control_context=control_context,
            i18n=i18n,
            settings=settings,
            launched_from_autostart=bool(args.autostart),
        )
        app.mainloop()
        return 0
    except Exception:
        _show_fatal_error(i18n, traceback.format_exc())
        return 1
