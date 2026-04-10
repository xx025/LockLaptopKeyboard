APP_ID = "xx025.LockLaptopKeyboard"
SERVICE_NAME = "i8042prt"

AUTOSTART_FLAG = "--autostart"
SET_STATE_FLAG = "--set-state"
REBOOT_NOW_FLAG = "--reboot-now"
CONTROL_MODE_FLAG = "--control-mode"
TARGET_ID_FLAG = "--target-id"
RESULT_FILE_FLAG = "--result-file"
SETTINGS_DIR_NAME = "LockLaptopKeyboard"
SETTINGS_FILE_NAME = "settings.json"
RUN_KEY_PATH = r"Software\Microsoft\Windows\CurrentVersion\Run"
RUN_VALUE_NAME = "LockLaptopKeyboard"

CONTROL_MODE_INSTANT = "instant"
CONTROL_MODE_DRIVER = "driver"

WORK_COMMANDS = {
    "enable": "sc config i8042prt start= demand",
    "stop": "sc config i8042prt start= disabled",
}
ENABLED_START_TYPES = {"AUTO_START", "DEMAND_START"}

DEFAULT_SETTINGS = {
    "autostart_enabled": False,
    "start_minimized_to_tray": True,
    "preferred_control_mode": CONTROL_MODE_INSTANT,
    "instant_target_ids": [],
}

TRAY_COMMAND_SHOW = "show"
TRAY_COMMAND_DISABLE = "disable"
TRAY_COMMAND_ENABLE = "enable"
TRAY_COMMAND_REBOOT = "reboot"
TRAY_COMMAND_EXIT = "exit"
