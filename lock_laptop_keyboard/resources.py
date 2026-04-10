import os
import sys
from pathlib import Path

from .constants import SETTINGS_DIR_NAME


def project_root():
    if hasattr(sys, "_MEIPASS"):
        return Path(sys._MEIPASS)
    return Path(__file__).resolve().parents[1]


def resource_path(*parts):
    return str(project_root().joinpath(*parts))


def entry_script_path():
    return project_root().joinpath("main.py")


def launcher_executable():
    executable = Path(sys.executable)
    if getattr(sys, "frozen", False):
        return str(executable)

    if executable.name.lower() == "python.exe":
        pythonw = executable.with_name("pythonw.exe")
        if pythonw.exists():
            return str(pythonw)

    return str(executable)


def app_data_dir():
    if sys.platform == "win32":
        base_dir = Path(os.getenv("APPDATA") or Path.home())
    else:
        base_dir = Path.home()
    return base_dir.joinpath(SETTINGS_DIR_NAME)
