import ctypes
import json
import subprocess
import sys
import tempfile
import time
from ctypes import wintypes
from pathlib import Path

from .constants import (
    APP_ID,
    CONTROL_MODE_DRIVER,
    CONTROL_MODE_FLAG,
    CONTROL_MODE_INSTANT,
    ENABLED_START_TYPES,
    REBOOT_NOW_FLAG,
    RESULT_FILE_FLAG,
    SERVICE_NAME,
    SET_STATE_FLAG,
    TARGET_ID_FLAG,
    WORK_COMMANDS,
)
from .resources import entry_script_path, launcher_executable, project_root


SEE_MASK_NOCLOSEPROCESS = 0x00000040
SW_HIDE = 0
INFINITE = 0xFFFFFFFF
ERROR_CANCELLED = 1223
CREATE_NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0x08000000)
STARTF_USESHOWWINDOW = getattr(subprocess, "STARTF_USESHOWWINDOW", 0x00000001)


class SHELLEXECUTEINFOW(ctypes.Structure):
    _fields_ = [
        ("cbSize", wintypes.DWORD),
        ("fMask", wintypes.ULONG),
        ("hwnd", wintypes.HWND),
        ("lpVerb", wintypes.LPCWSTR),
        ("lpFile", wintypes.LPCWSTR),
        ("lpParameters", wintypes.LPCWSTR),
        ("lpDirectory", wintypes.LPCWSTR),
        ("nShow", ctypes.c_int),
        ("hInstApp", wintypes.HINSTANCE),
        ("lpIDList", wintypes.LPVOID),
        ("lpClass", wintypes.LPCWSTR),
        ("hkeyClass", wintypes.HKEY),
        ("dwHotKey", wintypes.DWORD),
        ("hIconOrMonitor", wintypes.HANDLE),
        ("hProcess", wintypes.HANDLE),
    ]


shell32 = ctypes.WinDLL("shell32", use_last_error=True)
kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)

shell32.ShellExecuteExW.argtypes = [ctypes.POINTER(SHELLEXECUTEINFOW)]
shell32.ShellExecuteExW.restype = wintypes.BOOL
shell32.SetCurrentProcessExplicitAppUserModelID.argtypes = [wintypes.LPCWSTR]
shell32.SetCurrentProcessExplicitAppUserModelID.restype = ctypes.c_long
shell32.IsUserAnAdmin.argtypes = []
shell32.IsUserAnAdmin.restype = wintypes.BOOL
kernel32.WaitForSingleObject.argtypes = [wintypes.HANDLE, wintypes.DWORD]
kernel32.WaitForSingleObject.restype = wintypes.DWORD
kernel32.GetExitCodeProcess.argtypes = [wintypes.HANDLE, ctypes.POINTER(wintypes.DWORD)]
kernel32.GetExitCodeProcess.restype = wintypes.BOOL
kernel32.CloseHandle.argtypes = [wintypes.HANDLE]
kernel32.CloseHandle.restype = wintypes.BOOL


def decode_output(data):
    if isinstance(data, str):
        return data

    if data is None:
        return ""

    for encoding in ("utf-8", "gbk", sys.getfilesystemencoding() or "utf-8"):
        try:
            return data.decode(encoding).strip()
        except (LookupError, UnicodeDecodeError):
            continue

    return data.decode(errors="replace").strip()


def _hidden_subprocess_kwargs():
    if sys.platform != "win32":
        return {}

    startupinfo = subprocess.STARTUPINFO()
    startupinfo.dwFlags |= STARTF_USESHOWWINDOW
    startupinfo.wShowWindow = SW_HIDE
    return {
        "creationflags": CREATE_NO_WINDOW,
        "startupinfo": startupinfo,
    }


def run_command(command):
    args = command if isinstance(command, (list, tuple)) else str(command).split()
    completed = subprocess.run(
        args,
        capture_output=True,
        check=False,
        **_hidden_subprocess_kwargs(),
    )
    return completed.returncode, decode_output(completed.stdout), decode_output(completed.stderr)


def _as_list(value):
    if value is None:
        return []
    if isinstance(value, list):
        return value
    return [value]


INTERESTING_PNP_KEYS = {
    "DEVPKEY_Device_Service",
    "DEVPKEY_Device_EnumeratorName",
    "DEVPKEY_Device_Parent",
    "DEVPKEY_Device_LocationPaths",
    "DEVPKEY_Device_HardwareIds",
    "DEVPKEY_Device_ContainerId",
    "DEVPKEY_Device_InLocalMachineContainer",
    "DEVPKEY_Device_IsPresent",
    "DEVPKEY_Device_ProblemCode",
}
LIST_PNP_KEYS = {
    "DEVPKEY_Device_LocationPaths",
    "DEVPKEY_Device_HardwareIds",
}

INSTANCE_ID_PREFIXES = ("Instance ID:", "Instance ID：", "实例 ID:", "实例 ID：")
DEVICE_DESCRIPTION_PREFIXES = (
    "Device Description:",
    "Device Description：",
    "设备描述:",
    "设备描述：",
)
STATUS_PREFIXES = ("Status:", "Status：", "状态:", "状态：")
PROPERTIES_PREFIXES = ("Properties:", "Properties：", "属性:", "属性：")

RESTART_REQUIRED = "required"
RESTART_USUALLY_NOT_REQUIRED = "usually_not_required"
RESTART_VERIFY = "verify"


def _pnputil_bool(value, default=False):
    normalized = str(value or "").strip().upper()
    if normalized == "TRUE":
        return True
    if normalized == "FALSE":
        return False
    return default


def _extract_prefixed_value(line, prefixes):
    stripped = line.lstrip()
    for prefix in prefixes:
        if stripped.startswith(prefix):
            return stripped[len(prefix) :].strip()
    return None


def _parse_pnputil_properties_output(output):
    devices = []
    current = None
    current_property = None

    def finalize(device):
        if not device or not device.get("instance_id"):
            return
        properties = device.pop("_properties", {})
        device["service"] = str(properties.get("DEVPKEY_Device_Service", "") or "")
        device["enumerator_name"] = str(properties.get("DEVPKEY_Device_EnumeratorName", "") or "")
        device["parent"] = str(properties.get("DEVPKEY_Device_Parent", "") or "")
        device["location_paths"] = [
            str(item) for item in _as_list(properties.get("DEVPKEY_Device_LocationPaths")) if item
        ]
        device["hardware_ids"] = [
            str(item) for item in _as_list(properties.get("DEVPKEY_Device_HardwareIds")) if item
        ]
        device["container_id"] = str(properties.get("DEVPKEY_Device_ContainerId", "") or "")
        device["in_local_machine_container"] = _pnputil_bool(
            properties.get("DEVPKEY_Device_InLocalMachineContainer"),
            default=device.get("in_local_machine_container", False),
        )
        device["present"] = _pnputil_bool(
            properties.get("DEVPKEY_Device_IsPresent"),
            default=device.get("present", False),
        )
        problem_code = str(properties.get("DEVPKEY_Device_ProblemCode", "") or "").strip()
        if problem_code and problem_code not in {"0x00000000 (0)", "0"}:
            device["problem"] = problem_code
        devices.append(device)

    for raw_line in output.splitlines():
        line = raw_line.rstrip()
        stripped = line.strip()
        if not stripped:
            continue

        instance_id = _extract_prefixed_value(line, INSTANCE_ID_PREFIXES)
        if instance_id is not None:
            finalize(current)
            current = {
                "status": "",
                "friendly_name": "",
                "instance_id": instance_id,
                "present": True,
                "problem": "",
                "_properties": {},
            }
            current_property = None
            continue

        if current is None:
            continue

        device_description = _extract_prefixed_value(line, DEVICE_DESCRIPTION_PREFIXES)
        if device_description is not None:
            current["friendly_name"] = device_description
            continue

        status_value = _extract_prefixed_value(line, STATUS_PREFIXES)
        if status_value is not None:
            current["status"] = status_value
            continue

        if stripped in PROPERTIES_PREFIXES:
            current_property = None
            continue

        if line.startswith("    ") and not line.startswith("        "):
            property_header = stripped
            property_name = property_header.split(" [", 1)[0]
            if property_name in INTERESTING_PNP_KEYS:
                current_property = property_name
                if property_name in LIST_PNP_KEYS:
                    current["_properties"][property_name] = []
                else:
                    current["_properties"][property_name] = None
            else:
                current_property = None
            continue

        if current_property and line.startswith("        "):
            property_value = line.strip()
            if current_property in LIST_PNP_KEYS:
                if property_value:
                    current["_properties"][current_property].append(property_value)
            elif current["_properties"][current_property] is None and property_value:
                current["_properties"][current_property] = property_value

    finalize(current)
    return devices


def _query_keyboard_devices_connected():
    code, stdout, stderr = run_command(
        ["pnputil", "/enum-devices", "/class", "Keyboard", "/connected", "/properties"]
    )
    if code != 0:
        return [], stderr or stdout
    return _parse_pnputil_properties_output(stdout), ""


def _query_keyboard_device_by_id(instance_id):
    code, stdout, stderr = run_command(
        ["pnputil", "/enum-devices", "/instanceid", instance_id, "/properties"]
    )
    if code != 0:
        return None, stderr or stdout

    devices = _parse_pnputil_properties_output(stdout)
    return (devices[0] if devices else None), ""


def _collect_keyboard_devices(extra_instance_ids=None):
    connected_devices, details = _query_keyboard_devices_connected()
    if details:
        return [], details

    device_map = {
        device["instance_id"].upper(): device for device in connected_devices if device.get("instance_id")
    }

    extra_details = []
    for instance_id in extra_instance_ids or []:
        key = str(instance_id or "").strip().upper()
        if not key or key in device_map:
            continue
        device, error = _query_keyboard_device_by_id(instance_id)
        if device:
            device_map[key] = device
        elif error:
            extra_details.append(error)

    return list(device_map.values()), "\n".join(extra_details).strip()


def _classify_keyboard_device(device):
    instance_id = device.get("instance_id", "").upper()
    service = device.get("service", "").upper()
    enumerator_name = device.get("enumerator_name", "").upper()
    parent = device.get("parent", "").upper()
    location_text = " ".join(device.get("location_paths", [])).upper()
    hardware_text = " ".join(device.get("hardware_ids", [])).upper()
    in_local_container = bool(device.get("in_local_machine_container"))

    if instance_id.startswith("HID\\GVINPUT"):
        return {
            "priority": 35,
            "reason": "virtual_remote",
            "display_name": device.get("friendly_name") or device.get("instance_id"),
            "recommended_selected": False,
        }

    if parent.startswith("ROOT\\HIDCLASS\\"):
        return {
            "priority": 35,
            "reason": "virtual_remote",
            "display_name": device.get("friendly_name") or device.get("instance_id"),
            "recommended_selected": False,
        }

    if service == "I8042PRT" or enumerator_name == "ACPI" or "ACPI(PS2K)" in location_text:
        return {
            "priority": 100,
            "reason": "acpi_ps2",
            "display_name": device.get("friendly_name") or device.get("instance_id"),
            "recommended_selected": True,
        }

    if instance_id.startswith("HID\\CONVERTEDDEVICE") or parent.startswith("BUTTONCONVERTER\\CONVERTEDDEVICE"):
        return {
            "priority": 80,
            "reason": "converted_device",
            "display_name": device.get("friendly_name") or device.get("instance_id"),
            "recommended_selected": True,
        }

    if "MSFT0001" in instance_id or "*PNP0303" in hardware_text:
        return {
            "priority": 90,
            "reason": "keyboard_controller",
            "display_name": device.get("friendly_name") or device.get("instance_id"),
            "recommended_selected": True,
        }

    if parent.startswith("USB\\") or instance_id.startswith("USB\\"):
        return {
            "priority": 70,
            "reason": "external_usb",
            "display_name": device.get("friendly_name") or device.get("instance_id"),
            "recommended_selected": False,
        }

    if instance_id.startswith("HID\\VID_") and parent.startswith("USB\\"):
        return {
            "priority": 70,
            "reason": "external_usb",
            "display_name": device.get("friendly_name") or device.get("instance_id"),
            "recommended_selected": False,
        }

    if (
        parent.startswith("BTH")
        or enumerator_name in {"BTHENUM", "BLUETOOTH"}
        or "BTH" in location_text
    ):
        return {
            "priority": 60,
            "reason": "external_bluetooth",
            "display_name": device.get("friendly_name") or device.get("instance_id"),
            "recommended_selected": False,
        }

    if service == "KBDHID" and not in_local_container:
        return {
            "priority": 55,
            "reason": "external_hid",
            "display_name": device.get("friendly_name") or device.get("instance_id"),
            "recommended_selected": False,
        }

    if service == "KBDHID":
        return {
            "priority": 50,
            "reason": "hid_keyboard",
            "display_name": device.get("friendly_name") or device.get("instance_id"),
            "recommended_selected": False,
        }

    if enumerator_name in {"HID", "USB"}:
        return {
            "priority": 45,
            "reason": "external_hid" if not in_local_container else "hid_keyboard",
            "display_name": device.get("friendly_name") or device.get("instance_id"),
            "recommended_selected": False,
        }

    return {
        "priority": 40,
        "reason": "unknown",
        "display_name": device.get("friendly_name") or device.get("instance_id"),
        "recommended_selected": False,
    }


def _device_is_disabled(device):
    status = device.get("status", "").upper()
    problem = device.get("problem", "").upper()

    if "DISABLED" in problem:
        return True

    if "(22)" in problem:
        return True

    if status == "ERROR" and problem:
        return True

    return False


def _estimate_restart_requirement(device):
    reason = str(device.get("reason", "") or "").lower()
    service = str(device.get("service", "") or "").upper()
    parent = str(device.get("parent", "") or "").upper()
    enumerator_name = str(device.get("enumerator_name", "") or "").upper()

    if reason in {"acpi_ps2", "keyboard_controller"}:
        return RESTART_REQUIRED

    if service == "I8042PRT" or enumerator_name == "ACPI":
        return RESTART_REQUIRED

    if reason == "converted_device" or parent.startswith("BUTTONCONVERTER\\CONVERTEDDEVICE"):
        return RESTART_VERIFY

    if reason in {"external_usb", "external_bluetooth", "external_hid"}:
        return RESTART_VERIFY

    if reason == "virtual_remote":
        return RESTART_VERIFY

    if reason == "hid_keyboard":
        return RESTART_VERIFY

    if service == "KBDHID":
        return RESTART_VERIFY

    return RESTART_VERIFY


def _container_id_is_usable(container_id):
    normalized = str(container_id or "").strip("{} ").upper()
    if not normalized:
        return False
    if normalized == "00000000-0000-0000-FFFF-FFFFFFFFFFFF":
        return False
    if normalized == "00000000-0000-0000-0000-000000000000":
        return False
    return True


def _group_key_for_device(device):
    container_id = str(device.get("container_id", "") or "").strip().upper()
    reason = str(device.get("reason", "") or "").lower()

    if reason in {"external_usb", "external_bluetooth", "external_hid", "hid_keyboard", "virtual_remote"}:
        if _container_id_is_usable(container_id):
            return f"container:{container_id}"

    return f"instance:{str(device.get('instance_id', '') or '').strip().upper()}"


def _merge_restart_requirement(current, incoming):
    ranking = {
        RESTART_REQUIRED: 3,
        RESTART_VERIFY: 2,
        RESTART_USUALLY_NOT_REQUIRED: 1,
        "": 0,
    }
    if ranking.get(incoming, 0) > ranking.get(current, 0):
        return incoming
    return current


def _build_device_groups(discovered_targets, cached_target_ids=None):
    cached_target_keys = {target_id.upper() for target_id in _unique_strings(cached_target_ids)}
    grouped = {}
    ordered_groups = []

    for device in discovered_targets:
        group_key = _group_key_for_device(device)
        group = grouped.get(group_key)
        if group is None:
            group = {
                "group_key": group_key,
                "display_name": device.get("display_name") or device.get("friendly_name") or device.get("instance_id"),
                "reason": device.get("reason", "unknown"),
                "restart_requirement": device.get("restart_requirement", RESTART_VERIFY),
                "recommended_selected": bool(device.get("recommended_selected")),
                "priority": int(device.get("priority", 0)),
                "member_devices": [],
                "member_instance_ids": [],
                "present": False,
                "selected": False,
            }
            grouped[group_key] = group
            ordered_groups.append(group)

        group["member_devices"].append(device)
        if device.get("instance_id"):
            group["member_instance_ids"].append(device["instance_id"])
            if device["instance_id"].upper() in cached_target_keys:
                group["selected"] = True
        group["present"] = group["present"] or bool(device.get("present"))
        group["recommended_selected"] = group["recommended_selected"] or bool(
            device.get("recommended_selected")
        )
        group["priority"] = max(group["priority"], int(device.get("priority", 0)))
        group["restart_requirement"] = _merge_restart_requirement(
            group["restart_requirement"],
            device.get("restart_requirement", RESTART_VERIFY),
        )

    for group in ordered_groups:
        group["member_count"] = len(group["member_devices"])

    ordered_groups.sort(
        key=lambda item: (-int(item.get("priority", 0)), item.get("display_name", "").upper(), item.get("group_key", ""))
    )
    return ordered_groups


def _unique_strings(items):
    seen = set()
    result = []
    for item in items or []:
        value = str(item or "").strip()
        key = value.upper()
        if not value or key in seen:
            continue
        seen.add(key)
        result.append(value)
    return result


def _target_state_snapshot(target_ids):
    target_ids = _unique_strings(target_ids)
    devices, _ = _collect_keyboard_devices(extra_instance_ids=target_ids)
    devices_by_id = {device["instance_id"].upper(): device for device in devices if device.get("instance_id")}

    states = []
    for target_id in target_ids:
        device = devices_by_id.get(target_id.upper())
        if not device:
            states.append(None)
        elif _device_is_disabled(device):
            states.append(False)
        elif device.get("present"):
            states.append(True)
        else:
            states.append(None)
    return states


def _targets_match_desired_state(target_ids, enabled):
    states = _target_state_snapshot(target_ids)
    if not states:
        return False

    if enabled:
        return all(state is True for state in states)

    return all(state in {False, None} for state in states)


def _wait_for_target_state(target_ids, enabled):
    for delay_seconds in (0.2, 0.6, 1.2, 2.0, 3.0):
        time.sleep(delay_seconds)
        if _targets_match_desired_state(target_ids, enabled):
            return True
    return False


def _output_mentions_reboot(text):
    normalized = str(text or "").upper()
    return any(
        marker in normalized
        for marker in (
            "REBOOT",
            "RESTART",
            "\u91cd\u65b0\u542f\u52a8",
            "\u91cd\u65b0\u555f\u52d5",
            "\u91cd\u542f",
            "\u91cd\u65b0\u958b\u6a5f",
        )
    )


def _read_helper_result(result_file):
    if not result_file:
        return {}

    path = Path(result_file)
    try:
        if not path.exists():
            return {}

        raw = path.read_text(encoding="utf-8").strip()
        if not raw:
            return {}
        return json.loads(raw)
    except (OSError, ValueError, TypeError):
        return {}
    finally:
        try:
            path.unlink()
        except OSError:
            pass


def get_driver_keyboard_enabled():
    code, stdout, stderr = run_command(["sc", "qc", SERVICE_NAME])
    output = stdout or stderr

    if code != 0:
        return None, output

    for line in stdout.splitlines():
        if "START_TYPE" not in line.upper():
            continue

        upper_line = line.upper()
        if "DISABLED" in upper_line:
            return False, stdout
        if any(start_type in upper_line for start_type in ENABLED_START_TYPES):
            return True, stdout

    return None, stdout


def get_keyboard_control_context(cached_target_ids=None):
    cached_target_ids = _unique_strings(cached_target_ids)
    cached_target_keys = {target_id.upper() for target_id in cached_target_ids}
    devices, details = _collect_keyboard_devices(extra_instance_ids=cached_target_ids)
    devices_by_id = {device["instance_id"].upper(): device for device in devices if device.get("instance_id")}

    discovered_targets = []
    for device in devices:
        classification = _classify_keyboard_device(device)
        if not classification:
            continue

        if not device.get("present") and device.get("instance_id", "").upper() not in cached_target_keys:
            continue

        enriched = dict(device)
        enriched.update(classification)
        enriched["restart_requirement"] = _estimate_restart_requirement(enriched)
        discovered_targets.append(enriched)

    discovered_targets.sort(
        key=lambda item: (-int(item.get("priority", 0)), item.get("instance_id", "").upper())
    )
    discovered_target_ids = [item["instance_id"] for item in discovered_targets]
    discovered_targets_by_id = {
        item["instance_id"].upper(): item for item in discovered_targets if item.get("instance_id")
    }
    discovered_groups = _build_device_groups(discovered_targets, cached_target_ids=cached_target_ids)
    default_groups = [group for group in discovered_groups if group.get("recommended_selected")]

    selected_groups = [group for group in discovered_groups if group.get("selected")]
    if not selected_groups:
        selected_groups = list(default_groups)

    target_ids = _unique_strings(
        member_id
        for group in selected_groups
        for member_id in group.get("member_instance_ids", [])
    )
    target_devices = [
        discovered_targets_by_id[target_id.upper()]
        for target_id in target_ids
        if target_id.upper() in discovered_targets_by_id
    ]
    ignored_devices = [
        device
        for device in devices
        if device.get("instance_id", "").upper() not in discovered_targets_by_id
    ]

    instant_enabled = None
    if target_ids:
        states = []
        for target_id in target_ids:
            device = discovered_targets_by_id.get(target_id.upper()) or devices_by_id.get(target_id.upper())
            if not device:
                states.append(None)
            elif _device_is_disabled(device):
                states.append(False)
            elif device.get("present"):
                states.append(True)
            else:
                states.append(None)

        if states and all(state is True for state in states):
            instant_enabled = True
        elif any(state is False for state in states):
            instant_enabled = False

    driver_enabled, driver_details = get_driver_keyboard_enabled()

    return {
        "driver_enabled": driver_enabled,
        "driver_details": driver_details,
        "instant_available": bool(discovered_groups),
        "instant_enabled": instant_enabled,
        "instant_target_ids": target_ids,
        "instant_target_devices": target_devices,
        "instant_selected_target_groups": selected_groups,
        "instant_available_target_ids": discovered_target_ids,
        "instant_available_target_devices": discovered_targets,
        "instant_available_target_groups": discovered_groups,
        "instant_discovered_targets": discovered_targets,
        "instant_ignored_devices": ignored_devices,
        "instant_details": details,
    }


def set_driver_keyboard_enabled(enabled):
    command = WORK_COMMANDS["enable" if enabled else "stop"]
    code, stdout, stderr = run_command(command)
    return code == 0, stdout or stderr, []


def set_instant_keyboard_enabled(enabled, target_ids=None):
    target_ids = _unique_strings(target_ids)
    if not target_ids:
        context = get_keyboard_control_context()
        target_ids = context.get("instant_target_ids", [])

    if not target_ids:
        return False, "no_internal_targets", []

    outputs = []
    success = True
    for target_id in target_ids:
        args = ["pnputil", "/enable-device" if enabled else "/disable-device", target_id]
        if not enabled:
            args.append("/force")
        code, stdout, stderr = run_command(args)
        success = success and code == 0
        outputs.append((target_id, stdout or stderr))

    run_command(["pnputil", "/scan-devices"])
    combined_output = "\n".join(message for _, message in outputs if message)

    if success and _wait_for_target_state(target_ids, enabled):
        details = "\n\n".join(
            f"[{target_id}]\n{message}".strip() for target_id, message in outputs if message
        )
        return True, details, target_ids

    if _output_mentions_reboot(combined_output):
        return False, "instant_requires_reboot", target_ids

    if success:
        return False, "instant_no_state_change", target_ids

    details = "\n\n".join(
        f"[{target_id}]\n{message}".strip() for target_id, message in outputs if message
    )
    return success, details, target_ids


def set_keyboard_enabled(enabled, control_mode=CONTROL_MODE_DRIVER, target_ids=None):
    if control_mode == CONTROL_MODE_INSTANT:
        return set_instant_keyboard_enabled(enabled, target_ids=target_ids)
    return set_driver_keyboard_enabled(enabled)


def set_windows_app_id():
    if sys.platform != "win32":
        return

    try:
        shell32.SetCurrentProcessExplicitAppUserModelID(APP_ID)
    except (AttributeError, OSError):
        pass


def is_admin():
    if sys.platform != "win32":
        return False

    try:
        return bool(shell32.IsUserAnAdmin())
    except (AttributeError, OSError):
        return False


def _build_helper_parameters(extra_arguments):
    if getattr(sys, "frozen", False):
        parts = list(extra_arguments)
    else:
        parts = [str(entry_script_path()), *extra_arguments]
    return subprocess.list2cmdline(parts)


def run_elevated(extra_arguments):
    if sys.platform != "win32":
        return False, "Elevation is only supported on Windows."

    with tempfile.NamedTemporaryFile(prefix="llk-helper-", suffix=".json", delete=False) as handle:
        result_file = handle.name

    executable = sys.executable if getattr(sys, "frozen", False) else launcher_executable()
    execute_info = SHELLEXECUTEINFOW()
    execute_info.cbSize = ctypes.sizeof(SHELLEXECUTEINFOW)
    execute_info.fMask = SEE_MASK_NOCLOSEPROCESS
    execute_info.lpVerb = "runas"
    execute_info.lpFile = executable
    execute_info.lpParameters = _build_helper_parameters(
        [*extra_arguments, RESULT_FILE_FLAG, result_file]
    )
    execute_info.lpDirectory = str(project_root())
    execute_info.nShow = SW_HIDE

    if not shell32.ShellExecuteExW(ctypes.byref(execute_info)):
        error_code = ctypes.get_last_error()
        if error_code == ERROR_CANCELLED:
            _read_helper_result(result_file)
            return False, "cancelled", {}
        _read_helper_result(result_file)
        return False, ctypes.FormatError(error_code).strip(), {}

    if execute_info.hProcess:
        try:
            kernel32.WaitForSingleObject(execute_info.hProcess, INFINITE)
            exit_code = wintypes.DWORD()
            if not kernel32.GetExitCodeProcess(execute_info.hProcess, ctypes.byref(exit_code)):
                return False, ctypes.FormatError(ctypes.get_last_error()).strip(), {}
            if exit_code.value != 0:
                payload = _read_helper_result(result_file)
                return (
                    False,
                    str(payload.get("details") or f"helper exited with code {exit_code.value}"),
                    payload,
                )
        finally:
            kernel32.CloseHandle(execute_info.hProcess)

    payload = _read_helper_result(result_file)
    return bool(payload.get("success", True)), str(payload.get("details", "") or ""), payload


def set_keyboard_enabled_via_uac(enabled, control_mode=CONTROL_MODE_DRIVER, target_ids=None):
    extra_arguments = [
        SET_STATE_FLAG,
        "enable" if enabled else "disable",
        CONTROL_MODE_FLAG,
        control_mode or CONTROL_MODE_DRIVER,
    ]
    for target_id in _unique_strings(target_ids):
        extra_arguments.extend([TARGET_ID_FLAG, target_id])
    success, details, payload = run_elevated(extra_arguments)
    resolved_target_ids = _unique_strings(payload.get("resolved_target_ids") or target_ids)
    return success, details, resolved_target_ids


def reboot_computer():
    code, stdout, stderr = run_command(["shutdown", "/r", "/t", "0"])
    return code == 0, stdout or stderr


def reboot_computer_via_uac():
    success, details, _ = run_elevated([REBOOT_NOW_FLAG])
    return success, details
