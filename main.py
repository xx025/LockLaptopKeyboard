import subprocess

import webview
from pyuac import main_requires_admin

import ui


def run_command(command):
    process = subprocess.Popen(command, stdout=subprocess.PIPE, shell=True)
    result, err = process.communicate()
    try:
        result_str = result.decode("utf-8")
    except UnicodeDecodeError:
        result_str = result.decode("GBK")
    return result_str


def get_start_type(service_name):
    types = dict(
        auto="AUTO_START",
        disabled="DISABLED",
    )

    typeB = dict(
        auto='true',
        demand='true',
        disabled='false',
    )
    c = f'sc qc {service_name}'
    result_str = run_command(c)
    split_res = result_str.split("\n")
    for line in split_res:
        if "START_TYPE" in line:
            start_type = line.split(":")[1].strip()
            for k, v in types.items():
                if v in start_type:
                    return typeB[k]


work_commands = dict(
    plan1="sc config i8042prt start= demand",
    plan2="sc config i8042prt start= auto",
    stop="sc config i8042prt start= disabled",
)


class Api():
    def alter_state(self, status, plan):
        print(plan)
        if status == 'disable':
            command = work_commands['stop']
        else:
            command = work_commands[plan]
        print(command)
        run_command(command)

    def re_start_compute(self, n):
        run_command("shutdown -r -t 2")
        exit(1)


@main_requires_admin
def start_ui(html):
    webview.create_window("禁用笔记本内置键盘", html=html, width=330, height=500, fullscreen=False, js_api=Api())
    webview.start()


def main():
    service_name = 'i8042prt'
    service_status = get_start_type(service_name)
    # html=open('ui.html','r',encoding='utf8').read()
    html = ui.html
    html = html.replace("service_status", service_status)
    start_ui(html)


if __name__ == '__main__':
    main()
