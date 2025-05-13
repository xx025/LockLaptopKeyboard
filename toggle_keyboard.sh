#!/bin/bash

# 键盘名称
keyboard_name="AT Translated Set 2 keyboard"

# 获取键盘的设备 ID
device_info=$(xinput list | grep -i "$keyboard_name")
keyboard_id=$(echo "$device_info" | sed 's/.*id=\([0-9]*\).*/\1/')


# 判断键盘是否在线
if xinput list |  grep "$keyboard_id" | grep -q "floating slave"; then
    # 如果键盘是 floating slave 状态（离线状态），启用它
    xinput enable "$keyboard_id"
    echo "键盘已启用"
else
    # 如果键盘是 slave 状态（在线状态），禁用它
    xinput disable "$keyboard_id"
    echo "键盘已禁用"
fi

# 等待用户输入以保持终端窗口开启
echo "按 Enter 键退出..."
read
