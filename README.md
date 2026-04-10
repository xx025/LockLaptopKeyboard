# LockLaptopKeyboard

一个用于在 Windows 上禁用或恢复键盘设备的小工具，支持内置键盘和外置键盘。

当前版本已经完成这些能力：

- 移除 `pywebview`，改为原生 `tkinter + ttk` 窗口
- 支持 Windows 任务栏图标与系统托盘图标
- 支持“控制 / 设置”双标签页
- 即时模式下支持勾选具体键盘项，并显示每个目标的重启提示
- 支持当前用户开机自启与开机最小化到托盘
- 支持 `i18n`，会自动检测系统语言并在简体中文 / English 间切换
- 默认优先使用“即时设备模式”直接禁用/启用内置键盘设备，通常无需重启
- 保留“驱动兼容模式”，继续通过修改 `i8042prt` 服务启动类型实现键盘开关

## 使用方式

1. 运行 `python main.py`，或运行打包后的 `exe`
2. 查看当前键盘状态
3. 在即时设备模式下勾选要操作的键盘项，并查看每个目标旁边的“是否需要重启”提示
4. 点击“禁用键盘”或“启用键盘”
5. 即时设备模式下通常会立即生效；如果切到驱动模式，则需要重启电脑

关闭窗口时，程序会缩到 Windows 右下角系统托盘：

- 左键或双击托盘图标可恢复窗口
- 右键托盘图标可直接执行“显示窗口 / 禁用键盘 / 启用键盘 / 重启电脑 / 退出”

## 设置页

设置标签页目前支持：

- 开机自启
- 开机时最小化到托盘
- 查看当前跟随的系统语言
- 记住当前默认控制模式

开机自启使用当前用户的 Windows `Run` 启动项，不需要额外依赖。

## 项目结构

现在代码已经拆成更接近成熟项目的布局，不再全部堆在一个脚本里：

```text
main.py
lock_laptop_keyboard/
  app.py
  ui.py
  tray.py
  system_control.py
  i18n.py
  settings.py
  resources.py
  constants.py
```

## 控制模式

- 即时设备模式：识别当前可控的键盘 PnP 设备并直接禁用/启用，通常无需重启
- 驱动兼容模式：修改 `i8042prt` 启动类型，兼容旧机器或设备识别失败的情况，但需要重启
- 即时模式设备列表：会按“物理键盘项”展示候选目标；同一把键盘下的多个 HID 子接口会被合并为一项，并标记“需要 / 通常不需要 / 待验证”这类重启提示
- 默认会优先勾选更像内置键盘的目标；如果也想控制外置键盘，可以在列表中手动勾选

## 手动命令

禁用：

```shell
sc config i8042prt start= disabled
```

启用：

```shell
sc config i8042prt start= demand
```

## Linux

Linux 用户可以参考并使用 [toggle_keyboard.sh](toggle_keyboard.sh)。

## 下载

[下载软件](https://github.com/xx025/LockLaptopKeyboard/releases/download/1/LockLaptopKeyboard.exe)
