## Microsoft Icon Source

- Official icon source file:
  `img/ic_fluent_keyboard_24_regular.svg`
- Upstream repository:
  `https://github.com/microsoft/fluentui-system-icons`
- Upstream file URL:
  `https://raw.githubusercontent.com/microsoft/fluentui-system-icons/main/assets/Keyboard/SVG/ic_fluent_keyboard_24_regular.svg`
- License:
  MIT (copied to `img/MICROSOFT_FLUENT_ICONS_LICENSE.txt`)

## Packaged Asset

- Base icon file used by this app:
  `img/ms_keyboard_tray.ico`
- Used for:
  - window icon
  - tray icon
  - PyInstaller exe icon (see `LockLaptopKeyboard.spec`)
- Runtime icon rendering:
  - The app now generates a higher-contrast blue badge icon at startup for the
    window icon and tray icon.
  - The keyboard glyph still comes from `img/ic_fluent_keyboard_24_regular.svg`.
  - The generated tray `.ico` is cached under `%APPDATA%\\LockLaptopKeyboard`.
