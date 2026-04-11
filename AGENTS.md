# Agent Project Notes

## Environment
- This project uses **`uv`** to manage Python environment and dependencies.
- Prefer `uv run ...` instead of calling `python` directly.
- Prefer `uv pip ...` (or `uv sync` when lock/metadata exist) instead of `pip`.
- Do not assume global/site Python packages are available.

## Common commands
- Run app: `uv run python main.py`
- Compile check: `uv run python -m py_compile lock_laptop_keyboard\\app.py lock_laptop_keyboard\\ui.py`
- Build (if PyInstaller is installed in uv env): `uv run pyinstaller LockLaptopKeyboard.spec`

