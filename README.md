# Excel Toolkit ![Version](https://img.shields.io/badge/version-1.4.0-blue)

> AutoHotkey v2 script that supercharges Microsoft Excel with custom keyboard shortcuts, text-expansion hotstrings, and formatting helpers.

## Quick Start

- **Standalone** — Download the latest release from the [Releases page](https://github.com/ush-ruff/Excel-Toolkit/releases) (no dependencies required)
- **Source** — Double-click `excel-toolkit.ahk` (requires [AutoHotkey v2](https://www.autohotkey.com/))

Shortcuts are active only when an Excel window is focused.

## Features

### Keyboard Shortcuts

| Category | Shortcuts |
|----------|-----------|
| Merge & Unmerge | Merge across, merge center, unmerge |
| Fill Colors | Clear, light blue, rose, gray, red |
| Font Colors | Automatic, blue, red, white |
| Text & Alignment | Wrap text, left/center/right align, top/middle/down align |
| Font Size | Increase / decrease by 1 |
| Number Formats | Accounting, general, percentage |
| Borders | Remove, outside, all borders |
| Quick Actions | Toggle freeze panes, clean format, auto-fit, delete row |

> [Full shortcut reference &rarr; `shortcuts.md`](./shortcuts.md)

### Hotstring System (Text Expansion)

Type a trigger word followed by `;` inside a cell and it auto-expands.

- **Toggle on/off** — <kbd>Win</kbd> + <kbd>`</kbd>
- **Customize snippets** — Edit `snippets.ini` (or use the tray menu)
- **Example:** `dt;` expands to `=TODAY()`
- Supports `\s` (space), `\t` (tab), `\n` (newline), and AHK send specials like `{Enter}`

### System Tray

Right-click the tray icon to access:

| Menu Item | Action |
|-----------|--------|
| Open | Edit the script |
| Edit Snippets | Open `snippets.ini` in default editor |
| Toggle Hotstrings | Enable / disable hotstring expansion |
| Reload Script | Reload the AHK script |
| Pause Hotkeys | Suspend all hotkeys |
| Exit | Quit the script |

## Compiling from Source (Portable Method)

Generate `excel-toolkit.exe` yourself without installing AutoHotkey:

1. **Download portable AutoHotkey v2** — [ahk-v2.zip](https://www.autohotkey.com/download/ahk-v2.zip)
   - Extract to a folder (e.g., `C:\tools\AutoHotkey-v2`)
   - The zip includes both `AutoHotkey32.exe` (32-bit) and `AutoHotkey64.exe` (64-bit)
2. **Download Ahk2Exe** — latest release from [AutoHotkey/Ahk2Exe](https://github.com/AutoHotkey/Ahk2Exe/releases)
   - Extract `Ahk2Exe.exe` into the same folder
3. **Compile** — use the base file matching your system (32-bit or 64-bit):
   ```
   "C:\tools\AutoHotkey-v2\AutoHotkey64.exe" ^
     "C:\tools\AutoHotkey-v2\Ahk2Exe.exe" /in excel-toolkit.ahk /out excel-toolkit.exe
   ```
   The resulting `.exe` is fully standalone — no runtime needed.

## Generating Docs

```
python gen-shortcuts.py
```

Regenerates `shortcuts.md` from the `; Key:` comments in the AHK source.

## License

[GNU GPLv3](./LICENSE).
