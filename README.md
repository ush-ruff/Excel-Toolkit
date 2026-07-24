# Excel Toolkit 

![Version](https://img.shields.io/badge/version-1.4.0-blue)
[![Download](https://img.shields.io/badge/download-latest-green)](https://github.com/ush-ruff/Excel-Toolkit/releases/latest)

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

| Menu Item         | Action                                |
|-------------------|---------------------------------------|
| Open              | Edit the script                       |
| Edit Snippets     | Open `snippets.ini` in default editor |
| Toggle Hotstrings | Enable / disable hotstring expansion  |
| Reload Script     | Reload the AHK script                 |
| Pause Hotkeys     | Suspend all hotkeys                   |
| Exit              | Quit the script                       |

## Generating Docs

```
python gen-shortcuts.py
```

Regenerates `shortcuts.md` from the `; Key:` comments in the AHK source.

## Advanced Configuration

> Pre-built binaries are available on the [Releases page](https://github.com/ush-ruff/Excel-Toolkit/releases). Compiling is only needed if you want to modify the source.


<details>
<summary>Compiling (Installed Method)</summary>

You can compile the script into a standalone `.exe` using AutoHotkey’s compiler if it is already installed on your system.

### Steps:

1. Install [AutoHotkey v2](https://www.autohotkey.com/).
2. Locate `Ahk2Exe.exe` (included with AHK).
3. Run the compilation via command line, or right-click the script and choose **“Compile Script”** or **"Compile Script (GUI)"** if the option is available.
</details>

<details>
<summary>Compiling (Portable Method)</summary>

Generate `key-butler.exe` yourself without installing AutoHotkey:

1. **Download portable AutoHotkey v2** — [ahk-v2.zip](https://www.autohotkey.com/download/ahk-v2.zip)
   - Extract to a folder (e.g., `C:\tools\AutoHotkey-v2`)
2. **Download Ahk2Exe** — latest release from [AutoHotkey/Ahk2Exe](https://github.com/AutoHotkey/Ahk2Exe/releases)
   - Extract `Ahk2Exe.exe` into the same folder
3. **Compile** — use the base file matching your system (32-bit or 64-bit):
   ```powershell
   "C:\tools\AutoHotkey-v2\AutoHotkey64.exe" ^
     "C:\tools\AutoHotkey-v2\Ahk2Exe.exe" /in key-butler.ahk /out key-butler.exe
   ```
The resulting `.exe` is fully standalone — no runtime needed.
</details>

<details>
<summary>Running on Startup</summary>

To launch the script automatically on Windows startup:

1. Press **Win + R** and type
   ```
   shell:startup
   ```

2. Press Enter and place either:
   - `excel-toolkit.ahk`, or
   - the compiled `.exe` (recommended) — [download from releases](https://github.com/ush-ruff/Excel-Toolkit/releases/latest)
   
   > a shortcut to either of those files will also work

3. Place the `snippets.ini` in the same folder

The script will now start automatically when you log in.
</details>


## Version Support

- **AutoHotkey v2.0+ only**
- Tested on Windows 10 and Windows 11

Make sure you are running the **Unicode 64-bit** build of AutoHotkey v2 for best results.

## License

[GNU GPLv3](./LICENSE).
