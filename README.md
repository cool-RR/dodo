# Dodo Desktop Switcher

A lightweight Windows system tray application that adds convenient keyboard shortcuts for Windows 10/11 virtual desktops.

## The Problem

Windows 10 and 11 have excellent built-in virtual desktop support, but the default keyboard shortcuts are awkward:

- **Ctrl+Win+Left/Right** - Hard to remember which direction is which desktop
- **No direct jumping** - You must cycle through desktops one by one to reach desktop 5
- **Awkward window moving** - Moving windows between desktops requires multiple steps
- **No quick "return to previous"** - Can't easily toggle between two workspaces

If you're trying to organize your work across multiple virtual desktops, these limitations make the feature frustrating to use.

## The Solution

Dodo provides intuitive keyboard shortcuts that make virtual desktops actually usable:

### Basic Shortcuts
- **Alt+1 through Alt+9** - Jump directly to desktops 1-9
- **Alt+0** - Jump to desktop 10
- **Alt+-** - Return to your previous desktop (like Alt+Tab but for desktops)
- **Alt+Shift+1 through Alt+Shift+0** - Move the active window to a specific desktop
- **Alt+Shift+`** - Pin/unpin the active window to all desktops

### Advanced Features
- **Ctrl+Alt+-** - Generate a snapshot of your current window layout (Toxita feature)
- **Ctrl+Alt+Shift+-** - Show desktop number overlay on demand
- **Desktop Names** - Customize desktop names shown in overlays
- **Auto-Pinning** - Automatically pin windows based on title patterns
- **Desktop Monitoring** - Automatically show overlay when desktops change

Dodo automatically ensures you have 10 virtual desktops available and shows a brief on-screen indicator when you switch desktops.

## Installation

```bash
pip install dodo-desktop-switcher
```

### Auto-start with Windows

To have Dodo start automatically when you log in:

```bash
dodo --install
```

This creates a shortcut in your Windows Startup folder. Dodo will run silently in your system tray.

To remove from startup:

```bash
dodo --uninstall
```

To check installation status:

```bash
dodo --status
```

## Usage

After installation, simply run:

```bash
dodo
```

Dodo will appear in your system tray. The keyboard shortcuts are immediately active.

Right-click the system tray icon to:
- View all available shortcuts
- Manually switch desktops via menu
- Exit the application

## Requirements

- Windows 10 or Windows 11
- Python 3.12 or later
- Virtual desktops must be enabled (they are by default)

## How It Works

Dodo doesn't create or manage virtual desktops itself - Windows does all the heavy lifting. Dodo simply:

1. Uses the `pyvda` library to interface with Windows' built-in Virtual Desktop API
2. Registers global hotkeys using `pywin32`
3. Provides a system tray interface using `wxPython`

This means Dodo is lightweight, reliable, and works seamlessly with Windows' native virtual desktop features.

## Advanced Features

### Toxita: Window Layout Management

Dodo includes a powerful feature called "Toxita" that lets you save and restore window layouts across desktops:

1. **Generate a snapshot**: Press `Ctrl+Alt+-` to create `~/.dodo/toxita.yaml` with all your current windows organized by desktop
2. **Edit the layout**: The YAML file will open automatically - move window entries between desktops by editing the file
3. **Auto-apply changes**: Dodo watches the file - when you save it, windows are automatically moved to match your layout

This is perfect for:
- Quickly reorganizing your workspace
- Creating repeatable window layouts
- Moving many windows between desktops at once

### Desktop Names

Customize the names shown in desktop overlays by creating `~/.dodo/names.yaml`:

```yaml
desktop_1: Code
desktop_2: Browser
desktop_3: Communication
desktop_10: Music
```

Desktop names appear in the overlay when switching desktops, making it easier to remember what's on each desktop.

### Auto-Pinning

Automatically pin windows to all desktops based on their titles. Create `~/.dodo/always_pinned` with regex patterns (one per line):

```
.*Spotify.*
.*Discord.*
Task Manager
```

Any window whose title matches these patterns will be automatically pinned to all desktops. If you manually unpin a window using `Alt+Shift+\``, Dodo will respect that choice and won't re-pin it.

## Uninstallation

```bash
dodo --uninstall  # Remove from startup
pip uninstall dodo-desktop-switcher
```

## Troubleshooting

**Hotkeys don't work:**
- Make sure Dodo is running (check system tray)
- Check if another application is using the same keyboard shortcuts
- Try restarting Dodo

**"Failed to initialize Virtual Desktop Manager":**
- Virtual desktops must be enabled in Windows
- Try creating a virtual desktop manually first (Win+Tab, then "New Desktop")
- Restart Windows and try again

**Application doesn't start:**
- Make sure you're running Windows 10 or 11
- Check that all dependencies installed correctly: `pip install --upgrade dodo-desktop-switcher`

## License

MIT License - see LICENSE file for details.

## Author

Created by Ram Rachum (ram@rachum.com)
