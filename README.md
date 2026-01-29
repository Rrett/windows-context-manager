# Windows Context Manager

A lightweight Windows utility for managing windows, monitors, and per-app audio control.

![Python](https://img.shields.io/badge/python-3.8+-blue.svg)
![Platform](https://img.shields.io/badge/platform-Windows%2010%2F11-lightgrey.svg)

## Features

- **Multi-monitor support** - Move windows between displays with one click
- **Window splitting** - Snap windows side-by-side or top/bottom
- **Per-app audio control** - Mute/unmute and adjust volume for individual applications
- **Window state management** - Maximize, minimize, restore, fullscreen, snap left/right
- **Pin windows** - Keep frequently used windows at the top of the list
- **Dark theme UI** - Modern, clean interface

## Requirements

- Windows 10 or Windows 11
- Python 3.8 or higher

## Installation

### 1. Install Python

Download and install Python from [python.org](https://python.org)

> ⚠️ **Important:** Check "Add Python to PATH" during installation

### 2. Install Dependencies

Open Command Prompt or PowerShell and run:

bash
pip install pywin32 psutil pycaw comtypes keyboard


### 3. Download

Download `windows-context.py` from this repository.

### 4. Run
`bash
python windows-context.py


## Usage

### Window Management

| Action | How To |
|--------|--------|
| Move window to monitor | Select window(s) → Choose monitor from dropdown → Click **Move** |
| Split windows horizontally | Select 2+ windows → Click **Split H** |
| Split windows vertically | Select 2+ windows → Click **Split V** |
| Focus window | Click **◉** button on window card |
| Minimize/Maximize toggle | Click **□** button on window card |
| More window options | Right-click **□** button for context menu |

### Audio Control

| Action | How To |
|--------|--------|
| Mute/Unmute app | Click **🔊** button on window card |
| Adjust app volume | Right-click and drag on **🔊** button |
| Bulk mute selected | Select windows → Click **🔇 Mute** |
| Bulk unmute selected | Select windows → Click **🔊 Unmute** |
| View audio sessions | Click **🎧 Sessions** |

### Window Organization

| Action | How To |
|--------|--------|
| Pin window to top of list | Click **📌** on window card |
| Select all windows | Click **All** |
| Clear selection | Click **None** |
| Select windows on monitor | Choose monitor → Click **Monitor** |
| Refresh window list | Click **↻** |

### Window State Menu (Right-click □)

- **Maximize** - Maximize the window
- **Minimize** - Minimize the window
- **Restore** - Restore to normal size
- **Fullscreen** - Borderless fullscreen on current monitor
- **Split Left** - Snap to left half of monitor
- **Split Right** - Snap to right half of monitor

## Keyboard Shortcuts

| Shortcut | Action |
|----------|--------|
| Scroll wheel | Scroll through window list |

## Settings

Settings are automatically saved to `window_manager_settings.json`:

- Debug mode state
- Verbose logging preference
- Pinned window preferences

## Troubleshooting

### "pycaw not installed" warning
Audio control requires pycaw. Install it with:
bash
pip install pycaw comtypes


### Windows not appearing in list
- Hidden or minimized windows may not appear
- System windows and tool windows are filtered out
- Click **↻** to refresh the list

### Audio control not working for an app
- The app must be actively playing audio to appear in audio sessions
- Some apps use separate audio processes
- Click **🎧 Sessions** to see all active audio sessions

## Hotkey

Default hotkey: `Ctrl+Win+M` - Toggle app visibility (minimize/restore)

Configure via the ⌨ button in the header.

## Dependencies

| Package | Purpose |
|---------|---------|
| `pywin32` | Windows API access |
| `psutil` | Process information |
| `pycaw` | Audio session control |
| `comtypes` | COM interface support |
| `keyboard` | Global hotkey support |

## License

MIT License - Feel free to use and modify.

## Contributing

Contributions welcome! Feel free to submit issues and pull requests.
