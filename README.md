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

```bash
pip install pywin32 psutil pycaw comtypes
