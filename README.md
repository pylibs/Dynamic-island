# Dynamic Island for Windows

A feature-rich **Dynamic Island** application for Windows, built using **Python** and **PyQt6**. It brings the dynamic, interactive widget experience inspired by macOS to the Windows desktop, featuring media controls, live lyrics, system monitoring, clipboard management, timers, and an extensible plugin architecture.

## Features

* **🎵 Media Controls & Playback:**

  * Displays current track info and album art via Windows Media Transport Controls.

  * Play, pause, skip forward/backward.

  * Interactive track progress bar (click or drag to seek).

  * **Live Lyrics:** Fetches synchronized lyrics via LRCLIB with real-time line highlighting.

* **🔊 System Volume Integration:**

  * Control system volume effortlessly by scrolling the mouse wheel directly over the island (appears on scroll).

* **📋 Clipboard Manager:**

  * Unlimited text and image history, pinning support, search functionality, and quick copying.

  * 

  **📥 Network & Downloads Monitor:**

  * Real-time network speed indicators, live speed graph (sparkline), app network activity breakdown, and completion toast notifications.

* **⏱️ Multiple Timers:**

  * Run multiple simultaneous countdown timers with visual progress bars and audio alerts.

* **🔍 Alt + Space Launcher:**

  * Quick-access command palette and launcher.

  * Features: app search, built-in calculator, currency/unit converter, translation tool, weather info, Windows file search indexer, quick timers, custom commands, and clipboard search.

* **🔌 Plugin System:**

  * Easily extend functionality using Python modules (`.py`) or directories. Supports custom UI panels, background tasks, custom launcher actions, and compact widget stats.

* **🛡️ Privacy & Screen Recording:**

  * Privacy indicators for active microphone and webcam usage.

  * Built-in screen recording utility (requires `ffmpeg`) and quick screenshot tools.

## Requirements

* **Python 3.10+** (recommended)

* Operating System: **Windows 10 / 11**

## Installation & Setup

1. **Clone or download** the project repository.

2. **Install the required Python dependencies:**
   Open your terminal (Command Prompt or PowerShell) in the project folder and run:

   ```
   pip install PyQt6 pycaw comtypes psutil requests pywin32 winrt
   
   ```

3. **Optional (for Screen Recording):**

   * Download `ffmpeg.exe` and place it in the same directory as `main.py` (or ensure it is added to your system `PATH`).

## Running the Application

To run the application normally:

```
python main.py

```

To run it silently in the background (recommended for daily use):

```
pythonw main.py

```

## Shortcuts & Controls

* **`Alt + Space`**: Open/Close the global Launcher (search apps, perform calculations, convert units/currencies, translate text, check weather, search clipboard history, etc.).

* **Mouse Wheel (over the Island)**: Adjust system volume.

* **Right-Click (on the Island)**: Open context menu to toggle autostart, access screen recording, reload plugins, or exit.

## Plugin Development & Local Plugins

* **Pre-installed Plugins:** A few ready-to-use plugins are already included directly inside the local `plugins/` directory of this repository. You can explore or modify them to see how plugins work in practice.

* **AI & Developer Guide (`plg_guide.md`):** Check out the included [`plg_guide.md`](./plg_guide.md) file! It contains comprehensive instructions, API references, and structural guidelines designed specifically to help you or an **AI Assistant** quickly generate custom plugins for this application.
