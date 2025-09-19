# Audio Prioritization App (Windows)

A tiny Windows tool I wrote because I got tired of manually switching outputs every time my headset connected. It watches for audio devices coming and going, then sets the default output based on a priority list you control.

It has a simple PyQt5 UI for rearranging devices, and it talks to the Windows audio stack via `pycaw`. It also **depends on `svv`**, so make sure that’s installed.

---

## Why this exists.
- This app does one job: **pick the highest-priority available device** and make it default. That’s it.

---

## Features
- Watches for new/removed output devices and switches automatically.
- Drag-and-drop priority list in the UI.
- Reads/writes a plain JSON config (`audio_priority.json`).
- Lightweight: no services, no drivers.

---

## Requirements
- **Windows 10/11**
- **Python 3.9+**
- Packages:
  - `PyQt5`
  - `pycaw`
  - `svv`  ← **hard dependency**

> Heads-up: If Python isn’t running with enough rights, Windows may refuse default-device changes in some edge cases (corporate lockdowns, kiosk modes, etc.).

---

## Install
```bash
git clone https://github.com/jmorseusmc-cloud/Audio-Device-Prioritization-and-Auto-Switching-Software
cd audio-prioritization-app
pip install -r requirements.txt
# ensure svv is present (if not already pulled in)
pip install svv
