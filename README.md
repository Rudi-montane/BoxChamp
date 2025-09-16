# BoxChamp (Pre-release)

A Windows multibox manager for World of Warcraft 3.3.5a (12340). Borderless layouts, input broadcast, live HP/MP reading, and a lightweight rule-based rotations engine — all in a clean dark UI.

> ⚠️ **Pre-release**: APIs, file paths, and defaults may change.

---

## Features

- Modern **dark** UI (Qt/PySide6).
- **Borderless**/topmost window modes and simple layout tiling.
- Keyboard & **mouse broadcast** (mirror clicks while holding a modifier).
- Per-slot **Auto-Login** helper (type user/pass, enter, character select).
- Live **Memory Reader** (pymem) for player/target HP/MP %, levels, target name, combo points.
- **Combat Rotations**: ordered rules with conditions (e.g., `player_hp_percent < 50`).
- **Group Targeting**: hotkeys to target master/slaves (defaults below).
- Windows-only; DPI awareness enabled.

---

## Requirements

- **Windows 10/11**
- **Python 3.10+**
- World of Warcraft **3.3.5a (12340)** client (default process name: `Wow.exe`)

---

## Installation

```bash
# (Recommended) Create a virtual environment first
python -m venv .venv
. .venv/Scripts/activate  # on PowerShell: .venv\Scripts\Activate.ps1

# Install dependencies
python -m pip install --upgrade pip
pip install PySide6 pywin32 keyboard mouse screeninfo psutil pymem
