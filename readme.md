# Game Inventory Sorter

Automatically sorts your in-game Library inventory.

**(MIT License - Copyright (c) [2025] [Boerkie])**

## Quick Start Guide

**1. Setup (One-Time Only):**

   *   **Install Python:** Go to [python.org](https://www.python.org/downloads/), download Python (3.7+). **During install, check "Add Python to PATH".**
   *   **Install Tesseract OCR:** This helps read item counts.
        *   **Windows:** Get installer from [UB Mannheim Tesseract](https://github.com/UB-Mannheim/tesseract/wiki). During install, include "Language data" (for English). Note the install path (e.g., `C:\Program Files\Tesseract-OCR`).
   *   **Download Files:** Get `inventory_sorter.py` and `config.ini` from this page. Put them in the same folder.
   *   **Install Script Needs:** Open Command Prompt (search "cmd") or Terminal. Type and run:
        ```bash
        pip install pyautogui Pillow keyboard pygetwindow numpy pytesseract
        ```
   *   **Edit `config.ini` (Briefly!):**
        *   Open `config.ini` with Notepad.
        *   `GameWindowTitle`: Change to your game's exact window title.
        *   `TesseractCmdPath` (Windows): Set to where Tesseract installed (e.g., `C:\Program Files\Tesseract-OCR\tesseract.exe`).
        *   Save and close `config.ini`. Other settings will be calibrated in-script.

**2. Run & Calibrate (First time using with a game/UI):**

   *   **Start Script:** In Command Prompt/Terminal, go to the script's folder (e.g., `cd C:\MySorter`) and run `python inventory_sorter.py`.
   *   **Start Game:** Open your game to the inventory screen you want to sort.
   *   **Calibrate UI Layout (Numpad 4):**
        1.  With game inventory open, press **Numpad 4** (or your configured key).
        2.  The script console will now guide you step-by-step. **Follow the 8 steps carefully:** hover your mouse as instructed (e.g., "Top-Left of FIRST slot") and press **Numpad 4** again for each point.
        3.  After Step 8, console says "Calibration Complete!".
   *   **Save Calibration (Numpad 9):** Press **Numpad 9** to save these UI measurements to `config.ini`.

**3. Sort Your Inventory!**

   *   **Game Ready:** Inventory open, game window active.
   *   **Plan Sort (Numpad 1):** Press Numpad 1. The script console shows what it found and the planned moves.
        *   *Quick check: Does it look right? If not, you might need to re-calibrate (Step 2) or adjust `config.ini` values like OCR `ThresholdValue`.*
   *   **Execute Sort (Numpad 2):** If the plan is okay, press Numpad 2. **Don't touch your mouse/keyboard!**
   *   **Exit Script (Numpad 0):** Press Numpad 0 when done.

**Hotkeys (Defaults - Check `config.ini`):**
*   `Numpad 1`: Calculate Sort Plan
*   `Numpad 2`: Execute Sort
*   `Numpad 4`: Start Full UI Calibration
*   `Numpad 9`: Save Calibrated Settings
*   `Numpad 0`: Exit Sorter
*   *(Optional)* `Numpad 3, 5, 6, 7, 8`: Individual fine-tuning calibrations (see `config.ini`).

**Troubleshooting:**
*   **Not working?** Re-do calibration carefully. Check `config.ini` values.
*   **Numbers not read?** Adjust `[OCR]` `ThresholdValue` in `config.ini` (try 120-220). Check debug images in `execution_debug_images` folder.

Happy Sorting!
