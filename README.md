<!-- @format -->

# Wah Wah Button - Windows 11 productivity app to center and layer all open windows when pressed.

[▶ Watch a 30-second demo](https://www.youtube.com/shorts/37SKrFBWrjY)

A Windows utility that automatically arranges all open windows into a neat, cascading pyramid layout—perfect for quickly organizing a cluttered desktop.

## What It Does

**Wah Wah Button** detects all visible windows across all monitors and stacks them in a symmetric pyramid arrangement:

- **Fullscreen windows** stay in the back (never moved or resized).
- **Normal windows** are resized and positioned with:
  - Smallest windows on top, largest at the back.
  - Each layer gets progressively larger (both width and height increase).
  - Symmetric cascade: top-left and bottom-right margins expand per layer for a balanced pyramid effect.
  - Windows stay within the monitor's work area (no overflow off-screen).
- **"Always on Top" windows** are detected and handled appropriately.
- **Multi-monitor support**: Each monitor gets its own pyramid of windows.
- **Preserves Z-order**: The existing stacking order is respected; only size and position change.

### Bonus Feature

On startup, the script plays a fun sound effect (`stooges.wav`) to let you know it's running. The audio plays to completion in a background process, so it won't get cut off.

## Installation

1. **Copy the files** to any folder you like—all three files must be in the same directory:

   - `WahWahButton.ps1` – The main PowerShell script
   - `WahWahButton.bat` – Batch launcher (uses relative paths)
   - `stooges.wav` – Startup sound effect (optional but fun!)
   - `Wah Wah Button.lnk` – Pre-configured shortcut (optional)

   The scripts use relative paths, so the folder can be located anywhere on your system.

2. **Use the included shortcut** or create your own:

   **Option A: Use the included `Wah Wah Button.lnk`**

   - Copy or move `Wah Wah Button.lnk` to your desktop or taskbar.
   - Double-click to run!

   **Option B: Create a new shortcut manually**

   - Right-click on your desktop → **New** → **Shortcut**.
   - Browse to `WahWahButton.bat` in your folder.
   - Click **Next**, name it `Wah Wah Button`, and click **Finish**.
   - Right-click the shortcut → **Properties** → set **Start in** to the folder containing your files.
   - (Optional) Click **Change Icon...** to customize the icon.

3. **That's it!** The batch file automatically finds the PowerShell script and sound file in the same directory.

## Usage

### Quick Start

- **Double-click** the `Wah Wah Button` shortcut (or `WahWahButton.bat`).
- All open windows will be arranged into a pyramid on each monitor.
- Fullscreen apps (games, videos) remain untouched in the background.

### What Gets Arranged?

- **Included**: Normal application windows (browsers, editors, file explorers, UWP apps like Settings or WhatsApp, etc.).
- **Excluded**: The taskbar, desktop, minimized windows, and system UI elements.

### Keyboard Shortcut (Optional)

To trigger the script with a hotkey:

1. Right-click the shortcut → **Properties**.
2. Click in the **Shortcut key** field and press your desired key combination (e.g., `Ctrl+Alt+W`).
3. Click **OK**.

Now you can press that combo anytime to re-organize your windows.

## Features

- **Multi-monitor aware**: Each display gets its own independent pyramid.
- **Smart window detection**: Handles modern UWP apps (e.g., Windows Settings, Copilot, WhatsApp) and classic Win32 windows.
- **No overlap or clipping**: Windows are resized and clamped to stay within visible screen bounds.
- **Respects "Always on Top"**: Topmost windows are detected and kept above normal windows.
- **Z-order preservation**: The current stacking order is maintained; only positions and sizes change.
- **Lightweight & fast**: Runs in milliseconds, even with dozens of windows open.
- **Hidden execution**: The script runs silently in the background (no PowerShell window pops up).

## Troubleshooting

### The script doesn't run / nothing happens

- **Check the path**: Open `WahWahButton.bat` in a text editor and verify the path to `WahWahButton.ps1` is correct.
- **Execution Policy**: If Windows blocks the script, open PowerShell as Administrator and run:
  ```powershell
  Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
  ```
  Then try again.

### The sound doesn't play

- **Verify the WAV file**: Make sure `stooges.wav` is in the same folder as the script.
- **Check volume**: Ensure your system volume is up and the correct audio device is selected.
- **Test the WAV**: Double-click `stooges.wav` in File Explorer to confirm it plays in your default media player.
- **Check the log**: If playback fails, look for `%TEMP%\WahWahButton_SFX.log` for error details.

### Some windows aren't detected

- **Minimized windows are skipped** by design. Restore them first, then run the script.
- **System windows** (like the taskbar or desktop) are intentionally excluded.
- If a specific app isn't detected, [open an issue](https://github.com/michaelplzno/Utilities/issues) with the app name and I'll investigate.

### Windows are too small / too large

- The script calculates sizes based on your screen resolution and the number of windows.
- You can edit `WahWahButton.ps1` and adjust the `$minWidth`, `$minHeight`, and `$cascadeOffset` variables near the pyramid sizing logic to fine-tune the layout.

### I want to restore my original layout

- There's no built-in "undo" (Windows doesn't provide a snapshot API).
- **Workaround**: Use Windows' built-in window management:
  - Press `Win + Tab` to see all windows and drag them back.
  - Or press `Win + D` twice to minimize/restore all windows.

## Technical Details

- **Language**: PowerShell 5.1+ (compatible with Windows 10/11 out of the box)
- **Dependencies**: None (uses built-in Windows APIs via P/Invoke)
- **APIs Used**: `user32.dll` functions like `EnumWindows`, `GetWindowRect`, `MoveWindow`, `SetWindowPos`, `MonitorFromWindow`, etc.
- **Performance**: Typically arranges 20–30 windows across 3 monitors in under 1 second.

## Customization

Want to tweak the behavior? Open `WahWahButton.ps1` in your favorite editor and look for these sections:

- **Skip certain apps**: Edit the `$skipClasses` array to exclude specific window classes.
- **Adjust cascade spacing**: Change `$cascadeOffset` to increase/decrease the gap between layers.
- **Change minimum sizes**: Modify `$minWidth` and `$minHeight` to set a floor for how small windows can get.
- **Disable the sound**: Comment out or remove the SFX block at the top of the script.

## License

This project is licensed under the GNU General Public License v3.0 (GPL-3.0).

- You may copy, modify, and redistribute this software under the terms of the GPLv3.
- There is no warranty; see the license for details.

See the `LICENSE` file included in this repository for the full text.

## Credits

- **Author**: [michaelplzno](https://github.com/michaelplzno)
- **Inspiration**: The chaos of having 50 browser tabs across 3 monitors and needing a one-click "organize" button.
- **Sound Effect**: The classic "stooges.wav" (because why not make utility scripts fun?).

---

**Enjoy your newly organized desktop!** 🎉
