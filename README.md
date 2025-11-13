Simple Macro Recorder Python

A simple but powerful mouse macro recorder for Windows.
Supports global hotkeys, GUI editing, JSON save/load, looping, per-click delays, and Esc cancel during playback or recording.

This tool is designed for precise automation on the same machine and same screen resolution.

Features
🎯 Mouse Recording

Records absolute screen coordinates

Records delay before each click automatically

Automatically removes the last click (your stop action)

🎮 Global Hotkeys (work anywhere)

F9 — Start / stop recording

F10 — Play macro

Esc — Cancel playback or stop recording

These work even if the window is minimized or behind other apps.

📝 Macro Editing

Edit the delay before any click

Delete individual clicks

Clear the whole macro

Playback speed control (e.g., 1.5x faster)

Looping support (repeat macro N times)

💾 Save & Load

Save macros as JSON

Load macros back into the program

JSON files include:

{
  "version": 1,
  "events": [
    { "x": 123, "y": 456, "delay_before": 0.24 }
  ]
}

🖥️ Windows DPI-Aware

Works correctly on Windows with 100%, 125%, 150% DPI scaling.

📦 Installation

Install required dependencies:

pip install pyautogui pynput

▶️ Usage
1. Start the program
python mouse_macro_fullscreen.py


The GUI will open.

⏺️ Recording a Macro

Press F9 to start recording

Perform the clicks you want

Press F9 again to stop

The macro will appear in the listbox

Delays between clicks are recorded automatically.

🔁 Playing a Macro

Press F10 to play.

You can also customize:

Playback speed

Loop count

Playback speed:
    1.0 = same as recorded  
    2.0 = twice as fast  
    0.5 = half speed  

Loop count:
    Number of times the macro runs


Press Esc anytime during playback to stop immediately.

✏️ Editing Clicks

Select a click in the list

Enter a new delay

Click Update Delay

Or click Delete Selected Click

💾 Saving a Macro

Click: Save Macro

Choose a .json file to export your macro.

📂 Loading a Macro

Click: Load Macro

Choose a previously saved .json file and the macro will appear in the list.

🚫 Esc Cancel

Press Esc during:

Recording → stops recording

Playback → stops the macro immediately

Esc works even when window is not focused.

🧩 File Format

Example saved JSON macro:

{
  "version": 1,
  "events": [
    { "x": 523, "y": 742, "delay_before": 0.12 },
    { "x": 900, "y": 410, "delay_before": 0.20 }
  ]
}

⚠ Limitations

This macro uses absolute screen coordinates.

It works perfectly on the same device with the same screen layout.

If you move windows, reposition UI elements, or switch monitors, you must re-record.

🧰 Dependencies

Python 3.8+

pyautogui

pynput

tkinter (included with Python on Windows)
