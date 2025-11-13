import threading
import time
from dataclasses import dataclass
from typing import List, Optional

import tkinter as tk
from tkinter import messagebox, filedialog

import json
import platform
import pyautogui
from pynput import mouse, keyboard

# ==============================
# Configuration
# ==============================

DEFAULT_SPEED = 1.3          # Default playback speed (1.0 = real-time)
DEFAULT_LOOP_COUNT = 1       # Default loop count
MIN_SLEEP_CHUNK = 0.02       # Chunk size for long waits (seconds)
SMALL_WAIT_THRESHOLD = 0.2   # Below this, use single sleep

WINDOW_TITLE = "Mouse Macro (Full Screen, Editable Delays, Looping, Esc Cancel + JSON + Global Hotkeys)"
MACRO_FILE_VERSION = 1       # For future compatibility


# ==============================
# DPI awareness (Windows)
# ==============================

def set_dpi_aware() -> None:
    """Set process DPI awareness on Windows so coordinates match at scaled displays."""
    if platform.system() != "Windows":
        return
    try:
        import ctypes
        AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2 = -4
        user32 = ctypes.windll.user32
        if hasattr(user32, "SetProcessDpiAwarenessContext"):
            user32.SetProcessDpiAwarenessContext(AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2)
        else:
            shcore = ctypes.windll.shcore
            # 2 = PROCESS_PER_MONITOR_DPI_AWARE
            shcore.SetProcessDpiAwareness(2)
    except Exception:
        # If this fails, we still proceed; macro may still work fine
        pass


set_dpi_aware()


# ==============================
# Data model
# ==============================

@dataclass
class MouseEvent:
    """One click in the macro sequence."""
    x: int
    y: int
    delay_before: float  # seconds to wait before this click


# ==============================
# GUI application
# ==============================

class MouseMacroGUI:
    """
    Mouse macro recorder and player with JSON import/export and global hotkeys.

    Features:
    - Record left clicks with absolute screen coordinates (pyautogui).
    - Record per-click delay_before and show it in a list.
    - Edit delay_before for selected click.
    - Loop playback N times.
    - Global hotkeys (work even if window is not focused):
        F9: start/stop recording
        F10: play
        Esc: stop recording or cancel playback
    - Save / load macros to/from JSON files.
    """

    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title(WINDOW_TITLE)

        # State
        self.is_recording: bool = False
        self.is_playing: bool = False
        self.record_start_time: Optional[float] = None
        self.last_event_time: Optional[float] = None
        self.events: List[MouseEvent] = []
        self.mouse_listener: Optional[mouse.Listener] = None
        self.keyboard_listener: Optional[keyboard.Listener] = None
        self.stop_playback: bool = False  # Esc sets this during playback

        self._build_ui()
        self._bind_close()
        self._start_keyboard_listener()  # global hotkeys

    # ---------- UI setup ----------

    def _build_ui(self) -> None:
        self.status_label = tk.Label(self.root, text="Status: Idle")
        self.status_label.pack(pady=5)

        # Top buttons
        btn_frame = tk.Frame(self.root)
        btn_frame.pack(pady=5)

        self.record_button = tk.Button(
            btn_frame,
            text="Start Recording (F9)",
            width=20,
            command=self.toggle_recording
        )
        self.record_button.grid(row=0, column=0, padx=5)

        self.play_button = tk.Button(
            btn_frame,
            text="Play Macro (F10)",
            width=20,
            command=self.play_macro,
            state=tk.DISABLED
        )
        self.play_button.grid(row=0, column=1, padx=5)

        self.clear_button = tk.Button(
            btn_frame,
            text="Clear All",
            width=15,
            command=self.clear_events,
            state=tk.DISABLED
        )
        self.clear_button.grid(row=0, column=2, padx=5)

        # Secondary buttons
        btn2_frame = tk.Frame(self.root)
        btn2_frame.pack(pady=5)

        self.delete_selected_button = tk.Button(
            btn2_frame,
            text="Delete Selected Click",
            width=20,
            command=self.delete_selected,
            state=tk.DISABLED
        )
        self.delete_selected_button.grid(row=0, column=0, padx=5)

        self.save_button = tk.Button(
            btn2_frame,
            text="Save Macro",
            width=15,
            command=self.save_macro,
            state=tk.DISABLED
        )
        self.save_button.grid(row=0, column=1, padx=5)

        self.load_button = tk.Button(
            btn2_frame,
            text="Load Macro",
            width=15,
            command=self.load_macro
        )
        self.load_button.grid(row=0, column=2, padx=5)

        # Settings row: speed and loop count
        settings_frame = tk.Frame(self.root)
        settings_frame.pack(pady=5)

        tk.Label(settings_frame, text="Playback speed (1.0 = normal):").grid(row=0, column=0, sticky="e")
        self.speed_var = tk.StringVar(value=str(DEFAULT_SPEED))
        self.speed_entry = tk.Entry(settings_frame, textvariable=self.speed_var, width=6)
        self.speed_entry.grid(row=0, column=1, padx=5, sticky="w")

        tk.Label(settings_frame, text="Loop count:").grid(row=0, column=2, padx=(15, 0), sticky="e")
        self.loop_var = tk.StringVar(value=str(DEFAULT_LOOP_COUNT))
        self.loop_entry = tk.Entry(settings_frame, textvariable=self.loop_var, width=6)
        self.loop_entry.grid(row=0, column=3, padx=5, sticky="w")

        # Delay edit controls
        delay_frame = tk.Frame(self.root)
        delay_frame.pack(pady=5)

        tk.Label(delay_frame, text="New delay for selected (seconds):").pack(side=tk.LEFT)
        self.delay_edit_var = tk.StringVar(value="")
        self.delay_edit_entry = tk.Entry(delay_frame, textvariable=self.delay_edit_var, width=10)
        self.delay_edit_entry.pack(side=tk.LEFT, padx=5)
        self.update_delay_button = tk.Button(
            delay_frame,
            text="Update Delay",
            command=self.update_selected_delay,
            state=tk.DISABLED
        )
        self.update_delay_button.pack(side=tk.LEFT, padx=5)

        # Listbox
        self.listbox = tk.Listbox(self.root, width=80, height=10)
        self.listbox.pack(pady=5)

        # Help text
        info = (
            "Global keys (work anywhere while this app is running):\n"
            "- F9: start/stop recording.\n"
            "- F10: play macro.\n"
            "- Esc: stop recording or cancel playback.\n\n"
            "Notes:\n"
            "- Records absolute screen positions (full screen) and per-click delay_before.\n"
            "- Last recorded click is removed automatically when you stop.\n"
            "- Select a click, enter new delay, press 'Update Delay' to change just that wait.\n"
            "- Playback speed scales all delays; loop count repeats the macro.\n"
            "- 'Save Macro' and 'Load Macro' let you export/import macros as JSON files.\n"
        )
        self.info_label = tk.Label(self.root, text=info, justify="left")
        self.info_label.pack(pady=5)

    def _bind_close(self) -> None:
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    # ---------- Global keyboard listener ----------

    def _start_keyboard_listener(self) -> None:
        """Start global keyboard listener for F9/F10/Esc."""

        def on_press(key):
            try:
                if key == keyboard.Key.f9:
                    # Schedule on main thread
                    self.root.after(0, self.toggle_recording)
                elif key == keyboard.Key.f10:
                    self.root.after(0, self.play_macro)
                elif key == keyboard.Key.esc:
                    self.root.after(0, self.on_esc)
            except Exception:
                pass

        self.keyboard_listener = keyboard.Listener(on_press=on_press)
        self.keyboard_listener.start()

    def _stop_keyboard_listener(self) -> None:
        if self.keyboard_listener is not None:
            try:
                self.keyboard_listener.stop()
            except Exception:
                pass
            self.keyboard_listener = None

    # ---------- Event handlers (Tk) ----------

    def on_close(self) -> None:
        """Clean up listeners and close the window."""
        self._stop_mouse_listener()
        self._stop_keyboard_listener()
        self.root.destroy()

    def on_esc(self, event=None) -> None:
        """Esc cancels recording or playback."""
        if self.is_playing:
            self.stop_playback = True
            self.update_status("Cancel requested...")
        if self.is_recording:
            self.stop_recording()

    # ---------- Recording ----------

    def toggle_recording(self) -> None:
        if self.is_playing:  # do not start/stop while playing
            return
        if not self.is_recording:
            self.start_recording()
        else:
            self.stop_recording()

    def start_recording(self) -> None:
        if self.is_playing:
            messagebox.showwarning("Busy", "Wait for playback to finish.")
            return

        self.is_recording = True
        self.record_start_time = time.time()
        self.last_event_time = None
        self.events.clear()
        self.listbox.delete(0, tk.END)
        self.update_status("Recording...")
        self.record_button.config(text="Stop Recording (F9)")

        self.play_button.config(state=tk.DISABLED)
        self.clear_button.config(state=tk.DISABLED)
        self.delete_selected_button.config(state=tk.DISABLED)
        self.update_delay_button.config(state=tk.DISABLED)
        self.save_button.config(state=tk.DISABLED)

        self._start_mouse_listener()

    def stop_recording(self) -> None:
        if not self.is_recording:
            return

        self.is_recording = False
        self.record_button.config(text="Start Recording (F9)")
        self.update_status("Idle")

        self._stop_mouse_listener()

        # Remove the last recorded click (usually the stop action)
        if self.events:
            self.events.pop()

        self._refresh_listbox()
        self._update_buttons_based_on_events()

    def _start_mouse_listener(self) -> None:
        """Start global mouse listener for clicks."""
        if self.mouse_listener is not None:
            self.mouse_listener.stop()
            self.mouse_listener = None

        self.mouse_listener = mouse.Listener(on_click=self.on_mouse_click)
        self.mouse_listener.start()

    def _stop_mouse_listener(self) -> None:
        """Stop global mouse listener if running."""
        if self.mouse_listener is not None:
            try:
                self.mouse_listener.stop()
            except Exception:
                pass
            self.mouse_listener = None

    def on_mouse_click(self, x: int, y: int, button, pressed: bool) -> None:
        """pynput mouse callback used during recording."""
        if not self.is_recording:
            return
        if button != mouse.Button.left or not pressed:
            return

        pos_x, pos_y = pyautogui.position()
        now = time.time()

        if self.last_event_time is None:
            delay_before = 0.0
        else:
            delay_before = now - self.last_event_time

        self.last_event_time = now

        ev = MouseEvent(x=pos_x, y=pos_y, delay_before=delay_before)
        self.events.append(ev)

        index = len(self.events)
        text = f"{index}: Left click at ({pos_x}, {pos_y})  delay={delay_before:.3f}s"
        self.root.after(0, lambda: self.listbox.insert(tk.END, text))

    # ---------- Playback ----------

    def play_macro(self) -> None:
        if not self.events:
            messagebox.showinfo("No macro", "There are no recorded clicks.")
            return
        if self.is_recording:
            messagebox.showwarning("Busy", "Stop recording before playback.")
            return
        if self.is_playing:
            return

        self.is_playing = True
        self.stop_playback = False
        self.update_status("Playing macro...")

        self.play_button.config(state=tk.DISABLED)
        self.record_button.config(state=tk.DISABLED)
        self.clear_button.config(state=tk.DISABLED)
        self.delete_selected_button.config(state=tk.DISABLED)
        self.update_delay_button.config(state=tk.DISABLED)
        self.save_button.config(state=tk.DISABLED)

        thread = threading.Thread(target=self._play_macro_thread, daemon=True)
        thread.start()

    def _play_macro_thread(self) -> None:
        """Run the playback in a background thread."""
        try:
            speed = float(self.speed_var.get())
            if speed <= 0:
                speed = 1.0
        except Exception:
            speed = 1.0

        try:
            loops = int(self.loop_var.get())
            if loops <= 0:
                loops = 1
        except Exception:
            loops = 1

        if not self.events:
            self.root.after(0, self._playback_done)
            return

        cancel = False

        for _ in range(loops):
            for ev in self.events:
                if self.stop_playback:
                    cancel = True
                    break

                total_wait = ev.delay_before / speed if ev.delay_before > 0 else 0.0

                # Short waits: a single sleep
                if total_wait > 0:
                    if total_wait <= SMALL_WAIT_THRESHOLD:
                        time.sleep(total_wait)
                    else:
                        # Longer waits: chunked sleep so Esc can break in quickly
                        remaining = total_wait
                        while remaining > 0:
                            if self.stop_playback:
                                cancel = True
                                break
                            chunk = MIN_SLEEP_CHUNK if remaining > MIN_SLEEP_CHUNK else remaining
                            time.sleep(chunk)
                            remaining -= chunk
                        if cancel:
                            break

                if self.stop_playback or cancel:
                    break

                pyautogui.moveTo(ev.x, ev.y)
                pyautogui.click()

            if cancel or self.stop_playback:
                break

        self.root.after(0, self._playback_done)

    def _playback_done(self) -> None:
        """Reset UI after playback completes or is cancelled."""
        self.is_playing = False
        self.stop_playback = False
        self.update_status("Idle")

        self.record_button.config(text="Start Recording (F9)")
        self.record_button.config(state=tk.NORMAL)
        self._update_buttons_based_on_events()

    # ---------- Editing / utility ----------

    def clear_events(self) -> None:
        if self.is_recording or self.is_playing:
            messagebox.showwarning("Busy", "Stop recording or playback first.")
            return

        self.events.clear()
        self.listbox.delete(0, tk.END)
        self.update_status("Idle")
        self._update_buttons_based_on_events()

    def delete_selected(self) -> None:
        if self.is_recording or self.is_playing:
            messagebox.showwarning("Busy", "Stop recording or playback first.")
            return

        selection = self.listbox.curselection()
        if not selection:
            messagebox.showinfo("No selection", "Select a click in the list first.")
            return

        index = selection[0]
        if 0 <= index < len(self.events):
            del self.events[index]
            self._refresh_listbox()

        self._update_buttons_based_on_events()

    def update_selected_delay(self) -> None:
        if self.is_recording or self.is_playing:
            messagebox.showwarning("Busy", "Stop recording or playback first.")
            return

        selection = self.listbox.curselection()
        if not selection:
            messagebox.showinfo("No selection", "Select a click in the list first.")
            return

        index = selection[0]
        if not (0 <= index < len(self.events)):
            return

        new_delay_str = self.delay_edit_var.get().strip()
        if not new_delay_str:
            messagebox.showwarning("Invalid delay", "Enter a delay in seconds (0 or more).")
            return

        try:
            new_delay = float(new_delay_str)
            if new_delay < 0:
                messagebox.showwarning("Invalid delay", "Delay cannot be negative.")
                return
        except ValueError:
            messagebox.showwarning("Invalid delay", "Enter a valid number of seconds.")
            return

        self.events[index].delay_before = new_delay
        self._refresh_listbox()
        self.listbox.selection_set(index)
        self.listbox.activate(index)

    def _refresh_listbox(self) -> None:
        self.listbox.delete(0, tk.END)
        for i, ev in enumerate(self.events, start=1):
            self.listbox.insert(
                tk.END,
                f"{i}: Left click at ({ev.x}, {ev.y})  delay={ev.delay_before:.3f}s"
            )

    def _update_buttons_based_on_events(self) -> None:
        """Enable or disable buttons based on whether events exist."""
        if self.events:
            self.play_button.config(state=tk.NORMAL)
            self.clear_button.config(state=tk.NORMAL)
            self.delete_selected_button.config(state=tk.NORMAL)
            self.update_delay_button.config(state=tk.NORMAL)
            self.save_button.config(state=tk.NORMAL)
        else:
            self.play_button.config(state=tk.DISABLED)
            self.clear_button.config(state=tk.DISABLED)
            self.delete_selected_button.config(state=tk.DISABLED)
            self.update_delay_button.config(state=tk.DISABLED)
            self.save_button.config(state=tk.DISABLED)

    def update_status(self, text: str) -> None:
        self.status_label.config(text=f"Status: {text}")

    # ---------- Save / load ----------

    def save_macro(self) -> None:
        if not self.events:
            messagebox.showinfo("No macro", "There is nothing to save.")
            return

        path = filedialog.asksaveasfilename(
            title="Save Macro",
            defaultextension=".json",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
        if not path:
            return

        try:
            data = {
                "version": MACRO_FILE_VERSION,
                "events": [
                    {"x": ev.x, "y": ev.y, "delay_before": ev.delay_before}
                    for ev in self.events
                ],
            }
            with open(path, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=2)
            messagebox.showinfo("Saved", f"Macro saved to:\n{path}")
        except Exception as e:
            messagebox.showerror("Error", f"Could not save macro:\n{e}")

    def load_macro(self) -> None:
        if self.is_recording or self.is_playing:
            messagebox.showwarning("Busy", "Stop recording or playback first.")
            return

        path = filedialog.asksopenfilename(
            title="Load Macro",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
        if not path:
            return

        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)

            events_data = data.get("events")
            if not isinstance(events_data, list):
                raise ValueError("Invalid macro file format (missing 'events' list).")

            loaded_events: List[MouseEvent] = []
            for idx, item in enumerate(events_data):
                try:
                    x = int(item["x"])
                    y = int(item["y"])
                    delay = float(item.get("delay_before", 0.0))
                    if delay < 0:
                        delay = 0.0
                    loaded_events.append(MouseEvent(x=x, y=y, delay_before=delay))
                except Exception as inner_e:
                    raise ValueError(f"Invalid event at index {idx}: {inner_e}") from inner_e

            self.events = loaded_events
            self._refresh_listbox()
            self._update_buttons_based_on_events()
            self.update_status("Macro loaded.")
        except Exception as e:
            messagebox.showerror("Error", f"Could not load macro:\n{e}")


# ==============================
# Entry point
# ==============================

def main() -> None:
    root = tk.Tk()
    app = MouseMacroGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
