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

WINDOW_TITLE = "Mouse/Keyboard Macro (Global Hotkeys, Editable Delays, JSON, Looping)"
MACRO_FILE_VERSION = 3       # version for mixed mouse+keyboard events


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
        # If this fails, macro may still work fine
        pass


set_dpi_aware()


# ==============================
# Data model
# ==============================

@dataclass
class MacroEvent:
    """
    One event in the macro.
    type:
        - "mouse_click"
        - "key_down"
        - "key_up"
    """
    type: str
    delay_before: float
    x: Optional[int] = None
    y: Optional[int] = None
    key: Optional[str] = None  # string representation of the key


# ==============================
# Helpers for key encoding/decoding
# ==============================

def encode_key(k: keyboard.Key | keyboard.KeyCode) -> str:
    """Convert pynput key object to a string for JSON."""
    if isinstance(k, keyboard.Key):
        # special key
        return f"Key.{k.name}"
    if isinstance(k, keyboard.KeyCode):
        if k.char is not None:
            return k.char
        # fallback: use virtual key code if available
        if k.vk is not None:
            return f"KeyCode.vk.{k.vk}"
    # last resort string
    return str(k)


def decode_key(s: str) -> keyboard.Key | keyboard.KeyCode | None:
    """Convert stored string back to a pynput key, if possible."""
    # Special keys: "Key.ctrl", "Key.enter", etc.
    if s.startswith("Key."):
        name = s[4:]
        try:
            return getattr(keyboard.Key, name)
        except AttributeError:
            return None

    # Virtual key strings: "KeyCode.vk.13"
    if s.startswith("KeyCode.vk."):
        try:
            vk = int(s.split(".")[-1])
            return keyboard.KeyCode.from_vk(vk)
        except Exception:
            return None

    # Single char or simple printable
    if len(s) == 1:
        return keyboard.KeyCode.from_char(s)

    return None


# ==============================
# GUI application
# ==============================

class MouseKeyboardMacroGUI:
    """
    Mouse + keyboard macro recorder and player with JSON import/export and global hotkeys.

    Features:
    - Record left mouse clicks (absolute coordinates).
    - Record keyboard events (key_down / key_up).
    - Per-event delay_before stored and editable.
    - Add extra delay to multiple selected events.
    - Loop playback N times.
    - Global hotkeys:
        F9 : start/stop recording
        F10: play
        Esc: stop recording or cancel playback
    - Save / load macros as JSON (mouse + keyboard events).
    """

    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title(WINDOW_TITLE)

        # State
        self.is_recording: bool = False
        self.is_playing: bool = False
        self.record_start_time: Optional[float] = None
        self.last_event_time: Optional[float] = None
        self.events: List[MacroEvent] = []

        self.mouse_listener: Optional[mouse.Listener] = None
        self.kb_record_listener: Optional[keyboard.Listener] = None  # for recording
        self.kb_hotkey_listener: Optional[keyboard.Listener] = None  # for global hotkeys

        self.stop_playback: bool = False  # Esc sets this during playback
        self.kb_controller = keyboard.Controller()

        self._build_ui()
        self._bind_close()
        self._start_hotkey_listener()  # global hotkeys

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
            text="Delete Selected",
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

        # Delay edit controls (set one)
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

        # Add delay to multiple selected
        add_frame = tk.Frame(self.root)
        add_frame.pack(pady=5)

        tk.Label(add_frame, text="Add delay to selected (+seconds):").pack(side=tk.LEFT)
        self.add_delay_var = tk.StringVar(value="")
        self.add_delay_entry = tk.Entry(add_frame, textvariable=self.add_delay_var, width=10)
        self.add_delay_entry.pack(side=tk.LEFT, padx=5)
        self.add_delay_button = tk.Button(
            add_frame,
            text="Add to Selected",
            command=self.add_delay_to_selected,
            state=tk.DISABLED
        )
        self.add_delay_button.pack(side=tk.LEFT, padx=5)

        # Listbox (multi-select)
        self.listbox = tk.Listbox(self.root, width=90, height=14, selectmode=tk.EXTENDED)
        self.listbox.pack(pady=5)

        # Help text
        info = (
            "Global hotkeys (work anywhere while this app is running):\n"
            "  F9  : start/stop recording\n"
            "  F10 : play macro\n"
            "  Esc : stop recording or cancel playback\n\n"
            "During recording:\n"
            "  - Mouse left-clicks are recorded (absolute coordinates).\n"
            "  - All keyboard keys are recorded (key_down/key_up), except F9/F10/Esc.\n"
            "Delays:\n"
            "  - Each event stores 'delay_before'.\n"
            "  - You can set delay for one event or add delay to multiple selected.\n"
            "Playback:\n"
            "  - Uses recorded mouse moves/clicks and keyboard presses/releases.\n"
            "  - Speed and loop count adjust timing.\n"
            "JSON:\n"
            "  - Save/load mixed mouse+keyboard macros.\n"
        )
        self.info_label = tk.Label(self.root, text=info, justify="left")
        self.info_label.pack(pady=5)

    def _bind_close(self) -> None:
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    # ---------- Global hotkey listener (F9/F10/Esc) ----------

    def _start_hotkey_listener(self) -> None:
        """Start global keyboard listener for F9/F10/Esc."""

        def on_press(key):
            try:
                if key == keyboard.Key.f9:
                    self.root.after(0, self.toggle_recording)
                elif key == keyboard.Key.f10:
                    self.root.after(0, self.play_macro)
                elif key == keyboard.Key.esc:
                    self.root.after(0, self.on_esc)
            except Exception:
                pass

        self.kb_hotkey_listener = keyboard.Listener(on_press=on_press)
        self.kb_hotkey_listener.start()

    def _stop_hotkey_listener(self) -> None:
        if self.kb_hotkey_listener is not None:
            try:
                self.kb_hotkey_listener.stop()
            except Exception:
                pass
            self.kb_hotkey_listener = None

    # ---------- Event handlers (Tk) ----------

    def on_close(self) -> None:
        """Clean up listeners and close the window."""
        self._stop_mouse_listener()
        self._stop_kb_record_listener()
        self._stop_hotkey_listener()
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
        self.add_delay_button.config(state=tk.DISABLED)
        self.save_button.config(state=tk.DISABLED)

        self._start_mouse_listener()
        self._start_kb_record_listener()

    def stop_recording(self) -> None:
        if not self.is_recording:
            return

        self.is_recording = False
        self.record_button.config(text="Start Recording (F9)")
        self.update_status("Idle")

        self._stop_mouse_listener()
        self._stop_kb_record_listener()

        # Optional: remove the last event, which is often a stop action
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

    def _start_kb_record_listener(self) -> None:
        """Start keyboard listener for recording key_down/key_up."""

        def on_press(key):
            if not self.is_recording:
                return
            # Ignore control hotkeys so they aren't recorded
            if key in (keyboard.Key.f9, keyboard.Key.f10, keyboard.Key.esc):
                return

            self._record_key_event("key_down", key)

        def on_release(key):
            if not self.is_recording:
                return
            if key in (keyboard.Key.f9, keyboard.Key.f10, keyboard.Key.esc):
                return

            self._record_key_event("key_up", key)

        if self.kb_record_listener is not None:
            self.kb_record_listener.stop()
            self.kb_record_listener = None

        self.kb_record_listener = keyboard.Listener(on_press=on_press, on_release=on_release)
        self.kb_record_listener.start()

    def _stop_kb_record_listener(self) -> None:
        if self.kb_record_listener is not None:
            try:
                self.kb_record_listener.stop()
            except Exception:
                pass
            self.kb_record_listener = None

    def _get_delay(self) -> float:
        """Compute delay since last event."""
        now = time.time()
        if self.last_event_time is None:
            delay_before = 0.0
        else:
            delay_before = now - self.last_event_time
        self.last_event_time = now
        return delay_before

    def on_mouse_click(self, x: int, y: int, button, pressed: bool) -> None:
        """pynput mouse callback used during recording."""
        if not self.is_recording:
            return
        if button != mouse.Button.left or not pressed:
            return

        pos_x, pos_y = pyautogui.position()
        delay_before = self._get_delay()

        ev = MacroEvent(
            type="mouse_click",
            delay_before=delay_before,
            x=pos_x,
            y=pos_y,
            key=None,
        )
        self.events.append(ev)

        index = len(self.events)
        text = f"{index}: MOUSE click at ({pos_x}, {pos_y})  delay={delay_before:.3f}s"
        self.root.after(0, lambda: self.listbox.insert(tk.END, text))

    def _record_key_event(self, ev_type: str, key_obj) -> None:
        """Record a keyboard event (press or release)."""
        if not self.is_recording:
            return

        delay_before = self._get_delay()
        key_str = encode_key(key_obj)

        ev = MacroEvent(
            type=ev_type,
            delay_before=delay_before,
            x=None,
            y=None,
            key=key_str,
        )
        self.events.append(ev)

        index = len(self.events)
        text = f"{index}: {ev_type.upper()} {key_str}  delay={delay_before:.3f}s"
        self.root.after(0, lambda: self.listbox.insert(tk.END, text))

    # ---------- Playback ----------

    def play_macro(self) -> None:
        if not self.events:
            messagebox.showinfo("No macro", "There are no recorded events.")
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
        self.add_delay_button.config(state=tk.DISABLED)
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

                # Wait before event
                if total_wait > 0:
                    if total_wait <= SMALL_WAIT_THRESHOLD:
                        time.sleep(total_wait)
                    else:
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

                # Execute event
                if ev.type == "mouse_click" and ev.x is not None and ev.y is not None:
                    pyautogui.moveTo(ev.x, ev.y)
                    pyautogui.click()
                elif ev.type in ("key_down", "key_up") and ev.key is not None:
                    k = decode_key(ev.key)
                    if k is not None:
                        if ev.type == "key_down":
                            self.kb_controller.press(k)
                        else:
                            self.kb_controller.release(k)

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
            messagebox.showinfo("No selection", "Select one or more events first.")
            return

        for index in reversed(selection):
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
            messagebox.showinfo("No selection", "Select an event in the list first.")
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
        self.listbox.selection_clear(0, tk.END)
        self.listbox.selection_set(index)
        self.listbox.activate(index)

    def add_delay_to_selected(self) -> None:
        """Add a positive or negative delay to all selected events."""
        if self.is_recording or self.is_playing:
            messagebox.showwarning("Busy", "Stop recording or playback first.")
            return

        selection = self.listbox.curselection()
        if not selection:
            messagebox.showinfo("No selection", "Select one or more events first.")
            return

        add_str = self.add_delay_var.get().strip()
        if not add_str:
            messagebox.showwarning("Invalid delay", "Enter a delay in seconds (can be positive or negative).")
            return

        try:
            delta = float(add_str)
        except ValueError:
            messagebox.showwarning("Invalid delay", "Enter a valid number (e.g., 0.1 or -0.05).")
            return

        for idx in selection:
            if 0 <= idx < len(self.events):
                new_val = self.events[idx].delay_before + delta
                if new_val < 0:
                    new_val = 0.0
                self.events[idx].delay_before = new_val

        self._refresh_listbox()
        self.listbox.selection_clear(0, tk.END)
        for idx in selection:
            if 0 <= idx < self.listbox.size():
                self.listbox.selection_set(idx)
        if selection:
            self.listbox.activate(selection[0])

    def _refresh_listbox(self) -> None:
        self.listbox.delete(0, tk.END)
        for i, ev in enumerate(self.events, start=1):
            if ev.type == "mouse_click":
                txt = f"{i}: MOUSE click at ({ev.x}, {ev.y})  delay={ev.delay_before:.3f}s"
            elif ev.type == "key_down":
                txt = f"{i}: KEY DOWN {ev.key}  delay={ev.delay_before:.3f}s"
            elif ev.type == "key_up":
                txt = f"{i}: KEY UP   {ev.key}  delay={ev.delay_before:.3f}s"
            else:
                txt = f"{i}: {ev.type}  delay={ev.delay_before:.3f}s"
            self.listbox.insert(tk.END, txt)

    def _update_buttons_based_on_events(self) -> None:
        """Enable or disable buttons based on whether events exist."""
        if self.events:
            self.play_button.config(state=tk.NORMAL)
            self.clear_button.config(state=tk.NORMAL)
            self.delete_selected_button.config(state=tk.NORMAL)
            self.update_delay_button.config(state=tk.NORMAL)
            self.add_delay_button.config(state=tk.NORMAL)
            self.save_button.config(state=tk.NORMAL)
        else:
            self.play_button.config(state=tk.DISABLED)
            self.clear_button.config(state=tk.DISABLED)
            self.delete_selected_button.config(state=tk.DISABLED)
            self.update_delay_button.config(state=tk.DISABLED)
            self.add_delay_button.config(state=tk.DISABLED)
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
            data_events = []
            for ev in self.events:
                item = {
                    "type": ev.type,
                    "delay_before": ev.delay_before,
                }
                if ev.type == "mouse_click":
                    item["x"] = ev.x
                    item["y"] = ev.y
                if ev.type in ("key_down", "key_up"):
                    item["key"] = ev.key
                data_events.append(item)

            data = {
                "version": MACRO_FILE_VERSION,
                "events": data_events,
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

        path = filedialog.askopenfilename(
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

            loaded_events: List[MacroEvent] = []
            for idx, item in enumerate(events_data):
                try:
                    ev_type = item.get("type", "mouse_click")
                    delay = float(item.get("delay_before", 0.0))
                    if delay < 0:
                        delay = 0.0

                    if ev_type == "mouse_click":
                        x = int(item["x"])
                        y = int(item["y"])
                        loaded_events.append(
                            MacroEvent(
                                type="mouse_click",
                                delay_before=delay,
                                x=x,
                                y=y,
                                key=None,
                            )
                        )
                    elif ev_type in ("key_down", "key_up"):
                        key_str = str(item["key"])
                        loaded_events.append(
                            MacroEvent(
                                type=ev_type,
                                delay_before=delay,
                                x=None,
                                y=None,
                                key=key_str,
                            )
                        )
                    else:
                        # Unknown type: skip
                        continue

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
    app = MouseKeyboardMacroGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
