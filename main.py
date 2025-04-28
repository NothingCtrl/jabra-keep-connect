import sys
import time
import numpy as np
import pyaudio

try:
    import tkinter as tk
    from tkinter import ttk
    from PIL import Image, ImageDraw, ImageFont
except ImportError:
    print("Error: tkinter and Pillow are not available. Please install them.")
    print("You can install them using: pip install tkinter pillow")
    sys.exit(1)
import threading
import pystray
import os
import tempfile
import win32gui
import win32con


def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)


class KeepAliveApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Jabra Keep Connect")
        self.root.geometry("300x160")  # Increased height to accommodate status label
        self.root.resizable(False, False)
        try:
            self.root.iconbitmap(resource_path("resources/icon.ico"))
        except tk.TclError as e:
            print(f"Error loading icon: {e}")
        self.is_running = False
        self.thread = None
        self.stop_event = threading.Event()
        self.intervals = [5, 30, 60, 300, 900, 1800, 3600]
        self.selected_interval = tk.StringVar(value="900")
        self.tray_icon = None
        self.icon_path = None
        self.tray_thread = None
        self.last_click_time = 0
        self.double_click_interval = 0.5
        self.status_label_text = tk.StringVar(value="Idle.")
        self.inaudible_tone_data = self.generate_inaudible_tone()

        # Frame for Interval Label and Combobox
        interval_frame = tk.Frame(root)
        interval_frame.pack(pady=10, padx=10, fill='x')  # Add padx for some horizontal spacing

        interval_label = tk.Label(interval_frame, text="Interval (seconds):")
        interval_label.pack(side='left')

        interval_menu = ttk.Combobox(interval_frame, textvariable=self.selected_interval,
                                     values=[str(i) for i in self.intervals], state="readonly", width=10)
        interval_menu.pack(side='left', padx=10)  # Add padx to separate label and combobox

        button_frame = tk.Frame(root)
        button_frame.pack(pady=10)

        self.start_button = tk.Button(button_frame, text="Start", command=self.start_playback, width=10)
        self.start_button.pack(side=tk.LEFT, padx=5)

        self.stop_button = tk.Button(button_frame, text="Stop", command=self.stop_playback, state="disabled", width=10)
        self.stop_button.pack(side=tk.LEFT, padx=5)

        tk.Button(root, text="Minimize to system tray", command=self.minimize_to_tray, width=20).pack(pady=5)

        self.status_label = tk.Label(root, textvariable=self.status_label_text, anchor="w", fg="#808080")
        self.status_label.pack(pady=5, fill="x", padx=10)

        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    def generate_inaudible_tone(self):
        sample_rate = 44100
        duration = 3
        frequency = 20000
        amplitude = 1000
        t = np.linspace(0, duration, int(sample_rate * duration), False)
        audio = amplitude * np.sin(2 * np.pi * frequency * t)
        return audio.astype(np.int16)

    def play_audio(self, audio_data):
        try:
            p = pyaudio.PyAudio()
            stream = p.open(format=pyaudio.paInt16, channels=1, rate=44100, output=True)
            stream.write(audio_data.tobytes())
            stream.stop_stream()
            stream.close()
            p.terminate()
        except Exception as e:
            print(f"Error playing audio: {e}")

    def playback_loop(self):
        interval = int(self.selected_interval.get())
        while self.is_running and not self.stop_event.is_set():
            self.play_audio(self.inaudible_tone_data)
            for remaining_time in range(interval, -1, -1):
                self.status_label_text.set(f"Next playback in {remaining_time} seconds.")
                if self.stop_event.is_set():
                    break
                time.sleep(1)
            if not self.stop_event.is_set():
                self.status_label_text.set("Playing sound...")
            else:
                self.status_label_text.set("Playback stopped.")

    def start_playback(self):
        if not self.is_running:
            self.is_running = True
            self.stop_event.clear()
            self.start_button.config(state="disabled")
            self.stop_button.config(state="normal")
            self.thread = threading.Thread(target=self.playback_loop, daemon=True)
            self.thread.start()
            print(f"Started audio playback every {self.selected_interval.get()} seconds.")
            self.status_label_text.set("Starting playback...")

    def stop_playback(self):
        if self.is_running:
            self.is_running = False
            self.stop_event.set()
            if self.thread and self.thread.is_alive():
                self.thread.join(timeout=1.0)
                self.thread = None
            self.start_button.config(state="normal")
            self.stop_button.config(state="disabled")
            print("Stopped audio playback.")
            self.status_label_text.set("Playback stopped.")

    def create_tray_icon(self):
        image = Image.new('RGB', (64, 64), color='white')
        draw = ImageDraw.Draw(image)
        try:
            font = ImageFont.truetype("arial.ttf", 36)
        except IOError:
            font = ImageFont.load_default()
        draw.text((10, 10), "B", font=font, fill='blue')
        with tempfile.NamedTemporaryFile(suffix='.png', delete=False) as temp_file:
            image.save(temp_file, format='PNG')
            self.icon_path = temp_file.name
        icon = pystray.Icon("Bluetooth Keep Alive", Image.open(self.icon_path), "Bluetooth Keep Alive",
                            self.create_menu(), on_click=self._on_tray_click)
        return icon

    def create_menu(self):
        return pystray.Menu(
            pystray.MenuItem("Restore", self.restore_from_tray),
            pystray.MenuItem("Quit", self.exit_app)
        )

    def _run_tray(self):
        self.tray_icon.run()

    def _on_tray_click(self, icon, item):
        if item is None:  # Clicked on the icon itself
            current_time = time.time()
            if (current_time - self.last_click_time) <= self.double_click_interval:
                self._restore_window()
                self.last_click_time = 0  # Reset to prevent immediate re-trigger
            else:
                self.last_click_time = current_time
        elif str(item) == "Restore":
            self._restore_window()
        elif str(item) == "Quit":
            self._quit_app()

    def minimize_to_tray(self):
        self.root.withdraw()
        if self.tray_icon is None:
            self.tray_icon = self.create_tray_icon()
            self.tray_thread = threading.Thread(target=self._run_tray, daemon=True)
            self.tray_thread.start()
        else:
            self.tray_icon.visible = True

    def _restore_window(self):
        if self.tray_icon:
            self.tray_icon.visible = False
        self.root.deiconify()
        self.root.attributes('-topmost', True)
        self.root.lift()
        self.root.focus_force()
        self.root.wm_deiconify()
        self.root.focus_set()
        hwnd = self.root.winfo_id()
        try:
            win32gui.SetForegroundWindow(hwnd)
        except Exception as e:
            print(f"SetForegroundWindow failed: {e}")
        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
        self.root.after(300, lambda: self.root.attributes('-topmost', False))
        self.root.after(500, lambda: self.root.focus_force())
        print("Restored window from system tray")

    def restore_from_tray(self, icon, item):
        self._restore_window()

    def _quit_app(self):
        self.stop_playback()
        if self.tray_icon:
            self.tray_icon.stop()
            if self.tray_thread and self.tray_thread.is_alive():
                self.root.after(0, self._join_tray_thread)  # Join from the main thread
        else:
            self.root.destroy()
            sys.exit(0)

        if self.icon_path and os.path.exists(self.icon_path):
            os.remove(self.icon_path)
        self.root.destroy()
        sys.exit(0)

    def _join_tray_thread(self):
        if self.tray_thread and self.tray_thread.is_alive():
            self.tray_thread.join(timeout=1.0)
        if self.icon_path and os.path.exists(self.icon_path):
            os.remove(self.icon_path)
        self.root.destroy()
        sys.exit(0)

    def exit_app(self, icon, item):
        self._quit_app()

    def on_closing(self):
        self._quit_app()


if __name__ == "__main__":
    root = tk.Tk()
    app = KeepAliveApp(root)
    root.mainloop()
