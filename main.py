import sys
import time
import numpy as np
import pyaudio

try:
    import tkinter as tk
    from tkinter import ttk
except ImportError:
    print("Error: tkinter is not available. Please ensure Python is installed with Tcl/Tk support.")
    print("Try reinstalling Python from python.org and ensure 'tcl/tk and IDLE' is selected.")
    sys.exit(1)
import threading
import pystray
from PIL import Image, ImageDraw, ImageFont


# noinspection PyShadowingNames
class KeepAliveApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Bluetooth Keep Alive")
        self.root.geometry("300x200")
        self.root.resizable(False, False)  # Fixed window size
        self.is_running = False
        self.thread = None
        self.icon = None

        # Interval options
        self.intervals = [5, 30, 60, 300, 900, 1800, 3600]
        self.selected_interval = tk.StringVar(value=str(self.intervals[0]))

        # GUI Elements
        tk.Label(root, text="Select Interval (seconds):").pack(pady=10)
        interval_menu = ttk.Combobox(root, textvariable=self.selected_interval,
                                     values=[str(i) for i in self.intervals], state="readonly", width=10)
        interval_menu.pack(pady=10)

        # Frame for Start/Stop buttons
        button_frame = tk.Frame(root)
        button_frame.pack(pady=10)

        self.start_button = tk.Button(button_frame, text="Start", command=self.start_playback, width=10)
        self.start_button.pack(side=tk.LEFT, padx=5)

        self.stop_button = tk.Button(button_frame, text="Stop", command=self.stop_playback, state="disabled", width=10)
        self.stop_button.pack(side=tk.LEFT, padx=5)

        tk.Button(root, text="Minimize to Tray", command=self.minimize_to_tray, width=20).pack(pady=10)

        # Handle window close
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    def generate_inaudible_tone(self):
        sample_rate = 44100
        duration = 0.5
        frequency = 20
        amplitude = 1000

        t = np.linspace(0, duration, int(sample_rate * duration), False)
        audio = amplitude * np.sin(2 * np.pi * frequency * t)
        return audio.astype(np.int16)

    def play_audio(self, audio_data):
        p = pyaudio.PyAudio()
        stream = p.open(format=pyaudio.paInt16, channels=1, rate=44100, output=True)
        stream.write(audio_data.tobytes())
        stream.stop_stream()
        stream.close()
        p.terminate()

    def playback_loop(self):
        audio_data = self.generate_inaudible_tone()
        while self.is_running:
            self.play_audio(audio_data)
            time.sleep(int(self.selected_interval.get()))

    def start_playback(self):
        if not self.is_running:
            self.is_running = True
            self.start_button.config(state="disabled")
            self.stop_button.config(state="normal")
            self.thread = threading.Thread(target=self.playback_loop, daemon=True)
            self.thread.start()
            print(f"Started playing tone every {self.selected_interval.get()} seconds.")

    def stop_playback(self):
        if self.is_running:
            self.is_running = False
            self.start_button.config(state="normal")
            self.stop_button.config(state="disabled")
            print("Stopped playback.")

    def create_tray_icon(self):
        # Create a 64x64 icon with a blue "B" on white background
        image = Image.new('RGB', (64, 64), color='white')
        draw = ImageDraw.Draw(image)
        try:
            font = ImageFont.truetype("arial.ttf", 40)
        except Exception:
            font = ImageFont.load_default()
        draw.text((16, 10), "B", font=font, fill='blue')
        return image

    def minimize_to_tray(self):
        self.root.withdraw()  # Hide window
        if not self.icon:
            image = self.create_tray_icon()
            menu = (
                pystray.MenuItem("Restore", self.restore_from_tray),
                pystray.MenuItem("Exit", self.exit_app)
            )
            self.icon = pystray.Icon("Bluetooth Keep Alive", image, "Bluetooth Keep Alive", menu)
        self.icon.run()

    def restore_from_tray(self):
        self.icon.stop()
        self.icon = None
        self.root.deiconify()  # Show window

    def exit_app(self):
        self.stop_playback()
        if self.icon:
            self.icon.stop()
        self.root.destroy()
        sys.exit(0)

    def on_closing(self):
        self.exit_app()  # Close app instead of minimizing to tray


if __name__ == "__main__":
    root = tk.Tk()
    app = KeepAliveApp(root)
    root.mainloop()
