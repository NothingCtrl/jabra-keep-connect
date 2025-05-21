import sys
import time
import numpy as np
import pyaudio
from pycaw.pycaw import AudioUtilities
from comtypes import CoInitialize, CoUninitialize

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

audio_sessions = None


def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)


def refresh_audio_session():
    def _refresh():
        global audio_sessions
        while True:
            CoInitialize()
            audio_sessions = AudioUtilities.GetAllSessions()
            time.sleep(60)

    threading.Thread(target=_refresh, daemon=True).start()


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
        self.intervals = [15, 300, 600, 900, 1200, 1500, 1800, 2700, 3600]
        self.selected_interval = tk.StringVar(value="1500")
        self.tray_icon = None
        self.icon_path = None
        self.tray_thread = None
        self.last_click_time = 0
        self.double_click_interval = 0.5
        self.status_label_text = tk.StringVar(value="Idle.")
        self.inaudible_tone_data = self.generate_inaudible_tone()
        self.beep_tone_data = self.generate_beep()
        self.ting_tong_data = self.generate_ting_tong()

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

    def is_any_audio_playing_pycaw(self):
        """Check any sound is playing in the system. Only on Windows"""
        if audio_sessions:
            # noinspection PyTypeChecker
            for session in audio_sessions:
                try:
                    if session.State == 1:  # State 1 is "Active"
                        return True
                except Exception as e:
                    continue
        return False

    def generate_tone(self, frequency, duration, sample_rate, amplitude, stereo: bool = False):
        num_samples = int(sample_rate * duration)
        t = np.linspace(0, duration, num_samples, False)
        note = amplitude * np.sin(2 * np.pi * frequency * t)
        if not stereo:
            return (note * 32767).astype(np.int16).tobytes()
        else:
            stereo = np.stack([note * 32767, note * 32767], axis=1)  # Duplicate for left/right channels
            return stereo.astype(np.int16).tobytes()

    def generate_inaudible_tone(self):
        sample_rate = 44100
        duration = 3
        frequency = 20000
        amplitude = 1000
        return self.generate_tone(frequency, duration, sample_rate, amplitude)

    def generate_beep(self, stereo: bool = True):
        sample_rate = 44100
        duration = 0.2  # Short duration for a quick beep
        frequency = 1000  # Audible frequency (1 kHz)
        amplitude = 600  # Moderate amplitude for a soft beep
        return self.generate_tone(frequency, duration, sample_rate, amplitude, stereo)

    def generate_ting_tong(self):
        # Tạo dữ liệu âm thanh cho một tiếng "ting"
        ting_data = self.generate_tone(100, 0.15, 44100, 0.15)
        tong_data = self.generate_tone(150, 0.15, 44100, 0.15)
        # Tạo khoảng lặng
        num_silence_samples = int(44100 * .3)
        silence = np.zeros(num_silence_samples, dtype=np.int16).tobytes()
        # Kết hợp hai tiếng "ting" với khoảng lặng ở giữa
        audio_data = silence + ting_data + silence + tong_data + silence + ting_data + silence + tong_data + silence
        return audio_data

    def play_audio(self, audio_data: bytes, channels: int = 2):
        p = pyaudio.PyAudio()
        try:
            # Check for available output devices
            device_count = p.get_device_count()
            output_device = []
            for i in range(device_count):
                device_info = p.get_device_info_by_index(i)
                if device_info['maxOutputChannels'] > 0 and 'jabra' in device_info['name'].lower():  # Output device
                    output_device.append(i)
            if not output_device:
                print("No output device found. Waiting for device...")
                return False  # Signal to retry
            # Open stream with the found device
            is_ok = False
            try:
                for d in output_device:
                    stream = p.open(format=pyaudio.paInt16,
                                    channels=channels,
                                    rate=44100,
                                    output=True,
                                    output_device_index=d)
                    stream.write(audio_data)
                    stream.stop_stream()
                    stream.close()
                    is_ok = True
                    break
            except Exception:
                pass
            if is_ok:
                return True
            else:
                return False
        except Exception as e:
            print(f"Error playing audio: {e}")
            return False
        finally:
            p.terminate()

    def playback_loop(self):
        interval = int(self.selected_interval.get())
        while self.is_running and not self.stop_event.is_set():
            if not self.is_any_audio_playing_pycaw():
                if not self.play_audio(self.ting_tong_data, channels=1):
                    self.stop_event.wait(timeout=5)  # Wait 5 seconds before retry
                    continue
            for remaining_time in range(interval, -1, -1):
                self.status_label_text.set(f"Next playback in {remaining_time} seconds.")
                if self.stop_event.is_set():
                    break
                time.sleep(1)
            if not self.stop_event.is_set():
                if not self.is_any_audio_playing_pycaw():
                    self.status_label_text.set("Playing sound...")
                else:
                    self.status_label_text.set("Audio is playing...")
                    time.sleep(1)
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
        icon = pystray.Icon("Jabra Keep Connect", Image.open(self.icon_path), "Jabra Keep Connect",
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
    refresh_audio_session()
    root = tk.Tk()
    app = KeepAliveApp(root)
    root.mainloop()
