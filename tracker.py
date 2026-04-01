import threading
import time
import os
import subprocess
from datetime import datetime

import psutil
import pystray
from PIL import Image, ImageDraw, ImageFont
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
from comtypes import CLSCTX_ALL

# ── Config ────────────────────────────────────────────────────────────────────
POLL_INTERVAL = 0.5          # seconds between volume checks
CHANGE_THRESHOLD = 1         # minimum % change to trigger a log entry
LOG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "mic_log.txt")

KNOWN_AUDIO_APPS = {
    "discord.exe", "zoom.exe", "teams.exe", "obs64.exe", "obs32.exe",
    "skype.exe", "slack.exe", "loom.exe", "webex.exe", "chrome.exe",
    "firefox.exe", "streamlabs obs.exe", "streamlabsobs.exe", "msedge.exe",
    "audiodg.exe", "sndvol.exe", "realtek hd audio manager.exe",
    "rtkaudioservice.exe", "nahimicservice.exe", "soundswitch.exe",
}

# ── Tray icon drawing ─────────────────────────────────────────────────────────

def make_icon_image(volume_pct: int) -> Image.Image:
    size = 64
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)

    if volume_pct >= 70:
        color = (46, 184, 92)    # green
    elif volume_pct >= 40:
        color = (255, 165, 0)    # orange
    else:
        color = (220, 53, 69)    # red

    draw.ellipse([2, 2, size - 2, size - 2], fill=color)

    text = f"{volume_pct}%"
    font_size = 18 if volume_pct < 100 else 15
    try:
        font = ImageFont.truetype("arial.ttf", font_size)
    except Exception:
        font = ImageFont.load_default()

    bbox = draw.textbbox((0, 0), text, font=font)
    tw = bbox[2] - bbox[0]
    th = bbox[3] - bbox[1]
    draw.text(((size - tw) / 2, (size - th) / 2), text, fill="white", font=font)

    return img


# ── Volume reading ────────────────────────────────────────────────────────────

def get_mic_volume() -> int | None:
    try:
        mic = AudioUtilities.GetMicrophone()
        if mic is None:
            return None
        interface = mic.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
        volume = interface.QueryInterface(IAudioEndpointVolume)
        scalar = volume.GetMasterVolumeLevelScalar()
        return round(scalar * 100)
    except Exception:
        return None


# ── Process snapshot ──────────────────────────────────────────────────────────

def snapshot_processes():
    procs = set()
    for p in psutil.process_iter(["name"]):
        try:
            procs.add(p.info["name"].lower())
        except Exception:
            pass
    audio_procs = sorted(p for p in procs if p in KNOWN_AUDIO_APPS)
    all_procs = sorted(procs)
    return audio_procs, all_procs


# ── Logging ───────────────────────────────────────────────────────────────────

def log_change(old_vol: int, new_vol: int, audio_procs: list, all_procs: list):
    delta = new_vol - old_vol
    sign = "+" if delta > 0 else ""
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    audio_str = ", ".join(audio_procs) if audio_procs else "none detected"
    all_str = ", ".join(all_procs)

    line = (
        f"[{timestamp}] {old_vol}% → {new_vol}%  ({sign}{delta}%)\n"
        f"  Audio apps running: {audio_str}\n"
        f"  All processes: {all_str}\n"
        f"{'─' * 60}\n"
    )

    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(line)

    print(line, end="")


# ── Monitoring thread ─────────────────────────────────────────────────────────

class VolumeMonitor:
    def __init__(self):
        self.last_vol = get_mic_volume()
        self.icon = None
        self._stop = threading.Event()

    def start(self):
        t = threading.Thread(target=self._loop, daemon=True)
        t.start()

    def _loop(self):
        while not self._stop.is_set():
            time.sleep(POLL_INTERVAL)
            current = get_mic_volume()
            if current is None:
                continue
            if self.last_vol is not None and abs(current - self.last_vol) >= CHANGE_THRESHOLD:
                old = self.last_vol
                audio_procs, all_procs = snapshot_processes()
                log_change(old, current, audio_procs, all_procs)
                self._notify(old, current, audio_procs)
                self._update_icon(current)
            self.last_vol = current

    def _notify(self, old: int, new: int, audio_procs: list):
        if self.icon is None:
            return
        delta = new - old
        sign = "+" if delta > 0 else ""
        active = ", ".join(audio_procs[:3]) if audio_procs else "unknown"
        try:
            self.icon.notify(
                f"{old}% → {new}%  ({sign}{delta}%)\nActive: {active}",
                "🎤 Mic volume changed"
            )
        except Exception:
            pass

    def _update_icon(self, vol: int):
        if self.icon is None:
            return
        try:
            self.icon.icon = make_icon_image(vol)
            self.icon.title = f"Mic: {vol}%"
        except Exception:
            pass


# ── Tray setup ────────────────────────────────────────────────────────────────

def open_log(icon, item):
    subprocess.Popen(["notepad.exe", LOG_FILE])


def build_menu(monitor: VolumeMonitor):
    def current_vol_label(item):
        vol = monitor.last_vol
        return f"Current volume: {vol}%" if vol is not None else "Current volume: unknown"

    return pystray.Menu(
        pystray.MenuItem(current_vol_label, None, enabled=False),
        pystray.Menu.SEPARATOR,
        pystray.MenuItem("Open Log File", open_log),
        pystray.MenuItem("Exit", lambda icon, item: icon.stop()),
    )


def main():
    # Write log header
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(f"\n{'=' * 60}\n")
        f.write(f"Session started: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
        f.write(f"{'=' * 60}\n")

    monitor = VolumeMonitor()
    initial_vol = monitor.last_vol if monitor.last_vol is not None else 0

    print(f"🎤 Mic Volume Tracker started. Current mic volume: {initial_vol}%")
    print(f"📄 Logging changes to: {LOG_FILE}")
    print("Watching for changes... (right-click tray icon to exit)\n")

    icon_image = make_icon_image(initial_vol)
    icon = pystray.Icon(
        "mic_tracker",
        icon_image,
        f"Mic: {initial_vol}%",
        menu=build_menu(monitor),
    )
    monitor.icon = icon
    monitor.start()
    icon.run()


if __name__ == "__main__":
    main()
