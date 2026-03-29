import tkinter as tk
from tkinter import filedialog, messagebox
import pystray
from PIL import Image, ImageDraw
import os
import shutil
import sys
import time
import json
import threading
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import winreg as reg
import logging

# --- Delay before organizing (in seconds) ---
ORGANIZE_DELAY = 600  # 10 minutes

# Configuration handler
class Config:
    def __init__(self):
        self.config_path = os.path.join(os.getenv('APPDATA'), 'DownloadOrganizer', 'config.json')
        self.data = {
            'monitor_dir': os.path.join(os.path.expanduser('~'), 'Downloads'),
            'start_on_login': False
        }
        self.load()

    def load(self):
        try:
            with open(self.config_path, 'r') as f:
                self.data = json.load(f)
        except (FileNotFoundError, json.JSONDecodeError):
            self.save()
        except Exception as e:
            messagebox.showerror("Error", f"Config load error: {e}")

    def save(self):
        os.makedirs(os.path.dirname(self.config_path), exist_ok=True)
        with open(self.config_path, 'w') as f:
            json.dump(self.data, f)

    def set_startup(self, enable):
        self.data['start_on_login'] = enable
        self.save()
        key = reg.HKEY_CURRENT_USER
        key_path = r"Software\Microsoft\Windows\CurrentVersion\Run"
        app_name = "DownloadOrganizer"
        exe_path = f'"{sys.executable}" "{os.path.abspath(__file__)}"'
        try:
            with reg.OpenKey(key, key_path, 0, reg.KEY_ALL_ACCESS) as reg_key:
                if enable:
                    reg.SetValueEx(reg_key, app_name, 0, reg.REG_SZ, exe_path)
                else:
                    try:
                        reg.DeleteValue(reg_key, app_name)
                    except FileNotFoundError:
                        pass
        except Exception as e:
            messagebox.showerror("Error", f"Startup setting failed: {e}")


# File organization logic
class FileOrganizer:
    # Nested structure: category → sub-type → [extensions]
    # Flat entries (list value) have no sub-type folder.
    #
    # Final folder path layout:
    #   <monitor_dir>/<Category>/<Sub-type>/<Year>/<Month>/<file>   (nested)
    #   <monitor_dir>/<Category>/<Year>/<Month>/<file>              (flat)
    CATEGORIES = {
        'Images': {
            'Photos':   ['jpg', 'jpeg', 'png', 'gif', 'bmp', 'webp', 'tiff', 'raw', 'heic', 'heif'],
            'Graphics': ['svg', 'psd', 'ai', 'eps', 'xcf'],
            'Icons':    ['ico'],
        },
        'Videos': {
            'Movies': ['mp4', 'mov', 'avi', 'mkv'],
            'Clips':  ['flv', 'wmv', 'mpeg', 'mpg', '3gp', 'webm', 'web'],
        },
        'Documents': {
            'PDFs':          ['pdf'],
            'Word':          ['doc', 'docx', 'rtf', 'odt'],
            'Spreadsheets':  ['xls', 'xlsx', 'ods'],
            'Presentations': ['ppt', 'pptx', 'odp'],
            'Text':          ['txt', 'md'],
        },
        'Music': {
            'Lossless': ['flac', 'wav', 'aiff', 'alac'],
            'Lossy':    ['mp3', 'aac', 'ogg', 'wma', 'm4a'],
        },
        'Programs': {
            'Installers': ['exe', 'msi', 'pkg', 'dmg', 'deb', 'rpm'],
            'Portable':   ['jar', 'appimage'],
            'Mobile':     ['apk', 'ipa'],
            'Firmware':   ['uf2', 'iso', 'img'],
            'Scripts':    ['bat', 'sh', 'ps1'],
        },
        'Archives': {
            'Compressed': ['zip', 'rar', '7z', 'tar', 'gz', 'bz2', 'xz', 'zst'],
            'Partial':    ['part'],
        },
        '3D Models': {
            'Print Ready': ['stl', '3mf', 'gcode'],
            'Design':      ['obj', 'fbx', 'dae', 'gltf', 'blend'],
        },
        'Code': {
            'Python':  ['py'],
            'Web':     ['js', 'ts', 'html', 'css', 'php'],
            'C / C++': ['c', 'cpp', 'h', 'hpp'],
            'C#':      ['cs'],
            'Other':   ['rb', 'vb', 'go', 'rs', 'swift', 'java'],
        },
        'Fonts': {
            'TrueType': ['ttf'],
            'OpenType': ['otf'],
            'Web':      ['woff', 'woff2'],
        },
        'Data': {
            'Databases': ['db', 'sqlite', 'sql'],
            'Tabular':   ['csv', 'tsv'],
            'Config':    ['json', 'xml', 'yaml', 'yml', 'toml', 'ini'],
            'Logs':      ['log'],
        },
        'Torrents': ['torrent'],
        'System': {
            'Libraries': ['dll', 'so', 'dylib'],
            'Compiled':  ['pyd', 'pyo', 'pyc', 'bin'],
        },
        'Others': [],   # catch-all — no sub-type, no date folder
    }

    def __init__(self, base_path):
        self.base_path = base_path

    def get_category(self, filename):
        """Return (category, sub_category) or (category, None) for flat entries."""
        ext = filename.rsplit('.', 1)[-1].lower() if '.' in filename else ''
        for category, value in self.CATEGORIES.items():
            if isinstance(value, dict):
                for sub_category, exts in value.items():
                    if ext in exts:
                        return (category, sub_category)
            elif isinstance(value, list):
                if ext in value:
                    return (category, None)
        return ('Others', None)

    def _dest_dir(self, category, sub_category):
        """Build the full destination directory path by format only (no year/month)."""
        if category == 'Others':
            return os.path.join(self.base_path, 'Others')

        if sub_category:
            return os.path.join(self.base_path, category, sub_category)
        else:
            return os.path.join(self.base_path, category)

    def organize_file(self, file_path):
        if os.path.isdir(file_path):
            return
        # Wait for the file to be accessible (not still being written)
        while True:
            try:
                with open(file_path, 'rb') as f:
                    f.read(8)
                break
            except (IOError, OSError):
                time.sleep(10)

        filename = os.path.basename(file_path)
        self.wait_for_download_completion(file_path)

        category, sub_category = self.get_category(filename)
        dest_dir = self._dest_dir(category, sub_category)
        os.makedirs(dest_dir, exist_ok=True)
        dest_path = os.path.join(dest_dir, filename)

        label = f"{category}/{sub_category}" if sub_category else category

        if not os.path.exists(dest_path):
            shutil.move(file_path, dest_dir)
            logging.info(f"Moved '{filename}' → {label}")
        else:
            base, extension = os.path.splitext(filename)
            counter = 1
            while True:
                new_name = f"{base} ({counter}){extension}"
                new_path = os.path.join(dest_dir, new_name)
                if not os.path.exists(new_path):
                    shutil.move(file_path, new_path)
                    logging.info(f"Moved '{filename}' → {label} as '{new_name}'")
                    break
                counter += 1

    def wait_for_download_completion(self, file_path, check_interval=30):
        base_file_path = file_path[:-5] if file_path.lower().endswith('.part') else file_path

        while True:
            if file_path.lower().endswith('.part'):
                if not os.path.exists(file_path):
                    if os.path.exists(base_file_path):
                        logging.info(f"Download completed: '{base_file_path}' is available.")
                        break
                    else:
                        logging.warning(f"File '{file_path}' was removed or renamed unexpectedly.")
                        break
                else:
                    logging.info(f"Download still in progress: '{file_path}'")
            else:
                if os.path.exists(file_path):
                    break
                else:
                    logging.warning(f"File '{file_path}' does not exist.")
                    break
            time.sleep(check_interval)


# File system event handler — with 10-minute delay
class DownloadHandler(FileSystemEventHandler):
    def __init__(self, organizer, on_pending_change=None):
        super().__init__()
        self.organizer = organizer
        self.retries = 5
        self.delay = 2
        self._pending_timers = {}   # file_path -> threading.Timer
        self._lock = threading.Lock()
        self.on_pending_change = on_pending_change  # optional callback for UI updates

    @property
    def pending_count(self):
        with self._lock:
            return len(self._pending_timers)

    def on_created(self, event):
        if not event.is_directory:
            self._schedule(event.src_path)

    def _schedule(self, path):
        """Schedule a file to be organized after ORGANIZE_DELAY seconds."""
        with self._lock:
            # Cancel any existing timer for this path (e.g. file replaced)
            if path in self._pending_timers:
                self._pending_timers[path].cancel()

            timer = threading.Timer(ORGANIZE_DELAY, self._run, args=[path])
            self._pending_timers[path] = timer
            timer.start()
            logging.info(
                f"Queued '{os.path.basename(path)}' — will organize in "
                f"{ORGANIZE_DELAY // 60} min ({self.pending_count} pending)"
            )

        if self.on_pending_change:
            self.on_pending_change(self.pending_count)

    def _run(self, path):
        """Called after the delay; actually organizes the file."""
        with self._lock:
            self._pending_timers.pop(path, None)

        if self.on_pending_change:
            self.on_pending_change(self.pending_count)

        if not os.path.exists(path):
            logging.info(f"File no longer exists, skipping: '{path}'")
            return

        self.handle_file(path)

    def handle_file(self, path):
        for _ in range(self.retries):
            try:
                with open(path, 'rb'):
                    pass
                self.organizer.organize_file(path)
                break
            except PermissionError:
                time.sleep(self.delay)
            except Exception as e:
                logging.error(f"Error processing {path}: {e}")
                break

    def cancel_all(self):
        """Cancel all pending timers (called on shutdown)."""
        with self._lock:
            for timer in self._pending_timers.values():
                timer.cancel()
            self._pending_timers.clear()


# Main application class
class DownloadOrganizerApp:
    def __init__(self):
        self.config = Config()
        self.observer = None
        self.handler = None
        self.tray_icon = None
        self.setup_gui()
        # Auto-start monitoring immediately if the directory exists
        self._auto_start()

    def setup_gui(self):
        self.root = tk.Tk()
        self.root.title("Download Organizer")
        self.root.geometry("400x230")

        # Intercept window close → minimize to tray instead of quitting
        self.root.protocol("WM_DELETE_WINDOW", self._hide_to_tray)

        # Directory selection
        self.dir_var = tk.StringVar(value=self.config.data['monitor_dir'])
        tk.Label(self.root, text="Monitor Directory:").pack(pady=5)
        dir_frame = tk.Frame(self.root)
        dir_frame.pack(pady=5)
        tk.Entry(dir_frame, textvariable=self.dir_var, width=35).pack(side=tk.LEFT, padx=5)
        tk.Button(dir_frame, text="Browse", command=self.choose_directory).pack(side=tk.LEFT)

        # Startup checkbox
        self.startup_var = tk.BooleanVar(value=self.config.data['start_on_login'])
        tk.Checkbutton(self.root, text="Start with Windows", variable=self.startup_var,
                       command=self.update_startup).pack(pady=5)

        # Status label
        self.status_var = tk.StringVar(value="Status: Stopped")
        tk.Label(self.root, textvariable=self.status_var, fg="gray").pack(pady=2)

        # Buttons
        btn_frame = tk.Frame(self.root)
        btn_frame.pack(pady=8)
        self.start_btn = tk.Button(btn_frame, text="Start Monitoring", command=self.start_monitoring, width=16)
        self.start_btn.pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="Sort Now", command=self.sort_now, width=10).pack(side=tk.LEFT, padx=5)

    def _auto_start(self):
        """Start monitoring automatically on launch if the directory is valid."""
        monitor_dir = self.config.data.get('monitor_dir', '')
        if monitor_dir and os.path.exists(monitor_dir):
            self.start_monitoring(hide=True)

    def _hide_to_tray(self):
        """Hide the main window to the system tray (do not quit)."""
        self.root.withdraw()
        if self.tray_icon is None:
            self.create_tray_icon()

    def _update_pending(self, count):
        """Called when the pending file count changes — refresh tray tooltip."""
        if self.tray_icon:
            suffix = f" ({count} pending)" if count > 0 else ""
            self.tray_icon.title = f"Download Organizer{suffix}"

    def choose_directory(self):
        directory = filedialog.askdirectory(initialdir=self.dir_var.get())
        if directory:
            self.dir_var.set(directory)
            self.config.data['monitor_dir'] = directory
            self.config.save()

    def update_startup(self):
        self.config.set_startup(self.startup_var.get())

    def show_main_window(self):
        self.root.deiconify()
        self.root.lift()

    def start_monitoring(self, hide=False):
        if self.observer and self.observer.is_alive():
            return  # Already running

        monitor_dir = self.dir_var.get()
        if not os.path.exists(monitor_dir):
            messagebox.showerror("Error", "Invalid directory selected")
            return

        organizer = FileOrganizer(monitor_dir)
        self.handler = DownloadHandler(organizer, on_pending_change=self._update_pending)
        self.observer = Observer()
        self.observer.schedule(self.handler, monitor_dir, recursive=False)
        self.observer.start()

        self.status_var.set(f"Status: Monitoring  •  {ORGANIZE_DELAY // 60}-min delay active")
        self.start_btn.config(state=tk.DISABLED)
        logging.info(f"Started monitoring: {monitor_dir}")

        if hide:
            self.root.withdraw()
            self.create_tray_icon()

    def sort_now(self):
        monitor_dir = self.dir_var.get()
        if not os.path.exists(monitor_dir):
            messagebox.showerror("Error", "Invalid directory selected")
            return

        organizer = FileOrganizer(monitor_dir)
        handler = DownloadHandler(organizer)

        dirs_to_scan = [monitor_dir]
        others_folder = os.path.join(monitor_dir, 'Others')
        if os.path.isdir(others_folder):
            dirs_to_scan.append(others_folder)

        for directory in dirs_to_scan:
            for filename in os.listdir(directory):
                file_path = os.path.join(directory, filename)
                if os.path.isfile(file_path):
                    handler.handle_file(file_path)

        messagebox.showinfo("Info", "Files sorted successfully!")

    def create_tray_icon(self):
        if self.tray_icon is not None:
            return  # Already created

        def create_icon_image():
            icon_path = os.path.join(os.path.dirname(sys.argv[0]), 'fileorg.ico')
            if os.path.exists(icon_path):
                return Image.open(icon_path)
            image = Image.new('RGB', (64, 64), color=(0, 0, 0))
            draw = ImageDraw.Draw(image)
            draw.rectangle((0, 0, 64, 64), fill=(30, 120, 200))
            return image

        menu = pystray.Menu(
            pystray.MenuItem('Sort Now', self.sort_now),
            pystray.MenuItem('Show', self.show_main_window),
            pystray.MenuItem('Exit', self.exit_application)
        )
        self.tray_icon = pystray.Icon(
            "organizer_icon",
            create_icon_image(),
            "Download Organizer",
            menu
        )
        threading.Thread(target=self.tray_icon.run, daemon=True).start()

    def exit_application(self):
        if self.handler:
            self.handler.cancel_all()
        if self.observer:
            self.observer.stop()
            self.observer.join()
        if self.tray_icon:
            self.tray_icon.stop()
        self.root.destroy()
        sys.exit(0)


logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S',
    filename='app.log',
    filemode='a'   # append so logs survive restarts
)
logger = logging.getLogger(__name__)
logger.info('=== Download Organizer started ===')

if __name__ == "__main__":
    app = DownloadOrganizerApp()
    app.root.mainloop()
