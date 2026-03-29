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

# Partial/temp download extensions — never organize these
SKIP_EXTS = {'crdownload', 'tmp', 'downloading', 'part~'}

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
    COMPRESSED_EXTS = {'zip', 'rar', '7z', 'tar', 'gz', 'bz2', 'xz', 'zst'}

    # Directories created by the organizer that fresh_organize should not recurse into
    SKIP_DIRS = {'Open Archives', 'Archives W. Compressed'}

    # Nested structure: category → sub-type → [extensions]
    # Flat entries (list value) have no sub-type folder.
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
        'Others': [],   # catch-all — no sub-type
    }

    def __init__(self, base_path):
        self.base_path = base_path
        self._lock = threading.Lock()  # guards paired archive moves

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
        """Build the full destination directory path by format only."""
        if category == 'Others':
            return os.path.join(self.base_path, 'Others')
        if sub_category:
            return os.path.join(self.base_path, category, sub_category)
        else:
            return os.path.join(self.base_path, category)

    def _unique_dest(self, parent, name):
        """Return a collision-free path for name inside parent."""
        dest = os.path.join(parent, name)
        if not os.path.exists(dest):
            return dest
        base, ext = os.path.splitext(name)
        counter = 1
        while True:
            candidate = os.path.join(parent, f"{base} ({counter}){ext}")
            if not os.path.exists(candidate):
                return candidate
            counter += 1

    def _find_paired_archive(self, parent_dir, dirname):
        """Return the path of a compressed file matching dirname, or None."""
        for ext in self.COMPRESSED_EXTS:
            candidate = os.path.join(parent_dir, f"{dirname}.{ext}")
            if os.path.exists(candidate):
                return candidate
        return None

    # ------------------------------------------------------------------
    # Directory handling
    # ------------------------------------------------------------------

    def organize_item(self, path):
        """Unified entry point — routes to directory or file handler."""
        if not os.path.exists(path):
            return
        if os.path.isdir(path):
            self._organize_dir(path)
        else:
            self.organize_file(path)

    def _organize_dir(self, dir_path):
        """Move a directory to Open Archives or Archives W. Compressed."""
        with self._lock:
            if not os.path.exists(dir_path):
                return
            dirname = os.path.basename(dir_path)
            parent  = os.path.dirname(dir_path)
            paired  = self._find_paired_archive(parent, dirname)

            if paired and os.path.exists(paired):
                dest_parent = os.path.join(self.base_path, 'Archives W. Compressed')
                os.makedirs(dest_parent, exist_ok=True)
                dest = self._unique_dest(dest_parent, dirname)
                shutil.move(dir_path, dest)
                shutil.move(paired, os.path.join(dest, os.path.basename(paired)))
                logging.info(
                    f"Paired '{dirname}' + '{os.path.basename(paired)}'"
                    f" → Archives W. Compressed"
                )
            else:
                dest_parent = os.path.join(self.base_path, 'Open Archives')
                os.makedirs(dest_parent, exist_ok=True)
                dest = self._unique_dest(dest_parent, dirname)
                shutil.move(dir_path, dest)
                logging.info(f"'{dirname}' → Open Archives")

    # ------------------------------------------------------------------
    # File handling
    # ------------------------------------------------------------------

    def organize_file(self, file_path):
        if os.path.isdir(file_path):
            self._organize_dir(file_path)
            return

        ext = file_path.rsplit('.', 1)[-1].lower() if '.' in file_path else ''

        # Skip partial/temp downloads completely
        if ext in SKIP_EXTS:
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

        if not os.path.exists(file_path):
            return

        # If this is a compressed file and a matching directory exists,
        # let _organize_dir handle both together
        if ext in self.COMPRESSED_EXTS:
            stem = filename[:-(len(ext) + 1)]
            candidate_dir = os.path.join(os.path.dirname(file_path), stem)
            if os.path.isdir(candidate_dir):
                self._organize_dir(candidate_dir)
                return

        category, sub_category = self.get_category(filename)
        dest_dir  = self._dest_dir(category, sub_category)
        dest_path = os.path.join(dest_dir, filename)

        # Already in the right place — nothing to do
        if os.path.abspath(os.path.dirname(file_path)) == os.path.abspath(dest_dir):
            return

        os.makedirs(dest_dir, exist_ok=True)
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

    # ------------------------------------------------------------------
    # Fresh / recursive organize
    # ------------------------------------------------------------------

    def fresh_organize(self, root_dir):
        """Recursively organize all files and dirs under root_dir.

        - Skips Open Archives and Archives W. Compressed (to avoid a mess).
        - Skips hidden files/dirs (names starting with '.').
        - Recurses into known category folders; moves unknown dirs to
          Open Archives or Archives W. Compressed.
        - Files already in the correct folder are left untouched.
        """
        try:
            entries = list(os.scandir(root_dir))
        except PermissionError:
            return

        # Organize loose files at this level first
        for entry in entries:
            if entry.is_file(follow_symlinks=False) and not entry.name.startswith('.'):
                self.organize_file(entry.path)

        # Then handle subdirectories
        for entry in entries:
            if not entry.is_dir(follow_symlinks=False):
                continue
            if entry.name.startswith('.'):
                continue
            if entry.name in self.SKIP_DIRS:
                continue
            if not os.path.exists(entry.path):
                continue  # already moved (e.g. paired archive handled it)

            if entry.name in self.CATEGORIES:
                # Known category folder — recurse into it, don't move it
                self.fresh_organize(entry.path)
            else:
                # Unknown directory — treat as Open Archive / Archives W. Compressed
                self._organize_dir(entry.path)


# File system event handler — with 10-minute delay
class DownloadHandler(FileSystemEventHandler):
    def __init__(self, organizer, on_pending_change=None):
        super().__init__()
        self.organizer = organizer
        self.retries = 5
        self.delay = 2
        self._pending_timers = {}   # path -> threading.Timer
        self._lock = threading.Lock()
        self.on_pending_change = on_pending_change

    @property
    def pending_count(self):
        with self._lock:
            return len(self._pending_timers)

    def on_created(self, event):
        # Handle both new files and new directories
        self._schedule(event.src_path)

    def _schedule(self, path):
        """Schedule a path (file or dir) to be organized after ORGANIZE_DELAY seconds."""
        with self._lock:
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
        """Called after the delay; actually organizes the item."""
        with self._lock:
            self._pending_timers.pop(path, None)

        if self.on_pending_change:
            self.on_pending_change(self.pending_count)

        if not os.path.exists(path):
            logging.info(f"Item no longer exists, skipping: '{path}'")
            return

        self.handle_item(path)

    def handle_item(self, path):
        if os.path.isdir(path):
            self.organizer.organize_item(path)
            return
        for _ in range(self.retries):
            try:
                with open(path, 'rb'):
                    pass
                self.organizer.organize_item(path)
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
        self._auto_start()

    def setup_gui(self):
        self.root = tk.Tk()
        self.root.title("Download Organizer")
        self.root.geometry("400x320")

        self.root.protocol("WM_DELETE_WINDOW", self._hide_to_tray)

        # --- Monitor directory ---
        self.dir_var = tk.StringVar(value=self.config.data['monitor_dir'])
        tk.Label(self.root, text="Monitor Directory:").pack(pady=5)
        dir_frame = tk.Frame(self.root)
        dir_frame.pack(pady=5)
        tk.Entry(dir_frame, textvariable=self.dir_var, width=35).pack(side=tk.LEFT, padx=5)
        tk.Button(dir_frame, text="Browse", command=self.choose_directory).pack(side=tk.LEFT)

        # --- Startup checkbox ---
        self.startup_var = tk.BooleanVar(value=self.config.data['start_on_login'])
        tk.Checkbutton(self.root, text="Start with Windows", variable=self.startup_var,
                       command=self.update_startup).pack(pady=5)

        # --- Status ---
        self.status_var = tk.StringVar(value="Status: Stopped")
        tk.Label(self.root, textvariable=self.status_var, fg="gray").pack(pady=2)

        # --- Monitor buttons ---
        btn_frame = tk.Frame(self.root)
        btn_frame.pack(pady=5)
        self.start_btn = tk.Button(btn_frame, text="Start Monitoring",
                                   command=self.start_monitoring, width=16)
        self.start_btn.pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="Sort Now",
                  command=self.sort_now, width=10).pack(side=tk.LEFT, padx=5)

        # --- Separator ---
        tk.Frame(self.root, height=1, bg="gray").pack(fill=tk.X, padx=10, pady=6)

        # --- Fresh Organize section ---
        tk.Label(self.root, text="Fresh Organize (recursive):").pack()
        fresh_frame = tk.Frame(self.root)
        fresh_frame.pack(pady=5)
        self.fresh_dir_var = tk.StringVar(value=self.config.data['monitor_dir'])
        tk.Entry(fresh_frame, textvariable=self.fresh_dir_var, width=28).pack(side=tk.LEFT, padx=5)
        tk.Button(fresh_frame, text="Browse",
                  command=self.choose_fresh_directory).pack(side=tk.LEFT)

        self.fresh_btn = tk.Button(self.root, text="Fresh Organize",
                                   command=self.fresh_organize, width=16)
        self.fresh_btn.pack(pady=4)

    def _auto_start(self):
        monitor_dir = self.config.data.get('monitor_dir', '')
        if monitor_dir and os.path.exists(monitor_dir):
            self.start_monitoring(hide=True)

    def _hide_to_tray(self):
        self.root.withdraw()
        if self.tray_icon is None:
            self.create_tray_icon()

    def _update_pending(self, count):
        if self.tray_icon:
            suffix = f" ({count} pending)" if count > 0 else ""
            self.tray_icon.title = f"Download Organizer{suffix}"

    def choose_directory(self):
        directory = filedialog.askdirectory(initialdir=self.dir_var.get())
        if directory:
            self.dir_var.set(directory)
            self.config.data['monitor_dir'] = directory
            self.config.save()

    def choose_fresh_directory(self):
        directory = filedialog.askdirectory(initialdir=self.fresh_dir_var.get())
        if directory:
            self.fresh_dir_var.set(directory)

    def update_startup(self):
        self.config.set_startup(self.startup_var.get())

    def show_main_window(self):
        self.root.deiconify()
        self.root.lift()

    def start_monitoring(self, hide=False):
        if self.observer and self.observer.is_alive():
            return

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
                    handler.handle_item(file_path)

        messagebox.showinfo("Info", "Files sorted successfully!")

    def fresh_organize(self):
        target_dir = self.fresh_dir_var.get()
        if not os.path.exists(target_dir):
            messagebox.showerror("Error", "Invalid directory selected")
            return

        self.fresh_btn.config(state=tk.DISABLED, text="Working…")

        def run():
            try:
                monitor_dir = self.dir_var.get()
                organizer = FileOrganizer(monitor_dir)
                organizer.fresh_organize(target_dir)
                self.root.after(0, lambda: messagebox.showinfo(
                    "Fresh Organize", "Done! All files organized recursively."))
            except Exception as e:
                self.root.after(0, lambda: messagebox.showerror("Error", str(e)))
            finally:
                self.root.after(0, lambda: self.fresh_btn.config(
                    state=tk.NORMAL, text="Fresh Organize"))

        threading.Thread(target=run, daemon=True).start()

    def create_tray_icon(self):
        if self.tray_icon is not None:
            return

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
