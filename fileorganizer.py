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
import ctypes
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import winreg as reg
import logging
import re
import hashlib

# --- Delay before organizing (in seconds) ---
ORGANIZE_DELAY = 600  # 10 minutes

# Max retries for file accessibility (10s each = 5 minutes)
MAX_ACCESS_RETRIES = 30

# Max iterations waiting for .part download (30s each = ~60 minutes)
MAX_DOWNLOAD_WAIT = 120

# Partial/temp download extensions — never organize these
SKIP_EXTS = {'crdownload', 'tmp', 'downloading', 'part~'}

# OS-generated junk files — skip silently
SKIP_FILES = {'desktop.ini', 'thumbs.db', '.ds_store'}

# ---------------------------------------------------------------------------
# Theme definitions
# ---------------------------------------------------------------------------
THEME_DARK = {
    'bg':              '#1e1e1e',
    'fg':              '#d4d4d4',
    'entry_bg':        '#2d2d2d',
    'entry_fg':        '#d4d4d4',
    'btn_bg':          '#333333',
    'btn_fg':          '#d4d4d4',
    'btn_active_bg':   '#444444',
    'check_bg':        '#1e1e1e',
    'check_fg':        '#d4d4d4',
    'check_select':    '#1e1e1e',
    'labelframe_fg':   '#888888',
    'status_fg':       '#808080',
}

THEME_LIGHT = {
    'bg':              '#f0f0f0',
    'fg':              '#000000',
    'entry_bg':        '#ffffff',
    'entry_fg':        '#000000',
    'btn_bg':          '#e1e1e1',
    'btn_fg':          '#000000',
    'btn_active_bg':   '#cccccc',
    'check_bg':        '#f0f0f0',
    'check_fg':        '#000000',
    'check_select':    '#f0f0f0',
    'labelframe_fg':   '#000000',
    'status_fg':       '#808080',
}


# ---------------------------------------------------------------------------
# Configuration handler
# ---------------------------------------------------------------------------
class Config:
    def __init__(self):
        self.config_dir = os.path.join(os.getenv('APPDATA'), 'DownloadOrganizer')
        self.config_path = os.path.join(self.config_dir, 'config.json')
        self.data = {
            'monitor_dir': os.path.join(os.path.expanduser('~'), 'Downloads'),
            'start_on_login': False,
            'dark_mode': True,
        }
        self.load()

    def load(self):
        try:
            with open(self.config_path, 'r') as f:
                saved = json.load(f)
                self.data.update(saved)
        except (FileNotFoundError, json.JSONDecodeError):
            self.save()
        except Exception as e:
            messagebox.showerror("Error", f"Config load error: {e}")

    def save(self):
        os.makedirs(self.config_dir, exist_ok=True)
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


# ---------------------------------------------------------------------------
# File organization logic
# ---------------------------------------------------------------------------
class FileOrganizer:
    COMPRESSED_EXTS = {'zip', 'rar', '7z', 'tar', 'gz', 'bz2', 'xz', 'zst'}

    # Directories created by the organizer that should not be touched
    SKIP_DIRS = {'Open Archives', 'Archives W. Compressed'}

    # Matches "name (1).ext", "name (2).ext", etc.
    COPY_PATTERN = re.compile(r'^(.+?)\s*\((\d+)\)(\.[^.]+)$')

    # Category -> sub-type -> [extensions]
    # Flat entries (list) have no sub-type folder.
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
        'Others': [],   # catch-all
    }

    # Pre-built lookup tables (populated once, used for every file)
    _EXT_MAP = None       # ext -> (category, sub_category)
    _KNOWN_DIRS = None    # set of all folder names the organizer creates

    @classmethod
    def _get_ext_map(cls):
        if cls._EXT_MAP is None:
            m = {}
            for cat, value in cls.CATEGORIES.items():
                if isinstance(value, dict):
                    for sub, exts in value.items():
                        for e in exts:
                            m[e] = (cat, sub)
                elif isinstance(value, list):
                    for e in value:
                        m[e] = (cat, None)
            cls._EXT_MAP = m
        return cls._EXT_MAP

    @classmethod
    def _get_known_dirs(cls):
        """All directory names the organizer creates — never move these."""
        if cls._KNOWN_DIRS is None:
            dirs = set(cls.CATEGORIES.keys()) | cls.SKIP_DIRS
            for value in cls.CATEGORIES.values():
                if isinstance(value, dict):
                    dirs.update(value.keys())
            cls._KNOWN_DIRS = dirs
        return cls._KNOWN_DIRS

    def __init__(self, base_path):
        self.base_path = base_path
        self._lock = threading.Lock()

    def get_category(self, filename):
        """Return (category, sub_category) or ('Others', None)."""
        ext = filename.rsplit('.', 1)[-1].lower() if '.' in filename else ''
        return self._get_ext_map().get(ext, ('Others', None))

    def _dest_dir(self, category, sub_category):
        if category == 'Others':
            return os.path.join(self.base_path, 'Others')
        if sub_category:
            return os.path.join(self.base_path, category, sub_category)
        return os.path.join(self.base_path, category)

    def _unique_dest(self, parent, name):
        """Return a collision-free path for *name* inside *parent*."""
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

            # Never move known category / sub-category / archive folders
            if dirname in self._get_known_dirs():
                return

            # Skip empty directories
            try:
                if not any(os.scandir(dir_path)):
                    return
            except PermissionError:
                return

            parent = os.path.dirname(dir_path)
            paired = self._find_paired_archive(parent, dirname)

            if paired and os.path.exists(paired):
                dest_parent = os.path.join(self.base_path, 'Archives W. Compressed')
                os.makedirs(dest_parent, exist_ok=True)
                dest = self._unique_dest(dest_parent, dirname)
                shutil.move(dir_path, dest)
                shutil.move(paired, os.path.join(dest, os.path.basename(paired)))
                logging.info(
                    f"Paired '{dirname}' + '{os.path.basename(paired)}'"
                    f" -> Archives W. Compressed"
                )
            else:
                dest_parent = os.path.join(self.base_path, 'Open Archives')
                os.makedirs(dest_parent, exist_ok=True)
                dest = self._unique_dest(dest_parent, dirname)
                shutil.move(dir_path, dest)
                logging.info(f"'{dirname}' -> Open Archives")

    # ------------------------------------------------------------------
    # File handling
    # ------------------------------------------------------------------

    def organize_file(self, file_path):
        if os.path.isdir(file_path):
            self._organize_dir(file_path)
            return

        filename = os.path.basename(file_path)

        # Skip OS junk and temp downloads
        if filename.lower() in SKIP_FILES:
            return
        ext = filename.rsplit('.', 1)[-1].lower() if '.' in filename else ''
        if ext in SKIP_EXTS:
            return

        # Wait for the file to be accessible (max ~5 min)
        for _ in range(MAX_ACCESS_RETRIES):
            try:
                with open(file_path, 'rb') as f:
                    f.read(8)
                break
            except (IOError, OSError):
                time.sleep(10)
        else:
            logging.warning(f"File inaccessible after retries: '{filename}'")
            return

        self._wait_for_download(file_path)

        if not os.path.exists(file_path):
            return

        # Compressed file with a matching directory -> pair them
        if ext in self.COMPRESSED_EXTS:
            stem = filename[:-(len(ext) + 1)]
            candidate_dir = os.path.join(os.path.dirname(file_path), stem)
            if os.path.isdir(candidate_dir):
                self._organize_dir(candidate_dir)
                return

        category, sub_category = self.get_category(filename)
        dest_dir = self._dest_dir(category, sub_category)

        # Already in the right place
        if os.path.abspath(os.path.dirname(file_path)) == os.path.abspath(dest_dir):
            return

        os.makedirs(dest_dir, exist_ok=True)
        dest_path = os.path.join(dest_dir, filename)
        label = f"{category}/{sub_category}" if sub_category else category

        if not os.path.exists(dest_path):
            shutil.move(file_path, dest_dir)
            logging.info(f"Moved '{filename}' -> {label}")
        else:
            base, extension = os.path.splitext(filename)
            counter = 1
            while True:
                new_name = f"{base} ({counter}){extension}"
                new_path = os.path.join(dest_dir, new_name)
                if not os.path.exists(new_path):
                    shutil.move(file_path, new_path)
                    logging.info(f"Moved '{filename}' -> {label} as '{new_name}'")
                    break
                counter += 1

    def _wait_for_download(self, file_path, interval=30):
        """Block until a .part file finishes downloading (with timeout)."""
        if not file_path.lower().endswith('.part'):
            return
        base_path = file_path[:-5]
        for _ in range(MAX_DOWNLOAD_WAIT):
            if not os.path.exists(file_path):
                if os.path.exists(base_path):
                    logging.info(f"Download completed: '{os.path.basename(base_path)}'")
                else:
                    logging.warning(f"'{os.path.basename(file_path)}' removed unexpectedly.")
                return
            time.sleep(interval)
        logging.warning(f"Download wait timed out: '{os.path.basename(file_path)}'")

    # ------------------------------------------------------------------
    # Fresh / recursive organize
    # ------------------------------------------------------------------

    def fresh_organize(self, root_dir):
        """Recursively organize all files and dirs under root_dir.

        Skips Open Archives, Archives W. Compressed, hidden items,
        and known category/sub-category folders (recurses into them instead).
        """
        try:
            entries = list(os.scandir(root_dir))
        except PermissionError:
            return

        known = self._get_known_dirs()

        # Files first
        for entry in entries:
            if entry.is_file(follow_symlinks=False) and not entry.name.startswith('.'):
                self.organize_file(entry.path)

        # Directories
        for entry in entries:
            if not entry.is_dir(follow_symlinks=False):
                continue
            if entry.name.startswith('.'):
                continue
            if entry.name in self.SKIP_DIRS:
                continue
            if not os.path.exists(entry.path):
                continue

            if entry.name in known:
                self.fresh_organize(entry.path)
            else:
                self._organize_dir(entry.path)

    # ------------------------------------------------------------------
    # Duplicate removal
    # ------------------------------------------------------------------

    def find_duplicates(self, root_dir):
        """Recursively find duplicate files.

        A file is a duplicate when its name matches the copy pattern
        ``name (N).ext``, a same-name original exists in the same dir,
        and both share the same size and partial content hash.
        """
        groups = {}  # (dir, base, ext) -> [(path, size, is_copy)]

        for dirpath, dirnames, filenames in os.walk(root_dir):
            dirnames[:] = [d for d in dirnames if d not in self.SKIP_DIRS]

            for fname in filenames:
                fpath = os.path.join(dirpath, fname)
                try:
                    size = os.path.getsize(fpath)
                except OSError:
                    continue

                match = self.COPY_PATTERN.match(fname)
                if match:
                    base_name = match.group(1).strip()
                    ext       = match.group(3)
                    is_copy   = True
                else:
                    base_name, ext = os.path.splitext(fname)
                    is_copy = False

                key = (dirpath, base_name, ext.lower())
                groups.setdefault(key, []).append((fpath, size, is_copy))

        to_remove = []
        for _key, files in groups.items():
            if len(files) < 2:
                continue

            by_size = {}
            for fpath, size, is_copy in files:
                by_size.setdefault(size, []).append((fpath, is_copy))

            for _size, same_size in by_size.items():
                if len(same_size) < 2:
                    continue

                by_hash = {}
                for fpath, is_copy in same_size:
                    h = self._file_hash(fpath)
                    if h is not None:
                        by_hash.setdefault(h, []).append((fpath, is_copy))

                for _h, confirmed in by_hash.items():
                    if len(confirmed) < 2:
                        continue

                    originals = [f for f in confirmed if not f[1]]
                    copies    = [f for f in confirmed if f[1]]

                    if originals:
                        to_remove.extend(f[0] for f in copies)
                    else:
                        copies.sort(key=lambda f: f[0])
                        to_remove.extend(f[0] for f in copies[1:])

        return to_remove

    @staticmethod
    def _file_hash(path, chunk_size=8192):
        """Quick hash: first + last 8 KB."""
        h = hashlib.md5()
        try:
            size = os.path.getsize(path)
            with open(path, 'rb') as f:
                h.update(f.read(chunk_size))
                if size > chunk_size * 2:
                    f.seek(-chunk_size, 2)
                    h.update(f.read(chunk_size))
        except (IOError, OSError):
            return None
        return h.hexdigest()

    def remove_duplicates(self, root_dir):
        """Find and delete duplicate files. Returns (removed, total)."""
        to_remove = self.find_duplicates(root_dir)
        removed = 0
        for path in to_remove:
            try:
                os.remove(path)
                logging.info(f"Removed duplicate: '{path}'")
                removed += 1
            except OSError as e:
                logging.error(f"Failed to remove '{path}': {e}")
        return removed, len(to_remove)


# ---------------------------------------------------------------------------
# File system event handler (with configurable delay)
# ---------------------------------------------------------------------------
class DownloadHandler(FileSystemEventHandler):
    def __init__(self, organizer, on_pending_change=None):
        super().__init__()
        self.organizer = organizer
        self.retries = 5
        self.delay = 2
        self._pending_timers = {}
        self._lock = threading.Lock()
        self.on_pending_change = on_pending_change

    @property
    def pending_count(self):
        with self._lock:
            return len(self._pending_timers)

    def on_created(self, event):
        self._schedule(event.src_path)

    def _schedule(self, path):
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
        with self._lock:
            for timer in self._pending_timers.values():
                timer.cancel()
            self._pending_timers.clear()


# ---------------------------------------------------------------------------
# GUI application
# ---------------------------------------------------------------------------
class DownloadOrganizerApp:
    def __init__(self):
        self.config = Config()
        self.observer = None
        self.handler = None
        self.tray_icon = None
        self._theme = None
        self.setup_gui()
        self._apply_theme(THEME_DARK if self.config.data.get('dark_mode', True) else THEME_LIGHT)
        self._set_dark_title_bar(self.config.data.get('dark_mode', True))
        self._auto_start()

    # ── Theme ──────────────────────────────────────────────────────

    def _apply_theme(self, theme):
        self._theme = theme
        self.root.configure(bg=theme['bg'])
        self._apply_to_children(self.root, theme)

    def _apply_to_children(self, widget, theme):
        for child in widget.winfo_children():
            cls = child.winfo_class()
            try:
                if cls in ('Frame', 'Toplevel'):
                    child.configure(bg=theme['bg'])
                elif cls == 'Label':
                    fg = theme['status_fg'] if child is self._status_label else theme['fg']
                    child.configure(bg=theme['bg'], fg=fg)
                elif cls == 'Entry':
                    child.configure(bg=theme['entry_bg'], fg=theme['entry_fg'],
                                    insertbackground=theme['entry_fg'],
                                    disabledbackground=theme['entry_bg'])
                elif cls == 'Button':
                    child.configure(bg=theme['btn_bg'], fg=theme['btn_fg'],
                                    activebackground=theme['btn_active_bg'],
                                    activeforeground=theme['btn_fg'])
                elif cls == 'Checkbutton':
                    child.configure(bg=theme['check_bg'], fg=theme['check_fg'],
                                    selectcolor=theme['check_select'],
                                    activebackground=theme['check_bg'],
                                    activeforeground=theme['check_fg'])
                elif cls == 'Labelframe':
                    child.configure(bg=theme['bg'], fg=theme['labelframe_fg'])
            except tk.TclError:
                pass
            self._apply_to_children(child, theme)

    def _set_dark_title_bar(self, dark=True):
        try:
            hwnd = ctypes.windll.user32.GetParent(self.root.winfo_id())
            DWMWA_USE_IMMERSIVE_DARK_MODE = 20
            ctypes.windll.dwmapi.DwmSetWindowAttribute(
                hwnd, DWMWA_USE_IMMERSIVE_DARK_MODE,
                ctypes.byref(ctypes.c_int(1 if dark else 0)),
                ctypes.sizeof(ctypes.c_int))
        except Exception:
            pass

    def _toggle_dark_mode(self):
        dark = self.dark_var.get()
        self.config.data['dark_mode'] = dark
        self.config.save()
        self._apply_theme(THEME_DARK if dark else THEME_LIGHT)
        self._set_dark_title_bar(dark)

    # ── GUI setup ──────────────────────────────────────────────────

    def setup_gui(self):
        self.root = tk.Tk()
        self.root.title("Download Organizer")
        self.root.geometry("420x410")
        self.root.resizable(False, False)
        self.root.protocol("WM_DELETE_WINDOW", self._hide_to_tray)

        # ── Monitoring ─────────────────────────────────────────
        monitor_frame = tk.LabelFrame(self.root, text="Monitoring", padx=8, pady=6)
        monitor_frame.pack(fill=tk.X, padx=10, pady=(8, 4))

        dir_row = tk.Frame(monitor_frame)
        dir_row.pack(fill=tk.X, pady=2)
        self.dir_var = tk.StringVar(value=self.config.data['monitor_dir'])
        tk.Entry(dir_row, textvariable=self.dir_var, width=35).pack(side=tk.LEFT, padx=(0, 5))
        tk.Button(dir_row, text="Browse", command=self.choose_directory).pack(side=tk.LEFT)

        opts_row = tk.Frame(monitor_frame)
        opts_row.pack(fill=tk.X, pady=2)
        self.startup_var = tk.BooleanVar(value=self.config.data.get('start_on_login', False))
        tk.Checkbutton(opts_row, text="Start with Windows",
                       variable=self.startup_var,
                       command=self.update_startup).pack(side=tk.LEFT)
        self.dark_var = tk.BooleanVar(value=self.config.data.get('dark_mode', True))
        tk.Checkbutton(opts_row, text="Dark Mode",
                       variable=self.dark_var,
                       command=self._toggle_dark_mode).pack(side=tk.LEFT, padx=(12, 0))

        self.status_var = tk.StringVar(value="Status: Stopped")
        self._status_label = tk.Label(monitor_frame, textvariable=self.status_var)
        self._status_label.pack(anchor=tk.W)

        btn_row = tk.Frame(monitor_frame)
        btn_row.pack(fill=tk.X, pady=(4, 0))
        self.start_btn = tk.Button(btn_row, text="Start Monitoring",
                                   command=self.start_monitoring, width=16)
        self.start_btn.pack(side=tk.LEFT, padx=(0, 5))
        tk.Button(btn_row, text="Sort Now",
                  command=self.sort_now, width=10).pack(side=tk.LEFT)

        # ── Tools ──────────────────────────────────────────────
        tools_frame = tk.LabelFrame(self.root, text="Tools", padx=8, pady=6)
        tools_frame.pack(fill=tk.X, padx=10, pady=4)

        dir_row2 = tk.Frame(tools_frame)
        dir_row2.pack(fill=tk.X, pady=2)
        tk.Label(dir_row2, text="Target:").pack(side=tk.LEFT)
        self.tools_dir_var = tk.StringVar(value=self.config.data['monitor_dir'])
        tk.Entry(dir_row2, textvariable=self.tools_dir_var, width=28).pack(side=tk.LEFT, padx=5)
        tk.Button(dir_row2, text="Browse",
                  command=self.choose_tools_directory).pack(side=tk.LEFT)

        tool_btn_row = tk.Frame(tools_frame)
        tool_btn_row.pack(fill=tk.X, pady=(6, 0))
        self.fresh_btn = tk.Button(tool_btn_row, text="Fresh Organize",
                                   command=self.fresh_organize, width=16)
        self.fresh_btn.pack(side=tk.LEFT, padx=(0, 5))
        self.dedup_btn = tk.Button(tool_btn_row, text="Remove Duplicates",
                                   command=self.remove_duplicates, width=18)
        self.dedup_btn.pack(side=tk.LEFT)

    # ── Lifecycle ──────────────────────────────────────────────

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

    # ── Directory pickers ──────────────────────────────────────

    def choose_directory(self):
        directory = filedialog.askdirectory(initialdir=self.dir_var.get())
        if directory:
            self.dir_var.set(directory)
            self.config.data['monitor_dir'] = directory
            self.config.save()

    def choose_tools_directory(self):
        directory = filedialog.askdirectory(initialdir=self.tools_dir_var.get())
        if directory:
            self.tools_dir_var.set(directory)

    def update_startup(self):
        self.config.set_startup(self.startup_var.get())

    def show_main_window(self):
        self.root.deiconify()
        self.root.lift()

    # ── Monitoring ─────────────────────────────────────────────

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

        self.status_var.set(f"Status: Monitoring  \u2022  {ORGANIZE_DELAY // 60}-min delay active")
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
            for entry_name in os.listdir(directory):
                handler.handle_item(os.path.join(directory, entry_name))

        messagebox.showinfo("Info", "Files sorted successfully!")

    # ── Tools ──────────────────────────────────────────────────

    def _tools_target(self):
        target = self.tools_dir_var.get()
        if not os.path.exists(target):
            messagebox.showerror("Error", "Invalid target directory")
            return None
        return target

    def fresh_organize(self):
        target_dir = self._tools_target()
        if not target_dir:
            return
        self.fresh_btn.config(state=tk.DISABLED, text="Working...")

        def run():
            try:
                organizer = FileOrganizer(self.dir_var.get())
                organizer.fresh_organize(target_dir)
                self.root.after(0, lambda: messagebox.showinfo(
                    "Fresh Organize", "Done! All files organized recursively."))
            except Exception as e:
                self.root.after(0, lambda msg=str(e): messagebox.showerror("Error", msg))
            finally:
                self.root.after(0, lambda: self.fresh_btn.config(
                    state=tk.NORMAL, text="Fresh Organize"))

        threading.Thread(target=run, daemon=True).start()

    def remove_duplicates(self):
        target_dir = self._tools_target()
        if not target_dir:
            return

        organizer = FileOrganizer(self.dir_var.get())
        to_remove = organizer.find_duplicates(target_dir)

        if not to_remove:
            messagebox.showinfo("Remove Duplicates", "No duplicates found.")
            return

        preview = [f"{len(to_remove)} duplicate(s) found:\n"]
        for path in to_remove[:20]:
            preview.append(f"  {os.path.relpath(path, target_dir)}")
        if len(to_remove) > 20:
            preview.append(f"  ... and {len(to_remove) - 20} more")
        preview.append("\nDelete these files?")

        if not messagebox.askyesno("Remove Duplicates", "\n".join(preview)):
            return

        self.dedup_btn.config(state=tk.DISABLED, text="Working...")

        def run():
            try:
                removed, total = organizer.remove_duplicates(target_dir)
                self.root.after(0, lambda: messagebox.showinfo(
                    "Remove Duplicates",
                    f"Removed {removed} of {total} duplicate(s)."))
            except Exception as e:
                self.root.after(0, lambda msg=str(e): messagebox.showerror("Error", msg))
            finally:
                self.root.after(0, lambda: self.dedup_btn.config(
                    state=tk.NORMAL, text="Remove Duplicates"))

        threading.Thread(target=run, daemon=True).start()

    # ── System tray ────────────────────────────────────────────

    def create_tray_icon(self):
        if self.tray_icon is not None:
            return

        def create_icon_image():
            icon_path = os.path.join(os.path.dirname(sys.argv[0]), 'fileorg.ico')
            if os.path.exists(icon_path):
                return Image.open(icon_path)
            img = Image.new('RGB', (64, 64), color=(30, 120, 200))
            draw = ImageDraw.Draw(img)
            draw.rectangle((16, 12, 48, 52), fill=(255, 255, 255))
            draw.rectangle((20, 20, 44, 24), fill=(30, 120, 200))
            draw.rectangle((20, 30, 44, 34), fill=(30, 120, 200))
            draw.rectangle((20, 40, 44, 44), fill=(30, 120, 200))
            return img

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


# ---------------------------------------------------------------------------
# Logging — write to a stable location alongside the config
# ---------------------------------------------------------------------------
_log_dir = os.path.join(os.getenv('APPDATA'), 'DownloadOrganizer')
os.makedirs(_log_dir, exist_ok=True)

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S',
    filename=os.path.join(_log_dir, 'app.log'),
    filemode='a',
)
logger = logging.getLogger(__name__)
logger.info('=== Download Organizer started ===')

if __name__ == "__main__":
    app = DownloadOrganizerApp()
    app.root.mainloop()
