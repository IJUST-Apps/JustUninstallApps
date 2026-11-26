import ctypes
import gc
import json
import os
import re
import shutil
import stat
import subprocess
import sys
import threading
import time
import tkinter as tk
from tkinter import filedialog, messagebox, Menu
import webbrowser
import winreg

import customtkinter as ctk
import psutil
from PIL import Image, ImageDraw

# Optional pywin32 for native app icons and shortcut target inspection
try:
    import win32api
    import win32con
    import win32gui
    import win32ui
    HAS_WIN32 = True
except ImportError:
    HAS_WIN32 = False

try:
    from win32com.client import Dispatch as _Win32Dispatch
    HAS_WIN32COM = True
except ImportError:
    HAS_WIN32COM = False


def is_admin():
    if os.name != "nt":
        return True
    try:
        return bool(ctypes.windll.shell32.IsUserAnAdmin())
    except Exception:
        return False


def relaunch_as_admin():
    if os.name != "nt" or is_admin():
        return False
    try:
        if getattr(sys, "frozen", False):
            exe_path = sys.executable
            params = " ".join(f'"{arg}"' for arg in sys.argv[1:])
        else:
            exe_path = sys.executable
            script_path = os.path.abspath(sys.argv[0])
            rest_args = " ".join(f'"{arg}"' for arg in sys.argv[1:])
            params = f'"{script_path}" {rest_args}'.strip()

        ret = ctypes.windll.shell32.ShellExecuteW(
            None, "runas", exe_path, params, None, 1
        )
        return int(ret) > 32
    except Exception:
        return False


def get_settings_file_path():
    if getattr(sys, "frozen", False):
        base_dir = os.path.dirname(os.path.abspath(sys.executable))
    else:
        base_dir = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_dir, "justuninstallapps_settings.json")


def _format_elapsed(seconds):
    total_sec = max(0, int(seconds))
    mins, secs = divmod(total_sec, 60)
    hrs, mins = divmod(mins, 60)
    if hrs > 0:
        return f"{hrs:02d}:{mins:02d}:{secs:02d}"
    return f"{mins:02d}:{secs:02d}"


# Default appearance
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("dark-blue")


class UninstallProgressPopup(ctk.CTkToplevel):
    def __init__(self, parent, theme, on_cancel_callback, on_close_callback=None):
        super().__init__(parent)
        self.title("Uninstalling Process")
        self.geometry("620x430")
        self.minsize(500, 350)
        self.transient(parent)
        self.on_cancel_callback = on_cancel_callback
        self.on_close_callback = on_close_callback
        self._is_finished = False
        self._cancelled = False
        self._closed_notified = False
        self._start_time = time.monotonic()
        self._last_detail = "Initializing..."

        self.configure(fg_color=theme["bg"])

        self.card = ctk.CTkFrame(
            self,
            fg_color=theme["card"],
            corner_radius=12,
            border_width=1,
            border_color=theme["accent"],
        )
        self.card.pack(fill="both", expand=True, padx=14, pady=14)

        top_header = ctk.CTkFrame(self.card, fg_color="transparent")
        top_header.pack(fill="x", padx=16, pady=(14, 4))

        self.lbl_title = ctk.CTkLabel(
            top_header,
            text="⚙️ Uninstalling Process (Active)",
            font=ctk.CTkFont(size=17, weight="bold"),
            text_color=theme["accent"],
        )
        self.lbl_title.pack(side="left")

        self.lbl_elapsed = ctk.CTkLabel(
            top_header,
            text="Elapsed: 00:00  •  0%",
            font=ctk.CTkFont(family="Consolas", size=12, weight="bold"),
            text_color=theme["subtext"],
        )
        self.lbl_elapsed.pack(side="right")

        self.lbl_current = ctk.CTkLabel(
            self.card,
            text="Preparing uninstallation...",
            font=ctk.CTkFont(size=13, weight="bold"),
            text_color=theme["text"],
            anchor="w",
            justify="left",
            wraplength=560,
        )
        self.lbl_current.pack(anchor="w", padx=16, pady=(0, 2))

        self.lbl_detail = ctk.CTkLabel(
            self.card,
            text="Status: Starting worker thread — longer times during installer/leftover scans are normal.",
            font=ctk.CTkFont(size=11),
            text_color=theme["subtext"],
            anchor="w",
            justify="left",
            wraplength=560,
        )
        self.lbl_detail.pack(anchor="w", padx=16, pady=(0, 8))

        self.progress_bar = ctk.CTkProgressBar(
            self.card,
            height=12,
            fg_color=theme["fg"],
            progress_color=theme["accent"],
        )
        self.progress_bar.pack(fill="x", padx=16, pady=(0, 10))
        self.progress_bar.set(0.05)
        self._last_progress = 0.05

        self.step_box = ctk.CTkTextbox(
            self.card,
            wrap="word",
            font=ctk.CTkFont(family="Consolas", size=11),
            fg_color=theme["fg"],
            text_color=theme["text"],
            height=180,
        )
        self.step_box.pack(fill="both", expand=True, padx=16, pady=(0, 12))
        self.step_box.configure(state="disabled")

        btn_row = ctk.CTkFrame(self.card, fg_color="transparent")
        btn_row.pack(fill="x", padx=16, pady=(0, 14))

        self.btn_cancel = ctk.CTkButton(
            btn_row,
            text="✖ Cancel",
            width=120,
            height=36,
            fg_color="#dc2626",
            hover_color="#b91c1c",
            text_color="#ffffff",
            font=ctk.CTkFont(size=12, weight="bold"),
            command=self._handle_cancel,
        )
        self.btn_cancel.pack(side="left")

        self.lbl_hint = ctk.CTkLabel(
            btn_row,
            text="Tip: External uninstallers & leftover scans can take time — check live details above.",
            font=ctk.CTkFont(size=10),
            text_color=theme["subtext"],
        )
        self.lbl_hint.pack(side="left", padx=12)

        self.btn_ok = ctk.CTkButton(
            btn_row,
            text="✔ OK",
            width=120,
            height=36,
            fg_color="#16a34a",
            hover_color="#15803d",
            text_color="#ffffff",
            font=ctk.CTkFont(size=12, weight="bold"),
            command=self._handle_ok,
            state="disabled",
        )
        self.btn_ok.pack(side="right")

        self.protocol("WM_DELETE_WINDOW", self._handle_ok_or_hide)
        self.bind("<Destroy>", self._on_destroy_event, add="+")
        self._tick_clock()

    def _notify_closed_once(self):
        if not self._closed_notified:
            self._closed_notified = True
            if callable(self.on_close_callback):
                self.after(0, self.on_close_callback)

    def _on_destroy_event(self, event):
        if event.widget is self:
            self._notify_closed_once()

    def _tick_clock(self):
        if not self.winfo_exists():
            return
        elapsed_str = _format_elapsed(time.monotonic() - self._start_time)
        pct = int(round(self._last_progress * 100))
        self.lbl_elapsed.configure(text=f"Elapsed: {elapsed_str}  •  {pct}%")
        if not self._is_finished:
            self.after(500, self._tick_clock)

    def update_live_detail(self, detail_text):
        if not self.winfo_exists() or self._is_finished:
            return
        self._last_detail = detail_text
        self.lbl_detail.configure(text=f"Live Activity: {detail_text}")

    def append_step(self, text, progress=None, detail=None):
        if not self.winfo_exists():
            return
        self.lbl_current.configure(text=text)
        if progress is not None:
            clamped = max(0.0, min(1.0, progress))
            self._last_progress = clamped
            self.progress_bar.set(clamped)
        if detail is not None:
            self._last_detail = detail
            self.lbl_detail.configure(text=f"Live Activity: {detail}")
        elif not self._is_finished:
            self.lbl_detail.configure(text=f"Live Activity: {text}")

        elapsed_str = _format_elapsed(time.monotonic() - self._start_time)
        pct = int(round(self._last_progress * 100))
        self.lbl_elapsed.configure(text=f"Elapsed: {elapsed_str}  •  {pct}%")

        self.step_box.configure(state="normal")
        self.step_box.insert("end", f"[{elapsed_str}] • {text}\n")
        self.step_box.see("end")
        self.step_box.configure(state="disabled")

    def mark_cancelled(self):
        if not self.winfo_exists():
            return
        self._cancelled = True
        self._is_finished = True
        self.lbl_title.configure(text="⚙️ Uninstalling Process (Cancelled)")
        self.btn_cancel.configure(state="disabled")
        self.btn_ok.configure(state="normal")
        self.append_step("Uninstallation cancelled by user.", progress=1.0, detail="Stopped by user.")

    def mark_completed(self, summary_text="Completed uninstall."):
        if not self.winfo_exists():
            return
        self._is_finished = True
        self.lbl_title.configure(text="⚙️ Uninstalling Process (Completed)")
        self.btn_cancel.configure(state="disabled")
        self.btn_ok.configure(state="normal")
        self.append_step(summary_text, progress=1.0, detail="All steps completed.")

    def _handle_cancel(self):
        if self._is_finished:
            self._handle_ok()
            return
        self._cancelled = True
        self.btn_cancel.configure(state="disabled", text="Cancelling...")
        self.append_step("Cancellation requested — finishing current step...", detail="Waiting for current operation to stop safely...")
        if callable(self.on_cancel_callback):
            self.on_cancel_callback()

    def _handle_ok(self):
        self._notify_closed_once()
        if self.winfo_exists():
            self.destroy()

    def _handle_ok_or_hide(self):
        if self._is_finished:
            self._handle_ok()
        else:
            self._handle_cancel()


class JustUninstallApps(ctk.CTk):
    def __init__(self):
        super().__init__()

        # Clean window title
        self.title("JustUninstallApps")
        self.geometry("1280x780")
        self.minsize(980, 580)

        # App icon
        try:
            icon_path = "icon.ico"
            if hasattr(sys, "_MEIPASS"):
                icon_path = os.path.join(sys._MEIPASS, "icon.ico")
            elif not os.path.exists(icon_path):
                if os.path.exists(os.path.join("assets", "icon.ico")):
                    icon_path = os.path.join("assets", "icon.ico")

            if os.path.exists(icon_path):
                self.iconbitmap(icon_path)
        except Exception:
            pass

        # Curated distinct themes with matched contrast & harmonized RAM/Junk/Action colors
        self.themes = {
            "Dark Slate": {
                "mode": "Dark",
                "bg": "#0f172a",
                "fg": "#1e293b",
                "accent": "#38bdf8",
                "accent_hover": "#0ea5e9",
                "accent_text": "#0f172a",
                "tab_text": "#f8fafc",
                "text": "#f8fafc",
                "subtext": "#94a3b8",
                "card": "#162033",
                "row": "#1e293b",
                "row_hover": "#283548",
                "border": "#334155",
                "uninstall_btn": "#dc2626",
                "uninstall_hover": "#b91c1c",
                "contribute_btn": "#e11d48",
                "contribute_hover": "#be123c",
                "junk_btn": "#ea580c",
                "junk_hover": "#c2410c",
                "ram_btn": "#16a34a",
                "ram_hover": "#15803d",
            },
            "Yellow with Black": {
                "mode": "Dark",
                "bg": "#141414",
                "fg": "#1f1f1f",
                "accent": "#eab308",
                "accent_hover": "#ca8a04",
                "accent_text": "#141414",
                "tab_text": "#f8fafc",
                "text": "#f8fafc",
                "subtext": "#cbd5e1",
                "card": "#1a1a1a",
                "row": "#222222",
                "row_hover": "#2c2c2c",
                "border": "#3a3a3a",
                "uninstall_btn": "#dc2626",
                "uninstall_hover": "#b91c1c",
                "contribute_btn": "#e11d48",
                "contribute_hover": "#be123c",
                "junk_btn": "#ea580c",
                "junk_hover": "#c2410c",
                "ram_btn": "#16a34a",
                "ram_hover": "#15803d",
            },
            "Light": {
                "mode": "Light",
                "bg": "#f1f5f9",
                "fg": "#ffffff",
                "accent": "#2563eb",
                "accent_hover": "#1d4ed8",
                "accent_text": "#ffffff",
                "tab_text": "#0f172a",
                "text": "#0f172a",
                "subtext": "#475569",
                "card": "#ffffff",
                "row": "#f8fafc",
                "row_hover": "#e2e8f0",
                "border": "#cbd5e1",
                "uninstall_btn": "#dc2626",
                "uninstall_hover": "#b91c1c",
                "contribute_btn": "#e11d48",
                "contribute_hover": "#be123c",
                "junk_btn": "#ea580c",
                "junk_hover": "#c2410c",
                "ram_btn": "#15803d",
                "ram_hover": "#166534",
            },
            "Warm": {
                "mode": "Dark",
                "bg": "#1c1613",
                "fg": "#28201b",
                "accent": "#d97706",
                "accent_hover": "#b45309",
                "accent_text": "#ffffff",
                "tab_text": "#faf5f0",
                "text": "#faf5f0",
                "subtext": "#d6c7bc",
                "card": "#231c18",
                "row": "#2d241f",
                "row_hover": "#382d27",
                "border": "#4a3c34",
                "uninstall_btn": "#dc2626",
                "uninstall_hover": "#b91c1c",
                "contribute_btn": "#e11d48",
                "contribute_hover": "#be123c",
                "junk_btn": "#ea580c",
                "junk_hover": "#c2410c",
                "ram_btn": "#16a34a",
                "ram_hover": "#15803d",
            },
            "Dark Blue": {
                "mode": "Dark",
                "bg": "#0b1324",
                "fg": "#13203b",
                "accent": "#3b82f6",
                "accent_hover": "#2563eb",
                "accent_text": "#ffffff",
                "tab_text": "#eff6ff",
                "text": "#eff6ff",
                "subtext": "#93c5fd",
                "card": "#101c33",
                "row": "#172746",
                "row_hover": "#1e3259",
                "border": "#263d69",
                "uninstall_btn": "#dc2626",
                "uninstall_hover": "#b91c1c",
                "contribute_btn": "#e11d48",
                "contribute_hover": "#be123c",
                "junk_btn": "#ea580c",
                "junk_hover": "#c2410c",
                "ram_btn": "#16a34a",
                "ram_hover": "#15803d",
            },
            "Forest Green": {
                "mode": "Dark",
                "bg": "#0d1914",
                "fg": "#14261f",
                "accent": "#10b981",
                "accent_hover": "#059669",
                "accent_text": "#062016",
                "tab_text": "#ecfdf5",
                "text": "#ecfdf5",
                "subtext": "#a7f3d0",
                "card": "#11211b",
                "row": "#182e25",
                "row_hover": "#203d31",
                "border": "#2a4f40",
                "uninstall_btn": "#dc2626",
                "uninstall_hover": "#b91c1c",
                "contribute_btn": "#e11d48",
                "contribute_hover": "#be123c",
                "junk_btn": "#ea580c",
                "junk_hover": "#c2410c",
                "ram_btn": "#22c55e",
                "ram_hover": "#16a34a",
            },
            "Soft Sand": {
                "mode": "Light",
                "bg": "#f5f1ea",
                "fg": "#fdfbf7",
                "accent": "#b45309",
                "accent_hover": "#92400e",
                "accent_text": "#ffffff",
                "tab_text": "#292524",
                "text": "#292524",
                "subtext": "#57534e",
                "card": "#fcfaf5",
                "row": "#f3ede2",
                "row_hover": "#e7decb",
                "border": "#d6cbb8",
                "uninstall_btn": "#dc2626",
                "uninstall_hover": "#b91c1c",
                "contribute_btn": "#e11d48",
                "contribute_hover": "#be123c",
                "junk_btn": "#ea580c",
                "junk_hover": "#c2410c",
                "ram_btn": "#15803d",
                "ram_hover": "#166534",
            },
            "Midnight Teal": {
                "mode": "Dark",
                "bg": "#0b191e",
                "fg": "#12262e",
                "accent": "#14b8a6",
                "accent_hover": "#0d9488",
                "accent_text": "#061b1e",
                "tab_text": "#f0fdfa",
                "text": "#f0fdfa",
                "subtext": "#99f6e4",
                "card": "#0f2128",
                "row": "#16303a",
                "row_hover": "#1d3e4b",
                "border": "#265161",
                "uninstall_btn": "#dc2626",
                "uninstall_hover": "#b91c1c",
                "contribute_btn": "#e11d48",
                "contribute_hover": "#be123c",
                "junk_btn": "#ea580c",
                "junk_hover": "#c2410c",
                "ram_btn": "#16a34a",
                "ram_hover": "#15803d",
            },
            "Royal Purple": {
                "mode": "Dark",
                "bg": "#161124",
                "fg": "#211936",
                "accent": "#a855f7",
                "accent_hover": "#9333ea",
                "accent_text": "#ffffff",
                "tab_text": "#faf5ff",
                "text": "#faf5ff",
                "subtext": "#d8b4fe",
                "card": "#1c152e",
                "row": "#261d3d",
                "row_hover": "#31264f",
                "border": "#423369",
                "uninstall_btn": "#dc2626",
                "uninstall_hover": "#b91c1c",
                "contribute_btn": "#e11d48",
                "contribute_hover": "#be123c",
                "junk_btn": "#ea580c",
                "junk_hover": "#c2410c",
                "ram_btn": "#16a34a",
                "ram_hover": "#15803d",
            },
            "Cool Gray": {
                "mode": "Dark",
                "bg": "#18181b",
                "fg": "#27272a",
                "accent": "#a1a1aa",
                "accent_hover": "#d4d4d8",
                "accent_text": "#18181b",
                "tab_text": "#fafafa",
                "text": "#fafafa",
                "subtext": "#a1a1aa",
                "card": "#202024",
                "row": "#27272a",
                "row_hover": "#323236",
                "border": "#3f3f46",
                "uninstall_btn": "#dc2626",
                "uninstall_hover": "#b91c1c",
                "contribute_btn": "#e11d48",
                "contribute_hover": "#be123c",
                "junk_btn": "#ea580c",
                "junk_hover": "#c2410c",
                "ram_btn": "#16a34a",
                "ram_hover": "#15803d",
            },
        }

        self.default_theme_name = self._load_default_theme()
        self.current_theme_name = self.default_theme_name

        self.raw_installed_apps = []
        self.app_checkboxes = []
        self.action_logs = []
        self._cancel_uninstall_flag = threading.Event()
        self._active_progress_popup = None
        self._pending_uninstall_popups = set()
        self._auto_refresh_after_popups = False
        self._last_uninstall_success_count = 0
        self._last_uninstall_cancelled = False
        self._uninstalled_app_names = set()
        self._uninstalled_reg_keys = set()
        self._render_generation = 0

        self.icon_cache = {}
        self.fallback_icon = self._create_fallback_icon()

        self.create_header_frame()
        self.create_main_tabs()
        self._setup_smooth_scrolling()

        self.author = "IJUST-Apps"
        self.publisher = "IJUST-Apps"

        self.apply_theme(self.current_theme_name)
        self.refresh_installed_apps()
        self.refresh_space_info()

    def _register_uninstall_popup(self, token):
        self._pending_uninstall_popups.add(token)

    def _release_uninstall_popup(self, token):
        self._pending_uninstall_popups.discard(token)
        # Immediately refresh the apps list as soon as the progress popup or any cleanup popup closes
        self._sync_apps_list_now()
        if self._auto_refresh_after_popups and not self._pending_uninstall_popups:
            self._auto_refresh_after_popups = False
            self.after(50, self._run_post_uninstall_auto_refresh)

    def _strip_uninstalled_from_memory(self):
        if not self._uninstalled_app_names and not self._uninstalled_reg_keys:
            return
        self.raw_installed_apps = [
            app for app in self.raw_installed_apps
            if app.get("name", "").casefold() not in self._uninstalled_app_names
            and app.get("registry_key_name", "").casefold() not in self._uninstalled_reg_keys
        ]

    def _sync_apps_list_now(self):
        """Immediately removes uninstalled apps from the list view and triggers a fresh registry scan."""
        self._strip_uninstalled_from_memory()
        if hasattr(self, "raw_installed_apps"):
            self.populate_apps_list()
        self.refresh_installed_apps()

    def _run_post_uninstall_auto_refresh(self):
        self._sync_apps_list_now()
        if not self._pending_uninstall_popups:
            self.refresh_space_info()
            self.log_action("Auto-refreshed installed apps list after Completed uninstall.")

    def _setup_smooth_scrolling(self):
        """Routes mouse-wheel events smoothly to the active scrollable frame under the cursor."""
        def _on_mousewheel(event):
            widget = event.widget
            target_scroll = None
            curr = widget
            while curr is not None:
                if isinstance(curr, ctk.CTkScrollableFrame):
                    target_scroll = curr
                    break
                curr = getattr(curr, "master", None)

            if target_scroll is None:
                try:
                    current_tab = self.tab_view.get()
                    if current_tab == "Installed Apps":
                        target_scroll = self.apps_scroll
                    elif current_tab == "Space":
                        target_scroll = self.space_scroll
                except Exception:
                    pass

            if target_scroll is not None and hasattr(target_scroll, "_parent_canvas"):
                canvas = target_scroll._parent_canvas
                if event.delta:
                    step = int(-1 * (event.delta / 40))
                    if step == 0:
                        step = -1 if event.delta > 0 else 1
                    canvas.yview_scroll(step, "units")
                return "break"

        self.bind_all("<MouseWheel>", _on_mousewheel, add="+")

    def _load_default_theme(self):
        settings_path = get_settings_file_path()
        try:
            if os.path.exists(settings_path):
                with open(settings_path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                saved = data.get("default_theme", "")
                if saved in self.themes:
                    return saved
        except Exception:
            pass
        return "Dark Slate"

    def save_current_theme_as_default(self):
        settings_path = get_settings_file_path()
        try:
            data = {}
            if os.path.exists(settings_path):
                with open(settings_path, "r", encoding="utf-8") as f:
                    data = json.load(f)
            data["default_theme"] = self.current_theme_name
            with open(settings_path, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=2)
            self.default_theme_name = self.current_theme_name
            self.log_action(f"Default theme set to: {self.current_theme_name}")
            messagebox.showinfo(
                "Default Theme Saved",
                f"'{self.current_theme_name}' is now set as your default theme.",
            )
        except Exception as e:
            self.log_action(f"Failed to save default theme: {e}", success=False)
            messagebox.showerror("Error", f"Could not save default theme:\n{e}")

    def _create_fallback_icon(self):
        try:
            img = Image.new("RGBA", (20, 20), (0, 0, 0, 0))
            draw = ImageDraw.Draw(img)
            draw.rounded_rectangle([2, 2, 18, 18], radius=4, fill=(239, 68, 68, 255))
            draw.rectangle([6, 6, 14, 14], fill=(255, 255, 255, 220))
            return ctk.CTkImage(light_image=img, dark_image=img, size=(20, 20))
        except Exception:
            return None

    def resolve_exe_path(self, raw_path):
        if not raw_path:
            return ""
        cmd = os.path.expandvars(raw_path.strip())

        if cmd.startswith('"'):
            end_quote = cmd.find('"', 1)
            if end_quote != -1:
                candidate = cmd[1:end_quote]
                if os.path.exists(candidate) and candidate.lower().endswith((".exe", ".ico")):
                    return candidate
                elif os.path.isdir(candidate):
                    try:
                        for f in os.listdir(candidate):
                            if f.lower().endswith(".exe"):
                                return os.path.join(candidate, f)
                    except Exception:
                        pass

        match = re.search(r'([a-zA-Z]:\\[^"\n\r]+?\.(?:exe|ico))', cmd, re.IGNORECASE)
        if match:
            candidate = match.group(1).strip()
            if os.path.exists(candidate):
                return candidate

        first_token = cmd.split()[0].strip('"')
        if os.path.exists(first_token):
            if first_token.lower().endswith((".exe", ".ico")):
                return first_token
            elif os.path.isdir(first_token):
                try:
                    for f in os.listdir(first_token):
                        if f.lower().endswith(".exe"):
                            return os.path.join(first_token, f)
                except Exception:
                    pass

        return ""

    def get_app_mini_icon(self, raw_path):
        try:
            exe_path = self.resolve_exe_path(raw_path)
            if not exe_path or not os.path.exists(exe_path):
                return self.fallback_icon

            cache_key = exe_path.lower()
            if cache_key in self.icon_cache:
                return self.icon_cache[cache_key]

            pil_img = None

            if HAS_WIN32:
                try:
                    large, small = win32gui.ExtractIconEx(exe_path, 0)
                    hicon = large[0] if large else (small[0] if small else None)
                    if hicon:
                        ico_x = win32api.GetSystemMetrics(win32con.SM_CXICON)
                        ico_y = win32api.GetSystemMetrics(win32con.SM_CYICON)
                        hdc = win32ui.CreateDCFromHandle(win32gui.GetDC(0))
                        hbmp = win32ui.CreateBitmap()
                        hbmp.CreateCompatibleBitmap(hdc, ico_x, ico_y)
                        hdc_mem = hdc.CreateCompatibleDC()
                        hdc_mem.SelectObject(hbmp)
                        win32gui.DrawIconEx(
                            hdc_mem.GetHandleOutput(), 0, 0, hicon, ico_x, ico_y, 0, 0, win32con.DI_NORMAL
                        )
                        bmpinfo = hbmp.GetInfo()
                        bmpstr = hbmp.GetBitmapBits(True)
                        pil_img = Image.frombuffer(
                            "RGBA",
                            (bmpinfo["bmWidth"], bmpinfo["bmHeight"]),
                            bmpstr,
                            "raw",
                            "BGRA",
                            0,
                            1,
                        )
                        win32gui.DeleteObject(hbmp.GetHandle())
                        hdc_mem.DeleteDC()
                        hdc.DeleteDC()
                        for h in (large or []) + (small or []):
                            try:
                                win32gui.DestroyIcon(h)
                            except Exception:
                                pass
                except Exception:
                    pass

            if pil_img is None and exe_path.lower().endswith(".ico"):
                pil_img = Image.open(exe_path).convert("RGBA")

            if pil_img is not None:
                pil_img = pil_img.resize((20, 20), Image.Resampling.LANCZOS)
                ctk_img = ctk.CTkImage(light_image=pil_img, dark_image=pil_img, size=(20, 20))
                self.icon_cache[cache_key] = ctk_img
                return ctk_img
        except Exception:
            pass

        return self.fallback_icon

    def create_header_frame(self):
        self.header_frame = ctk.CTkFrame(self, corner_radius=0, height=75)
        self.header_frame.pack(side="top", fill="x", padx=0, pady=0)

        self.title_label = ctk.CTkLabel(
            self.header_frame,
            text="⚡🧹 JustUninstallApps",
            font=ctk.CTkFont(size=18, weight="bold"),
        )
        self.title_label.pack(side="left", padx=15, pady=15)

        self.admin_badge = ctk.CTkLabel(
            self.header_frame,
            text="🛡️ ADMIN" if is_admin() else "⚠️ STANDARD",
            font=ctk.CTkFont(size=11, weight="bold"),
            corner_radius=6,
            padx=8,
            pady=3,
        )
        self.admin_badge.pack(side="left", padx=(0, 10), pady=15)

        self.btn_refresh_all = ctk.CTkButton(
            self.header_frame,
            text="🔄 Refresh All App",
            width=130,
            font=ctk.CTkFont(size=12, weight="bold"),
            command=self.refresh_all_app,
        )
        self.btn_refresh_all.pack(side="right", padx=10, pady=15)

        self.btn_about = ctk.CTkButton(
            self.header_frame,
            text="ℹ️ About",
            width=80,
            command=self.show_about_dialog,
        )
        self.btn_about.pack(side="right", padx=5, pady=15)

        self.btn_UpdatesAndSupport = ctk.CTkButton(
            self.header_frame,
            text="🛠️ UpdatesAndSupport",
            width=145,
            command=lambda: webbrowser.open("https://github.com/IJUST-Apps"),
        )
        self.btn_UpdatesAndSupport.pack(side="right", padx=5, pady=15)

        self.btn_contribute = ctk.CTkButton(
            self.header_frame,
            text="❤️ Contribute",
            fg_color="#e11d48",
            hover_color="#be123c",
            text_color="#ffffff",
            width=115,
            command=self.show_donation_dialog,
        )
        self.btn_contribute.pack(side="right", padx=5, pady=15)

        self.opacity_slider = ctk.CTkSlider(
            self.header_frame,
            from_=0.5,
            to=1.0,
            number_of_steps=10,
            width=75,
            command=self.change_opacity,
        )
        self.opacity_slider.set(1.0)
        self.opacity_slider.pack(side="right", padx=5, pady=15)

        self.opacity_label = ctk.CTkLabel(
            self.header_frame, text="Opacity:", font=ctk.CTkFont(size=11)
        )
        self.opacity_label.pack(side="right", padx=(5, 0), pady=15)

        self.btn_set_default_theme = ctk.CTkButton(
            self.header_frame,
            text="⭐ Set Default",
            width=95,
            font=ctk.CTkFont(size=11, weight="bold"),
            command=self.save_current_theme_as_default,
        )
        self.btn_set_default_theme.pack(side="right", padx=(4, 8), pady=15)

        self.theme_menu = ctk.CTkOptionMenu(
            self.header_frame,
            values=list(self.themes.keys()),
            command=self.apply_theme,
            width=155,
        )
        self.theme_menu.set(self.current_theme_name)
        self.theme_menu.pack(side="right", padx=(8, 2), pady=15)

    def create_main_tabs(self):
        # Top pinned bar so Clean JUNK & Clean RAM stay in the exact same position across all tabs
        self.pinned_action_bar = ctk.CTkFrame(self, corner_radius=8)
        self.pinned_action_bar.pack(side="top", fill="x", padx=15, pady=(10, 0))

        self.pinned_info_label = ctk.CTkLabel(
            self.pinned_action_bar,
            text="Quick Maintenance: Clean temporary junk or free up system RAM anytime",
            font=ctk.CTkFont(size=12, weight="bold"),
        )
        self.pinned_info_label.pack(side="left", padx=12, pady=8)

        self.btn_clean_ram = ctk.CTkButton(
            self.pinned_action_bar,
            text="🧠 Clean RAM",
            fg_color="#16a34a",
            hover_color="#15803d",
            text_color="#ffffff",
            font=ctk.CTkFont(size=12, weight="bold"),
            width=140,
            command=self.clean_system_ram,
        )
        self.btn_clean_ram.pack(side="right", padx=(6, 10), pady=8)

        self.btn_clean_junk = ctk.CTkButton(
            self.pinned_action_bar,
            text="🧹 Clean JUNK",
            fg_color="#ea580c",
            hover_color="#c2410c",
            text_color="#ffffff",
            font=ctk.CTkFont(size=12, weight="bold"),
            width=140,
            command=lambda: self.clean_junk_temp_cache(),
        )
        self.btn_clean_junk.pack(side="right", padx=(6, 4), pady=8)

        self.tab_view = ctk.CTkTabview(self, corner_radius=10)
        self.tab_view.pack(side="top", fill="both", expand=True, padx=15, pady=(8, 15))

        self.tab_apps = self.tab_view.add("Installed Apps")
        self.tab_space = self.tab_view.add("Space")
        self.tab_logs = self.tab_view.add("📜 Logs & History")

        self.setup_apps_tab()
        self.setup_space_tab()
        self.setup_logs_tab()

    def setup_apps_tab(self):
        frame = self.tab_apps

        self.apps_toolbar = ctk.CTkFrame(frame, fg_color="transparent")
        self.apps_toolbar.pack(side="top", fill="x", padx=5, pady=5)

        self.search_var = ctk.StringVar()
        self.search_var.trace_add("write", lambda *args: self.populate_apps_list())
        self.txt_search = ctk.CTkEntry(
            self.apps_toolbar,
            placeholder_text="🔍 Search Installed Apps...",
            textvariable=self.search_var,
            width=190,
        )
        self.txt_search.pack(side="left", padx=5)

        self.hide_ms_var = ctk.BooleanVar(value=True)
        self.chk_hide_ms = ctk.CTkCheckBox(
            self.apps_toolbar,
            text="Hide Microsoft Apps",
            variable=self.hide_ms_var,
            command=self.populate_apps_list,
        )
        self.chk_hide_ms.pack(side="left", padx=(8, 6))

        self.show_uninstall_proc_var = ctk.BooleanVar(value=True)
        self.chk_show_uninstall_proc = ctk.CTkCheckBox(
            self.apps_toolbar,
            text="Show Uninstalling Process",
            variable=self.show_uninstall_proc_var,
        )
        self.chk_show_uninstall_proc.pack(side="left", padx=(4, 8))

        self.lbl_sort = ctk.CTkLabel(
            self.apps_toolbar, text="Arrange By:", font=ctk.CTkFont(size=11, weight="bold")
        )
        self.lbl_sort.pack(side="left", padx=(6, 2))

        self.sort_var = ctk.StringVar(value="Name (A-Z)")
        self.sort_menu = ctk.CTkOptionMenu(
            self.apps_toolbar,
            values=[
                "Name (A-Z)", "Name (Z-A)",
                "Size (Largest First)", "Size (Smallest First)",
                "Recently Installed", "Publisher (A-Z)",
            ],
            command=lambda v: self.populate_apps_list(),
            variable=self.sort_var,
            width=150,
        )
        self.sort_menu.pack(side="left", padx=5)

        self.btn_uninstall = ctk.CTkButton(
            self.apps_toolbar,
            text="🗑️ Complete Uninstall",
            fg_color="#dc2626",
            hover_color="#b91c1c",
            text_color="#ffffff",
            font=ctk.CTkFont(size=12, weight="bold"),
            width=165,
            command=self.uninstall_selected_apps,
        )
        self.btn_uninstall.pack(side="right", padx=5)

        self.btn_refresh_apps = ctk.CTkButton(
            self.apps_toolbar,
            text="🔄 Refresh",
            width=95,
            font=ctk.CTkFont(size=12, weight="bold"),
            command=self.refresh_installed_apps,
        )
        self.btn_refresh_apps.pack(side="right", padx=(5, 3))

        self.lbl_refresh_tip = ctk.CTkLabel(
            self.apps_toolbar,
            text="Tip: Refresh after installing a new app",
            font=ctk.CTkFont(size=10),
            text_color="#94a3b8",
        )
        self.lbl_refresh_tip.pack(side="right", padx=(2, 8))

        self.apps_scroll = ctk.CTkScrollableFrame(
            frame, label_text="Installed Applications Inventory"
        )
        self.apps_scroll.pack(side="bottom", fill="both", expand=True, padx=5, pady=10)

    def setup_space_tab(self):
        frame = self.tab_space

        self.space_top_bar = ctk.CTkFrame(frame, corner_radius=10)
        self.space_top_bar.pack(side="top", fill="x", padx=8, pady=(8, 6))

        self.space_title_lbl = ctk.CTkLabel(
            self.space_top_bar,
            text="💾 Storage & Disk Space",
            font=ctk.CTkFont(size=16, weight="bold"),
        )
        self.space_title_lbl.pack(side="left", padx=14, pady=12)

        self.btn_refresh_space = ctk.CTkButton(
            self.space_top_bar,
            text="🔄 REFRESH SPACE",
            height=42,
            width=185,
            font=ctk.CTkFont(size=13, weight="bold"),
            command=self.refresh_space_info,
        )
        self.btn_refresh_space.pack(side="right", padx=14, pady=10)

        # System Partition (C:)
        self.system_partition_frame = ctk.CTkFrame(frame, corner_radius=10, border_width=1)
        self.system_partition_frame.pack(side="top", fill="x", padx=8, pady=(0, 8))

        self.lbl_sys_section = ctk.CTkLabel(
            self.system_partition_frame,
            text="System Partition",
            font=ctk.CTkFont(size=14, weight="bold"),
        )
        self.lbl_sys_section.pack(anchor="w", padx=14, pady=(10, 4))

        self.sys_c_container = ctk.CTkFrame(self.system_partition_frame, fg_color="transparent")
        self.sys_c_container.pack(fill="x", padx=8, pady=(0, 8))

        # All Partitions & Drives (other drives)
        self.space_scroll = ctk.CTkScrollableFrame(
            frame, label_text="All Partitions & Drives"
        )
        self.space_scroll.pack(side="bottom", fill="both", expand=True, padx=8, pady=(0, 8))

    def refresh_all_app(self):
        """Refreshes the entire application state: search, filters, icon cache, apps list, and disk space."""
        self.search_var.set("")
        self.hide_ms_var.set(True)
        self.sort_var.set("Name (A-Z)")
        self.icon_cache.clear()
        self.refresh_installed_apps()
        self.refresh_space_info()
        self.log_action("Refreshed entire application (apps, filters, icon cache, and storage).")

    def refresh_space_info(self):
        self.btn_refresh_space.configure(state="disabled", text="⏳ Refreshing...")
        threading.Thread(target=self._load_space_worker, daemon=True).start()

    def _load_space_worker(self):
        partitions_data = []
        try:
            parts = psutil.disk_partitions(all=False)
            for p in parts:
                try:
                    usage = psutil.disk_usage(p.mountpoint)
                    drive_letter = os.path.splitdrive(p.mountpoint)[0].upper()
                    is_c_drive = drive_letter == "C:"
                    partitions_data.append({
                        "device": p.device,
                        "mountpoint": p.mountpoint,
                        "fstype": p.fstype,
                        "total": usage.total,
                        "used": usage.used,
                        "free": usage.free,
                        "percent": usage.percent,
                        "is_c": is_c_drive,
                    })
                except Exception:
                    continue
        except Exception as e:
            self.after(0, lambda err=e: self.log_action(f"Space scan error: {err}", success=False))

        partitions_data.sort(key=lambda x: x["mountpoint"].lower())
        self.after(0, lambda data=partitions_data: self._render_space_info(data))

    def _create_partition_card(self, parent, item, theme, highlight=False):
        card_bg = theme["row_hover"] if highlight else theme["row"]
        border_c = theme["accent"] if highlight else theme["border"]

        p_card = ctk.CTkFrame(
            parent,
            fg_color=card_bg,
            corner_radius=8,
            border_width=2 if highlight else 1,
            border_color=border_c,
        )
        p_card.pack(fill="x", padx=6, pady=5)

        top_line = ctk.CTkFrame(p_card, fg_color="transparent")
        top_line.pack(fill="x", padx=12, pady=(10, 4))

        lbl_drive = ctk.CTkLabel(
            top_line,
            text=f"💽 {item['mountpoint']} ({item['device']} • {item['fstype'] or 'NTFS'})",
            font=ctk.CTkFont(size=13, weight="bold"),
            text_color=theme["accent"] if highlight else theme["text"],
        )
        lbl_drive.pack(side="left")

        lbl_free = ctk.CTkLabel(
            top_line,
            text=f"Available: {self.format_size(item['free'])}  /  Total: {self.format_size(item['total'])}",
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color="#22c55e" if item["free"] > 10 * 1024**3 else "#f97316",
        )
        lbl_free.pack(side="right")

        bar = ctk.CTkProgressBar(
            p_card,
            height=11,
            fg_color=theme["bg"],
            progress_color="#dc2626" if item["percent"] >= 90 else theme["accent"],
        )
        bar.pack(fill="x", padx=12, pady=(2, 6))
        bar.set(min(1.0, max(0.0, item["percent"] / 100.0)))

        lbl_sub = ctk.CTkLabel(
            p_card,
            text=f"Used: {self.format_size(item['used'])}  •  Free: {self.format_size(item['free'])}  •  {item['percent']:.1f}% Used",
            font=ctk.CTkFont(size=11),
            text_color=theme["subtext"],
        )
        lbl_sub.pack(anchor="w", padx=12, pady=(0, 10))

    def _render_space_info(self, partitions_data):
        theme = self.themes.get(self.current_theme_name, self.themes["Dark Slate"])

        for widget in self.sys_c_container.winfo_children():
            widget.destroy()
        for widget in self.space_scroll.winfo_children():
            widget.destroy()

        c_entry = next((item for item in partitions_data if item["is_c"]), None)
        other_entries = [item for item in partitions_data if not item["is_c"]]

        if c_entry:
            self._create_partition_card(self.sys_c_container, c_entry, theme, highlight=True)
        else:
            lbl_empty_c = ctk.CTkLabel(
                self.sys_c_container,
                text="C: partition not found.",
                font=ctk.CTkFont(size=12),
                text_color=theme["subtext"],
            )
            lbl_empty_c.pack(anchor="w", padx=10, pady=6)

        if other_entries:
            for item in other_entries:
                self._create_partition_card(self.space_scroll, item, theme, highlight=False)
        else:
            lbl_no_others = ctk.CTkLabel(
                self.space_scroll,
                text="No additional partitions or drives detected.",
                font=ctk.CTkFont(size=12),
                text_color=theme["subtext"],
            )
            lbl_no_others.pack(anchor="w", padx=12, pady=12)

        self.btn_refresh_space.configure(state="normal", text="🔄 REFRESH SPACE")

    def setup_logs_tab(self):
        frame = self.tab_logs

        self.logs_toolbar = ctk.CTkFrame(frame, fg_color="transparent")
        self.logs_toolbar.pack(side="top", fill="x", padx=5, pady=5)

        self.btn_clear_logs = ctk.CTkButton(
            self.logs_toolbar,
            text="🧹 Clear Logs",
            width=110,
            command=self.clear_logs,
        )
        self.btn_clear_logs.pack(side="left", padx=5)

        self.logs_textbox = ctk.CTkTextbox(
            frame, wrap="word", font=ctk.CTkFont(family="Consolas", size=12)
        )
        self.logs_textbox.pack(side="bottom", fill="both", expand=True, padx=5, pady=10)
        self.logs_textbox.insert("0.0", "=== JustUninstallApps Log Center Initialized ===\n")
        self.logs_textbox.configure(state="disabled")

    def log_action(self, message, success=True):
        prefix = "[SUCCESS]" if success else "[FAILED/NOTE]"
        log_entry = f"{prefix} {message}\n"
        self.action_logs.append(log_entry)

        self.logs_textbox.configure(state="normal")
        self.logs_textbox.insert("end", log_entry)
        self.logs_textbox.see("end")
        self.logs_textbox.configure(state="disabled")

    def clear_logs(self):
        self.action_logs.clear()
        self.logs_textbox.configure(state="normal")
        self.logs_textbox.delete("0.0", "end")
        self.logs_textbox.insert("0.0", "=== Logs Cleared ===\n")
        self.logs_textbox.configure(state="disabled")

    def apply_theme(self, theme_name):
        self.current_theme_name = theme_name
        theme = self.themes.get(theme_name, self.themes["Dark Slate"])

        ctk.set_appearance_mode(theme["mode"])

        # Main window & header
        self.configure(fg_color=theme["bg"])
        self.header_frame.configure(fg_color=theme["fg"])
        self.title_label.configure(text_color=theme["accent"])
        self.admin_badge.configure(
            fg_color=theme["card"],
            text_color="#22c55e" if is_admin() else "#f59e0b",
        )
        self.opacity_label.configure(text_color=theme["text"])

        # Pinned bar, Clean JUNK & Clean RAM buttons
        self.pinned_action_bar.configure(
            fg_color=theme["card"],
            border_width=1,
            border_color=theme["border"],
        )
        self.pinned_info_label.configure(text_color=theme["text"])
        self.btn_clean_junk.configure(
            fg_color=theme["junk_btn"],
            hover_color=theme["junk_hover"],
            text_color="#ffffff",
        )
        self.btn_clean_ram.configure(
            fg_color=theme["ram_btn"],
            hover_color=theme["ram_hover"],
            text_color="#ffffff",
        )

        # Tabview (high-contrast readable tab labels)
        self.tab_view.configure(
            fg_color=theme["card"],
            segmented_button_fg_color=theme["fg"],
            segmented_button_selected_color=theme["accent"],
            segmented_button_selected_hover_color=theme["accent_hover"],
            segmented_button_unselected_color=theme["row"],
            segmented_button_unselected_hover_color=theme["row_hover"],
            text_color=theme["tab_text"],
            text_color_disabled=theme["subtext"],
        )

        # Header buttons & dropdown
        self.btn_refresh_all.configure(
            fg_color=theme["accent"],
            text_color=theme["accent_text"],
            hover_color=theme["accent_hover"],
        )
        self.btn_set_default_theme.configure(
            fg_color=theme["row"],
            hover_color=theme["row_hover"],
            text_color=theme["accent"],
            border_width=1,
            border_color=theme["border"],
        )
        self.btn_about.configure(
            fg_color=theme["row"],
            hover_color=theme["row_hover"],
            text_color=theme["text"],
        )
        self.btn_UpdatesAndSupport.configure(
            fg_color=theme["row"],
            hover_color=theme["row_hover"],
            text_color=theme["text"],
        )
        self.btn_contribute.configure(
            fg_color=theme["contribute_btn"],
            hover_color=theme["contribute_hover"],
            text_color="#ffffff",
        )
        self.theme_menu.configure(
            fg_color=theme["row"],
            button_color=theme["accent"],
            button_hover_color=theme["accent_hover"],
            dropdown_fg_color=theme["card"],
            dropdown_hover_color=theme["row_hover"],
            dropdown_text_color=theme["text"],
            text_color=theme["text"],
        )

        # Apps tab controls
        self.txt_search.configure(
            fg_color=theme["row"],
            border_color=theme["border"],
            text_color=theme["text"],
            placeholder_text_color=theme["subtext"],
        )
        self.chk_hide_ms.configure(
            text_color=theme["text"],
            fg_color=theme["accent"],
            hover_color=theme["accent_hover"],
            border_color=theme["border"],
            checkmark_color=theme["accent_text"],
        )
        self.chk_show_uninstall_proc.configure(
            text_color=theme["text"],
            fg_color=theme["accent"],
            hover_color=theme["accent_hover"],
            border_color=theme["border"],
            checkmark_color=theme["accent_text"],
        )
        self.lbl_sort.configure(text_color=theme["text"])
        self.sort_menu.configure(
            fg_color=theme["row"],
            button_color=theme["accent"],
            button_hover_color=theme["accent_hover"],
            dropdown_fg_color=theme["card"],
            dropdown_hover_color=theme["row_hover"],
            dropdown_text_color=theme["text"],
            text_color=theme["text"],
        )
        self.btn_refresh_apps.configure(
            fg_color=theme["row"],
            hover_color=theme["row_hover"],
            text_color=theme["accent"],
            border_width=1,
            border_color=theme["border"],
        )
        self.lbl_refresh_tip.configure(text_color=theme["subtext"])
        self.btn_uninstall.configure(
            fg_color=theme["uninstall_btn"],
            hover_color=theme["uninstall_hover"],
            text_color="#ffffff",
        )
        self.apps_scroll.configure(
            fg_color=theme["bg"],
            border_color=theme["border"],
            border_width=1,
            label_text_color=theme["accent"],
        )

        # Space tab controls
        self.space_top_bar.configure(
            fg_color=theme["fg"],
            border_width=1,
            border_color=theme["border"],
        )
        self.space_title_lbl.configure(text_color=theme["accent"])
        self.btn_refresh_space.configure(
            fg_color=theme["accent"],
            hover_color=theme["accent_hover"],
            text_color=theme["accent_text"],
        )
        self.system_partition_frame.configure(
            fg_color=theme["fg"],
            border_color=theme["border"],
        )
        self.lbl_sys_section.configure(text_color=theme["accent"])
        self.space_scroll.configure(
            fg_color=theme["bg"],
            border_color=theme["border"],
            border_width=1,
            label_text_color=theme["accent"],
        )

        # Logs tab controls
        self.btn_clear_logs.configure(
            fg_color=theme["row"],
            hover_color=theme["row_hover"],
            text_color=theme["text"],
        )
        self.logs_textbox.configure(
            fg_color=theme["fg"],
            text_color=theme["text"],
            border_color=theme["border"],
            border_width=1,
        )

        # Re-render dynamic lists so every row & partition card picks up the new theme
        if hasattr(self, "raw_installed_apps") and self.raw_installed_apps:
            self.populate_apps_list()
        if hasattr(self, "space_scroll"):
            self.refresh_space_info()

    def change_opacity(self, value):
        self.attributes("-alpha", value)

    def show_about_dialog(self):
        messagebox.showinfo(
            "About JustUninstallApps",
            "JustUninstallApps is a free open source app to uninstall apps made by IJUST-Apps",
        )

    def show_donation_dialog(self):
        webbrowser.open("https://ko-fi.com/ijustapps")

    def get_folder_size(self, folder_path):
        total_size = 0
        try:
            if folder_path and os.path.exists(folder_path) and os.path.isdir(folder_path):
                for dirpath, dirnames, filenames in os.walk(folder_path):
                    for f in filenames:
                        fp = os.path.join(dirpath, f)
                        try:
                            if not os.path.islink(fp):
                                total_size += os.path.getsize(fp)
                        except Exception:
                            pass
        except Exception:
            pass
        return total_size

    def format_size(self, size_bytes):
        if size_bytes <= 0:
            return "0 B"
        for unit in ["B", "KB", "MB", "GB"]:
            if size_bytes < 1024.0:
                return f"{size_bytes:.1f} {unit}"
            size_bytes /= 1024.0
        return f"{size_bytes:.1f} TB"

    def refresh_installed_apps(self):
        if hasattr(self, "btn_refresh_apps"):
            self.btn_refresh_apps.configure(state="disabled", text="⏳ Loading...")
        threading.Thread(target=self._load_registry_apps_worker, daemon=True).start()

    def _load_registry_apps_worker(self):
        apps = []
        try:
            reg_targets = [
                (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"),
                (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall"),
                (winreg.HKEY_CURRENT_USER, r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"),
            ]

            for hkey, path in reg_targets:
                try:
                    open_key = winreg.OpenKey(hkey, path, 0, winreg.KEY_READ)
                    subkeys_count = winreg.QueryInfoKey(open_key)[0]

                    for i in range(subkeys_count):
                        try:
                            sub_key_name = winreg.EnumKey(open_key, i)
                            if sub_key_name.casefold() in self._uninstalled_reg_keys:
                                continue
                            sub_key_full_path = f"{path}\\{sub_key_name}"
                            sub_key = winreg.OpenKey(open_key, sub_key_name)

                            try:
                                name = winreg.QueryValueEx(sub_key, "DisplayName")[0]
                            except Exception:
                                winreg.CloseKey(sub_key)
                                continue

                            if not name or not name.strip():
                                winreg.CloseKey(sub_key)
                                continue

                            if name.strip().casefold() in self._uninstalled_app_names:
                                winreg.CloseKey(sub_key)
                                continue

                            uninstall_string = ""
                            try:
                                uninstall_string = winreg.QueryValueEx(sub_key, "UninstallString")[0]
                            except Exception:
                                pass

                            install_location = ""
                            try:
                                install_location = winreg.QueryValueEx(sub_key, "InstallLocation")[0]
                            except Exception:
                                pass

                            publisher = "Unknown Publisher"
                            try:
                                publisher = winreg.QueryValueEx(sub_key, "Publisher")[0]
                            except Exception:
                                pass

                            version = ""
                            try:
                                version = winreg.QueryValueEx(sub_key, "DisplayVersion")[0]
                            except Exception:
                                pass

                            est_size = 0
                            try:
                                est_size = winreg.QueryValueEx(sub_key, "EstimatedSize")[0] * 1024
                            except Exception:
                                pass

                            install_date = ""
                            try:
                                install_date = winreg.QueryValueEx(sub_key, "InstallDate")[0]
                            except Exception:
                                pass

                            if est_size == 0 and install_location and os.path.exists(install_location):
                                est_size = self.get_folder_size(install_location)

                            l_name = name.lower()
                            l_pub = publisher.lower()
                            is_ms = (
                                "microsoft" in l_pub
                                or "microsoft corporation" in l_pub
                                or l_name.startswith("kb")
                                or "windows" in l_name
                            )

                            apps.append({
                                "name": name.strip(),
                                "publisher": publisher.strip(),
                                "version": version.strip(),
                                "uninstall_string": uninstall_string.strip(),
                                "install_location": install_location.strip(),
                                "registry_path": sub_key_full_path,
                                "registry_key_name": sub_key_name,
                                "hkey": hkey,
                                "is_ms": is_ms,
                                "size_bytes": est_size,
                                "install_date": install_date.strip(),
                            })
                            winreg.CloseKey(sub_key)
                        except Exception:
                            continue
                    winreg.CloseKey(open_key)
                except Exception:
                    pass
        except Exception as e:
            self.after(0, lambda err=e: self.log_action(f"Registry scan error: {err}", success=False))

        unique_apps = {}
        for app in apps:
            unique_apps[app["name"]] = app

        self.raw_installed_apps = list(unique_apps.values())
        self._strip_uninstalled_from_memory()
        self.after(0, self._on_apps_loaded)

    def _on_apps_loaded(self):
        if hasattr(self, "btn_refresh_apps"):
            self.btn_refresh_apps.configure(state="normal", text="🔄 Refresh")
        self.populate_apps_list()

    def populate_apps_list(self):
        self._render_generation += 1
        gen_id = self._render_generation
        theme = self.themes.get(self.current_theme_name, self.themes["Dark Slate"])

        for widget in self.apps_scroll.winfo_children():
            widget.destroy()

        try:
            self.apps_scroll._parent_canvas.yview_moveto(0.0)
        except Exception:
            pass

        hide_ms = self.hide_ms_var.get()
        search_query = self.search_var.get().lower() if hasattr(self, "search_var") else ""
        sort_mode = self.sort_var.get() if hasattr(self, "sort_var") else "Name (A-Z)"

        filtered = []
        for app in self.raw_installed_apps:
            if app.get("name", "").casefold() in self._uninstalled_app_names:
                continue
            if app.get("registry_key_name", "").casefold() in self._uninstalled_reg_keys:
                continue
            if hide_ms and app["is_ms"]:
                continue
            if (
                search_query
                and search_query not in app["name"].lower()
                and search_query not in app["publisher"].lower()
            ):
                continue
            filtered.append(app)

        if sort_mode == "Name (A-Z)":
            filtered.sort(key=lambda x: (x.get("name", "").casefold(), x.get("publisher", "").casefold()))
        elif sort_mode == "Name (Z-A)":
            filtered.sort(key=lambda x: (x.get("name", "").casefold(), x.get("publisher", "").casefold()), reverse=True)
        elif sort_mode == "Size (Largest First)":
            filtered.sort(key=lambda x: (int(x.get("size_bytes") or 0), x.get("name", "").casefold()), reverse=True)
        elif sort_mode == "Size (Smallest First)":
            filtered.sort(key=lambda x: (x.get("size_bytes") in (None, 0), int(x.get("size_bytes") or 0), x.get("name", "").casefold()))
        elif sort_mode == "Recently Installed":
            def _install_date_sort_key(item):
                raw = re.sub(r"[^0-9]", "", str(item.get("install_date", "")))
                valid = len(raw) == 8 and raw[:4].isdigit()
                return (valid, raw if valid else "", item.get("name", "").casefold())
            filtered.sort(key=_install_date_sort_key, reverse=True)
        elif sort_mode == "Publisher (A-Z)":
            filtered.sort(key=lambda x: (x.get("publisher", "").casefold(), x.get("name", "").casefold()))

        self.app_checkboxes = []
        self.apps_scroll.configure(
            label_text=f"Installed Applications Inventory ({len(filtered)} shown)"
        )

        # Render in smooth small chunks so the UI never freezes or lags while scrolling
        self._render_apps_batch(filtered, 0, 18, gen_id, theme)

    def _render_apps_batch(self, filtered, start_idx, batch_size, gen_id, theme):
        if gen_id != self._render_generation:
            return

        end_idx = min(start_idx + batch_size, len(filtered))
        for idx in range(start_idx, end_idx):
            app = filtered[idx]
            row_frame = ctk.CTkFrame(
                self.apps_scroll,
                fg_color=theme["row"],
                border_width=1,
                border_color=theme["border"],
                corner_radius=6,
            )
            row_frame.pack(fill="x", padx=5, pady=4)

            var = ctk.BooleanVar(value=False)

            chk = ctk.CTkCheckBox(
                row_frame,
                text="",
                variable=var,
                width=24,
                fg_color=theme["accent"],
                hover_color=theme["accent_hover"],
                border_color=theme["border"],
                checkmark_color=theme["accent_text"],
            )
            chk.pack(side="left", padx=(10, 4), pady=8)

            icon_source = app.get("uninstall_string") or app.get("install_location")
            mini_icon = self.get_app_mini_icon(icon_source)
            if mini_icon is not None:
                icon_lbl = ctk.CTkLabel(row_frame, text="", image=mini_icon, width=22)
                icon_lbl.pack(side="left", padx=(0, 8), pady=8)
            else:
                icon_lbl = None

            display_text = f"{app['name']}"
            if app["version"]:
                display_text += f" (v{app['version']})"
            size_str = self.format_size(app["size_bytes"]) if app["size_bytes"] > 0 else "Unknown Size"
            display_text += f"   • Publisher: {app['publisher']}   • Size: {size_str}"

            lbl_name = ctk.CTkLabel(
                row_frame,
                text=display_text,
                font=ctk.CTkFont(size=12, weight="bold"),
                text_color=theme["text"],
                anchor="w",
            )
            lbl_name.pack(side="left", fill="x", expand=True, padx=2, pady=8)
            lbl_name.bind("<Button-1>", lambda e, v=var: v.set(not v.get()))

            badge_color = theme["accent"] if not app["is_ms"] else theme["subtext"]
            badge_text = "[THIRD-PARTY]" if not app["is_ms"] else "[MICROSOFT]"
            lbl_badge = ctk.CTkLabel(
                row_frame,
                text=badge_text,
                font=ctk.CTkFont(size=11, weight="bold"),
                text_color=badge_color,
            )
            lbl_badge.pack(side="right", padx=10, pady=8)

            menu = Menu(row_frame, tearoff=0)
            menu.add_command(
                label="📂 Open Folder", command=lambda a=app: self.open_app_folder(a)
            )
            menu.add_command(
                label="🗑️ Complete Uninstall",
                command=lambda v=var: [v.set(True), self.uninstall_selected_apps()],
            )

            def show_menu(event, m=menu):
                try:
                    m.tk_popup(event.x_root, event.y_root)
                finally:
                    m.grab_release()

            for w in (row_frame, chk, lbl_name, lbl_badge, icon_lbl):
                if w is not None:
                    w.bind("<Button-3>", show_menu)

            self.app_checkboxes.append((var, app))

        if end_idx < len(filtered):
            self.after(
                1,
                lambda: self._render_apps_batch(filtered, end_idx, batch_size, gen_id, theme),
            )

    def browse_path_in_explorer(self, target_path):
        if not target_path:
            return
        folder_to_open = target_path if os.path.isdir(target_path) else os.path.dirname(target_path)
        if folder_to_open and os.path.exists(folder_to_open):
            try:
                if os.path.isfile(target_path):
                    subprocess.Popen(f'explorer /select,"{target_path}"', shell=True)
                else:
                    subprocess.Popen(f'explorer "{folder_to_open}"', shell=True)
                self.log_action(f"Opened manual cleanup location: {target_path}")
            except Exception as e:
                self.log_action(f"Failed to open folder {folder_to_open}: {e}", success=False)
        else:
            messagebox.showwarning(
                "Path Not Found",
                f"The path does not exist anymore:\n{target_path}",
            )

    def open_app_folder(self, app):
        folder = app["install_location"]
        if folder and os.path.exists(folder):
            self.browse_path_in_explorer(folder)
        else:
            messagebox.showinfo(
                "Folder Not Found",
                f"No valid installation directory registered for '{app['name']}'.",
            )
            self.log_action(
                f"Failed to open folder for {app['name']}: Path not registered.", success=False
            )

    def _show_ram_result_popup(self, summary_text, detail_lines):
        popup = ctk.CTkToplevel(self)
        popup.title("🧠 Clean RAM Results")
        popup.geometry("540x360")
        popup.minsize(430, 280)
        popup.transient(self)
        popup.grab_set()

        theme = self.themes.get(self.current_theme_name, self.themes["Dark Slate"])
        popup.configure(fg_color=theme["bg"])

        card = ctk.CTkFrame(
            popup,
            fg_color=theme["card"],
            corner_radius=12,
            border_width=1,
            border_color=theme["ram_btn"],
        )
        card.pack(fill="both", expand=True, padx=14, pady=14)

        lbl_header = ctk.CTkLabel(
            card,
            text="🧠 Clean RAM Complete",
            font=ctk.CTkFont(size=16, weight="bold"),
            text_color="#4ade80",
        )
        lbl_header.pack(anchor="w", padx=16, pady=(14, 4))

        lbl_summary = ctk.CTkLabel(
            card,
            text=summary_text,
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color=theme["text"],
            justify="left",
            wraplength=490,
        )
        lbl_summary.pack(anchor="w", padx=16, pady=(0, 8))

        info_box = ctk.CTkTextbox(
            card,
            wrap="word",
            font=ctk.CTkFont(family="Consolas", size=11),
            fg_color=theme["fg"],
            text_color=theme["text"],
            height=150,
        )
        info_box.pack(fill="both", expand=True, padx=16, pady=(0, 12))
        info_box.insert("0.0", "\n".join(detail_lines))
        info_box.configure(state="disabled")

        btn_ok = ctk.CTkButton(
            card,
            text="OK",
            width=120,
            height=36,
            fg_color=theme["ram_btn"],
            hover_color=theme["ram_hover"],
            text_color="#ffffff",
            font=ctk.CTkFont(size=12, weight="bold"),
            command=popup.destroy,
        )
        btn_ok.pack(pady=(0, 14))

    def clean_system_ram(self):
        self.btn_clean_ram.configure(state="disabled", text="🧠 Cleaning...")
        threading.Thread(target=self._clean_ram_worker, daemon=True).start()

    def _clean_ram_worker(self):
        before_mem = psutil.virtual_memory()
        gc.collect()
        trimmed_count = 0

        if os.name == "nt":
            try:
                kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
                psapi = ctypes.WinDLL("psapi", use_last_error=True)

                PROCESS_SET_QUOTA = 0x0100
                PROCESS_QUERY_INFORMATION = 0x0400

                for proc in psutil.process_iter(["pid", "name"]):
                    pid = proc.info.get("pid")
                    if not pid or pid <= 4:
                        continue
                    try:
                        h_proc = kernel32.OpenProcess(
                            PROCESS_SET_QUOTA | PROCESS_QUERY_INFORMATION, False, pid
                        )
                        if h_proc:
                            try:
                                if psapi.EmptyWorkingSet(h_proc):
                                    trimmed_count += 1
                                else:
                                    kernel32.SetProcessWorkingSetSize(
                                        h_proc, ctypes.c_size_t(-1), ctypes.c_size_t(-1)
                                    )
                            finally:
                                kernel32.CloseHandle(h_proc)
                    except Exception:
                        continue
            except Exception as e:
                self.after(0, lambda err=e: self.log_action(f"RAM clean note: {err}", success=False))

        after_mem = psutil.virtual_memory()
        freed_bytes = max(0, after_mem.available - before_mem.available)
        freed_str = self.format_size(freed_bytes)

        summary = f"Freed {freed_str} of RAM across {trimmed_count} background processes."
        details = [
            f"RAM Available Before: {self.format_size(before_mem.available)} ({100 - before_mem.percent:.1f}% free)",
            f"RAM Available After:  {self.format_size(after_mem.available)} ({100 - after_mem.percent:.1f}% free)",
            f"Memory Reclaimed:     {freed_str}",
            f"Processes Optimized:  {trimmed_count}",
        ]

        self.after(0, lambda: self.log_action(f"Clean RAM: {summary}"))
        self.after(0, lambda: self.btn_clean_ram.configure(state="normal", text="🧠 Clean RAM"))
        self.after(0, lambda: self._show_ram_result_popup(summary, details))

    def _show_junk_result_popup(self, title_text, summary_text, detail_lines, is_warning=False, on_close_callback=None):
        popup = ctk.CTkToplevel(self)
        popup.title(title_text)
        popup.geometry("580x400")
        popup.minsize(450, 310)
        popup.transient(self)

        theme = self.themes.get(self.current_theme_name, self.themes["Dark Slate"])
        popup.configure(fg_color=theme["bg"])

        notified = {"done": False}

        def _notify_once():
            if not notified["done"]:
                notified["done"] = True
                if callable(on_close_callback):
                    self.after(0, on_close_callback)

        def _close_popup():
            _notify_once()
            if popup.winfo_exists():
                popup.destroy()

        popup.protocol("WM_DELETE_WINDOW", _close_popup)
        popup.bind("<Destroy>", lambda e: _notify_once() if e.widget is popup else None, add="+")

        card = ctk.CTkFrame(
            popup,
            fg_color=theme["card"],
            corner_radius=12,
            border_width=1,
            border_color=theme["junk_btn"],
        )
        card.pack(fill="both", expand=True, padx=14, pady=14)

        header_color = "#fb923c" if is_warning else "#4ade80"
        lbl_header = ctk.CTkLabel(
            card,
            text=title_text,
            font=ctk.CTkFont(size=16, weight="bold"),
            text_color=header_color,
        )
        lbl_header.pack(anchor="w", padx=16, pady=(14, 4))

        lbl_summary = ctk.CTkLabel(
            card,
            text=summary_text,
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color=theme["text"],
            justify="left",
            wraplength=520,
        )
        lbl_summary.pack(anchor="w", padx=16, pady=(0, 8))

        info_box = ctk.CTkTextbox(
            card,
            wrap="word",
            font=ctk.CTkFont(family="Consolas", size=11),
            fg_color=theme["fg"],
            text_color=theme["text"],
            height=170,
        )
        info_box.pack(fill="both", expand=True, padx=16, pady=(0, 12))
        info_box.insert("0.0", "\n".join(detail_lines) if detail_lines else "No extra notes.")
        info_box.configure(state="disabled")

        btn_row = ctk.CTkFrame(card, fg_color="transparent")
        btn_row.pack(pady=(0, 14))

        temp_dir = os.environ.get("TEMP") or r"C:\Windows\Temp"
        btn_browse_temp = ctk.CTkButton(
            btn_row,
            text="📂 Browse Temp",
            width=130,
            height=36,
            fg_color=theme["row"],
            hover_color=theme["row_hover"],
            text_color=theme["text"],
            font=ctk.CTkFont(size=12, weight="bold"),
            command=lambda p=temp_dir: self.browse_path_in_explorer(p),
        )
        btn_browse_temp.pack(side="left", padx=6)

        btn_ok = ctk.CTkButton(
            btn_row,
            text="OK",
            width=120,
            height=36,
            fg_color=theme["junk_btn"],
            hover_color=theme["junk_hover"],
            text_color="#ffffff",
            font=ctk.CTkFont(size=12, weight="bold"),
            command=_close_popup,
        )
        btn_ok.pack(side="left", padx=6)

    def clean_junk_temp_cache(self, on_complete_callback=None):
        self.btn_clean_junk.configure(state="disabled", text="🧹 Cleaning...")
        threading.Thread(
            target=self._clean_junk_worker,
            args=(on_complete_callback,),
            daemon=True,
        ).start()

    def _is_dead_or_empty_shortcut(self, lnk_path, cached_target=None):
        """Return True only for clearly empty or broken local shortcuts."""
        try:
            if not os.path.isfile(lnk_path) or os.path.getsize(lnk_path) == 0:
                return True
        except OSError:
            return False

        low_p = lnk_path.lower()
        if low_p.endswith(".url"):
            try:
                with open(lnk_path, "r", encoding="utf-8-sig", errors="ignore") as f:
                    content = f.read()
                m = re.search(r"(?im)^URL\s*=\s*(.+)$", content)
                if not m or not m.group(1).strip():
                    return True
                url_val = m.group(1).strip().strip('"')
                if url_val.lower().startswith("file:///"):
                    from urllib.parse import unquote, urlparse
                    parsed = urlparse(url_val)
                    local_target = unquote(parsed.path).lstrip("/")
                    if re.match(r"^[a-zA-Z]:", local_target):
                        local_target = local_target[0] + ":\\" + local_target[2:].lstrip("/\\")
                    elif parsed.netloc:
                        local_target = "\\\\" + parsed.netloc + "\\" + local_target.replace("/", "\\")
                    if local_target and not os.path.exists(os.path.expandvars(local_target)):
                        return True
            except Exception:
                return False
            return False

        if low_p.endswith(".lnk"):
            if cached_target is not None:
                target_path, _ = cached_target
            else:
                target_path, _ = self._get_shortcut_target_info(lnk_path)
            target_path = os.path.expandvars((target_path or "").strip().strip('"'))
            # Only delete shortcuts with a resolvable, clearly missing local target.
            if target_path and re.match(r"^[a-zA-Z]:[\\/]", target_path):
                return not os.path.exists(target_path)
        return False

    def _clean_junk_worker(self, on_complete_callback=None):
        deleted_files = 0
        deleted_broken_shortcuts = 0
        freed_bytes = 0
        in_use_items = []
        cleaned_shortcuts_list = []

        target_folders = []
        env_temp = os.environ.get("TEMP")
        if env_temp:
            target_folders.append(env_temp)
        win_dir = os.environ.get("SystemRoot", r"C:\Windows")
        target_folders.append(os.path.join(win_dir, "Temp"))
        local_app = os.environ.get("LOCALAPPDATA")
        if local_app:
            for sub in (
                os.path.join(local_app, "Temp"),
                os.path.join(local_app, "CrashDumps"),
                os.path.join(local_app, "D3DSCache"),
                os.path.join(local_app, "NVIDIA", "GLCache"),
            ):
                if os.path.isdir(sub) and sub not in target_folders:
                    target_folders.append(sub)

        # 1. Clean Temp & Cache files
        for folder in target_folders:
            if not folder or not os.path.exists(folder):
                continue
            for root, dirs, files in os.walk(folder, topdown=False):
                for fname in files:
                    fpath = os.path.join(root, fname)
                    try:
                        if os.path.islink(fpath):
                            os.unlink(fpath)
                            deleted_files += 1
                            continue
                        fsize = os.path.getsize(fpath)
                        os.remove(fpath)
                        deleted_files += 1
                        freed_bytes += fsize
                    except (PermissionError, OSError) as exc:
                        if len(in_use_items) < 80:
                            reason = (
                                "In use by app/Windows"
                                if getattr(exc, "winerror", None) in (5, 32, 33)
                                else (exc.strerror or str(exc))
                            )
                            in_use_items.append(f"{fpath}  [{reason}]")

                for dname in dirs:
                    dpath = os.path.join(root, dname)
                    try:
                        if os.path.isdir(dpath) and not os.path.islink(dpath):
                            if not os.listdir(dpath):
                                os.rmdir(dpath)
                    except Exception:
                        pass

        # 2. Clean empty/dead/broken .lnk & .url leftover shortcuts across Desktop & Start Menu
        for root_dir in self._collect_shortcut_search_roots():
            for root, dirs, files in os.walk(root_dir, topdown=False):
                for fname in files:
                    if not fname.lower().endswith((".lnk", ".url")):
                        continue
                    fpath = os.path.join(root, fname)
                    if self._is_dead_or_empty_shortcut(fpath):
                        try:
                            fsize = os.path.getsize(fpath) if os.path.exists(fpath) else 0
                        except Exception:
                            fsize = 0
                        ok, err = self._force_remove_file(fpath)
                        if ok:
                            deleted_broken_shortcuts += 1
                            freed_bytes += fsize
                            if len(cleaned_shortcuts_list) < 25:
                                cleaned_shortcuts_list.append(fpath)
                        elif len(in_use_items) < 80:
                            in_use_items.append(f"{fpath}  [Broken shortcut locked: {err}]")

                # Remove empty Start Menu / Desktop subfolders left behind
                if root.lower() != root_dir.lower():
                    try:
                        if os.path.isdir(root) and not os.listdir(root):
                            os.rmdir(root)
                    except Exception:
                        pass

        freed_str = self.format_size(freed_bytes)
        summary = (
            f"Cleaned {deleted_files} temp/cache file(s) and {deleted_broken_shortcuts} "
            f"empty/broken shortcut(s), freeing {freed_str}."
        )
        details = [
            f"Deleted temp & cache files:       {deleted_files}",
            f"Removed dead/empty shortcuts:     {deleted_broken_shortcuts}",
            f"Total disk space reclaimed:       {freed_str}",
            f"Skipped (in use or locked):       {len(in_use_items)}",
        ]
        if cleaned_shortcuts_list:
            details.append("")
            details.append("Removed dead/empty shortcuts (.lnk / .url):")
            details.extend(cleaned_shortcuts_list)
        if in_use_items:
            details.append("")
            details.append("Skipped items (currently in use by running apps or Windows):")
            details.extend(in_use_items[:30])
            if len(in_use_items) > 30:
                details.append(f"... and {len(in_use_items) - 30} more in-use items.")

        self.after(
            0,
            lambda: self.log_action(
                f"Clean JUNK: {summary} ({len(in_use_items)} in-use/locked skipped)"
            ),
        )
        self.after(0, lambda: self.btn_clean_junk.configure(state="normal", text="🧹 Clean JUNK"))
        self.after(
            0,
            lambda: self._show_junk_result_popup(
                "🧹 Clean JUNK Results",
                summary,
                details,
                is_warning=bool(in_use_items),
                on_close_callback=on_complete_callback,
            ),
        )

    def _emit_progress(self, msg, progress=None, detail=None):
        self.after(0, lambda m=msg: self.log_action(m))
        if self._active_progress_popup is not None and self._active_progress_popup.winfo_exists():
            self.after(
                0,
                lambda m=msg, p=progress, d=detail: self._active_progress_popup.append_step(m, p, d),
            )

    def _emit_live_detail(self, detail_text):
        if self._active_progress_popup is not None and self._active_progress_popup.winfo_exists():
            self.after(
                0,
                lambda d=detail_text: self._active_progress_popup.update_live_detail(d),
            )

    def _get_shortcut_target_info(self, lnk_path):
        target_str = ""
        workdir_str = ""
        if not lnk_path.lower().endswith(".lnk"):
            return target_str, workdir_str

        if HAS_WIN32COM:
            try:
                shell = _Win32Dispatch("WScript.Shell")
                shortcut = shell.CreateShortcut(lnk_path)
                target_str = (getattr(shortcut, "TargetPath", "") or "").strip()
                workdir_str = (getattr(shortcut, "WorkingDirectory", "") or "").strip()
                if target_str or workdir_str:
                    return target_str.lower(), workdir_str.lower()
            except Exception:
                pass

        # Fallback: inspect binary .lnk target string directly without spawning PowerShell per file
        try:
            with open(lnk_path, "rb") as f:
                raw = f.read()
            ascii_matches = re.findall(rb"[A-Za-z]:\\[^\x00-\x1f\"*<>|?]{2,260}", raw)
            if ascii_matches:
                target_str = ascii_matches[0].decode("latin-1", errors="ignore").strip().lower()
        except Exception:
            pass
        return target_str, workdir_str

    def _infer_install_location(self, app_name, uninstall_string, install_location):
        if install_location and os.path.exists(install_location):
            return os.path.normpath(install_location)

        exe_path = self.resolve_exe_path(uninstall_string)
        if exe_path and os.path.exists(exe_path):
            parent_dir = os.path.dirname(exe_path)
            norm_parent = os.path.normcase(os.path.abspath(parent_dir))
            win_root = os.path.normcase(os.path.abspath(os.environ.get("SystemRoot", r"C:\Windows")))
            if not norm_parent.startswith(win_root):
                return os.path.normpath(parent_dir)

        s_name = re.sub(r"[^a-zA-Z0-9]", "", app_name or "").lower()
        if len(s_name) >= 3:
            candidate_roots = [
                os.environ.get("ProgramFiles", r"C:\Program Files"),
                os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)"),
                os.path.join(os.environ.get("LOCALAPPDATA", ""), "Programs"),
            ]
            for c_root in candidate_roots:
                if c_root and os.path.isdir(c_root):
                    try:
                        for entry in os.listdir(c_root):
                            clean_entry = re.sub(r"[^a-zA-Z0-9]", "", entry).lower()
                            full_c = os.path.join(c_root, entry)
                            if s_name in clean_entry and os.path.isdir(full_c):
                                return os.path.normpath(full_c)
                    except Exception:
                        pass

        return install_location

    def _terminate_processes_in_dir(self, dir_path):
        if not dir_path:
            return []
        terminated = []
        norm_dir = os.path.normcase(os.path.abspath(dir_path)) + os.sep
        current_pid = os.getpid()

        for proc in psutil.process_iter(["pid", "name", "exe"]):
            pid = proc.info.get("pid")
            if not pid or pid == current_pid or pid <= 4:
                continue
            exe = proc.info.get("exe") or ""
            if not exe:
                continue
            try:
                norm_exe = os.path.normcase(os.path.abspath(exe))
                if norm_exe.startswith(norm_dir):
                    proc.terminate()
                    terminated.append(proc.info.get("name") or str(pid))
            except Exception:
                continue

        if terminated:
            psutil.wait_procs([p for p in psutil.process_iter() if (p.info.get("name") if hasattr(p, "info") else "") in terminated], timeout=2.0)
        return terminated

    def _prepare_uninstall_command(self, uninstall_string):
        cmd = os.path.expandvars((uninstall_string or "").strip())
        if not cmd:
            return ""
        if "msiexec" in cmd.lower() and "/x" not in cmd.lower():
            cmd = re.sub(
                r"msiexec(?:\.exe)?",
                "msiexec.exe /x",
                cmd,
                count=1,
                flags=re.IGNORECASE,
            )

        # Ensure an unquoted path containing spaces at the start is properly quoted for cmd.exe
        if not cmd.startswith('"'):
            m = re.match(r'^([A-Za-z]:\\[^"\r\n]+?\.exe)(.*)$', cmd, re.IGNORECASE)
            if m and " " in m.group(1):
                exe_part = m.group(1).strip()
                args_part = m.group(2)
                if os.path.exists(exe_part):
                    cmd = f'"{exe_part}"{args_part}'
        return cmd

    def _force_remove_file(self, fpath):
        try:
            if os.path.islink(fpath):
                os.unlink(fpath)
                return True, ""
            try:
                os.chmod(fpath, stat.S_IWRITE)
            except Exception:
                pass
            if os.name == "nt":
                try:
                    ctypes.windll.kernel32.SetFileAttributesW(fpath, 0x80)  # FILE_ATTRIBUTE_NORMAL
                except Exception:
                    pass
            os.remove(fpath)
            return True, ""
        except Exception as exc:
            return False, str(exc)

    def _force_remove_dir(self, dir_path):
        if not dir_path or not os.path.exists(dir_path):
            return True, ""

        self._terminate_processes_in_dir(dir_path)

        def _handle_remove_readonly(func, path, exc_info):
            try:
                if os.name == "nt":
                    ctypes.windll.kernel32.SetFileAttributesW(path, 0x80)
                os.chmod(path, stat.S_IWRITE)
                func(path)
            except Exception:
                pass

        try:
            shutil.rmtree(dir_path, onerror=_handle_remove_readonly)
            if not os.path.exists(dir_path):
                return True, ""
        except Exception:
            pass

        # Second pass: clear attributes/read-only flags across tree and remove remaining files/empty dirs
        remaining_errors = []
        removed_count = 0
        for root, dirs, files in os.walk(dir_path, topdown=False):
            for fname in files:
                fp = os.path.join(root, fname)
                ok, err = self._force_remove_file(fp)
                if ok:
                    removed_count += 1
                    if removed_count % 25 == 0:
                        self._emit_live_detail(f"Cleaning folder files ({removed_count} removed): {fp}")
                else:
                    remaining_errors.append(err)
            for dname in dirs:
                dp = os.path.join(root, dname)
                try:
                    if os.name == "nt":
                        ctypes.windll.kernel32.SetFileAttributesW(dp, 0x80)
                    os.chmod(dp, stat.S_IWRITE)
                    os.rmdir(dp)
                except Exception:
                    pass

        try:
            if os.name == "nt":
                ctypes.windll.kernel32.SetFileAttributesW(dir_path, 0x80)
            os.chmod(dir_path, stat.S_IWRITE)
            os.rmdir(dir_path)
        except Exception as exc:
            remaining_errors.append(str(exc))

        if not os.path.exists(dir_path):
            return True, ""

        # Third pass on Windows: shell rd /s /q fallback
        if os.name == "nt":
            try:
                si = subprocess.STARTUPINFO()
                si.dwFlags |= subprocess.STARTUP_F_USESHOWWINDOW
                subprocess.run(
                    f'attrib -R -S -H "{dir_path}\\*.*" /S /D & rd /s /q "{dir_path}"',
                    shell=True,
                    startupinfo=si,
                    timeout=30,
                )
            except Exception:
                pass

        if not os.path.exists(dir_path):
            return True, ""
        return False, remaining_errors[0] if remaining_errors else "Locked or in use by Windows/app"

    def _collect_shortcut_search_roots(self):
        roots = []
        user_profile = os.environ.get("USERPROFILE", "")
        appdata = os.environ.get("APPDATA", "")
        programdata = os.environ.get("PROGRAMDATA", r"C:\ProgramData")

        # Read Windows' actual redirected Desktop/Start Menu paths (including OneDrive).
        if os.name == "nt":
            shell_folder_keys = [
                (winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders"),
                (winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"),
                (winreg.HKEY_LOCAL_MACHINE, r"Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders"),
            ]
            wanted = {"desktop", "commondesktop", "programs", "common programs"}
            for hive, key_path in shell_folder_keys:
                try:
                    with winreg.OpenKey(hive, key_path) as key:
                        value_count = winreg.QueryInfoKey(key)[1]
                        for idx in range(value_count):
                            try:
                                value_name, value, _ = winreg.EnumValue(key, idx)
                                if value_name.casefold() in wanted and isinstance(value, str):
                                    expanded = os.path.expandvars(value.strip())
                                    if expanded:
                                        roots.append(expanded)
                            except OSError:
                                continue
                except OSError:
                    continue

        if user_profile:
            roots.extend([
                os.path.join(user_profile, "Desktop"),
                os.path.join(user_profile, "OneDrive", "Desktop"),
                os.path.join(user_profile, "OneDrive - Personal", "Desktop"),
            ])
        for env_key in ("OneDrive", "OneDriveConsumer", "OneDriveCommercial"):
            od_path = os.environ.get(env_key, "")
            if od_path:
                roots.append(os.path.join(od_path, "Desktop"))

        public_dir = os.environ.get("PUBLIC", r"C:\Users\Public")
        roots.append(os.path.join(public_dir, "Desktop"))
        if appdata:
            roots.append(os.path.join(appdata, r"Microsoft\Windows\Start Menu\Programs"))
        roots.append(os.path.join(programdata, r"Microsoft\Windows\Start Menu\Programs"))

        users_dir = os.path.join(os.environ.get("SystemDrive", "C:") + os.sep, "Users")
        if os.path.isdir(users_dir):
            try:
                for u_entry in os.listdir(users_dir):
                    user_dir = os.path.join(users_dir, u_entry)
                    for relative in ("Desktop", r"AppData\Roaming\Microsoft\Windows\Start Menu\Programs"):
                        candidate = os.path.join(user_dir, relative)
                        if os.path.isdir(candidate):
                            roots.append(candidate)
            except OSError:
                pass

        unique_roots = []
        seen = set()
        for root in roots:
            if not root:
                continue
            norm = os.path.normcase(os.path.abspath(os.path.expandvars(root)))
            if norm not in seen and os.path.isdir(root):
                seen.add(norm)
                unique_roots.append(root)
        return unique_roots

    def _delete_matching_shortcuts(self, app_name, install_location="", publisher="", shortcut_roots=None):
        removed = []
        failed_shortcuts = []

        tokens = []
        s_name = re.sub(r"[^a-zA-Z0-9]", "", app_name or "").lower()
        if len(s_name) >= 3:
            tokens.append(s_name)
        ignored_words = {
            "microsoft", "corporation", "company", "software", "application",
            "launcher", "installer", "setup", "update", "version", "edition",
            "redistributable", "runtime", "service", "assistant", "windows",
        }
        for word in re.findall(r"[a-zA-Z0-9]{4,}", app_name or ""):
            word = word.lower()
            if word not in ignored_words and word not in tokens:
                tokens.append(word)
        if publisher and publisher.lower() != "unknown publisher":
            s_pub = re.sub(r"[^a-zA-Z0-9]", "", publisher).lower()
            if len(s_pub) >= 5 and s_pub not in ("microsoft", "microsoftcorporation"):
                tokens.append(s_pub)

        norm_loc = os.path.normpath(install_location).lower() if install_location else ""
        roots_to_scan = shortcut_roots if shortcut_roots is not None else self._collect_shortcut_search_roots()
        scanned_files = 0

        for root_dir in roots_to_scan:
            for dirpath, dirnames, filenames in os.walk(root_dir, topdown=False):
                for fname in filenames:
                    if not fname.lower().endswith((".lnk", ".url")):
                        continue
                    scanned_files += 1
                    if scanned_files % 20 == 0:
                        self._emit_live_detail(f"Checking shortcuts ({scanned_files} scanned): {fname}")
                    fpath = os.path.join(dirpath, fname)
                    clean_fname = re.sub(r"[^a-zA-Z0-9]", "", os.path.splitext(fname)[0]).lower()
                    matched = any(tok in clean_fname for tok in tokens)

                    if not matched and (norm_loc or tokens):
                        target_s, workdir_s = self._get_shortcut_target_info(fpath)
                        clean_target = re.sub(r"[^a-zA-Z0-9]", "", target_s).lower()
                        clean_work = re.sub(r"[^a-zA-Z0-9]", "", workdir_s).lower()
                        if norm_loc and (norm_loc in target_s or norm_loc in workdir_s):
                            matched = True
                        elif any(tok in clean_target or tok in clean_work for tok in tokens):
                            matched = True

                    if matched:
                        ok, err = self._force_remove_file(fpath)
                        if ok:
                            removed.append(fpath)
                        else:
                            failed_shortcuts.append((fpath, f"Shortcut locked/protected: {err}"))

                if dirpath.lower() != root_dir.lower():
                    clean_dname = re.sub(r"[^a-zA-Z0-9]", "", os.path.basename(dirpath)).lower()
                    if any(tok in clean_dname for tok in tokens):
                        try:
                            if not os.listdir(dirpath):
                                os.rmdir(dirpath)
                        except Exception:
                            pass

        return removed, failed_shortcuts

    def _delete_matching_software_reg_keys(self, app_name, publisher=""):
        cleaned_keys = []
        tokens = []
        s_name = re.sub(r"[^a-zA-Z0-9]", "", app_name).lower()
        if len(s_name) >= 4:
            tokens.append(s_name)
        if publisher and publisher.lower() != "unknown publisher":
            s_pub = re.sub(r"[^a-zA-Z0-9]", "", publisher).lower()
            if len(s_pub) >= 4 and s_pub != "microsoft" and s_pub != "microsoftcorporation":
                tokens.append(s_pub)

        if not tokens:
            return cleaned_keys

        software_roots = [
            (winreg.HKEY_CURRENT_USER, r"SOFTWARE"),
            (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE"),
            (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Wow6432Node"),
        ]

        for hkey, base_path in software_roots:
            self._emit_live_detail(f"Scanning registry hive: {base_path}")
            try:
                with winreg.OpenKey(hkey, base_path, 0, winreg.KEY_READ | winreg.KEY_WRITE) as parent_key:
                    sub_count = winreg.QueryInfoKey(parent_key)[0]
                    names_to_delete = []
                    for idx in range(sub_count):
                        try:
                            kname = winreg.EnumKey(parent_key, idx)
                            clean_k = re.sub(r"[^a-zA-Z0-9]", "", kname).lower()
                            if any(tok in clean_k for tok in tokens):
                                names_to_delete.append(kname)
                        except Exception:
                            continue

                    for kname in names_to_delete:
                        full_sub = f"{base_path}\\{kname}"
                        try:
                            self._delete_reg_tree(hkey, full_sub)
                            cleaned_keys.append(
                                f"{'HKCU' if hkey == winreg.HKEY_CURRENT_USER else 'HKLM'}\\{full_sub}"
                            )
                        except Exception:
                            pass
            except Exception:
                pass

        return cleaned_keys

    def _delete_reg_tree(self, hkey, key_path):
        with winreg.OpenKey(hkey, key_path, 0, winreg.KEY_READ | winreg.KEY_WRITE) as target:
            sub_count = winreg.QueryInfoKey(target)[0]
            for i in range(sub_count - 1, -1, -1):
                child_name = winreg.EnumKey(target, i)
                self._delete_reg_tree(hkey, f"{key_path}\\{child_name}")
        winreg.DeleteKey(hkey, key_path)

    def uninstall_selected_apps(self):
        selected = [app for var, app in self.app_checkboxes if var.get()]
        if not selected:
            messagebox.showwarning(
                "Selection Required", "Please check at least one application to uninstall."
            )
            return

        confirm = messagebox.askyesno(
            "Confirm Complete Uninstall",
            f"Are you sure you want to run complete uninstallation, force-delete app folders, and deep-clean shortcuts, registry, and leftovers for {len(selected)} app(s)?\n\n"
            "Note: Official uninstallers and deep leftover scans can take time — progress details will be shown live.",
        )
        if not confirm:
            return

        self._cancel_uninstall_flag.clear()
        self._auto_refresh_after_popups = True
        theme = self.themes.get(self.current_theme_name, self.themes["Dark Slate"])

        if self.show_uninstall_proc_var.get():
            progress_token = "uninstall_progress_popup"
            self._register_uninstall_popup(progress_token)
            popup = UninstallProgressPopup(
                self,
                theme=theme,
                on_cancel_callback=self._cancel_uninstall_flag.set,
                on_close_callback=lambda tok=progress_token: self._release_uninstall_popup(tok),
            )
            self._active_progress_popup = popup
        else:
            self._active_progress_popup = None

        threading.Thread(target=self._uninstall_worker, args=(selected,), daemon=True).start()

    def _uninstall_worker(self, selected):
        success_count = 0
        failed_items = []
        total = max(1, len(selected))
        shortcut_roots = self._collect_shortcut_search_roots()

        for idx, app in enumerate(selected):
            if self._cancel_uninstall_flag.is_set():
                break

            name = app["name"]
            publisher = app.get("publisher", "")
            uninst_str = app["uninstall_string"]
            loc = self._infer_install_location(name, uninst_str, app["install_location"])
            reg_key_name = app["registry_key_name"]
            base_prog = idx / total
            app_prefix = f"App {idx + 1}/{len(selected)} ({name})"

            self._emit_progress(
                f"[{idx + 1}/{len(selected)}] Starting complete uninstall: {name}",
                base_prog + 0.05 / total,
                detail=f"{app_prefix} • Preparing uninstall & force-cleanup steps...",
            )

            # 1. Run built-in uninstaller first if available.
            uninstall_returncode = 999 if not uninst_str else None
            if uninst_str and not self._cancel_uninstall_flag.is_set():
                proc = None
                try:
                    si = subprocess.STARTUPINFO()
                    si.dwFlags |= subprocess.STARTUP_F_USESHOWWINDOW

                    cmd_to_run = self._prepare_uninstall_command(uninst_str)

                    self._emit_progress(
                        f"Running uninstaller for {name} (may take a while — normal)...",
                        base_prog + 0.25 / total,
                        detail=f"{app_prefix} • Waiting on external uninstaller: {cmd_to_run[:90]}",
                    )
                    proc = subprocess.Popen(cmd_to_run, shell=True, startupinfo=si)
                    wait_start = time.monotonic()
                    timed_out = False
                    while proc.poll() is None:
                        if self._cancel_uninstall_flag.is_set():
                            try:
                                proc.terminate()
                            except Exception:
                                pass
                            break
                        waited = int(time.monotonic() - wait_start)
                        if waited > 600:
                            timed_out = True
                            try:
                                proc.terminate()
                            except Exception:
                                pass
                            break
                        self._emit_live_detail(
                            f"{app_prefix} • External uninstaller running ({waited}s elapsed — normal for large/MSI apps)"
                        )
                        threading.Event().wait(0.5)

                    if timed_out:
                        uninstall_returncode = 998
                        self.after(
                            0,
                            lambda n=name: self.log_action(
                                f"Uninstaller timed out after 600s for {n}; continuing force cleanup.", success=False
                            ),
                        )
                    else:
                        uninstall_returncode = proc.returncode if proc.returncode is not None else 999
                        self._emit_progress(
                            f"Uninstaller finished for {name} (exit code {uninstall_returncode})",
                            base_prog + 0.45 / total,
                            detail=f"{app_prefix} • External uninstaller exited with code {uninstall_returncode}",
                        )
                except Exception as e:
                    uninstall_returncode = 999
                    self.after(
                        0,
                        lambda n=name, err=e: self.log_action(
                            f"Uninstaller warning/error for {n}: {str(err)}; continuing force cleanup.", success=False
                        ),
                    )

            if self._cancel_uninstall_flag.is_set():
                break

            if uninstall_returncode not in (None, 0, 1641, 3010):
                self.after(0, lambda n=name, rc=uninstall_returncode: self.log_action(
                    f"Uninstaller for {n} returned exit code {rc}; proceeding with force cleanup of files, shortcuts, and registry.", success=False
                ))

            # 2. Clean Uninstall registry keys
            self._emit_live_detail(f"{app_prefix} • Removing Uninstall registry entries...")
            parent_paths = [
                r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
                r"SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall",
            ]
            for p_path in parent_paths:
                for hk in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
                    try:
                        self._delete_reg_tree(hk, f"{p_path}\\{reg_key_name}")
                    except Exception:
                        pass

            # 3. Clean matching app/publisher keys under SOFTWARE
            self._emit_live_detail(f"{app_prefix} • Scanning and cleaning SOFTWARE registry keys...")
            reg_cleaned = self._delete_matching_software_reg_keys(name, publisher)
            if reg_cleaned:
                self._emit_progress(
                    f"Cleaned {len(reg_cleaned)} registry key(s) for {name}",
                    base_prog + 0.65 / total,
                    detail=f"{app_prefix} • Removed {len(reg_cleaned)} registry tree(s)",
                )

            # 4. Force-remove Desktop & Start Menu shortcuts
            self._emit_live_detail(f"{app_prefix} • Scanning Desktop & Start Menu shortcuts...")
            removed_shortcuts, failed_shortcuts = self._delete_matching_shortcuts(
                name, install_location=loc, publisher=publisher, shortcut_roots=shortcut_roots
            )
            if not self._cancel_uninstall_flag.is_set():
                threading.Event().wait(0.35)
                removed_again, failed_again = self._delete_matching_shortcuts(
                    name, install_location=loc, publisher=publisher, shortcut_roots=shortcut_roots
                )
                removed_shortcuts.extend(path for path in removed_again if path not in removed_shortcuts)
                failed_shortcuts.extend(item for item in failed_again if item not in failed_shortcuts)
            if removed_shortcuts:
                self._emit_progress(
                    f"Removed {len(removed_shortcuts)} shortcut(s) for {name}",
                    base_prog + 0.75 / total,
                    detail=f"{app_prefix} • Removed {len(removed_shortcuts)} shortcut(s)",
                )
            for s_path, s_reason in failed_shortcuts:
                failed_items.append((name, s_path, s_reason))
                self.after(
                    0,
                    lambda n=name, sp=s_path, sr=s_reason: self.log_action(
                        f"Could not delete shortcut ({n}): {sp} -> {sr}", success=False
                    ),
                )

            # 5. Force-remove installation directory
            if loc and os.path.exists(loc):
                self._emit_live_detail(f"{app_prefix} • Force-deleting installation folder: {loc}")
                ok_dir, dir_err = self._force_remove_dir(loc)
                if ok_dir:
                    self._emit_progress(
                        f"Force-deleted install folder: {loc}",
                        base_prog + 0.85 / total,
                        detail=f"{app_prefix} • Force-deleted install directory",
                    )
                else:
                    failed_items.append((name, loc, f"Folder locked/in use: {dir_err}"))
                    self.after(
                        0,
                        lambda n=name, l=loc, err=dir_err: self.log_action(
                            f"Could not delete folder (locked): {l} -> {err}", success=False
                        ),
                    )

            # 6. Force-remove AppData & LocalAppData leftovers
            sanitized_name = re.sub(r"[^a-zA-Z0-9]", "", name).lower()
            for appdata_env in ["APPDATA", "LOCALAPPDATA"]:
                base_appdata = os.environ.get(appdata_env, "")
                if base_appdata and os.path.exists(base_appdata) and len(sanitized_name) > 3:
                    self._emit_live_detail(f"{app_prefix} • Cleaning matching folders in {appdata_env}...")
                    try:
                        for item in os.listdir(base_appdata):
                            clean_item = re.sub(r"[^a-zA-Z0-9]", "", item).lower()
                            if sanitized_name in clean_item:
                                target_dir = os.path.join(base_appdata, item)
                                if os.path.isdir(target_dir):
                                    self._emit_live_detail(f"{app_prefix} • Force-removing {target_dir}...")
                                    ok_ad, _ = self._force_remove_dir(target_dir)
                                    if ok_ad:
                                        self._emit_progress(
                                            f"Cleaned AppData folder: {target_dir}",
                                            base_prog + 0.95 / total,
                                            detail=f"{app_prefix} • Removed {target_dir}",
                                        )
                                    else:
                                        failed_items.append(
                                            (name, target_dir, "AppData folder locked/in use")
                                        )
                    except Exception:
                        pass

            # 7. Complete post-uninstall scan for any remaining leftovers + mini tip why
            self._emit_progress(
                f"Scanning for remaining leftovers of {name}...",
                base_prog + 0.98 / total,
                detail=f"{app_prefix} • Deep-scanning folders, shortcuts, and AppData for leftovers...",
            )
            app_failed_so_far = [item for item in failed_items if item[0] == name]
            other_failed = [item for item in failed_items if item[0] != name]
            scanned_for_app = self._scan_remaining_leftovers_for_app(
                name,
                install_location=loc,
                publisher=publisher,
                existing_failed=app_failed_so_far,
                shortcut_roots=shortcut_roots,
                app_prefix=app_prefix,
            )
            failed_items = other_failed + scanned_for_app
            if scanned_for_app:
                self._emit_progress(
                    f"Found {len(scanned_for_app)} leftover item(s) for {name} (see manual locator popup).",
                    base_prog + 1.0 / total,
                    detail=f"{app_prefix} • Leftover scan finished: {len(scanned_for_app)} item(s) flagged",
                )
            else:
                self._emit_live_detail(f"{app_prefix} • Leftover scan finished: 0 leftovers found")

            self._uninstalled_app_names.add(name.casefold())
            if reg_key_name:
                self._uninstalled_reg_keys.add(reg_key_name.casefold())
            self.after(0, self._sync_apps_list_now)
            success_count += 1

        cancelled = self._cancel_uninstall_flag.is_set()
        self.after(0, lambda: self._finish_uninstall_ui(success_count, failed_items, cancelled))

    def _build_leftover_tip(self, kind, reason=""):
        r_low = (reason or "").lower()
        if kind == "shortcut":
            if "lock" in r_low or "protect" in r_low or "permission" in r_low:
                return "Tip: Shortcut (.lnk/.url) is protected or synced by OneDrive/Windows—open in Explorer to delete manually if the app is gone."
            return "Tip: Leftover shortcut (.lnk/.url) still exists—safe to delete manually if no longer needed."
        if kind == "appdata":
            if "lock" in r_low or "use" in r_low:
                return "Tip: AppData folder may be in use by a background process or holds user settings/cache—close app or inspect before deleting."
            return "Tip: Holds user settings, profiles, or cache—delete to wipe config, or keep if reinstalling later."
        if "lock" in r_low or "use" in r_low or "permission" in r_low or "denied" in r_low:
            return "Tip: File/folder is locked by a running process, tray icon, or Windows permission—close related tasks or open in Explorer to delete."
        return "Tip: Leftover program files remained after uninstaller—open in Explorer to inspect and delete manually."

    def _scan_remaining_leftovers_for_app(
        self,
        app_name,
        install_location="",
        publisher="",
        existing_failed=None,
        shortcut_roots=None,
        app_prefix="",
    ):
        found = []
        seen_paths = set()
        prefix = f"{app_prefix} • " if app_prefix else ""

        if existing_failed:
            for a_n, p_val, r_val in existing_failed:
                norm_p = os.path.normpath(p_val).lower() if p_val else ""
                kind = "shortcut" if p_val and p_val.lower().endswith((".lnk", ".url")) else (
                    "appdata" if p_val and ("appdata" in p_val.lower()) else "folder"
                )
                tip = self._build_leftover_tip(kind, r_val)
                found.append((a_n or app_name, p_val, r_val, tip))
                if norm_p:
                    seen_paths.add(norm_p)

        # 1. Check install location
        if install_location and os.path.exists(install_location):
            self._emit_live_detail(f"{prefix}Checking install directory: {install_location}")
            norm_loc = os.path.normpath(install_location).lower()
            if norm_loc not in seen_paths:
                seen_paths.add(norm_loc)
                found.append((
                    app_name,
                    install_location,
                    "Install directory still exists",
                    self._build_leftover_tip("folder", "Install directory still exists"),
                ))

        # 2. Check leftover shortcuts (.lnk / .url)
        tokens = []
        s_name = re.sub(r"[^a-zA-Z0-9]", "", app_name or "").lower()
        if len(s_name) >= 3:
            tokens.append(s_name)
        if publisher and publisher.lower() != "unknown publisher":
            s_pub = re.sub(r"[^a-zA-Z0-9]", "", publisher).lower()
            if len(s_pub) >= 4 and s_pub not in ("microsoft", "microsoftcorporation"):
                tokens.append(s_pub)

        norm_loc_cmp = os.path.normpath(install_location).lower() if install_location else ""
        roots_to_scan = shortcut_roots if shortcut_roots is not None else self._collect_shortcut_search_roots()
        checked_shortcuts = 0
        for root_dir in roots_to_scan:
            for dirpath, _, filenames in os.walk(root_dir):
                for fname in filenames:
                    if not fname.lower().endswith((".lnk", ".url")):
                        continue
                    checked_shortcuts += 1
                    if checked_shortcuts % 25 == 0:
                        self._emit_live_detail(f"{prefix}Leftover shortcut check ({checked_shortcuts} files): {fname}")
                    fpath = os.path.join(dirpath, fname)
                    norm_f = os.path.normpath(fpath).lower()
                    if norm_f in seen_paths:
                        continue
                    clean_fname = re.sub(r"[^a-zA-Z0-9]", "", os.path.splitext(fname)[0]).lower()
                    matched = any(tok in clean_fname for tok in tokens)
                    if not matched and (norm_loc_cmp or tokens):
                        target_s, workdir_s = self._get_shortcut_target_info(fpath)
                        clean_target = re.sub(r"[^a-zA-Z0-9]", "", target_s).lower()
                        clean_work = re.sub(r"[^a-zA-Z0-9]", "", workdir_s).lower()
                        if norm_loc_cmp and (norm_loc_cmp in target_s or norm_loc_cmp in workdir_s):
                            matched = True
                        elif any(tok in clean_target or tok in clean_work for tok in tokens):
                            matched = True
                    if matched:
                        seen_paths.add(norm_f)
                        found.append((
                            app_name,
                            fpath,
                            "Leftover shortcut",
                            self._build_leftover_tip("shortcut", "Leftover shortcut"),
                        ))

        # 3. Check AppData / LocalAppData leftovers
        if len(s_name) > 3:
            for appdata_env in ("APPDATA", "LOCALAPPDATA"):
                base_appdata = os.environ.get(appdata_env, "")
                if base_appdata and os.path.isdir(base_appdata):
                    self._emit_live_detail(f"{prefix}Scanning {appdata_env} for leftover folders...")
                    try:
                        for item in os.listdir(base_appdata):
                            clean_item = re.sub(r"[^a-zA-Z0-9]", "", item).lower()
                            if s_name in clean_item:
                                target_dir = os.path.join(base_appdata, item)
                                norm_ad = os.path.normpath(target_dir).lower()
                                if norm_ad not in seen_paths and os.path.exists(target_dir):
                                    seen_paths.add(norm_ad)
                                    found.append((
                                        app_name,
                                        target_dir,
                                        f"Leftover {appdata_env} data",
                                        self._build_leftover_tip("appdata", "Leftover data"),
                                    ))
                    except Exception:
                        pass

        return found

    def _open_file_browser_dialog(self, initial_path=""):
        start_dir = ""
        if initial_path and os.path.exists(initial_path):
            start_dir = initial_path if os.path.isdir(initial_path) else os.path.dirname(initial_path)
        if not start_dir or not os.path.isdir(start_dir):
            start_dir = os.environ.get("USERPROFILE") or os.path.expanduser("~")

        chosen = filedialog.askopenfilename(
            title="Locate Leftover File or Shortcut to Delete Manually",
            initialdir=start_dir,
            filetypes=[
                ("All Leftover Items", "*.*"),
                ("Shortcuts (.lnk, .url)", "*.lnk;*.url"),
                ("Executables (.exe)", "*.exe"),
            ],
        )
        if chosen:
            self.browse_path_in_explorer(chosen)

    def _show_manual_cleanup_dialog(self, success_count, failed_items, cancelled, on_close_callback=None):
        popup = ctk.CTkToplevel(self)
        popup.title("Post-Uninstall Leftover Scan & Manual Locator")
        popup.geometry("680x470")
        popup.minsize(520, 350)
        popup.transient(self)

        theme = self.themes.get(self.current_theme_name, self.themes["Dark Slate"])
        popup.configure(fg_color=theme["bg"])

        notified = {"done": False}

        def _notify_once():
            if not notified["done"]:
                notified["done"] = True
                if callable(on_close_callback):
                    self.after(0, on_close_callback)

        def _close_popup():
            _notify_once()
            if popup.winfo_exists():
                popup.destroy()

        popup.protocol("WM_DELETE_WINDOW", _close_popup)
        popup.bind("<Destroy>", lambda e: _notify_once() if e.widget is popup else None, add="+")

        card = ctk.CTkFrame(
            popup,
            fg_color=theme["card"],
            corner_radius=12,
            border_width=1,
            border_color="#f59e0b",
        )
        card.pack(fill="both", expand=True, padx=14, pady=14)

        status_word = "Cancelled after processing" if cancelled else "Complete uninstall finished"
        lbl_head = ctk.CTkLabel(
            card,
            text=f"🔎 {status_word} ({success_count} app(s)) — {len(failed_items)} leftover item(s) found",
            font=ctk.CTkFont(size=15, weight="bold"),
            text_color="#f59e0b",
            wraplength=620,
            justify="left",
        )
        lbl_head.pack(anchor="w", padx=16, pady=(14, 4))

        lbl_note = ctk.CTkLabel(
            card,
            text="Select a leftover item below to open it in Explorer, or use 'Browse File Dialog' to locate any remaining shortcut/file manually:",
            font=ctk.CTkFont(size=12),
            text_color=theme["text"],
            wraplength=620,
            justify="left",
        )
        lbl_note.pack(anchor="w", padx=16, pady=(0, 8))

        listbox_frame = ctk.CTkScrollableFrame(card, fg_color=theme["fg"], height=210)
        listbox_frame.pack(fill="both", expand=True, padx=16, pady=(0, 12))

        first_item = failed_items[0] if failed_items else ("", "", "", "")
        selected_path_var = tk.StringVar(value=first_item[1])

        for entry in failed_items:
            app_name = entry[0]
            item_path = entry[1]
            reason = entry[2]
            tip_line = entry[3] if len(entry) > 3 else self._build_leftover_tip("folder", reason)

            row = ctk.CTkFrame(
                listbox_frame,
                fg_color=theme["row"],
                border_width=1,
                border_color=theme["border"],
                corner_radius=8,
            )
            row.pack(fill="x", padx=4, pady=4)

            rb = ctk.CTkRadioButton(
                row,
                text=f"{app_name}  →  {item_path}  [{reason}]",
                variable=selected_path_var,
                value=item_path,
                font=ctk.CTkFont(size=12, weight="bold"),
                text_color=theme["text"],
                fg_color=theme["accent"],
                hover_color=theme["accent_hover"],
            )
            rb.pack(anchor="w", padx=10, pady=(8, 2))

            lbl_tip = ctk.CTkLabel(
                row,
                text=f"💡 {tip_line}",
                font=ctk.CTkFont(size=11),
                text_color=theme["subtext"],
                wraplength=580,
                justify="left",
            )
            lbl_tip.pack(anchor="w", padx=32, pady=(0, 8))

        btn_row = ctk.CTkFrame(card, fg_color="transparent")
        btn_row.pack(pady=(0, 14))

        btn_browse = ctk.CTkButton(
            btn_row,
            text="📂 Locate in Explorer",
            width=155,
            height=36,
            fg_color=theme["accent"],
            hover_color=theme["accent_hover"],
            text_color=theme["accent_text"],
            font=ctk.CTkFont(size=12, weight="bold"),
            command=lambda: self.browse_path_in_explorer(selected_path_var.get()),
        )
        btn_browse.pack(side="left", padx=5)

        btn_file_picker = ctk.CTkButton(
            btn_row,
            text="🗂️ Browse File Dialog",
            width=155,
            height=36,
            fg_color=theme["row"],
            hover_color=theme["row_hover"],
            text_color=theme["text"],
            border_width=1,
            border_color=theme["border"],
            font=ctk.CTkFont(size=12, weight="bold"),
            command=lambda: self._open_file_browser_dialog(selected_path_var.get()),
        )
        btn_file_picker.pack(side="left", padx=5)

        btn_close = ctk.CTkButton(
            btn_row,
            text="✔ Done",
            width=110,
            height=36,
            fg_color="#16a34a",
            hover_color="#15803d",
            text_color="#ffffff",
            font=ctk.CTkFont(size=12, weight="bold"),
            command=_close_popup,
        )
        btn_close.pack(side="left", padx=5)

    def _prompt_post_uninstall_clean_junk(self, success_count, cancelled=False, on_close_callback=None):
        status_title = "Uninstallation Cancelled" if cancelled else "Complete Uninstall Done!"
        status_line = (
            f"Stopped after processing {success_count} app(s)."
            if cancelled
            else f"Successfully uninstalled and deep-cleaned {success_count} application(s)!"
        )
        ask_junk = messagebox.askyesno(
            f"🧹 {status_title} — Clean Junk Now?",
            f"{status_line}\n\n"
            "Would you like to run Clean JUNK now?\n\n"
            "💡 Tip: Running Clean JUNK after uninstalling helps remove:\n"
            "  • Leftover, empty, or broken (.lnk / .url) shortcuts\n"
            "  • Installer temporary files & leftover setup caches\n"
            "  • Crash dumps & shader cache files",
        )
        if ask_junk:
            self.clean_junk_temp_cache(on_complete_callback=on_close_callback)
        elif callable(on_close_callback):
            on_close_callback()

    def _on_manual_cleanup_closed(self, manual_token):
        self._release_uninstall_popup(manual_token)
        self.after(200, self._launch_clean_junk_prompt)

    def _launch_clean_junk_prompt(self):
        junk_token = "post_uninstall_clean_junk"
        self._register_uninstall_popup(junk_token)
        self._prompt_post_uninstall_clean_junk(
            self._last_uninstall_success_count,
            self._last_uninstall_cancelled,
            on_close_callback=lambda tok=junk_token: self._on_junk_prompt_closed(tok),
        )

    def _on_junk_prompt_closed(self, junk_token):
        self._release_uninstall_popup(junk_token)
        if not self._pending_uninstall_popups and self._auto_refresh_after_popups:
            self._auto_refresh_after_popups = False
            self.after(50, self._run_post_uninstall_auto_refresh)

    def _finish_uninstall_ui(self, success_count, failed_items, cancelled=False):
        self._last_uninstall_success_count = success_count
        self._last_uninstall_cancelled = cancelled
        self._sync_apps_list_now()

        if self._active_progress_popup is not None and self._active_progress_popup.winfo_exists():
            if cancelled:
                self._active_progress_popup.mark_cancelled()
            else:
                self._active_progress_popup.mark_completed(
                    f"Completed uninstall for {success_count} app(s)."
                )

        def _after_progress_closed():
            self._sync_apps_list_now()
            if failed_items:
                manual_token = "manual_cleanup_popup"
                self._register_uninstall_popup(manual_token)
                self._show_manual_cleanup_dialog(
                    success_count,
                    failed_items,
                    cancelled,
                    on_close_callback=lambda tok=manual_token: self._on_manual_cleanup_closed(tok),
                )
            else:
                self.after(150, self._launch_clean_junk_prompt)

        if self._active_progress_popup is not None and self._active_progress_popup.winfo_exists():
            orig_close = self._active_progress_popup.on_close_callback

            def _wrapped_progress_close():
                if callable(orig_close):
                    orig_close()
                _after_progress_closed()

            self._active_progress_popup.on_close_callback = _wrapped_progress_close
        else:
            _after_progress_closed()


if __name__ == "__main__":
    if os.name == "nt" and not is_admin():
        if relaunch_as_admin():
            sys.exit(0)

    app = JustUninstallApps()
    app.mainloop()