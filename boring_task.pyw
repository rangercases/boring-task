import os
import sys
import re
import datetime
import threading
import pickle
import json
import subprocess
import urllib.request
import tempfile
import time
from copy import copy
import ctypes
import zipfile

import customtkinter as ctk
from tkinter import filedialog, messagebox
import tkinterdnd2 as tkdnd
import openpyxl
from openpyxl.utils import get_column_letter
from PIL import Image

# Set Windows App User Model ID so Taskbar groups and shows custom icon
try:
    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("rangercases.boringtask.v1")
except Exception:
    pass

# Set Light Mode
ctk.set_appearance_mode("Light")

APP_VERSION = "v1.0"
GITHUB_REPO = "rangercases/claim-helper"
CACHE_FILE_NAME = ".overview_cache.pkl"
STATE_FILE_NAME = ".app_state.json"

# ----------------------------------------------------
# THEME: "Soft Ivory Bakery" (Japanese Cafe Style)
# ----------------------------------------------------
IVORY = "#f7f2e8"        # nền cửa sổ kem ngà
CARD = "#fffdf8"         # nền card trắng ấm nhẹ
LINE = "#e8e0d0"         # viền mảnh tao nhã
SAND = "#f1eadb"         # vùng kéo thả, ô kết quả
ROAST = "#2b211c"        # nâu rang rất đậm: chữ chính, nút chính
BROWN = "#6f5d50"        # chữ phụ ấm áp
FAINT = "#b3a898"        # chữ mờ, disabled, footer
MOSS = "#6b7a5a"         # xanh rêu: nhấn chính (progress, focus, thành công)
MOSS_SOFT = "#e4e8d8"    # nền nhạt của trạng thái thành công, dragover
CARAMEL = "#c8855a"      # nhấn phụ rất hạn chế (icon nhỏ)
DANGER = "#a8503a"       # đỏ gạch dịu cho Xóa/lỗi

FONT_SERIF = "Cambria"   # Font Serif chuẩn Windows, hỗ trợ 100% tiếng Việt
FONT_SANS = "Segoe UI"   # Font Sans-serif chuẩn Windows, sắc nét, không lỗi dấu

def parse_version(v_str):
    clean = re.sub(r'[^0-9.]', '', str(v_str))
    parts = []
    for p in clean.split('.'):
        if p.isdigit():
            parts.append(int(p))
    return parts

def is_newer_version(latest_tag, current_tag):
    v_latest = parse_version(latest_tag)
    v_curr = parse_version(current_tag)
    return v_latest > v_curr

def get_resource_path(relative_path):
    """Get absolute path to resource, works for dev and for PyInstaller bundle."""
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)

def get_cache_path():
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), CACHE_FILE_NAME)

def get_state_path():
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), STATE_FILE_NAME)

def parse_drop_paths(data_str):
    """Parse dropped file paths handling curly braces and spaces."""
    if not data_str:
        return []
    pattern = r'\{([^}]+)\}|(\S+)'
    matches = re.findall(pattern, data_str.strip())
    paths = []
    for m in matches:
        p = (m[0] if m[0] else m[1]).strip().strip('"').strip("'")
        if os.path.exists(p):
            paths.append(os.path.abspath(p))
    return paths

def format_file_size(size_bytes):
    if size_bytes < 1024:
        return f"{size_bytes} B"
    elif size_bytes < 1024 * 1024:
        return f"{size_bytes / 1024:.1f} KB"
    else:
        return f"{size_bytes / (1024 * 1024):.1f} MB"

def reveal_in_explorer(filepath):
    """Opens Windows Explorer with the specific file selected/highlighted."""
    try:
        norm_path = os.path.normpath(filepath)
        if os.path.exists(norm_path):
            subprocess.Popen(f'explorer /select,"{norm_path}"')
        else:
            folder = os.path.dirname(norm_path)
            if os.path.exists(folder):
                subprocess.Popen(f'explorer "{folder}"')
    except Exception as e:
        print(f"Error revealing in explorer: {e}")

def validate_excel_file(filepath):
    """Validates if file exists, non-empty, and has valid ZIP/XLSX signature."""
    fname = os.path.basename(filepath)
    if not os.path.exists(filepath):
        raise FileNotFoundError(f"Không tìm thấy file '{fname}'.")
    if os.path.getsize(filepath) == 0:
        raise ValueError(f"File '{fname}' bị rỗng (0 KB), có thể do tải dở từ OneDrive/Email.")
    
    # Check magic bytes for ZIP (.xlsx is a zip package)
    try:
        with open(filepath, 'rb') as f:
            header = f.read(4)
        if header != b'PK\x03\x04':
            if header.startswith(b'\xd0\xcf\x11\xe0'):
                raise ValueError(f"File '{fname}' là định dạng Excel cũ (.xls) bị đổi đuôi. Vui lòng mở bằng Excel và Save As thành Excel Workbook (*.xlsx).")
            elif header.startswith(b'<html') or header.startswith(b'<!DOC') or header.startswith(b'<?xml'):
                raise ValueError(f"File '{fname}' là file web (HTML/XML) bị đổi đuôi thành .xlsx. Vui lòng mở bằng Excel và Save As lại.")
            else:
                raise ValueError(f"File '{fname}' không phải định dạng Excel (.xlsx) chuẩn hoặc bị hỏng dữ liệu.")
    except (PermissionError, ValueError, FileNotFoundError):
        raise
    except Exception as e:
        raise ValueError(f"Không thể đọc file '{fname}': {e}")

class BoringTaskApp(ctk.CTk, tkdnd.TkinterDnD.DnDWrapper):
    def __init__(self):
        super().__init__()
        self.TkdndVersion = tkdnd.TkinterDnD._require(self)

        self.title("Boring Task")
        self.geometry("860x780")
        self.minsize(800, 700)
        self.configure(fg_color=IVORY)  # Soft Ivory Bakery background

        # Set Window & Taskbar Icon
        ico_path = get_resource_path(os.path.join("assets", "app_icon.ico"))
        if os.path.exists(ico_path):
            try:
                self.iconbitmap(ico_path)
            except Exception:
                pass

        # State Variables
        self.overview_path = ""
        self.claim_files = [] # list of absolute paths
        self.last_created_files = [] # list of created output files

        curr_dir = os.path.abspath(os.path.dirname(__file__))
        saved_state = self.load_saved_state()

        if saved_state is not None:
            # Restore saved overview path if it still exists
            saved_ov = saved_state.get("overview_path", "")
            if saved_ov and os.path.exists(saved_ov):
                self.overview_path = saved_ov
            
            # Restore saved claim files (keep only existing files)
            saved_claims = saved_state.get("claim_files", [])
            self.claim_files = [p for p in saved_claims if os.path.exists(p)]
        else:
            # First launch: Auto-detect initial versus files in current directory
            for f in os.listdir(curr_dir):
                if "versus" in f.lower() and f.endswith(".xlsx") and not f.endswith("_backup.xlsx") and not f.endswith("_filled.xlsx") and not f.startswith("~$"):
                    self.claim_files.append(os.path.join(curr_dir, f))
            
            self.save_state()

        # If overview not found in saved state, fallback to auto-detecting in current dir
        if not self.overview_path:
            for f in os.listdir(curr_dir):
                if f.lower().startswith("overview") and f.endswith(".xlsx") and not f.startswith("~$"):
                    self.overview_path = os.path.join(curr_dir, f)
                    break

        self.setup_ui()
        self.setup_drag_and_drop()
        self.apply_windows_titlebar_theme()

        # Check for remote updates silently in background
        threading.Thread(target=self.silent_auto_update, daemon=True).start()

    def apply_windows_titlebar_theme(self):
        """Customizes Windows title bar color to match IVORY & ROAST theme."""
        try:
            self.update_idletasks()
            hwnd = ctypes.windll.user32.GetParent(self.winfo_id())
            if not hwnd:
                hwnd = self.winfo_id()

            # DWM attributes for Windows 10 (20H1+) and Windows 11
            DWMWA_CAPTION_COLOR = 35
            DWMWA_TEXT_COLOR = 36

            # IVORY is #f7f2e8 -> BGR is 0x00E8F2F7
            caption_color = ctypes.c_int(0x00E8F2F7)
            # ROAST is #2b211c -> BGR is 0x001C212B
            text_color = ctypes.c_int(0x001C212B)

            ctypes.windll.dwmapi.DwmSetWindowAttribute(
                hwnd,
                DWMWA_CAPTION_COLOR,
                ctypes.byref(caption_color),
                ctypes.sizeof(caption_color)
            )
            ctypes.windll.dwmapi.DwmSetWindowAttribute(
                hwnd,
                DWMWA_TEXT_COLOR,
                ctypes.byref(text_color),
                ctypes.sizeof(text_color)
            )
        except Exception:
            pass

    def load_saved_state(self):
        """Loads persistent session state if available."""
        s_path = get_state_path()
        if not os.path.exists(s_path):
            return None
        try:
            with open(s_path, "r", encoding="utf-8") as f:
                data = json.load(f)
            return data
        except Exception as e:
            print(f"Error loading state: {e}")
            return None

    def save_state(self):
        """Saves current state to persistent storage."""
        s_path = get_state_path()
        try:
            data = {
                "overview_path": self.overview_path,
                "claim_files": self.claim_files
            }
            with open(s_path, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"Error saving state: {e}")

    # ----------------------------------------------------
    # Silent Lightning-Fast Background Auto-Updater
    # ----------------------------------------------------
    def silent_auto_update(self):
        """Runs silently in background. Checks GitHub Releases / raw code in 1-2s.
        If a newer release exists, silently downloads and overwrites local code in <1s.
        """
        try:
            url = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"
            req = urllib.request.Request(url, headers={"User-Agent": "BoringTask-App"})
            with urllib.request.urlopen(req, timeout=3) as resp:
                if resp.status == 200:
                    data = json.loads(resp.read().decode("utf-8"))
                    latest_tag = data.get("tag_name", "")
                    if latest_tag and is_newer_version(latest_tag, APP_VERSION):
                        raw_url = f"https://raw.githubusercontent.com/{GITHUB_REPO}/main/boring_task.pyw"
                        req_raw = urllib.request.Request(raw_url, headers={"User-Agent": "BoringTask-App"})
                        with urllib.request.urlopen(req_raw, timeout=5) as raw_resp:
                            if raw_resp.status == 200:
                                new_code = raw_resp.read()
                                if len(new_code) > 1000 and (b"BoringTask" in new_code or b"ClaimHelper" in new_code):
                                    curr_file = os.path.abspath(__file__)
                                    tmp_file = curr_file + ".new"
                                    with open(tmp_file, "wb") as f:
                                        f.write(new_code)
                                    os.replace(tmp_file, curr_file)
                                    self.after(0, self.set_update_badge, f"✓ Đã tự động cập nhật {latest_tag}")
        except Exception:
            pass

    def set_update_badge(self, text):
        if hasattr(self, 'badge_update'):
            self.badge_update.configure(text=f"  {text}  ", text_color=ROAST, fg_color=MOSS_SOFT)

    def setup_ui(self):
        # 1. Header Frame (Japanese Cafe aesthetic: airy, light, serif, brown tone)
        header = ctk.CTkFrame(self, fg_color="transparent")
        header.pack(fill="x", padx=36, pady=(20, 12))

        title_row = ctk.CTkFrame(header, fg_color="transparent")
        title_row.pack(anchor="w")

        # Boring Task Brand Logo Banner
        banner_path = get_resource_path(os.path.join("assets", "boring_task_banner.png"))
        if os.path.exists(banner_path):
            try:
                pil_banner = Image.open(banner_path)
                # Aspect ratio is 1024 / 132 = ~7.75. Display at (225, 29)
                self.header_logo_img = ctk.CTkImage(light_image=pil_banner, dark_image=pil_banner, size=(225, 29))
                banner_lbl = ctk.CTkLabel(title_row, text="", image=self.header_logo_img)
                banner_lbl.pack(side="left", padx=(0, 12))
            except Exception:
                pass
        else:
            title_lbl = ctk.CTkLabel(
                title_row,
                text="Boring Task",
                font=ctk.CTkFont(family=FONT_SERIF, size=28, weight="bold"),
                text_color=ROAST
            )
            title_lbl.pack(side="left", padx=(0, 12))

        # Pill version badge
        badge_lbl = ctk.CTkLabel(
            title_row,
            text=f"  {APP_VERSION}  ",
            font=ctk.CTkFont(family=FONT_SANS, size=10),
            fg_color="transparent",
            text_color=BROWN,
            corner_radius=10
        )
        badge_lbl.pack(side="left", pady=(6, 0))

        # Silent update status badge
        self.badge_update = ctk.CTkLabel(
            title_row,
            text="",
            font=ctk.CTkFont(family=FONT_SANS, size=10),
            fg_color="transparent",
            text_color=MOSS,
            corner_radius=10
        )
        self.badge_update.pack(side="left", padx=(10, 0), pady=(6, 0))

        # Dynamic Update Banner Container
        self.update_banner_container = ctk.CTkFrame(self, fg_color="transparent")

        # View container: hosts Home screen and feature screens (switched via pack/pack_forget)
        self.view_container = ctk.CTkFrame(self, fg_color="transparent")
        self.view_container.pack(fill="both", expand=True)

        self.home_view = ctk.CTkFrame(self.view_container, fg_color="transparent")
        self.claim_view = ctk.CTkFrame(self.view_container, fg_color="transparent")

        self.build_home()
        self.build_feature_nav(self.claim_view, "Purchase Cost Auto-Filled")

        # 2. Main Scrollable Container (Generous whitespace)
        self.main_scroll = ctk.CTkScrollableFrame(self.claim_view, fg_color="transparent")
        self.main_scroll.pack(fill="both", expand=True, padx=36, pady=(0, 16))

        # ========================================================
        # CARD 1: Master Overview (CARD background, 16px radius, LINE border)
        # ========================================================
        self.card_ov = ctk.CTkFrame(
            self.main_scroll,
            fg_color=CARD,
            corner_radius=16,
            border_width=1,
            border_color=LINE
        )
        self.card_ov.pack(fill="x", pady=(0, 20))

        ov_top = ctk.CTkFrame(self.card_ov, fg_color="transparent")
        ov_top.pack(fill="x", padx=24, pady=(18, 10))

        ctk.CTkLabel(
            ov_top,
            text="FILE MASTER OVERVIEW",
            font=ctk.CTkFont(family=FONT_SANS, size=11, weight="bold"),
            text_color=BROWN
        ).pack(side="left")

        ov_btn_box = ctk.CTkFrame(ov_top, fg_color="transparent")
        ov_btn_box.pack(side="right")

        # Pill button with LINE border
        self.btn_pick_ov = ctk.CTkButton(
            ov_btn_box,
            text="+ Chọn File Overview...",
            font=ctk.CTkFont(family=FONT_SANS, size=11),
            fg_color="transparent",
            text_color=ROAST,
            hover_color=SAND,
            border_width=1,
            border_color=LINE,
            corner_radius=14,
            width=165,
            height=28,
            command=self.browse_overview
        )
        self.btn_pick_ov.pack(side="left", padx=4)

        self.btn_refresh_cache = ctk.CTkButton(
            ov_btn_box,
            text="Đọc Lại Master",
            font=ctk.CTkFont(family=FONT_SANS, size=11),
            fg_color="transparent",
            text_color=ROAST,
            hover_color=SAND,
            border_width=1,
            border_color=LINE,
            corner_radius=14,
            width=120,
            height=28,
            command=self.force_reload_master
        )
        self.btn_refresh_cache.pack(side="left", padx=4)

        self.btn_clear_ov = ctk.CTkButton(
            ov_btn_box,
            text="Xóa File",
            font=ctk.CTkFont(family=FONT_SANS, size=11),
            fg_color="transparent",
            text_color=DANGER,
            hover_color=SAND,
            border_width=0,
            corner_radius=14,
            width=70,
            height=28,
            command=self.clear_overview
        )
        self.btn_clear_ov.pack(side="left", padx=4)

        # Hairline divider
        ctk.CTkFrame(self.card_ov, height=1, fg_color=LINE).pack(fill="x", padx=24, pady=(0, 14))

        # Drop Zone / Display container (SAND background, FAINT border, 12px radius)
        self.ov_display_container = ctk.CTkFrame(
            self.card_ov,
            fg_color=SAND,
            corner_radius=12,
            border_width=1,
            border_color=FAINT
        )
        self.ov_display_container.pack(fill="x", padx=24, pady=(0, 20))

        self.render_overview_display()

        # ========================================================
        # CARD 2: Claim Versus Files (CARD background, 16px radius)
        # ========================================================
        self.card_claims = ctk.CTkFrame(
            self.main_scroll,
            fg_color=CARD,
            corner_radius=16,
            border_width=1,
            border_color=LINE
        )
        self.card_claims.pack(fill="x", pady=(0, 20))

        claim_top = ctk.CTkFrame(self.card_claims, fg_color="transparent")
        claim_top.pack(fill="x", padx=24, pady=(18, 10))

        self.claim_count_lbl = ctk.CTkLabel(
            claim_top,
            text=f"DANH SÁCH FILE CLAIM VERSUS ({len(self.claim_files)})",
            font=ctk.CTkFont(family=FONT_SANS, size=11, weight="bold"),
            text_color=BROWN
        )
        self.claim_count_lbl.pack(side="left")

        actions_box = ctk.CTkFrame(claim_top, fg_color="transparent")
        actions_box.pack(side="right")

        ctk.CTkButton(
            actions_box,
            text="+ Chọn File(s)...",
            font=ctk.CTkFont(family=FONT_SANS, size=11),
            fg_color="transparent",
            text_color=ROAST,
            hover_color=SAND,
            border_width=1,
            border_color=LINE,
            corner_radius=14,
            width=165,
            height=28,
            command=self.browse_claim_files
        ).pack(side="left", padx=4)

        ctk.CTkButton(
            actions_box,
            text="📂 Thư Mục...",
            font=ctk.CTkFont(family=FONT_SANS, size=11),
            fg_color="transparent",
            text_color=ROAST,
            hover_color=SAND,
            border_width=1,
            border_color=LINE,
            corner_radius=14,
            width=120,
            height=28,
            command=self.browse_claim_folder
        ).pack(side="left", padx=4)

        ctk.CTkButton(
            actions_box,
            text="Xóa Hết",
            font=ctk.CTkFont(family=FONT_SANS, size=11),
            fg_color="transparent",
            text_color=DANGER,
            hover_color=SAND,
            border_width=0,
            corner_radius=14,
            width=70,
            height=28,
            command=self.clear_all_claims
        ).pack(side="left", padx=4)

        # Hairline divider
        ctk.CTkFrame(self.card_claims, height=1, fg_color=LINE).pack(fill="x", padx=24, pady=(0, 14))

        # File List Inner Container
        self.file_list_frame = ctk.CTkFrame(
            self.card_claims,
            fg_color=SAND,
            corner_radius=12,
            border_width=1,
            border_color=FAINT
        )
        self.file_list_frame.pack(fill="x", padx=24, pady=(0, 20))

        self.render_file_list()

        # ========================================================
        # ACTION: Start Processing Button (Pill 48px, ROAST, hover MOSS)
        # ========================================================
        act_box = ctk.CTkFrame(self.main_scroll, fg_color="transparent")
        act_box.pack(fill="x", pady=(4, 14))

        self.btn_run = ctk.CTkButton(
            act_box,
            text="Bắt Đầu Đối Chiếu & Điền Giá (Tạo File _filled)",
            font=ctk.CTkFont(family=FONT_SANS, size=13, weight="bold"),
            fg_color=ROAST,
            hover_color=MOSS,
            text_color=IVORY,
            corner_radius=24,
            height=48,
            command=self.start_processing
        )
        self.btn_run.pack(fill="x")

        # Progress Bar (3px high, LINE track, MOSS fill)
        self.prog_bar = ctk.CTkProgressBar(
            self.main_scroll,
            progress_color=MOSS,
            fg_color=LINE,
            height=3,
            corner_radius=2
        )
        self.prog_bar.set(0)
        self.prog_bar.pack(fill="x", pady=(0, 20))

        # ========================================================
        # CARD 3: Activity & Results Card
        # ========================================================
        self.card_results = ctk.CTkFrame(
            self.main_scroll,
            fg_color=CARD,
            corner_radius=16,
            border_width=1,
            border_color=LINE
        )
        self.card_results.pack(fill="both", expand=True, pady=(0, 8))

        results_head = ctk.CTkFrame(self.card_results, fg_color="transparent")
        results_head.pack(fill="x", padx=24, pady=(18, 10))

        ctk.CTkLabel(
            results_head,
            text="TIẾN ĐỘ & KẾT QUẢ",
            font=ctk.CTkFont(family=FONT_SANS, size=11, weight="bold"),
            text_color=BROWN
        ).pack(side="left")

        self.status_badge = ctk.CTkLabel(
            results_head,
            text="Sẵn sàng",
            font=ctk.CTkFont(family=FONT_SANS, size=11),
            text_color=BROWN
        )
        self.status_badge.pack(side="right")

        # Hairline divider
        ctk.CTkFrame(self.card_results, height=1, fg_color=LINE).pack(fill="x", padx=24, pady=(0, 14))

        # Container for concise result items
        self.results_feed = ctk.CTkFrame(
            self.card_results,
            fg_color=SAND,
            corner_radius=12,
            border_width=1,
            border_color=FAINT
        )
        self.results_feed.pack(fill="both", expand=True, padx=24, pady=(0, 20))

        self.render_initial_feed()

        # 3. Bottom Footer (Serif italic, FAINT)
        footer = ctk.CTkFrame(self, fg_color="transparent")
        footer.pack(fill="x", padx=36, pady=(0, 14))

        ctk.CTkLabel(
            footer,
            text="Powered by Quoc Hung",
            font=ctk.CTkFont(family=FONT_SERIF, size=11, slant="italic"),
            text_color=FAINT
        ).pack(side="left")

        self.show_home()

    # ----------------------------------------------------
    # Home Screen & Navigation
    # ----------------------------------------------------
    def build_home(self):
        wrap = ctk.CTkFrame(self.home_view, fg_color="transparent")
        wrap.pack(fill="both", expand=True, padx=36, pady=(18, 16))

        ctk.CTkLabel(
            wrap,
            text="CÔNG CỤ",
            font=ctk.CTkFont(family=FONT_SANS, size=11, weight="bold"),
            text_color=BROWN
        ).pack(anchor="w", pady=(0, 12))

        grid = ctk.CTkFrame(wrap, fg_color="transparent")
        grid.pack(anchor="w")

        # To add a new feature: append one entry here.
        features = [
            {
                "icon": "📋",
                "title": "Purchase Cost",
                "subtitle": "Auto-Filled",
                "command": lambda: self.show_feature(self.claim_view, "Purchase Cost Auto-Filled"),
            },
        ]

        cols = 3
        for i, f in enumerate(features):
            tile = self.make_feature_tile(grid, **f)
            tile.grid(row=i // cols, column=i % cols, padx=(0, 16), pady=(0, 16), sticky="nw")

    def make_feature_tile(self, parent, icon, title, subtitle, command):
        HOVER_BG = "#fbf5ea"

        tile = ctk.CTkFrame(
            parent, width=200, height=168,
            fg_color=CARD, corner_radius=18,
            border_width=1, border_color=LINE
        )
        tile.pack_propagate(False)

        icon_box = ctk.CTkFrame(tile, width=52, height=52, fg_color=SAND, corner_radius=14)
        icon_box.pack_propagate(False)
        icon_box.pack(anchor="w", padx=20, pady=(22, 0))

        icon_lbl = ctk.CTkLabel(
            icon_box, text=icon,
            font=ctk.CTkFont(family="Segoe UI Emoji", size=22),
            text_color=ROAST
        )
        icon_lbl.place(relx=0.5, rely=0.5, anchor="center")

        text_box = ctk.CTkFrame(tile, fg_color="transparent")
        text_box.pack(side="bottom", fill="x", padx=20, pady=(0, 20))

        t1 = ctk.CTkLabel(
            text_box, text=title,
            font=ctk.CTkFont(family=FONT_SERIF, size=16, weight="bold"),
            text_color=ROAST, anchor="w"
        )
        t1.pack(anchor="w")

        t2 = ctk.CTkLabel(
            text_box, text=subtitle,
            font=ctk.CTkFont(family=FONT_SANS, size=12),
            text_color=BROWN, anchor="w"
        )
        t2.pack(anchor="w")

        state = {"hover": False}

        def set_look(bg, border):
            tile.configure(fg_color=bg, border_color=border)

        def pointer_inside():
            try:
                x, y = tile.winfo_pointerxy()
                rx, ry = tile.winfo_rootx(), tile.winfo_rooty()
                return rx <= x < rx + tile.winfo_width() and ry <= y < ry + tile.winfo_height()
            except Exception:
                return False

        def on_enter(_e=None):
            if not state["hover"]:
                state["hover"] = True
                set_look(HOVER_BG, MOSS)

        def check_leave():
            if state["hover"] and not pointer_inside():
                state["hover"] = False
                set_look(CARD, LINE)

        def on_leave(_e=None):
            tile.after(15, check_leave)

        def on_press(_e=None):
            set_look(SAND, MOSS)

        def on_release(_e=None):
            if pointer_inside():
                set_look(HOVER_BG, MOSS)

                def go():
                    state["hover"] = False
                    set_look(CARD, LINE)
                    command()

                tile.after(70, go)
            else:
                state["hover"] = False
                set_look(CARD, LINE)

        for w in (tile, icon_box, icon_lbl, text_box, t1, t2):
            w.bind("<Enter>", on_enter)
            w.bind("<Leave>", on_leave)
            w.bind("<ButtonPress-1>", on_press)
            w.bind("<ButtonRelease-1>", on_release)
            try:
                w.configure(cursor="hand2")
            except Exception:
                pass

        return tile

    def build_feature_nav(self, view, title):
        nav = ctk.CTkFrame(view, fg_color="transparent", height=34)
        nav.pack(fill="x", padx=36, pady=(2, 12))
        nav.pack_propagate(False)

        ctk.CTkButton(
            nav,
            text="‹  Trang chủ",
            font=ctk.CTkFont(family=FONT_SANS, size=12),
            fg_color="transparent",
            hover_color=SAND,
            text_color=MOSS,
            corner_radius=14,
            width=104,
            height=30,
            anchor="w",
            command=self.show_home
        ).pack(side="left")

        ctk.CTkLabel(
            nav,
            text=title,
            font=ctk.CTkFont(family=FONT_SERIF, size=14, weight="bold"),
            text_color=ROAST
        ).place(relx=0.5, rely=0.5, anchor="center")

    def show_home(self):
        self.claim_view.pack_forget()
        self.home_view.pack(fill="both", expand=True)
        self.title("Boring Task")

    def show_feature(self, view, title):
        self.home_view.pack_forget()
        view.pack(fill="both", expand=True)
        self.title(f"Boring Task — {title}")

    # ----------------------------------------------------
    # UI Renderers
    # ----------------------------------------------------
    def render_initial_feed(self):
        for child in self.results_feed.winfo_children():
            child.destroy()
        
        hint = ctk.CTkLabel(
            self.results_feed,
            text="Kết quả sẽ hiển thị tại đây.",
            font=ctk.CTkFont(family=FONT_SANS, size=12),
            text_color=BROWN,
            pady=24
        )
        hint.pack(fill="both", expand=True)

    def render_overview_display(self):
        for child in self.ov_display_container.winfo_children():
            child.destroy()

        if self.overview_path and os.path.exists(self.overview_path):
            self.btn_refresh_cache.configure(state="normal", text_color=ROAST, border_color=LINE)
            self.btn_clear_ov.configure(state="normal", text_color=DANGER)
            size_str = format_file_size(os.path.getsize(self.overview_path))
            
            # Check cache status
            c_path = get_cache_path()
            has_cache = False
            if os.path.exists(c_path):
                try:
                    with open(c_path, 'rb') as f:
                        meta = pickle.load(f)
                    if (meta.get('file_path') == os.path.abspath(self.overview_path) and
                        meta.get('file_size') == os.path.getsize(self.overview_path)):
                        has_cache = True
                except Exception:
                    pass

            row = ctk.CTkFrame(self.ov_display_container, fg_color=CARD, corner_radius=10, border_width=1, border_color=LINE)
            row.pack(fill="x", padx=10, pady=10)

            ctk.CTkLabel(
                row,
                text="📊",
                font=ctk.CTkFont(size=14),
                text_color=CARAMEL
            ).pack(side="left", padx=(12, 8), pady=8)

            info_box = ctk.CTkFrame(row, fg_color="transparent")
            info_box.pack(side="left", fill="x", expand=True, pady=6)

            title_row = ctk.CTkFrame(info_box, fg_color="transparent")
            title_row.pack(anchor="w")

            ctk.CTkLabel(
                title_row,
                text=os.path.basename(self.overview_path),
                font=ctk.CTkFont(family=FONT_SERIF, size=12, weight="bold"),
                text_color=ROAST
            ).pack(side="left")

            if has_cache:
                ctk.CTkLabel(
                    title_row,
                    text="  ✓ ĐÃ GHI NHỚ  ",
                    font=ctk.CTkFont(family=FONT_SANS, size=9, weight="bold"),
                    fg_color=MOSS_SOFT,
                    text_color=MOSS,
                    corner_radius=6
                ).pack(side="left", padx=(8, 0))

            ctk.CTkLabel(
                info_box,
                text=f"{self.overview_path}  •  {size_str}",
                font=ctk.CTkFont(family=FONT_SANS, size=10),
                text_color=BROWN,
                anchor="w"
            ).pack(anchor="w", pady=(2, 0))

            # Inline Delete Button
            ctk.CTkButton(
                row,
                text="✕",
                font=ctk.CTkFont(size=11, weight="bold"),
                fg_color="transparent",
                text_color=BROWN,
                hover_color=SAND,
                border_width=0,
                corner_radius=12,
                width=24,
                height=24,
                command=self.clear_overview
            ).pack(side="right", padx=(4, 12), pady=8)
        else:
            self.btn_refresh_cache.configure(state="disabled", text_color=FAINT, border_color=LINE)
            self.btn_clear_ov.configure(state="disabled", text_color=FAINT)
            drop_hint = ctk.CTkLabel(
                self.ov_display_container,
                text="📥  Kéo thả file Master Overview vào đây",
                font=ctk.CTkFont(family=FONT_SANS, size=12),
                text_color=BROWN,
                pady=20
            )
            drop_hint.pack(fill="both", expand=True)

    def clear_overview(self):
        self.overview_path = ""
        self.save_state()
        self.render_overview_display()

    def render_file_list(self):
        for child in self.file_list_frame.winfo_children():
            child.destroy()

        if not self.claim_files:
            empty_lbl = ctk.CTkLabel(
                self.file_list_frame,
                text="📥  Kéo thả file Claim Versus vào đây",
                font=ctk.CTkFont(family=FONT_SANS, size=12),
                text_color=BROWN,
                pady=24
            )
            empty_lbl.pack(fill="both", expand=True)
            self.claim_count_lbl.configure(text="DANH SÁCH FILE CLAIM VERSUS (0)")
            return

        self.claim_count_lbl.configure(text=f"DANH SÁCH FILE CLAIM VERSUS ({len(self.claim_files)})")

        for fpath in self.claim_files:
            fname = os.path.basename(fpath)
            is_casa = "casa" in fname.lower()
            tag_text = "Casa (CS)" if is_casa else "Nhan Hoang (NH)"
            tag_color = SAND
            tag_text_color = BROWN

            base, ext = os.path.splitext(fname)
            out_preview = f"{base}_filled{ext}" if not base.endswith("_filled") else fname

            row = ctk.CTkFrame(self.file_list_frame, fg_color=CARD, corner_radius=10, border_width=1, border_color=LINE)
            row.pack(fill="x", padx=10, pady=4)

            # Icon & Name
            ctk.CTkLabel(
                row,
                text="📄",
                font=ctk.CTkFont(size=13),
                text_color=CARAMEL
            ).pack(side="left", padx=(12, 6), pady=6)

            info = ctk.CTkFrame(row, fg_color="transparent")
            info.pack(side="left", fill="x", expand=True, pady=4)

            ctk.CTkLabel(
                info,
                text=fname,
                font=ctk.CTkFont(family=FONT_SERIF, size=12, weight="bold"),
                text_color=ROAST,
                anchor="w"
            ).pack(anchor="w")

            ctk.CTkLabel(
                info,
                text=f"↳ File xuất sẽ tạo: {out_preview}",
                font=ctk.CTkFont(family=FONT_SANS, size=10),
                text_color=BROWN,
                anchor="w"
            ).pack(anchor="w")

            # Pill Badge
            badge = ctk.CTkLabel(
                row,
                text=f"  {tag_text}  ",
                font=ctk.CTkFont(family=FONT_SANS, size=10),
                fg_color=tag_color,
                text_color=tag_text_color,
                corner_radius=10
            )
            badge.pack(side="left", padx=8, pady=6)

            # Delete button
            del_btn = ctk.CTkButton(
                row,
                text="✕",
                font=ctk.CTkFont(size=11, weight="bold"),
                fg_color="transparent",
                text_color=DANGER,
                hover_color=SAND,
                border_width=0,
                corner_radius=12,
                width=24,
                height=24,
                command=lambda p=fpath: self.remove_claim_file(p)
            )
            del_btn.pack(side="right", padx=(4, 10), pady=6)

    def remove_claim_file(self, fpath):
        if fpath in self.claim_files:
            self.claim_files.remove(fpath)
            self.save_state()
            self.render_file_list()

    def clear_all_claims(self):
        if not self.claim_files:
            return
        if messagebox.askyesno("Xác nhận", "Bạn có chắc muốn xóa tất cả file khỏi danh sách không?"):
            self.claim_files.clear()
            self.save_state()
            self.render_file_list()

    # ----------------------------------------------------
    # Drag & Drop Handlers
    # ----------------------------------------------------
    def setup_drag_and_drop(self):
        targets = [self.card_ov, self.ov_display_container, self.card_claims, self.file_list_frame, self]
        for t in targets:
            try:
                t.drop_target_register(tkdnd.DND_FILES)
                t.dnd_bind('<<Drop>>', self.on_drop_generic)
            except Exception:
                pass

        try:
            self.card_ov.dnd_bind('<<Drop>>', self.on_drop_overview)
            self.ov_display_container.dnd_bind('<<Drop>>', self.on_drop_overview)

            def on_ov_enter(e):
                self.ov_display_container.configure(border_color=MOSS, fg_color=MOSS_SOFT)
            def on_ov_leave(e):
                self.ov_display_container.configure(border_color=FAINT, fg_color=SAND)

            self.card_ov.dnd_bind('<<DropEnter>>', on_ov_enter)
            self.ov_display_container.dnd_bind('<<DropEnter>>', on_ov_enter)
            self.card_ov.dnd_bind('<<DropLeave>>', on_ov_leave)
            self.ov_display_container.dnd_bind('<<DropLeave>>', on_ov_leave)
        except Exception:
            pass

        try:
            self.card_claims.dnd_bind('<<Drop>>', self.on_drop_claims)
            self.file_list_frame.dnd_bind('<<Drop>>', self.on_drop_claims)

            def on_cl_enter(e):
                self.file_list_frame.configure(border_color=MOSS, fg_color=MOSS_SOFT)
            def on_cl_leave(e):
                self.file_list_frame.configure(border_color=FAINT, fg_color=SAND)

            self.card_claims.dnd_bind('<<DropEnter>>', on_cl_enter)
            self.file_list_frame.dnd_bind('<<DropEnter>>', on_cl_enter)
            self.card_claims.dnd_bind('<<DropLeave>>', on_cl_leave)
            self.file_list_frame.dnd_bind('<<DropLeave>>', on_cl_leave)
        except Exception:
            pass

    def on_drop_overview(self, event):
        self.ov_display_container.configure(border_color=FAINT, fg_color=SAND)
        paths = parse_drop_paths(event.data)
        for p in paths:
            if os.path.isfile(p) and p.endswith(".xlsx"):
                self.overview_path = p
                self.save_state()
                self.render_overview_display()
                return

    def on_drop_claims(self, event):
        self.file_list_frame.configure(border_color=FAINT, fg_color=SAND)
        paths = parse_drop_paths(event.data)
        self.add_claim_paths(paths)

    def on_drop_generic(self, event):
        self.ov_display_container.configure(border_color=FAINT, fg_color=SAND)
        self.file_list_frame.configure(border_color=FAINT, fg_color=SAND)
        paths = parse_drop_paths(event.data)
        if not paths:
            return

        claims_to_add = []
        for p in paths:
            if os.path.isfile(p) and p.endswith(".xlsx"):
                fname = os.path.basename(p).lower()
                if "overview" in fname:
                    self.overview_path = p
                    self.save_state()
                    self.render_overview_display()
                else:
                    claims_to_add.append(p)
            elif os.path.isdir(p):
                claims_to_add.append(p)

        if claims_to_add:
            self.add_claim_paths(claims_to_add)

    def add_claim_paths(self, paths):
        added = 0
        for p in paths:
            if os.path.isfile(p) and p.endswith(".xlsx") and not p.endswith("_backup.xlsx") and not p.startswith("~$"):
                if p not in self.claim_files:
                    self.claim_files.append(p)
                    added += 1
            elif os.path.isdir(p):
                for f in os.listdir(p):
                    if "versus" in f.lower() and f.endswith(".xlsx") and not f.endswith("_backup.xlsx") and not f.endswith("_filled.xlsx") and not f.startswith("~$"):
                        fp = os.path.join(p, f)
                        if fp not in self.claim_files:
                            self.claim_files.append(fp)
                            added += 1
        if added > 0:
            self.save_state()
            self.render_file_list()

    # ----------------------------------------------------
    # File Dialogs
    # ----------------------------------------------------
    def browse_overview(self):
        f = filedialog.askopenfilename(
            title="Chọn file Master Overview",
            filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")]
        )
        if f:
            self.overview_path = os.path.abspath(f)
            self.save_state()
            self.render_overview_display()

    def browse_claim_files(self):
        files = filedialog.askopenfilenames(
            title="Chọn một hoặc nhiều file Claim Versus",
            filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")]
        )
        if files:
            self.add_claim_paths([os.path.abspath(f) for f in files])

    def browse_claim_folder(self):
        d = filedialog.askdirectory(title="Chọn thư mục chứa các file Claim Versus")
        if d:
            self.add_claim_paths([d])

    def force_reload_master(self):
        if not self.overview_path or not os.path.exists(self.overview_path):
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn file Master Overview trước!")
            return
        
        c_path = get_cache_path()
        if os.path.exists(c_path):
            try:
                os.remove(c_path)
            except Exception:
                pass
        self.render_overview_display()
        messagebox.showinfo("Thông báo", "Đã xóa bộ nhớ ghi nhớ! Lần chạy tới sẽ đọc lại file Master từ đầu.")

    # ----------------------------------------------------
    # Thread-Safe Results Feed Updates
    # ----------------------------------------------------
    def append_feed_item(self, filename, matched, total, rate_str, full_fpath=None):
        row = ctk.CTkFrame(self.results_feed, fg_color=CARD, corner_radius=10, border_width=1, border_color=LINE)
        row.pack(fill="x", padx=10, pady=3)

        ctk.CTkLabel(
            row,
            text="✓",
            font=ctk.CTkFont(family=FONT_SANS, size=12, weight="bold"),
            text_color=MOSS
        ).pack(side="left", padx=(12, 8), pady=6)

        ctk.CTkLabel(
            row,
            text=filename,
            font=ctk.CTkFont(family=FONT_SERIF, size=12, weight="bold"),
            text_color=ROAST
        ).pack(side="left", pady=6)

        ctk.CTkLabel(
            row,
            text=f"•  Khớp {matched}/{total} dòng ({rate_str})",
            font=ctk.CTkFont(family=FONT_SANS, size=11),
            text_color=BROWN
        ).pack(side="left", padx=8, pady=6)

        if full_fpath and os.path.exists(full_fpath):
            ctk.CTkButton(
                row,
                text="📂 Xem",
                font=ctk.CTkFont(family=FONT_SANS, size=10),
                fg_color="transparent",
                hover_color=SAND,
                border_width=1,
                border_color=LINE,
                text_color=ROAST,
                corner_radius=12,
                width=52,
                height=24,
                command=lambda p=full_fpath: reveal_in_explorer(p)
            ).pack(side="right", padx=(4, 10), pady=6)

    def append_feed_error_item(self, filename, error_msg):
        row = ctk.CTkFrame(self.results_feed, fg_color=CARD, corner_radius=10, border_width=1, border_color=DANGER)
        row.pack(fill="x", padx=10, pady=3)

        ctk.CTkLabel(
            row,
            text="✕",
            font=ctk.CTkFont(family=FONT_SANS, size=12, weight="bold"),
            text_color=DANGER
        ).pack(side="left", padx=(12, 8), pady=6)

        ctk.CTkLabel(
            row,
            text=filename,
            font=ctk.CTkFont(family=FONT_SERIF, size=12, weight="bold"),
            text_color=ROAST
        ).pack(side="left", pady=6)

        ctk.CTkLabel(
            row,
            text=f"•  {error_msg}",
            font=ctk.CTkFont(family=FONT_SANS, size=11),
            text_color=DANGER,
            wraplength=520,
            justify="left"
        ).pack(side="left", padx=8, pady=6)

    def set_status_text(self, text, color=BROWN, bg_color=None):
        if bg_color:
            self.status_badge.configure(text=f"  {text}  ", text_color=color, fg_color=bg_color, corner_radius=10)
        else:
            self.status_badge.configure(text=text, text_color=color, fg_color="transparent")

    # ----------------------------------------------------
    # Core Engine Processing (Thread-Safe & Lag-Free)
    # ----------------------------------------------------
    def start_processing(self):
        if not self.overview_path or not os.path.isfile(self.overview_path):
            messagebox.showerror("Thiếu thông tin", "Vui lòng chọn hoặc kéo thả file Master Overview (.xlsx)!")
            return
        if not self.claim_files:
            messagebox.showerror("Thiếu thông tin", "Danh sách file Claim đang trống. Vui lòng chọn hoặc kéo thả ít nhất 1 file Claim Versus!")
            return

        self.btn_run.configure(state="disabled", fg_color=FAINT, text="Đang xử lý dữ liệu...")
        self.prog_bar.set(0.05)
        self.set_status_text("● Đang xử lý...", MOSS)
        self.last_created_files = []

        # Clear feed
        for child in self.results_feed.winfo_children():
            child.destroy()

        # Run background thread
        thread = threading.Thread(target=self.run_engine, daemon=True)
        thread.start()

    def run_engine(self):
        try:
            # 0. Validate Master Overview
            try:
                validate_excel_file(self.overview_path)
            except Exception as ve:
                self.after(0, self.finish_processing_apple, 0, [], 0, f"Lỗi file Master Overview: {ve}")
                return

            # 1. Load Master Overview with Cache
            c_path = get_cache_path()
            ov_abs = os.path.abspath(self.overview_path)
            cur_size = os.path.getsize(ov_abs)
            cur_mtime = os.path.getmtime(ov_abs)

            loaded_from_cache = False
            cs_records = []
            nh_records = []

            if os.path.exists(c_path):
                try:
                    with open(c_path, 'rb') as f:
                        cache_obj = pickle.load(f)
                    if (cache_obj.get('file_path') == ov_abs and
                        cache_obj.get('file_size') == cur_size and
                        abs(cache_obj.get('file_mtime', 0) - cur_mtime) < 2):
                        cs_records = cache_obj['cs_records']
                        nh_records = cache_obj['nh_records']
                        loaded_from_cache = True
                except Exception:
                    pass

            if not loaded_from_cache:
                self.after(0, self.set_status_text, "● Đang đọc file Master...", MOSS)
                try:
                    wb_o = openpyxl.load_workbook(self.overview_path, read_only=True, data_only=True)
                except PermissionError:
                    self.after(0, self.finish_processing_apple, 0, [], 0, "File Master Overview đang mở trong Excel hoặc ứng dụng khác. Vui lòng đóng lại trước khi chạy.")
                    return
                except (zipfile.BadZipFile, Exception) as oe:
                    self.after(0, self.finish_processing_apple, 0, [], 0, f"Không thể đọc file Master Overview ({oe}).")
                    return

                s_o = wb_o['Overview'] if 'Overview' in wb_o.sheetnames else wb_o.active
                headers_o = next(s_o.iter_rows(max_row=1, values_only=True))
                col_idx = {h: i for i, h in enumerate(headers_o)}

                for r in s_o.iter_rows(min_row=2, values_only=True):
                    man = str(r[col_idx.get('Manufacturer', 0)] or '').strip().upper()
                    sup = str(r[col_idx.get('Supplier', 0)] or '').strip().upper()
                    cost = r[col_idx.get('Purchase cost', 0)]
                    note_p = str(r[col_idx.get('Item PO note', 0)] or '').lower()
                    note_s = str(r[col_idx.get('Note (sales)', 0)] or '').lower()

                    if 'box 1 of 2' in note_p or 'box 2 of 2' in note_p or 'box 1 of' in note_s:
                        continue
                    if cost is None or float(cost) <= 0:
                        continue

                    rec = {
                        'art': str(r[col_idx.get('Customer art No', 0)] or '').strip(),
                        'item_name': str(r[col_idx.get('Customer item name', 0)] or '').strip(),
                        'fabric': str(r[col_idx.get('Fabric type', 0)] or '').strip(),
                        'cost': float(cost),
                        'po_date': r[col_idx.get('P.O issued date', 0)],
                        'model': str(r[col_idx.get('Model', 0)] or '').strip()
                    }

                    if man in ['CS', 'WORKSHOP'] or 'CASA' in man or 'CASA' in sup:
                        cs_records.append(rec)
                    elif man in ['NH', 'NHF', 'NH FSC', 'NHAN HOANG'] or 'NHAN HOANG' in sup or 'NH ' in sup:
                        nh_records.append(rec)

                wb_o.close()

                # Save cache
                try:
                    cache_obj = {
                        'file_path': ov_abs,
                        'file_size': cur_size,
                        'file_mtime': cur_mtime,
                        'cached_at': datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
                        'cs_records': cs_records,
                        'nh_records': nh_records
                    }
                    with open(c_path, 'wb') as f:
                        pickle.dump(cache_obj, f, protocol=pickle.HIGHEST_PROTOCOL)
                except Exception:
                    pass

            self.after(0, self.render_overview_display)

            # 2. Process Files
            total_files = len(self.claim_files)
            created_files = []
            failed_count = 0

            for i, fpath in enumerate(self.claim_files, start=1):
                fname = os.path.basename(fpath)
                try:
                    validate_excel_file(fpath)

                    is_casa = "casa" in fname.lower()
                    ref_db = cs_records if is_casa else nh_records

                    wb = openpyxl.load_workbook(fpath)
                    s = wb['Export'] if 'Export' in wb.sheetnames else wb.active
                    headers = [cell.value for cell in s[1]]

                    # Add Purchase Cost (USD)
                    if 'Purchase Cost (USD)' in headers:
                        col_cost_idx = headers.index('Purchase Cost (USD)') + 1
                    else:
                        col_cost_idx = len(headers) + 1
                        headers.append('Purchase Cost (USD)')
                        c_h1 = s.cell(row=1, column=col_cost_idx, value='Purchase Cost (USD)')
                        sample_h = s.cell(row=1, column=headers.index('Unit Cost x QTY')+1 if 'Unit Cost x QTY' in headers else 1)
                        if sample_h.has_style:
                            c_h1.font, c_h1.border, c_h1.fill, c_h1.alignment = copy(sample_h.font), copy(sample_h.border), copy(sample_h.fill), copy(sample_h.alignment)

                    # Add Purchase Cost Note
                    if 'Purchase Cost Note' in headers:
                        col_note_idx = headers.index('Purchase Cost Note') + 1
                    else:
                        col_note_idx = len(headers) + 1
                        headers.append('Purchase Cost Note')
                        c_h2 = s.cell(row=1, column=col_note_idx, value='Purchase Cost Note')
                        sample_h = s.cell(row=1, column=headers.index('Unit Cost x QTY')+1 if 'Unit Cost x QTY' in headers else 1)
                        if sample_h.has_style:
                            c_h2.font, c_h2.border, c_h2.fill, c_h2.alignment = copy(sample_h.font), copy(sample_h.border), copy(sample_h.fill), copy(sample_h.alignment)

                    cost_letter = get_column_letter(col_cost_idx)
                    note_letter = get_column_letter(col_note_idx)
                    s.column_dimensions[cost_letter].width = 22
                    s.column_dimensions[note_letter].width = 85

                    max_row = s.max_row
                    art_col_idx = headers.index('No') + 1 if 'No' in headers else 7
                    unit_cost_col_idx = headers.index('Unit Cost x QTY') + 1 if 'Unit Cost x QTY' in headers else None

                    last_art = s.cell(row=max_row, column=art_col_idx).value
                    last_is_total = (last_art is None or str(last_art).strip() == '')
                    data_end = max_row - 1 if last_is_total else max_row
                    data_count = data_end - 1

                    rate_samples = []
                    if unit_cost_col_idx:
                        for r_idx in range(2, data_end + 1):
                            art = str(s.cell(row=r_idx, column=art_col_idx).value or '').strip()
                            dkk = s.cell(row=r_idx, column=unit_cost_col_idx).value
                            exact = [rec for rec in ref_db if rec['art'] == art]
                            if exact and isinstance(dkk, (int, float)) and dkk > 0:
                                avg_usd = sum(e['cost'] for e in exact) / len(exact)
                                rate_samples.append(dkk / avg_usd)

                    median_rate = sorted(rate_samples)[len(rate_samples)//2] if rate_samples else (8.0087 if is_casa else 7.8210)

                    file_matched = 0
                    for r_idx in range(2, data_end + 1):
                        art = str(s.cell(row=r_idx, column=art_col_idx).value or '').strip()
                        posting_date = s.cell(row=r_idx, column=headers.index('PostingDate')+1).value if 'PostingDate' in headers else None
                        days = s.cell(row=r_idx, column=headers.index('Days Between Order And Complaint')+1).value if 'Days Between Order And Complaint' in headers else 0
                        days = days if isinstance(days, (int, float)) else 0
                        unit_cost_dkk = s.cell(row=r_idx, column=unit_cost_col_idx).value if unit_cost_col_idx else 0
                        unit_cost_dkk = unit_cost_dkk if isinstance(unit_cost_dkk, (int, float)) else 0

                        order_date = None
                        if isinstance(posting_date, datetime.datetime):
                            order_date = posting_date - datetime.timedelta(days=int(days))
                        elif isinstance(posting_date, datetime.date):
                            order_date = posting_date - datetime.timedelta(days=int(days))

                        exact_matches = [rec for rec in ref_db if rec['art'] == art]
                        selected_cost = None
                        note = ""

                        if exact_matches:
                            file_matched += 1
                            if unit_cost_dkk > 0:
                                target_usd = unit_cost_dkk / median_rate
                                plausible = [e for e in exact_matches if 0.7 * target_usd <= e['cost'] <= 1.3 * target_usd]
                                if not plausible: plausible = exact_matches
                            else:
                                plausible = exact_matches

                            valid_dates = [e for e in plausible if isinstance(e['po_date'], (datetime.datetime, datetime.date))]
                            if valid_dates and order_date:
                                o_d = order_date.date() if isinstance(order_date, datetime.datetime) else order_date
                                best_match = min(valid_dates, key=lambda e: abs((e['po_date'].date() if isinstance(e['po_date'], datetime.datetime) else e['po_date']) - o_d))
                                selected_cost = best_match['cost']
                                po_str = best_match['po_date'].strftime('%d/%m/%Y')
                                od_str = order_date.strftime('%d/%m/%Y')
                                note = f"Khớp 100% Art No ({art}). Chọn {selected_cost}$ theo đợt PO ({po_str}) gần nhất ngày đặt hàng ({od_str})."
                            else:
                                best_match = min(plausible, key=lambda e: abs(e['cost'] - (unit_cost_dkk/median_rate if unit_cost_dkk > 0 else plausible[0]['cost'])))
                                selected_cost = best_match['cost']
                                note = f"Khớp 100% Art No ({art}). Chọn {selected_cost}$ từ danh mục trong Overview."
                        else:
                            candidates = [rec for rec in ref_db if rec['art'].startswith(art) or art in rec['art']]
                            if not candidates and len(art) >= 6:
                                candidates = [rec for rec in ref_db if rec['art'].startswith(art[:6])]

                            if candidates and unit_cost_dkk > 0:
                                file_matched += 1
                                est_usd = unit_cost_dkk / median_rate
                                best_cand = min(candidates, key=lambda c: abs(c['cost'] - est_usd))
                                selected_cost = best_cand['cost']
                                note = f"Mã rút gọn MM ({art}). Dựa vào Unit Cost {unit_cost_dkk:.2f} DKK (~{est_usd:.2f}$), khớp phân khúc vải '{best_cand['fabric']}' (giá {selected_cost}$)."
                            elif candidates:
                                file_matched += 1
                                selected_cost = round(sum(c['cost'] for c in candidates) / len(candidates), 2)
                                note = f"Mã rút gọn MM ({art}). Lấy giá trung bình phân khúc {selected_cost}$."
                            elif unit_cost_dkk > 0:
                                file_matched += 1
                                selected_cost = round(unit_cost_dkk / median_rate, 2)
                                note = f"Mã {art} không còn trong Overview. Tính theo tỷ giá quy đổi Unit Cost: {unit_cost_dkk:.2f} DKK / {median_rate:.4f} = {selected_cost}$."
                            else:
                                selected_cost = 0.0
                                note = f"Không tìm thấy mã {art} trong Overview."

                        c_cost = s.cell(row=r_idx, column=col_cost_idx, value=selected_cost)
                        c_cost.number_format = '#,##0.00'
                        sample_c = s.cell(row=r_idx, column=unit_cost_col_idx if unit_cost_col_idx else 1)
                        if sample_c.has_style:
                            c_cost.font, c_cost.border = copy(sample_c.font), copy(sample_c.border)

                        c_note = s.cell(row=r_idx, column=col_note_idx, value=note)
                        if sample_c.has_style:
                            c_note.font, c_note.border = copy(sample_c.font), copy(sample_c.border)

                    # Total Row
                    if last_is_total:
                        c_total = s.cell(row=max_row, column=col_cost_idx, value=f"=SUM({cost_letter}2:{cost_letter}{data_end})")
                        c_total.number_format = '#,##0.00'
                        sample_t = s.cell(row=max_row, column=unit_cost_col_idx if unit_cost_col_idx else 1)
                        if sample_t.has_style:
                            c_total.font, c_total.border = copy(sample_t.font), copy(sample_t.border)

                        c_t_note = s.cell(row=max_row, column=col_note_idx, value=f"Tổng Purchase Cost (USD) của {data_count} vụ khiếu nại")
                        if sample_t.has_style:
                            c_t_note.font, c_t_note.border = copy(sample_t.font), copy(sample_t.border)

                    # Save as adjacent file with '_filled' suffix
                    dir_name = os.path.dirname(fpath)
                    base_name, ext = os.path.splitext(fname)
                    if base_name.endswith("_filled"):
                        out_fpath = fpath
                    else:
                        out_fpath = os.path.join(dir_name, f"{base_name}_filled{ext}")

                    wb.save(out_fpath)
                    wb.close()

                    out_abs = os.path.abspath(out_fpath)
                    created_files.append(out_abs)

                    rate_str = f"{(file_matched / data_count * 100):.1f}%" if data_count > 0 else "100%"
                    out_name = os.path.basename(out_fpath)
                    self.after(0, self.append_feed_item, out_name, file_matched, data_count, rate_str, out_abs)

                except PermissionError:
                    failed_count += 1
                    err_msg = "File đang mở trong Excel hoặc ứng dụng khác. Vui lòng đóng file rồi thử lại."
                    self.after(0, self.append_feed_error_item, fname, err_msg)
                except (zipfile.BadZipFile, ValueError) as e:
                    failed_count += 1
                    err_msg = str(e) if isinstance(e, ValueError) else "File bị hỏng cấu trúc nén hoặc không phải định dạng .xlsx hợp lệ."
                    self.after(0, self.append_feed_error_item, fname, err_msg)
                except Exception as e:
                    failed_count += 1
                    err_msg = f"Lỗi: {e}"
                    self.after(0, self.append_feed_error_item, fname, err_msg)
                finally:
                    progress_val = 0.1 + 0.9 * (i / total_files)
                    self.after(0, self.prog_bar.set, progress_val)

            # Completion
            self.last_created_files = created_files
            self.after(0, self.finish_processing_apple, total_files, created_files, failed_count)

        except Exception as e:
            self.after(0, self.finish_processing_apple, 0, [], 0, str(e))

    def finish_processing_apple(self, total_files, created_files, failed_count=0, err_msg=""):
        self.prog_bar.set(1.0 if (len(created_files) > 0 and not err_msg) else 0)
        self.btn_run.configure(state="normal", fg_color=ROAST, text="Bắt Đầu Đối Chiếu & Điền Giá (Tạo File _filled)")

        if err_msg:
            self.set_status_text("❌ Lỗi", DANGER)
            err_box = ctk.CTkFrame(self.results_feed, fg_color=CARD, corner_radius=10, border_width=1, border_color=DANGER)
            err_box.pack(fill="x", padx=10, pady=6)
            ctk.CTkLabel(
                err_box,
                text=f"Đã xảy ra lỗi: {err_msg}",
                font=ctk.CTkFont(family=FONT_SANS, size=11),
                text_color=DANGER,
                wraplength=520,
                justify="left"
            ).pack(padx=12, pady=10)
            return

        if failed_count == 0 and len(created_files) > 0:
            self.set_status_text("✓ Hoàn tất", MOSS, MOSS_SOFT)
            title_text = f"Đã hoàn tất và điền giá {total_files} file."
            banner_bg = MOSS_SOFT
            banner_border = MOSS
            banner_text_color = ROAST
        elif len(created_files) > 0:
            self.set_status_text("Đã hoàn tất một phần", BROWN, SAND)
            title_text = f"Đã hoàn tất và điền giá {len(created_files)}/{total_files} file ({failed_count} file cần kiểm tra lại)."
            banner_bg = SAND
            banner_border = LINE
            banner_text_color = ROAST
        else:
            self.set_status_text("✕ Chưa thể xử lý", DANGER)
            title_text = "Chưa thể hoàn tất các file đã chọn. Vui lòng xem chi tiết ở danh sách trên."
            banner_bg = CARD
            banner_border = DANGER
            banner_text_color = DANGER

        # Determine subtitle message with paths
        if len(created_files) == 1:
            sub_text = f"Đã tạo file: {created_files[0]}"
            target_file_to_reveal = created_files[0]
        elif len(created_files) > 1:
            lines = [f"• {p}" for p in created_files]
            sub_text = f"Đã tạo {len(created_files)} file:\n" + "\n".join(lines)
            target_file_to_reveal = created_files[0]
        else:
            sub_text = ""
            target_file_to_reveal = ""

        # Render completion / status banner
        banner = ctk.CTkFrame(self.results_feed, fg_color=banner_bg, corner_radius=12, border_width=1, border_color=banner_border)
        banner.pack(fill="x", padx=10, pady=(8, 4))

        b_left = ctk.CTkFrame(banner, fg_color="transparent")
        b_left.pack(side="left", fill="x", expand=True, padx=16, pady=12)

        ctk.CTkLabel(
            b_left,
            text=title_text,
            font=ctk.CTkFont(family=FONT_SERIF, size=13, weight="bold"),
            text_color=banner_text_color
        ).pack(anchor="w")

        if sub_text:
            ctk.CTkLabel(
                b_left,
                text=sub_text,
                font=ctk.CTkFont(family=FONT_SANS, size=11),
                text_color=BROWN,
                wraplength=520,
                justify="left"
            ).pack(anchor="w", pady=(3, 0))

        if target_file_to_reveal:
            ctk.CTkButton(
                banner,
                text="📂 Xem File",
                font=ctk.CTkFont(family=FONT_SANS, size=11, weight="bold"),
                fg_color=ROAST,
                text_color=IVORY,
                hover_color=MOSS,
                corner_radius=16,
                height=32,
                command=lambda p=target_file_to_reveal: reveal_in_explorer(p)
            ).pack(side="right", padx=16, pady=12)

if __name__ == "__main__":
    app = BoringTaskApp()
    app.mainloop()
