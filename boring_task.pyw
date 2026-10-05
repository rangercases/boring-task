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
import shutil
import hashlib
import urllib.parse
from concurrent.futures import ThreadPoolExecutor, as_completed

import customtkinter as ctk
from tkinter import filedialog, messagebox
import tkinterdnd2 as tkdnd
import openpyxl
from openpyxl.utils import get_column_letter
from PIL import Image

from fabric_checker import FabricRuleManager, check_overview_file
import order_auditor

# Set Windows App User Model ID so Taskbar groups and shows custom icon
try:
    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("rangercases.boringtask.v1")
except Exception:
    pass

# Set Light Mode
ctk.set_appearance_mode("Light")

APP_VERSION = "v1.1"
GITHUB_REPO = "rangercases/claim-helper"
CACHE_FILE_NAME = ".overview_cache.pkl"
STATE_FILE_NAME = ".app_state.json"

# Per-machine module visibility (controlled centrally by code & local config.json)
CONFIG_FILE_NAME = "config.json"
APP_MODES = ("all", "cost", "images", "fabric", "auditor")

# BẢNG PHÂN QUYỀN TẬP TRUNG (Sửa tại đây để phân quyền từ xa qua Git update)
USER_PERMISSIONS = {
    "nhung": ["cost"],                                   # Ms Nhung: chỉ xem Purchase Cost
    "thuy":  ["fabric", "auditor"],                      # Ms Thuy: Fabric Checker & Order Auditor
    "admin": ["cost", "fabric", "images", "auditor"],   # Admin: toàn quyền xem tất cả các module
}

# Image Inserter module
IMAGE_CACHE_DIR = os.path.join(os.environ.get("LOCALAPPDATA") or tempfile.gettempdir(), "BoringTask", "ImageCache")
IMG_LINK_COLUMNS = ("Comp. Pic. 1", "Pic 2", "Pic 3", "Pic 4", "Pic 5")
IMG_EXTENSIONS = (".xlsx", ".xlsm", ".xls")
IMG_BTN_TEXT = "Bắt Đầu Chèn Ảnh (Place in Cell)"

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

def get_config_path():
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), CONFIG_FILE_NAME)

def load_app_mode():
    """Reads which modules this machine may see.
    Supports both user mapping (USER_PERMISSIONS) and legacy mode strings.
    Lightning-fast local read (<1ms).
    """
    try:
        with open(get_config_path(), "r", encoding="utf-8") as f:
            cfg = json.load(f)
        
        # Check user key first (Ms Nhung / Ms Thuy / Admin)
        user = str(cfg.get("user", "")).strip().lower()
        if user in USER_PERMISSIONS:
            return USER_PERMISSIONS[user]
        
        # Fallback to direct mode string if present
        mode = str(cfg.get("mode", "all")).strip().lower()
        if mode == "all":
            return ["cost", "fabric", "images", "auditor"]
        elif mode in ("cost", "fabric", "images", "auditor"):
            return [mode]
        return ["cost", "fabric", "images", "auditor"]
    except Exception:
        return ["cost", "fabric", "images", "auditor"]

def img_url_to_unc(url):
    """Converts a file:// link stored in Excel into a Windows UNC path."""
    if not isinstance(url, str) or not url.strip():
        return None
    url = url.strip()
    if url.startswith('file://'):
        url = url[7:]
    elif url.startswith('file:'):
        url = url[5:]
    url = urllib.parse.unquote(url)
    return '\\\\' + url.replace('/', '\\').lstrip('\\')

def img_thumb_path(unc_path):
    h = hashlib.md5(unc_path.lower().encode('utf-8')).hexdigest()
    return os.path.join(IMAGE_CACHE_DIR, f"{h}.jpg")

def img_prepare_thumb(unc_path):
    """Downloads + compresses one image to the local cache. Returns True if ready."""
    target = img_thumb_path(unc_path)
    if os.path.exists(target) and os.path.getsize(target) > 0:
        return True
    try:
        if not os.path.exists(unc_path):
            return False
        with Image.open(unc_path) as img:
            img.thumbnail((800, 800), Image.Resampling.LANCZOS)
            if img.mode in ('RGBA', 'P', 'LA'):
                img = img.convert('RGB')
            img.save(target, 'JPEG', quality=88)
        return True
    except Exception:
        return False

def com_rows(value):
    """Normalizes Excel Range.Value (scalar or tuple-of-tuples) to tuple-of-tuples."""
    if isinstance(value, tuple):
        return value
    return ((value,),)

class BoringTaskApp(ctk.CTk, tkdnd.TkinterDnD.DnDWrapper):
    def __init__(self):
        super().__init__()
        self.has_dnd = False
        try:
            self.TkdndVersion = tkdnd.TkinterDnD._require(self)
            self.has_dnd = True
        except Exception:
            pass

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
        self.app_mode = load_app_mode()
        self.current_view = None
        self.overview_path = ""
        self.claim_files = [] # list of absolute paths
        self.last_created_files = [] # list of created output files

        # Image Inserter state
        self.img_file = ""
        self.img_running = False
        self.img_active_step = -1

        # Fabric Checker state
        curr_dir = os.path.abspath(os.path.dirname(__file__))
        self.fabric_rule_mgr = FabricRuleManager(curr_dir)
        self.fabric_overview_file = ""
        self.fabric_running = False

        # Order Auditor state
        self.auditor_engine = order_auditor.OrderAuditEngine()
        self.auditor_po_file = ""
        self.auditor_cust_file = ""
        self.auditor_running = False
        self.auditor_last_report = ""
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
                                    app_dir = os.path.dirname(curr_file)
                                    tmp_file = curr_file + ".new"
                                    with open(tmp_file, "wb") as f:
                                        f.write(new_code)
                                    os.replace(tmp_file, curr_file)

                                    # Also ensure fabric_checker.py is kept up to date
                                    try:
                                        fc_url = f"https://raw.githubusercontent.com/{GITHUB_REPO}/main/fabric_checker.py"
                                        fc_req = urllib.request.Request(fc_url, headers={"User-Agent": "BoringTask-App"})
                                        with urllib.request.urlopen(fc_req, timeout=5) as fc_resp:
                                            if fc_resp.status == 200:
                                                fc_code = fc_resp.read()
                                                if len(fc_code) > 500:
                                                    fc_path = os.path.join(app_dir, "fabric_checker.py")
                                                    fc_tmp = fc_path + ".new"
                                                    with open(fc_tmp, "wb") as f_fc:
                                                        f_fc.write(fc_code)
                                                    os.replace(fc_tmp, fc_path)
                                    except Exception:
                                        pass

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
        self.image_view = ctk.CTkFrame(self.view_container, fg_color="transparent")
        self.fabric_view = ctk.CTkFrame(self.view_container, fg_color="transparent")
        self.auditor_view = ctk.CTkFrame(self.view_container, fg_color="transparent")

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

        if "images" in self.app_mode:
            self.build_image_inserter()
        if "fabric" in self.app_mode:
            self.build_fabric_checker()
        if "auditor" in self.app_mode:
            self.build_order_auditor()

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

        # To add a new feature: append one entry here (gated by config.json mode).
        features = []
        if "cost" in self.app_mode:
            features.append({
                "icon": "📋",
                "title": "Purchase Cost",
                "subtitle": "Auto-Filled",
                "command": lambda: self.show_feature(self.claim_view, "Purchase Cost Auto-Filled"),
            })
        if "images" in self.app_mode:
            features.append({
                "icon": "🖼️",
                "title": "Image Inserter",
                "subtitle": "Place in Cell",
                "command": lambda: self.show_feature(self.image_view, "Image Inserter"),
            })
        if "fabric" in self.app_mode:
            features.append({
                "icon": "🧵",
                "title": "Fabric Checker",
                "subtitle": "Mapping & Validate",
                "command": lambda: self.show_feature(self.fabric_view, "Fabric Checker"),
            })
        if "auditor" in self.app_mode:
            features.append({
                "icon": "⚖️",
                "title": "Order Auditor",
                "subtitle": "Multi-Layer PO Check",
                "command": lambda: self.show_feature(self.auditor_view, "Order Auditor"),
            })

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
        self.image_view.pack_forget()
        self.fabric_view.pack_forget()
        self.auditor_view.pack_forget()
        self.home_view.pack(fill="both", expand=True)
        self.current_view = self.home_view
        self.title("Boring Task")

    def show_feature(self, view, title):
        self.home_view.pack_forget()
        view.pack(fill="both", expand=True)
        self.current_view = view
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
        if not getattr(self, "has_dnd", False):
            return
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

        # Image Inserter drop zone
        if hasattr(self, "img_drop_zone"):
            def on_img_enter(e):
                self.img_drop_zone.configure(border_color=MOSS, fg_color=MOSS_SOFT)
            def on_img_leave(e):
                self.img_drop_zone.configure(border_color=FAINT, fg_color=SAND)

            for t in (self.img_card, self.img_drop_zone):
                try:
                    t.drop_target_register(tkdnd.DND_FILES)
                    t.dnd_bind('<<Drop>>', self.on_drop_image)
                    t.dnd_bind('<<DropEnter>>', on_img_enter)
                    t.dnd_bind('<<DropLeave>>', on_img_leave)
                except Exception:
                    pass

        # Fabric Checker drop zones
        if hasattr(self, "fabric_rule_status_box"):
            def on_fabric_rule_enter(e):
                self.fabric_rule_status_box.configure(border_color=MOSS, fg_color=MOSS_SOFT)
            def on_fabric_rule_leave(e):
                self.fabric_rule_status_box.configure(border_color=FAINT, fg_color=SAND)

            for t in (self.card_fabric_rule, self.fabric_rule_status_box):
                try:
                    t.drop_target_register(tkdnd.DND_FILES)
                    t.dnd_bind('<<Drop>>', self.on_drop_fabric_rule)
                    t.dnd_bind('<<DropEnter>>', on_fabric_rule_enter)
                    t.dnd_bind('<<DropLeave>>', on_fabric_rule_leave)
                except Exception:
                    pass

        if hasattr(self, "fabric_ov_drop_zone"):
            def on_fabric_ov_enter(e):
                self.fabric_ov_drop_zone.configure(border_color=MOSS, fg_color=MOSS_SOFT)
            def on_fabric_ov_leave(e):
                self.fabric_ov_drop_zone.configure(border_color=FAINT, fg_color=SAND)

            for t in (self.card_fabric_ov, self.fabric_ov_drop_zone):
                try:
                    t.drop_target_register(tkdnd.DND_FILES)
                    t.dnd_bind('<<Drop>>', self.on_drop_fabric_overview)
                    t.dnd_bind('<<DropEnter>>', on_fabric_ov_enter)
                    t.dnd_bind('<<DropLeave>>', on_fabric_ov_leave)
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
        # Route window-level drops to whichever module is on screen
        if self.current_view is self.image_view:
            return self.on_drop_image(event)
        if self.current_view is not self.claim_view:
            return
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

    # ====================================================
    # MODULE: Image Inserter / Place in Cell
    # ====================================================
    def build_image_inserter(self):
        self.build_feature_nav(self.image_view, "Image Inserter")

        scroll = ctk.CTkScrollableFrame(self.image_view, fg_color="transparent")
        scroll.pack(fill="both", expand=True, padx=36, pady=(0, 16))

        # ---------- CARD 1: Excel file ----------
        self.img_card = ctk.CTkFrame(scroll, fg_color=CARD, corner_radius=16, border_width=1, border_color=LINE)
        self.img_card.pack(fill="x", pady=(0, 20))

        top = ctk.CTkFrame(self.img_card, fg_color="transparent")
        top.pack(fill="x", padx=24, pady=(18, 10))

        ctk.CTkLabel(
            top,
            text="FILE EXCEL CẦN CHÈN ẢNH",
            font=ctk.CTkFont(family=FONT_SANS, size=11, weight="bold"),
            text_color=BROWN
        ).pack(side="left")

        btn_box = ctk.CTkFrame(top, fg_color="transparent")
        btn_box.pack(side="right")

        self.img_btn_pick = ctk.CTkButton(
            btn_box,
            text="+ Chọn File Excel...",
            font=ctk.CTkFont(family=FONT_SANS, size=11),
            fg_color="transparent",
            text_color=ROAST,
            hover_color=SAND,
            border_width=1,
            border_color=LINE,
            corner_radius=14,
            width=150,
            height=28,
            command=self.img_browse_file
        )
        self.img_btn_pick.pack(side="left", padx=4)

        self.img_btn_clear = ctk.CTkButton(
            btn_box,
            text="Xóa File",
            font=ctk.CTkFont(family=FONT_SANS, size=11),
            fg_color="transparent",
            text_color=DANGER,
            hover_color=SAND,
            border_width=0,
            corner_radius=14,
            width=70,
            height=28,
            command=self.img_clear_file
        )
        self.img_btn_clear.pack(side="left", padx=4)

        ctk.CTkFrame(self.img_card, height=1, fg_color=LINE).pack(fill="x", padx=24, pady=(0, 14))

        self.img_drop_zone = ctk.CTkFrame(
            self.img_card, fg_color=SAND, corner_radius=12, border_width=1, border_color=FAINT
        )
        self.img_drop_zone.pack(fill="x", padx=24, pady=(0, 20))
        self.img_render_file()

        # ---------- ACTION ----------
        act = ctk.CTkFrame(scroll, fg_color="transparent")
        act.pack(fill="x", pady=(4, 14))

        self.img_btn_run = ctk.CTkButton(
            act,
            text=IMG_BTN_TEXT,
            font=ctk.CTkFont(family=FONT_SANS, size=13, weight="bold"),
            fg_color=ROAST,
            hover_color=MOSS,
            text_color=IVORY,
            corner_radius=24,
            height=48,
            command=self.img_start
        )
        self.img_btn_run.pack(fill="x")

        self.img_prog = ctk.CTkProgressBar(scroll, progress_color=MOSS, fg_color=LINE, height=3, corner_radius=2)
        self.img_prog.set(0)
        self.img_prog.pack(fill="x", pady=(0, 20))

        # ---------- CARD 2: Progress & result ----------
        card_res = ctk.CTkFrame(scroll, fg_color=CARD, corner_radius=16, border_width=1, border_color=LINE)
        card_res.pack(fill="both", expand=True, pady=(0, 8))

        head = ctk.CTkFrame(card_res, fg_color="transparent")
        head.pack(fill="x", padx=24, pady=(18, 10))

        ctk.CTkLabel(
            head,
            text="TIẾN ĐỘ & KẾT QUẢ",
            font=ctk.CTkFont(family=FONT_SANS, size=11, weight="bold"),
            text_color=BROWN
        ).pack(side="left")

        self.img_status = ctk.CTkLabel(
            head, text="Sẵn sàng",
            font=ctk.CTkFont(family=FONT_SANS, size=11),
            text_color=BROWN, corner_radius=10
        )
        self.img_status.pack(side="right")

        ctk.CTkFrame(card_res, height=1, fg_color=LINE).pack(fill="x", padx=24, pady=(0, 14))

        self.img_feed = ctk.CTkFrame(card_res, fg_color=SAND, corner_radius=12, border_width=1, border_color=FAINT)
        self.img_feed.pack(fill="both", expand=True, padx=24, pady=(0, 20))
        self.img_render_feed()

    # ---------- Renderers ----------
    def img_render_file(self):
        for child in self.img_drop_zone.winfo_children():
            child.destroy()

        if self.img_file and os.path.exists(self.img_file):
            self.img_btn_clear.configure(state="normal", text_color=DANGER)
            row = ctk.CTkFrame(self.img_drop_zone, fg_color=CARD, corner_radius=10, border_width=1, border_color=LINE)
            row.pack(fill="x", padx=10, pady=10)

            ctk.CTkLabel(row, text="📊", font=ctk.CTkFont(size=14), text_color=CARAMEL).pack(side="left", padx=(12, 8), pady=8)

            info = ctk.CTkFrame(row, fg_color="transparent")
            info.pack(side="left", fill="x", expand=True, pady=6)

            ctk.CTkLabel(
                info, text=os.path.basename(self.img_file),
                font=ctk.CTkFont(family=FONT_SERIF, size=12, weight="bold"),
                text_color=ROAST, anchor="w"
            ).pack(anchor="w")

            ctk.CTkLabel(
                info, text=f"{self.img_file}  •  {format_file_size(os.path.getsize(self.img_file))}",
                font=ctk.CTkFont(family=FONT_SANS, size=10),
                text_color=BROWN, anchor="w"
            ).pack(anchor="w", pady=(2, 0))

            ctk.CTkButton(
                row, text="✕",
                font=ctk.CTkFont(size=11, weight="bold"),
                fg_color="transparent", text_color=BROWN, hover_color=SAND,
                border_width=0, corner_radius=12, width=24, height=24,
                command=self.img_clear_file
            ).pack(side="right", padx=(4, 12), pady=8)
        else:
            self.img_btn_clear.configure(state="disabled", text_color=FAINT)
            ctk.CTkLabel(
                self.img_drop_zone,
                text="📥  Kéo thả file Excel vào đây",
                font=ctk.CTkFont(family=FONT_SANS, size=12),
                text_color=BROWN,
                pady=20
            ).pack(fill="both", expand=True)

    def img_render_feed(self):
        for child in self.img_feed.winfo_children():
            child.destroy()

        self.img_steps = []
        self.img_active_step = -1
        titles = ("Đọc danh sách ảnh", "Tải & nén ảnh HD", "Chèn ảnh vào ô (Place in Cell)")

        steps_box = ctk.CTkFrame(self.img_feed, fg_color="transparent")
        steps_box.pack(fill="x", padx=14, pady=(12, 4))

        for t in titles:
            row = ctk.CTkFrame(steps_box, fg_color="transparent")
            row.pack(fill="x", pady=4)

            dot = ctk.CTkLabel(row, text="○", width=22, font=ctk.CTkFont(family=FONT_SANS, size=13), text_color=FAINT)
            dot.pack(side="left", anchor="n")

            txt = ctk.CTkFrame(row, fg_color="transparent")
            txt.pack(side="left", fill="x", expand=True, padx=(6, 0))

            title = ctk.CTkLabel(txt, text=t, font=ctk.CTkFont(family=FONT_SANS, size=12, weight="bold"), text_color=FAINT, anchor="w")
            title.pack(anchor="w")

            detail = ctk.CTkLabel(txt, text="", font=ctk.CTkFont(family=FONT_SANS, size=10), text_color=BROWN, anchor="w", height=14)
            detail.pack(anchor="w")

            self.img_steps.append((dot, title, detail))

        self.img_result_box = ctk.CTkFrame(self.img_feed, fg_color="transparent")
        self.img_result_box.pack(fill="x")

        ctk.CTkLabel(
            self.img_result_box,
            text="Cần Excel 365  ·  Không copy/paste trong lúc đang chèn ảnh",
            font=ctk.CTkFont(family=FONT_SANS, size=10),
            text_color=FAINT
        ).pack(anchor="w", padx=16, pady=(4, 12))

    def img_set_step(self, idx, state, detail=None):
        dot, title, det = self.img_steps[idx]
        look = {
            "pending": ("○", FAINT, FAINT),
            "active": ("●", MOSS, ROAST),
            "done": ("✓", MOSS, ROAST),
            "error": ("✕", DANGER, DANGER),
        }[state]
        dot.configure(text=look[0], text_color=look[1])
        title.configure(text_color=look[2])
        if detail is not None:
            det.configure(text=detail)
        if state == "active":
            self.img_active_step = idx

    def img_set_status(self, text, color, bg="transparent"):
        self.img_status.configure(text=f"  {text}  " if bg != "transparent" else text, text_color=color, fg_color=bg)

    def img_ui(self, fn, *args):
        """Thread-safe: schedule a UI update on the Tk main thread."""
        self.after(0, lambda: fn(*args))

    # ---------- File selection ----------
    def img_set_file(self, path):
        if self.img_running:
            return
        name = os.path.basename(path)
        if not os.path.isfile(path) or name.startswith("~$") or os.path.splitext(path)[1].lower() not in IMG_EXTENSIONS:
            messagebox.showwarning("Định dạng không hợp lệ", "Vui lòng chọn file Excel (.xlsx, .xlsm hoặc .xls).")
            return
        self.img_file = os.path.abspath(path)
        self.img_render_file()
        self.img_render_feed()
        self.img_prog.set(0)
        self.img_set_status("Sẵn sàng", BROWN)

    def img_browse_file(self):
        path = filedialog.askopenfilename(
            title="Chọn file Excel",
            filetypes=[("Excel Files", "*.xlsx *.xlsm *.xls")]
        )
        if path:
            self.img_set_file(path)

    def img_clear_file(self):
        if self.img_running:
            return
        self.img_file = ""
        self.img_render_file()
        self.img_render_feed()
        self.img_prog.set(0)
        self.img_set_status("Sẵn sàng", BROWN)

    def on_drop_image(self, event):
        self.img_drop_zone.configure(border_color=FAINT, fg_color=SAND)
        paths = [p for p in parse_drop_paths(event.data) if os.path.isfile(p)]
        if paths:
            self.img_set_file(paths[0])

    # ---------- Processing ----------
    def img_start(self):
        if self.img_running:
            return
        if not self.img_file or not os.path.exists(self.img_file):
            messagebox.showwarning("Thông báo", "Vui lòng chọn file Excel trước khi bắt đầu.")
            return

        self.img_running = True
        self.img_btn_run.configure(state="disabled", fg_color=FAINT, text="Đang chèn ảnh...")
        self.img_btn_pick.configure(state="disabled")
        self.img_btn_clear.configure(state="disabled", text_color=FAINT)
        self.img_prog.set(0)
        self.img_render_feed()
        self.img_set_status("Đang xử lý", BROWN, SAND)

        threading.Thread(target=self.img_worker, daemon=True).start()

    def img_worker(self):
        src = self.img_file
        base, ext = os.path.splitext(os.path.basename(src))
        out_file = os.path.join(os.path.dirname(src), f"{base} - With Images{ext}")

        excel = None
        wb = None
        com_ready = False
        out_created = False
        failed = False
        pythoncom = None

        try:
            # Lazy import: only machines that use this module need pywin32
            try:
                import pythoncom
                import win32com.client
            except ImportError:
                raise RuntimeError("Máy chưa có thư viện pywin32. Vui lòng chạy lại file Cai_Dat_Boring_Task.bat.")

            pythoncom.CoInitialize()
            com_ready = True

            # ---- Step 1: read image links ----
            self.img_ui(self.img_set_step, 0, "active", "Đang mở file...")
            os.makedirs(IMAGE_CACHE_DIR, exist_ok=True)

            try:
                shutil.copyfile(src, out_file)
            except PermissionError:
                raise RuntimeError(f"File '{os.path.basename(out_file)}' đang mở trong Excel. Vui lòng đóng file rồi thử lại.")
            out_created = True

            # Separate Excel instance: never touches workbooks the user has open
            excel = win32com.client.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            excel.ScreenUpdating = False

            wb = excel.Workbooks.Open(out_file)
            ws = wb.Worksheets(1)
            ws.Activate()

            used = ws.UsedRange
            last_row = used.Row + used.Rows.Count - 1
            last_col = used.Column + used.Columns.Count - 1
            if last_row < 2:
                raise ValueError("File không có dòng dữ liệu nào.")

            headers = com_rows(ws.Range(ws.Cells(1, 1), ws.Cells(1, last_col)).Value)[0]
            pairs = []
            for col_idx, h in enumerate(headers, 1):
                if h is not None and str(h).strip() in IMG_LINK_COLUMNS:
                    pairs.append((col_idx, col_idx + 1))  # link column -> picture column on its right

            if not pairs:
                raise ValueError("Không tìm thấy cột link ảnh (Comp. Pic. 1, Pic 2 … Pic 5) ở dòng tiêu đề.")

            links = {}
            all_unc = set()
            for url_col, _ in pairs:
                rows = com_rows(ws.Range(ws.Cells(2, url_col), ws.Cells(last_row, url_col)).Value)
                col_vals = [img_url_to_unc(str(r[0])) if r[0] else None for r in rows]
                links[url_col] = col_vals
                all_unc.update(u for u in col_vals if u)

            total_imgs = len(all_unc)
            num_rows = last_row - 1
            self.img_ui(self.img_set_step, 0, "done", f"{total_imgs} ảnh trong {num_rows} dòng")
            self.img_ui(self.img_prog.set, 0.08)

            # ---- Step 2: parallel download & compress ----
            self.img_ui(self.img_set_step, 1, "active", f"0/{total_imgs} ảnh")
            ready = 0
            done = 0
            with ThreadPoolExecutor(max_workers=16) as pool:
                futures = [pool.submit(img_prepare_thumb, p) for p in all_unc]
                for fut in as_completed(futures):
                    done += 1
                    if fut.result():
                        ready += 1
                    if done % 10 == 0 or done == total_imgs:
                        pct = done / max(total_imgs, 1)
                        self.img_ui(self.img_set_step, 1, "active", f"{done}/{total_imgs} ảnh ({int(pct * 100)}%)")
                        self.img_ui(self.img_prog.set, 0.08 + pct * 0.42)

            missing = total_imgs - ready
            detail = f"Sẵn sàng {ready}/{total_imgs} ảnh"
            if missing:
                detail += f"  ·  {missing} ảnh không truy cập được"
            self.img_ui(self.img_set_step, 1, "done", detail)

            # ---- Step 3: embed into cells ----
            self.img_ui(self.img_set_step, 2, "active", f"0/{num_rows} dòng")
            for _, pic_col in pairs:
                ws.Columns(pic_col).ColumnWidth = 14

            inserted = 0
            attempts = 0
            for r in range(2, last_row + 1):
                i = r - 2
                row_has_pic = False
                for url_col, pic_col in pairs:
                    unc = links[url_col][i]
                    if not unc:
                        continue
                    thumb = img_thumb_path(unc)
                    if not os.path.exists(thumb):
                        continue
                    attempts += 1
                    try:
                        cell = ws.Cells(r, pic_col)
                        cell.Select()
                        pic = ws.Pictures().Insert(thumb)
                        pic.Select()
                        pic.Copy()
                        pic.Delete()
                        time.sleep(0.03)
                        cell.PastePictureInCell()
                        inserted += 1
                        row_has_pic = True
                    except Exception:
                        # First few all failing => this Excel has no Place in Cell support
                        if inserted == 0 and attempts >= 3:
                            raise RuntimeError("Excel trên máy chưa hỗ trợ Place in Cell. Cần Microsoft Excel 365 bản mới.")
                if row_has_pic:
                    ws.Rows(r).RowHeight = 65

                if r % 10 == 0 or r == last_row:
                    pct = (r - 1) / max(num_rows, 1)
                    self.img_ui(self.img_set_step, 2, "active", f"{r - 1}/{num_rows} dòng  ·  đã chèn {inserted} ảnh")
                    self.img_ui(self.img_prog.set, 0.50 + pct * 0.50)

            excel.ScreenUpdating = True
            wb.Save()
            wb.Close(False)
            wb = None
            excel.Quit()
            excel = None

            self.img_ui(self.img_set_step, 2, "done", f"Đã chèn {inserted} ảnh")
            self.img_ui(self.img_finish, out_file, inserted, missing, "")

        except Exception as e:
            failed = True
            self.img_ui(self.img_finish, out_file, 0, 0, str(e))

        finally:
            if wb is not None:
                try:
                    wb.Close(False)
                except Exception:
                    pass
            if excel is not None:
                try:
                    excel.Quit()
                except Exception:
                    pass
            if failed and out_created:
                try:
                    os.remove(out_file)  # don't leave a half-finished result behind
                except Exception:
                    pass
            if com_ready and pythoncom is not None:
                try:
                    pythoncom.CoUninitialize()
                except Exception:
                    pass

    def img_finish(self, out_file, inserted, missing, err_msg):
        self.img_running = False
        self.img_btn_run.configure(state="normal", fg_color=ROAST, text=IMG_BTN_TEXT)
        self.img_btn_pick.configure(state="normal")
        self.img_btn_clear.configure(state="normal", text_color=DANGER)

        for child in self.img_result_box.winfo_children():
            child.destroy()

        if err_msg:
            self.img_prog.set(0)
            self.img_set_status("❌ Lỗi", DANGER)
            if self.img_active_step >= 0:
                self.img_set_step(self.img_active_step, "error")
            err_box = ctk.CTkFrame(self.img_result_box, fg_color=CARD, corner_radius=10, border_width=1, border_color=DANGER)
            err_box.pack(fill="x", padx=10, pady=(4, 10))
            ctk.CTkLabel(
                err_box, text=err_msg,
                font=ctk.CTkFont(family=FONT_SANS, size=11),
                text_color=DANGER, wraplength=520, justify="left"
            ).pack(padx=12, pady=10, anchor="w")
            return

        self.img_prog.set(1.0)
        self.img_set_status("✓ Hoàn tất", MOSS, MOSS_SOFT)

        banner = ctk.CTkFrame(self.img_result_box, fg_color=MOSS_SOFT, corner_radius=12, border_width=1, border_color=MOSS)
        banner.pack(fill="x", padx=10, pady=(4, 10))

        left = ctk.CTkFrame(banner, fg_color="transparent")
        left.pack(side="left", fill="x", expand=True, padx=16, pady=12)

        ctk.CTkLabel(
            left, text=f"Đã chèn {inserted} ảnh vào file.",
            font=ctk.CTkFont(family=FONT_SERIF, size=13, weight="bold"),
            text_color=ROAST
        ).pack(anchor="w")

        sub = f"Đã tạo file: {out_file}"
        if missing:
            sub += f"\n{missing} ảnh không truy cập được trên server."
        ctk.CTkLabel(
            left, text=sub,
            font=ctk.CTkFont(family=FONT_SANS, size=11),
            text_color=BROWN, wraplength=440, justify="left"
        ).pack(anchor="w", pady=(3, 0))

        btns = ctk.CTkFrame(banner, fg_color="transparent")
        btns.pack(side="right", padx=16, pady=12)

        ctk.CTkButton(
            btns, text="Mở File",
            font=ctk.CTkFont(family=FONT_SANS, size=11, weight="bold"),
            fg_color=ROAST, text_color=IVORY, hover_color=MOSS,
            corner_radius=16, height=32, width=96,
            command=lambda p=out_file: os.startfile(p)
        ).pack(pady=(0, 6))

        ctk.CTkButton(
            btns, text="📂 Xem File",
            font=ctk.CTkFont(family=FONT_SANS, size=11),
            fg_color="transparent", text_color=ROAST, hover_color=SAND,
            border_width=1, border_color=LINE,
            corner_radius=16, height=32, width=96,
            command=lambda p=out_file: reveal_in_explorer(p)
        ).pack()

        # Open the result automatically, as the original tool did
        try:
            os.startfile(out_file)
        except Exception:
            pass

    # ====================================================
    # MODULE: Fabric Mapping Checker
    # ====================================================
    def build_fabric_checker(self):
        self.build_feature_nav(self.fabric_view, "Fabric Checker")

        scroll = ctk.CTkScrollableFrame(self.fabric_view, fg_color="transparent")
        scroll.pack(fill="both", expand=True, padx=36, pady=(0, 16))

        # ========================================================
        # CARD 1: Rule File (Fabric_name.xlsx)
        # ========================================================
        self.card_fabric_rule = ctk.CTkFrame(scroll, fg_color=CARD, corner_radius=16, border_width=1, border_color=LINE)
        self.card_fabric_rule.pack(fill="x", pady=(0, 20))

        top_rule = ctk.CTkFrame(self.card_fabric_rule, fg_color="transparent")
        top_rule.pack(fill="x", padx=24, pady=(18, 10))

        ctk.CTkLabel(
            top_rule,
            text="FILE QUY TẮC VẢI (FABRIC RULES)",
            font=ctk.CTkFont(family=FONT_SANS, size=11, weight="bold"),
            text_color=BROWN
        ).pack(side="left")

        btn_rule_box = ctk.CTkFrame(top_rule, fg_color="transparent")
        btn_rule_box.pack(side="right")

        self.btn_pick_rule = ctk.CTkButton(
            btn_rule_box,
            text="+ Chọn File Quy Tắc...",
            font=ctk.CTkFont(family=FONT_SANS, size=11),
            fg_color="transparent",
            text_color=ROAST,
            hover_color=SAND,
            border_width=1,
            border_color=LINE,
            corner_radius=14,
            width=165,
            height=28,
            command=self.fabric_browse_rule_file
        )
        self.btn_pick_rule.pack(side="left", padx=4)

        self.btn_clear_rule = ctk.CTkButton(
            btn_rule_box,
            text="Xóa File",
            font=ctk.CTkFont(family=FONT_SANS, size=11),
            fg_color="transparent",
            text_color=DANGER,
            hover_color=SAND,
            border_width=0,
            corner_radius=14,
            width=70,
            height=28,
            command=self.fabric_clear_rule_file
        )
        self.btn_clear_rule.pack(side="left", padx=4)

        ctk.CTkFrame(self.card_fabric_rule, height=1, fg_color=LINE).pack(fill="x", padx=24, pady=(0, 14))

        self.fabric_rule_status_box = ctk.CTkFrame(
            self.card_fabric_rule, fg_color=SAND, corner_radius=12, border_width=1, border_color=FAINT
        )
        self.fabric_rule_status_box.pack(fill="x", padx=24, pady=(0, 20))
        self.fabric_render_rule_status()

        # ========================================================
        # CARD 2: File Overview Cần Kiểm Tra
        # ========================================================
        self.card_fabric_ov = ctk.CTkFrame(scroll, fg_color=CARD, corner_radius=16, border_width=1, border_color=LINE)
        self.card_fabric_ov.pack(fill="x", pady=(0, 20))

        top_ov = ctk.CTkFrame(self.card_fabric_ov, fg_color="transparent")
        top_ov.pack(fill="x", padx=24, pady=(18, 10))

        ctk.CTkLabel(
            top_ov,
            text="FILE OVERVIEW CẦN KIỂM TRA",
            font=ctk.CTkFont(family=FONT_SANS, size=11, weight="bold"),
            text_color=BROWN
        ).pack(side="left")

        btn_ov_box = ctk.CTkFrame(top_ov, fg_color="transparent")
        btn_ov_box.pack(side="right")

        self.btn_pick_fabric_ov = ctk.CTkButton(
            btn_ov_box,
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
            command=self.fabric_browse_overview_file
        )
        self.btn_pick_fabric_ov.pack(side="left", padx=4)

        self.btn_clear_fabric_ov = ctk.CTkButton(
            btn_ov_box,
            text="Xóa File",
            font=ctk.CTkFont(family=FONT_SANS, size=11),
            fg_color="transparent",
            text_color=DANGER,
            hover_color=SAND,
            border_width=0,
            corner_radius=14,
            width=70,
            height=28,
            command=self.fabric_clear_overview_file
        )
        self.btn_clear_fabric_ov.pack(side="left", padx=4)

        ctk.CTkFrame(self.card_fabric_ov, height=1, fg_color=LINE).pack(fill="x", padx=24, pady=(0, 14))

        self.fabric_ov_drop_zone = ctk.CTkFrame(
            self.card_fabric_ov, fg_color=SAND, corner_radius=12, border_width=1, border_color=FAINT
        )
        self.fabric_ov_drop_zone.pack(fill="x", padx=24, pady=(0, 20))
        self.fabric_render_overview_file()

        # ========================================================
        # ACTION BUTTON & PROGRESS
        # ========================================================
        act = ctk.CTkFrame(scroll, fg_color="transparent")
        act.pack(fill="x", pady=(4, 14))

        self.fabric_btn_run = ctk.CTkButton(
            act,
            text="Bắt Đầu Kiểm Tra Khớp Vải (Tạo File _checked)",
            font=ctk.CTkFont(family=FONT_SANS, size=13, weight="bold"),
            fg_color=ROAST,
            hover_color=MOSS,
            text_color=IVORY,
            corner_radius=24,
            height=48,
            command=self.fabric_start_check
        )
        self.fabric_btn_run.pack(fill="x")

        self.fabric_prog = ctk.CTkProgressBar(scroll, progress_color=MOSS, fg_color=LINE, height=3, corner_radius=2)
        self.fabric_prog.set(0)
        self.fabric_prog.pack(fill="x", pady=(0, 20))

        # ========================================================
        # CARD 3: TIẾN ĐỘ & KẾT QUẢ
        # ========================================================
        card_res = ctk.CTkFrame(scroll, fg_color=CARD, corner_radius=16, border_width=1, border_color=LINE)
        card_res.pack(fill="both", expand=True, pady=(0, 8))

        head = ctk.CTkFrame(card_res, fg_color="transparent")
        head.pack(fill="x", padx=24, pady=(18, 10))

        ctk.CTkLabel(
            head,
            text="TIẾN ĐỘ & KẾT QUẢ",
            font=ctk.CTkFont(family=FONT_SANS, size=11, weight="bold"),
            text_color=BROWN
        ).pack(side="left")

        self.fabric_status_lbl = ctk.CTkLabel(
            head, text="Sẵn sàng",
            font=ctk.CTkFont(family=FONT_SANS, size=11),
            text_color=BROWN
        )
        self.fabric_status_lbl.pack(side="right")

        ctk.CTkFrame(card_res, height=1, fg_color=LINE).pack(fill="x", padx=24, pady=(0, 14))

        self.fabric_feed = ctk.CTkFrame(card_res, fg_color=SAND, corner_radius=12, border_width=1, border_color=FAINT)
        self.fabric_feed.pack(fill="both", expand=True, padx=24, pady=(0, 20))
        self.fabric_render_initial_feed()

    # ---------- Fabric Checker Renderers & Handlers ----------
    def fabric_render_rule_status(self, delta_info=None):
        for child in self.fabric_rule_status_box.winfo_children():
            child.destroy()

        rules = self.fabric_rule_mgr.rules
        if rules:
            self.btn_clear_rule.configure(state="normal", text_color=DANGER)
            row = ctk.CTkFrame(self.fabric_rule_status_box, fg_color=CARD, corner_radius=10, border_width=1, border_color=LINE)
            row.pack(fill="x", padx=10, pady=10)

            ctk.CTkLabel(row, text="📖", font=ctk.CTkFont(size=14), text_color=CARAMEL).pack(side="left", padx=(12, 8), pady=8)

            info = ctk.CTkFrame(row, fg_color="transparent")
            info.pack(side="left", fill="x", expand=True, pady=6)

            title_row = ctk.CTkFrame(info, fg_color="transparent")
            title_row.pack(anchor="w")

            src_file = self.fabric_rule_mgr.meta.get("source_file", "Fabric_name.xlsx")
            ctk.CTkLabel(
                title_row,
                text=src_file,
                font=ctk.CTkFont(family=FONT_SERIF, size=12, weight="bold"),
                text_color=ROAST
            ).pack(side="left")

            # Pill Badge matching Purchase Cost
            ctk.CTkLabel(
                title_row,
                text=f"  ✓ ĐÃ GHI NHỚ ({len(rules)} cặp)  ",
                font=ctk.CTkFont(family=FONT_SANS, size=9, weight="bold"),
                fg_color=MOSS_SOFT,
                text_color=MOSS,
                corner_radius=6
            ).pack(side="left", padx=(8, 0))

            updated_at = self.fabric_rule_mgr.meta.get("updated_at", "Chưa rõ")
            sub_msg = f"Cập nhật lúc: {updated_at}"
            if delta_info:
                add_cnt = delta_info.get("added_count", 0)
                if add_cnt > 0:
                    sub_msg += f"  •  Vừa thêm {add_cnt} cặp mới"

            ctk.CTkLabel(
                info,
                text=sub_msg,
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
                command=self.fabric_clear_rule_file
            ).pack(side="right", padx=(4, 12), pady=8)
        else:
            self.btn_clear_rule.configure(state="disabled", text_color=FAINT)
            ctk.CTkLabel(
                self.fabric_rule_status_box,
                text="📥  Kéo thả file Fabric_name.xlsx vào đây để lưu quy tắc",
                font=ctk.CTkFont(family=FONT_SANS, size=12),
                text_color=BROWN,
                pady=20
            ).pack(fill="both", expand=True)

    def fabric_browse_rule_file(self):
        if self.fabric_running:
            return
        path = filedialog.askopenfilename(
            title="Chọn file quy tắc vải (Fabric_name.xlsx)",
            filetypes=[("Excel Files", "*.xlsx *.xlsm")]
        )
        if path:
            self.fabric_load_rule_file(path)

    def fabric_load_rule_file(self, path):
        if not os.path.isfile(path) or os.path.splitext(path)[1].lower() not in (".xlsx", ".xlsm"):
            messagebox.showwarning("Định dạng không hợp lệ", "Vui lòng chọn file Excel quy tắc (.xlsx hoặc .xlsm).")
            return
        try:
            delta = self.fabric_rule_mgr.import_rule_file(path)
            self.fabric_render_rule_status(delta_info=delta)
            self.fabric_set_status("Đã cập nhật quy tắc", MOSS, MOSS_SOFT)
        except Exception as e:
            messagebox.showerror("Lỗi đọc file quy tắc", f"Không thể đọc file quy tắc:\n{e}")

    def fabric_clear_rule_file(self):
        if self.fabric_running:
            return
        self.fabric_rule_mgr.rules = []
        self.fabric_rule_mgr.meta = {}
        self.fabric_rule_mgr.save_rules()
        self.fabric_render_rule_status()
        self.fabric_set_status("Sẵn sàng", BROWN)

    def on_drop_fabric_rule(self, event):
        self.fabric_rule_status_box.configure(border_color=FAINT, fg_color=SAND)
        paths = parse_drop_paths(event.data)
        for p in paths:
            if os.path.isfile(p) and p.lower().endswith((".xlsx", ".xlsm")):
                self.fabric_load_rule_file(p)
                return

    def fabric_render_overview_file(self):
        for child in self.fabric_ov_drop_zone.winfo_children():
            child.destroy()

        if self.fabric_overview_file and os.path.exists(self.fabric_overview_file):
            self.btn_clear_fabric_ov.configure(state="normal", text_color=DANGER)
            row = ctk.CTkFrame(self.fabric_ov_drop_zone, fg_color=CARD, corner_radius=10, border_width=1, border_color=LINE)
            row.pack(fill="x", padx=10, pady=10)

            ctk.CTkLabel(row, text="📊", font=ctk.CTkFont(size=14), text_color=CARAMEL).pack(side="left", padx=(12, 8), pady=8)

            info = ctk.CTkFrame(row, fg_color="transparent")
            info.pack(side="left", fill="x", expand=True, pady=6)

            title_row = ctk.CTkFrame(info, fg_color="transparent")
            title_row.pack(anchor="w")

            ctk.CTkLabel(
                title_row, text=os.path.basename(self.fabric_overview_file),
                font=ctk.CTkFont(family=FONT_SERIF, size=12, weight="bold"),
                text_color=ROAST
            ).pack(side="left")

            ctk.CTkLabel(
                title_row,
                text="  SẴN SÀNG  ",
                font=ctk.CTkFont(family=FONT_SANS, size=9, weight="bold"),
                fg_color=MOSS_SOFT,
                text_color=MOSS,
                corner_radius=6
            ).pack(side="left", padx=(8, 0))

            ctk.CTkLabel(
                info, text=f"{self.fabric_overview_file}  •  {format_file_size(os.path.getsize(self.fabric_overview_file))}",
                font=ctk.CTkFont(family=FONT_SANS, size=10),
                text_color=BROWN, anchor="w"
            ).pack(anchor="w", pady=(2, 0))

            ctk.CTkButton(
                row, text="✕",
                font=ctk.CTkFont(size=11, weight="bold"),
                fg_color="transparent", text_color=BROWN, hover_color=SAND,
                border_width=0, corner_radius=12, width=24, height=24,
                command=self.fabric_clear_overview_file
            ).pack(side="right", padx=(4, 12), pady=8)
        else:
            self.btn_clear_fabric_ov.configure(state="disabled", text_color=FAINT)
            ctk.CTkLabel(
                self.fabric_ov_drop_zone,
                text="📥  Kéo thả file Overview.xlsx vào đây",
                font=ctk.CTkFont(family=FONT_SANS, size=12),
                text_color=BROWN,
                pady=20
            ).pack(fill="both", expand=True)

    def fabric_browse_overview_file(self):
        if self.fabric_running:
            return
        path = filedialog.askopenfilename(
            title="Chọn file Overview cần kiểm tra",
            filetypes=[("Excel Files", "*.xlsx *.xlsm")]
        )
        if path:
            self.fabric_set_overview_file(path)

    def fabric_set_overview_file(self, path):
        if self.fabric_running:
            return
        if not os.path.isfile(path) or os.path.splitext(path)[1].lower() not in (".xlsx", ".xlsm"):
            messagebox.showwarning("Định dạng không hợp lệ", "Vui lòng chọn file Excel (.xlsx hoặc .xlsm).")
            return
        self.fabric_overview_file = os.path.abspath(path)
        self.fabric_render_overview_file()
        self.fabric_render_initial_feed()
        self.fabric_prog.set(0)
        self.fabric_set_status("Sẵn sàng", BROWN)

    def fabric_clear_overview_file(self):
        if self.fabric_running:
            return
        self.fabric_overview_file = ""
        self.fabric_render_overview_file()
        self.fabric_render_initial_feed()
        self.fabric_prog.set(0)
        self.fabric_set_status("Sẵn sàng", BROWN)

    def on_drop_fabric_overview(self, event):
        self.fabric_ov_drop_zone.configure(border_color=FAINT, fg_color=SAND)
        paths = parse_drop_paths(event.data)
        for p in paths:
            if os.path.isfile(p) and p.lower().endswith((".xlsx", ".xlsm")):
                self.fabric_set_overview_file(p)
                return

    def fabric_set_status(self, text, color, bg="transparent"):
        self.fabric_status_lbl.configure(text=f"  {text}  " if bg != "transparent" else text, text_color=color, fg_color=bg)

    def fabric_render_initial_feed(self):
        for child in self.fabric_feed.winfo_children():
            child.destroy()
        hint = ctk.CTkLabel(
            self.fabric_feed,
            text="Kết quả kiểm tra khớp vải và báo cáo thống kê sẽ hiển thị tại đây.",
            font=ctk.CTkFont(family=FONT_SANS, size=12),
            text_color=BROWN,
            pady=24
        )
        hint.pack(fill="both", expand=True)

    def fabric_start_check(self):
        if self.fabric_running:
            return
        if not self.fabric_overview_file or not os.path.exists(self.fabric_overview_file):
            messagebox.showwarning("Thiếu file Overview", "Vui lòng chọn hoặc kéo thả file Overview.xlsx trước.")
            return
        if not self.fabric_rule_mgr.rules:
            messagebox.showwarning("Thiếu quy tắc vải", "Chưa có quy tắc vải nào được nạp. Vui lòng nạp file Fabric_name.xlsx trước.")
            return

        self.fabric_running = True
        self.fabric_btn_run.configure(state="disabled", text="Đang So Khớp Dữ Liệu Vải...")
        self.btn_pick_fabric_ov.configure(state="disabled")
        self.btn_clear_fabric_ov.configure(state="disabled")
        self.btn_pick_rule.configure(state="disabled")
        self.btn_clear_rule.configure(state="disabled")
        self.fabric_prog.set(0.05)
        self.fabric_set_status("Đang kiểm tra...", MOSS)

        for child in self.fabric_feed.winfo_children():
            child.destroy()
        ctk.CTkLabel(
            self.fabric_feed,
            text="⏳ Đang phân tích so khớp từng dòng...",
            font=ctk.CTkFont(family=FONT_SANS, size=12),
            text_color=BROWN,
            pady=20
        ).pack(fill="both", expand=True)

        threading.Thread(target=self.fabric_worker, daemon=True).start()

    def fabric_worker(self):
        try:
            def on_progress(current, total):
                if total > 0:
                    pct = 0.05 + 0.90 * (current / total)
                    self.after(0, self.fabric_prog.set, pct)

            stats = check_overview_file(
                overview_path=self.fabric_overview_file,
                rule_manager=self.fabric_rule_mgr,
                fuzzy_threshold=0.85,
                progress_callback=on_progress
            )
            self.after(0, self.fabric_finish, stats, "")
        except PermissionError:
            self.after(0, self.fabric_finish, {}, "File Overview đang mở trong Excel hoặc ứng dụng khác. Vui lòng đóng file rồi thử lại.")
        except Exception as e:
            self.after(0, self.fabric_finish, {}, str(e))

    def fabric_finish(self, stats, err_msg=""):
        self.fabric_running = False
        self.fabric_btn_run.configure(state="normal", text="Bắt Đầu Kiểm Tra Khớp Vải (Tạo File _checked)")
        self.btn_pick_fabric_ov.configure(state="normal")
        self.btn_clear_fabric_ov.configure(state="normal")
        self.btn_pick_rule.configure(state="normal")
        self.btn_clear_rule.configure(state="normal")

        for child in self.fabric_feed.winfo_children():
            child.destroy()

        if err_msg:
            self.fabric_prog.set(0)
            self.fabric_set_status("❌ Lỗi", DANGER)
            err_box = ctk.CTkFrame(self.fabric_feed, fg_color=CARD, corner_radius=10, border_width=1, border_color=DANGER)
            err_box.pack(fill="x", padx=10, pady=10)
            ctk.CTkLabel(
                err_box, text=f"Đã xảy ra lỗi: {err_msg}",
                font=ctk.CTkFont(family=FONT_SANS, size=11),
                text_color=DANGER, wraplength=520, justify="left"
            ).pack(padx=14, pady=12, anchor="w")
            return

        self.fabric_prog.set(1.0)
        self.fabric_set_status("✓ Hoàn tất", MOSS, MOSS_SOFT)

        # 1. Summary Metrics Banner
        banner = ctk.CTkFrame(self.fabric_feed, fg_color=CARD, corner_radius=12, border_width=1, border_color=LINE)
        banner.pack(fill="x", padx=10, pady=(10, 8))

        top_b = ctk.CTkFrame(banner, fg_color="transparent")
        top_b.pack(fill="x", padx=16, pady=(12, 6))

        ctk.CTkLabel(
            top_b,
            text=f"Hoàn tất kiểm tra {stats['total_rows']} dòng dữ liệu",
            font=ctk.CTkFont(family=FONT_SERIF, size=14, weight="bold"),
            text_color=ROAST
        ).pack(side="left")

        # 3 Pills: OK / PARTIAL / ERROR
        pills_box = ctk.CTkFrame(top_b, fg_color="transparent")
        pills_box.pack(side="right")

        ctk.CTkLabel(
            pills_box,
            text=f"  ✓ OK: {stats['ok']} ({stats['ok_pct']:.1f}%)  ",
            font=ctk.CTkFont(family=FONT_SANS, size=10, weight="bold"),
            fg_color="#c6efce", text_color="#006100", corner_radius=10
        ).pack(side="left", padx=3)

        ctk.CTkLabel(
            pills_box,
            text=f"  ⚠ PARTIAL: {stats['partial']} ({stats['partial_pct']:.1f}%)  ",
            font=ctk.CTkFont(family=FONT_SANS, size=10, weight="bold"),
            fg_color="#ffeb9c", text_color="#9c6500", corner_radius=10
        ).pack(side="left", padx=3)

        ctk.CTkLabel(
            pills_box,
            text=f"  ✗ ERROR: {stats['error']} ({stats['error_pct']:.1f}%)  ",
            font=ctk.CTkFont(family=FONT_SANS, size=10, weight="bold"),
            fg_color="#ffc7ce", text_color="#9c0006", corner_radius=10
        ).pack(side="left", padx=3)

        # Divider
        ctk.CTkFrame(banner, height=1, fg_color=LINE).pack(fill="x", padx=16, pady=4)

        # Breakdown info & action buttons
        mid_b = ctk.CTkFrame(banner, fg_color="transparent")
        mid_b.pack(fill="x", padx=16, pady=(4, 12))

        b_info = ctk.CTkFrame(mid_b, fg_color="transparent")
        b_info.pack(side="left", fill="x", expand=True)

        ctk.CTkLabel(
            b_info,
            text=f"• Chưa điền tên vải (- / trống): {stats['empty_fabric']} dòng\n• ERROR thật sự (có vải nhưng không khớp): {stats['real_error']} dòng",
            font=ctk.CTkFont(family=FONT_SANS, size=11),
            text_color=BROWN, justify="left", anchor="w"
        ).pack(anchor="w")

        out_f = stats.get("output_file", "")
        btn_box = ctk.CTkFrame(mid_b, fg_color="transparent")
        btn_box.pack(side="right")

        if out_f:
            ctk.CTkButton(
                btn_box, text="Mở File",
                font=ctk.CTkFont(family=FONT_SANS, size=11, weight="bold"),
                fg_color=ROAST, text_color=IVORY, hover_color=MOSS,
                corner_radius=16, height=32, width=96,
                command=lambda p=out_f: os.startfile(p)
            ).pack(side="left", padx=4)

            ctk.CTkButton(
                btn_box, text="📂 Xem File",
                font=ctk.CTkFont(family=FONT_SANS, size=11, weight="bold"),
                fg_color=CARD, text_color=ROAST, hover_color=SAND,
                border_width=1, border_color=LINE,
                corner_radius=16, height=32, width=96,
                command=lambda p=out_f: reveal_in_explorer(p)
            ).pack(side="left", padx=4)

        # 2. Missing Fabrics Table (if real errors exist)
        missing_counts = stats.get("missing_fabrics", {})
        if missing_counts:
            miss_box = ctk.CTkFrame(self.fabric_feed, fg_color=CARD, corner_radius=12, border_width=1, border_color=LINE)
            miss_box.pack(fill="x", padx=10, pady=(0, 10))

            miss_head = ctk.CTkFrame(miss_box, fg_color="transparent")
            miss_head.pack(fill="x", padx=16, pady=(10, 6))

            ctk.CTkLabel(
                miss_head,
                text="DANH SÁCH TÊN VẢI LỖI THẬT SỰ (CẦN BỔ SUNG VÀO QUY TẮC)",
                font=ctk.CTkFont(family=FONT_SANS, size=10, weight="bold"),
                text_color=DANGER
            ).pack(side="left")

            ctk.CTkLabel(
                miss_head,
                text=f"{len(missing_counts)} loại vải chưa có trong mapping",
                font=ctk.CTkFont(family=FONT_SANS, size=10),
                text_color=BROWN
            ).pack(side="right")

            # Table rows
            for fab_name, count in missing_counts.most_common(8):
                t_row = ctk.CTkFrame(miss_box, fg_color="transparent")
                t_row.pack(fill="x", padx=16, pady=4)

                ctk.CTkLabel(
                    t_row, text=f"• {fab_name}",
                    font=ctk.CTkFont(family=FONT_SANS, size=11),
                    text_color=ROAST, anchor="w"
                ).pack(side="left")

                ctk.CTkLabel(
                    t_row,
                    text=f"  {count} lần  ",
                    font=ctk.CTkFont(family=FONT_SANS, size=10, weight="bold"),
                    fg_color="#ffc7ce",
                    text_color="#9c0006",
                    corner_radius=6
                ).pack(side="right")

            if len(missing_counts) > 8:
                ctk.CTkLabel(
                    miss_box,
                    text=f"...và {len(missing_counts) - 8} loại vải khác (xem chi tiết trong file Excel đã tạo).",
                    font=ctk.CTkFont(family=FONT_SANS, size=10, slant="italic"),
                    text_color=FAINT
                ).pack(padx=16, pady=(4, 8), anchor="w")

    # ====================================================
    # MODULE: Order Auditor (Multi-Layer PO Verification)
    # ====================================================
    def build_order_auditor(self):
        self.build_feature_nav(self.auditor_view, "Order Auditor")

        scroll = ctk.CTkScrollableFrame(self.auditor_view, fg_color="transparent")
        scroll.pack(fill="both", expand=True, padx=36, pady=(0, 16))

        # ========================================================
        # CARD 1: File PO Cần Duyệt (Purchase Order Thủy Lập)
        # ========================================================
        self.card_auditor_po = ctk.CTkFrame(scroll, fg_color=CARD, corner_radius=16, border_width=1, border_color=LINE)
        self.card_auditor_po.pack(fill="x", pady=(0, 20))

        top_po = ctk.CTkFrame(self.card_auditor_po, fg_color="transparent")
        top_po.pack(fill="x", padx=24, pady=(18, 10))

        ctk.CTkLabel(
            top_po,
            text="1. FILE PURCHASE ORDER (PO / PI CẦN KIỂM TOÁN)",
            font=ctk.CTkFont(family=FONT_SANS, size=11, weight="bold"),
            text_color=BROWN
        ).pack(side="left")

        btn_po_box = ctk.CTkFrame(top_po, fg_color="transparent")
        btn_po_box.pack(side="right")

        self.btn_pick_auditor_po = ctk.CTkButton(
            btn_po_box,
            text="+ Chọn File PO...",
            font=ctk.CTkFont(family=FONT_SANS, size=11),
            fg_color="transparent",
            text_color=ROAST,
            hover_color=SAND,
            border_width=1,
            border_color=LINE,
            corner_radius=14,
            width=140,
            height=28,
            command=self.auditor_browse_po_file
        )
        self.btn_pick_auditor_po.pack(side="left", padx=4)

        self.btn_clear_auditor_po = ctk.CTkButton(
            btn_po_box,
            text="Xóa File",
            font=ctk.CTkFont(family=FONT_SANS, size=11),
            fg_color="transparent",
            text_color=DANGER,
            hover_color=SAND,
            border_width=0,
            corner_radius=14,
            width=70,
            height=28,
            command=self.auditor_clear_po_file
        )
        self.btn_clear_auditor_po.pack(side="left", padx=4)

        ctk.CTkFrame(self.card_auditor_po, height=1, fg_color=LINE).pack(fill="x", padx=24, pady=(0, 14))

        self.auditor_po_status_box = ctk.CTkFrame(
            self.card_auditor_po, fg_color=SAND, corner_radius=12, border_width=1, border_color=FAINT
        )
        self.auditor_po_status_box.pack(fill="x", padx=24, pady=(0, 20))
        self.auditor_render_po_status()

        # ========================================================
        # CARD 2: File Đơn Hàng Gốc Của Khách (Customer Order / Master)
        # ========================================================
        self.card_auditor_cust = ctk.CTkFrame(scroll, fg_color=CARD, corner_radius=16, border_width=1, border_color=LINE)
        self.card_auditor_cust.pack(fill="x", pady=(0, 20))

        top_cust = ctk.CTkFrame(self.card_auditor_cust, fg_color="transparent")
        top_cust.pack(fill="x", padx=24, pady=(18, 10))

        ctk.CTkLabel(
            top_cust,
            text="2. FILE ĐƠN HÀNG GỐC CỦA KHÁCH (CUSTOMER ORDER / MASTER DATA)",
            font=ctk.CTkFont(family=FONT_SANS, size=11, weight="bold"),
            text_color=BROWN
        ).pack(side="left")

        btn_cust_box = ctk.CTkFrame(top_cust, fg_color="transparent")
        btn_cust_box.pack(side="right")

        self.btn_pick_auditor_cust = ctk.CTkButton(
            btn_cust_box,
            text="+ Chọn File Khách...",
            font=ctk.CTkFont(family=FONT_SANS, size=11),
            fg_color="transparent",
            text_color=ROAST,
            hover_color=SAND,
            border_width=1,
            border_color=LINE,
            corner_radius=14,
            width=150,
            height=28,
            command=self.auditor_browse_cust_file
        )
        self.btn_pick_auditor_cust.pack(side="left", padx=4)

        self.btn_clear_auditor_cust = ctk.CTkButton(
            btn_cust_box,
            text="Xóa File",
            font=ctk.CTkFont(family=FONT_SANS, size=11),
            fg_color="transparent",
            text_color=DANGER,
            hover_color=SAND,
            border_width=0,
            corner_radius=14,
            width=70,
            height=28,
            command=self.auditor_clear_cust_file
        )
        self.btn_clear_auditor_cust.pack(side="left", padx=4)

        ctk.CTkFrame(self.card_auditor_cust, height=1, fg_color=LINE).pack(fill="x", padx=24, pady=(0, 14))

        self.auditor_cust_status_box = ctk.CTkFrame(
            self.card_auditor_cust, fg_color=SAND, corner_radius=12, border_width=1, border_color=FAINT
        )
        self.auditor_cust_status_box.pack(fill="x", padx=24, pady=(0, 20))
        self.auditor_render_cust_status()

        # ========================================================
        # ACTION BUTTON & PROGRESS
        # ========================================================
        action_wrap = ctk.CTkFrame(scroll, fg_color="transparent")
        action_wrap.pack(fill="x", pady=(0, 20))

        self.auditor_btn_run = ctk.CTkButton(
            action_wrap,
            text="Bắt Đầu Kiểm Toán Đơn Hàng (6 Lớp Bảo Vệ)",
            font=ctk.CTkFont(family=FONT_SERIF, size=14, weight="bold"),
            fg_color=ROAST,
            text_color=IVORY,
            hover_color=MOSS,
            height=46,
            corner_radius=23,
            command=self.auditor_start_process
        )
        self.auditor_btn_run.pack(fill="x", pady=(0, 8))

        self.auditor_prog = ctk.CTkProgressBar(
            action_wrap,
            height=3,
            fg_color=LINE,
            progress_color=MOSS,
            corner_radius=2
        )
        self.auditor_prog.pack(fill="x")
        self.auditor_prog.set(0)

        # ========================================================
        # CARD 3: Results Feed
        # ========================================================
        self.card_auditor_results = ctk.CTkFrame(scroll, fg_color=CARD, corner_radius=16, border_width=1, border_color=LINE)
        self.card_auditor_results.pack(fill="both", expand=True, pady=(0, 8))

        r_head = ctk.CTkFrame(self.card_auditor_results, fg_color="transparent")
        r_head.pack(fill="x", padx=24, pady=(18, 10))

        ctk.CTkLabel(
            r_head,
            text="TIẾN ĐỘ & KẾT QUẢ KIỂM TOÁN",
            font=ctk.CTkFont(family=FONT_SANS, size=11, weight="bold"),
            text_color=BROWN
        ).pack(side="left")

        self.auditor_status_lbl = ctk.CTkLabel(
            r_head,
            text="Sẵn sàng",
            font=ctk.CTkFont(family=FONT_SANS, size=11),
            text_color=BROWN
        )
        self.auditor_status_lbl.pack(side="right")

        ctk.CTkFrame(self.card_auditor_results, height=1, fg_color=LINE).pack(fill="x", padx=24, pady=(0, 14))

        self.auditor_feed = ctk.CTkFrame(
            self.card_auditor_results, fg_color=SAND, corner_radius=12, border_width=1, border_color=FAINT
        )
        self.auditor_feed.pack(fill="both", expand=True, padx=24, pady=(0, 20))
        self.auditor_render_initial_feed()

        # Setup Drag and Drop for Order Auditor
        if getattr(self, "has_dnd", False):
            for t in [self.card_auditor_po, self.auditor_po_status_box]:
                try:
                    t.drop_target_register(tkdnd.DND_FILES)
                    t.dnd_bind('<<Drop>>', lambda e: self.auditor_on_drop_po(e))
                except Exception:
                    pass
            for t in [self.card_auditor_cust, self.auditor_cust_status_box]:
                try:
                    t.drop_target_register(tkdnd.DND_FILES)
                    t.dnd_bind('<<Drop>>', lambda e: self.auditor_on_drop_cust(e))
                except Exception:
                    pass

    def auditor_render_initial_feed(self):
        for child in self.auditor_feed.winfo_children():
            child.destroy()
        ctk.CTkLabel(
            self.auditor_feed,
            text="Chọn file PO và file Đơn hàng của khách rồi bấm 'Bắt Đầu Kiểm Toán'.\nHệ thống sẽ đối chiếu 6 lớp: SKU, Model, Fabric, Config, Set 1/2-2/2 và Giá bán.",
            font=ctk.CTkFont(family=FONT_SANS, size=11),
            text_color=BROWN,
            justify="center",
            pady=32
        ).pack(fill="both", expand=True)

    def auditor_render_po_status(self):
        for child in self.auditor_po_status_box.winfo_children():
            child.destroy()
        if self.auditor_po_file and os.path.exists(self.auditor_po_file):
            sz = format_file_size(os.path.getsize(self.auditor_po_file))
            fn = os.path.basename(self.auditor_po_file)
            row = ctk.CTkFrame(self.auditor_po_status_box, fg_color="transparent")
            row.pack(fill="x", padx=16, pady=12)

            ctk.CTkLabel(
                row, text="📄", font=ctk.CTkFont(size=16), text_color=ROAST
            ).pack(side="left", padx=(0, 8))

            info_box = ctk.CTkFrame(row, fg_color="transparent")
            info_box.pack(side="left", fill="x", expand=True)

            ctk.CTkLabel(
                info_box, text=fn, font=ctk.CTkFont(family=FONT_SANS, size=12, weight="bold"),
                text_color=ROAST, anchor="w"
            ).pack(anchor="w")

            ctk.CTkLabel(
                info_box, text=f"Kích thước: {sz}", font=ctk.CTkFont(family=FONT_SANS, size=10),
                text_color=BROWN, anchor="w"
            ).pack(anchor="w")
        else:
            ctk.CTkLabel(
                self.auditor_po_status_box,
                text="Kéo thả file Purchase Order (.xlsx) vào đây hoặc bấm nút Chọn File ở góc trên",
                font=ctk.CTkFont(family=FONT_SANS, size=11),
                text_color=BROWN,
                pady=18
            ).pack(fill="both", expand=True)

    def auditor_render_cust_status(self):
        for child in self.auditor_cust_status_box.winfo_children():
            child.destroy()
        if self.auditor_cust_file and os.path.exists(self.auditor_cust_file):
            sz = format_file_size(os.path.getsize(self.auditor_cust_file))
            fn = os.path.basename(self.auditor_cust_file)
            row = ctk.CTkFrame(self.auditor_cust_status_box, fg_color="transparent")
            row.pack(fill="x", padx=16, pady=12)

            ctk.CTkLabel(
                row, text="📂", font=ctk.CTkFont(size=16), text_color=ROAST
            ).pack(side="left", padx=(0, 8))

            info_box = ctk.CTkFrame(row, fg_color="transparent")
            info_box.pack(side="left", fill="x", expand=True)

            ctk.CTkLabel(
                info_box, text=fn, font=ctk.CTkFont(family=FONT_SANS, size=12, weight="bold"),
                text_color=ROAST, anchor="w"
            ).pack(anchor="w")

            ctk.CTkLabel(
                info_box, text=f"Kích thước: {sz}", font=ctk.CTkFont(family=FONT_SANS, size=10),
                text_color=BROWN, anchor="w"
            ).pack(anchor="w")
        else:
            ctk.CTkLabel(
                self.auditor_cust_status_box,
                text="Kéo thả file Đơn hàng gốc / Master Data (.xlsx) của khách vào đây",
                font=ctk.CTkFont(family=FONT_SANS, size=11),
                text_color=BROWN,
                pady=18
            ).pack(fill="both", expand=True)

    def auditor_browse_po_file(self):
        path = filedialog.askopenfilename(
            title="Chọn file Purchase Order",
            filetypes=[("Excel Files", "*.xlsx *.xls")]
        )
        if path:
            self.auditor_po_file = os.path.abspath(path)
            self.auditor_render_po_status()

    def auditor_clear_po_file(self):
        self.auditor_po_file = ""
        self.auditor_render_po_status()

    def auditor_browse_cust_file(self):
        path = filedialog.askopenfilename(
            title="Chọn file Đơn hàng gốc của khách",
            filetypes=[("Excel Files", "*.xlsx *.xls")]
        )
        if path:
            self.auditor_cust_file = os.path.abspath(path)
            self.auditor_render_cust_status()

    def auditor_clear_cust_file(self):
        self.auditor_cust_file = ""
        self.auditor_render_cust_status()

    def auditor_on_drop_po(self, event):
        paths = [p for p in parse_drop_paths(event.data) if os.path.isfile(p)]
        if paths:
            self.auditor_po_file = paths[0]
            self.auditor_render_po_status()

    def auditor_on_drop_cust(self, event):
        paths = [p for p in parse_drop_paths(event.data) if os.path.isfile(p)]
        if paths:
            self.auditor_cust_file = paths[0]
            self.auditor_render_cust_status()

    def auditor_start_process(self):
        if self.auditor_running:
            return
        if not self.auditor_po_file or not os.path.exists(self.auditor_po_file):
            messagebox.showwarning("Thiếu file", "Vui lòng chọn File Purchase Order (PO) cần kiểm toán.")
            return
        if not self.auditor_cust_file or not os.path.exists(self.auditor_cust_file):
            messagebox.showwarning("Thiếu file", "Vui lòng chọn File Đơn hàng gốc của khách.")
            return

        self.auditor_running = True
        self.auditor_btn_run.configure(state="disabled", fg_color=FAINT, text="Đang kiểm toán đa nguồn...")
        self.auditor_prog.set(0.15)
        self.auditor_status_lbl.configure(text="Đang xử lý...", text_color=MOSS)

        # Clear feed
        for child in self.auditor_feed.winfo_children():
            child.destroy()

        threading.Thread(target=self.auditor_worker, daemon=True).start()

    def auditor_worker(self):
        try:
            self.auditor_prog.set(0.3)
            po_doc = order_auditor.POParsedDoc(self.auditor_po_file)
            self.auditor_prog.set(0.6)
            cust_catalog = order_auditor.CustomerCatalog(self.auditor_cust_file)
            self.auditor_prog.set(0.85)

            results = self.auditor_engine.audit(po_doc, cust_catalog)

            # Export Excel
            po_dir = os.path.dirname(self.auditor_po_file)
            po_base = os.path.splitext(os.path.basename(self.auditor_po_file))[0]
            out_fn = os.path.join(po_dir, f"Audit_Report - {po_base}.xlsx")
            order_auditor.export_audit_excel(results, out_fn)
            self.auditor_last_report = out_fn

            self.after(0, self.auditor_finish_ui, len(results), results, out_fn, "")
        except Exception as e:
            self.after(0, self.auditor_finish_ui, 0, [], "", str(e))

    def auditor_finish_ui(self, total_items, results, out_file, err_msg):
        self.auditor_running = False
        self.auditor_btn_run.configure(state="normal", fg_color=ROAST, text="Bắt Đầu Kiểm Toán Đơn Hàng (6 Lớp Bảo Vệ)")
        self.auditor_prog.set(1.0 if not err_msg else 0)

        for child in self.auditor_feed.winfo_children():
            child.destroy()

        if err_msg:
            self.auditor_status_lbl.configure(text="✕ Lỗi xử lý", text_color=DANGER)
            err_box = ctk.CTkFrame(self.auditor_feed, fg_color=CARD, corner_radius=10, border_width=1, border_color=DANGER)
            err_box.pack(fill="x", padx=10, pady=10)
            ctk.CTkLabel(
                err_box, text=f"Lỗi: {err_msg}", font=ctk.CTkFont(family=FONT_SANS, size=11),
                text_color=DANGER, wraplength=520, justify="left"
            ).pack(padx=12, pady=10)
            return

        pass_c = sum(1 for r in results if r.overall_status == "PASS")
        rev_c = sum(1 for r in results if r.overall_status == "REVIEW")
        err_c = sum(1 for r in results if r.overall_status == "ERROR")

        if err_c > 0:
            self.auditor_status_lbl.configure(text=f"Phát hiện {err_c} lỗi", text_color=DANGER)
            b_bg, b_border = CARD, DANGER
        elif rev_c > 0:
            self.auditor_status_lbl.configure(text=f"Cần xem lại {rev_c} mục", text_color=BROWN)
            b_bg, b_border = SAND, LINE
        else:
            self.auditor_status_lbl.configure(text="✓ Khớp 100%", text_color=MOSS)
            b_bg, b_border = MOSS_SOFT, MOSS

        # Summary Banner
        banner = ctk.CTkFrame(self.auditor_feed, fg_color=b_bg, corner_radius=12, border_width=1, border_color=b_border)
        banner.pack(fill="x", padx=10, pady=10)

        b_left = ctk.CTkFrame(banner, fg_color="transparent")
        b_left.pack(side="left", fill="x", expand=True, padx=16, pady=12)

        ctk.CTkLabel(
            b_left,
            text=f"Đã kiểm toán {total_items} dòng sản phẩm",
            font=ctk.CTkFont(family=FONT_SERIF, size=13, weight="bold"),
            text_color=ROAST
        ).pack(anchor="w")

        stats_str = f"🟢 PASS: {pass_c} dòng  |  🟡 REVIEW: {rev_c} dòng  |  🔴 ERROR: {err_c} dòng"
        ctk.CTkLabel(
            b_left, text=stats_str,
            font=ctk.CTkFont(family=FONT_SANS, size=11, weight="bold"),
            text_color=BROWN
        ).pack(anchor="w", pady=(3, 2))

        ctk.CTkLabel(
            b_left, text=f"File báo cáo: {os.path.basename(out_file)}",
            font=ctk.CTkFont(family=FONT_SANS, size=10),
            text_color=FAINT, wraplength=420, justify="left"
        ).pack(anchor="w")

        # Action buttons
        btn_box = ctk.CTkFrame(banner, fg_color="transparent")
        btn_box.pack(side="right", padx=16, pady=12)

        ctk.CTkButton(
            btn_box, text="Mở Báo Cáo",
            font=ctk.CTkFont(family=FONT_SANS, size=11, weight="bold"),
            fg_color=ROAST, text_color=IVORY, hover_color=MOSS,
            corner_radius=16, height=32, width=105,
            command=lambda p=out_file: os.startfile(p)
        ).pack(side="left", padx=4)

        ctk.CTkButton(
            btn_box, text="📂 Xem File",
            font=ctk.CTkFont(family=FONT_SANS, size=11, weight="bold"),
            fg_color=CARD, text_color=ROAST, hover_color=SAND,
            border_width=1, border_color=LINE,
            corner_radius=16, height=32, width=96,
            command=lambda p=out_file: reveal_in_explorer(p)
        ).pack(side="left", padx=4)

        # Highlight Errors & Reviews list
        flagged = [r for r in results if r.overall_status in ["ERROR", "REVIEW"]]
        if flagged:
            list_card = ctk.CTkFrame(self.auditor_feed, fg_color=CARD, corner_radius=12, border_width=1, border_color=LINE)
            list_card.pack(fill="x", padx=10, pady=(0, 10))

            lh = ctk.CTkFrame(list_card, fg_color="transparent")
            lh.pack(fill="x", padx=16, pady=(10, 6))

            ctk.CTkLabel(
                lh,
                text=f"DANH SÁCH CHI TIẾT CÁC ĐIỂM CẦN LƯU Ý ({len(flagged)} mục)",
                font=ctk.CTkFont(family=FONT_SANS, size=10, weight="bold"),
                text_color=DANGER if err_c > 0 else BROWN
            ).pack(side="left")

            # Show top 6 items
            for item in flagged[:6]:
                p = item.po_item
                t_row = ctk.CTkFrame(list_card, fg_color="transparent")
                t_row.pack(fill="x", padx=16, pady=4)

                badge_color = "#ffc7ce" if item.overall_status == "ERROR" else "#ffeb9c"
                badge_text_color = "#9c0006" if item.overall_status == "ERROR" else "#9c6500"

                ctk.CTkLabel(
                    t_row,
                    text=f" {item.overall_status} ",
                    font=ctk.CTkFont(family=FONT_SANS, size=9, weight="bold"),
                    fg_color=badge_color,
                    text_color=badge_text_color,
                    corner_radius=4
                ).pack(side="left", padx=(0, 8))

                summary_reason = item.summary_reasons[0] if item.summary_reasons else ""
                label_text = f"Dòng {p.row_idx} ({p.md_number or p.item_no}): {summary_reason}"
                ctk.CTkLabel(
                    t_row, text=label_text,
                    font=ctk.CTkFont(family=FONT_SANS, size=10),
                    text_color=ROAST, anchor="w", wraplength=520, justify="left"
                ).pack(side="left", fill="x", expand=True)

            if len(flagged) > 6:
                ctk.CTkLabel(
                    list_card,
                    text=f"...và {len(flagged) - 6} mục khác (vui lòng mở file Báo Cáo Excel để xem toàn bộ chi tiết).",
                    font=ctk.CTkFont(family=FONT_SANS, size=10, slant="italic"),
                    text_color=FAINT
                ).pack(padx=16, pady=(4, 8), anchor="w")

if __name__ == "__main__":
    app = BoringTaskApp()
    app.mainloop()

