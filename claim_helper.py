import os
import sys
import re
import datetime
import threading
from copy import copy

import customtkinter as ctk
from tkinter import filedialog, messagebox
import tkinterdnd2 as tkdnd
import openpyxl
from openpyxl.utils import get_column_letter

# Set Apple-inspired Light Mode
ctk.set_appearance_mode("Light")
ctk.set_default_color_theme("blue")

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

class ClaimHelperAppleApp(ctk.CTk, tkdnd.TkinterDnD.DnDWrapper):
    def __init__(self):
        super().__init__()
        self.TkdndVersion = tkdnd.TkinterDnD._require(self)

        self.title("Claim Helper")
        self.geometry("820x780")
        self.minsize(780, 680)
        self.configure(fg_color="#F5F5F7")  # Signature Apple off-white

        # State Variables
        self.overview_path = ""
        self.claim_files = [] # list of absolute paths
        self.is_processing = False

        # Auto-detect Overview in current directory
        curr_dir = os.path.abspath(os.path.dirname(__file__))
        for f in os.listdir(curr_dir):
            if f.lower().startswith("overview") and f.endswith(".xlsx") and not f.startswith("~$"):
                self.overview_path = os.path.join(curr_dir, f)
                break

        # Auto-detect initial versus files in current directory
        for f in os.listdir(curr_dir):
            if "versus" in f.lower() and f.endswith(".xlsx") and not f.endswith("_backup.xlsx") and not f.startswith("~$"):
                self.claim_files.append(os.path.join(curr_dir, f))

        self.setup_ui()
        self.setup_drag_and_drop()

    def setup_ui(self):
        # 1. Header Frame
        header = ctk.CTkFrame(self, fg_color="transparent")
        header.pack(fill="x", padx=30, pady=(25, 12))

        title_lbl = ctk.CTkLabel(
            header,
            text="Claim Helper",
            font=ctk.CTkFont(family="Segoe UI", size=26, weight="bold"),
            text_color="#1D1D1F"
        )
        title_lbl.pack(anchor="w")

        sub_lbl = ctk.CTkLabel(
            header,
            text="Hệ thống tự động đối chiếu & điền Purchase Cost cho SOFACOMPANY",
            font=ctk.CTkFont(family="Segoe UI", size=12),
            text_color="#86868B"
        )
        sub_lbl.pack(anchor="w", pady=(2, 0))

        # 2. Main Scrollable Container
        self.main_scroll = ctk.CTkScrollableFrame(self, fg_color="transparent")
        self.main_scroll.pack(fill="both", expand=True, padx=30, pady=(0, 15))

        # CARD 1: Master Overview
        self.card_ov = ctk.CTkFrame(
            self.main_scroll,
            fg_color="#FFFFFF",
            corner_radius=14,
            border_width=1,
            border_color="#E5E5EA"
        )
        self.card_ov.pack(fill="x", pady=(0, 14))

        ov_top = ctk.CTkFrame(self.card_ov, fg_color="transparent")
        ov_top.pack(fill="x", padx=18, pady=(14, 6))

        ctk.CTkLabel(
            ov_top,
            text="📊  FILE MASTER OVERVIEW",
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            text_color="#0071E3"
        ).pack(side="left")

        btn_pick_ov = ctk.CTkButton(
            ov_top,
            text="Chọn File Overview...",
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            fg_color="#F2F2F7",
            text_color="#0071E3",
            hover_color="#E5E5EA",
            corner_radius=8,
            height=28,
            command=self.browse_overview
        )
        btn_pick_ov.pack(side="right")

        self.ov_path_lbl = ctk.CTkLabel(
            self.card_ov,
            text=self.get_ov_display_text(),
            font=ctk.CTkFont(family="Segoe UI", size=12),
            text_color="#1D1D1F" if self.overview_path else "#86868B",
            anchor="w"
        )
        self.ov_path_lbl.pack(fill="x", padx=18, pady=(0, 14))

        # CARD 2: Claim Versus Files (Multi-file select & Drag Drop)
        self.card_claims = ctk.CTkFrame(
            self.main_scroll,
            fg_color="#FFFFFF",
            corner_radius=14,
            border_width=1,
            border_color="#E5E5EA"
        )
        self.card_claims.pack(fill="x", pady=(0, 14))

        claim_top = ctk.CTkFrame(self.card_claims, fg_color="transparent")
        claim_top.pack(fill="x", padx=18, pady=(14, 8))

        self.claim_count_lbl = ctk.CTkLabel(
            claim_top,
            text=f"📋  DANH SÁCH FILE CLAIM VERSUS ({len(self.claim_files)})",
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            text_color="#0071E3"
        )
        self.claim_count_lbl.pack(side="left")

        actions_box = ctk.CTkFrame(claim_top, fg_color="transparent")
        actions_box.pack(side="right")

        ctk.CTkButton(
            actions_box,
            text="+ Chọn File(s)...",
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            fg_color="#F2F2F7",
            text_color="#0071E3",
            hover_color="#E5E5EA",
            corner_radius=8,
            height=26,
            command=self.browse_claim_files
        ).pack(side="left", padx=4)

        ctk.CTkButton(
            actions_box,
            text="📂 Thư Mục...",
            font=ctk.CTkFont(family="Segoe UI", size=11),
            fg_color="#F2F2F7",
            text_color="#1D1D1F",
            hover_color="#E5E5EA",
            corner_radius=8,
            height=26,
            command=self.browse_claim_folder
        ).pack(side="left", padx=4)

        ctk.CTkButton(
            actions_box,
            text="Xóa Hết",
            font=ctk.CTkFont(family="Segoe UI", size=11),
            fg_color="transparent",
            text_color="#FF3B30",
            hover_color="#FFECEB",
            corner_radius=8,
            height=26,
            command=self.clear_all_claims
        ).pack(side="left", padx=4)

        # File List Inner Container
        self.file_list_frame = ctk.CTkFrame(self.card_claims, fg_color="#F9F9FB", corner_radius=10)
        self.file_list_frame.pack(fill="x", padx=16, pady=(0, 14))

        self.render_file_list()

        # CARD 3: Action & Progress
        act_box = ctk.CTkFrame(self.main_scroll, fg_color="transparent")
        act_box.pack(fill="x", pady=(4, 12))

        self.btn_run = ctk.CTkButton(
            act_box,
            text="Bắt Đầu Đối Chiếu & Điền Giá",
            font=ctk.CTkFont(family="Segoe UI", size=14, weight="bold"),
            fg_color="#0071E3",
            hover_color="#0077ED",
            text_color="#FFFFFF",
            corner_radius=22,
            height=46,
            command=self.start_processing
        )
        self.btn_run.pack(fill="x")

        self.prog_bar = ctk.CTkProgressBar(
            self.main_scroll,
            progress_color="#0071E3",
            fg_color="#E5E5EA",
            height=4,
            corner_radius=2
        )
        self.prog_bar.set(0)
        self.prog_bar.pack(fill="x", pady=(0, 10))

        # CARD 4: Live Activity Console
        card_log = ctk.CTkFrame(
            self.main_scroll,
            fg_color="#1C1C1E",
            corner_radius=14
        )
        card_log.pack(fill="both", expand=True, pady=(0, 6))

        log_head = ctk.CTkFrame(card_log, fg_color="transparent")
        log_head.pack(fill="x", padx=16, pady=(10, 4))

        ctk.CTkLabel(
            log_head,
            text="NHẬT KÝ XỬ LÝ (LIVE ACTIVITY)",
            font=ctk.CTkFont(family="Consolas", size=10, weight="bold"),
            text_color="#8E8E93"
        ).pack(side="left")

        self.btn_open_folder = ctk.CTkButton(
            log_head,
            text="📂 Mở Thư Mục Chứa File",
            font=ctk.CTkFont(family="Segoe UI", size=11),
            fg_color="#2C2C2E",
            text_color="#FFFFFF",
            hover_color="#3A3A3C",
            corner_radius=6,
            height=24,
            command=self.open_current_folder
        )
        self.btn_open_folder.pack(side="right")

        self.log_text = ctk.CTkTextbox(
            card_log,
            fg_color="transparent",
            text_color="#E5E5EA",
            font=ctk.CTkFont(family="Consolas", size=11),
            height=160
        )
        self.log_text.pack(fill="both", expand=True, padx=12, pady=(0, 10))

        self.log("Claim Helper sẵn sàng.")
        self.log("Kéo thả file vào giao diện hoặc bấm 'Bắt Đầu Đối Chiếu & Điền Giá' để chạy.")

    def get_ov_display_text(self):
        if self.overview_path:
            return f"✓ {os.path.basename(self.overview_path)}  ({self.overview_path})"
        return "Chưa chọn file (Kéo thả file Master Overview vào đây hoặc bấm Chọn File...)"

    def render_file_list(self):
        # Clear existing items
        for child in self.file_list_frame.winfo_children():
            child.destroy()

        if not self.claim_files:
            empty_lbl = ctk.CTkLabel(
                self.file_list_frame,
                text="📥  Kéo thả một hoặc nhiều file Claim Versus (.xlsx) vào đây\nhoặc bấm nút '+ Chọn File(s)...' ở góc trên",
                font=ctk.CTkFont(family="Segoe UI", size=12),
                text_color="#86868B",
                pady=24
            )
            empty_lbl.pack(fill="both", expand=True)
            self.claim_count_lbl.configure(text="📋  DANH SÁCH FILE CLAIM VERSUS (0)")
            return

        self.claim_count_lbl.configure(text=f"📋  DANH SÁCH FILE CLAIM VERSUS ({len(self.claim_files)})")

        for idx, fpath in enumerate(self.claim_files):
            fname = os.path.basename(fpath)
            is_casa = "casa" in fname.lower()
            tag_text = "Casa (CS)" if is_casa else "Nhan Hoang (NH)"
            tag_color = "#EBF3FE" if is_casa else "#EBF8F2"
            tag_text_color = "#0071E3" if is_casa else "#16A34A"

            row = ctk.CTkFrame(self.file_list_frame, fg_color="#FFFFFF", corner_radius=8)
            row.pack(fill="x", padx=10, pady=4)

            # File icon & name
            ctk.CTkLabel(
                row,
                text="📄",
                font=ctk.CTkFont(size=13)
            ).pack(side="left", padx=(10, 6), pady=6)

            ctk.CTkLabel(
                row,
                text=fname,
                font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
                text_color="#1D1D1F",
                anchor="w"
            ).pack(side="left", fill="x", expand=True, pady=6)

            # Badge
            badge = ctk.CTkLabel(
                row,
                text=f"  {tag_text}  ",
                font=ctk.CTkFont(family="Segoe UI", size=10, weight="bold"),
                fg_color=tag_color,
                text_color=tag_text_color,
                corner_radius=6
            )
            badge.pack(side="left", padx=8, pady=6)

            # Delete button
            del_btn = ctk.CTkButton(
                row,
                text="✕",
                font=ctk.CTkFont(size=11, weight="bold"),
                fg_color="transparent",
                text_color="#86868B",
                hover_color="#FFECEB",
                corner_radius=6,
                width=24,
                height=24,
                command=lambda p=fpath: self.remove_claim_file(p)
            )
            del_btn.pack(side="right", padx=(4, 8), pady=6)

    def remove_claim_file(self, fpath):
        if fpath in self.claim_files:
            self.claim_files.remove(fpath)
            self.render_file_list()

    def clear_all_claims(self):
        if not self.claim_files:
            return
        if messagebox.askyesno("Xác nhận", "Bạn có chắc muốn xóa tất cả file khỏi danh sách không?"):
            self.claim_files.clear()
            self.render_file_list()

    def setup_drag_and_drop(self):
        # Register targets
        self.drop_target_register(tkdnd.DND_FILES)
        self.dnd_bind('<<Drop>>', self.on_drop)
        self.dnd_bind('<<DropEnter>>', self.on_drop_enter)
        self.dnd_bind('<<DropLeave>>', self.on_drop_leave)

    def on_drop_enter(self, event):
        self.card_claims.configure(border_color="#0071E3", fg_color="#F0F8FF")

    def on_drop_leave(self, event):
        self.card_claims.configure(border_color="#E5E5EA", fg_color="#FFFFFF")

    def on_drop(self, event):
        self.on_drop_leave(event)
        paths = parse_drop_paths(event.data)
        if not paths:
            return

        added_claims = 0
        for p in paths:
            if os.path.isfile(p) and p.endswith(".xlsx"):
                fname = os.path.basename(p).lower()
                if "overview" in fname:
                    self.overview_path = p
                    self.ov_path_lbl.configure(text=self.get_ov_display_text(), text_color="#1D1D1F")
                    self.log(f"Đã nạp file Master Overview: {os.path.basename(p)}")
                else:
                    if p not in self.claim_files:
                        self.claim_files.append(p)
                        added_claims += 1
            elif os.path.isdir(p):
                # Dropped a directory: scan for versus files
                for f in os.listdir(p):
                    if "versus" in f.lower() and f.endswith(".xlsx") and not f.endswith("_backup.xlsx") and not f.startswith("~$"):
                        full_p = os.path.join(p, f)
                        if full_p not in self.claim_files:
                            self.claim_files.append(full_p)
                            added_claims += 1

        if added_claims > 0:
            self.render_file_list()
            self.log(f"Đã thêm {added_claims} file Claim Versus vào danh sách.")

    def browse_overview(self):
        f = filedialog.askopenfilename(
            title="Chọn file Master Overview",
            filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")]
        )
        if f:
            self.overview_path = os.path.abspath(f)
            self.ov_path_lbl.configure(text=self.get_ov_display_text(), text_color="#1D1D1F")
            self.log(f"Đã chọn Master Overview: {os.path.basename(f)}")

    def browse_claim_files(self):
        files = filedialog.askopenfilenames(
            title="Chọn một hoặc nhiều file Claim Versus",
            filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")]
        )
        if files:
            added = 0
            for f in files:
                p = os.path.abspath(f)
                if p not in self.claim_files:
                    self.claim_files.append(p)
                    added += 1
            if added > 0:
                self.render_file_list()
                self.log(f"Đã chọn thêm {added} file Claim Versus.")

    def browse_claim_folder(self):
        d = filedialog.askdirectory(title="Chọn thư mục chứa các file Claim Versus")
        if d:
            added = 0
            for f in os.listdir(d):
                if "versus" in f.lower() and f.endswith(".xlsx") and not f.endswith("_backup.xlsx") and not f.startswith("~$"):
                    p = os.path.join(d, f)
                    if p not in self.claim_files:
                        self.claim_files.append(p)
                        added += 1
            if added > 0:
                self.render_file_list()
                self.log(f"Đã quét và thêm {added} file từ thư mục {os.path.basename(d)}.")
            else:
                self.log("Không tìm thấy file Excel 'versus' nào mới trong thư mục này.")

    def open_current_folder(self):
        if self.claim_files:
            folder = os.path.dirname(self.claim_files[0])
            os.startfile(folder)
        elif self.overview_path:
            folder = os.path.dirname(self.overview_path)
            os.startfile(folder)
        else:
            os.startfile(os.path.abspath(os.path.dirname(__file__)))

    def log(self, text):
        t_str = datetime.datetime.now().strftime("%H:%M:%S")
        self.log_text.insert("end", f"[{t_str}] {text}\n")
        self.log_text.see("end")

    def start_processing(self):
        if not self.overview_path or not os.path.isfile(self.overview_path):
            messagebox.showerror("Thiếu thông tin", "Vui lòng chọn hoặc kéo thả file Master Overview (.xlsx)!")
            return
        if not self.claim_files:
            messagebox.showerror("Thiếu thông tin", "Danh sách file Claim đang trống. Vui lòng chọn hoặc kéo thả ít nhất 1 file Claim Versus!")
            return

        self.btn_run.configure(state="disabled", text="Đang xử lý dữ liệu...")
        self.prog_bar.set(0.1)
        self.log("="*60)
        self.log(f"BẮT ĐẦU XỬ LÝ {len(self.claim_files)} FILE CLAIM VERSUS")

        # Run background thread
        thread = threading.Thread(target=self.run_engine, daemon=True)
        thread.start()

    def run_engine(self):
        try:
            # 1. Load Master Overview
            self.log("Đang nạp dữ liệu Master Overview (vui lòng đợi vài giây)...")
            wb_o = openpyxl.load_workbook(self.overview_path, read_only=True, data_only=True)
            s_o = wb_o['Overview'] if 'Overview' in wb_o.sheetnames else wb_o.active
            headers_o = next(s_o.iter_rows(max_row=1, values_only=True))
            col_idx = {h: i for i, h in enumerate(headers_o)}

            cs_records = []
            nh_records = []

            for r in s_o.iter_rows(min_row=2, values_only=True):
                man = str(r[col_idx.get('Manufacturer', 0)] or '').strip().upper()
                sup = str(r[col_idx.get('Supplier', 0)] or '').strip().upper()
                cost = r[col_idx.get('Purchase cost', 0)]
                note_p = str(r[col_idx.get('Item PO note', 0)] or '').lower()
                note_s = str(r[col_idx.get('Note (sales)', 0)] or '').lower()

                # Filter out partial boxes
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
            self.log(f"✓ Đã nạp thành công: {len(cs_records)} bản ghi Casa (CS) | {len(nh_records)} bản ghi Nhan Hoang (NH).")

            # 2. Process Files
            total_files = len(self.claim_files)
            grand_claims = 0
            grand_matched = 0

            for i, fpath in enumerate(self.claim_files, start=1):
                fname = os.path.basename(fpath)
                is_casa = "casa" in fname.lower()
                ref_db = cs_records if is_casa else nh_records
                supp_name = "Casa (CS)" if is_casa else "Nhan Hoang (NH)"

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

                # Median exchange rate
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
                    desc = str(s.cell(row=r_idx, column=headers.index('Description')+1).value or '').strip() if 'Description' in headers else ''
                    desc2 = str(s.cell(row=r_idx, column=headers.index('Description_2')+1).value or '').strip() if 'Description_2' in headers else ''
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

                wb.save(fpath)
                wb.close()

                rate_str = f"{(file_matched / data_count * 100):.1f}%" if data_count > 0 else "100%"
                self.log(f"[{i}/{total_files}] ✓ {fname} ({supp_name}): Khớp {file_matched}/{data_count} dòng ({rate_str})")
                grand_claims += data_count
                grand_matched += file_matched

                self.prog_bar.set(0.1 + 0.9 * (i / total_files))

            self.log("="*60)
            self.log(f"🎉 HOÀN TẤT XỬ LÝ TOÀN BỘ {total_files} FILE!")
            self.log(f"Tổng cộng: {grand_matched}/{grand_claims} dòng khiếu nại đã được điền chi phí chuẩn xác.")
            self.after(0, self.finish_processing, True, f"Xử lý thành công toàn bộ {total_files} file ({grand_matched} lượt claim)!")

        except Exception as e:
            self.log(f"❌ Lỗi: {str(e)}")
            self.after(0, self.finish_processing, False, str(e))

    def finish_processing(self, success, msg):
        self.prog_bar.set(1.0 if success else 0)
        self.btn_run.configure(state="normal", text="Bắt Đầu Đối Chiếu & Điền Giá")
        if success:
            messagebox.showinfo("Thành công", msg)
        else:
            messagebox.showerror("Thông báo lỗi", msg)

if __name__ == "__main__":
    app = ClaimHelperAppleApp()
    app.mainloop()
