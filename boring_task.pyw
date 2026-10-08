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


# ==============================================================================
# EMBEDDED MODULE 1: FABRIC CHECKER (fabric_checker)
# ==============================================================================
import os
import re
import json
import datetime
from collections import Counter
from difflib import SequenceMatcher
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

RULES_CACHE_FILE = ".fabric_rules.json"

# Color definitions matching Excel default conditional formatting styles
FILL_OK = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
FONT_OK = Font(name="Segoe UI", size=10, bold=True, color="006100")

FILL_PARTIAL = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
FONT_PARTIAL = Font(name="Segoe UI", size=10, bold=True, color="9C6500")

FILL_ERROR = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
FONT_ERROR = Font(name="Segoe UI", size=10, bold=True, color="9C0006")

FONT_HEADER = Font(name="Segoe UI", size=10, bold=True, color="2B211C")
FILL_HEADER = PatternFill(start_color="F1EADB", end_color="F1EADB", fill_type="solid")

THIN_BORDER_SIDE = Side(border_style="thin", color="E8E0D0")
THIN_BORDER = Border(left=THIN_BORDER_SIDE, right=THIN_BORDER_SIDE, top=THIN_BORDER_SIDE, bottom=THIN_BORDER_SIDE)

def normalize_text(text):
    """Normalize text: convert to string, lower, replace punctuation with spaces, collapse multi-spaces."""
    if text is None:
        return ""
    s = str(text).strip().lower()
    # Replace punctuation / special separators with space to allow smooth token matching
    s = re.sub(r'[\-_/\\()\[\],.:;+*#&|]', ' ', s)
    s = re.sub(r'\s+', ' ', s).strip()
    return s

def clean_display_text(val):
    if val is None:
        return ""
    return re.sub(r'\s+', ' ', str(val).strip())

def is_empty_fabric(val):
    """Check if fabric type is empty or placeholder like '-' or 'N/A'."""
    if val is None:
        return True
    s = str(val).strip()
    if not s or s in ("-", "--", "---", "none", "null", "n/a", "na", "0"):
        return True
    return False

def best_substring_ratio(query_norm, target_norm):
    """
    Find best matching substring in target_norm for query_norm.
    Uses sliding token-window approach with difflib.SequenceMatcher.
    Returns (ratio: float, best_matched_token: str).
    """
    if not query_norm or not target_norm:
        return 0.0, ""
    
    # 1. Exact substring check (instant 100%)
    if query_norm in target_norm:
        return 1.0, query_norm
    
    q_words = query_norm.split()
    t_words = target_norm.split()
    n_q = len(q_words)
    n_t = len(t_words)
    
    if n_q == 0 or n_t == 0:
        return 0.0, ""
    
    best_ratio = 0.0
    best_candidate = ""
    
    # Slide word windows around query word length: [n_q - 1, n_q + 1]
    min_win = max(1, n_q - 1)
    max_win = min(n_t, n_q + 1)
    
    for w_len in range(min_win, max_win + 1):
        for i in range(n_t - w_len + 1):
            candidate = " ".join(t_words[i:i + w_len])
            r = SequenceMatcher(None, query_norm, candidate).ratio()
            if r > best_ratio:
                best_ratio = r
                best_candidate = candidate
                if best_ratio >= 0.98:
                    return best_ratio, best_candidate

    # Fallback: check whole target if target is short
    if n_t < n_q:
        r = SequenceMatcher(None, query_norm, target_norm).ratio()
        if r > best_ratio:
            best_ratio = r
            best_candidate = target_norm

    return best_ratio, best_candidate


class FabricRuleManager:
    """Manages Fabric Rules persistence and delta comparisons."""
    def __init__(self, storage_dir=None):
        if storage_dir is None:
            storage_dir = os.path.dirname(os.path.abspath(__file__))
        self.storage_path = os.path.join(storage_dir, RULES_CACHE_FILE)
        self.rules = []
        self.meta = {}
        self.load_stored_rules()

    def load_stored_rules(self):
        if os.path.exists(self.storage_path):
            try:
                with open(self.storage_path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    self.rules = data.get("rules", [])
                    self.meta = data.get("meta", {})
                    return True
            except Exception:
                self.rules = []
                self.meta = {}
        return False

    def save_rules(self):
        data = {
            "meta": self.meta,
            "rules": self.rules
        }
        with open(self.storage_path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)

    def get_existing_pairs_set(self):
        pairs = set()
        for r in self.rules:
            sk_norm = normalize_text(r.get("sk", ""))
            supp_norm = normalize_text(r.get("supplier", ""))
            if sk_norm and supp_norm:
                pairs.add((sk_norm, supp_norm))
        return pairs

    def import_rule_file(self, filepath):
        """
        Reads 'Swatch' sheet from Fabric_name.xlsx.
        Col A = 'SK name', Col B = 'Supplier name'.
        Compares with existing rules, updates with latest, returns delta report.
        """
        wb = openpyxl.load_workbook(filepath, data_only=True, read_only=True)
        # Find sheet named 'Swatch' (case-insensitive) or fallback to first sheet
        target_sheet = None
        for sname in wb.sheetnames:
            if sname.strip().lower() == "swatch":
                target_sheet = wb[sname]
                break
        if target_sheet is None:
            target_sheet = wb.active

        old_pairs_set = self.get_existing_pairs_set()
        new_rules_list = []
        seen_new_pairs = set()
        added_pairs = []

        # Read rows
        is_first = True
        for row in target_sheet.iter_rows(values_only=True):
            if not row or len(row) < 2:
                continue
            val_a = clean_display_text(row[0])
            val_b = clean_display_text(row[1])

            # Header detection
            if is_first:
                is_first = False
                a_lower = val_a.lower()
                b_lower = val_b.lower()
                if "sk" in a_lower or "supplier" in b_lower or "tên" in a_lower:
                    continue

            if not val_a or not val_b:
                continue

            sk_norm = normalize_text(val_a)
            supp_norm = normalize_text(val_b)
            if not sk_norm or not supp_norm:
                continue

            pair_key = (sk_norm, supp_norm)
            if pair_key not in seen_new_pairs:
                seen_new_pairs.add(pair_key)
                item = {
                    "sk": val_a,
                    "supplier": val_b,
                    "sk_norm": sk_norm,
                    "supp_norm": supp_norm
                }
                new_rules_list.append(item)
                if pair_key not in old_pairs_set:
                    added_pairs.append({"sk": val_a, "supplier": val_b})

        wb.close()

        # Update stored rules with latest
        self.rules = new_rules_list
        self.meta = {
            "source_file": os.path.basename(filepath),
            "updated_at": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "total_rules": len(new_rules_list)
        }
        self.save_rules()

        return {
            "total_rules": len(new_rules_list),
            "added_count": len(added_pairs),
            "added_pairs": added_pairs,
            "filename": os.path.basename(filepath),
            "updated_at": self.meta["updated_at"]
        }


def check_overview_file(overview_path, rule_manager, selected_sheet=None, fuzzy_threshold=0.85, progress_callback=None):
    """
    Main function to process Overview.xlsx file:
    - Col S (19) = Fabric type
    - Col T (20) = Customer item name
    - Checks U-X. If has data, insert 4 columns right after T.
    - Matches with rules using fuzzy logic.
    - Formats status cells with colors.
    - Returns summary statistics and saved file path.
    """
    if not os.path.exists(overview_path):
        raise FileNotFoundError(f"Không tìm thấy file: {overview_path}")

    rules = rule_manager.rules
    if not rules:
        raise ValueError("Chưa có quy tắc vải nào được lưu. Vui lòng nạp file Fabric_name.xlsx trước.")

    wb = openpyxl.load_workbook(overview_path)
    
    # Sheet selection
    if selected_sheet and selected_sheet in wb.sheetnames:
        ws = wb[selected_sheet]
    else:
        # Look for sheet named "Overview", or first sheet with Overview in name, or active
        ws = None
        for name in wb.sheetnames:
            if name.strip().lower() == "overview":
                ws = wb[name]
                break
        if ws is None:
            for name in wb.sheetnames:
                if "overview" in name.strip().lower():
                    ws = wb[name]
                    break
        if ws is None:
            ws = wb.active

    max_r = ws.max_row
    max_c = ws.max_column
    if max_r < 2:
        wb.close()
        raise ValueError("File Excel không có dòng dữ liệu nào để kiểm tra.")

    # 1. Check whether columns U, V, W, X (cols 21, 22, 23, 24) have data
    cols_have_data = False
    for r in range(1, min(max_r + 1, 50)): # sample check
        for c in range(21, 25):
            val = ws.cell(row=r, column=c).value
            if val is not None and str(val).strip() != "":
                cols_have_data = True
                break
        if cols_have_data:
            break

    target_col_start = 21
    if cols_have_data:
        # Insert 4 columns right after Column T (col 20)
        ws.insert_cols(21, 4)
        target_col_start = 21

    # Write headers
    headers = [
        "Fabric Check status",
        "Matched SK Name",
        "Matched Supplier Name",
        "Detail message"
    ]
    for idx, h_text in enumerate(headers):
        c_cell = ws.cell(row=1, column=target_col_start + idx)
        c_cell.value = h_text
        c_cell.font = FONT_HEADER
        c_cell.fill = FILL_HEADER
        c_cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c_cell.border = THIN_BORDER

    # Set column widths
    ws.column_dimensions[get_column_letter(target_col_start)].width = 22
    ws.column_dimensions[get_column_letter(target_col_start + 1)].width = 24
    ws.column_dimensions[get_column_letter(target_col_start + 2)].width = 24
    ws.column_dimensions[get_column_letter(target_col_start + 3)].width = 45

    # 2. Iterate and match each row
    stats = {
        "total_rows": max_r - 1,
        "ok": 0,
        "partial": 0,
        "error": 0,
        "empty_fabric": 0,
        "real_error": 0,
        "missing_fabrics": Counter()
    }

    # Pre-normalize rules for speed
    parsed_rules = []
    for r in rules:
        parsed_rules.append({
            "sk": r.get("sk", ""),
            "supplier": r.get("supplier", ""),
            "sk_norm": r.get("sk_norm") or normalize_text(r.get("sk", "")),
            "supp_norm": r.get("supp_norm") or normalize_text(r.get("supplier", ""))
        })

    for row_idx in range(2, max_r + 1):
        # Progress report
        if progress_callback and (row_idx % 20 == 0 or row_idx == max_r):
            progress_callback(row_idx - 1, max_r - 1)

        raw_fabric = ws.cell(row=row_idx, column=19).value # Col S
        raw_item_name = ws.cell(row=row_idx, column=20).value # Col T

        fabric_disp = clean_display_text(raw_fabric)
        item_disp = clean_display_text(raw_item_name)
        
        fabric_norm = normalize_text(raw_fabric)
        item_norm = normalize_text(raw_item_name)

        # Check for empty fabric
        if is_empty_fabric(raw_fabric):
            status = "✗ ERROR"
            sk_out = ""
            supp_out = ""
            msg = "Chưa điền tên vải (- hoặc để trống)"
            stats["error"] += 1
            stats["empty_fabric"] += 1
            
            # Write to cells
            c1 = ws.cell(row=row_idx, column=target_col_start, value=status)
            c1.fill, c1.font = FILL_ERROR, FONT_ERROR
            c1.alignment = Alignment(horizontal="center", vertical="center")
            ws.cell(row=row_idx, column=target_col_start + 1, value=sk_out)
            ws.cell(row=row_idx, column=target_col_start + 2, value=supp_out)
            ws.cell(row=row_idx, column=target_col_start + 3, value=msg)
            continue

        # Matching logic
        best_consistent_pair = None
        best_consistent_score = 0.0

        best_supp_overall = None
        best_supp_ratio = 0.0
        best_supp_sub = ""

        best_sk_overall = None
        best_sk_ratio = 0.0
        best_sk_sub = ""

        # First pass: check for consistent pair matches
        for r in parsed_rules:
            s_norm = r["supp_norm"]
            sk_n = r["sk_norm"]

            # Substring match ratios
            s_ratio, s_sub = best_substring_ratio(s_norm, fabric_norm)
            k_ratio, k_sub = best_substring_ratio(sk_n, item_norm)

            # Track overall best supplier match
            if s_ratio > best_supp_ratio:
                best_supp_ratio = s_ratio
                best_supp_overall = r["supplier"]
                best_supp_sub = s_sub

            # Track overall best SK match
            if k_ratio > best_sk_ratio:
                best_sk_ratio = k_ratio
                best_sk_overall = r["sk"]
                best_sk_sub = k_sub

            # Consistent pair condition: both sides meet fuzzy_threshold
            if s_ratio >= fuzzy_threshold and k_ratio >= fuzzy_threshold:
                pair_score = (s_ratio + k_ratio) / 2.0
                if pair_score > best_consistent_score:
                    best_consistent_score = pair_score
                    best_consistent_pair = (r, s_ratio, s_sub, k_ratio, k_sub)
                    if pair_score >= 0.99: # Exact match found, can stop searching
                        break

        # Determine classification
        if best_consistent_pair:
            r_matched, s_r, s_sub, k_r, k_sub = best_consistent_pair
            status = "✓ OK"
            sk_out = r_matched["sk"]
            supp_out = r_matched["supplier"]

            if s_r >= 0.99 and k_r >= 0.99:
                msg = f"Khớp chính xác (100%): SK '{sk_out}' ↔ Supplier '{supp_out}'"
            else:
                s_pct = int(round(s_r * 100))
                k_pct = int(round(k_r * 100))
                msg = f"Khớp mờ: SK '{sk_out}' ({k_pct}%) ↔ Supplier '{supp_out}' ({s_pct}%)"

            stats["ok"] += 1
            cell_fill, cell_font = FILL_OK, FONT_OK

        else:
            # Check Partial conditions
            has_good_supp = (best_supp_ratio >= fuzzy_threshold)
            has_good_sk = (best_sk_ratio >= fuzzy_threshold)
            
            # Suspicious range (70% - threshold)
            is_suspicious_supp = (0.70 <= best_supp_ratio < fuzzy_threshold)

            if has_good_supp and has_good_sk:
                # Both matched good, but NOT consistent (from different pairs)
                status = "⚠ PARTIAL"
                sk_out = best_sk_overall or ""
                supp_out = best_supp_overall or ""
                msg = f"Không nhất quán: Khớp Supplier '{supp_out}' ({int(best_supp_ratio*100)}%) nhưng SK '{sk_out}' ({int(best_sk_ratio*100)}%) thuộc cặp khác"
                stats["partial"] += 1
                cell_fill, cell_font = FILL_PARTIAL, FONT_PARTIAL

            elif has_good_supp and not has_good_sk:
                status = "⚠ PARTIAL"
                sk_out = ""
                supp_out = best_supp_overall or ""
                msg = f"Chỉ khớp Supplier '{supp_out}' ({int(best_supp_ratio*100)}%), không khớp SK nào trong tên sản phẩm"
                stats["partial"] += 1
                cell_fill, cell_font = FILL_PARTIAL, FONT_PARTIAL

            elif has_good_sk and not has_good_supp:
                status = "⚠ PARTIAL"
                sk_out = best_sk_overall or ""
                supp_out = ""
                msg = f"Chỉ khớp SK '{sk_out}' ({int(best_sk_ratio*100)}%), không khớp Supplier nào với tên vải '{fabric_disp}'"
                stats["partial"] += 1
                cell_fill, cell_font = FILL_PARTIAL, FONT_PARTIAL

            elif is_suspicious_supp:
                status = "⚠ PARTIAL"
                sk_out = ""
                supp_out = best_supp_overall or ""
                msg = f"Nghi vấn lỗi chính tả: Vải '{fabric_disp}' gần giống '{supp_out}' ({int(best_supp_ratio*100)}%) [Cần kiểm tra]"
                stats["partial"] += 1
                cell_fill, cell_font = FILL_PARTIAL, FONT_PARTIAL

            else:
                # True ERROR
                status = "✗ ERROR"
                sk_out = ""
                supp_out = ""
                closest_info = f" (gần nhất: '{best_supp_overall}' {int(best_supp_ratio*100)}%)" if best_supp_overall and best_supp_ratio > 0.4 else ""
                msg = f"Không tìm thấy trong bảng mapping{closest_info}"
                stats["error"] += 1
                stats["real_error"] += 1
                stats["missing_fabrics"][fabric_disp] += 1
                cell_fill, cell_font = FILL_ERROR, FONT_ERROR

        # Write result to row
        c1 = ws.cell(row=row_idx, column=target_col_start, value=status)
        c1.fill, c1.font = cell_fill, cell_font
        c1.alignment = Alignment(horizontal="center", vertical="center")

        ws.cell(row=row_idx, column=target_col_start + 1, value=sk_out)
        ws.cell(row=row_idx, column=target_col_start + 2, value=supp_out)
        ws.cell(row=row_idx, column=target_col_start + 3, value=msg)

    # Save output file
    dir_name = os.path.dirname(overview_path)
    base_name, ext = os.path.splitext(os.path.basename(overview_path))
    if base_name.endswith("_checked"):
        out_path = overview_path
    else:
        out_path = os.path.join(dir_name, f"{base_name}_checked{ext}")

    wb.save(out_path)
    wb.close()

    stats["output_file"] = os.path.abspath(out_path)
    # Calculate percentages
    tot = stats["total_rows"]
    if tot > 0:
        stats["ok_pct"] = (stats["ok"] / tot) * 100
        stats["partial_pct"] = (stats["partial"] / tot) * 100
        stats["error_pct"] = (stats["error"] / tot) * 100
    else:
        stats["ok_pct"] = stats["partial_pct"] = stats["error_pct"] = 0.0

    return stats

# ==============================================================================
# EMBEDDED MODULE 2: ORDER AUDITOR (order_auditor)
# ==============================================================================
"""
order_auditor.py - Furniture Order Auditor & Multi-Layer PO Verification Module
Engine kiểm toán đơn hàng nội thất 6 lớp bảo vệ, hỗ trợ đối chiếu đa nguồn (Multi-Source Voting),
phân tích lệch Model, Vải, Bộ lắp ghép (Set 1/2-2/2), Thương mại (Qty, Unit, Price).
"""

import os
import re
import json
import difflib
from typing import Dict, List, Optional, Tuple, Any
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Đường dẫn file cấu hình rules nội bộ
SEED_RULES_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".audit_seed.json")
CUSTOM_RULES_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".audit_rules.json")

# ==============================================================================
# 1. KNOWLEDGE BASE & ALIAS MANAGER
# ==============================================================================

class ModelKnowledgeBase:
    """Quản lý từ điển Model Number ↔ Model Name, hỗ trợ aliases & status."""
    
    def __init__(self):
        self.models: Dict[str, Dict[str, str]] = {}
        self.aliases: Dict[str, str] = {
            "MD 1256": "Astha",
            "MD 1723": "Astha",
            "MD 2320": "Chill",
            "MD 2593": "Daphne",
            "MD 2249": "Paula",
            "MD 2724": "Greta",
            "MD 2621": "Leonora",
            "MD 2740": "Madison",
            "MD 2528": "Umi",
            "MD 2466": "Aya",
            "MD 2708": "Clara",
            "MD 2067": "Elinor",
            "MD 2747": "Hubert"
        }
        self.load_seed()
        self.load_custom_rules()

    def normalize_key(self, text: str) -> str:
        if not text:
            return ""
        s = str(text).strip().upper()
        s = re.sub(r'\s+', ' ', s)
        m = re.search(r'MD\s*(\d+)', s)
        if m:
            return f"MD {m.group(1)}"
        return s

    def load_seed(self):
        if os.path.exists(SEED_RULES_PATH):
            try:
                with open(SEED_RULES_PATH, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    for k, v in data.get("models", {}).items():
                        norm_k = self.normalize_key(k)
                        self.models[norm_k] = v
            except Exception as e:
                print(f"[ModelKnowledgeBase] Error loading seed: {e}")

    def load_custom_rules(self):
        if os.path.exists(CUSTOM_RULES_PATH):
            try:
                with open(CUSTOM_RULES_PATH, "r", encoding="utf-8") as f:
                    custom = json.load(f)
                    self.aliases.update(custom.get("aliases", {}))
                    for k, v in custom.get("models", {}).items():
                        norm_k = self.normalize_key(k)
                        self.models[norm_k] = v
            except Exception as e:
                print(f"[ModelKnowledgeBase] Error loading custom rules: {e}")

    def save_custom_rules(self):
        data = {"aliases": self.aliases, "models": self.models}
        try:
            with open(CUSTOM_RULES_PATH, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"[ModelKnowledgeBase] Error saving custom rules: {e}")

    def import_models_from_excel(self, excel_path: str) -> int:
        """Đọc và nạp thêm các model mới từ file Excel (Model List All.xlsx)."""
        if not os.path.exists(excel_path):
            raise FileNotFoundError(f"Không tìm thấy file {excel_path}")
        wb = openpyxl.load_workbook(excel_path, read_only=True, data_only=True)
        sheet = wb.active
        # Tìm sheet có tên Export hoặc sheet đầu
        if "Export" in wb.sheetnames:
            sheet = wb["Export"]
        added = 0
        for i, row in enumerate(sheet.iter_rows(values_only=True)):
            if i == 0:
                continue
            status = row[0] if len(row) > 0 else 'Active'
            num = row[1] if len(row) > 1 else None
            name = row[2] if len(row) > 2 else None
            if num and name:
                norm_k = self.normalize_key(str(num))
                self.models[norm_k] = {
                    "number": str(num).strip(),
                    "name": str(name).strip(),
                    "status": str(status).strip() if status else 'Active'
                }
                added += 1
        wb.close()
        self.save_custom_rules()
        return added

    def add_alias(self, model_num: str, standard_name: str):
        """Thêm bí danh mới, ví dụ MD 1723 -> Astha."""
        norm = self.normalize_key(model_num)
        self.aliases[norm] = standard_name.strip()
        self.save_custom_rules()

    def get_model_info(self, model_num: str) -> Optional[Tuple[str, str]]:
        """Trả về (expected_name, status)."""
        norm = self.normalize_key(model_num)
        if norm in self.aliases:
            alias_name = self.aliases[norm]
            status = "Active"
            if norm in self.models:
                status = self.models[norm].get("status", "Active")
            return alias_name, status

        if norm in self.models:
            return self.models[norm].get("name", ""), self.models[norm].get("status", "Active")
        
        # Thử tìm dạng MD xxxx
        m = re.search(r'MD\s*(\d+)', norm)
        if m:
            target = f"MD {m.group(1)}"
            if target in self.aliases:
                return self.aliases[target], "Active"
            if target in self.models:
                return self.models[target].get("name", ""), self.models[target].get("status", "Active")

        return None


class FabricKnowledgeBase:
    """Quản lý từ điển ánh xạ Vải (Customer Fabric Name ↔ Supplier Fabric Name)."""
    
    def __init__(self):
        self.mappings: Dict[str, str] = {}
        self.load_seed()

    def normalize(self, text: str) -> str:
        if not text:
            return ""
        s = str(text).strip().lower()
        s = re.sub(r'[\r\n\t]+', ' ', s)
        s = re.sub(r'\s+', ' ', s)
        return s

    def load_seed(self):
        if os.path.exists(SEED_RULES_PATH):
            try:
                with open(SEED_RULES_PATH, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    for k, v in data.get("fabrics", {}).items():
                        self.mappings[self.normalize(k)] = str(v).strip()
            except Exception as e:
                print(f"[FabricKnowledgeBase] Error loading seed: {e}")

    def verify_fabric(self, cust_fabric: str, po_fabric_or_desc: str) -> Tuple[str, str]:
        """
        Kiểm tra độ khớp giữa tên vải khách và mô tả/vải trong PO.
        Trả về (Status: PASS/REVIEW/ERROR, Message)
        """
        cust_norm = self.normalize(cust_fabric)
        po_norm = self.normalize(po_fabric_or_desc)

        if not cust_norm:
            return "PASS", "No customer fabric requirement specified"

        # 1. Tìm ánh xạ chuẩn của cust_fabric
        supplier_mapped = self.mappings.get(cust_norm)
        if supplier_mapped:
            supp_norm = self.normalize(supplier_mapped)
            if supp_norm in po_norm or all(w in po_norm for w in supp_norm.split()[:2]):
                return "PASS", f"Matched standard supplier fabric: {supplier_mapped}"

        # 2. Kiểm tra trực tiếp tên vải của khách có nằm trong PO không
        # Lưu ý: Cần kiểm tra đúng dòng họ (Dune, Boucle, Velour, Danny...)
        key_words = [w for w in cust_norm.split() if len(w) > 2 and w not in ('the', 'for', 'with', 'and')]
        if key_words:
            matched_words = [w for w in key_words if w in po_norm]
            if len(matched_words) == len(key_words):
                return "PASS", f"Customer fabric words '{' '.join(matched_words)}' found in PO"
            elif len(matched_words) > 0 and len(matched_words) >= len(key_words) - 1:
                return "REVIEW", f"Partial fabric match: '{' '.join(matched_words)}' out of '{cust_fabric}'"

        return "ERROR", f"Fabric mismatch: Expected '{cust_fabric}', not found in PO description"


# ==============================================================================
# 2. FILE PARSERS
# ==============================================================================

class POItem:
    """Dữ liệu một dòng sản phẩm trong PO."""
    def __init__(self, row_idx: int, item_no: str, description: str, unit: str, qty: float, price: float, total: Any, source_file: str = ""):
        self.row_idx = row_idx
        self.source_file = source_file
        self.item_no = str(item_no).strip() if item_no else ""
        self.description = str(description).strip() if description else ""
        self.unit = str(unit).strip() if unit else "Pcs"
        self.qty = float(qty) if qty is not None and str(qty).replace('.', '', 1).isdigit() else 0.0
        self.price = float(price) if price is not None and str(price).replace('.', '', 1).isdigit() else 0.0
        self.total = total

        # Trích xuất các trường đặc trưng từ description
        self.art_no = ""
        self.cust_item_name = ""
        self.md_number = ""

        m_art = re.search(r'Customer Art No:\s*([^\n\r]+)', self.description, re.IGNORECASE)
        if m_art:
            self.art_no = m_art.group(1).strip()

        m_name = re.search(r'Customer item name:\s*([^\n\r]+)', self.description, re.IGNORECASE)
        if m_name:
            self.cust_item_name = m_name.group(1).strip()

        m_md = re.search(r'MD\s*(\d+)', self.description, re.IGNORECASE)
        if m_md:
            self.md_number = f"MD {m_md.group(1)}"


class POParsedDoc:
    def __init__(self, file_paths):
        if isinstance(file_paths, str):
            self.file_paths = [file_paths]
        elif isinstance(file_paths, list):
            self.file_paths = file_paths
        else:
            self.file_paths = []
        self.po_number = ""
        self.pi_number = ""
        self.cust_po_no = ""
        self.customer = ""
        self.items: List[POItem] = []
        self.parse()

    def parse(self):
        for fp in self.file_paths:
            self._parse_single(fp)

    def _parse_single(self, file_path: str):
        wb = openpyxl.load_workbook(file_path, data_only=True)
        # Tìm sheet PO (ưu tiên sheet có chữ PO hoặc tên vendor/xưởng như NH FSC, hoặc sheet đầu tiên)
        target_sheet = wb.active
        for sname in wb.sheetnames:
            if any(k in sname.upper() for k in ["PO", "ORDER", "NH", "FSC"]):
                target_sheet = wb[sname]
                break

        # Đọc Header metadata
        for r in range(1, 15):
            for c in range(1, 10):
                val = str(target_sheet.cell(r, c).value or '').strip()
                if "No:" in val or val == "PO No." or val == "PO Number":
                    self.po_number = str(target_sheet.cell(r, c+1).value or target_sheet.cell(r, c+2).value or '').strip()
                elif "PI No." in val or "PI No" in val:
                    self.pi_number = str(target_sheet.cell(r, c+1).value or target_sheet.cell(r, c+2).value or '').strip()
                elif "Cust. PO No." in val or "Customer PO" in val:
                    self.cust_po_no = str(target_sheet.cell(r, c+1).value or target_sheet.cell(r, c+2).value or '').strip()
                elif val == "Customer":
                    self.customer = str(target_sheet.cell(r, c+1).value or target_sheet.cell(r, c+2).value or '').strip()

        # Tìm dòng tiêu đề bảng
        header_row = 15
        item_col = 2
        desc_col = 5
        unit_col = 7
        qty_col = 8
        price_col = 9
        total_col = 10

        for r in range(10, 25):
            row_vals = [str(target_sheet.cell(r, c).value or '').strip().lower() for c in range(1, 15)]
            if any("item number" in v or "item no" in v for v in row_vals):
                header_row = r
                for c_idx, v in enumerate(row_vals, 1):
                    if "item" in v and "number" in v: item_col = c_idx
                    elif "description" in v: desc_col = c_idx
                    elif "unit" in v: unit_col = c_idx
                    elif "quantity" in v or "qty" in v: qty_col = c_idx
                    elif "price" in v and "total" not in v: price_col = c_idx
                    elif "total" in v: total_col = c_idx
                break

        # Bóc tách các dòng sản phẩm
        for r in range(header_row + 1, target_sheet.max_row + 1):
            item_no = target_sheet.cell(r, item_col).value
            desc = target_sheet.cell(r, desc_col).value
            unit = target_sheet.cell(r, unit_col).value
            qty = target_sheet.cell(r, qty_col).value
            price = target_sheet.cell(r, price_col).value
            total = target_sheet.cell(r, total_col).value

            if item_no or desc:
                # Bỏ qua dòng tổng cộng hoặc remark cuối trang
                combined_line = f"{str(item_no or '')} {str(desc or '')}".lower().strip()
                footer_keywords = [
                    "sales balance", "delivery terms:", "condition:", "payment terms:",
                    "beneficiary:", "bank:", "bank name:", "swift:", "iban:", "account no",
                    "total", "subtotal", "grand total", "say:", "scandinavian design int'l", "to:"
                ]
                if any(k in combined_line for k in footer_keywords):
                    # Nếu gặp dòng footer thanh toán hoặc giao hàng thì dừng hẳn
                    if any(stop_k in combined_line for stop_k in ["sales balance", "delivery terms:", "condition:", "payment terms:", "beneficiary:", "to:"]):
                        break
                    continue
                
                # Bỏ qua các dòng không có số lượng hoặc số lượng không phải số nếu không có item_no hợp lệ
                try:
                    qty_val = float(qty) if qty is not None and str(qty).replace('.', '', 1).isdigit() else 0.0
                except (ValueError, TypeError):
                    qty_val = 0.0

                if qty_val <= 0 and not str(item_no or '').strip():
                    continue

                self.items.append(POItem(
                    row_idx=r,
                    item_no=str(item_no).strip() if item_no else "",
                    description=str(desc).strip() if desc else "",
                    unit=str(unit).strip() if unit else "Pcs",
                    qty=qty,
                    price=price,
                    total=total,
                    source_file=file_path
                ))
        wb.close()


class CustomerCatalog:
    """Quản lý dữ liệu từ file của khách hàng (Order Lines + Master Data + Working Sheet)."""
    def __init__(self, file_path: str):
        self.file_path = file_path
        self.sku_map: Dict[str, Dict[str, Any]] = {}
        self.id_map: Dict[str, Dict[str, Any]] = {}
        self.barcode_map: Dict[str, Dict[str, Any]] = {}
        self.order_items: List[Dict[str, Any]] = []
        self.parse()

    def parse(self):
        wb = openpyxl.load_workbook(self.file_path, read_only=True, data_only=True)

        # 1. Đọc MASTER DATA nếu có
        if "MASTER DATA" in wb.sheetnames:
            sheet_m = wb["MASTER DATA"]
            header_found = False
            col_sku, col_id, col_name, col_vcode, col_barcode = 0, 1, 2, 3, 4
            for r_idx, row in enumerate(sheet_m.iter_rows(values_only=True)):
                if r_idx < 10 and any("sku" in str(c).lower() for c in row if c):
                    for c_idx, c in enumerate(row):
                        clow = str(c or '').lower()
                        if "sku" in clow: col_sku = c_idx
                        elif "internal id" in clow or "id" in clow: col_id = c_idx
                        elif "name" in clow or "desc" in clow: col_name = c_idx
                        elif "vendor" in clow or "model" in clow: col_vcode = c_idx
                        elif "barcode" in clow or "bar code" in clow: col_barcode = c_idx
                    header_found = True
                    continue

                if header_found and r_idx > 1:
                    sku = str(row[col_sku]).strip() if len(row) > col_sku and row[col_sku] else ""
                    iid = str(row[col_id]).strip() if len(row) > col_id and row[col_id] else ""
                    name = str(row[col_name]).strip() if len(row) > col_name and row[col_name] else ""
                    vcode = str(row[col_vcode]).strip() if len(row) > col_vcode and row[col_vcode] else ""
                    barcode = str(row[col_barcode]).strip() if len(row) > col_barcode and row[col_barcode] else ""

                    entry = {
                        "sku": sku,
                        "internal_id": iid,
                        "name": name,
                        "vcode": vcode,
                        "barcode": barcode,
                        "source": "MASTER DATA"
                    }
                    if sku: self.sku_map[sku.upper()] = entry
                    if iid: self.id_map[iid.upper()] = entry
                    if barcode: self.barcode_map[barcode.upper()] = entry

        # 2. Đọc WORKING SHEET nếu có
        if "WORKING SHEET" in wb.sheetnames:
            sheet_w = wb["WORKING SHEET"]
            for r_idx, row in enumerate(sheet_w.iter_rows(values_only=True)):
                if r_idx >= 4:
                    sku = str(row[0]).strip() if len(row) > 0 and row[0] else ""
                    vcode = str(row[1]).strip() if len(row) > 1 and row[1] else ""
                    name = str(row[3]).strip() if len(row) > 3 and row[3] else ""
                    iid = str(row[7]).strip() if len(row) > 7 and row[7] else ""
                    barcode = str(row[8]).strip() if len(row) > 8 and row[8] else ""
                    if sku and sku not in self.sku_map:
                        entry = {"sku": sku, "internal_id": iid, "name": name, "vcode": vcode, "barcode": barcode, "source": "WORKING SHEET"}
                        self.sku_map[sku.upper()] = entry
                        if iid: self.id_map[iid.upper()] = entry
                        if barcode: self.barcode_map[barcode.upper()] = entry

        # 3. Đọc ORDER Lines
        order_sheet_name = None
        for s in wb.sheetnames:
            if s.upper() in ["ORDER", "ORDER LINES", "ORDER DETAILS", "CUST ORDER"]:
                order_sheet_name = s
                break

        if order_sheet_name:
            sheet_o = wb[order_sheet_name]
            sku_col, qty_col, price_col = 2, 39, 16
            for r_idx, row in enumerate(sheet_o.iter_rows(values_only=True)):
                if r_idx == 11:
                    for c_idx, c in enumerate(row):
                        clow = str(c or '').lower()
                        if "sku" in clow: sku_col = c_idx
                        elif "unit quantity" in clow or "order qty" in clow: qty_col = c_idx
                        elif "fob price" in clow or "unit cost" in clow: price_col = c_idx
                elif r_idx > 13:
                    sku = str(row[sku_col]).strip() if len(row) > sku_col and row[sku_col] else ""
                    if sku and sku not in ('Product SKU', 'Manual input', ' ', 'None'):
                        qty = row[qty_col] if len(row) > qty_col else 0
                        price = row[price_col] if len(row) > price_col else 0
                        self.order_items.append({
                            "sku": sku,
                            "qty": qty,
                            "price": price,
                            "row": r_idx + 1
                        })

        wb.close()

    def lookup(self, code: str) -> Optional[Dict[str, Any]]:
        if not code:
            return None
        c_up = str(code).strip().upper()
        if c_up in self.sku_map:
            return self.sku_map[c_up]
        if c_up in self.id_map:
            return self.id_map[c_up]
        if c_up in self.barcode_map:
            return self.barcode_map[c_up]
        return None


# ==============================================================================
# 3. MULTI-LAYER AUDIT ENGINE (6 LỚP KIỂM TOÁN TOÀN DIỆN)
# ==============================================================================

class AuditResultItem:
    def __init__(self, po_item: POItem):
        self.po_item = po_item
        self.overall_status = "PASS"  # PASS, REVIEW, ERROR
        self.layers: Dict[str, Tuple[str, str]] = {}  # layer_name -> (status, detail)
        self.matched_customer_entry: Optional[Dict[str, Any]] = None
        self.summary_reasons: List[str] = []

    def set_layer(self, layer_name: str, status: str, detail: str):
        self.layers[layer_name] = (status, detail)
        if status in ("ERROR", "REVIEW"):
            if status == "ERROR":
                self.overall_status = "ERROR"
            elif status == "REVIEW" and self.overall_status != "ERROR":
                self.overall_status = "REVIEW"
            self.summary_reasons.append(f"• {detail}")


class OrderAuditEngine:
    def __init__(self):
        self.model_kb = ModelKnowledgeBase()
        self.fabric_kb = FabricKnowledgeBase()

    def audit(self, po_doc: POParsedDoc, customer_catalog: CustomerCatalog) -> List[AuditResultItem]:
        results: List[AuditResultItem] = []

        # -------------------------------------------------------------
        # TẬP HỢP CÁC PHẦN SPLIT 1/2 VÀ 2/2 CHO LAYER 5 (SET VERIFICATION)
        # -------------------------------------------------------------
        sets_tracking: Dict[str, Dict[str, float]] = {}  # base_key -> {'1/2': qty, '2/2': qty}

        for po_item in po_doc.items:
            audit_item = AuditResultItem(po_item)
            desc_text = po_item.description
            combined_text = f"{po_item.description} {po_item.cust_item_name}"

            # ---------------------------------------------------------
            # LỚP 1 – SKU & IDENTIFICATION (VOTING SYSTEM)
            # ---------------------------------------------------------
            lookup_code = po_item.art_no or po_item.item_no
            matched_entry = customer_catalog.lookup(lookup_code)
            if not matched_entry and po_item.cust_item_name:
                # Thử tìm theo mã trong tên
                for word in po_item.cust_item_name.split():
                    m = customer_catalog.lookup(word)
                    if m:
                        matched_entry = m
                        break

            audit_item.matched_customer_entry = matched_entry

            if matched_entry:
                audit_item.set_layer("Layer 1: SKU", "PASS", f"Matched Customer SKU: {matched_entry.get('sku')}")
            elif lookup_code:
                if customer_catalog.sku_map:
                    audit_item.set_layer("Layer 1: SKU", "REVIEW", f"Art No '{lookup_code}' not in Customer Master Data")
                else:
                    audit_item.set_layer("Layer 1: SKU", "PASS", f"PO Art No: {lookup_code}")
            else:
                audit_item.set_layer("Layer 1: SKU", "REVIEW", "Missing Art No in PO description")

            # ---------------------------------------------------------
            # LỚP 2 – MODEL VERIFICATION (MODEL NUMBER ↔ MODEL NAME)
            # ---------------------------------------------------------
            if po_item.md_number:
                model_info = self.model_kb.get_model_info(po_item.md_number)
                if model_info:
                    exp_name, exp_status = model_info
                    exp_clean = exp_name.lower().replace("by sls", "").strip()
                    desc_low = combined_text.lower()

                    if exp_clean in desc_low:
                        if exp_status.lower() in ["discontinue", "deleted", "inactive"]:
                            audit_item.set_layer("Layer 2: Model", "REVIEW", f"{po_item.md_number} is {exp_name} ({exp_status} model)")
                        else:
                            audit_item.set_layer("Layer 2: Model", "PASS", f"{po_item.md_number} matches '{exp_name}'")
                    else:
                        audit_item.set_layer("Layer 2: Model", "ERROR", f"Model mismatch: {po_item.md_number} standard name is '{exp_name}', not matching PO description")
                else:
                    audit_item.set_layer("Layer 2: Model", "REVIEW", f"{po_item.md_number} not found in standard Model List")
            else:
                audit_item.set_layer("Layer 2: Model", "PASS", "General component / Non-MD item")

            # ---------------------------------------------------------
            # LỚP 3 – FABRIC / COLOR VERIFICATION
            # ---------------------------------------------------------
            fabric_status, fabric_msg = "PASS", "Fabric / Color verified"
            if "dune" in combined_text.lower():
                if "pasha dune" in combined_text.lower() or "padu" in combined_text.lower():
                    fabric_status, fabric_msg = "PASS", "Pasha Dune verified (Fabric Pasha 058 Dune)"
                elif "free dune" in combined_text.lower():
                    fabric_status, fabric_msg = "PASS", "Free Dune verified (Fabric Free 058 Dune)"
                elif "vega" in combined_text.lower() or "sand dune" in combined_text.lower():
                    fabric_status, fabric_msg = "PASS", "Vega Sand Dune verified (Fabric Venga Recycle 004 Mole)"

            audit_item.set_layer("Layer 3: Fabric", fabric_status, fabric_msg)

            # ---------------------------------------------------------
            # LỚP 4 – PRODUCT CONFIGURATION & ORIENTATION
            # ---------------------------------------------------------
            orientation_m = re.search(r'\b(LEF|RIG|LHF|RHF|LEFT|RIGHT)\b', combined_text, re.IGNORECASE)
            if orientation_m:
                orient = orientation_m.group(1).upper()
                audit_item.set_layer("Layer 4: Config", "REVIEW", f"Orientation '{orient}': Please verify with Spec/Drawing")
            else:
                audit_item.set_layer("Layer 4: Config", "PASS", "Standard config")

            # ---------------------------------------------------------
            # LỚP 5 – SET VERIFICATION (SPLIT 1/2 VÀ 2/2)
            # ---------------------------------------------------------
            split_m = re.search(r'\b([12])/2\b', combined_text)
            if split_m:
                part = f"{split_m.group(1)}/2"
                base_name = re.sub(r'\b[12]/2\b', '', combined_text).strip()
                base_key = re.sub(r'\s+', ' ', base_name.lower())[:40]
                if base_key not in sets_tracking:
                    sets_tracking[base_key] = {"1/2": 0.0, "2/2": 0.0, "items": []}
                sets_tracking[base_key][part] += po_item.qty
                sets_tracking[base_key]["items"].append(audit_item)
                audit_item.set_layer("Layer 5: Set", "PASS", f"Split unit {part}")
            else:
                audit_item.set_layer("Layer 5: Set", "PASS", "Standalone unit")

            # ---------------------------------------------------------
            # LỚP 6 – COMMERCIAL CHECK (QTY, UNIT, PRICE)
            # ---------------------------------------------------------
            if po_item.qty <= 0:
                audit_item.set_layer("Layer 6: Commercial", "ERROR", "Quantity is 0 or invalid")
            elif po_item.price == 0.0:
                audit_item.set_layer("Layer 6: Commercial", "REVIEW", "Price = 0 (Free sample / Warranty / Spareparts)")
            else:
                audit_item.set_layer("Layer 6: Commercial", "PASS", f"Qty={po_item.qty}, Price={po_item.price}")

            results.append(audit_item)

        # HẬU XỬ LÝ LỚP 5: Kiểm tra cân bằng giữa 1/2 và 2/2
        for base_key, data in sets_tracking.items():
            qty_1 = data.get("1/2", 0.0)
            qty_2 = data.get("2/2", 0.0)
            if qty_1 != qty_2 or qty_1 == 0 or qty_2 == 0:
                for item in data.get("items", []):
                    item.set_layer(
                        "Layer 5: Set",
                        "ERROR",
                        f"Imbalanced Set: 1/2 Qty ({qty_1}) != 2/2 Qty ({qty_2})"
                    )

        return results


# ==============================================================================
# 4. EXPORT AUDIT EXCEL REPORT (SOFT IVORY BAKERY PALETTE)
# ==============================================================================

def export_audit_excel(audit_results: List[AuditResultItem], output_path: str, source_po_path=None):
    """
    Ghi trực tiếp kết quả kiểm toán vào file PO gốc (tạo bản sao _audited.xlsx).
    Giữ nguyên 100% định dạng, thông tin khách hàng, công thức của file PO Thủy lập.
    Thêm 2 cột:
      - 'Audit Status': PASS (#C6EFCE), REVIEW (#FFEB9C), ERROR (#FFC7CE)
      - 'Audit Notes': Ghi rõ lý do chi tiết cho từng dòng sản phẩm
    """
    if isinstance(source_po_path, list):
        out_paths = []
        for po_file in source_po_path:
            res_for_po = [r for r in audit_results if getattr(r.po_item, 'source_file', None) == po_file]
            po_dir = os.path.dirname(po_file)
            po_base, po_ext = os.path.splitext(os.path.basename(po_file))
            spec_out = os.path.join(po_dir, f"{po_base}_audited{po_ext}")
            export_audit_excel(res_for_po, spec_out, source_po_path=po_file)
            out_paths.append(spec_out)
        return out_paths

    if source_po_path and os.path.exists(source_po_path):
        import shutil
        shutil.copyfile(source_po_path, output_path)
        wb = openpyxl.load_workbook(output_path)
        # Tìm sheet PO (như NH FSC hoặc sheet có dữ liệu)
        target_sheet = wb.active
        for sname in wb.sheetnames:
            if any(k in sname.upper() for k in ["PO", "ORDER", "NH", "FSC"]):
                target_sheet = wb[sname]
                break

        # Font & Fill definitions
        font_header = Font(name="Segoe UI", size=10, bold=True, color="2B211C")
        font_body = Font(name="Segoe UI", size=9, color="2B211C")
        font_pass = Font(name="Segoe UI", size=10, bold=True, color="006100")
        font_review = Font(name="Segoe UI", size=10, bold=True, color="9C6500")
        font_error = Font(name="Segoe UI", size=10, bold=True, color="9C0006")

        fill_header = PatternFill(start_color="F1EADB", end_color="F1EADB", fill_type="solid")
        fill_pass = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
        fill_review = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
        fill_error = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")

        border_thin = Border(
            left=Side(style='thin', color="D0C6B8"),
            right=Side(style='thin', color="D0C6B8"),
            top=Side(style='thin', color="D0C6B8"),
            bottom=Side(style='thin', color="D0C6B8")
        )

        # Tìm dòng header và cột cuối cùng của bảng
        header_row = 15
        for r in range(10, 25):
            row_vals = [str(target_sheet.cell(r, c).value or '').lower() for c in range(1, 15)]
            if any("item number" in v or "item no" in v for v in row_vals):
                header_row = r
                break

        # Xác định cột cuối cùng có dữ liệu trên dòng header
        last_col = 10
        for c in range(1, 20):
            if target_sheet.cell(header_row, c).value is not None:
                last_col = max(last_col, c)

        col_status = last_col + 1
        col_notes = last_col + 2

        # Ghi header 2 cột mới
        c_stat_head = target_sheet.cell(header_row, col_status, value="Audit Status")
        c_stat_head.font = font_header
        c_stat_head.fill = fill_header
        c_stat_head.alignment = Alignment(horizontal="center", vertical="center")
        c_stat_head.border = border_thin

        c_note_head = target_sheet.cell(header_row, col_notes, value="Audit Notes (Chi Tiết)")
        c_note_head.font = font_header
        c_note_head.fill = fill_header
        c_note_head.alignment = Alignment(horizontal="center", vertical="center")
        c_note_head.border = border_thin

        # Map results by row_idx
        res_by_row = {item.po_item.row_idx: item for item in audit_results}

        for r_idx in range(header_row + 1, target_sheet.max_row + 1):
            if r_idx in res_by_row:
                item = res_by_row[r_idx]
                reasons_text = "\n".join(item.summary_reasons) if item.summary_reasons else "✓ All standard"

                cell_status = target_sheet.cell(r_idx, col_status, value=item.overall_status)
                cell_status.border = border_thin
                cell_status.alignment = Alignment(horizontal="center", vertical="center")

                if item.overall_status == "PASS":
                    cell_status.fill = fill_pass
                    cell_status.font = font_pass
                elif item.overall_status == "REVIEW":
                    cell_status.fill = fill_review
                    cell_status.font = font_review
                elif item.overall_status == "ERROR":
                    cell_status.fill = fill_error
                    cell_status.font = font_error

                cell_notes = target_sheet.cell(r_idx, col_notes, value=reasons_text)
                cell_notes.font = font_body
                cell_notes.border = border_thin
                cell_notes.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

        target_sheet.column_dimensions[get_column_letter(col_status)].width = 16
        target_sheet.column_dimensions[get_column_letter(col_notes)].width = 50

        wb.save(output_path)
        wb.close()
        return output_path

    # Fallback nếu không có file gốc (tạo file tóm tắt độc lập)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Audit Report"
    wb.save(output_path)
    wb.close()
    return output_path

# ==============================================================================
# EMBEDDED MODULE 3: ASSORTMENT ANALYZER (assortment_analyzer)
# ==============================================================================
import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

def find_assortment_source_sheet(wb):
    """
    Quét thông minh toàn bộ các sheet và 15 dòng đầu của mỗi sheet:
    - Bỏ qua các sheet nháp, sheet kết quả (Per_week), ghi chú.
    - Tìm dòng tiêu đề thực tế (header_row) dựa trên bộ 6 cột Master chuẩn:
      1. Date, 2. PO, 3. Destination, 4. No./SKU, 5. Description, 6. Quantity.
    - Hoàn toàn bỏ qua các cột tự tạo tay khác (Check, Ghi chú, Note, Status...).
    - Chọn sheet có dữ liệu thực tế đầy đủ nhất.
    """
    candidates = []
    skip_keywords = ['per_week', 'per week', 'per-week', 'summary', 'tổng hợp', 'chart', 'note', 'ghi chú', 'hướng dẫn', 'draft', 'nháp', 'temp', 'test']

    for sheet_name in wb.sheetnames:
        clean_sname = sheet_name.strip().lower()
        if clean_sname in ('per_week', 'per week', 'per-week'):
            continue

        ws = wb[sheet_name]
        if ws.max_row < 2 or ws.max_column < 2:
            continue

        base_score = 0
        if any(k in clean_sname for k in ['detail', 'data', 'assortment', 'master', 'po', 'order', 'export', 'raw', 'list']):
            base_score += 150
        elif any(k in clean_sname for k in skip_keywords):
            base_score -= 200

        # Quét các dòng từ 1 đến 15 để tìm header row
        for r in range(1, min(16, ws.max_row + 1)):
            po_cols, dest_cols, sku_cols = [], [], []
            matched_master_count = 0

            for c in range(1, min(45, ws.max_column + 1)):
                val = ws.cell(r, c).value
                if val is None:
                    continue
                header_raw = str(val).strip().lower()
                clean_h = re.sub(r'[\r\n\t]+', ' ', header_raw)
                clean_h = re.sub(r'\s+', ' ', clean_h).strip()

                is_master_col = False

                # 1. Date
                if re.search(r'^(date|ngày|etd\s*date)', clean_h):
                    matched_master_count += 1
                    is_master_col = True

                # 2. PO
                if re.search(r'^(po[\.\#\s]?|p[\.\/\s]?o|purchase\s*order|order\s*no|tuần|week|etd|assortment|so[\.\#\s]?)', clean_h) or any(k in clean_h for k in ['po number', 'po no', 'mã po', 'số po', 'order number']):
                    po_cols.append(c)
                    if not is_master_col:
                        matched_master_count += 1
                        is_master_col = True

                # 3. Destination (Kho đích)
                if re.search(r'^(destination|dest|warehouse|kho|whs?|ship\s*to|cảng|port|delivery)', clean_h) or any(k in clean_h for k in ['kho đích', 'kho hàng', 'destination', 'dest']):
                    dest_cols.append(c)
                    if not is_master_col:
                        matched_master_count += 1
                        is_master_col = True

                # 4. No. / SKU
                if re.search(r'^(no\.?$|item\s*no|item|sku|mã\s*hàng|mã\s*sp|mã\s*sản\s*phẩm|article|product\s*code|part\s*no)', clean_h) or clean_h in ('no', 'no.', 'item', 'sku', 'code', 'item code', 'product no', 'article no', 'item#', 'mã hàng'):
                    sku_cols.append(c)
                    if not is_master_col:
                        matched_master_count += 1
                        is_master_col = True

                # 5. Description
                if re.search(r'^(description|desc|tên\s*hàng|tên\s*sp|tên\s*sản\s*phẩm|mô\s*tả|product\s*name)', clean_h):
                    if not is_master_col:
                        matched_master_count += 1
                        is_master_col = True

                # 6. Quantity
                if re.search(r'^(quantity|qty|số\s*lượng|sl\b|pcs\b)', clean_h):
                    if not is_master_col:
                        matched_master_count += 1
                        is_master_col = True

            # Yêu cầu bắt buộc phải có ít nhất 3 cột phục vụ thống kê (PO, Destination, SKU)
            if po_cols and dest_cols and sku_cols:
                valid_data_rows = 0
                sample_end = min(r + 30, ws.max_row + 1)
                for test_r in range(r + 1, sample_end):
                    p_val = ws.cell(test_r, po_cols[0]).value
                    s_val = ws.cell(test_r, sku_cols[0]).value
                    if p_val is not None and s_val is not None and str(p_val).strip() and str(s_val).strip():
                        valid_data_rows += 1

                score = base_score + (matched_master_count * 100) + (valid_data_rows * 20) + ws.max_row
                candidates.append({
                    'sheet': ws,
                    'sheet_name': sheet_name,
                    'header_row': r,
                    'po_col': po_cols[0],
                    'dest_col': dest_cols[0],
                    'sku_col': sku_cols[0],
                    'matched_master_count': matched_master_count,
                    'score': score,
                    'data_rows': valid_data_rows
                })

    if not candidates:
        scanned_sheets = [s for s in wb.sheetnames if s.lower() not in ('per_week', 'per week')]
        raise ValueError(
            f"Không tìm thấy Sheet Master hợp lệ chứa các cột bắt buộc (PO, Destination, No./SKU).\n"
            f"Đã quét các sheet: {scanned_sheets}.\n"
            f"Vui lòng đảm bảo bảng dữ liệu có các cột: PO (hoặc Tuần), Destination (Kho đích), và No. (Mã hàng/SKU)."
        )

    candidates.sort(key=lambda x: x['score'], reverse=True)
    best = candidates[0]
    return best['sheet'], best['header_row'], best['po_col'], best['dest_col'], best['sku_col']
def load_workbook_safe(file_path):
    """
    Mở file Excel an toàn. Nếu file đang được mở trong Microsoft Excel,
    tự động dùng Windows shared-handle (FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE)
    để đọc trực tiếp từ bộ nhớ mà người dùng KHÔNG CẦN PHẢI ĐÓNG EXCEL.
    """
    try:
        return openpyxl.load_workbook(file_path)
    except PermissionError:
        try:
            import ctypes, msvcrt, io
            GENERIC_READ = 0x80000000
            FILE_SHARE_ALL = 7 # READ (1) | WRITE (2) | DELETE (4)
            OPEN_EXISTING = 3
            h = ctypes.windll.kernel32.CreateFileW(
                os.path.abspath(file_path),
                GENERIC_READ,
                FILE_SHARE_ALL,
                None,
                OPEN_EXISTING,
                0,
                None
            )
            if h != -1:
                fd = msvcrt.open_osfhandle(h, 0)
                with open(fd, 'rb') as f:
                    content = f.read()
                return openpyxl.load_workbook(io.BytesIO(content))
        except Exception:
            pass
        fname = os.path.basename(file_path)
        raise ValueError(f"File '{fname}' đang được mở và khóa bởi ứng dụng khác. Vui lòng đóng file lại rồi thử lại.")

def analyze_assortment_file(file_path, output_path=None):
    """
    Phân tích file Assortment theo từng tuần PO và điểm đến (CW01, CW02).
    Tự động quét thông minh qua tất cả các sheet, bỏ qua sheet nháp/tự tạo,
    chỉ tập trung vào 3 cột master (PO, Destination, SKU/No.).
    Tạo hoặc cập nhật sheet 'Per_week' với đầy đủ công thức và định dạng chuẩn.
    """
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"Không tìm thấy file: {file_path}")

    # Mở workbook (hỗ trợ đọc trực tiếp kể cả khi file đang mở trong Excel)
    wb = load_workbook_safe(file_path)
    
    # Tìm sheet dữ liệu nguồn và định vị các cột master linh hoạt
    ws_source, header_row, po_col, dest_col, sku_col = find_assortment_source_sheet(wb)

    # Đọc dữ liệu và gom nhóm theo PO
    # Cấu trúc: po_groups[po_name] = {'CW01': set(sku), 'CW02': set(sku)}
    po_groups = {}
    po_order = []

    for r in range(header_row + 1, ws_source.max_row + 1):
        po_val = ws_source.cell(r, po_col).value
        dest_val = ws_source.cell(r, dest_col).value
        sku_val = ws_source.cell(r, sku_col).value

        if po_val is None or sku_val is None:
            continue

        po_str = str(po_val).strip()
        sku_str = str(sku_val).strip()

        if not po_str or not sku_str:
            continue

        low_po = po_str.lower()
        low_sku = sku_str.lower()

        # Bỏ qua dòng tổng cộng hoặc tiêu đề phụ thừa
        if any(bad in low_po for bad in ['total', 'grand total', 'tổng cộng', 'tổng', 'count', 'sum']):
            continue
        if any(bad in low_sku for bad in ['total', 'grand total', 'tổng cộng']):
            continue
        if sku_str in ('-', '--', 'N/A', 'NA', 'None', '0'):
            continue

        dest_str = str(dest_val).strip().upper() if dest_val is not None else ""
        clean_dest = re.sub(r'[^a-zA-Z0-9]', '', dest_str).upper()

        if po_str not in po_groups:
            po_groups[po_str] = {"CW01": set(), "CW02": set()}
            po_order.append(po_str)

        # Nhận diện linh hoạt kho đích
        if "CW01" in clean_dest or clean_dest in ("CW1", "DK", "DENMARK") or "DENMARK" in dest_str:
            po_groups[po_str]["CW01"].add(sku_str)
        elif "CW02" in clean_dest or clean_dest in ("CW2", "BE", "BELGIUM") or "BELGIUM" in dest_str:
            po_groups[po_str]["CW02"].add(sku_str)

    if not po_order:
        raise ValueError(f"Không tìm thấy dòng dữ liệu PO / Mã hàng hợp lệ nào trong sheet '{ws_source.title}'.")

    # Xử lý sheet 'Per_week'
    if 'Per_week' in wb.sheetnames:
        del wb['Per_week']
    ws_out = wb.create_sheet(title='Per_week')

    # Định nghĩa Styles Apple UI
    header_fill = PatternFill(start_color="3B302A", end_color="3B302A", fill_type="solid") # Dark Roast Brown
    header_font = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
    data_font = Font(name="Segoe UI", size=10)
    bold_data_font = Font(name="Segoe UI", size=10, bold=True)
    pct_font = Font(name="Segoe UI", size=10, bold=True, color="2E5A36") # Xanh lá đậm cho tỷ lệ

    thin_border_side = Side(border_style="thin", color="E0D6C8")
    data_border = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)

    # Nền xen kẽ dịu mắt
    zebra_fill = PatternFill(start_color="FAF7F2", end_color="FAF7F2", fill_type="solid")
    highlight_fill = PatternFill(start_color="F2EBD9", end_color="F2EBD9", fill_type="solid")

    columns = [
        "PO",
        "CW01",
        "CW02",
        "Duplicated",
        "Only CW01",
        "Only CW02",
        "SKU/week",
        "% Denmark Codes Also in Belgium"
    ]

    # Ghi Header
    for c_idx, col_name in enumerate(columns, 1):
        cell = ws_out.cell(row=1, column=c_idx, value=col_name)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    ws_out.row_dimensions[1].height = 28

    # Ghi dữ liệu từng tuần
    for r_idx, po_name in enumerate(po_order, 2):
        s_cw01 = po_groups[po_name]["CW01"]
        s_cw02 = po_groups[po_name]["CW02"]

        cnt_cw01 = len(s_cw01)
        cnt_cw02 = len(s_cw02)
        duplicated = len(s_cw01 & s_cw02)
        only_cw01 = len(s_cw01 - s_cw02)
        only_cw02 = len(s_cw02 - s_cw01)

        # Ghi các giá trị
        c_po = ws_out.cell(row=r_idx, column=1, value=po_name)
        c_cw01 = ws_out.cell(row=r_idx, column=2, value=cnt_cw01)
        c_cw02 = ws_out.cell(row=r_idx, column=3, value=cnt_cw02)
        c_dup = ws_out.cell(row=r_idx, column=4, value=duplicated)
        c_o1 = ws_out.cell(row=r_idx, column=5, value=only_cw01)
        c_o2 = ws_out.cell(row=r_idx, column=6, value=only_cw02)
        
        # Công thức tính Tổng SKU và Tỷ lệ theo đúng form của Nhung
        c_sku = ws_out.cell(row=r_idx, column=7, value=f"=SUM(D{r_idx}:F{r_idx})")
        c_pct = ws_out.cell(row=r_idx, column=8, value=f"=IF(B{r_idx}>0, D{r_idx}/B{r_idx}, 0)")
        c_pct.number_format = '0.0%'

        is_even = (r_idx % 2 == 0)
        curr_fill = zebra_fill if is_even else PatternFill(fill_type=None)

        # Căn chỉnh và border
        for c_idx in range(1, 9):
            cell = ws_out.cell(row=r_idx, column=c_idx)
            cell.border = data_border
            if c_idx == 1:
                cell.font = data_font
                cell.alignment = Alignment(horizontal="left", vertical="center")
            elif c_idx == 4: # Duplicated highlight
                cell.font = bold_data_font
                cell.fill = highlight_fill
                cell.alignment = Alignment(horizontal="right", vertical="center")
            elif c_idx == 7: # SKU/week
                cell.font = bold_data_font
                cell.alignment = Alignment(horizontal="right", vertical="center")
            elif c_idx == 8: # %
                cell.font = pct_font
                cell.alignment = Alignment(horizontal="right", vertical="center")
            else:
                cell.font = data_font
                cell.alignment = Alignment(horizontal="right", vertical="center")
            
            if c_idx not in (4,) and curr_fill.fill_type:
                cell.fill = curr_fill

        ws_out.row_dimensions[r_idx].height = 22

    # Tự động điều chỉnh độ rộng cột
    for col in ws_out.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            val_str = str(cell.value or '')
            if len(val_str) > max_len:
                max_len = len(val_str)
        ws_out.column_dimensions[col_letter].width = max(max_len + 4, 12)
    
    # Cột PO và Cột % cho rộng rãi hơn
    ws_out.column_dimensions['A'].width = 30
    ws_out.column_dimensions['H'].width = 32

    # Lưu file kết quả
    if not output_path:
        base, ext = os.path.splitext(file_path)
        output_path = f"{base}_analyzed{ext}"

    try:
        wb.save(output_path)
    except PermissionError:
        import time
        ts = time.strftime("%H%M%S")
        base, ext = os.path.splitext(output_path)
        alt_path = f"{base}_{ts}{ext}"
        try:
            wb.save(alt_path)
            output_path = alt_path
        except Exception:
            out_name = os.path.basename(output_path)
            raise ValueError(f"File kết quả '{out_name}' đang được mở trong Excel. Vui lòng lưu/đóng file đó lại rồi bấm chạy lại.")
    wb.close()
    return output_path, len(po_order)

# ==============================================================================
# EMBEDDED MODULE 4: WEEKLY PO CHECKER (weekly_po_checker)
# ==============================================================================
import os
import sys
import re
import copy
import win32file
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

W_FILL_OK = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
W_FONT_OK = Font(name="Segoe UI", size=10, bold=True, color="006100")

W_FILL_WARN = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
W_FONT_WARN = Font(name="Segoe UI", size=10, bold=True, color="9C6500")

W_FILL_ERR = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
W_FONT_ERR = Font(name="Segoe UI", size=10, bold=True, color="9C0006")

W_FILL_HEAD = PatternFill(start_color="F1EADB", end_color="F1EADB", fill_type="solid")
W_FONT_HEAD = Font(name="Segoe UI", size=10, bold=True, color="2B211C")

W_BORDER_SIDE = Side(border_style="thin", color="E8E0D0")
W_BORDER = Border(left=W_BORDER_SIDE, right=W_BORDER_SIDE, top=W_BORDER_SIDE, bottom=W_BORDER_SIDE)

def weekly_safe_copy_shared(src, dst):
    """Safely copies a file even if currently opened/locked by Excel with shared read/write."""
    handle = win32file.CreateFile(
        src,
        win32file.GENERIC_READ,
        win32file.FILE_SHARE_READ | win32file.FILE_SHARE_WRITE | win32file.FILE_SHARE_DELETE,
        None,
        win32file.OPEN_EXISTING,
        win32file.FILE_ATTRIBUTE_NORMAL,
        None
    )
    with open(dst, 'wb') as out_f:
        while True:
            hr, data = win32file.ReadFile(handle, 1024 * 1024)
            if not data:
                break
            out_f.write(data)
    win32file.CloseHandle(handle)

def weekly_normalize_key(val):
    """Normalize item art code/SKU: string, strip, lower, remove leading zeros if numeric."""
    if val is None:
        return ""
    s = str(val).strip().split(".")[0]
    return s.lower()

def weekly_normalize_text_loose(text):
    """Normalize text loosely for typo-resilient comparison (collapse spaces, commas, hyphens)."""
    if not text:
        return ""
    s = str(text).lower()
    s = re.sub(r'[\s,._\-\(\)]+', ' ', s).strip()
    return s

def weekly_clean_str(val):
    return str(val).strip() if val is not None else ""

def weekly_parse_int_qty(val):
    if val is None or val == "":
        return 0
    try:
        if isinstance(val, (int, float)):
            return int(round(val))
        s = str(val).replace(",", "").strip()
        return int(round(float(s)))
    except Exception:
        return 0

def weekly_detect_pi_number(filepath):
    """Extracts PI number (e.g. S498702) from filename or sheet cells."""
    fname = os.path.basename(filepath)
    m = re.search(r'(S\d{5,7})', fname, re.IGNORECASE)
    if m:
        return m.group(1).upper()
    try:
        tmp = filepath + ".tmp_peek.xlsx"
        weekly_safe_copy_shared(filepath, tmp)
        wb = openpyxl.load_workbook(tmp, read_only=True, data_only=True)
        for sname in wb.sheetnames:
            ws = wb[sname]
            for r_idx, row in enumerate(ws.iter_rows(values_only=True)):
                if r_idx > 30:
                    break
                for cell in row:
                    if cell and isinstance(cell, str):
                        m2 = re.search(r'(S\d{5,7})', cell, re.IGNORECASE)
                        if m2:
                            wb.close()
                            if os.path.exists(tmp): os.remove(tmp)
                            return m2.group(1).upper()
        wb.close()
        if os.path.exists(tmp): os.remove(tmp)
    except Exception:
        pass
    return ""

def weekly_get_merged_val(ws, r, c):
    val = ws.cell(r, c).value
    if val is not None:
        return val
    for rng in ws.merged_cells.ranges:
        if rng.min_row <= r <= rng.max_row and rng.min_col <= c <= rng.max_col:
            return ws.cell(rng.min_row, rng.min_col).value
    return None

def weekly_locate_table_in_sheet(ws):
    """Smartly detects the header row and column positions of an order table in any sheet."""
    for r in range(1, min(ws.max_row + 1, 60)):
        row_vals = [str(weekly_get_merged_val(ws, r, c) or '').strip() for c in range(1, min(ws.max_column + 1, 80))]
        found = {}
        for c_idx, val in enumerate(row_vals, start=1):
            vl = val.lower().replace(' ', '').replace('.', '').replace('_', '')
            if 'customerartno' in vl or vl == 'no' or 'itemno' in vl:
                if 'art' not in found: found['art'] = c_idx
            if 'description' in vl or 'customeritemname' in vl:
                if 'desc' not in found: found['desc'] = c_idx
            if 'quantity' in vl or 'qty' in vl:
                if 'qty' not in found: found['qty'] = c_idx
            if 'scitemno' in vl:
                if 'sc_item' not in found: found['sc_item'] = c_idx
            if 'scitemname' in vl:
                if 'sc_name' not in found: found['sc_name'] = c_idx
        if 'art' in found and ('desc' in found or 'qty' in found):
            return r, found
    return -1, {}

def weekly_find_best_order_sheet(wb):
    for sname in wb.sheetnames:
        h_row, col_map = weekly_locate_table_in_sheet(wb[sname])
        if h_row != -1:
            return wb[sname], h_row, col_map
    return wb.active, -1, {}

def weekly_load_master_overview_index(overview_path):
    """Loads Master Overview and builds comprehensive lookup index.
    Returns: (master_by_pi, master_all)
    """
    tmp_path = overview_path + ".tmp_load.xlsx"
    weekly_safe_copy_shared(overview_path, tmp_path)
    wb = openpyxl.load_workbook(tmp_path, read_only=True, data_only=True)
    ws = wb.active

    headers = None
    h_row = 1
    for r_idx, row in enumerate(ws.iter_rows(values_only=True), start=1):
        row_str = [str(c or '').lower() for c in row]
        if any('customer art no' in c or 'customerartno' in c or 'item no' in c for c in row_str):
            headers = [str(c or '').strip() for c in row]
            h_row = r_idx
            break
        if r_idx > 10:
            break

    if not headers:
        headers = [str(c or '').strip() for c in next(ws.iter_rows(values_only=True))]
        h_row = 1

    col_map = {}
    for idx, h in enumerate(headers):
        hl = h.lower().replace(" ", "").replace(".", "").replace("_", "")
        if "pino" in hl or "pi" == hl:
            col_map["pi"] = idx
        elif "pono" in hl or "po" == hl:
            col_map["po"] = idx
        elif "customerartno" in hl or "customerart" in hl:
            col_map["art"] = idx
        elif "scitemno" in hl or ("itemno" in hl and "customer" not in hl):
            col_map["sc_item"] = idx
        elif "scitemname" in hl:
            col_map["sc_name"] = idx
        elif "customeritemname" in hl:
            col_map["cust_name"] = idx
        elif "orderedquantity" in hl or "orderquantity" in hl or "orderedqty" in hl or "qty" == hl:
            col_map["qty"] = idx

    master_by_pi = {}
    master_all = {}

    for row in ws.iter_rows(min_row=h_row + 1, values_only=True):
        if not any(row):
            continue
        pi_val = weekly_clean_str(row[col_map.get("pi", 4)] if "pi" in col_map else "")
        art_val = weekly_clean_str(row[col_map.get("art", 12)] if "art" in col_map else "")
        if not art_val:
            continue

        m_pi = re.search(r'(S\d{5,7})', pi_val, re.IGNORECASE)
        pi_key = m_pi.group(1).upper() if m_pi else pi_val.upper()

        art_key = weekly_normalize_key(art_val)
        item_data = {
            "pi": pi_key,
            "po": weekly_clean_str(row[col_map.get("po", 5)] if "po" in col_map else ""),
            "art_raw": art_val,
            "art_key": art_key,
            "sc_item": weekly_clean_str(row[col_map.get("sc_item", 13)] if "sc_item" in col_map else ""),
            "sc_name": weekly_clean_str(row[col_map.get("sc_name", 17)] if "sc_name" in col_map else ""),
            "cust_name": weekly_clean_str(row[col_map.get("cust_name", 19)] if "cust_name" in col_map else ""),
            "qty": weekly_parse_int_qty(row[col_map.get("qty", 21)] if "qty" in col_map else 0),
        }

        if pi_key:
            if pi_key not in master_by_pi:
                master_by_pi[pi_key] = {}
            if art_key in master_by_pi[pi_key]:
                master_by_pi[pi_key][art_key]["qty"] += item_data["qty"]
            else:
                master_by_pi[pi_key][art_key] = copy.deepcopy(item_data)

        if art_key not in master_all:
            master_all[art_key] = copy.deepcopy(item_data)
        else:
            master_all[art_key]["qty"] += item_data["qty"]

    wb.close()
    if os.path.exists(tmp_path):
        os.remove(tmp_path)
    return master_by_pi, master_all

def weekly_find_safe_append_col(ws):
    """Safely finds an empty column index to append columns without colliding with merged or formatted cells."""
    max_c = 0
    for r in range(1, min(ws.max_row + 1, 150)):
        for c in range(1, ws.max_column + 1):
            if ws.cell(r, c).value is not None and c > max_c:
                max_c = c
    for rng in ws.merged_cells.ranges:
        if rng.max_col > max_c:
            max_c = rng.max_col
    for col_letter, dim in ws.column_dimensions.items():
        if dim.width is not None and dim.width > 0:
            try:
                c_idx = openpyxl.utils.column_index_from_string(col_letter)
                if c_idx > max_c:
                    max_c = c_idx
            except Exception:
                pass
    return max_c + 1

def weekly_audit_file_generic(file_path, master_dict, output_suffix="_checked.xlsx", factory_items_sum=None, role_name="NHUNG"):
    """Audits an order file (Nhung or Factory) against Master PO dictionary.
    Preserves 100% original layout, styling, and formulas, appending [CHECK STATUS] and [AUDIT NOTE].
    """
    tmp_in = file_path + ".tmp_audit.xlsx"
    weekly_safe_copy_shared(file_path, tmp_in)
    wb = openpyxl.load_workbook(tmp_in)

    ws, h_row, col_map = weekly_find_best_order_sheet(wb)
    if h_row == -1:
        wb.close()
        if os.path.exists(tmp_in): os.remove(tmp_in)
        raise ValueError(f"Không nhận diện được bảng đơn hàng trong file '{os.path.basename(file_path)}'.")

    # Safe append column location
    check_col = weekly_find_safe_append_col(ws)
    note_col = check_col + 1

    # Header styling
    ws.column_dimensions[get_column_letter(check_col)].width = 22
    ws.column_dimensions[get_column_letter(note_col)].width = 52
    ws.column_dimensions[get_column_letter(check_col)].hidden = False
    ws.column_dimensions[get_column_letter(note_col)].hidden = False

    cell_h_chk = ws.cell(h_row, check_col, "[CHECK STATUS]")
    cell_h_chk.fill = W_FILL_HEAD
    cell_h_chk.font = W_FONT_HEAD
    cell_h_chk.border = W_BORDER
    cell_h_chk.alignment = Alignment(horizontal="center", vertical="center")

    cell_h_not = ws.cell(h_row, note_col, "[AUDIT NOTE]")
    cell_h_not.fill = W_FILL_HEAD
    cell_h_not.font = W_FONT_HEAD
    cell_h_not.border = W_BORDER
    cell_h_not.alignment = Alignment(horizontal="left", vertical="center")

    # First pass: sum quantities and catalog items by art_key
    file_sum_qty = {}
    row_item_map = {}

    for r in range(h_row + 1, ws.max_row + 1):
        art_val = ws.cell(r, col_map.get('art', 1)).value
        if not art_val:
            continue
        art_str = str(art_val).strip()

        # Filtering: skip totals, address headers, non-SKU texts
        art_lower = art_str.lower()
        if any(w in art_lower for w in ('total', 'grand total', 'ship-to', 'address', 'warehouse', 'gate ', 'ghent', 'belgium', 'location')):
            continue
        if not any(ch.isdigit() for ch in art_str):
            continue

        art_k = weekly_normalize_key(art_str)
        desc_val = weekly_clean_str(weekly_get_merged_val(ws, r, col_map.get('desc', 2)))
        
        # Get quantity directly or via merged
        raw_qty = ws.cell(r, col_map.get('qty', 3)).value
        if raw_qty is None and 'qty' in col_map:
            # check neighbor col (e.g. Assortment has header at 41 but data at 42)
            c_target = col_map['qty']
            for c_try in (c_target, c_target + 1, c_target - 1):
                if 1 <= c_try <= ws.max_column:
                    v_try = ws.cell(r, c_try).value
                    if v_try is not None:
                        raw_qty = v_try
                        break
        qty_val = weekly_parse_int_qty(raw_qty)

        sc_item_val = weekly_clean_str(weekly_get_merged_val(ws, r, col_map.get('sc_item', 0))) if 'sc_item' in col_map else ""
        sc_name_val = weekly_clean_str(weekly_get_merged_val(ws, r, col_map.get('sc_name', 0))) if 'sc_name' in col_map else ""

        row_item_map[r] = (art_k, art_str, desc_val, qty_val, sc_item_val, sc_name_val)
        file_sum_qty[art_k] = file_sum_qty.get(art_k, 0) + qty_val

    # Statistics
    stats = {
        "total_lines": len(row_item_map),
        "total_qty": sum(file_sum_qty.values()),
        "ok": 0,
        "warn": 0,
        "err": 0,
        "missing_in_file": [],
        "diff_lines": []
    }

    # Second pass: check each row and decorate
    for r, item_info in row_item_map.items():
        art_k, art_str, desc_val, qty_val, sc_item_val, sc_name_val = item_info

        status = "OK"
        notes = []
        fill_to_apply = W_FILL_OK
        font_to_apply = W_FONT_OK

        if art_k not in master_dict:
            status = "KHÔNG CÓ TRONG MASTER"
            fill_to_apply = W_FILL_ERR
            font_to_apply = W_FONT_ERR
            notes.append(f"Mã '{art_str}' không tồn tại trong Master PO!")
        else:
            m_item = master_dict[art_k]
            tot_file_q = file_sum_qty.get(art_k, 0)
            tot_mast_q = m_item['qty']
            if tot_file_q != tot_mast_q:
                status = "LỆCH SỐ LƯỢNG"
                fill_to_apply = W_FILL_ERR
                font_to_apply = W_FONT_ERR
                notes.append(f"Tổng số lượng ({tot_file_q}) lệch Master ({tot_mast_q})")

            # 3-way check with Factory if available
            if factory_items_sum is not None:
                tot_fac_q = factory_items_sum.get(art_k, 0)
                if tot_fac_q != tot_file_q:
                    if status == "OK":
                        status = "LỆCH VỚI NHÀ MÁY"
                        fill_to_apply = W_FILL_WARN
                        font_to_apply = W_FONT_WARN
                    notes.append(f"Nhà máy nhận {tot_fac_q} (lệch đơn hàng {tot_file_q})")

            # Description check (fuzzy)
            if desc_val and m_item['cust_name']:
                if weekly_normalize_text_loose(desc_val) != weekly_normalize_text_loose(m_item['cust_name']):
                    if status == "OK":
                        status = "LỆCH TÊN"
                        fill_to_apply = W_FILL_WARN
                        font_to_apply = W_FONT_WARN
                    notes.append(f"Tên Master: '{m_item['cust_name']}'")

            # SC Item No check
            if sc_item_val and m_item['sc_item']:
                if weekly_normalize_key(sc_item_val) != weekly_normalize_key(m_item['sc_item']):
                    if status == "OK":
                        status = "SAI MÃ SC ITEM"
                        fill_to_apply = W_FILL_ERR
                        font_to_apply = W_FONT_ERR
                    notes.append(f"SC Item Master: '{m_item['sc_item']}'")

        if not notes:
            notes.append("Khớp 100% với Master Overview")

        # Write cell values safely
        c_chk = ws.cell(r, check_col, status)
        c_chk.fill = fill_to_apply
        c_chk.font = font_to_apply
        c_chk.border = W_BORDER
        c_chk.alignment = Alignment(horizontal="center", vertical="center")

        note_text = " | ".join(notes)
        c_not = ws.cell(r, note_col, note_text)
        c_not.border = W_BORDER
        c_not.font = font_to_apply if status != "OK" else Font(name="Segoe UI", size=10, color="006100")
        c_not.alignment = Alignment(horizontal="left", vertical="center")

        if status == "OK":
            stats["ok"] += 1
        elif "LỆCH" in status or "SAI" in status:
            stats["err"] += 1
            stats["diff_lines"].append((art_str, status, note_text))
        else:
            stats["warn"] += 1
            stats["diff_lines"].append((art_str, status, note_text))

    # Items in Master but missing in this file
    for art_k, m_item in master_dict.items():
        if art_k not in file_sum_qty:
            stats["missing_in_file"].append(m_item)

    # Save output file
    base, ext = os.path.splitext(file_path)
    out_path = f"{base}{output_suffix}"
    try:
        wb.save(out_path)
    except PermissionError:
        import time
        ts = time.strftime("%H%M%S")
        out_path = f"{base}_{ts}{output_suffix}"
        wb.save(out_path)
    wb.close()
    if os.path.exists(tmp_in):
        os.remove(tmp_in)

    return out_path, stats, file_sum_qty

# ==============================================================================
# MAIN BORING TASK APPLICATION CORE
# ==============================================================================
APP_VERSION = "v1.2"
GITHUB_REPO = "rangercases/boring-task"
CACHE_FILE_NAME = ".overview_cache.pkl"
STATE_FILE_NAME = ".app_state.json"

# Per-machine module visibility (controlled centrally by code & local config.json)
CONFIG_FILE_NAME = "config.json"
APP_MODES = ("all", "cost", "images", "fabric", "auditor", "assortment", "weekly_po")

# BẢNG PHÂN QUYỀN TẬP TRUNG (Sửa tại đây để phân quyền từ xa qua Git update)
USER_PERMISSIONS = {
    "nhung": ["cost", "assortment", "weekly_po"],                                # Ms Nhung: Purchase Cost, Assortment & Weekly PO
    "thuy":  ["fabric", "auditor"],                                                # Ms Thuy: Fabric Checker & Order Auditor
    "admin": ["cost", "fabric", "images", "auditor", "assortment", "weekly_po"], # Admin: toan quyen xem tat ca cac module
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
            return ["cost", "fabric", "images", "auditor", "assortment", "weekly_po"]
        elif mode == "cost":
            return ["cost", "assortment", "weekly_po"]
        elif mode in ("cost", "fabric", "images", "auditor", "assortment", "weekly_po"):
            return [mode]
        return ["cost", "fabric", "images", "auditor", "assortment", "weekly_po"]
    except Exception:
        return ["cost", "fabric", "images", "auditor", "assortment", "weekly_po"]

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
        self.fabric_rule_mgr = FabricRuleManager(curr_dir) if FabricRuleManager else None
        self.fabric_overview_file = ""
        self.fabric_running = False

        # Order Auditor state
        self.auditor_engine = OrderAuditEngine()
        self.auditor_po_files = []
        self.auditor_cust_file = ""
        self.auditor_running = False
        self.auditor_last_report = ""

        # Assortment Analyzer state
        self.assortment_file = ""
        self.assortment_running = False
        self.assortment_last_output = ""

        # Weekly PO Checker state
        self.weekly_po_master_file = ""
        self.weekly_po_nhung_file = ""
        self.weekly_po_factory_file = ""
        self.weekly_po_running = False
        self.weekly_po_last_outputs = []
        self.weekly_po_master_cache = None
        saved_state = self.load_saved_state()

        if saved_state is not None:
            # Restore saved overview path if it still exists
            saved_ov = saved_state.get("overview_path", "")
            if saved_ov and os.path.exists(saved_ov):
                self.overview_path = saved_ov
                self.weekly_po_master_file = saved_ov
            saved_po_m = saved_state.get("weekly_po_master", "")
            if saved_po_m and os.path.exists(saved_po_m):
                self.weekly_po_master_file = saved_po_m
            
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
        if not self.weekly_po_master_file:
            self.weekly_po_master_file = self.overview_path

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
                "claim_files": self.claim_files,
                "weekly_po_master": self.weekly_po_master_file
            }
            with open(s_path, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"Error saving state: {e}")

    # ----------------------------------------------------
    # Silent Lightning-Fast Background Auto-Updater
    # ----------------------------------------------------
    def silent_auto_update(self):
        """Runs silently in background. Checks latest commit SHA from GitHub.
        If a new commit is pushed, silently downloads and updates local files in 1-2s.
        """
        try:
            curr_file = os.path.abspath(__file__)
            app_dir = os.path.dirname(curr_file)
            sha_file = os.path.join(app_dir, ".git_sha")
            
            curr_sha = ""
            if os.path.exists(sha_file):
                try:
                    with open(sha_file, "r", encoding="utf-8") as f:
                        curr_sha = f.read().strip()
                except Exception:
                    pass

            # Check latest commit from GitHub main branch
            url = f"https://api.github.com/repos/{GITHUB_REPO}/commits/main"
            req = urllib.request.Request(url, headers={"User-Agent": "BoringTask-App"})
            with urllib.request.urlopen(req, timeout=4) as resp:
                if resp.status == 200:
                    data = json.loads(resp.read().decode("utf-8"))
                    latest_sha = data.get("sha", "")
                    
                    if latest_sha and latest_sha != curr_sha:
                        # Fetch latest boring_task.pyw using exact SHA to bypass GitHub Raw CDN cache (5-min TTL)
                        raw_url = f"https://raw.githubusercontent.com/{GITHUB_REPO}/{latest_sha}/boring_task.pyw"
                        req_raw = urllib.request.Request(raw_url, headers={"User-Agent": "BoringTask-App"})
                        with urllib.request.urlopen(req_raw, timeout=5) as raw_resp:
                            if raw_resp.status == 200:
                                new_code = raw_resp.read()
                                if len(new_code) > 1000 and (b"BoringTask" in new_code or b"ClaimHelper" in new_code):
                                    tmp_file = curr_file + ".new"
                                    with open(tmp_file, "wb") as f:
                                        f.write(new_code)
                                    os.replace(tmp_file, curr_file)

                                    # Also ensure extra modules are kept up to date
                                    # Single-file architecture: No extra module downloads needed!
                                    with open(sha_file, "w", encoding="utf-8") as f:
                                        f.write(latest_sha)

                                    self.after(0, self.set_update_badge, f"✓ Đã tự động cập nhật ({latest_sha[:7]})")
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
        self.assortment_view = ctk.CTkFrame(self.view_container, fg_color="transparent")
        self.weekly_po_view = ctk.CTkFrame(self.view_container, fg_color="transparent")

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
        if "assortment" in self.app_mode:
            self.build_assortment_analyzer()
        if "weekly_po" in self.app_mode:
            self.build_weekly_po_checker()

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
        if "assortment" in self.app_mode:
            features.append({
                "icon": "📊",
                "title": "Assortment Analyzer",
                "subtitle": "Weekly PO & Dest Stats",
                "command": lambda: self.show_feature(self.assortment_view, "Assortment Analyzer"),
            })
        if "weekly_po" in self.app_mode:
            features.append({
                "icon": "🔍",
                "title": "Weekly PO Check",
                "subtitle": "Master vs Order Audit",
                "command": lambda: self.show_feature(self.weekly_po_view, "Weekly PO Checking"),
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
        self.assortment_view.pack_forget()
        self.weekly_po_view.pack_forget()
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
        if self.current_view is self.weekly_po_view:
            return self.weekly_po_on_drop_generic(event)
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
        # TOP TOOLBAR: Import Model List / Aliases
        # ========================================================
        top_bar = ctk.CTkFrame(scroll, fg_color="transparent")
        top_bar.pack(fill="x", pady=(0, 10))

        ctk.CTkLabel(
            top_bar,
            text="KIỂM TOÁN ĐƠN HÀNG (PURCHASE ORDER AUDITOR)",
            font=ctk.CTkFont(family=FONT_SANS, size=11, weight="bold"),
            text_color=BROWN
        ).pack(side="left")

        ctk.CTkButton(
            top_bar,
            text="+ Nạp Thêm Model List...",
            font=ctk.CTkFont(family=FONT_SANS, size=11),
            fg_color="transparent",
            text_color=ROAST,
            hover_color=SAND,
            border_width=1,
            border_color=LINE,
            corner_radius=14,
            height=28,
            command=self.auditor_import_model_file
        ).pack(side="right")

        # ========================================================
        # CARD 1: File Đơn Hàng Gốc Của Khách (Customer Order / Master)
        # ========================================================
        self.card_auditor_cust = ctk.CTkFrame(scroll, fg_color=CARD, corner_radius=16, border_width=1, border_color=LINE)
        self.card_auditor_cust.pack(fill="x", pady=(0, 20))

        top_cust = ctk.CTkFrame(self.card_auditor_cust, fg_color="transparent")
        top_cust.pack(fill="x", padx=24, pady=(18, 10))

        ctk.CTkLabel(
            top_cust,
            text="1. FILE ĐƠN HÀNG GỐC CỦA KHÁCH (CUSTOMER ORDER / MASTER DATA)",
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
        # CARD 2: File PO Cần Duyệt (Purchase Order Thủy Lập)
        # ========================================================
        self.card_auditor_po = ctk.CTkFrame(scroll, fg_color=CARD, corner_radius=16, border_width=1, border_color=LINE)
        self.card_auditor_po.pack(fill="x", pady=(0, 20))

        top_po = ctk.CTkFrame(self.card_auditor_po, fg_color="transparent")
        top_po.pack(fill="x", padx=24, pady=(18, 10))

        ctk.CTkLabel(
            top_po,
            text="2. FILE PURCHASE ORDER (PO / PI CẦN KIỂM TOÁN)",
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
        if hasattr(self, 'auditor_po_files') and self.auditor_po_files:
            for pofile in self.auditor_po_files:
                if not os.path.exists(pofile): continue
                sz = format_file_size(os.path.getsize(pofile))
                fn = os.path.basename(pofile)
                row = ctk.CTkFrame(self.auditor_po_status_box, fg_color="transparent")
                row.pack(fill="x", padx=16, pady=2)

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
                text="Kéo thả file PO của Thuy vào đây",
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
                text="Kéo thả file của Life vào đây",
                font=ctk.CTkFont(family=FONT_SANS, size=11),
                text_color=BROWN,
                pady=18
            ).pack(fill="both", expand=True)

    def auditor_browse_po_file(self):
        paths = filedialog.askopenfilenames(
            title="Chọn file Purchase Order (có thể chọn nhiều file)",
            filetypes=[("Excel Files", "*.xlsx *.xls")]
        )
        if paths:
            self.auditor_po_files = [os.path.abspath(p) for p in paths]
            self.auditor_render_po_status()

    def auditor_clear_po_file(self):
        self.auditor_po_files = []
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

    def auditor_import_model_file(self):
        path = filedialog.askopenfilename(
            title="Chọn file Model List để nạp thêm (Model List All.xlsx)",
            filetypes=[("Excel Files", "*.xlsx *.xls")]
        )
        if path:
            try:
                count = self.auditor_engine.model_kb.import_models_from_excel(path)
                messagebox.showinfo("Thành công", f"Đã nạp thêm {count} models vào cơ sở dữ liệu kiểm toán!")
            except Exception as e:
                messagebox.showerror("Lỗi", f"Không thể đọc file Model List: {e}")

    def auditor_on_drop_po(self, event):
        paths = [p for p in parse_drop_paths(event.data) if os.path.isfile(p)]
        if paths:
            self.auditor_po_files = paths
            self.auditor_render_po_status()

    def auditor_on_drop_cust(self, event):
        paths = [p for p in parse_drop_paths(event.data) if os.path.isfile(p)]
        if paths:
            self.auditor_cust_file = paths[0]
            self.auditor_render_cust_status()

    def auditor_start_process(self):
        if self.auditor_running:
            return
        if not hasattr(self, 'auditor_po_files') or not self.auditor_po_files or not all(os.path.exists(p) for p in self.auditor_po_files):
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
            po_doc = POParsedDoc(self.auditor_po_files)
            self.auditor_prog.set(0.6)
            cust_catalog = CustomerCatalog(self.auditor_cust_file)
            self.auditor_prog.set(0.85)

            results = self.auditor_engine.audit(po_doc, cust_catalog)

            # Export Excel directly annotated on PO copy
            po_dir = os.path.dirname(self.auditor_po_files[0])
            out_paths = export_audit_excel(results, "", source_po_path=self.auditor_po_files)
            if out_paths:
                self.auditor_last_report = out_paths[0]
                out_fn = self.auditor_last_report
            else:
                self.auditor_last_report = ""
                out_fn = "" 

            self.after(0, self.auditor_finish_ui, len(results), results, out_fn, "")
        except Exception as e:
            self.after(0, self.auditor_finish_ui, 0, [], "", str(e))

    def auditor_finish_ui(self, total_items, results, out_file, err_msg):
        self.auditor_running = False
        self.auditor_btn_run.configure(state="normal", fg_color=ROAST, text="Bắt Đầu Kiểm Toán")
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

    # ====================================================
    # MODULE: Assortment Analyzer (Weekly PO & Dest Stats)
    # ====================================================
    def build_assortment_analyzer(self):
        self.build_feature_nav(self.assortment_view, "Assortment Analyzer")

        scroll = ctk.CTkScrollableFrame(self.assortment_view, fg_color="transparent")
        scroll.pack(fill="both", expand=True, padx=36, pady=(0, 16))

        # CARD 1: File Nguồn Assortment
        self.card_assort_file = ctk.CTkFrame(scroll, fg_color=CARD, corner_radius=16, border_width=1, border_color=LINE)
        self.card_assort_file.pack(fill="x", pady=(0, 20))

        top_file = ctk.CTkFrame(self.card_assort_file, fg_color="transparent")
        top_file.pack(fill="x", padx=24, pady=(18, 10))

        ctk.CTkLabel(
            top_file,
            text="FILE EXCEL ASSORTMENT HỆ THỐNG",
            font=ctk.CTkFont(family=FONT_SANS, size=11, weight="bold"),
            text_color=BROWN
        ).pack(side="left")

        btn_box = ctk.CTkFrame(top_file, fg_color="transparent")
        btn_box.pack(side="right")

        self.btn_pick_assort = ctk.CTkButton(
            btn_box,
            text="+ Chọn File Assortment...",
            font=ctk.CTkFont(family=FONT_SANS, size=11),
            fg_color="transparent",
            text_color=ROAST,
            hover_color=SAND,
            border_width=1,
            border_color=LINE,
            corner_radius=14,
            width=175,
            height=28,
            command=self.assortment_browse_file
        )
        self.btn_pick_assort.pack(side="left", padx=4)

        self.btn_clear_assort = ctk.CTkButton(
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
            command=self.assortment_clear_file
        )
        self.btn_clear_assort.pack(side="left", padx=4)

        ctk.CTkFrame(self.card_assort_file, height=1, fg_color=LINE).pack(fill="x", padx=24, pady=(0, 14))

        self.assort_status_box = ctk.CTkFrame(
            self.card_assort_file, fg_color=SAND, corner_radius=12, border_width=1, border_color=FAINT
        )
        self.assort_status_box.pack(fill="x", padx=24, pady=(0, 20))
        self.assortment_render_file_status()

        # ACTION BUTTON
        btn_run_box = ctk.CTkFrame(scroll, fg_color="transparent")
        btn_run_box.pack(fill="x", pady=(4, 14))

        self.btn_run_assort = ctk.CTkButton(
            btn_run_box,
            text="Bắt Đầu Phân Tích & Tạo Sheet Thống Kê (Per_week)",
            font=ctk.CTkFont(family=FONT_SANS, size=13, weight="bold"),
            fg_color=ROAST,
            hover_color=MOSS,
            text_color=IVORY,
            corner_radius=24,
            height=48,
            command=self.assortment_start_processing
        )
        self.btn_run_assort.pack(fill="x")

        # PROGRESS BAR
        self.prog_bar_assort = ctk.CTkProgressBar(scroll, progress_color=MOSS, fg_color=LINE, height=3, corner_radius=2)
        self.prog_bar_assort.set(0)
        self.prog_bar_assort.pack(fill="x", pady=(0, 20))

        # CARD 2: KẾT QUẢ
        self.card_assort_results = ctk.CTkFrame(scroll, fg_color=CARD, corner_radius=16, border_width=1, border_color=LINE)
        self.card_assort_results.pack(fill="both", expand=True, pady=(0, 20))

        top_res = ctk.CTkFrame(self.card_assort_results, fg_color="transparent")
        top_res.pack(fill="x", padx=24, pady=(18, 10))

        ctk.CTkLabel(
            top_res,
            text="KẾT QUẢ PHÂN TÍCH",
            font=ctk.CTkFont(family=FONT_SANS, size=11, weight="bold"),
            text_color=BROWN
        ).pack(side="left")

        ctk.CTkFrame(self.card_assort_results, height=1, fg_color=LINE).pack(fill="x", padx=24, pady=(0, 14))

        self.assort_feed = ctk.CTkFrame(
            self.card_assort_results, fg_color=SAND, corner_radius=12, border_width=1, border_color=FAINT
        )
        self.assort_feed.pack(fill="both", expand=True, padx=24, pady=(0, 20))
        self.assortment_render_initial_feed()

        # Drag and Drop support
        if getattr(self, "has_dnd", False):
            for t in [self.card_assort_file, self.assort_status_box]:
                try:
                    t.drop_target_register(tkdnd.DND_FILES)
                    t.dnd_bind('<<Drop>>', lambda e: self.assortment_on_drop(e))
                except Exception:
                    pass

    def assortment_render_initial_feed(self):
        for child in self.assort_feed.winfo_children():
            child.destroy()
        ctk.CTkLabel(
            self.assort_feed,
            text="Kéo thả hoặc chọn file Excel Assortment rồi bấm 'Bắt Đầu Phân Tích'.\nỨng dụng sẽ tự động tính SKU từng kho (CW01, CW02), mã trùng lặp và % tỷ lệ theo từng tuần.",
            font=ctk.CTkFont(family=FONT_SANS, size=11),
            text_color=BROWN,
            justify="center",
            pady=32
        ).pack(fill="both", expand=True)

    def assortment_render_file_status(self):
        for child in self.assort_status_box.winfo_children():
            child.destroy()
        if self.assortment_file and os.path.exists(self.assortment_file):
            sz = format_file_size(os.path.getsize(self.assortment_file))
            fn = os.path.basename(self.assortment_file)
            row = ctk.CTkFrame(self.assort_status_box, fg_color="transparent")
            row.pack(fill="x", padx=16, pady=12)

            ctk.CTkLabel(row, text="📊", font=ctk.CTkFont(size=16), text_color=ROAST).pack(side="left", padx=(0, 8))
            info_box = ctk.CTkFrame(row, fg_color="transparent")
            info_box.pack(side="left", fill="x", expand=True)

            ctk.CTkLabel(info_box, text=fn, font=ctk.CTkFont(family=FONT_SANS, size=11, weight="bold"), text_color=ROAST, anchor="w").pack(anchor="w")
            ctk.CTkLabel(info_box, text=f"Dung lượng: {sz}", font=ctk.CTkFont(family=FONT_SANS, size=10), text_color=BROWN, anchor="w").pack(anchor="w")
        else:
            ctk.CTkLabel(
                self.assort_status_box,
                text="Chưa chọn file Assortment.\nKéo & thả file Excel (.xlsx) vào đây hoặc bấm nút Chọn File ở trên.",
                font=ctk.CTkFont(family=FONT_SANS, size=11),
                text_color=FAINT,
                justify="center",
                pady=16
            ).pack(fill="both", expand=True)

    def assortment_browse_file(self):
        path = filedialog.askopenfilename(
            title="Chọn file Excel Assortment",
            filetypes=[("Excel Files", "*.xlsx;*.xlsm"), ("All Files", "*.*")]
        )
        if path:
            try:
                validate_excel_file(path)
                self.assortment_file = os.path.abspath(path)
                self.assortment_render_file_status()
            except Exception as e:
                messagebox.showerror("File không hợp lệ", str(e))

    def assortment_clear_file(self):
        self.assortment_file = ""
        self.assortment_render_file_status()

    def assortment_on_drop(self, event):
        paths = parse_drop_paths(event.data)
        for p in paths:
            if p.lower().endswith(('.xlsx', '.xlsm')) and not os.path.basename(p).startswith('~$'):
                try:
                    validate_excel_file(p)
                    self.assortment_file = os.path.abspath(p)
                    self.assortment_render_file_status()
                    break
                except Exception as e:
                    messagebox.showerror("File không hợp lệ", str(e))

    def assortment_start_processing(self):
        if self.assortment_running:
            return
        if not self.assortment_file or not os.path.exists(self.assortment_file):
            messagebox.showwarning("Thiếu dữ liệu", "Vui lòng chọn file Excel Assortment trước khi phân tích.")
            return

        self.assortment_running = True
        self.btn_run_assort.configure(state="disabled", text="Đang phân tích dữ liệu...")
        self.prog_bar_assort.set(0.3)

        threading.Thread(target=self._assortment_worker, daemon=True).start()

    def _assortment_worker(self):
        try:
            out_path, num_weeks = analyze_assortment_file(self.assortment_file)
            self.assortment_last_output = out_path
            self.after(0, self._assortment_success, out_path, num_weeks)
        except Exception as e:
            self.after(0, self._assortment_error, str(e))

    def _assortment_success(self, out_path, num_weeks):
        self.assortment_running = False
        self.btn_run_assort.configure(state="normal", text="Bắt Đầu Phân Tích & Tạo Sheet Thống Kê (Per_week)")
        self.prog_bar_assort.set(1.0)

        for child in self.assort_feed.winfo_children():
            child.destroy()

        res_box = ctk.CTkFrame(self.assort_feed, fg_color="transparent")
        res_box.pack(fill="x", padx=16, pady=16)

        ctk.CTkLabel(
            res_box,
            text=f"✓ Đã phân tích thành công {num_weeks} tuần PO!",
            font=ctk.CTkFont(family=FONT_SANS, size=13, weight="bold"),
            text_color=MOSS
        ).pack(anchor="w", pady=(0, 6))

        fn = os.path.basename(out_path)
        ctk.CTkLabel(
            res_box,
            text=f"File kết quả: {fn}\nĐã thêm sheet 'Per_week' với đầy đủ công thức và định dạng chuẩn.",
            font=ctk.CTkFont(family=FONT_SANS, size=11),
            text_color=ROAST,
            justify="left"
        ).pack(anchor="w", pady=(0, 14))

        act_row = ctk.CTkFrame(res_box, fg_color="transparent")
        act_row.pack(anchor="w")

        ctk.CTkButton(
            act_row,
            text="Mở File Excel",
            font=ctk.CTkFont(family=FONT_SANS, size=11, weight="bold"),
            fg_color=ROAST,
            hover_color=MOSS,
            text_color=IVORY,
            corner_radius=14,
            height=30,
            command=lambda: os.startfile(out_path)
        ).pack(side="left", padx=(0, 8))

        ctk.CTkButton(
            act_row,
            text="Mở Thư Mục",
            font=ctk.CTkFont(family=FONT_SANS, size=11),
            fg_color="transparent",
            text_color=ROAST,
            hover_color=SAND,
            border_width=1,
            border_color=LINE,
            corner_radius=14,
            height=30,
            command=lambda: reveal_in_explorer(out_path)
        ).pack(side="left")

    def _assortment_error(self, err_msg):
        self.assortment_running = False
        self.btn_run_assort.configure(state="normal", text="Bắt Đầu Phân Tích & Tạo Sheet Thống Kê (Per_week)")
        self.prog_bar_assort.set(0)

        for child in self.assort_feed.winfo_children():
            child.destroy()

        err_box = ctk.CTkFrame(self.assort_feed, fg_color="transparent")
        err_box.pack(fill="x", padx=16, pady=16)

        ctk.CTkLabel(
            err_box,
            text="❌ Có lỗi xảy ra trong quá trình phân tích:",
            font=ctk.CTkFont(family=FONT_SANS, size=12, weight="bold"),
            text_color=DANGER
        ).pack(anchor="w", pady=(0, 4))

        ctk.CTkLabel(
            err_box,
            text=err_msg,
            font=ctk.CTkFont(family=FONT_SANS, size=11),
            text_color=ROAST,
            wraplength=650,
            justify="left"
        ).pack(anchor="w")

# ====================================================
    # MODULE: Weekly PO Checking (Master vs Nhung vs Factory)
    # ====================================================
    def build_weekly_po_checker(self):
        self.build_feature_nav(self.weekly_po_view, "Weekly PO Checking")

        scroll = ctk.CTkScrollableFrame(self.weekly_po_view, fg_color="transparent")
        scroll.pack(fill="both", expand=True, padx=36, pady=(0, 16))

        # CARD 1: File Master Overview (Cached & Persistent)
        self.card_w_master = ctk.CTkFrame(scroll, fg_color=CARD, corner_radius=16, border_width=1, border_color=LINE)
        self.card_w_master.pack(fill="x", pady=(0, 16))

        top_m = ctk.CTkFrame(self.card_w_master, fg_color="transparent")
        top_m.pack(fill="x", padx=24, pady=(16, 8))

        ctk.CTkLabel(
            top_m,
            text="1. FILE MASTER OVERVIEW (DỮ LIỆU GỐC CHUẨN)",
            font=ctk.CTkFont(family=FONT_SANS, size=11, weight="bold"),
            text_color=BROWN
        ).pack(side="left")

        btn_box_m = ctk.CTkFrame(top_m, fg_color="transparent")
        btn_box_m.pack(side="right")

        self.btn_w_pick_master = ctk.CTkButton(
            btn_box_m,
            text="+ Cập Nhật Master...",
            font=ctk.CTkFont(family=FONT_SANS, size=11),
            fg_color="transparent",
            text_color=ROAST,
            hover_color=SAND,
            border_width=1,
            border_color=LINE,
            corner_radius=14,
            width=150,
            height=28,
            command=self.weekly_po_browse_master
        )
        self.btn_w_pick_master.pack(side="left", padx=4)

        ctk.CTkFrame(self.card_w_master, height=1, fg_color=LINE).pack(fill="x", padx=24, pady=(0, 12))

        self.weekly_master_status_box = ctk.CTkFrame(
            self.card_w_master, fg_color=SAND, corner_radius=12, border_width=1, border_color=FAINT
        )
        self.weekly_master_status_box.pack(fill="x", padx=24, pady=(0, 16))
        self.weekly_po_render_master_status()

        # CARD 2: Đơn Hàng Của Nhung Gửi Đi (Bắt Buộc)
        self.card_w_nhung = ctk.CTkFrame(scroll, fg_color=CARD, corner_radius=16, border_width=1, border_color=LINE)
        self.card_w_nhung.pack(fill="x", pady=(0, 16))

        top_nh = ctk.CTkFrame(self.card_w_nhung, fg_color="transparent")
        top_nh.pack(fill="x", padx=24, pady=(16, 8))

        ctk.CTkLabel(
            top_nh,
            text="2. ĐƠN HÀNG GỬI ĐI (FILE CỦA NHUNG — BẮT BUỘC)",
            font=ctk.CTkFont(family=FONT_SANS, size=11, weight="bold"),
            text_color=BROWN
        ).pack(side="left")

        btn_box_nh = ctk.CTkFrame(top_nh, fg_color="transparent")
        btn_box_nh.pack(side="right")

        self.btn_w_pick_nhung = ctk.CTkButton(
            btn_box_nh,
            text="+ Chọn File Đơn Hàng...",
            font=ctk.CTkFont(family=FONT_SANS, size=11),
            fg_color="transparent",
            text_color=ROAST,
            hover_color=SAND,
            border_width=1,
            border_color=LINE,
            corner_radius=14,
            width=165,
            height=28,
            command=self.weekly_po_browse_nhung
        )
        self.btn_w_pick_nhung.pack(side="left", padx=4)

        self.btn_w_clear_nhung = ctk.CTkButton(
            btn_box_nh,
            text="Xóa File",
            font=ctk.CTkFont(family=FONT_SANS, size=11),
            fg_color="transparent",
            text_color=DANGER,
            hover_color=SAND,
            border_width=0,
            corner_radius=14,
            width=65,
            height=28,
            command=self.weekly_po_clear_nhung
        )
        self.btn_w_clear_nhung.pack(side="left", padx=4)

        ctk.CTkFrame(self.card_w_nhung, height=1, fg_color=LINE).pack(fill="x", padx=24, pady=(0, 12))

        self.weekly_nhung_status_box = ctk.CTkFrame(
            self.card_w_nhung, fg_color=SAND, corner_radius=12, border_width=1, border_color=FAINT
        )
        self.weekly_nhung_status_box.pack(fill="x", padx=24, pady=(0, 16))
        self.weekly_po_render_nhung_status()

        # CARD 3: Đơn Hàng Nhà Máy Gửi Lại (Tùy Chọn / Linh Hoạt)
        self.card_w_factory = ctk.CTkFrame(scroll, fg_color=CARD, corner_radius=16, border_width=1, border_color=LINE)
        self.card_w_factory.pack(fill="x", pady=(0, 16))

        top_fac = ctk.CTkFrame(self.card_w_factory, fg_color="transparent")
        top_fac.pack(fill="x", padx=24, pady=(16, 8))

        ctk.CTkLabel(
            top_fac,
            text="3. ĐƠN HÀNG NHÀ MÁY PHẢN HỒI (TÙY CHỌN / CÓ THỂ CHECK SAU)",
            font=ctk.CTkFont(family=FONT_SANS, size=11, weight="bold"),
            text_color=BROWN
        ).pack(side="left")

        btn_box_fac = ctk.CTkFrame(top_fac, fg_color="transparent")
        btn_box_fac.pack(side="right")

        self.btn_w_pick_factory = ctk.CTkButton(
            btn_box_fac,
            text="+ Chọn File Nhà Máy...",
            font=ctk.CTkFont(family=FONT_SANS, size=11),
            fg_color="transparent",
            text_color=ROAST,
            hover_color=SAND,
            border_width=1,
            border_color=LINE,
            corner_radius=14,
            width=165,
            height=28,
            command=self.weekly_po_browse_factory
        )
        self.btn_w_pick_factory.pack(side="left", padx=4)

        self.btn_w_clear_factory = ctk.CTkButton(
            btn_box_fac,
            text="Xóa File",
            font=ctk.CTkFont(family=FONT_SANS, size=11),
            fg_color="transparent",
            text_color=DANGER,
            hover_color=SAND,
            border_width=0,
            corner_radius=14,
            width=65,
            height=28,
            command=self.weekly_po_clear_factory
        )
        self.btn_w_clear_factory.pack(side="left", padx=4)

        ctk.CTkFrame(self.card_w_factory, height=1, fg_color=LINE).pack(fill="x", padx=24, pady=(0, 12))

        self.weekly_factory_status_box = ctk.CTkFrame(
            self.card_w_factory, fg_color=SAND, corner_radius=12, border_width=1, border_color=FAINT
        )
        self.weekly_factory_status_box.pack(fill="x", padx=24, pady=(0, 16))
        self.weekly_po_render_factory_status()

        # ACTION BUTTON
        btn_run_box = ctk.CTkFrame(scroll, fg_color="transparent")
        btn_run_box.pack(fill="x", pady=(4, 12))

        self.btn_run_weekly_po = ctk.CTkButton(
            btn_run_box,
            text="Bắt Đầu Đối Soát & Thêm Cột Check Vào File",
            font=ctk.CTkFont(family=FONT_SANS, size=13, weight="bold"),
            fg_color=ROAST,
            hover_color=MOSS,
            text_color=IVORY,
            corner_radius=24,
            height=48,
            command=self.weekly_po_start_processing
        )
        self.btn_run_weekly_po.pack(fill="x")

        # PROGRESS BAR
        self.prog_bar_weekly_po = ctk.CTkProgressBar(scroll, progress_color=MOSS, fg_color=LINE, height=3, corner_radius=2)
        self.prog_bar_weekly_po.set(0)
        self.prog_bar_weekly_po.pack(fill="x", pady=(0, 16))

        # CARD 4: KẾT QUẢ ĐỐI SOÁT
        self.card_w_results = ctk.CTkFrame(scroll, fg_color=CARD, corner_radius=16, border_width=1, border_color=LINE)
        self.card_w_results.pack(fill="both", expand=True, pady=(0, 20))

        top_res = ctk.CTkFrame(self.card_w_results, fg_color="transparent")
        top_res.pack(fill="x", padx=24, pady=(16, 8))

        ctk.CTkLabel(
            top_res,
            text="KẾT QUẢ ĐỐI SOÁT & BÁO CÁO",
            font=ctk.CTkFont(family=FONT_SANS, size=11, weight="bold"),
            text_color=BROWN
        ).pack(side="left")

        ctk.CTkFrame(self.card_w_results, height=1, fg_color=LINE).pack(fill="x", padx=24, pady=(0, 12))

        self.weekly_po_feed = ctk.CTkFrame(
            self.card_w_results, fg_color=SAND, corner_radius=12, border_width=1, border_color=FAINT
        )
        self.weekly_po_feed.pack(fill="both", expand=True, padx=24, pady=(0, 20))
        self.weekly_po_render_initial_feed()

        # Drag and Drop support
        if getattr(self, "has_dnd", False):
            for t in [self.card_w_master, self.weekly_master_status_box]:
                try:
                    t.drop_target_register(tkdnd.DND_FILES)
                    t.dnd_bind('<<Drop>>', lambda e: self.weekly_po_on_drop_master(e))
                except Exception:
                    pass
            for t in [self.card_w_nhung, self.weekly_nhung_status_box]:
                try:
                    t.drop_target_register(tkdnd.DND_FILES)
                    t.dnd_bind('<<Drop>>', lambda e: self.weekly_po_on_drop_nhung(e))
                except Exception:
                    pass
            for t in [self.card_w_factory, self.weekly_factory_status_box]:
                try:
                    t.drop_target_register(tkdnd.DND_FILES)
                    t.dnd_bind('<<Drop>>', lambda e: self.weekly_po_on_drop_factory(e))
                except Exception:
                    pass

    def weekly_po_render_initial_feed(self):
        for child in self.weekly_po_feed.winfo_children():
            child.destroy()
        ctk.CTkLabel(
            self.weekly_po_feed,
            text="Chọn file đơn hàng cần thẩm định rồi bấm 'Bắt Đầu Đối Soát'.\n• Có thể check riêng file của Nhung trước khi gửi đi.\n• Khi có thêm file Nhà máy, ứng dụng sẽ đối chiếu 3 chiều và phát hiện mã thiếu, lệch số lượng, lệch tên.",
            font=ctk.CTkFont(family=FONT_SANS, size=11),
            text_color=BROWN,
            justify="center",
            pady=32
        ).pack(fill="both", expand=True)

    def weekly_po_render_master_status(self):
        for child in self.weekly_master_status_box.winfo_children():
            child.destroy()
        if self.weekly_po_master_file and os.path.exists(self.weekly_po_master_file):
            sz = format_file_size(os.path.getsize(self.weekly_po_master_file))
            fn = os.path.basename(self.weekly_po_master_file)
            row = ctk.CTkFrame(self.weekly_master_status_box, fg_color="transparent")
            row.pack(fill="x", padx=16, pady=10)

            ctk.CTkLabel(row, text="💾", font=ctk.CTkFont(size=16), text_color=MOSS).pack(side="left", padx=(0, 8))
            info_box = ctk.CTkFrame(row, fg_color="transparent")
            info_box.pack(side="left", fill="x", expand=True)

            ctk.CTkLabel(info_box, text=f"{fn}  [Đã nạp sẵn]", font=ctk.CTkFont(family=FONT_SANS, size=11, weight="bold"), text_color=ROAST, anchor="w").pack(anchor="w")
            ctk.CTkLabel(info_box, text=f"{self.weekly_po_master_file}  •  {sz}", font=ctk.CTkFont(family=FONT_SANS, size=10), text_color=BROWN, anchor="w").pack(anchor="w")
        else:
            ctk.CTkLabel(
                self.weekly_master_status_box,
                text="Chưa có file Master Overview. Kéo & thả file hoặc bấm 'Cập Nhật Master' ở trên.",
                font=ctk.CTkFont(family=FONT_SANS, size=11),
                text_color=FAINT,
                justify="center",
                pady=14
            ).pack(fill="both", expand=True)

    def weekly_po_render_nhung_status(self):
        for child in self.weekly_nhung_status_box.winfo_children():
            child.destroy()
        if self.weekly_po_nhung_file and os.path.exists(self.weekly_po_nhung_file):
            sz = format_file_size(os.path.getsize(self.weekly_po_nhung_file))
            fn = os.path.basename(self.weekly_po_nhung_file)
            row = ctk.CTkFrame(self.weekly_nhung_status_box, fg_color="transparent")
            row.pack(fill="x", padx=16, pady=10)

            ctk.CTkLabel(row, text="📄", font=ctk.CTkFont(size=16), text_color=ROAST).pack(side="left", padx=(0, 8))
            info_box = ctk.CTkFrame(row, fg_color="transparent")
            info_box.pack(side="left", fill="x", expand=True)

            ctk.CTkLabel(info_box, text=fn, font=ctk.CTkFont(family=FONT_SANS, size=11, weight="bold"), text_color=ROAST, anchor="w").pack(anchor="w")
            ctk.CTkLabel(info_box, text=f"Dung lượng: {sz}", font=ctk.CTkFont(family=FONT_SANS, size=10), text_color=BROWN, anchor="w").pack(anchor="w")
        else:
            ctk.CTkLabel(
                self.weekly_nhung_status_box,
                text="Chưa chọn file đơn hàng của Nhung.\nKéo & thả file Excel (.xlsx) vào đây hoặc bấm nút Chọn File ở trên.",
                font=ctk.CTkFont(family=FONT_SANS, size=11),
                text_color=FAINT,
                justify="center",
                pady=14
            ).pack(fill="both", expand=True)

    def weekly_po_render_factory_status(self):
        for child in self.weekly_factory_status_box.winfo_children():
            child.destroy()
        if self.weekly_po_factory_file and os.path.exists(self.weekly_po_factory_file):
            sz = format_file_size(os.path.getsize(self.weekly_po_factory_file))
            fn = os.path.basename(self.weekly_po_factory_file)
            row = ctk.CTkFrame(self.weekly_factory_status_box, fg_color="transparent")
            row.pack(fill="x", padx=16, pady=10)

            ctk.CTkLabel(row, text="🏭", font=ctk.CTkFont(size=16), text_color=ROAST).pack(side="left", padx=(0, 8))
            info_box = ctk.CTkFrame(row, fg_color="transparent")
            info_box.pack(side="left", fill="x", expand=True)

            ctk.CTkLabel(info_box, text=fn, font=ctk.CTkFont(family=FONT_SANS, size=11, weight="bold"), text_color=ROAST, anchor="w").pack(anchor="w")
            ctk.CTkLabel(info_box, text=f"Dung lượng: {sz}", font=ctk.CTkFont(family=FONT_SANS, size=10), text_color=BROWN, anchor="w").pack(anchor="w")
        else:
            ctk.CTkLabel(
                self.weekly_factory_status_box,
                text="(Tùy chọn) Kéo & thả file phản hồi của Nhà máy nếu có.\nNếu chưa có, ứng dụng sẽ chỉ kiểm tra đối chiếu file của Nhung với Master.",
                font=ctk.CTkFont(family=FONT_SANS, size=11),
                text_color=FAINT,
                justify="center",
                pady=14
            ).pack(fill="both", expand=True)

    def weekly_po_browse_master(self):
        path = filedialog.askopenfilename(
            title="Chọn file Master Overview",
            filetypes=[("Excel Files", "*.xlsx;*.xlsm"), ("All Files", "*.*")]
        )
        if path:
            try:
                validate_excel_file(path)
                self.weekly_po_master_file = os.path.abspath(path)
                self.weekly_po_master_cache = None # invalidate cache
                self.save_state()
                self.weekly_po_render_master_status()
            except Exception as e:
                messagebox.showerror("File không hợp lệ", str(e))

    def weekly_po_browse_nhung(self):
        path = filedialog.askopenfilename(
            title="Chọn file đơn hàng của Nhung",
            filetypes=[("Excel Files", "*.xlsx;*.xlsm"), ("All Files", "*.*")]
        )
        if path:
            try:
                validate_excel_file(path)
                self.weekly_po_nhung_file = os.path.abspath(path)
                self.weekly_po_render_nhung_status()
            except Exception as e:
                messagebox.showerror("File không hợp lệ", str(e))

    def weekly_po_clear_nhung(self):
        self.weekly_po_nhung_file = ""
        self.weekly_po_render_nhung_status()

    def weekly_po_browse_factory(self):
        path = filedialog.askopenfilename(
            title="Chọn file đơn hàng Nhà máy",
            filetypes=[("Excel Files", "*.xlsx;*.xlsm"), ("All Files", "*.*")]
        )
        if path:
            try:
                validate_excel_file(path)
                self.weekly_po_factory_file = os.path.abspath(path)
                self.weekly_po_render_factory_status()
            except Exception as e:
                messagebox.showerror("File không hợp lệ", str(e))

    def weekly_po_clear_factory(self):
        self.weekly_po_factory_file = ""
        self.weekly_po_render_factory_status()

    def weekly_po_on_drop_master(self, event):
        paths = parse_drop_paths(event.data)
        for p in paths:
            if p.lower().endswith(('.xlsx', '.xlsm')) and not os.path.basename(p).startswith('~$'):
                try:
                    validate_excel_file(p)
                    self.weekly_po_master_file = os.path.abspath(p)
                    self.weekly_po_master_cache = None
                    self.save_state()
                    self.weekly_po_render_master_status()
                    break
                except Exception as e:
                    messagebox.showerror("File không hợp lệ", str(e))

    def weekly_po_on_drop_nhung(self, event):
        paths = parse_drop_paths(event.data)
        for p in paths:
            if p.lower().endswith(('.xlsx', '.xlsm')) and not os.path.basename(p).startswith('~$'):
                try:
                    validate_excel_file(p)
                    self.weekly_po_nhung_file = os.path.abspath(p)
                    self.weekly_po_render_nhung_status()
                    break
                except Exception as e:
                    messagebox.showerror("File không hợp lệ", str(e))

    def weekly_po_on_drop_factory(self, event):
        paths = parse_drop_paths(event.data)
        for p in paths:
            if p.lower().endswith(('.xlsx', '.xlsm')) and not os.path.basename(p).startswith('~$'):
                try:
                    validate_excel_file(p)
                    self.weekly_po_factory_file = os.path.abspath(p)
                    self.weekly_po_render_factory_status()
                    break
                except Exception as e:
                    messagebox.showerror("File không hợp lệ", str(e))

    def weekly_po_on_drop_generic(self, event):
        paths = parse_drop_paths(event.data)
        for p in paths:
            if os.path.isfile(p) and p.endswith('.xlsx') and not os.path.basename(p).startswith('~$'):
                fn = os.path.basename(p).lower()
                if "overview" in fn:
                    self.weekly_po_master_file = os.path.abspath(p)
                    self.weekly_po_master_cache = None
                    self.save_state()
                    self.weekly_po_render_master_status()
                elif "orderitem" in fn or "sample" in fn:
                    self.weekly_po_factory_file = os.path.abspath(p)
                    self.weekly_po_render_factory_status()
                else:
                    self.weekly_po_nhung_file = os.path.abspath(p)
                    self.weekly_po_render_nhung_status()

    def weekly_po_start_processing(self):
        if self.weekly_po_running:
            return
        if not self.weekly_po_master_file or not os.path.exists(self.weekly_po_master_file):
            messagebox.showwarning("Thiếu Master", "Vui lòng chọn hoặc kéo thả file Master Overview trước khi đối soát.")
            return
        if not self.weekly_po_nhung_file or not os.path.exists(self.weekly_po_nhung_file):
            messagebox.showwarning("Thiếu đơn hàng", "Vui lòng chọn file đơn hàng gửi đi (file của Nhung) trước khi đối soát.")
            return

        self.weekly_po_running = True
        self.btn_run_weekly_po.configure(state="disabled", text="Đang tiến hành đối soát...")
        self.prog_bar_weekly_po.set(0.15)

        for child in self.weekly_po_feed.winfo_children():
            child.destroy()

        loading_box = ctk.CTkFrame(self.weekly_po_feed, fg_color="transparent")
        loading_box.pack(fill="both", expand=True, pady=32)
        ctk.CTkLabel(
            loading_box,
            text="⏳ Đang phân tích Master Overview và đối soát các thông tin...",
            font=ctk.CTkFont(family=FONT_SANS, size=12),
            text_color=BROWN
        ).pack()

        threading.Thread(target=self._weekly_po_worker, daemon=True).start()

    def _weekly_po_worker(self):
        try:
            self.prog_bar_weekly_po.set(0.3)
            # Load master cache if not loaded
            if not self.weekly_po_master_cache:
                m_by_pi, m_all = weekly_load_master_overview_index(self.weekly_po_master_file)
                self.weekly_po_master_cache = (m_by_pi, m_all)
            else:
                m_by_pi, m_all = self.weekly_po_master_cache

            self.prog_bar_weekly_po.set(0.5)

            # Auto detect PI from files
            pi_detected = weekly_detect_pi_number(self.weekly_po_nhung_file)
            if not pi_detected and self.weekly_po_factory_file:
                pi_detected = weekly_detect_pi_number(self.weekly_po_factory_file)

            master_dict = m_all
            if pi_detected and pi_detected in m_by_pi:
                master_dict = m_by_pi[pi_detected]

            out_fac = None
            stats_fac = None
            fac_sum = None

            # 1. Audit Factory if provided
            if self.weekly_po_factory_file and os.path.exists(self.weekly_po_factory_file):
                self.prog_bar_weekly_po.set(0.65)
                out_fac, stats_fac, fac_sum = weekly_audit_file_generic(
                    self.weekly_po_factory_file,
                    master_dict,
                    output_suffix="_checked.xlsx",
                    role_name="NHAMAY"
                )

            # 2. Audit Nhung
            self.prog_bar_weekly_po.set(0.85)
            out_nhung, stats_nhung, nhung_sum = weekly_audit_file_generic(
                self.weekly_po_nhung_file,
                master_dict,
                output_suffix="_checked.xlsx",
                factory_items_sum=fac_sum,
                role_name="NHUNG"
            )

            self.prog_bar_weekly_po.set(1.0)
            self.after(0, lambda: self._weekly_po_success(pi_detected, out_nhung, stats_nhung, out_fac, stats_fac))

        except Exception as e:
            self.after(0, lambda: self._weekly_po_error(str(e)))

    def _weekly_po_success(self, pi_num, out_nhung, stats_nhung, out_fac, stats_fac):
        self.weekly_po_running = False
        self.btn_run_weekly_po.configure(state="normal", text="Bắt Đầu Đối Soát & Thêm Cột Check Vào File")

        for child in self.weekly_po_feed.winfo_children():
            child.destroy()

        res_box = ctk.CTkFrame(self.weekly_po_feed, fg_color="transparent")
        res_box.pack(fill="x", padx=16, pady=12)

        # Header summary
        header_text = f"✅ HOÀN THÀNH ĐỐI SOÁT" + (f" (Mã PI: {pi_num})" if pi_num else "")
        ctk.CTkLabel(
            res_box,
            text=header_text,
            font=ctk.CTkFont(family=FONT_SANS, size=13, weight="bold"),
            text_color=MOSS
        ).pack(anchor="w", pady=(0, 6))

        # Result Card Nhung
        nh_card = ctk.CTkFrame(res_box, fg_color=CARD, corner_radius=10, border_width=1, border_color=LINE)
        nh_card.pack(fill="x", pady=6)

        nh_top = ctk.CTkFrame(nh_card, fg_color="transparent")
        nh_top.pack(fill="x", padx=14, pady=8)

        fn_nh = os.path.basename(out_nhung)
        ctk.CTkLabel(
            nh_top,
            text=f"📁 File Đơn Hàng (Nhung): {fn_nh}",
            font=ctk.CTkFont(family=FONT_SANS, size=11, weight="bold"),
            text_color=ROAST
        ).pack(side="left")

        btn_box_nh = ctk.CTkFrame(nh_top, fg_color="transparent")
        btn_box_nh.pack(side="right")

        ctk.CTkButton(
            btn_box_nh, text="Mở File", font=ctk.CTkFont(family=FONT_SANS, size=11),
            fg_color=ROAST, hover_color=MOSS, text_color=IVORY,
            corner_radius=12, height=26, width=70, command=lambda: os.startfile(out_nhung)
        ).pack(side="left", padx=2)

        ctk.CTkButton(
            btn_box_nh, text="Thư Mục", font=ctk.CTkFont(family=FONT_SANS, size=11),
            fg_color="transparent", hover_color=SAND, text_color=ROAST,
            border_width=1, border_color=LINE, corner_radius=12, height=26, width=70,
            command=lambda: reveal_in_explorer(out_nhung)
        ).pack(side="left", padx=2)

        status_nh_str = f"Tổng {stats_nhung['total_lines']} dòng • Khớp: {stats_nhung['ok']} • Lệch/Chú ý: {stats_nhung['err'] + stats_nhung['warn']}"
        ctk.CTkLabel(
            nh_card, text=status_nh_str, font=ctk.CTkFont(family=FONT_SANS, size=10),
            text_color=BROWN if (stats_nhung['err'] + stats_nhung['warn'] == 0) else DANGER
        ).pack(anchor="w", padx=14, pady=(0, 8))

        # Result Card Factory (if ran)
        if out_fac and stats_fac:
            fac_card = ctk.CTkFrame(res_box, fg_color=CARD, corner_radius=10, border_width=1, border_color=LINE)
            fac_card.pack(fill="x", pady=6)

            fac_top = ctk.CTkFrame(fac_card, fg_color="transparent")
            fac_top.pack(fill="x", padx=14, pady=8)

            fn_fac = os.path.basename(out_fac)
            ctk.CTkLabel(
                fac_top,
                text=f"🏭 File Nhà Máy: {fn_fac}",
                font=ctk.CTkFont(family=FONT_SANS, size=11, weight="bold"),
                text_color=ROAST
            ).pack(side="left")

            btn_box_fac = ctk.CTkFrame(fac_top, fg_color="transparent")
            btn_box_fac.pack(side="right")

            ctk.CTkButton(
                btn_box_fac, text="Mở File", font=ctk.CTkFont(family=FONT_SANS, size=11),
                fg_color=ROAST, hover_color=MOSS, text_color=IVORY,
                corner_radius=12, height=26, width=70, command=lambda: os.startfile(out_fac)
            ).pack(side="left", padx=2)

            ctk.CTkButton(
                btn_box_fac, text="Thư Mục", font=ctk.CTkFont(family=FONT_SANS, size=11),
                fg_color="transparent", hover_color=SAND, text_color=ROAST,
                border_width=1, border_color=LINE, corner_radius=12, height=26, width=70,
                command=lambda: reveal_in_explorer(out_fac)
            ).pack(side="left", padx=2)

            miss_cnt = len(stats_fac['missing_in_file'])
            status_fac_str = f"Tổng {stats_fac['total_lines']} dòng • Khớp: {stats_fac['ok']} • Lệch dòng: {stats_fac['err']} • THIẾU BỎ SÓT: {miss_cnt} MÃ"
            ctk.CTkLabel(
                fac_card, text=status_fac_str, font=ctk.CTkFont(family=FONT_SANS, size=10),
                text_color=BROWN if (stats_fac['err'] == 0 and miss_cnt == 0) else DANGER
            ).pack(anchor="w", padx=14, pady=(0, 8))

            if miss_cnt > 0:
                alert_box = ctk.CTkFrame(res_box, fg_color="#fdeeed", corner_radius=8, border_width=1, border_color="#f5c2c7")
                alert_box.pack(fill="x", pady=6)
                ctk.CTkLabel(
                    alert_box,
                    text=f"⚠️ CẢNH BÁO: Nhà máy bị bỏ sót {miss_cnt} mã so với Master PO!",
                    font=ctk.CTkFont(family=FONT_SANS, size=11, weight="bold"),
                    text_color=DANGER
                ).pack(anchor="w", padx=12, pady=(6, 2))
                for m in stats_fac['missing_in_file']:
                    ctk.CTkLabel(
                        alert_box,
                        text=f"  • Mã {m['art_raw']} — {m['cust_name']} (Số lượng: {m['qty']})",
                        font=ctk.CTkFont(family=FONT_SANS, size=10),
                        text_color=ROAST
                    ).pack(anchor="w", padx=12, pady=1)
                ctk.CTkFrame(alert_box, height=4, fg_color="transparent").pack()

    def _weekly_po_error(self, err_msg):
        self.weekly_po_running = False
        self.btn_run_weekly_po.configure(state="normal", text="Bắt Đầu Đối Soát & Thêm Cột Check Vào File")
        self.prog_bar_weekly_po.set(0)

        for child in self.weekly_po_feed.winfo_children():
            child.destroy()

        err_box = ctk.CTkFrame(self.weekly_po_feed, fg_color="transparent")
        err_box.pack(fill="x", padx=16, pady=16)

        ctk.CTkLabel(
            err_box,
            text="❌ Có lỗi xảy ra trong quá trình đối soát:",
            font=ctk.CTkFont(family=FONT_SANS, size=12, weight="bold"),
            text_color=DANGER
        ).pack(anchor="w", pady=(0, 4))

        ctk.CTkLabel(
            err_box,
            text=err_msg,
            font=ctk.CTkFont(family=FONT_SANS, size=11),
            text_color=ROAST,
            wraplength=650,
            justify="left"
        ).pack(anchor="w")

if __name__ == "__main__":
    app = BoringTaskApp()
    app.mainloop()

