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
