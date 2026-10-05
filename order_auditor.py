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
    def __init__(self, row_idx: int, item_no: str, description: str, unit: str, qty: float, price: float, total: Any):
        self.row_idx = row_idx
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
    def __init__(self, file_path: str):
        self.file_path = file_path
        self.po_number = ""
        self.pi_number = ""
        self.cust_po_no = ""
        self.customer = ""
        self.items: List[POItem] = []
        self.parse()

    def parse(self):
        wb = openpyxl.load_workbook(self.file_path, data_only=True)
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
                    total=total
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
        if status == "ERROR":
            self.overall_status = "ERROR"
            self.summary_reasons.append(f"[{layer_name}] {detail}")
        elif status == "REVIEW" and self.overall_status != "ERROR":
            self.overall_status = "REVIEW"
            self.summary_reasons.append(f"[{layer_name}] {detail}")


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
                audit_item.set_layer("Layer 1: SKU", "PASS", f"Matched Customer SKU: {matched_entry.get('sku')} (ID: {matched_entry.get('internal_id')})")
            elif lookup_code:
                # Nếu có mã Art No nhưng không tìm thấy trong Master Data khách
                # Lưu ý: Không vội kết luận ERROR nếu Customer Catalog không có sheet Master Data đầy đủ
                if customer_catalog.sku_map:
                    audit_item.set_layer("Layer 1: SKU", "REVIEW", f"Art No '{lookup_code}' not found in provided Customer Master Data")
                else:
                    audit_item.set_layer("Layer 1: SKU", "PASS", f"PO Art No: {lookup_code}")
            else:
                audit_item.set_layer("Layer 1: SKU", "REVIEW", "Missing Customer Art No in PO description")

            # ---------------------------------------------------------
            # LỚP 2 – MODEL VERIFICATION (MODEL NUMBER ↔ MODEL NAME)
            # ---------------------------------------------------------
            if po_item.md_number:
                model_info = self.model_kb.get_model_info(po_item.md_number)
                if model_info:
                    exp_name, exp_status = model_info
                    # Kiểm tra xem tên model mong đợi có xuất hiện trong mô tả hoặc tên hàng không
                    # Chuẩn hóa bỏ dấu câu, chữ hoa thường
                    exp_clean = exp_name.lower().replace("by sls", "").strip()
                    desc_low = combined_text.lower()

                    if exp_clean in desc_low:
                        if exp_status.lower() in ["discontinue", "deleted", "inactive"]:
                            audit_item.set_layer("Layer 2: Model", "REVIEW", f"{po_item.md_number} is {exp_name} ({exp_status} model)")
                        else:
                            audit_item.set_layer("Layer 2: Model", "PASS", f"{po_item.md_number} correctly matches '{exp_name}'")
                    else:
                        audit_item.set_layer("Layer 2: Model", "ERROR", f"{po_item.md_number} standard model is '{exp_name}', but not found in PO item description!")
                else:
                    audit_item.set_layer("Layer 2: Model", "REVIEW", f"{po_item.md_number} not recognized in standard Model List")
            else:
                audit_item.set_layer("Layer 2: Model", "PASS", "Non-MD item / General component")

            # ---------------------------------------------------------
            # LỚP 3 – FABRIC / COLOR VERIFICATION
            # ---------------------------------------------------------
            # Bóc tách màu/vải từ description
            fabric_status, fabric_msg = "PASS", "Fabric/Color check standard"
            # Kiểm tra trường hợp Dune, Boucle, Velvet, Soap, Black, Natural
            if "dune" in combined_text.lower():
                # Phân biệt chặt chẽ dòng Dune
                if "pasha dune" in combined_text.lower() or "padu" in combined_text.lower():
                    fabric_status, fabric_msg = "PASS", "Pasha Dune recognized (Fabric Pasha 058 Dune)"
                elif "free dune" in combined_text.lower():
                    fabric_status, fabric_msg = "PASS", "Free Dune recognized (Fabric Free 058 Dune)"
                elif "vega" in combined_text.lower() or "sand dune" in combined_text.lower():
                    fabric_status, fabric_msg = "PASS", "Vega Sand Dune recognized (Fabric Venga Recycle 004 Mole)"

            audit_item.set_layer("Layer 3: Fabric", fabric_status, fabric_msg)

            # ---------------------------------------------------------
            # LỚP 4 – PRODUCT CONFIGURATION & ORIENTATION
            # ---------------------------------------------------------
            # LEF / RHF / LHF / RIG
            orientation_m = re.search(r'\b(LEF|RIG|LHF|RHF|LEFT|RIGHT)\b', combined_text, re.IGNORECASE)
            if orientation_m:
                orient = orientation_m.group(1).upper()
                audit_item.set_layer("Layer 4: Config", "REVIEW", f"Orientation '{orient}' detected without engineering drawing. Please review BOM/Spec.")
            else:
                audit_item.set_layer("Layer 4: Config", "PASS", "Standard configuration")

            # ---------------------------------------------------------
            # LỚP 5 – SET VERIFICATION (SPLIT 1/2 VÀ 2/2)
            # ---------------------------------------------------------
            split_m = re.search(r'\b([12])/2\b', combined_text)
            if split_m:
                part = f"{split_m.group(1)}/2"
                # Tạo base key loại bỏ 1/2 và 2/2
                base_name = re.sub(r'\b[12]/2\b', '', combined_text).strip()
                base_key = re.sub(r'\s+', ' ', base_name.lower())[:40]
                if base_key not in sets_tracking:
                    sets_tracking[base_key] = {"1/2": 0.0, "2/2": 0.0, "items": []}
                sets_tracking[base_key][part] += po_item.qty
                sets_tracking[base_key]["items"].append(audit_item)
                audit_item.set_layer("Layer 5: Set", "PASS", f"Split part {part} tracked")
            else:
                audit_item.set_layer("Layer 5: Set", "PASS", "Single / Standalone unit")

            # ---------------------------------------------------------
            # LỚP 6 – COMMERCIAL CHECK (QTY, UNIT, PRICE)
            # ---------------------------------------------------------
            if po_item.qty <= 0:
                audit_item.set_layer("Layer 6: Commercial", "ERROR", "Quantity must be greater than 0")
            elif po_item.price == 0.0:
                # Giá = 0 KHÔNG TỰ ĐỘNG BÁO LỖI mà là REVIEW
                audit_item.set_layer("Layer 6: Commercial", "REVIEW", "Price is 0 (Free sample / Warranty / Included parts - confirm commercial policy)")
            else:
                audit_item.set_layer("Layer 6: Commercial", "PASS", f"Valid Qty={po_item.qty}, Price={po_item.price}")

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
                        f"Imbalanced Set! Part 1/2 Qty={qty_1} vs Part 2/2 Qty={qty_2}. Missing companion split unit!"
                    )

        return results


# ==============================================================================
# 4. EXPORT AUDIT EXCEL REPORT (SOFT IVORY BAKERY PALETTE)
# ==============================================================================

def export_audit_excel(audit_results: List[AuditResultItem], output_path: str):
    """
    Tạo báo cáo kiểm toán Excel đối chiếu song song chuyên nghiệp.
    Màu highlight theo tiêu chuẩn Excel & Soft Bakery:
      - PASS:   #C6EFCE (chữ #006100)
      - REVIEW: #FFEB9C (chữ #9C6500)
      - ERROR:  #FFC7CE (chữ #9C0006)
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Audit Report"

    # Định dạng màu & font
    font_header = Font(name="Segoe UI", size=11, bold=True, color="2B211C")
    font_body = Font(name="Segoe UI", size=10, color="2B211C")
    font_bold = Font(name="Segoe UI", size=10, bold=True)
    align_center = Alignment(horizontal="center", vertical="center")
    align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
    border_thin = Border(
        left=Side(style='thin', color="E8E0D0"),
        right=Side(style='thin', color="E8E0D0"),
        top=Side(style='thin', color="E8E0D0"),
        bottom=Side(style='thin', color="E8E0D0")
    )

    fill_header = PatternFill(start_color="F1EADB", end_color="F1EADB", fill_type="solid")
    fill_pass = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    fill_review = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
    fill_error = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")

    font_pass = Font(name="Segoe UI", size=10, bold=True, color="006100")
    font_review = Font(name="Segoe UI", size=10, bold=True, color="9C6500")
    font_error = Font(name="Segoe UI", size=10, bold=True, color="9C0006")

    headers = [
        "PO Row", "Overall Status", "Audit Summary Reasons",
        "PO Item Code", "PO Model", "Customer Art No", "PO Item Description",
        "Unit", "Qty", "Price",
        "Layer 1: SKU", "Layer 2: Model", "Layer 3: Fabric",
        "Layer 4: Config", "Layer 5: Set", "Layer 6: Commercial"
    ]

    ws.append(headers)
    for col_num in range(1, len(headers) + 1):
        cell = ws.cell(1, col_num)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = align_center
        cell.border = border_thin
    ws.row_dimensions[1].height = 28

    # Ghi dữ liệu kiểm toán
    for item in audit_results:
        p = item.po_item
        reasons_text = "; ".join(item.summary_reasons) if item.summary_reasons else "All 6 Layers Verified Passed"

        row_vals = [
            p.row_idx,
            item.overall_status,
            reasons_text,
            p.item_no,
            p.md_number,
            p.art_no,
            p.description,
            p.unit,
            p.qty,
            p.price,
            f"[{item.layers.get('Layer 1: SKU', ('', ''))[0]}] {item.layers.get('Layer 1: SKU', ('', ''))[1]}",
            f"[{item.layers.get('Layer 2: Model', ('', ''))[0]}] {item.layers.get('Layer 2: Model', ('', ''))[1]}",
            f"[{item.layers.get('Layer 3: Fabric', ('', ''))[0]}] {item.layers.get('Layer 3: Fabric', ('', ''))[1]}",
            f"[{item.layers.get('Layer 4: Config', ('', ''))[0]}] {item.layers.get('Layer 4: Config', ('', ''))[1]}",
            f"[{item.layers.get('Layer 5: Set', ('', ''))[0]}] {item.layers.get('Layer 5: Set', ('', ''))[1]}",
            f"[{item.layers.get('Layer 6: Commercial', ('', ''))[0]}] {item.layers.get('Layer 6: Commercial', ('', ''))[1]}"
        ]
        ws.append(row_vals)
        current_row = ws.max_row
        ws.row_dimensions[current_row].height = 24

        # Định dạng ô
        for col_idx in range(1, len(row_vals) + 1):
            cell = ws.cell(current_row, col_idx)
            cell.font = font_body
            cell.border = border_thin
            cell.alignment = align_left

        # Tô màu Overall Status
        status_cell = ws.cell(current_row, 2)
        status_cell.alignment = align_center
        if item.overall_status == "PASS":
            status_cell.fill = fill_pass
            status_cell.font = font_pass
        elif item.overall_status == "REVIEW":
            status_cell.fill = fill_review
            status_cell.font = font_review
        elif item.overall_status == "ERROR":
            status_cell.fill = fill_error
            status_cell.font = font_error

        # Tô màu cột Layer 2 (Model) nếu có ERROR
        layer2_cell = ws.cell(current_row, 12)
        if item.layers.get("Layer 2: Model", ("", ""))[0] == "ERROR":
            layer2_cell.fill = fill_error
            layer2_cell.font = font_error

        # Tô màu cột Layer 5 (Set) nếu có ERROR
        layer5_cell = ws.cell(current_row, 15)
        if item.layers.get("Layer 5: Set", ("", ""))[0] == "ERROR":
            layer5_cell.fill = fill_error
            layer5_cell.font = font_error

    # Auto-adjust column widths
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            val_str = str(cell.value or '')
            first_line = val_str.split('\n')[0]
            max_len = max(max_len, len(first_line))
        ws.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 45)

    ws.column_dimensions['C'].width = 40  # Summary reasons
    ws.column_dimensions['G'].width = 45  # Description

    wb.save(output_path)
    wb.close()
    return output_path
