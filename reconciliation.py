import pandas as pd
import numpy as np
import re
from difflib import SequenceMatcher
import io

def format_mst(val):
    """
    Format a tax code (MST) into a clean 10-digit string.
    Removes trailing decimals, spaces, and adds leading zeros if 9 digits.
    """
    if pd.isna(val) or str(val).strip().lower() in ["nan", "null", ""]:
        return ""
    val_str = str(val).strip()
    if val_str.endswith(".0"):
        val_str = val_str[:-2]
    # Keep only digits
    val_str = "".join([c for c in val_str if c.isdigit()])
    
    # Pad to 10 digits if it is 9 digits (common for leading zero drops)
    if len(val_str) == 9:
        val_str = "0" + val_str
    return val_str

def is_measurement(token):
    """
    Check if an alphanumeric token is a measurement dimension (like 160MM, 6INCH).
    These should not be treated as product model codes.
    """
    token = token.upper()
    units = ['MM', 'CM', 'M', 'KG', 'V', 'W', 'AH', 'INCH', 'PCS', 'SET', 'LIT', 'GR', 'G', 'INCHVA', 'VND']
    for unit in units:
        if token.endswith(unit):
            prefix = token[:-len(unit)]
            # If the part before unit is numeric, it is a measurement
            if prefix.isdigit() or prefix.replace('.', '', 1).replace(',', '', 1).isdigit():
                return True
    return False

def extract_model_codes(text):
    """
    Extract model numbers or unique alphanumeric codes from text (e.g. WFE7826, TCGLI2001).
    Excludes measurements like 160MM or 6INCH.
    """
    if not isinstance(text, str):
        return []
    tokens = re.split(r'[^a-zA-Z0-9]', text)
    codes = []
    for t in tokens:
        # Match tokens with both letters and numbers, length >= 3
        if len(t) >= 3 and any(c.isdigit() for c in t) and any(c.isalpha() for c in t):
            code_upper = t.upper()
            if not is_measurement(code_upper):
                codes.append(code_upper)
        # Or match pure numbers of length >= 4 (large model numbers)
        elif len(t) >= 4 and t.isdigit():
            codes.append(t)
    return codes

def extract_significant_words(text):
    """
    Extract significant unique keywords from a product name, ignoring common stop words.
    Used for fuzzy grouping of products without matching SKUs or model codes.
    """
    if not isinstance(text, str):
        return set()
    text = text.lower()
    # Split into words of length >= 4
    words = re.findall(r'\b\w{4,}\b', text)
    # Common Vietnamese and brand stop words
    stop_words = {
        'công', 'tnhh', 'thương', 'mại', 'dịch', 'vụ', 'viet', 'nam', 'hiệu', 'total', 
        'ingco', 'wadfow', 'chất', 'liệu', 'bằng', 'hàng', 'thiết', 'bị', 'phần', 'phiên',
        'bản', 'nhập', 'xuất', 'chính', 'sách', 'dụng', 'trong', 'dùng', 'nước', 'không',
        'bảo', 'đầu', 'bếp', 'giày', 'kèm', 'chân', 'xoay', 'đựng'
    }
    return set([w for w in words if w not in stop_words])

def share_words(name1, name2):
    """
    Check if two product names share at least one keyword of length >= 3.
    Provides a fallback when names are very different but contain key descriptors (like 'pin').
    """
    if not isinstance(name1, str) or not isinstance(name2, str):
        return False
    words1 = set(re.findall(r'\b\w{3,}\b', name1.lower()))
    words2 = set(re.findall(r'\b\w{3,}\b', name2.lower()))
    
    stop_words = {'cho', 'nha', 'của', 'với', 'cần', 'như', 'ghi', 'nhãn', 'thế'}
    w1 = words1 - stop_words
    w2 = words2 - stop_words
    
    return len(w1.intersection(w2)) > 0

def get_name_similarity(name1, name2):
    """
    Calculate text similarity ratio between two product names.
    """
    if not isinstance(name1, str) or not isinstance(name2, str):
        return 0.0
    return SequenceMatcher(None, name1.lower().strip(), name2.lower().strip()).ratio()

def clean_shoe_size(name):
    """
    Strip size suffixes at the end of product names (e.g. ' - 41', ' size 41', ' Size 45').
    Allows grouping Sapo items across sizes.
    """
    if not isinstance(name, str):
        return ""
    # Strip pattern " - 37" to " - 46" or " size 37" at the end of string
    name_clean = re.sub(r'\s*-\s*(3[5-9]|4[0-6])\s*$', '', name, flags=re.IGNORECASE)
    name_clean = re.sub(r'\s+size\s+(3[5-9]|4[0-6])\s*$', '', name_clean, flags=re.IGNORECASE)
    return name_clean.strip()

def clean_sku_size(sku):
    """
    Strip size suffixes at the end of SKUs (e.g. '-41', '-45').
    """
    if not isinstance(sku, str):
        return ""
    return re.sub(r'-(3[5-9]|4[0-6])$', '', sku).strip()

def load_excel_robust(file_input, key_column):
    """
    Load an Excel sheet robustly by looking for a row containing the key_column
    to use as the header. Supports both file paths and file-like objects.
    """
    if isinstance(file_input, bytes):
        df = pd.read_excel(io.BytesIO(file_input), sheet_name=0, header=None)
    else:
        df = pd.read_excel(file_input, sheet_name=0, header=None)
        
    header_row_idx = None
    for idx, row in df.iterrows():
        row_str_values = [str(val).strip() for val in row.values]
        if key_column in row_str_values:
            header_row_idx = idx
            break
            
    if header_row_idx is None:
        raise ValueError(f"Không tìm thấy dòng tiêu đề chứa cột '{key_column}'. Vui lòng kiểm tra lại định dạng file.")
        
    seen = {}
    cols = []
    for val in df.iloc[header_row_idx].values:
        val = str(val).strip()
        if val in seen:
            seen[val] += 1
            cols.append(f"{val}_{seen[val]}")
        else:
            seen[val] = 0
            cols.append(val)
            
    df.columns = cols
    df = df.iloc[header_row_idx + 1:].reset_index(drop=True)
    df = df.dropna(how='all')
    return df

def get_compared_values(s_item, i_item):
    """
    Get quantities and prices for comparison, applying UoM conversion rules.
    Specific rule: CSPS Wall Mesh Hook (VNDTCGO10BZ1) sets of 9 vs individual units (cái).
    """
    s_qty = s_item["qty"]
    s_price = s_item["unit_price_excl"]
    i_qty = i_item["qty"]
    i_price = i_item["unit_price"]
    uom_applied = False
    
    is_s_csps = "VNDTCGO10BZ1" in s_item["sku"].upper() or "VNDTCGO10BZ1" in s_item["name"].upper()
    is_i_csps = "VNDTCGO10BZ1" in i_item["name"].upper()
    
    if is_s_csps and is_i_csps:
        # If Sapo has Sets (Qty 100) and Invoice has Individuals (Qty 900)
        if s_qty * 9 == i_qty:
            s_qty = s_qty * 9
            s_price = s_price / 9
            uom_applied = True
        # If Sapo has Individuals (Qty 900) and Invoice has Sets (Qty 100)
        elif i_qty * 9 == s_qty:
            i_qty = i_qty * 9
            i_price = i_price / 9
            uom_applied = True
            
    return s_qty, s_price, i_qty, i_price, uom_applied

def reconcile_data(sapo_file, supplier_file, invoice_file):
    """
    Perform reconciliation of Sapo imports and Tax Invoices.
    """
    # 1. Load supplier list
    if isinstance(supplier_file, bytes):
        df_sup = pd.read_excel(io.BytesIO(supplier_file), sheet_name=0)
    else:
        df_sup = pd.read_excel(supplier_file, sheet_name=0)
        
    df_sup["Mã nhà cung cấp"] = df_sup["Mã nhà cung cấp"].astype(str).str.strip()
    df_sup["Tên nhà cung cấp *"] = df_sup["Tên nhà cung cấp *"].astype(str).str.strip()
    df_sup["Mã số thuế_clean"] = df_sup["Mã số thuế"].apply(format_mst)
    sup_mst_map = dict(zip(df_sup["Mã nhà cung cấp"], df_sup["Mã số thuế_clean"]))
    sup_name_map = dict(zip(df_sup["Mã số thuế_clean"], df_sup["Tên nhà cung cấp *"]))
    
    # 2. Load Sapo Imports
    df_sapo = load_excel_robust(sapo_file, "Mã đơn nhập hàng")
    df_sapo = df_sapo[df_sapo["Mã đơn nhập hàng"].notna()].copy()
    df_sapo["Mã nhà cung cấp"] = df_sapo["Mã nhà cung cấp"].astype(str).str.strip()
    df_sapo["Tên nhà cung cấp"] = df_sapo["Tên nhà cung cấp"].astype(str).str.strip()
    df_sapo["Mã đơn nhập hàng"] = df_sapo["Mã đơn nhập hàng"].astype(str).str.strip()
    df_sapo["Ngày nhập_dt"] = pd.to_datetime(df_sapo["Ngày nhập"], format="%d/%m/%Y %H:%M")
    df_sapo["MST_Supplier"] = df_sapo["Mã nhà cung cấp"].map(sup_mst_map).fillna("")
    
    # Apply clean shoe sizes
    df_sapo["Mã SKU"] = df_sapo["Mã SKU"].apply(clean_sku_size)
    df_sapo["Tên phiên bản sản phẩm"] = df_sapo["Tên phiên bản sản phẩm"].apply(clean_shoe_size)
    
    # Cast Sapo quantities and prices to numeric
    df_sapo["SL nhập"] = pd.to_numeric(df_sapo["SL nhập"], errors='coerce').fillna(0).astype(int)
    df_sapo["Thành tiền"] = pd.to_numeric(df_sapo["Thành tiền"], errors='coerce').fillna(0.0)
    df_sapo["Thuế"] = pd.to_numeric(df_sapo["Thuế"], errors='coerce').fillna(0.0)
    
    # Group Sapo items by order and size-stripped product details
    group_cols = [
        "Mã đơn nhập hàng", "Ngày nhập", "Mã nhà cung cấp", "Tên nhà cung cấp", 
        "MST_Supplier", "Mã SKU", "Tên phiên bản sản phẩm", "Áp dụng thuế", 
        "Trạng thái", "Trạng thái nhập", "Ngày nhập_dt"
    ]
    df_sapo = df_sapo.groupby(group_cols, as_index=False).agg({
        "SL nhập": "sum",
        "Thành tiền": "sum",
        "Thuế": "sum"
    })
    
    # 3. Load Invoices
    df_inv = load_excel_robust(invoice_file, "Số hóa đơn")
    df_inv = df_inv[(df_inv["Số hóa đơn"].notna()) & (df_inv["Số hóa đơn"].astype(str).str.strip() != "[3]")].copy()
    df_inv["Số hóa đơn"] = df_inv["Số hóa đơn"].astype(str).str.strip()
    df_inv["Mã số thuế"] = df_inv["Mã số thuế"].apply(format_mst)
    df_inv["Ngày hóa đơn_dt"] = pd.to_datetime(df_inv["Ngày hóa đơn"])
    df_inv["Tên người bán/Mua"] = df_inv["Tên người bán/Mua"].astype(str).str.strip()
    df_inv["Tên hàng"] = df_inv["Tên hàng"].astype(str).str.strip()
    
    # Store results
    order_summaries = []
    item_details = []
    unmatched_invoice_items = []
    invoice_row_matches = {}
    
    # Reconcile Supplier by Supplier
    sapo_by_supplier = df_sapo.groupby("MST_Supplier")
    
    for mst, sapo_group in sapo_by_supplier:
        sapo_items = []
        for idx, row in sapo_group.sort_values(by="Ngày nhập_dt").iterrows():
            qty = int(row["SL nhập"])
            total_incl = float(row["Thành tiền"])
            tax_val = float(row["Thuế"])
            total_excl = total_incl - tax_val
            unit_price_excl = total_excl / qty if qty > 0 else 0.0
            
            s_name = str(row["Tên phiên bản sản phẩm"])
            s_sku = str(row["Mã SKU"])
            
            sapo_items.append({
                "row_idx": idx,
                "order_id": str(row["Mã đơn nhập hàng"]),
                "sku": s_sku,
                "name": s_name,
                "qty": qty,
                "unit_price_excl": unit_price_excl,
                "total_excl": total_excl,
                "tax": tax_val,
                "total_incl": total_incl,
                "date": row["Ngày nhập_dt"],
                "matched": False,
                "match_info": None,
                "uom_conversion": 1,
                "status": "Chưa xuất hóa đơn",
                "codes": set(extract_model_codes(s_name) + extract_model_codes(s_sku)),
                "words": extract_significant_words(s_name),
                "words3": set(re.findall(r'\b\w{3,}\b', s_name.lower())) - {'cho', 'nha', 'của', 'với', 'cần', 'như', 'ghi', 'nhãn', 'thế'}
            })
            
        if mst == "":
            for order_id in sapo_group["Mã đơn nhập hàng"].unique():
                order_rows = sapo_group[sapo_group["Mã đơn nhập hàng"] == order_id]
                order_summaries.append({
                    "order_id": order_id,
                    "date": order_rows["Ngày nhập_dt"].iloc[0].strftime("%d/%m/%Y"),
                    "supplier_code": order_rows["Mã nhà cung cấp"].iloc[0],
                    "supplier_name": order_rows["Tên nhà cung cấp"].iloc[0],
                    "mst": "",
                    "total_incl": float(order_rows["Thành tiền"].sum()),
                    "total_excl": float((order_rows["Thành tiền"] - order_rows["Thuế"]).sum()),
                    "status": "Thiếu MST",
                    "notes": "Nhà cung cấp chưa được khai báo Mã số thuế (MST) trên Sapo."
                })
            for s_item in sapo_items:
                item_details.append({
                    "order_id": s_item["order_id"],
                    "sku": s_item["sku"],
                    "name": s_item["name"],
                    "qty": s_item["qty"],
                    "price_excl": s_item["unit_price_excl"],
                    "total_excl": s_item["total_excl"],
                    "invoice_no": "",
                    "inv_name": "",
                    "inv_qty": 0,
                    "inv_price": 0.0,
                    "inv_total": 0.0,
                    "status": "Thiếu MST",
                    "diff_qty": -s_item["qty"],
                    "diff_amt": -s_item["total_excl"]
                })
            continue
            
        invoices_group = df_inv[df_inv["Mã số thuế"] == mst].sort_values(by="Ngày hóa đơn_dt")
        
        if invoices_group.empty:
            for order_id in sapo_group["Mã đơn nhập hàng"].unique():
                order_rows = sapo_group[sapo_group["Mã đơn nhập hàng"] == order_id]
                order_summaries.append({
                    "order_id": order_id,
                    "date": order_rows["Ngày nhập_dt"].iloc[0].strftime("%d/%m/%Y"),
                    "supplier_code": order_rows["Mã nhà cung cấp"].iloc[0],
                    "supplier_name": order_rows["Tên nhà cung cấp"].iloc[0],
                    "mst": mst,
                    "total_incl": float(order_rows["Thành tiền"].sum()),
                    "total_excl": float((order_rows["Thành tiền"] - order_rows["Thuế"]).sum()),
                    "status": "Chưa xuất hóa đơn",
                    "notes": f"Không tìm thấy hóa đơn nào của nhà cung cấp có MST {mst}."
                })
            for s_item in sapo_items:
                item_details.append({
                    "order_id": s_item["order_id"],
                    "sku": s_item["sku"],
                    "name": s_item["name"],
                    "qty": s_item["qty"],
                    "price_excl": s_item["unit_price_excl"],
                    "total_excl": s_item["total_excl"],
                    "invoice_no": "",
                    "inv_name": "",
                    "inv_qty": 0,
                    "inv_price": 0.0,
                    "inv_total": 0.0,
                    "status": "Chưa xuất hóa đơn",
                    "diff_qty": -s_item["qty"],
                    "diff_amt": -s_item["total_excl"]
                })
            continue
            
        inv_items = []
        stop_words_3 = {'cho', 'nha', 'của', 'với', 'cần', 'như', 'ghi', 'nhãn', 'thế'}
        for idx, row in invoices_group.iterrows():
            qty = int(row["Số lượng"])
            total_excl = float(row["Thành tiền"])
            unit_price = float(row["Đơn giá"])
            i_name = str(row["Tên hàng"])
            
            inv_items.append({
                "row_idx": idx,
                "invoice_id": str(row["Số hóa đơn"]),
                "name": i_name,
                "qty": qty,
                "unit_price": unit_price,
                "total_excl": total_excl,
                "date": row["Ngày hóa đơn_dt"],
                "matched": False,
                "codes": set(extract_model_codes(i_name)),
                "words": extract_significant_words(i_name),
                "words3": set(re.findall(r'\b\w{3,}\b', i_name.lower())) - stop_words_3
            })
            
        def record_match(s_item, i_item, status_label, uom_applied=False, custom_qty=None, custom_price=None):
            s_item["matched"] = True
            if custom_qty is not None:
                match_view = dict(i_item)
                match_view["qty"] = custom_qty
                match_view["unit_price"] = custom_price if custom_price is not None else i_item["unit_price"]
                match_view["total_excl"] = custom_qty * match_view["unit_price"]
                s_item["match_info"] = match_view
            else:
                s_item["match_info"] = i_item
                
            s_item["status"] = status_label
            if uom_applied:
                s_item["uom_conversion"] = 9
            i_item["matched"] = True
            invoice_row_matches[i_item["row_idx"]] = {
                "sku": s_item["sku"],
                "order_id": s_item["order_id"]
            }
        
        # ==============================================================
        # PHASE 1: ORDER-TO-INVOICE MATCHING (Ưu tiên khớp trọn vẹn cấp Đơn hàng)
        # ==============================================================
        # Tránh lỗi FIFO: Khi có nhiều đơn hàng nhưng hóa đơn xuất cho đơn sau,
        # nếu đơn sau và hóa đơn khớp giá trị/mặt hàng, ta ghép cặp chúng trước.
        orders_dict = {}
        for s in sapo_items:
            orders_dict.setdefault(s["order_id"], []).append(s)
            
        invoices_dict = {}
        for i in inv_items:
            invoices_dict.setdefault(i["invoice_id"], []).append(i)
            
        # Tìm các cặp Đơn hàng - Hóa đơn khớp tổng tiền (< 500đ hoặc lệch thuế < 0.5%)
        amount_pairs = []
        for oid, o_item_list in orders_dict.items():
            o_tot = sum(s["total_excl"] for s in o_item_list)
            o_date = o_item_list[0]["date"]
            for inum, i_item_list in invoices_dict.items():
                i_tot = sum(i["total_excl"] for i in i_item_list)
                i_date = i_item_list[0]["date"]
                diff = abs(o_tot - i_tot)
                if diff < 500 or (diff / i_tot < 0.005 if i_tot > 0 else False):
                    days_diff = abs((o_date - i_date).days)
                    amount_pairs.append({
                        "order_id": oid,
                        "invoice_id": inum,
                        "diff": diff,
                        "days_diff": days_diff,
                        "o_tot": o_tot,
                        "i_tot": i_tot
                    })
                    
        # Ưu tiên cặp có độ lệch tiền nhỏ nhất, sau đó đến ngày gần nhất
        amount_pairs.sort(key=lambda x: (x["diff"], x["days_diff"]))
        
        matched_orders_phase1 = set()
        matched_invoices_phase1 = set()
        
        for pair in amount_pairs:
            oid = pair["order_id"]
            inum = pair["invoice_id"]
            if oid in matched_orders_phase1 or inum in matched_invoices_phase1:
                continue
                
            o_item_list = orders_dict[oid]
            i_item_list = invoices_dict[inum]
            
            # Bước 1.1: Khớp chính xác mã model và số lượng trong cặp này
            for s_item in o_item_list:
                if s_item["matched"]: continue
                s_codes = s_item["codes"]
                for i_item in i_item_list:
                    if i_item["matched"]: continue
                    s_qty_c, s_price_c, i_qty_c, i_price_c, uom = get_compared_values(s_item, i_item)
                    if s_qty_c != i_qty_c: continue
                    if s_codes and s_codes.intersection(i_item["codes"]):
                        if abs(s_price_c - i_price_c) < 500 or (abs(s_price_c - i_price_c) / i_price_c < 0.05 if i_price_c > 0 else False):
                            record_match(s_item, i_item, "Khớp", uom)
                            break

            # Bước 1.2: Khớp tên tương tự / từ khóa trong cặp này
            for s_item in o_item_list:
                if s_item["matched"]: continue
                for i_item in i_item_list:
                    if i_item["matched"]: continue
                    s_qty_c, s_price_c, i_qty_c, i_price_c, uom = get_compared_values(s_item, i_item)
                    if s_qty_c != i_qty_c: continue
                    is_match = (
                        (s_item["words"] and s_item["words"].intersection(i_item["words"])) or
                        (s_item["words3"] and s_item["words3"].intersection(i_item["words3"])) or
                        get_name_similarity(s_item["name"], i_item["name"]) >= 0.55
                    )
                    if is_match:
                        price_diff = abs(s_price_c - i_price_c)
                        if price_diff < 500 or (price_diff / i_price_c < 0.05 if i_price_c > 0 else False):
                            record_match(s_item, i_item, "Khớp", uom)
                            break
                            
            # Bước 1.3: Xử lý hóa đơn gộp quy cách (nhiều dòng size/mã phụ gộp thành 1 dòng HĐ)
            unmatched_s = [s for s in o_item_list if not s["matched"]]
            unmatched_i = [i for i in i_item_list if not i["matched"]]
            
            for i_item in list(unmatched_i):
                candidates = [
                    s for s in unmatched_s 
                    if abs(s["unit_price_excl"] - i_item["unit_price"]) < 500 or 
                       (s["codes"] and s["codes"].intersection(i_item["codes"])) or
                       (s["words3"] and s["words3"].intersection(i_item["words3"]))
                ]
                if sum(s["qty"] for s in candidates) == i_item["qty"]:
                    for s_item in candidates:
                        record_match(s_item, i_item, "Khớp", custom_qty=s_item["qty"], custom_price=i_item["unit_price"])
                        unmatched_s.remove(s_item)
                    unmatched_i.remove(i_item)
                    
            if len(unmatched_i) == 1 and len(unmatched_s) > 0:
                i_item = unmatched_i[0]
                if sum(s["qty"] for s in unmatched_s) == i_item["qty"]:
                    for s_item in unmatched_s:
                        record_match(s_item, i_item, "Khớp", custom_qty=s_item["qty"], custom_price=i_item["unit_price"])
                    unmatched_s.clear()
                    unmatched_i.clear()
                    
            matched_orders_phase1.add(oid)
            matched_invoices_phase1.add(inum)

        # ==============================================================
        # PHASE 2: FIFO RECONCILIATION CHO CÁC MẶT HÀNG CÒN LẠI
        # ==============================================================
        # Pass 1: Perfect Match (Same model, same qty, close price)
        for s_item in sapo_items:
            s_codes = s_item["codes"]
            if not s_codes:
                continue
            candidates = sorted(
                [i for i in inv_items if not i["matched"]],
                key=lambda x: abs((x["date"] - s_item["date"]).days)
            )
            for i_item in candidates:
                s_qty_c, s_price_c, i_qty_c, i_price_c, uom_applied = get_compared_values(s_item, i_item)
                if s_qty_c != i_qty_c:
                    continue
                if s_codes.intersection(i_item["codes"]):
                    price_diff = abs(s_price_c - i_price_c)
                    price_ratio = price_diff / i_price_c if i_price_c > 0 else 0
                    if price_ratio < 0.05 or price_diff < 500:
                        record_match(s_item, i_item, "Khớp", uom_applied)
                        break
                        
        # Pass 2: Fuzzy Name Match (Same qty, similar names, close price)
        for s_item in sapo_items:
            if s_item["matched"]:
                continue
            candidates = sorted(
                [i for i in inv_items if not i["matched"]],
                key=lambda x: abs((x["date"] - s_item["date"]).days)
            )
            for i_item in candidates:
                s_qty_c, s_price_c, i_qty_c, i_price_c, uom_applied = get_compared_values(s_item, i_item)
                if s_qty_c != i_qty_c:
                    continue
                similarity = get_name_similarity(s_item["name"], i_item["name"])
                if similarity >= 0.55:
                    price_diff = abs(s_price_c - i_price_c)
                    price_ratio = price_diff / i_price_c if i_price_c > 0 else 0
                    if price_ratio < 0.05 or price_diff < 500:
                        record_match(s_item, i_item, "Khớp", uom_applied)
                        break

        # Pass 2b: Significant Word Overlap (Same qty, significant word overlap, close price)
        for s_item in sapo_items:
            if s_item["matched"]:
                continue
            s_words = s_item["words"]
            if not s_words:
                continue
            candidates = sorted(
                [i for i in inv_items if not i["matched"]],
                key=lambda x: abs((x["date"] - s_item["date"]).days)
            )
            for i_item in candidates:
                s_qty_c, s_price_c, i_qty_c, i_price_c, uom_applied = get_compared_values(s_item, i_item)
                if s_qty_c != i_qty_c:
                    continue
                if s_words.intersection(i_item["words"]):
                    price_diff = abs(s_price_c - i_price_c)
                    price_ratio = price_diff / i_price_c if i_price_c > 0 else 0
                    if price_ratio < 0.05 or price_diff < 500:
                        record_match(s_item, i_item, "Khớp", uom_applied)
                        break

        # Pass 2c: Share Words >= 3 + Qty + Price Match (Fallback for items with different names but match in qty/price)
        for s_item in sapo_items:
            if s_item["matched"]:
                continue
            s_words3 = s_item["words3"]
            if not s_words3:
                continue
            candidates = sorted(
                [i for i in inv_items if not i["matched"]],
                key=lambda x: abs((x["date"] - s_item["date"]).days)
            )
            for i_item in candidates:
                s_qty_c, s_price_c, i_qty_c, i_price_c, uom_applied = get_compared_values(s_item, i_item)
                if s_qty_c != i_qty_c:
                    continue
                if s_words3.intersection(i_item["words3"]):
                    price_diff = abs(s_price_c - i_price_c)
                    price_ratio = price_diff / i_price_c if i_price_c > 0 else 0
                    if price_ratio < 0.05 or price_diff < 500:
                        record_match(s_item, i_item, "Khớp", uom_applied)
                        break

        # Pass 3: Price Discrepancy Match (Same model / similar name / word overlap, same qty, price differs > 5%)
        for s_item in sapo_items:
            if s_item["matched"]:
                continue
            s_codes = s_item["codes"]
            s_words = s_item["words"]
            s_words3 = s_item["words3"]
            candidates = sorted(
                [i for i in inv_items if not i["matched"]],
                key=lambda x: abs((x["date"] - s_item["date"]).days)
            )
            for i_item in candidates:
                s_qty_c, s_price_c, i_qty_c, i_price_c, uom_applied = get_compared_values(s_item, i_item)
                if s_qty_c != i_qty_c:
                    continue
                is_candidate = (
                    bool(s_codes and s_codes.intersection(i_item["codes"])) or
                    bool(s_words and s_words.intersection(i_item["words"])) or
                    bool(s_words3 and s_words3.intersection(i_item["words3"])) or
                    (get_name_similarity(s_item["name"], i_item["name"]) >= 0.55)
                )
                if is_candidate:
                    record_match(s_item, i_item, "Lệch đơn giá", uom_applied)
                    break
                    
        # Pass 4: Quantity Discrepancy Match (Same model / similar name / word overlap, close price, but qty differs)
        for s_item in sapo_items:
            if s_item["matched"]:
                continue
            s_codes = s_item["codes"]
            s_words = s_item["words"]
            s_words3 = s_item["words3"]
            candidates = sorted(
                [i for i in inv_items if not i["matched"]],
                key=lambda x: abs((x["date"] - s_item["date"]).days)
            )
            for i_item in candidates:
                s_qty_c, s_price_c, i_qty_c, i_price_c, uom_applied = get_compared_values(s_item, i_item)
                is_candidate = (
                    bool(s_codes and s_codes.intersection(i_item["codes"])) or
                    bool(s_words and s_words.intersection(i_item["words"])) or
                    bool(s_words3 and s_words3.intersection(i_item["words3"])) or
                    (get_name_similarity(s_item["name"], i_item["name"]) >= 0.55)
                )
                if is_candidate:
                    price_diff = abs(s_price_c - i_price_c)
                    price_ratio = price_diff / i_price_c if i_price_c > 0 else 0
                    if price_ratio < 0.05 or price_diff < 500:
                        record_match(s_item, i_item, "Lệch số lượng", uom_applied)
                        break

        # Save detailed reconciliation rows
        for s_item in sapo_items:
            if s_item["matched"]:
                i_item = s_item["match_info"]
                s_qty_show = s_item["qty"]
                s_price_show = s_item["unit_price_excl"]
                i_qty_show = i_item["qty"]
                i_price_show = i_item["unit_price"]
                
                if s_item["uom_conversion"] == 9:
                    s_qty_show = s_qty_show * 9
                    s_price_show = s_price_show / 9
                    
                qty_diff = i_qty_show - s_qty_show
                price_diff = i_price_show - s_price_show
                
                status = "Khớp"
                if s_item["uom_conversion"] == 9:
                    status = "Khớp (Quy đổi bộ x9)"
                    
                if qty_diff != 0 and abs(price_diff) >= 5.0:
                    status = "Lệch số lượng & đơn giá"
                elif qty_diff != 0:
                    status = "Lệch số lượng"
                elif abs(price_diff) >= 5.0:
                    status = "Lệch đơn giá"
                    
                item_details.append({
                    "order_id": s_item["order_id"],
                    "sku": s_item["sku"],
                    "name": s_item["name"] + (" (Quy đổi x9)" if s_item["uom_conversion"] == 9 else ""),
                    "qty": s_qty_show,
                    "price_excl": s_price_show,
                    "total_excl": s_item["total_excl"],
                    "invoice_no": i_item["invoice_id"],
                    "inv_name": i_item["name"],
                    "inv_qty": i_item["qty"],
                    "inv_price": i_item["unit_price"],
                    "inv_total": i_item["total_excl"],
                    "status": status,
                    "diff_qty": qty_diff,
                    "diff_amt": i_item["total_excl"] - s_item["total_excl"]
                })
            else:
                item_details.append({
                    "order_id": s_item["order_id"],
                    "sku": s_item["sku"],
                    "name": s_item["name"],
                    "qty": s_item["qty"],
                    "price_excl": s_item["unit_price_excl"],
                    "total_excl": s_item["total_excl"],
                    "invoice_no": "",
                    "inv_name": "",
                    "inv_qty": 0,
                    "inv_price": 0.0,
                    "inv_total": 0.0,
                    "status": "Chưa xuất hóa đơn",
                    "diff_qty": -s_item["qty"],
                    "diff_amt": -s_item["total_excl"]
                })
                
        # Record unmatched invoice items (surplus invoices)
        for i_item in inv_items:
            if not i_item["matched"]:
                unmatched_invoice_items.append({
                    "mst": mst,
                    "supplier_name": sup_name_map.get(mst, "Không rõ"),
                    "date": i_item["date"].strftime("%d/%m/%Y"),
                    "invoice_no": i_item["invoice_id"],
                    "name": i_item["name"],
                    "qty": i_item["qty"],
                    "price": i_item["unit_price"],
                    "total": i_item["total_excl"]
                })
                
        # Generate Sapo order summaries
        for order_id in sapo_group["Mã đơn nhập hàng"].unique():
            order_rows = [it for it in item_details if it["order_id"] == order_id]
            order_items_status = [it["status"] for it in order_rows]
            
            if all("Khớp" in stat for stat in order_items_status):
                status = "Khớp"
                inv_nos = list(set(str(it["invoice_no"]).strip() for it in order_rows if str(it["invoice_no"]).strip()))
                if inv_nos:
                    notes = f"Khớp hoàn toàn với HĐ số: {', '.join(inv_nos)}."
                else:
                    notes = "Hóa đơn và số lượng khớp hoàn toàn."
            elif all(stat == "Chưa xuất hóa đơn" for stat in order_items_status):
                status = "Chưa xuất hóa đơn"
                notes = "Toàn bộ sản phẩm trong đơn chưa có hóa đơn."
            else:
                status = "Lệch"
                mismatch_reasons = []
                unmatched_count = sum(1 for stat in order_items_status if stat == "Chưa xuất hóa đơn")
                qty_diff_count = sum(1 for stat in order_items_status if "Lệch số lượng" in stat)
                price_diff_count = sum(1 for stat in order_items_status if "Lệch đơn giá" in stat)
                
                if unmatched_count > 0:
                    mismatch_reasons.append(f"{unmatched_count} mặt hàng chưa HĐ")
                if qty_diff_count > 0:
                    mismatch_reasons.append(f"{qty_diff_count} mặt hàng lệch SL")
                if price_diff_count > 0:
                    mismatch_reasons.append(f"{price_diff_count} mặt hàng lệch đơn giá")
                notes = ", ".join(mismatch_reasons)
                
            group_order = sapo_group[sapo_group["Mã đơn nhập hàng"] == order_id]
            order_summaries.append({
                "order_id": order_id,
                "date": group_order["Ngày nhập_dt"].iloc[0].strftime("%d/%m/%Y"),
                "supplier_code": group_order["Mã nhà cung cấp"].iloc[0],
                "supplier_name": group_order["Tên nhà cung cấp"].iloc[0],
                "mst": mst,
                "total_incl": float(group_order["Thành tiền"].sum()),
                "total_excl": float((group_order["Thành tiền"] - group_order["Thuế"]).sum()),
                "status": status,
                "notes": notes
            })

    # Return results as DataFrames
    return (
        pd.DataFrame(order_summaries, columns=["order_id", "date", "supplier_code", "supplier_name", "mst", "total_incl", "total_excl", "status", "notes"]) if order_summaries else pd.DataFrame(columns=["order_id", "date", "supplier_code", "supplier_name", "mst", "total_incl", "total_excl", "status", "notes"]),
        pd.DataFrame(item_details, columns=["order_id", "sku", "name", "qty", "price_excl", "total_excl", "invoice_no", "inv_name", "inv_qty", "inv_price", "inv_total", "status", "diff_qty", "diff_amt"]) if item_details else pd.DataFrame(columns=["order_id", "sku", "name", "qty", "price_excl", "total_excl", "invoice_no", "inv_name", "inv_qty", "inv_price", "inv_total", "status", "diff_qty", "diff_amt"]),
        pd.DataFrame(unmatched_invoice_items, columns=["mst", "supplier_name", "date", "invoice_no", "name", "qty", "price", "total"]) if unmatched_invoice_items else pd.DataFrame(columns=["mst", "supplier_name", "date", "invoice_no", "name", "qty", "price", "total"]),
        invoice_row_matches
    )
