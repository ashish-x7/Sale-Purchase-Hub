import os
import math
import json
import traceback
import pickle
from datetime import datetime
from flask import Flask, render_template, request, jsonify, send_file
import xlrd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import pandas as pd

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024  # 100MB max upload

# ─── Persistent Data Storage ────────────────────────────────────────────────
# Store data in the workspace folder so it persists between restarts
WORKSPACE = os.path.dirname(os.path.abspath(__file__))
UPLOAD_FOLDER = os.path.join(WORKSPACE, 'uploads')
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
DATA_STORE_PATH = os.path.join(WORKSPACE, 'processed_data_store.pkl')


def get_financial_year():
    """Get current financial year string like '2026-27'."""
    now = datetime.now()
    if now.month >= 4:  # April onwards = new FY
        return f'{now.year}-{str(now.year + 1)[2:]}'
    else:
        return f'{now.year - 1}-{str(now.year)[2:]}'


def load_stored_data():
    """Load previously stored data from disk."""
    if os.path.exists(DATA_STORE_PATH):
        try:
            with open(DATA_STORE_PATH, 'rb') as f:
                return normalize_stored_data(pickle.load(f))
        except Exception:
            pass
    return {'sale': [], 'purchase': [], 'sale_ids': set(), 'purchase_ids': set()}


def save_stored_data(data):
    """Save data to disk for persistence."""
    with open(DATA_STORE_PATH, 'wb') as f:
        pickle.dump(normalize_stored_data(data), f)


def clean_json_value(value):
    """Return a value that can be safely serialized as strict JSON."""
    if isinstance(value, float):
        return value if math.isfinite(value) else 0
    if isinstance(value, dict):
        return {k: clean_json_value(v) for k, v in value.items()}
    if isinstance(value, list):
        return [clean_json_value(v) for v in value]
    return value


def clean_rows(rows):
    return [clean_json_value(dict(row)) for row in rows]


def normalize_stored_data(data):
    """Normalize old/new pickle data into the current strict JSON-safe shape."""
    if not isinstance(data, dict):
        return {'sale': [], 'purchase': [], 'sale_ids': set(), 'purchase_ids': set()}

    sale = clean_rows(data.get('sale', []))
    purchase = clean_rows(data.get('purchase', []))
    sale_ids = set(data.get('sale_ids') or [])
    purchase_ids = set(data.get('purchase_ids') or [])

    if not sale_ids:
        sale_ids = {row.get('sale_unique_id') for row in sale if row.get('sale_unique_id')}
    if not purchase_ids:
        purchase_ids = {row.get('purchase_unique_id') for row in purchase if row.get('purchase_unique_id')}

    return {
        'sale': sale,
        'purchase': purchase,
        'sale_ids': sale_ids,
        'purchase_ids': purchase_ids,
    }


def merge_with_dedup(existing_data, new_data, existing_ids, unique_id_key):
    """Merge new data into existing, skipping duplicates by UNIQUE ID.
    
    Returns: (merged_data, merged_ids, new_count, duplicate_count)
    """
    new_count = 0
    dup_count = 0
    merged = list(existing_data)
    merged_ids = set(existing_ids)
    
    for row in new_data:
        uid = row.get(unique_id_key, '')
        if uid and uid in merged_ids:
            dup_count += 1
        else:
            merged.append(row)
            if uid:
                merged_ids.add(uid)
            new_count += 1
    
    # Re-number all rows
    for i, row in enumerate(merged):
        row['no'] = i + 1
    
    return merged, merged_ids, new_count, dup_count

# ─── Helpers ──────────────────────────────────────────────────────────

def roundup2(value):
    """Round up to 2 decimal places (like Excel ROUNDUP)."""
    if value is None or value == '':
        return 0
    try:
        v = float(value)
        if not math.isfinite(v):
            return 0
        return math.ceil(v * 100) / 100
    except (ValueError, TypeError):
        return 0


def safe_float(value, default=0):
    """Safely convert to float."""
    if value is None or value == '':
        return default
    try:
        v = float(value)
        return v if math.isfinite(v) else default
    except (ValueError, TypeError):
        return default


def safe_str(value, default=''):
    """Safely convert to string."""
    if value is None:
        return default
    s = str(value).strip()
    # Remove trailing .0 from numeric strings
    if s.endswith('.0'):
        try:
            int_val = int(float(s))
            if float(s) == int_val:
                return str(int_val)
        except (ValueError, TypeError):
            pass
    return s


def extract_state_short(state_code_full):
    """Extract short state code: 'Gujarat(GJ)' → '(GJ)'"""
    if not state_code_full:
        return ''
    s = str(state_code_full)
    # Find content in parentheses
    start = s.find('(')
    if start != -1:
        end = s.find(')', start)
        if end != -1:
            return s[start:end + 1]
    return s


def read_xls_file(filepath):
    """Read .xls file handling corruption."""
    try:
        wb = xlrd.open_workbook(filepath, ignore_workbook_corruption=True)
        ws = wb.sheet_by_index(0)
        
        # Read all data into list of dicts
        headers = []
        seen = {}
        for c in range(ws.ncols):
            h = str(ws.cell_value(1, c)).strip()  # Row 2 = headers (index 1)
            if not h or h == '':
                h = f'Col_{c}'
            if h in seen:
                seen[h] += 1
                h = f"{h}_{seen[h]}"
            else:
                seen[h] = 1
            headers.append(h)
        
        rows = []
        for r in range(2, ws.nrows):  # Skip title row and header row
            row_data = {}
            for c in range(ws.ncols):
                cell_value = ws.cell_value(r, c)
                # Handle dates
                if ws.cell_type(r, c) == xlrd.XL_CELL_DATE:
                    try:
                        date_tuple = xlrd.xldate_as_tuple(cell_value, wb.datemode)
                        cell_value = f'{date_tuple[2]:02d}/{date_tuple[1]:02d}/{date_tuple[0]}'
                    except Exception:
                        pass
                row_data[headers[c]] = cell_value
            rows.append(row_data)
        
        return headers, rows
    except Exception as e:
        raise Exception(f'Error reading {os.path.basename(filepath)}: {str(e)}')


def read_xlsx_file(filepath):
    """Read .xlsx file."""
    try:
        wb = openpyxl.load_workbook(filepath, data_only=True)
        ws = wb.active
        
        headers = []
        seen = {}
        for c in range(1, ws.max_column + 1):
            h = ws.cell(row=2, column=c).value  # Row 2 = headers
            if h is None:
                h = f'Col_{c}'
            h = str(h).strip()
            if h in seen:
                seen[h] += 1
                h = f"{h}_{seen[h]}"
            else:
                seen[h] = 1
            headers.append(h)
        
        rows = []
        for r in range(3, ws.max_row + 1):  # Data starts at row 3
            row_data = {}
            for c in range(1, ws.max_column + 1):
                row_data[headers[c - 1]] = ws.cell(row=r, column=c).value
            rows.append(row_data)
        
        return headers, rows
    except Exception as e:
        raise Exception(f'Error reading {os.path.basename(filepath)}: {str(e)}')


def read_file(filepath):
    """Read Excel file (auto-detect format)."""
    ext = os.path.splitext(filepath)[1].lower()
    if ext == '.xls':
        return read_xls_file(filepath)
    elif ext == '.xlsx':
        return read_xlsx_file(filepath)
    else:
        raise Exception(f'Unsupported file format: {ext}')


# ─── LOOKUP Logic ────────────────────────────────────────────────────────

def build_sale_summary_lookup(sale_summary_rows, sale_summary_headers):
    """Build lookup dict from SALE SUMMARY keyed by New Invoice No."""
    lookup = {}
    for row in sale_summary_rows:
        # Try "New Invoice No." first, then "New Invoice No"
        key = None
        for possible_key in ['New Invoice No.', 'New Invoice No', 'New invoice No.']:
            if possible_key in row:
                key = safe_str(row[possible_key])
                break
        
        if key:
            lookup[key] = {
                'invoice_type': safe_str(row.get('Invoice Type', '')),
                'invoice_id': safe_str(row.get('Invoice No', row.get('Invoice No.', ''))),
                'state_code': safe_str(row.get('State Code', '')),
                'channel': safe_str(row.get('Channel', '')),
            }
    return lookup


def build_purchase_summary_lookup(purchase_summary_rows, purchase_summary_headers):
    """Build lookup dict from PURCHASE SUMMARY keyed by Invoice No."""
    lookup = {}
    for row in purchase_summary_rows:
        key = None
        for possible_key in ['Invoice No.', 'Invoice No', 'Invoice Number']:
            if possible_key in row:
                key = safe_str(row[possible_key])
                break
        
        if key:
            lookup[key] = {
                'invoice_type': safe_str(row.get('Invoice Type', '')),
                'order_no': safe_str(row.get('Order No', '')),
            }
    return lookup


# ─── Main Processing ───────────────────────────────────────────────────────

def process_files(sale_details_path, sale_summary_path, purchase_details_path, purchase_summary_path):
    """Process 4 files and generate combined data."""
    
    # Read all files
    sd_headers, sd_rows = read_file(sale_details_path)
    ss_headers, ss_rows = read_file(sale_summary_path)
    pd_headers, pd_rows = read_file(purchase_details_path)
    ps_headers, ps_rows = read_file(purchase_summary_path)
    
    # Build LOOKUP dicts
    sale_lookup = build_sale_summary_lookup(ss_rows, ss_headers)
    purchase_lookup = build_purchase_summary_lookup(ps_rows, ps_headers)
    
    # ─── Process SALE side ──────────────────────────────────────────────
    sale_processed = []
    for idx, row in enumerate(sd_rows):
        invoice_no = safe_str(row.get('Invoice No', ''))
        
        # VLOOKUP from SALE SUMMARY
        summary = sale_lookup.get(invoice_no, {})
        
        # Direct columns from SALE DETAILS
        quantity = safe_float(row.get('Quantity', 0))
        item_cost = safe_float(row.get('Item Cost', row.get('Item cost', 0)))
        gross = safe_float(row.get('Gross', 0))
        igst_rate = safe_float(row.get('IGST', 0))
        cgst_rate = safe_float(row.get('CGST', 0))
        sgst_rate = safe_float(row.get('SGST', 0))
        igst_amt = safe_float(row.get('IGST Amt', 0))
        cgst_amt = safe_float(row.get('CGST Amt', 0))
        sgst_amt = safe_float(row.get('SGST Amt', 0))
        invoice_val = safe_float(row.get('Invoice', 0))
        
        order_id = safe_str(row.get('Order ID', ''))
        item_asin = safe_str(row.get('Item Asin', ''))
        item_sku = safe_str(row.get('Item SKU', ''))
        
        # SALE UNIQUE ID = InvoiceNo - OrderID - ItemAsin - ItemSKU
        sale_unique_id = f'{invoice_no}-{order_id}-{item_asin}-{item_sku}'
        
        # State code
        state_code_full = summary.get('state_code', '')
        state_code_short = extract_state_short(state_code_full)
        
        # Duplicate calculated columns (Col 28-37)
        calc_qty = quantity  # Col 28 = Col 13
        calc_cost = item_cost  # Col 29 = Col 14
        calc_gross = roundup2(calc_qty * calc_cost)  # Col 30 = 28 * 29
        calc_igst = igst_rate  # Col 31 = Col 16
        calc_cgst = cgst_rate  # Col 32 = Col 17
        calc_sgst = sgst_rate  # Col 33 = Col 18
        calc_igst_amt = roundup2(calc_gross * calc_igst / 100) if calc_igst else 0  # Col 34
        calc_cgst_amt = roundup2(calc_gross * calc_cgst / 100) if calc_cgst else 0  # Col 35
        calc_sgst_amt = roundup2(calc_gross * calc_sgst / 100) if calc_sgst else 0  # Col 36
        calc_invoice = calc_gross + calc_igst + calc_cgst + calc_sgst  # Col 37
        
        sale_row = {
            'no': idx + 1,
            'invoice_no': invoice_no,
            'type': summary.get('invoice_type', ''),
            'invoice_date': safe_str(row.get('Invoice Date', '')),
            'warehouse_name': safe_str(row.get('Warehouse Name', '')),
            'warehouse_code': safe_str(row.get('Warehouse Code', '')),
            'gst_no': safe_str(row.get('GST No', '')),
            'order_id': order_id,
            'item_asin': item_asin,
            'item_sku': item_sku,
            'item_name': safe_str(row.get('Item Name', '')),
            'hsn_number': safe_str(row.get('HSN Number', '')),
            'quantity': quantity,
            'item_cost': item_cost,
            'gross': gross,
            'igst': igst_rate,
            'cgst': cgst_rate,
            'sgst': sgst_rate,
            'igst_amt': igst_amt,
            'cgst_amt': cgst_amt,
            'sgst_amt': sgst_amt,
            'invoice': invoice_val,
            'reason': safe_str(row.get('Reason', '')),
            'zoho_status': safe_str(row.get('Zoho Status', '')),
            'invoice_id': summary.get('invoice_id', ''),
            'state_code': state_code_full,
            'sale_unique_id': sale_unique_id,
            # Duplicate calculated columns
            'calc_qty': calc_qty,
            'calc_cost': calc_cost,
            'calc_gross': calc_gross,
            'calc_igst': calc_igst,
            'calc_cgst': calc_cgst,
            'calc_sgst': calc_sgst,
            'calc_igst_amt': calc_igst_amt,
            'calc_cgst_amt': calc_cgst_amt,
            'calc_sgst_amt': calc_sgst_amt,
            'calc_invoice': roundup2(calc_invoice),
            'state_code_short': state_code_short,
        }
        sale_processed.append(sale_row)
    
    # ─── Process PURCHASE side ──────────────────────────────────────────
    purchase_processed = []
    for idx, row in enumerate(pd_rows):
        invoice_no = safe_str(row.get('Invoice No', ''))
        
        order_id = safe_str(row.get('Order ID', ''))
        item_asin = safe_str(row.get('Item Asin', ''))
        item_sku = safe_str(row.get('Item SKU', ''))
        
        quantity = safe_float(row.get('Quantity', 0))
        
        # Handle the duplicate "Quantity" column that's actually "Item cost"
        # In PURCHASE DETAILS, Col 12 header says "Quantity" but value is Item cost
        item_cost = 0
        if 'Quantity_2' in row:
            item_cost = safe_float(row.get('Quantity_2', 0))
        else:
            item_cost_key = 'Item cost'
            if item_cost_key not in row:
                item_cost_key = 'Item Cost'
            item_cost = safe_float(row.get(item_cost_key, 0))
            if item_cost == 0:
                # Fallback: the item cost might be mislabeled
                # Look at all values and find the one that looks like a cost
                for k, v in row.items():
                    if k not in [item_cost_key] and 'cost' in k.lower():
                        item_cost = safe_float(v)
                        break
        
        gross = safe_float(row.get('Gross', 0))
        igst_rate = safe_float(row.get('IGST', 0))
        cgst_rate = safe_float(row.get('CGST', 0))
        sgst_rate = safe_float(row.get('SGST', 0))
        igst_amt = safe_float(row.get('IGST Amt', 0))
        cgst_amt = safe_float(row.get('CGST Amt', 0))
        sgst_amt = safe_float(row.get('SGST Amt', 0))
        invoice_val = safe_float(row.get('Invoice', 0))
        
        # PURCHASE UNIQUE ID = InvoiceNo - OrderID - ItemAsin - ItemSKU (with "-" separator)
        purchase_unique_id = f'{invoice_no}-{order_id}-{item_asin}-{item_sku}'
        
        # Duplicate calculated columns (Col 61-69)
        calc_qty = quantity  # Col 61
        calc_cost = item_cost  # Col 62
        calc_gross = roundup2(calc_qty * calc_cost)  # Col 63
        calc_igst = igst_rate  # Col 64
        calc_cgst = cgst_rate  # Col 65
        calc_sgst = sgst_rate  # Col 66
        calc_igst_amt = roundup2(calc_gross * calc_igst / 100) if calc_igst else 0  # Col 67
        calc_cgst_amt = roundup2(calc_gross * calc_cgst / 100) if calc_cgst else 0  # Col 68
        calc_sgst_amt = roundup2(calc_gross * calc_sgst / 100) if calc_sgst else 0  # Col 69
        calc_total_amt = roundup2(calc_gross + calc_igst_amt + calc_cgst_amt + calc_sgst_amt)  # Col 70
        
        purchase_row = {
            'no': idx + 1,
            'invoice_no': invoice_no,
            'warehouse_name': safe_str(row.get('Warehouse Name', '')),
            'warehouse_code': safe_str(row.get('Warehouse Code', '')),
            'gst_no': safe_str(row.get('GST No', '')),
            'order_id': order_id,
            'item_asin': item_asin,
            'item_sku': item_sku,
            'item_name': safe_str(row.get('Item Name', '')),
            'hsn_number': safe_str(row.get('HSN Number', '')),
            'quantity': quantity,
            'item_cost': item_cost,
            'gross': gross,
            'igst': igst_rate,
            'cgst': cgst_rate,
            'sgst': sgst_rate,
            'igst_amt': igst_amt,
            'cgst_amt': cgst_amt,
            'sgst_amt': sgst_amt,
            'invoice': invoice_val,
            'purchase_unique_id': purchase_unique_id,
            # Duplicate calculated columns
            'calc_qty': calc_qty,
            'calc_cost': calc_cost,
            'calc_gross': calc_gross,
            'calc_igst': calc_igst,
            'calc_cgst': calc_cgst,
            'calc_sgst': calc_sgst,
            'calc_igst_amt': calc_igst_amt,
            'calc_cgst_amt': calc_cgst_amt,
            'calc_sgst_amt': calc_sgst_amt,
            'calc_total_amt': calc_total_amt,
        }
        purchase_processed.append(purchase_row)
    
    return sale_processed, purchase_processed


# ─── Excel Export ────────────────────────────────────────────────────────

def export_to_excel(sale_data, purchase_data, output_path):
    """Export combined data to Excel in the final format."""
    wb = openpyxl.Workbook()
    ws = wb.active
    fy = get_financial_year()
    ws.title = f'AJIO & MYNTRA SALE-PURCHASE {fy}'
    
    # ─── Styles ─────────────────────────────────────────────────────────
    header_font = Font(name='Calibri', bold=True, size=12, color='FFFFFF')
    sale_fill = PatternFill(start_color='1B5E20', end_color='1B5E20', fill_type='solid')
    purchase_fill = PatternFill(start_color='0D47A1', end_color='0D47A1', fill_type='solid')
    subheader_font = Font(name='Calibri', bold=True, size=10)
    sale_subheader_fill = PatternFill(start_color='C8E6C9', end_color='C8E6C9', fill_type='solid')
    purchase_subheader_fill = PatternFill(start_color='BBDEFB', end_color='BBDEFB', fill_type='solid')
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    center_align = Alignment(horizontal='center', vertical='center')
    
    # ─── Row 1: Merged Headers ──────────────────────────────────────────
    ws.merge_cells('A1:AL1')  # SALE
    ws.merge_cells('AN1:BR1')  # PURCHASE
    
    sale_cell = ws.cell(row=1, column=1, value='SALE')
    sale_cell.font = header_font
    sale_cell.fill = sale_fill
    sale_cell.alignment = center_align
    
    purchase_cell = ws.cell(row=1, column=40, value='PURCHASE')
    purchase_cell.font = header_font
    purchase_cell.fill = purchase_fill
    purchase_cell.alignment = center_align
    
    # ─── Row 2: Column Headers ──────────────────────────────────────────
    sale_headers = [
        'No.', 'Invoice No', 'TYPE', 'Invoice Date', 'Warehouse Name',
        'Warehouse Code', 'GST No', 'Order ID', 'Item Asin', 'Item SKU',
        'Item Name', 'HSN Number', 'Quantity', 'Item Cost', 'Gross',
        'IGST', 'CGST', 'SGST', 'IGST Amt', 'CGST Amt', 'SGST Amt',
        'Invoice', 'Reason', 'Zoho Status', 'invoice id', 'State Code',
        'SALE UNIQUE ID', 'Quantity', 'Item Cost', 'Gross',
        'IGST', 'CGST', 'SGST', 'IGST Amt', 'CGST Amt', 'SGST Amt',
        'Invoice', 'STATE CODE'
    ]
    
    purchase_headers = [
        'No.', 'Invoice No', 'Warehouse Name', 'Warehouse Code', 'GST No',
        'Order ID', 'Item Asin', 'Item SKU', 'Item Name', 'HSN Number',
        'Quantity', 'Item cost', 'Gross', 'IGST', 'CGST', 'SGST',
        'IGST Amt', 'CGST Amt', 'SGST Amt', 'Invoice',
        'PURCHASE UNIQUE ID', 'Quantity', 'Item cost', 'Gross',
        'IGST', 'CGST', 'SGST', 'IGST Amt', 'CGST Amt', 'SGST Amt',
        'Total Amt'
    ]
    
    # Write SALE headers (Col 1-38)
    for i, h in enumerate(sale_headers):
        cell = ws.cell(row=2, column=i + 1, value=h)
        cell.font = subheader_font
        cell.fill = sale_subheader_fill
        cell.border = thin_border
        cell.alignment = center_align
    
    # Write PURCHASE headers (Col 40-70)
    for i, h in enumerate(purchase_headers):
        cell = ws.cell(row=2, column=40 + i, value=h)
        cell.font = subheader_font
        cell.fill = purchase_subheader_fill
        cell.border = thin_border
        cell.alignment = center_align
    
    # ─── Data Rows ──────────────────────────────────────────────────────
    max_rows = max(len(sale_data), len(purchase_data))
    
    sale_keys = [
        'no', 'invoice_no', 'type', 'invoice_date', 'warehouse_name',
        'warehouse_code', 'gst_no', 'order_id', 'item_asin', 'item_sku',
        'item_name', 'hsn_number', 'quantity', 'item_cost', 'gross',
        'igst', 'cgst', 'sgst', 'igst_amt', 'cgst_amt', 'sgst_amt',
        'invoice', 'reason', 'zoho_status', 'invoice_id', 'state_code',
        'sale_unique_id', 'calc_qty', 'calc_cost', 'calc_gross',
        'calc_igst', 'calc_cgst', 'calc_sgst', 'calc_igst_amt',
        'calc_cgst_amt', 'calc_sgst_amt', 'calc_invoice', 'state_code_short'
    ]
    
    purchase_keys = [
        'no', 'invoice_no', 'warehouse_name', 'warehouse_code', 'gst_no',
        'order_id', 'item_asin', 'item_sku', 'item_name', 'hsn_number',
        'quantity', 'item_cost', 'gross', 'igst', 'cgst', 'sgst',
        'igst_amt', 'cgst_amt', 'sgst_amt', 'invoice',
        'purchase_unique_id', 'calc_qty', 'calc_cost', 'calc_gross',
        'calc_igst', 'calc_cgst', 'calc_sgst', 'calc_igst_amt',
        'calc_cgst_amt', 'calc_sgst_amt', 'calc_total_amt'
    ]
    
    for r in range(max_rows):
        row_num = r + 3  # Data starts at row 3
        
        # SALE data
        if r < len(sale_data):
            for c, key in enumerate(sale_keys):
                cell = ws.cell(row=row_num, column=c + 1, value=sale_data[r].get(key, ''))
                cell.border = thin_border
        
        # PURCHASE data
        if r < len(purchase_data):
            for c, key in enumerate(purchase_keys):
                cell = ws.cell(row=row_num, column=40 + c, value=purchase_data[r].get(key, ''))
                cell.border = thin_border
    
    # ─── Column Widths ──────────────────────────────────────────────────
    for col in range(1, 72):
        ws.column_dimensions[get_column_letter(col)].width = 15
    # Wider columns for names
    ws.column_dimensions['K'].width = 40  # Item Name
    ws.column_dimensions['AW'].width = 40  # Purchase Item Name
    
    wb.save(output_path)
    return output_path


# ─── Routes ──────────────────────────────────────────────────────────

@app.route('/api/lookup-sale-quantities', methods=['POST', 'OPTIONS'])
def lookup_sale_quantities():
    if request.method == 'OPTIONS':
        response = jsonify({"status": "Success"})
        response.headers.add("Access-Control-Allow-Origin", "*")
        response.headers.add("Access-Control-Allow-Headers", "Content-Type")
        response.headers.add("Access-Control-Allow-Methods", "POST, OPTIONS")
        return response

    try:
        req_data = request.json or {}
        keys = req_data.get('keys', [])
        if not keys or not isinstance(keys, list):
            response = jsonify({"status": "Error", "message": "Invalid keys format"})
            response.headers.add("Access-Control-Allow-Origin", "*")
            return response, 400

        # Load the stored data
        stored = load_stored_data()
        sale_rows = stored.get('sale', [])

        # Build a fast mapping from sale_rows
        db_mapping = {}
        for row in sale_rows:
            uid = row.get('sale_unique_id')
            if uid:
                uid_str = str(uid).strip().upper()
                db_mapping[uid_str] = row.get('calc_qty', row.get('quantity', 0))

        # Match keys
        mapping = {}
        for key in keys:
            norm_key = str(key).strip().upper()
            if norm_key in db_mapping:
                mapping[key] = db_mapping[norm_key]

        response = jsonify({"status": "Success", "mapping": mapping})
        response.headers.add("Access-Control-Allow-Origin", "*")
        return response
    except Exception as e:
        traceback.print_exc()
        response = jsonify({"status": "Error", "message": str(e)})
        response.headers.add("Access-Control-Allow-Origin", "*")
        return response, 500

@app.route('/')
def index():
    return render_template('index.html')


@app.route('/upload', methods=['POST'])
def upload_files():
    """Handle file uploads, process data, merge with existing, deduplicate."""
    try:
        required_files = ['sale_details', 'sale_summary', 'purchase_details', 'purchase_summary']
        file_paths = {}
        
        for file_key in required_files:
            if file_key not in request.files:
                return jsonify({'error': f'Missing file: {file_key}'}), 400
            
            file = request.files[file_key]
            if file.filename == '':
                return jsonify({'error': f'No file selected for: {file_key}'}), 400
            
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], f'{file_key}_{file.filename}')
            file.save(filepath)
            file_paths[file_key] = filepath
        
        # Process new uploaded files
        new_sale_data, new_purchase_data = process_files(
            file_paths['sale_details'],
            file_paths['sale_summary'],
            file_paths['purchase_details'],
            file_paths['purchase_summary']
        )
        
        # Load existing stored data
        stored = load_stored_data()
        
        # Merge with deduplication using UNIQUE IDs
        merged_sale, sale_ids, sale_new, sale_dups = merge_with_dedup(
            stored['sale'], new_sale_data,
            stored['sale_ids'], 'sale_unique_id'
        )
        
        merged_purchase, purchase_ids, purchase_new, purchase_dups = merge_with_dedup(
            stored['purchase'], new_purchase_data,
            stored['purchase_ids'], 'purchase_unique_id'
        )
        
        # Save merged data persistently
        stored_data = {
            'sale': merged_sale,
            'purchase': merged_purchase,
            'sale_ids': sale_ids,
            'purchase_ids': purchase_ids,
        }
        save_stored_data(stored_data)
        
        # Also save to temp for export
        data_path = os.path.join(app.config['UPLOAD_FOLDER'], 'processed_data.pkl')
        with open(data_path, 'wb') as f:
            pickle.dump((merged_sale, merged_purchase), f)
        
        # Return preview (first 100 rows)
        preview_sale = merged_sale[:100]
        preview_purchase = merged_purchase[:100]
        
        return jsonify({
            'success': True,
            'sale_count': len(merged_sale),
            'purchase_count': len(merged_purchase),
            'sale_new': sale_new,
            'sale_duplicates': sale_dups,
            'purchase_new': purchase_new,
            'purchase_duplicates': purchase_dups,
            'financial_year': get_financial_year(),
            'sale_preview': preview_sale,
            'purchase_preview': preview_purchase,
        })
    
    except Exception as e:
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500


@app.route('/export', methods=['GET'])
def export_file():
    """Export processed data as Excel."""
    try:
        # Try persistent store first, then temp
        stored = load_stored_data()
        if stored['sale'] or stored['purchase']:
            sale_data = stored['sale']
            purchase_data = stored['purchase']
        else:
            data_path = os.path.join(app.config['UPLOAD_FOLDER'], 'processed_data.pkl')
            if not os.path.exists(data_path):
                return jsonify({'error': 'No processed data. Please upload files first.'}), 400
            with open(data_path, 'rb') as f:
                sale_data, purchase_data = pickle.load(f)
        
        fy = get_financial_year()
        output_path = os.path.join(app.config['UPLOAD_FOLDER'], f'AJIO_MYNTRA_FLIPKART_{fy}_Output.xlsx')
        export_to_excel(sale_data, purchase_data, output_path)
        
        return send_file(
            output_path,
            as_attachment=True,
            download_name=f'AJIO_MYNTRA_FLIPKART_Sale_Purchase_{fy}.xlsx',
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
    
    except Exception as e:
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500


@app.route('/preview-more', methods=['GET'])
def preview_more():
    """Get more rows for preview."""
    try:
        # Try persistent store first
        stored = load_stored_data()
        if stored['sale'] or stored['purchase']:
            sale_data = stored['sale']
            purchase_data = stored['purchase']
        else:
            data_path = os.path.join(app.config['UPLOAD_FOLDER'], 'processed_data.pkl')
            if not os.path.exists(data_path):
                return jsonify({'error': 'No data available'}), 400
            with open(data_path, 'rb') as f:
                sale_data, purchase_data = pickle.load(f)
        
        offset = int(request.args.get('offset', 0))
        limit = int(request.args.get('limit', 100))
        
        return jsonify({
            'sale_data': sale_data[offset:offset + limit],
            'purchase_data': purchase_data[offset:offset + limit],
            'sale_total': len(sale_data),
            'purchase_total': len(purchase_data),
        })
    
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/status', methods=['GET'])
def status():
    """Get current data status."""
    stored = load_stored_data()
    return jsonify({
        'sale_count': len(stored['sale']),
        'purchase_count': len(stored['purchase']),
        'financial_year': get_financial_year(),
        'has_data': bool(stored['sale'] or stored['purchase']),
    })


@app.route('/clear', methods=['POST'])
def clear_data():
    """Clear all stored data (fresh start)."""
    try:
        empty_store = {'sale': [], 'purchase': [], 'sale_ids': set(), 'purchase_ids': set()}
        save_stored_data(empty_store)

        data_path = os.path.join(app.config['UPLOAD_FOLDER'], 'processed_data.pkl')
        try:
            if os.path.exists(data_path):
                os.remove(data_path)
        except OSError:
            with open(data_path, 'wb') as f:
                pickle.dump(([], []), f)

        return jsonify({'success': True, 'message': 'All data cleared!'})
    except Exception as e:
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500


@app.route('/get-sync-data', methods=['POST'])
def get_sync_data():
    """Get processed and deduplicated data prepared for Google Sheets sync."""
    try:
        req_data = request.json or {}
        web_app_url = req_data.get('webAppUrl')
        sheet_name = req_data.get('sheetName')
        mode = req_data.get('mode', 'append')
        
        if not web_app_url:
            return jsonify({'error': 'Google Web App URL is required'}), 400
        if not sheet_name:
            return jsonify({'error': 'Target Sheet Name is required'}), 400
            
        stored = load_stored_data()
        sale_data = [dict(row) for row in stored['sale']]
        purchase_data = [dict(row) for row in stored['purchase']]
        
        if not sale_data and not purchase_data:
            return jsonify({'error': 'No data processed yet. Please upload files first.'}), 400
            
        import requests
        from urllib.parse import quote
        
        original_sale_count = len(sale_data)
        original_purchase_count = len(purchase_data)
        skipped_sale = 0
        skipped_purchase = 0
        
        if mode == 'append':
            try:
                url = f"{web_app_url}?action=getExistingUniqueIds&sheetName={quote(sheet_name)}"
                res = requests.get(url, allow_redirects=False)
                if res.status_code in [302, 307, 308]:
                    redirect_url = res.headers.get('Location')
                    res = requests.get(redirect_url)
                
                if res.status_code == 200:
                    res_data = res.json()
                    if res_data.get('success'):
                        existing_sale_ids = set(str(uid) for uid in res_data.get('saleIds', []))
                        existing_purchase_ids = set(str(uid) for uid in res_data.get('purchaseIds', []))
                        
                        sale_filtered = [row for row in sale_data if str(row.get('sale_unique_id', '')) not in existing_sale_ids]
                        purchase_filtered = [row for row in purchase_data if str(row.get('purchase_unique_id', '')) not in existing_purchase_ids]
                        
                        skipped_sale = len(sale_data) - len(sale_filtered)
                        skipped_purchase = len(purchase_data) - len(purchase_filtered)
                        
                        sale_data = sale_filtered
                        purchase_data = purchase_filtered
                        
                        for idx, row in enumerate(sale_data):
                            row['no'] = len(existing_sale_ids) + idx + 1
                        for idx, row in enumerate(purchase_data):
                            row['no'] = len(existing_purchase_ids) + idx + 1
            except Exception as e:
                print(f"Error fetching unique IDs: {str(e)}")
                
        # Map dict lists to 2D lists matching Google Sheet column layouts
        sale_keys = [
            'no', 'invoice_no', 'type', 'invoice_date', 'warehouse_name',
            'warehouse_code', 'gst_no', 'order_id', 'item_asin', 'item_sku',
            'item_name', 'hsn_number', 'quantity', 'item_cost', 'gross',
            'igst', 'cgst', 'sgst', 'igst_amt', 'cgst_amt', 'sgst_amt',
            'invoice', 'reason', 'zoho_status', 'invoice_id', 'state_code',
            'sale_unique_id', 'calc_qty', 'calc_cost', 'calc_gross',
            'calc_igst', 'calc_cgst', 'calc_sgst', 'calc_igst_amt',
            'calc_cgst_amt', 'calc_sgst_amt', 'calc_invoice', 'state_code_short'
        ]
        
        purchase_keys = [
            'no', 'invoice_no', 'warehouse_name', 'warehouse_code', 'gst_no',
            'order_id', 'item_asin', 'item_sku', 'item_name', 'hsn_number',
            'quantity', 'item_cost', 'gross', 'igst', 'cgst', 'sgst',
            'igst_amt', 'cgst_amt', 'sgst_amt', 'invoice',
            'purchase_unique_id', 'calc_qty', 'calc_cost', 'calc_gross',
            'calc_igst', 'calc_cgst', 'calc_sgst', 'calc_igst_amt',
            'calc_cgst_amt', 'calc_sgst_amt', 'calc_total_amt'
        ]
        
        sale_rows_2d = [[row.get(k, '') for k in sale_keys] for row in sale_data]
        purchase_rows_2d = [[row.get(k, '') for k in purchase_keys] for row in purchase_data]
        
        return jsonify({
            'success': True,
            'saleRows': sale_rows_2d,
            'purchaseRows': purchase_rows_2d,
            'originalSale': original_sale_count,
            'originalPurchase': original_purchase_count,
            'skippedSale': skipped_sale,
            'skippedPurchase': skipped_purchase,
            'toSyncSale': len(sale_rows_2d),
            'toSyncPurchase': len(purchase_rows_2d),
        })
    except Exception as e:
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500


@app.route('/sync-sheet', methods=['POST'])
def sync_sheet():
    """Sync stored data to Google Sheet via the Google Web App URL."""
    try:
        req_data = request.json or {}
        web_app_url = req_data.get('webAppUrl')
        sheet_name = req_data.get('sheetName')
        mode = req_data.get('mode', 'append')  # 'append' or 'overwrite'
        
        if not web_app_url:
            return jsonify({'error': 'Google Web App URL is required'}), 400
        if not sheet_name:
            return jsonify({'error': 'Target Sheet Name is required'}), 400
            
        # Load the stored data
        stored = load_stored_data()
        sale_data = [dict(row) for row in stored['sale']]
        purchase_data = [dict(row) for row in stored['purchase']]
        
        if not sale_data and not purchase_data:
            return jsonify({'error': 'No data processed yet. Please upload files first.'}), 400

        import requests
        from urllib.parse import quote

        # If in append mode, fetch existing unique IDs to avoid duplication in Google Sheets
        if mode == 'append':
            try:
                # Call getExistingUniqueIds from Google Sheets Web App
                url = f"{web_app_url}?action=getExistingUniqueIds&sheetName={quote(sheet_name)}"
                res = requests.get(url, allow_redirects=False)
                if res.status_code in [302, 307, 308]:
                    redirect_url = res.headers.get('Location')
                    res = requests.get(redirect_url)
                
                if res.status_code == 200:
                    res_data = res.json()
                    if res_data.get('success'):
                        existing_sale_ids = set(str(uid) for uid in res_data.get('saleIds', []))
                        existing_purchase_ids = set(str(uid) for uid in res_data.get('purchaseIds', []))
                        
                        # Filter out rows that already exist in the sheet
                        sale_data = [row for row in sale_data if str(row.get('sale_unique_id', '')) not in existing_sale_ids]
                        purchase_data = [row for row in purchase_data if str(row.get('purchase_unique_id', '')) not in existing_purchase_ids]
                        
                        # Re-number the appended rows starting after the existing counts
                        for idx, row in enumerate(sale_data):
                            row['no'] = len(existing_sale_ids) + idx + 1
                        for idx, row in enumerate(purchase_data):
                            row['no'] = len(existing_purchase_ids) + idx + 1
            except Exception as e:
                # Log the error but proceed with default append if it fails
                print(f"Error fetching unique IDs for append deduplication: {str(e)}")
            
        # Map dict lists to 2D lists matching Google Sheet column layouts
        sale_keys = [
            'no', 'invoice_no', 'type', 'invoice_date', 'warehouse_name',
            'warehouse_code', 'gst_no', 'order_id', 'item_asin', 'item_sku',
            'item_name', 'hsn_number', 'quantity', 'item_cost', 'gross',
            'igst', 'cgst', 'sgst', 'igst_amt', 'cgst_amt', 'sgst_amt',
            'invoice', 'reason', 'zoho_status', 'invoice_id', 'state_code',
            'sale_unique_id', 'calc_qty', 'calc_cost', 'calc_gross',
            'calc_igst', 'calc_cgst', 'calc_sgst', 'calc_igst_amt',
            'calc_cgst_amt', 'calc_sgst_amt', 'calc_invoice', 'state_code_short'
        ]
        
        purchase_keys = [
            'no', 'invoice_no', 'warehouse_name', 'warehouse_code', 'gst_no',
            'order_id', 'item_asin', 'item_sku', 'item_name', 'hsn_number',
            'quantity', 'item_cost', 'gross', 'igst', 'cgst', 'sgst',
            'igst_amt', 'cgst_amt', 'sgst_amt', 'invoice',
            'purchase_unique_id', 'calc_qty', 'calc_cost', 'calc_gross',
            'calc_igst', 'calc_cgst', 'calc_sgst', 'calc_igst_amt',
            'calc_cgst_amt', 'calc_sgst_amt', 'calc_total_amt'
        ]
        
        sale_rows_2d = [[row.get(k, '') for k in sale_keys] for row in sale_data]
        purchase_rows_2d = [[row.get(k, '') for k in purchase_keys] for row in purchase_data]
        
        # Prepare the payload for Google Apps Script Web App
        payload = {
            'action': 'writeData',
            'sheetName': sheet_name,
            'saleRows': sale_rows_2d,
            'purchaseRows': purchase_rows_2d,
            'isAppend': (mode == 'append')
        }
        
        # Call the Google Web App
        import requests
        # Manually handle 302 redirects to preserve POST method
        response = requests.post(web_app_url, json=payload, allow_redirects=False)
        if response.status_code in [302, 307, 308]:
            redirect_url = response.headers.get('Location')
            response = requests.post(redirect_url, json=payload)
            
        if response.status_code != 200:
            return jsonify({'error': f'Failed to write to Google Sheets. Status code: {response.status_code}'}), 500
            
        result = response.json()
        if not result.get('success'):
            return jsonify({'error': result.get('error', 'Unknown error writing to Google Sheets')}), 500
            
        return jsonify({
            'success': True,
            'message': result.get('message', 'Data synced successfully!'),
            'saleRows': result.get('saleRows', 0),
            'purchaseRows': result.get('purchaseRows', 0)
        })
        
    except Exception as e:
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500


if __name__ == '__main__':
    app.run(debug=True, port=5000)
