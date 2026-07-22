import pandas as pd
import numpy as np
import os
import sys
import warnings

# Ensure prints don't crash on Windows with charmap encoding
if sys.stdout.encoding.lower() != 'utf-8':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except AttributeError:
        pass

warnings.filterwarnings('ignore')

# ── Column used to detect NA rows ──────────────────────────────
NA_CHECK_COLS = ["FSN", "NC ID", "Sales Price", "lot weight ID", "customer_contact_number(req)"]

OUTPUT_COLS = [
    "customer_contact_number(req)",
    "NC ID",
    "NC Name",
    "QTY",
    "Date",
    "lot weight ID",
    "ordering_mode(optional)",
    "cancelled (optional)By default should be 0",
    "purchaseOrder",
    "Sales Price",
    "DELIVERY_CHARGE(opt)",
    "CITY_ID(req)",
    "sale_order_id(optional- leave empty)",
    "sub_type (optional- leave empty)",
    "CategoryId (if empty then default is 1)",
    "grocerFlow",
]

# ── City config ─────────────────────────────────────────────────
CITY_CONFIG = {
    "Bangalore": {
        "po_prefix":        "NCXDB",
        "city_initial":     "B",
        "city_id":          2,
        "alloc_city_name":  "Bangalore",
        "so_sheet":         "BLR FK GRO SO",
        "po_sheet":         "BLR FK Gro PO FIle",
        "sku_sheet":        "BLR GRO SKU config",
        "nlc_sheet":        "BLR FK Gro NLC",
        "cust_sheet":       "BLR FK Customers",
    },
    "Chennai": {
        "po_prefix":        "NCXDC",
        "city_initial":     "C",
        "city_id":          3,
        "alloc_city_name":  "Chennai",
        "so_sheet":         "CHN FK GRO SO",
        "po_sheet":         "CHN FK Gro PO FIle",
        "sku_sheet":        "CHN GRO SKU config",
        "nlc_sheet":        "CHN FK GRO NLC",
        "cust_sheet":       "CHN FK Customers",
    },
    "Mumbai": {
        "po_prefix":        "NCXDM",
        "city_initial":     "M",
        "city_id":          4,
        "alloc_city_name":  "Mumbai",
        "so_sheet":         "MUM FK GRO SO",
        "po_sheet":         "MUM FK Gro PO FIle",
        "sku_sheet":        "MUM GRO SKU config",
        "nlc_sheet":        "MUM FK Gro NLC",
        "cust_sheet":       "MUM FK Customers",
    },
    "Hyderabad": {
        "po_prefix":        "NCXDH",
        "city_initial":     "H",
        "city_id":          5,
        "alloc_city_name":  "Hyderabad",
        "so_sheet":         "HYD FK GRO SO",
        "po_sheet":         "HYD FK Gro PO FIle",
        "sku_sheet":        "HYD GRO SKU config",
        "nlc_sheet":        "HYD FK Gro NLC",
        "cust_sheet":       "HYD FK Customers",
    },
    "Trichy": {
        "po_prefix":        "NCXDT",
        "city_initial":     "T",
        "city_id":          6,
        "alloc_city_name":  "Trichy",
        "so_sheet":         "Trichy FK GRO SO",
        "po_sheet":         "Trichy FK Gro PO FIle",
        "sku_sheet":        "Trichy GRO SKU config",
        "nlc_sheet":        "Trichy FK GRO NLC",
        "cust_sheet":       "Trichy FK Customers",
    },
    "Coimbatore": {
        "po_prefix":        "NCXDCBE",
        "city_initial":     "CBE",
        "city_id":          7,
        "alloc_city_name":  "Coimbatore",
        "so_sheet":         "CBE FK GRO SO",
        "po_sheet":         "CBE FK Gro PO FIle",
        "sku_sheet":        "Coimbatore GRO SKU config",
        "nlc_sheet":        "CBE FK GRO NLC",
        "cust_sheet":       "Coimbatore FK customer",
    },
}

def get_gsheet_client(gsheet_url: str):
    """Authenticate and return the Google Sheets client and the sheet object."""
    import re, gspread
    from google.oauth2.service_account import Credentials
    
    match = re.search(r'/d/([a-zA-Z0-9_-]+)', gsheet_url)
    if not match:
        raise ValueError("Could not extract sheet ID from URL")
    sheet_id = match.group(1)
    
    creds_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'xd-allocation-9640b0ce66d2.json')
    if not os.path.exists(creds_path):
        raise FileNotFoundError(f"Service account credentials not found at {creds_path}")
        
    scopes = [
        'https://www.googleapis.com/auth/spreadsheets',
        'https://www.googleapis.com/auth/drive'
    ]
    credentials = Credentials.from_service_account_file(creds_path, scopes=scopes)
    gc = gspread.authorize(credentials)
    
    try:
        sh = gc.open_by_key(sheet_id)
        return sh, credentials.service_account_email
    except gspread.exceptions.APIError as e:
        if e.response.status_code in (403, 404):
            raise ValueError(f"Permission denied! Please ensure you have shared the Google Sheet with Editor access to: {credentials.service_account_email}")
        raise ValueError(f"Google Sheets API Error: {str(e)}")
        
def update_gsheet_po_file(gsheet_url: str, po_sheet_name: str, df: pd.DataFrame):
    """Clear the PO tab and upload the new Allocation data."""
    import gspread
    sh, sa_email = get_gsheet_client(gsheet_url)
    try:
        worksheet = sh.worksheet(po_sheet_name)
    except gspread.exceptions.WorksheetNotFound:
        raise ValueError(f"Could not find a tab named '{po_sheet_name}' in the Google Sheet.")
        
    worksheet.clear()
    
    # Fill NaN with empty string for gspread compatibility
    df_upload = df.fillna("")
    
    # Convert dataframe to list of lists (including header)
    data = [df_upload.columns.values.tolist()] + df_upload.values.tolist()
    
    # Use USER_ENTERED so that formulas starting with '=' are evaluated by Google Sheets
    try:
        worksheet.update(values=data, range_name="A1", value_input_option="USER_ENTERED")
    except TypeError:
        # Fallback for older gspread versions
        worksheet.update(data, value_input_option="USER_ENTERED")

def drag_formulas_in_so(gsheet_url: str, so_sheet_name: str, target_rows: int):
    """Automatically drags down formulas from row 2 to target_rows in the SO sheet."""
    import gspread
    sh, sa_email = get_gsheet_client(gsheet_url)
    try:
        so_ws = sh.worksheet(so_sheet_name)
    except Exception as e:
        print(f"       [WARN] Could not find SO tab '{so_sheet_name}' to drag formulas: {e}")
        return
        
    req = {
        "autoFill": {
            "useAlternateSeries": False,
            "sourceAndDestination": {
                "source": {
                    "sheetId": so_ws.id,
                    "startRowIndex": 1,  # Row 2 (0-indexed)
                    "endRowIndex": 2,    # Row 2
                    "startColumnIndex": 0,
                    "endColumnIndex": so_ws.col_count
                },
                "dimension": "ROWS",
                "length": max(0, target_rows - 1)  # How many additional rows to fill
            }
        }
    }
    try:
        sh.batch_update({"requests": [req]})
        print(f"       ✅ Formulas dragged down in '{so_sheet_name}' to cover {target_rows} rows.")
    except Exception as e:
        print(f"       [WARN] Failed to auto-fill formulas in SO tab: {e}")

def load_gsheet_so(gsheet_url: str, so_sheet_name: str) -> pd.DataFrame:
    """Fetch SO tab from a Google Sheet URL securely using service account."""
    import gspread
    sh, sa_email = get_gsheet_client(gsheet_url)
    
    try:
        worksheet = sh.worksheet(so_sheet_name)
    except gspread.exceptions.WorksheetNotFound:
        raise ValueError(f"Could not find a tab named '{so_sheet_name}' in the Google Sheet.")
    except Exception as e:
        # Fallback catch for the specific PermissionError wrapper in newer gspread
        if type(e).__name__ == 'PermissionError':
            raise ValueError(f"Permission denied! Please verify the Google Sheet is shared as Viewer with: {sa_email}")
        raise
    data = worksheet.get_all_values()  
    if not data:
        return pd.DataFrame()
    
    return pd.DataFrame(data[1:], columns=data[0])


def run_automation(
    allocation_path: str,
    ecom_path: str,
    city: str,
    delivery_date: str,
    output_dir: str = ".",
    gsheet_url: str = None,
    so_sheet_override: str = None,
    po_sheet_override: str = None,
):
    """
    Main automation.

    Steps:
      1. Read Allocation PO sheet → build Key & PO IDs
      2. Read E.com SO Placement sheet (BLR FK GRO SO) — from local file or G-sheet
      3. Separate valid vs NA rows
      4. Export: CSV (valid) + Excel with two tabs (valid + NA)
    """
    if city not in CITY_CONFIG:
        raise ValueError(f"Unknown city '{city}'. Options: {list(CITY_CONFIG.keys())}")

    cfg = CITY_CONFIG[city]
    os.makedirs(output_dir, exist_ok=True)
    
    so_sheet_name = so_sheet_override.strip() if so_sheet_override and so_sheet_override.strip() else cfg["so_sheet"]
    po_sheet_name = po_sheet_override.strip() if po_sheet_override and po_sheet_override.strip() else cfg["po_sheet"]

    print(f"\n{'='*60}")
    print(f"  {city}  |  PO Prefix: {cfg['po_prefix']}  |  Date: {delivery_date}")
    print(f"{'='*60}\n")

    # ── Step 1: Allocation PO sheet → build Key + PO IDs ────────
    print("► [1/3] Reading Allocation file...")
    alloc = pd.read_excel(allocation_path, sheet_name="PO")
    if "City" in alloc.columns:
        alloc = alloc[alloc["City"].astype(str).str.strip() == cfg["alloc_city_name"]].copy()
    else:
        print("       Detected processed format — using all rows directly")
    print(f"       {len(alloc)} rows for {city}")

    if "Store Site ID" in alloc.columns:
        alloc["Warehouse"] = alloc["Store Site ID"].astype(str).str.strip()
    elif "Warehouse" in alloc.columns:
        alloc["Warehouse"] = alloc["Warehouse"].astype(str).str.strip()

    if "Store ID" not in alloc.columns and "Store" in alloc.columns:
        alloc["Store ID"] = alloc["Store"]
    elif "Store ID" not in alloc.columns:
        alloc["Store ID"] = ""

    alloc["Supplier_ID"] = alloc["Supplier ID"].astype(str).str.strip()
    alloc["Key"] = alloc["Warehouse"] + alloc["Supplier_ID"]

    # Sequential PO ID per unique Key (Warehouse+SupplierID combo)
    unique_keys = list(dict.fromkeys(alloc["Key"].tolist()))  # preserve order
    key_to_poid = {k: f"{cfg['po_prefix']}{str(i+1).zfill(3)}" for i, k in enumerate(unique_keys)}
    alloc["PO_ID_generated"] = alloc["Key"].map(key_to_poid)

    # Build key reference table
    key_table = (
        alloc[["Key", "Warehouse", "Supplier_ID", "PO_ID_generated", "Store ID"]]
        .drop_duplicates("Key")
        .rename(columns={"PO_ID_generated": "PO ID", "Supplier_ID": "Supplier ID", "Store ID": "Store"})
    )

    # ── Step 2: Update PO Tab & Load E.com SO tab ─────────────────
    if gsheet_url:
        import time
        print(f"► [2/3] Updating {po_sheet_name} tab in Google Sheet...")
        
        # Prepare strictly formatted upload dataframe based on user spec
        upload_df = pd.DataFrame()
        upload_df["FSN/ISBN13"] = alloc["FSN"] if "FSN" in alloc.columns else ""
        
        if "FSN_Title" in alloc.columns:
            upload_df["Title"] = alloc["FSN_Title"]
        elif "Title" in alloc.columns:
            upload_df["Title"] = alloc["Title"]
        elif "NC Name" in alloc.columns:
            upload_df["Title"] = alloc["NC Name"]
        else:
            upload_df["Title"] = ""
            
        if "Final PO" in alloc.columns:
            upload_df["QTY"] = alloc["Final PO"]
        elif "QTY" in alloc.columns:
            upload_df["QTY"] = alloc["QTY"]
        elif "Quantity" in alloc.columns:
            upload_df["QTY"] = alloc["Quantity"]
        elif "Qty" in alloc.columns:
            upload_df["QTY"] = alloc["Qty"]
        elif "PO qty" in alloc.columns:
            upload_df["QTY"] = alloc["PO qty"]
        else:
            upload_df["QTY"] = ""
            
        cust_sheet = cfg.get("cust_sheet", "BLR FK Customers")
        
        # Fetch the customer sheet to get the FK Site Name mapping
        print(f"       Fetching '{cust_sheet}' to map FK Site Name...")
        try:
            cust_sh, _ = get_gsheet_client(gsheet_url)
            cust_ws = cust_sh.worksheet(cust_sheet)
            cust_data = cust_ws.get_all_values()
            
            if len(cust_data) > 1:
                cust_df = pd.DataFrame(cust_data[1:], columns=cust_data[0])
                wh_code_col = next((c for c in cust_df.columns if str(c).strip().lower() == "wh code"), None)
                fk_site_col = next((c for c in cust_df.columns if str(c).strip().lower() == "fk site name"), None)
                
                if wh_code_col and fk_site_col:
                    raw_map = cust_df.set_index(wh_code_col)[fk_site_col].to_dict()
                    fk_site_map = {str(k).strip().lower(): v for k, v in raw_map.items()}
                else:
                    fk_site_map = {}
            else:
                fk_site_map = {}
        except Exception as e:
            print(f"       [WARN] Failed to fetch customer sheet mapping: {e}")
            fk_site_map = {}

        upload_df["PO Number"] = alloc["PO_ID_generated"]
        upload_df["Store"] = alloc["Warehouse"].astype(str).str.strip().str.lower().map(fk_site_map).fillna(alloc["Warehouse"])
        
        # Filter out zero-qty or blank rows to prevent Google Sheets 10M cell limit error
        # BUT explicitly KEEP rows if they have a valid FSN or Title (so they go to NA tab)
        upload_df["QTY_NUM"] = pd.to_numeric(upload_df["QTY"], errors="coerce").fillna(0)
        has_qty = upload_df["QTY_NUM"] > 0
        has_fsn = upload_df["FSN/ISBN13"].astype(str).str.strip() != ""
        has_title = upload_df["Title"].astype(str).str.strip() != ""
        
        upload_df = upload_df[has_qty | has_fsn | has_title].copy()
        upload_df.drop(columns=["QTY_NUM"], inplace=True)
        
        num_rows_uploaded = len(upload_df)
        print(f"       {num_rows_uploaded} rows to upload (including missing qty/po ids)")
        upload_df = upload_df.reset_index(drop=True)
        
        # Inject dynamic VLOOKUP formula for the Contact column
        upload_df["Contact"] = [f"=VLOOKUP(E{i},'{cust_sheet}'!F:G,2,0)" for i in range(2, len(upload_df) + 2)]
        
        update_gsheet_po_file(gsheet_url, po_sheet_name, upload_df)
        
        print("       Dragging formulas in SO tab...")
        drag_formulas_in_so(gsheet_url, so_sheet_name, num_rows_uploaded)
        
        print("       Waiting 5 seconds for Google Sheet formulas to evaluate...")
        time.sleep(5)
        
        print("       Fetching updated SO data from Google Sheet...")
        so_df = load_gsheet_so(gsheet_url, so_sheet_name)
    else:
        print("► [2/3] Reading E.com SO Placement data locally...")
        so_df = pd.read_excel(ecom_path, sheet_name=so_sheet_name)

    so_df.columns = so_df.columns.str.strip()
    print(f"       {len(so_df)} total rows in SO tab")

    # Replace Google Sheets errors with NaN
    so_df.replace(["#N/A", "#REF!", "#VALUE!", "#DIV/0!", "#NAME?", "#NUM!", "#NULL!"], np.nan, inplace=True)

    # Convert numeric columns from strings back to proper numbers (since gspread returns all strings)
    for col in so_df.columns:
        if col not in ["Date", "NC Name", "ordering_mode(optional)", "purchaseOrder", "sale_order_id(optional- leave empty)", "customer_contact_number(req)", "sub_type (optional- leave empty)"]:
            try:
                # Remove commas from formatted numbers before parsing
                temp = so_df[col].astype(str).str.replace(',', '', regex=False)
                so_df[col] = pd.to_numeric(temp)
            except (ValueError, TypeError):
                pass

    # Fix contact number formatting (prevent scientific notation)
    def parse_contact(x):
        if pd.isna(x): return np.nan
        try:
            return str(int(float(x)))
        except (ValueError, TypeError):
            return x

    if "customer_contact_number(req)" in so_df.columns:
        so_df["customer_contact_number(req)"] = so_df["customer_contact_number(req)"].apply(parse_contact)

    # Ensure date column is formatted correctly
    if "Date" in so_df.columns:
        so_df["Date"] = pd.to_datetime(so_df["Date"], errors="coerce").dt.strftime("%d-%m-%Y")
        # If dates are all NaT, use delivery_date
        if so_df["Date"].isna().all():
            so_df["Date"] = delivery_date.replace("/", "-")

    # Ensure all output columns exist
    for col in OUTPUT_COLS:
        if col not in so_df.columns:
            so_df[col] = np.nan

    print(f"       Total QTY in Google Sheet before NA separation: {pd.to_numeric(so_df.get('QTY'), errors='coerce').sum()}")

    # ── Step 3: Separate valid vs NA rows ────────────────────────
    print("► [3/3] Separating valid and NA rows...")
    
    # Filter out rows if there is no FSN or no contact number (so they don't appear in NA tab)
    fsn_col = "FSN" if "FSN" in so_df.columns else "sku_id(req)"
    if fsn_col in so_df.columns and "customer_contact_number(req)" in so_df.columns:
        is_missing_fsn_or_contact = (so_df[fsn_col].fillna("").astype(str).str.strip() == "") | (so_df["customer_contact_number(req)"].fillna("").astype(str).str.strip() == "")
        so_df = so_df[~is_missing_fsn_or_contact].copy()
        
    NA_CHECK_COLS_WITH_PO = NA_CHECK_COLS + ["purchaseOrder"]
    actual_check_cols = [c for c in NA_CHECK_COLS_WITH_PO if c in so_df.columns]
    
    if actual_check_cols:
        is_null = so_df[actual_check_cols].isnull().any(axis=1)
        is_blank = (so_df[actual_check_cols].fillna("").astype(str).apply(lambda x: x.str.strip()) == "").any(axis=1)
        is_na = is_null | is_blank
    else:
        is_na = pd.Series(False, index=so_df.index)
        
    # Also NA if QTY is invalid
    if "QTY" in so_df.columns:
        is_invalid_qty = so_df["QTY"].isna() | (pd.to_numeric(so_df["QTY"], errors="coerce").fillna(0) <= 0)
        is_na = is_na | is_invalid_qty
        
    df_valid = so_df[~is_na][OUTPUT_COLS].copy()
    df_na    = so_df[is_na][OUTPUT_COLS].copy()

    print(f"       ✓ Valid rows : {len(df_valid)}")
    print(f"       ✗ NA rows    : {len(df_na)}")
    print(f"       ✓ Valid QTY  : {pd.to_numeric(df_valid.get('QTY'), errors='coerce').sum()}")
    print(f"       ✗ NA QTY     : {pd.to_numeric(df_na.get('QTY'), errors='coerce').sum()}")

    # ── Output ───────────────────────────────────────────────────
    from datetime import datetime
    try:
        dt = datetime.strptime(delivery_date, "%d-%m-%Y")
        date_tag = dt.strftime("%B %d")
    except:
        date_tag = delivery_date.replace("-", " ")

    base_name = f"{city} XD SO {date_tag}"

    csv_path  = os.path.join(output_dir, f"{base_name}.csv")
    xlsx_path = os.path.join(output_dir, f"{base_name}_full.xlsx")
    po_path   = os.path.join(output_dir, f"{city} XD PO Mapping {date_tag}.xlsx")

    # Main CSV — valid rows only
    df_valid.to_csv(csv_path, index=False)

    # Excel — two tabs: valid + NA
    with pd.ExcelWriter(xlsx_path, engine="openpyxl") as writer:
        df_valid.to_excel(writer, sheet_name="SO Output", index=False)
        if len(df_na) > 0:
            df_na.to_excel(writer, sheet_name="NA Rows", index=False)
        else:
            pd.DataFrame({"Message": ["No NA rows found"]}).to_excel(
                writer, sheet_name="NA Rows", index=False
            )

    # PO Key mapping reference
    key_table.to_excel(po_path, index=False)

    print(f"\n  ✅ CSV saved      → {csv_path}")
    print(f"  ✅ Excel saved    → {xlsx_path}")
    print(f"  ✅ PO Map saved   → {po_path}\n")

    return csv_path, xlsx_path, po_path, len(df_valid), len(df_na), len(so_df), len(unique_keys)

if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="E-com GRO SO Automation Tool")
    parser.add_argument("--allocation", required=True, help="Allocation Excel path")
    parser.add_argument("--ecom",       required=False, default=None, help="E_COM_SO_Placement.xlsx path (if not using G-sheet)")
    parser.add_argument("--gsheet",     required=False, default=None, help="Public Google Sheet URL (preferred)")
    parser.add_argument("--city",       required=True, choices=list(CITY_CONFIG.keys()))
    parser.add_argument("--date",       required=True, help="Delivery date e.g. 01-07-2026")
    parser.add_argument("--out",        default=".", help="Output directory")
    args = parser.parse_args()

    if not args.ecom and not args.gsheet:
        parser.error("Provide either --ecom (local file) or --gsheet (Google Sheet URL)")

    run_automation(
        allocation_path=args.allocation,
        ecom_path=args.ecom,
        city=args.city,
        delivery_date=args.date,
        output_dir=args.out,
        gsheet_url=args.gsheet,
    )
