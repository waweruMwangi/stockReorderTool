import os
import uuid
import io
import pandas as pd
import traceback
from flask import Flask, request, send_file, after_this_request, render_template

app = Flask(__name__)

UPLOAD_FOLDER = os.path.join(os.getcwd(), 'uploads')
os.makedirs(UPLOAD_FOLDER, exist_ok=True)


# -------------------------------
# HELPERS
# -------------------------------
def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in {
        'xls', 'xlsx', 'csv'
    }


# -------------------------------
# COLUMN NORMALIZER
# -------------------------------
def normalize_columns(df):
    df = df.copy()

    df.columns = (
        pd.Index(df.columns)
        .astype(str)
        .str.strip()
        .str.upper()
        .str.replace(r'\s+', ' ', regex=True)
    )

    rename_map = {
        "ITEM": "ITEM NAME",
        "ITEMNAME": "ITEM NAME",
        "PRODUCT": "ITEM NAME",
        "NAME": "ITEM NAME",
        "ITEM DESCRIPTION": "ITEM NAME",
        "QUANTITY": "QTY"
    }

    return df.rename(columns=rename_map)
def clean_item_names(df):
    df = df.copy()

    df['ITEM NAME'] = (
        df['ITEM NAME']
        .astype(str)
        .str.upper()
        .str.strip()

        # remove prefix up to first hyphen only
        .str.replace(r'^.{1,10}-\s*', '', regex=True)


        # normalize multiple spaces
        .str.replace(r'\s+', ' ', regex=True)
        .str.strip()
    )

    return df
# -------------------------------
# DETECT HEADER ROW (SAFE)
# -------------------------------
def detect_header(df):
    df = df.copy()

    for i in range(min(10, len(df))):
        row = df.iloc[i].fillna("").astype(str).str.lower().tolist()

        has_item = any("item" in x for x in row)
        has_qty = any("qty" in x or "quantity" in x for x in row)

        if has_item and has_qty:
            df.columns = df.iloc[i].astype(str)
            df = df[i + 1:].copy()
            return df

    return df


# -------------------------------
# SMART EXCEL READER (FIXED)
# -------------------------------
def read_excel(file_storage, header_row=0):
    content = file_storage.read()
    file_storage.seek(0)

    bio = io.BytesIO(content)

    # ALWAYS TRY OPENPYXL FIRST (SAFE FOR XLSX)
    try:
        return pd.read_excel(bio, engine="openpyxl", header=header_row)
    except Exception:
        # fallback without specifying engine
        bio.seek(0)
        return pd.read_excel(bio, header=header_row)


# -------------------------------
# CSV READER
# -------------------------------
def read_csv(file_storage, header_row=0):
    content = file_storage.read()
    file_storage.seek(0)

    bio = io.BytesIO(content)

    df = pd.read_csv(
        bio,
        sep=None,
        engine="python",
        header=header_row,
        encoding_errors="ignore",
        dtype=str
    )

    return df


# -------------------------------
# XML EXCEL READER (Excel 2003 XML)
# -------------------------------
import xml.etree.ElementTree as ET

def read_excel_xml(bio):
    bio.seek(0)
    tree = ET.parse(bio)
    root = tree.getroot()

    # Excel XML namespace
    ns = {'ss': 'urn:schemas-microsoft-com:office:spreadsheet'}

    rows = []

    # 🔥 Find actual worksheet table
    for worksheet in root.findall('.//ss:Worksheet', ns):
        table = worksheet.find('.//ss:Table', ns)
        if table is None:
            continue

        for row in table.findall('ss:Row', ns):
            row_data = []
            for cell in row.findall('ss:Cell', ns):
                data = cell.find('ss:Data', ns)
                row_data.append(data.text if data is not None else "")
            rows.append(row_data)

        break  # only first worksheet

    df = pd.DataFrame(rows)

    return df


# -------------------------------
# MAIN READER (FIXED)
# -------------------------------
def read_file(file_storage, header_row=0):
    content = file_storage.read()
    file_storage.seek(0)

    if not content:
        raise ValueError("Empty file uploaded")

    bio = io.BytesIO(content)

    # XLSX
    if content[:2] == b'PK':
        bio.seek(0)
        return pd.read_excel(bio, engine="openpyxl", header=header_row)

    # XLS
    if content[:4] == b'\xD0\xCF\x11\xE0':
        bio.seek(0)
        return pd.read_excel(bio, engine="xlrd", header=header_row)

    # 🔥 XML Excel
    if content.strip().startswith(b'<?xml') or b'<Workbook' in content[:200]:
        bio.seek(0)
        return read_excel_xml(bio)

    # CSV fallback
    bio.seek(0)
    return pd.read_csv(
        bio,
        sep=None,
        engine="python",
        header=header_row,
        encoding_errors="ignore",
        dtype=str,
        on_bad_lines='skip',
        quoting=3
    )

# -------------------------------
# ROUTES
# -------------------------------
@app.route('/')
def index():
    return render_template('index.html')


@app.route('/process', methods=['POST'])
def process():

    try:
        # ---------------- VALIDATION ----------------
        if 'sales' not in request.files or 'stock' not in request.files:
            return "Upload both sales and stock files.", 400

        sales_file = request.files['sales']
        stock_file = request.files['stock']

        if not allowed_file(sales_file.filename) or not allowed_file(stock_file.filename):
            return "Only Excel or CSV files allowed.", 400

        # ---------------- HEADER ROW RULES ----------------
        # sales headers row 2 -> index 1
        # stock headers row 3 -> index 2

        df_sales = clean_item_names(
            normalize_columns(detect_header(read_file(sales_file, header_row=1)))
        )

        df_stock = clean_item_names(
            normalize_columns(detect_header(read_file(stock_file, header_row=2)))
        )
        
        # ---------------- DEBUG ----------------
        print("SALES COLUMNS:", df_sales.columns.tolist())
        print("STOCK COLUMNS:", df_stock.columns.tolist())

        # ---------------- VALIDATION ----------------
        required = {'ITEM NAME'}

        if not required.issubset(df_sales.columns):
            return f"Sales file missing required columns. Found: {list(df_sales.columns)}", 400

        if 'ITEM NAME' not in df_stock.columns or 'AVAIL. QTY' not in df_stock.columns:
            return f"Stock file missing required columns. Found: {list(df_stock.columns)}", 400

        # ---------------- CLEAN ----------------
        df_sales['QTY'] = pd.to_numeric(df_sales['QTY'], errors='coerce').fillna(0)
        # CLEAN STOCK QTY
        df_stock['AVAIL. QTY'] = pd.to_numeric(
            df_stock['AVAIL. QTY'],
            errors='coerce'
        ).fillna(0)

        # 🔥 AGGREGATE STOCK PER ITEM (IMPORTANT FIX)
        df_stock = (
            df_stock
            .groupby('ITEM NAME', as_index=False)['AVAIL. QTY']
            .sum()
            .rename(columns={'AVAIL. QTY': 'CURRENT_STOCK'})
        )
        # ---------------- PROCESS ----------------
        sales_summary = df_sales.groupby('ITEM NAME', as_index=False)['QTY'].sum()


        merged = pd.merge(
            sales_summary,
            df_stock[['ITEM NAME', 'CURRENT_STOCK']],
            on='ITEM NAME',
            how='left'
        )

        merged['CURRENT_STOCK'] = merged['CURRENT_STOCK'].fillna(0)
        merged['REORDER_QTY'] = (merged['QTY'] - merged['CURRENT_STOCK']).clip(lower=0)

        # 🔥 SORT BY SALES QTY (DESCENDING)
        merged = merged.sort_values(by='QTY', ascending=False)

        # ---------------- OUTPUT ----------------
        filename = f"reorder_{uuid.uuid4().hex[:8]}.xlsx"
        path = os.path.join(UPLOAD_FOLDER, filename)

        merged.to_excel(path, index=False)

        from openpyxl import load_workbook

        wb = load_workbook(path)
        ws = wb.active

        # make ITEM NAME column wider
        for col in ws.columns:
            max_length = 0
            col_letter = col[0].column_letter

            for cell in col:
                try:
                    if cell.value:
                        max_length = max(max_length, len(str(cell.value)))
                except:
                    pass

            # special boost for ITEM NAME column
            if col[0].value == "ITEM NAME":
                ws.column_dimensions[col_letter].width = 100
            else:
                ws.column_dimensions[col_letter].width = max_length + 5

        wb.save(path)

        @after_this_request
        def cleanup(response):
            try:
                if os.path.exists(path):
                    os.remove(path)
            except Exception:
                pass
            return response

        return send_file(path, as_attachment=True)

    except Exception:
        error_trace = traceback.format_exc()
        print(error_trace)
        return f"<pre>{error_trace}</pre>", 500


# -------------------------------
# RUN
# -------------------------------
import os

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)