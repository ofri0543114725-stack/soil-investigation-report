import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import io
import re

# הגדרות עיצוב
YELLOW_FILL = PatternFill("solid", fgColor="FFFF00")
ORANGE_FILL = PatternFill("solid", fgColor="FFA500")
BLUE_HEADER = PatternFill("solid", fgColor="1F4E79")
LIGHT_BLUE_HEADER = PatternFill("solid", fgColor="2E75B6")
WHITE_FONT = Font(color="FFFFFF", bold=True, name="Arial", size=10)

def make_border():
    s = Side(style="thin", color="000000")
    return Border(left=s, right=s, top=s, bottom=s)

def natural_sort_key(s):
    return [int(text) if text.isdigit() else text.lower() for text in re.split('([0-9]+)', str(s))]

def norm_name(s):
    return re.sub(r"\s+", " ", str(s).strip().lower()) if s else ""

# פונקציה לקריאת קבצי ה-ALS
def parse_als_file(file_bytes):
    from openpyxl import load_workbook
    wb = load_workbook(io.BytesIO(file_bytes), data_only=True)
    sheet = next((wb[n] for n in wb.sheetnames if "Client" in n and "SOIL" in n), wb.active)
    rows = list(sheet.iter_rows(values_only=True))
    sample_row_idx = next(i for i, r in enumerate(rows) if any("Client Sample ID" in str(v) for v in r if v))
    param_row_idx = next(i for i, r in enumerate(rows) if r and r[0] == "Parameter")
    samples = {i: str(v).strip() for i, v in enumerate(rows[sample_row_idx]) if v and v != "Client Sample ID"}
    
    records = []
    current_group = "Unknown"
    for row in rows[param_row_idx + 1:]:
        param, unit = row[0], row[2]
        if not param: continue
        if not unit: current_group = str(param).strip(); continue
        for col_idx, sname in samples.items():
            val = row[col_idx]
            match = re.match(r"^(S-?\d+[A-Za-z]*)\s*\(([0-9.]+)\)", sname)
            sid, depth = (match.group(1), float(match.group(2))) if match else (sname, 0.0)
            res_str = str(val).strip() if val is not None else ""
            res_num = 0.0 if res_str.startswith("<") else (float(res_str) if res_str.replace('.','',1).isdigit() else None)
            records.append({
                "sample_id": sid, "depth": depth, "compound": str(param).strip(),
                "unit": unit, "result": res_num, "result_str": res_str, "group": current_group
            })
    return pd.DataFrame(records)

# --- ממשק Streamlit ---
st.title("🧪 הפקת דוח קרקע מאוחד - 4 לשוניות")

col1, col2 = st.columns(2)
with col1:
    thresh_file = st.file_uploader("העלה קובץ ערכי סף (גרסה 7)", type=["csv", "xlsx"])
with col2:
    als_files = st.file_uploader("העלה קבצי ALS", type=["xlsx"], accept_multiple_files=True)

if thresh_file and als_files:
    df_thresh = pd.read_csv(thresh_file) if thresh_file.name.endswith('csv') else pd.read_excel(thresh_file)
    df_thresh.columns = [c.strip() for c in df_thresh.columns]
    
    # זיהוי עמודות בקובץ הסף
    cas_col = next((c for c in df_thresh.columns if "CAS" in c), None)
    chem_col = next((c for c in df_thresh.columns if c.lower() in ["chemical", "parameter", "תרכובת"]), None)
    vsl_col = next((c for c in df_thresh.columns if "VSL" in c), None)
    tier1_col = st.sidebar.selectbox("בחר עמודת TIER 1 (כתום)", [c for c in df_thresh.columns if "Tier 1" in c])
    
    thresh_map = {norm_name(r[chem_col]): {"vsl": r.get(vsl_col), "tier1": r.get(tier1_col), "cas": r.get(cas_col)} for _, r in df_thresh.iterrows()}

    all_data = pd.concat([parse_als_file(f.read()) for f in als_files], ignore_index=True)
