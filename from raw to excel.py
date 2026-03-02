import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
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

# פונקציה לעיבוד קובץ ALS
def parse_als_file(file_bytes):
    from openpyxl import load_workbook
    wb = load_workbook(io.BytesIO(file_bytes), data_only=True)
    sheet = next((wb[n] for n in wb.sheetnames if "Client" in n and "SOIL" in n), wb.active)
    rows = list(sheet.iter_rows(values_only=True))
    
    # חיפוש שורות כותרת
    try:
        sample_row_idx = next(i for i, r in enumerate(rows) if any("Client Sample ID" in str(v) for v in r if v))
        param_row_idx = next(i for i, r in enumerate(rows) if r and r[0] == "Parameter")
    except StopIteration:
        return pd.DataFrame()

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

# --- ממשק המשתמש ---
try:
    st.set_page_config(page_title="דוח מאוחד", layout="wide")
    st.title("🧪 מערכת עיבוד תוצאות קרקע")

    with st.sidebar:
        st.header("📂 טעינת ערכי סף")
        f_main = st.file_uploader("1. ערכי סף כלליים", type=["csv", "xlsx"])
        f_pfas = st.file_uploader("2. ערכי סף PFAS", type=["csv", "xlsx"])
        t1_name = st.text_input("שם עמודת Tier 1", "Tier 1 industrial A 0-6m")
        st.markdown("---")
        als_files = st.file_uploader("3. קבצי ALS", type=["xlsx"], accept_multiple_files=True)

    if (f_main or f_pfas) and als_files:
        thresh_map = {}
        for f in [f_main, f_pfas]:
            if f:
                df_t = pd.read_csv(f) if f.name.endswith('csv') else pd.read_excel(f)
                df_t.columns = [c.strip() for c in df_t.columns]
                chem_c = next((c for c in df_t.columns if c.lower() in ["chemical", "parameter", "name", "תרכובת"]), None)
                vsl_c = next((c for c in df_t.columns if "VSL" in c), None)
                t1_c = t1_name if t1_name in df_t.columns else next((c for c in df_t.columns if "Tier 1" in c), None)
                cas_c = next((c for c in df_t.columns if "CAS" in c), None)
                if chem_c:
                    for _, r in df_t.iterrows():
                        thresh_map[norm_name(r[chem_c])] = {"vsl": r.get(vsl_c), "tier1": r.get(t1_c), "cas": r.get(cas_c)}

        all_data = pd.concat([parse_als_file(f.read()) for
