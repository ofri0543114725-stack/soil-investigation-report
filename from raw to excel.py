import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
import io
import re

# חובה: פקודה ראשונה
st.set_page_config(page_title="מערכת עיבוד תוצאות קרקע", layout="wide")

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
    """מיון לוגי של קידוחים (S329 לפני S3)"""
    return [int(text) if text.isdigit() else text.lower() for text in re.split('([0-9]+)', str(s))]

def norm_name(s):
    return re.sub(r"\s+", " ", str(s).strip().lower()) if s else ""

def parse_als_data(file_bytes):
    """סורק את כל הגיליונות ב-ALS ושולף נתונים ללא סינון מוקדם"""
    all_records = []
    try:
        from openpyxl import load_workbook
        wb = load_workbook(io.BytesIO(file_bytes), data_only=True)
        for sheet in wb.worksheets:
            rows = list(sheet.iter_rows(values_only=True))
            s_idx = next((i for i, r in enumerate(rows) if any("Client Sample ID" in str(v) for v in r if v)), None)
            p_idx = next((i for i, r in enumerate(rows) if r and r[0] == "Parameter"), None)
            
            if s_idx is not None and p_idx is not None:
                samples = {idx: str(v).strip() for idx, v in enumerate(rows[s_idx]) if v and v != "Client Sample ID"}
                current_group = "General"
                for row in rows[p_idx + 1:]:
                    param, unit = row[0], row[2]
                    if not param: continue
                    if not unit: current_group = str(param).strip(); continue
                    for col_idx, sname in samples.items():
                        val = row[col_idx]
                        match = re.match(r"^(S-?\d+[A-Za-z]*)\s*\(([0-9.]+)\)", sname)
                        sid, depth = (match.group(1), float(match.group(2))) if match else (sname, 0.0)
                        res_str = str(val).strip() if val is not None else ""
                        res_num = 0.0 if res_str.startswith("<") else (float(res_str) if res_str.replace('.','',1).isdigit() else None)
                        all_records.append({
                            "sample_id": sid, "depth": depth, "compound": str(param).strip(),
                            "unit": unit, "result": res_num, "result_str": res_str, "group": current_group
                        })
    except: pass
    return pd.DataFrame(all_records)

# --- ממשק המשתמש ---
st.title("🧪 הפקת דוח מאוחד - 4 לשוניות")

with st.sidebar:
    st.header("📂 טעינת ערכי סף")
    f_main = st.file_uploader("1. ערכי סף כלליים (גרסה 7)", type=["csv", "xlsx"])
    f_pfas = st.file_uploader("2. ערכי סף PFAS", type=["csv",
