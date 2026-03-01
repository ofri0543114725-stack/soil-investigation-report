
import streamlit as st
import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import io
import re

# --------------------------------------------------
# UI CONFIG
# --------------------------------------------------
st.set_page_config(page_title="דוח סקר קרקע", layout="wide", page_icon="🧪")
st.title("🧪 מערכת עיבוד תוצאות מעבדה")
st.markdown("---")

# --------------------------------------------------
# STYLES
# --------------------------------------------------
YELLOW_FILL  = PatternFill("solid", fgColor="FFFF00")
ORANGE_FILL  = PatternFill("solid", fgColor="FFA500")
HEADER_FILL  = PatternFill("solid", fgColor="1F4E79")
SUBHEAD_FILL = PatternFill("solid", fgColor="2E75B6")
WHITE_FONT   = Font(color="FFFFFF", bold=True, name="Arial", size=10)
BOLD_FONT    = Font(bold=True, name="Arial", size=10)
NORMAL_FONT  = Font(name="Arial", size=10)

def make_border():
    s = Side(style="thin", color="000000")
    return Border(left=s, right=s, top=s, bottom=s)

def style_header(cell, level=1):
    cell.font = WHITE_FONT
    cell.fill = HEADER_FILL if level == 1 else SUBHEAD_FILL
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border = make_border()

def style_data(cell, highlight=None):
    cell.font = BOLD_FONT if highlight else NORMAL_FONT
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.border = make_border()
    if highlight == "tier1":
        cell.fill = ORANGE_FILL
    elif highlight == "vsl":
        cell.fill = YELLOW_FILL

# --------------------------------------------------
# HELPERS
# --------------------------------------------------
def norm_name(s):
    s = "" if s is None else str(s)
    s = s.strip().lower()
    s = s.replace("\xa0", " ").replace("–", "-").replace("—", "-")
    s = re.sub(r"\s+", " ", s)
    return s

def safe_sheet_title(name, existing):
    s = "" if name is None else str(name)
    s = re.sub(r'[:\\/?*\[\]]', '-', s)
    s = s.replace('\n',' ').replace('\r',' ').strip()
    if not s:
        s = "Sheet"
    base = s[:28]
    title = base
    i = 1
    while title in existing or len(title) > 31:
        title = f"{base}-{i}"
        i += 1
    return title[:31]

# --------------------------------------------------
# ALS PARSER
# --------------------------------------------------
def parse_als_file(file_bytes, filename):
    try:
        wb = load_workbook(io.BytesIO(file_bytes), data_only=True)
    except Exception as e:
        return None, f"שגיאה בפתיחת הקובץ: {e}"

    ws = wb[wb.sheetnames[0]]
    rows = list(ws.iter_rows(values_only=True))

    header_row = None
    for i, r in enumerate(rows):
        if r and "Client Sample ID" in str(r):
            header_row = i
            break
    if header_row is None:
        return None, "לא נמצאה שורת Sample IDs"

    sample_row = rows[header_row]
    col_map = {}
    for i, val in enumerate(sample_row):
        if val and val != "Client Sample ID":
            col_map[i] = str(val).strip()

    param_row = None
    for i, r in enumerate(rows):
        if r and len(r) > 0 and r[0] == "Parameter":
            param_row = i
            break
    if param_row is None:
        return None, "לא נמצאה שורת Parameter"

    data = []
    group = "Unknown"

    for r in rows[param_row+1:]:
        if not r or not r[0]:
            continue

        if r[0] and not r[1] and not r[2]:
            group = str(r[0]).strip()
            continue

        for col_idx, sample in col_map.items():
            val = r[col_idx] if col_idx < len(r) else None
            result_str = str(val).strip() if val is not None else ""

            if result_str.startswith("<"):
                result = 0.0
            else:
                try:
                    result = float(result_str)
                except:
                    result = None

            if result is not None:
                data.append({
                    "sample_id": sample,
                    "compound": r[0],
                    "unit": r[2] if len(r) > 2 else "mg/kg",
                    "result": result,
                    "result_str": result_str,
                    "group": group
                })

    if not data:
        return None, "לא נמצאו נתונים"

    return pd.DataFrame(data), None

# --------------------------------------------------
# SIDEBAR
# --------------------------------------------------
st.sidebar.header("⚙️ הגדרות")
selected_tier = st.sidebar.selectbox("בחר TIER 1:", ["industrial","residential"])

# --------------------------------------------------
# FILE UPLOAD
# --------------------------------------------------
col1, col2 = st.columns(2)

with col1:
    threshold_file = st.file_uploader("קובץ ערכי סף", type=["xlsx"])

with col2:
    data_files = st.file_uploader("קבצי ALS", type=["xlsx"], accept_multiple_files=True)

# --------------------------------------------------
# MAIN LOGIC
# --------------------------------------------------
if threshold_file and data_files:

    df_thresh = pd.read_excel(threshold_file)
    df_thresh.columns = [str(c).lower().strip() for c in df_thresh.columns]

    comp_col = next((c for c in df_thresh.columns if c in ["chemical","compound","parameter"]), None)
    vsl_col = next((c for c in df_thresh.columns if "vsl" in c), None)
    tier_col = next((c for c in df_thresh.columns if selected_tier in c and "tier 1" in c), None)

    if not comp_col or not vsl_col:
        st.error("עמודות chemical / vsl לא נמצאו בקובץ הסף")
        st.stop()

    thresh_dict = {}
    for _, r in df_thresh.iterrows():
        key = norm_name(r[comp_col])
        thresh_dict[key] = {
            "vsl": r[vsl_col],
            "tier1": r[tier_col] if tier_col else None
        }

    all_data = []
    for f in data_files:
        df, err = parse_als_file(f.read(), f.name)
        if err:
            st.warning(err)
        else:
            all_data.append(df)

    if not all_data:
        st.stop()

    df_all = pd.concat(all_data)

    wb_out = Workbook()
    if wb_out.active:
        wb_out.remove(wb_out.active)

    existing_titles = set()

    for group in df_all["group"].unique():
        df_g = df_all[df_all["group"] == group]

        title = safe_sheet_title(group, existing_titles)
        existing_titles.add(title)

        ws = wb_out.create_sheet(title=title)

        headers = ["קידוח","תרכובת","תוצאה","VSL","TIER 1","סטטוס"]

        for i,h in enumerate(headers,1):
            c = ws.cell(row=1,column=i,value=h)
            style_header(c,2)

        row_i = 2
        for _,r in df_g.iterrows():
            key = norm_name(r["compound"])
            thresh = thresh_dict.get(key,{})
            vsl = thresh.get("vsl")
            tier = thresh.get("tier1")

            status = "תקין"
            highlight = None

            try:
                if tier and r["result"] > float(tier):
                    status="חריגת TIER 1"
                    highlight="tier1"
                elif vsl and r["result"] > float(vsl):
                    status="חריגת VSL"
                    highlight="vsl"
            except:
                pass

            values=[r["sample_id"],r["compound"],r["result"],vsl,tier,status]

            for col_i,val in enumerate(values,1):
                c=ws.cell(row=row_i,column=col_i,value=val)
                style_data(c,highlight if col_i==3 else None)

            row_i+=1

    buf=io.BytesIO()
    wb_out.save(buf)
    buf.seek(0)

    st.download_button(
        "⬇️ הורד קובץ Excel",
        data=buf,
        file_name="soil_report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

else:
    st.info("העלה קובץ ערכי סף וקבצי ALS")
