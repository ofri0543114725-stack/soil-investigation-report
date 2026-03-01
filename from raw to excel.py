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

def friendly_col_label(col: str) -> str:
    """Pretty label for sidebar, but keeps original col key available."""
    c = col.lower().strip()
    c = c.replace("_", " ")
    c = re.sub(r"\s+", " ", c)
    # title-casing lightly while keeping tier / vsl readable
    if c.startswith("tier"):
        return c.upper().replace(" 1", " 1")
    if c == "vsl" or " vsl" in c:
        return c.upper()
    if c.startswith("cas"):
        return "CAS"
    if c.startswith("chemical"):
        return "Chemical"
    return c

# --------------------------------------------------
# ALS PARSER (minimal - keep your original if needed)
# --------------------------------------------------
def parse_als_file(file_bytes, filename):
    try:
        wb = load_workbook(io.BytesIO(file_bytes), data_only=True)
    except Exception as e:
        return None, f"שגיאה בפתיחת הקובץ: {e}"

    # try to find client soil sheet, else first
    main_sheet = None
    for name in wb.sheetnames:
        if "Client" in name and "SOIL" in name:
            main_sheet = wb[name]
            break
    if main_sheet is None:
        main_sheet = wb[wb.sheetnames[0]]

    rows = list(main_sheet.iter_rows(values_only=True))

    sample_row_idx = None
    for i, row in enumerate(rows):
        if any("Client Sample ID" in str(v) for v in row if v):
            sample_row_idx = i
            break
    if sample_row_idx is None:
        return None, "לא נמצאה שורת Sample IDs"

    sample_row = rows[sample_row_idx]
    col_to_sample = {}
    for col_idx, val in enumerate(sample_row):
        if val and val != "Client Sample ID":
            col_to_sample[col_idx] = str(val).strip()

    header_row_idx = None
    for i, row in enumerate(rows):
        if row and len(row) > 0 and row[0] == "Parameter":
            header_row_idx = i
            break
    if header_row_idx is None:
        return None, "לא נמצאה שורת Parameter"

    records = []
    current_group = "Unknown"

    for row in rows[header_row_idx + 1:]:
        param  = row[0] if len(row) > 0 else None
        method = row[1] if len(row) > 1 else None
        unit   = row[2] if len(row) > 2 else None

        if not param:
            continue

        # group row
        if param and not method and not unit:
            current_group = str(param).strip()
            continue

        for col_idx, sample_name in col_to_sample.items():
            val = row[col_idx] if col_idx < len(row) else None

            # depth from sample name if exists like S12 (1.5)
            match = re.match(r"^(S-?\d+[A-Za-z]*)\s*\(([0-9.]+)\)", sample_name)
            if match:
                sid   = match.group(1)
                depth = float(match.group(2))
            else:
                sid   = sample_name
                depth = None

            result_str = str(val).strip() if val is not None else ""
            if result_str.startswith("<"):
                result = 0.0
            elif result_str and result_str != "None":
                try:
                    result = float(result_str)
                except:
                    result = None
            else:
                result = None

            if result is not None:
                records.append({
                    "sample_id":  sid,
                    "depth":      depth,
                    "compound":   str(param).strip(),
                    "unit":       str(unit).strip() if unit else "mg/kg",
                    "result":     result,
                    "result_str": result_str,
                    "group":      current_group,
                    "source":     filename
                })

    if not records:
        return None, "לא נמצאו נתונים"

    return pd.DataFrame(records), None

# --------------------------------------------------
# UPLOADS
# --------------------------------------------------
col1, col2 = st.columns(2)
with col1:
    st.subheader("📁 קובץ ערכי סף")
    threshold_file = st.file_uploader("העלה קובץ ערכי סף (Excel)", type=["xlsx","xls"], key="thresholds")
with col2:
    st.subheader("📂 קבצי נתונים מ-ALS")
    data_files = st.file_uploader("העלה קבצי ALS — אפשר כמה ביחד", type=["xlsx","xls"], accept_multiple_files=True, key="data")

# --------------------------------------------------
# SIDEBAR (dynamic - depends on threshold file)
# --------------------------------------------------
st.sidebar.header("⚙️ הגדרות")

if threshold_file is None:
    st.sidebar.info("כדי לראות את כל האופציות (VSL / TIER 1 ...), קודם תעלה קובץ ערכי סף.")
    st.stop()

# read thresholds early for sidebar options
try:
    df_thresh = pd.read_excel(threshold_file)
except Exception as e:
    st.sidebar.error(f"שגיאה בקריאת קובץ ערכי סף: {e}")
    st.stop()

df_thresh.columns = [str(c).lower().strip() for c in df_thresh.columns]

# detect compound column
comp_aliases = {
    "chemical","chemical name","compound","parameter","analyte","analyte name","name",
    "תרכובת","חומר","מזהם"
}
comp_col = next((c for c in df_thresh.columns if c in comp_aliases), None)
if not comp_col:
    st.sidebar.error("לא הצלחתי לזהות עמודת שם תרכובת (Chemical/Compound/Parameter).")
    st.sidebar.write("עמודות שנמצאו:", list(df_thresh.columns))
    st.stop()

# Build list of "threshold columns" user can select from: VSL + all Tier1 columns (and optionally more)
threshold_cols = []
for c in df_thresh.columns:
    if c == comp_col:
        continue
    if "tier" in c and "1" in c:
        threshold_cols.append(c)
    elif "vsl" in c:
        threshold_cols.append(c)

if not threshold_cols:
    st.sidebar.error("לא נמצאו עמודות VSL או TIER 1 בקובץ ערכי הסף.")
    st.sidebar.write("עמודות שנמצאו:", list(df_thresh.columns))
    st.stop()

# Map pretty labels -> real col
pretty_map = {friendly_col_label(c): c for c in threshold_cols}
pretty_options = list(pretty_map.keys())

default_choice = None
for lbl, col in pretty_map.items():
    if col == "vsl" or col.endswith(" vsl") or "vsl" in col:
        default_choice = lbl
        break
if default_choice is None:
    default_choice = pretty_options[0]

compare_label = st.sidebar.selectbox(
    "בחר ערך סף להשוואה (יצבע כתום מעליו)",
    options=pretty_options,
    index=pretty_options.index(default_choice) if default_choice in pretty_options else 0
)
compare_col = pretty_map[compare_label]

# VSL column (for yellow) - allow turning on/off
vsl_candidates = [c for c in df_thresh.columns if "vsl" in c]
vsl_col = vsl_candidates[0] if vsl_candidates else None
use_vsl = st.sidebar.checkbox("להציג/לחשב חריגות VSL (צהוב)", value=(vsl_col is not None))
st.sidebar.markdown("---")
st.sidebar.markdown("🟡 חריגה מעל VSL")
st.sidebar.markdown("🟠 חריגה מעל ערך סף שנבחר למעלה")

# --------------------------------------------------
# MAIN LOGIC
# --------------------------------------------------
if not data_files:
    st.info("👆 העלה גם קבצי ALS כדי להתחיל לעבד.")
    st.stop()

# Build threshold dict: compound -> values for vsl + chosen compare_col
thresh_dict = {}
for _, r in df_thresh.iterrows():
    key = norm_name(r.get(comp_col, ""))
    if not key or key == "nan":
        continue
    thresh_dict[key] = {
        "vsl": r.get(vsl_col) if vsl_col else None,
        "compare": r.get(compare_col)
    }

st.success(f"✅ נטענו {len(thresh_dict)} ערכי סף (Compound={comp_col}, Compare={compare_label})")

# Parse ALS files
all_data = []
for f in data_files:
    df, err = parse_als_file(f.read(), f.name)
    if err:
        st.warning(f"⚠️ {f.name}: {err}")
    else:
        all_data.append(df)
        st.success(f"✅ {f.name} — {len(df)} תוצאות")

if not all_data:
    st.error("לא נטענו נתונים.")
    st.stop()

df_all = pd.concat(all_data, ignore_index=True)

with st.expander(f"👁️ תצוגה מקדימה ({len(df_all)} שורות)"):
    st.dataframe(df_all.head(30))

# Output workbook
wb_out = Workbook()
# remove default sheet if exists
if wb_out.active:
    wb_out.remove(wb_out.active)

existing_titles = set()
stats = {"total": 0, "vsl": 0, "compare": 0}

groups = sorted(df_all["group"].unique())
for group in groups:
    title = safe_sheet_title(group, existing_titles)
    existing_titles.add(title)
    ws = wb_out.create_sheet(title=title)

    headers = ["שם קידוח", "עומק (מ')", "תרכובת", "יחידות", "תוצאה", "VSL", compare_label, "סטטוס"]
    ws.merge_cells(f"A1:{get_column_letter(len(headers))}1")
    c = ws["A1"]
    c.value = f"תוצאות — {group}"
    style_header(c, 1)
    ws.row_dimensions[1].height = 22

    for ci, h in enumerate(headers, 1):
        c = ws.cell(row=2, column=ci, value=h)
        style_header(c, 2)
    ws.row_dimensions[2].height = 30

    df_g = df_all[df_all["group"] == group].copy()

    for ri, (_, row) in enumerate(df_g.iterrows(), start=3):
        compound_key = norm_name(row["compound"])
        result_val   = row["result"]

        thr = thresh_dict.get(compound_key, {})
        vsl = thr.get("vsl")
        cmpv = thr.get("compare")

        highlight = None
        status = "תקין"
        try:
            r = float(result_val)
            if cmpv is not None and pd.notna(cmpv) and float(cmpv) > 0 and r > float(cmpv):
                highlight = "tier1"  # orange
                status = f"⚠️ חריגה מעל {compare_label}"
                stats["compare"] += 1
            elif use_vsl and vsl is not None and pd.notna(vsl) and float(vsl) > 0 and r > float(vsl):
                highlight = "vsl"     # yellow
                status = "⚡ חריגת VSL"
                stats["vsl"] += 1
        except:
            pass

        stats["total"] += 1

        values = [
            row["sample_id"], row["depth"], row["compound"], row["unit"],
            row.get("result_str", result_val),
            vsl if (use_vsl and vsl is not None and pd.notna(vsl)) else "—",
            cmpv if (cmpv is not None and pd.notna(cmpv)) else "—",
            status
        ]

        for ci, val in enumerate(values, 1):
            c = ws.cell(row=ri, column=ci, value=val)
            # highlight result + status cells
            style_data(c, highlight if ci in [5, 8] else None)

    for ci, w in enumerate([14,10,28,10,12,12,22,22], 1):
        ws.column_dimensions[get_column_letter(ci)].width = w
    ws.freeze_panes = "A3"

st.markdown("---")
c1, c2, c3 = st.columns(3)
c1.metric("סה״כ תוצאות", stats["total"])
c2.metric("🟡 חריגות VSL", stats["vsl"])
c3.metric(f"🟠 חריגות מעל {compare_label}", stats["compare"])

buf = io.BytesIO()
wb_out.save(buf)
buf.seek(0)

st.download_button(
    label="⬇️ הורד קובץ Excel מעובד",
    data=buf,
    file_name="soil_report.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=True
)
