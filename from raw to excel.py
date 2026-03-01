import streamlit as st
import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import io
import re

st.set_page_config(page_title="דוח סקר קרקע", layout="wide", page_icon="🧪")
st.title("🧪 מערכת עיבוד תוצאות מעבדה")
st.markdown("---")

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

def parse_als_file(file_bytes, filename):
    try:
        wb = load_workbook(io.BytesIO(file_bytes), data_only=True)
    except Exception as e:
        return None, f"לא ניתן לפתוח: {e}"

    main_sheet = None
    for name in wb.sheetnames:
        if "Client" in name and "SOIL" in name:
            main_sheet = wb[name]
            break
    if main_sheet is None:
        main_sheet = wb[wb.sheetnames[0]]

    all_rows = list(main_sheet.iter_rows(values_only=True))

    sample_row_idx = None
    for i, row in enumerate(all_rows):
        if any("Client Sample ID" in str(v) for v in row if v):
            sample_row_idx = i
            break
    if sample_row_idx is None:
        return None, "לא נמצאה שורת Sample IDs"

    sample_row = all_rows[sample_row_idx]
    col_to_sample = {}
    for col_idx, val in enumerate(sample_row):
        if val and val != "Client Sample ID":
            col_to_sample[col_idx] = str(val).strip()

    header_row_idx = None
    for i, row in enumerate(all_rows):
        if row[0] == "Parameter":
            header_row_idx = i
            break
    if header_row_idx is None:
        return None, "לא נמצאה שורת Parameter"

    records = []
    current_group = "Unknown"

    for row in all_rows[header_row_idx + 1:]:
        param  = row[0]
        method = row[1]
        unit   = row[2]

        if not param:
            continue
        if param and not method and not unit:
            current_group = str(param).strip()
            continue

        for col_idx, sample_name in col_to_sample.items():
            val = row[col_idx] if col_idx < len(row) else None

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

# Sidebar
st.sidebar.header("⚙️ הגדרות")
tier1_options = {"industrial": "תעשייתי", "residential": "מגורים", "recreational": "נופש"}
selected_tier1 = st.sidebar.selectbox(
    "בחר TIER 1:",
    options=list(tier1_options.keys()),
    format_func=lambda x: f"{x.capitalize()} ({tier1_options[x]})"
)
st.sidebar.markdown("---")
st.sidebar.markdown("🟡 חריגה מעל VSL")
st.sidebar.markdown("🟠 חריגה מעל TIER 1")

# העלאת קבצים
col1, col2 = st.columns(2)
with col1:
    st.subheader("📁 קובץ ערכי סף")
    threshold_file = st.file_uploader("העלה קובץ ערכי סף (Excel)", type=["xlsx","xls"], key="thresholds")
with col2:
    st.subheader("📂 קבצי נתונים מ-ALS")
    data_files = st.file_uploader("העלה קבצי ALS — אפשר כמה ביחד", type=["xlsx","xls"], accept_multiple_files=True, key="data")

if threshold_file and data_files:
    st.markdown("---")

    try:
        df_thresh = pd.read_excel(threshold_file)
        df_thresh.columns = [str(c).lower().strip() for c in df_thresh.columns]

        with st.expander("🔍 עמודות בקובץ ערכי הסף"):
            st.write(list(df_thresh.columns))

        comp_col  = next((c for c in df_thresh.columns if c in ["compound","parameter","תרכובת","name"]), None)
        vsl_col   = next((c for c in df_thresh.columns if "vsl" in c), None)
        tier1_col = next((c for c in df_thresh.columns if selected_tier1 in c.lower()), None)
        if tier1_col is None:
            tier1_col = next((c for c in df_thresh.columns if "tier" in c and "1" in c), None)

        if not comp_col or not vsl_col:
            st.error(f"לא נמצאו עמודות נדרשות. נמצא: {list(df_thresh.columns)}")
            st.stop()

        thresh_dict = {}
        for _, row in df_thresh.iterrows():
            name = str(row[comp_col]).strip().lower()
            thresh_dict[name] = {
                "vsl":   row.get(vsl_col),
                "tier1": row.get(tier1_col) if tier1_col else None
            }
        st.success(f"✅ נטענו {len(thresh_dict)} ערכי סף")

    except Exception as e:
        st.error(f"שגיאה בקובץ ערכי סף: {e}")
        st.stop()

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

    groups = sorted(df_all["group"].unique())
    wb_out = Workbook()
    wb_out.remove(wb_out.active)
    stats = {"total": 0, "vsl": 0, "tier1": 0}

    for group in groups:
        df_g  = df_all[df_all["group"] == group].copy()
        ws    = wb_out.create_sheet(title=str(group)[:31])
        headers = ["שם קידוח", "עומק (מ')", "תרכובת", "יחידות", "תוצאה", "VSL", f"TIER 1 ({selected_tier1})", "סטטוס"]

        ws.merge_cells(f"A1:{get_column_letter(len(headers))}1")
        c = ws["A1"]
        c.value = f"תוצאות — {group}"
        style_header(c, 1)
        ws.row_dimensions[1].height = 22

        for ci, h in enumerate(headers, 1):
            c = ws.cell(row=2, column=ci, value=h)
            style_header(c, 2)
        ws.row_dimensions[2].height = 30

        for ri, (_, row) in enumerate(df_g.iterrows(), start=3):
            compound_key = str(row["compound"]).strip().lower()
            result_val   = row["result"]
            thresh        = thresh_dict.get(compound_key, {})
            vsl           = thresh.get("vsl")
            tier1         = thresh.get("tier1")

            highlight = None
            status    = "תקין"
            try:
                r = float(result_val)
                if tier1 and pd.notna(tier1) and float(tier1) > 0 and r > float(tier1):
                    highlight = "tier1"
                    status    = "⚠️ חריגת TIER 1"
                    stats["tier1"] += 1
                elif vsl and pd.notna(vsl) and float(vsl) > 0 and r > float(vsl):
                    highlight = "vsl"
                    status    = "⚡ חריגת VSL"
                    stats["vsl"] += 1
            except:
                pass

            stats["total"] += 1
            values = [
                row["sample_id"], row["depth"], row["compound"], row["unit"],
                row.get("result_str", result_val),
                vsl   if vsl   and pd.notna(vsl)   else "—",
                tier1 if tier1 and pd.notna(tier1) else "—",
                status
            ]
            for ci, val in enumerate(values, 1):
                c = ws.cell(row=ri, column=ci, value=val)
                style_data(c, highlight if ci in [5, 8] else None)

        for ci, w in enumerate([14,10,28,10,12,12,16,18], 1):
            ws.column_dimensions[get_column_letter(ci)].width = w
        ws.freeze_panes = "A3"

    st.markdown("---")
    c1, c2, c3 = st.columns(3)
    c1.metric("סה״כ תוצאות", stats["total"])
    c2.metric("🟡 חריגות VSL",    stats["vsl"])
    c3.metric("🟠 חריגות TIER 1", stats["tier1"])

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

elif not threshold_file and not data_files:
    st.info("👆 העלה קובץ ערכי סף וקבצי ALS כדי להתחיל")
elif not threshold_file:
    st.warning("⚠️ חסר קובץ ערכי סף")
else:
    st.warning("⚠️ חסרים קבצי נתונים")
