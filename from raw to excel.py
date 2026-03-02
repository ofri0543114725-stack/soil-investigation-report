import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
import io
import re

# חובה: הפקודה הראשונה של Streamlit בקוד
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
    """מיון לוגי של קידוחים (למשל S329 לפני S3)"""
    return [int(text) if text.isdigit() else text.lower() for text in re.split('([0-9]+)', str(s))]

def norm_name(s):
    return re.sub(r"\s+", " ", str(s).strip().lower()) if s else ""

def parse_als_file(file_bytes):
    from openpyxl import load_workbook
    wb = load_workbook(io.BytesIO(file_bytes), data_only=True)
    sheet = next((wb[n] for n in wb.sheetnames if "Client" in n and "SOIL" in n), wb.active)
    rows = list(sheet.iter_rows(values_only=True))
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

# תחילת הממשק
try:
    st.title("🧪 מערכת עיבוד תוצאות קרקע - דוח מאוחד")

    with st.sidebar:
        st.header("📂 טעינת ערכי סף")
        f_main = st.file_uploader("1. ערכי סף כלליים (גרסה 7)", type=["csv", "xlsx"])
        f_pfas = st.file_uploader("2. ערכי סף PFAS", type=["csv", "xlsx"])
        tier1_name = st.text_input("שם עמודת Tier 1 לצביעה", "Tier 1 industrial A 0-6m")
        st.markdown("---")
        als_files = st.file_uploader("3. העלה קבצי ALS", type=["xlsx"], accept_multiple_files=True)

    if (f_main or f_pfas) and als_files:
        thresh_map = {}
        # איחוד נתונים מ-2 קבצי ערכי הסף
        for f in [f_main, f_pfas]:
            if f:
                df_t = pd.read_csv(f) if f.name.endswith('csv') else pd.read_excel(f)
                df_t.columns = [c.strip() for c in df_t.columns]
                chem_c = next((c for c in df_t.columns if c.lower() in ["chemical", "parameter", "name", "תרכובת"]), None)
                vsl_c = next((c for c in df_t.columns if "VSL" in c), None)
                t1_c = tier1_name if tier1_name in df_t.columns else next((c for c in df_t.columns if "Tier 1" in c), None)
                cas_c = next((c for c in df_t.columns if "CAS" in c), None)
                if chem_c:
                    for _, r in df_t.iterrows():
                        thresh_map[norm_name(r[chem_c])] = {"vsl": r.get(vsl_c), "tier1": r.get(t1_c), "cas": r.get(cas_c)}

        all_data = pd.concat([parse_als_file(f.read()) for f in als_files], ignore_index=True)
        wb = Workbook()
        wb.remove(wb.active)

        # הגדרת לשוניות
        METALS = ['Al', 'As', 'Ba', 'Be', 'B', 'Cd', 'Ca', 'Cr', 'Co', 'Cu', 'Fe', 'Pb', 'Li', 'Mg', 'Mn', 'Hg', 'Ni', 'K', 'Na', 'Se', 'Ag', 'V', 'Zn']
        TPH = ['TPH DRO', 'TPH ORO', 'Total TPH']

        def create_wide_sheet(name, analytes):
            ws = wb.create_sheet(name)
            # כותרות 1-5
            ws.append(["שם קידוח", "עומק", "אנליזה"] + analytes)
            ws.append(["", "", "יחידות"] + ["mg/kg"]*len(analytes))
            ws.append(["", "CAS", ""] + [thresh_map.get(norm_name(a), {}).get("cas", "-") for a in analytes])
            ws.append(["", "VSL", ""] + [thresh_map.get(norm_name(a), {}).get("vsl", "-") for a in analytes])
            ws.append(["", "TIER 1", ""] + [thresh_map.get(norm_name(a), {}).get("tier1", "-") for a in analytes])
            
            for r in range(1, 6):
                for c in range(1, len(analytes)+4):
                    cell = ws.cell(r, c)
                    cell.fill = BLUE_HEADER if r==1 else LIGHT_BLUE_HEADER
                    cell.font = WHITE_FONT; cell.border = make_border(); cell.alignment = Alignment(horizontal="center")
            
            df_sub = all_data[all_data['compound'].str.lower().isin([a.lower() for a in analytes])]
            if df_sub.empty: return
            samples = sorted(df_sub['sample_id'].unique(), key=natural_sort_key, reverse=True)
            row_idx = 6
            for sid in samples:
                df_s = df_sub[df_sub['sample_id'] == sid]
                depths = sorted(df_s['depth'].unique()); start_r = row_idx
                for d in depths:
                    ws.cell(row_idx, 1, sid if row_idx == start_r else ""); ws.cell(row_idx, 2, d)
                    for i, a in enumerate(analytes, 4):
                        res = df_s[(df_s['depth']==d) & (df_s['compound'].str.lower()==a.lower())]
                        if not res.empty:
                            c = ws.cell(row_idx, i, res['result_str'].values[0]); val = res['result'].values[0]
                            thr = thresh_map.get(norm_name(a), {})
                            try:
                                if thr.get("vsl") and val > float(thr["vsl"]): c.fill = YELLOW_FILL
                                elif thr.get("tier1") and val > float(thr["tier1"]): c.fill = ORANGE_FILL
                            except: pass
                        ws.cell(row_idx, i).border = make_border()
                    row_idx += 1
                if row_idx > start_r + 1: ws.merge_cells(start_row=start_r, start_column=1, end_row=row_idx-1, end_column=1)

        create_wide_sheet("מתכות", METALS)
        create_wide_sheet("TPH", TPH)
        # (כאן ניתן להוסיף PFAS ו-VOC בלוגיקה דומה)

        buf = io.BytesIO(); wb.save(buf)
        st.success("✅ העיבוד הושלם בהצלחה")
        st.download_button("⬇️ הורד קובץ מעובד", buf.getvalue(), "Soil_Report.xlsx", use_container_width=True)
    else:
        st.info("אנא העלה את קבצי ערכי הסף (לפחות אחד) ואת קבצי ה-ALS.")

except Exception as e:
    st.error(f"שגיאה בהרצת האפליקציה: {e}")
