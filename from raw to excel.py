import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
import io
import re

# חובה: הפקודה הראשונה של Streamlit
st.set_page_config(page_title="מערכת עיבוד תוצאות קרקע", layout="wide")

# הגדרות עיצוב (צבעים וגופנים)
YELLOW_FILL = PatternFill("solid", fgColor="FFFF00") # חריגת VSL
ORANGE_FILL = PatternFill("solid", fgColor="FFA500") # חריגת TIER 1
BLUE_HEADER = PatternFill("solid", fgColor="1F4E79") # כותרת ראשית
LIGHT_BLUE_HEADER = PatternFill("solid", fgColor="2E75B6") # כותרת משנית
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
            # חיפוש שורת הקידוחים ושורת הפרמטרים
            s_idx = next((i for i, r in enumerate(rows) if any("Client Sample ID" in str(v) for v in r if v)), None)
            p_idx = next((i for i, r in enumerate(rows) if r and r[0] == "Parameter"), None)
            
            if s_idx is not None and p_idx is not None:
                samples = {idx: str(v).strip() for idx, v in enumerate(rows[s_idx]) if v and v != "Client Sample ID"}
                for row in rows[p_idx + 1:]:
                    param, unit = row[0], row[2]
                    if not param or not unit: continue
                    for col_idx, sname in samples.items():
                        val = row[col_idx]
                        match = re.match(r"^(S-?\d+[A-Za-z]*)\s*\(([0-9.]+)\)", sname)
                        sid, depth = (match.group(1), float(match.group(2))) if match else (sname, 0.0)
                        res_str = str(val).strip() if val is not None else ""
                        # ניקוי תווים לא מספריים לצורך השוואה
                        clean_val = str(res_str).replace('<', '').replace('>', '').strip()
                        try:
                            res_num = float(clean_val)
                        except:
                            res_num = None
                        
                        all_records.append({
                            "sample_id": sid, "depth": depth, "compound": str(param).strip(),
                            "unit": unit, "result": res_num, "result_str": res_str
                        })
    except Exception as e:
        st.error(f"שגיאה בעיבוד קובץ ALS: {e}")
    return pd.DataFrame(all_records)

# --- ממשק האפליקציה ---
st.title("🧪 הפקת דוח מאוחד - 4 לשוניות")

with st.sidebar:
    st.header("📂 טעינת ערכי סף")
    f_main = st.file_uploader("1. ערכי סף כלליים (גרסה 7)", type=["csv", "xlsx"])
    f_pfas = st.file_uploader("2. ערכי סף PFAS", type=["csv", "xlsx"])
    tier1_col_input = st.text_input("שם עמודת Tier 1 לצביעה", "Tier 1 industrial A 0-6m")
    st.markdown("---")
    als_files = st.file_uploader("3. העלה קבצי ALS", type=["xlsx"], accept_multiple_files=True)

if (f_main or f_pfas) and als_files:
    # בניית מילון ערכי סף משני הקבצים
    thresh_map = {}
    for f in [f_main, f_pfas]:
        if f:
            df_t = pd.read_csv(f) if f.name.endswith('csv') else pd.read_excel(f)
            df_t.columns = [c.strip() for c in df_t.columns]
            chem_c = next((c for c in df_t.columns if c.lower() in ["chemical", "parameter", "name", "תרכובת", "analyte"]), None)
            vsl_c = next((c for c in df_t.columns if "VSL" in c), None)
            t1_c = tier1_col_input if tier1_col_input in df_t.columns else next((c for c in df_t.columns if "Tier 1" in c), None)
            cas_c = next((c for c in df_t.columns if "CAS" in c), None)
            if chem_c:
                for _, r in df_t.iterrows():
                    thresh_map[norm_name(r[chem_c])] = {"vsl": r.get(vsl_c), "tier1": r.get(t1_c), "cas": r.get(cas_c)}

    all_data = pd.concat([parse_als_data(f.read()) for f in als_files], ignore_index=True)
    wb = Workbook()
    wb.remove(wb.active)

    # שיוך חכם ללשוניות לפי מילות מפתח
    def get_sheet_name(compound):
        c = compound.lower()
        if any(x in c for x in ['tph', 'dro', 'oro', 'oil', 'c10', 'c40']): return "TPH"
        if any(x in c for x in ['pfos', 'pfoa', 'pfas', 'pfbs', 'pfhx']): return "PFAS"
        if any(x in c for x in ['arsenic', 'cadmium', 'chromium', 'copper', 'lead', 'mercury', 'nickel', 'zinc', 'aluminum', 'iron', 'manganese', 'silver', 'barium']): return "מתכות"
        return "VOC & SVOC"

    all_data['target_sheet'] = all_data['compound'].apply(get_sheet_name)

    for s_name in ["מתכות", "TPH", "PFAS", "VOC & SVOC"]:
        df_sub = all_data[all_data['target_sheet'] == s_name]
        if df_sub.empty: continue
        
        ws = wb.create_sheet(s_name)
        analytes = sorted(df_sub['compound'].unique())
        samples = sorted(df_sub['sample_id'].unique(), key=natural_sort_key, reverse=True) # מיון מהגבוה לנמוך

        # בניית כותרות 1-5
        ws.append(["שם קידוח", "עומק", "אנליזה"] + analytes)
        ws.append(["", "", "יחידות"] + [df_sub[df_sub['compound']==a]['unit'].iloc[0] if not df_sub[df_sub['compound']==a].empty else "mg/kg" for a in analytes])
        ws.append(["", "CAS", ""] + [thresh_map.get(norm_name(a), {}).get("cas", "-") for a in analytes])
        ws.append(["", "VSL", ""] + [thresh_map.get(norm_name(a), {}).get("vsl", "-") for a in analytes])
        ws.append(["", "TIER 1", ""] + [thresh_map.get(norm_name(a), {}).get("tier1", "-") for a in analytes])

        for r in range(1, 6):
            for c in range(1, len(analytes)+4):
                cell = ws.cell(r, c)
                cell.fill = BLUE_HEADER if r==1 else LIGHT_BLUE_HEADER
                cell.font = WHITE_FONT; cell.border = make_border(); cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        # מילוי נתונים ומיזוג תאים בעמודה A
        row_idx = 6
        for sid in samples:
            df_s = df_sub[df_sub['sample_id'] == sid]
            depths = sorted(df_s['depth'].unique()); start_r = row_idx
            for d in depths:
                ws.cell(row_idx, 1, sid if row_idx == start_r else ""); ws.cell(row_idx, 2, d)
                for i, a in enumerate(analytes, 4):
                    res = df_s[(df_s['depth']==d) & (df_s['compound']==a)]
                    if not res.empty:
                        c = ws.cell(row_idx, i, res['result_str'].values[0]); val = res['result'].values[0]
                        thr = thresh_map.get(norm_name(a), {})
                        if val is not None:
                            try:
                                # השוואה לערכי סף לצורך צביעה
                                if thr.get("vsl") and val > float(thr["vsl"]): c.fill = YELLOW_FILL
                                elif thr.get("tier1") and val > float(thr["tier1"]): c.fill = ORANGE_FILL
                            except: pass
                    ws.cell(row_idx, i).border = make_border(); ws.cell(row_idx, i).alignment = Alignment(horizontal="center")
                row_idx += 1
            if row_idx > start_r + 1: 
                ws.merge_cells(start_row=start_r, start_column=1, end_row=row_idx-1, end_column=1)
                ws.cell(start_r, 1).alignment = Alignment(vertical="center", horizontal="center")

    buf = io.BytesIO(); wb.save(buf)
    st.success("✅ העיבוד הושלם! ניתן להוריד את הקובץ.")
    st.download_button("⬇️ הורד קובץ אקסל סופי", buf.getvalue(), "Soil_Report_Final.xlsx", use_container_width=True)
