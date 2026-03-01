import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import io

st.set_page_config(page_title="דוח סקר קרקע", layout="wide")

st.title("🧪 מערכת עיבוד תוצאות מעבדה")
st.markdown("---")

# ==================== צבעים ====================
YELLOW_FILL = PatternFill("solid", fgColor="FFFF00")   # חריגת VSL
ORANGE_FILL = PatternFill("solid", fgColor="FFA500")   # חריגת TIER 1
HEADER_FILL = PatternFill("solid", fgColor="1F4E79")   # כותרת ראשית
SUBHEADER_FILL = PatternFill("solid", fgColor="2E75B6") # כותרת משנית
WHITE_FONT = Font(color="FFFFFF", bold=True, name="Arial", size=10)
BOLD_FONT = Font(bold=True, name="Arial", size=10)
NORMAL_FONT = Font(name="Arial", size=10)

def make_border():
    side = Side(style="thin", color="000000")
    return Border(left=side, right=side, top=side, bottom=side)

def style_header(cell, level=1):
    cell.font = WHITE_FONT
    cell.fill = HEADER_FILL if level == 1 else SUBHEADER_FILL
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border = make_border()

def style_data(cell, highlight=None):
    cell.font = NORMAL_FONT
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.border = make_border()
    if highlight == "vsl":
        cell.fill = YELLOW_FILL
        cell.font = Font(name="Arial", size=10, bold=True)
    elif highlight == "tier1":
        cell.fill = ORANGE_FILL
        cell.font = Font(name="Arial", size=10, bold=True)

# ==================== Sidebar - הגדרות ====================
st.sidebar.header("⚙️ הגדרות")

tier1_options = {
    "Industrial": "תעשייתי",
    "Residential": "מגורים",
    "Recreational": "נופש/פארק"
}
selected_tier1 = st.sidebar.selectbox(
    "בחר TIER 1:",
    options=list(tier1_options.keys()),
    format_func=lambda x: f"{x} ({tier1_options[x]})"
)

st.sidebar.markdown("---")
st.sidebar.markdown("**צבעים:**")
st.sidebar.markdown("🟡 חריגה מעל VSL")
st.sidebar.markdown("🟠 חריגה מעל TIER 1")

# ==================== העלאת קבצים ====================
col1, col2 = st.columns(2)

with col1:
    st.subheader("📁 קובץ ערכי סף")
    threshold_file = st.file_uploader(
        "העלה קובץ ערכי סף (Excel)",
        type=["xlsx", "xls"],
        key="thresholds"
    )

with col2:
    st.subheader("📂 קבצי נתונים מהמעבדה")
    data_files = st.file_uploader(
        "העלה קבצי תוצאות (Excel) — אפשר כמה ביחד",
        type=["xlsx", "xls"],
        accept_multiple_files=True,
        key="data"
    )

# ==================== מבנה התבנית ====================
st.markdown("---")
with st.expander("📋 מבנה התבנית הנדרשת לקבצי הנתונים", expanded=False):
    st.markdown("""
    **קבצי הנתונים** צריכים להכיל את העמודות הבאות (בשורה הראשונה):
    
    | עמודה | תיאור | דוגמה |
    |-------|-------|-------|
    | `sample_id` | שם הקידוח/דגימה | S-1 |
    | `depth` | עומק (מ') | 1.5 |
    | `compound` | שם התרכובת | Naphthalene |
    | `cas` | מספר CAS | 91-20-3 |
    | `result` | תוצאה (mg/kg) | 0.45 |
    | `unit` | יחידות | mg/kg |
    | `group` | קבוצת מזהמים | TPH / Metals / PFAS / VOCs / SVOCs |
    
    **קובץ ערכי הסף** צריך להכיל:
    
    | עמודה | תיאור |
    |-------|-------|
    | `cas` | מספר CAS |
    | `compound` | שם התרכובת |
    | `vsl` | ערך VSL |
    | `tier1_industrial` | TIER 1 תעשייתי |
    | `tier1_residential` | TIER 1 מגורים |
    | `tier1_recreational` | TIER 1 נופש |
    """)

# ==================== עיבוד ====================
if threshold_file and data_files:
    st.markdown("---")
    
    try:
        # קרא ערכי סף
        df_thresh = pd.read_excel(threshold_file)
        df_thresh.columns = [c.lower().strip() for c in df_thresh.columns]
        df_thresh['cas'] = df_thresh['cas'].astype(str).str.strip()
        
        # בנה dict ערכי סף
        thresh_dict = {}
        tier1_col = f"tier1_{selected_tier1.lower()}"
        
        for _, row in df_thresh.iterrows():
            cas = str(row['cas']).strip()
            thresh_dict[cas] = {
                'compound': row.get('compound', ''),
                'vsl': row.get('vsl', None),
                'tier1': row.get(tier1_col, None)
            }
        
        # קרא כל קבצי הנתונים
        all_data = []
        for f in data_files:
            df = pd.read_excel(f)
            df.columns = [c.lower().strip() for c in df.columns]
            df['source_file'] = f.name
            all_data.append(df)
        
        df_all = pd.concat(all_data, ignore_index=True)
        df_all['cas'] = df_all['cas'].astype(str).str.strip()
        
        st.success(f"✅ נטענו {len(df_all)} שורות נתונים מ-{len(data_files)} קבצים")
        
        # הצג תצוגה מקדימה
        with st.expander("👁️ תצוגה מקדימה של הנתונים"):
            st.dataframe(df_all.head(20))
        
        # ==================== בנה קובץ פלט ====================
        groups = df_all['group'].unique() if 'group' in df_all.columns else ['All']
        
        wb = Workbook()
        wb.remove(wb.active)  # הסר גיליון ריק
        
        stats = {"total": 0, "vsl_exceed": 0, "tier1_exceed": 0}
        
        for group in sorted(groups):
            ws = wb.create_sheet(title=str(group)[:31])
            
            if 'group' in df_all.columns:
                df_group = df_all[df_all['group'] == group].copy()
            else:
                df_group = df_all.copy()
            
            # כותרות עמודות
            headers = ["שם קידוח", "עומק (מ')", "תרכובת", "CAS", "תוצאה", "יחידות", "VSL", f"TIER 1\n({selected_tier1})", "סטטוס"]
            col_map = ["sample_id", "depth", "compound", "cas", "result", "unit", None, None, None]
            
            # שורת כותרת ראשית
            ws.merge_cells(f"A1:{get_column_letter(len(headers))}1")
            title_cell = ws["A1"]
            title_cell.value = f"תוצאות אנליזות — {group}"
            style_header(title_cell, level=1)
            ws.row_dimensions[1].height = 25
            
            # שורת כותרות עמודות
            for col_idx, header in enumerate(headers, 1):
                cell = ws.cell(row=2, column=col_idx, value=header)
                style_header(cell, level=2)
            ws.row_dimensions[2].height = 35
            
            # נתונים
            for row_idx, (_, row) in enumerate(df_group.iterrows(), start=3):
                cas = str(row.get('cas', '')).strip()
                result_val = row.get('result', None)
                
                # ערכי סף
                thresh = thresh_dict.get(cas, {})
                vsl = thresh.get('vsl', None)
                tier1 = thresh.get('tier1', None)
                
                # קבע חריגה
                highlight = None
                status = "תקין"
                try:
                    r = float(result_val)
                    if tier1 and float(tier1) > 0 and r > float(tier1):
                        highlight = "tier1"
                        status = f"⚠️ חריגת TIER 1"
                        stats["tier1_exceed"] += 1
                    elif vsl and float(vsl) > 0 and r > float(vsl):
                        highlight = "vsl"
                        status = f"⚡ חריגת VSL"
                        stats["vsl_exceed"] += 1
                except:
                    pass
                
                stats["total"] += 1
                
                # כתוב שורה
                values = [
                    row.get('sample_id', ''),
                    row.get('depth', ''),
                    row.get('compound', ''),
                    cas,
                    result_val,
                    row.get('unit', 'mg/kg'),
                    vsl if vsl else "—",
                    tier1 if tier1 else "—",
                    status
                ]
                
                for col_idx, val in enumerate(values, 1):
                    cell = ws.cell(row=row_idx, column=col_idx, value=val)
                    # צבע רק עמודות תוצאה וסטטוס
                    if col_idx in [5, 9]:
                        style_data(cell, highlight)
                    else:
                        style_data(cell)
            
            # רוחב עמודות
            col_widths = [15, 10, 25, 15, 12, 10, 12, 15, 18]
            for i, width in enumerate(col_widths, 1):
                ws.column_dimensions[get_column_letter(i)].width = width
            
            # הקפא שורות כותרת
            ws.freeze_panes = "A3"
        
        # ==================== הורדה ====================
        st.markdown("---")
        
        # סטטיסטיקות
        col1, col2, col3 = st.columns(3)
        col1.metric("סה״כ תוצאות", stats["total"])
        col2.metric("🟡 חריגות VSL", stats["vsl_exceed"])
        col3.metric("🟠 חריגות TIER 1", stats["tier1_exceed"])
        
        # שמור לזיכרון
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        
        st.download_button(
            label="⬇️ הורד קובץ Excel מעובד",
            data=output,
            file_name="soil_investigation_results.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
        
    except Exception as e:
        st.error(f"שגיאה בעיבוד: {str(e)}")
        st.exception(e)

elif not threshold_file and not data_files:
    st.info("👆 העלה קובץ ערכי סף וקבצי נתונים כדי להתחיל")
elif not threshold_file:
    st.warning("⚠️ חסר קובץ ערכי סף")
elif not data_files:
    st.warning("⚠️ חסרים קבצי נתונים מהמעבדה")
