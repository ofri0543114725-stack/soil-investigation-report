import streamlit as st
import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import io, re

st.set_page_config(page_title="דוח סקר קרקע", layout="wide", page_icon="🧪")
st.title("🧪 מערכת עיבוד תוצאות מעבדה")
st.markdown("---")

# ─────────────────────────────────────────────
# STYLES  (exact match to your reference files)
# ─────────────────────────────────────────────
YELLOW_FILL   = PatternFill("solid", fgColor="FFFF00")   # VSL exceed
ORANGE_FILL   = PatternFill("solid", fgColor="FFC000")   # TIER 1 exceed
HDR_BLUE_FILL = PatternFill("solid", fgColor="B7D7F0")   # metals header
HDR_CYAN_FILL = PatternFill("solid", fgColor="00B0F0")   # PFAS header
HDR_DARK_FILL = PatternFill("solid", fgColor="1F4E79")   # merged title row
NO_FILL       = PatternFill("none")

def thin_border():
    s = Side(style="thin", color="000000")
    return Border(left=s, right=s, top=s, bottom=s)

def hdr_font(size=11):  return Font(bold=True, name="Arial", size=size)
def data_font(size=10): return Font(name="Arial", size=size)
def white_font(size=11):return Font(bold=True, name="Arial", size=size, color="FFFFFF")
def center():           return Alignment(horizontal="center", vertical="center", wrap_text=True)
def right():            return Alignment(horizontal="right",  vertical="center", wrap_text=True)

# ─────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────
def norm(s):
    s = "" if s is None else str(s).strip().lower()
    return re.sub(r"\s+", " ", s.replace("\xa0"," ").replace("–","-").replace("—","-"))

def to_float(v):
    s = str(v).strip() if v is not None else ""
    if s.startswith("<") or s.startswith(">"):
        try: return float(s[1:])
        except: return None
    try:    return float(s)
    except: return None

def sort_key(sid):
    """Sort S1 < S2 < S10 < S100 numerically."""
    m = re.match(r"S-?(\d+)", str(sid), re.I)
    return int(m.group(1)) if m else 0

# ─────────────────────────────────────────────
# ALS PARSER  →  long-format DataFrame
# ─────────────────────────────────────────────
def parse_als_file(file_bytes, filename):
    try:
        wb = load_workbook(io.BytesIO(file_bytes), data_only=True)
    except Exception as e:
        return None, str(e)

    main = next((wb[n] for n in wb.sheetnames if "Client" in n and "SOIL" in n), wb.worksheets[0])
    rows = list(main.iter_rows(values_only=True))

    # find sample-id row
    sid_row = next((i for i,r in enumerate(rows) if any("Client Sample ID" in str(v) for v in r if v)), None)
    if sid_row is None: return None, "לא נמצאה שורת Sample IDs"

    col2sample = {ci: str(v).strip() for ci,v in enumerate(rows[sid_row]) if v and v != "Client Sample ID"}

    # find Parameter header
    param_row = next((i for i,r in enumerate(rows) if r and r[0] == "Parameter"), None)
    if param_row is None: return None, "לא נמצאה שורת Parameter"

    records, group = [], "Unknown"
    for row in rows[param_row+1:]:
        param, method, unit = (row[i] if len(row)>i else None for i in range(3))
        if not param: continue
        if not method and not unit:
            group = str(param).strip(); continue

        for ci, sname in col2sample.items():
            val = row[ci] if ci < len(row) else None
            m   = re.match(r"^(S-?\d+[A-Za-z0-9]*)\s*\(([0-9.]+)\)", sname)
            sid, depth = (m.group(1), float(m.group(2))) if m else (sname, None)
            rs = str(val).strip() if val is not None else ""
            result = 0.0 if rs.startswith("<") else to_float(val)
            if result is not None:
                records.append({"sample_id":sid,"depth":depth,"compound":str(param).strip(),
                                 "unit":str(unit).strip() if unit else "mg/kg",
                                 "result":result,"result_str":rs,"group":group,"source":filename})

    return (pd.DataFrame(records), None) if records else (None, "לא נמצאו נתונים")

# ─────────────────────────────────────────────
# BUILD OUTPUT SHEETS  (match reference format)
# ─────────────────────────────────────────────

def write_tph_sheet(ws, df, thresh):
    """
    Columns: שם קידוח | עומק | אנליזה | TPH DRO | TPH ORO | Total TPH
    Header rows 1-5 like reference file, data from row 6
    """
    metals_cols = ["TPH DRO","TPH ORO","Total TPH"]
    # header rows
    for c,h in enumerate(["שם קידוח","עומק","אנליזה"]+metals_cols, 1):
        cell = ws.cell(1,c,h)
        cell.font=hdr_font(); cell.alignment=center(); cell.border=thin_border()

    # sub-header rows 2-5
    labels = {2:"יחידות", 3:"CAS", 4:"VSL", 5:"TIER 1"}
    sub_data = {
        "TPH DRO": {"יחידות":"mg/kg","CAS":"C10-C40","VSL":350,"TIER 1":350},
        "TPH ORO": {"יחידות":"mg/kg","CAS":"C10-C40","VSL":350,"TIER 1":350},
        "Total TPH":{"יחידות":"mg/kg","CAS":"C10-C40","VSL":350,"TIER 1":350},
    }
    for row_i, lbl in labels.items():
        cell = ws.cell(row_i,2,lbl)
        cell.font=hdr_font(); cell.alignment=center(); cell.border=thin_border()
        for c,col in enumerate(metals_cols, 4):
            v = sub_data[col][lbl]
            cell = ws.cell(row_i,c,v)
            cell.font=hdr_font(); cell.alignment=center(); cell.border=thin_border()

    # data  –  group by sample, sorted
    df_s = df.copy()
    df_s["_key"] = df_s["sample_id"].apply(sort_key)
    df_s = df_s.sort_values(["_key","depth"])

    # pivot: sample+depth -> DRO/ORO/Total
    pivoted = {}
    for _, r in df_s.iterrows():
        k = (r["sample_id"], r["depth"])
        cmp = r["compound"].upper()
        if k not in pivoted: pivoted[k] = {"DRO":None,"ORO":None,"Total":None,"rs_DRO":"","rs_ORO":"","rs_Total":""}
        if "DRO" in cmp and "ORO" not in cmp:
            pivoted[k]["DRO"]=r["result"]; pivoted[k]["rs_DRO"]=r["result_str"]
        elif "ORO" in cmp:
            pivoted[k]["ORO"]=r["result"]; pivoted[k]["rs_ORO"]=r["result_str"]
        if "TOTAL" in cmp or ("DRO" in cmp and "ORO" in cmp):
            pass  # handled separately if exists
    # if Total not set, compute
    for k,v in pivoted.items():
        if v["Total"] is None and v["DRO"] is not None and v["ORO"] is not None:
            dro = to_float(v["rs_DRO"]) or 0
            oro = to_float(v["rs_ORO"]) or 0
            v["Total"] = dro + oro
            v["rs_Total"] = str(round(dro+oro,1)) if not (v["rs_DRO"].startswith("<") and v["rs_ORO"].startswith("<")) else f'<{round(dro+oro,1)}'

    VSL   = thresh.get("tph",{}).get("vsl",350)
    TIER1 = thresh.get("tph",{}).get("tier1",350)

    row_i = 6
    prev_sid = None
    for (sid, depth), v in sorted(pivoted.items(), key=lambda x:(sort_key(x[0][0]), x[0][1] or 0)):
        sid_val = sid if sid != prev_sid else None
        prev_sid = sid

        vals = [sid_val, depth, None, v["rs_DRO"] or "", v["rs_ORO"] or "", v["rs_Total"] or ""]
        total_f = v["Total"]
        hl = None
        if total_f is not None:
            if TIER1 and total_f > TIER1: hl = "tier1"
            elif VSL  and total_f > VSL:  hl = "vsl"

        for ci, val in enumerate(vals, 1):
            cell = ws.cell(row_i, ci, val)
            cell.font = data_font()
            cell.alignment = center()
            cell.border = thin_border()
            if hl and ci in [4,5,6]:
                cell.fill = ORANGE_FILL if hl=="tier1" else YELLOW_FILL
                cell.font = Font(bold=True, name="Arial", size=10)

        row_i += 1

    # col widths
    for col, w in zip("ABCDEF", [14,10,14,16,16,16]):
        ws.column_dimensions[get_column_letter(ord(col)-64)].width = w
    ws.freeze_panes = "A6"


def write_metals_sheet(ws, df, thresh):
    """
    Wide format: שם קידוח | עומק | (blank) | Al | As | Ba | ...
    Header rows 1-5 with blue fill, data from row 6
    """
    metals = ["Al","As","Ba","Be","B","Cd","Ca","Cr","Co","Cu","Fe","Pb","Li","Mg","Mn","Hg","Ni","K","Na","Se","Ag","V","Zn"]
    fixed_cols = ["שם קידוח","עומק",None]

    row1 = fixed_cols + metals
    for ci, h in enumerate(row1, 1):
        cell = ws.cell(1, ci, h)
        cell.fill = HDR_BLUE_FILL
        cell.font = hdr_font()
        cell.alignment = center()
        cell.border = thin_border()

    # sub-rows 2-5: units, CAS, VSL, TIER1
    cas_map  = {"Al":"7429-90-5","As":"7440-38-2","Ba":"7440-39-3","Be":"7440-41-7","B":"7440-42-8",
                "Cd":"7440-43-9","Ca":"-","Cr":"7440-47-3","Co":"7440-48-4","Cu":"7440-50-8",
                "Fe":"7439-89-6","Pb":"7439-92-1","Li":"7439-93-2","Mg":"7439-95-4","Mn":"7439-96-5",
                "Hg":"7439-97-6","Ni":"7440-02-0","K":"7440-09-7","Na":"7440-23-5","Se":"7782-49-2",
                "Ag":"7440-22-4","V":"7440-62-2","Zn":"7440-66-6"}
    vsl_map  = {"Al":78000,"As":16,"Ba":15600,"Be":156,"B":1230,"Cd":7.14,"Ca":"-","Cr":"-","Co":23.4,
                "Cu":191,"Fe":"-","Pb":40,"Li":"-","Mg":"-","Mn":1960,"Hg":6.4,"Ni":191,"K":"-",
                "Na":"-","Se":1.7,"Ag":"-","V":"-","Zn":191}
    tier1_map= {"Al":122000,"As":16,"Ba":16700,"Be":1280,"B":1230,"Cd":87.3,"Ca":"-","Cr":"-","Co":85.7,
                "Cu":191,"Fe":"-","Pb":400,"Li":"-","Mg":"-","Mn":16700,"Hg":6.4,"Ni":191,"K":"-",
                "Na":"-","Se":1.7,"Ag":"-","V":"-","Zn":191}

    sub_rows = {2:("יחידות","mg/kg"), 3:("CAS",cas_map), 4:("VSL",vsl_map), 5:("TIER 1",tier1_map)}
    for row_i,(lbl,data) in sub_rows.items():
        cell = ws.cell(row_i,2,lbl)
        cell.fill=HDR_BLUE_FILL; cell.font=hdr_font(); cell.alignment=center(); cell.border=thin_border()
        for ci,m in enumerate(metals,4):
            val = data if isinstance(data,str) else data.get(m,"-")
            cell = ws.cell(row_i,ci,val)
            cell.fill=HDR_BLUE_FILL; cell.font=hdr_font(); cell.alignment=center(); cell.border=thin_border()

    # data
    df_s = df.copy()
    df_s["_key"] = df_s["sample_id"].apply(sort_key)
    df_s = df_s.sort_values(["_key","depth"])

    # pivot
    pivoted = {}
    for _, r in df_s.iterrows():
        k=(r["sample_id"],r["depth"])
        if k not in pivoted: pivoted[k]={}
        pivoted[k][r["compound"]] = r["result_str"]

    row_i = 6
    prev_sid = None
    for (sid,depth), mvals in sorted(pivoted.items(), key=lambda x:(sort_key(x[0][0]),x[0][1] or 0)):
        sid_val = sid if sid != prev_sid else None
        prev_sid = sid
        row_data = [sid_val, depth, None] + [mvals.get(m,"") for m in metals]

        for ci, val in enumerate(row_data, 1):
            cell = ws.cell(row_i, ci, val)
            cell.font = data_font()
            cell.alignment = center()
            cell.border = thin_border()

            if ci >= 4:  # metal value column
                m = metals[ci-4]
                fval = to_float(val)
                t1   = tier1_map.get(m)
                vsl  = vsl_map.get(m)
                if fval is not None and fval > 0:
                    try:
                        if t1 and t1 != "-" and fval > float(t1):
                            cell.fill = ORANGE_FILL; cell.font=Font(bold=True,name="Arial",size=10)
                        elif vsl and vsl != "-" and fval > float(vsl):
                            cell.fill = YELLOW_FILL; cell.font=Font(bold=True,name="Arial",size=10)
                    except: pass
        row_i += 1

    for ci,w in enumerate([14,10,6]+[11]*len(metals),1):
        ws.column_dimensions[get_column_letter(ci)].width = w
    ws.freeze_panes = "A6"


def write_pfas_sheet(ws, df, thresh):
    """
    PFAS: תרכובת | CAS | ערך סף | יחידות | LOR | שם קידוח → columns
    Row 1 = headers (cyan), row 2 = depth, rows 3+ = compounds
    """
    samples = sorted(df["sample_id"].unique(), key=sort_key)
    sample_depths = {}
    for _, r in df.iterrows():
        if r["sample_id"] not in sample_depths:
            sample_depths[r["sample_id"]] = r["depth"]

    fixed = ["שם התרכובת","CAS","ערך סף","יחידות","LOR","שם הקידוח"]
    all_cols = fixed + samples

    # row 1 - headers
    for ci, h in enumerate(all_cols, 1):
        cell = ws.cell(1, ci, h)
        cell.fill = HDR_CYAN_FILL
        cell.font = hdr_font()
        cell.alignment = center()
        cell.border = thin_border()

    # row 2 - depths
    cell = ws.cell(2,6,"עומק")
    cell.fill=HDR_CYAN_FILL; cell.font=hdr_font(); cell.alignment=center(); cell.border=thin_border()
    for ci, sid in enumerate(samples, 7):
        cell = ws.cell(2, ci, sample_depths.get(sid,""))
        cell.fill=HDR_CYAN_FILL; cell.font=hdr_font(); cell.alignment=center(); cell.border=thin_border()

    # compound data
    compounds = df["compound"].unique()
    thresh_map = thresh.get("pfas",{})

    for row_i, cmp in enumerate(compounds, 3):
        df_c = df[df["compound"]==cmp]
        cas   = df_c.iloc[0]["compound"]  # no CAS in ALS PFAS usually
        unit  = df_c.iloc[0]["unit"]
        thr   = thresh_map.get(norm(cmp),{})
        vsl   = thr.get("vsl","-")
        lor   = ""

        row_vals = [cmp, "-", vsl, unit, lor, None]
        for ci, val in enumerate(row_vals, 1):
            cell = ws.cell(row_i, ci, val)
            cell.font=data_font(); cell.alignment=center(); cell.border=thin_border()

        for ci, sid in enumerate(samples, 7):
            row_sid = df_c[df_c["sample_id"]==sid]
            rs = row_sid.iloc[0]["result_str"] if not row_sid.empty else ""
            cell = ws.cell(row_i, ci, rs)
            cell.font=data_font(); cell.alignment=center(); cell.border=thin_border()

            if rs and not rs.startswith("<"):
                fval = to_float(rs)
                if fval and vsl and vsl != "-":
                    try:
                        if fval > float(vsl):
                            cell.fill=YELLOW_FILL; cell.font=Font(bold=True,name="Arial",size=10)
                    except: pass

    ws.column_dimensions["A"].width = 45
    for ci in range(2, len(all_cols)+1):
        ws.column_dimensions[get_column_letter(ci)].width = 12
    ws.freeze_panes = "G3"


def write_voc_sheet(ws, df, thresh):
    """
    VOCs: קבוצה | קבוצה | שם התרכובת | CAS | VSL | TIER 1 | יחידות | שם קידוח → cols
    Row 1 = headers, row 2 = depths
    """
    samples = sorted(df["sample_id"].unique(), key=sort_key)
    sample_depths = {}
    for _, r in df.iterrows():
        if r["sample_id"] not in sample_depths:
            sample_depths[r["sample_id"]] = r["depth"]

    fixed = ["קבוצה","קבוצה","שם התרכובת","CAS","VSL","TIER 1","יחידות","שם קידוח"]
    all_cols = fixed + samples

    for ci, h in enumerate(all_cols, 1):
        cell = ws.cell(1, ci, h)
        cell.font=hdr_font(9); cell.alignment=center(); cell.border=thin_border()

    cell = ws.cell(2,8,"עומק")
    cell.font=hdr_font(9); cell.alignment=center(); cell.border=thin_border()
    for ci, sid in enumerate(samples, 9):
        for di, (_, r) in enumerate(df[df["sample_id"]==sid].iterrows()):
            cell = ws.cell(2, ci+di, r["depth"])
            cell.font=hdr_font(9); cell.alignment=center(); cell.border=thin_border()
            break

    # group tracking
    group_col = df["group"].values if "group" in df.columns else [""] * len(df)
    compounds = df.drop_duplicates("compound")[["compound","group","unit"]].copy()

    thresh_map = thresh.get("voc",{})

    for row_i, (_, crow) in enumerate(compounds.iterrows(), 3):
        cmp   = crow["compound"]
        grp   = crow["group"]
        unit  = crow["unit"]
        thr   = thresh_map.get(norm(cmp),{})
        vsl   = thr.get("vsl","")
        tier1 = thr.get("tier1","")

        row_vals = [None, grp, cmp, "", vsl, tier1, unit, None]
        for ci, val in enumerate(row_vals, 1):
            cell = ws.cell(row_i, ci, val)
            cell.font=data_font(9); cell.alignment=center(); cell.border=thin_border()

        df_c = df[df["compound"]==cmp]
        for ci, sid in enumerate(samples, 9):
            row_sid = df_c[df_c["sample_id"]==sid]
            rs = row_sid.iloc[0]["result_str"] if not row_sid.empty else ""
            cell = ws.cell(row_i, ci, rs)
            cell.font=data_font(9); cell.alignment=center(); cell.border=thin_border()

            if rs and not rs.startswith("<"):
                fval = to_float(rs)
                try:
                    if tier1 and fval and float(tier1) > 0 and fval > float(tier1):
                        cell.fill=ORANGE_FILL; cell.font=Font(bold=True,name="Arial",size=9)
                    elif vsl and fval and float(vsl) > 0 and fval > float(vsl):
                        cell.fill=YELLOW_FILL; cell.font=Font(bold=True,name="Arial",size=9)
                except: pass

    ws.column_dimensions["A"].width = 8
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 35
    ws.column_dimensions["D"].width = 12
    ws.column_dimensions["E"].width = 10
    ws.column_dimensions["F"].width = 10
    ws.column_dimensions["G"].width = 10
    ws.freeze_panes = "I3"


# ─────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────
st.sidebar.header("⚙️ הגדרות")
st.sidebar.markdown("🟡 חריגה מעל VSL")
st.sidebar.markdown("🟠 חריגה מעל TIER 1")

# ─────────────────────────────────────────────
# FILE UPLOADS
# ─────────────────────────────────────────────
col1, col2 = st.columns(2)
with col1:
    st.subheader("📁 קובץ ערכי סף")
    threshold_file = st.file_uploader("העלה קובץ ערכי סף (Excel)", type=["xlsx","xls"], key="thr")
with col2:
    st.subheader("📂 קבצי נתונים מ-ALS")
    data_files = st.file_uploader("העלה קבצי ALS — אפשר כמה ביחד", type=["xlsx","xls"],
                                   accept_multiple_files=True, key="data")

if not threshold_file:
    st.info("👆 העלה קובץ ערכי סף וקבצי ALS כדי להתחיל")
    st.stop()

if not data_files:
    st.warning("⚠️ העלה קבצי נתונים מ-ALS")
    st.stop()

# ─────────────────────────────────────────────
# READ THRESHOLDS
# ─────────────────────────────────────────────
df_thr = pd.read_excel(threshold_file)
df_thr.columns = [str(c).lower().strip() for c in df_thr.columns]
comp_col = next((c for c in df_thr.columns if c in ["compound","chemical","parameter","analyte","name","תרכובת"]), None)
vsl_col  = next((c for c in df_thr.columns if "vsl" in c), None)
t1_col   = next((c for c in df_thr.columns if "tier" in c and "1" in c), None)

thresh = {"tph":{}, "metals":{}, "pfas":{}, "voc":{}}
if comp_col and vsl_col:
    for _, r in df_thr.iterrows():
        k = norm(r[comp_col])
        if k and k != "nan":
            thresh["pfas"][k] = {"vsl": r.get(vsl_col), "tier1": r.get(t1_col) if t1_col else None}
            thresh["voc"][k]  = {"vsl": r.get(vsl_col), "tier1": r.get(t1_col) if t1_col else None}
    st.success(f"✅ נטענו {len(thresh['pfas'])} ערכי סף")

# ─────────────────────────────────────────────
# PARSE ALS FILES
# ─────────────────────────────────────────────
all_data = []
for f in data_files:
    df, err = parse_als_file(f.read(), f.name)
    if err: st.warning(f"⚠️ {f.name}: {err}")
    else:
        all_data.append(df)
        st.success(f"✅ {f.name} — {len(df)} תוצאות")

if not all_data:
    st.error("לא נטענו נתונים.")
    st.stop()

df_all = pd.concat(all_data, ignore_index=True)

with st.expander(f"👁️ תצוגה מקדימה ({len(df_all)} שורות)"):
    st.dataframe(df_all.head(30))

# ─────────────────────────────────────────────
# BUILD OUTPUT WORKBOOK
# ─────────────────────────────────────────────
wb_out = Workbook()
wb_out.remove(wb_out.active)

groups = df_all["group"].unique()

# classify groups
tph_df    = df_all[df_all["group"].str.upper().str.contains("TPH|HYDROCARBON", na=False)]
metals_df = df_all[df_all["group"].str.upper().str.contains("METAL|CATION|EXTRACTABLE", na=False)]
pfas_df   = df_all[df_all["group"].str.upper().str.contains("PFAS|PERFLUORO|FLUORINATED", na=False)]
voc_df    = df_all[df_all["group"].str.upper().str.contains("VOC|SVOC|HALOGENATED|BTEX|AROMATIC|ANILINE|ALCOHOL|ALDEHYDE|PAH|PHENOL|PCB|NAPHTHALENE|KETONE|CHLORO", na=False)]

# if empty, put everything in VOC
if tph_df.empty and metals_df.empty and pfas_df.empty:
    voc_df = df_all

sheets_to_build = []
if not tph_df.empty:    sheets_to_build.append(("TPH",    tph_df))
if not metals_df.empty: sheets_to_build.append(("Metals", metals_df))
if not voc_df.empty:    sheets_to_build.append(("VOC+SVOC", voc_df))
if not pfas_df.empty:   sheets_to_build.append(("PFAS",   pfas_df))

if not sheets_to_build:
    sheets_to_build = [("Results", df_all)]

for sheet_name, df_sheet in sheets_to_build:
    ws = wb_out.create_sheet(title=sheet_name)
    sn = sheet_name.upper()
    if "TPH" in sn:
        write_tph_sheet(ws, df_sheet, thresh)
    elif "METAL" in sn:
        write_metals_sheet(ws, df_sheet, thresh)
    elif "PFAS" in sn:
        write_pfas_sheet(ws, df_sheet, thresh)
    else:
        write_voc_sheet(ws, df_sheet, thresh)

# ─────────────────────────────────────────────
# DOWNLOAD
# ─────────────────────────────────────────────
st.markdown("---")
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
