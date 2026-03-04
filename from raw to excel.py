import streamlit as st
import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import io, re

st.set_page_config(page_title="דוח סקר קרקע", layout="wide", page_icon="🧪")
st.title("🧪 מערכת עיבוד תוצאות מעבדה")
st.markdown("---")

# ── STYLES ───────────────────────────────────────────────────────────────────
YELLOW_FILL   = PatternFill("solid", fgColor="FFFF00")
ORANGE_FILL   = PatternFill("solid", fgColor="FFC000")
HDR_BLUE_FILL = PatternFill("solid", fgColor="B7D7F0")
HDR_CYAN_FILL = PatternFill("solid", fgColor="00B0F0")

def thin_border():
    s = Side(style="thin", color="000000")
    return Border(left=s, right=s, top=s, bottom=s)

def style_hdr(cell, fill=None, sz=11):
    cell.font      = Font(bold=True, name="Arial", size=sz)
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border    = thin_border()
    if fill: cell.fill = fill

def style_data(cell, hl=None, sz=10):
    cell.font      = Font(bold=bool(hl), name="Arial", size=sz)
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border    = thin_border()
    if   hl == "tier1": cell.fill = ORANGE_FILL
    elif hl == "vsl":   cell.fill = YELLOW_FILL

# ── HELPERS ──────────────────────────────────────────────────────────────────
def norm(s):
    s = "" if s is None else str(s).strip().lower()
    return re.sub(r"\s+", " ", s.replace("\xa0", " "))

def to_float(v):
    s = str(v).strip() if v is not None else ""
    try: return float(s.lstrip("<>").strip())
    except: return None

def sort_key(sid):
    sid = re.sub(r"\s*\(.*?\)", "", str(sid)).strip()
    m = re.match(r"S-?(\d+)", sid, re.I)
    return int(m.group(1)) if m else 9999

def clean_sid(sid):
    """Remove depth parenthetical: 'S85 (0.5)' -> 'S85'"""
    return re.sub(r"\s*\([^)]*\)", "", str(sid)).strip()

def check_exceed(val_str, vsl, tier1):
    if not val_str or str(val_str).strip().startswith("<"):
        return None
    f = to_float(val_str)
    if f is None: return None
    try:
        if tier1 is not None and pd.notna(tier1) and float(tier1) > 0 and f > float(tier1):
            return "tier1"
        if vsl is not None and pd.notna(vsl) and float(vsl) > 0 and f > float(vsl):
            return "vsl"
    except: pass
    return None

# ── PFAS NAME MATCHING ────────────────────────────────────────────────────────
def strip_abbrev(s):
    """Remove trailing (ABBREV) from chemical name, lowercase."""
    cleaned = re.sub(r"\s*\([A-Z0-9:_\-]+\)\s*$", "", str(s)).strip()
    return cleaned.lower()

def match_threshold(compound_name, thresh_dict):
    """Match compound to threshold entry by name (handles PFAS abbreviations)."""
    key = norm(compound_name)
    if key in thresh_dict:
        return thresh_dict[key]
    # strip abbreviations from both sides and compare
    full = strip_abbrev(compound_name)
    for k, v in thresh_dict.items():
        if strip_abbrev(k) == full:
            return v
        if len(full) > 12 and (full in strip_abbrev(k) or strip_abbrev(k) in full):
            return v
    return {}

# ── METAL NAME MAPPING ────────────────────────────────────────────────────────
METAL_MAP = {
    "aluminium":"Al","aluminum":"Al","antimony":"Sb","arsenic":"As","barium":"Ba",
    "beryllium":"Be","bismuth":"Bi","boron":"B","cadmium":"Cd","calcium":"Ca",
    "chromium":"Cr","cobalt":"Co","copper":"Cu","iron":"Fe","lead":"Pb",
    "lithium":"Li","magnesium":"Mg","manganese":"Mn","mercury":"Hg","nickel":"Ni",
    "potassium":"K","selenium":"Se","silver":"Ag","sodium":"Na","vanadium":"V",
    "zinc":"Zn","molybdenum":"Mo","tin":"Sn","titanium":"Ti","strontium":"Sr",
    "thallium":"Tl","phosphorus":"P","sulphur":"S","silicon":"Si",
}
METALS_ORDER = ["Al","Sb","As","Ba","Be","Bi","B","Cd","Ca","Cr","Co","Cu","Fe",
                "Pb","Li","Mg","Mn","Hg","Ni","K","Se","Ag","Na","V","Zn"]

# ── THRESHOLD FILE PARSER ─────────────────────────────────────────────────────
# Your file structure:
# Row 1:  Chemical | CAS | Group | Units | VSL | Tier 1 (merged) ...
# Row 2:  (land use)                     | Industrial      | Residential
# Row 3:  (aquifer)                      | A-1,A,B | B-1,C | A-1,A,B | B-1,C
# Row 4:  (depth)                        | 0-6m | >6m | n/a | 0-6m | >6m | n/a
# Row 6+: data
THRESH_COLS = {
    "VSL":      4,   # 0-indexed col
    "Ind_A_06": 8,
    "Ind_A_6p": 9,
    "Ind_B":    10,
    "Res_A_06": 11,
    "Res_A_6p": 12,
    "Res_B":    13,
}

def load_threshold_file(file_bytes):
    wb = load_workbook(io.BytesIO(file_bytes), data_only=True)
    ws = wb.active
    thresh = {}
    for row in ws.iter_rows(min_row=6, values_only=True):
        if not row[0]: continue
        name = str(row[0]).strip()
        cas  = str(row[1]).strip() if row[1] else "-"
        key  = norm(name)
        def g(ci):
            return row[ci] if ci < len(row) and row[ci] is not None else None
        thresh[key] = {
            "name":     name,
            "cas":      cas,
            "units":    str(row[3]) if row[3] else "mg/kg",
            "VSL":      g(THRESH_COLS["VSL"]),
            "Ind_A_06": g(THRESH_COLS["Ind_A_06"]),
            "Ind_A_6p": g(THRESH_COLS["Ind_A_6p"]),
            "Ind_B":    g(THRESH_COLS["Ind_B"]),
            "Res_A_06": g(THRESH_COLS["Res_A_06"]),
            "Res_A_6p": g(THRESH_COLS["Res_A_6p"]),
            "Res_B":    g(THRESH_COLS["Res_B"]),
        }
    return thresh

def get_tier1_col(land_use, aquifer, depth):
    ind = "industrial" in land_use.lower()
    b   = "b-1" in aquifer.lower()
    if b:
        return "Ind_B" if ind else "Res_B"
    deep = ">6" in depth
    if ind:
        return "Ind_A_06" if not deep else "Ind_A_6p"
    else:
        return "Res_A_06" if not deep else "Res_A_6p"

def get_thresh(compound, thresh_dict, tier1_col):
    """Returns (vsl, tier1, cas)."""
    t = match_threshold(compound, thresh_dict)
    return t.get("VSL"), t.get(tier1_col), t.get("cas", "-")

def build_metals_thresh(thresh_dict, tier1_col):
    result = {}
    for key, v in thresh_dict.items():
        sym = METAL_MAP.get(key)
        if sym:
            result[sym] = {"vsl": v.get("VSL"), "tier1": v.get(tier1_col), "cas": v.get("cas","-")}
    return result

# ── ALS PARSER ────────────────────────────────────────────────────────────────
def parse_als_file(file_bytes, filename):
    try:
        wb = load_workbook(io.BytesIO(file_bytes), data_only=True)
    except Exception as e:
        return None, str(e)

    main = next((wb[n] for n in wb.sheetnames if "Client" in n and "SOIL" in n), wb.worksheets[0])
    rows = list(main.iter_rows(values_only=True))

    sid_idx = next((i for i,r in enumerate(rows)
                    if any("Client Sample ID" in str(v) for v in r if v)), None)
    if sid_idx is None: return None, "לא נמצאה שורת Sample IDs"

    col2sample = {ci: str(v).strip() for ci,v in enumerate(rows[sid_idx])
                  if v and v != "Client Sample ID"}

    ph_idx = next((i for i,r in enumerate(rows) if r and r[0] == "Parameter"), None)
    if ph_idx is None: return None, "לא נמצאה שורת Parameter"

    records = []
    group = "Unknown"
    for row in rows[ph_idx+1:]:
        param  = row[0] if len(row) > 0 else None
        method = row[1] if len(row) > 1 else None
        unit   = row[2] if len(row) > 2 else None
        lor    = row[3] if len(row) > 3 else None
        if not param: continue
        if not method and not unit:
            group = str(param).strip(); continue

        for ci, sname in col2sample.items():
            # skip DUP
            if "DUP" in sname.upper(): continue

            val = row[ci] if ci < len(row) else None
            # parse "S85 (0.5)" -> sid=S85, depth=0.5
            m = re.match(r"^(S-?\d+[A-Za-z0-9]*)\s*\(([0-9.]+)\)", sname)
            if m:
                sid   = m.group(1)          # keep as-is, no parentheses
                depth = float(m.group(2))
            else:
                sid   = clean_sid(sname)
                depth = None

            rs = str(val).strip() if val is not None else ""
            result = None
            if rs.startswith("<"):
                result = 0.0
            elif rs and rs not in ("None", ""):
                try: result = float(rs)
                except: result = None

            if result is not None:
                records.append({
                    "sample_id":      sid,
                    "depth":          depth,
                    "compound":       str(param).strip(),
                    "compound_lower": norm(param),
                    "unit":           str(unit).strip() if unit else "mg/kg",
                    "lor":            lor,
                    "result":         result,
                    "result_str":     rs,
                    "group":          group,
                    "source":         filename,
                })

    if not records: return None, "לא נמצאו נתונים"
    return pd.DataFrame(records), None

# ── SHEET WRITERS ─────────────────────────────────────────────────────────────

def write_tph_sheet(ws, df, thresh_dict, tier1_col):
    def is_dro(c): return "dro" in c or ("c10" in c and "c28" in c) or ("c10" in c and "c40" in c and "oro" not in c)
    def is_oro(c): return "oro" in c or "c24" in c or ("c28" in c and "c40" in c)

    vsl_d, t1_d, _ = get_thresh("C10 - C28 Fraction (DRO)", thresh_dict, tier1_col)
    vsl_o, t1_o, _ = get_thresh("C24 - C40 Fraction (ORO)", thresh_dict, tier1_col)
    vals_vsl  = [v for v in [vsl_d, vsl_o] if v is not None]
    vals_t1   = [v for v in [t1_d,  t1_o]  if v is not None]
    vsl_tot   = min(vals_vsl) if vals_vsl else 350
    t1_tot    = min(vals_t1)  if vals_t1  else 350

    for ci, h in enumerate(["שם קידוח","עומק","","TPH DRO","TPH ORO","Total TPH"], 1):
        style_hdr(ws.cell(1, ci, h))
    sub = {"יחידות":"mg/kg","CAS":"C10-C40","VSL":vsl_tot,"TIER 1":t1_tot}
    for ri, lbl in enumerate(["יחידות","CAS","VSL","TIER 1"], 2):
        style_hdr(ws.cell(ri, 2, lbl))
        for ci in [4,5,6]:
            style_hdr(ws.cell(ri, ci, sub[lbl]))

    # pivot — one row per unique (sample_id, depth)
    pivoted = {}
    for _, r in df.iterrows():
        k = (r["sample_id"], r["depth"])
        if k not in pivoted: pivoted[k] = {"DRO":"","ORO":"","DRO_f":None,"ORO_f":None}
        c = r["compound_lower"]
        if is_dro(c):
            pivoted[k]["DRO"]   = r["result_str"]
            pivoted[k]["DRO_f"] = r["result"]
        elif is_oro(c):
            pivoted[k]["ORO"]   = r["result_str"]
            pivoted[k]["ORO_f"] = r["result"]

    ri = 6
    prev_sid = None
    for (sid, depth), v in sorted(pivoted.items(), key=lambda x:(sort_key(x[0][0]), x[0][1] or 0)):
        dro_f   = v["DRO_f"] or 0
        oro_f   = v["ORO_f"] or 0
        total_f = dro_f + oro_f
        b_lor   = str(v["DRO"]).startswith("<") and str(v["ORO"]).startswith("<")
        total_s = f"<{total_f:.0f}" if b_lor else f"{total_f:.0f}"
        hl      = check_exceed(total_s, vsl_tot, t1_tot)

        sid_val = sid if sid != prev_sid else None
        prev_sid = sid
        for ci, val in enumerate([sid_val, depth, None, v["DRO"], v["ORO"], total_s], 1):
            style_data(ws.cell(ri, ci, val), hl if ci in [4,5,6] else None)
        ri += 1

    for col,w in zip("ABCDEF",[14,10,14,16,16,16]):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A6"


def write_metals_sheet(ws, df, thresh_dict, tier1_col):
    df = df.copy()
    df["sym"] = df["compound_lower"].map(METAL_MAP)
    df = df[df["sym"].notna()]
    if df.empty:
        ws.cell(1,1,"אין נתוני מתכות"); return

    present = set(df["sym"].unique())
    metals  = [m for m in METALS_ORDER if m in present]
    metals += sorted(present - set(METALS_ORDER))
    mt = build_metals_thresh(thresh_dict, tier1_col)

    for ci, h in enumerate(["שם קידוח","עומק",None] + metals, 1):
        style_hdr(ws.cell(1, ci, h), HDR_BLUE_FILL)
    for ri, lbl in enumerate(["יחידות","CAS","VSL","TIER 1"], 2):
        style_hdr(ws.cell(ri, 2, lbl), HDR_BLUE_FILL)
        for ci, sym in enumerate(metals, 4):
            t = mt.get(sym, {})
            val = {"יחידות":"mg/kg","CAS":t.get("cas","-"),
                   "VSL":t.get("vsl","-"),"TIER 1":t.get("tier1","-")}[lbl]
            style_hdr(ws.cell(ri, ci, val), HDR_BLUE_FILL)

    pt = df.pivot_table(index=["sample_id","depth"], columns="sym",
                        values="result_str", aggfunc="first")
    pt = pt.reindex(sorted(pt.index, key=lambda x:(sort_key(x[0]), x[1] or 0)))

    ri = 6
    prev_sid = None
    for (sid, depth), row_data in pt.iterrows():
        sid_val = sid if sid != prev_sid else None
        prev_sid = sid
        for ci, val in enumerate([sid_val, depth, None], 1):
            style_data(ws.cell(ri, ci, val))
        for ci, sym in enumerate(metals, 4):
            val = row_data.get(sym, "") or ""
            val = "" if str(val) == "nan" else str(val)
            t   = mt.get(sym, {})
            hl  = check_exceed(val, t.get("vsl"), t.get("tier1"))
            style_data(ws.cell(ri, ci, val), hl)
        ri += 1

    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 10
    ws.column_dimensions["C"].width = 4
    for ci in range(4, len(metals)+4):
        ws.column_dimensions[get_column_letter(ci)].width = 11
    ws.freeze_panes = "D6"


def write_pfas_sheet(ws, df, thresh_dict, tier1_col):
    df = df.copy()
    samples = sorted(df["sample_id"].unique(), key=sort_key)
    sample_depth = {}
    for _, r in df.iterrows():
        if r["sample_id"] not in sample_depth:
            sample_depth[r["sample_id"]] = r["depth"]

    fixed = ["שם התרכובת","CAS","VSL","TIER 1","יחידות","LOR","שם הקידוח"]
    all_cols = fixed + samples

    for ci, h in enumerate(all_cols, 1):
        style_hdr(ws.cell(1, ci, h), HDR_CYAN_FILL)
    style_hdr(ws.cell(2, 7, "עומק"), HDR_CYAN_FILL)
    for ci, sid in enumerate(samples, 8):
        style_hdr(ws.cell(2, ci, sample_depth.get(sid, "")), HDR_CYAN_FILL)

    for row_i, cmp in enumerate(df["compound"].unique(), 3):
        df_c = df[df["compound"] == cmp]
        vsl, tier1, cas = get_thresh(cmp, thresh_dict, tier1_col)
        unit = df_c.iloc[0]["unit"] if not df_c.empty else "µg/kg"
        lor  = df_c.iloc[0]["lor"]  if not df_c.empty else ""
        for ci, val in enumerate([cmp, cas, vsl, tier1, unit, lor, None], 1):
            style_data(ws.cell(row_i, ci, val))
        for ci, sid in enumerate(samples, 8):
            row_sid = df_c[df_c["sample_id"] == sid]
            rs = row_sid.iloc[0]["result_str"] if not row_sid.empty else ""
            hl = check_exceed(rs, vsl, tier1)
            style_data(ws.cell(row_i, ci, rs), hl)

    ws.column_dimensions["A"].width = 50
    for ci in range(2, len(all_cols)+1):
        ws.column_dimensions[get_column_letter(ci)].width = 13
    ws.freeze_panes = "H3"


def write_voc_sheet(ws, df, thresh_dict, tier1_col):
    df = df.copy()
    samples = sorted(df["sample_id"].unique(), key=sort_key)
    sample_depth = {}
    for _, r in df.iterrows():
        if r["sample_id"] not in sample_depth:
            sample_depth[r["sample_id"]] = r["depth"]

    fixed = ["קבוצה","קבוצה","שם התרכובת","CAS","VSL","TIER 1","יחידות","שם קידוח"]
    all_cols = fixed + samples

    for ci, h in enumerate(all_cols, 1):
        style_hdr(ws.cell(1, ci, h), sz=9)
    style_hdr(ws.cell(2, 8, "עומק"), sz=9)
    for ci, sid in enumerate(samples, 9):
        style_hdr(ws.cell(2, ci, sample_depth.get(sid, "")), sz=9)

    seen = {}
    for _, r in df.iterrows():
        if r["compound"] not in seen:
            seen[r["compound"]] = r["group"]

    for row_i, (cmp, grp) in enumerate(seen.items(), 3):
        df_c = df[df["compound"] == cmp]
        vsl, tier1, cas = get_thresh(cmp, thresh_dict, tier1_col)
        unit = df_c.iloc[0]["unit"] if not df_c.empty else "mg/kg"
        for ci, val in enumerate([None, grp, cmp, cas, vsl, tier1, unit, None], 1):
            style_data(ws.cell(row_i, ci, val), sz=9)
        for ci, sid in enumerate(samples, 9):
            row_sid = df_c[df_c["sample_id"] == sid]
            rs = row_sid.iloc[0]["result_str"] if not row_sid.empty else ""
            hl = check_exceed(rs, vsl, tier1)
            style_data(ws.cell(row_i, ci, rs), hl, sz=9)

    ws.column_dimensions["A"].width = 8
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 35
    ws.column_dimensions["D"].width = 12
    for ci in range(5, len(all_cols)+1):
        ws.column_dimensions[get_column_letter(ci)].width = 11
    ws.freeze_panes = "I3"


# ── SIDEBAR ───────────────────────────────────────────────────────────────────
st.sidebar.header("⚙️ הגדרות ערכי סף")
st.sidebar.markdown("🟡 חריגה מעל VSL &nbsp;&nbsp; 🟠 חריגה מעל TIER 1")
st.sidebar.markdown("---")

land_use = st.sidebar.selectbox("Land Use", ["Industrial", "Residential"], index=0)
aquifer  = st.sidebar.selectbox("Aquifer Sensitivity", ["A-1, A, B", "B-1 or C"], index=0)

if "b-1" in aquifer.lower():
    depth_opts = ["Not Applicable"]
else:
    depth_opts = ["0 - 6 m", ">6 m"]

depth    = st.sidebar.selectbox(
    "Depth to Groundwater from the Contaminated Zone", depth_opts, index=0)

tier1_col = get_tier1_col(land_use, aquifer, depth)
st.sidebar.info(f"TIER 1: {land_use} | {aquifer} | {depth}")

# ── FILE UPLOADS ──────────────────────────────────────────────────────────────
col1, col2 = st.columns(2)
with col1:
    st.subheader("📁 קובץ ערכי סף")
    threshold_file = st.file_uploader("העלה את קובץ ערכי הסף המאוחד", type=["xlsx","xls"], key="thr")
with col2:
    st.subheader("📂 קבצי נתונים מ-ALS")
    data_files = st.file_uploader("העלה קבצי ALS", type=["xlsx","xls"],
                                   accept_multiple_files=True, key="data")

if not threshold_file:
    st.info("👆 העלה קובץ ערכי סף וקבצי ALS כדי להתחיל"); st.stop()
if not data_files:
    st.warning("⚠️ העלה קבצי נתונים מ-ALS"); st.stop()

# ── LOAD THRESHOLDS ───────────────────────────────────────────────────────────
thresh_dict = load_threshold_file(threshold_file.read())
st.success(f"✅ {len(thresh_dict)} ערכי סף | {land_use} | {aquifer} | {depth}")

# ── PARSE ALS ─────────────────────────────────────────────────────────────────
all_data = []
for f in data_files:
    df, err = parse_als_file(f.read(), f.name)
    if err: st.warning(f"⚠️ {f.name}: {err}")
    else:
        all_data.append(df)
        st.success(f"✅ {f.name} — {len(df)} תוצאות")

if not all_data:
    st.error("לא נטענו נתונים."); st.stop()

df_all = pd.concat(all_data, ignore_index=True)

with st.expander(f"👁️ תצוגה מקדימה ({len(df_all)} שורות)"):
    st.dataframe(df_all.head(30))
with st.expander("קבוצות שנמצאו"):
    st.write(df_all["group"].unique().tolist())

# ── CLASSIFY ──────────────────────────────────────────────────────────────────
def df_groups(keywords):
    pat = "|".join(keywords)
    return df_all[df_all["group"].str.contains(pat, case=False, na=False)]

tph_df    = df_groups(["petroleum","tph","hydrocarbon"])
metals_df = df_groups(["metal","cation","extractable"])
pfas_df   = df_groups(["perfluor","pfas","fluorin"])
voc_df    = df_groups(["voc","svoc","btex","aromatic","halogenated","volatile",
                        "alcohol","aldehyde","ketone","phenol","pah","aniline",
                        "nitro","phthalate","pesticide","pcb","other"])

# ── BUILD ─────────────────────────────────────────────────────────────────────
wb_out = Workbook()
wb_out.remove(wb_out.active)

if not tph_df.empty:
    write_tph_sheet(wb_out.create_sheet("TPH"), tph_df, thresh_dict, tier1_col)
    st.info(f"✅ TPH: {tph_df['sample_id'].nunique()} קידוחים")
if not metals_df.empty:
    write_metals_sheet(wb_out.create_sheet("Metals"), metals_df, thresh_dict, tier1_col)
    st.info(f"✅ Metals: {metals_df['sample_id'].nunique()} קידוחים")
if not voc_df.empty:
    write_voc_sheet(wb_out.create_sheet("VOC+SVOC"), voc_df, thresh_dict, tier1_col)
    st.info(f"✅ VOC+SVOC: {voc_df['sample_id'].nunique()} קידוחים")
if not pfas_df.empty:
    write_pfas_sheet(wb_out.create_sheet("PFAS"), pfas_df, thresh_dict, tier1_col)
    st.info(f"✅ PFAS: {pfas_df['sample_id'].nunique()} קידוחים")

if not wb_out.sheetnames:
    wb_out.create_sheet("Results")
    st.warning("לא זוהו קבוצות — בדוק שמות קבוצות בקבצי ALS")

# ── DOWNLOAD ──────────────────────────────────────────────────────────────────
st.markdown("---")
buf = io.BytesIO()
wb_out.save(buf); buf.seek(0)
st.download_button(
    label="⬇️ הורד קובץ Excel מעובד",
    data=buf, file_name="soil_report.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=True
)
