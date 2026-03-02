import streamlit as st
import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import io, re

st.set_page_config(page_title="דוח סקר קרקע", layout="wide", page_icon="🧪")
st.title("🧪 מערכת עיבוד תוצאות מעבדה")
st.markdown("---")

# ── STYLES ──────────────────────────────────────────────────────────────────
YELLOW_FILL   = PatternFill("solid", fgColor="FFFF00")
ORANGE_FILL   = PatternFill("solid", fgColor="FFC000")
HDR_BLUE_FILL = PatternFill("solid", fgColor="B7D7F0")   # metals
HDR_CYAN_FILL = PatternFill("solid", fgColor="00B0F0")   # PFAS
NO_FILL       = PatternFill("none")

def thin_border():
    s = Side(style="thin", color="000000")
    return Border(left=s, right=s, top=s, bottom=s)

def hdr_font(sz=11):   return Font(bold=True, name="Arial", size=sz)
def data_font(sz=10):  return Font(name="Arial", size=sz)
def center():          return Alignment(horizontal="center", vertical="center", wrap_text=True)

def style_hdr(cell, fill=None, sz=11):
    cell.font = hdr_font(sz)
    cell.alignment = center()
    cell.border = thin_border()
    if fill: cell.fill = fill

def style_data(cell, hl=None, sz=10):
    cell.font = Font(bold=bool(hl), name="Arial", size=sz)
    cell.alignment = center()
    cell.border = thin_border()
    if hl == "tier1": cell.fill = ORANGE_FILL
    elif hl == "vsl": cell.fill = YELLOW_FILL

# ── HELPERS ──────────────────────────────────────────────────────────────────
def norm(s):
    s = "" if s is None else str(s).strip().lower()
    return re.sub(r"\s+", " ", s.replace("\xa0"," "))

def to_float(v):
    s = str(v).strip() if v is not None else ""
    s2 = s.lstrip("<>").strip()
    try: return float(s2)
    except: return None

def sort_key(sid):
    m = re.match(r"S-?(\d+)", str(sid), re.I)
    return int(m.group(1)) if m else 9999

def check_exceed(val_str, vsl, tier1):
    """Returns 'tier1', 'vsl', or None"""
    f = to_float(val_str)
    if f is None: return None
    if str(val_str).strip().startswith("<"): return None  # below LOR = no exceedance
    try:
        if tier1 is not None and pd.notna(tier1) and float(tier1) > 0 and f > float(tier1):
            return "tier1"
        if vsl is not None and pd.notna(vsl) and float(vsl) > 0 and f > float(vsl):
            return "vsl"
    except: pass
    return None

# ── METAL NAME MAPPING (ALS full name → symbol) ─────────────────────────────
METAL_MAP = {
    "aluminium":"Al","aluminum":"Al","antimony":"Sb","arsenic":"As","barium":"Ba",
    "beryllium":"Be","bismuth":"Bi","boron":"B","cadmium":"Cd","calcium":"Ca",
    "chromium":"Cr","cobalt":"Co","copper":"Cu","iron":"Fe","lead":"Pb",
    "lithium":"Li","magnesium":"Mg","manganese":"Mn","mercury":"Hg","nickel":"Ni",
    "potassium":"K","selenium":"Se","silver":"Ag","sodium":"Na","vanadium":"V",
    "zinc":"Zn","molybdenum":"Mo","tin":"Sn","titanium":"Ti","tungsten":"W",
    "strontium":"Sr","thallium":"Tl","uranium":"U","cerium":"Ce","lanthanum":"La",
}
METALS_ORDER = ["Al","Sb","As","Ba","Be","Bi","B","Cd","Ca","Cr","Co","Cu","Fe",
                "Pb","Li","Mg","Mn","Hg","Ni","K","Se","Ag","Na","V","Zn"]

# ── ALS PARSER ───────────────────────────────────────────────────────────────
def parse_als_file(file_bytes, filename):
    try:
        wb = load_workbook(io.BytesIO(file_bytes), data_only=True)
    except Exception as e:
        return None, str(e)

    main = next((wb[n] for n in wb.sheetnames if "Client" in n and "SOIL" in n), wb.worksheets[0])
    rows = list(main.iter_rows(values_only=True))

    # sample IDs row
    sid_idx = next((i for i,r in enumerate(rows) if any("Client Sample ID" in str(v) for v in r if v)), None)
    if sid_idx is None: return None, "לא נמצאה שורת Sample IDs"
    col2sample = {ci: str(v).strip() for ci,v in enumerate(rows[sid_idx]) if v and v != "Client Sample ID"}

    # LOR row (col index 3)
    # Parameter header
    ph_idx = next((i for i,r in enumerate(rows) if r and r[0]=="Parameter"), None)
    if ph_idx is None: return None, "לא נמצאה שורת Parameter"

    # CAS row — look for a row where col0 is "CAS" or "CAS Number"
    cas_row_data = {}
    for row in rows[ph_idx+1:]:
        if row and str(row[0]).strip().upper().startswith("CAS"):
            for ci, v in enumerate(row):
                if ci in col2sample and v: cas_row_data[ci] = v
            break

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
            val = row[ci] if ci < len(row) else None
            m = re.match(r"^(S-?\d+[A-Za-z0-9]*)\s*\(([0-9.]+)\)", sname)
            sid, depth = (m.group(1), float(m.group(2))) if m else (sname, None)
            rs = str(val).strip() if val is not None else ""
            result = None
            if rs.startswith("<"):
                result = 0.0
            elif rs and rs != "None":
                try: result = float(rs)
                except: result = None

            if result is not None:
                records.append({
                    "sample_id": sid, "depth": depth,
                    "compound": str(param).strip(),
                    "compound_lower": norm(param),
                    "unit": str(unit).strip() if unit else "mg/kg",
                    "lor": lor,
                    "result": result, "result_str": rs,
                    "group": group, "source": filename
                })

    if not records: return None, "לא נמצאו נתונים"
    return pd.DataFrame(records), None

# ── SHEET WRITERS ─────────────────────────────────────────────────────────────

def write_tph_sheet(ws, df):
    """TPH: שם קידוח | עומק | אנליזה | DRO | ORO | Total — rows 1-5 header"""
    metals_cols = ["TPH DRO","TPH ORO","Total TPH"]
    hdr_vals = ["שם קידוח","עומק","אנליזה"] + metals_cols
    for ci, h in enumerate(hdr_vals, 1):
        style_hdr(ws.cell(1, ci, h))

    sub = {
        "יחידות": {"TPH DRO":"mg/kg","TPH ORO":"mg/kg","Total TPH":"mg/kg"},
        "CAS":     {"TPH DRO":"C10-C40","TPH ORO":"C10-C40","Total TPH":"C10-C40"},
        "VSL":     {"TPH DRO":350,"TPH ORO":350,"Total TPH":350},
        "TIER 1":  {"TPH DRO":350,"TPH ORO":350,"Total TPH":350},
    }
    for ri, lbl in enumerate(["יחידות","CAS","VSL","TIER 1"], 2):
        style_hdr(ws.cell(ri, 2, lbl))
        for ci, col in enumerate(metals_cols, 4):
            style_hdr(ws.cell(ri, ci, sub[lbl][col]))

    # pivot data: DRO, ORO per sample+depth
    pivoted = {}
    for _, r in df.iterrows():
        k = (r["sample_id"], r["depth"])
        if k not in pivoted: pivoted[k] = {"DRO":"","ORO":"","DRO_f":None,"ORO_f":None}
        c = r["compound_lower"]
        if "dro" in c or ("c10" in c and "c28" in c) or ("c10" in c and "c40" in c and "oro" not in c):
            pivoted[k]["DRO"] = r["result_str"]; pivoted[k]["DRO_f"] = r["result"]
        elif "oro" in c or ("c24" in c) or ("c28" in c and "c40" in c):
            pivoted[k]["ORO"] = r["result_str"]; pivoted[k]["ORO_f"] = r["result"]

    VSL, TIER1 = 350, 350
    ri = 6
    prev_sid = None
    for (sid, depth), v in sorted(pivoted.items(), key=lambda x:(sort_key(x[0][0]), x[0][1] or 0)):
        dro_f = v["DRO_f"] or 0
        oro_f = v["ORO_f"] or 0
        total_f = dro_f + oro_f
        both_lor = v["DRO"].startswith("<") and v["ORO"].startswith("<")
        total_str = f"<{total_f:.0f}" if both_lor else f"{total_f:.0f}"

        hl = check_exceed(total_str, VSL, TIER1)

        sid_val = sid if sid != prev_sid else None
        prev_sid = sid
        for ci, val in enumerate([sid_val, depth, None, v["DRO"], v["ORO"], total_str], 1):
            style_data(ws.cell(ri, ci, val), hl if ci in [4,5,6] else None)
        ri += 1

    for col, w in zip("ABCDEF", [14,10,14,16,16,16]):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A6"


def write_metals_sheet(ws, df, thresh_metals):
    """
    Wide: שם קידוח | עומק | (blank) | Al | As | ...
    thresh_metals: dict symbol -> {vsl, tier1, cas}
    """
    # build list of metals that actually appear in data
    present = set()
    for _, r in df.iterrows():
        sym = METAL_MAP.get(r["compound_lower"])
        if sym: present.add(sym)

    metals = [m for m in METALS_ORDER if m in present]
    if not metals: metals = list(present)

    hdr = ["שם קידוח","עומק",None] + metals
    for ci, h in enumerate(hdr, 1):
        style_hdr(ws.cell(1, ci, h), HDR_BLUE_FILL)

    # sub-rows
    for ri, lbl in enumerate(["יחידות","CAS","VSL","TIER 1"], 2):
        style_hdr(ws.cell(ri, 2, lbl), HDR_BLUE_FILL)
        for ci, sym in enumerate(metals, 4):
            t = thresh_metals.get(sym, {})
            val = {"יחידות":"mg/kg","CAS":t.get("cas","-"),
                   "VSL":t.get("vsl","-"),"TIER 1":t.get("tier1","-")}[lbl]
            style_hdr(ws.cell(ri, ci, val), HDR_BLUE_FILL)

    # pivot
    pivoted = {}
    for _, r in df.iterrows():
        sym = METAL_MAP.get(r["compound_lower"])
        if not sym: continue
        k = (r["sample_id"], r["depth"])
        if k not in pivoted: pivoted[k] = {}
        pivoted[k][sym] = r["result_str"]

    ri = 6
    prev_sid = None
    for (sid, depth), mvals in sorted(pivoted.items(), key=lambda x:(sort_key(x[0][0]), x[0][1] or 0)):
        sid_val = sid if sid != prev_sid else None
        prev_sid = sid
        for ci, val in enumerate([sid_val, depth, None], 1):
            style_data(ws.cell(ri, ci, val))
        for ci, sym in enumerate(metals, 4):
            val = mvals.get(sym, "")
            t = thresh_metals.get(sym, {})
            hl = check_exceed(val, t.get("vsl"), t.get("tier1"))
            style_data(ws.cell(ri, ci, val), hl)
        ri += 1

    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 10
    ws.column_dimensions["C"].width = 4
    for ci in range(4, len(metals)+4):
        ws.column_dimensions[get_column_letter(ci)].width = 11
    ws.freeze_panes = "D6"


def write_pfas_sheet(ws, df, thresh_pfas):
    """
    Vertical: תרכובת | CAS | ערך סף | יחידות | LOR | שם קידוח → sample cols
    """
    samples = sorted(df["sample_id"].unique(), key=sort_key)
    sample_depth = {r["sample_id"]: r["depth"] for _,r in df.iterrows()}

    fixed = ["שם התרכובת","CAS","ערך סף","יחידות","LOR","שם הקידוח"]
    all_cols = fixed + samples

    for ci, h in enumerate(all_cols, 1):
        style_hdr(ws.cell(1, ci, h), HDR_CYAN_FILL)
    cell = ws.cell(2, 6, "עומק")
    style_hdr(cell, HDR_CYAN_FILL)
    for ci, sid in enumerate(samples, 7):
        style_hdr(ws.cell(2, ci, sample_depth.get(sid, "")), HDR_CYAN_FILL)

    compounds = list(df["compound"].unique())
    for row_i, cmp in enumerate(compounds, 3):
        df_c = df[df["compound"]==cmp]
        key  = norm(cmp)
        t    = thresh_pfas.get(key, {})
        vsl  = t.get("vsl","")
        cas  = t.get("cas","-")
        unit = df_c.iloc[0]["unit"] if not df_c.empty else "µg/kg"
        lor  = df_c.iloc[0]["lor"]  if not df_c.empty else ""

        for ci, val in enumerate([cmp, cas, vsl, unit, lor, None], 1):
            style_data(ws.cell(row_i, ci, val))

        for ci, sid in enumerate(samples, 7):
            row_sid = df_c[df_c["sample_id"]==sid]
            rs = row_sid.iloc[0]["result_str"] if not row_sid.empty else ""
            hl = check_exceed(rs, vsl, None) if vsl else None
            style_data(ws.cell(row_i, ci, rs), hl)

    ws.column_dimensions["A"].width = 50
    for ci in range(2, len(all_cols)+1):
        ws.column_dimensions[get_column_letter(ci)].width = 13
    ws.freeze_panes = "G3"


def write_voc_sheet(ws, df, thresh_voc):
    """
    Vertical: קבוצה | קבוצה | תרכובת | CAS | VSL | TIER 1 | יחידות | → sample cols
    """
    samples = sorted(df["sample_id"].unique(), key=sort_key)
    sample_depth = {}
    for _,r in df.iterrows():
        if r["sample_id"] not in sample_depth:
            sample_depth[r["sample_id"]] = r["depth"]

    fixed = ["קבוצה","קבוצה","שם התרכובת","CAS","VSL","TIER 1","יחידות","שם קידוח"]
    all_cols = fixed + samples

    for ci, h in enumerate(all_cols, 1):
        style_hdr(ws.cell(1, ci, h), sz=9)
    cell = ws.cell(2, 8, "עומק")
    style_hdr(cell, sz=9)
    for ci, sid in enumerate(samples, 9):
        style_hdr(ws.cell(2, ci, sample_depth.get(sid,"")), sz=9)

    # compounds in order
    seen = {}
    for _,r in df.iterrows():
        if r["compound"] not in seen:
            seen[r["compound"]] = r["group"]

    for row_i, (cmp, grp) in enumerate(seen.items(), 3):
        key = norm(cmp)
        t   = thresh_voc.get(key, {})
        vsl = t.get("vsl",""); tier1 = t.get("tier1",""); cas = t.get("cas","-")
        df_c = df[df["compound"]==cmp]
        unit = df_c.iloc[0]["unit"] if not df_c.empty else "mg/kg"

        for ci, val in enumerate([None, grp, cmp, cas, vsl, tier1, unit, None], 1):
            style_data(ws.cell(row_i, ci, val), sz=9)

        for ci, sid in enumerate(samples, 9):
            row_sid = df_c[df_c["sample_id"]==sid]
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
st.sidebar.header("⚙️ הגדרות")
st.sidebar.markdown("🟡 חריגה מעל VSL")
st.sidebar.markdown("🟠 חריגה מעל TIER 1")

# ── FILE UPLOADS ──────────────────────────────────────────────────────────────
col1, col2 = st.columns(2)
with col1:
    st.subheader("📁 קבצי ערכי סף")
    threshold_files = st.file_uploader(
        "העלה קבצי ערכי סף (אפשר כמה — אחד לכל קבוצת מזהמים)",
        type=["xlsx","xls"], accept_multiple_files=True, key="thr"
    )
with col2:
    st.subheader("📂 קבצי נתונים מ-ALS")
    data_files = st.file_uploader(
        "העלה קבצי ALS — אפשר כמה ביחד",
        type=["xlsx","xls"], accept_multiple_files=True, key="data"
    )

if not threshold_files:
    st.info("👆 העלה לפחות קובץ ערכי סף אחד וקבצי ALS כדי להתחיל")
    st.stop()
if not data_files:
    st.warning("⚠️ העלה קבצי נתונים מ-ALS")
    st.stop()

# ── READ THRESHOLDS ───────────────────────────────────────────────────────────
# supports multiple threshold files
thresh_metals = {}  # symbol -> {vsl, tier1, cas}
thresh_pfas   = {}  # norm(name) -> {vsl, tier1, cas}
thresh_voc    = {}  # norm(name) -> {vsl, tier1, cas}

for tf in threshold_files:
    try:
        dft = pd.read_excel(tf)
        dft.columns = [str(c).lower().strip() for c in dft.columns]
        comp_col = next((c for c in dft.columns if c in
            ["compound","chemical","chemical name","parameter","analyte","name","תרכובת","חומר"]), None)
        vsl_col  = next((c for c in dft.columns if "vsl" in c), None)
        t1_col   = next((c for c in dft.columns if "tier" in c and "1" in c), None)
        cas_col  = next((c for c in dft.columns if c.startswith("cas")), None)
        sym_col  = next((c for c in dft.columns if c in ["symbol","sym","סמל","abbr"]), None)

        if not comp_col or not vsl_col:
            st.warning(f"⚠️ {tf.name}: לא נמצאו עמודות נדרשות. עמודות: {list(dft.columns)}")
            continue

        for _, r in dft.iterrows():
            name = str(r[comp_col]).strip()
            if not name or name.lower()=="nan": continue
            key   = norm(name)
            vsl   = r.get(vsl_col)
            tier1 = r.get(t1_col)  if t1_col  else None
            cas   = r.get(cas_col) if cas_col  else "-"
            sym   = r.get(sym_col) if sym_col  else None

            entry = {"vsl": vsl if pd.notna(vsl) else None,
                     "tier1": tier1 if (tier1 is not None and pd.notna(tier1)) else None,
                     "cas": str(cas).strip() if cas else "-"}

            # metals: try to match by symbol
            metal_sym = sym or METAL_MAP.get(key)
            if metal_sym and metal_sym in METALS_ORDER:
                thresh_metals[metal_sym] = entry
            else:
                thresh_pfas[key] = entry
                thresh_voc[key]  = entry

        st.success(f"✅ {tf.name} — {len(dft)} ערכי סף")
    except Exception as e:
        st.error(f"שגיאה ב-{tf.name}: {e}")

# ── PARSE ALS FILES ───────────────────────────────────────────────────────────
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

# show groups found
groups_found = df_all["group"].unique().tolist()
with st.expander("קבוצות שנמצאו בקבצי ALS"):
    st.write(groups_found)

# ── CLASSIFY GROUPS ───────────────────────────────────────────────────────────
def df_groups(keywords):
    pat = "|".join(keywords)
    return df_all[df_all["group"].str.contains(pat, case=False, na=False)]

tph_df    = df_groups(["petroleum","tph","hydrocarbon"])
metals_df = df_groups(["metal","cation","extractable"])
pfas_df   = df_groups(["perfluor","pfas","fluorin"])
voc_df    = df_groups(["voc","svoc","btex","aromatic","halogenated","volatile",
                        "alcohol","aldehyde","ketone","phenol","pah","aniline",
                        "nitro","phthalate","pesticide","pcb","other"])

# ── BUILD WORKBOOK ────────────────────────────────────────────────────────────
wb_out = Workbook()
wb_out.remove(wb_out.active)

if not tph_df.empty:
    write_tph_sheet(wb_out.create_sheet("TPH"), tph_df)
    st.info(f"✅ גיליון TPH: {len(tph_df)} שורות")

if not metals_df.empty:
    write_metals_sheet(wb_out.create_sheet("Metals"), metals_df, thresh_metals)
    st.info(f"✅ גיליון Metals: {len(metals_df)} שורות")

if not voc_df.empty:
    write_voc_sheet(wb_out.create_sheet("VOC+SVOC"), voc_df, thresh_voc)
    st.info(f"✅ גיליון VOC+SVOC: {len(voc_df)} שורות")

if not pfas_df.empty:
    write_pfas_sheet(wb_out.create_sheet("PFAS"), pfas_df, thresh_pfas)
    st.info(f"✅ גיליון PFAS: {len(pfas_df)} שורות")

if not wb_out.sheetnames:
    wb_out.create_sheet("Results")
    st.warning("לא זוהו קבוצות — בדוק את שמות הקבוצות בקבצי ALS")

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
