import streamlit as st
import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import io
import re
from docx import Document
from docx.shared import Pt, RGBColor, Twips
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.enum.section import WD_ORIENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

st.set_page_config(page_title="דוח סקר קרקע", layout="wide", page_icon="🧪")
st.title("🧪 מערכת עיבוד תוצאות מעבדה")
st.caption("v3.8 - RTL sheets, number formatting, compound names left-aligned")
st.markdown("---")

tab_excel, tab_word = st.tabs(["📊 יצוא Excel", "📄 יצוא Word"])

# ── העלאת קבצים - מחוץ לטאבים ────────────────────────────────────────────────
c1,c2=st.columns(2)
with c1:
    st.subheader("📁 קובץ ערכי סף")
    thr_file=st.file_uploader("העלה קובץ ערכי הסף המאוחד",type=["xlsx","xls"],key="thr")
with c2:
    st.subheader("📂 קבצי ALS")
    data_files=st.file_uploader("העלה קבצי ALS",type=["xlsx","xls"],accept_multiple_files=True,key="data")

if not thr_file: st.info("👆 העלה קובץ ערכי סף וקבצי ALS"); st.stop()
if not data_files: st.warning("⚠️ העלה קבצי ALS"); st.stop()

thresh_dict=load_threshold_file(thr_file.read())
st.success(f"✅ {len(thresh_dict)-1} ערכי סף | {land_use} | {aquifer} | {depth}")

all_data=[]
for f in data_files:
    df,err=parse_als_file(f.read(),f.name)
    if err: st.warning(f"⚠️ {f.name}: {err}")
    else: all_data.append(df); st.success(f"✅ {f.name} — {len(df)} תוצאות")

if not all_data: st.error("לא נטענו נתונים."); st.stop()
df_all=pd.concat(all_data,ignore_index=True)

def dg(kw): return df_all[df_all["group"].str.contains("|".join(kw),case=False,na=False)]
tph_df   =dg(["petroleum","tph","hydrocarbon"])
metals_df=dg(["metal","cation","extractable"])
pfas_df  =dg(["perfluor","pfas","fluorin"])
voc_df   =dg(["voc","svoc","btex","aromatic","halogenated","volatile",
               "alcohol","aldehyde","ketone","phenol","pah","aniline",
               "nitro","phthalate","pesticide","pcb","other"])

st.markdown("---")

with tab_excel:
    with st.expander(f"👁️ תצוגה מקדימה ({len(df_all)} שורות)"): st.dataframe(df_all.head(30),use_container_width=True)
    with st.expander("קבוצות שנמצאו"): st.write(df_all["group"].unique().tolist())
    wb_out=Workbook(); wb_out.remove(wb_out.active)
    if not tph_df.empty:
        write_tph_sheet(wb_out.create_sheet("TPH"),tph_df,thresh_dict,t1col,t1lbl)
        st.info(f"✅ TPH: {tph_df['sample_id'].nunique()} קידוחים")
    if not metals_df.empty:
        write_metals_sheet(wb_out.create_sheet("Metals"),metals_df,thresh_dict,t1col,t1lbl)
        st.info(f"✅ Metals: {metals_df['sample_id'].nunique()} קידוחים")
    if not voc_df.empty:
        write_voc_sheet(wb_out.create_sheet("VOC+SVOC"),voc_df,thresh_dict,t1col,t1lbl)
        st.info(f"✅ VOC+SVOC: {voc_df['sample_id'].nunique()} קידוחים")
    if not pfas_df.empty:
        write_pfas_sheet(wb_out.create_sheet("PFAS"),pfas_df,thresh_dict,t1col,t1lbl)
        st.info(f"✅ PFAS: {pfas_df['sample_id'].nunique()} קידוחים")
    if not wb_out.sheetnames:
        wb_out.create_sheet("Results"); st.warning("לא זוהו קבוצות")
    st.markdown("---")
    buf=io.BytesIO(); wb_out.save(buf); buf.seek(0)
    st.download_button("⬇️ הורד קובץ Excel מעובד",data=buf,file_name="soil_report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",use_container_width=True)

# ── WORD EXPORT FUNCTIONS ────────────────────────────────────────────────────


# ── צבעים ────────────────────────────────────────────────────────────────────
COLOR_YELLOW = "FFFF00"
COLOR_ORANGE = "FFC000"
COLOR_HEADER = "B7D7F0"
COLOR_HEADER2 = "00B0F0"
COLOR_WHITE   = "FFFFFF"
COLOR_LEGEND_BG = "F2F2F2"

# ── גדלי דף (EMU: 1 inch = 914400) ──────────────────────────────────────────
PAGE_SIZES = {
    # (width_portrait, height_portrait) in Twips (1/20 pt, 1 inch = 1440 twips)
    "A4":      (11906, 16838),   # 210mm × 297mm
    "Tabloid": (17280, 22320),   # 12in × 15.5in  (ANSI B)
}
MARGIN_NORMAL = 720   # 0.5 inch in twips
MARGIN_NARROW = 540   # 0.375 inch

def _set_cell_bg(cell, hex_color):
    """צבע רקע לתא"""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hex_color)
    # הסר shd קיים אם יש
    for old in tcPr.findall(qn('w:shd')):
        tcPr.remove(old)
    tcPr.append(shd)

def _set_cell_borders(cell, color="000000", size=4):
    """גבול דק לכל צדדי תא"""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    for old in tcPr.findall(qn('w:tcBorders')):
        tcPr.remove(old)
    tcBorders = OxmlElement('w:tcBorders')
    for side in ('top', 'left', 'bottom', 'right'):
        el = OxmlElement(f'w:{side}')
        el.set(qn('w:val'), 'single')
        el.set(qn('w:sz'), str(size))
        el.set(qn('w:space'), '0')
        el.set(qn('w:color'), color)
        tcBorders.append(el)
    tcPr.append(tcBorders)

def _cell_text(cell, text, bold=False, size=8, color=None, align="center",
               rtl=False, bg=None, valign="center"):
    """כתוב טקסט בתא עם עיצוב"""
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    if bg:
        _set_cell_bg(cell, bg)
    _set_cell_borders(cell)

    # נקה תוכן קיים
    for p in cell.paragraphs:
        for r in p.runs:
            r.text = ""

    p = cell.paragraphs[0]
    p.clear()
    if align == "center":
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    elif align == "right":
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    else:
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT

    # RTL
    if rtl:
        pPr = p._p.get_or_add_pPr()
        bidi = OxmlElement('w:bidi')
        pPr.append(bidi)

    run = p.add_run(str(text) if text is not None else "")
    run.bold = bold
    run.font.name = "Arial"
    run.font.size = Pt(size)
    if color:
        run.font.color.rgb = RGBColor.from_string(color)

def _set_section_props(section, page_size, landscape):
    """הגדר גודל דף וכיוון"""
    w_twips, h_twips = PAGE_SIZES[page_size]
    if landscape:
        section.page_width  = Twips(h_twips)
        section.page_height = Twips(w_twips)
        section.orientation = WD_ORIENT.LANDSCAPE
    else:
        section.page_width  = Twips(w_twips)
        section.page_height = Twips(h_twips)
        section.orientation = WD_ORIENT.PORTRAIT

    margin = MARGIN_NARROW
    section.top_margin    = Twips(margin)
    section.bottom_margin = Twips(margin)
    section.left_margin   = Twips(margin)
    section.right_margin  = Twips(margin)

def _content_width_twips(page_size, landscape):
    """רוחב תוכן שמיש בטוויפס"""
    w_twips, h_twips = PAGE_SIZES[page_size]
    if landscape:
        return h_twips - 2 * MARGIN_NARROW
    else:
        return w_twips - 2 * MARGIN_NARROW

def _add_section_break(doc, page_size, landscape):
    """הוסף מעבר דף + section חדש עם הגדרות"""
    new_section = doc.add_section()
    _set_section_props(new_section, page_size, landscape)
    return new_section

def _add_title_paragraph(doc, title, part_str, rtl=True):
    """כותרת טבלה מעל כל דף"""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT if rtl else WD_ALIGN_PARAGRAPH.LEFT
    pPr = p._p.get_or_add_pPr()
    if rtl:
        bidi = OxmlElement('w:bidi')
        pPr.append(bidi)
    sp = p.paragraph_format
    sp.space_before = Pt(0)
    sp.space_after  = Pt(3)

    run = p.add_run(f"{title}   {part_str}")
    run.bold = True
    run.font.name = "Arial"
    run.font.size = Pt(11)
    return p

def _add_legend(doc, has_yellow, has_orange):
    """מקרא בתחתית טבלה"""
    if not has_yellow and not has_orange:
        return
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    pPr = p._p.get_or_add_pPr()
    bidi = OxmlElement('w:bidi')
    pPr.append(bidi)
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after  = Pt(0)

    if has_yellow:
        run = p.add_run("■ ")
        run.font.name = "Arial"
        run.font.size = Pt(8)
        run.font.color.rgb = RGBColor.from_string(COLOR_YELLOW)
        run2 = p.add_run("בצהוב - חריגה מערך הסף VSL    ")
        run2.font.name = "Arial"
        run2.font.size = Pt(8)

    if has_orange:
        run3 = p.add_run("■ ")
        run3.font.name = "Arial"
        run3.font.size = Pt(8)
        run3.font.color.rgb = RGBColor.from_string(COLOR_ORANGE)
        run4 = p.add_run("בכתום - חריגה מערך הסף TIER 1")
        run4.font.name = "Arial"
        run4.font.size = Pt(8)

def _twips_to_emu(t):
    return int(t * 914400 / 1440)

def _set_table_width(table, width_twips):
    """הגדר רוחב טבלה"""
    tbl = table._tbl
    tblPr = tbl.find(qn('w:tblPr'))
    if tblPr is None:
        tblPr = OxmlElement('w:tblPr')
        tbl.insert(0, tblPr)
    for old in tblPr.findall(qn('w:tblW')):
        tblPr.remove(old)
    tblW = OxmlElement('w:tblW')
    tblW.set(qn('w:w'), str(width_twips))
    tblW.set(qn('w:type'), 'dxa')
    tblPr.append(tblW)

def _set_col_width(cell, width_twips):
    """רוחב תא"""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    for old in tcPr.findall(qn('w:tcW')):
        tcPr.remove(old)
    tcW = OxmlElement('w:tcW')
    tcW.set(qn('w:w'), str(width_twips))
    tcW.set(qn('w:type'), 'dxa')
    tcPr.append(tcW)

# ═══════════════════════════════════════════════════════════════════════════════
#  פונקציות בניית טבלאות לפי סוג
# ═══════════════════════════════════════════════════════════════════════════════

def _build_tph_table_data(df, thresh_dict, t1col, t1lbl):
    """
    מחזיר: headers (list of list), rows (list of dict with 'values' and 'colors')
    """
    import re

    def norm(s):
        s = "" if s is None else str(s).strip().lower()
        return re.sub(r"\s+", " ", s.replace("\xa0", " "))

    def sort_key(sid):
        m = re.match(r"S-?(\d+)", str(sid), re.I)
        return int(m.group(1)) if m else 9999

    def to_float(v):
        s = str(v).strip() if v is not None else ""
        try: return float(s.lstrip("<>").strip())
        except: return None

    def is_dro(c):
        if "(dro)" in c or "- dro" in c or c.strip()=="dro": return True
        if "(oro)" in c or "- oro" in c: return False
        if "c10" in c and "c28" in c and "c40" not in c: return True
        return False

    def is_oro(c):
        if "(oro)" in c or "- oro" in c or c.strip()=="oro": return True
        if "(dro)" in c or "- dro" in c: return False
        if "c24" in c and "c40" in c: return True
        if "c28" in c and "c40" in c: return True
        return False

    def is_total(c):
        if any(x in c for x in ["(dro)","(oro)","- dro","- oro"]): return False
        if "c10" in c and "c40" in c: return True
        if "total" in c and ("tph" in c or "hydrocarbon" in c): return True
        return False

    def get_thresh_local(compound, thresh_dict, t1col):
        from word_export import _match_thresh_simple
        t = _match_thresh_simple(compound, thresh_dict)
        return t.get("VSL"), t.get(t1col), t.get("cas", "-")

    vsl_d,t1_d,_ = get_thresh_local("C10 - C28 Fraction (DRO)", thresh_dict, t1col)
    vsl_o,t1_o,_ = get_thresh_local("C24 - C40 Fraction (ORO)", thresh_dict, t1col)
    vsl_t,t1_t,_ = get_thresh_local("TPH - DRO + ORO (Tier 1)", thresh_dict, t1col)
    vv = [v for v in [vsl_d,vsl_o,vsl_t] if v]
    tt = [v for v in [t1_d,t1_o,t1_t] if v]
    vsl_tot = min(float(v) for v in vv) if vv else 350
    t1_tot  = min(float(v) for v in tt) if tt else 350

    headers = [
        ["שם קידוח", "עומק", "TPH DRO", "TPH ORO", "Total TPH"],
        ["", "יחידות", "mg/kg", "mg/kg", "mg/kg"],
        ["", "CAS", "C10-C40", "C10-C40", "C10-C40"],
        ["", "VSL", str(vsl_tot), str(vsl_tot), str(vsl_tot)],
        ["", t1lbl.replace("\n", " "), str(t1_tot), str(t1_tot), str(t1_tot)],
    ]

    pivoted = {}
    for _, r in df.iterrows():
        k = (r["sample_id"], r["depth"])
        if k not in pivoted:
            pivoted[k] = {"DRO":"","ORO":"","TOT":"","DRO_f":None,"ORO_f":None,"DRO_lor":None,"ORO_lor":None}
        c = r["compound_lower"]
        if is_dro(c) and not pivoted[k]["DRO"]:
            pivoted[k]["DRO"] = r["result_str"]
            pivoted[k]["DRO_f"] = r["result"]
            pivoted[k]["DRO_lor"] = r.get("lor_val")
        elif is_oro(c) and not pivoted[k]["ORO"]:
            pivoted[k]["ORO"] = r["result_str"]
            pivoted[k]["ORO_f"] = r["result"]
            pivoted[k]["ORO_lor"] = r.get("lor_val")
        elif is_total(c) and not pivoted[k]["TOT"]:
            pivoted[k]["TOT"] = r["result_str"]

    def check_exceed(val_str, vsl, tier1):
        if not val_str or str(val_str).strip().startswith("<"): return None
        f = to_float(val_str)
        if f is None: return None
        try:
            t1f = float(tier1) if tier1 is not None else None
            vf  = float(vsl)   if vsl   is not None else None
            if t1f and t1f > 0 and f > t1f: return "tier1"
            if vf  and vf  > 0 and f > vf:  return "vsl"
        except: pass
        return None

    rows_out = []
    prev_sid = None
    for (sid,depth_val),v in sorted(pivoted.items(), key=lambda x:(sort_key(x[0][0]), x[0][1] or 0)):
        if v["TOT"]: total_s = v["TOT"]
        else:
            dro_lor = v["DRO"] and str(v["DRO"]).startswith("<")
            oro_lor = v["ORO"] and str(v["ORO"]).startswith("<")
            dro_empty = not v["DRO"]; oro_empty = not v["ORO"]
            dro_num = (v["DRO_lor"] if dro_lor and v["DRO_lor"] is not None else (v["DRO_f"] or 0))
            oro_num = (v["ORO_lor"] if oro_lor and v["ORO_lor"] is not None else (v["ORO_f"] or 0))
            total_f = dro_num + oro_num
            if (dro_lor or dro_empty) and (oro_lor or oro_empty) and not (dro_empty and oro_empty):
                total_s = f"<{total_f:.0f}"
            else:
                total_s = f"{total_f:.0f}"

        hl_dro   = check_exceed(v["DRO"],   vsl_tot, t1_tot)
        hl_oro   = check_exceed(v["ORO"],   vsl_tot, t1_tot)
        hl_total = check_exceed(total_s,    vsl_tot, t1_tot)

        sid_display = sid if sid != prev_sid else ""
        prev_sid = sid
        rows_out.append({
            "values": [sid_display, str(depth_val) if depth_val else "", v["DRO"], v["ORO"], total_s],
            "colors": [None, None, hl_dro, hl_oro, hl_total],
        })

    return headers, rows_out


def _build_metals_table_data(df, thresh_dict, t1col, t1lbl):
    import re
    METAL_MAP = {
        "aluminium":"Al","aluminum":"Al","antimony":"Sb","arsenic":"As",
        "barium":"Ba","beryllium":"Be","bismuth":"Bi","boron":"B",
        "cadmium":"Cd","calcium":"Ca","chromium":"Cr","cobalt":"Co",
        "copper":"Cu","iron":"Fe","lead":"Pb","lithium":"Li",
        "magnesium":"Mg","manganese":"Mn","mercury":"Hg","nickel":"Ni",
        "potassium":"K","selenium":"Se","silver":"Ag","sodium":"Na",
        "vanadium":"V","zinc":"Zn","molybdenum":"Mo","tin":"Sn",
        "titanium":"Ti","strontium":"Sr","thallium":"Tl",
        "phosphorus":"P","sulphur":"S","silicon":"Si",
    }
    METALS_ORDER = ["Al","Sb","As","Ba","Be","Bi","B","Cd","Ca","Cr","Co","Cu","Fe",
                    "Pb","Li","Mg","Mn","Hg","Ni","K","Se","Ag","Na","V","Zn"]
    THRESH_METAL_MAP = {
        "aluminum":"Al","antimony (metallic)":"Sb","antimony":"Sb",
        "arsenic, inorganic":"As","arsenic":"As","barium":"Ba",
        "beryllium and compounds":"Be","beryllium":"Be",
        "boron and borates only":"B","boron":"B",
        "cadmium (water) source: water and air":"Cd","cadmium":"Cd",
        "calcium":"Ca","chromium, total":"Cr","chromium":"Cr","cobalt":"Co",
        "copper":"Cu","iron":"Fe","lead and compounds":"Pb","lead":"Pb",
        "lithium":"Li","magnesium":"Mg","manganese (non-diet)":"Mn","manganese":"Mn",
        "mercuric chloride (and other mercury salts)":"Hg","mercury":"Hg",
        "nickel soluble salts":"Ni","nickel":"Ni","potassium":"K",
        "selenium":"Se","silver":"Ag","sodium":"Na",
        "vanadium and compounds":"V","vanadium":"V",
        "zinc and compounds":"Zn","zinc":"Zn","molybdenum":"Mo","tin":"Sn",
        "titanium":"Ti","strontium":"Sr","thallium":"Tl",
        "phosphorus":"P","sulphur":"S","silicon":"Si",
    }

    def norm(s):
        s = "" if s is None else str(s).strip().lower()
        return re.sub(r"\s+", " ", s.replace("\xa0", " "))

    def sort_key(sid):
        m = re.match(r"S-?(\d+)", str(sid), re.I)
        return int(m.group(1)) if m else 9999

    def to_float(v):
        s = str(v).strip() if v is not None else ""
        try: return float(s.lstrip("<>").strip())
        except: return None

    df = df.copy()
    df["sym"] = df["compound_lower"].map(METAL_MAP)
    df = df[df["sym"].notna()]
    if df.empty:
        return [["אין נתוני מתכות"]], []

    present = set(df["sym"].unique())
    metals = [m for m in METALS_ORDER if m in present] + sorted(present - set(METALS_ORDER))

    # build metals thresh
    mt = {}
    for key, v in thresh_dict.items():
        if key == "__CANONICAL_MAP_INTERNAL__": continue
        sym = THRESH_METAL_MAP.get(key)
        if sym and sym not in mt:
            mt[sym] = {"vsl": v.get("VSL"), "tier1": v.get(t1col), "cas": v.get("cas","-")}

    headers = [
        ["שם קידוח", "עומק"] + metals,
        ["", "יחידות"] + ["mg/kg"] * len(metals),
        ["", "CAS"] + [mt.get(m,{}).get("cas","-") for m in metals],
        ["", "VSL"] + [str(mt.get(m,{}).get("vsl","-")) for m in metals],
        ["", t1lbl.replace("\n"," ")] + [str(mt.get(m,{}).get("tier1","-")) for m in metals],
    ]

    pt = df.pivot_table(index=["sample_id","depth"], columns="sym", values="result_str", aggfunc="first")
    pt = pt.reindex(sorted(pt.index, key=lambda x:(sort_key(x[0]), x[1] or 0)))

    def check_exceed(val_str, vsl, tier1):
        if not val_str or str(val_str).strip().startswith("<"): return None
        f = to_float(val_str)
        if f is None: return None
        try:
            t1f = float(tier1) if tier1 is not None else None
            vf  = float(vsl)   if vsl   is not None else None
            if t1f and t1f > 0 and f > t1f: return "tier1"
            if vf  and vf  > 0 and f > vf:  return "vsl"
        except: pass
        return None

    rows_out = []
    prev_sid = None
    for (sid, depth_val), row_data in pt.iterrows():
        sid_display = sid if sid != prev_sid else ""
        prev_sid = sid
        values = [sid_display, str(depth_val) if depth_val else ""]
        colors = [None, None]
        for sym in metals:
            val = row_data.get(sym, "") or ""
            val = "" if str(val) == "nan" else str(val)
            hl = check_exceed(val, mt.get(sym,{}).get("vsl"), mt.get(sym,{}).get("tier1"))
            values.append(val)
            colors.append(hl)
        rows_out.append({"values": values, "colors": colors})

    return headers, rows_out


def _build_generic_table_data(df, thresh_dict, t1col, t1lbl, compound_order_key="compound"):
    """גנרי ל-VOC/SVOC ו-PFAS - מחזיר נתונים גולמיים"""
    import re

    def sort_key(sid):
        m = re.match(r"S-?(\d+)", str(sid), re.I)
        return int(m.group(1)) if m else 9999

    def to_float(v):
        s = str(v).strip() if v is not None else ""
        try: return float(s.lstrip("<>").strip())
        except: return None

    def check_exceed(val_str, vsl, tier1):
        if not val_str or str(val_str).strip().startswith("<"): return None
        f = to_float(val_str)
        if f is None: return None
        try:
            t1f = float(tier1) if tier1 is not None else None
            vf  = float(vsl)   if vsl   is not None else None
            if t1f and t1f > 0 and f > t1f: return "tier1"
            if vf  and vf  > 0 and f > vf:  return "vsl"
        except: pass
        return None

    pairs = sorted(df[["sample_id","depth"]].drop_duplicates().values.tolist(),
                   key=lambda x: (sort_key(x[0]), -(x[1] or 0)))

    compounds = df[compound_order_key].unique().tolist()

    headers = [
        ["שם התרכובת", "CAS", "VSL", t1lbl.replace("\n"," "), "יחידות"] + [f"{sid}\n{d}" for sid,d in pairs],
    ]

    als_data = {}
    for _, r in df.iterrows():
        k = str(r[compound_order_key]).strip()
        als_data.setdefault(k, {})[(r["sample_id"], r["depth"])] = r["result_str"]

    rows_out = []
    for cmp in compounds:
        t = _match_thresh_simple(cmp, thresh_dict)
        vsl   = t.get("VSL")
        tier1 = t.get(t1col)
        cas   = t.get("cas", "-")
        unit  = df[df[compound_order_key]==cmp]["unit"].iloc[0] if len(df[df[compound_order_key]==cmp]) > 0 else "mg/kg"
        values = [cmp, str(cas), str(vsl) if vsl else "-", str(tier1) if tier1 else "-", str(unit)]
        colors = [None, None, None, None, None]
        for sid, depth_val in pairs:
            rs = als_data.get(cmp, {}).get((sid, depth_val), "")
            hl = check_exceed(rs, vsl, tier1)
            values.append(str(rs) if rs else "")
            colors.append(hl)
        rows_out.append({"values": values, "colors": colors})

    return headers, rows_out


def _match_thresh_simple(compound, thresh_dict):
    """חיפוש פשוט במילון ערכי סף"""
    import re
    def norm(s):
        s = "" if s is None else str(s).strip().lower()
        return re.sub(r"\s+", " ", s.replace("\xa0", " "))
    key = norm(compound)
    for k in (key, key.replace(".",","), key.replace(",",".")):
        if k in thresh_dict and k != "__CANONICAL_MAP_INTERNAL__":
            return thresh_dict[k]
    return {}


# ═══════════════════════════════════════════════════════════════════════════════
#  פונקציה ראשית: כתוב טבלה ל-document
# ═══════════════════════════════════════════════════════════════════════════════

def _add_table_to_doc(doc, headers, data_rows, title, page_size, landscape,
                      is_first_section=False):
    """
    הוסף טבלה אחת (או חלק ממנה) למסמך.
    מחלק אוטומטית לדפים לפי מספר שורות שמתאימות.
    """
    if not data_rows:
        return

    n_cols = len(headers[0])
    content_w = _content_width_twips(page_size, landscape)

    # חישוב רוחב עמודות: עמודה ראשונה רחבה יותר
    if n_cols <= 5:
        col_w_first = int(content_w * 0.25)
        col_w_rest  = int((content_w - col_w_first) / max(n_cols - 1, 1))
        col_widths = [col_w_first] + [col_w_rest] * (n_cols - 1)
    else:
        col_w_first = int(content_w * 0.18)
        col_w_second = int(content_w * 0.08)
        remaining = content_w - col_w_first - col_w_second
        col_w_rest = int(remaining / max(n_cols - 2, 1))
        col_widths = [col_w_first, col_w_second] + [col_w_rest] * (n_cols - 2)

    # וודא שהסכום שווה לתוכן
    total = sum(col_widths)
    if total != content_w:
        col_widths[-1] += (content_w - total)

    # כמה שורות נתונים מתאימות לדף (הערכה)
    PAGE_H_TWIPS = PAGE_SIZES[page_size][1] if not landscape else PAGE_SIZES[page_size][0]
    available_h = PAGE_H_TWIPS - 2 * MARGIN_NARROW
    HEADER_ROWS_H = 600   # גובה כותרת טבלה בטוויפס
    TITLE_H = 400
    LEGEND_H = 300
    DATA_ROW_H = 280      # גובה שורת נתונים
    n_header_rows = len(headers)

    rows_per_page = max(5, int(
        (available_h - TITLE_H - LEGEND_H - n_header_rows * HEADER_ROWS_H) / DATA_ROW_H
    ))

    # פצל לחלקים
    chunks = []
    for i in range(0, len(data_rows), rows_per_page):
        chunks.append(data_rows[i:i+rows_per_page])

    total_parts = len(chunks)

    for part_idx, chunk in enumerate(chunks):
        part_num = part_idx + 1
        part_str = f"(חלק {part_num} מתוך {total_parts})" if total_parts > 1 else ""

        # section break (לא לחלק הראשון בראשון)
        if not (is_first_section and part_idx == 0):
            _add_section_break(doc, page_size, landscape)

        # כותרת
        _add_title_paragraph(doc, title, part_str, rtl=True)

        # בנה טבלה
        n_rows_total = n_header_rows + len(chunk)
        table = doc.add_table(rows=n_rows_total, cols=n_cols)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        _set_table_width(table, content_w)

        # כתוב header rows
        for hi, hrow in enumerate(headers):
            for ci, val in enumerate(hrow):
                cell = table.cell(hi, ci)
                _set_col_width(cell, col_widths[ci])
                is_first_hdr_row = (hi == 0)
                bg = COLOR_HEADER2 if is_first_hdr_row else COLOR_HEADER
                _cell_text(cell, val, bold=True, size=8, bg=bg, rtl=True)

        # כתוב data rows
        has_yellow = False
        has_orange = False
        for ri, row_data in enumerate(chunk):
            tr_idx = n_header_rows + ri
            for ci, (val, color) in enumerate(zip(row_data["values"], row_data["colors"])):
                cell = table.cell(tr_idx, ci)
                _set_col_width(cell, col_widths[ci])
                bg = None
                if color == "vsl":
                    bg = COLOR_YELLOW
                    has_yellow = True
                elif color == "tier1":
                    bg = COLOR_ORANGE
                    has_orange = True
                left_align = (ci == 0)
                _cell_text(cell, val, bold=bool(color), size=8, bg=bg or COLOR_WHITE,
                           align="right" if left_align else "center", rtl=True)

        # מקרא
        _add_legend(doc, has_yellow, has_orange)


# ═══════════════════════════════════════════════════════════════════════════════
#  פונקציה ציבורית: בנה קובץ Word
# ═══════════════════════════════════════════════════════════════════════════════

def build_word_report(table_configs, thresh_dict, t1col, t1lbl):
    """
    table_configs: list of dicts, each:
      {
        "type": "TPH" | "Metals" | "VOC+SVOC" | "PFAS",
        "df": DataFrame,
        "title": str,
        "page_size": "A4" | "Tabloid",
        "landscape": bool,
      }
    מחזיר: bytes של קובץ docx
    """
    doc = Document()

    # הסר section ראשוני ריק
    for section in doc.sections:
        section.page_width  = Twips(PAGE_SIZES["A4"][0])
        section.page_height = Twips(PAGE_SIZES["A4"][1])
        section.top_margin    = Twips(MARGIN_NARROW)
        section.bottom_margin = Twips(MARGIN_NARROW)
        section.left_margin   = Twips(MARGIN_NARROW)
        section.right_margin  = Twips(MARGIN_NARROW)

    # הסר פסקה ריקה ראשונית
    for p in doc.paragraphs:
        p._element.getparent().remove(p._element)

    is_first = True

    for cfg in table_configs:
        df        = cfg["df"]
        ttype     = cfg["type"]
        title     = cfg["title"]
        page_size = cfg.get("page_size", "A4")
        landscape = cfg.get("landscape", False)

        if df is None or df.empty:
            continue

        # בנה נתוני טבלה
        if ttype == "TPH":
            headers, data_rows = _build_tph_table_data(df, thresh_dict, t1col, t1lbl)
        elif ttype == "Metals":
            headers, data_rows = _build_metals_table_data(df, thresh_dict, t1col, t1lbl)
        elif ttype in ("VOC+SVOC", "PFAS"):
            headers, data_rows = _build_generic_table_data(df, thresh_dict, t1col, t1lbl)
        else:
            headers, data_rows = _build_generic_table_data(df, thresh_dict, t1col, t1lbl)

        if not data_rows:
            continue

        # הגדר section ראשון
        if is_first:
            section = doc.sections[0]
            _set_section_props(section, page_size, landscape)
            is_first = False

        _add_table_to_doc(doc, headers, data_rows, title, page_size, landscape,
                          is_first_section=(len(doc.element.body) <= 2))

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.getvalue()



with tab_word:
    # ── WORD EXPORT SECTION ──────────────────────────────────────────────────────
    st.markdown("---")
    st.header("📄 יצוא דוח Word")
    st.caption("בחר עד 4 טבלאות, סדר אותן, הגדר סוג דף וכיוון – ויוצר קובץ Word מסודר")

    # ── סשן-סטייט לניהול סדר הטבלאות ──────────────────────────────────────────
    if "word_table_order" not in st.session_state:
        st.session_state.word_table_order = ["TPH", "Metals", "VOC+SVOC", "PFAS"]

    TABLE_TYPES = ["TPH", "Metals", "VOC+SVOC", "PFAS"]
    TABLE_LABELS = {
        "TPH":      "🛢️ TPH (פחמימנים)",
        "Metals":   "⚗️ Metals (מתכות)",
        "VOC+SVOC": "🧪 VOC+SVOC",
        "PFAS":     "🔬 PFAS",
    }
    TABLE_DF_MAP = {
        "TPH":      tph_df,
        "Metals":   metals_df,
        "VOC+SVOC": voc_df,
        "PFAS":     pfas_df,
    }

    # ── בחירת סדר טבלאות ────────────────────────────────────────────────────────
    st.subheader("1️⃣ סדר הטבלאות בדוח")
    st.caption("גרור או שנה את הסדר באמצעות הבחירה למטה")

    order_cols = st.columns(4)
    new_order = []
    used = []
    for i, col in enumerate(order_cols):
        with col:
            default_idx = i if i < len(st.session_state.word_table_order) else 0
            default_val = st.session_state.word_table_order[i] if i < len(st.session_state.word_table_order) else TABLE_TYPES[i]
            available = [t for t in TABLE_TYPES if t not in used]
            if default_val not in available:
                default_val = available[0] if available else TABLE_TYPES[i]
            choice = st.selectbox(
                f"מיקום {i+1}",
                TABLE_TYPES,
                index=TABLE_TYPES.index(default_val),
                key=f"word_order_{i}"
            )
            new_order.append(choice)
            used.append(choice)

    st.session_state.word_table_order = new_order

    # ── הגדרות לכל טבלה ─────────────────────────────────────────────────────────
    st.subheader("2️⃣ הגדרות לכל טבלה")

    word_configs = {}
    for ttype in TABLE_TYPES:
        df_avail = TABLE_DF_MAP[ttype]
        has_data = df_avail is not None and not df_avail.empty

        with st.expander(f"{TABLE_LABELS[ttype]}  {'✅ נטען' if has_data else '⚠️ אין נתונים'}", expanded=has_data):
            col_a, col_b, col_c, col_d = st.columns([2, 1, 1, 1])

            with col_a:
                include = st.checkbox("כלול בדוח", value=has_data, key=f"word_include_{ttype}", disabled=not has_data)
                custom_title = st.text_input(
                    "שם הטבלה (ריק = ברירת מחדל)",
                    value="",
                    key=f"word_title_{ttype}",
                    placeholder=f"למשל: תוצאות {ttype} - אתר X"
                )

            with col_b:
                page_size = st.selectbox(
                    "סוג דף",
                    ["A4", "Tabloid"],
                    index=0,
                    key=f"word_page_{ttype}"
                )

            with col_c:
                landscape = st.selectbox(
                    "כיוון דף",
                    ["לאורך (Portrait)", "לרוחב (Landscape)"],
                    index=1 if ttype in ("Metals", "VOC+SVOC", "PFAS") else 0,
                    key=f"word_orient_{ttype}"
                ) == "לרוחב (Landscape)"

            with col_d:
                if has_data:
                    n_samples = df_avail["sample_id"].nunique() if "sample_id" in df_avail.columns else 0
                    n_compounds = df_avail["compound"].nunique() if "compound" in df_avail.columns else 0
                    st.metric("קידוחים", n_samples)
                    if n_compounds > 0:
                        st.metric("תרכובות", n_compounds)

            word_configs[ttype] = {
                "include":    include if has_data else False,
                "title":      custom_title if custom_title.strip() else f"טבלת {ttype}",
                "page_size":  page_size,
                "landscape":  landscape,
                "df":         df_avail if has_data else None,
                "type":       ttype,
            }

    # ── כפתור יצוא ───────────────────────────────────────────────────────────────
    st.subheader("3️⃣ יצוא")

    col_export, col_info = st.columns([1, 2])
    with col_export:
        export_word = st.button("📄 צור קובץ Word", type="primary", use_container_width=True)

    with col_info:
        tables_to_export = [t for t in st.session_state.word_table_order if word_configs[t]["include"]]
        if tables_to_export:
            st.info(f"יכיל {len(tables_to_export)} טבלאות לפי הסדר: {' → '.join(tables_to_export)}")
        else:
            st.warning("לא נבחרו טבלאות לייצוא")

    if export_word:
        if not tables_to_export:
            st.error("❌ לא נבחרו טבלאות לייצוא")
        else:
            try:
                ordered_configs = []
                for ttype in st.session_state.word_table_order:
                    cfg = word_configs[ttype]
                    if cfg["include"]:
                        ordered_configs.append({
                            "type":      cfg["type"],
                            "df":        cfg["df"],
                            "title":     cfg["title"],
                            "page_size": cfg["page_size"],
                            "landscape": cfg["landscape"],
                        })

                with st.spinner("⏳ מכין קובץ Word..."):
                    docx_bytes = build_word_report(ordered_configs, thresh_dict, t1col, t1lbl)

                st.success(f"✅ הדוח נוצר בהצלחה! ({len(tables_to_export)} טבלאות)")
                st.download_button(
                    "⬇️ הורד דוח Word",
                    data=docx_bytes,
                    file_name="soil_report.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )

            except Exception as e:
                st.error(f"❌ שגיאה ביצירת הדוח: {e}")
                import traceback
                st.code(traceback.format_exc())

