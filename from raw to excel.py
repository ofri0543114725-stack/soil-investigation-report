import streamlit as st
import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import io, re

st.set_page_config(page_title="דוח סקר קרקע", layout="wide", page_icon="🧪")
st.title("🧪 מערכת עיבוד תוצאות מעבדה")
st.markdown("---")

# ── STYLES ────────────────────────────────────────────────────────────────────
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

# ── HELPERS ───────────────────────────────────────────────────────────────────
def norm(s):
    s = "" if s is None else str(s).strip().lower()
    return re.sub(r"\s+", " ", s.replace("\xa0", " "))

def to_float(v):
    s = str(v).strip() if v is not None else ""
    try: return float(s.lstrip("<>").strip())
    except: return None

def sort_key(sid):
    m = re.match(r"S-?(\d+)", str(sid), re.I)
    return int(m.group(1)) if m else 9999

def parse_sample(sname):
    sname = str(sname).strip()
    if "DUP" in sname.upper(): return None, None
    m = re.match(r"^(S\d+[A-Za-z0-9]*)\s*\(([0-9.]+)\)", sname)
    if m: return m.group(1), float(m.group(2))
    m = re.match(r"^(S\d+)-([0-9]+\.?[0-9]*)$", sname)
    if m: return m.group(1), float(m.group(2))
    return sname, None

def check_exceed(val_str, vsl, tier1):
    """VSL=yellow, TIER1=orange. Only colors actual detections (not <LOR)."""
    if not val_str or str(val_str).strip().startswith("<"): return None
    f = to_float(val_str)
    if f is None: return None
    try:
        t1f = float(tier1) if tier1 is not None and str(tier1) not in ("-","NA","") and pd.notna(tier1) else None
        vf  = float(vsl)   if vsl   is not None and str(vsl)   not in ("-","NA","") and pd.notna(vsl)   else None
        if t1f and t1f > 0 and f > t1f: return "tier1"
        if vf  and vf  > 0 and f > vf:  return "vsl"
    except: pass
    return None

def apply_sid_merge(ws, sid_rows, col=1):
    for sid, rows_list in sid_rows.items():
        if len(rows_list) > 1:
            ws.merge_cells(start_row=rows_list[0], start_column=col,
                           end_row=rows_list[-1], end_column=col)
            c = ws.cell(rows_list[0], col)
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            c.border = thin_border()

# ── METAL MAP ─────────────────────────────────────────────────────────────────
# Maps ALS compound names (lowercased) -> symbol
METAL_MAP = {
    # standard names from ALS
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

# Maps threshold file names (lowercased, stripped) -> symbol
# These are the exact names in the user's threshold file
THRESH_METAL_MAP = {
    "aluminum":                     "Al",
    "antimony (metallic)":          "Sb",
    "antimony":                     "Sb",
    "arsenic, inorganic":           "As",
    "arsenic":                      "As",
    "barium":                       "Ba",
    "beryllium and compounds":      "Be",
    "beryllium":                    "Be",
    "boron and borates only":       "B",
    "boron":                        "B",
    "cadmium (water) source: water and air": "Cd",
    "cadmium":                      "Cd",
    "calcium":                      "Ca",
    "chromium, total":              "Cr",
    "chromium":                     "Cr",
    "cobalt":                       "Co",
    "copper":                       "Cu",
    "iron":                         "Fe",
    "lead and compounds":           "Pb",
    "lead":                         "Pb",
    "lithium":                      "Li",
    "magnesium":                    "Mg",
    "manganese (non-diet)":         "Mn",
    "manganese":                    "Mn",
    "mercuric chloride (and other mercury salts)": "Hg",
    "mercury":                      "Hg",
    "nickel soluble salts":         "Ni",
    "nickel":                       "Ni",
    "potassium":                    "K",
    "selenium":                     "Se",
    "silver":                       "Ag",
    "sodium":                       "Na",
    "vanadium and compounds":       "V",
    "vanadium":                     "V",
    "zinc and compounds":           "Zn",
    "zinc":                         "Zn",
    "molybdenum":                   "Mo",
    "tin":                          "Sn",
    "titanium":                     "Ti",
    "strontium":                    "Sr",
    "thallium":                     "Tl",
    "phosphorus":                   "P",
    "sulphur":                      "S",
    "silicon":                      "Si",
}

# ── PFAS ALIAS ────────────────────────────────────────────────────────────────
PFAS_ALIAS = {
    "2,3,3,3-tetrafluoro-2-(heptafluoropropoxy)propanoic acid (hfpo-da)": "hexafluoropropylene oxide dimer acid (hfpo-da)",
    "7h-perfluoroheptanoic acid (hpfhpa)":         "perfluoroheptanoic acid (pfhpa)",
    "perfluorobutane sulfonic acid (pfbs)":         "perfluorobutanesulfonic acid (pfbs)",
    "perfluorobutane sulfonate (pfbs)":             "perfluorobutanesulfonic acid (pfbs)",
    "perfluorohexane sulfonic acid (pfhxs)":        "perfluorohexanesulfonic acid (pfhxs)",
    "perfluorohexane sulfonate (pfhxs)":            "perfluorohexanesulfonic acid (pfhxs)",
    "perfluorooctane sulfonic acid (pfos)":         "perfluorooctanesulfonic acid (pfos)",
    "perfluorooctane sulfonate (pfos)":             "perfluorooctanesulfonic acid (pfos)",
    "perfluorooctadecanoic acid (pfocda)":          "perfluorooctadecanoic acid (pfoda)",
    "perfluoroundecanoic acid (pfunda)":            "perfluoroundecanoic acid (pfuda)",
    "perfluorotetradecanoic acid (pfcpda)":         "perfluorotetradecanoic acid (pfteta)",
    "perfluorodecane sulfonic acid (pfds)":         "perfluorodecanesulfonic acid (pfds)",
    "perfluoroheptane sulfonic acid (pfhps)":       "perfluoroheptanesulfonic acid (pfhps)",
    "perfluoropentane sulfonic acid (pfpes)":       "perfluoropentanesulfonic acid (pfpes)",
    "perfluorooctane sulfonamide (fosa)":           "perfluorooctanesulfonamide (fosa)",
    "perfluoropentanoic acid (pfpea)":              "perfluoropentanoic acid (pfpea)",
    "perfluorodecanoic acid (pfda)":                "perfluorodecanoic acid (pfda)",
    "perfluorododecanoic acid (pfdoda)":            "perfluorododecanoic acid (pfdoda)",
    "perfluoroheptanoic acid (pfhpa)":              "perfluoroheptanoic acid (pfhpa)",
    "perfluorotridecanoic acid (pftrda)":           "perfluorotridecanoic acid (pftrda)",
    "perfluorooctanesulfonic acid (pfos)":          "perfluorooctanesulfonic acid (pfos)",
}

def match_threshold(compound_name, thresh_dict):
    key = norm(compound_name)
    if key in thresh_dict: return thresh_dict[key]
    aliased = PFAS_ALIAS.get(key)
    if aliased and aliased in thresh_dict: return thresh_dict[aliased]
    stripped = re.sub(r"\s*\([A-Z0-9:_\-]+\)\s*$", "", compound_name).strip().lower()
    for k, v in thresh_dict.items():
        k_s = re.sub(r"\s*\([A-Z0-9:_\-]+\)\s*$", "", k).strip()
        if len(stripped) > 8 and stripped == k_s: return v
    return {}

# ── THRESHOLD FILE ────────────────────────────────────────────────────────────
def load_threshold_file(file_bytes):
    wb = load_workbook(io.BytesIO(file_bytes), data_only=True)
    ws = wb.active
    thresh = {}
    for row in ws.iter_rows(min_row=6, values_only=True):
        if not row[0]: continue
        name = str(row[0]).strip()
        cas  = str(row[1]).strip() if row[1] else "-"
        def g(ci): return row[ci] if ci < len(row) and row[ci] is not None and str(row[ci]) not in ("NA","") else None
        thresh[norm(name)] = {
            "name":name,"cas":cas,"units":str(row[3]) if row[3] else "mg/kg",
            "VSL":g(4),"Ind_A_06":g(8),"Ind_A_6p":g(9),"Ind_B":g(10),
            "Res_A_06":g(11),"Res_A_6p":g(12),"Res_B":g(13),
        }
    return thresh

def get_tier1_col(land_use, aquifer, depth):
    ind = "industrial" in land_use.lower()
    b   = "b-1" in aquifer.lower()
    if b: return "Ind_B" if ind else "Res_B"
    deep = ">6" in depth
    if ind: return "Ind_A_06" if not deep else "Ind_A_6p"
    else:   return "Res_A_06" if not deep else "Res_A_6p"

def tier1_label(land_use, aquifer, depth):
    return f"TIER 1\n{land_use}\n{aquifer}\n{depth}"

def get_thresh(compound, thresh_dict, t1col):
    t = match_threshold(compound, thresh_dict)
    return t.get("VSL"), t.get(t1col), t.get("cas","-")

def build_metals_thresh(thresh_dict, t1col):
    """Build {symbol: {vsl,tier1,cas}} using THRESH_METAL_MAP to match threshold names."""
    result = {}
    for key, v in thresh_dict.items():
        sym = THRESH_METAL_MAP.get(key)
        if sym and sym not in result:   # first match wins
            result[sym] = {"vsl":v.get("VSL"),"tier1":v.get(t1col),"cas":v.get("cas","-")}
    return result

# ── ALS PARSER ────────────────────────────────────────────────────────────────
def parse_als_file(file_bytes, filename):
    try: wb = load_workbook(io.BytesIO(file_bytes), data_only=True)
    except Exception as e: return None, str(e)
    main = next((wb[n] for n in wb.sheetnames if "Client" in n and "SOIL" in n), wb.worksheets[0])
    rows = list(main.iter_rows(values_only=True))
    sid_idx = next((i for i,r in enumerate(rows) if any("Client Sample ID" in str(v) for v in r if v)), None)
    if sid_idx is None: return None, "לא נמצאה שורת Sample IDs"
    col2sample = {ci:str(v).strip() for ci,v in enumerate(rows[sid_idx]) if v and v!="Client Sample ID"}
    ph_idx = next((i for i,r in enumerate(rows) if r and r[0]=="Parameter"), None)
    if ph_idx is None: return None, "לא נמצאה שורת Parameter"

    records = []; group = "Unknown"
    for row in rows[ph_idx+1:]:
        p=row[0] if len(row)>0 else None; m=row[1] if len(row)>1 else None
        u=row[2] if len(row)>2 else None; lor=row[3] if len(row)>3 else None
        if not p: continue
        if not m and not u: group=str(p).strip(); continue
        for ci,sname in col2sample.items():
            sid,depth = parse_sample(sname)
            if sid is None: continue
            val=row[ci] if ci<len(row) else None
            rs=str(val).strip() if val is not None else ""
            result=None
            if rs.startswith("<"): result=0.0
            elif rs and rs not in ("None",""):
                try: result=float(rs)
                except: result=None
            if result is not None:
                lor_val = None
                if rs.startswith("<"):
                    try: lor_val = float(rs[1:].strip())
                    except: lor_val = 0.0
                records.append({"sample_id":sid,"depth":depth,"compound":str(p).strip(),
                    "compound_lower":norm(p),"unit":str(u).strip() if u else "mg/kg",
                    "lor":lor,"result":result,"result_str":rs,"lor_val":lor_val,"group":group,"source":filename})
    if not records: return None, "לא נמצאו נתונים"
    return pd.DataFrame(records), None

# ── TPH SHEET ─────────────────────────────────────────────────────────────────
def write_tph_sheet(ws, df, thresh_dict, t1col, t1lbl):
    def is_dro(c):
        # explicit DRO in name, OR c10-c28 (without c40, without oro)
        if "dro" in c: return True
        if "oro" in c: return False
        if "c10" in c and "c28" in c and "c40" not in c: return True
        return False
    def is_oro(c):
        # explicit ORO in name, OR c24-c40 or c28-c40 (without dro)
        if "oro" in c: return True
        if "dro" in c: return False
        if "c24" in c and "c40" in c: return True
        if "c28" in c and "c40" in c: return True
        return False
    def is_total(c):
        # c10-c40 total (no dro/oro) — use as Total TPH directly
        if "dro" in c or "oro" in c: return False
        if "c10" in c and "c40" in c: return True
        if "total" in c and ("tph" in c or "hydrocarbon" in c): return True
        return False

    vsl_d,t1_d,_ = get_thresh("C10 - C28 Fraction (DRO)", thresh_dict, t1col)
    vsl_o,t1_o,_ = get_thresh("C24 - C40 Fraction (ORO)", thresh_dict, t1col)
    # use TPH total entry if available
    vsl_t,t1_t,_ = get_thresh("TPH - DRO + ORO (Tier 1)", thresh_dict, t1col)
    vv=[v for v in [vsl_d,vsl_o,vsl_t] if v]; tt=[v for v in [t1_d,t1_o,t1_t] if v]
    vsl_tot=min(vv) if vv else 350; t1_tot=min(tt) if tt else 350

    # Header row 1 (no col C)
    for ci,h in enumerate(["שם קידוח","עומק","TPH DRO","TPH ORO","Total TPH"],1):
        style_hdr(ws.cell(1,ci,h), HDR_BLUE_FILL)
    ws.merge_cells(start_row=1,start_column=1,end_row=5,end_column=1)
    ws.cell(1,1).value="שם קידוח"
    ws.cell(1,1).font=Font(bold=True,name="Arial",size=11)
    ws.cell(1,1).fill=HDR_BLUE_FILL
    ws.cell(1,1).alignment=Alignment(horizontal="center",vertical="center",wrap_text=True)
    ws.cell(1,1).border=thin_border()

    sub_rows=["יחידות","CAS","VSL",t1lbl]
    sub_vals={"יחידות":"mg/kg","CAS":"C10-C40","VSL":vsl_tot,t1lbl:t1_tot}
    for ri,lbl in enumerate(sub_rows,2):
        style_hdr(ws.cell(ri,2,lbl), HDR_BLUE_FILL)
        for ci in [3,4,5]: style_hdr(ws.cell(ri,ci,sub_vals[lbl]), HDR_BLUE_FILL)

    # pivot — one row per (sample_id, depth)
    pivoted = {}
    for _,r in df.iterrows():
        k=(r["sample_id"],r["depth"])
        if k not in pivoted: pivoted[k]={"DRO":"","ORO":"","TOT":"","DRO_f":None,"ORO_f":None,"DRO_lor":None,"ORO_lor":None}
        c=r["compound_lower"]
        if is_dro(c) and not pivoted[k]["DRO"]:
            pivoted[k]["DRO"]=r["result_str"]; pivoted[k]["DRO_f"]=r["result"]
            pivoted[k]["DRO_lor"]=r.get("lor_val")
        elif is_oro(c) and not pivoted[k]["ORO"]:
            pivoted[k]["ORO"]=r["result_str"]; pivoted[k]["ORO_f"]=r["result"]
            pivoted[k]["ORO_lor"]=r.get("lor_val")
        elif is_total(c) and not pivoted[k]["TOT"]:
            pivoted[k]["TOT"]=r["result_str"]

    ri=6; prev_sid=None; sid_rows={}
    for (sid,depth),v in sorted(pivoted.items(),key=lambda x:(sort_key(x[0][0]),x[0][1] or 0)):
        dro_f=v["DRO_f"] or 0; oro_f=v["ORO_f"] or 0

        # if we have an explicit total, use it; else compute from DRO+ORO
        if v["TOT"]:
            total_s = v["TOT"]
        else:
            dro_lor = v["DRO"] and str(v["DRO"]).startswith("<")
            oro_lor = v["ORO"] and str(v["ORO"]).startswith("<")
            dro_empty = not v["DRO"]
            oro_empty = not v["ORO"]
            # use actual LOR number (not 0.0) for correct total
            dro_num = v["DRO_lor"] if dro_lor and v["DRO_lor"] is not None else (v["DRO_f"] or 0)
            oro_num = v["ORO_lor"] if oro_lor and v["ORO_lor"] is not None else (v["ORO_f"] or 0)
            total_f = dro_num + oro_num
            if (dro_lor or dro_empty) and (oro_lor or oro_empty) and not (dro_empty and oro_empty):
                total_s = f"<{total_f:.0f}"
            else:
                total_s = f"{total_f:.0f}"

        hl_dro   = check_exceed(v["DRO"],  vsl_tot, t1_tot)
        hl_oro   = check_exceed(v["ORO"],  vsl_tot, t1_tot)
        hl_total = check_exceed(total_s,   vsl_tot, t1_tot)

        if sid!=prev_sid: sid_rows[sid]=[]
        sid_rows[sid].append(ri)
        sid_val=sid if sid!=prev_sid else None; prev_sid=sid

        style_data(ws.cell(ri,1,sid_val))
        style_data(ws.cell(ri,2,depth))
        style_data(ws.cell(ri,3,v["DRO"]),  hl_dro)
        style_data(ws.cell(ri,4,v["ORO"]),  hl_oro)
        style_data(ws.cell(ri,5,total_s),   hl_total)
        ri+=1

    apply_sid_merge(ws, sid_rows, col=1)
    for col,w in zip("ABCDE",[14,10,16,16,16]):
        ws.column_dimensions[col].width=w
    ws.row_dimensions[1].height=15
    ws.freeze_panes="A6"

# ── METALS SHEET ──────────────────────────────────────────────────────────────
def write_metals_sheet(ws, df, thresh_dict, t1col, t1lbl):
    df=df.copy()
    df["sym"]=df["compound_lower"].map(METAL_MAP)
    df=df[df["sym"].notna()]
    if df.empty: ws.cell(1,1,"אין נתוני מתכות"); return

    present=set(df["sym"].unique())
    metals=[m for m in METALS_ORDER if m in present]+sorted(present-set(METALS_ORDER))
    mt=build_metals_thresh(thresh_dict,t1col)

    ws.merge_cells(start_row=1,start_column=1,end_row=5,end_column=1)
    c=ws.cell(1,1,"שם קידוח")
    c.font=Font(bold=True,name="Arial",size=11); c.fill=HDR_BLUE_FILL
    c.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True); c.border=thin_border()

    for ci,h in enumerate(["עומק"]+metals,2):
        style_hdr(ws.cell(1,ci,h),HDR_BLUE_FILL)

    for ri,lbl in enumerate(["יחידות","CAS","VSL",t1lbl],2):
        style_hdr(ws.cell(ri,2,lbl),HDR_BLUE_FILL)
        for ci,sym in enumerate(metals,3):
            t=mt.get(sym,{})
            val={"יחידות":"mg/kg","CAS":t.get("cas","-"),"VSL":t.get("vsl","-"),t1lbl:t.get("tier1","-")}.get(lbl,"-")
            style_hdr(ws.cell(ri,ci,val),HDR_BLUE_FILL)

    pt=df.pivot_table(index=["sample_id","depth"],columns="sym",values="result_str",aggfunc="first")
    pt=pt.reindex(sorted(pt.index,key=lambda x:(sort_key(x[0]),x[1] or 0)))

    ri=6; prev_sid=None; sid_rows={}
    for (sid,depth),row_data in pt.iterrows():
        if sid!=prev_sid: sid_rows[sid]=[]
        sid_rows[sid].append(ri)
        sid_val=sid if sid!=prev_sid else None; prev_sid=sid
        style_data(ws.cell(ri,1,sid_val))
        style_data(ws.cell(ri,2,depth))
        for ci,sym in enumerate(metals,3):
            val=row_data.get(sym,"") or ""; val="" if str(val)=="nan" else str(val)
            hl=check_exceed(val,mt.get(sym,{}).get("vsl"),mt.get(sym,{}).get("tier1"))
            style_data(ws.cell(ri,ci,val),hl)
        ri+=1

    apply_sid_merge(ws,sid_rows,col=1)
    ws.column_dimensions["A"].width=14; ws.column_dimensions["B"].width=10
    for ci in range(3,len(metals)+3): ws.column_dimensions[get_column_letter(ci)].width=11
    ws.freeze_panes="C6"

# ── PFAS SHEET ────────────────────────────────────────────────────────────────
def write_pfas_sheet(ws, df, thresh_dict, t1col, t1lbl):
    df=df.copy()
    samples=sorted(df["sample_id"].unique(),key=sort_key)
    sdepth={}
    for _,r in df.iterrows():
        if r["sample_id"] not in sdepth: sdepth[r["sample_id"]]=r["depth"]

    fixed=["שם התרכובת","CAS","VSL [µg/kg]",f"{t1lbl}\n[µg/kg]","יחידות","LOR","שם הקידוח"]
    all_cols=fixed+samples
    for ci,h in enumerate(all_cols,1): style_hdr(ws.cell(1,ci,h),HDR_BLUE_FILL)
    style_hdr(ws.cell(2,7,"עומק"),HDR_BLUE_FILL)
    for ci,sid in enumerate(samples,8): style_hdr(ws.cell(2,ci,sdepth.get(sid,"")),HDR_BLUE_FILL)

    for row_i,cmp in enumerate(df["compound"].unique(),3):
        df_c=df[df["compound"]==cmp]
        vsl_mg,tier1_mg,cas=get_thresh(cmp,thresh_dict,t1col)
        # convert mg/kg -> µg/kg (×1000)
        def to_ug(v):
            if v is None: return None
            try: return round(float(v)*1000, 6)
            except: return v
        vsl   = to_ug(vsl_mg)
        tier1 = to_ug(tier1_mg)
        unit = df_c.iloc[0]["unit"] if not df_c.empty else "µg/kg"
        lor  = df_c.iloc[0]["lor"]  if not df_c.empty else ""
        for ci,val in enumerate([cmp,cas,vsl,tier1,unit,lor,None],1): style_data(ws.cell(row_i,ci,val))
        for ci,sid in enumerate(samples,8):
            rs=df_c[df_c["sample_id"]==sid].iloc[0]["result_str"] if not df_c[df_c["sample_id"]==sid].empty else ""
            style_data(ws.cell(row_i,ci,rs),check_exceed(rs,vsl,tier1))

    ws.column_dimensions["A"].width=50
    for ci in range(2,len(all_cols)+1): ws.column_dimensions[get_column_letter(ci)].width=13
    ws.freeze_panes="H3"

# ── VOC SHEET ─────────────────────────────────────────────────────────────────
def write_voc_sheet(ws, df, thresh_dict, t1col, t1lbl):
    df=df.copy()
    samples=sorted(df["sample_id"].unique(),key=sort_key)
    sdepth={}
    for _,r in df.iterrows():
        if r["sample_id"] not in sdepth: sdepth[r["sample_id"]]=r["depth"]

    fixed=["קבוצה","קבוצה","שם התרכובת","CAS","VSL",t1lbl,"יחידות","שם קידוח"]
    all_cols=fixed+samples
    for ci,h in enumerate(all_cols,1): style_hdr(ws.cell(1,ci,h),sz=9)
    style_hdr(ws.cell(2,8,"עומק"),sz=9)
    for ci,sid in enumerate(samples,9): style_hdr(ws.cell(2,ci,sdepth.get(sid,"")),sz=9)

    seen={}
    for _,r in df.iterrows():
        if r["compound"] not in seen: seen[r["compound"]]=r["group"]
    for row_i,(cmp,grp) in enumerate(seen.items(),3):
        df_c=df[df["compound"]==cmp]
        vsl,tier1,cas=get_thresh(cmp,thresh_dict,t1col)
        unit=df_c.iloc[0]["unit"] if not df_c.empty else "mg/kg"
        for ci,val in enumerate([None,grp,cmp,cas,vsl,tier1,unit,None],1): style_data(ws.cell(row_i,ci,val),sz=9)
        for ci,sid in enumerate(samples,9):
            rs=df_c[df_c["sample_id"]==sid].iloc[0]["result_str"] if not df_c[df_c["sample_id"]==sid].empty else ""
            style_data(ws.cell(row_i,ci,rs),check_exceed(rs,vsl,tier1),sz=9)

    ws.column_dimensions["A"].width=8; ws.column_dimensions["B"].width=18
    ws.column_dimensions["C"].width=35; ws.column_dimensions["D"].width=12
    for ci in range(5,len(all_cols)+1): ws.column_dimensions[get_column_letter(ci)].width=11
    ws.freeze_panes="I3"

# ── SIDEBAR ────────────────────────────────────────────────────────────────────
st.sidebar.header("⚙️ הגדרות ערכי סף")
st.sidebar.markdown("🟡 חריגה מ-VSL &nbsp;&nbsp;&nbsp; 🟠 חריגה מ-TIER 1")
st.sidebar.markdown("---")
land_use=st.sidebar.selectbox("Land Use",["Industrial","Residential"],index=0)
aquifer=st.sidebar.selectbox("Aquifer Sensitivity",["A-1, A, B","B-1 or C"],index=0)
depth_opts=["Not Applicable"] if "b-1" in aquifer.lower() else ["0 - 6 m",">6 m"]
depth=st.sidebar.selectbox("Depth to Groundwater",depth_opts,index=0)
t1col=get_tier1_col(land_use,aquifer,depth)
t1lbl=tier1_label(land_use,aquifer,depth)
st.sidebar.info(f"TIER 1: **{land_use}** | {aquifer} | {depth}")

# ── UPLOADS ────────────────────────────────────────────────────────────────────
c1,c2=st.columns(2)
with c1:
    st.subheader("📁 קובץ ערכי סף")
    thr_file=st.file_uploader("העלה קובץ ערכי הסף המאוחד",type=["xlsx","xls"],key="thr")
with c2:
    st.subheader("📂 קבצי ALS")
    data_files=st.file_uploader("העלה קבצי ALS",type=["xlsx","xls"],accept_multiple_files=True,key="data")

if not thr_file: st.info("👆 העלה קובץ ערכי סף וקבצי ALS"); st.stop()
if not data_files: st.warning("⚠️ העלה קבצי ALS"); st.stop()

# ── LOAD ───────────────────────────────────────────────────────────────────────
thresh_dict=load_threshold_file(thr_file.read())
st.success(f"✅ {len(thresh_dict)} ערכי סף | {land_use} | {aquifer} | {depth}")

all_data=[]
for f in data_files:
    df,err=parse_als_file(f.read(),f.name)
    if err: st.warning(f"⚠️ {f.name}: {err}")
    else: all_data.append(df); st.success(f"✅ {f.name} — {len(df)} תוצאות")

if not all_data: st.error("לא נטענו נתונים."); st.stop()
df_all=pd.concat(all_data,ignore_index=True)

with st.expander(f"👁️ תצוגה מקדימה ({len(df_all)} שורות)"): st.dataframe(df_all.head(30))
with st.expander("קבוצות שנמצאו"): st.write(df_all["group"].unique().tolist())

# ── CLASSIFY ───────────────────────────────────────────────────────────────────
def dg(kw): return df_all[df_all["group"].str.contains("|".join(kw),case=False,na=False)]
tph_df    = dg(["petroleum","tph","hydrocarbon"])
metals_df = dg(["metal","cation","extractable"])
pfas_df   = dg(["perfluor","pfas","fluorin"])
voc_df    = dg(["voc","svoc","btex","aromatic","halogenated","volatile",
                 "alcohol","aldehyde","ketone","phenol","pah","aniline",
                 "nitro","phthalate","pesticide","pcb","other"])

# ── BUILD ──────────────────────────────────────────────────────────────────────
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

# ── DOWNLOAD ───────────────────────────────────────────────────────────────────
st.markdown("---")
buf=io.BytesIO(); wb_out.save(buf); buf.seek(0)
st.download_button("⬇️ הורד קובץ Excel מעובד",data=buf,file_name="soil_report.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",use_container_width=True)
