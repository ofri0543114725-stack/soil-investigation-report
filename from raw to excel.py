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

def fmt_number(val):
    """הוסף פסיקים למספרים מעל 999"""
    if val is None or val == "": return val
    s = str(val).strip()
    if s.startswith("<") or s.startswith(">"): return s
    try:
        f = float(s)
        if abs(f) < 1000: return val
        if "." in s:
            decimals = len(s.split(".")[-1])
            return f"{f:,.{decimals}f}"
        return f"{int(f):,}"
    except (ValueError, TypeError):
        return val

def style_data(cell, hl=None, sz=10, left_align=False):
    cell.font      = Font(bold=bool(hl), name="Arial", size=sz)
    cell.alignment = Alignment(horizontal="left" if left_align else "center",
                               vertical="center", wrap_text=True)
    cell.border    = thin_border()
    if   hl == "tier1": cell.fill = ORANGE_FILL
    elif hl == "vsl":   cell.fill = YELLOW_FILL
    if cell.value is not None:
        cell.value = fmt_number(cell.value)

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
    # S85 (0.5) או S1 (1.0)
    m = re.match(r"^(S\d+[A-Za-z0-9]*)\s*\(([0-9.]+)\)", sname)
    if m: return m.group(1), float(m.group(2))
    # S1-1.0
    m = re.match(r"^(S\d+)-([0-9]+\.?[0-9]*)$", sname)
    if m: return m.group(1), float(m.group(2))
    # 24.15 (3.0) — שם קידוח מספרי עם נקודה
    m = re.match(r"^(\d+[\.\d]*)\s*\(([0-9.]+)\)", sname)
    if m: return m.group(1), float(m.group(2))
    # שם קידוח בלי עומק — קבל כמות שהוא
    return sname, None

def check_exceed(val_str, vsl, tier1):
    if not val_str or str(val_str).strip().startswith("<"): return None
    f = to_float(val_str)
    if f is None: return None
    try:
        t1f = float(tier1) if (tier1 is not None and str(tier1) not in ("-","NA","") and pd.notna(tier1)) else None
        vf  = float(vsl)   if (vsl   is not None and str(vsl)   not in ("-","NA","") and pd.notna(vsl))   else None
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

METALS_ORDER = [
    "Al","Sb","As","Ba","Be","Bi","B","Cd","Ca","Cr","Co","Cu","Fe",
    "Pb","Li","Mg","Mn","Hg","Ni","K","Se","Ag","Na","V","Zn"
]

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

VOC_COMPOUND_ORDER = [
    ("VOCs","Non-Halogenated VOCs","1.2.4-Trimethylbenzene"),
    ("VOCs","Non-Halogenated VOCs","1.3.5-Trimethylbenzene"),
    ("VOCs","Non-Halogenated VOCs","MTBE"),
    ("VOCs","Non-Halogenated VOCs","Styrene"),
    ("VOCs","Non-Halogenated VOCs","n-Butylbenzene"),
    ("VOCs","Non-Halogenated VOCs","n-Propylbenzene"),
    ("VOCs","Non-Halogenated VOCs","Isopropylbenzene"),
    ("VOCs","Non-Halogenated VOCs","Acetone"),
    ("VOCs","Non-Halogenated VOCs","2-Butanone (MEK)"),
    ("VOCs","Non-Halogenated VOCs","1.4-Dioxane"),
    ("VOCs","BTEX","Benzene"),
    ("VOCs","BTEX","Toluene"),
    ("VOCs","BTEX","Ethylbenzene"),
    ("VOCs","BTEX","Sum of Xylenes"),
    ("VOCs","Halogenated VOCs","1.1-Dichloroethane"),
    ("VOCs","Halogenated VOCs","1.1-Dichloroethene"),
    ("VOCs","Halogenated VOCs","1.2-Dichloroethane"),
    ("VOCs","Halogenated VOCs","1.2-Dichloropropane"),
    ("VOCs","Halogenated VOCs","Chlorobenzene"),
    ("VOCs","Halogenated VOCs","Chloroform"),
    ("VOCs","Halogenated VOCs","Dichloromethane"),
    ("VOCs","Halogenated VOCs","Tetrachloroethene"),
    ("VOCs","Halogenated VOCs","Tetrachloromethane"),
    ("VOCs","Halogenated VOCs","Trichloroethene"),
    ("VOCs","Halogenated VOCs","Vinyl chloride"),
    ("VOCs","Halogenated VOCs","cis-1.2-Dichloroethene"),
    ("VOCs","Halogenated VOCs","trans-1.2-Dichloroethene"),
    ("VOCs","Halogenated VOCs","1.4-Dichlorobenzene"),
    ("VOCs","Halogenated VOCs","1.2-Dichlorobenzene"),
    ("VOCs","Halogenated VOCs","1.3-Dichlorobenzene"),
    ("SVOCs","Phenols & Naphtols","2.4-Dimethylphenol"),
    ("SVOCs","Phenols & Naphtols","2-Methylphenol"),
    ("SVOCs","Phenols & Naphtols","3 & 4-Methylphenol"),
    ("SVOCs","Phenols & Naphtols","4-Chloro-3-methylphenol"),
    ("SVOCs","Phenols & Naphtols","Phenol"),
    ("SVOCs","PAHs","Acenaphthene"),
    ("SVOCs","PAHs","Acenaphthylene"),
    ("SVOCs","PAHs","Anthracene"),
    ("SVOCs","PAHs","Benz(a)anthracene"),
    ("SVOCs","PAHs","Benzo(a)pyrene"),
    ("SVOCs","PAHs","Benzo(b)fluoranthene"),
    ("SVOCs","PAHs","Benzo(g.h.i)perylene"),
    ("SVOCs","PAHs","Benzo(k)fluoranthene"),
    ("SVOCs","PAHs","Chrysene"),
    ("SVOCs","PAHs","Dibenz(a.h)anthracene"),
    ("SVOCs","PAHs","Fluoranthene"),
    ("SVOCs","PAHs","Fluorene"),
    ("SVOCs","PAHs","Indeno(1.2.3.cd)pyrene"),
    ("SVOCs","PAHs","Naphthalene"),
    ("SVOCs","PAHs","Phenanthrene"),
    ("SVOCs","PAHs","Pyrene"),
    ("SVOCs","Anilines","4-Chloroaniline"),
    ("SVOCs","Anilines","Aniline"),
    ("SVOCs","Anilines","Benzidine"),
    ("SVOCs","Anilines","Diphenylamine"),
    ("SVOCs","Aromatic Compounds","1,1'-Biphenyl"),
    ("SVOCs","Aromatic Compounds","1-Chloronaphthalene"),
    ("SVOCs","Aromatic Compounds","2-Chloronaphthalene"),
    ("SVOCs","Aromatic Compounds","2-Methylnaphthalene"),
    ("SVOCs","Aromatic Compounds","4-Bromophenyl phenyl ether"),
    ("SVOCs","Aromatic Compounds","4-Chlorophenyl phenyl ether"),
    ("SVOCs","Aromatic Compounds","Carbazole"),
    ("SVOCs","Aromatic Compounds","Dibenzofuran"),
    ("SVOCs","Alcohols","Benzyl Alcohol"),
    ("SVOCs","Aldehydes / Ketones","6-Caprolactam"),
    ("SVOCs","Aldehydes / Ketones","Acetophenone"),
    ("SVOCs","Aldehydes / Ketones","Isophorone"),
    ("SVOCs","Chlorophenols","2-Chlorophenol"),
    ("SVOCs","Chlorophenols","2.4.5-Trichlorophenol"),
    ("SVOCs","Chlorophenols","2.4.6-Trichlorophenol"),
    ("SVOCs","Chlorophenols","2.4-Dichlorophenol"),
    ("SVOCs","Chlorophenols","2.6-Dichlorophenol"),
    ("SVOCs","Chlorophenols","Pentachlorophenol"),
    ("SVOCs","Nitroaromatic Compounds","2.4-Dinitrophenol"),
    ("SVOCs","Nitroaromatic Compounds","2.4-Dinitrotoluene"),
    ("SVOCs","Nitroaromatic Compounds","2-Nitroaniline"),
    ("SVOCs","Nitroaromatic Compounds","2-Nitrophenol"),
    ("SVOCs","Nitroaromatic Compounds","2.6-Dinitrotoluene"),
    ("SVOCs","Nitroaromatic Compounds","3-Nitroaniline"),
    ("SVOCs","Nitroaromatic Compounds","4.6-Dinitro-2-methylphenol"),
    ("SVOCs","Nitroaromatic Compounds","4-Nitroaniline"),
    ("SVOCs","Nitroaromatic Compounds","4-Nitrophenol"),
    ("SVOCs","Nitroaromatic Compounds","Nitrobenzene"),
    ("SVOCs","Chlorinated Hydrocarbons","Bis(2-chloroethoxy)methane"),
    ("SVOCs","Chlorinated Hydrocarbons","Bis(2-chloroethyl)ether"),
    ("SVOCs","Chlorinated Hydrocarbons","Bis(2-chloroisopropyl)ether"),
    ("SVOCs","Nitrosoamines","N-Nitrosodi-n-propylamine"),
    ("SVOCs","Pesticides","Dinoseb"),
    ("SVOCs","Phthalates","Bis(2-ethylhexyl)phthalate"),
    ("SVOCs","Phthalates","Butyl benzyl phthalate"),
    ("SVOCs","Phthalates","Di-n-butyl phthalate"),
    ("SVOCs","Phthalates","Di-n-octyl phthalate"),
    ("SVOCs","Phthalates","Diethyl phthalate"),
    ("SVOCs","Phthalates","Dimethyl phthalate"),
]

# ── VOC ALIAS ─────────────────────────────────────────────────────────────────────
VOC_ALIAS = {
    "1.2.4-trimethylbenzene":           "Trimethylbenzene, 1,2,4-",
    "1,2,4-trimethylbenzene":           "Trimethylbenzene, 1,2,4-",
    "1.3.5-trimethylbenzene":           "Trimethylbenzene, 1,3,5-",
    "1,3,5-trimethylbenzene":           "Trimethylbenzene, 1,3,5-",
    "mtbe":                             "Methyl tert-Butyl Ether (MTBE)",
    "methyl tert-butyl ether":          "Methyl tert-Butyl Ether (MTBE)",
    "n-propylbenzene":                  "Propyl benzene",
    "propylbenzene":                    "Propyl benzene",
    "isopropylbenzene":                 "Cumene",
    "cumene":                           "Cumene",
    "2-butanone (mek)":                 "Methyl Ethyl Ketone - MEK (2-Butanone)",
    "2-butanone":                       "Methyl Ethyl Ketone - MEK (2-Butanone)",
    "methyl ethyl ketone":              "Methyl Ethyl Ketone - MEK (2-Butanone)",
    "mek":                              "Methyl Ethyl Ketone - MEK (2-Butanone)",
    "sum of xylenes":                   "Xylenes",
    "xylenes, total":                   "Xylenes",
    "total xylenes":                    "Xylenes",
    "1.1-dichloroethene":               "1,1-Dichloroethylene",
    "1,1-dichloroethene":               "1,1-Dichloroethylene",
    "1,1-dichloroethylene":             "1,1-Dichloroethylene",
    "1.2-dichloroethane":               "1,2- (EDC) Dichloroethane",
    "1,2-dichloroethane":               "1,2- (EDC) Dichloroethane",
    "ethylene dichloride":              "1,2- (EDC) Dichloroethane",
    "dichloromethane":                  "Methylene Chloride",
    "methylene chloride":               "Methylene Chloride",
    "tetrachloroethene":                "Tetrachloroethylene (PCE)",
    "tetrachloroethylene":              "Tetrachloroethylene (PCE)",
    "perchloroethylene":                "Tetrachloroethylene (PCE)",
    "pce":                              "Tetrachloroethylene (PCE)",
    "tetrachloromethane":               "Carbon Tetrachloride",
    "carbon tetrachloride":             "Carbon Tetrachloride",
    "trichloroethene":                  "Trichloroethylene (TCE)",
    "trichloroethylene":                "Trichloroethylene (TCE)",
    "tce":                              "Trichloroethylene (TCE)",
    "cis-1.2-dichloroethene":           "1,2-cis-Dichloroethylene",
    "cis-1,2-dichloroethene":           "1,2-cis-Dichloroethylene",
    "cis-1,2-dichloroethylene":         "1,2-cis-Dichloroethylene",
    "1,2-cis-dichloroethylene":         "1,2-cis-Dichloroethylene",
    "trans-1.2-dichloroethene":         "1,2-trans-Dichloroethylene",
    "trans-1,2-dichloroethene":         "1,2-trans-Dichloroethylene",
    "trans-1,2-dichloroethylene":       "1,2-trans-Dichloroethylene",
    "1,2-trans-dichloroethylene":       "1,2-trans-Dichloroethylene",
    "benz(a)anthracene":                "Benz[a]anthracene",
    "benz[a]anthracene":                "Benz[a]anthracene",
    "benzo(g.h.i)perylene":             "h,i)perylene Benzo(g",
    "benzo(g,h,i)perylene":             "h,i)perylene Benzo(g",
    "benzo[g,h,i]perylene":             "h,i)perylene Benzo(g",
    "dibenz(a.h)anthracene":            "h]anthracene Dibenz[a",
    "dibenz(a,h)anthracene":            "h]anthracene Dibenz[a",
    "dibenz[a,h]anthracene":            "h]anthracene Dibenz[a",
    "indeno(1.2.3.cd)pyrene":           "2,3-cd]pyrene Indeno[1",
    "indeno(1,2,3-cd)pyrene":           "2,3-cd]pyrene Indeno[1",
    "indeno[1,2,3-cd]pyrene":           "2,3-cd]pyrene Indeno[1",
    "4-chloroaniline":                  "p-Chloroaniline",
    "p-chloroaniline":                  "p-Chloroaniline",
    "1-chloronaphthalene":              "Beta-Chloronaphthalene",
    "2-chloronaphthalene":              "Beta-Chloronaphthalene",
    "6-caprolactam":                    "Caprolactam",
    "2.4.5-trichlorophenol":            "Trichlorophenol, 2,4,5-",
    "2,4,5-trichlorophenol":            "Trichlorophenol, 2,4,5-",
    "2.4.6-trichlorophenol":            "Trichlorophenol, 2,4,6-",
    "2,4,6-trichlorophenol":            "Trichlorophenol, 2,4,6-",
    "2.6-dichlorophenol":               "2,6-Dimethylphenol",
    "2,6-dichlorophenol":               "2,6-Dimethylphenol",
    "4.6-dinitro-2-methylphenol":       "4,6-Dinitro-o-cresol",
    "4,6-dinitro-2-methylphenol":       "4,6-Dinitro-o-cresol",
    "3-nitroaniline":                   "3,5-Dinitroaniline",
    "bis(2-chloroisopropyl)ether":      "Bis(2-chloro-1-methylethyl) ether",
    "n-nitrosodi-n-propylamine":        "N-Nitroso-di-N-propylamine",
    "di-n-butyl phthalate":             "Dibutyl Phthalate",
    "butyl benzyl phthalate":           "Butyl Benzyl Phthalate",
    "di-n-octyl phthalate":             "di-N-Octyl Phthalate",
    "diethyl phthalate":                "Diethyl Phthalate",
    "dimethyl phthalate":               "Dimethylterephthalate",
    "4-chloro-3-methylphenol":          "p-chloro-m-Cresol",
    "1.4-dioxane":                      "1,4-Dioxane",
    "1.2-dichlorobenzene":              "1,2-Dichlorobenzene",
    "1.4-dichlorobenzene":              "1,4-Dichlorobenzene",
    "1.1-dichloroethane":               "1,1-Dichloroethane",
    "1.2-dichloropropane":              "1,2-Dichloropropane",
    "2.4-dimethylphenol":               "2,4-Dimethylphenol",
    "2.4-dichlorophenol":               "2,4-Dichlorophenol",
    "2.4-dinitrophenol":                "2,4-Dinitrophenol",
    "2.4-dinitrotoluene":               "2,4-Dinitrotoluene",
    "2.6-dinitrotoluene":               "2,6-Dinitrotoluene",
    "bis(2-ethylhexyl)phthalate":       "Bis(2-ethylhexyl)phthalate",
    "1,1'-biphenyl":                   "1,1'-Biphenyl",
    "n-butylbenzene":                   "n-Butylbenzene",
    "vinyl chloride":                   "Vinyl Chloride",
}

# ── PFAS ALIAS ────────────────────────────────────────────────────────────────────
PFAS_ALIAS = {
    "2,3,3,3-tetrafluoro-2-(heptafluoropropoxy)propanoic acid (hfpo-da)":
        "hexafluoropropylene oxide dimer acid (hfpo-da)",
    "7h-perfluoroheptanoic acid (hpfhpa)":      "perfluoroheptanoic acid (pfhpa)",
    "perfluorobutane sulfonic acid (pfbs)":     "perfluorobutanesulfonic acid (pfbs)",
    "perfluorobutane sulfonate (pfbs)":         "perfluorobutanesulfonic acid (pfbs)",
    "perfluorohexane sulfonic acid (pfhxs)":    "perfluorohexanesulfonic acid (pfhxs)",
    "perfluorohexane sulfonate (pfhxs)":        "perfluorohexanesulfonic acid (pfhxs)",
    "perfluorooctane sulfonic acid (pfos)":     "perfluorooctanesulfonic acid (pfos)",
    "perfluorooctane sulfonate (pfos)":         "perfluorooctanesulfonic acid (pfos)",
    "perfluorooctadecanoic acid (pfocda)":      "perfluorooctadecanoic acid (pfoda)",
    "perfluoroundecanoic acid (pfunda)":        "perfluoroundecanoic acid (pfuda)",
    "perfluorotetradecanoic acid (pfcpda)":     "perfluorotetradecanoic acid (pfteta)",
    "perfluorodecane sulfonic acid (pfds)":     "perfluorodecanesulfonic acid (pfds)",
    "perfluoroheptane sulfonic acid (pfhps)":   "perfluoroheptanesulfonic acid (pfhps)",
    "perfluoropentane sulfonic acid (pfpes)":   "perfluoropentanesulfonic acid (pfpes)",
    "perfluorooctane sulfonamide (fosa)":       "perfluorooctanesulfonamide (fosa)",
    "perfluoropentanoic acid (pfpea)":          "perfluoropentanoic acid (pfpea)",
    "perfluorodecanoic acid (pfda)":            "perfluorodecanoic acid (pfda)",
    "perfluorododecanoic acid (pfdoda)":        "perfluorododecanoic acid (pfdoda)",
    "perfluoroheptanoic acid (pfhpa)":          "perfluoroheptanoic acid (pfhpa)",
    "perfluorotridecanoic acid (pftrda)":       "perfluorotridecanoic acid (pftrda)",
    "perfluorooctanesulfonic acid (pfos)":      "perfluorooctanesulfonic acid (pfos)",
}

CANONICAL_KEY = "__CANONICAL_MAP_INTERNAL__"

def canonical_compound(name: str) -> str:
    s = norm(name)
    s = s.replace("ethylene", "ethen").replace("ethene", "ethen")
    s = re.sub(r"\s*\([^)]+\)\s*$", "", s)
    nums = re.findall(r"\d+", s)
    num_part = ",".join(nums) if nums else ""
    base_no_nums = re.sub(r"[0-9.,/\-]", " ", s)
    base_no_nums = re.sub(r"\s+", "", base_no_nums)
    canon = (num_part + " " + base_no_nums).strip()
    return canon

def match_threshold(compound_name, thresh_dict):
    key = norm(compound_name)
    canon_map = thresh_dict.get(CANONICAL_KEY)

    for k in (key, key.replace(".",","), key.replace(",",".")):
        if k in thresh_dict and k != CANONICAL_KEY: return thresh_dict[k]

    # VOC alias
    aliased_voc = VOC_ALIAS.get(key) or VOC_ALIAS.get(key.replace(".",","))
    if aliased_voc:
        a_key = norm(aliased_voc)
        for k in (a_key, a_key.replace(".",","), a_key.replace(",",".")):
            if k in thresh_dict and k != CANONICAL_KEY: return thresh_dict[k]

    # PFAS alias
    aliased = PFAS_ALIAS.get(key)
    if aliased:
        a_key = norm(aliased)
        for k in (a_key, a_key.replace(".",","), a_key.replace(",",".")):
            if k in thresh_dict and k != CANONICAL_KEY: return thresh_dict[k]

    stripped = re.sub(r"\s*\([A-Z0-9:_\-]+\)\s*$", "", compound_name).strip().lower()
    for k, v in thresh_dict.items():
        if k == CANONICAL_KEY: continue
        k_s = re.sub(r"\s*\([A-Z0-9:_\-]+\)\s*$", "", k).strip().lower()
        if len(stripped) > 8 and stripped == k_s: return v

    if canon_map:
        ck = canonical_compound(compound_name)
        hit = canon_map.get(ck)
        if hit: return hit

    return {}

def load_threshold_file(file_bytes):
    wb = load_workbook(io.BytesIO(file_bytes), data_only=True)
    ws = wb.active
    thresh = {}
    for row in ws.iter_rows(min_row=6, values_only=True):
        if not row[0]: continue
        name = str(row[0]).strip()
        cas  = str(row[1]).strip() if row[1] else "-"
        def g(ci):
            return (row[ci] if ci < len(row) and row[ci] is not None
                    and str(row[ci]) not in ("NA","") else None)
        thresh[norm(name)] = {
            "name": name, "cas": cas,
            "units": str(row[3]) if row[3] else "mg/kg",
            "VSL": g(4), "Ind_A_06": g(8), "Ind_A_6p": g(9),
            "Ind_B": g(10), "Res_A_06": g(11), "Res_A_6p": g(12), "Res_B": g(13),
        }
    canon_map = {}
    for k, v in thresh.items():
        ck = canonical_compound(k)
        canon_map.setdefault(ck, v)
    thresh[CANONICAL_KEY] = canon_map
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
    return t.get("VSL"), t.get(t1col), t.get("cas", "-")

def build_metals_thresh(thresh_dict, t1col):
    result = {}
    for key, v in thresh_dict.items():
        if key == CANONICAL_KEY: continue
        sym = THRESH_METAL_MAP.get(key)
        if sym and sym not in result:
            result[sym] = {"vsl": v.get("VSL"), "tier1": v.get(t1col), "cas": v.get("cas","-")}
    return result

def parse_als_file(file_bytes, filename):
    """
    Parser גנרי - עובד עם כל קבצי ALS:
    - מחפש Client Sample ID לפי תוכן (לא לפי מיקום שורה)
    - מחפש Parameter לפי תוכן
    - מזהה אוטומטית עמודות Unit, LOR, VSL
    - עובד עם PFAS / VOC / Metals / TPH באותו קוד
    """
    try:
        wb = load_workbook(io.BytesIO(file_bytes), data_only=True)
    except Exception as e:
        return None, str(e)

    main = next((wb[n] for n in wb.sheetnames if "Client" in n and "SOIL" in n), wb.worksheets[0])
    rows = list(main.iter_rows(values_only=True))

    # ── מצא שורת Client Sample ID ──────────────────────────────────────────────
    sid_row_idx = next(
        (i for i,r in enumerate(rows)
         if any("Client Sample ID" in str(v) for v in r if v)), None)
    if sid_row_idx is None:
        return None, "לא נמצאה שורת Client Sample ID"

    sid_row = rows[sid_row_idx]
    # מצא את העמודה שמכילה "Client Sample ID"
    sid_label_col = next(ci for ci,v in enumerate(sid_row) if v and "Client Sample ID" in str(v))
    # כל העמודות אחריה = שמות הדגימות
    col2sample = {
        ci: str(v).strip()
        for ci,v in enumerate(sid_row)
        if ci > sid_label_col and v and str(v).strip() not in ("", "None")
    }

    # ── מצא שורת Parameter ─────────────────────────────────────────────────────
    ph_idx = next(
        (i for i,r in enumerate(rows) if r and r[0] == "Parameter"), None)
    if ph_idx is None:
        return None, "לא נמצאה שורת Parameter"

    param_row = rows[ph_idx]

    # זהה אוטומטית עמודות Unit ו-LOR לפי כותרת (לא לפי מיקום קשיח)
    unit_col = next(
        (ci for ci,v in enumerate(param_row)
         if v and str(v).strip().lower() == "unit"), 2)
    lor_col = next(
        (ci for ci,v in enumerate(param_row)
         if v and str(v).strip().lower() == "lor"), unit_col + 1)

    # ── קרא נתונים ─────────────────────────────────────────────────────────────
    records = []
    group = "Unknown"

    for row in rows[ph_idx + 1:]:
        p = row[0] if len(row) > 0 else None
        if not p or str(p).strip() in ("", "None"):
            continue

        # שורת קבוצה: יש שם אבל אין Method (עמודה 1)
        method = row[1] if len(row) > 1 else None
        if not method or str(method).strip() in ("", "None"):
            group = str(p).strip()
            continue

        u   = row[unit_col] if unit_col < len(row) else None
        lor = row[lor_col]  if lor_col  < len(row) else None

        for ci, sname in col2sample.items():
            sid, depth_val = parse_sample(sname)
            if sid is None:
                continue
            val = row[ci] if ci < len(row) else None
            rs  = str(val).strip() if val is not None else ""
            if rs in ("", "None"):
                continue
            result = None
            if rs.startswith("<"):
                result = 0.0
            else:
                try: result = float(rs)
                except: result = None
            if result is not None:
                lor_val = None
                if rs.startswith("<"):
                    try: lor_val = float(rs[1:].strip())
                    except: lor_val = 0.0
                records.append({
                    "sample_id":      sid,
                    "depth":          depth_val,
                    "compound":       str(p).strip(),
                    "compound_lower": norm(p),
                    "unit":           str(u).strip() if u else "mg/kg",
                    "lor":            lor,
                    "result":         result,
                    "result_str":     rs,
                    "lor_val":        lor_val,
                    "group":          group,
                    "source":         filename,
                })

    if not records:
        return None, "לא נמצאו נתונים"
    return pd.DataFrame(records), None

# ── TPH SHEET ─────────────────────────────────────────────────────────────────────
def write_tph_sheet(ws, df, thresh_dict, t1col, t1lbl):
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

    vsl_d,t1_d,_ = get_thresh("C10 - C28 Fraction (DRO)", thresh_dict, t1col)
    vsl_o,t1_o,_ = get_thresh("C24 - C40 Fraction (ORO)", thresh_dict, t1col)
    vsl_t,t1_t,_ = get_thresh("TPH - DRO + ORO (Tier 1)", thresh_dict, t1col)
    vv = [v for v in [vsl_d,vsl_o,vsl_t] if v]
    tt = [v for v in [t1_d,t1_o,t1_t] if v]
    vsl_tot = min(vv) if vv else 350
    t1_tot  = min(tt) if tt else 350

    for ci,h in enumerate(["שם קידוח","עומק","TPH DRO","TPH ORO","Total TPH"],1):
        style_hdr(ws.cell(1,ci,h),HDR_BLUE_FILL)
    ws.merge_cells(start_row=1,start_column=1,end_row=5,end_column=1)
    c=ws.cell(1,1,"שם קידוח"); c.font=Font(bold=True,name="Arial",size=11)
    c.fill=HDR_BLUE_FILL; c.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True); c.border=thin_border()
    sub_rows=["יחידות","CAS","VSL",t1lbl]
    sub_vals={"יחידות":"mg/kg","CAS":"C10-C40","VSL":vsl_tot,t1lbl:t1_tot}
    for ri,lbl in enumerate(sub_rows,2):
        style_hdr(ws.cell(ri,2,lbl),HDR_BLUE_FILL)
        for ci in [3,4,5]: style_hdr(ws.cell(ri,ci,sub_vals[lbl]),HDR_BLUE_FILL)

    pivoted = {}
    for _,r in df.iterrows():
        k=(r["sample_id"],r["depth"])
        if k not in pivoted: pivoted[k]={"DRO":"","ORO":"","TOT":"","DRO_f":None,"ORO_f":None,"DRO_lor":None,"ORO_lor":None}
        c=r["compound_lower"]
        if is_dro(c) and not pivoted[k]["DRO"]:
            pivoted[k]["DRO"]=r["result_str"]; pivoted[k]["DRO_f"]=r["result"]; pivoted[k]["DRO_lor"]=r.get("lor_val")
        elif is_oro(c) and not pivoted[k]["ORO"]:
            pivoted[k]["ORO"]=r["result_str"]; pivoted[k]["ORO_f"]=r["result"]; pivoted[k]["ORO_lor"]=r.get("lor_val")
        elif is_total(c) and not pivoted[k]["TOT"]:
            pivoted[k]["TOT"]=r["result_str"]

    ri=6; prev_sid=None; sid_rows={}
    for (sid,depth_val),v in sorted(pivoted.items(),key=lambda x:(sort_key(x[0][0]),x[0][1] or 0)):
        if v["TOT"]: total_s=v["TOT"]
        else:
            dro_lor=v["DRO"] and str(v["DRO"]).startswith("<")
            oro_lor=v["ORO"] and str(v["ORO"]).startswith("<")
            dro_empty=not v["DRO"]; oro_empty=not v["ORO"]
            dro_num=(v["DRO_lor"] if dro_lor and v["DRO_lor"] is not None else (v["DRO_f"] or 0))
            oro_num=(v["ORO_lor"] if oro_lor and v["ORO_lor"] is not None else (v["ORO_f"] or 0))
            total_f=dro_num+oro_num
            if (dro_lor or dro_empty) and (oro_lor or oro_empty) and not (dro_empty and oro_empty):
                total_s=f"<{total_f:.0f}"
            else: total_s=f"{total_f:.0f}"
        hl_dro=check_exceed(v["DRO"],vsl_tot,t1_tot)
        hl_oro=check_exceed(v["ORO"],vsl_tot,t1_tot)
        hl_total=check_exceed(total_s,vsl_tot,t1_tot)
        if sid!=prev_sid: sid_rows[sid]=[]
        sid_rows[sid].append(ri)
        sid_val=sid if sid!=prev_sid else None; prev_sid=sid
        style_data(ws.cell(ri,1,sid_val)); style_data(ws.cell(ri,2,depth_val))
        style_data(ws.cell(ri,3,v["DRO"]),hl_dro); style_data(ws.cell(ri,4,v["ORO"]),hl_oro)
        style_data(ws.cell(ri,5,total_s),hl_total); ri+=1
    apply_sid_merge(ws,sid_rows,col=1)
    for col,w in zip("ABCDE",[14,10,16,16,16]): ws.column_dimensions[col].width=w
    ws.row_dimensions[1].height=15; ws.freeze_panes="A6"
    ws.sheet_view.rightToLeft=True

# ── METALS SHEET ──────────────────────────────────────────────────────────────────
def write_metals_sheet(ws, df, thresh_dict, t1col, t1lbl):
    df=df.copy(); df["sym"]=df["compound_lower"].map(METAL_MAP); df=df[df["sym"].notna()]
    if df.empty: ws.cell(1,1,"אין נתוני מתכות"); return
    present=set(df["sym"].unique())
    metals=[m for m in METALS_ORDER if m in present]+sorted(present-set(METALS_ORDER))
    mt=build_metals_thresh(thresh_dict,t1col)
    ws.merge_cells(start_row=1,start_column=1,end_row=5,end_column=1)
    c=ws.cell(1,1,"שם קידוח"); c.font=Font(bold=True,name="Arial",size=11)
    c.fill=HDR_BLUE_FILL; c.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True); c.border=thin_border()
    for ci,h in enumerate(["עומק"]+metals,2): style_hdr(ws.cell(1,ci,h),HDR_BLUE_FILL)
    for ri,lbl in enumerate(["יחידות","CAS","VSL",t1lbl],2):
        style_hdr(ws.cell(ri,2,lbl),HDR_BLUE_FILL)
        for ci,sym in enumerate(metals,3):
            t=mt.get(sym,{})
            val={"יחידות":"mg/kg","CAS":t.get("cas","-"),"VSL":t.get("vsl","-"),t1lbl:t.get("tier1","-")}.get(lbl,"-")
            style_hdr(ws.cell(ri,ci,val),HDR_BLUE_FILL)
    pt=df.pivot_table(index=["sample_id","depth"],columns="sym",values="result_str",aggfunc="first")
    pt=pt.reindex(sorted(pt.index,key=lambda x:(sort_key(x[0]),x[1] or 0)))
    ri=6; prev_sid=None; sid_rows={}
    for (sid,depth_val),row_data in pt.iterrows():
        if sid!=prev_sid: sid_rows[sid]=[]
        sid_rows[sid].append(ri)
        sid_val=sid if sid!=prev_sid else None; prev_sid=sid
        style_data(ws.cell(ri,1,sid_val)); style_data(ws.cell(ri,2,depth_val))
        for ci,sym in enumerate(metals,3):
            val=row_data.get(sym,"") or ""
            val="" if str(val)=="nan" else str(val)
            hl=check_exceed(val,mt.get(sym,{}).get("vsl"),mt.get(sym,{}).get("tier1"))
            style_data(ws.cell(ri,ci,val),hl)
        ri+=1
    apply_sid_merge(ws,sid_rows,col=1)
    ws.column_dimensions["A"].width=14; ws.column_dimensions["B"].width=10
    for ci in range(3,len(metals)+3): ws.column_dimensions[get_column_letter(ci)].width=11
    ws.freeze_panes="C6"
    ws.sheet_view.rightToLeft=True

# ── PFAS SHEET ────────────────────────────────────────────────────────────────────
def write_pfas_sheet(ws, df, thresh_dict, t1col, t1lbl):
    df=df.copy()
    pairs_pfas=sorted(df[["sample_id","depth"]].drop_duplicates().values.tolist(),
                      key=lambda x:(sort_key(x[0]),-(x[1] or 0)))
    def to_ug(v):
        if v is None: return None
        try: return round(float(v)*1000,6)
        except: return v
    fixed_hdrs=["שם התרכובת","CAS","VSL [µg/kg]",t1lbl,"יחידות"]
    for ci,h in enumerate(fixed_hdrs,1):
        style_hdr(ws.cell(1,ci,h),HDR_BLUE_FILL)
        ws.merge_cells(start_row=1,start_column=ci,end_row=2,end_column=ci)
        ws.cell(1,ci).alignment=Alignment(horizontal="center",vertical="center",wrap_text=True)
        ws.cell(1,ci).fill=HDR_BLUE_FILL; ws.cell(1,ci).border=thin_border()
    style_hdr(ws.cell(1,6,"LOR"),HDR_BLUE_FILL); style_hdr(ws.cell(2,6,"[µg/kg]"),HDR_BLUE_FILL)
    style_hdr(ws.cell(1,7,"שם קידוח"),HDR_BLUE_FILL); style_hdr(ws.cell(2,7,"עומק"),HDR_BLUE_FILL)
    prev_sid=None; sid_merge_start_p={}
    for ci,(sid,depth_val) in enumerate(pairs_pfas,8):
        sid_val=sid if sid!=prev_sid else None
        style_hdr(ws.cell(1,ci,sid_val),HDR_BLUE_FILL)
        style_hdr(ws.cell(2,ci,depth_val),HDR_BLUE_FILL)
        if sid!=prev_sid: sid_merge_start_p[sid]=ci
        prev_sid=sid
    for sid,start_ci in sid_merge_start_p.items():
        cols_p=[ci for ci,(s,_) in enumerate(pairs_pfas,8) if s==sid]
        if len(cols_p)>1:
            ws.merge_cells(start_row=1,start_column=start_ci,end_row=1,end_column=cols_p[-1])
            c=ws.cell(1,start_ci); c.alignment=Alignment(horizontal="center",vertical="center")
            c.fill=HDR_BLUE_FILL; c.border=thin_border()
    for row_i,cmp in enumerate(df["compound"].unique(),3):
        df_c=df[df["compound"]==cmp]
        vsl_mg,tier1_mg,cas=get_thresh(cmp,thresh_dict,t1col)
        vsl=to_ug(vsl_mg); tier1=to_ug(tier1_mg)
        unit=df_c.iloc[0]["unit"] if not df_c.empty else "µg/kg"
        lor=df_c.iloc[0]["lor"] if not df_c.empty else ""
        style_data(ws.cell(row_i,1,cmp),left_align=True)
        for ci,val in enumerate([cas,vsl,tier1,unit],2): style_data(ws.cell(row_i,ci,val))
        ws.merge_cells(start_row=row_i,start_column=6,end_row=row_i,end_column=6)
        style_data(ws.cell(row_i,6,lor)); style_data(ws.cell(row_i,7,None))
        for ci,(sid,depth_val) in enumerate(pairs_pfas,8):
            sub=df_c[(df_c["sample_id"]==sid)&(df_c["depth"]==depth_val)]
            rs=sub.iloc[0]["result_str"] if not sub.empty else ""
            style_data(ws.cell(row_i,ci,rs),check_exceed(rs,vsl,tier1))
    ws.column_dimensions["A"].width=50
    for ci in range(2,8): ws.column_dimensions[get_column_letter(ci)].width=13
    for ci in range(8,8+len(pairs_pfas)): ws.column_dimensions[get_column_letter(ci)].width=12
    ws.row_dimensions[1].height=60
    ws.freeze_panes="H3"
    ws.sheet_view.rightToLeft=True

# ── VOC+SVOC SHEET ────────────────────────────────────────────────────────────────
def write_voc_sheet(ws, df, thresh_dict, t1col, t1lbl):
    df=df.copy()
    pairs=sorted(df[["sample_id","depth"]].drop_duplicates().values.tolist(),
                 key=lambda x:(sort_key(x[0]),-(x[1] or 0)))
    # Headers A-F: merge rows 1-2
    for ci,h in enumerate(["קבוצה","קבוצה","שם התרכובת","CAS","VSL",t1lbl],1):
        ws.merge_cells(start_row=1,start_column=ci,end_row=2,end_column=ci)
        style_hdr(ws.cell(1,ci,h),HDR_BLUE_FILL,sz=9)
        ws.cell(1,ci).alignment=Alignment(horizontal="center",vertical="center",wrap_text=True)
        ws.cell(1,ci).fill=HDR_BLUE_FILL; ws.cell(1,ci).border=thin_border()
    # G:H merged = יחידות
    ws.merge_cells(start_row=1,start_column=7,end_row=2,end_column=8)
    style_hdr(ws.cell(1,7,"יחידות"),HDR_BLUE_FILL,sz=9)
    ws.cell(1,7).alignment=Alignment(horizontal="center",vertical="center",wrap_text=True)
    ws.cell(1,7).fill=HDR_BLUE_FILL; ws.cell(1,7).border=thin_border()
    # col I: שם קידוח / עומק
    style_hdr(ws.cell(1,9,"שם קידוח"),HDR_BLUE_FILL,sz=9)
    style_hdr(ws.cell(2,9,"עומק"),HDR_BLUE_FILL,sz=9)
    # sample columns from col 10
    prev_sid=None; sid_col_start={}
    for ci,(sid,depth_val) in enumerate(pairs,10):
        sid_val=sid if sid!=prev_sid else None
        style_hdr(ws.cell(1,ci,sid_val),HDR_BLUE_FILL,sz=9)
        style_hdr(ws.cell(2,ci,depth_val),HDR_BLUE_FILL,sz=9)
        if sid!=prev_sid: sid_col_start[sid]=ci
        prev_sid=sid
    for sid,sc in sid_col_start.items():
        cols=[ci for ci,(s,_) in enumerate(pairs,10) if s==sid]
        if len(cols)>1:
            ws.merge_cells(start_row=1,start_column=sc,end_row=1,end_column=cols[-1])
            c=ws.cell(1,sc); c.alignment=Alignment(horizontal="center",vertical="center")
            c.fill=HDR_BLUE_FILL; c.border=thin_border()
    # ALS lookup
    als_data={}
    for _,r in df.iterrows():
        k=norm(r["compound"]); als_data.setdefault(k,{})[(r["sample_id"],r["depth"])]=r["result_str"]
    for k in list(als_data.keys()):
        for alt in (k.replace(".",","),k.replace(",",".")):
            if alt not in als_data: als_data[alt]=als_data[k]
    # data rows
    for row_i,(vs,grp,cmp) in enumerate(VOC_COMPOUND_ORDER,3):
        vsl,tier1,cas=get_thresh(cmp,thresh_dict,t1col)
        cmp_key=norm(cmp)
        cmp_data=als_data.get(cmp_key) or als_data.get(cmp_key.replace(".",",")) or {}
        style_data(ws.cell(row_i,1,vs),sz=9); style_data(ws.cell(row_i,2,grp),sz=9)
        style_data(ws.cell(row_i,3,cmp),sz=9,left_align=True); style_data(ws.cell(row_i,4,cas),sz=9)
        style_data(ws.cell(row_i,5,vsl),sz=9); style_data(ws.cell(row_i,6,tier1),sz=9)
        ws.merge_cells(start_row=row_i,start_column=7,end_row=row_i,end_column=8)
        c=ws.cell(row_i,7,"mg/kg"); c.font=Font(name="Arial",size=9)
        c.alignment=Alignment(horizontal="center",vertical="center"); c.border=thin_border()
        style_data(ws.cell(row_i,9,None),sz=9)
        for ci,(sid,depth_val) in enumerate(pairs,10):
            rs=cmp_data.get((sid,depth_val),"")
            style_data(ws.cell(row_i,ci,rs),check_exceed(rs,vsl,tier1),sz=9)
    # merge col A: VOCs rows 3-32, SVOCs rows 33-96
    for r1,r2,val in [(3,32,"VOCs"),(33,96,"SVOCs")]:
        ws.merge_cells(start_row=r1,start_column=1,end_row=r2,end_column=1)
        c=ws.cell(r1,1,val); c.font=Font(bold=True,name="Arial",size=9)
        c.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True); c.border=thin_border()
    # merge col B
    b_ranges=[
        (3,12,"Non-Halogenated VOCs"),(13,16,"BTEX"),(17,32,"Halogenated VOCs"),
        (33,37,"Phenols & Naphtols"),(38,53,"PAHs"),(54,57,"Anilines"),
        (58,65,"Aromatic Compounds"),(66,66,"Alcohols"),(67,69,"Aldehydes / Ketones"),
        (70,75,"Chlorophenols"),(76,85,"Nitroaromatic Compounds"),
        (86,88,"Chlorinated Hydrocarbons"),(89,89,"Nitrosoamines"),
        (90,90,"Pesticides"),(91,96,"Phthalates"),
    ]
    for r1,r2,val in b_ranges:
        if r2>r1: ws.merge_cells(start_row=r1,start_column=2,end_row=r2,end_column=2)
        c=ws.cell(r1,2,val); c.font=Font(bold=True,name="Arial",size=9)
        c.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True); c.border=thin_border()
    ws.column_dimensions["A"].width=8; ws.column_dimensions["B"].width=22
    ws.column_dimensions["C"].width=35; ws.column_dimensions["D"].width=12
    ws.column_dimensions["E"].width=10; ws.column_dimensions["F"].width=12
    ws.column_dimensions["G"].width=7;  ws.column_dimensions["H"].width=7
    ws.column_dimensions["I"].width=12
    for ci in range(10,10+len(pairs)): ws.column_dimensions[get_column_letter(ci)].width=10
    ws.row_dimensions[1].height=60; ws.row_dimensions[2].height=15
    ws.freeze_panes="J3"
    ws.sheet_view.rightToLeft=True

# ── SIDEBAR ───────────────────────────────────────────────────────────────────────
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

with st.expander(f"👁️ תצוגה מקדימה ({len(df_all)} שורות)"): st.dataframe(df_all.head(30),use_container_width=True)
with st.expander("קבוצות שנמצאו"): st.write(df_all["group"].unique().tolist())

def dg(kw): return df_all[df_all["group"].str.contains("|".join(kw),case=False,na=False)]
tph_df   =dg(["petroleum","tph","hydrocarbon"])
metals_df=dg(["metal","cation","extractable"])
pfas_df  =dg(["perfluor","pfas","fluorin"])
voc_df   =dg(["voc","svoc","btex","aromatic","halogenated","volatile",
               "alcohol","aldehyde","ketone","phenol","pah","aniline",
               "nitro","phthalate","pesticide","pcb","other"])

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

