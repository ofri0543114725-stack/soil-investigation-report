import streamlit as st
import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import io
import re

# ── BASIC PAGE SETUP ─────────────────────────────────────────────────────────────
st.set_page_config(page_title="דוח סקר קרקע", layout="wide", page_icon="🧪")
st.title("🧪 מערכת עיבוד תוצאות מעבדה")
st.caption("v3.5 - Robust VOC threshold matching (spaces, dots/commas, cis/trans)")
st.markdown("---")

# ── STYLES ───────────────────────────────────────────────────────────────────────
YELLOW_FILL   = PatternFill("solid", fgColor="FFFF00")   # VSL
ORANGE_FILL   = PatternFill("solid", fgColor="FFC000")   # TIER1
HDR_BLUE_FILL = PatternFill("solid", fgColor="B7D7F0")
HDR_CYAN_FILL = PatternFill("solid", fgColor="00B0F0")


def thin_border():
    s = Side(style="thin", color="000000")
    return Border(left=s, right=s, top=s, bottom=s)


def style_hdr(cell, fill=None, sz=11):
    cell.font      = Font(bold=True, name="Arial", size=sz)
    cell.alignment = Alignment(horizontal="center",
                               vertical="center",
                               wrap_text=True)
    cell.border    = thin_border()
    if fill:
        cell.fill = fill


def style_data(cell, hl=None, sz=10):
    cell.font      = Font(bold=bool(hl), name="Arial", size=sz)
    cell.alignment = Alignment(horizontal="center",
                               vertical="center",
                               wrap_text=True)
    cell.border    = thin_border()
    if   hl == "tier1":
        cell.fill = ORANGE_FILL
    elif hl == "vsl":
        cell.fill = YELLOW_FILL


# ── HELPERS ──────────────────────────────────────────────────────────────────────
def norm(s):
    """Normalize strings for matching (lowercase, trim, collapse spaces)."""
    s = "" if s is None else str(s).strip().lower()
    return re.sub(r"\s+", " ", s.replace("\xa0", " "))


def to_float(v):
    s = str(v).strip() if v is not None else ""
    try:
        return float(s.lstrip("<>").strip())
    except Exception:
        return None


def sort_key(sid):
    """Sort S1, S2, S10… numerically."""
    m = re.match(r"S-?(\d+)", str(sid), re.I)
    return int(m.group(1)) if m else 9999


def parse_sample(sname):
    """
    Parse ALS 'Client Sample ID' like:
    - S1 (1.0)
    - S1-1.0
    Skip DUP.
    """
    sname = str(sname).strip()
    if "DUP" in sname.upper():
        return None, None
    m = re.match(r"^(S\d+[A-Za-z0-9]*)\s*\(([0-9.]+)\)", sname)
    if m:
        return m.group(1), float(m.group(2))
    m = re.match(r"^(S\d+)-([0-9]+\.?[0-9]*)$", sname)
    if m:
        return m.group(1), float(m.group(2))
    return sname, None


def check_exceed(val_str, vsl, tier1):
    """VSL=yellow, TIER1=orange. Only color real detections, not <LOR."""
    if not val_str or str(val_str).strip().startswith("<"):
        return None
    f = to_float(val_str)
    if f is None:
        return None
    try:
        t1f = float(tier1) if (tier1 is not None and
                               str(tier1) not in ("-", "NA", "") and
                               pd.notna(tier1)) else None
        vf  = float(vsl)   if (vsl   is not None and
                               str(vsl)   not in ("-", "NA", "") and
                               pd.notna(vsl))   else None
        if t1f and t1f > 0 and f > t1f:
            return "tier1"
        if vf and vf > 0 and f > vf:
            return "vsl"
    except Exception:
        pass
    return None


def apply_sid_merge(ws, sid_rows, col=1):
    """Merge sample ID column where same קידוח מופיע בכמה עומקים."""
    for sid, rows_list in sid_rows.items():
        if len(rows_list) > 1:
            ws.merge_cells(start_row=rows_list[0], start_column=col,
                           end_row=rows_list[-1], end_column=col)
            c = ws.cell(rows_list[0], col)
            c.alignment = Alignment(horizontal="center",
                                    vertical="center",
                                    wrap_text=True)
            c.border = thin_border()


# ── METALS MAPS ──────────────────────────────────────────────────────────────────
METAL_MAP = {
    "aluminium": "Al", "aluminum": "Al", "antimony": "Sb", "arsenic": "As",
    "barium": "Ba", "beryllium": "Be", "bismuth": "Bi", "boron": "B",
    "cadmium": "Cd", "calcium": "Ca", "chromium": "Cr", "cobalt": "Co",
    "copper": "Cu", "iron": "Fe", "lead": "Pb", "lithium": "Li",
    "magnesium": "Mg", "manganese": "Mn", "mercury": "Hg", "nickel": "Ni",
    "potassium": "K", "selenium": "Se", "silver": "Ag", "sodium": "Na",
    "vanadium": "V", "zinc": "Zn", "molybdenum": "Mo", "tin": "Sn",
    "titanium": "Ti", "strontium": "Sr", "thallium": "Tl",
    "phosphorus": "P", "sulphur": "S", "silicon": "Si",
}

METALS_ORDER = [
    "Al", "Sb", "As", "Ba", "Be", "Bi", "B", "Cd", "Ca", "Cr", "Co", "Cu", "Fe",
    "Pb", "Li", "Mg", "Mn", "Hg", "Ni", "K", "Se", "Ag", "Na", "V", "Zn"
]

THRESH_METAL_MAP = {
    "aluminum": "Al",
    "antimony (metallic)": "Sb",
    "antimony": "Sb",
    "arsenic, inorganic": "As",
    "arsenic": "As",
    "barium": "Ba",
    "beryllium and compounds": "Be",
    "beryllium": "Be",
    "boron and borates only": "B",
    "boron": "B",
    "cadmium (water) source: water and air": "Cd",
    "cadmium": "Cd",
    "calcium": "Ca",
    "chromium, total": "Cr",
    "chromium": "Cr",
    "cobalt": "Co",
    "copper": "Cu",
    "iron": "Fe",
    "lead and compounds": "Pb",
    "lead": "Pb",
    "lithium": "Li",
    "magnesium": "Mg",
    "manganese (non-diet)": "Mn",
    "manganese": "Mn",
    "mercuric chloride (and other mercury salts)": "Hg",
    "mercury": "Hg",
    "nickel soluble salts": "Ni",
    "nickel": "Ni",
    "potassium": "K",
    "selenium": "Se",
    "silver": "Ag",
    "sodium": "Na",
    "vanadium and compounds": "V",
    "vanadium": "V",
    "zinc and compounds": "Zn",
    "zinc": "Zn",
    "molybdenum": "Mo",
    "tin": "Sn",
    "titanium": "Ti",
    "strontium": "Sr",
    "thallium": "Tl",
    "phosphorus": "P",
    "sulphur": "S",
    "silicon": "Si",
}

# ── VOC/SVOC ORDER (טמפלט ייחוס) ────────────────────────────────────────────────
VOC_COMPOUND_ORDER = [
    # VOCs - Non-Halogenated VOCs (10)
    ("VOCs", "Non-Halogenated VOCs", "1.2.4-Trimethylbenzene"),
    ("VOCs", "Non-Halogenated VOCs", "1.3.5-Trimethylbenzene"),
    ("VOCs", "Non-Halogenated VOCs", "MTBE"),
    ("VOCs", "Non-Halogenated VOCs", "Styrene"),
    ("VOCs", "Non-Halogenated VOCs", "n-Butylbenzene"),
    ("VOCs", "Non-Halogenated VOCs", "n-Propylbenzene"),
    ("VOCs", "Non-Halogenated VOCs", "Isopropylbenzene"),
    ("VOCs", "Non-Halogenated VOCs", "Acetone"),
    ("VOCs", "Non-Halogenated VOCs", "2-Butanone (MEK)"),
    ("VOCs", "Non-Halogenated VOCs", "1.4-Dioxane"),
    # VOCs - BTEX (4)
    ("VOCs", "BTEX", "Benzene"),
    ("VOCs", "BTEX", "Toluene"),
    ("VOCs", "BTEX", "Ethylbenzene"),
    ("VOCs", "BTEX", "Sum of Xylenes"),
    # VOCs - Halogenated VOCs (16)
    ("VOCs", "Halogenated VOCs", "1.1-Dichloroethane"),
    ("VOCs", "Halogenated VOCs", "1.1-Dichloroethene"),
    ("VOCs", "Halogenated VOCs", "1.2-Dichloroethane"),
    ("VOCs", "Halogenated VOCs", "1.2-Dichloropropane"),
    ("VOCs", "Halogenated VOCs", "Chlorobenzene"),
    ("VOCs", "Halogenated VOCs", "Chloroform"),
    ("VOCs", "Halogenated VOCs", "Dichloromethane"),
    ("VOCs", "Halogenated VOCs", "Tetrachloroethene"),
    ("VOCs", "Halogenated VOCs", "Tetrachloromethane"),
    ("VOCs", "Halogenated VOCs", "Trichloroethene"),
    ("VOCs", "Halogenated VOCs", "Vinyl chloride"),
    ("VOCs", "Halogenated VOCs", "cis-1.2-Dichloroethene"),
    ("VOCs", "Halogenated VOCs", "trans-1.2-Dichloroethene"),
    ("VOCs", "Halogenated VOCs", "1.4-Dichlorobenzene"),
    ("VOCs", "Halogenated VOCs", "1.2-Dichlorobenzene"),
    ("VOCs", "Halogenated VOCs", "1.3-Dichlorobenzene"),
    # SVOCs - Phenols & Naphtols (5)
    ("SVOCs", "Phenols & Naphtols", "2.4-Dimethylphenol"),
    ("SVOCs", "Phenols & Naphtols", "2-Methylphenol"),
    ("SVOCs", "Phenols & Naphtols", "3 & 4-Methylphenol"),
    ("SVOCs", "Phenols & Naphtols", "4-Chloro-3-methylphenol"),
    ("SVOCs", "Phenols & Naphtols", "Phenol"),
    # SVOCs - PAHs (16)
    ("SVOCs", "PAHs", "Acenaphthene"),
    ("SVOCs", "PAHs", "Acenaphthylene"),
    ("SVOCs", "PAHs", "Anthracene"),
    ("SVOCs", "PAHs", "Benz(a)anthracene"),
    ("SVOCs", "PAHs", "Benzo(a)pyrene"),
    ("SVOCs", "PAHs", "Benzo(b)fluoranthene"),
    ("SVOCs", "PAHs", "Benzo(g.h.i)perylene"),
    ("SVOCs", "PAHs", "Benzo(k)fluoranthene"),
    ("SVOCs", "PAHs", "Chrysene"),
    ("SVOCs", "PAHs", "Dibenz(a.h)anthracene"),
    ("SVOCs", "PAHs", "Fluoranthene"),
    ("SVOCs", "PAHs", "Fluorene"),
    ("SVOCs", "PAHs", "Indeno(1.2.3.cd)pyrene"),
    ("SVOCs", "PAHs", "Naphthalene"),
    ("SVOCs", "PAHs", "Phenanthrene"),
    ("SVOCs", "PAHs", "Pyrene"),
    # SVOCs - Anilines (4)
    ("SVOCs", "Anilines", "4-Chloroaniline"),
    ("SVOCs", "Anilines", "Aniline"),
    ("SVOCs", "Anilines", "Benzidine"),
    ("SVOCs", "Anilines", "Diphenylamine"),
    # SVOCs - Aromatic Compounds (8)
    ("SVOCs", "Aromatic Compounds", "1,1'-Biphenyl"),
    ("SVOCs", "Aromatic Compounds", "1-Chloronaphthalene"),
    ("SVOCs", "Aromatic Compounds", "2-Chloronaphthalene"),
    ("SVOCs", "Aromatic Compounds", "2-Methylnaphthalene"),
    ("SVOCs", "Aromatic Compounds", "4-Bromophenyl phenyl ether"),
    ("SVOCs", "Aromatic Compounds", "4-Chlorophenyl phenyl ether"),
    ("SVOCs", "Aromatic Compounds", "Carbazole"),
    ("SVOCs", "Aromatic Compounds", "Dibenzofuran"),
    # SVOCs - Alcohols (1)
    ("SVOCs", "Alcohols", "Benzyl Alcohol"),
    # SVOCs - Aldehydes / Ketones (3)
    ("SVOCs", "Aldehydes / Ketones", "6-Caprolactam"),
    ("SVOCs", "Aldehydes / Ketones", "Acetophenone"),
    ("SVOCs", "Aldehydes / Ketones", "Isophorone"),
    # SVOCs - Chlorophenols (6)
    ("SVOCs", "Chlorophenols", "2-Chlorophenol"),
    ("SVOCs", "Chlorophenols", "2.4.5-Trichlorophenol"),
    ("SVOCs", "Chlorophenols", "2.4.6-Trichlorophenol"),
    ("SVOCs", "Chlorophenols", "2.4-Dichlorophenol"),
    ("SVOCs", "Chlorophenols", "2.6-Dichlorophenol"),
    ("SVOCs", "Chlorophenols", "Pentachlorophenol"),
    # SVOCs - Nitroaromatic Compounds (10)
    ("SVOCs", "Nitroaromatic Compounds", "2.4-Dinitrophenol"),
    ("SVOCs", "Nitroaromatic Compounds", "2.4-Dinitrotoluene"),
    ("SVOCs", "Nitroaromatic Compounds", "2-Nitroaniline"),
    ("SVOCs", "Nitroaromatic Compounds", "2-Nitrophenol"),
    ("SVOCs", "Nitroaromatic Compounds", "2.6-Dinitrotoluene"),
    ("SVOCs", "Nitroaromatic Compounds", "3-Nitroaniline"),
    ("SVOCs", "Nitroaromatic Compounds", "4.6-Dinitro-2-methylphenol"),
    ("SVOCs", "Nitroaromatic Compounds", "4-Nitroaniline"),
    ("SVOCs", "Nitroaromatic Compounds", "4-Nitrophenol"),
    ("SVOCs", "Nitroaromatic Compounds", "Nitrobenzene"),
    # SVOCs - Chlorinated Hydrocarbons (3)
    ("SVOCs", "Chlorinated Hydrocarbons",
     "Bis(2-chloroethoxy)methane"),
    ("SVOCs", "Chlorinated Hydrocarbons",
     "Bis(2-chloroethyl)ether"),
    ("SVOCs", "Chlorinated Hydrocarbons",
     "Bis(2-chloroisopropyl)ether"),
    # SVOCs - Nitrosoamines (1)
    ("SVOCs", "Nitrosoamines", "N-Nitrosodi-n-propylamine"),
    # SVOCs - Pesticides (1)
    ("SVOCs", "Pesticides", "Dinoseb"),
    # SVOCs - Phthalates (6)
    ("SVOCs", "Phthalates", "Bis(2-ethylhexyl)phthalate"),
    ("SVOCs", "Phthalates", "Butyl benzyl phthalate"),
    ("SVOCs", "Phthalates", "Di-n-butyl phthalate"),
    ("SVOCs", "Phthalates", "Di-n-octyl phthalate"),
    ("SVOCs", "Phthalates", "Diethyl phthalate"),
    ("SVOCs", "Phthalates", "Dimethyl phthalate"),
]

# ── PFAS ALIAS + THRESHOLD HELPERS ──────────────────────────────────────────────
PFAS_ALIAS = {
    "2,3,3,3-tetrafluoro-2-(heptafluoropropoxy)propanoic acid (hfpo-da)":
        "hexafluoropropylene oxide dimer acid (hfpo-da)",
    "7h-perfluoroheptanoic acid (hpfhpa)":
        "perfluoroheptanoic acid (pfhpa)",
    "perfluorobutane sulfonic acid (pfbs)":
        "perfluorobutanesulfonic acid (pfbs)",
    "perfluorobutane sulfonate (pfbs)":
        "perfluorobutanesulfonic acid (pfbs)",
    "perfluorohexane sulfonic acid (pfhxs)":
        "perfluorohexanesulfonic acid (pfhxs)",
    "perfluorohexane sulfonate (pfhxs)":
        "perfluorohexanesulfonic acid (pfhxs)",
    "perfluorooctane sulfonic acid (pfos)":
        "perfluorooctanesulfonic acid (pfos)",
    "perfluorooctane sulfonate (pfos)":
        "perfluorooctanesulfonic acid (pfos)",
    "perfluorooctadecanoic acid (pfocda)":
        "perfluorooctadecanoic acid (pfoda)",
    "perfluoroundecanoic acid (pfunda)":
        "perfluoroundecanoic acid (pfuda)",
    "perfluorotetradecanoic acid (pfcpda)":
        "perfluorotetradecanoic acid (pfteta)",
    "perfluorodecane sulfonic acid (pfds)":
        "perfluorodecanesulfonic acid (pfds)",
    "perfluoroheptane sulfonic acid (pfhps)":
        "perfluoroheptanesulfonic acid (pfhps)",
    "perfluoropentane sulfonic acid (pfpes)":
        "perfluoropentanesulfonic acid (pfpes)",
    "perfluorooctane sulfonamide (fosa)":
        "perfluorooctanesulfonamide (fosa)",
    "perfluoropentanoic acid (pfpea)":
        "perfluoropentanoic acid (pfpea)",
    "perfluorodecanoic acid (pfda)":
        "perfluorodecanoic acid (pfda)",
    "perfluorododecanoic acid (pfdoda)":
        "perfluorododecanoic acid (pfdoda)",
    "perfluoroheptanoic acid (pfhpa)":
        "perfluoroheptanoic acid (pfhpa)",
    "perfluorotridecanoic acid (pftrda)":
        "perfluorotridecanoic acid (pftrda)",
    "perfluorooctanesulfonic acid (pfos)":
        "perfluorooctanesulfonic acid (pfos)",
}

CANONICAL_KEY = "__CANONICAL_MAP_INTERNAL__"


def canonical_compound(name: str) -> str:
    """
    צורה קנונית:
    - מאחד ethene/ethylene/ethen
    - מוחק ספרות וסימני . , - / מהאותיות, ומוחק גם רווחים → Propylbenzene == Propyl benzene
    - אוסף את כל המספרים לפי סדר הופעה → 1.2.4 / 1,2,4 / 1-2-4 → "1,2,4"
    לדוגמה:
    - "1.2.4-Trimethylbenzene" -> "1,2,4 trimethylbenzene"
    - "Trimethylbenzene, 1,2,4-" -> "1,2,4 trimethylbenzene"
    - "cis-1.2-Dichloroethene" -> "1,2 cisdichloroethen"
    - "1,2-cis-Dichloroethylene" -> "1,2 cisdichloroethen"
    """
    s = norm(name)

    # unify ethene/ethylene/ethen
    s = s.replace("ethylene", "ethen")
    s = s.replace("ethene", "ethen")

    # remove trailing brackets content
    s = re.sub(r"\s*\([^)]+\)\s*$", "", s)

    # digits in order
    nums = re.findall(r"\d+", s)
    num_part = ",".join(nums) if nums else ""

    # letters part: remove digits and punctuation, then remove spaces
    base_no_nums = re.sub(r"[0-9.,/\-]", " ", s)
    base_no_nums = re.sub(r"\s+", "", base_no_nums)  # no spaces at all
    word_part = base_no_nums

    canon = (num_part + " " + word_part).strip()
    return canon


def match_threshold(compound_name, thresh_dict):
    """
    Match compound name from ALS/VOC list to threshold table.
    מטפל:
    - נקודות מול פסיקים במספרים
    - רווחים בתוך שם (Propylbenzene / Propyl benzene)
    - cis/trans לפני/אחרי המספרים
    """
    key = norm(compound_name)
    canon_map = thresh_dict.get(CANONICAL_KEY)

    # 1. ניסיון ישיר + וריאציות נקודה/פסיק
    for k in (key, key.replace(".", ","), key.replace(",", ".")):
        if k in thresh_dict and k != CANONICAL_KEY:
            return thresh_dict[k]

    # 2. PFAS alias
    aliased = PFAS_ALIAS.get(key)
    if aliased:
        a_key = norm(aliased)
        for k in (a_key, a_key.replace(".", ","), a_key.replace(",", ".")):
            if k in thresh_dict and k != CANONICAL_KEY:
                return thresh_dict[k]

    # 3. הורדת סוגריים בסוף
    stripped = re.sub(r"\s*\([A-Z0-9:_\-]+\)\s*$", "",
                      compound_name).strip().lower()
    for k, v in thresh_dict.items():
        if k == CANONICAL_KEY:
            continue
        k_s = re.sub(r"\s*\([A-Z0-9:_\-]+\)\s*$", "", k).strip().lower()
        if len(stripped) > 8 and stripped == k_s:
            return v

    # 4. canonical
    if canon_map:
        ck = canonical_compound(compound_name)
        hit = canon_map.get(ck)
        if hit:
            return hit

    return {}


def load_threshold_file(file_bytes):
    wb = load_workbook(io.BytesIO(file_bytes), data_only=True)
    ws = wb.active
    thresh = {}
    for row in ws.iter_rows(min_row=6, values_only=True):
        if not row[0]:
            continue
        name = str(row[0]).strip()
        cas  = str(row[1]).strip() if row[1] else "-"
        def g(ci):
            return (row[ci] if ci < len(row)
                    and row[ci] is not None
                    and str(row[ci]) not in ("NA", "")
                    else None)
        thresh[norm(name)] = {
            "name":  name,
            "cas":   cas,
            "units": str(row[3]) if row[3] else "mg/kg",
            "VSL":       g(4),
            "Ind_A_06":  g(8),
            "Ind_A_6p":  g(9),
            "Ind_B":     g(10),
            "Res_A_06":  g(11),
            "Res_A_6p":  g(12),
            "Res_B":     g(13),
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
    if b:
        return "Ind_B" if ind else "Res_B"
    deep = ">6" in depth
    if ind:
        return "Ind_A_06" if not deep else "Ind_A_6p"
    else:
        return "Res_A_06" if not deep else "Res_A_6p"


def tier1_label(land_use, aquifer, depth):
    return f"TIER 1\n{land_use}\n{aquifer}\n{depth}"


def get_thresh(compound, thresh_dict, t1col):
    t = match_threshold(compound, thresh_dict)
    return t.get("VSL"), t.get(t1col), t.get("cas", "-")


def build_metals_thresh(thresh_dict, t1col):
    result = {}
    for key, v in thresh_dict.items():
        if key == CANONICAL_KEY:
            continue
        sym = THRESH_METAL_MAP.get(key)
        if sym and sym not in result:
            result[sym] = {
                "vsl":   v.get("VSL"),
                "tier1": v.get(t1col),
                "cas":   v.get("cas", "-"),
            }
    return result


# ── ALS PARSER ───────────────────────────────────────────────────────────────────
def parse_als_file(file_bytes, filename):
    try:
        wb = load_workbook(io.BytesIO(file_bytes), data_only=True)
    except Exception as e:
        return None, str(e)

    main = next(
        (wb[n] for n in wb.sheetnames
         if "Client" in n and "SOIL" in n),
        wb.worksheets[0]
    )
    rows = list(main.iter_rows(values_only=True))

    sid_idx = next(
        (i for i, r in enumerate(rows)
         if any("Client Sample ID" in str(v)
                for v in r if v)),
        None
    )
    if sid_idx is None:
        return None, "לא נמצאה שורת Sample IDs"

    col2sample = {
        ci: str(v).strip()
        for ci, v in enumerate(rows[sid_idx])
        if v and v != "Client Sample ID"
    }

    ph_idx = next(
        (i for i, r in enumerate(rows)
         if r and r[0] == "Parameter"),
        None
    )
    if ph_idx is None:
        return None, "לא נמצאה שורת Parameter"

    records = []
    group   = "Unknown"

    for row in rows[ph_idx + 1:]:
        p   = row[0] if len(row) > 0 else None
        u   = row[2] if len(row) > 2 else None
        lor = row[3] if len(row) > 3 else None

        if not p:
            continue
        if not u and not row[1]:
            group = str(p).strip()
            continue

        for ci, sname in col2sample.items():
            sid, depth_val = parse_sample(sname)
            if sid is None:
                continue
            val = row[ci] if ci < len(row) else None
            rs  = str(val).strip() if val is not None else ""
            result = None
            if rs.startswith("<"):
                result = 0.0
            elif rs and rs not in ("None", ""):
                try:
                    result = float(rs)
                except Exception:
                    result = None
            if result is not None:
                lor_val = None
                if rs.startswith("<"):
                    try:
                        lor_val = float(rs[1:].strip())
                    except Exception:
                        lor_val = 0.0
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


# ── TPH SHEET / METALS / PFAS / VOC+SVOC ────────────────────────────────────────
# (אותו קוד כמו בגרסאות הקודמות שלך – השארתי אותו כמו שהוא, כדי לא לגעת ב‑TPH / Metals / PFAS
#  ושיניתי רק את הלוגיקה של התאמת ערכי הסף. אם תרצה/י גם אותו כאן מלא שוב – תגיד/י.)

# בגלל אורך ההודעה אני לא מעתיק שוב את כל הגיליונות TPH/Metals/PFAS/VOC,
# אלא משאיר אותם כמו בקובץ האחרון שנתתי (v3.4) – רק הפונקציות
# canonical_compound / match_threshold / load_threshold_file עודכנו.

# אם את/ה רוצה, אוכל גם לצרף שוב את כל הקובץ כולל כל הפונקציות, אבל השינוי
# המהותי שנוגע לדוגמאות שציינת נמצא בדיוק בשלוש הפונקציות למעלה.
