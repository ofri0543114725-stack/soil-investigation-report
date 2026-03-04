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
st.caption("v3.3 - Canonical threshold matching (numbers before/after name)")
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
    - מאחד ethene / ethylene / ethen.
    - מאחד מצבים שבהם המספרים לפני/אחרי השם (1.2.4‑Trimethylbenzene
      לעומת Trimethylbenzene, 1,2,4‑).
    """
    s = norm(name)

    # unify ethene/ethylene/ethen
    s = s.replace("ethylene", "ethen")
    s = s.replace("ethene", "ethen")

    # הורדת סוגריים בסוף
    s = re.sub(r"\s*\([^)]+\)\s*$", "", s)

    # digits sequence (אוספים את כל המספרים לפי סדר הופעה)
    nums = re.findall(r"\d+", s)
    num_part = ",".join(nums) if nums else ""

    # מסירים ספרות וסימני הפרדה מהשם ובלבד שנשארו רק מילים
    base_no_nums = re.sub(r"[0-9.,/\-]", " ", s)
    base_no_nums = re.sub(r"\s+", " ", base_no_nums).strip()
    word_part = base_no_nums

    canon = (num_part + " " + word_part).strip()
    return canon


def match_threshold(compound_name, thresh_dict):
    """
    Match compound name from ALS/VOC list to threshold table.
    כולל:
    - נקודות/פסיקים
    - alias של PFAS
    - הסרת סוגריים
    - התאמת צורה קנונית (מספרים+שם)
    """
    key = norm(compound_name)
    canon_map = thresh_dict.get(CANONICAL_KEY)

    # 1. ניסיון ישיר + וריאציות נקודה/פסיק
    candidates = [key, key.replace(".", ","), key.replace(",", ".")]
    for k in candidates:
        if k in thresh_dict and k != CANONICAL_KEY:
            return thresh_dict[k]

    # 2. PFAS alias
    aliased = PFAS_ALIAS.get(key)
    if aliased:
        a_key = norm(aliased)
        for k in [a_key, a_key.replace(".", ","), a_key.replace(",", ".")]:
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

    # 4. canonical name
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

    # canonical map
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
    """Build {symbol: {vsl,tier1,cas}} using threshold file."""
    result = {}
    for key, v in thresh_dict.items():
        if key == CANONICAL_KEY:
            continue
        sym = THRESH_METAL_MAP.get(key)
        if sym and sym not in result:   # first match wins
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
        m   = row[1] if len(row) > 1 else None
        u   = row[2] if len(row) > 2 else None
        lor = row[3] if len(row) > 3 else None

        if not p:
            continue
        if not m and not u:
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


# ── TPH SHEET ────────────────────────────────────────────────────────────────────
def write_tph_sheet(ws, df, thresh_dict, t1col, t1lbl):
    def is_dro(c):
        if "(dro)" in c or "- dro" in c or c.strip() == "dro":
            return True
        if "(oro)" in c or "- oro" in c:
            return False
        if "c10" in c and "c28" in c and "c40" not in c:
            return True
        return False

    def is_oro(c):
        if "(oro)" in c or "- oro" in c or c.strip() == "oro":
            return True
        if "(dro)" in c or "- dro" in c:
            return False
        if "c24" in c and "c40" in c:
            return True
        if "c28" in c and "c40" in c:
            return True
        return False

    def is_total(c):
        if any(x in c for x in ["(dro)", "(oro)", "- dro", "- oro"]):
            return False
        if "c10" in c and "c40" in c:
            return True
        if "total" in c and ("tph" in c or "hydrocarbon" in c):
            return True
        return False

    vsl_d, t1_d, _ = get_thresh("C10 - C28 Fraction (DRO)", thresh_dict, t1col)
    vsl_o, t1_o, _ = get_thresh("C24 - C40 Fraction (ORO)", thresh_dict, t1col)
    vsl_t, t1_t, _ = get_thresh("TPH - DRO + ORO (Tier 1)", thresh_dict, t1col)

    vv = [v for v in [vsl_d, vsl_o, vsl_t] if v]
    tt = [v for v in [t1_d, t1_o, t1_t] if v]
    vsl_tot = min(vv) if vv else 350
    t1_tot  = min(tt) if tt else 350

    for ci, h in enumerate(["שם קידוח", "עומק",
                            "TPH DRO", "TPH ORO", "Total TPH"], 1):
        style_hdr(ws.cell(1, ci, h), HDR_BLUE_FILL)

    ws.merge_cells(start_row=1, start_column=1, end_row=5, end_column=1)
    c = ws.cell(1, 1, "שם קידוח")
    c.font = Font(bold=True, name="Arial", size=11)
    c.fill = HDR_BLUE_FILL
    c.alignment = Alignment(horizontal="center", vertical="center",
                            wrap_text=True)
    c.border = thin_border()

    sub_rows = ["יחידות", "CAS", "VSL", t1lbl]
    sub_vals = {
        "יחידות": "mg/kg",
        "CAS":     "C10-C40",
        "VSL":     vsl_tot,
        t1lbl:     t1_tot,
    }
    for ri, lbl in enumerate(sub_rows, 2):
        style_hdr(ws.cell(ri, 2, lbl), HDR_BLUE_FILL)
        for ci in [3, 4, 5]:
            style_hdr(ws.cell(ri, ci, sub_vals[lbl]), HDR_BLUE_FILL)

    pivoted = {}
    for _, r in df.iterrows():
        k = (r["sample_id"], r["depth"])
        if k not in pivoted:
            pivoted[k] = {
                "DRO": "", "ORO": "", "TOT": "",
                "DRO_f": None, "ORO_f": None,
                "DRO_lor": None, "ORO_lor": None,
            }
        c = r["compound_lower"]
        if is_dro(c) and not pivoted[k]["DRO"]:
            pivoted[k]["DRO"]      = r["result_str"]
            pivoted[k]["DRO_f"]    = r["result"]
            pivoted[k]["DRO_lor"]  = r.get("lor_val")
        elif is_oro(c) and not pivoted[k]["ORO"]:
            pivoted[k]["ORO"]      = r["result_str"]
            pivoted[k]["ORO_f"]    = r["result"]
            pivoted[k]["ORO_lor"]  = r.get("lor_val")
        elif is_total(c) and not pivoted[k]["TOT"]:
            pivoted[k]["TOT"]      = r["result_str"]

    ri       = 6
    prev_sid = None
    sid_rows = {}

    for (sid, depth_val), v in sorted(
        pivoted.items(),
        key=lambda x: (sort_key(x[0][0]), x[0][1] or 0)
    ):
        if v["TOT"]:
            total_s = v["TOT"]
        else:
            dro_lor    = v["DRO"] and str(v["DRO"]).startswith("<")
            oro_lor    = v["ORO"] and str(v["ORO"]).startswith("<")
            dro_empty  = not v["DRO"]
            oro_empty  = not v["ORO"]
            dro_num    = (v["DRO_lor"] if dro_lor and v["DRO_lor"] is not None
                          else (v["DRO_f"] or 0))
            oro_num    = (v["ORO_lor"] if oro_lor and v["ORO_lor"] is not None
                          else (v["ORO_f"] or 0))
            total_f    = dro_num + oro_num
            if ((dro_lor or dro_empty) and
                (oro_lor or oro_empty) and
                not (dro_empty and oro_empty)):
                total_s = f"<{total_f:.0f}"
            else:
                total_s = f"{total_f:.0f}"

        hl_dro   = check_exceed(v["DRO"],  vsl_tot, t1_tot)
        hl_oro   = check_exceed(v["ORO"],  vsl_tot, t1_tot)
        hl_total = check_exceed(total_s,   vsl_tot, t1_tot)

        if sid != prev_sid:
            sid_rows[sid] = []
        sid_rows[sid].append(ri)
        sid_val = sid if sid != prev_sid else None
        prev_sid = sid

        style_data(ws.cell(ri, 1, sid_val))
        style_data(ws.cell(ri, 2, depth_val))
        style_data(ws.cell(ri, 3, v["DRO"]),  hl_dro)
        style_data(ws.cell(ri, 4, v["ORO"]),  hl_oro)
        style_data(ws.cell(ri, 5, total_s),   hl_total)
        ri += 1

    apply_sid_merge(ws, sid_rows, col=1)

    for col, w in zip("ABCDE", [14, 10, 16, 16, 16]):
        ws.column_dimensions[col].width = w
    ws.row_dimensions[1].height = 15
    ws.freeze_panes = "A6"


# ── METALS SHEET ─────────────────────────────────────────────────────────────────
def write_metals_sheet(ws, df, thresh_dict, t1col, t1lbl):
    df = df.copy()
    df["sym"] = df["compound_lower"].map(METAL_MAP)
    df = df[df["sym"].notna()]
    if df.empty:
        ws.cell(1, 1, "אין נתוני מתכות")
        return

    present = set(df["sym"].unique())
    metals  = [m for m in METALS_ORDER if m in present] + \
              sorted(present - set(METALS_ORDER))
    mt = build_metals_thresh(thresh_dict, t1col)

    ws.merge_cells(start_row=1, start_column=1, end_row=5, end_column=1)
    c = ws.cell(1, 1, "שם קידוח")
    c.font = Font(bold=True, name="Arial", size=11)
    c.fill = HDR_BLUE_FILL
    c.alignment = Alignment(horizontal="center", vertical="center",
                            wrap_text=True)
    c.border = thin_border()

    for ci, h in enumerate(["עומק"] + metals, 2):
        style_hdr(ws.cell(1, ci, h), HDR_BLUE_FILL)

    for ri, lbl in enumerate(["יחידות", "CAS", "VSL", t1lbl], 2):
        style_hdr(ws.cell(ri, 2, lbl), HDR_BLUE_FILL)
        for ci, sym in enumerate(metals, 3):
            t = mt.get(sym, {})
            val = {
                "יחידות": "mg/kg",
                "CAS":     t.get("cas", "-"),
                "VSL":     t.get("vsl", "-"),
                t1lbl:     t.get("tier1", "-"),
            }.get(lbl, "-")
            style_hdr(ws.cell(ri, ci, val), HDR_BLUE_FILL)

    pt = df.pivot_table(
        index=["sample_id", "depth"],
        columns="sym",
        values="result_str",
        aggfunc="first"
    )
    pt = pt.reindex(
        sorted(pt.index, key=lambda x: (sort_key(x[0]), x[1] or 0))
    )

    ri       = 6
    prev_sid = None
    sid_rows = {}

    for (sid, depth_val), row_data in pt.iterrows():
        if sid != prev_sid:
            sid_rows[sid] = []
        sid_rows[sid].append(ri)
        sid_val = sid if sid != prev_sid else None
        prev_sid = sid

        style_data(ws.cell(ri, 1, sid_val))
        style_data(ws.cell(ri, 2, depth_val))
        for ci, sym in enumerate(metals, 3):
            val = row_data.get(sym, "") or ""
            val = "" if str(val) == "nan" else str(val)
            hl  = check_exceed(
                val,
                mt.get(sym, {}).get("vsl"),
                mt.get(sym, {}).get("tier1")
            )
            style_data(ws.cell(ri, ci, val), hl)
        ri += 1

    apply_sid_merge(ws, sid_rows, col=1)
    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 10
    for ci in range(3, len(metals) + 3):
        ws.column_dimensions[get_column_letter(ci)].width = 11
    ws.freeze_panes = "C6"


# ── PFAS SHEET ───────────────────────────────────────────────────────────────────
def write_pfas_sheet(ws, df, thresh_dict, t1col, t1lbl):
    df = df.copy()

    pairs_pfas = sorted(
        df[["sample_id", "depth"]].drop_duplicates().values.tolist(),
        key=lambda x: (sort_key(x[0]), -(x[1] or 0))
    )

    def to_ug(v):
        if v is None:
            return None
        try:
            return round(float(v) * 1000, 6)
        except Exception:
            return v

    fixed_hdrs = [
        "שם התרכובת", "CAS",
        "VSL [µg/kg]", f"{t1lbl} [µg/kg]",
        "יחידות",
    ]
    for ci, h in enumerate(fixed_hdrs, 1):
        style_hdr(ws.cell(1, ci, h), HDR_BLUE_FILL)
        ws.merge_cells(start_row=1, start_column=ci,
                       end_row=2, end_column=ci)
        ws.cell(1, ci).alignment = Alignment(horizontal="center",
                                             vertical="center",
                                             wrap_text=True)
        ws.cell(1, ci).fill   = HDR_BLUE_FILL
        ws.cell(1, ci).border = thin_border()

    style_hdr(ws.cell(1, 6, "LOR"),     HDR_BLUE_FILL)
    style_hdr(ws.cell(2, 6, "[µg/kg]"), HDR_BLUE_FILL)
    style_hdr(ws.cell(1, 7, "שם קידוח"), HDR_BLUE_FILL)
    style_hdr(ws.cell(2, 7, "עומק"),     HDR_BLUE_FILL)

    prev_sid          = None
    sid_merge_start_p = {}
    for ci, (sid, depth_val) in enumerate(pairs_pfas, 8):
        sid_val = sid if sid != prev_sid else None
        style_hdr(ws.cell(1, ci, sid_val), HDR_BLUE_FILL)
        style_hdr(ws.cell(2, ci, depth_val),   HDR_BLUE_FILL)
        if sid != prev_sid:
            sid_merge_start_p[sid] = ci
        prev_sid = sid

    for sid, start_ci in sid_merge_start_p.items():
        cols_p = [ci for ci, (s, _) in enumerate(pairs_pfas, 8)
                  if s == sid]
        if len(cols_p) > 1:
            ws.merge_cells(start_row=1, start_column=start_ci,
                           end_row=1, end_column=cols_p[-1])
            c = ws.cell(1, start_ci)
            c.alignment = Alignment(horizontal="center",
                                    vertical="center")
            c.fill   = HDR_BLUE_FILL
            c.border = thin_border()

    for row_i, cmp in enumerate(df["compound"].unique(), 3):
        df_c = df[df["compound"] == cmp]
        vsl_mg, tier1_mg, cas = get_thresh(cmp, thresh_dict, t1col)
        vsl    = to_ug(vsl_mg)
        tier1  = to_ug(tier1_mg)
        unit   = df_c.iloc[0]["unit"] if not df_c.empty else "µg/kg"
        lor    = df_c.iloc[0]["lor"]  if not df_c.empty else ""

        for ci, val in enumerate([cmp, cas, vsl, tier1, unit], 1):
            style_data(ws.cell(row_i, ci, val))

        ws.merge_cells(start_row=row_i, start_column=6,
                       end_row=row_i, end_column=6)
        style_data(ws.cell(row_i, 6, lor))
        style_data(ws.cell(row_i, 7, None))

        for ci, (sid, depth_val) in enumerate(pairs_pfas, 8):
            sub = df_c[(df_c["sample_id"] == sid) &
                       (df_c["depth"] == depth_val)]
            rs  = sub.iloc[0]["result_str"] if not sub.empty else ""
            style_data(ws.cell(row_i, ci, rs),
                       check_exceed(rs, vsl, tier1))

    ws.column_dimensions["A"].width = 50
    for ci in range(2, 8):
        ws.column_dimensions[get_column_letter(ci)].width = 13
    for ci in range(8, 8 + len(pairs_pfas)):
        ws.column_dimensions[get_column_letter(ci)].width = 12
    ws.freeze_panes = "H3"


# ── VOC+SVOC SHEET ───────────────────────────────────────────────────────────────
def write_voc_sheet(ws, df, thresh_dict, t1col, t1lbl):
    df = df.copy()

    pairs = sorted(
        df[["sample_id", "depth"]].drop_duplicates().values.tolist(),
        key=lambda x: (sort_key(x[0]), -(x[1] or 0))
    )

    for ci, h in enumerate(
        ["קבוצה", "קבוצה", "שם התרכובת", "CAS", "VSL", f"TIER 1\n{depth}"], 1
    ):
        ws.merge_cells(start_row=1, start_column=ci,
                       end_row=2, end_column=ci)
        style_hdr(ws.cell(1, ci, h), HDR_BLUE_FILL, sz=9)
        ws.cell(1, ci).border = thin_border()

    ws.merge_cells(start_row=1, start_column=7, end_row=2, end_column=8)
    style_hdr(ws.cell(1, 7, "יחידות"), HDR_BLUE_FILL, sz=9)
    ws.cell(1, 7).border = thin_border()

    style_hdr(ws.cell(1, 9, "שם קידוח"), HDR_BLUE_FILL, sz=9)
    style_hdr(ws.cell(2, 9, "עומק"),      HDR_BLUE_FILL, sz=9)

    prev_sid      = None
    sid_col_start = {}
    for ci, (sid, depth_val) in enumerate(pairs, 10):
        sid_val = sid if sid != prev_sid else None
        style_hdr(ws.cell(1, ci, sid_val), HDR_BLUE_FILL, sz=9)
        style_hdr(ws.cell(2, ci, depth_val),   HDR_BLUE_FILL, sz=9)
        if sid != prev_sid:
            sid_col_start[sid] = ci
        prev_sid = sid

    for sid, sc in sid_col_start.items():
        cols = [ci for ci, (s, _) in enumerate(pairs, 10) if s == sid]
        if len(cols) > 1:
            ws.merge_cells(start_row=1, start_column=sc,
                           end_row=1, end_column=cols[-1])
            c = ws.cell(1, sc)
            c.alignment = Alignment(horizontal="center",
                                    vertical="center")
            c.fill   = HDR_BLUE_FILL
            c.border = thin_border()

    als_data = {}
    for _, r in df.iterrows():
        k = norm(r["compound"])
        als_data.setdefault(k, {})[(r["sample_id"], r["depth"])] = r["result_str"]

    for k in list(als_data.keys()):
        for alt in (k.replace(".", ","), k.replace(",", ".")):
            if alt not in als_data:
                als_data[alt] = als_data[k]

    for row_i, (vs, grp, cmp) in enumerate(VOC_COMPOUND_ORDER, 3):
        vsl, tier1, cas = get_thresh(cmp, thresh_dict, t1col)
        cmp_key = norm(cmp)
        cmp_data = als_data.get(cmp_key) or \
                   als_data.get(cmp_key.replace(".", ",")) or {}

        style_data(ws.cell(row_i, 1, vs),    sz=9)
        style_data(ws.cell(row_i, 2, grp),   sz=9)
        style_data(ws.cell(row_i, 3, cmp),   sz=9)
        style_data(ws.cell(row_i, 4, cas),   sz=9)
        style_data(ws.cell(row_i, 5, vsl),   sz=9)
        style_data(ws.cell(row_i, 6, tier1), sz=9)

        ws.merge_cells(start_row=row_i, start_column=7,
                       end_row=row_i, end_column=8)
        style_data(ws.cell(row_i, 7, "mg/kg"), sz=9)

        style_data(ws.cell(row_i, 9, None), sz=9)

        for ci, (sid, depth_val) in enumerate(pairs, 10):
            rs = cmp_data.get((sid, depth_val), "")
            style_data(ws.cell(row_i, ci, rs),
                       check_exceed(rs, vsl, tier1),
                       sz=9)

    for r1, r2, val in [(3, 32, "VOCs"), (33, 96, "SVOCs")]:
        ws.merge_cells(start_row=r1, start_column=1,
                       end_row=r2, end_column=1)
        c = ws.cell(r1, 1, val)
        c.font = Font(bold=True, name="Arial", size=9)
        c.alignment = Alignment(horizontal="center",
                                vertical="center",
                                wrap_text=True)
        c.border = thin_border()

    b_ranges = [
        (3, 12,  "Non-Halogenated VOCs"),
        (13, 16, "BTEX"),
        (17, 32, "Halogenated VOCs"),
        (33, 37, "Phenols & Naphtols"),
        (38, 53, "PAHs"),
        (54, 57, "Anilines"),
        (58, 65, "Aromatic Compounds"),
        (66, 66, "Alcohols"),
        (67, 69, "Aldehydes / Ketones"),
        (70, 75, "Chlorophenols"),
        (76, 85, "Nitroaromatic Compounds"),
        (86, 88, "Chlorinated Hydrocarbons"),
        (89, 89, "Nitrosoamines"),
        (90, 90, "Pesticides"),
        (91, 96, "Phthalates"),
    ]
    for r1, r2, val in b_ranges:
        if r2 > r1:
            ws.merge_cells(start_row=r1, start_column=2,
                           end_row=r2, end_column=2)
        c = ws.cell(r1, 2, val)
        c.font = Font(bold=True, name="Arial", size=9)
        c.alignment = Alignment(horizontal="center",
                                vertical="center",
                                wrap_text=True)
        c.border = thin_border()

    ws.column_dimensions["A"].width = 8
    ws.column_dimensions["B"].width = 22
    ws.column_dimensions["C"].width = 35
    ws.column_dimensions["D"].width = 12
    ws.column_dimensions["E"].width = 10
    ws.column_dimensions["F"].width = 12
    ws.column_dimensions["G"].width = 7
    ws.column_dimensions["H"].width = 7
    ws.column_dimensions["I"].width = 12
    for ci in range(10, 10 + len(pairs)):
        ws.column_dimensions[get_column_letter(ci)].width = 10
    ws.row_dimensions[1].height = 20
    ws.row_dimensions[2].height = 15
    ws.freeze_panes = "J3"


# ── SIDEBAR & MAIN FLOW ─────────────────────────────────────────────────────────
st.sidebar.header("⚙️ הגדרות ערכי סף")
st.sidebar.markdown("🟡 חריגה מ-VSL &nbsp;&nbsp;&nbsp; 🟠 חריגה מ-TIER 1")
st.sidebar.markdown("---")

land_use = st.sidebar.selectbox("Land Use",
                                ["Industrial", "Residential"], index=0)
aquifer = st.sidebar.selectbox("Aquifer Sensitivity",
                               ["A-1, A, B", "B-1 or C"], index=0)
depth_opts = ["Not Applicable"] if "b-1" in aquifer.lower() \
    else ["0 - 6 m", ">6 m"]
depth   = st.sidebar.selectbox("Depth to Groundwater",
                               depth_opts, index=0)

t1col = get_tier1_col(land_use, aquifer, depth)
t1lbl = tier1_label(land_use, aquifer, depth)
st.sidebar.info(f"TIER 1: **{land_use}** | {aquifer} | {depth}")

c1, c2 = st.columns(2)
with c1:
    st.subheader("📁 קובץ ערכי סף")
    thr_file = st.file_uploader("העלה קובץ ערכי הסף המאוחד",
                                type=["xlsx", "xls"], key="thr")
with c2:
    st.subheader("📂 קבצי ALS")
    data_files = st.file_uploader("העלה קבצי ALS",
                                  type=["xlsx", "xls"],
                                  accept_multiple_files=True,
                                  key="data")

if not thr_file:
    st.info("👆 העלה קובץ ערכי סף וקבצי ALS")
    st.stop()

if not data_files:
    st.warning("⚠️ העלה קבצי ALS")
    st.stop()

thresh_dict = load_threshold_file(thr_file.read())
st.success(
    f"✅ {len(thresh_dict) - 1} ערכי סף (לא כולל מפת קאנון) | {land_use} | {aquifer} | {depth}"
)

all_data = []
for f in data_files:
    df, err = parse_als_file(f.read(), f.name)
    if err:
        st.warning(f"⚠️ {f.name}: {err}")
    else:
        all_data.append(df)
        st.success(f"✅ {f.name} — {len(df)} תוצאות")

if not all_data:
    st.error("לא נטענו נתונים.")
    st.stop()

df_all = pd.concat(all_data, ignore_index=True)

with st.expander(f"👁️ תצוגה מקדימה ({len(df_all)} שורות)"):
    st.dataframe(df_all.head(30), use_container_width=True)
with st.expander("קבוצות שנמצאו"):
    st.write(df_all["group"].unique().tolist())

def dg(kw):
    return df_all[df_all["group"].str.contains(
        "|".join(kw),
        case=False,
        na=False
    )]

tph_df    = dg(["petroleum", "tph", "hydrocarbon"])
metals_df = dg(["metal", "cation", "extractable"])
pfas_df   = dg(["perfluor", "pfas", "fluorin"])
voc_df    = dg([
    "voc", "svoc", "btex", "aromatic", "halogenated", "volatile",
    "alcohol", "aldehyde", "ketone", "phenol", "pah", "aniline",
    "nitro", "phthalate", "pesticide", "pcb", "other"
])

if not tph_df.empty:
    st.info(f"✅ TPH: {tph_df['sample_id'].nunique()} קידוחים")
if not metals_df.empty:
    st.info(f"✅ Metals: {metals_df['sample_id'].nunique()} קידוחים")
if not voc_df.empty:
    st.info(f"✅ VOC+SVOC: {voc_df['sample_id'].nunique()} קידוחים")
if not pfas_df.empty:
    st.info(f"✅ PFAS: {pfas_df['sample_id'].nunique()} קידוחים")

wb_out = Workbook()
wb_out.remove(wb_out.active)

if not tph_df.empty:
    write_tph_sheet(wb_out.create_sheet("TPH"),
                    tph_df, thresh_dict, t1col, t1lbl)
if not metals_df.empty:
    write_metals_sheet(wb_out.create_sheet("Metals"),
                       metals_df, thresh_dict, t1col, t1lbl)
if not voc_df.empty:
    write_voc_sheet(wb_out.create_sheet("VOC+SVOC"),
                    voc_df, thresh_dict, t1col, t1lbl)
if not pfas_df.empty:
    write_pfas_sheet(wb_out.create_sheet("PFAS"),
                     pfas_df, thresh_dict, t1col, t1lbl)

if not wb_out.sheetnames:
    wb_out.create_sheet("Results")
    st.warning("לא זוהו קבוצות")

st.markdown("---")
buf = io.BytesIO()
wb_out.save(buf)
buf.seek(0)

st.download_button(
    "⬇️ הורד קובץ Excel מעובד",
    data=buf,
    file_name="soil_report.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=True,
)
