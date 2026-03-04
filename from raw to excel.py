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
        if len(full) > 12 and (full in strip_abbrev(k) or strip
