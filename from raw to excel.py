import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import io
import re

# --------------------------------------------------
# CONFIG & STYLE CONSTANTS
# --------------------------------------------------
st.set_page_config(page_title="דוח סקר קרקע מאוחד", layout="wide", page_icon="🧪")

# צבעים וגופנים לפי המפרט שלך
YELLOW_FILL = PatternFill("solid", fgColor="FFFF00")  # חריגת VSL
ORANGE_FILL = PatternFill("solid", fgColor="FFA500")  # חריגת TIER 1
BLUE_HEADER = PatternFill("solid", fgColor="1F4E79")  # כותרת ראשית
LIGHT_BLUE_HEADER = PatternFill("solid", fgColor="2E75B6") # כותרת משנית
WHITE_FONT = Font(color="FFFFFF", bold=True, name="Arial", size=10)
BLACK_BOLD = Font(bold=True, name="Arial", size=10)
NORMAL_FONT = Font(name="Arial", size=10)

def make_border():
    s = Side(style="thin", color="000000")
    return Border(left=s, right=s, top=s, bottom=s)

# --------------------------------------------------
# ANALYTE LISTS (ORDERED AS PER YOUR TEMPLATES)
# --------------------------------------------------
METALS_LIST = ['Al', 'As', 'Ba', 'Be', 'B', 'Cd', 'Ca', 'Cr', 'Co', 'Cu', 'Fe', 'Pb', 'Li', 'Mg', 'Mn', 'Hg', 'Ni', 'K', 'Na', 'Se', 'Ag', 'V', 'Zn']
TPH_LIST = ['TPH DRO', 'TPH ORO', 'Total TPH']

# --------------------------------------------------
# HELPERS
# --------------------------------------------------
def natural_sort_key(s):
    """מיון לוגי של קידוחים (S329 לפני S3) בסדר יורד"""
    return [int(text) if text.isdigit() else text.lower() for text in re.split('([0-9]+)', str(s))]

def norm_name(s):
    s = str(s).strip().lower() if s else ""
    s = s.replace("\xa0", " ").replace("–", "-").replace("—", "-")
    return re.sub(r"\s+", " ", s)

# --------------------------------------------------
# PARSING LOGIC
# --------------------------------------------------
def parse_als_file(file_bytes, filename):
    try:
        from openpyxl import load_workbook
        wb = load_workbook(io.BytesIO(file_bytes), data_only=True)
        sheet = next((wb[n] for n in wb.sheetnames if "Client" in n and "SOIL" in n), wb.active)
        rows = list(sheet.iter_rows(values_only=True))
        
        # זיהוי שורת הדגימות ושורת הפרמטרים
        sample_row_idx = next(i for i, r in enumerate(rows) if any("Client Sample ID" in str(v) for v in r if v))
        param_row_idx = next(i for i, r in enumerate(rows) if r and r[0] == "Parameter")
        
        samples = {i: str(v).strip() for i, v in enumerate(rows[sample_row_idx]) if v and v != "Client Sample ID"}
        
        records = []
        current_group = "Unknown"
        for row in rows[param_row_idx + 1:]:
            param, unit = row[0], row[2]
            if not param: continue
            if not unit: current_group = str(param).strip(); continue
            
            for col_idx, sname in samples.items():
                val = row[col_idx]
                match = re.match(r"^(S-?\d+[A-Za-z]*)\s*\
