import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
import io
import re

# הגדרות עיצוב (צבעים וגופנים)
YELLOW_FILL = PatternFill("solid", fgColor="FFFF00") # חריגת VSL
ORANGE_FILL = PatternFill("solid", fgColor="FFA500") # חריגת TIER 1
BLUE_HEADER = PatternFill("solid", fgColor="1F4E79") # כותרת ראשית
LIGHT_BLUE_HEADER = PatternFill("solid", fgColor="2E75B6") # כותרת משנית
WHITE_FONT = Font(color="FFFFFF", bold=True, name="Arial", size=10)

def make_border():
    s = Side(style="thin", color="000000")
    return Border(left=s, right=s, top=s, bottom=s)

def natural_sort_key(s):
    """מיון לוגי (S329 לפני S3)"""
    return [int(text) if text.isdigit() else text.lower() for text in re.split('([0-9]+)', str(s))]

def norm_name(s):
    return re.sub(r"\s+", " ", str(s).strip().lower()) if s else ""

def parse_als_file(file_bytes):
    from openpyxl import load_workbook
    wb = load_workbook(io.BytesIO(file_bytes), data_only=True)
    sheet = next((wb[n] for n in wb.sheetnames if "Client" in n and "SOIL" in n), wb.active)
    rows = list(sheet.iter_rows(values_only=True))
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
            match = re.match(r"^(S-?\d+[A-Za-z]*)\s*\(([0-9.]+)\)", sname)
            sid,
