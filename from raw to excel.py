import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
import io
import re

# חובה: הפקודה הראשונה של Streamlit
st.set_page_config(page_title="מערכת עיבוד תוצאות קרקע", layout="wide")

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
    """מיון לוגי של קידוחים (S329 לפני S3)"""
    return [int(text) if text.isdigit() else text.lower() for text in re.split('([0-9]+)', str(s))]

def norm_name(s):
    return re.sub(r"\s+", " ", str(s).strip().lower()) if s else ""

def parse_als_data(file_bytes):
    """סורק את כל הגיליונות ב-ALS ושולף נתונים ללא סינון מוקדם"""
    all_records = []
    try:
        from openpyxl import load_workbook
        wb = load_workbook(io.BytesIO(file_bytes), data_only=True)
        for sheet in wb.worksheets:
            rows = list(sheet.iter_rows(values_only=True))
            # חיפוש שורת הקידוחים ושורת הפרמטרים
            s_idx = next((i for i, r in enumerate(rows) if any("Client Sample ID" in str(v) for v in r if v)), None)
            p_idx = next((i for i, r in enumerate(rows) if r and r[0] == "Parameter"), None)
            
            if s_idx is not None and p_idx is not None:
                samples = {idx: str(v).strip() for idx, v in enumerate(rows[s_idx]) if v and v != "Client Sample ID"}
                for row in rows[p_idx + 1:]:
                    param, unit = row[0], row[2]
                    if
