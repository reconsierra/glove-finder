
# glove-app.py
import os
from typing import Dict, List, Optional

import pandas as pd
import streamlit as st

# ---------------------------------------------------------
# App config
# ---------------------------------------------------------
st.set_page_config(page_title="Glove Finder", page_icon="🧤", layout="wide")
st.title("Glove Finder")
st.caption("Find the right glove by cut level, colour, and safety attributes.")

# ---------------------------------------------------------
# Creator-controlled layout + display labels
# ---------------------------------------------------------
ORDER_LEFT = ["Colour", "Cut Category", "Cut rating"]
ORDER_RIGHT = ["Food Safe?", "Chemical rated?", "Heat rated?"]

DISPLAY_LABELS = {
    "Colour": "Glove Colour",
    "Cut Category": "Cut Category (A–F)",
    "Cut rating": "Cut Level",
    "Food Safe?": "Food Safe",
    "Chemical rated?": "Chemical Resistant",
    "Heat rated?": "Heat Resistant",
}

DISPLAY_ATTRS = {
    "Article Numbers": "Article #",
    "Colour": "Glove Colour",
    "EN 388 Code": "EN388 Rating",
    "Abrasion": "Abrasion Level",
    "Cut": "Cut Level",
    "Tear": "Tear Strength",
    "Puncture": "Puncture Resistance",
    "Cut Category": "Cut Category",
    "Impact": "Impact Protection",
    "Chemical Resistance": "Chemical Resistant",
    "Heat Resistance": "Heat Resistant",
    "Food Safe": "Food Safe",
    "Tactile": "Tactile Sensitivity",
}

PLACEHOLDERS = {"", "-", "n/a", "na", "none", "null", "not en 388 rated"}


# ---------------------------------------------------------
# Helpers
# ---------------------------------------------------------
def norm(v) -> str:
    return "" if pd.isna(v) else str(v).strip().casefold()

def display(v) -> str:
    return "" if pd.isna(v) else str(v).strip()

def int_str(v: str) -> str:
    try:
        f = float(v)
        return str(int(f)) if f.is_integer() else v
    except Exception:
        return v


def resolve_column(aliases: List[str], columns: Dict[str, str]) -> Optional[str]:
    """Return the actual column name matching any alias"""
    for a in aliases:
        if a in columns:
            return columns[a]
    return None


# ---------------------------------------------------------
# Load data
# ---------------------------------------------------------
@st.cache_data
def load_data(xlsx: str) -> pd.DataFrame:
    df = pd.read_excel(xlsx, engine="openpyxl")
    df.columns = [c.strip() for c in df.columns]

    cols = {c.lower(): c for c in df.columns}

    colour_col = resolve_column(["colour", "color"], cols)
    cutcat_col = resolve_column(["cut category"], cols)
    cut_col = resolve_column(["cut"], cols)

    if not all([colour_col, cutcat_col, cut_col]):
        st.error("One or more required columns are missing.")
        st.stop()

    # Normalised helper columns for filtering
    df["_colour"] = df[colour_col].astype(str).str.strip().str.casefold()
    df["_cutcat"] = df[cutcat_col].astype(str).str.strip().str.casefold()
    df["_cut_display"] = df[cut_col].astype(str).map(int_str)
    df["_cut"] = df["_cut_display"].str.casefold()

    # Boolean helpers
    for col in ["Food Safe", "Chemical Resistance", "Heat Resistance"]:
        if col in df.columns:
            df[col + " (bool)"] = (
                df[col].astype(str).str.lower().isin(["yes", "y", "true", "1"])
            )
