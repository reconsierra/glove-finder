
# app.py
import os
from typing import Dict, List, Optional

import pandas as pd
import streamlit as st

st.set_page_config(page_title="Glove Finder", page_icon="🧤", layout="wide")
st.title("Glove Finder")
st.caption("Find the right glove by cut level/category, colour and safety attributes.")

# =========================================================
# 🔧 DISPLAY ORDER — adjust here (app creator only)
# Left and right columns render top-to-bottom in this order.
# Valid labels: "Colour", "Cut Category", "Cut rating",
#               "Food Safe?", "Chemical rated?", "Heat rated?"
ORDER_LEFT  = ["Cut Category - (new style letter type)", "Cut rating (old style number type)", "Colour (predominant colour only)"]
ORDER_RIGHT = ["Food Safe?", "Chemical rated?", "Heat rated?"]
# =========================================================


# -----------------------------
# Embedded image extraction
# -----------------------------
@st.cache_data(show_spinner=False)
def extract_embedded_images(xlsx_path: str, images_dir: str = "images") -> Dict[int, str]:
    """
    Extract embedded images from the first worksheet and map Excel row -> saved image path.
    Returns: {excel_row_number: image_file_path}
    """
    mapping: Dict[int, str] = {}
    try:
        from openpyxl import load_workbook
        from openpyxl.utils.cell import coordinate_to_tuple
    except Exception:
        return mapping

    try:
        wb = load_workbook(xlsx_path, data_only=True)
        ws = wb.active
        os.makedirs(images_dir, exist_ok=True)
        imgs = getattr(ws, "_images", [])
        for idx, img in enumerate(imgs, start=1):
            anchor = getattr(img, "anchor", None)
            row = None
            try:
                if isinstance(anchor, str):
                    row = coordinate_to_tuple(anchor)[0]  # (row, col)
                else:
                    from_anchor = getattr(anchor, "_from", None)
                    if from_anchor is not None:
                        row = int(from_anchor.row) + 1
            except Exception:
                row = None

            filename = f"row_{row or idx}.png"
            out_path = os.path.join(images_dir, filename)
            try:
                data = img._data()  # type: ignore[attr-defined]
                with open(out_path, "wb") as f:
                    f.write(data)
                if row:
                    mapping[row] = out_path
            except Exception:
                try:
                    from PIL import Image as PILImage
                    pil = getattr(img, "_image", None)
                    if pil is not None and isinstance(pil, PILImage.Image):
                        pil.save(out_path)
                        if row:
                            mapping[row] = out_path
                except Exception:
                    pass
    except Exception:
        return mapping
    return mapping


# -----------------------------
# Utilities
# -----------------------------
PLACEHOLDERS = {"", "-", "n/a", "na", "none", "null", "not en 388 rated"}

def _norm_str(x: object) -> str:
    """Return a casefolded, stripped string for comparisons."""
    if pd.isna(x):
        return ""
    return str(x).strip().casefold()

def _display_str(x: object) -> str:
    """Return a clean display string (preserve case, strip)."""
    if pd.isna(x):
        return ""
    return str(x).strip()

def _as_int_string_if_number(s: str) -> str:
    """Turn '1.0' -> '1', leave 'X' or non-numeric as-is."""
    try:
        f = float(s)
        if f.is_integer():
            return str(int(f))
        return s
    except Exception:
        return s

def _first_present(mapping: Dict[str, str], df_cols_lower: Dict[str, str]) -> Optional[str]:
    """
    Given desired logical names -> possible header aliases (lowercased),
    return the actual DataFrame column name that matches, if any.
    """
    for alias in mapping.values():
        if alias in df_cols_lower:
            return df_cols_lower[alias]
    return None


# -----------------------------
# Data loader
# -----------------------------
@st.cache_data(show_spinner=False)
def load_data(path: str) -> pd.DataFrame:
    row_to_img = extract_embedded_images(path)

    df = pd.read_excel(path, engine="openpyxl")
    # Strip header whitespace but preserve original case for display
    df.columns = [c.strip() for c in df.columns]

    # Build a lowercase->original-name mapping to support aliases
    df_cols_lower = {c.lower(): c for c in df.columns}

    # Resolve the three dropdown columns by common aliases
    # Colour: support 'colour' or 'color'
    colour_col = _first_present(
        {"Colour": "colour", "Color": "color"}, df_cols_lower
    ) or "Colour"
    cutcat_col = _first_present(
        {"Cut Category": "cut category", "CutCategory": "cutcategory"}, df_cols_lower
    ) or "Cut Category"
    cut_col = _first_present(
        {"Cut": "cut"}, df_cols_lower
    ) or "Cut"

    # Fallback: warn if columns are not present
    missing_display = []
    for col in [colour_col, cutcat_col, cut_col]:
        if col not in df.columns:
            missing_display.append(col)
    if missing_display:
        st.error(
            f"Missing expected column(s): {missing_display}. "
            "Please check your workbook headers."
        )

    # Ensure Image column exists; fill with extracted images by Excel row index
    if "Image" not in df.columns:
        df["Image"] = None
    for i in range(len(df)):
        excel_row = i + 2  # header at row 1
        if pd.isna(df.at[i, "Image"]) or not str(df.at[i, "Image"]).strip():
            if excel_row in row_to_img:
                df.at[i, "Image"] = row_to_img[excel_row]

    # Boolean normalisation for Yes-only filters
    def to_bool(x):
        s = _norm_str(x)
        if s in {"yes", "y", "true", "1"}:
            return True
        if s in {"no", "n", "false", "0", "-"}:
            return False
        return None

    for col in ["Food Safe", "Chemical Resistance", "Heat Resistance"]:
        if col in df.columns:
            df[col + " (bool)"] = df[col].apply(to_bool)

    # Create normalised copies for reliable filtering (keep originals for display)
    if colour_col in df.columns:
        df["_colour_norm"] = df[colour_col].astype(str).str.strip().str.casefold()
    if cutcat_col in df.columns:
        df["_cutcat_norm"] = df[cutcat_col].astype(str).str.strip().str.casefold()
    if cut_col in df.columns:
        # Turn 1.0 -> 1 for better UX, but keep original for display
        df["_cut_display"] = df[cut_col].astype(str).str.strip().map(_as_int_string_if_number)
        df["_cut_norm"] = df["_cut_display"].str.strip().str.casefold()

    # Save the resolved column names for downstream use
    df.attrs["colour_col"] = colour_col
    df.attrs["cutcat_col"] = cutcat_col
    df.attrs["cut_col"] = cut_col

    return df


DATA_PATH = "wurth_safety_glove_MASTER List.xlsx"
df = load_data(DATA_PATH)

# Pull resolved columns
COLOUR_COL = df.attrs.get("colour_col", "Colour")
CUTCAT_COL = df.attrs.get("cutcat_col", "Cut Category")
CUT_COL = df.attrs.get("cut_col", "Cut")


# -----------------------------
# Option builders
# -----------------------------
def _clean_distinct(series: pd.Series) -> List[str]:
    """Distinct, non-placeholder strings for dropdowns, preserving display case."""
    if series is None or series.empty:
        return []
    series = series.dropna()
    vals: List[str] = []
    for v in series:
        s_disp = _display_str(v)
        s_norm = _norm_str(v)
        if not s_disp:
            continue
        if s_norm in PLACEHOLDERS:
            continue
        vals.append(s_disp)
    # Deduplicate by normalised value but keep the first seen display text
    seen = set()
    uniq = []
    for v in vals:
        k = _norm_str(v)
        if k not in seen:
            seen.add(k)
            uniq.append(v)
    return uniq

def options_for(label: str) -> List[str]:
    """
    Build option list for non-boolean filters (first option 'Any') from the DataFrame.
    """
    if label == "Colour" and COLOUR_COL in df.columns:
        uniq = _clean_distinct(df[COLOUR_COL])
        uniq.sort(key=lambda x: x.casefold())
        return ["Any"] + uniq

    if label == "Cut Category" and CUTCAT_COL in df.columns:
        uniq = _clean_distinct(df[CUTCAT_COL])
        # Preferred order A..F; 'X' or others after
        order_map = {k: i for i, k in enumerate(list("ABCDEF"))}
        uniq.sort(key=lambda x: (order_map.get(x.upper(), 99), x.upper()))
        return ["Any"] + uniq

    if label == "Cut rating":
        # Use the cleaned display copy if available
        if "_cut_display" in df.columns:
            uniq = _clean_distinct(df["_cut_display"])
        elif CUT_COL in df.columns:
            uniq = _clean_distinct(df[CUT_COL].astype(str).map(_as_int_string_if_number))
        else:
            uniq = []
        # Numbers first (1,2,3...), then tokens like 'X'
        def _key(x: str):
            return (0, int(x)) if x.isdigit() else (1, x.upper())
        uniq.sort(key=_key)
        return ["Any"] + uniq

    return ["Any"]


# -----------------------------
# Filter UI (two columns, code-ordered)
# -----------------------------
with st.container():
    st.subheader("Filter")

    left_col, right_col = st.columns(2)

    selections: Dict[str, object] = {}

    def render_filter(label: str, col):
        # Yes-only checkboxes
        if label in {"Food Safe?", "Chemical rated?", "Heat rated?"}:
            selections[label] = col.checkbox(label + " (Yes only)", value=False, key=f"cb_{label}")
            return

        # Dropdowns
        opts = options_for(label)
        if DEBUG:
            col.caption(f"{label} options: {opts}")
        selections[label] = col.selectbox(label, opts, index=0, key=f"sb_{label}")

    for label in ORDER_LEFT:
        render_filter(label, left_col)
    for label in ORDER_RIGHT:
        render_filter(label, right_col)

    go = st.button("Search", type="primary")


# -----------------------------
# Apply filters
# -----------------------------
if go:
    filtered = df.copy()

    # Dropdown filters (case-insensitive, trimmed)
    if "Colour" in selections and isinstance(selections["Colour"], str) and selections["Colour"] != "Any":
        right = _norm_str(selections["Colour"])
        if "_colour_norm" in filtered.columns:
            filtered = filtered[filtered["_colour_norm"] == right]

    if "Cut Category" in selections and isinstance(selections["Cut Category"], str) and selections["Cut Category"] != "Any":
        right = _norm_str(selections["Cut Category"])
        if "_cutcat_norm" in filtered.columns:
            filtered = filtered[filtered["_cutcat_norm"] == right]

    if "Cut rating" in selections and isinstance(selections["Cut rating"], str) and selections["Cut rating"] != "Any":
        right = _norm_str(_as_int_string_if_number(selections["Cut rating"]))
        if "_cut_norm" in filtered.columns:
            filtered = filtered[filtered["_cut_norm"] == right]

    # Yes-only boolean filters
    def apply_yes_only(src_col_label: str, sel_key: str) -> None:
        yes_checked = bool(selections.get(sel_key, False))
        col_bool = src_col_label + " (bool)"
        nonlocal filtered  # use outer scope
        if yes_checked and col_bool in filtered.columns:
            mask = filtered[col_bool].fillna(False) == True
            filtered = filtered[mask]

    apply_yes_only("Food Safe", "Food Safe?")
    apply_yes_only("Chemical Resistance", "Chemical rated?")
    apply_yes_only("Heat Resistance", "Heat rated?")

    st.subheader("Search results")
    st.write(f"{len(filtered)} result(s)")

    for _, row in filtered.iterrows():
        with st.container(border=True):
            left, right = st.columns([1, 3])

            # Image
            img_src = row.get("Image", None)
            if isinstance(img_src, str) and img_src.strip():
                try:
                    left.image(img_src, use_column_width=True)
                except Exception:
                    left.info("No image available")
            else:
                left.info("No image available")

            # Details
            name = row.get("Glove Name", "(no name)")
            right.markdown(f"### {name}")
            link = row.get("Product Link", "")
            if isinstance(link, str) and link.strip():
                right.link_button("View product", link)

            # Attributes grid (compact) with integer formatting for EN388 sub-ratings
            attrs: List[tuple] = []
            for label in [
                "Article Numbers", "Colour", "EN 388 Code", "Abrasion", "Cut", "Tear", "Puncture",
                "Cut Category", "Impact", "Chemical Resistance", "Heat Resistance", "Food Safe", "Tactile"
            ]:
                val = row.get(label, None)
                if pd.isna(val):
                    continue
                # Show EN388 sub-ratings with no decimals
                if label in ["Abrasion", "Cut", "Tear", "Puncture"] and isinstance(val, (int, float)):
                    val = int(val)
                attrs.append((label, str(val)))

            a1, a2 = right.columns(2)
            half = (len(attrs) + 1) // 2
            for col, items in [(a1, attrs[:half]), (a2, attrs[half:])]:
                for lab, v in items:
                    col.markdown(f"**{lab}:** {v}")
            st.divider()

    if not filtered.empty:
        csv = filtered.to_csv(index=False)
        st.download_button(
            "Download results (CSV)",
            data=csv,
            file_name="glove_finder_results.csv",
            mime="text/csv",
        )
else:
    st.info("Choose filters above and press **Search** to see matching gloves.")
