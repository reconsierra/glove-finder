# app.py
import os
from typing import Dict, List

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
            # Try to derive the row number from the image anchor
            try:
                if isinstance(anchor, str):
                    # e.g. "D2" -> (row, col)
                    _row = coordinate_to_tuple(anchor)[0]
                    row = _row
                else:
                    # OneCellAnchor / TwoCellAnchor -> ._from.row (0-based)
                    from_anchor = getattr(anchor, "_from", None)
                    if from_anchor is not None:
                        row = int(from_anchor.row) + 1
            except Exception:
                row = None

            filename = f"row_{row or idx}.png"
            out_path = os.path.join(images_dir, filename)
            try:
                # Preferred: direct bytes (openpyxl internal)
                data = img._data()  # type: ignore[attr-defined]
                with open(out_path, "wb") as f:
                    f.write(data)
                if row:
                    mapping[row] = out_path
            except Exception:
                # Fallback: try PIL image object, if present
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
# Data loader
# -----------------------------
@st.cache_data(show_spinner=False)
def load_data(path: str) -> pd.DataFrame:
    # Extract any embedded images first; create row -> path mapping
    row_to_img = extract_embedded_images(path)

    df = pd.read_excel(path, engine="openpyxl")
    df.columns = [c.strip() for c in df.columns]

    # Ensure Image column exists; if empty, fill from extracted mapping by Excel row number
    if "Image" not in df.columns:
        df["Image"] = None
    for i in range(len(df)):
        excel_row = i + 2  # header is row 1
        if pd.isna(df.at[i, "Image"]) or not str(df.at[i, "Image"]).strip():
            if excel_row in row_to_img:
                df.at[i, "Image"] = row_to_img[excel_row]

    expected = [
        "Glove Name","Article Numbers","Colour","Image","EN 388 Code","Abrasion","Cut",
        "Tear","Puncture","Cut Category","Impact","Chemical Resistance","Heat Resistance",
        "Food Safe","Tactile","Product Link"
    ]
    missing = [c for c in expected if c not in df.columns]
    if missing:
        st.warning(
            f"Missing columns in data: {missing}. The app will still run but some filters/cards may be incomplete."
        )

    # Normalise Yes/No values to booleans for filtering
    def to_bool(x):
        if pd.isna(x):
            return None
        s = str(x).strip().lower()
        if s in {"yes", "y", "true", "1"}:
            return True
        if s in {"no", "n", "false", "0", "-"}:
            return False
        return None

    for col in ["Food Safe", "Chemical Resistance", "Heat Resistance"]:
        if col in df.columns:
            df[col + " (bool)"] = df[col].apply(to_bool)

    # Tidy strings
    for col in ["EN 388 Code", "Colour", "Cut Category"]:
        if col in df.columns:
            df[col] = df[col].fillna("").astype(str).str.strip()

    return df


DATA_PATH = "wurth_safety_glove_MASTER List.xlsx"
df = load_data(DATA_PATH)


# -----------------------------
# Filter definitions & options
# -----------------------------
# Display labels -> underlying data columns (for filtering)
LABEL_TO_COL = {
    "Colour": "Colour",
    "Cut Category": "Cut Category",
    "Cut rating": "Cut",
    "Food Safe?": "Food Safe",
    "Chemical rated?": "Chemical Resistance",
    "Heat rated?": "Heat Resistance",
}
BOOL_LABELS = {"Food Safe?", "Chemical rated?", "Heat rated?"}

def options_for(label: str) -> List:
    """
    Build option list for non-boolean filters (first option 'Any').
    """
    if label == "Colour":
        vals = sorted(set(df.get("Colour", pd.Series(dtype=str)).dropna().astype(str)))
    elif label == "Cut Category":
        vals = sorted(set(df.get("Cut Category", pd.Series(dtype=str)).dropna().astype(str)))
    elif label == "Cut rating":
        vals = sorted(set(df.get("Cut", pd.Series(dtype=str)).dropna().astype(str)), key=lambda x: str(x))
    else:
        vals = []
    return ["Any"] + [v for v in vals if v.strip()]


# -----------------------------
# Filter UI (two columns, code-ordered)
# -----------------------------
with st.container():
    st.subheader("Filter")

    left_col, right_col = st.columns(2)

    # Store selections
    # For non-boolean filters: value string ("Any" or selected value)
    # For boolean filters: True (Yes only) or False (Any)
    selections: Dict[str, object] = {}

    def render_filter(label: str, col):
        if label in BOOL_LABELS:
            # Single checkbox: if checked -> filter to Yes; if unchecked -> no filter applied
            key = f"yes_{label}"
            selections[label] = col.checkbox(label + " (Yes)", value=False, key=key)
        else:
            opts = options_for(label)
            selections[label] = col.selectbox(label, opts, index=0, key=f"sb_{label}")

    # Render ordered filters — creator defines ORDER_LEFT / ORDER_RIGHT above
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

    # Non-boolean filters
    for label in ["Colour", "Cut Category", "Cut rating"]:
        sel = selections.get(label, "Any")
        if isinstance(sel, str) and sel != "Any":
            colname = LABEL_TO_COL[label]
            filtered = filtered[filtered[colname].astype(str) == str(sel)]

    # Boolean filters (single Yes-only checkbox)
    def apply_yes_only(src_col_label: str, yes_checked: bool) -> pd.DataFrame:
        col_bool = src_col_label + " (bool)"
        if not yes_checked or col_bool not in filtered.columns:
            # Not checked -> no filtering; or no bool column available
            return filtered                   
        mask = filtered[col_bool].fillna(False) == True
        return filtered[mask]



    filtered = apply_yes_only(LABEL_TO_COL["Food Safe?"], bool(selections.get("Food Safe?", False)))
    filtered = apply_yes_only(LABEL_TO_COL["Chemical rated?"], bool(selections.get("Chemical rated?", False)))
    filtered = apply_yes_only(LABEL_TO_COL["Heat rated?"], bool(selections.get("Heat rated?", False)))

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

            # Attributes grid (compact)
            
attrs = []
for label in [
    "Article Numbers", "Colour", "EN 388 Code", "Abrasion", "Cut", "Tear", "Puncture",
    "Cut Category", "Impact", "Chemical Resistance", "Heat Resistance", "Food Safe", "Tactile"
]:

val = row.get(label, None)
    if pd.isna(val):
        continue
    attrs.append((label, str(val)))

            a1, a2 = right.columns(2)
            half = (len(attrs) + 1) // 2
            for col, items in [(a1, attrs[:half]), (a2, attrs[half:])]:
                for label, val in items:
                    col.markdown(f"**{label}:** {val}")
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
