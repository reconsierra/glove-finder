
# app.py
import streamlit as st
import pandas as pd
from typing import Dict, List
import os

st.set_page_config(page_title="Glove Finder", page_icon="🧤", layout="wide")
st.title("Glove Finder")
st.caption("Find the right glove by cut level/category, colour and safety attributes.")

# -----------------------------
# Embedded image extraction
# -----------------------------
@st.cache_data(show_spinner=False)
def extract_embedded_images(xlsx_path: str, images_dir: str = "images") -> Dict[int, str]:
    """
    Extract embedded images from the first worksheet and map Excel row -> saved image path.
    Returns a dict: {excel_row_number: image_file_path}
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
                    _row = coordinate_to_tuple(anchor)[0]  # returns (row, col)
                    row = _row
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
        excel_row = i + 2  # header assumed on row 1
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
        st.warning(f"Missing columns in data: {missing}. The app will still run but some filters/cards may be incomplete.")

    # Normalise Yes/No values to booleans for filtering
    def to_bool(x):
        if pd.isna(x):
            return None
        s = str(x).strip().lower()
        if s in {"yes","y","true","1"}:
            return True
        if s in {"no","n","false","0","-"}:
            return False
        return None

    for col in ["Food Safe","Chemical Resistance","Heat Resistance"]:
        if col in df.columns:
            df[col + " (bool)"] = df[col].apply(to_bool)

    for col in ["EN 388 Code","Colour","Cut Category"]:
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
# Filter order editor (simple)
# -----------------------------
DEFAULT_LEFT = ["Colour", "Cut Category", "Cut rating"]
DEFAULT_RIGHT = ["Food Safe?", "Chemical rated?", "Heat rated?"]
ALL_FILTERS = DEFAULT_LEFT + DEFAULT_RIGHT

with st.expander("Change filter display order", expanded=False):
    st.caption("Choose which filters appear (top-to-bottom) in the left and right columns.")
    # Left column order
    left_order = st.multiselect(
        "Left column (top to bottom)",
        ALL_FILTERS,
        default=DEFAULT_LEFT,
        key="order_left",
    )
    # Right column options exclude those already placed left
    remaining = [f for f in ALL_FILTERS if f not in left_order]
    right_order = st.multiselect(
        "Right column (top to bottom)",
        remaining,
        default=[f for f in DEFAULT_RIGHT if f in remaining] or remaining,
        key="order_right",
    )
    # Finalise (append any missing filters to the shorter column, preserving order)
    missing = [f for f in ALL_FILTERS if f not in left_order and f not in right_order]
    if missing:
        if len(left_order) <= len(right_order):
            left_order += missing
        else:
            right_order += missing

# -----------------------------
# Filter UI (two columns)
# -----------------------------
with st.container():
    st.subheader("Filter")

    left_col, right_col = st.columns(2)

    selections: Dict[str, str] = {}

    # Render helper for a single filter
    def render_filter(label: str, col: st.delta_generator.DeltaGenerator):
        if label in BOOL_LABELS:
            # Apply checkbox gives 'Any' state; value checkbox gives Yes/No
            apply_key = f"apply_{label}"
            val_key = f"val_{label}"
            apply = col.checkbox(f"Filter {label}", value=False, key=apply_key)
            val = col.checkbox(label + " (Yes)", value=True, key=val_key)
            selections[label] = "Any" if not apply else ("Yes" if val else "No")
        else:
            opts = options_for(label)
            selections[label] = col.selectbox(label, opts, index=0, key=f"sb_{label}")

    # Left column filters
    for label in left_order:
        render_filter(label, left_col)

    # Right column filters
    for label in right_order:
        render_filter(label, right_col)

    go = st.button("Search", type="primary")

# -----------------------------
# Apply filters
# -----------------------------
if go:
    filtered = df.copy()

    # Non-boolean filters
    for label in ["Colour", "Cut Category", "Cut rating"]:
        val = selections.get(label, "Any")
        if val != "Any":
            colname = LABEL_TO_COL[label]
            filtered = filtered[filtered[colname].astype(str) == str(val)]

    # Boolean filters
    def apply_bool(col_label: str, choice: str) -> pd.DataFrame:
        col_bool = col_label + " (bool)"
        if choice == "Any" or col_bool not in filtered.columns:
            return filtered
        want = True if choice == "Yes" else False
        return filtered[filtered[col_bool] == want]

    filtered = apply_bool(LABEL_TO_COL["Food Safe?"], selections.get("Food Safe?", "Any"))
    filtered = apply_bool(LABEL_TO_COL["Chemical rated?"], selections.get("Chemical rated?", "Any"))
    filtered = apply_bool(LABEL_TO_COL["Heat rated?"], selections.get("Heat rated?", "Any"))

    st.subheader("Search results")
    st.write(f"{len(filtered)} result(s)")

    for _, row in filtered.iterrows():
        with st.container(border=True):
            left, right = st.columns([1,3])
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

            attrs = []
            for label in [
                "Article Numbers","Colour","EN 388 Code","Abrasion","Cut","Tear","Puncture",
                "Cut Category","Impact","Chemical Resistance","Heat Resistance","Food Safe","Tactile"
            ]:
                val = row.get(label, None)
                if pd.isna(val):
                    continue
                attrs.append((label, str(val)))

            a1, a2 = right.columns(2)
            half = (len(attrs)+1)//2
            for col, items in [(a1, attrs[:half]), (a2, attrs[half:])]:
                for label, val in items:
                    col.markdown(f"**{label}:** {val}")
            st.divider()

    if not filtered.empty:
        csv = filtered.to_csv(index=False)
        st.download_button("Download results (CSV)", data=csv,
                           file_name="glove_finder_results.csv", mime="text/csv")
else:
    st.info("Choose filters above and press **Search** to see matching gloves.")
