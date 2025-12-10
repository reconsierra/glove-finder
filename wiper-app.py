
# app.py
import streamlit as st
import pandas as pd
from typing import Dict
import os

st.set_page_config(page_title="Glove Finder", page_icon="🧤", layout="wide")
st.title("Glove Finder")
st.caption("Find the right glove by cut level/category, colour and safety attributes.")

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
        # Ensure output dir
        os.makedirs(images_dir, exist_ok=True)
        # openpyxl stores images in ws._images with anchors
        imgs = getattr(ws, "_images", [])
        for idx, img in enumerate(imgs, start=1):
            anchor = getattr(img, "anchor", None)
            row = None
            # Try to derive row from anchor
            try:
                # Anchor may be a single-cell coordinate like 'D2'
                if isinstance(anchor, str):
                    _row = coordinate_to_tuple(anchor)[0]  # returns (row, col)
                    row = _row
                else:
                    # OneCellAnchor / TwoCellAnchor with ._from.row (0-based)
                    from_anchor = getattr(anchor, "_from", None)
                    if from_anchor is not None:
                        row = int(from_anchor.row) + 1
            except Exception:
                row = None

            # Save image bytes to file
            filename = f"row_{row or idx}.png"
            out_path = os.path.join(images_dir, filename)
            try:
                # Many openpyxl Image objects expose a private _data() with PNG/JPEG bytes
                data = img._data()  # type: ignore[attr-defined]
                with open(out_path, "wb") as f:
                    f.write(data)
                if row:
                    mapping[row] = out_path
            except Exception:
                # If direct bytes not available, try PIL fallback
                try:
                    from PIL import Image as PILImage
                    pil = getattr(img, "_image", None)
                    if pil is not None and isinstance(pil, PILImage.Image):
                        pil.save(out_path)
                        if row:
                            mapping[row] = out_path
                except Exception:
                    # Give up on this image
                    pass
    except Exception:
        return mapping
    return mapping

@st.cache_data(show_spinner=False)
def load_data(path: str) -> pd.DataFrame:
    # Extract any embedded images first; create row -> path mapping
    row_to_img = extract_embedded_images(path)

    df = pd.read_excel(path, engine="openpyxl")
    # Normalise column names (Australian spelling)
    df.columns = [c.strip() for c in df.columns]

    # Ensure Image column exists; if empty, fill from extracted mapping by Excel row number
    if "Image" not in df.columns:
        df["Image"] = None
    # Map by Excel row: pandas index 0 corresponds to Excel row 2 (assuming header row = 1)
    for i in range(len(df)):
        excel_row = i + 2
        if pd.isna(df.at[i, "Image"]) or not str(df.at[i, "Image"]).strip():
            if excel_row in row_to_img:
                df.at[i, "Image"] = row_to_img[excel_row]

    # Expected columns
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

    # Tidy strings
    for col in ["EN 388 Code","Colour","Cut Category"]:
        if col in df.columns:
            df[col] = df[col].fillna("").astype(str).str.strip()

    return df

DATA_PATH = "wurth_safety_glove_MASTER List.xlsx"
df = load_data(DATA_PATH)

with st.container():
    st.subheader("Filter")
    # EN 388 selector removed; rebalance the layout: 3 + 3 columns
    col1, col2, col3 = st.columns(3)
    col4, col5, col6 = st.columns(3)

    # Cut Category (A-F, X, etc.)
    cut_cat_vals = sorted(set(df.get("Cut Category", pd.Series(dtype=str)).dropna().astype(str)))
    cut_cat_opts = ["Any"] + [v for v in cut_cat_vals if v.strip()]
    cut_cat = col2.selectbox("Cut Category", cut_cat_opts, index=0)

    # Cut rating (numeric or X)
    cut_vals = sorted(set(df.get("Cut", pd.Series(dtype=str)).dropna().astype(str)), key=lambda x: str(x))
    cut_opts = ["Any"] + [v for v in cut_vals if v.strip()]
    cut_rating = col1.selectbox("Cut rating", cut_opts, index=0)
    
    # Colour
    colour_vals = sorted(set(df.get("Colour", pd.Series(dtype=str)).dropna().astype(str)))
    colour_opts = ["Any"] + [v for v in colour_vals if v.strip()]
    colour = col3.selectbox("Colour", colour_opts, index=0)

    # Food Safe?
    food_opts = ["Any","Yes","No"]
    food_safe = col4.selectbox("Food Safe?", food_opts, index=0)

    # Chemical rated?
    chem_opts = ["Any","Yes","No"]
    chem = col5.selectbox("Chemical rated?", chem_opts, index=0)

    # Heat rated?
    heat_opts = ["Any","Yes","No"]
    heat = col6.selectbox("Heat rated?", heat_opts, index=0)

    go = st.button("Search", type="primary")

# Apply filters
if go:
    filtered = df.copy()

    # EN 388 filter removed

    if cut_rating != "Any" and "Cut" in filtered.columns:
        filtered = filtered[filtered["Cut"].astype(str) == str(cut_rating)]
    if cut_cat != "Any" and "Cut Category" in filtered.columns:
        filtered = filtered[filtered["Cut Category"].astype(str) == str(cut_cat)]
    if colour != "Any" and "Colour" in filtered.columns:
        filtered = filtered[filtered["Colour"] == colour]

    # Booleans
    def apply_bool(col_label, selection):
        col_bool = col_label + " (bool)"
        if selection == "Any" or col_bool not in filtered.columns:
            return filtered
        val = True if selection == "Yes" else False
        return filtered[filtered[col_bool] == val]

    filtered = apply_bool("Food Safe", food_safe)
    filtered = apply_bool("Chemical Resistance", chem)
    filtered = apply_bool("Heat Resistance", heat)

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
