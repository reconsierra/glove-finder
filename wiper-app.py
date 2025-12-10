
# app.py
import streamlit as st
import pandas as pd
from typing import List

st.set_page_config(page_title="Glove Finder", page_icon="🧤", layout="wide")
st.title("Glove Finder")
st.caption("Find the right glove by EN388, cut level, category, colour and safety attributes.")

@st.cache_data(show_spinner=False)
def load_data(path: str) -> pd.DataFrame:
    df = pd.read_excel(path, engine="openpyxl")
    # Normalise column names
    df.columns = [c.strip() for c in df.columns]
    # Ensure expected columns exist
    expected = [
        "Glove Name","Article Numbers","Colour","Image","EN 388 Code","Abrasion","Cut",
        "Tear","Puncture","Cut Category","Impact","Chemical Resistance","Heat Resistance",
        "Food Safe","Tactile","Product Link"
    ]
    missing = [c for c in expected if c not in df.columns]
    if missing:
        st.warning(f"Missing columns in data: {missing}. The app will still run but some filters/cards may be incomplete.")
    # Clean Yes/No style values
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
    # Tidy EN 388 entries
    if "EN 388 Code" in df.columns:
        df["EN 388 Code"] = df["EN 388 Code"].fillna("").astype(str).str.strip()
    # Convenience columns
    if "Colour" in df.columns:
        df["Colour"] = df["Colour"].astype(str).str.strip()
    if "Cut Category" in df.columns:
        df["Cut Category"] = df["Cut Category"].astype(str).str.strip()
    return df

DATA_PATH = "wurth_safety_glove_MASTER List.xlsx"
df = load_data(DATA_PATH)

with st.container():
    st.subheader("Filter")
    col1, col2, col3 = st.columns(3)
    col4, col5, col6, col7 = st.columns(4)

    # EN388
    en388_vals = sorted(v for v in df.get("EN 388 Code", pd.Series(dtype=str)).dropna().unique())
    en388_opts = ["Any"] + en388_vals
    en388 = col1.selectbox("EN388 rating", en388_opts, index=0)

    # Cut rating (numeric or X)
    cut_vals = sorted(df.get("Cut", pd.Series(dtype=str)).dropna().astype(str).unique(), key=lambda x: (str(x)))
    cut_opts = ["Any"] + cut_vals
    cut_rating = col2.selectbox("Cut rating", cut_opts, index=0)

    # Cut Category (A-F, X, etc.)
    cut_cat_vals = sorted(df.get("Cut Category", pd.Series(dtype=str)).dropna().unique())
    cut_cat_opts = ["Any"] + cut_cat_vals
    cut_cat = col3.selectbox("Cut Category", cut_cat_opts, index=0)

    # Colour
    colour_vals = sorted(df.get("Colour", pd.Series(dtype=str)).dropna().unique())
    colour_opts = ["Any"] + colour_vals
    colour = col4.selectbox("Colour", colour_opts, index=0)

    # Food Safe?
    food_opts = ["Any","Yes","No"]
    food_safe = col5.selectbox("Food Safe?", food_opts, index=0)

    # Chemical rated?
    chem_opts = ["Any","Yes","No"]
    chem = col6.selectbox("Chemical rated?", chem_opts, index=0)

    # Heat rated?
    heat_opts = ["Any","Yes","No"]
    heat = col7.selectbox("Heat rated?", heat_opts, index=0)

    go = st.button("Search", type="primary")

# Apply filters
if go:
    filtered = df.copy()
    if en388 != "Any" and "EN 388 Code" in filtered.columns:
        filtered = filtered[filtered["EN 388 Code"] == en388]
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

    # Compact cards
    for i, row in filtered.iterrows():
        with st.container(border=True):
            left, right = st.columns([1,3])
            # Image: expects URL or local path in 'Image' column; shows placeholder if not usable
            img_src = row.get("Image", None)
            if isinstance(img_src, str) and img_src.strip():
                try:
                    left.image(img_src, use_column_width=True)
                except Exception:
                    left.info("No image available")
            else:
                left.info("No image available")

            # Right-hand details
            name = row.get("Glove Name", "(no name)")
            right.markdown(f"### {name}")
            # Product link button
            link = row.get("Product Link", "")
            if isinstance(link, str) and link.strip():
                right.link_button("View product", link)
            # Attributes grid
            attrs = []
            for label in [
                "Article Numbers","Colour","EN 388 Code","Abrasion","Cut","Tear","Puncture",
                "Cut Category","Impact","Chemical Resistance","Heat Resistance","Food Safe","Tactile"
            ]:
                val = row.get(label, None)
                if pd.isna(val):
                    continue
                attrs.append((label, str(val)))
            # Two-column compact layout
            a1, a2 = right.columns(2)
            half = (len(attrs)+1)//2
            for col, items in [(a1, attrs[:half]), (a2, attrs[half:])]:
                for label, val in items:
                    col.markdown(f"**{label}:** {val}")
            st.divider()

    # Export button
    if not filtered.empty:
        csv = filtered.to_csv(index=False)
        st.download_button("Download results (CSV)", data=csv, file_name="glove_finder_results.csv", mime="text/csv")
else:
    st.info("Choose filters above and press **Search** to see matching gloves.")
