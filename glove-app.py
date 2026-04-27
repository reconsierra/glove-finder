
# glove-app.py
import os
from typing import Dict, List, Optional, Tuple

import pandas as pd
import streamlit as st

# -----------------------------
# App config
# -----------------------------
st.set_page_config(page_title="Glove Finder", page_icon="🧤", layout="wide")
st.title("Glove Finder")
st.caption("Find the right glove by cut level/category, colour and safety attributes.")

# -----------------------------
# Creator-controlled layout + labels
# -----------------------------
ORDER_LEFT:  List[str] = ["Colour", "Cut Category", "Cut rating"]
ORDER_RIGHT: List[str] = ["Food Safe?", "Chemical rated?", "Heat rated?"]

DISPLAY_LABELS: Dict[str, str] = {
    "Colour": "Glove Colour",
    "Cut Category": "Cut Category (A–F)",
    "Cut rating": "Cut Level",
    "Food Safe?": "Food Safe",
    "Chemical rated?": "Chemical Resistant",
    "Heat rated?": "Heat Resistant",
}

DISPLAY_ATTRS: Dict[str, str] = {
    "Article Numbers": "Article #",
    "Colour": "Glove Colour",
    "EN 388 Code": "EN388 Rating",
    "Abrasion": "Abrasion",
    "Cut": "Cut",
    "Tear": "Tear",
    "Puncture": "Puncture",
    "Cut Category": "Cut Category",
    "Impact": "Impact",
    "Chemical Resistance": "Chemical Resistant",
    "Heat Resistance": "Heat Resistant",
    "Food Safe": "Food Safe",
    "Tactile": "Tactile",
}

# If True, shows debug details (resolved columns + option counts)
DEBUG = False

PLACEHOLDERS = {"", "-", "n/a", "na", "none", "null", "not en 388 rated", "nan"}


# -----------------------------
# Small helpers
# -----------------------------
def norm(x: object) -> str:
    if pd.isna(x):
        return ""
    return str(x).strip().casefold()

def disp(x: object) -> str:
    if pd.isna(x):
        return ""
    return str(x).strip()

def as_int_string_if_number(s: str) -> str:
    # turns "1.0" -> "1", leaves "X" etc
    try:
        f = float(s)
        if f.is_integer():
            return str(int(f))
        return s
    except Exception:
        return s

def resolve_column(aliases: List[str], df_cols_lower: Dict[str, str]) -> Optional[str]:
    """
    Resolve a column name using a list of aliases (already lowercased).
    Returns the original column name or None.
    """
    for a in aliases:
        if a in df_cols_lower:
            return df_cols_lower[a]
    return None


# -----------------------------
# Embedded image extraction (safe / non-fatal)
# -----------------------------
@st.cache_data(show_spinner=False)
def extract_embedded_images(xlsx_path: str, images_dir: str = "images") -> Dict[int, str]:
    """
    Best-effort extract embedded images from the first worksheet.
    Returns a dict: {excel_row_number: saved_image_path}
    Never raises; returns {} on failure.
    """
    mapping: Dict[int, str] = {}
    try:
        from openpyxl import load_workbook
        from openpyxl.utils.cell import coordinate_to_tuple

        wb = load_workbook(xlsx_path, data_only=True)
        ws = wb.active
        os.makedirs(images_dir, exist_ok=True)

        imgs = getattr(ws, "_images", [])
        for idx, img in enumerate(imgs, start=1):
            anchor = getattr(img, "anchor", None)
            row = None

            try:
                if isinstance(anchor, str):
                    row = coordinate_to_tuple(anchor)[0]
                else:
                    from_anchor = getattr(anchor, "_from", None)
                    if from_anchor is not None:
                        row = int(from_anchor.row) + 1
            except Exception:
                row = None

            out_path = os.path.join(images_dir, f"row_{row or idx}.png")

            try:
                data = img._data()  # type: ignore[attr-defined]
                with open(out_path, "wb") as f:
                    f.write(data)
                if row:
                    mapping[row] = out_path
            except Exception:
                # quietly skip image if extraction fails
                pass

    except Exception:
        return {}
    return mapping


# -----------------------------
# Load + normalise data (guarantees dropdowns have values when data exists)
# -----------------------------
@st.cache_data(show_spinner=False)
def load_data(path: str) -> Tuple[pd.DataFrame, Dict[str, str]]:
    if not os.path.exists(path):
        raise FileNotFoundError(f"Excel file not found: {path}")

    df = pd.read_excel(path, engine="openpyxl")
    df.columns = [c.strip() for c in df.columns]
    cols_lower = {c.lower(): c for c in df.columns}

    # Resolve key columns with aliases
    colour_col = resolve_column(["colour", "color"], cols_lower) or "Colour"
    cutcat_col = resolve_column(["cut category", "cutcategory"], cols_lower) or "Cut Category"
    cut_col = resolve_column(["cut"], cols_lower) or "Cut"

    # Validate presence
    missing = [c for c in [colour_col, cutcat_col, cut_col] if c not in df.columns]
    if missing:
        raise KeyError(f"Missing required column(s): {missing}. Found columns: {list(df.columns)}")

    # Create normalised columns for filtering (string-safe)
    df["_colour_norm"] = df[colour_col].astype(str).str.strip().str.casefold()
    df["_cutcat_norm"] = df[cutcat_col].astype(str).str.strip().str.casefold()

    # Cut display: 1.0 -> 1
    cut_display = df[cut_col].astype(str).str.strip().map(as_int_string_if_number)
    df["_cut_display"] = cut_display
    df["_cut_norm"] = cut_display.str.strip().str.casefold()

    # Yes-only boolean columns
    def to_bool(x):
        s = norm(x)
        if s in {"yes", "y", "true", "1"}:
            return True
        if s in {"no", "n", "false", "0", "-"}:
            return False
        return None

    for col in ["Food Safe", "Chemical Resistance", "Heat Resistance"]:
        if col in df.columns:
            df[col + " (bool)"] = df[col].apply(to_bool)

    resolved = {"Colour": colour_col, "Cut Category": cutcat_col, "Cut": cut_col}
    return df, resolved


def clean_options(series: pd.Series) -> List[str]:
    """
    Builds a non-empty option list from series values when any non-placeholder values exist.
    Falls back to raw uniques if cleaning strips everything.
    """
    if series is None or series.empty:
        return []

    # Primary: cleaned
    vals: List[str] = []
    for v in series.dropna():
        d = disp(v)
        if not d:
            continue
        if norm(d) in PLACEHOLDERS:
            continue
        vals.append(d)

    uniq = list(dict.fromkeys(vals))  # preserve order
    if uniq:
        return uniq

    # Fallback: raw uniques (trim only)
    raw = series.dropna().astype(str).str.strip()
    raw_uniq = [v for v in raw.unique().tolist() if v and norm(v) not in PLACEHOLDERS]
    return raw_uniq


# -----------------------------
# Main app (never blank: shows errors)
# -----------------------------
try:
    DATA_PATH = "wurth_safety_glove_MASTER List.xlsx"
    df, resolved_cols = load_data(DATA_PATH)

    # Attempt image extraction; non-fatal
    row_to_img = extract_embedded_images(DATA_PATH)
    if "Image" not in df.columns:
        df["Image"] = None
    for i in range(len(df)):
        excel_row = i + 2
        if (pd.isna(df.at[i, "Image"]) or not str(df.at[i, "Image"]).strip()) and excel_row in row_to_img:
            df.at[i, "Image"] = row_to_img[excel_row]

    # Build dropdown option sets (guaranteed to include data if present)
    colour_opts = ["Any"] + sorted(set(clean_options(df[resolved_cols["Colour"]])), key=lambda x: x.casefold())

    cutcat_opts = ["Any"] + clean_options(df[resolved_cols["Cut Category"]])
    # Sort Cut Category A..F then others
    order_map = {k: i for i, k in enumerate(list("ABCDEF"))}
    cutcat_opts = ["Any"] + sorted(
        [x for x in cutcat_opts[1:] if x],
        key=lambda x: (order_map.get(x.upper(), 99), x.upper())
    )

    cut_opts = ["Any"] + clean_options(df["_cut_display"])
    # Sort cut numeric first then tokens like X
    def cut_key(x: str):
        return (0, int(x)) if x.isdigit() else (1, x.upper())
    cut_opts = ["Any"] + sorted([x for x in cut_opts[1:] if x], key=cut_key)

    if DEBUG:
        st.write("Resolved columns:", resolved_cols)
        st.write("Colour options:", len(colour_opts) - 1)
        st.write("Cut Category options:", len(cutcat_opts) - 1)
        st.write("Cut options:", len(cut_opts) - 1)

    # -----------------------------
    # Filter UI
    # -----------------------------
    st.subheader("Filters")
    left, right = st.columns(2)
    selections: Dict[str, object] = {}

    def render(label: str, col):
        ui_label = DISPLAY_LABELS.get(label, label)
        if label.endswith("?"):
            selections[label] = col.checkbox(ui_label + " (Yes only)", value=False, key=f"cb_{label}")
        else:
            if label == "Colour":
                selections[label] = col.selectbox(ui_label, colour_opts, index=0, key="sb_colour")
            elif label == "Cut Category":
                selections[label] = col.selectbox(ui_label, cutcat_opts, index=0, key="sb_cutcat")
            elif label == "Cut rating":
                selections[label] = col.selectbox(ui_label, cut_opts, index=0, key="sb_cut")
            else:
                selections[label] = col.selectbox(ui_label, ["Any"], index=0, key=f"sb_{label}")

    for l in ORDER_LEFT:
        render(l, left)
    for l in ORDER_RIGHT:
        render(l, right)

    go = st.button("Search", type="primary")

    # -----------------------------
    # Apply filters
    # -----------------------------
    if go:
        filtered = df.copy()

        # Dropdown filters
        if selections.get("Colour") and selections["Colour"] != "Any":
            target = norm(selections["Colour"])
            filtered = filtered[filtered["_colour_norm"] == target]

        if selections.get("Cut Category") and selections["Cut Category"] != "Any":
            target = norm(selections["Cut Category"])
            filtered = filtered[filtered["_cutcat_norm"] == target]

        if selections.get("Cut rating") and selections["Cut rating"] != "Any":
            target = norm(selections["Cut rating"])
            filtered = filtered[filtered["_cut_norm"] == target]

        # Yes-only filters (unchecked = no filter)
        def yes_only(df_in: pd.DataFrame, src_col: str, key: str) -> pd.DataFrame:
            if bool(selections.get(key, False)) and (src_col + " (bool)") in df_in.columns:
                mask = df_in[src_col + " (bool)"].fillna(False) == True
                return df_in[mask]
            return df_in

        filtered = yes_only(filtered, "Food Safe", "Food Safe?")
        filtered = yes_only(filtered, "Chemical Resistance", "Chemical rated?")
        filtered = yes_only(filtered, "Heat Resistance", "Heat rated?")

        st.subheader(f"Results ({len(filtered)})")

        for _, r in filtered.iterrows():
            with st.container(border=True):
                c1, c2 = st.columns([1, 3])

                img = r.get("Image", "")
                if isinstance(img, str) and img.strip():
                    try:
                        c1.image(img, use_column_width=True)
                    except Exception:
                        c1.info("No image available")
                else:
                    c1.info("No image available")

                c2.markdown(f"### {r.get('Glove Name','(no name)')}")
                link = r.get("Product Link", "")
                if isinstance(link, str) and link.strip():
                    c2.link_button("View product", link)

                # Attributes (compact) + no decimals for EN388 sub-ratings
                attrs = []
                for key in [
                    "Article Numbers", "Colour", "EN 388 Code", "Abrasion", "Cut", "Tear", "Puncture",
                    "Cut Category", "Impact", "Chemical Resistance", "Heat Resistance", "Food Safe", "Tactile"
                ]:
                    if key not in r or pd.isna(r[key]):
                        continue
                    val = r[key]
                    if key in ["Abrasion", "Cut", "Tear", "Puncture"] and isinstance(val, (int, float)):
                        val = int(val)
                    attrs.append((DISPLAY_ATTRS.get(key, key), str(val)))

                a1, a2 = c2.columns(2)
                half = (len(attrs) + 1) // 2
                for col, items in [(a1, attrs[:half]), (a2, attrs[half:])]:
                    for k, v in items:
                        col.markdown(f"**{k}:** {v}")

    else:
        st.info("Select filters and press **Search**.")

except Exception as e:
    st.error("The app failed to start. Here’s the error:")
    st.exception(e)
    st.info(
        "Common causes: missing Excel file in repo root, workbook headers not matching, "
        "or a dependency missing in requirements.txt (needs streamlit, pandas, openpyxl, Pillow)."
    )
