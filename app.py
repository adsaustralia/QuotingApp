
import streamlit as st
import pandas as pd
import numpy as np
import re, json, math, io
from pathlib import Path
from datetime import datetime

# =========================
# Storage paths (local)
# =========================
APP_DIR = Path(__file__).parent
DATA_DIR = APP_DIR / "data"
MAPPING_DIR = DATA_DIR / "mappings"
MAPPING_DIR.mkdir(parents=True, exist_ok=True)

DEFAULT_STANDARD_STOCKS_CSV = DATA_DIR / "standard_stocks.csv"

# =========================
# Helpers
# =========================
def clean_text(s: str) -> str:
    if s is None or (isinstance(s, float) and np.isnan(s)):
        return ""
    s = str(s).strip().lower()
    s = re.sub(r"\s+", " ", s)
    # normalize common symbols
    s = s.replace("Ø", "dia ").replace("⌀", "dia ")
    return s

def load_mapping(customer: str) -> dict:
    fp = MAPPING_DIR / f"{clean_text(customer).replace(' ', '_')}_stock_map.json"
    if fp.exists():
        return json.loads(fp.read_text(encoding="utf-8"))
    return {"customer": customer, "mappings": {}}

def save_mapping(customer: str, mapping: dict) -> None:
    fp = MAPPING_DIR / f"{clean_text(customer).replace(' ', '_')}_stock_map.json"
    fp.write_text(json.dumps(mapping, indent=2), encoding="utf-8")

def parse_size_from_text(text: str):
    """
    Returns dict with keys:
      shape: rectangle|circle|unknown
      width_mm, height_mm, diameter_mm
    """
    t = clean_text(text)
    if not t:
        return {"shape":"unknown", "width_mm":None, "height_mm":None, "diameter_mm":None}

    # circle patterns: dia 600, diameter 600, ø600
    m = re.search(r"(dia(?:meter)?)\s*[:=]?\s*(\d+(?:\.\d+)?)", t)
    if m:
        d = float(m.group(2))
        return {"shape":"circle", "width_mm":None, "height_mm":None, "diameter_mm":d}

    # rectangle patterns: 1200 x 900, 1200*900, 1200 by 900
    m = re.search(r"(\d+(?:\.\d+)?)\s*(x|\*|by)\s*(\d+(?:\.\d+)?)", t)
    if m:
        w = float(m.group(1))
        h = float(m.group(3))
        return {"shape":"rectangle", "width_mm":w, "height_mm":h, "diameter_mm":None}

    return {"shape":"unknown", "width_mm":None, "height_mm":None, "diameter_mm":None}

def sqm_from_geometry(shape: str, width_mm=None, height_mm=None, diameter_mm=None, odd_factor=1.15):
    if shape == "rectangle" and width_mm and height_mm:
        return (width_mm/1000.0) * (height_mm/1000.0)
    if shape == "circle" and diameter_mm:
        r = (diameter_mm/1000.0) / 2.0
        return math.pi * (r**2)
    if shape == "odd" and width_mm and height_mm:
        return (width_mm/1000.0) * (height_mm/1000.0) * odd_factor
    return np.nan

def parse_variant_spec(col_name: str):
    """
    Adidas file uses headers with newline-separated spec, e.g.
      "1000 x 2700\n4 colour\n1pp\nPrinted Mattex Skin\nRubber Edging"

    Returns:
      dict(size_text, colour, sides, stock_customer, finishing)
    """
    raw = "" if col_name is None else str(col_name)
    parts = [p.strip() for p in raw.split("\n") if str(p).strip()]
    out = {"size_text": None, "colour": None, "sides": None, "stock_customer": None, "finishing": None}

    if len(parts) >= 1:
        out["size_text"] = parts[0]
    if len(parts) >= 2:
        out["colour"] = parts[1]
    if len(parts) >= 3:
        out["sides"] = parts[2]
    if len(parts) >= 4:
        out["stock_customer"] = parts[3]
    if len(parts) >= 5:
        out["finishing"] = parts[4]
    return out

def is_probably_variant_col(series: pd.Series) -> bool:
    # variant columns are numeric-ish and contain at least one > 0 value
    s = pd.to_numeric(series, errors="coerce")
    if s.notna().sum() == 0:
        return False
    return (s.fillna(0) > 0).any()

def export_quote_excel(df_lines: pd.DataFrame, df_summary: pd.DataFrame) -> bytes:
    bio = io.BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        df_summary.to_excel(writer, index=False, sheet_name="Quote Summary")
        df_lines.to_excel(writer, index=False, sheet_name="Line Items")
    return bio.getvalue()

def export_quote_pdf(df_summary: pd.DataFrame, df_lines: pd.DataFrame, title="Quote") -> bytes:
    # Lightweight PDF using reportlab
    from reportlab.lib.pagesizes import A4
    from reportlab.pdfgen import canvas
    from reportlab.lib.units import mm

    bio = io.BytesIO()
    c = canvas.Canvas(bio, pagesize=A4)
    w, h = A4

    y = h - 20*mm
    c.setFont("Helvetica-Bold", 14)
    c.drawString(20*mm, y, title)
    y -= 8*mm
    c.setFont("Helvetica", 10)
    c.drawString(20*mm, y, f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}")
    y -= 10*mm

    # Summary
    c.setFont("Helvetica-Bold", 11)
    c.drawString(20*mm, y, "Summary")
    y -= 6*mm
    c.setFont("Helvetica", 10)
    for _, row in df_summary.iterrows():
        line = f"{row['Label']}: {row['Value']}"
        c.drawString(22*mm, y, line)
        y -= 5*mm

    y -= 5*mm
    c.setFont("Helvetica-Bold", 11)
    c.drawString(20*mm, y, "Line Items")
    y -= 6*mm

    # Table header
    headers = ["Description", "Qty", "Size", "Stock", "Total SQM", "Rate", "Line Total"]
    col_x = [20, 70, 95, 135, 175, 195, 210]  # in mm
    c.setFont("Helvetica-Bold", 8)
    for hx, head in zip(col_x, headers):
        c.drawString(hx*mm, y, head)
    y -= 4*mm
    c.setFont("Helvetica", 8)

    def money(x):
        try:
            return f"{float(x):,.2f}"
        except:
            return str(x)

    # rows
    for _, r in df_lines.head(60).iterrows():  # keep MVP simple
        if y < 20*mm:
            c.showPage()
            y = h - 20*mm
            c.setFont("Helvetica-Bold", 8)
            for hx, head in zip(col_x, headers):
                c.drawString(hx*mm, y, head)
            y -= 4*mm
            c.setFont("Helvetica", 8)

        desc = str(r.get("description",""))[:26]
        qty = r.get("qty", "")
        size = str(r.get("size_text",""))[:14]
        stock = str(r.get("stock_std",""))[:18]
        tsqm = r.get("total_sqm", "")
        rate = r.get("sqm_rate", "")
        total = r.get("line_total", "")

        values = [desc, str(qty), size, stock, money(tsqm), money(rate), money(total)]
        for hx, val in zip(col_x, values):
            c.drawString(hx*mm, y, val)
        y -= 4*mm

    c.showPage()
    c.save()
    return bio.getvalue()

# =========================
# Streamlit UI
# =========================
st.set_page_config(page_title="Quoting App (Adidas MVP)", layout="wide")

st.title("Quoting App — Adidas Sheet MVP (Material SQM Pricing)")

with st.sidebar:
    st.header("1) Upload")
    customer = st.text_input("Customer", value="Adidas")
    uploaded = st.file_uploader("Upload customer Excel", type=["xlsx"])

    st.divider()
    st.header("2) Standard Stocks")
    st.caption("This is your internal list with sqm rates. You can replace the sample CSV in /data.")
    if DEFAULT_STANDARD_STOCKS_CSV.exists():
        st.success(f"Loaded default: {DEFAULT_STANDARD_STOCKS_CSV.name}")
    std_upload = st.file_uploader("Optional: upload standard_stocks.csv", type=["csv"])

# Load standard stocks
if std_upload is not None:
    df_std = pd.read_csv(std_upload)
else:
    if DEFAULT_STANDARD_STOCKS_CSV.exists():
        df_std = pd.read_csv(DEFAULT_STANDARD_STOCKS_CSV)
    else:
        df_std = pd.DataFrame(columns=["stock_id","stock_name_std","sqm_rate"])

std_options = df_std["stock_name_std"].dropna().astype(str).tolist()
std_rate_map = dict(zip(df_std["stock_name_std"].astype(str), pd.to_numeric(df_std["sqm_rate"], errors="coerce")))

if uploaded is None:
    st.info("Upload an Adidas Excel file to start.")
    st.stop()

# -------------------------
# Step A: Read sheet + header row
# -------------------------
col1, col2, col3 = st.columns([1,1,2])
# Streamlit doesn't allow dynamic options_kwargs; do it properly
import openpyxl
wb = openpyxl.load_workbook(uploaded, read_only=True, data_only=True)
sheet_names = wb.sheetnames
wb.close()

sheet_name = col1.selectbox("Sheet", sheet_names, index=sheet_names.index("Standard_ WHS Pr MY AUS") if "Standard_ WHS Pr MY AUS" in sheet_names else 0)
header_row_1idx = col2.number_input("Header row (1-indexed)", min_value=1, max_value=100, value=6, step=1)
header = int(header_row_1idx) - 1

df_raw = pd.read_excel(uploaded, sheet_name=sheet_name, header=header)
st.caption(f"Loaded {df_raw.shape[0]} rows × {df_raw.shape[1]} cols from '{sheet_name}' using header row {header_row_1idx}.")

# Show preview
with st.expander("Preview (top 20 rows)", expanded=True):
    st.dataframe(df_raw.head(20), use_container_width=True)

# -------------------------
# Step B: Picker / Mapping
# -------------------------
st.subheader("Pick Values (Mapping)")

# Defaults tuned to Adidas sheet
default_base_cols = [c for c in ["Line Reference","Description","Width","Height","Stock","Finishing","Colour","Sides","Orientation","Qty"] if c in df_raw.columns]

m1, m2, m3, m4 = st.columns([1.2, 1.2, 1.2, 1.2])

with m1:
    col_desc = st.selectbox("Description column", df_raw.columns, index=df_raw.columns.get_loc("Description") if "Description" in df_raw.columns else 0)
with m2:
    col_stock = st.selectbox("Customer stock column", df_raw.columns, index=df_raw.columns.get_loc("Stock") if "Stock" in df_raw.columns else 0)
with m3:
    size_mode = st.radio("Size source", ["Two columns (W+H)", "Single column"], index=0, horizontal=True)
with m4:
    units = st.selectbox("Units", ["mm","cm","m"], index=0)

if size_mode.startswith("Two"):
    cW, cH = st.columns(2)
    with cW:
        col_w = st.selectbox("Width column", df_raw.columns, index=df_raw.columns.get_loc("Width") if "Width" in df_raw.columns else 0)
    with cH:
        col_h = st.selectbox("Height column", df_raw.columns, index=df_raw.columns.get_loc("Height") if "Height" in df_raw.columns else 0)
    col_size = None
else:
    col_size = st.selectbox("Size column", df_raw.columns, index=df_raw.columns.get_loc("Description") if "Description" in df_raw.columns else 0)
    col_w, col_h = None, None

q1, q2, q3 = st.columns([1.2,1.2,1.6])
with q1:
    qty_mode = st.radio("Qty mode", ["Multiple variant columns (Adidas)", "Single column"], index=0)
with q2:
    col_qty_single = st.selectbox("Qty column (if single)", df_raw.columns, index=df_raw.columns.get_loc("Qty") if "Qty" in df_raw.columns else 0)
with q3:
    st.caption("For Adidas variant columns, pick which columns represent quantities. The app will melt them into rows.")

# auto-suggest variant columns (everything not base, numeric and has >0)
base_cols = st.multiselect("ID/Base columns to keep", df_raw.columns.tolist(), default=default_base_cols)
candidate_variant_cols = [c for c in df_raw.columns if c not in base_cols and is_probably_variant_col(df_raw[c])]
default_variant_cols = candidate_variant_cols

variant_cols = st.multiselect("Variant qty columns (melt these)", df_raw.columns.tolist(), default=default_variant_cols, disabled=(qty_mode!="Multiple variant columns (Adidas)"))

shape_default = st.selectbox("Default shape", ["rectangle","circle","odd"], index=0)
odd_factor = st.number_input("Odd-shape factor (used only for shape=odd)", min_value=1.00, max_value=2.00, value=1.15, step=0.01)

# -------------------------
# Step C: Build normalized line items
# -------------------------
st.subheader("Normalized Line Items")

df = df_raw.copy()

# unit conversion factor to mm
unit_to_mm = {"mm":1.0, "cm":10.0, "m":1000.0}
u = unit_to_mm[units]

lines = None

if qty_mode == "Multiple variant columns (Adidas)":
    if not variant_cols:
        st.warning("No variant columns selected. Select at least one.")
        st.stop()
    # Melt
    lines = df.melt(
        id_vars=base_cols,
        value_vars=variant_cols,
        var_name="variant_spec",
        value_name="qty"
    )
    lines["qty"] = pd.to_numeric(lines["qty"], errors="coerce").fillna(0)
    lines = lines[lines["qty"] > 0].copy()

    # Parse spec from column header
    parsed = lines["variant_spec"].apply(parse_variant_spec).apply(pd.Series)
    lines = pd.concat([lines, parsed], axis=1)

    # Decide size source: use parsed size_text if present; else use W/H
    # Convert W/H to mm if available
    if size_mode.startswith("Two") and col_w in lines.columns and col_h in lines.columns:
        lines["width_mm"] = pd.to_numeric(lines[col_w], errors="coerce") * u
        lines["height_mm"] = pd.to_numeric(lines[col_h], errors="coerce") * u
    else:
        lines["width_mm"] = np.nan
        lines["height_mm"] = np.nan

    # From size_text in variant_spec if available
    geo = lines["size_text"].apply(parse_size_from_text).apply(pd.Series)
    lines = pd.concat([lines, geo], axis=1)

    # If parsed width/height missing, fall back to columns
    lines["width_mm"] = np.where(lines["width_mm"].notna(), lines["width_mm"], lines["width_mm"])  # keep
    lines["height_mm"] = np.where(lines["height_mm"].notna(), lines["height_mm"], lines["height_mm"])  # keep
    # Prefer parsed size for width/height
    lines["width_mm"] = np.where(pd.to_numeric(lines["width_mm"], errors="coerce").notna(), lines["width_mm"], lines["width_mm"])
    # If parsed values exist, overwrite
    lines["width_mm"] = np.where(geo["width_mm"].notna(), geo["width_mm"]*u, lines["width_mm"])
    lines["height_mm"] = np.where(geo["height_mm"].notna(), geo["height_mm"]*u, lines["height_mm"])
    lines["diameter_mm"] = np.where(geo["diameter_mm"].notna(), geo["diameter_mm"]*u, np.nan)

    # Set final shape
    lines["shape_final"] = lines["shape"].replace({"unknown": None})
    lines["shape_final"] = lines["shape_final"].fillna(shape_default)

    # Choose customer stock: prefer spec stock if present, else base stock column
    lines["stock_customer_final"] = lines["stock_customer"].fillna(lines[col_stock] if col_stock in lines.columns else None)
    lines["stock_customer_final"] = lines["stock_customer_final"].astype(str)

    # Description
    lines["description"] = lines[col_desc].astype(str) if col_desc in lines.columns else ""

else:
    # single qty column path (still normalized)
    lines = df.copy()
    lines["qty"] = pd.to_numeric(lines[col_qty_single], errors="coerce").fillna(0)
    lines = lines[lines["qty"] > 0].copy()

    if size_mode.startswith("Two"):
        lines["width_mm"] = pd.to_numeric(lines[col_w], errors="coerce") * u
        lines["height_mm"] = pd.to_numeric(lines[col_h], errors="coerce") * u
        lines["diameter_mm"] = np.nan
        lines["shape_final"] = shape_default
        lines["size_text"] = lines.apply(lambda r: f"{r[col_w]} x {r[col_h]}", axis=1)
    else:
        geo = lines[col_size].apply(parse_size_from_text).apply(pd.Series)
        lines = pd.concat([lines, geo], axis=1)
        lines["width_mm"] = geo["width_mm"]*u
        lines["height_mm"] = geo["height_mm"]*u
        lines["diameter_mm"] = geo["diameter_mm"]*u
        lines["shape_final"] = lines["shape"].replace({"unknown": shape_default})
        lines["size_text"] = lines[col_size].astype(str)

    lines["stock_customer_final"] = lines[col_stock].astype(str)
    lines["description"] = lines[col_desc].astype(str)

# SQM + totals
lines["sqm_each"] = lines.apply(lambda r: sqm_from_geometry(
    r.get("shape_final","rectangle"),
    r.get("width_mm", None),
    r.get("height_mm", None),
    r.get("diameter_mm", None),
    odd_factor=odd_factor
), axis=1)

lines["total_sqm"] = lines["sqm_each"] * lines["qty"]

# -------------------------
# Step D: Stock mapping + rate
# -------------------------
st.subheader("Stock Mapping (Customer → Standard)")

mapping = load_mapping(customer)

unique_customer_stocks = sorted(set([s for s in lines["stock_customer_final"].dropna().astype(str).tolist() if clean_text(s)]))

# Build editable mapping table
map_rows = []
for cs in unique_customer_stocks:
    cs_key = clean_text(cs)
    chosen = mapping["mappings"].get(cs_key, "")
    map_rows.append({"customer_stock_name": cs, "standard_stock": chosen})

df_map = pd.DataFrame(map_rows)

st.caption("Pick the matching standard stock for each customer stock name. This is saved for this customer.")
edited_map = st.data_editor(
    df_map,
    use_container_width=True,
    column_config={
        "standard_stock": st.column_config.SelectboxColumn("Standard stock", options=[""] + std_options)
    },
    num_rows="fixed",
    hide_index=True
)

if st.button("Save stock mappings"):
    # persist
    new_map = load_mapping(customer)
    for _, r in edited_map.iterrows():
        cs_key = clean_text(r["customer_stock_name"])
        std_name = (r["standard_stock"] or "").strip()
        if cs_key and std_name:
            new_map["mappings"][cs_key] = std_name
    save_mapping(customer, new_map)
    st.success("Saved mappings.")

# Apply mapping
mapping = load_mapping(customer)
lines["stock_std"] = lines["stock_customer_final"].apply(lambda x: mapping["mappings"].get(clean_text(x), ""))

# Join sqm rates
lines["sqm_rate"] = lines["stock_std"].map(std_rate_map).astype(float)
lines["line_total"] = lines["total_sqm"] * lines["sqm_rate"]

# -------------------------
# Step E: Review + Export
# -------------------------
st.subheader("Quote Review")

missing_map = lines[(lines["stock_std"]=="") | (lines["sqm_rate"].isna())]
if len(missing_map) > 0:
    st.warning(f"{len(missing_map)} line(s) are missing stock mapping or sqm rate. Map them above to get totals.")
else:
    st.success("All lines mapped and rated.")

review_cols = ["description","qty","size_text","shape_final","stock_customer_final","stock_std","sqm_each","total_sqm","sqm_rate","line_total"]
show = lines.copy()

# Keep readable
for c in ["sqm_each","total_sqm","sqm_rate","line_total"]:
    show[c] = pd.to_numeric(show[c], errors="coerce")

st.dataframe(show[review_cols].sort_values(["stock_std","description"]).reset_index(drop=True), use_container_width=True)

# Summary
total_sqm = float(pd.to_numeric(lines["total_sqm"], errors="coerce").fillna(0).sum())
subtotal = float(pd.to_numeric(lines["line_total"], errors="coerce").fillna(0).sum())
df_summary = pd.DataFrame([
    {"Label":"Customer", "Value": customer},
    {"Label":"Sheet", "Value": sheet_name},
    {"Label":"Total SQM", "Value": f"{total_sqm:,.3f}"},
    {"Label":"Subtotal (Material)", "Value": f"{subtotal:,.2f}"},
])

st.markdown("**Totals**")
st.table(df_summary)

c1, c2 = st.columns(2)

with c1:
    xlsx_bytes = export_quote_excel(show[review_cols], df_summary)
    st.download_button(
        "Download Quote Excel",
        data=xlsx_bytes,
        file_name=f"Quote_{customer}_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

with c2:
    pdf_bytes = export_quote_pdf(df_summary, show[review_cols], title=f"Quote - {customer}")
    st.download_button(
        "Download Quote PDF",
        data=pdf_bytes,
        file_name=f"Quote_{customer}_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf",
        mime="application/pdf"
    )

st.divider()
with st.expander("Developer notes / how this MVP is wired for Adidas", expanded=False):
    st.markdown("""
- Default sheet: **Standard_ WHS Pr MY AUS**
- Default header row: **6**
- Variant qty columns: auto-detected as numeric columns with at least one value > 0, excluding base/id columns
- Variant column headers are parsed by splitting on newline:
  - line 1 → size
  - line 2 → colour
  - line 3 → sides
  - line 4 → stock
  - line 5 → finishing
- Size parsing supports:
  - `1200 x 900`
  - `dia 600` / `diameter 600`
- Pricing:
  - Material-only = **total_sqm × sqm_rate** using your standard stocks table
""")
