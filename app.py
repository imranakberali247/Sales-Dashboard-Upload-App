import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.pagesizes import LETTER
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors
from datetime import datetime

# -------------------- App config --------------------
st.set_page_config(page_title="Sales Dashboard (Robust Upload)", layout="wide")
st.title("Sales Dashboard — Upload & Filter (Robust)")
st.caption("Upload Sales & Weight Excel files (messy formats OK). Auto-detects sheet + header; flexible column names.")

# -------------------- Safe helpers --------------------
def _as_series(df, col):
    """Return a single Series for df[col], even if duplicate columns exist or column missing."""
    if df is None or not hasattr(df, "get"):
        return pd.Series([], dtype=str)
    obj = df.get(col)
    if obj is None:
        return pd.Series([], dtype=str)
    if isinstance(obj, pd.DataFrame):  # duplicate column names -> first column
        obj = obj.iloc[:, 0]
    return obj

def _clean_headers(cols):
    return [str(c).strip() for c in cols]

def _normalize_column_names(df: pd.DataFrame):
    """Map many possible header variants to canonical names."""
    if df is None or df.empty:
        return df
    rename = {}
    for c in df.columns:
        cl = str(c).strip().lower()

        # Dates
        if cl in ("date", "order date", "order_date", "txn date", "transaction date"):
            rename[c] = "Order Date"
        elif cl in ("ship date", "ship_date", "shipment date"):
            rename[c] = "Ship Date"

        # Quantities
        elif cl in ("qty", "quantity", "units", "unit qty", "qty sold"):
            rename[c] = "Quantity"
        elif cl in ("net qty", "net quantity", "net_qty"):
            rename[c] = "Net Quantity"

        # Location
        elif cl in ("loc", "location", "site", "sales office", "sales office site"):
            rename[c] = "Location"

        # Product / Size
        elif any(k in cl for k in ("product", "item", "sku", "material", "product name")):
            rename[c] = "Product"
        elif cl in ("size", "pack size", "uom", "package size"):
            rename[c] = "Size"

        # Weight
        elif ("weight" in cl and "lb" in cl) or cl in (
            "weight of indv. product (lb)", "weight_lb", "weight (lb)"
        ):
            rename[c] = "Weight_lb"

    df = df.rename(columns=rename)
    # drop duplicate header names: keep first
    df = df.loc[:, ~pd.Index(df.columns).duplicated()]
    return df

def _coerce_types_sales(df: pd.DataFrame) -> pd.DataFrame:
    """Ensure the minimum Sales columns exist with correct dtypes."""
    if df is None or df.empty:
        # skeleton to keep app running
        df = pd.DataFrame(columns=["Order Date", "Quantity", "Location", "Product", "Size"])

    # Dates
    od = _as_series(df, "Order Date")
    sd = _as_series(df, "Ship Date")
    if not od.empty:
        df["Order Date"] = pd.to_datetime(od, errors="coerce")
    if not sd.empty:
        df["Ship Date"] = pd.to_datetime(sd, errors="coerce")

    # Numbers
    q = _as_series(df, "Quantity")
    if not q.empty:
        df["Quantity"] = pd.to_numeric(q, errors="coerce")

    nq = _as_series(df, "Net Quantity")
    if not nq.empty:
        df["Net Quantity"] = pd.to_numeric(nq, errors="coerce")

    # Text
    for c in ("Location", "Product", "Size"):
        s = _as_series(df, c)
        if not s.empty:
            df[c] = s.astype(str).str.strip()

    # Ensure minimum columns exist
    for col in ["Order Date", "Quantity", "Location", "Product", "Size"]:
        if col not in df.columns:
            df[col] = pd.NA

    keep = ["Order Date", "Quantity", "Location", "Product", "Size"]
    if "Net Quantity" in df.columns:
        keep.append("Net Quantity")
    df = df[keep]
    # drop duplicate header names again, just in case
    df = df.loc[:, ~pd.Index(df.columns).duplicated()]
    return df

def _coerce_types_weight(df: pd.DataFrame) -> pd.DataFrame:
    """Ensure canonical Weight columns and build Key_Product safely."""
    if df is None or df.empty:
        df = pd.DataFrame(columns=["Product", "Weight_lb"])

    df = df.loc[:, ~pd.Index(df.columns).duplicated()]

    prod = _as_series(df, "Product")
    if not prod.empty:
        df["Product"] = prod.astype(str).str.strip()

    w = _as_series(df, "Weight_lb")
    if not w.empty:
        df["Weight_lb"] = pd.to_numeric(w, errors="coerce")

    prod_col = _as_series(df, "Product").astype(str)
    df["Key_Product"] = (
        prod_col.str.lower()
                .str.replace(r"[^a-z0-9]+", " ", regex=True)
                .str.replace(r"\s+", " ", regex=True)
                .str.strip()
    )
    return df[["Product", "Weight_lb", "Key_Product"]]

def _find_candidate_header(df_no_header, required_like, max_scan=40):
    """Scan first N rows for a header row whose columns normalize to required fields."""
    rows = len(df_no_header)
    limit = min(max_scan, rows)
    for i in range(limit):
        header = _clean_headers(df_no_header.iloc[i].tolist())
        data = df_no_header.iloc[i + 1 :].reset_index(drop=True)
        data.columns = header
        data = _normalize_column_names(data)

        cols = set(map(str.lower, data.columns))
        # does it contain all required-like fragments somewhere in the set of names?
        if all(any(req in c for c in cols) for req in required_like):
            return i
    return None

def _read_best_sheet(xls, mode="sales"):
    """Read all sheets (header=None), auto-detect header row, normalize and coerce."""
    raw = pd.read_excel(xls, sheet_name=None, header=None, engine="openpyxl")
    best_df = None
    best_rows = -1
    picked = None

    for name, df in raw.items():
        if df is None or df.empty:
            continue
        df = df.dropna(how="all", axis=1)  # drop fully empty columns
        if df.empty:
            continue

        required = ["product", "weight", "lb"] if mode == "weights" else [
            "date", "order", "quantity", "location", "product", "size"
        ]
        idx = _find_candidate_header(df, required_like=required, max_scan=40)
        if idx is None:
            continue

        header = _clean_headers(df.iloc[idx].tolist())
        data = df.iloc[idx + 1 :].reset_index(drop=True)
        if data.empty:
            continue
        data.columns = header
        # drop duplicate header names: keep first
        data = data.loc[:, ~pd.Index(data.columns).duplicated()]
        data = _normalize_column_names(data)
        data = _coerce_types_weight(data) if mode == "weights" else _coerce_types_sales(data)

        rows = len(data)
        if rows > best_rows:
            best_df = data
            best_rows = rows
            picked = (name, idx)

    return best_df, picked, raw

def robust_read_sales(file):
    return _read_best_sheet(file, "sales")

def robust_read_weights(file):
    return _read_best_sheet(file, "weights")

def _norm_series(s: pd.Series) -> pd.Series:
    return (
        s.astype(str)
         .str.lower()
         .str.replace(r"[^a-z0-9]+", " ", regex=True)
         .str.replace(r"\s+", " ", regex=True)
         .str.strip()
    )

def make_keys(df: pd.DataFrame) -> pd.DataFrame:
    """Create Key_Product and Key_Size safely (handles duplicate columns)."""
    if df is None or df.empty:
        return pd.DataFrame(columns=["Product", "Size", "Key_Product", "Key_Size", "Key_PS"])
    x = df.copy()
    x = x.loc[:, ~pd.Index(x.columns).duplicated()]
    prod = _norm_series(_as_series(x, "Product"))
    size = _norm_series(_as_series(x, "Size"))
    x["Key_Product"] = prod
    x["Key_Size"] = size
    x["Key_PS"] = (prod + " " + size).str.strip()
    return x

def merge_data(sales: pd.DataFrame, weights: pd.DataFrame) -> pd.DataFrame:
    s = make_keys(sales)
    w = weights.copy()
    merged = s.merge(w[["Key_Product", "Weight_lb"]], on="Key_Product", how="left")
    merged = merged[~merged["Weight_lb"].isna()].copy()  # exclude missing weights
    merged["Weight_Total_lb"] = pd.to_numeric(merged["Quantity"], errors="coerce").fillna(0) * merged["Weight_lb"]
    merged["Weight_Total_ton"] = merged["Weight_Total_lb"] / 2000.0
    return merged

def filter_df(df: pd.DataFrame, locations, date_range):
    x = df.copy()
    if locations:
        x = x[x["Location"].isin(locations)]
    if date_range and len(date_range) == 2 and pd.notna(date_range[0]) and pd.notna(date_range[1]):
        start, end = pd.to_datetime(date_range[0]), pd.to_datetime(date_range[1])
        x = x[(x["Order Date"] >= start) & (x["Order Date"] <= end)]
    return x

def to_excel_bytes(dfs: dict) -> bytes:
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        for name, d in dfs.items():
            d.to_excel(writer, index=False, sheet_name=name[:31])
    buf.seek(0)
    return buf.read()

def pdf_bytes(title_suffix: str, kpi: dict, tbl1: pd.DataFrame, tbl2: pd.DataFrame) -> bytes:
    styles = getSampleStyleSheet()
    buf = BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=LETTER, leftMargin=24, rightMargin=24, topMargin=24, bottomMargin=24)

    def _tbl(df: pd.DataFrame, max_rows=25):
        data = [list(df.columns)] + df.head(max_rows).astype(str).values.tolist()
        t = Table(data, repeatRows=1)
        t.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#f0f0f0")),
            ('GRID', (0,0), (-1,-1), 0.25, colors.grey),
            ('FONT', (0,0), (-1,0), 'Helvetica-Bold'),
            ('ALIGN', (1,1), (-1,-1), 'RIGHT'),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ]))
        return t

    story = []
    story.append(Paragraph(f"CTR Report — {title_suffix}", styles["Title"]))
    story.append(Paragraph(datetime.now().strftime("%Y-%m-%d %H:%M"), styles["Normal"]))
    story.append(Spacer(1, 12))
    kpi_lines = "<br/>".join([f"<b>{k}:</b> {v}" for k, v in kpi.items()])
    story.append(Paragraph(kpi_lines, styles["Normal"]))
    story.append(Spacer(1, 12))
    story.append(Paragraph("<b>By Product</b>", styles["Heading2"]))
    story.append(_tbl(tbl1))
    story.append(Spacer(1, 12))
    story.append(Paragraph("<b>By Product → Size</b>", styles["Heading2"]))
    story.append(_tbl(tbl2))
    doc.build(story)
    buf.seek(0)
    return buf.read()

# -------------------- UI --------------------
left, right = st.columns(2)
with left:
    sales_file = st.file_uploader("Upload Sales Data Excel", type=["xlsx"], key="sales")
with right:
    weight_file = st.file_uploader("Upload Weight Reference Excel", type=["xlsx"], key="weights")

if sales_file and weight_file:
    try:
        # Auto-detect & parse
        sales_df, sales_info, sales_raw = robust_read_sales(sales_file)
        weight_df, weight_info, weight_raw = robust_read_weights(weight_file)

        # Manual override if detection failed
        if sales_df is None:
            st.warning("Couldn't auto-detect the Sales sheet/header. Pick manually below.")
            sales_sheet = st.selectbox("Sales sheet", list(sales_raw.keys()), key="sales_sheet")
            df_raw = sales_raw[sales_sheet].dropna(how="all", axis=1)
            hdr_idx = st.number_input("Sales header row index (0-based)", min_value=0, max_value=max(0, len(df_raw)-1), value=0, step=1)
            temp = df_raw.copy()
            temp.columns = _clean_headers(temp.iloc[int(hdr_idx)].tolist())
            temp = temp.iloc[int(hdr_idx)+1:].reset_index(drop=True)
            temp = _normalize_column_names(temp)
            sales_df = _coerce_types_sales(temp)

        if weight_df is None:
            st.warning("Couldn't auto-detect the Weight sheet/header. Pick manually below.")
            weight_sheet = st.selectbox("Weight sheet", list(weight_raw.keys()), key="weight_sheet")
            df_raw = weight_raw[weight_sheet].dropna(how="all", axis=1)
            hdr_idx = st.number_input("Weight header row index (0-based)", min_value=0, max_value=max(0, len(df_raw)-1), value=0, step=1)
            temp = df_raw.copy()
            temp.columns = _clean_headers(temp.iloc[int(hdr_idx)].tolist())
            temp = temp.iloc[int(hdr_idx)+1:].reset_index(drop=True)
            temp = _normalize_column_names(temp)
            weight_df = _coerce_types_weight(temp)

        st.caption(f"Sales detected: sheet {sales_info[0] if sales_info else '?'} header row {sales_info[1] if sales_info else '?'}")
        st.caption(f"Weight detected: sheet {weight_info[0] if weight_info else '?'} header row {weight_info[1] if weight_info else '?'}")

        # Merge + filter
        merged = merge_data(sales_df, weight_df)

        locations = sorted(merged["Location"].dropna().unique().tolist())
        loc_sel = st.multiselect("Location(s)", options=locations, default=locations)

        min_d = pd.to_datetime(merged["Order Date"]).min()
        max_d = pd.to_datetime(merged["Order Date"]).max()
        if pd.isna(min_d) or pd.isna(max_d):
            date_range = None
        else:
            date_range = st.date_input("Order Date range", (min_d.date(), max_d.date()))

        df = filter_df(merged, loc_sel, date_range)

        # KPIs
        kpi_units = int(pd.to_numeric(df["Quantity"], errors="coerce").fillna(0).sum())
        kpi_wlb = round(pd.to_numeric(df["Weight_Total_lb"], errors="coerce").fillna(0).sum(), 2)
        kpi_wtn = round(pd.to_numeric(df["Weight_Total_ton"], errors="coerce").fillna(0).sum(), 3)

        st.subheader("KPIs")
        ka, kb, kc = st.columns(3)
        ka.metric("Total Quantity", f"{kpi_units:,}")
        kb.metric("Total Weight (lb)", f"{kpi_wlb:,}")
        kc.metric("Total Weight (tons)", f"{kpi_wtn:,}")

        # Tables
        st.subheader("By Product")
        by_prod = (
            df.groupby("Product", dropna=False)
              .agg(Quantity=("Quantity", "sum"),
                   Weight_lb=("Weight_Total_lb", "sum"),
                   Weight_ton=("Weight_Total_ton", "sum"))
              .reset_index()
              .sort_values(["Quantity", "Weight_lb"], ascending=False)
        )
        st.dataframe(by_prod, use_container_width=True)

        st.subheader("By Product → Size")
        by_ps = (
            df.groupby(["Product", "Size"], dropna=False)
              .agg(Quantity=("Quantity", "sum"),
                   Weight_lb=("Weight_Total_lb", "sum"),
                   Weight_ton=("Weight_Total_ton", "sum"))
              .reset_index()
              .sort_values(["Product", "Quantity"], ascending=[True, False])
        )
        st.dataframe(by_ps, use_container_width=True)

        # Downloads
        st.divider()
        excel_bytes = to_excel_bytes({"By_Product": by_prod, "By_Product_Size": by_ps})
        st.download_button(
            "Download filtered Excel",
            data=excel_bytes,
            file_name="sales_filtered.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        loc_label = ", ".join(loc_sel) if loc_sel else "All Locations"
        title = f"{loc_label} — Order Date"
        pdf = pdf_bytes(
            title,
            {"Total Quantity": kpi_units, "Total Weight (lb)": kpi_wlb, "Total Weight (tons)": kpi_wtn},
            by_prod, by_ps
        )
        st.download_button("Download PDF report", data=pdf, file_name="CTR_Report.pdf", mime="application/pdf")

        st.caption("Robust parsing enabled. Rows with missing weights are excluded from weight totals.")
    except Exception as e:
        st.error(f"Parser error: {e}")
else:
    st.info("Upload both files to begin.")
