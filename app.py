
import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.pagesizes import LETTER
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors
from datetime import datetime

# --- safe helpers ---
def _as_series(df, col):
    """Return a single Series for df[col], even if duplicate columns exist."""
    obj = df.get(col)
    if obj is None:
        return pd.Series([], dtype=str)
    # if there are duplicate columns, df[col] can be a DataFrame
    if isinstance(obj, pd.DataFrame):
        obj = obj.iloc[:, 0]
    return obj

header = _clean_headers(df.iloc[idx].tolist())
data = df.iloc[idx+1:].reset_index(drop=True)
data.columns = header
# ADD THIS LINE:
data = data.loc[:, ~pd.Index(data.columns).duplicated()]
data = _normalize_column_names(data)
data = _coerce_types_sales(data) if mode=="sales" else _coerce_types_weight(data)

st.set_page_config(page_title="Sales Dashboard (Robust Upload)", layout="wide")
st.title("Sales Dashboard — Upload & Filter (Robust)")
st.caption("Handles multi-sheet files, banners above headers, and flexible column names.")

def _normalize_column_names(df: pd.DataFrame):
    rename = {}
    for c in df.columns:
        cl = str(c).strip().lower()
        if cl in ("date","order date","order_date","txn date","transaction date"):
            rename[c] = "Order Date"
        elif cl in ("ship date","ship_date","shipment date"):
            rename[c] = "Ship Date"
        elif cl in ("qty","quantity","units","unit qty","qty sold"):
            rename[c] = "Quantity"
        elif cl in ("net qty","net quantity","net_qty"):
            rename[c] = "Net Quantity"
        elif cl in ("loc","location","site","sales office","sales office site"):
            rename[c] = "Location"
        elif any(k in cl for k in ("product","item","sku","material","product name")):
            rename[c] = "Product"
        elif cl in ("size","pack size","uom","package size"):
            rename[c] = "Size"
        elif ("weight" in cl and "lb" in cl) or cl in ("weight of indv. product (lb)","weight_lb","weight (lb)"):
            rename[c] = "Weight_lb"
    return df.rename(columns=rename)

def _coerce_types_sales(df):
    if "Order Date" in df.columns:
        df["Order Date"] = pd.to_datetime(df["Order Date"], errors="coerce")
    if "Ship Date" in df.columns:
        df["Ship Date"] = pd.to_datetime(df["Ship Date"], errors="coerce")
    for c in ("Quantity","Net Quantity"):
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce")
    for c in ("Location","Product","Size"):
        if c in df.columns:
            df[c] = df[c].astype(str).str.strip()
    for k in ["Order Date","Quantity","Location","Product","Size"]:
        if k not in df.columns:
            df[k] = pd.NA
    keep = ["Order Date","Quantity","Location","Product","Size"]
    if "Net Quantity" in df.columns: keep.append("Net Quantity")
    return df[keep]

def _coerce_types_weight(df):
    if "Product" in df.columns:
        df["Product"] = df["Product"].astype(str).str.strip()
    if "Weight_lb" in df.columns:
        df["Weight_lb"] = pd.to_numeric(df["Weight_lb"], errors="coerce")
    prod_col = df.get("Product")
    if prod_col is None or not hasattr(prod_col, "astype"):
        prod_col = pd.Series([], dtype=str)
    df["Key_Product"] = (prod_col.astype(str)
                         .str.lower()
                         .str.replace(r"[^a-z0-9]+", " ", regex=True)
                         .str.replace(r"\s+", " ", regex=True)
                         .str.strip())
    return df[["Product","Weight_lb","Key_Product"]]

def _clean_headers(cols):
    return [str(c).strip() for c in cols]

def _find_candidate_header(df, required_like, max_scan=35):
    for i in range(min(max_scan, len(df))):
        header = _clean_headers(df.iloc[i].tolist())
        test = pd.DataFrame(df.iloc[i+1:].values, columns=header)
        test = _normalize_column_names(test)
        cols = set(map(str.lower, test.columns))
        if all(any(r in c for c in cols) for r in required_like):
            return i
    return None

def _read_best_sheet(xls, mode="sales"):
    raw = pd.read_excel(xls, sheet_name=None, header=None, engine="openpyxl")
    best = None; best_rows = -1; pick = None
    for name, df in raw.items():
        df = df.dropna(how="all", axis=1)
        if df.empty: continue
        req = ["date","order","quantity","location","product","size"] if mode=="sales" else ["product","weight","lb"]
        idx = _find_candidate_header(df, req)
        if idx is None: continue
        header = _clean_headers(df.iloc[idx].tolist())
        data = df.iloc[idx+1:].reset_index(drop=True)
        data.columns = header
        data = _normalize_column_names(data)
        data = _coerce_types_sales(data) if mode=="sales" else _coerce_types_weight(data)
        rows = len(data)
        if rows > best_rows:
            best = data; best_rows = rows; pick = (name, idx)
    return best, pick, raw

def robust_read_sales(file):
    return _read_best_sheet(file, "sales")

def robust_read_weights(file):
    return _read_best_sheet(file, "weights")

def make_keys(df: pd.DataFrame) -> pd.DataFrame:
    def norm(s: pd.Series) -> pd.Series:
        return (s.astype(str).str.lower()
                .str.replace(r"[^a-z0-9]+"," ",regex=True)
                .str.replace(r"\s+"," ",regex=True)
                .str.strip())
    x = df.copy()
    x["Key_Product"] = norm(x["Product"])
    x["Key_Size"] = norm(x["Size"])
    x["Key_PS"] = (x["Key_Product"] + " " + x["Key_Size"]).str.strip()
    return x

def merge_data(sales: pd.DataFrame, weights: pd.DataFrame) -> pd.DataFrame:
    s = make_keys(sales)
    w = weights.copy()
    merged = s.merge(w[["Key_Product","Weight_lb"]], on="Key_Product", how="left")
    merged = merged[~merged["Weight_lb"].isna()].copy()
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
    buf.seek(0); return buf.read()

def pdf_bytes(title_suffix: str, kpi: dict, tbl1: pd.DataFrame, tbl2: pd.DataFrame) -> bytes:
    styles = getSampleStyleSheet()
    buf = BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=LETTER, leftMargin=24, rightMargin=24, topMargin=24, bottomMargin=24)

    def _tbl(df: pd.DataFrame, max_rows=25):
        data = [list(df.columns)] + df.head(max_rows).astype(str).values.tolist()
        t = Table(data, repeatRows=1)
        t.setStyle(TableStyle([
            ('BACKGROUND',(0,0),(-1,0),colors.HexColor("#f0f0f0")),
            ('GRID',(0,0),(-1,-1),0.25,colors.grey),
            ('FONT',(0,0),(-1,0),'Helvetica-Bold'),
            ('ALIGN',(1,1),(-1,-1),'RIGHT'),
            ('VALIGN',(0,0),(-1,-1),'MIDDLE'),
        ]))
        return t

    story = []
    story.append(Paragraph(f"CTR Report — {title_suffix}", styles["Title"]))
    story.append(Paragraph(datetime.now().strftime("%Y-%m-%d %H:%M"), styles["Normal"]))
    story.append(Spacer(1, 12))
    kpi_lines = "<br/>".join([f"<b>{k}:</b> {v}" for k,v in kpi.items()])
    story.append(Paragraph(kpi_lines, styles["Normal"]))
    story.append(Spacer(1, 12))
    story.append(Paragraph("<b>By Product</b>", styles["Heading2"])); story.append(_tbl(tbl1)); story.append(Spacer(1, 12))
    story.append(Paragraph("<b>By Product → Size</b>", styles["Heading2"])); story.append(_tbl(tbl2))
    doc.build(story)
    buf.seek(0)
    return buf.read()

left, right = st.columns(2)
with left:
    sales_file = st.file_uploader("Upload Sales Data Excel", type=["xlsx"], key="sales")
with right:
    weight_file = st.file_uploader("Upload Weight Reference Excel", type=["xlsx"], key="weights")

if sales_file and weight_file:
    try:
        sales_df, sales_info, sales_raw = robust_read_sales(sales_file)
        weight_df, weight_info, weight_raw = robust_read_weights(weight_file)

        if sales_df is None:
            st.warning("Couldn't auto-detect Sales header. Pick sheet/header below.")
            sheet = st.selectbox("Sales sheet", list(sales_raw.keys()), key="sales_sheet")
            df_raw = sales_raw[sheet]
            header_row = st.number_input("Sales Header row index (0-based)", min_value=0, max_value=max(0, len(df_raw)-1), value=0, step=1, key="sales_hdr")
            tmp = df_raw.copy().dropna(how="all", axis=1)
            tmp.columns = tmp.iloc[int(header_row)].astype(str)
            tmp = tmp.iloc[int(header_row)+1:]
            tmp = _normalize_column_names(tmp)
            sales_df = _coerce_types_sales(tmp)

        if weight_df is None:
            st.warning("Couldn't auto-detect Weight header. Pick sheet/header below.")
            sheet = st.selectbox("Weight sheet", list(weight_raw.keys()), key="weight_sheet")
            df_raw = weight_raw[sheet]
            header_row = st.number_input("Weight Header row index (0-based)", min_value=0, max_value=max(0, len(df_raw)-1), value=0, step=1, key="weight_hdr")
            tmp = df_raw.copy().dropna(how="all", axis=1)
            tmp.columns = tmp.iloc[int(header_row)].astype(str)
            tmp = tmp.iloc[int(header_row)+1:]
            tmp = _normalize_column_names(tmp)
            weight_df = _coerce_types_weight(tmp)

        st.caption(f"Sales detected: sheet {sales_info[0] if sales_info else '?'} header row {sales_info[1] if sales_info else '?'}")
        st.caption(f"Weight detected: sheet {weight_info[0] if weight_info else '?'} header row {weight_info[1] if weight_info else '?'}")

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

        kpi_units = int(pd.to_numeric(df["Quantity"], errors="coerce").fillna(0).sum())
        kpi_wlb = round(pd.to_numeric(df["Weight_Total_lb"], errors="coerce").fillna(0).sum(), 2)
        kpi_wtn = round(pd.to_numeric(df["Weight_Total_ton"], errors="coerce").fillna(0).sum(), 3)
        st.subheader("KPIs")
        a,b,c = st.columns(3)
        a.metric("Total Quantity", f"{kpi_units:,}")
        b.metric("Total Weight (lb)", f"{kpi_wlb:,}")
        c.metric("Total Weight (tons)", f"{kpi_wtn:,}")

        st.subheader("By Product")
        by_prod = (df.groupby("Product", dropna=False)
                     .agg(Quantity=("Quantity","sum"),
                          Weight_lb=("Weight_Total_lb","sum"),
                          Weight_ton=("Weight_Total_ton","sum"))
                     .reset_index()
                     .sort_values(["Quantity","Weight_lb"], ascending=False))
        st.dataframe(by_prod, use_container_width=True)

        st.subheader("By Product → Size")
        by_ps = (df.groupby(["Product","Size"], dropna=False)
                   .agg(Quantity=("Quantity","sum"),
                        Weight_lb=("Weight_Total_lb","sum"),
                        Weight_ton=("Weight_Total_ton","sum"))
                   .reset_index()
                   .sort_values(["Product","Quantity"], ascending=[True, False]))
        st.dataframe(by_ps, use_container_width=True)

        st.divider()
        ex_bytes = to_excel_bytes({"By_Product": by_prod, "By_Product_Size": by_ps})
        st.download_button("Download filtered Excel", data=ex_bytes, file_name="sales_filtered.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        loc_label = ", ".join(loc_sel) if loc_sel else "All Locations"
        title = f"{loc_label} — Order Date"
        pdf = pdf_bytes(title, {"Total Quantity": kpi_units,
                                "Total Weight (lb)": kpi_wlb,
                                "Total Weight (tons)": kpi_wtn},
                        by_prod, by_ps)
        st.download_button("Download PDF report", data=pdf, file_name="CTR_Report.pdf", mime="application/pdf")

        st.caption("Robust parsing enabled. Rows with missing weights are excluded from weight-based totals.")
    except Exception as e:
        st.error(f"Parser error: {e}")
else:
    st.info("Upload both files to begin.")
