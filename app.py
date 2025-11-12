
import streamlit as st
import pandas as pd
from io import BytesIO
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.pagesizes import LETTER
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors
from datetime import datetime

st.set_page_config(page_title="Sales Dashboard (Upload-Based)", layout="wide")

st.title("Sales Dashboard — Upload & Filter")
st.caption("Upload two Excel files: Sales + Weight. Then filter and download outputs.")

def normalize_sales(df: pd.DataFrame) -> pd.DataFrame:
    mapping = {}
    for c in df.columns:
        cl = c.lower().strip()
        if cl in ["date","order date"]:
            mapping[c] = "Order Date"
        elif cl == "ship date":
            mapping[c] = "Ship Date"
        elif cl == "quantity":
            mapping[c] = "Quantity"
        elif cl == "net quantity":
            mapping[c] = "Net Quantity"
        elif cl == "location":
            mapping[c] = "Location"
        elif cl == "product":
            mapping[c] = "Product"
        elif cl == "size":
            mapping[c] = "Size"
    df = df.rename(columns=mapping)
    if "Order Date" in df.columns:
        df["Order Date"] = pd.to_datetime(df["Order Date"], errors="coerce")
    elif "Date" in df.columns:
        df["Order Date"] = pd.to_datetime(df["Date"], errors="coerce")
    for col in ["Quantity","Net Quantity"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")
    for col in ["Location","Product","Size"]:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip()
    keep = ["Order Date","Quantity","Location","Product","Size"]
    for k in keep:
        if k not in df.columns:
            df[k] = pd.NA
    return df[keep]

def normalize_weights(df: pd.DataFrame) -> pd.DataFrame:
    mapping = {}
    for c in df.columns:
        cl = c.lower().strip()
        if cl.startswith("product"):
            mapping[c] = "Product"
        elif "weight" in cl and "lb" in cl:
            mapping[c] = "Weight_lb"
    df = df.rename(columns=mapping)
    df["Product"] = df["Product"].astype(str).str.strip()
    df["Weight_lb"] = pd.to_numeric(df["Weight_lb"], errors="coerce")
    df["Key_Product"] = (df["Product"].str.lower()
                         .str.replace(r"[^a-z0-9]+"," ",regex=True)
                         .str.replace(r"\\s+"," ",regex=True)
                         .str.strip())
    return df[["Product","Weight_lb","Key_Product"]]

def make_keys(df: pd.DataFrame) -> pd.DataFrame:
    def norm(s: pd.Series) -> pd.Series:
        return (s.astype(str).str.lower()
                .str.replace(r"[^a-z0-9]+"," ",regex=True)
                .str.replace(r"\\s+"," ",regex=True)
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
    merged["Weight_Total_lb"] = merged["Quantity"].fillna(0) * merged["Weight_lb"]
    merged["Weight_Total_ton"] = merged["Weight_Total_lb"] / 2000.0
    return merged

def filter_df(df: pd.DataFrame, locations, date_range):
    x = df.copy()
    if locations:
        x = x[x["Location"].isin(locations)]
    if date_range and len(date_range) == 2:
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
    sales_file = st.file_uploader("Upload Sales Data Excel", type=["xlsx"])
with right:
    weight_file = st.file_uploader("Upload Weight Reference Excel", type=["xlsx"])

if sales_file and weight_file:
    try:
        sales_raw = pd.read_excel(sales_file)
        weights_raw = pd.read_excel(weight_file)
        sales = normalize_sales(sales_raw)
        weights = normalize_weights(weights_raw)
        merged = merge_data(sales, weights)

        locations = sorted(merged["Location"].dropna().unique().tolist())
        loc_sel = st.multiselect("Location(s)", options=locations, default=locations)
        min_d = pd.to_datetime(merged["Order Date"]).min()
        max_d = pd.to_datetime(merged["Order Date"]).max()
        date_range = st.date_input("Order Date range", (min_d.date(), max_d.date()))

        df = filter_df(merged, loc_sel, date_range)

        kpi_units = int(df["Quantity"].fillna(0).sum())
        kpi_wlb = round(df["Weight_Total_lb"].fillna(0).sum(), 2)
        kpi_wtn = round(df["Weight_Total_ton"].fillna(0).sum(), 3)
        st.subheader("KPIs")
        a,b,c = st.columns(3)
        a.metric("Total Quantity", f"{kpi_units:,}")
        b.metric("Total Weight (lb)", f"{kpi_wlb:,}")
        c.metric("Total Weight (tons)", f"{kpi_wtn:,}")

        st.subheader("By Product")
        by_prod = (df.groupby("Product", dropna=False)
                     .agg(Quantity=("Quantity","sum"),
                          Weight_lb=("Weight_Total_lb","sum"),
                          Weight_ton=("Weight_Total_Ton","sum"))
                     .reset_index())
        # fix column name typo if exists
        if "Weight_Total_Ton" in by_prod.columns and "Weight_ton" not in by_prod.columns:
            by_prod = by_prod.rename(columns={"Weight_Total_Ton":"Weight_ton"})
        by_prod = by_prod.sort_values(["Quantity","Weight_lb"], ascending=False)
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

        st.caption("Rows with missing weights are excluded from weight-based totals.")
    except Exception as e:
        st.error(f"Something went wrong while reading your files: {e}")
else:
    st.info("Upload both files to begin.")
