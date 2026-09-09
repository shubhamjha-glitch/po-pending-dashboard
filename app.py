import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from pathlib import Path
from datetime import date

# =========================================================
# PO PENDING DASHBOARD
# DAILY UPLOADS:
#   1. PO Pending Excel
#   2. Schedule Date Excel
# FIXED MASTER:
#   BARCODE STATUS.xlsx in the same folder as this app.py
# =========================================================

st.set_page_config(
    page_title="PO Pending Dashboard",
    page_icon="📊",
    layout="wide"
)

st.title("📊 PO PENDING DASHBOARD")
st.caption("Upload the daily PO Pending and Schedule Date files. Barcode Status is read from the fixed master file.")

# =========================================================
# FIXED BARCODE MASTER PATH
# =========================================================

APP_FOLDER = Path(__file__).resolve().parent
MASTER_BARCODE_FILE = APP_FOLDER / "BARCODE STATUS.xlsx"

# =========================================================
# OUTPUT COLUMNS
# =========================================================

FINAL_COLUMNS = [
    "ORDER NO.BARCODE",
    "DIVISION",
    "BRAND",
    "VENDOR NAME",
    "VENDOR ID",
    "VENDOR CITY",
    "ORDER DATE",
    "DELIVERY DATE",
    "DELIVERY MNTH",
    "ORDER NO.",
    "ORDER REMARKS",
    "PREPARED BY",
    "ATTRIBUTE1",
    "SECTION",
    "DEPARTMENT",
    "ARTCLE NAME",
    "CAT-1",
    "CAT-3",
    "CAT-4",
    "CAT-5",
    "CAT-6",
    "DESC1",
    "DESC2",
    "DESC3",
    "DESC4",
    "OUTER PACK",
    "DESC5",
    "DESC6",
    "SEASON",
    "STANDARD RATE",
    "LAST IN RATE",
    "ORDER RATE",
    "MRP",
    "HSN CODE",
    "REF_BAR",
    "BARCODE",
    "ORDER QTY",
    "RECEIVED QTY",
    "CANCELLED QTY",
    "PENDING QTY",
    "ORDER VAL",
    "RECEIVED VAL",
    "CANCELLED VAL",
    "PENDING VAL",
    "PKT QTY",
    "BRCD STATUS",
    "EXPIRE DAYS",
    "DELAY GROUP",
    "DELAY GROUP2"
]

# =========================================================
# HELPERS
# =========================================================

def clean_text(value):
    if pd.isna(value):
        return ""
    return str(value).strip()

def normalize_columns(df):
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    return df

def read_po_file(file):
    # The supplied PO report has the actual header on Excel row 3.
    df = pd.read_excel(file, sheet_name="PO PENDING", skiprows=2)
    df = normalize_columns(df)
    df = df.dropna(how="all")

    required = [
        "DIVISION", "CAT-2", "VENDOR NAME", "VENDOR ID", "VENDOR CITY",
        "ORDER DATE", "ORDER NO.", "ORDER REMARKS", "PREPARED BY",
        "ATTRIBUTE1", "SECTION", "DEPARTMENT", "ARTCLE NAME",
        "CAT-1", "CAT-3", "CAT-4", "CAT-5", "CAT-6",
        "DESC1", "DESC2", "DESC3", "DESC4", "OUTER PACK",
        "DESC5", "DESC6", "SEASON", "STANDARD RATE", "LAST IN RATE",
        "ORDER RATE", "MRP", "HSN CODE", "REF_BAR", "BARCODE",
        "ORDER QTY", "RECEIVED QTY", "CANCELLED QTY", "PENDING QTY"
    ]

    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(
            "These columns are missing in the PO Pending file: "
            + ", ".join(missing)
        )

    return df

def read_schedule_file(file):
    df = pd.read_excel(file, sheet_name="SCHEDULE DATE")
    df = normalize_columns(df)
    df = df.dropna(how="all")

    required = ["ORDER NO", "SCHEDULE DATE"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(
            "These columns are missing in the Schedule Date file: "
            + ", ".join(missing)
        )

    df["ORDER NO"] = df["ORDER NO"].map(clean_text)
    df["SCHEDULE DATE"] = pd.to_datetime(
        df["SCHEDULE DATE"], errors="coerce"
    )

    # Prevent duplicate ORDER NO rows from multiplying PO records.
    df = (
        df.sort_values("SCHEDULE DATE")
        .drop_duplicates(subset=["ORDER NO"], keep="last")
    )

    return df[["ORDER NO", "SCHEDULE DATE"]]

def read_barcode_master(path):
    df = pd.read_excel(path)
    df = normalize_columns(df)
    df = df.dropna(how="all")

    required = ["BARCODE", "BRCD STATUS"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(
            "Barcode Status master must contain: BARCODE and BRCD STATUS"
        )

    df["BARCODE"] = df["BARCODE"].map(clean_text)
    df["BRCD STATUS"] = df["BRCD STATUS"].astype(str).str.strip()

    df = df.drop_duplicates(subset=["BARCODE"], keep="last")

    return df[["BARCODE", "BRCD STATUS"]]

def make_delay_group(expire_days):
    if pd.isna(expire_days):
        return "NO DELIVERY DATE"
    if expire_days <= 0:
        return "NOT EXPIRED"
    if expire_days <= 7:
        return "1-7 DAYS"
    if expire_days <= 15:
        return "8-15 DAYS"
    if expire_days <= 30:
        return "16-30 DAYS"
    if expire_days <= 60:
        return "31-60 DAYS"
    if expire_days <= 90:
        return "61-90 DAYS"
    return "90+ DAYS"

def make_delay_group2(expire_days):
    if pd.isna(expire_days):
        return "NO DELIVERY DATE"
    if expire_days <= 0:
        return "NOT DUE"
    return "EXPIRED"

def create_final_data(po_df, schedule_df, barcode_df):
    df = po_df.copy()

    # -----------------------------------------------------
    # Rename CAT-2 to BRAND
    # -----------------------------------------------------
    df = df.rename(columns={"CAT-2": "BRAND"})

    # -----------------------------------------------------
    # Normalize merge keys
    # -----------------------------------------------------
    df["ORDER NO."] = df["ORDER NO."].map(clean_text)
    df["BARCODE"] = df["BARCODE"].map(clean_text)

    # -----------------------------------------------------
    # SCHEDULE DATE = DELIVERY DATE
    # Merge using ORDER NO.
    # -----------------------------------------------------
    df = df.merge(
        schedule_df,
        how="left",
        left_on="ORDER NO.",
        right_on="ORDER NO"
    )

    df = df.drop(columns=["ORDER NO"], errors="ignore")
    df = df.rename(columns={"SCHEDULE DATE": "DELIVERY DATE"})

    # -----------------------------------------------------
    # BARCODE STATUS
    # -----------------------------------------------------
    df = df.merge(
        barcode_df,
        how="left",
        on="BARCODE"
    )

    df["BRCD STATUS"] = df["BRCD STATUS"].fillna("NOT FOUND")

    # -----------------------------------------------------
    # DATE CONVERSION
    # -----------------------------------------------------
    df["ORDER DATE"] = pd.to_datetime(
        df["ORDER DATE"], errors="coerce"
    )

    df["DELIVERY DATE"] = pd.to_datetime(
        df["DELIVERY DATE"], errors="coerce"
    )

    # DELIVERY MNTH
    df["DELIVERY MNTH"] = df["DELIVERY DATE"].dt.strftime("%b-%Y")
    df["DELIVERY MNTH"] = df["DELIVERY MNTH"].fillna("NO DATE")

    # -----------------------------------------------------
    # NUMERIC COLUMNS
    # -----------------------------------------------------
    numeric_cols = [
        "STANDARD RATE", "LAST IN RATE", "ORDER RATE", "MRP",
        "OUTER PACK", "ORDER QTY", "RECEIVED QTY",
        "CANCELLED QTY", "PENDING QTY"
    ]

    for col in numeric_cols:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    # -----------------------------------------------------
    # REQUIRED CALCULATIONS
    # -----------------------------------------------------

    # ORDER NO.BARCODE
    df["ORDER NO.BARCODE"] = (
        df["ORDER NO."].astype(str).str.strip()
        + "."
        + df["BARCODE"].astype(str).str.strip()
    )

    # Value calculations
    df["ORDER VAL"] = df["ORDER QTY"] * df["ORDER RATE"]
    df["RECEIVED VAL"] = df["RECEIVED QTY"] * df["ORDER RATE"]
    df["CANCELLED VAL"] = df["CANCELLED QTY"] * df["ORDER RATE"]
    df["PENDING VAL"] = df["PENDING QTY"] * df["ORDER RATE"]

    # PKT QTY = PENDING QTY / OUTER PACK
    df["PKT QTY"] = np.where(
        df["OUTER PACK"] > 0,
        df["PENDING QTY"] / df["OUTER PACK"],
        0
    )

    # EXPIRE DAYS = Today - Delivery Date
    today = pd.Timestamp(date.today()).normalize()
    df["EXPIRE DAYS"] = (
        today - df["DELIVERY DATE"].dt.normalize()
    ).dt.days

    # DELAY GROUPS
    df["DELAY GROUP"] = df["EXPIRE DAYS"].apply(make_delay_group)
    df["DELAY GROUP2"] = df["EXPIRE DAYS"].apply(make_delay_group2)

    # -----------------------------------------------------
    # ENSURE ALL OUTPUT COLUMNS EXIST
    # -----------------------------------------------------
    for col in FINAL_COLUMNS:
        if col not in df.columns:
            df[col] = ""

    df = df[FINAL_COLUMNS]

    # Sort
    df = df.sort_values(
        by=["DELIVERY DATE", "DIVISION", "BRAND", "VENDOR NAME"],
        na_position="last"
    ).reset_index(drop=True)

    return df


def create_pivot_tables(final_df):
    """Create Excel-style pivot summary tables from the final data."""

    # Division x Delay Group
    pivot_division_delay = pd.pivot_table(
        final_df,
        index="DIVISION",
        columns="DELAY GROUP",
        values="PENDING QTY",
        aggfunc="sum",
        fill_value=0,
        margins=True,
        margins_name="GRAND TOTAL"
    )

    # Brand x Delivery Month
    pivot_brand_month = pd.pivot_table(
        final_df,
        index="BRAND",
        columns="DELIVERY MNTH",
        values="PENDING QTY",
        aggfunc="sum",
        fill_value=0,
        margins=True,
        margins_name="GRAND TOTAL"
    )

    # Vendor summary
    pivot_vendor = pd.pivot_table(
        final_df,
        index=["VENDOR NAME", "BRAND"],
        values=["ORDER QTY", "RECEIVED QTY", "CANCELLED QTY", "PENDING QTY", "PENDING VAL", "PKT QTY"],
        aggfunc="sum",
        fill_value=0,
        margins=True,
        margins_name="GRAND TOTAL"
    )

    # Department x Delay Group
    pivot_department_delay = pd.pivot_table(
        final_df,
        index="DEPARTMENT",
        columns="DELAY GROUP",
        values="PENDING QTY",
        aggfunc="sum",
        fill_value=0,
        margins=True,
        margins_name="GRAND TOTAL"
    )

    return {
        "Division x Delay Group": pivot_division_delay,
        "Brand x Delivery Month": pivot_brand_month,
        "Vendor Summary": pivot_vendor,
        "Department x Delay Group": pivot_department_delay,
    }


def create_pivot_excel(final_df):
    """Create a downloadable workbook containing multiple pivot-style summaries."""
    output = BytesIO()
    pivots = create_pivot_tables(final_df)

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for sheet_name, pivot_df in pivots.items():
            pivot_df.to_excel(writer, sheet_name=sheet_name[:31])

        # Also include source data for reference.
        final_df.to_excel(
            writer,
            sheet_name="Source Data",
            index=False
        )

    output.seek(0)
    return output.getvalue()


def create_excel_file(final_df):
    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        final_df.to_excel(
            writer,
            sheet_name="PO PENDING",
            index=False
        )

        # Summary by Division
        division_summary = (
            final_df.groupby("DIVISION", dropna=False)
            .agg(
                **{
                    "ORDER QTY": ("ORDER QTY", "sum"),
                    "RECEIVED QTY": ("RECEIVED QTY", "sum"),
                    "CANCELLED QTY": ("CANCELLED QTY", "sum"),
                    "PENDING QTY": ("PENDING QTY", "sum"),
                    "PENDING VAL": ("PENDING VAL", "sum"),
                    "PKT QTY": ("PKT QTY", "sum")
                }
            )
            .reset_index()
        )

        division_summary.to_excel(
            writer,
            sheet_name="Division Summary",
            index=False
        )

        # Summary by Brand
        brand_summary = (
            final_df.groupby("BRAND", dropna=False)
            .agg(
                **{
                    "PENDING QTY": ("PENDING QTY", "sum"),
                    "PENDING VAL": ("PENDING VAL", "sum"),
                    "PKT QTY": ("PKT QTY", "sum")
                }
            )
            .reset_index()
            .sort_values("PENDING QTY", ascending=False)
        )

        brand_summary.to_excel(
            writer,
            sheet_name="Brand Summary",
            index=False
        )

        # Summary by Vendor
        vendor_summary = (
            final_df.groupby("VENDOR NAME", dropna=False)
            .agg(
                **{
                    "PENDING QTY": ("PENDING QTY", "sum"),
                    "PENDING VAL": ("PENDING VAL", "sum"),
                    "PKT QTY": ("PKT QTY", "sum")
                }
            )
            .reset_index()
            .sort_values("PENDING QTY", ascending=False)
        )

        vendor_summary.to_excel(
            writer,
            sheet_name="Vendor Summary",
            index=False
        )

    output.seek(0)
    return output.getvalue()

# =========================================================
# FILE UPLOAD
# =========================================================

st.sidebar.header("📁 DAILY FILE UPLOAD")

po_file = st.sidebar.file_uploader(
    "1. Upload PO Pending file",
    type=["xlsx", "xls"],
    key="po_file"
)

schedule_file = st.sidebar.file_uploader(
    "2. Upload Schedule Date file",
    type=["xlsx", "xls"],
    key="schedule_file"
)

# =========================================================
# MAIN PROCESS
# =========================================================

if not MASTER_BARCODE_FILE.exists():
    st.error(
        "BARCODE STATUS.xlsx was not found in the same folder as app.py."
    )
    st.info(
        f"Please place BARCODE STATUS.xlsx here:\n{APP_FOLDER}"
    )
    st.stop()

if po_file is None or schedule_file is None:
    st.info(
        "👈 Upload both daily files from the left side to create the final PO Pending data and dashboard."
    )
    st.stop()

try:
    with st.spinner("Processing files..."):
        po_df = read_po_file(po_file)
        schedule_df = read_schedule_file(schedule_file)
        barcode_df = read_barcode_master(MASTER_BARCODE_FILE)

        final_df = create_final_data(
            po_df,
            schedule_df,
            barcode_df
        )

    # =====================================================
    # FILTERS
    # =====================================================

    st.sidebar.divider()
    st.sidebar.header("🔍 DASHBOARD FILTERS")

    filtered_df = final_df.copy()

    filter_columns = [
        "DIVISION",
        "BRAND",
        "VENDOR NAME",
        "DEPARTMENT",
        "DELIVERY MNTH",
        "BRCD STATUS",
        "DELAY GROUP2"
    ]

    for col in filter_columns:
        options = sorted(
            filtered_df[col]
            .fillna("BLANK")
            .astype(str)
            .unique()
            .tolist()
        )

        selected = st.sidebar.multiselect(
            col,
            options=options,
            default=options,
            key=f"filter_{col}"
        )

        filtered_df = filtered_df[
            filtered_df[col]
            .fillna("BLANK")
            .astype(str)
            .isin(selected)
        ]

    # =====================================================
    # KPI CARDS
    # =====================================================

    st.subheader("📌 Overall Summary")

    total_order = filtered_df["ORDER QTY"].sum()
    total_received = filtered_df["RECEIVED QTY"].sum()
    total_cancelled = filtered_df["CANCELLED QTY"].sum()
    total_pending = filtered_df["PENDING QTY"].sum()
    total_pending_val = filtered_df["PENDING VAL"].sum()
    total_pkt = filtered_df["PKT QTY"].sum()

    c1, c2, c3, c4, c5, c6 = st.columns(6)

    c1.metric("Order Qty", f"{total_order:,.0f}")
    c2.metric("Received Qty", f"{total_received:,.0f}")
    c3.metric("Cancelled Qty", f"{total_cancelled:,.0f}")
    c4.metric("Pending Qty", f"{total_pending:,.0f}")
    c5.metric("Pending Value", f"₹ {total_pending_val:,.0f}")
    c6.metric("Packet Qty", f"{total_pkt:,.2f}")

    # =====================================================
    # DELAY SUMMARY
    # =====================================================

    st.divider()

    col_a, col_b = st.columns(2)

    with col_a:
        st.subheader("⏳ Delay Group Summary")

        delay_summary = (
            filtered_df.groupby("DELAY GROUP", dropna=False)
            .agg(
                **{
                    "Pending Qty": ("PENDING QTY", "sum"),
                    "Pending Value": ("PENDING VAL", "sum"),
                    "Packet Qty": ("PKT QTY", "sum")
                }
            )
            .reset_index()
        )

        st.dataframe(
            delay_summary,
            use_container_width=True,
            hide_index=True
        )

    with col_b:
        st.subheader("🏷️ Brand-wise Pending")

        brand_summary = (
            filtered_df.groupby("BRAND", dropna=False)
            .agg(
                **{
                    "Pending Qty": ("PENDING QTY", "sum"),
                    "Pending Value": ("PENDING VAL", "sum")
                }
            )
            .reset_index()
            .sort_values("Pending Qty", ascending=False)
        )

        st.dataframe(
            brand_summary,
            use_container_width=True,
            hide_index=True
        )

    # =====================================================
    # VENDOR SUMMARY
    # =====================================================

    st.subheader("🏭 Vendor-wise Pending")

    vendor_summary = (
        filtered_df.groupby("VENDOR NAME", dropna=False)
        .agg(
            **{
                "Pending Qty": ("PENDING QTY", "sum"),
                "Pending Value": ("PENDING VAL", "sum"),
                "Packet Qty": ("PKT QTY", "sum")
            }
        )
        .reset_index()
        .sort_values("Pending Qty", ascending=False)
    )

    st.dataframe(
        vendor_summary,
        use_container_width=True,
        hide_index=True,
        height=400
    )

    # =====================================================
    # PIVOT TABLES
    # =====================================================

    st.divider()
    st.subheader("📊 Pivot Table Summary")

    pivot_tables = create_pivot_tables(filtered_df)

    pivot_tabs = st.tabs(list(pivot_tables.keys()))

    for tab, (pivot_name, pivot_df) in zip(pivot_tabs, pivot_tables.items()):
        with tab:
            st.dataframe(
                pivot_df,
                use_container_width=True,
                height=400
            )

    # =====================================================
    # FINAL DATA
    # =====================================================

    st.divider()
    st.subheader("📋 Final PO Pending Data")
    st.write(f"Total records after filters: **{len(filtered_df):,}**")

    st.dataframe(
        filtered_df,
        use_container_width=True,
        hide_index=True,
        height=600
    )

    # =====================================================
    # DOWNLOAD
    # =====================================================

    st.divider()

    excel_data = create_excel_file(filtered_df)

    st.download_button(
        label="⬇️ Download Final PO Pending Excel",
        data=excel_data,
        file_name="FINAL_PO_PENDING.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # =====================================================
    # PIVOT DOWNLOAD
    # =====================================================

    pivot_excel_data = create_pivot_excel(filtered_df)

    st.download_button(
        label="📊 Download Pivot Table Excel",
        data=pivot_excel_data,
        file_name="PO_PENDING_PIVOT_TABLES.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

except Exception as e:
    st.error("An error occurred while processing the files.")
    st.exception(e)
