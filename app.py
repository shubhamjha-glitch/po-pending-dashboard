import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from pathlib import Path
from datetime import date, timedelta
import tempfile
import os


# ============================================================
# PAGE CONFIG
# ============================================================

st.set_page_config(
    page_title="PO Pending Dashboard",
    page_icon="📊",
    layout="wide"
)


# ============================================================
# APP FOLDER
# ============================================================

APP_FOLDER = Path(__file__).resolve().parent

BARCODE_MASTER_FILE = APP_FOLDER / "BARCODE STATUS.xlsx"


# ============================================================
# CSS
# ============================================================

st.markdown(
    """
    <style>

    .main-title {
        font-size: 30px;
        font-weight: 700;
        margin-bottom: 0px;
    }

    .sub-title {
        font-size: 14px;
        color: #666;
        margin-top: 0px;
        margin-bottom: 20px;
    }

    div[data-testid="stMetric"] {
        border: 1px solid #dddddd;
        padding: 10px;
        border-radius: 8px;
        background-color: #fafafa;
    }

    </style>
    """,
    unsafe_allow_html=True
)


# ============================================================
# TITLE
# ============================================================

st.markdown(
    '<div class="main-title">📊 PO Pending Dashboard</div>',
    unsafe_allow_html=True
)

updated_till = date.today() - timedelta(days=1)

st.markdown(
    f'<div class="sub-title">Updated Till: '
    f'{updated_till.strftime("%d-%b-%Y")}</div>',
    unsafe_allow_html=True
)


# ============================================================
# FINAL COLUMNS
# ============================================================

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


# ============================================================
# HELPERS
# ============================================================

def clean_text(value):

    if pd.isna(value):
        return ""

    return str(value).strip()


def normalize_columns(df):

    df.columns = [
        str(col).strip()
        for col in df.columns
    ]

    return df


# ============================================================
# READ PO FILE
# ============================================================

@st.cache_data(show_spinner=False)
def read_po_file(file):

    df = pd.read_excel(
        file,
        sheet_name="PO PENDING",
        skiprows=2
    )

    df = normalize_columns(df)

    df = df.dropna(
        how="all"
    )

    return df


# ============================================================
# READ SCHEDULE FILE
# ============================================================

@st.cache_data(show_spinner=False)
def read_schedule_file(file):

    df = pd.read_excel(
        file,
        sheet_name="SCHEDULE DATE"
    )

    df = normalize_columns(df)

    required = [
        "ORDER NO",
        "SCHEDULE DATE"
    ]

    missing = [
        col
        for col in required
        if col not in df.columns
    ]

    if missing:

        raise ValueError(
            "Schedule file missing columns: "
            + ", ".join(missing)
        )

    df["ORDER NO"] = (
        df["ORDER NO"]
        .apply(clean_text)
    )

    df["SCHEDULE DATE"] = pd.to_datetime(
        df["SCHEDULE DATE"],
        errors="coerce"
    )

    df = df.sort_values(
        "SCHEDULE DATE"
    )

    df = df.drop_duplicates(
        subset=["ORDER NO"],
        keep="last"
    )

    return df[
        ["ORDER NO", "SCHEDULE DATE"]
    ]


# ============================================================
# READ BARCODE MASTER
# ============================================================

@st.cache_data(show_spinner=False)
def read_barcode_master(file_path):

    df = pd.read_excel(
        file_path
    )

    df = normalize_columns(df)

    # Accept different names used in barcode-status master files.
    canonical = {str(c).strip().upper().replace(" ", "_"): c for c in df.columns}
    lookup_col = next((canonical[k] for k in
                       ("NEW_BAR", "BARCODE", "NEW_BARCODE", "ITEM_BARCODE", "BAR_CODE")
                       if k in canonical), None)
    status_col = next((canonical[k] for k in
                       ("BRCD_STATUS", "BARCODE_STATUS", "STATUS")
                       if k in canonical), None)
    if lookup_col is None or status_col is None:
        raise ValueError(
            "Barcode master columns not recognized. Found: "
            + ", ".join(map(str, df.columns))
            + ". Expected NEW_BAR or BARCODE and BRCD STATUS."
        )
    df = df.rename(columns={lookup_col: "_MASTER_BARCODE", status_col: "BRCD STATUS"})
    df["_MASTER_BARCODE"] = df["_MASTER_BARCODE"].apply(clean_text)
    df["BRCD STATUS"] = df["BRCD STATUS"].fillna("NOT FOUND").astype(str).str.strip()
    return df.drop_duplicates(subset=["_MASTER_BARCODE"], keep="last")[["_MASTER_BARCODE", "BRCD STATUS"]]


# ============================================================
# DELAY GROUP
# ============================================================

def get_delay_group(expire_days):

    if pd.isna(expire_days):
        return "NO DELIVERY DATE"

    if expire_days <= 0:
        return "NOT EXPIRED"

    elif expire_days <= 7:
        return "1-7 DAYS"

    elif expire_days <= 15:
        return "8-15 DAYS"

    elif expire_days <= 30:
        return "16-30 DAYS"

    elif expire_days <= 60:
        return "31-60 DAYS"

    elif expire_days <= 90:
        return "61-90 DAYS"

    else:
        return "90+ DAYS"


# ============================================================
# DELAY GROUP 2
# ============================================================

def get_delay_group2(expire_days):

    if pd.isna(expire_days):
        return "NO DELIVERY DATE"

    if expire_days <= 0:
        return "NOT DUE"

    return "EXPIRED"


# ============================================================
# CREATE FINAL DATA
# ============================================================

def create_final_data(
    po_df,
    schedule_df,
    barcode_df
):

    df = po_df.copy()
    original_input_columns = list(df.columns)

    # --------------------------------------------------------
    # BRAND
    # --------------------------------------------------------

    if "CAT-2" in df.columns:

        if "BRAND" in df.columns:

            df["BRAND"] = df["CAT-2"]

        else:

            df["BRAND"] = df["CAT-2"]

    # --------------------------------------------------------
    # REQUIRED COLUMNS
    # --------------------------------------------------------

    required = [
        "ORDER NO.",
        "BARCODE",
        "PENDING QTY",
        "OUTER PACK",
        "ORDER RATE"
    ]

    missing = [
        col
        for col in required
        if col not in df.columns
    ]

    if missing:

        raise ValueError(
            "PO file missing required columns: "
            + ", ".join(missing)
        )

    # --------------------------------------------------------
    # CLEAN KEYS
    # --------------------------------------------------------

    df["ORDER NO."] = (
        df["ORDER NO."]
        .apply(clean_text)
    )

    df["BARCODE"] = (
        df["BARCODE"]
        .apply(clean_text)
    )

    # --------------------------------------------------------
    # MERGE SCHEDULE
    # --------------------------------------------------------

    df = df.drop(columns=["DELIVERY DATE", "SCHEDULE DATE"], errors="ignore")
    df = df.merge(
        schedule_df,
        left_on="ORDER NO.",
        right_on="ORDER NO",
        how="left"
    )

    df.rename(
        columns={
            "SCHEDULE DATE": "DELIVERY DATE"
        },
        inplace=True
    )

    if "ORDER NO" in df.columns:

        df.drop(
            columns=["ORDER NO"],
            inplace=True
        )

    # --------------------------------------------------------
    # MERGE BARCODE MASTER
    # --------------------------------------------------------

    # Prefer NEW_BAR, but retain compatibility with older PO input files.
    lookup_field = "NEW_BAR" if "NEW_BAR" in df.columns else "BARCODE"
    df[lookup_field] = df[lookup_field].apply(clean_text)
    df = df.drop(columns=["BRCD STATUS"], errors="ignore")
    df = df.merge(
        barcode_df,
        left_on=lookup_field,
        right_on="_MASTER_BARCODE",
        how="left"
    ).drop(columns=["_MASTER_BARCODE"])

    df["BRCD STATUS"] = (
        df["BRCD STATUS"]
        .fillna("NOT FOUND")
        .astype(str)
        .str.strip()
    )

    # --------------------------------------------------------
    # DATES
    # --------------------------------------------------------

    if "ORDER DATE" in df.columns:

        df["ORDER DATE"] = pd.to_datetime(
            df["ORDER DATE"],
            errors="coerce"
        )

    else:

        df["ORDER DATE"] = pd.NaT

    df["DELIVERY DATE"] = pd.to_datetime(
        df["DELIVERY DATE"],
        errors="coerce"
    )

    # --------------------------------------------------------
    # DELIVERY MONTH
    # --------------------------------------------------------

    df["DELIVERY MNTH"] = (
        df["DELIVERY DATE"]
        .dt.strftime("%b-%Y")
    )

    df["DELIVERY MNTH"] = (
        df["DELIVERY MNTH"]
        .fillna("NO DATE")
    )

    # --------------------------------------------------------
    # NUMERIC COLUMNS
    # --------------------------------------------------------

    numeric_columns = [
        "STANDARD RATE",
        "LAST IN RATE",
        "ORDER RATE",
        "MRP",
        "OUTER PACK",
        "ORDER QTY",
        "RECEIVED QTY",
        "CANCELLED QTY",
        "PENDING QTY"
    ]

    for col in numeric_columns:

        if col not in df.columns:

            df[col] = 0

        df[col] = pd.to_numeric(
            df[col],
            errors="coerce"
        ).fillna(0)

    # --------------------------------------------------------
    # REMOVE ZERO PENDING
    # --------------------------------------------------------

    df = df[
        df["PENDING QTY"] > 0
    ].copy()

    # --------------------------------------------------------
    # ORDER NO + BARCODE
    # --------------------------------------------------------

    df["ORDER NO.BARCODE"] = (
        df["ORDER NO."].astype(str)
        + "."
        + df["BARCODE"].astype(str)
    )

    # --------------------------------------------------------
    # VALUES
    # --------------------------------------------------------

    df["ORDER VAL"] = (
        df["ORDER QTY"]
        * df["ORDER RATE"]
    )

    df["RECEIVED VAL"] = (
        df["RECEIVED QTY"]
        * df["ORDER RATE"]
    )

    df["CANCELLED VAL"] = (
        df["CANCELLED QTY"]
        * df["ORDER RATE"]
    )

    df["PENDING VAL"] = (
        df["PENDING QTY"]
        * df["ORDER RATE"]
    )

    # --------------------------------------------------------
    # PKT QTY
    # --------------------------------------------------------

    df["PKT QTY"] = np.where(
        df["OUTER PACK"] > 0,
        df["PENDING QTY"]
        / df["OUTER PACK"],
        0
    )

    # --------------------------------------------------------
    # EXPIRE DAYS
    # --------------------------------------------------------

    today = pd.Timestamp.today().normalize()

    df["EXPIRE DAYS"] = (
        today - df["DELIVERY DATE"]
    ).dt.days

    # --------------------------------------------------------
    # DELAY GROUPS
    # --------------------------------------------------------

    df["DELAY GROUP"] = (
        df["EXPIRE DAYS"]
        .apply(get_delay_group)
    )

    df["DELAY GROUP2"] = (
        df["EXPIRE DAYS"]
        .apply(get_delay_group2)
    )

    # --------------------------------------------------------
    # ADD MISSING COLUMNS
    # --------------------------------------------------------

    for col in FINAL_COLUMNS:

        if col not in df.columns:

            df[col] = ""

    # --------------------------------------------------------
    # FINAL ORDER
    # --------------------------------------------------------

    # Keep the established output layout, but insert any additional input
    # columns beside their original neighboring headers (not at the end).
    output_columns = list(FINAL_COLUMNS)
    input_columns = [col for col in original_input_columns if col in df.columns]
    extra_columns = [col for col in input_columns if col not in output_columns and col != "CAT-2"]

    for extra_col in extra_columns:
        input_position = input_columns.index(extra_col)
        preceding = [col for col in input_columns[:input_position]
                     if col in output_columns]
        following = [col for col in input_columns[input_position + 1:]
                     if col in output_columns]

        if preceding:
            # Insert after the closest preceding input header, and after
            # earlier new columns that were inserted at the same location.
            anchor = preceding[-1]
            insert_at = output_columns.index(anchor) + 1
            while (insert_at < len(output_columns)
                   and output_columns[insert_at] in extra_columns
                   and input_columns.index(output_columns[insert_at]) < input_position):
                insert_at += 1
            output_columns.insert(insert_at, extra_col)
        elif following:
            output_columns.insert(output_columns.index(following[0]), extra_col)
        else:
            output_columns.append(extra_col)

    # Show BRCD STATUS directly after NEW_BAR, when NEW_BAR is supplied.
    if "NEW_BAR" in output_columns:
        output_columns.remove("BRCD STATUS")
        output_columns.insert(output_columns.index("NEW_BAR") + 1, "BRCD STATUS")

    # CAT-2 is represented by BRAND; do not export CAT-2 separately.
    output_columns = [col for col in output_columns if col != "CAT-2"]
    df = df[output_columns]

    # --------------------------------------------------------
    # SORT
    # --------------------------------------------------------

    sort_columns = [
        "DIVISION",
        "SECTION",
        "DEPARTMENT",
        "ARTCLE NAME",
        "BRAND",
        "DELIVERY DATE"
    ]

    available_sort = [
        col
        for col in sort_columns
        if col in df.columns
    ]

    if available_sort:

        df = df.sort_values(
            available_sort,
            na_position="last"
        )

    return df.reset_index(
        drop=True
    )


# ============================================================
# HIERARCHY SUMMARY
# ============================================================

def create_hierarchy_summary(
    df,
    hierarchy_column
):

    if hierarchy_column not in df.columns:

        return pd.DataFrame()

    summary = (
        df.groupby(
            hierarchy_column,
            dropna=False
        )
        .agg(
            {
                "ORDER QTY": "sum",
                "RECEIVED QTY": "sum",
                "CANCELLED QTY": "sum",
                "PENDING QTY": "sum",
                "PENDING VAL": "sum",
                "PKT QTY": "sum"
            }
        )
        .reset_index()
    )

    summary = summary.sort_values(
        "PENDING QTY",
        ascending=False
    )

    return summary


# ============================================================
# CREATE FINAL EXCEL
# ============================================================

def create_final_excel(df):

    output = BytesIO()

    with pd.ExcelWriter(
        output,
        engine="openpyxl"
    ) as writer:

        df.to_excel(
            writer,
            sheet_name="PO PENDING",
            index=False
        )

        # ----------------------------------------------------
        # FORMAT ALL SHEETS
        # APTOS FONT - SIZE 8
        # ----------------------------------------------------

        for worksheet in writer.book.worksheets:

            for row in worksheet.iter_rows():

                for cell in row:

                    cell.font = cell.font.copy(
                        name="Aptos",
                        sz=8
                    )

    output.seek(0)

    return output.getvalue()


# ============================================================
# CREATE LIVE PIVOT EXCEL
# ============================================================

def create_pivot_excel(df):

    try:

        import win32com.client as win32

    except ImportError:

        raise RuntimeError(
            "pywin32 is required for Live PivotTable.\n\n"
            "Run this command in Command Prompt:\n"
            "pip install pywin32"
        )

    export_df = df.copy()

    # ========================================================
    # PIVOT STRUCTURE
    # ========================================================

    # --------------------------------------------------------
    # FILTERS
    # --------------------------------------------------------

    filter_fields = [
        "DELAY GROUP2",
        "DIVISION",
        "ATTRIBUTE1",
        "DELIVERY MNTH"
    ]

    # --------------------------------------------------------
    # ROWS
    # --------------------------------------------------------

    row_fields = [
        "VENDOR NAME",
        "BRAND",
        "ORDER DATE",
        "DELIVERY DATE",
        "ORDER NO.",
        "ORDER REMARKS",
        "DEPARTMENT",
        "ARTCLE NAME",
        "BARCODE",
        "REF_BAR",
        "CAT-1",
        "ORDER RATE",
        "MRP",
        "BRCD STATUS",
        "DESC4",
        "OUTER PACK"
    ]

    # --------------------------------------------------------
    # VALUES
    # --------------------------------------------------------

    value_fields = [
        "ORDER QTY",
        "RECEIVED QTY",
        "CANCELLED QTY",
        "PENDING QTY",
        "PKT QTY"
    ]

    # ========================================================
    # CHECK REQUIRED FIELDS
    # ========================================================

    all_pivot_fields = (
        filter_fields
        + row_fields
        + value_fields
    )

    missing_fields = [
        field
        for field in all_pivot_fields
        if field not in export_df.columns
    ]

    if missing_fields:

        raise RuntimeError(
            "The following Pivot fields are missing:\n\n"
            + ", ".join(missing_fields)
        )

    # ========================================================
    # FORCE NUMERIC VALUE COLUMNS
    # ========================================================

    for field in value_fields:

        export_df[field] = pd.to_numeric(
            export_df[field],
            errors="coerce"
        ).fillna(0)

    # ========================================================
    # CLEAN DATE COLUMNS
    # ========================================================

    for field in [
        "ORDER DATE",
        "DELIVERY DATE"
    ]:

        export_df[field] = pd.to_datetime(
            export_df[field],
            errors="coerce"
        )

    # ========================================================
    # CLEAN TEXT COLUMNS
    # ========================================================

    text_columns = [
        "DELAY GROUP2",
        "DIVISION",
        "ATTRIBUTE1",
        "DELIVERY MNTH",
        "VENDOR NAME",
        "BRAND",
        "ORDER NO.",
        "ORDER REMARKS",
        "DEPARTMENT",
        "ARTCLE NAME",
        "BARCODE",
        "REF_BAR",
        "CAT-1",
        "BRCD STATUS",
        "DESC4"
    ]

    for field in text_columns:

        export_df[field] = (
            export_df[field]
            .fillna("")
            .astype(str)
            .str.strip()
        )

    # ========================================================
    # CREATE TEMP FOLDER
    # ========================================================

    temp_dir = tempfile.mkdtemp()

    temp_file = os.path.join(
        temp_dir,
        "PO_PENDING_WITH_LIVE_PIVOT.xlsx"
    )

    excel = None
    workbook = None

    try:

        # ====================================================
        # WRITE SOURCE EXCEL
        # ====================================================

        with pd.ExcelWriter(
            temp_file,
            engine="openpyxl"
        ) as writer:

            export_df.to_excel(
                writer,
                sheet_name="PO PENDING",
                index=False
            )

        # ====================================================
        # OPEN EXCEL
        # ====================================================

        excel = win32.DispatchEx(
            "Excel.Application"
        )

        excel.Visible = False
        excel.DisplayAlerts = False

        workbook = excel.Workbooks.Open(
            os.path.abspath(temp_file)
        )

        source_sheet = workbook.Worksheets(
            "PO PENDING"
        )

        # ====================================================
        # SOURCE RANGE
        # ====================================================

        last_row = len(export_df) + 1
        last_col = len(export_df.columns)

        source_range = source_sheet.Range(
            source_sheet.Cells(1, 1),
            source_sheet.Cells(
                last_row,
                last_col
            )
        )

        # ====================================================
        # FORMAT SOURCE HEADER
        # ====================================================

        source_header = source_sheet.Range(
            source_sheet.Cells(1, 1),
            source_sheet.Cells(1, last_col)
        )

        source_header.Font.Bold = True

        # ====================================================
        # CREATE EXCEL TABLE
        # ====================================================

        table = None

        try:

            table = source_sheet.ListObjects.Add(
                1,
                source_range,
                None,
                1
            )

            table.Name = "POPendingData"

        except Exception:

            table = None

        # ====================================================
        # REMOVE OLD PIVOT SHEET IF EXISTS
        # ====================================================

        try:

            for ws in workbook.Worksheets:

                if ws.Name == "PIVOT TABLE":

                    ws.Delete()

                    break

        except Exception:

            pass

        # ====================================================
        # CREATE PIVOT SHEET
        # ====================================================

        pivot_sheet = workbook.Worksheets.Add(
            After=source_sheet
        )

        pivot_sheet.Name = "PIVOT TABLE"

        # ====================================================
        # CREATE PIVOT CACHE
        # ====================================================

        if table is not None:

            pivot_cache = workbook.PivotCaches().Create(
                SourceType=1,
                SourceData="POPendingData"
            )

        else:

            source_address = source_range.Address(
                True,
                True,
                1,
                True
            )

            pivot_cache = workbook.PivotCaches().Create(
                SourceType=1,
                SourceData=source_address
            )

        # ====================================================
        # CREATE PIVOT TABLE
        # ====================================================

        pivot_destination = pivot_sheet.Range(
            "A3"
        )

        pivot_table = pivot_cache.CreatePivotTable(
            TableDestination=pivot_destination,
            TableName="POPendingPivot"
        )

        # ====================================================
        # PIVOT BASIC SETTINGS
        # ====================================================

        try:

            pivot_table.ManualUpdate = True

        except Exception:

            pass

        try:

            pivot_table.InGridDropZones = False

        except Exception:

            pass

        try:

            pivot_table.DisplayFieldCaptions = True

        except Exception:

            pass

        # ====================================================
        # FILTER FIELDS
        # ====================================================

        for position, field_name in enumerate(
            filter_fields,
            start=1
        ):

            pf = pivot_table.PivotFields(
                field_name
            )

            # xlPageField
            pf.Orientation = 3
            pf.Position = position

        # ====================================================
        # ROW FIELDS
        # ====================================================

        for position, field_name in enumerate(
            row_fields,
            start=1
        ):

            pf = pivot_table.PivotFields(
                field_name
            )

            # xlRowField
            pf.Orientation = 1
            pf.Position = position

        # ====================================================
        # REMOVE ALL SUBTOTALS
        # ====================================================

        for field_name in row_fields:

            try:

                pf = pivot_table.PivotFields(
                    field_name
                )

                pf.Subtotals = (
                    False,
                    False,
                    False,
                    False,
                    False,
                    False,
                    False,
                    False,
                    False,
                    False,
                    False,
                    False
                )

            except Exception:

                pass

        # ====================================================
        # VENDOR NAME ONLY SUBTOTAL
        # ====================================================

        try:

            vendor_pf = pivot_table.PivotFields(
                "VENDOR NAME"
            )

            vendor_pf.Subtotals = (
                True,
                False,
                False,
                False,
                False,
                False,
                False,
                False,
                False,
                False,
                False,
                False
            )

        except Exception:

            pass

        # ====================================================
        # ADD VALUE FIELDS
        # ====================================================

        for position, field_name in enumerate(
            value_fields,
            start=1
        ):

            try:

                pf = pivot_table.PivotFields(
                    field_name
                )

                # xlDataField
                pf.Orientation = 4

                # xlSum
                pf.Function = -4157

                pf.Position = position

                # Keep original field caption
                try:

                    pf.Caption = field_name

                except Exception:

                    pass

                # Number format
                try:

                    pf.NumberFormat = "#,##0.00"

                except Exception:

                    pass

            except Exception as value_error:

                raise RuntimeError(
                    f"Unable to add value field "
                    f"'{field_name}' to PivotTable.\n\n"
                    f"Excel Error: {value_error}"
                )

        # ====================================================
        # TABULAR FORM
        # ====================================================

        try:

            # xlTabularRow
            pivot_table.RowAxisLayout(1)

        except Exception:

            pass

        # ====================================================
        # REPEAT ALL ITEM LABELS
        # ====================================================

        try:

            # xlRepeatLabels
            pivot_table.RepeatAllLabels(2)

        except Exception:

            pass

        # ====================================================
        # VALUES / DATA FIELDS AS COLUMNS
        # ====================================================

        try:

            data_pivot_field = pivot_table.DataPivotField

            # xlColumnField
            data_pivot_field.Orientation = 2

        except Exception:

            pass

        # ====================================================
        # ENABLE COLUMN GRAND TOTAL
        # ====================================================

        try:

            pivot_table.ColumnGrand = True

        except Exception:

            pass

        # ====================================================
        # ENABLE ROW GRAND TOTAL
        # ====================================================

        try:

            pivot_table.RowGrand = True

        except Exception:

            pass

        # ====================================================
        # REFRESH PIVOT
        # ====================================================

        try:

            pivot_table.ManualUpdate = False

        except Exception:

            pass

        try:

            pivot_table.RefreshTable()

        except Exception:

            pass

        # ====================================================
        # REFRESH ON FILE OPEN
        # ====================================================

        try:

            pivot_cache.RefreshOnFileOpen = True

        except Exception:

            pass

        try:

            pivot_cache.SaveData = True

        except Exception:

            pass

        try:

            pivot_cache.MissingItemsLimit = 0

        except Exception:

            pass

        # ====================================================
        # FREEZE SOURCE HEADER
        # ====================================================

        try:

            source_sheet.Activate()

            excel.ActiveWindow.FreezePanes = False

            source_sheet.Range("A2").Select()

            excel.ActiveWindow.FreezePanes = True

        except Exception:

            pass

        # ====================================================
        # FORMAT SOURCE SHEET
        # APTOS - SIZE 8
        # ====================================================

        try:

            source_used_range = source_sheet.UsedRange

            source_used_range.Font.Name = "Aptos"
            source_used_range.Font.Size = 8

            source_sheet.Columns.AutoFit()

        except Exception:

            pass

        # ====================================================
        # FORMAT PIVOT SHEET
        # APTOS - SIZE 8
        # ====================================================

        try:

            pivot_sheet.Activate()

            pivot_used_range = pivot_sheet.UsedRange

            pivot_used_range.Font.Name = "Aptos"
            pivot_used_range.Font.Size = 8

            pivot_sheet.Columns.AutoFit()

        except Exception:

            pass

        # ====================================================
        # FINAL FORMAT - ALL SHEETS
        # APTOS FONT / SIZE 8
        # ====================================================

        try:

            for ws in workbook.Worksheets:

                used_range = ws.UsedRange

                used_range.Font.Name = "Aptos"
                used_range.Font.Size = 8

        except Exception:

            pass

        # ====================================================
        # SAVE WORKBOOK
        # ====================================================

        workbook.Save()

        # ====================================================
        # CLOSE WORKBOOK
        # ====================================================

        workbook.Close(
            SaveChanges=True
        )

        workbook = None

        # ====================================================
        # QUIT EXCEL
        # ====================================================

        excel.Quit()

        excel = None

        # ====================================================
        # READ FINAL FILE
        # ====================================================

        with open(
            temp_file,
            "rb"
        ) as file:

            file_bytes = file.read()

        return file_bytes

    finally:

        # ====================================================
        # CLOSE WORKBOOK IF STILL OPEN
        # ====================================================

        try:

            if workbook is not None:

                workbook.Close(
                    SaveChanges=False
                )

        except Exception:

            pass

        # ====================================================
        # QUIT EXCEL IF STILL RUNNING
        # ====================================================

        try:

            if excel is not None:

                excel.Quit()

        except Exception:

            pass

        # ====================================================
        # DELETE TEMP FILE
        # ====================================================

        try:

            if os.path.exists(temp_file):

                os.remove(temp_file)

        except Exception:

            pass

        try:

            if os.path.exists(temp_dir):

                os.rmdir(temp_dir)

        except Exception:

            pass


# ============================================================
# SIDEBAR
# ============================================================

st.sidebar.title(
    "📁 FILE UPLOAD"
)


# ------------------------------------------------------------
# BARCODE MASTER
# ------------------------------------------------------------

if not BARCODE_MASTER_FILE.exists():

    st.sidebar.error(
        "BARCODE STATUS.xlsx not found."
    )

    st.info(
        "Place BARCODE STATUS.xlsx in the same folder as app.py."
    )

    st.stop()


# ------------------------------------------------------------
# UPLOADERS
# ------------------------------------------------------------

po_file = st.sidebar.file_uploader(
    "Upload PO Pending File",
    type=["xlsx", "xls"]
)

schedule_file = st.sidebar.file_uploader(
    "Upload Schedule Date File",
    type=["xlsx", "xls"]
)


if po_file is None or schedule_file is None:

    st.info(
        "Please upload both PO Pending File and Schedule Date File."
    )

    st.stop()


# ============================================================
# MAIN
# ============================================================

try:

    # ========================================================
    # READ FILES
    # ========================================================

    with st.spinner(
        "Reading PO Pending File..."
    ):

        po_df = read_po_file(
            po_file
        )

    with st.spinner(
        "Reading Schedule Date File..."
    ):

        schedule_df = read_schedule_file(
            schedule_file
        )

    with st.spinner(
        "Reading Barcode Master..."
    ):

        barcode_df = read_barcode_master(
            BARCODE_MASTER_FILE
        )

    # ========================================================
    # FINAL DATA
    # ========================================================

    with st.spinner(
        "Preparing PO Pending Data..."
    ):

        final_df = create_final_data(
            po_df,
            schedule_df,
            barcode_df
        )

    if final_df.empty:

        st.warning(
            "No PO Pending data available after removing "
            "rows where Pending Qty is zero."
        )

        st.stop()

    # ========================================================
    # FILTERS
    # ========================================================

    st.sidebar.markdown("---")

    st.sidebar.subheader(
        "🔎 FILTERS"
    )

    filter_columns = [

        "DIVISION",
        "SECTION",
        "DEPARTMENT",
        "ARTCLE NAME",
        "BRAND",
        "VENDOR NAME",
        "DELIVERY MNTH",
        "BARCODE",
        "BRCD STATUS",
        "DELAY GROUP",
        "DELAY GROUP2",
        "ATTRIBUTE1",
        "CAT-1"

    ]

    # Automatically add filters for newly introduced input columns.
    filter_columns = list(dict.fromkeys(filter_columns + list(final_df.columns)))
    filtered_df = final_df.copy()

    for column in filter_columns:

        if column not in filtered_df.columns:

            continue

        options = (
            filtered_df[column]
            .dropna()
            .astype(str)
            .str.strip()
        )

        options = sorted(
            [
                x
                for x in options.unique()
                if x != ""
            ]
        )

        selected = st.sidebar.multiselect(
            column,
            options,
            key=f"filter_{column}"
        )

        if selected:

            filtered_df = filtered_df[
                filtered_df[column]
                .astype(str)
                .isin(selected)
            ]

    # ========================================================
    # CLEAR FILTERS
    # ========================================================

    if st.sidebar.button(
        "🔄 CLEAR ALL FILTERS"
    ):

        keys_to_delete = [
            key
            for key in st.session_state.keys()
            if key.startswith("filter_")
        ]

        for key in keys_to_delete:

            del st.session_state[key]

        st.rerun()

    # ========================================================
    # KPI
    # ========================================================

    st.markdown(
        "### 📌 PO Pending Summary"
    )

    col1, col2, col3 = st.columns(3)
    col4, col5, col6 = st.columns(3)

    with col1:

        st.metric(
            "ORDER QTY",
            f"{filtered_df['ORDER QTY'].sum():,.0f}"
        )

    with col2:

        st.metric(
            "RECEIVED QTY",
            f"{filtered_df['RECEIVED QTY'].sum():,.0f}"
        )

    with col3:

        st.metric(
            "CANCELLED QTY",
            f"{filtered_df['CANCELLED QTY'].sum():,.0f}"
        )

    with col4:

        st.metric(
            "PENDING QTY",
            f"{filtered_df['PENDING QTY'].sum():,.0f}"
        )

    with col5:

        st.metric(
            "PENDING VALUE",
            f"₹{filtered_df['PENDING VAL'].sum():,.0f}"
        )

    with col6:

        st.metric(
            "PACKET QTY",
            f"{filtered_df['PKT QTY'].sum():,.2f}"
        )

    # ========================================================
    # HIERARCHY SUMMARY
    # ========================================================

    st.markdown("---")

    st.markdown(
        "### 📊 Hierarchy Summary"
    )

    hierarchy_columns = [

        "DIVISION",
        "SECTION",
        "DEPARTMENT",
        "ARTCLE NAME",
        "BRAND",
        "DELIVERY MNTH",
        "BARCODE",
        "BRCD STATUS",
        "ORDER RATE",
        "MRP",
        "CAT-1"

    ]

    hierarchy_selected = st.selectbox(
        "Select Summary Level",
        hierarchy_columns
    )

    hierarchy_df = create_hierarchy_summary(
        filtered_df,
        hierarchy_selected
    )

    st.dataframe(
        hierarchy_df,
        use_container_width=True,
        height=350
    )

    # ========================================================
    # DRILL DOWN
    # ========================================================

    st.markdown("---")

    st.markdown(
        "### 🔍 Drill Down Analysis"
    )

    drill_levels = [
        "DIVISION",
        "SECTION",
        "DEPARTMENT",
        "ARTCLE NAME",
        "BRAND"
    ]

    available_drill = [
        col
        for col in drill_levels
        if col in filtered_df.columns
    ]

    if available_drill:

        drill_level = st.selectbox(
            "Select Drill Level",
            available_drill
        )

        drill_summary = create_hierarchy_summary(
            filtered_df,
            drill_level
        )

        st.dataframe(
            drill_summary,
            use_container_width=True,
            height=300
        )

    # ========================================================
    # FINAL DATA
    # ========================================================

    st.markdown("---")

    st.markdown(
        "### 📋 Final PO Pending Data"
    )

    st.caption(
        f"Showing {len(filtered_df):,} rows"
    )

    st.dataframe(
        filtered_df,
        use_container_width=True,
        height=500
    )

    # Reset the prepared Pivot download when source data or filters change.
    import hashlib
    signature = hashlib.sha256(
        pd.util.hash_pandas_object(filtered_df, index=True).values.tobytes()
        + "|".join(filtered_df.columns).encode("utf-8")
    ).hexdigest()
    if st.session_state.get("pivot_signature") != signature:
        st.session_state["pivot_signature"] = signature
        st.session_state["pivot_ready"] = False
        st.session_state["pivot_excel"] = None

    # ========================================================
    # DOWNLOADS
    # ========================================================

    st.markdown("---")

    st.markdown(
        "### ⬇️ DOWNLOAD"
    )

    # ========================================================
    # DOWNLOAD 1
    # ========================================================

    final_excel = create_final_excel(
        filtered_df
    )

    st.download_button(
        label="⬇️ DOWNLOAD FINAL PO PENDING",
        data=final_excel,
        file_name="FINAL_PO_PENDING.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
        use_container_width=True
    )

    # ========================================================
    # DOWNLOAD 2 - LIVE PIVOT
    # ========================================================

    st.markdown(
        "### 📊 LIVE PIVOT DOWNLOAD"
    )

    if st.button(
        "⚙️ PREPARE LIVE PIVOT DOWNLOAD",
        use_container_width=True
    ):

        with st.spinner(
            "Creating Live Excel PivotTable..."
        ):

            pivot_bytes = create_pivot_excel(
                filtered_df
            )

        st.session_state[
            "pivot_excel"
        ] = pivot_bytes

        st.session_state[
            "pivot_ready"
        ] = True

        st.success(
            "Live PivotTable Excel file is ready."
        )

    if st.session_state.get(
        "pivot_ready",
        False
    ):

        st.download_button(
            label="📊 DOWNLOAD PO PENDING + LIVE PIVOT",
            data=st.session_state[
                "pivot_excel"
            ],
            file_name="PO_PENDING_WITH_LIVE_PIVOT.xlsx",
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
            use_container_width=True
        )


# ============================================================
# ERROR HANDLER
# ============================================================

except Exception as error:

    st.error(
        "An error occurred while processing the files."
    )

    st.exception(error)