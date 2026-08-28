import streamlit as st
import pandas as pd
import plotly.express as px
from io import BytesIO
import re

# =========================================================
# PAGE CONFIGURATION
# =========================================================

st.set_page_config(
    page_title="GST Reconciliation Pro",
    page_icon="🧾",
    layout="wide"
)

# =========================================================
# PAGE HEADER
# =========================================================

st.title("🧾 GST Reconciliation Pro")
st.caption(
    "GST 2B vs Books • ITC Reconciliation • Invoice Exception Analysis"
)

st.info(
    "📢 Automated GSTR-2B vs Books Reconciliation • "
    "Fast & Accurate ITC Matching Engine"
)

# =========================================================
# HELPER FUNCTIONS
# =========================================================

def clean_text(value):
    if pd.isna(value):
        return ""
    return str(value).strip()


def clean_gstin(value):
    if pd.isna(value):
        return ""
    return str(value).upper().replace(" ", "").strip()


def clean_invoice(value):
    if pd.isna(value):
        return ""

    value = str(value).upper().strip()

    if value.endswith(".0"):
        value = value[:-2]

    return re.sub(r"[^A-Z0-9]", "", value)


def clean_amount(value):
    if pd.isna(value):
        return 0.0

    try:
        text = str(value)

        text = (
            text
            .replace(",", "")
            .replace("₹", "")
            .replace("Rs.", "")
            .replace("Rs", "")
            .strip()
        )

        return float(text)

    except Exception:
        return 0.0


# =========================================================
# COLUMN ALIASES
# =========================================================

ALIASES = {
    "GSTIN": [
        "GSTIN",
        "GSTIN/UIN",
        "GSTIN of Supplier",
        "GSTIN of supplier",
        "Supplier GSTIN",
        "GST Number",
        "GST No",
        "GST No.",
        "GSTIN of Vendor",
        "GSTIN/UIN of Supplier"
    ],

    "Invoice": [
        "Invoice Number",
        "Invoice number",
        "Invoice No",
        "Invoice No.",
        "Document Number",
        "Document No",
        "Bill No",
        "Bill Number",
        "Invoice"
    ],

    "Date": [
        "Invoice Date",
        "Invoice date",
        "Document Date",
        "Document date",
        "Bill Date",
        "Date"
    ],

    "Party": [
        "Party Name",
        "Supplier Name",
        "Supplier",
        "Vendor Name",
        "Vendor",
        "Trade Name",
        "Legal Name",
        "Party",
        "Supplier Trade Name"
    ],

    "Taxable": [
        "Taxable Value",
        "Taxable value",
        "Taxable Amount",
        "Taxable Amount (₹)",
        "Taxable Amt",
        "Taxable"
    ],

    "IGST": [
        "IGST",
        "IGST Amount",
        "IGST Amt",
        "Integrated Tax",
        "Integrated Tax Amount",
        "Integrated Tax Amount (₹)"
    ],

    "CGST": [
        "CGST",
        "CGST Amount",
        "CGST Amt",
        "Central Tax",
        "Central Tax Amount",
        "Central Tax Amount (₹)"
    ],

    "SGST": [
        "SGST",
        "SGST Amount",
        "SGST Amt",
        "UTGST",
        "State Tax",
        "State Tax Amount",
        "State/UT Tax",
        "State/UT Tax Amount",
        "State/UT Tax Amount (₹)"
    ],

    "InvoiceValue": [
        "Invoice Value",
        "Invoice value",
        "Total Invoice Value",
        "Total Value",
        "Document Value",
        "Invoice Amount",
        "Total Invoice Amount",
        "Invoice Value (₹)"
    ]
}


def normalize_column(value):
    return re.sub(
        r"[^A-Z0-9]",
        "",
        str(value).upper()
    )


def find_column(df, aliases):

    columns = {
        normalize_column(c): c
        for c in df.columns
    }

    # Exact match
    for alias in aliases:
        key = normalize_column(alias)

        if key in columns:
            return columns[key]

    # Partial match
    for column in df.columns:

        current = normalize_column(column)

        for alias in aliases:

            target = normalize_column(alias)

            if target in current or current in target:
                return column

    return None


# =========================================================
# READ EXCEL
# =========================================================

def read_excel(file):

    excel = pd.ExcelFile(file)

    all_data = []
    valid_sheets = []

    for sheet in excel.sheet_names:

        try:

            df = pd.read_excel(
                file,
                sheet_name=sheet
            )

            if df.empty:
                continue

            gstin_column = find_column(
                df,
                ALIASES["GSTIN"]
            )

            invoice_column = find_column(
                df,
                ALIASES["Invoice"]
            )

            if gstin_column and invoice_column:

                df["_SOURCE_SHEET"] = sheet

                all_data.append(df)
                valid_sheets.append(sheet)

        except Exception:
            continue

    if not all_data:

        raise ValueError(
            "GSTIN aur Invoice Number wale columns nahi mile."
        )

    return (
        pd.concat(
            all_data,
            ignore_index=True
        ),
        valid_sheets
    )


# =========================================================
# STANDARDIZE
# =========================================================

def standardize(df):

    result = pd.DataFrame()

    detected = {
        key: find_column(
            df,
            ALIASES[key]
        )
        for key in ALIASES
    }

    # GSTIN
    if detected["GSTIN"]:
        result["GSTIN"] = (
            df[detected["GSTIN"]]
            .apply(clean_gstin)
        )
    else:
        result["GSTIN"] = ""

    # Invoice
    if detected["Invoice"]:
        result["Invoice Number"] = (
            df[detected["Invoice"]]
            .apply(clean_invoice)
        )
    else:
        result["Invoice Number"] = ""

    # Date
    if detected["Date"]:
        result["Invoice Date"] = pd.to_datetime(
            df[detected["Date"]],
            errors="coerce"
        )
    else:
        result["Invoice Date"] = pd.NaT

    # Party
    if detected["Party"]:
        result["Party Name"] = (
            df[detected["Party"]]
            .apply(clean_text)
        )
    else:
        result["Party Name"] = ""

    # Taxable
    if detected["Taxable"]:
        result["Taxable Value"] = (
            df[detected["Taxable"]]
            .apply(clean_amount)
        )
    else:
        result["Taxable Value"] = 0.0

    # IGST
    if detected["IGST"]:
        result["IGST"] = (
            df[detected["IGST"]]
            .apply(clean_amount)
        )
    else:
        result["IGST"] = 0.0

    # CGST
    if detected["CGST"]:
        result["CGST"] = (
            df[detected["CGST"]]
            .apply(clean_amount)
        )
    else:
        result["CGST"] = 0.0

    # SGST
    if detected["SGST"]:
        result["SGST"] = (
            df[detected["SGST"]]
            .apply(clean_amount)
        )
    else:
        result["SGST"] = 0.0

    # Invoice value
    if detected["InvoiceValue"]:

        result["Invoice Value"] = (
            df[detected["InvoiceValue"]]
            .apply(clean_amount)
        )

    else:

        result["Invoice Value"] = (
            result["Taxable Value"]
            + result["IGST"]
            + result["CGST"]
            + result["SGST"]
        )

    # Source sheet
    if "_SOURCE_SHEET" in df.columns:
        result["Source Sheet"] = df["_SOURCE_SHEET"]
    else:
        result["Source Sheet"] = ""

    # Remove blank rows
    result = result[
        (result["GSTIN"] != "")
        &
        (result["Invoice Number"] != "")
    ].copy()

    return result


# =========================================================
# RECONCILIATION
# =========================================================

def reconcile(two_b, books, tolerance):

    two_b = two_b.copy()
    books = books.copy()

    two_b["KEY"] = (
        two_b["GSTIN"]
        + "|"
        + two_b["Invoice Number"]
    )

    books["KEY"] = (
        books["GSTIN"]
        + "|"
        + books["Invoice Number"]
    )

    result = pd.merge(
        books,
        two_b,
        on="KEY",
        how="outer",
        suffixes=("_Books", "_2B"),
        indicator=True
    )

    statuses = []
    differences = []

    for _, row in result.iterrows():

        # -------------------------------------------------
        # Missing in 2B
        # -------------------------------------------------

        if row["_merge"] == "left_only":

            statuses.append("Missing in 2B")

            differences.append(
                row["IGST_Books"]
                + row["CGST_Books"]
                + row["SGST_Books"]
            )

        # -------------------------------------------------
        # Missing in Books
        # -------------------------------------------------

        elif row["_merge"] == "right_only":

            statuses.append("Missing in Books")

            differences.append(
                row["IGST_2B"]
                + row["CGST_2B"]
                + row["SGST_2B"]
            )

        # -------------------------------------------------
        # Present in Both
        # -------------------------------------------------

        else:

            taxable_diff = abs(
                row["Taxable Value_Books"]
                -
                row["Taxable Value_2B"]
            )

            igst_diff = abs(
                row["IGST_Books"]
                -
                row["IGST_2B"]
            )

            cgst_diff = abs(
                row["CGST_Books"]
                -
                row["CGST_2B"]
            )

            sgst_diff = abs(
                row["SGST_Books"]
                -
                row["SGST_2B"]
            )

            invoice_diff = abs(
                row["Invoice Value_Books"]
                -
                row["Invoice Value_2B"]
            )

            total_tax_difference = (
                igst_diff
                + cgst_diff
                + sgst_diff
            )

            if (
                taxable_diff <= tolerance
                and igst_diff <= tolerance
                and cgst_diff <= tolerance
                and sgst_diff <= tolerance
                and invoice_diff <= tolerance
            ):

                statuses.append("Matched")

            else:

                statuses.append("Value Mismatch")

            differences.append(
                total_tax_difference
            )

    result["Status"] = statuses
    result["ITC Difference"] = differences

    return result


# =========================================================
# PREPARE DISPLAY
# =========================================================

def prepare_display(df):

    output = pd.DataFrame()

    def get_col(name):

        if name in df.columns:

            return df[name]

        return pd.Series(
            [""] * len(df),
            index=df.index
        )

    output["GSTIN"] = (
        get_col("GSTIN_Books")
        .where(
            get_col("GSTIN_Books") != "",
            get_col("GSTIN_2B")
        )
    )

    output["Party Name"] = (
        get_col("Party Name_Books")
        .where(
            get_col("Party Name_Books") != "",
            get_col("Party Name_2B")
        )
    )

    output["Invoice Number"] = (
        get_col("Invoice Number_Books")
        .where(
            get_col("Invoice Number_Books") != "",
            get_col("Invoice Number_2B")
        )
    )

    output["Invoice Date"] = (
        get_col("Invoice Date_Books")
        .where(
            get_col("Invoice Date_Books").notna(),
            get_col("Invoice Date_2B")
        )
    )

    output["Taxable - Books"] = get_col(
        "Taxable Value_Books"
    )

    output["Taxable - 2B"] = get_col(
        "Taxable Value_2B"
    )

    output["IGST - Books"] = get_col(
        "IGST_Books"
    )

    output["IGST - 2B"] = get_col(
        "IGST_2B"
    )

    output["CGST - Books"] = get_col(
        "CGST_Books"
    )

    output["CGST - 2B"] = get_col(
        "CGST_2B"
    )

    output["SGST - Books"] = get_col(
        "SGST_Books"
    )

    output["SGST - 2B"] = get_col(
        "SGST_2B"
    )

    output["Invoice Value - Books"] = get_col(
        "Invoice Value_Books"
    )

    output["Invoice Value - 2B"] = get_col(
        "Invoice Value_2B"
    )

    output["ITC Difference"] = get_col(
        "ITC Difference"
    )

    output["Status"] = get_col(
        "Status"
    )

    return output.reset_index(drop=True)


# =========================================================
# CREATE EXCEL
# =========================================================

def create_excel(result):

    output = BytesIO()

    display = prepare_display(result)

    with pd.ExcelWriter(
        output,
        engine="openpyxl"
    ) as writer:

        display.to_excel(
            writer,
            sheet_name="Complete Reconciliation",
            index=False
        )

        for status in [
            "Matched",
            "Missing in 2B",
            "Missing in Books",
            "Value Mismatch"
        ]:

            temp = display[
                display["Status"] == status
            ]

            temp.to_excel(
                writer,
                sheet_name=status[:31],
                index=False
            )

    output.seek(0)

    return output


# =========================================================
# SIDEBAR
# =========================================================

with st.sidebar:

    st.header("⚙️ Reconciliation Settings")

    tolerance = st.number_input(
        "Mismatch Tolerance ₹",
        min_value=0.0,
        value=2.0,
        step=0.50
    )

    st.divider()

    st.subheader("🔑 Matching Key")

    st.write(
        "GSTIN + Invoice Number"
    )

    st.divider()

    st.subheader("📌 Status Categories")

    st.write(
        "🟢 Matched — Exact match within tolerance"
    )

    st.write(
        "🔴 Missing in 2B — Present in Books only"
    )

    st.write(
        "🟠 Missing in Books — Present in 2B only"
    )

    st.write(
        "🟣 Value Mismatch — Amount difference"
    )


# =========================================================
# FILE UPLOAD
# =========================================================

st.header("📂 Upload GST Data")

col1, col2 = st.columns(2)

with col1:

    st.subheader("📥 GST 2B File")

    st.caption(
        "Upload GST portal 2B Excel file. Multi-sheet supported."
    )

    two_b_file = st.file_uploader(
        "Choose GST 2B Excel",
        type=["xlsx", "xls"],
        key="gst_2b_upload"
    )


with col2:

    st.subheader("📚 Books File")

    st.caption(
        "Upload Purchase Register / Accounting Books Excel."
    )

    books_file = st.file_uploader(
        "Choose Books Excel",
        type=["xlsx", "xls"],
        key="books_upload"
    )


# =========================================================
# CLEAR OLD RESULT WHEN FILES ARE MISSING
# =========================================================

if not two_b_file or not books_file:

    st.session_state.pop(
        "result",
        None
    )

    st.session_state.pop(
        "selected_status",
        None
    )

    st.info(
        "👋 Welcome to GST Reconciliation Pro"
    )

    st.write(
        "Upload your GSTR-2B and Purchase Register Excel files above "
        "to reconcile your ITC automatically."
    )

    st.write(
        "🟢 Matched — Invoices matching in both files within tolerance."
    )

    st.write(
        "🔴 Missing in 2B — Invoices recorded in Books "
        "but not filed by vendor in 2B."
    )

    st.write(
        "🟠 Missing in Books — Invoices in 2B "
        "but missed in your accounting register."
    )

    st.write(
        "🟣 Value Mismatch — Differences in Taxable, "
        "IGST, CGST, or SGST values."
    )


# =========================================================
# START RECONCILIATION
# =========================================================

if two_b_file and books_file:

    if st.button(
        "🚀 START GST RECONCILIATION",
        type="primary",
        use_container_width=True
    ):

        try:

            with st.spinner(
                "Analysing GST 2B and Books..."
            ):

                # Read 2B
                two_b_raw, two_b_sheets = read_excel(
                    two_b_file
                )

                # Read Books
                books_raw, books_sheets = read_excel(
                    books_file
                )

                # Standardize
                two_b = standardize(
                    two_b_raw
                )

                books = standardize(
                    books_raw
                )

                if two_b.empty:

                    raise ValueError(
                        "GST 2B file mein valid invoice data nahi mila."
                    )

                if books.empty:

                    raise ValueError(
                        "Books file mein valid invoice data nahi mila."
                    )

                # Reconciliation
                result = reconcile(
                    two_b,
                    books,
                    tolerance
                )

                # Save in session
                st.session_state["result"] = result

                st.session_state["two_b"] = two_b

                st.session_state["books"] = books

                st.session_state[
                    "selected_status"
                ] = "All Records"

            st.success(
                "✅ GST Reconciliation completed successfully."
            )

        except Exception as error:

            st.error(
                f"❌ File processing error: {error}"
            )


# =========================================================
# DASHBOARD
# ONLY SHOW AFTER SUCCESSFUL RECONCILIATION
# =========================================================

if (
    two_b_file
    and books_file
    and "result" in st.session_state
):

    result = st.session_state["result"]

    # =====================================================
    # COUNTS
    # =====================================================

    total = len(result)

    matched = int(
        (
            result["Status"]
            ==
            "Matched"
        ).sum()
    )

    missing_2b = int(
        (
            result["Status"]
            ==
            "Missing in 2B"
        ).sum()
    )

    missing_books = int(
        (
            result["Status"]
            ==
            "Missing in Books"
        ).sum()
    )

    mismatch = int(
        (
            result["Status"]
            ==
            "Value Mismatch"
        ).sum()
    )

    # =====================================================
    # DASHBOARD
    # =====================================================

    st.divider()

    st.header("📊 Reconciliation Dashboard")

    st.caption(
        "Click any category below to view its invoice details."
    )

    # =====================================================
    # KPI CARDS
    # =====================================================

    c1, c2, c3, c4, c5 = st.columns(5)

    with c1:

        st.metric(
            label="📊 TOTAL RECORDS",
            value=f"{total:,}"
        )

        if st.button(
            "View Total Records",
            key="btn_total",
            use_container_width=True
        ):

            st.session_state[
                "selected_status"
            ] = "All Records"

    with c2:

        st.metric(
            label="🟢 MATCHED",
            value=f"{matched:,}"
        )

        if st.button(
            "View Matched",
            key="btn_matched",
            use_container_width=True
        ):

            st.session_state[
                "selected_status"
            ] = "Matched"

    with c3:

        st.metric(
            label="🔴 MISSING IN 2B",
            value=f"{missing_2b:,}"
        )

        if st.button(
            "View Missing in 2B",
            key="btn_missing_2b",
            use_container_width=True
        ):

            st.session_state[
                "selected_status"
            ] = "Missing in 2B"

    with c4:

        st.metric(
            label="🟠 MISSING IN BOOKS",
            value=f"{missing_books:,}"
        )

        if st.button(
            "View Missing in Books",
            key="btn_missing_books",
            use_container_width=True
        ):

            st.session_state[
                "selected_status"
            ] = "Missing in Books"

    with c5:

        st.metric(
            label="🟣 VALUE MISMATCH",
            value=f"{mismatch:,}"
        )

        if st.button(
            "View Value Mismatch",
            key="btn_mismatch",
            use_container_width=True
        ):

            st.session_state[
                "selected_status"
            ] = "Value Mismatch"

    # =====================================================
    # CHARTS
    # =====================================================

    st.subheader("📈 Reconciliation Summary")

    ch1, ch2 = st.columns([1, 2])

    with ch1:

        fig_pie = px.pie(

            names=[
                "Matched",
                "Missing in 2B",
                "Missing in Books",
                "Value Mismatch"
            ],

            values=[
                matched,
                missing_2b,
                missing_books,
                mismatch
            ],

            color=[
                "Matched",
                "Missing in 2B",
                "Missing in Books",
                "Value Mismatch"
            ],

            color_discrete_map={
                "Matched": "#16a34a",
                "Missing in 2B": "#dc2626",
                "Missing in Books": "#ea580c",
                "Value Mismatch": "#7c3aed"
            },

            hole=0.45
        )

        fig_pie.update_layout(
            margin=dict(
                t=20,
                b=20,
                l=10,
                r=10
            ),
            height=320
        )

        st.plotly_chart(
            fig_pie,
            use_container_width=True
        )

    # =====================================================
    # DETAILS
    # =====================================================

    display_df = prepare_display(
        result
    )

    if "selected_status" not in st.session_state:

        st.session_state[
            "selected_status"
        ] = "All Records"

    selected_status = st.session_state[
        "selected_status"
    ]

    if selected_status == "All Records":

        filtered_df = display_df

        heading = (
            "📋 All Reconciliation Records"
        )

    else:

        filtered_df = display_df[
            display_df["Status"]
            ==
            selected_status
        ]

        heading = (
            f"📌 {selected_status} - Invoice Details"
        )

    st.divider()

    st.subheader(
        heading
    )

    st.caption(
        f"{len(filtered_df):,} invoice(s) found"
    )

    if filtered_df.empty:

        st.warning(
            "Is category mein koi invoice nahi mila."
        )

    else:

        st.dataframe(
            filtered_df,
            use_container_width=True,
            height=350,
            hide_index=True
        )

    # =====================================================
    # DOWNLOAD
    # =====================================================

    st.divider()

    excel_file = create_excel(
        result
    )

    st.download_button(
        label="📥 Download Full Reconciliation Report (Excel)",
        data=excel_file,
        file_name="GST_Reconciliation_Report.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
        use_container_width=True
    )
