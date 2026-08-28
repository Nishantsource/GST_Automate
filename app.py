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
# CSS & STYLES
# =========================================================

st.markdown("""
<style>

.stApp {
    background-color: #f4f7fb;
    color: #1e293b;
}

[data-testid="stHeader"] {
    background-color: transparent;
}

/* News Ticker Headline */
.ticker-wrap {
    width: 100%;
    background-color: #0f2a5f;
    color: #f8fafc;
    padding: 10px 0;
    overflow: hidden;
    border-radius: 8px;
    margin-bottom: 20px;
    box-shadow: 0 2px 10px rgba(0,0,0,0.1);
}

.ticker {
    display: inline-block;
    white-space: nowrap;
    padding-left: 100%;
    animation: marquee 20s linear infinite;
    font-size: 15px;
    font-weight: 600;
    letter-spacing: 0.5px;
}

@keyframes marquee {
    0% {
        transform: translate(0, 0);
    }

    100% {
        transform: translate(-100%, 0);
    }
}

/* Hero */
.hero {
    background: linear-gradient(135deg, #0f2a5f, #1769aa);
    padding: 25px 30px;
    border-radius: 16px;
    margin-bottom: 25px;
}

.hero h1 {
    color: #ffffff !important;
    font-size: 30px;
    font-weight: 800;
    margin: 0;
}

.hero p {
    color: #dbeafe !important;
    font-size: 14px;
    margin-top: 6px;
    margin-bottom: 0;
}

.section-title {
    color: #0f2a5f;
    font-size: 20px;
    font-weight: 800;
    margin-top: 20px;
    margin-bottom: 12px;
}

/* Original KPI Look */
.kpi-button button {
    width: 100%;
    min-height: 110px;
    background: #ffffff !important;
    color: #0f172a !important;
    border: 1px solid #dbe3ef !important;
    border-radius: 12px !important;
    box-shadow: 0 4px 12px rgba(15,42,95,.05);
    transition: all 0.2s ease;
    text-align: left !important;
    padding: 15px 16px !important;
    white-space: pre-line !important;
}

.kpi-button button:hover {
    transform: translateY(-3px);
    box-shadow: 0 8px 20px rgba(15,42,95,.12);
    border-color: #2563eb !important;
}

/* Colored top border for each clickable card */
.kpi-blue button {
    border-top: 4px solid #2563eb !important;
}

.kpi-green button {
    border-top: 4px solid #16a34a !important;
}

.kpi-red button {
    border-top: 4px solid #dc2626 !important;
}

.kpi-orange button {
    border-top: 4px solid #ea580c !important;
}

.kpi-purple button {
    border-top: 4px solid #7c3aed !important;
}

/* Welcome box */
.welcome-box {
    background: #ffffff;
    padding: 18px 22px;
    border-radius: 12px;
    border-left: 5px solid #2563eb;
    border-top: 1px solid #e2e8f0;
    border-right: 1px solid #e2e8f0;
    border-bottom: 1px solid #e2e8f0;
    margin: 15px 0 25px 0;
    box-shadow: 0 4px 12px rgba(15,42,95,.05);
}

.welcome-title {
    font-size: 16px;
    font-weight: 800;
    color: #0f2a5f;
    margin-bottom: 8px;
}

.welcome-text {
    color: #334155;
    font-size: 13.5px;
    line-height: 1.7;
}

/* Details box */
.details-box {
    background: #ffffff;
    padding: 18px 22px;
    border-radius: 12px;
    border-left: 5px solid #2563eb;
    border-top: 1px solid #dbe3ef;
    border-right: 1px solid #dbe3ef;
    border-bottom: 1px solid #dbe3ef;
    margin-top: 18px;
    box-shadow: 0 4px 12px rgba(15,42,95,.05);
}

.details-title {
    color: #0f2a5f;
    font-size: 19px;
    font-weight: 800;
    margin-bottom: 5px;
}

/* File uploader */
[data-testid="stFileUploaderDropzone"] {
    background: #ffffff;
    border: 2px dashed #94a3b8;
    border-radius: 12px;
}

/* Download */
.stDownloadButton button {
    background: #0f2a5f !important;
    color: white !important;
    border-radius: 10px !important;
    font-weight: 700 !important;
}

/* Start button */
.stButton > button[kind="primary"] {
    border-radius: 10px !important;
    font-weight: 800 !important;
}

</style>
""", unsafe_allow_html=True)

# =========================================================
# RUNNING TICKER & HERO HEADER
# =========================================================

st.markdown("""
<div class="ticker-wrap">
    <div class="ticker">
        📢 Welcome to GST RECONCILIATION PRO • Automated GSTR-2B vs Books Reconciliation • Fast & Accurate ITC Matching Engine • Reconcile Invoices with One Click
    </div>
</div>
""", unsafe_allow_html=True)

st.markdown("""
<div class="hero">
    <h1>🧾 GST Reconciliation Pro</h1>
    <p>GST 2B vs Books • ITC Reconciliation • Invoice Exception Analysis</p>
</div>
""", unsafe_allow_html=True)

# =========================================================
# HELPERS
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

    for alias in aliases:

        key = normalize_column(alias)

        if key in columns:
            return columns[key]

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
        key: find_column(df, ALIASES[key])
        for key in ALIASES
    }

    result["GSTIN"] = (
        df[detected["GSTIN"]].apply(clean_gstin)
        if detected["GSTIN"]
        else ""
    )

    result["Invoice Number"] = (
        df[detected["Invoice"]].apply(clean_invoice)
        if detected["Invoice"]
        else ""
    )

    result["Invoice Date"] = (
        pd.to_datetime(
            df[detected["Date"]],
            errors="coerce"
        )
        if detected["Date"]
        else pd.NaT
    )

    result["Party Name"] = (
        df[detected["Party"]].apply(clean_text)
        if detected["Party"]
        else ""
    )

    result["Taxable Value"] = (
        df[detected["Taxable"]].apply(clean_amount)
        if detected["Taxable"]
        else 0.0
    )

    result["IGST"] = (
        df[detected["IGST"]].apply(clean_amount)
        if detected["IGST"]
        else 0.0
    )

    result["CGST"] = (
        df[detected["CGST"]].apply(clean_amount)
        if detected["CGST"]
        else 0.0
    )

    result["SGST"] = (
        df[detected["SGST"]].apply(clean_amount)
        if detected["SGST"]
        else 0.0
    )

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

    result["Source Sheet"] = (
        df["_SOURCE_SHEET"]
        if "_SOURCE_SHEET" in df.columns
        else ""
    )

    return result[
        (result["GSTIN"] != "")
        &
        (result["Invoice Number"] != "")
    ].copy()


# =========================================================
# RECONCILE
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

        if row["_merge"] == "left_only":

            statuses.append("Missing in 2B")

            differences.append(
                row["IGST_Books"]
                + row["CGST_Books"]
                + row["SGST_Books"]
            )

        elif row["_merge"] == "right_only":

            statuses.append("Missing in Books")

            differences.append(
                row["IGST_2B"]
                + row["CGST_2B"]
                + row["SGST_2B"]
            )

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

            inv_diff = abs(
                row["Invoice Value_Books"]
                -
                row["Invoice Value_2B"]
            )

            tot_diff = (
                igst_diff
                + cgst_diff
                + sgst_diff
            )

            if (
                taxable_diff <= tolerance
                and igst_diff <= tolerance
                and cgst_diff <= tolerance
                and sgst_diff <= tolerance
                and inv_diff <= tolerance
            ):

                statuses.append("Matched")

            else:

                statuses.append("Value Mismatch")

            differences.append(tot_diff)

    result["Status"] = statuses
    result["ITC Difference"] = differences

    return result


# =========================================================
# DISPLAY
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

    st.info(
        "GSTIN + Invoice Number"
    )

    st.divider()

    st.subheader("📌 Status Categories")

    st.write(
        "🟢 **Matched:** Exact match within tolerance"
    )

    st.write(
        "🔴 **Missing in 2B:** Present in Books only"
    )

    st.write(
        "🟠 **Missing in Books:** Present in 2B only"
    )

    st.write(
        "🟣 **Value Mismatch:** Amount difference"
    )


# =========================================================
# UPLOAD SECTION
# =========================================================

st.markdown(
    '<div class="section-title">📂 Upload GST Data</div>',
    unsafe_allow_html=True
)

col1, col2 = st.columns(2)

with col1:

    st.markdown("##### 📥 GST 2B File")

    st.caption(
        "Upload GST portal 2B Excel file (multi-sheet supported)."
    )

    two_b_file = st.file_uploader(
        "Choose GST 2B Excel",
        type=["xlsx", "xls"],
        key="gst_2b_upload",
        label_visibility="collapsed"
    )


with col2:

    st.markdown("##### 📚 Books File")

    st.caption(
        "Upload Purchase Register / Accounting Books Excel."
    )

    books_file = st.file_uploader(
        "Choose Books Excel",
        type=["xlsx", "xls"],
        key="books_upload",
        label_visibility="collapsed"
    )


# =========================================================
# FILE SIGNATURE
# Used so old dashboard does not appear for a new upload
# =========================================================

current_files = (
    two_b_file.name if two_b_file else None,
    books_file.name if books_file else None
)

saved_files = st.session_state.get(
    "processed_files",
    None
)


# =========================================================
# NO FILE / NOT PROCESSED
# =========================================================

if not two_b_file or not books_file:

    st.session_state.pop(
        "result",
        None
    )

    st.session_state.pop(
        "processed_files",
        None
    )

    st.session_state.pop(
        "selected_status",
        None
    )

    # HTML-free welcome section
    with st.container(border=True):

        st.subheader(
            "👋 Welcome to GST Reconciliation Pro"
        )

        st.write(
            "Upload your GSTR-2B and Purchase Register Excel "
            "files above to reconcile your ITC automatically."
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
            "🟣 Value Mismatch — Differences found in "
            "Taxable, IGST, CGST, or SGST values."
        )


# =========================================================
# START RECONCILIATION
# =========================================================

if two_b_file and books_file:

    st.write("")

    if st.button(
        "🚀 START GST RECONCILIATION",
        type="primary",
        use_container_width=True
    ):

        try:

            with st.spinner(
                "Analysing GST 2B and Books..."
            ):

                two_b_raw, two_b_sheets = read_excel(
                    two_b_file
                )

                books_raw, books_sheets = read_excel(
                    books_file
                )

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

                result = reconcile(
                    two_b,
                    books,
                    tolerance
                )

                st.session_state["result"] = result

                st.session_state["two_b"] = two_b

                st.session_state["books"] = books

                st.session_state["processed_files"] = (
                    two_b_file.name,
                    books_file.name
                )

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
# ONLY AFTER SUCCESSFUL RECONCILIATION
# =========================================================

if (
    two_b_file
    and books_file
    and "result" in st.session_state
    and saved_files == current_files
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
    # DASHBOARD TITLE
    # =====================================================

    st.markdown(
        '<div class="section-title">📊 Reconciliation Dashboard</div>',
        unsafe_allow_html=True
    )

    # =====================================================
    # CLICKABLE KPI CARDS
    # =====================================================

    c1, c2, c3, c4, c5 = st.columns(5)

    # TOTAL
    with c1:

        st.markdown(
            '<div class="kpi-button kpi-blue">',
            unsafe_allow_html=True
        )

        if st.button(
            f"📊 TOTAL RECORDS\n\n{total:,}\n\nAnalysed Invoices",
            key="card_total",
            use_container_width=True
        ):

            st.session_state[
                "selected_status"
            ] = "All Records"

        st.markdown(
            "</div>",
            unsafe_allow_html=True
        )

    # MATCHED
    with c2:

        st.markdown(
            '<div class="kpi-button kpi-green">',
            unsafe_allow_html=True
        )

        if st.button(
            f"🟢 MATCHED\n\n{matched:,}\n\nFully Reconciled",
            key="card_matched",
            use_container_width=True
        ):

            st.session_state[
                "selected_status"
            ] = "Matched"

        st.markdown(
            "</div>",
            unsafe_allow_html=True
        )

    # MISSING IN 2B
    with c3:

        st.markdown(
            '<div class="kpi-button kpi-red">',
            unsafe_allow_html=True
        )

        if st.button(
            f"🔴 MISSING IN 2B\n\n{missing_2b:,}\n\nITC at Risk",
            key="card_missing_2b",
            use_container_width=True
        ):

            st.session_state[
                "selected_status"
            ] = "Missing in 2B"

        st.markdown(
            "</div>",
            unsafe_allow_html=True
        )

    # MISSING IN BOOKS
    with c4:

        st.markdown(
            '<div class="kpi-button kpi-orange">',
            unsafe_allow_html=True
        )

        if st.button(
            f"🟠 MISSING IN BOOKS\n\n{missing_books:,}\n\nUnrecorded Invoices",
            key="card_missing_books",
            use_container_width=True
        ):

            st.session_state[
                "selected_status"
            ] = "Missing in Books"

        st.markdown(
            "</div>",
            unsafe_allow_html=True
        )

    # VALUE MISMATCH
    with c5:

        st.markdown(
            '<div class="kpi-button kpi-purple">',
            unsafe_allow_html=True
        )

        if st.button(
            f"🟣 VALUE MISMATCH\n\n{mismatch:,}\n\nAmount Discrepancies",
            key="card_mismatch",
            use_container_width=True
        ):

            st.session_state[
                "selected_status"
            ] = "Value Mismatch"

        st.markdown(
            "</div>",
            unsafe_allow_html=True
        )

    # =====================================================
    # CHARTS
    # =====================================================

    st.write("")

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
            height=300
        )

        st.plotly_chart(
            fig_pie,
            use_container_width=True
        )

    # =====================================================
    # DETAILS OF CLICKED CATEGORY
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

    # =====================================================
    # DETAILS BOX
    # =====================================================

    st.markdown(
        '<div class="details-box">',
        unsafe_allow_html=True
    )

    st.markdown(
        f'<div class="details-title">{heading}</div>',
        unsafe_allow_html=True
    )

    st.caption(
        f"{len(filtered_df):,} invoice(s) found"
    )

    st.markdown(
        "</div>",
        unsafe_allow_html=True
    )

    if filtered_df.empty:

        st.info(
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
