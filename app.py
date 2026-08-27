import streamlit as st
import pandas as pd
import plotly.express as px
from io import BytesIO
import re

# =========================================================
# PAGE CONFIG
# =========================================================

st.set_page_config(
    page_title="GST Reconciliation Pro",
    page_icon="🧾",
    layout="wide",
    initial_sidebar_state="expanded"
)

# =========================================================
# PROFESSIONAL GST SOFTWARE THEME
# =========================================================

st.markdown("""
<style>

html, body, [class*="css"] {
    font-family: Inter, Arial, sans-serif;
}

html {
    color-scheme: light !important;
}

.stApp {
    background: #f4f7fb !important;
    color: #172033 !important;
}

[data-testid="stAppViewContainer"] {
    background: #f4f7fb !important;
}

[data-testid="stHeader"] {
    background: transparent !important;
}

[data-testid="stMainBlockContainer"] {
    max-width: 1500px;
    padding-top: 25px;
}

/* ================= HEADER ================= */

.hero {
    background: linear-gradient(135deg, #0f2a5f 0%, #1769aa 100%);
    padding: 30px 34px;
    border-radius: 20px;
    margin-bottom: 25px;
    box-shadow: 0 12px 30px rgba(15,42,95,.18);
}

.hero h1 {
    color: white !important;
    font-size: 34px;
    margin: 0;
    font-weight: 800;
}

.hero p {
    color: #dbeafe !important;
    font-size: 15px;
    margin-top: 8px;
    margin-bottom: 0;
}

/* ================= SIDEBAR ================= */

[data-testid="stSidebar"] {
    background: #ffffff !important;
    border-right: 1px solid #e2e8f0;
}

[data-testid="stSidebar"] * {
    color: #172033 !important;
}

/* ================= SECTION ================= */

.section-title {
    font-size: 23px;
    font-weight: 800;
    color: #0f2a5f !important;
    margin-top: 28px;
    margin-bottom: 14px;
}

/* ================= UPLOAD CARDS ================= */

.upload-card {
    background: white;
    border: 1px solid #dbe3ef;
    border-radius: 18px;
    padding: 20px;
    margin-bottom: 15px;
    box-shadow: 0 5px 18px rgba(15,42,95,.06);
}

.upload-title {
    color: #0f2a5f !important;
    font-size: 19px;
    font-weight: 800;
    margin-bottom: 5px;
}

.upload-subtitle {
    color: #64748b !important;
    font-size: 13px;
    margin-bottom: 12px;
}

/* ================= KPI CARDS ================= */

.kpi {
    background: white;
    border: 1px solid #dbe3ef;
    border-radius: 16px;
    padding: 20px;
    min-height: 135px;
    box-shadow: 0 5px 18px rgba(15,42,95,.07);
    transition: all .2s ease;
}

.kpi:hover {
    transform: translateY(-4px);
    box-shadow: 0 12px 28px rgba(15,42,95,.13);
}

.kpi-label {
    color: #64748b !important;
    font-size: 12px;
    font-weight: 800;
    letter-spacing: .5px;
}

.kpi-value {
    color: #172033 !important;
    font-size: 30px;
    font-weight: 900;
    margin-top: 8px;
}

.kpi-description {
    color: #94a3b8 !important;
    font-size: 12px;
    margin-top: 4px;
}

.kpi-blue {
    border-top: 4px solid #2563eb;
}

.kpi-green {
    border-top: 4px solid #16a34a;
}

.kpi-red {
    border-top: 4px solid #dc2626;
}

.kpi-orange {
    border-top: 4px solid #ea580c;
}

.kpi-purple {
    border-top: 4px solid #7c3aed;
}

/* ================= BUTTONS ================= */

.stButton > button {
    border-radius: 11px !important;
    min-height: 46px !important;
    font-weight: 800 !important;
    border: 1px solid #d1d9e6 !important;
    background: white !important;
    color: #172033 !important;
}

.stButton > button:hover {
    border-color: #2563eb !important;
    color: #2563eb !important;
    box-shadow: 0 5px 15px rgba(37,99,235,.12);
}

.stButton > button[kind="primary"] {
    background: linear-gradient(135deg,#155eef,#1769aa) !important;
    color: white !important;
    border: none !important;
    min-height: 52px !important;
    font-size: 15px !important;
}

/* ================= DOWNLOAD ================= */

.stDownloadButton > button {
    background: #0f2a5f !important;
    color: white !important;
    border: none !important;
    border-radius: 11px !important;
    min-height: 50px !important;
    font-weight: 800 !important;
}

/* ================= FILE UPLOADER ================= */

[data-testid="stFileUploaderDropzone"] {
    background: #f8fbff !important;
    border: 2px dashed #b8c9df !important;
    border-radius: 14px !important;
}

[data-testid="stFileUploaderDropzone"] * {
    color: #172033 !important;
}

/* ================= INPUTS ================= */

input {
    color: #172033 !important;
    background: white !important;
}

[data-baseweb="select"] > div {
    background: white !important;
    color: #172033 !important;
}

/* ================= INFO BOX ================= */

.info-box {
    background: white;
    border: 1px solid #dbe3ef;
    border-left: 5px solid #2563eb;
    border-radius: 13px;
    padding: 18px;
    margin: 14px 0;
    color: #475569 !important;
    box-shadow: 0 4px 14px rgba(15,42,95,.05);
}

.info-box h3 {
    color: #0f2a5f !important;
    margin-top: 0;
}

/* ================= TABLE ================= */

.custom-table {
    width: 100%;
    border-collapse: collapse;
    background: white;
    border: 1px solid #dbe3ef;
    border-radius: 12px;
    overflow: hidden;
    font-size: 12px;
}

.custom-table th {
    background: #eef4fb;
    color: #0f2a5f !important;
    padding: 12px;
    text-align: left;
    font-weight: 800;
    border-bottom: 1px solid #dbe3ef;
    white-space: nowrap;
}

.custom-table td {
    background: white;
    color: #172033 !important;
    padding: 11px;
    border-bottom: 1px solid #edf1f6;
    white-space: nowrap;
}

.custom-table tr:hover td {
    background: #f7faff;
}

/* ================= EXPANDER ================= */

[data-testid="stExpander"] {
    background: white !important;
    border: 1px solid #dbe3ef !important;
    border-radius: 12px !important;
}

[data-testid="stExpander"] * {
    color: #172033 !important;
}

/* ================= MOBILE ================= */

@media (max-width: 768px) {

    .hero {
        padding: 22px;
    }

    .hero h1 {
        font-size: 25px;
    }

    .hero p {
        font-size: 13px;
    }

    .section-title {
        font-size: 20px;
    }

    .kpi {
        margin-bottom: 10px;
    }

}

</style>
""", unsafe_allow_html=True)


# =========================================================
# HELPERS
# =========================================================

def clean_text(x):
    if pd.isna(x):
        return ""
    return str(x).strip()


def clean_gstin(x):
    if pd.isna(x):
        return ""

    value = str(x).upper().strip()
    value = value.replace(" ", "")
    value = value.replace("-", "")

    return value


def clean_invoice(x):
    if pd.isna(x):
        return ""

    x = str(x).upper().strip()

    if x.endswith(".0"):
        x = x[:-2]

    return re.sub(r"[^A-Z0-9]", "", x)


def amount(x):
    if pd.isna(x):
        return 0.0

    try:
        value = str(x).strip()

        value = value.replace(",", "")
        value = value.replace("₹", "")
        value = value.replace("Rs.", "")
        value = value.replace("Rs", "")
        value = value.replace("INR", "")
        value = value.replace("(", "-")
        value = value.replace(")", "")

        return float(value)

    except Exception:
        return 0.0


# =========================================================
# COLUMN ALIASES
# =========================================================
# Software will look for these names.
# Different names between Books and GST 2B are supported.

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
        "Vendor GSTIN",
        "Party GSTIN",
        "GSTIN/UIN of Recipient"
    ],

    "Invoice": [
        "Invoice Number",
        "Invoice number",
        "Invoice No",
        "Invoice No.",
        "Invoice No:",
        "Document Number",
        "Document No",
        "Document No.",
        "Bill No",
        "Bill Number",
        "Bill No.",
        "Invoice"
    ],

    "Date": [
        "Invoice Date",
        "Invoice date",
        "Document Date",
        "Document date",
        "Bill Date",
        "Bill date",
        "Date"
    ],

    "Party": [
        "Party Name",
        "Party name",
        "Supplier Name",
        "Supplier name",
        "Supplier",
        "Vendor Name",
        "Vendor name",
        "Vendor",
        "Trade Name",
        "Trade name",
        "Legal Name",
        "Legal name",
        "Party"
    ],

    "Taxable": [
        "Taxable Value",
        "Taxable value",
        "Taxable Amount",
        "Taxable Amount (₹)",
        "Taxable Amt",
        "Taxable",
        "Taxable Value (₹)",
        "Taxable Value INR",
        "Taxable Amount INR"
    ],

    "IGST": [
        "IGST",
        "IGST Amount",
        "IGST Amt",
        "Integrated Tax",
        "Integrated Tax Amount",
        "Integrated Tax Amount (₹)",
        "IGST Amount (₹)"
    ],

    "CGST": [
        "CGST",
        "CGST Amount",
        "CGST Amt",
        "Central Tax",
        "Central Tax Amount",
        "Central Tax Amount (₹)",
        "CGST Amount (₹)"
    ],

    "SGST": [
        "SGST",
        "SGST Amount",
        "SGST Amt",
        "UTGST",
        "UTGST Amount",
        "State Tax",
        "State Tax Amount",
        "State/UT Tax",
        "State/UT Tax Amount",
        "State Tax Amount (₹)",
        "SGST Amount (₹)"
    ],

    "InvoiceValue": [
        "Invoice Value",
        "Invoice value",
        "Total Invoice Value",
        "Total Value",
        "Document Value",
        "Invoice Amount",
        "Total Invoice Amount",
        "Total Invoice Value (₹)",
        "Invoice Value (₹)",
        "Document Value (₹)"
    ]
}


def normalize_column(x):

    return re.sub(
        r"[^A-Z0-9]",
        "",
        str(x).upper()
    )


def find_column(df, aliases):

    # First: exact normalized match
    columns = {
        normalize_column(c): c
        for c in df.columns
    }

    for alias in aliases:

        key = normalize_column(alias)

        if key in columns:
            return columns[key]

    # Second: partial match
    for alias in aliases:

        aa = normalize_column(alias)

        for c in df.columns:

            cc = normalize_column(c)

            if aa in cc or cc in aa:
                return c

    return None


# =========================================================
# READ EXCEL
# =========================================================

def read_excel(file):

    excel = pd.ExcelFile(file)

    all_rows = []
    sheets = []

    for sheet in excel.sheet_names:

        try:

            df = pd.read_excel(
                file,
                sheet_name=sheet
            )

            if df.empty:
                continue

            gstin = find_column(
                df,
                ALIASES["GSTIN"]
            )

            invoice = find_column(
                df,
                ALIASES["Invoice"]
            )

            # Only invoice-related sheets
            if gstin or invoice:

                df["_SHEET"] = sheet

                all_rows.append(df)

                sheets.append(sheet)

        except Exception:
            continue

    if not all_rows:

        raise ValueError(
            "Excel file mein GST invoice data nahi mila."
        )

    return pd.concat(
        all_rows,
        ignore_index=True
    ), sheets


# =========================================================
# STANDARDIZE DATA
# =========================================================

def standardize(df):

    result = pd.DataFrame()

    detected = {}

    # Detect required columns
    for key in ALIASES:

        detected[key] = find_column(
            df,
            ALIASES[key]
        )

    # -----------------------------------------------------
    # GSTIN
    # -----------------------------------------------------

    if detected["GSTIN"]:

        result["GSTIN"] = df[
            detected["GSTIN"]
        ].apply(clean_gstin)

    else:

        result["GSTIN"] = ""

    # -----------------------------------------------------
    # INVOICE
    # -----------------------------------------------------

    if detected["Invoice"]:

        result["Invoice Number"] = df[
            detected["Invoice"]
        ].apply(clean_invoice)

    else:

        result["Invoice Number"] = ""

    # -----------------------------------------------------
    # DATE
    # -----------------------------------------------------

    if detected["Date"]:

        result["Invoice Date"] = pd.to_datetime(
            df[detected["Date"]],
            errors="coerce",
            dayfirst=True
        )

    else:

        result["Invoice Date"] = pd.NaT

    # -----------------------------------------------------
    # PARTY
    # -----------------------------------------------------

    if detected["Party"]:

        result["Party Name"] = df[
            detected["Party"]
        ].apply(clean_text)

    else:

        result["Party Name"] = ""

    # -----------------------------------------------------
    # TAXABLE
    # -----------------------------------------------------

    if detected["Taxable"]:

        result["Taxable Value"] = df[
            detected["Taxable"]
        ].apply(amount)

    else:

        result["Taxable Value"] = 0.0

    # -----------------------------------------------------
    # IGST
    # -----------------------------------------------------

    if detected["IGST"]:

        result["IGST"] = df[
            detected["IGST"]
        ].apply(amount)

    else:

        result["IGST"] = 0.0

    # -----------------------------------------------------
    # CGST
    # -----------------------------------------------------

    if detected["CGST"]:

        result["CGST"] = df[
            detected["CGST"]
        ].apply(amount)

    else:

        result["CGST"] = 0.0

    # -----------------------------------------------------
    # SGST
    # -----------------------------------------------------

    if detected["SGST"]:

        result["SGST"] = df[
            detected["SGST"]
        ].apply(amount)

    else:

        result["SGST"] = 0.0

    # -----------------------------------------------------
    # INVOICE VALUE
    # -----------------------------------------------------

    if detected["InvoiceValue"]:

        result["Invoice Value"] = df[
            detected["InvoiceValue"]
        ].apply(amount)

    else:

        result["Invoice Value"] = (
            result["Taxable Value"]
            + result["IGST"]
            + result["CGST"]
            + result["SGST"]
        )

    # -----------------------------------------------------
    # SOURCE SHEET
    # -----------------------------------------------------

    if "_SHEET" in df.columns:

        result["Source Sheet"] = df[
            "_SHEET"
        ]

    else:

        result["Source Sheet"] = ""

    # -----------------------------------------------------
    # CLEAN NUMERIC VALUES
    # -----------------------------------------------------

    numeric_columns = [
        "Taxable Value",
        "IGST",
        "CGST",
        "SGST",
        "Invoice Value"
    ]

    for col in numeric_columns:

        result[col] = pd.to_numeric(
            result[col],
            errors="coerce"
        ).fillna(0.0)

    # -----------------------------------------------------
    # REMOVE INVALID RECORDS
    # -----------------------------------------------------

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

    # Matching key
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

    # -----------------------------------------------------
    # IMPORTANT
    # Fill missing numeric values
    # -----------------------------------------------------

    numeric_columns = [
        "Taxable Value_Books",
        "Taxable Value_2B",
        "IGST_Books",
        "IGST_2B",
        "CGST_Books",
        "CGST_2B",
        "SGST_Books",
        "SGST_2B",
        "Invoice Value_Books",
        "Invoice Value_2B"
    ]

    for col in numeric_columns:

        if col in result.columns:

            result[col] = pd.to_numeric(
                result[col],
                errors="coerce"
            ).fillna(0.0)

    statuses = []
    differences = []

    # -----------------------------------------------------
    # CHECK EACH INVOICE
    # -----------------------------------------------------

    for _, row in result.iterrows():

        # ---------------------------------------------
        # Missing in 2B
        # ---------------------------------------------

        if row["_merge"] == "left_only":

            statuses.append(
                "Missing in 2B"
            )

            diff = (
                row["IGST_Books"]
                + row["CGST_Books"]
                + row["SGST_Books"]
            )

            differences.append(
                round(diff, 2)
            )

        # ---------------------------------------------
        # Missing in Books
        # ---------------------------------------------

        elif row["_merge"] == "right_only":

            statuses.append(
                "Missing in Books"
            )

            diff = (
                row["IGST_2B"]
                + row["CGST_2B"]
                + row["SGST_2B"]
            )

            differences.append(
                round(diff, 2)
            )

        # ---------------------------------------------
        # Present in both
        # ---------------------------------------------

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

            total_diff = (
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

                statuses.append(
                    "Matched"
                )

                differences.append(0.0)

            else:

                statuses.append(
                    "Value Mismatch"
                )

                differences.append(
                    round(total_diff, 2)
                )

    result["Status"] = statuses

    result["ITC Difference"] = differences

    result["ITC Difference"] = pd.to_numeric(
        result["ITC Difference"],
        errors="coerce"
    ).fillna(0.0)

    return result


# =========================================================
# DISPLAY
# =========================================================

def prepare_display(df):

    out = pd.DataFrame()

    def col(name, default=""):

        if name in df.columns:

            return df[name]

        return pd.Series(
            [default] * len(df),
            index=df.index
        )

    out["GSTIN"] = col(
        "GSTIN_Books"
    ).where(
        col("GSTIN_Books") != "",
        col("GSTIN_2B")
    )

    out["Party Name"] = col(
        "Party Name_Books"
    ).where(
        col("Party Name_Books") != "",
        col("Party Name_2B")
    )

    out["Invoice Number"] = col(
        "Invoice Number_Books"
    ).where(
        col("Invoice Number_Books") != "",
        col("Invoice Number_2B")
    )

    out["Invoice Date"] = col(
        "Invoice Date_Books"
    ).where(
        col("Invoice Date_Books").notna(),
        col("Invoice Date_2B")
    )

    out["Taxable - Books"] = col(
        "Taxable Value_Books"
    )

    out["Taxable - 2B"] = col(
        "Taxable Value_2B"
    )

    out["IGST - Books"] = col(
        "IGST_Books"
    )

    out["IGST - 2B"] = col(
        "IGST_2B"
    )

    out["CGST - Books"] = col(
        "CGST_Books"
    )

    out["CGST - 2B"] = col(
        "CGST_2B"
    )

    out["SGST - Books"] = col(
        "SGST_Books"
    )

    out["SGST - 2B"] = col(
        "SGST_2B"
    )

    out["Invoice Value - Books"] = col(
        "Invoice Value_Books"
    )

    out["Invoice Value - 2B"] = col(
        "Invoice Value_2B"
    )

    out["ITC Difference"] = col(
        "ITC Difference",
        0.0
    )

    out["Status"] = col(
        "Status"
    )

    return out.reset_index(drop=True)


# =========================================================
# TABLE
# =========================================================

def show_table(df):

    display = prepare_display(df)

    if display.empty:

        st.info(
            "Is category mein koi invoice nahi hai."
        )

        return

    display["Invoice Date"] = pd.to_datetime(
        display["Invoice Date"],
        errors="coerce"
    ).dt.strftime("%d-%m-%Y")

    display["Invoice Date"] = display[
        "Invoice Date"
    ].fillna("-")

    amount_columns = [
        "Taxable - Books",
        "Taxable - 2B",
        "IGST - Books",
        "IGST - 2B",
        "CGST - Books",
        "CGST - 2B",
        "SGST - Books",
        "SGST - 2B",
        "Invoice Value - Books",
        "Invoice Value - 2B",
        "ITC Difference"
    ]

    for c in amount_columns:

        if c in display.columns:

            display[c] = pd.to_numeric(
                display[c],
                errors="coerce"
            ).fillna(0)

            display[c] = display[c].apply(
                lambda x:
                f"₹{float(x):,.2f}"
            )

    html = display.to_html(
        index=False,
        classes="custom-table",
        escape=True
    )

    st.markdown(
        html,
        unsafe_allow_html=True
    )


# =========================================================
# EXCEL REPORT
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
# HEADER
# =========================================================

st.markdown("""
<div class="hero">

    <h1>🧾 GST Reconciliation Pro</h1>

    <p>
        GST 2B vs Books &nbsp;•&nbsp;
        ITC Reconciliation &nbsp;•&nbsp;
        Invoice Exception Analysis
    </p>

</div>
""", unsafe_allow_html=True)


# =========================================================
# SIDEBAR
# =========================================================

with st.sidebar:

    st.markdown(
        "## ⚙️ Reconciliation Settings"
    )

    tolerance = st.number_input(
        "Mismatch Tolerance ₹",
        min_value=0.0,
        value=2.0,
        step=0.50
    )

    st.divider()

    st.markdown(
        "### 🔑 Matching Key"
    )

    st.info(
        "GSTIN + Invoice Number"
    )

    st.divider()

    st.markdown(
        "### 📌 Status"
    )

    st.caption(
        "🟢 Matched\n\n"
        "🔴 Missing in 2B\n\n"
        "🟠 Missing in Books\n\n"
        "🟣 Value Mismatch"
    )


# =========================================================
# FILE UPLOAD
# =========================================================

st.markdown(
    '<div class="section-title">📂 Upload GST Data</div>',
    unsafe_allow_html=True
)

u1, u2 = st.columns(2)

with u1:

    st.markdown("""
    <div class="upload-card">

        <div class="upload-title">
            📥 GST 2B File
        </div>

        <div class="upload-subtitle">
            Upload your GST portal 2B Excel file.
            Multiple sheets are supported.
        </div>

    </div>
    """, unsafe_allow_html=True)

    two_b_file = st.file_uploader(
        "Choose GST 2B Excel",
        type=["xlsx", "xls"],
        key="two_b_file"
    )


with u2:

    st.markdown("""
    <div class="upload-card">

        <div class="upload-title">
            📚 Books File
        </div>

        <div class="upload-subtitle">
            Upload your purchase/books Excel file.
        </div>

    </div>
    """, unsafe_allow_html=True)

    books_file = st.file_uploader(
        "Choose Books Excel",
        type=["xlsx", "xls"],
        key="books_file"
    )


# =========================================================
# PROCESS
# =========================================================

if two_b_file and books_file:

    st.markdown("")

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

                # Safety check
                if two_b.empty:

                    raise ValueError(
                        "GST 2B mein valid GSTIN + Invoice Number records nahi mile."
                    )

                if books.empty:

                    raise ValueError(
                        "Books file mein valid GSTIN + Invoice Number records nahi mile."
                    )

                result = reconcile(
                    two_b,
                    books,
                    tolerance
                )

                st.session_state[
                    "reco_result"
                ] = result

                st.session_state[
                    "reco_two_b"
                ] = two_b

                st.session_state[
                    "reco_books"
                ] = books

                st.session_state[
                    "reco_two_b_sheets"
                ] = two_b_sheets

                st.session_state[
                    "reco_books_sheets"
                ] = books_sheets

            st.success(
                "✅ GST Reconciliation completed successfully."
            )

        except Exception as e:

            st.error(
                "❌ Error while processing files."
            )

            st.exception(e)


# =========================================================
# DASHBOARD
# =========================================================

if "reco_result" in st.session_state:

    result = st.session_state[
        "reco_result"
    ]

    two_b = st.session_state[
        "reco_two_b"
    ]

    books = st.session_state[
        "reco_books"
    ]

    # =====================================================
    # COUNTS
    # =====================================================

    total = len(result)

    matched = (
        result["Status"] == "Matched"
    ).sum()

    missing_2b = (
        result["Status"] == "Missing in 2B"
    ).sum()

    missing_books = (
        result["Status"] == "Missing in Books"
    ).sum()

    mismatch = (
        result["Status"] == "Value Mismatch"
    ).sum()

    # =====================================================
    # IMPORTANT:
    # VALUE MISMATCH AMOUNT
    # =====================================================

    mismatch_amount = pd.to_numeric(
        result.loc[
            result["Status"] == "Value Mismatch",
            "ITC Difference"
        ],
        errors="coerce"
    ).fillna(0).sum()

    missing_2b_amount = pd.to_numeric(
        result.loc[
            result["Status"] == "Missing in 2B",
            "ITC Difference"
        ],
        errors="coerce"
    ).fillna(0).sum()

    missing_books_amount = pd.to_numeric(
        result.loc[
            result["Status"] == "Missing in Books",
            "ITC Difference"
        ],
        errors="coerce"
    ).fillna(0).sum()

    # =====================================================
    # KPI
    # =====================================================

    st.markdown(
        '<div class="section-title">📊 Reconciliation Dashboard</div>',
        unsafe_allow_html=True
    )

    c1, c2, c3, c4, c5 = st.columns(5)

    cards = [

        (
            c1,
            "TOTAL INVOICES",
            f"{total:,}",
            "All analysed records",
            "kpi-blue"
        ),

        (
            c2,
            "MATCHED",
            f"{matched:,}",
            "Successfully matched",
            "kpi-green"
        ),

        (
            c3,
            "MISSING IN 2B",
            f"{missing_2b:,}",
            f"ITC ₹{missing_2b_amount:,.2f}",
            "kpi-red"
        ),

        (
            c4,
            "MISSING IN BOOKS",
            f"{missing_books:,}",
            f"ITC ₹{missing_books_amount:,.2f}",
            "kpi-orange"
        ),

        (
            c5,
            "VALUE MISMATCH",
            f"{mismatch:,}",
            f"Mismatch Amount ₹{mismatch_amount:,.2f}",
            "kpi-purple"
        )
    ]

    for col, title, number, desc, colour in cards:

        with col:

            st.markdown(
                f"""
                <div class="kpi {colour}">

                    <div class="kpi-label">
                        {title}
                    </div>

                    <div class="kpi-value">
                        {number}
                    </div>

                    <div class="kpi-description">
                        {desc}
                    </div>

                </div>
                """,
                unsafe_allow_html=True
            )


    # =====================================================
    # EXCEPTIONS
    # =====================================================

    st.markdown(
        '<div class="section-title">🔎 Invoice Analysis</div>',
        unsafe_allow_html=True
    )

    b1, b2, b3, b4 = st.columns(4)

    with b1:

        if st.button(
            f"🔴 Missing in 2B • {missing_2b}",
            key="category_2b",
            use_container_width=True
        ):

            st.session_state[
                "selected_status"
            ] = "Missing in 2B"


    with b2:

        if st.button(
            f"🟠 Missing in Books • {missing_books}",
            key="category_books",
            use_container_width=True
        ):

            st.session_state[
                "selected_status"
            ] = "Missing in Books"


    with b3:

        if st.button(
            f"🟣 Value Mismatch • {mismatch}",
            key="category_mismatch",
            use_container_width=True
        ):

            st.session_state[
                "selected_status"
            ] = "Value Mismatch"


    with b4:

        if st.button(
            f"🟢 Matched • {matched}",
            key="category_matched",
            use_container_width=True
        ):

            st.session_state[
                "selected_status"
            ] = "Matched"


    if "selected_status" not in st.session_state:

        st.session_state[
            "selected_status"
        ] = "Missing in 2B"


    selected = st.session_state[
        "selected_status"
    ]

    detail = result[
        result["Status"] == selected
    ].copy()


    st.markdown(
        f"""
        <div class="section-title">
            📋 {selected} — {len(detail):,} Invoices
        </div>
        """,
        unsafe_allow_html=True
    )


    info_text = {

        "Missing in 2B":
        "These invoices are present in Books but are not available in GST 2B.",

        "Missing in Books":
        "These invoices are present in GST 2B but are not recorded in Books.",

        "Value Mismatch":
        "GSTIN and invoice number matched, but taxable/tax/invoice values are different.",

        "Matched":
        "These invoices successfully matched between Books and GST 2B."
    }


    st.markdown(
        f"""
        <div class="info-box">

            <b>ℹ️ Analysis:</b><br>

            {info_text.get(selected, "")}

        </div>
        """,
        unsafe_allow_html=True
    )

    # Show mismatch amount when category selected
    if selected == "Value Mismatch":

        st.metric(
            "💰 Total ITC Mismatch Amount",
            f"₹{mismatch_amount:,.2f}"
        )

    elif selected == "Missing in 2B":

        st.metric(
            "💰 ITC Missing from 2B",
            f"₹{missing_2b_amount:,.2f}"
        )

    elif selected == "Missing in Books":

        st.metric(
            "💰 ITC Missing from Books",
            f"₹{missing_books_amount:,.2f}"
        )

    show_table(detail)


    # =====================================================
    # ITC ANALYSIS
    # =====================================================

    st.markdown(
        '<div class="section-title">💰 ITC Analysis</div>',
        unsafe_allow_html=True
    )

    b_igst = books["IGST"].sum()
    b_cgst = books["CGST"].sum()
    b_sgst = books["SGST"].sum()

    t_igst = two_b["IGST"].sum()
    t_cgst = two_b["CGST"].sum()
    t_sgst = two_b["SGST"].sum()

    books_itc = (
        b_igst
        + b_cgst
        + b_sgst
    )

    two_b_itc = (
        t_igst
        + t_cgst
        + t_sgst
    )

    itc_difference = (
        books_itc
        - two_b_itc
    )


    i1, i2, i3 = st.columns(3)

    with i1:

        st.metric(
            "📚 Books Total ITC",
            f"₹{books_itc:,.2f}"
        )

    with i2:

        st.metric(
            "📥 GST 2B Total ITC",
            f"₹{two_b_itc:,.2f}"
        )

    with i3:

        st.metric(
            "⚖️ ITC Difference",
            f"₹{itc_difference:,.2f}"
        )


    # =====================================================
    # TAX TABLE
    # =====================================================

    tax_df = pd.DataFrame({

        "Tax Component":
        [
            "IGST",
            "CGST",
            "SGST",
            "TOTAL ITC"
        ],

        "Books":
        [
            b_igst,
            b_cgst,
            b_sgst,
            books_itc
        ],

        "GST 2B":
        [
            t_igst,
            t_cgst,
            t_sgst,
            two_b_itc
        ],

        "Difference":
        [
            b_igst - t_igst,
            b_cgst - t_cgst,
            b_sgst - t_sgst,
            itc_difference
        ]
    })


    tax_display = tax_df.copy()

    for c in [
        "Books",
        "GST 2B",
        "Difference"
    ]:

        tax_display[c] = tax_display[
            c
        ].apply(
            lambda x:
            f"₹{x:,.2f}"
        )


    st.markdown(
        tax_display.to_html(
            index=False,
            classes="custom-table"
        ),
        unsafe_allow_html=True
    )


    # =====================================================
    # CHARTS
    # =====================================================

    st.markdown(
        '<div class="section-title">📈 Visual Analysis</div>',
        unsafe_allow_html=True
    )

    chart1, chart2 = st.columns(2)

    with chart1:

        status_data = pd.DataFrame({

            "Status":
            [
                "Matched",
                "Missing in 2B",
                "Missing in Books",
                "Value Mismatch"
            ],

            "Count":
            [
                matched,
                missing_2b,
                missing_books,
                mismatch
            ]
        })


        fig = px.pie(
            status_data,
            names="Status",
            values="Count",
            hole=0.58,
            title="Reconciliation Status"
        )


        fig.update_traces(
            textinfo="percent+label"
        )


        fig.update_layout(
            paper_bgcolor="white",
            plot_bgcolor="white",
            font_color="#172033",
            margin=dict(
                l=20,
                r=20,
                t=60,
                b=20
            )
        )


        st.plotly_chart(
            fig,
            use_container_width=True
        )


    with chart2:

        itc_data = pd.DataFrame({

            "Tax":
            [
                "IGST",
                "CGST",
                "SGST"
            ],

            "Amount":
            [
                t_igst,
                t_cgst,
                t_sgst
            ]
        })


        fig2 = px.pie(
            itc_data,
            names="Tax",
            values="Amount",
            hole=0.58,
            title="GST 2B ITC Distribution"
        )


        fig2.update_traces(
            textinfo="percent+label"
        )


        fig2.update_layout(
            paper_bgcolor="white",
            plot_bgcolor="white",
            font_color="#172033",
            margin=dict(
                l=20,
                r=20,
                t=60,
                b=20
            )
        )


        st.plotly_chart(
            fig2,
            use_container_width=True
        )


    # =====================================================
    # SHEETS
    # =====================================================

    with st.expander(
        "📂 Excel Sheets Automatically Analysed"
    ):

        s1, s2 = st.columns(2)

        with s1:

            st.markdown(
                "### 📥 GST 2B"
            )

            for sheet in st.session_state[
                "reco_two_b_sheets"
            ]:

                st.write(
                    "•",
                    sheet
                )

        with s2:

            st.markdown(
                "### 📚 Books"
            )

            for sheet in st.session_state[
                "reco_books_sheets"
            ]:

                st.write(
                    "•",
                    sheet
                )


    # =====================================================
    # EXPORT
    # =====================================================

    st.markdown(
        '<div class="section-title">📥 Export Report</div>',
        unsafe_allow_html=True
    )


    excel_file = create_excel(
        result
    )


    st.download_button(
        "📊 DOWNLOAD COMPLETE EXCEL REPORT",
        data=excel_file,
        file_name="GST_Reconciliation_Report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )


# =========================================================
# WELCOME SCREEN
# =========================================================

else:

    st.markdown("""
    <div class="info-box">

        <h3>👋 Welcome to GST Reconciliation Pro</h3>

        <p>
        Upload your <b>GST 2B</b> and
        <b>Books</b> Excel files to begin reconciliation.
        </p>

        <p>
        The software automatically detects the required
        GST columns even when Books and GST 2B use
        different column names.
        </p>

        <h3>🔑 Matching Key</h3>

        <p>
        <b>GSTIN + Invoice Number</b>
        </p>

        <p>
        The software identifies:
        </p>

        <p>
        🔴 Missing in 2B<br>
        🟠 Missing in Books<br>
        🟣 Value Mismatch<br>
        🟢 Matched
        </p>

    </div>
    """, unsafe_allow_html=True)
