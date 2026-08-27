import streamlit as st
import pandas as pd
import plotly.express as px
from io import BytesIO
import re


# =========================================================
# PAGE
# =========================================================

st.set_page_config(
    page_title="GST Reconciliation Pro",
    page_icon="📊",
    layout="wide"
)


# =========================================================
# PROFESSIONAL WHITE UI  (FIXED — full coverage of Streamlit's
# internal containers so dark theme / OS dark-mode can never
# leak through. Root fix also lives in .streamlit/config.toml)
# =========================================================

st.markdown("""
<style>

html {
    color-scheme: light !important;
}

/* Root app containers */
.stApp,
[data-testid="stAppViewContainer"],
[data-testid="stHeader"],
[data-testid="stToolbar"],
[data-testid="stBottomBlockContainer"],
[data-testid="stMain"],
[data-testid="stMainBlockContainer"],
body {
    background: #ffffff !important;
    color: #111827 !important;
}

[data-testid="stHeader"] {
    background: transparent !important;
}

.main .block-container {
    max-width: 1450px;
    padding-top: 30px;
    background: #ffffff !important;
}

h1, h2, h3, h4, h5, h6, p, span, label, li,
.stMarkdown, .stMarkdown p, .stText,
[data-testid="stMarkdownContainer"] {
    color: #111827 !important;
}

/* Sidebar */
[data-testid="stSidebar"],
[data-testid="stSidebarContent"],
[data-testid="stSidebarUserContent"] {
    background: #f8fafc !important;
    color: #111827 !important;
    border-right: 1px solid #e5e7eb;
}

[data-testid="stSidebar"] * {
    color: #111827 !important;
}


/* Header */

.header {
    background: #ffffff;
    border: 1px solid #e5e7eb;
    border-radius: 16px;
    padding: 25px 30px;
    margin-bottom: 25px;
    box-shadow: 0 4px 15px rgba(0,0,0,.05);
}

.header h1 {
    margin: 0;
    font-size: 32px;
}

.header p {
    color: #667085 !important;
    margin-top: 7px;
}


/* KPI CARDS */

.kpi {
    background: #ffffff;
    border: 1px solid #e5e7eb;
    border-radius: 14px;
    padding: 20px;
    min-height: 125px;
    box-shadow: 0 3px 12px rgba(0,0,0,.05);
    transition: .2s;
}

.kpi:hover {
    transform: translateY(-5px);
    box-shadow: 0 12px 25px rgba(0,0,0,.10);
}

.kpi-title {
    font-size: 13px;
    font-weight: 700;
    color: #667085;
}

.kpi-number {
    font-size: 30px;
    font-weight: 800;
    margin-top: 8px;
    color: #111827;
}

.blue   { border-top: 4px solid #2563eb; }
.green  { border-top: 4px solid #16a34a; }
.red    { border-top: 4px solid #dc2626; }
.orange { border-top: 4px solid #ea580c; }
.purple { border-top: 4px solid #7c3aed; }


/* BUTTONS */

.stButton > button,
.stDownloadButton > button {
    background: #ffffff !important;
    color: #111827 !important;
    border: 1px solid #d1d5db !important;
    border-radius: 10px !important;
    min-height: 48px;
    font-weight: 700;
    transition: .2s;
}

.stButton > button:hover,
.stDownloadButton > button:hover {
    background: #f8fafc !important;
    color: #2563eb !important;
    border-color: #2563eb !important;
    transform: translateY(-3px);
    box-shadow: 0 7px 18px rgba(37,99,235,.12);
}

.stButton > button[kind="primary"] {
    background: #2563eb !important;
    color: #ffffff !important;
    border-color: #2563eb !important;
}

.stButton > button[kind="primary"]:hover {
    background: #1d4ed8 !important;
    color: #ffffff !important;
}


/* UPLOAD */

[data-testid="stFileUploader"],
[data-testid="stFileUploaderDropzone"] {
    background: #ffffff !important;
    border: 1px solid #d1d5db !important;
    border-radius: 12px;
    color: #111827 !important;
}

[data-testid="stFileUploaderDropzone"] * {
    color: #111827 !important;
}


/* INPUTS (number input, selectbox, text input) */

.stNumberInput input,
.stTextInput input,
.stSelectbox div[data-baseweb="select"] > div,
.stDateInput input {
    background: #ffffff !important;
    color: #111827 !important;
    border: 1px solid #d1d5db !important;
}

.stNumberInput label,
.stTextInput label,
.stSelectbox label {
    color: #111827 !important;
}


/* METRICS */

[data-testid="stMetric"] {
    background: #ffffff !important;
    border: 1px solid #e5e7eb;
    border-radius: 12px;
    padding: 12px;
}

[data-testid="stMetricLabel"],
[data-testid="stMetricValue"],
[data-testid="stMetricDelta"] {
    color: #111827 !important;
}


/* EXPANDER */

[data-testid="stExpander"] {
    background: #ffffff !important;
    border: 1px solid #e5e7eb !important;
    border-radius: 10px;
}

[data-testid="stExpander"] summary,
[data-testid="stExpander"] * {
    color: #111827 !important;
    background: #ffffff !important;
}


/* NATIVE / BUILT-IN DATAFRAME (if used anywhere) */

[data-testid="stDataFrame"],
[data-testid="stTable"] {
    background: #ffffff !important;
}


/* CUSTOM TABLE (used for reconciliation results) */

.custom-table {
    width: 100%;
    border-collapse: collapse;
    background: #ffffff;
    border: 1px solid #e5e7eb;
    border-radius: 10px;
    overflow: hidden;
    font-size: 13px;
}

.custom-table th {
    background: #f8fafc;
    color: #344054;
    font-weight: 700;
    padding: 12px;
    border-bottom: 1px solid #e5e7eb;
    text-align: left;
    white-space: nowrap;
}

.custom-table td {
    background: #ffffff;
    color: #111827;
    padding: 11px 12px;
    border-bottom: 1px solid #f0f1f3;
    white-space: nowrap;
}

.custom-table tr:hover td {
    background: #f8fafc;
}


/* INFO */

.info-box {
    background: #f8fafc;
    border: 1px solid #e5e7eb;
    border-radius: 12px;
    padding: 15px;
    color: #344054;
    margin: 12px 0;
}

.stAlert, [data-testid="stNotification"] {
    background: #f8fafc !important;
    color: #111827 !important;
}


/* SECTION */

.section {
    font-size: 22px;
    font-weight: 800;
    margin-top: 28px;
    margin-bottom: 12px;
    color: #111827;
}


/* STATUS */

.status-red    { color: #dc2626; font-weight: 800; }
.status-orange { color: #ea580c; font-weight: 800; }
.status-purple { color: #7c3aed; font-weight: 800; }
.status-green  { color: #16a34a; font-weight: 800; }

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
    return str(x).upper().replace(" ", "").strip()


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
        return float(
            str(x)
            .replace(",", "")
            .replace("₹", "")
            .strip()
        )
    except:
        return 0.0


# =========================================================
# COLUMN DETECTION
# =========================================================

ALIASES = {

    "GSTIN": [
        "GSTIN",
        "GSTIN/UIN",
        "GSTIN of Supplier",
        "Supplier GSTIN",
        "GST Number",
        "GST No",
        "GST No.",
        "GSTIN of Vendor"
    ],

    "Invoice": [
        "Invoice Number",
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
        "Document Date",
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
        "Party"
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
        "Integrated Tax Amount"
    ],

    "CGST": [
        "CGST",
        "CGST Amount",
        "CGST Amt",
        "Central Tax",
        "Central Tax Amount"
    ],

    "SGST": [
        "SGST",
        "SGST Amount",
        "SGST Amt",
        "UTGST",
        "State Tax",
        "State Tax Amount"
    ],

    "InvoiceValue": [
        "Invoice Value",
        "Invoice value",
        "Total Invoice Value",
        "Total Value",
        "Document Value",
        "Invoice Amount",
        "Total Invoice Amount"
    ]
}


def normalize_column(x):
    return re.sub(
        r"[^A-Z0-9]",
        "",
        str(x).upper()
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

    for c in df.columns:

        cc = normalize_column(c)

        for alias in aliases:

            aa = normalize_column(alias)

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

            if gstin or invoice:

                df["_SHEET"] = sheet

                all_rows.append(df)

                sheets.append(sheet)

        except Exception:
            pass

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

    c = {}

    for key in ALIASES:
        c[key] = find_column(
            df,
            ALIASES[key]
        )

    # GSTIN

    if c["GSTIN"]:
        result["GSTIN"] = df[c["GSTIN"]].apply(
            clean_gstin
        )
    else:
        result["GSTIN"] = ""

    # Invoice

    if c["Invoice"]:
        result["Invoice Number"] = df[
            c["Invoice"]
        ].apply(clean_invoice)
    else:
        result["Invoice Number"] = ""

    # Date

    if c["Date"]:
        result["Invoice Date"] = pd.to_datetime(
            df[c["Date"]],
            errors="coerce"
        )
    else:
        result["Invoice Date"] = pd.NaT

    # Party

    if c["Party"]:
        result["Party Name"] = df[
            c["Party"]
        ].apply(clean_text)
    else:
        result["Party Name"] = ""

    # Taxable

    if c["Taxable"]:
        result["Taxable Value"] = df[
            c["Taxable"]
        ].apply(amount)
    else:
        result["Taxable Value"] = 0.0

    # IGST

    if c["IGST"]:
        result["IGST"] = df[
            c["IGST"]
        ].apply(amount)
    else:
        result["IGST"] = 0.0

    # CGST

    if c["CGST"]:
        result["CGST"] = df[
            c["CGST"]
        ].apply(amount)
    else:
        result["CGST"] = 0.0

    # SGST

    if c["SGST"]:
        result["SGST"] = df[
            c["SGST"]
        ].apply(amount)
    else:
        result["SGST"] = 0.0

    # Invoice Value

    if c["InvoiceValue"]:

        result["Invoice Value"] = df[
            c["InvoiceValue"]
        ].apply(amount)

    else:

        result["Invoice Value"] = (
            result["Taxable Value"]
            + result["IGST"]
            + result["CGST"]
            + result["SGST"]
        )

    # Source sheet

    if "_SHEET" in df.columns:
        result["Source Sheet"] = df[
            "_SHEET"
        ]
    else:
        result["Source Sheet"] = ""

    # Remove blank records

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

        # Missing in 2B

        if row["_merge"] == "left_only":

            statuses.append("Missing in 2B")

            differences.append(
                row["IGST_Books"]
                + row["CGST_Books"]
                + row["SGST_Books"]
            )

        # Missing in Books

        elif row["_merge"] == "right_only":

            statuses.append("Missing in Books")

            differences.append(
                row["IGST_2B"]
                + row["CGST_2B"]
                + row["SGST_2B"]
            )

        # Both available

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

                statuses.append("Matched")

            else:

                statuses.append(
                    "Value Mismatch"
                )

            differences.append(
                total_diff
            )

    result["Status"] = statuses
    result["ITC Difference"] = differences

    return result


# =========================================================
# DISPLAY TABLE
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

    out["Taxable - Books"] = col("Taxable Value_Books")
    out["Taxable - 2B"] = col("Taxable Value_2B")

    out["IGST - Books"] = col("IGST_Books")
    out["IGST - 2B"] = col("IGST_2B")

    out["CGST - Books"] = col("CGST_Books")
    out["CGST - 2B"] = col("CGST_2B")

    out["SGST - Books"] = col("SGST_Books")
    out["SGST - 2B"] = col("SGST_2B")

    out["Invoice Value - Books"] = col("Invoice Value_Books")
    out["Invoice Value - 2B"] = col("Invoice Value_2B")

    out["ITC Difference"] = col("ITC Difference")
    out["Status"] = col("Status")

    return out.reset_index(drop=True)


# =========================================================
# HTML TABLE
# =========================================================

def show_table(df):

    display = prepare_display(df)

    if display.empty:

        st.info("Is category mein koi invoice nahi hai.")

        return

    # Dates

    if "Invoice Date" in display.columns:

        display["Invoice Date"] = pd.to_datetime(
            display["Invoice Date"],
            errors="coerce"
        ).dt.strftime("%d-%m-%Y")

        display["Invoice Date"] = display[
            "Invoice Date"
        ].fillna("-")

    # Amount formatting

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

            display[c] = display[c].apply(
                lambda x: (
                    f"₹{float(x):,.2f}"
                    if pd.notna(x)
                    else "-"
                )
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
<div class="header">
<h1>GST Reconciliation Pro</h1>
<p>GST 2B vs Books • ITC Reconciliation • Invoice Exception Analysis</p>
</div>
""", unsafe_allow_html=True)


# =========================================================
# SIDEBAR
# =========================================================

with st.sidebar:

    st.markdown("## ⚙️ Reconciliation Settings")

    tolerance = st.number_input(
        "Mismatch Tolerance ₹",
        min_value=0.0,
        value=2.0,
        step=0.50
    )

    st.divider()

    st.markdown("### Matching Key")

    st.info("GSTIN + Invoice Number")


# =========================================================
# UPLOAD
# =========================================================

u1, u2 = st.columns(2)

with u1:
    st.markdown("### 📥 GST 2B File")
    two_b_file = st.file_uploader(
        "Upload 2B Excel",
        type=["xlsx", "xls"],
        key="two_b_file"
    )

with u2:
    st.markdown("### 📚 Books File")
    books_file = st.file_uploader(
        "Upload Books Excel",
        type=["xlsx", "xls"],
        key="books_file"
    )


# =========================================================
# PROCESS
# =========================================================

if two_b_file and books_file:

    st.markdown("")

    if st.button(
        "🚀 START RECONCILIATION",
        type="primary",
        use_container_width=True
    ):

        try:

            with st.spinner("Analysing Excel files..."):

                two_b_raw, two_b_sheets = read_excel(two_b_file)
                books_raw, books_sheets = read_excel(books_file)

                two_b = standardize(two_b_raw)
                books = standardize(books_raw)

                result = reconcile(two_b, books, tolerance)

                st.session_state["reco_result"] = result
                st.session_state["reco_two_b"] = two_b
                st.session_state["reco_books"] = books
                st.session_state["reco_two_b_sheets"] = two_b_sheets
                st.session_state["reco_books_sheets"] = books_sheets

            st.success("✅ Reconciliation completed successfully.")

        except Exception as e:

            st.error("❌ Error while processing files")
            st.exception(e)


# =========================================================
# DASHBOARD
# =========================================================

if "reco_result" in st.session_state:

    result = st.session_state["reco_result"]
    two_b = st.session_state["reco_two_b"]
    books = st.session_state["reco_books"]

    # COUNTS

    total = len(result)
    matched = (result["Status"] == "Matched").sum()
    missing_2b = (result["Status"] == "Missing in 2B").sum()
    missing_books = (result["Status"] == "Missing in Books").sum()
    mismatch = (result["Status"] == "Value Mismatch").sum()

    # =====================================================
    # KPI
    # =====================================================

    st.markdown(
        '<div class="section">📊 Reconciliation Overview</div>',
        unsafe_allow_html=True
    )

    c1, c2, c3, c4, c5 = st.columns(5)

    cards = [
        (c1, "TOTAL INVOICES", total, "All analysed records", "blue"),
        (c2, "MATCHED", matched, "Successfully matched", "green"),
        (c3, "MISSING IN 2B", missing_2b, "Books → not in 2B", "red"),
        (c4, "MISSING IN BOOKS", missing_books, "2B → not in Books", "orange"),
        (c5, "VALUE MISMATCH", mismatch, "Tax/value difference", "purple")
    ]

    for col, title, number, desc, colour in cards:

        with col:

            st.markdown(
                f"""
                <div class="kpi {colour}">
                    <div class="kpi-title">{title}</div>
                    <div class="kpi-number">{number:,}</div>
                    <div style="color:#667085;font-size:12px;">{desc}</div>
                </div>
                """,
                unsafe_allow_html=True
            )

    # =====================================================
    # CATEGORY BUTTONS
    # =====================================================

    st.markdown(
        '<div class="section">🔎 Invoice Exceptions</div>',
        unsafe_allow_html=True
    )

    b1, b2, b3, b4 = st.columns(4)

    with b1:
        if st.button(f"🔴 Missing in 2B  |  {missing_2b}", key="category_2b", use_container_width=True):
            st.session_state["selected_status"] = "Missing in 2B"

    with b2:
        if st.button(f"🟠 Missing in Books  |  {missing_books}", key="category_books", use_container_width=True):
            st.session_state["selected_status"] = "Missing in Books"

    with b3:
        if st.button(f"🟣 Value Mismatch  |  {mismatch}", key="category_mismatch", use_container_width=True):
            st.session_state["selected_status"] = "Value Mismatch"

    with b4:
        if st.button(f"🟢 Matched  |  {matched}", key="category_matched", use_container_width=True):
            st.session_state["selected_status"] = "Matched"

    if "selected_status" not in st.session_state:
        st.session_state["selected_status"] = "Missing in 2B"

    selected = st.session_state["selected_status"]

    # =====================================================
    # SELECTED LIST
    # =====================================================

    detail = result[result["Status"] == selected].copy()

    st.markdown(
        f'<div class="section">📋 {selected} — {len(detail):,} Invoices</div>',
        unsafe_allow_html=True
    )

    info_text = {
        "Missing in 2B": "These invoices are present in Books but are not available in GST 2B.",
        "Missing in Books": "These invoices are present in GST 2B but are not recorded in Books.",
        "Value Mismatch": "GSTIN and invoice number matched, but taxable/tax/invoice values are different.",
        "Matched": "These invoices successfully matched between Books and GST 2B."
    }

    st.markdown(
        f'<div class="info-box">{info_text.get(selected, "")}</div>',
        unsafe_allow_html=True
    )

    show_table(detail)

    # =====================================================
    # ITC ANALYSIS
    # =====================================================

    st.markdown(
        '<div class="section">💰 ITC Analysis</div>',
        unsafe_allow_html=True
    )

    b_igst = books["IGST"].sum()
    b_cgst = books["CGST"].sum()
    b_sgst = books["SGST"].sum()

    t_igst = two_b["IGST"].sum()
    t_cgst = two_b["CGST"].sum()
    t_sgst = two_b["SGST"].sum()

    books_itc = b_igst + b_cgst + b_sgst
    two_b_itc = t_igst + t_cgst + t_sgst
    itc_difference = books_itc - two_b_itc

    i1, i2, i3 = st.columns(3)

    with i1:
        st.metric("Books Total ITC", f"₹{books_itc:,.2f}")

    with i2:
        st.metric("GST 2B Total ITC", f"₹{two_b_itc:,.2f}")

    with i3:
        st.metric("ITC Difference", f"₹{itc_difference:,.2f}")

    # =====================================================
    # TAX TABLE
    # =====================================================

    tax_df = pd.DataFrame({
        "Tax Component": ["IGST", "CGST", "SGST", "TOTAL ITC"],
        "Books": [b_igst, b_cgst, b_sgst, books_itc],
        "GST 2B": [t_igst, t_cgst, t_sgst, two_b_itc],
        "Difference": [b_igst - t_igst, b_cgst - t_cgst, b_sgst - t_sgst, itc_difference]
    })

    tax_display = tax_df.copy()

    for c in ["Books", "GST 2B", "Difference"]:
        tax_display[c] = tax_display[c].apply(lambda x: f"₹{x:,.2f}")

    st.markdown(
        tax_display.to_html(index=False, classes="custom-table"),
        unsafe_allow_html=True
    )

    # =====================================================
    # CHARTS
    # =====================================================

    st.markdown(
        '<div class="section">📈 Visual Analysis</div>',
        unsafe_allow_html=True
    )

    chart1, chart2 = st.columns(2)

    with chart1:

        status_data = pd.DataFrame({
            "Status": ["Matched", "Missing in 2B", "Missing in Books", "Value Mismatch"],
            "Count": [matched, missing_2b, missing_books, mismatch]
        })

        fig = px.pie(
            status_data, names="Status", values="Count",
            hole=0.55, title="Reconciliation Status"
        )

        fig.update_traces(textinfo="percent+label")

        fig.update_layout(
            paper_bgcolor="white",
            plot_bgcolor="white",
            font_color="#111827"
        )

        st.plotly_chart(fig, use_container_width=True)

    with chart2:

        itc_data = pd.DataFrame({
            "Tax": ["IGST", "CGST", "SGST"],
            "Amount": [t_igst, t_cgst, t_sgst]
        })

        fig2 = px.pie(
            itc_data, names="Tax", values="Amount",
            hole=0.55, title="GST 2B ITC Distribution"
        )

        fig2.update_traces(textinfo="percent+label")

        fig2.update_layout(
            paper_bgcolor="white",
            plot_bgcolor="white",
            font_color="#111827"
        )

        st.plotly_chart(fig2, use_container_width=True)

    # =====================================================
    # SHEETS
    # =====================================================

    with st.expander("📂 Excel Sheets Automatically Analysed"):

        s1, s2 = st.columns(2)

        with s1:
            st.markdown("### GST 2B")
            for sheet in st.session_state["reco_two_b_sheets"]:
                st.write("•", sheet)

        with s2:
            st.markdown("### Books")
            for sheet in st.session_state["reco_books_sheets"]:
                st.write("•", sheet)

    # =====================================================
    # EXPORT
    # =====================================================

    st.markdown(
        '<div class="section">📥 Export Report</div>',
        unsafe_allow_html=True
    )

    excel_file = create_excel(result)

    st.download_button(
        "📊 Download Complete Excel Report",
        data=excel_file,
        file_name="GST_Reconciliation_Report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

else:

    st.markdown(
        """
        <div class="info-box">
        <h3>👋 Welcome to GST Reconciliation Pro</h3>
        Upload your <b>GST 2B</b> and <b>Books</b> Excel files.
        <br><br>
        The system will automatically identify useful GST columns
        and reconcile invoices using:
        <br><br>
        <b>GSTIN + Invoice Number</b>
        <br><br>
        It will identify Missing in 2B, Missing in Books,
        Value Mismatch and Matched invoices.
        </div>
        """,
        unsafe_allow_html=True
    )