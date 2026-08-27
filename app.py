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
    page_icon="🧾",
    layout="wide"
)

# =========================================================
# CSS
# =========================================================

st.markdown("""
<style>
.stApp {
    background-color: #f4f7fb;
}

[data-testid="stHeader"] {
    background-color: transparent;
}

.hero {
    background: linear-gradient(135deg, #0f2a5f, #1769aa);
    padding: 30px;
    border-radius: 20px;
    margin-bottom: 25px;
}

.hero h1 {
    color: white !important;
    font-size: 34px;
    margin: 0;
}

.hero p {
    color: #dbeafe !important;
    margin-top: 8px;
}

.section-title {
    color: #0f2a5f;
    font-size: 23px;
    font-weight: 800;
    margin-top: 25px;
    margin-bottom: 15px;
}

.kpi {
    background: white;
    border-radius: 16px;
    padding: 20px;
    border: 1px solid #dbe3ef;
    box-shadow: 0 5px 18px rgba(15,42,95,.07);
    min-height: 125px;
}

.kpi-title {
    color: #64748b;
    font-size: 12px;
    font-weight: 800;
}

.kpi-value {
    color: #172033;
    font-size: 30px;
    font-weight: 900;
    margin-top: 8px;
}

.kpi-desc {
    color: #94a3b8;
    font-size: 12px;
}

.blue {
    border-top: 4px solid #2563eb;
}

.green {
    border-top: 4px solid #16a34a;
}

.red {
    border-top: 4px solid #dc2626;
}

.orange {
    border-top: 4px solid #ea580c;
}

.purple {
    border-top: 4px solid #7c3aed;
}

.info-box {
    background: white;
    padding: 18px;
    border-radius: 13px;
    border-left: 5px solid #2563eb;
    border-top: 1px solid #dbe3ef;
    border-right: 1px solid #dbe3ef;
    border-bottom: 1px solid #dbe3ef;
    margin: 15px 0;
}

[data-testid="stFileUploaderDropzone"] {
    background: white;
    border: 2px dashed #b8c9df;
    border-radius: 14px;
}

.stButton button {
    border-radius: 10px;
    font-weight: 700;
}

.stDownloadButton button {
    background: #0f2a5f !important;
    color: white !important;
    border-radius: 10px;
    font-weight: 700;
}
</style>
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
        text = text.replace(",", "")
        text = text.replace("₹", "")
        text = text.replace("Rs.", "")
        text = text.replace("Rs", "")
        text = text.strip()

        return float(text)

    except:
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

    return pd.concat(
        all_data,
        ignore_index=True
    ), valid_sheets


# =========================================================
# STANDARDIZE
# =========================================================

def standardize(df):

    result = pd.DataFrame()

    detected = {}

    for key in ALIASES:

        detected[key] = find_column(
            df,
            ALIASES[key]
        )

    # GSTIN

    if detected["GSTIN"]:
        result["GSTIN"] = df[
            detected["GSTIN"]
        ].apply(clean_gstin)
    else:
        result["GSTIN"] = ""

    # Invoice

    if detected["Invoice"]:
        result["Invoice Number"] = df[
            detected["Invoice"]
        ].apply(clean_invoice)
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

        result["Party Name"] = df[
            detected["Party"]
        ].apply(clean_text)

    else:

        result["Party Name"] = ""

    # Taxable

    if detected["Taxable"]:

        result["Taxable Value"] = df[
            detected["Taxable"]
        ].apply(clean_amount)

    else:

        result["Taxable Value"] = 0.0

    # IGST

    if detected["IGST"]:

        result["IGST"] = df[
            detected["IGST"]
        ].apply(clean_amount)

    else:

        result["IGST"] = 0.0

    # CGST

    if detected["CGST"]:

        result["CGST"] = df[
            detected["CGST"]
        ].apply(clean_amount)

    else:

        result["CGST"] = 0.0

    # SGST

    if detected["SGST"]:

        result["SGST"] = df[
            detected["SGST"]
        ].apply(clean_amount)

    else:

        result["SGST"] = 0.0

    # Invoice Value

    if detected["InvoiceValue"]:

        result["Invoice Value"] = df[
            detected["InvoiceValue"]
        ].apply(clean_amount)

    else:

        result["Invoice Value"] = (
            result["Taxable Value"]
            + result["IGST"]
            + result["CGST"]
            + result["SGST"]
        )

    # Source sheet

    if "_SOURCE_SHEET" in df.columns:

        result["Source Sheet"] = df[
            "_SOURCE_SHEET"
        ]

    else:

        result["Source Sheet"] = ""

    # Remove blank invoices

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

        if row["_merge"] == "left_only":

            statuses.append("Missing in 2B")

            difference = (
                row["IGST_Books"]
                + row["CGST_Books"]
                + row["SGST_Books"]
            )

            differences.append(difference)

        elif row["_merge"] == "right_only":

            statuses.append("Missing in Books")

            difference = (
                row["IGST_2B"]
                + row["CGST_2B"]
                + row["SGST_2B"]
            )

            differences.append(difference)

        else:

            taxable_difference = abs(
                row["Taxable Value_Books"]
                -
                row["Taxable Value_2B"]
            )

            igst_difference = abs(
                row["IGST_Books"]
                -
                row["IGST_2B"]
            )

            cgst_difference = abs(
                row["CGST_Books"]
                -
                row["CGST_2B"]
            )

            sgst_difference = abs(
                row["SGST_Books"]
                -
                row["SGST_2B"]
            )

            invoice_difference = abs(
                row["Invoice Value_Books"]
                -
                row["Invoice Value_2B"]
            )

            total_difference = (
                igst_difference
                + cgst_difference
                + sgst_difference
            )

            if (
                taxable_difference <= tolerance
                and
                igst_difference <= tolerance
                and
                cgst_difference <= tolerance
                and
                sgst_difference <= tolerance
                and
                invoice_difference <= tolerance
            ):

                statuses.append("Matched")

            else:

                statuses.append("Value Mismatch")

            differences.append(
                total_difference
            )

    result["Status"] = statuses
    result["ITC Difference"] = differences

    return result


# =========================================================
# DISPLAY DATA
# =========================================================

def prepare_display(df):

    output = pd.DataFrame()

    def get_column(name):

        if name in df.columns:
            return df[name]

        return pd.Series(
            [""] * len(df),
            index=df.index
        )

    output["GSTIN"] = get_column(
        "GSTIN_Books"
    ).where(
        get_column("GSTIN_Books") != "",
        get_column("GSTIN_2B")
    )

    output["Party Name"] = get_column(
        "Party Name_Books"
    ).where(
        get_column("Party Name_Books") != "",
        get_column("Party Name_2B")
    )

    output["Invoice Number"] = get_column(
        "Invoice Number_Books"
    ).where(
        get_column("Invoice Number_Books") != "",
        get_column("Invoice Number_2B")
    )

    output["Invoice Date"] = get_column(
        "Invoice Date_Books"
    ).where(
        get_column("Invoice Date_Books").notna(),
        get_column("Invoice Date_2B")
    )

    output["Taxable - Books"] = get_column(
        "Taxable Value_Books"
    )

    output["Taxable - 2B"] = get_column(
        "Taxable Value_2B"
    )

    output["IGST - Books"] = get_column(
        "IGST_Books"
    )

    output["IGST - 2B"] = get_column(
        "IGST_2B"
    )

    output["CGST - Books"] = get_column(
        "CGST_Books"
    )

    output["CGST - 2B"] = get_column(
        "CGST_2B"
    )

    output["SGST - Books"] = get_column(
        "SGST_Books"
    )

    output["SGST - 2B"] = get_column(
        "SGST_2B"
    )

    output["Invoice Value - Books"] = get_column(
        "Invoice Value_Books"
    )

    output["Invoice Value - 2B"] = get_column(
        "Invoice Value_2B"
    )

    output["ITC Difference"] = get_column(
        "ITC Difference"
    )

    output["Status"] = get_column(
        "Status"
    )

    return output.reset_index(drop=True)


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

        statuses = [
            "Matched",
            "Missing in 2B",
            "Missing in Books",
            "Value Mismatch"
        ]

        for status in statuses:

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
        GST 2B vs Books • ITC Reconciliation • Invoice Exception Analysis
    </p>
</div>
""", unsafe_allow_html=True)


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

    st.subheader("📌 Status")

    st.write("🟢 Matched")
    st.write("🔴 Missing in 2B")
    st.write("🟠 Missing in Books")
    st.write("🟣 Value Mismatch")


# =========================================================
# UPLOAD
# =========================================================

st.markdown(
    '<div class="section-title">📂 Upload GST Data</div>',
    unsafe_allow_html=True
)

col1, col2 = st.columns(2)

with col1:

    st.subheader("📥 GST 2B File")

    st.caption(
        "Upload GST portal 2B Excel file. Multiple sheets supported."
    )

    two_b_file = st.file_uploader(
        "Choose GST 2B Excel",
        type=["xlsx", "xls"],
        key="gst_2b_upload"
    )


with col2:

    st.subheader("📚 Books File")

    st.caption(
        "Upload your purchase/books Excel file."
    )

    books_file = st.file_uploader(
        "Choose Books Excel",
        type=["xlsx", "xls"],
        key="books_upload"
    )


# =========================================================
# PROCESS
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
                st.session_state["two_b_sheets"] = two_b_sheets
                st.session_state["books_sheets"] = books_sheets
                st.session_state["selected_status"] = "Missing in 2B"

            st.success(
                "✅ GST Reconciliation completed successfully."
            )

        except Exception as error:

            st.error(
                "❌ File processing error"
            )

            st.error(
                str(error)
            )


# =========================================================
# DASHBOARD
# =========================================================

if "result" in st.session_state:

    result = st.session_state["result"]

    two_b = st.session_state["two_b"]

    books = st.session_state["books"]

    # =====================================================
    # COUNTS
    # =====================================================

    total = len(result)

    matched = int(
        (result["Status"] == "Matched").sum()
    )

    missing_2b = int(
        (result["Status"] == "Missing in 2B").sum()
    )

    missing_books = int(
        (result["Status"] == "Missing in Books").sum()
    )

    mismatch = int(
        (result["Status"] == "Value Mismatch").sum()
    )

    # =====================================================
    # DASHBOARD
    # =====================================================

    st.markdown(
        '<div class="section-title">📊 Reconciliation Dashboard</div>',
        unsafe_allow_html=True
    )

    c1, c2, c3, c4, c5 = st.columns(5)

    with c1:

        st.markdown(
            f"""
            <div class="kpi blue">
                <div class="kpi-title">TOTAL INVOICES</div>
                <div class="kpi-value">{total:,}</div>
                <div class="kpi-desc">All analysed records</div>
            </div>
            """,
            unsafe_allow_html=True
        )

    with c2:

        st.markdown(
            f"""
            <div class="kpi green">
                <div class="kpi-title">MATCHED</div>
                <div class="kpi-value">{matched:,}</div>
                <div class="kpi-desc">Successfully matched</div>
            </div>
            """,
            unsafe_allow_html=True
        )

    with c3:

        st.markdown(
            f"""
            <div class="kpi red">
                <div class="kpi-title">MISSING IN 2B</div>
                <div class="kpi-value">{missing_2b:,}</div>
                <div class="kpi-desc">Books → not in 2B</div>
            </div>
            """,
            unsafe_allow_html=True
        )

    with c4:

        st.markdown(
            f"""
            <div class="kpi orange">
                <div class="kpi-title">MISSING IN BOOKS</div>
                <div class="kpi-value">{missing_books:,}</div>
                <div class="kpi-desc">2B → not in Books</div>
            </div>
            """,
            unsafe_allow_html=True
        )

    with c5:

        st.markdown(
            f"""
            <div class="kpi purple">
                <div class="kpi-title">VALUE MISMATCH</div>
                <div class="kpi-value">{mismatch:,}</div>
                <div class="kpi-desc">Tax/value difference</div>
            </div>
            """,
            unsafe_allow_html=True
        )


    # =====================================================
    # INVOICE ANALYSIS
    # =====================================================

    st.markdown(
        '<div class="section-title">🔎 Invoice Analysis</div>',
        unsafe_allow_html=True
    )

    b1, b2, b3, b4 = st.columns(4)

    with b1:

        if st.button(
            f"🔴 Missing in 2B • {missing_2b}",
            use_container_width=True
        ):

            st.session_state["selected_status"] = "Missing in 2B"

    with b2:

        if st.button(
            f"🟠 Missing in Books • {missing_books}",
            use_container_width=True
        ):

            st.session_state["selected_status"] = "Missing in Books"

    with b3:

        if st.button(
            f"🟣 Value Mismatch • {mismatch}",
            use_container_width=True
        ):

            st.session_state["selected_status"] = "Value Mismatch"

    with b4:

        if st.button(
            f"🟢 Matched • {matched}",
            use_container_width=True
        ):

            st.session_state["selected_status"] = "Matched"


    selected_status = st.session_state.get(
        "selected_status",
        "Missing in 2B"
    )

    detail = result[
        result["Status"] == selected_status
    ].copy()


    st.subheader(
        f"📋 {selected_status} — {len(detail):,} Invoices"
    )


    explanations = {

        "Missing in 2B":
        "Books mein invoice hai, lekin GST 2B mein nahi mila.",

        "Missing in Books":
        "GST 2B mein invoice hai, lekin Books mein nahi mila.",

        "Value Mismatch":
        "GSTIN aur invoice number match hain, lekin taxable/tax/invoice value different hai.",

        "Matched":
        "Books aur GST 2B mein invoice successfully match hua."
    }


    st.info(
        explanations[selected_status]
    )


    # =====================================================
    # TABLE
    # =====================================================

    display = prepare_display(detail)

    if not display.empty:

        display["Invoice Date"] = pd.to_datetime(
            display["Invoice Date"],
            errors="coerce"
        ).dt.strftime("%d-%m-%Y")

        display["Invoice Date"] = display[
            "Invoice Date"
        ].fillna("-")

        st.dataframe(
            display,
            use_container_width=True,
            hide_index=True
        )

    else:

        st.success(
            "Is category mein koi invoice nahi hai."
        )


    # =====================================================
    # ITC ANALYSIS
    # =====================================================

    st.markdown(
        '<div class="section-title">💰 ITC Analysis</div>',
        unsafe_allow_html=True
    )

    books_igst = books["IGST"].sum()
    books_cgst = books["CGST"].sum()
    books_sgst = books["SGST"].sum()

    two_b_igst = two_b["IGST"].sum()
    two_b_cgst = two_b["CGST"].sum()
    two_b_sgst = two_b["SGST"].sum()

    books_itc = (
        books_igst
        + books_cgst
        + books_sgst
    )

    two_b_itc = (
        two_b_igst
        + two_b_cgst
        + two_b_sgst
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
    # TAX SUMMARY
    # =====================================================

    tax_df = pd.DataFrame({

        "Tax Component": [
            "IGST",
            "CGST",
            "SGST",
            "TOTAL ITC"
        ],

        "Books": [
            books_igst,
            books_cgst,
            books_sgst,
            books_itc
        ],

        "GST 2B": [
            two_b_igst,
            two_b_cgst,
            two_b_sgst,
            two_b_itc
        ],

        "Difference": [
            books_igst - two_b_igst,
            books_cgst - two_b_cgst,
            books_sgst - two_b_sgst,
            itc_difference
        ]
    })


    st.dataframe(
        tax_df.style.format({
            "Books": "₹{:,.2f}",
            "GST 2B": "₹{:,.2f}",
            "Difference": "₹{:,.2f}"
        }),
        use_container_width=True,
        hide_index=True
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

        status_df = pd.DataFrame({

            "Status": [
                "Matched",
                "Missing in 2B",
                "Missing in Books",
                "Value Mismatch"
            ],

            "Count": [
                matched,
                missing_2b,
                missing_books,
                mismatch
            ]
        })

        status_df = status_df[
            status_df["Count"] > 0
        ]

        if not status_df.empty:

            fig = px.pie(
                status_df,
                names="Status",
                values="Count",
                hole=0.55,
                title="Reconciliation Status"
            )

            fig.update_traces(
                textinfo="percent+label"
            )

            st.plotly_chart(
                fig,
                use_container_width=True
            )


    with chart2:

        itc_df = pd.DataFrame({

            "Tax": [
                "IGST",
                "CGST",
                "SGST"
            ],

            "Amount": [
                two_b_igst,
                two_b_cgst,
                two_b_sgst
            ]
        })

        itc_df = itc_df[
            itc_df["Amount"] > 0
        ]

        if not itc_df.empty:

            fig2 = px.pie(
                itc_df,
                names="Tax",
                values="Amount",
                hole=0.55,
                title="GST 2B ITC Distribution"
            )

            fig2.update_traces(
                textinfo="percent+label"
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

        col_a, col_b = st.columns(2)

        with col_a:

            st.subheader("📥 GST 2B")

            for sheet in st.session_state["two_b_sheets"]:

                st.write(
                    "•",
                    sheet
                )

        with col_b:

            st.subheader("📚 Books")

            for sheet in st.session_state["books_sheets"]:

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

    excel_report = create_excel(result)

    st.download_button(
        "📊 DOWNLOAD COMPLETE EXCEL REPORT",
        data=excel_report,
        file_name="GST_Reconciliation_Report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )


# =========================================================
# WELCOME
# =========================================================

else:

    st.markdown("""
    <div class="info-box">

    <h3>👋 Welcome to GST Reconciliation Pro</h3>

    Upload your <b>GST 2B</b> and <b>Books</b> Excel files.

    <br><br>

    The software automatically detects important GST columns
    even when the column names are different.

    <br><br>

    <b>Matching Key:</b>

    <br>

    GSTIN + Invoice Number

    <br><br>

    🔴 Missing in 2B<br>
    🟠 Missing in Books<br>
    🟣 Value Mismatch<br>
    🟢 Matched

    </div>
    """, unsafe_allow_html=True)
