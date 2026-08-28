import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
from io import BytesIO
import re
import difflib

from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

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
    letter-spacing: .03em;
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

.blue { border-top: 4px solid #2563eb; }
.green { border-top: 4px solid #16a34a; }
.red { border-top: 4px solid #dc2626; }
.orange { border-top: 4px solid #ea580c; }
.purple { border-top: 4px solid #7c3aed; }
.teal { border-top: 4px solid #0d9488; }

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

.warn-box {
    background: #fff7ed;
    padding: 14px 18px;
    border-radius: 13px;
    border-left: 5px solid #ea580c;
    margin: 10px 0;
    color: #7c2d12;
    font-size: 13px;
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

.badge {
    display: inline-block;
    padding: 2px 10px;
    border-radius: 999px;
    font-size: 11px;
    font-weight: 700;
}
</style>
""", unsafe_allow_html=True)

REQUIRED_FIELDS = ["GSTIN", "Invoice", "Party", "Taxable"]
OPTIONAL_FIELDS = ["Date", "IGST", "CGST", "SGST", "InvoiceValue"]

STATUS_COLORS = {
    "Matched": "16A34A",
    "Missing in 2B": "DC2626",
    "Missing in Books": "EA580C",
    "Value Mismatch": "7C3AED",
}

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
        if text in ("", "-", "nan"):
            return 0.0
        return float(text)
    except Exception:
        return 0.0


def is_valid_gstin(gstin):
    """Loose structural check: 15 chars, standard pattern."""
    if not gstin or len(gstin) != 15:
        return False
    pattern = r"^[0-9]{2}[A-Z]{5}[0-9]{4}[A-Z]{1}[1-9A-Z]{1}Z[0-9A-Z]{1}$"
    return bool(re.match(pattern, gstin))


# =========================================================
# COLUMN ALIASES
# =========================================================

ALIASES = {
    "GSTIN": [
        "GSTIN", "GSTIN/UIN", "GSTIN of Supplier", "GSTIN of supplier",
        "Supplier GSTIN", "GST Number", "GST No", "GST No.",
        "GSTIN of Vendor", "GSTIN/UIN of Supplier", "Vendor GSTIN"
    ],
    "Invoice": [
        "Invoice Number", "Invoice number", "Invoice No", "Invoice No.",
        "Document Number", "Document No", "Bill No", "Bill Number", "Invoice"
    ],
    "Date": [
        "Invoice Date", "Invoice date", "Document Date", "Document date",
        "Bill Date", "Date"
    ],
    "Party": [
        "Party Name", "Supplier Name", "Supplier", "Vendor Name", "Vendor",
        "Trade Name", "Legal Name", "Party", "Supplier Trade Name"
    ],
    "Taxable": [
        "Taxable Value", "Taxable value", "Taxable Amount",
        "Taxable Amount (₹)", "Taxable Amt", "Taxable"
    ],
    "IGST": [
        "IGST", "IGST Amount", "IGST Amt", "Integrated Tax",
        "Integrated Tax Amount", "Integrated Tax Amount (₹)"
    ],
    "CGST": [
        "CGST", "CGST Amount", "CGST Amt", "Central Tax",
        "Central Tax Amount", "Central Tax Amount (₹)"
    ],
    "SGST": [
        "SGST", "SGST Amount", "SGST Amt", "UTGST", "State Tax",
        "State Tax Amount", "State/UT Tax", "State/UT Tax Amount",
        "State/UT Tax Amount (₹)"
    ],
    "InvoiceValue": [
        "Invoice Value", "Invoice value", "Total Invoice Value",
        "Total Value", "Document Value", "Invoice Amount",
        "Total Invoice Amount", "Invoice Value (₹)"
    ]
}

FIELD_LABELS = {
    "GSTIN": "GSTIN",
    "Invoice": "Invoice Number",
    "Date": "Invoice Date",
    "Party": "Party / Supplier Name",
    "Taxable": "Taxable Value",
    "IGST": "IGST",
    "CGST": "CGST",
    "SGST": "SGST",
    "InvoiceValue": "Invoice Value (Total)",
}


def normalize_column(value):
    return re.sub(r"[^A-Z0-9]", "", str(value).upper())


def find_column(df, aliases):
    columns = {normalize_column(c): c for c in df.columns}
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


def auto_detect(df):
    return {key: find_column(df, ALIASES[key]) for key in ALIASES}


# =========================================================
# READ EXCEL
# =========================================================

def read_excel(file):
    excel = pd.ExcelFile(file)
    all_data = []
    valid_sheets = []
    skipped_sheets = []

    for sheet in excel.sheet_names:
        try:
            df = pd.read_excel(file, sheet_name=sheet)
            if df.empty:
                skipped_sheets.append(sheet)
                continue

            gstin_column = find_column(df, ALIASES["GSTIN"])
            invoice_column = find_column(df, ALIASES["Invoice"])

            if gstin_column and invoice_column:
                df["_SOURCE_SHEET"] = sheet
                all_data.append(df)
                valid_sheets.append(sheet)
            else:
                skipped_sheets.append(sheet)
        except Exception:
            skipped_sheets.append(sheet)
            continue

    if not all_data:
        raise ValueError("GSTIN aur Invoice Number wale columns kisi bhi sheet mein nahi mile.")

    combined = pd.concat(all_data, ignore_index=True)
    return combined, valid_sheets, skipped_sheets


# =========================================================
# STANDARDIZE  (accepts optional manual column overrides)
# =========================================================

def standardize(df, overrides=None):
    """
    overrides: dict like {"GSTIN": "Some Column Name", ...} to force a
    specific source column instead of auto-detection. Use "— None —" or
    None to mean 'not mapped'.
    """
    overrides = overrides or {}
    result = pd.DataFrame()
    detected = {}

    for key in ALIASES:
        manual = overrides.get(key)
        if manual and manual != "— None —" and manual in df.columns:
            detected[key] = manual
        else:
            detected[key] = find_column(df, ALIASES[key])

    if detected["GSTIN"]:
        result["GSTIN"] = df[detected["GSTIN"]].apply(clean_gstin)
    else:
        result["GSTIN"] = ""

    if detected["Invoice"]:
        result["Invoice Number"] = df[detected["Invoice"]].apply(clean_invoice)
    else:
        result["Invoice Number"] = ""

    if detected["Date"]:
        result["Invoice Date"] = pd.to_datetime(df[detected["Date"]], errors="coerce")
    else:
        result["Invoice Date"] = pd.NaT

    if detected["Party"]:
        result["Party Name"] = df[detected["Party"]].apply(clean_text)
    else:
        result["Party Name"] = ""

    if detected["Taxable"]:
        result["Taxable Value"] = df[detected["Taxable"]].apply(clean_amount)
    else:
        result["Taxable Value"] = 0.0

    if detected["IGST"]:
        result["IGST"] = df[detected["IGST"]].apply(clean_amount)
    else:
        result["IGST"] = 0.0

    if detected["CGST"]:
        result["CGST"] = df[detected["CGST"]].apply(clean_amount)
    else:
        result["CGST"] = 0.0

    if detected["SGST"]:
        result["SGST"] = df[detected["SGST"]].apply(clean_amount)
    else:
        result["SGST"] = 0.0

    if detected["InvoiceValue"]:
        result["Invoice Value"] = df[detected["InvoiceValue"]].apply(clean_amount)
    else:
        result["Invoice Value"] = (
            result["Taxable Value"] + result["IGST"] + result["CGST"] + result["SGST"]
        )

    result["Source Sheet"] = df["_SOURCE_SHEET"] if "_SOURCE_SHEET" in df.columns else ""

    total_rows = len(result)
    blank_mask = (result["GSTIN"] == "") | (result["Invoice Number"] == "")
    dropped_rows = int(blank_mask.sum())

    cleaned = result[~blank_mask].copy()

    # Duplicate invoice detection (same GSTIN + Invoice appearing >1 time)
    dup_mask = cleaned.duplicated(subset=["GSTIN", "Invoice Number"], keep=False)
    duplicate_count = int(dup_mask.sum())

    invalid_gstin_count = int(cleaned["GSTIN"].apply(lambda g: not is_valid_gstin(g)).sum())

    quality = {
        "total_rows": total_rows,
        "dropped_rows": dropped_rows,
        "duplicate_rows": duplicate_count,
        "invalid_gstin_rows": invalid_gstin_count,
        "detected_columns": detected,
    }

    return cleaned.reset_index(drop=True), quality


# =========================================================
# RECONCILIATION
# =========================================================

def build_dedup_key(df):
    """
    GSTIN + Invoice Number can repeat (rare, but real files have it).
    A plain merge on that key would create a cross-product of every
    duplicate pair, which silently corrupts the reconciliation.
    We instead pair up duplicates in the order they appear (1st with 1st,
    2nd with 2nd, ...), which mirrors how an accountant would actually
    tie invoices together.
    """
    df = df.copy()
    base_key = df["GSTIN"] + "|" + df["Invoice Number"]
    occurrence = base_key.groupby(base_key).cumcount()
    df["KEY"] = base_key + "|" + occurrence.astype(str)
    return df


def suggest_close_matches(missing_row, other_df, field="Invoice Number", cutoff=0.82):
    """For an unmatched invoice, look for a same-GSTIN invoice in the other
    file whose number is *almost* identical (formatting / typo difference)
    and flag it as a probable match rather than a genuine gap."""
    same_gstin = other_df[other_df["GSTIN"] == missing_row["GSTIN"]]
    if same_gstin.empty:
        return None
    candidates = same_gstin[field].tolist()
    best = difflib.get_close_matches(missing_row[field], candidates, n=1, cutoff=cutoff)
    if best:
        match_row = same_gstin[same_gstin[field] == best[0]].iloc[0]
        return best[0], match_row.get("Invoice Value", None)
    return None


def reconcile(two_b, books, tolerance, enable_fuzzy=True):
    two_b_k = build_dedup_key(two_b)
    books_k = build_dedup_key(books)

    result = pd.merge(
        books_k, two_b_k, on="KEY", how="outer",
        suffixes=("_Books", "_2B"), indicator=True
    )

    statuses = []
    differences = []
    suggestions = []

    for _, row in result.iterrows():
        if row["_merge"] == "left_only":
            statuses.append("Missing in 2B")
            differences.append(row["IGST_Books"] + row["CGST_Books"] + row["SGST_Books"])
            suggestion = ""
            if enable_fuzzy:
                probe = {"GSTIN": row["GSTIN_Books"], "Invoice Number": row["Invoice Number_Books"]}
                found = suggest_close_matches(probe, two_b)
                if found:
                    suggestion = f"Possible match in 2B: '{found[0]}' — check formatting/typo"
            suggestions.append(suggestion)

        elif row["_merge"] == "right_only":
            statuses.append("Missing in Books")
            differences.append(row["IGST_2B"] + row["CGST_2B"] + row["SGST_2B"])
            suggestion = ""
            if enable_fuzzy:
                probe = {"GSTIN": row["GSTIN_2B"], "Invoice Number": row["Invoice Number_2B"]}
                found = suggest_close_matches(probe, books)
                if found:
                    suggestion = f"Possible match in Books: '{found[0]}' — check formatting/typo"
            suggestions.append(suggestion)

        else:
            taxable_diff = abs(row["Taxable Value_Books"] - row["Taxable Value_2B"])
            igst_diff = abs(row["IGST_Books"] - row["IGST_2B"])
            cgst_diff = abs(row["CGST_Books"] - row["CGST_2B"])
            sgst_diff = abs(row["SGST_Books"] - row["SGST_2B"])
            invoice_diff = abs(row["Invoice Value_Books"] - row["Invoice Value_2B"])
            total_diff = igst_diff + cgst_diff + sgst_diff

            if (taxable_diff <= tolerance and igst_diff <= tolerance and
                    cgst_diff <= tolerance and sgst_diff <= tolerance and
                    invoice_diff <= tolerance):
                statuses.append("Matched")
            else:
                statuses.append("Value Mismatch")

            differences.append(total_diff)
            suggestions.append("")

    result["Status"] = statuses
    result["ITC Difference"] = differences
    result["Match Suggestion"] = suggestions

    return result


# =========================================================
# DISPLAY DATA
# =========================================================

def prepare_display(df):
    output = pd.DataFrame()

    def get_column(name):
        if name in df.columns:
            return df[name]
        return pd.Series([""] * len(df), index=df.index)

    # NOTE: an outer merge fills the *missing* side with NaN, not "" — so the
    # fallback must treat NaN as empty too, or it silently keeps the NaN
    # instead of falling back to the other file's value.
    gstin_books = get_column("GSTIN_Books").fillna("")
    gstin_2b = get_column("GSTIN_2B").fillna("")
    output["GSTIN"] = gstin_books.where(gstin_books != "", gstin_2b)

    party_books = get_column("Party Name_Books").fillna("")
    party_2b = get_column("Party Name_2B").fillna("")
    output["Party Name"] = party_books.where(party_books != "", party_2b)

    inv_books = get_column("Invoice Number_Books").fillna("")
    inv_2b = get_column("Invoice Number_2B").fillna("")
    output["Invoice Number"] = inv_books.where(inv_books != "", inv_2b)
    output["Invoice Date"] = get_column("Invoice Date_Books").where(get_column("Invoice Date_Books").notna(), get_column("Invoice Date_2B"))
    output["Taxable - Books"] = get_column("Taxable Value_Books")
    output["Taxable - 2B"] = get_column("Taxable Value_2B")
    output["IGST - Books"] = get_column("IGST_Books")
    output["IGST - 2B"] = get_column("IGST_2B")
    output["CGST - Books"] = get_column("CGST_Books")
    output["CGST - 2B"] = get_column("CGST_2B")
    output["SGST - Books"] = get_column("SGST_Books")
    output["SGST - 2B"] = get_column("SGST_2B")
    output["Invoice Value - Books"] = get_column("Invoice Value_Books")
    output["Invoice Value - 2B"] = get_column("Invoice Value_2B")
    output["ITC Difference"] = get_column("ITC Difference")
    output["Status"] = get_column("Status")
    output["Match Suggestion"] = get_column("Match Suggestion")

    return output.reset_index(drop=True)


def vendor_summary(display_df):
    """GSTIN-wise rollup — the fastest way to see which vendors are causing
    the most reconciliation trouble."""
    grouped = display_df.groupby(["GSTIN", "Party Name"], dropna=False).agg(
        Total_Invoices=("Invoice Number", "count"),
        Matched=("Status", lambda s: (s == "Matched").sum()),
        Missing_in_2B=("Status", lambda s: (s == "Missing in 2B").sum()),
        Missing_in_Books=("Status", lambda s: (s == "Missing in Books").sum()),
        Value_Mismatch=("Status", lambda s: (s == "Value Mismatch").sum()),
        ITC_at_Risk=("ITC Difference", "sum"),
    ).reset_index()

    grouped = grouped.sort_values("ITC_at_Risk", ascending=False)
    grouped["ITC_at_Risk"] = grouped["ITC_at_Risk"].round(2)
    return grouped


# =========================================================
# EXCEL REPORT (formatted)
# =========================================================

def _style_worksheet_header(ws, ncols):
    header_fill = PatternFill(start_color="0F2A5F", end_color="0F2A5F", fill_type="solid")
    header_font = Font(name="Arial", bold=True, color="FFFFFF", size=10)
    thin = Side(style="thin", color="D9D9D9")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for col in range(1, ncols + 1):
        c = ws.cell(row=1, column=col)
        c.fill = header_fill
        c.font = header_font
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = border
    ws.freeze_panes = "A2"


def _autofit(ws, max_width=45):
    for col_cells in ws.columns:
        length = max((len(str(c.value)) if c.value is not None else 0) for c in col_cells)
        col_letter = get_column_letter(col_cells[0].column)
        ws.column_dimensions[col_letter].width = min(max(length + 3, 10), max_width)


def _format_data_rows(ws, ncols, status_col_index=None):
    thin = Side(style="thin", color="EDEDED")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    money_cols = set()
    for idx, cell in enumerate(ws[1], start=1):
        header = str(cell.value or "")
        if any(tag in header for tag in ["Taxable", "IGST", "CGST", "SGST", "Invoice Value", "ITC"]):
            money_cols.add(idx)

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=ncols):
        for cell in row:
            cell.font = Font(name="Arial", size=10)
            cell.border = border
            if cell.column in money_cols:
                cell.number_format = "#,##0.00"

        if status_col_index:
            status_val = row[status_col_index - 1].value
            color = STATUS_COLORS.get(status_val)
            if color:
                fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
                font = Font(name="Arial", size=9, bold=True, color="FFFFFF")
                cell = row[status_col_index - 1]
                cell.fill = fill
                cell.font = font
                cell.alignment = Alignment(horizontal="center")


def create_excel(result, quality_2b, quality_books, tolerance):
    display = prepare_display(result)
    v_summary = vendor_summary(display)

    total = len(display)
    matched = int((display["Status"] == "Matched").sum())
    missing_2b = int((display["Status"] == "Missing in 2B").sum())
    missing_books = int((display["Status"] == "Missing in Books").sum())
    mismatch = int((display["Status"] == "Value Mismatch").sum())
    itc_at_risk = display.loc[display["Status"] != "Matched", "ITC Difference"].sum()

    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:

        summary_df = pd.DataFrame({
            "Metric": [
                "Total Invoices Analysed", "Matched", "Missing in 2B",
                "Missing in Books", "Value Mismatch", "Match Rate (%)",
                "Total ITC at Risk (Rs.)", "Mismatch Tolerance Used (Rs.)",
            ],
            "Value": [
                total, matched, missing_2b, missing_books, mismatch,
                round((matched / total * 100), 2) if total else 0,
                round(itc_at_risk, 2), tolerance,
            ]
        })
        summary_df.to_excel(writer, sheet_name="Summary", index=False)

        display_export = display.copy()
        if "Invoice Date" in display_export.columns:
            display_export["Invoice Date"] = pd.to_datetime(
                display_export["Invoice Date"], errors="coerce"
            ).dt.strftime("%d-%m-%Y")
        display_export.to_excel(writer, sheet_name="Complete Reconciliation", index=False)

        for status in ["Matched", "Missing in 2B", "Missing in Books", "Value Mismatch"]:
            temp = display_export[display_export["Status"] == status]
            temp.to_excel(writer, sheet_name=status[:31], index=False)

        v_summary.to_excel(writer, sheet_name="Vendor Summary", index=False)

    output.seek(0)

    # ---- post-process formatting with openpyxl ----
    wb = load_workbook(output)

    ws_summary = wb["Summary"]
    _style_worksheet_header(ws_summary, ws_summary.max_column)
    _format_data_rows(ws_summary, ws_summary.max_column)
    _autofit(ws_summary)

    main_sheet_names = ["Complete Reconciliation", "Matched", "Missing in 2B",
                        "Missing in Books", "Value Mismatch"]
    for name in main_sheet_names:
        ws = wb[name]
        if ws.max_row < 1:
            continue
        _style_worksheet_header(ws, ws.max_column)
        status_idx = None
        for i, cell in enumerate(ws[1], start=1):
            if cell.value == "Status":
                status_idx = i
                break
        _format_data_rows(ws, ws.max_column, status_col_index=status_idx)
        _autofit(ws)

    ws_vendor = wb["Vendor Summary"]
    _style_worksheet_header(ws_vendor, ws_vendor.max_column)
    _format_data_rows(ws_vendor, ws_vendor.max_column)
    _autofit(ws_vendor)

    final_output = BytesIO()
    wb.save(final_output)
    final_output.seek(0)
    return final_output


# =========================================================
# HEADER
# =========================================================

st.markdown("""
<div class="hero">
    <h1>🧾 GST Reconciliation Pro</h1>
    <p>GST 2B vs Books • ITC Reconciliation • Vendor & Invoice Exception Analysis</p>
</div>
""", unsafe_allow_html=True)

# =========================================================
# SIDEBAR
# =========================================================

with st.sidebar:
    st.header("⚙️ Reconciliation Settings")

    tolerance = st.number_input(
        "Mismatch Tolerance ₹", min_value=0.0, value=2.0, step=0.50,
        help="Differences within this amount are still treated as a Match."
    )

    enable_fuzzy = st.checkbox(
        "🔍 Suggest possible matches for gaps", value=True,
        help="Flags likely formatting/typo differences instead of a genuine missing invoice."
    )

    st.divider()
    st.subheader("🔑 Matching Key")
    st.info("GSTIN + Invoice Number\n\n(duplicate invoices are paired in sequence)")

    st.divider()
    st.subheader("📌 Status Legend")
    st.markdown("🟢 **Matched** — tallies within tolerance")
    st.markdown("🔴 **Missing in 2B** — in Books, not filed by supplier")
    st.markdown("🟠 **Missing in Books** — filed by supplier, not booked")
    st.markdown("🟣 **Value Mismatch** — amounts differ beyond tolerance")

# =========================================================
# UPLOAD
# =========================================================

st.markdown('<div class="section-title">📂 Upload GST Data</div>', unsafe_allow_html=True)

col1, col2 = st.columns(2)

with col1:
    st.subheader("📥 GST 2B File")
    st.caption("Upload GST portal 2B Excel file. Multiple sheets supported.")
    two_b_file = st.file_uploader("Choose GST 2B Excel", type=["xlsx", "xls"], key="gst_2b_upload")

with col2:
    st.subheader("📚 Books File")
    st.caption("Upload your purchase/books Excel file.")
    books_file = st.file_uploader("Choose Books Excel", type=["xlsx", "xls"], key="books_upload")

# =========================================================
# LOAD + OPTIONAL MANUAL COLUMN MAPPING
# =========================================================

overrides_2b, overrides_books = {}, {}

if two_b_file and books_file:
    try:
        if ("two_b_raw" not in st.session_state or
                st.session_state.get("two_b_name") != two_b_file.name):
            raw, sheets, skipped = read_excel(two_b_file)
            st.session_state["two_b_raw"] = raw
            st.session_state["two_b_sheets"] = sheets
            st.session_state["two_b_skipped"] = skipped
            st.session_state["two_b_name"] = two_b_file.name

        if ("books_raw" not in st.session_state or
                st.session_state.get("books_name") != books_file.name):
            raw, sheets, skipped = read_excel(books_file)
            st.session_state["books_raw"] = raw
            st.session_state["books_sheets"] = sheets
            st.session_state["books_skipped"] = skipped
            st.session_state["books_name"] = books_file.name

        two_b_raw = st.session_state["two_b_raw"]
        books_raw = st.session_state["books_raw"]

        with st.expander("🔧 Column Mapping (verify or override auto-detection)"):
            st.caption(
                "Har alag Excel export column names alag rakhta hai — humne "
                "best-guess mapping kar di hai. Zaroorat ho to yahan se change karein."
            )
            m1, m2 = st.columns(2)

            with m1:
                st.markdown("**GST 2B columns**")
                auto_2b = auto_detect(two_b_raw)
                options_2b = ["— None —"] + list(two_b_raw.columns)
                for field in REQUIRED_FIELDS + OPTIONAL_FIELDS:
                    default_val = auto_2b.get(field) or "— None —"
                    default_idx = options_2b.index(default_val) if default_val in options_2b else 0
                    overrides_2b[field] = st.selectbox(
                        FIELD_LABELS[field], options_2b, index=default_idx, key=f"2b_{field}"
                    )

            with m2:
                st.markdown("**Books columns**")
                auto_books = auto_detect(books_raw)
                options_books = ["— None —"] + list(books_raw.columns)
                for field in REQUIRED_FIELDS + OPTIONAL_FIELDS:
                    default_val = auto_books.get(field) or "— None —"
                    default_idx = options_books.index(default_val) if default_val in options_books else 0
                    overrides_books[field] = st.selectbox(
                        FIELD_LABELS[field], options_books, index=default_idx, key=f"bk_{field}"
                    )

    except Exception as error:
        st.error("❌ File padhne mein error aayi")
        st.error(str(error))

# =========================================================
# PROCESS
# =========================================================

if two_b_file and books_file and "two_b_raw" in st.session_state:
    st.write("")

    if st.button("🚀 START GST RECONCILIATION", type="primary", use_container_width=True):
        try:
            with st.spinner("Analysing GST 2B and Books..."):
                two_b, quality_2b = standardize(st.session_state["two_b_raw"], overrides_2b)
                books, quality_books = standardize(st.session_state["books_raw"], overrides_books)

                if two_b.empty:
                    raise ValueError("GST 2B file mein valid invoice data nahi mila.")
                if books.empty:
                    raise ValueError("Books file mein valid invoice data nahi mila.")

                result = reconcile(two_b, books, tolerance, enable_fuzzy=enable_fuzzy)

                st.session_state["result"] = result
                st.session_state["two_b"] = two_b
                st.session_state["books"] = books
                st.session_state["quality_2b"] = quality_2b
                st.session_state["quality_books"] = quality_books
                st.session_state["tolerance_used"] = tolerance
                st.session_state["selected_status"] = "Missing in 2B"

            st.success("✅ GST Reconciliation completed successfully.")

        except Exception as error:
            st.error("❌ File processing error")
            st.error(str(error))

# =========================================================
# DATA QUALITY PANEL
# =========================================================

if "quality_2b" in st.session_state and "result" in st.session_state:
    q2b = st.session_state["quality_2b"]
    qbk = st.session_state["quality_books"]

    quality_flags = []
    if q2b["dropped_rows"] or qbk["dropped_rows"]:
        quality_flags.append(
            f"⚠️ Rows skipped due to blank GSTIN/Invoice — 2B: {q2b['dropped_rows']}, Books: {qbk['dropped_rows']}"
        )
    if q2b["duplicate_rows"] or qbk["duplicate_rows"]:
        quality_flags.append(
            f"⚠️ Duplicate GSTIN+Invoice rows detected — 2B: {q2b['duplicate_rows']}, Books: {qbk['duplicate_rows']} "
            "(paired in sequence during matching)"
        )
    if q2b["invalid_gstin_rows"] or qbk["invalid_gstin_rows"]:
        quality_flags.append(
            f"⚠️ Rows with non-standard GSTIN format — 2B: {q2b['invalid_gstin_rows']}, Books: {qbk['invalid_gstin_rows']}"
        )

    with st.expander("🩺 Data Quality Checks", expanded=bool(quality_flags)):
        if quality_flags:
            for flag in quality_flags:
                st.markdown(f'<div class="warn-box">{flag}</div>', unsafe_allow_html=True)
        else:
            st.success("Koi data quality issue nahi mila — files clean hain.")

# =========================================================
# DASHBOARD
# =========================================================

if "result" in st.session_state:
    result = st.session_state["result"]
    two_b = st.session_state["two_b"]
    books = st.session_state["books"]

    total = len(result)
    matched = int((result["Status"] == "Matched").sum())
    missing_2b = int((result["Status"] == "Missing in 2B").sum())
    missing_books = int((result["Status"] == "Missing in Books").sum())
    mismatch = int((result["Status"] == "Value Mismatch").sum())
    match_rate = round((matched / total * 100), 1) if total else 0.0

    st.markdown('<div class="section-title">📊 Reconciliation Dashboard</div>', unsafe_allow_html=True)

    c1, c2, c3, c4, c5, c6 = st.columns(6)

    with c1:
        st.markdown(f"""<div class="kpi blue"><div class="kpi-title">TOTAL INVOICES</div>
        <div class="kpi-value">{total:,}</div><div class="kpi-desc">All analysed records</div></div>""",
        unsafe_allow_html=True)

    with c2:
        st.markdown(f"""<div class="kpi green"><div class="kpi-title">MATCHED</div>
        <div class="kpi-value">{matched:,}</div><div class="kpi-desc">{match_rate}% match rate</div></div>""",
        unsafe_allow_html=True)

    with c3:
        st.markdown(f"""<div class="kpi red"><div class="kpi-title">MISSING IN 2B</div>
        <div class="kpi-value">{missing_2b:,}</div><div class="kpi-desc">Books → not in 2B</div></div>""",
        unsafe_allow_html=True)

    with c4:
        st.markdown(f"""<div class="kpi orange"><div class="kpi-title">MISSING IN BOOKS</div>
        <div class="kpi-value">{missing_books:,}</div><div class="kpi-desc">2B → not in Books</div></div>""",
        unsafe_allow_html=True)

    with c5:
        st.markdown(f"""<div class="kpi purple"><div class="kpi-title">VALUE MISMATCH</div>
        <div class="kpi-value">{mismatch:,}</div><div class="kpi-desc">Tax/value difference</div></div>""",
        unsafe_allow_html=True)

    with c6:
        itc_at_risk = result.loc[result["Status"] != "Matched", "ITC Difference"].sum()
        st.markdown(f"""<div class="kpi teal"><div class="kpi-title">ITC AT RISK</div>
        <div class="kpi-value">₹{itc_at_risk:,.0f}</div><div class="kpi-desc">Sum of open differences</div></div>""",
        unsafe_allow_html=True)

    # =====================================================
    # INVOICE ANALYSIS
    # =====================================================

    st.markdown('<div class="section-title">🔎 Invoice Analysis</div>', unsafe_allow_html=True)

    b1, b2, b3, b4 = st.columns(4)
    with b1:
        if st.button(f"🔴 Missing in 2B • {missing_2b}", use_container_width=True):
            st.session_state["selected_status"] = "Missing in 2B"
    with b2:
        if st.button(f"🟠 Missing in Books • {missing_books}", use_container_width=True):
            st.session_state["selected_status"] = "Missing in Books"
    with b3:
        if st.button(f"🟣 Value Mismatch • {mismatch}", use_container_width=True):
            st.session_state["selected_status"] = "Value Mismatch"
    with b4:
        if st.button(f"🟢 Matched • {matched}", use_container_width=True):
            st.session_state["selected_status"] = "Matched"

    selected_status = st.session_state.get("selected_status", "Missing in 2B")
    detail = result[result["Status"] == selected_status].copy()

    st.subheader(f"📋 {selected_status} — {len(detail):,} Invoices")

    explanations = {
        "Missing in 2B": "Books mein invoice hai, lekin GST 2B mein nahi mila.",
        "Missing in Books": "GST 2B mein invoice hai, lekin Books mein nahi mila.",
        "Value Mismatch": "GSTIN aur invoice number match hain, lekin taxable/tax/invoice value different hai.",
        "Matched": "Books aur GST 2B mein invoice successfully match hua.",
    }
    st.info(explanations[selected_status])

    display = prepare_display(detail)

    search_term = st.text_input(
        "🔍 Search within this list (GSTIN, Invoice No, or Party Name)",
        placeholder="e.g. 27AABCU9603R1ZM or INV/2025/1001"
    )

    if not display.empty:
        display["Invoice Date"] = pd.to_datetime(display["Invoice Date"], errors="coerce").dt.strftime("%d-%m-%Y")
        display["Invoice Date"] = display["Invoice Date"].fillna("-")

        if search_term.strip():
            term = search_term.strip().lower()
            mask = (
                display["GSTIN"].str.lower().str.contains(term, na=False) |
                display["Invoice Number"].str.lower().str.contains(term, na=False) |
                display["Party Name"].str.lower().str.contains(term, na=False)
            )
            display = display[mask]

        st.dataframe(display, use_container_width=True, hide_index=True)

        csv_bytes = display.to_csv(index=False).encode("utf-8-sig")
        st.download_button(
            f"⬇️ Download '{selected_status}' as CSV",
            data=csv_bytes,
            file_name=f"GST_{selected_status.replace(' ', '_')}.csv",
            mime="text/csv",
        )
    else:
        st.success("Is category mein koi invoice nahi hai.")

    # =====================================================
    # VENDOR-WISE SUMMARY
    # =====================================================

    st.markdown('<div class="section-title">🏢 Vendor-wise Summary</div>', unsafe_allow_html=True)
    st.caption("Sabse zyada ITC-at-risk wale vendors sabse upar dikhaye gaye hain.")

    full_display = prepare_display(result)
    v_summary = vendor_summary(full_display)

    st.dataframe(
        v_summary.rename(columns={
            "Total_Invoices": "Total Invoices",
            "Missing_in_2B": "Missing in 2B",
            "Missing_in_Books": "Missing in Books",
            "Value_Mismatch": "Value Mismatch",
            "ITC_at_Risk": "ITC at Risk (Rs.)",
        }).style.format({"ITC at Risk (Rs.)": "₹{:,.2f}"}),
        use_container_width=True,
        hide_index=True,
    )

    # =====================================================
    # ITC ANALYSIS
    # =====================================================

    st.markdown('<div class="section-title">💰 ITC Analysis</div>', unsafe_allow_html=True)

    books_igst, books_cgst, books_sgst = books["IGST"].sum(), books["CGST"].sum(), books["SGST"].sum()
    two_b_igst, two_b_cgst, two_b_sgst = two_b["IGST"].sum(), two_b["CGST"].sum(), two_b["SGST"].sum()
    books_itc = books_igst + books_cgst + books_sgst
    two_b_itc = two_b_igst + two_b_cgst + two_b_sgst
    itc_difference = books_itc - two_b_itc

    i1, i2, i3 = st.columns(3)
    with i1:
        st.metric("📚 Books Total ITC", f"₹{books_itc:,.2f}")
    with i2:
        st.metric("📥 GST 2B Total ITC", f"₹{two_b_itc:,.2f}")
    with i3:
        st.metric("⚖️ ITC Difference", f"₹{itc_difference:,.2f}")

    tax_df = pd.DataFrame({
        "Tax Component": ["IGST", "CGST", "SGST", "TOTAL ITC"],
        "Books": [books_igst, books_cgst, books_sgst, books_itc],
        "GST 2B": [two_b_igst, two_b_cgst, two_b_sgst, two_b_itc],
        "Difference": [
            books_igst - two_b_igst, books_cgst - two_b_cgst,
            books_sgst - two_b_sgst, itc_difference
        ]
    })

    st.dataframe(
        tax_df.style.format({"Books": "₹{:,.2f}", "GST 2B": "₹{:,.2f}", "Difference": "₹{:,.2f}"}),
        use_container_width=True, hide_index=True
    )

    # =====================================================
    # CHARTS
    # =====================================================

    st.markdown('<div class="section-title">📈 Visual Analysis</div>', unsafe_allow_html=True)

    chart1, chart2 = st.columns(2)

    with chart1:
        status_df = pd.DataFrame({
            "Status": ["Matched", "Missing in 2B", "Missing in Books", "Value Mismatch"],
            "Count": [matched, missing_2b, missing_books, mismatch]
        })
        status_df = status_df[status_df["Count"] > 0]
        if not status_df.empty:
            fig = px.pie(
                status_df, names="Status", values="Count", hole=0.55,
                title="Reconciliation Status",
                color="Status",
                color_discrete_map={
                    "Matched": "#16a34a", "Missing in 2B": "#dc2626",
                    "Missing in Books": "#ea580c", "Value Mismatch": "#7c3aed",
                }
            )
            fig.update_traces(textinfo="percent+label")
            st.plotly_chart(fig, use_container_width=True)

    with chart2:
        top_vendors = v_summary[v_summary["ITC_at_Risk"] > 0].head(10)
        if not top_vendors.empty:
            fig3 = px.bar(
                top_vendors.sort_values("ITC_at_Risk"),
                x="ITC_at_Risk", y="Party Name", orientation="h",
                title="Top 10 Vendors by ITC at Risk",
                labels={"ITC_at_Risk": "ITC at Risk (Rs.)", "Party Name": ""},
            )
            fig3.update_traces(marker_color="#7c3aed")
            st.plotly_chart(fig3, use_container_width=True)
        else:
            itc_df = pd.DataFrame({
                "Tax": ["IGST", "CGST", "SGST"],
                "Amount": [two_b_igst, two_b_cgst, two_b_sgst]
            })
            itc_df = itc_df[itc_df["Amount"] > 0]
            if not itc_df.empty:
                fig2 = px.pie(itc_df, names="Tax", values="Amount", hole=0.55, title="GST 2B ITC Distribution")
                fig2.update_traces(textinfo="percent+label")
                st.plotly_chart(fig2, use_container_width=True)

    # Monthly trend, if invoice dates are usable
    combined_dates = pd.concat([
        two_b[["Invoice Date", "Taxable Value"]].assign(Source="GST 2B"),
        books[["Invoice Date", "Taxable Value"]].assign(Source="Books"),
    ])
    combined_dates = combined_dates.dropna(subset=["Invoice Date"])
    if not combined_dates.empty:
        combined_dates["Month"] = combined_dates["Invoice Date"].dt.to_period("M").astype(str)
        monthly = combined_dates.groupby(["Month", "Source"])["Taxable Value"].sum().reset_index()
        if monthly["Month"].nunique() > 1:
            fig4 = px.bar(
                monthly, x="Month", y="Taxable Value", color="Source", barmode="group",
                title="Month-wise Taxable Value: GST 2B vs Books",
                color_discrete_map={"GST 2B": "#2563eb", "Books": "#16a34a"},
            )
            st.plotly_chart(fig4, use_container_width=True)

    # =====================================================
    # SHEETS
    # =====================================================

    with st.expander("📂 Excel Sheets Automatically Analysed"):
        col_a, col_b = st.columns(2)
        with col_a:
            st.subheader("📥 GST 2B")
            for sheet in st.session_state.get("two_b_sheets", []):
                st.write("•", sheet)
            for sheet in st.session_state.get("two_b_skipped", []):
                st.caption(f"⏭️ Skipped (no GSTIN/Invoice column): {sheet}")
        with col_b:
            st.subheader("📚 Books")
            for sheet in st.session_state.get("books_sheets", []):
                st.write("•", sheet)
            for sheet in st.session_state.get("books_skipped", []):
                st.caption(f"⏭️ Skipped (no GSTIN/Invoice column): {sheet}")

    # =====================================================
    # EXPORT
    # =====================================================

    st.markdown('<div class="section-title">📥 Export Report</div>', unsafe_allow_html=True)

    excel_report = create_excel(
        result,
        st.session_state["quality_2b"],
        st.session_state["quality_books"],
        st.session_state["tolerance_used"],
    )

    st.download_button(
        "📊 DOWNLOAD COMPLETE EXCEL REPORT",
        data=excel_report,
        file_name="GST_Reconciliation_Report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
    st.caption(
        "Report mein Summary, Complete Reconciliation, status-wise sheets, "
        "aur Vendor Summary — sab color-coded aur formatted hain."
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
    The software automatically detects important GST columns even when the
    column names are different — and lets you fine-tune the mapping if needed.
    <br><br>
    <b>Matching Key:</b><br>
    GSTIN + Invoice Number (duplicates paired in sequence)
    <br><br>
    🔴 Missing in 2B<br>
    🟠 Missing in Books<br>
    🟣 Value Mismatch<br>
    🟢 Matched
    <br><br>
    <b>What's new in this version:</b>
    <ul>
        <li>Manual column-mapping override for non-standard exports</li>
        <li>Data quality checks — blank rows, duplicate invoices, invalid GSTIN format</li>
        <li>Smart match suggestions for likely typo/formatting gaps</li>
        <li>Vendor-wise ITC-at-risk summary and chart</li>
        <li>Search + CSV export for any category</li>
        <li>Professionally formatted, color-coded Excel report with a Summary sheet</li>
    </ul>
    </div>
    """, unsafe_allow_html=True)
