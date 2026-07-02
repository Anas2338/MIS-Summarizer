"""Streamlit web frontend for MIS Data Converter.

Supports two modes:
  - MIS Summarizer — tax payment data grouped by financial year & section
  - Sale/Purchase Summarizer — invoice data grouped by buyer with totals
"""

from __future__ import annotations

import io

import pandas as pd
import streamlit as st
from pandas import DataFrame

from mis_converter.converter import (
    add_financial_year,
    clean_data,
    read_input,
    sort_data,
    write_output,
)
from mis_converter.sale_purchase import (
    KEEP_COLS as SP_KEEP_COLS,
    convert as sp_convert,
    read_input as sp_read_input,
)

# ── Page configuration ─────────────────────────────────────────────────
st.set_page_config(
    page_title="MIS Summarizer",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Session-state keys ─────────────────────────────────────────────────
if "mode" not in st.session_state:
    st.session_state.mode = "mis"

if "mis_clean_df" not in st.session_state:
    st.session_state.mis_clean_df = None
if "mis_buffer" not in st.session_state:
    st.session_state.mis_buffer = None
if "mis_uploaded_name" not in st.session_state:
    st.session_state.mis_uploaded_name = None

if "sp_buffer" not in st.session_state:
    st.session_state.sp_buffer = None
if "sp_uploaded_name" not in st.session_state:
    st.session_state.sp_uploaded_name = None

if "sp_sort_by" not in st.session_state:
    st.session_state.sp_sort_by = "Buyer Name"


# ── Helper: format file size for display ───────────────────────────────
def _safe_display_df(df: pd.DataFrame) -> pd.DataFrame:
    """Return a copy safe for PyArrow rendering in ``st.dataframe``.
    Converts all columns to string when they contain mixed types
    (e.g. int + NaN + str) which Arrow cannot serialize.
    """
    df = df.copy()
    for col in df.columns:
        unique_types = df[col].dropna().apply(type).unique()
        if len(unique_types) > 1:
            df[col] = df[col].astype(str)
            continue
        dtype = df[col].dtype
        if dtype in ("object", "str"):
            try:
                sample = df[col].dropna().head(50)
                has_big = sample.apply(
                    lambda v: isinstance(v, int) and abs(v) > 2**53
                )
                if has_big.any():
                    df[col] = df[col].astype(str)
            except Exception:
                pass
        if dtype in ("int64", "Int64"):
            try:
                if df[col].max() > 2**53 or df[col].min() < -(2**53):
                    df[col] = df[col].astype(str)
            except Exception:
                pass
    return df


def _fmt_size(bytes_: int) -> str:
    if bytes_ < 1024:
        return f"{bytes_} B"
    elif bytes_ < 1024 * 1024:
        return f"{bytes_ / 1024:.1f} KB"
    else:
        return f"{bytes_ / (1024 * 1024):.1f} MB"


# ── Sidebar ────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## 🛠️ Tools")

    # ── Mode selector ──────────────────────────────────────────────────
    prev_mode = st.session_state.mode

    mode = st.segmented_control(
        "Mode",
        options=["mis", "sp"],
        format_func=lambda m: (
            "📊 MIS" if m == "mis" else "🧾 Sale/Purchase"
        ),
        key="mode",
        selection_mode="single",
        label_visibility="collapsed",
    )
    if mode is None:
        st.session_state.mode = prev_mode

    if st.session_state.mode != prev_mode:
        if st.session_state.mode == "mis":
            st.session_state.sp_buffer = None
            st.session_state.sp_uploaded_name = None
        else:
            st.session_state.mis_clean_df = None
            st.session_state.mis_buffer = None
            st.session_state.mis_uploaded_name = None

    st.divider()

    # ── File uploader ─────────────────────────────────────────────────
    uploaded_file = st.file_uploader(
        "Upload Excel File",
        type=["xls", "xlsx"],
        help="Upload a raw MIS Excel export (.xls or .xlsx)",
        key="uploader",
    )

    if uploaded_file is not None:
        ext = uploaded_file.name.rsplit(".", 1)[-1].upper()
        st.caption(f"📄 {uploaded_file.name}  ·  {_fmt_size(uploaded_file.size)}  ·  `.{ext}`")


# ═══════════════════════════════════════════════════════════════════════
# MIS SUMMARIZER MODE
# ═══════════════════════════════════════════════════════════════════════

MIS_EXPECTED_COLS = [
    "Wa Name", "Section", "Tax Year",
    "Taxable Amount", "Paid Amount", "Payment Date",
]


def _render_mis_ui(raw_df: pd.DataFrame) -> None:
    """MIS Summarizer UI."""
    if st.session_state.mis_uploaded_name != uploaded_file.name:
        st.session_state.mis_clean_df = None
        st.session_state.mis_buffer = None
        st.session_state.mis_uploaded_name = uploaded_file.name

    missing = [c for c in MIS_EXPECTED_COLS if c not in raw_df.columns]
    if missing:
        st.warning(f"Missing columns: **{', '.join(missing)}**. Will skip them.")

    if st.session_state.mis_clean_df is not None:
        _render_mis_results()
        return

    # ---- Preview + Convert ----
    _, btn_col, _ = st.columns([6, 2, 1])
    with btn_col:
        convert_clicked = st.button("▶ Convert", type="primary", key="mis_convert", width="stretch")

    st.dataframe(
        _safe_display_df(raw_df),
        height=280,
        width="stretch",
    )

    if convert_clicked:
        with st.spinner("Converting..."):
            _run_mis_conversion(raw_df)
        st.rerun()


def _run_mis_conversion(raw_df: DataFrame) -> None:
    """Run the MIS conversion pipeline."""
    try:
        with st.status("Converting…", expanded=False) as status:
            status.write("Cleaning data…")
            clean_df = clean_data(raw_df)
            if clean_df.empty or len(clean_df.columns) == 0:
                st.error("None of the expected columns were found. Conversion cannot proceed.")
                return

            status.write("Adding financial year…")
            fy_df = add_financial_year(clean_df)

            status.write("Sorting data…")
            sorted_df = sort_data(fy_df)

            if sorted_df["Payment Date"].isna().sum() > 0:
                unparseable = sorted_df["Payment Date"].isna().sum()
                st.warning(
                    f"{unparseable} row(s) have unparseable payment "
                    f"date{'s' if unparseable > 1 else ''}."
                )

            status.write("Generating Excel file…")
            buf = io.BytesIO()
            write_output(sorted_df, buf)
            buf.seek(0)

            st.session_state.mis_clean_df = sorted_df
            st.session_state.mis_buffer = buf
            st.session_state.mis_uploaded_name = uploaded_file.name

            status.write("✅ Conversion complete!")
    except Exception as exc:
        st.error("An error occurred during conversion.")
        st.exception(exc)


def _render_mis_results() -> None:
    """Render the MIS conversion results."""
    sorted_df = st.session_state.mis_clean_df

    fy_count = sorted_df["Financial Year"].nunique()
    section_count = sorted_df["Section"].nunique()

    m1, m2, m3, _, dl = st.columns([1, 1, 1, 2, 2])
    m1.metric("Total Rows", len(sorted_df))
    m2.metric("Financial Years", fy_count)
    m3.metric("Sections", section_count)
    with dl:
        st.download_button(
            label="📥 Download Excel",
            data=st.session_state.mis_buffer,
            file_name="mis_formatted_output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            width="stretch",
        )

    processed_preview = sorted_df.drop(columns=["_fy_sort"], errors="ignore")
    st.dataframe(
        _safe_display_df(processed_preview),
        height=280,
        width="stretch",
    )


# ═══════════════════════════════════════════════════════════════════════
# SALE / PURCHASE SUMMARIZER MODE
# ═══════════════════════════════════════════════════════════════════════


def _on_sp_sort_change() -> None:
    """Clear cached conversion results when sort-by option changes."""
    st.session_state.sp_buffer = None
    st.session_state.sp_uploaded_name = None


def _render_sp_ui(raw_df: pd.DataFrame) -> None:
    """Sale/Purchase Summarizer UI."""
    if st.session_state.sp_uploaded_name != uploaded_file.name:
        st.session_state.sp_buffer = None
        st.session_state.sp_uploaded_name = uploaded_file.name

    missing = [c for c in SP_KEEP_COLS if c not in raw_df.columns]
    if missing:
        st.warning(f"Missing columns: **{', '.join(missing)}**. Will skip them.")

    # ── Sort-by selector ──────────────────────────────────────────────────
    sort_col, _ = st.columns([4, 6])
    with sort_col:
        st.session_state.sp_sort_by = st.radio(
            "Sort by",
            options=["Buyer Name", "Seller Name", "Tax Year"],
            key="sp_sort_radio",
            horizontal=True,
            label_visibility="visible",
            on_change=_on_sp_sort_change,
        )

    if st.session_state.sp_buffer is not None:
        _render_sp_results()
        return

    # ---- Preview + Convert ----
    _, btn_col, _ = st.columns([6, 2, 1])
    with btn_col:
        convert_clicked = st.button("▶ Convert", type="primary", key="sp_convert", width="stretch")

    st.dataframe(
        _safe_display_df(raw_df),
        height=280,
        width="stretch",
    )

    if convert_clicked:
        with st.spinner("Converting..."):
            _run_sp_conversion(raw_df)
        st.rerun()


def _run_sp_conversion(raw_df: DataFrame) -> None:
    """Run the Sale/Purchase conversion."""
    try:
        with st.status("Converting…", expanded=False) as status:
            status.write("Cleaning & grouping invoices…")

            import tempfile
            import os

            sort_by = st.session_state.sp_sort_by

            suffix = ".xls" if uploaded_file.name.endswith(".xls") else ".xlsx"
            with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
                tmp.write(uploaded_file.getbuffer())
                tmp_path = tmp.name

            status.write(f"Generating Excel file sorted by {sort_by}…")
            buf = io.BytesIO()
            sp_convert(tmp_path, buf, sort_by=sort_by)
            buf.seek(0)
            os.unlink(tmp_path)

            st.session_state.sp_buffer = buf
            st.session_state.sp_uploaded_name = uploaded_file.name

            status.write("✅ Conversion complete!")
    except Exception as exc:
        st.error("An error occurred during conversion.")
        st.exception(exc)


def _render_sp_results() -> None:
    """Render the Sale/Purchase conversion results."""
    buf = st.session_state.sp_buffer
    buf.seek(0)
    preview_df = pd.read_excel(buf, header=None)
    buf.seek(0)

    # Parse counts from the output
    grand_total_row = None
    buyer_count = 0
    for i in range(len(preview_df)):
        qty_val = str(preview_df.iloc[i, 5]).strip()
        if qty_val == "TOTAL":
            buyer_count += 1
        elif qty_val == "GRAND TOTAL":
            grand_total_row = i

    # Read raw data for actual invoice/buyer counts
    raw_df = sp_read_input(uploaded_file)
    raw_df.columns = raw_df.columns.str.strip()
    existing = [c for c in SP_KEEP_COLS if c in raw_df.columns]
    raw_df = raw_df[existing]

    unique_buyers = raw_df["Buyer Name"].nunique() if "Buyer Name" in raw_df.columns else 0
    invoice_count = len(raw_df)

    grand_val = ""
    if grand_total_row is not None:
        try:
            grand_val = f"{int(preview_df.iloc[grand_total_row, 7]):,}"
        except Exception:
            pass

    m1, m2, m3, _, dl = st.columns([1, 1, 1, 2, 2])
    m1.metric("Total Invoices", invoice_count)
    m2.metric("Buyers", unique_buyers)
    m3.metric("Grand Total", grand_val)
    with dl:
        st.download_button(
            label="📥 Download Excel",
            data=st.session_state.sp_buffer,
            file_name="sale_purchase_output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            width="stretch",
        )

    st.dataframe(
        _safe_display_df(preview_df),
        height=280,
        width="stretch",
    )


# ═══════════════════════════════════════════════════════════════════════
# MAIN
# ═══════════════════════════════════════════════════════════════════════

def main() -> None:
    """Route to the correct UI based on selected mode."""
    # --- State A: No file uploaded -----------------------------------------
    if uploaded_file is None:
        _clear_state_for_current_mode()

        mode_label = (
            "MIS Summarizer" if st.session_state.mode == "mis"
            else "Sale/Purchase Invoice Summarizer"
        )
        icon = "📊" if st.session_state.mode == "mis" else "🧾"

        st.markdown(
            f"<h1 style='text-align: center; margin-bottom: 0;'>{icon} {mode_label}</h1>",
            unsafe_allow_html=True,
        )
        st.markdown(
            "<p style='text-align: center; margin-top: 0;'>"
            "Upload a file from the sidebar to get started.</p>",
            unsafe_allow_html=True,
        )
        c1, c2, c3 = st.columns(3)
        c1.info("📤 **Upload**\n\n.xls or .xlsx file")
        c2.info("🔄 **Convert**\n\nAuto-formatted report")
        c3.info("📥 **Download**\n\nStyled .xlsx file")
        return

    # --- Read raw file ----------------------------------------------------
    try:
        if st.session_state.mode == "mis":
            raw_df = read_input(uploaded_file)
        else:
            raw_df = sp_read_input(uploaded_file)
    except Exception as exc:
        st.error("Failed to read the file. Please ensure it is a valid MIS Excel export.")
        st.exception(exc)
        return

    if raw_df.empty:
        st.error("The file appears to be empty. Please check the source file.")
        return

    # --- Route to mode-specific UI ----------------------------------------
    if st.session_state.mode == "mis":
        _render_mis_ui(raw_df)
    else:
        _render_sp_ui(raw_df)


def _clear_state_for_current_mode() -> None:
    """Clear stale session state for the active mode when file is removed."""
    if st.session_state.mode == "mis":
        st.session_state.mis_clean_df = None
        st.session_state.mis_buffer = None
        st.session_state.mis_uploaded_name = None
    else:
        st.session_state.sp_buffer = None
        st.session_state.sp_uploaded_name = None


if __name__ == "__main__":
    main()
