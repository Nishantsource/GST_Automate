# =====================================================
# KPI DASHBOARD
# =====================================================

st.markdown(
    '<div class="section-title">📊 Reconciliation Dashboard</div>',
    unsafe_allow_html=True
)

c1, c2, c3, c4, c5 = st.columns(5)

with c1:
    st.metric(
        label="📄 TOTAL INVOICES",
        value=f"{total:,}"
    )

with c2:
    st.metric(
        label="🟢 MATCHED",
        value=f"{matched:,}"
    )

with c3:
    st.metric(
        label="🔴 MISSING IN 2B",
        value=f"{missing_2b:,}"
    )

with c4:
    st.metric(
        label="🟠 MISSING IN BOOKS",
        value=f"{missing_books:,}"
    )

with c5:
    st.metric(
        label="🟣 VALUE MISMATCH",
        value=f"{mismatch:,}"
    )
