import streamlit as st
import pandas as pd


def render(df):
    st.markdown(
        """
        <div style='background: linear-gradient(90deg, #1a3c5e 0%, #2d6a9f 100%);
                    padding: 18px 28px; border-radius: 10px; margin-bottom: 20px;'>
            <h2 style='color: white; margin: 0; font-size: 1.6rem; letter-spacing: 1px;'>
                📋 Dataset Overview & Metadata
            </h2>
            <p style='color: #c8dff0; margin: 4px 0 0 0; font-size: 0.95rem;'>
                Structural summary of the Digital Seva Setu dataset
            </p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    total_rows = len(df)
    total_cols = len(df.columns)
    total_districts = df["District"].nunique() if "District" in df.columns else "—"
    total_offices = df["Office"].nunique() if "Office" in df.columns else "—"
    total_services = df["Service"].nunique() if "Service" in df.columns else "—"
    date_range = (
        f"{df['month_dt'].min().strftime('%b %Y')}  →  {df['month_dt'].max().strftime('%b %Y')}"
        if "month_dt" in df.columns
        else "—"
    )

    # ── KPI Cards ──────────────────────────────────────────────────────────────
    st.markdown("#### 🗂️ At a Glance")
    k1, k2, k3, k4, k5, k6 = st.columns(6)
    def kpi(col, label, value, icon=""):
        col.markdown(
            f"""
            <div style='background:#f0f6ff; border-left: 5px solid #2d6a9f;
                        padding: 14px 12px; border-radius: 8px; text-align:center;'>
                <div style='font-size:1.5rem;'>{icon}</div>
                <div style='font-size:1.55rem; font-weight:700; color:#1a3c5e;'>{value}</div>
                <div style='font-size:0.78rem; color:#555; font-weight:600;
                             text-transform:uppercase; letter-spacing:0.5px;'>{label}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    kpi(k1, "Total Records",   f"{total_rows:,}",       "📝")
    kpi(k2, "Total Columns",   total_cols,               "🗃️")
    kpi(k3, "Districts",       total_districts,          "🏛️")
    kpi(k4, "Offices",         total_offices,            "🏢")
    kpi(k5, "Services",        total_services,           "⚙️")
    kpi(k6, "Date Range",      date_range,               "📅")

    st.markdown("<br>", unsafe_allow_html=True)

    # ── Column Information ─────────────────────────────────────────────────────
    st.markdown("#### 📌 Column-level Information")

    TYPE_COLOR = {
        "object":          ("#e8f0fe", "#1a56db"),
        "int64":           ("#e8f8f0", "#1e7e4a"),
        "float64":         ("#fff8e1", "#c17f00"),
        "datetime64[ns]":  ("#fce8f3", "#9c27b0"),
    }

    meta_rows = []
    for col in df.columns:
        dtype      = str(df[col].dtype)
        non_null   = int(df[col].notna().sum())
        null_count = int(df[col].isna().sum())
        unique     = int(df[col].nunique())
        completeness = f"{(non_null / total_rows * 100):.1f}%" if total_rows else "—"
        sample     = df[col].dropna().iloc[0] if non_null > 0 else "N/A"
        meta_rows.append(
            {
                "Column":        col,
                "Data Type":     dtype,
                "Non-Null Count": f"{non_null:,}",
                "Null Count":    null_count,
                "Completeness":  completeness,
                "Unique Values": f"{unique:,}",
                "Sample Value":  str(sample),
            }
        )

    meta_df = pd.DataFrame(meta_rows)
    st.dataframe(
        meta_df.style.apply(
            lambda _: [
                f"background-color: {TYPE_COLOR.get(str(df[r['Column']].dtype), ('#fff','#000'))[0]};"
                f"color: {TYPE_COLOR.get(str(df[r['Column']].dtype), ('#fff','#333'))[1]};"
                f"font-weight:600;"
                if c == "Data Type" else ""
                for c in meta_df.columns
            ],
            axis=1,
        ),
        use_container_width=True,
        hide_index=True,
    )

    st.markdown("<br>", unsafe_allow_html=True)

    # ── Summary Statistics ─────────────────────────────────────────────────────
    st.markdown("#### 📊 Summary Statistics — Numeric Columns")
    numeric_df = df.select_dtypes(include="number")
    if not numeric_df.empty:
        desc = numeric_df.describe().T.reset_index().rename(columns={"index": "Column"})
        desc = desc.round(2)
        st.dataframe(
            desc.style.background_gradient(cmap="Blues", subset=["mean", "max"]),
            use_container_width=True,
            hide_index=True,
        )
    else:
        st.info("No numeric columns found.")

    st.markdown("<br>", unsafe_allow_html=True)

    # ── Categorical Distribution ───────────────────────────────────────────────
    st.markdown("#### 🔠 Categorical Column Distribution")
    cat_cols = df.select_dtypes(include="object").columns.tolist()
    if cat_cols:
        selected_col = st.selectbox(
            "Select a categorical column to explore:", cat_cols, key="meta_cat_col"
        )
        dist = df[selected_col].value_counts().reset_index()
        dist.columns = ["Value", "Count"]
        dist["Share (%)"] = (dist["Count"] / dist["Count"].sum() * 100).round(2)

        c1, c2 = st.columns([2, 1])
        with c1:
            st.dataframe(
                dist.style.bar(subset=["Count"], color="#2d6a9f")
                         .format({"Share (%)": "{:.2f}%"}),
                use_container_width=True,
                hide_index=True,
            )
        with c2:
            st.markdown(
                f"""
                <div style='background:#f0f6ff; border-radius:10px; padding:18px;'>
                    <p style='margin:0; font-size:0.85rem; color:#555;'>Unique Values</p>
                    <p style='margin:0; font-size:2rem; font-weight:700;
                               color:#1a3c5e;'>{dist.shape[0]:,}</p>
                    <hr style='border-color:#c8dff0;'>
                    <p style='margin:0; font-size:0.85rem; color:#555;'>Top Value</p>
                    <p style='margin:0; font-size:1rem; font-weight:600;
                               color:#2d6a9f;'>{dist.iloc[0]['Value']}</p>
                    <p style='margin:4px 0 0 0; font-size:0.85rem;
                               color:#888;'>{dist.iloc[0]['Count']:,} records
                               ({dist.iloc[0]['Share (%)']:.1f}%)</p>
                </div>
                """,
                unsafe_allow_html=True,
            )
    else:
        st.info("No categorical columns found.")

    st.markdown("<br>", unsafe_allow_html=True)

    # ── Footer Note ────────────────────────────────────────────────────────────
    st.markdown(
        """
        <div style='background:#f8f9fa; border: 1px solid #d0d7de;
                    border-radius:8px; padding:12px 18px;
                    color:#555; font-size:0.82rem;'>
            <strong>Note:</strong> This metadata view is auto-generated from the loaded dataset.
            Completeness is calculated as the percentage of non-null values per column.
            Sample values are drawn from the first available non-null record.
        </div>
        """,
        unsafe_allow_html=True,
    )