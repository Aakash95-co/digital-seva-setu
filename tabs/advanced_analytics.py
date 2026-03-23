import streamlit as st
import pandas as pd
import plotly.express as px
import numpy as np


# ─────────────────────────────────────────────────────────────────────────────
# Helper: compute OOT rate safely
# ─────────────────────────────────────────────────────────────────────────────
def _oot_rate(oot, total):
    return round((oot / total * 100), 2) if total > 0 else 0.0


# ─────────────────────────────────────────────────────────────────────────────
# Scoring engine  →  returns scored DataFrame for all offices in the district
# ─────────────────────────────────────────────────────────────────────────────
def _score_offices(df, selected_district, selected_year, selected_month, min_count):
    """
    Three factors per office (for the chosen district / year / month):

    F1 – District comparison   : OOT rate vs district average  (weight 35 %)
    F2 – State average         : OOT rate vs statewide average (weight 35 %)
    F3 – Consistency penalty   : flagged if bad for 3 / 6 / 9 consecutive
                                  months ending at the selected month        (weight 30 %)

    Services with total < min_count are excluded before aggregation.
    """

    # ── 1. filter low-volume services globally ────────────────────────────────
    svc_totals = df.groupby('Service')['Total'].sum()
    valid_services = svc_totals[svc_totals >= min_count].index
    df_valid = df[df['Service'].isin(valid_services)].copy()

    # ── 2. snapshot: selected district + month/year ───────────────────────────
    snap = df_valid[
        (df_valid['District'] == selected_district) &
        (df_valid['month_dt'].dt.year == selected_year) &
        (df_valid['month_dt'].dt.month == selected_month)
    ]
    if snap.empty:
        return pd.DataFrame()

    office_snap = (
        snap.groupby('Office')
        .agg(Total=('Total', 'sum'), Out_of_Time=('Out_of_Time', 'sum'))
        .reset_index()
    )
    office_snap['OOT_Rate'] = office_snap.apply(
        lambda r: _oot_rate(r['Out_of_Time'], r['Total']), axis=1
    )

    # ── 3. district average for that month ────────────────────────────────────
    dist_total    = office_snap['Total'].sum()
    dist_oot      = office_snap['Out_of_Time'].sum()
    district_avg  = _oot_rate(dist_oot, dist_total)

    # ── 4. state average for that month ───────────────────────────────────────
    state_snap = df_valid[
        (df_valid['month_dt'].dt.year == selected_year) &
        (df_valid['month_dt'].dt.month == selected_month)
    ]
    state_total = state_snap['Total'].sum()
    state_oot   = state_snap['Out_of_Time'].sum()
    state_avg   = _oot_rate(state_oot, state_total)

    # ── 5. consistency: rolling OOT over history ──────────────────────────────
    # Get the full monthly series for this district (all admitted months)
    dist_history = (
        df_valid[df_valid['District'] == selected_district]
        .groupby(['Office', 'month_dt'])
        .agg(Total=('Total', 'sum'), Out_of_Time=('Out_of_Time', 'sum'))
        .reset_index()
    )
    dist_history['OOT_Rate'] = dist_history.apply(
        lambda r: _oot_rate(r['Out_of_Time'], r['Total']), axis=1
    )

    # Build a monthly district-level average to define "bad" threshold
    dist_monthly_avg = (
        dist_history.groupby('month_dt')
        .apply(lambda g: _oot_rate(g['Out_of_Time'].sum(), g['Total'].sum()))
        .reset_index(name='Dist_Avg')
    )
    dist_history = dist_history.merge(dist_monthly_avg, on='month_dt')
    dist_history['is_bad'] = dist_history['OOT_Rate'] > dist_history['Dist_Avg']

    # For each office count consecutive bad months up to selected month
    cutoff = pd.Timestamp(year=selected_year, month=selected_month, day=1)
    consistency_scores = {}
    for office, grp in dist_history[dist_history['month_dt'] <= cutoff].groupby('Office'):
        grp_sorted = grp.sort_values('month_dt')
        # Count trailing consecutive "bad" months
        bad_flags = grp_sorted['is_bad'].tolist()
        streak = 0
        for flag in reversed(bad_flags):
            if flag:
                streak += 1
            else:
                break
        consistency_scores[office] = streak

    office_snap['Streak'] = office_snap['Office'].map(consistency_scores).fillna(0).astype(int)

    # ── 6. composite score ────────────────────────────────────────────────────
    # F1: how much above district avg (normalised 0-100, capped at 100)
    max_oot = office_snap['OOT_Rate'].max() or 1
    office_snap['F1_District'] = ((office_snap['OOT_Rate'] - district_avg) / max(max_oot, 1) * 100).clip(0)

    # F2: how much above state avg
    office_snap['F2_State'] = ((office_snap['OOT_Rate'] - state_avg) / max(max_oot, 1) * 100).clip(0)

    # F3: streak penalty   (3–5 → 33, 6–8 → 66, 9+ → 100)
    def streak_score(s):
        if s >= 9:
            return 100
        elif s >= 6:
            return 66
        elif s >= 3:
            return 33
        return 0

    office_snap['F3_Consistency'] = office_snap['Streak'].apply(streak_score)

    office_snap['Composite_Score'] = (
        0.35 * office_snap['F1_District'] +
        0.35 * office_snap['F2_State']    +
        0.30 * office_snap['F3_Consistency']
    ).round(2)

    office_snap['District_Avg_OOT'] = district_avg
    office_snap['State_Avg_OOT']    = state_avg

    return office_snap.sort_values('Composite_Score', ascending=False).reset_index(drop=True)


# ─────────────────────────────────────────────────────────────────────────────
# Reason builder
# ─────────────────────────────────────────────────────────────────────────────
def _build_reason(row, rank_label):
    reasons = []
    delta_dist  = row['OOT_Rate'] - row['District_Avg_OOT']
    delta_state = row['OOT_Rate'] - row['State_Avg_OOT']

    if delta_dist > 0:
        reasons.append(
            f"📍 **District comparison:** OOT rate is **{delta_dist:.1f}%** higher than the "
            f"district average ({row['District_Avg_OOT']:.1f}%)."
        )
    else:
        reasons.append(
            f"📍 **District comparison:** OOT rate is **{abs(delta_dist):.1f}%** *below* the "
            f"district average ({row['District_Avg_OOT']:.1f}%) — performing better than peers."
        )

    if delta_state > 0:
        reasons.append(
            f"🌐 **State average comparison:** OOT rate exceeds the state average "
            f"({row['State_Avg_OOT']:.1f}%) by **{delta_state:.1f}%**."
        )
    else:
        reasons.append(
            f"🌐 **State average comparison:** OOT rate is **{abs(delta_state):.1f}%** *below* "
            f"the state average ({row['State_Avg_OOT']:.1f}%)."
        )

    streak = int(row['Streak'])
    if streak >= 9:
        reasons.append(
            f"🚨 **Consistency flag:** This office has been performing **above the district "
            f"average OOT rate for {streak} consecutive months** — a serious persistent concern."
        )
    elif streak >= 6:
        reasons.append(
            f"⚠️ **Consistency flag:** Flagged for **{streak} consecutive months** of above-average "
            f"OOT — indicates a sustained performance issue."
        )
    elif streak >= 3:
        reasons.append(
            f"🔔 **Consistency flag:** Above-average OOT observed for **{streak} consecutive months** "
            f"— requires monitoring."
        )
    else:
        reasons.append(
            f"✅ **Consistency:** No prolonged OOT streak detected ({streak} month(s))."
        )

    return reasons


# ─────────────────────────────────────────────────────────────────────────────
# Main render
# ─────────────────────────────────────────────────────────────────────────────
def render(df):
    # ── Header ──────────────────────────────────────────────────────────────
    st.markdown(
        """
        <div style='background: linear-gradient(90deg, #1a3c5e 0%, #2d6a9f 100%);
                    padding: 18px 28px; border-radius: 10px; margin-bottom: 20px;'>
            <h2 style='color: white; margin: 0; font-size: 1.6rem; letter-spacing: 1px;'>
                🔍 Advanced Analytics & Performance Report
            </h2>
            <p style='color: #c8dff0; margin: 4px 0 0 0; font-size: 0.95rem;'>
                Identify worst-performing offices using district benchmarking,
                state averages, and consistency analysis
            </p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # ── Filters row ─────────────────────────────────────────────────────────
    fc1, fc2, fc3, fc4 = st.columns([3, 2, 2, 2])

    with fc1:
        districts = sorted(df['District'].dropna().unique())
        selected_district = st.selectbox("🏛️ Select District", districts, key="aa_district")

    df_dist = df[df['District'] == selected_district]

    with fc2:
        years = sorted(df_dist['month_dt'].dt.year.unique())
        selected_year = st.selectbox("📅 Select Year", years, key="aa_year")

    with fc3:
        months_avail = sorted(
            df_dist[df_dist['month_dt'].dt.year == selected_year]['month_dt'].dt.month.unique()
        )
        MONTH_NAMES = {
            1: "January", 2: "February", 3: "March", 4: "April",
            5: "May",     6: "June",     7: "July",  8: "August",
            9: "September", 10: "October", 11: "November", 12: "December"
        }
        month_options = {MONTH_NAMES[m]: m for m in months_avail}
        selected_month_name = st.selectbox(
            "🗓️ Select Month", list(month_options.keys()), key="aa_month"
        )
        selected_month = month_options[selected_month_name]

    with fc4:
        min_count = st.number_input(
            "⚙️ Min. Service Count Threshold",
            min_value=0, max_value=10000,
            value=100, step=10,
            help="Services whose total count (across all months) is below this threshold "
                 "will be excluded from analysis.",
            key="aa_min_count"
        )

    # ── Score ────────────────────────────────────────────────────────────────
    scored = _score_offices(df, selected_district, selected_year, selected_month, int(min_count))

    if scored.empty:
        st.warning("No data available for the selected filters.")
        return

    district_avg = scored['District_Avg_OOT'].iloc[0]
    state_avg    = scored['State_Avg_OOT'].iloc[0]

    # ── KPI bar ──────────────────────────────────────────────────────────────
    st.markdown("<br>", unsafe_allow_html=True)
    kb1, kb2, kb3, kb4 = st.columns(4)

    def _kpi_card(col, label, value, color="#1a3c5e", bg="#f0f6ff", border="#2d6a9f"):
        col.markdown(
            f"""
            <div style='background:{bg}; border-left:5px solid {border};
                        padding:14px 12px; border-radius:8px; text-align:center;'>
                <div style='font-size:1.5rem; font-weight:700; color:{color};'>{value}</div>
                <div style='font-size:0.78rem; color:#555; font-weight:600;
                             text-transform: uppercase; letter-spacing:0.5px;'>{label}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    _kpi_card(kb1, "Offices Analysed",   len(scored))
    _kpi_card(kb2, "District OOT Avg",   f"{district_avg:.1f}%", color="#c0392b", bg="#fff5f5", border="#e74c3c")
    _kpi_card(kb3, "State OOT Avg",      f"{state_avg:.1f}%",    color="#7d3c98", bg="#fdf5ff", border="#9b59b6")
    _kpi_card(
        kb4, "Flagged (≥3 months)",
        int((scored['Streak'] >= 3).sum()),
        color="#e67e22", bg="#fff8f0", border="#e67e22"
    )

    st.markdown("<br>", unsafe_allow_html=True)

    # ── Office-wise performance table ────────────────────────────────────────
    st.markdown(
        f"### 📊 Office-wise Performance — {selected_month_name} {selected_year}  |  {selected_district}"
    )

    display_cols = ['Office', 'Total', 'Out_of_Time', 'OOT_Rate',
                    'District_Avg_OOT', 'State_Avg_OOT', 'Streak', 'Composite_Score']
    rename_map = {
        'OOT_Rate':         'OOT Rate (%)',
        'District_Avg_OOT': 'District Avg OOT (%)',
        'State_Avg_OOT':    'State Avg OOT (%)',
        'Streak':           'Bad-Month Streak',
        'Composite_Score':  'Composite Score',
    }
    table_df = scored[display_cols].rename(columns=rename_map)

    st.dataframe(
        table_df.style
            .background_gradient(cmap='Reds', subset=['OOT Rate (%)', 'Composite Score'])
            .format({
                'OOT Rate (%)':          '{:.2f}%',
                'District Avg OOT (%)':  '{:.2f}%',
                'State Avg OOT (%)':     '{:.2f}%',
                'Composite Score':       '{:.2f}',
            }),
        use_container_width=True,
        hide_index=True,
    )

    # ── Bar chart ────────────────────────────────────────────────────────────
    fig = px.bar(
        scored.sort_values('Composite_Score', ascending=False),
        x='Office', y='OOT_Rate',
        color='Composite_Score',
        color_continuous_scale='Reds',
        title=f"Out-of-Time Rate by Office — {selected_month_name} {selected_year}",
        labels={'OOT_Rate': '% Out of Time', 'Composite_Score': 'Composite Score'},
        text='OOT_Rate',
    )
    fig.add_hline(
        y=district_avg, line_dash='dash', line_color='#2980b9',
        annotation_text=f"District Avg: {district_avg:.1f}%",
        annotation_position="top left"
    )
    fig.add_hline(
        y=state_avg, line_dash='dot', line_color='#8e44ad',
        annotation_text=f"State Avg: {state_avg:.1f}%",
        annotation_position="bottom right"
    )
    fig.update_layout(xaxis_tickangle=-45, height=480)
    fig.update_traces(texttemplate='%{text:.1f}%', textposition='outside')
    st.plotly_chart(fig, use_container_width=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # ────────────────────────────────────────────────────────────────────────
    # WORST 3 OFFICES
    # ────────────────────────────────────────────────────────────────────────
    st.markdown(
        """
        <div style='background:#fff0f0; border:1.5px solid #e74c3c;
                    border-radius:10px; padding:14px 22px; margin-bottom:12px;'>
            <h3 style='margin:0; color:#c0392b;'>⚠️ Worst Performing Offices</h3>
            <p style='margin:4px 0 0 0; color:#7f8c8d; font-size:0.9rem;'>
                Ranked by composite score: district benchmark (35%) +
                state average (35%) + consistency (30%)
            </p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    worst_n = min(3, len(scored))
    for rank, (_, row) in enumerate(scored.head(worst_n).iterrows(), 1):
        badge_color = ["#e74c3c", "#e67e22", "#f1c40f"][rank - 1]
        label       = ["🥇 Rank 1 — Worst",  "🥈 Rank 2",  "🥉 Rank 3"][rank - 1]

        with st.expander(
            f"{label}   |   {row['Office']}   |   OOT: {row['OOT_Rate']:.1f}%   "
            f"|   Score: {row['Composite_Score']:.1f}",
            expanded=(rank == 1),
        ):
            m1, m2, m3, m4 = st.columns(4)
            _kpi_card(m1, "OOT Rate",        f"{row['OOT_Rate']:.1f}%",    color="#c0392b", bg="#fff5f5", border="#e74c3c")
            _kpi_card(m2, "District Avg",    f"{row['District_Avg_OOT']:.1f}%", color="#2980b9", bg="#eaf4ff", border="#3498db")
            _kpi_card(m3, "State Avg",       f"{row['State_Avg_OOT']:.1f}%",    color="#7d3c98", bg="#fdf5ff", border="#9b59b6")
            _kpi_card(m4, "Consecutive Bad", f"{int(row['Streak'])} month(s)",
                      color="#e67e22", bg="#fff8f0", border="#e67e22")

            st.markdown("<br>", unsafe_allow_html=True)
            st.markdown("**📝 Reason for Flagging:**")
            for reason in _build_reason(row, label):
                st.markdown(f"> {reason}")

    st.markdown("<br>", unsafe_allow_html=True)

    # ────────────────────────────────────────────────────────────────────────
    # BEST 3 OFFICES
    # ────────────────────────────────────────────────────────────────────────
    st.markdown(
        """
        <div style='background:#f0fff4; border:1.5px solid #27ae60;
                    border-radius:10px; padding:14px 22px; margin-bottom:12px;'>
            <h3 style='margin:0; color:#1e8449;'>🏆 Best Performing Offices</h3>
            <p style='margin:4px 0 0 0; color:#7f8c8d; font-size:0.9rem;'>
                Offices with lowest composite score (lowest OOT burden relative to benchmarks)
            </p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    best_rows = scored.tail(min(3, len(scored))).iloc[::-1].reset_index(drop=True)
    best_labels = ["🥇 Best Office", "🥈 2nd Best", "🥉 3rd Best"]

    for rank, (_, row) in enumerate(best_rows.iterrows(), 1):
        with st.expander(
            f"{best_labels[rank-1]}   |   {row['Office']}   |   OOT: {row['OOT_Rate']:.1f}%   "
            f"|   Score: {row['Composite_Score']:.1f}",
            expanded=(rank == 1),
        ):
            m1, m2, m3, m4 = st.columns(4)
            _kpi_card(m1, "OOT Rate",        f"{row['OOT_Rate']:.1f}%",    color="#1e8449", bg="#f0fff4", border="#27ae60")
            _kpi_card(m2, "District Avg",    f"{row['District_Avg_OOT']:.1f}%", color="#2980b9", bg="#eaf4ff", border="#3498db")
            _kpi_card(m3, "State Avg",       f"{row['State_Avg_OOT']:.1f}%",    color="#7d3c98", bg="#fdf5ff", border="#9b59b6")
            _kpi_card(m4, "Consecutive Bad", f"{int(row['Streak'])} month(s)",
                      color="#27ae60", bg="#f0fff4", border="#27ae60")

            st.markdown("<br>", unsafe_allow_html=True)
            st.markdown("**📝 Performance Summary:**")
            for reason in _build_reason(row, best_labels[rank - 1]):
                st.markdown(f"> {reason}")

    # ── Disclaimer ────────────────────────────────────────────────────────────
    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown(
        f"""
        <div style='background:#f8f9fa; border:1px solid #d0d7de;
                    border-radius:8px; padding:12px 18px;
                    color:#555; font-size:0.82rem;'>
            <strong>Methodology Note:</strong>
            Services with a total count below <strong>{int(min_count)}</strong> records
            (configurable above) are excluded to avoid skewed statistics on low-volume services.
            The composite score weights: District Benchmark 35% · State Average 35% ·
            Consistency (streak) 30%.
            A "bad month" is defined as a month in which the office's OOT rate exceeds
            the district average for that month.
            Streaks of ≥ 3, ≥ 6, and ≥ 9 consecutive bad months are flagged with
            🔔, ⚠️, and 🚨 respectively.
        </div>
        """,
        unsafe_allow_html=True,
    )