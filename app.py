import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import re as _re

st.set_page_config(page_title="Dividend Quality & Earnings Analytics", layout="wide", initial_sidebar_state="collapsed")

st.markdown("""
<style>
.stApp { background-color: #0d1117; }
.eyebrow { color:#3fb950; font-size:12px; letter-spacing:2px; font-weight:700; text-transform:uppercase; }
.main-title { font-size:28px; font-weight:800; color:#fff; margin-top:2px; }
.subtitle { color:#8b949e; font-size:13px; margin-top:2px; }
.card { background:#161b22; border:1px solid #30363d; border-radius:10px; padding:16px 18px; margin-bottom:16px; }
.card-title { font-size:14px; font-weight:700; color:#e6edf3; margin-bottom:10px; }
.dot-orange::before { content:"● "; color:#d29922; }
.dot-pink::before { content:"● "; color:#f778ba; }
.dot-cyan::before { content:"● "; color:#39c5cf; }
.pill { padding:2px 10px; border-radius:12px; font-size:11px; font-weight:700; display:inline-block; }
.pill-excellent { background:rgba(63,185,80,0.15); color:#3fb950; }
.pill-good { background:rgba(63,185,80,0.12); color:#56d364; }
.pill-fair { background:rgba(210,153,34,0.15); color:#d29922; }
.pill-weak { background:rgba(219,109,40,0.15); color:#db6d28; }
.pill-poor { background:rgba(248,81,73,0.15); color:#f85149; }
.pill-buy { background:rgba(56,139,253,0.15); color:#58a6ff; }
.pill-hold { background:rgba(139,148,158,0.15); color:#8b949e; }
.pill-sell { background:rgba(248,81,73,0.15); color:#f85149; }
.rank-badge { display:inline-flex; align-items:center; justify-content:center; width:22px; height:22px;
    border-radius:50%; font-size:11px; font-weight:700; color:#000; }
.rank-gold { background:#d4a72c; }
.rank-silver { background:#6e7681; color:#fff; }
.rank-plain { color:#8b949e; padding-left:6px; }
.bar-wrap { background:#21262d; border-radius:4px; height:6px; width:50px; display:inline-block; vertical-align:middle; margin-right:6px; }
.bar-fill { height:6px; border-radius:4px; display:inline-block; }
.tbl-header { display:flex; padding:6px 4px; border-bottom:1px solid #30363d; color:#8b949e; font-size:11px;
    text-transform:uppercase; font-weight:700; }
.tbl-row-cell { display:flex; align-items:center; padding:2px 4px; font-size:13px; color:#e6edf3; min-height:34px; }
/* ทำให้ st.button ดูเหมือนตัวหนังสือ ไม่มีกรอบปุ่ม */
div[data-testid="stButton"] > button {
    background: transparent !important; border: none !important; color:#e6edf3 !important;
    font-weight:700 !important; padding: 4px 2px !important; text-align:left !important;
    width:100% !important; box-shadow:none !important;
}
div[data-testid="stButton"] > button:hover { color:#58a6ff !important; text-decoration: underline; }
div[data-testid="stButton"] > button:focus:not(:active) { color:#58a6ff !important; }
</style>
""", unsafe_allow_html=True)

EXCEL_PATH = "model_predictคะแนน68__1_.xlsx"
YEARS = ["2565", "2566", "2567", "2568"]

SHEET1_BASE = {
    "revenue": "รายได้รวม", "netprofit": "กำไรสุทธิ", "eps": "EPS (บาท)",
    "assets": "สินทรัพย์รวม", "liab": "หนี้สินรวม", "equity": "ส่วนของผู้ถือหุ้น", "price": "ราคาหุ้น (บาท)",
}
SHEET2_BASE = {"divyield": "Div Yield (%)", "cfo": "CFO"}
SHEET1_NAME, SHEET2_NAME = "ข้อมูล65-68", "Historical Trends"

FACTOR_LABELS = {
    "F1": "Dividend Safety", "F2": "Dividend Track Record", "F3": "Financial Strength",
    "F4": "Growth", "F5": "Profitability", "F6": "Risk", "F7": "Earnings Quality",
}

def _norm(s): return _re.sub(r"\s+", "", str(s))
def find_col(columns, base, year):
    target = _norm(base + year)
    for c in columns:
        if _norm(c) == target: return c
    return None
def pct_to_score(pct_series): return (pct_series * 9 + 1).round(1)

@st.cache_data
def load_data():
    df1 = pd.read_excel(EXCEL_PATH, sheet_name=SHEET1_NAME)
    df1.columns = [str(c).replace("\n", " ").strip() for c in df1.columns]
    df2 = pd.read_excel(EXCEL_PATH, sheet_name=SHEET2_NAME)
    df2.columns = [str(c).replace("\n", " ").strip() for c in df2.columns]
    sym_col1, ind_col1, sec_col1 = df1.columns[0], df1.columns[1], df1.columns[2]
    sym_col2 = df2.columns[0]
    df2_indexed = df2.set_index(sym_col2)
    records = []
    for _, row in df1.iterrows():
        symbol = row[sym_col1]
        for yr in YEARS:
            rec = {"symbol": symbol, "industry": row[ind_col1], "sector": row[sec_col1], "year": yr}
            for key, base in SHEET1_BASE.items():
                col = find_col(df1.columns, base, yr)
                rec[key] = row[col] if col else np.nan
            if symbol in df2_indexed.index:
                r2 = df2_indexed.loc[symbol]
                for key, base in SHEET2_BASE.items():
                    col = find_col(df2.columns, base, yr)
                    rec[key] = r2[col] if col else np.nan
            else:
                rec["divyield"], rec["cfo"] = np.nan, np.nan
            records.append(rec)
    return pd.DataFrame(records)

@st.cache_data
def compute_scores(long_df):
    d = long_df.sort_values(["symbol", "year"]).copy()
    d["dps"] = d["price"] * (d["divyield"] / 100)
    d["payout_ratio"] = (d["dps"] / d["eps"].replace(0, np.nan)) * 100
    d["payout_ratio_3y"] = d.groupby("symbol")["payout_ratio"].transform(lambda s: s.rolling(3, min_periods=1).mean())
    d["dps_growth"] = d.groupby("symbol")["dps"].pct_change()
    d["dps_cagr"] = d["symbol"].map(d.groupby("symbol")["dps_growth"].mean())
    d["dps_never_decreased"] = d["symbol"].map(d.groupby("symbol")["dps_growth"].apply(lambda s: (s.dropna() >= 0).all()))
    d["equity_assets_pct"] = (d["equity"] / d["assets"].replace(0, np.nan)) * 100
    d["revenue_growth"] = d.groupby("symbol")["revenue"].pct_change()
    d["eps_growth"] = d.groupby("symbol")["eps"].pct_change()
    d["growth_avg"] = d[["revenue_growth", "eps_growth"]].mean(axis=1)
    d["roe_pct"] = (d["netprofit"] / d["equity"].replace(0, np.nan)) * 100
    d["de_pct"] = (d["liab"] / d["equity"].replace(0, np.nan)) * 100
    d["F7"] = (0.35 * (d["cfo"] / d["revenue"].replace(0, np.nan)) * 100) / 10
    d["QOE"] = d["cfo"] / d["netprofit"].replace(0, np.nan)
    d["F1"] = pct_to_score(d.groupby("year")["payout_ratio_3y"].rank(pct=True))
    d["F2"] = pct_to_score(d.groupby("year")["dps_cagr"].rank(pct=True))
    d.loc[d["dps_never_decreased"] == True, "F2"] = (d["F2"] + 0.5).clip(upper=10)
    d["F3"] = pct_to_score(d.groupby("year")["equity_assets_pct"].rank(pct=True))
    d["F4"] = pct_to_score(d.groupby("year")["growth_avg"].rank(pct=True))
    d["F5"] = pct_to_score(d.groupby("year")["roe_pct"].rank(pct=True))
    d["F6"] = pct_to_score(1 - d.groupby("year")["de_pct"].rank(pct=True))
    d["Scorecard"] = d[["F1","F2","F3","F4","F5","F6"]].sum(axis=1)
    d["TotalScore"] = d["Scorecard"] + d["F7"]
    return d

def assign_label(score, year_scores):
    if pd.isna(score): return "N/A", "N/A"
    valid = year_scores.dropna()
    if len(valid) == 0: return "N/A", "N/A"
    pct = (valid < score).mean() * 100
    bands = [(85,"Excellent","Strong Buy"),(65,"Good","Buy"),(40,"Fair","Hold"),(15,"Weak","Reduce"),(0,"Poor","Sell")]
    for th, lab, rec in bands:
        if pct >= th: return lab, rec
    return "Poor", "Sell"

PILL_MAP = {"Excellent":"pill-excellent","Good":"pill-good","Fair":"pill-fair","Weak":"pill-weak","Poor":"pill-poor",
            "Strong Buy":"pill-buy","Buy":"pill-buy","Hold":"pill-hold","Reduce":"pill-hold","Sell":"pill-sell","N/A":"pill-hold"}

long_df = load_data()
scored_df = compute_scores(long_df)
n_stocks = scored_df["symbol"].nunique()

h1, h2, h3 = st.columns([3, 1, 1])
with h1:
    st.markdown(f'<div class="eyebrow">STOCK SCREENING · SET {n_stocks} UNIVERSE</div>'
                f'<div class="main-title">Dividend Quality & Earnings Analytics</div>'
                f'<div class="subtitle">Evaluating Earnings Quality and Investment Potential</div>', unsafe_allow_html=True)
with h2:
    st.markdown('<div style="color:#8b949e;font-size:11px;">INDUSTRY</div>', unsafe_allow_html=True)
    industries = ["All"] + sorted(scored_df["industry"].dropna().unique().tolist())
    sel_industry = st.selectbox("industry", industries, label_visibility="collapsed")
with h3:
    st.markdown('<div style="color:#8b949e;font-size:11px;">YEAR</div>', unsafe_allow_html=True)
    sel_year = st.selectbox("year", ["All"] + YEARS, label_visibility="collapsed")

st.write("")

view_year = "2568" if sel_year == "All" else sel_year
view_df = scored_df[scored_df["year"] == view_year].copy()
if sel_industry != "All":
    view_df = view_df[view_df["industry"] == sel_industry]

year_scores = scored_df[scored_df["year"] == view_year]["TotalScore"]
view_df[["QualityLabel","Recommend"]] = view_df["TotalScore"].apply(lambda s: pd.Series(assign_label(s, year_scores)))
view_df = view_df.sort_values("TotalScore", ascending=False, na_position="last").reset_index(drop=True)
view_df.insert(0, "No.", range(1, len(view_df)+1))

if "sel_symbol" not in st.session_state:
    st.session_state.sel_symbol = None
if st.session_state.sel_symbol not in view_df["symbol"].tolist():
    st.session_state.sel_symbol = view_df.iloc[0]["symbol"] if len(view_df) else None

colL, colR = st.columns([2, 1])

with colL:
    st.markdown(f'<div class="card"><div class="card-title dot-orange">Stock Overview'
                f'<span style="float:right;color:#8b949e;font-weight:400;">{len(view_df)} stocks</span></div>',
                unsafe_allow_html=True)

    COLW = [0.4, 0.9, 1.5, 0.9, 0.9, 0.9, 0.9, 0.9, 0.9]
    hdr = st.columns(COLW)
    for c, label in zip(hdr, ["No.","Symbol","Industry",f"Div Yield ({view_year})","Scorecard","Earning Q.","Total Score","Quality","Rec."]):
        c.markdown(f'<div class="tbl-row-cell" style="color:#8b949e;font-size:11px;text-transform:uppercase;font-weight:700;">{label}</div>', unsafe_allow_html=True)

    max_eq = view_df["F7"].max(skipna=True) if len(view_df) else 1
    max_eq = max_eq if max_eq and max_eq > 0 else 1

    with st.container(height=480):
        for _, r in view_df.iterrows():
            rank = int(r["No."])
            badge = f'<span class="rank-badge rank-gold">{rank}</span>' if rank <= 5 else \
                    (f'<span class="rank-badge rank-silver">{rank}</span>' if rank <= 10 else f'<span class="rank-plain">{rank}</span>')
            sc_pct = min(100, (r["Scorecard"]/60)*100) if pd.notna(r["Scorecard"]) else 0
            eq_val = r["F7"] if pd.notna(r["F7"]) else 0
            eq_pct = min(100, (eq_val/max_eq)*100)
            tot_disp = "N/A" if pd.isna(r["TotalScore"]) else f"{r['TotalScore']:.0f}/100"
            dy_disp = "N/A" if pd.isna(r["divyield"]) else f"{r['divyield']:.2f}%"

            row_cols = st.columns(COLW)
            row_cols[0].markdown(f'<div class="tbl-row-cell">{badge}</div>', unsafe_allow_html=True)
            with row_cols[1]:
                if st.button(r["symbol"], key=f"btn_{r['symbol']}_{view_year}_{sel_industry}", use_container_width=True):
                    st.session_state.sel_symbol = r["symbol"]
                    st.rerun()
            row_cols[2].markdown(f'<div class="tbl-row-cell" style="color:#8b949e;">{r["industry"]}</div>', unsafe_allow_html=True)
            row_cols[3].markdown(f'<div class="tbl-row-cell">{dy_disp}</div>', unsafe_allow_html=True)
            row_cols[4].markdown(f'<div class="tbl-row-cell"><span class="bar-wrap"><span class="bar-fill" style="width:{sc_pct}%;background:#d29922;"></span></span>{r["Scorecard"]:.0f}</div>' if pd.notna(r["Scorecard"]) else '<div class="tbl-row-cell">N/A</div>', unsafe_allow_html=True)
            row_cols[5].markdown(f'<div class="tbl-row-cell"><span class="bar-wrap"><span class="bar-fill" style="width:{eq_pct}%;background:#39c5cf;"></span></span>{r["F7"]:.2f}</div>' if pd.notna(r["F7"]) else '<div class="tbl-row-cell">N/A</div>', unsafe_allow_html=True)
            row_cols[6].markdown(f'<div class="tbl-row-cell"><b>{tot_disp}</b></div>', unsafe_allow_html=True)
            row_cols[7].markdown(f'<div class="tbl-row-cell"><span class="pill {PILL_MAP.get(r["QualityLabel"],"pill-fair")}">{r["QualityLabel"]}</span></div>', unsafe_allow_html=True)
            row_cols[8].markdown(f'<div class="tbl-row-cell"><span class="pill {PILL_MAP.get(r["Recommend"],"pill-hold")}">{r["Recommend"]}</span></div>', unsafe_allow_html=True)

    st.markdown('</div>', unsafe_allow_html=True)

selected_symbol = st.session_state.sel_symbol

with colR:
    if selected_symbol:
        srow = view_df[view_df["symbol"] == selected_symbol].iloc[0]
        tot_disp = "N/A" if pd.isna(srow['TotalScore']) else f"{srow['TotalScore']:.0f}/100"
        st.markdown(f"""
        <div class="card">
            <div class="card-title dot-orange">Selected Stock</div>
            <div style="font-size:22px;font-weight:800;color:#fff;">{srow['symbol']}</div>
            <div style="color:#8b949e;font-size:12px;margin-bottom:10px;">{srow['industry']}</div>
            <div style="display:flex;justify-content:space-between;margin-bottom:6px;">
                <span style="color:#8b949e;font-size:12px;">Rank</span>
                <span style="font-weight:700;">{int(srow['No.'])}/{n_stocks}</span>
            </div>
            <div style="display:flex;justify-content:space-between;margin-bottom:6px;">
                <span style="color:#8b949e;font-size:12px;">Total Score</span>
                <span style="font-weight:700;">{tot_disp}</span>
            </div>
            <div style="display:flex;justify-content:space-between;margin-bottom:6px;">
                <span style="color:#8b949e;font-size:12px;">Quality Label</span>
                <span class="pill {PILL_MAP.get(srow['QualityLabel'],'pill-fair')}">{srow['QualityLabel']}</span>
            </div>
            <div style="display:flex;justify-content:space-between;">
                <span style="color:#8b949e;font-size:12px;">Investment Rec.</span>
                <span class="pill {PILL_MAP.get(srow['Recommend'],'pill-hold')}">{srow['Recommend']}</span>
            </div>
        </div>""", unsafe_allow_html=True)

        st.markdown('<div class="card"><div class="card-title dot-pink">Score Breakdown (F1–F7)</div>', unsafe_allow_html=True)
        factors = ["F1","F2","F3","F4","F5","F6","F7"]
        factor_disp = [f"{f} - {FACTOR_LABELS[f]}" for f in factors]
        values = [srow[f] for f in factors]
        fig = go.Figure(data=go.Scatterpolar(r=values+[values[0]], theta=factor_disp+[factor_disp[0]], fill="toself",
                                              line_color="#39c5cf", fillcolor="rgba(57,197,207,0.25)"))
        fig.update_layout(polar=dict(bgcolor="#161b22", radialaxis=dict(visible=True, range=[0,10], color="#8b949e", gridcolor="#30363d"),
                       angularaxis=dict(color="#e6edf3", gridcolor="#30363d")),
            showlegend=False, height=340, margin=dict(l=40,r=40,t=20,b=20),
            paper_bgcolor="rgba(0,0,0,0)", font_color="#e6edf3", font_size=10)
        st.plotly_chart(fig, use_container_width=True)
        st.markdown('</div>', unsafe_allow_html=True)

if selected_symbol:
    srow = view_df[view_df["symbol"] == selected_symbol].iloc[0]
    ind = srow["industry"]
    trend_df = scored_df[scored_df["industry"] == ind].groupby("year").agg(
        QOE=("QOE","mean"), PayoutRatio=("payout_ratio","mean"),
        DPS=("dps","mean"), DivYield=("divyield","mean")
    ).reindex(YEARS)

    st.markdown(f'<div class="card"><div class="card-title dot-cyan">Historical Trends'
                f'<span style="float:right;color:#8b949e;font-weight:400;">อุตสาหกรรม: {ind} · 2565–2568</span></div>',
                unsafe_allow_html=True)

    tc1, tc2 = st.columns(2)
    with tc1:
        st.markdown('<div style="font-size:12px;color:#8b949e;margin-bottom:6px;">คุณภาพกำไร & ความยั่งยืน (QOE & Payout Ratio)</div>', unsafe_allow_html=True)
        fig1 = go.Figure()
        fig1.add_trace(go.Scatter(x=YEARS, y=trend_df["QOE"], name="QOE (CFO/NetProfit)", mode="lines+markers", line_color="#39c5cf"))
        fig1.add_trace(go.Scatter(x=YEARS, y=trend_df["PayoutRatio"], name="Payout Ratio (%)", mode="lines+markers",
                                   line_color="#f778ba", yaxis="y2"))
        fig1.update_layout(
            yaxis=dict(title="QOE", color="#8b949e", gridcolor="#21262d"),
            yaxis2=dict(title="Payout Ratio (%)", overlaying="y", side="right", color="#8b949e"),
           xaxis=dict(color="#8b949e", gridcolor="#21262d", type="category"),
            height=280, margin=dict(l=20,r=20,t=10,b=20), legend=dict(orientation="h", font_color="#e6edf3", font_size=10),
            plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)", font_color="#e6edf3")
        st.plotly_chart(fig1, use_container_width=True)

    with tc2:
        st.markdown('<div style="font-size:12px;color:#8b949e;margin-bottom:6px;">ผลตอบแทนปันผล (DPS & Dividend Yield)</div>', unsafe_allow_html=True)
        fig2 = go.Figure()
        fig2.add_trace(go.Scatter(x=YEARS, y=trend_df["DPS"], name="DPS (บาท/หุ้น)", mode="lines+markers", line_color="#d29922"))
        fig2.add_trace(go.Scatter(x=YEARS, y=trend_df["DivYield"], name="Dividend Yield (%)", mode="lines+markers",
                                   line_color="#3fb950", yaxis="y2"))
        fig2.update_layout(
            yaxis=dict(title="DPS (บาท)", color="#8b949e", gridcolor="#21262d"),
            yaxis2=dict(title="Dividend Yield (%)", overlaying="y", side="right", color="#8b949e"),
            xaxis=dict(color="#8b949e", gridcolor="#21262d"),
            height=280, margin=dict(l=20,r=20,t=10,b=20), legend=dict(orientation="h", font_color="#e6edf3", font_size=10),
            plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)", font_color="#e6edf3")
        st.plotly_chart(fig2, use_container_width=True)

    st.markdown('</div>', unsafe_allow_html=True)
