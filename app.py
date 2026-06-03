"""
app.py - Equity Research Dashboard (Einzeltitel)
================================================
Ein Ticker rein -> Kennzahlen, Zeitreihen-Charts, Bewertung (DCF, Reverse-DCF,
optional Peer-Median), Scorecard und Verdikt. Excel-Export.
 
    streamlit run app.py
"""
import numpy as np
import pandas as pd
import streamlit as st
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import equity_engine as eng
 
st.set_page_config(page_title="Equity Research Dashboard", layout="wide")
st.markdown(
    "<style>#MainMenu{visibility:hidden}footer{visibility:hidden}"
    "div[data-testid='stMetricValue']{font-size:1.4rem}</style>",
    unsafe_allow_html=True)
 
# ---- Farbpalette ----
NAVY, TEAL, GREEN, RED, AMBER, GREY = "#1F3864", "#2E86AB", "#1E7D32", "#B71C1C", "#B8860B", "#8A8A8A"
PLOT_FONT = dict(family="Arial, Helvetica, sans-serif", size=12, color="#222")
 
def _layout(fig, h=320, title=None):
    fig.update_layout(template="plotly_white", height=h, font=PLOT_FONT,
                      margin=dict(l=10, r=10, t=40 if title else 20, b=10),
                      title=dict(text=title, font=dict(size=14, color=NAVY)) if title else None,
                      legend=dict(orientation="h", yanchor="bottom", y=1.02, x=0))
    return fig
 
# ---- Helfer ----
def f_pct(x, d=1): return "n/v" if x is None or (isinstance(x, float) and np.isnan(x)) else f"{x*100:.{d}f}%"
def f_x(x):        return "n/v" if x is None or (isinstance(x, float) and np.isnan(x)) else f"{x:.1f}x"
def f_eur(x, d=2): return "n/v" if x is None or (isinstance(x, float) and np.isnan(x)) else f"{x:,.{d}f}"
def row(r, name):  return r["ratios"].loc[name].values.astype(float) if name in r["ratios"].index else None
def frow(r, name): return r["financials"].loc[name].values.astype(float) if name in r["financials"].index else None
def verdict_color(v): return GREEN if "KAUFEN" in v else (RED if "MEIDEN" in v else AMBER)
 
# ================================================================ Sidebar
with st.sidebar:
    st.markdown("### Eingaben")
    ticker = st.text_input("Ticker", "DG.PA",
                           help="EU-Suffixe: .PA .DE .MC .MI .AS .SW .L .VI")
    st.markdown("**Bewertungsannahmen**")
    wacc = st.slider("WACC", 0.04, 0.15, 0.08, 0.005, format="%.3f")
    auto_g = st.checkbox("FCF-Wachstum automatisch (aus Historie, max. 10%)", True)
    growth = None if auto_g else st.slider("FCF-Wachstum explizit", 0.0, 0.15, 0.06, 0.005, format="%.3f")
    terminal = st.slider("Terminales Wachstum", 0.0, 0.04, 0.025, 0.005, format="%.3f")
    st.markdown("**Qualitatives Scoring (1-5)**")
    qb = st.slider("Geschaeftsmodell / Burggraben", 1, 5, 3)
    qm = st.slider("Management / Kapitalallokation", 1, 5, 3)
    peers = st.text_input("Peers fuer relative Bewertung (optional)", "",
                          help="Komma-getrennt; nur fuer die Football-Field-Grafik")
    run = st.button("Analyse starten", type="primary", use_container_width=True)
 
st.title("Equity Research Dashboard")
 
if not run:
    st.info("Ticker links eingeben und Analyse starten.")
    st.stop()
 
# ================================================================ Daten laden
peer_list = [p.strip() for p in peers.replace(";", ",").split(",") if p.strip()]
all_tk = [ticker.strip()] + peer_list
with st.spinner("Lade Daten von Yahoo Finance ..."):
    results, errors = eng.analyse(all_tk, wacc=wacc, growth=growth, terminal=terminal,
                                  qual_business=qb, qual_management=qm, verbose=False)
 
main = next((r for r in results if r["ticker"].upper() == ticker.strip().upper()), None)
if main is None:
    st.error(f"Keine Daten fuer '{ticker}'. Ticker/Suffix pruefen oder erneut versuchen "
             f"(yfinance/Yahoo ist zeitweise instabil).")
    for tk, msg in errors:
        st.caption(f"{tk}: {msg}")
    st.stop()
peer_rs = [r for r in results if r is not main]
 
h = main["headline"]; m = main["multiples"]; d = main["dcf"]; yrs = [str(y) for y in main["years"]]
 
# Peer-Median (ohne Selbst) fuer Football Field
def med(vals):
    vals = [v for v in vals if v is not None and not np.isnan(v)]
    return float(np.median(vals)) if vals else np.nan
med_evebit = med([r["multiples"]["ev_ebit"] for r in peer_rs])
med_pe = med([r["multiples"]["pe"] for r in peer_rs])
nd_abs = frow(main, "Nettoverschuldung")[-1] * eng.MM
impl_evebit = (med_evebit * main["latest"]["ebit"] - nd_abs) / main["latest"]["shares"] \
    if peer_rs and not np.isnan(med_evebit) else np.nan
impl_pe = med_pe * main["latest"]["eps"] if peer_rs and not np.isnan(med_pe) else np.nan
 
# ================================================================ Kopf
c1, c2 = st.columns([3, 2])
with c1:
    st.markdown(f"#### {main['name']}  ·  {main['ticker']}  ·  {main['currency']}")
    st.markdown(
        f"<div style='display:inline-block;padding:6px 16px;border-radius:4px;"
        f"background:{verdict_color(main['verdict'])};color:#fff;font-weight:700;font-size:1.05rem'>"
        f"{main['verdict']}</div>"
        f"<span style='margin-left:14px;color:{GREY}'>Conviction {main['conviction']:.1f} / 5,0</span>",
        unsafe_allow_html=True)
with c2:
    v1, v2, v3 = st.columns(3)
    v1.metric("Kurs", f_eur(main["latest"]["price"]))
    v2.metric("Fair Value (DCF)", f_eur(h["fair_value"]))
    v3.metric("Sicherheitsmarge", f_pct(h["mos"], 0))
 
st.divider()
 
# ================================================================ KPI-Leiste
nm_ = row(main, "Nettomarge"); roe_ = row(main, "ROE"); zd_ = row(main, "Zinsdeckung")
k = st.columns(6)
k[0].metric("ROIC", f_pct(h["roic"]), f"{f_pct(h['roic_wacc'])} vs WACC")
k[1].metric("EBIT-Marge", f_pct(h["ebit_margin"]))
k[2].metric("Nettomarge", f_pct(nm_[-1] if nm_ is not None else np.nan))
k[3].metric("FCF-Marge", f_pct(h["fcf_margin"]))
k[4].metric("Net Debt / EBITDA", f_x(h["nd_ebitda"]))
k[5].metric("Cash Conversion", f_pct(h["cash_conv"], 0))
k2 = st.columns(6)
k2[0].metric("Umsatz-CAGR", f_pct(h["rev_cagr"]))
k2[1].metric("ROE", f_pct(roe_[-1] if roe_ is not None else np.nan))
k2[2].metric("Zinsdeckung", f_x(zd_[-1] if zd_ is not None else np.nan))
k2[3].metric("KGV", f_x(m["pe"]))
k2[4].metric("EV / EBIT", f_x(m["ev_ebit"]))
k2[5].metric("FCF-Rendite", f_pct(m["fcf_yield"]))
 
# Excel-Download
tmp = "Equity_Analyse.xlsx"; eng.to_excel(results, tmp)
with open(tmp, "rb") as fh:
    st.download_button("Excel-Bericht herunterladen", fh.read(), file_name=f"Equity_{main['ticker']}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
 
# ================================================================ Tabs
t_dev, t_val, t_fin, t_score = st.tabs(["Entwicklung", "Bewertung", "Kennzahlen", "Scorecard"])
 
# ---------------------------------------------------- Entwicklung
with t_dev:
    a, b = st.columns(2)
    rev = frow(main, "Umsatz")
    g = np.concatenate([[np.nan], rev[1:] / rev[:-1] - 1]) if rev is not None else None
    fig = make_subplots(specs=[[{"secondary_y": True}]])
    fig.add_bar(x=yrs, y=rev, name="Umsatz (Mio.)", marker_color=NAVY)
    fig.add_scatter(x=yrs, y=g, name="Wachstum", line=dict(color=TEAL, width=2), secondary_y=True)
    fig.update_yaxes(tickformat=".0%", secondary_y=True)
    a.plotly_chart(_layout(fig, title="Umsatz & Wachstum"), use_container_width=True)
 
    fig = go.Figure()
    for nm, col in [("Bruttomarge", GREY), ("EBITDA-Marge", TEAL), ("EBIT-Marge", NAVY),
                    ("Nettomarge", GREEN), ("FCF-Marge", AMBER)]:
        s = row(main, nm)
        if s is not None:
            fig.add_scatter(x=yrs, y=s, name=nm, mode="lines+markers", line=dict(color=col))
    fig.update_yaxes(tickformat=".0%")
    b.plotly_chart(_layout(fig, title="Margenentwicklung"), use_container_width=True)
 
    c, e = st.columns(2)
    roic = row(main, "ROIC")
    fig = go.Figure()
    if roic is not None:
        fig.add_scatter(x=yrs, y=roic, name="ROIC", mode="lines+markers", line=dict(color=NAVY, width=2.5))
    fig.add_scatter(x=yrs, y=[wacc] * len(yrs), name="WACC", mode="lines",
                    line=dict(color=RED, width=2, dash="dash"))
    for nm, col in [("ROE", TEAL), ("ROA", GREY)]:
        s = row(main, nm)
        if s is not None:
            fig.add_scatter(x=yrs, y=s, name=nm, mode="lines", line=dict(color=col, width=1.5))
    fig.update_yaxes(tickformat=".0%")
    c.plotly_chart(_layout(fig, title="Kapitalrendite vs. WACC"), use_container_width=True)
 
    fcf = frow(main, "Free Cashflow")
    fig = make_subplots(specs=[[{"secondary_y": True}]])
    fig.add_bar(x=yrs, y=fcf, name="Free Cashflow (Mio.)", marker_color=GREEN)
    s = row(main, "FCF-Marge")
    if s is not None:
        fig.add_scatter(x=yrs, y=s, name="FCF-Marge", line=dict(color=NAVY, width=2), secondary_y=True)
    fig.update_yaxes(tickformat=".0%", secondary_y=True)
    e.plotly_chart(_layout(fig, title="Free Cashflow & FCF-Marge"), use_container_width=True)
 
    f, gcol = st.columns(2)
    nde = row(main, "Nettoverschuldung/EBITDA"); icov = row(main, "Zinsdeckung")
    fig = make_subplots(specs=[[{"secondary_y": True}]])
    if nde is not None:
        fig.add_bar(x=yrs, y=nde, name="Net Debt/EBITDA", marker_color=NAVY)
    if icov is not None:
        fig.add_scatter(x=yrs, y=icov, name="Zinsdeckung", line=dict(color=TEAL, width=2), secondary_y=True)
    f.plotly_chart(_layout(fig, title="Verschuldung & Zinsdeckung"), use_container_width=True)
 
    cc = row(main, "Cash Conversion (CFO/NI)")
    fig = go.Figure()
    if cc is not None:
        fig.add_scatter(x=yrs, y=cc, name="Cash Conversion (CFO/NI)", mode="lines+markers", line=dict(color=GREEN))
    fig.add_scatter(x=yrs, y=[1.0] * len(yrs), name="100%", mode="lines",
                    line=dict(color=GREY, width=1, dash="dot"))
    fig.update_yaxes(tickformat=".0%")
    gcol.plotly_chart(_layout(fig, title="Cashflow-Qualitaet"), use_container_width=True)
 
# ---------------------------------------------------- Bewertung
with t_val:
    left, right = st.columns([3, 2])
    with left:
        labels, vals, base = [], [], main["latest"]["price"]
        if not np.isnan(h["fair_value"]): labels.append("DCF Fair Value"); vals.append(h["fair_value"])
        if not np.isnan(impl_evebit): labels.append("Peer EV/EBIT"); vals.append(impl_evebit)
        if not np.isnan(impl_pe): labels.append("Peer KGV"); vals.append(impl_pe)
        fig = go.Figure()
        if vals:
            fig.add_bar(x=vals, y=labels, orientation="h",
                        marker_color=[GREEN if v >= base else RED for v in vals],
                        text=[f_eur(v) for v in vals], textposition="outside")
        if base and not np.isnan(base):
            fig.add_vline(x=base, line=dict(color=NAVY, width=2, dash="dash"),
                          annotation_text=f"Kurs {f_eur(base)}", annotation_position="top")
        st.plotly_chart(_layout(fig, h=260, title="Bewertungsanker (Football Field)"), use_container_width=True)
        if not peer_rs:
            st.caption("Peer-Anker erscheinen, sobald du oben Peers eingibst.")
    with right:
        st.markdown("**Aktuelle Multiplikatoren**")
        mult_df = pd.DataFrame({
            "Wert": [f_x(m["pe"]), f_x(m["ev_ebit"]), f_x(m["ev_ebitda"]), f_x(m["ev_fcf"]),
                     f_x(m["pb"]), f_pct(m["fcf_yield"]), f_pct(m["div_yield"])]},
            index=["KGV", "EV/EBIT", "EV/EBITDA", "EV/FCF", "KBV", "FCF-Rendite", "Dividendenrendite"])
        st.table(mult_df)
 
    st.markdown("**DCF (Free Cashflow to Firm)**")
    dc = st.columns(5)
    dc[0].metric("Basis-FCF (Mio.)", f_eur(d.get("base_fcf", np.nan) / eng.MM, 0) if d else "n/v")
    dc[1].metric("Wachstum (J1-5)", f_pct(d.get("growth", np.nan)) if d else "n/v")
    dc[2].metric("Terminal", f_pct(d.get("terminal", np.nan)) if d else "n/v")
    dc[3].metric("WACC", f_pct(d.get("wacc", np.nan)) if d else "n/v")
    dc[4].metric("Reverse-DCF impl. Wachstum", f_pct(main["reverse_growth"]))
    rc = st.columns(3)
    rc[0].metric("Fair Value je Aktie", f_eur(h["fair_value"]))
    rc[1].metric("Aktueller Kurs", f_eur(main["latest"]["price"]))
    rc[2].metric("Sicherheitsmarge", f_pct(h["mos"], 0))
    st.caption("Reverse-DCF: welches Dauerwachstum der aktuelle EV bereits einpreist (einstufige Naeherung).")
 
# ---------------------------------------------------- Kennzahlen
with t_fin:
    st.markdown("**Finanzdaten (Mio.)**")
    st.dataframe(main["financials"].style.format("{:,.0f}", na_rep="-"), use_container_width=True)
    st.markdown("**Kennzahlen**")
    mult_idx = ("Kapitalumschlag", "EK-Multiplikator", "Nettoverschuldung/EBITDA",
                "Zinsdeckung", "Verschuldungsgrad", "Current Ratio")
    def _fmt_ratio(v, idx):
        if pd.isna(v): return "-"
        return f"{v:.1f}x" if idx in mult_idx else f"{v:.1%}"
    disp = main["ratios"].copy()
    disp = disp.apply(lambda rr: [_fmt_ratio(v, rr.name) for v in rr], axis=1, result_type="expand")
    disp.columns = main["ratios"].columns
    st.dataframe(disp, use_container_width=True)
 
# ---------------------------------------------------- Scorecard
with t_score:
    labels = dict(geschaeftsmodell="Geschaeftsmodell*", management="Management*", wachstum="Wachstum",
                  profitabilitaet="Profitabilitaet (ROIC>WACC)", bilanz="Bilanz & Verschuldung",
                  cashflow="Cashflow-Qualitaet", margen="Margen", bewertung="Bewertung & Marge")
    order = list(main["weights"].keys())
    sc = [main["scores"][k] for k in order]
    fig = go.Figure(go.Bar(x=sc, y=[labels[k] for k in order], orientation="h",
                           marker_color=NAVY, text=sc, textposition="outside"))
    fig.update_xaxes(range=[0, 5])
    st.plotly_chart(_layout(fig, h=360, title="Scores je Kriterium (1-5)"), use_container_width=True)
    sc_df = pd.DataFrame({
        "Gewicht": [f_pct(main["weights"][k], 0) for k in order],
        "Score": [main["scores"][k] for k in order],
        "Beitrag": [round(main["weights"][k] * main["scores"][k], 2) for k in order],
    }, index=[labels[k] for k in order])
    st.table(sc_df)
    st.markdown(f"**Conviction-Score: {main['conviction']:.1f} / 5,0**  ·  Verdikt: "
                f"<span style='color:{verdict_color(main['verdict'])};font-weight:700'>{main['verdict']}</span>",
                unsafe_allow_html=True)
    st.caption("* Geschaeftsmodell und Management sind qualitativ (Slider links). "
               "Verdikt-Regel: Kaufen = Score >= 3,5 und Marge >= 20%; Halten = Score >= 3 und Marge >= 0; "
               "Meiden = Score < 2,5 oder Marge <= -10%.")
 
if main["missing"]:
    st.warning("Nicht gefundene Posten (manuell gegen Geschaeftsbericht pruefen): " + ", ".join(main["missing"]))
for tk, msg in errors:
    st.caption(f"Hinweis {tk}: {msg}")
