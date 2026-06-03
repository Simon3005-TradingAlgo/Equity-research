"""
app.py - Equity Research Dashboard (Einzeltitel)
================================================
Ein Ticker rein -> Profil, Kennzahlen, Zeitreihen-Charts, interaktive Bewertung
(DCF, Sensitivitaet, DDM, Multiplikatoren, kombinierter fairer Wert), Scorecard,
Verdikt, Excel-Export. Daten werden einmal geladen (gecacht); Annahmen-Slider
rechnen live ohne erneuten Abruf.

    streamlit run app.py
"""
import numpy as np
import pandas as pd
import streamlit as st
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import equity_engine as eng

st.set_page_config(page_title="Equity Research Dashboard", layout="wide")
st.markdown("<style>#MainMenu{visibility:hidden}footer{visibility:hidden}"
            "div[data-testid='stMetricValue']{font-size:1.25rem}</style>", unsafe_allow_html=True)

NAVY, TEAL, GREEN, RED, AMBER, GREY = "#1F3864", "#2E86AB", "#1E7D32", "#B71C1C", "#B8860B", "#8A8A8A"
FONT = dict(family="Arial, Helvetica, sans-serif", size=12, color="#222")

def _layout(fig, h=340, title=None, legend=True):
    fig.update_layout(template="plotly_white", height=h, font=FONT,
                      margin=dict(l=12, r=20, t=44 if title else 14, b=48 if legend else 14),
                      title=dict(text=title, font=dict(size=14, color=NAVY), x=0, xanchor="left") if title else None,
                      showlegend=legend,
                      legend=dict(orientation="h", yanchor="top", y=-0.14, x=0, font=dict(size=11)))
    return fig

def f_pct(x, d=1): return "n/v" if x is None or (isinstance(x, float) and np.isnan(x)) else f"{x*100:.{d}f}%"
def f_x(x):        return "n/v" if x is None or (isinstance(x, float) and np.isnan(x)) else f"{x:.1f}x"
def f_eur(x, d=2): return "n/v" if x is None or (isinstance(x, float) and np.isnan(x)) else f"{x:,.{d}f}"
def f_int(x):      return "n/v" if not x else f"{int(x):,}"
def row(r, n):  return r["ratios"].loc[n].values.astype(float) if n in r["ratios"].index else None
def frow(r, n): return r["financials"].loc[n].values.astype(float) if n in r["financials"].index else None
def vcolor(v):  return GREEN if "KAUFEN" in v else (RED if "MEIDEN" in v else AMBER)

# ---- lokale Bewertungsmodelle (rechnen live, ohne Netz) ----
def dcf_value(base_fcf, net_debt, shares, wacc, growth, years, terminal, exit_mult=None, ebitda=None):
    if not base_fcf or np.isnan(base_fcf) or not shares or np.isnan(shares):
        return dict(fair=np.nan, proj=[], pv=[], pv_tv=np.nan, ev=np.nan)
    proj = [base_fcf * (1 + growth) ** k for k in range(1, years + 1)]
    dfac = [(1 + wacc) ** -k for k in range(1, years + 1)]
    pv = [p * df for p, df in zip(proj, dfac)]
    if exit_mult and ebitda and not np.isnan(ebitda):
        tv = exit_mult * ebitda * (1 + growth) ** years
    elif wacc > terminal:
        tv = proj[-1] * (1 + terminal) / (wacc - terminal)
    else:
        return dict(fair=np.nan, proj=proj, pv=pv, pv_tv=np.nan, ev=np.nan)
    pv_tv = tv * dfac[-1]
    ev = sum(pv) + pv_tv
    return dict(fair=(ev - net_debt) / shares, proj=proj, pv=pv, pv_tv=pv_tv, ev=ev, tv=tv)

def ddm_value(dps, g, ke):
    if not dps or np.isnan(dps) or dps <= 0 or ke <= g: return np.nan
    return dps * (1 + g) / (ke - g)

# ================================================================ Sidebar
with st.sidebar:
    st.markdown("### Eingaben")
    ticker = st.text_input("Ticker", "DG.PA", help="EU-Suffixe: .PA .DE .MC .MI .AS .SW .L .VI")
    st.markdown("**Bewertungsannahmen**")
    wacc = st.slider("WACC", 0.04, 0.15, 0.08, 0.005, format="%.3f")
    auto_g = st.checkbox("FCF-Wachstum automatisch (max. 10%)", True)
    growth = None if auto_g else st.slider("FCF-Wachstum explizit", 0.0, 0.15, 0.06, 0.005, format="%.3f")
    terminal = st.slider("Terminales Wachstum", 0.0, 0.04, 0.025, 0.005, format="%.3f")
    st.markdown("**Qualitatives Scoring (1-5)**")
    qb = st.slider("Geschaeftsmodell / Burggraben", 1, 5, 3)
    qm = st.slider("Management / Kapitalallokation", 1, 5, 3)
    peers = st.text_input("Peers fuer relative Bewertung (optional)", "")
    run = st.button("Analyse laden / aktualisieren", type="primary", use_container_width=True)

st.title("Equity Research Dashboard")

@st.cache_data(show_spinner=False, ttl=3600)
def fetch_one(tk): return eng.fetch_fundamentals(tk)

if run:
    st.session_state.active = dict(ticker=ticker.strip(),
                                   peers=[p.strip() for p in peers.replace(";", ",").split(",") if p.strip()])
if "active" not in st.session_state:
    st.info("Ticker links eingeben und Analyse laden.")
    st.stop()

act = st.session_state.active
raw, errors = {}, []
with st.spinner("Lade Daten von Yahoo Finance ..."):
    for tk in [act["ticker"]] + act["peers"]:
        try:
            raw[tk] = fetch_one(tk)
        except Exception as ex:
            errors.append((tk, str(ex)))

if act["ticker"] not in raw:
    st.error(f"Keine Daten fuer '{act['ticker']}'. Ticker/Suffix pruefen oder erneut laden "
             f"(yfinance/Yahoo ist zeitweise instabil).")
    for tk, msg in errors:
        st.caption(f"{tk}: {msg}")
    st.stop()

# Live-Berechnung (gecachte Rohdaten -> kein erneuter Abruf bei Slider-Aenderung)
main = eng.compute(raw[act["ticker"]], wacc=wacc, growth=growth, terminal=terminal,
                   qual_business=qb, qual_management=qm)
peer_rs = [eng.compute(raw[tk], wacc=wacc, growth=growth, terminal=terminal)
           for tk in act["peers"] if tk in raw]

h, m, d = main["headline"], main["multiples"], main["dcf"]
yrs = [str(y) for y in main["years"]]
price = main["latest"]["price"]
shares = main["latest"]["shares"]
ebitda_l = main["latest"]["ebitda"]
eps_l = main["latest"]["eps"]
dps_l = main["latest"]["dps"]
nd_abs = frow(main, "Nettoverschuldung")[-1] * eng.MM
base_fcf = main["latest"]["fcf"]

def med(vals):
    vals = [v for v in vals if v is not None and not np.isnan(v)]
    return float(np.median(vals)) if vals else np.nan
peer_evebit = med([r["multiples"]["ev_ebit"] for r in peer_rs])
peer_pe = med([r["multiples"]["pe"] for r in peer_rs])

# ================================================================ Kopf
c1, c2 = st.columns([3, 2])
with c1:
    st.markdown(f"#### {main['name']}  ·  {main['ticker']}  ·  {main['currency']}")
    st.markdown(f"<div style='display:inline-block;padding:6px 16px;border-radius:4px;"
                f"background:{vcolor(main['verdict'])};color:#fff;font-weight:700'>{main['verdict']}</div>"
                f"<span style='margin-left:14px;color:{GREY}'>Conviction {main['conviction']:.1f} / 5,0</span>",
                unsafe_allow_html=True)
with c2:
    v = st.columns(3)
    v[0].metric("Aktueller Kurs", f_eur(price))
    v[1].metric("Fair Value (DCF)", f_eur(h["fair_value"]))
    v[2].metric("Sicherheitsmarge", f_pct(h["mos"], 0))
st.divider()

# KPI-Raster (4 breit -> kein Ueberlappen)
r1 = st.columns(4)
r1[0].metric("ROIC", f_pct(h["roic"]), f"{h['roic_wacc']*100:+.1f} Pp vs WACC")
r1[1].metric("EBIT-Marge", f_pct(h["ebit_margin"]))
r1[2].metric("Nettomarge", f_pct(row(main, "Nettomarge")[-1] if row(main, "Nettomarge") is not None else np.nan))
r1[3].metric("FCF-Marge", f_pct(h["fcf_margin"]))
r2 = st.columns(4)
r2[0].metric("Net Debt / EBITDA", f_x(h["nd_ebitda"]))
r2[1].metric("Zinsdeckung", f_x(row(main, "Zinsdeckung")[-1] if row(main, "Zinsdeckung") is not None else np.nan))
r2[2].metric("Cash Conversion", f_pct(h["cash_conv"], 0))
r2[3].metric("Umsatz-CAGR", f_pct(h["rev_cagr"]))
r3 = st.columns(4)
r3[0].metric("ROE", f_pct(row(main, "ROE")[-1] if row(main, "ROE") is not None else np.nan))
r3[1].metric("KGV", f_x(m["pe"]))
r3[2].metric("EV / EBIT", f_x(m["ev_ebit"]))
r3[3].metric("FCF-Rendite", f_pct(m["fcf_yield"]))

eng.to_excel([main] + peer_rs, "Equity_Analyse.xlsx")
with open("Equity_Analyse.xlsx", "rb") as fh:
    st.download_button("Excel-Bericht herunterladen", fh.read(), file_name=f"Equity_{main['ticker']}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

t_comp, t_dev, t_val, t_fin, t_score = st.tabs(
    ["Unternehmen", "Entwicklung", "Bewertung & Prognose", "Kennzahlen", "Scorecard"])

# ---------------------------------------------------- Unternehmen
with t_comp:
    p = main.get("profile") or {}
    facts = pd.DataFrame({"Angabe": [
        p.get("sector") or "n/v", p.get("industry") or "n/v",
        (f"{p.get('city')+', ' if p.get('city') else ''}{p.get('country') or 'n/v'}"),
        f_int(p.get("employees")), p.get("website") or "n/v"]},
        index=["Sektor", "Industrie", "Sitz", "Mitarbeiter", "Website"])
    cc = st.columns([2, 3])
    with cc[0]:
        st.table(facts)
        mm = st.columns(2)
        mm[0].metric("Marktkap. (Mio.)", f_eur(p["market_cap"] / eng.MM, 0) if p.get("market_cap") else "n/v")
        mm[1].metric("Beta", f"{p['beta']:.2f}" if p.get("beta") else "n/v")
    with cc[1]:
        lo, hi = p.get("lo52"), p.get("hi52")
        if lo and hi and price and not np.isnan(price):
            fig = go.Figure()
            fig.add_trace(go.Scatter(x=[lo, hi], y=[0, 0], mode="lines", line=dict(color=GREY, width=8)))
            fig.add_trace(go.Scatter(x=[price], y=[0], mode="markers+text", marker=dict(color=NAVY, size=16),
                                     text=[f"Kurs {f_eur(price)}"], textposition="top center"))
            fig.add_annotation(x=lo, y=0, text=f"Tief {f_eur(lo)}", showarrow=False, yshift=-22)
            fig.add_annotation(x=hi, y=0, text=f"Hoch {f_eur(hi)}", showarrow=False, yshift=-22)
            fig.update_yaxes(visible=False, range=[-1, 1])
            st.plotly_chart(_layout(fig, h=150, title="52-Wochen-Spanne", legend=False), use_container_width=True)
    st.markdown("**Geschaeftsbeschreibung**")
    st.write(p.get("summary") or "Keine Beschreibung von der Datenquelle verfuegbar.")
    st.caption("Segment-, Regions- und Wettbewerbsdetails sind in den Yahoo-Daten nicht strukturiert enthalten "
               "und stammen aus Geschaeftsbericht/IR. Die Beschreibung oben nennt diese oft im Fliesstext.")

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
        if s is not None: fig.add_scatter(x=yrs, y=s, name=nm, mode="lines+markers", line=dict(color=col))
    fig.update_yaxes(tickformat=".0%")
    b.plotly_chart(_layout(fig, title="Margenentwicklung"), use_container_width=True)

    c, e = st.columns(2)
    roic = row(main, "ROIC")
    fig = go.Figure()
    if roic is not None: fig.add_scatter(x=yrs, y=roic, name="ROIC", mode="lines+markers", line=dict(color=NAVY, width=2.5))
    fig.add_scatter(x=yrs, y=[wacc] * len(yrs), name="WACC", mode="lines", line=dict(color=RED, width=2, dash="dash"))
    for nm, col in [("ROE", TEAL), ("ROA", GREY)]:
        s = row(main, nm)
        if s is not None: fig.add_scatter(x=yrs, y=s, name=nm, mode="lines", line=dict(color=col, width=1.5))
    fig.update_yaxes(tickformat=".0%")
    c.plotly_chart(_layout(fig, title="Kapitalrendite vs. WACC"), use_container_width=True)

    fcf = frow(main, "Free Cashflow")
    fig = make_subplots(specs=[[{"secondary_y": True}]])
    fig.add_bar(x=yrs, y=fcf, name="Free Cashflow (Mio.)", marker_color=GREEN)
    s = row(main, "FCF-Marge")
    if s is not None: fig.add_scatter(x=yrs, y=s, name="FCF-Marge", line=dict(color=NAVY, width=2), secondary_y=True)
    fig.update_yaxes(tickformat=".0%", secondary_y=True)
    e.plotly_chart(_layout(fig, title="Free Cashflow & FCF-Marge"), use_container_width=True)

    f, gcol = st.columns(2)
    nde, icov = row(main, "Nettoverschuldung/EBITDA"), row(main, "Zinsdeckung")
    fig = make_subplots(specs=[[{"secondary_y": True}]])
    if nde is not None: fig.add_bar(x=yrs, y=nde, name="Net Debt/EBITDA", marker_color=NAVY)
    if icov is not None: fig.add_scatter(x=yrs, y=icov, name="Zinsdeckung", line=dict(color=TEAL, width=2), secondary_y=True)
    f.plotly_chart(_layout(fig, title="Verschuldung & Zinsdeckung"), use_container_width=True)

    cconv = row(main, "Cash Conversion (CFO/NI)")
    fig = go.Figure()
    if cconv is not None: fig.add_scatter(x=yrs, y=cconv, name="Cash Conversion", mode="lines+markers", line=dict(color=GREEN))
    fig.add_scatter(x=yrs, y=[1.0] * len(yrs), name="100%", mode="lines", line=dict(color=GREY, width=1, dash="dot"))
    fig.update_yaxes(tickformat=".0%")
    gcol.plotly_chart(_layout(fig, title="Cashflow-Qualitaet"), use_container_width=True)

# ---------------------------------------------------- Bewertung & Prognose
with t_val:
    st.markdown("**DCF-Annahmen (interaktiv - rechnet sofort)**")
    cs = st.columns(4)
    v_wacc = cs[0].slider("WACC", 0.04, 0.15, float(wacc), 0.005, format="%.3f", key="vw")
    v_g = cs[1].slider("FCF-Wachstum (Phase 1)", 0.0, 0.15,
                       float(d.get("growth", 0.06)) if d else 0.06, 0.005, format="%.3f", key="vg")
    v_years = cs[2].slider("Explizite Jahre", 3, 10, 5, key="vy")
    v_term = cs[3].slider("Terminales Wachstum", 0.0, 0.04, float(terminal), 0.005, format="%.3f", key="vt")
    method = st.radio("Terminalwert-Methode", ["Gordon-Growth", "Exit-Multiple (EV/EBITDA)"], horizontal=True)
    exit_mult = None
    if method.startswith("Exit"):
        default_x = float(np.clip(m["ev_ebitda"], 4, 20)) if not np.isnan(m["ev_ebitda"]) else 10.0
        exit_mult = st.slider("Exit EV/EBITDA", 4.0, 20.0, round(default_x, 1), 0.5)

    res = dcf_value(base_fcf, nd_abs, shares, v_wacc, v_g, v_years, v_term, exit_mult, ebitda_l)
    mos = (res["fair"] / price - 1) if (price and not np.isnan(price) and not np.isnan(res["fair"])) else np.nan

    mc = st.columns(4)
    mc[0].metric("Enterprise Value (Mio.)", f_eur(res["ev"] / eng.MM, 0) if not np.isnan(res["ev"]) else "n/v")
    mc[1].metric("Fair Value je Aktie", f_eur(res["fair"]))
    mc[2].metric("Aktueller Kurs", f_eur(price))
    mc[3].metric("Sicherheitsmarge", f_pct(mos, 0))

    g1, g2 = st.columns(2)
    # Barwert-Komposition
    if res["pv"]:
        labels = [f"J{k}" for k in range(1, v_years + 1)] + ["Terminal"]
        vals = [v / eng.MM for v in res["pv"]] + [res["pv_tv"] / eng.MM]
        cols = [NAVY] * v_years + [TEAL]
        fig = go.Figure(go.Bar(x=labels, y=vals, marker_color=cols))
        g1.plotly_chart(_layout(fig, title="Barwert-Komposition des Enterprise Value (Mio.)", legend=False),
                        use_container_width=True)
    # Sensitivitaet WACC x Terminal
    waccs = [round(v_wacc + x, 4) for x in (-0.01, -0.005, 0, 0.005, 0.01)]
    terms = [round(v_term + x, 4) for x in (-0.01, -0.005, 0, 0.005, 0.01)]
    Z = [[dcf_value(base_fcf, nd_abs, shares, w, v_g, v_years, tm)["fair"] for tm in terms] for w in waccs]
    fig = go.Figure(go.Heatmap(z=Z, x=[f"{t*100:.1f}%" for t in terms], y=[f"{w*100:.1f}%" for w in waccs],
                               colorscale="Blues", text=[[f_eur(v, 0) for v in r_] for r_ in Z],
                               texttemplate="%{text}", colorbar=dict(title="Fair Value")))
    fig.update_xaxes(title="Terminal-Wachstum"); fig.update_yaxes(title="WACC")
    g2.plotly_chart(_layout(fig, title="Sensitivitaet: Fair Value", legend=False), use_container_width=True)

    st.divider()
    st.markdown("**Weitere Bewertungsmodelle**")
    mcol = st.columns(3)
    with mcol[0]:
        st.caption("Dividend Discount Model (Gordon)")
        ke = st.slider("Eigenkapitalkosten (ke)", 0.04, 0.15, float(v_wacc), 0.005, format="%.3f", key="ke")
        g_div = st.slider("Dividendenwachstum", 0.0, 0.08, min(float(v_g), 0.04), 0.005, format="%.3f", key="gd")
        fair_ddm = ddm_value(dps_l, g_div, ke)
        st.metric("DDM Fair Value", f_eur(fair_ddm))
        st.caption(f"DPS letztes GJ: {f_eur(dps_l)}")
    with mcol[1]:
        st.caption("Multiplikator EV/EBIT")
        tgt_ev = st.number_input("Ziel EV/EBIT", 1.0, 60.0,
                                 round(float(peer_evebit if not np.isnan(peer_evebit) else m["ev_ebit"]), 1)
                                 if not np.isnan(m["ev_ebit"]) else 12.0, 0.5, key="tev")
        impl_ev = (tgt_ev * main["latest"]["ebit"] - nd_abs) / shares if shares else np.nan
        st.metric("Impl. Fair Value", f_eur(impl_ev))
    with mcol[2]:
        st.caption("Multiplikator KGV")
        tgt_pe = st.number_input("Ziel KGV", 1.0, 60.0,
                                 round(float(peer_pe if not np.isnan(peer_pe) else m["pe"]), 1)
                                 if not np.isnan(m["pe"]) else 15.0, 0.5, key="tpe")
        impl_pe = tgt_pe * eps_l if not np.isnan(eps_l) else np.nan
        st.metric("Impl. Fair Value", f_eur(impl_pe))

    # Kombinierter fairer Wert
    methods = {"DCF": res["fair"], "DDM": fair_ddm, "EV/EBIT": impl_ev, "KGV": impl_pe}
    methods = {k: v for k, v in methods.items() if v is not None and not np.isnan(v)}
    blended = float(np.median(list(methods.values()))) if methods else np.nan
    s1, s2 = st.columns([3, 2])
    with s1:
        fig = go.Figure(go.Bar(x=list(methods.values()), y=list(methods.keys()), orientation="h",
                               marker_color=[GREEN if v >= (price or 0) else RED for v in methods.values()],
                               text=[f_eur(v) for v in methods.values()], textposition="outside"))
        if price and not np.isnan(price):
            fig.add_vline(x=price, line=dict(color=NAVY, width=2, dash="dash"),
                          annotation_text=f"Kurs {f_eur(price)}", annotation_position="top")
        mx = max(list(methods.values()) + [price or 0]) * 1.2 if methods else 1
        fig.update_xaxes(range=[0, mx])
        s2.metric("Fair Value (Median der Modelle)", f_eur(blended),
                  f"{(blended/price-1)*100:+.0f}% vs Kurs" if (price and not np.isnan(price) and not np.isnan(blended)) else None)
        st.plotly_chart(_layout(fig, h=280, title="Fairer Wert je Modell vs. Kurs", legend=False),
                        use_container_width=True)
    st.caption("Reverse-DCF (einstufige Naeherung): der aktuelle EV preist ein Dauerwachstum von rund "
               f"{f_pct(main['reverse_growth'])} ein. Blended = Median der oben verfuegbaren Modelle, "
               "bewusst robust gegen Ausreisser.")

# ---------------------------------------------------- Kennzahlen
with t_fin:
    st.markdown("**Finanzdaten (Mio.)**")
    st.dataframe(main["financials"].style.format("{:,.0f}", na_rep="-"), use_container_width=True)
    st.markdown("**Kennzahlen**")
    mult_idx = ("Kapitalumschlag", "EK-Multiplikator", "Nettoverschuldung/EBITDA",
                "Zinsdeckung", "Verschuldungsgrad", "Current Ratio")
    def fmt(v, idx): return "-" if pd.isna(v) else (f"{v:.1f}x" if idx in mult_idx else f"{v:.1%}")
    disp = main["ratios"].apply(lambda rr: [fmt(v, rr.name) for v in rr], axis=1, result_type="expand")
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
    fig.update_xaxes(range=[0, 5.5])
    st.plotly_chart(_layout(fig, h=360, title="Scores je Kriterium (1-5)", legend=False), use_container_width=True)
    sc_df = pd.DataFrame({"Gewicht": [f_pct(main["weights"][k], 0) for k in order],
                          "Score": [main["scores"][k] for k in order],
                          "Beitrag": [round(main["weights"][k] * main["scores"][k], 2) for k in order]},
                         index=[labels[k] for k in order])
    st.table(sc_df)
    st.markdown(f"**Conviction-Score: {main['conviction']:.1f} / 5,0**  ·  Verdikt: "
                f"<span style='color:{vcolor(main['verdict'])};font-weight:700'>{main['verdict']}</span>",
                unsafe_allow_html=True)
    st.caption("* Geschaeftsmodell & Management sind qualitativ (Slider links). Verdikt-Regel: "
               "Kaufen = Score >= 3,5 und Marge >= 20%; Halten = Score >= 3 und Marge >= 0; "
               "Meiden = Score < 2,5 oder Marge <= -10%.")

if main["missing"]:
    st.warning("Nicht gefundene Posten (gegen Geschaeftsbericht pruefen): " + ", ".join(main["missing"]))
for tk, msg in errors:
    st.caption(f"Hinweis {tk}: {msg}")
