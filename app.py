"""
app.py - Equity Research Dashboard (Einzeltitel)
================================================
Ein Ticker rein -> Profil, Kennzahlen, Zeitreihen-Charts, interaktive Bewertung
(DCF, Sensitivitaet, DDM, Multiplikatoren, kombinierter fairer Wert), Scorecard,
Qualitaets-Score, Excel-/PPT-Export. Daten werden einmal geladen (gecacht); Slider
rechnen live ohne erneuten Abruf.

    streamlit run app.py
"""
import numpy as np
import pandas as pd
import streamlit as st
import io
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import equity_engine as eng
import report
import excel_report

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
    st.caption("FMP_API_KEY in den Streamlit-Secrets aktiviert die Kombination: FMP liefert "
               "Beschreibung/Profil/Kurs und fuellt Luecken in den Finanzdaten, Yahoo bleibt Basis.")

st.title("Equity Research Dashboard")

@st.cache_data(show_spinner=False, ttl=3600)
def fetch_one(tk, fmp_key=None): return eng.fetch_fundamentals(tk, fmp_key=fmp_key)

@st.cache_data(show_spinner=False, ttl=3600)
def fetch_history(tk):
    import yfinance as yf
    hh = yf.Ticker(tk).history(period="max", interval="1d")
    if hh is None or hh.empty: return None
    hh = hh[["Close"]].dropna()
    try: hh.index = hh.index.tz_localize(None)
    except (TypeError, AttributeError): pass
    return hh

def _ffill_to(index, dates, vals):
    s = pd.Series(np.asarray(vals, float), index=pd.DatetimeIndex(dates)).sort_index()
    s = s[~s.index.duplicated(keep="last")]
    return s.reindex(index.union(s.index)).sort_index().ffill().reindex(index)

def valuation_band(hist, series, kind):
    if hist is None or series is None or series.get("dates") is None: return None
    idx, close = hist.index, hist["Close"]
    if kind == "pe":
        eps_ff = _ffill_to(idx, series["dates"], series["eps"])
        mult = close / eps_ff.where(eps_ff > 0)
    else:
        sh = _ffill_to(idx, series["dates"], series["shares"])
        nd = _ffill_to(idx, series["dates"], series["net_debt"])
        eb = _ffill_to(idx, series["dates"], series["ebit"])
        mult = (close * sh + nd) / eb.where(eb > 0)
    mult = mult.replace([np.inf, -np.inf], np.nan).dropna()
    return mult if len(mult) > 30 else None

if run:
    st.session_state.active = dict(ticker=ticker.strip(),
                                   peers=[p.strip() for p in peers.replace(";", ",").split(",") if p.strip()])
if "active" not in st.session_state:
    st.info("Ticker links eingeben und Analyse laden.")
    st.stop()

act = st.session_state.active
try:
    fmp_key = st.secrets.get("FMP_API_KEY")
except Exception:
    fmp_key = None
raw, errors = {}, []
with st.spinner("Lade Daten von Yahoo Finance ..."):
    for tk in [act["ticker"]] + act["peers"]:
        try:
            raw[tk] = fetch_one(tk, fmp_key)
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
ccy = main.get("price_ccy") or main.get("currency") or ""
def f_cur(x, dd=2): return "n/v" if x is None or (isinstance(x, float) and np.isnan(x)) else f"{x:,.{dd}f} {ccy}"
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

# Analystenkonsens + Blended Fair Value (fuer Kopf/Export)
prof = main.get("profile") or {}
consensus = dict(target_mean=prof.get("target_mean"), target_high=prof.get("target_high"),
                 target_low=prof.get("target_low"), n_analysts=prof.get("n_analysts"),
                 rec_key=prof.get("rec_key"), rec_mean=prof.get("rec_mean"))
_g_used = float(d.get("growth", 0.04)) if d else 0.04
_ddm_rep = (dps_l * (1 + min(_g_used, 0.04)) / (wacc - min(_g_used, 0.04))
            if dps_l and not np.isnan(dps_l) and dps_l > 0 and wacc > min(_g_used, 0.04) else np.nan)
_methods_rep = [h["fair_value"], _ddm_rep]
if peer_rs:
    _methods_rep += [(peer_evebit * main["latest"]["ebit"] - nd_abs) / shares if shares and not np.isnan(peer_evebit) else np.nan,
                     peer_pe * eps_l if not np.isnan(peer_pe) else np.nan]
_methods_rep = [v for v in _methods_rep if v is not None and not np.isnan(v)]
blended_report = float(np.median(_methods_rep)) if _methods_rep else np.nan

models = {"DCF (FCFF)": h["fair_value"]}
if not np.isnan(_ddm_rep): models["DDM (Gordon)"] = _ddm_rep
if peer_rs:
    models["EV/EBIT (Peer)"] = ((peer_evebit * main["latest"]["ebit"] - nd_abs) / shares
                               if shares and not np.isnan(peer_evebit) else np.nan)
    models["KGV (Peer)"] = peer_pe * eps_l if not np.isnan(peer_pe) else np.nan
if consensus.get("target_mean"): models["Analysten-Ziel"] = consensus["target_mean"]
models = {k: v for k, v in models.items() if v is not None and not np.isnan(v)}

# ================================================================ Kopf
c1, c2 = st.columns([3, 2])
with c1:
    st.markdown(f"#### {main['name']}  ·  {main['ticker']}  ·  {ccy}")
    _vals = [v for v in models.values() if not np.isnan(v)]
    _lo, _hi = (min(_vals), max(_vals)) if _vals else (np.nan, np.nan)
    st.markdown(f"<span style='display:inline-block;padding:5px 14px;border-radius:4px;"
                f"background:#EBF0F7;color:#1F3864;font-weight:700'>Qualitaets-Score "
                f"{main['conviction']:.1f} / 5,0</span>", unsafe_allow_html=True)
    st.caption(f"Bewertungsspanne {f_cur(_lo)} – {f_cur(_hi)}  ·  Median {f_cur(blended_report)}")
with c2:
    v = st.columns(3)
    v[0].metric("Aktueller Kurs", f_cur(price))
    v[1].metric("Fair Value (Median)", f_cur(blended_report))
    _mos_h = (blended_report / price - 1) if price and not np.isnan(price) and not np.isnan(blended_report) else np.nan
    v[2].metric("Auf-/Abschlag", f_pct(_mos_h, 0))
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

def _artifact(key, sig, build_fn):
    if st.session_state.get(key + "_sig") != sig:
        try:
            st.session_state[key] = build_fn(); st.session_state[key + "_sig"] = sig; st.session_state[key + "_err"] = None
        except Exception as ex:
            st.session_state[key] = None; st.session_state[key + "_err"] = repr(ex)
    return st.session_state.get(key), st.session_state.get(key + "_err")

_sig = (main["ticker"], wacc, terminal, growth, qb, qm, len(peer_rs))
xls, xls_err = _artifact("xls", _sig, lambda: excel_report.build_excel(main, peer_rs, models, blended_report, consensus, wacc))
ppt, ppt_err = _artifact("ppt", _sig, lambda: report.build_pptx(main, models, blended_report, consensus))

dlc = st.columns(2)
if xls:
    dlc[0].download_button("Excel-Bericht herunterladen", xls, file_name=f"Equity_{main['ticker']}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           use_container_width=True)
else:
    dlc[0].error("Excel-Export fehlgeschlagen")
    st.caption(f"Excel: {xls_err}")
if ppt:
    dlc[1].download_button("PPT-Bericht herunterladen", ppt, file_name=f"Equity_{main['ticker']}.pptx",
                           mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                           use_container_width=True)
else:
    dlc[1].error("PPT-Export fehlgeschlagen")
    st.caption(f"PPT: {ppt_err}  (pruefe, ob python-pptx und matplotlib installiert sind / Repo neu deployt)")

t_comp, t_dev, t_val, t_fin, t_score, t_gloss = st.tabs(
    ["Unternehmen", "Entwicklung", "Bewertung & Prognose", "Kennzahlen", "Qualitaet", "Glossar"])

# ---------------------------------------------------- Unternehmen
with t_comp:
    hist = fetch_history(act["ticker"])
    if hist is not None and not hist.empty:
        oc = st.columns([1, 1, 2])
        ma50 = oc[0].checkbox("MA50", True)
        ma200 = oc[1].checkbox("MA200", True)
        span = oc[2].selectbox("Zeitraum", ["Max", "10 Jahre", "5 Jahre", "1 Jahr"], index=0)
        close = hist["Close"]
        ma50s, ma200s = close.rolling(50).mean(), close.rolling(200).mean()
        if span != "Max":
            n = {"10 Jahre": 10, "5 Jahre": 5, "1 Jahr": 1}[span]
            cut = close.index.max() - pd.DateOffset(years=n)
            mask = close.index >= cut
            close, ma50s, ma200s = close[mask], ma50s[mask], ma200s[mask]
        fig = go.Figure()
        fig.add_scatter(x=close.index, y=close.values, name="Kurs", line=dict(color=NAVY, width=1.6))
        if ma50: fig.add_scatter(x=ma50s.index, y=ma50s.values, name="MA50", line=dict(color=TEAL, width=1.3))
        if ma200: fig.add_scatter(x=ma200s.index, y=ma200s.values, name="MA200", line=dict(color=AMBER, width=1.3))
        st.plotly_chart(_layout(fig, h=340, title=f"Kursverlauf ({ccy})"), use_container_width=True)
    else:
        st.caption("Keine Kurshistorie verfuegbar (Yahoo zeitweise instabil - erneut laden).")

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
    summary = p.get("summary")
    if summary:
        st.write(summary)
    else:
        st.info("Aktuell keine Beschreibung von Yahoo verfuegbar. Das Profil-Feld ist auf geteilten "
                "Cloud-IPs oft leer - erneutes Laden hilft manchmal. Zuverlaessig kommt das Profil "
                "(inkl. Segmenten/Regionen) aus einer API wie Financial Modeling Prep.")
    st.caption("Segment-, Regions- und Wettbewerbsdetails sind in den Yahoo-Daten nicht strukturiert enthalten "
               "und stammen aus Geschaeftsbericht/IR. Die Beschreibung oben nennt diese oft im Fliesstext.")
    src = main.get("sources") or {}
    if src:
        st.caption(f"Datenquellen — Profil: {src.get('profile','Yahoo')}  ·  Kurs: {src.get('price','Yahoo')}  ·  "
                   f"FMP-Ergaenzungen Finanzdaten: {src.get('fmp_fill', 0)} Werte  ·  FMP-Status: {src.get('fmp_msg','-')}")

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
    mc[1].metric("Fair Value je Aktie", f_cur(res["fair"]))
    mc[2].metric("Aktueller Kurs", f_cur(price))
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
        st.metric("DDM Fair Value", f_cur(fair_ddm))
        st.caption(f"DPS letztes GJ: {f_cur(dps_l)}")
    with mcol[1]:
        st.caption("Multiplikator EV/EBIT")
        tgt_ev = st.number_input("Ziel EV/EBIT", 1.0, 60.0,
                                 round(float(peer_evebit if not np.isnan(peer_evebit) else m["ev_ebit"]), 1)
                                 if not np.isnan(m["ev_ebit"]) else 12.0, 0.5, key="tev")
        impl_ev = (tgt_ev * main["latest"]["ebit"] - nd_abs) / shares if shares else np.nan
        st.metric("Impl. Fair Value", f_cur(impl_ev))
    with mcol[2]:
        st.caption("Multiplikator KGV")
        tgt_pe = st.number_input("Ziel KGV", 1.0, 60.0,
                                 round(float(peer_pe if not np.isnan(peer_pe) else m["pe"]), 1)
                                 if not np.isnan(m["pe"]) else 15.0, 0.5, key="tpe")
        impl_pe = tgt_pe * eps_l if not np.isnan(eps_l) else np.nan
        st.metric("Impl. Fair Value", f_cur(impl_pe))

    # Kombinierter fairer Wert
    methods = {"DCF": res["fair"], "DDM": fair_ddm, "EV/EBIT": impl_ev, "KGV": impl_pe}
    if consensus.get("target_mean"):
        methods["Analysten-Ziel"] = consensus["target_mean"]
    methods = {k: v for k, v in methods.items() if v is not None and not np.isnan(v)}
    blended = float(np.median(list(methods.values()))) if methods else np.nan
    s1, s2 = st.columns([3, 2])
    with s1:
        fig = go.Figure(go.Bar(x=list(methods.values()), y=list(methods.keys()), orientation="h",
                               marker_color=[GREEN if v >= (price or 0) else RED for v in methods.values()],
                               text=[f"{v:,.1f}" for v in methods.values()], textposition="outside"))
        if price and not np.isnan(price):
            fig.add_vline(x=price, line=dict(color=NAVY, width=2, dash="dash"),
                          annotation_text=f"Kurs {price:,.1f}", annotation_position="top")
        mx = max(list(methods.values()) + [price or 0]) * 1.2 if methods else 1
        fig.update_xaxes(range=[0, mx])
        s2.metric("Fair Value (Median der Modelle)", f_cur(blended),
                  f"{(blended/price-1)*100:+.0f}% vs Kurs" if (price and not np.isnan(price) and not np.isnan(blended)) else None)
        st.plotly_chart(_layout(fig, h=280, title=f"Fairer Wert je Modell vs. Kurs ({ccy})", legend=False),
                        use_container_width=True)
    st.caption("Reverse-DCF (einstufige Naeherung): der aktuelle EV preist ein Dauerwachstum von rund "
               f"{f_pct(main['reverse_growth'])} ein. Blended = Median der oben verfuegbaren Modelle, "
               "bewusst robust gegen Ausreisser.")

    st.divider()
    st.markdown("**Analysten-Kursziele (externe Referenz)**")
    if consensus.get("target_mean"):
        ac = st.columns(4)
        ac[0].metric("Kursziel (Mittel)", f_cur(consensus["target_mean"]),
                     f"{(consensus['target_mean']/price-1)*100:+.0f}% vs Kurs"
                     if price and not np.isnan(price) else None)
        ac[1].metric("Kursziel (Hoch)", f_cur(consensus.get("target_high")))
        ac[2].metric("Kursziel (Tief)", f_cur(consensus.get("target_low")))
        ac[3].metric("Anzahl Analysten", f_int(consensus.get("n_analysts")))
    else:
        st.caption("Keine Analysten-Kursziele von der Datenquelle verfuegbar (bei EU-Titeln oft leer; "
                   "zuverlaessig ueber eine API mit Key).")

    st.divider()
    st.markdown("**Historische Bewertungsbaender**")
    hist_b = fetch_history(act["ticker"])
    bcols = st.columns(2)
    for col_, (kind, lab) in zip(bcols, [("pe", "KGV (P/E)"), ("ev_ebit", "EV/EBIT")]):
        mult = valuation_band(hist_b, main.get("series"), kind)
        if mult is None:
            col_.caption(f"{lab}: zu wenig Historie/Daten fuer ein Band.")
            continue
        mu, sd, cur = float(mult.mean()), float(mult.std()), float(mult.iloc[-1])
        fig = go.Figure()
        fig.add_scatter(x=mult.index, y=mult.values, name=lab, line=dict(color=NAVY, width=1.4))
        fig.add_hline(y=mu, line=dict(color=GREY, width=1.5, dash="dash"),
                      annotation_text=f"Mittel {mu:.1f}x", annotation_position="right")
        fig.add_hrect(y0=mu - sd, y1=mu + sd, fillcolor=TEAL, opacity=0.12, line_width=0)
        fig.add_scatter(x=[mult.index[-1]], y=[cur], mode="markers", name="aktuell",
                        marker=dict(color=AMBER, size=10))
        col_.plotly_chart(_layout(fig, h=300, title=f"{lab}-Band  ·  aktuell {cur:.1f}x"),
                          use_container_width=True)
    st.caption("Band = Mittelwert ± 1 Standardabweichung des Multiplikators ueber die verfuegbare Kurshistorie "
               "(annuelle Fundamentaldaten auf Tagesbasis fortgeschrieben). Aktueller Punkt = heutiger Multiplikator.")

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

# ---------------------------------------------------- Qualitaet
with t_score:
    labels = dict(geschaeftsmodell="Geschaeftsmodell*", management="Management*", wachstum="Wachstum",
                  profitabilitaet="Profitabilitaet (ROIC>WACC)", bilanz="Bilanz & Verschuldung",
                  cashflow="Cashflow-Qualitaet", margen="Margen", bewertung="Bewertung (Niveau)")
    order = list(main["weights"].keys())
    sc = [main["scores"][k] for k in order]
    fig = go.Figure(go.Bar(x=sc, y=[labels[k] for k in order], orientation="h",
                           marker_color=NAVY, text=sc, textposition="outside"))
    fig.update_xaxes(range=[0, 5.5])
    st.plotly_chart(_layout(fig, h=360, title="Qualitaets-Scores je Kriterium (1-5)", legend=False), use_container_width=True)
    sc_df = pd.DataFrame({"Gewicht": [f_pct(main["weights"][k], 0) for k in order],
                          "Score": [main["scores"][k] for k in order],
                          "Beitrag": [round(main["weights"][k] * main["scores"][k], 2) for k in order]},
                         index=[labels[k] for k in order])
    st.table(sc_df)
    st.markdown(f"**Qualitaets-Score (gewichtet): {main['conviction']:.1f} / 5,0**")
    st.caption("Der Qualitaets-Score bewertet ausschliesslich fundamentale Merkmale (Profitabilitaet, Bilanz, "
               "Cashflow, Wachstum, Bewertungsniveau). Er ist eine Einordnung der Unternehmensqualitaet, keine "
               "Kauf- oder Verkaufsempfehlung. * Geschaeftsmodell & Management qualitativ (Slider links).")

# ---------------------------------------------------- Glossar
with t_gloss:
    st.markdown("Erklaerung aller im Dashboard verwendeten Werte und Kennzahlen.")
    glossar = {
        "Margen & Rentabilitaet": [
            ("Bruttomarge", "Bruttogewinn / Umsatz - Umsatzanteil nach Herstellkosten."),
            ("EBITDA-Marge", "EBITDA / Umsatz - operative Marge vor Abschreibungen."),
            ("EBIT-Marge", "EBIT / Umsatz - operative Marge nach Abschreibungen."),
            ("Nettomarge", "Nettoergebnis / Umsatz - was nach allen Kosten und Steuern bleibt."),
            ("FCF-Marge", "Free Cashflow / Umsatz - wie viel Umsatz zu freiem Cash wird."),
        ],
        "Kapitalrendite & Wertschoepfung": [
            ("ROE", "Nettoergebnis / Eigenkapital - Rendite auf das Eigenkapital."),
            ("ROA", "Nettoergebnis / Bilanzsumme - Rendite auf die Aktiva."),
            ("NOPAT", "EBIT x (1 - Steuersatz) - operativer Nachsteuergewinn ohne Finanzierungseffekt."),
            ("Investiertes Kapital", "Finanzschulden + Eigenkapital - liquide Mittel."),
            ("ROIC", "NOPAT / investiertes Kapital - Rendite auf das eingesetzte Kapital."),
            ("WACC", "Gewichtete Kapitalkosten - Mindestrendite, die Kapitalgeber erwarten."),
            ("ROIC - WACC (Spread)", "Wertschoepfung: positiv = verdient mehr als die Kapitalkosten."),
        ],
        "DuPont (ROE-Zerlegung)": [
            ("DuPont", "Nettomarge x Kapitalumschlag x EK-Multiplikator = ROE."),
            ("Kapitalumschlag", "Umsatz / Bilanzsumme - wie effizient Aktiva Umsatz erzeugen."),
            ("EK-Multiplikator", "Bilanzsumme / Eigenkapital - finanzieller Hebel."),
        ],
        "Cashflow-Qualitaet": [
            ("Cash Conversion (CFO/NI)", "Operativer Cashflow / Nettoergebnis - nahe/ueber 100% = Gewinne durch Cash gedeckt."),
            ("FCF / NI", "Free Cashflow / Nettoergebnis."),
            ("Accruals-Ratio", "(Nettoergebnis - operativer Cashflow) / Bilanzsumme - hoch = Gewinn stark abgegrenzt (Warnsignal)."),
        ],
        "Verschuldung & Solvenz": [
            ("Nettoverschuldung", "Finanzschulden - liquide Mittel."),
            ("Net Debt / EBITDA", "Verschuldung relativ zum operativen Ergebnis - niedriger = solider."),
            ("Zinsdeckung", "EBIT / Zinsaufwand - wie oft die Zinsen verdient werden."),
            ("Verschuldungsgrad", "Schulden / Eigenkapital."),
            ("Current Ratio", "Kurzfristige Aktiva / kurzfristige Passiva - kurzfristige Liquiditaet."),
        ],
        "Bewertungs-Multiplikatoren": [
            ("Enterprise Value (EV)", "Marktkapitalisierung + Nettoverschuldung - Wert des Gesamtunternehmens."),
            ("KGV (P/E)", "Kurs / Gewinn je Aktie."),
            ("EV / EBIT, EV / EBITDA, EV / FCF", "EV relativ zu operativen Groessen - kapitalstrukturneutral."),
            ("KBV (P/B)", "Kurs / Buchwert je Aktie."),
            ("FCF-Rendite", "Free Cashflow / Marktkapitalisierung."),
            ("Dividendenrendite", "Dividende je Aktie / Kurs."),
        ],
        "Bewertungsmodelle": [
            ("DCF (FCFF)", "Barwert kuenftiger freier Cashflows; Fair Value = (Summe Barwerte + Terminalwert) - Nettoverschuldung, je Aktie."),
            ("Terminalwert (Gordon)", "Letzter FCF x (1+g) / (WACC - g) - ewiges Wachstum."),
            ("Terminalwert (Exit-Multiple)", "EV/EBITDA-Multiplikator auf das Endjahres-EBITDA."),
            ("Reverse-DCF", "Welches Dauerwachstum der aktuelle Kurs bereits einpreist (Disziplin-Check)."),
            ("DDM (Gordon)", "Dividende x (1+g) / (ke - g) - Bewertung ueber Dividenden."),
            ("Sicherheitsmarge", "Fair Value / Kurs - 1 - Puffer zwischen Schaetzung und Marktpreis."),
            ("CAGR", "Durchschnittliche jaehrliche Wachstumsrate ueber die Periode."),
        ],
        "Qualitaets-Score": [
            ("Qualitaets-Score", "Gewichteter Mittelwert von sieben fundamentalen Kriterien (Skala 1-5). "
                                 "Einordnung der Unternehmensqualitaet, keine Kauf-/Verkaufsempfehlung."),
            ("Bewertungsspanne", "Minimum bis Maximum der fairen Werte ueber alle Methoden; Median als zentrale Schaetzung."),
        ],
    }
    for grp, items in glossar.items():
        with st.expander(grp, expanded=(grp == "Kapitalrendite & Wertschoepfung")):
            st.table(pd.DataFrame({"Erklaerung": [e for _, e in items]}, index=[b for b, _ in items]))

if main["missing"]:
    st.warning("Nicht gefundene Posten (gegen Geschaeftsbericht pruefen): " + ", ".join(main["missing"]))
for tk, msg in errors:
    st.caption(f"Hinweis {tk}: {msg}")
