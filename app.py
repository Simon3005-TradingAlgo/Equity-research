"""
app.py · Interaktives Web-Dashboard für die Aktienanalyse-Engine
================================================================
Lokal:   streamlit run app.py
Deploy:  equity_engine.py + app.py + requirements.txt -> GitHub -> share.streamlit.io
"""
import io
import numpy as np
import pandas as pd
import streamlit as st
import equity_engine as eng

st.set_page_config(page_title="Aktienanalyse-Engine", page_icon="📊", layout="wide")
st.title("📊 Fundamentale Aktienanalyse")
st.caption("Ticker eingeben → Bilanz/Cashflow/Gewinn-Kennzahlen, Bewertung (DCF · Reverse-DCF · Peer-Median) "
           "und regelbasiertes Verdikt. EU-Suffixe: .PA .DE .MC .MI .AS .SW .L .VI")

with st.sidebar:
    st.header("Eingaben")
    tickers = st.text_input("Ticker (Komma-getrennt)", "DG.PA, FER.MC, EIFF.PA")
    wacc = st.slider("WACC", 0.04, 0.15, 0.08, 0.005, format="%.3f")
    auto_g = st.checkbox("Wachstum automatisch (aus Historie, max. 10%)", True)
    growth = None if auto_g else st.slider("FCF-Wachstum explizit", 0.0, 0.15, 0.06, 0.005, format="%.3f")
    terminal = st.slider("Terminales Wachstum", 0.0, 0.04, 0.025, 0.005, format="%.3f")
    qb = st.slider("Geschäftsmodell / Burggraben (1–5)", 1, 5, 3)
    qm = st.slider("Management / Kapitalallokation (1–5)", 1, 5, 3)
    run = st.button("Analyse starten", type="primary", use_container_width=True)

@st.cache_data(show_spinner=False)
def _run(tk, wacc, growth, terminal, qb, qm):
    return eng.analyse(tk, wacc=wacc, growth=growth, terminal=terminal,
                       qual_business=qb, qual_management=qm, verbose=False)

def _fmt_verdict(v):
    color = {"KAUFEN": "#2e7d32"}.get(v, "#b71c1c" if "MEIDEN" in v else "#f9a825")
    return f"<span style='color:{color};font-weight:700'>{v}</span>"

if run:
    with st.spinner("Lade Daten von Yahoo Finance …"):
        results, errors = _run(tickers, wacc, growth, terminal, qb, qm)
    for tk, msg in errors:
        st.warning(f"{tk}: {msg}")
    if not results:
        st.stop()

    # --- Vergleichstabelle ---
    st.subheader("Vergleich")
    rows = []
    for r in results:
        h = r["headline"]
        rows.append(dict(Ticker=r["ticker"], Name=(r["name"] or "")[:30], Verdikt=r["verdict"],
                         Conviction=round(r["conviction"], 1), Kurs=r["latest"]["price"],
                         FairValue=h["fair_value"], Marge=h["mos"], ROIC=h["roic"],
                         **{"ROIC-WACC": h["roic_wacc"], "ND/EBITDA": h["nd_ebitda"],
                            "EBIT-Marge": h["ebit_margin"], "FCF-Marge": h["fcf_margin"],
                            "Cash Conv": h["cash_conv"], "KGV": r["multiples"]["pe"],
                            "EV/EBIT": r["multiples"]["ev_ebit"], "Rev-CAGR": h["rev_cagr"]}))
    df = pd.DataFrame(rows).set_index("Ticker")
    st.dataframe(df.style.format({
        "Kurs": "{:.2f}", "FairValue": "{:.2f}", "Conviction": "{:.1f}",
        "Marge": "{:.1%}", "ROIC": "{:.1%}", "ROIC-WACC": "{:.1%}", "EBIT-Marge": "{:.1%}",
        "FCF-Marge": "{:.1%}", "Cash Conv": "{:.0%}", "Rev-CAGR": "{:.1%}",
        "ND/EBITDA": "{:.1f}x", "KGV": "{:.1f}x", "EV/EBIT": "{:.1f}x"}, na_rep="–"),
        use_container_width=True)

    c1, c2 = st.columns(2)
    c1.bar_chart(df["ROIC-WACC"], color="#1F3864")
    c2.bar_chart(df["Marge"].rename("Sicherheitsmarge"), color="#2e7d32")

    # --- Excel-Download ---
    tmp = "Equity_Analyse.xlsx"
    eng.to_excel(results, tmp)
    with open(tmp, "rb") as f:
        st.download_button("⬇️ Excel herunterladen", f.read(), file_name="Equity_Analyse.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           use_container_width=True)

    # --- Detail je Titel ---
    st.subheader("Detail")
    for r in results:
        h = r["headline"]
        with st.expander(f"{r['ticker']} — {r['name']}", expanded=len(results) == 1):
            st.markdown(f"Verdikt: {_fmt_verdict(r['verdict'])} · "
                        f"Conviction **{r['conviction']:.1f}/5** · Fair Value **{h['fair_value']:.2f}** · "
                        f"Marge **{0 if np.isnan(h['mos']) else h['mos']:.0%}**", unsafe_allow_html=True)
            k = st.columns(4)
            k[0].metric("ROIC", f"{h['roic']:.1%}", f"{h['roic_wacc']:+.1%} vs WACC")
            k[1].metric("EBIT-Marge", f"{h['ebit_margin']:.1%}")
            k[2].metric("Net Debt/EBITDA", f"{h['nd_ebitda']:.1f}x")
            k[3].metric("Cash Conversion", f"{h['cash_conv']:.0%}")
            st.markdown("**Kennzahlen**")
            st.dataframe(r["ratios"].style.format("{:.1%}", na_rep="–"), use_container_width=True)
            st.markdown("**Finanzdaten (Mio.)**")
            st.dataframe(r["financials"].style.format("{:,.0f}", na_rep="–"), use_container_width=True)
            if r["missing"]:
                st.caption("⚠️ Nicht gefunden (manuell prüfen): " + ", ".join(r["missing"]))
    st.caption("Hinweis: yfinance-Daten immer gegen den Geschäftsbericht plausibilisieren. "
               "Zwei Score-Kategorien (Geschäftsmodell, Management) sind qualitativ und über die Slider gesetzt.")
else:
    st.info("Ticker links eingeben und **Analyse starten**.")
