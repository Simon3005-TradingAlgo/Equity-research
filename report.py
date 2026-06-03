"""
report.py - professioneller PPTX-Bericht (Einzeltitel, ohne Anlageempfehlung)
=============================================================================
build_pptx(main, models, blended, consensus) -> bytes (PPTX)

Sachliche Aufbereitung: Bewertungsspanne, Kennzahlen, Qualitaets-Score und
Charts. Keine Kauf-/Verkaufsempfehlung. Charts via matplotlib (kein Kaleido).
"""
import io
from datetime import date
import numpy as np

NAVY = (31, 56, 100); TEAL = (46, 134, 171); GREEN = (30, 125, 50); RED = (183, 28, 28); GREY = (90, 90, 90)
LIGHT = (235, 240, 247)
MM = 1_000_000

def _ok(x): return x is not None and not (isinstance(x, float) and np.isnan(x))
def _eur(x, ccy="", d=2): return "n/v" if not _ok(x) else f"{x:,.{d}f}{(' ' + ccy) if ccy else ''}"
def _pct(x, d=1): return "n/v" if not _ok(x) else f"{x*100:.{d}f}%"
def _xx(x): return "n/v" if not _ok(x) else f"{x:.1f}x"

def _fig_rev_margin(main, ccy):
    import matplotlib; matplotlib.use("Agg"); import matplotlib.pyplot as plt
    yrs = [str(y) for y in main["years"]]
    rev = main["financials"].loc["Umsatz"].values.astype(float)
    ebm = main["ratios"].loc["EBIT-Marge"].values.astype(float) * 100
    fig, ax = plt.subplots(figsize=(6.2, 3.2), dpi=160)
    ax.bar(yrs, rev, color="#1F3864", width=0.6); ax.set_ylabel(f"Umsatz (Mio. {ccy})", color="#1F3864", fontsize=9)
    ax2 = ax.twinx(); ax2.plot(yrs, ebm, color="#2E86AB", marker="o", lw=2); ax2.set_ylabel("EBIT-Marge (%)", color="#2E86AB", fontsize=9)
    ax.set_title("Umsatz & EBIT-Marge", fontsize=11, color="#1F3864", loc="left")
    for s in ("top",): ax.spines[s].set_visible(False); ax2.spines[s].set_visible(False)
    ax.tick_params(labelsize=8); ax2.tick_params(labelsize=8)
    fig.tight_layout(); buf = io.BytesIO(); fig.savefig(buf, format="png"); plt.close(fig); buf.seek(0); return buf

def _fig_valuation(main, models, blended, ccy):
    import matplotlib; matplotlib.use("Agg"); import matplotlib.pyplot as plt
    price = main["latest"]["price"]
    items = [(k, v) for k, v in models.items() if _ok(v)]
    fig, ax = plt.subplots(figsize=(6.2, 3.2), dpi=160)
    cols = ["#1E7D32" if v >= (price or 0) else "#B71C1C" for _, v in items]
    ax.barh([k for k, _ in items], [v for _, v in items], color=cols)
    for i, (_, v) in enumerate(items): ax.text(v, i, f" {v:,.1f}", va="center", fontsize=8)
    if _ok(price):
        ax.axvline(price, color="#1F3864", ls="--", lw=2)
        ax.text(price, len(items) - 0.4, f" Kurs {price:,.1f}", color="#1F3864", fontsize=8)
    ax.set_title(f"Fairer Wert je Methode vs. Kurs ({ccy})", fontsize=11, color="#1F3864", loc="left")
    for s in ("top", "right"): ax.spines[s].set_visible(False)
    ax.tick_params(labelsize=8)
    fig.tight_layout(); buf = io.BytesIO(); fig.savefig(buf, format="png"); plt.close(fig); buf.seek(0); return buf

def build_pptx(main, models, blended, consensus=None):
    from pptx import Presentation
    from pptx.util import Inches, Pt
    from pptx.dml.color import RGBColor
    from pptx.enum.text import PP_ALIGN

    ccy = main.get("price_ccy") or main.get("currency") or ""
    price = main["latest"]["price"]; h, m = main["headline"], main["multiples"]
    vals = [v for v in models.values() if _ok(v)]
    lo, hi = (min(vals), max(vals)) if vals else (np.nan, np.nan)
    mos = (blended / price - 1) if _ok(price) and _ok(blended) else np.nan

    prs = Presentation(); prs.slide_width = Inches(13.333); prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]

    def tb(s, l, t, w, ht, text, size=14, bold=False, color=(0, 0, 0), align=PP_ALIGN.LEFT):
        b = s.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(ht)); f = b.text_frame
        f.word_wrap = True; p = f.paragraphs[0]; p.alignment = align
        r = p.add_run(); r.text = text; r.font.size = Pt(size); r.font.bold = bold
        r.font.color.rgb = RGBColor(*color); r.font.name = "Arial"; return b

    def chip(s, l, t, w, ht, text, color, fg=(255, 255, 255), size=14):
        sp = s.shapes.add_shape(1, Inches(l), Inches(t), Inches(w), Inches(ht))
        sp.fill.solid(); sp.fill.fore_color.rgb = RGBColor(*color); sp.line.fill.background()
        p = sp.text_frame.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
        r = p.add_run(); r.text = text; r.font.size = Pt(size); r.font.bold = True
        r.font.color.rgb = RGBColor(*fg); r.font.name = "Arial"

    def table(s, l, t, w, ht, rows, header):
        tbl = s.shapes.add_table(len(rows) + 1, len(header), Inches(l), Inches(t), Inches(w), Inches(ht)).table
        for j, hd in enumerate(header):
            c = tbl.cell(0, j); c.text = hd; c.fill.solid(); c.fill.fore_color.rgb = RGBColor(*NAVY)
            rr = c.text_frame.paragraphs[0].runs[0]; rr.font.color.rgb = RGBColor(255, 255, 255); rr.font.size = Pt(11); rr.font.bold = True
        for i, rowv in enumerate(rows, 1):
            for j, v in enumerate(rowv):
                c = tbl.cell(i, j); c.text = str(v) if v not in (None, "") else " "
                para = c.text_frame.paragraphs[0]
                if para.runs: para.runs[0].font.size = Pt(10)
        return tbl

    def footer(s):
        tb(s, 0.6, 7.05, 12, 0.4, "Nur zu Informationszwecken - keine Anlageempfehlung. "
           "Datenquelle: Yahoo Finance / FMP (zu plausibilisieren).", 9, False, GREY)

    # Slide 1 - Titel & Bewertungsspanne
    s = prs.slides.add_slide(blank); chip(s, 0, 0, 13.333, 0.18, "", NAVY)
    tb(s, 0.6, 1.2, 12, 1.0, main["name"], 32, True, NAVY)
    tb(s, 0.6, 2.25, 12, 0.5, f"{main['ticker']}  ·  {ccy}  ·  Equity Research  ·  {date.today().strftime('%d.%m.%Y')}", 15, False, GREY)
    tb(s, 0.6, 3.3, 6, 0.4, "Bewertungsspanne (alle Methoden)", 13, True, NAVY)
    tb(s, 0.6, 3.8, 7, 0.6, f"{_eur(lo, ccy)}   bis   {_eur(hi, ccy)}", 20, True, (0, 0, 0))
    tb(s, 0.6, 4.6, 7, 0.4, f"Median {_eur(blended, ccy)}   ·   Auf-/Abschlag vs. Kurs {_pct(mos, 0)}", 14, False, (0, 0, 0))
    tb(s, 8.0, 3.3, 4.5, 0.4, "Aktueller Kurs", 13, True, NAVY)
    tb(s, 8.0, 3.7, 4.5, 0.6, _eur(price, ccy), 22, True, (0, 0, 0))
    tb(s, 8.0, 4.7, 4.5, 0.4, f"Qualitaets-Score {main['conviction']:.1f} / 5,0", 14, False, GREY)
    footer(s)

    # Slide 2 - Kennzahlen & Methoden
    s = prs.slides.add_slide(blank)
    tb(s, 0.6, 0.4, 12, 0.6, "Kennzahlen & Bewertungsmethoden", 24, True, NAVY)
    kpis = [["ROIC", _pct(h["roic"])], ["ROIC - WACC", _pct(h["roic_wacc"])], ["EBIT-Marge", _pct(h["ebit_margin"])],
            ["FCF-Marge", _pct(h["fcf_margin"])], ["Net Debt/EBITDA", _xx(h["nd_ebitda"])],
            ["Cash Conversion", _pct(h["cash_conv"], 0)], ["Umsatz-CAGR", _pct(h["rev_cagr"])],
            ["KGV", _xx(m["pe"])], ["EV/EBIT", _xx(m["ev_ebit"])], ["FCF-Rendite", _pct(m["fcf_yield"])]]
    table(s, 0.6, 1.2, 5.4, 4.8, kpis, ["Kennzahl", "Wert"])
    meth = []
    for k, v in models.items():
        if not _ok(v): continue
        up = (v / price - 1) if _ok(price) else np.nan
        meth.append([k, _eur(v, ccy), _pct(up, 0)])
    meth.append(["Median", _eur(blended, ccy), _pct(mos, 0)])
    meth.append(["Aktueller Kurs", _eur(price, ccy), "-"])
    table(s, 6.4, 1.2, 6.3, 4.0, meth, ["Methode", "Fairer Wert", "Auf-/Abschlag"])
    footer(s)

    # Slide 3 - Charts
    s = prs.slides.add_slide(blank)
    tb(s, 0.6, 0.4, 12, 0.6, "Entwicklung & Bewertung", 24, True, NAVY)
    s.shapes.add_picture(_fig_rev_margin(main, ccy), Inches(0.5), Inches(1.5), width=Inches(6.1))
    s.shapes.add_picture(_fig_valuation(main, models, blended, ccy), Inches(6.8), Inches(1.5), width=Inches(6.1))
    footer(s)

    # Slide 4 - Qualitaets-Scorecard (kein Verdikt)
    s = prs.slides.add_slide(blank)
    tb(s, 0.6, 0.4, 12, 0.6, "Qualitaets-Scorecard", 24, True, NAVY)
    labels = dict(geschaeftsmodell="Geschaeftsmodell*", management="Management*", wachstum="Wachstum",
                  profitabilitaet="Profitabilitaet (ROIC>WACC)", bilanz="Bilanz & Verschuldung",
                  cashflow="Cashflow-Qualitaet", margen="Margen", bewertung="Bewertung (Niveau)")
    rows = [[labels[k], _pct(main["weights"][k], 0), str(main["scores"][k])] for k in main["weights"]]
    rows.append(["Qualitaets-Score (gewichtet)", "", f"{main['conviction']:.1f} / 5"])
    table(s, 0.6, 1.3, 8.2, 4.8, rows, ["Kriterium", "Gewicht", "Score (1-5)"])
    tb(s, 9.1, 1.4, 3.8, 3.5,
       "Der Qualitaets-Score bewertet ausschliesslich fundamentale Merkmale (Profitabilitaet, Bilanz, "
       "Cashflow, Wachstum, Bewertungsniveau). Er ist eine Einordnung der Unternehmensqualitaet, keine "
       "Kauf- oder Verkaufsempfehlung.\n\n* qualitativ gesetzt.", 11, False, GREY)
    footer(s)

    bio = io.BytesIO(); prs.save(bio); return bio.getvalue()

if __name__ == "__main__":
    import equity_engine as eng
    r = eng.compute(eng._synthetic())
    models = {"DCF (FCFF)": r["headline"]["fair_value"], "DDM (Gordon)": 22.9, "Analysten-Ziel": 42.0}
    cons = dict(target_mean=42.0, target_high=55.0, target_low=30.0, n_analysts=14, rec_key="buy")
    open("selftest_report.pptx", "wb").write(build_pptx(r, models, 33.4, cons)); print("PPTX ok")
