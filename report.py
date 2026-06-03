"""
report.py - kundenfertiger PPTX-Export des Equity-Research-Dashboards
=====================================================================
build_pptx(main, fair_value, mos, blended, consensus) -> bytes (PPTX)

Charts werden mit matplotlib gerendert (kein Kaleido/Chrome noetig -> laeuft
zuverlaessig auf Streamlit Cloud). python-pptx und matplotlib in requirements.
"""
import io
from datetime import date
import numpy as np

NAVY = (31, 56, 100); GREEN = (30, 125, 50); RED = (183, 28, 28); AMBER = (184, 134, 11); GREY = (90, 90, 90)
MM = 1_000_000

def _vc(v): return GREEN if "KAUFEN" in v else (RED if "MEIDEN" in v else AMBER)
def _f_eur(x, d=2):
    return "n/v" if x is None or (isinstance(x, float) and np.isnan(x)) else f"{x:,.{d}f}"
def _f_pct(x, d=1):
    return "n/v" if x is None or (isinstance(x, float) and np.isnan(x)) else f"{x*100:.{d}f}%"
def _f_x(x):
    return "n/v" if x is None or (isinstance(x, float) and np.isnan(x)) else f"{x:.1f}x"

def _fig_rev_margin(main):
    import matplotlib; matplotlib.use("Agg"); import matplotlib.pyplot as plt
    yrs = [str(y) for y in main["years"]]
    rev = main["financials"].loc["Umsatz"].values.astype(float)
    ebm = main["ratios"].loc["EBIT-Marge"].values.astype(float) * 100
    fig, ax = plt.subplots(figsize=(6.2, 3.1), dpi=160)
    ax.bar(yrs, rev, color="#1F3864", width=0.6)
    ax.set_ylabel("Umsatz (Mio.)", color="#1F3864")
    ax2 = ax.twinx(); ax2.plot(yrs, ebm, color="#2E86AB", marker="o", lw=2)
    ax2.set_ylabel("EBIT-Marge (%)", color="#2E86AB")
    ax.set_title("Umsatz & EBIT-Marge", fontsize=11, color="#1F3864", loc="left")
    for s in ("top",): ax.spines[s].set_visible(False); ax2.spines[s].set_visible(False)
    fig.tight_layout(); buf = io.BytesIO(); fig.savefig(buf, format="png"); plt.close(fig); buf.seek(0)
    return buf

def _fig_fair(main, fair, blended, consensus):
    import matplotlib; matplotlib.use("Agg"); import matplotlib.pyplot as plt
    price = main["latest"]["price"]
    items = [("DCF", fair), ("Median Modelle", blended),
             ("Analysten-Ziel", (consensus or {}).get("target_mean"))]
    items = [(k, v) for k, v in items if v is not None and not (isinstance(v, float) and np.isnan(v))]
    fig, ax = plt.subplots(figsize=(6.2, 3.1), dpi=160)
    cols = ["#1E7D32" if v >= (price or 0) else "#B71C1C" for _, v in items]
    ax.barh([k for k, _ in items], [v for _, v in items], color=cols)
    for i, (_, v) in enumerate(items):
        ax.text(v, i, f" {_f_eur(v)}", va="center", fontsize=9)
    if price and not np.isnan(price):
        ax.axvline(price, color="#1F3864", ls="--", lw=2)
        ax.text(price, len(items) - 0.4, f" Kurs {_f_eur(price)}", color="#1F3864", fontsize=9)
    ax.set_title("Fairer Wert je Modell vs. Kurs", fontsize=11, color="#1F3864", loc="left")
    for s in ("top", "right"): ax.spines[s].set_visible(False)
    fig.tight_layout(); buf = io.BytesIO(); fig.savefig(buf, format="png"); plt.close(fig); buf.seek(0)
    return buf

def build_pptx(main, fair_value, mos, blended, consensus=None):
    from pptx import Presentation
    from pptx.util import Inches, Pt
    from pptx.dml.color import RGBColor
    from pptx.enum.text import PP_ALIGN

    prs = Presentation(); prs.slide_width = Inches(13.333); prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]
    h, m = main["headline"], main["multiples"]

    def textbox(slide, l, t, w, ht, text, size=14, bold=False, color=(0, 0, 0), align=PP_ALIGN.LEFT):
        tb = slide.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(ht)); tf = tb.text_frame
        tf.word_wrap = True; p = tf.paragraphs[0]; p.alignment = align
        r = p.add_run(); r.text = text; r.font.size = Pt(size); r.font.bold = bold
        r.font.color.rgb = RGBColor(*color); r.font.name = "Arial"
        return tb

    def bar(slide, l, t, w, ht, text, color):
        sp = slide.shapes.add_shape(1, Inches(l), Inches(t), Inches(w), Inches(ht))
        sp.fill.solid(); sp.fill.fore_color.rgb = RGBColor(*color); sp.line.fill.background()
        tf = sp.text_frame; p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
        r = p.add_run(); r.text = text; r.font.size = Pt(16); r.font.bold = True
        r.font.color.rgb = RGBColor(255, 255, 255); r.font.name = "Arial"

    # ---- Slide 1: Titel ----
    s = prs.slides.add_slide(blank)
    bar(s, 0, 0, 13.333, 0.18, "", NAVY)
    textbox(s, 0.6, 1.4, 12, 1.0, f"{main['name']}", 34, True, NAVY)
    textbox(s, 0.6, 2.5, 12, 0.5, f"{main['ticker']}  ·  {main['currency']}  ·  Equity Research", 16, False, GREY)
    bar(s, 0.6, 3.4, 3.2, 0.7, main["verdict"], _vc(main["verdict"]))
    textbox(s, 4.0, 3.5, 8, 0.6, f"Conviction {main['conviction']:.1f} / 5,0", 18, True, (0, 0, 0))
    textbox(s, 0.6, 4.4, 12, 0.5,
            f"Kurs {_f_eur(main['latest']['price'])}   ·   Fair Value (DCF) {_f_eur(fair_value)}   ·   "
            f"Sicherheitsmarge {_f_pct(mos, 0)}", 16, False, (0, 0, 0))
    textbox(s, 0.6, 6.8, 12, 0.4, f"Stand: {date.today().strftime('%d.%m.%Y')}  ·  Quelle: Yahoo Finance (zu pruefen)",
            10, False, GREY)

    # ---- Slide 2: Kennzahlen & Bewertung ----
    s = prs.slides.add_slide(blank)
    textbox(s, 0.6, 0.4, 12, 0.6, "Kennzahlen & Bewertung", 24, True, NAVY)
    kpis = [("ROIC", _f_pct(h["roic"])), ("ROIC - WACC", _f_pct(h["roic_wacc"])),
            ("EBIT-Marge", _f_pct(h["ebit_margin"])), ("FCF-Marge", _f_pct(h["fcf_margin"])),
            ("Net Debt/EBITDA", _f_x(h["nd_ebitda"])), ("Cash Conversion", _f_pct(h["cash_conv"], 0)),
            ("Umsatz-CAGR", _f_pct(h["rev_cagr"])), ("KGV", _f_x(m["pe"])),
            ("EV/EBIT", _f_x(m["ev_ebit"])), ("FCF-Rendite", _f_pct(m["fcf_yield"]))]
    rows = len(kpis) + 1
    tbl = s.shapes.add_table(rows, 2, Inches(0.6), Inches(1.2), Inches(5.6), Inches(5.2)).table
    tbl.cell(0, 0).text = "Kennzahl"; tbl.cell(0, 1).text = "Wert"
    for i, (kk, vv) in enumerate(kpis, 1):
        tbl.cell(i, 0).text = kk; tbl.cell(i, 1).text = vv
    for ci in range(2):
        c0 = tbl.cell(0, ci); c0.fill.solid(); c0.fill.fore_color.rgb = RGBColor(*NAVY)
        c0.text_frame.paragraphs[0].runs[0].font.color.rgb = RGBColor(255, 255, 255)
    val = [("Fair Value (DCF)", _f_eur(fair_value)), ("Sicherheitsmarge", _f_pct(mos, 0)),
           ("Fair Value (Median Modelle)", _f_eur(blended)),
           ("Analysten-Kursziel (Mittel)", _f_eur((consensus or {}).get("target_mean"))),
           ("Analysten (Hoch/Tief)", f"{_f_eur((consensus or {}).get('target_high'))} / "
            f"{_f_eur((consensus or {}).get('target_low'))}"),
           ("Empfehlung", str((consensus or {}).get("rec_key") or "n/v")),
           ("Reverse-DCF impl. Wachstum", _f_pct(main["reverse_growth"]))]
    tbl2 = s.shapes.add_table(len(val) + 1, 2, Inches(6.5), Inches(1.2), Inches(6.2), Inches(3.7)).table
    tbl2.cell(0, 0).text = "Bewertung"; tbl2.cell(0, 1).text = "Wert"
    for i, (kk, vv) in enumerate(val, 1):
        tbl2.cell(i, 0).text = kk; tbl2.cell(i, 1).text = vv
    for ci in range(2):
        c0 = tbl2.cell(0, ci); c0.fill.solid(); c0.fill.fore_color.rgb = RGBColor(*NAVY)
        c0.text_frame.paragraphs[0].runs[0].font.color.rgb = RGBColor(255, 255, 255)

    # ---- Slide 3: Charts ----
    s = prs.slides.add_slide(blank)
    textbox(s, 0.6, 0.4, 12, 0.6, "Entwicklung & Bewertung", 24, True, NAVY)
    s.shapes.add_picture(_fig_rev_margin(main), Inches(0.5), Inches(1.4), width=Inches(6.1))
    s.shapes.add_picture(_fig_fair(main, fair_value, blended, consensus), Inches(6.8), Inches(1.4), width=Inches(6.1))

    # ---- Slide 4: Scorecard ----
    s = prs.slides.add_slide(blank)
    textbox(s, 0.6, 0.4, 12, 0.6, "Scorecard & These", 24, True, NAVY)
    labels = dict(geschaeftsmodell="Geschaeftsmodell*", management="Management*", wachstum="Wachstum",
                  profitabilitaet="Profitabilitaet (ROIC>WACC)", bilanz="Bilanz & Verschuldung",
                  cashflow="Cashflow-Qualitaet", margen="Margen", bewertung="Bewertung & Marge")
    order = list(main["weights"].keys())
    tbl = s.shapes.add_table(len(order) + 2, 3, Inches(0.6), Inches(1.3), Inches(8.0), Inches(4.8)).table
    for j, hd in enumerate(["Kriterium", "Gewicht", "Score (1-5)"]):
        tbl.cell(0, j).text = hd; c = tbl.cell(0, j); c.fill.solid(); c.fill.fore_color.rgb = RGBColor(*NAVY)
        c.text_frame.paragraphs[0].runs[0].font.color.rgb = RGBColor(255, 255, 255)
    for i, k in enumerate(order, 1):
        tbl.cell(i, 0).text = labels[k]; tbl.cell(i, 1).text = _f_pct(main["weights"][k], 0)
        tbl.cell(i, 2).text = str(main["scores"][k])
    tbl.cell(len(order) + 1, 0).text = "Conviction-Score"; tbl.cell(len(order) + 1, 2).text = f"{main['conviction']:.1f}"
    bar(s, 9.0, 1.5, 3.5, 0.8, main["verdict"], _vc(main["verdict"]))
    textbox(s, 9.0, 2.6, 3.8, 2.0,
            "Verdikt-Regel: Kaufen = Score >= 3,5 und Marge >= 20%; Halten = Score >= 3 und Marge >= 0; "
            "Meiden = Score < 2,5 oder Marge <= -10%.\n\n* qualitativ gesetzt.", 11, False, GREY)

    bio = io.BytesIO(); prs.save(bio); return bio.getvalue()

if __name__ == "__main__":
    import equity_engine as eng
    r = eng.compute(eng._synthetic())
    cons = dict(target_mean=42.0, target_high=55.0, target_low=30.0, n_analysts=14,
                rec_key="buy", rec_mean=2.1)
    data = build_pptx(r, r["headline"]["fair_value"], r["headline"]["mos"], 33.4, cons)
    open("selftest_report.pptx", "wb").write(data); print("PPTX ok, bytes:", len(data))
