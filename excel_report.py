"""
excel_report.py - umfassender, professioneller Excel-Bericht (Einzeltitel)
===========================================================================
build_excel(main, peer_rs, models, blended, consensus, wacc) -> bytes (XLSX)

Blaetter: Uebersicht | Finanzdaten | Kennzahlen | Bewertung | Scorecard | (Peers)
Werte sind in Python berechnet und werden direkt geschrieben (kein Neu-Rechnen
in Excel noetig). Professionelles, konsistentes Layout mit Charts.
"""
from io import BytesIO
import numpy as np
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import BarChart, Reference

MM = 1_000_000
F = "Arial"
HDR = PatternFill("solid", fgColor="1F3864"); SEC = PatternFill("solid", fgColor="D9E1F2")
GRN = PatternFill("solid", fgColor="C6EFCE"); REDF = PatternFill("solid", fgColor="FFC7CE")
YEL = PatternFill("solid", fgColor="FFF2CC"); LITE = PatternFill("solid", fgColor="F2F5FA")
thin = Side(style="thin", color="C8D0DC"); BORD = Border(thin, thin, thin, thin)
PCT = "0.0%"; MUL = '0.0"x"'; CUR = '#,##0;(#,##0);"-"'; PS = "#,##0.00"; UP = "+0.0%;-0.0%;0.0%"

def _vfill(v): return GRN if "KAUFEN" in v else (REDF if "MEIDEN" in v else YEL)
def _num(x): return None if x is None or (isinstance(x, float) and np.isnan(x)) else float(x)

def C(ws, coord, v=None, *, b=False, col="000000", fill=None, fmt=None, al=None, sz=10, wrap=False, bd=False):
    c = ws[coord]
    if v is not None: c.value = v
    c.font = Font(name=F, bold=b, color=col, size=sz)
    if fill: c.fill = fill
    if fmt: c.number_format = fmt
    if al or wrap: c.alignment = Alignment(horizontal=al, vertical="center", wrap_text=wrap)
    if bd: c.border = BORD
    return c

def title(ws, coord, span, text):
    C(ws, coord, text, b=True, col="FFFFFF", fill=HDR, sz=13)
    for col in range(ws[coord].column, ws[span].column + 1):
        ws.cell(ws[coord].row, col).fill = HDR

def section(ws, row, last, text):
    C(ws, f"A{row}", text, b=True, sz=11, fill=SEC)
    for col in range(1, ws[f"{last}{row}"].column + 1):
        ws.cell(row, col).fill = SEC

# --------------------------------------------------------------- Uebersicht
def _overview(ws, main, models, blended, consensus):
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 30
    for col in "BCDEF": ws.column_dimensions[col].width = 15
    title(ws, "A1", "F1", "EQUITY RESEARCH · UEBERSICHT")
    C(ws, "A2", main["name"], b=True, sz=14); ws.merge_cells("A2:D2")
    C(ws, "E2", f"{main['ticker']} · {main['currency']}", al="right", col="5A5A5A")
    C(ws, "A4", "VERDIKT", b=True, fill=SEC); C(ws, "B4", main["verdict"], b=True, fill=_vfill(main["verdict"]), al="center"); ws.merge_cells("B4:C4")
    C(ws, "D4", "Conviction", b=True, fill=SEC, al="right"); C(ws, "E4", round(main["conviction"], 1), b=True, fmt='0.0" / 5,0"', al="right")

    h, m = main["headline"], main["multiples"]
    section(ws, 6, "F", "Kennzahlen-Snapshot (letztes GJ)")
    kpis = [("Aktueller Kurs", _num(main["latest"]["price"]), PS), ("ROIC", _num(h["roic"]), PCT),
            ("ROIC - WACC", _num(h["roic_wacc"]), PCT), ("EBIT-Marge", _num(h["ebit_margin"]), PCT),
            ("FCF-Marge", _num(h["fcf_margin"]), PCT), ("Nettomarge", _num(main["ratios"].loc["Nettomarge"].iloc[-1]), PCT),
            ("Umsatz-CAGR", _num(h["rev_cagr"]), PCT), ("Net Debt/EBITDA", _num(h["nd_ebitda"]), MUL),
            ("Zinsdeckung", _num(main["ratios"].loc["Zinsdeckung"].iloc[-1]), MUL), ("Cash Conversion", _num(h["cash_conv"]), PCT),
            ("KGV", _num(m["pe"]), MUL), ("EV/EBIT", _num(m["ev_ebit"]), MUL),
            ("EV/EBITDA", _num(m["ev_ebitda"]), MUL), ("FCF-Rendite", _num(m["fcf_yield"]), PCT),
            ("Dividendenrendite", _num(m["div_yield"]), PCT), ("KBV", _num(m["pb"]), MUL)]
    r = 7
    for i in range(0, len(kpis), 2):
        for j in range(2):
            if i + j < len(kpis):
                lab, val, fmt = kpis[i + j]
                C(ws, f"{'A' if j==0 else 'D'}{r}", lab, bd=True)
                C(ws, f"{'B' if j==0 else 'E'}{r}", val, fmt=fmt, al="right", bd=True)
        r += 1

    # Bewertungsmethoden-Gegenueberstellung
    price = main["latest"]["price"]
    section(ws, r + 1, "F", "Bewertungsmethoden im Vergleich")
    hr = r + 2
    C(ws, f"A{hr}", "Methode", b=True, fill=LITE, bd=True); C(ws, f"B{hr}", "Fairer Wert", b=True, fill=LITE, al="right", bd=True)
    C(ws, f"C{hr}", "Upside vs. Kurs", b=True, fill=LITE, al="right", bd=True)
    rr = hr + 1
    for lab, val in models.items():
        up = (val / price - 1) if price and not np.isnan(price) else np.nan
        C(ws, f"A{rr}", lab, bd=True); C(ws, f"B{rr}", _num(val), fmt=PS, al="right", bd=True)
        c = C(ws, f"C{rr}", _num(up), fmt=UP, al="right", bd=True)
        if up is not None and not np.isnan(up): c.fill = GRN if up >= 0 else REDF
        rr += 1
    C(ws, f"A{rr}", "Median der Modelle", b=True, bd=True); C(ws, f"B{rr}", _num(blended), b=True, fmt=PS, al="right", bd=True)
    bup = (blended / price - 1) if price and not np.isnan(price) and not np.isnan(blended) else np.nan
    C(ws, f"C{rr}", _num(bup), b=True, fmt=UP, al="right", bd=True)
    C(ws, f"A{rr+1}", "Aktueller Kurs", b=True, fill=YEL, bd=True); C(ws, f"B{rr+1}", _num(price), b=True, fill=YEL, fmt=PS, al="right", bd=True)

    chart = BarChart(); chart.type = "bar"; chart.title = "Fairer Wert je Methode"; chart.height = 7; chart.width = 14
    chart.add_data(Reference(ws, min_col=2, min_row=hr + 1, max_row=rr + 1), titles_from_data=False)
    chart.set_categories(Reference(ws, min_col=1, min_row=hr + 1, max_row=rr + 1))
    chart.legend = None
    ws.add_chart(chart, f"D{r+1}")

# --------------------------------------------------------------- Finanzdaten
def _financials(ws, main):
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 26
    yrs = main["years"]
    for k in range(len(yrs)): ws.column_dimensions[chr(66 + k)].width = 13
    title(ws, "A1", f"{chr(65+len(yrs))}1", "FINANZDATEN (Mio., sofern nicht je Aktie)")
    C(ws, "A3", "Geschaeftsjahr", b=True, fill=SEC)
    for c, y in enumerate(yrs, 2): C(ws, f"{chr(64+c)}3", str(y), b=True, fill=SEC, al="right")
    row = 4
    for idx, vals in main["financials"].iterrows():
        fmt = PS if idx == "EPS" else CUR
        C(ws, f"A{row}", idx, bd=True)
        for c, v in enumerate(vals.values, 2): C(ws, f"{chr(64+c)}{row}", _num(v), fmt=fmt, al="right", bd=True)
        row += 1
    # Umsatzchart
    chart = BarChart(); chart.title = "Umsatz (Mio.)"; chart.height = 7; chart.width = 14; chart.legend = None
    chart.add_data(Reference(ws, min_col=2, max_col=1 + len(yrs), min_row=4, max_row=4), from_rows=True)
    chart.set_categories(Reference(ws, min_col=2, max_col=1 + len(yrs), min_row=3, max_row=3))
    ws.add_chart(chart, f"A{row+1}")

# --------------------------------------------------------------- Kennzahlen
def _ratios(ws, main):
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 30
    yrs = main["years"]
    for k in range(len(yrs)): ws.column_dimensions[chr(66 + k)].width = 13
    title(ws, "A1", f"{chr(65+len(yrs))}1", "KENNZAHLEN")
    groups = {
        "Margen": ["Bruttomarge", "EBITDA-Marge", "EBIT-Marge", "Nettomarge", "FCF-Marge"],
        "Kapitalrendite": ["ROE", "ROA", "ROIC", "ROIC - WACC"],
        "DuPont": ["DuPont Nettomarge", "Kapitalumschlag", "EK-Multiplikator"],
        "Cashflow-Qualitaet": ["Cash Conversion (CFO/NI)", "FCF/NI", "Accruals-Ratio"],
        "Verschuldung & Solvenz": ["Nettoverschuldung/EBITDA", "Zinsdeckung", "Verschuldungsgrad", "Current Ratio"],
    }
    mult_idx = {"Kapitalumschlag", "EK-Multiplikator", "Nettoverschuldung/EBITDA",
                "Zinsdeckung", "Verschuldungsgrad", "Current Ratio"}
    rat, row = main["ratios"], 3
    for grp, items in groups.items():
        section(ws, row, chr(65 + len(yrs)), grp)
        for c, y in enumerate(yrs, 2): C(ws, f"{chr(64+c)}{row}", str(y), b=True, fill=SEC, al="right")
        row += 1
        for name in items:
            if name not in rat.index: continue
            C(ws, f"A{row}", name, bd=True)
            for c, v in enumerate(rat.loc[name].values, 2):
                C(ws, f"{chr(64+c)}{row}", _num(v), fmt=(MUL if name in mult_idx else PCT), al="right", bd=True)
            row += 1
        row += 1

# --------------------------------------------------------------- Bewertung
def _valuation(ws, main, models, blended, consensus, wacc):
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 32
    for col in "BCDEF": ws.column_dimensions[col].width = 13
    title(ws, "A1", "F1", "BEWERTUNG")
    m, d = main["multiples"], main["dcf"]
    price = main["latest"]["price"]

    section(ws, 3, "B", "Aktuelle Multiplikatoren")
    mr = [("KGV", m["pe"], MUL), ("EV/EBIT", m["ev_ebit"], MUL), ("EV/EBITDA", m["ev_ebitda"], MUL),
          ("EV/FCF", m["ev_fcf"], MUL), ("KBV", m["pb"], MUL), ("FCF-Rendite", m["fcf_yield"], PCT),
          ("Dividendenrendite", m["div_yield"], PCT)]
    row = 4
    for lab, v, fmt in mr:
        C(ws, f"A{row}", lab, bd=True); C(ws, f"B{row}", _num(v), fmt=fmt, al="right", bd=True); row += 1

    section(ws, row + 1, "F", "DCF (Free Cashflow to Firm)")
    row += 2
    base = d.get("base_fcf", np.nan); g = d.get("growth", np.nan); term = d.get("terminal", np.nan); w = d.get("wacc", wacc)
    assum = [("Basis-FCF (Mio.)", base / MM if base and not np.isnan(base) else np.nan, CUR),
             ("Wachstum (J1-5)", g, PCT), ("Terminales Wachstum", term, PCT), ("WACC", w, PCT)]
    for lab, v, fmt in assum:
        C(ws, f"A{row}", lab, bd=True); C(ws, f"B{row}", _num(v), fmt=fmt, fill=YEL, al="right", bd=True); row += 1
    # Projektion
    if base and not np.isnan(base) and not np.isnan(g) and not np.isnan(w):
        C(ws, f"A{row}", "Jahr", b=True, fill=LITE, bd=True)
        for k in range(1, 6): C(ws, f"{chr(65+k)}{row}", f"J{k}", b=True, fill=LITE, al="right", bd=True)
        row += 1
        proj = [base * (1 + g) ** k for k in range(1, 6)]
        pv = [proj[k - 1] / (1 + w) ** k for k in range(1, 6)]
        C(ws, f"A{row}", "FCF (Mio.)", bd=True)
        for k in range(5): C(ws, f"{chr(66+k)}{row}", _num(proj[k] / MM), fmt=CUR, al="right", bd=True)
        row += 1
        C(ws, f"A{row}", "Barwert (Mio.)", bd=True)
        for k in range(5): C(ws, f"{chr(66+k)}{row}", _num(pv[k] / MM), fmt=CUR, al="right", bd=True)
        row += 1
    out = [("Enterprise Value (Mio.)", d.get("ev_dcf", np.nan) / MM if d.get("ev_dcf") else np.nan, CUR),
           ("Eigenkapitalwert (Mio.)", d.get("equity_value", np.nan) / MM if d.get("equity_value") else np.nan, CUR),
           ("Fair Value je Aktie", d.get("fair_value", np.nan), PS),
           ("Sicherheitsmarge", d.get("mos", np.nan), PCT),
           ("Reverse-DCF impl. Wachstum", main["reverse_growth"], PCT)]
    for lab, v, fmt in out:
        bold = lab.startswith("Fair") or lab.startswith("Sicher")
        C(ws, f"A{row}", lab, b=bold, bd=True); C(ws, f"B{row}", _num(v), b=bold, fmt=fmt, al="right",
                                                  fill=(SEC if bold else None), bd=True); row += 1

    section(ws, row + 1, "C", "Methoden-Gegenueberstellung")
    row += 2
    C(ws, f"A{row}", "Methode", b=True, fill=LITE, bd=True); C(ws, f"B{row}", "Fairer Wert", b=True, fill=LITE, al="right", bd=True)
    C(ws, f"C{row}", "Upside vs. Kurs", b=True, fill=LITE, al="right", bd=True); row += 1
    for lab, v in models.items():
        up = (v / price - 1) if price and not np.isnan(price) else np.nan
        C(ws, f"A{row}", lab, bd=True); C(ws, f"B{row}", _num(v), fmt=PS, al="right", bd=True)
        c = C(ws, f"C{row}", _num(up), fmt=UP, al="right", bd=True)
        if up is not None and not np.isnan(up): c.fill = GRN if up >= 0 else REDF
        row += 1
    C(ws, f"A{row}", "Median", b=True, bd=True); C(ws, f"B{row}", _num(blended), b=True, fmt=PS, al="right", bd=True)
    C(ws, f"A{row+1}", "Aktueller Kurs", b=True, fill=YEL, bd=True); C(ws, f"B{row+1}", _num(price), b=True, fill=YEL, fmt=PS, al="right", bd=True)

    if consensus and consensus.get("target_mean"):
        section(ws, row + 3, "B", "Analystenkonsens")
        rr = row + 4
        for lab, v, fmt in [("Kursziel Mittel", consensus.get("target_mean"), PS),
                            ("Kursziel Hoch", consensus.get("target_high"), PS),
                            ("Kursziel Tief", consensus.get("target_low"), PS),
                            ("Empfehlung", consensus.get("rec_key"), None),
                            ("Anzahl Analysten", consensus.get("n_analysts"), "0")]:
            C(ws, f"A{rr}", lab, bd=True)
            C(ws, f"B{rr}", (_num(v) if fmt else (v or "n/v")), fmt=fmt, al="right", bd=True); rr += 1

# --------------------------------------------------------------- Scorecard
def _scorecard(ws, main):
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 34
    for col in "BCD": ws.column_dimensions[col].width = 13
    title(ws, "A1", "D1", "SCORECARD & VERDIKT")
    labels = dict(geschaeftsmodell="Geschaeftsmodell*", management="Management*", wachstum="Wachstum",
                  profitabilitaet="Profitabilitaet (ROIC>WACC)", bilanz="Bilanz & Verschuldung",
                  cashflow="Cashflow-Qualitaet", margen="Margen", bewertung="Bewertung & Marge")
    C(ws, "A3", "Kriterium", b=True, fill=SEC, bd=True); C(ws, "B3", "Gewicht", b=True, fill=SEC, al="right", bd=True)
    C(ws, "C3", "Score (1-5)", b=True, fill=SEC, al="right", bd=True); C(ws, "D3", "Beitrag", b=True, fill=SEC, al="right", bd=True)
    row = 4
    for k in main["weights"]:
        C(ws, f"A{row}", labels[k], bd=True); C(ws, f"B{row}", main["weights"][k], fmt=PCT, al="right", bd=True)
        C(ws, f"C{row}", main["scores"][k], al="right", bd=True)
        C(ws, f"D{row}", round(main["weights"][k] * main["scores"][k], 2), fmt="0.00", al="right", bd=True); row += 1
    C(ws, f"A{row}", "Conviction-Score", b=True, bd=True); C(ws, f"D{row}", round(main["conviction"], 2), b=True, fmt="0.00", al="right", bd=True)
    C(ws, f"A{row+2}", "VERDIKT", b=True, fill=SEC); C(ws, f"B{row+2}", main["verdict"], b=True, fill=_vfill(main["verdict"]), al="center"); ws.merge_cells(f"B{row+2}:D{row+2}")
    C(ws, f"A{row+4}", "Regel: Kaufen = Score >= 3,5 und Marge >= 20%; Halten = Score >= 3 und Marge >= 0; "
                       "Meiden = Score < 2,5 oder Marge <= -10%. * qualitativ gesetzt.", col="5A5A5A", sz=9, wrap=True)
    ws.merge_cells(f"A{row+4}:D{row+5}")

# --------------------------------------------------------------- Peers
def _peers(ws, main, peer_rs):
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 26
    cols = [("Verdikt", None), ("Conviction", "0.0"), ("Kurs", PS), ("Fair Value", PS), ("Marge", PCT),
            ("ROIC", PCT), ("ROIC-WACC", PCT), ("ND/EBITDA", MUL), ("EBIT-Marge", PCT), ("KGV", MUL), ("EV/EBIT", MUL)]
    for i in range(len(cols)): ws.column_dimensions[chr(66 + i)].width = 12
    title(ws, "A1", f"{chr(66+len(cols))}1", "PEER-VERGLEICH")
    C(ws, "A2", "Titel", b=True, fill=SEC, bd=True)
    for j, (hn, _) in enumerate(cols, 2): C(ws, f"{chr(64+j)}2", hn, b=True, fill=SEC, al="right", bd=True)
    row = 3
    for r in [main] + peer_rs:
        hh, mm = r["headline"], r["multiples"]
        C(ws, f"A{row}", f"{r['ticker']}", b=(r is main), bd=True)
        vals = [r["verdict"], _num(r["conviction"]), _num(r["latest"]["price"]), _num(hh["fair_value"]),
                _num(hh["mos"]), _num(hh["roic"]), _num(hh["roic_wacc"]), _num(hh["nd_ebitda"]),
                _num(hh["ebit_margin"]), _num(mm["pe"]), _num(mm["ev_ebit"])]
        for j, (v, (_, fmt)) in enumerate(zip(vals, cols), 2):
            cc = C(ws, f"{chr(64+j)}{row}", v, fmt=fmt, al="right", bd=True)
            if j == 2: cc.fill = _vfill(r["verdict"])
        row += 1

def build_excel(main, peer_rs, models, blended, consensus, wacc):
    wb = Workbook()
    _overview(wb.active, main, models, blended, consensus); wb.active.title = "Uebersicht"
    _financials(wb.create_sheet("Finanzdaten"), main)
    _ratios(wb.create_sheet("Kennzahlen"), main)
    _valuation(wb.create_sheet("Bewertung"), main, models, blended, consensus, wacc)
    _scorecard(wb.create_sheet("Scorecard"), main)
    if peer_rs: _peers(wb.create_sheet("Peers"), main, peer_rs)
    bio = BytesIO(); wb.save(bio); return bio.getvalue()

if __name__ == "__main__":
    import equity_engine as eng
    r = eng.compute(eng._synthetic())
    models = {"DCF (FCFF)": r["headline"]["fair_value"], "DDM (Gordon)": 22.9,
              "EV/EBIT (Peer)": 49.4, "Analysten-Ziel": 42.0}
    cons = dict(target_mean=42.0, target_high=55.0, target_low=30.0, n_analysts=14, rec_key="buy")
    data = build_excel(r, [], models, 33.4, cons, 0.08)
    open("selftest_report.xlsx", "wb").write(data); print("Excel ok, bytes:", len(data))
