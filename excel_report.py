"""
excel_report.py - professioneller Excel-Bericht (Einzeltitel, ohne Anlageempfehlung)
====================================================================================
build_excel(main, peer_rs, models, blended, consensus, wacc) -> bytes (XLSX)

Blaetter: Uebersicht | Finanzdaten | Kennzahlen | Bewertung | Qualitaet | (Peers)
Sachliche Aufbereitung mit Bewertungsspanne und Methoden-Gegenueberstellung.
Keine Kauf-/Verkaufsempfehlung.
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
    for col in range(ws[coord].column, ws[span].column + 1): ws.cell(ws[coord].row, col).fill = HDR

def section(ws, row, last, text):
    C(ws, f"A{row}", text, b=True, sz=11, fill=SEC)
    for col in range(1, ws[f"{last}{row}"].column + 1): ws.cell(row, col).fill = SEC

# --------------------------------------------------------------- Uebersicht
def _overview(ws, main, models, blended, consensus, ccy):
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 30
    for col in "BCDEF": ws.column_dimensions[col].width = 15
    title(ws, "A1", "F1", "EQUITY RESEARCH · UEBERSICHT")
    C(ws, "A2", main["name"], b=True, sz=14); ws.merge_cells("A2:D2")
    C(ws, "E2", f"{main['ticker']} · {ccy}", al="right", col="5A5A5A")

    price = main["latest"]["price"]; h, m = main["headline"], main["multiples"]
    vals = [v for v in models.values() if v is not None and not np.isnan(v)]
    lo, hi = (min(vals), max(vals)) if vals else (np.nan, np.nan)
    mos = (blended / price - 1) if price and not np.isnan(price) and not np.isnan(blended) else np.nan
    C(ws, "A4", "Aktueller Kurs", b=True, fill=SEC, bd=True); C(ws, "B4", _num(price), b=True, fmt=PS, al="right", bd=True)
    C(ws, "C4", "Qualitaets-Score", b=True, fill=SEC, al="right", bd=True); C(ws, "D4", round(main["conviction"], 1), b=True, fmt='0.0" / 5,0"', al="right", bd=True)
    C(ws, "A5", "Bewertungsspanne (Methoden)", b=True, fill=SEC, bd=True)
    C(ws, "B5", f"{_num(lo):,.2f}  bis  {_num(hi):,.2f}" if vals else "n/v", b=True, al="right", bd=True); ws.merge_cells("B5:C5")
    C(ws, "D5", "Median / Auf-Abschlag", b=True, fill=SEC, al="right", bd=True)
    C(ws, "E5", _num(blended), b=True, fmt=PS, al="right", bd=True); C(ws, "F5", _num(mos), b=True, fmt=UP, al="right", bd=True)

    section(ws, 7, "F", f"Kennzahlen-Snapshot (letztes GJ, je-Aktie in {ccy})")
    kpis = [("Aktueller Kurs", _num(price), PS), ("ROIC", _num(h["roic"]), PCT),
            ("ROIC - WACC", _num(h["roic_wacc"]), PCT), ("EBIT-Marge", _num(h["ebit_margin"]), PCT),
            ("FCF-Marge", _num(h["fcf_margin"]), PCT), ("Nettomarge", _num(main["ratios"].loc["Nettomarge"].iloc[-1]), PCT),
            ("Umsatz-CAGR", _num(h["rev_cagr"]), PCT), ("Net Debt/EBITDA", _num(h["nd_ebitda"]), MUL),
            ("Zinsdeckung", _num(main["ratios"].loc["Zinsdeckung"].iloc[-1]), MUL), ("Cash Conversion", _num(h["cash_conv"]), PCT),
            ("KGV", _num(m["pe"]), MUL), ("EV/EBIT", _num(m["ev_ebit"]), MUL),
            ("EV/EBITDA", _num(m["ev_ebitda"]), MUL), ("FCF-Rendite", _num(m["fcf_yield"]), PCT),
            ("Dividendenrendite", _num(m["div_yield"]), PCT), ("KBV", _num(m["pb"]), MUL)]
    r = 8
    for i in range(0, len(kpis), 2):
        for j in range(2):
            if i + j < len(kpis):
                lab, val, fmt = kpis[i + j]
                C(ws, f"{'A' if j==0 else 'D'}{r}", lab, bd=True)
                C(ws, f"{'B' if j==0 else 'E'}{r}", val, fmt=fmt, al="right", bd=True)
        r += 1

    section(ws, r + 1, "F", "Bewertungsmethoden im Vergleich")
    hr = r + 2
    C(ws, f"A{hr}", "Methode", b=True, fill=LITE, bd=True); C(ws, f"B{hr}", "Fairer Wert", b=True, fill=LITE, al="right", bd=True)
    C(ws, f"C{hr}", "Auf-/Abschlag vs. Kurs", b=True, fill=LITE, al="right", bd=True)
    rr = hr + 1
    for lab, val in models.items():
        if val is None or np.isnan(val): continue
        up = (val / price - 1) if price and not np.isnan(price) else np.nan
        C(ws, f"A{rr}", lab, bd=True); C(ws, f"B{rr}", _num(val), fmt=PS, al="right", bd=True)
        c = C(ws, f"C{rr}", _num(up), fmt=UP, al="right", bd=True)
        if up is not None and not np.isnan(up): c.fill = GRN if up >= 0 else REDF
        rr += 1
    C(ws, f"A{rr}", "Median der Modelle", b=True, bd=True); C(ws, f"B{rr}", _num(blended), b=True, fmt=PS, al="right", bd=True)
    C(ws, f"C{rr}", _num(mos), b=True, fmt=UP, al="right", bd=True)
    C(ws, f"A{rr+1}", "Aktueller Kurs", b=True, fill=YEL, bd=True); C(ws, f"B{rr+1}", _num(price), b=True, fill=YEL, fmt=PS, al="right", bd=True)

    chart = BarChart(); chart.type = "bar"; chart.title = f"Fairer Wert je Methode ({ccy})"; chart.height = 7; chart.width = 13; chart.legend = None
    chart.add_data(Reference(ws, min_col=2, min_row=hr + 1, max_row=rr + 1), titles_from_data=False)
    chart.set_categories(Reference(ws, min_col=1, min_row=hr + 1, max_row=rr + 1))
    ws.add_chart(chart, f"D{r+1}")
    C(ws, f"A{rr+3}", "Nur zu Informationszwecken - keine Anlageempfehlung. Quelle: Yahoo Finance / FMP (zu pruefen).",
      col="5A5A5A", sz=9); ws.merge_cells(f"A{rr+3}:F{rr+3}")

# --------------------------------------------------------------- Finanzdaten
def _financials(ws, main, ccy):
    ws.sheet_view.showGridLines = False; ws.column_dimensions["A"].width = 26
    yrs = main["years"]
    for k in range(len(yrs)): ws.column_dimensions[chr(66 + k)].width = 13
    title(ws, "A1", f"{chr(65+len(yrs))}1", f"FINANZDATEN (Mio. {ccy}, EPS je Aktie)")
    C(ws, "A3", "Geschaeftsjahr", b=True, fill=SEC)
    for c, y in enumerate(yrs, 2): C(ws, f"{chr(64+c)}3", str(y), b=True, fill=SEC, al="right")
    row = 4
    for idx, vals in main["financials"].iterrows():
        fmt = PS if idx == "EPS" else CUR
        C(ws, f"A{row}", idx, bd=True)
        for c, v in enumerate(vals.values, 2): C(ws, f"{chr(64+c)}{row}", _num(v), fmt=fmt, al="right", bd=True)
        row += 1
    chart = BarChart(); chart.title = f"Umsatz (Mio. {ccy})"; chart.height = 7; chart.width = 14; chart.legend = None
    chart.add_data(Reference(ws, min_col=2, max_col=1 + len(yrs), min_row=4, max_row=4), from_rows=True)
    chart.set_categories(Reference(ws, min_col=2, max_col=1 + len(yrs), min_row=3, max_row=3))
    ws.add_chart(chart, f"A{row+1}")

# --------------------------------------------------------------- Kennzahlen
def _ratios(ws, main):
    ws.sheet_view.showGridLines = False; ws.column_dimensions["A"].width = 30
    yrs = main["years"]
    for k in range(len(yrs)): ws.column_dimensions[chr(66 + k)].width = 13
    title(ws, "A1", f"{chr(65+len(yrs))}1", "KENNZAHLEN")
    groups = {"Margen": ["Bruttomarge", "EBITDA-Marge", "EBIT-Marge", "Nettomarge", "FCF-Marge"],
              "Kapitalrendite": ["ROE", "ROA", "ROIC", "ROIC - WACC"],
              "DuPont": ["DuPont Nettomarge", "Kapitalumschlag", "EK-Multiplikator"],
              "Cashflow-Qualitaet": ["Cash Conversion (CFO/NI)", "FCF/NI", "Accruals-Ratio"],
              "Verschuldung & Solvenz": ["Nettoverschuldung/EBITDA", "Zinsdeckung", "Verschuldungsgrad", "Current Ratio"]}
    mult_idx = {"Kapitalumschlag", "EK-Multiplikator", "Nettoverschuldung/EBITDA", "Zinsdeckung", "Verschuldungsgrad", "Current Ratio"}
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
def _valuation(ws, main, models, blended, consensus, wacc, ccy):
    ws.sheet_view.showGridLines = False; ws.column_dimensions["A"].width = 32
    for col in "BCDEF": ws.column_dimensions[col].width = 13
    title(ws, "A1", "F1", f"BEWERTUNG (je Aktie in {ccy})")
    m, d = main["multiples"], main["dcf"]; price = main["latest"]["price"]

    section(ws, 3, "B", "Aktuelle Multiplikatoren")
    row = 4
    for lab, v, fmt in [("KGV", m["pe"], MUL), ("EV/EBIT", m["ev_ebit"], MUL), ("EV/EBITDA", m["ev_ebitda"], MUL),
                        ("EV/FCF", m["ev_fcf"], MUL), ("KBV", m["pb"], MUL), ("FCF-Rendite", m["fcf_yield"], PCT),
                        ("Dividendenrendite", m["div_yield"], PCT)]:
        C(ws, f"A{row}", lab, bd=True); C(ws, f"B{row}", _num(v), fmt=fmt, al="right", bd=True); row += 1

    section(ws, row + 1, "F", "DCF (Free Cashflow to Firm)"); row += 2
    base = d.get("base_fcf", np.nan); g = d.get("growth", np.nan); term = d.get("terminal", np.nan); w = d.get("wacc", wacc)
    for lab, v, fmt in [("Basis-FCF (Mio.)", base / MM if base and not np.isnan(base) else np.nan, CUR),
                        ("Wachstum (J1-5)", g, PCT), ("Terminales Wachstum", term, PCT), ("WACC", w, PCT)]:
        C(ws, f"A{row}", lab, bd=True); C(ws, f"B{row}", _num(v), fmt=fmt, fill=YEL, al="right", bd=True); row += 1
    if base and not np.isnan(base) and not np.isnan(g) and not np.isnan(w):
        C(ws, f"A{row}", "Jahr", b=True, fill=LITE, bd=True)
        for k in range(1, 6): C(ws, f"{chr(65+k)}{row}", f"J{k}", b=True, fill=LITE, al="right", bd=True)
        row += 1
        proj = [base * (1 + g) ** k for k in range(1, 6)]; pv = [proj[k - 1] / (1 + w) ** k for k in range(1, 6)]
        C(ws, f"A{row}", "FCF (Mio.)", bd=True)
        for k in range(5): C(ws, f"{chr(66+k)}{row}", _num(proj[k] / MM), fmt=CUR, al="right", bd=True)
        row += 1
        C(ws, f"A{row}", "Barwert (Mio.)", bd=True)
        for k in range(5): C(ws, f"{chr(66+k)}{row}", _num(pv[k] / MM), fmt=CUR, al="right", bd=True)
        row += 1
    for lab, v, fmt in [("Enterprise Value (Mio.)", d.get("ev_dcf", np.nan) / MM if d.get("ev_dcf") else np.nan, CUR),
                        ("Eigenkapitalwert (Mio.)", d.get("equity_value", np.nan) / MM if d.get("equity_value") else np.nan, CUR),
                        ("Fair Value je Aktie", d.get("fair_value", np.nan), PS),
                        ("Sicherheitsmarge", d.get("mos", np.nan), PCT),
                        ("Reverse-DCF impl. Wachstum", main["reverse_growth"], PCT)]:
        bold = lab.startswith("Fair") or lab.startswith("Sicher")
        C(ws, f"A{row}", lab, b=bold, bd=True)
        C(ws, f"B{row}", _num(v), b=bold, fmt=fmt, al="right", fill=(SEC if bold else None), bd=True); row += 1

    section(ws, row + 1, "C", "Methoden-Gegenueberstellung"); row += 2
    C(ws, f"A{row}", "Methode", b=True, fill=LITE, bd=True); C(ws, f"B{row}", "Fairer Wert", b=True, fill=LITE, al="right", bd=True)
    C(ws, f"C{row}", "Auf-/Abschlag", b=True, fill=LITE, al="right", bd=True); row += 1
    for lab, v in models.items():
        if v is None or np.isnan(v): continue
        up = (v / price - 1) if price and not np.isnan(price) else np.nan
        C(ws, f"A{row}", lab, bd=True); C(ws, f"B{row}", _num(v), fmt=PS, al="right", bd=True)
        c = C(ws, f"C{row}", _num(up), fmt=UP, al="right", bd=True)
        if up is not None and not np.isnan(up): c.fill = GRN if up >= 0 else REDF
        row += 1
    C(ws, f"A{row}", "Median", b=True, bd=True); C(ws, f"B{row}", _num(blended), b=True, fmt=PS, al="right", bd=True)
    C(ws, f"A{row+1}", "Aktueller Kurs", b=True, fill=YEL, bd=True); C(ws, f"B{row+1}", _num(price), b=True, fill=YEL, fmt=PS, al="right", bd=True)

    if consensus and consensus.get("target_mean"):
        section(ws, row + 3, "B", "Analystenkonsens (externe Referenz)")
        rr = row + 4
        for lab, v, fmt in [("Kursziel Mittel", consensus.get("target_mean"), PS),
                            ("Kursziel Hoch", consensus.get("target_high"), PS),
                            ("Kursziel Tief", consensus.get("target_low"), PS),
                            ("Anzahl Analysten", consensus.get("n_analysts"), "0")]:
            C(ws, f"A{rr}", lab, bd=True); C(ws, f"B{rr}", _num(v), fmt=fmt, al="right", bd=True); rr += 1

# --------------------------------------------------------------- Qualitaet (kein Verdikt)
def _quality(ws, main):
    ws.sheet_view.showGridLines = False; ws.column_dimensions["A"].width = 34
    for col in "BCD": ws.column_dimensions[col].width = 13
    title(ws, "A1", "D1", "QUALITAETS-SCORECARD")
    labels = dict(geschaeftsmodell="Geschaeftsmodell*", management="Management*", wachstum="Wachstum",
                  profitabilitaet="Profitabilitaet (ROIC>WACC)", bilanz="Bilanz & Verschuldung",
                  cashflow="Cashflow-Qualitaet", margen="Margen", bewertung="Bewertung (Niveau)")
    C(ws, "A3", "Kriterium", b=True, fill=SEC, bd=True); C(ws, "B3", "Gewicht", b=True, fill=SEC, al="right", bd=True)
    C(ws, "C3", "Score (1-5)", b=True, fill=SEC, al="right", bd=True); C(ws, "D3", "Beitrag", b=True, fill=SEC, al="right", bd=True)
    row = 4
    for k in main["weights"]:
        C(ws, f"A{row}", labels[k], bd=True); C(ws, f"B{row}", main["weights"][k], fmt=PCT, al="right", bd=True)
        C(ws, f"C{row}", main["scores"][k], al="right", bd=True)
        C(ws, f"D{row}", round(main["weights"][k] * main["scores"][k], 2), fmt="0.00", al="right", bd=True); row += 1
    C(ws, f"A{row}", "Qualitaets-Score (gewichtet)", b=True, bd=True)
    C(ws, f"D{row}", round(main["conviction"], 2), b=True, fmt='0.00" / 5"', al="right", bd=True)
    C(ws, f"A{row+2}", "Der Qualitaets-Score bewertet ausschliesslich fundamentale Merkmale (Profitabilitaet, "
                       "Bilanz, Cashflow, Wachstum, Bewertungsniveau). Es handelt sich um eine Einordnung der "
                       "Unternehmensqualitaet, NICHT um eine Kauf- oder Verkaufsempfehlung. * qualitativ gesetzt.",
      col="5A5A5A", sz=9, wrap=True); ws.merge_cells(f"A{row+2}:D{row+4}")

# --------------------------------------------------------------- Peers
def _peers(ws, main, peer_rs):
    ws.sheet_view.showGridLines = False; ws.column_dimensions["A"].width = 26
    cols = [("Qual.-Score", "0.0"), ("Kurs", PS), ("Fair Value", PS), ("Auf-/Abschlag", UP), ("ROIC", PCT),
            ("ROIC-WACC", PCT), ("ND/EBITDA", MUL), ("EBIT-Marge", PCT), ("KGV", MUL), ("EV/EBIT", MUL)]
    for i in range(len(cols)): ws.column_dimensions[chr(66 + i)].width = 12
    title(ws, "A1", f"{chr(65+len(cols))}1", "PEER-VERGLEICH")
    C(ws, "A2", "Titel", b=True, fill=SEC, bd=True)
    for j, (hn, _) in enumerate(cols, 2): C(ws, f"{chr(64+j)}2", hn, b=True, fill=SEC, al="right", bd=True)
    row = 3
    for r in [main] + peer_rs:
        hh, mm = r["headline"], r["multiples"]
        C(ws, f"A{row}", r["ticker"], b=(r is main), bd=True)
        vals = [_num(r["conviction"]), _num(r["latest"]["price"]), _num(hh["fair_value"]), _num(hh["mos"]),
                _num(hh["roic"]), _num(hh["roic_wacc"]), _num(hh["nd_ebitda"]), _num(hh["ebit_margin"]),
                _num(mm["pe"]), _num(mm["ev_ebit"])]
        for j, (v, (_, fmt)) in enumerate(zip(vals, cols), 2):
            C(ws, f"{chr(64+j)}{row}", v, fmt=fmt, al="right", bd=True)
        row += 1

def build_excel(main, peer_rs, models, blended, consensus, wacc):
    ccy = main.get("price_ccy") or main.get("currency") or ""
    wb = Workbook()
    _overview(wb.active, main, models, blended, consensus, ccy); wb.active.title = "Uebersicht"
    _financials(wb.create_sheet("Finanzdaten"), main, ccy)
    _ratios(wb.create_sheet("Kennzahlen"), main)
    _valuation(wb.create_sheet("Bewertung"), main, models, blended, consensus, wacc, ccy)
    _quality(wb.create_sheet("Qualitaet"), main)
    if peer_rs: _peers(wb.create_sheet("Peers"), main, peer_rs)
    bio = BytesIO(); wb.save(bio); return bio.getvalue()

if __name__ == "__main__":
    import equity_engine as eng
    r = eng.compute(eng._synthetic())
    models = {"DCF (FCFF)": r["headline"]["fair_value"], "DDM (Gordon)": 22.9, "EV/EBIT (Peer)": 49.4, "Analysten-Ziel": 42.0}
    cons = dict(target_mean=42.0, target_high=55.0, target_low=30.0, n_analysts=14)
    open("selftest_report.xlsx", "wb").write(build_excel(r, [], models, 33.4, cons, 0.08)); print("Excel ok")
