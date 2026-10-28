#!/usr/bin/env python3
"""Consolidated portfolio workbook: build from a JSON state, or extract the state back out of a workbook.

  python3 consolidate.py extract <book.xlsx> <state.json>   # read the state saved inside a previous book
  python3 consolidate.py build   <state.json> <out.xlsx>    # write the workbook (then run recalc.py on it)
  python3 consolidate.py check   <state.json>               # print per-portfolio row sums vs statement totals

State schema (JSON):
{
  "built": "dd.mm.yyyy", "close_date": "dd.mm.yyyy", "prev_date": "dd.mm.yyyy",
  "closes": {"<security>": [close, prev_close]},          # exchange closes; missing -> statement price
  "screener_urls": {"<security>": "/company/SYMBOL/"},    # Screener.in paths used to fetch closes
  "notes": ["free-text note", ...],                       # shown on Summary under NOTES & FLAGS
  "portfolios": [{
      "key": "Umesh", "tab": "Dr Umesh Choudhry", "name": "Dr Umesh Choudhry", "broker": "ICICI Direct",
      "holdings_date": "...", "price_basis": "...", "basis_label": "Umesh stmt – 09.10", "stale_prices": false,
      "stmt_invested": 0.0, "stmt_mv": 0.0,
      "rows": [["<security>", "<broker>", qty, invested_or_null, stmt_market_value_or_null, "note"]],
      "trades_title": "...", "trades": [["date/time", "BUY/SELL", "<security>", qty, price, value]],
      "base": {"label": "Year 1 base", "amount": 0.0, "cash": [["Cash added", 0.0]]}   # optional, supplied by the user
  }]
}
Portfolio list order = price priority (freshest statement first) for securities with no exchange close.
Rows whose security starts with "Cash" are priced at 1.
"""
import json, sys
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.formatting.rule import FormulaRule
from openpyxl.comments import Comment

FONT, BLUE, GREEN, NAVY = "Arial", "0000FF", "008000", "1F3864"
PALETTE = [("1E8449", "D5F5E3"), ("2874A6", "D6EAF8"), ("27AE60", "E9F7EF"), ("1F618D", "EBF5FB"),
           ("145A32", "ABEBC6"), ("1B4F72", "AED6F1"), ("52BE80", "D4EFDF"), ("5499C7", "D4E6F1")]   # greens and blues
COMB_DARK, COMB_LIGHT = "0B5345", "D4EFDF"
RS = '#,##0;[Red](#,##0);"-"'
PX = '#,##0;[Red](#,##0);"-"'
PCT = '0%;[Red](0%);"-"'
QTY = '#,##0;(#,##0);"-"'
FIRST, LAST = 7, 300
GREEN_NAMES = {"AFCOM Holdings Ltd", "Affle 3i Ltd", "Avenue Supermarts (DMart) Ltd", "Bajaj Finance Ltd",
    "Bajaj Housing Finance Ltd", "Connplex Cinemas Ltd", "Eicher Motors Ltd", "Fineotex Chemical Ltd", "HDFC Bank Ltd",
    "HDFC Gold ETF", "L&T Technology Services Ltd", "LTM Ltd (formerly LTIMindtree)", "Nestle India Ltd",
    "Nippon India ETF Gold BeES", "Persistent Systems Ltd", "Poonawalla Fincorp Ltd", "Rategain Travel Technologies Ltd",
    "Tata Consultancy Services Ltd", "Tata Elxsi Ltd", "Titan Company Ltd", "Trent Ltd", "Vinati Organics Ltd",
    "Waaree Energies Ltd"}   # manual list - change only when the user says so
NAME_GREEN = "00B050"
TAB_COLOURS = ["1E8449", "2874A6", "27AE60", "1F618D", "52BE80", "5DADE2", "145A32", "154360"]   # greens and blues, alternating
STATE_SHEET = "_state"
thin, med = Side(style="thin", color="BFBFBF"), Side(style="medium", color="7F7F7F")
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)


def f(bold=False, color="000000", size=10, italic=False):
    return Font(name=FONT, bold=bold, color=color, size=size, italic=italic)


def fill(h):
    return PatternFill("solid", start_color=h, end_color=h)


AMBER = fill("FCE4B6")


def q(tab):
    return "'" + tab.replace("'", "''") + "'"


def title(ws, text, sub):
    ws["A1"], ws["A2"] = text, sub
    ws["A1"].font, ws["A2"].font = f(True, NAVY, 14), f(False, "595959", 9, True)
    ws.sheet_view.showGridLines = False


def hdr(cell, text, bg=NAVY):
    cell.value = text
    cell.font = f(True, "FFFFFF", 9)
    cell.fill = fill(bg)
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border = BOX


def build(state, out):
    blob = json.dumps(state, ensure_ascii=False)   # saved as-is before any defaults are filled in
    state = json.loads(blob)
    P = state["portfolios"]
    n = len(P)
    CL = state.get("closes", {})
    cdate, pdate = state.get("close_date", ""), state.get("prev_date", "")
    dark = {p["key"]: PALETTE[i % len(PALETTE)][0] for i, p in enumerate(P)}
    light = {p["key"]: PALETTE[i % len(PALETTE)][1] for i, p in enumerate(P)}
    dark["Combined"], light["Combined"] = COMB_DARK, COMB_LIGHT
    for p in P:
        p.setdefault("basis_label", p["name"] + " stmt")
        if p.get("stale_prices") and "stale" not in p["basis_label"]:
            p["basis_label"] += " (stale)"

    stmt_px, all_secs = {}, []
    for p in P:
        for sec, br, qty, inv, mv, note in p["rows"]:
            if mv is not None and qty:
                stmt_px[(p["key"], sec)] = mv / qty
            if sec not in all_secs:
                all_secs.append(sec)

    def ref_px(sec):
        if sec.startswith("Cash"):
            return 1.0
        if sec in CL:
            return CL[sec][0]
        for p in P:
            if (p["key"], sec) in stmt_px:
                return stmt_px[(p["key"], sec)]
        return 0.0

    comb_mv = {s: sum(r[2] * ref_px(s) for p in P for r in p["rows"] if r[0] == s) for s in all_secs}
    sbs_order = sorted(all_secs, key=lambda s: -comb_mv[s])

    wb = Workbook()
    ws_sum = wb.active
    ws_sum.title = "Summary"
    ws_sbs, ws_px = wb.create_sheet("Side by Side"), wb.create_sheet("Prices")
    tabs = {p["key"]: wb.create_sheet(p["tab"]) for p in P}

    # ---------------- portfolio tabs ----------------
    cols = ["#", "Company", "Broker", "No. of Shares", "Total Invested (₹)", "Wtd Avg Cost (₹)", "Ref Price (₹)",
            "Market Value (₹)", "Unrealised P&L (₹)", "P&L %", "Portfolio Weight %", "Note"]
    tot_row, base_rows = {}, {}
    for p in P:
        ws, k = tabs[p["key"]], p["key"]
        title(ws, f"{p['name'].upper()} – {p['broker']} – {p['holdings_date']}",
              f"Holdings as of {p['holdings_date']}. Statement prices: {p['price_basis']}. "
              "Market value uses the common Ref Price from the Prices sheet so all portfolios are valued alike.")
        ws["A3"] = "Blue = figures from the statement (inputs). Black = formulas. Green figures = pulled from another sheet. Green company names = marked holdings. TOTAL is at the bottom."
        ws["A3"].font = f(False, "595959", 9, True)
        for i, h in enumerate(cols, 1):
            hdr(ws.cell(6, i), h, dark[k])
        ws.row_dimensions[6].height = 32
        end = FIRST + len(p["rows"]) - 1
        T = end + 1   # TOTAL row sits at the bottom
        ws[f"B{T}"] = "TOTAL"
        for c in "ABCDEFGHIJKL":
            ws[f"{c}{T}"].fill, ws[f"{c}{T}"].font, ws[f"{c}{T}"].border = fill(light[k]), f(True), Border(left=thin, right=thin, top=med, bottom=med)
        ws[f"E{T}"], ws[f"H{T}"], ws[f"I{T}"] = f"=SUM(E{FIRST}:E{end})", f"=SUM(H{FIRST}:H{end})", f"=H{T}-E{T}"
        ws[f"J{T}"], ws[f"K{T}"] = f'=IF(E{T}>0,I{T}/E{T},"")', f"=SUM(K{FIRST}:K{end})"
        for c, fmt in zip("EHIJK", [RS, RS, RS, PCT, PCT]):
            ws[f"{c}{T}"].number_format = fmt
        tot_row[p["key"]] = T
        bs = p.get("base")
        if bs:
            r0 = T + 2
            ws[f"B{r0}"] = "BASE – last high-water mark (supplied)"
            ws[f"B{r0}"].font = f(True, NAVY)
            r = r0 + 1
            ws[f"B{r}"], ws[f"E{r}"] = f"Base – {bs['label']}", bs["amount"]
            cash_rows = []
            for lab, amt in bs.get("cash", []):
                r += 1
                ws[f"B{r}"], ws[f"E{r}"] = "+ " + lab, amt
                cash_rows.append(r)
            for rr in range(r0 + 1, r + 1):
                ws[f"B{rr}"].font, ws[f"E{rr}"].font, ws[f"E{rr}"].number_format = f(), f(color=BLUE), RS
            r += 1
            ws[f"B{r}"], ws[f"E{r}"] = "TOTAL BASE (base + cash added)", f"=SUM(E{r0+1}:E{r-1})"
            tb = r
            r += 1
            ws[f"B{r}"], ws[f"E{r}"] = "Market value", f"=H{T}"
            r += 1
            ws[f"B{r}"], ws[f"E{r}"] = "Gain / (loss) vs total base", f"=E{r-1}-E{tb}"
            ws[f"F{r}"], ws[f"F{r}"].number_format = f'=IF(E{tb}>0,E{r}/E{tb},"")', PCT
            for rr in range(tb, r + 1):
                ws[f"B{rr}"].font = f(rr in (tb, r))
                ws[f"E{rr}"].font = f(rr in (tb, r))
                ws[f"E{rr}"].number_format = RS
                for cc in "BCDEF":
                    ws[f"{cc}{rr}"].fill = fill(light[k]) if rr in (tb, r) else PatternFill()
            ws[f"F{r}"].font = f(True)
            base_rows[p["key"]] = (tb, r)
        ws.sheet_properties.tabColor = TAB_COLOURS[(list(tabs).index(k)) % len(TAB_COLOURS)]
        rows = sorted(p["rows"], key=lambda r: -(r[2] * ref_px(r[0])))
        for i, (sec, br, qty, inv, mv, note) in enumerate(rows):
            r = FIRST + i
            vals = {"A": i + 1, "B": sec, "C": br, "D": qty, "E": inv,
                    "F": f'=IF(AND(D{r}>0,N(E{r})>0),E{r}/D{r},"")',
                    "G": f'=IFERROR(INDEX(Prices!$C$6:$C${LAST},MATCH(B{r},Prices!$A$6:$A${LAST},0)),"")',
                    "H": f"=IF(ISNUMBER(G{r}),D{r}*G{r},0)", "I": f"=H{r}-N(E{r})",
                    "J": f'=IF(N(E{r})>0,I{r}/E{r},"")', "K": f'=IF($H${T}=0,"",H{r}/$H${T})', "L": note}
            for c, v in vals.items():
                ws[f"{c}{r}"] = v
                ws[f"{c}{r}"].font, ws[f"{c}{r}"].border = f(), BOX
                if inv is None:
                    ws[f"{c}{r}"].fill = fill("F2F2F2")
            for c in "DE":
                ws[f"{c}{r}"].font = f(color=BLUE)
            ws[f"B{r}"].font = f(sec in GREEN_NAMES, NAME_GREEN if sec in GREEN_NAMES else "000000")
            ws[f"G{r}"].font, ws[f"L{r}"].font = f(color=GREEN), f(False, "7F7F7F", 9, True)
            ws[f"D{r}"].number_format = QTY
            for c in "EHI":
                ws[f"{c}{r}"].number_format = RS
            for c in "FG":
                ws[f"{c}{r}"].number_format = PX
            for c in "JK":
                ws[f"{c}{r}"].number_format = PCT
        for c, w in zip("ABCDEFGHIJKL", [4, 38, 17, 10, 14, 12, 11, 14, 13, 8, 9, 40]):
            ws.column_dimensions[c].width = w
        ws.freeze_panes = "C7"
        ws.conditional_formatting.add(f"G{FIRST}:G{LAST}", FormulaRule(formula=[
            f'ISNUMBER(SEARCH("stale",INDEX(Prices!$D$6:$D${LAST},MATCH($B{FIRST},Prices!$A$6:$A${LAST},0))))'], fill=AMBER))
        if p.get("trades"):
            ws["Q5"] = p.get("trades_title", "Trades (already blended into Wtd Avg Cost)")
            ws["Q5"].font = f(True, NAVY)
            for i, h in enumerate(["Date / Time", "Type", "Instrument", "Qty", "Fill Price (₹)", "Value (₹)"]):
                hdr(ws.cell(6, 17 + i), h, dark[k])
            for j, t in enumerate(p["trades"]):
                for i, v in enumerate(t):
                    c = ws.cell(7 + j, 17 + i, v)
                    c.font, c.border = f(color=BLUE), BOX
                ws.cell(7 + j, 21).number_format = PX
                ws.cell(7 + j, 22).number_format = RS
            r = 7 + len(p["trades"])
            ws.cell(r, 19, "Net bought (buys − sells)").font = f(True)
            ws.cell(r, 22, f'=SUMIFS(V7:V{r-1},R7:R{r-1},"BUY")-SUMIFS(V7:V{r-1},R7:R{r-1},"SELL")').font = f(True)
            ws.cell(r, 22).number_format = RS
            for c, w in zip("QRSTUV", [19, 6, 24, 6, 12, 11]):
                ws.column_dimensions[c].width = w

    # ---------------- prices ----------------
    ws = ws_px
    title(ws, "COMMON PRICE SHEET",
          "One price per security, used to value every portfolio. Without an exchange close, the statement price is used, "
          "freshest statement first: " + " → ".join(p["basis_label"] for p in P) + ".")
    ws["A3"] = (f"Column B holds NSE/BSE prices as at {cdate} (Screener.in). TO UPDATE: overwrite the yellow column B and the "
                "source text beside it — every portfolio and the combined view revalue. Blank B = statement price.")
    ws["A3"].font = f(True, "C00000", 9)
    stc = {p["key"]: L(5 + i) for i, p in enumerate(P)}
    SRC, PRV, CHG = L(5 + n), L(6 + n), L(7 + n)
    heads = (["Security", "Override Price (₹)\n← type here", "Ref Price (₹)\n(used everywhere)", "Price Basis"]
             + [p["name"] + " stmt (₹)" for p in P] + ["Override source", f"Prev close {pdate[:5]} (₹)", "1-day chg %"])
    for i, h in enumerate(heads, 1):
        hdr(ws.cell(5, i), h)
    ws.row_dimensions[5].height = 32
    for i, sec in enumerate(sorted(all_secs, key=str.lower)):
        r = 6 + i
        ws[f"A{r}"], ws[f"A{r}"].font = sec, f()
        ws[f"B{r}"].fill, ws[f"B{r}"].font = fill("FFFF00"), f(color=BLUE)
        if sec.startswith("Cash"):
            ws[f"B{r}"], ws[f"{SRC}{r}"] = 1, "Cash at par (₹1)"
        elif sec in CL:
            ws[f"B{r}"], ws[f"{SRC}{r}"] = CL[sec][0], f"Close {cdate} (Screener.in)"
            if len(CL[sec]) > 1 and CL[sec][1] is not None:
                ws[f"{PRV}{r}"] = CL[sec][1]
        else:
            ws[f"{SRC}{r}"] = "No exchange close found – statement price used"
        ws[f"{SRC}{r}"].font, ws[f"{PRV}{r}"].font = f(False, BLUE, 9), f(color=BLUE)
        ws[f"{PRV}{r}"].number_format = PX
        ws[f"{CHG}{r}"] = f'=IF(AND(ISNUMBER(B{r}),ISNUMBER({PRV}{r}),N({PRV}{r})>0),B{r}/{PRV}{r}-1,"")'
        ws[f"{CHG}{r}"].number_format, ws[f"{CHG}{r}"].font = PCT, f()
        for p in P:
            t, c = q(p["tab"]), stc[p["key"]]
            ws[f"{c}{r}"] = stmt_px.get((p["key"], sec), "")
            ws[f"{c}{r}"].font, ws[f"{c}{r}"].number_format = f(color=BLUE), PX
        cpx, cbs = '""', '"No price on any statement"'
        for p in reversed(P):
            c = stc[p["key"]]
            cpx = f"IF(ISNUMBER({c}{r}),{c}{r},{cpx})"
            cbs = f'IF(ISNUMBER({c}{r}),"{p["basis_label"]}",{cbs})'
        ws[f"C{r}"] = f"=IF(ISNUMBER(B{r}),B{r},{cpx})"
        ws[f"D{r}"] = f'=IF(ISNUMBER(B{r}),IF({SRC}{r}<>"",{SRC}{r},"Manual override"),{cbs})'
        ws[f"C{r}"].font, ws[f"D{r}"].font = f(True), f(False, "595959", 9)
        for c in "BC":
            ws[f"{c}{r}"].number_format = PX
        for ci in range(1, 8 + n):
            ws.cell(r, ci).border = BOX
    pend = 5 + len(all_secs)
    ws.conditional_formatting.add(f"C6:D{pend}", FormulaRule(formula=['ISNUMBER(SEARCH("stale",$D6))'], fill=AMBER))
    ws.conditional_formatting.add(f"C6:D{pend}", FormulaRule(formula=['ISNUMBER(SEARCH("No price",$D6))'], fill=fill("F2F2F2")))
    for ci, w in enumerate([44, 13, 13, 30] + [12] * n + [30, 12, 9], 1):
        ws.column_dimensions[L(ci)].width = w
    ws.freeze_panes = "B6"
    ws["B5"].comment = Comment("Leave blank to use statement price. Type a number to override for all portfolios.", "Claude")

    # ---------------- side by side ----------------
    ws = ws_sbs
    title(ws, "SIDE-BY-SIDE HOLDINGS – COMBINED AND BY PORTFOLIO",
          "All values at the common Ref Price (Prices sheet). Sorted by combined market value at build; '-' = not held. "
          "Amber price = only a stale statement price available.")
    for i, h in enumerate(["#", "Security", "Ref Price (₹)", "Price Basis", "Held in\n(# a/cs)"], 1):
        hdr(ws.cell(5, i), h)
        ws.cell(4, i).fill = fill(NAVY)
    groups = [("Combined", f"COMBINED (ALL {n})", None)] + [(p["key"], p["name"].upper() + " – " + p["broker"], p["tab"]) for p in P]
    sub = ["Qty", "Wtd Avg Cost (₹)", "Invested (₹)", "Mkt Value (₹)", "P&L (₹)", "P&L %", "Weight %"]
    gcol, col = {}, 6
    for key, label, tab in groups:
        gcol[key] = col
        ws.merge_cells(start_row=4, start_column=col, end_row=4, end_column=col + 6)
        hdr(ws.cell(4, col), label, dark[key])
        for j, s in enumerate(sub):
            hdr(ws.cell(5, col + j), s, dark[key])
        col += 7
    lastcol = col - 1
    ws.row_dimensions[4].height = ws.row_dimensions[5].height = 30
    end = FIRST + len(sbs_order) - 1
    for i, sec in enumerate(sbs_order):
        r = FIRST + i
        ws[f"A{r}"], ws[f"B{r}"] = i + 1, sec
        ws[f"C{r}"] = f'=IFERROR(INDEX(Prices!$C$6:$C${LAST},MATCH($B{r},Prices!$A$6:$A${LAST},0)),"")'
        ws[f"D{r}"] = f'=IFERROR(INDEX(Prices!$D$6:$D${LAST},MATCH($B{r},Prices!$A$6:$A${LAST},0)),"")'
        ws[f"A{r}"].font, ws[f"B{r}"].font, ws[f"C{r}"].font, ws[f"D{r}"].font = f(), f(True), f(color=GREEN), f(False, GREEN, 8)
        ws[f"C{r}"].number_format = PX
        qcells = []
        for key, label, tab in groups[1:]:
            c0, t = gcol[key], q(tab)
            ws[f"{L(c0)}{r}"] = f"=SUMIFS({t}!$D${FIRST}:$D${LAST},{t}!$B${FIRST}:$B${LAST},$B{r})"
            ws[f"{L(c0+2)}{r}"] = f"=SUMIFS({t}!$E${FIRST}:$E${LAST},{t}!$B${FIRST}:$B${LAST},$B{r})"
            ws[f"{L(c0+3)}{r}"] = f"=SUMIFS({t}!$H${FIRST}:$H${LAST},{t}!$B${FIRST}:$B${LAST},$B{r})"
            qcells.append(f"{L(c0)}{r}")
        c0 = gcol["Combined"]
        ws[f"{L(c0)}{r}"] = "=" + "+".join(qcells)
        ws[f"{L(c0+2)}{r}"] = "=" + "+".join(f"{L(gcol[k]+2)}{r}" for k, _, _ in groups[1:])
        ws[f"{L(c0+3)}{r}"] = "=" + "+".join(f"{L(gcol[k]+3)}{r}" for k, _, _ in groups[1:])
        ws[f"E{r}"] = "=" + "+".join(f"({c}>0)" for c in qcells)
        ws[f"E{r}"].font, ws[f"E{r}"].alignment = f(), Alignment(horizontal="center")
        for key, label, tab in groups:
            c0 = gcol[key]
            Q, A, I, M, Pl, PC, W = [L(c0 + j) for j in range(7)]
            ws[f"{A}{r}"] = f'=IF(AND({Q}{r}>0,{I}{r}>0),{I}{r}/{Q}{r},"")'
            ws[f"{Pl}{r}"] = f"={M}{r}-{I}{r}"
            ws[f"{PC}{r}"] = f'=IF({I}{r}>0,{Pl}{r}/{I}{r},"")'
            ws[f"{W}{r}"] = f'=IF({M}$6=0,"",{M}{r}/{M}$6)'
            for j, fmt in enumerate([QTY, PX, RS, RS, RS, PCT, PCT]):
                cell = ws[f"{L(c0 + j)}{r}"]
                cell.number_format = fmt
                cell.font = f(key == "Combined" and j in (3, 6))
                if key == "Combined":
                    cell.fill = fill(COMB_LIGHT)
        for ci in range(1, lastcol + 1):
            ws.cell(r, ci).border = BOX
    ws["B6"] = "TOTAL"
    for ci in range(1, lastcol + 1):
        c = ws.cell(6, ci)
        c.font, c.fill, c.border = f(True), fill("D9D9D9"), Border(left=thin, right=thin, top=med, bottom=med)
    ws["E6"] = f'=COUNTIFS(E{FIRST}:E{end},">0",B{FIRST}:B{end},"<>Cash*")'
    ws["E6"].alignment = Alignment(horizontal="center")
    ws["E6"].comment = Comment("Distinct securities held across all portfolios (cash excluded)", "Claude")
    for key, label, tab in groups:
        c0 = gcol[key]
        Q, A, I, M, Pl, PC, W = [L(c0 + j) for j in range(7)]
        ws[f"{Q}6"] = f'=COUNTIFS({Q}{FIRST}:{Q}{end},">0",$B{FIRST}:$B{end},"<>Cash*")'
        ws[f"{I}6"], ws[f"{M}6"] = f"=SUM({I}{FIRST}:{I}{end})", f"=SUM({M}{FIRST}:{M}{end})"
        ws[f"{Pl}6"], ws[f"{PC}6"], ws[f"{W}6"] = f"={M}6-{I}6", f'=IF({I}6>0,{Pl}6/{I}6,"")', f"=SUM({W}{FIRST}:{W}{end})"
        for j, fmt in enumerate(['0 "stocks"', PX, RS, RS, RS, PCT, PCT]):
            ws[f"{L(c0 + j)}6"].number_format = fmt
    ws.conditional_formatting.add(f"C{FIRST}:C{end}", FormulaRule(formula=[f'ISNUMBER(SEARCH("stale",$D{FIRST}))'], fill=AMBER))
    for c, w in {"A": 4, "B": 36, "C": 11, "D": 20, "E": 8}.items():
        ws.column_dimensions[c].width = w
    for key, *_ in groups:
        for j, w in enumerate([8, 10, 12, 12, 11, 7.5, 7.5]):
            ws.column_dimensions[L(gcol[key] + j)].width = w
    ws.freeze_panes = "C7"

    # ---------------- summary ----------------
    ws = ws_sum
    title(ws, "CONSOLIDATED PORTFOLIO SUMMARY – ALL MANAGED ACCOUNTS",
          f"Built {state.get('built', '')} from {n} statements. All market values at the Screener.in price of {cdate} (Prices sheet); "
          "statement-reported values kept alongside for reconciliation.")
    sh = ["Portfolio", "Broker", "Holdings as of", "Statement price basis", "# Holdings", "Invested (₹)", "Market Value (₹)",
          "Unrealised P&L (₹)", "P&L %", "Share of Combined MV",
          "Total Base incl. cash added (₹)", "Gain vs Base (₹)", "Gain vs Base %"]
    for i, h in enumerate(sh, 1):
        hdr(ws.cell(5, i), h)
    ws.row_dimensions[5].height = 32
    cr = 6 + n
    for i, p in enumerate(P):
        r, t = 6 + i, q(p["tab"])
        vals = [p["name"], p["broker"], p["holdings_date"], p["price_basis"],
                f'=COUNTIFS({t}!$D${FIRST}:$D${LAST},">0",{t}!$B${FIRST}:$B${LAST},"<>Cash*")', f"={t}!E{tot_row[p['key']]}", f"={t}!H{tot_row[p['key']]}",
                f"=G{r}-F{r}", f'=IF(F{r}>0,H{r}/F{r},"")', f'=IF($G${cr}=0,"",G{r}/$G${cr})']
        if p["key"] in base_rows:
            vals += [f"={t}!E{base_rows[p['key']][0]}", f"=G{r}-K{r}", f'=IF(K{r}>0,L{r}/K{r},"")']
        for j, v in enumerate(vals, 1):
            c = ws.cell(r, j, v)
            c.font, c.border = f(j == 1), BOX
            if j == 1:
                c.fill = fill(light[p["key"]])
        for c in "FG":
            ws[f"{c}{r}"].font = f(color=GREEN)
    ws[f"A{cr}"], ws[f"B{cr}"] = f"COMBINED (ALL {n})", " + ".join(dict.fromkeys(p["broker"] for p in P))
    ws[f"E{cr}"] = "='Side by Side'!E6"
    for c in "FGHKL":
        ws[f"{c}{cr}"] = f"=SUM({c}6:{c}{cr-1})"
    ws[f"M{cr}"] = f'=IF(K{cr}>0,L{cr}/K{cr},"")'
    ws[f"I{cr}"], ws[f"J{cr}"] = f'=IF(F{cr}>0,H{cr}/F{cr},"")', f"=SUM(J6:J{cr-1})"
    for j in range(1, 14):
        c = ws.cell(cr, j)
        c.font, c.fill, c.border = f(True, "FFFFFF"), fill(COMB_DARK), BOX
    for rr in range(6, cr + 1):
        for c in "FGHKL":
            ws[f"{c}{rr}"].number_format = RS
        for c in "IJM":
            ws[f"{c}{rr}"].number_format = PCT
        ws[f"E{rr}"].alignment = Alignment(horizontal="center")

    tr = n + 9
    ws[f"A{tr}"] = "TOP 10 COMBINED EXPOSURES – weight in each portfolio"
    ws[f"A{tr}"].font = f(True, NAVY, 11)
    th = ["Rank", "Security", "Combined MV (₹)", "Combined Weight"] + [p["name"] for p in P] + ["Held in (# a/cs)"]
    for i, h in enumerate(th, 1):
        bg = COMB_DARK if i <= 4 else (dark[P[i - 5]["key"]] if i <= 4 + n else NAVY)
        hdr(ws.cell(tr + 1, i), h, bg)
    ws.row_dimensions[tr + 1].height = 30
    sbs, cm, cw = "'Side by Side'", L(gcol["Combined"] + 3), L(gcol["Combined"] + 6)
    for k in range(1, min(10, len(sbs_order)) + 1):
        r = tr + 1 + k
        ws[f"A{r}"] = k
        ws[f"C{r}"] = f"=LARGE({sbs}!${cm}${FIRST}:${cm}${end},A{r})"
        match = f"MATCH(C{r},{sbs}!${cm}${FIRST}:${cm}${end},0)"
        ws[f"B{r}"] = f"=INDEX({sbs}!$B${FIRST}:$B${end},{match})"
        ws[f"D{r}"] = f"=INDEX({sbs}!${cw}${FIRST}:${cw}${end},{match})"
        for i, p in enumerate(P):
            ws.cell(r, 5 + i, f"=INDEX({sbs}!${L(gcol[p['key']] + 6)}${FIRST}:${L(gcol[p['key']] + 6)}${end},{match})").number_format = PCT
        ws.cell(r, 5 + n, f"=INDEX({sbs}!$E${FIRST}:$E${end},{match})").alignment = Alignment(horizontal="center")
        ws[f"C{r}"].number_format, ws[f"D{r}"].number_format = RS, PCT
        ws[f"A{r}"].alignment = Alignment(horizontal="center")
        for j in range(1, 6 + n):
            ws.cell(r, j).border = BOX
            ws.cell(r, j).font = f(j == 2, GREEN if j > 1 else "000000")

    notes = ["NOTES & FLAGS",
             f"• Pricing: every holding is valued at the Screener.in price of {cdate}, one price per "
             "security across all accounts.",
             f'="• Securities on an exchange close: "&COUNTIF(Prices!${SRC}$6:${SRC}${LAST},"Close*")&"   |   on a statement price: "'
             f'&COUNTIF(Prices!${SRC}$6:${SRC}${LAST},"No exchange*")&"   |   held in all {n} accounts: "&COUNTIF({sbs}!$E${FIRST}:$E${end},{n})']
    notes += ["• " + t for t in state.get("notes", [])]
    for i, t in enumerate(notes):
        c = ws.cell(tr + 14 + i, 1, t)
        c.font = f(i == 0, NAVY if i == 0 else "404040", 10 if i == 0 else 9)
    for ci, w in enumerate([24, 34, 40, 26, 11, 14, 14, 14, 9, 12, 16, 14, 11], 1):
        ws.column_dimensions[L(ci)].width = max(w, 14 if ci > 4 + n else 0)
    ws.freeze_panes = "B6"

    ws_sum.sheet_properties.tabColor, ws_sbs.sheet_properties.tabColor, ws_px.sheet_properties.tabColor = "0B5345", "117864", "1A5276"

    # ---------------- hidden state (lets the next run start from this book) ----------------
    st = wb.create_sheet(STATE_SHEET)
    for i in range(0, len(blob), 30000):
        st.cell(1 + i // 30000, 1, blob[i:i + 30000])
    st.sheet_state = "hidden"

    for w in wb.worksheets:
        w.page_setup.orientation, w.page_setup.fitToWidth, w.page_setup.fitToHeight = "landscape", 1, 0
        w.sheet_properties.pageSetUpPr.fitToPage = True
    wb.save(out)
    print(f"saved {out}: {n} portfolios, {len(all_secs)} securities, {len(CL)} exchange closes")


def extract(book, out):
    wb = load_workbook(book)
    if STATE_SHEET in wb.sheetnames:
        ws = wb[STATE_SHEET]
        blob = "".join(ws.cell(r, 1).value or "" for r in range(1, ws.max_row + 1))
        state = json.loads(blob)
    else:  # older book without _state: rebuild state from the visible sheets
        state = _extract_visible(wb)
    json.dump(state, open(out, "w"), ensure_ascii=False, indent=1)
    print(f"state -> {out}: {len(state['portfolios'])} portfolios, {len(state.get('closes', {}))} closes, "
          f"close_date {state.get('close_date')}")


def _extract_visible(wb):
    import re
    summ, prices = wb["Summary"], wb["Prices"]
    hdrs = {prices.cell(5, c).value: c for c in range(1, prices.max_column + 1)}
    src = hdrs.get("Override source")
    prv = next((c for h, c in hdrs.items() if h and str(h).startswith("Prev close")), None)
    closes, cdate = {}, ""
    for r in range(6, prices.max_row + 1):
        sec, px, s = prices.cell(r, 1).value, prices.cell(r, 2).value, prices.cell(r, src).value if src else None
        if sec and isinstance(px, (int, float)) and s and str(s).startswith("Close"):
            closes[sec] = [px, prices.cell(r, prv).value if prv else None]
            cdate = cdate or re.search(r"\d{2}\.\d{2}\.\d{4}", s).group(0)
    stmt = {summ.cell(r, 1).value: (summ.cell(r, 3).value, summ.cell(r, 4).value, None, None)
            for r in range(6, 30) if summ.cell(r, 1).value}
    P = []
    for name in wb.sheetnames:
        if name in ("Summary", "Side by Side", "Prices", STATE_SHEET):
            continue
        ws = wb[name]
        disp, broker = (ws["A1"].value or name).split(" – ", 1) if " – " in (ws["A1"].value or "") else (name, "")
        rows, r = [], FIRST
        while ws.cell(r, 2).value and ws.cell(r, 2).value != "TOTAL":
            rows.append([ws.cell(r, 2).value, ws.cell(r, 3).value, ws.cell(r, 4).value, ws.cell(r, 5).value,
                         None, ws.cell(r, 12).value or ""])
            r += 1
        trades, r = [], 7
        while ws.cell(r, 17).value and ws.cell(r, 18).value:
            trades.append([ws.cell(r, c).value for c in range(17, 23)])
            r += 1
        sname = next((k for k in stmt if k and k.upper() == disp.upper()), name)
        hd, pb, si, sm = stmt.get(sname, ("", "", None, None))
        P.append({"key": name.split()[0], "tab": name, "name": sname, "broker": broker, "holdings_date": hd,
                  "price_basis": pb, "basis_label": sname + " stmt", "stale_prices": "older" in str(pb).lower() or "not printed" in str(pb).lower(),
                  "stmt_invested": si, "stmt_mv": sm, "rows": rows,
                  "trades_title": ws["Q5"].value or "", "trades": trades})
    return {"built": "", "close_date": cdate, "prev_date": "", "closes": closes, "screener_urls": {},
            "notes": [], "portfolios": P}


def check(path):
    s = json.load(open(path))
    for p in s["portfolios"]:
        inv = sum(r[3] or 0 for r in p["rows"])
        mv = sum(r[4] or 0 for r in p["rows"])
        print(f"{p['name'][:28]:28s} rows {len(p['rows']):3d}  invested {inv:>14,.2f} vs stmt {p['stmt_invested'] or 0:>14,.2f}"
              f"  (diff {inv - (p['stmt_invested'] or 0):,.2f})   mv {mv:>14,.2f} vs stmt {p['stmt_mv'] or 0:>14,.2f}"
              f"  (diff {mv - (p['stmt_mv'] or 0):,.2f})")
    secs = {r[0] for p in s["portfolios"] for r in p["rows"]}
    missing = sorted(x for x in secs if x not in s.get("closes", {}) and not x.startswith("Cash"))
    print("no exchange close:", missing)
    no_url = sorted(x for x in secs if x not in s.get("screener_urls", {}) and not x.startswith("Cash"))
    print("no Screener url:", no_url)


if __name__ == "__main__":
    cmd = sys.argv[1]
    if cmd == "build":
        build(json.load(open(sys.argv[2])), sys.argv[3])
    elif cmd == "extract":
        extract(sys.argv[2], sys.argv[3])
    elif cmd == "check":
        check(sys.argv[2])
