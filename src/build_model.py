"""Build the public Excel budget model from synthetic line items.

Run after clean_budget.py:
    python src/build_model.py
Writes output/Production_Cost_Review_Demo.xlsx (all totals are live formulas).
"""

from __future__ import annotations

from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parent.parent
FONT = "Arial"
BLUE, GREEN, BLACK, GREY = "0000FF", "008000", "000000", "6B7280"
F_IN = Font(name=FONT, color=BLUE)
F_LINK = Font(name=FONT, color=GREEN)
F = Font(name=FONT, color=BLACK)
F_B = Font(name=FONT, bold=True)
F_H = Font(name=FONT, bold=True, color="FFFFFF")
F_T = Font(name=FONT, bold=True, size=14)
F_NOTE = Font(name=FONT, italic=True, color=GREY, size=9)
FILL_H = PatternFill("solid", fgColor="1F3A5F")
FILL_KEY = PatternFill("solid", fgColor="FFFF00")
FILL_SUB = PatternFill("solid", fgColor="F3F4F6")
THIN = Side(style="thin", color="D1D5DB")
TOP = Border(top=Side(style="thin", color=BLACK))
USD = '$#,##0;($#,##0);"-"'
USD2 = '$#,##0.00;($#,##0.00);"-"'
PCT = '0.0%;(0.0%);"-"'
PTS = '+0.0" pts";-0.0" pts";"-"'

# Cost driver per item, used by the Scenario sheet. Assumptions, documented on the README sheet.
FIXED_ITEMS = {"Director", "Props", "Wardrobe purchases", "Art department purchases",
               "Art department purchases (hardware)", "Production insurance", "Auto insurance",
               "Equipment rental insurance", "Tools", "Sound equipment", "Heaters & hand warmers",
               "Trash cans", "Fitting room"}
EPISODE_ITEMS = {"Editing (quoted as CNY 40,000)", "Trailer", "Poster"}

CATEGORIES = ["Below-the-line crew", "Locations", "Above-the-line crew", "Cast", "Gear & equipment",
              "Insurance", "Art, props & wardrobe", "Post", "Catering & crafty", "Transport", "Set supplies"]
# Benchmark split: Real Reel line-producer breakdown, as reported by Filmustage (2026)
BENCHMARK = {"Cast": 0.22, "Locations": 0.20, "Above-the-line crew": 0.17, "Below-the-line crew": 0.18,
             "Gear & equipment": 0.12, "Post": 0.06, "Contingency": 0.05}


def header(ws, row, labels, widths=None):
    for c, label in enumerate(labels, start=1):
        cell = ws.cell(row, c, label)
        cell.font, cell.fill = F_H, FILL_H
        cell.alignment = Alignment(vertical="center", wrap_text=True)
    ws.row_dimensions[row].height = 30
    if widths:
        for c, w in enumerate(widths, start=1):
            ws.column_dimensions[get_column_letter(c)].width = w


def driver(item: str, section: str) -> str:
    if item in FIXED_ITEMS:
        return "Fixed"
    if item in EPISODE_ITEMS:
        return "Episodes"
    return "Shoot days"


def flag(v) -> str:
    if pd.isna(v):
        return "Not recorded"
    return "Yes" if bool(v) else "No"


def build() -> Path:
    items = pd.read_csv(ROOT / "data" / "sample_line_items.csv")
    dq = pd.read_csv(ROOT / "data" / "data_quality.csv")
    # Lines with no amount in the source: label them so a reader knows why they are $0
    items.loc[items["item"].isin(["Trailer", "Poster"]), "item"] += " (made in-house, not costed)"
    items.loc[items["stated_total"].isna() & items["section"].eq("Catering & crafty"), "item"] += " (amount not recorded)"
    EPISODE_ITEMS.update({"Trailer (made in-house, not costed)", "Poster (made in-house, not costed)"})

    wb = Workbook()

    # ------------------------------------------------------------------ README
    ws = wb.active
    ws.title = "README"
    ws.column_dimensions["A"].width = 110
    lines = [
        ("Vertical short-drama production: budget model", F_T),
        ("Public portfolio edition built from synthetic cost data modelled on a multi-episode vertical-drama workflow.", F),
        ("The original production workbook, client identities, personal information and real financial values are not included.", F),
        ("", F),
        ("Sheets", F_B),
        ("Summary: totals by cost category, comparison with an industry cost split, unit costs.", F),
        ("Scenario: what a next series would cost under different episode counts, shoot days and contingency.", F),
        ("People & Paperwork: cast tiers, and how much people spend has invoices, payments and contracts recorded.", F),
        ("Line Items: synthetic cost lines used by all other sheets.", F),
        ("Data Quality: illustrative controls and exceptions for the demo dataset.", F),
        ("Benchmarks: external figures used, with sources.", F),
        ("", F),
        ("Colour key", F_B),
        ("Blue text = input you can change (actual amounts, scenario levers, benchmark values).", F_IN),
        ("Black text = formula.   Green text = link to another sheet.", F),
        ("Yellow fill = key scenario levers: change these to test a next series.", F),
        ("", F),
        ("Key assumptions", F_B),
        ("1. Cost drivers (Line Items, column M): day-rate crew, cast, locations, catering, equipment rental and fuel scale with shoot days;", F),
        ("   editing scales with episode count; director fee, insurance, art/props/wardrobe purchases and small purchases are fixed per series.", F),
        ("2. Scaling is linear: halving shoot days halves day-driven costs. Real productions have minimum bookings and overtime, so treat results as planning estimates.", F),
        ("3. Trailer and poster were made in-house and carry no cost. One catering receipt (Mar 23) has no recorded amount.", F),
        ("4. Editing was quoted in Chinese yuan (CNY 40,000) and recorded as $5,500.", F),
        ("5. The benchmark split does not include insurance, catering, art or transport, so shares are compared only where categories match.", F),
    ]
    for r, (text, font) in enumerate(lines, start=1):
        c = ws.cell(r, 1, text)
        c.font = font
    ws.cell(16, 1).fill = FILL_KEY

    # ------------------------------------------------------------------ Line Items
    li = wb.create_sheet("Line Items")
    cols = ["ID", "Section", "Category", "Role group", "Item", "Date", "Days", "Day rate ($)", "Prep",
            "Days x rate ($)", "Amount ($)", "Amount less days x rate ($)", "Cost driver", "Invoice", "Paid", "Contract"]
    header(li, 1, cols, [6, 18, 22, 13, 40, 11, 7, 11, 7, 13, 12, 14, 12, 13, 13, 13])
    n = len(items)
    for i, row in enumerate(items.itertuples(index=False), start=2):
        vals = [i - 1, row.section, row.category, row.role_group if pd.notna(row.role_group) else "", row.item,
                pd.to_datetime(row.date).date() if pd.notna(row.date) else None,
                row.days if pd.notna(row.days) else None, row.rate if pd.notna(row.rate) else None,
                row.prep if pd.notna(row.prep) else None]
        for c, v in enumerate(vals, start=1):
            li.cell(i, c, v).font = F
        li.cell(i, 6).number_format = "d mmm yyyy"
        li.cell(i, 8).number_format = USD
        li.cell(i, 10, f'=IF(AND(ISNUMBER(G{i}),ISNUMBER(H{i})),G{i}*H{i},"")').number_format = USD
        amt = li.cell(i, 11, float(row.amount))
        amt.font, amt.number_format = F_IN, USD2
        li.cell(i, 12, f'=IF(J{i}="","",K{i}-J{i})').number_format = USD2
        li.cell(i, 13, driver(row.item, row.section)).font = F_IN
        people = row.section in ("Crew", "Cast")
        for c, v in zip((14, 15, 16), (row.invoice, row.paid, row.contract)):
            if people or c == 15:
                li.cell(i, c, flag(v)).font = F
    last = n + 1
    tr = last + 1
    li.cell(tr, 5, "Total").font = F_B
    li.cell(tr, 11, f"=SUM(K2:K{last})").number_format = USD2
    li.cell(tr, 11).font, li.cell(tr, 11).border = F_B, TOP
    li.freeze_panes = "F2"
    li.auto_filter.ref = f"A1:P{last}"
    rng = lambda col: f"'Line Items'!${col}$2:${col}${last}"  # noqa: E731
    AMT, CAT, SEC, ROLE, DRV, INV, PAID, CON = (rng(c) for c in "KCBDMNOP")

    # ------------------------------------------------------------------ Summary
    s = wb.create_sheet("Summary", 1)
    s["A1"], s["A1"].font = "Where the money went", F_T
    s["A2"], s["A2"].font = "Actual spend by category vs a typical LA vertical-drama cost split", F_NOTE
    header(s, 4, ["Category", "Actual ($)", "Share of total", "Benchmark share", "Difference (pts)", "Benchmark $ at this total"],
           [26, 14, 14, 16, 16, 22])
    r0 = 5
    for k, cat in enumerate(CATEGORIES + ["Contingency"]):
        r = r0 + k
        s.cell(r, 1, cat).font = F
        if cat == "Contingency":
            s.cell(r, 2, 0).font = F_IN
            s.cell(r, 2).comment = Comment("No contingency line in the original budget.", "Model")
        else:
            s.cell(r, 2, f'=SUMIFS({AMT},{CAT},A{r})').font = F_LINK
        s.cell(r, 2).number_format = USD
        s.cell(r, 3, f"=IF($B${r0 + len(CATEGORIES) + 1}=0,0,B{r}/$B${r0 + len(CATEGORIES) + 1})").number_format = PCT
        if cat in BENCHMARK:
            b = s.cell(r, 4, BENCHMARK[cat])
            b.font, b.number_format = F_IN, PCT
            s.cell(r, 5, f"=(C{r}-D{r})*100").number_format = PTS
            s.cell(r, 6, f"=D{r}*$B${r0 + len(CATEGORIES) + 1}").number_format = USD
        else:
            s.cell(r, 4, "not in benchmark").font = F_NOTE
    tot = r0 + len(CATEGORIES) + 1
    s.cell(tot, 1, "Total").font = F_B
    s.cell(tot, 2, f"=SUM(B{r0}:B{tot - 1})")
    s.cell(tot, 3, f"=SUM(C{r0}:C{tot - 1})")
    s.cell(tot, 4, f"=SUM(D{r0}:D{tot - 1})")
    for c, fmt in ((2, USD), (3, PCT), (4, PCT)):
        s.cell(tot, c).number_format, s.cell(tot, c).font, s.cell(tot, c).border = fmt, F_B, TOP
    s.cell(tot + 1, 1, "Check: equals Line Items total").font = F_NOTE
    s.cell(tot + 1, 2, f"=IF(ABS(B{tot}-'Line Items'!K{tr})<0.01,\"OK\",\"MISMATCH\")").font = F_NOTE
    s.cell(tot + 2, 1, "Benchmark: Real Reel line-producer split, reported by Filmustage (2026). See Benchmarks sheet.").font = F_NOTE

    # Unit costs
    u = tot + 4
    s.cell(u, 1, "Unit costs").font = F_B
    s.cell(u + 1, 1, "Episodes"); s.cell(u + 1, 2, "=Scenario!B6").font = F_LINK
    s.cell(u + 2, 1, "Shoot days"); s.cell(u + 2, 2, "=Scenario!B7").font = F_LINK
    s.cell(u + 3, 1, "Cost per episode ($)"); s.cell(u + 3, 2, f"=B{tot}/B{u + 1}")
    s.cell(u + 4, 1, "Cost per shoot day ($)"); s.cell(u + 4, 2, f"=B{tot}/B{u + 2}")
    s.cell(u + 5, 1, "Episodes per shoot day"); s.cell(u + 5, 2, f"=B{u + 1}/B{u + 2}")
    s.cell(u + 6, 1, "Cost per finished minute, 1-min episodes ($)"); s.cell(u + 6, 2, f"=B{tot}/(B{u + 1}*1)")
    s.cell(u + 7, 1, "Cost per finished minute, 2-min episodes ($)"); s.cell(u + 7, 2, f"=B{tot}/(B{u + 1}*2)")
    for r in range(u + 1, u + 8):
        s.cell(r, 1).font = F
        s.cell(r, 2).number_format = "0.0" if r == u + 5 else (USD if r > u + 2 else "0")
    s.cell(u + 1, 3, "Benchmark range").font = F_B
    # Benchmarks rows: 6 = LA budget, 8 = episodes, 9 = shoot days, 10 = cost per minute
    bench_rows = [(u + 3, "=Benchmarks!B6/Benchmarks!C8", "=Benchmarks!C6/Benchmarks!B8"),
                  (u + 4, "=Benchmarks!B6/Benchmarks!C9", "=Benchmarks!C6/Benchmarks!B9"),
                  (u + 5, "=Benchmarks!B8/Benchmarks!C9", "=Benchmarks!C8/Benchmarks!B9")]
    for r, lo, hi in bench_rows:
        s.cell(r, 3, lo).font = F_LINK
        s.cell(r, 4, hi).font = F_LINK
        fmt = "0.0" if r == u + 5 else USD
        s.cell(r, 3).number_format = s.cell(r, 4).number_format = fmt
    s.cell(u + 2, 3, "Low").font = F_NOTE
    s.cell(u + 2, 4, "High").font = F_NOTE
    s.cell(u + 6, 3, "=Benchmarks!B10").font = F_LINK
    s.cell(u + 6, 4, "=Benchmarks!C10").font = F_LINK
    s.cell(u + 6, 3).number_format = s.cell(u + 6, 4).number_format = USD
    s.cell(u + 8, 1, "Benchmark ranges derived from LA budgets of $150k-$200k for 60-100 episodes over 8-10 days (Variety via Filmustage).").font = F_NOTE

    # Chart: actual vs benchmark share for categories that exist in both
    ch_r = u + 11
    s.cell(ch_r, 1, "Chart data (categories in both)").font = F_NOTE
    s.cell(ch_r + 1, 1, "Category"); s.cell(ch_r + 1, 2, "This series"); s.cell(ch_r + 1, 3, "Benchmark")
    comp = [c for c in CATEGORIES + ["Contingency"] if c in BENCHMARK]
    for k, cat in enumerate(comp):
        rr = ch_r + 2 + k
        src = r0 + (CATEGORIES + ["Contingency"]).index(cat)
        s.cell(rr, 1, f"=A{src}")
        s.cell(rr, 2, f"=C{src}").number_format = PCT
        s.cell(rr, 3, f"=D{src}").number_format = PCT
    chart = BarChart()
    chart.type = "bar"
    chart.title = "Budget share vs benchmark"
    chart.x_axis.scaling.orientation = "maxMin"  # first category at the top
    chart.y_axis.title = "Share of total"
    chart.y_axis.number_format = "0%"
    data = Reference(s, min_col=2, max_col=3, min_row=ch_r + 1, max_row=ch_r + 1 + len(comp))
    cats = Reference(s, min_col=1, min_row=ch_r + 2, max_row=ch_r + 1 + len(comp))
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    chart.series[0].graphicalProperties.solidFill = "1F3A5F"
    chart.series[1].graphicalProperties.solidFill = "9CA3AF"
    chart.height, chart.width = 9, 18
    chart.roundedCorners = False
    chart.legend.position = "b"
    chart.legend.overlay = False
    chart.title.overlay = False
    chart.y_axis.title.overlay = False
    chart.x_axis.delete = False
    chart.y_axis.delete = False
    s.add_chart(chart, "H4")
    s.freeze_panes = "A5"

    # ------------------------------------------------------------------ Scenario
    sc = wb.create_sheet("Scenario", 2)
    sc["A1"], sc["A1"].font = "Next-series scenarios", F_T
    sc["A2"], sc["A2"].font = "Change the yellow cells. Costs scale with each line's driver (see README, assumption 1).", F_NOTE
    header(sc, 4, ["", "Actual (this series)", "A: recommended, 8-day shoot", "B: tight, 7-day shoot", "C: your scenario"],
           [40, 20, 26, 24, 20])
    sc.cell(5, 1, "Inputs").font = F_B
    inputs = [("Episodes", [54, 54, 54, 54], "0"),
              ("Shoot days", [9, 8, 7, 9], "0"),
              ("Average episode length (minutes)", [1.5, 1.5, 1.5, 1.5], "0.0"),
              ("Contingency (% of subtotal)", [0, 0.05, 0.05, 0.05], "0%")]
    for k, (label, vals, fmt) in enumerate(inputs):
        r = 6 + k
        sc.cell(r, 1, label).font = F
        for c, v in enumerate(vals, start=2):
            cell = sc.cell(r, c, v)
            cell.font, cell.number_format = F_IN, fmt
            if c > 2:
                cell.fill = FILL_KEY
    sc.cell(8, 1).comment = Comment("Episodes run 1-2 minutes; 1.5 is the midpoint (assumption).", "Model")
    sc.cell(11, 1, "Costs ($)").font = F_B
    rows = [
        ("Shoot-day driven costs", lambda col: f'=SUMIFS({AMT},{DRV},"Shoot days")*{col}7/$B$7'),
        ("Episode driven costs", lambda col: f'=SUMIFS({AMT},{DRV},"Episodes")*{col}6/$B$6'),
        ("Fixed costs", lambda col: f'=SUMIFS({AMT},{DRV},"Fixed")'),
        ("Subtotal", lambda col: f"=SUM({col}12:{col}14)"),
        ("Contingency", lambda col: f"={col}15*{col}9"),
        ("Total", lambda col: f"={col}15+{col}16"),
    ]
    for k, (label, fn) in enumerate(rows):
        r = 12 + k
        sc.cell(r, 1, label).font = F_B if label in ("Subtotal", "Total") else F
        for c in range(2, 6):
            col = get_column_letter(c)
            cell = sc.cell(r, c, fn(col))
            cell.number_format = USD
            if label in ("Subtotal", "Total"):
                cell.font, cell.border = F_B, TOP
            elif label != "Contingency":
                cell.font = F_LINK
    sc.cell(19, 1, "Compared with actual").font = F_B
    outs = [("Change vs actual ($)", lambda col: f"={col}17-$B$17", USD),
            ("Change vs actual (%)", lambda col: f"=IF($B$17=0,0,{col}17/$B$17-1)", PCT),
            ("Cost per episode ($)", lambda col: f"={col}17/{col}6", USD),
            ("Cost per shoot day ($)", lambda col: f"={col}17/{col}7", USD),
            ("Cost per finished minute ($)", lambda col: f"={col}17/({col}6*{col}8)", USD),
            ("Episodes per shoot day", lambda col: f"={col}6/{col}7", "0.0")]
    for k, (label, fn, fmt) in enumerate(outs):
        r = 20 + k
        sc.cell(r, 1, label).font = F
        for c in range(2, 6):
            cell = sc.cell(r, c, fn(get_column_letter(c)))
            cell.number_format = fmt
    sc.cell(27, 1, "Producer's judgement: 8 days was achievable for this series; 7 days leaves no room for problems and risks overtime that would erase the saving.").font = F_NOTE
    sc.cell(28, 1, "Contingency is not a saving: it reserves money for overruns. The actual series had none.").font = F_NOTE

    # Scenario by category
    sc.cell(30, 1, "Scenario total by category ($)").font = F_B
    header(sc, 31, ["Category", "Actual", "A", "B", "C"])
    for k, cat in enumerate(CATEGORIES):
        r = 32 + k
        sc.cell(r, 1, cat).font = F
        for c in range(2, 6):
            col = get_column_letter(c)
            sc.cell(r, c, (f'=SUMIFS({AMT},{CAT},$A{r},{DRV},"Shoot days")*{col}$7/$B$7'
                           f'+SUMIFS({AMT},{CAT},$A{r},{DRV},"Episodes")*{col}$6/$B$6'
                           f'+SUMIFS({AMT},{CAT},$A{r},{DRV},"Fixed")')).number_format = USD
    rc = 32 + len(CATEGORIES)
    sc.cell(rc, 1, "Contingency").font = F
    for c in range(2, 6):
        sc.cell(rc, c, f"={get_column_letter(c)}16").number_format = USD
    sc.cell(rc + 1, 1, "Total").font = F_B
    for c in range(2, 6):
        col = get_column_letter(c)
        cell = sc.cell(rc + 1, c, f"=SUM({col}32:{col}{rc})")
        cell.number_format, cell.font, cell.border = USD, F_B, TOP
    sc.cell(rc + 2, 1, "Check: equals Total above").font = F_NOTE
    for c in range(2, 6):
        col = get_column_letter(c)
        sc.cell(rc + 2, c, f'=IF(ABS({col}{rc + 1}-{col}17)<0.01,"OK","MISMATCH")').font = F_NOTE
    sc.freeze_panes = "B5"

    # ------------------------------------------------------------------ People & Paperwork
    pp = wb.create_sheet("People & Paperwork", 3)
    pp["A1"], pp["A1"].font = "People costs and paperwork", F_T
    header(pp, 3, ["Cast tier", "Spend ($)", "People", "Share of cast", "Avg per person ($)"], [34, 14, 10, 14, 18])
    tiers = ["Lead", "Supporting", "Day player", "Background"]
    for k, t in enumerate(tiers):
        r = 4 + k
        pp.cell(r, 1, t).font = F
        pp.cell(r, 2, f'=SUMIFS({AMT},{SEC},"Cast",{ROLE},A{r})').number_format = USD
        pp.cell(r, 3, f'=COUNTIFS({SEC},"Cast",{ROLE},A{r})')
        pp.cell(r, 4, f"=IF($B$8=0,0,B{r}/$B$8)").number_format = PCT
        pp.cell(r, 5, f"=IF(C{r}=0,0,B{r}/C{r})").number_format = USD
    pp.cell(8, 1, "Total cast").font = F_B
    pp.cell(8, 2, "=SUM(B4:B7)").number_format = USD
    pp.cell(8, 3, "=SUM(C4:C7)")
    for c in (1, 2, 3):
        pp.cell(8, c).border = TOP

    header(pp, 10, ["Paperwork status (crew + cast)", "Spend ($)", "Lines", "Share of people spend", ""])
    people = f'SUMIFS({AMT},{SEC},"Crew")+SUMIFS({AMT},{SEC},"Cast")'
    stat = [("All people spend", f"={people}", f'=COUNTIFS({SEC},"Crew")+COUNTIFS({SEC},"Cast")'),
            ("Invoice recorded", f'=SUMIFS({AMT},{SEC},"Crew",{INV},"Yes")+SUMIFS({AMT},{SEC},"Cast",{INV},"Yes")',
             f'=COUNTIFS({SEC},"Crew",{INV},"Yes")+COUNTIFS({SEC},"Cast",{INV},"Yes")'),
            ("Payment recorded", f'=SUMIFS({AMT},{SEC},"Crew",{PAID},"Yes")+SUMIFS({AMT},{SEC},"Cast",{PAID},"Yes")',
             f'=COUNTIFS({SEC},"Crew",{PAID},"Yes")+COUNTIFS({SEC},"Cast",{PAID},"Yes")'),
            ("Contract complete", f'=SUMIFS({AMT},{SEC},"Crew",{CON},"Yes")+SUMIFS({AMT},{SEC},"Cast",{CON},"Yes")',
             f'=COUNTIFS({SEC},"Crew",{CON},"Yes")+COUNTIFS({SEC},"Cast",{CON},"Yes")'),
            ("All three recorded", f'=SUMIFS({AMT},{SEC},"Crew",{INV},"Yes",{PAID},"Yes",{CON},"Yes")+SUMIFS({AMT},{SEC},"Cast",{INV},"Yes",{PAID},"Yes",{CON},"Yes")',
             f'=COUNTIFS({SEC},"Crew",{INV},"Yes",{PAID},"Yes",{CON},"Yes")+COUNTIFS({SEC},"Cast",{INV},"Yes",{PAID},"Yes",{CON},"Yes")'),
            ("No complete contract recorded", "=B11-B14", "=C11-C14")]
    for k, (label, f_amt, f_cnt) in enumerate(stat):
        r = 11 + k
        pp.cell(r, 1, label).font = F_B if k in (0, 5) else F
        pp.cell(r, 2, f_amt).number_format = USD
        pp.cell(r, 3, f_cnt)
        pp.cell(r, 4, f"=IF($B$11=0,0,B{r}/$B$11)").number_format = PCT
    pp.cell(18, 1, '"Not recorded" means the status cell was blank in the original tracker; it may be untracked rather than missing.').font = F_NOTE

    # ------------------------------------------------------------------ Data Quality
    d = wb.create_sheet("Data Quality")
    header(d, 1, ["Source row", "Item", "Issue", "Detail"], [11, 36, 32, 70])
    for i, row in enumerate(dq.itertuples(index=False), start=2):
        for c, v in enumerate([row.source_row, row.item, row.issue, row.detail], start=1):
            d.cell(i, c, v).font = F
    d.cell(len(dq) + 3, 1, "Source row = row number in the synthetic public sample.").font = F_NOTE
    d.freeze_panes = "A2"

    # ------------------------------------------------------------------ Benchmarks
    b = wb.create_sheet("Benchmarks")
    b["A1"], b["A1"].font = "External benchmarks", F_T
    header(b, 3, ["Measure", "Low", "High", "Source"], [44, 14, 14, 90])
    bench = [
        ("Vertical drama budget, Los Angeles ($)", 150000, 200000, "Variety, via Filmustage, 'How to Produce a Vertical Drama on a Budget' (2026): https://filmustage.com/blog/how-to-produce-a-vertical-drama-on-a-budget/", USD),
        ("Vertical drama budget, US range ($)", 150000, 300000, "Variety and TheWrap, via Filmustage (2026), same article", USD),
        ("Episodes per series", 60, 100, "Variety, via Filmustage (2026), same article", "0"),
        ("Shoot days", 8, 10, "Variety, via Filmustage (2026), same article", "0"),
        ("Cost per finished minute ($)", 1500, 4000, "Vitrina, 'Fund Short-Form Series' (2026): https://vitrina.ai/?p=138851", USD),
        ("Cost split: cast / locations / ATL / BTL / gear / post / contingency", None, None,
         "22% / 20% / 17% / 18% / 12% / 6% / 5%. Real Reel line-producer breakdown, via Filmustage (2026)", None),
        ("Fast shoot: episodes covered in 7 days", 70, 100, "Filmustage, 'How to Schedule a Vertical Drama Shoot in 7 Days' (2026): https://filmustage.com/blog/how-to-schedule-a-vertical-drama-shoot-in-7-days/", "0"),
        ("Global micro-drama revenue, 2025 / 2026 ($bn)", 11, 14, "Omdia estimates, reported by Powerdrill (2026): https://powerdrill.ai/blog/the-global-short-drama-economy-is-booming", "0"),
    ]
    # Rows 6-13 so Summary formulas can reference fixed cells
    b.cell(5, 1, "(rows 6-10 feed the Summary unit-cost benchmarks)").font = F_NOTE
    for k, (label, lo, hi, src, fmt) in enumerate(bench):
        r = 6 + k
        b.cell(r, 1, label).font = F
        for c, v in ((2, lo), (3, hi)):
            if v is not None:
                cell = b.cell(r, c, v)
                cell.font, cell.number_format = F_IN, fmt
        b.cell(r, 4, src).font = F
        b.cell(r, 4).alignment = Alignment(wrap_text=True, vertical="top")
    # Summary references: B6/C6 budget, B8/C8 episodes, B9 shoot days? -> keep explicit mapping below
    return wb, items


def main():
    wb, _ = build()
    out = ROOT / "output" / "Production_Cost_Review_Demo.xlsx"
    out.parent.mkdir(exist_ok=True)
    wb.save(out)
    print(out)


if __name__ == "__main__":
    main()

