#!/usr/bin/env python3
"""
Helix MedTech Strategy Simulation — builder.

Single source of truth for the simulation's financial model. Running this script
produces two deliverables that share *identical* maths:

  1. helix-medtech-simulation.xlsx  — a self-contained Excel workbook with
     decision dropdowns, a financial engine, results charts, and a teams
     dashboard. Participants play three rounds entirely in Excel.

  2. medtech-simulation.html        — an interactive web "shell": a teams
     dashboard plus a per-team workspace with one tab per decision area,
     a live financial readout, and a leaderboard. No install, no server.

Company backdrop (chosen with the user):
  "Helix MedTech" — a medical-technology company that combines a mature
  MEDICAL-DEVICES business (the core to *exploit*) with a fast-growing
  DIGITAL-HEALTH SaaS platform (recurring revenue / the growth engine).

The model is deliberately transparent: every option carries additive impacts
that flow through a P&L and an enterprise-value calculation. See MODEL below.
"""

import json

# --------------------------------------------------------------------------- #
#  THE MODEL                                                                   #
# --------------------------------------------------------------------------- #
# Starting position (Round 1 opening balance), all $ in millions.
START = {
    "core_revenue": 1800.0,      # Medical devices (mature, hardware-led)
    "explore_revenue": 200.0,    # Digital-health SaaS + new ventures
    "margin_pct": 60.0,          # Gross margin %
    "opex": 700.0,               # Recurring operating costs (R&D + SG&A)
    "cash": 300.0,               # Cash balance
    "risk": 30.0,                # Risk index 0-100 (lower is better)
    "recurring_mix": 15.0,       # % of revenue that is recurring (SaaS / contracts)
    "base_multiple": 9.0,        # EV/EBITDA base multiple
    "base_core_growth": 1.5,     # organic core growth %/round before decisions
    "base_explore_growth": 8.0,  # organic explore growth %/round before decisions
    "tax_note": 0.0,
}

# How the EV multiple reacts (kept identical in Excel + JS):
#   multiple = clamp(6..18,
#       base_multiple
#       + cumulative_multiple_adjustments
#       + recurring_mix%/100 * 6        (investors pay up for SaaS/recurring)
#       + revenue_growth% * 0.15        (growth premium)
#       - (risk - 30) * 0.05 )          (risk discount)
MULT = {"recur_weight": 6.0, "growth_weight": 0.15, "risk_weight": 0.05,
        "lo": 6.0, "hi": 18.0, "risk_anchor": 30.0}

# Each decision area has 4 options on an intensity scale (0..3).
# Impacts (all additive for the round unless noted):
#   spend   : one-time investment ($M) — reduces cash, drives the benefits
#   coreG   : added to core (devices) growth %, this round
#   expG    : added to explore (SaaS/ventures) growth %, this round
#   margin  : change to gross margin % (carries forward)
#   opex    : change to recurring operating costs $M (carries forward)
#   risk    : change to risk index (carries forward)
#   mult    : change to cumulative EV multiple adjustment (carries forward)
#   recur   : change to recurring-revenue mix % (carries forward)
AREAS = [
    {
        "key": "core",
        "name": "Managing the Core (Exploit)",
        "icon": "⚙️",
        "desc": "How hard you run, defend and modernise the mature medical-devices business.",
        "options": [
            {"name": "Harvest & cut", "scale": "Defensive",
             "blurb": "Maximise short-term cash from devices; minimal reinvestment.",
             "spend": 0,  "coreG": -3, "expG": 0, "margin": 1.5, "opex": -40, "risk": 5,  "mult": -0.5, "recur": 0},
            {"name": "Maintain", "scale": "Steady",
             "blurb": "Keep the core running as-is; protect share without big bets.",
             "spend": 20, "coreG": 0,  "expG": 0, "margin": 0,   "opex": 0,   "risk": 0,  "mult": 0,    "recur": 0},
            {"name": "Optimise & defend", "scale": "Efficiency",
             "blurb": "Lean + automation in plants; targeted upgrades to defend share.",
             "spend": 60, "coreG": 2,  "expG": 0, "margin": 2,   "opex": -10, "risk": -3, "mult": 0.3,  "recur": 0},
            {"name": "Invest to grow core", "scale": "Offensive",
             "blurb": "New indications & geographies for the device line.",
             "spend": 120,"coreG": 5,  "expG": 0, "margin": -0.5,"opex": 20,  "risk": 3,  "mult": 0.4,  "recur": 0},
        ],
    },
    {
        "key": "explore",
        "name": "Explore Portfolio",
        "icon": "\U0001f9ed",
        "desc": "Investment in the digital-health SaaS platform and new growth ventures.",
        "options": [
            {"name": "Pause exploration", "scale": "None",
             "blurb": "Freeze venture spend; bank the savings, accept slower SaaS.",
             "spend": 0,  "coreG": 0, "expG": -10, "margin": 0.5, "opex": 0,  "risk": 4,  "mult": -1.0, "recur": 0},
            {"name": "Selective bets", "scale": "Focused",
             "blurb": "Fund the 1-2 strongest digital-health products only.",
             "spend": 50, "coreG": 0, "expG": 10,  "margin": 0,   "opex": 10, "risk": 0,  "mult": 0.5,  "recur": 2},
            {"name": "Balanced portfolio", "scale": "Balanced",
             "blurb": "A managed portfolio of SaaS bets across the pipeline.",
             "spend": 110,"coreG": 0, "expG": 25,  "margin": -0.5,"opex": 20, "risk": 2,  "mult": 1.2,  "recur": 3},
            {"name": "Aggressive venture build", "scale": "Aggressive",
             "blurb": "Go big on the platform; chase category leadership in SaaS.",
             "spend": 200,"coreG": 0, "expG": 45,  "margin": -1.5,"opex": 40, "risk": 6,  "mult": 2.0,  "recur": 5},
        ],
    },
    {
        "key": "disruption",
        "name": "Managing Disruption",
        "icon": "⚡",
        "desc": "Response to AI diagnostics and connected-device entrants attacking the core.",
        "options": [
            {"name": "Wait & see", "scale": "Reactive",
             "blurb": "Hold position; assume the threat is over-hyped.",
             "spend": 0,  "coreG": -2, "expG": 0,  "margin": 0, "opex": 0,  "risk": 12, "mult": -1.0, "recur": 0},
            {"name": "Partner / license", "scale": "Collaborate",
             "blurb": "License AI / connectivity from a fast mover; integrate quickly.",
             "spend": 40, "coreG": 1,  "expG": 0,  "margin": 0, "opex": 0,  "risk": -4, "mult": 0.5,  "recur": 2},
            {"name": "Acquire capability", "scale": "M&A",
             "blurb": "Buy a disruptor to own the new platform outright.",
             "spend": 180,"coreG": 3,  "expG": 10, "margin": 0, "opex": 15, "risk": -6, "mult": 1.0,  "recur": 3},
            {"name": "Build internal moonshot", "scale": "Build",
             "blurb": "Stand up a skunkworks to leapfrog the entrants.",
             "spend": 130,"coreG": 0,  "expG": 15, "margin": -1,"opex": 0,  "risk": -2, "mult": 0.8,  "recur": 2},
        ],
    },
    {
        "key": "supply",
        "name": "Supply Chain",
        "icon": "\U0001f69a",
        "desc": "Trading off landed cost against resilience for devices and components.",
        "options": [
            {"name": "Lowest cost (single-source)", "scale": "Cost",
             "blurb": "Offshore, single-source for the best unit cost.",
             "spend": 0,  "coreG": 0, "expG": 0, "margin": 2,   "opex": 0,  "risk": 10, "mult": -0.5, "recur": 0},
            {"name": "Dual-source", "scale": "Balanced",
             "blurb": "Second suppliers on critical parts to cut single points of failure.",
             "spend": 40, "coreG": 0, "expG": 0, "margin": 0.5, "opex": 5,  "risk": -3, "mult": 0,    "recur": 0},
            {"name": "Regionalise / nearshore", "scale": "Resilient",
             "blurb": "Nearshore key flows; shorten and de-risk the network.",
             "spend": 90, "coreG": 0, "expG": 0, "margin": -0.5,"opex": 10, "risk": -8, "mult": 0.3,  "recur": 0},
            {"name": "Vertical integration", "scale": "Integrate",
             "blurb": "Automate & bring critical manufacturing in-house.",
             "spend": 160,"coreG": 1, "expG": 0, "margin": 1.5, "opex": -5, "risk": -6, "mult": 0.4,  "recur": 0},
        ],
    },
    {
        "key": "revenue",
        "name": "Revenue & Commercial Model",
        "icon": "\U0001f4b0",
        "desc": "Pricing and business model — from one-off device sales to recurring outcomes.",
        "options": [
            {"name": "Discount for volume", "scale": "Volume",
             "blurb": "Cut price to win units and defend share.",
             "spend": 10, "coreG": 3,  "expG": 0, "margin": -2.5,"opex": 0,  "risk": 0, "mult": -0.3, "recur": 0},
            {"name": "Hold premium pricing", "scale": "Premium",
             "blurb": "Protect price and brand; accept slower volume.",
             "spend": 10, "coreG": -1, "expG": 0, "margin": 2,   "opex": 0,  "risk": 0, "mult": 0.3,  "recur": 0},
            {"name": "Value-based selling", "scale": "Value",
             "blurb": "Sell on clinical & economic outcomes to hospitals.",
             "spend": 50, "coreG": 3,  "expG": 0, "margin": 1,   "opex": 15, "risk": 0, "mult": 0.5,  "recur": 3},
            {"name": "Device-as-a-Service (recurring)", "scale": "Recurring",
             "blurb": "Bundle hardware + SaaS into subscriptions / outcomes contracts.",
             "spend": 90, "coreG": 2,  "expG": 5, "margin": -1,  "opex": 20, "risk": 0, "mult": 1.5,  "recur": 10},
        ],
    },
    {
        "key": "people",
        "name": "People & Talent",
        "icon": "\U0001f465",
        "desc": "Workforce strategy — the execution engine behind every other choice.",
        "options": [
            {"name": "Cost reduction / layoffs", "scale": "Cut",
             "blurb": "Reduce headcount to protect near-term margin.",
             "spend": 0,  "coreG": -2, "expG": -5, "margin": 1,   "opex": -60, "risk": 8,  "mult": -0.8, "recur": 0},
            {"name": "Maintain", "scale": "Steady",
             "blurb": "Hold the workforce steady; no major change.",
             "spend": 20, "coreG": 0,  "expG": 0,  "margin": 0,   "opex": 0,   "risk": 0,  "mult": 0,    "recur": 0},
            {"name": "Upskill & reskill", "scale": "Develop",
             "blurb": "Reskill for software & data; raise productivity.",
             "spend": 60, "coreG": 1,  "expG": 5,  "margin": 1,   "opex": 10,  "risk": -3, "mult": 0.5,  "recur": 0},
            {"name": "Transform culture & talent", "scale": "Transform",
             "blurb": "Top-quartile talent + culture change for agility.",
             "spend": 110,"coreG": 2,  "expG": 10, "margin": 1.5, "opex": 30,  "risk": -5, "mult": 1.0,  "recur": 0},
        ],
    },
    {
        "key": "corporate",
        "name": "Corporate Functions",
        "icon": "\U0001f3db️",
        "desc": "Overhead and the digital backbone (finance, IT, shared services).",
        "options": [
            {"name": "Slash overhead", "scale": "Cut",
             "blurb": "Aggressively cut corporate cost; accept fragility.",
             "spend": 0,  "coreG": 0, "expG": -3, "margin": 0,   "opex": -50, "risk": 6,  "mult": -0.4, "recur": 0},
            {"name": "Maintain", "scale": "Steady",
             "blurb": "Run functions as-is.",
             "spend": 15, "coreG": 0, "expG": 0,  "margin": 0,   "opex": 0,   "risk": 0,  "mult": 0,    "recur": 0},
            {"name": "Selective digitisation", "scale": "Digitise",
             "blurb": "Automate the highest-cost processes first.",
             "spend": 70, "coreG": 0, "expG": 0,  "margin": 0.5, "opex": -20, "risk": -2, "mult": 0.3,  "recur": 0},
            {"name": "Full digital core + shared services", "scale": "Transform",
             "blurb": "One digital backbone; shared services at scale.",
             "spend": 140,"coreG": 1, "expG": 0,  "margin": 1,   "opex": -45, "risk": -4, "mult": 0.6,  "recur": 0},
        ],
    },
]

ROUNDS = [
    {"n": 1, "title": "Pressure on the Core",
     "story": ("Reimbursement cuts squeeze device margins and an AI-diagnostics start-up "
               "lands its first hospital contracts. The board wants a plan that protects "
               "the core while standing up the digital-health platform.")},
    {"n": 2, "title": "Disruption Accelerates",
     "story": ("A rival launches a connected-device + subscription bundle, and a component "
               "shock exposes single-source supply. Investors start rewarding recurring "
               "revenue and punishing fragility.")},
    {"n": 3, "title": "The New Normal",
     "story": ("Value-based care mandates reshape buying. The market now prizes recurring "
               "SaaS revenue, proven innovation and low operating risk. Time to convert "
               "your build-up into enterprise value.")},
]

IMPACT_KEYS = ["spend", "coreG", "expG", "margin", "opex", "risk", "mult", "recur"]

MODEL = {"start": START, "mult": MULT, "areas": AREAS, "rounds": ROUNDS,
         "impact_keys": IMPACT_KEYS}


# --------------------------------------------------------------------------- #
#  EXCEL WORKBOOK                                                              #
# --------------------------------------------------------------------------- #
def build_excel(path):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.worksheet.datavalidation import DataValidation
    from openpyxl.utils import get_column_letter
    from openpyxl.chart import LineChart, BarChart, Reference

    # palette
    NAVY = "0B2545"; TEAL = "0E7C86"; LIGHT = "E8EEF4"; ACCENT = "13B5A6"
    GOLD = "C9A227"; GREY = "5B6B7B"; WHITE = "FFFFFF"; BAND = "F4F7FB"
    f_title = Font(name="Calibri", size=18, bold=True, color=WHITE)
    f_sub   = Font(name="Calibri", size=11, color=WHITE)
    f_h     = Font(name="Calibri", size=11, bold=True, color=WHITE)
    f_b     = Font(name="Calibri", size=11, bold=True, color=NAVY)
    f_n     = Font(name="Calibri", size=11, color="22303C")
    f_small = Font(name="Calibri", size=9, color=GREY, italic=True)
    fill_navy = PatternFill("solid", fgColor=NAVY)
    fill_teal = PatternFill("solid", fgColor=TEAL)
    fill_light= PatternFill("solid", fgColor=LIGHT)
    fill_band = PatternFill("solid", fgColor=BAND)
    fill_gold = PatternFill("solid", fgColor=GOLD)
    fill_acc  = PatternFill("solid", fgColor=ACCENT)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left   = Alignment(horizontal="left", vertical="center", wrap_text=True)
    right  = Alignment(horizontal="right", vertical="center")
    thin = Side(style="thin", color="C7D2DD")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    wb = Workbook()

    def style_header_row(ws, row, c1, c2, fill=fill_navy, font=f_h):
        for c in range(c1, c2 + 1):
            cell = ws.cell(row=row, column=c)
            cell.fill = fill; cell.font = font
            cell.alignment = center; cell.border = border

    # ---------- WELCOME ----------------------------------------------------- #
    ws = wb.active; ws.title = "Welcome"
    ws.sheet_view.showGridLines = False
    ws.merge_cells("B2:H2"); ws["B2"] = "HELIX MEDTECH  —  STRATEGY SIMULATION"
    ws["B2"].font = f_title; ws["B2"].fill = fill_navy; ws["B2"].alignment = left
    for c in range(2, 9): ws.cell(row=2, column=c).fill = fill_navy
    ws.merge_cells("B3:H3")
    ws["B3"] = "Medical devices (the core to exploit) + a digital-health SaaS platform (the growth engine) — three rounds."
    ws["B3"].font = f_sub; ws["B3"].fill = fill_navy; ws["B3"].alignment = left
    for c in range(2, 9): ws.cell(row=3, column=c).fill = fill_navy
    ws.row_dimensions[2].height = 34; ws.row_dimensions[3].height = 20

    def section(row, text):
        ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=8)
        c = ws.cell(row=row, column=2, value=text)
        c.font = Font(bold=True, size=12, color=NAVY)
        c.fill = fill_light; c.alignment = left
        for k in range(2, 9): ws.cell(row=row, column=k).fill = fill_light

    def para(row, text, italic=False):
        ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=8)
        c = ws.cell(row=row, column=2, value=text)
        c.font = Font(size=10, color="22303C", italic=italic); c.alignment = left
        ws.row_dimensions[row].height = 28

    r = 5
    section(r, "The brief"); r += 1
    para(r, "You are the executive team of Helix MedTech. Your mature medical-devices business throws off cash but "
            "is under attack from AI diagnostics and connected-device entrants. Your digital-health SaaS platform is "
            "small but growing fast and is what investors reward. Across three rounds you make seven strategic "
            "decisions each round. The model converts them into financial results and enterprise value (EV)."); r += 2
    section(r, "How to play"); r += 1
    para(r, "1.  Open each 'Round' tab. For all seven decision areas, pick ONE option from the dropdown in the yellow 'Your Choice' cell."); r += 1
    para(r, "2.  Watch the 'Engine' and 'Results' tabs update live — revenue, EBITDA, cash, risk and enterprise value."); r += 1
    para(r, "3.  Play Round 1, then Round 2, then Round 3. Each round opens from the prior round's closing position."); r += 1
    para(r, "4.  Compare teams on the 'Teams Dashboard'. Highest enterprise value at the end of Round 3 wins."); r += 2
    section(r, "How you are scored"); r += 1
    para(r, "Enterprise Value = EBITDA × EV multiple.  The multiple rises with recurring (SaaS) revenue mix and "
            "growth, and falls with operating risk. So short-term cuts can flatter EBITDA but shrink the multiple. "
            "Score = EV indexed to the starting EV (Start = 100), with a penalty if you run cash negative."); r += 2
    section(r, "The seven decision areas"); r += 1
    for a in AREAS:
        para(r, f"•  {a['name']}  —  {a['desc']}"); r += 1

    ws.column_dimensions["A"].width = 2
    for col in "BCDEFGH": ws.column_dimensions[col].width = 17

    # ---------- REFERENCE (decision library + impacts) ---------------------- #
    ref = wb.create_sheet("Reference")
    ref.sheet_view.showGridLines = False
    headers = ["Key", "Area", "Option", "Spend $M", "Core g%", "Explore g%",
               "Margin pp", "Opex $M", "Risk", "Mult", "Recur pp", "Description"]
    for j, h in enumerate(headers, start=1):
        cell = ref.cell(row=1, column=j, value=h)
        cell.fill = fill_navy; cell.font = f_h; cell.alignment = center; cell.border = border
    row = 2
    area_rows = {}  # area key -> (first_row, last_row) for dropdown ranges
    for a in AREAS:
        first = row
        for opt in a["options"]:
            key = f"{a['name']}|{opt['name']}"
            vals = [key, a["name"], opt["name"], opt["spend"], opt["coreG"], opt["expG"],
                    opt["margin"], opt["opex"], opt["risk"], opt["mult"], opt["recur"], opt["blurb"]]
            for j, v in enumerate(vals, start=1):
                cell = ref.cell(row=row, column=j, value=v)
                cell.border = border
                cell.font = f_n
                if j in (2, 3, 12): cell.alignment = left
                else: cell.alignment = center
                if row % 2 == 0: cell.fill = fill_band
            row += 1
        area_rows[a["key"]] = (first, row - 1)
    widths = [1, 30, 30, 10, 9, 11, 10, 9, 7, 7, 9, 52]
    ref.column_dimensions["A"].hidden = True
    for j, w in enumerate(widths, start=1):
        ref.column_dimensions[get_column_letter(j)].width = w
    ref.freeze_panes = "B2"

    # ---------- ROUND SHEETS ------------------------------------------------ #
    # round sheet column map: A area, B choice, C key(hidden),
    #   D spend E coreG F expG G margin H opex I risk J mult K recur
    impact_cols = {"spend": 4, "coreG": 5, "expG": 6, "margin": 7,
                   "opex": 8, "risk": 9, "mult": 10, "recur": 11}
    AREA_FIRST = 4
    TOTAL_ROW = AREA_FIRST + len(AREAS) + 1   # blank row then totals

    for rd in ROUNDS:
        sh = wb.create_sheet(f"Round {rd['n']}")
        sh.sheet_view.showGridLines = False
        sh.merge_cells("A1:K1")
        t = sh.cell(row=1, column=1, value=f"ROUND {rd['n']}  —  {rd['title']}")
        t.font = f_title; t.fill = fill_navy; t.alignment = left
        for c in range(1, 12): sh.cell(row=1, column=1+c-1).fill = fill_navy
        sh.row_dimensions[1].height = 32
        sh.merge_cells("A2:K2")
        s = sh.cell(row=2, column=1, value="Scenario:  " + rd["story"])
        s.font = Font(size=10, italic=True, color="22303C"); s.alignment = left
        sh.row_dimensions[2].height = 44

        hdr = ["Decision Area", "Your Choice", "key", "Spend $M", "Core g%",
               "Explore g%", "Margin pp", "Opex $M", "Risk", "Mult", "Recur pp"]
        for j, h in enumerate(hdr, start=1):
            cell = sh.cell(row=3, column=j, value=h)
            cell.fill = fill_teal; cell.font = f_h; cell.alignment = center; cell.border = border

        for i, a in enumerate(AREAS):
            r = AREA_FIRST + i
            sh.cell(row=r, column=1, value=a["name"]).font = f_b
            sh.cell(row=r, column=1).alignment = left
            sh.cell(row=r, column=1).border = border
            # choice cell (yellow input) — default to "Maintain"-style option index 1
            default = a["options"][1]["name"]
            choice = sh.cell(row=r, column=2, value=default)
            choice.fill = fill_gold; choice.font = Font(bold=True, color=NAVY)
            choice.alignment = center; choice.border = border
            # hidden key
            sh.cell(row=r, column=3, value=f'=A{r}&"|"&B{r}')
            # data validation dropdown for this area
            f0, f1 = area_rows[a["key"]]
            dv = DataValidation(
                type="list",
                formula1=f"=Reference!$C${f0}:$C${f1}",
                allow_blank=False, showDropDown=False)
            dv.error = "Pick one of the listed options."
            dv.prompt = a["desc"]
            sh.add_data_validation(dv); dv.add(choice)
            # impact lookups
            for key, col in impact_cols.items():
                ref_col = {"spend":4,"coreG":5,"expG":6,"margin":7,"opex":8,
                           "risk":9,"mult":10,"recur":11}[key]
                f = (f"=IFERROR(VLOOKUP($C{r},Reference!$A:$K,{ref_col},FALSE),0)")
                c = sh.cell(row=r, column=col, value=f)
                c.alignment = center; c.border = border; c.font = f_n
                if i % 2 == 1: c.fill = fill_band

        # totals row
        sh.cell(row=TOTAL_ROW, column=1, value="ROUND TOTALS").font = f_b
        sh.cell(row=TOTAL_ROW, column=1).fill = fill_light
        sh.cell(row=TOTAL_ROW, column=1).border = border
        sh.cell(row=TOTAL_ROW, column=2).fill = fill_light
        sh.cell(row=TOTAL_ROW, column=2).border = border
        for col in impact_cols.values():
            L = get_column_letter(col)
            c = sh.cell(row=TOTAL_ROW, column=col,
                        value=f"=SUM({L}{AREA_FIRST}:{L}{AREA_FIRST+len(AREAS)-1})")
            c.font = f_b; c.alignment = center; c.fill = fill_light; c.border = border

        sh.column_dimensions["A"].width = 30
        sh.column_dimensions["B"].width = 30
        sh.column_dimensions["C"].hidden = True
        for col in "DEFGHIJK": sh.column_dimensions[col].width = 10
        sh.freeze_panes = "A4"

    # totals cell refs per round (used by Engine)
    def tot(rnd, key):
        col = get_column_letter(impact_cols[key])
        return f"'Round {rnd}'!{col}{TOTAL_ROW}"

    # ---------- ENGINE ------------------------------------------------------ #
    eng = wb.create_sheet("Engine")
    eng.sheet_view.showGridLines = False
    eng.merge_cells("A1:E1")
    e = eng.cell(row=1, column=1, value="FINANCIAL ENGINE  ($M unless noted)")
    e.font = f_title; e.fill = fill_navy; e.alignment = left
    for c in range(1, 6): eng.cell(row=1, column=c).fill = fill_navy
    eng.row_dimensions[1].height = 30
    col_titles = ["Metric", "Start", "Round 1", "Round 2", "Round 3"]
    for j, h in enumerate(col_titles, start=1):
        cell = eng.cell(row=2, column=j, value=h)
        cell.fill = fill_teal; cell.font = f_h; cell.alignment = center; cell.border = border

    # row map
    R = {
        "core": 3, "explore": 4, "total": 5, "growth": 6, "margin": 7,
        "gp": 8, "opex": 9, "ebitda": 10, "ebitda_m": 11, "spend": 12,
        "cash": 13, "risk": 14, "recur": 15, "cummult": 16, "mult": 17,
        "ev": 18, "evchg": 19, "score": 20,
    }
    labels = {
        "core": "Core revenue (devices)", "explore": "Explore revenue (SaaS/ventures)",
        "total": "Total revenue", "growth": "Revenue growth %", "margin": "Gross margin %",
        "gp": "Gross profit", "opex": "Operating costs (recurring)", "ebitda": "EBITDA",
        "ebitda_m": "EBITDA margin %", "spend": "Investment spend (one-time)",
        "cash": "Cash balance", "risk": "Risk index (0-100, lower better)",
        "recur": "Recurring revenue mix %", "cummult": "Cumulative multiple adj.",
        "mult": "EV multiple (x)", "ev": "ENTERPRISE VALUE", "evchg": "EV change vs Start",
        "score": "SCORE (Start = 100)",
    }
    for k, rr in R.items():
        c = eng.cell(row=rr, column=1, value=labels[k])
        c.font = f_b if k in ("ev", "score") else f_n
        c.alignment = left; c.border = border
        if k in ("ev", "score"): c.fill = fill_gold
        elif rr % 2 == 1: c.fill = fill_band

    bc, bcg, bxg = START["base_core_growth"], START["base_core_growth"], START["base_explore_growth"]
    mw = MULT

    # START column (B)
    B = {  # python-side literal start values written as numbers/formulas
        "core": START["core_revenue"], "explore": START["explore_revenue"],
        "margin": START["margin_pct"], "opex": START["opex"], "cash": START["cash"],
        "risk": START["risk"], "recur": START["recurring_mix"],
    }
    eng[f'B{R["core"]}'] = B["core"]
    eng[f'B{R["explore"]}'] = B["explore"]
    eng[f'B{R["total"]}'] = f'=B{R["core"]}+B{R["explore"]}'
    eng[f'B{R["growth"]}'] = 0
    eng[f'B{R["margin"]}'] = B["margin"]
    eng[f'B{R["gp"]}'] = f'=B{R["total"]}*B{R["margin"]}/100'
    eng[f'B{R["opex"]}'] = B["opex"]
    eng[f'B{R["ebitda"]}'] = f'=B{R["gp"]}-B{R["opex"]}'
    eng[f'B{R["ebitda_m"]}'] = f'=B{R["ebitda"]}/B{R["total"]}'
    eng[f'B{R["spend"]}'] = 0
    eng[f'B{R["cash"]}'] = B["cash"]
    eng[f'B{R["risk"]}'] = B["risk"]
    eng[f'B{R["recur"]}'] = B["recur"]
    eng[f'B{R["cummult"]}'] = 0
    eng[f'B{R["mult"]}'] = (
        f'=MAX({mw["lo"]},MIN({mw["hi"]},{START["base_multiple"]}+B{R["cummult"]}'
        f'+B{R["recur"]}/100*{mw["recur_weight"]}+B{R["growth"]}*100*{mw["growth_weight"]}'
        f'-(B{R["risk"]}-{mw["risk_anchor"]})*{mw["risk_weight"]}))')
    eng[f'B{R["ev"]}'] = f'=B{R["ebitda"]}*B{R["mult"]}'
    eng[f'B{R["evchg"]}'] = 0
    eng[f'B{R["score"]}'] = 100

    # Round columns C (R1), D (R2), E (R3)
    round_cols = {1: "C", 2: "D", 3: "E"}
    prev_col = {1: "B", 2: "C", 3: "D"}
    for rnd in (1, 2, 3):
        col = round_cols[rnd]; pcol = prev_col[rnd]
        eng[f'{col}{R["core"]}'] = (
            f'={pcol}{R["core"]}*(1+({START["base_core_growth"]}+{tot(rnd,"coreG")})/100)')
        eng[f'{col}{R["explore"]}'] = (
            f'={pcol}{R["explore"]}*(1+({START["base_explore_growth"]}+{tot(rnd,"expG")})/100)')
        eng[f'{col}{R["total"]}'] = f'={col}{R["core"]}+{col}{R["explore"]}'
        eng[f'{col}{R["growth"]}'] = f'={col}{R["total"]}/{pcol}{R["total"]}-1'
        eng[f'{col}{R["margin"]}'] = f'={pcol}{R["margin"]}+{tot(rnd,"margin")}'
        eng[f'{col}{R["gp"]}'] = f'={col}{R["total"]}*{col}{R["margin"]}/100'
        eng[f'{col}{R["opex"]}'] = f'={pcol}{R["opex"]}+{tot(rnd,"opex")}'
        eng[f'{col}{R["ebitda"]}'] = f'={col}{R["gp"]}-{col}{R["opex"]}'
        eng[f'{col}{R["ebitda_m"]}'] = f'={col}{R["ebitda"]}/{col}{R["total"]}'
        eng[f'{col}{R["spend"]}'] = f'={tot(rnd,"spend")}'
        eng[f'{col}{R["cash"]}'] = f'={pcol}{R["cash"]}+{col}{R["ebitda"]}-{col}{R["spend"]}'
        eng[f'{col}{R["risk"]}'] = f'=MEDIAN(0,{pcol}{R["risk"]}+{tot(rnd,"risk")},100)'
        eng[f'{col}{R["recur"]}'] = f'=MEDIAN(0,{pcol}{R["recur"]}+{tot(rnd,"recur")},90)'
        eng[f'{col}{R["cummult"]}'] = f'={pcol}{R["cummult"]}+{tot(rnd,"mult")}'
        eng[f'{col}{R["mult"]}'] = (
            f'=MAX({mw["lo"]},MIN({mw["hi"]},{START["base_multiple"]}+{col}{R["cummult"]}'
            f'+{col}{R["recur"]}/100*{mw["recur_weight"]}+{col}{R["growth"]}*100*{mw["growth_weight"]}'
            f'-({col}{R["risk"]}-{mw["risk_anchor"]})*{mw["risk_weight"]}))')
        eng[f'{col}{R["ev"]}'] = f'={col}{R["ebitda"]}*{col}{R["mult"]}'
        eng[f'{col}{R["evchg"]}'] = f'={col}{R["ev"]}-B{R["ev"]}'
        eng[f'{col}{R["score"]}'] = (
            f'={col}{R["ev"]}/$B${R["ev"]}*100-MAX(0,-{col}{R["cash"]})*0.1')

    # number formats + styling for engine body
    money_rows = {R["core"], R["explore"], R["total"], R["gp"], R["opex"],
                  R["ebitda"], R["spend"], R["cash"], R["ev"], R["evchg"]}
    pct_rows = {R["growth"], R["margin"], R["ebitda_m"], R["recur"]}
    x_rows = {R["mult"], R["cummult"]}
    for rr in R.values():
        for cc in range(2, 6):
            cell = eng.cell(row=rr, column=cc)
            cell.border = border; cell.alignment = right
            if rr in money_rows: cell.number_format = '#,##0'
            elif rr == R["growth"] or rr == R["ebitda_m"]: cell.number_format = '0.0%'
            elif rr in (R["margin"], R["recur"]): cell.number_format = '0.0'
            elif rr in x_rows: cell.number_format = '0.0"x"'
            elif rr in (R["risk"], R["score"]): cell.number_format = '0.0'
            if rr in (R["ev"], R["score"]):
                cell.font = Font(bold=True, color=NAVY); cell.fill = fill_light
            elif rr % 2 == 1:
                cell.fill = fill_band
    eng.column_dimensions["A"].width = 30
    for col in "BCDE": eng.column_dimensions[col].width = 13
    eng.freeze_panes = "B3"

    # ---------- RESULTS (charts) ------------------------------------------- #
    res = wb.create_sheet("Results")
    res.sheet_view.showGridLines = False
    res.merge_cells("A1:H1")
    rc = res.cell(row=1, column=1, value="RESULTS")
    rc.font = f_title; rc.fill = fill_navy; rc.alignment = left
    for c in range(1, 9): res.cell(row=1, column=c).fill = fill_navy
    res.row_dimensions[1].height = 30

    cats = Reference(eng, min_col=2, max_col=5, min_row=2, max_row=2)

    ln = LineChart(); ln.title = "Enterprise Value ($M)"; ln.height = 8; ln.width = 16
    data = Reference(eng, min_col=2, max_col=5, min_row=R["ev"], max_row=R["ev"])
    ln.add_data(data, from_rows=True, titles_from_data=False)
    ln.set_categories(cats); ln.y_axis.title = "$M"; ln.legend = None
    res.add_chart(ln, "A3")

    bar = BarChart(); bar.title = "Revenue mix ($M)"; bar.height = 8; bar.width = 16
    d2 = Reference(eng, min_col=2, max_col=5, min_row=R["core"], max_row=R["explore"])
    bar.add_data(d2, from_rows=True, titles_from_data=False); bar.set_categories(cats)
    bar.type = "col"; bar.grouping = "stacked"; bar.overlap = 100
    res.add_chart(bar, "A20")

    ln2 = LineChart(); ln2.title = "EBITDA ($M) & EV multiple (x)"; ln2.height = 8; ln2.width = 16
    d3 = Reference(eng, min_col=2, max_col=5, min_row=R["ebitda"], max_row=R["ebitda"])
    ln2.add_data(d3, from_rows=True, titles_from_data=False); ln2.set_categories(cats); ln2.legend = None
    res.add_chart(ln2, "J3")

    ln3 = LineChart(); ln3.title = "Risk index (lower is better)"; ln3.height = 8; ln3.width = 16
    d4 = Reference(eng, min_col=2, max_col=5, min_row=R["risk"], max_row=R["risk"])
    ln3.add_data(d4, from_rows=True, titles_from_data=False); ln3.set_categories(cats); ln3.legend = None
    res.add_chart(ln3, "J20")

    # ---------- TEAMS DASHBOARD -------------------------------------------- #
    td = wb.create_sheet("Teams Dashboard")
    td.sheet_view.showGridLines = False
    td.merge_cells("A1:G1")
    tdc = td.cell(row=1, column=1, value="TEAMS DASHBOARD  —  LEADERBOARD (end of Round 3)")
    tdc.font = f_title; tdc.fill = fill_navy; tdc.alignment = left
    for c in range(1, 8): td.cell(row=1, column=c).fill = fill_navy
    td.row_dimensions[1].height = 30
    td.merge_cells("A2:G2")
    td.cell(row=2, column=1,
            value="Row 1 reflects THIS workbook automatically. Facilitator: type each other team's "
                  "final figures into rows below to rank the room.").font = f_small
    td.cell(row=2, column=1).alignment = left

    head = ["Rank", "Team", "Enterprise Value $M", "EV vs Start $M",
            "Recurring mix %", "Risk", "Score"]
    for j, h in enumerate(head, start=1):
        cell = td.cell(row=3, column=j, value=h)
        cell.fill = fill_teal; cell.font = f_h; cell.alignment = center; cell.border = border

    # row 4 = this team (formulas), rows 5-11 blank for facilitator
    td.cell(row=4, column=2, value="Your Team").font = f_b
    td.cell(row=4, column=3, value=f"=Engine!E{R['ev']}")
    td.cell(row=4, column=4, value=f"=Engine!E{R['evchg']}")
    td.cell(row=4, column=5, value=f"=Engine!E{R['recur']}")
    td.cell(row=4, column=6, value=f"=Engine!E{R['risk']}")
    td.cell(row=4, column=7, value=f"=Engine!E{R['score']}")
    for i in range(5, 12):
        td.cell(row=i, column=2, value=f"Team {i-3}").font = f_n
    for i in range(4, 12):
        td.cell(row=i, column=1, value=f"=IF(C{i}=\"\",\"\",RANK(C{i},$C$4:$C$11))")
        for j in range(1, 8):
            cell = td.cell(row=i, column=j)
            cell.border = border
            cell.alignment = center if j != 2 else left
            if j in (3, 4): cell.number_format = '#,##0'
            if j in (5, 6, 7): cell.number_format = '0.0'
            if i % 2 == 1: cell.fill = fill_band
    td.column_dimensions["A"].width = 7
    td.column_dimensions["B"].width = 18
    for col in "CDEFG": td.column_dimensions[col].width = 17

    wb.active = 0
    wb.save(path)
    return path


# --------------------------------------------------------------------------- #
#  HTML WEB SHELL                                                              #
# --------------------------------------------------------------------------- #
HTML_TEMPLATE = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8"/>
<meta name="viewport" content="width=device-width, initial-scale=1"/>
<title>Helix MedTech — Strategy Simulation</title>
<style>
  :root{
    --navy:#0B2545; --navy2:#13315C; --teal:#0E7C86; --accent:#13B5A6;
    --gold:#C9A227; --light:#E8EEF4; --band:#F4F7FB; --ink:#22303C;
    --grey:#5B6B7B; --good:#1E9E6A; --bad:#D64545; --line:#D7E0EA;
  }
  *{box-sizing:border-box}
  body{margin:0;font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Helvetica,Arial,sans-serif;
       color:var(--ink);background:#F0F4F8}
  header{background:linear-gradient(110deg,var(--navy),var(--navy2));color:#fff;padding:14px 22px;
         display:flex;align-items:center;gap:18px;flex-wrap:wrap;box-shadow:0 2px 10px rgba(0,0,0,.15)}
  header h1{font-size:18px;margin:0;letter-spacing:.4px}
  header .tag{font-size:12px;opacity:.85;margin-top:2px}
  header .spacer{flex:1}
  .nav{display:flex;gap:8px}
  .nav button{background:rgba(255,255,255,.10);color:#fff;border:1px solid rgba(255,255,255,.25);
       padding:8px 14px;border-radius:8px;cursor:pointer;font-size:13px;font-weight:600}
  .nav button.active{background:var(--accent);border-color:var(--accent)}
  .pill{background:rgba(255,255,255,.12);border:1px solid rgba(255,255,255,.25);padding:6px 10px;
        border-radius:20px;font-size:12px;display:flex;align-items:center;gap:8px}
  select,input{font:inherit}
  .wrap{max-width:1180px;margin:18px auto;padding:0 16px}
  .card{background:#fff;border:1px solid var(--line);border-radius:12px;padding:18px;
        box-shadow:0 1px 4px rgba(11,37,69,.06)}
  .grid{display:grid;gap:14px}
  h2{margin:.1em 0 .5em;color:var(--navy);font-size:18px}
  h3{margin:.2em 0;color:var(--navy);font-size:15px}
  .muted{color:var(--grey);font-size:13px}
  /* dashboard */
  .teams{grid-template-columns:repeat(auto-fill,minmax(260px,1fr))}
  .team-card{cursor:pointer;transition:transform .08s,box-shadow .12s;position:relative}
  .team-card:hover{transform:translateY(-2px);box-shadow:0 6px 18px rgba(11,37,69,.12)}
  .team-card .ev{font-size:26px;font-weight:800;color:var(--navy)}
  .team-card .delta{font-size:13px;font-weight:700}
  .rankbadge{position:absolute;top:12px;right:12px;background:var(--gold);color:#fff;border-radius:20px;
        width:26px;height:26px;display:flex;align-items:center;justify-content:center;font-weight:800;font-size:13px}
  .bar{height:8px;background:var(--band);border-radius:6px;overflow:hidden;margin-top:8px}
  .bar > i{display:block;height:100%;background:var(--accent)}
  .row{display:flex;justify-content:space-between;align-items:center;gap:10px}
  .btn{background:var(--teal);color:#fff;border:none;padding:9px 16px;border-radius:8px;cursor:pointer;
       font-weight:700;font-size:13px}
  .btn.ghost{background:#fff;color:var(--teal);border:1px solid var(--teal)}
  .btn.gold{background:var(--gold)}
  .btn.warn{background:var(--bad)}
  .btn:disabled{opacity:.45;cursor:not-allowed}
  /* workspace */
  .ws{display:grid;grid-template-columns:230px 1fr;gap:16px;align-items:start}
  .tabs{display:flex;flex-direction:column;gap:6px}
  .tabs button{text-align:left;background:#fff;border:1px solid var(--line);border-radius:9px;padding:10px 12px;
       cursor:pointer;font-size:13px;font-weight:600;color:var(--navy);display:flex;gap:9px;align-items:center}
  .tabs button.active{background:var(--navy);color:#fff;border-color:var(--navy)}
  .tabs button .dot{margin-left:auto;width:9px;height:9px;border-radius:50%;background:var(--line)}
  .tabs button.done .dot{background:var(--good)}
  .opt{border:1.5px solid var(--line);border-radius:10px;padding:12px 14px;margin:8px 0;cursor:pointer;
       display:grid;grid-template-columns:22px 1fr;gap:10px;align-items:start}
  .opt:hover{border-color:var(--accent)}
  .opt.sel{border-color:var(--teal);background:#F2FBFA;box-shadow:0 0 0 2px rgba(14,124,134,.15)}
  .opt .rad{width:18px;height:18px;border-radius:50%;border:2px solid var(--grey);margin-top:2px}
  .opt.sel .rad{border-color:var(--teal);background:var(--teal);box-shadow:inset 0 0 0 3px #fff}
  .opt .scale{font-size:11px;font-weight:700;color:var(--teal);text-transform:uppercase;letter-spacing:.4px}
  .opt .nm{font-weight:700;color:var(--navy)}
  .opt .bl{font-size:13px;color:var(--ink);margin-top:2px}
  .chips{display:flex;flex-wrap:wrap;gap:6px;margin-top:7px}
  .chip{font-size:11px;padding:3px 8px;border-radius:20px;background:var(--band);border:1px solid var(--line);font-weight:600}
  .chip.pos{color:var(--good);border-color:#bfe6d4}
  .chip.neg{color:var(--bad);border-color:#f0c6c6}
  /* readout */
  .readout{position:sticky;top:14px}
  .kpi{display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-top:6px}
  .kpi .k{background:var(--band);border:1px solid var(--line);border-radius:9px;padding:9px 11px}
  .kpi .k .lab{font-size:11px;color:var(--grey);font-weight:600}
  .kpi .k .val{font-size:18px;font-weight:800;color:var(--navy)}
  .kpi .k.big{grid-column:1 / -1;background:linear-gradient(110deg,var(--navy),var(--navy2));border:none}
  .kpi .k.big .lab{color:#bcd0e6}
  .kpi .k.big .val{color:#fff;font-size:26px}
  .delta.pos{color:var(--good)} .delta.neg{color:var(--bad)}
  table{width:100%;border-collapse:collapse;font-size:13px}
  th,td{padding:7px 9px;border-bottom:1px solid var(--line);text-align:right}
  th:first-child,td:first-child{text-align:left}
  thead th{background:var(--band);color:var(--navy)}
  .roundbar{display:flex;gap:8px;align-items:center;flex-wrap:wrap;margin-bottom:12px}
  .roundbar .step{padding:6px 12px;border-radius:20px;border:1px solid var(--line);background:#fff;font-size:13px;font-weight:700;color:var(--grey)}
  .roundbar .step.active{background:var(--gold);color:#fff;border-color:var(--gold)}
  .roundbar .step.done{background:var(--good);color:#fff;border-color:var(--good)}
  .scenario{background:var(--band);border-left:4px solid var(--teal);padding:10px 14px;border-radius:8px;font-size:13px;margin-bottom:12px}
  .modal-bg{position:fixed;inset:0;background:rgba(11,37,69,.45);display:flex;align-items:center;justify-content:center;z-index:50}
  .modal{background:#fff;border-radius:12px;padding:22px;max-width:440px;width:92%}
  .hide{display:none!important}
  .spark{display:flex;align-items:flex-end;gap:6px;height:46px;margin-top:6px}
  .spark > i{flex:1;background:var(--accent);border-radius:3px 3px 0 0;min-height:3px}
  footer{max-width:1180px;margin:10px auto 30px;padding:0 16px;color:var(--grey);font-size:12px}
</style>
</head>
<body>
<header>
  <div>
    <h1>⚕️ Helix MedTech — Strategy Simulation</h1>
    <div class="tag">Medical devices (exploit) + digital-health SaaS (explore) · three rounds</div>
  </div>
  <div class="spacer"></div>
  <div class="nav">
    <button id="navDash" class="active" onclick="go('dash')">Teams Dashboard</button>
    <button id="navWork" onclick="go('work')">Team Workspace</button>
  </div>
</header>

<!-- DASHBOARD -->
<div id="view-dash" class="wrap">
  <div class="card" style="margin-bottom:14px">
    <div class="row">
      <div>
        <h2 style="margin:0">Teams</h2>
        <div class="muted">Create a team per table. Click a card to enter its workspace. Ranked by enterprise value.</div>
      </div>
      <div style="display:flex;gap:8px">
        <input id="newTeam" placeholder="New team name" style="padding:9px 12px;border:1px solid var(--line);border-radius:8px"/>
        <button class="btn" onclick="addTeam()">+ Add team</button>
        <button class="btn ghost" onclick="resetAll()">Reset all</button>
      </div>
    </div>
  </div>
  <div id="teamGrid" class="grid teams"></div>
</div>

<!-- WORKSPACE -->
<div id="view-work" class="wrap hide">
  <div class="card" style="margin-bottom:14px">
    <div class="row">
      <div>
        <h2 style="margin:0" id="wsTitle">Team</h2>
        <div class="muted" id="wsSub"></div>
      </div>
      <button class="btn ghost" onclick="go('dash')">← Back to teams</button>
    </div>
    <div class="roundbar" id="roundBar" style="margin-top:12px"></div>
    <div class="scenario" id="scenario"></div>
  </div>

  <div class="ws">
    <div>
      <div class="tabs" id="tabs"></div>
      <div style="margin-top:12px">
        <button class="btn gold" id="advanceBtn" style="width:100%" onclick="advance()">Lock round &amp; advance →</button>
      </div>
    </div>

    <div class="grid" style="grid-template-columns:1fr 320px;gap:16px;align-items:start">
      <div class="card" id="decisionPane"></div>
      <div class="card readout" id="readout"></div>
    </div>
  </div>
</div>

<footer>
  Single-file shell · progress saved in your browser · model mirrors the companion Excel workbook
  <span id="cfgNote"></span>.
</footer>

<script>
const MODEL = __MODEL_JSON__;
const AREAS = MODEL.areas, ROUNDS = MODEL.rounds, START = MODEL.start, MW = MODEL.mult;
const NROUNDS = ROUNDS.length;
const LS_KEY = "helix_medtech_sim_v1";

/* ---------------- state ---------------- */
function blankChoices(){ // choices[round][areaKey] = optionIndex (default 1 = maintain-ish)
  return ROUNDS.map(()=>{ const o={}; AREAS.forEach(a=>o[a.key]=1); return o; });
}
let state = load();
function load(){
  try{ const s=JSON.parse(localStorage.getItem(LS_KEY)); if(s&&s.teams) return s; }catch(e){}
  return {teams:[ mkTeam("Team 1"), mkTeam("Team 2") ], current:null, ui:{round:0, tab:AREAS[0].key}};
}
function mkTeam(name){ return {id:"t"+Math.random().toString(36).slice(2,8), name, choices:blankChoices(), locked:[false,false,false]}; }
function save(){ localStorage.setItem(LS_KEY, JSON.stringify(state)); }

/* ---------------- model maths (mirror of Excel engine) ---------------- */
function simulate(team){
  // returns array of per-round states incl. "start" at index 0
  let st = {
    label:"Start", core:START.core_revenue, explore:START.explore_revenue,
    margin:START.margin_pct, opex:START.opex, cash:START.cash, risk:START.risk,
    recur:START.recurring_mix, cummult:0
  };
  st.total = st.core+st.explore; st.growth=0; st.gp=st.total*st.margin/100;
  st.ebitda=st.gp-st.opex; st.spend=0;
  st.mult = clampMult(START.base_multiple+st.cummult+st.recur/100*MW.recur_weight+st.growth*100*MW.growth_weight-(st.risk-MW.risk_anchor)*MW.risk_weight);
  st.ev = st.ebitda*st.mult; st.evchg=0; st.score=100;
  const out=[st];
  for(let r=0;r<NROUNDS;r++){
    const prev=out[out.length-1];
    const t=totals(team, r);
    const s={label:"R"+(r+1)};
    s.core = prev.core*(1+(START.base_core_growth+t.coreG)/100);
    s.explore = prev.explore*(1+(START.base_explore_growth+t.expG)/100);
    s.total = s.core+s.explore;
    s.growth = s.total/prev.total-1;
    s.margin = prev.margin+t.margin;
    s.gp = s.total*s.margin/100;
    s.opex = prev.opex+t.opex;
    s.ebitda = s.gp-s.opex;
    s.ebitda_m = s.ebitda/s.total;
    s.spend = t.spend;
    s.cash = prev.cash + s.ebitda - s.spend;
    s.risk = Math.max(0,Math.min(100, prev.risk+t.risk));
    s.recur = Math.max(0,Math.min(90, prev.recur+t.recur));
    s.cummult = prev.cummult + t.mult;
    s.mult = clampMult(START.base_multiple+s.cummult+s.recur/100*MW.recur_weight+s.growth*100*MW.growth_weight-(s.risk-MW.risk_anchor)*MW.risk_weight);
    s.ev = s.ebitda*s.mult;
    s.evchg = s.ev - out[0].ev;
    s.score = s.ev/out[0].ev*100 - Math.max(0,-s.cash)*0.1;
    out.push(s);
  }
  return out;
}
function clampMult(m){ return Math.max(MW.lo, Math.min(MW.hi, m)); }
function totals(team, r){
  const t={spend:0,coreG:0,expG:0,margin:0,opex:0,risk:0,mult:0,recur:0};
  AREAS.forEach(a=>{
    const idx=team.choices[r][a.key]; const o=a.options[idx];
    MODEL.impact_keys.forEach(k=>{ t[k]+= (o[k]||0); });
  });
  return t;
}

/* ---------------- routing ---------------- */
function go(v){
  document.getElementById("view-dash").classList.toggle("hide", v!=="dash");
  document.getElementById("view-work").classList.toggle("hide", v!=="work");
  document.getElementById("navDash").classList.toggle("active", v==="dash");
  document.getElementById("navWork").classList.toggle("active", v==="work");
  if(v==="dash") renderDash();
  if(v==="work") renderWork();
}

/* ---------------- dashboard ---------------- */
function addTeam(){
  const el=document.getElementById("newTeam"); const n=(el.value||"").trim();
  if(!n) return; state.teams.push(mkTeam(n)); el.value=""; save(); renderDash();
}
function resetAll(){
  if(!confirm("Reset ALL teams and decisions? This cannot be undone.")) return;
  state={teams:[mkTeam("Team 1"),mkTeam("Team 2")], current:null, ui:{round:0,tab:AREAS[0].key}};
  save(); renderDash();
}
function delTeam(id,ev){ ev.stopPropagation();
  if(!confirm("Delete this team?")) return;
  state.teams=state.teams.filter(t=>t.id!==id); save(); renderDash();
}
function fmt(n){ return Math.round(n).toLocaleString(); }
function renderDash(){
  const grid=document.getElementById("teamGrid"); grid.innerHTML="";
  const sims=state.teams.map(t=>({t, s:simulate(t)}));
  const startEV=sims.length?sims[0].s[0].ev:1;
  const ranked=[...sims].sort((a,b)=>b.s[NROUNDS].ev-a.s[NROUNDS].ev);
  const rankOf={}; ranked.forEach((x,i)=>rankOf[x.t.id]=i+1);
  const maxEV=Math.max(...sims.map(x=>x.s[NROUNDS].ev),1);
  sims.forEach(({t,s})=>{
    const fin=s[NROUNDS]; const up=fin.evchg>=0;
    const spark=s.map(x=>`<i style="height:${Math.max(4,(x.ev/maxEV)*46)}px"></i>`).join("");
    const card=document.createElement("div");
    card.className="card team-card";
    card.onclick=()=>openTeam(t.id);
    card.innerHTML=`
      <div class="rankbadge">${rankOf[t.id]}</div>
      <h3 style="margin:0 36px 2px 0">${esc(t.name)}</h3>
      <div class="muted">Enterprise value (end R${NROUNDS})</div>
      <div class="ev">$${fmt(fin.ev)}M</div>
      <div class="delta ${up?'pos':'neg'}">${up?'▲':'▼'} $${fmt(Math.abs(fin.evchg))}M vs start (${(fin.ev/startEV*100).toFixed(0)} index)</div>
      <div class="spark">${spark}</div>
      <div class="bar"><i style="width:${(fin.ev/maxEV*100).toFixed(0)}%"></i></div>
      <table style="margin-top:10px">
        <tr><td>Revenue</td><td>$${fmt(fin.total)}M</td><td>Recurring</td><td>${fin.recur.toFixed(0)}%</td></tr>
        <tr><td>EBITDA</td><td>$${fmt(fin.ebitda)}M</td><td>Risk</td><td>${fin.risk.toFixed(0)}</td></tr>
        <tr><td>Cash</td><td>$${fmt(fin.cash)}M</td><td>Score</td><td><b>${fin.score.toFixed(0)}</b></td></tr>
      </table>
      <div class="row" style="margin-top:10px">
        <button class="btn" onclick="event.stopPropagation();openTeam('${t.id}')">Open workspace →</button>
        <button class="btn warn ghost" style="background:#fff;color:var(--bad);border:1px solid var(--bad)" onclick="delTeam('${t.id}',event)">Delete</button>
      </div>`;
    grid.appendChild(card);
  });
}

/* ---------------- workspace ---------------- */
function openTeam(id){ state.current=id; state.ui={round:firstOpenRound(team()),tab:AREAS[0].key}; save(); go("work"); }
function team(){ return state.teams.find(t=>t.id===state.current); }
function firstOpenRound(t){ for(let r=0;r<NROUNDS;r++){ if(!t.locked[r]) return r; } return NROUNDS-1; }

function renderWork(){
  const t=team(); if(!t){ go("dash"); return; }
  document.getElementById("wsTitle").textContent=t.name;
  const sims=simulate(t); const cur=sims[state.ui.round+1];
  document.getElementById("wsSub").textContent=
    `Round ${state.ui.round+1} of ${NROUNDS} · ${ROUNDS[state.ui.round].title}`;
  // round bar
  const rb=document.getElementById("roundBar"); rb.innerHTML="";
  ROUNDS.forEach((rd,i)=>{
    const cls = t.locked[i]?"done":(i===state.ui.round?"active":"");
    const b=document.createElement("button"); b.className="step "+cls;
    b.textContent=`Round ${rd.n}: ${rd.title}`+(t.locked[i]?" ✓":"");
    b.onclick=()=>{ state.ui.round=i; save(); renderWork(); };
    rb.appendChild(b);
  });
  document.getElementById("scenario").innerHTML =
    `<b>Scenario — Round ${ROUNDS[state.ui.round].n}: ${ROUNDS[state.ui.round].title}.</b> ${ROUNDS[state.ui.round].story}`;
  // tabs
  const tabs=document.getElementById("tabs"); tabs.innerHTML="";
  AREAS.forEach(a=>{
    const b=document.createElement("button");
    b.className="tab"+(a.key===state.ui.tab?" active":"")+ " done";
    if(a.key===state.ui.tab) b.classList.add("active");
    b.innerHTML=`<span>${a.icon}</span><span>${a.name}</span><span class="dot"></span>`;
    b.onclick=()=>{ state.ui.tab=a.key; save(); renderWork(); };
    tabs.appendChild(b);
  });
  renderDecision(t);
  renderReadout(t, sims);
  const adv=document.getElementById("advanceBtn");
  const locked=t.locked[state.ui.round];
  adv.textContent = locked ? (state.ui.round<NROUNDS-1?"Go to next round →":"Simulation complete ✓")
                           : (state.ui.round<NROUNDS-1?"Lock round & advance →":"Lock final round ✓");
}

function renderDecision(t){
  const a=AREAS.find(x=>x.key===state.ui.tab);
  const r=state.ui.round; const locked=t.locked[r];
  const pane=document.getElementById("decisionPane");
  let html=`<div class="row"><h2>${a.icon} ${a.name}</h2><span class="muted">Round ${r+1}</span></div>
            <div class="muted" style="margin-bottom:6px">${a.desc}</div>`;
  if(locked) html+=`<div class="scenario" style="border-color:var(--good)">This round is locked. Reopen it from the round bar only by resetting — decisions are final.</div>`;
  a.options.forEach((o,idx)=>{
    const sel = t.choices[r][a.key]===idx;
    html+=`<div class="opt ${sel?'sel':''}" onclick="${locked?'':`pick('${a.key}',${idx})`}">
      <div class="rad"></div>
      <div>
        <div class="scale">${o.scale}</div>
        <div class="nm">${o.name}</div>
        <div class="bl">${o.bl||o.blurb}</div>
        <div class="chips">${impactChips(o)}</div>
      </div></div>`;
  });
  pane.innerHTML=html;
}
function impactChips(o){
  const defs=[["spend","Invest","$",true],["coreG","Core","%"],["expG","SaaS","%"],
              ["margin","Margin","pp"],["opex","Opex","$",true],["risk","Risk","",true],
              ["mult","Multiple","x"],["recur","Recurring","pp"]];
  let out="";
  defs.forEach(([k,lab,unit,inv])=>{
    const v=o[k]||0; if(v===0) return;
    // "good" direction: lower spend/opex/risk is good; otherwise higher is good
    let good = inv ? v<0 : v>0;
    const cls = good?"pos":"neg";
    const sign = v>0?"+":"";
    out+=`<span class="chip ${cls}">${lab} ${sign}${v}${unit}</span>`;
  });
  return out||'<span class="chip">no change</span>';
}
function pick(areaKey,idx){ const t=team(); if(t.locked[state.ui.round])return;
  t.choices[state.ui.round][areaKey]=idx; save(); renderWork(); }

function renderReadout(t, sims){
  const r=state.ui.round; const cur=sims[r+1]; const prev=sims[r];
  const ro=document.getElementById("readout");
  const d=(now,was,inv)=>{ const x=now-was; const good=inv?x<0:x>0;
    return `<span class="delta ${x===0?'':(good?'pos':'neg')}">${x>=0?'▲':'▼'} ${fmt(Math.abs(x))}</span>`; };
  ro.innerHTML=`
    <h3>Live financials — end of Round ${r+1}</h3>
    <div class="muted">Reflects choices across rounds 1–${r+1}.</div>
    <div class="kpi">
      <div class="k big"><div class="lab">ENTERPRISE VALUE</div>
        <div class="val">$${fmt(cur.ev)}M</div>
        <div class="delta ${cur.evchg>=0?'pos':'neg'}">${cur.evchg>=0?'▲':'▼'} $${fmt(Math.abs(cur.evchg))}M vs start · multiple ${cur.mult.toFixed(1)}x</div></div>
      <div class="k"><div class="lab">Revenue</div><div class="val">$${fmt(cur.total)}M</div>${d(cur.total,prev.total)}</div>
      <div class="k"><div class="lab">EBITDA</div><div class="val">$${fmt(cur.ebitda)}M</div>${d(cur.ebitda,prev.ebitda)}</div>
      <div class="k"><div class="lab">Cash</div><div class="val">$${fmt(cur.cash)}M</div>${cur.cash<0?'<span class="delta neg">negative!</span>':d(cur.cash,prev.cash)}</div>
      <div class="k"><div class="lab">Risk index</div><div class="val">${cur.risk.toFixed(0)}</div>${d(cur.risk,prev.risk,true)}</div>
      <div class="k"><div class="lab">Recurring mix</div><div class="val">${cur.recur.toFixed(0)}%</div>${d(cur.recur,prev.recur)}</div>
      <div class="k"><div class="lab">Score</div><div class="val">${cur.score.toFixed(0)}</div>${d(cur.score,prev.score)}</div>
    </div>
    <h3 style="margin-top:14px">Trajectory</h3>
    <table>
      <thead><tr><th>Metric</th>${sims.map(s=>`<th>${s.label}</th>`).join("")}</tr></thead>
      <tbody>
        ${trajRow("Revenue",sims,s=>'$'+fmt(s.total))}
        ${trajRow("EBITDA",sims,s=>'$'+fmt(s.ebitda))}
        ${trajRow("EV mult",sims,s=>s.mult.toFixed(1)+'x')}
        ${trajRow("Ent. value",sims,s=>'$'+fmt(s.ev))}
      </tbody>
    </table>`;
}
function trajRow(lab,sims,f){ return `<tr><td>${lab}</td>${sims.map(s=>`<td>${f(s)}</td>`).join("")}</tr>`; }

function advance(){
  const t=team(); const r=state.ui.round;
  if(!t.locked[r]){ t.locked[r]=true; }
  if(r<NROUNDS-1){ state.ui.round=r+1; }
  save(); renderWork();
}
function esc(s){ return (s+"").replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c])); }

/* boot */
go("dash");
</script>
</body>
</html>
"""


def build_html(path):
    js = json.dumps(MODEL, ensure_ascii=False)
    html = HTML_TEMPLATE.replace("__MODEL_JSON__", js)
    with open(path, "w", encoding="utf-8") as f:
        f.write(html)
    return path


if __name__ == "__main__":
    import os
    here = os.path.dirname(os.path.abspath(__file__))
    xlsx = build_excel(os.path.join(here, "helix-medtech-simulation.xlsx"))
    html = build_html(os.path.join(here, "medtech-simulation.html"))
    # index.html is an identical copy so GitHub Pages serves the app at the site root.
    index = build_html(os.path.join(here, "index.html"))
    print("Built:")
    print("  ", xlsx)
    print("  ", html)
    print("  ", index)
