#!/usr/bin/env python3
"""
Builds "Portfolio_Planner_CA.xlsx" — a Google-Sheets-compatible planning workbook
for a Canadian (Ontario) investor:

  Dashboard | Inputs | Projection | Tax Estimate | Engine (Monthly) | Assumptions & Notes
  (+ 3 hidden scenario engines: no-withdrawal counterfactual, return -1%, return +1%)

Everything except the yellow input cells is formula-driven, so the file stays fully
recalculable after the user edits inputs in Google Sheets or Excel.

Tax parameters are for tax year 2026 (verified July 2026 — see Assumptions tab for sources).
"""

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, Protection
from openpyxl.utils import get_column_letter as gcl
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.chart import LineChart, BarChart, AreaChart, Reference
from openpyxl.chart.marker import Marker

# ----------------------------------------------------------------------------- constants

MAXM = 600          # engine rows: months 1..600 (50 years hard cap)
MAXY = 50
ROW0 = 3            # engine row holding "month 0" (opening state)
ENG_FIRST = 4       # engine row of month 1
ENG_LAST = ROW0 + MAXM  # 603

# 2026 tax parameters (verified 2026-07-10; sources documented on Assumptions tab)
FED_BRACKETS = [(0, 0.14), (58523, 0.205), (117045, 0.26), (181440, 0.29), (258482, 0.33)]
FED_BPA_MAX, FED_BPA_MIN = 16452, 14829
FED_PHASE_LOW, FED_PHASE_HIGH = 181440, 258482
ON_BRACKETS = [(0, 0.0505), (53891, 0.0915), (107785, 0.1116), (150000, 0.1216), (220000, 0.1316)]
ON_BPA = 12989
ON_SUR1, ON_SUR2 = 5818, 7446
ON_SURR1, ON_SURR2 = 0.20, 0.36
GUE, GUN = 0.38, 0.15                    # dividend gross-ups
FDTCE, FDTCN = 0.150198, 0.090301        # federal DTC as % of grossed-up dividend
ODTCE, ODTCN = 0.10, 0.029863            # Ontario DTC as % of grossed-up dividend
CG_INC = 0.50                            # capital gains inclusion rate

ACCTS = ["NR", "TF", "RR"]
ACCT_LABEL = {"NR": "Non-Registered", "TF": "TFSA", "RR": "RRSP/RRIF"}

E_BASE = "Engine (Monthly)"
E_NOWD = "Engine (No Withdrawals)"
E_M1 = "Engine (-1%)"
E_P1 = "Engine (+1%)"

# palette (muted, professional)
NAVY = "33475B"
NAVY_LIGHT = "50647A"
INPUT_FILL = "FFF2CC"
CALC_FILL = "EFF2F6"
CARD_FILL = "F7F9FB"
BORDER_GRAY = "C9D1D9"
SERIES = ["4C78A8", "E0A458", "72B7B2", "9C755F", "C25B5B", "7F7F7F"]

MONEY = '$#,##0'
MONEY2 = '$#,##0.00'
PCT = '0.00%'

thin = Side(style="thin", color=BORDER_GRAY)
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)

F_TITLE = Font(name="Calibri", size=16, bold=True, color=NAVY)
F_H2 = Font(name="Calibri", size=12, bold=True, color="FFFFFF")
F_LABEL = Font(name="Calibri", size=11, color="222222")
F_SMALL = Font(name="Calibri", size=9, color="666666")
F_KPI = Font(name="Calibri", size=20, bold=True, color=NAVY)
FILL_H2 = PatternFill("solid", fgColor=NAVY)
FILL_IN = PatternFill("solid", fgColor=INPUT_FILL)
FILL_CALC = PatternFill("solid", fgColor=CALC_FILL)
FILL_CARD = PatternFill("solid", fgColor=CARD_FILL)

UNLOCKED = Protection(locked=False)


def defname(wb, name, ref):
    wb.defined_names[name] = DefinedName(name, attr_text=ref)


def q(sheet):
    return f"'{sheet}'"


# ------------------------------------------------------------------- engine column map

def engine_cols():
    """name -> column index. Streams are (account, sub 1..6)."""
    c = {}
    c["T"], c["ACT"], c["YR"], c["CT"], c["WT"] = 1, 2, 3, 4, 5
    i = 6
    for a in ACCTS:
        for s in range(6):
            c[f"P_{a}{s+1}"] = i + s          # post-growth balance (pre-sale)
        for s in range(6):
            c[f"D_{a}{s+1}"] = i + 6 + s      # dividends generated
        for s in range(6):
            c[f"K_{a}{s+1}"] = i + 12 + s     # closing balance
        for j, agg in enumerate(["CONTRIB", "PG", "DIV", "GROWTH", "CAP", "CAPB",
                                 "DRAW", "FROMCASH", "SALE", "REINV", "CASH", "BAL"]):
            c[f"{agg}_{a}"] = i + 18 + j
        i += 30
    for j, x in enumerate(["DIVEL", "DIVNE", "DIVFOR", "ACBB", "GAIN", "ACB",
                           "BALTOT", "DRAWTOT", "FCTOT", "SALETOT", "UNMET"]):
        c[x] = i + j
    return c


EC = engine_cols()
EL = {k: gcl(v) for k, v in EC.items()}      # name -> column letter
ENG_NCOLS = max(EC.values())


def eref(sheet, name, row=None, rng=None):
    """absolute reference into an engine sheet column."""
    L = EL[name]
    if rng:
        return f"{q(sheet)}!${L}${rng[0]}:${L}${rng[1]}"
    return f"{q(sheet)}!${L}${row}"


# ------------------------------------------------------------------------------ Inputs

def build_inputs(wb):
    ws = wb.create_sheet("Inputs")
    ws.sheet_properties.tabColor = "E8B93E"
    ws.sheet_view.showGridLines = False
    for col, w in {"A": 30, "B": 16, "C": 12, "D": 10, "E": 10, "F": 16, "G": 12,
                   "H": 12, "I": 12, "J": 26, "K": 4, "L": 24, "M": 14}.items():
        ws.column_dimensions[col].width = w

    ws["A1"] = "Inputs — fill the yellow cells only"
    ws["A1"].font = F_TITLE
    ws["A2"] = "Everything else in this workbook recalculates automatically. Grey cells are computed — do not edit."
    ws["A2"].font = F_SMALL

    def section(row, text):
        ws.cell(row=row, column=1, value=text).font = F_H2
        for cc in range(1, 11):
            ws.cell(row=row, column=cc).fill = FILL_H2

    def input_cell(addr, value, fmt=None, note=None, note_col="D"):
        ws[addr] = value
        ws[addr].fill = FILL_IN
        ws[addr].border = BOX
        ws[addr].protection = UNLOCKED
        if fmt:
            ws[addr].number_format = fmt
        if note:
            r = int(addr[1:])
            ws[f"{note_col}{r}"] = note
            ws[f"{note_col}{r}"].font = F_SMALL

    # --- 2.1 portfolio setup
    section(3, "1 · Portfolio setup")
    ws["A4"] = "Current total portfolio value (CAD)"
    input_cell("B4", 500000, MONEY, "Market value of all accounts today.")
    ws["A5"] = "Projection horizon"
    input_cell("B5", 20, "0", None)
    input_cell("C5", "Years")
    ws["D5"] = "Number + unit. Engine runs monthly (max 50 years)."
    ws["D5"].font = F_SMALL
    ws["A6"] = "Annual inflation assumption"
    input_cell("B6", 0.02, PCT, "Used for withdrawal indexing and the real-dollar view.")
    ws["A7"] = "Expected-return override (± all sub-sections)"
    input_cell("B7", 0.0, PCT, "Adds to every sub-section's return. Leave 0% normally; scenario table uses ±1% around this.")

    # --- 2.2 sub-portfolios
    section(9, "2 · Sub-portfolio allocation (up to 6)")
    heads = ["Name", "Risk", "Alloc %", "Return %/yr (ex-div)", "Dividend yield %/yr",
             "Dividend type", "% Non-Reg", "% TFSA", "% RRSP/RRIF", "Row check"]
    for j, h in enumerate(heads):
        cell = ws.cell(row=10, column=1 + j, value=h)
        cell.font = Font(bold=True, size=10, color=NAVY)
        cell.alignment = Alignment(wrap_text=True, vertical="top")
    defaults = [
        ("Canadian Dividend Stocks", "Medium", 0.30, 0.040, 0.040, "Eligible", 0.60, 0.20, 0.20),
        ("US / Global Growth Equities", "High", 0.30, 0.070, 0.012, "Foreign / interest", 0.40, 0.30, 0.30),
        ("Bonds & GICs", "Low", 0.25, 0.015, 0.035, "Foreign / interest", 0.20, 0.20, 0.60),
        ("Canadian REITs", "Medium", 0.15, 0.030, 0.045, "Non-eligible", 0.30, 0.40, 0.30),
        ("", "", 0, 0, 0, "Eligible", 0, 0, 0),
        ("", "", 0, 0, 0, "Eligible", 0, 0, 0),
    ]
    for s in range(6):
        r = 11 + s
        name, risk, alloc, ret, yld, typ, nr, tf, rr = defaults[s]
        for col, val, fmt in [(1, name, None), (2, risk, None), (3, alloc, PCT), (4, ret, PCT),
                              (5, yld, PCT), (6, typ, None), (7, nr, PCT), (8, tf, PCT), (9, rr, PCT)]:
            cc = ws.cell(row=r, column=col, value=val)
            cc.fill = FILL_IN
            cc.border = BOX
            cc.protection = UNLOCKED
            if fmt:
                cc.number_format = fmt
        ws.cell(row=r, column=10,
                value=f'=IF(N($C{r})=0,"",IF(ABS(N($G{r})+N($H{r})+N($I{r})-1)<0.0001,"✓","⚠ accounts must total 100%"))')
        ws.cell(row=r, column=10).font = F_SMALL
        n = s + 1
        for nm, col in [("Name", 1), ("Alloc", 3), ("Ret", 4), ("Yld", 5),
                        ("Type", 6), ("NR", 7), ("TF", 8), ("RR", 9)]:
            defname(wb, f"S{n}{nm}", f"Inputs!${gcl(col)}${r}")
    ws["A17"] = "Allocation check:"
    ws["A17"].font = Font(bold=True, size=10)
    ws["C17"] = "=SUM(C11:C16)"
    ws["C17"].number_format = PCT
    ws["C17"].font = Font(bold=True)
    ws["E17"] = '=IF(ABS(SUM($C$11:$C$16)-1)<0.0001,"✓ Allocations total 100%","⚠ Allocations must total 100% — currently "&TEXT(SUM($C$11:$C$16),"0.0%"))'

    # --- 2.3 contributions
    section(19, "3 · Contributions")
    ws["A20"] = "Contribution amount (CAD)"
    input_cell("B20", 500, MONEY, "Per period (see frequency).")
    ws["A21"] = "Frequency"
    input_cell("B21", "Bi-weekly", None, "Weekly ×52/12, bi-weekly ×26/12 → monthly equivalent.")
    ws["A22"] = "Start month (1 = first month)"
    input_cell("B22", 1, "0")
    ws["A23"] = "End month (blank = end of horizon)"
    input_cell("B23", 60, "0", "Many planners stop contributing when withdrawals start.")
    ws["A24"] = "Annual contribution escalation"
    input_cell("B24", 0.03, PCT, "Contribution grows this much every 12 months.")
    ws["A25"] = "Contribution destination"
    input_cell("B25", "Follow allocation %", None, "Or route 100% to one named sub-section.")

    # --- 2.4 withdrawals
    section(27, "4 · Withdrawals")
    ws["A28"] = "Withdrawal amount (CAD)"
    input_cell("B28", 2000, MONEY, "Per period (see frequency).")
    ws["A29"] = "Frequency"
    input_cell("B29", "Monthly", None, "Bi-weekly ×26/12; Annual = withdrawn once every 12th month.")
    ws["A30"] = "Start month"
    input_cell("B30", 61, "0", "e.g. 61 = start of year 6.")
    ws["A31"] = "End month (blank = end of horizon)"
    input_cell("B31", "", "0")
    ws["A32"] = "Index withdrawals to inflation?"
    input_cell("B32", "Yes", None, "Steps up once per year by the inflation assumption.")
    ws["A33"] = "Funding order — 1st account drawn"
    input_cell("B33", "Non-Registered", None, "That month's dividends/cash in the account are used first,")
    ws["A34"] = "Funding order — 2nd account drawn"
    input_cell("B34", "TFSA", None, "then assets are sold pro-rata across its sub-sections.")
    ws["A35"] = "Funding order — 3rd account drawn"
    input_cell("B35", "RRSP/RRIF", None, "Each account must appear exactly once.")
    ws["A36"] = "Order check:"
    ws["B36"] = ('=IF(AND(COUNTIF($B$33:$B$35,"Non-Registered")=1,COUNTIF($B$33:$B$35,"TFSA")=1,'
                 'COUNTIF($B$33:$B$35,"RRSP/RRIF")=1),"✓ OK","⚠ each account exactly once")')
    ws["B36"].font = F_SMALL

    # --- 2.5 tax settings
    section(38, "5 · Tax settings")
    ws["A39"] = "Province"
    ws["B39"] = "Ontario"
    ws["B39"].fill = FILL_CALC
    ws["D39"] = "Fixed for now; tax tables on the Assumptions tab are structured so other provinces can be added."
    ws["D39"].font = F_SMALL
    ws["A40"] = "Other annual taxable income (CAD)"
    input_cell("B40", 60000, MONEY, 'Used only for the "stacked" tax view.')
    ws["A41"] = "ACB today, as % of non-registered value"
    input_cell("B41", 0.70, PCT, "Adjusted cost base — gains are taxed on sale proceeds minus cost base.")
    ws["A42"] = "Reinvest dividends when not withdrawing?"
    input_cell("B42", "Yes", None, "If No, unused dividends accumulate as cash (0% return) in the account.")

    # --- calculated block
    ws["A44"] = "Calculated values (do not edit)"
    ws["A44"].font = Font(bold=True, size=11, color=NAVY)
    calc = [
        ("Horizon (months)", "HorizonM", '=MIN(600,MAX(1,IF($C$5="Years",N($B$5)*12,N($B$5))))', "0"),
        ("Horizon (years, rounded up)", "HorizonY", "=ROUNDUP(HorizonM/12,0)", "0"),
        ("Contribution / month", "ContribMonthly",
         '=N($B$20)*IF($B$21="Weekly",52/12,IF($B$21="Bi-weekly",26/12,1))', MONEY2),
        ("Contribution start (effective)", "CStartEff", "=IF(N($B$22)<1,1,$B$22)", "0"),
        ("Contribution end (effective)", "CEndEff", "=IF(N($B$23)<1,HorizonM,MIN($B$23,HorizonM))", "0"),
        ("Withdrawal / month (non-annual)", "WdMonthlyBase",
         '=N($B$28)*IF($B$29="Bi-weekly",26/12,1)', MONEY2),
        ("Withdrawal start (effective)", "WdStartEff", "=IF(N($B$30)<1,1,$B$30)", "0"),
        ("Withdrawal end (effective)", "WdEndEff", "=IF(N($B$31)<1,HorizonM,MIN($B$31,HorizonM))", "0"),
        ("Rank: Non-Registered", "RankNR", '=IFERROR(MATCH("Non-Registered",$B$33:$B$35,0),99)', "0"),
        ("Rank: TFSA", "RankTF", '=IFERROR(MATCH("TFSA",$B$33:$B$35,0),99)', "0"),
        ("Rank: RRSP/RRIF", "RankRR", '=IFERROR(MATCH("RRSP/RRIF",$B$33:$B$35,0),99)', "0"),
    ]
    r = 45
    for label, nm, f, fmt in calc:
        ws.cell(row=r, column=1, value=label).font = F_SMALL
        cc = ws.cell(row=r, column=2, value=f)
        cc.fill = FILL_CALC
        cc.number_format = fmt
        defname(wb, nm, f"Inputs!$B${r}")
        r += 1
    # destination shares DS1..DS6
    for s in range(6):
        n = s + 1
        ws.cell(row=r, column=1, value=f"Contribution share → sub {n}").font = F_SMALL
        cc = ws.cell(row=r, column=2,
                     value=f'=IF($B$25="Follow allocation %",N(S{n}Alloc),IF(AND(S{n}Name<>"",$B$25=S{n}Name),1,0))')
        cc.fill = FILL_CALC
        cc.number_format = PCT
        defname(wb, f"DS{n}", f"Inputs!$B${r}")
        r += 1

    # destination dropdown source list (col L)
    ws["L44"] = "Destination options"
    ws["L44"].font = F_SMALL
    ws["L45"] = "Follow allocation %"
    for s in range(6):
        ws[f"L{46+s}"] = f'=IF(S{s+1}Name="","",S{s+1}Name)'

    # names for the simple inputs
    for nm, addr in [("StartValue", "B4"), ("HorizonN", "B5"), ("HorizonUnit", "C5"),
                     ("Inflation", "B6"), ("RetOverride", "B7"),
                     ("ContribAmt", "B20"), ("ContribFreq", "B21"), ("ContribStart", "B22"),
                     ("ContribEnd", "B23"), ("ContribEsc", "B24"), ("ContribDest", "B25"),
                     ("WdAmt", "B28"), ("WdFreq", "B29"), ("WdStart", "B30"), ("WdEnd", "B31"),
                     ("WdIndexed", "B32"), ("OtherInc", "B40"), ("ACBPct", "B41"), ("DRIP", "B42")]:
        defname(wb, nm, f"Inputs!${addr[0]}${addr[1:]}")

    # data validation
    def dv_list(formula, cells, allow_blank=True):
        dv = DataValidation(type="list", formula1=formula, allow_blank=allow_blank, showDropDown=False)
        ws.add_data_validation(dv)
        for cell in cells:
            dv.add(cell)

    dv_list('"Months,Years"', ["C5"])
    dv_list('"Low,Medium,High"', [f"B{r}" for r in range(11, 17)])
    dv_list('"Eligible,Non-eligible,Foreign / interest"', [f"F{r}" for r in range(11, 17)])
    dv_list('"Weekly,Bi-weekly,Monthly"', ["B21"])
    dv_list('"Bi-weekly,Monthly,Annual"', ["B29"])
    dv_list('"Yes,No"', ["B32", "B42"])
    dv_list('"Non-Registered,TFSA,RRSP/RRIF"', ["B33", "B34", "B35"])
    dv_list("=$L$45:$L$51", ["B25"])
    dv_num = DataValidation(type="decimal", operator="between", formula1="0", formula2="1",
                            allow_blank=True, errorStyle="warning",
                            error="Enter a percentage between 0% and 100% (e.g. 30%).")
    ws.add_data_validation(dv_num)
    for rr in range(11, 17):
        for col in ("C", "E", "G", "H", "I"):
            dv_num.add(f"{col}{rr}")
    dv_num.add("B6"); dv_num.add("B24"); dv_num.add("B41")

    ws.protection.sheet = True
    return ws


# --------------------------------------------------------------------------- engines

def build_engine(wb, title, delta_name, with_wd, hidden):
    ws = wb.create_sheet(title)
    ws.sheet_properties.tabColor = "8AA0B4"
    if hidden:
        ws.sheet_state = "hidden"
    ws.freeze_panes = "F4"

    ws["A1"] = "Row 1 = calculation helpers (monthly rate multipliers & contribution shares)."
    ws["A1"].font = F_SMALL
    # headers
    hdr = {v: k for k, v in EC.items()}
    NICE = {"T": "Month", "ACT": "Active", "YR": "Year", "CT": "Contribution $", "WT": "Withdrawal target $",
            "DIVEL": "NR divs — eligible", "DIVNE": "NR divs — non-eligible", "DIVFOR": "NR divs — foreign/interest",
            "ACBB": "ACB base (pre-sale)", "GAIN": "Realized gain (NR)", "ACB": "ACB (closing)",
            "BALTOT": "Total balance", "DRAWTOT": "Total withdrawn", "FCTOT": "Funded by divs/cash",
            "SALETOT": "Funded by sales", "UNMET": "Shortfall flag"}
    AGG = {"CONTRIB": "contribution", "PG": "balance post-growth", "DIV": "dividends", "GROWTH": "growth",
           "CAP": "capacity", "CAPB": "capacity ahead", "DRAW": "withdrawn", "FROMCASH": "from divs/cash",
           "SALE": "sold", "REINV": "reinvested", "CASH": "cash bucket", "BAL": "balance"}
    for idx in range(1, ENG_NCOLS + 1):
        name = hdr[idx]
        if name in NICE:
            label = NICE[name]
        else:
            kind, rest = name.split("_")
            a, s = rest[:2], rest[2:]
            base = {"P": "post-growth", "D": "dividends", "K": "closing"}.get(kind)
            label = (f"{a} S{s} {base}" if base else f"{a} {AGG[kind]}")
        c = ws.cell(row=2, column=idx, value=label)
        c.font = Font(bold=True, size=9, color="FFFFFF")
        c.fill = FILL_H2
        c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.column_dimensions[gcl(idx)].width = 11

    # helper row 1
    for a in ACCTS:
        for s in range(1, 7):
            ws.cell(row=1, column=EC[f"P_{a}{s}"],
                    value=f"=POWER(1+N(S{s}Ret)+{delta_name},1/12)")
            ws.cell(row=1, column=EC[f"D_{a}{s}"],
                    value=f"=POWER(1+N(S{s}Yld),1/12)-1")
            ws.cell(row=1, column=EC[f"K_{a}{s}"],
                    value=f"=N(DS{s})*N(S{s}{a})")

    # month 0 row
    ws.cell(row=ROW0, column=EC["T"], value=0)
    ws.cell(row=ROW0, column=EC["ACT"], value=1)
    ws.cell(row=ROW0, column=EC["YR"], value=0)
    ws.cell(row=ROW0, column=EC["CT"], value=0)
    ws.cell(row=ROW0, column=EC["WT"], value=0)
    for a in ACCTS:
        for s in range(1, 7):
            ws.cell(row=ROW0, column=EC[f"K_{a}{s}"],
                    value=f"=StartValue*N(S{s}Alloc)*N(S{s}{a})")
            ws.cell(row=ROW0, column=EC[f"P_{a}{s}"], value=0)
            ws.cell(row=ROW0, column=EC[f"D_{a}{s}"], value=0)
        for agg in ["CONTRIB", "PG", "DIV", "GROWTH", "CAP", "CAPB", "DRAW", "FROMCASH", "SALE", "REINV", "CASH"]:
            ws.cell(row=ROW0, column=EC[f"{agg}_{a}"], value=0)
        k1, k6 = EL[f"K_{a}1"], EL[f"K_{a}6"]
        ws.cell(row=ROW0, column=EC[f"BAL_{a}"], value=f"=SUM({k1}{ROW0}:{k6}{ROW0})")
    kn1, kn6 = EL["K_NR1"], EL["K_NR6"]
    ws.cell(row=ROW0, column=EC["DIVEL"], value=0)
    ws.cell(row=ROW0, column=EC["DIVNE"], value=0)
    ws.cell(row=ROW0, column=EC["DIVFOR"], value=0)
    ws.cell(row=ROW0, column=EC["ACBB"], value=f"=N(ACBPct)*SUM({kn1}{ROW0}:{kn6}{ROW0})")
    ws.cell(row=ROW0, column=EC["GAIN"], value=0)
    ws.cell(row=ROW0, column=EC["ACB"], value=f"=N(ACBPct)*SUM({kn1}{ROW0}:{kn6}{ROW0})")
    ws.cell(row=ROW0, column=EC["BALTOT"],
            value=f"={EL['BAL_NR']}{ROW0}+{EL['BAL_TF']}{ROW0}+{EL['BAL_RR']}{ROW0}")
    for x in ["DRAWTOT", "FCTOT", "SALETOT", "UNMET"]:
        ws.cell(row=ROW0, column=EC[x], value=0)

    rank = {"NR": "RankNR", "TF": "RankTF", "RR": "RankRR"}

    for r in range(ENG_FIRST, ENG_LAST + 1):
        pr = r - 1
        ws.cell(row=r, column=EC["T"], value=f"=A{pr}+1")
        ws.cell(row=r, column=EC["ACT"], value=f"=IF($A{r}<=HorizonM,1,0)")
        ws.cell(row=r, column=EC["YR"], value=f"=INT(($A{r}-1)/12)+1")
        ws.cell(row=r, column=EC["CT"],
                value=(f"=IF($B{r}=0,0,IF(AND($A{r}>=CStartEff,$A{r}<=CEndEff),"
                       f"ContribMonthly*POWER(1+N(ContribEsc),INT(($A{r}-CStartEff)/12)),0))"))
        if with_wd:
            ws.cell(row=r, column=EC["WT"],
                    value=(f'=IF($B{r}=0,0,IF(AND($A{r}>=WdStartEff,$A{r}<=WdEndEff),'
                           f'IF(WdFreq="Annual",IF(MOD($A{r}-WdStartEff,12)=0,N(WdAmt),0),WdMonthlyBase)'
                           f'*IF(WdIndexed="Yes",POWER(1+N(Inflation),INT(($A{r}-WdStartEff)/12)),1),0))'))
        else:
            ws.cell(row=r, column=EC["WT"], value=0)

        for a in ACCTS:
            PG, DIVc, CASH = EL[f"PG_{a}"], EL[f"DIV_{a}"], EL[f"CASH_{a}"]
            CAP, CAPB, DRAW = EL[f"CAP_{a}"], EL[f"CAPB_{a}"], EL[f"DRAW_{a}"]
            FC, SALE, REINV = EL[f"FROMCASH_{a}"], EL[f"SALE_{a}"], EL[f"REINV_{a}"]
            p1, p6 = EL[f"P_{a}1"], EL[f"P_{a}6"]
            d1, d6 = EL[f"D_{a}1"], EL[f"D_{a}6"]
            k1, k6 = EL[f"K_{a}1"], EL[f"K_{a}6"]
            for s in range(1, 7):
                p, d, k = EL[f"P_{a}{s}"], EL[f"D_{a}{s}"], EL[f"K_{a}{s}"]
                ws.cell(row=r, column=EC[f"P_{a}{s}"],
                        value=f"=IF($B{r}=0,{k}{pr},({k}{pr}+$D{r}*{k}$1)*{p}$1)")
                ws.cell(row=r, column=EC[f"D_{a}{s}"],
                        value=f"=IF($B{r}=0,0,({k}{pr}+$D{r}*{k}$1)*{d}$1)")
                ws.cell(row=r, column=EC[f"K_{a}{s}"],
                        value=f"=IF({PG}{r}=0,{p}{r},{p}{r}*(1-({SALE}{r}-{REINV}{r})/{PG}{r}))")
            shares = "+".join(f"{EL[f'K_{a}{s}']}$1" for s in range(1, 7))
            ws.cell(row=r, column=EC[f"CONTRIB_{a}"], value=f"=$D{r}*({shares})")
            ws.cell(row=r, column=EC[f"PG_{a}"], value=f"=SUM({p1}{r}:{p6}{r})")
            ws.cell(row=r, column=EC[f"DIV_{a}"], value=f"=SUM({d1}{r}:{d6}{r})")
            ws.cell(row=r, column=EC[f"GROWTH_{a}"],
                    value=f"={PG}{r}-SUM({k1}{pr}:{k6}{pr})-{EL[f'CONTRIB_{a}']}{r}")
            ws.cell(row=r, column=EC[f"CAP_{a}"], value=f"={CASH}{pr}+{DIVc}{r}+{PG}{r}")
            others = [b for b in ACCTS if b != a]
            terms = "+".join(f"IF({rank[b]}<{rank[a]},{EL[f'CAP_{b}']}{r},0)" for b in others)
            ws.cell(row=r, column=EC[f"CAPB_{a}"], value=f"={terms}")
            ws.cell(row=r, column=EC[f"DRAW_{a}"], value=f"=MAX(0,MIN({CAP}{r},$E{r}-{CAPB}{r}))")
            ws.cell(row=r, column=EC[f"FROMCASH_{a}"], value=f"=MIN({DRAW}{r},{CASH}{pr}+{DIVc}{r})")
            ws.cell(row=r, column=EC[f"SALE_{a}"], value=f"={DRAW}{r}-{FC}{r}")
            ws.cell(row=r, column=EC[f"REINV_{a}"],
                    value=f'=IF(DRIP="Yes",{CASH}{pr}+{DIVc}{r}-{FC}{r},0)')
            ws.cell(row=r, column=EC[f"CASH_{a}"],
                    value=f'=IF(DRIP="Yes",0,{CASH}{pr}+{DIVc}{r}-{FC}{r})')
            ws.cell(row=r, column=EC[f"BAL_{a}"], value=f"=SUM({k1}{r}:{k6}{r})+{CASH}{r}")

        # NR tax detail
        for colname, typ in [("DIVEL", "Eligible"), ("DIVNE", "Non-eligible"), ("DIVFOR", "Foreign / interest")]:
            terms = "+".join(f'IF(S{s}Type="{typ}",{EL[f"D_NR{s}"]}{r},0)' for s in range(1, 7))
            ws.cell(row=r, column=EC[colname], value=f"={terms}")
        ACBB, ACB = EL["ACBB"], EL["ACB"]
        PGN, SALEN, REINVN = EL["PG_NR"], EL["SALE_NR"], EL["REINV_NR"]
        ws.cell(row=r, column=EC["ACBB"], value=f"={ACB}{pr}+{EL['CONTRIB_NR']}{r}")
        ws.cell(row=r, column=EC["GAIN"],
                value=f"=IF({PGN}{r}=0,0,{SALEN}{r}*(1-{ACBB}{r}/{PGN}{r}))")
        ws.cell(row=r, column=EC["ACB"],
                value=f"=IF({PGN}{r}=0,{ACBB}{r},{ACBB}{r}*(1-{SALEN}{r}/{PGN}{r}))+{REINVN}{r}")
        ws.cell(row=r, column=EC["BALTOT"],
                value=f"={EL['BAL_NR']}{r}+{EL['BAL_TF']}{r}+{EL['BAL_RR']}{r}")
        ws.cell(row=r, column=EC["DRAWTOT"],
                value=f"={EL['DRAW_NR']}{r}+{EL['DRAW_TF']}{r}+{EL['DRAW_RR']}{r}")
        ws.cell(row=r, column=EC["FCTOT"],
                value=f"={EL['FROMCASH_NR']}{r}+{EL['FROMCASH_TF']}{r}+{EL['FROMCASH_RR']}{r}")
        ws.cell(row=r, column=EC["SALETOT"],
                value=f"={EL['SALE_NR']}{r}+{EL['SALE_TF']}{r}+{EL['SALE_RR']}{r}")
        ws.cell(row=r, column=EC["UNMET"],
                value=f"=IF(AND($B{r}=1,$E{r}-{EL['DRAWTOT']}{r}>0.01),1,0)")

    # number formats
    money_cols = [i for i in range(4, ENG_NCOLS + 1)
                  if hdr_name(i) not in ("UNMET",)]
    for i in money_cols:
        for r in range(ROW0, ENG_LAST + 1):
            ws.cell(row=r, column=i).number_format = '#,##0'
    ws.protection.sheet = True
    return ws


def hdr_name(idx):
    for k, v in EC.items():
        if v == idx:
            return k
    return ""


# ------------------------------------------------------------------- assumptions & tax tables

def build_assumptions(wb):
    ws = wb.create_sheet("Assumptions & Notes")
    ws.sheet_properties.tabColor = "9C9C9C"
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 46
    for c in "BCDE":
        ws.column_dimensions[c].width = 15
    ws.column_dimensions["F"].width = 90

    ws["A1"] = "Assumptions, tax tables & conventions"
    ws["A1"].font = F_TITLE
    ws["A2"] = "Tax year 2026 parameters, verified 2026-07-10. The engine references these named cells — update here when CRA publishes new figures."
    ws["A2"].font = F_SMALL

    ws["A4"] = "Federal brackets (2026)"
    ws["A4"].font = Font(bold=True, color=NAVY)
    ws["B4"], ws["C4"], ws["D4"] = "Lower bound", "Rate", "Rate step"
    for j, (lo, rate) in enumerate(FED_BRACKETS):
        r = 5 + j
        ws.cell(row=r, column=2, value=lo).number_format = MONEY
        ws.cell(row=r, column=3, value=rate).number_format = PCT
        step = rate - (FED_BRACKETS[j - 1][1] if j else 0)
        ws.cell(row=r, column=4, value=round(step, 6)).number_format = PCT
    defname(wb, "FedThr", "'Assumptions & Notes'!$B$5:$B$9")
    defname(wb, "FedDlt", "'Assumptions & Notes'!$D$5:$D$9")

    ws["A11"] = "Ontario brackets (2026)"
    ws["A11"].font = Font(bold=True, color=NAVY)
    for j, (lo, rate) in enumerate(ON_BRACKETS):
        r = 12 + j
        ws.cell(row=r, column=2, value=lo).number_format = MONEY
        ws.cell(row=r, column=3, value=rate).number_format = PCT
        step = rate - (ON_BRACKETS[j - 1][1] if j else 0)
        ws.cell(row=r, column=4, value=round(step, 6)).number_format = PCT
    defname(wb, "ONThr", "'Assumptions & Notes'!$B$12:$B$16")
    defname(wb, "ONDlt", "'Assumptions & Notes'!$D$12:$D$16")

    singles = [
        ("Federal lowest rate", "FedLowRate", FED_BRACKETS[0][1], PCT),
        ("Federal BPA (max)", "FedBPAMax", FED_BPA_MAX, MONEY),
        ("Federal BPA (min)", "FedBPAMin", FED_BPA_MIN, MONEY),
        ("Federal BPA phase-out floor", "FedPhLow", FED_PHASE_LOW, MONEY),
        ("Federal BPA phase-out ceiling", "FedPhHigh", FED_PHASE_HIGH, MONEY),
        ("Ontario lowest rate", "ONLowRate", ON_BRACKETS[0][1], PCT),
        ("Ontario BPA", "ONBPA", ON_BPA, MONEY),
        ("Ontario surtax threshold 1 (20%)", "ONSur1", ON_SUR1, MONEY),
        ("Ontario surtax threshold 2 (36%)", "ONSur2", ON_SUR2, MONEY),
        ("Ontario surtax rate 1", "ONSurR1", ON_SURR1, PCT),
        ("Ontario surtax rate 2", "ONSurR2", ON_SURR2, PCT),
        ("Eligible dividend gross-up", "GUE", GUE, PCT),
        ("Eligible DTC — federal (% of grossed-up)", "FDTCE", FDTCE, "0.0000%"),
        ("Eligible DTC — Ontario", "ODTCE", ODTCE, PCT),
        ("Non-eligible dividend gross-up", "GUN", GUN, PCT),
        ("Non-eligible DTC — federal", "FDTCN", FDTCN, "0.0000%"),
        ("Non-eligible DTC — Ontario", "ODTCN", ODTCN, "0.0000%"),
        ("Capital gains inclusion rate", "CGInc", CG_INC, PCT),
        ("Scenario delta — base", "DeltaBase", "=N(RetOverride)", PCT),
        ("Scenario delta — minus 1%", "DeltaM1", "=N(RetOverride)-0.01", PCT),
        ("Scenario delta — plus 1%", "DeltaP1", "=N(RetOverride)+0.01", PCT),
    ]
    r = 18
    for label, nm, val, fmt in singles:
        ws.cell(row=r, column=1, value=label)
        cc = ws.cell(row=r, column=2, value=val)
        cc.number_format = fmt
        defname(wb, nm, f"'Assumptions & Notes'!$B${r}")
        r += 1

    notes = [
        ("SOURCES (verified 2026-07-10)", [
            "Federal brackets & 14% lowest rate: CRA 2026 tax-year page (canada.ca) + Bill C-4, Royal Assent 2026-03-12; thresholds = 2025 × 1.020 indexation.",
            "Federal BPA $16,452 / $14,829 with phase-out $181,440→$258,482: CRA indexation release; KPMG personal tax credits table (Dec 31, 2025).",
            "Ontario brackets, BPA $12,989, surtax $5,818 / $7,446 (indexation 1.019): taxtips.ca/on.htm; EY Ontario 2026 table (2026-01-15).",
            "Dividend gross-ups & credits (38% / 15.0198% / 10%; 15% / 9.0301% / 2.9863%): CRA line 40425; ontario.ca dividend tax credit page.",
            "Capital gains inclusion 50%: proposed 2/3 rate increase cancelled 2025-03-21 (Dept. of Finance), never enacted.",
        ]),
        ("ENGINE CONVENTIONS", [
            "Monthly engine. Annual rates converted to effective monthly: (1+annual)^(1/12)−1. Weekly amounts ×52/12, bi-weekly ×26/12.",
            "Order of operations each month: contributions in → price growth on (opening+contribution) → dividends on (opening+contribution) → withdrawal funded (dividends/cash first, then pro-rata asset sale per funding order) → leftover dividends reinvested (or held as 0%-cash if DRIP = No).",
            "Escalations/indexing step once every 12 months from their own start month.",
            "'Annual' withdrawal frequency = full amount taken every 12th month starting at the start month.",
            "ACB tracking (non-registered): opening ACB = ACB% × opening non-reg value; contributions & reinvested dividends add at full value; sales remove ACB in proportion to market value sold. Realized gain = proceeds × (1 − ACB/market value). A negative result is a capital loss; annual taxable gains are floored at zero (losses are not carried between years).",
            "After the horizon's final month the engine freezes (no growth) so annual sums are unaffected.",
            "Charts have a fixed 50-year / 600-month axis; series simply stop at your horizon (#N/A beyond it is intentional and only appears on the hidden ChartData sheet).",
        ]),
        ("TAX CONVENTIONS", [
            "All non-registered dividends are taxed in the year generated, whether reinvested or withdrawn. TFSA is fully tax-free. Every dollar out of RRSP/RRIF (dividends used or sales) is ordinary income.",
            "Federal BPA phases down between $181,440 and $258,482 of taxable income; credits cannot push tax below zero.",
            "Ontario: surtax is applied to Ontario tax after the BPA credit; the Ontario dividend tax credit is applied after the surtax (standard T3-ON ordering).",
            "The 'stacked' view treats Other annual taxable income as fully ordinary income (e.g. salary) and shows only the incremental tax the portfolio causes.",
            "Tax-by-source chart shows the INCREMENTAL tax stacked on top of Other annual taxable income, attributed sequentially (interest → RRSP → capital gains → dividends); the dividend layer can be negative when dividend tax credits exceed the tax on the dividends themselves. With Other income = 0 it equals the standalone view.",
        ]),
        ("EXCLUDED BY DESIGN", [
            "AMT, Ontario Health Premium (up to $900/yr), OAS clawback, CPP/EI, RRIF minimum-withdrawal rules, TFSA/RRSP contribution-room limits, foreign withholding tax, capital-loss carryovers, RRSP deduction for contributions.",
            "A large single-year sale may trigger Alternative Minimum Tax — confirm with an accountant before acting on any single-year figure.",
        ]),
    ]
    r += 2
    for title, lines in notes:
        ws.cell(row=r, column=1, value=title).font = Font(bold=True, color=NAVY)
        r += 1
        for ln in lines:
            cc = ws.cell(row=r, column=1, value="• " + ln)
            cc.alignment = Alignment(wrap_text=True, vertical="top")
            cc.font = F_LABEL
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
            ws.row_dimensions[r].height = 28
            r += 1
        r += 1
    ws.protection.sheet = True
    return ws


# ----------------------------------------------------------------------------- tax tab

def _fed_bracket(inc):
    return f"SUMPRODUCT(({inc}>FedThr)*(({inc})-FedThr)*FedDlt)"


def _on_bracket(inc):
    return f"SUMPRODUCT(({inc}>ONThr)*(({inc})-ONThr)*ONDlt)"


def _fed_bpa(inc):
    return (f"(FedBPAMax-(FedBPAMax-FedBPAMin)*MIN(1,MAX(0,(({inc})-FedPhLow)/(FedPhHigh-FedPhLow))))")


def _fed_tax_plain(inc):
    """federal tax, no dividend credits (single formula)"""
    return f"MAX(0,{_fed_bracket(inc)}-{_fed_bpa(inc)}*FedLowRate)"


def _on_basic_plain(inc):
    return f"MAX(0,{_on_bracket(inc)}-ONBPA*ONLowRate)"


def build_tax(wb):
    ws = wb.create_sheet("Tax Estimate")
    ws.sheet_properties.tabColor = "72B7B2"
    ws.freeze_panes = "C4"
    ws.sheet_view.showGridLines = False
    E = E_BASE

    ws["A1"] = "Tax Estimate — CRA federal + Ontario, tax year 2026 rules"
    ws["A1"].font = F_TITLE
    ws["A2"] = ("Standalone = portfolio is the only income.  Stacked = portfolio on top of other income "
                "(shows the incremental tax only).  Excluded by design: AMT, Ontario Health Premium, OAS clawback, CPP/EI — "
                "a large single-year sale may trigger AMT; confirm with an accountant.")
    ws["A2"].font = F_SMALL
    ws.merge_cells("A2:T2")

    # tax on other income alone (for the stacked view)
    ws["V1"] = "Tax on other income alone:"
    ws["V1"].font = F_SMALL
    ws["X1"] = f"={_fed_tax_plain('N(OtherInc)')}"
    ws["Y1"] = f"={_on_basic_plain('N(OtherInc)')}"
    ws["Z1"] = "=$X$1+$Y$1+ONSurR1*MAX(0,$Y$1-ONSur1)+ONSurR2*MAX(0,$Y$1-ONSur2)"
    for a in ("X1", "Y1", "Z1"):
        ws[a].number_format = MONEY
    defname(wb, "BaseTaxTotal", "'Tax Estimate'!$Z$1")

    cols = [
        ("A", "Year"), ("B", "·"),
        ("C", "Eligible dividends (NR)"), ("D", "Non-eligible dividends (NR)"),
        ("E", "Foreign / interest (NR)"), ("F", "Realized capital gains"),
        ("G", "RRSP/RRIF withdrawals"), ("H", "TFSA withdrawals (tax-free)"),
        ("I", "Grossed-up eligible"), ("J", "Grossed-up non-eligible"),
        ("K", "Taxable income (standalone)"), ("L", "Fed tax before credits"),
        ("M", "Fed BPA"), ("N", "Federal tax"), ("O", "ON tax before credits"),
        ("P", "ON basic tax"), ("Q", "ON surtax"), ("R", "Ontario tax"),
        ("S", "TOTAL TAX (standalone)"), ("T", "Effective rate on cash income"),
        ("U", "Taxable income (stacked)"), ("V", "Stk fed before credits"), ("W", "Stk fed BPA"),
        ("X", "Stk federal tax"), ("Y", "Stk ON before credits"), ("Z", "Stk ON basic"),
        ("AA", "Stk ON surtax"), ("AB", "Stk Ontario tax"), ("AC", "Stk total (all income)"),
        ("AD", "INCREMENTAL TAX (stacked view)"),
        ("AE", "T·A fed"), ("AF", "T·A ON basic"), ("AG", "T·A total"),
        ("AH", "T·B fed"), ("AI", "T·B ON basic"), ("AJ", "T·B total"),
        ("AK", "T·C fed"), ("AL", "T·C ON basic"), ("AM", "T·C total"),
        ("AN", "Tax: interest/foreign"), ("AO", "Tax: RRSP withdrawals"),
        ("AP", "Tax: capital gains"), ("AQ", "Tax: dividends (net of credits)"),
    ]
    for L, h in cols:
        c = ws[f"{L}3"]
        c.value = h
        c.font = Font(bold=True, size=9, color="FFFFFF")
        c.fill = FILL_H2
        c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.column_dimensions[L].width = 13
    ws.column_dimensions["A"].width = 6
    ws.column_dimensions["B"].width = 3
    ws.row_dimensions[3].height = 42

    yr_rng = eref(E, "YR", rng=(ENG_FIRST, ENG_LAST))

    def sumy(col, r):
        return f"SUMIFS({eref(E, col, rng=(ENG_FIRST, ENG_LAST))},{yr_rng},$A{r})"

    for y in range(1, MAXY + 1):
        r = y + 3
        ws[f"A{r}"] = y
        ws[f"B{r}"] = f"=IF($A{r}<=HorizonY,1,0)"
        ws[f"C{r}"] = f"={sumy('DIVEL', r)}"
        ws[f"D{r}"] = f"={sumy('DIVNE', r)}"
        ws[f"E{r}"] = f"={sumy('DIVFOR', r)}"
        ws[f"F{r}"] = f"=MAX(0,{sumy('GAIN', r)})"
        ws[f"G{r}"] = f"={sumy('DRAW_RR', r)}"
        ws[f"H{r}"] = f"={sumy('DRAW_TF', r)}"
        ws[f"I{r}"] = f"=$C{r}*(1+GUE)"
        ws[f"J{r}"] = f"=$D{r}*(1+GUN)"
        ws[f"K{r}"] = f"=$I{r}+$J{r}+$E{r}+CGInc*$F{r}+$G{r}"
        ws[f"L{r}"] = f"={_fed_bracket(f'$K{r}')}"
        ws[f"M{r}"] = f"={_fed_bpa(f'$K{r}')}"
        ws[f"N{r}"] = f"=MAX(0,$L{r}-$M{r}*FedLowRate-$I{r}*FDTCE-$J{r}*FDTCN)"
        ws[f"O{r}"] = f"={_on_bracket(f'$K{r}')}"
        ws[f"P{r}"] = f"=MAX(0,$O{r}-ONBPA*ONLowRate)"
        ws[f"Q{r}"] = f"=ONSurR1*MAX(0,$P{r}-ONSur1)+ONSurR2*MAX(0,$P{r}-ONSur2)"
        ws[f"R{r}"] = f"=MAX(0,$P{r}+$Q{r}-$I{r}*ODTCE-$J{r}*ODTCN)"
        ws[f"S{r}"] = f"=$N{r}+$R{r}"
        ws[f"T{r}"] = f"=IF($C{r}+$D{r}+$E{r}+$F{r}+$G{r}>0,$S{r}/($C{r}+$D{r}+$E{r}+$F{r}+$G{r}),0)"
        ws[f"U{r}"] = f"=$K{r}+N(OtherInc)"
        ws[f"V{r}"] = f"={_fed_bracket(f'$U{r}')}"
        ws[f"W{r}"] = f"={_fed_bpa(f'$U{r}')}"
        ws[f"X{r}"] = f"=MAX(0,$V{r}-$W{r}*FedLowRate-$I{r}*FDTCE-$J{r}*FDTCN)"
        ws[f"Y{r}"] = f"={_on_bracket(f'$U{r}')}"
        ws[f"Z{r}"] = f"=MAX(0,$Y{r}-ONBPA*ONLowRate)"
        ws[f"AA{r}"] = f"=ONSurR1*MAX(0,$Z{r}-ONSur1)+ONSurR2*MAX(0,$Z{r}-ONSur2)"
        ws[f"AB{r}"] = f"=MAX(0,$Z{r}+$AA{r}-$I{r}*ODTCE-$J{r}*ODTCN)"
        ws[f"AC{r}"] = f"=$X{r}+$AB{r}"
        ws[f"AD{r}"] = f"=IF($B{r}=0,0,$AC{r}-BaseTaxTotal)"
        # attribution scenarios, stacked on other income (no dividends → no DTC):
        # A = other+interest, B = +RRSP, C = +gains; dividend layer = stacked total − C
        incA = f"N(OtherInc)+$E{r}"
        incB = f"N(OtherInc)+$E{r}+$G{r}"
        incC = f"N(OtherInc)+$E{r}+$G{r}+CGInc*$F{r}"
        for (LF, LB, LT), inc in [(("AE", "AF", "AG"), incA), (("AH", "AI", "AJ"), incB),
                                  (("AK", "AL", "AM"), incC)]:
            ws[f"{LF}{r}"] = f"={_fed_tax_plain(inc)}"
            ws[f"{LB}{r}"] = f"={_on_basic_plain(inc)}"
            ws[f"{LT}{r}"] = (f"=${LF}{r}+${LB}{r}+ONSurR1*MAX(0,${LB}{r}-ONSur1)"
                              f"+ONSurR2*MAX(0,${LB}{r}-ONSur2)")
        ws[f"AN{r}"] = f"=IF($B{r}=0,0,$AG{r}-BaseTaxTotal)"
        ws[f"AO{r}"] = f"=$AJ{r}-$AG{r}"
        ws[f"AP{r}"] = f"=$AM{r}-$AJ{r}"
        ws[f"AQ{r}"] = f"=$AC{r}-$AM{r}"
        for L, _ in cols:
            if L in ("A", "B"):
                continue
            ws[f"{L}{r}"].number_format = PCT if L == "T" else MONEY
    # hide helper columns
    for L in ["I", "J", "L", "M", "O", "V", "W", "Y", "AE", "AF", "AG", "AH", "AI", "AJ", "AK", "AL", "AM"]:
        ws.column_dimensions[L].hidden = True
    ws.column_dimensions["B"].hidden = True

    tips_r = MAXY + 6
    ws.cell(row=tips_r, column=1, value="General strategies — not tax advice").font = Font(bold=True, color=NAVY, size=12)
    tips = [
        "Prioritize TFSA room — everything inside is tax-free, including withdrawals.",
        "Hold interest-bearing assets (bonds, GICs, foreign dividends) inside registered accounts; hold eligible-dividend payers in non-registered accounts where the dividend tax credit applies.",
        "Spread large sales across tax years to stay in lower brackets and reduce AMT risk.",
        "RRSP contributions in high-income years can offset RRSP/RRIF withdrawal income.",
        "Harvest capital losses in non-registered accounts to offset realized gains.",
        "Spousal RRSPs and pension income splitting can shift income to a lower-bracket spouse.",
    ]
    for i, t in enumerate(tips):
        cc = ws.cell(row=tips_r + 1 + i, column=1, value="• " + t)
        cc.alignment = Alignment(wrap_text=True)
        ws.merge_cells(start_row=tips_r + 1 + i, start_column=1, end_row=tips_r + 1 + i, end_column=8)
    ws.protection.sheet = True
    return ws


# ------------------------------------------------------------------------- projection

def build_projection(wb):
    ws = wb.create_sheet("Projection")
    ws.sheet_properties.tabColor = "4C78A8"
    ws.freeze_panes = "C4"
    ws.sheet_view.showGridLines = False
    E = E_BASE

    ws["A1"] = "Projection — annual summary (all figures trace to the monthly Engine)"
    ws["A1"].font = F_TITLE

    groups = [("Non-Registered", "C"), ("TFSA", "I"), ("RRSP/RRIF", "O"), ("Total", "U")]
    for gname, start in groups:
        c0 = openpyxl.utils.column_index_from_string(start)
        span = 7 if gname == "Total" else 6
        ws.cell(row=2, column=c0, value=gname).font = Font(bold=True, color=NAVY)
        ws.merge_cells(start_row=2, start_column=c0, end_row=2, end_column=c0 + span - 1)
    ws.cell(row=2, column=28, value="Withdrawals Summary (cumulative)").font = Font(bold=True, color=NAVY)
    ws.merge_cells(start_row=2, start_column=28, end_row=2, end_column=32)

    sub = ["Opening", "Contributions", "Growth", "Dividends", "Withdrawn", "Closing"]
    headers = (["Year", "·"] + sub * 3 +
               ["Opening", "Contributions", "Growth", "Dividends", "Withdrawn", "Tax (standalone)", "Closing"] +
               ["Total withdrawn", "Funded by dividends/cash", "Funded by asset sales", "Realized gains", "Status"])
    for j, h in enumerate(headers):
        c = ws.cell(row=3, column=1 + j, value=h)
        c.font = Font(bold=True, size=9, color="FFFFFF")
        c.fill = FILL_H2
        c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.column_dimensions[gcl(1 + j)].width = 13
    ws.column_dimensions["A"].width = 6
    ws.column_dimensions["B"].width = 3
    ws.column_dimensions["B"].hidden = True
    ws.row_dimensions[3].height = 40

    yr_rng = eref(E, "YR", rng=(ENG_FIRST, ENG_LAST))

    def sumy(col, r):
        return f"SUMIFS({eref(E, col, rng=(ENG_FIRST, ENG_LAST))},{yr_rng},$A{r})"

    def cum(col, r):
        return f'SUMIFS({eref(E, col, rng=(ENG_FIRST, ENG_LAST))},{yr_rng},"<="&$A{r})'

    for y in range(1, MAXY + 1):
        r = y + 3
        ws[f"A{r}"] = y
        ws[f"B{r}"] = f"=IF($A{r}<=HorizonY,1,0)"

        def w(L, core):
            ws[f"{L}{r}"] = f'=IF($B{r}=0,"",{core})'
            ws[f"{L}{r}"].number_format = MONEY

        for a, start in [("NR", "C"), ("TF", "I"), ("RR", "O")]:
            c0 = openpyxl.utils.column_index_from_string(start)
            bal_rng = eref(E, f"BAL_{a}", rng=(ROW0, ENG_LAST))
            w(gcl(c0), f"INDEX({bal_rng},MIN(($A{r}-1)*12,HorizonM)+1)")
            w(gcl(c0 + 1), sumy(f"CONTRIB_{a}", r))
            w(gcl(c0 + 2), sumy(f"GROWTH_{a}", r))
            w(gcl(c0 + 3), sumy(f"DIV_{a}", r))
            w(gcl(c0 + 4), sumy(f"DRAW_{a}", r))
            w(gcl(c0 + 5), f"INDEX({bal_rng},MIN($A{r}*12,HorizonM)+1)")
        # totals
        w("U", f"$C{r}+$I{r}+$O{r}")
        w("V", f"$D{r}+$J{r}+$P{r}")
        w("W", f"$E{r}+$K{r}+$Q{r}")
        w("X", f"$F{r}+$L{r}+$R{r}")
        w("Y", f"$G{r}+$M{r}+$S{r}")
        w("Z", f"'Tax Estimate'!$S{r}")
        w("AA", f"$H{r}+$N{r}+$T{r}")
        w("AB", cum("DRAWTOT", r))
        w("AC", cum("FCTOT", r))
        w("AD", cum("SALETOT", r))
        w("AE", cum("GAIN", r))
        ws[f"AF{r}"] = (f'=IF($B{r}=0,"",IF(COUNTIFS({yr_rng},$A{r},'
                        f'{eref(E, "UNMET", rng=(ENG_FIRST, ENG_LAST))},1)>0,"⚠ shortfall","OK"))')
    ws.protection.sheet = True
    return ws


# -------------------------------------------------------------------------- chart data

def build_chartdata(wb):
    ws = wb.create_sheet("ChartData")
    ws.sheet_state = "hidden"
    E = E_BASE

    ws["A1"], ws["B1"], ws["C1"], ws["D1"] = "Month", "Balance (plan)", "Balance (no withdrawals)", "Withdrawals begin"
    bal_rng = eref(E, "BALTOT", rng=(ROW0, ENG_LAST))
    bal_nowd = eref(E_NOWD, "BALTOT", rng=(ROW0, ENG_LAST))
    for m in range(0, MAXM + 1):
        r = m + 2
        ws[f"A{r}"] = m
        ws[f"B{r}"] = f"=IF($A{r}<=HorizonM,INDEX({bal_rng},$A{r}+1),NA())"
        ws[f"C{r}"] = f"=IF($A{r}<=HorizonM,INDEX({bal_nowd},$A{r}+1),NA())"
        ws[f"D{r}"] = f"=IF(AND(N(WdAmt)>0,$A{r}=WdStartEff),$B{r},NA())"

    hdr = ["Year", "Dividends/cash used", "Asset sales", "Contributions",
           "Tax: interest/foreign", "Tax: RRSP", "Tax: capital gains", "Tax: dividends",
           "Non-Registered", "TFSA", "RRSP/RRIF"]
    for j, h in enumerate(hdr):
        ws.cell(row=1, column=6 + j, value=h)  # F..P
    yr_rng = eref(E, "YR", rng=(ENG_FIRST, ENG_LAST))
    for y in range(1, MAXY + 1):
        r = y + 1
        ws[f"F{r}"] = y
        act = f"$F{r}<=HorizonY"
        ws[f"G{r}"] = f"=IF({act},SUMIFS({eref(E,'FCTOT',rng=(ENG_FIRST,ENG_LAST))},{yr_rng},$F{r}),NA())"
        ws[f"H{r}"] = f"=IF({act},SUMIFS({eref(E,'SALETOT',rng=(ENG_FIRST,ENG_LAST))},{yr_rng},$F{r}),NA())"
        ws[f"I{r}"] = f"=IF({act},SUMIFS({eref(E,'CT',rng=(ENG_FIRST,ENG_LAST))},{yr_rng},$F{r}),NA())"
        ws[f"J{r}"] = f"=IF({act},'Tax Estimate'!$AN{y+3},NA())"
        ws[f"K{r}"] = f"=IF({act},'Tax Estimate'!$AO{y+3},NA())"
        ws[f"L{r}"] = f"=IF({act},'Tax Estimate'!$AP{y+3},NA())"
        ws[f"M{r}"] = f"=IF({act},'Tax Estimate'!$AQ{y+3},NA())"
        for L, a in [("N", "NR"), ("O", "TF"), ("P", "RR")]:
            ws[f"{L}{r}"] = (f"=IF({act},INDEX({eref(E, f'BAL_{a}', rng=(ROW0, ENG_LAST))},"
                             f"MIN($F{r}*12,HorizonM)+1),NA())")
    return ws


# --------------------------------------------------------------------------- dashboard

def style_chart(ch, title, w=17, h=9):
    ch.title = title
    ch.width = w
    ch.height = h
    ch.style = 2
    if ch.y_axis:
        ch.y_axis.numFmt = '$#,##0'
        ch.y_axis.majorGridlines = None
    ch.legend.position = "b"


def paint(series, color, line=False):
    from openpyxl.chart.shapes import GraphicalProperties
    gp = GraphicalProperties(solidFill=color)
    if line:
        from openpyxl.drawing.line import LineProperties
        gp = GraphicalProperties()
        gp.line = LineProperties(solidFill=color, w=22000)
    series.graphicalProperties = gp


def build_dashboard(wb):
    ws = wb.create_sheet("Dashboard", 0)
    ws.sheet_properties.tabColor = NAVY
    ws.sheet_view.showGridLines = False
    widths = {"A": 2, "B": 20, "C": 16, "D": 16, "E": 14, "F": 20, "G": 16, "H": 16,
              "I": 3, "J": 20, "K": 16, "L": 16, "M": 3}
    for c, w in widths.items():
        ws.column_dimensions[c].width = w

    ws["B1"] = "Portfolio, Withdrawal & Tax Planner — Dashboard"
    ws["B1"].font = Font(size=18, bold=True, color=NAVY)
    ws["B2"] = "Canada (Ontario) · monthly engine · tax year 2026 rules. Edit the yellow cells on the Inputs tab; everything here updates automatically."
    ws["B2"].font = F_SMALL

    # year selector
    ws["B4"] = "Selected year"
    ws["B4"].font = Font(bold=True, size=12)
    ws["C4"] = 8
    ws["C4"].fill = FILL_IN
    ws["C4"].border = BOX
    ws["C4"].font = Font(bold=True, size=12)
    ws["C4"].protection = UNLOCKED
    defname(wb, "SelYear", "Dashboard!$C$4")
    ws["D4"] = '="of "&HorizonY&" years"'
    ws["D4"].font = F_LABEL
    ws["M4"] = "=MIN(MAX(1,N(SelYear)),HorizonY)"   # effective year (hidden helper)
    ws["M4"].font = F_SMALL
    defname(wb, "SelYearEff", "Dashboard!$M$4")
    dv = DataValidation(type="whole", operator="between", formula1="1", formula2="50",
                        allow_blank=False, error="Enter a year between 1 and 50.")
    ws.add_data_validation(dv)
    dv.add("C4")

    # quick-edit mirrors
    ws["F4"] = "Key inputs (change on the Inputs tab):"
    ws["F4"].font = F_SMALL
    ws["F5"] = '="Withdrawal: "&TEXT(N(WdAmt),"$#,##0")&" "&WdFreq&" from month "&WdStartEff'
    ws["F6"] = '="Horizon: "&HorizonY&" yrs · Return override: "&TEXT(N(RetOverride),"+0.0%;-0.0%;0%")'
    ws["F5"].font = F_SMALL
    ws["F6"].font = F_SMALL
    ws.merge_cells("F5:L5")
    ws.merge_cells("F6:L6")

    # ---- KPI cards (rows 8-12 and 14-18)
    P = "Projection"
    cum_tax = f"SUMIFS('Tax Estimate'!$S$4:$S$53,'Tax Estimate'!$A$4:$A$53,\"<=\"&SelYearEff)"
    cum_tax_stk = f"SUMIFS('Tax Estimate'!$AD$4:$AD$53,'Tax Estimate'!$A$4:$A$53,\"<=\"&SelYearEff)"
    cum_wd = f"INDEX('{P}'!$AB$4:$AB$53,SelYearEff)"
    cum_contrib = f"SUMIFS('{P}'!$V$4:$V$53,'{P}'!$A$4:$A$53,\"<=\"&SelYearEff)"
    end_bal = f"INDEX('{P}'!$AA$4:$AA$53,SelYearEff)"
    unmet_rng = eref(E_BASE, "UNMET", rng=(ENG_FIRST, ENG_LAST))

    cards = [
        ("B8", "PORTFOLIO BALANCE", f"={end_bal}", MONEY,
         '="at end of year "&SelYearEff'),
        ("F8", "TOTAL WITHDRAWN", f"={cum_wd}", MONEY,
         f"=\"to date · \"&TEXT(INDEX('{P}'!$Y$4:$Y$53,SelYearEff),\"$#,##0\")&\" in year \"&SelYearEff"),
        ("J8", "TOTAL TAX PAID (standalone)", f"={cum_tax}", MONEY,
         f'="to date · stacked view adds "&TEXT({cum_tax_stk},"$#,##0")&" on top of other income"'),
        ("B14", "DIVIDENDS RECEIVED", f"=SUMIFS('{P}'!$X$4:$X$53,'{P}'!$A$4:$A$53,\"<=\"&SelYearEff)", MONEY,
         '="all accounts, to end of year "&SelYearEff'),
        ("F14", "NET AFTER-TAX RETURN (annualized)",
         f"=IFERROR(POWER(({end_bal}+{cum_wd}-{cum_tax})/(N(StartValue)+{cum_contrib}),1/SelYearEff)-1,0)", PCT,
         '="growth + dividends − tax, on money invested, to year "&SelYearEff'),
        ("J14", "PORTFOLIO LONGEVITY",
         f'=IFERROR("Depletes in year "&(INT((MATCH(1,{unmet_rng},0)-1)/12)+1)&" ⚠",'
         f'"Lasts the full horizon ✅")', None,
         '="the single most important wellbeing number"'),
    ]
    for anchor, label, formula, fmt, ctx in cards:
        col = anchor[0]
        row = int(anchor[1:])
        c0 = openpyxl.utils.column_index_from_string(col)
        for rr in range(row, row + 4):
            for cc in range(c0, c0 + 3):
                cell = ws.cell(row=rr, column=cc)
                cell.fill = FILL_CARD
                cell.border = Border()
        for cc in range(c0, c0 + 3):
            ws.cell(row=row, column=cc).border = Border(top=Side(style="medium", color=NAVY))
        ws.cell(row=row, column=c0, value=label).font = Font(size=9, bold=True, color=NAVY_LIGHT)
        v = ws.cell(row=row + 1, column=c0, value=formula)
        v.font = F_KPI if fmt else Font(size=13, bold=True, color=NAVY)
        if fmt:
            v.number_format = fmt
        ws.merge_cells(start_row=row + 1, start_column=c0, end_row=row + 2, end_column=c0 + 2)
        v.alignment = Alignment(vertical="center")
        ws.cell(row=row + 3, column=c0, value=ctx).font = F_SMALL
        ws.merge_cells(start_row=row + 3, start_column=c0, end_row=row + 3, end_column=c0 + 2)

    # ---- scenario mini-table
    ws["B20"] = "Sensitivity — expected return ±1%"
    ws["B20"].font = Font(bold=True, size=12, color=NAVY)
    for j, h in enumerate(["", "Return −1%", "Base plan", "Return +1%"]):
        ws.cell(row=21, column=2 + j, value=h).font = Font(bold=True, size=10)
    ws["B22"] = "Ending balance"
    ws["B23"] = "Depletes"
    for j, eng in enumerate([E_M1, E_BASE, E_P1]):
        bal = eref(eng, "BALTOT", rng=(ROW0, ENG_LAST))
        unm = eref(eng, "UNMET", rng=(ENG_FIRST, ENG_LAST))
        c = ws.cell(row=22, column=3 + j, value=f"=INDEX({bal},HorizonM+1)")
        c.number_format = MONEY
        c.font = Font(bold=True)
        ws.cell(row=23, column=3 + j,
                value=f'=IFERROR("Year "&(INT((MATCH(1,{unm},0)-1)/12)+1)&" ⚠","Never (lasts)")')
    for rr in range(21, 24):
        for cc in range(2, 6):
            ws.cell(row=rr, column=cc).border = BOX

    # ---- charts
    cd = wb["ChartData"]

    ch1 = LineChart()
    style_chart(ch1, "Portfolio balance — plan vs no withdrawals", w=19, h=10)
    data = Reference(cd, min_col=2, max_col=4, min_row=1, max_row=MAXM + 2)
    cats = Reference(cd, min_col=1, min_row=2, max_row=MAXM + 2)
    ch1.add_data(data, titles_from_data=True)
    ch1.set_categories(cats)
    ch1.x_axis.title = "Month"
    for s, colr in zip(ch1.series, [SERIES[0], SERIES[2], SERIES[4]]):
        paint(s, colr, line=True)
        s.smooth = False
    mk = ch1.series[2]
    mk.marker = Marker(symbol="circle", size=8)
    from openpyxl.drawing.line import LineProperties
    from openpyxl.chart.shapes import GraphicalProperties
    mk.graphicalProperties = GraphicalProperties()
    mk.graphicalProperties.line = LineProperties(noFill=True)
    ws.add_chart(ch1, "B26")

    ch2 = BarChart()
    ch2.type = "col"
    ch2.grouping = "stacked"
    ch2.overlap = 100
    style_chart(ch2, "Annual income & funding composition")
    data = Reference(cd, min_col=7, max_col=9, min_row=1, max_row=MAXY + 1)
    cats = Reference(cd, min_col=6, min_row=2, max_row=MAXY + 1)
    ch2.add_data(data, titles_from_data=True)
    ch2.set_categories(cats)
    ch2.x_axis.title = "Year"
    for s, colr in zip(ch2.series, [SERIES[0], SERIES[1], SERIES[2]]):
        paint(s, colr)
    ws.add_chart(ch2, "J26")

    ch3 = BarChart()
    ch3.type = "col"
    ch3.grouping = "stacked"
    ch3.overlap = 100
    style_chart(ch3, "Annual tax by source (on top of other income)")
    data = Reference(cd, min_col=10, max_col=13, min_row=1, max_row=MAXY + 1)
    cats = Reference(cd, min_col=6, min_row=2, max_row=MAXY + 1)
    ch3.add_data(data, titles_from_data=True)
    ch3.set_categories(cats)
    ch3.x_axis.title = "Year"
    for s, colr in zip(ch3.series, [SERIES[2], SERIES[3], SERIES[1], SERIES[0]]):
        paint(s, colr)
    ws.add_chart(ch3, "B47")

    ch4 = AreaChart()
    ch4.grouping = "stacked"
    style_chart(ch4, "Account mix over time (year-end balances)")
    data = Reference(cd, min_col=14, max_col=16, min_row=1, max_row=MAXY + 1)
    cats = Reference(cd, min_col=6, min_row=2, max_row=MAXY + 1)
    ch4.add_data(data, titles_from_data=True)
    ch4.set_categories(cats)
    ch4.x_axis.title = "Year"
    for s, colr in zip(ch4.series, [SERIES[0], SERIES[2], SERIES[1]]):
        paint(s, colr)
    ws.add_chart(ch4, "J47")

    ws.protection.sheet = True
    return ws


# --------------------------------------------------------------------------------- main

def main(path):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    build_inputs(wb)
    build_assumptions(wb)
    build_engine(wb, E_BASE, "DeltaBase", with_wd=True, hidden=False)
    build_engine(wb, E_NOWD, "DeltaBase", with_wd=False, hidden=True)
    build_engine(wb, E_M1, "DeltaM1", with_wd=True, hidden=True)
    build_engine(wb, E_P1, "DeltaP1", with_wd=True, hidden=True)
    build_tax(wb)
    build_projection(wb)
    build_chartdata(wb)
    build_dashboard(wb)

    order = ["Dashboard", "Inputs", "Projection", "Tax Estimate", E_BASE,
             "Assumptions & Notes", E_NOWD, E_M1, E_P1, "ChartData"]
    wb._sheets = [wb[n] for n in order]
    wb.active = 0
    wb.save(path)
    print(f"wrote {path}")


if __name__ == "__main__":
    import sys
    main(sys.argv[1] if len(sys.argv) > 1 else "Portfolio_Planner_CA.xlsx")
