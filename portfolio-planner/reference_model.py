#!/usr/bin/env python3
"""
Independent Python re-implementation of the workbook's monthly engine and tax
pipeline, used to verify the spreadsheet formulas (see verify_workbook.py).
Semantics intentionally mirror the formulas 1:1.
"""

from dataclasses import dataclass, field

ACCTS = ["NR", "TF", "RR"]

# 2026 tax parameters (must match build_workbook.py)
FED_BRACKETS = [(0, 0.14), (58523, 0.205), (117045, 0.26), (181440, 0.29), (258482, 0.33)]
FED_BPA_MAX, FED_BPA_MIN = 16452, 14829
FED_PHASE_LOW, FED_PHASE_HIGH = 181440, 258482
ON_BRACKETS = [(0, 0.0505), (53891, 0.0915), (107785, 0.1116), (150000, 0.1216), (220000, 0.1316)]
ON_BPA = 12989
ON_SUR1, ON_SUR2 = 5818, 7446
ON_SURR1, ON_SURR2 = 0.20, 0.36
GUE, GUN = 0.38, 0.15
FDTCE, FDTCN = 0.150198, 0.090301
ODTCE, ODTCN = 0.10, 0.029863
CG_INC = 0.50


@dataclass
class Sub:
    name: str
    alloc: float
    ret: float
    yld: float
    dtype: str          # "Eligible" | "Non-eligible" | "Foreign / interest"
    acct: dict          # {"NR": .., "TF": .., "RR": ..}


@dataclass
class Inputs:
    start_value: float = 500_000
    horizon_m: int = 240
    inflation: float = 0.02
    ret_override: float = 0.0
    subs: list = field(default_factory=lambda: [
        Sub("Canadian Dividend Stocks", 0.30, 0.040, 0.040, "Eligible",
            {"NR": 0.60, "TF": 0.20, "RR": 0.20}),
        Sub("US / Global Growth Equities", 0.30, 0.070, 0.012, "Foreign / interest",
            {"NR": 0.40, "TF": 0.30, "RR": 0.30}),
        Sub("Bonds & GICs", 0.25, 0.015, 0.035, "Foreign / interest",
            {"NR": 0.20, "TF": 0.20, "RR": 0.60}),
        Sub("Canadian REITs", 0.15, 0.030, 0.045, "Non-eligible",
            {"NR": 0.30, "TF": 0.40, "RR": 0.30}),
        Sub("", 0, 0, 0, "Eligible", {"NR": 0, "TF": 0, "RR": 0}),
        Sub("", 0, 0, 0, "Eligible", {"NR": 0, "TF": 0, "RR": 0}),
    ])
    contrib_amt: float = 500
    contrib_freq: str = "Bi-weekly"     # Weekly | Bi-weekly | Monthly
    contrib_start: int = 1
    contrib_end: int = 60               # 0/None -> horizon
    contrib_esc: float = 0.03
    contrib_dest: str = "Follow allocation %"
    wd_amt: float = 2000
    wd_freq: str = "Monthly"            # Bi-weekly | Monthly | Annual
    wd_start: int = 61
    wd_end: int = 0                     # 0 -> horizon
    wd_indexed: bool = True
    order: tuple = ("NR", "TF", "RR")
    other_inc: float = 60_000
    acb_pct: float = 0.70
    drip: bool = True
    delta: float = 0.0                  # scenario shift, added to ret_override
    with_wd: bool = True


def simulate(inp: Inputs):
    H = inp.horizon_m
    delta = inp.ret_override + inp.delta
    retm = [pow(1 + s.ret + delta, 1 / 12) for s in inp.subs]           # multiplier
    yldm = [pow(1 + s.yld, 1 / 12) - 1 for s in inp.subs]
    ds = [(s.alloc if inp.contrib_dest == "Follow allocation %" else
           (1.0 if (s.name and inp.contrib_dest == s.name) else 0.0)) for s in inp.subs]
    cm = inp.contrib_amt * {"Weekly": 52 / 12, "Bi-weekly": 26 / 12, "Monthly": 1}[inp.contrib_freq]
    c_start = max(1, inp.contrib_start or 1)
    c_end = inp.contrib_end if (inp.contrib_end or 0) >= 1 else H
    c_end = min(c_end, H)
    wm = inp.wd_amt * (26 / 12 if inp.wd_freq == "Bi-weekly" else 1)
    w_start = max(1, inp.wd_start or 1)
    w_end = inp.wd_end if (inp.wd_end or 0) >= 1 else H
    w_end = min(w_end, H)
    rank = {a: inp.order.index(a) + 1 for a in ACCTS}

    K = {(a, i): inp.start_value * inp.subs[i].alloc * inp.subs[i].acct[a]
         for a in ACCTS for i in range(6)}
    cash = {a: 0.0 for a in ACCTS}
    acb = inp.acb_pct * sum(K[("NR", i)] for i in range(6))

    months = []   # dicts of per-month aggregates
    rows = 600
    for t in range(1, rows + 1):
        active = t <= H
        if active and c_start <= t <= c_end:
            C = cm * (1 + inp.contrib_esc) ** int((t - c_start) / 12)
        else:
            C = 0.0
        if inp.with_wd and active and w_start <= t <= w_end and inp.wd_amt > 0:
            base = (inp.wd_amt if ((t - w_start) % 12 == 0) else 0.0) if inp.wd_freq == "Annual" else wm
            E = base * ((1 + inp.inflation) ** int((t - w_start) / 12) if inp.wd_indexed else 1)
        else:
            E = 0.0

        P, D = {}, {}
        for a in ACCTS:
            for i in range(6):
                share = ds[i] * inp.subs[i].acct[a]
                if active:
                    P[(a, i)] = (K[(a, i)] + C * share) * retm[i]
                    D[(a, i)] = (K[(a, i)] + C * share) * yldm[i]
                else:
                    P[(a, i)] = K[(a, i)]
                    D[(a, i)] = 0.0
        PG = {a: sum(P[(a, i)] for i in range(6)) for a in ACCTS}
        DIV = {a: sum(D[(a, i)] for i in range(6)) for a in ACCTS}
        contrib_a = {a: C * sum(ds[i] * inp.subs[i].acct[a] for i in range(6)) for a in ACCTS}
        cap = {a: cash[a] + DIV[a] + PG[a] for a in ACCTS}
        capb = {a: sum(cap[b] for b in ACCTS if rank[b] < rank[a]) for a in ACCTS}
        draw = {a: max(0.0, min(cap[a], E - capb[a])) for a in ACCTS}
        fromcash = {a: min(draw[a], cash[a] + DIV[a]) for a in ACCTS}
        sale = {a: draw[a] - fromcash[a] for a in ACCTS}
        reinv, ncash = {}, {}
        for a in ACCTS:
            leftover = cash[a] + DIV[a] - fromcash[a]
            reinv[a] = leftover if inp.drip else 0.0
            ncash[a] = 0.0 if inp.drip else leftover
        newK = {}
        for a in ACCTS:
            for i in range(6):
                if PG[a] == 0:
                    newK[(a, i)] = P[(a, i)]
                else:
                    newK[(a, i)] = P[(a, i)] * (1 - (sale[a] - reinv[a]) / PG[a])

        acbb = acb + contrib_a["NR"]
        gain = 0.0 if PG["NR"] == 0 else sale["NR"] * (1 - acbb / PG["NR"])
        acb = (acbb if PG["NR"] == 0 else acbb * (1 - sale["NR"] / PG["NR"])) + reinv["NR"]

        divel = sum(D[("NR", i)] for i in range(6) if inp.subs[i].dtype == "Eligible")
        divne = sum(D[("NR", i)] for i in range(6) if inp.subs[i].dtype == "Non-eligible")
        divfor = sum(D[("NR", i)] for i in range(6) if inp.subs[i].dtype == "Foreign / interest")

        K, cash = newK, ncash
        bal = {a: sum(K[(a, i)] for i in range(6)) + cash[a] for a in ACCTS}
        months.append(dict(
            t=t, year=(t - 1) // 12 + 1, active=active, C=C, E=E,
            div=dict(DIV), draw=dict(draw), fromcash=dict(fromcash), sale=dict(sale),
            gain=gain, divel=divel, divne=divne, divfor=divfor,
            bal=dict(bal), baltot=sum(bal.values()),
            drawtot=sum(draw.values()), fctot=sum(fromcash.values()), saletot=sum(sale.values()),
            unmet=1 if (active and E - sum(draw.values()) > 0.01) else 0,
        ))
    return months


# ------------------------------------------------------------------- tax pipeline

def bracket_tax(inc, brackets):
    tax, prev_rate = 0.0, 0.0
    for lo, rate in brackets:
        if inc > lo:
            tax += (inc - lo) * (rate - prev_rate)
        prev_rate = rate
    return tax


def fed_bpa(inc):
    frac = min(1.0, max(0.0, (inc - FED_PHASE_LOW) / (FED_PHASE_HIGH - FED_PHASE_LOW)))
    return FED_BPA_MAX - (FED_BPA_MAX - FED_BPA_MIN) * frac


def fed_tax(inc, gu_e=0.0, gu_n=0.0):
    return max(0.0, bracket_tax(inc, FED_BRACKETS) - fed_bpa(inc) * FED_BRACKETS[0][1]
               - gu_e * FDTCE - gu_n * FDTCN)


def on_tax(inc, gu_e=0.0, gu_n=0.0):
    basic = max(0.0, bracket_tax(inc, ON_BRACKETS) - ON_BPA * ON_BRACKETS[0][1])
    surtax = ON_SURR1 * max(0.0, basic - ON_SUR1) + ON_SURR2 * max(0.0, basic - ON_SUR2)
    return max(0.0, basic + surtax - gu_e * ODTCE - gu_n * ODTCN)


def annual_tax(months, inp: Inputs, years=50):
    out = []
    for y in range(1, years + 1):
        ms = [m for m in months if m["year"] == y]
        divel = sum(m["divel"] for m in ms)
        divne = sum(m["divne"] for m in ms)
        divfor = sum(m["divfor"] for m in ms)
        gains = max(0.0, sum(m["gain"] for m in ms))
        rrsp = sum(m["draw"]["RR"] for m in ms)
        gu_e, gu_n = divel * (1 + GUE), divne * (1 + GUN)
        inc = gu_e + gu_n + divfor + CG_INC * gains + rrsp
        standalone = fed_tax(inc, gu_e, gu_n) + on_tax(inc, gu_e, gu_n)
        stacked = (fed_tax(inc + inp.other_inc, gu_e, gu_n)
                   + on_tax(inc + inp.other_inc, gu_e, gu_n))
        base = fed_tax(inp.other_inc) + on_tax(inp.other_inc)
        out.append(dict(year=y, divel=divel, divne=divne, divfor=divfor, gains=gains,
                        rrsp=rrsp, taxable=inc, standalone=standalone,
                        incremental=stacked - base))
    return out
