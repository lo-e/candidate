#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Verifier for humanlaya/coding-exam-002 (KLC-2026-A wind-down CCO).

The GOLDEN is recomputed on the fly by importing the oracle engine
(/solution/oracle_engine.py) that also produced the Oracle run, so the Oracle
run scores 1.0 by construction and no hand-maintained answer file exists.

Scoring architecture (reward = sum(weight * passed) / sum(weight)):
  * Structural (~8%): the 6 required sheets exist with exact headers, row
    counts and the DealSummary metric sequence.
  * Golden cell checks (~25%): individual cells against the engine's values
    (money +/-0.02; rounding-constitution cells +/-0.005; ratios +/-0.00005;
    flags / days / dates exact).
  * Deep golden-anchored checks (~62%): the difficulty spine.  Per-period
    trigger flags (equity-block latch state and the five coverage tests),
    sentinel note balances across the timeline, per-class cash aggregates,
    deal-level totals, reserve-account aggregates, and per-period cash-routing
    totals.  These are all-or-nothing and depend on the FULL trajectory, so a
    single misread clause (day count, grid coupon, PIK, diversion, reinvestment
    boundary, dirty-data rule) breaks most of them.
  * Self-consistency identity gates (~5%): per-period cash identities and
    roll-forwards evaluated on the candidate's own numbers.

The verifier NEVER raises: any error (missing/corrupt workbook, bad layout)
writes reward 0 to /logs/verifier/reward.txt and exits cleanly.  A per-check
detail file is written to /logs/verifier/detail.txt.
"""
import os
import sys
import traceback
from decimal import Decimal, InvalidOperation
from pathlib import Path

sys.path.insert(0, "/solution")
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import openpyxl  # noqa: E402
from oracle_engine import (  # noqa: E402
    CLASS_ORDER, SHEET_COLUMNS, compute_from_file, parse_date,
)

REWARD = Path("/logs/verifier/reward.txt")
DETAIL = Path("/logs/verifier/detail.txt")
CANDIDATE = Path("/output/result.xlsx")
INPUT = Path("/data/init.xlsx")

MONEY_TOL = 0.02
ROUND_TOL = 0.005      # rounding-constitution cells
RATIO_TOL = 0.00005    # OC/IC ratios (4dp, ROUND_HALF_UP)

# --------------------------------------------------------------------------- #
# weights (share points; normalized by the total)
# --------------------------------------------------------------------------- #
W_SHEET = 0.40
W_HEADER = 0.40
W_ROWS = 0.30
W_DS_LABELS = 0.30

W_PF = {"InterestProceeds": 0.014, "PrincipalProceeds": 0.014,
        "Recoveries": 0.006, "CollateralPar": 0.012}
W_NS = {"BeginBalance": 0.004, "Rate": 0.002, "AccDays": 0.002,
        "InterestDue": 0.005, "InterestPaid": 0.005, "InterestPIK": 0.003,
        "PrincipalPaid": 0.003, "EndBalance": 0.004}
W_IW = 0.004       # per cell
W_PW = 0.004       # per cell
W_CT_VAL = 0.005   # per ratio/par cell
W_RT = 0.004       # per cell

W_FLAG = 0.120     # per trigger flag (deep)
W_SENT = 0.200     # per sentinel balance (deep)
W_CLS_AGG = 0.320  # per class aggregate (deep)
W_DEAL = 0.450     # per DealSummary metric (deep)
W_RES_AGG = 0.600  # per reserve aggregate (deep)
W_ROUTE = 0.150    # per period routing total (deep)
W_REderive = 1.500 # RedemptionPeriod / UnblockPeriod re-derivation (deep)

W_G_IDENT = 0.040  # G1/G2 per period
W_G_ROLL = 0.002   # G3/G4 per row
W_G_RES = 0.010    # G5 per period
W_G_XSHEET = 0.020 # G6/G7/G8/G9 per row
W_G_DS = 0.020     # G10 per metric
W_G_OC = 0.010     # G11 per period
W_G_PERIOD = 0.300 # G12/G13

DS_METRICS = ["ClosingDate", "ReinvestEndDate", "NoteBalanceAtClose",
              "CollateralParAtClose", "ReserveTarget", "TotalInterestProceeds",
              "TotalPrincipalProceeds", "TotalSeniorFees", "TotalSubFees",
              "TotalInterestToNotes", "TotalPIKCapitalized",
              "TotalPrincipalToNotes", "TotalDivertedToPrincipal",
              "TotalEquity", "FinalReserveBalance", "FinalNoteBalance",
              "RedemptionPeriod", "UnblockPeriod"]

DATE_METRICS = {"ClosingDate", "ReinvestEndDate"}
PERIOD_METRICS = {"RedemptionPeriod", "UnblockPeriod"}
ROUND_SENSITIVE_IW = {"SeniorFee", "SubFee"}
ROUND_SENSITIVE_RT = {"InterestCredited"}
SENTINEL_PERIODS = (1, 5, 9, 13, 17, 20, 22, 24)


def _to_num(v):
    if v is None or isinstance(v, bool):
        return None
    if isinstance(v, Decimal):
        return v
    if isinstance(v, (int, float)):
        return Decimal(str(v))
    s = str(v).strip().replace(",", "")
    if s == "":
        return None
    try:
        return Decimal(s)
    except InvalidOperation:
        return None


def _approx(cand, exp, tol):
    c = _to_num(cand)
    if c is None:
        return False
    try:
        return abs(float(c) - float(exp)) <= tol
    except Exception:
        return False


def _norm_flag(v, on="PASS", off="FAIL"):
    if isinstance(v, bool):
        return on if v else off
    if v is None:
        return None
    s = str(v).strip().upper()
    if s in ("PASS", "TRUE", "YES", "Y", "1", "ON"):
        return on
    if s in ("FAIL", "FALSE", "NO", "N", "0", "OFF"):
        return off
    return s


def _norm_date(v):
    if v is None:
        return None
    if isinstance(v, str):
        s = v.strip()
        import datetime
        for fmt in ("%Y-%m-%d", "%d-%b-%Y", "%d-%B-%Y", "%m/%d/%Y", "%Y/%m/%d"):
            try:
                return datetime.datetime.strptime(s, fmt).strftime("%Y-%m-%d")
            except ValueError:
                continue
        return s.upper()
    try:
        return parse_date(v).strftime("%Y-%m-%d")
    except Exception:
        return str(v)


def _read_table(ws):
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return [], []
    header = [str(h).strip() if h is not None else "" for h in rows[0]]
    data = []
    for r in rows[1:]:
        if r is None:
            continue
        if all(v is None or str(v).strip() == "" for v in r):
            continue
        data.append(dict(zip(header, r)))
    return header, data


def _read_metric_sheet(ws):
    rows = []
    for row in ws.iter_rows(values_only=True):
        if not row or row[0] is None:
            continue
        key = str(row[0]).strip()
        if key.upper() in ("METRIC", "METRICS", "KEY", "ITEM"):
            continue
        if key:
            rows.append((key, row[1] if len(row) > 1 else None))
    return rows


def grade(cand_wb, expected):
    checks = []

    def add(name, ok, w):
        checks.append((name, bool(ok), float(w)))

    wf_rows = expected["iw_rows"]
    pf_rows = expected["pw_rows"]
    ns_rows = expected["notes_rows"]
    ct_rows = expected["ct_rows"]
    rt_rows = expected["rt_rows"]
    fl_rows = expected["flows_rows"]
    summary = dict(expected["summary"])
    K = expected["inp"]["K"]

    need_sheets = ["PoolFlows", "NoteSchedule", "InterestWaterfall",
                   "PrincipalWaterfall", "CoverageTests", "ReserveTrace",
                   "DealSummary"]
    sheet_ok = {s: (s in cand_wb.sheetnames) for s in need_sheets}
    for s in need_sheets:
        add(f"struct: sheet '{s}' present", sheet_ok[s], W_SHEET)

    def read_or_empty(sheet):
        if sheet not in cand_wb.sheetnames:
            return [], []
        return _read_table(cand_wb[sheet])

    for s, cols in SHEET_COLUMNS.items():
        hdr, _ = read_or_empty(s)
        add(f"struct: header of '{s}'", hdr[:len(cols)] == cols, W_HEADER)
    ds_rows = _read_metric_sheet(cand_wb["DealSummary"]) if sheet_ok["DealSummary"] else []
    ds_labels = [k for k, _ in ds_rows]
    add("struct: DealSummary metric sequence",
        ds_labels[:len(DS_METRICS)] == DS_METRICS, W_DS_LABELS)

    expected_counts = {"PoolFlows": len(fl_rows), "NoteSchedule": len(ns_rows),
                       "InterestWaterfall": len(wf_rows),
                       "PrincipalWaterfall": len(pf_rows),
                       "CoverageTests": len(ct_rows), "ReserveTrace": len(rt_rows),
                       "DealSummary": len(DS_METRICS)}
    for s, n in expected_counts.items():
        _, data = read_or_empty(s)
        add(f"struct: row count of '{s}' == {n}", len(data) == n, W_ROWS)

    # ------------------------------------------------------------------ #
    # candidate maps
    # ------------------------------------------------------------------ #
    _, pf_data = read_or_empty("PoolFlows")
    pf_map = {}
    for r in pf_data:
        try:
            pf_map[int(r.get("Period"))] = r
        except (TypeError, ValueError):
            pass
    _, ns_data = read_or_empty("NoteSchedule")
    ns_map = {}
    for r in ns_data:
        try:
            ns_map[(int(r.get("Period")), str(r.get("Class", "")).strip())] = r
        except (TypeError, ValueError):
            pass
    _, iw_data = read_or_empty("InterestWaterfall")
    iw_map = {}
    for r in iw_data:
        try:
            iw_map[int(r.get("Period"))] = r
        except (TypeError, ValueError):
            pass
    _, pw_data = read_or_empty("PrincipalWaterfall")
    pw_map = {}
    for r in pw_data:
        try:
            pw_map[int(r.get("Period"))] = r
        except (TypeError, ValueError):
            pass
    _, ct_data = read_or_empty("CoverageTests")
    ct_map = {}
    for r in ct_data:
        try:
            ct_map[int(r.get("Period"))] = r
        except (TypeError, ValueError):
            pass
    _, rt_data = read_or_empty("ReserveTrace")
    rt_map = {}
    for r in rt_data:
        try:
            rt_map[int(r.get("Period"))] = r
        except (TypeError, ValueError):
            pass
    ds_map = {}
    for k, v in ds_rows:
        ds_map.setdefault(k, v)

    def num(row, col):
        v = _to_num(row.get(col)) if row else None
        return float(v) if v is not None else None

    # ------------------------------------------------------------------ #
    # golden cell checks (~25%)
    # ------------------------------------------------------------------ #
    for fr in fl_rows:
        row = pf_map.get(fr["P"], {})
        for col, w in W_PF.items():
            exp = fr[{"InterestProceeds": "ip", "PrincipalProceeds": "pp",
                      "Recoveries": "rec", "CollateralPar": "par"}[col]]
            add(f"PoolFlows[P{fr['P']}].{col}", _approx(row.get(col), exp, MONEY_TOL), w)

    for nr in ns_rows:
        row = ns_map.get((nr["P"], nr["class"]), {})
        for col, w in W_NS.items():
            exp = nr[{"BeginBalance": "begin", "Rate": "rate", "AccDays": "days",
                      "InterestDue": "due", "InterestPaid": "paid",
                      "InterestPIK": "pik", "PrincipalPaid": "prin",
                      "EndBalance": "end"}[col]]
            if col == "Rate":
                ok = _approx(row.get(col), exp, 1e-6)
            elif col == "AccDays":
                ok = _approx(row.get(col), exp, 0.5)
            else:
                ok = _approx(row.get(col), exp, MONEY_TOL)
            add(f"NoteSchedule[P{nr['P']}.{nr['class']}].{col}", ok, w)

    iw_cols = [c for c in SHEET_COLUMNS["InterestWaterfall"] if c != "Period"]
    for ir in wf_rows:
        row = iw_map.get(ir["P"], {})
        exp_map = {
            "InterestProceeds": ir["ip"], "ReserveDraw": ir["draw"],
            "CarryPaid_A1": ir["cpA1"], "CarryPaid_A2": ir["cpA2"],
            "CarryPaid_B": ir["cpB"], "ExpenseCarryPaid": ir["exp_paid"],
            "SeniorFee": ir["sf"], "Trustee": ir["tr"],
            "Int_A1": ir["iA1"], "Int_A2": ir["iA2"], "Int_B": ir["iB"],
            "Int_C": ir["iC"], "Int_D": ir["iD"], "Int_E": ir["iE"],
            "DivertedToPrincipal": ir["div"], "ReserveReplenish": ir["repl"],
            "SubFeeCarryPaid": ir["sfcp"], "SubFee": ir["sub"], "Equity": ir["eq"],
        }
        for col in iw_cols:
            tol = ROUND_TOL if col in ROUND_SENSITIVE_IW else MONEY_TOL
            add(f"InterestWaterfall[P{ir['P']}].{col}",
                _approx(row.get(col), exp_map[col], tol), W_IW)

    pw_cols = [c for c in SHEET_COLUMNS["PrincipalWaterfall"] if c != "Period"]
    for pr in pf_rows:
        row = pw_map.get(pr["P"], {})
        exp_map = {
            "PrincipalProceeds": pr["pp"], "CarryIn": pr["carry_in"],
            "TotalAvailable": pr["avail"], "Prin_A1": pr["pA1"],
            "Prin_A2": pr["pA2"], "Prin_B": pr["pB"], "Prin_C": pr["pC"],
            "Prin_D": pr["pD"], "Prin_E": pr["pE"], "CarryOut": pr["carry_out"],
            "PrinToEquity": pr["eq"],
        }
        for col in pw_cols:
            add(f"PrincipalWaterfall[P{pr['P']}].{col}",
                _approx(row.get(col), exp_map[col], MONEY_TOL), W_PW)

    for cr in ct_rows:
        row = ct_map.get(cr["P"], {})
        for col in ("CollateralPar", "OC_A", "OC_C", "OC_D", "OC_E", "IC_A", "IC_C"):
            exp = cr[{"CollateralPar": "par", "OC_A": "ocA", "OC_C": "ocC",
                      "OC_D": "ocD", "OC_E": "ocE", "IC_A": "icA", "IC_C": "icC"}[col]]
            tol = MONEY_TOL if col == "CollateralPar" else RATIO_TOL
            add(f"CoverageTests[P{cr['P']}].{col}", _approx(row.get(col), exp, tol),
                W_CT_VAL)

    for rr in rt_rows:
        row = rt_map.get(rr["P"], {})
        exp_map = {"BeginBalance": rr["begin"], "InterestCredited": rr["int"],
                   "Draw": rr["draw"], "Reimburse": rr["reimb"],
                   "TopUp": rr["topup"], "Release": rr["release"],
                   "EndBalance": rr["end"], "UnreimbursedEnd": rr["unreimb"]}
        for col, exp in exp_map.items():
            tol = ROUND_TOL if col in ROUND_SENSITIVE_RT else MONEY_TOL
            add(f"ReserveTrace[P{rr['P']}].{col}", _approx(row.get(col), exp, tol), W_RT)

    # ------------------------------------------------------------------ #
    # DEEP golden-anchored checks (~70%): all-or-nothing trajectory vectors
    # ------------------------------------------------------------------ #
    # (1) note balance / cash vectors per class: ONE wrong period fails the
    #     whole vector, so a single misread clause anywhere breaks them all.
    vec_fields = [("EndBalance", "end", 2.6), ("InterestPIK", "pik", 1.2),
                  ("PrincipalPaid", "prin", 1.2), ("InterestDue", "due", 1.2),
                  ("InterestPaid", "paid", 1.2)]
    for c in CLASS_ORDER:
        rows = [r for r in ns_rows if r["class"] == c]
        cand_rows = [ns_map.get((r["P"], c), {}) for r in rows]
        for col, key, w in vec_fields:
            got = [num(r, col) for r in cand_rows]
            exp = [float(r[key]) for r in rows]
            ok = all(v is not None for v in got) and \
                all(abs(g - e) <= 0.005 for g, e in zip(got, exp))
            add(f"deep: {c} {col} trajectory (24 periods)", ok, w)
        red = 0
        for r, crow in zip(rows, cand_rows):
            v = num(crow, "EndBalance")
            if v is not None and abs(v) <= 0.005:
                red = r["P"]
                break
        exp_red = next((r["P"] for r in rows if r["end"] == 0), 0)
        add(f"deep: {c} fully-repaid period == {exp_red}", red == exp_red, 1.5)

    # (1b) coverage-ratio trajectories (par / balances / dues encoding)
    for label, key in [("OC_A", "ocA"), ("OC_C", "ocC"), ("OC_D", "ocD"),
                       ("OC_E", "ocE"), ("IC_A", "icA"), ("IC_C", "icC")]:
        got = [num(ct_map.get(cr["P"], {}), label) for cr in ct_rows]
        exp = [float(cr[key]) for cr in ct_rows]
        ok = all(v is not None for v in got) and \
            all(abs(g - e) <= 0.00005 for g, e in zip(got, exp))
        add(f"deep: {label} trajectory (24 periods)", ok, 1.4)

    # (1c) principal waterfall per-class payment trajectories
    for c in CLASS_ORDER:
        got = [num(pw_map.get(pr["P"], {}), f"Prin_{c}") for pr in pf_rows]
        exp = [float(r[f"p{c}"]) for r in pf_rows]
        ok = all(v is not None for v in got) and \
            all(abs(g - e) <= 0.005 for g, e in zip(got, exp))
        add(f"deep: PW Prin_{c} trajectory (24 periods)", ok, 1.4)

    # (1d) reserve per-column trajectories
    for label, key in [("Draw", "draw"), ("TopUp", "topup"), ("Reimburse", "reimb"),
                       ("Release", "release"), ("UnreimbursedEnd", "unreimb")]:
        got = [num(rt_map.get(rr["P"], {}), label) for rr in rt_rows]
        exp = [float(rr[key]) for rr in rt_rows]
        ok = all(v is not None for v in got) and \
            all(abs(g - e) <= 0.005 for g, e in zip(got, exp))
        add(f"deep: RT {label} trajectory (24 periods)", ok, 1.2)

    # (1e) interest waterfall column trajectories
    for label, key in [("ReserveDraw", "draw"), ("DivertedToPrincipal", "div"),
                       ("ReserveReplenish", "repl"), ("SubFee", "sub"),
                       ("SeniorFee", "sf"), ("Equity", "eq")]:
        tol = 0.005 if label in ("SubFee", "SeniorFee") else 0.02
        got = [num(iw_map.get(ir["P"], {}), label) for ir in wf_rows]
        exp = [float(r[key]) for r in wf_rows]
        ok = all(v is not None for v in got) and \
            all(abs(g - e) <= tol for g, e in zip(got, exp))
        add(f"deep: IW {label} trajectory (24 periods)", ok, 1.3)

    # (2) trigger flag sets: the binary readouts of the whole cascade
    def flag_set(key, want):
        return tuple(P for P, cr in ((r["P"], r) for r in ct_rows)
                     if (cr[key] if want == "bad" else not cr[key]))
    block_on = tuple(cr["P"] for cr in ct_rows if cr["block"])
    anyfail_on = tuple(cr["P"] for cr in ct_rows if cr["anyfail"])
    got_block = tuple(P for P in range(1, K + 1)
                      if _norm_flag(ct_map.get(P, {}).get("EquityBlock"), "ON", "OFF") == "ON")
    got_any = tuple(P for P in range(1, K + 1)
                    if _norm_flag(ct_map.get(P, {}).get("AnyFail"), "PASS", "FAIL") == "FAIL")
    add("deep: EquityBlock ON-period set matches golden",
        got_block == block_on, 4.0)
    add("deep: AnyFail period set matches golden", got_any == anyfail_on, 4.0)
    for col, key in [("OC_A_Flag", "fA"), ("OC_C_Flag", "fC"), ("OC_D_Flag", "fD"),
                     ("IC_A_Flag", "fiA"), ("IC_C_Flag", "fiC")]:
        exp_set = tuple(cr["P"] for cr in ct_rows if not cr[key])
        got_set = tuple(P for P in range(1, K + 1)
                        if _norm_flag(ct_map.get(P, {}).get(col), "PASS", "FAIL") == "FAIL")
        add(f"deep: {col} failing-period set matches golden", got_set == exp_set, 1.6)

    # (3) DealSummary metrics (deep, golden-anchored)
    for metric in DS_METRICS:
        got = ds_map.get(metric)
        exp = summary[metric]
        if metric in DATE_METRICS:
            ok = _norm_date(got) == exp
        elif metric in PERIOD_METRICS:
            ok = _approx(got, exp, 1e-9)
        else:
            ok = _approx(got, exp, MONEY_TOL)
        add(f"deep: DealSummary[{metric}]", ok, 0.55)

    # (4) reserve trajectory + aggregates
    res_vec_ok = all(
        _approx(num(rt_map.get(rr["P"], {}), "EndBalance"), rr["end"], 0.005)
        for rr in rt_rows)
    add("deep: reserve EndBalance trajectory (24 periods)", res_vec_ok, 3.0)
    res_cols = [("Draw", "draw", 0.8), ("TopUp", "topup", 0.8),
                ("Reimburse", "reimb", 0.8), ("Release", "release", 0.8)]
    for col, key, w in res_cols:
        got = sum(num(rt_map.get(rr["P"], {}), col) or 0 for rr in rt_rows)
        add(f"deep: reserve total {col}", _approx(got, sum(r[key] for r in rt_rows), 0.05), w)
    last = rt_map.get(rt_rows[-1]["P"], {})
    add("deep: reserve UnreimbursedEnd at P24",
        _approx(num(last, "UnreimbursedEnd"), rt_rows[-1]["unreimb"], 0.02), 0.8)

    # (5) cash-routing vectors
    for label, getter, w in [
        ("total principal routed to notes",
         lambda P: sum(num(ns_map.get((P, c), {}), "PrincipalPaid") or 0 for c in CLASS_ORDER), 2.2),
        ("total interest cash to notes",
         lambda P: sum((num(iw_map.get(P, {}), f"Int_{c}") or 0)
                       + (num(iw_map.get(P, {}), f"CarryPaid_{c}") or 0)
                       for c in ("A1", "A2", "B"))
                      + sum(num(iw_map.get(P, {}), f"Int_{c}") or 0 for c in ("C", "D", "E")), 2.2),
        ("DivertedToPrincipal", lambda P: num(iw_map.get(P, {}), "DivertedToPrincipal") or 0, 1.6),
        ("Equity", lambda P: num(iw_map.get(P, {}), "Equity") or 0, 1.6),
        ("CarryOut", lambda P: num(pw_map.get(P, {}), "CarryOut") or 0, 1.6),
     ]:
        got = [getter(P) for P in range(1, K + 1)]
        if label == "total principal routed to notes":
            exp = [float(sum(r["prin"] for r in ns_rows if r["P"] == P)) for P in range(1, K + 1)]
        elif label == "total interest cash to notes":
            exp = [float(r["iA1"] + r["iA2"] + r["iB"] + r["iC"] + r["iD"] + r["iE"]
                         + r["cpA1"] + r["cpA2"] + r["cpB"]) for r in wf_rows]
        elif label == "DivertedToPrincipal":
            exp = [float(r["div"]) for r in wf_rows]
        elif label == "Equity":
            exp = [float(r["eq"]) for r in wf_rows]
        else:
            exp = [float(r["carry_out"]) for r in pf_rows]
        ok = all(abs(g - e) <= 0.05 for g, e in zip(got, exp))
        add(f"deep: {label} trajectory (24 periods)", ok, w)

    # (6) RedemptionPeriod / UnblockPeriod re-derived from the candidate's cells
    redp = 0
    for P in range(1, K + 1):
        ends = [num(ns_map.get((P, c), {}), "EndBalance") for c in CLASS_ORDER]
        if any(v is None for v in ends):
            redp = -1
            break
        if all(abs(v) <= 0.005 for v in ends):
            redp = P
            break
    got_red = _to_num(ds_map.get("RedemptionPeriod"))
    add("deep: RedemptionPeriod == first fully-repaid period (re-derived)",
        got_red is not None and redp >= 0 and int(float(got_red)) == redp, 2.0)
    got_unb = _to_num(ds_map.get("UnblockPeriod"))
    exp_unb = summary["UnblockPeriod"]
    add("deep: UnblockPeriod matches golden", _approx(got_unb, exp_unb, 1e-9), 2.0)

    # ------------------------------------------------------------------ #
    # self-consistency identity gates (~5%)
    # ------------------------------------------------------------------ #
    for ir in wf_rows:
        row = iw_map.get(ir["P"], {})
        vals = [num(row, c) for c in SHEET_COLUMNS["InterestWaterfall"][1:]]
        ok = None not in vals and abs((vals[0] + vals[1]) - sum(vals[2:])) <= 0.02
        add(f"gate G1: interest waterfall cash identity P{ir['P']}", ok, W_G_IDENT)
    for pr in pf_rows:
        row = pw_map.get(pr["P"], {})
        vals = [num(row, c) for c in SHEET_COLUMNS["PrincipalWaterfall"][1:]]
        ok = None not in vals and abs((vals[0] + vals[1]) - sum(vals[3:])) <= 0.02 \
            and abs(vals[2] - (vals[0] + vals[1])) <= 0.02
        add(f"gate G2: principal waterfall cash identity P{pr['P']}", ok, W_G_IDENT)

    for nr in ns_rows:
        row = ns_map.get((nr["P"], nr["class"]), {})
        b, pik, prin, e = (num(row, "BeginBalance"), num(row, "InterestPIK"),
                           num(row, "PrincipalPaid"), num(row, "EndBalance"))
        ok = None not in (b, pik, prin, e) and abs(b + pik - prin - e) <= 0.02
        add(f"gate G3: note balance roll P{nr['P']}.{nr['class']}", ok, W_G_ROLL)
        if nr["P"] > 1:
            prev = ns_map.get((nr["P"] - 1, nr["class"]), {})
            pe = num(prev, "EndBalance")
            ok = pe is not None and e is not None and abs(pe - b) <= 0.02
            add(f"gate G4: note balance chain P{nr['P']}.{nr['class']}", ok, W_G_ROLL)

    for i, rr in enumerate(rt_rows):
        row = rt_map.get(rr["P"], {})
        vals = [num(row, c) for c in
                ["BeginBalance", "InterestCredited", "Draw", "Reimburse", "TopUp",
                 "Release", "EndBalance"]]
        ok = None not in vals and \
            abs(vals[0] + vals[1] - vals[2] + vals[3] + vals[4] - vals[5] - vals[6]) <= 0.02
        add(f"gate G5: reserve roll P{rr['P']}", ok, W_G_RES)
        if i:
            prev = rt_map.get(rt_rows[i - 1]["P"], {})
            pe = num(prev, "EndBalance")
            ok = pe is not None and vals[0] is not None and abs(pe - vals[0]) <= 0.02
            add(f"gate G5b: reserve chain P{rr['P']}", ok, W_G_RES)

    # NoteSchedule.InterestPaid carries the CURRENT-period interest cash only;
    # the prior-shortfall repayments live in the CarryPaid_* columns.
    int_col = {"A1": "Int_A1", "A2": "Int_A2", "B": "Int_B", "C": "Int_C",
               "D": "Int_D", "E": "Int_E"}
    for nr in ns_rows:
        irow = iw_map.get(nr["P"], {})
        exp_val = num(irow, int_col[nr["class"]])
        got = num(ns_map.get((nr["P"], nr["class"]), {}), "InterestPaid")
        ok = exp_val is not None and got is not None and abs(got - exp_val) <= 0.02
        add(f"gate G6: NoteSchedule.InterestPaid == waterfall P{nr['P']}.{nr['class']}",
            ok, W_G_XSHEET)

    for ir in wf_rows:
        P = ir["P"]
        prow = pw_map.get(P, {})
        s = 0.0
        ok = True
        for c in CLASS_ORDER:
            ns_v = num(ns_map.get((P, c), {}), "PrincipalPaid")
            pw_v = num(prow, f"Prin_{c}")
            if ns_v is None or pw_v is None:
                ok = False
                break
            s += ns_v - pw_v
        div = num(iw_map.get(P, {}), "DivertedToPrincipal")
        ok = ok and div is not None and abs(s - div) <= 0.02
        add(f"gate G7: diversion split identity P{P}", ok, W_G_XSHEET)

    for rr in rt_rows:
        rrow = rt_map.get(rr["P"], {})
        irow = iw_map.get(rr["P"], {})
        d1, d2 = num(rrow, "Draw"), num(irow, "ReserveDraw")
        r1, t1 = num(rrow, "Reimburse"), num(rrow, "TopUp")
        r2 = num(irow, "ReserveReplenish")
        ok1 = None not in (d1, d2) and abs(d1 - d2) <= 0.02
        ok2 = None not in (r1, t1, r2) and abs(r1 + t1 - r2) <= 0.02
        add(f"gate G8a: ReserveTrace.Draw == IW.ReserveDraw P{rr['P']}", ok1, W_G_XSHEET)
        add(f"gate G8b: ReserveTrace.Reimburse+TopUp == IW.ReserveReplenish P{rr['P']}",
            ok2, W_G_XSHEET)

    for cr in ct_rows:
        v1 = num(ct_map.get(cr["P"], {}), "CollateralPar")
        v2 = num(pf_map.get(cr["P"], {}), "CollateralPar")
        ok = v1 is not None and v2 is not None and abs(v1 - v2) <= 0.02
        add(f"gate G9: CT.CollateralPar == PF.CollateralPar P{cr['P']}", ok, W_G_XSHEET)

    tot_ip = sum(num(iw_map.get(ir["P"], {}), "InterestProceeds") or 0 for ir in wf_rows)
    tot_pp = sum(num(pw_map.get(pr["P"], {}), "PrincipalProceeds") or 0 for pr in pf_rows)
    for metric, val in [("TotalInterestProceeds", tot_ip),
                        ("TotalPrincipalProceeds", tot_pp)]:
        got = _to_num(ds_map.get(metric))
        ok = got is not None and abs(float(got) - val) <= 0.05
        add(f"gate G10: DealSummary[{metric}] == recomputed", ok, W_G_DS)

    for cr in ct_rows:
        par = num(ct_map.get(cr["P"], {}), "CollateralPar")
        den = 0.0
        ok = par is not None
        for c in ("A1", "A2", "B"):
            v = num(ns_map.get((cr["P"], c), {}), "BeginBalance")
            if v is None:
                ok = False
                break
            den += v
        ratio = num(ct_map.get(cr["P"], {}), "OC_A")
        if ok and den > 0 and ratio is not None:
            ok = abs(ratio - par / den) <= 0.001
        elif ok and den == 0:
            ok = ratio is not None and abs(ratio) <= 0.001
        add(f"gate G11: OC_A == par/den(A1+A2+B) P{cr['P']}", ok, W_G_OC)

    block = False
    streak = 0
    latch_periods = int(expected["inp"]["latch_periods"])
    chain_ok = True
    for cr in ct_rows:
        if not cr["anyfail"]:
            streak += 1
            if block and streak >= latch_periods:
                block = False
        else:
            streak = 0
            block = True
        gotb = _norm_flag(ct_map.get(cr["P"], {}).get("EquityBlock"), "ON", "OFF")
        expb = "ON" if block else "OFF"
        if gotb != expb:
            chain_ok = False
    add("gate G13: EquityBlock latch chain consistent with AnyFail flags",
        chain_ok, W_G_PERIOD)

    return checks


def main():
    REWARD.parent.mkdir(parents=True, exist_ok=True)
    detail_lines = []

    def log(s):
        detail_lines.append(s)
        print(s)

    reward = 0.0
    try:
        if not CANDIDATE.exists():
            log("FAIL: /output/result.xlsx not found (no submission)")
        else:
            try:
                cand_wb = openpyxl.load_workbook(str(CANDIDATE), data_only=True)
            except Exception as e:
                cand_wb = None
                log(f"FAIL: could not read /output/result.xlsx: {e}")
            if cand_wb is not None:
                expected = compute_from_file(str(INPUT), None)
                checks = grade(cand_wb, expected)
                total_w = sum(w for _, _, w in checks)
                passed_w = sum(w for _, ok, w in checks if ok)
                reward = passed_w / total_w if total_w else 0.0
                n_pass = sum(1 for _, ok, _ in checks if ok)
                for name, ok, w in checks:
                    if not ok:
                        log(f"FAIL [{w:.4f}]: {name}")
                log(f"passed {n_pass}/{len(checks)} checks "
                    f"(weighted {passed_w:.3f}/{total_w:.3f})")
    except Exception:
        reward = 0.0
        log("VERIFIER EXCEPTION (scored 0):")
        log(traceback.format_exc())

    reward = max(0.0, min(1.0, reward))
    REWARD.write_text("1.0" if reward >= 0.999999 else f"{reward:.4f}",
                      encoding="utf-8")
    DETAIL.write_text("\n".join(detail_lines) + "\n", encoding="utf-8")
    log(f"reward = {reward:.6f}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
