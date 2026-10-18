#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Oracle engine — coding-exam-002 "KLC-2026-A 清仓型合成信贷凭证 (wind-down CCO)".

This module is the SINGLE SOURCE OF TRUTH for the correct answer. It is used by:
  * solution/oracle_convert.py  -> writes /output/result.xlsx (the Oracle run)
  * tests/test_outputs.py       -> recomputes the golden and grades a candidate

Because the grader recomputes expected values from the *same* engine that
produces the Oracle output, the Oracle test is 1.0 by construction, and the
golden is never a hand-maintained static file.

DESIGN SUMMARY (all rules are explicit in instruction.md; nothing hidden):
  * Monthly determination dates; two day counts (Act/360 for floating classes,
    30/360 for fixed classes); floating classes use a per-period base-rate curve.
  * Dirty-data ingestion: negatives->abs, blank->0, comma-number strings,
    duplicate (Period,AssetID) rows, unknown AssetIDs, sparse rows, and
    EndPar carry-forward (blank vs 0 semantics).
  * Dual waterfall: interest waterfall (14 ordered steps incl. carryforward
    repayments, reserve draws, coverage-test interest diversion) and principal
    waterfall (reinvestment retention, pro-rata group A1/A2, sequential B..E).
  * Class A2's fixed coupon is derived from the period-1 base rate via a
    "floor to the 0.125% grid" clause (adversarial parameter).
  * Coverage tests (OC_A/OC_C/OC_D, IC_A/IC_C) computed on beginning-of-period
    note balances + current-period collateral par snapshot; an equity-block
    LATCH with a 2-consecutive-pass release (hysteresis) gates the tail of the
    interest waterfall.
  * Reserve account: draws cover A1/A2/B interest shortfalls; draws create an
    unreimbursed balance; reimburse-then-top-up step; interest credited
    (ROUND_DOWN); balance released to equity when all notes are repaid.
  * Deferrable classes C/D/E PIK their unpaid interest (capitalization).

All money math uses Decimal. Rounding constitution (instruction.md section 7):
  senior fee / trustee->expense carry : ROUND_DOWN to cent (senior fee)
  sub management fee                  : ROUND_UP to cent
  note interest                       : ROUND_HALF_UP to cent
  reserve interest                    : ROUND_DOWN to cent
  OC/IC ratios                        : ROUND_HALF_UP to 4 decimals (both for
                                        comparison against thresholds and output)
"""

import datetime
import calendar
from decimal import Decimal, ROUND_DOWN, ROUND_UP, ROUND_HALF_UP, getcontext

getcontext().prec = 60

CENT = Decimal("0.01")
Q4 = Decimal("0.0001")
ZERO = Decimal(0)
ONE = Decimal(1)

CLASS_ORDER = ["A1", "A2", "B", "C", "D", "E"]
GROUP1 = ("A1", "A2")            # pro-rata amortisation group
SEQ_AFTER = ("B", "C", "D", "E")  # sequential after the pro-rata group
DEFERRABLE = {"C", "D", "E"}      # unpaid interest PIKs (capitalises)
DRAW_CLASSES = ("A1", "A2", "B")  # interest shortfall supported by reserve draws
FLOAT_TYPE = "FLOAT"
FIXED_TYPE = "FIXED"

MONTHS = {
    "jan": 1, "feb": 2, "mar": 3, "apr": 4, "may": 5, "jun": 6,
    "jul": 7, "aug": 8, "sep": 9, "oct": 10, "nov": 11, "dec": 12,
}

# Optional mutation switches used ONLY by solution/mutations.py for the
# discrimination experiments. The oracle (and the shipped golden) always runs
# with an empty dict.
MUT = {}


# --------------------------------------------------------------------------- #
# Decimal helpers
# --------------------------------------------------------------------------- #
def q2(x, mode=ROUND_HALF_UP):
    return Decimal(x).quantize(CENT, rounding=mode)


def floor2(x):
    return q2(x, ROUND_DOWN)


def ceil2(x):
    return q2(x, ROUND_UP)


def q4(x):
    return Decimal(x).quantize(Q4, rounding=ROUND_HALF_UP)


def parse_money(v, absolute=True):
    """Parse a money cell. Blank -> 0; '1,234.56' strings supported;
    negatives are data-noise (absolute value) for the three collection columns."""
    if v is None:
        return ZERO
    if isinstance(v, Decimal):
        d = v
    elif isinstance(v, bool):
        return ZERO
    elif isinstance(v, int):
        d = Decimal(v)
    elif isinstance(v, float):
        d = Decimal(repr(v))
    else:
        s = str(v).strip().replace(",", "")
        if s == "":
            return ZERO
        d = Decimal(s)
    if d == 0:
        return ZERO
    return abs(d) if absolute else d


def parse_date(value):
    """Parse a date that may be DD-Mon-YYYY, MM/DD/YYYY, YYYY-MM-DD, a
    datetime/date object, or an Excel serial number."""
    if value is None:
        raise ValueError("empty date")
    if isinstance(value, datetime.datetime):
        return value.date()
    if isinstance(value, datetime.date):
        return value
    s = str(value).strip()
    if s == "":
        raise ValueError("empty date")
    if "/" in s:
        parts = s.split("/")
        if len(parts) == 3:
            m, d, y = parts
            return datetime.date(int(y), int(m), int(d))
    if "-" in s and any(c.isalpha() for c in s):
        d, mon, y = s.split("-")
        mk = mon.strip()[:3].lower()
        if mk in MONTHS:
            return datetime.date(int(y), MONTHS[mk], int(d))
    if "-" in s:
        y, m, d = s.split("-")
        return datetime.date(int(y), int(m), int(d))
    return datetime.date(1899, 12, 30) + datetime.timedelta(days=int(float(s)))


def days_act(d1, d2):
    """Actual days between two dates (from-and-including d1, to-but-excluding d2)."""
    return (d2 - d1).days


def days360(d1, d2):
    """30/360 (US) day count: 360*(Y2-Y1) + 30*(M2-M1) + (min(D2,30)-min(D1,30))."""
    d1d = min(d1.day, 30)
    d2d = min(d2.day, 30)
    return (d2.year - d1.year) * 360 + (d2.month - d1.month) * 30 + (d2d - d1d)


def floor_grid(x, step):
    """Largest multiple of `step` not greater than x."""
    x = Decimal(x)
    step = Decimal(step)
    return (x / step).to_integral_value(rounding=ROUND_DOWN) * step


def allocate_pro_rata(amount, caps_ordered):
    """Pro-rata allocation used by the diversion step and the principal
    waterfall's group-1 step.

    caps_ordered: list of (class, capacity) in priority order, capacity >= 0.
    Step 1: share_c = floor-to-cent(amount * cap_c / total_cap), capped at cap_c
            (each share is rounded DOWN so the shares never over-allocate).
    Step 2 (sweep): any leftover cash is paid to classes in priority order up
    to their remaining capacity.
    Returns (pays dict, leftover).
    """
    pays = {c: ZERO for c, _ in caps_ordered}
    amount = Decimal(amount)
    tot = sum((Decimal(cap) for _, cap in caps_ordered), ZERO)
    if amount > 0 and tot > 0:
        for c, cap in caps_ordered:
            cap = Decimal(cap)
            if cap <= 0:
                continue
            share = q2(amount * cap / tot, ROUND_DOWN)
            pays[c] = min(share, cap)
    leftover = amount - sum(pays.values())
    for c, cap in caps_ordered:
        if leftover <= 0:
            break
        cap = Decimal(cap)
        room = cap - pays[c]
        if room > 0:
            g = min(room, leftover)
            pays[c] += g
            leftover -= g
    return pays, leftover


# --------------------------------------------------------------------------- #
# Input loading
# --------------------------------------------------------------------------- #
def _sheet_rows(ws):
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return [], []
    header = [str(h).strip() if h is not None else "" for h in rows[0]]
    data = [dict(zip(header, r)) for r in rows[1:]
            if r is not None and any(v is not None and str(v).strip() != "" for v in r)]
    return header, data


def load_params(wb):
    params = {}
    ws = wb["Params"]
    for row in ws.iter_rows(values_only=True):
        if not row or row[0] is None:
            continue
        key = str(row[0]).strip()
        val = row[1]
        if key and val is not None and str(val).strip() != "":
            params[key] = val
    return params


def load_inputs(path):
    import openpyxl
    wb = openpyxl.load_workbook(path, data_only=True)

    params = load_params(wb)
    closing = parse_date(params["ClosingDate"])
    setdates = [parse_date(x.strip())
                for x in str(params["SettlementDates"]).split(";") if x.strip()]
    K = len(setdates)

    inp = {
        "closing": closing,
        "setdates": setdates,
        "K": K,
        "reinvest_months": int(params.get("ReinvestMonths", 12)),
        "senior_fee_pct": parse_money(params["SeniorFeePct"], absolute=False),
        "sub_fee_pct": parse_money(params["SubFeePct"], absolute=False),
        "trustee": parse_money(params["TrusteeExpense"]),
        "reserve_init": parse_money(params["ReserveInitial"]),
        "reserve_target_pct": parse_money(params["ReserveTargetPct"], absolute=False),
        "reserve_rate": parse_money(params["ReserveRate"], absolute=False),
        "oc_thr": {
            "A": parse_money(params["OCThreshold_A"], absolute=False),
            "C": parse_money(params["OCThreshold_C"], absolute=False),
            "D": parse_money(params["OCThreshold_D"], absolute=False),
            "E": parse_money(params["OCThreshold_E"], absolute=False),
        },
        "ic_thr": {
            "A": parse_money(params["ICThreshold_A"], absolute=False),
            "C": parse_money(params["ICThreshold_C"], absolute=False),
        },
        "latch_periods": int(params.get("LatchReleasePeriods", 2)),
        "grid_step": parse_money(params["GridStep"], absolute=False),
        "a2_spread": parse_money(params["A2Spread"], absolute=False),
        "params": params,
    }

    # --- Classes sheet ---
    _, cls_rows = _sheet_rows(wb["Classes"])
    classes = {}
    for r in cls_rows:
        cid = str(r.get("ClassID", "")).strip()
        if not cid:
            continue
        rate_type = str(r.get("RateType", "")).strip().upper()
        dc = str(r.get("DayCount", "")).strip().upper()
        deferr = str(r.get("Deferrable", "")).strip().upper()
        prg = r.get("ProRataGroup")
        classes[cid] = {
            "priority": int(r.get("Priority")),
            "orig_bal": parse_money(r.get("OrigBalance")),
            "rate_type": rate_type,
            "fixed": parse_money(r.get("FixedRate"), absolute=False) if rate_type == FIXED_TYPE else None,
            "spread": parse_money(r.get("Spread"), absolute=False) if rate_type == FLOAT_TYPE else ZERO,
            "day_count": dc,
            "deferrable": deferr in ("Y", "YES", "TRUE", "1"),
            "group1": prg is not None and str(prg).strip() not in ("", "0") and parse_money(prg) == 1,
        }
    inp["classes"] = classes

    # --- RateCurve sheet ---
    _, curve_rows = _sheet_rows(wb["RateCurve"])
    curve = {}
    for r in curve_rows:
        p = r.get("Period")
        if p is None:
            continue
        curve[int(p)] = parse_money(r.get("BaseRate"), absolute=False)
    inp["curve"] = curve

    # --- AssetTape sheet ---
    _, asset_rows = _sheet_rows(wb["AssetTape"])
    assets = {}
    asset_order = []
    for r in asset_rows:
        aid = str(r.get("AssetID", "")).strip()
        if not aid:
            continue
        assets[aid] = {
            "par": parse_money(r.get("ParAtClose")),
            "coupon": parse_money(r.get("Coupon"), absolute=False),
        }
        asset_order.append(aid)
    inp["assets"] = assets
    inp["asset_order"] = asset_order

    # --- PeriodFlows sheet ---
    _, flow_rows = _sheet_rows(wb["PeriodFlows"])
    flows = {}
    for r in flow_rows:
        p = r.get("Period")
        aid = r.get("AssetID")
        if p is None or aid is None:
            continue
        try:
            p = int(p)
        except (TypeError, ValueError):
            continue
        aid = str(aid).strip()
        row = {
            "interest_raw": r.get("InterestCollected"),
            "principal_raw": r.get("PrincipalCollected"),
            "recovery_raw": r.get("RecoveryCollected"),
            "endpar_raw": r.get("EndPar"),
        }
        flows.setdefault((p, aid), []).append(row)
    inp["flows"] = flows

    return inp


# --------------------------------------------------------------------------- #
# Collateral cash-flow aggregation (dirty-data rules)
# --------------------------------------------------------------------------- #
def aggregate_flows(inp):
    """Returns (period_flows, par_snap):
    period_flows[P] = {ip, pp, rec, }
    par_snap[P]     = aggregate collateral par snapshot at determination P.
    """
    K = inp["K"]
    assets = inp["assets"]
    flows = inp["flows"]

    # Deduplicate (Period, AssetID): identical parsed payload -> one row;
    # conflicting payload -> LAST occurrence wins.
    dedup = {}
    for (p, aid), rows in flows.items():
        if p < 1 or p > K or aid not in assets:
            continue  # out-of-range period / unknown asset -> noise, dropped
        parsed = []
        for r in rows:
            parsed.append({
                "interest": parse_money(r["interest_raw"]),
                "principal": parse_money(r["principal_raw"]),
                "recovery": parse_money(r["recovery_raw"]),
                "endpar": None if r["endpar_raw"] is None or str(r["endpar_raw"]).strip() == ""
                          else parse_money(r["endpar_raw"], absolute=False),
            })
        kept = []
        for item in parsed:
            if any(item == k for k in kept):
                continue
            kept.append(item)
        dedup[(p, aid)] = kept[-1] if len(kept) > 1 else (kept[0] if kept else None)

    period_flows = {}
    par_snap = {}
    par = {aid: assets[aid]["par"] for aid in inp["asset_order"]}
    for P in range(1, K + 1):
        ip = ZERO
        pp = ZERO
        rec = ZERO
        for aid in inp["asset_order"]:
            item = dedup.get((P, aid))
            if item is not None:
                ip += item["interest"]
                if MUT.get("rec_to_interest"):
                    ip += item["recovery"]
                else:
                    pp += item["principal"] + item["recovery"]
                rec += item["recovery"]
                if item["endpar"] is not None:
                    par[aid] = item["endpar"]
                elif MUT.get("endpar_zero_fill"):
                    par[aid] = ZERO
            # no row -> no cash this period; par carries forward
        period_flows[P] = {"ip": ip, "pp": pp, "rec": rec}
        par_snap[P] = sum(par.values(), ZERO)
    return period_flows, par_snap


# --------------------------------------------------------------------------- #
# Note interest accrual
# --------------------------------------------------------------------------- #
def compute_rates(inp):
    """Per-class contractual rate sources. A2 uses the floor-to-grid clause on
    the period-1 base rate."""
    classes = inp["classes"]
    curve = inp["curve"]
    a2_rate = None
    if "A2" in classes:
        base1 = curve[1]
        if MUT.get("a2_no_grid"):
            a2_rate = base1 + inp["a2_spread"]
        else:
            a2_rate = floor_grid(base1, inp["grid_step"]) + inp["a2_spread"]
    rates = {}
    for cid, c in classes.items():
        if cid == "A2":
            rates[cid] = ("fixed", a2_rate)
        elif c["rate_type"] == FLOAT_TYPE:
            rates[cid] = ("float", c["spread"])
        else:
            rates[cid] = ("fixed", c["fixed"])
    return rates


def class_rate_for_P(inp, rates, cid, P):
    kind, v = rates[cid]
    if kind == "float":
        return inp["curve"][P] + v
    return v


def class_days(inp, cid, P):
    c = inp["classes"][cid]
    cur = inp["setdates"][P - 1]
    prev = inp["closing"] if P == 1 else inp["setdates"][P - 2]
    if c["day_count"] == "ACT360":
        return days_act(prev, cur)
    return days360(prev, cur)


# --------------------------------------------------------------------------- #
# Main deal engine
# --------------------------------------------------------------------------- #
def run_deal(inp):
    if MUT.get("reinvest_off_by_one"):
        months = inp["reinvest_months"] - 1
    else:
        months = inp["reinvest_months"]
    # Reinvestment end date = last day of the calendar month that is
    # `reinvest_months` months AFTER the closing month (closing month = month 0).
    m0 = inp["closing"].month - 1 + months
    y = inp["closing"].year + m0 // 12
    mm = m0 % 12 + 1
    reinvest_end = datetime.date(y, mm, calendar.monthrange(y, mm)[1])
    inp["reinvest_end"] = reinvest_end

    K = inp["K"]
    classes = inp["classes"]
    period_flows, par_snap = aggregate_flows(inp)
    rates = compute_rates(inp)

    note_bal_at_close = sum((classes[c]["orig_bal"] for c in CLASS_ORDER), ZERO)
    reserve_target = q2(inp["reserve_target_pct"] * note_bal_at_close)
    senior_fee_pct = inp["senior_fee_pct"]
    sub_fee_pct = inp["sub_fee_pct"]
    trustee_amt = inp["trustee"]
    reserve_rate = inp["reserve_rate"]
    # mutation hooks (experiments only): "no_latch" = divert exactly in failing
    # periods; "latch_release_1" = release after a single all-pass period.
    latch_periods = 1 if (MUT.get("no_latch") or MUT.get("latch_release_1")) \
        else inp["latch_periods"]

    # state entering P1
    bal = {c: classes[c]["orig_bal"] for c in CLASS_ORDER}
    unp_int = {c: ZERO for c in DRAW_CLASSES}
    exp_carry = ZERO
    sub_carry = ZERO
    unreimb = ZERO
    res_bal = inp["reserve_init"]
    carry = ZERO          # retained principal proceeds (reinvestment)
    block = False         # equity-block latch
    ever_on = False
    streak = 0
    unblock_period = 0
    prev_par = None       # OC numerator variant hook

    flows_rows = []
    notes_rows = []
    iw_rows = []
    pw_rows = []
    ct_rows = []
    rt_rows = []
    totals = {
        "ip": ZERO, "pp": ZERO, "rec": ZERO,
        "senior_fee": ZERO, "sub_fee": ZERO, "trustee": ZERO,
        "int_to_notes": ZERO, "pik": ZERO, "prin_to_notes": ZERO,
        "diverted": ZERO, "equity": ZERO,
        "draw": ZERO, "replenish": ZERO,
    }
    redemption_period = 0

    for P in range(1, K + 1):
        cur = inp["setdates"][P - 1]
        prev = inp["closing"] if P == 1 else inp["setdates"][P - 2]
        f = period_flows[P]
        IP = f["ip"]
        PP = f["pp"]
        PAR = par_snap[P]
        oc_numerator = PAR if not MUT.get("oc_prev_par") or P == 1 else par_snap[P - 1]

        # ---- interest accrual on beginning-of-period balances ----
        rate = {}
        due = {}
        adays = {}
        for c in CLASS_ORDER:
            rate[c] = class_rate_for_P(inp, rates, c, P)
            ad = class_days(inp, c, P)
            if MUT.get("act360_all"):
                ad = days_act(prev, cur)
            adays[c] = ad
            if bal[c] > 0:
                prod = bal[c] * rate[c] * Decimal(ad) / Decimal(360)
                due[c] = ZERO if prod == 0 else q2(prod)
                if MUT.get("floor_interest") and due[c] != 0:
                    due[c] = floor2(prod)
            else:
                due[c] = ZERO

        # ---- coverage tests (beginning-of-period balances) ----
        denA = bal["A1"] + bal["A2"] + bal["B"]
        denC = denA + bal["C"]
        denD = denC + bal["D"]
        denE = denD + bal["E"]
        ocA = q4(oc_numerator / denA) if denA > 0 else ZERO
        ocC = q4(oc_numerator / denC) if denC > 0 else ZERO
        ocD = q4(oc_numerator / denD) if denD > 0 else ZERO
        ocE = q4(oc_numerator / denE) if denE > 0 else ZERO
        duesen = due["A1"] + due["A2"] + due["B"]
        duesC = duesen + due["C"]
        icA = q4(IP / duesen) if duesen > 0 else ZERO
        icC = q4(IP / duesC) if duesC > 0 else ZERO
        # a zero denominator means the tested classes are fully repaid -> PASS
        fA = True if denA == 0 else ocA >= inp["oc_thr"]["A"]
        fC = True if denC == 0 else ocC >= inp["oc_thr"]["C"]
        fD = True if denD == 0 else ocD >= inp["oc_thr"]["D"]
        fE = True if denE == 0 else ocE >= inp["oc_thr"]["E"]
        fiA = True if duesen == 0 else icA >= inp["ic_thr"]["A"]
        fiC = True if duesC == 0 else icC >= inp["ic_thr"]["C"]
        # anyfail drives the equity-block latch: while ANY of the six tests
        # fails, the latch is ON.  The junior tests therefore carry teeth too:
        # their flicker keeps the latch ON and the diversion/retention running.
        anyfail = not (fA and fC and fD and fE and fiA and fiC)

        # ---- latch update (activation AND release take effect in the same
        #      period, before the diversion decision) ----
        if anyfail:
            streak = 0
            block = True
            ever_on = True
        else:
            streak += 1
            if block and streak >= latch_periods:
                block = False
                if unblock_period == 0:
                    unblock_period = P

        # ---- interest waterfall ----
        av = IP
        draw_tot = ZERO
        carry_paid = {c: ZERO for c in DRAW_CLASSES}
        exp_paid = ZERO
        sf_paid = ZERO
        tr_paid = ZERO
        int_paid = {c: ZERO for c in CLASS_ORDER}
        pik = {c: ZERO for c in CLASS_ORDER}
        div_pays = {c: ZERO for c in CLASS_ORDER}
        divert = ZERO
        replenish = ZERO
        reimb = ZERO
        topup = ZERO
        sf_carry_paid = ZERO
        sub_paid = ZERO
        equity = ZERO
        div_equity_fallback = ZERO

        # S1: repay prior-period interest shortfalls (A1 -> A2 -> B)
        for c in DRAW_CLASSES:
            pay = min(av, unp_int[c])
            av -= pay
            unp_int[c] -= pay
            carry_paid[c] = pay
        # S2: repay prior-period senior expense carryforward
        exp_paid = min(av, exp_carry)
        av -= exp_paid
        exp_carry -= exp_paid
        # S3: senior management fee (ROUND_DOWN); shortfall -> expense carry
        fee = floor2(IP * senior_fee_pct)
        sf_paid = min(av, fee)
        av -= sf_paid
        exp_carry += fee - sf_paid
        # S4: trustee expense (fixed); shortfall -> expense carry
        tr_paid = min(av, trustee_amt)
        av -= tr_paid
        exp_carry += trustee_amt - tr_paid
        # S5-S7: A1/A2/B interest; shortfall -> reserve draw; then carryforward
        for c in DRAW_CLASSES:
            cash = min(av, due[c])
            av -= cash
            short = due[c] - cash
            dr = ZERO
            if not MUT.get("no_reserve"):
                dr = min(res_bal, short)
                res_bal -= dr
                unreimb += dr
            cash += dr
            draw_tot += dr
            unp_int[c] += due[c] - cash
            int_paid[c] = cash
        # S8: coverage-test interest diversion while the latch is ON. While the
        # block is active ALL remaining interest proceeds are diverted to note
        # principal (senior pro-rata group first, then B..E, then any truly
        # unplaceable remainder to equity), so the junior classes below this
        # step receive no INTEREST cash this period.
        if block:
            divert_raw = av
            av = ZERO
            caps1 = [(c, bal[c]) for c in GROUP1]
            if sum(cap for _, cap in caps1) > 0:
                pays, leftover = allocate_pro_rata(divert_raw, caps1)
                for c in GROUP1:
                    div_pays[c] += pays[c]
            else:
                leftover = divert_raw
            if leftover > 0:
                # while blocked, C/D/E will PIK their full dues below (av is 0
                # from here on); B's shortfall never capitalises -> balance
                # capacity only for B.
                caps2 = []
                for c in SEQ_AFTER:
                    cap_c = bal[c] + (due[c] if c in DEFERRABLE else ZERO)
                    caps2.append((c, cap_c))
                pays2, leftover2 = allocate_pro_rata(leftover, caps2)
                for c in SEQ_AFTER:
                    div_pays[c] += pays2[c]
                leftover = leftover2
            # only the cash actually applied to note principal counts as the
            # period's diversion; unplaceable remainder flows to equity
            divert = sum(div_pays.values(), ZERO)
            if leftover > 0:
                div_equity_fallback = leftover
                equity += leftover

        # S9/S10: class C and D interest.  Both are deferrable; the reserve
        # also provides liquidity support for them while the block is OFF
        # (during a block there is no cash left and the reserve must not fund
        # the juniors -- their dues PIK instead).
        for c in ("C", "D"):
            cash = min(av, due[c])
            av -= cash
            short = due[c] - cash
            dr = ZERO
            if c in ("C", "D") and not block and not MUT.get("no_reserve"):
                dr = min(res_bal, short)
                res_bal -= dr
                unreimb += dr
            cash += dr
            draw_tot += dr
            int_paid[c] = cash
            if not MUT.get("no_pik"):
                pik[c] = due[c] - cash

        # S11: reserve step — reimburse draws, then top-up to target
        reimb = min(av, unreimb)
        av -= reimb
        unreimb -= reimb
        res_bal += reimb
        if not MUT.get("no_topup"):
            topup = min(av, max(ZERO, reserve_target - res_bal))
            av -= topup
            res_bal += topup
        # S12: subordinated management fee (ROUND_UP); carryforward pool
        sf_carry_paid = min(av, sub_carry)
        av -= sf_carry_paid
        sub_carry -= sf_carry_paid
        subfee = ceil2(IP * sub_fee_pct)
        sub_paid = min(av, subfee)
        av -= sub_paid
        sub_carry += subfee - sub_paid
        # S13: class E interest (deferrable; no reserve support)
        cash = min(av, due["E"])
        av -= cash
        int_paid["E"] = cash
        if not MUT.get("no_pik"):
            pik["E"] = due["E"] - cash
        # S14: equity residual
        equity += av
        av = ZERO

        # ---- principal waterfall ----
        # Four cases (latch x reinvestment):
        #   block ON  + reinvest : flush the RETAINED carry to the notes now;
        #                          the current period's collections are retained.
        #   block ON  + after    : pay all available principal to the notes.
        #   block OFF + reinvest : retain everything (carry_out = pv).
        #   block OFF + after    : pay all available principal to the notes.
        pv = PP + carry
        in_reinv = cur <= reinvest_end
        prin = {c: ZERO for c in CLASS_ORDER}
        prin_to_equity = ZERO
        carry_out = ZERO

        def pay_notes(remaining):
            """Pay `remaining` cash to the notes: group-1 pro rata, then B..E."""
            remaining = Decimal(remaining)
            if not MUT.get("seq_not_pro_rata"):
                caps1 = [(c, bal[c] + pik[c] - div_pays[c]) for c in GROUP1]
                pays, leftover = allocate_pro_rata(remaining, caps1)
                for c in GROUP1:
                    prin[c] = pays[c]
                remaining = leftover
            else:
                # mutated: pay A1 then A2 sequentially, no pro-rata
                for c in GROUP1:
                    cap = bal[c] + pik[c] - div_pays[c]
                    g = min(remaining, cap)
                    prin[c] = g
                    remaining -= g
            for c in SEQ_AFTER:
                cap = bal[c] + pik[c] - div_pays[c]
                g = min(remaining, cap)
                prin[c] = g
                remaining -= g
            return remaining

        if in_reinv and not block:
            # reinvestment: retain everything
            carry_out = pv
        elif in_reinv and block:
            # turbo: flush the retained carry now; retain current collections
            if carry > 0:
                prin_to_equity = pay_notes(carry)
            carry_out = PP
        else:
            leftover = pay_notes(pv)
            prin_to_equity = leftover
            carry_out = ZERO
            remaining = pv
            if not MUT.get("seq_not_pro_rata"):
                caps1 = [(c, bal[c] + pik[c] - div_pays[c]) for c in GROUP1]
                pays, leftover = allocate_pro_rata(remaining, caps1)
                for c in GROUP1:
                    prin[c] = pays[c]
                remaining = leftover
            else:
                for c in GROUP1:
                    cap = bal[c] + pik[c] - div_pays[c]
                    g = min(remaining, cap)
                    prin[c] = g
                    remaining -= g
            for c in SEQ_AFTER:
                cap = bal[c] + pik[c] - div_pays[c]
                g = min(remaining, cap)
                prin[c] = g
                remaining -= g
            prin_to_equity = remaining
            carry_out = ZERO

        # ---- balance roll-forward ----
        bal_end = {}
        for c in CLASS_ORDER:
            bal_end[c] = bal[c] - div_pays[c] - prin[c] + pik[c]
            if bal_end[c] < 0:
                raise AssertionError(f"negative balance {c} P{P}")

        # ---- reserve trace ----
        res_begin = None
        # reserve interest accrues on the BEGINNING-of-period reserve balance
        res_begin = rt_rows[-1]["end"] if rt_rows else inp["reserve_init"]
        r_int = floor2(res_begin * reserve_rate * Decimal(days_act(prev, cur)) / Decimal(360))
        if MUT.get("no_reserve"):
            r_int = ZERO
        res_pre_release = res_begin + r_int - draw_tot + reimb + topup
        release = ZERO
        if all(bal_end[c] == 0 for c in CLASS_ORDER):
            release = res_pre_release
        res_end = res_pre_release - release

        # ---- record rows ----
        flows_rows.append({
            "P": P, "ip": IP, "pp": PP, "rec": f["rec"], "par": PAR,
        })
        for c in CLASS_ORDER:
            notes_rows.append({
                "P": P, "class": c,
                "begin": bal[c], "rate": rate[c], "days": adays[c],
                "due": due[c], "paid": int_paid[c], "pik": pik[c],
                "prin": div_pays[c] + prin[c], "end": bal_end[c],
            })
        iw_rows.append({
            "P": P, "ip": IP, "draw": draw_tot,
            "cpA1": carry_paid["A1"], "cpA2": carry_paid["A2"], "cpB": carry_paid["B"],
            "exp_paid": exp_paid, "sf": sf_paid, "tr": tr_paid,
            "iA1": int_paid["A1"], "iA2": int_paid["A2"], "iB": int_paid["B"],
            "iC": int_paid["C"], "iD": int_paid["D"], "iE": int_paid["E"],
            "div": divert,
            "repl": reimb + topup, "sfcp": sf_carry_paid, "sub": sub_paid,
            "eq": equity,
        })
        pw_rows.append({
            "P": P, "pp": PP, "carry_in": carry, "avail": pv,
            "pA1": prin["A1"] + (div_pays["A1"] if False else ZERO),
            "pA2": prin["A2"],
            "pB": prin["B"], "pC": prin["C"], "pD": prin["D"], "pE": prin["E"],
            "carry_out": carry_out, "eq": prin_to_equity,
        })
        ct_rows.append({
            "P": P, "par": oc_numerator,
            "ocA": ocA, "ocC": ocC, "ocD": ocD, "icA": icA, "icC": icC,
            "fA": fA, "fC": fC, "fD": fD, "fE": fE, "fiA": fiA, "fiC": fiC,
            "anyfail": anyfail, "block": block,
            "ocE": ocE,
        })
        rt_rows.append({
            "P": P, "begin": res_begin, "int": r_int, "draw": draw_tot,
            "reimb": reimb, "topup": topup,
            "release": release, "end": res_end, "unreimb": unreimb,
        })

        # totals
        totals["ip"] += IP
        totals["pp"] += PP
        totals["rec"] += f["rec"]
        totals["senior_fee"] += sf_paid
        totals["sub_fee"] += sf_carry_paid + sub_paid
        totals["trustee"] += tr_paid
        totals["int_to_notes"] += sum(int_paid.values()) + sum(carry_paid.values())
        totals["pik"] += sum(pik.values())
        totals["prin_to_notes"] += sum(div_pays.values()) + sum(prin.values())
        totals["diverted"] += divert
        totals["equity"] += equity + prin_to_equity + release
        totals["draw"] += draw_tot
        totals["replenish"] += reimb + topup

        if redemption_period == 0 and all(bal_end[c] == 0 for c in CLASS_ORDER):
            redemption_period = P

        # state update for next period
        bal = bal_end
        carry = carry_out

    final_note_bal = sum(bal.values(), ZERO)
    summary = [
        ("ClosingDate", inp["closing"].strftime("%Y-%m-%d")),
        ("ReinvestEndDate", reinvest_end.strftime("%Y-%m-%d")),
        ("NoteBalanceAtClose", q2(note_bal_at_close)),
        ("CollateralParAtClose", q2(sum((inp["assets"][a]["par"] for a in inp["asset_order"]), ZERO))),
        ("ReserveTarget", reserve_target),
        ("TotalInterestProceeds", q2(totals["ip"])),
        ("TotalPrincipalProceeds", q2(totals["pp"])),
        ("TotalSeniorFees", q2(totals["senior_fee"])),
        ("TotalSubFees", q2(totals["sub_fee"])),
        ("TotalInterestToNotes", q2(totals["int_to_notes"])),
        ("TotalPIKCapitalized", q2(totals["pik"])),
        ("TotalPrincipalToNotes", q2(totals["prin_to_notes"])),
        ("TotalDivertedToPrincipal", q2(totals["diverted"])),
        ("TotalEquity", q2(totals["equity"])),
        ("FinalReserveBalance", q2(res_end)),
        ("FinalNoteBalance", q2(final_note_bal)),
        ("RedemptionPeriod", redemption_period),
        ("UnblockPeriod", unblock_period),
    ]

    # ---------------- internal self-consistency assertions ----------------
    for i, row in enumerate(iw_rows):
        lhs = row["ip"] + row["draw"]
        rhs = (row["cpA1"] + row["cpA2"] + row["cpB"] + row["exp_paid"]
               + row["sf"] + row["tr"] + row["iA1"] + row["iA2"] + row["iB"]
               + row["iC"] + row["iD"] + row["iE"] + row["div"] + row["repl"]
               + row["sfcp"] + row["sub"] + row["eq"])
        assert abs(lhs - rhs) <= Decimal("0.005"), f"IW identity P{i+1}: {lhs} vs {rhs}"
    for i, row in enumerate(pw_rows):
        lhs = row["pp"] + row["carry_in"]
        rhs = (row["pA1"] + row["pA2"] + row["pB"] + row["pC"] + row["pD"]
               + row["pE"] + row["carry_out"] + row["eq"])
        assert abs(lhs - rhs) <= Decimal("0.005"), f"PW identity P{i+1}: {lhs} vs {rhs}"
        assert row["avail"] == row["pp"] + row["carry_in"]
    for r in notes_rows:
        lhs = r["begin"] + r["pik"] - r["prin"]
        assert abs(lhs - r["end"]) <= Decimal("0.005"), f"NS roll {r['P']}/{r['class']}"
    for c in CLASS_ORDER:
        rows = [r for r in notes_rows if r["class"] == c]
        for i in range(1, len(rows)):
            assert rows[i]["begin"] == rows[i - 1]["end"], f"NS chain {c} P{i+1}"
    for i, r in enumerate(rt_rows):
        lhs = r["begin"] + r["int"] - r["draw"] + r["reimb"] + r["topup"] - r["release"]
        assert abs(lhs - r["end"]) <= Decimal("0.005"), f"RT roll P{i+1}"
        if i:
            assert r["begin"] == rt_rows[i - 1]["end"], f"RT chain P{i+1}"
    # Global cash conservation: every yuan sourced from IP + PP + reserve
    # interest + the initial reserve funding ends up paid out (fees / note
    # interest / note principal / equity streams), still in the reserve, or
    # retained in the principal collection account.
    #   iwfees = ExpenseCarryPaid + SeniorFee + Trustee + SubFeeCarryPaid + SubFee
    res_interest_total = sum(r["int"] for r in rt_rows)
    end_res = rt_rows[-1]["end"]
    iwfees = sum(r["exp_paid"] + r["sf"] + r["tr"] + r["sfcp"] + r["sub"]
                 for r in iw_rows)
    eq_iw = sum(r["eq"] for r in iw_rows)
    ptE = sum(r["eq"] for r in pw_rows)
    release_total = sum(r["release"] for r in rt_rows)
    lhs = (iwfees + totals["int_to_notes"] + totals["prin_to_notes"]
           + eq_iw + ptE + release_total + end_res + carry)
    rhs = (totals["ip"] + totals["pp"] + res_interest_total + inp["reserve_init"])
    assert abs(lhs - rhs) <= Decimal("0.05"), \
        f"global cash conservation fail: {float(lhs)} vs {float(rhs)}"

    return {
        "inp": inp,
        "flows_rows": flows_rows,
        "notes_rows": notes_rows,
        "iw_rows": iw_rows,
        "pw_rows": pw_rows,
        "ct_rows": ct_rows,
        "rt_rows": rt_rows,
        "summary": summary,
        "totals": totals,
        "final_res": res_end,
        "final_note": final_note_bal,
        "redemption_period": redemption_period,
        "unblock_period": unblock_period,
        "note_bal_at_close": note_bal_at_close,
        "reserve_target": reserve_target,
    }


# --------------------------------------------------------------------------- #
# Workbook builder (the answer format)
# --------------------------------------------------------------------------- #
SHEET_COLUMNS = {
    "PoolFlows": ["Period", "InterestProceeds", "PrincipalProceeds", "Recoveries", "CollateralPar"],
    "NoteSchedule": ["Period", "Class", "BeginBalance", "Rate", "AccDays", "InterestDue",
                     "InterestPaid", "InterestPIK", "PrincipalPaid", "EndBalance"],
    "InterestWaterfall": ["Period", "InterestProceeds", "ReserveDraw", "CarryPaid_A1",
                          "CarryPaid_A2", "CarryPaid_B", "ExpenseCarryPaid", "SeniorFee",
                          "Trustee", "Int_A1", "Int_A2", "Int_B", "Int_C", "Int_D",
                          "Int_E", "DivertedToPrincipal", "ReserveReplenish",
                          "SubFeeCarryPaid", "SubFee", "Equity"],
    "PrincipalWaterfall": ["Period", "PrincipalProceeds", "CarryIn", "TotalAvailable",
                           "Prin_A1", "Prin_A2", "Prin_B", "Prin_C", "Prin_D", "Prin_E",
                           "CarryOut", "PrinToEquity"],
    "CoverageTests": ["Period", "CollateralPar", "OC_A", "OC_C", "OC_D", "OC_E",
                      "IC_A", "IC_C",
                      "OC_A_Flag", "OC_C_Flag", "OC_D_Flag", "OC_E_Flag",
                      "IC_A_Flag", "IC_C_Flag",
                      "AnyFail", "EquityBlock"],
    "ReserveTrace": ["Period", "BeginBalance", "InterestCredited", "Draw", "Reimburse",
                     "TopUp", "Release", "EndBalance", "UnreimbursedEnd"],
}


def _f(x):
    """Decimal -> float for openpyxl (keeps cells numeric)."""
    if isinstance(x, Decimal):
        return float(x)
    return x


def build_workbook(res, path):
    import openpyxl
    wb = openpyxl.Workbook()

    ws = wb.active
    ws.title = "PoolFlows"
    ws.append(SHEET_COLUMNS["PoolFlows"])
    for r in res["flows_rows"]:
        ws.append([r["P"], _f(r["ip"]), _f(r["pp"]), _f(r["rec"]), _f(r["par"])])

    ws = wb.create_sheet("NoteSchedule")
    ws.append(SHEET_COLUMNS["NoteSchedule"])
    for r in res["notes_rows"]:
        ws.append([r["P"], r["class"], _f(r["begin"]), _f(r["rate"]), r["days"],
                   _f(r["due"]), _f(r["paid"]), _f(r["pik"]), _f(r["prin"]), _f(r["end"])])

    ws = wb.create_sheet("InterestWaterfall")
    ws.append(SHEET_COLUMNS["InterestWaterfall"])
    for r in res["iw_rows"]:
        ws.append([r["P"], _f(r["ip"]), _f(r["draw"]), _f(r["cpA1"]), _f(r["cpA2"]),
                   _f(r["cpB"]), _f(r["exp_paid"]), _f(r["sf"]), _f(r["tr"]),
                   _f(r["iA1"]), _f(r["iA2"]), _f(r["iB"]), _f(r["iC"]), _f(r["iD"]),
                   _f(r["iE"]), _f(r["div"]), _f(r["repl"]), _f(r["sfcp"]),
                   _f(r["sub"]), _f(r["eq"])])

    ws = wb.create_sheet("PrincipalWaterfall")
    ws.append(SHEET_COLUMNS["PrincipalWaterfall"])
    for r in res["pw_rows"]:
        ws.append([r["P"], _f(r["pp"]), _f(r["carry_in"]), _f(r["avail"]), _f(r["pA1"]),
                   _f(r["pA2"]), _f(r["pB"]), _f(r["pC"]), _f(r["pD"]), _f(r["pE"]),
                   _f(r["carry_out"]), _f(r["eq"])])

    ws = wb.create_sheet("CoverageTests")
    ws.append(SHEET_COLUMNS["CoverageTests"])
    for r in res["ct_rows"]:
        ws.append([r["P"], _f(r["par"]), _f(r["ocA"]), _f(r["ocC"]), _f(r["ocD"]),
                   _f(r["ocE"]), _f(r["icA"]), _f(r["icC"]),
                   "PASS" if r["fA"] else "FAIL", "PASS" if r["fC"] else "FAIL",
                   "PASS" if r["fD"] else "FAIL", "PASS" if r["fE"] else "FAIL",
                   "PASS" if r["fiA"] else "FAIL", "PASS" if r["fiC"] else "FAIL",
                   "FAIL" if r["anyfail"] else "PASS",
                   "ON" if r["block"] else "OFF"])

    ws = wb.create_sheet("ReserveTrace")
    ws.append(SHEET_COLUMNS["ReserveTrace"])
    for r in res["rt_rows"]:
        ws.append([r["P"], _f(r["begin"]), _f(r["int"]), _f(r["draw"]), _f(r["reimb"]),
                   _f(r["topup"]), _f(r["release"]), _f(r["end"]), _f(r["unreimb"])])

    ws = wb.create_sheet("DealSummary")
    ws.append(["Metric", "Value"])
    for k, v in res["summary"]:
        ws.append([k, _f(v)])

    wb.save(path)
    return path


def compute_from_file(input_path, output_path=None):
    inp = load_inputs(input_path)
    res = run_deal(inp)
    if output_path:
        build_workbook(res, output_path)
    return res


# --------------------------------------------------------------------------- #
# Diagnostics (authoring aid)
# --------------------------------------------------------------------------- #
def print_diagnostics(res):
    inp = res["inp"]
    print("=== KLC-2026-A diagnostics ===")
    print(f"K={inp['K']} reinvest_end={inp['reinvest_end']} "
          f"note_bal_at_close={res['note_bal_at_close']} target={res['reserve_target']}")
    print("thr  OC_A=%.4f OC_C=%.4f OC_D=%.4f IC_A=%.4f IC_C=%.4f" % (
        float(inp["oc_thr"]["A"]), float(inp["oc_thr"]["C"]), float(inp["oc_thr"]["D"]),
        float(inp["ic_thr"]["A"]), float(inp["ic_thr"]["C"])))
    hdr = ("P   IP        PP        PAR        denA      OC_A   OC_C   IC_A   IC_C   "
           "fails blk draw      pikCD      prinNotes  equity    resEnd")
    print(hdr)
    prev_block = False
    for i in range(inp["K"]):
        ct = res["ct_rows"][i]
        iw = res["iw_rows"][i]
        pw = res["pw_rows"][i]
        rt = res["rt_rows"][i]
        ns = [r for r in res["notes_rows"] if r["P"] == ct["P"]]
        denA = sum(r["begin"] for r in ns if r["class"] in ("A1", "A2", "B"))
        pikCD = sum(r["pik"] for r in ns if r["class"] in ("C", "D", "E"))
        prin = sum(r["prin"] for r in ns)
        fails = ""
        if not ct["fA"]:
            fails += "A"
        if not ct["fC"]:
            fails += "c"
        if not ct["fD"]:
            fails += "d"
        if not ct["fiA"]:
            fails += "i"
        if not ct["fE"]:
            fails += "e"
        if not ct["fiC"]:
            fails += "j"
        trans = ""
        if ct["block"] != prev_block:
            trans = " <-ON" if ct["block"] else " <-OFF"
        prev_block = ct["block"]
        print(f"{ct['P']:<3} {float(iw['ip']):>9} {float(pw['pp'] + pw['carry_in']):>9} "
              f"{float(ct['par'])/1e6:>8.1f}M {float(denA)/1e6:>8.1f}M {float(ct['ocA']):>6} "
              f"{float(ct['ocC']):>6} {float(ct['icA']):>6} {float(ct['icC']):>6} "
              f"{fails:<5} {'ON' if ct['block'] else 'off':<3} {float(iw['draw']):>9} "
              f"{float(pikCD):>10} {float(prin):>10} {float(iw['eq']):>9} "
              f"{float(rt['end']):>9}{trans}")
    sm = dict(res["summary"])
    print(f"redemption P{res['redemption_period']} unblock P{res['unblock_period']} "
          f"finalRes={sm['FinalReserveBalance']} finalNote={sm['FinalNoteBalance']} "
          f"totEquity={sm['TotalEquity']} diverted={sm['TotalDivertedToPrincipal']}")


def ct_rows_ip(iw):
    return iw["ip"]


if __name__ == "__main__":
    import sys
    inp_path = sys.argv[1] if len(sys.argv) > 1 else "/data/init.xlsx"
    out_path = sys.argv[2] if len(sys.argv) > 2 else None
    res = compute_from_file(inp_path, out_path)
    print_diagnostics(res)
    if out_path:
        print(f"oracle wrote {out_path}")
