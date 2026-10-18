#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
gen_data.py — fixed-seed generator for the coding-exam-002 input workbook
/data/init.xlsx (KLC-2026-A wind-down CCO).

Output: environment/sheet/init.xlsx  (relative to this file's parent).

The deal is calibrated so that the coverage tests (OC_A/OC_C/OC_D/IC_A/IC_C)
OSCILLATE AROUND THEIR THRESHOLDS through the middle of the deal: a
default-scheduling controller reads back the oracle engine's ratio path and
places asset defaults to push the OC_A ratio below its threshold whenever it
drifts too far above it.  The resulting flickering triggers exercise the whole
interlocked machinery (equity-block latch with 2-pass release, interest
diversion, reinvestment turbo, PIK capitalisation, reserve draws) and makes any
single-clause misreading diverge catastrophically.

Everything is seeded; rerunning reproduces the workbook.
"""

import os
import random
import datetime
import tempfile
from decimal import Decimal

import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
sys_root = os.path.dirname(HERE)
OUT = os.path.join(HERE, "..", "environment", "sheet", "init.xlsx")

SEED = 20260118
R = random.Random(SEED)

# ----------------------------- TUNE section -------------------------------- #
N_ASSETS = 64
TOTAL_PAR = 380_000_000.00
COUPON_LO, COUPON_HI = 0.0500, 0.0560          # annual, on par
CLOSING = datetime.date(2025, 1, 20)
K = 24                                          # determination dates
# sale pacing: (first_period, last_period, monthly_target); P1-P3 quiet;
# period 23 force-liquidates whatever is left.  Roughly 25-35M of performing
# par is left outstanding at the end (the notes are called while part of the
# pool is still paying).
SALE_PROFILE = [(4, 9, 1_200_000), (10, 14, 5_000_000), (15, 22, 33_000_000)]
RECOVERY_RATE_LO, RECOVERY_RATE_HI = 0.44, 0.52
RECOVERY_LAG_START, RECOVERY_LAG_N = 2, 6       # start d+2, spread over 6 months
# controller knobs
CTRL_FIRST_PERIOD = 5           # no defaults scheduled before this period
CTRL_LAST_PERIOD = 20           # no defaults scheduled after this period
CTRL_MAX_TOTAL_PAR = 170_000_000  # cap on cumulative defaulted par
# forced crisis defaults: schedule assets defaulting together so 2-3 periods
# later interest proceeds dip below the senior interest dues -> reserve draws
FORCED_CRISIS = [(6, 10_000_000), (12, 34_000_000), (16, 18_000_000)]
CTRL_TRIGGER = 0.007            # schedule a default if OC_A >= thr + this
CTRL_OVERSHOOT = 0.004          # target OC_A lands thr - this after a default
RECOVERY_R4 = Decimal("0.0001")

CLASSES = [
    # ClassID, Priority, OrigBalance, RateType, Fixed, Spread, DayCount, Deferrable, Group1
    ("A1", 1,  20_000_000.00, "FLOAT",  None,   0.0250, "ACT360", "N", 1),
    ("A2", 2, 190_000_000.00, "FIXED",  None,   None,   "30360",  "N", 1),  # rate via grid clause
    ("B",  3,  30_000_000.00, "FIXED",  0.0690, None,   "30360",  "N", None),
    ("C",  4,  18_000_000.00, "FIXED",  0.0825, None,   "30360",  "Y", None),
    ("D",  5,  14_000_000.00, "FLOAT",  None,   0.0450, "ACT360", "Y", None),
    ("E",  6,   8_000_000.00, "FIXED",  0.1250, None,   "30360",  "Y", None),
]

BASE_THRESHOLDS = {
    "OCThreshold_A": 1.45, "OCThreshold_C": 1.28, "OCThreshold_D": 1.12,
    "OCThreshold_E": 1.05, "ICThreshold_A": 1.158, "ICThreshold_C": 1.02,
}

PARAMS_FIXED = [
    ("ClosingDate", "20-Jan-2025"),
    ("ReinvestMonths", 12),
    ("SeniorFeePct", 0.004),
    ("SubFeePct", 0.0025),
    ("TrusteeExpense", 12500.00),
    ("ReserveInitial", 500000.00),
    ("ReserveTargetPct", 0.007),
    ("ReserveRate", 0.03),
    ("LatchReleasePeriods", 2),
    ("GridStep", 0.00125),
    ("A2Spread", 0.0225),
]
# --------------------------------------------------------------------------- #

D = Decimal
Q2 = D("0.01")
ZERO = D(0)


def r2(x):
    return D(x).quantize(Q2)


def build_settlement_dates():
    ds = []
    y, m = CLOSING.year, CLOSING.month
    for i in range(1, K + 1):
        mm = m + i
        yy = y + (mm - 1) // 12
        mm = (mm - 1) % 12 + 1
        ds.append(datetime.date(yy, mm, CLOSING.day))
    return ds


def build_assets():
    """64 assets, total par exactly TOTAL_PAR, coupons, sectors, regions."""
    sizes = []
    for _ in range(N_ASSETS):
        if R.random() < 0.22:
            sizes.append(R.uniform(9_000_000, 15_000_000))
        else:
            sizes.append(R.uniform(2_500_000, 8_500_000))
    scale = TOTAL_PAR / sum(sizes)
    sizes = [r2(s * scale) for s in sizes]
    drift = r2(TOTAL_PAR) - sum(sizes)
    sizes[0] += drift
    assert sum(sizes) == r2(TOTAL_PAR), sum(sizes)

    assets = []
    sectors = ["CRE", "SME", "ABS", "Shipping", "Aviation", "Infra", "Consumer"]
    regions = ["NA", "EMEA", "APAC", "LATAM"]
    for i, par in enumerate(sizes, start=1):
        assets.append({
            "id": f"L-{i:03d}",
            "par": par,
            "coupon": D(str(round(R.uniform(COUPON_LO, COUPON_HI), 4))),
            "sector": R.choice(sectors),
            "region": R.choice(regions),
        })
    return assets


def assign_events(assets, K, defaults):
    """Per-asset event plan keyed by asset id.  `defaults` maps asset id to the
    period in which the asset defaults (par -> 0 that period)."""
    plans = {a["id"]: {"sales": [], "default": defaults.get(a["id"]),
                       "recoveries": []} for a in assets}

    # ---- sales: greedy pacing toward the profile; liquidate by P23 ----
    remaining = {a["id"]: a["par"] for a in assets}
    order = [a["id"] for a in assets]
    R.shuffle(order)
    for per in range(1, 24):
        target = ZERO
        for lo, hi, monthly in SALE_PROFILE:
            if lo <= per <= hi:
                target = D(str(monthly))
        if per == 23:
            target = None  # force full liquidation of what is left
        if target == 0:
            continue
        left = target
        for aid in order:
            if left is not None and left <= 0:
                break
            rem = remaining[aid]
            if rem <= 0:
                continue
            dp = plans[aid]["default"]
            if dp is not None and per >= dp:
                continue
            if left is None:
                amt = rem
            else:
                amt = min(rem, left)
            plans[aid]["sales"].append((per, amt))
            remaining[aid] -= amt
            if left is not None:
                left -= amt

    # ---- recoveries for defaulted assets ----
    for a in assets:
        p = plans[a["id"]]
        if p["default"] is None:
            continue
        rate = R.uniform(RECOVERY_RATE_LO, RECOVERY_RATE_HI)
        tot = r2(a["par"] * D(str(rate)))
        base = p["default"] + RECOVERY_LAG_START
        w = [R.uniform(0.5, 1.5) for _ in range(RECOVERY_LAG_N)]
        tw = sum(w)
        acc = D(0)
        for j in range(RECOVERY_LAG_N):
            per = min(base + j, 24)
            amt = tot - acc if j == RECOVERY_LAG_N - 1 else r2(D(str(w[j] / tw)) * tot)
            acc += amt
            if amt > 0:
                p["recoveries"].append((per, amt))
    return plans


def build_flows(assets, plans, setdates):
    """Rows: (Period, AssetID, AsOfDate, InterestCollected, PrincipalCollected,
    RecoveryCollected, EndPar). Missing cells -> None. Missing rows -> nothing."""
    K = len(setdates)
    par = {a["id"]: a["par"] for a in assets}
    sale_by = {}
    rec_by = {}
    for a in assets:
        p = plans[a["id"]]
        for per, amt in p["sales"]:
            sale_by.setdefault((per, a["id"]), D(0))
            sale_by[(per, a["id"])] += amt
        for per, amt in p["recoveries"]:
            rec_by.setdefault((per, a["id"]), D(0))
            rec_by[(per, a["id"])] += amt

    rows = []
    for per in range(1, K + 1):
        for a in assets:
            aid = a["id"]
            par_begin = par[aid]
            sold = sale_by.get((per, aid), D(0))
            rec = rec_by.get((per, aid), D(0))
            default_here = plans[aid]["default"] == per

            if par_begin > 0:
                if default_here:
                    ip = r2(par_begin * a["coupon"] / D(12) * D(str(R.uniform(0.40, 0.60))))
                else:
                    ip = r2(par_begin * a["coupon"] / D(12) * D(str(R.uniform(0.98, 1.02))))
            else:
                ip = D(0)

            par_end = par_begin - sold
            if default_here:
                par_end = D(0)

            endpar_cell = None
            if par_end != par_begin or per == 1:
                endpar_cell = par_end
                if per == 1 and R.random() < 0.18:
                    endpar_cell = None  # exercise carry-forward from ParAtClose

            if ip == 0 and sold == 0 and rec == 0 and endpar_cell is None:
                continue  # sparse row: no activity, par carries forward

            asof = format_asof(setdates[per - 1])
            rows.append({
                "P": per, "aid": aid, "asof": asof,
                "ip": ip if ip != 0 else None,
                "pp": sold if sold != 0 else None,
                "rec": rec if rec != 0 else None,
                "endpar": endpar_cell,
            })
            par[aid] = par_end
    return rows


def format_asof(d):
    k = d.month % 3
    if k == 0:
        return d.strftime("%d-%b-%Y")
    if k == 1:
        return d.strftime("%m/%d/%Y")
    return d.strftime("%Y-%m-%d")


def apply_dirty(rows):
    """Inject noise: negatives, blanks, comma strings, duplicates, unknown
    AssetIDs, out-of-range periods."""
    R2 = random.Random(SEED + 7)
    idx_candidates = [i for i, r in enumerate(rows) if isinstance(r["ip"], D) and r["ip"] > 0]
    neg = R2.sample(idx_candidates, 3)
    for i in neg:
        rows[i]["ip"] = -rows[i]["ip"]                      # noise -> abs
    blank = R2.sample([i for i in idx_candidates if i not in neg], 5)
    for i in blank:
        rows[i]["ip"] = None                                # blank -> 0
    comma_c = [i for i, r in enumerate(rows)
               if isinstance(r["endpar"], D) and r["endpar"] > 0]
    for i in R2.sample(comma_c, 4):
        rows[i]["endpar"] = f"{rows[i]['endpar']:,.2f}"
    ip_c = [i for i, r in enumerate(rows) if isinstance(r["ip"], D) and r["ip"] > 0]
    for i in R2.sample(ip_c, 3):
        rows[i]["ip"] = f"{rows[i]['ip']:,.2f}"
    pp_c = [i for i, r in enumerate(rows) if r["pp"]]
    for i in R2.sample(pp_c, 2):
        rows[i]["pp"] = -abs(rows[i]["pp"])
    rc_c = [i for i, r in enumerate(rows) if r["rec"]]
    if rc_c:
        rows[R2.choice(rc_c)]["rec"] = -abs(rows[R2.choice(rc_c)]["rec"])
    dup_pool = [i for i, r in enumerate(rows)
                if isinstance(r["ip"], D) and r["ip"] > 0 and r["pp"] is None]
    picks = R2.sample(dup_pool, 4)
    for j, i in enumerate(picks):
        copy = dict(rows[i])
        if j < 2:
            rows.insert(i + 1, copy)          # exact duplicate
        else:
            v = copy["ip"]
            copy["ip"] = v + D("100.00")
            rows.insert(i + 1, copy)          # conflicting duplicate (last wins)
    rows.append({"P": 7, "aid": "L-099", "asof": "2025-08-20", "ip": D("123456.78"),
                 "pp": None, "rec": None, "endpar": None})
    rows.append({"P": 0, "aid": "L-001", "asof": "2025-01-20", "ip": D("9999.99"),
                 "pp": None, "rec": None, "endpar": None})
    rows.append({"P": 25, "aid": "L-002", "asof": "2027-02-20", "ip": D("8888.88"),
                 "pp": None, "rec": None, "endpar": None})
    return rows


def write_workbook(assets, plans, rows, setdates, thresholds, path):
    wb = openpyxl.Workbook()

    ws = wb.active
    ws.title = "Params"
    ws.append(["Key", "Value"])
    for k, v in PARAMS_FIXED:
        ws.append([k, v])
    for k, v in sorted(thresholds.items()):
        ws.append([k, v])
    sd = ";".join(d.strftime("%d-%b-%Y") for d in setdates)
    ws.append(["SettlementDates", sd])

    ws = wb.create_sheet("Classes")
    ws.append(["ClassID", "Priority", "OrigBalance", "RateType", "FixedRate",
               "Spread", "DayCount", "Deferrable", "ProRataGroup"])
    for cid, pr, bal, rt, fx, sp, dc, df, grp in CLASSES:
        ws.append([cid, pr, float(bal), rt,
                   float(fx) if fx is not None else None,
                   float(sp) if sp is not None else None,
                   dc, df, grp])

    ws = wb.create_sheet("RateCurve")
    ws.append(["Period", "BaseRate"])
    base0 = 0.0437
    for p in range(1, K + 1):
        wiggle = 0.0002 * ((p * 7) % 3 - 1)
        v = round(base0 - 0.00028 * (p - 1) + wiggle, 4)
        ws.append([p, v])

    ws = wb.create_sheet("AssetTape")
    ws.append(["AssetID", "ParAtClose", "Coupon", "Sector", "Region"])
    for a in assets:
        ws.append([a["id"], float(a["par"]), float(a["coupon"]), a["sector"], a["region"]])

    ws = wb.create_sheet("PeriodFlows")
    ws.append(["Period", "AssetID", "AsOfDate", "InterestCollected",
               "PrincipalCollected", "RecoveryCollected", "EndPar"])
    rows = sorted(rows, key=lambda r: (r["P"], r["aid"]))
    for r in rows:
        ws.append([
            r["P"], r["aid"], r["asof"],
            float(r["ip"]) if isinstance(r["ip"], D) else r["ip"],
            float(r["pp"]) if isinstance(r["pp"], D) else r["pp"],
            float(r["rec"]) if isinstance(r["rec"], D) else r["rec"],
            float(r["endpar"]) if isinstance(r["endpar"], D) else r["endpar"],
        ])

    os.makedirs(os.path.dirname(path), exist_ok=True)
    wb.save(path)
    return path


# --------------------------------------------------------------------------- #
# Calibration controller: schedule defaults so OC_A oscillates around its
# threshold, then place the remaining thresholds on the realised paths.
# --------------------------------------------------------------------------- #
def simulate(assets, defaults, setdates, thresholds):
    plans = assign_events(assets, K, defaults)
    rows = build_flows(assets, plans, setdates)
    path = os.path.join(tempfile.gettempdir(), "klc_calib.xlsx")
    write_workbook(assets, plans, rows, setdates, thresholds, path)
    import sys
    sys.path.insert(0, HERE)
    import oracle_engine as eng
    res = eng.compute_from_file(path, None)
    return res, plans, rows


def calibrate(assets, setdates, verbose=True):
    import sys
    sys.path.insert(0, HERE)
    import oracle_engine as eng

    defaults = {}       # aid -> default period
    thresholds = dict(BASE_THRESHOLDS)
    order = sorted(assets, key=lambda a: -float(a["par"]))
    defaulted_par = D(0)
    # forced crisis defaults (largest assets first until the target is met)
    for period, target in FORCED_CRISIS:
        acc = D(0)
        for a in order:
            if a["id"] in defaults:
                continue
            if acc >= D(str(target)):
                break
            defaults[a["id"]] = period
            defaulted_par += a["par"]
            acc += a["par"]

    for it in range(10):
        res, plans, rows = simulate(assets, defaults, setdates, thresholds)
        ct = res["ct_rows"]
        ns = res["notes_rows"]
        thr_a = D(str(thresholds["OCThreshold_A"]))

        # --- schedule new defaults where OC_A drifted too far above threshold ---
        added = []
        by_P = {r["P"]: r for r in ct}
        den_by_P = {}
        for r in ns:
            if r["class"] in ("A1", "A2", "B"):
                den_by_P[r["P"]] = den_by_P.get(r["P"], ZERO) + r["begin"]
        for P in range(max(2, CTRL_FIRST_PERIOD), CTRL_LAST_PERIOD):
            if len(added) and P <= max(a[1] for a in added):
                continue
            row = by_P.get(P)
            if row is None:
                continue
            gap = row["ocA"] - thr_a
            if gap >= D(str(CTRL_TRIGGER)) and defaulted_par + D(0) <= D(str(CTRL_MAX_TOTAL_PAR)):
                # want OC_A(P) ~= thr - overshoot after a default lands at P
                nxt = by_P.get(P + 1)
                if nxt is None:
                    continue
                target = nxt["par"] - (thr_a - D(str(CTRL_OVERSHOOT))) * den_by_P[P + 1]
                if target <= 0:
                    continue
                target = min(target, D(str(CTRL_MAX_TOTAL_PAR)) - defaulted_par)
                # pick the closest-par asset not yet defaulted and not defaulted too late
                best, bestdiff = None, None
                for a in order:
                    if a["id"] in defaults:
                        continue
                    if a["par"] < D("2000000"):
                        continue
                    diff = abs(a["par"] - target)
                    if bestdiff is None or diff < bestdiff:
                        best, bestdiff = a, diff
                if best is None:
                    continue
                defaults[best["id"]] = P
                defaulted_par += best["par"]
                added.append((best["id"], P))
        # --- recalibrate the other thresholds onto the realised paths ---
        if not added or it >= 3:
            ocs = [r["ocC"] for r in ct if 1 <= r["P"] <= 9]
            ocDs = [r["ocD"] for r in ct if 1 <= r["P"] <= 9]
            ocEs = [r["ocE"] for r in ct if 1 <= r["P"] <= 9]
            ics = [r["icA"] for r in ct if 1 <= r["P"] <= 9]
            icCs = [r["icC"] for r in ct if 1 <= r["P"] <= 9]
            if ocs and ics:
                # just below the clean-window minimum: the tests start failing
                # as the pool deteriorates, then oscillate with the noise band
                thresholds["OCThreshold_C"] = float(round(min(ocs) - D("0.028"), 4))
                thresholds["OCThreshold_D"] = float(round(min(ocDs) - D("0.045"), 4))
                thresholds["OCThreshold_E"] = float(round(min(ocEs) - D("0.030"), 4))
                thresholds["ICThreshold_A"] = float(round(min(ics) - D("0.032"), 4))
                thresholds["ICThreshold_C"] = float(round(min(icCs) - D("0.030"), 4))
                thresholds["OCThreshold_A"] = float(round(thr_a, 4))
        if verbose:
            print(f"  iter {it}: scheduled {len(added)} defaults "
                  f"(cum par {float(defaulted_par)/1e6:.1f}M) "
                  f"thr={ {k: round(v, 4) for k, v in thresholds.items()} }")
        if not added and it >= 3:
            break

    res, plans, rows = simulate(assets, defaults, setdates, thresholds)
    return res, plans, rows, defaults, thresholds


def validate():
    """Authoring-time sanity checks on the generated workbook."""
    import sys
    sys.path.insert(0, HERE)
    import oracle_engine as eng

    res = eng.compute_from_file(os.path.normpath(OUT), None)
    bad = []
    for fr in res["flows_rows"]:
        cents = int((fr["ip"] * 100).to_integral_value())
        if fr["ip"] != 0 and (cents % 250 == 0 or cents % 400 == 0):
            bad.append((fr["P"], cents))
    if bad:
        print("WARN: IP divisible by fee denominators:", bad)
    eng.print_diagnostics(res)
    return res


if __name__ == "__main__":
    setdates = build_settlement_dates()
    assets = build_assets()
    print("calibrating default schedule ...")
    res, plans, rows, defaults, thresholds = calibrate(assets, setdates)
    print(f"defaults scheduled: {len(defaults)} assets")
    rows = apply_dirty(rows)
    # nudge interest by 1 cent where a period's InterestProceeds would be an
    # exact multiple of 250 / 400 cents (fee-rounding boundary safety)
    import sys
    sys.path.insert(0, HERE)
    import oracle_engine as eng
    for attempt in range(6):
        write_workbook(assets, plans, rows, setdates, thresholds, OUT)
        probe = eng.compute_from_file(OUT, None)
        bad = []
        for fr in probe["flows_rows"]:
            cents = int((fr["ip"] * 100).to_integral_value())
            if fr["ip"] != 0 and (cents % 250 == 0 or cents % 400 == 0):
                bad.append(fr["P"])
        if not bad:
            break
        for P in bad:
            cands = [r for r in rows if r["P"] == P and isinstance(r["ip"], D) and r["ip"] > 0]
            if cands:
                cands[0]["ip"] += D("0.01")
        print(f"  IP divisibility fixup round {attempt}: nudged periods {bad}")
    print(f"wrote {OUT}  rows={len(rows)}  thresholds={thresholds}")
    validate()
