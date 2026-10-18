#!/usr/bin/env python3
"""Authoring sweep: try a few crisis/controller configs, pick the one that
minimizes the worst tested-variant reward. Not part of the graded container."""
import itertools
import os
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
sys.path.insert(0, os.path.join(HERE, "..", "tests"))

import gen_data as G
import oracle_engine as eng
from mutations import VARIANTS
import test_outputs as T

CONFIGS = [
    dict(name="H1", coupons=(0.0500, 0.0560), res_init=500_000, res_pct=0.007,
         a2=199_000_000, cds=(0.0825, 0.0450, 0.1250),
         forced=[(6, 10_000_000), (12, 34_000_000), (16, 18_000_000)], first=5),
    dict(name="H2", coupons=(0.0500, 0.0560), res_init=500_000, res_pct=0.010,
         a2=199_000_000, cds=(0.0825, 0.0450, 0.1250),
         forced=[(6, 10_000_000), (12, 34_000_000), (16, 18_000_000)], first=5),
    dict(name="H3", coupons=(0.0500, 0.0560), res_init=400_000, res_pct=0.009,
         a2=205_000_000, cds=(0.0825, 0.0450, 0.1250),
         forced=[(6, 10_000_000), (12, 34_000_000), (16, 18_000_000)], first=5),
]


def run_config(cfg):
    G.FORCED_CRISIS = cfg["forced"]
    G.CTRL_FIRST_PERIOD = cfg["first"]
    G.COUPON_LO, G.COUPON_HI = cfg.get("coupons", (0.0500, 0.0560))
    if "res_init" in cfg:
        PARAMS = [(k, cfg["res_init"] if k == "ReserveInitial" else
                   (cfg["res_pct"] if k == "ReserveTargetPct" else v))
                  for k, v in G.PARAMS_FIXED]
        G.PARAMS_FIXED = PARAMS
        cds = cfg["cds"]
        G.CLASSES = [("A1", 1, 20_000_000.00, "FLOAT", None, 0.0250, "ACT360", "N", 1),
                     ("A2", 2, cfg.get("a2", 190_000_000) * 1.0, "FIXED", None, None, "30360", "N", 1),
                     ("B", 3, 30_000_000.00, "FIXED", 0.0690, None, "30360", "N", None),
                     ("C", 4, 18_000_000.00, "FIXED", cds[0], None, "30360", "Y", None),
                     ("D", 5, 14_000_000.00, "FLOAT", None, cds[1], "ACT360", "Y", None),
                     ("E", 6, 8_000_000.00, "FIXED", cds[2], None, "30360", "Y", None)]
    setdates = G.build_settlement_dates()
    assets = G.build_assets()
    res, plans, rows, defaults, thresholds = G.calibrate(assets, setdates, verbose=False)
    rows = G.apply_dirty(rows)
    path = os.path.join(tempfile.gettempdir(), f"sweep_{cfg['name']}.xlsx")
    G.write_workbook(assets, plans, rows, setdates, thresholds, path)
    eng.MUT.clear()
    baseline = eng.compute_from_file(path, None)
    out = {"oracle": None}
    # oracle self-check
    tmp = path + ".g.xlsx"
    eng.build_workbook(baseline, tmp)
    wb = openpyxl_load(tmp)
    checks = T.grade(wb, baseline)
    tot = sum(w for _, _, w in checks)
    ok = sum(w for _, o, w in checks if o)
    out["oracle"] = ok / tot
    rewards = {}
    for name, flags in VARIANTS.items():
        eng.MUT.clear()
        eng.MUT.update(flags)
        r = eng.compute_from_file(path, None)
        t2 = os.path.join(tempfile.gettempdir(), f"sweep_v_{cfg['name']}_{name}.xlsx")
        eng.build_workbook(r, t2)
        wb2 = openpyxl_load(t2)
        checks = T.grade(wb2, baseline)
        tot2 = sum(w for _, _, w in checks)
        ok2 = sum(w for _, o, w in checks if o)
        rewards[name] = ok2 / tot2
    eng.MUT.clear()
    return out["oracle"], rewards, res


def openpyxl_load(path):
    import openpyxl
    return openpyxl.load_workbook(path, data_only=True)


if __name__ == "__main__":
    best = None
    for cfg in CONFIGS:
        oracle_r, rewards, res = run_config(cfg)
        worst = max(rewards.values())
        sm = dict(res["summary"])
        print(f"cfg {cfg['name']}: oracle={oracle_r:.4f} worst={worst:.4f} "
              f"red=P{sm['RedemptionPeriod']} unblk=P{sm['UnblockPeriod']} "
              f"diverted={sm['TotalDivertedToPrincipal']}")
        for k, v in sorted(rewards.items(), key=lambda kv: -kv[1]):
            print(f"    {k:<22} {v:.4f}")
        if oracle_r >= 0.9999 and (best is None or worst < best[0]):
            best = (worst, cfg["name"])
    print("BEST:", best)
