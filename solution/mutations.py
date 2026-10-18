#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Mutation variants for the discrimination experiments (authoring tool only).

Each variant is a dict of oracle_engine.MUT flags representing a plausible
single-clause misreading of the instruction. The shipped oracle runs with an
empty MUT dict.
"""

# Primary variants: single-clause misreadings used for the discrimination
# gate (each must score reward < 0.5 against the golden).
VARIANTS = {
    # pro-rata group A1/A2 amortised sequentially instead of pro rata
    "seq_not_pro_rata": {"seq_not_pro_rata": True},
    # reinvestment end date taken one calendar month earlier
    "reinvest_off_by_one": {"reinvest_off_by_one": True},
    # A2 coupon taken without the floor-to-grid clause
    "a2_no_grid": {"a2_no_grid": True},
    # deferrable classes' unpaid interest dropped instead of capitalised
    "no_pik": {"no_pik": True},
    # post-default recoveries booked as interest proceeds instead of principal
    "rec_to_interest": {"rec_to_interest": True},
    # reserve account never draws / tops up
    "no_reserve": {"no_reserve": True},
    # OC numerator taken from the prior period's par snapshot
    "oc_prev_par": {"oc_prev_par": True},
    # every class accrued on Act/360 regardless of its DayCount column
    "act360_all": {"act360_all": True},
    # blank EndPar treated as zero instead of carried forward
    "endpar_zero_fill": {"endpar_zero_fill": True},
    # reserve top-up step dropped
    "no_topup": {"no_topup": True},
}

# Composite variants: pairs of independent single-clause misreadings.  This is
# the agent-realistic error profile (an agent that misimplements the spec in a
# 20-minute window almost never errs in exactly one lever).
COMPOSITE = {
    "act360_all+no_pik": {"act360_all": True, "no_pik": True},
    "a2_no_grid+seq": {"a2_no_grid": True, "seq_not_pro_rata": True},
    "rec_to_int+no_topup": {"rec_to_interest": True, "no_topup": True},
    "reinvest+no_pik": {"reinvest_off_by_one": True, "no_pik": True},
    "act360+no_reserve": {"act360_all": True, "no_reserve": True},
    "a2_no_grid+no_topup": {"a2_no_grid": True, "no_topup": True},
}

# Secondary variants: real misreadings whose cash stakes are structurally
# bounded (they shift only a small tail of the trajectory).  Reported for
# diagnostics; not part of the < 0.5 discrimination gate.
SECONDARY = {
    "no_latch": {"no_latch": True},
    "latch_release_1": {"latch_release_1": True},
    "floor_interest": {"floor_interest": True},
    "no_junior_gate": {"no_junior_gate": True},
}
