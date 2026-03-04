#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Internet benchmark check for rc_beam_optimizer.py

Sources:
1) SkyCiv ACI-318 verification table:
   https://skyciv.com/docs/tech-notes/verification/aci-318-verification/
   - Book 1, Example 4.1 (manual): design shear strength = 8750 lb (8.75 kip)
   - Book 1, Example 4.2 (manual): design shear strength = 83.6 kip
"""

from __future__ import annotations

import rc_beam_optimizer as m

KIP_TO_KN = 4.4482216152605


def run_case(name, expected, got, unit, tol_pct):
    err = (got - expected) / expected * 100.0 if expected else 0.0
    ok = abs(err) <= tol_pct
    print(
        f"{name}: expected={expected:.3f} {unit}, got={got:.3f} {unit}, "
        f"err={err:+.2f}% -> {'PASS' if ok else 'FAIL'}"
    )
    return ok


def main():
    cost = {"c_conc": 0.0, "c_steel": 0.0, "c_form": 0.0, "e_conc": 0.0, "e_steel": 0.0, "e_form": 0.0}
    setts = {"gD": 1.0, "gL": 1.0, "mode": "COST", "w_cost": 1.0, "w_co2": 0.0, "max_rows": 1000, "all_checks": True}
    beam = {"id": "B", "span": 6.0, "trib": 0.0, "qD": 0.0, "qL": 0.0, "PD": 0.0, "PL": 0.0, "r": 0.5}

    # Example 4.1 analogue (no stirrups)
    # Target from source: 8.75 kip
    mat_41 = {
        "fc": 27.579, "fy": 413.685, "fyv": 413.685, "cover": 57.15, "gamma": 0.0,
        "phi_m": 0.9, "phi_v": 0.75, "rho_max": 0.025, "fs_lim": 1e9,
        "defl_ratio": 1e9, "ei_fac": 1.0, "s_clear_min": 25.0, "n_div": 200,
    }
    cand_41 = {
        "rank": 1, "sec": "EX41", "b": 260.35, "h": 520.85,
        "nb": 5, "db": "D13", "nt": 0, "dt": "",
        "legs": 0, "ds": "", "s": 0.0,
    }
    r41 = m.eval_cand(beam, cand_41, mat_41, setts, cost)
    phi_vn_41_kip = r41["phiVn"] / KIP_TO_KN

    # Example 4.2 analogue (with stirrups)
    # Target from source (manual): 83.6 kip
    mat_42 = {
        "fc": 27.579, "fy": 413.685, "fyv": 413.685, "cover": 63.5, "gamma": 0.0,
        "phi_m": 0.9, "phi_v": 0.75, "rho_max": 0.025, "fs_lim": 1e9,
        "defl_ratio": 1e9, "ei_fac": 1.0, "s_clear_min": 25.0, "n_div": 200,
    }
    cand_42 = {
        "rank": 1, "sec": "EX42", "b": 457.2, "h": 921.2,
        "nb": 7, "db": "D19", "nt": 0, "dt": "",
        "legs": 2, "ds": "D10", "s": 304.8,
    }
    r42 = m.eval_cand(beam, cand_42, mat_42, setts, cost)
    phi_vn_42_kip = r42["phiVn"] / KIP_TO_KN

    ok = True
    ok &= run_case("Ex4.1 shear (no stirrups)", 8.75, phi_vn_41_kip, "kip", 2.0)
    ok &= run_case("Ex4.2 shear (with stirrups)", 83.6, phi_vn_42_kip, "kip", 3.0)
    print("Overall:", "PASS" if ok else "FAIL")
    return 0 if ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
