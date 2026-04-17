"""
solver.py — Newton-Raphson para taxa e spread, duration, overTP.
Replica sGetTaxa, sGetSpreadYield, sDuration, sGetOverTP, sGetPar, sGetPU, etc.
"""

from .pv import pv_calc, pv_spread_inp, pv_spread_res, get_price
from .daycount import brworkdays, brworkday


def get_par(bond, pbs):
    """Calcula valor par. Replica sGetPar."""
    d_par = sum(pb.dPVpmtPar for pb in pbs)
    return round(d_par, 6)


def _safe_real(v):
    """Extrai parte real se complexo, clamp se infinito."""
    if isinstance(v, complex):
        v = v.real
    if v != v:  # NaN
        return 0.0
    if abs(v) > 1e15:
        return 0.0
    return v


def get_taxa(calc, bond, periods, pbs, results):
    """
    Newton-Raphson: PU → yield.
    Replica sGetTaxa do VBA (sem clamp de yield).
    """
    i_times = 0
    d_yield = 0.0
    d_price = 0.0
    d_delta = 1e-10

    while True:
        d_price = _safe_real(get_price(d_yield, "Taxa", bond, calc, periods, pbs))

        if round(d_price, 6) == round(float(calc.dPU), 6):
            break

        # TryAgain loop: replica GoTo TryAgain do VBA
        d_try = 0
        while True:
            d_yield += d_delta
            d_price_1 = _safe_real(get_price(d_yield, "Taxa", bond, calc, periods, pbs))
            d_derivada = (d_price - d_price_1) / d_delta
            if d_derivada != 0:
                break
            d_yield -= d_delta
            d_delta *= 10
            d_try += 1
            if d_try > 50:
                break

        if d_derivada == 0:
            break

        d_yield = d_yield + ((d_price - calc.dPU) / d_derivada)
        d_yield = _safe_real(d_yield)

        i_times += 1
        if i_times > 20:
            break

    results.dYield = round(float(d_yield), 6)
    results.dPrice = round(float(d_price), 6)
    if results.dPar != 0:
        results.dPercPar = round(results.dPrice / results.dPar, 15)

    get_spread_yield(calc, bond, periods, pbs, results)


def get_spread_yield(calc, bond, periods, pbs, results):
    """
    Newton-Raphson: PU → spread.
    Replica sGetSpreadYield do VBA (sem clamp de yield).
    """
    i_times = 0
    d_yield = 0.0
    d_price = 0.0
    d_delta = 1e-10

    while True:
        d_price = _safe_real(get_price(d_yield, "Spread", bond, calc, periods, pbs))

        if round(d_price, 6) == round(float(calc.dPU), 6):
            break

        # TryAgain loop: replica GoTo TryAgain do VBA
        d_try = 0
        while True:
            d_yield += d_delta
            d_price_1 = _safe_real(get_price(d_yield, "Spread", bond, calc, periods, pbs))
            d_derivada = (d_price - d_price_1) / d_delta
            if d_derivada != 0:
                break
            d_yield -= d_delta
            d_delta *= 10
            d_try += 1
            if d_try > 50:
                break

        if d_derivada == 0:
            break

        d_yield = d_yield + ((d_price - calc.dPU) / d_derivada)
        d_yield = _safe_real(d_yield)

        i_times += 1
        if i_times > 20:
            break

    results.dSpread = round(float(d_yield), 15)
    get_over_tp(calc, bond, results, pbs)


def get_pu(calc, bond, periods, pbs, results):
    """
    Yield → PU (forward pricing).
    Replica sGetPU do VBA.
    """
    d_price = 0.0
    for i in range(len(periods)):
        pv_calc(i, calc.dYield, bond, calc, periods, pbs)
        d_price += pbs[i].dPVpmtCalc

    results.dPrice = round(d_price, 6)
    results.dYield = round(calc.dYield, 6)
    if results.dPar != 0:
        results.dPercPar = round(results.dPrice / results.dPar, 15)
    calc.dPU = round(results.dPrice, 6)


def get_perc_pu_par(calc, results):
    """% PU Par → PU. Replica sGetPercPuPar."""
    calc.dPU = round(results.dPar * calc.dPercPuPar, 6)
    results.dPrice = calc.dPU


def get_spread(calc, bond, periods, pbs, results):
    """
    Spread → PU (forward pricing com spread).
    Replica sGetSpread do VBA.
    """
    d_price = 0.0
    for i in range(len(periods)):
        pv_spread_inp(i, calc.dSpread, bond, calc, periods, pbs)
        d_price += pbs[i].dPVpmtSpread

    results.dSpread = round(calc.dSpread, 6)
    results.dPrice = round(d_price, 6)
    calc.dPU = round(results.dPrice, 6)
    if results.dPar != 0:
        results.dPercPar = round(results.dPrice / results.dPar, 15)


def duration(calc, bond, periods, pbs, results):
    """Calcula Macaulay duration. Replica sDuration."""
    d_sum_dur = 0.0
    for i in range(len(periods)):
        i_days = brworkdays(calc.dtDay, periods[i].dtDay) - 1
        d_sum_dur += pbs[i].dPVpmtCalc * i_days

    if results.dPrice != 0:
        results.dDurationMacaulay = round((d_sum_dur / results.dPrice) / 252, 6)
    results.dDuration = results.dDurationMacaulay


def get_over_tp(calc, bond, results, curves_or_pbs=None):
    """
    Calcula spread sobre benchmark (Tesouro Prefixado).
    Replica sGetOverTP do VBA.
    Nota: precisa das curvas para lookup — recebe via global ou parâmetro.
    """
    # Este método será chamado pelo fpricing.py que passa as curvas
    pass


def get_over_tp_with_curves(calc, bond, results, curves):
    """Versão completa de sGetOverTP com acesso às curvas."""
    dur_days = int(results.dDuration * 252)
    dt_target = brworkday(calc.dtDay, dur_days)
    dt_key = str(dt_target)

    if bond.sIndex == "% CDI":
        di1_rate = curves.dicDI1.get(dt_key, list(curves.dicDI1.values())[-1] if curves.dicDI1 else 0)
        results.dOverTP = (results.dYield - 1) * di1_rate
    elif bond.sIndex == "IPCA +":
        ntnb_rate = curves.dicNTNB.get(dt_key, list(curves.dicNTNB.values())[-1] if curves.dicNTNB else 0)
        results.dOverTP = results.dYield - ntnb_rate
    elif bond.sIndex in ("IGPM +", "IGPDI +"):
        igp_rate = curves.dicIGP.get(dt_key, list(curves.dicIGP.values())[-1] if curves.dicIGP else 0)
        results.dOverTP = results.dYield - igp_rate
    elif bond.sIndex == "Pré":
        di1_rate = curves.dicDI1.get(dt_key, list(curves.dicDI1.values())[-1] if curves.dicDI1 else 0)
        results.dOverTP = results.dYield - di1_rate
    elif bond.sIndex == "CDI +":
        results.dOverTP = results.dYield
