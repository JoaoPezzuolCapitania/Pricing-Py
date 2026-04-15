"""
fpricing.py — Main pricing function, equivalent to VBA fPricing.
"""

import pandas as pd
from tqdm import tqdm

from .curves import CurveData, load_curves
from .bond import (
    BondInfo, CalcParams, BondResults, PeriodBond,
    load_bond, run_payments
)
from .solver import (
    get_par, get_taxa, get_spread_yield, get_pu,
    get_perc_pu_par, get_spread, duration,
    get_over_tp_with_curves
)


def _setup_calc(calc, i_info, d_value, bond):
    """Configures calculation type. Equivalent to VBA sTypeCalc."""
    calc.dPU = 0
    calc.dYield = 0
    calc.dSpread = 0
    calc.dPercPuPar = 0

    if i_info == 1:
        calc.dYield = d_value
        calc.sType = "Yield"
    elif i_info == 2:
        calc.dPU = d_value
        calc.sType = "PU"
    elif i_info == 3:
        calc.dPercPuPar = d_value
        calc.sType = "%PUPar"
    elif i_info == 4:
        calc.dSpread = d_value
        bond.sSpreadIndexInp = "CDI"
        calc.sType = "Taxa Spread"
    elif i_info == 5:
        calc.dSpread = d_value
        bond.sSpreadIndexInp = "IPCA"
        calc.sType = "Taxa Spread"
    elif i_info == 6:
        calc.dSpread = d_value
        bond.sSpreadIndexInp = "IGPM"
        calc.sType = "Taxa Spread"
    elif i_info == 7:
        calc.dSpread = d_value
        bond.sSpreadIndexInp = "IGPDI"
        calc.sType = "Taxa Spread"


def _setup_result(calc, i_result, bond):
    """Configures result type. Equivalent to VBA sTypeResult."""
    if i_result == 1:
        calc.sResult = "Yield"
    elif i_result == 2:
        calc.sResult = "PU"
    elif i_result == 3:
        calc.sResult = "%PUPar"
    elif i_result == 4:
        bond.sSpreadIndexRes = "CDI"
        calc.sResult = "Taxa Spread"
    elif i_result == 5:
        bond.sSpreadIndexRes = "IPCA"
        calc.sResult = "Taxa Spread"
    elif i_result == 6:
        bond.sSpreadIndexRes = "IGPM"
        calc.sResult = "Taxa Spread"
    elif i_result == 7:
        bond.sSpreadIndexRes = "IGPDI"
        calc.sResult = "Taxa Spread"
    elif i_result == 8:
        calc.sResult = "Duration"
    elif i_result == 9:
        calc.sResult = "OverB"
    elif i_result == 0:
        calc.sResult = "CashFlow"


def _get_result(i_result, results):
    """Returns result value by type. Equivalent to VBA fResult."""
    if i_result == 1:
        return results.dYield
    elif i_result == 2:
        return results.dPrice
    elif i_result == 3:
        return results.dPercPar
    elif i_result in (4, 5, 6, 7):
        return results.dSpread
    elif i_result == 8:
        return results.dDuration
    elif i_result == 9:
        return results.dOverTP
    return None


def fPricing(s_bond, dt_day, d_value, i_info, i_result, curves=None):
    """
    Exact replica of VBA fPricing function.

    Args:
        s_bond:   bond CETIP code (e.g. 'CBAN72')
        dt_day:   calculation date (str 'YYYY-MM-DD' or date object)
        d_value:  input value (PU, yield, spread, etc.)
        i_info:   input type
                    1 = Yield/Taxa
                    2 = PU
                    3 = %PUPar
                    4 = CDI Spread
                    5 = IPCA Spread
                    6 = IGPM Spread
                    7 = IGPDI Spread
        i_result: output type
                    1 = Yield/Taxa
                    2 = PU
                    3 = %PUPar
                    4 = CDI Spread
                    5 = IPCA Spread
                    8 = Duration
                    9 = OverTP
                    0 = CashFlow
        curves:   pre-loaded CurveData (optional; loads automatically if None)

    Returns:
        float or error string
    """
    from .daycount import _to_date
    dt_day = _to_date(dt_day)

    if curves is None:
        curves = load_curves(dt_day)

    result = load_bond(s_bond, dt_day)
    if result is None:
        return "Bond not found!"

    bond, calc, periods = result

    _setup_calc(calc, i_info, d_value, bond)
    _setup_result(calc, i_result, bond)

    pbs = run_payments(bond, calc, periods, curves)

    results = BondResults()
    results.dPar = get_par(bond, pbs)

    if calc.sType == "PU":
        get_taxa(calc, bond, periods, pbs, results)
    elif calc.sType == "Yield":
        get_pu(calc, bond, periods, pbs, results)
    elif calc.sType == "%PUPar":
        get_perc_pu_par(calc, results)
    elif calc.sType == "Taxa Spread":
        get_spread(calc, bond, periods, pbs, results)

    if calc.sResult == "Yield" and results.dYield == 0:
        get_taxa(calc, bond, periods, pbs, results)
    elif calc.sResult == "%PUPar":
        if results.dPar != 0:
            results.dPercPar = results.dPrice / results.dPar
    elif calc.sResult == "Taxa Spread":
        get_spread_yield(calc, bond, periods, pbs, results)
    elif calc.sResult in ("Duration", "OverB"):
        if results.dYield == 0:
            get_taxa(calc, bond, periods, pbs, results)
        duration(calc, bond, periods, pbs, results)
        get_over_tp_with_curves(calc, bond, results, curves)

    return _get_result(i_result, results)


def fPricing_batch(df, dt, inp=2, curves=None):
    """
    Batch pricing for multiple bonds.
    Loads curves ONCE and processes all bonds.

    Args:
        df:     DataFrame with columns 'cod_cetip' (or 'ativo') and 'pu'
        dt:     calculation date (str 'YYYY-MM-DD')
        inp:    input type (default 2 = PU)
        curves: pre-loaded CurveData (optional)

    Returns:
        DataFrame with additional columns: taxa, spread, duration, over_tp
    """
    from .daycount import _to_date
    dt = _to_date(dt)

    if curves is None:
        print("Loading curves...")
        curves = load_curves(dt)
        print("Curves loaded!")

    results_taxa = []
    results_spread = []
    results_duration = []
    results_over_tp = []

    for i in tqdm(range(len(df)), desc="Pricing"):
        cetip = df.iloc[i]['cod_cetip'] if 'cod_cetip' in df.columns else df.iloc[i]['ativo']
        pu = df.iloc[i]['pu']

        taxa = fPricing(cetip, dt, pu, inp, 1, curves)
        spread = fPricing(cetip, dt, pu, inp, 4, curves)
        dur = fPricing(cetip, dt, pu, inp, 8, curves)
        over = fPricing(cetip, dt, pu, inp, 9, curves)

        results_taxa.append(taxa)
        results_spread.append(spread)
        results_duration.append(dur)
        results_over_tp.append(over)

    df = df.copy()
    df['taxa'] = results_taxa
    df['spread'] = results_spread
    df['duration'] = results_duration
    df['over_tp'] = results_over_tp

    return df
