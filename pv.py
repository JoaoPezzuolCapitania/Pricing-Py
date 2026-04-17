"""
pv.py — Cálculos de valor presente.
Replica sPvCalc, sPvSpreadInp, sPvSpreadRes, fGetSpread, fGetSpreadPerc, fGetPrice.
"""

from .daycount import brworkdays, brworkday, fDays, _to_date
from datetime import timedelta


def fGetSpread(i_run, d_yield, calc, periods):
    """
    Fator de desconto para uma taxa (convenção de dias).
    Replica fGetSpread do VBA.
    """
    base = 1 + d_yield
    if base <= 0:
        return 1e-10  # Protege contra 0^negativo

    if calc.sYdays == "Úteis":
        if calc.dtDay > periods[i_run - 1].dtDay and calc.dtDay < periods[i_run].dtDay:
            i_days = brworkdays(calc.dtDay, periods[i_run].dtDay) - 1
        else:
            i_days = brworkdays(periods[i_run - 1].dtDay, periods[i_run].dtDay) - 1
        return (1 + d_yield) ** (i_days / 252)

    elif calc.sYdays == "30/360":
        if calc.dtDay > periods[i_run - 1].dtDay and calc.dtDay < periods[i_run].dtDay:
            i_days = (periods[i_run].dtDay - calc.dtDay).days
        else:
            i_days = (periods[i_run].dtDay - periods[i_run - 1].dtDay).days
        i_month_days = fDays(periods[i_run].dtDay, "DC")
        return ((1 + d_yield) ** (30 / 360)) ** (i_days / i_month_days) if i_month_days else 1.0

    elif calc.sYdays == "1/360":
        if calc.dtDay > periods[i_run - 1].dtDay and calc.dtDay < periods[i_run].dtDay:
            i_days = (periods[i_run].dtDay - calc.dtDay).days
        else:
            i_days = (periods[i_run].dtDay - periods[i_run - 1].dtDay).days
        return ((1 + d_yield) ** (1 / 360)) ** i_days

    elif calc.sYdays == "1/365":
        if calc.dtDay > periods[i_run - 1].dtDay and calc.dtDay < periods[i_run].dtDay:
            i_days = (periods[i_run].dtDay - calc.dtDay).days
        else:
            i_days = (periods[i_run].dtDay - periods[i_run - 1].dtDay).days
        return ((1 + d_yield) ** (1 / 365)) ** i_days

    elif calc.sYdays == "21/252":
        if calc.dtDay > periods[i_run - 1].dtDay and calc.dtDay < periods[i_run].dtDay:
            i_days = brworkdays(calc.dtDay, periods[i_run].dtDay) - 1
        else:
            i_days = brworkdays(periods[i_run - 1].dtDay, periods[i_run].dtDay) - 1
        i_month_days = fDays(periods[i_run].dtDay, "DU")
        return ((1 + d_yield) ** (21 / 252)) ** (i_days / i_month_days) if i_month_days else 1.0

    return 1.0


def fGetSpreadPerc(i_run, d_yield, calc, periods, pbs):
    """
    Fator de desconto para % CDI.
    Replica fGetSpreadPerc do VBA.
    """
    i_days = brworkdays(periods[i_run - 1].dtDay, periods[i_run].dtDay) - 1

    if calc.dtDay > periods[i_run - 1].dtDay:
        dt_wd = brworkday(periods[i_run].dtDay - timedelta(days=1), 1)
        i_d = brworkdays(calc.dtDay, dt_wd) - 1
    else:
        i_d = i_days

    if i_d == 0:
        i_d = 1

    return ((d_yield) * ((pbs[i_run].dYdi1) ** (1 / i_d) - 1) + 1) ** (i_d / 1)


def pv_calc(i_run, d_pv_yield, bond, calc, periods, pbs):
    """
    Calcula PV de um período com yield arbitrária.
    Replica sPvCalc do VBA.
    """
    if calc.dtDay > periods[i_run].dtDay or i_run == 0:
        pbs[i_run].dPVfactorCalc = 1.0
        return

    if calc.dtDay == periods[i_run].dtDay or i_run == 0:
        pbs[i_run].dPVfactorCalc = 1.0
        if calc.bPmtIncorp:
            pbs[i_run].dPVpmtCalc = pbs[i_run].dPMTTotal / pbs[i_run].dPVfactorCalc if pbs[i_run].dPVfactorCalc else 0
        return

    # Factor
    if '+' in bond.sIndex or bond.sIndex == "Pré":
        pbs[i_run].dPVfactorCalc = pbs[i_run - 1].dPVfactorCalc * fGetSpread(i_run, d_pv_yield, calc, periods)
    elif '%' in bond.sIndex:
        pbs[i_run].dPVfactorCalc = pbs[i_run - 1].dPVfactorCalc * fGetSpreadPerc(i_run, d_pv_yield, calc, periods, pbs)

    # Adjust DI1
    if bond.sIndex == "CDI +":
        pbs[i_run].dPVfactorCalc *= pbs[i_run].dYdi1
    elif '%' not in bond.sIndex:
        pbs[i_run].dPVfactorCalc *= pbs[i_run].dYinf

    # PMT
    pbs[i_run].dPVpmtCalc = pbs[i_run].dPMTTotal / pbs[i_run].dPVfactorCalc if pbs[i_run].dPVfactorCalc else 0


def pv_spread_inp(i_run, d_pv_yield, bond, calc, periods, pbs):
    """
    Calcula PV para spread input.
    Replica sPvSpreadInp do VBA.
    """
    if calc.dtDay > periods[i_run].dtDay or i_run == 0:
        pbs[i_run].dPVfactorSpread = 1.0
        return

    if calc.dtDay == periods[i_run].dtDay or i_run == 0:
        pbs[i_run].dPVfactorSpread = 1.0
        if calc.bPmtIncorp:
            pbs[i_run].dPVpmtSpread = pbs[i_run].dPMTTotal / pbs[i_run].dPVfactorSpread if pbs[i_run].dPVfactorSpread else 0
        return

    pbs[i_run].dPVfactorSpread = pbs[i_run - 1].dPVfactorSpread * fGetSpread(i_run, d_pv_yield, calc, periods)

    if bond.sSpreadIndexInp == "CDI":
        pbs[i_run].dPVfactorSpread *= pbs[i_run].dYdi1
    else:
        pbs[i_run].dPVfactorSpread *= pbs[i_run].dYAuxCurveInp

    pbs[i_run].dPVpmtSpread = pbs[i_run].dPMTTotal / pbs[i_run].dPVfactorSpread if pbs[i_run].dPVfactorSpread else 0


def pv_spread_res(i_run, d_pv_yield, bond, calc, periods, pbs):
    """
    Calcula PV para spread result.
    Replica sPvSpreadRes do VBA.
    """
    if calc.dtDay > periods[i_run].dtDay or i_run == 0:
        pbs[i_run].dPVfactorSpread = 1.0
        return

    if calc.dtDay == periods[i_run].dtDay or i_run == 0:
        pbs[i_run].dPVfactorSpread = 1.0
        if calc.bPmtIncorp:
            pbs[i_run].dPVpmtSpread = pbs[i_run].dPMTTotal / pbs[i_run].dPVfactorSpread if pbs[i_run].dPVfactorSpread else 0
        return

    pbs[i_run].dPVfactorSpread = pbs[i_run - 1].dPVfactorSpread * fGetSpread(i_run, d_pv_yield, calc, periods)

    if bond.sSpreadIndexRes in ("CDI", ""):
        pbs[i_run].dPVfactorSpread *= pbs[i_run].dYdi1
    else:
        pbs[i_run].dPVfactorSpread *= pbs[i_run].dYAuxCurveRes

    pbs[i_run].dPVpmtSpread = pbs[i_run].dPMTTotal / pbs[i_run].dPVfactorSpread if pbs[i_run].dPVfactorSpread else 0


def get_price(d_yield, s_type, bond, calc, periods, pbs):
    """
    Calcula preço total dado yield/spread. Replica fGetPrice.
    """
    d_sum = 0.0
    for i in range(len(periods)):
        if s_type == "Taxa":
            pv_calc(i, d_yield, bond, calc, periods, pbs)
            d_sum += pbs[i].dPVpmtCalc
        elif s_type == "Spread":
            pv_spread_res(i, d_yield, bond, calc, periods, pbs)
            d_sum += pbs[i].dPVpmtSpread
    return d_sum
