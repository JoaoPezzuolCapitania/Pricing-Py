"""
bond.py — Bond data loading and payment schedule generation.
Equivalent to VBA fLoadBondInfo + sRunPayments + helper subs.

Configure the bond data directory via environment variable:
    BOND_DATA_PATH  — directory containing {CETIP}.txt files
                      Defaults to './data/bonds'
"""

import os
import re
from datetime import date, timedelta
from dataclasses import dataclass, field
from typing import List, Optional
from .daycount import (
    _to_date, brworkday, brworkdays, fDays, fFactorDays,
    next_month, n_months, lag_month_key, lag_date, am_month_check
)

BOND_DATA_PATH = os.environ.get('BOND_DATA_PATH', os.path.join(os.path.dirname(__file__), 'data', 'bonds'))


@dataclass
class BondInfo:
    sCETIP: str = ""
    sType: str = ""
    dtIssuance: date = None
    dtMaturity: date = None
    sIndex: str = ""
    dYield: float = 0.0
    dPU: float = 0.0
    iPeriods: int = 0
    sSpreadIndexInp: str = ""
    sSpreadIndexRes: str = ""
    dicYields: dict = field(default_factory=dict)


@dataclass
class CalcParams:
    sType: str = ""
    sResult: str = ""
    dtDay: date = None
    dYield: float = 0.0
    dSpread: float = 0.0
    dPU: float = 0.0
    dPercPuPar: float = 0.0
    iAMlag: int = 0
    iCDIlag: int = 0
    iPmtlag: int = 0
    bAMTotal: bool = False
    sAMmonth: str = ""
    sAMcarac: str = ""
    bAmort100: bool = True
    bPmtIncorp: bool = False
    sYdays: str = ""
    dAccInflFactor: float = 0.0
    dError: float = 0.0
    dFatorAMacc: float = 0.0


@dataclass
class PeriodInfo:
    dtDay: date = None
    dtDayPMT: date = None
    dIncorpYield: float = 0.0
    dAmort: float = 0.0
    dExtrAmort: float = 0.0
    dMultaFee: float = 0.0


@dataclass
class PeriodBond:
    dSN: float = 0.0
    dSNA: float = 0.0
    dFatAm: float = 1.0
    dFatAmAcc: float = 1.0
    dYinf: float = 1.0
    dYcdi: float = 1.0
    dYdi1: float = 1.0
    dYAuxCurveInp: float = 1.0
    dYAuxCurveRes: float = 1.0
    dYspread: float = 1.0
    dYtotal: float = 1.0
    dPMTJuros: float = 0.0
    dPMTIncorpJuros: float = 0.0
    dPMTAmort: float = 0.0
    dPMTAmortExtr: float = 0.0
    dPMTAMprincipal: float = 0.0
    dPMTTotal: float = 0.0
    dPMTMulta: float = 0.0
    dPVfactorPar: float = 1.0
    dPVfactorCalc: float = 1.0
    dPVfactorSpread: float = 1.0
    dPVpmtPar: float = 0.0
    dPVpmtCalc: float = 0.0
    dPVpmtSpread: float = 0.0


@dataclass
class BondResults:
    dPar: float = 0.0
    dYield: float = 0.0
    dSpread: float = 0.0
    dOverTP: float = 0.0
    dPrice: float = 0.0
    dPercPar: float = 0.0
    dDurationMacaulay: float = 0.0
    dDuration: float = 0.0


def _parse_date_bond(s):
    """
    Parses date from bond file. Supports three formats:
      MM/DD/YYYY — US format (VBA default)
      DD/MM/YYYY — detected when first part > 12
      YYYY-MM-DD — ISO format
    """
    s = s.strip()
    if not s:
        return None

    if '-' in s:
        parts = s.split('-')
        if len(parts) != 3:
            return None
        try:
            y, m, d = int(parts[0]), int(parts[1]), int(parts[2])
            return date(y, m, d)
        except (ValueError, IndexError):
            return None

    if '/' not in s:
        return None
    parts = s.split('/')
    if len(parts) != 3:
        return None
    try:
        a, b, c = int(parts[0]), int(parts[1]), int(parts[2])
        if a > 12:
            return date(c, b, a)
        if b > 12:
            return date(c, a, b)
        return date(c, a, b)
    except (ValueError, IndexError):
        return None


def _load_yield(s_yield, dt_calc):
    """Parses rate with possible scheduled changes (step-up). Equivalent to VBA fLoadYield."""
    s_yield = str(s_yield).strip()
    dic = {}

    if '#' in s_yield:
        splitter = '#'
    elif '@' in s_yield:
        splitter = '@'
    else:
        try:
            return float(s_yield) if s_yield else 0.0, dic
        except ValueError:
            return 0.0, dic

    entries = s_yield.split('_')
    first_yield = None
    last_yield = None
    for entry in entries:
        parts = entry.split(splitter)
        if len(parts) < 2:
            continue
        dt = _parse_date_bond(parts[0])
        if dt is None:
            continue
        if splitter == '#' or (splitter == '@' and dt < dt_calc):
            val_str = parts[1].strip()
            if val_str.endswith('%'):
                val_str = val_str[:-1]
            try:
                last_yield = float(val_str) / 100
                if first_yield is None:
                    first_yield = last_yield
                dic[str(dt)] = last_yield
            except ValueError:
                continue

    return first_yield if first_yield is not None else 0.0, dic


def load_bond(cetip, dt_calc):
    """
    Loads bond information from a .txt file.
    Equivalent to VBA fLoadBondInfo.

    The bond file is looked up at: {BOND_DATA_PATH}/{cetip}.txt

    Returns: (BondInfo, CalcParams, List[PeriodInfo]) or None if not found.

    Bond file format (tab-separated key-value pairs, then payment schedule):

        Código CETIP    XXXX99
        Tipo            CRI
        Data de Emissão 01/15/2020
        Vencimento      01/15/2030
        Indexador       CDI +
        Taxa            0.012
        PU              1000.00
        Dias (juros)    Úteis
        Atraso AM (meses)   0
        Atraso PMT (du) 0
        AM Total        False
        Data
        01/15/2021  0   0.1   0   0
        01/15/2022  0   0.1   0   0
        ...

    Payment schedule columns: date, incorp_yield, amort, extra_amort, fee
    """
    dt_calc = _to_date(dt_calc)
    filepath = os.path.join(BOND_DATA_PATH, f"{cetip}.txt")

    if not os.path.exists(filepath):
        return None

    bond = BondInfo()
    calc = CalcParams(dtDay=dt_calc)
    dates_dict = {}

    with open(filepath, 'r', encoding='latin-1') as f:
        lines = f.readlines()

    i = 0
    while i < len(lines):
        line = lines[i].strip()
        i += 1
        if not line:
            continue

        parts = line.split('\t')
        name_raw = parts[0].strip()
        value = parts[1].strip() if len(parts) > 1 else ""
        aux_a = parts[2].strip() if len(parts) > 2 else ""
        aux_b = parts[3].strip() if len(parts) > 3 else ""

        name = re.sub(r'\s+', ' ', name_raw)
        _name_map = {
            "CódigoCETIP": "Código CETIP",
            "CodigoCETIP": "Código CETIP",
            "DatadeEmissão": "Data de Emissão",
            "DatadeEmissao": "Data de Emissão",
            "Dias(juros)": "Dias (juros)",
            "AtrasoAM(meses)": "Atraso AM (meses)",
            "AtrasoPMT(du)": "Atraso PMT (du)",
            "AMTotal": "AM Total",
        }
        name = _name_map.get(name, name)

        def _safe_int(v, default=0):
            try:
                return int(v) if v else default
            except (ValueError, TypeError):
                return default

        def _safe_float(v, default=0.0):
            try:
                return float(v) if v else default
            except (ValueError, TypeError):
                return default

        if name == "Código CETIP":
            bond.sCETIP = value
        elif name == "Tipo":
            bond.sType = value
        elif name == "Data de Emissão":
            bond.dtIssuance = _parse_date_bond(value)
        elif name == "Vencimento":
            bond.dtMaturity = _parse_date_bond(value)
        elif name == "Indexador":
            bond.sIndex = value
        elif name == "Taxa":
            bond.dYield, bond.dicYields = _load_yield(value, dt_calc)
        elif name == "PU":
            bond.dPU = _safe_float(value)
        elif name == "Dias (juros)":
            calc.sYdays = value
        elif name == "Atraso AM (meses)":
            if "CDI" not in bond.sIndex:
                calc.iAMlag = _safe_int(value)
            else:
                calc.iCDIlag = _safe_int(value)
        elif name == "Atraso PMT (du)":
            calc.iPmtlag = _safe_int(value)
        elif name == "AM Total":
            calc.bAMTotal = value.lower() == 'true' if value else False
            calc.sAMmonth = aux_a
            calc.sAMcarac = aux_b
        elif name == "Data":
            while i < len(lines):
                line = lines[i].strip()
                i += 1
                if not line:
                    break
                parts = line.split('\t')
                if len(parts) >= 5:
                    dt_key = parts[0].strip()
                    incorp = float(parts[1])
                    amort = float(parts[2])
                    extra = float(parts[3])
                    multa = float(parts[4])
                    dates_dict[dt_key] = (incorp, amort, extra, multa)

    periods = []
    for dt_str, (incorp, amort, extra, multa) in dates_dict.items():
        dt_parsed = _parse_date_bond(dt_str)
        if dt_parsed is None:
            continue

        pi = PeriodInfo()
        pi.dIncorpYield = incorp
        pi.dAmort = amort
        pi.dExtrAmort = extra
        pi.dMultaFee = multa
        pi.dtDay = dt_parsed

        if calc.iPmtlag == 0:
            pi.dtDayPMT = brworkday(pi.dtDay - timedelta(days=1), 1)
        else:
            if pi.dMultaFee != 0 and pi.dAmort == 0:
                pi.dtDayPMT = brworkday(pi.dtDay - timedelta(days=1), 1)
            else:
                pi.dtDayPMT = brworkday(pi.dtDay, calc.iPmtlag)

        periods.append(pi)

    dt_adjust = None
    for idx in range(1, len(periods)):
        if (periods[idx].dIncorpYield == 1 and
                periods[idx - 1].dIncorpYield == 0 and
                "CDI" in bond.sIndex):
            dt_adjust = periods[idx].dtDayPMT

    if dt_adjust and dt_calc < dt_adjust:
        periods[-1].dAmort = 1

    bond.iPeriods = len(periods)

    if not periods:
        return bond, calc, periods

    if periods[-1].dAmort == 1:
        first_amort_idx = 0
        while (first_amort_idx < len(periods) and
               periods[first_amort_idx].dAmort == 0 and
               periods[first_amort_idx].dExtrAmort == 0):
            first_amort_idx += 1
        calc.bAmort100 = (first_amort_idx == len(periods) - 1)
    else:
        calc.bAmort100 = True
        extra_amort_adj = 0.0
        for idx, pi in enumerate(periods):
            extra_amort_adj += pi.dAmort
            if dt_calc >= pi.dtDay:
                extra_amort_adj += pi.dExtrAmort / bond.dPU if bond.dPU != 0 else 0
        if extra_amort_adj != 1:
            periods[-1].dAmort += (1 - extra_amort_adj)

    calc.bPmtIncorp = False

    if bond.dPU < 1000:
        calc.dError = 0.00000005
    elif bond.dPU < 1000000:
        calc.dError = 0.0000005
    else:
        calc.dError = 0.0005

    return bond, calc, periods


# ============================================================
# Payment schedule generation (sRunPayments and helpers)
# ============================================================

def _get_future(dt_from, dt_to, s_curve, calc, bond, curves):
    """Equivalent to VBA fGetFuture — computes forward factor between two dates."""
    import datetime as _dt
    dt_from = _to_date(dt_from)
    dt_to = _to_date(dt_to)

    if dt_to <= calc.dtDay:
        return 1.0

    d_am = 1.0
    dt_wd = brworkday(dt_to - _dt.timedelta(days=1), 1)

    if dt_to > calc.dtDay and calc.dtDay > dt_from:
        i_days = brworkdays(calc.dtDay, dt_to) - 1
    else:
        i_days = brworkdays(dt_from, dt_to) - 1

    if s_curve not in ("CDI", "DI1", "Pré"):
        if '+' in s_curve:
            s_am_key = lag_month_key(dt_from, calc.iAMlag)
            dt_am1 = dt_from
            dt_am2 = next_month(dt_from)

            while _test_am(s_am_key, bond, curves) and dt_am2 <= dt_to:
                if calc.dtDay > dt_am1 and calc.dtDay < dt_am2:
                    num = brworkdays(calc.dtDay, dt_am2) - 1
                    den = brworkdays(dt_am1, dt_am2) - 1
                    d_am *= _am_calc(dt_am1, dt_am2, bond, calc, curves) ** (num / den if den else 0)
                elif calc.dtDay < dt_am1:
                    d_am *= _am_calc(dt_am1, dt_am2, bond, calc, curves)

                dt_am1 = dt_am2
                dt_am2 = next_month(dt_am2)
                s_am_key = lag_month_key(dt_am2, calc.iAMlag)

            if d_am == 1.0:
                i_days = brworkdays(dt_from, dt_to) - 1
            else:
                i_days = brworkdays(dt_am1, dt_to) - 1

            if calc.iAMlag != 0:
                dt_wd = lag_date(dt_to, calc.iAMlag)
                dt_wd = brworkday(dt_wd - _dt.timedelta(days=1), 1)

    result = 1.0
    if s_curve in ("IPCA +", "IPCA"):
        if i_days <= 0:
            result = d_am
        else:
            dt_key = str(dt_wd)
            rate = curves.dicIPCA.get(dt_key, list(curves.dicIPCA.values())[-1] if curves.dicIPCA else 0)
            result = d_am * (1 + rate) ** (i_days / 252)
    elif s_curve in ("IGPM +", "IGPM"):
        if i_days <= 0:
            result = d_am
        else:
            dt_key = str(dt_wd)
            rate = curves.dicIGPM.get(dt_key, list(curves.dicIGPM.values())[-1] if curves.dicIGPM else 0)
            result = d_am * (1 + rate) ** (i_days / 252)
    elif s_curve in ("IGPDI +", "IGPDI"):
        if i_days <= 0:
            result = d_am
        else:
            dt_key = str(dt_wd)
            rate = curves.dicIGPDI.get(dt_key, list(curves.dicIGPDI.values())[-1] if curves.dicIGPDI else 0)
            result = d_am * (1 + rate) ** (i_days / 252)
    elif s_curve in ("DI1", "CDI"):
        dt_key = str(dt_wd)
        rate = curves.dicDI1.get(dt_key, list(curves.dicDI1.values())[-1] if curves.dicDI1 else 0)
        result = (1 + rate) ** (i_days / 252)
    elif s_curve == "Pré":
        dt_key = str(dt_wd)
        rate = curves.dicDI1.get(dt_key, list(curves.dicDI1.values())[-1] if curves.dicDI1 else 0)
        result = (1 + rate) ** (i_days / 252)

    if calc.sAMmonth and s_curve not in ("CDI", "DI1") and '+' in s_curve:
        if am_month_check(calc.sAMmonth, dt_to) != 0:
            if calc.dAccInflFactor != 0:
                result = calc.dAccInflFactor * result
            calc.dAccInflFactor = 1.0
        else:
            if calc.dAccInflFactor == 0:
                calc.dAccInflFactor = result
            else:
                calc.dAccInflFactor *= result
            result = 1.0

    return result


def _test_am(key, bond, curves):
    """Equivalent to VBA fTestAM."""
    if bond.sIndex == "IPCA +":
        return key in curves.dicAMipca
    elif bond.sIndex == "IGPM +":
        return key in curves.dicAMigpm
    elif bond.sIndex == "IGPDI +":
        return key in curves.dicAMigpdi
    return False


def _am_calc(dt1, dt2, bond, calc, curves):
    """Equivalent to VBA fAMcalc."""
    key_from = lag_month_key(dt1, calc.iAMlag)
    key_to = lag_month_key(dt2, calc.iAMlag)

    if _test_am(key_from, bond, curves) and _test_am(key_to, bond, curves):
        if bond.sIndex == "IPCA +":
            d_am = curves.dicAMipca[key_to] / curves.dicAMipca[key_from]
        elif bond.sIndex == "IGPM +":
            d_am = curves.dicAMigpm[key_to] / curves.dicAMigpm[key_from]
        elif bond.sIndex == "IGPDI +":
            d_am = curves.dicAMigpdi[key_to] / curves.dicAMigpdi[key_from]
        else:
            d_am = 1.0
    else:
        d_am = 1.0

    return d_am if d_am != 0 else 1.0


def _compound_cdi(dt_from, dt_to, d_factor, calc, curves):
    """Equivalent to VBA fCompoundCDI — accumulated CDI between two dates."""
    dt_from = _to_date(dt_from)
    dt_to = _to_date(dt_to)
    i_days = brworkdays(dt_from, dt_to) - 1
    dt_day = brworkday(dt_from, -calc.iCDIlag)
    d_compound = 1.0

    for _ in range(i_days):
        dt_key = str(dt_day)
        cdi_val = curves.dicCDI.get(dt_key)
        if cdi_val is None or cdi_val == 0:
            di1_val = curves.dicDI1.get(dt_key, 0)
            d_compound *= (((1 + di1_val) ** (1 / 252) - 1) * d_factor + 1)
        else:
            d_compound *= (((1 + cdi_val) ** (1 / 252) - 1) * d_factor + 1)
        dt_day = brworkday(dt_day, 1)

    return d_compound


def _get_cdi(dt_from, dt_to, bond, calc, curves):
    """Equivalent to VBA fGetCDI."""
    dt_from = _to_date(dt_from)
    dt_to = _to_date(dt_to)

    if bond.sIndex == "% CDI":
        d_f = bond.dYield
    else:
        d_f = 1.0

    result = 1.0
    if dt_to <= calc.dtDay:
        result = _compound_cdi(dt_from, dt_to, d_f, calc, curves)
    else:
        if dt_from < calc.dtDay:
            result = _compound_cdi(dt_from, calc.dtDay, d_f, calc, curves)

    return result


def _get_am_principal(dt_from, dt_to, i_run, bond, calc, periods, period_bonds, curves):
    """Equivalent to VBA fGetAMprincipal — computes monetary adjustment factor."""
    dt_from = _to_date(dt_from)
    dt_to = _to_date(dt_to)

    if ("CDI" in bond.sIndex or bond.sIndex == "Pré" or
            (dt_from > calc.dtDay and not calc.sAMmonth and not calc.sAMcarac)):
        return 1.0

    if calc.bAmort100 and periods[i_run].dIncorpYield != 0:
        return 1.0

    saved_am_lag = calc.iAMlag
    if calc.sAMcarac == "Não Atualiza com Prévia" and calc.dtDay < dt_to and calc.dtDay > dt_from:
        calc.iAMlag += 1

    dt_n1 = dt_from
    dt_n2 = dt_to
    i_months = 0

    if calc.sAMmonth:
        i_amm = am_month_check(calc.sAMmonth, dt_to)
        if i_amm != 0:
            dt_n2_t = dt_to
            try:
                dt_n1_t = date(dt_to.year - 1, i_amm, dt_to.day)
            except ValueError:
                dt_n1_t = date(dt_to.year - 1, i_amm, 28)

            if dt_n1_t < periods[0].dtDay:
                if calc.sYdays != "Úteis":
                    try:
                        dt_n1_t = date(dt_to.year - 1, periods[0].dtDay.month, dt_to.day)
                    except ValueError:
                        dt_n1_t = date(dt_to.year - 1, periods[0].dtDay.month, 28)
                    if n_months(dt_n1_t, dt_n2_t) > 12:
                        try:
                            dt_n1_t = date(dt_to.year, periods[0].dtDay.month, dt_to.day)
                        except ValueError:
                            dt_n1_t = date(dt_to.year, periods[0].dtDay.month, 28)

            i_months = n_months(dt_n1_t, dt_n2_t)

            if dt_n2_t > calc.dtDay:
                if dt_n1_t > calc.dtDay:
                    calc.iAMlag = saved_am_lag
                    return 1.0
                else:
                    dt_n2_t = calc.dtDay

            dt_n1 = dt_n1_t
            dt_n2 = dt_n2_t
        else:
            calc.iAMlag = saved_am_lag
            return 1.0
    else:
        if calc.bAmort100 and i_run > 0 and periods[i_run - 1].dIncorpYield != 0:
            i_corr = 1
            while (i_run - i_corr >= 0 and
                   periods[i_run - i_corr].dIncorpYield == 1 and
                   i_corr < i_run):
                i_corr += 1
            dt_n1 = periods[i_run - i_corr].dtDay

        i_months = n_months(dt_n1, dt_n2)

    if i_months == 0:
        day_to_offset = dt_to.day - 1
        dt_n1 = dt_from - timedelta(days=day_to_offset)
        dt_n2 = dt_to - timedelta(days=day_to_offset)
        i_months = 1
    elif i_months > 1:
        try:
            dt_n2 = next_month(date(dt_n1.year, dt_n1.month, dt_to.day))
        except ValueError:
            dt_n2 = next_month(dt_n1)

    d_am_acc = 1.0
    for _ in range(i_months):
        if dt_n2 <= calc.dtDay:
            if calc.sYdays == "21/252":
                d_factor = 1.0
            else:
                days_dc = fDays(dt_n2, "DC")
                d_factor = (dt_n2 - dt_n1).days / days_dc if days_dc else 1.0
        elif dt_n1 < calc.dtDay:
            d_factor = fFactorDays(calc.dtDay, dt_n1, dt_n2, calc.sYdays)
        else:
            break

        d_am_acc *= _am_calc(dt_n1, dt_n2, bond, calc, curves) ** d_factor
        dt_n1 = dt_n2
        dt_n2 = next_month(dt_n2)

    calc.iAMlag = saved_am_lag
    return d_am_acc


def run_payments(bond, calc, periods, curves):
    """
    Generates payment schedule. Equivalent to VBA sRunPayments.
    Returns: List[PeriodBond]
    """
    from datetime import timedelta
    n = len(periods)
    pbs = [PeriodBond() for _ in range(n)]

    pbs[0].dSN = bond.dPU
    pbs[0].dSNA = bond.dPU
    pbs[0].dFatAmAcc = 1.0

    for i in range(1, n):
        if bond.dicYields:
            dt_key = str(periods[i - 1].dtDay)
            if dt_key in bond.dicYields:
                bond.dYield = bond.dicYields[dt_key]

        _calc_yield(i, bond, calc, periods, pbs, curves)
        _calc_am(i, bond, calc, periods, pbs, curves)
        _calc_pmt(i, bond, calc, periods, pbs)
        _calc_pv_par(i, bond, calc, periods, pbs)

    return pbs


def _calc_yield(i, bond, calc, periods, pbs, curves):
    """Equivalent to VBA sYield — computes interest factors for the period."""
    if "CDI" not in bond.sIndex and bond.sIndex != "Pré":
        pbs[i].dYinf = _get_future(periods[i - 1].dtDay, periods[i].dtDay, bond.sIndex, calc, bond, curves)
    else:
        pbs[i].dYinf = 1.0

    pbs[i].dYAuxCurveInp = _get_future(periods[i - 1].dtDay, periods[i].dtDay, bond.sSpreadIndexInp, calc, bond, curves)
    pbs[i].dYAuxCurveRes = _get_future(periods[i - 1].dtDay, periods[i].dtDay, bond.sSpreadIndexRes, calc, bond, curves)
    pbs[i].dYdi1 = _get_future(periods[i - 1].dtDay, periods[i].dtDay, "DI1", calc, bond, curves)
    pbs[i].dYcdi = _get_cdi(periods[i - 1].dtDay, periods[i].dtDay, bond, calc, curves)

    if bond.sIndex != "% CDI":
        i_days, i_month_days = _get_period_days(i, calc, periods)

        if calc.sYdays == "Úteis":
            pbs[i].dYspread = (1 + bond.dYield) ** (i_days / 252)
        elif calc.sYdays == "30/360":
            pbs[i].dYspread = ((1 + bond.dYield) ** (30 / 360)) ** (i_days / i_month_days) if i_month_days else 1.0
        elif calc.sYdays == "1/360":
            pbs[i].dYspread = ((1 + bond.dYield) ** (1 / 360)) ** i_days
        elif calc.sYdays == "1/365":
            pbs[i].dYspread = ((1 + bond.dYield) ** (1 / 365)) ** i_days
        elif calc.sYdays == "21/252":
            pbs[i].dYspread = ((1 + bond.dYield) ** (21 / 252)) ** (i_days / i_month_days) if i_month_days else 1.0
    else:
        i_days = brworkdays(periods[i - 1].dtDay, periods[i].dtDay) - 1
        if calc.dtDay > periods[i - 1].dtDay:
            i_d = brworkdays(calc.dtDay, periods[i].dtDay) - 1
            if i_d == 0:
                i_d = 1
        else:
            i_d = i_days

        pbs[i].dYspread = pbs[i].dYcdi * ((bond.dYield) * ((pbs[i].dYdi1) ** (1 / i_d) - 1) + 1) ** (i_d / 1)

    if "CDI +" in bond.sIndex:
        pbs[i].dYtotal = pbs[i].dYcdi * pbs[i].dYdi1 * pbs[i].dYspread
    else:
        pbs[i].dYtotal = pbs[i].dYspread

    _incorporated_yield(i, bond, calc, periods, pbs)


def _get_period_days(i, calc, periods):
    """Calculates period days according to day-count convention."""
    if calc.sYdays == "Úteis":
        i_days = brworkdays(periods[i - 1].dtDay, periods[i].dtDay) - 1
        return i_days, 0
    elif calc.sYdays in ("30/360", "1/360", "1/365"):
        i_days = (periods[i].dtDay - periods[i - 1].dtDay).days
        i_month_days = fDays(periods[i].dtDay, "DC")
        return i_days, i_month_days
    elif calc.sYdays == "21/252":
        i_days = brworkdays(periods[i - 1].dtDay, periods[i].dtDay) - 1
        i_month_days = fDays(periods[i].dtDay, "DU")
        return i_days, i_month_days
    return 0, 0


def _incorporated_yield(i, bond, calc, periods, pbs):
    """Equivalent to VBA sIncorporatedYield."""
    if not calc.bAmort100:
        return

    if periods[i].dIncorpYield != 0:
        if periods[i].dIncorpYield > 1:
            pbs[i].dPMTIncorpJuros = 0
            denom = pbs[i - 1].dSN * (pbs[i].dYtotal - 1)
            if denom != 0:
                periods[i].dIncorpYield = periods[i].dIncorpYield / denom

        pbs[i].dYinf = (pbs[i].dYinf - 1) * periods[i].dIncorpYield + 1
        pbs[i].dYcdi = (pbs[i].dYcdi - 1) * periods[i].dIncorpYield + 1
        pbs[i].dYdi1 = (pbs[i].dYdi1 - 1) * periods[i].dIncorpYield + 1
        pbs[i].dYspread = (pbs[i].dYspread - 1) * periods[i].dIncorpYield + 1
        pbs[i].dYtotal = (pbs[i].dYtotal - 1) * periods[i].dIncorpYield + 1

    if i > 0 and periods[i - 1].dIncorpYield != 0:
        pbs[i].dYinf *= pbs[i - 1].dYinf
        pbs[i].dYcdi *= pbs[i - 1].dYcdi
        pbs[i].dYdi1 *= pbs[i - 1].dYdi1
        pbs[i].dYspread *= pbs[i - 1].dYspread
        pbs[i].dYtotal *= pbs[i - 1].dYtotal


def _calc_am(i, bond, calc, periods, pbs, curves):
    """Equivalent to VBA sAM — monetary adjustment."""
    pbs[i].dFatAm = _get_am_principal(
        periods[i - 1].dtDay, periods[i].dtDay, i, bond, calc, periods, pbs, curves
    )

    if pbs[i].dFatAm < 1 and calc.sAMcarac == "Só AM Positiva":
        pbs[i].dFatAm = 1.0

    pbs[i].dFatAmAcc = pbs[i - 1].dFatAmAcc * pbs[i].dFatAm * pbs[i].dYinf
    pbs[i].dSNA = pbs[i - 1].dSN * pbs[i].dFatAm
    pbs[i].dSNA = pbs[i].dSNA * pbs[i].dYinf


def _calc_pmt(i, bond, calc, periods, pbs):
    """Equivalent to VBA sPMT — payment calculation."""
    if not calc.bAmort100:
        if periods[i].dIncorpYield != 0:
            if periods[i].dIncorpYield <= 1:
                pbs[i].dPMTIncorpJuros = periods[i].dIncorpYield * pbs[i].dSNA * (pbs[i].dYtotal - 1)
            else:
                pbs[i].dPMTIncorpJuros = periods[i].dIncorpYield

        pbs[i].dPMTJuros = round(pbs[i].dSNA * (pbs[i].dYtotal - 1) - pbs[i].dPMTIncorpJuros, 10)

        if calc.bAMTotal:
            pbs[i].dPMTAmort = (pbs[i].dSNA * periods[i].dAmort) / (pbs[i].dFatAmAcc / pbs[i - 1].dFatAmAcc)
            pbs[i].dPMTAMprincipal = pbs[i].dSNA - pbs[i - 1].dSN
        else:
            fat = pbs[i].dFatAmAcc if pbs[i].dFatAmAcc != 0 else 1
            pbs[i].dPMTAmort = (pbs[i].dSNA * periods[i].dAmort) / fat
            pbs[i].dPMTAMprincipal = (periods[i].dAmort * pbs[i].dSNA) * (1 - (1 / fat))

    else:
        if periods[i].dIncorpYield != 0:
            pbs[i].dPMTJuros = round(
                pbs[i].dSNA * ((pbs[i].dYtotal - 1) / periods[i].dIncorpYield) * (1 - periods[i].dIncorpYield), 10)
        else:
            pbs[i].dPMTJuros = round(pbs[i].dSNA * (pbs[i].dYtotal - 1), 10)

        pbs[i].dPMTAmort = bond.dPU * periods[i].dAmort

        if calc.bAMTotal:
            pbs[i].dPMTAMprincipal = pbs[i].dSNA - pbs[i - 1].dSN
        else:
            pbs[i].dPMTAMprincipal = bond.dPU * periods[i].dAmort * (pbs[i].dFatAmAcc - 1)

    if calc.dtDay >= periods[i].dtDay:
        if periods[i].dExtrAmort <= 1:
            pbs[i].dPMTAmortExtr = periods[i].dExtrAmort * pbs[i].dSNA
        else:
            pbs[i].dPMTAmortExtr = periods[i].dExtrAmort

        if not calc.bAMTotal:
            fat = pbs[i].dFatAmAcc if pbs[i].dFatAmAcc != 0 else 1
            pbs[i].dPMTAMprincipal += periods[i].dExtrAmort * (1 - (1 / fat))
            pbs[i].dPMTAmortExtr = pbs[i].dPMTAmortExtr / fat

        if periods[i].dMultaFee <= 1:
            pbs[i].dPMTMulta = periods[i].dMultaFee * pbs[i].dSNA
        else:
            pbs[i].dPMTMulta = periods[i].dMultaFee

    if pbs[i].dPMTAMprincipal < 0 and pbs[i].dPMTAmort == 0:
        pbs[i].dPMTAMprincipal = 0

    pbs[i].dPMTTotal = (pbs[i].dPMTJuros + pbs[i].dPMTAmort + pbs[i].dPMTAmortExtr +
                         pbs[i].dPMTAMprincipal + pbs[i].dPMTMulta)

    pbs[i].dSN = (pbs[i].dSNA - pbs[i].dPMTAmort - pbs[i].dPMTAmortExtr -
                   pbs[i].dPMTAMprincipal + pbs[i].dPMTIncorpJuros)


def _calc_pv_par(i, bond, calc, periods, pbs):
    """Equivalent to VBA sPvPar — present value for par calculation."""
    from .pv import fGetSpread, fGetSpreadPerc

    d_pv_yield = bond.dYield

    if calc.dtDay > periods[i].dtDay:
        pbs[i].dPVfactorPar = pbs[i - 1].dPVfactorPar

    elif calc.dtDay == periods[i].dtDay:
        pbs[i].dPVfactorPar = pbs[i - 1].dPVfactorPar
        if calc.bPmtIncorp:
            pbs[i].dPVpmtPar = pbs[i].dPMTTotal / pbs[i].dPVfactorPar if pbs[i].dPVfactorPar else 0
    else:
        if '+' in bond.sIndex or bond.sIndex == "Pré":
            pbs[i].dPVfactorPar = pbs[i - 1].dPVfactorPar * fGetSpread(i, d_pv_yield, calc, periods)
        elif '%' in bond.sIndex:
            pbs[i].dPVfactorPar = pbs[i - 1].dPVfactorPar * fGetSpreadPerc(i, d_pv_yield, calc, periods, pbs)

        if bond.sIndex == "CDI +":
            pbs[i].dPVfactorPar *= pbs[i].dYdi1
        elif '%' not in bond.sIndex:
            pbs[i].dPVfactorPar *= pbs[i].dYinf

        pbs[i].dPVpmtPar = pbs[i].dPMTTotal / pbs[i].dPVfactorPar if pbs[i].dPVfactorPar else 0
