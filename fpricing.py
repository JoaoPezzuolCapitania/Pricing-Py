"""
fpricing.py — Função principal que substitui fPricing do VBA.
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
    """Configura tipo de cálculo. Replica sTypeCalc."""
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
    """Configura tipo de resultado. Replica sTypeResult."""
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
    """Retorna resultado conforme tipo. Replica fResult."""
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


# State global de tCalc.dAccInflFactor — replica o comportamento do VBA, onde
# tCalc é módulo-level e dAccInflFactor persiste entre chamadas de fPricing.
# Necessário pra bonds com sAMmonth: workflow batch (PROD) ativa essa pollution.
_PERSISTENT_STATE = {'dAccInflFactor': 0.0}


def reset_persistent_state():
    """Reseta o estado global. Use no início de cada workflow novo."""
    _PERSISTENT_STATE['dAccInflFactor'] = 0.0


def fPricing(s_bond, dt_day, d_value, i_info, i_result, curves=None):
    """
    Réplica exata do fPricing do VBA.

    Args:
        s_bond: código CETIP do ativo (ex: 'CBAN72')
        dt_day: data de cálculo (str 'YYYY-MM-DD' ou date)
        d_value: valor de input (PU, taxa, spread, etc)
        i_info: tipo de input (1=Taxa, 2=PU, 3=%PUPar, 4-7=Spread)
        i_result: tipo de output (1=Taxa, 2=PU, 3=%PUPar, 4=Spread, 8=Duration, 9=OverTP)
        curves: CurveData pré-carregadas (opcional, carrega automaticamente se None)

    Returns:
        float ou str de erro
    """
    from .daycount import _to_date
    dt_day = _to_date(dt_day)

    # Carregar curvas (se não fornecidas)
    if curves is None:
        curves = load_curves(dt_day)

    # Carregar bond
    result = load_bond(s_bond, dt_day)
    if result is None:
        return "Ativo não cadastrado!"

    bond, calc, periods = result

    # Restaurar dAccInflFactor da chamada anterior (replica state global do VBA)
    calc.dAccInflFactor = _PERSISTENT_STATE['dAccInflFactor']

    # Configurar cálculo
    _setup_calc(calc, i_info, d_value, bond)
    _setup_result(calc, i_result, bond)

    # Gerar pagamentos
    pbs = run_payments(bond, calc, periods, curves)

    # Calcular par
    results = BondResults()
    results.dPar = get_par(bond, pbs)

    # Cálculo principal (baseado no input)
    if calc.sType == "PU":
        get_taxa(calc, bond, periods, pbs, results)
    elif calc.sType == "Yield":
        get_pu(calc, bond, periods, pbs, results)
    elif calc.sType == "%PUPar":
        get_perc_pu_par(calc, results)
    elif calc.sType == "Taxa Spread":
        get_spread(calc, bond, periods, pbs, results)

    # Resultado adicional (baseado no output)
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

    # Persistir dAccInflFactor pra proxima chamada (replica state global do VBA)
    _PERSISTENT_STATE['dAccInflFactor'] = calc.dAccInflFactor

    return _get_result(i_result, results)


def fPricing_batch(df, dt, inp=2, curves=None):
    """
    Versão batch para múltiplos ativos.
    Carrega curvas UMA vez e processa todos os ativos.

    Args:
        df: DataFrame com colunas 'cod_cetip' e 'pu'
        dt: data de cálculo (str 'YYYY-MM-DD')
        inp: tipo de input (default 2 = PU)
        curves: CurveData pré-carregadas (opcional)

    Returns:
        DataFrame com colunas adicionais: taxa, spread, duration, over_tp
    """
    from .daycount import _to_date
    dt = _to_date(dt)

    if curves is None:
        print("Carregando curvas...")
        curves = load_curves(dt)
        print("Curvas carregadas!")

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
