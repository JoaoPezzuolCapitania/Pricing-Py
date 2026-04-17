"""
diag_duration.py — Diagnóstico detalhado de duration CDI+ outliers.
Compara per-period dPVpmtCalc entre Python e VBA COM.
"""

import sys
import os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from datetime import date
from pricing_engine.curves import load_curves
from pricing_engine.bond import load_bond, run_payments, BondResults
from pricing_engine.solver import get_par, get_taxa, duration
from pricing_engine.pv import pv_calc, fGetSpread, get_price
from pricing_engine.daycount import brworkdays


def dump_period_data(cetip, dt_str='2026-03-26'):
    """Dump complete per-period data for a bond."""
    dt = date(2026, 3, 26) if dt_str == '2026-03-26' else None

    print(f"\n{'='*80}")
    print(f"  DIAGNÓSTICO: {cetip} — data {dt_str}")
    print(f"{'='*80}\n")

    # Load
    curves = load_curves(dt)
    result = load_bond(cetip, dt)
    if result is None:
        print("ERRO: Ativo não encontrado!")
        return

    bond, calc, periods = result
    print(f"  Bond: {bond.sCETIP}")
    print(f"  Type: {bond.sType}")
    print(f"  Index: {bond.sIndex}")
    print(f"  Yield: {bond.dYield}")
    print(f"  PU: {bond.dPU}")
    print(f"  Periods: {bond.iPeriods}")
    print(f"  bAmort100: {calc.bAmort100}")
    print(f"  sYdays: {calc.sYdays}")
    print(f"  iPmtlag: {calc.iPmtlag}")
    print(f"  iCDIlag: {calc.iCDIlag}")
    print()

    # Run payments
    pbs = run_payments(bond, calc, periods, curves)

    # Print period schedule
    print(f"{'i':>3} {'dtDay':>12} {'IncYld':>7} {'Amort':>10} {'ExtrAm':>12} {'dYcdi':>14} {'dYdi1':>14} {'dYsprd':>14} {'dYtotal':>14} {'dPMTTotal':>14} {'dSN':>14}")
    print("-" * 170)
    for i, (pi, pb) in enumerate(zip(periods, pbs)):
        print(f"{i:3d} {str(pi.dtDay):>12} {pi.dIncorpYield:7.1f} {pi.dAmort:10.6f} {pi.dExtrAmort:12.10f} "
              f"{pb.dYcdi:14.10f} {pb.dYdi1:14.10f} {pb.dYspread:14.10f} {pb.dYtotal:14.10f} "
              f"{pb.dPMTTotal:14.6f} {pb.dSN:14.6f}")

    print()

    # Get par
    results = BondResults()
    results.dPar = get_par(bond, pbs)
    print(f"  Par: {results.dPar}")

    # Setup for PU -> Taxa
    calc.sType = "PU"
    calc.sResult = "Duration"

    # Read PU from comparison (hardcode for known outliers)
    # PUs corretos do Comparativo Excel 20260326
    pus = {
        'LCAMC2': 1079.6402,
        '22L1212138': 920.64758,
        '5483424UN1': 762.39,
        '6039625SR1': 698.9885,
        '6083125SR1': 1003.0545,
        '5241224SR3': 676.860773,
    }
    pu = pus.get(cetip, 1000.0)
    calc.dPU = pu
    print(f"  PU input: {pu}")

    # Run Newton-Raphson
    get_taxa(calc, bond, periods, pbs, results)
    print(f"  Yield (NR): {results.dYield}")
    print(f"  Price (NR): {results.dPrice}")

    # Duration
    duration(calc, bond, periods, pbs, results)
    print(f"  Duration PY: {results.dDuration}")
    print()

    # Print PV per period
    print(f"{'i':>3} {'dtDay':>12} {'days_to':>7} {'dPVfactCalc':>18} {'dPVpmtCalc':>18} {'dPMTTotal':>14} {'contrib_dur':>14}")
    print("-" * 100)

    d_sum_dur = 0.0
    for i, (pi, pb) in enumerate(zip(periods, pbs)):
        days_to = brworkdays(calc.dtDay, pi.dtDay) - 1
        contrib = pb.dPVpmtCalc * days_to
        d_sum_dur += contrib
        if pb.dPVpmtCalc != 0:
            print(f"{i:3d} {str(pi.dtDay):>12} {days_to:7d} {pb.dPVfactorCalc:18.12f} {pb.dPVpmtCalc:18.10f} "
                  f"{pb.dPMTTotal:14.6f} {contrib:14.6f}")

    print()
    print(f"  Sum(dPVpmtCalc * days) = {d_sum_dur:.10f}")
    print(f"  Duration = {d_sum_dur} / {results.dPrice} / 252 = {d_sum_dur / results.dPrice / 252:.6f}")
    print()

    return bond, calc, periods, pbs, results


def compare_nr_final_state(cetip, dt_str='2026-03-26'):
    """
    Compare what dPVpmtCalc looks like with:
    1. converged yield (Python behavior)
    2. converged yield + 1e-10 (VBA behavior — derivative step overwrites)
    """
    dt = date(2026, 3, 26)
    curves = load_curves(dt)
    result = load_bond(cetip, dt)
    if result is None:
        return

    bond, calc, periods = result
    pbs = run_payments(bond, calc, periods, curves)
    results = BondResults()
    results.dPar = get_par(bond, pbs)

    # PUs corretos do Comparativo Excel 20260326
    pus = {
        'LCAMC2': 1079.6402,
        '22L1212138': 920.64758,
        '5483424UN1': 762.39,
        '6039625SR1': 698.9885,
        '6083125SR1': 1003.0545,
        '5241224SR3': 676.860773,
    }
    calc.dPU = pus.get(cetip, 1000.0)
    calc.sType = "PU"

    # Run NR to get converged yield
    get_taxa(calc, bond, periods, pbs, results)
    yield_converged = results.dYield

    # Now compute dPVpmtCalc with exact converged yield
    price_exact = get_price(yield_converged, "Taxa", bond, calc, periods, pbs)
    dur_data_exact = []
    for i in range(len(periods)):
        days_to = brworkdays(calc.dtDay, periods[i].dtDay) - 1
        dur_data_exact.append((i, pbs[i].dPVpmtCalc, days_to))

    dur_exact = sum(d[1] * d[2] for d in dur_data_exact) / price_exact / 252

    # Now compute with yield + 1e-10 (simulating VBA behavior)
    price_delta = get_price(yield_converged + 1e-10, "Taxa", bond, calc, periods, pbs)
    dur_data_delta = []
    for i in range(len(periods)):
        days_to = brworkdays(calc.dtDay, periods[i].dtDay) - 1
        dur_data_delta.append((i, pbs[i].dPVpmtCalc, days_to))

    dur_delta = sum(d[1] * d[2] for d in dur_data_delta) / price_exact / 252

    print(f"\n  LCAMC2 NR final state comparison:")
    print(f"  Yield converged: {yield_converged}")
    print(f"  Duration (yield exact):    {dur_exact:.6f}")
    print(f"  Duration (yield + 1e-10):  {dur_delta:.6f}")
    print(f"  Difference:                {dur_delta - dur_exact:.12f}")


def compare_with_vba_com(cetip, dt_str='2026-03-26'):
    """
    Use COM to get VBA duration and compare.
    Requires Excel with PricingRFE.xlam loaded.
    """
    try:
        import pythoncom
        import win32com.client
        import tempfile, uuid, openpyxl
    except ImportError:
        print("win32com not available — skipping COM comparison")
        return

    dt = date(2026, 3, 26)

    # PUs corretos do Comparativo Excel 20260326
    pus = {
        'LCAMC2': 1079.6402,
        '22L1212138': 920.64758,
        '5483424UN1': 762.39,
        '6039625SR1': 698.9885,
        '6083125SR1': 1003.0545,
        '5241224SR3': 676.860773,
    }
    pu = pus.get(cetip, 1000.0)

    print(f"\n  COM comparison for {cetip}...")

    pythoncom.CoInitialize()
    xl = win32com.client.DispatchEx("Excel.Application")
    xl.Visible = False

    try:
        xl.Workbooks.Open(r'C:\Add-in\Oficial\PricingRFE.xlam')

        wb_temp = openpyxl.Workbook()
        tf = tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx')
        wb_temp.save(tf.name)
        wb = xl.Workbooks.Open(tf.name)
        ws = wb.ActiveSheet

        # Get taxa
        ws.Range('A1').Value = cetip
        ws.Range('A2').Value = dt_str
        ws.Range('A3').Value = pu
        ws.Range('A4').Formula = f'=fPricing(A1, A2, A3, 2, 1)'
        taxa_vba = ws.Range('A4').Value

        # Get duration
        ws.Range('A5').Formula = f'=fPricing(A1, A2, A3, 2, 8)'
        dur_vba = ws.Range('A5').Value

        # Get spread
        ws.Range('A6').Formula = f'=fPricing(A1, A2, A3, 2, 4)'
        spread_vba = ws.Range('A6').Value

        print(f"  VBA Taxa:     {taxa_vba}")
        print(f"  VBA Duration: {dur_vba}")
        print(f"  VBA Spread:   {spread_vba}")

        wb.Close(False)
    finally:
        xl.Quit()
        pythoncom.CoUninitialize()
        try:
            os.remove(tf.name)
        except:
            pass


def test_vba_yield_in_python(cetip, vba_yield, dt_str='2026-03-26'):
    """
    Injeta a yield EXATA do VBA no Python e compara dPVpmtCalc / duration.
    Se duration ainda difere, o problema está na distribuição de pagamentos (dPMTTotal),
    não no NR convergence.
    """
    dt = date(2026, 3, 26)
    curves = load_curves(dt)
    result = load_bond(cetip, dt)
    if result is None:
        return

    bond, calc, periods = result
    pbs = run_payments(bond, calc, periods, curves)
    results = BondResults()
    results.dPar = get_par(bond, pbs)

    # PUs corretos
    pus = {
        'LCAMC2': 1079.6402,
        '22L1212138': 920.64758,
        '5483424UN1': 762.39,
        '6039625SR1': 698.9885,
        '6083125SR1': 1003.0545,
        '5241224SR3': 676.860773,
    }
    calc.dPU = pus.get(cetip, 1000.0)
    calc.sType = "PU"

    # 1) Python NR yield
    get_taxa(calc, bond, periods, pbs, results)
    py_yield = results.dYield
    py_price = results.dPrice

    # Duration com yield Python
    duration(calc, bond, periods, pbs, results)
    py_dur = results.dDuration

    # Salvar dPVpmtCalc com yield Python
    py_pvpmt = [(i, pbs[i].dPVpmtCalc, pbs[i].dPVfactorCalc, pbs[i].dPMTTotal)
                for i in range(len(periods))]

    # 2) Agora usar yield VBA
    price_vba_yield = get_price(vba_yield, "Taxa", bond, calc, periods, pbs)
    vba_pvpmt = [(i, pbs[i].dPVpmtCalc, pbs[i].dPVfactorCalc, pbs[i].dPMTTotal)
                 for i in range(len(periods))]

    # Duration com yield VBA
    d_sum_dur_vba = 0.0
    for i in range(len(periods)):
        days_to = brworkdays(calc.dtDay, periods[i].dtDay) - 1
        d_sum_dur_vba += pbs[i].dPVpmtCalc * days_to
    dur_vba_yield = d_sum_dur_vba / price_vba_yield / 252 if price_vba_yield else 0

    print(f"\n{'='*80}")
    print(f"  TEST VBA YIELD IN PYTHON: {cetip}")
    print(f"{'='*80}")
    print(f"  PU: {calc.dPU}")
    print(f"  Python yield: {py_yield} -> price={py_price} -> duration={py_dur}")
    print(f"  VBA yield:    {vba_yield} -> price={price_vba_yield:.6f} -> duration={dur_vba_yield:.6f}")
    print()

    # Comparar dPVpmtCalc por período (apenas futuros)
    print(f"{'i':>3} {'dtDay':>12} {'dPVpmtCalc_PY':>18} {'dPVpmtCalc_VBA_Y':>18} {'diff':>14} {'dPVfact_PY':>18} {'dPVfact_VBA_Y':>18} {'fact_diff':>14}")
    print("-" * 140)
    for idx in range(len(periods)):
        py_pv = py_pvpmt[idx][1]
        vba_pv = vba_pvpmt[idx][1]
        py_fact = py_pvpmt[idx][2]
        vba_fact = vba_pvpmt[idx][2]
        if py_pv != 0 or vba_pv != 0:
            print(f"{idx:3d} {str(periods[idx].dtDay):>12} {py_pv:18.10f} {vba_pv:18.10f} {vba_pv-py_pv:14.10f} "
                  f"{py_fact:18.12f} {vba_fact:18.12f} {vba_fact-py_fact:14.12f}")


def deep_compare_5483424UN1():
    """
    Comparação profunda para 5483424UN1: yield bate (diff=3e-6)
    mas duration difere 0.027. Vamos comparar dYdi1 per-período.
    """
    cetip = '5483424UN1'
    dt = date(2026, 3, 26)
    curves = load_curves(dt)
    result = load_bond(cetip, dt)
    if result is None:
        return

    bond, calc, periods = result

    print(f"\n{'='*80}")
    print(f"  DEEP COMPARE: {cetip} — {bond.sIndex}, {bond.iPeriods} periods")
    print(f"  bAmort100={calc.bAmort100}, sYdays={calc.sYdays}, iCDIlag={calc.iCDIlag}")
    print(f"{'='*80}")

    # Run payments
    pbs = run_payments(bond, calc, periods, curves)

    # Mostrar apenas períodos futuros com detalhe de dYdi1, dYcdi
    print(f"\n{'i':>3} {'dtDay':>12} {'dtDayPMT':>12} {'dYcdi':>14} {'dYdi1':>14} {'dYsprd':>14} {'dYtotal':>14} {'dPMTTotal':>14} {'dSN':>14} {'Incorp':>7}")
    print("-" * 155)
    for i, (pi, pb) in enumerate(zip(periods, pbs)):
        if pi.dtDay >= date(2026, 1, 1):  # apenas períodos recentes/futuros
            print(f"{i:3d} {str(pi.dtDay):>12} {str(pi.dtDayPMT):>12} "
                  f"{pb.dYcdi:14.10f} {pb.dYdi1:14.10f} {pb.dYspread:14.10f} {pb.dYtotal:14.10f} "
                  f"{pb.dPMTTotal:14.6f} {pb.dSN:14.6f} {pi.dIncorpYield:7.1f}")

    # Par e taxa
    results = BondResults()
    results.dPar = get_par(bond, pbs)
    calc.dPU = 762.39
    calc.sType = "PU"
    get_taxa(calc, bond, periods, pbs, results)
    duration(calc, bond, periods, pbs, results)
    print(f"\n  Par={results.dPar}, Yield={results.dYield}, Price={results.dPrice}, Duration={results.dDuration}")


if __name__ == '__main__':
    # Bonds onde yield bate mas duration difere (mais informativos)
    # 5483424UN1: taxa diff=3e-6, dur diff=-0.027
    # 6083125SR1: taxa diff=1e-6, dur diff=-0.016
    # 6039625SR1: taxa diff=-1e-6, dur diff=-0.018

    # Taxas VBA do comparativo
    vba_yields = {
        '5483424UN1': 0.048297,
        '6083125SR1': 0.060144,
        '6039625SR1': 0.029704,
        'LCAMC2': 0.016955,
    }

    # Testar yield VBA no Python
    for cetip, vba_y in vba_yields.items():
        test_vba_yield_in_python(cetip, vba_y)

    # Deep compare do melhor candidato
    deep_compare_5483424UN1()
