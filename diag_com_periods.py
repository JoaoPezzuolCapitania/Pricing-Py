"""
diag_com_periods.py - Extrai dPVpmtCalc per-periodo do VBA via COM
e compara com Python.
"""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from datetime import date
import pythoncom
import win32com.client
import tempfile, openpyxl

from pricing_engine.curves import load_curves
from pricing_engine.bond import load_bond, run_payments, BondResults
from pricing_engine.solver import get_par, get_taxa, duration
from pricing_engine.pv import get_price
from pricing_engine.daycount import brworkdays


def extract_vba_periods(xl, ws, addin_wb, cetip, dt_str, pu):
    """
    Roda sPricing no VBA e extrai tPeriodBond per-periodo.
    """
    # Chamar sPricing via COM (sub publica)
    addin_module = None
    for wb in xl.Workbooks:
        if 'PricingRFE' in wb.Name:
            addin_wb = wb
            break

    # Rodar sPricing
    xl.Run("sPricing", cetip, dt_str, pu, 2, 8)

    # Ler dados publicos do modulo
    n_periods = xl.Run("PricingAddIn.tBondInfo.iPeriods")

    vba_data = []
    for i in range(1, n_periods + 1):
        try:
            pvfact = xl.Evaluate(f"PricingAddIn.tPeriodBond({i}).dPVfactorCalc")
            pvpmt = xl.Evaluate(f"PricingAddIn.tPeriodBond({i}).dPVpmtCalc")
            pmttot = xl.Evaluate(f"PricingAddIn.tPeriodBond({i}).dPMTTotal")
            ydi1 = xl.Evaluate(f"PricingAddIn.tPeriodBond({i}).dYdi1")
            ycdi = xl.Evaluate(f"PricingAddIn.tPeriodBond({i}).dYcdi")
            ysprd = xl.Evaluate(f"PricingAddIn.tPeriodBond({i}).dYspread")
            ytot = xl.Evaluate(f"PricingAddIn.tPeriodBond({i}).dYtotal")
            vba_data.append({
                'i': i,
                'dPVfactorCalc': pvfact,
                'dPVpmtCalc': pvpmt,
                'dPMTTotal': pmttot,
                'dYdi1': ydi1,
                'dYcdi': ycdi,
                'dYspread': ysprd,
                'dYtotal': ytot,
            })
        except Exception as e:
            print(f"  Erro periodo {i}: {e}")
            break

    return vba_data


def compare_periods(cetip, dt_str='2026-03-26'):
    """Compara per-periodo Python vs VBA COM usando macro helper injetada."""
    dt = date(2026, 3, 26)

    pus = {
        'LCAMC2': 1079.6402,
        '22L1212138': 920.64758,
        '5483424UN1': 762.39,
        '6039625SR1': 698.9885,
        '6083125SR1': 1003.0545,
        '5241224SR3': 676.860773,
    }
    pu = pus.get(cetip, 1000.0)

    # === PYTHON ===
    curves = load_curves(dt)
    result = load_bond(cetip, dt)
    if result is None:
        print(f"ERRO: {cetip} nao encontrado")
        return

    bond, calc, periods = result
    pbs = run_payments(bond, calc, periods, curves)
    results = BondResults()
    results.dPar = get_par(bond, pbs)
    calc.dPU = pu
    calc.sType = "PU"
    calc.sResult = "Duration"
    get_taxa(calc, bond, periods, pbs, results)
    duration(calc, bond, periods, pbs, results)

    py_yield = results.dYield
    py_dur = results.dDuration

    # === VBA COM com macro helper ===
    print(f"\nAbrindo Excel para {cetip}...")
    pythoncom.CoInitialize()
    xl = win32com.client.DispatchEx("Excel.Application")
    xl.Visible = False

    # Macro VBA helper para extrair dados de tPeriodBond
    helper_code = '''
Sub DumpPeriodData()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    Dim i As Integer
    Dim n As Integer
    n = PricingAddIn.tBondInfo.iPeriods

    ' Header row
    ws.Range("A1").Value = "iPeriods"
    ws.Range("B1").Value = n
    ws.Range("C1").Value = "Yield"
    ws.Range("D1").Value = PricingAddIn.tBondResults.dYield
    ws.Range("E1").Value = "Duration"
    ws.Range("F1").Value = PricingAddIn.tBondResults.dDuration
    ws.Range("G1").Value = "Price"
    ws.Range("H1").Value = PricingAddIn.tBondResults.dPrice

    ' Period data starting row 3
    ws.Range("A2").Value = "i"
    ws.Range("B2").Value = "dPVpmtCalc"
    ws.Range("C2").Value = "dPVfactorCalc"
    ws.Range("D2").Value = "dPMTTotal"
    ws.Range("E2").Value = "dYdi1"
    ws.Range("F2").Value = "dYcdi"
    ws.Range("G2").Value = "dYspread"
    ws.Range("H2").Value = "dYtotal"
    ws.Range("I2").Value = "dSN"

    For i = 1 To n
        ws.Cells(i + 2, 1).Value = i
        ws.Cells(i + 2, 2).Value = PricingAddIn.tPeriodBond(i).dPVpmtCalc
        ws.Cells(i + 2, 3).Value = PricingAddIn.tPeriodBond(i).dPVfactorCalc
        ws.Cells(i + 2, 4).Value = PricingAddIn.tPeriodBond(i).dPMTTotal
        ws.Cells(i + 2, 5).Value = PricingAddIn.tPeriodBond(i).dYdi1
        ws.Cells(i + 2, 6).Value = PricingAddIn.tPeriodBond(i).dYcdi
        ws.Cells(i + 2, 7).Value = PricingAddIn.tPeriodBond(i).dYspread
        ws.Cells(i + 2, 8).Value = PricingAddIn.tPeriodBond(i).dYtotal
        ws.Cells(i + 2, 9).Value = PricingAddIn.tPeriodBond(i).dSN
    Next
End Sub
'''

    try:
        addin = xl.Workbooks.Open(r'C:\Add-in\Oficial\PricingRFE.xlam')

        # Criar workbook .xlsm (com macros)
        wb = xl.Workbooks.Add()
        ws = wb.ActiveSheet

        # Injetar macro helper
        try:
            vbc = wb.VBProject.VBComponents.Add(1)  # vbext_ct_StdModule
            vbc.CodeModule.AddFromString(helper_code)
        except Exception as e:
            print(f"  ERRO ao injetar macro: {e}")
            print("  -> Habilite 'Trust access to the VBA project object model'")
            print("     (File > Options > Trust Center > Macro Settings)")
            wb.Close(False)
            xl.Quit()
            pythoncom.CoUninitialize()
            return

        # Rodar sPricing para popular as variaveis publicas
        xl.Run("sPricing", cetip, dt_str, pu, 2, 8)

        # Rodar macro helper para dumpar dados
        xl.Run("DumpPeriodData")

        # Ler resultados
        n_periods_vba = int(ws.Range("B1").Value)
        vba_yield_val = ws.Range("D1").Value
        vba_dur_val = ws.Range("F1").Value
        vba_price_val = ws.Range("H1").Value

        print(f"\n  VBA yield:    {vba_yield_val}")
        print(f"  VBA duration: {vba_dur_val}")
        print(f"  VBA price:    {vba_price_val}")
        print(f"  VBA periods:  {n_periods_vba}")

        # Ler per-periodo
        vba_pvpmt = []
        vba_pvfact = []
        vba_pmttot = []
        vba_ydi1 = []
        vba_ycdi = []
        vba_ysprd = []
        vba_ytot = []
        vba_sn = []

        for i in range(n_periods_vba):
            row = i + 3  # data starts at row 3
            vba_pvpmt.append(ws.Cells(row, 2).Value or 0)
            vba_pvfact.append(ws.Cells(row, 3).Value or 0)
            vba_pmttot.append(ws.Cells(row, 4).Value or 0)
            vba_ydi1.append(ws.Cells(row, 5).Value or 0)
            vba_ycdi.append(ws.Cells(row, 6).Value or 0)
            vba_ysprd.append(ws.Cells(row, 7).Value or 0)
            vba_ytot.append(ws.Cells(row, 8).Value or 0)
            vba_sn.append(ws.Cells(row, 9).Value or 0)

        wb.Close(False)

    except Exception as e:
        print(f"  ERRO: {e}")
        import traceback
        traceback.print_exc()
        try:
            xl.Quit()
        except:
            pass
        pythoncom.CoUninitialize()
        return

    xl.Quit()
    pythoncom.CoUninitialize()

    # === COMPARACAO ===
    print(f"\n{'='*100}")
    print(f"  COMPARACAO PER-PERIODO: {cetip}")
    print(f"  PY yield={py_yield}, PY dur={py_dur}")
    print(f"  VBA yield={vba_yield_val}, VBA dur={vba_dur_val}")
    print(f"{'='*100}")

    n = min(len(periods), len(vba_pvpmt))

    # dPMTTotal comparison (payments - cause of duration diff)
    print(f"\n--- dPMTTotal (pagamentos) ---")
    print(f"{'i':>3} {'py_pmt':>14} {'vba_pmt':>14} {'diff':>14}")
    print("-" * 50)
    total_pmt_diff = 0
    for i in range(n):
        py_pmt = pbs[i].dPMTTotal
        vba_pmt = vba_pmttot[i]
        diff = vba_pmt - py_pmt
        total_pmt_diff += abs(diff)
        if abs(diff) > 1e-6:
            print(f"{i:3d} {py_pmt:14.6f} {vba_pmt:14.6f} {diff:14.10f}")
    print(f"  Total abs diff PMT: {total_pmt_diff:.10f}")

    # dYcdi comparison (CDI compound - likely root cause)
    print(f"\n--- dYcdi (CDI composto) ---")
    print(f"{'i':>3} {'py_ycdi':>16} {'vba_ycdi':>16} {'diff':>16}")
    print("-" * 60)
    for i in range(n):
        py_cdi = pbs[i].dYcdi
        vba_cdi = vba_ycdi[i]
        diff = vba_cdi - py_cdi
        if abs(diff) > 1e-10:
            print(f"{i:3d} {py_cdi:16.10f} {vba_cdi:16.10f} {diff:16.12f}")

    # dYdi1 comparison
    print(f"\n--- dYdi1 (DI1 forward) ---")
    print(f"{'i':>3} {'py_ydi1':>16} {'vba_ydi1':>16} {'diff':>16}")
    print("-" * 60)
    for i in range(n):
        py_di1 = pbs[i].dYdi1
        vba_di1 = vba_ydi1[i]
        diff = vba_di1 - py_di1
        if abs(diff) > 1e-10:
            print(f"{i:3d} {py_di1:16.10f} {vba_di1:16.10f} {diff:16.12f}")

    # dPVpmtCalc comparison (PV of payments - drives duration)
    print(f"\n--- dPVpmtCalc (PV pagamentos) ---")
    print(f"{'i':>3} {'py_pvpmt':>16} {'vba_pvpmt':>16} {'diff':>14} {'py_pvfact':>16} {'vba_pvfact':>16} {'fact_diff':>14}")
    print("-" * 120)
    for i in range(n):
        py_pv = pbs[i].dPVpmtCalc
        vba_pv = vba_pvpmt[i]
        py_f = pbs[i].dPVfactorCalc
        vba_f = vba_pvfact[i]
        if py_pv != 0 or vba_pv != 0:
            print(f"{i:3d} {py_pv:16.8f} {vba_pv:16.8f} {vba_pv-py_pv:14.10f} "
                  f"{py_f:16.10f} {vba_f:16.10f} {vba_f-py_f:14.10f}")


if __name__ == '__main__':
    # Bonds com yield match mas duration diff (melhor para isolar causa)
    # 5483424UN1: taxa diff=3e-6, dur diff=-0.027
    compare_periods('5483424UN1')
