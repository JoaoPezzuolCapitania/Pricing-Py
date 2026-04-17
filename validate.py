"""
validate.py — Compara resultados do Python vs Excel (via COM).
Roda N ativos pelos dois caminhos e gera relatório de diferenças.
"""

import pandas as pd
import numpy as np
import pythoncom
import win32com.client
from tqdm import tqdm

from .fpricing import fPricing
from .curves import load_curves


def _excel_fpricing(xl, ws, s_bond, dt_day, d_value, i_info, i_result):
    """Calcula fPricing via Excel COM (mesmo que o workflow atual)."""
    try:
        ws.Range('A1').Value = s_bond
        ws.Range('A2').Value = dt_day
        ws.Range('A3').Value = d_value
        ws.Range('A4').Value = f'=fPricing(A1, A2, A3, {i_info}, {i_result})'
        return ws.Range('A4').Value
    except Exception as e:
        return f"ERRO: {e}"


def validate(df, dt, inp=2, tolerance=1e-4, max_ativos=None):
    """
    Compara Python vs Excel para um DataFrame de ativos.

    Args:
        df: DataFrame com 'cod_cetip' e 'pu'
        dt: data de cálculo (str 'YYYY-MM-DD')
        inp: tipo de input (default 2 = PU)
        tolerance: diferença máxima aceitável
        max_ativos: limitar número de ativos (para teste rápido)

    Returns:
        DataFrame com comparação detalhada
    """
    import tempfile, uuid, os, openpyxl

    if max_ativos:
        df = df.head(max_ativos).copy()
    else:
        df = df.copy()

    n = len(df)
    print(f"Validando {n} ativos para data {dt}...")
    print()

    # Carregar curvas Python (1x)
    print("Carregando curvas (Python)...")
    curves = load_curves(dt)
    print("Curvas carregadas!")
    print()

    # Preparar Excel
    print("Abrindo Excel...")
    pythoncom.CoInitialize()
    xl = win32com.client.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.Workbooks.Open(r'C:\Add-in\Oficial\PricingRFE.xlam')

    wb_temp = openpyxl.Workbook()
    tf = tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx', prefix=f'validate_{uuid.uuid4()}_')
    wb_temp.save(tf.name)
    wb = xl.Workbooks.Open(tf.name)
    ws = wb.ActiveSheet
    print("Excel pronto!")
    print()

    metrics = [
        ('taxa', 1),
        ('spread', 4),
        ('duration', 8),
        ('over_tp', 9),
    ]

    results = []

    try:
        for i in tqdm(range(n), desc="Validando"):
            cetip = df.iloc[i]['cod_cetip'] if 'cod_cetip' in df.columns else df.iloc[i]['ativo']
            pu = df.iloc[i]['pu']

            row = {'ativo': cetip, 'pu': pu}

            for metric_name, i_result in metrics:
                # Python
                py_val = fPricing(cetip, dt, pu, inp, i_result, curves)
                # Excel
                xl_val = _excel_fpricing(xl, ws, cetip, dt, pu, inp, i_result)

                row[f'{metric_name}_py'] = py_val
                row[f'{metric_name}_xl'] = xl_val

                # Diferença
                try:
                    py_f = float(py_val) if py_val is not None else np.nan
                    xl_f = float(xl_val) if xl_val is not None else np.nan
                    diff = abs(py_f - xl_f)
                    row[f'{metric_name}_diff'] = diff
                    row[f'{metric_name}_ok'] = diff <= tolerance
                except (TypeError, ValueError):
                    row[f'{metric_name}_diff'] = np.nan
                    row[f'{metric_name}_ok'] = str(py_val) == str(xl_val)

            results.append(row)

    finally:
        wb.Close(False)
        xl.Quit()
        pythoncom.CoUninitialize()
        try:
            tf.close()
            os.remove(tf.name)
        except Exception:
            pass

    df_results = pd.DataFrame(results)

    # Resumo
    print()
    print("=" * 60)
    print("  RESUMO DA VALIDAÇÃO")
    print("=" * 60)
    print(f"  Ativos testados: {n}")
    for metric_name, _ in metrics:
        ok_col = f'{metric_name}_ok'
        if ok_col in df_results.columns:
            n_ok = df_results[ok_col].sum()
            n_fail = n - n_ok
            pct = n_ok / n * 100
            status = "OK" if n_fail == 0 else "DIFERENÇAS"
            print(f"  {metric_name:>10s}: {n_ok}/{n} ({pct:.1f}%) — {status}")

            if n_fail > 0:
                diff_col = f'{metric_name}_diff'
                diffs = df_results[df_results[ok_col] == False]
                if len(diffs) > 0 and diff_col in diffs.columns:
                    max_diff = diffs[diff_col].max()
                    mean_diff = diffs[diff_col].mean()
                    print(f"             max_diff={max_diff:.8f}  mean_diff={mean_diff:.8f}")
    print("=" * 60)

    return df_results
