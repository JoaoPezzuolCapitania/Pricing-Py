"""
run_pricing_hml.py - Replica o workflow_pricings.py (linhas 31-50)
mas usando a pricing engine Python em vez do VBA COM.

Le carteiras RF do PROD, roda fPricing Python, salva em Chaves_Pricing no HML.

Uso:
    python -m pricing_engine.run_pricing_hml 2025-06-23
"""

import sys
import os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, r'X:\BDM\CRI')

import pyodbc
import pandas as pd
import numpy as np
from tqdm import tqdm
from sqlalchemy import create_engine
import urllib
from pricing_engine.fpricing import fPricing
from pricing_engine.curves import load_curves

# ============================================================
# Conexoes (mesmas do fpricing_capitania_parallel_hml.py)
# ============================================================
SERVER_PROD = "rds01.capitania.net"
SERVER_HML = "s14-db-hml.cb6sndi8pxjt.sa-east-1.rds.amazonaws.com"
DATABASE = "db_asset_carteiras"
UID = "SISTEMA_LOCAL"
PWD = "7Vl2D@HH0n"


def _connect(server):
    return pyodbc.connect(
        f"DRIVER={{ODBC Driver 17 for SQL Server}};"
        f"Server={server};Database={DATABASE};"
        f"UID={UID};PWD={PWD};"
    )


def _engine(server):
    params = urllib.parse.quote_plus(
        f"DRIVER={{ODBC Driver 17 for SQL Server}};"
        f"SERVER={server};"
        f"DATABASE={DATABASE};"
        f"UID={UID};"
        f"PWD={PWD};"
    )
    return create_engine(f"mssql+pyodbc:///?odbc_connect={params}",
                         fast_executemany=True)


# ============================================================
# 1) Ler carteiras RF do PROD (replica read_carteiras_rf)
# ============================================================
def read_carteiras_rf(dt):
    conn = _connect(SERVER_PROD)

    print("Carregando Renda_Fixa_Carteiras...\n")

    query_rf = f"SELECT * FROM Renda_Fixa_Carteiras WHERE data = '{dt}'"
    df_rf = pd.read_sql_query(query_rf, conn).reset_index(drop=True)

    query_infos = f"SELECT * FROM Infos_Carteiras WHERE data = '{dt}'"
    df_infos = pd.read_sql_query(query_infos, conn).reset_index(drop=True)
    df_infos = df_infos.rename(columns={'administrador': 'adm_fundo'})
    df_infos = df_infos[['fundo', 'adm_fundo']]

    conn.close()

    # Mesma logica do fpricing_capitania_parallel_hml
    ativos_no_corp = (
        df_rf[df_rf['fundo'] != 'CAPITANIA CORP FIDC']['cod_cetip']
        .unique().tolist()
    )
    ativos_corp_import = (
        df_rf[(df_rf['fundo'] == 'CAPITANIA CORP FIDC') &
              (~df_rf['cod_cetip'].isin(ativos_no_corp))]
        .drop('id', axis=1)
    )

    fundos_ignorar = ['FIRFEMBVCP', 'CAPITANIA CORP FIDC']
    df_rf = (
        df_rf[~df_rf['fundo'].isin(fundos_ignorar)]
        .drop('id', axis=1)
        .reset_index(drop=True)
    )

    if len(ativos_corp_import) > 0:
        df_rf = pd.concat([df_rf, ativos_corp_import], ignore_index=True)

    list_rf = ['Debenture', 'Debênture', 'FIDC', 'CRI', 'CRA', 'TP', 'NC']
    df_rf = df_rf[df_rf['tipo'].isin(list_rf)]
    df_rf = df_rf[~df_rf['cod_cetip'].str.startswith('LFT_')].reset_index(drop=True)
    df_rf = df_rf[~df_rf['cod_cetip'].str.startswith('NTNB_')].reset_index(drop=True)
    df_rf = df_rf[df_rf['cod_cetip'] != 'Compromissada']

    df_rf_infos = pd.merge(df_rf, df_infos, how='left', on='fundo')
    df_rf_infos = df_rf_infos.assign(
        key=df_rf_infos['cod_cetip'].astype(str)
            + df_rf_infos['pu'].astype(str)
            + df_rf_infos['adm_fundo'].astype(str)
    )
    df_rf_infos = df_rf_infos.drop_duplicates('key')
    df_rf_infos = df_rf_infos[['data', 'adm_fundo', 'cod_cetip', 'tipo', 'pu']]

    print(f"Total ativos RF: {len(df_rf_infos)}")
    return df_rf_infos


# ============================================================
# 2) Rodar fPricing Python (replica fpricing_carteiras)
# ============================================================
def run_fpricing_python(df, dt):
    print(f"\nCarregando curvas para {dt}...")
    curves = load_curves(dt)
    print("Curvas carregadas!\n")

    results = []

    for i in tqdm(range(len(df)), desc="Pricing Python"):
        ativo = df.iloc[i]['cod_cetip']
        pu = df.iloc[i]['pu']
        adm = df.iloc[i].get('adm_fundo', '')

        row = {
            'data': dt,
            'adm': adm,
            'ativo': ativo,
            'pu': pu,
        }

        for col, out_code in [('taxa', 1), ('spread', 4), ('duration', 8), ('overtp', 9)]:
            try:
                val = fPricing(ativo, dt, pu, 2, out_code, curves)
                row[col] = val if not isinstance(val, str) else None
            except Exception:
                row[col] = None

        results.append(row)

    df_result = pd.DataFrame(results)

    # Filtrar erros (mesma logica do workflow: remove -2146826273)
    for col in ['taxa', 'spread', 'duration', 'overtp']:
        df_result[col] = pd.to_numeric(df_result[col], errors='coerce')
        df_result = df_result[df_result[col] != -2146826273].reset_index(drop=True)

    n_ok = df_result['taxa'].notna().sum()
    print(f"\nResultado: {len(df_result)} registros, {n_ok} com taxa OK")

    return df_result


# ============================================================
# 3) Salvar no HML (replica append_to_table)
# ============================================================
def append_to_hml(df):
    engine = _engine(SERVER_HML)
    df.to_sql(name='Chaves_Pricing', con=engine,
              if_exists='append', index=False, schema='dbo')
    print(f"Salvo {len(df)} registros em HML.Chaves_Pricing")


# ============================================================
# Main
# ============================================================
if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("Uso: python -m pricing_engine.run_pricing_hml <data>")
        print("Ex:  python -m pricing_engine.run_pricing_hml 2025-06-23")
        sys.exit(1)

    dt = sys.argv[1]

    print(f"{'='*60}")
    print(f"  PRICING PYTHON -> HML")
    print(f"  Data: {dt}")
    print(f"{'='*60}\n")

    # 1) Ler carteiras
    df_ativos = read_carteiras_rf(dt)

    # 2) Rodar pricing Python
    df_result = run_fpricing_python(df_ativos, dt)

    # 3) Salvar no HML
    append_to_hml(df_result)

    print(f"\nConcluido!")
