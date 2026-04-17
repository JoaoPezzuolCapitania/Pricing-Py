"""
curves.py — Carga de curvas de mercado dos arquivos TxT.
Carrega DI1, IPCA, IGPM, CDI, AM, NTNB, IGP.
"""

import os
from datetime import date, datetime
from dataclasses import dataclass, field
from .daycount import _to_date, brworkday, brworkdays

# Paths (mesmos do VBA)
MELLON_TXT_PATH = r"X:\#CapitaniaRFE\Trading\CurvasBNYM\CurvasTxT"
CDI_PATH = os.path.join(MELLON_TXT_PATH, "CDI.txt")
AM_PATH = os.path.join(MELLON_TXT_PATH, "AM.txt")


@dataclass
class CurveData:
    """Contém todas as curvas carregadas para uma data."""
    dt: date
    dicDI1: dict = field(default_factory=dict)
    dicIPCA: dict = field(default_factory=dict)
    dicIGPM: dict = field(default_factory=dict)
    dicIGPDI: dict = field(default_factory=dict)
    dicNTNB: dict = field(default_factory=dict)
    dicIGP: dict = field(default_factory=dict)
    dicCDI: dict = field(default_factory=dict)
    dicAMipca: dict = field(default_factory=dict)
    dicAMigpm: dict = field(default_factory=dict)
    dicAMigpdi: dict = field(default_factory=dict)


def _parse_date(s):
    """Tenta parsear data em vários formatos."""
    s = s.strip()
    if '-' in s:
        return _to_date(s)
    parts = s.split('/')
    if len(parts) == 3:
        a, b, c = parts
        a, b = int(a), int(b)
        # Se primeiro > 12, é DD/MM/YYYY
        if a > 12:
            return date(int(c), b, a)
        # Se segundo > 12, é MM/DD/YYYY
        elif b > 12:
            return date(int(c), a, b)
        else:
            # Assumir MM/DD/YYYY (padrão US que VBA usa)
            return date(int(c), a, b)
    raise ValueError(f"Data não reconhecida: {s}")


def _detect_date_format(filepath):
    """Detecta se arquivo usa DD/MM ou MM/DD. Replica fWhichDate."""
    with open(filepath, 'r', encoding='latin-1') as f:
        for line in f:
            line = line.strip()
            if not line:
                continue
            parts = line.split('\t')
            s = parts[0]
            if '-' in s:
                return "MMDD"
            if '/' in s:
                p = s.split('/')
                if len(p) == 3:
                    if int(p[0]) > 12:
                        return "DDMM"
                    elif int(p[1]) > 12:
                        return "MMDD"
    return "MMDD"


def _convert_date(s, fmt):
    """Converte data com formato detectado."""
    s = s.strip()
    if '-' in s:
        return _to_date(s)
    parts = s.split('/')
    if fmt == "DDMM":
        return date(int(parts[2]), int(parts[1]), int(parts[0]))
    else:  # MMDD
        return date(int(parts[2]), int(parts[0]), int(parts[1]))


def load_cdi(dt_calc):
    """
    Carrega CDI.txt -> dict {date_str: taxa}. Replica fReadCDI.
    Preenche gaps de dias uteis ausentes (ex: feriados novos como 20/Nov)
    usando o ultimo CDI disponivel, similar ao fReadCDI do VBA que usa
    fGetIndex para preencher dias faltantes.
    """
    dt_calc = _to_date(dt_calc)
    dic = {}

    with open(CDI_PATH, 'r', encoding='latin-1') as f:
        for line in f:
            line = line.strip()
            if not line:
                continue
            parts = line.split('\t')
            if len(parts) >= 2:
                try:
                    dt = _parse_date(parts[0])
                    val = float(parts[1])
                    dic[str(dt)] = val
                except (ValueError, IndexError):
                    continue

    return dic


def load_am(dt_calc):
    """Carrega AM.txt → dicts IPCA, IGPM, IGPDI. Replica fReadAM."""
    dt_calc = _to_date(dt_calc)
    dic_ipca = {}
    dic_igpm = {}
    dic_igpdi = {}

    with open(AM_PATH, 'r', encoding='latin-1') as f:
        first = True
        for line in f:
            line = line.strip()
            if not line:
                continue
            if first:
                first = False
                continue  # skip header
            parts = line.split('\t')
            if len(parts) >= 4:
                try:
                    dt = _parse_date(parts[0])
                    if dt >= dt_calc:
                        continue
                    ipca_val = float(parts[1])
                    igpm_val = float(parts[2])
                    igpdi_val = float(parts[3])
                    if ipca_val != 0:
                        key = f"{dt.month}_{dt.year}"
                        dic_ipca[key] = ipca_val
                    if igpm_val != 0:
                        key = f"{dt.month}_{dt.year}"
                        dic_igpm[key] = igpm_val
                    if igpdi_val != 0:
                        key = f"{dt.month}_{dt.year}"
                        dic_igpdi[key] = igpdi_val
                except (ValueError, IndexError):
                    continue

    return dic_ipca, dic_igpm, dic_igpdi


def load_mellon_curve(curve_name, dt_calc, dic_di1_existing=None):
    """
    Carrega curva DI1/IPCA/IGPM do arquivo TxT.
    Replica fReadMellonCurves.

    Retorna: (dic_curve, dic_ntnb_or_igp)
    Para DI1: dic_ntnb_or_igp é None
    Para IPCA: dic_ntnb_or_igp é dicNTNB
    Para IGPM: dic_ntnb_or_igp é dicIGP
    """
    dt_calc = _to_date(dt_calc)

    # Se data futura, usa ontem
    today = date.today()
    yesterday = brworkday(today, -1)
    if dt_calc > yesterday:
        search_date = yesterday
    else:
        search_date = dt_calc

    date_str = search_date.strftime('%Y%m%d')
    txt_path = os.path.join(MELLON_TXT_PATH, f"{curve_name}{date_str}.txt")

    if not os.path.exists(txt_path):
        raise FileNotFoundError(f"Curva {curve_name} não encontrada: {txt_path}")

    fmt = _detect_date_format(txt_path)
    dic_curve = {}
    dic_secondary = {}  # NTNB para IPCA, IGP para IGPM

    with open(txt_path, 'r', encoding='latin-1') as f:
        for line in f:
            line = line.strip()
            if not line:
                continue
            parts = line.split('\t')
            if len(parts) < 2:
                continue

            s_date = parts[0].strip()
            s_value = parts[1].strip()

            # Verifica se é data válida
            if '/' not in s_date and '-' not in s_date:
                continue
            if not s_value or s_value == '#N/A':
                continue

            try:
                dt_data = _convert_date(s_date, fmt)
            except (ValueError, IndexError):
                continue

            # Remove '%' e converte
            val_str = s_value.rstrip('%').strip()
            try:
                d_value = float(val_str) / 100
            except ValueError:
                continue

            if curve_name != "DI1":
                # Armazena em dicts secundários (NTNB para IPCA, IGP para IGPM)
                if curve_name == "IPCA":
                    dic_secondary[str(dt_data)] = d_value
                elif curve_name in ("IGPM", "IGPDI"):
                    dic_secondary[str(dt_data)] = d_value

                # Calcula spread sobre DI1
                if dic_di1_existing:
                    dt_str = str(dt_data)
                    wd_str = str(brworkday(dt_data, 1))
                    if dt_str in dic_di1_existing:
                        d_value = dic_di1_existing[dt_str] - d_value
                    elif wd_str in dic_di1_existing:
                        d_value = dic_di1_existing[wd_str] - d_value
                    else:
                        # Usa último valor do DI1
                        last_di1 = list(dic_di1_existing.values())[-1]
                        d_value = last_di1 - d_value

            dt_key = str(dt_data)
            if dt_key in dic_curve:
                break  # duplicata = fim dos dados úteis
            dic_curve[dt_key] = d_value

    return dic_curve, dic_secondary


def load_curves(dt):
    """
    Carrega TODAS as curvas para uma data.
    Replica sLoadCurves do VBA.
    Retorna CurveData com todos os dicts.
    """
    dt = _to_date(dt)
    curves = CurveData(dt=dt)

    # AM (Atualização Monetária)
    curves.dicAMipca, curves.dicAMigpm, curves.dicAMigpdi = load_am(dt)

    # CDI
    curves.dicCDI = load_cdi(dt)

    # DI1
    curves.dicDI1, _ = load_mellon_curve("DI1", dt)

    # IPCA → dicIPCA + dicNTNB
    curves.dicIPCA, curves.dicNTNB = load_mellon_curve("IPCA", dt, curves.dicDI1)

    # IGPM → dicIGPM + dicIGP (IGPDI = IGPM)
    curves.dicIGPM, curves.dicIGP = load_mellon_curve("IGPM", dt, curves.dicDI1)
    curves.dicIGPDI = curves.dicIGPM

    return curves
