"""
curves.py — Market curve loading from TxT files.
Loads DI1, IPCA, IGPM, CDI, AM (monetary adjustment) curves.

Configure the data directory via environment variables:
    CURVES_DATA_PATH  — directory containing CDI.txt, AM.txt, DI1{date}.txt, etc.
                        Defaults to './data/curves'
"""

import os
from datetime import date, datetime
from dataclasses import dataclass, field
from .daycount import _to_date, brworkday, brworkdays

CURVES_DATA_PATH = os.environ.get('CURVES_DATA_PATH', os.path.join(os.path.dirname(__file__), 'data', 'curves'))
CDI_PATH = os.path.join(CURVES_DATA_PATH, "CDI.txt")
AM_PATH = os.path.join(CURVES_DATA_PATH, "AM.txt")


@dataclass
class CurveData:
    """Contains all loaded curves for a given date."""
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
    """Tries to parse date in multiple formats."""
    s = s.strip()
    if '-' in s:
        return _to_date(s)
    parts = s.split('/')
    if len(parts) == 3:
        a, b, c = parts
        a, b = int(a), int(b)
        if a > 12:
            return date(int(c), b, a)
        elif b > 12:
            return date(int(c), a, b)
        else:
            return date(int(c), a, b)
    raise ValueError(f"Unrecognized date: {s}")


def _detect_date_format(filepath):
    """Detects whether file uses DD/MM or MM/DD. Equivalent to VBA fWhichDate."""
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
    """Converts date with detected format."""
    s = s.strip()
    if '-' in s:
        return _to_date(s)
    parts = s.split('/')
    if fmt == "DDMM":
        return date(int(parts[2]), int(parts[1]), int(parts[0]))
    else:
        return date(int(parts[2]), int(parts[0]), int(parts[1]))


def load_cdi(dt_calc):
    """
    Loads CDI.txt → dict {date_str: rate}.
    Equivalent to VBA fReadCDI.

    File format (tab-separated):
        date    cdi_daily_rate
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
    """
    Loads AM.txt → dicts for IPCA, IGPM, IGPDI monetary adjustment factors.
    Equivalent to VBA fReadAM.

    File format (tab-separated, first row is header):
        date    ipca_factor    igpm_factor    igpdi_factor
    """
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
                continue
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
    Loads a DI1/IPCA/IGPM curve from a TxT file.
    Equivalent to VBA fReadMellonCurves.

    File naming: {curve_name}{YYYYMMDD}.txt  (e.g. DI120260415.txt)
    File format (tab-separated):
        date    rate_pct    [optional secondary rate]

    Returns: (dic_curve, dic_secondary)
        For DI1:  dic_secondary is None
        For IPCA: dic_secondary contains NTN-B rates
        For IGPM: dic_secondary contains IGP rates
    """
    dt_calc = _to_date(dt_calc)

    today = date.today()
    yesterday = brworkday(today, -1)
    if dt_calc > yesterday:
        search_date = yesterday
    else:
        search_date = dt_calc

    date_str = search_date.strftime('%Y%m%d')
    txt_path = os.path.join(CURVES_DATA_PATH, f"{curve_name}{date_str}.txt")

    if not os.path.exists(txt_path):
        raise FileNotFoundError(f"Curve {curve_name} not found: {txt_path}")

    fmt = _detect_date_format(txt_path)
    dic_curve = {}
    dic_secondary = {}

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

            if '/' not in s_date and '-' not in s_date:
                continue
            if not s_value or s_value == '#N/A':
                continue

            try:
                dt_data = _convert_date(s_date, fmt)
            except (ValueError, IndexError):
                continue

            val_str = s_value.rstrip('%').strip()
            try:
                d_value = float(val_str) / 100
            except ValueError:
                continue

            if curve_name != "DI1":
                if curve_name == "IPCA":
                    dic_secondary[str(dt_data)] = d_value
                elif curve_name in ("IGPM", "IGPDI"):
                    dic_secondary[str(dt_data)] = d_value

                if dic_di1_existing:
                    dt_str = str(dt_data)
                    wd_str = str(brworkday(dt_data, 1))
                    if dt_str in dic_di1_existing:
                        d_value = dic_di1_existing[dt_str] - d_value
                    elif wd_str in dic_di1_existing:
                        d_value = dic_di1_existing[wd_str] - d_value
                    else:
                        last_di1 = list(dic_di1_existing.values())[-1]
                        d_value = last_di1 - d_value

            dt_key = str(dt_data)
            if dt_key in dic_curve:
                break
            dic_curve[dt_key] = d_value

    return dic_curve, dic_secondary


def load_curves(dt):
    """
    Loads ALL curves for a given date.
    Equivalent to VBA sLoadCurves.
    Returns CurveData with all dicts populated.
    """
    dt = _to_date(dt)
    curves = CurveData(dt=dt)

    curves.dicAMipca, curves.dicAMigpm, curves.dicAMigpdi = load_am(dt)
    curves.dicCDI = load_cdi(dt)
    curves.dicDI1, _ = load_mellon_curve("DI1", dt)
    curves.dicIPCA, curves.dicNTNB = load_mellon_curve("IPCA", dt, curves.dicDI1)
    curves.dicIGPM, curves.dicIGP = load_mellon_curve("IGPM", dt, curves.dicDI1)
    curves.dicIGPDI = curves.dicIGPM

    return curves
