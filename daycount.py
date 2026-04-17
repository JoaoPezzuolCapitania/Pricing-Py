"""
daycount.py — Funções de dias úteis e convenções de contagem de dias.
Usa a lista EXATA de feriados do FeriadosAddin.xlam para garantir
match perfeito com o VBA.
"""

from datetime import date, datetime, timedelta
from functools import lru_cache
import os
import zipfile
import xml.etree.ElementTree as ET
import bizdays
from bizdays import Calendar


def _load_vba_holidays():
    """Carrega lista de feriados do FeriadosAddin.xlam (mesma usada pelo VBA)."""
    xlam_path = r'C:\Add-in\Oficial\FeriadosAddin.xlam'
    if not os.path.exists(xlam_path):
        # Fallback para ANBIMA sem Nov 20
        cal_base = Calendar.load(name='ANBIMA')
        return [h for h in cal_base.holidays if not (h.month == 11 and h.day == 20)]

    with zipfile.ZipFile(xlam_path, 'r') as z:
        sheet_data = z.read('xl/worksheets/sheet1.xml')

    root = ET.fromstring(sheet_data)
    ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    feriados = []
    for row in root.findall(f'.//{{{ns}}}row'):
        for cell in row.findall(f'{{{ns}}}c'):
            col = cell.get('r', '')
            if col.startswith('A') and col != 'A1':
                val_elem = cell.find(f'{{{ns}}}v')
                if val_elem is not None:
                    try:
                        serial = float(val_elem.text)
                        dt = date(1899, 12, 30) + timedelta(days=int(serial))
                        feriados.append(dt)
                    except (ValueError, TypeError):
                        pass
    return feriados


_vba_holidays = _load_vba_holidays()
_cal = Calendar(
    holidays=_vba_holidays,
    weekdays=['Saturday', 'Sunday'],
    startdate='1994-01-01',
    enddate='2100-12-31',
    name='VBA_FERIADOS'
)


def _to_date(d):
    """Converte string ou datetime para date."""
    if isinstance(d, str):
        for fmt in ('%m/%d/%Y', '%Y-%m-%d', '%d/%m/%Y'):
            try:
                return datetime.strptime(d, fmt).date()
            except ValueError:
                continue
        raise ValueError(f"Formato de data não reconhecido: {d}")
    if isinstance(d, datetime):
        return d.date()
    return d


def brworkday(dt, n):
    """
    Retorna a data N dias úteis a partir de dt.
    Equivale a fbrworkday do VBA (usa Excel WORKDAY internamente).
    VBA WORKDAY(date, 0) retorna a própria data mesmo se não for dia útil.
    """
    dt = _to_date(dt)
    if n == 0:
        return dt
    return _cal.offset(dt, n)


def brworkdays(dt_from, dt_to):
    """
    Retorna contagem de dias úteis entre dt_from e dt_to.
    Equivale a fbrworkdays do VBA:
      iResult = NETWORKDAYS(dtA, dtB, feriados)
      If dtA <> WORKDAY(dtA - 1, 1, feriados) Then iResult = iResult + 1
    O VBA adiciona +1 quando dt_from NÃO é dia útil (weekend OU feriado em dia de semana).
    """
    dt_from = _to_date(dt_from)
    dt_to = _to_date(dt_to)
    # Ajuste: se início não é dia útil (weekend ou feriado), ir para dia útil anterior
    if not _cal.isbizday(dt_from):
        dt_from = _cal.preceding(dt_from)
    return _cal.bizdays(dt_from, dt_to) + 1


def fDays(dt, tipo):
    """
    Calcula dias no mês relativo à data.
    tipo = "DC": dias corridos (calendar days)
    tipo = "DU": dias úteis (business days)
    Replica fDays do VBA.
    """
    dt = _to_date(dt)
    i_month = dt.month
    i_year = dt.year
    i_day = dt.day

    # Mês anterior
    if i_month == 1:
        i_month = 12
        i_year = i_year - 1
    else:
        i_month = i_month - 1

    # Ajusta para dia válido
    i_error = 0
    while True:
        try:
            prev_date = date(i_year, i_month, i_day - i_error)
            break
        except ValueError:
            i_error += 1

    if tipo == "DC":
        return (dt - prev_date).days + i_error
    elif tipo == "DU":
        return brworkdays(prev_date, dt) - 1
    else:
        raise ValueError(f"Tipo de dias não reconhecido: {tipo}")


def fFactorDays(dt_today, dt_last, dt_next, sYdays):
    """
    Fator de accrual pro-rata entre duas datas.
    Replica fFactorDays do VBA.
    """
    dt_today = _to_date(dt_today)
    dt_last = _to_date(dt_last)
    dt_next = _to_date(dt_next)

    if sYdays in ("Úteis", "21/252"):
        num = brworkdays(dt_last, dt_today) - 1
        den = brworkdays(dt_last, dt_next) - 1
        return num / den if den != 0 else 0
    elif sYdays in ("30/360", "1/360", "1/365"):
        num = (dt_today - dt_last).days
        den = (dt_next - dt_last).days
        return num / den if den != 0 else 0
    else:
        # Fallback para Úteis quando convenção não especificada
        if not sYdays:
            num = brworkdays(dt_last, dt_today) - 1
            den = brworkdays(dt_last, dt_next) - 1
            return num / den if den != 0 else 0
        raise ValueError(f"Convenção de dias não reconhecida: {sYdays}")


def next_month(dt):
    """Avança um mês, mantendo o dia (ou último dia válido)."""
    dt = _to_date(dt)
    if dt.month + 1 > 12:
        m = 1
        y = dt.year + 1
    else:
        m = dt.month + 1
        y = dt.year

    error = 0
    while True:
        try:
            return date(y, m, dt.day - error)
        except ValueError:
            error += 1


def last_month(dt):
    """Recua um mês, mantendo o dia."""
    dt = _to_date(dt)
    if dt.month - 1 < 1:
        m = 12
        y = dt.year - 1
    else:
        m = dt.month - 1
        y = dt.year
    return date(y, m, dt.day)


def n_months(dt1, dt2):
    """Número de meses entre duas datas."""
    dt1 = _to_date(dt1)
    dt2 = _to_date(dt2)
    return (dt2.year - dt1.year) * 12 + (dt2.month - dt1.month)


def lag_month_key(dt, lag):
    """Retorna chave 'month_year' com lag de meses. Replica fLagMonth."""
    dt = _to_date(dt)
    m = dt.month - lag
    y = dt.year
    if m < 1:
        m = 12 + m
        y = y - 1
    return f"{m}_{y}"


def lag_date(dt, am_lag):
    """Retorna data com lag de meses. Replica fLagDate."""
    dt = _to_date(dt)
    m = dt.month - am_lag
    y = dt.year
    if m < 1:
        m = 12 + m
        y = y - 1
    d = dt.day
    while True:
        try:
            return date(y, m, d)
        except ValueError:
            d -= 1


def am_month_check(s_month, dt_to):
    """Replica fAmMonth — verifica se o mês de dt_to está na lista de meses AM."""
    dt_to = _to_date(dt_to)
    i_month = dt_to.month

    if not s_month:
        return 0

    if '_' in str(s_month):
        months = [int(x) for x in str(s_month).split('_')]
        if i_month in months:
            return i_month
    else:
        if i_month == int(s_month):
            return i_month

    return 0
