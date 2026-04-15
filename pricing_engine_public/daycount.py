"""
daycount.py — Business day functions and day-count conventions for the Brazilian market.

Uses the ANBIMA holiday calendar by default. To match exact VBA behavior, provide the
path to a FeriadosAddin.xlam file via the FERIADOS_XLAM_PATH environment variable.
"""

import os
from datetime import date, datetime, timedelta
from functools import lru_cache
import zipfile
import xml.etree.ElementTree as ET
import bizdays
from bizdays import Calendar


def _load_holidays():
    """
    Loads holiday list. Uses ANBIMA calendar by default.

    To load from a custom Excel file (xlam/xlsx with dates in column A),
    set the FERIADOS_XLAM_PATH environment variable to the file path.
    """
    xlam_path = os.environ.get('FERIADOS_XLAM_PATH', '')
    if xlam_path and os.path.exists(xlam_path):
        with zipfile.ZipFile(xlam_path, 'r') as z:
            sheet_data = z.read('xl/worksheets/sheet1.xml')

        root = ET.fromstring(sheet_data)
        ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
        holidays = []
        for row in root.findall(f'.//{{{ns}}}row'):
            for cell in row.findall(f'{{{ns}}}c'):
                col = cell.get('r', '')
                if col.startswith('A') and col != 'A1':
                    val_elem = cell.find(f'{{{ns}}}v')
                    if val_elem is not None:
                        try:
                            serial = float(val_elem.text)
                            dt = date(1899, 12, 30) + timedelta(days=int(serial))
                            holidays.append(dt)
                        except (ValueError, TypeError):
                            pass
        return holidays

    # Fallback: ANBIMA without Nov 20 (Dia da Consciência Negra)
    cal_base = Calendar.load(name='ANBIMA')
    return [h for h in cal_base.holidays if not (h.month == 11 and h.day == 20)]


_holidays = _load_holidays()
_cal = Calendar(
    holidays=_holidays,
    weekdays=['Saturday', 'Sunday'],
    startdate='1994-01-01',
    enddate='2100-12-31',
    name='BR_HOLIDAYS'
)


def _to_date(d):
    """Converts string or datetime to date."""
    if isinstance(d, str):
        for fmt in ('%m/%d/%Y', '%Y-%m-%d', '%d/%m/%Y'):
            try:
                return datetime.strptime(d, fmt).date()
            except ValueError:
                continue
        raise ValueError(f"Unrecognized date format: {d}")
    if isinstance(d, datetime):
        return d.date()
    return d


def brworkday(dt, n):
    """
    Returns the date N business days from dt.
    Equivalent to VBA fbrworkday (uses Excel WORKDAY internally).
    VBA WORKDAY(date, 0) returns the date itself even if not a business day.
    """
    dt = _to_date(dt)
    if n == 0:
        return dt
    return _cal.offset(dt, n)


def brworkdays(dt_from, dt_to):
    """
    Returns count of business days between dt_from and dt_to.
    Equivalent to VBA fbrworkdays:
      iResult = NETWORKDAYS(dtA, dtB, holidays)
      If dtA <> WORKDAY(dtA - 1, 1, holidays) Then iResult = iResult + 1
    Adds +1 when dt_from is NOT a business day (weekend OR weekday holiday).
    """
    dt_from = _to_date(dt_from)
    dt_to = _to_date(dt_to)
    if not _cal.isbizday(dt_from):
        dt_from = _cal.preceding(dt_from)
    return _cal.bizdays(dt_from, dt_to) + 1


def fDays(dt, tipo):
    """
    Calculates days in the month relative to the date.
    tipo = "DC": calendar days
    tipo = "DU": business days
    Equivalent to VBA fDays.
    """
    dt = _to_date(dt)
    i_month = dt.month
    i_year = dt.year
    i_day = dt.day

    if i_month == 1:
        i_month = 12
        i_year = i_year - 1
    else:
        i_month = i_month - 1

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
        raise ValueError(f"Unrecognized day type: {tipo}")


def fFactorDays(dt_today, dt_last, dt_next, sYdays):
    """
    Pro-rata accrual factor between two dates.
    Equivalent to VBA fFactorDays.
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
        if not sYdays:
            num = brworkdays(dt_last, dt_today) - 1
            den = brworkdays(dt_last, dt_next) - 1
            return num / den if den != 0 else 0
        raise ValueError(f"Unrecognized day convention: {sYdays}")


def next_month(dt):
    """Advances one month, keeping the day (or last valid day)."""
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
    """Recedes one month, keeping the day."""
    dt = _to_date(dt)
    if dt.month - 1 < 1:
        m = 12
        y = dt.year - 1
    else:
        m = dt.month - 1
        y = dt.year
    return date(y, m, dt.day)


def n_months(dt1, dt2):
    """Number of months between two dates."""
    dt1 = _to_date(dt1)
    dt2 = _to_date(dt2)
    return (dt2.year - dt1.year) * 12 + (dt2.month - dt1.month)


def lag_month_key(dt, lag):
    """Returns 'month_year' key with month lag. Equivalent to VBA fLagMonth."""
    dt = _to_date(dt)
    m = dt.month - lag
    y = dt.year
    if m < 1:
        m = 12 + m
        y = y - 1
    return f"{m}_{y}"


def lag_date(dt, am_lag):
    """Returns date with month lag. Equivalent to VBA fLagDate."""
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
    """Equivalent to VBA fAmMonth — checks if dt_to month is in the AM month list."""
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
