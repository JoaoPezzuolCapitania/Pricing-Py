# Brazilian Fixed Income Pricing Engine

A pure Python pricing engine for Brazilian fixed income instruments (CRI, CRA, Debentures, LCA, FIDC).

Replicates the exact pricing logic of a VBA Excel add-in, supporting multiple indexers and day-count conventions used in the Brazilian market.

---

## Features

- Supported indexers: **CDI+**, **IPCA+**, **% CDI**, **Pre-fixed**, **IGPM+**, **IGPDI+**
- Day-count conventions: **Úteis (252)**, **30/360**, **1/360**, **1/365**, **21/252**
- Outputs: **Yield**, **PU**, **% PU Par**, **CDI Spread**, **IPCA Spread**, **Duration (Macaulay)**, **Over-TP**
- Newton-Raphson solver (no yield clamping, matches VBA behavior)
- Step-up yield schedules (`dicYields`)
- Incorporated yield (IncorpYield) support
- Monetary adjustment (AM): IPCA, IGPM, IGPDI with lag and pro-rata accrual
- Business day calendar: ANBIMA-based (or custom via env var)
- Batch processing with `fPricing_batch`

---

## Installation

```bash
pip install -r requirements.txt
```

---

## Configuration

The engine reads data files from directories configured via **environment variables**:

| Variable | Description | Default |
|---|---|---|
| `BOND_DATA_PATH` | Directory with `{CETIP}.txt` bond files | `./data/bonds` |
| `CURVES_DATA_PATH` | Directory with curve files (CDI.txt, AM.txt, DI1*.txt, etc.) | `./data/curves` |
| `FERIADOS_XLAM_PATH` | Path to an `.xlam`/`.xlsx` file with holidays in column A | *(uses ANBIMA)* |

Example (PowerShell):
```powershell
$env:BOND_DATA_PATH   = "C:\mydata\bonds"
$env:CURVES_DATA_PATH = "C:\mydata\curves"
```

Example (bash):
```bash
export BOND_DATA_PATH=/mydata/bonds
export CURVES_DATA_PATH=/mydata/curves
```

---

## Data File Formats

### Bond file: `{CETIP}.txt`

Tab-separated key-value pairs, followed by the payment schedule:

```
Código CETIP	XXXX99
Tipo	CRI
Data de Emissão	01/15/2020
Vencimento	01/15/2030
Indexador	CDI +
Taxa	0.012
PU	1000.00
Dias (juros)	Úteis
Atraso AM (meses)	0
Atraso PMT (du)	0
AM Total	False
Data
01/15/2021	0	0.1	0	0
01/15/2022	0	0.1	0	0
01/15/2030	0	0.8	0	0
```

Payment schedule columns: `date`, `incorp_yield`, `amort`, `extra_amort`, `fee`

Supported date formats in files: `MM/DD/YYYY`, `DD/MM/YYYY`, `YYYY-MM-DD`

### CDI.txt

Tab-separated, no header:
```
01/02/2026    0.1275
01/05/2026    0.1275
...
```

### AM.txt

Tab-separated, first row is header:
```
Date    IPCA    IGPM    IGPDI
01/2026    4.83    6.12    5.90
02/2026    4.91    ...
```

### DI1{YYYYMMDD}.txt / IPCA{YYYYMMDD}.txt / IGPM{YYYYMMDD}.txt

Tab-separated, rates in % (divided by 100 internally):
```
01/15/2026    12.75%
01/15/2027    12.40%
...
```

---

## Usage

### Single bond

```python
from pricing_engine_public.fpricing import fPricing

# PU → Yield (inp=2, out=1)
taxa = fPricing('XXXX99', '2026-04-15', 1050.50, 2, 1)

# PU → CDI Spread (inp=2, out=4)
spread = fPricing('XXXX99', '2026-04-15', 1050.50, 2, 4)

# PU → Duration (inp=2, out=8)
dur = fPricing('XXXX99', '2026-04-15', 1050.50, 2, 8)

# PU → Over-TP (inp=2, out=9)
over = fPricing('XXXX99', '2026-04-15', 1050.50, 2, 9)

# Yield → PU (inp=1, out=2)
pu = fPricing('XXXX99', '2026-04-15', 0.012, 1, 2)
```

### inp/out codes

| Code | inp meaning | out meaning |
|------|------------|-------------|
| 1 | Yield/Taxa | Yield/Taxa |
| 2 | PU | PU |
| 3 | % PU Par | % PU Par |
| 4 | CDI Spread | CDI Spread |
| 5 | IPCA Spread | IPCA Spread |
| 6 | IGPM Spread | — |
| 7 | IGPDI Spread | — |
| 8 | — | Duration |
| 9 | — | Over-TP |
| 0 | — | CashFlow |

### Batch (multiple bonds, single curve load)

```python
import pandas as pd
from pricing_engine_public.fpricing import fPricing_batch
from pricing_engine_public.curves import load_curves

df = pd.DataFrame({
    'cod_cetip': ['XXXX99', 'YYYY88'],
    'pu': [1050.50, 980.00]
})

curves = load_curves('2026-04-15')
df_result = fPricing_batch(df, '2026-04-15', inp=2, curves=curves)
# Adds columns: taxa, spread, duration, over_tp
```

---

## Module Structure

```
pricing_engine_public/
├── __init__.py       — exports fPricing, fPricing_batch
├── daycount.py       — business day calendar, fDays, fFactorDays
├── curves.py         — market curve loading (DI1, IPCA, IGPM, CDI, AM)
├── bond.py           — bond data loading, payment schedule generation
├── pv.py             — present value calculations
├── solver.py         — Newton-Raphson: yield, spread, duration, over-TP
└── fpricing.py       — main API: fPricing(), fPricing_batch()
```

---

## VBA → Python mapping

| VBA Function | Python Module | Python Function |
|---|---|---|
| `fBrWorkdays` | daycount.py | `brworkdays(d1, d2)` |
| `fBrWorkday` | daycount.py | `brworkday(dt, n)` |
| `fDays` | daycount.py | `fDays(dt, tipo)` |
| `fFactorDays` | daycount.py | `fFactorDays(...)` |
| `fLoadBondInfo` | bond.py | `load_bond(cetip, dt)` |
| `sRunPayments` | bond.py | `run_payments(...)` |
| `fReadCDI` | curves.py | `load_cdi(dt)` |
| `fReadAM` | curves.py | `load_am(dt)` |
| `fReadMellonCurves` | curves.py | `load_mellon_curve(...)` |
| `sLoadCurves` | curves.py | `load_curves(dt)` |
| `sPvCalc` | pv.py | `pv_calc(...)` |
| `fGetSpread` | pv.py | `fGetSpread(...)` |
| `fGetPrice` | pv.py | `get_price(...)` |
| `sGetTaxa` | solver.py | `get_taxa(...)` |
| `sGetSpreadYield` | solver.py | `get_spread_yield(...)` |
| `sDuration` | solver.py | `duration(...)` |
| `sGetOverTP` | solver.py | `get_over_tp_with_curves(...)` |
| `fPricing` | fpricing.py | `fPricing(...)` |

---

## Dependencies

```
bizdays
pandas
tqdm
```
