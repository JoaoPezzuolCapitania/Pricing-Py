"""
exemplo_uso.py — Usage examples for the Brazilian Fixed Income Pricing Engine.

Before running, set the environment variables pointing to your data:

    Windows (PowerShell):
        $env:BOND_DATA_PATH   = "C:\\path\\to\\bonds"
        $env:CURVES_DATA_PATH = "C:\\path\\to\\curves"

    Linux/macOS:
        export BOND_DATA_PATH=/path/to/bonds
        export CURVES_DATA_PATH=/path/to/curves

Bond files must be named {CETIP}.txt (e.g. XXXX99.txt).
See README.md for file format details.
"""

import os
import pandas as pd
from pricing_engine_public.fpricing import fPricing, fPricing_batch
from pricing_engine_public.curves import load_curves

# ----------------------------------------------------------------
# Configuration — edit these or set environment variables instead
# ----------------------------------------------------------------
BOND_DATA_PATH   = os.environ.get('BOND_DATA_PATH',   './data/bonds')
CURVES_DATA_PATH = os.environ.get('CURVES_DATA_PATH', './data/curves')

os.environ.setdefault('BOND_DATA_PATH',   BOND_DATA_PATH)
os.environ.setdefault('CURVES_DATA_PATH', CURVES_DATA_PATH)

DT = '2026-04-15'
CETIP = 'XXXX99'   # replace with an actual bond code
PU = 1050.50       # replace with actual PU


# ----------------------------------------------------------------
# Example 1: single bond, PU → all outputs
# ----------------------------------------------------------------
def example_single_bond():
    print("=" * 50)
    print("Example 1 — Single bond pricing")
    print("=" * 50)

    curves = load_curves(DT)

    taxa    = fPricing(CETIP, DT, PU, 2, 1, curves)
    spread  = fPricing(CETIP, DT, PU, 2, 4, curves)
    dur     = fPricing(CETIP, DT, PU, 2, 8, curves)
    over_tp = fPricing(CETIP, DT, PU, 2, 9, curves)

    print(f"Bond:     {CETIP}")
    print(f"Date:     {DT}")
    print(f"PU:       {PU}")
    print(f"Yield:    {taxa}")
    print(f"Spread:   {spread}")
    print(f"Duration: {dur}")
    print(f"Over-TP:  {over_tp}")


# ----------------------------------------------------------------
# Example 2: yield → PU (forward pricing)
# ----------------------------------------------------------------
def example_yield_to_pu():
    print()
    print("=" * 50)
    print("Example 2 — Yield → PU")
    print("=" * 50)

    curves = load_curves(DT)
    yield_input = 0.012   # 1.2% a.a.
    pu_result = fPricing(CETIP, DT, yield_input, 1, 2, curves)

    print(f"Bond:        {CETIP}")
    print(f"Input yield: {yield_input:.4%}")
    print(f"PU result:   {pu_result}")


# ----------------------------------------------------------------
# Example 3: batch pricing for multiple bonds
# ----------------------------------------------------------------
def example_batch():
    print()
    print("=" * 50)
    print("Example 3 — Batch pricing")
    print("=" * 50)

    df = pd.DataFrame({
        'cod_cetip': [CETIP, CETIP],
        'pu': [1050.50, 1045.00]
    })

    curves = load_curves(DT)
    df_result = fPricing_batch(df, DT, inp=2, curves=curves)

    print(df_result[['cod_cetip', 'pu', 'taxa', 'spread', 'duration', 'over_tp']].to_string(index=False))


if __name__ == '__main__':
    example_single_bond()
    example_yield_to_pu()
    example_batch()
