# Pricing Engine Python — Documentação do Projeto

## O que é este projeto

Tradução completa do add-in VBA `PricingRFE.xlam` para Python puro.
O objetivo é substituir a dependência do Excel COM (win32com) por cálculos nativos em Python,
replicando bit-a-bit o comportamento da função VBA `fPricing(ativo, data, pu, inp, out)`.

---

## Status atual (2026-04-14, comparação COM direta, 318 ativos)

### Taxa (yield)

| Métrica Taxa          | Valor              |
|-----------------------|--------------------|
| Match exato (<=1e-6)  | 303/318 (95.3%)    |
| Diff < 0.0001         | 314/318 (98.7%)    |
| Diff < 0.001          | 317/318 (99.7%)    |

### Duration

| Métrica Duration      | Valor              |
|-----------------------|--------------------|
| Match exato (<=1e-6)  | 304/318 (95.6%)    |
| Diff < 0.0001         | 307/318 (96.5%)    |
| Diff < 0.001          | 309/318 (97.2%)    |

### Distribuição de outliers por indexador

| Indexador | Exact | Com diff |
|-----------|-------|----------|
| CDI +     | 150   | 12       |
| IPCA +    | 147   | 2 (tiny) |
| % CDI     | 2     | 0        |
| Pré       | 5     | 0        |

### Planilha de comparação

`X:\BDM\CRI\Comparativo_Python_vs_VBA_20260326.xlsx` — 318 ativos com colunas:
Ativo, PU, Taxa PY/VBA/Diff, Spread PY/VBA/Diff, Duration PY/VBA/Diff, OverTP PY/VBA/Diff

---

## Bugs corrigidos (histórico completo)

### Sessão 2026-04-14

4. **`_parse_date_bond` sem suporte YYYY-MM-DD** — O parser só reconhecia datas com `/`
   (MM/DD/YYYY ou DD/MM/YYYY). Bonds como DESK19 usam formato `YYYY-MM-DD` com separador `-`.
   O parser retornava 0 períodos e fPricing falhava com `IndexError`.
   **Fix**: Adicionado bloco para detectar `-` como separador e parsear como YYYY-MM-DD.
   **Impacto**: DESK19 agora precifica (taxa match exato com VBA).

5. **Calendário bizdays limitado a 2078** — O calendário `bizdays` herdava o range dos
   feriados (1994-2078). Bonds como 6039625SR1 (FIDC, vencimento 2099) causavam
   `DateOutOfRange` na linha `brworkday(pi.dtDay - timedelta(days=1), 1)`.
   **Fix**: Estendido o calendário com `startdate='1994-01-01', enddate='2100-12-31'`.
   **Impacto**: 6039625SR1 agora precifica (taxa diff = -0.000001 vs VBA).

### Sessão anterior

1. **`brworkday(n=0)` para non-business days** — VBA `WORKDAY(date, 0)` retorna a
   própria data mesmo se for weekend/feriado. Python retornava o dia útil seguinte.
   Fix: `if n == 0: return dt`. **Impacto: +31 ativos em match <0.0001**.

2. **Newton-Raphson sem clamp** — VBA `sGetTaxa` não limita yield a [-1, 5000].
   Python tinha clamp que impedia convergência para yields extremos (ex: 6361%).
   Fix: removido clamp, ajustado TryAgain loop.

3. **`brworkdays` para feriados em dia de semana** — VBA `fBrWorkdays` adiciona +1
   quando `dtFrom` não é dia útil (weekend OU feriado). Python só checava weekends.
   Fix: `if not _cal.isbizday(dt_from)`. **Impacto: +73 ativos em match exato**.

---

## Investigação em andamento: Duration CDI+ (12 outliers)

### O problema

12 bonds CDI+ têm duration Python **sempre menor** que VBA. A diferença NÃO vem da yield
(testado usando a yield exata do VBA → duration continua diferindo). O problema está na
**distribuição dos `dPVpmtCalc`** por período.

### O que já foi verificado

- Fórmula `sDuration` do VBA vs `duration()` Python: **IDÊNTICAS**
- Fórmula `sPvCalc` vs `pv_calc`: **IDÊNTICAS**
- Fórmula `fGetSpread` vs `fGetSpread`: **IDÊNTICAS**
- Fórmula `fGetFuture` (DI1) vs `_get_future`: **IDÊNTICAS**
- DI1 curve lookups para BBML13 (63 períodos): **TODOS BATEM**
- `brworkdays` para todos os pares de período de BBML13: **TODOS BATEM**
- Fluxo `sGetTaxa` → `sDuration` no VBA vs `get_taxa` → `duration` no Python: **MESMO FLUXO**
- `sGetSpreadYield` (chamado dentro de `sGetTaxa`) usa `fGetPrice("Spread")` que modifica
  `dPVpmtSpread` e NÃO `dPVpmtCalc` → duration não é afetada

### Hipótese atual

A cadeia `dPVfactorCalc[i] = dPVfactorCalc[i-1] * fGetSpread(yield) * dYdi1[i]` acumula uma
diferença microscópica nos fatores DI1 ou CDI que se amplifica ao longo de muitos períodos.
Mesmo que PU total bata (Newton-Raphson ajusta yield para isso), a distribuição de PVpmtCalc
entre períodos difere, causando um weighted average (duration) diferente.

### Top 6 maiores outliers

| Ativo       | Dur PY    | Dur VBA   | Diff      |
|-------------|-----------|-----------|-----------|
| LCAMC2      | 2.982798  | 3.237431  | -0.254633 |
| 22L1212138  | 3.113936  | 3.141400  | -0.027464 |
| 5483424UN1  | 3.065510  | 3.092634  | -0.027124 |
| 6039625SR1  | 5.676398  | 5.694390  | -0.017992 |
| 6083125SR1  | 1.846871  | 1.862805  | -0.015934 |
| 5241224SR3  | 1.771320  | 1.784846  | -0.013526 |

### Próximo passo

Extrair os `dPVpmtCalc` per-período do VBA via COM para comparação direta.
A injeção de macro (`VBProject.VBComponents.Add`) requer "Trust access to VBA project
object model" habilitado no Excel. Alternativa: escrever macro standalone no VBA que
exporta dados para arquivo texto.

---

## BBML13 — investigação detalhada (deprioritizado)

Bond CDI+ com IncorpYield + StepUp. PU bate exato, mas duration difere por ~0.006 e
taxa difere por ~0.030. Investigação extensiva feita sem encontrar root cause:
- Todas as fórmulas idênticas ao VBA
- DI1 lookups todos batem
- brworkdays todos batem
- Outros bonds CDI+ com IncorpYield+StepUp (19E0322333, 22K0640841) batem perfeitamente
- Causa provavelmente é a mesma dos 12 outliers CDI+ (acúmulo na cadeia PV)

---

## Arquitetura dos módulos

```
pricing_engine/
├── CLAUDE.md       ← este arquivo
├── __init__.py
├── daycount.py     ← calendário (1994-2100), dias úteis, fDays, fFactorDays
├── curves.py       ← carga de curvas DI1/IPCA/IGPM/CDI/AM
├── bond.py         ← load_bond, run_payments, estruturas de dados
├── pv.py           ← pv_calc, fGetSpread, fGetSpreadPerc, get_price
├── solver.py       ← Newton-Raphson: get_taxa, get_spread_yield, duration
├── fpricing.py     ← fPricing() e fPricing_batch() — API principal
└── validate.py     ← compara Python vs Excel COM (para testes)
```

---

## Mapeamento VBA → Python

| Função VBA               | Módulo Python    | Função Python            |
|--------------------------|------------------|--------------------------|
| `fBrWorkdays`            | daycount.py      | `brworkdays(d1, d2)`     |
| `fBrWorkday`             | daycount.py      | `brworkday(dt, n)`       |
| `fDays`                  | daycount.py      | `fDays(dt, tipo)`        |
| `fFactorDays`            | daycount.py      | `fFactorDays(...)`       |
| `fLoadBondInfo`          | bond.py          | `load_bond(cetip, dt)`   |
| `sRunPayments`           | bond.py          | `run_payments(...)`      |
| `sIncorporatedYield`     | bond.py          | `_incorporated_yield(...)` |
| `fGetFuture`             | bond.py          | `_get_future(...)`       |
| `fGetCDI`                | bond.py          | `_get_cdi(...)`          |
| `fReadCDI`               | curves.py        | `load_cdi(dt)`           |
| `fReadAM`                | curves.py        | `load_am(dt)`            |
| `fReadMellonCurves`      | curves.py        | `load_mellon_curve(...)` |
| `sLoadCurves`            | curves.py        | `load_curves(dt)`        |
| `sPvCalc`                | pv.py            | `pv_calc(...)`           |
| `sPvSpreadInp`           | pv.py            | `pv_spread_inp(...)`     |
| `sPvSpreadRes`           | pv.py            | `pv_spread_res(...)`     |
| `fGetSpread`             | pv.py            | `fGetSpread(...)`        |
| `fGetSpreadPerc`         | pv.py            | `fGetSpreadPerc(...)`    |
| `fGetPrice`              | pv.py            | `get_price(...)`         |
| `sGetTaxa`               | solver.py        | `get_taxa(...)`          |
| `sGetSpreadYield`        | solver.py        | `get_spread_yield(...)`  |
| `sDuration`              | solver.py        | `duration(...)`          |
| `sGetOverTP`             | solver.py        | `get_over_tp_with_curves(...)` |
| `sGetPar`                | solver.py        | `get_par(...)`           |
| `sGetPU`                 | solver.py        | `get_pu(...)`            |
| `fPricing`               | fpricing.py      | `fPricing(...)`          |

---

## Estruturas de dados principais (bond.py)

### `BondInfo` — dados estáticos do ativo
```python
sCETIP: str           # código CETIP
sType: str            # tipo (CRI, CRA, LCA, Debênture, FIDC)
dtIssuance: date      # emissão
dtMaturity: date      # vencimento
sIndex: str           # "CDI +", "IPCA +", "Pré", "% CDI", "IGPM +", "IGPDI +"
dYield: float         # taxa base
dPU: float            # PU emissão
dicYields: dict       # mudanças de taxa programadas {date_str: taxa} (step-up)
sSpreadIndexInp: str  # index do spread input
sSpreadIndexRes: str  # index do spread result
```

### `CalcParams` — parâmetros de cálculo
```python
sType: str            # "PU", "Yield", "%PUPar", "Taxa Spread"
sResult: str          # "Yield", "PU", "%PUPar", "Duration", "OverB", "CashFlow"
dtDay: date           # data de cálculo
dYield, dSpread, dPU, dPercPuPar: float
iAMlag: int           # lag de AM em meses
iCDIlag: int          # lag CDI em meses
iPmtlag: int          # lag de pagamento em du
bAmort100: bool       # amortização 100% (controla IncorpYield)
bPmtIncorp: bool      # pagamento incorporado
sAMmonth: str         # meses de AM ("1_7" = jan e jul)
sAMcarac: str         # característica AM ("Só AM Positiva", etc.)
sYdays: str           # convenção: "Úteis", "30/360", "1/360", "1/365", "21/252"
dAccInflFactor: float # fator inflação acumulado (para AM month check)
```

### `PeriodInfo` — schedule de pagamentos (do arquivo .txt)
```python
dtDay: date           # data do período
dtDayPMT: date        # data de pagamento (com lag)
dIncorpYield: float   # taxa de incorporação (0 = sem, 1 = incorpora)
dAmort: float         # amortização programada (%)
dExtrAmort: float     # amortização extraordinária
dMultaFee: float      # multa/fee
```

### `PeriodBond` — resultado calculado por período
```python
dSN, dSNA: float      # saldo nominal e saldo nominal ajustado
dFatAm, dFatAmAcc: float  # fator AM e acumulado
dYinf: float          # yield inflação (IPCA/IGPM composto)
dYcdi: float          # yield CDI diário acumulado
dYdi1: float          # fator DI1 forward acumulado no período
dYspread: float       # fator spread (yield * DI1)
dYtotal: float        # fator total
dPMTJuros: float      # pagamento de juros
dPMTIncorpJuros: float # juros incorporados
dPMTAmort: float      # pagamento de amortização
dPMTAmortExtr: float  # amortização extraordinária
dPMTTotal: float      # pagamento total
dPVfactorCalc: float  # fator de desconto para taxa
dPVpmtCalc: float     # PMT descontado (para taxa)
dPVfactorSpread: float # fator de desconto para spread
dPVpmtSpread: float   # PMT descontado (para spread)
dPVfactorPar: float   # fator de desconto para par
dPVpmtPar: float      # PMT descontado (para par)
dYAuxCurveInp: float  # curva auxiliar (spread input)
dYAuxCurveRes: float  # curva auxiliar (spread result)
```

---

## Caminhos dos arquivos de dados

```
X:\#CapitaniaRFE\Trading\Ativos Capitânia\{CETIP}.txt  ← dados dos ativos
X:\#CapitaniaRFE\Trading\CurvasBNYM\CurvasTxT\CDI.txt  ← CDI diário
X:\#CapitaniaRFE\Trading\CurvasBNYM\CurvasTxT\AM.txt   ← AM (IPCA/IGPM/IGPDI)
X:\#CapitaniaRFE\Trading\CurvasBNYM\CurvasTxT\DI1{YYYYMMDD}.txt   ← curva DI1
X:\#CapitaniaRFE\Trading\CurvasBNYM\CurvasTxT\IPCA{YYYYMMDD}.txt  ← curva IPCA
X:\#CapitaniaRFE\Trading\CurvasBNYM\CurvasTxT\IGPM{YYYYMMDD}.txt  ← curva IGPM
C:\Add-in\Oficial\FeriadosAddin.xlam  ← feriados do calendário VBA (908 feriados, 1994-2078)
C:\Add-in\Oficial\PricingRFE.xlam     ← add-in VBA (referência / validação)
```

---

## API de uso (fpricing.py)

### Uso simples (1 ativo)
```python
from pricing_engine.fpricing import fPricing

# PU → Taxa (inp=2, out=1)
taxa = fPricing('CBAN72', '2026-03-27', 1050.50, 2, 1)

# PU → Spread CDI (inp=2, out=4)
spread = fPricing('CBAN72', '2026-03-27', 1050.50, 2, 4)

# PU → Duration (inp=2, out=8)
dur = fPricing('CBAN72', '2026-03-27', 1050.50, 2, 8)

# PU → OverTP (inp=2, out=9)
over = fPricing('CBAN72', '2026-03-27', 1050.50, 2, 9)
```

### Códigos de inp/out
```
inp: 1=Taxa/Yield, 2=PU, 3=%PUPar, 4=Spread CDI, 5=Spread IPCA, 6=Spread IGPM, 7=Spread IGPDI
out: 1=Taxa, 2=PU, 3=%PUPar, 4=Spread CDI, 5=Spread IPCA, 8=Duration, 9=OverTP, 0=CashFlow
```

### Uso batch (múltiplos ativos)
```python
from pricing_engine.fpricing import fPricing_batch
from pricing_engine.curves import load_curves

# Carrega curvas 1x para toda a lista
curves = load_curves('2026-03-27')

# df deve ter colunas: 'cod_cetip' (ou 'ativo') e 'pu'
df_result = fPricing_batch(df, '2026-03-27', inp=2, curves=curves)
# Retorna df com colunas adicionais: taxa, spread, duration, over_tp
```

---

## Fluxo de cálculo (fPricing completo)

```
fPricing(ativo, dt, value, inp, out)
  ├── load_curves(dt)            ← carrega DI1, IPCA, IGPM, CDI, AM
  ├── load_bond(ativo, dt)       ← carrega BondInfo, CalcParams, PeriodInfo[]
  ├── _setup_calc(inp, value)    ← configura tipo: PU, Yield, Spread, %PUPar
  ├── _setup_result(out)         ← configura resultado: Taxa, Duration, etc.
  ├── run_payments(bond, calc, periods, curves)
  │     ├── _load_yield(i)       ← yield com step-up (dicYields)
  │     ├── _calc_yield(i)       ← dYcdi, dYdi1, dYspread, dYtotal
  │     ├── _incorporated_yield(i) ← ajuste IncorpYield (se bAmort100)
  │     ├── _calc_am(i)          ← dFatAm, dFatAmAcc, dSNA
  │     ├── _calc_pmt(i)         ← dPMTJuros, dPMTAmort, dPMTTotal
  │     └── _calc_pv_par(i)      ← dPVfactorPar, dPVpmtPar
  ├── get_par()                  ← sum(dPVpmtPar)
  ├── [PU input] get_taxa()      ← Newton-Raphson: PU → yield
  │     ├── get_price(yield, "Taxa") → pv_calc() → dPVpmtCalc
  │     └── get_spread_yield()   ← Newton-Raphson: PU → spread
  ├── [Duration out] duration()  ← sum(dPVpmtCalc * days) / price / 252
  └── get_over_tp_with_curves()  ← spread sobre benchmark
```

---

## Formatos de data em arquivos .txt de bonds

O parser `_parse_date_bond` reconhece 3 formatos:
- `MM/DD/YYYY` — padrão US do VBA (separador `/`, primeiro campo <= 12)
- `DD/MM/YYYY` — detectado quando primeiro campo > 12
- `YYYY-MM-DD` — formato ISO (separador `-`, primeiro campo = 4 dígitos)

---

## Calendário bizdays

```python
_cal = Calendar(
    holidays=_vba_holidays,         # 908 feriados de FeriadosAddin.xlam
    weekdays=['Saturday', 'Sunday'],
    startdate='1994-01-01',
    enddate='2100-12-31',           # Estendido para cobrir bonds até 2099
    name='VBA_FERIADOS'
)
```

- Se `FeriadosAddin.xlam` não existir, usa ANBIMA sem 20/Nov como fallback
- Range estendido para 2100 garante bonds de longo prazo (ex: 6039625SR1 vencimento 2099)

---

## Dependências Python

```
bizdays          ← calendário ANBIMA base
pandas           ← manipulação de dados
tqdm             ← barra de progresso no batch
pyodbc           ← conexão SQL Server (para ler ativos de PROD)
win32com         ← apenas para validate.py / diagnostico_com.py
pythoncom        ← apenas para validate.py
openpyxl         ← leitura/escrita de Excel
```

---

## Notas importantes

- **Servidor PROD**: `rds01.capitania.net`, database `db_asset_carteiras`
- **Calendário VBA**: carregado de `FeriadosAddin.xlam` → 908 feriados (1994-2078)
  - Calendário Python estendido até 2100 (sem feriados após 2078, apenas weekends)
- **Formato datas nos arquivos .txt**: MM/DD/YYYY, DD/MM/YYYY, ou YYYY-MM-DD
- **Curvas**: arquivos com data D-1 (usa `yesterday` se `dt_calc > yesterday`)
- **CDI.txt**: contém CDI de todos os dias históricos
- **AM.txt**: inclui IPCA, IGPM, IGPDI mensais acumulados
- **VBA `sGetTaxa` chama `sGetSpreadYield` internamente** — ambos Python e VBA fazem isso.
  `sGetSpreadYield` usa `fGetPrice("Spread")` que modifica `dPVpmtSpread`, NÃO `dPVpmtCalc`.
- **VBA `sPricing` (sub)** difere de `fPricing` (function): `sPricing` sempre roda
  `sGetTaxa → sDuration → sGetSpreadYield`, enquanto `fPricing` é condicional.
  A comparação COM usa `sPricing` + `fResult`.
