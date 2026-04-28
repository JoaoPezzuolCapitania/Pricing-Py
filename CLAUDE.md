# Pricing Engine Python — Documentação do Projeto

## O que é este projeto

Tradução completa do add-in VBA `PricingRFE.xlam` para Python puro.
O objetivo é substituir a dependência do Excel COM (win32com) por cálculos nativos em Python,
replicando bit-a-bit o comportamento da função VBA `fPricing(ativo, data, pu, inp, out)`.

---

## Status atual (2026-04-28, comparação PROD vs HML, dia 2026-04-24, 1047 ativos)

### Resumo PROD (VBA `PricingRFE.xlam`) vs HML (`pricing_engine` Python)

| Métrica   | Match exato (<=1e-6) | Diff < 1e-4    | Diff < 5e-4    | Diff < 1e-3    |
|-----------|----------------------|-----------------|-----------------|-----------------|
| taxa      | 1021/1047 (97,52%)   | 1031 (98,47%)  | 1042 (99,43%)  | 1043 (99,52%)  |
| spread    | 1025/1047 (97,90%)   | 1031 (98,47%)  | 1037 (98,95%)  | 1038 (99,04%)  |
| duration  | 1030/1047 (98,38%)   | 1038 (99,14%)  | 1038 (99,14%)  | 1038 (99,14%)  |
| overtp    | 833/1047 (79,56%)    | 1026 (97,99%)  | 1042 (99,43%)  | 1043 (99,52%)  |

**Match nas 4 métricas batendo TODAS (até 1e-3):** 1033/1047 = **98,66%**.
Excluindo `4083221SN1` (distressed): **1033/1042 = 99,14%**.

### Planilhas de comparação

- `X:\BDM\CRI\Comparativo_PROD_vs_HML_20260424.xlsx` — todas as linhas, ordem alfabética
- `X:\BDM\CRI\Comparativo_PROD_vs_HML_20260424_ordenado.xlsx` — ordenado por maior diff no topo

---

## Patch aplicado (sessão 2026-04-28)

### Patch 1 — `dAccInflFactor` persistente

**Problema:** o VBA `fGetFuture` mantém `tCalc.dAccInflFactor` GLOBAL entre chamadas de `fPricing`. Bonds com `sAMmonth` setado (IPCA+ com lag, ex: 19L0840477, 20L0687041) acumulam estado neste campo. PROD é gerado em batch — múltiplas chamadas seguidas — então herda esse estado entre bonds.

Antes do patch, Python resetava `dAccInflFactor = 0` a cada chamada (default da dataclass `CalcParams`). Resultado: 25 bonds IPCA+ com diff sistemática de ~0,022 no spread.

**Mudanças:**

1. **`bond.py:_get_future`** (linha 391-398): para períodos passados, retorna `1.0` direto sem mexer em `dAccInflFactor`. Antes Python tinha um bloco extra que rodava a acumulação para datas passadas — VBA não faz isso (sai com `Exit Function`).
2. **`fpricing.py`**: adicionado `_PERSISTENT_STATE = {'dAccInflFactor': 0.0}` módulo-level. `fPricing` lê esse valor antes de `run_payments` e salva o valor final ao retornar. Replica o comportamento do VBA onde `tCalc.dAccInflFactor` persiste entre chamadas.
3. **`fpricing.py`**: exposta `reset_persistent_state()` para resetar manualmente em testes.

**Impacto:** spread foi de 96,0% → 97,9% match exato, sem regressões em outras métricas. Net +25 ativos.

---

## Divergências remanescentes (14 bonds, dia 2026-04-24)

Investigação detalhada validou que **Python = VBA `PricingRFE.xlam` fresh** para esses bonds (testado via COM direta). O PROD difere por motivos externos ao motor.

### `4083221SN1` (5 linhas) — Distressed FIDC, yield ~190%
- PROD: yield = 1,990482 — gerado pela `RFE Pricing 3.0.xlsm` (calculadora SEPARADA) usada manualmente
- VBA `PricingRFE` fresh = Python = 1,910744 (Newton-Raphson)
- A função PV(yield) é flat entre 1,91 e 1,99 — múltiplos yields satisfazem PV ≈ PU
- Implementar a bisseção da RFE 3.0 em Python NÃO resolve (Python converge a 1,910760, RFE 3.0 a 1,990481, divergência por precisão Double específica do VBA)
- **Sem fix viável** sem replicar bit-a-bit o caminho de iteração VBA

### `AGIZ11` (4 linhas) — IPCA+ Debênture
- PROD: 0,077890; VBA fresh = Python = 0,077591
- `AGIZ11.txt` foi modificado em 2026-04-28 11:05 (MESMO DIA da nossa investigação)
- PROD foi gerado ANTES dessa modificação, com versão antiga do .txt
- **Causa: input mudou após geração do PROD** (não é bug do motor)

### `ALGTA4` (3 linhas) e `CASN23` (2 linhas) — IPCA+ Debênture
- PROD: ALGTA4=0,082733, CASN23=0,080740
- VBA fresh = Python = ALGTA4=0,083037, CASN23=0,081040
- **TODAS as 3 calculadoras** (PricingRFE, RFE 3.0, Python) dão o MESMO valor — só PROD difere
- `.txt` desses bonds não foi modificado recentemente
- **Origem do valor PROD desconhecida** — provavelmente outra calculadora manual ou edição direta no banco

---

## Calculadoras VBA paralelas (descobertas durante investigação)

Existem múltiplos engines de pricing convivendo:

1. **`C:\Add-in\Oficial\PricingRFE.xlam`** — addin oficial. Usado pelo `workflow_pricings.py` automático.
   - Algoritmo: Newton-Raphson clássico
   - Convergência: `Round(dPrice, 6) == Round(dPU, 6)`
2. **`X:\#CapitaniaRFE\Trading\Pricing\Operacional\RFE Pricing 3.0.xlsm`** — calculadora interativa.
   - Algoritmo: bisseção adaptativa (`dStart` halve em sign change)
   - Convergência: `Abs(dError) <= calc.dError` (tolerância absoluta dependente de PU)
   - Usada manualmente para distressed (ex: 4083221SN1)
3. **`X:\#CapitaniaRFE\Trading\Pricing\Calculadora FIDC CDI.xlsm`** — variante para FIDCs.

Para bonds normais, todas as calculadoras convergem ao mesmo ponto. Para distressed (yield extrema, função flat), divergem.

---

## Bugs corrigidos (histórico completo)

### Sessão 2026-04-28

1. **`_get_future` rodava acumulação para datas passadas** — `bond.py:391-411` tinha bloco
   especial que mexia em `dAccInflFactor` para `dt_to <= calc.dtDay`. VBA `fGetFuture`
   sai com `Exit Function` para datas passadas SEM tocar em `dAccInflFactor`. Removido.
   **Impacto**: parte do fix do spread IPCA+ (combinado com #2 abaixo).

2. **`dAccInflFactor` resetado a cada `fPricing`** — Python criava nova `CalcParams` por
   chamada (default 0.0). VBA mantém `tCalc.dAccInflFactor` global entre chamadas.
   Adicionado `_PERSISTENT_STATE` módulo-level em `fpricing.py` que persiste o valor
   entre chamadas, replicando o comportamento do VBA.
   **Impacto**: spread foi de 96,0% → 97,9% match exato (+25 ativos IPCA+ com sAMmonth).

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

## Tentativa abandonada: bisseção do RFE 3.0 para distressed (2026-04-28)

Tentei replicar o solver de bisseção do `RFE Pricing 3.0.xlsm` em Python para tentar bater
o `4083221SN1` (yield 1,99 no PROD vs 1,91 no Python).

**Implementação:**
```python
def get_taxa_bisection(calc, bond, periods, pbs, results):
    d_y = bond.dYield
    d_error = bond.dPU
    d_start = 1.0
    while abs(d_error) > calc.dError:
        d_price = sum(pv_calc(i, d_y, ...) for i in periods)
        if d_error * (d_price - calc.dPU) < 0:
            d_start /= 2
        d_error = d_price - calc.dPU
        d_mult = d_start * abs(d_error / calc.dPU)
        if d_error > 0: d_y += d_mult
        else: d_y -= d_mult
```

**Resultado:** Python bisseção converge a 1,910760 (basicamente igual ao Newton-Raphson).
NÃO bate com o 1,990481 do RFE 3.0 real.

**Por quê:** a função `PV(yield) - PU` é completamente flat entre 1,91 e 1,99 para esse
bond (yield extrema, função mal-condicionada). Múltiplos yields são "soluções" válidas.
Newton-Raphson e bisseção convergem em pontos diferentes da faixa flat. O VBA do RFE 3.0
termina em 1,99 por idiossincrasias da precisão Double do VB Office — não dá pra
replicar bit-a-bit em Python.

Patch revertido. Código removido. **Não é fixável sem replicar o caminho de iteração
exato do VBA, incluindo numeric subtleties do VB Office.**

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
  ├── _PERSISTENT_STATE.read     ← restaura dAccInflFactor da chamada anterior
  ├── load_curves(dt)            ← carrega DI1, IPCA, IGPM, CDI, AM
  ├── load_bond(ativo, dt)       ← carrega BondInfo, CalcParams, PeriodInfo[]
  ├── _setup_calc(inp, value)    ← configura tipo: PU, Yield, Spread, %PUPar
  ├── _setup_result(out)         ← configura resultado: Taxa, Duration, etc.
  ├── run_payments(bond, calc, periods, curves)
  │     ├── _load_yield(i)       ← yield com step-up (dicYields)
  │     ├── _calc_yield(i)       ← dYcdi, dYdi1, dYspread, dYtotal
  │     │   └── _get_future()    ← LE/ESCREVE calc.dAccInflFactor (estado!)
  │     ├── _incorporated_yield(i) ← ajuste IncorpYield (se bAmort100)
  │     ├── _calc_am(i)          ← dFatAm, dFatAmAcc, dSNA
  │     ├── _calc_pmt(i)         ← dPMTJuros, dPMTAmort, dPMTTotal
  │     └── _calc_pv_par(i)      ← dPVfactorPar, dPVpmtPar
  ├── get_par()                  ← sum(dPVpmtPar)
  ├── [PU input] get_taxa()      ← Newton-Raphson: PU → yield
  │     ├── get_price(yield, "Taxa") → pv_calc() → dPVpmtCalc
  │     └── get_spread_yield()   ← Newton-Raphson: PU → spread
  ├── [Duration out] duration()  ← sum(dPVpmtCalc * days) / price / 252
  ├── get_over_tp_with_curves()  ← spread sobre benchmark
  └── _PERSISTENT_STATE.write    ← persiste dAccInflFactor para próxima chamada
```

**Importante:** `_PERSISTENT_STATE` é módulo-level em `fpricing.py`. Replica
`tCalc.dAccInflFactor` global do VBA. Só fica diferente de zero para bonds com
`sAMmonth` setado (IPCA+ com lag mensal). Para bonds sem `sAMmonth`, é sempre 0.0
(efeito nulo). Use `reset_persistent_state()` no início de workflow novo se quiser
estado limpo.

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
- **Servidor HML**: `s14-db-hml.cb6sndi8pxjt.sa-east-1.rds.amazonaws.com`
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

## Como rodar o workflow oficial

```powershell
cd X:\BDM\CRI
python -m pricing_engine.run_pricing_hml 2026-04-24
```

- Lê `Renda_Fixa_Carteiras` do PROD (rds01).
- Aplica `read_carteiras_rf` (filtros, dedup por chave).
- Chama `fPricing` para cada bond (4 métricas: taxa, spread, duration, overtp).
- Salva em `Chaves_Pricing` no HML (s14-db-hml).

## Comparar HML vs PROD

Use os Excel gerados:
- `X:\BDM\CRI\Comparativo_PROD_vs_HML_20260424.xlsx`
- `X:\BDM\CRI\Comparativo_PROD_vs_HML_20260424_ordenado.xlsx`

Cores:
- 🟢 verde: diff <= 1e-4
- 🟡 amarelo: 1e-4 < diff <= 1e-3
- 🔴 vermelho: diff > 1e-3
