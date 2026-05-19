# Pulsar — Cross-Asset Scoring Dashboards

Streamlit app replicating the *Pulsar* visual style for two cross-asset
scoring dashboards:

- **Global Rates Scoring** — 10 sovereign bond markets
- **Global Equity Scoring** — 17 equity index futures

Both run off a single Excel workbook of Bloomberg data and recompute when
you change the as-of date or pillar weights.

## Layout

```
.
├── app.py                  # loader + factor math + renderer
├── requirements.txt
├── README.md
└── data/
    └── pulsar_data.xlsx    # Bloomberg data, one sheet per factor
```

## Run

```bash
pip install -r requirements.txt
streamlit run app.py
```

Or deploy to Streamlit Community Cloud by pointing it at this repo.

## Model

### Rates dashboard

| Pillar  | Factors                                                |
|---------|--------------------------------------------------------|
| Macro   | GDP YoY, CPI YoY (inverted), Fiscal balance (% GDP)    |
| Markets | 3M yield momentum (inverted), 10Y carry, real yield    |

Score = weighted blend of the two pillars. Each factor z-scored across
the 10-country panel and sign-adjusted so higher = constructive for the
bond market.

### Equity dashboard

| Pillar  | Factors                                                |
|---------|--------------------------------------------------------|
| Macro   | Growth, Inflation (inv), Deficit, ToT 3M, FCI          |
| EPS Δ   | 3M % change in bottom-up FY1 EPS estimate              |

Macro factors are looked up by country (each index → its country via map).
FCI is keyed by region (US, EZ, UK, CN — only four FCIs published reliably).
Trailing performance (5D/1M/3M) and 30-day realized vol shown alongside
the score but not part of it.

## Reading the dashboard

- **Green pip** — top 2 ranks
- **Red pip** — bottom rank
- **Orange pip** — pillar-score tension, or incomplete data coverage
- **Shaded columns** — pillar composites
- **Green/red text** — sign of the z-score

## Refreshing

Update `data/pulsar_data.xlsx` with fresh Bloomberg pulls and the dashboard
recomputes on next load.
