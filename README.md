# SaaS ARR Movement Dashboard — FP&A Analytics Suite

> "We need someone young who understands AI and knows how to use it efficiently."  
> — VP Finance, B2B SaaS company ($3B valuation)

## Overview
End-to-end FP&A analytics pipeline simulating a real B2B SaaS monthly reporting workflow — from raw data to board-ready report.

## Stack
| Layer | Tool |
|---|---|
| Data generation | Python |
| Aggregation | SQL (SQLite-compatible) |
| Dashboard | Excel + VBA |
| AI Commentary | Anthropic Claude API via VBA |
| Board Report | PowerPoint (10 slides) |

## Key Metrics Computed
- ARR Movements: New MRR, Expansion, Contraction, Churn
- NRR / GRR
- Rule of 40
- CAC Payback
- EBITDA Margin (benchmarked against Datadog FY2024 10-K)

## How the Claude API Integration Works
A VBA macro reads live metrics from the Excel dashboard, builds a structured prompt, calls the Anthropic Claude API via AppleScript/curl, and writes board-ready executive commentary directly into the dashboard in under 5 seconds.

## Dataset
Synthetic — 1,006 rows, 24 months (Jan 2023 – Dec 2024). No real customer or revenue data used. P&L ratios benchmarked against Datadog Inc. (NASDAQ: DDOG) quarterly filings FY2023–FY2024 (10-K).

## Files
- `data/saas_arr_raw.csv` — Raw synthetic dataset
- `sql/saas_arr_waterfall_sqlite.sql` — Multi-CTE SQL aggregation script
- `excel/SaaS_ARR_Dashboard_charts.xlsx` — Excel dashboard with VBA macros and charts
- `excel/SaaS_Dashboard_VBA.bas` — VBA source code (Claude API integration)
- `report/SaaS_Monthly_Report_Dec2024.pptx` — 10-slide monthly board report
