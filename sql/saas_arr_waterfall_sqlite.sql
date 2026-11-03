-- ============================================================
-- SaaS ARR Waterfall Metrics | SQLite-compatible
-- Source table: saas_arr_raw
-- Output: Monthly Starting MRR, movements, Ending MRR, NRR, GRR
-- Author: Théo | Hult International Business School
-- ============================================================

-- ============================================================
-- CTE 1: Aggregate raw movements into monthly buckets
-- ============================================================
WITH monthly_movements AS (
    SELECT
        Month,
        COALESCE(SUM(CASE WHEN Type = 'New'         THEN MRR ELSE 0 END), 0) AS new_mrr,
        COALESCE(SUM(CASE WHEN Type = 'Expansion'   THEN MRR ELSE 0 END), 0) AS expansion_mrr,
        COALESCE(SUM(CASE WHEN Type = 'Contraction' THEN MRR ELSE 0 END), 0) AS contraction_mrr,
        COALESCE(SUM(CASE WHEN Type = 'Churn'       THEN MRR ELSE 0 END), 0) AS churn_mrr
    FROM saas_arr_raw
    GROUP BY Month
),

-- ============================================================
-- CTE 2: Compute Running Ending MRR
-- ============================================================
mrr_running AS (
    SELECT
        Month,
        new_mrr,
        expansion_mrr,
        contraction_mrr,
        churn_mrr,
        SUM(new_mrr + expansion_mrr + contraction_mrr + churn_mrr)
            OVER (ORDER BY Month ASC ROWS BETWEEN UNBOUNDED PRECEDING AND CURRENT ROW)
            AS ending_mrr
    FROM monthly_movements
),

-- ============================================================
-- CTE 3: Derive Starting MRR via LAG
-- ============================================================
mrr_with_starting AS (
    SELECT
        Month,
        new_mrr,
        expansion_mrr,
        contraction_mrr,
        churn_mrr,
        ending_mrr,
        COALESCE(LAG(ending_mrr, 1) OVER (ORDER BY Month ASC), 0) AS starting_mrr
    FROM mrr_running
)

-- ============================================================
-- FINAL OUTPUT
-- SQLite change: ROUND(x, 2) directly, no ::NUMERIC cast
-- ============================================================
SELECT
    Month,
    ROUND(starting_mrr,    2) AS starting_mrr,
    ROUND(new_mrr,         2) AS new_mrr,
    ROUND(expansion_mrr,   2) AS expansion_mrr,
    ROUND(contraction_mrr, 2) AS contraction_mrr,
    ROUND(churn_mrr,       2) AS churn_mrr,
    ROUND(ending_mrr,      2) AS ending_mrr,
    ROUND(ending_mrr - starting_mrr, 2) AS net_mrr_change,
    ROUND(ending_mrr * 12, 2) AS arr,

    -- NRR
    CASE
        WHEN starting_mrr = 0 THEN NULL
        ELSE ROUND(
            (starting_mrr + expansion_mrr + contraction_mrr + churn_mrr) / starting_mrr * 100,
            2
        )
    END AS nrr_pct,

    -- GRR (capped at 100)
    CASE
        WHEN starting_mrr = 0 THEN NULL
        ELSE ROUND(
            MIN(
                (starting_mrr + contraction_mrr + churn_mrr) / starting_mrr * 100,
                100
            ),
            2
        )
    END AS grr_pct

FROM mrr_with_starting
ORDER BY Month ASC;
