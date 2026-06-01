-- ============================================================
-- MinervaDB XL View - Finance KPI Queries
-- finance_kpis.sql
-- ============================================================
-- Financial performance KPIs for the MinervaDB XL View
-- Finance Dashboard. Covers P&L, cash flow, budget variance,
-- and financial health metrics.
-- ============================================================


-- ============================================================
-- KPI 1: Monthly P&L Summary
-- ============================================================
SELECT
    DATE_TRUNC('month', period_date) AS period_month,
    TO_CHAR(DATE_TRUNC('month', period_date), 'Mon YYYY') AS period_label,
    SUM(CASE WHEN account_category = 'Revenue' THEN actual_amount ELSE 0 END) AS total_revenue,
    SUM(CASE WHEN account_category = 'COGS' THEN actual_amount ELSE 0 END) AS total_cogs,
    SUM(CASE WHEN account_category = 'Revenue' THEN actual_amount ELSE 0 END) -
    SUM(CASE WHEN account_category = 'COGS' THEN actual_amount ELSE 0 END) AS gross_profit,
    SUM(CASE WHEN account_category = 'Operating Expense' THEN actual_amount ELSE 0 END) AS operating_expenses,
    SUM(CASE WHEN account_category = 'Revenue' THEN actual_amount ELSE 0 END) -
    SUM(CASE WHEN account_category IN ('COGS','Operating Expense') THEN actual_amount ELSE 0 END) AS ebitda,
    ROUND(
        100.0 * (
            SUM(CASE WHEN account_category = 'Revenue' THEN actual_amount ELSE 0 END) -
            SUM(CASE WHEN account_category = 'COGS' THEN actual_amount ELSE 0 END)
        ) / NULLIF(SUM(CASE WHEN account_category = 'Revenue' THEN actual_amount ELSE 0 END), 0),
        2
    ) AS gross_margin_pct
FROM financial_data
WHERE period_date >= DATE_TRUNC('year', CURRENT_DATE) - INTERVAL '1 year'
GROUP BY DATE_TRUNC('month', period_date)
ORDER BY period_month DESC;


-- ============================================================
-- KPI 2: Budget vs Actual Variance by Department
-- ============================================================
SELECT
    department,
    account_category,
    SUM(actual_amount)                                          AS actual_ytd,
    SUM(budget_amount)                                          AS budget_ytd,
    SUM(actual_amount) - SUM(budget_amount)                     AS variance_amount,
    ROUND(
        100.0 * (SUM(actual_amount) - SUM(budget_amount))
        / NULLIF(ABS(SUM(budget_amount)), 0),
        2
    )                                                           AS variance_pct,
    CASE
        WHEN ABS(SUM(actual_amount) - SUM(budget_amount)) / NULLIF(ABS(SUM(budget_amount)), 0) <= 0.05
            THEN 'On Track'
        WHEN (SUM(actual_amount) - SUM(budget_amount)) < 0
            THEN 'Under Budget'
        ELSE 'Over Budget'
    END                                                         AS budget_status
FROM financial_data
WHERE period_date BETWEEN DATE_TRUNC('year', CURRENT_DATE) AND CURRENT_DATE
GROUP BY department, account_category
ORDER BY department, account_category;


-- ============================================================
-- KPI 3: Year-over-Year Revenue Growth
-- ============================================================
WITH monthly_revenue AS (
    SELECT
        DATE_TRUNC('month', period_date)                        AS period_month,
        EXTRACT(YEAR FROM period_date)                          AS fiscal_year,
        EXTRACT(MONTH FROM period_date)                         AS fiscal_month,
        SUM(CASE WHEN account_category = 'Revenue' THEN actual_amount ELSE 0 END) AS revenue
    FROM financial_data
    WHERE period_date >= CURRENT_DATE - INTERVAL '2 years'
    GROUP BY DATE_TRUNC('month', period_date), EXTRACT(YEAR FROM period_date), EXTRACT(MONTH FROM period_date)
)
SELECT
    cy.period_month,
    cy.fiscal_year,
    cy.fiscal_month,
    cy.revenue                                                  AS current_year_revenue,
    py.revenue                                                  AS prior_year_revenue,
    cy.revenue - COALESCE(py.revenue, 0)                        AS yoy_variance,
    ROUND(
        100.0 * (cy.revenue - COALESCE(py.revenue, 0))
        / NULLIF(py.revenue, 0),
        2
    )                                                           AS yoy_growth_pct
FROM monthly_revenue cy
LEFT JOIN monthly_revenue py
    ON cy.fiscal_month = py.fiscal_month
    AND cy.fiscal_year = py.fiscal_year + 1
ORDER BY cy.period_month DESC;


-- ============================================================
-- KPI 4: Cash Flow Summary (Rolling 12 Months)
-- ============================================================
SELECT
    DATE_TRUNC('month', period_date)                            AS period_month,
    TO_CHAR(DATE_TRUNC('month', period_date), 'Mon YYYY')       AS period_label,
    SUM(CASE WHEN account_category = 'Operating Cash Flow'
             THEN actual_amount ELSE 0 END)                     AS operating_cf,
    SUM(CASE WHEN account_category = 'Investing Cash Flow'
             THEN actual_amount ELSE 0 END)                     AS investing_cf,
    SUM(CASE WHEN account_category = 'Financing Cash Flow'
             THEN actual_amount ELSE 0 END)                     AS financing_cf,
    SUM(CASE WHEN account_category IN (
             'Operating Cash Flow','Investing Cash Flow','Financing Cash Flow')
             THEN actual_amount ELSE 0 END)                     AS net_cash_flow
FROM financial_data
WHERE period_date >= CURRENT_DATE - INTERVAL '12 months'
GROUP BY DATE_TRUNC('month', period_date)
ORDER BY period_month DESC;


-- ============================================================
-- KPI 5: Expense Category Breakdown (Current Quarter)
-- ============================================================
SELECT
    account_category,
    account_name,
    department,
    SUM(actual_amount)                                          AS actual_amount,
    SUM(budget_amount)                                          AS budget_amount,
    ROUND(
        100.0 * SUM(actual_amount)
        / NULLIF(SUM(SUM(actual_amount)) OVER (PARTITION BY account_category), 0),
        2
    )                                                           AS pct_of_category
FROM financial_data
WHERE account_category = 'Operating Expense'
  AND period_date >= DATE_TRUNC('quarter', CURRENT_DATE)
GROUP BY account_category, account_name, department
ORDER BY actual_amount DESC;


-- ============================================================
-- KPI 6: Accounts Receivable Aging Summary
-- ============================================================
SELECT
    customer_segment,
    COUNT(*)                                                    AS invoice_count,
    SUM(invoice_amount)                                         AS total_ar,
    SUM(CASE WHEN days_outstanding <= 30 THEN invoice_amount ELSE 0 END) AS current_0_30,
    SUM(CASE WHEN days_outstanding BETWEEN 31 AND 60 THEN invoice_amount ELSE 0 END) AS days_31_60,
    SUM(CASE WHEN days_outstanding BETWEEN 61 AND 90 THEN invoice_amount ELSE 0 END) AS days_61_90,
    SUM(CASE WHEN days_outstanding > 90 THEN invoice_amount ELSE 0 END) AS over_90_days,
    ROUND(
        100.0 * SUM(CASE WHEN days_outstanding > 90 THEN invoice_amount ELSE 0 END)
        / NULLIF(SUM(invoice_amount), 0),
        2
    )                                                           AS pct_over_90
FROM accounts_receivable
WHERE invoice_status = 'Outstanding'
GROUP BY customer_segment
ORDER BY total_ar DESC;


-- ============================================================
-- KPI 7: Executive Finance Scorecard (Single Row)
-- ============================================================
WITH ytd_finance AS (
    SELECT
        SUM(CASE WHEN account_category = 'Revenue' THEN actual_amount ELSE 0 END) AS ytd_revenue,
        SUM(CASE WHEN account_category = 'Revenue' THEN budget_amount ELSE 0 END) AS ytd_revenue_budget,
        SUM(CASE WHEN account_category = 'COGS' THEN actual_amount ELSE 0 END) AS ytd_cogs,
        SUM(CASE WHEN account_category = 'Operating Expense' THEN actual_amount ELSE 0 END) AS ytd_opex
    FROM financial_data
    WHERE period_date BETWEEN DATE_TRUNC('year', CURRENT_DATE) AND CURRENT_DATE
)
SELECT
    ytd_revenue                                                 AS ytd_revenue,
    ytd_revenue_budget                                          AS ytd_revenue_budget,
    ROUND(100.0 * ytd_revenue / NULLIF(ytd_revenue_budget, 0), 1) AS revenue_attainment_pct,
    ytd_revenue - ytd_cogs                                      AS ytd_gross_profit,
    ROUND(100.0 * (ytd_revenue - ytd_cogs) / NULLIF(ytd_revenue, 0), 1) AS gross_margin_pct,
    ytd_revenue - ytd_cogs - ytd_opex                           AS ytd_ebitda,
    ROUND(100.0 * (ytd_revenue - ytd_cogs - ytd_opex) / NULLIF(ytd_revenue, 0), 1) AS ebitda_margin_pct,
    TO_CHAR(CURRENT_DATE, 'Month DD, YYYY')                     AS as_of_date
FROM ytd_finance;
