-- =============================================================================
-- sales_kpis.sql
-- Sales Performance KPI Queries for Excel Dashboard
-- Excel-PostgreSQL-Dashboard Repository
-- Author: Shiv Iyer | MinervaDB | ChistaDATA
-- License: MIT
-- =============================================================================
-- NOTE: Adjust table/column names to match your actual schema.
--       Queries are designed for PostgreSQL 13+
-- =============================================================================


-- -----------------------------------------------------------------------------
-- KPI 1: Total Revenue (Current Month)
-- Returns: Single scalar value
-- Usage: modQueryRunner.RunQuery_Scalar()
-- -----------------------------------------------------------------------------
SELECT
    ROUND(SUM(order_amount)::NUMERIC, 2) AS total_revenue_current_month
FROM orders
WHERE
    order_status  NOT IN ('cancelled', 'refunded')
    AND DATE_TRUNC('month', order_date) = DATE_TRUNC('month', CURRENT_DATE);


-- -----------------------------------------------------------------------------
-- KPI 2: Monthly Revenue Trend (Last 13 Months) - Time Series
-- Returns: Tabular result for chart
-- Usage: modQueryRunner.RunQuery_ToSheet()
-- -----------------------------------------------------------------------------
SELECT
    TO_CHAR(DATE_TRUNC('month', order_date), 'Mon YYYY')    AS month_label,
    DATE_TRUNC('month', order_date)                          AS month_date,
    ROUND(SUM(order_amount)::NUMERIC, 2)                     AS revenue,
    COUNT(DISTINCT order_id)                                  AS order_count,
    COUNT(DISTINCT customer_id)                              AS unique_customers,
    ROUND(SUM(order_amount)::NUMERIC / NULLIF(COUNT(DISTINCT order_id), 0), 2)
                                                             AS avg_order_value,

    -- Month-over-Month change
    ROUND(SUM(order_amount)::NUMERIC
          - LAG(SUM(order_amount)) OVER (ORDER BY DATE_TRUNC('month', order_date))
      , 2)                                                     AS mom_revenue_change,

    -- Month-over-Month % change
    ROUND(
          (SUM(order_amount)
              - LAG(SUM(order_amount)) OVER (ORDER BY DATE_TRUNC('month', order_date)))
          / NULLIF(LAG(SUM(order_amount)) OVER (ORDER BY DATE_TRUNC('month', order_date)), 0)
          * 100
      , 2)                                                     AS mom_pct_change,

    -- Year-over-Year comparison
    ROUND(SUM(order_amount)::NUMERIC
          - LAG(SUM(order_amount)) OVER (
              ORDER BY DATE_TRUNC('month', order_date)
              ROWS BETWEEN 12 PRECEDING AND 12 PRECEDING
            )
      , 2)                                                     AS yoy_revenue_change

FROM orders
WHERE
    order_status NOT IN ('cancelled', 'refunded')
    AND order_date >= CURRENT_DATE - INTERVAL '13 months'
GROUP BY
    DATE_TRUNC('month', order_date)
ORDER BY
    month_date;


-- -----------------------------------------------------------------------------
-- KPI 3: Sales Pipeline by Stage
-- Returns: Stage breakdown for funnel chart
-- -----------------------------------------------------------------------------
SELECT
    stage_name,
    COUNT(*)                                        AS deal_count,
    ROUND(SUM(deal_value)::NUMERIC, 2)              AS total_value,
    ROUND(AVG(deal_value)::NUMERIC, 2)              AS avg_deal_size,
    ROUND(SUM(deal_value) * win_probability / 100
          ::NUMERIC, 2)                               AS weighted_value,
    win_probability
FROM sales_pipeline
WHERE is_active = TRUE
GROUP BY
    stage_name,
    win_probability,
    stage_order
ORDER BY
    stage_order;


-- -----------------------------------------------------------------------------
-- KPI 4: Sales Rep Performance (Current Quarter)
-- Returns: Leaderboard table
-- -----------------------------------------------------------------------------
SELECT
    sr.rep_name,
    sr.region,
    sr.team,
    ROUND(SUM(o.order_amount)::NUMERIC, 2)          AS actual_revenue,
    sr.quarterly_quota                               AS quota,
    ROUND(
          SUM(o.order_amount) / NULLIF(sr.quarterly_quota, 0) * 100
      ::NUMERIC, 2)                                   AS quota_attainment_pct,
    COUNT(DISTINCT o.order_id)                       AS deals_closed,
    COUNT(DISTINCT o.customer_id)                    AS customers_served,
    ROUND(AVG(o.order_amount)::NUMERIC, 2)           AS avg_deal_size,
    ROUND(AVG(o.days_to_close)::NUMERIC, 1)          AS avg_sales_cycle_days
FROM orders o
JOIN sales_reps sr ON sr.rep_id = o.rep_id
WHERE
    o.order_status NOT IN ('cancelled', 'refunded')
    AND DATE_TRUNC('quarter', o.order_date) = DATE_TRUNC('quarter', CURRENT_DATE)
GROUP BY
    sr.rep_name,
    sr.region,
    sr.team,
    sr.quarterly_quota
ORDER BY
    actual_revenue DESC;


-- -----------------------------------------------------------------------------
-- KPI 5: Top 10 Products by Revenue (Current Month)
-- -----------------------------------------------------------------------------
SELECT
    p.product_name,
    p.category,
    COUNT(oi.order_item_id)                          AS units_sold,
    ROUND(SUM(oi.line_amount)::NUMERIC, 2)           AS revenue,
    ROUND(AVG(oi.unit_price)::NUMERIC, 2)            AS avg_selling_price,
    ROUND(
          SUM(oi.line_amount) - SUM(oi.unit_cost * oi.quantity)
      ::NUMERIC, 2)                                    AS gross_profit,
    ROUND(
          (SUM(oi.line_amount) - SUM(oi.unit_cost * oi.quantity))
          / NULLIF(SUM(oi.line_amount), 0) * 100
      ::NUMERIC, 2)                                    AS gross_margin_pct
FROM order_items oi
JOIN products p ON p.product_id = oi.product_id
JOIN orders o ON o.order_id = oi.order_id
WHERE
    o.order_status NOT IN ('cancelled', 'refunded')
    AND DATE_TRUNC('month', o.order_date) = DATE_TRUNC('month', CURRENT_DATE)
GROUP BY
    p.product_name,
    p.category
ORDER BY
    revenue DESC
LIMIT 10;


-- -----------------------------------------------------------------------------
-- KPI 6: Customer Acquisition & Retention
-- -----------------------------------------------------------------------------
SELECT
    DATE_TRUNC('month', first_order_date)           AS cohort_month,
    TO_CHAR(DATE_TRUNC('month', first_order_date), 'Mon YYYY')
                                                     AS cohort_label,
    COUNT(DISTINCT customer_id)                      AS new_customers,

    -- Retained in month 1 (30-day retention)
    COUNT(DISTINCT CASE
          WHEN last_order_date >= first_order_date + INTERVAL '30 days'
          THEN customer_id
      END)                                             AS retained_30d,

    -- Retained in month 3 (90-day retention)
    COUNT(DISTINCT CASE
          WHEN last_order_date >= first_order_date + INTERVAL '90 days'
          THEN customer_id
      END)                                             AS retained_90d,

    -- LTV proxy (total spend per customer)
    ROUND(AVG(total_spend)::NUMERIC, 2)              AS avg_ltv

FROM (
      SELECT
          customer_id,
          MIN(order_date) AS first_order_date,
          MAX(order_date) AS last_order_date,
          SUM(order_amount) AS total_spend
      FROM orders
      WHERE order_status NOT IN ('cancelled', 'refunded')
      GROUP BY customer_id
  ) customer_metrics
GROUP BY
    DATE_TRUNC('month', first_order_date)
ORDER BY
    cohort_month DESC
LIMIT 12;


-- -----------------------------------------------------------------------------
-- KPI 7: Win Rate by Source / Channel
-- -----------------------------------------------------------------------------
SELECT
    lead_source,
    COUNT(*)                                         AS total_opportunities,
    COUNT(*) FILTER (WHERE stage_name = 'Closed Won')   AS won,
    COUNT(*) FILTER (WHERE stage_name = 'Closed Lost')  AS lost,
    ROUND(
          COUNT(*) FILTER (WHERE stage_name = 'Closed Won')::NUMERIC
          / NULLIF(COUNT(*) FILTER (WHERE stage_name IN ('Closed Won','Closed Lost')), 0)
          * 100
      , 2)                                             AS win_rate_pct,
    ROUND(AVG(deal_value) FILTER (WHERE stage_name = 'Closed Won')::NUMERIC, 2)
                                                     AS avg_won_deal_value,
    ROUND(AVG(days_in_pipeline) FILTER (WHERE stage_name = 'Closed Won')::NUMERIC, 1)
                                                     AS avg_sales_cycle_won
FROM sales_pipeline
WHERE close_date >= CURRENT_DATE - INTERVAL '12 months'
GROUP BY lead_source
ORDER BY total_opportunities DESC;
