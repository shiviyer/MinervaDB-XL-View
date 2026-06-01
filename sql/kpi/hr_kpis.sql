-- ============================================================
-- MinervaDB XL View - HR KPI Queries
-- hr_kpis.sql
-- ============================================================
-- Human Resources analytics KPIs for the MinervaDB XL View
-- HR Dashboard. Covers headcount, attrition, performance,
-- compensation, and workforce planning metrics.
-- ============================================================


-- ============================================================
-- KPI 1: Headcount by Department (Current)
-- ============================================================
SELECT
    department,
    COUNT(*)                                                    AS total_headcount,
    COUNT(CASE WHEN employment_status = 'Active' THEN 1 END)   AS active_headcount,
    COUNT(CASE WHEN gender = 'Female' THEN 1 END)              AS female_count,
    COUNT(CASE WHEN gender = 'Male' THEN 1 END)                AS male_count,
    ROUND(100.0 * COUNT(CASE WHEN gender = 'Female' THEN 1 END) / NULLIF(COUNT(*), 0), 1) AS female_pct,
    ROUND(AVG(salary), 2)                                      AS avg_salary,
    ROUND(AVG(EXTRACT(EPOCH FROM (CURRENT_DATE - hire_date)) / (365.25 * 86400)), 2) AS avg_tenure_years
FROM employees
WHERE employment_status = 'Active'
GROUP BY department
ORDER BY active_headcount DESC;


-- ============================================================
-- KPI 2: Monthly Attrition Rate (Rolling 12 Months)
-- ============================================================
WITH monthly_data AS (
    SELECT
        DATE_TRUNC('month', gs.month_start) AS month,
        COUNT(CASE WHEN e.hire_date <= gs.month_start
                        AND (t.termination_date IS NULL OR t.termination_date > gs.month_start)
                   THEN 1 END) AS beginning_headcount,
        COUNT(CASE WHEN DATE_TRUNC('month', t.termination_date) = DATE_TRUNC('month', gs.month_start)
                   THEN 1 END) AS terminations
    FROM generate_series(
        CURRENT_DATE - INTERVAL '12 months',
        CURRENT_DATE,
        INTERVAL '1 month'
    ) AS gs(month_start)
    CROSS JOIN employees e
    LEFT JOIN terminations t ON e.employee_id = t.employee_id
    GROUP BY gs.month_start
)
SELECT
    month,
    TO_CHAR(month, 'Mon YYYY')                                  AS month_label,
    beginning_headcount,
    terminations,
    ROUND(100.0 * terminations / NULLIF(beginning_headcount, 0), 2) AS monthly_attrition_rate,
    ROUND(100.0 * terminations / NULLIF(beginning_headcount, 0) * 12, 2) AS annualized_attrition_rate
FROM monthly_data
ORDER BY month DESC;


-- ============================================================
-- KPI 3: Performance Rating Distribution by Department
-- ============================================================
SELECT
    department,
    performance_rating,
    COUNT(*)                                                    AS employee_count,
    ROUND(100.0 * COUNT(*) / SUM(COUNT(*)) OVER (PARTITION BY department), 1) AS pct_of_dept,
    ROUND(AVG(salary), 2)                                      AS avg_salary_for_rating
FROM employees
WHERE employment_status = 'Active'
  AND performance_rating IS NOT NULL
GROUP BY department, performance_rating
ORDER BY department, performance_rating DESC;


-- ============================================================
-- KPI 4: New Hires vs Terminations (Rolling 6 Months)
-- ============================================================
SELECT
    DATE_TRUNC('month', event_date)                             AS month,
    TO_CHAR(DATE_TRUNC('month', event_date), 'Mon YYYY')        AS month_label,
    SUM(CASE WHEN event_type = 'Hire' THEN 1 ELSE 0 END)       AS new_hires,
    SUM(CASE WHEN event_type = 'Termination' THEN 1 ELSE 0 END) AS terminations,
    SUM(CASE WHEN event_type = 'Hire' THEN 1 ELSE 0 END) -
    SUM(CASE WHEN event_type = 'Termination' THEN 1 ELSE 0 END) AS net_headcount_change
FROM (
    SELECT hire_date AS event_date, 'Hire' AS event_type
    FROM employees
    WHERE hire_date >= CURRENT_DATE - INTERVAL '6 months'
    UNION ALL
    SELECT termination_date AS event_date, 'Termination' AS event_type
    FROM terminations
    WHERE termination_date >= CURRENT_DATE - INTERVAL '6 months'
) events
GROUP BY DATE_TRUNC('month', event_date)
ORDER BY month DESC;


-- ============================================================
-- KPI 5: Compensation Band Analysis by Job Title
-- ============================================================
SELECT
    job_title,
    department,
    COUNT(*)                                                    AS employee_count,
    MIN(salary)                                                 AS min_salary,
    PERCENTILE_CONT(0.25) WITHIN GROUP (ORDER BY salary)       AS p25_salary,
    PERCENTILE_CONT(0.50) WITHIN GROUP (ORDER BY salary)       AS median_salary,
    PERCENTILE_CONT(0.75) WITHIN GROUP (ORDER BY salary)       AS p75_salary,
    MAX(salary)                                                 AS max_salary,
    ROUND(AVG(salary), 2)                                      AS avg_salary
FROM employees
WHERE employment_status = 'Active'
GROUP BY job_title, department
HAVING COUNT(*) >= 3
ORDER BY department, avg_salary DESC;


-- ============================================================
-- KPI 6: HR Executive Scorecard (Single Row)
-- ============================================================
WITH hr_summary AS (
    SELECT
        COUNT(CASE WHEN employment_status = 'Active' THEN 1 END) AS active_hc,
        ROUND(AVG(salary), 2) AS avg_salary,
        ROUND(AVG(EXTRACT(EPOCH FROM (CURRENT_DATE - hire_date)) / (365.25 * 86400)), 2) AS avg_tenure_years,
        COUNT(CASE WHEN performance_rating IN ('Exceeds Expectations','Outstanding') THEN 1 END) AS high_performers
    FROM employees
),
ytd_attrition AS (
    SELECT COUNT(*) AS ytd_terminations
    FROM terminations
    WHERE termination_date >= DATE_TRUNC('year', CURRENT_DATE)
)
SELECT
    h.active_hc                                                AS active_headcount,
    h.avg_salary                                               AS avg_salary,
    h.avg_tenure_years                                         AS avg_tenure_years,
    ROUND(100.0 * h.high_performers / NULLIF(h.active_hc, 0), 1) AS high_performer_pct,
    a.ytd_terminations                                         AS ytd_terminations,
    ROUND(100.0 * a.ytd_terminations / NULLIF(h.active_hc, 0), 1) AS ytd_attrition_pct,
    TO_CHAR(CURRENT_DATE, 'Month DD, YYYY')                    AS as_of_date
FROM hr_summary h, ytd_attrition a;
